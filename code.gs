/**
 * Sistema de Revisão Espaçada para Residência Médica
 * Backend em Google Apps Script
 */

// ============================================================================
// CONFIGURAÇÃO E CONSTANTES
// ============================================================================

const SHEET_NAMES = {
  LOG: 'LOG',
  STATS: 'STATS',
  SPACED: 'SPACED',
  REVER_HOJE: 'REVER_HOJE',
  MODEL: 'MODEL',
  REVISAO_LOG: 'REVISAO_LOG',
  SETTINGS: 'SETTINGS'
};

const HEADERS = {
  LOG: ['data', 'area', 'subarea', 'total', 'acertos', 'tempoMedioSeg', 'difPercebida', 'flags', 'obs', 'uid'],
  STATS: ['area', 'subarea', 'total_blocos', 'questoes', 'acertos', 'acerto_vida', 'acerto_28d', 'acerto_7d', 'tempo_medio', 'flags_28d', 'dif_media', 'ultimaData'],
  SPACED: ['alvo', 'ultimaRevisao', 'estabilidade', 'dificuldade_media', 'proximaRevisao', 'lapses', 'prioridade'],
  REVER_HOJE: ['alvo', 'prioridade', 'proximaRevisao', 'estabilidade', 'feito'],
  MODEL: ['alvo', 'theta0', 'theta1', 'theta2', 'S_atual', 'ultima_atualizacao'],
  REVISAO_LOG: ['data', 'alvo', 'tDias', 'metaUsada', 'p_prev', 'acertou', 'tempoSeg', 'difPercebida', 'flags', 'obs'],
  SETTINGS: ['retentionTarget', 'wPeg', 'wTempo', 'wDif', 'alpha', 'overdueMode', 'lrEta', 'regLambda', 'halfLifeDecayDays', 'reviewOutcomeWeight', 'Smin', 'Smax', 'Imin', 'Imax', 'betaUncertainty', 'shrinkageC', 'planGainMix']
};

const DEFAULT_SETTINGS = {
  retentionTarget: 0.90,
  wPeg: 0.20,
  wTempo: 0.10,
  wDif: 0.20,
  alpha: 0.35,
  overdueMode: 'linear',
  lrEta: 0.08,
  regLambda: 0.02,
  halfLifeDecayDays: 56,
  reviewOutcomeWeight: 3.0,
  Smin: 2,
  Smax: 120,
  Imin: 2,
  Imax: 90,
  betaUncertainty: 0.50,
  shrinkageC: 8.0,
  planGainMix: 0.5
};

// ============================================================================
// SERVIDOR WEB
// ============================================================================

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Sistema de Revisão Espaçada')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ============================================================================
// UTILITÁRIOS DE PLANILHA
// ============================================================================

function getOrCreateSheet(sheetName, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    if (headers && headers.length > 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    }
  } else {
    // Verificar e garantir headers
    const existingHeaders = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
    let needsUpdate = false;
    for (let i = 0; i < headers.length; i++) {
      if (existingHeaders[i] !== headers[i]) {
        needsUpdate = true;
        break;
      }
    }
    if (needsUpdate) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
  }
  
  return sheet;
}

function readSheetData(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return [];
  
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  
  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  return data.map(row => {
    const obj = {};
    headers.forEach((header, idx) => {
      obj[header] = row[idx];
    });
    return obj;
  });
}

function writeSheetRow(sheetName, rowData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error(`Sheet ${sheetName} não encontrada`);
  
  const lastRow = sheet.getLastRow();
  const targetRow = lastRow + 1;
  
  sheet.getRange(targetRow, 1, 1, rowData.length).setValues([rowData]);
  SpreadsheetApp.flush();
}

function updateSheetRow(sheetName, rowIndex, rowData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error(`Sheet ${sheetName} não encontrada`);
  
  // rowIndex é baseado em 0, então +2 (1 para header, 1 para converter de 0-based)
  sheet.getRange(rowIndex + 2, 1, 1, rowData.length).setValues([rowData]);
  SpreadsheetApp.flush();
}

function clearSheetData(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clear();
  }
  SpreadsheetApp.flush();
}

function toIsoDate(date) {
  const copy = new Date(date);
  copy.setHours(0, 0, 0, 0);
  return Utilities.formatDate(copy, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function parseDateValue(value) {
  if (!value && value !== 0) return null;
  if (value instanceof Date) {
    const d = new Date(value.getTime());
    d.setHours(0, 0, 0, 0);
    return d;
  }
  if (typeof value === 'string' || typeof value === 'number') {
    const d = new Date(value);
    if (!isNaN(d.getTime())) {
      d.setHours(0, 0, 0, 0);
      return d;
    }
  }
  return null;
}

function diffInDays(later, earlier) {
  if (!(later instanceof Date) || !(earlier instanceof Date)) return 0;
  const msPerDay = 1000 * 60 * 60 * 24;
  return Math.floor((later.getTime() - earlier.getTime()) / msPerDay);
}

function ensureNumber(value, fallback) {
  const num = parseFloat(value);
  return isNaN(num) ? fallback : num;
}

// ============================================================================
// API: INICIALIZAÇÃO
// ============================================================================

function apiInit() {
  try {
    const lock = LockService.getScriptLock();
    lock.tryLock(10000);
    
    // Criar todas as abas com headers
    Object.keys(SHEET_NAMES).forEach(key => {
      const sheetName = SHEET_NAMES[key];
      const headers = HEADERS[key];
      getOrCreateSheet(sheetName, headers);
    });
    
    // Garantir defaults em SETTINGS
    const settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.SETTINGS);
    if (settingsSheet.getLastRow() <= 1) {
      const values = HEADERS.SETTINGS.map(key => DEFAULT_SETTINGS[key]);
      settingsSheet.appendRow(values);
      SpreadsheetApp.flush();
    }
    
    lock.releaseLock();
    return { ok: true };
  } catch (e) {
    return { ok: false, error: e.toString() };
  }
}

function apiWhereAmI() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return {
    name: ss.getName(),
    url: ss.getUrl(),
    id: ss.getId()
  };
}

// ============================================================================
// API: SETTINGS
// ============================================================================

function apiGetSettings() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.SETTINGS);
    if (!sheet || sheet.getLastRow() <= 1) {
      return DEFAULT_SETTINGS;
    }
    
    const headers = sheet.getRange(1, 1, 1, HEADERS.SETTINGS.length).getValues()[0];
    const values = sheet.getRange(2, 1, 1, HEADERS.SETTINGS.length).getValues()[0];
    
    const settings = {};
    headers.forEach((header, idx) => {
      settings[header] = values[idx];
    });
    
    return settings;
  } catch (e) {
    return DEFAULT_SETTINGS;
  }
}

function apiSaveSettings(obj) {
  try {
    const lock = LockService.getScriptLock();
    lock.tryLock(10000);
    
    const sheet = getOrCreateSheet(SHEET_NAMES.SETTINGS, HEADERS.SETTINGS);
    
    // Validações básicas
    obj.retentionTarget = Math.max(0.70, Math.min(0.98, parseFloat(obj.retentionTarget)));
    obj.Smin = Math.max(1, parseFloat(obj.Smin));
    obj.Smax = Math.max(obj.Smin, parseFloat(obj.Smax));
    obj.Imin = Math.max(1, parseFloat(obj.Imin));
    obj.Imax = Math.max(obj.Imin, parseFloat(obj.Imax));
    
    const values = HEADERS.SETTINGS.map(key => obj[key] !== undefined ? obj[key] : DEFAULT_SETTINGS[key]);
    
    if (sheet.getLastRow() <= 1) {
      sheet.appendRow(values);
    } else {
      sheet.getRange(2, 1, 1, values.length).setValues([values]);
    }
    
    SpreadsheetApp.flush();
    lock.releaseLock();
    
    return { ok: true };
  } catch (e) {
    return { ok: false, error: e.toString() };
  }
}

// ============================================================================
// API: LANÇAR BLOCO
// ============================================================================


function apiLogBlock(payload) {
  try {
    Logger.log('Iniciando apiLogBlock...');
    Logger.log('Payload recebido: ' + JSON.stringify(payload));
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = ss.getSheetByName('LOG');
    
    if (!logSheet) {
      Logger.log('ERRO: Aba LOG não encontrada');
      return { ok: false, error: 'Aba LOG não encontrada' };
    }
    
    // Converter data
    let dataObj = new Date();
    if (payload.data) {
      dataObj = new Date(payload.data);
    }
    
    const uid = Utilities.getUuid();
    
    // Dados da linha
    const rowData = [
      dataObj,
      payload.area || '',
      payload.subarea || '',
      parseInt(payload.total) || 0,
      parseInt(payload.acertos) || 0,
      parseFloat(payload.tempoMedioSeg) || 0,
      parseInt(payload.difPercebida) || 3,
      payload.flags || '',
      payload.obs || '',
      uid
    ];
    
    Logger.log('Dados a gravar: ' + JSON.stringify(rowData));
    
    // Escrever diretamente
    logSheet.appendRow(rowData);
    SpreadsheetApp.flush();
    
    Logger.log('Bloco salvo com sucesso! UID: ' + uid);
    
    return { ok: true, uid: uid };
    
  } catch (e) {
    Logger.log('ERRO em apiLogBlock: ' + e.toString());
    Logger.log('Stack: ' + e.stack);
    return { ok: false, error: e.toString() };
  }
}

function testLogBlock() {
  const payload = {
    data: '2025-01-15',
    area: 'Clínica Médica',
    subarea: 'Cardiologia',
    total: 10,
    acertos: 7,
    tempoMedioSeg: 60,
    difPercebida: 3,
    obs: 'Teste'
  };
  
  const result = apiLogBlock(payload);
  Logger.log('Resultado do teste: ' + JSON.stringify(result));
  return result;
}

// ============================================================================
// PROCESSAMENTO: ATUALIZAR STATS A PARTIR DO LOG
// ============================================================================
// ============================================================================
// PROCESSAMENTO: ATUALIZAR STATS A PARTIR DO LOG
// ============================================================================

function apiProcessLog() {
  try {
    const lock = LockService.getScriptLock();
    lock.tryLock(30000);
    
    const logData = readSheetData(SHEET_NAMES.LOG);
    const settings = apiGetSettings();
    
    if (logData.length === 0) {
      lock.releaseLock();
      return { ok: true, message: 'Nenhum dado no LOG para processar' };
    }
    
    // Agrupar por área::subárea
    const statsMap = {};
    const hoje = new Date();
    hoje.setHours(0, 0, 0, 0);
    
    logData.forEach(row => {
      const area = row.area || 'Sem área';
      const subarea = row.subarea || 'Sem subárea';
      const alvo = `${area}::${subarea}`;
      
      const dataBloco = new Date(row.data);
      dataBloco.setHours(0, 0, 0, 0);
      const diasAtras = Math.floor((hoje - dataBloco) / (1000 * 60 * 60 * 24));
      
      if (!statsMap[alvo]) {
        statsMap[alvo] = {
          area: area,
          subarea: subarea,
          total_blocos: 0,
          questoes: 0,
          acertos: 0,
          questoes_28d: 0,
          acertos_28d: 0,
          questoes_7d: 0,
          acertos_7d: 0,
          tempos: [],
          flags_28d: 0,
          dificuldades: [],
          ultimaData: dataBloco
        };
      }
      
      const stat = statsMap[alvo];
      stat.total_blocos++;
      
      const total = parseInt(row.total) || 0;
      const acertos = parseInt(row.acertos) || 0;
      const tempo = parseFloat(row.tempoMedioSeg) || 0;
      const dif = parseInt(row.difPercebida) || 3;
      
      // Totais gerais
      stat.questoes += total;
      stat.acertos += acertos;
      
      // Últimos 28 dias
      if (diasAtras <= 28) {
        stat.questoes_28d += total;
        stat.acertos_28d += acertos;
        if (row.flags) stat.flags_28d++;
      }
      
      // Últimos 7 dias
      if (diasAtras <= 7) {
        stat.questoes_7d += total;
        stat.acertos_7d += acertos;
      }
      
      // Tempo e dificuldade
      if (tempo > 0) stat.tempos.push(tempo);
      stat.dificuldades.push(dif);
      
      // Data mais recente
      if (dataBloco > stat.ultimaData) {
        stat.ultimaData = dataBloco;
      }
    });
    
    // Atualizar aba STATS
    clearSheetData(SHEET_NAMES.STATS);
    
    Object.keys(statsMap).forEach(alvo => {
      const s = statsMap[alvo];
      
      const acerto_vida = s.questoes > 0 ? s.acertos / s.questoes : 0;
      const acerto_28d = s.questoes_28d > 0 ? s.acertos_28d / s.questoes_28d : acerto_vida;
      const acerto_7d = s.questoes_7d > 0 ? s.acertos_7d / s.questoes_7d : acerto_28d;
      
      const tempo_medio = s.tempos.length > 0 
        ? s.tempos.reduce((a, b) => a + b, 0) / s.tempos.length 
        : 60;
      
      const dif_media = s.dificuldades.length > 0
        ? s.dificuldades.reduce((a, b) => a + b, 0) / s.dificuldades.length
        : 3;
      
      const rowData = [
        s.area,
        s.subarea,
        s.total_blocos,
        s.questoes,
        s.acertos,
        acerto_vida,
        acerto_28d,
        acerto_7d,
        tempo_medio,
        s.flags_28d,
        dif_media,
        s.ultimaData
      ];
      
      writeSheetRow(SHEET_NAMES.STATS, rowData);
    });
    
    // Criar/atualizar alvos em SPACED
    const spacedData = readSheetData(SHEET_NAMES.SPACED);
    
    Object.keys(statsMap).forEach(alvo => {
      const existeSpaced = spacedData.find(s => s.alvo === alvo);
      
      if (!existeSpaced) {
        // Criar novo alvo em SPACED
        const s = statsMap[alvo];
        const acerto_28d = s.questoes_28d > 0 ? s.acertos_28d / s.questoes_28d : 0.5;
        const dif_media = s.dificuldades.length > 0
          ? s.dificuldades.reduce((a, b) => a + b, 0) / s.dificuldades.length
          : 3;
        
        // Estabilidade inicial baseada na competência
        let S_inicial = settings.Smin;
        if (acerto_28d > 0.8) {
          S_inicial = settings.Smin * 2;
        } else if (acerto_28d > 0.6) {
          S_inicial = settings.Smin * 1.5;
        }
        S_inicial = Math.min(S_inicial, settings.Smax);
        
        // Primeira revisão: logo após estudar
        const proximaRevisao = new Date(s.ultimaData);
        proximaRevisao.setDate(proximaRevisao.getDate() + Math.round(S_inicial * 0.3));
        
        const newSpacedRow = [
          alvo,
          s.ultimaData,
          S_inicial,
          dif_media,
          proximaRevisao,
          0, // lapses
          0  // prioridade
        ];
        
        writeSheetRow(SHEET_NAMES.SPACED, newSpacedRow);
        
        // Criar modelo inicial
        const modelData = readSheetData(SHEET_NAMES.MODEL);
        const existeModel = modelData.find(m => m.alvo === alvo);
        
        if (!existeModel) {
          const newModelRow = [
            alvo,
            Math.log(S_inicial), // theta0
            0, // theta1
            0, // theta2
            S_inicial, // S_atual
            hoje
          ];
          writeSheetRow(SHEET_NAMES.MODEL, newModelRow);
        }
      }
    });
    
    SpreadsheetApp.flush();
    lock.releaseLock();
    
    return { 
      ok: true, 
      alvosProcessados: Object.keys(statsMap).length,
      message: `${Object.keys(statsMap).length} alvos processados com sucesso`
    };
  } catch (e) {
    Logger.log('Erro em apiProcessLog: ' + e.toString());
    return { ok: false, error: e.toString() };
  }
}

// ============================================================================
// API: PROCESSAR TUDO (LOG → STATS → SPACED → FILA)
// ============================================================================

function apiProcessAll() {
  try {
    // 1. Processar LOG → STATS + SPACED
    const processResult = apiProcessLog();
    if (!processResult.ok) {
      return processResult;
    }
    
    // 2. Gerar fila de revisões
    const reviewResult = apiMakeReviewToday();
    if (!reviewResult.ok) {
      return reviewResult;
    }
    
    return {
      ok: true,
      alvosProcessados: processResult.alvosProcessados,
      revisoesHoje: reviewResult.count,
      message: `Processamento completo: ${processResult.alvosProcessados} alvos, ${reviewResult.count} revisões hoje`
    };
  } catch (e) {
    return { ok: false, error: e.toString() };
  }
}

function apiProcessLog() {
  try {
    const lock = LockService.getScriptLock();
    lock.tryLock(30000);
    
    const logData = readSheetData(SHEET_NAMES.LOG);
    const settings = apiGetSettings();
    
    if (logData.length === 0) {
      lock.releaseLock();
      return { ok: true, message: 'Nenhum dado no LOG para processar' };
    }
    
    // Agrupar por área::subárea
    const statsMap = {};
    const hoje = new Date();
    hoje.setHours(0, 0, 0, 0);
    
    logData.forEach(row => {
      const area = row.area || 'Sem área';
      const subarea = row.subarea || 'Sem subárea';
      const alvo = `${area}::${subarea}`;
      
      const dataBloco = new Date(row.data);
      dataBloco.setHours(0, 0, 0, 0);
      const diasAtras = Math.floor((hoje - dataBloco) / (1000 * 60 * 60 * 24));
      
      if (!statsMap[alvo]) {
        statsMap[alvo] = {
          area: area,
          subarea: subarea,
          total_blocos: 0,
          questoes: 0,
          acertos: 0,
          questoes_28d: 0,
          acertos_28d: 0,
          questoes_7d: 0,
          acertos_7d: 0,
          tempos: [],
          flags_28d: 0,
          dificuldades: [],
          ultimaData: dataBloco
        };
      }
      
      const stat = statsMap[alvo];
      stat.total_blocos++;
      
      const total = parseInt(row.total) || 0;
      const acertos = parseInt(row.acertos) || 0;
      const tempo = parseFloat(row.tempoMedioSeg) || 0;
      const dif = parseInt(row.difPercebida) || 3;
      
      // Totais gerais
      stat.questoes += total;
      stat.acertos += acertos;
      
      // Últimos 28 dias
      if (diasAtras <= 28) {
        stat.questoes_28d += total;
        stat.acertos_28d += acertos;
        if (row.flags) stat.flags_28d++;
      }
      
      // Últimos 7 dias
      if (diasAtras <= 7) {
        stat.questoes_7d += total;
        stat.acertos_7d += acertos;
      }
      
      // Tempo e dificuldade
      if (tempo > 0) stat.tempos.push(tempo);
      stat.dificuldades.push(dif);
      
      // Data mais recente
      if (dataBloco > stat.ultimaData) {
        stat.ultimaData = dataBloco;
      }
    });
    
    // Atualizar aba STATS
    clearSheetData(SHEET_NAMES.STATS);
    const statsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.STATS);
    
    Object.keys(statsMap).forEach(alvo => {
      const s = statsMap[alvo];
      
      const acerto_vida = s.questoes > 0 ? s.acertos / s.questoes : 0;
      const acerto_28d = s.questoes_28d > 0 ? s.acertos_28d / s.questoes_28d : acerto_vida;
      const acerto_7d = s.questoes_7d > 0 ? s.acertos_7d / s.questoes_7d : acerto_28d;
      
      const tempo_medio = s.tempos.length > 0 
        ? s.tempos.reduce((a, b) => a + b, 0) / s.tempos.length 
        : 60;
      
      const dif_media = s.dificuldades.length > 0
        ? s.dificuldades.reduce((a, b) => a + b, 0) / s.dificuldades.length
        : 3;
      
      const rowData = [
        s.area,
        s.subarea,
        s.total_blocos,
        s.questoes,
        s.acertos,
        acerto_vida,
        acerto_28d,
        acerto_7d,
        tempo_medio,
        s.flags_28d,
        dif_media,
        s.ultimaData
      ];
      
      writeSheetRow(SHEET_NAMES.STATS, rowData);
    });
    
    // Criar/atualizar alvos em SPACED
    const spacedData = readSheetData(SHEET_NAMES.SPACED);
    const spacedSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.SPACED);
    
    Object.keys(statsMap).forEach(alvo => {
      const existeSpaced = spacedData.find(s => s.alvo === alvo);
      
      if (!existeSpaced) {
        // Criar novo alvo em SPACED
        const s = statsMap[alvo];
        const acerto_28d = s.questoes_28d > 0 ? s.acertos_28d / s.questoes_28d : 0.5;
        const dif_media = s.dificuldades.length > 0
          ? s.dificuldades.reduce((a, b) => a + b, 0) / s.dificuldades.length
          : 3;
        
        // Estabilidade inicial baseada na competência
        let S_inicial = settings.Smin;
        if (acerto_28d > 0.8) {
          S_inicial = settings.Smin * 2;
        } else if (acerto_28d > 0.6) {
          S_inicial = settings.Smin * 1.5;
        }
        S_inicial = Math.min(S_inicial, settings.Smax);
        
        // Primeira revisão: logo após estudar
        const proximaRevisao = new Date(s.ultimaData);
        proximaRevisao.setDate(proximaRevisao.getDate() + Math.round(S_inicial * 0.3));
        
        const newSpacedRow = [
          alvo,
          s.ultimaData,
          S_inicial,
          dif_media,
          proximaRevisao,
          0, // lapses
          0  // prioridade
        ];
        
        writeSheetRow(SHEET_NAMES.SPACED, newSpacedRow);
        
        // Criar modelo inicial
        const modelData = readSheetData(SHEET_NAMES.MODEL);
        const existeModel = modelData.find(m => m.alvo === alvo);
        
        if (!existeModel) {
          const newModelRow = [
            alvo,
            Math.log(S_inicial), // theta0
            0, // theta1
            0, // theta2
            S_inicial, // S_atual
            hoje
          ];
          writeSheetRow(SHEET_NAMES.MODEL, newModelRow);
        }
      }
    });
    
    SpreadsheetApp.flush();
    lock.releaseLock();
    
    return { 
      ok: true, 
      alvosProcessados: Object.keys(statsMap).length,
      message: `${Object.keys(statsMap).length} alvos processados com sucesso`
    };
  } catch (e) {
    Logger.log('Erro em apiProcessLog: ' + e.toString());
    return { ok: false, error: e.toString() };
  }
}

// ============================================================================
// API: PROCESSAR TUDO (LOG → STATS → SPACED → FILA)
// ============================================================================

function apiProcessAll() {
  try {
    // Validar que as abas existem
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss.getSheetByName(SHEET_NAMES.LOG)) {
      return { ok: false, error: 'Aba LOG não encontrada' };
    }
    
    // 1. Processar LOG → STATS + SPACED
    Logger.log('Iniciando processamento...');
    const processResult = apiProcessLog();
    
    if (!processResult || !processResult.ok) {
      Logger.log('Erro no processamento: ' + JSON.stringify(processResult));
      return { ok: false, error: processResult ? processResult.error : 'Erro desconhecido' };
    }
    
    Logger.log('Processamento OK, gerando fila...');
    
    // 2. Gerar fila de revisões (com timeout protection)
    let reviewResult;
    try {
      reviewResult = apiMakeReviewToday();
    } catch (e) {
      Logger.log('Erro ao gerar fila: ' + e.toString());
      reviewResult = { ok: true, count: 0 }; // Continuar mesmo sem fila
    }
    
    return {
      ok: true,
      alvosProcessados: processResult.alvosProcessados || 0,
      revisoesHoje: reviewResult.count || 0,
      message: `✓ ${processResult.alvosProcessados || 0} alvos processados, ${reviewResult.count || 0} revisões hoje`
    };
  } catch (e) {
    Logger.log('Erro em apiProcessAll: ' + e.toString());
    return { ok: false, error: e.toString() };
  }
}
// ============================================================================
// API: ESTATÍSTICAS E GRÁFICOS
// ============================================================================


function apiGetStats() {
  try {
    const stats = readSheetData(SHEET_NAMES.STATS);
    return { ok: true, data: stats };
  } catch (e) {
    return { ok: false, error: e.toString() };
  }
}

function apiChartData() {
  try {
    const stats = readSheetData(SHEET_NAMES.STATS);
    
    // Agrupar por área
    const areaMap = {};
    stats.forEach(row => {
      const area = row.area || 'Sem área';
      if (!areaMap[area]) {
        areaMap[area] = { total: 0, acertos: 0 };
      }
      areaMap[area].total += parseFloat(row.questoes) || 0;
      areaMap[area].acertos += parseFloat(row.acertos) || 0;
    });
    
    const chartData = [['Área', 'Acerto %']];
    Object.keys(areaMap).forEach(area => {
      const pct = areaMap[area].total > 0 ? (areaMap[area].acertos / areaMap[area].total) * 100 : 0;
      chartData.push([area, pct]);
    });
    
    return { ok: true, data: chartData };
  } catch (e) {
    return { ok: false, error: e.toString() };
  }
}

// ============================================================================
// ALGORITMO: FUNÇÕES AUXILIARES
// ============================================================================

function calcRecall(t, S) {
  // R(t) = e^(-t/S)
  return Math.exp(-t / S);
}

function calcOptimalInterval(S, retentionTarget) {
  // I = -S * ln(meta)
  return -S * Math.log(retentionTarget);
}

function calcStability(theta0, theta1, theta2, competencia, difNorm) {
  // ln(S) = θ0 + θ1·competência + θ2·difNorm
  const lnS = theta0 + theta1 * competencia + theta2 * difNorm;
  return Math.exp(lnS);
}

function calcSobs(t, retentionTarget) {
  // S_obs = t / (-ln(meta))
  return t / (-Math.log(retentionTarget));
}

function applyCapS(S, Smin, Smax) {
  return Math.max(Smin, Math.min(Smax, S));
}

function applyCapI(I, Imin, Imax) {
  return Math.max(Imin, Math.min(Imax, I));
}

function calcOverdue(atrasoDias, S, alpha, mode) {
  if (mode === 'softplus') {
    const x = atrasoDias / S;
    return alpha * Math.log(1 + Math.exp(x));
  }
  // linear (default)
  return alpha * Math.min(1.5, atrasoDias / S);
}

function normalizeDif(difPercebida) {
  // dif ∈ [1,5] → [0,1]
  return (difPercebida - 1) / 4;
}

function buildReviewItem(alvo, context, settings, dataReferencia, doneMap) {
  const spacedRow = context.spacedMap[alvo];
  const statRow = context.statsMap[alvo];
  const modelRow = context.modelMap[alvo];

  if (!spacedRow && !statRow) {
    return null;
  }

  const hoje = new Date(dataReferencia.getTime());
  hoje.setHours(0, 0, 0, 0);

  let area = '';
  let subarea = '';
  if (alvo && alvo.indexOf('::') >= 0) {
    const partes = alvo.split('::');
    area = partes[0] || '';
    subarea = partes[1] || '';
  } else if (statRow) {
    area = statRow.area || '';
    subarea = statRow.subarea || '';
  }

  let S = null;
  if (modelRow && modelRow.S_atual !== undefined) {
    S = ensureNumber(modelRow.S_atual, null);
  }
  if (!S && spacedRow && spacedRow.estabilidade !== undefined) {
    S = ensureNumber(spacedRow.estabilidade, null);
  }
  if (!S) {
    S = settings ? settings.Smin : 1;
  }
  S = applyCapS(S, settings.Smin, settings.Smax);

  const ultimaRevisao = spacedRow ? parseDateValue(spacedRow.ultimaRevisao) : null;
  const ultimaFallback = !ultimaRevisao && statRow ? parseDateValue(statRow.ultimaData) : null;
  const ultima = ultimaRevisao || ultimaFallback;

  let proxima = spacedRow ? parseDateValue(spacedRow.proximaRevisao) : null;
  if (!proxima && ultima) {
    let intervalo = calcOptimalInterval(S, settings.retentionTarget);
    intervalo = applyCapI(intervalo, settings.Imin, settings.Imax);
    proxima = new Date(ultima);
    proxima.setDate(proxima.getDate() + Math.round(intervalo));
  }
  if (!proxima) {
    proxima = new Date(hoje);
  }

  if (proxima.getTime() > hoje.getTime()) {
    return null;
  }

  let tDias = ultima ? Math.max(1, diffInDays(hoje, ultima)) : Math.max(1, diffInDays(hoje, proxima));

  const Rhoje = calcRecall(tDias, S);
  const basePrioridade = 1 - Rhoje;

  const atrasoDias = Math.max(0, diffInDays(hoje, proxima));
  const overdue = atrasoDias > 0 ? calcOverdue(atrasoDias, S, settings.alpha, settings.overdueMode) : 0;

  const stat = statRow || {};
  const flags28d = ensureNumber(stat.flags_28d, 0);
  const totalBlocos = ensureNumber(stat.total_blocos, 0);
  const peg = Math.min(1, totalBlocos > 0 ? flags28d / totalBlocos : flags28d / 10);

  const tempoMedio = ensureNumber(stat.tempo_medio, 60);
  const tempo_rel = Math.min(1, tempoMedio / 120);

  const difMedia = ensureNumber(stat.dif_media, 3);
  const dif_norm = normalizeDif(difMedia);

  // Calcula prioridade consolidada com pesos (1 − R(t)) + overdue + custos
  const prioridade = basePrioridade + overdue + settings.wPeg * peg + settings.wTempo * tempo_rel + settings.wDif * dif_norm;

  const mediaQuestoes = totalBlocos > 0
    ? Math.max(1, Math.round(ensureNumber(stat.questoes, 0) / totalBlocos))
    : 10;
  const tempoEstMin = Math.max(1, Math.round((tempoMedio * mediaQuestoes) / 60));

  const feito = !!(doneMap && doneMap[alvo]);
  const lapses = spacedRow ? ensureNumber(spacedRow.lapses, 0) : 0;

  return {
    alvo: alvo,
    area: area || 'Geral',
    subarea: subarea || 'Geral',
    prioridade: parseFloat(prioridade.toFixed(3)),
    S: parseFloat(S.toFixed(2)),
    Rhoje: parseFloat(Rhoje.toFixed(3)),
    overdue: parseFloat(overdue.toFixed(3)),
    tempoEstMin: tempoEstMin,
    peg: parseFloat(peg.toFixed(3)),
    tempo_rel: parseFloat(tempo_rel.toFixed(3)),
    dif_norm: parseFloat(dif_norm.toFixed(3)),
    feito: feito,
    proximaRevisao: Utilities.formatDate(proxima, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
    caps: {
      Smin: settings.Smin,
      Smax: settings.Smax,
      Imin: settings.Imin,
      Imax: settings.Imax
    },
    lapses: lapses,
    tDias: tDias,
    R_prev: parseFloat(Rhoje.toFixed(4))
  };
}

// ============================================================================
// API: REVISÕES - CRIAR FILA DO DIA
// ============================================================================

function apiMakeReviewToday() {
  try {
    const settings = apiGetSettings();
    const hoje = new Date();
    hoje.setHours(0, 0, 0, 0);

    const spaced = readSheetData(SHEET_NAMES.SPACED);
    const stats = readSheetData(SHEET_NAMES.STATS);
    const model = readSheetData(SHEET_NAMES.MODEL);
    const reviewHoje = readSheetData(SHEET_NAMES.REVER_HOJE);

    const spacedMap = {};
    spaced.forEach(row => {
      if (row.alvo) spacedMap[row.alvo] = row;
    });

    const statsMap = {};
    stats.forEach(row => {
      const alvo = `${row.area || ''}::${row.subarea || ''}`;
      statsMap[alvo] = row;
    });

    const modelMap = {};
    model.forEach(row => {
      if (row.alvo) modelMap[row.alvo] = row;
    });

    const docProps = PropertiesService.getDocumentProperties();
    const storedDate = docProps.getProperty('REVER_HOJE_DATE');
    const todayIso = toIsoDate(hoje);

    const doneMap = {};
    if (storedDate === todayIso) {
      reviewHoje.forEach(row => {
        if (row.alvo) {
          const value = row.feito;
          doneMap[row.alvo] = value === 1 || value === true || value === '1';
        }
      });
    }

    const context = { spacedMap: spacedMap, statsMap: statsMap, modelMap: modelMap };
    const candidatos = new Set();

    spaced.forEach(row => {
      if (row.alvo) {
        const prox = parseDateValue(row.proximaRevisao);
        if (prox && prox.getTime() <= hoje.getTime()) {
          candidatos.add(row.alvo);
        }
      }
    });

    reviewHoje.forEach(row => {
      if (row.alvo) candidatos.add(row.alvo);
    });

    Object.keys(statsMap).forEach(alvo => candidatos.add(alvo));

    const lista = [];
    candidatos.forEach(alvo => {
      const item = buildReviewItem(alvo, context, settings, hoje, doneMap);
      if (item) {
        lista.push(item);
      }
    });

    lista.sort((a, b) => b.prioridade - a.prioridade);

    const reviewSheet = getOrCreateSheet(SHEET_NAMES.REVER_HOJE, HEADERS.REVER_HOJE);
    clearSheetData(SHEET_NAMES.REVER_HOJE);

    // Reconstrói REVER_HOJE a partir de SPACED mantendo marcações do dia
    lista.forEach(item => {
      const rowData = [
        item.alvo,
        item.prioridade,
        item.proximaRevisao,
        item.S,
        item.feito ? 1 : 0
      ];
      reviewSheet.appendRow(rowData);
    });

    SpreadsheetApp.flush();
    docProps.setProperty('REVER_HOJE_DATE', todayIso);

    return { ok: true, date: todayIso, items: lista };
  } catch (e) {
    return { ok: false, error: e.toString() };
  }
}

function debugSpaced() {
  try {
    const spaced = readSheetData(SHEET_NAMES.SPACED);
    Logger.log('Total de registros em SPACED: ' + spaced.length);
    
    if (spaced.length > 0) {
      Logger.log('Primeiro registro: ' + JSON.stringify(spaced[0]));
      Logger.log('Colunas: ' + Object.keys(spaced[0]).join(', '));
    }
    
    const hoje = new Date();
    hoje.setHours(0, 0, 0, 0);
    
    spaced.forEach((item, idx) => {
      Logger.log(`\n--- Item ${idx} ---`);
      Logger.log('Alvo: ' + item.alvo);
      Logger.log('Próxima revisão: ' + item.proximaRevisao);
      Logger.log('Estabilidade: ' + item.estabilidade);
      
      if (item.proximaRevisao) {
        const proxRev = new Date(item.proximaRevisao);
        proxRev.setHours(0, 0, 0, 0);
        Logger.log('Próxima revisão (processada): ' + proxRev);
        Logger.log('Hoje: ' + hoje);
        Logger.log('Está vencido? ' + (proxRev <= hoje));
      }
    });
    
    return 'Ver logs';
  } catch (e) {
    Logger.log('Erro: ' + e.toString());
    return e.toString();
  }
}

function apiGetDayDetails(dateISO) {
  try {
    const settings = apiGetSettings();
    const referencia = dateISO ? parseDateValue(`${dateISO}T00:00:00`) : new Date();
    if (!referencia) {
      return { ok: false, error: 'Data inválida' };
    }
    referencia.setHours(0, 0, 0, 0);

    const spaced = readSheetData(SHEET_NAMES.SPACED);
    const stats = readSheetData(SHEET_NAMES.STATS);
    const model = readSheetData(SHEET_NAMES.MODEL);
    const reviewHoje = readSheetData(SHEET_NAMES.REVER_HOJE);

    const spacedMap = {};
    spaced.forEach(row => {
      if (row.alvo) spacedMap[row.alvo] = row;
    });

    const statsMap = {};
    stats.forEach(row => {
      const alvo = `${row.area || ''}::${row.subarea || ''}`;
      statsMap[alvo] = row;
    });

    const modelMap = {};
    model.forEach(row => {
      if (row.alvo) modelMap[row.alvo] = row;
    });

    const docProps = PropertiesService.getDocumentProperties();
    const storedDate = docProps.getProperty('REVER_HOJE_DATE');
    const dateIso = toIsoDate(referencia);

    const doneMap = {};
    if (storedDate === dateIso) {
      reviewHoje.forEach(row => {
        if (row.alvo) {
          const value = row.feito;
          doneMap[row.alvo] = value === 1 || value === true || value === '1';
        }
      });
    }

    const context = { spacedMap: spacedMap, statsMap: statsMap, modelMap: modelMap };
    const candidatos = new Set();
    spaced.forEach(row => {
      if (row.alvo) {
        const prox = parseDateValue(row.proximaRevisao);
        if (prox && prox.getTime() <= referencia.getTime()) {
          candidatos.add(row.alvo);
        }
      }
    });
    reviewHoje.forEach(row => {
      if (row.alvo) candidatos.add(row.alvo);
    });
    Object.keys(statsMap).forEach(alvo => candidatos.add(alvo));

    const lista = [];
    candidatos.forEach(alvo => {
      const item = buildReviewItem(alvo, context, settings, referencia, doneMap);
      if (item) lista.push(item);
    });

    lista.sort((a, b) => b.prioridade - a.prioridade);

    const pendentes = lista.filter(item => !item.feito);
    const tempoTotal = pendentes.reduce((acc, item) => acc + (item.tempoEstMin || 0), 0);
    const top3 = lista.slice(0, 3).map(item => ({ alvo: item.alvo, prioridade: item.prioridade }));

    return {
      ok: true,
      data: {
        date: dateIso,
        total: lista.length,
        pendentes: pendentes.length,
        tempoTotalMin: tempoTotal,
        top: top3
      }
    };
  } catch (e) {
    return { ok: false, error: e.toString() };
  }
}

function apiApplyReviewDone(payload) {
  const lock = LockService.getScriptLock();
  lock.tryLock(30000);

  try {
    if (!payload || !payload.alvo) {
      throw new Error('Payload inválido para aplicar revisão');
    }

    const alvo = payload.alvo;
    const undo = payload.undo === true;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const reviewSheet = getOrCreateSheet(SHEET_NAMES.REVER_HOJE, HEADERS.REVER_HOJE);
    const reviewData = readSheetData(SHEET_NAMES.REVER_HOJE);
    const reviewIdx = reviewData.findIndex(row => row.alvo === alvo);

    if (reviewIdx === -1) {
      throw new Error('Alvo não encontrado na fila do dia');
    }

    if (undo) {
      // Marca como não feito sem alterar demais dados
      reviewSheet.getRange(reviewIdx + 2, 5).setValue(0); // coluna feito
      SpreadsheetApp.flush();
      return { ok: true, alvo: alvo, undone: true, recomputeHint: false };
    }

    const settings = apiGetSettings();
    const hoje = new Date();
    hoje.setHours(0, 0, 0, 0);

    const totalQuestoes = Math.max(1, parseInt(payload.total, 10) || 0);
    const acertos = Math.min(totalQuestoes, Math.max(0, parseInt(payload.acertos, 10)));
    const tempoSeg = Math.max(10, parseFloat(payload.tempoSeg) || totalQuestoes * 60);
    const difPercebida = Math.min(5, Math.max(1, parseInt(payload.difPercebida, 10) || 3));
    const flags = payload.flags || '';
    const obs = payload.obs || '';

    const spacedData = readSheetData(SHEET_NAMES.SPACED);
    const statsData = readSheetData(SHEET_NAMES.STATS);
    const modelData = readSheetData(SHEET_NAMES.MODEL);

    const spacedMap = {};
    spacedData.forEach((row, idx) => {
      if (row.alvo) {
        spacedMap[row.alvo] = { row: row, idx: idx };
      }
    });

    const statsMap = {};
    statsData.forEach((row, idx) => {
      const key = `${row.area || ''}::${row.subarea || ''}`;
      statsMap[key] = { row: row, idx: idx };
    });

    const modelMap = {};
    modelData.forEach((row, idx) => {
      if (row.alvo) {
        modelMap[row.alvo] = { row: row, idx: idx };
      }
    });

    const [areaRaw, subRaw] = alvo.split('::');
    const area = areaRaw || 'Geral';
    const subarea = subRaw || 'Geral';

    const statsSheet = getOrCreateSheet(SHEET_NAMES.STATS, HEADERS.STATS);
    let statBundle = statsMap[alvo];
    if (!statBundle) {
      const novoStat = [area, subarea, 0, 0, 0, 0, 0, 0, 60, 0, 3, hoje];
      writeSheetRow(SHEET_NAMES.STATS, novoStat);
      statsData.push({
        area: area,
        subarea: subarea,
        total_blocos: 0,
        questoes: 0,
        acertos: 0,
        acerto_vida: 0,
        acerto_28d: 0,
        acerto_7d: 0,
        tempo_medio: 60,
        flags_28d: 0,
        dif_media: 3,
        ultimaData: hoje
      });
      statBundle = { row: statsData[statsData.length - 1], idx: statsData.length - 1 };
      statsMap[alvo] = statBundle;
    }

    const spacedSheet = getOrCreateSheet(SHEET_NAMES.SPACED, HEADERS.SPACED);
    let spacedBundle = spacedMap[alvo];
    if (!spacedBundle) {
      const novoSpaced = [alvo, hoje, settings.Smin, difPercebida, hoje, 0, 0];
      writeSheetRow(SHEET_NAMES.SPACED, novoSpaced);
      spacedData.push({
        alvo: alvo,
        ultimaRevisao: hoje,
        estabilidade: settings.Smin,
        dificuldade_media: difPercebida,
        proximaRevisao: hoje,
        lapses: 0,
        prioridade: 0
      });
      spacedBundle = { row: spacedData[spacedData.length - 1], idx: spacedData.length - 1 };
      spacedMap[alvo] = spacedBundle;
    }

    const modelSheet = getOrCreateSheet(SHEET_NAMES.MODEL, HEADERS.MODEL);
    let modelBundle = modelMap[alvo];
    if (!modelBundle) {
      const theta0 = Math.log(settings.Smin);
      const novoModel = [alvo, theta0, 0, 0, settings.Smin, hoje];
      writeSheetRow(SHEET_NAMES.MODEL, novoModel);
      modelData.push({
        alvo: alvo,
        theta0: theta0,
        theta1: 0,
        theta2: 0,
        S_atual: settings.Smin,
        ultima_atualizacao: hoje
      });
      modelBundle = { row: modelData[modelData.length - 1], idx: modelData.length - 1 };
      modelMap[alvo] = modelBundle;
    }

    const spacedRow = spacedBundle.row;
    const statRow = statBundle.row;
    const modelRow = modelBundle.row;

    const ultimaRevisao = parseDateValue(spacedRow.ultimaRevisao) || parseDateValue(statRow.ultimaData) || hoje;
    let tDias = Math.max(1, diffInDays(hoje, ultimaRevisao));

    let S_atual = ensureNumber(modelRow.S_atual, settings.Smin);
    S_atual = applyCapS(S_atual, settings.Smin, settings.Smax);
    const pPrev = calcRecall(tDias, S_atual);

    const reviewLogRow = [
      hoje,
      alvo,
      tDias,
      settings.retentionTarget,
      pPrev,
      0,
      tempoSeg,
      difPercebida,
      flags,
      obs
    ];

    const competencia = ensureNumber(statRow.acerto_28d, ensureNumber(statRow.acerto_vida, 0.5));
    const difNorm = normalizeDif(difPercebida);

    const theta0 = ensureNumber(modelRow.theta0, Math.log(settings.Smin));
    const theta1 = ensureNumber(modelRow.theta1, 0);
    const theta2 = ensureNumber(modelRow.theta2, 0);
    const x = [1, competencia, difNorm];

    const S_hat = calcStability(theta0, theta1, theta2, competencia, difNorm);
    const S_obs = calcSobs(Math.max(1, tDias), settings.retentionTarget);

    const lnS_obs = Math.log(S_obs);
    const lnS_hat = Math.log(S_hat);
    const erro = lnS_obs - lnS_hat;

    const thetas = [theta0, theta1, theta2];
    for (let i = 0; i < thetas.length; i++) {
      thetas[i] = (1 - settings.regLambda) * thetas[i] + settings.lrEta * erro * x[i];
    }

    let S_novo = applyCapS(S_obs, settings.Smin, settings.Smax);

    const accuracy = acertos / totalQuestoes;
    const acertou = accuracy >= 0.7;
    reviewLogRow[5] = acertou ? 1 : 0;
    if (!acertou) {
      S_novo = Math.max(settings.Smin, S_novo * 0.7);
    }

    writeSheetRow(SHEET_NAMES.REVISAO_LOG, reviewLogRow);

    let intervalo = calcOptimalInterval(S_novo, settings.retentionTarget);
    intervalo = applyCapI(intervalo, settings.Imin, settings.Imax);

    const acerto7d = ensureNumber(statRow.acerto_7d, ensureNumber(statRow.acerto_vida, 0.5));
    if (acerto7d < 0.5) {
      intervalo = Math.max(settings.Imin, intervalo * 0.7);
    }
    if (!acertou) {
      intervalo = Math.max(settings.Imin, intervalo * 0.5);
    }

    const proximaRevisao = new Date(hoje);
    proximaRevisao.setDate(proximaRevisao.getDate() + Math.max(1, Math.round(intervalo)));

    const lapsesAtual = ensureNumber(spacedRow.lapses, 0);
    const lapsesNovos = acertou ? lapsesAtual : lapsesAtual + 1;

    const tempoPrevTotal = ensureNumber(statRow.tempo_medio, 60) * ensureNumber(statRow.questoes, 0);
    const questoesPrev = ensureNumber(statRow.questoes, 0);
    const acertosPrev = ensureNumber(statRow.acertos, 0);
    const totalBlocosPrev = ensureNumber(statRow.total_blocos, 0);
    const difMediaPrev = ensureNumber(statRow.dif_media, 3);
    const flagsPrev = ensureNumber(statRow.flags_28d, 0);
    const questoes28Prev = ensureNumber(statRow.questoes_28d, 0);
    const acertos28Prev = ensureNumber(statRow.acertos_28d, 0);
    const questoes7Prev = ensureNumber(statRow.questoes_7d, 0);
    const acertos7Prev = ensureNumber(statRow.acertos_7d, 0);

    const totalBlocosNovo = totalBlocosPrev + 1;
    const questoesNovo = questoesPrev + totalQuestoes;
    const acertosNovo = acertosPrev + acertos;
    const questoes28Novo = questoes28Prev + totalQuestoes;
    const acertos28Novo = acertos28Prev + acertos;
    const questoes7Novo = questoes7Prev + totalQuestoes;
    const acertos7Novo = acertos7Prev + acertos;

    const tempoMedioNovo = (tempoPrevTotal + tempoSeg) / Math.max(1, questoesNovo);
    const difMediaNova = ((difMediaPrev * totalBlocosPrev) + difPercebida) / Math.max(1, totalBlocosNovo);
    const flagsNovo = flagsPrev + (flags ? 1 : 0);

    const acertoVida = questoesNovo > 0 ? acertosNovo / questoesNovo : 0;
    const acerto28 = questoes28Novo > 0 ? acertos28Novo / questoes28Novo : acertoVida;
    const acerto7 = questoes7Novo > 0 ? acertos7Novo / questoes7Novo : acerto28;

    const updatedStatRow = [
      area,
      subarea,
      totalBlocosNovo,
      questoesNovo,
      acertosNovo,
      acertoVida,
      acerto28,
      acerto7,
      tempoMedioNovo,
      flagsNovo,
      difMediaNova,
      hoje
    ];
    updateSheetRow(SHEET_NAMES.STATS, statBundle.idx, updatedStatRow);

    const updatedSpacedRow = [
      alvo,
      hoje,
      S_novo,
      difMediaNova,
      proximaRevisao,
      lapsesNovos,
      spacedRow.prioridade || 0
    ];
    updateSheetRow(SHEET_NAMES.SPACED, spacedBundle.idx, updatedSpacedRow);

    const updatedModelRow = [
      alvo,
      thetas[0],
      thetas[1],
      thetas[2],
      S_novo,
      hoje
    ];
    updateSheetRow(SHEET_NAMES.MODEL, modelBundle.idx, updatedModelRow);

    // Marca feito na planilha de hoje
    reviewSheet.getRange(reviewIdx + 2, 5).setValue(1);
    SpreadsheetApp.flush();

    return { ok: true, alvo: alvo, recomputeHint: true };
  } catch (e) {
    return { ok: false, error: e.toString() };
  } finally {
    try {
      lock.releaseLock();
    } catch (err) {
      // Ignora erro ao liberar lock
    }
  }
}

// ============================================================================
// API: CALENDÁRIO DE REVISÕES
// ============================================================================

function apiGetReviewCalendar(days) {
  try {
    days = days || 28;
    const spaced = readSheetData(SHEET_NAMES.SPACED);
    const hoje = new Date();
    hoje.setHours(0, 0, 0, 0);
    
    const calendar = {};
    
    for (let i = 0; i < days; i++) {
      const date = new Date(hoje);
      date.setDate(date.getDate() + i);
      const dateStr = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      calendar[dateStr] = 0;
    }
    
    spaced.forEach(item => {
      const proxRevisao = new Date(item.proximaRevisao);
      proxRevisao.setHours(0, 0, 0, 0);
      const dateStr = Utilities.formatDate(proxRevisao, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      
      if (calendar[dateStr] !== undefined) {
        calendar[dateStr]++;
      }
    });
    
    const result = Object.keys(calendar).map(date => ({
      data: date,
      quantidade: calendar[date]
    }));
    
    return { ok: true, data: result };
  } catch (e) {
    return { ok: false, error: e.toString() };
  }
}

// ============================================================================
// API: REGISTRAR DESFECHO DE REVISÃO
// ============================================================================

function apiLogReviewOutcome(payload) {
  try {
    const lock = LockService.getScriptLock();
    lock.tryLock(30000);
    
    const settings = apiGetSettings();
    const alvo = payload.alvo;
    const tDias = parseFloat(payload.tDias) || 0;
    const acertou = payload.acertou === true || payload.acertou === 1;
    const difPercebida = parseInt(payload.difPercebida) || 3;
    const tempoSeg = parseFloat(payload.tempoSeg) || 0;
    
    // Gravar em REVISAO_LOG
    const hoje = new Date();
    const metaUsada = settings.retentionTarget;
    
    const logRow = [
      hoje,
      alvo,
      tDias,
      metaUsada,
      0, // p_prev (será calculado)
      acertou ? 1 : 0,
      tempoSeg,
      difPercebida,
      payload.flags || '',
      payload.obs || ''
    ];
    
    writeSheetRow(SHEET_NAMES.REVISAO_LOG, logRow);
    
    // Atualizar MODEL e SPACED
    const modelData = readSheetData(SHEET_NAMES.MODEL);
    const spacedData = readSheetData(SHEET_NAMES.SPACED);
    const statsData = readSheetData(SHEET_NAMES.STATS);
    
    let modelRow = modelData.find(m => m.alvo === alvo);
    let spacedRow = spacedData.find(s => s.alvo === alvo);
    const statRow = statsData.find(s => `${s.area}::${s.subarea}` === alvo);
    
    // Inicializar MODEL se não existir
    if (!modelRow) {
      modelRow = {
        alvo: alvo,
        theta0: Math.log(settings.Smin),
        theta1: 0,
        theta2: 0,
        S_atual: settings.Smin,
        ultima_atualizacao: hoje
      };
      const newModelRow = [alvo, modelRow.theta0, modelRow.theta1, modelRow.theta2, modelRow.S_atual, hoje];
      writeSheetRow(SHEET_NAMES.MODEL, newModelRow);
    }
    
    // Inicializar SPACED se não existir
    if (!spacedRow) {
      spacedRow = {
        alvo: alvo,
        ultimaRevisao: null,
        estabilidade: settings.Smin,
        dificuldade_media: 3,
        proximaRevisao: hoje,
        lapses: 0,
        prioridade: 0
      };
      const newSpacedRow = [alvo, null, settings.Smin, 3, hoje, 0, 0];
      writeSheetRow(SHEET_NAMES.SPACED, newSpacedRow);
    }
    
    // Calcular competência
    let competencia = 0.5;
    if (statRow) {
      competencia = parseFloat(statRow.acerto_28d) || parseFloat(statRow.acerto_vida) || 0.5;
    }
    
    const difNorm = normalizeDif(difPercebida);
    
    // Calcular S_hat
    const theta0 = parseFloat(modelRow.theta0);
    const theta1 = parseFloat(modelRow.theta1);
    const theta2 = parseFloat(modelRow.theta2);
    
    const S_hat = calcStability(theta0, theta1, theta2, competencia, difNorm);
    
    // Calcular S_obs
    const S_obs = calcSobs(tDias, metaUsada);
    
    // Atualizar θ (ridge-like)
    const lnS_obs = Math.log(S_obs);
    const lnS_hat = Math.log(S_hat);
    const erro = lnS_obs - lnS_hat;
    
    const x = [1, competencia, difNorm];
    const thetas = [theta0, theta1, theta2];
    
    for (let i = 0; i < 3; i++) {
      thetas[i] = (1 - settings.regLambda) * thetas[i] + settings.lrEta * erro * x[i];
    }
    
    // Aplicar caps em S_obs
    const S_novo = applyCapS(S_obs, settings.Smin, settings.Smax);
    
    // Atualizar MODEL
    const modelSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.MODEL);
    const modelIdx = modelData.findIndex(m => m.alvo === alvo);
    if (modelIdx >= 0) {
      const updatedModelRow = [alvo, thetas[0], thetas[1], thetas[2], S_novo, hoje];
      updateSheetRow(SHEET_NAMES.MODEL, modelIdx, updatedModelRow);
    }
    
    // Atualizar SPACED
    const spacedSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.SPACED);
    const spacedIdx = spacedData.findIndex(s => s.alvo === alvo);
    
    let novaUltimaRevisao = spacedRow.ultimaRevisao;
    let novoLapses = parseInt(spacedRow.lapses) || 0;
    
    if (acertou) {
      novaUltimaRevisao = hoje;
    } else {
      novoLapses++;
    }
    
    // Calcular próximo intervalo
    let I = calcOptimalInterval(S_novo, metaUsada);
    I = applyCapI(I, settings.Imin, settings.Imax);
    
    // Reset suave se acerto_7d < 0.5
    if (statRow && parseFloat(statRow.acerto_7d) < 0.5) {
      I = Math.max(settings.Imin, I * 0.5);
    }
    
    const proximaRevisao = new Date(hoje);
    proximaRevisao.setDate(proximaRevisao.getDate() + Math.round(I));
    
    if (spacedIdx >= 0) {
      const updatedSpacedRow = [
        alvo,
        novaUltimaRevisao,
        S_novo,
        difPercebida,
        proximaRevisao,
        novoLapses,
        0 // prioridade será recalculada
      ];
      updateSheetRow(SHEET_NAMES.SPACED, spacedIdx, updatedSpacedRow);
    }
    
    SpreadsheetApp.flush();
    lock.releaseLock();
    
    return { ok: true };
  } catch (e) {
    return { ok: false, error: e.toString() };
  }
}

// ============================================================================
// API: RECALCULAR (REPROCESSAR MODELO)
// ============================================================================

function apiRecompute() {
  try {
    const lock = LockService.getScriptLock();
    lock.tryLock(60000);
    
    const settings = apiGetSettings();
    const revisaoLog = readSheetData(SHEET_NAMES.REVISAO_LOG);
    const statsData = readSheetData(SHEET_NAMES.STATS);
    
    // Ordenar por data
    revisaoLog.sort((a, b) => new Date(a.data) - new Date(b.data));
    
    // Mapa de modelos
    const models = {};
    
    revisaoLog.forEach(log => {
      const alvo = log.alvo;
      const tDias = parseFloat(log.tDias) || 0;
      const metaUsada = parseFloat(log.metaUsada) || settings.retentionTarget;
      const difPercebida = parseInt(log.difPercebida) || 3;
      
      // Inicializar modelo se não existir
      if (!models[alvo]) {
        models[alvo] = {
          theta0: Math.log(settings.Smin),
          theta1: 0,
          theta2: 0,
          S_atual: settings.Smin
        };
      }
      
      const model = models[alvo];
      const statRow = statsData.find(s => `${s.area}::${s.subarea}` === alvo);
      
      let competencia = 0.5;
      if (statRow) {
        competencia = parseFloat(statRow.acerto_28d) || parseFloat(statRow.acerto_vida) || 0.5;
      }
      
      const difNorm = normalizeDif(difPercebida);
      
      // S_hat
      const S_hat = calcStability(model.theta0, model.theta1, model.theta2, competencia, difNorm);
      
      // S_obs
      const S_obs = calcSobs(tDias, metaUsada);
      
      // Atualizar θ
      const lnS_obs = Math.log(S_obs);
      const lnS_hat = Math.log(S_hat);
      const erro = lnS_obs - lnS_hat;
      
      const x = [1, competencia, difNorm];
      const thetas = [model.theta0, model.theta1, model.theta2];
      
      for (let i = 0; i < 3; i++) {
        thetas[i] = (1 - settings.regLambda) * thetas[i] + settings.lrEta * erro * x[i];
      }
      
      model.theta0 = thetas[0];
      model.theta1 = thetas[1];
      model.theta2 = thetas[2];
      model.S_atual = applyCapS(S_obs, settings.Smin, settings.Smax);
    });
    // Atualizar aba MODEL
    clearSheetData(SHEET_NAMES.MODEL);
    const modelSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.MODEL);
    
    Object.keys(models).forEach(alvo => {
      const m = models[alvo];
      const rowData = [
        alvo,
        m.theta0,
        m.theta1,
        m.theta2,
        m.S_atual,
        new Date()
      ];
      writeSheetRow(SHEET_NAMES.MODEL, rowData);
    });
    
    // Atualizar SPACED com novas estabilidades
    const spacedData = readSheetData(SHEET_NAMES.SPACED);
    const spacedSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.SPACED);
    
    spacedData.forEach((item, idx) => {
      const alvo = item.alvo;
      if (models[alvo]) {
        const S_novo = models[alvo].S_atual;
        let I = calcOptimalInterval(S_novo, settings.retentionTarget);
        I = applyCapI(I, settings.Imin, settings.Imax);
        
        const ultimaRevisao = item.ultimaRevisao ? new Date(item.ultimaRevisao) : new Date();
        const proximaRevisao = new Date(ultimaRevisao);
        proximaRevisao.setDate(proximaRevisao.getDate() + Math.round(I));
        
        const updatedRow = [
          alvo,
          item.ultimaRevisao,
          S_novo,
          item.dificuldade_media,
          proximaRevisao,
          item.lapses,
          item.prioridade
        ];
        updateSheetRow(SHEET_NAMES.SPACED, idx, updatedRow);
      }
    });
    
    SpreadsheetApp.flush();
    lock.releaseLock();
    
    return { ok: true, modelsUpdated: Object.keys(models).length };
  } catch (e) {
    return { ok: false, error: e.toString() };
  }
}