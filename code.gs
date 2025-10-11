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
  REVER_HOJE: ['alvo', 'prioridade', 'proximaRevisao', 'estabilidade'],
  MODEL: ['alvo', 'theta0', 'theta1', 'theta2', 'S_atual', 'ultima_atualizacao'],
  REVISAO_LOG: ['data', 'alvo', 'tDias', 'metaUsada', 'p_prev', 'acertou', 'tempoSeg', 'difPercebida', 'flags', 'obs', 'total', 'acertos'],
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
  }

  if (headers && headers.length > 0) {
    const lastColumn = sheet.getLastColumn();
    if (lastColumn < headers.length) {
      sheet.insertColumnsAfter(Math.max(1, lastColumn), headers.length - lastColumn);
    }

    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    const existingHeaders = headerRange.getValues()[0];
    let needsUpdate = false;
    for (let i = 0; i < headers.length; i++) {
      if (existingHeaders[i] !== headers[i]) {
        needsUpdate = true;
        break;
      }
    }
    if (needsUpdate) {
      headerRange.setValues([headers]);
    }

    if (lastColumn > headers.length) {
      sheet.getRange(1, headers.length + 1, 1, lastColumn - headers.length).clearContent();
    }

    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
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

function parseIsoDateToLocal(dateInput) {
  if (!dateInput && dateInput !== 0) {
    return null;
  }

  if (dateInput instanceof Date && !isNaN(dateInput)) {
    const copy = new Date(dateInput.getTime());
    copy.setHours(0, 0, 0, 0);
    return copy;
  }

  if (typeof dateInput === 'string') {
    const normalized = dateInput.slice(0, 10);
    const parts = normalized.split('-');
    if (parts.length === 3) {
      const year = Number(parts[0]);
      const month = Number(parts[1]) - 1;
      const day = Number(parts[2]);
      if ([year, month, day].every(num => Number.isFinite(num))) {
        const parsed = new Date(year, month, day);
        parsed.setHours(0, 0, 0, 0);
        return parsed;
      }
    }
  }

  return null;
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

function clamp(value, min, max) {
  return Math.min(max, Math.max(min, value));
}

function parseAlvoParts(alvo) {
  const parts = (alvo || '').split('::');
  return {
    area: (parts[0] || '').trim(),
    subarea: (parts[1] || '').trim()
  };
}

function calculatePriorityForRow(spacedItem, statsRow, settings, referenceDate) {
  if (!spacedItem || !settings) return 0;

  const today = new Date(referenceDate || new Date());
  today.setHours(0, 0, 0, 0);

  const msPerDay = 1000 * 60 * 60 * 24;
  const S = Math.max(settings.Smin, parseFloat(spacedItem.estabilidade) || settings.Smin);

  let ultimaRevisaoDias = 0;
  if (spacedItem.ultimaRevisao) {
    const ultima = new Date(spacedItem.ultimaRevisao);
    if (!isNaN(ultima)) {
      ultima.setHours(0, 0, 0, 0);
      ultimaRevisaoDias = Math.max(0, Math.floor((today - ultima) / msPerDay));
    }
  } else if (spacedItem.proximaRevisao) {
    const prox = new Date(spacedItem.proximaRevisao);
    if (!isNaN(prox)) {
      prox.setHours(0, 0, 0, 0);
      ultimaRevisaoDias = Math.max(0, Math.floor((today - prox) / msPerDay));
    }
  }

  const R_t = Math.exp(-ultimaRevisaoDias / S);
  const base = 1 - R_t;

  let peg = 0;
  let tempoRel = 0;
  let difNorm = 0;

  if (statsRow) {
    const flags28d = parseFloat(statsRow.flags_28d) || 0;
    peg = clamp(flags28d / 10, 0, 1);

    const tempoMedio = parseFloat(statsRow.tempo_medio) || 60;
    tempoRel = clamp(tempoMedio / 120, 0, 1);

    const difMedia = parseFloat(statsRow.dif_media);
    if (!isNaN(difMedia)) {
      difNorm = clamp((difMedia - 1) / 4, 0, 1);
    }
  } else {
    const difMedia = parseFloat(spacedItem.dificuldade_media);
    if (!isNaN(difMedia)) {
      difNorm = clamp((difMedia - 1) / 4, 0, 1);
    }
  }

  let atrasoDias = 0;
  if (spacedItem.proximaRevisao) {
    const proxima = new Date(spacedItem.proximaRevisao);
    if (!isNaN(proxima)) {
      proxima.setHours(0, 0, 0, 0);
      atrasoDias = Math.max(0, Math.floor((today - proxima) / msPerDay));
    }
  }

  const overdue = settings.alpha * Math.min(1.5, atrasoDias / S);

  return base +
    settings.wPeg * peg +
    settings.wTempo * tempoRel +
    settings.wDif * difNorm +
    overdue;
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

// ============================================================================
// API: REVISÕES - CRIAR FILA DO DIA
// ============================================================================

function apiMakeReviewToday() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const spacedSheet = ss.getSheetByName(SHEET_NAMES.SPACED);
    const reviewSheet = getOrCreateSheet(SHEET_NAMES.REVER_HOJE, HEADERS.REVER_HOJE);

    if (!spacedSheet) {
      clearSheetData(SHEET_NAMES.REVER_HOJE);
      return { ok: true, count: 0, data: [] };
    }

    const settings = apiGetSettings();
    const spaced = readSheetData(SHEET_NAMES.SPACED);
    const statsSheet = ss.getSheetByName(SHEET_NAMES.STATS);
    const statsData = statsSheet ? readSheetData(SHEET_NAMES.STATS) : [];

    const statsMap = {};
    statsData.forEach(row => {
      const key = `${row.area}::${row.subarea}`;
      statsMap[key] = row;
    });

    const hoje = new Date();
    hoje.setHours(0, 0, 0, 0);

    const reviewList = [];
    const priorityValues = [];
    const priorityCol = HEADERS.SPACED.indexOf('prioridade') + 1;

    spaced.forEach(item => {
      if (!item.alvo) {
        priorityValues.push([0]);
        return;
      }

      const statsRow = statsMap[item.alvo];
      const prioridade = calculatePriorityForRow(item, statsRow, settings, hoje) || 0;
      priorityValues.push([prioridade]);

      const proxima = item.proximaRevisao ? new Date(item.proximaRevisao) : null;
      if (!proxima || isNaN(proxima)) {
        return;
      }

      proxima.setHours(0, 0, 0, 0);
      if (proxima <= hoje) {
        reviewList.push({
          alvo: item.alvo,
          prioridade: prioridade,
          proximaRevisao: Utilities.formatDate(proxima, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
          estabilidade: parseFloat(item.estabilidade) || settings.Smin
        });
      }
    });

    if (spaced.length > 0 && priorityCol > 0 && priorityValues.length === spaced.length) {
      spacedSheet.getRange(2, priorityCol, priorityValues.length, 1).setValues(priorityValues);
    }

    reviewList.sort((a, b) => b.prioridade - a.prioridade);

    clearSheetData(SHEET_NAMES.REVER_HOJE);
    reviewList.forEach(item => {
      reviewSheet.appendRow([
        item.alvo,
        item.prioridade,
        item.proximaRevisao,
        item.estabilidade
      ]);
    });

    SpreadsheetApp.flush();

    return { ok: true, count: reviewList.length, data: reviewList };

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
    if (!dateISO) {
      return { ok: false, error: 'Data inválida' };
    }

    const timezone = Session.getScriptTimeZone();
    let targetDate = parseIsoDateToLocal(dateISO);
    if (!targetDate) {
      const fallback = new Date(dateISO);
      if (!(fallback instanceof Date) || isNaN(fallback)) {
        return { ok: false, error: 'Data inválida' };
      }
      fallback.setHours(0, 0, 0, 0);
      targetDate = fallback;
    }
    const targetKey = Utilities.formatDate(targetDate, timezone, 'yyyy-MM-dd');

    const spacedData = readSheetData(SHEET_NAMES.SPACED);
    const statsData = readSheetData(SHEET_NAMES.STATS);
    const logData = readSheetData(SHEET_NAMES.LOG);

    const makeKey = function(area, subarea) {
      const safeArea = area ? area.toString().trim() : '';
      const safeSub = subarea ? subarea.toString().trim() : '';
      return `${safeArea}::${safeSub}`;
    };

    const statsMap = {};
    statsData.forEach(function(row) {
      const key = makeKey(row.area, row.subarea);
      statsMap[key] = row;
    });

    const logsMap = {};
    logData.forEach(function(entry) {
      const key = makeKey(entry.area, entry.subarea);
      const rawDate = entry.data;
      let dateKey = '';
      if (rawDate instanceof Date && !isNaN(rawDate)) {
        dateKey = Utilities.formatDate(rawDate, timezone, 'yyyy-MM-dd');
      } else if (typeof rawDate === 'string' && rawDate) {
        dateKey = rawDate.slice(0, 10);
      }

      if (!logsMap[key]) {
        logsMap[key] = [];
      }

      const total = Number(entry.total) || 0;
      const acertos = Number(entry.acertos) || 0;
      const pct = total > 0 ? (acertos / total) * 100 : null;

      logsMap[key].push({
        data: dateKey,
        total: total,
        acertos: acertos,
        acertoPct: pct
      });
    });

    Object.keys(logsMap).forEach(function(key) {
      logsMap[key].sort(function(a, b) {
        const safeA = a.data ? a.data : '1970-01-01';
        const safeB = b.data ? b.data : '1970-01-01';
        const dateA = parseIsoDateToLocal(safeA) || new Date(safeA);
        const dateB = parseIsoDateToLocal(safeB) || new Date(safeB);
        return dateB - dateA;
      });
    });

    const details = [];

    spacedData.forEach(function(item) {
      if (!item || !item.proximaRevisao) return;

      let prox = item.proximaRevisao;
      let proxDate = parseIsoDateToLocal(prox);
      if (!proxDate) {
        proxDate = new Date(prox);
        if (!(proxDate instanceof Date) || isNaN(proxDate)) return;
        proxDate.setHours(0, 0, 0, 0);
      }
      const proxKey = Utilities.formatDate(proxDate, timezone, 'yyyy-MM-dd');
      if (proxKey !== targetKey) return;

      const alvo = item.alvo || '';
      const partes = alvo.split('::');
      const area = (item.area || partes[0] || '').toString().trim();
      const subarea = (item.subarea || partes[1] || '').toString().trim();
      const key = makeKey(area, subarea);

      const statsRow = statsMap[key] || null;
      const history = logsMap[key] || [];

      let ultimaRevisao = '';
      if (item.ultimaRevisao instanceof Date && !isNaN(item.ultimaRevisao)) {
        ultimaRevisao = Utilities.formatDate(item.ultimaRevisao, timezone, 'yyyy-MM-dd');
      } else if (typeof item.ultimaRevisao === 'string' && item.ultimaRevisao) {
        ultimaRevisao = item.ultimaRevisao.slice(0, 10);
      }

      const detalhe = {
        alvo: alvo,
        area: area,
        subarea: subarea,
        prioridade: item.prioridade !== undefined ? Number(item.prioridade) : null,
        estabilidade: item.estabilidade !== undefined ? Number(item.estabilidade) : null,
        proximaRevisao: proxKey,
        ultimaRevisao: ultimaRevisao,
        lapses: item.lapses !== undefined ? Number(item.lapses) : 0,
        history: history,
        stats: statsRow
          ? {
              acerto_28d: statsRow.acerto_28d !== undefined ? Number(statsRow.acerto_28d) : null,
              acerto_vida: statsRow.acerto_vida !== undefined ? Number(statsRow.acerto_vida) : null,
              total_blocos: statsRow.total_blocos !== undefined ? Number(statsRow.total_blocos) : null
            }
          : null
      };

      details.push(detalhe);
    });

    details.sort(function(a, b) {
      const pA = isNaN(a.prioridade) ? -Infinity : a.prioridade;
      const pB = isNaN(b.prioridade) ? -Infinity : b.prioridade;
      return pB - pA;
    });

    return { ok: true, date: targetKey, revisoes: details };
  } catch (e) {
    return { ok: false, error: e.toString() };
  }
}

function apiApplyReviewDone(alvos) {
  try {
    return { ok: true };
  } catch (e) {
    return { ok: false, error: e.toString() };
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
  const lock = LockService.getScriptLock();
  const acquired = lock.tryLock(30000);
  if (!acquired) {
    return { ok: false, error: 'Não foi possível obter lock para registrar revisão.' };
  }

  try {
    const settings = apiGetSettings();
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const hoje = new Date();
    const hojeSemHora = new Date(hoje);
    hojeSemHora.setHours(0, 0, 0, 0);

    let alvo = (payload.alvo || '').trim();
    let area = (payload.area || '').trim();
    let subarea = (payload.subarea || '').trim();

    if (!alvo && area && subarea) {
      alvo = `${area}::${subarea}`;
    }

    if (alvo && (!area || !subarea)) {
      const parts = parseAlvoParts(alvo);
      if (!area) area = parts.area;
      if (!subarea) subarea = parts.subarea;
    }

    if (!alvo) {
      throw new Error('Informe o alvo (Área::Subárea) da revisão.');
    }
    if (!area || !subarea) {
      throw new Error('Área e subárea são obrigatórias para registrar a revisão.');
    }

    const total = parseInt(payload.total, 10);
    const acertos = parseInt(payload.acertos, 10);

    if (!isFinite(total) || total <= 0) {
      throw new Error('Total de questões deve ser maior que zero.');
    }
    if (!isFinite(acertos) || acertos < 0) {
      throw new Error('Quantidade de acertos inválida.');
    }
    if (acertos > total) {
      throw new Error('Acertos não podem exceder o total de questões.');
    }

    const tempoSegRaw = parseFloat(payload.tempoSeg);
    const tempoSeg = isFinite(tempoSegRaw) && tempoSegRaw >= 0 ? tempoSegRaw : 0;

    let difPercebida = parseInt(payload.difPercebida, 10);
    if (!isFinite(difPercebida)) difPercebida = 3;
    difPercebida = clamp(difPercebida, 1, 5);

    const flags = payload.flags || '';
    const obs = payload.obs || '';
    const hasPprev = payload.p_prev !== undefined && payload.p_prev !== null && payload.p_prev !== '';
    const pPrev = hasPprev ? parseFloat(payload.p_prev) : '';

    const metaOverride = parseFloat(payload.metaOverride);
    let metaUsada = isFinite(metaOverride) && metaOverride > 0 && metaOverride < 1
      ? metaOverride
      : settings.retentionTarget;
    metaUsada = clamp(metaUsada, 0.01, 0.99);

    const spacedSheet = getOrCreateSheet(SHEET_NAMES.SPACED, HEADERS.SPACED);
    const spacedData = readSheetData(SHEET_NAMES.SPACED);
    const spacedIdx = spacedData.findIndex(row => row.alvo === alvo);
    const spacedRow = spacedIdx >= 0 ? spacedData[spacedIdx] : null;

    let tDias = parseFloat(payload.tDias);
    if (!isFinite(tDias) || tDias <= 0) {
      if (spacedRow && spacedRow.ultimaRevisao) {
        const ultima = new Date(spacedRow.ultimaRevisao);
        if (!isNaN(ultima)) {
          ultima.setHours(0, 0, 0, 0);
          tDias = Math.max(1, Math.floor((hojeSemHora - ultima) / (1000 * 60 * 60 * 24)));
        }
      }
    }
    if (!isFinite(tDias) || tDias <= 0) {
      const baseS = spacedRow ? (parseFloat(spacedRow.estabilidade) || settings.Smin) : settings.Smin;
      const estimativa = calcOptimalInterval(baseS, metaUsada);
      tDias = Math.max(1, Math.round(estimativa));
    }

    const acertouPredominante = total > 0 ? (acertos / total) >= 0.5 : false;

    getOrCreateSheet(SHEET_NAMES.REVISAO_LOG, HEADERS.REVISAO_LOG);
    getOrCreateSheet(SHEET_NAMES.LOG, HEADERS.LOG);

    writeSheetRow(SHEET_NAMES.REVISAO_LOG, [
      hoje,
      alvo,
      tDias,
      metaUsada,
      pPrev,
      acertouPredominante ? 1 : 0,
      tempoSeg,
      difPercebida,
      flags,
      obs,
      total,
      acertos
    ]);

    const uid = Utilities.getUuid();
    writeSheetRow(SHEET_NAMES.LOG, [
      hoje,
      area,
      subarea,
      total,
      acertos,
      tempoSeg,
      difPercebida,
      flags,
      obs,
      uid
    ]);

    const modelSheet = getOrCreateSheet(SHEET_NAMES.MODEL, HEADERS.MODEL);
    const modelData = readSheetData(SHEET_NAMES.MODEL);
    let modelIdx = modelData.findIndex(row => row.alvo === alvo);
    let theta0;
    let theta1;
    let theta2;

    if (modelIdx >= 0) {
      const modelRow = modelData[modelIdx];
      theta0 = parseFloat(modelRow.theta0);
      theta1 = parseFloat(modelRow.theta1);
      theta2 = parseFloat(modelRow.theta2);
    } else {
      theta0 = Math.log(settings.Smin);
      theta1 = 0;
      theta2 = 0;
    }

    const statsData = readSheetData(SHEET_NAMES.STATS);
    const statsRow = statsData.find(row => `${row.area}::${row.subarea}` === alvo);
    let competencia = 0.5;
    if (statsRow) {
      const acc28 = parseFloat(statsRow.acerto_28d);
      const accVida = parseFloat(statsRow.acerto_vida);
      if (isFinite(acc28) && acc28 > 0) {
        competencia = acc28;
      } else if (isFinite(accVida) && accVida > 0) {
        competencia = accVida;
      }
    }
    competencia = clamp(competencia, 0, 1);

    const difNorm = clamp((difPercebida - 1) / 4, 0, 1);
    const x = [1, competencia, difNorm];

    const safeTDias = Math.max(tDias, 0.25);
    const S_obs = Math.max(calcSobs(safeTDias, metaUsada), settings.Smin / 4);

    const lnS_hat = theta0 + theta1 * competencia + theta2 * difNorm;
    const lnS_obs = Math.log(S_obs);
    const erro = lnS_obs - lnS_hat;

    const weightBase = settings.reviewOutcomeWeight || 1;
    const weight = clamp(weightBase * total, 1, 50);

    theta0 = (1 - settings.regLambda) * theta0 + settings.lrEta * weight * erro * x[0];
    theta1 = (1 - settings.regLambda) * theta1 + settings.lrEta * weight * erro * x[1];
    theta2 = (1 - settings.regLambda) * theta2 + settings.lrEta * weight * erro * x[2];

    const lnS_pred = theta0 + theta1 * competencia + theta2 * difNorm;
    const S_pred = applyCapS(Math.exp(lnS_pred), settings.Smin, settings.Smax);

    let I = calcOptimalInterval(S_pred, metaUsada);
    if (!isFinite(I)) {
      I = settings.Imin;
    }
    I = applyCapI(Math.round(I), settings.Imin, settings.Imax);

    const proximaRevisao = new Date(hojeSemHora);
    proximaRevisao.setDate(proximaRevisao.getDate() + I);

    let ultimaRevisaoValor = '';
    if (acertouPredominante) {
      ultimaRevisaoValor = hojeSemHora;
    } else if (spacedRow && spacedRow.ultimaRevisao) {
      const ultima = new Date(spacedRow.ultimaRevisao);
      if (!isNaN(ultima)) {
        ultimaRevisaoValor = ultima;
      }
    }

    const lapsesAnterior = spacedRow ? parseInt(spacedRow.lapses) || 0 : 0;
    const lapsesAtual = acertouPredominante ? lapsesAnterior : lapsesAnterior + 1;

    const difAnterior = spacedRow ? parseFloat(spacedRow.dificuldade_media) : NaN;
    const difMedia = isNaN(difAnterior)
      ? difPercebida
      : clamp(difAnterior * 0.7 + difPercebida * 0.3, 1, 5);

    const spacedObjForPriority = {
      alvo: alvo,
      ultimaRevisao: ultimaRevisaoValor || '',
      estabilidade: S_pred,
      dificuldade_media: difMedia,
      proximaRevisao: proximaRevisao,
      lapses: lapsesAtual
    };

    const prioridade = calculatePriorityForRow(spacedObjForPriority, statsRow, settings, hojeSemHora);

    const spacedRowValues = [
      alvo,
      ultimaRevisaoValor,
      S_pred,
      difMedia,
      proximaRevisao,
      lapsesAtual,
      prioridade
    ];

    if (spacedIdx >= 0) {
      updateSheetRow(SHEET_NAMES.SPACED, spacedIdx, spacedRowValues);
    } else {
      writeSheetRow(SHEET_NAMES.SPACED, spacedRowValues);
    }

    const modelRowValues = [
      alvo,
      theta0,
      theta1,
      theta2,
      S_pred,
      hoje
    ];

    if (modelIdx >= 0) {
      updateSheetRow(SHEET_NAMES.MODEL, modelIdx, modelRowValues);
    } else {
      writeSheetRow(SHEET_NAMES.MODEL, modelRowValues);
    }

    SpreadsheetApp.flush();
    apiMakeReviewToday();

    return {
      ok: true,
      alvo: alvo,
      S: S_pred,
      I: I,
      theta: {
        theta0: theta0,
        theta1: theta1,
        theta2: theta2
      },
      updated: ['REVISAO_LOG', 'LOG', 'MODEL', 'SPACED', 'REVER_HOJE']
    };
  } catch (e) {
    return { ok: false, error: e.toString() };
  } finally {
    SpreadsheetApp.flush();
    lock.releaseLock();
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