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

// ============================================================================
// API: REVISÕES - CRIAR FILA DO DIA
// ============================================================================

function apiMakeReviewToday() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const spacedSheet = ss.getSheetByName(SHEET_NAMES.SPACED);
    const statsSheet = ss.getSheetByName(SHEET_NAMES.STATS);
    
    if (!spacedSheet) {
      return { ok: true, count: 0, data: [] };
    }
    
    const settings = apiGetSettings();
    const spaced = readSheetData(SHEET_NAMES.SPACED);
    const stats = statsSheet ? readSheetData(SHEET_NAMES.STATS) : [];
    
    if (spaced.length === 0) {
      return { ok: true, count: 0, data: [] };
    }
    
    const hoje = new Date();
    hoje.setHours(0, 0, 0, 0);
    
    const reviewList = [];
    
    spaced.forEach(item => {
      try {
        const alvo = item.alvo;
        if (!alvo) return;
        
        const proximaRevisao = new Date(item.proximaRevisao);
        proximaRevisao.setHours(0, 0, 0, 0);
        
        if (isNaN(proximaRevisao.getTime())) return;
        
        if (proximaRevisao <= hoje) {
          const S = parseFloat(item.estabilidade) || settings.Smin;
          const ultimaRevisao = item.ultimaRevisao ? new Date(item.ultimaRevisao) : null;
          
          let t = 0;
          if (ultimaRevisao) {
            ultimaRevisao.setHours(0, 0, 0, 0);
            t = Math.max(0, Math.floor((hoje - ultimaRevisao) / (1000 * 60 * 60 * 24)));
          }
          
          const R_t = Math.exp(-t / S);
          const base = 1 - R_t;
          
          // Buscar stats
          const statRow = stats.find(s => `${s.area}::${s.subarea}` === alvo);
          
          let peg = 0;
          let tempo_rel = 0;
          let dif_norm = 0;
          
          if (statRow) {
            const flags28d = parseFloat(statRow.flags_28d) || 0;
            peg = Math.min(1, flags28d / 10);
            
            const tempoMedio = parseFloat(statRow.tempo_medio) || 60;
            tempo_rel = Math.min(1, tempoMedio / 120);
            
            const difMedia = parseFloat(statRow.dif_media) || 3;
            dif_norm = (difMedia - 1) / 4;
          }
          
          const atrasoDias = Math.max(0, Math.floor((hoje - proximaRevisao) / (1000 * 60 * 60 * 24)));
          const overdue = settings.alpha * Math.min(1.5, atrasoDias / S);
          
          const prioridade = base + 
            settings.wPeg * peg + 
            settings.wTempo * tempo_rel + 
            settings.wDif * dif_norm + 
            overdue;
          
          reviewList.push({
            alvo: alvo,
            prioridade: prioridade,
            proximaRevisao: Utilities.formatDate(proximaRevisao, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
            estabilidade: S,
            feito: 0
          });
        }
      } catch (itemErr) {
        // Skip item com erro
      }
    });
    
    reviewList.sort((a, b) => b.prioridade - a.prioridade);
    
    // Escrever em REVER_HOJE
    const reviewSheet = ss.getSheetByName(SHEET_NAMES.REVER_HOJE);
    if (reviewSheet) {
      clearSheetData(SHEET_NAMES.REVER_HOJE);
      
      reviewList.forEach(item => {
        const rowData = [
          item.alvo,
          item.prioridade,
          item.proximaRevisao,
          item.estabilidade,
          item.feito
        ];
        reviewSheet.appendRow(rowData);
      });
      
      SpreadsheetApp.flush();
    }
    
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
    const reviewToday = readSheetData(SHEET_NAMES.REVER_HOJE);
    return { ok: true, data: reviewToday };
  } catch (e) {
    return { ok: false, error: e.toString() };
  }
}

function apiApplyReviewDone(alvos) {
  try {
    const lock = LockService.getScriptLock();
    lock.tryLock(10000);
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.REVER_HOJE);
    const data = readSheetData(SHEET_NAMES.REVER_HOJE);
    
    data.forEach((row, idx) => {
      if (alvos.includes(row.alvo)) {
        sheet.getRange(idx + 2, 5).setValue(1); // coluna 'feito'
      }
    });
    
    SpreadsheetApp.flush();
    lock.releaseLock();
    
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
  try {
    Logger.log('Iniciando apiLogReviewOutcome...');
    
    const lock = LockService.getScriptLock();
    lock.tryLock(30000);
    
    const settings = apiGetSettings();
    const alvo = payload.alvo;
    const total = parseInt(payload.total) || 1;
    const acertos = parseInt(payload.acertos) || 0;
    const taxaAcerto = parseFloat(payload.taxaAcerto) || (acertos / total);
    const difPercebida = parseInt(payload.difPercebida) || 3;
    const tempoSeg = parseFloat(payload.tempoSeg) || 0;
    
    // Data da revisão
    const dataRevisao = payload.dataRevisao ? new Date(payload.dataRevisao) : new Date();
    dataRevisao.setHours(12, 0, 0, 0);
    
    const hoje = new Date();
    hoje.setHours(0, 0, 0, 0);
    
    // Buscar dados do alvo
    const spacedData = readSheetData(SHEET_NAMES.SPACED);
    const modelData = readSheetData(SHEET_NAMES.MODEL);
    const statsData = readSheetData(SHEET_NAMES.STATS);
    
    let spacedRow = spacedData.find(s => s.alvo === alvo);
    let modelRow = modelData.find(m => m.alvo === alvo);
    const statRow = statsData.find(s => `${s.area}::${s.subarea}` === alvo);
    
    if (!spacedRow) {
      lock.releaseLock();
      return { ok: false, error: 'Alvo não encontrado em SPACED. Lance blocos primeiro.' };
    }
    
    // Calcular tDias (dias desde última revisão)
    let tDias = 0;
    if (spacedRow.ultimaRevisao) {
      const ultimaRev = new Date(spacedRow.ultimaRevisao);
      ultimaRev.setHours(0, 0, 0, 0);
      tDias = Math.max(1, Math.floor((dataRevisao - ultimaRev) / (1000 * 60 * 60 * 24)));
    } else {
      // Se nunca revisou, usar dias desde o lançamento do bloco
      tDias = 1;
    }
    
    const metaUsada = settings.retentionTarget;
    
    // Gravar em REVISAO_LOG
    const logRow = [
      dataRevisao,
      alvo,
      tDias,
      metaUsada,
      0, // p_prev (será calculado)
      taxaAcerto >= 0.7 ? 1 : 0, // considera acerto se >= 70%
      tempoSeg,
      difPercebida,
      payload.flags || '',
      payload.obs || `${acertos}/${total} questões`
    ];
    
    writeSheetRow(SHEET_NAMES.REVISAO_LOG, logRow);
    
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
    
    // Calcular competência
    let competencia = taxaAcerto;
    if (statRow) {
      competencia = parseFloat(statRow.acerto_28d) || parseFloat(statRow.acerto_vida) || taxaAcerto;
    }
    
    const difNorm = (difPercebida - 1) / 4;
    
    // Calcular S_hat
    const theta0 = parseFloat(modelRow.theta0);
    const theta1 = parseFloat(modelRow.theta1);
    const theta2 = parseFloat(modelRow.theta2);
    
    const lnS_hat = theta0 + theta1 * competencia + theta2 * difNorm;
    const S_hat = Math.exp(lnS_hat);
    
    // Calcular S_obs
    const S_obs = tDias / (-Math.log(metaUsada));
    
    // Atualizar θ (ridge-like)
    const lnS_obs = Math.log(Math.max(settings.Smin, S_obs));
    const erro = lnS_obs - lnS_hat;
    
    const x = [1, competencia, difNorm];
    const thetas = [theta0, theta1, theta2];
    
    for (let i = 0; i < 3; i++) {
      thetas[i] = (1 - settings.regLambda) * thetas[i] + settings.lrEta * erro * x[i];
    }
    
    // Aplicar caps em S_obs
    const S_novo = Math.max(settings.Smin, Math.min(settings.Smax, S_obs));
    
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
    
    let novaUltimaRevisao = dataRevisao;
    let novoLapses = parseInt(spacedRow.lapses) || 0;
    
    // Se errou (taxa < 70%), incrementar lapses
    if (taxaAcerto < 0.7) {
      novoLapses++;
    }
    
    // Calcular próximo intervalo
    let I = -S_novo * Math.log(metaUsada);
    I = Math.max(settings.Imin, Math.min(settings.Imax, I));
    
    // Reset suave se performance ruim
    if (taxaAcerto < 0.5) {
      I = Math.max(settings.Imin, I * 0.5);
    }
    
    const proximaRevisao = new Date(dataRevisao);
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
    
    const proximaRevisaoStr = Utilities.formatDate(proximaRevisao, Session.getScriptTimeZone(), 'dd/MM/yyyy');
    
    return { 
      ok: true, 
      proximaRevisao: proximaRevisaoStr,
      estabilidade: S_novo.toFixed(1),
      intervalo: Math.round(I)
    };
    
  } catch (e) {
    Logger.log('ERRO em apiLogReviewOutcome: ' + e.toString());
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