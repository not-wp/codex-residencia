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
  SETTINGS: 'SETTINGS',
  EXAM_CONFIG: 'EXAM_CONFIG',
  POLICY_LOG: 'POLICY_LOG',
  EFFECTS: 'EFFECTS'
};

const HEADERS = {
  LOG: ['data', 'area', 'subarea', 'total', 'acertos', 'tempoMedioSeg', 'difPercebida', 'flags', 'obs', 'uid'],
  STATS: ['area', 'subarea', 'total_blocos', 'questoes', 'acertos', 'acerto_vida', 'acerto_28d', 'acerto_7d', 'tempo_medio', 'flags_28d', 'dif_media', 'ultimaData'],
  SPACED: ['alvo', 'ultimaRevisao', 'estabilidade', 'dificuldade_media', 'proximaRevisao', 'lapses', 'prioridade'],
  REVER_HOJE: ['alvo', 'prioridade', 'proximaRevisao', 'estabilidade', 'feito'],
  MODEL: ['alvo', 'theta0', 'theta1', 'theta2', 'S_atual', 'ultima_atualizacao', 'sigma', 'n_eff'],
  REVISAO_LOG: ['data', 'alvo', 'tDias', 'metaUsada', 'p_prev', 'acertou', 'tempoSeg', 'difPercebida', 'flags', 'obs', 'total', 'acertos'],
  SETTINGS: ['retentionTarget', 'wPeg', 'wTempo', 'wDif', 'alpha', 'overdueMode', 'lrEta', 'regLambda', 'halfLifeDecayDays', 'reviewOutcomeWeight', 'Smin', 'Smax', 'Imin', 'Imax', 'betaUncertainty', 'shrinkageC', 'planGainMix', 'useAdvancedPriority', 'useGainLCB', 'useRLSKalman', 'useDiversityReg', 'useWeibull', 'useBanditPlanner', 'useABTesting'],
  EXAM_CONFIG: ['area', 'peso', 'dataProva'],
  POLICY_LOG: ['timestamp', 'alvo', 'area', 'subarea', 'pri', 'eviPerMin', 'overdue', 'diversity', 'custos', 'tempoPrev', 'decisao', 'policyVersion'],
  EFFECTS: ['alvo', 'ATE_pct', 'lo', 'hi', 'n_pairs', 'updated']
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
  planGainMix: 0.5,
  useAdvancedPriority: false,
  useGainLCB: true,
  useRLSKalman: true,
  useDiversityReg: false,
  useWeibull: false,
  useBanditPlanner: false,
  useABTesting: false
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

  const sanitized = rowData.map(value => {
    if (value instanceof Date && !isNaN(value)) {
      const copy = new Date(value.getTime());
      copy.setHours(0, 0, 0, 0);
      return copy;
    }
    return value;
  });

  sheet.getRange(targetRow, 1, 1, sanitized.length).setValues([sanitized]);

  sanitized.forEach((value, idx) => {
    if (value instanceof Date && !isNaN(value)) {
      sheet.getRange(targetRow, idx + 1).setNumberFormat('dd/mm/yyyy');
    }
  });

  SpreadsheetApp.flush();
}

function updateSheetRow(sheetName, rowIndex, rowData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error(`Sheet ${sheetName} não encontrada`);

  const sanitized = rowData.map(value => {
    if (value instanceof Date && !isNaN(value)) {
      const copy = new Date(value.getTime());
      copy.setHours(0, 0, 0, 0);
      return copy;
    }
    return value;
  });

  // rowIndex é baseado em 0, então +2 (1 para header, 1 para converter de 0-based)
  sheet.getRange(rowIndex + 2, 1, 1, sanitized.length).setValues([sanitized]);

  sanitized.forEach((value, idx) => {
    if (value instanceof Date && !isNaN(value)) {
      sheet.getRange(rowIndex + 2, idx + 1).setNumberFormat('dd/mm/yyyy');
    }
  });

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
  if (dateInput === null || dateInput === undefined || dateInput === '') {
    return null;
  }

  if (dateInput instanceof Date && !isNaN(dateInput)) {
    const copy = new Date(dateInput.getTime());
    copy.setHours(0, 0, 0, 0);
    return copy;
  }

  if (typeof dateInput === 'number' && isFinite(dateInput)) {
    const fromNumber = new Date(dateInput);
    if (!isNaN(fromNumber)) {
      fromNumber.setHours(0, 0, 0, 0);
      return fromNumber;
    }
  }

  if (typeof dateInput === 'string') {
    const trimmed = dateInput.trim();
    if (!trimmed) return null;

    // dd/mm/yyyy
    const brMatch = trimmed.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
    if (brMatch) {
      const day = Number(brMatch[1]);
      const month = Number(brMatch[2]) - 1;
      const year = Number(brMatch[3]);
      const parsed = new Date(year, month, day);
      if (!isNaN(parsed)) {
        parsed.setHours(0, 0, 0, 0);
        return parsed;
      }
    }

    // yyyy-MM-dd or yyyy-MM-ddTHH:MM:SS
    const isoMatch = trimmed.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (isoMatch) {
      const year = Number(isoMatch[1]);
      const month = Number(isoMatch[2]) - 1;
      const day = Number(isoMatch[3]);
      const parsed = new Date(year, month, day);
      if (!isNaN(parsed)) {
        parsed.setHours(0, 0, 0, 0);
        return parsed;
      }
    }

    const fallback = new Date(trimmed);
    if (!isNaN(fallback)) {
      fallback.setHours(0, 0, 0, 0);
      return fallback;
    }
  }

  return null;
}

function parseSheetDate(value) {
  const parsed = parseIsoDateToLocal(value);
  if (parsed) return parsed;
  return null;
}

function formatDateDDMMYYYY(date) {
  const parsed = parseSheetDate(date);
  if (!parsed) return '';
  const day = String(parsed.getDate()).padStart(2, '0');
  const month = String(parsed.getMonth() + 1).padStart(2, '0');
  const year = parsed.getFullYear();
  return `${day}/${month}/${year}`;
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
      return Object.assign({}, DEFAULT_SETTINGS);
    }

    const headers = sheet.getRange(1, 1, 1, HEADERS.SETTINGS.length).getValues()[0];
    const values = sheet.getRange(2, 1, 1, HEADERS.SETTINGS.length).getValues()[0];

    const settings = Object.assign({}, DEFAULT_SETTINGS);
    headers.forEach((header, idx) => {
      settings[header] = values[idx];
    });

    return settings;
  } catch (e) {
    return Object.assign({}, DEFAULT_SETTINGS);
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
    obj.halfLifeDecayDays = Math.max(1, parseFloat(obj.halfLifeDecayDays) || DEFAULT_SETTINGS.halfLifeDecayDays);
    obj.reviewOutcomeWeight = Math.max(0, parseFloat(obj.reviewOutcomeWeight) || DEFAULT_SETTINGS.reviewOutcomeWeight);
    obj.betaUncertainty = parseFloat(obj.betaUncertainty);
    if (!isFinite(obj.betaUncertainty)) obj.betaUncertainty = DEFAULT_SETTINGS.betaUncertainty;
    obj.shrinkageC = parseFloat(obj.shrinkageC);
    if (!isFinite(obj.shrinkageC)) obj.shrinkageC = DEFAULT_SETTINGS.shrinkageC;
    obj.planGainMix = clamp(parseFloat(obj.planGainMix), 0, 1);
    if (!isFinite(obj.planGainMix)) obj.planGainMix = DEFAULT_SETTINGS.planGainMix;
    obj.useAdvancedPriority = asBoolean(obj.useAdvancedPriority);
    obj.useGainLCB = asBoolean(obj.useGainLCB);
    obj.useRLSKalman = asBoolean(obj.useRLSKalman);
    obj.useDiversityReg = asBoolean(obj.useDiversityReg);
    obj.useWeibull = asBoolean(obj.useWeibull);
    obj.useBanditPlanner = asBoolean(obj.useBanditPlanner);
    obj.useABTesting = asBoolean(obj.useABTesting);

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
    
    // Converter data (dd/mm/yyyy)
    let dataObj = parseSheetDate(payload.data);
    if (!dataObj) {
      dataObj = new Date();
    }
    dataObj.setHours(0, 0, 0, 0);
    
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
    
    // Escrever diretamente com formatação de data
    writeSheetRow(SHEET_NAMES.LOG, rowData);

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

function apiProcessLogInternal() {
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
      
      let dataBloco = parseSheetDate(row.data);
      if (!dataBloco) {
        dataBloco = new Date();
      }
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
      if (dataBloco && dataBloco > stat.ultimaData) {
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
        const proximaRevisaoBase = parseSheetDate(s.ultimaData) || new Date();
        const proximaRevisao = new Date(proximaRevisaoBase.getTime());
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
            hoje,
            0.2,
            0
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
  return apiProcessLogInternal();
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

function asBoolean(value) {
  return value === true || value === 'true' || value === 1 || value === '1';
}

function parseAlvoParts(alvo) {
  const parts = (alvo || '').split('::');
  return {
    area: (parts[0] || '').trim(),
    subarea: (parts[1] || '').trim()
  };
}

function identityMatrix(size, scale) {
  const matrix = [];
  for (let i = 0; i < size; i++) {
    const row = [];
    for (let j = 0; j < size; j++) {
      row.push(i === j ? scale : 0);
    }
    matrix.push(row);
  }
  return matrix;
}

function multiplyMatrixVector(matrix, vector) {
  const result = [];
  for (let i = 0; i < matrix.length; i++) {
    let sum = 0;
    for (let j = 0; j < vector.length; j++) {
      sum += (matrix[i][j] || 0) * vector[j];
    }
    result.push(sum);
  }
  return result;
}

function outerProduct(vecA, vecB) {
  const result = [];
  for (let i = 0; i < vecA.length; i++) {
    const row = [];
    for (let j = 0; j < vecB.length; j++) {
      row.push((vecA[i] || 0) * (vecB[j] || 0));
    }
    result.push(row);
  }
  return result;
}

function addMatrices(matA, matB) {
  const result = [];
  for (let i = 0; i < matA.length; i++) {
    const row = [];
    for (let j = 0; j < matA[i].length; j++) {
      row.push((matA[i][j] || 0) + (matB[i][j] || 0));
    }
    result.push(row);
  }
  return result;
}

function scaleMatrix(matrix, scalar) {
  return matrix.map(row => row.map(value => value * scalar));
}

function ensureRlsState(alvo, featureCount, settings) {
  const props = PropertiesService.getDocumentProperties();
  const key = `RLS_${alvo}`;
  let state;
  try {
    const raw = props.getProperty(key);
    if (raw) {
      state = JSON.parse(raw);
    }
  } catch (e) {
    state = null;
  }

  if (!state || !Array.isArray(state.P)) {
    const scale = settings && settings.regLambda ? 1 / Math.max(settings.regLambda, 1e-3) : 10;
    state = {
      P: identityMatrix(featureCount, scale),
      sigma2: 0.04,
      nEff: 0
    };
  }

  return state;
}

function persistRlsState(alvo, state) {
  const props = PropertiesService.getDocumentProperties();
  const key = `RLS_${alvo}`;
  props.setProperty(key, JSON.stringify({
    P: state.P,
    sigma2: state.sigma2,
    nEff: state.nEff
  }));
}

function dotProduct(vecA, vecB) {
  let sum = 0;
  for (let i = 0; i < vecA.length; i++) {
    sum += (vecA[i] || 0) * (vecB[i] || 0);
  }
  return sum;
}

function performLearningStep(alvo, thetaVec, xVec, lnSObs, settings, options) {
  const theta = thetaVec.slice();
  const total = Math.max(1, options && options.total ? options.total : 1);
  const useRls = asBoolean(settings.useRLSKalman);
  let sigma2 = options && options.sigma2 !== undefined ? Math.max(1e-6, options.sigma2) : 0.04;
  let nEff = options && options.nEff !== undefined ? Math.max(0, options.nEff) : 0;
  let rlsState = options && options.state ? options.state : null;

  const lnSHat = dotProduct(theta, xVec);
  const innovation = lnSObs - lnSHat;

  if (useRls) {
    const featureCount = xVec.length;
    rlsState = rlsState || ensureRlsState(alvo, featureCount, settings);
    const halfLife = Math.max(1, Number(settings.halfLifeDecayDays) || 56);
    const forgetting = clamp(Math.pow(2, -1 / halfLife), 0.01, 0.999);

    const scaledP = scaleMatrix(rlsState.P, 1 / forgetting);
    const Px = multiplyMatrixVector(scaledP, xVec);
    const denom = 1 + dotProduct(xVec, Px);
    const gain = Px.map(value => value / denom);

    for (let i = 0; i < theta.length; i++) {
      theta[i] = theta[i] + gain[i] * innovation;
    }

    const adjustment = outerProduct(gain, xVec);
    const newP = [];
    for (let i = 0; i < scaledP.length; i++) {
      const row = [];
      for (let j = 0; j < scaledP[i].length; j++) {
        row.push(scaledP[i][j] - adjustment[i][j]);
      }
      newP.push(row);
    }
    rlsState.P = newP;

    const gainScalar = clamp(dotProduct(xVec, gain), 0, 1);
    sigma2 = (1 - gainScalar) * sigma2 + gainScalar * (innovation * innovation);
    nEff = (1 - forgetting) * nEff + gainScalar;
  } else {
    const weightBase = settings.reviewOutcomeWeight || 1;
    const stepWeight = clamp(weightBase * total, 1, 50);
    const lr = settings.lrEta || 0.05;
    const reg = settings.regLambda || 0;
    for (let i = 0; i < theta.length; i++) {
      theta[i] = (1 - reg) * theta[i] + lr * stepWeight * innovation * xVec[i];
    }
    sigma2 = (1 - reg) * sigma2 + reg * (innovation * innovation);
    nEff = Math.min(1000, nEff + stepWeight);
  }

  const lnSPred = dotProduct(theta, xVec);
  let S_pred = Math.exp(lnSPred);
  S_pred = applyCapS(S_pred, settings.Smin, settings.Smax);

  return {
    theta,
    sigma2,
    nEff,
    S_pred,
    lnSPred,
    innovation,
    state: rlsState
  };
}

function appendPolicyLogEntries(entries) {
  if (!entries || entries.length === 0) {
    return;
  }
  const sheet = getOrCreateSheet(SHEET_NAMES.POLICY_LOG, HEADERS.POLICY_LOG);
  const startRow = sheet.getLastRow() + 1;
  const values = entries.map(entry => HEADERS.POLICY_LOG.map(header => entry[header] !== undefined ? entry[header] : ''));
  sheet.getRange(startRow, 1, values.length, HEADERS.POLICY_LOG.length).setValues(values);
  sheet.getRange(startRow, 1, values.length, 1).setNumberFormat('dd/mm/yyyy hh:mm:ss');
}

function buildPriorityContext(spacedItem, statsRow, settings, referenceDate) {
  if (!spacedItem || !settings) return null;

  const today = new Date(referenceDate || new Date());
  today.setHours(0, 0, 0, 0);

  const msPerDay = 1000 * 60 * 60 * 24;
  const S = Math.max(settings.Smin, parseFloat(spacedItem.estabilidade) || settings.Smin);

  const alvoParts = parseAlvoParts(spacedItem.alvo || '');

  let ultimaRevisaoDias = 0;
  if (spacedItem.ultimaRevisao) {
    const ultima = parseSheetDate(spacedItem.ultimaRevisao);
    if (ultima) {
      ultimaRevisaoDias = Math.max(0, Math.floor((today - ultima) / msPerDay));
    }
  } else if (spacedItem.proximaRevisao) {
    const prox = parseSheetDate(spacedItem.proximaRevisao);
    if (prox) {
      ultimaRevisaoDias = Math.max(0, Math.floor((today - prox) / msPerDay));
    }
  }

  const R_t = Math.exp(-ultimaRevisaoDias / Math.max(1, S));
  const baseRecall = 1 - R_t;

  let peg = 0;
  let tempoRel = 0;
  let difNorm = 0;
  let tempoPrevSeg = 60;
  let competencia = 0.5;

  if (statsRow) {
    const flags28d = parseFloat(statsRow.flags_28d);
    if (!isNaN(flags28d)) {
      peg = clamp(flags28d / 10, 0, 1);
    }

    const tempoMedio = parseFloat(statsRow.tempo_medio);
    if (!isNaN(tempoMedio) && tempoMedio > 0) {
      tempoPrevSeg = tempoMedio;
      tempoRel = clamp(tempoMedio / 120, 0, 1);
    } else {
      tempoRel = clamp(tempoPrevSeg / 120, 0, 1);
    }

    const difMedia = parseFloat(statsRow.dif_media);
    if (!isNaN(difMedia)) {
      difNorm = clamp((difMedia - 1) / 4, 0, 1);
    }

    const acc28 = parseFloat(statsRow.acerto_28d);
    const accVida = parseFloat(statsRow.acerto_vida);
    if (!isNaN(acc28) && acc28 > 0) {
      competencia = clamp(acc28, 0, 1);
    } else if (!isNaN(accVida) && accVida > 0) {
      competencia = clamp(accVida, 0, 1);
    }
  } else {
    const difMedia = parseFloat(spacedItem.dificuldade_media);
    if (!isNaN(difMedia)) {
      difNorm = clamp((difMedia - 1) / 4, 0, 1);
    }
    tempoRel = clamp(tempoPrevSeg / 120, 0, 1);
  }

  let atrasoDias = 0;
  let proximaDate = null;
  if (spacedItem.proximaRevisao) {
    const proxima = parseSheetDate(spacedItem.proximaRevisao);
    if (proxima) {
      proximaDate = proxima;
      atrasoDias = Math.max(0, Math.floor((today - proxima) / msPerDay));
    }
  }

  const overdueRaw = Math.min(1.5, atrasoDias / Math.max(1, S));
  const overdueValue = calcOverdue(atrasoDias, Math.max(1, S), settings.alpha, settings.overdueMode);

  return {
    hoje: today,
    alvo: spacedItem.alvo || '',
    area: alvoParts.area,
    subarea: alvoParts.subarea,
    S,
    ultimaRevisaoDias,
    baseRecall,
    peg,
    tempoRel,
    tempoPrevSeg,
    difNorm,
    atrasoDias,
    overdueRaw,
    overdueValue,
    proximaDate,
    competencia
  };
}

function calculateClassicPriority(context, settings) {
  if (!context) return { score: 0, components: {} };
  const custos =
    settings.wPeg * context.peg +
    settings.wTempo * context.tempoRel +
    settings.wDif * context.difNorm;
  const score = context.baseRecall + custos + context.overdueValue;
  return {
    score,
    components: {
      base: context.baseRecall,
      custos,
      overdue: context.overdueValue
    }
  };
}

function estimateExpectedDeltaS(modelRow, settings, context) {
  if (!modelRow) {
    return Math.max(0.05 * context.S, 0.1);
  }
  const sigma = parseFloat(modelRow.sigma);
  const nEff = parseFloat(modelRow.n_eff);
  const sigmaAbs = isNaN(sigma) ? 0.1 : Math.max(0.01, Math.abs(sigma));
  const effective = isNaN(nEff) ? 1 : Math.max(0.25, nEff);
  const weight = settings.reviewOutcomeWeight || 1;
  const mix = settings.planGainMix || 0.5;
  const scale = clamp(weight * mix / effective, 0.02, 1);
  return clamp(sigmaAbs * context.S * scale, 0.05, context.S * 0.75);
}

function determineHorizonDays(referenceDate, examConfig) {
  const today = new Date(referenceDate || new Date());
  today.setHours(0, 0, 0, 0);
  let horizon = 42;
  if (Array.isArray(examConfig)) {
    examConfig.forEach(row => {
      if (!row || !row.dataProva) return;
      const examDate = parseSheetDate(row.dataProva);
      if (!examDate) return;
      const diff = Math.floor((examDate.getTime() - today.getTime()) / (1000 * 60 * 60 * 24));
      if (diff > 0 && (horizon === null || diff < horizon)) {
        horizon = diff;
      }
    });
  }
  if (horizon === null || !isFinite(horizon) || horizon <= 0) {
    horizon = 42;
  }
  return horizon;
}

function calculateAdvancedPriority(context, settings, extras) {
  if (!context) {
    return { score: 0, components: {} };
  }

  const modelRow = extras && extras.modelRow ? extras.modelRow : null;
  const horizonDays = extras && extras.horizonDays ? Math.max(1, extras.horizonDays) : 42;
  const expectedDeltaS = estimateExpectedDeltaS(modelRow, settings, context);
  const S = Math.max(context.S, 1);
  const derivative = (horizonDays / (S * S)) * Math.exp(-horizonDays / S);
  const deltaR = derivative * expectedDeltaS;
  const tempoMin = Math.max(context.tempoPrevSeg / 60, 0.25);
  let eviPerMin = deltaR / tempoMin;

  if (asBoolean(settings.useGainLCB)) {
    const sigma = modelRow && modelRow.sigma !== undefined ? Math.abs(parseFloat(modelRow.sigma)) || 0 : 0;
    const beta = settings.betaUncertainty || 0;
    eviPerMin -= beta * sigma;
  }

  const custos =
    settings.wPeg * context.peg +
    settings.wTempo * context.tempoRel +
    settings.wDif * context.difNorm;

  const overdueComponent = context.overdueValue;
  let diversityPenalty = 0;
  if (extras && extras.diversityPenalty) {
    diversityPenalty = extras.diversityPenalty;
  }

  const score = eviPerMin + overdueComponent - diversityPenalty + custos;

  return {
    score,
    components: {
      eviPerMin,
      overdue: overdueComponent,
      custos,
      diversity: diversityPenalty,
      tempoPrev: context.tempoPrevSeg,
      deltaR
    }
  };
}

function calculatePriorityForRow(spacedItem, statsRow, settings, referenceDate, extras) {
  const context = buildPriorityContext(spacedItem, statsRow, settings, referenceDate);
  if (!context) {
    return { score: 0, components: {}, context: null };
  }

  let result;
  if (asBoolean(settings.useAdvancedPriority)) {
    result = calculateAdvancedPriority(context, settings, extras);
  } else {
    result = calculateClassicPriority(context, settings);
  }

  return {
    score: result.score,
    components: result.components || {},
    context
  };
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
    const modelData = readSheetData(SHEET_NAMES.MODEL);
    const examConfig = readSheetData(SHEET_NAMES.EXAM_CONFIG);

    const statsMap = {};
    statsData.forEach(row => {
      const key = `${row.area}::${row.subarea}`;
      statsMap[key] = row;
    });

    const modelMap = {};
    modelData.forEach(row => {
      if (!row || !row.alvo) return;
      modelMap[row.alvo] = row;
    });

    const hoje = new Date();
    hoje.setHours(0, 0, 0, 0);
    const horizonDays = determineHorizonDays(hoje, examConfig);
    const useAdvanced = asBoolean(settings.useAdvancedPriority);

    const reviewList = [];
    const priorityUpdates = [];
    const priorityByAlvo = {};
    const priorityCol = HEADERS.SPACED.indexOf('prioridade') + 1;

    const existingToday = readSheetData(SHEET_NAMES.REVER_HOJE);
    const feitoMap = {};
    existingToday.forEach(row => {
      if (row && row.alvo) {
        feitoMap[row.alvo] = row.feito;
      }
    });

    spaced.forEach((item, idx) => {
      if (!item.alvo) {
        priorityUpdates[idx] = { index: idx, alvo: '', value: 0 };
        return;
      }

      const statsRow = statsMap[item.alvo];
      const extras = {
        modelRow: modelMap[item.alvo] || null,
        horizonDays
      };
      const priorityInfo = calculatePriorityForRow(item, statsRow, settings, hoje, extras);
      const prioridade = priorityInfo.score || 0;
      priorityUpdates[idx] = { index: idx, alvo: item.alvo, value: prioridade };
      priorityByAlvo[item.alvo] = prioridade;

      const proxima = item.proximaRevisao ? parseSheetDate(item.proximaRevisao) : null;
      if (!proxima || isNaN(proxima)) {
        return;
      }

      proxima.setHours(0, 0, 0, 0);
      if (proxima <= hoje) {
        reviewList.push({
          alvo: item.alvo,
          prioridade: prioridade,
          prioridadeBase: prioridade,
          proximaRevisao: proxima,
          estabilidade: parseFloat(item.estabilidade) || settings.Smin,
          components: priorityInfo.components || {},
          context: priorityInfo.context || null,
          feito: feitoMap[item.alvo] || '',
          modelRow: extras.modelRow || null
        });
      }
    });

    if (useAdvanced && asBoolean(settings.useDiversityReg)) {
      reviewList.sort((a, b) => (b.prioridade || 0) - (a.prioridade || 0));
      const diversityCount = {};
      reviewList.forEach(item => {
        const areaKey = item.context && item.context.area ? item.context.area : 'Sem área';
        const penalty = (settings.shrinkageC || 0) * (diversityCount[areaKey] || 0);
        item.components = item.components || {};
        item.components.diversity = penalty;
        item.prioridade = (item.prioridadeBase || 0) - penalty;
        diversityCount[areaKey] = (diversityCount[areaKey] || 0) + 1;
        priorityByAlvo[item.alvo] = item.prioridade;
      });
    }

    reviewList.sort((a, b) => (b.prioridade || 0) - (a.prioridade || 0));

    if (spaced.length > 0 && priorityCol > 0 && priorityUpdates.length === spaced.length) {
      const values = spaced.map((row, idx) => {
        const alvo = row.alvo;
        const update = priorityUpdates[idx];
        const fallback = update ? update.value : 0;
        const val = alvo && priorityByAlvo[alvo] !== undefined ? priorityByAlvo[alvo] : fallback;
        return [val];
      });
      spacedSheet.getRange(2, priorityCol, values.length, 1).setValues(values);
    }

    clearSheetData(SHEET_NAMES.REVER_HOJE);
    const rowsToWrite = reviewList.map(item => [
      item.alvo,
      item.prioridade,
      item.proximaRevisao,
      item.estabilidade,
      item.feito || ''
    ]);
    if (rowsToWrite.length > 0) {
      const startRow = reviewSheet.getLastRow() + 1;
      reviewSheet.getRange(startRow, 1, rowsToWrite.length, HEADERS.REVER_HOJE.length).setValues(rowsToWrite);
      reviewSheet.getRange(startRow, 3, rowsToWrite.length, 1).setNumberFormat('dd/mm/yyyy');
    }

    SpreadsheetApp.flush();

    if (reviewList.length > 0) {
      const policyVersion = useAdvanced ? 'advanced_v1' : 'classic_v1';
      const policyEntries = reviewList.map(item => {
        const area = item.context && item.context.area ? item.context.area : parseAlvoParts(item.alvo).area;
        const subarea = item.context && item.context.subarea ? item.context.subarea : parseAlvoParts(item.alvo).subarea;
        const components = item.components || {};
        return {
          timestamp: new Date(),
          alvo: item.alvo,
          area: area,
          subarea: subarea,
          pri: item.prioridade,
          eviPerMin: components.eviPerMin !== undefined ? components.eviPerMin : '',
          overdue: components.overdue !== undefined ? components.overdue : '',
          diversity: components.diversity !== undefined ? components.diversity : '',
          custos: components.custos !== undefined ? components.custos : '',
          tempoPrev: components.tempoPrev !== undefined ? components.tempoPrev : (item.context ? item.context.tempoPrevSeg : ''),
          decisao: 'selected',
          policyVersion: policyVersion
        };
      });
      appendPolicyLogEntries(policyEntries);
    }

    const responseList = reviewList.map(item => ({
      alvo: item.alvo,
      prioridade: item.prioridade,
      proximaRevisao: formatDateDDMMYYYY(item.proximaRevisao),
      estabilidade: item.estabilidade,
      feito: item.feito || ''
    }));

    return { ok: true, count: reviewList.length, data: responseList };

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
        const proxRev = parseSheetDate(item.proximaRevisao);
        if (proxRev) {
          Logger.log('Próxima revisão (processada): ' + proxRev);
          Logger.log('Hoje: ' + hoje);
          Logger.log('Está vencido? ' + (proxRev <= hoje));
        }
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
    const targetKeyDisplay = formatDateDDMMYYYY(targetDate);

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
      const parsed = parseSheetDate(rawDate);
      if (parsed) {
        dateKey = formatDateDDMMYYYY(parsed);
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
        const dateA = parseSheetDate(a.data) || new Date(0);
        const dateB = parseSheetDate(b.data) || new Date(0);
        return dateB - dateA;
      });
    });

    const details = [];

    spacedData.forEach(function(item) {
      if (!item || !item.proximaRevisao) return;

      let prox = item.proximaRevisao;
      let proxDate = parseSheetDate(prox);
      if (!proxDate) {
        return;
      }
      if (proxDate.getTime() !== targetDate.getTime()) return;

      const alvo = item.alvo || '';
      const partes = alvo.split('::');
      const area = (item.area || partes[0] || '').toString().trim();
      const subarea = (item.subarea || partes[1] || '').toString().trim();
      const key = makeKey(area, subarea);

      const statsRow = statsMap[key] || null;
      const history = logsMap[key] || [];

      let ultimaRevisao = '';
      if (item.ultimaRevisao) {
        ultimaRevisao = formatDateDDMMYYYY(item.ultimaRevisao);
      }

      const detalhe = {
        alvo: alvo,
        area: area,
        subarea: subarea,
        prioridade: item.prioridade !== undefined ? Number(item.prioridade) : null,
        estabilidade: item.estabilidade !== undefined ? Number(item.estabilidade) : null,
        proximaRevisao: formatDateDDMMYYYY(proxDate),
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

    const formattedDetails = details.map(function(detail) {
      const history = Array.isArray(detail.history)
        ? detail.history.map(function(entry) {
            const parsedDate = parseSheetDate(entry.data);
            const formattedDate = formatDateDDMMYYYY(parsedDate || entry.data);
            return {
              data: formattedDate || (entry.data || ''),
              total: entry.total,
              acertos: entry.acertos,
              acertoPct: entry.acertoPct
            };
          })
        : [];

      return Object.assign({}, detail, { history: history });
    });

    return { ok: true, date: targetKeyDisplay, revisoes: formattedDetails };
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
      const dateStr = formatDateDDMMYYYY(date);
      calendar[dateStr] = 0;
    }

    spaced.forEach(item => {
      const proxRevisao = parseSheetDate(item.proximaRevisao);
      if (!proxRevisao) return;
      const dateStr = formatDateDDMMYYYY(proxRevisao);

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
        const ultima = parseSheetDate(spacedRow.ultimaRevisao);
        if (ultima) {
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

    let sigmaAtual = 0.2;
    let nEffAtual = 0;
    let rlsState = null;

    if (modelIdx >= 0) {
      const modelRow = modelData[modelIdx];
      theta0 = parseFloat(modelRow.theta0);
      theta1 = parseFloat(modelRow.theta1);
      theta2 = parseFloat(modelRow.theta2);
      const sigmaSheet = parseFloat(modelRow.sigma);
      const nEffSheet = parseFloat(modelRow.n_eff);
      if (!isNaN(sigmaSheet) && sigmaSheet > 0) {
        sigmaAtual = sigmaSheet;
      }
      if (!isNaN(nEffSheet) && nEffSheet >= 0) {
        nEffAtual = nEffSheet;
      }
      if (asBoolean(settings.useRLSKalman)) {
        rlsState = ensureRlsState(alvo, 3, settings);
      }
    } else {
      theta0 = Math.log(settings.Smin);
      theta1 = 0;
      theta2 = 0;
      if (asBoolean(settings.useRLSKalman)) {
        rlsState = ensureRlsState(alvo, 3, settings);
      }
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
    const lnS_obs = Math.log(S_obs);

    const learningResult = performLearningStep(alvo, [theta0, theta1, theta2], x, lnS_obs, settings, {
      total,
      sigma2: sigmaAtual * sigmaAtual,
      nEff: nEffAtual,
      state: rlsState
    });

    theta0 = learningResult.theta[0];
    theta1 = learningResult.theta[1];
    theta2 = learningResult.theta[2];
    const S_pred = learningResult.S_pred;
    sigmaAtual = Math.sqrt(Math.max(1e-6, learningResult.sigma2));
    nEffAtual = learningResult.nEff;
    if (asBoolean(settings.useRLSKalman) && learningResult.state) {
      persistRlsState(alvo, learningResult.state);
    }

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
      const ultima = parseSheetDate(spacedRow.ultimaRevisao);
      if (ultima) {
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

    const examConfig = readSheetData(SHEET_NAMES.EXAM_CONFIG);
    const horizonDays = determineHorizonDays(hojeSemHora, examConfig);
    const prioridadeInfo = calculatePriorityForRow(
      spacedObjForPriority,
      statsRow,
      settings,
      hojeSemHora,
      {
        modelRow: { sigma: sigmaAtual, n_eff: nEffAtual },
        horizonDays
      }
    );
    const prioridade = prioridadeInfo.score;

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
      hoje,
      sigmaAtual,
      nEffAtual
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
  let lock;
  try {
    lock = LockService.getScriptLock();
    lock.tryLock(60000);

    const settings = apiGetSettings();
    const revisaoLog = readSheetData(SHEET_NAMES.REVISAO_LOG);
    const statsData = readSheetData(SHEET_NAMES.STATS);
    const spacedData = readSheetData(SHEET_NAMES.SPACED);

    const statsMap = {};
    statsData.forEach(row => {
      if (!row) return;
      const key = `${row.area}::${row.subarea}`;
      statsMap[key] = row;
    });

    revisaoLog.sort((a, b) => {
      const dateA = parseSheetDate(a.data) || new Date(0);
      const dateB = parseSheetDate(b.data) || new Date(0);
      return dateA - dateB;
    });

    const models = {};
    const rlsStates = {};
    const useRls = asBoolean(settings.useRLSKalman);

    revisaoLog.forEach(log => {
      if (!log || !log.alvo) return;
      const alvo = log.alvo;
      const metaUsada = clamp(parseFloat(log.metaUsada) || settings.retentionTarget, 0.01, 0.99);
      const difPercebida = parseInt(log.difPercebida, 10);
      const difNorm = clamp(((isNaN(difPercebida) ? 3 : difPercebida) - 1) / 4, 0, 1);
      const totalQuestoes = Math.max(1, parseFloat(log.total) || 1);
      const tDias = Math.max(0.25, parseFloat(log.tDias) || 0.25);

      if (!models[alvo]) {
        models[alvo] = {
          theta: [Math.log(settings.Smin), 0, 0],
          sigma2: 0.04,
          nEff: 0,
          S_atual: settings.Smin
        };
      }

      const modelState = models[alvo];
      const statsRow = statsMap[alvo];

      let competencia = 0.5;
      if (statsRow) {
        const acc28 = parseFloat(statsRow.acerto_28d);
        const accVida = parseFloat(statsRow.acerto_vida);
        if (!isNaN(acc28) && acc28 > 0) {
          competencia = clamp(acc28, 0, 1);
        } else if (!isNaN(accVida) && accVida > 0) {
          competencia = clamp(accVida, 0, 1);
        }
      }

      const xVec = [1, competencia, difNorm];
      const S_obs = Math.max(calcSobs(tDias, metaUsada), settings.Smin / 4);
      const lnS_obs = Math.log(S_obs);

      const rlsState = useRls ? (rlsStates[alvo] || ensureRlsState(alvo, xVec.length, settings)) : null;
      const learningResult = performLearningStep(alvo, modelState.theta, xVec, lnS_obs, settings, {
        total: totalQuestoes,
        sigma2: modelState.sigma2,
        nEff: modelState.nEff,
        state: rlsState
      });

      modelState.theta = learningResult.theta;
      modelState.sigma2 = learningResult.sigma2;
      modelState.nEff = learningResult.nEff;
      modelState.S_atual = learningResult.S_pred;
      if (useRls && learningResult.state) {
        rlsStates[alvo] = learningResult.state;
      }
    });

    if (useRls) {
      Object.keys(rlsStates).forEach(alvo => {
        persistRlsState(alvo, rlsStates[alvo]);
      });
    }

    clearSheetData(SHEET_NAMES.MODEL);
    const modelSheet = getOrCreateSheet(SHEET_NAMES.MODEL, HEADERS.MODEL);
    const modelRows = Object.keys(models).map(alvo => {
      const state = models[alvo];
      const sigma = Math.sqrt(Math.max(1e-6, state.sigma2));
      return [
        alvo,
        state.theta[0],
        state.theta[1],
        state.theta[2],
        state.S_atual,
        new Date(),
        sigma,
        state.nEff
      ];
    });
    if (modelRows.length > 0) {
      const startRow = modelSheet.getLastRow() + 1;
      modelSheet.getRange(startRow, 1, modelRows.length, HEADERS.MODEL.length).setValues(modelRows);
    }

    const spacedSheet = getOrCreateSheet(SHEET_NAMES.SPACED, HEADERS.SPACED);
    const examConfig = readSheetData(SHEET_NAMES.EXAM_CONFIG);
    const horizonDays = determineHorizonDays(new Date(), examConfig);

    spacedData.forEach((item, idx) => {
      if (!item || !item.alvo || !models[item.alvo]) return;
      const alvo = item.alvo;
      const modelState = models[alvo];
      const S_novo = modelState.S_atual;
      let I = calcOptimalInterval(S_novo, settings.retentionTarget);
      if (!isFinite(I)) {
        I = settings.Imin;
      }
      I = applyCapI(Math.round(I), settings.Imin, settings.Imax);

      const ultimaRevisao = parseSheetDate(item.ultimaRevisao) || new Date();
      const proximaRevisao = new Date(ultimaRevisao.getTime());
      proximaRevisao.setDate(proximaRevisao.getDate() + I);

      const statsRow = statsMap[alvo];
      const spacedObjForPriority = {
        alvo: alvo,
        ultimaRevisao: ultimaRevisao,
        estabilidade: S_novo,
        dificuldade_media: item.dificuldade_media,
        proximaRevisao: proximaRevisao,
        lapses: item.lapses
      };

      const prioridadeInfo = calculatePriorityForRow(
        spacedObjForPriority,
        statsRow,
        settings,
        new Date(),
        {
          modelRow: { sigma: Math.sqrt(Math.max(1e-6, modelState.sigma2)), n_eff: modelState.nEff },
          horizonDays
        }
      );

      const updatedRow = [
        alvo,
        ultimaRevisao,
        S_novo,
        item.dificuldade_media,
        proximaRevisao,
        item.lapses,
        prioridadeInfo.score
      ];
      updateSheetRow(SHEET_NAMES.SPACED, idx, updatedRow);
    });

    SpreadsheetApp.flush();
    apiMakeReviewToday();

    return { ok: true, modelsUpdated: Object.keys(models).length };
  } catch (e) {
    return { ok: false, error: e.toString() };
  } finally {
    if (lock) {
      try {
        lock.releaseLock();
      } catch (err) {
        // ignore
      }
    }
  }
}