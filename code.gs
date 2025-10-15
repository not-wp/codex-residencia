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
  EFFECTS: 'EFFECTS',
  EXAMS: 'EXAMS',
  STUDY_SESSIONS: 'STUDY_SESSIONS',
  STUDY_REVIEWS: 'STUDY_REVIEWS'
};

const HEADERS = {
  LOG: ['data', 'area', 'subarea', 'total', 'acertos', 'tempoMedioSeg', 'difPercebida', 'flags', 'obs', 'uid'],
  STATS: ['area', 'subarea', 'total_blocos', 'questoes', 'acertos', 'acerto_vida', 'acerto_28d', 'acerto_7d', 'tempo_medio', 'flags_28d', 'dif_media', 'ultimaData'],
  SPACED: ['alvo', 'ultimaRevisao', 'estabilidade', 'dificuldade_media', 'proximaRevisao', 'lapses', 'prioridade'],
  REVER_HOJE: ['alvo', 'prioridade', 'proximaRevisao', 'estabilidade', 'feito'],
  MODEL: ['alvo', 'theta0', 'theta1', 'theta2', 'S_atual', 'ultima_atualizacao', 'sigma', 'n_eff', 'weibull_k'],
  REVISAO_LOG: ['data', 'alvo', 'tDias', 'metaUsada', 'p_prev', 'acertou', 'tempoSeg', 'difPercebida', 'flags', 'obs', 'total', 'acertos'],
  SETTINGS: ['retentionTarget', 'wPeg', 'wTempo', 'wDif', 'alpha', 'overdueMode', 'lrEta', 'regLambda', 'halfLifeDecayDays', 'reviewOutcomeWeight', 'Smin', 'Smax', 'Imin', 'Imax', 'betaUncertainty', 'shrinkageC', 'lambdaDiversity', 'planGainMix', 'flashcardsPerMinBase', 'minD1ReadMin', 'banditEnabledForGuide', 'kappaPriToDelta', 'fatigueFactor', 'lambdaSurprise', 'coverageTarget7d', 'powerAlphaScale', 'powerBetaScale', 'powerDiversityScale', 'maintAlphaScale', 'maintBetaScale', 'maintDiversityScale', 'useAdvancedPriority', 'useGainLCB', 'useRLSKalman', 'useDiversityReg', 'useWeibull', 'useBanditPlanner', 'useABTesting'],
  EXAM_CONFIG: ['area', 'peso', 'dataProva'],
  POLICY_LOG: ['timestamp', 'alvo', 'area', 'subarea', 'pri', 'eviPerMin', 'overdue', 'diversity', 'custos', 'tempoPrev', 'decisao', 'policyVersion'],
  EFFECTS: ['alvo', 'ATE_pct', 'lo', 'hi', 'n_pairs', 'updated'],
  EXAMS: ['id', 'title', 'dateTime', 'hoursPerDay', 'topics', 'importance', 'createdAt', 'updatedAt', 'calendarEventId'],
  STUDY_SESSIONS: ['id', 'examId', 'day', 'sequence', 'start', 'end', 'minutes', 'technique', 'mode', 'topic', 'confusableWith', 'pomodoroType', 'status', 'accuracy', 'difficulty', 'notes', 'deltaRpp', 'createdAt', 'updatedAt'],
  STUDY_REVIEWS: ['id', 'examId', 'sourceSessionId', 'day', 'minutes', 'gap', 'type', 'status', 'notes', 'createdAt', 'updatedAt']
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
  lambdaDiversity: 0.25,
  planGainMix: 0.5,
  flashcardsPerMinBase: 2.0,
  minD1ReadMin: 15,
  banditEnabledForGuide: true,
  kappaPriToDelta: 0.2,
  fatigueFactor: 0.3,
  lambdaSurprise: 0.4,
  coverageTarget7d: 0.25,
  powerAlphaScale: 0.9,
  powerBetaScale: 0.8,
  powerDiversityScale: 0.7,
  maintAlphaScale: 1.2,
  maintBetaScale: 1.2,
  maintDiversityScale: 1.3,
  useAdvancedPriority: false,
  useGainLCB: true,
  useRLSKalman: true,
  useDiversityReg: false,
  useWeibull: false,
  useBanditPlanner: false,
  useABTesting: false
};

const WEIBULL_STATE_PREFIX = 'WEIBULL_SHAPE_';
const AB_STATS_KEY = 'AB_TEST_STATS';
const weibullShapeCache = {};

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

function ensureStudyPlanningSheets() {
  getOrCreateSheet(SHEET_NAMES.EXAMS, HEADERS.EXAMS);
  getOrCreateSheet(SHEET_NAMES.STUDY_SESSIONS, HEADERS.STUDY_SESSIONS);
  getOrCreateSheet(SHEET_NAMES.STUDY_REVIEWS, HEADERS.STUDY_REVIEWS);
}

function generateRowId(prefix) {
  const now = new Date();
  const ts = now.getTime().toString(36);
  const rand = Math.floor(Math.random() * 1e6).toString(36);
  return `${prefix}_${ts}_${rand}`;
}

function parseExamTopicsCell(value) {
  if (!value && value !== 0) return [];
  if (Array.isArray(value)) return value;
  if (typeof value === 'string') {
    const trimmed = value.trim();
    if (!trimmed) return [];
    try {
      const parsed = JSON.parse(trimmed);
      if (Array.isArray(parsed)) return parsed;
    } catch (e) {
      // tentar split simples "area::sub|peso"
      const parts = trimmed.split(',').map(item => item.trim()).filter(Boolean);
      if (parts.length) {
        return parts.map(item => ({ topic: item, weight: 1, confusables: [] }));
      }
    }
  }
  return [];
}

function stringifyExamTopics(topics) {
  if (!topics || !Array.isArray(topics)) return '[]';
  try {
    return JSON.stringify(topics);
  } catch (e) {
    return '[]';
  }
}

function loadExamRecords() {
  ensureStudyPlanningSheets();
  const rows = readSheetData(SHEET_NAMES.EXAMS) || [];
  return rows.map(normalizeExamRecord);
}

function normalizeExamRecord(row) {
  if (!row) return null;
  const id = (row.id || row.ID || '').toString() || generateRowId('exam');
  const title = (row.title || row.nome || '').toString();
  const dateTime = parseExamDateTime(row.dateTime || row.dataProva || row.data || '');
  const hoursPerDay = Number(row.hoursPerDay || row.horasDia || 3) || 3;
  const importance = Number(row.importance || row.peso || 1) || 1;
  const topics = parseExamTopicsCell(row.topics || row.topicos || row.temas);
  const calendarEventId = (row.calendarEventId || row.eventId || '').toString();
  const createdAt = row.createdAt ? parseExamDateTime(row.createdAt) : null;
  const updatedAt = row.updatedAt ? parseExamDateTime(row.updatedAt) : null;
  return {
    id,
    title,
    dateTime,
    hoursPerDay,
    importance,
    topics,
    calendarEventId,
    createdAt,
    updatedAt
  };
}

function persistExamRecord(record) {
  ensureStudyPlanningSheets();
  const sheet = getOrCreateSheet(SHEET_NAMES.EXAMS, HEADERS.EXAMS);
  const data = readSheetData(SHEET_NAMES.EXAMS) || [];
  const rowIndex = data.findIndex(item => item.id === record.id);
  const rowValues = [
    record.id,
    record.title,
    record.dateTime ? formatDateTimeISO(record.dateTime) : '',
    record.hoursPerDay,
    stringifyExamTopics(record.topics),
    record.importance,
    record.createdAt ? formatDateTimeISO(record.createdAt) : formatDateTimeISO(new Date()),
    record.updatedAt ? formatDateTimeISO(record.updatedAt) : formatDateTimeISO(new Date()),
    record.calendarEventId || ''
  ];

  if (rowIndex >= 0) {
    updateSheetRow(SHEET_NAMES.EXAMS, rowIndex, rowValues);
  } else {
    writeSheetRow(SHEET_NAMES.EXAMS, rowValues);
  }
}

function replaceRowsForExam(sheetName, headers, examId, newRows) {
  ensureStudyPlanningSheets();
  const sheet = getOrCreateSheet(sheetName, headers);
  const existing = readSheetData(sheetName) || [];
  const remaining = examId
    ? existing.filter(row => (row.examId || row.examID || row.exameId) !== examId)
    : existing.slice();

  const outputRows = remaining.map(row => headers.map(header => row[header] !== undefined ? row[header] : ''));
  if (Array.isArray(newRows) && newRows.length > 0) {
    newRows.forEach(row => {
      outputRows.push(headers.map(header => row[header] !== undefined ? row[header] : ''));
    });
  }

  clearSheetData(sheetName);
  if (outputRows.length > 0) {
    sheet.getRange(2, 1, outputRows.length, headers.length).setValues(outputRows);
  }
  SpreadsheetApp.flush();
}

function loadStudySessions(examId) {
  ensureStudyPlanningSheets();
  const rows = readSheetData(SHEET_NAMES.STUDY_SESSIONS) || [];
  const filtered = examId ? rows.filter(row => row.examId === examId) : rows;
  return filtered.map(row => ({
    id: row.id || generateRowId('sess'),
    examId: row.examId || examId || '',
    day: row.day || '',
    sequence: Number(row.sequence) || 0,
    start: row.start || '',
    end: row.end || '',
    minutes: Number(row.minutes) || 0,
    technique: row.technique || '',
    mode: row.mode || '',
    topic: row.topic || '',
    confusableWith: row.confusableWith || '',
    pomodoroType: row.pomodoroType || '25-5',
    status: row.status || 'pending',
    accuracy: row.accuracy !== undefined && row.accuracy !== '' ? Number(row.accuracy) : null,
    difficulty: row.difficulty !== undefined && row.difficulty !== '' ? Number(row.difficulty) : null,
    notes: row.notes || '',
    deltaRpp: row.deltaRpp !== undefined && row.deltaRpp !== '' ? Number(row.deltaRpp) : null,
    createdAt: row.createdAt || '',
    updatedAt: row.updatedAt || ''
  }));
}

function loadStudyReviews(examId) {
  ensureStudyPlanningSheets();
  const rows = readSheetData(SHEET_NAMES.STUDY_REVIEWS) || [];
  const filtered = examId ? rows.filter(row => row.examId === examId) : rows;
  return filtered.map(row => ({
    id: row.id || generateRowId('review'),
    examId: row.examId || examId || '',
    sourceSessionId: row.sourceSessionId || '',
    day: row.day || '',
    minutes: Number(row.minutes) || 0,
    gap: Number(row.gap) || 0,
    type: row.type || 'review',
    status: row.status || 'pending',
    notes: row.notes || '',
    createdAt: row.createdAt || '',
    updatedAt: row.updatedAt || ''
  }));
}

function persistStudySessions(examId, sessions) {
  const headers = HEADERS.STUDY_SESSIONS;
  const normalized = (sessions || []).map(session => ({
    id: session.id || generateRowId('sess'),
    examId,
    day: session.day || '',
    sequence: session.sequence || 0,
    start: session.start || '',
    end: session.end || '',
    minutes: session.minutes || 0,
    technique: session.technique || '',
    mode: session.mode || '',
    topic: session.topic || '',
    confusableWith: session.confusableWith || '',
    pomodoroType: session.pomodoroType || '25-5',
    status: session.status || 'pending',
    accuracy: session.accuracy !== undefined && session.accuracy !== null ? session.accuracy : '',
    difficulty: session.difficulty !== undefined && session.difficulty !== null ? session.difficulty : '',
    notes: session.notes || '',
    deltaRpp: session.deltaRpp !== undefined && session.deltaRpp !== null ? session.deltaRpp : '',
    createdAt: session.createdAt ? session.createdAt : formatDateTimeISO(new Date()),
    updatedAt: session.updatedAt ? session.updatedAt : formatDateTimeISO(new Date())
  }));
  replaceRowsForExam(SHEET_NAMES.STUDY_SESSIONS, headers, examId, normalized);
}

function persistStudyReviews(examId, reviews) {
  const headers = HEADERS.STUDY_REVIEWS;
  const normalized = (reviews || []).map(review => ({
    id: review.id || generateRowId('review'),
    examId,
    sourceSessionId: review.sourceSessionId || '',
    day: review.day || '',
    minutes: review.minutes || 0,
    gap: review.gap || 0,
    type: review.type || 'review',
    status: review.status || 'pending',
    notes: review.notes || '',
    createdAt: review.createdAt ? review.createdAt : formatDateTimeISO(new Date()),
    updatedAt: review.updatedAt ? review.updatedAt : formatDateTimeISO(new Date())
  }));
  replaceRowsForExam(SHEET_NAMES.STUDY_REVIEWS, headers, examId, normalized);
}

// ============================================================================
// EXAM PLANNER & SCHEDULER
// ============================================================================

function normalizeExamTopicsList(rawTopics, fallbackTitle) {
  const topics = Array.isArray(rawTopics) ? rawTopics.slice() : [];
  if (!topics.length && fallbackTitle) {
    topics.push({ alvo: fallbackTitle, weight: 1, confusables: [] });
  }
  return topics.map(item => {
    const alvo = (item.alvo || item.topic || '').toString().trim();
    const weightRaw = item.weight !== undefined ? parseFloat(item.weight) : 1;
    const weight = isFinite(weightRaw) && weightRaw > 0 ? weightRaw : 1;
    const confusables = Array.isArray(item.confusables) ? item.confusables.map(String) : [];
    const parts = parseAlvoParts(alvo);
    return {
      alvo,
      area: parts.area,
      subarea: parts.subarea,
      weight,
      confusables
    };
  }).filter(topic => topic.alvo);
}

function determineExamRegime(days) {
  const normalized = Math.max(0, Math.round(days));
  if (normalized <= 2) {
    return {
      recallShare: 0.9,
      reviewShare: 0.1,
      revisionOffsets: [0, 1],
      simuladoOffsets: [],
      reviewLabel: 'curta'
    };
  }
  if (normalized <= 7) {
    const offset = Math.max(1, normalized - 2);
    return {
      recallShare: 0.7,
      reviewShare: 0.3,
      revisionOffsets: [1, 3, 6].filter(off => off <= normalized + 1),
      simuladoOffsets: [offset],
      reviewLabel: 'mixta'
    };
  }
  if (normalized <= 28) {
    const sims = [];
    for (let d = 7; d < normalized; d += 7) {
      sims.push(d);
    }
    return {
      recallShare: 0.6,
      reviewShare: 0.4,
      revisionOffsets: [1, 3, 7, 14].filter(off => off <= normalized + 3),
      simuladoOffsets: sims,
      reviewLabel: 'intermediaria'
    };
  }
  const offsets = [3, 10, 30, 60].filter(off => off <= normalized + 7);
  const sims = [];
  for (let d = 7; d < normalized; d += 7) {
    sims.push(d);
  }
  return {
    recallShare: 0.6,
    reviewShare: 0.4,
    revisionOffsets: offsets.length ? offsets : [7, 21, 45],
    simuladoOffsets: sims,
    reviewLabel: 'longa'
  };
}

function allocateTopicSlots(topics, totalSlots) {
  if (!Array.isArray(topics) || topics.length === 0 || !isFinite(totalSlots) || totalSlots <= 0) {
    return [];
  }
  const weights = topics.map(topic => Math.max(0.5, Number(topic.weight) || 1));
  const totalWeight = weights.reduce((sum, w) => sum + w, 0);
  const baseCounts = topics.map(weight => Math.max(0, Math.round((totalSlots * weight) / Math.max(totalWeight, 1e-6))));
  let assigned = baseCounts.reduce((sum, count) => sum + count, 0);
  const counts = baseCounts.slice();
  while (assigned < totalSlots) {
    let bestIdx = 0;
    let bestWeight = -Infinity;
    for (let i = 0; i < topics.length; i++) {
      const potential = weights[i] - counts[i];
      if (potential > bestWeight) {
        bestWeight = potential;
        bestIdx = i;
      }
    }
    counts[bestIdx] += 1;
    assigned += 1;
  }
  while (assigned > totalSlots) {
    let worstIdx = 0;
    let worstValue = Infinity;
    for (let i = 0; i < topics.length; i++) {
      if (counts[i] > 0 && counts[i] < worstValue) {
        worstValue = counts[i];
        worstIdx = i;
      }
    }
    counts[worstIdx] -= 1;
    assigned -= 1;
  }

  const slots = [];
  let lastIdx = -1;
  let streak = 0;
  for (let slot = 0; slot < totalSlots; slot++) {
    let selected = -1;
    let preferred = -1;
    const lastTopic = lastIdx >= 0 ? topics[lastIdx] : null;
    for (let i = 0; i < topics.length; i++) {
      if (counts[i] <= 0) continue;
      if (lastTopic && Array.isArray(lastTopic.confusables) && lastTopic.confusables.indexOf(topics[i].alvo) !== -1) {
        preferred = i;
        break;
      }
    }
    if (preferred >= 0 && !(preferred === lastIdx && streak >= 2)) {
      selected = preferred;
    }
    if (selected < 0) {
      let bestCount = -1;
      for (let i = 0; i < topics.length; i++) {
        if (counts[i] <= 0) continue;
        if (i === lastIdx && streak >= 2) continue;
        if (counts[i] > bestCount) {
          bestCount = counts[i];
          selected = i;
        }
      }
    }
    if (selected < 0) {
      for (let i = 0; i < topics.length; i++) {
        if (counts[i] > 0) {
          selected = i;
          break;
        }
      }
    }
    if (selected < 0) {
      selected = 0;
    }
    slots.push(selected);
    counts[selected] -= 1;
    if (selected === lastIdx) {
      streak += 1;
    } else {
      lastIdx = selected;
      streak = 1;
    }
  }
  return slots;
}

function rotateTopicsForDay(topics, dayIndex) {
  if (!Array.isArray(topics) || topics.length === 0) {
    return [];
  }
  const len = topics.length;
  if (len === 1) {
    return topics.slice();
  }
  const offsetRaw = isFinite(dayIndex) ? Number(dayIndex) : 0;
  const shift = ((offsetRaw % len) + len) % len;
  if (shift === 0) {
    return topics.slice();
  }
  const rotated = topics.slice(shift).concat(topics.slice(0, shift));
  return rotated;
}

function calculateDaysUntil(date) {
  const target = parseExamDateTime(date);
  if (!target) return 0;
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const cleanTarget = new Date(target.getTime());
  cleanTarget.setHours(0, 0, 0, 0);
  const diff = cleanTarget.getTime() - today.getTime();
  return Math.max(0, Math.round(diff / (1000 * 60 * 60 * 24)));
}

function estimateSessionGain(alvo, minutes, settings, dataCache, logRows, revisaoRows, examConfig, today, cache) {
  if (!alvo || !isFinite(minutes) || minutes <= 0) {
    return { deltaRpp: 0, plan: null };
  }
  const key = `${alvo}::${Math.round(minutes)}`;
  if (cache[key]) {
    return cache[key];
  }
  const prefs = {
    flashcardsPerMinBase: settings.flashcardsPerMinBase !== undefined ? settings.flashcardsPerMinBase : DEFAULT_SETTINGS.flashcardsPerMinBase,
    blockQuestionsTarget: 25,
    readShare: 0.35
  };
  const plan = computeStudyGuidePlanV2(alvo, minutes, prefs, settings, dataCache, logRows, revisaoRows, examConfig, today);
  const delta = plan && plan.planToday && plan.planToday.deltaRpp ? Number(plan.planToday.deltaRpp) : 0;
  const result = { deltaRpp: Math.max(0, delta), plan };
  cache[key] = result;
  return result;
}

function buildExamStudyPlan(exam, options, settings) {
  if (!exam) {
    return { sessions: [], reviews: [], summary: {}, totalDelta: 0, days: 0 };
  }

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const examDate = parseExamDateTime(options && options.dateTime ? options.dateTime : exam.dateTime) || new Date(today.getTime() + 7 * 24 * 60 * 60 * 1000);
  const daysUntil = Math.max(0, calculateDaysUntil(examDate));
  const regime = determineExamRegime(daysUntil);
  const hoursPerDayInput = options && options.hoursPerDay !== undefined ? Number(options.hoursPerDay) : exam.hoursPerDay;
  const hoursPerDay = isFinite(hoursPerDayInput) && hoursPerDayInput > 0 ? hoursPerDayInput : 3;
  const baseMinutes = Math.max(30, Math.round(hoursPerDay * 60));
  const topics = normalizeExamTopicsList(options && options.topics ? options.topics : exam.topics, exam.title);
  const topicCount = topics.length > 0 ? topics.length : 1;

  const dataCache = buildStudyGuideData(settings);
  const logRows = readSheetData(SHEET_NAMES.LOG);
  const revisaoRows = readSheetData(SHEET_NAMES.REVISAO_LOG);
  const examConfig = readSheetData(SHEET_NAMES.EXAM_CONFIG);
  const gainCache = {};

  const msPerDay = 24 * 60 * 60 * 1000;
  const sessions = [];
  const reviews = [];
  const areaSummary = {};
  const dailyPlans = new Map();
  let totalDeltaRaw = 0;

  const ensureDayEntry = (date, dayIndex) => {
    const key = formatDateDDMMYYYY(date);
    if (!dailyPlans.has(key)) {
      const base = new Date(date.getTime());
      base.setHours(0, 0, 0, 0);
      dailyPlans.set(key, {
        day: key,
        dayIndex: isFinite(dayIndex) ? Number(dayIndex) : 0,
        date: base,
        topics: [],
        totalMinutesRaw: 0,
        totalDeltaRaw: 0,
        _topicMap: {}
      });
    }
    return dailyPlans.get(key);
  };

  const addTaskToDay = (entry, topic, minutes, deltaRaw, taskMeta) => {
    if (!entry) return;
    const safeMinutes = isFinite(minutes) ? Number(minutes) : 0;
    const safeDelta = isFinite(deltaRaw) ? Number(deltaRaw) : 0;
    const topicKey = topic && topic.alvo ? topic.alvo : (taskMeta && taskMeta.topicKey ? taskMeta.topicKey : `misc_${entry.topics.length}`);
    const map = entry._topicMap || (entry._topicMap = {});
    let topicEntry = map[topicKey];
    if (!topicEntry) {
      topicEntry = {
        topic: topic && topic.alvo ? topic.alvo : (taskMeta && taskMeta.topicLabel ? taskMeta.topicLabel : 'Sessão'),
        area: topic && topic.area ? topic.area : (taskMeta && taskMeta.area ? taskMeta.area : ''),
        stage: taskMeta && taskMeta.stage ? taskMeta.stage : '',
        status: taskMeta && taskMeta.status ? taskMeta.status : '',
        tasks: [],
        totalMinutesRaw: 0,
        totalDeltaRaw: 0
      };
      map[topicKey] = topicEntry;
      entry.topics.push(topicEntry);
    }
    topicEntry.tasks.push({
      technique: taskMeta && taskMeta.technique ? taskMeta.technique : '',
      minutes: Math.round(Math.max(0, safeMinutes)),
      notes: taskMeta && taskMeta.notes ? taskMeta.notes : '',
      mode: taskMeta && taskMeta.mode ? taskMeta.mode : '',
      deltaRpp: Math.max(0, Math.round(safeDelta * 10) / 10)
    });
    topicEntry.totalMinutesRaw += safeMinutes;
    topicEntry.totalDeltaRaw += Math.max(0, safeDelta);
    entry.totalMinutesRaw += safeMinutes;
    entry.totalDeltaRaw += Math.max(0, safeDelta);
  };

  const scheduleSessionsForTopic = (topic, baseMinutes, gainResult, context) => {
    const contextDate = context && context.currentDate instanceof Date ? context.currentDate : today;
    const dayIndexValue = context && isFinite(context.dayIndex) ? Number(context.dayIndex) : 0;
    let sequenceCounter = context && isFinite(context.sequence) ? Number(context.sequence) : 1;
    const dailyEntry = context && context.dailyEntry ? context.dailyEntry : null;
    const planGuide = gainResult && gainResult.plan ? gainResult.plan : null;
    const todayPlan = planGuide && planGuide.planToday ? planGuide.planToday : null;
    const statusGuide = planGuide && planGuide.status ? planGuide.status : null;
    const stageGuide = planGuide && planGuide.stage ? planGuide.stage : null;
    const tasks = [];

    if (todayPlan) {
      if (statusGuide === 'NOVO') {
        if (todayPlan.readMin > 0) {
          tasks.push({
            technique: 'LEITURA',
            rawMinutes: todayPlan.readMin,
            mode: 'power',
            notes: 'Leitura ativa do conteúdo',
            scheduleReviews: false
          });
        }
        if (todayPlan.flashcardsCreateMin > 0) {
          const cards = Math.round(todayPlan.flashcardsCreate || 0);
          tasks.push({
            technique: 'FLASHCARDS',
            rawMinutes: todayPlan.flashcardsCreateMin,
            mode: 'power',
            notes: cards > 0 ? `Criar ${cards} cards conceituais` : 'Criar cards conceituais',
            scheduleReviews: false
          });
        }
        if (todayPlan.questionsEstMin > 0) {
          const qCount = Math.round(todayPlan.questionsNew || 0);
          tasks.push({
            technique: 'QUESTOES',
            rawMinutes: todayPlan.questionsEstMin,
            mode: 'power',
            notes: qCount > 0 ? `Bloco de ${qCount} questões iniciais` : 'Questões iniciais',
            scheduleReviews: true
          });
        }
      } else {
        if (todayPlan.flashcardsReviewMin > 0) {
          tasks.push({
            technique: 'FLASHCARDS',
            rawMinutes: todayPlan.flashcardsReviewMin,
            mode: 'maintenance',
            notes: 'Revisar flashcards existentes',
            scheduleReviews: false
          });
        }
        if (todayPlan.flashcardsCreateMin > 0) {
          const cards = Math.round(todayPlan.flashcardsCreate || 0);
          tasks.push({
            technique: 'FLASHCARDS',
            rawMinutes: todayPlan.flashcardsCreateMin,
            mode: 'maintenance',
            notes: cards > 0 ? `Criar ${cards} cards de erro` : 'Criar cards de erro',
            scheduleReviews: false
          });
        }
        if (todayPlan.questionsEstMin > 0) {
          const qCount = Math.round(todayPlan.questionsNew || 0);
          const helper = stageGuide === 'S1'
            ? 'Questões pós-D1'
            : (stageGuide === 'S3' ? 'Questões de manutenção' : 'Questões de consolidação');
          tasks.push({
            technique: 'QUESTOES',
            rawMinutes: todayPlan.questionsEstMin,
            mode: stageGuide === 'S3' ? 'maintenance' : 'power',
            notes: qCount > 0 ? `${helper} (${qCount})` : helper,
            scheduleReviews: true
          });
        }
      }
    }

    if (tasks.length === 0) {
      tasks.push({
        technique: 'QUESTOES',
        rawMinutes: baseMinutes,
        mode: 'power',
        notes: 'Prática orientada',
        scheduleReviews: true
      });
    }

    let totalRawMinutes = tasks.reduce((sum, task) => sum + (Number(task.rawMinutes) || 0), 0);
    if (!isFinite(totalRawMinutes) || totalRawMinutes <= 0) {
      totalRawMinutes = tasks.length > 0 ? tasks.length * 10 : 0;
      tasks.forEach(task => { task.rawMinutes = totalRawMinutes > 0 ? totalRawMinutes / tasks.length : 0; });
    }

    let effectiveTotalMinutes = isFinite(baseMinutes) && baseMinutes > 0 ? baseMinutes : totalRawMinutes;
    if (!isFinite(effectiveTotalMinutes) || effectiveTotalMinutes <= 0) {
      effectiveTotalMinutes = totalRawMinutes > 0 ? totalRawMinutes : (tasks.length > 0 ? tasks.length * 10 : 25);
    }
    const minTotal = tasks.length > 0 ? tasks.length * 5 : 0;
    if (effectiveTotalMinutes < minTotal) {
      effectiveTotalMinutes = minTotal;
    }

    const deltaTotalRaw = Math.max(0, Number(gainResult && gainResult.deltaRpp ? gainResult.deltaRpp : 0));
    let accumulatedRaw = 0;
    let allocatedMinutes = 0;

    tasks.forEach((task, idx) => {
      const raw = Number(task.rawMinutes) || 0;
      const share = totalRawMinutes > 0 ? raw / totalRawMinutes : (1 / tasks.length);
      let minutes = Math.max(5, Math.round(share * effectiveTotalMinutes));
      if (idx === tasks.length - 1) {
        const remaining = effectiveTotalMinutes - allocatedMinutes;
        if (isFinite(remaining) && remaining > 0) {
          minutes = Math.max(5, Math.round(remaining));
        }
      }
      allocatedMinutes += minutes;

      let taskDeltaRaw = deltaTotalRaw * share;
      if (idx === tasks.length - 1) {
        taskDeltaRaw = Math.max(0, deltaTotalRaw - accumulatedRaw);
      }
      accumulatedRaw += taskDeltaRaw;
      const roundedDelta = Math.max(0, Math.round(taskDeltaRaw * 10) / 10);

      const sessionId = generateRowId('sess');
      const parts = parseAlvoParts(topic.alvo || '');
      const areaLabel = parts.area || 'Geral';
      if (!areaSummary[areaLabel]) {
        areaSummary[areaLabel] = { allocMin: 0, deltaRaw: 0 };
      }
      areaSummary[areaLabel].allocMin += minutes;
      areaSummary[areaLabel].deltaRaw += Math.max(0, taskDeltaRaw);
      totalDeltaRaw += Math.max(0, taskDeltaRaw);

      addTaskToDay(dailyEntry, topic, minutes, taskDeltaRaw, {
        technique: task.technique,
        notes: task.notes || '',
        mode: task.mode || 'power',
        stage: stageGuide || (statusGuide === 'NOVO' ? 'D1' : ''),
        status: statusGuide || '',
        area: parts.area || '',
        topicLabel: topic.alvo || '',
        topicKey: topic.alvo || ''
      });

      sessions.push({
        id: sessionId,
        examId: exam.id,
        day: formatDateDDMMYYYY(contextDate),
        sequence: dayIndexValue * 100 + sequenceCounter,
        start: '',
        end: '',
        minutes,
        technique: task.technique,
        mode: task.mode || 'power',
        topic: topic.alvo,
        confusableWith: Array.isArray(topic.confusables) ? topic.confusables.join(', ') : '',
        pomodoroType: '25-5',
        status: 'pending',
        accuracy: '',
        difficulty: '',
        notes: task.notes || '',
        createdAt: formatDateTimeISO(new Date()),
        updatedAt: formatDateTimeISO(new Date()),
        deltaRpp: roundedDelta
      });
      sequenceCounter += 1;

      if (task.scheduleReviews && Array.isArray(regime.revisionOffsets)) {
        regime.revisionOffsets.forEach(offset => {
          const gap = Math.max(0, offset);
          const reviewDate = new Date(contextDate.getTime() + gap * msPerDay);
          if (reviewDate > examDate) return;
          const reviewMinutes = gap <= 3 ? 15 : 20;
          reviews.push({
            id: generateRowId('rev'),
            examId: exam.id,
            sourceSessionId: sessionId,
            day: formatDateDDMMYYYY(reviewDate),
            minutes: reviewMinutes,
            gap,
            type: gap <= 1 ? 'curta' : 'revisao',
            status: 'pending',
            notes: '',
            createdAt: formatDateTimeISO(new Date()),
            updatedAt: formatDateTimeISO(new Date())
          });
        });
      }
    });

    return sequenceCounter;
  };

  for (let dayIndex = 0; dayIndex <= daysUntil; dayIndex++) {
    const currentDate = new Date(today.getTime() + dayIndex * msPerDay);
    const dateLabel = formatDateDDMMYYYY(currentDate);
    let remainingMinutes = baseMinutes;
    let sequence = 1;
    const dayEntry = ensureDayEntry(currentDate, dayIndex);
    const dayTopics = rotateTopicsForDay(topics, dayIndex);
    const effectiveTopics = dayTopics.length > 0 ? dayTopics : topics;

    let simMinutes = 0;
    if (regime.simuladoOffsets && regime.simuladoOffsets.indexOf(dayIndex) !== -1) {
      simMinutes = dayIndex === daysUntil ? Math.min(remainingMinutes, 60) : Math.min(remainingMinutes, 90);
      remainingMinutes = Math.max(0, remainingMinutes - simMinutes);
      if (simMinutes > 0) {
        const sessionId = generateRowId('sess');
        const deltaEstimate = simMinutes * 0.05;
        const simTopic = effectiveTopics.length > 0
          ? effectiveTopics[dayIndex % effectiveTopics.length]
          : (topics[dayIndex % topicCount] || topics[0]);
        addTaskToDay(dayEntry, simTopic, simMinutes, deltaEstimate, {
          technique: 'SIMULADO',
          notes: 'Simulado completo',
          mode: 'power',
          stage: 'Simulado',
          status: '',
          topicLabel: simTopic && simTopic.alvo ? simTopic.alvo : 'Simulado',
          area: simTopic && simTopic.area ? simTopic.area : ''
        });
        sessions.push({
          id: sessionId,
          examId: exam.id,
          day: dateLabel,
          sequence: dayIndex * 100 + sequence,
          start: '',
          end: '',
          minutes: simMinutes,
          technique: 'SIMULADO',
          mode: 'power',
          topic: simTopic && simTopic.alvo ? simTopic.alvo : exam.title,
          confusableWith: '',
          pomodoroType: 'simulado',
          status: 'pending',
          accuracy: '',
          difficulty: '',
          notes: '',
          createdAt: formatDateTimeISO(new Date()),
          updatedAt: formatDateTimeISO(new Date()),
          deltaRpp: Math.round(deltaEstimate * 10) / 10
        });
        totalDeltaRaw += Math.max(0, deltaEstimate);
        sequence += 1;
      }
    }

    const effectiveMinutes = remainingMinutes;
    const recallMinutes = Math.max(0, Math.round(effectiveMinutes * regime.recallShare));
    const reviewMinutes = Math.max(0, effectiveMinutes - recallMinutes);

    const recallBlock = 25;
    const reviewBlock = 20;
    const recallSlots = recallMinutes > 0 ? Math.max(1, Math.round(recallMinutes / recallBlock)) : 0;
    const reviewSlots = reviewMinutes > 0 ? Math.max(1, Math.round(reviewMinutes / reviewBlock)) : 0;

    const recallAllocation = allocateTopicSlots(effectiveTopics, recallSlots);
    const reviewAllocation = allocateTopicSlots(effectiveTopics, reviewSlots);

    const topicMinutesMap = new Map();
    const topicOrder = [];

    const addMinutesForTopic = (topic, minutes) => {
      if (!topic || !topic.alvo) return;
      if (!isFinite(minutes) || minutes <= 0) return;
      const key = topic.alvo;
      if (!topicMinutesMap.has(key)) {
        topicMinutesMap.set(key, { topic, minutes: 0 });
        topicOrder.push(key);
      }
      const entry = topicMinutesMap.get(key);
      entry.minutes += minutes;
    };

    recallAllocation.forEach(slotIndex => {
      const topic = effectiveTopics[slotIndex] || effectiveTopics[0];
      addMinutesForTopic(topic, recallBlock);
    });

    reviewAllocation.forEach(slotIndex => {
      const topic = effectiveTopics[slotIndex] || effectiveTopics[0];
      addMinutesForTopic(topic, reviewBlock);
    });

    topicOrder.forEach(alvoKey => {
      const entry = topicMinutesMap.get(alvoKey);
      if (!entry || !entry.topic) return;
      const minutes = Number(entry.minutes) || 0;
      if (minutes <= 0) return;
      const gain = estimateSessionGain(entry.topic.alvo, minutes, settings, dataCache, logRows, revisaoRows, examConfig, today, gainCache);
      sequence = scheduleSessionsForTopic(entry.topic, minutes, gain, {
        currentDate,
        dayIndex,
        sequence,
        dailyEntry: dayEntry
      });
    });
  }

  const dailyArray = Array.from(dailyPlans.values()).sort((a, b) => {
    if (a.date instanceof Date && b.date instanceof Date) {
      if (a.date.getTime() !== b.date.getTime()) {
        return a.date.getTime() - b.date.getTime();
      }
    }
    return (a.dayIndex || 0) - (b.dayIndex || 0);
  });

  const daily = dailyArray.map(entry => {
    const iso = formatDateTimeISO(entry.date);
    const topicsList = entry.topics.map(topic => ({
      topic: topic.topic,
      area: topic.area,
      stage: topic.stage,
      status: topic.status,
      totalMinutes: Math.round(topic.totalMinutesRaw),
      deltaRpp: Math.round(topic.totalDeltaRaw * 10) / 10,
      tasks: topic.tasks.slice()
    })).sort((a, b) => {
      if (b.totalMinutes !== a.totalMinutes) {
        return b.totalMinutes - a.totalMinutes;
      }
      return (a.topic || '').localeCompare(b.topic || '');
    });
    return {
      day: entry.day,
      dayIndex: entry.dayIndex,
      dayIso: iso ? iso.slice(0, 10) : '',
      totalMinutes: Math.round(entry.totalMinutesRaw),
      totalDeltaRpp: Math.round(entry.totalDeltaRaw * 10) / 10,
      topics: topicsList
    };
  });

  const areaList = Object.keys(areaSummary).map(area => ({
    area,
    allocMin: Math.round(areaSummary[area].allocMin),
    deltaRpp: Math.round(areaSummary[area].deltaRaw * 10) / 10
  }));

  sessions.sort((a, b) => {
    const seqA = isFinite(a.sequence) ? Number(a.sequence) : 0;
    const seqB = isFinite(b.sequence) ? Number(b.sequence) : 0;
    if (Math.floor(seqA / 100) !== Math.floor(seqB / 100)) {
      return Math.floor(seqA / 100) - Math.floor(seqB / 100);
    }
    if (seqA !== seqB) {
      return seqA - seqB;
    }
    return (a.topic || '').localeCompare(b.topic || '');
  });

  return {
    sessions,
    reviews,
    summary: {
      areas: areaList,
      totalDelta: Math.round(totalDeltaRaw * 10) / 10,
      totalMinutes: daily.length > 0
        ? Math.round(daily.reduce((sum, entry) => sum + (entry.totalMinutes || 0), 0) / daily.length)
        : Math.round(sessions.reduce((sum, item) => sum + (item.minutes || 0), 0))
    },
    days: daysUntil,
    daily
  };
}

function serializeExamRecord(exam) {
  if (!exam) return null;
  return {
    id: exam.id,
    title: exam.title,
    dateTime: exam.dateTime ? formatDateTimeISO(exam.dateTime) : '',
    dateLabel: exam.dateTime ? formatDateTimeLabel(exam.dateTime) : '',
    hoursPerDay: exam.hoursPerDay,
    topics: exam.topics || [],
    importance: exam.importance,
    calendarEventId: exam.calendarEventId || '',
    createdAt: exam.createdAt ? formatDateTimeISO(exam.createdAt) : '',
    updatedAt: exam.updatedAt ? formatDateTimeISO(exam.updatedAt) : ''
  };
}

function apiGetExams() {
  try {
    ensureStudyPlanningSheets();
    const exams = loadExamRecords().map(serializeExamRecord).filter(Boolean);
    return { ok: true, exams };
  } catch (e) {
    return { ok: false, error: e.toString() };
  }
}

function apiCreateExam(payload) {
  try {
    ensureStudyPlanningSheets();
    const data = payload || {};
    const title = (data.title || '').toString().trim();
    if (!title) {
      return { ok: false, error: 'Título da prova obrigatório' };
    }
    const dateTime = parseExamDateTime(data.dateTime || data.date || data.dateISO);
    if (!dateTime) {
      return { ok: false, error: 'Data/hora da prova inválida' };
    }
    const importanceRaw = data.importance !== undefined ? Number(data.importance) : 1;
    const importance = isFinite(importanceRaw) && importanceRaw > 0 ? importanceRaw : 1;
    const hoursRaw = data.hoursPerDay !== undefined ? Number(data.hoursPerDay) : 3;
    const hoursPerDay = isFinite(hoursRaw) && hoursRaw > 0 ? hoursRaw : 3;
    const topics = normalizeExamTopicsList(data.topics || [], title);

    const record = {
      id: generateRowId('exam'),
      title,
      dateTime,
      hoursPerDay,
      topics,
      importance,
      calendarEventId: '',
      createdAt: new Date(),
      updatedAt: new Date()
    };

    persistExamRecord(record);
    return { ok: true, exam: serializeExamRecord(record) };
  } catch (e) {
    return { ok: false, error: e.toString() };
  }
}

function apiUpdateExam(payload) {
  try {
    ensureStudyPlanningSheets();
    const data = payload || {};
    const examId = (data.examId || data.id || '').toString().trim();
    if (!examId) {
      return { ok: false, error: 'examId obrigatório' };
    }

    const exams = loadExamRecords();
    const exam = exams.find(item => item && item.id === examId);
    if (!exam) {
      return { ok: false, error: 'Prova não encontrada' };
    }

    if (data.title !== undefined) {
      const title = (data.title || '').toString().trim();
      if (title) {
        exam.title = title;
      }
    }

    if (data.dateTime || data.date || data.dateISO) {
      const parsed = parseExamDateTime(data.dateTime || data.date || data.dateISO);
      if (parsed) {
        exam.dateTime = parsed;
      }
    }

    if (data.hoursPerDay !== undefined) {
      const hours = Number(data.hoursPerDay);
      if (isFinite(hours) && hours > 0) {
        exam.hoursPerDay = hours;
      }
    }

    if (data.importance !== undefined) {
      const importance = Number(data.importance);
      if (isFinite(importance) && importance > 0) {
        exam.importance = importance;
      }
    }

    if (Array.isArray(data.topics)) {
      exam.topics = normalizeExamTopicsList(data.topics, exam.title);
    }

    exam.updatedAt = new Date();
    persistExamRecord(exam);

    return { ok: true, exam: serializeExamRecord(exam) };
  } catch (e) {
    return { ok: false, error: e.toString() };
  }
}

function apiPlanExam(payload) {
  try {
    ensureStudyPlanningSheets();
    const data = payload || {};
    const examId = (data.examId || data.id || '').toString().trim();
    if (!examId) {
      return { ok: false, error: 'examId obrigatório' };
    }
    const exams = loadExamRecords();
    const exam = exams.find(item => item && item.id === examId);
    if (!exam) {
      return { ok: false, error: 'Prova não encontrada' };
    }

    if (data.title) {
      exam.title = data.title.toString();
    }
    if (data.dateTime || data.date || data.dateISO) {
      const parsed = parseExamDateTime(data.dateTime || data.date || data.dateISO);
      if (parsed) {
        exam.dateTime = parsed;
      }
    }
    if (data.hoursPerDay !== undefined) {
      const hours = Number(data.hoursPerDay);
      if (isFinite(hours) && hours > 0) {
        exam.hoursPerDay = hours;
      }
    }
    if (Array.isArray(data.topics)) {
      exam.topics = normalizeExamTopicsList(data.topics, exam.title);
    }
    exam.updatedAt = new Date();
    persistExamRecord(exam);

    const plan = buildExamStudyPlan(exam, {
      hoursPerDay: exam.hoursPerDay,
      topics: exam.topics,
      dateTime: exam.dateTime
    }, apiGetSettings());

    persistStudySessions(exam.id, plan.sessions);
    persistStudyReviews(exam.id, plan.reviews);

    return {
      ok: true,
      exam: serializeExamRecord(exam),
      plan: {
        sessions: plan.sessions,
        reviews: plan.reviews,
        summary: plan.summary,
        days: plan.days,
        daily: plan.daily
      }
    };
  } catch (e) {
    return { ok: false, error: e.toString() };
  }
}

function apiSyncExamCalendar(payload) {
  try {
    ensureStudyPlanningSheets();
    const data = payload || {};
    const examId = (data.examId || data.id || '').toString().trim();
    if (!examId) {
      return { ok: false, error: 'examId obrigatório' };
    }
    const exams = loadExamRecords();
    const exam = exams.find(item => item && item.id === examId);
    if (!exam) {
      return { ok: false, error: 'Prova não encontrada' };
    }
    const examDate = exam.dateTime ? new Date(exam.dateTime.getTime()) : new Date();
    examDate.setHours(0, 0, 0, 0);
    const start = new Date(examDate.getTime());
    start.setHours(8, 0, 0, 0);
    const end = new Date(start.getTime() + 2 * 60 * 60 * 1000);

    let eventId = exam.calendarEventId || '';
    let calendarLinked = false;
    const descriptionLines = [
      `Plano gerado no app em ${formatDateDDMMYYYY(new Date())}.`,
      'Revise o cronograma diário na aba Provas para acompanhar sessões e revisões.'
    ];

    try {
      if (typeof CalendarApp !== 'undefined' && CalendarApp.getDefaultCalendar) {
        const calendar = CalendarApp.getDefaultCalendar();
        const eventTitle = `Plano de estudos: ${exam.title}`;
        let event = null;
        if (eventId) {
          try {
            event = CalendarApp.getEventById(eventId);
          } catch (err) {
            event = null;
          }
        }
        if (event) {
          event.setTitle(eventTitle);
          event.setTime(start, end);
          event.setDescription(descriptionLines.join('\n'));
        } else {
          event = calendar.createEvent(eventTitle, start, end, {
            description: descriptionLines.join('\n')
          });
          eventId = event.getId();
        }
        calendarLinked = true;
      }
    } catch (calendarError) {
      // Caso o CalendarApp não esteja disponível ou falhe, seguimos com fallback.
      calendarLinked = false;
    }

    if (!calendarLinked && !eventId) {
      eventId = `CAL_${generateRowId('evt')}`;
    }

    exam.calendarEventId = eventId;
    exam.updatedAt = new Date();
    persistExamRecord(exam);
    const response = serializeExamRecord(exam);
    if (!calendarLinked) {
      response.calendarStatus = 'fallback';
    }
    return { ok: true, exam: response };
  } catch (e) {
    return { ok: false, error: e.toString() };
  }
}

function apiGetSessions(params) {
  try {
    ensureStudyPlanningSheets();
    const examId = params && params.examId ? params.examId.toString().trim() : '';
    const sessions = loadStudySessions(examId).map(session => Object.assign({}, session));
    return { ok: true, sessions };
  } catch (e) {
    return { ok: false, error: e.toString() };
  }
}

function apiUpdateSession(payload) {
  try {
    ensureStudyPlanningSheets();
    const data = payload || {};
    const examId = (data.examId || '').toString().trim();
    const sessionId = (data.id || data.sessionId || '').toString().trim();
    if (!examId || !sessionId) {
      return { ok: false, error: 'examId e id obrigatórios' };
    }
    const sessions = loadStudySessions(examId);
    const index = sessions.findIndex(item => item.id === sessionId);
    if (index === -1) {
      return { ok: false, error: 'Sessão não encontrada' };
    }
    const session = sessions[index];
    if (data.status) {
      session.status = data.status.toString();
    }
    if (data.notes !== undefined) {
      session.notes = data.notes.toString();
    }
    if (data.accuracy !== undefined && data.accuracy !== null && data.accuracy !== '') {
      const acc = Number(data.accuracy);
      if (isFinite(acc)) {
        session.accuracy = acc;
      }
    }
    if (data.difficulty !== undefined && data.difficulty !== null && data.difficulty !== '') {
      const dif = Number(data.difficulty);
      if (isFinite(dif)) {
        session.difficulty = dif;
      }
    }
    session.updatedAt = formatDateTimeISO(new Date());
    sessions[index] = session;
    persistStudySessions(examId, sessions);

    let reviews = loadStudyReviews(examId);
    const related = reviews.filter(item => item.sourceSessionId === sessionId);
    const accuracy = session.accuracy;
    let adjusted = false;
    if (accuracy !== null && accuracy !== undefined && accuracy !== '') {
      if (isFinite(accuracy) && accuracy < 60) {
        const hasImmediate = related.some(item => item.gap <= 1);
        if (!hasImmediate) {
          const reviewId = generateRowId('rev');
          const baseDate = parseExamDateTime(session.day) || parseSheetDate(session.day) || new Date();
          const reviewDate = new Date(baseDate.getTime() + 1 * 24 * 60 * 60 * 1000);
          reviews.push({
            id: reviewId,
            examId,
            sourceSessionId: sessionId,
            day: formatDateDDMMYYYY(reviewDate),
            minutes: 15,
            gap: 1,
            type: 'curta',
            status: 'pending',
            notes: '',
            createdAt: formatDateTimeISO(new Date()),
            updatedAt: formatDateTimeISO(new Date())
          });
          adjusted = true;
        }
      }
      if (isFinite(accuracy) && accuracy > 85) {
        const removable = related
          .filter(item => item.gap <= 3)
          .sort((a, b) => (a.gap || 0) - (b.gap || 0));
        if (removable.length > 0) {
          const removeId = removable[0].id;
          reviews = reviews.filter(item => item.id !== removeId);
          adjusted = true;
        }
      }
    }
    if (adjusted) {
      persistStudyReviews(examId, reviews);
    }
    const refreshedReviews = loadStudyReviews(examId).filter(item => item.sourceSessionId === sessionId);
    return {
      ok: true,
      session,
      reviews: refreshedReviews
    };
  } catch (e) {
    return { ok: false, error: e.toString() };
  }
}

function apiGetReviews(params) {
  try {
    ensureStudyPlanningSheets();
    const examId = params && params.examId ? params.examId.toString().trim() : '';
    const reviews = loadStudyReviews(examId).map(item => Object.assign({}, item));
    return { ok: true, reviews };
  } catch (e) {
    return { ok: false, error: e.toString() };
  }
}

function apiUpdateReview(payload) {
  try {
    ensureStudyPlanningSheets();
    const data = payload || {};
    const examId = (data.examId || '').toString().trim();
    const reviewId = (data.id || data.reviewId || '').toString().trim();
    if (!examId || !reviewId) {
      return { ok: false, error: 'examId e id obrigatórios' };
    }
    const reviews = loadStudyReviews(examId);
    const index = reviews.findIndex(item => item.id === reviewId);
    if (index === -1) {
      return { ok: false, error: 'Revisão não encontrada' };
    }
    const review = reviews[index];
    if (data.status) {
      review.status = data.status.toString();
    }
    if (data.notes !== undefined) {
      review.notes = data.notes.toString();
    }
    if (data.day) {
      review.day = data.day.toString();
    }
    review.updatedAt = formatDateTimeISO(new Date());
    reviews[index] = review;
    persistStudyReviews(examId, reviews);
    return { ok: true, review };
  } catch (e) {
    return { ok: false, error: e.toString() };
  }
}

function clearWeibullCache(area) {
  if (area) {
    delete weibullShapeCache[area];
  } else {
    Object.keys(weibullShapeCache).forEach(key => delete weibullShapeCache[key]);
  }
}

function loadWeibullState(area) {
  if (!area) return { sum: 0, count: 0 };
  const props = PropertiesService.getDocumentProperties();
  const raw = props.getProperty(WEIBULL_STATE_PREFIX + area);
  if (!raw) {
    return { sum: 0, count: 0 };
  }
  try {
    const parsed = JSON.parse(raw);
    if (parsed && typeof parsed === 'object') {
      return {
        sum: isFinite(parsed.sum) ? parsed.sum : 0,
        count: isFinite(parsed.count) ? parsed.count : 0
      };
    }
  } catch (e) {
    // se falhar, ignora e retorna estado vazio
  }
  return { sum: 0, count: 0 };
}

function saveWeibullState(area, state) {
  if (!area) return;
  const props = PropertiesService.getDocumentProperties();
  const payload = {
    sum: isFinite(state.sum) ? state.sum : 0,
    count: isFinite(state.count) ? state.count : 0
  };
  props.setProperty(WEIBULL_STATE_PREFIX + area, JSON.stringify(payload));
}

function computeWeibullShapeFromState(state, settings) {
  const shrink = Math.max(1, parseFloat(settings.shrinkageC) || DEFAULT_SETTINGS.shrinkageC);
  const sum = isFinite(state.sum) ? state.sum : 0;
  const count = isFinite(state.count) ? state.count : 0;
  const mean = (sum + shrink * 1) / (count + shrink);
  return clamp(mean, 0.5, 3.0);
}

function getWeibullShape(area, settings) {
  if (!area) return 1;
  if (weibullShapeCache.hasOwnProperty(area)) {
    return weibullShapeCache[area];
  }
  const state = loadWeibullState(area);
  const shape = computeWeibullShapeFromState(state, settings || DEFAULT_SETTINGS);
  weibullShapeCache[area] = shape;
  return shape;
}

function updateWeibullShape(area, sampleK, settings) {
  if (!area || !isFinite(sampleK)) {
    return;
  }
  const boundedSample = clamp(sampleK, 0.5, 3.0);
  const state = loadWeibullState(area);
  state.sum = (isFinite(state.sum) ? state.sum : 0) + boundedSample;
  state.count = (isFinite(state.count) ? state.count : 0) + 1;
  saveWeibullState(area, state);
  const updatedShape = computeWeibullShapeFromState(state, settings || DEFAULT_SETTINGS);
  weibullShapeCache[area] = updatedShape;
}

function estimateWeibullSample(tDias, meta, lambda) {
  if (!isFinite(tDias) || tDias <= 0) return null;
  if (!isFinite(lambda) || lambda <= 0) return null;
  const clampedMeta = clamp(meta, 0.01, 0.99);
  const numerator = Math.log(-Math.log(clampedMeta));
  const ratio = tDias / lambda;
  if (!isFinite(ratio) || ratio <= 0) return null;
  const denominator = Math.log(ratio);
  if (!isFinite(denominator) || Math.abs(denominator) < 1e-6) return null;
  const sample = numerator / denominator;
  if (!isFinite(sample) || sample <= 0) return null;
  return sample;
}

function loadAbStats() {
  const props = PropertiesService.getDocumentProperties();
  const raw = props.getProperty(AB_STATS_KEY);
  if (!raw) {
    return {
      classic: { count: 0, sum: 0, sumSquares: 0 },
      evi: { count: 0, sum: 0, sumSquares: 0 }
    };
  }
  try {
    const parsed = JSON.parse(raw);
    if (parsed && typeof parsed === 'object') {
      const ensure = variant => ({
        count: isFinite(variant.count) ? variant.count : 0,
        sum: isFinite(variant.sum) ? variant.sum : 0,
        sumSquares: isFinite(variant.sumSquares) ? variant.sumSquares : 0
      });
      return {
        classic: ensure(parsed.classic || {}),
        evi: ensure(parsed.evi || {})
      };
    }
  } catch (e) {
    // se falhar o parse, retorna estado vazio
  }
  return {
    classic: { count: 0, sum: 0, sumSquares: 0 },
    evi: { count: 0, sum: 0, sumSquares: 0 }
  };
}

function saveAbStats(stats) {
  const props = PropertiesService.getDocumentProperties();
  props.setProperty(AB_STATS_KEY, JSON.stringify(stats));
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

function parseExamDateTime(value) {
  if (value === null || value === undefined || value === '') return null;

  if (value instanceof Date && !isNaN(value)) {
    return new Date(value.getTime());
  }

  const str = value.toString().trim();
  if (!str) return null;

  const isoMatch = str.match(/^(\d{4})-(\d{2})-(\d{2})(?:[ T](\d{2})(?::(\d{2}))?)?/);
  if (isoMatch) {
    const year = Number(isoMatch[1]);
    const month = Number(isoMatch[2]) - 1;
    const day = Number(isoMatch[3]);
    const hour = isoMatch[4] !== undefined ? Number(isoMatch[4]) : 0;
    const minute = isoMatch[5] !== undefined ? Number(isoMatch[5]) : 0;
    const parsed = new Date(year, month, day, hour, minute, 0, 0);
    if (!isNaN(parsed)) return parsed;
  }

  const brMatch = str.match(/^(\d{2})\/(\d{2})\/(\d{4})(?:\s+(\d{2}):(\d{2}))?$/);
  if (brMatch) {
    const day = Number(brMatch[1]);
    const month = Number(brMatch[2]) - 1;
    const year = Number(brMatch[3]);
    const hour = brMatch[4] !== undefined ? Number(brMatch[4]) : 0;
    const minute = brMatch[5] !== undefined ? Number(brMatch[5]) : 0;
    const parsed = new Date(year, month, day, hour, minute, 0, 0);
    if (!isNaN(parsed)) return parsed;
  }

  const fallback = new Date(str);
  if (!isNaN(fallback)) return fallback;
  return null;
}

function formatDateTimeISO(date) {
  if (!(date instanceof Date) || isNaN(date)) return '';
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  const hour = String(date.getHours()).padStart(2, '0');
  const minute = String(date.getMinutes()).padStart(2, '0');
  return `${year}-${month}-${day}T${hour}:${minute}`;
}

function formatDateTimeLabel(date) {
  if (!(date instanceof Date) || isNaN(date)) return '';
  return date.toLocaleString('pt-BR', {
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit'
  });
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
    obj.lambdaDiversity = parseFloat(obj.lambdaDiversity);
    if (!isFinite(obj.lambdaDiversity)) obj.lambdaDiversity = DEFAULT_SETTINGS.lambdaDiversity;
    obj.planGainMix = clamp(parseFloat(obj.planGainMix), 0, 1);
    if (!isFinite(obj.planGainMix)) obj.planGainMix = DEFAULT_SETTINGS.planGainMix;
    obj.kappaPriToDelta = parseFloat(obj.kappaPriToDelta);
    if (!isFinite(obj.kappaPriToDelta)) obj.kappaPriToDelta = DEFAULT_SETTINGS.kappaPriToDelta;
    obj.kappaPriToDelta = Math.max(0, obj.kappaPriToDelta);
    obj.fatigueFactor = parseFloat(obj.fatigueFactor);
    if (!isFinite(obj.fatigueFactor)) obj.fatigueFactor = DEFAULT_SETTINGS.fatigueFactor;
    obj.fatigueFactor = clamp(obj.fatigueFactor, 0, 1);
    obj.lambdaSurprise = parseFloat(obj.lambdaSurprise);
    if (!isFinite(obj.lambdaSurprise) || obj.lambdaSurprise < 0) obj.lambdaSurprise = DEFAULT_SETTINGS.lambdaSurprise;
    obj.coverageTarget7d = parseFloat(obj.coverageTarget7d);
    if (!isFinite(obj.coverageTarget7d)) obj.coverageTarget7d = DEFAULT_SETTINGS.coverageTarget7d;
    obj.coverageTarget7d = clamp(obj.coverageTarget7d, 0, 1);
    obj.powerAlphaScale = parseFloat(obj.powerAlphaScale);
    if (!isFinite(obj.powerAlphaScale) || obj.powerAlphaScale <= 0) obj.powerAlphaScale = DEFAULT_SETTINGS.powerAlphaScale;
    obj.powerBetaScale = parseFloat(obj.powerBetaScale);
    if (!isFinite(obj.powerBetaScale) || obj.powerBetaScale <= 0) obj.powerBetaScale = DEFAULT_SETTINGS.powerBetaScale;
    obj.powerDiversityScale = parseFloat(obj.powerDiversityScale);
    if (!isFinite(obj.powerDiversityScale) || obj.powerDiversityScale <= 0) obj.powerDiversityScale = DEFAULT_SETTINGS.powerDiversityScale;
    obj.maintAlphaScale = parseFloat(obj.maintAlphaScale);
    if (!isFinite(obj.maintAlphaScale) || obj.maintAlphaScale <= 0) obj.maintAlphaScale = DEFAULT_SETTINGS.maintAlphaScale;
    obj.maintBetaScale = parseFloat(obj.maintBetaScale);
    if (!isFinite(obj.maintBetaScale) || obj.maintBetaScale <= 0) obj.maintBetaScale = DEFAULT_SETTINGS.maintBetaScale;
    obj.maintDiversityScale = parseFloat(obj.maintDiversityScale);
    if (!isFinite(obj.maintDiversityScale) || obj.maintDiversityScale <= 0) obj.maintDiversityScale = DEFAULT_SETTINGS.maintDiversityScale;
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
            0,
            1
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

function calcRecall(t, S, shape) {
  const lambda = Math.max(S, 1e-6);
  const tPos = Math.max(0, t);
  const k = isFinite(shape) && shape > 0 ? shape : 1;
  if (k !== 1) {
    const ratio = tPos / lambda;
    return Math.exp(-Math.pow(ratio, k));
  }
  return Math.exp(-tPos / lambda);
}

function calcOptimalInterval(S, retentionTarget, shape) {
  const meta = clamp(retentionTarget, 0.01, 0.99);
  if (isFinite(shape) && shape > 0 && Math.abs(shape - 1) > 1e-6) {
    const factor = Math.pow(-Math.log(meta), 1 / shape);
    return S * factor;
  }
  return -S * Math.log(meta);
}

function calcStability(theta0, theta1, theta2, competencia, difNorm) {
  // ln(S) = θ0 + θ1·competência + θ2·difNorm
  const lnS = theta0 + theta1 * competencia + theta2 * difNorm;
  return Math.exp(lnS);
}

function calcSobs(t, retentionTarget, shape) {
  const meta = clamp(retentionTarget, 0.01, 0.99);
  if (isFinite(shape) && shape > 0 && Math.abs(shape - 1) > 1e-6) {
    const denom = Math.pow(-Math.log(meta), 1 / shape);
    if (!isFinite(denom) || denom <= 0) {
      return t / (-Math.log(meta));
    }
    return t / denom;
  }
  return t / (-Math.log(meta));
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
  const useWeibull = asBoolean(settings.useWeibull);
  const weibullK = useWeibull ? getWeibullShape(alvoParts.area, settings) : 1;

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

  const R_t = calcRecall(ultimaRevisaoDias, Math.max(1, S), weibullK);
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
    competencia,
    weibullK,
    useWeibull
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

function estimateDeltaSStats(context, modelRow, settings) {
  const baseMean = Math.max(0.05 * context.S, 0.1);
  let mean = baseMean;
  let variance = Math.pow(baseMean * 0.5, 2);

  if (modelRow) {
    const sigma = parseFloat(modelRow.sigma);
    const nEff = parseFloat(modelRow.n_eff);
    const sigmaAbs = isNaN(sigma) ? 0.2 : Math.max(0.01, Math.abs(sigma));
    const effective = isNaN(nEff) ? 1 : Math.max(0.25, nEff);
    const weight = settings.reviewOutcomeWeight || 1;
    const mix = settings.planGainMix || 0.5;
    const scale = clamp(weight * mix / effective, 0.02, 1);
    mean = clamp(sigmaAbs * context.S * scale, 0.05, context.S * 0.75);
    variance = Math.max(variance, Math.pow(mean * 0.5, 2));
  }

  if (asBoolean(settings.useRLSKalman) && context && context.alvo) {
    try {
      const xVec = [1, clamp(context.competencia || 0.5, 0, 1), clamp(context.difNorm || 0, 0, 1)];
      const state = ensureRlsState(context.alvo, xVec.length, settings);
      if (state && Array.isArray(state.P)) {
        const Px = multiplyMatrixVector(state.P, xVec);
        const leverage = Math.max(0, dotProduct(xVec, Px));
        let sigma2State = state.sigma2;
        if (!isFinite(sigma2State) || sigma2State <= 0) {
          const sigmaFallback = modelRow && modelRow.sigma !== undefined ? Math.max(0.01, Math.abs(parseFloat(modelRow.sigma))) : 0.2;
          sigma2State = sigmaFallback * sigmaFallback;
        }
        const varLnS = Math.max(0, leverage * sigma2State);
        const baseS = Math.max(1, context.S);
        const varS = Math.pow(baseS, 2) * varLnS;
        if (isFinite(varS) && varS > 0) {
          variance = Math.max(variance, varS);
        }
      }
    } catch (err) {
      // Se algo falhar, mantém variância básica
    }
  }

  const minVar = Math.pow(baseMean * 0.25, 2);
  const maxVar = Math.pow(context.S * 0.75, 2);
  variance = clamp(variance, minVar, maxVar);

  return { mean, variance };
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
  const deltaStats = estimateDeltaSStats(context, modelRow, settings);
  const S = Math.max(context.S, 1);
  const k = context.useWeibull ? (context.weibullK || 1) : 1;
  let derivative;
  if (context.useWeibull && k !== 1) {
    const ratio = Math.max(1e-6, horizonDays / S);
    const powTerm = Math.pow(ratio, k);
    const recallAtHorizon = Math.exp(-powTerm);
    derivative = recallAtHorizon * (k / S) * powTerm;
  } else {
    derivative = (horizonDays / (S * S)) * Math.exp(-horizonDays / S);
  }

  const deltaRMean = derivative * (deltaStats.mean || 0);
  const deltaRVar = derivative * derivative * Math.max(0, deltaStats.variance || 0);
  const tempoMin = Math.max(context.tempoPrevSeg / 60, 0.25);
  const totalStd = Math.sqrt(Math.max(0, deltaRVar));
  const totalMean = deltaRMean;
  let totalLCB = totalMean;
  let nEffModel = 0;

  if (asBoolean(settings.useGainLCB)) {
    const betaBase = settings.betaUncertainty || 0;
    if (extras && extras.modelRow && extras.modelRow.n_eff !== undefined) {
      const nEffRaw = parseFloat(extras.modelRow.n_eff);
      if (isFinite(nEffRaw) && nEffRaw >= 0) {
        nEffModel = nEffRaw;
      }
    }
    const betaScale = Math.sqrt(Math.max(1, nEffModel));
    const beta = betaScale > 0 ? betaBase / betaScale : betaBase;
    totalLCB = totalMean - beta * totalStd;
  }

  const eviPerMin = totalLCB / tempoMin;
  const eviPerMinMean = totalMean / tempoMin;
  const sdPerMin = totalStd / tempoMin;

  const custos =
    settings.wPeg * context.peg +
    settings.wTempo * context.tempoRel +
    settings.wDif * context.difNorm;

  const overdueComponent = context.overdueValue;
  let diversityBoost = 0;
  if (extras && extras.diversityBoost) {
    diversityBoost = extras.diversityBoost;
  }

  const score = eviPerMin + overdueComponent + diversityBoost + custos;

  return {
    score,
    components: {
      eviPerMin,
      eviPerMinMean,
      eviStdPerMin: sdPerMin,
      eviTotalMean: totalMean,
      eviTotalLCB: totalLCB,
      eviStdTotal: totalStd,
      overdue: overdueComponent,
      custos,
      diversity: diversityBoost,
      tempoPrev: context.tempoPrevSeg,
      costMinutes: tempoMin,
      deltaR: totalMean,
      weibullK: k,
      nEff: nEffModel
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

  let finalScore = result.score;
  const components = Object.assign({}, result.components || {});
  const lambdaSurpriseRaw = extras && extras.surpriseLambda !== undefined
    ? parseFloat(extras.surpriseLambda)
    : parseFloat(settings.lambdaSurprise);
  const lambdaSurprise = isFinite(lambdaSurpriseRaw) ? Math.max(0, lambdaSurpriseRaw) : 0;
  if (lambdaSurprise > 0) {
    const residual = extras && extras.residualValue !== undefined ? extras.residualValue : null;
    if (residual !== null && residual !== undefined && isFinite(residual) && residual < 0) {
      const bonus = lambdaSurprise * Math.max(0, -residual);
      if (bonus > 0) {
        finalScore += bonus;
        components.surprise = bonus;
      }
    }
  }

  if (extras && extras.coverage7d !== undefined) {
    components.coverage7d = extras.coverage7d;
    if (extras.meta7d !== undefined) {
      components.coverageTarget = extras.meta7d;
    }
  }

  return {
    score: finalScore,
    components,
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

function computeSurpriseResidualMap(rows) {
  const map = {};
  try {
    if (!Array.isArray(rows) || rows.length === 0) {
      return map;
    }
    const byAlvo = {};
    rows.forEach(entry => {
      if (!entry || !entry.alvo) return;
      const alvo = entry.alvo.toString();
      const dateObj = parseSheetDate(entry.data) || parseIsoDateToLocal(entry.data);
      const timestamp = dateObj instanceof Date && !isNaN(dateObj) ? dateObj.getTime() : 0;
      let predicted = parseFloat(entry.p_prev);
      if (!isFinite(predicted)) {
        const meta = parseFloat(entry.metaUsada);
        if (isFinite(meta)) {
          predicted = meta;
        }
      }
      if (!isFinite(predicted)) {
        predicted = 0.85;
      }
      predicted = clamp(predicted, 0, 1);
      let observed = null;
      const total = parseFloat(entry.total);
      const acertos = parseFloat(entry.acertos);
      if (isFinite(total) && total > 0 && isFinite(acertos)) {
        observed = clamp(acertos / total, 0, 1);
      } else {
        const acertou = parseFloat(entry.acertou);
        if (isFinite(acertou)) {
          observed = clamp(acertou, 0, 1);
        }
      }
      if (observed === null) return;
      const residual = predicted - observed;
      if (!byAlvo[alvo]) {
        byAlvo[alvo] = [];
      }
      byAlvo[alvo].push({ residual, timestamp });
    });

    const windowSize = 5;
    Object.keys(byAlvo).forEach(alvo => {
      const samples = byAlvo[alvo];
      samples.sort((a, b) => (b.timestamp || 0) - (a.timestamp || 0));
      const slice = samples.slice(0, windowSize);
      if (slice.length === 0) return;
      const avg = slice.reduce((sum, item) => sum + (item.residual || 0), 0) / slice.length;
      map[alvo] = avg;
    });
  } catch (err) {
    Logger.log('computeSurpriseResidualMap error: ' + err);
  }
  return map;
}

function computeCoverage7d(logData, referenceDate, settings) {
  const result = {
    fractions: {},
    total: 0,
    meta: clamp(parseFloat(settings.coverageTarget7d) || DEFAULT_SETTINGS.coverageTarget7d, 0, 1)
  };
  try {
    if (!Array.isArray(logData) || logData.length === 0) {
      return result;
    }
    const today = new Date(referenceDate || new Date());
    today.setHours(0, 0, 0, 0);
    const cutoff = new Date(today.getTime() - 6 * 24 * 60 * 60 * 1000);
    const areaCounts = {};
    let total = 0;
    logData.forEach(row => {
      if (!row || !row.area) return;
      const data = parseSheetDate(row.data) || parseIsoDateToLocal(row.data);
      if (!(data instanceof Date) || isNaN(data)) return;
      data.setHours(0, 0, 0, 0);
      if (data < cutoff || data > today) return;
      const area = row.area.toString().trim();
      if (!area) return;
      total += 1;
      areaCounts[area] = (areaCounts[area] || 0) + 1;
    });
    if (total <= 0) {
      return result;
    }
    result.total = total;
    Object.keys(areaCounts).forEach(area => {
      result.fractions[area] = areaCounts[area] / total;
    });
  } catch (err) {
    Logger.log('computeCoverage7d error: ' + err);
  }
  return result;
}

function gatherReviewCandidates(settings, referenceDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const spacedSheet = ss.getSheetByName(SHEET_NAMES.SPACED);
  const reviewSheet = getOrCreateSheet(SHEET_NAMES.REVER_HOJE, HEADERS.REVER_HOJE);
  if (!spacedSheet) {
    return {
      spacedSheet: null,
      reviewSheet,
      spacedData: [],
      reviewList: [],
      priorityUpdates: [],
      priorityByAlvo: {},
      feitoMap: {},
      useAdvanced: asBoolean(settings.useAdvancedPriority),
      horizonDays: determineHorizonDays(referenceDate, []),
      priorityCol: HEADERS.SPACED.indexOf('prioridade') + 1
    };
  }

  const spacedData = readSheetData(SHEET_NAMES.SPACED);
  const statsSheet = ss.getSheetByName(SHEET_NAMES.STATS);
  const statsData = statsSheet ? readSheetData(SHEET_NAMES.STATS) : [];
  const logData = readSheetData(SHEET_NAMES.LOG);
  const revisaoLogData = readSheetData(SHEET_NAMES.REVISAO_LOG);
  const modelData = readSheetData(SHEET_NAMES.MODEL);
  const examConfig = readSheetData(SHEET_NAMES.EXAM_CONFIG);

  const statsMap = {};
  statsData.forEach(row => {
    if (!row) return;
    const key = `${row.area}::${row.subarea}`;
    statsMap[key] = row;
  });

  const modelMap = {};
  modelData.forEach(row => {
    if (!row || !row.alvo) return;
    modelMap[row.alvo] = row;
  });

  const hoje = new Date(referenceDate || new Date());
  hoje.setHours(0, 0, 0, 0);
  const horizonDays = determineHorizonDays(hoje, examConfig);
  const useAdvanced = asBoolean(settings.useAdvancedPriority);
  const residualMap = computeSurpriseResidualMap(revisaoLogData);
  const coverageInfo = computeCoverage7d(logData, hoje, settings);

  const priorityCol = HEADERS.SPACED.indexOf('prioridade') + 1;
  const existingToday = readSheetData(SHEET_NAMES.REVER_HOJE);
  const feitoMap = {};
  existingToday.forEach(row => {
    if (row && row.alvo) {
      feitoMap[row.alvo] = row.feito;
    }
  });

  const reviewList = [];
  const priorityUpdates = [];
  const priorityByAlvo = {};

  spacedData.forEach((item, idx) => {
    if (!item || !item.alvo) {
      priorityUpdates[idx] = { index: idx, alvo: '', value: 0 };
      return;
    }

    const statsRow = statsMap[item.alvo];
    const alvoParts = parseAlvoParts(item.alvo || '');
    const areaKeyRaw = (alvoParts.area || '').toString().trim() || 'Sem área';
    const extras = {
      modelRow: modelMap[item.alvo] || null,
      horizonDays,
      residualValue: residualMap.hasOwnProperty(item.alvo) ? residualMap[item.alvo] : null,
      surpriseLambda: settings.lambdaSurprise,
      coverage7d: coverageInfo.fractions && coverageInfo.fractions.hasOwnProperty(areaKeyRaw)
        ? coverageInfo.fractions[areaKeyRaw]
        : 0,
      meta7d: coverageInfo.meta
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
    const areaKey = priorityInfo.context && priorityInfo.context.area
      ? priorityInfo.context.area
      : areaKeyRaw;
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
        modelRow: extras.modelRow || null,
        areaKey
      });
    }
  });

  if (useAdvanced && asBoolean(settings.useDiversityReg) && reviewList.length > 0) {
    const lambdaDiv = parseFloat(settings.lambdaDiversity);
    const diversityWeight = isFinite(lambdaDiv) ? lambdaDiv : DEFAULT_SETTINGS.lambdaDiversity;
    if (diversityWeight !== 0) {
      const areaCounts = {};
      reviewList.forEach(item => {
        const key = item.areaKey || 'Sem área';
        areaCounts[key] = (areaCounts[key] || 0) + 1;
      });
      const totalItems = reviewList.length;
      const coverageReal = {};
      Object.keys(areaCounts).forEach(area => {
        coverageReal[area] = areaCounts[area] / totalItems;
      });

      const targetsRaw = {};
      let totalPeso = 0;
      examConfig.forEach(row => {
        if (!row || !row.area) return;
        const peso = parseFloat(row.peso);
        if (!isNaN(peso) && peso > 0) {
          targetsRaw[row.area] = (targetsRaw[row.area] || 0) + peso;
          totalPeso += peso;
        }
      });

      const uniqueAreas = Array.from(new Set(reviewList.map(item => item.areaKey || 'Sem área')));
      const targetShares = {};
      if (totalPeso > 0) {
        Object.keys(targetsRaw).forEach(area => {
          targetShares[area] = targetsRaw[area] / totalPeso;
        });
        const configuredSum = Object.keys(targetShares).reduce((sum, area) => sum + targetShares[area], 0);
        const remainingAreas = uniqueAreas.filter(area => !targetShares.hasOwnProperty(area));
        const remainingShare = Math.max(0, 1 - configuredSum);
        const defaultShare = remainingAreas.length > 0 ? remainingShare / remainingAreas.length : 0;
        remainingAreas.forEach(area => {
          targetShares[area] = defaultShare;
        });
      } else if (uniqueAreas.length > 0) {
        const equalShare = 1 / uniqueAreas.length;
        uniqueAreas.forEach(area => {
          targetShares[area] = equalShare;
        });
      }

      reviewList.forEach(item => {
        const area = item.areaKey || 'Sem área';
        const target = targetShares.hasOwnProperty(area)
          ? targetShares[area]
          : (uniqueAreas.length > 0 ? 1 / uniqueAreas.length : 0);
        const atual = coverageReal[area] || 0;
        let deficit = Math.max(0, target - atual);
        const coverageArea = coverageInfo.fractions && coverageInfo.fractions.hasOwnProperty(area)
          ? coverageInfo.fractions[area]
          : 0;
        const meta7d = coverageInfo && coverageInfo.meta !== undefined ? coverageInfo.meta : null;
        if (meta7d !== null && isFinite(meta7d)) {
          deficit += Math.max(0, meta7d - coverageArea);
        }
        const boost = diversityWeight * deficit;
        item.components = item.components || {};
        item.components.diversity = boost;
        item.components.coverage7d = coverageArea;
        item.prioridade = (item.prioridadeBase || 0) + boost;
        priorityByAlvo[item.alvo] = item.prioridade;
      });
    }
  }

  reviewList.sort((a, b) => (b.prioridade || 0) - (a.prioridade || 0));

  return {
    spacedSheet,
    reviewSheet,
    spacedData,
    reviewList,
    priorityUpdates,
    priorityByAlvo,
    feitoMap,
    useAdvanced,
    horizonDays,
    priorityCol
  };
}

function getCandidateCostMinutes(item) {
  if (!item) return 1;
  const components = item.components || {};
  if (components.costMinutes && isFinite(components.costMinutes)) {
    return Math.max(0.25, components.costMinutes);
  }
  const tempoPrev = components.tempoPrev !== undefined
    ? components.tempoPrev
    : (item.context && item.context.tempoPrevSeg ? item.context.tempoPrevSeg : 60);
  return Math.max(0.25, (tempoPrev || 60) / 60);
}

function runBanditPlanner(reviewList, settings, customBudgetMinutes) {
  const candidates = reviewList || [];
  if (candidates.length === 0) {
    return { selected: [], metrics: {}, budget: 0, totalCost: 0 };
  }

  const metrics = {};
  let totalCost = 0;
  const enriched = candidates.map(item => {
    const cost = getCandidateCostMinutes(item);
    totalCost += cost;
    const components = item.components || {};
    const totalValue = components.eviTotalLCB !== undefined
      ? components.eviTotalLCB
      : (components.eviPerMin !== undefined ? components.eviPerMin * cost : item.prioridade || 0);
    const ratio = cost > 0 ? totalValue / cost : totalValue;
    metrics[item.alvo] = { cost, totalValue, ratio };
    return { item, cost, totalValue, ratio };
  });

  let budget = totalCost;
  if (isFinite(customBudgetMinutes) && customBudgetMinutes > 0) {
    budget = Math.min(totalCost, Math.max(0.25, customBudgetMinutes));
  } else {
    const mixRaw = parseFloat(settings.planGainMix);
    const mix = isFinite(mixRaw) ? clamp(mixRaw, 0, 1) : 0.5;
    if (mix > 0 && mix < 1) {
      budget = Math.max(0.25, totalCost * mix);
    } else if (mix <= 0) {
      budget = Math.max(0.25, totalCost * 0.5);
    }
  }

  const sorted = enriched.slice().sort((a, b) => (b.ratio || 0) - (a.ratio || 0));
  const selected = [];
  let remaining = budget;
  sorted.forEach(entry => {
    if (entry.cost <= remaining + 1e-6 || selected.length === 0) {
      selected.push(entry.item);
      remaining = Math.max(0, remaining - entry.cost);
    }
  });

  if (selected.length === 0) {
    return { selected: candidates.slice(), metrics, budget: totalCost, totalCost };
  }

  selected.sort((a, b) => (b.prioridade || 0) - (a.prioridade || 0));
  return { selected, metrics, budget, totalCost };
}

function apiMakeReviewToday() {
  try {
    const settings = apiGetSettings();
    const hoje = new Date();
    hoje.setHours(0, 0, 0, 0);

    const gather = gatherReviewCandidates(settings, hoje);
    const spacedSheet = gather.spacedSheet;
    const reviewSheet = gather.reviewSheet;

    if (!spacedSheet) {
      clearSheetData(SHEET_NAMES.REVER_HOJE);
      return { ok: true, count: 0, data: [] };
    }

    const reviewList = gather.reviewList;
    const priorityUpdates = gather.priorityUpdates;
    const priorityByAlvo = gather.priorityByAlvo;
    const spacedData = gather.spacedData;
    const priorityCol = gather.priorityCol;
    const useAdvanced = gather.useAdvanced;

    if (spacedData.length > 0 && priorityCol > 0 && priorityUpdates.length === spacedData.length) {
      const values = spacedData.map((row, idx) => {
        const alvo = row.alvo;
        const update = priorityUpdates[idx];
        const fallback = update ? update.value : 0;
        const val = alvo && priorityByAlvo[alvo] !== undefined ? priorityByAlvo[alvo] : fallback;
        return [val];
      });
      spacedSheet.getRange(2, priorityCol, values.length, 1).setValues(values);
    }

    let finalList = reviewList.slice();
    let policyEntries = [];
    const policyVersion = useAdvanced ? 'advanced_v1' : 'classic_v1';
    let banditResult = null;

    if (useAdvanced && asBoolean(settings.useBanditPlanner) && reviewList.length > 0) {
      banditResult = runBanditPlanner(reviewList, settings);
      const selectedSet = new Set((banditResult.selected || []).map(item => item.alvo));
      const metrics = banditResult.metrics || {};
      finalList = banditResult.selected;

      const allDecisions = reviewList.map(item => {
        const metric = metrics[item.alvo] || {};
        const components = item.components || {};
        components.banditRatio = metric.ratio;
        components.banditCost = metric.cost;
        components.banditValue = metric.totalValue;
        item.components = components;
        return {
          item,
          decisao: selectedSet.has(item.alvo) ? 'selected' : 'skipped',
          metric
        };
      });

      policyEntries = allDecisions.map(decision => {
        const item = decision.item;
        const areaParts = item.context && item.context.area ? {
          area: item.context.area,
          subarea: item.context.subarea
        } : parseAlvoParts(item.alvo);
        const components = item.components || {};
        return {
          timestamp: new Date(),
          alvo: item.alvo,
          area: areaParts.area,
          subarea: areaParts.subarea,
          pri: item.prioridade,
          eviPerMin: components.eviPerMin !== undefined ? components.eviPerMin : '',
          overdue: components.overdue !== undefined ? components.overdue : '',
          diversity: components.diversity !== undefined ? components.diversity : '',
          custos: components.custos !== undefined ? components.custos : '',
          tempoPrev: components.tempoPrev !== undefined ? components.tempoPrev : (item.context ? item.context.tempoPrevSeg : ''),
          decisao: decision.decisao,
          policyVersion: policyVersion
        };
      });
    } else if (finalList.length > 0) {
      policyEntries = finalList.map(item => {
        const areaParts = item.context && item.context.area ? {
          area: item.context.area,
          subarea: item.context.subarea
        } : parseAlvoParts(item.alvo);
        const components = item.components || {};
        return {
          timestamp: new Date(),
          alvo: item.alvo,
          area: areaParts.area,
          subarea: areaParts.subarea,
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
    }

    clearSheetData(SHEET_NAMES.REVER_HOJE);
    const rowsToWrite = finalList.map(item => [
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

    if (policyEntries.length > 0) {
      appendPolicyLogEntries(policyEntries);
    }

    const responseList = finalList.map(item => {
      const components = item.components || {};
      const diag = item.modelRow ? {
        sigma: item.modelRow.sigma,
        n_eff: item.modelRow.n_eff
      } : null;
      const context = item.context || null;
      const alvoParts = context && (context.area || context.subarea)
        ? { area: context.area, subarea: context.subarea }
        : parseAlvoParts(item.alvo || '');
      const tempoPrevSeg = context && context.tempoPrevSeg !== undefined
        ? context.tempoPrevSeg
        : (components.costMinutes !== undefined && components.costMinutes !== null
          ? components.costMinutes * 60
          : null);
      const tempoPrevMin = tempoPrevSeg !== null && tempoPrevSeg !== undefined
        ? tempoPrevSeg / 60
        : (components.costMinutes !== undefined ? components.costMinutes : null);
      return {
        alvo: item.alvo,
        prioridade: item.prioridade,
        proximaRevisao: formatDateDDMMYYYY(item.proximaRevisao),
        estabilidade: item.estabilidade,
        feito: item.feito || '',
        eviPerMin: components.eviPerMin !== undefined ? components.eviPerMin : null,
        eviPerMinMean: components.eviPerMinMean !== undefined ? components.eviPerMinMean : null,
        eviTotal: components.eviTotalLCB !== undefined ? components.eviTotalLCB : null,
        custos: components.custos !== undefined ? components.custos : null,
        overdue: components.overdue !== undefined ? components.overdue : null,
        diversity: components.diversity !== undefined ? components.diversity : null,
        costMinutes: components.costMinutes !== undefined ? components.costMinutes : null,
        diagnostics: diag,
        atrasoDias: context && context.atrasoDias !== undefined ? context.atrasoDias : null,
        tempoPrevMin: tempoPrevMin,
        baseRecall: context && context.baseRecall !== undefined ? context.baseRecall : null,
        area: alvoParts.area || '',
        subarea: alvoParts.subarea || ''
      };
    });

    return {
      ok: true,
      count: finalList.length,
      data: responseList,
      policyVersion,
      bandit: banditResult ? { budget: banditResult.budget, totalCost: banditResult.totalCost } : null
    };

  } catch (e) {
    return { ok: false, error: e.toString() };
  }
}

function apiEffortPlanner(budgetMinutes) {
  try {
    const settings = apiGetSettings();
    if (!asBoolean(settings.useAdvancedPriority) || !asBoolean(settings.useBanditPlanner)) {
      return { ok: false, disabled: true };
    }

    const hoje = new Date();
    hoje.setHours(0, 0, 0, 0);
    const gather = gatherReviewCandidates(settings, hoje);
    const reviewList = gather.reviewList || [];
    if (reviewList.length === 0) {
      return { ok: true, selected: [], budget: 0, totalCost: 0, count: 0 };
    }

    const result = runBanditPlanner(reviewList, settings, budgetMinutes);
    const metrics = result.metrics || {};
    const selectedSet = new Set((result.selected || []).map(item => item.alvo));

    const policyEntries = reviewList.map(item => {
      const areaParts = item.context && item.context.area ? {
        area: item.context.area,
        subarea: item.context.subarea
      } : parseAlvoParts(item.alvo);
      const components = item.components || {};
      const metric = metrics[item.alvo] || {};
      components.banditRatio = metric.ratio;
      components.banditCost = metric.cost;
      components.banditValue = metric.totalValue;
      return {
        timestamp: new Date(),
        alvo: item.alvo,
        area: areaParts.area,
        subarea: areaParts.subarea,
        pri: item.prioridade,
        eviPerMin: components.eviPerMin !== undefined ? components.eviPerMin : '',
        overdue: components.overdue !== undefined ? components.overdue : '',
        diversity: components.diversity !== undefined ? components.diversity : '',
        custos: components.custos !== undefined ? components.custos : '',
        tempoPrev: components.tempoPrev !== undefined ? components.tempoPrev : (item.context ? item.context.tempoPrevSeg : ''),
        decisao: selectedSet.has(item.alvo) ? 'planner_selected' : 'planner_skipped',
        policyVersion: 'advanced_v1'
      };
    });
    if (policyEntries.length > 0) {
      appendPolicyLogEntries(policyEntries);
    }

    const selected = (result.selected || []).map(item => {
      const metric = metrics[item.alvo] || {};
      return {
        alvo: item.alvo,
        prioridade: item.prioridade,
        custoMin: metric.cost,
        ganhoLCB: metric.totalValue,
        ratio: metric.ratio,
        estabilidade: item.estabilidade,
        proximaRevisao: formatDateDDMMYYYY(item.proximaRevisao)
      };
    });

    return {
      ok: true,
      selected,
      budget: result.budget,
      totalCost: result.totalCost,
      count: selected.length
    };

  } catch (e) {
    return { ok: false, error: e.toString() };
  }
}

function apiLogRetrospective(payload) {
  let lock = null;
  try {
    const params = payload || {};
    lock = LockService.getScriptLock();
    lock.tryLock(10000);

    const sheet = getOrCreateSheet(SHEET_NAMES.LOG, HEADERS.LOG);
    const hoje = new Date();
    hoje.setHours(0, 0, 0, 0);
    const data = formatDateDDMMYYYY(hoje);
    const totalMinutes = Number(params.totalMinutes) || 0;
    const totalGain = Number(params.totalGain) || 0;
    const executionPct = Number(params.executionPct) || 0;
    const phrase = params.phrase ? params.phrase.toString() : '';
    const topAreas = Array.isArray(params.topAreas) ? params.topAreas : [];

    const obsParts = [];
    if (phrase) obsParts.push(phrase);
    obsParts.push(`Execução ${executionPct.toFixed(1)}%`);
    obsParts.push(`Ganho +${totalGain.toFixed(1)}pp`);
    if (topAreas.length > 0) {
      obsParts.push(`Top: ${topAreas.join(', ')}`);
    }
    const obs = obsParts.join(' | ');

    const row = [
      data,
      'RETROSPECTIVA',
      '',
      Math.round(totalMinutes),
      0,
      Math.round(totalMinutes * 60),
      '',
      'retrospectiva',
      obs,
      Utilities.getUuid()
    ];

    sheet.appendRow(row);
    SpreadsheetApp.flush();
    if (lock) lock.releaseLock();
    return { ok: true };
  } catch (e) {
    if (lock) {
      try { lock.releaseLock(); } catch (err) {}
    }
    return { ok: false, error: e.toString() };
  }
}

function createPlannerItem(entry, statsMap, settings, index) {
  if (!entry || !entry.alvo) return null;
  const context = entry.context || {};
  const alvoParts = (context.area || context.subarea)
    ? { area: context.area, subarea: context.subarea }
    : parseAlvoParts(entry.alvo || '');
  const area = (alvoParts.area || '').toString().trim();
  const subarea = (alvoParts.subarea || '').toString().trim();
  const statsKey = `${area}::${subarea}`;
  const statsRow = statsKey && statsMap ? statsMap[statsKey] : null;

  let tempoMedioMin = null;
  if (statsRow && statsRow.tempo_medio !== undefined && statsRow.tempo_medio !== '') {
    const tempoStats = parseFloat(statsRow.tempo_medio);
    if (isFinite(tempoStats) && tempoStats > 0) {
      tempoMedioMin = tempoStats / 60;
    }
  }
  if (tempoMedioMin === null) {
    const components = entry.components || {};
    if (components.costMinutes !== undefined && components.costMinutes !== null) {
      const costMin = parseFloat(components.costMinutes);
      if (isFinite(costMin) && costMin > 0) {
        tempoMedioMin = costMin;
      }
    }
  }
  if (tempoMedioMin === null && context && context.tempoPrevSeg !== undefined) {
    const tempoSeg = parseFloat(context.tempoPrevSeg);
    if (isFinite(tempoSeg) && tempoSeg > 0) {
      tempoMedioMin = tempoSeg / 60;
    }
  }
  if (tempoMedioMin === null || !isFinite(tempoMedioMin) || tempoMedioMin <= 0) {
    tempoMedioMin = 1;
  }

  const tempoScore = Math.max(tempoMedioMin, 1);
  const components = entry.components || {};
  let eviTotalLCB = parseFloat(components.eviTotalLCB);
  let eviPerMinLCB = parseFloat(components.eviPerMin);
  if (!isFinite(eviTotalLCB) && isFinite(eviPerMinLCB)) {
    eviTotalLCB = eviPerMinLCB * tempoMedioMin;
  }
  if (!isFinite(eviPerMinLCB) && isFinite(eviTotalLCB) && tempoScore > 0) {
    eviPerMinLCB = eviTotalLCB / tempoScore;
  }
  let eviTotalMean = parseFloat(components.eviTotalMean);
  let eviPerMinMean = parseFloat(components.eviPerMinMean);
  if (!isFinite(eviPerMinMean) && isFinite(eviTotalMean) && tempoScore > 0) {
    eviPerMinMean = eviTotalMean / tempoScore;
  }
  if (!isFinite(eviTotalMean) && isFinite(eviPerMinMean)) {
    eviTotalMean = eviPerMinMean * tempoMedioMin;
  }
  const prioridade = Number(entry.prioridade) || 0;
  const priPerMinClassic = tempoScore > 0 ? Math.max(prioridade, 0.01) / tempoScore : Math.max(prioridade, 0.01);
  const atrasoDias = context && context.atrasoDias !== undefined
    ? parseFloat(context.atrasoDias)
    : null;
  const estabilidade = parseFloat(entry.estabilidade);
  const baseRecall = context && context.baseRecall !== undefined ? context.baseRecall : null;

  return {
    alvo: entry.alvo,
    area,
    subarea,
    prioridade,
    prioridadeBase: entry.prioridadeBase || prioridade,
    S: isFinite(estabilidade) ? estabilidade : null,
    t: atrasoDias !== null && isFinite(atrasoDias) ? atrasoDias : null,
    tempoMedio: tempoMedioMin,
    tempoScore,
    eviPerMinMean: isFinite(eviPerMinMean) ? eviPerMinMean : null,
    eviLCBPerMin: isFinite(eviPerMinLCB) ? eviPerMinLCB : null,
    eviTotalMean: isFinite(eviTotalMean) ? eviTotalMean : null,
    priPerMinClassic,
    scoreAdvanced: tempoScore > 0 ? Math.max(eviPerMinLCB || 0, 0) : 0,
    scoreClassic: Math.max(priPerMinClassic, 0),
    baseRecall: baseRecall !== null && baseRecall !== undefined ? baseRecall : null,
    feito: entry.feito || '',
    sheetIndex: entry.sheetIndex !== undefined ? entry.sheetIndex : index,
    components,
    dueDate: entry.proximaRevisao || null,
    context,
    score: 0,
    modelRow: entry.modelRow || null
  };
}

function calculateGainPerMinute(item, useAdvanced, fallbackToPriority, kappaPriToDelta) {
  const tempo = Math.max(item.tempoMedio || 1, 0.5);
  let gainPerMin = 0;
  if (useAdvanced && !fallbackToPriority) {
    if (item.eviPerMinMean !== null && item.eviPerMinMean !== undefined) {
      gainPerMin = Math.max(0, item.eviPerMinMean) * 100;
    } else if (item.eviLCBPerMin !== null && item.eviLCBPerMin !== undefined) {
      gainPerMin = Math.max(0, item.eviLCBPerMin) * 100;
    }
  }
  if (!isFinite(gainPerMin) || gainPerMin <= 0) {
    const priPerMin = Math.max(item.priPerMinClassic || 0, 0);
    gainPerMin = Math.max(0, kappaPriToDelta * priPerMin * 100);
  }
  if (!isFinite(gainPerMin) || gainPerMin < 0) {
    gainPerMin = 0;
  }
  return gainPerMin;
}

function applyStreakAndFatigueOrdering(targets, settings) {
  const fatigueRaw = settings && settings.fatigueFactor !== undefined
    ? parseFloat(settings.fatigueFactor)
    : DEFAULT_SETTINGS.fatigueFactor;
  const fatigue = clamp(isFinite(fatigueRaw) ? fatigueRaw : DEFAULT_SETTINGS.fatigueFactor, 0, 0.9);
  const allocated = targets.filter(item => (item.allocMin || 0) > 0);
  const remainder = targets.filter(item => (item.allocMin || 0) <= 0);
  const total = allocated.length;
  if (total <= 2) {
    return allocated.concat(remainder);
  }
  const pending = allocated.slice();
  const ordered = [];
  const lastAreas = [];
  const normalizer = Math.max(1, total - 1);
  while (pending.length > 0) {
    let bestIndex = -1;
    let bestScore = -Infinity;
    for (let i = 0; i < pending.length; i++) {
      const candidate = pending[i];
      const area = candidate.area || 'Sem área';
      const streak = lastAreas.length >= 2 && lastAreas[lastAreas.length - 1] === area && lastAreas[lastAreas.length - 2] === area;
      if (streak) {
        continue;
      }
      const position = ordered.length;
      const fatigueFactor = 1 - fatigue * (position / normalizer);
      const difficulty = clamp((candidate.tempoMedio || 1) / 6, 0, 1);
      const difficultyFactor = 1 - fatigue * difficulty * (position / normalizer);
      const baseScore = Math.max(candidate.score || 0, 0);
      const adjusted = baseScore * fatigueFactor * difficultyFactor;
      if (adjusted > bestScore + 1e-9) {
        bestScore = adjusted;
        bestIndex = i;
      }
    }
    if (bestIndex === -1) {
      bestIndex = 0;
    }
    const chosen = pending.splice(bestIndex, 1)[0];
    ordered.push(chosen);
    const area = chosen.area || 'Sem área';
    lastAreas.push(area);
    if (lastAreas.length > 2) {
      lastAreas.shift();
    }
  }
  return ordered.concat(remainder);
}

function applyPlanningModeAdjustments(settings, mode) {
  if (!settings) return settings;
  const normalized = (mode || '').toString().toLowerCase();
  const alpha = parseFloat(settings.alpha);
  const lambdaDiv = parseFloat(settings.lambdaDiversity);
  const beta = parseFloat(settings.betaUncertainty);
  if (normalized === 'power') {
    settings.alpha = alpha * (parseFloat(settings.powerAlphaScale) || DEFAULT_SETTINGS.powerAlphaScale || 1);
    settings.lambdaDiversity = lambdaDiv * (parseFloat(settings.powerDiversityScale) || DEFAULT_SETTINGS.powerDiversityScale || 1);
    settings.betaUncertainty = beta * (parseFloat(settings.powerBetaScale) || DEFAULT_SETTINGS.powerBetaScale || 1);
  } else if (normalized === 'maintenance') {
    settings.alpha = alpha * (parseFloat(settings.maintAlphaScale) || DEFAULT_SETTINGS.maintAlphaScale || 1);
    settings.lambdaDiversity = lambdaDiv * (parseFloat(settings.maintDiversityScale) || DEFAULT_SETTINGS.maintDiversityScale || 1);
    settings.betaUncertainty = beta * (parseFloat(settings.maintBetaScale) || DEFAULT_SETTINGS.maintBetaScale || 1);
  }
  return settings;
}

function normalizePlannerConfig(config) {
  const params = config || {};
  const budgetInput = parseFloat(params.budgetMin);
  const budgetMin = isFinite(budgetInput) && budgetInput > 0 ? budgetInput : 0;
  const roundInput = parseFloat(params.roundTo);
  const roundTo = isFinite(roundInput) && roundInput > 0 ? Math.max(1, Math.round(roundInput)) : 5;
  const minPriorityInput = params.minPriority !== undefined ? parseFloat(params.minPriority) : 0;
  const minPriority = isFinite(minPriorityInput) ? minPriorityInput : 0;
  const maxTargetsInput = parseInt(params.maxTargets, 10);
  const maxTargets = isFinite(maxTargetsInput) && maxTargetsInput > 0 ? maxTargetsInput : 12;
  const useReviewHoje = params.useReviewHoje === undefined ? true : !!params.useReviewHoje;

  return {
    budgetMin,
    roundTo,
    minPriority,
    maxTargets,
    useReviewHoje
  };
}

function buildProportionalPlan(items, options, settings) {
  const result = {
    targets: [],
    totalAlloc: 0,
    totalDelta: 0,
    fallback: false
  };
  const list = Array.isArray(items) ? items.slice() : [];
  const budget = Math.max(0, options && options.budgetMin !== undefined ? Number(options.budgetMin) : 0);
  const roundTo = Math.max(1, options && options.roundTo !== undefined ? Number(options.roundTo) : 5);
  const minUnit = Math.max(roundTo, 5);
  const useAdvanced = !!(options && options.useAdvanced);
  const kappa = options && options.kappaPriToDelta !== undefined ? Number(options.kappaPriToDelta) : DEFAULT_SETTINGS.kappaPriToDelta;

  if (list.length === 0) {
    return result;
  }

  const limited = list.slice();
  let fallback = false;

  let totalScore = 0;
  limited.forEach(item => {
    const baseScore = useAdvanced ? Math.max(item.scoreAdvanced || 0, 0) : Math.max(item.scoreClassic || 0, 0);
    item.score = baseScore;
    totalScore += item.score;
  });

  if (useAdvanced && totalScore <= 0 && limited.length > 0) {
    fallback = true;
    totalScore = 0;
    limited.forEach(item => {
      item.score = Math.max(item.scoreClassic || 0, 0);
      totalScore += item.score;
    });
  }

  if (totalScore <= 0 && limited.length > 0) {
    fallback = true;
    totalScore = 0;
    limited.forEach(item => {
      item.score = Math.max(item.scoreClassic || 0.01, 0.01);
      totalScore += item.score;
    });
  }

  if (budget <= 0 || totalScore <= 0) {
    result.fallback = fallback;
    result.targets = applyStreakAndFatigueOrdering(limited.map(item => ({
      alvo: item.alvo,
      area: item.area,
      subarea: item.subarea,
      prioridade: item.prioridade,
      S: item.S,
      t: item.t,
      tempoMedio: Math.round((item.tempoMedio || 0) * 100) / 100,
      eviPerMin: useAdvanced && !fallback ? item.eviPerMinMean : null,
      eviLCBPerMin: useAdvanced && !fallback ? item.eviLCBPerMin : null,
      priPerMin: (!useAdvanced || fallback) ? item.priPerMinClassic : null,
      allocMin: 0,
      deltaRpp: 0,
      feito: item.feito || '',
      baseRecall: item.baseRecall,
      score: item.score
    })), settings || {});
    return result;
  }

  const states = limited.map((item, index) => {
    const share = (budget * item.score) / totalScore;
    const down = roundTo > 0 ? Math.floor(share / roundTo) * roundTo : 0;
    const remainder = share - down;
    return {
      item,
      index,
      share,
      alloc: Math.max(0, down),
      remainder
    };
  });

  let allocated = states.reduce((sum, state) => sum + state.alloc, 0);
  let leftover = Math.max(0, budget - allocated);

  const zeroCandidates = states
    .filter(state => state.share > 0 && state.alloc === 0)
    .sort((a, b) => (b.share || 0) - (a.share || 0) || a.index - b.index);
  zeroCandidates.forEach(state => {
    if (leftover >= minUnit) {
      state.alloc += minUnit;
      leftover -= minUnit;
    }
  });

  if (roundTo > 0 && states.length > 0) {
    const remainderOrder = states.slice().sort((a, b) => {
      const diff = (b.remainder || 0) - (a.remainder || 0);
      if (Math.abs(diff) > 1e-9) return diff;
      return a.index - b.index;
    });
    while (leftover >= roundTo) {
      let progress = false;
      for (let i = 0; i < remainderOrder.length && leftover >= roundTo; i++) {
        const state = remainderOrder[i];
        if (state.share <= 0) continue;
        state.alloc += roundTo;
        leftover -= roundTo;
        progress = true;
      }
      if (!progress) {
        break;
      }
    }
  }

  const allocationMap = new Map();
  states.forEach(state => allocationMap.set(state.item.alvo, Math.max(0, state.alloc)));

  const prepared = limited.map(item => {
    const alloc = allocationMap.has(item.alvo) ? allocationMap.get(item.alvo) : 0;
    const allocRounded = Math.round(alloc * 100) / 100;
    const gainPerMin = calculateGainPerMinute(item, useAdvanced, fallback, kappa);
    const deltaRaw = allocRounded > 0 ? gainPerMin * allocRounded : 0;
    const delta = Math.max(0, Math.round(deltaRaw * 10) / 10);
    const tempoOut = Math.round((item.tempoMedio || 0) * 100) / 100;
    const diagSource = item.modelRow || item.diagnostics || null;
    const sigmaRaw = diagSource && diagSource.sigma !== undefined
      ? parseFloat(diagSource.sigma)
      : (item.sigma !== undefined ? parseFloat(item.sigma) : NaN);
    const nEffRaw = diagSource && (diagSource.n_eff !== undefined || diagSource.nEff !== undefined)
      ? parseFloat(diagSource.n_eff !== undefined ? diagSource.n_eff : diagSource.nEff)
      : (item.nEff !== undefined ? parseFloat(item.nEff) : NaN);
    const sigmaVal = isFinite(sigmaRaw) && sigmaRaw > 0 ? sigmaRaw : null;
    const nEffVal = isFinite(nEffRaw) && nEffRaw >= 0 ? nEffRaw : null;
    return {
      alvo: item.alvo,
      area: item.area,
      subarea: item.subarea,
      prioridade: item.prioridade,
      S: item.S,
      t: item.t,
      tempoMedio: tempoOut,
      eviPerMin: useAdvanced && !fallback ? item.eviPerMinMean : null,
      eviLCBPerMin: useAdvanced && !fallback ? item.eviLCBPerMin : null,
      priPerMin: (!useAdvanced || fallback) ? item.priPerMinClassic : null,
      allocMin: allocRounded,
      deltaRpp: delta,
      feito: item.feito || '',
      baseRecall: item.baseRecall,
      score: item.score,
      sigma: sigmaVal,
      nEff: nEffVal
    };
  });

  const orderedTargets = applyStreakAndFatigueOrdering(prepared, settings || {});
  let totalAlloc = 0;
  let totalDelta = 0;
  orderedTargets.forEach(target => {
    totalAlloc += target.allocMin || 0;
    totalDelta += target.deltaRpp || 0;
  });

  result.targets = orderedTargets;
  result.totalAlloc = Math.round(totalAlloc);
  result.totalDelta = Math.round(totalDelta * 10) / 10;
  result.fallback = fallback;
  return result;
}

function apiPlanDayBudget(params) {
  try {
    const settings = apiGetSettings();
    const config = params || {};

    if (!asBoolean(settings.useBanditPlanner)) {
      return { ok: false, disabled: true };
    }

    const planningSettings = applyPlanningModeAdjustments(Object.assign({}, settings), config.mode);
    const normalizedConfig = normalizePlannerConfig(config);
    const useAdvanced = asBoolean(planningSettings.useAdvancedPriority);
    const budgetMin = normalizedConfig.budgetMin;
    const roundTo = normalizedConfig.roundTo;
    const minPriority = normalizedConfig.minPriority;
    const maxTargets = normalizedConfig.maxTargets;
    const useReviewHoje = normalizedConfig.useReviewHoje;
    const kappaInput = planningSettings.kappaPriToDelta !== undefined && planningSettings.kappaPriToDelta !== ''
      ? parseFloat(planningSettings.kappaPriToDelta)
      : DEFAULT_SETTINGS.kappaPriToDelta;
    const kappaPriToDelta = isFinite(kappaInput) ? kappaInput : DEFAULT_SETTINGS.kappaPriToDelta;

    const hoje = new Date();
    hoje.setHours(0, 0, 0, 0);

    const dataCache = buildStudyGuideData(planningSettings);
    const logRows = readSheetData(SHEET_NAMES.LOG);
    const revisaoRows = readSheetData(SHEET_NAMES.REVISAO_LOG);
    const examConfig = readSheetData(SHEET_NAMES.EXAM_CONFIG);

    const gather = gatherReviewCandidates(planningSettings, hoje);
    const reviewList = gather.reviewList || [];
    const reviewMap = {};
    reviewList.forEach((item, idx) => {
      if (item && item.alvo) {
        reviewMap[item.alvo] = { item, idx };
      }
    });

    const statsData = readSheetData(SHEET_NAMES.STATS);
    const statsMap = {};
    statsData.forEach(row => {
      if (!row) return;
      const key = `${row.area}::${row.subarea}`;
      statsMap[key] = row;
    });

    let ordered = reviewList.slice();
    if (useReviewHoje) {
      const reviewHoje = readSheetData(SHEET_NAMES.REVER_HOJE);
      if (reviewHoje.length > 0) {
        const seen = new Set();
        ordered = [];
        reviewHoje.forEach((row, idx) => {
          if (!row || !row.alvo) return;
          const key = row.alvo;
          seen.add(key);
          if (reviewMap[key]) {
            const existing = reviewMap[key].item;
            existing.sheetIndex = idx;
            ordered.push(existing);
          } else {
            const parsed = parseAlvoParts(key);
            ordered.push({
              alvo: key,
              prioridade: Number(row.prioridade) || 0,
              prioridadeBase: Number(row.prioridade) || 0,
              proximaRevisao: row.proximaRevisao ? parseSheetDate(row.proximaRevisao) : null,
              estabilidade: Number(row.estabilidade) || settings.Smin,
              components: {},
              context: {
                alvo: key,
                area: parsed.area,
                subarea: parsed.subarea,
                tempoPrevSeg: 60,
                atrasoDias: null,
                baseRecall: 0
              },
              feito: row.feito || '',
              areaKey: parsed.area || 'Sem área',
              sheetIndex: idx
            });
          }
        });
        reviewList.forEach((candidate, idx) => {
          if (!candidate || !candidate.alvo) return;
          if (!seen.has(candidate.alvo)) {
            candidate.sheetIndex = reviewHoje.length + idx;
            ordered.push(candidate);
          }
        });
      }
    }

    if (ordered.length === 0) {
      return {
        ok: true,
        budgetMin: budgetMin,
        allocatedMin: 0,
        totalDeltaRpp: 0,
        targets: [],
        areas: [],
        mode: useAdvanced ? 'advanced' : 'classic'
      };
    }

    const items = [];
    ordered.forEach((entry, idx) => {
      if (!entry || !entry.alvo) return;
      const prioridade = Number(entry.prioridade) || 0;
      if (prioridade < minPriority) {
        return;
      }
      const item = createPlannerItem(entry, statsMap, planningSettings, idx);
      if (item) {
        items.push(item);
      }
    });

    if (items.length === 0) {
      return {
        ok: true,
        budgetMin: budgetMin,
        allocatedMin: 0,
        totalDeltaRpp: 0,
        targets: [],
        areas: [],
        mode: useAdvanced ? 'advanced' : 'classic',
        fallbackToPriority: false
      };
    }

    const limited = maxTargets > 0 && items.length > maxTargets ? items.slice(0, maxTargets) : items.slice();
    const plan = buildProportionalPlan(limited, {
      budgetMin,
      roundTo,
      useAdvanced,
      kappaPriToDelta
    }, planningSettings);

    const areaAgg = {};
    plan.targets.forEach(target => {
      const alloc = target.allocMin || 0;
      if (alloc <= 0) return;
      const delta = target.deltaRpp || 0;
      const key = target.area || 'Sem área';
      if (!areaAgg[key]) {
        areaAgg[key] = { area: key, allocMin: 0, deltaRpp: 0 };
      }
      areaAgg[key].allocMin += alloc;
      areaAgg[key].deltaRpp += delta;
    });

    const areas = Object.keys(areaAgg).map(key => ({
      area: key,
      allocMin: Math.round(areaAgg[key].allocMin),
      deltaRpp: Math.round(areaAgg[key].deltaRpp * 10) / 10
    })).sort((a, b) => (b.allocMin || 0) - (a.allocMin || 0));

    const prefsForGuide = {
      flashcardsPerMinBase: planningSettings.flashcardsPerMinBase !== undefined
        ? planningSettings.flashcardsPerMinBase
        : DEFAULT_SETTINGS.flashcardsPerMinBase,
      blockQuestionsTarget: 25,
      readShare: 0.35
    };
    const todayDate = new Date(hoje.getTime());

    let newCount = 0;
    let historicalCount = 0;
    let guideMinutesSum = 0;
    let guideDeltaSum = 0;

    const enrichedTargets = plan.targets.map(target => {
      const enriched = Object.assign({}, target);
      const allocMinutes = Math.max(Math.round(target.allocMin || 0), 15);
      const guidePlan = computeStudyGuidePlanV2(target.alvo, allocMinutes, prefsForGuide, planningSettings, dataCache, logRows, revisaoRows, examConfig, todayDate);
      if (guidePlan) {
        if (guidePlan.status) {
          enriched.guideStatus = guidePlan.status;
          if (guidePlan.status === 'NOVO') newCount += 1;
          else if (guidePlan.status === 'HISTORICO') historicalCount += 1;
        }
        if (guidePlan.stage) {
          enriched.guideStage = guidePlan.stage;
        }
        if (guidePlan.planToday) {
          const todayPlan = guidePlan.planToday;
          if (todayPlan.totalMin !== undefined && todayPlan.totalMin !== null) {
            enriched.guideTotalMin = Number(todayPlan.totalMin);
            if (isFinite(enriched.guideTotalMin)) {
              guideMinutesSum += enriched.guideTotalMin;
            }
          }
          if (todayPlan.deltaRpp !== undefined && todayPlan.deltaRpp !== null) {
            enriched.guideDelta = Number(todayPlan.deltaRpp);
            if (isFinite(enriched.guideDelta)) {
              guideDeltaSum += enriched.guideDelta;
            }
          }
        }
      }
      return enriched;
    });

    const guideSummary = {
      countNew: newCount,
      countHistorical: historicalCount,
      totalMinutes: Math.round(guideMinutesSum),
      totalDeltaRpp: Math.round(guideDeltaSum * 10) / 10
    };

    return {
      ok: true,
      mode: useAdvanced ? (plan.fallback ? 'classic' : 'advanced') : 'classic',
      fallbackToPriority: plan.fallback,
      budgetMin: budgetMin,
      allocatedMin: Math.round(plan.totalAlloc),
      totalDeltaRpp: plan.totalDelta,
      targets: enrichedTargets,
      areas,
      guideSummary
    };
  } catch (e) {
    return { ok: false, error: e.toString() };
  }
}

function computeModeDiagnostics(targets) {
  const response = {
    avgSigma: null,
    avgNEff: null,
    lowDataShare: 0,
    lowDataCount: 0,
    count: 0
  };
  if (!Array.isArray(targets) || targets.length === 0) {
    return response;
  }

  let sigmaSum = 0;
  let sigmaCount = 0;
  let nEffSum = 0;
  let nEffCount = 0;
  let lowDataCount = 0;
  const threshold = 20;

  targets.forEach(target => {
    if (!target) return;
    const sigmaVal = target.sigma !== undefined && target.sigma !== null
      ? parseFloat(target.sigma)
      : (target.modelRow && target.modelRow.sigma !== undefined ? parseFloat(target.modelRow.sigma) : NaN);
    if (isFinite(sigmaVal) && sigmaVal > 0) {
      sigmaSum += sigmaVal;
      sigmaCount++;
    }

    const nEffVal = target.nEff !== undefined && target.nEff !== null
      ? parseFloat(target.nEff)
      : (target.modelRow && (target.modelRow.n_eff !== undefined || target.modelRow.nEff !== undefined)
        ? parseFloat(target.modelRow.n_eff !== undefined ? target.modelRow.n_eff : target.modelRow.nEff)
        : NaN);
    if (isFinite(nEffVal) && nEffVal >= 0) {
      nEffSum += nEffVal;
      nEffCount++;
      if (nEffVal < threshold) {
        lowDataCount++;
      }
    }
  });

  response.count = targets.length;
  response.lowDataCount = lowDataCount;
  response.lowDataShare = targets.length > 0 ? lowDataCount / targets.length : 0;
  response.avgSigma = sigmaCount > 0 ? sigmaSum / sigmaCount : null;
  response.avgNEff = nEffCount > 0 ? nEffSum / nEffCount : null;
  return response;
}

function apiCompareModes(params) {
  try {
    const settings = apiGetSettings();
    if (!asBoolean(settings.useBanditPlanner)) {
      return { ok: false, disabled: true };
    }

    const baseConfig = normalizePlannerConfig(params || {});
    const powerParams = Object.assign({}, baseConfig, { mode: 'power' });
    const maintenanceParams = Object.assign({}, baseConfig, { mode: 'maintenance' });

    const powerPlan = apiPlanDayBudget(powerParams);
    if (!powerPlan || powerPlan.disabled) {
      return { ok: false, disabled: true };
    }
    if (!powerPlan.ok) {
      return { ok: false, error: powerPlan.error || 'Falha ao simular modo Power' };
    }

    const maintenancePlan = apiPlanDayBudget(maintenanceParams);
    if (!maintenancePlan || maintenancePlan.disabled) {
      return { ok: false, disabled: true };
    }
    if (!maintenancePlan.ok) {
      return { ok: false, error: maintenancePlan.error || 'Falha ao simular modo Maintenance' };
    }

    const powerDiag = computeModeDiagnostics(powerPlan.targets || []);
    const maintenanceDiag = computeModeDiagnostics(maintenancePlan.targets || []);

    const sanitizeModeOutput = function(plan, diag) {
      const guide = plan.guideSummary || {};
      return {
        totalDeltaRpp: Number(plan.totalDeltaRpp) || 0,
        areas: Array.isArray(plan.areas) ? plan.areas : [],
        avgSigma: diag.avgSigma,
        avgNEff: diag.avgNEff,
        lowDataShare: diag.lowDataShare,
        lowDataCount: diag.lowDataCount,
        targetCount: diag.count,
        allocatedMin: Number(plan.allocatedMin) || 0,
        fallbackToPriority: !!plan.fallbackToPriority,
        modeUsed: plan.mode || '',
        guideSummary: {
          countNew: Number(guide.countNew) || 0,
          countHistorical: Number(guide.countHistorical) || 0,
          totalMinutes: Number(guide.totalMinutes) || 0,
          totalDeltaRpp: Number(guide.totalDeltaRpp) || 0
        }
      };
    };

    return {
      ok: true,
      budgetMin: baseConfig.budgetMin,
      power: sanitizeModeOutput(powerPlan, powerDiag),
      maintenance: sanitizeModeOutput(maintenancePlan, maintenanceDiag)
    };
  } catch (e) {
    return { ok: false, error: e.toString() };
  }
}

function buildStudyGuideData(settings) {
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const spacedRows = readSheetData(SHEET_NAMES.SPACED);
  const spacedMap = {};
  spacedRows.forEach(row => {
    if (!row || !row.alvo) return;
    const key = row.alvo.toString().trim();
    if (key) {
      spacedMap[key] = row;
    }
  });

  const statsRows = readSheetData(SHEET_NAMES.STATS);
  const statsMap = {};
  statsRows.forEach(row => {
    if (!row) return;
    const area = (row.area || '').toString().trim();
    const subarea = (row.subarea || '').toString().trim();
    const key = `${area}::${subarea}`;
    statsMap[key] = row;
  });

  const modelRows = readSheetData(SHEET_NAMES.MODEL);
  const modelMap = {};
  modelRows.forEach(row => {
    if (!row || !row.alvo) return;
    const key = row.alvo.toString().trim();
    if (key) {
      modelMap[key] = row;
    }
  });

  const gather = gatherReviewCandidates(settings, today);
  const reviewList = Array.isArray(gather.reviewList) ? gather.reviewList : [];
  const reviewMap = {};
  reviewList.forEach((entry, idx) => {
    if (!entry || !entry.alvo) return;
    const key = entry.alvo.toString().trim();
    if (key && !reviewMap[key]) {
      reviewMap[key] = { entry, index: idx };
    }
  });

  return {
    today,
    spacedMap,
    statsMap,
    modelMap,
    reviewMap,
    reviewList
  };
}

function computeStudyGuidePlanForTarget(alvoRaw, budgetMin, prefs, settings, data) {
  if (!alvoRaw || !settings || !data) return null;
  const alvoKey = alvoRaw.toString().trim();
  if (!alvoKey) return null;

  const parts = parseAlvoParts(alvoKey);
  let area = parts.area;
  let subarea = parts.subarea;

  const spacedRow = data.spacedMap && data.spacedMap[alvoKey] ? data.spacedMap[alvoKey] : null;
  if ((!area || !subarea) && spacedRow && spacedRow.alvo) {
    const spacedParts = parseAlvoParts(spacedRow.alvo);
    area = area || spacedParts.area;
    subarea = subarea || spacedParts.subarea;
  }

  const reviewWrapper = data.reviewMap && data.reviewMap[alvoKey] ? data.reviewMap[alvoKey] : null;
  const reviewEntry = reviewWrapper ? reviewWrapper.entry : null;
  if ((!area || !subarea) && reviewEntry && reviewEntry.context) {
    area = area || (reviewEntry.context.area || '').toString().trim();
    subarea = subarea || (reviewEntry.context.subarea || '').toString().trim();
  }

  const fallbackParts = parseAlvoParts(alvoKey);
  area = (area || fallbackParts.area || '').trim();
  subarea = (subarea || fallbackParts.subarea || '').trim();

  const statsKeyBase = `${area}::${subarea}`;
  let statsRow = data.statsMap ? data.statsMap[statsKeyBase] : null;
  if (!statsRow && data.statsMap) {
    const altKey = Object.keys(data.statsMap).find(key => key && key.toLowerCase() === statsKeyBase.toLowerCase());
    if (altKey) {
      statsRow = data.statsMap[altKey];
    }
  }

  const modelRow = data.modelMap && data.modelMap[alvoKey] ? data.modelMap[alvoKey] : null;

  const SminSetting = parseFloat(settings.Smin);
  const Smin = isFinite(SminSetting) ? SminSetting : DEFAULT_SETTINGS.Smin;
  const metaSetting = parseFloat(settings.retentionTarget);
  const retentionTarget = isFinite(metaSetting) && metaSetting > 0 ? metaSetting : DEFAULT_SETTINGS.retentionTarget;

  const prefsFlash = parseFloat(prefs.flashcardsPerMinBase);
  const prefsBlock = parseFloat(prefs.blockQuestionsTarget);
  let prefsReadShare = parseFloat(prefs.readShare);
  if (isFinite(prefsReadShare) && prefsReadShare > 1.5) {
    prefsReadShare = prefsReadShare / 100;
  }
  const flashRate = clamp(isFinite(prefsFlash) ? prefsFlash : 2, 0.5, 5);
  const blockTarget = isFinite(prefsBlock) ? clamp(prefsBlock, 5, 120) : 25;
  const readShare = clamp(isFinite(prefsReadShare) ? prefsReadShare : 0.35, 0.1, 0.6);

  const msPerDay = 1000 * 60 * 60 * 24;
  let S = null;
  if (spacedRow && spacedRow.estabilidade !== undefined && spacedRow.estabilidade !== '') {
    const value = parseFloat(spacedRow.estabilidade);
    if (isFinite(value)) {
      S = value;
    }
  }
  if (!isFinite(S) && modelRow && modelRow.S_atual !== undefined && modelRow.S_atual !== '') {
    const value = parseFloat(modelRow.S_atual);
    if (isFinite(value)) {
      S = value;
    }
  }

  let context = null;
  if (spacedRow) {
    context = buildPriorityContext(spacedRow, statsRow || null, settings, data.today);
  } else if (reviewEntry && reviewEntry.context) {
    context = reviewEntry.context;
  }
  if (!isFinite(S) && context && context.S !== undefined) {
    const value = parseFloat(context.S);
    if (isFinite(value)) {
      S = value;
    }
  }
  if (!isFinite(S) || S <= 0) {
    S = Smin;
  }
  S = applyCapS(S, Smin, parseFloat(settings.Smax) || DEFAULT_SETTINGS.Smax);

  const useWeibull = asBoolean(settings.useWeibull);
  let weibullK = 1;
  if (useWeibull) {
    if (modelRow && modelRow.weibull_k !== undefined && modelRow.weibull_k !== '') {
      const kCandidate = parseFloat(modelRow.weibull_k);
      if (isFinite(kCandidate) && kCandidate > 0) {
        weibullK = kCandidate;
      }
    }
    if (!isFinite(weibullK) || weibullK <= 0) {
      weibullK = getWeibullShape(area, settings);
    }
    if (!isFinite(weibullK) || weibullK <= 0) {
      weibullK = 1;
    }
  }

  let lastReviewDays = 0;
  if (spacedRow && spacedRow.ultimaRevisao) {
    const ultima = parseSheetDate(spacedRow.ultimaRevisao);
    if (ultima) {
      lastReviewDays = Math.max(0, Math.floor((data.today - ultima) / msPerDay));
    }
  } else if (context && context.atrasoDias !== undefined) {
    const atraso = parseFloat(context.atrasoDias);
    if (isFinite(atraso)) {
      lastReviewDays = Math.max(0, atraso);
    }
  }

  const Rhoje = calcRecall(lastReviewDays, Math.max(1, S), weibullK);

  let tempoMedioMin = null;
  if (statsRow && statsRow.tempo_medio !== undefined && statsRow.tempo_medio !== '') {
    const tempoStats = parseFloat(statsRow.tempo_medio);
    if (isFinite(tempoStats) && tempoStats > 0) {
      tempoMedioMin = tempoStats / 60;
    }
  }
  if ((tempoMedioMin === null || !isFinite(tempoMedioMin) || tempoMedioMin <= 0) && context && context.tempoPrevSeg !== undefined) {
    const tempoSeg = parseFloat(context.tempoPrevSeg);
    if (isFinite(tempoSeg) && tempoSeg > 0) {
      tempoMedioMin = tempoSeg / 60;
    }
  }
  if (tempoMedioMin === null || !isFinite(tempoMedioMin) || tempoMedioMin <= 0) {
    tempoMedioMin = 1;
  }

  let acerto28 = statsRow && statsRow.acerto_28d !== undefined ? parseFloat(statsRow.acerto_28d) : NaN;
  if (!isFinite(acerto28)) {
    acerto28 = statsRow && statsRow.acerto_vida !== undefined ? parseFloat(statsRow.acerto_vida) : NaN;
  }
  if (!isFinite(acerto28)) {
    acerto28 = 0.7;
  }
  acerto28 = clamp(acerto28, 0, 1);

  let difMedia = statsRow && statsRow.dif_media !== undefined ? parseFloat(statsRow.dif_media) : NaN;
  if (!isFinite(difMedia)) {
    difMedia = 3;
  }

  const kappaValue = settings.kappaPriToDelta !== undefined ? parseFloat(settings.kappaPriToDelta) : NaN;
  const kappaPriToDelta = isFinite(kappaValue) ? kappaValue : DEFAULT_SETTINGS.kappaPriToDelta;
  const useAdvanced = asBoolean(settings.useAdvancedPriority);

  let priorityClassic = null;
  if (reviewEntry && reviewEntry.prioridade !== undefined && reviewEntry.prioridade !== '') {
    const pri = parseFloat(reviewEntry.prioridade);
    if (isFinite(pri)) {
      priorityClassic = pri;
    }
  }
  if (!isFinite(priorityClassic) && context) {
    const classic = calculateClassicPriority(context, settings);
    if (classic && classic.score !== undefined) {
      priorityClassic = classic.score;
    }
  }
  if (!isFinite(priorityClassic)) {
    priorityClassic = 0.1;
  }

  let plannerItem = null;
  if (reviewEntry) {
    plannerItem = createPlannerItem(reviewEntry, data.statsMap, settings, reviewWrapper ? reviewWrapper.index : 0);
  }

  const fallbackToPriority = !useAdvanced || !plannerItem ||
    ((!plannerItem.eviPerMinMean && plannerItem.eviPerMinMean !== 0) && (!plannerItem.eviLCBPerMin && plannerItem.eviLCBPerMin !== 0));

  const gainPerMin = calculateGainPerMinute({
    tempoMedio: tempoMedioMin,
    eviPerMinMean: plannerItem ? plannerItem.eviPerMinMean : null,
    eviLCBPerMin: plannerItem ? plannerItem.eviLCBPerMin : null,
    priPerMinClassic: Math.max(priorityClassic, 0.01) / Math.max(tempoMedioMin, 0.25)
  }, useAdvanced, fallbackToPriority, kappaPriToDelta);

  const deltaPerMin = isFinite(gainPerMin) ? Math.max(gainPerMin, 0) : 0;
  const budget = isFinite(budgetMin) && budgetMin > 0 ? budgetMin : 0;

  let readMin = Math.round(budget * readShare);
  if (budget >= 30) {
    readMin = Math.max(15, readMin);
  } else {
    readMin = Math.max(0, readMin);
  }
  if (readMin > budget) {
    readMin = Math.max(0, Math.floor(budget * 0.6));
  }

  let remaining = Math.max(0, budget - readMin);
  let blockMin = Math.round(Math.max(remaining * 0.45, budget >= 40 ? 20 : Math.min(remaining, 20)));
  if (blockMin > remaining) {
    blockMin = Math.round(Math.max(remaining, 0));
  }
  remaining = Math.max(0, budget - readMin - blockMin);
  let flashMin = Math.round(remaining);

  const totalUsed = readMin + blockMin + flashMin;
  const budgetRounded = Math.round(budget);
  if (totalUsed !== budgetRounded) {
    const diff = budgetRounded - totalUsed;
    if (diff !== 0) {
      flashMin = Math.max(0, flashMin + diff);
      if (flashMin < 0) {
        blockMin = Math.max(0, blockMin + flashMin);
        flashMin = 0;
        if (blockMin < 0) {
          readMin = Math.max(0, readMin + blockMin);
          blockMin = 0;
        }
      }
    }
  }

  const flashcardsMin = Math.max(0, flashMin);
  const difFactor = clamp(1 + 0.1 * (difMedia - 3), 0.6, 1.4);
  let acertoFactor = 1;
  if (acerto28 < 0.6) {
    acertoFactor = 1.1;
  } else if (acerto28 > 0.8) {
    acertoFactor = 0.9;
  }
  const rawCards = flashcardsMin > 0 ? flashRate * flashcardsMin * difFactor * acertoFactor : 0;
  const flashcards = flashcardsMin > 0 ? Math.round(clamp(rawCards, 20, 120)) : 0;

  const tempoQuest = Math.max(tempoMedioMin, 0.5);
  let blockQuestions = blockMin > 0 ? Math.max(5, Math.round(blockTarget)) : 0;
  if (blockMin > 0 && tempoQuest > 0) {
    const maxByTime = Math.floor(blockMin / tempoQuest);
    if (isFinite(maxByTime) && maxByTime > 0) {
      if (blockQuestions > maxByTime) {
        blockQuestions = Math.max(5, maxByTime);
      } else {
        const fill = Math.max(5, Math.round(blockMin / tempoQuest));
        blockQuestions = Math.max(5, Math.min(Math.max(blockQuestions, fill), Math.max(maxByTime, blockQuestions)));
      }
    }
  }
  if (blockMin <= 0) {
    blockQuestions = 0;
  }

  const blockEst = blockQuestions > 0 ? Math.round(blockQuestions * tempoQuest) : Math.round(blockMin);
  const deltaHoje = Math.max(0, Math.round(deltaPerMin * budget * 10) / 10);

  const firstInterval = calcOptimalInterval(S, retentionTarget, weibullK);
  let firstOffset = null;
  if (spacedRow && spacedRow.proximaRevisao) {
    const prox = parseSheetDate(spacedRow.proximaRevisao);
    if (prox) {
      firstOffset = Math.max(1, Math.round((prox - data.today) / msPerDay));
    }
  }
  if (!isFinite(firstOffset) || firstOffset === null) {
    firstOffset = Math.max(2, Math.round(firstInterval));
  }
  if (!isFinite(firstOffset) || firstOffset <= 0) {
    firstOffset = 2;
  }

  const offsets = [
    firstOffset,
    Math.max(firstOffset + 3, Math.round(firstOffset * 2)),
    Math.max(firstOffset + 7, Math.round(firstOffset * 1.5))
  ];
  const uniqueOffsets = [];
  offsets.forEach(off => {
    const val = Math.max(1, Math.round(off));
    if (uniqueOffsets.indexOf(val) === -1) {
      uniqueOffsets.push(val);
    }
  });

  const nextReviews = [];
  uniqueOffsets.slice(0, 3).forEach((offset, idx) => {
    const estMin = offset <= 3 ? 15 : 20;
    const tipo = offset <= 3 ? 'revisão curta' : 'revisão média';
    const decay = Math.pow(0.85, idx);
    const delta = Math.max(0, Math.round(deltaPerMin * estMin * decay * 10) / 10);
    nextReviews.push({ diaOffset: offset, tipo, estMin, deltaRpp: delta });
  });

  const total7 = nextReviews
    .filter(item => item && item.diaOffset <= 7)
    .reduce((sum, item) => sum + (item.deltaRpp || 0), 0);
  const total28 = nextReviews
    .filter(item => item && item.diaOffset <= 28)
    .reduce((sum, item) => sum + (item.deltaRpp || 0), 0);

  return {
    alvo: alvoKey,
    area,
    subarea,
    diagnostics: {
      S,
      Rhoje,
      tempoMedio: tempoMedioMin,
      acerto_28d: acerto28,
      difMedia
    },
    D1: {
      readMin: Math.round(readMin),
      flashcards,
      blockQuestions,
      blockEstMin: Math.max(0, blockEst),
      deltaRpp: deltaHoje,
      rationale: {
        cardsRate: Number(flashRate.toFixed ? flashRate.toFixed(2) : flashRate),
        difficultyFactor: Number(difFactor.toFixed ? difFactor.toFixed(2) : difFactor),
        efficiencyAdj: Number(acertoFactor.toFixed ? acertoFactor.toFixed(2) : acertoFactor)
      }
    },
    D2plus: {
      nextReviews,
      totalDeltaRpp7d: Math.round(total7 * 10) / 10,
      totalDeltaRpp28d: Math.round(total28 * 10) / 10
    }
  };
}

function apiStudyGuidePlanV2(params) {
  try {
    const settings = apiGetSettings();
    const payload = params || {};
    const alvoRaw = (payload.alvo || '').toString().trim();
    if (!alvoRaw) {
      return { ok: false, error: 'Alvo inválido' };
    }

    const budgetInput = parseFloat(payload.budgetMin !== undefined ? payload.budgetMin : payload.budgetD1Min);
    const budgetMin = isFinite(budgetInput) && budgetInput > 0 ? budgetInput : 0;

    const prefsInput = payload.prefs || payload.preferencias || {};
    const flashRateSetting = settings.flashcardsPerMinBase !== undefined ? parseFloat(settings.flashcardsPerMinBase) : DEFAULT_SETTINGS.flashcardsPerMinBase;
    const readShareDefault = prefsInput.readShare !== undefined ? prefsInput.readShare : 0.35;
    const prefs = {
      flashcardsPerMinBase: prefsInput.flashcardsPerMinBase !== undefined ? prefsInput.flashcardsPerMinBase : flashRateSetting,
      blockQuestionsTarget: prefsInput.blockQuestionsTarget !== undefined ? prefsInput.blockQuestionsTarget : 25,
      readShare: readShareDefault
    };

    const today = new Date();
    today.setHours(0, 0, 0, 0);

    const planningSettings = Object.assign({}, settings);
    const data = buildStudyGuideData(planningSettings);
    const logRows = readSheetData(SHEET_NAMES.LOG);
    const revisaoRows = readSheetData(SHEET_NAMES.REVISAO_LOG);
    const examConfig = readSheetData(SHEET_NAMES.EXAM_CONFIG);

    const plan = computeStudyGuidePlanV2(alvoRaw, budgetMin, prefs, planningSettings, data, logRows, revisaoRows, examConfig, today);
    if (!plan) {
      return { ok: false, error: 'Alvo não encontrado ou sem dados suficientes' };
    }
    return Object.assign({ ok: true }, plan);
  } catch (e) {
    return { ok: false, error: e.toString() };
  }
}

function computeStudyGuidePlanV2(alvoRaw, budgetMin, prefs, settings, data, logRows, revisaoRows, examConfig, today) {
  if (!alvoRaw || !settings || !data) return null;
  const alvoKey = alvoRaw.toString().trim();
  if (!alvoKey) return null;

  const parts = parseAlvoParts(alvoKey);
  let area = parts.area;
  let subarea = parts.subarea;

  const statsRow = resolveStatsRow(data.statsMap, area, subarea);
  if (statsRow) {
    area = area || (statsRow.area || '').toString().trim();
    subarea = subarea || (statsRow.subarea || '').toString().trim();
  }

  const spacedRow = resolveSpacedRow(data.spacedMap, alvoKey);
  const modelRow = resolveModelRow(data.modelMap, alvoKey);
  const reviewWrapper = data.reviewMap && data.reviewMap[alvoKey] ? data.reviewMap[alvoKey] : null;
  const reviewEntry = reviewWrapper ? reviewWrapper.entry : null;

  const budgetMinutes = Math.max(0, Math.round(budgetMin || 0));

  const flashRateRaw = parseFloat(prefs.flashcardsPerMinBase);
  const flashRate = isFinite(flashRateRaw) ? clamp(flashRateRaw, 0.5, 5) : DEFAULT_SETTINGS.flashcardsPerMinBase;
  const blockTargetRaw = parseFloat(prefs.blockQuestionsTarget);
  const blockTarget = isFinite(blockTargetRaw) ? clamp(blockTargetRaw, 5, 120) : 25;
  let readShareRaw = prefs.readShare;
  if (isFinite(readShareRaw) && readShareRaw > 1) {
    readShareRaw = readShareRaw / 100;
  }
  const readShare = isFinite(readShareRaw) ? clamp(readShareRaw, 0.1, 0.6) : 0.35;

  const Smin = parseFloat(settings.Smin);
  const retentionTarget = parseFloat(settings.retentionTarget);
  const kappaInput = settings.kappaPriToDelta !== undefined ? parseFloat(settings.kappaPriToDelta) : DEFAULT_SETTINGS.kappaPriToDelta;
  const kappaPriToDelta = isFinite(kappaInput) ? kappaInput : DEFAULT_SETTINGS.kappaPriToDelta;
  const minReadSetting = settings.minD1ReadMin !== undefined ? parseFloat(settings.minD1ReadMin) : DEFAULT_SETTINGS.minD1ReadMin;
  const guideBanditEnabled = settings.banditEnabledForGuide !== undefined ? asBoolean(settings.banditEnabledForGuide) : DEFAULT_SETTINGS.banditEnabledForGuide;

  const todayDate = new Date(today || new Date());
  todayDate.setHours(0, 0, 0, 0);
  const msPerDay = 1000 * 60 * 60 * 24;

  const useWeibull = asBoolean(settings.useWeibull);
  const weibullShape = useWeibull && modelRow && modelRow.weibull_k !== undefined
    ? (isFinite(parseFloat(modelRow.weibull_k)) ? parseFloat(modelRow.weibull_k) : 1)
    : 1;

  let estabilidade = null;
  if (spacedRow && spacedRow.estabilidade !== undefined && spacedRow.estabilidade !== '') {
    const value = parseFloat(spacedRow.estabilidade);
    if (isFinite(value)) {
      estabilidade = value;
    }
  }
  if (!isFinite(estabilidade) && modelRow && modelRow.S_atual !== undefined) {
    const value = parseFloat(modelRow.S_atual);
    if (isFinite(value)) {
      estabilidade = value;
    }
  }
  if (!isFinite(estabilidade)) {
    estabilidade = isFinite(Smin) ? Smin : DEFAULT_SETTINGS.Smin;
  }

  const tempoMedioSec = statsRow && statsRow.tempo_medio !== undefined ? parseFloat(statsRow.tempo_medio) : null;
  const tempoMedioMin = isFinite(tempoMedioSec) && tempoMedioSec > 0 ? tempoMedioSec / 60 : 1;
  const difMedia = statsRow && statsRow.dif_media !== undefined ? parseFloat(statsRow.dif_media) : 3;
  const acerto28 = statsRow && statsRow.acerto_28d !== undefined ? parseFloat(statsRow.acerto_28d) : (statsRow && statsRow.acerto_vida !== undefined ? parseFloat(statsRow.acerto_vida) : 0.65);

  const historyInfo = gatherGuideHistoryInfo(alvoKey, area, subarea, spacedRow, logRows, revisaoRows);
  const lastReviewDate = historyInfo.lastDate ? new Date(historyInfo.lastDate.getTime()) : null;
  const lapses = spacedRow && spacedRow.lapses !== undefined ? parseFloat(spacedRow.lapses) : NaN;

  let daysSinceLast = null;
  if (lastReviewDate instanceof Date && !isNaN(lastReviewDate)) {
    lastReviewDate.setHours(0, 0, 0, 0);
    daysSinceLast = Math.max(0, Math.floor((todayDate - lastReviewDate) / msPerDay));
  }

  let recallToday = null;
  if (daysSinceLast !== null && isFinite(daysSinceLast)) {
    recallToday = calcRecall(daysSinceLast, Math.max(1, estabilidade), weibullShape);
  }
  if (!isFinite(recallToday)) {
    recallToday = 0.4;
  }
  recallToday = clamp(recallToday, 0, 1);

  const modelNEff = modelRow && modelRow.n_eff !== undefined ? parseFloat(modelRow.n_eff) : NaN;
  const hasSheetHistory = !!statsRow || !!spacedRow || !!modelRow;
  const fallbackNEffBase = historyInfo.totalCount;
  const fallbackNEff = fallbackNEffBase > 0 ? fallbackNEffBase : (hasSheetHistory ? 1 : 0);
  const nEff = isFinite(modelNEff) && modelNEff >= 0 ? modelNEff : fallbackNEff;

  const hasHistoryFlag = hasSheetHistory || historyInfo.hasHistory;
  const status = hasHistoryFlag ? 'HISTORICO' : 'NOVO';
  let stage = status === 'NOVO' ? 'D1' : 'S2';
  if (status === 'HISTORICO') {
    stage = determineGuideStage(nEff, daysSinceLast, recallToday, lapses, acerto28, estabilidade);
  }

  const totalMinutes = budgetMinutes;
  const tempoQuest = Math.max(tempoMedioMin, 0.5);

  let planToday;
  let usesBandit = false;
  let rationaleExtra = {};

  if (status === 'NOVO') {
    const minRead = isFinite(minReadSetting) ? Math.max(5, Math.round(minReadSetting)) : 15;
    let readMin = Math.round(totalMinutes * readShare);
    if (readMin < minRead && totalMinutes >= minRead) {
      readMin = minRead;
    }
    if (readMin > totalMinutes) {
      readMin = Math.max(0, Math.min(totalMinutes, minRead));
    }

    let remaining = Math.max(0, totalMinutes - readMin);
    let blockMin = Math.round(remaining * 0.45);
    if (blockMin < 20 && remaining >= 20) {
      blockMin = 20;
    }
    if (blockMin > remaining) {
      blockMin = remaining;
    }
    remaining = Math.max(0, totalMinutes - readMin - blockMin);
    let flashMin = Math.round(remaining);
    if (readMin + blockMin + flashMin !== totalMinutes) {
      const diff = totalMinutes - (readMin + blockMin + flashMin);
      flashMin = Math.max(0, flashMin + diff);
      if (flashMin < 0) {
        blockMin = Math.max(0, blockMin + flashMin);
        flashMin = 0;
      }
      if (blockMin < 0) {
        readMin = Math.max(0, readMin + blockMin);
        blockMin = 0;
      }
    }

    const difFactor = clamp(1 + 0.1 * (difMedia - 3), 0.6, 1.4);
    let acertoFactor = 1;
    if (isFinite(acerto28)) {
      if (acerto28 < 0.6) acertoFactor = 1.1;
      else if (acerto28 > 0.8) acertoFactor = 0.9;
    }
    const flashcardsRaw = flashMin > 0 ? flashRate * flashMin * difFactor * acertoFactor : 0;
    const flashcardsCreate = flashMin > 0 ? Math.round(clamp(flashcardsRaw, 20, 120)) : 0;

    let questionsNew = 0;
    let questionsEstMin = Math.round(blockMin);
    if (blockMin > 0) {
      const maxByTime = Math.max(1, Math.floor(blockMin / tempoQuest));
      questionsNew = Math.max(5, Math.round(Math.min(blockTarget, maxByTime)));
      questionsEstMin = Math.round(questionsNew * tempoQuest);
      if (questionsEstMin > blockMin) {
        questionsEstMin = Math.round(blockMin);
      }
    }

    const deltaRpp = Math.max(0, 100 * kappaPriToDelta * (1 - recallToday));
    const flashcardsCreateMin = Math.round(Math.max(0, flashMin));
    planToday = {
      readMin: Math.round(readMin),
      flashcardsCreate,
      flashcardsReviewMin: 0,
      flashcardsCreateMin,
      questionsNew,
      questionsEstMin: Math.round(Math.max(questionsEstMin, blockMin)),
      deltaRpp,
      totalMin: totalMinutes,
      rationale: {
        base: 'novo',
        stage: 'D1',
        usesBandit: false,
        cardsRate: flashRate,
        difficultyFactor: difFactor,
        accuracyAdj: acertoFactor,
        readShare
      }
    };
    rationaleExtra = planToday.rationale;
  } else {
    let reviewMin = Math.round(Math.min(20, totalMinutes * 0.15));
    if (stage === 'S3') {
      reviewMin = Math.max(reviewMin, Math.round(Math.min(30, totalMinutes * 0.25)));
    }
    if (stage === 'S1') {
      reviewMin = Math.max(reviewMin, Math.round(Math.min(20, totalMinutes * 0.2)));
    }
    reviewMin = Math.min(reviewMin, totalMinutes);

    const remainingAfterReview = Math.max(0, totalMinutes - reviewMin);
    let questionShare = 0.5;
    if (stage === 'S1') questionShare = 0.6;
    if (stage === 'S3') questionShare = 0.35;
    let questionsMin = Math.round(remainingAfterReview * questionShare);
    if (remainingAfterReview > 0 && questionsMin < 15) {
      questionsMin = Math.min(remainingAfterReview, 15);
    }
    if (questionsMin > remainingAfterReview) {
      questionsMin = remainingAfterReview;
    }

    const creationMin = Math.max(0, remainingAfterReview - questionsMin);

    const cardsErroBase = isFinite(acerto28) ? Math.round((1 - acerto28) * 40) : 30;
    const flashcardsCreate = Math.max(10, Math.min(60, cardsErroBase));

    let questionsNew = 0;
    let questionsEstMin = Math.round(questionsMin);
    if (questionsMin > 0) {
      const possible = Math.floor(questionsMin / tempoQuest);
      questionsNew = Math.max(5, possible);
      if (questionsNew <= 0 && questionsMin > 0) {
        questionsNew = Math.max(5, Math.round(questionsMin / tempoQuest));
      }
      questionsEstMin = Math.round(questionsNew * tempoQuest);
      if (questionsEstMin > remainingAfterReview) {
        const adjusted = Math.floor(remainingAfterReview / tempoQuest);
        questionsNew = Math.max(3, adjusted);
        questionsEstMin = Math.round(questionsNew * tempoQuest);
      }
    }

    const horizonDays = determineHorizonDays(todayDate, examConfig);
    const components = guideResolveEviComponents(alvoKey, settings, statsRow, spacedRow, modelRow, reviewEntry, todayDate, horizonDays);
    let deltaPerMin = 0;
    if (guideBanditEnabled && asBoolean(settings.useAdvancedPriority)) {
      const eviLCB = components && components.eviLCBPerMin !== null && components.eviLCBPerMin !== undefined ? parseFloat(components.eviLCBPerMin) : null;
      const eviMean = components && components.eviPerMinMean !== null && components.eviPerMinMean !== undefined ? parseFloat(components.eviPerMinMean) : null;
      const candidate = asBoolean(settings.useGainLCB) ? eviLCB : eviMean;
      if (isFinite(candidate) && candidate > 0) {
        deltaPerMin = candidate;
        usesBandit = true;
      }
    }
    if (!usesBandit) {
      deltaPerMin = Math.max(0, kappaPriToDelta * (1 - recallToday) / Math.max(1, totalMinutes));
    }

    const effectiveMinutes = Math.max(1, totalMinutes);
    const deltaRpp = Math.max(0, Math.round(deltaPerMin * effectiveMinutes * 1000) / 10);

    const creationEstMin = Math.round(creationMin);
    planToday = {
      readMin: 0,
      flashcardsCreate,
      flashcardsReviewMin: Math.round(reviewMin),
      flashcardsCreateMin: creationEstMin,
      questionsNew: Math.max(0, Math.round(questionsNew)),
      questionsEstMin: Math.round(Math.max(questionsEstMin, questionsMin)),
      deltaRpp,
      totalMin: totalMinutes,
      rationale: {
        base: 'historico',
        stage,
        usesBandit,
        reviewShare: reviewMin / Math.max(1, totalMinutes),
        questionShare: questionsMin / Math.max(1, totalMinutes),
        creationShare: creationMin / Math.max(1, totalMinutes),
        fallback: !usesBandit
      }
    };
    rationaleExtra = planToday.rationale;
  }

  if (planToday) {
    const readComponent = planToday.readMin || 0;
    const reviewComponent = planToday.flashcardsReviewMin || 0;
    const createComponent = planToday.flashcardsCreateMin || 0;
    const questionsComponent = planToday.questionsEstMin || 0;
    const sumMinutes = Math.round(readComponent + reviewComponent + createComponent + questionsComponent);
    if (!isFinite(planToday.totalMin) || planToday.totalMin <= 0) {
      planToday.totalMin = sumMinutes;
    } else if (sumMinutes > planToday.totalMin) {
      planToday.totalMin = sumMinutes;
    }
  }

  const perMinuteGainPP = planToday && planToday.totalMin > 0 ? (planToday.deltaRpp || 0) / Math.max(1, planToday.totalMin) : 0;
  const reviewsProjection = simulateGuideReviews(estabilidade, retentionTarget, weibullShape, perMinuteGainPP, planToday.totalMin, todayDate, spacedRow);

  return {
    alvo: alvoKey,
    area,
    subarea,
    status,
    stage,
    diagnostics: {
      S: estabilidade,
      Rhoje: recallToday,
      tempoMedio: tempoMedioMin,
      acerto_28d: acerto28,
      difMedia,
      n_eff: nEff
    },
    planToday,
    nextReviews: reviewsProjection.nextReviews,
    totalDeltaRpp7d: reviewsProjection.total7d,
    totalDeltaRpp28d: reviewsProjection.total28d,
    rationale: rationaleExtra
  };
}

function resolveStatsRow(statsMap, area, subarea) {
  if (!statsMap) return null;
  const key = `${(area || '').toString().trim()}::${(subarea || '').toString().trim()}`;
  if (statsMap[key]) return statsMap[key];
  const lower = key.toLowerCase();
  const matchKey = Object.keys(statsMap).find(k => k && k.toLowerCase() === lower);
  return matchKey ? statsMap[matchKey] : null;
}

function resolveSpacedRow(spacedMap, alvoKey) {
  if (!spacedMap) return null;
  if (spacedMap[alvoKey]) return spacedMap[alvoKey];
  const lower = alvoKey.toLowerCase();
  const matchKey = Object.keys(spacedMap).find(k => k && k.toLowerCase() === lower);
  return matchKey ? spacedMap[matchKey] : null;
}

function resolveModelRow(modelMap, alvoKey) {
  if (!modelMap) return null;
  if (modelMap[alvoKey]) return modelMap[alvoKey];
  const lower = alvoKey.toLowerCase();
  const matchKey = Object.keys(modelMap).find(k => k && k.toLowerCase() === lower);
  return matchKey ? modelMap[matchKey] : null;
}

function determineGuideStage(nEff, daysSinceLast, recallToday, lapses, acerto28, estabilidade) {
  const eff = isFinite(nEff) ? nEff : 0;
  const days = isFinite(daysSinceLast) ? daysSinceLast : null;
  const recall = isFinite(recallToday) ? recallToday : 0.5;
  const laps = isFinite(lapses) ? lapses : 0;
  const acerto = isFinite(acerto28) ? acerto28 : null;
  const est = isFinite(estabilidade) ? estabilidade : null;
  if (eff < 10 || (days !== null && days <= 7)) {
    return 'S1';
  }
  const highRecall = recall >= 0.75;
  const highAcerto = acerto !== null && acerto >= 0.85;
  const highEst = est !== null && est >= 40;
  if ((eff > 30 || (highRecall && highAcerto) || highEst) && laps <= 3) {
    return 'S3';
  }
  return 'S2';
}

function gatherGuideHistoryInfo(alvoKey, area, subarea, spacedRow, logRows, revisaoRows) {
  const info = {
    hasHistory: false,
    totalCount: 0,
    lastDate: null
  };

  const normalizedArea = (area || '').toString().trim().toLowerCase();
  const normalizedSub = (subarea || '').toString().trim().toLowerCase();

  const updateDate = (date) => {
    if (!date || !(date instanceof Date) || isNaN(date)) return;
    if (!info.lastDate || date > info.lastDate) {
      info.lastDate = new Date(date.getTime());
    }
  };

  if (spacedRow) {
    info.hasHistory = true;
    if (spacedRow.ultimaRevisao) {
      const ultima = parseSheetDate(spacedRow.ultimaRevisao);
      if (ultima) updateDate(ultima);
    }
  }

  if (revisaoRows && revisaoRows.length) {
    revisaoRows.forEach(row => {
      if (!row) return;
      const alvoRow = (row.alvo || '').toString().trim().toLowerCase();
      if (alvoRow && alvoRow === alvoKey.toLowerCase()) {
        info.hasHistory = true;
        info.totalCount += 1;
        const data = parseSheetDate(row.data);
        if (data) updateDate(data);
        return;
      }
      if (!normalizedArea && !normalizedSub) return;
      const parts = parseAlvoParts(row.alvo || '');
      const areaMatch = parts.area && parts.area.toLowerCase() === normalizedArea;
      const subMatch = parts.subarea && parts.subarea.toLowerCase() === normalizedSub;
      if ((normalizedArea && areaMatch) || (normalizedSub && subMatch)) {
        info.hasHistory = true;
        info.totalCount += 1;
        const data = parseSheetDate(row.data);
        if (data) updateDate(data);
      }
    });
  }

  if (logRows && logRows.length) {
    logRows.forEach(row => {
      if (!row) return;
      const areaRow = (row.area || '').toString().trim().toLowerCase();
      const subRow = (row.subarea || '').toString().trim().toLowerCase();
      if ((normalizedArea && areaRow === normalizedArea) && (normalizedSub && subRow === normalizedSub)) {
        info.hasHistory = true;
        info.totalCount += 1;
        const data = parseSheetDate(row.data);
        if (data) updateDate(data);
      }
    });
  }

  if (!info.hasHistory && info.totalCount > 0) {
    info.hasHistory = true;
  }

  if (info.totalCount === 0 && info.hasHistory) {
    info.totalCount = 1;
  }

  if (!info.hasHistory) {
    info.totalCount = 0;
  }

  return info;
}

function guideResolveEviComponents(alvoKey, settings, statsRow, spacedRow, modelRow, reviewEntry, today, horizonDays) {
  if (reviewEntry && reviewEntry.components) {
    return {
      eviLCBPerMin: reviewEntry.components.eviPerMin !== undefined ? reviewEntry.components.eviPerMin : null,
      eviPerMinMean: reviewEntry.components.eviPerMinMean !== undefined ? reviewEntry.components.eviPerMinMean : null
    };
  }

  const pseudo = spacedRow ? Object.assign({}, spacedRow) : { alvo: alvoKey };
  if (!pseudo.alvo) pseudo.alvo = alvoKey;
  if (pseudo.estabilidade === undefined || pseudo.estabilidade === '') {
    const SfromModel = modelRow && modelRow.S_atual !== undefined ? parseFloat(modelRow.S_atual) : null;
    pseudo.estabilidade = isFinite(SfromModel) ? SfromModel : settings.Smin;
  }
  if (!pseudo.ultimaRevisao && modelRow && modelRow.ultima_atualizacao) {
    pseudo.ultimaRevisao = modelRow.ultima_atualizacao;
  }

  const extras = {
    modelRow: modelRow || null,
    horizonDays: horizonDays,
    residualValue: null,
    surpriseLambda: settings.lambdaSurprise,
    coverage7d: 0,
    meta7d: settings.coverageTarget7d !== undefined ? parseFloat(settings.coverageTarget7d) : 0
  };

  const result = calculatePriorityForRow(pseudo, statsRow || null, settings, today, extras);
  if (result && result.components) {
    return {
      eviLCBPerMin: result.components.eviPerMin !== undefined ? result.components.eviPerMin : null,
      eviPerMinMean: result.components.eviPerMinMean !== undefined ? result.components.eviPerMinMean : null
    };
  }
  return {
    eviLCBPerMin: null,
    eviPerMinMean: null
  };
}

function simulateGuideReviews(S, retentionTarget, weibullK, deltaPerMinPP, totalMinutes, today, spacedRow) {
  let baseS = isFinite(S) && S > 0 ? S : DEFAULT_SETTINGS.Smin;
  const meta = isFinite(retentionTarget) && retentionTarget > 0 ? retentionTarget : DEFAULT_SETTINGS.retentionTarget;
  const shape = isFinite(weibullK) && weibullK > 0 ? weibullK : 1;
  const basePerMin = isFinite(deltaPerMinPP) ? Math.max(0, deltaPerMinPP) : 0;
  const msPerDay = 1000 * 60 * 60 * 24;
  let firstInterval = calcOptimalInterval(baseS, meta, shape);

  if (spacedRow && spacedRow.proximaRevisao) {
    const proxima = parseSheetDate(spacedRow.proximaRevisao);
    if (proxima) {
      const offset = Math.max(1, Math.round((proxima - today) / msPerDay));
      firstInterval = Math.max(1, offset);
    }
  }

  let firstOffset = isFinite(firstInterval) ? Math.round(firstInterval) : 2;
  if (!isFinite(firstOffset) || firstOffset <= 0) {
    firstOffset = 2;
  }

  const offsets = [
    firstOffset,
    Math.max(firstOffset + 3, Math.round(firstOffset * 2)),
    Math.max(firstOffset + 7, Math.round(firstOffset * 1.5))
  ];

  const uniqueOffsets = [];
  offsets.forEach(off => {
    const val = Math.max(1, Math.round(off));
    if (uniqueOffsets.indexOf(val) === -1) {
      uniqueOffsets.push(val);
    }
  });

  const nextReviews = [];
  let total7d = 0;
  let total28d = 0;
  uniqueOffsets.slice(0, 3).forEach((offset, idx) => {
    const estMin = offset <= 3 ? 15 : 20;
    const tipo = offset <= 3 ? 'curta' : 'media';
    const decay = Math.pow(0.85, idx);
    const delta = Math.max(0, Math.round(basePerMin * estMin * decay * 10) / 10);
    if (offset <= 7) total7d += delta;
    if (offset <= 28) total28d += delta;
    nextReviews.push({ diaOffset: offset, tipo, estMin, deltaRpp: delta });
  });

  return {
    nextReviews,
    total7d,
    total28d
  };
}

function apiStudyGuidePlan(params) {
  try {
    const settings = apiGetSettings();
    const payload = params || {};
    const alvoRaw = (payload.alvo || '').toString().trim();
    if (!alvoRaw) {
      return { ok: false, error: 'Alvo inválido' };
    }

    const budgetInput = parseFloat(payload.budgetD1Min);
    const budgetMin = isFinite(budgetInput) && budgetInput > 0 ? budgetInput : 0;

    const prefsInput = payload.preferencias || {};
    const prefs = {
      flashcardsPerMinBase: prefsInput.flashcardsPerMinBase,
      blockQuestionsTarget: prefsInput.blockQuestionsTarget,
      readShare: prefsInput.readShare
    };

    const planningSettings = applyPlanningModeAdjustments(Object.assign({}, settings), payload.mode);
    const data = buildStudyGuideData(planningSettings);
    const plan = computeStudyGuidePlanForTarget(alvoRaw, budgetMin, prefs, planningSettings, data);
    if (!plan) {
      return { ok: false, error: 'Alvo não encontrado ou sem dados suficientes' };
    }
    return Object.assign({ ok: true }, plan);
  } catch (e) {
    return { ok: false, error: e.toString() };
  }
}

function apiStudyGuideCompareAlvos(params) {
  try {
    const settings = apiGetSettings();
    const payload = params || {};
    const budgetInput = parseFloat(payload.budgetD1Min);
    const budgetMin = isFinite(budgetInput) && budgetInput > 0 ? budgetInput : 0;

    const prefsInput = payload.preferencias || {};
    const prefs = {
      flashcardsPerMinBase: prefsInput.flashcardsPerMinBase,
      blockQuestionsTarget: prefsInput.blockQuestionsTarget,
      readShare: prefsInput.readShare
    };

    const maxTargetsInput = payload.maxTargets !== undefined ? parseInt(payload.maxTargets, 10) : 5;
    const maxTargets = isFinite(maxTargetsInput) && maxTargetsInput > 0 ? maxTargetsInput : 5;

    const areaFilter = (payload.area || '').toString().trim().toLowerCase();
    const subFilter = (payload.subarea || '').toString().trim().toLowerCase();

    const planningSettings = applyPlanningModeAdjustments(Object.assign({}, settings), payload.mode);
    const data = buildStudyGuideData(planningSettings);

    const alvoSet = new Set();
    Object.keys(data.spacedMap || {}).forEach(key => alvoSet.add(key));
    (data.reviewList || []).forEach(entry => {
      if (entry && entry.alvo) {
        alvoSet.add(entry.alvo.toString().trim());
      }
    });

    const results = [];
    alvoSet.forEach(key => {
      if (!key) return;
      const parts = parseAlvoParts(key);
      if (areaFilter && parts.area.toLowerCase() !== areaFilter) return;
      if (subFilter && parts.subarea.toLowerCase() !== subFilter) return;
      const plan = computeStudyGuidePlanForTarget(key, budgetMin, prefs, planningSettings, data);
      if (plan && plan.D1) {
        results.push(plan);
      }
    });

    results.sort((a, b) => {
      const deltaA = a && a.D1 && a.D1.deltaRpp ? Number(a.D1.deltaRpp) : 0;
      const deltaB = b && b.D1 && b.D1.deltaRpp ? Number(b.D1.deltaRpp) : 0;
      return deltaB - deltaA;
    });

    const limited = maxTargets > 0 ? results.slice(0, maxTargets) : results;

    return {
      ok: true,
      budgetD1Min: budgetMin,
      targets: limited
    };
  } catch (e) {
    return { ok: false, error: e.toString() };
  }
}

function apiWeeklyPlan(params) {
  try {
    const settings = apiGetSettings();
    const config = params || {};

    if (!asBoolean(settings.useBanditPlanner)) {
      return { ok: false, disabled: true };
    }

    const planningSettings = applyPlanningModeAdjustments(Object.assign({}, settings), config.mode);
    const useAdvanced = asBoolean(planningSettings.useAdvancedPriority);
    const budgetInput = parseFloat(config.budgetPerDayMin);
    const budgetPerDay = isFinite(budgetInput) && budgetInput > 0 ? budgetInput : 0;
    const roundInput = parseFloat(config.roundTo);
    const roundTo = isFinite(roundInput) && roundInput > 0 ? Math.max(1, Math.round(roundInput)) : 5;
    const maxTargetsInput = parseInt(config.maxTargets, 10);
    const maxTargets = isFinite(maxTargetsInput) && maxTargetsInput > 0 ? maxTargetsInput : 12;
    const kappaInput = planningSettings.kappaPriToDelta !== undefined && planningSettings.kappaPriToDelta !== ''
      ? parseFloat(planningSettings.kappaPriToDelta)
      : DEFAULT_SETTINGS.kappaPriToDelta;
    const kappaPriToDelta = isFinite(kappaInput) ? kappaInput : DEFAULT_SETTINGS.kappaPriToDelta;

    const hoje = new Date();
    hoje.setHours(0, 0, 0, 0);

    const gather = gatherReviewCandidates(planningSettings, hoje);
    const reviewList = gather.reviewList || [];
    const statsData = readSheetData(SHEET_NAMES.STATS);
    const statsMap = {};
    statsData.forEach(row => {
      if (!row) return;
      const key = `${row.area}::${row.subarea}`;
      statsMap[key] = row;
    });

    const candidates = reviewList.map((entry, idx) => createPlannerItem(entry, statsMap, planningSettings, idx)).filter(Boolean);
    if (candidates.length === 0) {
      return {
        ok: true,
        mode: useAdvanced ? 'advanced' : 'classic',
        budgetPerDay: budgetPerDay,
        days: [],
        carryOver: []
      };
    }

    candidates.forEach(item => {
      const due = item.dueDate instanceof Date ? item.dueDate : (item.dueDate ? parseSheetDate(item.dueDate) : null);
      if (due instanceof Date && !isNaN(due)) {
        due.setHours(0, 0, 0, 0);
        item.dueDate = due;
      } else {
        item.dueDate = null;
      }
      item._assigned = false;
    });

    candidates.sort((a, b) => {
      const dueA = a.dueDate instanceof Date ? a.dueDate.getTime() : Number.MAX_SAFE_INTEGER;
      const dueB = b.dueDate instanceof Date ? b.dueDate.getTime() : Number.MAX_SAFE_INTEGER;
      if (dueA !== dueB) return dueA - dueB;
      return (b.prioridade || 0) - (a.prioridade || 0);
    });

    const days = [];
    const msPerDay = 24 * 60 * 60 * 1000;
    for (let offset = 0; offset < 7; offset++) {
      const currentDate = new Date(hoje.getTime() + offset * msPerDay);
      const available = candidates.filter(item => !item._assigned && (item.dueDate === null || item.dueDate <= currentDate));
      let selection = available.slice();
      if (selection.length === 0) {
        const remaining = candidates.filter(item => !item._assigned);
        selection = remaining.slice(0, maxTargets);
      } else if (selection.length > maxTargets) {
        selection = selection.slice(0, maxTargets);
      }

      const planInputs = selection.map(item => Object.assign({}, item));
      const dayPlan = buildProportionalPlan(planInputs, {
        budgetMin: budgetPerDay,
        roundTo,
        useAdvanced,
        kappaPriToDelta
      }, planningSettings);

      dayPlan.targets.forEach(target => {
        if (!target || !target.alvo) return;
        if ((target.allocMin || 0) <= 0) return;
        const original = candidates.find(item => !item._assigned && item.alvo === target.alvo);
        if (original) {
          original._assigned = true;
        }
      });

      const areaAgg = {};
      (dayPlan.targets || []).forEach(target => {
        const alloc = target.allocMin || 0;
        if (alloc <= 0) return;
        const delta = target.deltaRpp || 0;
        const key = target.area || 'Sem área';
        if (!areaAgg[key]) {
          areaAgg[key] = { area: key, allocMin: 0, deltaRpp: 0 };
        }
        areaAgg[key].allocMin += alloc;
        areaAgg[key].deltaRpp += delta;
      });

      const areas = Object.keys(areaAgg).map(key => ({
        area: key,
        allocMin: Math.round(areaAgg[key].allocMin),
        deltaRpp: Math.round(areaAgg[key].deltaRpp * 10) / 10
      })).sort((a, b) => (b.allocMin || 0) - (a.allocMin || 0));

      days.push({
        dateISO: currentDate.toISOString(),
        date: formatDateDDMMYYYY(currentDate),
        allocatedMin: Math.round(dayPlan.totalAlloc),
        totalDeltaRpp: dayPlan.totalDelta,
        targets: dayPlan.targets,
        areas,
        fallbackToPriority: dayPlan.fallback
      });
    }

    const carryOver = candidates
      .filter(item => !item._assigned)
      .map(item => ({
        alvo: item.alvo,
        area: item.area,
        prioridade: item.prioridade,
        dueDate: item.dueDate instanceof Date ? formatDateDDMMYYYY(item.dueDate) : '',
        score: useAdvanced ? item.scoreAdvanced : item.scoreClassic
      }));

    return {
      ok: true,
      mode: useAdvanced ? 'advanced' : 'classic',
      budgetPerDay: budgetPerDay,
      days,
      carryOver
    };
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

    const useWeibull = asBoolean(settings.useWeibull);

    const spacedSheet = getOrCreateSheet(SHEET_NAMES.SPACED, HEADERS.SPACED);
    const spacedData = readSheetData(SHEET_NAMES.SPACED);
    const spacedIdx = spacedData.findIndex(row => row.alvo === alvo);
    const spacedRow = spacedIdx >= 0 ? spacedData[spacedIdx] : null;
    const baseEstabilidade = spacedRow ? parseFloat(spacedRow.estabilidade) : NaN;
    const estabilidadeAnterior = isFinite(baseEstabilidade) && baseEstabilidade > 0 ? baseEstabilidade : settings.Smin;
    const weibullKBefore = useWeibull ? getWeibullShape(area, settings) : 1;

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
      const estimativa = calcOptimalInterval(baseS, metaUsada, weibullKBefore);
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
    const S_obs = Math.max(calcSobs(safeTDias, metaUsada, weibullKBefore), settings.Smin / 4);
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

    if (useWeibull) {
      const sampleK = estimateWeibullSample(safeTDias, metaUsada, Math.max(1, estabilidadeAnterior));
      if (sampleK) {
        updateWeibullShape(area, sampleK, settings);
      }
    }

    const weibullKAfter = useWeibull ? getWeibullShape(area, settings) : 1;

    let I = calcOptimalInterval(S_pred, metaUsada, weibullKAfter);
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
        modelRow: { sigma: sigmaAtual, n_eff: nEffAtual, weibull_k: weibullKAfter },
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
      nEffAtual,
      weibullKAfter
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
      pri: prioridade,
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
      const alvoParts = parseAlvoParts(alvo);
      const weibullK = asBoolean(settings.useWeibull) ? getWeibullShape(alvoParts.area, settings) : 1;

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
      const S_obs = Math.max(calcSobs(tDias, metaUsada, weibullK), settings.Smin / 4);
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
      const alvoParts = parseAlvoParts(alvo);
      const weibullK = asBoolean(settings.useWeibull) ? getWeibullShape(alvoParts.area, settings) : 1;
      return [
        alvo,
        state.theta[0],
        state.theta[1],
        state.theta[2],
        state.S_atual,
        new Date(),
        sigma,
        state.nEff,
        weibullK
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
      const alvoParts = parseAlvoParts(alvo);
      const weibullK = asBoolean(settings.useWeibull) ? getWeibullShape(alvoParts.area, settings) : 1;
      let I = calcOptimalInterval(S_novo, settings.retentionTarget, weibullK);
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
          modelRow: {
            sigma: Math.sqrt(Math.max(1e-6, modelState.sigma2)),
            n_eff: modelState.nEff,
            weibull_k: asBoolean(settings.useWeibull)
              ? getWeibullShape(alvoParts.area, settings)
              : 1
          },
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

function apiFitDoseResponse(alvoOrAll) {
  try {
    const settings = apiGetSettings();
    if (!asBoolean(settings.useAdvancedPriority)) {
      return { ok: false, disabled: true };
    }

    const logData = readSheetData(SHEET_NAMES.LOG);
    if (!logData || logData.length === 0) {
      return { ok: false, error: 'Sem dados suficientes' };
    }

    const target = (alvoOrAll || '').toString().trim();
    const isAll = !target || target.toLowerCase() === 'all';

    const byTarget = {};
    logData.forEach(row => {
      if (!row) return;
      const area = (row.area || '').toString().trim();
      const subarea = (row.subarea || '').toString().trim();
      if (!area || !subarea) return;
      const alvo = `${area}::${subarea}`;
      if (!isAll && alvo !== target) return;
      const total = Number(row.total) || 0;
      const acertos = Number(row.acertos) || 0;
      if (total <= 0) return;
      if (!byTarget[alvo]) {
        byTarget[alvo] = [];
      }
      byTarget[alvo].push({ effort: total, gain: acertos / total });
    });

    const results = Object.keys(byTarget).map(key => {
      const points = byTarget[key];
      if (!points || points.length < 2) {
        return { alvo: key, slope: 0, intercept: points && points.length === 1 ? points[0].gain : 0, n: points.length };
      }
      let sumX = 0;
      let sumY = 0;
      let sumXY = 0;
      let sumXX = 0;
      points.forEach(p => {
        sumX += p.effort;
        sumY += p.gain;
        sumXY += p.effort * p.gain;
        sumXX += p.effort * p.effort;
      });
      const n = points.length;
      const denom = (n * sumXX) - (sumX * sumX);
      const slope = denom !== 0 ? ((n * sumXY) - (sumX * sumY)) / denom : 0;
      const intercept = (sumY - slope * sumX) / n;
      return { alvo: key, slope, intercept, n };
    });

    if (!isAll) {
      return results.length > 0 ? { ok: true, result: results[0] } : { ok: false, error: 'Alvo sem dados' };
    }

    return { ok: true, results };
  } catch (e) {
    return { ok: false, error: e.toString() };
  }
}

function apiPredictGain(alvo, effortPlan) {
  try {
    const settings = apiGetSettings();
    if (!asBoolean(settings.useAdvancedPriority)) {
      return { ok: false, disabled: true };
    }
    if (!alvo) {
      return { ok: false, error: 'Informe o alvo' };
    }

    const fit = apiFitDoseResponse(alvo);
    if (!fit || !fit.ok || !fit.result) {
      return { ok: false, error: 'Não foi possível ajustar dose-resposta' };
    }

    const effort = isFinite(effortPlan) ? Number(effortPlan) : 10;
    const slope = fit.result.slope || 0;
    const intercept = fit.result.intercept || 0;
    const predicted = intercept + slope * effort;
    return { ok: true, alvo, effort, predictedGain: clamp(predicted, 0, 1), model: fit.result };
  } catch (e) {
    return { ok: false, error: e.toString() };
  }
}

function apiEstimateAttribution(window) {
  try {
    const settings = apiGetSettings();
    if (!asBoolean(settings.useAdvancedPriority)) {
      return { ok: false, disabled: true };
    }

    const windowDays = Math.max(7, parseInt(window, 10) || 28);
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const cutoff = new Date(today.getTime());
    cutoff.setDate(cutoff.getDate() - windowDays);

    const logData = readSheetData(SHEET_NAMES.LOG);
    if (!logData || logData.length === 0) {
      return { ok: false, error: 'Sem dados para estimar efeitos' };
    }

    const effectsMap = {};
    logData.forEach(row => {
      if (!row) return;
      const area = (row.area || '').toString().trim();
      const subarea = (row.subarea || '').toString().trim();
      if (!area || !subarea) return;
      const alvo = `${area}::${subarea}`;
      const total = Number(row.total) || 0;
      const acertos = Number(row.acertos) || 0;
      if (total <= 0) return;
      const data = parseSheetDate(row.data);
      const bucket = data && data >= cutoff ? 'recent' : 'baseline';
      if (!effectsMap[alvo]) {
        effectsMap[alvo] = { recent: { total: 0, acertos: 0 }, baseline: { total: 0, acertos: 0 } };
      }
      effectsMap[alvo][bucket].total += total;
      effectsMap[alvo][bucket].acertos += acertos;
    });

    const rows = Object.keys(effectsMap).map(alvo => {
      const stats = effectsMap[alvo];
      const recentTotal = stats.recent.total || 0;
      const baselineTotal = stats.baseline.total || 0;
      const recentRate = recentTotal > 0 ? stats.recent.acertos / recentTotal : 0;
      const baselineRate = baselineTotal > 0 ? stats.baseline.acertos / baselineTotal : recentRate;
      const ate = (recentRate - baselineRate) * 100;
      const seRecent = recentTotal > 0 ? (recentRate * (1 - recentRate)) / recentTotal : 0;
      const seBase = baselineTotal > 0 ? (baselineRate * (1 - baselineRate)) / baselineTotal : 0;
      const se = Math.sqrt(Math.max(0, seRecent + seBase));
      const margin = 1.96 * se * 100;
      const lo = ate - margin;
      const hi = ate + margin;
      const updated = new Date();
      return [
        alvo,
        ate,
        lo,
        hi,
        Math.min(recentTotal, baselineTotal),
        updated
      ];
    });

    const sheet = getOrCreateSheet(SHEET_NAMES.EFFECTS, HEADERS.EFFECTS);
    clearSheetData(SHEET_NAMES.EFFECTS);
    if (rows.length > 0) {
      const startRow = sheet.getLastRow() + 1;
      sheet.getRange(startRow, 1, rows.length, HEADERS.EFFECTS.length).setValues(rows);
      sheet.getRange(startRow, HEADERS.EFFECTS.length, rows.length, 1).setNumberFormat('dd/mm/yyyy');
    }

    return { ok: true, updated: rows.length };
  } catch (e) {
    return { ok: false, error: e.toString() };
  }
}

function hashUidToInt(uid) {
  const str = (uid || '').toString();
  let hash = 0;
  for (let i = 0; i < str.length; i++) {
    hash = ((hash << 5) - hash) + str.charCodeAt(i);
    hash |= 0;
  }
  return Math.abs(hash);
}

function approxZPValue(z) {
  const absZ = Math.abs(z);
  const t = 1 / (1 + 0.2316419 * absZ);
  const d = Math.exp(-0.5 * absZ * absZ) / Math.sqrt(2 * Math.PI);
  const prob = d * t * (0.31938153 + t * (-0.356563782 + t * (1.781477937 + t * (-1.821255978 + t * 1.330274429))));
  return clamp(2 * prob, 0, 1);
}

function approxChiSquarePValue(chi2) {
  const x = Math.max(chi2, 0);
  return Math.exp(-0.5 * x);
}

function apiABAssign(uid) {
  try {
    const settings = apiGetSettings();
    if (!asBoolean(settings.useABTesting)) {
      return { ok: false, disabled: true };
    }
    if (!uid) {
      return { ok: false, error: 'uid obrigatório' };
    }
    const hash = hashUidToInt(uid);
    const variant = hash % 2 === 0 ? 'classic' : 'evi';
    return { ok: true, uid, variant };
  } catch (e) {
    return { ok: false, error: e.toString() };
  }
}

function apiABLog(payload) {
  try {
    const settings = apiGetSettings();
    if (!asBoolean(settings.useABTesting)) {
      return { ok: false, disabled: true };
    }
    if (!payload || !payload.policyVersion) {
      return { ok: false, error: 'policyVersion obrigatório' };
    }
    const stats = loadAbStats();
    const variantKey = payload.policyVersion === 'classic' ? 'classic' : 'evi';
    const outcome = isFinite(payload.outcome) ? Number(payload.outcome) : null;
    if (outcome !== null) {
      const record = stats[variantKey] || { count: 0, sum: 0, sumSquares: 0 };
      record.count += 1;
      record.sum += outcome;
      record.sumSquares += outcome * outcome;
      stats[variantKey] = record;
      saveAbStats(stats);
    }

    appendPolicyLogEntries([{
      timestamp: new Date(),
      alvo: payload.uid || '',
      area: '',
      subarea: '',
      pri: '',
      eviPerMin: '',
      overdue: '',
      diversity: '',
      custos: payload.endpoint || '',
      tempoPrev: '',
      decisao: `ab_${variantKey}`,
      policyVersion: payload.policyVersion
    }]);

    return { ok: true };
  } catch (e) {
    return { ok: false, error: e.toString() };
  }
}

function apiABReport() {
  try {
    const settings = apiGetSettings();
    if (!asBoolean(settings.useABTesting)) {
      return { ok: false, disabled: true };
    }
    const stats = loadAbStats();
    const classic = stats.classic || { count: 0, sum: 0, sumSquares: 0 };
    const evi = stats.evi || { count: 0, sum: 0, sumSquares: 0 };

    const meanClassic = classic.count > 0 ? classic.sum / classic.count : 0;
    const meanEvi = evi.count > 0 ? evi.sum / evi.count : 0;
    const varClassic = classic.count > 1
      ? Math.max(0, (classic.sumSquares - (classic.sum * classic.sum) / classic.count) / (classic.count - 1))
      : meanClassic * (1 - meanClassic);
    const varEvi = evi.count > 1
      ? Math.max(0, (evi.sumSquares - (evi.sum * evi.sum) / evi.count) / (evi.count - 1))
      : meanEvi * (1 - meanEvi);

    const denom = Math.sqrt(Math.max(1e-6, (varClassic / Math.max(1, classic.count)) + (varEvi / Math.max(1, evi.count))));
    const z = denom > 0 ? (meanEvi - meanClassic) / denom : 0;
    const pZ = approxZPValue(z);

    const successClassic = classic.sum;
    const successEvi = evi.sum;
    const failClassic = Math.max(0, classic.count - successClassic);
    const failEvi = Math.max(0, evi.count - successEvi);
    const total = successClassic + successEvi + failClassic + failEvi;
    const expectedClassic = ((successClassic + failClassic) * (successClassic + successEvi)) / Math.max(1, total);
    const expectedEvi = ((successEvi + failEvi) * (successClassic + successEvi)) / Math.max(1, total);
    const expectedClassicFail = ((successClassic + failClassic) * (failClassic + failEvi)) / Math.max(1, total);
    const expectedEviFail = ((successEvi + failEvi) * (failClassic + failEvi)) / Math.max(1, total);
    let chi2 = 0;
    if (expectedClassic > 0) chi2 += Math.pow(successClassic - expectedClassic, 2) / expectedClassic;
    if (expectedEvi > 0) chi2 += Math.pow(successEvi - expectedEvi, 2) / expectedEvi;
    if (expectedClassicFail > 0) chi2 += Math.pow(failClassic - expectedClassicFail, 2) / expectedClassicFail;
    if (expectedEviFail > 0) chi2 += Math.pow(failEvi - expectedEviFail, 2) / expectedEviFail;
    const pChi = approxChiSquarePValue(chi2);

    return {
      ok: true,
      variants: {
        classic: classic,
        evi: evi
      },
      comparison: {
        z,
        pZ,
        chi2,
        pChi
      }
    };
  } catch (e) {
    return { ok: false, error: e.toString() };
  }
}