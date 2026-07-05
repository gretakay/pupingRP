const SPREADSHEET_ID = '1uFOUf4oeVafkBxjkxOet2SEe4vAbkLBO902TxzMgKEc';
const RESPONSES_SHEET_NAME = '表單回覆 1';
const SUMMARY_SHEET_NAME = '進度摘要表';

const SUMMARY_HEADERS = [
  '姓名',
  '法名(沒有不必填寫)',
  '手機末四碼',
  '第一關狀態',
  '第一關分數',
  '第一關複習次數',
  '第二關狀態',
  '第二關複習次數',
  '最後完成進度',
  '最後更新時間'
];

const COLUMN_NAMES = {
  timestamp: '時間戳記',
  playerName: '姓名',
  dharmaName: '法名(沒有不必填寫)',
  phoneLast4: '手機末四碼',
  stageCompleted: '完成進度',
  monopolyScore: '大富翁分數',
  stageId: '關卡代碼',
  eventType: '事件類型'
};

function doGet(e) {
  const params = e && e.parameter ? e.parameter : {};
  const callback = String(params.callback || '').trim();
  const playerName = normalizeText(params.playerName);
  const dharmaName = normalizeText(params.dharmaName);
  const phoneLast4 = normalizePhoneLast4(params.phoneLast4);

  const payload = buildProgressPayload(playerName, dharmaName, phoneLast4);
  const text = JSON.stringify(payload);

  if (callback) {
    return ContentService
      .createTextOutput(`${callback}(${text});`)
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(text)
    .setMimeType(ContentService.MimeType.JSON);
}

function buildProgressPayload(playerName, dharmaName, phoneLast4) {
  if (!playerName || phoneLast4.length !== 4) {
    return {
      ok: false,
      found: false,
      error: 'MISSING_IDENTITY'
    };
  }

  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(RESPONSES_SHEET_NAME);
  if (!sheet) {
    return {
      ok: false,
      found: false,
      error: 'SHEET_NOT_FOUND'
    };
  }

  const summary = getSummaryRecord(playerName, dharmaName, phoneLast4);
  if (!summary) {
    syncProgressSummaryFromResponses();
  }

  const refreshedSummary = getSummaryRecord(playerName, dharmaName, phoneLast4);
  if (!refreshedSummary) {
    return createEmptyPayload(playerName, dharmaName, phoneLast4);
  }

  return {
    ok: true,
    found: true,
    playerName,
    dharmaName,
    phoneLast4,
    progress: {
      DressCode: refreshedSummary.firstStageStatus === 'Pass' ? 'Pass' : '',
      MonopolyScore: Number(refreshedSummary.firstStageScore || 0),
      DressCodeReviewCount: Number(refreshedSummary.firstStageReviewCount || 0),
      'reception-duty': refreshedSummary.secondStageStatus === 'Completed' ? 'Completed' : '',
      'reception-dutyReviewCount': Number(refreshedSummary.secondStageReviewCount || 0)
    },
    lastUpdated: refreshedSummary.lastUpdated || ''
  };
}

function onFormSubmit() {
  syncProgressSummaryFromResponses();
}

function syncProgressSummaryFromResponses() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const responsesSheet = spreadsheet.getSheetByName(RESPONSES_SHEET_NAME);
  if (!responsesSheet) {
    throw new Error('找不到表單回覆工作表');
  }

  const values = responsesSheet.getDataRange().getDisplayValues();
  const summarySheet = getOrCreateSummarySheet(spreadsheet);
  if (values.length < 2) {
    resetSummarySheet(summarySheet);
    return;
  }

  const headers = values[0];
  const rows = values.slice(1);
  const columnIndexes = mapColumnIndexes(headers);
  const summaryMap = new Map();

  rows
    .map((row) => toRowObject(row, columnIndexes))
    .forEach((row) => {
      const playerName = normalizeText(row.playerName);
      const phoneLast4 = normalizePhoneLast4(row.phoneLast4);
      if (!playerName || phoneLast4.length !== 4) {
        return;
      }

      const key = buildSummaryKey(playerName, row.dharmaName, phoneLast4);
      const existing = summaryMap.get(key) || createSummaryRecord(row);
      applyProgressRowToSummary(existing, row);
      summaryMap.set(key, existing);
    });

  writeSummarySheet(summarySheet, Array.from(summaryMap.values()));
}

function createEmptyPayload(playerName, dharmaName, phoneLast4) {
  return {
    ok: true,
    found: false,
    playerName,
    dharmaName,
    phoneLast4,
    progress: {},
    lastUpdated: ''
  };
}

function mapColumnIndexes(headers) {
  return {
    timestamp: headers.indexOf(COLUMN_NAMES.timestamp),
    playerName: headers.indexOf(COLUMN_NAMES.playerName),
    dharmaName: headers.indexOf(COLUMN_NAMES.dharmaName),
    phoneLast4: headers.indexOf(COLUMN_NAMES.phoneLast4),
    stageCompleted: headers.indexOf(COLUMN_NAMES.stageCompleted),
    monopolyScore: headers.indexOf(COLUMN_NAMES.monopolyScore),
    stageId: headers.indexOf(COLUMN_NAMES.stageId),
    eventType: headers.indexOf(COLUMN_NAMES.eventType)
  };
}

function toRowObject(row, columnIndexes) {
  return {
    timestamp: getCell(row, columnIndexes.timestamp),
    playerName: getCell(row, columnIndexes.playerName),
    dharmaName: getCell(row, columnIndexes.dharmaName),
    phoneLast4: getCell(row, columnIndexes.phoneLast4),
    stageCompleted: getCell(row, columnIndexes.stageCompleted),
    monopolyScore: getCell(row, columnIndexes.monopolyScore),
    stageId: getCell(row, columnIndexes.stageId),
    eventType: getCell(row, columnIndexes.eventType)
  };
}

function getCell(row, index) {
  return index >= 0 ? String(row[index] || '').trim() : '';
}

function buildProgressState(rows) {
  const state = {};

  rows.forEach((row) => {
    const stageCompleted = normalizeText(row.stageCompleted);
    const stageId = normalizeText(row.stageId);
    const eventType = normalizeText(row.eventType);
    const monopolyScore = Number(row.monopolyScore || 0);

    if (eventType === 'monopoly_pass' || stageCompleted === '第一關完成') {
      state.DressCode = 'Pass';
      state.MonopolyScore = Math.max(Number(state.MonopolyScore || 0), monopolyScore || 75);
    }

    if (eventType === 'monopoly_review') {
      state.DressCode = 'Pass';
    }

    if (eventType === 'stage_pass' && stageId === 'reception-duty') {
      state['reception-duty'] = 'Completed';
    }

    if (eventType === 'stage_review' && stageId === 'reception-duty') {
      state['reception-duty'] = 'Completed';
    }

    if (stageCompleted === '第二關完成') {
      state['reception-duty'] = 'Completed';
    }

    const firstStageReviewCount = extractReviewCount(stageCompleted, '第一關');
    if (firstStageReviewCount > 0) {
      state.DressCode = 'Pass';
      state.DressCodeReviewCount = Math.max(Number(state.DressCodeReviewCount || 0), firstStageReviewCount);
      state.MonopolyScore = Math.max(Number(state.MonopolyScore || 0), monopolyScore || 75);
    }

    const secondStageReviewCount = extractReviewCount(stageCompleted, '第二關');
    if (secondStageReviewCount > 0) {
      state['reception-duty'] = 'Completed';
      state['reception-dutyReviewCount'] = Math.max(Number(state['reception-dutyReviewCount'] || 0), secondStageReviewCount);
    }
  });

  return state;
}

function extractReviewCount(stageCompleted, stageLabel) {
  const match = String(stageCompleted || '').match(new RegExp(`${stageLabel}複習第(\\d+)次完成`));
  return match ? Number(match[1]) || 0 : 0;
}

function normalizeText(value) {
  return String(value || '').trim();
}

function normalizePhoneLast4(value) {
  return String(value || '').replace(/\D/g, '').slice(0, 4);
}

function getSummaryRecord(playerName, dharmaName, phoneLast4) {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const summarySheet = spreadsheet.getSheetByName(SUMMARY_SHEET_NAME);
  if (!summarySheet) {
    return null;
  }

  const values = summarySheet.getDataRange().getDisplayValues();
  if (values.length < 2) {
    return null;
  }

  const rows = values.slice(1);
  const summaryKey = buildSummaryKey(playerName, dharmaName, phoneLast4);
  const found = rows.find((row) => buildSummaryKey(row[0], row[1], row[2]) === summaryKey);
  if (!found) {
    return null;
  }

  return {
    playerName: found[0],
    dharmaName: found[1],
    phoneLast4: found[2],
    firstStageStatus: found[3],
    firstStageScore: found[4],
    firstStageReviewCount: found[5],
    secondStageStatus: found[6],
    secondStageReviewCount: found[7],
    lastStageCompleted: found[8],
    lastUpdated: found[9]
  };
}

function getOrCreateSummarySheet(spreadsheet) {
  const summarySheet = spreadsheet.getSheetByName(SUMMARY_SHEET_NAME) || spreadsheet.insertSheet(SUMMARY_SHEET_NAME);
  if (summarySheet.getLastRow() === 0) {
    summarySheet.getRange(1, 1, 1, SUMMARY_HEADERS.length).setValues([SUMMARY_HEADERS]);
  } else {
    summarySheet.getRange(1, 1, 1, SUMMARY_HEADERS.length).setValues([SUMMARY_HEADERS]);
  }
  return summarySheet;
}

function resetSummarySheet(summarySheet) {
  summarySheet.clearContents();
  summarySheet.getRange(1, 1, 1, SUMMARY_HEADERS.length).setValues([SUMMARY_HEADERS]);
}

function writeSummarySheet(summarySheet, records) {
  resetSummarySheet(summarySheet);
  if (!records.length) {
    return;
  }

  const rows = records
    .sort((left, right) => {
      return normalizeText(left.playerName).localeCompare(normalizeText(right.playerName), 'zh-Hant');
    })
    .map((record) => {
      return [
        record.playerName,
        record.dharmaName,
        record.phoneLast4,
        record.firstStageStatus,
        record.firstStageScore,
        record.firstStageReviewCount,
        record.secondStageStatus,
        record.secondStageReviewCount,
        record.lastStageCompleted,
        record.lastUpdated
      ];
    });

  summarySheet.getRange(2, 1, rows.length, SUMMARY_HEADERS.length).setValues(rows);
  summarySheet.autoResizeColumns(1, SUMMARY_HEADERS.length);
}

function createSummaryRecord(row) {
  return {
    playerName: normalizeText(row.playerName),
    dharmaName: normalizeText(row.dharmaName),
    phoneLast4: normalizePhoneLast4(row.phoneLast4),
    firstStageStatus: '',
    firstStageScore: 0,
    firstStageReviewCount: 0,
    secondStageStatus: '',
    secondStageReviewCount: 0,
    lastStageCompleted: '',
    lastUpdated: ''
  };
}

function applyProgressRowToSummary(summary, row) {
  const stageCompleted = normalizeText(row.stageCompleted);
  const eventType = normalizeText(row.eventType);
  const stageId = normalizeText(row.stageId);
  const monopolyScore = Number(row.monopolyScore || 0);

  if (eventType === 'monopoly_pass' || stageCompleted === '第一關完成') {
    summary.firstStageStatus = 'Pass';
    summary.firstStageScore = Math.max(Number(summary.firstStageScore || 0), monopolyScore || 75);
  }

  const firstStageReviewCount = extractReviewCount(stageCompleted, '第一關');
  if (eventType === 'monopoly_review' || firstStageReviewCount > 0) {
    summary.firstStageStatus = 'Pass';
    summary.firstStageReviewCount = Math.max(Number(summary.firstStageReviewCount || 0), firstStageReviewCount || 1);
    summary.firstStageScore = Math.max(Number(summary.firstStageScore || 0), monopolyScore || 75);
  }

  if (eventType === 'stage_pass' && stageId === 'reception-duty' || stageCompleted === '第二關完成') {
    summary.secondStageStatus = 'Completed';
  }

  const secondStageReviewCount = extractReviewCount(stageCompleted, '第二關');
  if (eventType === 'stage_review' && stageId === 'reception-duty' || secondStageReviewCount > 0) {
    summary.secondStageStatus = 'Completed';
    summary.secondStageReviewCount = Math.max(Number(summary.secondStageReviewCount || 0), secondStageReviewCount || 1);
  }

  if (stageCompleted) {
    summary.lastStageCompleted = stageCompleted;
  }

  if (row.timestamp) {
    summary.lastUpdated = row.timestamp;
  }
}

function buildSummaryKey(playerName, dharmaName, phoneLast4) {
  return [
    normalizeText(playerName),
    normalizeText(dharmaName),
    normalizePhoneLast4(phoneLast4)
  ].join('::');
}