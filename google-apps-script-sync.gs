const SPREADSHEET_ID = '1uFOUf4oeVafkBxjkxOet2SEe4vAbkLBO902TxzMgKEc';
const RESPONSES_SHEET_NAME = '表單回覆 1';

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

  const values = sheet.getDataRange().getDisplayValues();
  if (values.length < 2) {
    return createEmptyPayload(playerName, dharmaName, phoneLast4);
  }

  const headers = values[0];
  const rows = values.slice(1);
  const columnIndexes = mapColumnIndexes(headers);

  const matchedRows = rows
    .map((row) => toRowObject(row, columnIndexes))
    .filter((row) => {
      return normalizeText(row.playerName) === playerName &&
        normalizeText(row.dharmaName) === dharmaName &&
        normalizePhoneLast4(row.phoneLast4) === phoneLast4;
    });

  if (!matchedRows.length) {
    return createEmptyPayload(playerName, dharmaName, phoneLast4);
  }

  matchedRows.sort((left, right) => {
    return new Date(left.timestamp || 0).getTime() - new Date(right.timestamp || 0).getTime();
  });

  const progress = buildProgressState(matchedRows);
  const lastRow = matchedRows[matchedRows.length - 1];

  return {
    ok: true,
    found: true,
    playerName,
    dharmaName,
    phoneLast4,
    progress,
    lastUpdated: lastRow.timestamp || ''
  };
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