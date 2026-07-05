const SPREADSHEET_ID = '1uFOUf4oeVafkBxjkxOet2SEe4vAbkLBO902TxzMgKEc';
const RESPONSES_SHEET_NAME = '表單回覆 1';
const SUMMARY_SHEET_NAME = '進度摘要表';

const STAGE_DEFINITIONS = [
  {
    id: 'DressCode',
    label: '第一關',
    passValue: 'Pass',
    scoreKey: 'MonopolyScore'
  },
  {
    id: 'reception-duty',
    label: '第二關',
    passValue: 'Completed'
  },
  {
    id: 'desk-duty',
    label: '第三關',
    passValue: 'Completed'
  }
];

const SUMMARY_HEADERS = [
  '姓名',
  '法名(沒有不必填寫)',
  '手機末四碼',
  ...STAGE_DEFINITIONS.flatMap((stage) => {
    const columns = [`${stage.label}狀態`];
    if (stage.scoreKey) {
      columns.push(`${stage.label}分數`);
    }
    columns.push(`${stage.label}複習次數`);
    return columns;
  }),
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
    progress: buildProgressFromSummary(refreshedSummary),
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
  const headers = values[0];
  const summaryIndexes = mapSummaryIndexes(headers);
  const summaryKey = buildSummaryKey(playerName, dharmaName, phoneLast4);
  const found = rows.find((row) => buildSummaryKey(
    row[summaryIndexes.playerName],
    row[summaryIndexes.dharmaName],
    row[summaryIndexes.phoneLast4]
  ) === summaryKey);
  if (!found) {
    return null;
  }

  return summaryRowToRecord(found, summaryIndexes);
}

function getOrCreateSummarySheet(spreadsheet) {
  const summarySheet = spreadsheet.getSheetByName(SUMMARY_SHEET_NAME) || spreadsheet.insertSheet(SUMMARY_SHEET_NAME);
  if (summarySheet.getLastRow() === 0) {
    summarySheet.getRange(1, 1, 1, SUMMARY_HEADERS.length).setValues([SUMMARY_HEADERS]);
  } else {
    summarySheet.getRange(1, 1, 1, SUMMARY_HEADERS.length).setValues([SUMMARY_HEADERS]);
  }
  applySummaryColumnFormats(summarySheet);
  return summarySheet;
}

function resetSummarySheet(summarySheet) {
  summarySheet.clearContents();
  summarySheet.getRange(1, 1, 1, SUMMARY_HEADERS.length).setValues([SUMMARY_HEADERS]);
  applySummaryColumnFormats(summarySheet);
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
    .map((record) => summaryRecordToRow(record));

  summarySheet.getRange(2, 1, rows.length, SUMMARY_HEADERS.length).setValues(rows);
  summarySheet.autoResizeColumns(1, SUMMARY_HEADERS.length);
}

function createSummaryRecord(row) {
  const record = {
    playerName: normalizeText(row.playerName),
    dharmaName: normalizeText(row.dharmaName),
    phoneLast4: normalizePhoneLast4(row.phoneLast4),
    lastStageCompleted: '',
    lastUpdated: ''
  };

  STAGE_DEFINITIONS.forEach((stage) => {
    record[`${stage.id}Status`] = '';
    if (stage.scoreKey) {
      record[`${stage.id}Score`] = 0;
    }
    record[`${stage.id}ReviewCount`] = 0;
  });

  return record;
}

function applyProgressRowToSummary(summary, row) {
  const stageCompleted = normalizeText(row.stageCompleted);
  const eventType = normalizeText(row.eventType);
  const stageId = normalizeText(row.stageId);
  const monopolyScore = Number(row.monopolyScore || 0);

  STAGE_DEFINITIONS.forEach((stage) => {
    const completedByEvent = isStageCompletedByEvent(stage, eventType, stageId);
    const completedByText = stageCompleted === `${stage.label}完成`;
    const reviewCount = extractReviewCount(stageCompleted, stage.label);
    const reviewedByEvent = isStageReviewedByEvent(stage, eventType, stageId);

    if (completedByEvent || completedByText) {
      summary[`${stage.id}Status`] = stage.passValue;
    }

    if (reviewedByEvent || reviewCount > 0) {
      summary[`${stage.id}Status`] = stage.passValue;
      summary[`${stage.id}ReviewCount`] = Math.max(Number(summary[`${stage.id}ReviewCount`] || 0), reviewCount || 1);
    }

    if (stage.scoreKey && (completedByEvent || completedByText || reviewedByEvent || reviewCount > 0)) {
      summary[`${stage.id}Score`] = Math.max(Number(summary[`${stage.id}Score`] || 0), monopolyScore || 75);
    }
  });

  const isProfileSavedEvent = stageCompleted === '資料填寫完成';

  if (stageCompleted && !isProfileSavedEvent) {
    summary.lastStageCompleted = stageCompleted;
  } else if (!summary.lastStageCompleted && stageCompleted) {
    // 保留首次資料建立紀錄，但不覆蓋真正關卡完成結果。
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

function buildProgressFromSummary(summaryRecord) {
  const progress = {};

  STAGE_DEFINITIONS.forEach((stage) => {
    if (summaryRecord[`${stage.id}Status`]) {
      progress[stage.id] = summaryRecord[`${stage.id}Status`];
    }

    const reviewCount = Number(summaryRecord[`${stage.id}ReviewCount`] || 0);
    if (reviewCount > 0) {
      progress[`${stage.id}ReviewCount`] = reviewCount;
    }

    if (stage.scoreKey) {
      const score = Number(summaryRecord[`${stage.id}Score`] || 0);
      if (score > 0) {
        progress[stage.scoreKey] = score;
      }
    }
  });

  return progress;
}

function mapSummaryIndexes(headers) {
  const indexes = {
    playerName: headers.indexOf('姓名'),
    dharmaName: headers.indexOf('法名(沒有不必填寫)'),
    phoneLast4: headers.indexOf('手機末四碼'),
    lastStageCompleted: headers.indexOf('最後完成進度'),
    lastUpdated: headers.indexOf('最後更新時間')
  };

  STAGE_DEFINITIONS.forEach((stage) => {
    indexes[`${stage.id}Status`] = headers.indexOf(`${stage.label}狀態`);
    indexes[`${stage.id}ReviewCount`] = headers.indexOf(`${stage.label}複習次數`);
    if (stage.scoreKey) {
      indexes[`${stage.id}Score`] = headers.indexOf(`${stage.label}分數`);
    }
  });

  return indexes;
}

function summaryRowToRecord(row, indexes) {
  const record = {
    playerName: row[indexes.playerName],
    dharmaName: row[indexes.dharmaName],
    phoneLast4: row[indexes.phoneLast4],
    lastStageCompleted: row[indexes.lastStageCompleted],
    lastUpdated: row[indexes.lastUpdated]
  };

  STAGE_DEFINITIONS.forEach((stage) => {
    record[`${stage.id}Status`] = indexes[`${stage.id}Status`] >= 0 ? row[indexes[`${stage.id}Status`]] : '';
    record[`${stage.id}ReviewCount`] = indexes[`${stage.id}ReviewCount`] >= 0 ? row[indexes[`${stage.id}ReviewCount`]] : 0;
    if (stage.scoreKey) {
      record[`${stage.id}Score`] = indexes[`${stage.id}Score`] >= 0 ? row[indexes[`${stage.id}Score`]] : 0;
    }
  });

  return record;
}

function summaryRecordToRow(record) {
  return [
    record.playerName,
    record.dharmaName,
    record.phoneLast4,
    ...STAGE_DEFINITIONS.flatMap((stage) => {
      const values = [record[`${stage.id}Status`] || ''];
      if (stage.scoreKey) {
        values.push(Number(record[`${stage.id}Score`] || 0));
      }
      values.push(Number(record[`${stage.id}ReviewCount`] || 0));
      return values;
    }),
    record.lastStageCompleted,
    record.lastUpdated
  ];
}

function isStageCompletedByEvent(stage, eventType, stageId) {
  if (stage.id === 'DressCode') {
    return eventType === 'monopoly_pass' && stageId === stage.id;
  }
  return eventType === 'stage_pass' && stageId === stage.id;
}

function isStageReviewedByEvent(stage, eventType, stageId) {
  if (stage.id === 'DressCode') {
    return eventType === 'monopoly_review' && stageId === stage.id;
  }
  return eventType === 'stage_review' && stageId === stage.id;
}

function applySummaryColumnFormats(summarySheet) {
  if (summarySheet.getMaxColumns() < SUMMARY_HEADERS.length) {
    summarySheet.insertColumnsAfter(summarySheet.getMaxColumns(), SUMMARY_HEADERS.length - summarySheet.getMaxColumns());
  }

  const headers = summarySheet.getRange(1, 1, 1, SUMMARY_HEADERS.length).getValues()[0];
  const lastRow = Math.max(summarySheet.getLastRow(), 2);

  const setFormatByHeader = (headerName, format) => {
    const colIndex = headers.indexOf(headerName);
    if (colIndex >= 0) {
      summarySheet.getRange(2, colIndex + 1, lastRow - 1, 1).setNumberFormat(format);
    }
  };

  setFormatByHeader('姓名', '@');
  setFormatByHeader('法名(沒有不必填寫)', '@');
  setFormatByHeader('手機末四碼', '@');
  setFormatByHeader('最後完成進度', '@');
  setFormatByHeader('最後更新時間', 'yyyy/m/d h:mm:ss');

  STAGE_DEFINITIONS.forEach((stage) => {
    setFormatByHeader(`${stage.label}狀態`, '@');
    setFormatByHeader(`${stage.label}複習次數`, '0');
    if (stage.scoreKey) {
      setFormatByHeader(`${stage.label}分數`, '0');
    }
  });
}