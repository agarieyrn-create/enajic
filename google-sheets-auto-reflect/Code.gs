const SHEETS = {
  input: '入荷入力',
  rules: '振り分けルール',
  inventory: '3施設別在庫',
  report: '3施設別報告表',
  previous: '前日在庫',
  settings: '設定',
  parts: '部位マスタ',
};

const DEFAULT_SOURCE_SPREADSHEET_ID = '1K-cR1Cw4ycldHk6wl4MLwjnL0UejdTYJXIpycNaE8ck';

const FACILITIES = [
  'カンゲンファームブロック肉',
  '今帰仁冷凍施設',
  'アロマ加工場',
];

const SOURCE_CONFIGS = [
  {
    settingKey: '今帰仁元シート名',
    defaultSheetName: '2026.3～今帰仁冷凍施設(ブロック肉)',
    facility: '今帰仁冷凍施設',
    dateRow: 90,
    headerRow: 91,
    dataStartRow: 92,
    partCol: 5,
    firstDateCol: 8,
  },
  {
    settingKey: 'アロマ元シート名',
    defaultSheetName: '2026.3～アロマ加工場（ブロック肉）',
    facility: 'アロマ加工場',
    dateRow: 89,
    headerRow: 90,
    dataStartRow: 91,
    partCol: 3,
    firstDateCol: 6,
  },
];

const PARTS = [
  'ヒレ',
  'リブロース',
  'サーロイン',
  '肩ロース',
  'トウガラシ',
  'ウデ（しゃくし）',
  '前スネ',
  'ネック',
  'ブリスケ',
  '三角',
  '内バラ',
  'カイノミ',
  '外バラ',
  '内モモ',
  'シンタマ',
  '外モモ',
  'ランプ',
  '小肉',
  'スジ肉',
  'イチボ',
  'フランク',
  'トモズネ',
  'ランイチ',
  'ホルモンミックス',
  'ミスジ',
  'くず肉',
  'ハツ',
  'ハツモト',
  'シマ腸',
  'センマイ',
  '赤センマイ',
  '小腸',
  'アキレス',
  '直腸',
  'ミノ',
  'フク',
  'メンブルン',
  'ハチノス',
  '食道',
  '丸腸',
  '肩バラ',
  'ミンチ',
  'レバー',
  'ホホ',
  'テール',
  'タン',
  '横隔膜',
];

const INPUT_HEADERS = [
  '日付',
  '個体識別番号',
  '部位名',
  '左右',
  '重量kg',
  '摘要',
  '反映先施設',
  '反映状況',
  'エラー',
];

const RULE_HEADERS = ['部位名', '反映先施設', '有効', '備考'];
const INVENTORY_HEADERS = ['反映日', '施設', '部位名', '前日', 'IN', 'OUT', '残り', '摘要'];
const PREVIOUS_HEADERS = ['施設', '部位名', '前日重量kg'];

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('3施設 自動反映')
    .addItem('初期セットアップ', 'setupWorkbook')
    .addItem('自動反映（日付シート作成）', 'autoReflect')
    .addItem('日付別報告作成', 'createDailyReport')
    .addItem('PDF出力', 'exportReportPdf')
    .addToUi();
}

function setupWorkbook() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupSettingsSheet_(ss);
  setupPartsSheet_(ss);
  setupInventorySheet_(ss);
  setupReportSheet_(ss);
  autoResizeAll_(ss);
  SpreadsheetApp.getActive().toast('初期セットアップが完了しました。', '3施設 自動反映', 5);
}

function autoReflect() {
  createDailyReport();
}

function createDailyReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureRequiredSheets_(ss);
  const targetDateKey = getTargetDateKey_(ss);
  const rows = readSourceInventory_(ss, targetDateKey);
  writeInventory_(ss, rows);
  const sheetName = getDailyReportSheetName_(targetDateKey);
  writeReport_(ss, rows, targetDateKey, sheetName);
  ss.setActiveSheet(ss.getSheetByName(sheetName));
  SpreadsheetApp.getActive().toast(`${sheetName} を作成/更新しました。`, '3施設 自動反映', 5);
}

function exportReportPdf() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dailySheetName = getDailyReportSheetName_(getTargetDateKey_(ss));
  const sheet = ss.getSheetByName(dailySheetName) || ss.getSheetByName(SHEETS.report);
  if (!sheet) {
    throw new Error(`シート「${SHEETS.report}」がありません。先に初期セットアップを実行してください。`);
  }

  const folderId = getSettingValue_(ss, 'PDF保存先フォルダID');
  const folder = folderId ? DriveApp.getFolderById(folderId) : DriveApp.getRootFolder();
  const dateKey = getTargetDateKey_(ss);
  const fileName = `${sheet.getName()}_3施設別報告表.pdf`;
  const url = buildPdfExportUrl_(ss.getId(), sheet.getSheetId());
  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, {
    headers: { Authorization: `Bearer ${token}` },
  });
  folder.createFile(response.getBlob().setName(fileName));
  SpreadsheetApp.getActive().toast(`PDFを出力しました: ${fileName}`, '3施設 自動反映', 5);
}

function setupSettingsSheet_(ss) {
  const sheet = getOrCreateSheet_(ss, SHEETS.settings, 20, 2);
  resizeSheet_(sheet, 20, 2);
  sheet.clear();
  sheet.getRange(1, 1, 1, 2).setValues([['項目', '値']]).setFontWeight('bold');
  sheet.getRange(2, 1, 6, 2).setValues([
    ['報告対象日', new Date()],
    ['PDF保存先フォルダID', ''],
    ['元データSpreadsheetID', DEFAULT_SOURCE_SPREADSHEET_ID],
    ['今帰仁元シート名', '2026.3～今帰仁冷凍施設(ブロック肉)'],
    ['アロマ元シート名', '2026.3～アロマ加工場（ブロック肉）'],
    ['施設名', FACILITIES.join(', ')],
  ]);
  sheet.getRange('B2').setNumberFormat('yyyy/mm/dd');
  sheet.setFrozenRows(1);
}

function setupPartsSheet_(ss) {
  const sheet = getOrCreateSheet_(ss, SHEETS.parts, PARTS.length + 1, 1);
  resizeSheet_(sheet, PARTS.length + 1, 1);
  sheet.clear();
  sheet.getRange(1, 1).setValue('部位名').setFontWeight('bold');
  sheet.getRange(2, 1, PARTS.length, 1).setValues(PARTS.map((part) => [part]));
  sheet.hideSheet();
}

function setupRulesSheet_(ss) {
  const sheet = getOrCreateSheet_(ss, SHEETS.rules, PARTS.length + 1, RULE_HEADERS.length);
  resizeSheet_(sheet, PARTS.length + 1, RULE_HEADERS.length);
  sheet.clear();
  sheet.getRange(1, 1, 1, RULE_HEADERS.length).setValues([RULE_HEADERS]).setFontWeight('bold');
  const defaults = PARTS.map((part) => [part, FACILITIES[1], true, '初期値。必要に応じて施設を変更してください。']);
  sheet.getRange(2, 1, defaults.length, RULE_HEADERS.length).setValues(defaults);
  applyFacilityValidation_(sheet.getRange(2, 2, Math.max(PARTS.length, 200), 1));
  sheet.getRange(2, 3, Math.max(PARTS.length, 200), 1).insertCheckboxes();
  sheet.setFrozenRows(1);
}

function setupInputSheet_(ss) {
  const sheet = getOrCreateSheet_(ss, SHEETS.input, 250, INPUT_HEADERS.length);
  resizeSheet_(sheet, 250, INPUT_HEADERS.length);
  sheet.clear();
  sheet.getRange(1, 1, 1, INPUT_HEADERS.length).setValues([INPUT_HEADERS]).setFontWeight('bold');
  sheet.getRange(2, 1, 249, 1).setNumberFormat('yyyy/mm/dd');
  sheet.getRange(2, 5, 249, 1).setNumberFormat('0.00');
  applyPartValidation_(ss, sheet.getRange(2, 3, 249, 1));
  applySideValidation_(sheet.getRange(2, 4, 249, 1));
  sheet.getRange(2, 7, 249, 3).setBackground('#f3f4f6');
  sheet.setFrozenRows(1);
}

function setupInventorySheet_(ss) {
  const sheet = getOrCreateSheet_(ss, SHEETS.inventory, 250, INVENTORY_HEADERS.length);
  resizeSheet_(sheet, 250, INVENTORY_HEADERS.length);
  sheet.clear();
  sheet.getRange(1, 1, 1, INVENTORY_HEADERS.length).setValues([INVENTORY_HEADERS]).setFontWeight('bold');
  sheet.getRange(2, 1, 249, 1).setNumberFormat('yyyy/mm/dd');
  sheet.getRange(2, 4, 249, 4).setNumberFormat('0.00');
  sheet.setFrozenRows(1);
}

function setupPreviousSheet_(ss) {
  const sheet = getOrCreateSheet_(ss, SHEETS.previous, PARTS.length * FACILITIES.length + 1, PREVIOUS_HEADERS.length);
  resizeSheet_(sheet, PARTS.length * FACILITIES.length + 1, PREVIOUS_HEADERS.length);
  sheet.clear();
  sheet.getRange(1, 1, 1, PREVIOUS_HEADERS.length).setValues([PREVIOUS_HEADERS]).setFontWeight('bold');
  const rows = [];
  FACILITIES.forEach((facility) => {
    PARTS.forEach((part) => rows.push([facility, part, 0]));
  });
  sheet.getRange(2, 1, rows.length, PREVIOUS_HEADERS.length).setValues(rows);
  applyFacilityValidation_(sheet.getRange(2, 1, rows.length, 1));
  applyPartValidation_(ss, sheet.getRange(2, 2, rows.length, 1));
  sheet.getRange(2, 3, rows.length, 1).setNumberFormat('0.00');
  sheet.setFrozenRows(1);
}

function setupReportSheet_(ss) {
  writeReport_(ss, [], getTargetDateKey_(ss), SHEETS.report);
}

function ensureRequiredSheets_(ss) {
  [SHEETS.inventory, SHEETS.report, SHEETS.settings].forEach((name) => {
    if (!ss.getSheetByName(name)) {
      throw new Error(`シート「${name}」がありません。先に「初期セットアップ」を実行してください。`);
    }
  });
}

function getRoutingRules_(ss) {
  const sheet = ss.getSheetByName(SHEETS.rules);
  const values = sheet.getDataRange().getValues();
  const rules = new Map();
  for (let i = 1; i < values.length; i += 1) {
    const part = normalizeText_(values[i][0]);
    const facility = normalizeText_(values[i][1]);
    const enabled = values[i][2] === true || normalizeText_(values[i][2]) !== 'FALSE';
    if (!part || !enabled) continue;
    if (!FACILITIES.includes(facility)) continue;
    rules.set(part, facility);
  }
  return rules;
}

function getInputRows_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, INPUT_HEADERS.length).getValues();
  return values
    .map((row, index) => {
      const dateValue = row[0];
      const part = normalizeText_(row[2]);
      return {
        rowNo: index + 2,
        dateValue,
        dateKey: toDateKey_(dateValue),
        cattleId: normalizeText_(row[1]),
        part,
        side: normalizeText_(row[3]),
        weight: toNumber_(row[4]),
        note: normalizeText_(row[5]),
      };
    })
    .filter((row) => row.dateValue || row.cattleId || row.part || row.side || row.weight || row.note);
}

function validateInputRow_(row, rules) {
  const errors = [];
  if (!row.dateKey) errors.push('日付が未入力または不正です');
  if (!row.cattleId) errors.push('個体識別番号が未入力です');
  if (!row.part) errors.push('部位名が未入力です');
  if (row.part && !rules.has(row.part)) errors.push('振り分けルールがありません');
  if (!['右', '左', '左右なし'].includes(row.side)) errors.push('左右は「右」「左」「左右なし」から選択してください');
  if (!(row.weight > 0)) errors.push('重量kgは0より大きい数値にしてください');
  return errors;
}

function writeInputUpdates_(sheet, updates) {
  if (updates.length === 0) return;
  updates.forEach((update) => {
    sheet.getRange(update.rowNo, 7, 1, 3).setValues([[update.facility, update.status, update.error]]);
  });
}

function writeInventory_(ss, rows) {
  const sheet = ss.getSheetByName(SHEETS.inventory);
  sheet.getRange(1, 1, 1, INVENTORY_HEADERS.length).setValues([INVENTORY_HEADERS]).setFontWeight('bold');
  sheet.getRange(2, 1, Math.max(sheet.getMaxRows() - 1, 1), INVENTORY_HEADERS.length).clearContent();
  const sorted = rows.sort(compareAggregateRows_);
  if (sorted.length === 0) return;
  const values = sorted.map((row) => [
    row.dateValue,
    row.facility,
    row.part,
    round2_(row.previous),
    round2_(row.inWeight),
    round2_(row.outWeight),
    round2_(row.remaining),
    row.notes ? row.notes.join(' / ') : '',
  ]);
  sheet.getRange(2, 1, values.length, INVENTORY_HEADERS.length).setValues(values);
  sheet.getRange(2, 1, values.length, 1).setNumberFormat('yyyy/mm/dd');
  sheet.getRange(2, 4, values.length, 4).setNumberFormat('0.00');
}

function readInventory_(ss) {
  const sheet = ss.getSheetByName(SHEETS.inventory);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, INVENTORY_HEADERS.length).getValues()
    .filter((row) => row[0] && row[1] && row[2])
    .map((row) => ({
      dateValue: row[0],
      dateKey: toDateKey_(row[0]),
      facility: normalizeText_(row[1]),
      part: normalizeText_(row[2]),
      count: toNumber_(row[3]),
      weight: toNumber_(row[4]),
      notes: normalizeText_(row[5]) ? [normalizeText_(row[5])] : [],
    }));
}

function readSourceInventory_(ss, targetDateKey) {
  const sourceSpreadsheetId = normalizeText_(getSettingValue_(ss, '元データSpreadsheetID')) || DEFAULT_SOURCE_SPREADSHEET_ID;
  const sourceSpreadsheet = SpreadsheetApp.openById(sourceSpreadsheetId);
  const previousDateKey = addDaysToDateKey_(targetDateKey, -1);
  const aggregate = new Map();

  SOURCE_CONFIGS.forEach((config) => {
    const sheetName = normalizeText_(getSettingValue_(ss, config.settingKey)) || config.defaultSheetName;
    const sourceSheet = sourceSpreadsheet.getSheetByName(sheetName);
    if (!sourceSheet) {
      throw new Error(`元データシート「${sheetName}」が見つかりません。設定を確認してください。`);
    }

    const sourceRows = readSourceSheet_(sourceSheet, config, targetDateKey, previousDateKey);
    sourceRows.forEach((row) => {
      const key = makeKey_(targetDateKey, config.facility, row.part);
      if (!aggregate.has(key)) {
        aggregate.set(key, {
          dateKey: targetDateKey,
          dateValue: targetDateKey,
          facility: config.facility,
          part: row.part,
          previous: 0,
          inWeight: 0,
          outWeight: 0,
          remaining: 0,
          notes: [`元: ${sheetName}`],
        });
      }
      const item = aggregate.get(key);
      item.previous += row.previous;
      item.inWeight += row.inWeight;
      item.outWeight += row.outWeight;
      item.remaining += row.remaining;
    });
  });

  return Array.from(aggregate.values());
}

function readSourceSheet_(sourceSheet, config, targetDateKey, previousDateKey) {
  const lastColumn = sourceSheet.getLastColumn();
  const lastRow = sourceSheet.getLastRow();
  const dateRowValues = sourceSheet.getRange(config.dateRow, 1, 1, lastColumn).getValues()[0];
  const headerRowValues = sourceSheet.getRange(config.headerRow, 1, 1, lastColumn).getValues()[0];
  const targetCol = findDateColumn_(dateRowValues, headerRowValues, targetDateKey, config);
  const previousCol = findDateColumn_(dateRowValues, headerRowValues, previousDateKey, config);
  const requiredLastCol = Math.max(config.partCol, targetCol + 2, previousCol + 2);
  const values = sourceSheet
    .getRange(config.dataStartRow, 1, lastRow - config.dataStartRow + 1, requiredLastCol)
    .getValues();
  const rows = [];

  values.forEach((row) => {
    const part = normalizePartName_(row[config.partCol - 1]);
    if (!part || !PARTS.includes(part)) return;
    rows.push({
      part,
      previous: toNumber_(row[previousCol + 1]),
      inWeight: toNumber_(row[targetCol - 1]),
      outWeight: toNumber_(row[targetCol]),
      remaining: toNumber_(row[targetCol + 1]),
    });
  });

  return rows;
}

function findDateColumn_(dateRowValues, headerRowValues, dateKey, config) {
  for (let index = config.firstDateCol - 1; index < dateRowValues.length; index += 1) {
    if (toDateKey_(dateRowValues[index]) !== dateKey) continue;
    const inHeader = normalizeText_(headerRowValues[index]);
    const outHeader = normalizeText_(headerRowValues[index + 1]);
    const remainingHeader = normalizeText_(headerRowValues[index + 2]);
    if (inHeader === '入庫' && outHeader === '出庫' && remainingHeader === '残数') {
      return index + 1;
    }
  }
  throw new Error(`元データシート「${config.defaultSheetName}」に ${dateKey} の入庫/出庫/残数列が見つかりません。`);
}

function writeReport_(ss, rows, targetDateKey, sheetName) {
  const sheet = getOrCreateSheet_(ss, sheetName || SHEETS.report, PARTS.length + 4, 16);
  resizeSheet_(sheet, PARTS.length + 4, 16);
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).breakApart();
  sheet.clear();
  sheet.getRange('A1').setValue('報告対象日');
  sheet.getRange('B1').setValue(targetDateKey).setFontWeight('bold');

  const byFacilityPart = new Map();
  rows
    .filter((row) => row.dateKey === targetDateKey)
    .forEach((row) => {
      byFacilityPart.set(makeKey_('', row.facility, row.part), row);
    });

  sheet.getRange('D1:G1').merge().setValue(FACILITIES[0] + '在庫');
  sheet.getRange('H1:K1').merge().setValue(FACILITIES[1] + 'ブロック肉在庫');
  sheet.getRange('L1:O1').merge().setValue(FACILITIES[2] + 'ブロック肉在庫');
  sheet.getRange('C2:P2').setValues([[
    '部位名',
    '前日',
    'IN',
    'OUT',
    '残り',
    '前日',
    'IN',
    'OUT',
    '残り',
    '前日',
    'IN',
    'OUT',
    '残り',
    '摘要',
  ]]);
  sheet.getRange('D1:O2').setFontWeight('bold').setHorizontalAlignment('center').setBackground('#dbeafe');
  sheet.getRange('C2:P2').setBackground('#f3f4f6');

  const body = PARTS.map((part, index) => {
    const rowNumber = index + 3;
    const row = [part];
    const notes = [];
    FACILITIES.forEach((facility, facilityIndex) => {
      const data = byFacilityPart.get(makeKey_('', facility, part));
      const prev = data ? data.previous : 0;
      const inWeight = data ? data.inWeight : 0;
      const outWeight = data ? data.outWeight : 0;
      const remaining = data ? data.remaining : 0;
      if (data && data.notes) {
        data.notes.forEach((note) => {
          if (note && !notes.includes(note)) notes.push(note);
        });
      }
      row.push(round2_(prev), round2_(inWeight), round2_(outWeight), round2_(remaining));
    });
    row.push(notes.join(' / '));
    return row;
  });

  sheet.getRange(3, 3, body.length, 14).setValues(body);
  const totalRow = body.length + 3;
  sheet.getRange(totalRow, 3).setValue('合計').setFontWeight('bold');
  ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O'].forEach((column) => {
    sheet.getRange(`${column}${totalRow}`).setFormula(`=SUM(${column}3:${column}${totalRow - 1})`).setFontWeight('bold');
  });
  sheet.getRange(3, 4, body.length + 1, 12).setNumberFormat('0.00');
  sheet.getRange(1, 3, totalRow, 14).setBorder(true, true, true, true, true, true);
  sheet.setFrozenRows(2);
  sheet.setFrozenColumns(3);
  sheet.setColumnWidth(3, 150);
  sheet.setColumnWidths(4, 12, 70);
  sheet.setColumnWidth(16, 220);
}

function applyPartValidation_(ss, range) {
  const partSheet = ss.getSheetByName(SHEETS.parts);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(partSheet.getRange(2, 1, PARTS.length, 1), true)
    .setAllowInvalid(false)
    .build();
  range.setDataValidation(rule);
}

function applyFacilityValidation_(range) {
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(FACILITIES, true)
    .setAllowInvalid(false)
    .build();
  range.setDataValidation(rule);
}

function applySideValidation_(range) {
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['右', '左', '左右なし'], true)
    .setAllowInvalid(false)
    .build();
  range.setDataValidation(rule);
}

function getTargetDateKey_(ss) {
  const value = getSettingValue_(ss, '報告対象日');
  return toDateKey_(value) || toDateKey_(new Date());
}

function getDailyReportSheetName_(dateKey) {
  const parts = dateKey.split('-').map((value) => Number(value));
  if (parts.length !== 3 || parts.some((value) => !Number.isFinite(value))) {
    return `${dateKey}(3)`;
  }
  return `${parts[0]}.${parts[1]}.${parts[2]}(3)`;
}

function getSettingValue_(ss, key) {
  const sheet = ss.getSheetByName(SHEETS.settings);
  if (!sheet) return '';
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i += 1) {
    if (normalizeText_(values[i][0]) === key) return values[i][1];
  }
  return '';
}

function getPreviousStock_(ss) {
  const sheet = ss.getSheetByName(SHEETS.previous);
  const result = new Map();
  if (!sheet || sheet.getLastRow() < 2) return result;
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, PREVIOUS_HEADERS.length).getValues();
  values.forEach((row) => {
    const facility = normalizeText_(row[0]);
    const part = normalizeText_(row[1]);
    if (!facility || !part) return;
    result.set(makeKey_('', facility, part), toNumber_(row[2]));
  });
  return result;
}

function buildPdfExportUrl_(spreadsheetId, sheetId) {
  const params = [
    'format=pdf',
    'size=A4',
    'portrait=true',
    'fitw=true',
    'sheetnames=false',
    'printtitle=false',
    'pagenumbers=false',
    'gridlines=false',
    'fzr=false',
    `gid=${sheetId}`,
  ].join('&');
  return `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?${params}`;
}

function getOrCreateSheet_(ss, name, rowCount, columnCount) {
  const existing = ss.getSheetByName(name);
  if (existing) return existing;
  const rows = rowCount || 60;
  const columns = columnCount || 16;
  const template = findSmallestSheet_(ss);
  const created = template ? template.copyTo(ss).setName(name) : ss.insertSheet(name);
  resizeSheet_(created, rows, columns);
  return created;
}

function resizeSheet_(sheet, rowCount, columnCount) {
  const maxRows = sheet.getMaxRows();
  const maxColumns = sheet.getMaxColumns();
  if (maxRows > rowCount) {
    sheet.deleteRows(rowCount + 1, maxRows - rowCount);
  } else if (maxRows < rowCount) {
    sheet.insertRowsAfter(maxRows, rowCount - maxRows);
  }
  if (maxColumns > columnCount) {
    sheet.deleteColumns(columnCount + 1, maxColumns - columnCount);
  } else if (maxColumns < columnCount) {
    sheet.insertColumnsAfter(maxColumns, columnCount - maxColumns);
  }
}

function findSmallestSheet_(ss) {
  const sheets = ss.getSheets().filter((sheet) => sheet.getSheetType() === SpreadsheetApp.SheetType.GRID);
  if (sheets.length === 0) return null;
  return sheets
    .map((sheet) => ({
      sheet,
      cells: sheet.getMaxRows() * sheet.getMaxColumns(),
    }))
    .sort((a, b) => a.cells - b.cells)[0].sheet;
}

function autoResizeAll_(ss) {
  Object.keys(SHEETS).forEach((key) => {
    const sheet = ss.getSheetByName(SHEETS[key]);
    if (sheet && !sheet.isSheetHidden()) {
      sheet.autoResizeColumns(1, Math.max(sheet.getLastColumn(), 1));
    }
  });
}

function normalizeText_(value) {
  return value === null || value === undefined ? '' : String(value).trim();
}

function normalizePartName_(value) {
  const part = normalizeText_(value).replace(/\s+/g, '');
  const aliases = {
    'ミンチ1kg': 'ミンチ',
    'ミンチ1ｋｇ': 'ミンチ',
    'ウデ': 'ウデ（しゃくし）',
  };
  return aliases[part] || part;
}

function toNumber_(value) {
  if (value === null || value === undefined || value === '') return 0;
  const number = Number(value);
  return Number.isFinite(number) ? number : 0;
}

function toDateKey_(value) {
  if (!value) return '';
  if (typeof value === 'number' && Number.isFinite(value)) {
    const serialDate = new Date(Math.round((value - 25569) * 86400 * 1000));
    return Utilities.formatDate(serialDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  let date = value;
  if (!(date instanceof Date)) {
    date = new Date(value);
  }
  if (Number.isNaN(date.getTime())) return '';
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function addDaysToDateKey_(dateKey, days) {
  const date = new Date(`${dateKey}T00:00:00`);
  if (Number.isNaN(date.getTime())) {
    throw new Error(`日付が不正です: ${dateKey}`);
  }
  date.setDate(date.getDate() + days);
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function makeKey_(dateKey, facility, part) {
  return [dateKey, facility, part].join('\u0001');
}

function compareAggregateRows_(a, b) {
  return (
    a.dateKey.localeCompare(b.dateKey) ||
    FACILITIES.indexOf(a.facility) - FACILITIES.indexOf(b.facility) ||
    a.part.localeCompare(b.part, 'ja')
  );
}

function round2_(value) {
  return Math.round((value + Number.EPSILON) * 100) / 100;
}
