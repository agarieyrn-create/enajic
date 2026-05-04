const SHEETS = {
  input: '入荷入力',
  rules: '振り分けルール',
  inventory: '3施設別在庫',
  report: '3施設別報告表',
  previous: '前日在庫',
  settings: '設定',
  parts: '部位マスタ',
};

const FACILITIES = [
  'カンゲンファームブロック肉',
  '今帰仁冷凍施設',
  'アロマ加工場',
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
const INVENTORY_HEADERS = ['反映日', '施設', '部位名', '明細数', '合計重量kg', '摘要'];
const PREVIOUS_HEADERS = ['施設', '部位名', '前日重量kg'];

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('3施設 自動反映')
    .addItem('初期セットアップ', 'setupWorkbook')
    .addItem('自動反映', 'autoReflect')
    .addItem('日付別報告作成', 'createDailyReport')
    .addItem('PDF出力', 'exportReportPdf')
    .addToUi();
}

function setupWorkbook() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupSettingsSheet_(ss);
  setupPartsSheet_(ss);
  setupRulesSheet_(ss);
  setupInputSheet_(ss);
  setupPreviousSheet_(ss);
  setupInventorySheet_(ss);
  setupReportSheet_(ss);
  autoResizeAll_(ss);
  SpreadsheetApp.getActive().toast('初期セットアップが完了しました。', '3施設 自動反映', 5);
}

function autoReflect() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureRequiredSheets_(ss);

  const inputSheet = ss.getSheetByName(SHEETS.input);
  const rules = getRoutingRules_(ss);
  const rows = getInputRows_(inputSheet);
  const aggregate = new Map();
  const updates = [];

  rows.forEach((row) => {
    const rowNo = row.rowNo;
    const errors = validateInputRow_(row, rules);
    if (errors.length > 0) {
      updates.push({ rowNo, facility: '', status: '未反映', error: errors.join(' / ') });
      return;
    }

    const facility = rules.get(row.part);
    const key = makeKey_(row.dateKey, facility, row.part);
    if (!aggregate.has(key)) {
      aggregate.set(key, {
        dateKey: row.dateKey,
        dateValue: row.dateValue,
        facility,
        part: row.part,
        count: 0,
        weight: 0,
        notes: [],
      });
    }
    const item = aggregate.get(key);
    item.count += 1;
    item.weight += row.weight;
    if (row.note && !item.notes.includes(row.note)) item.notes.push(row.note);
    updates.push({ rowNo, facility, status: '反映済', error: '' });
  });

  writeInputUpdates_(inputSheet, updates);
  writeInventory_(ss, Array.from(aggregate.values()));
  writeReport_(ss, Array.from(aggregate.values()), getTargetDateKey_(ss));
  SpreadsheetApp.getActive().toast('3施設別在庫と報告表を更新しました。', '3施設 自動反映', 5);
}

function createDailyReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureRequiredSheets_(ss);
  const rows = readInventory_(ss);
  writeReport_(ss, rows, getTargetDateKey_(ss));
  SpreadsheetApp.getActive().toast('日付別報告表を作成しました。', '3施設 自動反映', 5);
}

function exportReportPdf() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.report);
  if (!sheet) {
    throw new Error(`シート「${SHEETS.report}」がありません。先に初期セットアップを実行してください。`);
  }

  const folderId = getSettingValue_(ss, 'PDF保存先フォルダID');
  const folder = folderId ? DriveApp.getFolderById(folderId) : DriveApp.getRootFolder();
  const dateKey = getTargetDateKey_(ss);
  const fileName = `3施設別報告表_${dateKey}.pdf`;
  const url = buildPdfExportUrl_(ss.getId(), sheet.getSheetId());
  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, {
    headers: { Authorization: `Bearer ${token}` },
  });
  folder.createFile(response.getBlob().setName(fileName));
  SpreadsheetApp.getActive().toast(`PDFを出力しました: ${fileName}`, '3施設 自動反映', 5);
}

function setupSettingsSheet_(ss) {
  const sheet = getOrCreateSheet_(ss, SHEETS.settings);
  sheet.clear();
  sheet.getRange(1, 1, 1, 2).setValues([['項目', '値']]).setFontWeight('bold');
  sheet.getRange(2, 1, 3, 2).setValues([
    ['報告対象日', new Date()],
    ['PDF保存先フォルダID', ''],
    ['施設名', FACILITIES.join(', ')],
  ]);
  sheet.getRange('B2').setNumberFormat('yyyy/mm/dd');
  sheet.setFrozenRows(1);
}

function setupPartsSheet_(ss) {
  const sheet = getOrCreateSheet_(ss, SHEETS.parts);
  sheet.clear();
  sheet.getRange(1, 1).setValue('部位名').setFontWeight('bold');
  sheet.getRange(2, 1, PARTS.length, 1).setValues(PARTS.map((part) => [part]));
  sheet.hideSheet();
}

function setupRulesSheet_(ss) {
  const sheet = getOrCreateSheet_(ss, SHEETS.rules);
  sheet.clear();
  sheet.getRange(1, 1, 1, RULE_HEADERS.length).setValues([RULE_HEADERS]).setFontWeight('bold');
  const defaults = PARTS.map((part) => [part, FACILITIES[1], true, '初期値。必要に応じて施設を変更してください。']);
  sheet.getRange(2, 1, defaults.length, RULE_HEADERS.length).setValues(defaults);
  applyFacilityValidation_(sheet.getRange(2, 2, Math.max(PARTS.length, 200), 1));
  sheet.getRange(2, 3, Math.max(PARTS.length, 200), 1).insertCheckboxes();
  sheet.setFrozenRows(1);
}

function setupInputSheet_(ss) {
  const sheet = getOrCreateSheet_(ss, SHEETS.input);
  sheet.clear();
  sheet.getRange(1, 1, 1, INPUT_HEADERS.length).setValues([INPUT_HEADERS]).setFontWeight('bold');
  sheet.getRange(2, 1, 500, 1).setNumberFormat('yyyy/mm/dd');
  sheet.getRange(2, 5, 500, 1).setNumberFormat('0.00');
  applyPartValidation_(ss, sheet.getRange(2, 3, 500, 1));
  applySideValidation_(sheet.getRange(2, 4, 500, 1));
  sheet.getRange(2, 7, 500, 3).setBackground('#f3f4f6');
  sheet.setFrozenRows(1);
}

function setupInventorySheet_(ss) {
  const sheet = getOrCreateSheet_(ss, SHEETS.inventory);
  sheet.clear();
  sheet.getRange(1, 1, 1, INVENTORY_HEADERS.length).setValues([INVENTORY_HEADERS]).setFontWeight('bold');
  sheet.getRange(2, 1, 1000, 1).setNumberFormat('yyyy/mm/dd');
  sheet.getRange(2, 5, 1000, 1).setNumberFormat('0.00');
  sheet.setFrozenRows(1);
}

function setupPreviousSheet_(ss) {
  const sheet = getOrCreateSheet_(ss, SHEETS.previous);
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
  const sheet = getOrCreateSheet_(ss, SHEETS.report);
  writeReport_(ss, [], getTargetDateKey_(ss));
}

function ensureRequiredSheets_(ss) {
  [SHEETS.input, SHEETS.rules, SHEETS.inventory, SHEETS.report, SHEETS.previous, SHEETS.settings].forEach((name) => {
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
  sheet.getRange(2, 1, Math.max(sheet.getMaxRows() - 1, 1), INVENTORY_HEADERS.length).clearContent();
  const sorted = rows.sort(compareAggregateRows_);
  if (sorted.length === 0) return;
  const values = sorted.map((row) => [
    row.dateValue,
    row.facility,
    row.part,
    row.count,
    round2_(row.weight),
    row.notes ? row.notes.join(' / ') : '',
  ]);
  sheet.getRange(2, 1, values.length, INVENTORY_HEADERS.length).setValues(values);
  sheet.getRange(2, 1, values.length, 1).setNumberFormat('yyyy/mm/dd');
  sheet.getRange(2, 5, values.length, 1).setNumberFormat('0.00');
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

function writeReport_(ss, rows, targetDateKey) {
  const sheet = ss.getSheetByName(SHEETS.report);
  const previous = getPreviousStock_(ss);
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
      const prev = previous.get(makeKey_('', facility, part)) || 0;
      const inWeight = data ? data.weight : 0;
      if (data && data.notes) {
        data.notes.forEach((note) => {
          if (note && !notes.includes(note)) notes.push(note);
        });
      }
      row.push(round2_(prev), round2_(inWeight), 0, '');
    });
    row.push(notes.join(' / '));
    return row;
  });

  sheet.getRange(3, 3, body.length, 14).setValues(body);
  const totalRow = body.length + 3;
  PARTS.forEach((_, index) => {
    const rowNumber = index + 3;
    sheet.getRange(`G${rowNumber}`).setFormula(`=D${rowNumber}+E${rowNumber}-F${rowNumber}`);
    sheet.getRange(`K${rowNumber}`).setFormula(`=H${rowNumber}+I${rowNumber}-J${rowNumber}`);
    sheet.getRange(`O${rowNumber}`).setFormula(`=L${rowNumber}+M${rowNumber}-N${rowNumber}`);
  });
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

function getOrCreateSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
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

function toNumber_(value) {
  if (value === null || value === undefined || value === '') return 0;
  const number = Number(value);
  return Number.isFinite(number) ? number : 0;
}

function toDateKey_(value) {
  if (!value) return '';
  let date = value;
  if (!(date instanceof Date)) {
    date = new Date(value);
  }
  if (Number.isNaN(date.getTime())) return '';
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
