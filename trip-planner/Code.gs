/**
 * 旅行共編網頁：Google Apps Script 後端
 * 資料存在綁定的 Google 試算表（行程 / 分攤 / 設定 三個工作表）
 */

var SHEETS = {
  itinerary: {
    name: '行程',
    headers: ['id', 'date', 'time', 'endTime', 'title', 'category', 'location', 'note', 'status', 'author', 'updatedAt']
  },
  expenses: {
    name: '分攤',
    headers: ['id', 'date', 'item', 'amount', 'payer', 'splitAmong', 'note', 'author', 'updatedAt']
  },
  settings: {
    name: '設定',
    headers: ['key', 'value']
  }
};

var DEFAULT_SETTINGS = {
  tripName: '我們的旅行',
  startDate: '',
  endDate: '',
  members: '成員A,成員B,成員C,成員D',
  currency: 'TWD'
};

function doGet() {
  ensureSheets_();
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(getSettings_().tripName || '旅行共編')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

/** 第一次使用可在編輯器手動執行，建立工作表 */
function setup() {
  ensureSheets_();
}

function ensureSheets_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(SHEETS).forEach(function (key) {
    var def = SHEETS[key];
    var sheet = ss.getSheetByName(def.name);
    if (!sheet) {
      sheet = ss.insertSheet(def.name);
      // 全部存成純文字，避免日期、時間被試算表自動轉格式
      sheet.getRange(1, 1, sheet.getMaxRows(), def.headers.length).setNumberFormat('@');
      sheet.getRange(1, 1, 1, def.headers.length).setValues([def.headers]).setFontWeight('bold');
      sheet.setFrozenRows(1);
      if (key === 'settings') {
        var rows = Object.keys(DEFAULT_SETTINGS).map(function (k) { return [k, DEFAULT_SETTINGS[k]]; });
        sheet.getRange(2, 1, rows.length, 2).setValues(rows);
      }
    }
  });
}

function readRows_(key) {
  var def = SHEETS[key];
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(def.name);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var values = sheet.getRange(2, 1, lastRow - 1, def.headers.length).getDisplayValues();
  return values
    .filter(function (r) { return r[0]; })
    .map(function (r) {
      var obj = {};
      def.headers.forEach(function (h, i) { obj[h] = r[i]; });
      return obj;
    });
}

function getSettings_() {
  var settings = {};
  Object.keys(DEFAULT_SETTINGS).forEach(function (k) { settings[k] = DEFAULT_SETTINGS[k]; });
  readRows_('settings').forEach(function (r) { settings[r.key] = r.value; });
  return settings;
}

/** 前端讀取全部資料 */
function getData() {
  ensureSheets_();
  return {
    settings: getSettings_(),
    itinerary: readRows_('itinerary'),
    expenses: readRows_('expenses'),
    serverTime: new Date().toISOString()
  };
}

/** 新增或更新一筆（type: itinerary | expenses） */
function saveItem(type, item) {
  if (type !== 'itinerary' && type !== 'expenses') throw new Error('未知的類型');
  var def = SHEETS[type];
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    ensureSheets_();
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(def.name);
    item.id = item.id || Utilities.getUuid();
    item.updatedAt = new Date().toISOString();
    var row = def.headers.map(function (h) {
      var v = item[h] == null ? '' : String(item[h]);
      return v.slice(0, 2000);
    });
    var rowIndex = findRow_(sheet, item.id);
    if (rowIndex) {
      sheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
    } else {
      var next = sheet.getLastRow() + 1;
      sheet.getRange(next, 1, 1, row.length).setNumberFormat('@').setValues([row]);
    }
  } finally {
    lock.releaseLock();
  }
  return getData();
}

function deleteItem(type, id) {
  if (type !== 'itinerary' && type !== 'expenses') throw new Error('未知的類型');
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS[type].name);
    var rowIndex = findRow_(sheet, id);
    if (rowIndex) sheet.deleteRow(rowIndex);
  } finally {
    lock.releaseLock();
  }
  return getData();
}

function saveSettings(newSettings) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    ensureSheets_();
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.settings.name);
    Object.keys(DEFAULT_SETTINGS).forEach(function (k) {
      if (newSettings[k] == null) return;
      var value = String(newSettings[k]).slice(0, 500);
      var rowIndex = findRow_(sheet, k);
      if (rowIndex) {
        sheet.getRange(rowIndex, 2).setValue(value);
      } else {
        sheet.appendRow([k, value]);
      }
    });
  } finally {
    lock.releaseLock();
  }
  return getData();
}

function findRow_(sheet, id) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  var ids = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues();
  for (var i = 0; i < ids.length; i++) {
    if (ids[i][0] === id) return i + 2;
  }
  return 0;
}
