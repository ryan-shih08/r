/**
 * 旅行共編網頁：Google Apps Script 後端
 * 資料存在綁定的 Google 試算表（行程 / 分攤 / 候選 / 清單 / 設定 / 匯率 等工作表）
 */

var SHEETS = {
  itinerary: {
    name: '行程',
    headers: ['id', 'date', 'time', 'endTime', 'title', 'category', 'location', 'note', 'status', 'author', 'updatedAt']
  },
  expenses: {
    name: '分攤',
    headers: ['id', 'date', 'item', 'amount', 'payer', 'splitAmong', 'note', 'author', 'updatedAt', 'currency', 'rate']
  },
  candidates: {
    name: '候選',
    headers: ['id', 'title', 'category', 'location', 'link', 'note', 'votes', 'author', 'createdAt', 'updatedAt']
  },
  checklist: {
    name: '清單',
    headers: ['id', 'kind', 'title', 'scope', 'owner', 'dueDate', 'note', 'done', 'checkedBy', 'author', 'updatedAt']
  },
  settings: {
    name: '設定',
    headers: ['key', 'value']
  }
};

var ITEM_TYPES = ['itinerary', 'expenses', 'candidates', 'checklist'];

function checkType_(type) {
  if (ITEM_TYPES.indexOf(type) < 0) throw new Error('未知的類型');
}

var DEFAULT_SETTINGS = {
  tripName: '我們的旅行',
  startDate: '',
  endDate: '',
  members: '成員A,成員B,成員C,成員D',
  currency: 'TWD',
  extraCurrencies: 'THB'
};

var RATE_SHEET = '匯率';

function doGet() {
  ensureSheets_();
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(getSettings_().tripName || '旅行共編')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

/** 第一次使用可在編輯器手動執行，建立工作表 */
function setup() {
  ensureSheets_();
  var fx = getRates_(getSettings_());
  Logger.log('目前匯率（換算成 ' + fx.base + '）：' + JSON.stringify(fx.rates));
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
    } else if (sheet.getLastColumn() < def.headers.length) {
      // 舊版工作表：補上新增的欄位
      var from = Math.max(sheet.getLastColumn(), 1);
      sheet.getRange(1, from, sheet.getMaxRows(), def.headers.length - from + 1).setNumberFormat('@');
      sheet.getRange(1, 1, 1, def.headers.length).setValues([def.headers]).setFontWeight('bold');
    }
  });
}

/* ---------- 匯率 ---------- */

function currencyList_(text) {
  return String(text || '').toUpperCase().split(/[^A-Z]+/)
    .filter(function (c) { return /^[A-Z]{3}$/.test(c); })
    .filter(function (c, i, arr) { return arr.indexOf(c) === i; });
}

/**
 * 取得「1 外幣 = ? 基準幣」的匯率。
 * 先用試算表內建的 GOOGLEFINANCE，抓不到再用免費的 open.er-api.com 備援。
 */
function getRates_(settings) {
  var base = currencyList_(settings.currency)[0] || 'TWD';
  var list = currencyList_(settings.extraCurrencies).filter(function (c) { return c !== base; });
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(RATE_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(RATE_SHEET);
    sheet.getRange(1, 1, 1, 3).setValues([['幣別', '匯率（1 單位 = ? 基準幣）', '基準幣']]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  var rates = {};
  var changed = false;
  list.forEach(function (cur, i) {
    var row = i + 2;
    var formula = '=GOOGLEFINANCE("CURRENCY:' + cur + base + '")';
    if (sheet.getRange(row, 1).getValue() !== cur || sheet.getRange(row, 2).getFormula() !== formula) {
      sheet.getRange(row, 1).setValue(cur);
      sheet.getRange(row, 2).setFormula(formula);
      sheet.getRange(row, 3).setValue(base);
      changed = true;
    }
  });
  if (changed) SpreadsheetApp.flush();
  var lastRow = sheet.getLastRow();
  if (lastRow > list.length + 1) sheet.getRange(list.length + 2, 1, lastRow - list.length - 1, 3).clearContent();

  list.forEach(function (cur, i) {
    var v = sheet.getRange(i + 2, 2).getValue();
    if (typeof v === 'number' && v > 0) {
      rates[cur] = { rate: v, source: 'Google 財經' };
    } else {
      var backup = fetchRate_(cur, base);
      rates[cur] = backup ? { rate: backup, source: 'open.er-api.com' } : { rate: 0, source: '' };
    }
  });
  return { base: base, rates: rates };
}

function fetchRate_(cur, base) {
  var cache = CacheService.getScriptCache();
  var key = 'rate_' + cur + base;
  var hit = cache.get(key);
  if (hit) return Number(hit);
  try {
    var res = UrlFetchApp.fetch('https://open.er-api.com/v6/latest/' + cur, { muteHttpExceptions: true });
    var json = JSON.parse(res.getContentText());
    var rate = json && json.rates && Number(json.rates[base]);
    if (rate > 0) {
      cache.put(key, String(rate), 3 * 60 * 60);
      return rate;
    }
  } catch (e) {}
  return 0;
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
  var settings = getSettings_();
  var fx = { base: settings.currency || 'TWD', rates: {} };
  try { fx = getRates_(settings); } catch (e) {}
  return {
    settings: settings,
    fx: fx,
    itinerary: readRows_('itinerary'),
    expenses: readRows_('expenses'),
    candidates: readRows_('candidates'),
    checklist: readRows_('checklist'),
    serverTime: new Date().toISOString()
  };
}

/** 新增或更新一筆（type: itinerary | expenses | candidates） */
function saveItem(type, item) {
  checkType_(type);
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    ensureSheets_();
    writeItem_(type, item);
  } finally {
    lock.releaseLock();
  }
  return getData();
}

function writeItem_(type, item) {
  var def = SHEETS[type];
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(def.name);
  item.id = item.id || Utilities.getUuid();
  item.updatedAt = new Date().toISOString();
  var rowIndex = findRow_(sheet, item.id);
  if (type === 'candidates') {
    // 投票只能透過 vote() 修改，編輯內容時保留原本的票與建立時間
    if (rowIndex) {
      var old = sheet.getRange(rowIndex, 1, 1, def.headers.length).getDisplayValues()[0];
      item.votes = old[def.headers.indexOf('votes')];
      item.createdAt = old[def.headers.indexOf('createdAt')];
    } else {
      item.votes = '{}';
      item.createdAt = item.updatedAt;
    }
  }
  if (type === 'checklist') {
    // 勾選狀態只能透過 toggleCheck() 修改，編輯內容時保留
    if (rowIndex) {
      var prev = sheet.getRange(rowIndex, 1, 1, def.headers.length).getDisplayValues()[0];
      item.done = prev[def.headers.indexOf('done')];
      item.checkedBy = prev[def.headers.indexOf('checkedBy')];
    } else {
      item.done = '';
      item.checkedBy = '[]';
    }
  }
  var row = def.headers.map(function (h) {
    var v = item[h] == null ? '' : String(item[h]);
    return v.slice(0, 2000);
  });
  if (rowIndex) {
    sheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
  } else {
    var next = sheet.getLastRow() + 1;
    sheet.getRange(next, 1, 1, row.length).setNumberFormat('@').setValues([row]);
  }
}

function deleteItem(type, id) {
  checkType_(type);
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

/** 候選投票：value 為 1（讚）、-1（倒讚）或 0（取消） */
function vote(id, name, value) {
  name = String(name || '').trim().slice(0, 50);
  if (!name) throw new Error('請先選擇你的名字');
  value = Number(value);
  if ([1, -1, 0].indexOf(value) < 0) throw new Error('投票值錯誤');
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    ensureSheets_();
    var def = SHEETS.candidates;
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(def.name);
    var rowIndex = findRow_(sheet, id);
    if (!rowIndex) throw new Error('這個候選項目已被刪除或排入行程');
    var col = def.headers.indexOf('votes') + 1;
    var votes = {};
    try { votes = JSON.parse(sheet.getRange(rowIndex, col).getDisplayValue() || '{}') || {}; } catch (e) {}
    if (value === 0) delete votes[name]; else votes[name] = value;
    sheet.getRange(rowIndex, col).setValue(JSON.stringify(votes));
  } finally {
    lock.releaseLock();
  }
  return getData();
}

/** 把候選項目排入行程：新增一筆行程並移除該候選 */
function scheduleCandidate(candidateId, event) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    ensureSheets_();
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.candidates.name);
    var rowIndex = findRow_(sheet, candidateId);
    if (!rowIndex) throw new Error('這個候選項目已被刪除或排入行程');
    event.id = '';
    writeItem_('itinerary', event);
    sheet.deleteRow(rowIndex);
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

/** 一次新增多筆（例如匯入建議行李清單） */
function addItems(type, items) {
  checkType_(type);
  if (!items || !items.length) return getData();
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    ensureSheets_();
    items.slice(0, 100).forEach(function (item) {
      item.id = '';
      writeItem_(type, item);
    });
  } finally {
    lock.releaseLock();
  }
  return getData();
}

/** 清單勾選：「每人都帶」的行李記錄每個人各自的勾選，其餘項目只有一個完成狀態 */
function toggleCheck(id, name, checked) {
  name = String(name || '').trim().slice(0, 50);
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    ensureSheets_();
    var def = SHEETS.checklist;
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(def.name);
    var rowIndex = findRow_(sheet, id);
    if (!rowIndex) throw new Error('這個項目已被刪除');
    var row = sheet.getRange(rowIndex, 1, 1, def.headers.length).getDisplayValues()[0];
    var kind = row[def.headers.indexOf('kind')];
    var scope = row[def.headers.indexOf('scope')];
    if (kind === 'pack' && scope === 'personal') {
      if (!name) throw new Error('請先選擇你的名字');
      var col = def.headers.indexOf('checkedBy') + 1;
      var list = [];
      try { list = JSON.parse(row[col - 1] || '[]') || []; } catch (e) {}
      list = list.filter(function (n) { return n !== name; });
      if (checked) list.push(name);
      sheet.getRange(rowIndex, col).setValue(JSON.stringify(list));
    } else {
      sheet.getRange(rowIndex, def.headers.indexOf('done') + 1).setValue(checked ? '1' : '');
    }
  } finally {
    lock.releaseLock();
  }
  return getData();
}
