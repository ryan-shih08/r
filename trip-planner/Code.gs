/**
 * 旅行共編網頁：Google Apps Script 後端
 * 資料存在綁定的 Google 試算表（行程 / 分攤 / 候選 / 清單 / 設定 / 紀錄 / 匯率 等工作表）
 */

var SHEETS = {
  itinerary: {
    name: '行程',
    headers: ['id', 'date', 'time', 'endTime', 'title', 'category', 'location', 'note', 'status', 'author', 'updatedAt', 'images', 'booking']
  },
  expenses: {
    name: '分攤',
    headers: ['id', 'date', 'item', 'amount', 'payer', 'splitAmong', 'note', 'author', 'updatedAt', 'currency', 'rate']
  },
  candidates: {
    name: '候選',
    headers: ['id', 'title', 'category', 'location', 'link', 'note', 'votes', 'author', 'createdAt', 'updatedAt', 'images', 'booking']
  },
  checklist: {
    name: '清單',
    headers: ['id', 'kind', 'title', 'scope', 'owner', 'dueDate', 'note', 'done', 'checkedBy', 'author', 'updatedAt']
  },
  settings: {
    name: '設定',
    headers: ['key', 'value']
  },
  comments: {
    name: '留言',
    headers: ['id', 'targetType', 'targetId', 'text', 'author', 'createdAt', 'updatedAt']
  },
  log: {
    name: '紀錄',
    headers: ['id', 'time', 'actor', 'action', 'summary', 'changes', 'undoneBy']
  }
};

var TYPE_LABELS = { itinerary: '行程', expenses: '支出', candidates: '候選', checklist: '清單', comments: '留言', settings: '設定' };
var LOG_KEEP = 300;

/** 程式版本：要跟 Index.html 裡的 APP_VERSION 一樣，不一樣代表其中一個檔案沒更新到 */
var APP_VERSION = '2026-10-03.1';

/**
 * 上傳圖片存放的 Google 雲端硬碟資料夾 ID
 * （打開資料夾後，網址 drive.google.com/drive/folders/ 後面那一段）
 */
var UPLOAD_FOLDER_ID = '1Mqg3eG1sAW4KBstuNXrX7SQFJGc8niWL';
var MAX_IMAGES = { candidates: 1, itinerary: 6 };
var MAX_IMAGE_BYTES = 3 * 1024 * 1024;

var ITEM_TYPES = ['itinerary', 'expenses', 'candidates', 'checklist', 'comments'];

function checkType_(type) {
  if (ITEM_TYPES.indexOf(type) < 0) throw new Error('未知的類型');
}

var DEFAULT_SETTINGS = {
  tripName: '我們的旅行',
  startDate: '',
  endDate: '',
  members: '成員A,成員B,成員C,成員D',
  currency: 'TWD',
  extraCurrencies: 'THB',
  baseLocation: ''
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
  Logger.log('圖片會存到資料夾：' + DriveApp.getFolderById(UPLOAD_FOLDER_ID).getName());
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
    comments: readRows_('comments'),
    version: APP_VERSION,
    serverTime: new Date().toISOString()
  };
}

/* ---------- 讀寫工具 ---------- */

function sheetOf_(type) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS[type].name);
}

function readItem_(type, rowIndex) {
  var def = SHEETS[type];
  var r = sheetOf_(type).getRange(rowIndex, 1, 1, def.headers.length).getDisplayValues()[0];
  var obj = {};
  def.headers.forEach(function (h, i) { obj[h] = r[i]; });
  return obj;
}

/** 原樣寫入一列（rowIndex 為 0 時新增在最後） */
function writeRow_(type, rowIndex, obj) {
  var def = SHEETS[type];
  var sheet = sheetOf_(type);
  var row = def.headers.map(function (h) {
    var v = obj[h] == null ? '' : String(obj[h]);
    return v.slice(0, type === 'log' ? 45000 : 2000);
  });
  if (rowIndex) {
    sheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
  } else {
    sheet.getRange(sheet.getLastRow() + 1, 1, 1, row.length).setNumberFormat('@').setValues([row]);
  }
}

function titleOf_(type, obj) {
  if (!obj) return '';
  if (type === 'comments') {
    var t = String(obj.text || '').replace(/\s+/g, ' ');
    return t.length > 20 ? t.slice(0, 20) + '…' : t;
  }
  return String(obj.title || obj.item || obj.id || '').slice(0, 60);
}

function actorOf_(opts, item) {
  return String((opts && opts.actor) || (item && item.author) || '').trim().slice(0, 50);
}

/** 衝突錯誤：前端會解析 CONFLICT: 後面的 JSON */
function conflict_(info) {
  return new Error('CONFLICT:' + JSON.stringify(info));
}

/** 新增或更新一筆，回傳 { before, after } */
function writeItem_(type, item) {
  var sheet = sheetOf_(type);
  item.id = item.id || Utilities.getUuid();
  item.updatedAt = new Date().toISOString();
  var rowIndex = findRow_(sheet, item.id);
  var before = rowIndex ? readItem_(type, rowIndex) : null;
  if (type === 'comments') {
    if (['itinerary', 'candidates'].indexOf(item.targetType) < 0) throw new Error('留言對象錯誤');
    item.text = String(item.text || '').trim().slice(0, 500);
    if (!item.text) throw new Error('留言是空的');
    item.createdAt = before ? before.createdAt : item.updatedAt;
  }
  if (type === 'candidates') {
    // 投票只能透過 vote() 修改，編輯內容時保留原本的票與建立時間
    item.votes = before ? before.votes : '{}';
    item.createdAt = before ? before.createdAt : item.updatedAt;
  }
  if (MAX_IMAGES[type]) item.images = JSON.stringify(imageIds_(item.images).slice(0, MAX_IMAGES[type]));
  if (type === 'checklist') {
    // 勾選狀態只能透過 toggleCheck() 修改，編輯內容時保留
    item.done = before ? before.done : '';
    item.checkedBy = before ? before.checkedBy : '[]';
  }
  writeRow_(type, rowIndex, item);
  return { before: before, after: rowIndex ? readItem_(type, rowIndex) : readItem_(type, sheet.getLastRow()) };
}

/* ---------- 修改紀錄 ---------- */

function addLog_(actor, action, summary, changes) {
  syncImages_(changes);
  var id = Utilities.getUuid();
  writeRow_('log', 0, {
    id: id, time: new Date().toISOString(), actor: actor, action: action,
    summary: summary, changes: JSON.stringify(changes), undoneBy: ''
  });
  // 只保留最近的紀錄
  var sheet = sheetOf_('log');
  var extra = sheet.getLastRow() - 1 - LOG_KEEP;
  if (extra > 50) sheet.deleteRows(2, extra);
  return id;
}

/** 讀取最近的修改紀錄（新的在前） */
function getLog(limit) {
  ensureSheets_();
  limit = Math.min(Number(limit) || 80, 200);
  var sheet = sheetOf_('log');
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var n = Math.min(limit, lastRow - 1);
  var headers = SHEETS.log.headers;
  var values = sheet.getRange(lastRow - n + 1, 1, n, headers.length).getDisplayValues();
  return values.reverse().filter(function (r) { return r[0]; }).map(function (r) {
    var obj = {};
    headers.forEach(function (h, i) { obj[h] = r[i]; });
    try { obj.changes = JSON.parse(obj.changes || '[]'); } catch (e) { obj.changes = []; }
    return obj;
  });
}

/* ---------- 對外 API ---------- */

/**
 * 新增或更新一筆（type: itinerary | expenses | candidates | checklist）
 * opts.base：開始編輯時這筆的 updatedAt，用來偵測別人是否在這期間改過
 * opts.force：發生衝突時仍要覆蓋
 */
function saveItem(type, item, opts) {
  checkType_(type);
  opts = opts || {};
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    ensureSheets_();
    if (item.id) {
      var rowIndex = findRow_(sheetOf_(type), item.id);
      if (!rowIndex) {
        if (!opts.force) throw conflict_({ kind: 'deleted', title: titleOf_(type, item) });
        item.id = '';  // 已被刪除：以新項目重新建立
      } else if (opts.base && !opts.force) {
        var cur = readItem_(type, rowIndex);
        if (cur.updatedAt !== opts.base) {
          throw conflict_({ kind: 'modified', by: cur.author, at: cur.updatedAt, title: titleOf_(type, cur) });
        }
      }
    }
    var res = writeItem_(type, item);
    addLog_(actorOf_(opts, item), res.before ? 'update' : 'create',
      (res.before ? '修改' : '新增') + TYPE_LABELS[type] + '「' + titleOf_(type, res.after) + '」',
      [{ type: type, id: res.after.id, before: res.before, after: res.after }]);
  } finally {
    lock.releaseLock();
  }
  return getData();
}

function deleteItem(type, id, opts) {
  checkType_(type);
  opts = opts || {};
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    ensureSheets_();
    var sheet = sheetOf_(type);
    var rowIndex = findRow_(sheet, id);
    if (rowIndex) {
      var before = readItem_(type, rowIndex);
      if (opts.base && !opts.force && before.updatedAt !== opts.base) {
        throw conflict_({ kind: 'modified', by: before.author, at: before.updatedAt, title: titleOf_(type, before) });
      }
      sheet.deleteRow(rowIndex);
      addLog_(actorOf_(opts), 'delete', '刪除' + TYPE_LABELS[type] + '「' + titleOf_(type, before) + '」',
        [{ type: type, id: id, before: before, after: null }]);
    }
  } finally {
    lock.releaseLock();
  }
  return getData();
}

/** 候選投票：value 為 1（讚）、-1（倒讚）或 0（取消）。投票不記入修改紀錄 */
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
    var sheet = sheetOf_('candidates');
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
function scheduleCandidate(candidateId, event, opts) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    ensureSheets_();
    var sheet = sheetOf_('candidates');
    var rowIndex = findRow_(sheet, candidateId);
    if (!rowIndex) throw new Error('這個候選項目已被刪除或排入行程');
    var cand = readItem_('candidates', rowIndex);
    event.id = '';
    var res = writeItem_('itinerary', event);
    sheet.deleteRow(rowIndex);
    var changes = [{ type: 'candidates', id: cand.id, before: cand, after: null },
                   { type: 'itinerary', id: res.after.id, before: null, after: res.after }];
    // 候選底下的留言跟著搬到新的行程
    var cSheet = sheetOf_('comments');
    readRows_('comments').forEach(function (cm) {
      if (cm.targetType !== 'candidates' || cm.targetId !== cand.id) return;
      var r = findRow_(cSheet, cm.id);
      var moved = {};
      Object.keys(cm).forEach(function (k) { moved[k] = cm[k]; });
      moved.targetType = 'itinerary';
      moved.targetId = res.after.id;
      moved.updatedAt = new Date().toISOString();
      writeRow_('comments', r, moved);
      changes.push({ type: 'comments', id: cm.id, before: cm, after: readItem_('comments', r) });
    });
    addLog_(actorOf_(opts, event), 'schedule', '把「' + titleOf_('candidates', cand) + '」排入行程', changes);
  } finally {
    lock.releaseLock();
  }
  return getData();
}

function saveSettings(newSettings, opts) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    ensureSheets_();
    var sheet = sheetOf_('settings');
    var before = getSettings_(), changedBefore = {}, changedAfter = {};
    Object.keys(DEFAULT_SETTINGS).forEach(function (k) {
      if (newSettings[k] == null) return;
      var value = String(newSettings[k]).slice(0, 500);
      if (value === String(before[k])) return;
      changedBefore[k] = before[k];
      changedAfter[k] = value;
      var rowIndex = findRow_(sheet, k);
      if (rowIndex) {
        sheet.getRange(rowIndex, 2).setValue(value);
      } else {
        sheet.appendRow([k, value]);
      }
    });
    if (Object.keys(changedAfter).length) {
      addLog_(actorOf_(opts), 'settings', '修改旅行設定',
        [{ type: 'settings', id: 'settings', before: changedBefore, after: changedAfter }]);
    }
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
function addItems(type, items, opts) {
  checkType_(type);
  if (!items || !items.length) return getData();
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    ensureSheets_();
    var changes = items.slice(0, 100).map(function (item) {
      item.id = '';
      var res = writeItem_(type, item);
      return { type: type, id: res.after.id, before: null, after: res.after };
    });
    addLog_(actorOf_(opts), 'import', '匯入 ' + changes.length + ' 項' + TYPE_LABELS[type], changes);
  } finally {
    lock.releaseLock();
  }
  return getData();
}

/** 清單勾選：「每人都帶」的行李記錄每個人各自的勾選，其餘項目只有一個完成狀態。勾選不記入修改紀錄 */
function toggleCheck(id, name, checked) {
  name = String(name || '').trim().slice(0, 50);
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    ensureSheets_();
    var def = SHEETS.checklist;
    var sheet = sheetOf_('checklist');
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

/* ---------- 復原 ---------- */

// 投票、勾選這類欄位不靠 updatedAt 追蹤，復原內容修改時保留目前的狀態
var LIVE_FIELDS = { candidates: ['votes'], checklist: ['done', 'checkedBy'] };

/**
 * 復原一筆修改紀錄：把每個變更還原成修改前的樣子。
 * 如果之後又有人改過同一筆，會先回傳衝突讓使用者確認（opts.force 為 true 才覆蓋）。
 * 復原本身也會記一筆紀錄，所以可以再「復原這次復原」。
 */
function undo(logId, opts) {
  opts = opts || {};
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    ensureSheets_();
    var logSheet = sheetOf_('log');
    var logRow = findRow_(logSheet, logId);
    if (!logRow) throw new Error('找不到這筆紀錄（可能太舊已被清除）');
    var entry = readItem_('log', logRow);
    if (entry.undoneBy) throw new Error('這筆已經復原過了');
    var changes = [];
    try { changes = JSON.parse(entry.changes || '[]'); } catch (e) {}
    if (!changes.length) throw new Error('這筆紀錄沒有可以復原的內容');

    // 1. 檢查之後是否又被修改
    var settingsNow = getSettings_();
    var conflicts = [];
    changes.forEach(function (c) {
      if (c.type === 'settings') {
        Object.keys(c.after || {}).forEach(function (k) {
          if (String(settingsNow[k]) !== String(c.after[k])) conflicts.push({ title: '設定：' + k, by: '' });
        });
        return;
      }
      var rowIndex = findRow_(sheetOf_(c.type), c.id);
      var cur = rowIndex ? readItem_(c.type, rowIndex) : null;
      var expected = c.after;
      var same = (!cur && !expected) || (cur && expected && cur.updatedAt === expected.updatedAt);
      if (!same) {
        conflicts.push({
          title: titleOf_(c.type, cur || expected || c.before),
          by: cur ? cur.author : '',
          kind: cur ? (expected ? 'modified' : 'recreated') : 'deleted'
        });
      }
    });
    if (conflicts.length && !opts.force) throw conflict_({ kind: 'undo', items: conflicts });

    // 2. 還原
    var inverse = [];
    changes.slice().reverse().forEach(function (c) {
      if (c.type === 'settings') {
        var setSheet = sheetOf_('settings');
        var nowVals = getSettings_(), was = {};
        Object.keys(c.before || {}).forEach(function (k) {
          was[k] = nowVals[k];
          var r = findRow_(setSheet, k);
          if (r) setSheet.getRange(r, 2).setValue(String(c.before[k])); else setSheet.appendRow([k, String(c.before[k])]);
        });
        inverse.push({ type: 'settings', id: 'settings', before: was, after: c.before });
        return;
      }
      var sheet = sheetOf_(c.type);
      var rowIndex = findRow_(sheet, c.id);
      var cur = rowIndex ? readItem_(c.type, rowIndex) : null;
      var after = null;
      if (!c.before) {
        if (rowIndex) sheet.deleteRow(rowIndex);
      } else {
        var restored = {};
        Object.keys(c.before).forEach(function (k) { restored[k] = c.before[k]; });
        if (cur) (LIVE_FIELDS[c.type] || []).forEach(function (k) { restored[k] = cur[k]; });
        restored.updatedAt = new Date().toISOString();
        writeRow_(c.type, rowIndex, restored);
        after = readItem_(c.type, rowIndex || sheet.getLastRow());
      }
      inverse.push({ type: c.type, id: c.id, before: cur, after: after });
    });

    var newId = addLog_(actorOf_(opts), 'undo', '復原：' + entry.summary, inverse);
    // addLog_ 可能清掉舊紀錄，重新找一次位置
    var again = findRow_(logSheet, logId);
    if (again) logSheet.getRange(again, SHEETS.log.headers.indexOf('undoneBy') + 1).setValue(newId);
  } finally {
    lock.releaseLock();
  }
  return getData();
}

/* ---------- 圖片 ---------- */

function imageIds_(value) {
  var list = value;
  if (typeof list === 'string') {
    try { list = JSON.parse(list || '[]'); } catch (e) { list = []; }
  }
  if (!Array.isArray(list)) return [];
  return list.map(String).filter(function (id, i, arr) {
    return /^[A-Za-z0-9_-]{20,100}$/.test(id) && arr.indexOf(id) === i;
  });
}

/** 只處理放在上傳資料夾裡的檔案，避免動到雲端硬碟其他東西 */
function uploadedFile_(id) {
  try {
    var file = DriveApp.getFileById(id);
    var parents = file.getParents();
    while (parents.hasNext()) {
      if (parents.next().getId() === UPLOAD_FOLDER_ID) return file;
    }
  } catch (e) {}
  return null;
}

/**
 * 依照一次操作的前後差異整理圖片：不再被使用的移到垃圾桶，重新被使用的（例如復原）從垃圾桶救回。
 * 圖片處理失敗不影響資料儲存。
 */
function syncImages_(changes) {
  var before = {}, after = {};
  (changes || []).forEach(function (c) {
    if (c.type === 'settings') return;
    if (c.before) imageIds_(c.before.images).forEach(function (id) { before[id] = true; });
    if (c.after) imageIds_(c.after.images).forEach(function (id) { after[id] = true; });
  });
  Object.keys(before).forEach(function (id) {
    if (after[id]) return;
    var f = uploadedFile_(id);
    if (f) try { f.setTrashed(true); } catch (e) {}
  });
  Object.keys(after).forEach(function (id) {
    if (before[id]) return;
    var f = uploadedFile_(id);
    if (f) try { if (f.isTrashed()) f.setTrashed(false); } catch (e) {}
  });
}

/** 上傳一張圖片（前端已壓縮成 JPEG），回傳檔案 ID */
function uploadImage(base64, mimeType, opts) {
  if (!/^image\/(jpeg|png|webp)$/.test(String(mimeType))) throw new Error('只能上傳 JPG、PNG 或 WebP 圖片');
  var bytes = Utilities.base64Decode(String(base64 || ''));
  if (!bytes.length) throw new Error('圖片是空的');
  if (bytes.length > MAX_IMAGE_BYTES) throw new Error('圖片太大了（上限 3MB）');
  var actor = actorOf_(opts) || '匿名';
  var stamp = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyyMMdd_HHmmss');
  var ext = mimeType === 'image/png' ? 'png' : (mimeType === 'image/webp' ? 'webp' : 'jpg');
  var blob = Utilities.newBlob(bytes, mimeType, stamp + '_' + actor + '.' + ext);
  var file = DriveApp.getFolderById(UPLOAD_FOLDER_ID).createFile(blob);
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (e) {
    file.setTrashed(true);
    throw new Error('無法把圖片設成「知道連結的人可以查看」，可能是帳號的共用限制：' + e.message);
  }
  return { id: file.getId() };
}

/** 表單取消時，把剛上傳但最後沒用到的圖片移到垃圾桶 */
function discardImages(ids) {
  ids = imageIds_(ids);
  if (!ids.length) return true;
  var used = {};
  ['candidates', 'itinerary'].forEach(function (type) {
    readRows_(type).forEach(function (r) { imageIds_(r.images).forEach(function (id) { used[id] = true; }); });
  });
  ids.forEach(function (id) {
    if (used[id]) return;
    var f = uploadedFile_(id);
    if (f) try { f.setTrashed(true); } catch (e) {}
  });
  return true;
}

/* ---------- 離線 PDF ---------- */

var PDF_CATEGORY = { transport: '交通', sight: '景點', food: '餐廳', stay: '住宿', other: '其他' };
var PDF_WEEK = ['日', '一', '二', '三', '四', '五', '六'];
var PDF_MAX_IMAGE_BYTES = 40 * 1024 * 1024;

function htmlEsc_(s) {
  return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
  });
}

function pdfDate_(s) {
  var m = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(String(s || ''));
  if (!m) return null;
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
}

function pdfDateLabel_(s) {
  var d = pdfDate_(s);
  return d ? (d.getMonth() + 1) + '/' + d.getDate() + '（' + PDF_WEEK[d.getDay()] + '）' : s;
}

/** PDF 存放的資料夾：上傳圖片資料夾的上一層（例如「2027 泰國曼谷」） */
function pdfFolder_() {
  var upload = DriveApp.getFolderById(UPLOAD_FOLDER_ID);
  var parents = upload.getParents();
  return parents.hasNext() ? parents.next() : upload;
}

function buildItineraryHtml_(settings, items, generatedAt, actor) {
  var start = pdfDate_(settings.startDate);
  var groups = {}, keys = [];
  items.forEach(function (it) {
    var k = pdfDate_(it.date) ? it.date : '';
    if (!groups[k]) { groups[k] = []; keys.push(k); }
    groups[k].push(it);
  });
  keys.sort(function (a, b) { return a === '' ? 1 : (b === '' ? -1 : a.localeCompare(b)); });

  var imageBytes = 0, skipped = 0;
  function imagesHtml(it) {
    var ids = imageIds_(it.images);
    if (!ids.length) return '';
    var out = [];
    ids.forEach(function (id) {
      var f = uploadedFile_(id);
      if (!f || f.isTrashed()) return;
      var blob = f.getBlob();
      var bytes = blob.getBytes();
      if (imageBytes + bytes.length > PDF_MAX_IMAGE_BYTES) { skipped++; return; }
      imageBytes += bytes.length;
      out.push('<img src="data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(bytes) + '" style="width:150px;margin:6px 6px 0 0;border:1px solid #ddd;">');
    });
    return out.length ? '<div>' + out.join('') + '</div>' : '';
  }

  var dateRange = settings.startDate && settings.endDate
    ? pdfDateLabel_(settings.startDate) + ' – ' + pdfDateLabel_(settings.endDate) : '';
  var html = '<html><head><meta charset="utf-8"><style>' +
    'body{font-family:Arial,"Noto Sans TC","Noto Sans Thai",sans-serif;color:#222;font-size:12pt;}' +
    'h1{font-size:22pt;margin:0 0 4px;color:#c85a28;}' +
    '.sub{color:#666;font-size:10pt;margin-bottom:16px;}' +
    'h2{font-size:14pt;background:#2b2a28;color:#fff;padding:5px 10px;margin:18px 0 6px;}' +
    'table{width:100%;border-collapse:collapse;}' +
    'td{vertical-align:top;padding:8px 6px;border-bottom:1px solid #ddd;}' +
    'td.t{width:70px;font-weight:bold;white-space:nowrap;}' +
    '.title{font-size:13pt;font-weight:bold;}' +
    '.cat{color:#888;font-size:9pt;font-weight:normal;}' +
    '.tent{color:#c98a1b;font-size:9pt;font-weight:normal;}' +
    '.loc{font-size:13pt;margin-top:3px;}' +
    '.note{color:#555;font-size:10pt;margin-top:3px;}' +
    '.foot{color:#999;font-size:9pt;margin-top:24px;}' +
    '</style></head><body>' +
    '<h1>' + htmlEsc_(settings.tripName || '旅行行程') + '</h1>' +
    '<div class="sub">' + htmlEsc_(dateRange) + (dateRange ? '　' : '') + '成員：' + htmlEsc_(String(settings.members || '').split(',').join('、')) +
    '<br>離線版產生於 ' + htmlEsc_(generatedAt) + (actor ? '（' + htmlEsc_(actor) + '）' : '') + '，之後的修改請看網頁</div>';

  if (!keys.length) html += '<p>目前還沒有任何行程。</p>';
  keys.forEach(function (k) {
    var list = groups[k].slice().sort(function (a, b) { return (a.time || '99:99').localeCompare(b.time || '99:99'); });
    var dayNo = '';
    if (k && start) {
      var n = Math.round((pdfDate_(k) - start) / 86400000) + 1;
      if (n >= 1) dayNo = 'Day ' + n + '　';
    }
    html += '<h2>' + dayNo + (k ? htmlEsc_(pdfDateLabel_(k)) : '日期未定') + '</h2><table>';
    list.forEach(function (it) {
      html += '<tr><td class="t">' + htmlEsc_(it.time || '—') + (it.endTime ? '<br><span class="cat">~' + htmlEsc_(it.endTime) + '</span>' : '') + '</td><td>' +
        '<div class="title">' + htmlEsc_(it.title) + ' <span class="cat">［' + (PDF_CATEGORY[it.category] || '其他') + '］</span>' +
        (it.status !== 'confirmed' ? ' <span class="tent">待確認</span>' : '') +
        (it.booking === 'need' ? ' <span class="tent">需預約・未訂</span>' : (it.booking === 'booked' ? ' <span class="cat">已預約</span>' : '')) + '</div>' +
        (it.location ? '<div class="loc">地點：' + htmlEsc_(it.location) + '</div>' : '') +
        (it.note ? '<div class="note">' + htmlEsc_(it.note).replace(/\n/g, '<br>') + '</div>' : '') +
        imagesHtml(it) + '</td></tr>';
    });
    html += '</table>';
  });
  if (skipped) html += '<div class="foot">※ 照片太多，有 ' + skipped + ' 張沒有放進 PDF，請到網頁查看。</div>';
  html += '<div class="foot">' + htmlEsc_(settings.tripName || '') + '・旅行共編</div></body></html>';
  return html;
}

/**
 * 產生行程的離線 PDF：存到旅行資料夾、設成知道連結可看，並刪掉舊的離線版。
 * 回傳最新的 PDF 資訊（也會記在設定的 lastPdf，大家都能下載）。
 */
function exportPdf(opts) {
  ensureSheets_();
  var actor = actorOf_(opts);
  var settings = getSettings_();
  var now = new Date();
  var generatedAt = Utilities.formatDate(now, 'Asia/Taipei', 'yyyy/MM/dd HH:mm');
  var html = buildItineraryHtml_(settings, readRows_('itinerary'), generatedAt, actor);
  var baseName = String(settings.tripName || '旅行').replace(/[\\\/:*?"<>|]/g, '').slice(0, 40) + '_離線行程_';
  var pdf = Utilities.newBlob(html, 'text/html', 'itinerary.html').getAs('application/pdf')
    .setName(baseName + Utilities.formatDate(now, 'Asia/Taipei', 'yyyyMMdd_HHmm') + '.pdf');

  var folder = pdfFolder_();
  var file = folder.createFile(pdf);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    // 只清掉之前產生的離線版（檔名格式相同），資料夾裡其他檔案不動
    var old = null;
    try { old = JSON.parse(getSettings_().lastPdf || 'null'); } catch (e) {}
    if (old && old.id && old.id !== file.getId()) {
      try {
        var prev = DriveApp.getFileById(old.id);
        if (/_離線行程_\d{8}_\d{4}\.pdf$/.test(prev.getName())) prev.setTrashed(true);
      } catch (e) {}
    }
    var info = { id: file.getId(), name: file.getName(), at: now.toISOString(), by: actor, size: file.getSize() };
    var sheet = sheetOf_('settings');
    var r = findRow_(sheet, 'lastPdf');
    if (r) sheet.getRange(r, 2).setValue(JSON.stringify(info)); else sheet.appendRow(['lastPdf', JSON.stringify(info)]);
  } finally {
    lock.releaseLock();
  }
  return getData();
}

/* ---------- 距離與地圖（Google 地圖服務，免 API 金鑰） ---------- */

var GEO_SHEET = '地點快取';
var GEO_BUDGET = 40;          // 每次呼叫最多查詢幾次 Google 地圖（避免逾時，剩下的下次再算）
var GEO_HINT = '曼谷';         // 地點沒寫城市時，加上這個再查一次

function geoSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(GEO_SHEET);
  if (!sh) {
    sh = ss.insertSheet(GEO_SHEET);
    sh.getRange(1, 1, sh.getMaxRows(), 3).setNumberFormat('@');
    sh.getRange(1, 1, 1, 3).setValues([['key', 'value', 'at']]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function geoCacheRead_() {
  var sh = geoSheet_(), last = sh.getLastRow(), map = {};
  if (last < 2) return map;
  sh.getRange(2, 1, last - 1, 2).getDisplayValues().forEach(function (r) {
    if (!r[0]) return;
    try { map[r[0]] = JSON.parse(r[1]); } catch (e) {}
  });
  return map;
}

function geoKeyOf_(loc) { return 'g|' + String(loc || '').trim().toLowerCase().replace(/\s+/g, ' '); }
function round5_(n) { return Math.round(Number(n) * 1e5) / 1e5; }

/** Google 地圖分享網址（含短網址）→ 座標或地名 */
function resolveMapsUrl_(url) {
  var target = url;
  try {
    for (var i = 0; i < 3 && /goo\.gl|g\.co\//.test(target); i++) {
      var res = UrlFetchApp.fetch(target, { followRedirects: false, muteHttpExceptions: true });
      var headers = res.getAllHeaders();
      var next = headers.Location || headers.location;
      if (!next) break;
      target = String(next);
    }
  } catch (e) {}
  target = decodeURIComponent(target);
  var m = /!3d(-?\d+\.\d+)!4d(-?\d+\.\d+)/.exec(target) || /@(-?\d+\.\d+),(-?\d+\.\d+)/.exec(target) ||
          /[?&](?:q|query|ll)=(-?\d+\.\d+),\s*(-?\d+\.\d+)/.exec(target);
  var name = (/\/place\/([^\/@?]+)/.exec(target) || [])[1] || (/[?&](?:q|query)=([^&]+)/.exec(target) || [])[1] || '';
  name = name.replace(/\+/g, ' ');
  if (m) return { lat: Number(m[1]), lng: Number(m[2]), name: name };
  return name ? { query: name } : null;
}

function geocodeOnce_(query) {
  var res = Maps.newGeocoder().setLanguage('zh-TW').setRegion('th').geocode(query);
  if (res.status === 'OVER_QUERY_LIMIT' || res.status === 'UNKNOWN_ERROR') throw new Error(res.status);
  if (res.status !== 'OK' || !res.results || !res.results.length) return null;
  return res.results[0];
}

function areaOf_(result) {
  var comps = (result && result.address_components) || [];
  var pick = function (type) {
    for (var i = 0; i < comps.length; i++) if (comps[i].types.indexOf(type) >= 0) return comps[i].long_name;
    return '';
  };
  return pick('sublocality_level_1') || pick('administrative_area_level_2') || pick('locality') || pick('administrative_area_level_1');
}

/** 查一個地點的座標（先看快取） */
function geocode_(loc, cache, budget) {
  var key = geoKeyOf_(loc);
  if (cache[key]) return cache[key];
  if (budget.left <= 0) return null;
  var out = { ok: false };
  try {
    var text = String(loc).trim();
    var query = text, preset = null;
    if (/^https?:\/\//i.test(text)) {
      budget.left--;
      preset = resolveMapsUrl_(text);
      if (preset && preset.query) query = preset.query;
    }
    if (preset && preset.lat != null) {
      budget.left--;
      var rev = null;
      try { rev = Maps.newGeocoder().setLanguage('zh-TW').reverseGeocode(preset.lat, preset.lng); } catch (e) {}
      var r0 = rev && rev.results && rev.results[0];
      out = { ok: true, lat: round5_(preset.lat), lng: round5_(preset.lng), area: areaOf_(r0), addr: preset.name || (r0 && r0.formatted_address) || '' };
    } else if (!/^https?:\/\//i.test(query)) {
      budget.left--;
      var r = geocodeOnce_(query);
      if (!r && !/曼谷|bangkok|กรุงเทพ|泰國|thailand/i.test(query) && budget.left > 0) {
        budget.left--;
        r = geocodeOnce_(query + ' ' + GEO_HINT);
      }
      if (r) out = { ok: true, lat: round5_(r.geometry.location.lat), lng: round5_(r.geometry.location.lng), area: areaOf_(r), addr: r.formatted_address || '' };
    }
  } catch (e) {
    // 暫時性錯誤（例如超過每日用量）不寫入快取，之後會再試
    cache[key] = { ok: false, error: String(e.message || e).slice(0, 100), temp: true };
    return cache[key];
  }
  cache[key] = out;
  cache.__new.push([key, JSON.stringify(out), new Date().toISOString()]);
  return out;
}

var GEO_MODES = { drive: 'DRIVING', transit: 'TRANSIT', walk: 'WALKING' };

function route_(a, b, mode, cache, budget) {
  var key = 'r|' + mode + '|' + a.lat + ',' + a.lng + '|' + b.lat + ',' + b.lng;
  if (cache[key]) return cache[key];
  if (budget.left <= 0) return null;
  budget.left--;
  var out = { ok: false };
  try {
    var dir = Maps.newDirectionFinder().setOrigin(a.lat, a.lng).setDestination(b.lat, b.lng)
      .setMode(Maps.DirectionFinder.Mode[GEO_MODES[mode]]).setLanguage('zh-TW').getDirections();
    var leg = dir && dir.routes && dir.routes[0] && dir.routes[0].legs && dir.routes[0].legs[0];
    if (leg) out = { ok: true, m: leg.distance.value, s: leg.duration.value };
  } catch (e) {
    // 暫時性錯誤（例如超過每日用量）不寫入快取，之後會再試
    cache[key] = { ok: false, error: String(e.message || e).slice(0, 100), temp: true };
    return cache[key];
  }
  cache[key] = out;
  cache.__new.push([key, JSON.stringify(out), new Date().toISOString()]);
  return out;
}

/**
 * 前端要什麼就算什麼：
 * req.locations：地點文字陣列 → 回傳座標、區域
 * req.routes：[{from, to, mode}]（mode: drive | transit | walk）→ 回傳距離（公尺）、時間（秒）
 * 結果會存在「地點快取」工作表，同樣的地點和路線不會重查。
 */
function getGeo(req) {
  req = req || {};
  var cache = geoCacheRead_();
  cache.__new = [];
  var budget = { left: GEO_BUDGET };
  var places = {}, routes = {}, pending = 0;
  (req.locations || []).slice(0, 300).forEach(function (loc) {
    loc = String(loc || '').trim();
    if (!loc || places[loc]) return;
    var g = geocode_(loc, cache, budget);
    if (g) places[loc] = g; else pending++;
  });
  (req.routes || []).slice(0, 300).forEach(function (r) {
    var mode = GEO_MODES[r.mode] ? r.mode : 'drive';
    var k = mode + '|' + r.from + '|' + r.to;
    if (routes[k]) return;
    var a = geocode_(r.from, cache, budget), b = geocode_(r.to, cache, budget);
    if (!a || !b) { pending++; return; }
    if (!a.ok || !b.ok) { routes[k] = { ok: false, error: 'place' }; return; }
    var x = route_(a, b, mode, cache, budget);
    if (x) routes[k] = x; else pending++;
  });
  if (cache.__new.length) {
    var lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      var sh = geoSheet_();
      sh.getRange(sh.getLastRow() + 1, 1, cache.__new.length, 3).setNumberFormat('@').setValues(cache.__new);
    } finally {
      lock.releaseLock();
    }
  }
  return { places: places, routes: routes, pending: pending };
}

/** 候選地圖總覽：points = [{lat, lng, label, color}]，回傳圖片 data URL */
function getOverviewMap(points, base) {
  var map = Maps.newStaticMap().setSize(640, 640).setLanguage('zh-TW').setMapType(Maps.StaticMap.Type.ROADMAP);
  if (base && base.lat != null) {
    map.setMarkerStyle(Maps.StaticMap.MarkerSize.MID, Maps.StaticMap.Color.BLACK, 'H');
    map.addMarker(Number(base.lat), Number(base.lng));
  }
  (points || []).slice(0, 35).forEach(function (p) {
    map.setMarkerStyle(Maps.StaticMap.MarkerSize.MID, p.color === 'green' ? Maps.StaticMap.Color.GREEN :
      (p.color === 'red' ? Maps.StaticMap.Color.RED : Maps.StaticMap.Color.ORANGE), String(p.label || '').slice(0, 1));
    map.addMarker(Number(p.lat), Number(p.lng));
  });
  var blob = map.getBlob();
  return 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
}
