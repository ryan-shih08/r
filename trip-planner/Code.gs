/**
 * 旅行共編網頁：Google Apps Script 後端
 * 資料存在綁定的 Google 試算表（行程 / 分攤 / 候選 / 清單 / 設定 / 紀錄 / 匯率 等工作表）
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
  },
  log: {
    name: '紀錄',
    headers: ['id', 'time', 'actor', 'action', 'summary', 'changes', 'undoneBy']
  }
};

var TYPE_LABELS = { itinerary: '行程', expenses: '支出', candidates: '候選', checklist: '清單', settings: '設定' };
var LOG_KEEP = 300;

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
  if (type === 'candidates') {
    // 投票只能透過 vote() 修改，編輯內容時保留原本的票與建立時間
    item.votes = before ? before.votes : '{}';
    item.createdAt = before ? before.createdAt : item.updatedAt;
  }
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
    addLog_(actorOf_(opts, event), 'schedule', '把「' + titleOf_('candidates', cand) + '」排入行程',
      [{ type: 'candidates', id: cand.id, before: cand, after: null },
       { type: 'itinerary', id: res.after.id, before: null, after: res.after }]);
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
