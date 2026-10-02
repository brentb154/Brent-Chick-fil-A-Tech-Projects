/**
 * ============================================================
 * CHECKLIST SYSTEM - Shared Helpers
 * ============================================================
 * Tab access by header name, Settings lookup, and ALL date and
 * time handling.
 *
 * TIME RULES - read before changing anything in this file:
 *  1. Clock times (Due by, Late after, the Settings times) are
 *     read as the text the sheet DISPLAYS ("11:00 PM") and parsed
 *     into minutes after midnight. Never read them with
 *     getValues(): a time-only cell comes back as a Date on
 *     12/30/1899 whose hour depends on time zone settings.
 *  2. Dates are stored as plain text "yyyy-MM-dd" in cells
 *     formatted as text, so Sheets can't convert them.
 *  3. Business clock: a time before "Business day ends at" belongs
 *     to the night before and gets +1440 minutes. 12:30 AM = 1470,
 *     which sorts after 11:30 PM = 1410. Every deadline check
 *     compares minutes on this clock.
 *  4. Never add or subtract milliseconds to move between days.
 *     DST days are 23 or 25 hours long. Day math uses UTC date
 *     parts only (addDays_).
 */

var TABS = {
  settings: 'Settings',
  checklists: 'Checklists',
  positions: 'Positions',
  items: 'Items',
  submissions: 'Submissions',
  results: 'Item Results',
  status: 'Daily Status',
  alerts: 'Alert Log',
  formMap: 'Form Map'
};

var DAY_SHORT = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
var DAY_LONG = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
var MONTH_LONG = ['January', 'February', 'March', 'April', 'May', 'June', 'July',
  'August', 'September', 'October', 'November', 'December'];

// -- Sheet access ---------------------------------------------

function getSS_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) return ss;
  return SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SS_ID'));
}

function getSheet_(name) {
  var sheet = getSS_().getSheetByName(name);
  if (!sheet) throw new Error('Missing tab: "' + name + '"');
  return sheet;
}

// "Form ID (script fills)" -> "form id"
function normHeader_(text) {
  return String(text || '').replace(/\(.*?\)/g, '').replace(/\s+/g, ' ').trim().toLowerCase();
}

// Header row only. Use for appending without reading the whole tab.
function openTab_(name) {
  var sheet = getSheet_(name);
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0].map(normHeader_);
  return {
    sheet: sheet,
    name: name,
    headers: headers,
    width: headers.length,
    col: function (header) { return headers.indexOf(normHeader_(header)); }
  };
}

// Whole tab. display = what the sheet shows (use for text, dates, clock times).
// values = raw values, only when asked (use for full timestamps like "Submitted at").
function readTab_(name, withValues) {
  var tab = openTab_(name);
  var rowCount = tab.sheet.getLastRow() - 1;
  tab.display = [];
  tab.values = [];
  if (rowCount > 0) {
    var range = tab.sheet.getRange(2, 1, rowCount, tab.width);
    tab.display = range.getDisplayValues();
    if (withValues) tab.values = range.getValues();
  }
  return tab;
}

function colOrThrow_(tab, header) {
  var i = tab.col(header);
  if (i < 0) throw new Error('"' + tab.name + '" tab is missing the "' + header + '" column.');
  return i;
}

// { 'Header': value } -> row array in this tab's column order
function toRow_(tab, obj) {
  var row = [];
  for (var c = 0; c < tab.width; c++) row.push('');
  Object.keys(obj).forEach(function (header) {
    var value = obj[header];
    if (typeof value === 'string' && value.charAt(0) === '=') value = "'" + value; // never a formula
    row[colOrThrow_(tab, header)] = value;
  });
  return row;
}

// Writes rows starting at startRow. Grows the sheet if needed. Any column holding text is
// formatted as plain text first so Sheets can't turn "11:00 PM" or "2026-10-06" into a time/date.
function writeRows_(tab, startRow, rows) {
  if (!rows.length) return;
  var needed = startRow + rows.length - 1 - tab.sheet.getMaxRows();
  if (needed > 0) tab.sheet.insertRowsAfter(tab.sheet.getMaxRows(), needed);
  for (var c = 0; c < tab.width; c++) {
    var hasText = rows.some(function (r) { return typeof r[c] === 'string' && r[c] !== ''; });
    if (hasText) tab.sheet.getRange(startRow, c + 1, rows.length, 1).setNumberFormat('@');
  }
  tab.sheet.getRange(startRow, 1, rows.length, tab.width).setValues(rows);
}

function appendRows_(tab, objects) {
  if (!objects.length) return;
  writeRows_(tab, tab.sheet.getLastRow() + 1, objects.map(function (o) { return toRow_(tab, o); }));
}

// rowIndex is 0-based into tab.display
function setCell_(tab, rowIndex, header, value) {
  tab.sheet.getRange(rowIndex + 2, colOrThrow_(tab, header) + 1).setValue(value);
}

function withLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function isTrue_(v) {
  return String(v).trim().toUpperCase() === 'TRUE';
}

// -- Settings -------------------------------------------------

function loadSettings_() {
  var tab = readTab_(TABS.settings);
  var labelCol = colOrThrow_(tab, 'Setting');
  var valueCol = colOrThrow_(tab, 'Value');
  var map = {};
  tab.display.forEach(function (r) {
    var label = normHeader_(r[labelCol]);
    if (label) map[label] = String(r[valueCol]).trim();
  });
  var get = function (label) { return map[normHeader_(label)] || ''; };
  var reminder = parseInt(get('Reminder before due (minutes)'), 10);

  // Bad or blank times fall back to the spec defaults so tracking keeps running; Validate flags them.
  return {
    get: get,
    tz: getSS_().getSpreadsheetTimeZone(),
    dayEndMin: minutesOr_(get('Business day ends at'), 4 * 60),
    rebuildMin: minutesOr_(get('Rebuild forms at'), 4 * 60 + 30),
    summaryMin: minutesOr_(get('Daily summary at'), 7 * 60),
    reminderMin: isNaN(reminder) ? 30 : reminder,
    live: /^live$/i.test(get('Alerts mode'))
  };
}

// -- Clock times ----------------------------------------------

// Displayed time -> minutes after midnight (0-1439), or null if it isn't a time.
// Accepts "11:00 PM", "11 PM", "11pm", "11:00:00 PM", "23:00", "12/30/1899 23:00:00".
function parseTimeToMinutes_(text) {
  var s = String(text || '').trim().toUpperCase();
  if (!s) return null;
  var m = s.match(/(\d{1,2})(?::(\d{2}))?(?::\d{2})?\s*([AP])\.?\s*M\b/);
  if (m) {
    var h = Number(m[1]);
    var min = Number(m[2] || 0);
    if (h < 1 || h > 12 || min > 59) return null;
    if (h === 12) h = 0;          // 12:xx AM -> 0:xx, 12:xx PM -> 12:xx (after +12)
    if (m[3] === 'P') h += 12;
    return h * 60 + min;
  }
  m = s.match(/(\d{1,2}):(\d{2})/);
  if (m) {
    var h24 = Number(m[1]);
    var min24 = Number(m[2]);
    if (h24 > 23 || min24 > 59) return null;
    return h24 * 60 + min24;
  }
  return null;
}

function minutesOr_(text, fallback) {
  var m = parseTimeToMinutes_(text);
  return m === null ? fallback : m;
}

// Wall-clock minutes -> business-clock minutes (see TIME RULES #3)
function bizMinutes_(wallMin, cfg) {
  return wallMin < cfg.dayEndMin ? wallMin + 1440 : wallMin;
}

// -- Dates ----------------------------------------------------

function dateKey_(date, tz) {
  return Utilities.formatDate(date, tz, 'yyyy-MM-dd');
}

function wallMinutes_(date, tz) {
  var hm = Utilities.formatDate(date, tz, 'HH:mm').split(':');
  return Number(hm[0]) * 60 + Number(hm[1]);
}

// Which business day an instant counts toward, and where it sits on that day's clock.
function businessMoment_(date, cfg) {
  var key = dateKey_(date, cfg.tz);
  var wall = wallMinutes_(date, cfg.tz);
  if (wall < cfg.dayEndMin) return { dateKey: addDays_(key, -1), minutes: wall + 1440 };
  return { dateKey: key, minutes: wall };
}

function pad2_(n) {
  return ('0' + Number(n)).slice(-2);
}

// Day math on "yyyy-MM-dd" in UTC, so DST can never shift the date.
function addDays_(key, n) {
  var p = key.split('-');
  var d = new Date(Date.UTC(Number(p[0]), Number(p[1]) - 1, Number(p[2]) + n));
  return d.getUTCFullYear() + '-' + pad2_(d.getUTCMonth() + 1) + '-' + pad2_(d.getUTCDate());
}

// 0 = Sunday ... 6 = Saturday
function dayIndex_(key) {
  var p = key.split('-');
  return new Date(Date.UTC(Number(p[0]), Number(p[1]) - 1, Number(p[2]))).getUTCDay();
}

// "Tue 10/6"
function shortLabel_(key) {
  var p = key.split('-');
  return DAY_SHORT[dayIndex_(key)] + ' ' + Number(p[1]) + '/' + Number(p[2]);
}

// "Tuesday, October 6, 2026"
function longLabel_(key) {
  var p = key.split('-');
  return DAY_LONG[dayIndex_(key)] + ', ' + MONTH_LONG[Number(p[1]) - 1] + ' ' + Number(p[2]) + ', ' + p[0];
}

// Date cell -> "yyyy-MM-dd". Handles text keys, "10/6/2026", and Date objects.
function toDateKey_(v, tz) {
  if (v instanceof Date) return Utilities.formatDate(v, tz || getSS_().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  var s = String(v || '').trim();
  var m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return m[1] + '-' + pad2_(m[2]) + '-' + pad2_(m[3]);
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return m[3] + '-' + pad2_(m[1]) + '-' + pad2_(m[2]);
  return '';
}

// -- Email ----------------------------------------------------

function isEmail_(s) {
  return /^[^@\s,;]+@[^@\s,;]+\.[^@\s,;]+$/.test(s);
}

// "a@x.com, b@y.com" -> valid addresses only
function emailList_(text) {
  return String(text || '').split(/[,;\s]+/).filter(isEmail_);
}

function esc_(s) {
  return String(s === null || s === undefined ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
