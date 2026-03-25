/**
 * Schedule Variance Analyzer — Parsing & Formatting
 *
 * Reads 2D arrays from sheet tabs, parses timestamps,
 * groups records by employee + date. No CSV parsing needed
 * since data is pasted directly into sheets.
 */

/* ── Schedule Parser ── */

function parseScheduleFromSheet(data) {
  var headers = data[0].map(function(h) { return String(h).toUpperCase().replace(/[^A-Z_]/g, ''); });
  var iN  = headers.indexOf('FULL_NAME');
  var iSS = headers.indexOf('SCHEDULED_START_TIME');
  var iSE = headers.indexOf('SCHEDULED_END_TIME');

  if (iN < 0 || iSS < 0 || iSE < 0) {
    throw new Error('Schedule tab missing required columns: FULL_NAME, SCHEDULED_START_TIME, SCHEDULED_END_TIME');
  }

  var result = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var name = String(row[iN] || '').trim();
    if (!name) continue;

    var st = parseTimestamp(row[iSS]);
    var en = parseTimestamp(row[iSE]);
    if (!st || !en) continue;

    var mins = Math.round((en - st) / 60000);
    if (mins <= 0) continue;

    result.push({
      name: name,
      date: formatDateKey(st),
      start: st,
      end: en,
      minutes: mins
    });
  }

  if (!result.length) throw new Error('No valid schedule rows found. Check that data includes FULL_NAME, SCHEDULED_START_TIME, SCHEDULED_END_TIME columns.');
  return result;
}

/* ── Punch Parser ── */

function parsePunchesFromSheet(data, cfg) {
  var headers = data[0].map(function(h) { return String(h).toUpperCase().replace(/[^A-Z_]/g, ''); });
  var iE  = headers.indexOf('EMPLOYEE');
  var iPT = headers.indexOf('PUNCHTYPE');
  var iTI = headers.indexOf('TIMEIN');
  var iTO = headers.indexOf('TIMEOUT');
  var iRE = headers.indexOf('REMARKS');

  if (iE < 0 || iTI < 0 || iTO < 0) {
    throw new Error('Punch tab missing required columns: EMPLOYEE, TIMEIN, TIMEOUT');
  }

  var result = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var name = String(row[iE] || '').trim();
    if (!name) continue;

    if (iPT >= 0 && row[iPT] && String(row[iPT]).toLowerCase() !== 'regular') continue;

    var tIn  = parseTimestamp(row[iTI]);
    var tOut = parseTimestamp(row[iTO]);
    if (!tIn || !tOut) continue;

    var mins = Math.round((tOut - tIn) / 60000);
    if (mins <= 0) continue;

    var remarks = (iRE >= 0 && row[iRE]) ? String(row[iRE]).replace(/\n/g, ' ').trim() : '';

    result.push({
      name: name,
      date: formatDateKey(tIn),
      timeIn: tIn,
      timeOut: tOut,
      minutes: mins,
      remarks: remarks
    });
  }

  if (!result.length) throw new Error('No valid punch rows found. Check that data includes EMPLOYEE, TIMEIN, TIMEOUT columns and that PUNCHTYPE is "Regular".');
  return result;
}

/* ── Timestamp Parsing ── */

function parseTimestamp(val) {
  if (!val) return null;

  if (val instanceof Date && !isNaN(val.getTime())) {
    return val;
  }

  var s = String(val).trim();
  if (!s) return null;

  // YYYY-MM-DD HH:MM:SS.0
  var m = s.match(/(\d{4})-(\d{2})-(\d{2})\s+(\d{2}):(\d{2}):(\d{2})/);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3], +m[4], +m[5], +m[6]);

  // M/D/YYYY H:MM
  m = s.match(/(\d{1,2})\/(\d{1,2})\/(\d{2,4})\s+(\d{1,2}):(\d{2})/);
  if (m) {
    var y = +m[3];
    if (y < 100) y += 2000;
    return new Date(y, +m[1] - 1, +m[2], +m[4], +m[5], 0);
  }

  return null;
}

/* ── Formatting Utilities ── */

function formatDateKey(dt) {
  return dt.getFullYear() + '-' +
    String(dt.getMonth() + 1).padStart(2, '0') + '-' +
    String(dt.getDate()).padStart(2, '0');
}

function formatTime12(dt) {
  if (!dt || isNaN(dt.getTime())) return '—';
  var h = dt.getHours(), m = dt.getMinutes(), suffix = 'AM';
  if (h >= 12) { suffix = 'PM'; if (h > 12) h -= 12; }
  if (h === 0) h = 12;
  return h + ':' + String(m).padStart(2, '0') + ' ' + suffix;
}

function formatHours(m) {
  if (m === null || m === undefined || isNaN(m)) return '—';
  return (m / 60).toFixed(1) + ' hrs';
}

function formatVariance(v) {
  if (v === 0 || v === null || v === undefined) return '0 min';
  var sign = v > 0 ? '+' : '-';
  var abs = Math.abs(v);
  if (abs < 60) return sign + abs + ' min';
  var hrs = Math.floor(abs / 60);
  var mins = abs % 60;
  return sign + hrs + ' hr' + (mins ? ' ' + mins + ' min' : '');
}

function formatFriendlyDate(dateStr) {
  var DAYS = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
  if (!dateStr) return 'Unknown';
  var p = dateStr.split('-');
  if (p.length !== 3) return dateStr;
  var dt = new Date(parseInt(p[0]), parseInt(p[1]) - 1, parseInt(p[2]));
  return DAYS[dt.getDay()] + ' ' + parseInt(p[1]) + '/' + parseInt(p[2]);
}
