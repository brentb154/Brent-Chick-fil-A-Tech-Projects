/**
 * Standalone diagnostic: runs the exact same parse + OT logic as MondaySheet.gs
 * on the two raw CSVs, then prints where the Sched OT number comes from.
 *
 * Not part of the GAS project — run with `node diagnose_ot.js`.
 */

var fs = require('fs');

// ── Mirrored verbatim from Parser.gs ──────────────────────────────────────

function normalizeName(raw) {
  if (!raw) return '';
  var s = String(raw).replace(/\s+/g, ' ').trim();
  if (!s) return '';

  var comma = s.indexOf(',');
  if (comma > 0 && comma < s.length - 1) {
    var last = s.substring(0, comma).trim();
    var rest = s.substring(comma + 1).trim();
    if (last && rest) s = rest + ' ' + last;
  }

  s = s.toLowerCase().replace(/\b([a-z])/g, function(_, c) { return c.toUpperCase(); });
  return s; // no alias map in this environment
}

function parseTimestamp(val) {
  if (!val) return null;
  if (val instanceof Date && !isNaN(val.getTime())) return val;
  var s = String(val).trim();
  if (!s) return null;

  var m = s.match(/(\d{4})-(\d{2})-(\d{2})\s+(\d{2}):(\d{2}):(\d{2})/);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3], +m[4], +m[5], +m[6]);

  m = s.match(/(\d{1,2})\/(\d{1,2})\/(\d{2,4})\s+(\d{1,2}):(\d{2})/);
  if (m) {
    var y = +m[3];
    if (y < 100) y += 2000;
    return new Date(y, +m[1] - 1, +m[2], +m[4], +m[5], 0);
  }
  return null;
}

function parseScheduleFromSheet(data) {
  var headers = data[0].map(function(h) { return String(h).toUpperCase().replace(/[^A-Z_]/g, ''); });
  var iN  = headers.indexOf('FULL_NAME');
  var iSS = headers.indexOf('SCHEDULED_START_TIME');
  var iSE = headers.indexOf('SCHEDULED_END_TIME');

  if (iN < 0 || iSS < 0 || iSE < 0) throw new Error('Missing required columns');

  var result = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var name = normalizeName(row[iN]);
    if (!name) continue;
    var st = parseTimestamp(row[iSS]);
    var en = parseTimestamp(row[iSE]);
    if (!st || !en) continue;
    var mins = Math.round((en - st) / 60000);
    if (mins <= 0) continue;
    result.push({ name: name, start: st, end: en, minutes: mins });
  }
  return result;
}

// ── CSV parser (mirrors parseCSVToArray_ in Code.gs) ──────────────────────

function parseCSVToArray(text) {
  if (!text || !text.trim()) return [];
  var firstNL = text.indexOf('\n');
  if (firstNL < 0) firstNL = text.length;
  var firstLine = text.substring(0, firstNL);
  var delim = (firstLine.match(/\t/g) || []).length > (firstLine.match(/,/g) || []).length ? '\t' : ',';

  var rows = [], row = [], field = '', inQ = false;
  for (var i = 0; i < text.length; i++) {
    var ch = text[i];
    if (inQ) {
      if (ch === '"') {
        if (i + 1 < text.length && text[i + 1] === '"') { field += '"'; i++; }
        else inQ = false;
      } else { field += ch; }
    } else {
      if (ch === '"') { inQ = true; }
      else if (ch === delim) { row.push(field.trim()); field = ''; }
      else if (ch === '\n' || ch === '\r') {
        if (ch === '\r' && i + 1 < text.length && text[i + 1] === '\n') i++;
        row.push(field.trim()); field = '';
        if (row.length > 1 || (row.length === 1 && row[0] !== '')) rows.push(row);
        row = [];
      } else { field += ch; }
    }
  }
  if (field || row.length > 0) {
    row.push(field.trim());
    if (row.length > 1 || (row.length === 1 && row[0] !== '')) rows.push(row);
  }
  return rows;
}

// ── Run on real files ─────────────────────────────────────────────────────

var chText  = fs.readFileSync(__dirname + '/Cockrell Hill.csv', 'utf8');
var dbuText = fs.readFileSync(__dirname + '/DBU.csv', 'utf8');

var chRows  = parseScheduleFromSheet(parseCSVToArray(chText));
var dbuRows = parseScheduleFromSheet(parseCSVToArray(dbuText));

var OT_THRESHOLD_MIN = 40 * 60;
var LOC1 = 'Cockrell Hill DTO';
var LOC2 = 'DBU OCV';

// Build empMap[name][loc] = minutes (same shape as MondaySheet.gs:228-253)
var empMap = {};
function addRows(rows, loc) {
  rows.forEach(function(r) {
    if (!empMap[r.name]) empMap[r.name] = {};
    if (!empMap[r.name][loc]) empMap[r.name][loc] = 0;
    empMap[r.name][loc] += r.minutes;
  });
}
addRows(chRows, LOC1);
addRows(dbuRows, LOC2);

// Replay the exact OT distribution from MondaySheet.gs:269-281
var locOT = {}; locOT[LOC1] = 0; locOT[LOC2] = 0;
var totalOT = 0;
var otBreakdown = [];

Object.keys(empMap).forEach(function(name) {
  var locs = empMap[name];
  var totalMin = 0;
  Object.keys(locs).forEach(function(l) { totalMin += locs[l]; });
  var otMin = Math.max(0, totalMin - OT_THRESHOLD_MIN);
  if (otMin <= 0) return;

  totalOT += otMin;
  var row = { name: name, totalHrs: totalMin / 60, otHrs: otMin / 60, allocations: {} };
  Object.keys(locs).forEach(function(l) {
    var share = totalMin > 0 ? locs[l] / totalMin : 0;
    var credit = otMin * share;
    locOT[l] += credit;
    row.allocations[l] = { scheduledHrs: locs[l] / 60, otCredit: credit / 60 };
  });
  otBreakdown.push(row);
});

// Also compute "location-only OT" — i.e. OT if you ONLY looked at one sheet
function locationOnlyOT(rows) {
  var m = {};
  rows.forEach(function(r) { m[r.name] = (m[r.name] || 0) + r.minutes; });
  var total = 0, list = [];
  Object.keys(m).forEach(function(n) {
    var ot = Math.max(0, m[n] - OT_THRESHOLD_MIN);
    if (ot > 0) { total += ot; list.push({ name: n, hrs: m[n] / 60, ot: ot / 60 }); }
  });
  return { total: total / 60, list: list };
}

var chOnly  = locationOnlyOT(chRows);
var dbuOnly = locationOnlyOT(dbuRows);

// ── Output ────────────────────────────────────────────────────────────────

function hrs(v) { return (Math.round(v * 10) / 10).toFixed(1); }

console.log('\n========================================================');
console.log('  PARSER SANITY CHECK');
console.log('========================================================');
console.log('CH  rows parsed: ' + chRows.length + '  (' + new Set(chRows.map(function(r){return r.name;})).size + ' unique employees)');
console.log('DBU rows parsed: ' + dbuRows.length + '  (' + new Set(dbuRows.map(function(r){return r.name;})).size + ' unique employees)');
var chSum  = chRows.reduce(function(t,r){return t+r.minutes;}, 0) / 60;
var dbuSum = dbuRows.reduce(function(t,r){return t+r.minutes;}, 0) / 60;
console.log('CH  total scheduled hours: ' + hrs(chSum));
console.log('DBU total scheduled hours: ' + hrs(dbuSum));

console.log('\n========================================================');
console.log('  WHAT THE CODE COMPUTES (cross-location OT, distributed)');
console.log('========================================================');
console.log('CH  Sched OT: ' + hrs(locOT[LOC1] / 60));
console.log('DBU Sched OT: ' + hrs(locOT[LOC2] / 60));
console.log('Total Sched OT: ' + hrs(totalOT / 60));

console.log('\n========================================================');
console.log('  NEW LAYOUT — what will appear in the Monday sheet');
console.log('========================================================');
console.log('  Sched OT (combined)    CH=' + hrs(locOT[LOC1]/60) + '   DBU=' + hrs(locOT[LOC2]/60) + '   Combined=' + hrs(totalOT/60));
console.log('  Sched OT (on-site)     CH=' + hrs(chOnly.total) + '    DBU=' + hrs(dbuOnly.total) + '    Combined=' + hrs(chOnly.total+dbuOnly.total));

console.log('\n========================================================');
console.log('  WHAT YOU\'D EXPECT (location-only: >40h at that location)');
console.log('========================================================');
console.log('CH-only OT:  ' + hrs(chOnly.total) + '  (' + chOnly.list.length + ' employees)');
chOnly.list.forEach(function(e){ console.log('   ' + e.name + ': ' + hrs(e.hrs) + 'h → ' + hrs(e.ot) + 'h OT'); });
console.log('DBU-only OT: ' + hrs(dbuOnly.total) + '  (' + dbuOnly.list.length + ' employees)');
dbuOnly.list.forEach(function(e){ console.log('   ' + e.name + ': ' + hrs(e.hrs) + 'h → ' + hrs(e.ot) + 'h OT'); });

console.log('\n========================================================');
console.log('  PER-EMPLOYEE OT BREAKDOWN (who is driving the number)');
console.log('========================================================');
otBreakdown.sort(function(a,b){return b.otHrs - a.otHrs;});
otBreakdown.forEach(function(row) {
  var locs = Object.keys(row.allocations);
  var multi = locs.length > 1 ? '  [MULTI-LOC]' : '';
  console.log('\n' + row.name + multi);
  console.log('   Total scheduled: ' + hrs(row.totalHrs) + 'h  →  OT: ' + hrs(row.otHrs) + 'h');
  locs.forEach(function(l) {
    var a = row.allocations[l];
    console.log('     • ' + l + ': scheduled ' + hrs(a.scheduledHrs) + 'h, credited ' + hrs(a.otCredit) + 'h OT');
  });
});

console.log('\n========================================================');
console.log('  CROSS-CHECK: names appearing in BOTH sheets');
console.log('========================================================');
var chNames  = new Set(chRows.map(function(r){return r.name;}));
var dbuNames = new Set(dbuRows.map(function(r){return r.name;}));
var both = [];
chNames.forEach(function(n){ if (dbuNames.has(n)) both.push(n); });
if (both.length === 0) {
  console.log('None. (No one is scheduled at both locations this week.)');
} else {
  both.forEach(function(n) {
    var chH  = chRows.filter(function(r){return r.name===n;}).reduce(function(t,r){return t+r.minutes;},0)/60;
    var dbuH = dbuRows.filter(function(r){return r.name===n;}).reduce(function(t,r){return t+r.minutes;},0)/60;
    console.log('   ' + n + ': CH=' + hrs(chH) + 'h, DBU=' + hrs(dbuH) + 'h, combined=' + hrs(chH+dbuH) + 'h');
  });
}
console.log('');
