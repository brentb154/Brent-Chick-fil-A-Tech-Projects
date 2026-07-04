// ── Tool Adoption Tracker ────────────────────────────────────
// Bound to the "Tool Adoption" spreadsheet. Each hub tool appends one row
// per day to the Pings tab (via its logAdoptionPing_ helper). This script
// emails a weekly digest of which tools actually got opened.
//
// Setup: run runInitialSetup() once, put your email in the Settings tab,
// then run createWeeklyTrigger(). See README.md.

var PINGS_TAB = 'Pings';
var SETTINGS_TAB = 'Settings';

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Adoption Tracker')
    .addItem('Send Digest Now', 'sendWeeklyAdoptionDigest')
    .addSeparator()
    .addItem('Install Weekly Trigger (Mon 6 AM)', 'createWeeklyTrigger')
    .addItem('Remove Triggers', 'deleteTriggers')
    .addItem('Run Initial Setup', 'runInitialSetup')
    .addToUi();
}

// Idempotent — safe to re-run, never touches existing data.
function runInitialSetup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var pings = ss.getSheetByName(PINGS_TAB);
  if (!pings) {
    pings = ss.insertSheet(PINGS_TAB);
    pings.getRange(1, 1, 1, 2).setValues([['Date', 'Tool']]).setFontWeight('bold');
    pings.setFrozenRows(1);
    pings.getRange('A:A').setNumberFormat('@'); // keep yyyy-MM-dd strings as strings
  }

  var settings = ss.getSheetByName(SETTINGS_TAB);
  if (!settings) {
    settings = ss.insertSheet(SETTINGS_TAB);
    settings.getRange(1, 1, 2, 2).setValues([
      ['Setting', 'Value'],
      ['Digest Email', Session.getEffectiveUser().getEmail()]
    ]);
    settings.getRange(1, 1, 1, 2).setFontWeight('bold');
  }

  SpreadsheetApp.flush();
}

function getDigestEmail_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SETTINGS_TAB);
  if (!sheet) return Session.getEffectiveUser().getEmail();
  var rows = sheet.getDataRange().getValues();
  for (var r = 1; r < rows.length; r++) {
    if (String(rows[r][0]).trim() === 'Digest Email' && String(rows[r][1]).trim()) {
      return String(rows[r][1]).trim();
    }
  }
  return Session.getEffectiveUser().getEmail();
}

// Trigger handler — alerts by email if it fails.
function sendWeeklyAdoptionDigest() {
  try {
    var email = getDigestEmail_();
    MailApp.sendEmail({
      to: email,
      subject: 'Tool Adoption — week of ' + formatDate_(weekStart_(new Date())),
      htmlBody: buildDigestHtml_()
    });
  } catch (err) {
    try {
      MailApp.sendEmail(
        Session.getEffectiveUser().getEmail(),
        'ALERT: Adoption digest failed',
        'sendWeeklyAdoptionDigest threw:\n\n' + (err && err.stack ? err.stack : err)
      );
    } catch (ignored) {}
  }
}

function buildDigestHtml_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PINGS_TAB);
  if (!sheet || sheet.getLastRow() < 2) {
    return '<p>No pings recorded yet. Make sure each tool has its ADOPTION_SHEET_ID script property set.</p>';
  }
  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();

  var now = new Date();
  var thisWeekStart = weekStart_(now);
  var fourWeeksAgo = new Date(thisWeekStart.getTime() - 28 * 86400000);

  // Per tool: days touched this week, days touched in prior 4 weeks, last seen
  var tools = {};
  rows.forEach(function(row) {
    var dateStr = String(row[0]).slice(0, 10);
    var tool = String(row[1]).trim();
    if (!tool || !dateStr) return;
    var d = new Date(dateStr + 'T12:00:00');
    if (isNaN(d.getTime())) return;
    if (!tools[tool]) tools[tool] = { thisWeek: 0, prior4: 0, lastSeen: '' };
    if (d >= thisWeekStart) tools[tool].thisWeek++;
    else if (d >= fourWeeksAgo) tools[tool].prior4++;
    if (dateStr > tools[tool].lastSeen) tools[tool].lastSeen = dateStr;
  });

  var names = Object.keys(tools).sort(function(a, b) {
    return tools[b].thisWeek - tools[a].thisWeek || a.localeCompare(b);
  });

  var html = '<h2 style="font-family:Arial,sans-serif;">Tool Adoption — week of ' +
    formatDate_(thisWeekStart) + '</h2>' +
    '<table cellpadding="8" cellspacing="0" style="border-collapse:collapse;font-family:Arial,sans-serif;font-size:14px;">' +
    '<tr style="background:#f3f4f6;"><th align="left">Tool</th><th>Days used this week</th><th>Avg days/week (prior 4)</th><th>Last opened</th></tr>';

  names.forEach(function(name) {
    var t = tools[name];
    var avg = (t.prior4 / 4).toFixed(1);
    var quiet = t.thisWeek === 0;
    html += '<tr style="border-bottom:1px solid #e5e7eb;' + (quiet ? 'color:#9ca3af;' : '') + '">' +
      '<td>' + escHtml_(name) + '</td>' +
      '<td align="center">' + (quiet ? '—' : t.thisWeek) + '</td>' +
      '<td align="center">' + avg + '</td>' +
      '<td align="center">' + t.lastSeen + '</td></tr>';
  });

  html += '</table>' +
    '<p style="font-family:Arial,sans-serif;font-size:12px;color:#6b7280;">' +
    'One ping per tool per day, logged when someone actually opens or uses the tool. ' +
    'A tool missing from this table has never pinged — check its ADOPTION_SHEET_ID property. ' +
    'Note: shared-table-email-summary delivers value via its daily email, so quiet is expected there.</p>';
  return html;
}

// Monday 00:00 of the week containing d
function weekStart_(d) {
  var x = new Date(d.getFullYear(), d.getMonth(), d.getDate());
  var day = x.getDay(); // 0=Sun
  x.setDate(x.getDate() - (day === 0 ? 6 : day - 1));
  return x;
}

function formatDate_(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'MMM d, yyyy');
}

function escHtml_(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

// Checks for an existing trigger before creating — safe to re-run.
function createWeeklyTrigger() {
  var exists = ScriptApp.getProjectTriggers().some(function(t) {
    return t.getHandlerFunction() === 'sendWeeklyAdoptionDigest';
  });
  if (exists) return;
  ScriptApp.newTrigger('sendWeeklyAdoptionDigest')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(6)
    .create();
}

function deleteTriggers() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'sendWeeklyAdoptionDigest') {
      ScriptApp.deleteTrigger(t);
    }
  });
}
