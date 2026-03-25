/**
 * Schedule Variance Analyzer — Orchestration & Config
 *
 * Reads all configuration from the Settings tab so directors can
 * change thresholds, recipients, and location names without touching code.
 */

/* ── Settings reader ── */

function getConfig() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Settings');
  if (!sheet) throw new Error('Missing "Settings" tab. Please create it with the required configuration rows.');

  var data = sheet.getDataRange().getValues();
  var map = {};
  data.forEach(function(row) {
    var key = String(row[0] || '').trim();
    var val = row[1];
    if (key) map[key] = val;
  });

  function str(k, def)  { return map[k] !== undefined && map[k] !== '' ? String(map[k]).trim() : def; }
  function num(k, def)  { var v = parseFloat(map[k]); return isNaN(v) ? def : v; }
  function list(k)      { var v = str(k, ''); return v ? v.split(',').map(function(s) { return s.trim(); }).filter(Boolean) : []; }

  return {
    LOC1_NAME:              str('Location 1 Name', 'Cockrell Hill DTO'),
    LOC2_NAME:              str('Location 2 Name', 'DBU OCV'),
    MONDAY_TAB:             str('Monday Tab Name', 'Current Week Schedule'),
    DIRECTOR_EMAILS:        list('Director Email Recipients'),
    MANAGER_EMAILS:         list('Manager Email Recipients'),
    OT_THRESHOLD:           num('OT Threshold (hours)', 40),
    HIGH_VARIANCE_THRESHOLD:num('High Variance Threshold (min)', 60),
    LATENESS_THRESHOLD:     num('Lateness Threshold (min)', 10),
    PATTERN_PERCENT:        num('Pattern Frequency (%)', 40) / 100,
    CHRONIC_WINDOW:         num('Chronic Flag Window (weeks)', 3),
    CHRONIC_TRIGGER:        num('Chronic Flag Trigger (weeks)', 2),
    SWAP_THRESHOLD:         num('Swap Detection Threshold (min)', 120),
    MIDNIGHT_WINDOW:        num('Midnight Window (min)', 5)
  };
}

/* ── Tab name constants ── */

var TABS = {
  INSTRUCTIONS:   'Instructions',
  CH_SCHED:       'CH Schedule Data',
  CH_PUNCH:       'CH Punch Data',
  DBU_SCHED:      'DBU Schedule Data',
  DBU_PUNCH:      'DBU Punch Data',
  CH_SCHED_NEXT:  'CH Schedule Data (This Week)',
  DBU_SCHED_NEXT: 'DBU Schedule Data (This Week)',
  REPORT:         'Weekly Report',
  HISTORY:        'History',
  SETTINGS:       'Settings',
  REPORT_OUTPUT:  'Report Output'
};

/* ── Menu ── */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Schedule Variance')
    .addItem('Upload & Analyze', 'showUploadDialog')
    .addItem('Re-run Analysis (from existing tabs)', 'analyze')
    .addSeparator()
    .addItem('Send Emails (Director + Manager)', 'sendEmails')
    .addItem('Fill Monday Sheet', 'fillMondaySheet')
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('Historical Reports')
      .addItem('Worst Offenders — Rolling', 'reportWorstOffenders')
      .addItem('Individual Employee Timeline', 'reportEmployeeTimeline')
      .addItem('OT Trend by Location', 'reportOTTrend')
      .addItem('Chronic Lateness Leaderboard', 'reportChronicLateness')
      .addItem('Missed Clock-Out Frequency', 'reportMissedClockOuts')
      .addItem('OT Reduction Trend', 'reportOTReduction'))
    .addSeparator()
    .addItem('Archive .ch Tabs Now', 'archiveWeeklySnapshot')
    .addItem('Setup Weekly Archive (Monday 5 PM)', 'setupWeeklyArchiveTrigger')
    .addItem('Remove Weekly Archive Trigger', 'removeWeeklyArchiveTrigger')
    .addSeparator()
    .addItem('Clear Input Tabs', 'clearInputs')
    .addItem('Initialize Sheet (first-time setup)', 'initializeSheet')
    .addToUi();
}

/* ── Upload Dialog ── */

function showUploadDialog() {
  var html = HtmlService.createHtmlOutputFromFile('Upload')
    .setWidth(580)
    .setHeight(780)
    .setTitle('Upload Schedule Data');
  SpreadsheetApp.getUi().showModalDialog(html, 'Upload Schedule Data');
}

/**
 * Pre-flight check: parses a date from the uploaded schedule CSV,
 * determines the week label, and checks History for existing rows.
 * Called from the dialog before running the full analysis.
 */
function checkExistingWeekData(schedCSV) {
  if (!schedCSV) return { exists: false };

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var historySheet = ss.getSheetByName(TABS.HISTORY);
  if (!historySheet || historySheet.getLastRow() < 2) return { exists: false };

  var rows = parseCSVToArray_(schedCSV);
  if (rows.length < 2) return { exists: false };

  var headers = rows[0].map(function(h) { return String(h).toUpperCase().replace(/[^A-Z_]/g, ''); });
  var iSS = headers.indexOf('SCHEDULED_START_TIME');
  if (iSS < 0) return { exists: false };

  var ts = parseTimestamp(rows[1][iSS]);
  if (!ts) return { exists: false };

  var weekLabel = makeWeekLabel(getWeekStart(formatDateKey(ts)));

  var histData = historySheet.getDataRange().getValues();
  var count = 0;
  for (var i = 1; i < histData.length; i++) {
    if (String(histData[i][1] || '').trim() === weekLabel) count++;
  }

  return { exists: count > 0, weekLabel: weekLabel, rowCount: count };
}

/**
 * Called from the Upload dialog. Receives raw CSV text for each slot,
 * writes it into the input tabs, then runs the full analysis.
 */
function receiveUploadAndAnalyze(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var mapping = [
    { key: 'chSched',     tab: TABS.CH_SCHED },
    { key: 'chPunch',     tab: TABS.CH_PUNCH },
    { key: 'dbuSched',    tab: TABS.DBU_SCHED },
    { key: 'dbuPunch',    tab: TABS.DBU_PUNCH },
    { key: 'chSchedNext', tab: TABS.CH_SCHED_NEXT },
    { key: 'dbuSchedNext',tab: TABS.DBU_SCHED_NEXT }
  ];

  mapping.forEach(function(m) {
    var csv = payload[m.key];
    if (!csv) return;

    var sheet = ss.getSheetByName(m.tab);
    if (!sheet) sheet = ss.insertSheet(m.tab);
    sheet.clearContents();

    var rows = parseCSVToArray_(csv);
    if (rows.length) {
      sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
    }
  });

  analyzeAfterUpload_();
  return { message: 'Analysis complete — see the Weekly Report tab.', hasRecipients: hasEmailRecipients_() };
}

function hasEmailRecipients_() {
  try {
    var cfg = getConfig();
    return cfg.DIRECTOR_EMAILS.length > 0 || cfg.MANAGER_EMAILS.length > 0;
  } catch (e) { return false; }
}

function sendEmailsAfterUpload() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cfg = getConfig();
  var latestRun = getLatestRunData_(ss);
  var chronicFlags = getChronicFlags(ss, cfg);
  sendTwoTierEmails(ss, latestRun.rows, latestRun.crossLocOT, chronicFlags, latestRun.locationData, latestRun.weeks, cfg);
  var count = cfg.DIRECTOR_EMAILS.length + cfg.MANAGER_EMAILS.length;
  return 'Emails sent to ' + count + ' recipient(s).';
}

/**
 * Parse raw CSV text into a 2D array suitable for setValues().
 * Handles quoted fields, commas inside quotes, and tab-delimited files.
 */
function parseCSVToArray_(text) {
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
      } else {
        field += ch;
      }
    } else {
      if (ch === '"') { inQ = true; }
      else if (ch === delim) { row.push(field.trim()); field = ''; }
      else if (ch === '\n' || ch === '\r') {
        if (ch === '\r' && i + 1 < text.length && text[i + 1] === '\n') i++;
        row.push(field.trim()); field = '';
        if (row.length > 1 || (row.length === 1 && row[0] !== '')) rows.push(row);
        row = [];
      } else {
        field += ch;
      }
    }
  }
  if (field || row.length > 0) {
    row.push(field.trim());
    if (row.length > 1 || (row.length === 1 && row[0] !== '')) rows.push(row);
  }

  // Normalize column count (setValues requires uniform width)
  var maxCols = 0;
  rows.forEach(function(r) { if (r.length > maxCols) maxCols = r.length; });
  rows.forEach(function(r) { while (r.length < maxCols) r.push(''); });

  return rows;
}

/**
 * Internal analysis runner called after upload writes data to tabs.
 * Same as analyze() but without UI prompts/alerts (runs silently
 * since the dialog handles user feedback).
 */
function analyzeAfterUpload_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cfg = getConfig();

  var loc1Sched = ss.getSheetByName(TABS.CH_SCHED);
  var loc1Punch = ss.getSheetByName(TABS.CH_PUNCH);
  var loc2Sched = ss.getSheetByName(TABS.DBU_SCHED);
  var loc2Punch = ss.getSheetByName(TABS.DBU_PUNCH);

  var hasLoc1 = loc1Sched && loc1Punch && loc1Sched.getLastRow() > 1 && loc1Punch.getLastRow() > 1;
  var hasLoc2 = loc2Sched && loc2Punch && loc2Sched.getLastRow() > 1 && loc2Punch.getLastRow() > 1;

  if (!hasLoc1 && !hasLoc2) {
    throw new Error('No valid data found in uploaded files. Check that CSVs have the correct columns.');
  }

  var allConsolidated = [];
  var locationData = [];

  if (hasLoc1) {
    var s1 = parseScheduleFromSheet(loc1Sched.getDataRange().getValues());
    var p1 = parsePunchesFromSheet(loc1Punch.getDataRange().getValues(), cfg);
    var c1 = consolidate(s1, p1, cfg);
    c1.forEach(function(r) { r.locationName = cfg.LOC1_NAME; });
    allConsolidated = allConsolidated.concat(c1);
    locationData.push({ name: cfg.LOC1_NAME, records: c1 });
  }

  if (hasLoc2) {
    var s2 = parseScheduleFromSheet(loc2Sched.getDataRange().getValues());
    var p2 = parsePunchesFromSheet(loc2Punch.getDataRange().getValues(), cfg);
    var c2 = consolidate(s2, p2, cfg);
    c2.forEach(function(r) { r.locationName = cfg.LOC2_NAME; });
    allConsolidated = allConsolidated.concat(c2);
    locationData.push({ name: cfg.LOC2_NAME, records: c2 });
  }

  var weeks = detectWeeks(allConsolidated);
  var allRows = [];

  weeks.forEach(function(w) {
    w.records.forEach(function(r) { r.weekLabel = w.label; });
    var rolled = rollUp(w.records, cfg);
    rolled.forEach(function(row) { row.weekLabel = w.label; });
    allRows = allRows.concat(rolled);
  });

  var crossLocOT = reconcileCrossLocationOT(allConsolidated, cfg);
  var chronicFlags = getChronicFlags(ss, cfg);

  writeReport(ss, allRows, weeks, locationData, crossLocOT, chronicFlags, cfg);
  appendHistory(ss, allRows, crossLocOT);
}

/* ── First-time setup: creates all tabs and populates Settings + Instructions ── */

function initializeSheet() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var tabNames = [
    TABS.INSTRUCTIONS, TABS.CH_SCHED, TABS.CH_PUNCH,
    TABS.DBU_SCHED, TABS.DBU_PUNCH,
    TABS.CH_SCHED_NEXT, TABS.DBU_SCHED_NEXT,
    TABS.REPORT, TABS.HISTORY, TABS.SETTINGS
  ];

  tabNames.forEach(function(name) {
    if (!ss.getSheetByName(name)) ss.insertSheet(name);
  });

  var settingsSheet = ss.getSheetByName(TABS.SETTINGS);
  if (settingsSheet.getLastRow() < 2) {
    buildSettingsTab_(ss);
  }

  var instrSheet = ss.getSheetByName(TABS.INSTRUCTIONS);
  if (instrSheet.getLastRow() < 2) {
    buildInstructionsTab_(ss);
  }

  protectHistoryTab_(ss);

  ui.alert('Setup Complete', 'All tabs created. Settings and Instructions were only written if they were empty — your existing data is untouched.', ui.ButtonSet.OK);
}

function buildSettingsTab_(ss) {
  var sheet = ss.getSheetByName(TABS.SETTINGS);

  var settings = [
    ['Setting', 'Value', 'Description'],
    ['Location 1 Name', 'Cockrell Hill DTO', 'Display name for location 1 in reports and emails'],
    ['Location 2 Name', 'DBU OCV', 'Display name for location 2 in reports and emails'],
    ['Monday Tab Name', 'Current Week Schedule', 'Name of the tab in this sheet where the Monday Huddle OT block lives'],
    ['Director Email Recipients', '', 'Comma-separated emails for full-detail tier'],
    ['Manager Email Recipients', '', 'Comma-separated emails for summary tier'],
    ['OT Threshold (hours)', 40, 'Weekly hours above which an employee is flagged for OT'],
    ['High Variance Threshold (min)', 60, 'Total weekly variance above which an employee is flagged'],
    ['Lateness Threshold (min)', 10, 'Minimum minutes late to count toward lateness pattern'],
    ['Pattern Frequency (%)', 40, '% of shifts that must show a behavior to trigger a pattern flag'],
    ['Chronic Flag Window (weeks)', 3, 'How many weeks back to look when evaluating chronic behavior'],
    ['Chronic Flag Trigger (weeks)', 2, 'How many flagged weeks in the window triggers a chronic alert'],
    ['Swap Detection Threshold (min)', 120, 'Start variance above which a shift is flagged as possible swap'],
    ['Midnight Window (min)', 5, 'Minutes from midnight to flag as likely missed clock-out']
  ];

  sheet.getRange(1, 1, settings.length, 3).setValues(settings);
  sheet.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#F3F4F6');
  sheet.setColumnWidth(1, 280);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 420);
  sheet.setFrozenRows(1);
}

function buildInstructionsTab_(ss) {
  var sheet = ss.getSheetByName(TABS.INSTRUCTIONS);
  sheet.clearContents();
  sheet.clearFormats();

  var r = 1;
  sheet.getRange(r, 1).setValue('SCHEDULE VARIANCE ANALYZER — INSTRUCTIONS').setFontSize(16).setFontWeight('bold');
  r += 2;

  var sections = [
    ['WHAT THIS TOOL DOES', [
      'This tool compares scheduled hours (from HotSchedules) against actual time punches to identify:',
      '• Late arrivals and early departures',
      '• Missed clock-outs (punched out at midnight)',
      '• Absences (scheduled but no punch)',
      '• Overtime, including cross-location OT that Analytics Hub misses',
      '• Behavioral patterns (chronic lateness, works extra, etc.)',
      '',
      'It also automatically archives your Monday Huddle presentation tabs every week,',
      'building a permanent running history of what was discussed each Monday.',
      '',
      'It replaces the manual 30–60 minute weekly prep with a 3–5 minute workflow.'
    ]],
    ['WEEKLY WORKFLOW (STEP BY STEP)', [
      '1. Go to CFA Home → Analytics Hub → Data Feeds',
      '2. Download CSV files:',
      '   LAST WEEK (required — at least one location):',
      '   • Cockrell Hill: "HotSchedules Schedule Data" + "Time Punches"',
      '   • DBU: "HotSchedules Schedule Data" + "Time Punches"',
      '   THIS WEEK (optional — schedule only, no punches):',
      '   • Cockrell Hill: "HotSchedules Schedule Data"',
      '   • DBU: "HotSchedules Schedule Data"',
      '3. Click "Schedule Variance" menu → "Upload & Analyze"',
      '4. Drag and drop (or click) to upload each file into its box',
      '5. Click "Analyze Variance" — the tool writes the data, runs analysis, and sends emails',
      '6. Review the "Weekly Report" tab',
      '7. (Optional) Click "Fill Monday Sheet" to push OT data to the Monday Huddle tab',
      '',
      'Note: You only need one location\'s last-week data to run.',
      'This-week schedule files are optional — they power the "This Week Plan" on the Monday Sheet.'
    ]],
    ['UNDERSTANDING THE REPORT', [
      'GREEN = Good (on track, high reliability)',
      'YELLOW = Warning (stayed late, missed clock-out, possible swap)',
      'RED = Action needed (absent, chronic lateness, low reliability, high variance)',
      '',
      'Reliability % = (matched days / scheduled days) × 100',
      '  Green ≥ 90%  |  Yellow ≥ 75%  |  Red < 75%',
      '',
      'Variance = difference between scheduled and actual time in minutes',
      '  Positive (+) means more time worked than scheduled',
      '  Negative (−) means less time worked than scheduled'
    ]],
    ['WEEKLY ARCHIVE (.ch TABS)', [
      'Every Monday at 5:00 PM, all tabs ending in ".ch" are automatically copied into the',
      '"Huddle Archive" tab. This creates a permanent, append-only history of your Monday',
      'Huddle presentations — nothing is ever deleted.',
      '',
      'Archived tabs:',
      '• Catering Presentation.ch',
      '• Current Week Schedule.ch',
      '• Talent Presentation.ch',
      '• Facilities Presentation.ch',
      '• CEM Presentation.ch',
      '• Any future tab you create ending in ".ch"',
      '',
      'Each snapshot includes a date/time banner and preserves formatting and merged cells.',
      '',
      'From the "Schedule Variance" menu you can:',
      '• "Archive .ch Tabs Now" — take a snapshot immediately (for testing or ad-hoc use)',
      '• "Setup Weekly Archive (Monday 5 PM)" — one-time setup to enable the automatic trigger',
      '• "Remove Weekly Archive Trigger" — turn off automatic archiving'
    ]],
    ['TABS IN THIS SHEET', [
      'Instructions — You are here',
      '',
      'Input tabs (auto-populated via upload):',
      '  CH Schedule Data — Last week Cockrell Hill schedule',
      '  CH Punch Data — Last week Cockrell Hill time punches',
      '  DBU Schedule Data — Last week DBU schedule',
      '  DBU Punch Data — Last week DBU time punches',
      '  CH Schedule Data (This Week) — Upcoming week Cockrell Hill schedule',
      '  DBU Schedule Data (This Week) — Upcoming week DBU schedule',
      '',
      'Output tabs (auto-generated, do not edit):',
      '  Weekly Report — Analysis results with color-coded flags',
      '  History — Backend data store for chronic flags and historical reports (locked)',
      '  Huddle Archive — Permanent weekly snapshots of all .ch presentation tabs',
      '',
      'Presentation tabs (your content, archived automatically):',
      '  Catering Presentation.ch',
      '  Current Week Schedule.ch',
      '  Talent Presentation.ch',
      '  Facilities Presentation.ch',
      '  CEM Presentation.ch',
      '  Follow up things to do',
      '',
      'Configuration:',
      '  Settings — All configurable values (emails, thresholds, etc.)'
    ]],
    ['SETTINGS YOU CAN CHANGE', [
      'Open the "Settings" tab to modify any of these without touching code:',
      '• Location names — change if locations change',
      '• Email recipients — comma-separated list for Director and Manager tiers',
      '• OT Threshold — default 40 hours',
      '• Variance and lateness thresholds',
      '• Monday Tab Name — the tab in this sheet where the OT block gets written (default: "Current Week Schedule")',
      '• Chronic Flag Window / Trigger — how many weeks of repeated issues trigger a chronic alert',
      '• Swap Detection Threshold — minutes of start variance that flag a possible shift swap',
      '• Midnight Window — minutes from midnight to flag as likely missed clock-out'
    ]],
    ['HISTORICAL REPORTS', [
      'Under Schedule Variance → Historical Reports:',
      '• Worst Offenders — Rolling: lowest reliability over 4/8/12 weeks',
      '• Individual Employee Timeline: full history for one person',
      '• OT Trend by Location: week-over-week OT comparison',
      '• Chronic Lateness Leaderboard: who\'s late the most over 4 weeks',
      '• Missed Clock-Out Frequency: by employee',
      '• OT Reduction Trend: month-over-month OT trajectory'
    ]],
    ['TROUBLESHOOTING', [
      'Q: "No data found" error — Make sure you uploaded the correct CSV files with the expected columns.',
      'Q: Employees don\'t match — Names must match exactly between schedule and punch exports.',
      'Q: OT numbers look wrong — Run both locations in the same analysis. Cross-location OT is only calculated when both are present.',
      'Q: Email not sending — Check that email addresses are entered correctly in Settings. The script uses your Google account to send.',
      'Q: Archive is empty — Run "Setup Weekly Archive (Monday 5 PM)" once from the menu to enable the trigger.',
      'Q: Want to archive right now — Click "Archive .ch Tabs Now" from the Schedule Variance menu.',
      'Q: New presentation tab not archiving — Make sure its name ends in ".ch" (e.g., "Safety Presentation.ch").'
    ]]
  ];

  sections.forEach(function(sec) {
    sheet.getRange(r, 1).setValue(sec[0]).setFontSize(13).setFontWeight('bold').setFontColor('#1a1a2e');
    r++;
    sec[1].forEach(function(line) {
      sheet.getRange(r, 1).setValue(line).setFontSize(10).setFontColor(line === '' ? '#ffffff' : '#555555');
      r++;
    });
    r++;
  });

  sheet.setColumnWidth(1, 800);
}

function protectHistoryTab_(ss) {
  var sheet = ss.getSheetByName(TABS.HISTORY);
  if (!sheet) return;
  var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  if (protections.length === 0) {
    var protection = sheet.protect().setDescription('History — automated data only');
    protection.setWarningOnly(true);
  }
}

/* ── Main Analysis ── */

function analyze() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cfg;

  try {
    cfg = getConfig();
  } catch (e) {
    ui.alert('Configuration Error', e.message + '\n\nRun "Initialize Sheet" from the menu first.', ui.ButtonSet.OK);
    return;
  }

  var loc1Sched = ss.getSheetByName(TABS.CH_SCHED);
  var loc1Punch = ss.getSheetByName(TABS.CH_PUNCH);
  var loc2Sched = ss.getSheetByName(TABS.DBU_SCHED);
  var loc2Punch = ss.getSheetByName(TABS.DBU_PUNCH);

  var hasLoc1 = loc1Sched && loc1Punch && loc1Sched.getLastRow() > 1 && loc1Punch.getLastRow() > 1;
  var hasLoc2 = loc2Sched && loc2Punch && loc2Sched.getLastRow() > 1 && loc2Punch.getLastRow() > 1;

  if (!hasLoc1 && !hasLoc2) {
    ui.alert('No Data', 'No data in the input tabs. Use "Upload & Analyze" to upload your CSV files, or paste data into the input tabs manually.', ui.ButtonSet.OK);
    return;
  }

  try {
    var allConsolidated = [];
    var locationData = [];

    if (hasLoc1) {
      var s1 = parseScheduleFromSheet(loc1Sched.getDataRange().getValues());
      var p1 = parsePunchesFromSheet(loc1Punch.getDataRange().getValues(), cfg);
      var c1 = consolidate(s1, p1, cfg);
      c1.forEach(function(r) { r.locationName = cfg.LOC1_NAME; });
      allConsolidated = allConsolidated.concat(c1);
      locationData.push({ name: cfg.LOC1_NAME, records: c1 });
    }

    if (hasLoc2) {
      var s2 = parseScheduleFromSheet(loc2Sched.getDataRange().getValues());
      var p2 = parsePunchesFromSheet(loc2Punch.getDataRange().getValues(), cfg);
      var c2 = consolidate(s2, p2, cfg);
      c2.forEach(function(r) { r.locationName = cfg.LOC2_NAME; });
      allConsolidated = allConsolidated.concat(c2);
      locationData.push({ name: cfg.LOC2_NAME, records: c2 });
    }

    var weeks = detectWeeks(allConsolidated);
    var allRows = [];

    weeks.forEach(function(w) {
      w.records.forEach(function(r) { r.weekLabel = w.label; });
      var rolled = rollUp(w.records, cfg);
      rolled.forEach(function(row) { row.weekLabel = w.label; });
      allRows = allRows.concat(rolled);
    });

    var crossLocOT = reconcileCrossLocationOT(allConsolidated, cfg);

    var chronicFlags = getChronicFlags(ss, cfg);

    writeReport(ss, allRows, weeks, locationData, crossLocOT, chronicFlags, cfg);
    appendHistory(ss, allRows, crossLocOT);

    var totalDays = allConsolidated.length;
    var matched = allConsolidated.filter(function(r) { return r.status === 'matched'; }).length;
    var locNames = locationData.map(function(l) { return l.name; }).join(' + ');

    ui.alert(
      'Analysis Complete',
      locNames + '\n' +
      totalDays + ' day-records processed across ' + weeks.length + ' week(s).\n' +
      matched + ' matched, ' + (totalDays - matched) + ' unmatched.\n\n' +
      'See the "' + TABS.REPORT + '" tab for results.',
      ui.ButtonSet.OK
    );

    if (cfg.DIRECTOR_EMAILS.length || cfg.MANAGER_EMAILS.length) {
      var sendConfirm = ui.alert(
        'Send Emails?',
        'Send variance report emails to ' + (cfg.DIRECTOR_EMAILS.length + cfg.MANAGER_EMAILS.length) + ' recipient(s)?',
        ui.ButtonSet.YES_NO
      );
      if (sendConfirm === ui.Button.YES) {
        sendTwoTierEmails(ss, allRows, crossLocOT, chronicFlags, locationData, weeks, cfg);
        ui.alert('Emails sent.');
      }
    }
  } catch (e) {
    ui.alert('Error during analysis:\n' + e.message + '\n\n' + e.stack);
  }
}

/* ── Clear Inputs ── */

function clearInputs() {
  var ui = SpreadsheetApp.getUi();
  var confirm = ui.alert(
    'Clear Input Tabs',
    'This will erase all data from the six input tabs (last-week and this-week schedule/punch data). Continue?',
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  [TABS.CH_SCHED, TABS.CH_PUNCH, TABS.DBU_SCHED, TABS.DBU_PUNCH,
   TABS.CH_SCHED_NEXT, TABS.DBU_SCHED_NEXT].forEach(function(name) {
    var s = ss.getSheetByName(name);
    if (s) s.clearContents();
  });
  ui.alert('All six input tabs cleared.');
}

/* ── Email trigger (manual, if auto-send was skipped) ── */

function sendEmails() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cfg;

  try {
    cfg = getConfig();
  } catch (e) {
    ui.alert('Configuration Error', e.message, ui.ButtonSet.OK);
    return;
  }

  if (!cfg.DIRECTOR_EMAILS.length && !cfg.MANAGER_EMAILS.length) {
    ui.alert('No email recipients configured in the Settings tab.');
    return;
  }

  var historySheet = ss.getSheetByName(TABS.HISTORY);
  if (!historySheet || historySheet.getLastRow() < 2) {
    ui.alert('No analysis data. Run "Run Full Analysis" first.');
    return;
  }

  try {
    var latestRun = getLatestRunData_(ss);
    var chronicFlags = getChronicFlags(ss, cfg);
    sendTwoTierEmails(ss, latestRun.rows, latestRun.crossLocOT, chronicFlags, latestRun.locationData, latestRun.weeks, cfg);
    ui.alert('Emails sent to ' + (cfg.DIRECTOR_EMAILS.length + cfg.MANAGER_EMAILS.length) + ' recipient(s).');
  } catch (e) {
    ui.alert('Email error: ' + e.message);
  }
}

/**
 * Reconstruct the latest run's data from History for re-sending emails.
 */
function getLatestRunData_(ss) {
  var historySheet = ss.getSheetByName(TABS.HISTORY);
  var data = historySheet.getDataRange().getValues();
  var headers = data[0];
  var latestRunDate = data[data.length - 1][0];

  var rows = [];
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(latestRunDate)) {
      rows.push({
        name: data[i][3],
        locationName: data[i][2],
        weekLabel: data[i][1],
        scheduledDays: data[i][4],
        workedDays: data[i][5],
        absentDays: data[i][6],
        unscheduledDays: data[i][7],
        reliability: data[i][8],
        startVar: data[i][9],
        endVar: data[i][10],
        totalVar: data[i][11],
        avgVar: data[i][12],
        earlyIn: data[i][13],
        lateIn: data[i][14],
        lateOut: data[i][15],
        earlyOut: data[i][16],
        midnightCount: data[i][17],
        pattern: { label: data[i][18], type: patternTypeFromLabel_(data[i][18]) },
        swapCount: data[i][19],
        otHours: data[i][20],
        lateCount: data[i][21],
        absTotal: Math.abs(data[i][11] || 0)
      });
    }
  }

  var locSet = {};
  var weekSet = {};
  rows.forEach(function(r) {
    locSet[r.locationName] = true;
    weekSet[r.weekLabel] = true;
  });

  return {
    rows: rows,
    crossLocOT: {},
    locationData: Object.keys(locSet).map(function(n) { return { name: n }; }),
    weeks: Object.keys(weekSet).map(function(l) { return { label: l }; })
  };
}

function patternTypeFromLabel_(label) {
  if (!label) return 'none';
  if (label === 'On Track') return 'good';
  if (label === 'Late In' || label === 'Works Less' || label === 'Leaves Early') return 'bad';
  if (label === 'Works Extra' || label === 'Stays Late' || label === 'Shift Late') return 'warn';
  if (label === 'Early In' || label === 'Shift Early') return 'neutral';
  return 'none';
}
