/**
 * Schedule Variance Analyzer — Orchestration & Config
 *
 * Reads all configuration from the Settings tab so directors can
 * change thresholds, recipients, and location names without touching code.
 */

/* ── Trigger failure notifier ──
 *
 * Wrap any trigger-invoked function in try/catch and call this in the
 * catch block. Emails the configured alert recipient (or the script
 * owner as fallback) so silent failures in background triggers
 * actually get noticed.
 */
function notifyTriggerFailure_(functionName, error) {
  var recipient = '';
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var settingsSheet = ss.getSheetByName('Settings');
    if (settingsSheet && settingsSheet.getLastRow() > 1) {
      var data = settingsSheet.getRange(2, 1, settingsSheet.getLastRow() - 1, 2).getValues();
      for (var i = 0; i < data.length; i++) {
        if (String(data[i][0] || '').trim() === 'Trigger Alert Email') {
          recipient = String(data[i][1] || '').trim();
          break;
        }
      }
    }
  } catch (e) {
    // Ignore — fall through to owner fallback
  }

  if (!recipient) {
    try {
      recipient = Session.getEffectiveUser().getEmail();
    } catch (e) {
      Logger.log('Trigger failure (no recipient available): ' + functionName + ' — ' + error);
      return;
    }
  }
  if (!recipient) {
    Logger.log('Trigger failure (empty recipient): ' + functionName + ' — ' + error);
    return;
  }

  var ssName = 'Schedule Variance Analyzer';
  var ssUrl = '';
  try {
    var ss2 = SpreadsheetApp.getActiveSpreadsheet();
    ssName = ss2.getName();
    ssUrl = ss2.getUrl();
  } catch (e) {}

  var subject = '⚠ Trigger Failure: ' + functionName + ' — ' + ssName;
  var body =
    'A scheduled background function failed.\n\n' +
    'Spreadsheet: ' + ssName + '\n' +
    (ssUrl ? 'Link: ' + ssUrl + '\n' : '') +
    'Function:    ' + functionName + '\n' +
    'Time:        ' + new Date() + '\n\n' +
    'Error:\n' + (error && error.stack ? error.stack : String(error)) + '\n\n' +
    'This alert comes from the Schedule Variance Analyzer trigger monitor.\n' +
    'Change the recipient in Settings → "Trigger Alert Email".';

  try {
    MailApp.sendEmail({ to: recipient, subject: subject, body: body });
  } catch (e) {
    Logger.log('Failed to send trigger failure email: ' + e);
  }
}

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
    MONDAY_TAB:             str('Monday Tab Name', 'Current Week Schedule.ch'),
    DIRECTOR_EMAILS:        list('Director Email Recipients'),
    MANAGER_EMAILS:         list('Manager Email Recipients'),
    OT_THRESHOLD:           num('OT Threshold (hours)', 40),
    HIGH_VARIANCE_THRESHOLD:num('High Variance Threshold (min)', 60),
    LATENESS_THRESHOLD:     num('Lateness Threshold (min)', 10),
    PATTERN_PERCENT:        num('Pattern Frequency (%)', 40) / 100,
    CHRONIC_WINDOW:         num('Chronic Flag Window (weeks)', 3),
    CHRONIC_TRIGGER:        num('Chronic Flag Trigger (weeks)', 2),
    SWAP_THRESHOLD:         num('Swap Detection Threshold (min)', 120),
    MIDNIGHT_WINDOW:        num('Midnight Window (min)', 5),
    MISSED_CLOCKOUT_THRESHOLD: num('Missed Clock-Out Threshold (min)', 90),
    TRIGGER_ALERT_EMAIL:    str('Trigger Alert Email', '')
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
  REPORT_OUTPUT:  'Report Output',
  NAME_ALIASES:   'Name Aliases'
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
    .addItem('Rebuild Instructions Tab', 'rebuildInstructionsTab')
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
    TABS.REPORT, TABS.HISTORY, TABS.SETTINGS,
    TABS.NAME_ALIASES
  ];

  tabNames.forEach(function(name) {
    if (!ss.getSheetByName(name)) ss.insertSheet(name);
  });

  var aliasSheet = ss.getSheetByName(TABS.NAME_ALIASES);
  if (aliasSheet && aliasSheet.getLastRow() < 1) {
    buildNameAliasesTab_(ss);
  }

  var settingsSheet = ss.getSheetByName(TABS.SETTINGS);
  var settingsMsg;
  if (settingsSheet.getLastRow() < 2) {
    buildSettingsTab_(ss);
    settingsMsg = 'Settings tab populated with defaults.';
  } else {
    var added = migrateSettingsTab_(ss);
    settingsMsg = added > 0
      ? 'Settings tab migrated: added ' + added + ' new setting row(s).'
      : 'Settings tab already up to date.';
  }

  var instrSheet = ss.getSheetByName(TABS.INSTRUCTIONS);
  if (instrSheet.getLastRow() < 2) {
    buildInstructionsTab_(ss);
  }

  protectHistoryTab_(ss);

  ui.alert('Setup Complete', 'All tabs created.\n\n' + settingsMsg + '\n\nExisting data was not touched.', ui.ButtonSet.OK);
}

/**
 * Canonical list of every Settings row the tool expects. Single source of
 * truth — buildSettingsTab_() uses this for fresh installs, and
 * migrateSettingsTab_() uses it to append missing rows on existing installs.
 */
function getSettingsSchema_() {
  return [
    ['Location 1 Name', 'Cockrell Hill DTO', 'Display name for location 1 in reports and emails'],
    ['Location 2 Name', 'DBU OCV', 'Display name for location 2 in reports and emails'],
    ['Monday Tab Name', 'Current Week Schedule.ch', 'Name of the tab in this sheet where the Monday Huddle OT block lives'],
    ['Director Email Recipients', '', 'Comma-separated emails for full-detail tier'],
    ['Manager Email Recipients', '', 'Comma-separated emails for summary tier'],
    ['Trigger Alert Email', '', 'Email address to notify when a background trigger (e.g. weekly archive) fails. Leave blank to use the spreadsheet owner.'],
    ['OT Threshold (hours)', 40, 'Weekly hours above which an employee is flagged for OT'],
    ['High Variance Threshold (min)', 60, 'Total weekly variance above which an employee is flagged'],
    ['Lateness Threshold (min)', 10, 'Minimum minutes late to count toward lateness pattern'],
    ['Pattern Frequency (%)', 40, '% of shifts that must show a behavior to trigger a pattern flag'],
    ['Chronic Flag Window (weeks)', 3, 'How many weeks back to look when evaluating chronic behavior'],
    ['Chronic Flag Trigger (weeks)', 2, 'How many flagged weeks in the window triggers a chronic alert'],
    ['Swap Detection Threshold (min)', 120, 'Start variance above which a shift is flagged as possible swap'],
    ['Midnight Window (min)', 5, 'Fallback only — minutes from midnight to flag unscheduled shifts as likely missed clock-out'],
    ['Missed Clock-Out Threshold (min)', 90, 'Minutes past scheduled end to flag a matched shift as a missed clock-out']
  ];
}

function buildSettingsTab_(ss) {
  var sheet = ss.getSheetByName(TABS.SETTINGS);
  var rows = [['Setting', 'Value', 'Description']].concat(getSettingsSchema_());

  sheet.getRange(1, 1, rows.length, 3).setValues(rows);
  sheet.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#F3F4F6');
  sheet.setColumnWidth(1, 280);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 420);
  sheet.setFrozenRows(1);
}

/**
 * Append any missing setting rows to an existing Settings tab without
 * touching rows that already exist (including user-edited values).
 * Returns the number of rows added.
 */
function migrateSettingsTab_(ss) {
  var sheet = ss.getSheetByName(TABS.SETTINGS);
  if (!sheet) return 0;
  if (sheet.getLastRow() < 2) return 0; // empty — buildSettingsTab_ handles this case

  var existing = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  var existingKeys = {};
  existing.forEach(function(row) {
    var k = String(row[0] || '').trim();
    if (k) existingKeys[k] = true;
  });

  var schema = getSettingsSchema_();
  var toAdd = schema.filter(function(row) { return !existingKeys[row[0]]; });
  if (!toAdd.length) return 0;

  var startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, toAdd.length, 3).setValues(toAdd);
  return toAdd.length;
}

/**
 * Builds the Instructions tab — a standalone operations manual so any
 * non-technical successor can run, maintain, and troubleshoot this
 * spreadsheet without needing a phone call.
 *
 * Line prefix convention inside each section's body array:
 *   '### '  → sub-heading (bold, size 11)
 *   anything else → body line (size 10)
 *   empty string → blank spacer row
 */
function buildInstructionsTab_(ss) {
  var sheet = ss.getSheetByName(TABS.INSTRUCTIONS);
  sheet.clearContents();
  sheet.clearFormats();

  var sections = [
    ['1. TABLE OF CONTENTS', [
      'This tab is the full operating manual for the Schedule Variance Analyzer.',
      'Scroll down to the section you need:',
      '',
      '  1. Table of Contents (you are here)',
      '  2. What This Tool Does & Why It Exists',
      '  3. Who Uses This Tool',
      '  4. First-Time Setup',
      '  5. Weekly Workflow (Step by Step)',
      '  6. Understanding the Weekly Report',
      '  7. What a "Missed Clock-Out" Means',
      '  8. Overtime — How It\'s Calculated',
      '  9. The Two Email Tiers',
      ' 10. The Monday Huddle Sheet',
      ' 11. The Weekly Archive (.ch Tabs)',
      ' 12. Tab Reference',
      ' 13. Settings — Full Reference',
      ' 14. Historical Reports',
      ' 15. Menu Reference',
      ' 16. Glossary',
      ' 17. Common Operator Mistakes',
      ' 18. Troubleshooting',
      ' 19. If You\'re Stuck'
    ]],

    ['2. WHAT THIS TOOL DOES & WHY IT EXISTS', [
      'The Schedule Variance Analyzer powers the Monday Morning Huddle with two views:',
      '',
      '  LAST WEEK REVIEW — compares last week\'s scheduled hours against actual time',
      '  punches. Flags variance, lateness, absences, missed clock-outs, OT, and',
      '  behavior patterns. This is the accountability piece: what actually happened.',
      '',
      '  THIS WEEK PLAN — shows this week\'s projected scheduled hours and OT. This',
      '  is the planning piece: what are we heading into, and do we need to cut hours',
      '  before the week plays out.',
      '',
      'It also archives every Monday Huddle presentation tab automatically.',
      '',
      '### What it finds (last week)',
      '• Late arrivals and early departures',
      '• Missed clock-outs (shifts where someone forgot to punch out)',
      '• Absences (scheduled but no punch at all)',
      '• Overtime — including CROSS-LOCATION OT that Analytics Hub does not show',
      '• Behavior patterns (chronic lateness, works extra, shift swaps)',
      '',
      '### What it projects (this week)',
      '• Total scheduled hours per location',
      '• Scheduled OT — if the schedule as written would put people over 40 hours',
      '',
      '### Why it exists',
      'Before this tool, weekly prep for the Monday Huddle took 30–60 minutes of manual',
      'Analytics Hub pulls, Excel work, and spreadsheet shuffling. This tool replaces',
      'that with a 3–5 minute workflow: upload six CSVs, click Analyze, done.',
      '',
      'The biggest single reason this exists: Analytics Hub does NOT combine hours worked',
      'across both locations when an employee works at both Cockrell Hill AND DBU. That',
      'meant cross-location OT was invisible. This tool fixes that.',
      '',
      '### What it is NOT',
      '• It is NOT a timekeeping system of record. HotSchedules is.',
      '• It does NOT edit punches or schedules in HotSchedules.',
      '• It does NOT pay people. Payroll runs from the corporate system.',
      '• Small differences from payroll numbers are expected and normal.'
    ]],

    ['3. WHO USES THIS TOOL', [
      '### Chief of Staff / Operator (primary user)',
      'Runs the weekly workflow every Monday morning. Uploads the CSVs, reviews the',
      'Weekly Report, sends the emails, fills the Monday Sheet, and handles exceptions.',
      'This is the person reading these instructions right now.',
      '',
      '### Directors',
      'Receive the full-detail email (every flagged employee, every alert, every date).',
      'Review flagged employees and coach managers. Do not need to open the spreadsheet.',
      '',
      '### Managers',
      'Receive the summary email (counts and top-line alerts only). Handle day-to-day',
      'corrections and coach individual team members. Do not need to open the spreadsheet.',
      '',
      'Editing the spreadsheet directly is almost never needed. Everything you need to',
      'run or configure lives in the "Schedule Variance" menu and the "Settings" tab.'
    ]],

    ['4. FIRST-TIME SETUP', [
      'Do these steps ONCE when you inherit or copy this spreadsheet. After that, you',
      'just run the Weekly Workflow (section 5) each Monday.',
      '',
      '### Step 1 — Initialize',
      'Click Schedule Variance → "Initialize Sheet (first-time setup)". This creates',
      'all the tabs it needs (Instructions, Settings, History, etc.) and writes default',
      'values into Settings if they are empty. Safe to re-run — it will NOT overwrite',
      'anything that already has data.',
      '',
      '### Step 2 — Fill in Settings',
      'Open the "Settings" tab. Fill in at minimum:',
      '  • Director Email Recipients — comma-separated list of director emails',
      '  • Manager Email Recipients — comma-separated list of manager emails',
      '  • Trigger Alert Email — YOUR email. If a background trigger fails (like',
      '    the Monday 5 PM archive), this is where the alert will go. Leave blank',
      '    to default to the spreadsheet owner.',
      'You can leave every other setting at its default to start.',
      '',
      '### Step 3 — Enable the weekly archive',
      'Click Schedule Variance → "Setup Weekly Archive (Monday 5 PM)". This creates a',
      'time-driven trigger that automatically snapshots your Monday Huddle presentation',
      'tabs every Monday at 5:00 PM. You only need to run this once.',
      '',
      '### Step 4 — Test run',
      'Do a practice upload with a known past week of data. Run Upload & Analyze, check',
      'that the Weekly Report looks right, and send the emails to yourself only (temporarily',
      'edit the recipient list in Settings) before going live with real recipients.'
    ]],

    ['5. WEEKLY WORKFLOW (STEP BY STEP)', [
      'Target time: 3–5 minutes from start to finish.',
      '',
      '### Step 1 — Download CSVs from Analytics Hub',
      'Go to CFA Home → Analytics Hub → Data Feeds. Download the following:',
      '',
      'LAST WEEK (schedule + punches — powers the "Last Week Review" section):',
      '  • Cockrell Hill: "HotSchedules Schedule Data" + "Time Punches"',
      '  • DBU: "HotSchedules Schedule Data" + "Time Punches"',
      '',
      'THIS WEEK (schedule only — powers the "This Week Plan" section):',
      '  • Cockrell Hill: "HotSchedules Schedule Data"',
      '  • DBU: "HotSchedules Schedule Data"',
      '  (No punches for this week — the week just started.)',
      '',
      'You should upload ALL SIX files every week: both locations\' schedule + punches',
      'for last week, plus both locations\' schedules for this week. The Monday Huddle',
      'sheet uses all of it — last week for review, this week for planning.',
      '',
      '### Step 2 — Open the upload dialog',
      'Click Schedule Variance → "Upload & Analyze". A dialog box opens with six slots:',
      'four for last-week data and two for this-week schedules.',
      '',
      '### Step 3 — Upload each file',
      'Drag and drop each CSV into its matching slot, or click the slot to browse.',
      'Minimum required: at least one location\'s last-week schedule AND punch files.',
      'But for the full Monday Huddle experience (last week review + this week plan),',
      'upload all six files.',
      '',
      '### Step 4 — Click "Analyze Variance"',
      'The tool writes the data into the input tabs, runs the analysis, updates the',
      'Weekly Report and History tabs, and sends the director + manager emails. This',
      'normally takes about 30 seconds. Do not close the dialog until it says "Done."',
      '',
      '### Step 5 — Review the Weekly Report tab',
      'Open the "Weekly Report" tab and scan the Flags & Alerts section at the top',
      'for anything unusual. See section 6 for how to read every column.',
      '',
      '### Step 6 — Fill the Monday Sheet',
      'Click Schedule Variance → "Fill Monday Sheet". This pushes the OT block onto',
      'the tab named in Settings as "Monday Tab Name" (default: Current Week Schedule.ch).',
      'Safe to re-run — it overwrites the same block rather than duplicating.',
      '',
      '### If the week has already been analyzed',
      'The upload dialog checks History before running and warns you if it finds',
      'existing rows for that week. You can:',
      '  • Cancel and not re-run (if you already did it earlier)',
      '  • Proceed and overwrite (if you need to re-run with corrected data)',
      '',
      '### If you only have one location',
      'You can run with just one location\'s data. Cross-location OT will not be',
      'calculated (because the tool cannot see the other location), but everything',
      'else works normally.'
    ]],

    ['6. UNDERSTANDING THE WEEKLY REPORT', [
      'The Weekly Report tab has three main areas: Flags & Alerts (top), Employee',
      'Roll-Up (middle), and Day-by-Day Detail (bottom).',
      '',
      '### Color legend',
      'GREEN   = Good. On track, high reliability, no action needed.',
      'YELLOW  = Warning. Stayed late, missed clock-out, possible shift swap.',
      'RED     = Action needed. Absent, chronic lateness, low reliability, high variance.',
      '',
      'Reliability thresholds: Green ≥ 90%  |  Yellow ≥ 75%  |  Red < 75%',
      '',
      '### Variance — what the numbers mean',
      'Variance is measured in minutes and compares scheduled time to actual time.',
      '  Positive (+) = MORE time worked than scheduled (stayed late / started early)',
      '  Negative (−) = LESS time worked than scheduled (left early / came in late)',
      '',
      '### Employee Roll-Up columns (left to right)',
      '• Name — team member name as it appears in HotSchedules',
      '• Location — which store this row belongs to',
      '• Matched Days — days where both a schedule AND a punch exist',
      '• Scheduled Days — days the person was on the schedule',
      '• Worked Days — days they actually punched in',
      '• Reliability % — matched days ÷ scheduled days × 100',
      '• Start Var — total minutes of start-time variance for the week',
      '• End Var — total minutes of end-time variance for the week',
      '• Total Var — sum of all variance for the week',
      '• Avg Var — Total Var ÷ matched days',
      '• Late In — total minutes arrived late across the week',
      '• Early In — total minutes arrived early',
      '• Late Out — total minutes stayed past scheduled end',
      '• Early Out — total minutes left before scheduled end',
      '• Missed Punches — count of shifts flagged as missed clock-outs (see section 7)',
      '• Scheduled Hours — total hours on the schedule',
      '• Actual Hours — total hours worked (from punches)',
      '• OT Hours — hours above the OT threshold (default 40) for the week',
      '• Pattern — behavioral label (see below)',
      '• Swap Count — shifts where the start time was off by more than the Swap',
      '   Detection Threshold (default 120 min) — usually indicates a shift trade',
      '',
      '### Pattern labels — exactly when each one fires',
      'A pattern requires at least 40% of a person\'s matched shifts to show the',
      'behavior. Adjustable in Settings → Pattern Frequency (%).',
      '',
      '• On Track — no pattern triggered; behavior is normal',
      '• Late In — frequently arriving late',
      '• Early In — frequently arriving early',
      '• Stays Late — frequently staying past scheduled end',
      '• Leaves Early — frequently leaving before scheduled end',
      '• Works Extra — combines Early In + Stays Late (more hours than scheduled)',
      '• Works Less — combines Late In + Leaves Early (fewer hours than scheduled)',
      '• Shift Early — combines Early In + Leaves Early (shifted earlier)',
      '• Shift Late — combines Late In + Stays Late (shifted later)',
      '',
      '### Day-by-Day Detail columns',
      'One row per employee per day. Shows Date, Scheduled In, Scheduled Out,',
      'Actual In, Actual Out, Start Var, End Var, Status (matched/absent/unscheduled),',
      'and Remarks from the punch record. Missed clock-outs are marked "⚠ MISSED"',
      'next to the Actual Out time.'
    ]],

    ['7. WHAT A "MISSED CLOCK-OUT" MEANS', [
      'A "missed clock-out" means the team member forgot to punch out at the end of',
      'their shift, so the time-out recorded in the system is wrong (usually because',
      'a manager had to fix it later, or because the system auto-closed the shift).',
      '',
      '### How the tool detects it',
      'PRIMARY RULE (matched shifts — the common case):',
      'If the actual clock-out is 90 or more minutes past the scheduled end, the',
      'shift is flagged. Example: scheduled off at 10:00pm, actual out at 11:34pm',
      '= 94 minutes over = flagged.',
      '',
      'FALLBACK RULE (unscheduled shifts with no schedule to compare against):',
      'If an unscheduled shift\'s actual clock-out falls within 5 minutes of midnight',
      '(23:55–00:05), the shift is flagged as a likely auto-punchout.',
      '',
      '### How to tune sensitivity',
      'Settings → "Missed Clock-Out Threshold (min)". Default 90. Lower it if you want',
      'more flags; raise it if you want only the worst cases. A full week of data',
      'after any change is usually enough to tell if the new threshold feels right.',
      '',
      '### What to do when you see one',
      '1. Open HotSchedules and find the shift.',
      '2. Verify the correct clock-out time with the manager who was on duty.',
      '3. Correct the punch in HotSchedules.',
      '4. Coach the team member on remembering to clock out.',
      'Note: correcting it in HotSchedules does not automatically update this',
      'spreadsheet — re-run the weekly analysis if you want the report refreshed.'
    ]],

    ['8. OVERTIME — HOW IT\'S CALCULATED', [
      '### The basic rule',
      'If a team member\'s total actual hours for the week exceed the OT Threshold',
      '(default 40, configurable in Settings), the overage is counted as OT hours.',
      'Example: 43.5 actual hours - 40 threshold = 3.5 OT hours.',
      '',
      '### Cross-location OT (the important part)',
      'If a team member works at BOTH Cockrell Hill AND DBU in the same week, the',
      'tool sums their hours across both locations BEFORE checking the OT threshold.',
      '',
      'Example:',
      '  Cockrell Hill: 24 hours',
      '  DBU:           20 hours',
      '  Combined:      44 hours  → 4 hours of OT',
      '',
      'Analytics Hub looks at each location in isolation and would show "no OT" at',
      'both stores even though the person is clearly in OT. This tool catches that.',
      '',
      'Cross-location OT only works if you upload BOTH locations\' data in the same',
      'analysis run. If you only run one location, the tool cannot see the other.',
      '',
      '### What shows up on the Monday Huddle OT block',
      'The Fill Monday Sheet command writes a block showing per-location OT hours,',
      'scheduled OT (hours the schedule itself would create if everyone worked it',
      'exactly), multi-location team members, and any missed-clock-out callouts.'
    ]],

    ['9. THE TWO EMAIL TIERS', [
      'The tool sends TWO different emails from the same run:',
      '',
      '### Director Email (full detail)',
      'Goes to everyone listed in Settings → "Director Email Recipients".',
      'Contains every flagged employee, every alert, every date, and the full',
      'cross-location OT breakdown. Directors use this to review and coach managers.',
      '',
      '### Manager Email (summary)',
      'Goes to everyone listed in Settings → "Manager Email Recipients".',
      'Contains counts and top-line alerts only. Managers use this for day-to-day',
      'corrections without being overwhelmed by detail.',
      '',
      '### How to add or remove a recipient',
      '1. Open the Settings tab.',
      '2. Find the row for "Director Email Recipients" or "Manager Email Recipients".',
      '3. Edit the value cell. Format: email1@example.com, email2@example.com',
      '   — just email addresses separated by commas, no quotes, no brackets.',
      '4. Save (Sheets saves automatically). The change takes effect on the next run.',
      '',
      '### Who sends the email',
      'The email is sent from the Google account of whoever clicks "Send Emails".',
      'There is no service account. Recipients will see the sender as the person who',
      'ran the analysis. If an email ends up in spam, the recipient should whitelist',
      'the sender address.',
      '',
      '### Re-sending emails without re-running analysis',
      'Click Schedule Variance → "Send Emails (Director + Manager)" to re-send using',
      'the existing Weekly Report data. Useful if you forgot a recipient the first time.'
    ]],

    ['10. THE MONDAY HUDDLE SHEET', [
      '"Fill Monday Sheet" writes the scheduling and OT analysis onto a presentation',
      'tab that gets shown in the Monday Morning Huddle. The layout has two main',
      'sections that mirror the huddle conversation: what happened last week, and',
      'what\'s coming this week.',
      '',
      '### Where it writes',
      'It writes to the tab named in Settings → "Monday Tab Name" (default:',
      '"Current Week Schedule.ch"). If you rename that tab, update the Settings row',
      'to match or Fill Monday Sheet will not find it.',
      '',
      '### LAST WEEK REVIEW section (rows 32–37)',
      'Powered by last week\'s schedule + punch data. Shows:',
      '• Scheduled vs Actual hours per location and combined',
      '• Percentage delta (how far off actual was from scheduled)',
      '• Actual OT per location (hours above threshold, from real punches)',
      '• Scheduled OT per location (what the schedule alone would have created)',
      '• Delta between actual OT and scheduled OT',
      '',
      '### THIS WEEK PLAN section (rows 39–42)',
      'Powered by this week\'s schedule data. Shows:',
      '• Projected scheduled hours per location and combined',
      '• Projected scheduled OT — OT that the current schedule would create',
      '   if everyone works it exactly as written',
      'This is the forward-looking piece. If scheduled OT is too high, you know',
      'you need to cut hours before the week plays out.',
      '',
      '### ATTENDANCE FLAGS section (rows 44–51)',
      '• OT Offenders — who actually went into OT last week',
      '• Late Arrivals — frequent late-ins or "Late In" / "Works Less" patterns',
      '• Absences — scheduled but never clocked in',
      '• Missed Clock-Outs — forgot to punch out (see section 7)',
      '',
      '### Safe to re-run',
      'Fill Monday Sheet overwrites the same block every time — it does not duplicate',
      'rows or pile up stale data. You can run it as many times as you want.'
    ]],

    ['11. THE WEEKLY ARCHIVE (.CH TABS)', [
      'Every Monday at 5:00 PM, all tabs whose name ENDS in ".ch" are automatically',
      'copied into the "Huddle Archive" tab. This creates an append-only, permanent',
      'history of every Monday Huddle presentation — nothing is ever deleted.',
      '',
      '### What gets archived',
      'Any tab whose name ends in .ch — that is how the tool identifies presentations',
      'to snapshot. Typical tabs:',
      '  • Catering Presentation.ch',
      '  • Current Week Schedule.ch',
      '  • Talent Presentation.ch',
      '  • Facilities Presentation.ch',
      '  • CEM Presentation.ch',
      '',
      'Each snapshot includes a date/time banner and preserves formatting, colors,',
      'and merged cells so you can later look at exactly what was presented.',
      '',
      '### How to add a new presentation',
      'Rename the tab so its name ends in ".ch" (for example: "Safety Presentation.ch").',
      'It will automatically be included in the next Monday snapshot. No code change needed.',
      '',
      '### How to check that it ran',
      'Open the Huddle Archive tab and scroll to the bottom. You should see a date',
      'banner for the most recent Monday. If the date is stale, the trigger did not',
      'run — see the Troubleshooting section.',
      '',
      '### Manual controls',
      '• "Archive .ch Tabs Now" — take a snapshot immediately (good for testing)',
      '• "Setup Weekly Archive (Monday 5 PM)" — one-time setup to enable the trigger',
      '• "Remove Weekly Archive Trigger" — turn off automatic archiving'
    ]],

    ['12. TAB REFERENCE', [
      '### Documentation',
      '• Instructions — This tab. You are here.',
      '',
      '### Input tabs (auto-populated by Upload & Analyze)',
      '• CH Schedule Data — Last week Cockrell Hill schedule (for review)',
      '• CH Punch Data — Last week Cockrell Hill time punches (for review)',
      '• DBU Schedule Data — Last week DBU schedule (for review)',
      '• DBU Punch Data — Last week DBU time punches (for review)',
      '• CH Schedule Data (This Week) — This week Cockrell Hill schedule (for planning)',
      '• DBU Schedule Data (This Week) — This week DBU schedule (for planning)',
      '',
      '### Output tabs (auto-generated — DO NOT EDIT)',
      '• Weekly Report — Analysis results with color-coded flags. Overwritten on each run.',
      '• History — Permanent backend data store. Powers chronic flags and historical',
      '  reports. PROTECTED — editing this tab will corrupt historical reports.',
      '• Report Output — Destination for historical reports. Overwritten on each report run.',
      '• Huddle Archive — Permanent weekly snapshots of all .ch presentation tabs.',
      '',
      '### Presentation tabs (your content, archived automatically)',
      'Any tab whose name ends in ".ch" is a presentation tab and will be archived',
      'every Monday at 5 PM. Edit these freely — your edits are what show up in the',
      'archive snapshot.',
      '',
      '### Configuration',
      '• Settings — All configurable values. Edit here instead of in code.',
      '• Name Aliases — Optional map of "raw name" to "canonical name" for',
      '  fixing HotSchedules name inconsistencies. See section 17.'
    ]],

    ['13. SETTINGS — FULL REFERENCE', [
      'Every setting lives in the "Settings" tab. To change one, edit the Value column',
      'and save. Changes take effect on the next analysis run.',
      '',
      '### Location 1 Name / Location 2 Name',
      'Display names used in reports and emails. Default: Cockrell Hill DTO, DBU OCV.',
      'Change if either location changes names. Does not affect which CSVs get loaded.',
      '',
      '### Monday Tab Name',
      'Name of the tab where "Fill Monday Sheet" writes the OT block.',
      'Default: Current Week Schedule.ch. Must match an existing tab name exactly.',
      '',
      '### Director Email Recipients / Manager Email Recipients',
      'Comma-separated email addresses. See section 9 for details.',
      '',
      '### Trigger Alert Email',
      'Single email address that receives an alert if a background trigger',
      '(currently: the Monday 5 PM weekly archive) fails. Default: blank, which',
      'falls back to the spreadsheet owner. Set this to your own email so silent',
      'trigger failures do not go unnoticed.',
      '',
      '### OT Threshold (hours)',
      'Default 40. Weekly hours above which a team member is flagged for OT. You',
      'rarely need to change this — it should match the legal/payroll OT threshold.',
      '',
      '### High Variance Threshold (min)',
      'Default 60. Total weekly variance (in minutes) above which an employee gets',
      'flagged in the Flags & Alerts section. Lower = more flags.',
      '',
      '### Lateness Threshold (min)',
      'Default 10. Minimum minutes late for a shift to count toward the "Late In"',
      'pattern. A 5-minute arrival delay will not trigger; a 15-minute one will.',
      '',
      '### Pattern Frequency (%)',
      'Default 40. The percentage of a person\'s matched shifts that must show a',
      'behavior for a pattern label to be applied. Lower = pattern labels appear',
      'more easily; higher = only the most consistent patterns show up.',
      '',
      '### Chronic Flag Window (weeks) / Chronic Flag Trigger (weeks)',
      'Defaults: 3 / 2. "Look back over the last 3 weeks — if at least 2 of them',
      'had a flag, mark this employee as chronic." Lower the window to react faster;',
      'raise the trigger to be more conservative.',
      '',
      '### Swap Detection Threshold (min)',
      'Default 120. Minutes of start-time variance above which a shift is flagged as',
      'a probable shift swap. A team member clocking in 3 hours early or late almost',
      'certainly traded shifts with someone.',
      '',
      '### Midnight Window (min)',
      'Default 5. FALLBACK ONLY — only used for unscheduled shifts. See section 7.',
      '',
      '### Missed Clock-Out Threshold (min)',
      'Default 90. Primary rule for detecting missed clock-outs on matched shifts.',
      'See section 7 for full detail. This is the main knob if you feel the tool is',
      'flagging too many or too few missed clock-outs.'
    ]],

    ['14. HISTORICAL REPORTS', [
      'Under Schedule Variance → Historical Reports, you can run analyses that span',
      'multiple weeks of History. Each report writes its output to the "Report Output"',
      'tab, overwriting whatever was there before.',
      '',
      '### Worst Offenders — Rolling',
      'Question: "Who has the lowest reliability over the last 4, 8, or 12 weeks?"',
      'Use for: quarterly reviews, termination decisions, finding trends across the team.',
      '',
      '### Individual Employee Timeline',
      'Question: "What has this one person\'s history looked like week by week?"',
      'Use for: 1:1 conversations, performance reviews, documenting a pattern.',
      '',
      '### OT Trend by Location',
      'Question: "Is our OT at CH or DBU getting better or worse over time?"',
      'Use for: director-level staffing and labor conversations.',
      '',
      '### Chronic Lateness Leaderboard',
      'Question: "Who has been late the most over the last 4 weeks?"',
      'Use for: coaching conversations and attendance accountability.',
      '',
      '### Missed Clock-Out Frequency',
      'Question: "Who forgets to clock out the most?"',
      'Use for: training reminders and identifying team members who need extra coaching.',
      '',
      '### OT Reduction Trend',
      'Question: "Are we trending toward less OT month over month?"',
      'Use for: measuring the impact of scheduling or labor-model changes.'
    ]],

    ['15. MENU REFERENCE', [
      'Every item under the "Schedule Variance" menu, in order:',
      '',
      '### Upload & Analyze',
      'Opens the file upload dialog. The main weekly workflow starts here.',
      '',
      '### Re-run Analysis (from existing tabs)',
      'Re-runs the analysis using whatever is already in the input tabs without',
      'requiring you to re-upload. Useful if you manually corrected the input data.',
      '',
      '### Send Emails (Director + Manager)',
      'Re-sends the two-tier email using the latest analysis. Does not re-run the',
      'analysis itself.',
      '',
      '### Fill Monday Sheet',
      'Writes the OT block onto the tab named in Settings → "Monday Tab Name".',
      '',
      '### Historical Reports (submenu)',
      'Runs multi-week analyses. See section 14.',
      '',
      '### Archive .ch Tabs Now',
      'Takes an immediate snapshot of all .ch tabs into Huddle Archive. Same thing',
      'the Monday 5 PM trigger does, but manual.',
      '',
      '### Setup Weekly Archive (Monday 5 PM)',
      'One-time setup to enable the automatic Monday archive trigger.',
      '',
      '### Remove Weekly Archive Trigger',
      'Turns off the automatic archive. The manual button still works.',
      '',
      '### Clear Input Tabs',
      'Wipes the six input tabs (schedules + punches). Does NOT touch History,',
      'Weekly Report, or any presentation tab.',
      '',
      '### Rebuild Instructions Tab',
      'Force-rebuilds THIS tab from the latest content in code. Use this after a',
      'code update to refresh the instructions.',
      '',
      '### Initialize Sheet (first-time setup)',
      'Creates any missing tabs and populates Settings/Instructions if empty.',
      'Safe to re-run — never overwrites existing data.'
    ]],

    ['16. GLOSSARY', [
      '### Variance',
      'The difference between scheduled time and actual time, measured in minutes.',
      'Positive = more time than scheduled; negative = less.',
      '',
      '### Matched',
      'A day where the employee had BOTH a schedule entry AND a punch record. This',
      'is the normal case and is the only case used for variance calculations.',
      '',
      '### Absent',
      'A day where the employee had a schedule entry but NO punch. They were',
      'supposed to work but did not show up (or did not clock in).',
      '',
      '### Unscheduled',
      'A day where the employee had a punch but NO schedule entry. They worked',
      'without being on the schedule (covering a shift, walk-in, etc.).',
      '',
      '### Reliability',
      'Matched days divided by scheduled days, expressed as a percentage. 100%',
      'means every scheduled day was also worked. Below 75% is red.',
      '',
      '### Pattern',
      'A label (e.g., "Stays Late", "Works Extra") applied when at least 40% of',
      'an employee\'s matched shifts show the same behavior. See section 6.',
      '',
      '### Roll-Up',
      'The middle section of the Weekly Report — one row per employee with all',
      'their stats aggregated for the week.',
      '',
      '### Swap',
      'A shift where the start variance exceeds the Swap Detection Threshold',
      '(default 120 min). Almost always means the person traded shifts with',
      'someone who was on the schedule.',
      '',
      '### Chronic',
      'An employee who has been flagged in at least N of the last M weeks.',
      'Defaults: 2 of the last 3. Used to escalate repeat issues.',
      '',
      '### Missed Clock-Out Threshold',
      'Minutes past scheduled end before a matched shift is flagged as a missed',
      'clock-out. Default 90. See section 7.',
      '',
      '### Midnight Window',
      'Fallback rule for unscheduled shifts only — minutes from midnight to flag',
      'as a likely auto-punchout. Default 5. See section 7.',
      '',
      '### Cross-Location OT',
      'Overtime calculated on combined hours from both CH and DBU for employees',
      'who work at both stores. See section 8.',
      '',
      '### .ch Tab',
      'Any spreadsheet tab whose name ends in ".ch" — marks it as a Monday Huddle',
      'presentation tab that should be archived automatically.',
      '',
      '### Huddle Archive',
      'Append-only tab that stores weekly snapshots of all .ch tabs. See section 11.'
    ]],

    ['17. COMMON OPERATOR MISTAKES', [
      '### Uploading the wrong week',
      'The upload dialog checks History and warns you, but read the warning. If',
      'you proceed, you WILL overwrite the existing week\'s data.',
      '',
      '### Uploading schedule but forgetting punches (or vice versa)',
      'The tool needs BOTH files from a location to produce matched-day results.',
      'Only schedule = everyone shows as "absent". Only punches = everyone shows',
      'as "unscheduled". If the report looks strange, check that both files loaded.',
      '',
      '### Forgetting to update email recipients after a management change',
      'When a manager leaves or a new director joins, update the Settings tab',
      'BEFORE the next Monday run. Otherwise the wrong people get the emails.',
      '',
      '### Manually editing the History tab',
      'History feeds the chronic flag logic and every historical report. Editing',
      'it by hand will corrupt those reports. If a History row is wrong, re-run',
      'the analysis for that week instead of editing directly.',
      '',
      '### Renaming the Monday Huddle tab without updating Settings',
      '"Fill Monday Sheet" looks for the tab name listed in Settings → "Monday Tab',
      'Name". If you rename the tab, update Settings to match.',
      '',
      '### Letting name mismatches accumulate',
      'HotSchedules sometimes exports the same person differently across schedule',
      'vs punch files. The tool handles the common cases automatically:',
      '  • Trims and collapses whitespace',
      '  • Flips "Last, First" into "First Last"',
      '  • Title-cases everything (JOHN SMITH = john smith = John Smith)',
      '',
      'For the weirder cases (nicknames, "Tyler J." vs "Tyler Johnson", legal vs',
      'preferred names) use the "Name Aliases" tab. Add a row with the raw name',
      'from HotSchedules on the left and the canonical name on the right. The',
      'tool applies the alias on every parse, so next run the duplicate goes away.',
      '',
      'If a duplicate persists: check both CSVs directly to see exactly what text',
      'HotSchedules exported for that person, then add the raw form as an alias.',
      '',
      '### Closing the upload dialog before analysis finishes',
      'The analysis takes ~30 seconds. Do not close the dialog until it says "Done."',
      'Closing early will leave the input tabs partially written and the History',
      'tab not updated.'
    ]],

    ['18. TROUBLESHOOTING', [
      '### "No data found" error',
      'The input tabs are empty. Use Upload & Analyze to upload CSVs, or paste',
      'data into the input tabs manually if needed.',
      '',
      '### Employees don\'t match between schedule and punches',
      'The tool auto-normalizes names (whitespace, "Last, First" flip, title',
      'case) at parse time, so most minor mismatches are already handled. If',
      'a person still shows up twice, open the "Name Aliases" tab and add a',
      'row: raw name on the left (exactly as it appears in HotSchedules),',
      'canonical name on the right. Re-run the analysis and the duplicate',
      'should collapse into one row.',
      '',
      '### OT numbers look wrong',
      'Cross-location OT only works when BOTH locations are loaded in the same',
      'run. If you only uploaded one location, combined-hours OT is invisible.',
      'Re-run with both locations if you suspect you\'re missing OT.',
      '',
      '### Numbers don\'t match payroll exactly',
      'This tool reads HotSchedules; payroll runs off a separate corporate system.',
      'Small differences are normal and expected. This tool is NOT a payroll',
      'system of record.',
      '',
      '### Email not sending',
      'Check that email addresses in Settings are spelled correctly and separated',
      'by commas only. Emails send from your Google account — if Gmail is having',
      'issues, sends will fail. Try sending a test email from Gmail first.',
      '',
      '### Email went to spam',
      'Ask the recipient to whitelist the sender Google account address.',
      '',
      '### Archive is empty or stale',
      'Run Schedule Variance → "Setup Weekly Archive (Monday 5 PM)" once to',
      'create the trigger. Then run "Archive .ch Tabs Now" to take an immediate',
      'snapshot so the tab is not blank. If the trigger was created but still',
      'did not run, check the Trigger Alert Email recipient (Settings tab) —',
      'if the archive fails during the Monday run, the tool sends an email',
      'there with the error details.',
      '',
      '### New presentation tab not getting archived',
      'Its name must END in ".ch" exactly (e.g., "Safety Presentation.ch"). Tabs',
      'that only have .ch in the middle will be ignored.',
      '',
      '### Report shows last week but it\'s a new week',
      'You have not uploaded the new week yet. Run Upload & Analyze with the',
      'new week\'s CSVs.',
      '',
      '### Monday Huddle OT block did not update',
      'Run Schedule Variance → "Fill Monday Sheet" manually. Check that the',
      '"Monday Tab Name" setting matches the actual tab name.',
      '',
      '### Script timed out',
      'Very rare. If it happens, upload one location at a time (run CH first,',
      'then DBU) instead of both together.',
      '',
      '### "I broke something"',
      'Google Sheets keeps full version history. Go to File → Version history →',
      'See version history, pick a version from before the problem, and restore.',
      'Nothing this tool writes is irreversible as long as you have version',
      'history enabled (it is on by default).'
    ]],

    ['19. IF YOU\'RE STUCK', [
      'The spreadsheet is built to be self-maintaining. If something is broken',
      'and the troubleshooting section did not help, try these in order:',
      '',
      '1. Re-read the Troubleshooting section (17) carefully — most issues are',
      '   covered there.',
      '2. Click Schedule Variance → "Rebuild Instructions Tab" to make sure',
      '   you are reading the latest version of this manual.',
      '3. Click Schedule Variance → "Initialize Sheet (first-time setup)".',
      '   This is safe to re-run and will recreate any missing tabs.',
      '4. Use File → Version history → See version history to restore the',
      '   spreadsheet to a known-good state from before the problem appeared.',
      '5. As a last resort, contact the current Chief of Staff or technical',
      '   admin at CFA Cockrell Hill for help.',
      '',
      '### For technical successors',
      'This tool is built in Google Apps Script. To view the code: Extensions →',
      'Apps Script. The files are numbered for load order. Settings are read',
      'from the Settings tab at runtime — you almost never need to touch code',
      'to change behavior. If you are a non-technical operator, you should not',
      'need to open the Apps Script editor at all.'
    ]]
  ];

  // Build the full value array in one shot for batch write
  var rows = [];
  rows.push(['SCHEDULE VARIANCE ANALYZER — INSTRUCTIONS']);
  rows.push(['']);

  // Track row numbers for formatting pass
  var titleRow = 1;
  var sectionHeadRows = [];
  var subHeadRows = [];
  var bodyRows = [];

  sections.forEach(function(sec) {
    sectionHeadRows.push(rows.length + 1);
    rows.push([sec[0]]);
    sec[1].forEach(function(line) {
      var r = rows.length + 1;
      if (typeof line === 'string' && line.indexOf('### ') === 0) {
        rows.push([line.substring(4)]);
        subHeadRows.push(r);
      } else {
        rows.push([line]);
        bodyRows.push(r);
      }
    });
    rows.push(['']);
  });

  sheet.getRange(1, 1, rows.length, 1).setValues(rows);

  sheet.getRange(titleRow, 1).setFontSize(16).setFontWeight('bold').setFontColor('#1a1a2e');

  sectionHeadRows.forEach(function(r) {
    sheet.getRange(r, 1).setFontSize(13).setFontWeight('bold').setFontColor('#1a1a2e');
  });

  subHeadRows.forEach(function(r) {
    sheet.getRange(r, 1).setFontSize(11).setFontWeight('bold').setFontColor('#2d3748');
  });

  bodyRows.forEach(function(r) {
    sheet.getRange(r, 1).setFontSize(10).setFontColor('#555555');
  });

  sheet.setColumnWidth(1, 900);
  sheet.setFrozenRows(1);
}

function rebuildInstructionsTab() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(TABS.INSTRUCTIONS);
  if (!sheet) sheet = ss.insertSheet(TABS.INSTRUCTIONS);
  buildInstructionsTab_(ss);
  SpreadsheetApp.getUi().alert('Instructions Rebuilt', 'The Instructions tab has been refreshed with the latest content.', SpreadsheetApp.getUi().ButtonSet.OK);
}

function buildNameAliasesTab_(ss) {
  var sheet = ss.getSheetByName(TABS.NAME_ALIASES);
  sheet.getRange(1, 1, 1, 2).setValues([['Raw Name (from HotSchedules)', 'Canonical Name (what the tool should use)']]);
  sheet.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#F3F4F6');
  sheet.setColumnWidth(1, 320);
  sheet.setColumnWidth(2, 320);
  sheet.setFrozenRows(1);
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
        scheduledHours: data[i][22] || 0,
        actualHours: data[i][23] || 0,
        absTotal: Math.abs(data[i][11] || 0)
      });
    }
  }

  var locSet = {};
  var weekSet = {};
  var empMap = {};   // name → { totalHours, otHours, locations }
  var locHours = {}; // location → { schedHours, actualHours }

  rows.forEach(function(r) {
    locSet[r.locationName] = true;
    weekSet[r.weekLabel] = true;

    // Build per-employee OT data
    var schedH = r.scheduledHours || 0;
    var actualH = r.actualHours || 0;
    if (!empMap[r.name]) {
      empMap[r.name] = { totalHours: 0, otHours: 0, locations: {} };
    }
    empMap[r.name].totalHours += actualH;
    empMap[r.name].otHours = Math.max(empMap[r.name].otHours, r.otHours || 0);
    empMap[r.name].locations[r.locationName] = true;

    // Build per-location hour totals
    if (!locHours[r.locationName]) {
      locHours[r.locationName] = { schedHours: 0, actualHours: 0 };
    }
    locHours[r.locationName].schedHours += schedH;
    locHours[r.locationName].actualHours += actualH;
  });

  // Reconstruct crossLocOT structure that Email.gs expects
  var crossLocOT = { byLocation: {}, employees: {}, totalOTHours: 0, totalScheduledOTHours: 0 };
  Object.keys(locHours).forEach(function(loc) {
    crossLocOT.byLocation[loc] = {
      schedHours: rd_(locHours[loc].schedHours),
      actualHours: rd_(locHours[loc].actualHours)
    };
  });
  Object.keys(empMap).forEach(function(name) {
    var e = empMap[name];
    crossLocOT.employees[name] = {
      totalHours: rd_(e.totalHours),
      otHours: rd_(e.otHours),
      isMultiLocation: Object.keys(e.locations).length > 1
    };
    crossLocOT.totalOTHours += (e.otHours || 0);
  });
  crossLocOT.totalOTHours = rd_(crossLocOT.totalOTHours);

  return {
    rows: rows,
    crossLocOT: crossLocOT,
    locationData: Object.keys(locSet).map(function(n) { return { name: n }; }),
    weeks: Object.keys(weekSet).map(function(l) { return { label: l }; })
  };
}

function rd_(n) {
  return Math.round((n || 0) * 10) / 10;
}

function patternTypeFromLabel_(label) {
  if (!label) return 'none';
  if (label === 'On Track') return 'good';
  if (label === 'Late In' || label === 'Works Less' || label === 'Leaves Early') return 'bad';
  if (label === 'Works Extra' || label === 'Stays Late' || label === 'Shift Late') return 'warn';
  if (label === 'Early In' || label === 'Shift Early') return 'neutral';
  return 'none';
}
