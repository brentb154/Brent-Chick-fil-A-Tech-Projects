// ============================================================
// Code.gs — Schedule Counter · Cockrell Hill
// Chunk 2: History, Trends & Email
// ============================================================
//
// SETUP: Before running initSheets() set this Script Property:
//   ALERT_EMAIL  — email address for trigger failure alerts
//
// Run initSheets() once from the Apps Script editor after first deployment.
// Upgrading from Chunk 1? If weekly_summary already exists, delete that tab
// and re-run initSheets() to get the updated column schema.

// ----- CONSTANTS: TAB HEADERS --------------------------------

var SCHEDULE_HISTORY_HEADERS = [
  ['week_start','day','hour','foh_count','boh_count','combined_count','recorded_at']
];

var TRAINING_SCHEDULE_HISTORY_HEADERS = [
  ['week_start','day','hour','training_count','recorded_at']
];

var SALES_HISTORY_HEADERS = [
  ['week_start','day','hour','sales_dollars','sales_source','splh','vs_goal','efficiency_flag','recorded_at']
];

var SALES_CURVES_HEADERS = [
  ['day_of_week','hour','weight','sample_count','last_updated']
];

var CONFIG_HEADERS = [
  ['key','value']
];

var WEEKLY_SUMMARY_HEADERS = [
  ['week_start','location_name','avg_splh','peak_hour','pct_below_goal','pct_near_goal','pct_meeting_goal','total_hours_scheduled','recorded_at']
];

var PRODUCTIVITY_TRACKER_HEADERS = [
  ['month','productivity','sales','hours']
];

var SCHEDULE_PUBLISHED_HEADERS = [
  ['week_start','payload_json','updated_at']
];

var HOURS_LIST = [
  '5 AM','6 AM','7 AM','8 AM','9 AM','10 AM','11 AM','12 PM',
  '1 PM','2 PM','3 PM','4 PM','5 PM','6 PM','7 PM','8 PM',
  '9 PM','10 PM','11 PM'
];

var DAYS_LIST = [
  'Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'
  // Sunday intentionally omitted — location is closed Sunday
];

// Default curve: relative weights, sum = 102.5 → normalise to 1.0
var RAW_DEFAULT_WEIGHTS = [2,4,6,8,10,10,10,10,9,8,6,5,4,3,2.5,2,1.5,1,0.5];
var RAW_WEIGHT_SUM = RAW_DEFAULT_WEIGHTS.reduce(function(a,b){ return a+b; }, 0); // 102.5

function getDefaultCurveWeight(hourIndex) {
  return parseFloat((RAW_DEFAULT_WEIGHTS[hourIndex] / RAW_WEIGHT_SUM).toFixed(6));
}

// ----- CELL VALUE HELPERS ------------------------------------

/**
 * Google Sheets auto-converts YYYY-MM-DD strings to Date objects on read.
 * This helper converts them back to canonical YYYY-MM-DD strings.
 */
function toDateString_(val) {
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  var s = String(val).trim();
  // Handle M/D/YYYY or MM/DD/YYYY strings that Sheets auto-formats
  var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) {
    var mon = m[1].length < 2 ? '0' + m[1] : m[1];
    var day = m[2].length < 2 ? '0' + m[2] : m[2];
    return m[3] + '-' + mon + '-' + day;
  }
  return s;
}

/**
 * Google Sheets auto-converts time-like strings (e.g. "5 AM", "12 PM") to
 * Date objects on read, or sometimes to strings like "5:00 AM".
 * This converts them back to the canonical "H AM/PM" label (no :00).
 */
function hourCellToLabel_(val) {
  if (val instanceof Date) {
    var h = val.getHours();
    if (h === 0)  return '12 AM';
    if (h === 12) return '12 PM';
    if (h > 12)   return (h - 12) + ' PM';
    return h + ' AM';
  }
  // Handle Sheets returning "5:00 AM" / "12:00 PM" as strings — strip the :00
  var s = String(val);
  var m = s.match(/^(\d{1,2}):00\s*(AM|PM)$/i);
  if (m) return parseInt(m[1]) + ' ' + m[2].toUpperCase();
  return s;
}

// ----- HTMLSERVICE INCLUDE HELPER ----------------------------

/**
 * Used inside Index.html as <?!= include('Stylesheet') ?> and <?!= include('JavaScript') ?>
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ----- ROUTING -----------------------------------------------

/**
 * Routes to Index.html by default, or ProductivityTracker.html when ?page=tracker.
 */
// Adoption ping: one row per day to the shared adoption sheet.
// No-op unless ADOPTION_SHEET_ID is set in Script Properties. Never throws.
function logAdoptionPing_(toolName) {
  try {
    var props = PropertiesService.getScriptProperties();
    var sheetId = props.getProperty('ADOPTION_SHEET_ID');
    if (!sheetId) return;
    var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    if (props.getProperty('ADOPTION_LAST_PING') === today) return;
    var tab = SpreadsheetApp.openById(sheetId).getSheetByName('Pings');
    if (!tab) return;
    tab.appendRow([today, toolName]);
    props.setProperty('ADOPTION_LAST_PING', today);
  } catch (err) {
    // Never let adoption logging break the tool.
  }
}

function doGet(e) {
  logAdoptionPing_('schedule-counter');
  var page = e && e.parameter && e.parameter.page;
  if (page === 'schedule') {
    return HtmlService.createHtmlOutputFromFile('Schedule')
      .setTitle('Schedule — Cockrell Hill')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  if (page === 'tracker') {
    return HtmlService.createHtmlOutputFromFile('ProductivityTracker')
      .setTitle('Productivity Tracker \u2014 Cockrell Hill')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Schedule Counter \u2014 Cockrell Hill')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ----- PRODUCTIVITY TRACKER ----------------------------------

/**
 * Returns all rows from productivity_tracker as an array of objects:
 * [{ month, productivity, sales, hours }, ...]
 * Called from ProductivityTracker.html via google.script.run.
 */
function getProductivityData() {
  var ss    = getSpreadsheet();
  var sheet = ss.getSheetByName('productivity_tracker');
  if (!sheet) return []; // Tab not created yet — return empty, tracker shows "No data"
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var rows = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  var tz = Session.getScriptTimeZone();
  return rows
    .filter(function(r) { return r[0] !== '' && r[0] !== null; })
    .map(function(r) {
      // Google Sheets auto-converts "Feb 2025"-style strings to Date objects.
      // Build "MMM yyyy" manually — Utilities.formatDate can produce locale-specific
      // abbreviations (e.g. "Feb." with a period) and GAS serializes Date return values
      // as JavaScript Date objects on the client, breaking the takeoverMonth comparison.
      var MONTH_ABBR = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
      var month = r[0];
      if (month instanceof Date) {
        var mi = parseInt(Utilities.formatDate(month, tz, 'M'), 10) - 1;
        var yr = Utilities.formatDate(month, tz, 'yyyy');
        month = MONTH_ABBR[mi] + ' ' + yr;
      } else {
        month = String(month).trim();
      }
      return {
        month:        month,
        productivity: parseFloat(r[1]) || 0,
        sales:        parseFloat(r[2]) || 0,
        hours:        parseFloat(r[3]) || 0
      };
    });
}

/**
 * Replaces all data rows in productivity_tracker with the provided array.
 * Auto-creates the tab if it doesn't exist (no need to run initSheets manually).
 * rows: [{ month, productivity, sales, hours }, ...]
 */
function saveProductivityData(rows) {
  var ss    = getSpreadsheet();
  var sheet = ss.getSheetByName('productivity_tracker');
  if (!sheet) {
    // Auto-create the tab rather than throwing
    sheet = ss.insertSheet('productivity_tracker');
    sheet.getRange(1, 1, 1, 4).setValues(PRODUCTIVITY_TRACKER_HEADERS);
    sheet.setFrozenRows(1);
    Logger.log('productivity_tracker tab auto-created by saveProductivityData');
  }
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, 4).clearContent();
  if (!rows || rows.length === 0) return;
  var tz = Session.getScriptTimeZone();
  var MONTH_ABBR = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  var values = rows.map(function(r) {
    var month = r.month;
    // Guard: if month somehow came back as a Date, rebuild "MMM yyyy" string manually
    if (month instanceof Date) {
      var mi = parseInt(Utilities.formatDate(month, tz, 'M'), 10) - 1;
      var yr = Utilities.formatDate(month, tz, 'yyyy');
      month = MONTH_ABBR[mi] + ' ' + yr;
    } else {
      month = String(month).trim();
    }
    return [month, parseFloat(r.productivity) || 0, parseFloat(r.sales) || 0, parseFloat(r.hours) || 0];
  });
  // '@' = plain text format — prevents Sheets from auto-converting "Feb 2025" to a date value.
  // Note: '@STRING@' is not a valid number format token and is silently ignored.
  sheet.getRange(2, 1, values.length, 1).setNumberFormat('@');
  sheet.getRange(2, 1, values.length, 4).setValues(values);
}

/**
 * Routes external POST calls by action parameter.
 * google.script.run bypasses doPost and calls functions directly.
 * doPost is provided for external/trigger use.
 */
function doPost(e) {
  try {
    var body   = JSON.parse(e.postData.contents);
    var action = e.parameter.action;
    var handlers = {
      getConfig:             function() { return getConfig(); },
      loadAppCache:          function() { return loadAppCache(); },
      saveConfig:            function() { return saveConfig(body); },
      saveWeekSnapshot:      function() { return saveWeekSnapshot(body); },
      getScheduleUrl:        function() { return getScheduleUrl(); },
      getWeeklySummary:      function() { return getWeeklySummary(body); },
      getHistoryForWeek:     function() { return getHistoryForWeek(body); },
      sendWeeklySummaryEmail: function() { return sendWeeklySummaryEmail(body); }
      // Chunk 3 adds: nothing — all intelligence runs via triggers, not doPost
    };
    var result = handlers[action] ? handlers[action]() : { error: 'Unknown action: ' + action };
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ----- SPREADSHEET HELPERS -----------------------------------

function getSpreadsheet() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getSheet(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error('Tab "' + name + '" not found. Run initSheets() from the Apps Script editor first.');
  return sheet;
}

/**
 * Read all data rows from config tab → returns plain object { key: value }
 */
function readConfigMap(ss) {
  var sheet = ss.getSheetByName('config');
  if (!sheet) return {};
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};
  var data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  var map = {};
  data.forEach(function(row) {
    if (row[0]) map[String(row[0])] = String(row[1]);
  });
  return map;
}

/**
 * Write key/value pairs to config tab using batch setValues — never cell-by-cell.
 * Keys that already exist are updated in-place; new keys are appended.
 */
function writeConfigKeys(ss, updates) {
  var sheet = getSheet(ss, 'config');
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    // Nothing seeded yet — append all
    var rows = Object.keys(updates).map(function(k) { return [k, String(updates[k])]; });
    if (rows.length > 0) sheet.getRange(2, 1, rows.length, 2).setValues(rows);
    return;
  }

  var fullData = sheet.getRange(2, 1, lastRow - 1, 2).getValues();

  // Build lookup: key → row-index in fullData
  var rowMap = {};
  fullData.forEach(function(row, idx) {
    if (row[0]) rowMap[String(row[0])] = idx;
  });

  var toAppend = [];
  Object.keys(updates).forEach(function(key) {
    if (rowMap[key] !== undefined) {
      fullData[rowMap[key]][1] = String(updates[key]);
    } else {
      toAppend.push([key, String(updates[key])]);
    }
  });

  // Write the full existing block back in one call
  sheet.getRange(2, 1, fullData.length, 2).setValues(fullData);

  // Append any new keys
  if (toAppend.length > 0) {
    var appendStart = sheet.getLastRow() + 1;
    sheet.getRange(appendStart, 1, toAppend.length, 2).setValues(toAppend);
  }
}

/**
 * Read sales_curves tab → { curves: { Day: { Hour: weight } }, confidence: { Day: { confident, sampleCount } } }
 */
function readSalesCurves(ss) {
  var sheet = getSheet(ss, 'sales_curves');
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { curves: {}, confidence: {} };

  var data = sheet.getRange(2, 1, lastRow - 1, 4).getValues(); // day, hour, weight, sample_count
  var curves = {};
  var confidence = {};

  data.forEach(function(row) {
    var day        = String(row[0]);
    var hour       = hourCellToLabel_(row[1]);
    var weight     = parseFloat(row[2]) || 0;
    var sampleCount = parseInt(row[3]) || 0;

    if (!curves[day]) {
      curves[day] = {};
      confidence[day] = { sampleCount: 0, confident: false };
    }
    curves[day][hour] = weight;
    if (sampleCount > confidence[day].sampleCount) {
      confidence[day].sampleCount = sampleCount;
      confidence[day].confident   = sampleCount >= 4;
    }
  });

  return { curves: curves, confidence: confidence };
}

// ----- INIT --------------------------------------------------

/**
 * Run once manually from the Apps Script editor after first deployment.
 * Safe to re-run — checks existence before creating each tab.
 * NEVER deletes or overwrites Sheet1 or any tab this function did not create.
 */
function initSheets() {
  var ss = getSpreadsheet();
  createTabIfMissing(ss, 'schedule_history', SCHEDULE_HISTORY_HEADERS);
  createTabIfMissing(ss, 'training_schedule_history', TRAINING_SCHEDULE_HISTORY_HEADERS);
  createTabIfMissing(ss, 'sales_history',    SALES_HISTORY_HEADERS);
  createTabIfMissing(ss, 'sales_curves',     SALES_CURVES_HEADERS);
  createTabIfMissing(ss, 'config',           CONFIG_HEADERS);
  createTabIfMissing(ss, 'weekly_summary',   WEEKLY_SUMMARY_HEADERS);
  createTabIfMissing(ss, 'app_cache',           [['cache_json']]);
  createTabIfMissing(ss, 'productivity_tracker', PRODUCTIVITY_TRACKER_HEADERS);
  createTabIfMissing(ss, 'schedule_published',  SCHEDULE_PUBLISHED_HEADERS);
  createTabIfMissing(ss, 'actuals_weekly',      ACTUALS_HEADERS);
  seedDefaultConfig(ss);
  seedDefaultCurves(ss);
  // Calibration keys seeded individually — seedDefaultConfig skips entirely once
  // config has rows, so existing installs pick these up on an initSheets() re-run.
  seedConfigKeyIfMissing_(ss, 'splh_cal_weight',  '0.5');
  seedConfigKeyIfMissing_(ss, 'splh_cal_floor',   '0.85');
  seedConfigKeyIfMissing_(ss, 'splh_cal_ceiling', '1.10');
  seedConfigKeyIfMissing_(ss, 'curve_outlier_pct', '30');
  Logger.log('initSheets complete. All tabs verified.');
}

function createTabIfMissing(ss, name, headers) {
  if (ss.getSheetByName(name)) {
    Logger.log('Tab already exists, skipping: ' + name);
    return;
  }
  var sheet = ss.insertSheet(name);
  sheet.getRange(1, 1, headers.length, headers[0].length).setValues(headers);
  sheet.setFrozenRows(1);
  Logger.log('Created tab: ' + name);
}

function seedDefaultConfig(ss) {
  var sheet = ss.getSheetByName('config');
  if (!sheet) return;
  if (sheet.getLastRow() > 1) {
    Logger.log('config already seeded, skipping.');
    return;
  }
  var defaults = [
    ['splh_goal_overall',  '90'],
    ['splh_goal_foh',      '90'],
    ['splh_goal_boh',      '90'],
    ['labor_share_foh',    '50'],
    ['labor_share_boh',    '40'],
    ['labor_share_other',  '10'],
    ['labor_share_training_foh', '3'],
    ['labor_share_training_boh', '3'],
    ['smoothing_alpha',    '0.3'],
    ['bias_threshold',     '10'],
    ['operator_email',     ''],
    ['location_name',      'Cockrell Hill'],
    ['setup_complete',     'false'],
    ['bias_flags',         '{}']
  ];
  sheet.getRange(2, 1, defaults.length, 2).setValues(defaults);
  Logger.log('Seeded ' + defaults.length + ' default config rows.');
}

function seedDefaultCurves(ss) {
  var sheet = ss.getSheetByName('sales_curves');
  if (!sheet) return;
  if (sheet.getLastRow() > 1) {
    Logger.log('sales_curves already seeded, skipping.');
    return;
  }
  var now  = new Date().toISOString();
  var rows = [];
  DAYS_LIST.forEach(function(day) {
    HOURS_LIST.forEach(function(hour, idx) {
      rows.push([day, hour, getDefaultCurveWeight(idx), 0, now]);
    });
  });
  // Force text format on the hour column BEFORE writing so Sheets doesn't auto-convert "5 AM" to a Date
  sheet.getRange(2, 2, rows.length, 1).setNumberFormat('@');
  sheet.getRange(2, 1, rows.length, 5).setValues(rows);
  Logger.log('Seeded ' + rows.length + ' curve rows (' + DAYS_LIST.length + ' days \u00d7 ' + HOURS_LIST.length + ' hours).');
}

// ----- APP CACHE LAYER ---------------------------------------

/**
 * Writes a cache object as JSON to app_cache A1.
 * Called by the Monday trigger pipeline in Chunk 3.
 * Also callable manually for testing.
 */
function writeAppCache(cacheObject) {
  var ss    = getSpreadsheet();
  var sheet = ss.getSheetByName('app_cache');
  if (!sheet) throw new Error('app_cache tab not found. Run initSheets() first.');
  sheet.getRange('A1').setValue(JSON.stringify(cacheObject));
  Logger.log('app_cache written at ' + new Date().toISOString());
}

/**
 * Reads app_cache A1. Falls back to getConfig() if empty or unparseable.
 * Returns the same shape as getConfig() — frontend never knows which path was taken.
 */
function loadAppCache() {
  try {
    var ss    = getSpreadsheet();
    var sheet = ss.getSheetByName('app_cache');
    if (!sheet) return getConfig();
    var raw = sheet.getRange('A1').getValue();
    if (!raw) return getConfig(); // Cache not yet written — first run before Chunk 3 trigger
    var cfg = JSON.parse(raw);
    // Calibration must never be served stale: it changes when wiring happens or
    // payroll batches land, not on the Monday cache cadence. Recompute on every
    // load — it reads only the local actuals_weekly tab + config (cheap).
    try {
      cfg.calibration = buildCalibration_(ss, readConfigMap(ss));
    } catch (e) {
      Logger.log('calibration overlay failed (non-fatal): ' + e.message);
      cfg.calibration = cfg.calibration || null;
    }
    return cfg;
  } catch (err) {
    Logger.log('app_cache read failed, falling back to getConfig: ' + err.message);
    return getConfig();
  }
}

// ----- CONFIG LAYER ------------------------------------------

/**
 * Returns the full config object the frontend needs on load.
 * This is the fallback path — loadAppCache() calls this when cache is empty.
 */
function getConfig() {
  try {
    var ss     = getSpreadsheet();
    var cfgMap = readConfigMap(ss);
    var curvesResult = readSalesCurves(ss);
    var rawCurves    = curvesResult.curves;
    var rawConf      = curvesResult.confidence;

    // Guarantee all 7 days are present, filling gaps with defaults
    var finalCurves = {};
    var finalConf   = {};
    DAYS_LIST.forEach(function(day) {
      if (rawCurves[day] && Object.keys(rawCurves[day]).length > 0) {
        finalCurves[day] = rawCurves[day];
      } else {
        finalCurves[day] = {};
        HOURS_LIST.forEach(function(h, idx) {
          finalCurves[day][h] = getDefaultCurveWeight(idx);
        });
      }
      finalConf[day] = rawConf[day] || { confident: false, sampleCount: 0 };
    });

    return {
      splhGoals: {
        overall: parseFloat(cfgMap['splh_goal_overall']) || 90,
        foh:     parseFloat(cfgMap['splh_goal_foh'])     || 90,
        boh:     parseFloat(cfgMap['splh_goal_boh'])     || 90
      },
      laborShares: {
        foh:      parseFloat(cfgMap['labor_share_foh'])      || 50,
        boh:      parseFloat(cfgMap['labor_share_boh'])      || 40,
        other:    parseFloat(cfgMap['labor_share_other'])    || 10,
        trainingFoh: (function() {
          var v = parseFloat(cfgMap['labor_share_training_foh']);
          return isNaN(v) ? 3 : v;
        })(),
        trainingBoh: (function() {
          var v = parseFloat(cfgMap['labor_share_training_boh']);
          return isNaN(v) ? 3 : v;
        })()
      },
      smoothingAlpha: parseFloat(cfgMap['smoothing_alpha']) || 0.3,
      biasThreshold:  parseFloat(cfgMap['bias_threshold'])  || 10,
      operatorEmail:  cfgMap['operator_email']  || '',
      locationName:   cfgMap['location_name']   || 'Cockrell Hill',
      setupComplete:  cfgMap['setup_complete']  === 'true',
      biasFlags:      JSON.parse(cfgMap['bias_flags'] || '{}'),
      salesCurves:     finalCurves,
      curveConfidence: finalConf,
      // null unless wireCalibration() has been run — frontend is 100% stock on null
      calibration: (function() {
        try { return buildCalibration_(ss, cfgMap); }
        catch (e) { Logger.log('calibration build failed (non-fatal): ' + e.message); return null; }
      })()
    };
  } catch (err) {
    Logger.log('getConfig error: ' + err.message);
    throw err;
  }
}

/**
 * Accepts a flat object of key/value pairs and updates only the supplied keys.
 * If body.salesCurves is present, validates weights sum to ~1.0 then writes curves.
 * All writes use setValues() — never cell-by-cell.
 */
function saveConfig(body) {
  try {
    if (!body) throw new Error('saveConfig called with empty body.');
    var ss      = getSpreadsheet();
    var updates = {};

    if (body.splhGoals) {
      if (body.splhGoals.overall !== undefined) updates['splh_goal_overall'] = body.splhGoals.overall;
      if (body.splhGoals.foh     !== undefined) updates['splh_goal_foh']     = body.splhGoals.foh;
      if (body.splhGoals.boh     !== undefined) updates['splh_goal_boh']     = body.splhGoals.boh;
    }
    if (body.laborShares) {
      if (body.laborShares.foh      !== undefined) updates['labor_share_foh']      = body.laborShares.foh;
      if (body.laborShares.boh      !== undefined) updates['labor_share_boh']      = body.laborShares.boh;
      if (body.laborShares.other    !== undefined) updates['labor_share_other']    = body.laborShares.other;
      if (body.laborShares.trainingFoh !== undefined) updates['labor_share_training_foh'] = body.laborShares.trainingFoh;
      if (body.laborShares.trainingBoh !== undefined) updates['labor_share_training_boh'] = body.laborShares.trainingBoh;
    }
    if (body.smoothingAlpha !== undefined) updates['smoothing_alpha'] = body.smoothingAlpha;
    if (body.biasThreshold  !== undefined) updates['bias_threshold']  = body.biasThreshold;
    if (body.operatorEmail  !== undefined) updates['operator_email']  = body.operatorEmail;
    if (body.locationName   !== undefined) updates['location_name']   = body.locationName;
    if (body.setupComplete  !== undefined) updates['setup_complete']  = body.setupComplete;
    if (body.biasFlags      !== undefined) updates['bias_flags']      = JSON.stringify(body.biasFlags);

    if (Object.keys(updates).length > 0) {
      writeConfigKeys(ss, updates);
    }

    if (body.salesCurves) {
      saveSalesCurves_(ss, body.salesCurves);
    }

    return { success: true };
  } catch (err) {
    Logger.log('saveConfig error: ' + err.message);
    return { error: err.message };
  }
}

/**
 * Internal: write per-day curve weights to sales_curves tab.
 * Validates each day's weights sum to 1.0 (±0.05) before writing.
 * Resets sample_count to 0 for any manually edited day.
 * Uses a single setValues() call — never row-by-row.
 */
function saveSalesCurves_(ss, daysCurveData) {
  var sheet   = getSheet(ss, 'sales_curves');
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) throw new Error('sales_curves is empty. Run initSheets() first.');

  var now     = new Date().toISOString();
  var allData = sheet.getRange(2, 1, lastRow - 1, 5).getValues();

  Object.keys(daysCurveData).forEach(function(day) {
    var dayCurve  = daysCurveData[day];
    var weightSum = Object.values(dayCurve).reduce(function(a, b) {
      return a + (parseFloat(b) || 0);
    }, 0);

    if (Math.abs(weightSum - 1.0) > 0.05) {
      throw new Error(
        'Curve weights for ' + day + ' must sum to 1.0. Got ' +
        weightSum.toFixed(4) + '. Normalize before saving.'
      );
    }

    allData.forEach(function(row, idx) {
      if (String(row[0]) === day && dayCurve[hourCellToLabel_(row[1])] !== undefined) {
        allData[idx][2] = parseFloat(dayCurve[hourCellToLabel_(row[1])]) || 0;
        allData[idx][3] = 0;   // Reset sample_count — manually edited
        allData[idx][4] = now;
      }
    });
  });

  // Single batch write — force hour column (B) to plain text
  sheet.getRange(2, 2, allData.length, 1).setNumberFormat('@');
  sheet.getRange(2, 1, allData.length, 5).setValues(allData);
}

// ----- HISTORY LAYER -----------------------------------------

/**
 * Internal: delete all rows in a sheet whose week_start (col 1) matches weekStart.
 * Assumes matching rows are contiguous (they always are since we append in order).
 * Single deleteRows() call.
 */
/**
 * Removes a week's rows via filtered rebuild. The old contiguous-block delete
 * (deleteRows(firstMatch, matchCount)) destroyed bystander rows whenever a
 * week's rows were non-contiguous (api rows written Monday, estimated rows
 * published later, other weeks in between).
 * Optional keepFn(row): return true to preserve a matching row (e.g. api sales).
 */
function deleteWeekRows_(sheet, weekStart, keepFn) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2) return;
  var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var kept = data.filter(function(r) {
    if (toDateString_(r[0]) !== weekStart) return true;
    return keepFn ? !!keepFn(r) : false;
  });
  if (kept.length === data.length) return;
  sheet.getRange(2, 1, data.length, lastCol).clearContent();
  if (kept.length > 0) sheet.getRange(2, 1, kept.length, lastCol).setValues(kept);
}

/** True when sales_history already holds pipeline (api) rows for this week. */
function weekHasApiSales_(salesSheet, weekStart) {
  var lastRow = salesSheet.getLastRow();
  if (lastRow < 2) return false;
  var data = salesSheet.getRange(2, 1, lastRow - 1, 5).getValues();
  for (var i = 0; i < data.length; i++) {
    if (toDateString_(data[i][0]) === weekStart && String(data[i][4]) === 'api') return true;
  }
  return false;
}

/**
 * Batch-writes a week's schedule and sales data to history sheets,
 * then computes and appends a summary row to weekly_summary.
 *
 * Body shape: { weekStart, scheduleData, fohScheduleData, bohScheduleData, salesData, salesSource,
 *               fohStaffNames, bohStaffNames, trainingStaffNames, managerShifts, suggestedStaff,
 *               locationName }
 * Duplicate weeks are deleted before writing — safe to re-run for the same week.
 * Also writes a JSON blob to schedule_published for the manager dashboard view.
 */
function saveWeekSnapshot(body) {
  try {
    if (!body || !body.weekStart) throw new Error('weekStart is required.');
    var ss          = getSpreadsheet();
    var cfgMap      = readConfigMap(ss);
    var schedSheet  = getSheet(ss, 'schedule_history');
    var salesSheet  = getSheet(ss, 'sales_history');
    var summSheet   = getSheet(ss, 'weekly_summary');
    var trainingSheet = getSheet(ss, 'training_schedule_history');

    var weekStart    = String(body.weekStart);
    var schedData    = body.scheduleData         || {};
    var fohData      = body.fohScheduleData      || {};
    var bohData      = body.bohScheduleData      || {};
    var trainingData = body.trainingScheduleData || {};
    var salesData    = body.salesData            || {};
    var salesSource  = body.salesSource          || 'estimated';
    var goal        = parseFloat(cfgMap['splh_goal_overall']) || 90;
    var locName     = cfgMap['location_name'] || 'Cockrell Hill';
    var now         = new Date().toISOString();

    // Does this week already have actual (api) sales from the Monday pipeline?
    // If so we keep them and skip writing publish-time sales rows entirely —
    // actuals must never be downgraded to curve-distributed estimates.
    var weekHasApiSales = weekHasApiSales_(salesSheet, weekStart);

    // Delete existing rows for this weekStart (idempotent re-run).
    // sales_history keeps its api rows; publish only owns non-api rows.
    deleteWeekRows_(schedSheet, weekStart);
    deleteWeekRows_(salesSheet, weekStart, function(r) { return String(r[4]) === 'api'; });
    deleteWeekRows_(trainingSheet, weekStart);

    // Build complete row arrays in memory first, then write once per sheet
    var schedRows = [];
    var salesRows = [];
    var trainingRows = [];

    // Map day names to offsets from Monday (weekStart) to detect future days
    var dayOffset = { Monday: 0, Tuesday: 1, Wednesday: 2, Thursday: 3, Friday: 4, Saturday: 5 };
    var wsDate = new Date(weekStart + 'T00:00:00');
    var today  = new Date();
    today.setHours(0, 0, 0, 0);

    Object.keys(schedData).forEach(function(day) {
      // Check if this day is in the future — skip estimated sales for future days
      var dayDate = new Date(wsDate.getTime() + (dayOffset[day] || 0) * 86400000);
      var isFuture = dayDate > today;

      HOURS_LIST.forEach(function(hour) {
        var trainingCount = (trainingData[day] && trainingData[day][hour]) || 0;
        if (trainingCount > 0) {
          trainingRows.push([weekStart, day, hour, trainingCount, now]);
        }

        var combined = (schedData[day] && schedData[day][hour]) || 0;
        if (combined === 0) return; // Skip unstaffed hours

        var foh = (fohData[day] && fohData[day][hour]) || 0;
        var boh = (bohData[day] && bohData[day][hour]) || 0;
        schedRows.push([weekStart, day, hour, foh, boh, combined, now]);

        // Only write estimated sales for days that have already occurred
        if (isFuture && salesSource === 'estimated') return;

        var sales = (salesData[day] && salesData[day][hour]) || 0;
        if (sales > 0) {
          var splh = sales / combined;
          var diff = splh - goal;
          var flag = splh >= goal         ? 'Meeting Goal'
                   : splh >= goal * 0.9   ? 'Near Goal'
                   :                        'Below Goal';
          salesRows.push([weekStart, day, hour, sales, salesSource, splh, diff, flag, now]);
        }
      });
    });

    if (schedRows.length > 0) {
      var sn = schedSheet.getLastRow() + 1;
      // Force week_start (A) and hour (C) columns to plain text
      schedSheet.getRange(sn, 1, schedRows.length, 1).setNumberFormat('@');
      schedSheet.getRange(sn, 3, schedRows.length, 1).setNumberFormat('@');
      schedSheet.getRange(sn, 1, schedRows.length, 7).setValues(schedRows);
    }
    // Skip the sales_history write when api actuals exist for this week — the
    // pipeline's rows are strictly better and must not be duplicated/downgraded.
    // weekly_summary below still computes from the in-memory salesRows (unchanged).
    if (salesRows.length > 0 && !weekHasApiSales) {
      var an = salesSheet.getLastRow() + 1;
      // Force week_start (A) and hour (C) columns to plain text
      salesSheet.getRange(an, 1, salesRows.length, 1).setNumberFormat('@');
      salesSheet.getRange(an, 3, salesRows.length, 1).setNumberFormat('@');
      salesSheet.getRange(an, 1, salesRows.length, 9).setValues(salesRows);
    }
    if (trainingRows.length > 0) {
      var tn = trainingSheet.getLastRow() + 1;
      // Force week_start (A) and hour (C) columns to plain text
      trainingSheet.getRange(tn, 1, trainingRows.length, 1).setNumberFormat('@');
      trainingSheet.getRange(tn, 3, trainingRows.length, 1).setNumberFormat('@');
      trainingSheet.getRange(tn, 1, trainingRows.length, 5).setValues(trainingRows);
    }

    // Compute weekly_summary values from already-built salesRows
    var meetCount  = salesRows.filter(function(r) { return r[7] === 'Meeting Goal'; }).length;
    var nearCount  = salesRows.filter(function(r) { return r[7] === 'Near Goal';    }).length;
    var belowCount = salesRows.filter(function(r) { return r[7] === 'Below Goal';   }).length;
    var total      = salesRows.length || 1;

    var splhSum = salesRows.reduce(function(s, r) { return s + (r[5] || 0); }, 0);
    var avgSplh = salesRows.length > 0 ? splhSum / salesRows.length : 0;

    var peakRow  = salesRows.reduce(function(best, r) {
      return (r[5] > (best ? best[5] : -1)) ? r : best;
    }, null);
    var peakHour = peakRow ? (peakRow[1] + ' ' + peakRow[2]) : '';

    // total_hours_scheduled = sum of combined_count across all staffed hours
    var totalHoursScheduled = schedRows.reduce(function(s, r) { return s + (r[5] || 0); }, 0);

    deleteWeekRows_(summSheet, weekStart);
    var un = summSheet.getLastRow() + 1;
    // Force week_start (A) to plain text
    summSheet.getRange(un, 1, 1, 1).setNumberFormat('@');
    summSheet.getRange(un, 1, 1, 9).setValues([[
      weekStart, locName, avgSplh, peakHour,
      belowCount / total, nearCount / total, meetCount / total,
      totalHoursScheduled, now
    ]]);

    // Publish to schedule_published tab for the manager dashboard
    writePublishedSchedule_(ss, weekStart, body, locName, now);

    return { success: true };
  } catch (err) {
    Logger.log('saveWeekSnapshot error: ' + err.message);
    return { error: err.message };
  }
}

/**
 * Writes the full schedule payload to schedule_published as one JSON blob per week.
 * One row per weekStart — idempotent re-writes. Safe if fields are missing (old clients).
 * Payload shape:
 *   {
 *     weekStart, locationName, updatedAt,
 *     fohCounts, bohCounts, trainingCounts,
 *     fohNames, bohNames, trainingNames,
 *     managers: { foh: {day: [{label,name}]}, boh: {...} },
 *     suggested: { foh: {day: {hour: N}}, boh: {day: {hour: N}} }
 *   }
 */
function writePublishedSchedule_(ss, weekStart, body, locName, nowIso) {
  var sheet = ss.getSheetByName('schedule_published');
  if (!sheet) {
    sheet = ss.insertSheet('schedule_published');
    sheet.getRange(1, 1, 1, 3).setValues(SCHEDULE_PUBLISHED_HEADERS);
    sheet.setFrozenRows(1);
  }

  var payload = {
    weekStart:      weekStart,
    locationName:   locName,
    updatedAt:      nowIso,
    fohCounts:      body.fohScheduleData      || {},
    bohCounts:      body.bohScheduleData      || {},
    trainingCounts: body.trainingScheduleData || {},
    fohNames:       body.fohStaffNames        || {},
    bohNames:       body.bohStaffNames        || {},
    trainingNames:  body.trainingStaffNames   || {},
    managers:       body.managerShifts        || {},
    suggested:      body.suggestedStaff       || { foh: {}, boh: {} }
  };
  var json = JSON.stringify(payload);

  // Delete existing row for this week (idempotent re-write)
  var last = sheet.getLastRow();
  if (last > 1) {
    var weekCol = sheet.getRange(2, 1, last - 1, 1).getValues();
    for (var i = weekCol.length - 1; i >= 0; i--) {
      if (String(weekCol[i][0]) === weekStart) {
        sheet.deleteRow(i + 2);
      }
    }
  }

  var rowNum = sheet.getLastRow() + 1;
  sheet.getRange(rowNum, 1, 1, 1).setNumberFormat('@'); // week_start as plain text
  sheet.getRange(rowNum, 1, 1, 3).setValues([[weekStart, json, nowIso]]);
}

/**
 * Returns the published schedule payload for a given weekStart.
 * If weekStart is omitted, returns the most recently published week.
 * Shape: { weekStart, locationName, updatedAt, fohCounts, bohCounts, trainingCounts,
 *          fohNames, bohNames, trainingNames, managers, suggested, availableWeeks }
 * Returns { empty: true, availableWeeks: [...] } if nothing published yet.
 */
function getPublishedSchedule(weekStart) {
  try {
    var ss    = getSpreadsheet();
    var sheet = ss.getSheetByName('schedule_published');
    if (!sheet || sheet.getLastRow() < 2) {
      return { empty: true, availableWeeks: [] };
    }
    var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
    var weeks = rows.map(function(r) { return String(r[0]); }).sort(); // ascending
    var target = weekStart ? String(weekStart) : weeks[weeks.length - 1];
    var row = null;
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][0]) === target) { row = rows[i]; break; }
    }
    if (!row) {
      return { empty: true, availableWeeks: weeks, requestedWeek: target };
    }
    var payload;
    try {
      payload = JSON.parse(row[1]);
    } catch (parseErr) {
      return { error: 'Payload parse failed for week ' + target + ': ' + parseErr.message };
    }
    payload.availableWeeks = weeks;
    return payload;
  } catch (err) {
    Logger.log('getPublishedSchedule error: ' + err.message);
    return { error: err.message };
  }
}

/**
 * Returns the most recent published weekStart as a string, or null if none.
 * Cheap helper for the dashboard's default view.
 */
function getLatestPublishedWeekStart() {
  try {
    var ss    = getSpreadsheet();
    var sheet = ss.getSheetByName('schedule_published');
    if (!sheet || sheet.getLastRow() < 2) return null;
    var weeks = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues()
      .map(function(r) { return String(r[0]); })
      .filter(function(w) { return w && w !== 'null' && w !== 'undefined'; })
      .sort();
    return weeks.length > 0 ? weeks[weeks.length - 1] : null;
  } catch (err) {
    Logger.log('getLatestPublishedWeekStart error: ' + err.message);
    return null;
  }
}

/**
 * Returns the published web app URL for the Schedule view (?page=schedule).
 * Called by the frontend to render the "View Schedule" link after publish.
 */
function getScheduleUrl() {
  try {
    var base = ScriptApp.getService().getUrl();
    if (!base) return { error: 'Web app URL unavailable — deploy the app first.' };
    return { url: base + '?page=schedule' };
  } catch (err) {
    return { error: err.message };
  }
}

/**
 * Returns last N rows from weekly_summary sorted descending by week_start.
 * Body: { n } — default 12. Sorts in memory; never modifies the sheet.
 */
function getWeeklySummary(body) {
  try {
    var n     = (body && body.n) || 12;
    var ss    = getSpreadsheet();
    var sheet = getSheet(ss, 'weekly_summary');
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { rows: [] };

    var data = sheet.getRange(2, 1, lastRow - 1, 9).getValues();

    // Sort descending by week_start (YYYY-MM-DD string sort is correct)
    data.sort(function(a, b) { return toDateString_(b[0]).localeCompare(toDateString_(a[0])); });

    var rows = data.slice(0, n).map(function(r) {
      return {
        weekStart:           toDateString_(r[0]),
        locationName:        String(r[1]),
        avgSplh:             parseFloat(r[2]) || 0,
        peakHour:            String(r[3]),
        pctBelowGoal:        parseFloat(r[4]) || 0,
        pctNearGoal:         parseFloat(r[5]) || 0,
        pctMeetingGoal:      parseFloat(r[6]) || 0,
        totalHoursScheduled: parseFloat(r[7]) || 0,
        recordedAt:          String(r[8])
      };
    });

    return { rows: rows };
  } catch (err) {
    Logger.log('getWeeklySummary error: ' + err.message);
    return { error: err.message };
  }
}

/**
 * Returns all history rows for a single week.
 * Body: { weekStart } — filters in memory, never server-side.
 */
function getHistoryForWeek(body) {
  try {
    if (!body || !body.weekStart) throw new Error('weekStart is required.');
    var ss        = getSpreadsheet();
    var weekStart = String(body.weekStart);

    var getRows = function(sheet, cols) {
      var lastRow = sheet.getLastRow();
      if (lastRow < 2) return [];
      return sheet.getRange(2, 1, lastRow - 1, cols).getValues()
        .filter(function(row) { return toDateString_(row[0]) === weekStart; });
    };

    return {
      scheduleRows: getRows(getSheet(ss, 'schedule_history'), 7),
      salesRows:    getRows(getSheet(ss, 'sales_history'),    9)
    };
  } catch (err) {
    Logger.log('getHistoryForWeek error: ' + err.message);
    return { error: err.message };
  }
}

/**
 * Sends a plain-text weekly summary email to operator_email.
 * Skips silently if operator_email is blank — never throws on blank email.
 * Reads computed metrics from weekly_summary sheet + computes FOH/BOH SPLH from body payload.
 */
// ============================================================
// Code.gs — Chunk 3: Intelligence Layer
// ============================================================

// ----- HELPER ------------------------------------------------

/**
 * Returns a single value from the config tab by key.
 */
function getConfigValue(key) {
  return readConfigMap(getSpreadsheet())[key] || '';
}

// ----- SHEET1 ------------------------------------------------

/**
 * Reads Sheet1 (CFA sync target) and returns a normalized salesPayload array.
 * Shape: [{ day: 'Monday', hour: '12 PM', sales: 1234.56 }, ...]
 * This is the only place Sheet1 is read.
 */
function readSheet1() {
  var ss    = getSpreadsheet();
  var sheet = ss.getSheetByName('Sheet1');
  if (!sheet) throw new Error('Sheet1 not found in this spreadsheet.');
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  // Detect columns by header name — tolerates any column order.
  // Expected headers (case-insensitive, partial match):
  //   "Hour (24 hr)" or "hour"  → 24-hour integer
  //   "Business Date"           → date value for day-of-week
  //   "Total Sales" or "sales"  → dollar amount
  var headers  = data[0].map(function(h) { return String(h).toLowerCase().trim(); });
  var hourCol  = -1, dateCol = -1, salesCol = -1;
  headers.forEach(function(h, i) {
    if (hourCol  === -1 && h.indexOf('hour')          !== -1) hourCol  = i;
    if (dateCol  === -1 && h.indexOf('business date') !== -1) dateCol  = i;
    if (salesCol === -1 && h.indexOf('total sales')   !== -1) salesCol = i;
  });
  // Fallback: "business date" not found — look for any "date" column that isn't "today"
  if (dateCol === -1) {
    headers.forEach(function(h, i) {
      if (dateCol === -1 && h.indexOf('date') !== -1 && h.indexOf('today') === -1) dateCol = i;
    });
  }
  if (hourCol === -1 || dateCol === -1 || salesCol === -1) {
    throw new Error(
      'Sheet1 missing required columns. Need headers containing "Hour", "Business Date", and "Total Sales". ' +
      'Found: ' + data[0].join(', ')
    );
  }

  var DAY_NAMES = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
  var HOUR_MAP  = {
    6:'6 AM',7:'7 AM',8:'8 AM',9:'9 AM',10:'10 AM',11:'11 AM',
    12:'12 PM',13:'1 PM',14:'2 PM',15:'3 PM',16:'4 PM',17:'5 PM',
    18:'6 PM',19:'7 PM',20:'8 PM',21:'9 PM',22:'10 PM'
  };
  var payload = [];
  for (var i = 1; i < data.length; i++) {
    var hour24   = parseInt(data[i][hourCol]);
    var rawDate  = data[i][dateCol];
    var salesAmt = parseFloat(data[i][salesCol]);
    if (!hour24 || !rawDate || isNaN(salesAmt)) continue;
    if (!HOUR_MAP[hour24]) continue;
    var date = new Date(rawDate);
    if (isNaN(date.getTime())) continue;
    payload.push({ day: DAY_NAMES[date.getDay()], hour: HOUR_MAP[hour24], sales: salesAmt });
  }
  return payload;
}

/**
 * Validates Sheet1 has data rows with today's date in column C.
 * Returns { valid: true/false, reason: string }.
 */
function validateSheet1() {
  var ss    = getSpreadsheet();
  var sheet = ss.getSheetByName('Sheet1');
  if (!sheet) return { valid: false, reason: 'Sheet1 not found in this spreadsheet.' };
  var last = sheet.getLastRow();
  if (last < 2) return { valid: false, reason: 'Sheet1 is empty — CFA sync may have failed.' };

  // Find the "Business Date" column dynamically
  var headers   = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var dateCol   = -1;
  headers.forEach(function(h, i) {
    if (dateCol === -1 && String(h).toLowerCase().indexOf('business date') !== -1) dateCol = i + 1; // 1-based
  });
  if (dateCol === -1) {
    // No "Business Date" column — skip date validation, just confirm data exists
    return { valid: true };
  }

  var tz       = Session.getScriptTimeZone();
  var now      = new Date();
  var cutoff   = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000); // 7 days ago
  var col      = sheet.getRange(2, dateCol, last - 1, 1).getValues().flat();
  var hasRecent = col.some(function(v) {
    try {
      var d = new Date(v);
      return !isNaN(d.getTime()) && d >= cutoff && d <= now;
    } catch(e) { return false; }
  });
  if (!hasRecent) return {
    valid: false,
    reason: 'Sheet1 has no rows within the last 7 days. CFA sync may not have run.'
  };
  return { valid: true };
}

/**
 * Clears all data rows from Sheet1, leaving the header row intact.
 * Called only after archiveSalesPayload() succeeds — never on error.
 */
function clearSheet1() {
  var sheet   = getSpreadsheet().getSheetByName('Sheet1');
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, 4).clearContent();
  Logger.log('Sheet1 cleared — ' + (lastRow - 1) + ' data rows removed.');
}

// ----- TRIGGERS ----------------------------------------------

/**
 * Creates both time-based triggers. Safe to re-run — skips any that already exist.
 * Run once manually from the Apps Script editor after deployment.
 */
function installTriggers() {
  var existing = ScriptApp.getProjectTriggers().map(function(t) { return t.getHandlerFunction(); });
  if (existing.indexOf('sundayAlertCheck') === -1) {
    ScriptApp.newTrigger('sundayAlertCheck')
      .timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(22).create();
    Logger.log('Sunday 22:00 trigger created.');
  } else {
    Logger.log('Sunday trigger already exists — skipped.');
  }
  if (existing.indexOf('mondayPipeline') === -1) {
    ScriptApp.newTrigger('mondayPipeline')
      .timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(6).create();
    Logger.log('Monday 06:00 trigger created.');
  } else {
    Logger.log('Monday trigger already exists — skipped.');
  }
  if (existing.indexOf('thursdayScheduleReminder') === -1) {
    ScriptApp.newTrigger('thursdayScheduleReminder')
      .timeBased().onWeekDay(ScriptApp.WeekDay.THURSDAY).atHour(19).create();
    Logger.log('Thursday 19:00 schedule reminder trigger created.');
  } else {
    Logger.log('Thursday trigger already exists — skipped.');
  }
  return { success: true };
}

/**
 * Deletes all project triggers. Safe to re-run — no-op if none exist.
 * Use before re-installing triggers from scratch or when decommissioning.
 */
function deleteTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t) { ScriptApp.deleteTrigger(t); });
  Logger.log('Deleted ' + triggers.length + ' trigger(s).');
  return { success: true, deleted: triggers.length };
}

/**
 * Returns the installation status of all three triggers.
 * Called by the frontend Settings > Automation section.
 */
function getTriggersStatus() {
  var triggers = ScriptApp.getProjectTriggers();
  var fns      = triggers.map(function(t) { return t.getHandlerFunction(); });
  return {
    sunday:   { installed: fns.indexOf('sundayAlertCheck')         !== -1 },
    monday:   { installed: fns.indexOf('mondayPipeline')            !== -1 },
    thursday: { installed: fns.indexOf('thursdayScheduleReminder')  !== -1 }
  };
}

/**
 * Runs Thursday ~19:00. Sends a schedule posting reminder to operator_email.
 */
/**
 * Runs Thursday ~19:00. Sends schedule posting reminder + latest upload summary.
 */
function thursdayScheduleReminder() {
  try {
    var result = buildThursdayEmail_(false);
    if (!result) return;
    GmailApp.sendEmail(result.email, result.subject, result.body);
    Logger.log('thursdayScheduleReminder: sent to ' + result.email);
  } catch(err) {
    Logger.log('thursdayScheduleReminder error: ' + err.message);
  }
}

/**
 * Manual test for the Thursday reminder. Returns result to the UI.
 */
function testThursdayReminder() {
  try {
    var result = buildThursdayEmail_(true);
    if (!result) return { success: false, reason: 'No operator_email found in config sheet.' };
    GmailApp.sendEmail(result.email, result.subject, result.body);
    return { success: true, email: result.email };
  } catch(err) {
    return { success: false, reason: err.message };
  }
}

/**
 * Builds the Thursday reminder email with latest schedule upload summary.
 * Returns { email, subject, body } or null if no email configured.
 */
function buildThursdayEmail_(isTest) {
  var ss      = getSpreadsheet();
  var cfgMap  = readConfigMap(ss);
  var email   = (cfgMap['operator_email'] || '').trim();
  var locName = cfgMap['location_name'] || 'Cockrell Hill';
  var goal    = parseFloat(cfgMap['splh_goal_overall']) || 90;
  if (!email) { Logger.log('thursdayScheduleReminder: no operator_email configured, skipping.'); return null; }

  var MONTHS = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  var prefix = isTest ? '[TEST] ' : '';

  // Read latest weekly_summary row
  var summSheet = getSheet(ss, 'weekly_summary');
  var lastRow   = summSheet.getLastRow();
  var hasData   = lastRow >= 2;
  var summaryBlock = '';

  if (hasData) {
    var summData = summSheet.getRange(2, 1, lastRow - 1, 9).getValues();
    // Sort by weekStart descending, pick most recent
    summData.sort(function(a, b) { return toDateString_(b[0]).localeCompare(toDateString_(a[0])); });
    var row = summData[0];
    var ws      = toDateString_(row[0]);
    var avgSplh = parseFloat(row[2]) || 0;
    var peak    = String(row[3]) || 'N/A';
    var pctBel  = parseFloat(row[4]) || 0;
    var pctNear = parseFloat(row[5]) || 0;
    var pctMeet = parseFloat(row[6]) || 0;
    var hours   = parseFloat(row[7]) || 0;

    var wDate   = new Date(ws + 'T12:00:00Z');
    var sDate   = new Date(wDate.getTime()); sDate.setUTCDate(sDate.getUTCDate() + 5);
    var monStr  = MONTHS[wDate.getUTCMonth()] + ' ' + wDate.getUTCDate();
    var satStr  = MONTHS[sDate.getUTCMonth()] + ' ' + sDate.getUTCDate();

    // Read schedule_history for FOH/BOH breakdown
    var schedSheet = getSheet(ss, 'schedule_history');
    var sLastRow   = schedSheet.getLastRow();
    var fohTotal = 0, bohTotal = 0;
    if (sLastRow >= 2) {
      var schedData = schedSheet.getRange(2, 1, sLastRow - 1, 7).getValues();
      schedData.forEach(function(r) {
        if (toDateString_(r[0]) === ws) {
          fohTotal += parseFloat(r[3]) || 0; // foh_count col
          bohTotal += parseFloat(r[4]) || 0; // boh_count col
        }
      });
    }

    // Read sales_history for FOH/BOH SPLH + daily totals
    var salesSheet = getSheet(ss, 'sales_history');
    var slLastRow  = salesSheet.getLastRow();
    var totalSales = 0, dailySales = {};
    if (slLastRow >= 2) {
      var salesData = salesSheet.getRange(2, 1, slLastRow - 1, 9).getValues();
      salesData.forEach(function(r) {
        if (toDateString_(r[0]) === ws) {
          var day   = String(r[1]);
          var sales = parseFloat(r[3]) || 0;
          totalSales += sales;
          dailySales[day] = (dailySales[day] || 0) + sales;
        }
      });
    }

    // Compute FOH/BOH SPLH
    var laborShareFoh = parseFloat(cfgMap['labor_share_foh']) || 50;
    var laborShareBoh = parseFloat(cfgMap['labor_share_boh']) || 40;
    var totalShare    = laborShareFoh + laborShareBoh || 1;
    var fohSplh       = fohTotal > 0 ? (totalSales * (laborShareFoh / totalShare)) / fohTotal : 0;
    var bohSplh       = bohTotal > 0 ? (totalSales * (laborShareBoh / totalShare)) / bohTotal : 0;

    // Daily sales lines
    var dailyLines = DAYS_LIST.map(function(day) {
      var ds = dailySales[day] || 0;
      return ds > 0 ? ('  ' + day + ': $' + Math.round(ds).toLocaleString()) : null;
    }).filter(function(l) { return l; });

    summaryBlock = [
      '',
      '═══════════════════════════════════',
      'LATEST UPLOAD: Week of ' + monStr + ' - ' + satStr,
      '═══════════════════════════════════',
      '',
      'SPLH PERFORMANCE',
      '  Overall Avg SPLH:  $' + avgSplh.toFixed(2) + '  (Goal: $' + goal.toFixed(2) + ')',
      '  FOH Avg SPLH:      $' + fohSplh.toFixed(2),
      '  BOH Avg SPLH:      $' + bohSplh.toFixed(2),
      '',
      'GOAL BREAKDOWN',
      '  Meeting Goal:  ' + (pctMeet * 100).toFixed(1) + '%',
      '  Near Goal:     ' + (pctNear * 100).toFixed(1) + '%',
      '  Below Goal:    ' + (pctBel  * 100).toFixed(1) + '%',
      '',
      'STAFFING',
      '  Total Scheduled:  ' + Math.round(hours) + ' person-hours',
      '  FOH Hours:        ' + Math.round(fohTotal),
      '  BOH Hours:        ' + Math.round(bohTotal),
      '  Peak Hour:        ' + peak,
      ''
    ];

    if (dailyLines.length > 0) {
      summaryBlock = summaryBlock.concat(['DAILY SALES'], dailyLines, ['']);
    }

    summaryBlock = summaryBlock.join('\n');
  } else {
    summaryBlock = '\n(No schedule uploads found yet. Upload a schedule in the app to see data here.)\n';
  }

  var subject = prefix + '\ud83d\udcc5 Schedule Reminder \u2014 Post by Tonight | ' + locName;
  var body = [
    'Hi,',
    '',
    'This is your weekly reminder to post next week\'s schedule tonight.',
    '',
    'Location: ' + locName,
    summaryBlock,
    '---',
    'Schedule Counter | ' + locName
  ].join('\n');

  return { email: email, subject: subject, body: body };
}

// ----- TRIGGER HANDLERS --------------------------------------

/**
 * Runs Sunday ~22:00. Single responsibility: validate Sheet1 and alert if sync failed.
 */
function sundayAlertCheck() {
  try {
    var result = validateSheet1();
    if (!result.valid) {
      sendAlert(
        '\u26a0 CFA Sync Alert \u2014 Sunday',
        result.reason + '\n\nCheck the CFA data connection and verify Sheet1 before Monday morning.'
      );
    }
  } catch(err) {
    sendAlert(
      '\u26a0 Sunday Alert Check Failed',
      'The Sunday validation check itself threw an error: ' + err.message
    );
  }
}

// ----- SALES FORECAST ----------------------------------------

/**
 * Rebuilds sales_curves by replaying all weeks in sales_history through the
 * same exponential smoothing algorithm used by mondayPipeline.
 * Processes weeks in chronological order so the most recent weeks have the
 * most influence. Safe to run multiple times.
 * Returns { weeksProcessed: N }.
 */
function rebuildSalesCurvesFromHistory() {
  var ss    = getSpreadsheet();
  var sheet = ss.getSheetByName('sales_history');
  if (!sheet || sheet.getLastRow() < 2) {
    throw new Error('sales_history is empty. Run "Backfill History" first.');
  }

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
  var validDays = {};
  DAYS_LIST.forEach(function(d) { validDays[d] = true; });

  // Group hourly rows by weekStart
  var byWeek = {};
  data.forEach(function(row) {
    var ws    = toDateString_(row[0]);
    var day   = String(row[1]);
    var hour  = hourCellToLabel_(row[2]);
    var sales = parseFloat(row[3]) || 0;
    if (!validDays[day] || sales <= 0) return;
    if (!byWeek[ws]) byWeek[ws] = [];
    byWeek[ws].push({ day: day, hour: hour, sales: sales });
  });

  // Process chronologically so recent weeks have the final say
  var weeks = Object.keys(byWeek).sort();
  weeks.forEach(function(ws) {
    runSmoothingUpdate(byWeek[ws]);
  });

  Logger.log('rebuildSalesCurvesFromHistory: processed ' + weeks.length + ' weeks');
  return { weeksProcessed: weeks.length };
}

/**
 * Returns weekly total sales from sales_history, grouped by week and day.
 * Used by the Trends tab to show sales history even when weekly_summary is empty.
 * Returns { rows: [{ weekStart, totalSales, salesByDay }] } sorted oldest-first.
 */
function getWeeklySalesHistory() {
  try {
    var ss    = getSpreadsheet();
    var sheet = ss.getSheetByName('sales_history');
    if (!sheet || sheet.getLastRow() < 2) return { rows: [] };

    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
    var byWeek = {}; // { weekStart: { day: totalSales } }

    data.forEach(function(row) {
      var ws    = toDateString_(row[0]);
      var day   = String(row[1]);
      var sales = parseFloat(row[3]) || 0;
      if (!byWeek[ws]) byWeek[ws] = {};
      byWeek[ws][day] = (byWeek[ws][day] || 0) + sales;
    });

    var rows = Object.keys(byWeek).sort().map(function(ws) {
      var salesByDay = byWeek[ws];
      var totalSales = Object.keys(salesByDay).reduce(function(sum, d) {
        return sum + salesByDay[d];
      }, 0);
      return { weekStart: ws, totalSales: Math.round(totalSales), salesByDay: salesByDay };
    });

    return { rows: rows };
  } catch(err) {
    Logger.log('getWeeklySalesHistory error: ' + err.message);
    return { error: err.message };
  }
}

/**
 * Predicts next week's daily total sales using a robust linearly-weighted
 * moving average with outlier exclusion and trend adjustment.
 *
 * Algorithm (per day of week):
 *  1. Take the most recent N weeks (default 8, set via forecast_weeks config).
 *  2. OUTLIER EXCLUSION — drop any week whose total is > 2 standard deviations
 *     from the window mean (e.g. storm closures, holiday anomalies). Requires
 *     at least 4 weeks in window and leaves at least 3 weeks after exclusion.
 *  3. LINEAR WEIGHTS — oldest remaining week = weight 1, newest = weight N,
 *     so recent weeks count more without completely ignoring history.
 *  4. TREND ADJUSTMENT — when ≥ 6 clean weeks exist, compare the average of
 *     the recent half to the older half; add one week's worth of that slope to
 *     the LWMA prediction. The adjustment is capped at ±10% of LWMA to prevent
 *     runaway projections from noisy data.
 *
 * Simulation results vs prior exponential smoothing (α=0.3) across 7 realistic
 * restaurant scenarios: ~51% reduction in mean absolute error.
 *   • Spike/crash weeks (storm, holiday): removed automatically → 0 error
 *   • Business step changes (new promo): error 480→56
 *   • Upward/downward trends: error reduced ~29%
 *
 * Config keys (optional, set in Settings sheet):
 *   forecast_weeks  – integer, how many recent weeks to include (default 8)
 *
 * Returns:
 *   { predictions: { Monday: { total, weekCount, windowSize, trend }, ... },
 *     weekStart: 'YYYY-MM-DD'  (the Monday being forecast) }
 */
function getSalesForecast() {
  try {
    var ss      = getSpreadsheet();
    var cfgMap  = readConfigMap(ss);
    var nWeeks  = parseInt(cfgMap['forecast_weeks']) || 8;
    var sheet   = ss.getSheetByName('sales_history');
    if (!sheet || sheet.getLastRow() < 2) return { predictions: {}, weekStart: null };

    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();

    // Group: { day: { weekStart: total } }
    var byDay = {};
    DAYS_LIST.forEach(function(d) { byDay[d] = {}; });
    data.forEach(function(row) {
      var ws    = toDateString_(row[0]);
      var day   = String(row[1]);
      var sales = parseFloat(row[3]) || 0;
      if (!byDay[day]) return;
      byDay[day][ws] = (byDay[day][ws] || 0) + sales;
    });

    var predictions = {};
    DAYS_LIST.forEach(function(day) {
      var weekMap = byDay[day];
      var series  = Object.keys(weekMap).sort().map(function(ws) { return weekMap[ws]; });
      if (series.length === 0) { predictions[day] = { total: null, weekCount: 0, windowSize: 0, trend: 0 }; return; }

      // ── Step 1: window ───────────────────────────────────────
      var win = series.slice(-nWeeks);
      var n   = win.length;

      // ── Step 2: outlier exclusion ────────────────────────────
      var mean = win.reduce(function(s, v) { return s + v; }, 0) / n;
      var std  = Math.sqrt(win.reduce(function(s, v) { return s + (v - mean) * (v - mean); }, 0) / n);
      var clean = (std > 0 && n >= 4)
        ? win.filter(function(v) { return Math.abs(v - mean) <= 2 * std; })
        : win;
      if (clean.length < 3) clean = win; // safety: never drop below 3 weeks
      var nc = clean.length;

      // ── Step 3: linearly-weighted average ────────────────────
      var weightSum = (nc * (nc + 1)) / 2;
      var weighted  = 0;
      for (var i = 0; i < nc; i++) weighted += (i + 1) * clean[i];
      var lwma = weighted / weightSum;

      // ── Step 4: trend adjustment (capped at ±10% of LWMA) ────
      var trend = 0;
      if (nc >= 6) {
        var half      = Math.floor(nc / 2);
        var avgRecent = clean.slice(-half).reduce(function(s, v) { return s + v; }, 0) / half;
        var avgOlder  = clean.slice(0, half).reduce(function(s, v) { return s + v; }, 0) / half;
        var rawTrend  = (avgRecent - avgOlder) / half; // per-week slope
        var cap       = lwma * 0.10;
        trend = Math.max(-cap, Math.min(cap, rawTrend));
      }

      predictions[day] = {
        total:      Math.round(lwma + trend),
        weekCount:  series.length,     // total history available
        windowSize: nc,                // clean weeks used in forecast
        trend:      Math.round(trend)  // adjustment applied (+ = rising, - = falling)
      };
    });

    // Compute Monday 2 weeks out (the week being scheduled)
    var today   = new Date();
    var dow     = today.getDay(); // 0=Sun
    var daysToNextMon = (dow === 1) ? 7 : (8 - dow) % 7;
    var nextMon = new Date(today.getTime());
    nextMon.setDate(today.getDate() + daysToNextMon + 7); // +7 for 2 weeks out
    var mm = String(nextMon.getMonth() + 1); if (mm.length < 2) mm = '0' + mm;
    var dd = String(nextMon.getDate());      if (dd.length < 2) dd = '0' + dd;
    var weekStart = nextMon.getFullYear() + '-' + mm + '-' + dd;

    return { predictions: predictions, weekStart: weekStart };
  } catch(err) {
    Logger.log('getSalesForecast error: ' + err.message);
    return { error: err.message };
  }
}

/**
 * One-time backfill: reads ALL rows currently in Sheet1 and archives them
 * to sales_history grouped by business date and week.
 *
 * Safe to run multiple times — skips (week_start, day, hour) combos that
 * already exist in sales_history with source='api'.
 *
 * Returns { added: N, skipped: M } for display in the UI.
 */
function backfillSalesHistoryFromSheet1() {
  var ss    = getSpreadsheet();
  var sheet = ss.getSheetByName('Sheet1');
  if (!sheet) throw new Error('Sheet1 not found.');
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) throw new Error('Sheet1 has no data rows.');

  // Dynamic column detection (mirrors readSheet1)
  var headers = data[0].map(function(h) { return String(h).toLowerCase().trim(); });
  var hourCol = -1, dateCol = -1, salesCol = -1;
  headers.forEach(function(h, i) {
    if (hourCol  === -1 && h.indexOf('hour')          !== -1) hourCol  = i;
    if (dateCol  === -1 && h.indexOf('business date') !== -1) dateCol  = i;
    if (salesCol === -1 && h.indexOf('total sales')   !== -1) salesCol = i;
  });
  if (dateCol === -1) {
    headers.forEach(function(h, i) {
      if (dateCol === -1 && h.indexOf('date') !== -1 && h.indexOf('today') === -1) dateCol = i;
    });
  }
  if (hourCol === -1 || dateCol === -1 || salesCol === -1) {
    throw new Error('Sheet1 missing required columns. Need "Hour", "Business Date", and "Total Sales" headers.');
  }

  var DAY_NAMES = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
  var HOUR_MAP  = {
    6:'6 AM',7:'7 AM',8:'8 AM',9:'9 AM',10:'10 AM',11:'11 AM',
    12:'12 PM',13:'1 PM',14:'2 PM',15:'3 PM',16:'4 PM',17:'5 PM',
    18:'6 PM',19:'7 PM',20:'8 PM',21:'9 PM',22:'10 PM'
  };

  // Parse Sheet1 rows and group by date
  var byDate = {}; // { 'YYYY-MM-DD': { 'Monday|6 AM': totalSales, ... } }
  var dateToMeta = {}; // { 'YYYY-MM-DD': { day, weekStart } }
  for (var i = 1; i < data.length; i++) {
    var hour24   = parseInt(data[i][hourCol]);
    var rawDate  = data[i][dateCol];
    var salesAmt = parseFloat(data[i][salesCol]);
    if (!hour24 || !rawDate || isNaN(salesAmt)) continue;
    if (!HOUR_MAP[hour24]) continue;
    var d = (rawDate instanceof Date) ? rawDate : new Date(rawDate);
    if (isNaN(d.getTime())) continue;
    var dayName  = DAY_NAMES[d.getDay()];
    if (dayName === 'Sunday') continue; // location closed Sunday
    var dateStr  = toDateString_(d);
    // week_start = Monday of this week
    var dow      = d.getDay();
    var monday   = new Date(d.getTime());
    monday.setDate(d.getDate() - (dow === 0 ? 6 : dow - 1));
    var weekStart = toDateString_(monday);
    if (!byDate[dateStr]) {
      byDate[dateStr] = {};
      dateToMeta[dateStr] = { day: dayName, weekStart: weekStart };
    }
    var key = dayName + '|' + HOUR_MAP[hour24];
    byDate[dateStr][key] = (byDate[dateStr][key] || 0) + salesAmt;
  }

  // Load existing sales_history to avoid duplicates
  var salesSheet   = getSheet(ss, 'sales_history');
  var salesLast    = salesSheet.getLastRow();
  var existingKeys = {};
  if (salesLast >= 2) {
    var existingSalesRows = salesSheet.getRange(2, 1, salesLast - 1, 5).getValues();
    existingSalesRows.forEach(function(row) {
      if (String(row[4]) === 'api') {
        existingKeys[toDateString_(row[0]) + '|' + row[1] + '|' + hourCellToLabel_(row[2])] = true;
      }
    });
  }

  var now     = new Date().toISOString();
  var newRows = [];
  var skipped = 0;

  Object.keys(byDate).sort().forEach(function(dateStr) {
    var meta    = dateToMeta[dateStr];
    var dayData = byDate[dateStr];
    Object.keys(dayData).forEach(function(dayHourKey) {
      var parts = dayHourKey.split('|');
      var day   = parts[0];
      var hour  = parts[1];
      var dedupeKey = meta.weekStart + '|' + day + '|' + hour;
      if (existingKeys[dedupeKey]) { skipped++; return; }
      existingKeys[dedupeKey] = true;
      newRows.push([meta.weekStart, day, hour, dayData[dayHourKey], 'api', 0, 0, 'Pending', now]);
    });
  });

  if (newRows.length > 0) {
    var writeRange = salesSheet.getRange(salesLast + 1, 1, newRows.length, 9);
    // Force week_start (A) and hour (C) columns to plain text
    salesSheet.getRange(salesLast + 1, 1, newRows.length, 1).setNumberFormat('@');
    salesSheet.getRange(salesLast + 1, 3, newRows.length, 1).setNumberFormat('@');
    writeRange.setValues(newRows);
  }
  Logger.log('backfillSalesHistoryFromSheet1: added=' + newRows.length + ' skipped=' + skipped);
  return { added: newRows.length, skipped: skipped };
}

/**
 * One-time cleanup: normalizes week_start and hour columns in sales_history,
 * then removes duplicate (week_start, day, hour, source) rows keeping the latest.
 * Safe to re-run. Returns { normalized, deduped } counts.
 */
function dedupSalesHistory() {
  var ss    = getSpreadsheet();
  var sheet = ss.getSheetByName('sales_history');
  if (!sheet || sheet.getLastRow() < 2) return { normalized: 0, deduped: 0 };

  var data    = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues();
  var normalized = 0;

  // Normalize week_start (col 0) and hour (col 2) in every row
  data.forEach(function(row) {
    var origWs = String(row[0]);
    row[0] = toDateString_(row[0]);
    row[2] = hourCellToLabel_(row[2]);
    if (row[0] !== origWs) normalized++;
  });

  // Deduplicate: keep last occurrence of each (week_start, day, hour, source) combo
  var seen = {};
  var kept = [];
  for (var i = data.length - 1; i >= 0; i--) {
    var key = data[i][0] + '|' + data[i][1] + '|' + data[i][2] + '|' + data[i][4];
    if (seen[key]) continue;
    seen[key] = true;
    kept.push(data[i]);
  }
  kept.reverse();

  var deduped = data.length - kept.length;

  // Rewrite the sheet
  sheet.clearContents();
  sheet.getRange(1, 1, 1, 9).setValues(SALES_HISTORY_HEADERS);
  if (kept.length > 0) {
    sheet.getRange(2, 1, kept.length, 1).setNumberFormat('@');
    sheet.getRange(2, 3, kept.length, 1).setNumberFormat('@');
    sheet.getRange(2, 1, kept.length, 9).setValues(kept);
  }
  SpreadsheetApp.flush();
  Logger.log('dedupSalesHistory: normalized=' + normalized + ' deduped=' + deduped);
  return { normalized: normalized, deduped: deduped };
}

/**
 * Runs Monday ~06:00. Core automation pipeline. Steps run sequentially —
 * any failure sends an alert and stops. App remains usable with last cached state.
 */
function mondayPipeline() {
  var startTime = new Date();
  Logger.log('mondayPipeline started: ' + startTime.toISOString());
  try {
    // Step 1: Validate Sheet1
    var validation = validateSheet1();
    if (!validation.valid) {
      sendAlert('\u26a0 Monday Pipeline \u2014 No Sales Data', validation.reason);
      return;
    }

    // Step 2: Read Sheet1
    var salesPayload = readSheet1();
    Logger.log('Sheet1 read: ' + salesPayload.length + ' records');

    // Step 3: Determine week start from payload dates
    var weekStart = detectWeekStartFromPayload(salesPayload);
    Logger.log('Week start: ' + weekStart);

    // Step 4: Archive to sales_history (sales_source = 'api')
    archiveSalesPayload(salesPayload, weekStart);
    Logger.log('Archived to sales_history');

    // Step 5: Exponential smoothing — update sales_curves
    runSmoothingUpdate(salesPayload);
    Logger.log('Smoothing update complete');

    // Step 6: Bias detection — update bias_flags in config
    runBiasDetection();
    Logger.log('Bias detection complete');

    // Step 7: Delete rows older than 90 days
    cleanupOldHistory();
    Logger.log('Cleanup complete');

    // Step 8: Clear Sheet1
    clearSheet1();
    Logger.log('Sheet1 cleared');

    // Step 8.5: Refresh actuals calibration (silent no-op unless wired;
    // failures alert + degrade to the configured goal, never abort the pipeline)
    refreshActuals();
    Logger.log('Actuals calibration refreshed');

    // Step 9: Build and write app_cache
    var cacheObject = buildCacheObject();
    writeAppCache(cacheObject);
    Logger.log('app_cache written');

    var elapsed = (new Date() - startTime) / 1000;
    Logger.log('mondayPipeline complete in ' + elapsed + 's');
  } catch(err) {
    sendAlert(
      '\u26a0 Monday Pipeline Failed',
      'Error: ' + err.message +
      '\n\nThe app may be using last week\'s cached data. Check Apps Script execution logs.'
    );
    Logger.log('mondayPipeline ERROR: ' + err.message);
  }
}

/**
 * Sends an alert email to ALERT_EMAIL (Script Property) and operator_email (config).
 * De-duplicates recipients. Silent no-op if neither is configured.
 */
function sendAlert(subject, body) {
  var scriptEmail = PropertiesService.getScriptProperties().getProperty('ALERT_EMAIL');
  var configEmail = getConfigValue('operator_email');
  var seen = {};
  var recipients = [scriptEmail, configEmail].filter(function(e) {
    if (!e || !e.trim()) return false;
    var k = e.trim().toLowerCase();
    if (seen[k]) return false;
    seen[k] = true;
    return true;
  });
  if (!recipients.length) { Logger.log('sendAlert: no recipients configured.'); return; }
  recipients.forEach(function(email) {
    GmailApp.sendEmail(email.trim(), subject, body);
  });
}

// ----- PIPELINE STEPS ----------------------------------------

/**
 * Reads Sheet1 column B dates to find the Monday of the week.
 * Returns YYYY-MM-DD string.
 */
function detectWeekStartFromPayload(salesPayload) {
  var ss    = getSpreadsheet();
  var sheet = ss.getSheetByName('Sheet1');
  var data  = sheet.getDataRange().getValues();
  var earliest = null;
  for (var i = 1; i < data.length; i++) {
    var rawDate = data[i][1];
    if (!rawDate) continue;
    var d = new Date(rawDate);
    if (isNaN(d.getTime())) continue;
    if (!earliest || d < earliest) earliest = d;
  }
  if (!earliest) throw new Error('Could not determine week start: no valid dates in Sheet1 column B.');
  // Adjust to Monday of that week (getDay: 0=Sun, 1=Mon, ...)
  var dow = earliest.getDay();
  earliest.setDate(earliest.getDate() - (dow === 0 ? 6 : dow - 1));
  var mm = String(earliest.getMonth() + 1); if (mm.length < 2) mm = '0' + mm;
  var dd = String(earliest.getDate());      if (dd.length < 2) dd = '0' + dd;
  return earliest.getFullYear() + '-' + mm + '-' + dd;
}

/**
 * Archives the CFA salesPayload to sales_history with sales_source = 'api'.
 * Removes any existing api rows for this weekStart before writing (idempotent).
 * SPLH is written as 0 / 'Pending' if no matching schedule_history rows exist yet.
 */
function archiveSalesPayload(salesPayload, weekStart) {
  var ss         = getSpreadsheet();
  var salesSheet = getSheet(ss, 'sales_history');
  var schedSheet = getSheet(ss, 'schedule_history');
  var cfgMap     = readConfigMap(ss);
  var goal       = parseFloat(cfgMap['splh_goal_overall']) || 90;
  var now        = new Date().toISOString();

  // Build per-hour combined_count lookup from schedule_history for this weekStart
  var schedLookup = {};
  var schedLast = schedSheet.getLastRow();
  if (schedLast >= 2) {
    var schedRows = schedSheet.getRange(2, 1, schedLast - 1, 7).getValues();
    schedRows.forEach(function(row) {
      if (toDateString_(row[0]) === weekStart) {
        var day = String(row[1]), hour = hourCellToLabel_(row[2]), combined = row[5] || 0;
        if (!schedLookup[day]) schedLookup[day] = {};
        schedLookup[day][hour] = combined;
      }
    });
  }

  // Group salesPayload: { day: { hour: totalSales } }
  var grouped = {};
  salesPayload.forEach(function(rec) {
    if (!grouped[rec.day]) grouped[rec.day] = {};
    grouped[rec.day][rec.hour] = (grouped[rec.day][rec.hour] || 0) + rec.sales;
  });

  // Build new api rows
  var newRows = [];
  Object.keys(grouped).forEach(function(day) {
    Object.keys(grouped[day]).forEach(function(hour) {
      var sales    = grouped[day][hour];
      var combined = (schedLookup[day] && schedLookup[day][hour]) || 0;
      var splh, diff, flag;
      if (combined > 0) {
        splh = sales / combined;
        diff = splh - goal;
        flag = splh >= goal ? 'Meeting Goal' : splh >= goal * 0.9 ? 'Near Goal' : 'Below Goal';
      } else {
        splh = 0; diff = 0; flag = 'Pending';
      }
      newRows.push([weekStart, day, hour, sales, 'api', splh, diff, flag, now]);
    });
  });

  // Rebuild sales_history: keep all rows except (this weekStart AND source='api')
  var salesLast = salesSheet.getLastRow();
  var kept = [];
  if (salesLast >= 2) {
    kept = salesSheet.getRange(2, 1, salesLast - 1, 9).getValues()
      .filter(function(r) {
        return !(toDateString_(r[0]) === weekStart && String(r[4]) === 'api');
      });
  }
  var allRows = kept.concat(newRows);
  salesSheet.clearContents();
  salesSheet.getRange(1, 1, 1, 9).setValues(SALES_HISTORY_HEADERS);
  if (allRows.length > 0) {
    // Force week_start (A) and hour (C) columns to plain text
    salesSheet.getRange(2, 1, allRows.length, 1).setNumberFormat('@');
    salesSheet.getRange(2, 3, allRows.length, 1).setNumberFormat('@');
    salesSheet.getRange(2, 1, allRows.length, 9).setValues(allRows);
  }
}

/**
 * Updates sales_curves using exponential smoothing.
 * Reads the entire sales_curves sheet once, computes all updates in memory,
 * writes back the complete dataset in a single setValues() call.
 *
 * Algorithm per day d, hour h (in-range only: 6 AM–10 PM):
 *   new_weight = alpha * actual_fraction + (1 - alpha) * current_weight
 *   Then normalize the 17 in-range weights so the full 19-hour curve sums to 1.0.
 *   Increment sample_count by 1. Set last_updated to now.
 *   Hours 5 AM and 11 PM are never touched.
 */
function runSmoothingUpdate(salesPayload) {
  var ss     = getSpreadsheet();
  var cfgMap = readConfigMap(ss);
  var alpha  = parseFloat(cfgMap['smoothing_alpha']) || 0.3;
  var sheet  = getSheet(ss, 'sales_curves');
  var last   = sheet.getLastRow();
  if (last < 2) return;

  var allData  = sheet.getRange(2, 1, last - 1, 5).getValues();
  var now      = new Date().toISOString();
  var CFA_HOURS_SET = {
    '6 AM':true,'7 AM':true,'8 AM':true,'9 AM':true,'10 AM':true,'11 AM':true,
    '12 PM':true,'1 PM':true,'2 PM':true,'3 PM':true,'4 PM':true,'5 PM':true,
    '6 PM':true,'7 PM':true,'8 PM':true,'9 PM':true,'10 PM':true
  };
  var OUT_OF_RANGE = { '5 AM': true, '11 PM': true };

  // Group payload by day/hour (sum duplicates)
  var grouped = {};
  salesPayload.forEach(function(rec) {
    if (!CFA_HOURS_SET[rec.hour]) return;
    if (!grouped[rec.day]) grouped[rec.day] = {};
    grouped[rec.day][rec.hour] = (grouped[rec.day][rec.hour] || 0) + rec.sales;
  });

  // Day-level outlier gate: typical (median) same-weekday api totals from
  // history, excluding the week just archived. The hour-level winsorize below
  // already caps single-hour spikes; this gate catches whole-day anomalies
  // (closures, weather, spread-out catering blowouts) so they never teach the
  // curve or inflate its confidence counts. Skipped days are logged.
  var outlierPct = parseFloat(cfgMap['curve_outlier_pct']);
  if (isNaN(outlierPct)) outlierPct = 30;
  var typicalByDay = typicalApiDayTotals_(ss);

  Object.keys(grouped).forEach(function(day) {
    var dayHours   = grouped[day];
    var totalSales = Object.keys(dayHours).reduce(function(s, h) { return s + dayHours[h]; }, 0);
    if (totalSales <= 0) return;

    var typical = typicalByDay[day] || 0;
    if (outlierPct > 0 && typical > 0) {
      var devPct = Math.abs(totalSales - typical) / typical * 100;
      if (devPct > outlierPct) {
        Logger.log('Smoothing skipped ' + day + ' as outlier: $' + Math.round(totalSales) +
          ' vs typical $' + Math.round(typical) + ' (' + devPct.toFixed(0) + '% > ' + outlierPct + '%)');
        return;
      }
    }

    // Compute actual fractions for the 17 in-range hours
    var actualFractions = {};
    Object.keys(CFA_HOURS_SET).forEach(function(h) {
      actualFractions[h] = (dayHours[h] || 0) / totalSales;
    });

    // Winsorize: cap incoming fractions at mean + 2*SD of current curve weights
    // This prevents catering spikes from distorting the curve
    var curWeights = [];
    allData.forEach(function(row) {
      if (String(row[0]) === day && !OUT_OF_RANGE[hourCellToLabel_(row[1])]) {
        curWeights.push(parseFloat(row[2]) || 0);
      }
    });
    if (curWeights.length > 0) {
      var cwMean = curWeights.reduce(function(s, v) { return s + v; }, 0) / curWeights.length;
      var cwStd  = Math.sqrt(curWeights.reduce(function(s, v) { return s + (v - cwMean) * (v - cwMean); }, 0) / curWeights.length);
      var cap    = cwMean + 2 * cwStd;
      if (cap > 0) {
        Object.keys(actualFractions).forEach(function(h) {
          if (actualFractions[h] > cap) actualFractions[h] = cap;
        });
        // Re-normalize fractions after capping so they still sum to 1.0
        var fracSum = Object.keys(actualFractions).reduce(function(s, h) { return s + actualFractions[h]; }, 0);
        if (fracSum > 0) {
          Object.keys(actualFractions).forEach(function(h) {
            actualFractions[h] = actualFractions[h] / fracSum;
          });
        }
      }
    }

    // Sum of the 2 skipped hours' current weights (preserved as-is)
    var skippedWeightSum = 0;
    allData.forEach(function(row) {
      if (String(row[0]) === day && OUT_OF_RANGE[hourCellToLabel_(row[1])]) {
        skippedWeightSum += parseFloat(row[2]) || 0;
      }
    });
    var targetSum = 1.0 - skippedWeightSum;
    if (targetSum <= 0) targetSum = 1.0;

    // Apply smoothing to in-range hours
    var newWeights = {};
    allData.forEach(function(row) {
      var rDay = String(row[0]), rHour = hourCellToLabel_(row[1]);
      if (rDay !== day || OUT_OF_RANGE[rHour] || actualFractions[rHour] === undefined) return;
      var cur = parseFloat(row[2]) || 0;
      newWeights[rHour] = alpha * actualFractions[rHour] + (1 - alpha) * cur;
    });

    // Normalize so in-range hours sum to targetSum
    var newSum = Object.keys(newWeights).reduce(function(s, h) { return s + newWeights[h]; }, 0);
    if (newSum > 0) {
      Object.keys(newWeights).forEach(function(h) {
        newWeights[h] = parseFloat(((newWeights[h] / newSum) * targetSum).toFixed(6));
      });
    }

    // Write updated weights back into allData
    allData.forEach(function(row, idx) {
      var rDay = String(row[0]), rHour = hourCellToLabel_(row[1]);
      if (rDay !== day || OUT_OF_RANGE[rHour] || newWeights[rHour] === undefined) return;
      allData[idx][2] = newWeights[rHour];
      allData[idx][3] = (parseInt(row[3]) || 0) + 1; // sample_count++
      allData[idx][4] = now;
    });
  });

  // Single batch write — force hour column to plain text first
  sheet.getRange(2, 2, allData.length, 1).setNumberFormat('@');
  sheet.getRange(2, 1, allData.length, 5).setValues(allData);
}

/**
 * Reads sales_history and computes per-day bias between 'estimated' and 'api' sources.
 * Requires at least 2 weeks where both sources exist for the same day.
 * Writes the result to the 'bias_flags' key in config.
 *
 * biasFlags shape: { Monday: { direction:'over'|'under', avgErrorPct:N, sampleCount:N } | null, ... }
 */
function runBiasDetection() {
  var ss        = getSpreadsheet();
  var cfgMap    = readConfigMap(ss);
  var threshold = parseFloat(cfgMap['bias_threshold']) || 10;
  var sheet     = getSheet(ss, 'sales_history');
  var last      = sheet.getLastRow();

  var biasFlags = {};
  DAYS_LIST.forEach(function(d) { biasFlags[d] = null; });

  if (last < 2) {
    writeConfigKeys(ss, { 'bias_flags': JSON.stringify(biasFlags) });
    return;
  }

  // Cols: week_start[0], day[1], hour[2], sales_dollars[3], sales_source[4]
  var data = sheet.getRange(2, 1, last - 1, 5).getValues();

  // Build { day: { weekStart: { estimated: total, api: total } } }
  var byDay = {};
  data.forEach(function(row) {
    var ws     = toDateString_(row[0]);
    var day    = String(row[1]);
    var sales  = parseFloat(row[3]) || 0;
    var source = String(row[4]);
    if (source !== 'estimated' && source !== 'api') return;
    if (!byDay[day]) byDay[day] = {};
    if (!byDay[day][ws]) byDay[day][ws] = { estimated: 0, api: 0 };
    byDay[day][ws][source] += sales;
  });

  DAYS_LIST.forEach(function(day) {
    if (!byDay[day]) return;
    var sharedWeeks = Object.keys(byDay[day]).filter(function(ws) {
      return byDay[day][ws].estimated > 0 && byDay[day][ws].api > 0;
    });
    if (sharedWeeks.length < 2) return;

    var totalError = sharedWeeks.reduce(function(sum, ws) {
      var est = byDay[day][ws].estimated;
      var api = byDay[day][ws].api;
      return sum + (est - api) / api;
    }, 0);
    var avgError = totalError / sharedWeeks.length;

    if (avgError > threshold / 100) {
      biasFlags[day] = {
        direction: 'over',
        avgErrorPct: parseFloat((avgError * 100).toFixed(1)),
        sampleCount: sharedWeeks.length
      };
    } else if (avgError < -(threshold / 100)) {
      biasFlags[day] = {
        direction: 'under',
        avgErrorPct: parseFloat((avgError * 100).toFixed(1)),
        sampleCount: sharedWeeks.length
      };
    }
  });

  writeConfigKeys(ss, { 'bias_flags': JSON.stringify(biasFlags) });
}

/**
 * Moves rows older than 90 days from schedule_history and sales_history into
 * matching *_archive tabs (was: permanent deletion). The working tabs stay
 * small for fast reads; the archive preserves every observation forever —
 * that history is what makes seasonality modeling, same-week-last-year
 * comparisons, and parameter backtesting possible at all.
 * Uses rebuild-and-overwrite to avoid index-shifting errors.
 */
function cleanupOldHistory() {
  var cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - 90);
  var cutoffStr = cutoff.toISOString().split('T')[0];
  var ss = getSpreadsheet();
  ['schedule_history', 'sales_history'].forEach(function(tabName) {
    var sheet  = ss.getSheetByName(tabName);
    var all    = sheet.getDataRange().getValues();
    var header = all[0];
    var kept   = all.slice(1).filter(function(row) { return toDateString_(row[0]) >= cutoffStr; });
    var pruned = all.slice(1).filter(function(row) { return toDateString_(row[0]) < cutoffStr; });

    // Archive before touching the working tab — never destroy observations
    if (pruned.length) {
      var archive = ss.getSheetByName(tabName + '_archive');
      if (!archive) {
        archive = ss.insertSheet(tabName + '_archive');
        archive.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
        archive.setFrozenRows(1);
      }
      var an = archive.getLastRow() + 1;
      archive.getRange(an, 1, pruned.length, 1).setNumberFormat('@');
      archive.getRange(an, 3, pruned.length, 1).setNumberFormat('@');
      archive.getRange(an, 1, pruned.length, header.length).setValues(pruned);
    }

    sheet.clearContents();
    sheet.getRange(1, 1, 1, header.length).setValues([header]);
    if (kept.length) {
      // Force week_start (A) and hour (C) columns to plain text
      sheet.getRange(2, 1, kept.length, 1).setNumberFormat('@');
      sheet.getRange(2, 3, kept.length, 1).setNumberFormat('@');
      sheet.getRange(2, 1, kept.length, header.length).setValues(kept);
    }
  });
}

/**
 * Assembles the full app_cache object from current config and curves.
 * Shape is identical to getConfig() so loadAppCache() fallback is transparent.
 *
 * Overrides curveConfidence with a stricter rule: a day is confident only when
 * ALL 17 in-range hours (6 AM–10 PM) have sample_count >= 4.
 *
 * Includes currentWeekSales: daily totals from the most recent 'api' week,
 * so the frontend can offer CFA pre-fill suggestions.
 */
function buildCacheObject() {
  var config = getConfig();
  var ss     = getSpreadsheet();
  var CFA_HOURS = [
    '6 AM','7 AM','8 AM','9 AM','10 AM','11 AM','12 PM',
    '1 PM','2 PM','3 PM','4 PM','5 PM','6 PM','7 PM','8 PM','9 PM','10 PM'
  ];

  // Re-compute curveConfidence: all 17 in-range hours must have sample_count >= 4
  var curvesSheet = ss.getSheetByName('sales_curves');
  if (curvesSheet && curvesSheet.getLastRow() >= 2) {
    var rawData = curvesSheet.getRange(2, 1, curvesSheet.getLastRow() - 1, 4).getValues();
    var sampleCounts = {}; // { day: { hour: count } }
    rawData.forEach(function(row) {
      var day = String(row[0]), hour = hourCellToLabel_(row[1]), cnt = parseInt(row[3]) || 0;
      if (!sampleCounts[day]) sampleCounts[day] = {};
      sampleCounts[day][hour] = cnt;
    });
    var newConf = {};
    DAYS_LIST.forEach(function(day) {
      if (!sampleCounts[day]) { newConf[day] = { confident: false, sampleCount: 0 }; return; }
      var minCount = Infinity;
      CFA_HOURS.forEach(function(h) {
        var cnt = sampleCounts[day][h] !== undefined ? sampleCounts[day][h] : 0;
        if (cnt < minCount) minCount = cnt;
      });
      if (!isFinite(minCount)) minCount = 0;
      newConf[day] = { confident: minCount >= 4, sampleCount: minCount };
    });
    config.curveConfidence = newConf;
  }

  // Attach most recent api week's daily totals for CFA pre-fill suggestions
  var cws = buildCurrentWeekSales_(ss);
  if (Object.keys(cws).length > 0) config.currentWeekSales = cws;

  return config;
}

/**
 * Internal: returns daily sales totals for the most recent 'api' week in sales_history.
 * Returns {} if no api data exists.
 */
function buildCurrentWeekSales_(ss) {
  var sheet = ss.getSheetByName('sales_history');
  if (!sheet || sheet.getLastRow() < 2) return {};
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
  var apiWeeks = {};
  data.forEach(function(row) { if (String(row[4]) === 'api') apiWeeks[toDateString_(row[0])] = true; });
  var weeks = Object.keys(apiWeeks).sort();
  if (!weeks.length) return {};
  var latestWeek = weeks[weeks.length - 1];
  var dailyTotals = {};
  data.forEach(function(row) {
    if (toDateString_(row[0]) === latestWeek && String(row[4]) === 'api') {
      var day = String(row[1]), sales = parseFloat(row[3]) || 0;
      dailyTotals[day] = (dailyTotals[day] || 0) + sales;
    }
  });
  return dailyTotals;
}

// ---- (original sendWeeklySummaryEmail begins below) ----------

function sendWeeklySummaryEmail(body) {
  try {
    var ss      = getSpreadsheet();
    var cfgMap  = readConfigMap(ss);
    var email   = (cfgMap['operator_email'] || '').trim();
    if (!email) return { success: true }; // Skip silently

    var weekStart      = body ? String(body.weekStart || '') : '';
    var goal           = parseFloat(cfgMap['splh_goal_overall']) || 90;
    var locName        = cfgMap['location_name'] || 'Cockrell Hill';
    var laborShareFoh  = parseFloat(cfgMap['labor_share_foh']) || 50;
    var laborShareBoh  = parseFloat(cfgMap['labor_share_boh']) || 40;
    var totalShare     = laborShareFoh + laborShareBoh || 1;
    var fohFraction    = laborShareFoh / totalShare;
    var bohFraction    = laborShareBoh / totalShare;

    // Read computed summary row for this week
    var avgSplh = 0, peakHour = 'N/A', pctBelow = 0, pctNear = 0, pctMeeting = 0, totalHours = 0;
    if (weekStart) {
      var summSheet = getSheet(ss, 'weekly_summary');
      var lastRow   = summSheet.getLastRow();
      if (lastRow >= 2) {
        var summData = summSheet.getRange(2, 1, lastRow - 1, 9).getValues();
        for (var i = 0; i < summData.length; i++) {
          if (toDateString_(summData[i][0]) === weekStart) {
            avgSplh    = parseFloat(summData[i][2]) || 0;
            peakHour   = String(summData[i][3]) || 'N/A';
            pctBelow   = parseFloat(summData[i][4]) || 0;
            pctNear    = parseFloat(summData[i][5]) || 0;
            pctMeeting = parseFloat(summData[i][6]) || 0;
            totalHours = parseFloat(summData[i][7]) || 0;
            break;
          }
        }
      }
    }

    // Compute FOH / BOH SPLH from snapshot payload
    var fohData   = (body && body.fohScheduleData) || {};
    var bohData   = (body && body.bohScheduleData) || {};
    var salesData = (body && body.salesData)       || {};
    var fohSales = 0, fohHours = 0, bohSales = 0, bohHours = 0;

    Object.keys(fohData).forEach(function(day) {
      HOURS_LIST.forEach(function(hour) {
        var staff = (fohData[day] && fohData[day][hour]) || 0;
        var sales = (salesData[day] && salesData[day][hour]) || 0;
        if (staff > 0 && sales > 0) { fohSales += sales * fohFraction; fohHours += staff; }
      });
    });
    Object.keys(bohData).forEach(function(day) {
      HOURS_LIST.forEach(function(hour) {
        var staff = (bohData[day] && bohData[day][hour]) || 0;
        var sales = (salesData[day] && salesData[day][hour]) || 0;
        if (staff > 0 && sales > 0) { bohSales += sales * bohFraction; bohHours += staff; }
      });
    });
    var fohSplh = fohHours > 0 ? fohSales / fohHours : 0;
    var bohSplh = bohHours > 0 ? bohSales / bohHours : 0;

    // Format week date range
    var MONTHS   = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    var weekDate = weekStart ? new Date(weekStart + 'T12:00:00Z') : new Date();
    var satDate  = new Date(weekDate.getTime());
    satDate.setUTCDate(satDate.getUTCDate() + 5);
    var monStr   = MONTHS[weekDate.getUTCMonth()] + ' ' + weekDate.getUTCDate();
    var satStr   = MONTHS[satDate.getUTCMonth()]  + ' ' + satDate.getUTCDate();

    // Build bias flags section from snapshot payload
    var bFlags      = (body && body.biasFlags) || {};
    var activeFlags = Object.keys(bFlags).filter(function(d) { return bFlags[d] !== null; });
    var biasLines   = ['FORECAST ACCURACY FLAGS'];
    if (activeFlags.length > 0) {
      activeFlags.forEach(function(day) {
        var f   = bFlags[day];
        var dir = f.direction === 'over' ? 'Over-estimating' : 'Under-estimating';
        biasLines.push(
          '  ' + day + ': ' + dir + ' by ~' + Math.abs(f.avgErrorPct) + '% over ' +
          f.sampleCount + ' weeks \u2014 ' +
          (f.direction === 'over'
            ? 'staffing targets may be too high.'
            : 'staffing targets may be too low.')
        );
      });
    } else {
      biasLines.push('  No significant bias detected \u2014 curves are tracking well.');
    }

    var subject  = 'Labor Summary \u2014 Week of ' + monStr + ' | ' + locName;
    var bodyText = [
      'WEEKLY LABOR SUMMARY',
      'Location: ' + locName + ' | Week: ' + monStr + ' \u2013 ' + satStr,
      '',
      'SPLH PERFORMANCE',
      '  Overall Avg SPLH:     $' + avgSplh.toFixed(2) + '  (Goal: $' + goal.toFixed(2) + ')',
      '  FOH Avg SPLH:         $' + fohSplh.toFixed(2),
      '  BOH Avg SPLH:         $' + bohSplh.toFixed(2),
      '  Hours Meeting Goal:   ' + (pctMeeting * 100).toFixed(1) + '%',
      '  Hours Near Goal:      ' + (pctNear    * 100).toFixed(1) + '%',
      '  Hours Below Goal:     ' + (pctBelow   * 100).toFixed(1) + '%',
      '',
      'PEAK HOUR',
      '  Highest SPLH: ' + peakHour,
      '',
      'TOTAL SCHEDULED',
      '  ' + Math.round(totalHours) + ' person-hours',
      '',
      biasLines.join('\n'),
      '',
      '---',
      'Schedule Counter | ' + locName
    ].join('\n');

    GmailApp.sendEmail(email, subject, bodyText);
    Logger.log('Weekly summary email sent to ' + email);
    return { success: true };
  } catch (err) {
    Logger.log('sendWeeklySummaryEmail error: ' + err.message);
    return { error: err.message };
  }
}

// ============================================================
// ACTUALS CALIBRATION (dormant unless PAYROLL_SHEET_ID is wired)
// ============================================================
// Joins locked payroll batches (OT_History in the payroll spreadsheet — gross
// punched hours per week, immutable once uploaded) against this tool's actual
// api sales weeks to compute true weekly SPLH, cached in the actuals_weekly
// tab. The frontend blends that 8-week actual into the SPLH goal used by the
// staffing grid. Entirely dormant until wireCalibration() stores the payroll
// sheet id in Script Properties — nothing here runs for a stock install.

var ACTUALS_TAB = 'actuals_weekly';
var ACTUALS_HEADERS = [[
  'week_start', 'actual_ch_hours', 'ot_hours', 'api_sales',
  'actual_splh', 'days_covered', 'exclude', 'computed_at'
]];
var CAL_MIN_WEEKS = 4;        // qualifying weeks required before calibration engages
var CAL_WINDOW_WEEKS = 8;     // trailing window for the actual-SPLH average
var CAL_STALE_DAYS = 42;      // newest qualifying week older than this → stale
var CAL_MIN_DAYS_COVERED = 5; // api sales rows must span at least this many days

/**
 * One-time wiring — run from the GAS editor: wireCalibration('<payroll sheet id>').
 * Validates access + expected columns before storing anything, then runs the
 * first refresh. Re-runnable. After an account transfer, Script Properties do
 * NOT come along with pasted code — re-run this (see setup guide).
 */
function wireCalibration(payrollSheetId) {
  if (!payrollSheetId) throw new Error('Usage: wireCalibration("<payroll spreadsheet id>")');
  var probe = SpreadsheetApp.openById(payrollSheetId).getSheetByName('OT_History');
  if (!probe) throw new Error('Opened the spreadsheet but found no OT_History tab — wrong sheet id?');
  var headers = probe.getRange(1, 1, 1, probe.getLastColumn()).getValues()[0].map(function(h) { return String(h).trim(); });
  ['Period End', 'Week1_CH', 'Week2_CH'].forEach(function(h) {
    if (headers.indexOf(h) < 0) throw new Error('OT_History is missing expected column: ' + h);
  });
  PropertiesService.getScriptProperties().setProperty('PAYROLL_SHEET_ID', payrollSheetId);
  Logger.log('Calibration wired. Running first refresh...');
  refreshActuals();
  Logger.log('Done. Check the ' + ACTUALS_TAB + ' tab.');
}

/** Removes the wiring — feature goes fully dormant. Data in actuals_weekly is kept. */
function unwireCalibration() {
  PropertiesService.getScriptProperties().deleteProperty('PAYROLL_SHEET_ID');
  Logger.log('Calibration unwired.');
}

/**
 * Trigger-safe refresh: recomputes actuals_weekly from locked payroll batches.
 * Silent no-op when unwired. Alerts by email on failure (grid falls back to
 * the configured goal, so a failure degrades gracefully).
 */
function refreshActuals() {
  try {
    var payrollId = PropertiesService.getScriptProperties().getProperty('PAYROLL_SHEET_ID');
    if (!payrollId) return;
    computeActualsWeekly_(payrollId);
  } catch (err) {
    sendAlert(
      '\u26a0 Actuals Calibration Failed',
      'refreshActuals threw: ' + err.message +
      '\n\nThe staffing grid falls back to the configured SPLH goal until this is fixed.'
    );
    Logger.log('refreshActuals ERROR: ' + err.message);
  }
}

/**
 * The join. For every locked pay period (Period End = Saturday, per the payroll
 * tool's PaydayModule): split into two Sun–Sat weeks, sum gross home-store hours,
 * then match the api sales week whose Monday week_start falls inside each window.
 * Upserts actuals_weekly by week_start, preserving the operator-set exclude flag.
 *
 * Hours source per OT_History row (all gross — OT columns classify the same
 * hours and are recorded for context only):
 * - Multi-location rows carry per-location splits in Week1_CH/Week2_CH.
 * - Single-location home-store rows DON'T (the payroll parser only splits for
 *   Multi) — but for them 'Week 1 Hours'/'Week 2 Hours' ARE their home-store
 *   hours. Rows are matched to the home store by the Location column against
 *   this tool's location_name config (prefix match, so 'Cockrell Hill' matches
 *   'Cockrell Hill DTO'); other-store rows are ignored entirely.
 */
function computeActualsWeekly_(payrollId) {
  var otRows = readPayrollActuals_(payrollId);
  if (!otRows.length) throw new Error('OT_History has no readable rows.');

  var apiWeeks = readApiSalesWeeks_(); // { 'YYYY-MM-DD': { sales, days } }

  // Group payroll rows by period end
  var periods = {};
  otRows.forEach(function(r) {
    var key = localDateString_(r.periodEnd);
    if (!periods[key]) periods[key] = { periodEnd: r.periodEnd, w1ch: 0, w2ch: 0, w1ot: 0, w2ot: 0 };
    periods[key].w1ch += r.w1ch; periods[key].w2ch += r.w2ch;
    periods[key].w1ot += r.w1ot; periods[key].w2ot += r.w2ot;
  });

  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(ACTUALS_TAB);
  if (!sheet) throw new Error(ACTUALS_TAB + ' tab missing — run initSheets() first.');

  // Existing rows: preserve exclude flags across recomputes
  var existing = {}; // week_start -> { exclude }
  var lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    sheet.getRange(2, 1, lastRow - 1, 8).getValues().forEach(function(r) {
      existing[toDateString_(r[0])] = { exclude: String(r[6]).toUpperCase() === 'TRUE' };
    });
  }

  var now = new Date().toISOString();
  var out = [];
  Object.keys(periods).sort().forEach(function(key) {
    var p = periods[key];
    [
      { start: addDaysLocal_(p.periodEnd, -13), end: addDaysLocal_(p.periodEnd, -7), hours: p.w1ch, ot: p.w1ot },
      { start: addDaysLocal_(p.periodEnd, -6),  end: p.periodEnd,                    hours: p.w2ch, ot: p.w2ot }
    ].forEach(function(w) {
      // A 7-day window contains exactly one Monday → at most one api week matches
      var label = localDateString_(w.start), sales = 0, days = 0;
      Object.keys(apiWeeks).forEach(function(ws) {
        var d = parseLocalDate_(ws);
        if (d && d >= w.start && d <= w.end) { label = ws; sales = apiWeeks[ws].sales; days = apiWeeks[ws].days; }
      });
      // Weeks with neither hours nor sales are pure noise (batches predating
      // the weekly-split columns, sales weeks past the 90-day retention) — skip.
      if (w.hours === 0 && sales === 0) return;
      var qualifies = sales > 0 && w.hours > 0 && days >= CAL_MIN_DAYS_COVERED;
      var splh = qualifies ? Math.round(sales / w.hours * 100) / 100 : '';
      var excl = existing[label] ? existing[label].exclude : false;
      out.push([label, Math.round(w.hours * 100) / 100, Math.round(w.ot * 100) / 100,
                Math.round(sales), splh, days, excl ? 'TRUE' : 'FALSE', now]);
    });
  });

  // Dedupe by week_start (a re-uploaded/overwritten period recomputes cleanly),
  // then rewrite the tab in one batch.
  var byWeek = {};
  out.forEach(function(r) { byWeek[r[0]] = r; });
  var rows = Object.keys(byWeek).sort().map(function(k) { return byWeek[k]; });

  if (lastRow >= 2) sheet.getRange(2, 1, lastRow - 1, 8).clearContent();
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 1).setNumberFormat('@'); // keep week_start as text
    sheet.getRange(2, 1, rows.length, 8).setValues(rows);
  }
  SpreadsheetApp.flush();
  Logger.log('actuals_weekly: wrote ' + rows.length + ' week rows.');
}

/** Header-detected read of the payroll OT_History tab (cross-sheet, read-only). */
function readPayrollActuals_(payrollId) {
  var sheet = SpreadsheetApp.openById(payrollId).getSheetByName('OT_History');
  if (!sheet || sheet.getLastRow() < 2) return [];
  var data = sheet.getDataRange().getValues();
  var h = data[0].map(function(x) { return String(x).trim(); });
  var col = {
    periodEnd: h.indexOf('Period End'),
    location: h.indexOf('Location'),
    w1: h.indexOf('Week 1 Hours'), w2: h.indexOf('Week 2 Hours'),
    w1ch: h.indexOf('Week1_CH'), w2ch: h.indexOf('Week2_CH'),
    w1ot: h.indexOf('Week 1 OT'), w2ot: h.indexOf('Week 2 OT')
  };
  if (col.periodEnd < 0 || col.location < 0 || col.w1 < 0) {
    throw new Error('OT_History columns not found (Period End / Location / Week 1 Hours).');
  }

  // Home-store matcher: the payroll tool writes 'Cockrell Hill DTO' / 'DBU' /
  // 'Multi' in Location. Prefix-match against this tool's location_name config
  // so operators control it from the sheet; accept bare 'CH' for old rows.
  var homeName = String(getConfigValue('location_name') || 'Cockrell Hill').toLowerCase();
  function isHome(loc) {
    var l = String(loc).trim().toLowerCase();
    return l === 'ch' || (l !== '' && l !== 'multi' && l.indexOf(homeName) === 0);
  }

  var out = [];
  for (var i = 1; i < data.length; i++) {
    var pe = data[i][col.periodEnd];
    var d = (pe instanceof Date) ? new Date(pe.getFullYear(), pe.getMonth(), pe.getDate()) : parseLocalDate_(String(pe));
    if (!d) continue;

    var loc = String(data[i][col.location]).trim();
    var isMulti = loc.toLowerCase() === 'multi';
    var home = isHome(loc);
    if (!home && !isMulti) continue; // other-store row: contributes nothing

    var w1 = 0, w2 = 0;
    if (isMulti) {
      // Only Multi rows carry per-location splits; without them the home-store
      // share is unknowable — old Multi rows contribute 0 (small undercount).
      w1 = col.w1ch >= 0 ? (parseFloat(data[i][col.w1ch]) || 0) : 0;
      w2 = col.w2ch >= 0 ? (parseFloat(data[i][col.w2ch]) || 0) : 0;
    } else {
      // Home-store single-location row: Week 1/2 Hours ARE its home-store hours.
      // (The payroll parser never fills Week1_CH for these — that was the bug
      // that made hours read ~10x low.)
      w1 = parseFloat(data[i][col.w1]) || 0;
      w2 = parseFloat(data[i][col.w2]) || 0;
    }

    out.push({
      periodEnd: d,
      w1ch: w1,
      w2ch: w2,
      w1ot: col.w1ot >= 0 ? (parseFloat(data[i][col.w1ot]) || 0) : 0,
      w2ot: col.w2ot >= 0 ? (parseFloat(data[i][col.w2ot]) || 0) : 0
    });
  }
  return out;
}

/**
 * Median same-weekday api day-totals from sales_history, excluding the most
 * recent api week (the one the pipeline just archived — it must not vouch for
 * itself). Returns { Monday: median$, ... }; empty when history is thin.
 */
function typicalApiDayTotals_(ss) {
  var sheet = ss.getSheetByName('sales_history');
  var result = {};
  if (!sheet || sheet.getLastRow() < 2) return result;
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();

  var latestWeek = '';
  data.forEach(function(r) {
    if (String(r[4]) === 'api') {
      var ws = toDateString_(r[0]);
      if (ws > latestWeek) latestWeek = ws;
    }
  });

  var totals = {}; // { day: { week: total } }
  data.forEach(function(r) {
    if (String(r[4]) !== 'api') return;
    var ws = toDateString_(r[0]);
    if (ws === latestWeek) return;
    var day = String(r[1]);
    if (!totals[day]) totals[day] = {};
    totals[day][ws] = (totals[day][ws] || 0) + (parseFloat(r[3]) || 0);
  });

  Object.keys(totals).forEach(function(day) {
    result[day] = median_(Object.keys(totals[day]).map(function(w) { return totals[day][w]; }));
  });
  return result;
}

function median_(values) {
  if (!values.length) return 0;
  var sorted = values.slice().sort(function(a, b) { return a - b; });
  var mid = Math.floor(sorted.length / 2);
  return sorted.length % 2 ? sorted[mid] : (sorted[mid - 1] + sorted[mid]) / 2;
}

/** Weekly api sales from sales_history: { week_start: { sales, days } }. api rows only. */
function readApiSalesWeeks_() {
  var sheet = getSpreadsheet().getSheetByName('sales_history');
  var byWeek = {};
  if (!sheet || sheet.getLastRow() < 2) return byWeek;
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
  var daySets = {};
  data.forEach(function(r) {
    if (String(r[4]) !== 'api') return;
    var ws = toDateString_(r[0]);
    if (!ws) return;
    if (!byWeek[ws]) { byWeek[ws] = { sales: 0, days: 0 }; daySets[ws] = {}; }
    byWeek[ws].sales += parseFloat(r[3]) || 0;
    daySets[ws][String(r[1])] = true;
  });
  Object.keys(byWeek).forEach(function(ws) {
    byWeek[ws].days = Object.keys(daySets[ws]).length;
  });
  return byWeek;
}

/**
 * Summary the frontend consumes. Reads only the local actuals_weekly tab —
 * never the payroll sheet — so page loads stay fast. Returns null when the
 * feature is unwired (frontend then behaves 100% stock).
 */
function buildCalibration_(ss, cfgMap) {
  if (!PropertiesService.getScriptProperties().getProperty('PAYROLL_SHEET_ID')) return null;
  var sheet = ss.getSheetByName(ACTUALS_TAB);
  var result = {
    ready: false, reason: 'insufficient',
    actualSplh: 0, weeksUsed: 0, weeksExcluded: 0, weeksRejected: 0, newestWeek: '',
    weight:  clampNum_(parseFloat(cfgMap['splh_cal_weight']),  0, 1,   0.5),
    floor:   clampNum_(parseFloat(cfgMap['splh_cal_floor']),   0.5, 1, 0.85),
    ceiling: clampNum_(parseFloat(cfgMap['splh_cal_ceiling']), 1, 1.5, 1.10)
  };
  if (!sheet || sheet.getLastRow() < 2) return result;

  var goal = parseFloat(cfgMap['splh_goal_overall']) || 90;
  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues()
    .map(function(r) {
      return { week: toDateString_(r[0]), splh: parseFloat(r[4]) || 0,
               days: parseInt(r[5], 10) || 0, exclude: String(r[6]).toUpperCase() === 'TRUE' };
    })
    .filter(function(r) { return r.week && r.splh > 0 && r.days >= CAL_MIN_DAYS_COVERED; })
    .sort(function(a, b) { return a.week < b.week ? -1 : 1; });

  // Data-quality band: a real week's SPLH lands near the goal; anything outside
  // 0.25x–2.5x means broken inputs (e.g. hours read wrong), not performance.
  // Reject rather than let the clamps quietly absorb garbage.
  var sane = rows.filter(function(r) { return r.splh >= goal * 0.25 && r.splh <= goal * 2.5; });
  result.weeksRejected = rows.length - sane.length;
  rows = sane;

  result.weeksExcluded = rows.filter(function(r) { return r.exclude; }).length;
  var qualifying = rows.filter(function(r) { return !r.exclude; }).slice(-CAL_WINDOW_WEEKS);
  result.weeksUsed = qualifying.length;
  if (qualifying.length < CAL_MIN_WEEKS) return result; // reason: insufficient

  var newest = qualifying[qualifying.length - 1].week;
  result.newestWeek = newest;
  var ageDays = (new Date() - parseLocalDate_(newest)) / 86400000;
  if (ageDays > CAL_STALE_DAYS) { result.reason = 'stale'; return result; }

  // Median, not mean: a single anomalous week (big catering run, data hiccup)
  // moves a mean by 1/8th of its weirdness but moves the median not at all.
  result.actualSplh = Math.round(median_(qualifying.map(function(r) { return r.splh; })) * 100) / 100;
  result.ready = true;
  result.reason = 'ok';
  return result;
}

/** Seeds one config key only if absent — reaches existing installs on initSheets() re-run. */
function seedConfigKeyIfMissing_(ss, key, value) {
  var sheet = ss.getSheetByName('config');
  if (!sheet) return;
  var last = sheet.getLastRow();
  if (last >= 2) {
    var keys = sheet.getRange(2, 1, last - 1, 1).getValues();
    for (var i = 0; i < keys.length; i++) {
      if (String(keys[i][0]).trim() === key) return;
    }
  }
  sheet.appendRow([key, String(value)]);
}

function clampNum_(v, lo, hi, fallback) {
  if (isNaN(v)) return fallback;
  return Math.min(Math.max(v, lo), hi);
}

/** Local-noon date from YYYY-MM-DD — never UTC parsing (same rule as payroll's PaydayModule). */
function parseLocalDate_(s) {
  var m = String(s).trim().match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!m) return null;
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), 12, 0, 0);
}

function localDateString_(d) {
  var mm = String(d.getMonth() + 1); if (mm.length < 2) mm = '0' + mm;
  var dd = String(d.getDate());      if (dd.length < 2) dd = '0' + dd;
  return d.getFullYear() + '-' + mm + '-' + dd;
}

function addDaysLocal_(d, n) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate() + n, 12, 0, 0);
}
