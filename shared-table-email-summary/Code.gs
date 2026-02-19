/**
 * ============================================================
 * Weekly Waste Summary — Google Apps Script
 * ============================================================
 *
 * SETUP INSTRUCTIONS:
 *
 * 1. Open your Google Sheet that receives form responses.
 *    The response sheet should be named "Form Responses 2".
 *
 * 2. Create a new sheet tab named "Config" with the following layout:
 *
 *    Row 1:  A = "Send Day"         B = "Monday"              (day of week to send email)
 *    Row 2:  A = "Recipient Email"  B = "you@email.com"       (primary recipient)
 *    Row 3:  A = "Week Start Day"   B = "Sunday"              (Sunday or Monday)
 *    Row 4:  A = "Last Sent"        B = (leave blank)         (script auto-fills this)
 *    Row 5:  A = "CC Email"         B = (optional CC address)
 *    Row 6:  A = "Send Time"        B = "7"                   (hour in 24h format, 0-23)
 *
 * 3. In the Apps Script editor (Extensions > Apps Script), paste this file.
 *
 * 4. Run the function createDailyTrigger() ONE TIME from the editor
 *    — OR use the menu:  Waste Summary > Reinstall Trigger
 *    You will be prompted to authorize the script.
 *
 * 5. That's it. A "Waste Summary" menu appears in the toolbar.
 *    Use  Waste Summary > Settings  to open the sidebar and
 *    adjust all options without touching the Config sheet.
 *
 * To stop the automation: Waste Summary > Remove Trigger
 * ============================================================
 */

// ── Sheet names ──────────────────────────────────────────────
var RESPONSE_SHEET_NAME = 'Form Responses 2';
var CONFIG_SHEET_NAME   = 'Config';

// ── Category groupings for email layout ──────────────────────
var CATEGORIES = [
  {
    title: 'PROTEINS (pounds)',
    items: [
      'Bacon, Full Strip',
      'Breakfast Filet Spicy',
      'Chicken, Breakfast Filets',
      'Chicken, Filet Spicy',
      'Chicken, Filets',
      'Chicken, Filets, Grilled',
      'Chicken, Nuggets',
      'Chicken, Nuggets, Grilled',
      'Chicken, Tenders',
      'Sausage, Precooked 4/10# Boxes',
      'Egg, Scrambled Blend'
    ]
  },
  {
    title: 'BREAD & BREAKFAST (units/pounds)',
    items: [
      'Biscuit w/ Jelly',
      'Muffin, English Multigrain',
      'Roll, Mini Yeast',
      'Tortilla, 10 Flour',
      'Potato, Hashrounds',
      'Potato, Waffle Fries'
    ]
  },
  {
    title: 'SIDES (units/pounds)',
    items: [
      'CFA Side Salad',
      'Kale Crunch Side, Large',
      'Kale Crunch Side, Small',
      'Mac & Cheese, Large',
      'Fruit Cup, Large',
      'Fruit Cup, Medium',
      'Fruit Cup, Small',
      'Greek Yogurt Parfait, Granola'
    ]
  },
  {
    title: 'SALADS & WRAPS (units)',
    items: [
      'Salad, Cobb Base',
      'Salad, Mkt Base',
      'Salad, Spicy SW Base',
      'Wrap, Grilled Chicken Cool'
    ]
  },
  {
    title: 'BEVERAGES (units/gallons)',
    items: [
      'Iced Tea (Gallon)',
      'Lemonade (Gallon)'
    ]
  },
  {
    title: 'SOUPS & SWEETS (pounds/units)',
    items: [
      'Soup, Base Chicken Noodle',
      'Soup, Base Chkn Tortilla',
      'Chocolate Fudge Brownie',
      'Cookie, Choc Chunk No Thaw'
    ]
  }
];

// ═══════════════════════════════════════════════════════════════
//  CUSTOM MENU & SIDEBAR
// ═══════════════════════════════════════════════════════════════

/**
 * Adds the "Waste Summary" menu to the Google Sheets toolbar
 * every time the spreadsheet is opened.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Waste Summary')
    .addItem('Settings', 'openSettingsSidebar')
    .addSeparator()
    .addItem('Send Test Email Now', 'sendTestEmail')
    .addSeparator()
    .addItem('Reinstall Trigger', 'createDailyTriggerWithAlert')
    .addItem('Remove Trigger', 'deleteTriggerWithAlert')
    .addToUi();
}

/**
 * Opens the settings sidebar panel.
 */
function openSettingsSidebar() {
  var html = HtmlService.createHtmlOutput(getSidebarHtml())
    .setTitle('Waste Summary Settings')
    .setWidth(320);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Returns all current settings to the sidebar.
 * Called from client-side JS via google.script.run.
 */
function getSettings() {
  var settings = {};
  try { settings.sendDay        = String(getConfigValue('Send Day')).trim();        } catch(_) { settings.sendDay = 'Monday'; }
  try { settings.recipientEmail = String(getConfigValue('Recipient Email')).trim();  } catch(_) { settings.recipientEmail = ''; }
  try { settings.weekStartDay   = String(getConfigValue('Week Start Day')).trim();   } catch(_) { settings.weekStartDay = 'Sunday'; }
  try { settings.ccEmail        = String(getConfigValue('CC Email')).trim();         } catch(_) { settings.ccEmail = ''; }
  try { settings.sendTime       = String(getConfigValue('Send Time')).trim();        } catch(_) { settings.sendTime = '7'; }
  try { settings.lastSent       = String(getConfigValue('Last Sent')).trim();        } catch(_) { settings.lastSent = ''; }
  return settings;
}

/**
 * Saves settings from the sidebar back to the Config sheet,
 * and reinstalls the trigger if the send time changed.
 * Called from client-side JS via google.script.run.
 */
function saveSettings(settings) {
  // Read old send time before saving
  var oldSendTime = '7';
  try { oldSendTime = String(getConfigValue('Send Time')).trim(); } catch(_) {}

  ensureConfigRow('Send Day',        settings.sendDay);
  ensureConfigRow('Recipient Email', settings.recipientEmail);
  ensureConfigRow('Week Start Day',  settings.weekStartDay);
  ensureConfigRow('CC Email',        settings.ccEmail);
  ensureConfigRow('Send Time',       settings.sendTime);

  // If send time changed, reinstall the trigger at the new hour
  if (settings.sendTime !== oldSendTime) {
    createDailyTrigger();
    return 'Settings saved. Trigger updated to ' + formatHour(parseInt(settings.sendTime, 10)) + '.';
  }

  return 'Settings saved.';
}

/**
 * Ensures a config row exists. If the label is missing, appends it.
 */
function ensureConfigRow(label, value) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  if (!sheet) {
    throw new Error('Config sheet not found.');
  }

  var data = sheet.getDataRange().getValues();
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === label.toLowerCase()) {
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }

  // Label not found — append a new row
  sheet.appendRow([label, value]);
}

/**
 * Formats an hour integer (0-23) into a readable string like "7:00 AM".
 */
function formatHour(h) {
  if (h === 0)  return '12:00 AM';
  if (h === 12) return '12:00 PM';
  if (h < 12)   return h + ':00 AM';
  return (h - 12) + ':00 PM';
}

// ═══════════════════════════════════════════════════════════════
//  SIDEBAR HTML
// ═══════════════════════════════════════════════════════════════

function getSidebarHtml() {
  // Build hour options for the dropdown
  var hourOptions = '';
  for (var h = 0; h < 24; h++) {
    var label;
    if (h === 0) label = '12:00 AM';
    else if (h === 12) label = '12:00 PM';
    else if (h < 12) label = h + ':00 AM';
    else label = (h - 12) + ':00 PM';
    hourOptions += '<option value="' + h + '">' + label + '</option>';
  }

  return '<!DOCTYPE html>'
  + '<html><head>'
  + '<style>'
  + '  * { box-sizing: border-box; margin: 0; padding: 0; }'
  + '  body { font-family: "Google Sans", Roboto, Arial, sans-serif; font-size: 13px; color: #333; padding: 16px; background: #fff; }'
  + '  h2 { font-size: 16px; font-weight: 600; margin-bottom: 4px; color: #1a1a1a; }'
  + '  .subtitle { font-size: 11px; color: #888; margin-bottom: 18px; }'
  + '  .field { margin-bottom: 14px; }'
  + '  label { display: block; font-size: 11px; font-weight: 600; text-transform: uppercase; letter-spacing: 0.5px; color: #555; margin-bottom: 4px; }'
  + '  select, input[type="email"], input[type="text"] {'
  + '    width: 100%; padding: 8px 10px; font-size: 13px; border: 1px solid #dadce0;'
  + '    border-radius: 6px; background: #fff; color: #333; outline: none; transition: border 0.2s;'
  + '  }'
  + '  select:focus, input:focus { border-color: #1a73e8; }'
  + '  .hint { font-size: 11px; color: #999; margin-top: 3px; }'
  + '  .section-label { font-size: 11px; font-weight: 700; color: #1a73e8; text-transform: uppercase; letter-spacing: 0.8px; margin: 18px 0 10px; }'
  + '  .btn-row { display: flex; gap: 8px; margin-top: 20px; }'
  + '  .btn { flex: 1; padding: 10px 0; font-size: 13px; font-weight: 600; border: none; border-radius: 6px; cursor: pointer; transition: background 0.15s; text-align: center; }'
  + '  .btn-primary { background: #1a73e8; color: #fff; }'
  + '  .btn-primary:hover { background: #1557b0; }'
  + '  .btn-secondary { background: #f1f3f4; color: #333; }'
  + '  .btn-secondary:hover { background: #e2e4e6; }'
  + '  .btn:disabled { opacity: 0.5; cursor: not-allowed; }'
  + '  .status { margin-top: 12px; padding: 10px 12px; border-radius: 6px; font-size: 12px; display: none; }'
  + '  .status.success { display: block; background: #e6f4ea; color: #137333; }'
  + '  .status.error { display: block; background: #fce8e6; color: #c5221f; }'
  + '  .divider { border: none; border-top: 1px solid #eee; margin: 6px 0 2px; }'
  + '  .meta { font-size: 11px; color: #aaa; margin-top: 20px; padding-top: 12px; border-top: 1px solid #eee; }'
  + '  .loading-overlay { position: fixed; inset: 0; background: rgba(255,255,255,0.9); display: flex; align-items: center; justify-content: center; z-index: 10; }'
  + '  .spinner { width: 28px; height: 28px; border: 3px solid #dadce0; border-top-color: #1a73e8; border-radius: 50%; animation: spin 0.7s linear infinite; }'
  + '  @keyframes spin { to { transform: rotate(360deg); } }'
  + '</style>'
  + '</head><body>'

  // Loading overlay
  + '<div class="loading-overlay" id="loader"><div class="spinner"></div></div>'

  + '<h2>Waste Summary</h2>'
  + '<div class="subtitle">Configure your weekly email report</div>'

  // ── SCHEDULE SECTION ──
  + '<div class="section-label">Schedule</div>'

  + '<div class="field">'
  + '  <label>Send Day</label>'
  + '  <select id="sendDay">'
  + '    <option>Sunday</option><option>Monday</option><option>Tuesday</option>'
  + '    <option>Wednesday</option><option>Thursday</option><option>Friday</option><option>Saturday</option>'
  + '  </select>'
  + '</div>'

  + '<div class="field">'
  + '  <label>Send Time</label>'
  + '  <select id="sendTime">' + hourOptions + '</select>'
  + '  <div class="hint">Trigger fires within ~15 min of this hour</div>'
  + '</div>'

  + '<div class="field">'
  + '  <label>Week Starts On</label>'
  + '  <select id="weekStartDay">'
  + '    <option>Sunday</option><option>Monday</option>'
  + '  </select>'
  + '  <div class="hint">Defines the "last week" date range</div>'
  + '</div>'

  + '<hr class="divider">'

  // ── RECIPIENTS SECTION ──
  + '<div class="section-label">Recipients</div>'

  + '<div class="field">'
  + '  <label>Recipient Email</label>'
  + '  <input type="email" id="recipientEmail" placeholder="you@email.com">'
  + '</div>'

  + '<div class="field">'
  + '  <label>CC Email <span style="font-weight:400;text-transform:none;">(optional)</span></label>'
  + '  <input type="email" id="ccEmail" placeholder="optional@email.com">'
  + '</div>'

  // ── Buttons ──
  + '<div class="btn-row">'
  + '  <button class="btn btn-primary" id="saveBtn" onclick="save()">Save Settings</button>'
  + '  <button class="btn btn-secondary" onclick="sendTest()">Send Test</button>'
  + '</div>'

  + '<div class="status" id="status"></div>'

  // ── Meta info ──
  + '<div class="meta" id="metaInfo"></div>'

  // ── Script ──
  + '<script>'

  + 'function init() {'
  + '  google.script.run.withSuccessHandler(function(s) {'
  + '    document.getElementById("sendDay").value = s.sendDay || "Monday";'
  + '    document.getElementById("sendTime").value = s.sendTime || "7";'
  + '    document.getElementById("weekStartDay").value = s.weekStartDay || "Sunday";'
  + '    document.getElementById("recipientEmail").value = s.recipientEmail || "";'
  + '    document.getElementById("ccEmail").value = s.ccEmail || "";'
  + '    if (s.lastSent && s.lastSent !== "undefined" && s.lastSent !== "") {'
  + '      document.getElementById("metaInfo").innerHTML = "Last email sent: <strong>" + s.lastSent + "</strong>";'
  + '    }'
  + '    document.getElementById("loader").style.display = "none";'
  + '  }).withFailureHandler(function(e) {'
  + '    showStatus("Failed to load settings: " + e.message, "error");'
  + '    document.getElementById("loader").style.display = "none";'
  + '  }).getSettings();'
  + '}'

  + 'function save() {'
  + '  var btn = document.getElementById("saveBtn");'
  + '  btn.disabled = true; btn.textContent = "Saving...";'
  + '  hideStatus();'
  + '  var s = {'
  + '    sendDay: document.getElementById("sendDay").value,'
  + '    sendTime: document.getElementById("sendTime").value,'
  + '    weekStartDay: document.getElementById("weekStartDay").value,'
  + '    recipientEmail: document.getElementById("recipientEmail").value.trim(),'
  + '    ccEmail: document.getElementById("ccEmail").value.trim()'
  + '  };'
  + '  if (!s.recipientEmail) {'
  + '    showStatus("Recipient email is required.", "error");'
  + '    btn.disabled = false; btn.textContent = "Save Settings";'
  + '    return;'
  + '  }'
  + '  google.script.run.withSuccessHandler(function(msg) {'
  + '    showStatus(msg, "success");'
  + '    btn.disabled = false; btn.textContent = "Save Settings";'
  + '  }).withFailureHandler(function(e) {'
  + '    showStatus("Error: " + e.message, "error");'
  + '    btn.disabled = false; btn.textContent = "Save Settings";'
  + '  }).saveSettings(s);'
  + '}'

  + 'function sendTest() {'
  + '  hideStatus();'
  + '  showStatus("Sending test email...", "success");'
  + '  google.script.run.withSuccessHandler(function(msg) {'
  + '    showStatus(msg, "success");'
  + '  }).withFailureHandler(function(e) {'
  + '    showStatus("Error: " + e.message, "error");'
  + '  }).sendTestEmail();'
  + '}'

  + 'function showStatus(msg, type) {'
  + '  var el = document.getElementById("status");'
  + '  el.className = "status " + type;'
  + '  el.textContent = msg;'
  + '}'

  + 'function hideStatus() {'
  + '  var el = document.getElementById("status");'
  + '  el.className = "status";'
  + '  el.textContent = "";'
  + '}'

  + 'init();'
  + '</script>'
  + '</body></html>';
}

// ═══════════════════════════════════════════════════════════════
//  MAIN ENTRY POINT
// ═══════════════════════════════════════════════════════════════

/**
 * Runs daily via time-based trigger.
 * Only sends the email on the configured Send Day.
 */
function onDailyTrigger() {
  try {
    // 1. Check if today is the send day
    var sendDay = getConfigValue('Send Day');
    var sendDayIndex = getDayIndex(sendDay);
    var today = new Date();
    var todayIndex = today.getDay();

    Logger.log('Today: ' + today.toDateString() + ' (index ' + todayIndex + ')');
    Logger.log('Send Day: ' + sendDay + ' (index ' + sendDayIndex + ')');

    if (todayIndex !== sendDayIndex) {
      Logger.log('Not the send day — exiting.');
      return;
    }

    // 2. Check Last Sent to prevent duplicate sends
    var lastSent = getConfigValue('Last Sent');
    var todayStr = Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy-MM-dd');

    if (lastSent && String(lastSent).trim() === todayStr) {
      Logger.log('Already sent today (' + todayStr + ') — exiting.');
      return;
    }

    // 3. Calculate date range
    var range = getLastWeekRange();
    Logger.log('Date range: ' + range.startDate + ' to ' + range.endDate);

    // 4. Filter rows
    var rows = getFilteredRows(range);
    Logger.log('Filtered rows: ' + rows.length);

    // 5. Calculate sums
    var sums = calculateSums(rows);
    Logger.log('Row count in sums: ' + sums._rowCount);

    // 6. Send email
    sendSummaryEmail(sums, range);
    Logger.log('Email sent successfully.');

    // 7. Update Last Sent
    setConfigValue('Last Sent', todayStr);
    Logger.log('Last Sent updated to ' + todayStr);

  } catch (e) {
    Logger.log('ERROR in onDailyTrigger: ' + e.message + '\n' + e.stack);
    try {
      var recipient = getConfigValue('Recipient Email');
      MailApp.sendEmail({
        to: recipient,
        subject: 'Weekly Waste Summary — SCRIPT ERROR',
        htmlBody: '<div style="font-family:Arial,sans-serif;padding:20px;">'
          + '<h2 style="color:#cc0000;">⚠ Script Error</h2>'
          + '<p>The Weekly Waste Summary script encountered an error:</p>'
          + '<pre style="background:#f5f5f5;padding:12px;border-radius:4px;overflow-x:auto;">'
          + e.message + '\n\n' + e.stack
          + '</pre>'
          + '<p style="color:#666;font-size:12px;">Check the Apps Script execution log for more details.</p>'
          + '</div>'
      });
    } catch (mailErr) {
      Logger.log('Failed to send error email: ' + mailErr.message);
    }
  }
}

// ═══════════════════════════════════════════════════════════════
//  SEND TEST EMAIL (manual trigger from sidebar / menu)
// ═══════════════════════════════════════════════════════════════

/**
 * Immediately runs the full pipeline for last week's data
 * and sends the email, bypassing the day-of-week check.
 * Returns a status message to the sidebar.
 */
function sendTestEmail() {
  try {
    var range = getLastWeekRange();
    var rows = getFilteredRows(range);
    var sums = calculateSums(rows);
    sendSummaryEmail(sums, range);
    var dateLabel = formatDateRange(range.startDate, range.endDate);
    return 'Test email sent for ' + dateLabel + ' (' + sums._rowCount + ' submissions).';
  } catch (e) {
    throw new Error('Test email failed: ' + e.message);
  }
}

// ═══════════════════════════════════════════════════════════════
//  DATE RANGE
// ═══════════════════════════════════════════════════════════════

/**
 * Returns { startDate, endDate } for the previous full week,
 * based on the configured Week Start Day.
 */
function getLastWeekRange() {
  var weekStartDay = getConfigValue('Week Start Day');
  var weekStartIndex = getDayIndex(weekStartDay);

  var today = new Date();
  var todayIndex = today.getDay();

  var daysSinceWeekStart = (todayIndex - weekStartIndex + 7) % 7;
  if (daysSinceWeekStart === 0) {
    daysSinceWeekStart = 7;
  }

  var currentWeekStart = new Date(today);
  currentWeekStart.setDate(today.getDate() - daysSinceWeekStart);

  var startDate = new Date(currentWeekStart);
  startDate.setDate(currentWeekStart.getDate() - 7);
  startDate.setHours(0, 0, 0, 0);

  var endDate = new Date(currentWeekStart);
  endDate.setDate(currentWeekStart.getDate() - 1);
  endDate.setHours(23, 59, 59, 999);

  return { startDate: startDate, endDate: endDate };
}

// ═══════════════════════════════════════════════════════════════
//  ROW FILTERING
// ═══════════════════════════════════════════════════════════════

/**
 * Returns rows whose Timestamp falls within [startDate, endDate].
 */
function getFilteredRows(range) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(RESPONSE_SHEET_NAME);
    if (!sheet) {
      throw new Error('Sheet "' + RESPONSE_SHEET_NAME + '" not found.');
    }

    var data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      Logger.log('No data rows found (only header or empty).');
      return [];
    }

    var filtered = [];
    for (var i = 1; i < data.length; i++) {
      var raw = data[i][0];
      var ts;

      if (raw instanceof Date) {
        ts = raw;
      } else {
        ts = new Date(raw);
      }

      if (isNaN(ts.getTime())) {
        Logger.log('Warning: Could not parse timestamp on row ' + (i + 1) + ': ' + raw);
        continue;
      }

      if (ts >= range.startDate && ts <= range.endDate) {
        filtered.push(data[i]);
      }
    }

    return filtered;
  } catch (e) {
    Logger.log('Error in getFilteredRows: ' + e.message);
    throw e;
  }
}

// ═══════════════════════════════════════════════════════════════
//  AGGREGATION
// ═══════════════════════════════════════════════════════════════

/**
 * Sums every numeric column (B onward) across all filtered rows.
 * Returns { headerName: total, ..., _rowCount: n }
 */
function calculateSums(rows) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(RESPONSE_SHEET_NAME);
    var headers = sheet.getDataRange().getValues()[0];

    var sums = { _rowCount: rows.length };

    for (var col = 1; col < headers.length; col++) {
      var headerName = String(headers[col]).trim();
      if (!headerName) continue;

      var total = 0;
      for (var r = 0; r < rows.length; r++) {
        total += sanitizeNumeric(rows[r][col]);
      }

      sums[headerName] = Math.round(total * 100) / 100;
    }

    return sums;
  } catch (e) {
    Logger.log('Error in calculateSums: ' + e.message);
    throw e;
  }
}

// ═══════════════════════════════════════════════════════════════
//  EMAIL
// ═══════════════════════════════════════════════════════════════

/**
 * Composes and sends the HTML summary email (or no-data email).
 */
function sendSummaryEmail(sums, range) {
  try {
    var recipient = getConfigValue('Recipient Email');
    var ccRaw = '';
    try { ccRaw = getConfigValue('CC Email'); } catch (_) {}
    var cc = (ccRaw && String(ccRaw).trim()) ? String(ccRaw).trim() : '';

    var dateLabel = formatDateRange(range.startDate, range.endDate);
    var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

    // ── No-data email ──
    if (sums._rowCount === 0) {
      var noDataSubject = 'Weekly Waste Summary — No Submissions Found (' + dateLabel + ')';
      var noDataBody = '<div style="font-family:Arial,Helvetica,sans-serif;max-width:600px;margin:0 auto;padding:20px;">'
        + '<div style="background:#1a1a1a;color:#ffffff;padding:16px 20px;border-radius:6px 6px 0 0;">'
        + '<h2 style="margin:0;font-size:18px;">Weekly Waste Summary</h2>'
        + '<p style="margin:4px 0 0;font-size:13px;color:#cccccc;">' + dateLabel + '</p>'
        + '</div>'
        + '<div style="padding:24px 20px;border:1px solid #e0e0e0;border-top:none;border-radius:0 0 6px 6px;">'
        + '<p style="font-size:15px;color:#333333;">No form submissions were found for the week of <strong>' + dateLabel + '</strong>.</p>'
        + '<p style="font-size:14px;color:#666666;">If submissions were expected, verify that the Google Form is still linked to the <em>' + RESPONSE_SHEET_NAME + '</em> sheet and that dates are being recorded correctly.</p>'
        + '</div>'
        + '<p style="font-size:11px;color:#999999;text-align:center;margin-top:12px;">Generated by Weekly Summary Script | ' + now + '</p>'
        + '</div>';

      var opts = { to: recipient, subject: noDataSubject, htmlBody: noDataBody };
      if (cc) opts.cc = cc;
      MailApp.sendEmail(opts);
      return;
    }

    // ── Full summary email ──
    var subject = 'Weekly Waste Summary — ' + dateLabel;

    var html = '';
    html += '<div style="font-family:Arial,Helvetica,sans-serif;max-width:640px;margin:0 auto;">';
    html += '<div style="background:#1a1a1a;color:#ffffff;padding:16px 20px;border-radius:6px 6px 0 0;">';
    html += '<h2 style="margin:0;font-size:18px;">Weekly Waste Summary</h2>';
    html += '<p style="margin:4px 0 0;font-size:13px;color:#cccccc;">' + dateLabel + '</p>';
    html += '</div>';

    html += '<div style="padding:20px;border:1px solid #e0e0e0;border-top:none;border-radius:0 0 6px 6px;background:#ffffff;">';
    html += '<p style="font-size:15px;color:#333333;margin-top:0;">Submissions this week: <strong>' + sums._rowCount + '</strong></p>';

    // Build a lookup of all items that belong to a known category
    var categorizedItems = {};
    for (var c = 0; c < CATEGORIES.length; c++) {
      var cat = CATEGORIES[c];
      html += buildCategorySection(cat.title, cat.items, sums);
      for (var ci = 0; ci < cat.items.length; ci++) {
        categorizedItems[cat.items[ci]] = true;
      }
    }

    // Find any items in sums that are NOT in a known category (new/rearranged columns)
    var uncategorized = [];
    for (var key in sums) {
      if (key === '_rowCount') continue;
      if (!categorizedItems[key]) {
        uncategorized.push(key);
      }
    }

    // If there are uncategorized items, add a catch-all section
    if (uncategorized.length > 0) {
      uncategorized.sort();
      html += buildCategorySection('OTHER / NEW ITEMS', uncategorized, sums);
    }

    html += '</div>';
    html += '<p style="font-size:11px;color:#999999;text-align:center;margin-top:12px;">Generated by Weekly Summary Script | ' + now + '</p>';
    html += '</div>';

    var opts = { to: recipient, subject: subject, htmlBody: html };
    if (cc) opts.cc = cc;
    MailApp.sendEmail(opts);

  } catch (e) {
    Logger.log('Error in sendSummaryEmail: ' + e.message);
    throw e;
  }
}

/**
 * Builds the HTML for one category section (header + table).
 */
function buildCategorySection(title, items, sums) {
  var html = '';

  html += '<div style="margin-top:20px;margin-bottom:6px;padding:8px 12px;background:#f0f0f0;border-left:4px solid #1a1a1a;font-weight:bold;font-size:14px;color:#333333;">'
    + title + '</div>';

  html += '<table style="width:100%;border-collapse:collapse;font-size:13px;margin-bottom:4px;" cellpadding="0" cellspacing="0">';
  html += '<tr style="background:#1a1a1a;color:#ffffff;">';
  html += '<th style="text-align:left;padding:8px 12px;font-weight:bold;">Item</th>';
  html += '<th style="text-align:right;padding:8px 12px;font-weight:bold;">Total</th>';
  html += '</tr>';

  for (var i = 0; i < items.length; i++) {
    var itemName = items[i];
    var value = (sums[itemName] !== undefined) ? sums[itemName] : 0;
    var rowBg = (i % 2 === 0) ? '#ffffff' : '#f9f9f9';
    var valueColor = (value === 0) ? '#999999' : '#333333';

    html += '<tr style="background:' + rowBg + ';">';
    html += '<td style="padding:6px 12px;border-bottom:1px solid #eeeeee;color:#333333;">' + itemName + '</td>';
    html += '<td style="padding:6px 12px;border-bottom:1px solid #eeeeee;text-align:right;color:' + valueColor + ';">' + value + '</td>';
    html += '</tr>';
  }

  html += '</table>';
  return html;
}

// ═══════════════════════════════════════════════════════════════
//  TRIGGER MANAGEMENT
// ═══════════════════════════════════════════════════════════════

/**
 * Installs the daily trigger at the configured Send Time hour.
 * Safe to re-run — removes old triggers first.
 */
function createDailyTrigger() {
  deleteTrigger();

  var sendHour = 7;
  try {
    sendHour = parseInt(getConfigValue('Send Time'), 10);
    if (isNaN(sendHour) || sendHour < 0 || sendHour > 23) sendHour = 7;
  } catch(_) {}

  ScriptApp.newTrigger('onDailyTrigger')
    .timeBased()
    .everyDays(1)
    .atHour(sendHour)
    .create();

  Logger.log('Daily trigger created — onDailyTrigger will run every day at ~' + formatHour(sendHour) + '.');
}

/**
 * Wrapper for menu item — shows a confirmation alert.
 */
function createDailyTriggerWithAlert() {
  createDailyTrigger();
  var sendHour = 7;
  try { sendHour = parseInt(getConfigValue('Send Time'), 10); } catch(_) {}
  SpreadsheetApp.getUi().alert('Trigger installed. The script will run daily at ~' + formatHour(sendHour) + '.');
}

/**
 * Removes all triggers pointing to onDailyTrigger.
 */
function deleteTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'onDailyTrigger') {
      ScriptApp.deleteTrigger(triggers[i]);
      Logger.log('Deleted existing trigger: ' + triggers[i].getUniqueId());
    }
  }
}

/**
 * Wrapper for menu item — shows a confirmation alert.
 */
function deleteTriggerWithAlert() {
  deleteTrigger();
  SpreadsheetApp.getUi().alert('All triggers removed. No more automatic emails will be sent.');
}

// ═══════════════════════════════════════════════════════════════
//  UTILITY / HELPER FUNCTIONS
// ═══════════════════════════════════════════════════════════════

/**
 * Reads a value from the Config sheet by label (Column A).
 */
function getConfigValue(label) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  if (!sheet) {
    throw new Error('Config sheet "' + CONFIG_SHEET_NAME + '" not found. Please create it.');
  }

  var data = sheet.getDataRange().getValues();
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === label.toLowerCase()) {
      return data[i][1];
    }
  }

  throw new Error('Config label "' + label + '" not found in the Config sheet.');
}

/**
 * Writes a value to the Config sheet by label (Column A > Column B).
 */
function setConfigValue(label, value) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  if (!sheet) {
    throw new Error('Config sheet "' + CONFIG_SHEET_NAME + '" not found.');
  }

  var data = sheet.getDataRange().getValues();
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === label.toLowerCase()) {
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }

  throw new Error('Config label "' + label + '" not found — cannot write value.');
}

/**
 * Converts a day name to its getDay() index (0=Sunday ... 6=Saturday).
 */
function getDayIndex(dayName) {
  var days = {
    'sunday': 0, 'monday': 1, 'tuesday': 2, 'wednesday': 3,
    'thursday': 4, 'friday': 5, 'saturday': 6
  };
  var key = String(dayName).trim().toLowerCase();
  if (days[key] === undefined) {
    throw new Error('Invalid day name: "' + dayName + '". Use Sunday through Saturday.');
  }
  return days[key];
}

/**
 * Returns a readable date range, e.g. "Jan 5 to Jan 11, 2025".
 */
function formatDateRange(startDate, endDate) {
  var months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];

  var sMonth = months[startDate.getMonth()];
  var sDay   = startDate.getDate();
  var eMonth = months[endDate.getMonth()];
  var eDay   = endDate.getDate();
  var eYear  = endDate.getFullYear();

  return sMonth + ' ' + sDay + ' to ' + eMonth + ' ' + eDay + ', ' + eYear;
}

/**
 * Cleans a cell value and returns a number.
 * Handles blanks, "O" (letter), whitespace, and non-numeric values.
 */
function sanitizeNumeric(value) {
  if (value === null || value === undefined) return 0;

  var str = String(value).trim();

  if (str === '' || str.toUpperCase() === 'O') return 0;

  var num = parseFloat(str);
  if (isNaN(num)) {
    Logger.log('Warning: Non-numeric value encountered and treated as 0: "' + str + '"');
    return 0;
  }
  return num;
}