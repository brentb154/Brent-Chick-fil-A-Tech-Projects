/**
 * ============================================================
 * TRAINING TRACKING SYSTEM - UI Functions
 * ============================================================
 * Sidebar and dialog launchers, data getters for HTML UIs,
 * and action handlers called from the HTML front-ends.
 */

// -- Deduplication Sidebar ------------------------------------

function showDeduplicationSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('DeduplicationSidebar')
    .setTitle('Review Duplicate Names')
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Returns all pending duplicate suggestions for the sidebar.
 */
function getPendingDuplicates() {
  var dedupSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Name Deduplication');
  if (dedupSheet.getLastRow() < 2) return [];

  var data    = dedupSheet.getDataRange().getValues();
  var pending = [];

  data.slice(1).forEach(function (row, index) {
    if (String(row[4]).trim() === 'Pending') {
      pending.push({
        rowIndex:  index + 2, // +2 for header + 0-indexing
        canonical: row[0],
        variants:  row[1],
        count:     row[2],
        action:    row[3]
      });
    }
  });

  return pending;
}

/**
 * Applies a merge or ignore action from the sidebar.
 */
function applyDedupAction(rowIndex, action) {
  var dedupSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Name Deduplication');
  var row        = dedupSheet.getRange(rowIndex, 1, 1, 5).getValues()[0];

  var canonicalName = String(row[0]).trim();
  var variantName   = String(row[1]).trim();

  if (action === 'Merge') {
    var count = mergeDuplicateNames(canonicalName, variantName);
    return 'Merged "' + variantName + '" -> "' + canonicalName + '" (' + count + ' entries updated)';
  } else if (action === 'Ignore') {
    dedupSheet.getRange(rowIndex, 4).setValue('Ignore');
    dedupSheet.getRange(rowIndex, 5).setValue('Completed');
    return 'Ignored suggestion for "' + variantName + '"';
  }

  return 'Unknown action';
}


// -- Alert Settings Dialog ------------------------------------

function showAlertSettings() {
  var html = HtmlService.createHtmlOutputFromFile('AlertSettingsDialog')
    .setWidth(600)
    .setHeight(420);
  SpreadsheetApp.getUi().showModalDialog(html, 'Alert Settings');
}

/**
 * Returns current alert settings for the dialog.
 */
function getAlertSettings() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var settingsSheet = ss.getSheetByName('Alert Settings');

  // Create sheet if missing
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet('Alert Settings');
  }

  // Check if there is actual data in row 2, column A (not just formatting)
  var hasData = false;
  if (settingsSheet.getLastRow() >= 2) {
    var checkVal = settingsSheet.getRange('A2').getValue();
    if (checkVal && String(checkVal).trim() !== '') {
      hasData = true;
    }
  }

  // Populate defaults if no real data
  if (!hasData) {
    settingsSheet.clear();

    var headers = [['Alert Type', 'Enabled', 'Recipient 1 Email', 'Recipient 2 Email', 'Recipient 3 Email', 'Send Time']];
    var defaults = [
      ['Daily Training Log Reminder',    true, '', '', '', '20:00'],
      ['Trainee Inactive 3+ Days',       true, '', '', '', '06:00'],
      ['Position Completion Milestone',   true, '', '', '', ''],
      ['Trainee Ready for Certification', true, '', '', '', ''],
      ['Duplicate Name Detected',         true, '', '', '', '']
    ];

    settingsSheet.getRange(1, 1, 1, 6).setValues(headers).setFontWeight('bold').setBackground('#4285f4').setFontColor('white');
    settingsSheet.getRange(2, 1, defaults.length, 6).setValues(defaults);
    settingsSheet.getRange(2, 2, defaults.length, 1).insertCheckboxes();
    settingsSheet.autoResizeColumns(1, 6);

    SpreadsheetApp.flush(); // Force write before reading back
  }

  var data     = settingsSheet.getDataRange().getValues();
  var settings = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    // Skip rows where alert type is blank
    if (!row[0] || String(row[0]).trim() === '') continue;

    settings.push({
      rowIndex:   i + 1,
      alertType:  String(row[0]),
      enabled:    row[1] === true || String(row[1]) === 'TRUE',
      recipient1: String(row[2] || ''),
      recipient2: String(row[3] || ''),
      recipient3: String(row[4] || ''),
      sendTime:   String(row[5] || '')
    });
  }

  return settings;
}

/**
 * Saves updated alert settings from the dialog.
 */
function saveAlertSettings(settings) {
  var settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Alert Settings');

  settings.forEach(function (setting) {
    settingsSheet.getRange(setting.rowIndex, 1, 1, 6).setValues([[
      setting.alertType,
      setting.enabled,
      setting.recipient1,
      setting.recipient2,
      setting.recipient3,
      setting.sendTime
    ]]);
  });

  return 'Settings saved successfully!';
}
