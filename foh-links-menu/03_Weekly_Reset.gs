// ============================================================
// FOH Links Menu — Sunday Weekly Reset
// ============================================================

var RESET_DAYS = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

// Desired tab order — tabs not in this list stay at the end
var TAB_ORDER = [
  'Links', 'Manager in Charge',
  'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday',
  'Monday (Reset)', 'Tuesday (Reset)', 'Wednesday (Reset)',
  'Thursday (Reset)', 'Friday (Reset)', 'Saturday (Reset)',
  'DonationLog', 'Settings'
];

/**
 * Resets a single day sheet from its (Reset) template.
 * Deletes the day tab, copies the reset tab, renames it.
 * True copy-paste — all formatting, borders, merges, everything.
 * Returns null on success, or an error string on failure.
 */
function resetSingleDay_(ss, day) {
  var daySheet = ss.getSheetByName(day);
  var resetSheet = ss.getSheetByName(day + ' (Reset)');

  if (!daySheet) return day + ': tab not found';
  if (!resetSheet) return day + ' (Reset): tab not found';

  // Remember the position of the day tab
  var dayIndex = daySheet.getIndex();

  // Delete the old day sheet
  ss.deleteSheet(daySheet);

  // Copy the reset template (creates a new tab named "Copy of Monday (Reset)")
  var newSheet = resetSheet.copyTo(ss);

  // Rename to the day name and move to original position
  newSheet.setName(day);
  ss.moveActiveSheet(dayIndex);

  return null; // success
}

/**
 * Reorders all tabs to match TAB_ORDER. Tabs not in the list go to the end.
 */
function reorderTabs() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var position = 1;

  TAB_ORDER.forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    if (sheet) {
      ss.setActiveSheet(sheet);
      ss.moveActiveSheet(position);
      position++;
    }
  });

  SpreadsheetApp.flush();
}

/**
 * Copies each "(Reset)" tab over the matching day tab.
 * Preserves formatting, clears all names. Safe to re-run.
 */
function resetWeeklySheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var errors = [];
  var resetCount = 0;

  RESET_DAYS.forEach(function(day) {
    var err = resetSingleDay_(ss, day);
    if (err) {
      errors.push(err);
    } else {
      resetCount++;
    }
  });

  // Reorder tabs after reset (copyTo puts new sheets at the end)
  reorderTabs();

  SpreadsheetApp.flush();

  // Alert on errors (email if running from trigger)
  if (errors.length > 0) {
    var msg = 'Weekly reset completed (' + resetCount + '/' + RESET_DAYS.length + ' sheets).\n\nIssues:\n' + errors.join('\n');
    try {
      var alertEmail = getSetting('Alert Email') || Session.getActiveUser().getEmail();
      GmailApp.sendEmail(
        alertEmail,
        'FOH Sheet Reset — Errors',
        msg
      );
    } catch(e) {
      // If no email available, log it
      Logger.log(msg);
    }
  }

  return { reset: resetCount, errors: errors };
}

/**
 * Sets up the Sunday 9am trigger. Safe to re-run — removes existing reset triggers first.
 */
function setupResetTrigger() {
  deleteResetTriggers();

  var resetDay = getSetting('Reset Day');
  var resetHour = getSetting('Reset Hour');

  ScriptApp.newTrigger('resetWeeklySheets')
    .timeBased()
    .onWeekDay(parseDaySetting_(resetDay))
    .atHour(parseHourSetting_(resetHour))
    .create();

  SpreadsheetApp.getUi().alert('Auto-reset trigger created: ' + resetDay + ' at ' + resetHour + '.\n\nSheets will reset automatically. To change the day or time, edit the Settings tab and re-run this.');
}

/**
 * Removes all existing reset triggers.
 */
function deleteResetTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'resetWeeklySheets') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

/**
 * Test reset — only resets Monday. Use to verify before enabling the full trigger.
 */
function testResetMonday() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(
    'Test: Reset Monday Only?',
    'This will overwrite the Monday sheet with the Monday (Reset) template.\n\nAre you sure?',
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var err = resetSingleDay_(ss, 'Monday');
  reorderTabs();
  SpreadsheetApp.flush();

  if (err) {
    ui.alert('Monday reset failed:\n' + err);
  } else {
    ui.alert('Monday sheet has been reset and tabs reordered. Check it and verify everything looks right before enabling the full Sunday trigger.');
  }
}

/**
 * Manual reset — runs immediately with confirmation.
 */
function manualReset() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(
    'Reset All Day Sheets?',
    'This will overwrite Monday through Saturday with the (Reset) templates. Names will be cleared.\n\nAre you sure?',
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;

  var result = resetWeeklySheets();
  if (result.errors.length > 0) {
    ui.alert('Reset completed with issues:\n' + result.errors.join('\n'));
  } else {
    ui.alert('All ' + result.reset + ' day sheets have been reset.');
  }
}
