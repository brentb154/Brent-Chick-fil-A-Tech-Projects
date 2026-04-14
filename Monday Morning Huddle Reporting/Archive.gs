/**
 * Monday Huddle Archive — Weekly Snapshot
 *
 * Every Monday at 5 PM, copies all tabs ending in ".ch" into a single
 * "Huddle Archive" tab. Each snapshot is separated by a banner row
 * showing the date and tab name. Nothing is ever deleted.
 *
 * Setup:  Run setupWeeklyArchiveTrigger() once from the script editor
 *         or the menu (Schedule Variance → Setup Weekly Archive).
 * Manual: Run archiveWeeklySnapshot() to take a snapshot right now.
 */

var ARCHIVE_TAB = 'Huddle Archive';
var CH_SUFFIX   = '.ch';

/* ── Main snapshot function ── */

function archiveWeeklySnapshot() {
  try {
    archiveWeeklySnapshot_();
  } catch (e) {
    notifyTriggerFailure_('archiveWeeklySnapshot', e);
    throw e;
  }
}

function archiveWeeklySnapshot_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var archive = ss.getSheetByName(ARCHIVE_TAB);

  if (!archive) {
    archive = ss.insertSheet(ARCHIVE_TAB);
    // Move it to the end so it stays out of the way
    ss.setActiveSheet(archive);
    ss.moveActiveSheet(ss.getNumSheets());
  }

  var sheets = ss.getSheets();
  var chTabs = sheets.filter(function(s) {
    var name = s.getName();
    return name.length > CH_SUFFIX.length &&
           name.substring(name.length - CH_SUFFIX.length).toLowerCase() === CH_SUFFIX;
  });

  if (chTabs.length === 0) {
    Logger.log('No .ch tabs found — nothing to archive.');
    return;
  }

  var now       = new Date();
  var dateStr   = Utilities.formatDate(now, Session.getScriptTimeZone(), 'EEEE, MMMM d, yyyy \'at\' h:mm a');
  var startRow  = archive.getLastRow() + 1;

  // If the archive already has data, add a blank spacer row first
  if (startRow > 1) {
    startRow += 1;
  }

  var BORDER = SpreadsheetApp.BorderStyle.SOLID;

  // ── Master banner: date stamp ──
  var maxCols = getMaxCols_(chTabs);
  if (maxCols < 5) maxCols = 5;

  archive.getRange(startRow, 1, 1, maxCols).merge()
    .setValue('SNAPSHOT — ' + dateStr)
    .setFontSize(14).setFontWeight('bold').setFontColor('#FFFFFF')
    .setBackground('#1a1a2e').setHorizontalAlignment('center');
  startRow++;

  // ── Copy each .ch tab ──
  chTabs.forEach(function(tab) {
    var tabName = tab.getName();
    var lastRow = tab.getLastRow();
    var lastCol = tab.getLastColumn();

    // Tab sub-banner
    archive.getRange(startRow, 1, 1, maxCols).merge()
      .setValue(tabName)
      .setFontSize(12).setFontWeight('bold').setFontColor('#FFFFFF')
      .setBackground('#374151').setHorizontalAlignment('center');
    startRow++;

    if (lastRow === 0 || lastCol === 0) {
      archive.getRange(startRow, 1).setValue('(empty tab)')
        .setFontSize(10).setFontColor('#9CA3AF').setFontStyle('italic');
      startRow += 2;
      return;
    }

    // Pull all data and formatting info
    var sourceRange = tab.getRange(1, 1, lastRow, lastCol);
    var values      = sourceRange.getValues();
    var bgs         = sourceRange.getBackgrounds();
    var colors      = sourceRange.getFontColors();
    var weights     = sourceRange.getFontWeights();
    var formats     = sourceRange.getNumberFormats();

    // Write values
    var destRange = archive.getRange(startRow, 1, lastRow, lastCol);
    destRange.setValues(values);
    destRange.setBackgrounds(bgs);
    destRange.setFontColors(colors);
    destRange.setFontWeights(weights);
    destRange.setNumberFormats(formats);

    // Copy merged regions
    var merges = tab.getRange(1, 1, lastRow, lastCol).getMergedRanges();
    merges.forEach(function(m) {
      var rowOff = m.getRow() - 1;
      var colOff = m.getColumn() - 1;
      var numR   = m.getNumRows();
      var numC   = m.getNumColumns();
      try {
        archive.getRange(startRow + rowOff, 1 + colOff, numR, numC).merge();
      } catch (e) {
        // Skip if merge conflicts with existing content
      }
    });

    startRow += lastRow + 1; // +1 for spacer between tabs
  });

  Logger.log('Archived ' + chTabs.length + ' tab(s) at ' + dateStr);
}

/* ── Trigger setup — run once ── */

function setupWeeklyArchiveTrigger() {
  // Remove any existing archive triggers to avoid duplicates
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t) {
    if (t.getHandlerFunction() === 'archiveWeeklySnapshot') {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('archiveWeeklySnapshot')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(17)        // 5 PM
    .nearMinute(0)     // as close to :00 as possible
    .create();

  SpreadsheetApp.getUi().alert(
    'Archive Trigger Set',
    'A weekly snapshot will run every Monday around 5:00 PM.\n\n' +
    'All tabs ending in ".ch" will be copied into the "' + ARCHIVE_TAB + '" tab.\n' +
    'Nothing is ever deleted — it just keeps growing.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/* ── Remove trigger ── */

function removeWeeklyArchiveTrigger() {
  var removed = 0;
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'archiveWeeklySnapshot') {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });

  SpreadsheetApp.getUi().alert(
    removed > 0
      ? 'Archive trigger removed.'
      : 'No archive trigger was found.'
  );
}

/* ── Helper: widest column count across tabs ── */

function getMaxCols_(tabs) {
  var max = 0;
  tabs.forEach(function(t) {
    var c = t.getLastColumn();
    if (c > max) max = c;
  });
  return max;
}
