/**
 * ============================================================
 * CHECKLIST SYSTEM - Data Cleanup
 * ============================================================
 * Each morning, rows older than Settings "Keep data for (days)"
 * (60 = about 2 months) are deleted from the log tabs below, so
 * the sheet keeps a rolling window instead of growing forever.
 * Blank = keep everything. Nothing is copied first: deleted rows
 * are gone.
 *
 * Checklists, Positions, Items and Settings are never touched.
 * Photo files have their own setting ("Keep photos for (days)").
 * Form Map trims itself to 14 days when forms are rebuilt.
 */

var MIN_KEEP_DATA_DAYS = 30; // the weekly scorecard and the 28-day photo checks need this much

// Log tab -> the column holding each row's date
var CLEANUP_TABS = [
  { name: TABS.status, date: 'Date' },
  { name: TABS.submissions, date: 'Business date' },
  { name: TABS.results, date: 'Business date' },
  { name: TABS.alerts, date: 'Date' },
  { name: TABS.photos, date: 'Business date' }
];

// Under the script lock, so no submission or status check is changing rows meanwhile
function clearOldData_(cfg, today) {
  if (!cfg.keepDataDays) return;
  var cutoff = addDays_(today, -cfg.keepDataDays);
  withLock_(function () {
    CLEANUP_TABS.forEach(function (t) { deleteOldRows_(t.name, t.date, cutoff); });
    SpreadsheetApp.flush();
  });
}

// Deletes rows dated before cutoff, one block of neighboring rows at a time from the bottom up so
// row numbers above stay right. Rows with no readable date are kept. Returns how many were deleted.
function deleteOldRows_(tabName, dateHeader, cutoff) {
  var tab = readTab_(tabName);
  var col = colOrThrow_(tab, dateHeader);
  var old = [];
  tab.display.forEach(function (r, i) {
    var d = toDateKey_(r[col]);
    if (d && d < cutoff) old.push(i + 2);
  });
  if (!old.length) return 0;
  // Sheets won't delete every row under the header, so keep a spare blank row at the bottom
  if (old.length >= tab.sheet.getMaxRows() - 1) tab.sheet.insertRowAfter(tab.sheet.getMaxRows());
  var k = old.length - 1;
  while (k >= 0) {
    var end = old[k];
    var start = end;
    while (k > 0 && old[k - 1] === start - 1) {
      start--;
      k--;
    }
    tab.sheet.deleteRows(start, end - start + 1);
    k--;
  }
  return old.length;
}
