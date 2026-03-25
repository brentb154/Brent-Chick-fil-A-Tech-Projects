/**
 * Schedule Variance Analyzer — History Tab
 *
 * Appends per-employee per-week roll-up rows after each analysis run.
 * If data for the same week label already exists, replaces it (upsert).
 * Powers chronic flag logic and historical reports.
 */

var HISTORY_HEADERS = [
  'Run Date', 'Week Label', 'Location', 'Employee',
  'Sched Days', 'Worked Days', 'Absent Days', 'Unscheduled',
  'Reliability %', 'Start Var', 'End Var', 'Total Var', 'Avg/Day',
  'Early In', 'Late In', 'Stayed Late', 'Left Early',
  'Missed Clock-Outs', 'Pattern', 'Swap Flags',
  'OT Hours', 'Late Count', 'Sched Hours', 'Actual Hours'
];

function appendHistory(ss, allRows, crossLocOT) {
  var sheet = ss.getSheetByName(TABS.HISTORY);
  if (!sheet) sheet = ss.insertSheet(TABS.HISTORY);

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, HISTORY_HEADERS.length).setValues([HISTORY_HEADERS]);
    sheet.getRange(1, 1, 1, HISTORY_HEADERS.length).setFontWeight('bold').setBackground('#F3F4F6');
    sheet.setFrozenRows(1);
  }

  // Determine which week labels are being written
  var incomingWeeks = {};
  allRows.forEach(function(r) { if (r.weekLabel) incomingWeeks[r.weekLabel] = true; });
  var weekLabelList = Object.keys(incomingWeeks);

  // Remove existing rows for those week labels (prevents duplication on re-run)
  if (weekLabelList.length > 0 && sheet.getLastRow() > 1) {
    var data = sheet.getDataRange().getValues();
    var rowsToDelete = [];
    for (var i = 1; i < data.length; i++) {
      var wl = String(data[i][1] || '').trim();
      if (weekLabelList.indexOf(wl) >= 0) {
        rowsToDelete.push(i + 1);
      }
    }
    for (var j = rowsToDelete.length - 1; j >= 0; j--) {
      sheet.deleteRow(rowsToDelete[j]);
    }
  }

  var runDate = new Date().toLocaleDateString();
  var rows = allRows.map(function(r) {
    var empOT = crossLocOT && crossLocOT.employees && crossLocOT.employees[r.name];
    var otHrs = empOT ? empOT.otHours : (r.otHours || 0);

    return [
      runDate,
      r.weekLabel || '',
      r.locationName || '',
      r.name,
      r.scheduledDays,
      r.workedDays,
      r.absentDays,
      r.unscheduledDays,
      r.reliability !== null ? r.reliability : '',
      r.startVar,
      r.endVar,
      r.totalVar,
      r.avgVar,
      r.earlyIn,
      r.lateIn,
      r.lateOut,
      r.earlyOut,
      r.midnightCount,
      r.pattern.label,
      r.swapCount,
      otHrs,
      r.lateCount || 0,
      r.scheduledHours || 0,
      r.actualHours || 0
    ];
  });

  if (rows.length) {
    var startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, rows.length, HISTORY_HEADERS.length).setValues(rows);
  }
}

/* ── Chronic Flag Lookback ── */

/**
 * Scans the History tab for employees flagged in N+ of the last M weeks.
 * A flag means: absent 1+ days, or pattern is 'Late In' / 'Works Less' / 'Leaves Early'.
 *
 * Returns array of { name, reason, weeksTriggered }
 */
function getChronicFlags(ss, cfg) {
  var historySheet = ss.getSheetByName(TABS.HISTORY);
  if (!historySheet || historySheet.getLastRow() < 2) return [];

  var data = historySheet.getDataRange().getValues();
  var window   = cfg.CHRONIC_WINDOW  || 3;
  var trigger  = cfg.CHRONIC_TRIGGER || 2;

  // Collect distinct week labels from History, sorted
  var weekSet = {};
  for (var i = 1; i < data.length; i++) {
    var wl = String(data[i][1] || '').trim();
    if (wl) weekSet[wl] = true;
  }
  var allWeeks = Object.keys(weekSet).sort();
  if (allWeeks.length < 1) return [];

  var recentWeeks = allWeeks.slice(-window);

  // Build per-employee per-week flag map
  var empFlags = {};  // { name: { weekLabel: [reasons] } }
  for (var i = 1; i < data.length; i++) {
    var wl   = String(data[i][1] || '').trim();
    var name = String(data[i][3] || '').trim();
    if (!wl || !name) continue;
    if (recentWeeks.indexOf(wl) < 0) continue;

    var absentDays = data[i][6] || 0;
    var pattern    = String(data[i][18] || '');
    var lateCount  = data[i][21] || 0;

    var reasons = [];
    if (absentDays > 0) reasons.push('absent');
    if (pattern === 'Late In' || pattern === 'Works Less' || pattern === 'Leaves Early') reasons.push(pattern.toLowerCase());
    if (lateCount >= 3) reasons.push('frequent lateness');

    if (reasons.length > 0) {
      if (!empFlags[name]) empFlags[name] = {};
      empFlags[name][wl] = reasons;
    }
  }

  var results = [];
  Object.keys(empFlags).forEach(function(name) {
    var flaggedWeeks = Object.keys(empFlags[name]).length;
    if (flaggedWeeks >= trigger) {
      var allReasons = {};
      Object.keys(empFlags[name]).forEach(function(wl) {
        empFlags[name][wl].forEach(function(r) { allReasons[r] = true; });
      });
      results.push({
        name: name,
        reason: Object.keys(allReasons).join(', '),
        weeksTriggered: flaggedWeeks
      });
    }
  });

  results.sort(function(a, b) { return b.weeksTriggered - a.weeksTriggered; });
  return results;
}
