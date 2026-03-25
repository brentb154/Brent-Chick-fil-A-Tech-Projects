/**
 * Schedule Variance Analyzer — Historical Reports
 *
 * Six reports generated from the History tab, accessible through the custom menu.
 * Each report writes to the "Report Output" tab, clearing it first.
 */

/* ── Shared: read history into structured rows ── */

function readHistoryRows_(ss) {
  var sheet = ss.getSheetByName(TABS.HISTORY);
  if (!sheet || sheet.getLastRow() < 2) return [];
  var data = sheet.getDataRange().getValues();
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    rows.push({
      runDate:       String(data[i][0] || ''),
      weekLabel:     String(data[i][1] || ''),
      location:      String(data[i][2] || ''),
      name:          String(data[i][3] || ''),
      schedDays:     data[i][4] || 0,
      workedDays:    data[i][5] || 0,
      absentDays:    data[i][6] || 0,
      unschedDays:   data[i][7] || 0,
      reliability:   data[i][8],
      startVar:      data[i][9] || 0,
      endVar:        data[i][10] || 0,
      totalVar:      data[i][11] || 0,
      avgVar:        data[i][12] || 0,
      earlyIn:       data[i][13] || 0,
      lateIn:        data[i][14] || 0,
      lateOut:       data[i][15] || 0,
      earlyOut:      data[i][16] || 0,
      midnightCount: data[i][17] || 0,
      pattern:       String(data[i][18] || ''),
      swapCount:     data[i][19] || 0,
      otHours:       data[i][20] || 0,
      lateCount:     data[i][21] || 0
    });
  }
  return rows;
}

function getReportSheet_(ss) {
  var sheet = ss.getSheetByName(TABS.REPORT_OUTPUT);
  if (!sheet) sheet = ss.insertSheet(TABS.REPORT_OUTPUT);
  sheet.clearContents();
  sheet.clearFormats();
  ss.setActiveSheet(sheet);
  return sheet;
}

function getDistinctWeeks_(rows) {
  var set = {};
  rows.forEach(function(r) { if (r.weekLabel) set[r.weekLabel] = true; });
  return Object.keys(set).sort();
}

/* ═══════════════════════════════════════════════════════
   1. Worst Offenders — Rolling
   ═══════════════════════════════════════════════════════ */

function reportWorstOffenders() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var rows = readHistoryRows_(ss);
  if (!rows.length) { ui.alert('No history data. Run an analysis first.'); return; }

  var resp = ui.prompt('Rolling Window', 'Enter number of weeks to look back (4, 8, or 12):', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  var windowSize = parseInt(resp.getResponseText()) || 4;

  var allWeeks = getDistinctWeeks_(rows);
  var recentWeeks = allWeeks.slice(-windowSize);

  var filtered = rows.filter(function(r) { return recentWeeks.indexOf(r.weekLabel) >= 0; });

  // Aggregate per employee
  var empMap = {};
  filtered.forEach(function(r) {
    if (!empMap[r.name]) empMap[r.name] = { name: r.name, schedDays: 0, workedDays: 0, absentDays: 0, totalVar: 0, midnightCount: 0, weeks: 0, otHours: 0, lateCount: 0, weekSet: {} };
    var e = empMap[r.name];
    e.schedDays += r.schedDays;
    e.workedDays += r.workedDays;
    e.absentDays += r.absentDays;
    e.totalVar += r.totalVar;
    e.midnightCount += r.midnightCount;
    e.otHours += r.otHours;
    e.lateCount += r.lateCount;
    if (!e.weekSet[r.weekLabel]) { e.weekSet[r.weekLabel] = true; e.weeks++; }
  });

  var list = Object.keys(empMap).map(function(k) {
    var e = empMap[k];
    e.reliability = e.schedDays > 0 ? Math.round(e.workedDays / e.schedDays * 1000) / 10 : 0;
    return e;
  });
  list.sort(function(a, b) { return a.reliability - b.reliability; });

  var sheet = getReportSheet_(ss);
  var row = 1;

  sheet.getRange(row, 1).setValue('WORST OFFENDERS — ROLLING ' + windowSize + ' WEEKS').setFontSize(14).setFontWeight('bold');
  row++;
  sheet.getRange(row, 1).setValue('Covering: ' + recentWeeks[0] + ' through ' + recentWeeks[recentWeeks.length - 1]).setFontSize(10).setFontColor('#666');
  row += 2;

  var headers = ['Employee', 'Weeks Present', 'Sched Days', 'Worked Days', 'Absent', 'Reliability %', 'Total Var (min)', 'OT Hours', 'Late Count', 'Missed Clock-Outs'];
  sheet.getRange(row, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#F3F4F6');
  row++;

  list.forEach(function(e) {
    sheet.getRange(row, 1, 1, headers.length).setValues([[
      e.name, e.weeks, e.schedDays, e.workedDays, e.absentDays,
      e.reliability + '%', e.totalVar, e.otHours, e.lateCount, e.midnightCount
    ]]);
    if (e.reliability < 75) sheet.getRange(row, 6).setFontColor('#DC2626');
    else if (e.reliability < 90) sheet.getRange(row, 6).setFontColor('#D97706');
    else sheet.getRange(row, 6).setFontColor('#059669');
    row++;
  });

  sheet.autoResizeColumns(1, headers.length);
  ui.alert('Report generated on the "Report Output" tab.');
}

/* ═══════════════════════════════════════════════════════
   2. Individual Employee Timeline
   ═══════════════════════════════════════════════════════ */

function reportEmployeeTimeline() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var rows = readHistoryRows_(ss);
  if (!rows.length) { ui.alert('No history data.'); return; }

  var resp = ui.prompt('Employee Timeline', 'Enter the employee\'s full name (exact match):', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  var targetName = resp.getResponseText().trim();
  if (!targetName) { ui.alert('Name is required.'); return; }

  var filtered = rows.filter(function(r) { return r.name.toLowerCase() === targetName.toLowerCase(); });
  if (!filtered.length) {
    // Try partial match
    filtered = rows.filter(function(r) { return r.name.toLowerCase().indexOf(targetName.toLowerCase()) >= 0; });
  }
  if (!filtered.length) { ui.alert('No records found for "' + targetName + '".'); return; }

  var actualName = filtered[0].name;
  filtered.sort(function(a, b) { return a.weekLabel < b.weekLabel ? -1 : a.weekLabel > b.weekLabel ? 1 : 0; });

  var sheet = getReportSheet_(ss);
  var row = 1;

  sheet.getRange(row, 1).setValue('EMPLOYEE TIMELINE: ' + actualName).setFontSize(14).setFontWeight('bold');
  row++;
  sheet.getRange(row, 1).setValue(filtered.length + ' weekly records found').setFontSize(10).setFontColor('#666');
  row += 2;

  var headers = ['Week', 'Location', 'Sched Days', 'Worked', 'Absent', 'Reliability %', 'Total Var', 'Avg/Day', 'OT Hrs', 'Late Count', 'Missed Punches', 'Pattern'];
  sheet.getRange(row, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#F3F4F6');
  row++;

  filtered.forEach(function(r) {
    var rel = r.schedDays > 0 ? Math.round(r.workedDays / r.schedDays * 1000) / 10 : '';
    sheet.getRange(row, 1, 1, headers.length).setValues([[
      r.weekLabel, r.location, r.schedDays, r.workedDays, r.absentDays,
      rel ? rel + '%' : 'N/A', formatVariance(r.totalVar), formatVariance(r.avgVar),
      r.otHours, r.lateCount, r.midnightCount, r.pattern
    ]]);
    if (r.absentDays > 0) sheet.getRange(row, 5).setFontColor('#DC2626').setFontWeight('bold');
    if (r.pattern === 'Late In' || r.pattern === 'Works Less' || r.pattern === 'Leaves Early') {
      sheet.getRange(row, 12).setFontColor('#DC2626');
    } else if (r.pattern === 'On Track') {
      sheet.getRange(row, 12).setFontColor('#059669');
    }
    row++;
  });

  // Summary
  row++;
  sheet.getRange(row, 1).setValue('SUMMARY').setFontWeight('bold').setFontSize(11);
  row++;

  var totalSched = 0, totalWorked = 0, totalAbsent = 0, totalTV = 0, totalOT = 0, totalLate = 0, totalMid = 0;
  filtered.forEach(function(r) {
    totalSched += r.schedDays; totalWorked += r.workedDays; totalAbsent += r.absentDays;
    totalTV += r.totalVar; totalOT += r.otHours; totalLate += r.lateCount; totalMid += r.midnightCount;
  });
  var overallRel = totalSched > 0 ? Math.round(totalWorked / totalSched * 1000) / 10 : 0;

  var summaryStats = [
    ['Total Weeks', filtered.length],
    ['Overall Reliability', overallRel + '%'],
    ['Total Absent Days', totalAbsent],
    ['Total Net Variance', formatVariance(totalTV)],
    ['Total OT Hours', totalOT],
    ['Total Late Arrivals', totalLate],
    ['Total Missed Clock-Outs', totalMid]
  ];
  summaryStats.forEach(function(s) {
    sheet.getRange(row, 1).setValue(s[0]).setFontWeight('bold').setFontColor('#555');
    sheet.getRange(row, 2).setValue(s[1]);
    row++;
  });

  sheet.autoResizeColumns(1, headers.length);
  ui.alert('Timeline for "' + actualName + '" generated on the "Report Output" tab.');
}

/* ═══════════════════════════════════════════════════════
   3. OT Trend by Location
   ═══════════════════════════════════════════════════════ */

function reportOTTrend() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var rows = readHistoryRows_(ss);
  if (!rows.length) { ui.alert('No history data.'); return; }

  var allWeeks = getDistinctWeeks_(rows);

  // Build week × location → OT totals
  var grid = {};
  var locSet = {};
  rows.forEach(function(r) {
    locSet[r.location] = true;
    var k = r.weekLabel + '|||' + r.location;
    if (!grid[k]) grid[k] = { schedOT: 0, actualOT: 0 };
    grid[k].actualOT += r.otHours;
  });

  var locations = Object.keys(locSet).sort();

  var sheet = getReportSheet_(ss);
  var row = 1;

  sheet.getRange(row, 1).setValue('OT TREND BY LOCATION').setFontSize(14).setFontWeight('bold');
  row += 2;

  // Header row
  var headers = ['Week'];
  locations.forEach(function(loc) {
    headers.push(loc + ' (OT hrs)');
  });
  headers.push('Combined OT');
  sheet.getRange(row, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#F3F4F6');
  row++;

  allWeeks.forEach(function(wk) {
    var vals = [wk];
    var combined = 0;
    locations.forEach(function(loc) {
      var k = wk + '|||' + loc;
      var ot = grid[k] ? grid[k].actualOT : 0;
      combined += ot;
      vals.push(Math.round(ot * 10) / 10);
    });
    vals.push(Math.round(combined * 10) / 10);
    sheet.getRange(row, 1, 1, vals.length).setValues([vals]);
    row++;
  });

  sheet.autoResizeColumns(1, headers.length);
  ui.alert('OT Trend report generated on the "Report Output" tab.');
}

/* ═══════════════════════════════════════════════════════
   4. Chronic Lateness Leaderboard
   ═══════════════════════════════════════════════════════ */

function reportChronicLateness() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var rows = readHistoryRows_(ss);
  if (!rows.length) { ui.alert('No history data.'); return; }

  var allWeeks = getDistinctWeeks_(rows);
  var recentWeeks = allWeeks.slice(-4);
  var filtered = rows.filter(function(r) { return recentWeeks.indexOf(r.weekLabel) >= 0; });

  var empMap = {};
  filtered.forEach(function(r) {
    if (!empMap[r.name]) empMap[r.name] = { name: r.name, lateCount: 0, lateInMin: 0, matchedDays: 0, weeks: 0, weekSet: {} };
    var e = empMap[r.name];
    e.lateCount += r.lateCount;
    e.lateInMin += r.lateIn;
    e.matchedDays += r.workedDays;
    if (!e.weekSet[r.weekLabel]) { e.weekSet[r.weekLabel] = true; e.weeks++; }
  });

  var list = Object.keys(empMap).map(function(k) {
    var e = empMap[k];
    e.avgLateMin = e.lateCount > 0 ? Math.round(e.lateInMin / e.lateCount) : 0;
    e.frequency = e.matchedDays > 0 ? Math.round(e.lateCount / e.matchedDays * 1000) / 10 : 0;
    return e;
  }).filter(function(e) { return e.lateCount > 0; });

  list.sort(function(a, b) { return b.lateCount - a.lateCount; });

  var sheet = getReportSheet_(ss);
  var row = 1;

  sheet.getRange(row, 1).setValue('CHRONIC LATENESS LEADERBOARD — LAST 4 WEEKS').setFontSize(14).setFontWeight('bold');
  row++;
  sheet.getRange(row, 1).setValue(recentWeeks[0] + ' through ' + recentWeeks[recentWeeks.length - 1]).setFontSize(10).setFontColor('#666');
  row += 2;

  var headers = ['Rank', 'Employee', 'Late Arrivals', 'Total Late (min)', 'Avg Late (min)', 'Matched Days', 'Late Frequency %', 'Weeks Present'];
  sheet.getRange(row, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#F3F4F6');
  row++;

  list.forEach(function(e, i) {
    sheet.getRange(row, 1, 1, headers.length).setValues([[
      i + 1, e.name, e.lateCount, e.lateInMin, e.avgLateMin, e.matchedDays, e.frequency + '%', e.weeks
    ]]);
    if (e.frequency >= 40) sheet.getRange(row, 7).setFontColor('#DC2626');
    row++;
  });

  if (!list.length) {
    sheet.getRange(row, 1).setValue('No late arrivals recorded in the last 4 weeks.').setFontColor('#059669');
  }

  sheet.autoResizeColumns(1, headers.length);
  ui.alert('Chronic Lateness report generated on the "Report Output" tab.');
}

/* ═══════════════════════════════════════════════════════
   5. Missed Clock-Out Frequency
   ═══════════════════════════════════════════════════════ */

function reportMissedClockOuts() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var rows = readHistoryRows_(ss);
  if (!rows.length) { ui.alert('No history data.'); return; }

  var empMap = {};
  rows.forEach(function(r) {
    if (r.midnightCount > 0) {
      if (!empMap[r.name]) empMap[r.name] = { name: r.name, total: 0, weeks: [], byWeek: {} };
      empMap[r.name].total += r.midnightCount;
      if (!empMap[r.name].byWeek[r.weekLabel]) {
        empMap[r.name].byWeek[r.weekLabel] = 0;
        empMap[r.name].weeks.push(r.weekLabel);
      }
      empMap[r.name].byWeek[r.weekLabel] += r.midnightCount;
    }
  });

  var list = Object.keys(empMap).map(function(k) { return empMap[k]; });
  list.sort(function(a, b) { return b.total - a.total; });

  var sheet = getReportSheet_(ss);
  var row = 1;

  sheet.getRange(row, 1).setValue('MISSED CLOCK-OUT FREQUENCY — ALL TIME').setFontSize(14).setFontWeight('bold');
  row += 2;

  if (!list.length) {
    sheet.getRange(row, 1).setValue('No missed clock-outs recorded in history.').setFontColor('#059669');
    ui.alert('Report generated. No missed clock-outs found.');
    return;
  }

  // Summary
  var totalMissed = list.reduce(function(t, e) { return t + e.total; }, 0);
  sheet.getRange(row, 1).setValue('Total missed clock-outs across all history: ' + totalMissed).setFontWeight('bold');
  row++;
  sheet.getRange(row, 1).setValue(list.length + ' employees affected').setFontColor('#666');
  row += 2;

  var headers = ['Employee', 'Total Missed', 'Weeks Affected', 'Week Details'];
  sheet.getRange(row, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#F3F4F6');
  row++;

  list.forEach(function(e) {
    var details = e.weeks.map(function(w) { return w + ' (' + e.byWeek[w] + 'x)'; }).join(', ');
    sheet.getRange(row, 1, 1, headers.length).setValues([[e.name, e.total, e.weeks.length, details]]);
    if (e.total >= 3) sheet.getRange(row, 2).setFontColor('#DC2626').setFontWeight('bold');
    else if (e.total >= 2) sheet.getRange(row, 2).setFontColor('#D97706');
    row++;
  });

  sheet.autoResizeColumns(1, headers.length);
  ui.alert('Missed Clock-Out Frequency report generated on the "Report Output" tab.');
}

/* ═══════════════════════════════════════════════════════
   6. OT Reduction Trend (month-over-month)
   ═══════════════════════════════════════════════════════ */

function reportOTReduction() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var rows = readHistoryRows_(ss);
  if (!rows.length) { ui.alert('No history data.'); return; }

  // Group by month from runDate (M/D/YYYY format)
  var monthMap = {};
  rows.forEach(function(r) {
    var parts = r.runDate.match(/(\d+)\/(\d+)\/(\d+)/);
    if (!parts) return;
    var monthKey = parts[3] + '-' + String(parseInt(parts[1])).padStart(2, '0');
    if (!monthMap[monthKey]) monthMap[monthKey] = { schedOT: 0, actualOT: 0, empCount: 0, otCount: 0 };
    monthMap[monthKey].actualOT += r.otHours;
    monthMap[monthKey].empCount++;
    if (r.otHours > 0) monthMap[monthKey].otCount++;
  });

  var months = Object.keys(monthMap).sort();
  if (months.length < 1) { ui.alert('Not enough data for OT trend.'); return; }

  var sheet = getReportSheet_(ss);
  var row = 1;

  sheet.getRange(row, 1).setValue('OT REDUCTION TREND — MONTH OVER MONTH').setFontSize(14).setFontWeight('bold');
  row++;
  sheet.getRange(row, 1).setValue('Tracks whether scheduling discipline is improving over time').setFontSize(10).setFontColor('#666');
  row += 2;

  var headers = ['Month', 'Total OT Hours', 'Employees with OT', 'Δ from Previous'];
  sheet.getRange(row, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#F3F4F6');
  row++;

  var prevOT = null;
  months.forEach(function(m) {
    var d = monthMap[m];
    var ot = Math.round(d.actualOT * 10) / 10;
    var delta = prevOT !== null ? Math.round((ot - prevOT) * 10) / 10 : '';
    var deltaStr = delta !== '' ? (delta > 0 ? '+' + delta + ' hrs' : delta + ' hrs') : '—';

    var MONTH_NAMES = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    var parts = m.split('-');
    var label = MONTH_NAMES[parseInt(parts[1]) - 1] + ' ' + parts[0];

    sheet.getRange(row, 1, 1, headers.length).setValues([[label, ot, d.otCount, deltaStr]]);

    if (delta !== '' && delta < 0) {
      sheet.getRange(row, 4).setFontColor('#059669');
    } else if (delta !== '' && delta > 0) {
      sheet.getRange(row, 4).setFontColor('#DC2626');
    }

    prevOT = ot;
    row++;
  });

  // Overall trend line
  if (months.length >= 2) {
    row++;
    var first = monthMap[months[0]].actualOT;
    var last = monthMap[months[months.length - 1]].actualOT;
    var change = Math.round((last - first) * 10) / 10;
    var direction = change < 0 ? 'DECREASED' : change > 0 ? 'INCREASED' : 'UNCHANGED';
    sheet.getRange(row, 1).setValue('Overall: OT ' + direction + ' by ' + Math.abs(change) + ' hrs from ' + months[0] + ' to ' + months[months.length - 1])
      .setFontWeight('bold').setFontColor(change <= 0 ? '#059669' : '#DC2626');
  }

  sheet.autoResizeColumns(1, headers.length);
  ui.alert('OT Reduction Trend report generated on the "Report Output" tab.');
}
