/**
 * Schedule Variance Analyzer — Weekly Report Writer
 *
 * Generates the formatted, print-ready Weekly Report tab with:
 * 1. Summary stats (with week-over-week delta)
 * 2. Flags & alerts (midnight, absences, behavioral, chronic)
 * 3. OT summary (cross-location reconciled)
 * 4. Employee roll-up table
 * 5. Day-by-day detail
 */

function writeReport(ss, allRows, weeks, locationData, crossLocOT, chronicFlags, cfg) {
  var sheet = ss.getSheetByName(TABS.REPORT);
  if (!sheet) sheet = ss.insertSheet(TABS.REPORT);
  sheet.clearContents();
  sheet.clearFormats();

  var row = 1;
  var locNames = locationData.map(function(l) { return l.name; }).join(' + ');
  var weekLabels = weeks.map(function(w) { return w.label; }).join(', ');

  // ── Header ──
  sheet.getRange(row, 1).setValue('SCHEDULE VARIANCE REPORT').setFontSize(16).setFontWeight('bold');
  row++;
  sheet.getRange(row, 1).setValue(locNames + ' — ' + weekLabels).setFontSize(11).setFontColor('#666666');
  row++;
  sheet.getRange(row, 1).setValue('Generated: ' + new Date().toLocaleString()).setFontSize(9).setFontColor('#999999');
  row += 2;

  // ── Section 1: Summary Stats ──
  var totalEmp = allRows.length;
  var totalMatched = 0, totalAbsent = 0, totalMidnight = 0, netVar = 0, totalSched = 0;
  allRows.forEach(function(r) {
    totalMatched += r.matchedDays;
    totalAbsent  += r.absentDays;
    totalMidnight += r.midnightCount;
    netVar += r.totalVar;
    totalSched += r.scheduledDays;
  });
  var adherence = totalSched > 0 ? Math.round(totalMatched / totalSched * 1000) / 10 : 0;
  var highVar = allRows.filter(function(r) { return r.absTotal >= cfg.HIGH_VARIANCE_THRESHOLD; }).length;

  var prevAdherence = getPreviousAdherence_(ss);
  var adhDelta = prevAdherence !== null ? (adherence - prevAdherence) : null;

  sheet.getRange(row, 1).setValue('SUMMARY').setFontSize(12).setFontWeight('bold');
  row++;

  var stats = [
    ['Employees Analyzed', totalEmp],
    ['Schedule Adherence', adherence + '%' + (adhDelta !== null ? '  (' + (adhDelta >= 0 ? '+' : '') + adhDelta.toFixed(1) + '% vs last week)' : '')],
    ['Matched Days', totalMatched],
    ['Absences', totalAbsent],
    ['Missed Clock-Outs', totalMidnight],
    ['High Variance (>' + cfg.HIGH_VARIANCE_THRESHOLD + ' min)', highVar],
    ['Net Variance', formatVariance(netVar)],
    ['Cross-Location OT', crossLocOT.totalOTHours + ' hrs actual  /  ' + crossLocOT.totalScheduledOTHours + ' hrs scheduled']
  ];

  stats.forEach(function(s) {
    sheet.getRange(row, 1).setValue(s[0]).setFontWeight('bold').setFontColor('#555555');
    sheet.getRange(row, 2).setValue(s[1]);
    if (s[0] === 'Schedule Adherence' && adhDelta !== null) {
      sheet.getRange(row, 2).setFontColor(adhDelta >= 0 ? '#059669' : '#DC2626');
    }
    row++;
  });
  row++;

  // ── Section 2: Flags & Alerts ──
  sheet.getRange(row, 1).setValue('FLAGS & ALERTS').setFontSize(12).setFontWeight('bold');
  row++;

  // Midnight clock-outs
  var midnighters = [];
  allRows.forEach(function(r) {
    if (r.midnightCount > 0) {
      var days = r.dayList.filter(function(d) { return d.midnightFlag; })
        .map(function(d) { return formatFriendlyDate(d.date); }).join(', ');
      midnighters.push([r.name, r.locationName, days, r.midnightCount]);
    }
  });

  if (midnighters.length) {
    sheet.getRange(row, 1).setValue('Missed Clock-Outs (punched out near midnight)').setFontWeight('bold').setFontColor('#D97706');
    row++;
    midnighters.forEach(function(m) {
      sheet.getRange(row, 1).setValue(m[0]);
      sheet.getRange(row, 2).setValue(m[1]);
      sheet.getRange(row, 3).setValue(m[2]);
      sheet.getRange(row, 4).setValue(m[3] + 'x');
      sheet.getRange(row, 1, 1, 4).setBackground('#FEF3C7');
      row++;
    });
    row++;
  }

  // Absences
  var absents = [];
  allRows.forEach(function(r) {
    if (r.absentDays > 0) {
      var days = r.dayList.filter(function(d) { return d.status === 'absent'; })
        .map(function(d) { return formatFriendlyDate(d.date); }).join(', ');
      absents.push([r.name, r.locationName, days, r.absentDays]);
    }
  });

  if (absents.length) {
    sheet.getRange(row, 1).setValue('Absences (scheduled but no punch)').setFontWeight('bold').setFontColor('#DC2626');
    row++;
    absents.forEach(function(a) {
      sheet.getRange(row, 1).setValue(a[0]);
      sheet.getRange(row, 2).setValue(a[1]);
      sheet.getRange(row, 3).setValue(a[2]);
      sheet.getRange(row, 4).setValue(a[3] + ' day(s)');
      sheet.getRange(row, 1, 1, 4).setBackground('#FEE2E2');
      row++;
    });
    row++;
  }

  // High variance
  var highVarEmps = allRows.filter(function(r) { return r.absTotal >= cfg.HIGH_VARIANCE_THRESHOLD; })
    .sort(function(a, b) { return b.absTotal - a.absTotal; });
  if (highVarEmps.length) {
    sheet.getRange(row, 1).setValue('High Variance (>' + cfg.HIGH_VARIANCE_THRESHOLD + ' min total)').setFontWeight('bold').setFontColor('#DC2626');
    row++;
    highVarEmps.forEach(function(r) {
      sheet.getRange(row, 1).setValue(r.name);
      sheet.getRange(row, 2).setValue(r.locationName);
      sheet.getRange(row, 3).setValue(formatVariance(r.totalVar));
      sheet.getRange(row, 3).setFontColor('#DC2626');
      row++;
    });
    row++;
  }

  // Possible swaps
  var swapEmps = allRows.filter(function(r) { return r.swapCount > 0; });
  if (swapEmps.length) {
    sheet.getRange(row, 1).setValue('Possible Schedule Mismatches / Swaps').setFontWeight('bold').setFontColor('#D97706');
    row++;
    swapEmps.forEach(function(r) {
      sheet.getRange(row, 1).setValue(r.name);
      sheet.getRange(row, 2).setValue(r.locationName);
      sheet.getRange(row, 3).setValue(r.swapCount + ' shift(s) with >' + cfg.SWAP_THRESHOLD + ' min start variance');
      sheet.getRange(row, 1, 1, 3).setBackground('#FFF7ED');
      row++;
    });
    row++;
  }

  // Behavioral flags
  var latePatterns = allRows.filter(function(r) { return r.pattern.type === 'bad'; });
  if (latePatterns.length) {
    sheet.getRange(row, 1).setValue('Behavioral Flags').setFontWeight('bold').setFontColor('#DC2626');
    row++;
    latePatterns.forEach(function(r) {
      sheet.getRange(row, 1).setValue(r.name);
      sheet.getRange(row, 2).setValue(r.locationName);
      sheet.getRange(row, 3).setValue(r.pattern.label);
      sheet.getRange(row, 4).setValue('Reliability: ' + (r.reliability !== null ? r.reliability + '%' : 'N/A'));
      row++;
    });
    row++;
  }

  // Chronic flags (from History lookback)
  if (chronicFlags && chronicFlags.length) {
    sheet.getRange(row, 1).setValue('ONGOING CONCERNS (flagged ' + cfg.CHRONIC_TRIGGER + '+ of last ' + cfg.CHRONIC_WINDOW + ' weeks)')
      .setFontWeight('bold').setFontColor('#7C3AED');
    row++;
    chronicFlags.forEach(function(f) {
      sheet.getRange(row, 1).setValue(f.name);
      sheet.getRange(row, 2).setValue(f.reason);
      sheet.getRange(row, 3).setValue(f.weeksTriggered + ' of last ' + cfg.CHRONIC_WINDOW + ' weeks');
      sheet.getRange(row, 1, 1, 3).setBackground('#F3E8FF');
      row++;
    });
    row++;
  }

  row++;

  // ── Section 3: OT Summary ──
  sheet.getRange(row, 1).setValue('OVERTIME SUMMARY').setFontSize(12).setFontWeight('bold');
  row++;

  Object.keys(crossLocOT.byLocation).forEach(function(loc) {
    var d = crossLocOT.byLocation[loc];
    sheet.getRange(row, 1).setValue(loc).setFontWeight('bold');
    sheet.getRange(row, 2).setValue('Scheduled: ' + d.schedHours + ' hrs');
    sheet.getRange(row, 3).setValue('Actual: ' + d.actualHours + ' hrs');
    row++;
  });

  sheet.getRange(row, 1).setValue('Cross-Location Reconciled OT').setFontWeight('bold').setFontColor('#7C3AED');
  sheet.getRange(row, 2).setValue(crossLocOT.totalOTHours + ' hrs actual OT');
  sheet.getRange(row, 3).setValue(crossLocOT.totalScheduledOTHours + ' hrs scheduled OT');
  row++;

  // Top OT offenders
  var otList = [];
  Object.keys(crossLocOT.employees).forEach(function(name) {
    var e = crossLocOT.employees[name];
    if (e.otHours > 0) otList.push({ name: name, ot: e.otHours, total: e.totalHours, multi: e.isMultiLocation });
  });
  otList.sort(function(a, b) { return b.ot - a.ot; });

  if (otList.length) {
    row++;
    sheet.getRange(row, 1).setValue('Top OT Employees').setFontWeight('bold');
    row++;
    otList.slice(0, 10).forEach(function(e) {
      sheet.getRange(row, 1).setValue(e.name + (e.multi ? ' ★ multi-location' : ''));
      sheet.getRange(row, 2).setValue(e.ot + ' hrs OT');
      sheet.getRange(row, 3).setValue(e.total + ' hrs total');
      if (e.multi) sheet.getRange(row, 1).setFontColor('#7C3AED');
      row++;
    });
  }
  row += 2;

  // ── Section 4: Employee Roll-Up Table ──
  sheet.getRange(row, 1).setValue('EMPLOYEE ROLL-UP').setFontSize(12).setFontWeight('bold');
  row++;

  var headers = [
    'Employee', 'Location', 'Sched Days', 'Worked Days', 'Absent', 'Unsched',
    'Reliability', 'Start Var', 'End Var', 'Total Var', 'Avg/Day',
    'Early In', 'Late In', 'Stayed Late', 'Left Early',
    'Missed Punches', 'OT Hrs', 'Pattern', 'Flags'
  ];
  var headerRange = sheet.getRange(row, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setFontWeight('bold').setBackground('#F3F4F6').setFontSize(9);
  row++;

  allRows.forEach(function(r) {
    var vals = [
      r.name,
      r.locationName,
      r.scheduledDays,
      r.workedDays,
      r.absentDays,
      r.unscheduledDays,
      r.reliability !== null ? r.reliability + '%' : 'N/A',
      formatVariance(r.startVar),
      formatVariance(r.endVar),
      formatVariance(r.totalVar),
      formatVariance(r.avgVar),
      formatVariance(r.earlyIn),
      formatVariance(r.lateIn),
      formatVariance(r.lateOut),
      formatVariance(r.earlyOut),
      r.midnightCount,
      r.otHours,
      r.pattern.label,
      r.swapCount > 0 ? r.swapCount + ' swap(s)' : ''
    ];
    var rowRange = sheet.getRange(row, 1, 1, vals.length);
    rowRange.setValues([vals]);

    if (r.absentDays > 0) {
      sheet.getRange(row, 5).setFontColor('#DC2626').setFontWeight('bold');
    }
    if (r.reliability !== null && r.reliability < 75) {
      sheet.getRange(row, 7).setFontColor('#DC2626');
    } else if (r.reliability !== null && r.reliability < 90) {
      sheet.getRange(row, 7).setFontColor('#D97706');
    } else if (r.reliability !== null) {
      sheet.getRange(row, 7).setFontColor('#059669');
    }
    if (r.absTotal >= cfg.HIGH_VARIANCE_THRESHOLD) {
      sheet.getRange(row, 10).setFontColor('#DC2626').setFontWeight('bold');
    }
    if (r.midnightCount > 0) {
      sheet.getRange(row, 16).setBackground('#FEF3C7').setFontWeight('bold');
    }
    if (r.otHours > 0) {
      sheet.getRange(row, 17).setFontColor('#DC2626').setFontWeight('bold');
    }
    if (r.pattern.type === 'bad') {
      sheet.getRange(row, 18).setFontColor('#DC2626');
    } else if (r.pattern.type === 'good') {
      sheet.getRange(row, 18).setFontColor('#059669');
    } else if (r.pattern.type === 'warn') {
      sheet.getRange(row, 18).setFontColor('#D97706');
    }
    row++;
  });

  row += 2;

  // ── Section 5: Day-by-Day Detail ──
  sheet.getRange(row, 1).setValue('DAY-BY-DAY DETAIL').setFontSize(12).setFontWeight('bold');
  row++;

  allRows.forEach(function(emp) {
    sheet.getRange(row, 1).setValue(emp.name + '  (' + emp.locationName + ')');
    sheet.getRange(row, 1).setFontWeight('bold').setFontSize(10).setBackground('#E5E7EB');
    sheet.getRange(row, 1, 1, 9).setBackground('#E5E7EB');
    row++;

    var detailHeaders = ['Day', 'Status', 'Sched In', 'Sched Out', 'Actual In', 'Actual Out', 'Sched Hrs', 'Actual Hrs', 'Net Var'];
    sheet.getRange(row, 1, 1, detailHeaders.length).setValues([detailHeaders]);
    sheet.getRange(row, 1, 1, detailHeaders.length).setFontWeight('bold').setFontSize(8).setFontColor('#888888');
    row++;

    var sortedDays = emp.dayList.slice().sort(function(a, b) {
      return a.date < b.date ? -1 : a.date > b.date ? 1 : 0;
    });

    sortedDays.forEach(function(d) {
      var vals = [
        formatFriendlyDate(d.date),
        d.status.charAt(0).toUpperCase() + d.status.slice(1),
        d.schedStart ? formatTime12(d.schedStart) : '—',
        d.schedEnd ? formatTime12(d.schedEnd) : '—',
        d.actualStart ? formatTime12(d.actualStart) : '—',
        d.actualEnd ? formatTime12(d.actualEnd) + (d.midnightFlag ? ' ⚠ MISSED' : '') : '—',
        d.schedMinutes ? formatHours(d.schedMinutes) : '—',
        d.actualMinutes ? formatHours(d.actualMinutes) : '—',
        d.totalVar !== null ? formatVariance(d.totalVar) : '—'
      ];
      sheet.getRange(row, 1, 1, vals.length).setValues([vals]);

      if (d.status === 'absent') {
        sheet.getRange(row, 1, 1, vals.length).setBackground('#FEE2E2');
      } else if (d.status === 'unscheduled') {
        sheet.getRange(row, 1, 1, vals.length).setBackground('#DBEAFE');
      }
      if (d.midnightFlag) {
        sheet.getRange(row, 6).setFontColor('#D97706').setFontWeight('bold');
      }
      row++;
    });
    row++;
  });

  sheet.autoResizeColumns(1, headers.length);
}

/**
 * Look at the History tab's most recent previous run to get adherence
 * for week-over-week comparison.
 */
function getPreviousAdherence_(ss) {
  var historySheet = ss.getSheetByName(TABS.HISTORY);
  if (!historySheet || historySheet.getLastRow() < 2) return null;

  var data = historySheet.getDataRange().getValues();
  var runDates = {};
  for (var i = 1; i < data.length; i++) {
    var rd = String(data[i][0]);
    if (!runDates[rd]) runDates[rd] = [];
    runDates[rd].push(data[i]);
  }

  var sortedRuns = Object.keys(runDates).sort();
  if (sortedRuns.length < 1) return null;

  var prevRun = runDates[sortedRuns[sortedRuns.length - 1]];
  var totalSched = 0, totalMatched = 0;
  prevRun.forEach(function(row) {
    totalSched += (row[4] || 0);
    totalMatched += (row[5] || 0);
  });

  return totalSched > 0 ? Math.round(totalMatched / totalSched * 1000) / 10 : null;
}
