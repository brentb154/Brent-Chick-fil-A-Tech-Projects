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
  // Reset every cell to Sheets' default ("General") format. clearFormats() alone
  // doesn't always strip column-level time/date formats inherited from prior runs,
  // which is why integer day counts were rendering as "12:00 AM".
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).setNumberFormat('General');

  // The whole report is buffered into parallel 2D arrays and written with one
  // setValues + one call per format layer at the end (was ~300 per-cell calls).
  var W = 19;
  var vals = [], colors = [], weights = [], bgs = [], sizes = [];
  var numFmtBlocks = []; // {row, col, nrows, ncols, fmt} applied after the batch write

  // cells: sparse array of values by 0-based col; fmts: {colIndex: {color,weight,bg,size}}
  function addRow(cells, fmts) {
    var v = [], c = [], w = [], b = [], s = [];
    for (var i = 0; i < W; i++) {
      v.push(cells && cells[i] !== undefined ? cells[i] : '');
      c.push('#000000'); w.push('normal'); b.push('#ffffff'); s.push(10);
    }
    if (fmts) {
      Object.keys(fmts).forEach(function(k) {
        var f = fmts[k], i = Number(k);
        if (f.color) c[i] = f.color;
        if (f.weight) w[i] = f.weight;
        if (f.bg) b[i] = f.bg;
        if (f.size) s[i] = f.size;
      });
    }
    vals.push(v); colors.push(c); weights.push(w); bgs.push(b); sizes.push(s);
    return vals.length; // 1-based sheet row just added
  }
  function addBlank(n) { for (var i = 0; i < (n || 1); i++) addRow([]); }
  // bg across the first n columns of a row
  function rowBg(bg, n, extra) {
    var f = extra || {};
    for (var i = 0; i < n; i++) f[i] = f[i] ? f[i] : {};
    for (var j = 0; j < n; j++) f[j].bg = bg;
    return f;
  }

  var locNames = locationData.map(function(l) { return l.name; }).join(' + ');
  var weekLabels = weeks.map(function(w) { return w.label; }).join(', ');

  // ── Header ──
  addRow(['SCHEDULE VARIANCE REPORT'], { 0: { size: 16, weight: 'bold' } });
  addRow([locNames + ' — ' + weekLabels], { 0: { size: 11, color: '#666666' } });
  addRow(['Generated: ' + new Date().toLocaleString()], { 0: { size: 9, color: '#999999' } });
  addBlank();

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

  addRow(['SUMMARY'], { 0: { size: 12, weight: 'bold' } });

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
    var f = { 0: { weight: 'bold', color: '#555555' } };
    if (s[0] === 'Schedule Adherence' && adhDelta !== null) {
      f[1] = { color: adhDelta >= 0 ? '#059669' : '#DC2626' };
    }
    addRow([s[0], s[1]], f);
  });
  addBlank();

  // ── Section 2: Flags & Alerts ──
  addRow(['FLAGS & ALERTS'], { 0: { size: 12, weight: 'bold' } });

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
    addRow(['Missed Clock-Outs (punched out near midnight)'], { 0: { weight: 'bold', color: '#D97706' } });
    midnighters.forEach(function(m) {
      addRow([m[0], m[1], m[2], m[3] + 'x'], rowBg('#FEF3C7', 4));
    });
    addBlank();
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
    addRow(['Absences (scheduled but no punch)'], { 0: { weight: 'bold', color: '#DC2626' } });
    absents.forEach(function(a) {
      addRow([a[0], a[1], a[2], a[3] + ' day(s)'], rowBg('#FEE2E2', 4));
    });
    addBlank();
  }

  // High variance
  var highVarEmps = allRows.filter(function(r) { return r.absTotal >= cfg.HIGH_VARIANCE_THRESHOLD; })
    .sort(function(a, b) { return b.absTotal - a.absTotal; });
  if (highVarEmps.length) {
    addRow(['High Variance (>' + cfg.HIGH_VARIANCE_THRESHOLD + ' min total)'], { 0: { weight: 'bold', color: '#DC2626' } });
    highVarEmps.forEach(function(r) {
      addRow([r.name, r.locationName, formatVariance(r.totalVar)], { 2: { color: '#DC2626' } });
    });
    addBlank();
  }

  // Possible swaps
  var swapEmps = allRows.filter(function(r) { return r.swapCount > 0; });
  if (swapEmps.length) {
    addRow(['Possible Schedule Mismatches / Swaps'], { 0: { weight: 'bold', color: '#D97706' } });
    swapEmps.forEach(function(r) {
      addRow([r.name, r.locationName, r.swapCount + ' shift(s) with >' + cfg.SWAP_THRESHOLD + ' min start variance'],
        rowBg('#FFF7ED', 3));
    });
    addBlank();
  }

  // Behavioral flags
  var latePatterns = allRows.filter(function(r) { return r.pattern.type === 'bad'; });
  if (latePatterns.length) {
    addRow(['Behavioral Flags'], { 0: { weight: 'bold', color: '#DC2626' } });
    latePatterns.forEach(function(r) {
      addRow([r.name, r.locationName, r.pattern.label,
        'Reliability: ' + (r.reliability !== null ? r.reliability + '%' : 'N/A')]);
    });
    addBlank();
  }

  // Chronic flags (from History lookback)
  if (chronicFlags && chronicFlags.length) {
    addRow(['ONGOING CONCERNS (flagged ' + cfg.CHRONIC_TRIGGER + '+ of last ' + cfg.CHRONIC_WINDOW + ' weeks)'],
      { 0: { weight: 'bold', color: '#7C3AED' } });
    chronicFlags.forEach(function(f) {
      addRow([f.name, f.reason, f.weeksTriggered + ' of last ' + cfg.CHRONIC_WINDOW + ' weeks'], rowBg('#F3E8FF', 3));
    });
    addBlank();
  }

  addBlank();

  // ── Section 3: OT Summary ──
  addRow(['OVERTIME SUMMARY'], { 0: { size: 12, weight: 'bold' } });

  Object.keys(crossLocOT.byLocation).forEach(function(loc) {
    var d = crossLocOT.byLocation[loc];
    addRow([loc, 'Scheduled: ' + d.schedHours + ' hrs', 'Actual: ' + d.actualHours + ' hrs'],
      { 0: { weight: 'bold' } });
  });

  addRow(['Cross-Location Reconciled OT', crossLocOT.totalOTHours + ' hrs actual OT',
    crossLocOT.totalScheduledOTHours + ' hrs scheduled OT'],
    { 0: { weight: 'bold', color: '#7C3AED' } });

  // Top OT offenders
  var otList = [];
  Object.keys(crossLocOT.employees).forEach(function(name) {
    var e = crossLocOT.employees[name];
    if (e.otHours > 0) otList.push({ name: name, ot: e.otHours, total: e.totalHours, multi: e.isMultiLocation });
  });
  otList.sort(function(a, b) { return b.ot - a.ot; });

  if (otList.length) {
    addBlank();
    addRow(['Top OT Employees'], { 0: { weight: 'bold' } });
    otList.slice(0, 10).forEach(function(e) {
      addRow([e.name + (e.multi ? ' ★ multi-location' : ''), e.ot + ' hrs OT', e.total + ' hrs total'],
        e.multi ? { 0: { color: '#7C3AED' } } : null);
    });
  }
  addBlank(2);

  // ── Section 4: Employee Roll-Up Table ──
  addRow(['EMPLOYEE ROLL-UP'], { 0: { size: 12, weight: 'bold' } });

  var headers = [
    'Employee', 'Location', 'Sched Days', 'Worked Days', 'Absent', 'Unsched',
    'Reliability', 'Start Var', 'End Var', 'Total Var', 'Avg/Day',
    'Early In', 'Late In', 'Stayed Late', 'Left Early',
    'Missed Punches', 'OT Hrs', 'Pattern', 'Flags'
  ];
  addRow(headers, rowBg('#F3F4F6', headers.length, (function() {
    var f = {};
    for (var i = 0; i < headers.length; i++) f[i] = { weight: 'bold', size: 9 };
    return f;
  })()));

  var rollupStart = vals.length + 1;
  allRows.forEach(function(r) {
    var f = {};
    if (r.absentDays > 0) f[4] = { color: '#DC2626', weight: 'bold' };
    if (r.reliability !== null && r.reliability < 75) f[6] = { color: '#DC2626' };
    else if (r.reliability !== null && r.reliability < 90) f[6] = { color: '#D97706' };
    else if (r.reliability !== null) f[6] = { color: '#059669' };
    if (r.absTotal >= cfg.HIGH_VARIANCE_THRESHOLD) f[9] = { color: '#DC2626', weight: 'bold' };
    if (r.midnightCount > 0) f[15] = { bg: '#FEF3C7', weight: 'bold' };
    if (r.otHours > 0) f[16] = { color: '#DC2626', weight: 'bold' };
    if (r.pattern.type === 'bad') f[17] = { color: '#DC2626' };
    else if (r.pattern.type === 'good') f[17] = { color: '#059669' };
    else if (r.pattern.type === 'warn') f[17] = { color: '#D97706' };

    addRow([
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
    ], f);
  });
  if (allRows.length) {
    // Force integer display on day-count columns — Sheets sometimes inherits
    // a time format on these cells from prior runs, displaying 5 as "12:00 AM".
    numFmtBlocks.push({ row: rollupStart, col: 3, nrows: allRows.length, ncols: 4, fmt: '0' });
  }

  addBlank(2);

  // ── Section 5: Day-by-Day Detail ──
  addRow(['DAY-BY-DAY DETAIL'], { 0: { size: 12, weight: 'bold' } });

  allRows.forEach(function(emp) {
    addRow([emp.name + '  (' + emp.locationName + ')'],
      rowBg('#E5E7EB', 9, { 0: { weight: 'bold', size: 10 } }));

    var detailHeaders = ['Day', 'Status', 'Sched In', 'Sched Out', 'Actual In', 'Actual Out', 'Sched Hrs', 'Actual Hrs', 'Net Var'];
    addRow(detailHeaders, (function() {
      var f = {};
      for (var i = 0; i < detailHeaders.length; i++) f[i] = { weight: 'bold', size: 8, color: '#888888' };
      return f;
    })());

    var sortedDays = emp.dayList.slice().sort(function(a, b) {
      return a.date < b.date ? -1 : a.date > b.date ? 1 : 0;
    });

    sortedDays.forEach(function(d) {
      var f = {};
      if (d.status === 'absent') f = rowBg('#FEE2E2', 9);
      else if (d.status === 'unscheduled') f = rowBg('#DBEAFE', 9);
      if (d.midnightFlag) {
        f[5] = f[5] || {};
        f[5].color = '#D97706';
        f[5].weight = 'bold';
      }
      addRow([
        formatFriendlyDate(d.date),
        d.status.charAt(0).toUpperCase() + d.status.slice(1),
        d.schedStart ? formatTime12(d.schedStart) : '—',
        d.schedEnd ? formatTime12(d.schedEnd) : '—',
        d.actualStart ? formatTime12(d.actualStart) : '—',
        d.actualEnd ? formatTime12(d.actualEnd) + (d.midnightFlag ? ' ⚠ MISSED' : '') : '—',
        d.schedMinutes ? formatHours(d.schedMinutes) : '—',
        d.actualMinutes ? formatHours(d.actualMinutes) : '—',
        d.totalVar !== null ? formatVariance(d.totalVar) : '—'
      ], f);
    });
    addBlank();
  });

  // ── Single batched write ──
  if (sheet.getMaxRows() < vals.length) {
    sheet.insertRowsAfter(sheet.getMaxRows(), vals.length - sheet.getMaxRows());
  }
  var out = sheet.getRange(1, 1, vals.length, W);
  out.setValues(vals);
  out.setBackgrounds(bgs);
  out.setFontColors(colors);
  out.setFontWeights(weights);
  out.setFontSizes(sizes);
  numFmtBlocks.forEach(function(nf) {
    sheet.getRange(nf.row, nf.col, nf.nrows, nf.ncols).setNumberFormat(nf.fmt);
  });
  SpreadsheetApp.flush();

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
