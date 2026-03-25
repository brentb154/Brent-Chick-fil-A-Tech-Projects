/**
 * Schedule Variance Analyzer — Analysis Engine
 *
 * Consolidation, variance calc, midnight detection, pattern detection,
 * swap detection, roll-up, cross-location OT reconciliation.
 */

/* ── Consolidation: match schedule to punches by employee+date ── */

function consolidate(schedule, punches, cfg) {
  var sMap = {}, pMap = {};

  schedule.forEach(function(s) {
    var k = s.name + '|||' + s.date;
    if (!sMap[k]) sMap[k] = { name: s.name, date: s.date, segs: [] };
    sMap[k].segs.push(s);
  });

  punches.forEach(function(p) {
    var k = p.name + '|||' + p.date;
    if (!pMap[k]) pMap[k] = { name: p.name, date: p.date, pnch: [], rem: [] };
    pMap[k].pnch.push(p);
    if (p.remarks) pMap[k].rem.push(p.remarks);
  });

  var allKeys = {};
  Object.keys(sMap).forEach(function(k) { allKeys[k] = 1; });
  Object.keys(pMap).forEach(function(k) { allKeys[k] = 1; });

  var result = [];
  Object.keys(allKeys).forEach(function(key) {
    var s = sMap[key], p = pMap[key], pts = key.split('|||');
    var rec = {
      name: pts[0], date: pts[1], status: 'matched',
      locationName: '',
      schedStart: null, schedEnd: null, schedMinutes: 0,
      actualStart: null, actualEnd: null, actualMinutes: 0,
      startVar: null, endVar: null, totalVar: null,
      midnightFlag: false, remarks: ''
    };

    if (s) {
      rec.schedStart = s.segs.reduce(function(m, x) { return !m || x.start < m ? x.start : m; }, null);
      rec.schedEnd   = s.segs.reduce(function(m, x) { return !m || x.end > m ? x.end : m; }, null);
      rec.schedMinutes = s.segs.reduce(function(t, x) { return t + x.minutes; }, 0);
    }

    if (p) {
      rec.actualStart = p.pnch.reduce(function(m, x) { return !m || x.timeIn < m ? x.timeIn : m; }, null);
      rec.actualEnd   = p.pnch.reduce(function(m, x) { return !m || x.timeOut > m ? x.timeOut : m; }, null);
      rec.actualMinutes = p.pnch.reduce(function(t, x) { return t + x.minutes; }, 0);
      rec.remarks = p.rem.join('; ');
    }

    if (s && p) {
      rec.status = 'matched';
      rec.startVar = Math.round((rec.schedStart - rec.actualStart) / 60000);
      rec.endVar   = Math.round((rec.actualEnd - rec.schedEnd) / 60000);
      rec.totalVar = rec.actualMinutes - rec.schedMinutes;
    } else if (s && !p) {
      rec.status = 'absent';
    } else {
      rec.status = 'unscheduled';
    }

    rec.midnightFlag = rec.actualEnd ? isMidnight(rec.actualEnd, cfg) : false;
    result.push(rec);
  });

  result.sort(function(a, b) {
    return a.name < b.name ? -1 : a.name > b.name ? 1 : a.date < b.date ? -1 : a.date > b.date ? 1 : 0;
  });
  return result;
}

/* ── Midnight Detection ── */

function isMidnight(dt, cfg) {
  if (!dt || isNaN(dt.getTime())) return false;
  var h = dt.getHours(), m = dt.getMinutes();
  var window = (cfg && cfg.MIDNIGHT_WINDOW) || 5;
  return (h === 0 && m <= window) || (h === 23 && m >= (60 - window));
}

/* ── Week Detection ── */

function detectWeeks(records) {
  var wMap = {};
  records.forEach(function(r) {
    var ws = getWeekStart(r.date);
    if (!wMap[ws]) wMap[ws] = [];
    wMap[ws].push(r);
  });

  return Object.keys(wMap).sort().map(function(ws) {
    return { label: makeWeekLabel(ws), weekStart: ws, records: wMap[ws] };
  });
}

function getWeekStart(dateStr) {
  var p = dateStr.split('-');
  var d = new Date(parseInt(p[0]), parseInt(p[1]) - 1, parseInt(p[2]));
  var day = d.getDay();
  var diff = day === 0 ? 6 : day - 1;
  d.setDate(d.getDate() - diff);
  return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
}

function makeWeekLabel(wsStr) {
  var p = wsStr.split('-');
  var mon = new Date(parseInt(p[0]), parseInt(p[1]) - 1, parseInt(p[2]));
  var sun = new Date(mon);
  sun.setDate(sun.getDate() + 6);
  return 'Week of ' + (mon.getMonth() + 1) + '/' + mon.getDate() + ' – ' + (sun.getMonth() + 1) + '/' + sun.getDate();
}

/* ── Pattern Detection ── */

function detectPattern(matchedDays, cfg) {
  if (!matchedDays.length) return { label: '—', type: 'none' };
  var n = matchedDays.length;
  var lateIn = 0, earlyIn = 0, staysLate = 0, leavesEarly = 0;
  var thresh = (cfg && cfg.LATENESS_THRESHOLD) || 10;

  matchedDays.forEach(function(d) {
    if (d.startVar <= -thresh) lateIn++;
    if (d.startVar >= thresh)  earlyIn++;
    if (d.endVar >= thresh)    staysLate++;
    if (d.endVar <= -thresh)   leavesEarly++;
  });

  var pct = (cfg && cfg.PATTERN_PERCENT) || 0.4;
  var li = lateIn / n >= pct, ei = earlyIn / n >= pct;
  var sl = staysLate / n >= pct, le = leavesEarly / n >= pct;

  if (ei && sl) return { label: 'Works Extra', type: 'warn' };
  if (li && le) return { label: 'Works Less',  type: 'bad' };
  if (ei && le) return { label: 'Shift Early', type: 'neutral' };
  if (li && sl) return { label: 'Shift Late',  type: 'warn' };
  if (li)       return { label: 'Late In',     type: 'bad' };
  if (ei)       return { label: 'Early In',    type: 'neutral' };
  if (sl)       return { label: 'Stays Late',  type: 'warn' };
  if (le)       return { label: 'Leaves Early',type: 'bad' };
  return { label: 'On Track', type: 'good' };
}

/* ── Swap Detection ── */

function isSwap(d, cfg) {
  var thresh = (cfg && cfg.SWAP_THRESHOLD) || 120;
  return d.startVar !== null && Math.abs(d.startVar) >= thresh;
}

/* ── Roll-Up: day records → per-employee summary ── */

function rollUp(dayRecords, cfg) {
  var g = {};
  var otThresholdMin = ((cfg && cfg.OT_THRESHOLD) || 40) * 60;

  dayRecords.forEach(function(r) {
    var k = r.name;
    if (!g[k]) {
      g[k] = {
        name: r.name, locationName: r.locationName || '',
        dayList: [], matched: 0, scheduled: 0, worked: 0,
        absent: 0, unscheduled: 0, midnightCount: 0,
        totalSV: 0, totalEV: 0, totalTV: 0,
        totalSchedMin: 0, totalActualMin: 0, lateCount: 0
      };
    }
    var o = g[k];
    o.dayList.push(r);
    if (r.midnightFlag) o.midnightCount++;

    if (r.status === 'matched') {
      o.matched++; o.scheduled++; o.worked++;
      o.totalSV += r.startVar; o.totalEV += r.endVar; o.totalTV += r.totalVar;
      o.totalSchedMin += r.schedMinutes;
      o.totalActualMin += r.actualMinutes;
      if (r.startVar <= -(cfg && cfg.LATENESS_THRESHOLD || 10)) o.lateCount++;
    } else if (r.status === 'absent') {
      o.absent++; o.scheduled++;
      o.totalSchedMin += r.schedMinutes;
    } else {
      o.unscheduled++; o.worked++;
      o.totalActualMin += r.actualMinutes;
    }
  });

  return Object.keys(g).map(function(k) {
    var o = g[k];
    var md = o.dayList.filter(function(d) { return d.status === 'matched'; });
    var pat = detectPattern(md, cfg);
    var swaps = md.filter(function(d) { return isSwap(d, cfg); }).length;

    var earlyIn = 0, lateIn = 0, lateOut = 0, earlyOut = 0;
    md.forEach(function(d) {
      if (d.startVar > 0) earlyIn += d.startVar; else if (d.startVar < 0) lateIn += Math.abs(d.startVar);
      if (d.endVar > 0) lateOut += d.endVar; else if (d.endVar < 0) earlyOut += Math.abs(d.endVar);
    });

    var reliability = o.scheduled > 0 ? Math.round(o.matched / o.scheduled * 1000) / 10 : null;
    var scheduledHours = Math.round(o.totalSchedMin / 60 * 10) / 10;
    var actualHours    = Math.round(o.totalActualMin / 60 * 10) / 10;
    var otHours        = Math.max(0, Math.round((o.totalActualMin - otThresholdMin) / 60 * 10) / 10);

    return {
      name: o.name,
      locationName: o.locationName,
      weekLabel: '',
      totalDays: o.dayList.length,
      matchedDays: o.matched,
      scheduledDays: o.scheduled,
      workedDays: o.worked,
      absentDays: o.absent,
      unscheduledDays: o.unscheduled,
      startVar: o.totalSV,
      endVar: o.totalEV,
      totalVar: o.totalTV,
      avgVar: o.matched ? Math.round(o.totalTV / o.matched * 10) / 10 : 0,
      absTotal: Math.abs(o.totalTV),
      reliability: reliability,
      earlyIn: earlyIn,
      lateIn: lateIn,
      lateOut: lateOut,
      earlyOut: earlyOut,
      midnightCount: o.midnightCount,
      pattern: pat,
      swapCount: swaps,
      scheduledHours: scheduledHours,
      actualHours: actualHours,
      otHours: otHours,
      lateCount: o.lateCount,
      dayList: o.dayList
    };
  }).sort(function(a, b) { return b.absTotal - a.absTotal; });
}

/* ── Cross-Location OT Reconciliation ── */

/**
 * Aggregate actual hours per employee across ALL locations for the week,
 * then determine who is truly OT when combining both locations.
 *
 * Returns: {
 *   employees: { "Name": { loc1Hours, loc2Hours, totalHours, otHours, scheduledTotal } },
 *   byLocation: { "Cockrell Hill DTO": { schedHours, actualHours, otHours }, ... },
 *   totalOTHours: number,
 *   totalScheduledOTHours: number
 * }
 */
function reconcileCrossLocationOT(allConsolidated, cfg) {
  var otThresholdMin = ((cfg && cfg.OT_THRESHOLD) || 40) * 60;
  var empMap = {};
  var locMap = {};

  allConsolidated.forEach(function(r) {
    var loc = r.locationName || 'Unknown';

    if (!empMap[r.name]) empMap[r.name] = {};
    if (!empMap[r.name][loc]) empMap[r.name][loc] = { schedMin: 0, actualMin: 0 };

    if (r.status === 'matched' || r.status === 'absent') {
      empMap[r.name][loc].schedMin += r.schedMinutes;
    }
    if (r.status === 'matched' || r.status === 'unscheduled') {
      empMap[r.name][loc].actualMin += r.actualMinutes;
    }

    if (!locMap[loc]) locMap[loc] = { schedMin: 0, actualMin: 0, otMin: 0, schedOtMin: 0 };
    if (r.status === 'matched' || r.status === 'absent') {
      locMap[loc].schedMin += r.schedMinutes;
    }
    if (r.status === 'matched' || r.status === 'unscheduled') {
      locMap[loc].actualMin += r.actualMinutes;
    }
  });

  var employees = {};
  var totalOTMin = 0;
  var totalScheduledOTMin = 0;

  Object.keys(empMap).forEach(function(name) {
    var locs = empMap[name];
    var totalActual = 0, totalSched = 0;
    var perLoc = {};

    Object.keys(locs).forEach(function(loc) {
      totalActual += locs[loc].actualMin;
      totalSched  += locs[loc].schedMin;
      perLoc[loc] = Math.round(locs[loc].actualMin / 60 * 10) / 10;
    });

    var otMin = Math.max(0, totalActual - otThresholdMin);
    var schedOtMin = Math.max(0, totalSched - otThresholdMin);
    totalOTMin += otMin;
    totalScheduledOTMin += schedOtMin;

    employees[name] = {
      perLocation: perLoc,
      totalHours: Math.round(totalActual / 60 * 10) / 10,
      scheduledTotal: Math.round(totalSched / 60 * 10) / 10,
      otHours: Math.round(otMin / 60 * 10) / 10,
      scheduledOtHours: Math.round(schedOtMin / 60 * 10) / 10,
      isMultiLocation: Object.keys(locs).length > 1
    };
  });

  var byLocation = {};
  Object.keys(locMap).forEach(function(loc) {
    byLocation[loc] = {
      schedHours:  Math.round(locMap[loc].schedMin / 60 * 10) / 10,
      actualHours: Math.round(locMap[loc].actualMin / 60 * 10) / 10
    };
  });

  return {
    employees: employees,
    byLocation: byLocation,
    totalOTHours: Math.round(totalOTMin / 60 * 10) / 10,
    totalScheduledOTHours: Math.round(totalScheduledOTMin / 60 * 10) / 10
  };
}
