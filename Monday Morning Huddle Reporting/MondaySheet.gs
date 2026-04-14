/**
 * Schedule Variance Analyzer — Monday Morning Huddle Integration
 *
 * Writes the scheduling/OT block and attendance flags into the
 * "Current Week Schedule.ch" tab of THIS spreadsheet.
 *
 * Layout:
 *   Rows 1–31:  untouched (user content)
 *   Row  32:    LAST WEEK REVIEW banner
 *   Rows 33–37: last week hours table (sched, actual, actual OT, sched OT)
 *   Row  38:    spacer
 *   Row  39:    THIS WEEK PLAN banner
 *   Rows 40–42: this week hours table (sched hrs, sched OT)
 *   Row  43:    spacer
 *   Rows 44–63: attendance flags (OT/Late/Absent table, Missed Clock-Outs, Largest Variance)
 *   Rows 64–65: buffer
 */

var MONDAY_CLEAR_FROM = 32;
var MONDAY_START_ROW  = 32;
var MONDAY_END_ROW    = 65;
var MONDAY_COLS       = 5;

function fillMondaySheet() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cfg;

  try { cfg = getConfig(); } catch (e) {
    ui.alert('Configuration Error', e.message, ui.ButtonSet.OK); return;
  }

  var tabName = cfg.MONDAY_TAB || 'Current Week Schedule.ch';
  var targetTab = ss.getSheetByName(tabName);
  if (!targetTab) {
    ui.alert('Tab Not Found', 'No tab named "' + tabName + '".', ui.ButtonSet.OK); return;
  }

  var historySheet = ss.getSheetByName(TABS.HISTORY);
  if (!historySheet || historySheet.getLastRow() < 2) {
    ui.alert('No analysis data. Run "Upload & Analyze" first.'); return;
  }

  try {
    var runData = buildRunData_(ss, cfg);
    writeMondayBlock_(targetTab, runData, cfg);
    ss.setActiveSheet(targetTab);
    ui.alert('Monday Sheet Updated', '"' + tabName + '" updated.', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error: ' + e.message);
  }
}

/* ── Build run data: try input tabs first (fresh hours), fall back to History ── */

function buildRunData_(ss, cfg) {
  var loc1 = cfg.LOC1_NAME, loc2 = cfg.LOC2_NAME;
  var locTotals = {};
  locTotals[loc1] = { schedHrs: 0, actualHrs: 0 };
  locTotals[loc2] = { schedHrs: 0, actualHrs: 0 };

  var freshHours = getFreshHoursFromInputTabs_(ss, cfg);
  if (freshHours) {
    locTotals[loc1] = freshHours[loc1] || locTotals[loc1];
    locTotals[loc2] = freshHours[loc2] || locTotals[loc2];
  }

  // Last week's scheduled OT — re-derive from input tabs for accuracy
  var lastWeekSchedOT = getLastWeekScheduledOT_(ss, cfg);

  // This week's projected data from the "This Week" schedule tabs
  var thisWeek = getThisWeekScheduleData_(ss, cfg);

  // Pull employee-level flags from History
  var historySheet = ss.getSheetByName(TABS.HISTORY);
  var data = historySheet.getDataRange().getValues();
  var latestRunDate = String(data[data.length - 1][0]);
  var weekLabel = '', byEmployee = {};

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) !== latestRunDate) continue;
    var name = String(data[i][3] || '');
    var loc  = String(data[i][2] || '');
    weekLabel = String(data[i][1] || '') || weekLabel;

    if (!byEmployee[name]) {
      byEmployee[name] = { otHours: 0, lateCount: 0, midnightCount: 0, absentDays: 0, totalVar: 0, pattern: '', locs: {} };
    }
    var emp = byEmployee[name];
    emp.otHours = Math.max(emp.otHours, data[i][20] || 0);
    emp.lateCount += data[i][21] || 0;
    emp.midnightCount += data[i][17] || 0;
    emp.absentDays += data[i][6] || 0;
    emp.totalVar += data[i][11] || 0;
    emp.pattern = data[i][18] || emp.pattern;
    emp.locs[loc] = (emp.locs[loc] || 0) + (data[i][23] || 0);

    if (!freshHours) {
      var sh = data[i][22] || 0;
      var ah = data[i][23] || 0;
      locTotals[loc].schedHrs  += sh;
      locTotals[loc].actualHrs += ah;
    }
  }

  return {
    weekLabel: weekLabel,
    byEmployee: byEmployee,
    locTotals: locTotals,
    loc1: loc1,
    loc2: loc2,
    lastWeekSchedOT: lastWeekSchedOT,
    thisWeek: thisWeek
  };
}

/**
 * Re-read the input tabs and sum scheduled/actual minutes to get real hours.
 * Returns null if input tabs are empty.
 */
function getFreshHoursFromInputTabs_(ss, cfg) {
  var result = {};

  var pairs = [
    { schedTab: TABS.CH_SCHED, punchTab: TABS.CH_PUNCH, locName: cfg.LOC1_NAME },
    { schedTab: TABS.DBU_SCHED, punchTab: TABS.DBU_PUNCH, locName: cfg.LOC2_NAME }
  ];

  var anyData = false;
  pairs.forEach(function(p) {
    var sSheet = ss.getSheetByName(p.schedTab);
    var pSheet = ss.getSheetByName(p.punchTab);
    if (!sSheet || !pSheet || sSheet.getLastRow() < 2 || pSheet.getLastRow() < 2) {
      result[p.locName] = { schedHrs: 0, actualHrs: 0 };
      return;
    }
    anyData = true;
    try {
      var sched = parseScheduleFromSheet(sSheet.getDataRange().getValues());
      var punch = parsePunchesFromSheet(pSheet.getDataRange().getValues(), cfg);
      var schedMin = sched.reduce(function(t, s) { return t + s.minutes; }, 0);
      var actualMin = punch.reduce(function(t, p) { return t + p.minutes; }, 0);
      result[p.locName] = {
        schedHrs:  Math.round(schedMin / 60 * 100) / 100,
        actualHrs: Math.round(actualMin / 60 * 100) / 100
      };
    } catch (e) {
      result[p.locName] = { schedHrs: 0, actualHrs: 0 };
    }
  });

  return anyData ? result : null;
}

/**
 * Compute last week's scheduled OT by re-parsing input tabs and running
 * cross-location aggregation. Returns per-location and total scheduled OT,
 * or null if input tabs are empty.
 */
function getLastWeekScheduledOT_(ss, cfg) {
  var otThresholdMin = ((cfg && cfg.OT_THRESHOLD) || 40) * 60;
  var loc1 = cfg.LOC1_NAME, loc2 = cfg.LOC2_NAME;

  var tabPairs = [
    { tab: TABS.CH_SCHED, loc: loc1 },
    { tab: TABS.DBU_SCHED, loc: loc2 }
  ];

  var empMap = {};
  var anyData = false;

  tabPairs.forEach(function(p) {
    var sheet = ss.getSheetByName(p.tab);
    if (!sheet || sheet.getLastRow() < 2) return;
    try {
      var rows = parseScheduleFromSheet(sheet.getDataRange().getValues());
      anyData = true;
      rows.forEach(function(r) {
        if (!empMap[r.name]) empMap[r.name] = {};
        if (!empMap[r.name][p.loc]) empMap[r.name][p.loc] = 0;
        empMap[r.name][p.loc] += r.minutes;
      });
    } catch (e) { /* skip bad data */ }
  });

  if (!anyData) return null;

  var locOT = {};
  locOT[loc1] = 0;
  locOT[loc2] = 0;
  var totalOT = 0;

  Object.keys(empMap).forEach(function(name) {
    var locs = empMap[name];
    var totalMin = 0;
    Object.keys(locs).forEach(function(l) { totalMin += locs[l]; });
    var otMin = Math.max(0, totalMin - otThresholdMin);
    if (otMin <= 0) return;
    totalOT += otMin;

    // Distribute OT proportionally across locations
    Object.keys(locs).forEach(function(l) {
      var share = totalMin > 0 ? locs[l] / totalMin : 0;
      locOT[l] += otMin * share;
    });
  });

  return {
    loc1OT: Math.round(locOT[loc1] / 60 * 10) / 10,
    loc2OT: Math.round(locOT[loc2] / 60 * 10) / 10,
    totalOT: Math.round(totalOT / 60 * 10) / 10
  };
}

/**
 * Parse the "This Week" schedule tabs to get projected hours and OT.
 * Returns null if both tabs are empty.
 */
function getThisWeekScheduleData_(ss, cfg) {
  var otThresholdMin = ((cfg && cfg.OT_THRESHOLD) || 40) * 60;
  var loc1 = cfg.LOC1_NAME, loc2 = cfg.LOC2_NAME;

  var tabPairs = [
    { tab: TABS.CH_SCHED_NEXT, loc: loc1 },
    { tab: TABS.DBU_SCHED_NEXT, loc: loc2 }
  ];

  var empMap = {};
  var locHrs = {};
  locHrs[loc1] = 0;
  locHrs[loc2] = 0;
  var anyData = false;
  var weekLabel = '';

  tabPairs.forEach(function(p) {
    var sheet = ss.getSheetByName(p.tab);
    if (!sheet || sheet.getLastRow() < 2) return;
    try {
      var rows = parseScheduleFromSheet(sheet.getDataRange().getValues());
      anyData = true;

      // Derive week label from earliest date in this schedule
      if (!weekLabel && rows.length) {
        var dates = rows.map(function(r) { return r.date; }).sort();
        weekLabel = makeWeekLabel(getWeekStart(dates[0]));
      }

      rows.forEach(function(r) {
        locHrs[p.loc] += r.minutes;
        if (!empMap[r.name]) empMap[r.name] = {};
        if (!empMap[r.name][p.loc]) empMap[r.name][p.loc] = 0;
        empMap[r.name][p.loc] += r.minutes;
      });
    } catch (e) { /* skip bad data */ }
  });

  if (!anyData) return null;

  // Convert location totals to hours
  var loc1Hrs = Math.round(locHrs[loc1] / 60 * 10) / 10;
  var loc2Hrs = Math.round(locHrs[loc2] / 60 * 10) / 10;

  // Compute scheduled OT (cross-location per employee)
  var locOT = {};
  locOT[loc1] = 0;
  locOT[loc2] = 0;
  var totalOT = 0;

  Object.keys(empMap).forEach(function(name) {
    var locs = empMap[name];
    var totalMin = 0;
    Object.keys(locs).forEach(function(l) { totalMin += locs[l]; });
    var otMin = Math.max(0, totalMin - otThresholdMin);
    if (otMin <= 0) return;
    totalOT += otMin;

    Object.keys(locs).forEach(function(l) {
      var share = totalMin > 0 ? locs[l] / totalMin : 0;
      locOT[l] += otMin * share;
    });
  });

  return {
    weekLabel: weekLabel,
    loc1SchedHrs: loc1Hrs,
    loc2SchedHrs: loc2Hrs,
    totalSchedHrs: Math.round((loc1Hrs + loc2Hrs) * 10) / 10,
    loc1OT: Math.round(locOT[loc1] / 60 * 10) / 10,
    loc2OT: Math.round(locOT[loc2] / 60 * 10) / 10,
    totalOT: Math.round(totalOT / 60 * 10) / 10
  };
}

/* ── Shorten a name to "First L." for compact flag display ── */
function shortName_(fullName) {
  if (!fullName) return '';
  var s = fullName.trim();

  // Legacy "Last, First" format (pre-normalization or alias override)
  var comma = s.indexOf(',');
  if (comma > 0 && comma < s.length - 1) {
    var first = s.substring(comma + 1).trim().split(' ')[0];
    var last  = s.substring(0, comma).trim();
    return first + ' ' + last.charAt(0) + '.';
  }

  // Normalized "First [Middle] Last" format
  var words = s.split(' ');
  if (words.length >= 2) {
    return words[0] + ' ' + words[words.length - 1].charAt(0) + '.';
  }
  return s;
}

/* ── Visual layout writer ── */

function writeMondayBlock_(sheet, data, cfg) {
  // Clear the full zone (old remnants + new area)
  var clearRows = MONDAY_END_ROW - MONDAY_CLEAR_FROM + 1;
  var clearRange = sheet.getRange(MONDAY_CLEAR_FROM, 1, clearRows, MONDAY_COLS);
  clearRange.breakApart();
  clearRange.clearContent().clearFormat()
    .setFontFamily('Arial').setFontSize(10).setVerticalAlignment('middle')
    .setBackground(null).setFontColor('#000000').setFontWeight('normal')
    .setBorder(false, false, false, false, false, false);

  var r = MONDAY_START_ROW;
  var loc1 = data.loc1, loc2 = data.loc2;
  var t1 = data.locTotals[loc1] || { schedHrs: 0, actualHrs: 0 };
  var t2 = data.locTotals[loc2] || { schedHrs: 0, actualHrs: 0 };
  var rd = function(v) { return Math.round(v * 10) / 10; };
  var BORDER = SpreadsheetApp.BorderStyle.SOLID;

  // ── Compute actual OT per location ──
  var loc1OT = 0, loc2OT = 0;
  Object.keys(data.byEmployee).forEach(function(name) {
    var emp = data.byEmployee[name];
    if (emp.otHours <= 0) return;
    var locs = Object.keys(emp.locs);
    if (locs.length === 1) {
      if (locs[0] === loc1) loc1OT += emp.otHours; else loc2OT += emp.otHours;
    } else {
      var total = 0;
      locs.forEach(function(l) { total += emp.locs[l]; });
      locs.forEach(function(l) {
        var share = total > 0 ? emp.locs[l] / total : 0;
        if (l === loc1) loc1OT += emp.otHours * share; else loc2OT += emp.otHours * share;
      });
    }
  });

  // ═════════════════════════════════════════════════
  //  ROW 21: LAST WEEK BANNER
  // ═════════════════════════════════════════════════
  sheet.getRange(r, 1, 1, MONDAY_COLS).merge()
    .setValue('LAST WEEK REVIEW  ·  ' + data.weekLabel)
    .setFontSize(17).setFontWeight('bold').setFontColor('#FFFFFF')
    .setBackground('#1a1a2e').setHorizontalAlignment('center');
  r++;

  // ═════════════════════════════════════════════════
  //  ROWS 22–26: LAST WEEK HOURS TABLE
  // ═════════════════════════════════════════════════

  // Row 22: Column headers
  sheet.getRange(r, 1, 1, MONDAY_COLS)
    .setValues([[' ', loc1, loc2, 'Combined', 'Δ']])
    .setFontWeight('bold').setFontSize(12).setFontColor('#555555')
    .setBackground('#E5E7EB').setHorizontalAlignment('center');
  sheet.getRange(r, 1).setHorizontalAlignment('left');
  r++;

  // Row 23: Scheduled Hours
  var cS = rd(t1.schedHrs + t2.schedHrs);
  sheet.getRange(r, 1, 1, 4).setValues([['Scheduled', rd(t1.schedHrs), rd(t2.schedHrs), cS]]);
  styleHoursRow_(sheet, r, '#FFFFFF');
  r++;

  // Row 24: Actual Hours
  var cA = rd(t1.actualHrs + t2.actualHrs);
  var vPct = cS > 0 ? Math.round((cA - cS) / cS * 1000) / 10 : 0;
  sheet.getRange(r, 1, 1, 4).setValues([['Actual', rd(t1.actualHrs), rd(t2.actualHrs), cA]]);
  sheet.getRange(r, 5).setValue((vPct >= 0 ? '+' : '') + vPct + '%')
    .setFontSize(12).setFontWeight('bold').setFontColor(vPct > 0 ? '#DC2626' : '#059669')
    .setHorizontalAlignment('center');
  styleHoursRow_(sheet, r, '#F9FAFB');
  r++;

  // Row 25: Actual OT (amber highlight)
  var tOT = rd(loc1OT + loc2OT);
  sheet.getRange(r, 1, 1, 4).setValues([['Actual OT', rd(loc1OT), rd(loc2OT), tOT]]);
  sheet.getRange(r, 1).setFontWeight('bold').setFontSize(12).setFontColor('#92400E');
  sheet.getRange(r, 2, 1, 3).setFontWeight('bold').setFontSize(12).setFontColor('#92400E')
    .setHorizontalAlignment('center').setNumberFormat('#,##0.0');
  sheet.getRange(r, 4).setFontSize(16).setFontColor(tOT > 0 ? '#DC2626' : '#059669');
  sheet.getRange(r, 1, 1, MONDAY_COLS).setBackground('#FEF3C7')
    .setBorder(true, true, true, true, true, true, '#F59E0B', BORDER);
  r++;

  // Row 26: Scheduled OT (what was originally on the schedule)
  var sOT = data.lastWeekSchedOT;
  var sLoc1OT = sOT ? rd(sOT.loc1OT) : 0;
  var sLoc2OT = sOT ? rd(sOT.loc2OT) : 0;
  var sTotalOT = sOT ? rd(sOT.totalOT) : 0;
  sheet.getRange(r, 1, 1, 4).setValues([['Sched OT', sLoc1OT, sLoc2OT, sTotalOT]]);
  sheet.getRange(r, 1).setFontWeight('bold').setFontSize(12).setFontColor('#78350F');
  sheet.getRange(r, 2, 1, 3).setFontSize(12).setFontColor('#78350F')
    .setHorizontalAlignment('center').setNumberFormat('#,##0.0');
  var otDelta = rd(tOT - sTotalOT);
  var otDeltaStr = (otDelta >= 0 ? '+' : '') + otDelta + 'h';
  sheet.getRange(r, 5).setValue(otDeltaStr)
    .setFontSize(11).setFontWeight('bold')
    .setFontColor(otDelta > 0 ? '#DC2626' : '#059669')
    .setHorizontalAlignment('center');
  sheet.getRange(r, 1, 1, MONDAY_COLS).setBackground('#FFF7ED')
    .setBorder(true, true, true, true, true, true, '#F59E0B', BORDER);
  r++;

  // Format last-week hours table numbers
  sheet.getRange(MONDAY_START_ROW + 2, 2, 2, 3).setNumberFormat('#,##0.0').setHorizontalAlignment('center');

  // Row 27: spacer
  r++;

  // ═════════════════════════════════════════════════
  //  ROWS 28–31: THIS WEEK PLAN
  // ═════════════════════════════════════════════════
  var tw = data.thisWeek;
  var twLabel = tw ? tw.weekLabel : 'No data uploaded';

  // Row 28: THIS WEEK banner
  sheet.getRange(r, 1, 1, MONDAY_COLS).merge()
    .setValue('THIS WEEK PLAN  ·  ' + twLabel)
    .setFontSize(17).setFontWeight('bold').setFontColor('#FFFFFF')
    .setBackground('#0F766E').setHorizontalAlignment('center');
  r++;

  if (tw) {
    // Row 29: Column headers
    sheet.getRange(r, 1, 1, MONDAY_COLS)
      .setValues([[' ', loc1, loc2, 'Combined', ' ']])
      .setFontWeight('bold').setFontSize(12).setFontColor('#555555')
      .setBackground('#E5E7EB').setHorizontalAlignment('center');
    sheet.getRange(r, 1).setHorizontalAlignment('left');
    r++;

    // Row 30: Scheduled Hrs
    sheet.getRange(r, 1, 1, 4)
      .setValues([['Scheduled', rd(tw.loc1SchedHrs), rd(tw.loc2SchedHrs), rd(tw.totalSchedHrs)]]);
    styleHoursRow_(sheet, r, '#F0FDFA');
    r++;

    // Row 31: Scheduled OT
    sheet.getRange(r, 1, 1, 4)
      .setValues([['Sched OT', rd(tw.loc1OT), rd(tw.loc2OT), rd(tw.totalOT)]]);
    sheet.getRange(r, 1).setFontWeight('bold').setFontSize(12).setFontColor('#134E4A');
    sheet.getRange(r, 2, 1, 3).setFontWeight('bold').setFontSize(12).setFontColor('#134E4A')
      .setHorizontalAlignment('center').setNumberFormat('#,##0.0');
    sheet.getRange(r, 4).setFontSize(16).setFontColor(tw.totalOT > 0 ? '#DC2626' : '#059669');
    sheet.getRange(r, 1, 1, MONDAY_COLS).setBackground('#CCFBF1')
      .setBorder(true, true, true, true, true, true, '#14B8A6', BORDER);
    r++;
  } else {
    // No this-week data: show a note
    sheet.getRange(r, 1, 1, MONDAY_COLS).merge()
      .setValue('Upload this week\'s schedule CSVs to see projected hours & OT here.')
      .setFontSize(11).setFontColor('#6B7280').setHorizontalAlignment('center')
      .setBackground('#F3F4F6');
    r++;
  }

  // Row 32: spacer
  r++;

  // ═════════════════════════════════════════════════
  //  ATTENDANCE FLAGS (3-column table)
  // ═════════════════════════════════════════════════
  var flagsStartRow = r;

  // Collect flag data
  var otList = [], lateList = [], midList = [], absentList = [], varList = [];
  Object.keys(data.byEmployee).forEach(function(name) {
    var e = data.byEmployee[name];
    if (e.otHours > 0) otList.push({ name: name, val: rd(e.otHours) + 'h' });
    if (e.lateCount >= 2 || e.pattern === 'Late In' || e.pattern === 'Works Less')
      lateList.push({ name: name, val: e.lateCount + 'x' });
    if (e.midnightCount > 0) midList.push({ name: name, val: e.midnightCount + 'x' });
    if (e.absentDays > 0) absentList.push({ name: name, val: e.absentDays + 'd' });
    if (Math.abs(e.totalVar) >= 240) {
      var hrs = rd(Math.abs(e.totalVar) / 60);
      varList.push({ name: name, absVar: Math.abs(e.totalVar), val: (e.totalVar >= 0 ? '+' : '-') + hrs + 'h' });
    }
  });
  otList.sort(function(a, b) { return parseFloat(b.val) - parseFloat(a.val); });
  lateList.sort(function(a, b) { return parseInt(b.val) - parseInt(a.val); });
  varList.sort(function(a, b) { return b.absVar - a.absVar; });
  if (varList.length > 10) varList = varList.slice(0, 10);

  // FLAGS banner
  sheet.getRange(r, 1, 1, MONDAY_COLS).merge()
    .setValue('ATTENDANCE FLAGS')
    .setFontSize(17).setFontWeight('bold').setFontColor('#FFFFFF')
    .setBackground('#374151').setHorizontalAlignment('center');
  r++;

  // 3-column headers
  sheet.getRange(r, 1, 1, 2).merge()
    .setValue('OT OFFENDERS').setFontSize(12).setFontWeight('bold')
    .setFontColor('#FFFFFF').setBackground('#DC2626').setHorizontalAlignment('center');
  sheet.getRange(r, 3)
    .setValue('LATE ARRIVALS').setFontSize(12).setFontWeight('bold')
    .setFontColor('#FFFFFF').setBackground('#D97706').setHorizontalAlignment('center');
  sheet.getRange(r, 4, 1, 2).merge()
    .setValue('ABSENCES').setFontSize(12).setFontWeight('bold')
    .setFontColor('#FFFFFF').setBackground('#1D4ED8').setHorizontalAlignment('center');
  r++;

  var maxNames = 5;
  for (var n = 0; n < maxNames; n++) {
    // OT column (A-B merged)
    sheet.getRange(r, 1, 1, 2).merge();
    if (n < otList.length) {
      sheet.getRange(r, 1).setValue(shortName_(otList[n].name) + '  ' + otList[n].val)
        .setFontSize(11).setFontColor('#991B1B');
    } else if (n === 0) {
      sheet.getRange(r, 1).setValue('None').setFontSize(11).setFontColor('#059669');
    }
    sheet.getRange(r, 1, 1, 2).setBackground('#FEF2F2');

    // Late column (C)
    if (n < lateList.length) {
      sheet.getRange(r, 3).setValue(shortName_(lateList[n].name) + '  ' + lateList[n].val)
        .setFontSize(11).setFontColor('#92400E');
    } else if (n === 0) {
      sheet.getRange(r, 3).setValue('None').setFontSize(11).setFontColor('#059669');
    }
    sheet.getRange(r, 3).setBackground('#FFFBEB');

    // Absent column (D-E merged)
    sheet.getRange(r, 4, 1, 2).merge();
    if (n < absentList.length) {
      sheet.getRange(r, 4).setValue(shortName_(absentList[n].name) + '  ' + absentList[n].val)
        .setFontSize(11).setFontColor('#1E40AF');
    } else if (n === 0) {
      sheet.getRange(r, 4).setValue('None').setFontSize(11).setFontColor('#059669');
    }
    sheet.getRange(r, 4, 1, 2).setBackground('#EFF6FF');

    r++;
  }

  // ── Missed Clock-Outs (table layout, 2 columns, up to 10 people) ──
  midList.sort(function(a, b) { return parseInt(b.val) - parseInt(a.val); });
  if (midList.length > 10) midList = midList.slice(0, 10);

  sheet.getRange(r, 1, 1, MONDAY_COLS).merge()
    .setValue('MISSED CLOCK-OUTS').setFontSize(12).setFontWeight('bold')
    .setFontColor('#FFFFFF').setBackground('#7C3AED').setHorizontalAlignment('center');
  r++;

  var midRows = midList.length > 0 ? Math.min(5, Math.ceil(midList.length / 2)) : 1;
  for (var mi = 0; mi < midRows; mi++) {
    var mL = mi * 2, mR = mi * 2 + 1;

    sheet.getRange(r, 1, 1, 3).merge();
    if (mL < midList.length) {
      sheet.getRange(r, 1).setValue(shortName_(midList[mL].name) + '  ' + midList[mL].val)
        .setFontSize(11).setFontColor('#6B21A8');
    } else if (mL === 0) {
      sheet.getRange(r, 1).setValue('None').setFontSize(11).setFontColor('#059669');
    }
    sheet.getRange(r, 1, 1, 3).setBackground('#F5F3FF');

    sheet.getRange(r, 4, 1, 2).merge();
    if (mR < midList.length) {
      sheet.getRange(r, 4).setValue(shortName_(midList[mR].name) + '  ' + midList[mR].val)
        .setFontSize(11).setFontColor('#6B21A8');
    }
    sheet.getRange(r, 4, 1, 2).setBackground('#F5F3FF');
    r++;
  }

  // ── Largest Variance (table layout, 2 columns, up to 10 people, ±4hr threshold) ──
  sheet.getRange(r, 1, 1, MONDAY_COLS).merge()
    .setValue('LARGEST VARIANCE').setFontSize(12).setFontWeight('bold')
    .setFontColor('#FFFFFF').setBackground('#C2410C').setHorizontalAlignment('center');
  r++;

  var varRows = varList.length > 0 ? Math.min(5, Math.ceil(varList.length / 2)) : 1;
  for (var vi = 0; vi < varRows; vi++) {
    var vL = vi * 2, vR = vi * 2 + 1;

    sheet.getRange(r, 1, 1, 3).merge();
    if (vL < varList.length) {
      sheet.getRange(r, 1).setValue(shortName_(varList[vL].name) + '  ' + varList[vL].val)
        .setFontSize(11).setFontColor('#9A3412');
    } else if (vL === 0) {
      sheet.getRange(r, 1).setValue('None').setFontSize(11).setFontColor('#059669');
    }
    sheet.getRange(r, 1, 1, 3).setBackground('#FFF7ED');

    sheet.getRange(r, 4, 1, 2).merge();
    if (vR < varList.length) {
      sheet.getRange(r, 4).setValue(shortName_(varList[vR].name) + '  ' + varList[vR].val)
        .setFontSize(11).setFontColor('#9A3412');
    }
    sheet.getRange(r, 4, 1, 2).setBackground('#FFF7ED');
    r++;
  }

  // Outer borders around the flags section
  var flagRows = r - flagsStartRow;
  sheet.getRange(flagsStartRow, 1, flagRows, MONDAY_COLS)
    .setBorder(true, true, true, true, false, false, '#D1D5DB', BORDER);
}

function styleHoursRow_(sheet, r, bg) {
  sheet.getRange(r, 1).setFontWeight('bold').setFontSize(12).setFontColor('#333333');
  sheet.getRange(r, 2, 1, 3).setFontSize(12).setHorizontalAlignment('center').setFontColor('#333333').setNumberFormat('#,##0.0');
  sheet.getRange(r, 1, 1, MONDAY_COLS).setBackground(bg)
    .setBorder(true, true, true, true, true, true, '#D1D5DB', SpreadsheetApp.BorderStyle.SOLID);
}
