/**
 * ============================================================
 * TRAINING TRACKING SYSTEM - Timeline & Schedule Management
 * ============================================================
 * Generates training timelines, stores planned assignments in
 * a Training Schedule data sheet, and auto-populates the
 * Training Needs sheet weekly.
 *
 * KEY ARCHITECTURE:
 *   Timeline Dialog -> createTimeline() -> Training Schedule sheet
 *   Monday trigger  -> populateWeeklyTraining() -> Training Needs sheet
 *   Manual menu     -> loadThisWeeksTraining()  -> Training Needs sheet
 */


// ==============================================================
// DIALOG LAUNCHER
// ==============================================================

function generateTrainingTimeline() {
  var html = HtmlService.createHtmlOutputFromFile('TimelineDialog')
    .setWidth(580)
    .setHeight(680);
  SpreadsheetApp.getUi().showModalDialog(html, 'Generate Training Timeline');
}


// ==============================================================
// TIMELINE BUILDER (called from TimelineDialog.html)
// ==============================================================

/**
 * @param {Object} params
 *   name, house, hoursPerShift, shift,
 *   selectedDays: [[bool x 7] x 3] (3 weeks x 7 days, Mon=0)
 *   weekStartDate: "YYYY-MM-DD" string
 */
function createTimeline(params) {
  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var reqSheet = ss.getSheetByName('Position Requirements');
  var tz       = Session.getScriptTimeZone();

  if (!reqSheet) {
    return 'Error: "Position Requirements" sheet not found. Run Initial Setup first.';
  }

  // -- Build list of actual calendar work dates ---------------
  // IMPORTANT: Append T12:00:00 to avoid UTC midnight timezone shift
  var weekStartDate = new Date(params.weekStartDate + 'T12:00:00');
  var workDates = [];
  var dayNames = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];
  var fullDayNames = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];

  for (var week = 0; week < 3; week++) {
    for (var day = 0; day < 7; day++) {
      if (params.selectedDays[week][day]) {
        var d = new Date(weekStartDate);
        d.setDate(d.getDate() + (week * 7) + day);
        workDates.push({
          date: d,
          dayName: dayNames[day],
          fullDayName: fullDayNames[day],
          weekNum: week + 1,
          dayIndex: day
        });
      }
    }
  }

  if (workDates.length === 0) {
    return 'Error: No work days selected. Please select at least one day.';
  }

  // -- Calculate effective training hours per day -------------
  var effectiveHoursPerDay = params.hoursPerShift * 0.85;

  // -- Get positions and build day-by-day schedule -----------
  var positions = getPositionsForHouse(params.house, reqSheet);
  var daySchedule = [];

  var posIndex = 0;
  var posHoursRemaining = positions.length > 0 ? positions[0].targetHours : 0;

  for (var i = 0; i < workDates.length; i++) {
    if (posIndex >= positions.length) break;
    var hoursLeftToday = effectiveHoursPerDay;

    while (hoursLeftToday > 0.1 && posIndex < positions.length) {
      var hoursThisBlock = Math.min(hoursLeftToday, posHoursRemaining);
      daySchedule.push({
        date:        workDates[i].date,
        dayName:     workDates[i].dayName,
        fullDayName: workDates[i].fullDayName,
        weekNum:     workDates[i].weekNum,
        position:    positions[posIndex].name,
        hours:       hoursThisBlock
      });
      hoursLeftToday    -= hoursThisBlock;
      posHoursRemaining -= hoursThisBlock;

      if (posHoursRemaining <= 0.1) {
        posIndex++;
        if (posIndex < positions.length) {
          posHoursRemaining = positions[posIndex].targetHours;
        }
      }
    }
  }

  // -- Check if training fits in 3 weeks ---------------------
  var totalTargetHours = 0;
  positions.forEach(function(p) { totalTargetHours += p.targetHours; });
  var totalAvailableHours = workDates.length * effectiveHoursPerDay;
  var fitsIn3Weeks = totalTargetHours <= totalAvailableHours;

  // -- Build weekly summary ----------------------------------
  var weeklySummary = {};
  daySchedule.forEach(function(entry) {
    if (!weeklySummary[entry.weekNum]) weeklySummary[entry.weekNum] = {};
    if (!weeklySummary[entry.weekNum][entry.position]) weeklySummary[entry.weekNum][entry.position] = 0;
    weeklySummary[entry.weekNum][entry.position] += entry.hours;
  });

  var daysPerWeek = {};
  for (var w = 0; w < 3; w++) {
    var count = 0;
    for (var dd = 0; dd < 7; dd++) {
      if (params.selectedDays[w][dd]) count++;
    }
    daysPerWeek[w + 1] = count;
  }

  // -- Create timeline sheet ---------------------------------
  var sheetName = 'Timeline - ' + params.name;
  var existingSheet = ss.getSheetByName(sheetName);
  if (existingSheet) ss.deleteSheet(existingSheet);
  var timelineSheet = ss.insertSheet(sheetName);

  // -- Compose content ---------------------------------------
  var content = [];
  content.push(['TRAINING TIMELINE FOR: ' + params.name.toUpperCase()]);
  content.push(['']);
  content.push(['House: ' + params.house]);
  content.push(['Shift: ' + params.shift]);
  content.push(['Hours per shift: ' + params.hoursPerShift + ' (' + effectiveHoursPerDay.toFixed(1) + ' effective training hrs)']);
  content.push(['']);

  for (var w = 1; w <= 3; w++) {
    var weekMon = new Date(weekStartDate);
    weekMon.setDate(weekMon.getDate() + ((w - 1) * 7));
    var selectedList = [];
    for (var dd = 0; dd < 7; dd++) {
      if (params.selectedDays[w - 1][dd]) selectedList.push(dayNames[dd]);
    }
    content.push(['Week ' + w + ' (' + Utilities.formatDate(weekMon, tz, 'MMM d') + '): ' +
                  selectedList.join(', ') + ' (' + daysPerWeek[w] + ' days)']);
  }
  content.push(['']);

  if (!fitsIn3Weeks) {
    content.push(['** NOTE: Training may extend beyond 3 weeks. ' +
                  'Need ' + totalTargetHours.toFixed(0) + ' hrs, have ' + totalAvailableHours.toFixed(0) + ' hrs available **']);
    content.push(['']);
  }

  content.push(['-------------------------------------------------']);
  content.push(['']);

  for (var w = 1; w <= 3; w++) {
    if (!weeklySummary[w]) continue;
    var weekMon = new Date(weekStartDate);
    weekMon.setDate(weekMon.getDate() + ((w - 1) * 7));
    var weekSun = new Date(weekMon);
    weekSun.setDate(weekSun.getDate() + 6);

    content.push(['WEEK ' + w + '  (' +
                  Utilities.formatDate(weekMon, tz, 'MMM d') + ' - ' +
                  Utilities.formatDate(weekSun, tz, 'MMM d') + ')  |  ' +
                  daysPerWeek[w] + ' work days']);
    content.push(['']);

    var posNames = Object.keys(weeklySummary[w]);
    posNames.forEach(function(posName) {
      var hrs = weeklySummary[w][posName];
      var days = hrs / effectiveHoursPerDay;
      content.push(['  - ' + posName + ':  ~' + days.toFixed(1) + ' days  (' + hrs.toFixed(1) + ' hrs)']);
    });
    content.push(['']);

    content.push(['  Day-by-day:']);
    var weekEntries = daySchedule.filter(function(e) { return e.weekNum === w; });
    var dateGroups = {};
    weekEntries.forEach(function(entry) {
      var dateKey = Utilities.formatDate(entry.date, tz, 'EEE MMM d');
      if (!dateGroups[dateKey]) dateGroups[dateKey] = [];
      if (dateGroups[dateKey].indexOf(entry.position) === -1) {
        dateGroups[dateKey].push(entry.position);
      }
    });
    Object.keys(dateGroups).forEach(function(dateKey) {
      content.push(['    ' + dateKey + ':  ' + dateGroups[dateKey].join(' + ')]);
    });
    content.push(['']);
  }

  content.push(['-------------------------------------------------']);
  content.push(['']);
  content.push(['NOTES:']);
  content.push(['- 85% efficiency factor applied (breaks, setup, etc.)']);
  content.push(['- Positions assigned in order; adjust based on trainee pace']);
  content.push(['- Manager should reassign if trainee masters a position early']);

  // -- Write to sheet ----------------------------------------
  timelineSheet.getRange(1, 1, content.length, 1).setValues(content);
  timelineSheet.getRange('A1').setFontSize(14).setFontWeight('bold');
  timelineSheet.setColumnWidth(1, 650);
  timelineSheet.getRange('A:A').setWrap(true);

  for (var i = 0; i < content.length; i++) {
    var val = String(content[i][0]);
    if (val.indexOf('WEEK ') === 0) {
      timelineSheet.getRange(i + 1, 1).setBackground('#E2EFDA').setFontWeight('bold');
    }
    if (val.indexOf('** NOTE') === 0) {
      timelineSheet.getRange(i + 1, 1).setBackground('#FFC7CE').setFontWeight('bold');
    }
  }

  // -- Save to Training Schedule data sheet ------------------
  writeToTrainingSchedule(ss, daySchedule, params);

  // -- Auto-populate Training Needs for first training week ---
  var populatedMsg = '';
  try {
    if (workDates.length > 0) {
      // Find the Monday of the first training week
      var firstDate = workDates[0].date;
      var firstDow = firstDate.getDay(); // 0=Sun, 1=Mon...
      var firstMonday = new Date(firstDate);
      firstMonday.setDate(firstDate.getDate() - (firstDow === 0 ? 6 : firstDow - 1));
      firstMonday.setHours(0, 0, 0, 0);

      var pCount = populateWeeklyTraining(0, false, ss, firstMonday);
      if (pCount > 0) {
        var weekLabel = Utilities.formatDate(firstMonday, tz, 'MMM d');
        populatedMsg = '\n' + pCount + ' entries loaded into Training Needs (week of ' + weekLabel + ').';
      }
    }
  } catch (e) {
    Logger.log('Auto-populate Training Needs failed: ' + e.message);
  }

  // -- Refresh dashboard ------------------------------------
  updateDashboard();
  ss.setActiveSheet(timelineSheet);

  return 'Timeline created for ' + params.name + '!' + populatedMsg;
}


// ==============================================================
// TRAINING SCHEDULE DATA SHEET
// ==============================================================

/**
 * Creates the Training Schedule sheet if it doesn't exist.
 */
function setupTrainingScheduleSheet(ss) {
  var sheet = ss.getSheetByName('Training Schedule');
  if (sheet) {
    // Check if it's the old 8-column format (missing House column)
    var headerCount = sheet.getLastColumn();
    if (headerCount > 0) {
      var firstHeader = String(sheet.getRange(1, 2).getValue()).trim();
      if (firstHeader === 'Date') {
        // Old format detected - House column missing. Clear and rebuild.
        Logger.log('Training Schedule: migrating from 8-col to 9-col format (adding House)');
        sheet.clear();
        var headers = [['Trainee', 'House', 'Date', 'Day', 'Shift', 'Position', 'Hours', 'Week #', 'Created']];
        sheet.getRange(1, 1, 1, 9).setValues(headers).setFontWeight('bold').setBackground('#D9E2F3');
        sheet.setFrozenRows(1);
      }
    }
    return sheet;
  }

  sheet = ss.insertSheet('Training Schedule');
  var headers = [['Trainee', 'House', 'Date', 'Day', 'Shift', 'Position', 'Hours', 'Week #', 'Created']];
  sheet.getRange(1, 1, 1, 9).setValues(headers).setFontWeight('bold').setBackground('#D9E2F3');
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 60);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(6, 130);
  return sheet;
}

/**
 * Writes day schedule to Training Schedule.
 * Clears old entries for this trainee+shift before writing.
 */
function writeToTrainingSchedule(ss, daySchedule, params) {
  var schedSheet = setupTrainingScheduleSheet(ss);
  var now = new Date();

  // Delete existing entries for this trainee+shift (bottom-up)
  if (schedSheet.getLastRow() > 1) {
    var existing = schedSheet.getRange(2, 1, schedSheet.getLastRow() - 1, 5).getValues();
    for (var i = existing.length - 1; i >= 0; i--) {
      if (String(existing[i][0]).trim() === params.name &&
          String(existing[i][4]).trim() === params.shift) {
        schedSheet.deleteRow(i + 2);
      }
    }
  }

  // Build and write new rows in batch
  var newRows = daySchedule.map(function(entry) {
    return [
      params.name,
      params.house,
      entry.date,
      entry.fullDayName,
      params.shift,
      entry.position,
      Math.round(entry.hours * 100) / 100,
      entry.weekNum,
      now
    ];
  });

  if (newRows.length > 0) {
    var startRow = schedSheet.getLastRow() + 1;
    schedSheet.getRange(startRow, 1, newRows.length, 9).setValues(newRows);
  }

  Logger.log('Training Schedule: wrote ' + newRows.length + ' entries for ' + params.name);
}


// ==============================================================
// TRAINING NEEDS POPULATION
// ==============================================================

/** Menu-callable: loads current week into Training Needs. */
function loadThisWeeksTraining() {
  populateWeeklyTraining(0, true);
}

/** Menu-callable: loads next week into Training Needs. */
function loadNextWeeksTraining() {
  populateWeeklyTraining(1, true);
}

/** Trigger-callable: auto-populates on Monday mornings. */
function mondayAutoPopulate() {
  populateWeeklyTraining(0, false);
  Logger.log('Monday auto-populate completed: ' + new Date());
}

/** Quiet version called from createTimeline. Returns count. */
function populateWeeklyTrainingQuiet(ss) {
  return populateWeeklyTraining(0, false, ss);
}

/**
 * Main population logic.
 *
 * 1. Finds the target week's Monday-Saturday range
 * 2. Reads Training Schedule for matching dates (FOH only)
 * 3. Scans Training Needs sheet to detect layout dynamically
 * 4. Clears existing data in Training Needs
 * 5. Groups entries by day + shift + trainee
 * 6. Sorts positions (unstarted first = new learning priority)
 * 7. Writes combined position strings with hours
 *
 * @param {number} weekOffset  0=current week, 1=next week, etc.
 * @param {boolean} showAlerts  Show UI alerts (for manual runs)
 * @param {Spreadsheet} ssOverride  Optional spreadsheet reference
 * @param {Date} targetMonday  Optional: populate this specific week instead of using offset
 * @return {number} Count of entries populated
 */
function populateWeeklyTraining(weekOffset, showAlerts, ssOverride, targetMonday) {
  weekOffset = weekOffset || 0;
  var ss = ssOverride || SpreadsheetApp.getActiveSpreadsheet();
  var tz = Session.getScriptTimeZone();

  var schedSheet = ss.getSheetByName('Training Schedule');
  var needsSheet = ss.getSheetByName('Training Needs');

  if (!schedSheet || schedSheet.getLastRow() < 2) {
    if (showAlerts) SpreadsheetApp.getUi().alert('No Training Schedule data found.\nGenerate a timeline first.');
    return 0;
  }
  if (!needsSheet) {
    if (showAlerts) SpreadsheetApp.getUi().alert('"Training Needs" sheet not found.');
    return 0;
  }

  // -- Find target week's Monday through Saturday ------------
  var monday;
  if (targetMonday) {
    monday = new Date(targetMonday);
  } else {
    var today = new Date();
    var dow = today.getDay(); // 0=Sun,1=Mon...
    monday = new Date(today);
    monday.setDate(today.getDate() - (dow === 0 ? 6 : dow - 1) + (weekOffset * 7));
  }
  monday.setHours(0, 0, 0, 0);

  var saturday = new Date(monday);
  saturday.setDate(monday.getDate() + 5);

  var mondayStr   = Utilities.formatDate(monday, tz, 'yyyy-MM-dd');
  var saturdayStr = Utilities.formatDate(saturday, tz, 'yyyy-MM-dd');

  // -- Read Training Schedule and filter ---------------------
  // Columns: A=Trainee, B=House, C=Date, D=Day, E=Shift, F=Position, G=Hours, H=Week#, I=Created
  var schedData = schedSheet.getRange(2, 1, schedSheet.getLastRow() - 1, 9).getValues();

  var weekEntries = [];
  schedData.forEach(function(row) {
    var house = String(row[1]).trim();

    // FOH ONLY - Training Needs sheet is for FOH trainees
    if (house !== 'FOH') return;

    var entryDate = new Date(row[2]);
    var entryStr = Utilities.formatDate(entryDate, tz, 'yyyy-MM-dd');
    if (entryStr >= mondayStr && entryStr <= saturdayStr) {
      weekEntries.push({
        trainee:   String(row[0]).trim(),
        house:     house,
        date:      entryDate,
        dateStr:   entryStr,
        dayOfWeek: String(row[3]).trim(),
        shift:     String(row[4]).trim(),
        position:  String(row[5]).trim(),
        hours:     parseFloat(row[6]) || 0
      });
    }
  });

  if (weekEntries.length === 0) {
    if (showAlerts) {
      SpreadsheetApp.getUi().alert(
        'No training scheduled for week of ' +
        Utilities.formatDate(monday, tz, 'MMM d, yyyy') + '.');
    }
    return 0;
  }

  // -- Detect Training Needs sheet layout --------------------
  var layout = loadTrainingNeedsLayout(needsSheet);
  if (!layout || Object.keys(layout).length === 0) {
    if (showAlerts) {
      SpreadsheetApp.getUi().alert(
        'Could not detect Training Needs layout.\n' +
        'Check that day names (MONDAY, Tuesday, etc.) and\n' +
        '"FOH Training" headers are present on the sheet.');
    }
    return 0;
  }

  // -- Clear existing Training Needs data --------------------
  clearTrainingNeedsData(needsSheet, layout);

  // -- Get logged hours for position priority ----------------
  var loggedHours = getAllTraineeLoggedHours(ss);

  // -- Group entries by day + shift + trainee ----------------
  var grouped = {};
  weekEntries.forEach(function(entry) {
    var key = entry.dayOfWeek + '|' + entry.shift + '|' + entry.trainee;
    if (!grouped[key]) {
      grouped[key] = {
        trainee:   entry.trainee,
        dayOfWeek: entry.dayOfWeek,
        shift:     entry.shift,
        date:      entry.date,
        positions: []
      };
    }
    grouped[key].positions.push({ name: entry.position, hours: entry.hours });
  });

  // -- Write to Training Needs -------------------------------
  var populated = 0;

  Object.keys(grouped).forEach(function(key) {
    var entry = grouped[key];
    var dayLayout = layout[entry.dayOfWeek];
    if (!dayLayout) {
      Logger.log('populateWeekly: no layout for day "' + entry.dayOfWeek + '"');
      return;
    }

    // Sort: unstarted positions first (new learning priority)
    var traineeHours = loggedHours[entry.trainee] || {};
    entry.positions.sort(function(a, b) {
      var aLogged = traineeHours[a.name] || 0;
      var bLogged = traineeHours[b.name] || 0;
      if (aLogged === 0 && bLogged > 0) return -1;
      if (aLogged > 0 && bLogged === 0) return 1;
      return aLogged - bLogged;
    });

    // Format: "iPOS (2.0 hrs) / Register/POS (4.0 hrs)"
    var posStr = entry.positions.map(function(p) {
      return p.name + ' (' + p.hours.toFixed(1) + ' hrs)';
    }).join(' / ');

    // Determine all dayparts this shift spans based on total hours
    var totalEffHours = entry.positions.reduce(function(sum, p) { return sum + p.hours; }, 0);
    var coveredDayparts = getDaypartsForShift(entry.shift, totalEffHours);

    coveredDayparts.forEach(function(daypartName) {
      var section = dayLayout.sections[daypartName];
      if (!section) {
        Logger.log('populateWeekly: no section for "' + daypartName + '" on ' + entry.dayOfWeek);
        return;
      }

      var cols = section.cols;
      if (!cols || !cols.nameCol || !cols.posCol) {
        Logger.log('populateWeekly: missing columns for ' + entry.dayOfWeek + ' ' + daypartName);
        return;
      }

      // Find empty row
      var emptyRow = -1;
      for (var r = section.dataStart; r <= section.dataEnd; r++) {
        var cellVal = needsSheet.getRange(r, cols.nameCol).getValue();
        if (!cellVal || String(cellVal).trim() === '') {
          emptyRow = r;
          break;
        }
      }

      if (emptyRow === -1) {
        Logger.log('No empty row for ' + entry.dayOfWeek + ' ' + daypartName);
        return;
      }

      needsSheet.getRange(emptyRow, cols.nameCol).setValue(entry.trainee);
      if (cols.shiftCol) needsSheet.getRange(emptyRow, cols.shiftCol).setValue(entry.shift);
      needsSheet.getRange(emptyRow, cols.posCol).setValue(posStr);
      if (cols.dateCol) {
        needsSheet.getRange(emptyRow, cols.dateCol).setValue(
          Utilities.formatDate(entry.date, tz, 'M/d/yyyy'));
      }
      populated++;
    });
  });

  if (showAlerts) {
    SpreadsheetApp.getUi().alert(
      'Training Needs Updated!\n\n' +
      'Week of: ' + Utilities.formatDate(monday, tz, 'MMM d, yyyy') + '\n' +
      'Entries loaded: ' + populated);
  }

  return populated;
}


// ==============================================================
// TRAINING NEEDS LAYOUT DETECTION
// ==============================================================

/**
 * Dynamically scans the Training Needs sheet to find where each
 * day's sections are and what columns contain what data.
 *
 * The sheet has all 6 days (Mon-Sat) stacked vertically with
 * 4 meal periods per day:
 *   Left side:  Breakfast + Lunch
 *   Right side: Afternoon + Dinner
 *
 * Detection logic:
 *   1. Scan all rows for day name text (MONDAY, Tuesday, etc.)
 *   2. For each day, find "FOH Training" header rows below it
 *   3. First header = Breakfast/Afternoon, second = Lunch/Dinner
 *   4. Auto-detect column positions from header text
 *
 * Returns: { 'Monday': { dayRow, sections: { 'Breakfast': {...}, ... } }, ... }
 */
function loadTrainingNeedsLayout(sheet) {
  var numRows = Math.min(sheet.getLastRow(), 200);
  if (numRows < 5) return null;

  // Read columns A through N
  var allData = sheet.getRange(1, 1, numRows, 14).getValues();
  var days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  var layout = {};

  // -- Pass 1: Find each day's row ---------------------------
  for (var r = 0; r < allData.length; r++) {
    var rowText = '';
    for (var c = 0; c < allData[r].length; c++) {
      rowText += ' ' + String(allData[r][c]);
    }
    rowText = rowText.toLowerCase();

    for (var di = 0; di < days.length; di++) {
      if (rowText.indexOf(days[di].toLowerCase()) > -1 && !layout[days[di]]) {
        layout[days[di]] = { dayRow: r + 1, sections: {} };
      }
    }
  }

  // -- Pass 2: Find FOH Training headers for each day --------
  Object.keys(layout).forEach(function(dayName) {
    var dayRow = layout[dayName].dayRow;
    var fohHeaderRows = [];

    // Look within 20 rows after the day marker
    // Scan ALL columns in each row for "FOH Training" / "BOH Training"
    // (header may be in col A, B, J, or elsewhere depending on layout)
    var searchEnd = Math.min(dayRow + 19, allData.length);
    for (var r = dayRow - 1; r < searchEnd; r++) {
      var isHeaderRow = false;
      for (var c = 0; c < allData[r].length; c++) {
        var cellVal = String(allData[r][c]).trim().toLowerCase();
        if (cellVal.indexOf('foh training') > -1 || cellVal.indexOf('boh training') > -1) {
          isHeaderRow = true;
          break;
        }
      }
      if (isHeaderRow) {
        fohHeaderRows.push(r + 1);
      }
    }

    // First header = Breakfast/Afternoon, second = Lunch/Dinner
    if (fohHeaderRows.length >= 1) {
      var h1 = fohHeaderRows[0];
      var leftCols1 = detectSectionColumns(sheet, h1, 1, 8);
      var rightCols1 = detectSectionColumns(sheet, h1, 9, 14);

      layout[dayName].sections['Breakfast'] = {
        headerRow: h1, dataStart: h1 + 1, dataEnd: h1 + 5,
        side: 'left', cols: leftCols1
      };
      if (rightCols1.nameCol) {
        layout[dayName].sections['Afternoon'] = {
          headerRow: h1, dataStart: h1 + 1, dataEnd: h1 + 5,
          side: 'right', cols: rightCols1
        };
      }
    }

    if (fohHeaderRows.length >= 2) {
      var h2 = fohHeaderRows[1];
      var leftCols2 = detectSectionColumns(sheet, h2, 1, 8);
      var rightCols2 = detectSectionColumns(sheet, h2, 9, 14);

      layout[dayName].sections['Lunch'] = {
        headerRow: h2, dataStart: h2 + 1, dataEnd: h2 + 5,
        side: 'left', cols: leftCols2
      };
      if (rightCols2.nameCol) {
        layout[dayName].sections['Dinner'] = {
          headerRow: h2, dataStart: h2 + 1, dataEnd: h2 + 5,
          side: 'right', cols: rightCols2
        };
      }
    }
  });

  // Log detected layout for debugging
  Object.keys(layout).forEach(function(dayName) {
    var sNames = Object.keys(layout[dayName].sections);
    Logger.log('Layout: ' + dayName + ' (row ' + layout[dayName].dayRow + ') -> ' + sNames.join(', '));
    sNames.forEach(function(sn) {
      var s = layout[dayName].sections[sn];
      Logger.log('  ' + sn + ': rows ' + s.dataStart + '-' + s.dataEnd +
                 ' nameCol=' + (s.cols.nameCol || '?') +
                 ' posCol=' + (s.cols.posCol || '?') +
                 ' dateCol=' + (s.cols.dateCol || 'none'));
    });
  });

  return layout;
}

/**
 * Reads header cells in a row range to auto-detect column purposes.
 * Looks for: FOH Training/name, Shift, Position, Trainer, Date.
 */
function detectSectionColumns(sheet, headerRow, startCol, endCol) {
  var cols = {};
  for (var c = startCol; c <= endCol; c++) {
    var val = String(sheet.getRange(headerRow, c).getValue()).trim().toLowerCase();
    if (!val) continue;

    if ((val.indexOf('training') > -1 || val.indexOf('foh') > -1 ||
         val.indexOf('boh') > -1 || val === 'name' || val === 'trainee') && !cols.nameCol) {
      cols.nameCol = c;
    } else if (val === 'shift' && !cols.shiftCol) {
      cols.shiftCol = c;
    } else if (val === 'position' && !cols.posCol) {
      cols.posCol = c;
    } else if (val === 'trainer' && !cols.trainerCol) {
      cols.trainerCol = c;
    } else if (val.indexOf('date') > -1 && !cols.dateCol) {
      cols.dateCol = c;
    }
  }
  return cols;
}


// ==============================================================
// DATA HELPERS
// ==============================================================

/**
 * Clears Name, Shift, Position, Date cells in Training Needs
 * data rows. Preserves dropdowns, trainer assignments, and formatting.
 */
function clearTrainingNeedsData(sheet, layout) {
  Object.keys(layout).forEach(function(dayName) {
    var sections = layout[dayName].sections;
    Object.keys(sections).forEach(function(shiftName) {
      var section = sections[shiftName];
      var cols = section.cols;

      for (var r = section.dataStart; r <= section.dataEnd; r++) {
        if (cols.nameCol) sheet.getRange(r, cols.nameCol).clearContent();
        if (cols.shiftCol) sheet.getRange(r, cols.shiftCol).clearContent();
        if (cols.posCol) sheet.getRange(r, cols.posCol).clearContent();
        // Trainer column intentionally NOT cleared (manually assigned)
        if (cols.dateCol) sheet.getRange(r, cols.dateCol).clearContent();
      }
    });
  });
}

/**
 * Gets all trainee logged hours from Daily Training Log or Form_Responses.
 * Used for position priority (unstarted positions sort first).
 *
 * Returns: { 'Name': { 'iPOS': 4.0, 'Register/POS': 2.0 }, ... }
 */
function getAllTraineeLoggedHours(ss) {
  var result = {};
  var data = [];

  // Try Daily Training Log first
  var logSheet = ss.getSheetByName('Daily Training Log');
  if (logSheet && logSheet.getLastRow() > 1) {
    data = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 9).getValues();
  }

  // Fall back to Form_Responses
  if (data.length === 0) {
    var formNames = ['Form_Responses', 'Form Responses 1', 'Form responses 1', 'Form_Responses_1'];
    for (var i = 0; i < formNames.length; i++) {
      var fs = ss.getSheetByName(formNames[i]);
      if (fs && fs.getLastRow() > 1) {
        var rawData = fs.getRange(2, 1, fs.getLastRow() - 1, 7).getValues();
        rawData.forEach(function(row) {
          var name = String(row[2] || '').trim();
          var tcName = name.split(' ').map(function(w) {
            return w ? w.charAt(0).toUpperCase() + w.slice(1).toLowerCase() : '';
          }).join(' ');
          data.push([row[0], row[1], name, row[3], row[4], row[5], row[6], tcName, false]);
        });
        break;
      }
    }
  }

  data.forEach(function(row) {
    var name = String(row[7] || row[2] || '').trim();
    var position = String(row[3] || '').trim();
    var hours = parseFloat(row[4]) || 0;
    if (!name || !position) return;

    // Title-case normalize
    name = name.split(' ').map(function(w) {
      return w ? w.charAt(0).toUpperCase() + w.slice(1).toLowerCase() : '';
    }).join(' ');

    if (!result[name]) result[name] = {};

    // Handle multi-position entries
    var positions;
    if (position.indexOf(',') > -1) {
      positions = position.split(',').map(function(p) {
        return (typeof normalizePositionName === 'function') ? normalizePositionName(p) : p.trim();
      });
    } else {
      positions = [(typeof normalizePositionName === 'function') ? normalizePositionName(position) : position.trim()];
    }

    var hrsPerPos = hours / positions.length;
    positions.forEach(function(pos) {
      if (!result[name][pos]) result[name][pos] = 0;
      result[name][pos] += hrsPerPos;
    });
  });

  return result;
}


// ==============================================================
// TRIGGER MANAGEMENT
// ==============================================================

/**
 * Sets up Monday 5 AM trigger for auto-populating Training Needs.
 */
function setupMondayTrigger() {
  // Remove existing triggers for this function
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'mondayAutoPopulate') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('mondayAutoPopulate')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(5)
    .create();

  SpreadsheetApp.getUi().alert(
    'Monday auto-populate trigger set!\n\n' +
    'Every Monday at ~5 AM, Training Needs will be\n' +
    'automatically loaded from the Training Schedule.');
}


// ==============================================================
// EXISTING HELPERS
// ==============================================================

/**
 * Returns all daypart names that a shift covers, based on the starting
 * daypart and total effective training hours for that day.
 *
 * Daypart time windows:
 *   Breakfast  6:00 - 11:00
 *   Lunch     11:00 - 15:00
 *   Afternoon 15:00 - 17:00
 *   Dinner    17:00 - 22:00
 *
 * Example: Breakfast + 8-hour shift (6am-2pm) covers Breakfast AND Lunch.
 *
 * @param {string} startShift  The named starting daypart (e.g. "Breakfast")
 * @param {number} totalEffectiveHours  Sum of position hours for the day (after 85% factor)
 * @return {Array} Daypart names covered by this shift
 */
function getDaypartsForShift(startShift, totalEffectiveHours) {
  var DAYPARTS = [
    { name: 'Breakfast', start: 6,  end: 11 },
    { name: 'Lunch',     start: 11, end: 15 },
    { name: 'Afternoon', start: 15, end: 17 },
    { name: 'Dinner',    start: 17, end: 22 }
  ];

  // Find the start hour for the named shift
  var startHour = -1;
  for (var i = 0; i < DAYPARTS.length; i++) {
    if (DAYPARTS[i].name === startShift) {
      startHour = DAYPARTS[i].start;
      break;
    }
  }
  if (startHour === -1) return [startShift]; // unknown shift, fall back to named only

  // Reverse the 85% efficiency factor to get actual clock hours
  var shiftHours = totalEffectiveHours / 0.85;
  var endHour = startHour + shiftHours;

  // Collect all dayparts whose window overlaps [startHour, endHour)
  var result = [];
  for (var j = 0; j < DAYPARTS.length; j++) {
    var dp = DAYPARTS[j];
    if (dp.start < endHour && dp.end > startHour) {
      result.push(dp.name);
    }
  }
  return result.length > 0 ? result : [startShift];
}

function weeklyDashboardRefresh() {
  updateDashboard();
  Logger.log('Weekly dashboard refresh completed: ' + new Date());
}

function getPositionsForHouse(house, reqSheet) {
  var data      = reqSheet.getDataRange().getValues();
  var positions = [];

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === house) {
      positions.push({
        name:        data[i][1],
        minHours:    data[i][2],
        maxHours:    data[i][3],
        targetHours: data[i][4]
      });
    }
  }

  return positions;
}
