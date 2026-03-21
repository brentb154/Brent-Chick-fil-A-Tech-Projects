/**
 * ============================================================
 * TRAINING TRACKING SYSTEM - Dashboard & Position Tracking
 * ============================================================
 * Recalculates all dashboard metrics, active trainee summaries,
 * position-level progress, and certification readiness.
 */

// -- Main Dashboard Refresh -----------------------------------

/**
 * Recalculates every section of the Master Dashboard.
 * Called after form submissions and manually via menu.
 */
function updateDashboard() {
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet  = ss.getSheetByName('Daily Training Log');
  var dashSheet = ss.getSheetByName('Master Dashboard');
  var reqSheet  = ss.getSheetByName('Position Requirements');
  var certSheet = ss.getSheetByName('Certification Log');

  if (!dashSheet || !reqSheet || !certSheet) {
    Logger.log('updateDashboard: missing required sheets');
    return;
  }

  // -- Gather data - try Daily Training Log first, fall back to Form_Responses --
  var logData = [];
  var dataSource = '';

  // Try Daily Training Log
  if (logSheet && logSheet.getLastRow() > 1) {
    logData = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 9).getValues();
    dataSource = 'Daily Training Log';
  }

  // If empty, try Form_Responses variants
  if (logData.length === 0) {
    var formSheetNames = ['Form_Responses', 'Form Responses 1', 'Form responses 1', 'Form_Responses_1'];
    for (var fi = 0; fi < formSheetNames.length; fi++) {
      var formSheet = ss.getSheetByName(formSheetNames[fi]);
      if (formSheet && formSheet.getLastRow() > 1) {
        var rawData = formSheet.getRange(2, 1, formSheet.getLastRow() - 1, formSheet.getLastColumn()).getValues();
        // Map form columns to expected log format:
        // Form: A=Timestamp, B=Date, C=Name, D=Position, E=Hours, F=OnTrack, G=Notes
        // Log:  A=Timestamp, B=Date, C=Name, D=Position, E=Hours, F=OnTrack, G=Notes, H=CanonicalName, I=Synced
        rawData.forEach(function(row) {
          var name = String(row[2] || '').trim();
          if (!name) return;
          // Basic title-case normalization (handles "bald" -> "Bald")
          var normalizedName = name.split(' ').map(function(word) {
            if (!word) return '';
            return word.charAt(0).toUpperCase() + word.slice(1).toLowerCase();
          }).join(' ');
          logData.push([
            row[0],                   // Timestamp
            row[1],                   // Date
            name,                     // Trainee Name (original)
            String(row[3] || ''),     // Position
            row[4] || 0,              // Hours
            String(row[5] || ''),     // On Track
            String(row[6] || ''),     // Notes
            normalizedName,           // Canonical Name (title-cased)
            false                     // Synced
          ]);
        });
        dataSource = formSheetNames[fi];
        break;
      }
    }
  }

  if (logData.length === 0) {
    dashSheet.getRange('B4').setValue(0);
    dashSheet.getRange('B5').setValue(0);
    dashSheet.getRange('B6').setValue(0);
    dashSheet.getRange('B7').setValue(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM/dd/yyyy h:mm a'));
    return;
  }

  Logger.log('updateDashboard: using ' + dataSource + ' (' + logData.length + ' rows)');

  var certifiedNames = [];
  if (certSheet.getLastRow() > 1) {
    certifiedNames = certSheet.getRange(2, 1, certSheet.getLastRow() - 1, 1)
      .getValues().map(function (r) { return String(r[0]).trim(); });
  }

  // -- Build trainee map ------------------------------------
  var trainees = {};

  // First: scan for Timeline sheets to pick up scheduled trainees
  // who may not have any log entries yet
  var allSheets = ss.getSheets();
  allSheets.forEach(function (sheet) {
    var name = sheet.getName();
    if (name.indexOf('Timeline - ') === 0) {
      var traineeName = name.replace('Timeline - ', '').trim();
      if (!traineeName) return;
      if (certifiedNames.indexOf(traineeName) > -1) return;

      // Read the timeline sheet to get house info (row with "House:")
      var timelineData = sheet.getRange('A1:A10').getValues();
      var house = '';
      for (var i = 0; i < timelineData.length; i++) {
        var cellVal = String(timelineData[i][0]);
        if (cellVal.indexOf('House: ') === 0) {
          house = cellVal.replace('House: ', '').trim();
          break;
        }
      }

      // Add as scheduled trainee (will be overwritten if log data exists)
      if (!trainees[traineeName]) {
        trainees[traineeName] = {
          totalHours:    0,
          positions:     {},
          firstDate:     new Date(),
          lastDate:      new Date(),
          lastPosition:  'N/A',
          onTrackCount:  0,
          totalEntries:  0,
          house:         house || 'TBD',
          isScheduledOnly: true  // Flag: no log entries yet
        };
      }
    }
  });

  // Then: overlay actual log data (overrides scheduled-only status)
  logData.forEach(function (row) {
    var canonicalName = String(row[7]).trim(); // Column H
    if (!canonicalName) canonicalName = String(row[2]).trim(); // fallback to Col C
    if (!canonicalName) return;

    var position = String(row[3]).trim();
    var hours    = parseFloat(row[4]) || 0;
    var date     = new Date(row[1]);
    var onTrack  = String(row[5]);

    // Handle multi-position entries (e.g., "iPOS, Register")
    var positionList = position.indexOf(',') > -1
      ? position.split(',').map(function(p) { return normalizePositionName(p); }).filter(function(p) { return p; })
      : [normalizePositionName(position)];
    var hoursPerPos = positionList.length > 0 ? hours / positionList.length : hours;

    // Skip certified trainees
    if (certifiedNames.indexOf(canonicalName) > -1) return;

    if (!trainees[canonicalName]) {
      trainees[canonicalName] = {
        totalHours:    0,
        positions:     {},
        firstDate:     date,
        lastDate:      date,
        lastPosition:  positionList[0],
        onTrackCount:  0,
        totalEntries:  0,
        house:         inferHouse(positionList[0]),
        isScheduledOnly: false
      };
    } else {
      // Clear scheduled-only flag since we now have real data
      trainees[canonicalName].isScheduledOnly = false;
    }

    var t = trainees[canonicalName];
    t.totalHours   += hours;
    t.totalEntries += 1;

    if (onTrack === 'Yes' || onTrack === 'Y' || onTrack === 'yes') {
      t.onTrackCount++;
    }

    if (date < t.firstDate) t.firstDate = date;
    if (date > t.lastDate) {
      t.lastDate     = date;
      t.lastPosition = positionList[0];
    }

    // Credit hours to each position in the list
    positionList.forEach(function(pos) {
      if (!t.positions[pos]) t.positions[pos] = 0;
      t.positions[pos] += hoursPerPos;
    });
  });

  // -- Write Active Trainees Summary (rows 14+) -------------
  var traineeNames = Object.keys(trainees).sort();
  var summaryData  = [];

  traineeNames.forEach(function (name) {
    var t = trainees[name];

    // Scheduled-only trainees (have timeline but no log entries yet)
    if (t.isScheduledOnly) {
      summaryData.push([
        name,
        t.house,
        '0.0',
        '-',
        'N/A',
        '-',
        '-',
        'Scheduled'
      ]);
      return;
    }

    var daysSinceLast = Math.max(0, Math.round((new Date() - t.lastDate) / (1000 * 60 * 60 * 24)));
    var onTrackPct    = t.totalEntries > 0 ? (t.onTrackCount / t.totalEntries) * 100 : 0;
    var certStatus    = isCertificationReady(name, t.house, t.positions) ? '✅ READY' : 'In Progress';

    summaryData.push([
      name,
      t.house,
      t.totalHours.toFixed(1),
      daysSinceLast,
      t.lastPosition,
      Utilities.formatDate(t.lastDate, Session.getScriptTimeZone(), 'MM/dd/yyyy'),
      onTrackPct.toFixed(0) + '%',
      certStatus
    ]);
  });

  // Clear old trainee data AND formatting (rows 14 through buffer)
  var traineeEndRow = Math.max(dashSheet.getLastRow(), 30);
  if (traineeEndRow >= 14) {
    dashSheet.getRange(14, 1, traineeEndRow - 13, 8).clear();
  }

  if (summaryData.length > 0) {
    dashSheet.getRange(14, 1, summaryData.length, 8).setValues(summaryData);

    // Conditional formatting for cert status
    for (var i = 0; i < summaryData.length; i++) {
      var row = 14 + i;
      var statusVal = summaryData[i][7];

      if (statusVal === '✅ READY') {
        dashSheet.getRange(row, 8).setBackground('#C6EFCE').setFontColor('#006100');
      } else if (statusVal === 'Scheduled') {
        dashSheet.getRange(row, 8).setBackground('#D9E2F3').setFontColor('#1F4E79');
      } else {
        dashSheet.getRange(row, 8).setBackground('#FFF2CC').setFontColor('#9C5700');
      }

      // Highlight inactive trainees (3+ days)
      if (summaryData[i][3] >= 3) {
        dashSheet.getRange(row, 4).setBackground('#FFC7CE').setFontColor('#9C0006');
      } else {
        dashSheet.getRange(row, 4).setBackground(null).setFontColor(null);
      }
    }
  }

  // -- Write Position Progress Detail (rows 33+) ------------
  updatePositionProgress(trainees, dashSheet, reqSheet);

  // -- Write Quick Stats ------------------------------------
  updateQuickStats(trainees, logData, dashSheet);
}


// -- Position Progress (Matrix Grid) --------------------------

function updatePositionProgress(trainees, dashSheet, reqSheet) {
  var reqData = reqSheet.getDataRange().getValues();

  // Get positions grouped by house
  var housePositions = {};  // { 'FOH': [{name, min}, ...], 'BOH': [...] }
  reqData.slice(1).forEach(function(req) {
    var house = req[0];
    var pos   = req[1];
    var min   = req[2];
    if (!housePositions[house]) housePositions[house] = [];
    housePositions[house].push({ name: pos, min: min });
  });

  // Group trainees by house
  var traineesByHouse = {};
  Object.keys(trainees).sort().forEach(function(name) {
    var house = trainees[name].house || 'TBD';
    if (!traineesByHouse[house]) traineesByHouse[house] = [];
    traineesByHouse[house].push({ name: name, data: trainees[name] });
  });

  // Clear everything below row 31 (old position data + formatting)
  var clearEndRow = Math.max(dashSheet.getLastRow(), 60);
  if (clearEndRow >= 31) {
    dashSheet.getRange(31, 1, clearEndRow - 30, 20).clear();
  }

  var currentRow = 31;

  // Section header
  dashSheet.getRange(currentRow, 1).setValue('-- POSITION PROGRESS --')
    .setFontWeight('bold').setBackground('#E2EFDA');
  dashSheet.getRange(currentRow, 2, 1, 15).setBackground('#E2EFDA');
  currentRow++;

  // Build matrix for each house
  var houses = Object.keys(housePositions).sort();

  houses.forEach(function(house) {
    var positions = housePositions[house];
    var houseTrainees = traineesByHouse[house] || [];

    if (houseTrainees.length === 0) return;

    // House sub-header
    dashSheet.getRange(currentRow, 1).setValue(house).setFontWeight('bold')
      .setFontSize(11).setBackground('#D9E2F3');
    dashSheet.getRange(currentRow, 2, 1, positions.length).setBackground('#D9E2F3');
    currentRow++;

    // Column headers: Trainee | pos1 | pos2 | pos3 | ...
    var headerRow = ['Trainee'];
    positions.forEach(function(pos) {
      // Abbreviate long names to fit columns
      var label = pos.name;
      if (label.length > 10) {
        label = label.replace('Register/', 'Reg/')
                     .replace('Primary ', '')
                     .replace('Raw ', '');
      }
      headerRow.push(label);
    });

    var numCols = headerRow.length;
    dashSheet.getRange(currentRow, 1, 1, numCols).setValues([headerRow]);
    dashSheet.getRange(currentRow, 1, 1, numCols).setFontWeight('bold')
      .setFontSize(9).setBackground('#F3F3F3');
    // Rotate position headers for compactness
    dashSheet.getRange(currentRow, 2, 1, numCols - 1)
      .setHorizontalAlignment('center').setVerticalAlignment('bottom');
    currentRow++;

    // One row per trainee
    houseTrainees.forEach(function(entry) {
      var t = entry.data;
      var rowData = [entry.name];

      positions.forEach(function(pos) {
        var logged = t.positions[pos.name] || 0;
        var loggedRound = Math.round(logged * 10) / 10;
        rowData.push(loggedRound + '/' + pos.min);
      });

      dashSheet.getRange(currentRow, 1, 1, numCols).setValues([rowData]);
      dashSheet.getRange(currentRow, 1).setFontWeight('bold').setFontSize(9);
      dashSheet.getRange(currentRow, 2, 1, numCols - 1)
        .setHorizontalAlignment('center').setFontSize(9);

      // Color-code each cell
      for (var c = 0; c < positions.length; c++) {
        var logged = t.positions[positions[c].name] || 0;
        var min = positions[c].min;
        var cell = dashSheet.getRange(currentRow, c + 2);

        if (logged >= min) {
          // Complete - green
          cell.setBackground('#C6EFCE').setFontColor('#006100');
        } else if (logged > 0) {
          // In progress - yellow
          cell.setBackground('#FFF2CC').setFontColor('#9C5700');
        } else {
          // Not started - light gray
          cell.setBackground('#F2F2F2').setFontColor('#999999');
        }
      }

      currentRow++;
    });

    // Blank row between houses
    currentRow++;
  });

  // Note: column widths not modified here to preserve trainee summary layout above
}


// -- Quick Stats ----------------------------------------------

function updateQuickStats(trainees, logData, dashSheet) {
  var traineeCount = Object.keys(trainees).length;

  // Total hours this week
  var weekStart = new Date();
  weekStart.setDate(weekStart.getDate() - weekStart.getDay());
  weekStart.setHours(0, 0, 0, 0);

  var weekHours = 0;
  logData.forEach(function (row) {
    var entryDate = new Date(row[1]);
    if (entryDate >= weekStart) {
      weekHours += (parseFloat(row[4]) || 0);
    }
  });

  // Cert-ready count (exclude scheduled-only trainees)
  var certReady = 0;
  var activeTraineeCount = 0;
  Object.keys(trainees).forEach(function (name) {
    var t = trainees[name];
    if (t.isScheduledOnly) return;  // Don't count scheduled trainees
    activeTraineeCount++;
    if (isCertificationReady(name, t.house, t.positions)) {
      certReady++;
    }
  });

  dashSheet.getRange('B4').setValue(traineeCount + ' (' + activeTraineeCount + ' active)');
  dashSheet.getRange('B5').setValue(weekHours.toFixed(1));
  dashSheet.getRange('B6').setValue(certReady);
  dashSheet.getRange('B7').setValue(
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM/dd/yyyy h:mm a')
  );
}


// -- House & Certification Helpers ----------------------------

/**
 * Infer FOH or BOH from the position name.
 */
function inferHouse(position) {
  // Normalize first so abbreviations are caught
  var normalized = normalizePositionName(position);
  var fohPositions = [
    'iPOS', 'Register/POS', 'Cash Cart', 'Server', 'FC Drinks',
    'Desserts', 'DT Drinks', 'DT Stuffer', 'FC Bagger', 'DT Bagger', 'Window'
  ];

  return fohPositions.indexOf(normalized) > -1 ? 'FOH' : 'BOH';
}


/**
 * Maps common position name variations from form entries
 * to the official names in Position Requirements.
 * Add new mappings here as needed.
 */
function normalizePositionName(position) {
  var aliases = {
    'register':     'Register/POS',
    'pos':          'Register/POS',
    'ipos':         'iPOS',
    'cash':         'Cash Cart',
    'fc drink':     'FC Drinks',
    'fc drinks':    'FC Drinks',
    'dt drink':     'DT Drinks',
    'dt drinks':    'DT Drinks',
    'dt stuff':     'DT Stuffer',
    'fc bag':       'FC Bagger',
    'dt bag':       'DT Bagger',
    'filet':        'Raw (Filet)',
    'raw':          'Raw (Filet)',
    'raw filet':    'Raw (Filet)',
    'buns':         'Primary (Buns)',
    'primary':      'Primary (Buns)',
    'primary buns': 'Primary (Buns)'
  };

  var lower = position.toLowerCase().trim();

  // Exact alias match
  if (aliases[lower]) return aliases[lower];

  // If no alias, return original (preserving case)
  return position.trim();
}

/**
 * Returns true if a trainee has met the minimum hours
 * for EVERY position required in their house.
 */
function isCertificationReady(traineeName, house, positionsMap) {
  if (!house || house === 'TBD') return false;

  var reqSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Position Requirements');
  var reqData  = reqSheet.getDataRange().getValues();

  for (var i = 1; i < reqData.length; i++) {
    var reqHouse    = reqData[i][0];
    var reqPosition = reqData[i][1];
    var minHours    = reqData[i][2];

    if (reqHouse !== house) continue;

    var logged = positionsMap[reqPosition] || 0;
    if (logged < minHours) return false;
  }

  return true;
}

/**
 * Gets all position hours for a given trainee.
 */
function getTraineePositions(traineeName, logSheet) {
  if (logSheet.getLastRow() < 2) return {};

  var data      = logSheet.getDataRange().getValues();
  var positions = {};

  data.slice(1).forEach(function (row) {
    var name = String(row[7]).trim(); // Canonical name
    if (!name) name = String(row[2]).trim();
    if (name !== traineeName) return;

    var position = String(row[3]).trim();
    var hours    = parseFloat(row[4]) || 0;

    if (!positions[position]) positions[position] = 0;
    positions[position] += hours;
  });

  return positions;
}

