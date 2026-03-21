/**
 * ============================================================
 * TRAINING TRACKING SYSTEM - Diagnostic Utility
 * ============================================================
 * Debug tool for troubleshooting Training Needs layout detection
 * and Training Schedule population issues.
 *
 * HOW TO USE:
 *   1. Add this file to your Apps Script project
 *   2. Run runTrainingDiagnostic() from the editor (play button)
 *   3. A scrollable report dialog will appear showing:
 *      - Training Schedule data (first 10 rows)
 *      - Training Needs sheet raw content
 *      - Layout detection results (which days/dayparts were found)
 *      - Date matching analysis (scheduled weeks vs. current week)
 *
 * This file is optional and can be removed once setup is verified.
 */

function runTrainingDiagnostic() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tz = Session.getScriptTimeZone();
  var report = [];

  // ============================================
  // 1. CHECK TRAINING SCHEDULE DATA
  // ============================================
  report.push('=== TRAINING SCHEDULE ===');
  var schedSheet = ss.getSheetByName('Training Schedule');
  if (!schedSheet) {
    report.push('NOT FOUND - Training Schedule sheet does not exist!');
  } else {
    var lastRow = schedSheet.getLastRow();
    report.push('Rows: ' + lastRow);
    if (lastRow >= 1) {
      var headers = schedSheet.getRange(1, 1, 1, schedSheet.getLastColumn()).getValues()[0];
      report.push('Headers: ' + headers.join(' | '));
    }
    if (lastRow >= 2) {
      var data = schedSheet.getRange(2, 1, Math.min(lastRow - 1, 10), schedSheet.getLastColumn()).getValues();
      data.forEach(function(row, i) {
        report.push('  Row ' + (i+2) + ': ' + row.map(function(v) {
          if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM-dd EEE');
          return String(v);
        }).join(' | '));
      });
    }
  }

  // ============================================
  // 2. CHECK TRAINING NEEDS SHEET EXISTS
  // ============================================
  report.push('');
  report.push('=== TRAINING NEEDS SHEET ===');
  var needsSheet = ss.getSheetByName('Training Needs');
  if (!needsSheet) {
    report.push('NOT FOUND!');
    showDiagReport(report);
    return;
  }
  report.push('Found. Last row: ' + needsSheet.getLastRow() + ', Last col: ' + needsSheet.getLastColumn());

  // ============================================
  // 3. RAW SCAN - show first 35 rows of key columns
  // ============================================
  report.push('');
  report.push('=== RAW CONTENT (rows 1-35, cols A,B,D,G,H,J,K,L) ===');
  var numRows = Math.min(needsSheet.getLastRow(), 35);
  if (numRows > 0) {
    var raw = needsSheet.getRange(1, 1, numRows, 14).getValues();
    for (var r = 0; r < raw.length; r++) {
      var a = String(raw[r][0]).substring(0, 15).trim();  // A
      var b = String(raw[r][1]).substring(0, 10).trim();  // B
      var d = String(raw[r][3]).substring(0, 12).trim();  // D
      var g = String(raw[r][6]).substring(0, 15).trim();  // G
      var h = String(raw[r][7]).substring(0, 10).trim();  // H
      var j = String(raw[r][9]).substring(0, 15).trim();  // J
      var k = String(raw[r][10]).substring(0, 10).trim(); // K
      var l = String(raw[r][11]).substring(0, 12).trim(); // L

      if (a || b || d || g || h || j || k || l) {
        report.push('Row ' + (r+1).toString().padStart(2,' ') + ': A=[' + a + '] B=[' + b + '] D=[' + d + '] G=[' + g + '] H=[' + h + '] J=[' + j + '] K=[' + k + '] L=[' + l + ']');
      }
    }
  }

  // ============================================
  // 4. RUN LAYOUT DETECTION
  // ============================================
  report.push('');
  report.push('=== LAYOUT DETECTION ===');

  var allData = needsSheet.getRange(1, 1, Math.min(needsSheet.getLastRow(), 200), 14).getValues();
  var days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  var dayRows = {};

  // Find day markers
  for (var r = 0; r < allData.length; r++) {
    var rowText = '';
    for (var c = 0; c < allData[r].length; c++) {
      rowText += ' ' + String(allData[r][c]);
    }
    var rowLower = rowText.toLowerCase();

    for (var di = 0; di < days.length; di++) {
      if (rowLower.indexOf(days[di].toLowerCase()) > -1 && !dayRows[days[di]]) {
        dayRows[days[di]] = r + 1;
        report.push('Found "' + days[di] + '" at row ' + (r+1) + ' | raw text: ' + rowText.trim().substring(0, 80));
      }
    }
  }

  if (Object.keys(dayRows).length === 0) {
    report.push('NO DAY NAMES FOUND! The sheet may not contain Monday/Tuesday/etc text.');
  }

  // Find FOH Training headers per day
  Object.keys(dayRows).forEach(function(dayName) {
    var dayRow = dayRows[dayName];
    report.push('');
    report.push('--- ' + dayName + ' (starts row ' + dayRow + ') ---');

    var fohRows = [];
    var searchEnd = Math.min(dayRow + 19, allData.length);
    for (var r = dayRow - 1; r < searchEnd; r++) {
      var cellA = String(allData[r][0]).trim();
      var cellJ = String(allData[r][9] || '').trim();

      if (cellA.toLowerCase().indexOf('training') > -1 ||
          cellA.toLowerCase().indexOf('foh') > -1 ||
          cellA.toLowerCase().indexOf('boh') > -1) {
        fohRows.push(r + 1);
        report.push('  FOH header at row ' + (r+1) + ' col A: "' + cellA + '"');
      }
      if (cellJ.toLowerCase().indexOf('training') > -1 ||
          cellJ.toLowerCase().indexOf('foh') > -1 ||
          cellJ.toLowerCase().indexOf('boh') > -1) {
        report.push('  FOH header at row ' + (r+1) + ' col J: "' + cellJ + '"');
      }
    }

    if (fohRows.length === 0) {
      report.push('  NO FOH TRAINING HEADERS FOUND within 20 rows of day marker!');
    }

    // Detect columns for each header row
    fohRows.forEach(function(hr, idx) {
      var headerVals = needsSheet.getRange(hr, 1, 1, 14).getValues()[0];
      report.push('  Header row ' + hr + ' full content:');
      for (var c = 0; c < headerVals.length; c++) {
        var v = String(headerVals[c]).trim();
        if (v) {
          report.push('    Col ' + String.fromCharCode(65 + c) + ' = "' + v + '"');
        }
      }
    });
  });

  // ============================================
  // 5. DATE MATCHING CHECK
  // ============================================
  report.push('');
  report.push('=== DATE MATCHING ===');
  var today = new Date();
  var dow = today.getDay();
  var thisMonday = new Date(today);
  thisMonday.setDate(today.getDate() - (dow === 0 ? 6 : dow - 1));
  thisMonday.setHours(0, 0, 0, 0);
  report.push('Today: ' + Utilities.formatDate(today, tz, 'yyyy-MM-dd EEE'));
  report.push('This week Monday: ' + Utilities.formatDate(thisMonday, tz, 'yyyy-MM-dd'));

  if (schedSheet && schedSheet.getLastRow() >= 2) {
    var schedData = schedSheet.getRange(2, 1, schedSheet.getLastRow() - 1, schedSheet.getLastColumn()).getValues();
    var uniqueWeeks = {};
    schedData.forEach(function(row) {
      var d = new Date(row[1] || row[2]); // could be col B or C depending on format
      if (!isNaN(d.getTime())) {
        var wk = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
        var weekday = d.getDay();
        var mon = new Date(d);
        mon.setDate(d.getDate() - (weekday === 0 ? 6 : weekday - 1));
        var monStr = Utilities.formatDate(mon, tz, 'yyyy-MM-dd');
        uniqueWeeks[monStr] = (uniqueWeeks[monStr] || 0) + 1;
      }
    });
    report.push('Weeks with scheduled training:');
    Object.keys(uniqueWeeks).sort().forEach(function(wk) {
      report.push('  Week of ' + wk + ': ' + uniqueWeeks[wk] + ' entries');
    });
  }

  showDiagReport(report);
}

function showDiagReport(lines) {
  var text = lines.join('\n');
  Logger.log(text);

  // Show in a scrollable dialog since alert truncates
  var html = HtmlService
    .createHtmlOutput('<pre style="font-size:12px;font-family:monospace;white-space:pre-wrap;max-height:500px;overflow:auto;">' +
                      text.replace(/</g, '&lt;').replace(/>/g, '&gt;') +
                      '</pre>')
    .setWidth(700)
    .setHeight(550);
  SpreadsheetApp.getUi().showModalDialog(html, 'Training System Diagnostic');
}
