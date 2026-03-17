var SETTINGS_SHEET_NAME = 'Settings';

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Reimbursement Reports')
    .addItem('Send Report Now (Previous Month)', 'sendMonthlyReport')
    .addItem('Send Test Report (Current Month)', 'sendTestReport')
    .addSeparator()
    .addItem('Setup Monthly Trigger', 'setupMonthlyTrigger')
    .addItem('Remove All Triggers', 'removeAllTriggers')
    .addSeparator()
    .addItem('Create/Reset Settings Sheet', 'createSettingsSheet')
    .addToUi();
}

// ── Settings Sheet ──────────────────────────────────────────────────

function createSettingsSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SETTINGS_SHEET_NAME);

  if (sheet) {
    sheet.clear();
  } else {
    sheet = ss.insertSheet(SETTINGS_SHEET_NAME);
  }

  sheet.getRange(1, 1, 6, 3).setValues([
    ['Setting', 'Value', 'Notes'],
    ['Email Recipients', '', 'Comma-separated email addresses'],
    ['Report Day of Month', 1, 'Day of month to send the report (1-28)'],
    ['Email Subject Template', 'Monthly Mileage Reimbursement Report - {month} {year}', '{month} and {year} are auto-replaced'],
    ['Data Sheet Name', 'Form Responses 1', 'Name of the tab with form responses'],
    ['Date Column', 'Date of trip', '"Date of trip" or "Timestamp"']
  ]);

  sheet.getRange(1, 1, 1, 3)
    .setFontWeight('bold')
    .setBackground('#4285f4')
    .setFontColor('#ffffff');

  sheet.getRange(2, 1, 5, 1).setFontWeight('bold');
  sheet.getRange(2, 3, 5, 1).setFontColor('#888888').setFontStyle('italic');

  sheet.setColumnWidth(1, 220);
  sheet.setColumnWidth(2, 420);
  sheet.setColumnWidth(3, 350);

  sheet.getRange(3, 2).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireNumberBetween(1, 28)
      .setAllowInvalid(false)
      .build()
  );

  sheet.getRange(6, 2).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['Date of trip', 'Timestamp'])
      .setAllowInvalid(false)
      .build()
  );

  SpreadsheetApp.getUi().alert(
    'Settings sheet created!\n\nPlease enter your email address(es) in the Email Recipients row.'
  );
}

function getSettings_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SETTINGS_SHEET_NAME);

  if (!sheet) {
    throw new Error('Settings sheet not found. Use menu: Reimbursement Reports > Create/Reset Settings Sheet');
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    throw new Error('Settings sheet is empty. Please re-create it from the menu.');
  }

  var data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  var map = {};
  for (var i = 0; i < data.length; i++) {
    var key = String(data[i][0]).trim();
    if (key) map[key] = data[i][1];
  }

  return {
    recipients: String(map['Email Recipients'] || '').trim(),
    reportDay: parseInt(map['Report Day of Month'], 10) || 1,
    subjectTemplate: String(map['Email Subject Template'] || 'Monthly Mileage Reimbursement Report - {month} {year}'),
    dataSheetName: String(map['Data Sheet Name'] || 'Form Responses 1').trim(),
    dateColumn: String(map['Date Column'] || 'Date of trip').trim()
  };
}

// ── Trigger Management ──────────────────────────────────────────────

function setupMonthlyTrigger() {
  clearTriggers_();

  var settings = getSettings_();
  var day = settings.reportDay;

  ScriptApp.newTrigger('sendMonthlyReport')
    .timeBased()
    .onMonthDay(day)
    .atHour(7)
    .create();

  SpreadsheetApp.getUi().alert(
    'Done! The report will automatically send on day ' + day + ' of each month around 7-8 AM.'
  );
}

function removeAllTriggers() {
  clearTriggers_();
  SpreadsheetApp.getUi().alert('All triggers have been removed.');
}

function clearTriggers_() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
}

// ── Report Entry Points ─────────────────────────────────────────────

function sendMonthlyReport() {
  try {
    var today = new Date();
    var month = today.getMonth() - 1;
    var year = today.getFullYear();
    if (month < 0) { month = 11; year--; }
    sendReport_(month, year);
  } catch (e) {
    Logger.log('sendMonthlyReport error: ' + e.message);
    notifyUser_('Error: ' + e.message);
  }
}

function sendTestReport() {
  try {
    var today = new Date();
    sendReport_(today.getMonth(), today.getFullYear());
  } catch (e) {
    Logger.log('sendTestReport error: ' + e.message);
    notifyUser_('Error: ' + e.message);
  }
}

// ── Core Report Logic ───────────────────────────────────────────────

var MONTH_NAMES = [
  'January','February','March','April','May','June',
  'July','August','September','October','November','December'
];

function sendReport_(month, year) {
  var monthName = MONTH_NAMES[month];
  var settings = getSettings_();

  if (!settings.recipients) {
    throw new Error('No email recipients set. Please update the Settings sheet.');
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataSheet = ss.getSheetByName(settings.dataSheetName);
  if (!dataSheet) {
    throw new Error('Sheet "' + settings.dataSheetName + '" not found. Check your Settings sheet.');
  }

  var allData = dataSheet.getDataRange().getValues();
  if (allData.length < 2) {
    sendNoActivityEmail_(settings, monthName, year);
    return;
  }

  var headers = allData[0];
  var colIdx = findColumns_(headers, settings.dateColumn);
  var rows = filterMonth_(allData, colIdx.date, month, year);

  if (rows.length === 0) {
    sendNoActivityEmail_(settings, monthName, year);
    return;
  }

  var agg = aggregate_(rows, colIdx);

  var images = {};
  try {
    images = buildCharts_(agg, monthName, year);
  } catch (e) {
    Logger.log('Chart generation failed, sending without charts: ' + e.message);
  }

  var subject = settings.subjectTemplate
    .replace('{month}', monthName)
    .replace('{year}', String(year));

  var html = buildHtml_(monthName, year, agg, images);

  MailApp.sendEmail({
    to: settings.recipients,
    subject: subject,
    htmlBody: html,
    inlineImages: images
  });

  Logger.log('Report sent for ' + monthName + ' ' + year + ' to ' + settings.recipients);
  notifyUser_('Report sent for ' + monthName + ' ' + year + ' to ' + settings.recipients);
}

// ── Column Resolution ───────────────────────────────────────────────

function findColumns_(headers, dateColumnSetting) {
  var normalize = function(s) {
    return String(s).toLowerCase().trim().replace(/[?:]+$/, '');
  };

  var normalizedHeaders = [];
  for (var i = 0; i < headers.length; i++) {
    normalizedHeaders.push(normalize(headers[i]));
  }

  var find = function(label) {
    var target = normalize(label);
    for (var i = 0; i < normalizedHeaders.length; i++) {
      if (normalizedHeaders[i] === target) return i;
    }
    return -1;
  };

  var dateCol = (dateColumnSetting === 'Timestamp')
    ? find('Timestamp')
    : find('Date of trip');

  var amount  = find('Cash total for reimbursement');
  var person  = find("Team Member's name");
  if (person === -1) person = find("Team Member\u2019s name");
  var purpose = find('What was the purpose of the trip');
  var miles   = find('What was the total miles driven round trip');

  if (dateCol === -1)  throw new Error('Cannot find the date column in your headers.');
  if (amount === -1)   throw new Error('Cannot find the "Cash total for reimbursement" column.');
  if (person === -1)   throw new Error('Cannot find the "Team Member\'s name" column.');
  if (purpose === -1)  throw new Error('Cannot find the "What was the purpose of the trip" column.');

  return { date: dateCol, amount: amount, person: person, purpose: purpose, miles: miles };
}

// ── Data Filtering & Aggregation ────────────────────────────────────

function filterMonth_(allData, dateColIdx, month, year) {
  var rows = [];
  for (var i = 1; i < allData.length; i++) {
    var raw = allData[i][dateColIdx];
    if (!raw) continue;
    var d = (raw instanceof Date) ? raw : new Date(raw);
    if (isNaN(d.getTime())) continue;
    if (d.getMonth() === month && d.getFullYear() === year) {
      rows.push(allData[i]);
    }
  }
  return rows;
}

function aggregate_(rows, colIdx) {
  var byPerson = {};
  var byPurpose = {};
  var grandTotal = 0;
  var totalMiles = 0;

  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var person  = String(row[colIdx.person]  || 'Unknown').trim();
    var purpose = String(row[colIdx.purpose] || 'Unknown').trim();
    var amount  = parseFloat(row[colIdx.amount]) || 0;
    var miles   = (colIdx.miles >= 0) ? (parseFloat(row[colIdx.miles]) || 0) : 0;

    if (!byPerson[person]) byPerson[person] = { total: 0, trips: 0, miles: 0 };
    byPerson[person].total += amount;
    byPerson[person].trips += 1;
    byPerson[person].miles += miles;

    if (!byPurpose[purpose]) byPurpose[purpose] = { total: 0, trips: 0 };
    byPurpose[purpose].total += amount;
    byPurpose[purpose].trips += 1;

    grandTotal += amount;
    totalMiles += miles;
  }

  return {
    byPerson: byPerson,
    byPurpose: byPurpose,
    grandTotal: grandTotal,
    totalMiles: totalMiles,
    totalTrips: rows.length
  };
}

// ── Chart Generation ────────────────────────────────────────────────

function buildCharts_(agg, monthName, year) {
  var images = {};

  var personKeys = Object.keys(agg.byPerson).sort();
  if (personKeys.length > 0) {
    var dt = Charts.newDataTable()
      .addColumn(Charts.ColumnType.STRING, 'Member')
      .addColumn(Charts.ColumnType.NUMBER, 'Amount');
    for (var i = 0; i < personKeys.length; i++) {
      dt.addRow([personKeys[i], round2_(agg.byPerson[personKeys[i]].total)]);
    }
    var chart = Charts.newPieChart()
      .setDataTable(dt.build())
      .setTitle('Reimbursements by Team Member - ' + monthName + ' ' + year)
      .setDimension(500, 320)
      .setLegendPosition(Charts.Position.RIGHT)
      .build();
    images.personChart = chart.getAs('image/png');
  }

  var purposeKeys = Object.keys(agg.byPurpose).sort();
  if (purposeKeys.length > 0) {
    var dt2 = Charts.newDataTable()
      .addColumn(Charts.ColumnType.STRING, 'Purpose')
      .addColumn(Charts.ColumnType.NUMBER, 'Amount');
    for (var i = 0; i < purposeKeys.length; i++) {
      dt2.addRow([purposeKeys[i], round2_(agg.byPurpose[purposeKeys[i]].total)]);
    }
    var chart2 = Charts.newPieChart()
      .setDataTable(dt2.build())
      .setTitle('Reimbursements by Purpose - ' + monthName + ' ' + year)
      .setDimension(500, 320)
      .setLegendPosition(Charts.Position.RIGHT)
      .build();
    images.purposeChart = chart2.getAs('image/png');
  }

  return images;
}

// ── HTML Email Builder ──────────────────────────────────────────────

function buildHtml_(monthName, year, agg, images) {
  var h = '';

  h += '<div style="font-family:Arial,Helvetica,sans-serif;max-width:680px;margin:0 auto;color:#333;">';

  h += '<div style="background:#4285f4;color:#fff;padding:20px 24px;border-radius:8px 8px 0 0;">';
  h += '<h1 style="margin:0;font-size:22px;">Mileage Reimbursement Report</h1>';
  h += '<p style="margin:4px 0 0;font-size:14px;opacity:0.9;">' + monthName + ' ' + year + '</p>';
  h += '</div>';

  h += '<table style="width:100%;background:#f8f9fa;border-bottom:1px solid #e0e0e0;" cellpadding="0" cellspacing="0"><tr>';
  h += summaryCell_('$' + agg.grandTotal.toFixed(2), 'Total Reimbursed');
  h += summaryCell_(String(agg.totalTrips), 'Total Trips');
  h += summaryCell_(agg.totalMiles.toFixed(1), 'Total Miles');
  h += '</tr></table>';

  h += '<div style="padding:24px;">';

  h += '<h2 style="font-size:16px;margin-top:0;">By Team Member</h2>';
  h += '<table style="width:100%;border-collapse:collapse;margin-bottom:16px;">';
  h += tableHeader_(['Team Member', 'Trips', 'Miles', 'Reimbursed']);

  var personKeys = Object.keys(agg.byPerson).sort();
  for (var i = 0; i < personKeys.length; i++) {
    var p = agg.byPerson[personKeys[i]];
    h += tableRow_([personKeys[i], p.trips, p.miles.toFixed(1), '$' + p.total.toFixed(2)], i);
  }
  h += totalRow_(['Total', agg.totalTrips, agg.totalMiles.toFixed(1), '$' + agg.grandTotal.toFixed(2)]);
  h += '</table>';

  if (images.personChart) {
    h += '<div style="text-align:center;margin-bottom:28px;"><img src="cid:personChart" style="max-width:100%;"></div>';
  }

  h += '<h2 style="font-size:16px;">By Trip Purpose</h2>';
  h += '<table style="width:100%;border-collapse:collapse;margin-bottom:16px;">';
  h += tableHeader_(['Purpose', 'Trips', 'Reimbursed']);

  var purposeKeys = Object.keys(agg.byPurpose).sort();
  for (var i = 0; i < purposeKeys.length; i++) {
    var pp = agg.byPurpose[purposeKeys[i]];
    h += tableRow_([purposeKeys[i], pp.trips, '$' + pp.total.toFixed(2)], i);
  }
  h += totalRow_(['Total', agg.totalTrips, '$' + agg.grandTotal.toFixed(2)]);
  h += '</table>';

  if (images.purposeChart) {
    h += '<div style="text-align:center;margin-bottom:28px;"><img src="cid:purposeChart" style="max-width:100%;"></div>';
  }

  h += '<p style="font-size:11px;color:#999;margin-top:24px;border-top:1px solid #eee;padding-top:12px;">';
  h += 'Automatically generated from the mileage reimbursement spreadsheet.</p>';
  h += '</div></div>';

  return h;
}

// ── HTML Helpers ────────────────────────────────────────────────────

function summaryCell_(value, label) {
  return '<td style="text-align:center;padding:16px 8px;">'
    + '<div style="font-size:24px;font-weight:bold;color:#4285f4;">' + value + '</div>'
    + '<div style="font-size:12px;color:#666;">' + label + '</div></td>';
}

function tableHeader_(cols) {
  var row = '<tr style="background:#4285f4;color:#fff;">';
  for (var i = 0; i < cols.length; i++) {
    var align = (i === 0) ? 'left' : 'right';
    row += '<th style="padding:10px 12px;text-align:' + align + ';">' + cols[i] + '</th>';
  }
  return row + '</tr>';
}

function tableRow_(cells, rowIndex) {
  var bg = (rowIndex % 2 === 0) ? '#ffffff' : '#f8f9fa';
  var row = '<tr style="background:' + bg + ';">';
  for (var i = 0; i < cells.length; i++) {
    var align = (i === 0) ? 'left' : 'right';
    var bold = (i === cells.length - 1) ? 'font-weight:bold;' : '';
    row += '<td style="padding:8px 12px;border-bottom:1px solid #eee;text-align:' + align + ';' + bold + '">' + cells[i] + '</td>';
  }
  return row + '</tr>';
}

function totalRow_(cells) {
  var row = '<tr style="background:#e8f0fe;font-weight:bold;">';
  for (var i = 0; i < cells.length; i++) {
    var align = (i === 0) ? 'left' : 'right';
    row += '<td style="padding:10px 12px;text-align:' + align + ';">' + cells[i] + '</td>';
  }
  return row + '</tr>';
}

// ── No-Activity Email ───────────────────────────────────────────────

function sendNoActivityEmail_(settings, monthName, year) {
  var subject = settings.subjectTemplate
    .replace('{month}', monthName)
    .replace('{year}', String(year));

  var html = '<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;">'
    + '<div style="background:#4285f4;color:#fff;padding:20px 24px;border-radius:8px 8px 0 0;">'
    + '<h1 style="margin:0;font-size:22px;">Mileage Reimbursement Report</h1>'
    + '<p style="margin:4px 0 0;font-size:14px;opacity:0.9;">' + monthName + ' ' + year + '</p>'
    + '</div>'
    + '<div style="padding:24px;background:#f8f9fa;border-radius:0 0 8px 8px;text-align:center;">'
    + '<p style="font-size:16px;color:#666;">No mileage reimbursement trips were recorded for ' + monthName + ' ' + year + '.</p>'
    + '</div></div>';

  MailApp.sendEmail({
    to: settings.recipients,
    subject: subject + ' (No Activity)',
    htmlBody: html
  });

  Logger.log('No-activity email sent for ' + monthName + ' ' + year);
  notifyUser_('No trips found for ' + monthName + ' ' + year + '. A no-activity notice was sent.');
}

// ── Utilities ───────────────────────────────────────────────────────

function round2_(n) {
  return Math.round(n * 100) / 100;
}

function notifyUser_(message) {
  try {
    SpreadsheetApp.getUi().alert(message);
  } catch (e) {
    Logger.log(message);
  }
}
