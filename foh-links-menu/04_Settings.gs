// ============================================================
// FOH Links Menu — Settings
// ============================================================

var SETTINGS_DEFAULTS = {
  'Donation Subject':  'Chick-fil-A Cockrell Hill — Donation Request',
  'Donation From Name': 'Chick-fil-A Cockrell Hill',
  'Donation Message':  'Hi {NAME},\n\nThank you for reaching out to Chick-fil-A Cockrell Hill regarding a donation! We take serving our community seriously, and we love being able to support organizations and causes that make a difference.\n\nTo help us process your request, please fill out our giving request form using the link below:\n\n{FORM_LINK}\n\nOnce we receive your completed form, our team will review it and follow up with you.\n\nThank you for thinking of us!\n\nWarm regards,\nChick-fil-A Cockrell Hill',
  'Alert Email':       '',
  'Reset Day':         'Sunday',
  'Reset Hour':        '9 AM'
};

var SETTINGS_DESCRIPTIONS = {
  'Donation Subject':  'Default email subject line for donation requests',
  'Donation From Name': 'Sender display name on donation emails',
  'Donation Message':  'Default email body. Use {NAME} for recipient name, {FORM_LINK} for the giving form URL',
  'Alert Email':       'Who gets emailed if the Sunday reset fails (leave blank to use your own email)',
  'Reset Day':         'Day of week for automatic sheet reset',
  'Reset Hour':        'Time of day for automatic sheet reset'
};

/**
 * Reads a setting from the Settings sheet. Falls back to default if not found.
 */
function getSetting(name) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  if (!sheet) return SETTINGS_DEFAULTS[name] || '';

  var data = sheet.getDataRange().getValues();
  for (var r = 1; r < data.length; r++) {
    if (String(data[r][0]).trim() === name) {
      var val = data[r][1];
      return (val === '' || val === null || val === undefined) ? (SETTINGS_DEFAULTS[name] || '') : String(val);
    }
  }
  return SETTINGS_DEFAULTS[name] || '';
}

/**
 * Returns all settings as an object. Used by the donation dialog.
 */
function getAllSettings() {
  var result = {};
  Object.keys(SETTINGS_DEFAULTS).forEach(function(key) {
    result[key] = getSetting(key);
  });
  return result;
}

/**
 * Creates the Settings sheet with defaults and validation. Safe to re-run.
 */
function setupSettingsSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Settings');

  if (!sheet) {
    sheet = ss.insertSheet('Settings');
  }

  // Check if already set up
  if (sheet.getRange('A1').getValue() === 'Setting') return sheet;

  // Headers
  var headers = ['Setting', 'Value', 'Description'];
  sheet.getRange(1, 1, 1, 3).setValues([headers]);
  sheet.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');

  // Write setting rows
  var settingNames = Object.keys(SETTINGS_DEFAULTS);
  var rows = settingNames.map(function(name) {
    return [name, SETTINGS_DEFAULTS[name], SETTINGS_DESCRIPTIONS[name]];
  });
  sheet.getRange(2, 1, rows.length, 3).setValues(rows);

  // Column widths
  sheet.setColumnWidth(1, 180); // Setting
  sheet.setColumnWidth(2, 400); // Value
  sheet.setColumnWidth(3, 400); // Description

  // Lock Setting and Description columns (light gray background to signal "don't edit")
  sheet.getRange(2, 1, rows.length, 1).setBackground('#f3f3f3').setFontWeight('bold');
  sheet.getRange(2, 3, rows.length, 1).setBackground('#f3f3f3').setFontColor('#5f6368').setFontStyle('italic');

  // Wrap text on the message row
  var messageRow = settingNames.indexOf('Donation Message') + 2;
  sheet.getRange(messageRow, 2).setWrap(true);
  sheet.setRowHeight(messageRow, 120);

  // --- Data Validation ---

  // Reset Day dropdown
  var dayRow = settingNames.indexOf('Reset Day') + 2;
  var dayRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(dayRow, 2).setDataValidation(dayRule);

  // Reset Hour dropdown
  var hourRow = settingNames.indexOf('Reset Hour') + 2;
  var hours = ['12 AM', '1 AM', '2 AM', '3 AM', '4 AM', '5 AM', '6 AM', '7 AM', '8 AM', '9 AM', '10 AM', '11 AM',
               '12 PM', '1 PM', '2 PM', '3 PM', '4 PM', '5 PM', '6 PM', '7 PM', '8 PM', '9 PM', '10 PM', '11 PM'];
  var hourRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(hours, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(hourRow, 2).setDataValidation(hourRule);

  // Freeze header
  sheet.setFrozenRows(1);

  // Protect Setting and Description columns
  var protection = sheet.getRange(2, 1, rows.length, 1).protect().setDescription('Setting names — do not edit');
  protection.setWarningOnly(true);
  var descProtection = sheet.getRange(2, 3, rows.length, 1).protect().setDescription('Descriptions — do not edit');
  descProtection.setWarningOnly(true);

  SpreadsheetApp.flush();
  return sheet;
}

/**
 * Converts "9 AM" style hour string to 24hr integer (0-23).
 */
function parseHourSetting_(hourStr) {
  hourStr = String(hourStr).trim().toUpperCase();
  var match = hourStr.match(/^(\d{1,2})\s*(AM|PM)$/);
  if (!match) return 9; // fallback
  var hour = parseInt(match[1]);
  var ampm = match[2];
  if (ampm === 'AM') {
    return hour === 12 ? 0 : hour;
  } else {
    return hour === 12 ? 12 : hour + 12;
  }
}

/**
 * Converts day name string to ScriptApp.WeekDay constant.
 */
function parseDaySetting_(dayStr) {
  var map = {
    'SUNDAY': ScriptApp.WeekDay.SUNDAY,
    'MONDAY': ScriptApp.WeekDay.MONDAY,
    'TUESDAY': ScriptApp.WeekDay.TUESDAY,
    'WEDNESDAY': ScriptApp.WeekDay.WEDNESDAY,
    'THURSDAY': ScriptApp.WeekDay.THURSDAY,
    'FRIDAY': ScriptApp.WeekDay.FRIDAY,
    'SATURDAY': ScriptApp.WeekDay.SATURDAY
  };
  return map[String(dayStr).trim().toUpperCase()] || ScriptApp.WeekDay.SUNDAY;
}
