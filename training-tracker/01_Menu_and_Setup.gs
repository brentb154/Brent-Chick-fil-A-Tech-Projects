/**
 * ============================================================
 * TRAINING TRACKING SYSTEM - Menu & Setup
 * ============================================================
 * Creates the custom "Training Tools" menu and handles
 * initial sheet setup / data population.
 */

// -- Menu -----------------------------------------------------

function onOpen() {
  var ui = SpreadsheetApp.getUi();

  ui.createMenu('Training Tools')
    .addItem('Generate Training Timeline', 'generateTrainingTimeline')
    .addItem('Load This Week\'s Training', 'loadThisWeeksTraining')
    .addItem('Load Next Week\'s Training', 'loadNextWeeksTraining')
    .addSeparator()
    .addItem('Review Duplicate Names', 'showDeduplicationSidebar')
    .addItem('Check for Duplicates Now', 'checkForDuplicates')
    .addSeparator()
    .addItem('Alert Settings', 'showAlertSettings')
    .addSeparator()
    .addItem('Refresh Dashboard', 'updateDashboard')
    .addItem('Sync Form Data', 'syncFormData')
    .addItem('Manually Certify Trainee', 'manuallyCertifyTrainee')
    .addItem('Setup Monday Auto-Populate', 'setupMondayTrigger')
    .addItem('Initial Setup (Run Once)', 'runInitialSetup')
    .addToUi();
}

// -- Initial Setup --------------------------------------------

/**
 * One-time setup: creates all required sheets, populates
 * Position Requirements and Alert Settings with defaults.
 * Safe to re-run - skips sheets that already exist.
 */
function runInitialSetup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  // 1. Create sheets if missing
  var sheetsToCreate = [
    'Daily Training Log',
    'Position Requirements',
    'Master Dashboard',
    'Certification Log',
    'Name Deduplication',
    'Alert Settings',
    'Training Schedule'
  ];

  sheetsToCreate.forEach(function (name) {
    if (!ss.getSheetByName(name)) {
      ss.insertSheet(name);
    }
  });

  // 2. Populate Position Requirements
  setupPositionRequirements(ss);

  // 3. Set up Daily Training Log headers
  setupDailyTrainingLog(ss);

  // 4. Set up Master Dashboard layout
  setupMasterDashboard(ss);

  // 5. Set up Certification Log headers
  setupCertificationLog(ss);

  // 6. Set up Name Deduplication headers
  setupNameDeduplication(ss);

  // 7. Populate Alert Settings
  setupAlertSettings(ss);

  // 8. Set up Training Schedule headers
  setupTrainingScheduleSheet(ss);

  ui.alert(
    'Setup Complete',
    'All sheets have been created and populated. Next steps:\n\n' +
      '1. Link your Google Form to the "Daily Training Log" sheet\n' +
      '2. Set up installable triggers (see documentation)\n' +
      '3. Configure alert recipients in Training Tools -> Alert Settings',
    ui.ButtonSet.OK
  );
}

// -- Sheet Setup Helpers --------------------------------------

function setupPositionRequirements(ss) {
  var sheet = ss.getSheetByName('Position Requirements');
  if (sheet.getLastRow() > 1) return; // Already has data

  var headers = [['House', 'Position Name', 'Minimum Hours', 'Maximum Hours', 'Target Hours']];

  var data = [
    ['FOH', 'iPOS',        12, 16, 14],
    ['FOH', 'Register/POS', 8, 12, 10],
    ['FOH', 'Cash Cart',    3,  5,  4],
    ['FOH', 'Server',       6,  9,  7.5],
    ['FOH', 'FC Drinks',    6,  8,  7],
    ['FOH', 'Desserts',     4,  6,  5],
    ['FOH', 'DT Drinks',    6,  8,  7],
    ['FOH', 'DT Stuffer',   8, 10,  9],
    ['FOH', 'FC Bagger',    4,  6,  5],
    ['FOH', 'DT Bagger',    4,  6,  5],
    ['FOH', 'Window',       4,  6,  5],
    ['BOH', 'Breading',    12, 16, 14],
    ['BOH', 'Raw (Filet)', 10, 12, 11],
    ['BOH', 'Fries',        8, 10,  9],
    ['BOH', 'Machines',     8, 10,  9],
    ['BOH', 'Dishes',       6,  8,  7],
    ['BOH', 'Prep',         8, 10,  9],
    ['BOH', 'Secondary',    8, 10,  9],
    ['BOH', 'Primary (Buns)', 12, 16, 14],
    ['BOH', 'Truck',        4,  6,  5]
  ];

  sheet.getRange(1, 1, 1, 5).setValues(headers).setFontWeight('bold').setBackground('#4285f4').setFontColor('white');
  sheet.getRange(2, 1, data.length, 5).setValues(data);
  sheet.autoResizeColumns(1, 5);
}

function setupDailyTrainingLog(ss) {
  var sheet = ss.getSheetByName('Daily Training Log');
  if (sheet.getLastRow() > 0 && sheet.getRange('A1').getValue() !== '') return;

  var headers = [['Timestamp', 'Date', 'Trainee Name', 'Position Trained', 'Hours', 'On Track?', 'Notes', 'Canonical Name', 'Synced to Dashboard']];
  sheet.getRange(1, 1, 1, 9).setValues(headers).setFontWeight('bold').setBackground('#4285f4').setFontColor('white');
  sheet.autoResizeColumns(1, 9);
}

function setupMasterDashboard(ss) {
  var sheet = ss.getSheetByName('Master Dashboard');
  if (sheet.getRange('A1').getValue() !== '') return;

  // Title
  sheet.getRange('A1').setValue('TRAINING DASHBOARD').setFontSize(18).setFontWeight('bold');
  sheet.getRange('A2').setValue('Auto-updated from Daily Training Log').setFontColor('#666666');

  // Quick Stats section
  sheet.getRange('A3').setValue('-- QUICK STATS --').setFontWeight('bold').setBackground('#E2EFDA');
  sheet.getRange('A4').setValue('Total Active Trainees:');
  sheet.getRange('A5').setValue('Total Hours This Week:');
  sheet.getRange('A6').setValue('Trainees Ready for Certification:');
  sheet.getRange('A7').setValue('Last Updated:');

  // Active Trainees header
  sheet.getRange('A12').setValue('-- ACTIVE TRAINEES --').setFontWeight('bold').setBackground('#E2EFDA');
  var traineeHeaders = [['Trainee Name', 'House', 'Total Hours', 'Days Since Last', 'Last Position', 'Last Date', 'On Track %', 'Cert Status']];
  sheet.getRange(13, 1, 1, 8).setValues(traineeHeaders).setFontWeight('bold').setBackground('#D9E2F3');

  // Position Progress section (headers written dynamically by updatePositionProgress)
  sheet.getRange('A31').setValue('-- POSITION PROGRESS --').setFontWeight('bold').setBackground('#E2EFDA');

  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 120);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 120);
  sheet.setColumnWidth(5, 130);
  sheet.setColumnWidth(6, 110);
  sheet.setColumnWidth(7, 100);
  sheet.setColumnWidth(8, 130);
}

function setupCertificationLog(ss) {
  var sheet = ss.getSheetByName('Certification Log');
  if (sheet.getLastRow() > 0 && sheet.getRange('A1').getValue() !== '') return;

  var headers = [['Trainee Name', 'House', 'Certification Date', 'Total Training Hours', 'Training Duration (days)', 'Form Response ID', 'Notes']];
  sheet.getRange(1, 1, 1, 7).setValues(headers).setFontWeight('bold').setBackground('#4285f4').setFontColor('white');
  sheet.autoResizeColumns(1, 7);
}

function setupNameDeduplication(ss) {
  var sheet = ss.getSheetByName('Name Deduplication');
  if (sheet.getLastRow() > 0 && sheet.getRange('A1').getValue() !== '') return;

  var headers = [['Suggested Canonical Name', 'Variant Names', 'Entry Count', 'Action', 'Status']];
  sheet.getRange(1, 1, 1, 5).setValues(headers).setFontWeight('bold').setBackground('#4285f4').setFontColor('white');
  sheet.autoResizeColumns(1, 5);
}

function setupAlertSettings(ss) {
  var sheet = ss.getSheetByName('Alert Settings');
  if (sheet.getLastRow() > 1) return;

  var headers = [['Alert Type', 'Enabled', 'Recipient 1 Email', 'Recipient 2 Email', 'Recipient 3 Email', 'Send Time']];

  var data = [
    ['Daily Training Log Reminder',  true, '', '', '', '20:00'],
    ['Trainee Inactive 3+ Days',     true, '', '', '', '06:00'],
    ['Position Completion Milestone', true, '', '', '', ''],
    ['Trainee Ready for Certification', true, '', '', '', ''],
    ['Duplicate Name Detected',       true, '', '', '', '']
  ];

  sheet.getRange(1, 1, 1, 6).setValues(headers).setFontWeight('bold').setBackground('#4285f4').setFontColor('white');
  sheet.getRange(2, 1, data.length, 6).setValues(data);

  // Add data validation for Enabled column (TRUE/FALSE checkboxes)
  var enabledRange = sheet.getRange(2, 2, data.length, 1);
  enabledRange.insertCheckboxes();

  sheet.autoResizeColumns(1, 6);
}
