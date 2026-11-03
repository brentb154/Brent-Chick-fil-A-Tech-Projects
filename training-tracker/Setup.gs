// -- FIRST-TIME SETUP -----------------------------------------
// A new store opens its copy of the sheet -> "🎓 Training Setup" menu -> Start Setup.
// The wizard builds every tab, creates the Daily Training Log form, links it to this
// sheet, fills in alert emails and turns on the automatic triggers.
// Safe to re-run: reuses the same form and refreshes its position list.

var SETUP_FORM_PROP = 'TRAINING_FORM_ID';

function showSetupWizard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Building your tabs. This takes about 20 seconds the first time.', 'Training Setup', 20);
  buildTrainingTabs_(ss);
  var html = HtmlService.createHtmlOutputFromFile('SetupWizard').setWidth(520).setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Training Setup');
}

// Same steps as runInitialSetup(), without its "link your form by hand" popup.
function buildTrainingTabs_(ss) {
  ['Daily Training Log', 'Position Requirements', 'Master Dashboard', 'Certification Log',
   'Name Deduplication', 'Alert Settings', 'Training Schedule', 'Training Needs'].forEach(function (name) {
    if (!ss.getSheetByName(name)) ss.insertSheet(name);
  });
  setupPositionRequirements(ss);
  setupDailyTrainingLog(ss);
  setupMasterDashboard(ss);
  setupCertificationLog(ss);
  setupNameDeduplication(ss);
  setupAlertSettings(ss);
  setupTrainingScheduleSheet(ss);
  setupTrainingNeedsSheet(ss);
  SpreadsheetApp.flush();
}

// Current answers so a re-run shows what's already there.
function getSetupStatus() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var form = getSetupForm_(ss);
  var emails = [];
  var alertSheet = ss.getSheetByName('Alert Settings');
  if (alertSheet && alertSheet.getLastRow() > 1) {
    emails = alertSheet.getRange(2, 3, 1, 3).getValues()[0].filter(function (v) { return String(v).trim(); });
  }
  return {
    storeName: form ? form.getTitle().replace(/\s*Daily Training Log$/, '') : '',
    emails: emails.join(', '),
    formUrl: form ? form.getPublishedUrl() : ''
  };
}

function saveSetupAnswers(answers) {
  answers = answers || {};
  var storeName = String(answers.storeName || '').trim();
  var emails = String(answers.emails || '').split(/[,;\s]+/).filter(function (v) { return v.indexOf('@') > -1; });
  if (!storeName) return { success: false, message: 'Please enter your store name.' };
  if (!emails.length) return { success: false, message: 'Please enter at least one email for alerts.' };
  if (emails.length > 3) return { success: false, message: 'Please enter up to 3 emails.' };

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    buildTrainingTabs_(ss);

    // Alert emails go on every alert row
    var alertSheet = ss.getSheetByName('Alert Settings');
    var rows = alertSheet.getLastRow() - 1;
    var recips = [emails[0] || '', emails[1] || '', emails[2] || ''];
    var block = [];
    for (var i = 0; i < rows; i++) block.push(recips);
    alertSheet.getRange(2, 3, rows, 3).setValues(block);

    var form = buildTrainingForm_(ss, storeName);
    setupTrainingTriggers_(ss);
    SpreadsheetApp.flush();
    return { success: true, formUrl: form.getPublishedUrl(), editUrl: form.getEditUrl() };
  } finally {
    lock.releaseLock();
  }
}

// The form saved for THIS sheet, or null (a copied sheet may carry another sheet's form ID).
function getSetupForm_(ss) {
  var id = PropertiesService.getScriptProperties().getProperty(SETUP_FORM_PROP);
  if (!id) return null;
  try {
    var form = FormApp.openById(id);
    return form.getDestinationId() === ss.getId() ? form : null;
  } catch (e) {
    return null;
  }
}

// Question order must match onFormSubmit: Date, Name, Position, Hours, On Track?, Notes
function buildTrainingForm_(ss, storeName) {
  var positions = ss.getSheetByName('Position Requirements').getDataRange().getValues().slice(1)
    .map(function (r) { return String(r[1]).trim(); })
    .filter(function (p) { return p; });

  var form = getSetupForm_(ss);
  if (form) {
    form.getItems(FormApp.ItemType.LIST).forEach(function (item) {
      if (item.getTitle() === 'Position Trained') item.asListItem().setChoiceValues(positions);
    });
  } else {
    form = FormApp.create(storeName + ' Daily Training Log');
    form.setDescription('Trainers: fill this out after every training shift.');
    form.addDateItem().setTitle('Date').setRequired(true);
    form.addTextItem().setTitle('Trainee Name').setHelpText('First and last name').setRequired(true);
    form.addListItem().setTitle('Position Trained').setChoiceValues(positions).setRequired(true);
    form.addTextItem().setTitle('Hours').setRequired(true)
      .setValidation(FormApp.createTextValidation().requireNumber().build());
    form.addMultipleChoiceItem().setTitle('On Track?').setChoiceValues(['Yes', 'No']).setRequired(true);
    form.addParagraphTextItem().setTitle('Notes');
    form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
    PropertiesService.getScriptProperties().setProperty(SETUP_FORM_PROP, form.getId());
  }
  form.setTitle(storeName + ' Daily Training Log');
  return form;
}

// Form submit + daily/weekly triggers. Skips any that already exist.
function setupTrainingTriggers_(ss) {
  var have = {};
  ScriptApp.getProjectTriggers().forEach(function (t) { have[t.getHandlerFunction()] = true; });
  if (!have.onFormSubmit) ScriptApp.newTrigger('onFormSubmit').forSpreadsheet(ss).onFormSubmit().create();
  if (!have.dailyReminderCheck) ScriptApp.newTrigger('dailyReminderCheck').timeBased().everyDays(1).atHour(20).create();
  if (!have.checkInactiveTrainees) ScriptApp.newTrigger('checkInactiveTrainees').timeBased().everyDays(1).atHour(6).create();
  if (!have.checkForDuplicates) ScriptApp.newTrigger('checkForDuplicates').timeBased().everyDays(1).atHour(7).create();
  if (!have.mondayAutoPopulate) {
    ScriptApp.newTrigger('mondayAutoPopulate').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(5).create();
  }
}
