// ============================================================
// Setup.gs — first-time setup for template copies
// ============================================================
// A new store opens its copy of the sheet -> "📊 Schedule Counter Setup" -> Start Setup.
// Saves the alert email, builds the tabs (initSheets) and installs the weekly jobs
// (installTriggers). Safe to re-run. Store name and goals are asked by the web
// app's own first-run onboarding.

function onOpen() {
  SpreadsheetApp.getUi().createMenu('📊 Schedule Counter Setup')
    .addItem('Start Setup', 'showSetupWizard')
    .addToUi();
}

function showSetupWizard() {
  var html = HtmlService.createHtmlOutputFromFile('SetupWizard').setWidth(520).setHeight(620);
  SpreadsheetApp.getUi().showModalDialog(html, 'Schedule Counter Setup');
}

function getSetupStatus() {
  return { alertEmail: PropertiesService.getScriptProperties().getProperty('ALERT_EMAIL') || '' };
}

function saveSetupAnswers(answers) {
  var email = String((answers && answers.alertEmail) || '').trim();
  if (email.indexOf('@') < 1) return { success: false, message: 'Please enter the email that should get alerts.' };

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    PropertiesService.getScriptProperties().setProperty('ALERT_EMAIL', email);
    initSheets();
    installTriggers();
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return { success: true };
}
