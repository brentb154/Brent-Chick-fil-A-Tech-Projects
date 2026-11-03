// ------------ FIRST-TIME SETUP (template copies) -------------
// A new store opens its copy of the sheet -> "👂 H.E.A.R.D. Setup" menu -> Start Setup.
// The wizard saves the two passcodes and alert emails to Script Properties and
// builds the data tab. Safe to re-run: it shows the current answers so they can be edited.

function onOpen() {
  SpreadsheetApp.getUi().createMenu('👂 H.E.A.R.D. Setup')
    .addItem('Start Setup', 'showSetupWizard')
    .addToUi();
}

function showSetupWizard() {
  const html = HtmlService.createHtmlOutputFromFile('SetupWizard').setWidth(520).setHeight(640);
  SpreadsheetApp.getUi().showModalDialog(html, 'H.E.A.R.D. Setup');
}

// Current answers, so a re-run shows what's already there.
function getSetupStatus() {
  const props = PropertiesService.getScriptProperties();
  return {
    loginPasscode:    props.getProperty('loginPasscode') || '',
    settingsPasscode: props.getProperty('settingsPasscode') || '',
    emails:           readAlertEmails().join(', ')
  };
}

function saveSetupAnswers(answers) {
  answers = answers || {};
  const login    = String(answers.loginPasscode || '').trim();
  const settings = String(answers.settingsPasscode || '').trim();
  const emails   = String(answers.emails || '').split(/[,;\s]+/).filter(function(v){ return v.indexOf('@') > -1; });
  if (login.length < 4 || login.length > 12)       return { success:false, message:'The login passcode needs 4 to 12 characters.' };
  if (settings.length < 4 || settings.length > 12) return { success:false, message:'The settings passcode needs 4 to 12 characters.' };
  if (login === settings)  return { success:false, message:'Use a different passcode for settings than for login.' };

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const props = PropertiesService.getScriptProperties();
    // The web app has no "active" sheet, so remember which sheet this is.
    props.setProperty('sheetId', SpreadsheetApp.getActiveSpreadsheet().getId());
    props.setProperty('loginPasscode', login);
    props.setProperty('settingsPasscode', settings);
    runFirstTimeSetup();   // data tab, headers, alert storage, API token
    props.setProperty('alertEmails', JSON.stringify(emails));
  } finally {
    lock.releaseLock();
  }
  return { success:true, status:getSetupStatus() };
}
