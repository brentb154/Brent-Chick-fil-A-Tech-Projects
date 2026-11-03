/**
 * ============================================================
 * CHECKLIST SYSTEM - First-time setup (template copies)
 * ============================================================
 * A new store opens its copy -> "✅ Checklist Setup" -> Start Setup.
 * Saves the alert and summary emails to Settings, then lists the
 * three menu steps that build the forms and turn on the triggers.
 * Safe to re-run: it shows what's already in Settings.
 */

var SETUP_EMAIL_ROWS = {
  foh: 'FOH escalation email',
  boh: 'BOH escalation email',
  summary: 'Daily summary recipients',
  errors: 'System error emails'
};

// Like the other pop-ups, the wizard gets a pass from the menu (see newPopupPass_ in 02_Helpers.gs),
// so the public photo page can't call these functions.
function showSetupWizard() {
  var t = HtmlService.createTemplateFromFile('SetupWizard');
  t.pass = newPopupPass_();
  SpreadsheetApp.getUi().showModalDialog(t.evaluate().setWidth(540).setHeight(660), 'Checklist Setup');
}

function getSetupStatus(pass) {
  checkPopupPass_(pass);
  var rows = settingsRows_();
  var out = {};
  Object.keys(SETUP_EMAIL_ROWS).forEach(function (k) {
    var r = rows[SETUP_EMAIL_ROWS[k]];
    out[k] = r ? String(r.value).trim() : '';
  });
  return out;
}

function saveSetupAnswers(answers, pass) {
  checkPopupPass_(pass);
  answers = answers || {};
  var clean = {};
  var bad = [];
  Object.keys(SETUP_EMAIL_ROWS).forEach(function (k) {
    var list = String(answers[k] || '').split(/[,;\s]+/).filter(function (v) { return v; });
    list.forEach(function (v) { if (v.indexOf('@') < 1) bad.push(v); });
    clean[k] = list.join(',');
  });
  if (bad.length) return { success: false, message: 'Check these emails: ' + bad.join(', ') };
  if (!clean.foh) return { success: false, message: 'Please enter at least one email for late alerts.' };
  if (!clean.summary) return { success: false, message: 'Please enter at least one email for the morning summary.' };

  var sheet = getSheet_(TABS.settings);
  var rows = settingsRows_();
  Object.keys(SETUP_EMAIL_ROWS).forEach(function (k) {
    var r = rows[SETUP_EMAIL_ROWS[k]];
    if (r) sheet.getRange(r.row, 2).setValue(clean[k]);
  });
  SpreadsheetApp.flush();
  return { success: true };
}

// Settings label -> { row, value }
function settingsRows_() {
  var values = getSheet_(TABS.settings).getDataRange().getValues();
  var out = {};
  for (var i = 1; i < values.length; i++) {
    var label = String(values[i][0]).trim();
    if (label) out[label] = { row: i + 1, value: values[i][1] };
  }
  return out;
}
