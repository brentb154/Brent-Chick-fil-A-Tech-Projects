/**
 * ============================================================
 * CHECKLIST SYSTEM - Menu and Triggers
 * ============================================================
 * Container-bound to "Checklist Master (FOH + BOH)".
 *
 * Triggers (Checklists > Install / repair triggers):
 *   quarterHourTick   - every 15 minutes. Runs the morning rebuild
 *                       and daily summary once each day after the
 *                       times in Settings, and checks statuses on
 *                       every run. Changing a time in Settings takes
 *                       effect on the next run - no reinstall needed.
 *   onChecklistSubmit - one onFormSubmit trigger per checklist form.
 *
 * The photo upload page (doGet in 10_Photos.gs) is a web app, deployed
 * separately from the store account. See the top of 10_Photos.gs.
 *
 * Install triggers from the store account (sheet owner). They run
 * as whoever installs them.
 */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Checklists')
    .addItem('Add a task', 'menuAddItem')
    .addItem('Edit a task', 'menuEditItem')
    .addItem('Rebuild forms now', 'menuRebuildForms')
    .addItem('Validate sheet', 'menuValidate')
    .addItem('Send test alert to me', 'menuSendTestAlert')
    .addItem('Show form links', 'menuShowFormLinks')
    .addItem('View photos', 'menuViewPhotos')
    .addSeparator()
    .addItem('Install / repair triggers', 'menuInstallTriggers')
    .addItem('Remove my triggers', 'menuRemoveTriggers')
    .addToUi();
}

// -- Trigger entry points -------------------------------------

function quarterHourTick() {
  if (!isTriggerOwner_()) return;
  var cfg;
  try {
    cfg = loadSettings_();
  } catch (err) {
    reportError_('Reading Settings', err);
    return;
  }
  var props = PropertiesService.getScriptProperties();
  var started = Date.now();
  var now = new Date();
  var today = dateKey_(now, cfg.tz);   // calendar date, not business date
  var wall = wallMinutes_(now, cfg.tz);

  if (wall >= cfg.rebuildMin && props.getProperty('REBUILD_DONE') !== today) {
    runJob_('Morning rebuild', function () { morningRebuild_(cfg, today); });
  }

  runJob_('Status check', function () { checkStatuses_(cfg); });

  // A long rebuild can use most of the 6-minute limit; the summary waits for the next run if so
  var hasTime = Date.now() - started < 4 * 60 * 1000;
  if (hasTime && wall >= cfg.summaryMin && props.getProperty('SUMMARY_DONE') !== today) {
    runJob_('Daily summary', function () {
      sendDailySummary_(cfg, today);
      props.setProperty('SUMMARY_DONE', today);
    });
  }
}

function runJob_(name, fn) {
  try {
    fn();
  } catch (err) {
    reportError_(name, err);
  }
}

// Only the account that installed the triggers does any work. Leftover triggers
// from another account (they can't be deleted by anyone else) just exit.
function isTriggerOwner_() {
  var owner = PropertiesService.getScriptProperties().getProperty('TRIGGER_OWNER');
  if (!owner) return true;
  var me = '';
  try {
    me = Session.getEffectiveUser().getEmail();
  } catch (err) {}
  return !me || me === owner;
}

// -- Menu actions ---------------------------------------------

function menuRebuildForms() {
  var ui = SpreadsheetApp.getUi();
  var ok = ui.alert('Rebuild forms now',
    'Rebuilds today\'s forms from the sheet. Anyone filling out a checklist right now will lose their answers. Continue?',
    ui.ButtonSet.YES_NO);
  if (ok !== ui.Button.YES) return;

  var check = validateSheet_();
  if (check.blocking.length) {
    showValidation_(check);
    return;
  }
  var cfg = loadSettings_();
  var today = dateKey_(new Date(), cfg.tz);
  clearRebuildFlags_();

  // The forms belong to the account that runs the triggers. Anyone else (a director editing the
  // sheet) can't open them, so their request is handed to the 15-minute check, which runs as the owner.
  if (!isTriggerOwner_()) {
    ui.alert('Rebuild queued',
      'The forms are owned by ' + PropertiesService.getScriptProperties().getProperty('TRIGGER_OWNER') +
      ', so the rebuild runs automatically on the next check, within 15 minutes.', ui.ButtonSet.OK);
    return;
  }
  var done = rebuildForms_(cfg, today);
  if (done) PropertiesService.getScriptProperties().setProperty('REBUILD_DONE', today);
  ensureDailyStatus_(cfg, today);
  ui.alert(done
    ? 'Forms rebuilt for ' + longLabel_(today) + '.'
    : 'Ran out of time partway through. The rest will finish on the next automatic check (within 15 minutes).');
}

function menuValidate() {
  showValidation_(validateSheet_());
}

function menuSendTestAlert() {
  var sentTo = sendTestAlert_();
  SpreadsheetApp.getUi().alert('Test alert sent to ' + sentTo + '.');
}

function menuShowFormLinks() {
  var model = loadModel_();
  var rows = model.checklists.filter(function (c) { return c.active; }).map(function (c) {
    var link = c.formLink
      ? '<a href="' + esc_(c.formLink) + '" target="_blank">' + esc_(c.formLink) + '</a>'
      : '<i>No form yet. Use Rebuild forms now.</i>';
    return '<p><b>' + esc_(c.name) + '</b><br>' + link + '</p>';
  });
  var photoLink = photoLink_(loadSettings_(), '');
  rows.push('<p><b>Photo upload page</b> (for the QR code)<br>' + (photoLink
    ? '<a href="' + esc_(photoLink) + '" target="_blank">' + esc_(photoLink) + '</a>'
    : '<i>Not set up. See "Photo upload page" in Settings.</i>') + '</p>');
  var html = '<div style="font-family:Arial,sans-serif;font-size:13px;word-break:break-all">' +
    (rows.join('') || '<p>No active checklists.</p>') + '</div>';
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(620).setHeight(400), 'Form links');
}

function menuInstallTriggers() {
  var ui = SpreadsheetApp.getUi();
  var ss = getSS_();
  var props = PropertiesService.getScriptProperties();
  var me = Session.getEffectiveUser().getEmail();
  var sheetOwner = ss.getOwner() ? ss.getOwner().getEmail() : '';
  var previous = props.getProperty('TRIGGER_OWNER');

  var msg = 'Triggers will run as ' + me + '.';
  if (sheetOwner && sheetOwner !== me) msg += '\n\nThis sheet is owned by ' + sheetOwner + '. Triggers should be installed from that account.';
  if (previous && previous !== me) msg += '\n\nTriggers were last installed by ' + previous + '. Theirs will stop doing anything; they can clear them with Checklists > Remove my triggers.';
  if (ui.alert('Install / repair triggers', msg + '\n\nContinue?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

  deleteTriggers();
  props.setProperty('SS_ID', ss.getId());
  props.setProperty('TRIGGER_OWNER', me);
  ScriptApp.newTrigger('quarterHourTick').timeBased().everyMinutes(15).create();
  var count = ensureSubmitTriggers_();
  ui.alert('Installed the 15-minute check and ' + count + ' form submit trigger(s).\n\n' +
    'Checklists without a form yet get their trigger automatically when their form is first built.');
}

function menuRemoveTriggers() {
  var ui = SpreadsheetApp.getUi();
  var ok = ui.alert('Remove my triggers',
    'Removes every trigger YOUR account installed for this project. If you are the account running the system, ' +
    'forms stop rebuilding and submissions stop logging until someone runs Install / repair triggers. Continue?',
    ui.ButtonSet.YES_NO);
  if (ok !== ui.Button.YES) return;
  deleteTriggers();
  ui.alert('Your triggers were removed.');
}

// Removes this project's triggers owned by the account running it
function deleteTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (t) { ScriptApp.deleteTrigger(t); });
}
