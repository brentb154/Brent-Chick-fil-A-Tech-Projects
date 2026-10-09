/**
 * ============================================================
 * CHECKLIST SYSTEM - Late Alerts and Error Alerts
 * ============================================================
 * Alerts only go out when something wasn't done. There is no
 * "heads-up" before the due time, only a late alert after it.
 *
 * Recipients are read from Settings on every run:
 *   "<Area> escalation email"   e.g. FOH escalation email
 * Changing that cell is the only step needed to change who gets
 * alerted. A new Area works by adding the same row for it.
 *
 * Each late alert is sent once per Daily Status row: the
 * Escalation sent cell is stamped in the same locked run.
 */

// Live mode only. Caller holds the lock. One message per checklist listing every position still out.
function sendDueAlerts_(cfg, tab, todayRows, current) {
  var model = loadModel_();
  var groups = {};
  todayRows.forEach(function (r) {
    if (r.status !== 'Pending' && r.status !== 'Late') return;
    (groups[r.checklistId] = groups[r.checklistId] || []).push(r);
  });

  Object.keys(groups).forEach(function (id) {
    var checklist = model.byId[id];
    var open = groups[id];
    var first = open[0];
    if (!checklist || first.lateMin === null) return;
    if (current.minutes < bizMinutes_(first.lateMin, cfg)) return;

    var toEscalate = open.filter(function (r) { return !r.escalationSent; });
    if (toEscalate.length && sendEscalation_(cfg, checklist, open, first)) {
      stampRows_(tab, toEscalate, 'Escalation sent');
    }
  });
}

function stampRows_(tab, rows, header) {
  var now = new Date();
  rows.forEach(function (r) { setCell_(tab, r.index, header, now); });
}

function missingList_(checklist, open) {
  return checklist.perPosition ? open.map(function (r) { return r.position; }) : [];
}

// Returns true if it went out (so the rows get stamped)
function sendEscalation_(cfg, checklist, open, first) {
  var to = emailList_(cfg.get(checklist.area + ' escalation email'));
  if (!to.length) return false;
  var missing = missingList_(checklist, open);
  var what = checklist.perPosition
    ? missing.length + (missing.length === 1 ? ' position' : ' positions') + ' missing'
    : 'not submitted';
  var subject = 'LATE: ' + shortName_(checklist) + ' – ' + what + ' (' + shortLabel_(first.dateKey) + ')';
  var body = checklist.name + ' was due at ' + first.dueText + ' and is now late (after ' + first.lateText + ').\n\n' +
    (missing.length ? 'Still missing:\n- ' + missing.join('\n- ') + '\n\n' : 'No submission yet.\n\n') +
    'Form: ' + checklist.formLink + '\n' +
    'Checklist Master: ' + getSS_().getUrl();
  MailApp.sendEmail({ to: to.join(','), subject: subject, body: body });
  logAlert_(first.dateKey, checklist.id, missing.join(', '), 'Escalation', to.join(', '));
  return true;
}

// Never throws: if logging failed after an email went out, the caller would never mark
// the alert as sent and the same email would go out again every 15 minutes.
function logAlert_(dateKey, checklistId, position, type, sentTo, details) {
  try {
    var tab = openTab_(TABS.alerts);
    var row = {
      'Sent at': new Date(),
      'Date': dateKey,
      'Checklist ID': checklistId,
      'Position': position,
      'Alert type': type,
      'Sent to': sentTo
    };
    if (details && tab.col('Details') > -1) row['Details'] = String(details).slice(0, 1000); // optional column
    appendRows_(tab, [row]);
  } catch (err) {
    // The email already went out; a missing log row is the lesser problem
  }
}

// FOH escalation email, or the account the script runs as if that's blank
// Who hears about system errors and blocked rebuilds: Settings "System error emails",
// else the FOH escalation email, else the account running the script
function adminEmails_(cfg) {
  var to = emailList_(cfg.get('System error emails'));
  if (!to.length) to = emailList_(cfg.get('FOH escalation email'));
  if (!to.length) to = [Session.getEffectiveUser().getEmail()];
  return to;
}

function sendBlockedAlert_(cfg, today, problems) {
  var to = adminEmails_(cfg);
  MailApp.sendEmail({
    to: to.join(','),
    subject: 'Checklist forms NOT rebuilt for ' + shortLabel_(today) + ' – sheet needs a fix',
    body: 'Validate found problems in the Checklist Master, so yesterday\'s forms were left in place.\n\n- ' +
      problems.join('\n- ') + '\n\nFix them, then use Checklists > Rebuild forms now (or wait for the next ' +
      'automatic check, within 15 minutes).\n\n' + getSS_().getUrl()
  });
  logAlert_(today, '', '', 'Rebuild blocked', to.join(', '));
}

// One email per job per day, so a stuck trigger can't burn the daily email quota.
function reportError_(job, err) {
  try {
    var today = dateKey_(new Date(), Session.getScriptTimeZone());
    var props = PropertiesService.getScriptProperties();
    var key = 'ERROR_' + job.replace(/\W+/g, '_').toUpperCase();
    if (props.getProperty(key) === today) return;
    props.setProperty(key, today);
    var to;
    try {
      to = adminEmails_(loadSettings_());
    } catch (settingsErr) {
      to = [Session.getEffectiveUser().getEmail()]; // Settings itself may be what broke
    }
    MailApp.sendEmail({
      to: to.join(','),
      subject: 'Checklist system error: ' + job,
      body: job + ' failed:\n\n' + (err && err.stack ? err.stack : err) + '\n\n' +
        'Further errors from this step today won\'t be emailed.\n' + getSS_().getUrl()
    });
    logAlert_(today, '', '', 'Error: ' + job, to.join(', '), err && err.message ? err.message : String(err));
  } catch (ignore) {
    // Nothing else to do if even the alert fails
  }
}

function sendTestAlert_() {
  var cfg = loadSettings_();
  var model = loadModel_();
  var me = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail();
  var checklist = model.checklists.filter(function (c) { return c.active; })[0];
  var today = dateKey_(new Date(), cfg.tz);
  var name = checklist ? shortName_(checklist) : 'Positional Closing';
  MailApp.sendEmail({
    to: me,
    subject: 'TEST – LATE: ' + name + ' – 3 positions missing (' + shortLabel_(today) + ')',
    body: 'This is a test of the late alert. A real one goes to the "' + (checklist ? checklist.area : 'FOH') +
      ' escalation email" in Settings.\n\nStill missing:\n- Dishes\n- Stocker\n- Team Leader\n\n' +
      'Form: ' + (checklist ? checklist.formLink : '') + '\nChecklist Master: ' + getSS_().getUrl()
  });
  logAlert_(today, checklist ? checklist.id : '', '', 'Test', me);
  return me;
}
