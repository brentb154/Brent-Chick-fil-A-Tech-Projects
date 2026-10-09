/**
 * ============================================================
 * CHECKLIST SYSTEM - Daily Form Rebuild
 * ============================================================
 * Updates each active checklist's form to that day's tasks,
 * changing only the questions that differ (15_Form_Sync.gs). A
 * form is created once; after that it's reopened by the Form ID
 * in Checklists so the link never changes.
 */

var Q_LEADER = 'Leader completing this checklist';
var Q_POSITION = 'Which position are you checking off?';
var Q_NOTES = 'Anything not completed or needing attention?';
var Q_STATION_CODE = 'Station code (filled in when you scan the station\'s QR code)';
var ANSWER_DONE = 'Complete';
var ANSWER_NOT_DONE = 'Could not complete';
var PHOTO_SUFFIX = ' **TAKE PICTURES**';
var REBUILD_BUDGET_MS = 4 * 60 * 1000; // stops well inside Google's 6-minute limit; the next run carries on

// Morning step: validate, rebuild, then create today's Daily Status rows.
// Blocking problems keep yesterday's forms and send one warning a day; each run retries.
// Today's rows and the upkeep steps run even when a form fails to rebuild, so one broken
// form can't leave the whole day untracked. Each step reports its own errors.
function morningRebuild_(cfg, today) {
  var props = PropertiesService.getScriptProperties();
  try {
    var check = validateSheet_();
    if (check.blocking.length) {
      if (props.getProperty('BLOCK_ALERT') !== today) {
        sendBlockedAlert_(cfg, today, check.blocking);
        props.setProperty('BLOCK_ALERT', today);
      }
    } else if (rebuildForms_(cfg, today)) {
      props.setProperty('REBUILD_DONE', today);
    }
  } finally {
    runJob_('Daily Status rows', function () { ensureDailyStatus_(cfg, today); });
    runJob_('Form submit triggers', ensureSubmitTriggers_); // every active form has a submit trigger owned by this account
    runJob_('Photo cleanup', function () { trashOldPhotos_(cfg, today); });
  }
}

// Returns true when every active checklist is done for dateKey. If it runs long, or another
// rebuild is already running, it returns false; the next 15-minute run picks up what's left.
function rebuildForms_(cfg, dateKey) {
  if (!claimRebuild_()) return false;
  try {
    var deadline = Date.now() + REBUILD_BUDGET_MS;
    var props = PropertiesService.getScriptProperties();
    var model = loadModel_();
    var active = model.checklists.filter(function (c) { return c.active; });
    var failures = [];

    for (var i = 0; i < active.length; i++) {
      var checklist = active[i];
      var doneKey = 'REBUILT_' + checklist.id;
      if (props.getProperty(doneKey) === dateKey) continue;
      if (Date.now() > deadline - 20 * 1000) return false;

      // One broken form shouldn't stop the others from rebuilding
      try {
        var form = openOrCreateForm_(checklist);
        if (scheduledOn_(checklist, dateKey, cfg)) {
          if (!buildForm_(form, checklist, dateKey, cfg, deadline)) return false; // out of time; the next run carries on
        } else {
          closeForm_(form);
        }
        props.setProperty(doneKey, dateKey);
      } catch (err) {
        failures.push(checklist.name + ': ' + err.message);
      }
    }
    if (failures.length) throw new Error('Some forms did not rebuild. ' + failures.join(' | '));
    return true;
  } finally {
    PropertiesService.getScriptProperties().deleteProperty('REBUILD_RUNNING');
  }
}

// A day the checklist doesn't run: stop responses and, where Google allows it, show
// "No checklist today." Google rejects the closed message on these forms ("Invalid data
// updating form", seen 10/8/2026); without it, responders get Google's own "no longer
// accepting responses" page, so it's best-effort. A form left open is a real failure.
function closeForm_(form) {
  form.setAcceptingResponses(false);
  try {
    form.setCustomClosedFormMessage('No checklist today.');
  } catch (err) {
    // cosmetic; see above
  }
}

// Only one rebuild at a time: two runs editing the same form at once would scramble it.
// A claim older than 6 minutes is from a run that died, so it's ignored.
function claimRebuild_() {
  return withLock_(function () {
    var props = PropertiesService.getScriptProperties();
    var running = Number(props.getProperty('REBUILD_RUNNING') || 0);
    if (running && Date.now() - running < 6 * 60 * 1000) return false;
    props.setProperty('REBUILD_RUNNING', String(Date.now()));
    return true;
  });
}

function clearRebuildFlags_() {
  var props = PropertiesService.getScriptProperties();
  props.getKeys().forEach(function (k) {
    if (k.indexOf('REBUILT_') === 0 || k === 'REBUILD_DONE') props.deleteProperty(k);
  });
}

// Never creates a second form for a checklist. If the saved Form ID can't be opened, it stops and says why.
function openOrCreateForm_(checklist) {
  if (checklist.formId) {
    try {
      return FormApp.openById(checklist.formId);
    } catch (err) {
      throw new Error('Can\'t open the form for ' + checklist.id + ' (Form ID ' + checklist.formId + '). ' +
        'If it was deleted on purpose, clear the Form ID and Form link cells and rebuild. ' + err.message);
    }
  }
  var form = FormApp.create(checklist.name, true);
  var link = form.getPublishedUrl();
  withLock_(function () {
    // Find the row again by Checklist ID: rows may have been added or sorted since the sheet was read
    var tab = readTab_(TABS.checklists);
    var idCol = colOrThrow_(tab, 'Checklist ID');
    var index = -1;
    tab.display.forEach(function (r, i) { if (index < 0 && r[idCol].trim() === checklist.id) index = i; });
    if (index < 0) throw new Error('Created a form for ' + checklist.id + ' (Form ID ' + form.getId() + ') but that Checklist ID is no longer on the Checklists tab.');
    setCell_(tab, index, 'Form ID', form.getId());
    setCell_(tab, index, 'Form link', link);
    SpreadsheetApp.flush();
  });
  ensureSubmitTriggers_();
  checklist.formId = form.getId();
  checklist.formLink = link;
  return form;
}

// Brings the form up to date for dateKey: header, then only the questions that changed
// (15_Form_Sync.gs). Returns false if it ran out of time; the next run carries on.
function buildForm_(form, checklist, dateKey, cfg, deadline) {
  form.setTitle(checklist.name);
  // Photos: the upload page when it's set up, the photo email while it's still filled in (either or both)
  var uploadLink = photoLink_(cfg, checklist.id);
  var photoEmail = cfg.get('Photo email');
  var photoSubject = cfg.get('Photo email subject');
  var photoLines = [];
  if (uploadLink) photoLines.push('Upload all pictures here (no sign-in): ' + uploadLink);
  if (photoEmail) photoLines.push((uploadLink ? 'Or email them to ' : 'Email all pictures to ') + photoEmail +
    (photoSubject ? ' with the subject "' + photoSubject + '"' : '') + '.');
  form.setDescription(checklist.name + '\n' + longLabel_(dateKey) + (photoLines.length ? '\n\n' + photoLines.join('\n') : ''));
  form.setConfirmationMessage(uploadLink
    ? 'Thanks, your checklist is in. Now upload your pictures: ' + uploadLink
    : 'Thanks, your checklist is in.');
  if (form.supportsAdvancedResponderPermissions() && !form.isPublished()) form.setPublished(true);

  var plan = formPlan_(checklist, dateKey);
  var synced = syncForm_(form, checklist, plan, dateKey, deadline || Infinity);
  if (!synced) return false;

  // Station QR links need today's entry IDs, which only change when these questions are new
  if (checklist.stationQr) {
    var c = synced.created;
    var props = PropertiesService.getScriptProperties();
    if (c.leader || c.position || c.code || !props.getProperty('PREFILL_' + checklist.id)) {
      var at = function (kind) { return form.getItemById(synced.ids[indexOfKind_(plan, kind)]); };
      var first = plan[indexOfKind_(plan, 'page')];
      if (first) savePrefill_(form, checklist, at('leader').asTextItem(), at('position').asMultipleChoiceItem(), at('code').asTextItem(), first.title);
    }
  }

  form.setAcceptingResponses(true);
  var mapRows = [];
  plan.forEach(function (s, j) {
    if (s.kind !== 'task') return;
    mapRows.push({
      'Checklist ID': checklist.id,
      'Built for date': dateKey,
      'Form question ID': String(synced.ids[j]),
      'Item ID': s.itemId,
      'Position': s.position
    });
  });
  saveFormMap_(checklist.id, dateKey, mapRows);
  return true;
}

// Replaces this checklist's rows for dateKey and drops anything older than 14 days.
function saveFormMap_(checklistId, dateKey, newRows) {
  withLock_(function () {
    var tab = readTab_(TABS.formMap);
    var cId = colOrThrow_(tab, 'Checklist ID');
    var cDate = colOrThrow_(tab, 'Built for date');
    var cutoff = addDays_(dateKey, -14);
    var keep = tab.display.filter(function (r) {
      var d = toDateKey_(r[cDate]);
      if (!r[cId] || d < cutoff) return false;
      return !(r[cId] === checklistId && d === dateKey);
    });
    if (tab.display.length) tab.sheet.getRange(2, 1, tab.display.length, tab.width).clearContent();
    var rows = keep.concat(newRows.map(function (o) { return toRow_(tab, o); }));
    writeRows_(tab, 2, rows);
    SpreadsheetApp.flush();
  });
}

// questionId -> { itemId, position, date }, plus all rows for counting expected tasks
function loadFormMap_(checklistId) {
  var tab = readTab_(TABS.formMap);
  var c = {
    id: colOrThrow_(tab, 'Checklist ID'),
    date: colOrThrow_(tab, 'Built for date'),
    q: colOrThrow_(tab, 'Form question ID'),
    item: colOrThrow_(tab, 'Item ID'),
    pos: colOrThrow_(tab, 'Position')
  };
  var byQuestion = {};
  var rows = [];
  tab.display.forEach(function (r) {
    if (r[c.id] !== checklistId) return;
    var row = { itemId: r[c.item], position: r[c.pos], date: toDateKey_(r[c.date]) };
    byQuestion[String(r[c.q]).trim()] = row;
    rows.push(row);
  });
  return { byQuestion: byQuestion, rows: rows };
}

// Gives every active checklist form an onFormSubmit trigger owned by this account.
// Only the trigger owner does this; a trigger owned by anyone else would never log anything.
// Returns how many active forms have a trigger.
function ensureSubmitTriggers_() {
  if (!isTriggerOwner_()) return 0;
  var have = {};
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'onChecklistSubmit') have[t.getTriggerSourceId()] = true;
  });
  var count = 0;
  loadModel_().checklists.forEach(function (c) {
    if (!c.active || !c.formId) return;
    if (!have[c.formId]) {
      ScriptApp.newTrigger('onChecklistSubmit').forForm(c.formId).onFormSubmit().create();
      have[c.formId] = true;
    }
    count++;
  });
  return count;
}
