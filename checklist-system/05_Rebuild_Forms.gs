/**
 * ============================================================
 * CHECKLIST SYSTEM - Daily Form Rebuild
 * ============================================================
 * Rebuilds each active checklist's form with only that day's
 * tasks. A form is created once; after that it's reopened by
 * the Form ID in Checklists so the link never changes.
 */

var Q_LEADER = 'Leader completing this checklist';
var Q_POSITION = 'Which position are you checking off?';
var Q_NOTES = 'Anything not completed or needing attention?';
var ANSWER_DONE = 'Complete';
var ANSWER_NOT_DONE = 'Could not complete';
var PHOTO_SUFFIX = ' **EMAIL PICTURES**';
var REBUILD_TIME_LIMIT_MS = 3 * 60 * 1000; // leaves room for one long form inside the 6-minute limit

// Morning step: validate, rebuild, then create today's Daily Status rows.
// Blocking problems keep yesterday's forms and send one warning a day; each run retries.
function morningRebuild_(cfg, today) {
  var props = PropertiesService.getScriptProperties();
  var check = validateSheet_();
  if (check.blocking.length) {
    if (props.getProperty('BLOCK_ALERT') !== today) {
      sendBlockedAlert_(cfg, today, check.blocking);
      props.setProperty('BLOCK_ALERT', today);
    }
  } else if (rebuildForms_(cfg, today)) {
    props.setProperty('REBUILD_DONE', today);
  }
  ensureDailyStatus_(cfg, today);
  ensureSubmitTriggers_(); // daily repair: every active form has a submit trigger owned by this account
  trashOldPhotos_(cfg, today);
}

// Returns true when every active checklist is done for dateKey. If it runs long, or another
// rebuild is already running, it returns false; the next 15-minute run picks up what's left.
function rebuildForms_(cfg, dateKey) {
  if (!claimRebuild_()) return false;
  try {
    var started = Date.now();
    var props = PropertiesService.getScriptProperties();
    var model = loadModel_();
    var active = model.checklists.filter(function (c) { return c.active; });

    for (var i = 0; i < active.length; i++) {
      var checklist = active[i];
      var doneKey = 'REBUILT_' + checklist.id;
      if (props.getProperty(doneKey) === dateKey) continue;
      if (Date.now() - started > REBUILD_TIME_LIMIT_MS) return false;

      var form = openOrCreateForm_(checklist);
      if (scheduledOn_(checklist, dateKey, cfg)) {
        buildForm_(form, checklist, dateKey, cfg);
      } else {
        form.setAcceptingResponses(false).setCustomClosedFormMessage('No checklist today.');
      }
      props.setProperty(doneKey, dateKey);
    }
    return true;
  } finally {
    PropertiesService.getScriptProperties().deleteProperty('REBUILD_RUNNING');
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

function buildForm_(form, checklist, dateKey, cfg) {
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

  var old = form.getItems();
  for (var i = old.length - 1; i >= 0; i--) form.deleteItem(old[i]);

  form.addTextItem().setTitle(Q_LEADER).setRequired(true);
  var items = itemsOn_(checklist, dateKey);
  var mapRows = [];

  if (checklist.perPosition) {
    var positionQ = form.addMultipleChoiceItem().setTitle(Q_POSITION).setRequired(true);
    var choices = [];
    requiredPositions_(checklist).forEach(function (pos, idx) {
      var page = form.addPageBreakItem().setTitle(pos.name);
      // A page break's "go to" controls the END of the section BEFORE it. Setting SUBMIT here
      // makes the previous position submit instead of running into this one (the old Stocker bug).
      // The first break follows the intro page, where the position answer decides the route.
      if (idx > 0) page.setGoToPage(FormApp.PageNavigationType.SUBMIT);
      choices.push(positionQ.createChoice(pos.name, page));
      var posItems = items.filter(function (it) { return it.positionKey === pos.key; });
      addTaskQuestions_(form, posItems, checklist.id, dateKey, mapRows);
      form.addParagraphTextItem().setTitle(Q_NOTES);
    });
    positionQ.setChoices(choices);
  } else {
    addTaskQuestions_(form, items, checklist.id, dateKey, mapRows);
    form.addParagraphTextItem().setTitle(Q_NOTES);
  }

  form.setAcceptingResponses(true);
  saveFormMap_(checklist.id, dateKey, mapRows);
}

function addTaskQuestions_(form, items, checklistId, dateKey, mapRows) {
  items.forEach(function (it) {
    var q = form.addMultipleChoiceItem()
      .setTitle(it.task + (it.photo ? PHOTO_SUFFIX : ''))
      .setChoiceValues([ANSWER_DONE, ANSWER_NOT_DONE])
      .setRequired(true);
    if (it.reference) q.setHelpText(it.reference);
    mapRows.push({
      'Checklist ID': checklistId,
      'Built for date': dateKey,
      'Form question ID': String(q.getId()),
      'Item ID': it.id,
      'Position': it.positionName
    });
  });
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
