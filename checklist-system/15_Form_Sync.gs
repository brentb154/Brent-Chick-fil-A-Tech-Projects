/**
 * ============================================================
 * CHECKLIST SYSTEM - Form Sync (the fast daily rebuild)
 * ============================================================
 * Updates a form to match today's questions instead of deleting
 * and recreating every one. Google Forms makes each change a
 * separate slow call, so a full rebuild of the closing form
 * (~800 calls) ran into Google's 6-minute limit. Most days only a
 * few questions change, so a sync takes a few dozen calls.
 *
 * After each sync the form's layout (each question's ID plus a
 * short key of its type, title and help text) is kept in Script
 * Properties. The next sync compares it with today's plan and
 * only deletes and adds what differs. When the layout can't be
 * trusted (first run, a form edited by hand, or a week old), the
 * form is read again first.
 *
 * Every step checks the deadline. Out of time, it saves how far it
 * got and the next 15-minute run carries on from there.
 */

var LAYOUT_REREAD_DAYS = 7;
var TYPE_CODES = { TEXT: 'T', MULTIPLE_CHOICE: 'M', PAGE_BREAK: 'B', PARAGRAPH_TEXT: 'N' };

// -- Today's plan ---------------------------------------------

// What today's form should hold, top to bottom
function formPlan_(checklist, dateKey) {
  var plan = [slot_('leader', 'TEXT', Q_LEADER)];
  var tasks = itemsOn_(checklist, dateKey).map(function (it) { return taskSlot_(it, dateKey); });
  if (checklist.perPosition) {
    plan.push(slot_('position', 'MULTIPLE_CHOICE', Q_POSITION));
    if (checklist.stationQr) plan.push(slot_('code', 'TEXT', Q_STATION_CODE));
    requiredPositions_(checklist).forEach(function (pos, idx) {
      plan.push(slot_('page', 'PAGE_BREAK', pos.name, '', idx === 0));
      tasks.forEach(function (t) { if (t.positionKey === pos.key) plan.push(t); });
      plan.push(slot_('notes', 'PARAGRAPH_TEXT', Q_NOTES));
    });
  } else {
    plan = plan.concat(tasks);
    plan.push(slot_('notes', 'PARAGRAPH_TEXT', Q_NOTES));
  }
  return plan;
}

function slot_(kind, type, title, help, first) {
  return { kind: kind, type: type, title: title, help: help || '', first: !!first, key: slotKey_(type, title, help, first) };
}

function taskSlot_(it, dateKey) {
  var s = slot_('task', it.typeIn ? 'TEXT' : 'MULTIPLE_CHOICE', pickOne_(it.task, dateKey) + (it.photo ? PHOTO_SUFFIX : ''), it.reference);
  s.itemId = it.id;
  s.position = it.positionName;
  s.positionKey = it.positionKey;
  return s;
}

// "{Walk In Cooler|Fry Freezer}" -> one of the options. Picked from the date, so it holds all
// day (a re-run or Rebuild forms now won't switch it) and changes from day to day.
function pickOne_(text, dateKey) {
  return String(text).replace(/\{([^{}]*\|[^{}]*)\}/g, function (all, list) {
    var options = list.split('|').map(function (s) { return s.trim(); }).filter(function (s) { return s; });
    if (!options.length) return all;
    var d = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, dateKey + '|' + list, Utilities.Charset.UTF_8);
    return options[((d[0] & 255) * 256 + (d[1] & 255)) % options.length];
  });
}

// A question's identity for matching: type, title, help text, and for sections whether it's the
// first one (the first section has no "go to"). The first letter keeps the type readable.
function slotKey_(type, title, help, first) {
  var d = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, [type, title, help || '', first ? 1 : 0].join('\u0001'), Utilities.Charset.UTF_8);
  return (TYPE_CODES[type] || '?') + Utilities.base64Encode(d).slice(0, 10);
}

// -- Sync -----------------------------------------------------

// Makes the form match the plan, changing only what differs. Returns the question IDs (one per
// plan slot) and which special questions were made new, or null if it ran out of time.
function syncForm_(form, checklist, plan, dateKey, deadline) {
  var formId = form.getId();
  var saved = loadLayout_(checklist.id);
  var items = form.getItems();
  var savedList = saved ? saved.items.map(function (x) { return { id: x[0], key: x[1] }; }) : [];
  var state;
  if (saved && !saved.reading && saved.form === formId && savedList.length === items.length &&
      dayNumber_(dateKey) - dayNumber_(saved.read) < LAYOUT_REREAD_DAYS) {
    state = { read: saved.read, route: !!saved.route, list: savedList };
  } else {
    // Read the form, picking up where an unfinished read of this same form left off
    var resume = saved && saved.reading && saved.form === formId && saved.count === items.length;
    var fresh = readLayout_(items, resume ? savedList : [], deadline);
    if (fresh.length < items.length) {
      saveLayout_(checklist.id, formId, { read: dateKey, reading: true, count: items.length, list: fresh });
      return null;
    }
    state = { read: dateKey, route: true, list: fresh }; // after a fresh read, set routing again to be safe
  }

  var list = state.list;
  var match = matchKeys_(list.map(function (x) { return x.key; }), plan.map(function (s) { return s.key; }));
  var created = {};
  try {
    // A section can't be deleted while the position question still routes to it
    var dropsSection = list.some(function (x, i) { return !match.old[i] && x.key.charAt(0) === 'B'; });
    if (dropsSection) {
      unroutePositions_(form, plan, list);
      state.route = true;
    }
    for (var i = list.length - 1; i >= 0; i--) { // last to first, so earlier indexes stay put
      if (match.old[i]) continue;
      if (Date.now() > deadline) return null;
      form.deleteItem(i);
      list.splice(i, 1);
    }
    for (var j = 0; j < plan.length; j++) {
      if (match.fresh[j]) continue;
      if (Date.now() > deadline) return null;
      var item;
      try {
        item = addSlot_(form, plan[j]);
      } catch (err) {
        try { form.deleteItem(list.length); } catch (ignore) {} // don't leave a half-made question behind
        throw err;
      }
      var entry = { id: item.getId(), key: plan[j].key };
      if (j < list.length) form.moveItem(list.length, j);
      list.splice(j, 0, entry);
      created[plan[j].kind] = true;
      if (plan[j].kind === 'page' || plan[j].kind === 'position') state.route = true;
    }
    if (state.route) {
      if (Date.now() > deadline) return null;
      routePositions_(form, plan, list);
      state.route = false;
    }
    return { ids: list.map(function (x) { return x.id; }), created: created };
  } finally {
    saveLayout_(checklist.id, formId, state); // also on the way out early, so the next run carries on
  }
}

// Longest common subsequence of keys: the questions that can stay. old[i] / fresh[j] mark
// kept indexes in the current layout and in today's plan.
function matchKeys_(oldKeys, newKeys) {
  var n = oldKeys.length;
  var m = newKeys.length;
  var dp = [];
  for (var a = 0; a <= n; a++) {
    dp.push([]);
    for (var b = 0; b <= m; b++) dp[a].push(0);
  }
  for (var i = n - 1; i >= 0; i--) {
    for (var j = m - 1; j >= 0; j--) {
      dp[i][j] = oldKeys[i] === newKeys[j] ? dp[i + 1][j + 1] + 1 : Math.max(dp[i + 1][j], dp[i][j + 1]);
    }
  }
  var out = { old: {}, fresh: {} };
  var x = 0;
  var y = 0;
  while (x < n && y < m) {
    if (oldKeys[x] === newKeys[y]) {
      out.old[x++] = true;
      out.fresh[y++] = true;
    } else if (dp[x + 1][y] >= dp[x][y + 1]) {
      x++;
    } else {
      y++;
    }
  }
  return out;
}

function addSlot_(form, s) {
  if (s.kind === 'page') {
    var page = form.addPageBreakItem().setTitle(s.title);
    // A section's "go to" controls the END of the section before it. SUBMIT keeps one position
    // from running into the next (the old Stocker bug). The first section follows the intro
    // page, where the position answer decides the route.
    if (!s.first) page.setGoToPage(FormApp.PageNavigationType.SUBMIT);
    return page;
  }
  if (s.kind === 'notes') return form.addParagraphTextItem().setTitle(s.title);
  if (s.kind === 'position') return form.addMultipleChoiceItem().setTitle(s.title).setRequired(true);
  var q = s.type === 'TEXT'
    ? form.addTextItem().setTitle(s.title).setRequired(true)
    : form.addMultipleChoiceItem().setTitle(s.title).setChoiceValues([ANSWER_DONE, ANSWER_NOT_DONE]).setRequired(true);
  if (s.help) q.setHelpText(s.help);
  return q;
}

// Each position's answer jumps to its section
function routePositions_(form, plan, list) {
  var p = indexOfKind_(plan, 'position');
  if (p < 0) return;
  var mc = form.getItemById(list[p].id).asMultipleChoiceItem();
  var choices = [];
  plan.forEach(function (s, j) {
    if (s.kind === 'page') choices.push(mc.createChoice(s.title, form.getItemById(list[j].id).asPageBreakItem()));
  });
  if (choices.length) mc.setChoices(choices);
}

// Plain choices (no jumps), so sections can be deleted
function unroutePositions_(form, plan, list) {
  var key = slotKey_('MULTIPLE_CHOICE', Q_POSITION, '', false);
  var at = -1;
  list.forEach(function (x, i) { if (x.key === key) at = i; });
  if (at < 0) return;
  var names = plan.filter(function (s) { return s.kind === 'page'; }).map(function (s) { return s.title; });
  form.getItemById(list[at].id).asMultipleChoiceItem().setChoiceValues(names.length ? names : ['-']);
}

function indexOfKind_(plan, kind) {
  for (var i = 0; i < plan.length; i++) if (plan[i].kind === kind) return i;
  return -1;
}

// -- Layout memory --------------------------------------------

// Reads the form's questions (about 3 calls each) after the ones already read. Out of time, it
// returns what it has; the caller saves that and the next run continues.
function readLayout_(items, done, deadline) {
  var list = done.slice();
  var firstSection = !list.some(function (x) { return x.key.charAt(0) === 'B'; });
  for (var i = list.length; i < items.length; i++) {
    if (Date.now() > deadline) return list;
    var type = String(items[i].getType());
    var hasHelp = type === 'TEXT' || type === 'MULTIPLE_CHOICE';
    list.push({ id: items[i].getId(), key: slotKey_(type, items[i].getTitle(), hasHelp ? items[i].getHelpText() : '', type === 'PAGE_BREAK' && firstSection) });
    if (type === 'PAGE_BREAK') firstSection = false;
  }
  return list;
}

function loadLayout_(checklistId) {
  try {
    return JSON.parse(PropertiesService.getScriptProperties().getProperty('FORM_LAYOUT_' + checklistId) || 'null');
  } catch (err) {
    return null;
  }
}

// If it's ever too big for a Script Property, the next sync just reads the form again
function saveLayout_(checklistId, formId, state) {
  var props = PropertiesService.getScriptProperties();
  var json = JSON.stringify({ form: formId, read: state.read, route: !!state.route, reading: !!state.reading, count: state.count || 0,
    items: state.list.map(function (x) { return [x.id, x.key]; }) });
  if (json.length > 8500) props.deleteProperty('FORM_LAYOUT_' + checklistId);
  else props.setProperty('FORM_LAYOUT_' + checklistId, json);
}
