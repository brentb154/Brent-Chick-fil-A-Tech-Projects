/**
 * ============================================================
 * CHECKLIST SYSTEM - Edit a Task (Checklists > Edit a task)
 * ============================================================
 * The Add a task pop-up in edit mode. Changes a task's wording,
 * days, rotation weeks, photo, reference material, on/off
 * (Active) and place in the list.
 *
 * The Item ID never changes, so the task's history in Item
 * Results stays together. Tasks are turned off, never deleted.
 * To move a task to another position, turn it off and add it
 * there.
 */

var MSG_TASK_CHANGED = 'Someone changed this task on the Items tab after you opened it, so nothing was saved. Close this and open Edit a task again to see the latest.';

function menuEditItem() {
  openTaskDialog_('edit');
}

// p = Add's fields plus { id, active, was }. was = the task as the pop-up loaded it; if the row
// has changed since, nothing is saved. place: 'keep' (default), 'top', 'end' or 'after:<Item ID>'.
function serverEditItem(p) {
  return saveTask_(function () { return updateItem_(p || {}); });
}

function updateItem_(p) {
  var model = loadModel_();
  var checklist = checklistFor_(model, p.checklistId);
  var tab = openTab_(TABS.items);
  var dayIdx = DAY_DISPLAY.filter(function (i) { return tab.col(DAY_SHORT[i]) > -1; });

  // The task must still be on the same row data the pop-up showed
  var matches = model.items.filter(function (it) { return it.id && it.id === p.id; });
  if (matches.length > 1) throw new Error('Item ID ' + p.id + ' is on more than one row of the Items tab. Fix that first (Validate sheet shows where).');
  var item = matches[0];
  var section = item && item.checklistId === checklist.id && sectionsOf_(checklist).filter(function (s) {
    return s.items.indexOf(item) > -1 && posKey_(s.name) === posKey_(p.section);
  })[0];
  if (!section || !sameTask_(taskSnapshot_(item, dayIdx), p.was)) throw new Error(MSG_TASK_CHANGED);

  var input = taskInput_(p, checklist, tab);
  var active = p.active !== false;
  checkDuplicate_(section, input.task, item);

  var others = section.items.filter(function (it) { return it !== item; });
  var place = String(p.place || 'keep');
  var at = placeIndex_(others, place);
  var unchanged = at < 0 && sameTask_(taskSnapshot_(item, dayIdx), {
    id: item.id,
    task: input.task,
    days: dayIdx.map(function (i) { return !!input.picked[DAY_SHORT[i]]; }),
    weeks: input.weeks,
    photo: input.photo,
    reference: input.reference,
    active: active
  });
  if (unchanged) return { id: item.id, section: section.name, message: 'Nothing changed on ' + item.id + '.' };

  // New spot: renumber Order with it there, then move the row next to its new neighbors.
  // dest = the row it goes in front of, counted before the move (0 = it stays put).
  var dest = 0;
  if (at > -1) {
    var list = others.slice();
    list.splice(at, 0, item);
    writeOrders_(tab, list);
    if (others.length) {
      dest = place.indexOf('after:') === 0 ? others[at - 1].row + 1
        : place === 'top' ? minRow_(others)
        : maxRow_(others) + 1;
    }
  }

  // Patch this row's task columns; everything else on the row is written back as it was
  var range = tab.sheet.getRange(item.row, 1, 1, tab.width);
  var row = range.getValues()[0];
  var fields = taskFields_(tab, input);
  fields['Active'] = active;
  Object.keys(fields).forEach(function (h) { row[colOrThrow_(tab, h)] = fields[h]; });
  writeRows_(tab, item.row, [row.map(function (v) {
    return typeof v === 'string' && v.charAt(0) === '=' ? "'" + v : v; // never a formula
  })]);
  if (dest && dest !== item.row && dest !== item.row + 1) tab.sheet.moveRows(range, dest);
  SpreadsheetApp.flush();

  return {
    id: item.id,
    section: section.name,
    message: active
      ? 'Saved ' + item.id + ' (' + whenText_(checklist, input) + '). ' + formNote_(checklist)
      : 'Saved ' + item.id + '. It\'s turned off, so it comes off the form after the next morning rebuild. Turn it back on here anytime.'
  };
}

// Same task settings? (Compares the fields the pop-up shows.)
function sameTask_(a, b) {
  if (!a || !b) return false;
  return ['id', 'task', 'days', 'weeks', 'photo', 'reference', 'active'].every(function (k) {
    return JSON.stringify(a[k]) === JSON.stringify(b[k]);
  });
}
