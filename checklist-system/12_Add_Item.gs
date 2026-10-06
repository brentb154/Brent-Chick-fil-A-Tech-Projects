/**
 * ============================================================
 * CHECKLIST SYSTEM - Add a Task (Checklists > Add a task)
 * ============================================================
 * A pop-up for adding a task to the Items tab without typing a
 * row by hand. It picks the next Item ID for that checklist,
 * inserts the row under the chosen task (or at the end of that
 * position's tasks) so the tab stays grouped, and renumbers
 * Order for the tasks after it.
 *
 * New tasks show up on the forms at the next morning rebuild, or
 * right away with Checklists > Rebuild forms now.
 */

var PHOTO_MARK = 'Email'; // the Photo column value 03_Model.gs reads as a photo task

function menuAddItem() {
  var t = HtmlService.createTemplateFromFile('AddItem');
  t.data = JSON.stringify(addItemPageData_()).replace(/</g, '\\u003c');
  SpreadsheetApp.getUi().showModalDialog(t.evaluate().setWidth(640).setHeight(660), 'Add a task');
}

// Every checklist with its positions (or sections) and their tasks, in sheet order
function addItemPageData_() {
  var model = loadModel_();
  var tab = openTab_(TABS.items);
  var dayIdx = [];
  DAY_SHORT.forEach(function (d, i) { if (tab.col(d) > -1) dayIdx.push(i); });
  return {
    days: dayIdx.map(function (i) { return DAY_SHORT[i]; }),
    photoNote: PHOTO_SUFFIX.trim(),
    checklists: model.checklists
      .filter(function (c) { return model.byId[c.id] === c; }) // a duplicate ID row is a Validate problem
      .map(function (c) {
        return {
          id: c.id,
          name: c.name,
          active: c.active,
          perPosition: c.perPosition,
          runs: dayIdx.map(function (i) { return c.days[i]; }),
          sections: sectionsOf_(c).map(function (s) {
            return {
              name: s.name,
              note: s.note,
              tasks: s.items.map(function (it) { return { id: it.id, task: it.task, active: it.active }; })
            };
          })
        };
      })
  };
}

// Positional checklists: the Positions tab in Order. Others: the section names their tasks
// already use. Each section's tasks are in form order (Order, then row).
function sectionsOf_(checklist) {
  var sections = [];
  var byKey = {};
  function add(name, note) {
    var key = posKey_(name);
    if (!byKey[key]) {
      byKey[key] = { name: name, note: note || '', items: [] };
      sections.push(byKey[key]);
    }
    return byKey[key];
  }
  if (checklist.perPosition) {
    checklist.positions.slice().sort(byOrder_).forEach(function (p) {
      add(p.name, p.required ? '' : 'not on the form');
    });
  }
  checklist.items.forEach(function (it) {
    if (!it.positionName || (checklist.perPosition && !byKey[it.positionKey])) return; // Validate flags these
    add(it.positionName).items.push(it);
  });
  sections.forEach(function (s) { s.items.sort(byOrder_); });
  return sections;
}

function byOrder_(a, b) {
  return a.order - b.order || a.row - b.row;
}

// p = { checklistId, section, task, days: ['Mon', ...], photo, reference, afterId ('' = at the end) }
function serverAddItem(p) {
  p = p || {};
  var added;
  try {
    added = withLock_(function () { return insertItem_(p); });
  } catch (err) {
    if (/lock timeout/i.test(err.message)) throw new Error('The sheet is busy saving something else. Try again in a minute.');
    throw err;
  }
  added.data = addItemPageData_();
  return added;
}

function insertItem_(p) {
  var task = oneLine_(p.task);
  var reference = oneLine_(p.reference);
  var model = loadModel_();
  var checklist = model.byId[oneLine_(p.checklistId)];
  if (!checklist) throw new Error('Pick a checklist.');

  // Position / section: positional checklists only take positions from the Positions tab
  var key = posKey_(p.section);
  var section = sectionsOf_(checklist).filter(function (s) { return posKey_(s.name) === key; })[0];
  if (!key || (!section && checklist.perPosition)) throw new Error(checklist.perPosition ? 'Pick a position.' : 'Pick a section.');
  if (!section) section = { name: oneLine_(p.section), items: [] }; // first task in a new section

  if (!task) throw new Error('Type the task.');
  var same = section.items.filter(function (it) { return oneLine_(it.task).toLowerCase() === task.toLowerCase(); })[0];
  if (same) {
    throw new Error('That task is already on ' + section.name + ' (' + same.id + ')' +
      (same.active ? '.' : ', turned off. Check its Active box on the Items tab instead.'));
  }

  var tab = openTab_(TABS.items);
  var dayCols = DAY_SHORT.filter(function (d) { return tab.col(d) > -1; });
  var picked = {};
  [].concat(p.days || []).forEach(function (d) { picked[d] = true; });
  var days = dayCols.filter(function (d) { return picked[d]; });
  if (!days.length) throw new Error('Pick at least one day.');
  if (!days.some(function (d) { return checklist.days[DAY_SHORT.indexOf(d)]; })) {
    throw new Error(checklist.name + ' doesn\'t run on ' + days.join(', ') + '.');
  }

  // Where it goes: right after the chosen task, else after this section's last row,
  // else after this checklist's last row, else after the last task on the tab
  var list = section.items.slice();
  var at = list.length;
  if (p.afterId) {
    at = -1;
    list.forEach(function (it, i) { if (it.id === p.afterId) at = i + 1; });
    if (at < 0) throw new Error('The task you picked to put it after was moved or removed. Close this and open Add a task again.');
  }
  var anchorRow = p.afterId ? list[at - 1].row : (maxRow_(list) || maxRow_(checklist.items) || maxRow_(model.items) || 1);

  // Order: number the section 1, 2, 3... in its current form order with the new task in place,
  // so later tasks move down one. Only cells that change are written.
  list.splice(at, 0, null);
  var newOrder = at + 1;
  var changes = [];
  list.forEach(function (it, i) {
    if (it && it.order !== i + 1) changes.push({ row: it.row > anchorRow ? it.row + 1 : it.row, order: i + 1 });
  });

  var id = nextItemId_(model, checklist);
  var row = { 'Item ID': id, 'Checklist ID': checklist.id, 'Position / section': section.name, 'Order': newOrder,
    'Task': task, 'Photo': p.photo ? PHOTO_MARK : '', 'Reference material': reference, 'Active': true };
  dayCols.forEach(function (d) { row[d] = !!picked[d]; });

  var sheet = tab.sheet;
  var newRow = anchorRow + 1;
  sheet.insertRowAfter(anchorRow);
  if (anchorRow > 1) { // checkboxes and colors from the row above (never the header)
    var from = sheet.getRange(anchorRow, 1, 1, tab.width);
    var to = sheet.getRange(newRow, 1, 1, tab.width);
    from.copyTo(to, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    from.copyTo(to, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
  }
  writeRows_(tab, newRow, [toRow_(tab, row)]);

  if (changes.length) {
    var rows = changes.map(function (c) { return c.row; });
    var first = Math.min.apply(null, rows);
    var range = sheet.getRange(first, colOrThrow_(tab, 'Order') + 1, Math.max.apply(null, rows) - first + 1, 1);
    var orders = range.getValues();
    changes.forEach(function (c) { orders[c.row - first][0] = c.order; });
    range.setValues(orders);
  }
  SpreadsheetApp.flush();

  var when = days.length === dayCols.length ? 'every day' : days.join(', ');
  return {
    id: id,
    section: section.name,
    message: 'Added ' + id + ' to ' + shortName_(checklist) + ' – ' + section.name + ' (' + when + '). ' +
      (checklist.active
        ? 'It will be on the form after the next morning rebuild. To add it to today\'s form now, use Checklists > Rebuild forms now.'
        : checklist.name + ' isn\'t active, so it won\'t be on a form until it is.')
  };
}

// Next number for this checklist's ID prefix (PC-, MC-...). Item Results is checked too,
// so the ID of a deleted task is never handed out again.
function nextItemId_(model, checklist) {
  var counts = {};
  checklist.items.forEach(function (it) {
    var s = splitItemId_(it.id);
    if (s) counts[s.prefix] = (counts[s.prefix] || 0) + 1;
  });
  var prefix = Object.keys(counts).sort(function (a, b) { return counts[b] - counts[a]; })[0] || checklist.id + '-';

  var max = 0;
  function see(id) {
    var s = splitItemId_(id);
    if (s && s.prefix === prefix && s.number > max) max = s.number;
  }
  model.items.forEach(function (it) { see(it.id); });
  var results = openTab_(TABS.results);
  var col = results.col('Item ID');
  var count = results.sheet.getLastRow() - 1;
  if (col > -1 && count > 0) {
    results.sheet.getRange(2, col + 1, count, 1).getDisplayValues().forEach(function (r) { see(r[0]); });
  }

  var next = String(max + 1);
  while (next.length < 3) next = '0' + next;
  return prefix + next;
}

// "PC-094" -> { prefix: 'PC-', number: 94 }. null when it doesn't end in a number.
function splitItemId_(id) {
  var m = String(id || '').trim().match(/^(.*\D)(\d+)$/);
  return m ? { prefix: m[1], number: Number(m[2]) } : null;
}

function maxRow_(items) {
  return items.reduce(function (max, it) { return Math.max(max, it.row); }, 0);
}

// Typed text as one tidy line
function oneLine_(v) {
  return String(v === null || v === undefined ? '' : v).replace(/\s+/g, ' ').trim();
}
