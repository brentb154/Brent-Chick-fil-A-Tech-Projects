/**
 * ============================================================
 * CHECKLIST SYSTEM - Add a Task (Checklists > Add a task)
 * ============================================================
 * A pop-up for adding a task to the Items tab without typing a
 * row by hand. It picks the next Item ID for that checklist,
 * inserts the row with that position's tasks (or wherever it was
 * put) so the tab stays grouped, and renumbers Order around it.
 *
 * The same pop-up edits tasks (13_Edit_Item.gs); the checks and
 * placement code below are shared.
 *
 * Changes show up on the forms at the next morning rebuild, or
 * right away with Checklists > Rebuild forms now.
 */

var PHOTO_MARK = 'Email';               // the Photo column value 03_Model.gs reads as a photo task
var DAY_DISPLAY = [1, 2, 3, 4, 5, 6, 0]; // pop-up shows Mon..Sat, then Sun

function menuAddItem() {
  openTaskDialog_('add');
}

function openTaskDialog_(mode) {
  var data = taskPageData_();
  data.mode = mode;
  var t = HtmlService.createTemplateFromFile('AddItem');
  t.data = JSON.stringify(data).replace(/</g, '\\u003c');
  SpreadsheetApp.getUi().showModalDialog(t.evaluate().setWidth(660).setHeight(740), 'Checklist tasks');
}

// Every checklist with its positions (or sections) and their tasks, in sheet order
function taskPageData_() {
  var model = loadModel_();
  var tab = openTab_(TABS.items);
  var dayIdx = DAY_DISPLAY.filter(function (i) { return tab.col(DAY_SHORT[i]) > -1; });
  var today = dateKey_(new Date(), getSS_().getSpreadsheetTimeZone());
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
          rotation: rotationInfo_(c, today),
          sections: sectionsOf_(c).map(function (s) {
            return {
              name: s.name,
              note: s.note,
              tasks: s.items.map(function (it) { return taskSnapshot_(it, dayIdx); })
            };
          })
        };
      })
  };
}

// One task as the pop-up shows it. Edit sends it back to prove the row hasn't changed since.
function taskSnapshot_(it, dayIdx) {
  return {
    id: it.id,
    task: it.task,
    days: dayIdx.map(function (i) { return it.days[i]; }),
    weeks: it.weeks,
    photo: it.photo,
    reference: it.reference,
    active: it.active
  };
}

// Rotating checklist: { length: 6, next: 'Sun 10/11', week: 1 } for its next day. Else null.
function rotationInfo_(c, today) {
  if (!c.rotationLength || !c.rotationStart) return null;
  var info = { length: c.rotationLength, next: '', week: 0 };
  for (var k = 0; k < 7; k++) {
    var d = addDays_(today, k);
    if (c.days[dayIndex_(d)]) {
      info.next = shortLabel_(d);
      info.week = rotationWeek_(c, d);
      break;
    }
  }
  return info;
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

// p = { checklistId, section, task, days: ['Mon', ...], weeks: [1, 4], photo, reference, place }
// place: 'end' (default), 'top', or 'after:<Item ID>'
function serverAddItem(p) {
  return saveTask_(function () { return insertItem_(p || {}); });
}

// Runs a save under the script lock, then adds fresh pop-up data to its result
function saveTask_(fn) {
  var result;
  try {
    result = withLock_(fn);
  } catch (err) {
    if (/lock timeout/i.test(err.message)) throw new Error('The sheet is busy saving something else. Try again in a minute.');
    throw err;
  }
  result.data = taskPageData_();
  return result;
}

function insertItem_(p) {
  var model = loadModel_();
  var checklist = checklistFor_(model, p.checklistId);
  var section = sectionFor_(checklist, p.section);
  var tab = openTab_(TABS.items);
  var input = taskInput_(p, checklist, tab);
  checkDuplicate_(section, input.task, null);

  // Where it goes: under the chosen task, above the section, or after the section's last row
  // (else the checklist's last row, else the last task on the tab)
  var list = section.items.slice();
  var place = String(p.place || 'end');
  var at = placeIndex_(list, place);
  if (at < 0) at = list.length;
  var anchorRow = place.indexOf('after:') === 0 ? list[at - 1].row
    : place === 'top' && list.length ? minRow_(list) - 1
    : maxRow_(list) || maxRow_(checklist.items) || maxRow_(model.items) || 1;

  list.splice(at, 0, null);
  writeOrders_(tab, list);
  var id = nextItemId_(model, checklist);
  var row = taskFields_(tab, input);
  row['Item ID'] = id;
  row['Checklist ID'] = checklist.id;
  row['Position / section'] = section.name;
  row['Order'] = at + 1;
  row['Active'] = true;

  var sheet = tab.sheet;
  var newRow = anchorRow + 1;
  sheet.insertRowAfter(anchorRow);
  var source = anchorRow > 1 ? anchorRow : newRow + 1; // a task row next to it, never the header
  if (source <= sheet.getLastRow()) {
    var from = sheet.getRange(source, 1, 1, tab.width);
    var to = sheet.getRange(newRow, 1, 1, tab.width);
    from.copyTo(to, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false); // checkboxes and colors
    from.copyTo(to, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
  }
  writeRows_(tab, newRow, [toRow_(tab, row)]);
  SpreadsheetApp.flush();

  return {
    id: id,
    section: section.name,
    message: 'Added ' + id + ' to ' + shortName_(checklist) + ' – ' + section.name + ' (' + whenText_(checklist, input) + '). ' +
      formNote_(checklist)
  };
}

// -- Shared with Edit -----------------------------------------

function checklistFor_(model, id) {
  var checklist = model.byId[oneLine_(id)];
  if (!checklist) throw new Error('Pick a checklist.');
  return checklist;
}

// Positional checklists only take positions from the Positions tab. Others can start a new section.
function sectionFor_(checklist, name) {
  var key = posKey_(name);
  var section = sectionsOf_(checklist).filter(function (s) { return posKey_(s.name) === key; })[0];
  if (!key || (!section && checklist.perPosition)) throw new Error(checklist.perPosition ? 'Pick a position.' : 'Pick a section.');
  return section || { name: oneLine_(name), items: [] };
}

// Checks the typed task, days and rotation weeks, and returns them ready to write
function taskInput_(p, checklist, tab) {
  var task = oneLine_(p.task);
  if (!task) throw new Error('Type the task.');

  var dayCols = DAY_SHORT.filter(function (d) { return tab.col(d) > -1; });
  var picked = {};
  [].concat(p.days || []).forEach(function (d) { picked[d] = true; });
  var days = DAY_DISPLAY.map(function (i) { return DAY_SHORT[i]; })
    .filter(function (d) { return picked[d] && dayCols.indexOf(d) > -1; });
  if (!days.length) throw new Error('Pick at least one day.');
  if (!days.some(function (d) { return checklist.days[DAY_SHORT.indexOf(d)]; })) {
    throw new Error(checklist.name + ' doesn\'t run on ' + days.join(', ') + '.');
  }

  // Rotation weeks: none (or all) checked = every week
  var weeks = [];
  if (checklist.rotationLength) {
    [].concat(p.weeks || []).forEach(function (w) {
      w = Number(w);
      if (w % 1 === 0 && w >= 1 && w <= checklist.rotationLength && weeks.indexOf(w) < 0) weeks.push(w);
    });
    weeks.sort(function (a, b) { return a - b; });
    if (weeks.length === checklist.rotationLength) weeks = [];
    if (weeks.length && tab.col('Rotation weeks') < 0) throw new Error('Add a "Rotation weeks" column to the Items tab first.');
  }

  return { task: task, reference: oneLine_(p.reference), photo: !!p.photo, dayCols: dayCols, picked: picked, days: days, weeks: weeks };
}

// The columns a task's settings live in, by header
function taskFields_(tab, input) {
  var row = { 'Task': input.task, 'Photo': input.photo ? PHOTO_MARK : '', 'Reference material': input.reference };
  input.dayCols.forEach(function (d) { row[d] = !!input.picked[d]; });
  if (tab.col('Rotation weeks') > -1) row['Rotation weeks'] = input.weeks.join(', ');
  return row;
}

function checkDuplicate_(section, task, self) {
  var same = section.items.filter(function (it) {
    return it !== self && oneLine_(it.task).toLowerCase() === task.toLowerCase();
  })[0];
  if (same) {
    throw new Error('That task is already on ' + section.name + ' (' + same.id + ')' +
      (same.active ? '.' : ', turned off. Check its Active box on the Items tab instead.'));
  }
}

// 'end', 'top', 'keep' or 'after:<Item ID>' -> index in `list` (the section's tasks, minus the
// one being placed). -1 = keep it where it is.
function placeIndex_(list, place) {
  place = String(place || 'end');
  if (place === 'keep') return -1;
  if (place === 'top') return 0;
  if (place.indexOf('after:') === 0) {
    var id = place.slice(6);
    for (var i = 0; i < list.length; i++) if (list[i].id === id) return i + 1;
    throw new Error('The task you picked to put it after was moved or removed. Close this and open it again.');
  }
  return list.length;
}

// Numbers a section 1, 2, 3... in list order (null = the new task). Writes only the Order cells
// that change, in one range write. Uses the rows the model read, so call it before inserting or
// moving rows.
function writeOrders_(tab, list) {
  var changes = [];
  list.forEach(function (it, i) { if (it && it.order !== i + 1) changes.push({ row: it.row, order: i + 1 }); });
  if (!changes.length) return;
  var rows = changes.map(function (c) { return c.row; });
  var first = Math.min.apply(null, rows);
  var range = tab.sheet.getRange(first, colOrThrow_(tab, 'Order') + 1, Math.max.apply(null, rows) - first + 1, 1);
  var orders = range.getValues();
  changes.forEach(function (c) { orders[c.row - first][0] = c.order; });
  range.setValues(orders);
}

// "Mon, Wed", "every day", "Sun, weeks 1, 5", "Sun, every week"
function whenText_(checklist, input) {
  var running = DAY_SHORT.filter(function (d, i) { return checklist.days[i]; });
  var all = running.length > 1 && running.every(function (d) { return input.picked[d]; });
  var text = all ? 'every day' : input.days.join(', ');
  if (checklist.rotationLength) {
    text += !input.weeks.length ? ', every week' : ', week' + (input.weeks.length > 1 ? 's ' : ' ') + input.weeks.join(', ');
  }
  return text;
}

function formNote_(checklist) {
  return checklist.active
    ? 'The form updates after the next morning rebuild. To update today\'s form now, use Checklists > Rebuild forms now.'
    : checklist.name + ' isn\'t active, so it won\'t be on a form until it is.';
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

function minRow_(items) {
  return items.reduce(function (min, it) { return Math.min(min, it.row); }, Infinity);
}

// Typed text as one tidy line
function oneLine_(v) {
  return String(v === null || v === undefined ? '' : v).replace(/\s+/g, ' ').trim();
}
