/**
 * ============================================================
 * CHECKLIST SYSTEM - Checklist Model
 * ============================================================
 * Reads Checklists, Positions and Items (by header name) into
 * plain objects and answers "what's scheduled on this date?".
 * Everything is read as displayed text - see 02_Helpers.gs.
 */

function loadModel_() {
  var model = { checklists: [], byId: {}, items: [], positionRows: [] };

  // Checklists
  var cl = readTab_(TABS.checklists);
  var c = {
    id: colOrThrow_(cl, 'Checklist ID'),
    name: colOrThrow_(cl, 'Checklist name'),
    area: colOrThrow_(cl, 'Area'),
    due: colOrThrow_(cl, 'Due by'),
    late: colOrThrow_(cl, 'Late after'),
    perPos: colOrThrow_(cl, 'One submission per position?'),
    active: colOrThrow_(cl, 'Active'),
    formId: colOrThrow_(cl, 'Form ID'),
    formLink: colOrThrow_(cl, 'Form link')
  };
  var clDays = DAY_SHORT.map(function (d) { return cl.col(d); }); // -1 = no column (e.g. Sun)
  var clRotation = cl.col('Rotation length (weeks)');               // optional columns
  var clRotationStart = cl.col('Rotation start');
  var clStationQr = cl.col('Station QR codes');
  var clLanguage = cl.col('Language');

  cl.display.forEach(function (r, i) {
    var id = r[c.id].trim();
    if (!id) return;
    var rotationText = clRotation > -1 ? r[clRotation].trim() : '';
    var rotationStartText = clRotationStart > -1 ? r[clRotationStart].trim() : '';
    var checklist = {
      row: i + 2,
      id: id,
      name: r[c.name].trim() || id,
      area: r[c.area].trim().toUpperCase(),
      days: clDays.map(function (k) { return k > -1 && isTrue_(r[k]); }),
      dueText: r[c.due].trim(),
      lateText: r[c.late].trim(),
      dueMin: parseTimeToMinutes_(r[c.due]),
      lateMin: parseTimeToMinutes_(r[c.late]),
      perPosition: isTrue_(r[c.perPos]),
      active: isTrue_(r[c.active]),
      formId: r[c.formId].trim(),
      formLink: r[c.formLink].trim(),
      rotationText: rotationText,
      rotationStartText: rotationStartText,
      rotationLength: /^\d+$/.test(rotationText) && Number(rotationText) > 0 ? Number(rotationText) : 0,
      rotationStart: toDateKey_(rotationStartText),
      stationQr: clStationQr > -1 && isTrue_(r[clStationQr]), // positions are checked off by QR code (14_Station_QR.gs)
      languageText: clLanguage > -1 ? r[clLanguage].trim() : '',
      language: languageCode_(clLanguage > -1 ? r[clLanguage] : ''), // 'en' or 'es': the form's language (16_Languages.gs)
      positions: [],
      items: []
    };
    model.checklists.push(checklist);
    if (!model.byId[id]) model.byId[id] = checklist;
  });

  // Positions
  var pt = readTab_(TABS.positions);
  var p = {
    id: colOrThrow_(pt, 'Checklist ID'),
    order: colOrThrow_(pt, 'Order'),
    name: colOrThrow_(pt, 'Position'),
    required: colOrThrow_(pt, 'Required each night')
  };
  pt.display.forEach(function (r, i) {
    var id = r[p.id].trim();
    var name = r[p.name].trim();
    if (!id || !name) return;
    var pos = { row: i + 2, checklistId: id, name: name, key: posKey_(name), order: orderOf_(r[p.order]), required: isTrue_(r[p.required]) };
    model.positionRows.push(pos);
    if (model.byId[id]) model.byId[id].positions.push(pos);
  });

  // Items
  var it = readTab_(TABS.items);
  var t = {
    id: colOrThrow_(it, 'Item ID'),
    checklist: colOrThrow_(it, 'Checklist ID'),
    position: colOrThrow_(it, 'Position / section'),
    order: colOrThrow_(it, 'Order'),
    task: colOrThrow_(it, 'Task'),
    photo: colOrThrow_(it, 'Photo'),
    reference: colOrThrow_(it, 'Reference material'),
    active: colOrThrow_(it, 'Active')
  };
  var itDays = DAY_SHORT.map(function (d) { return it.col(d); });
  var itWeeks = it.col('Rotation weeks'); // optional columns
  var itAnswer = it.col('Answer');
  var itSpanish = it.col('Spanish task');
  it.display.forEach(function (r, i) {
    var id = r[t.id].trim();
    var task = r[t.task].trim();
    if (!id && !task) return;
    var weeksText = itWeeks > -1 ? r[itWeeks].trim() : '';
    var weeks = parseWeeks_(weeksText);
    var item = {
      row: i + 2,
      id: id,
      checklistId: r[t.checklist].trim(),
      positionName: r[t.position].trim(),
      positionKey: posKey_(r[t.position]),
      order: orderOf_(r[t.order]),
      task: task,
      days: itDays.map(function (k) { return k > -1 && isTrue_(r[k]); }),
      weeksText: weeksText,
      weeks: weeks || [],           // [] = every week
      weeksBad: weeks === null,     // Validate flags it
      photo: /^email$/i.test(r[t.photo].trim()),
      answerText: itAnswer > -1 ? r[itAnswer].trim() : '',
      taskEs: itSpanish > -1 ? r[itSpanish].trim() : '',
      typeIn: itAnswer > -1 && /^type in$/i.test(r[itAnswer].trim()), // a typed answer instead of Complete / Could not complete
      reference: r[t.reference].trim(),
      active: isTrue_(r[t.active])
    };
    model.items.push(item);
    if (model.byId[item.checklistId]) model.byId[item.checklistId].items.push(item);
  });

  return model;
}

function posKey_(name) {
  return String(name || '').replace(/^'/, '').replace(/\s+/g, ' ').trim().toLowerCase();
}

// "1, 4" -> [1, 4]; "" -> []; null if anything isn't a whole number
function parseWeeks_(text) {
  var parts = String(text || '').split(/[,;\s]+/).filter(function (s) { return s; });
  if (parts.some(function (s) { return !/^\d+$/.test(s); })) return null;
  return parts.map(Number);
}

function orderOf_(v) {
  var n = Number(String(v).trim());
  return String(v).trim() === '' || isNaN(n) ? 99999 : n;
}

// Active, checked for that weekday, and not a store-closed date (Settings "Closed dates")
function scheduledOn_(checklist, dateKey, cfg) {
  if (cfg && cfg.closed[dateKey]) return false;
  return checklist.active && checklist.days[dayIndex_(dateKey)];
}

// Required positions in Order
function requiredPositions_(checklist) {
  return checklist.positions
    .filter(function (p) { return p.required; })
    .sort(function (a, b) { return a.order - b.order || a.row - b.row; });
}

// Rotating checklists (e.g. a different set of Sunday tasks each week for 6 weeks): which
// week dateKey is in, 1..length. Week 1 starts on "Rotation start" and each week is 7 days.
// 0 when the checklist doesn't rotate.
function rotationWeek_(checklist, dateKey) {
  var n = checklist.rotationLength;
  if (!n || !checklist.rotationStart) return 0;
  var weeks = Math.floor((dayNumber_(dateKey) - dayNumber_(checklist.rotationStart)) / 7);
  return ((weeks % n) + n) % n + 1;
}

// Active tasks checked for that weekday (and that rotation week), sorted by position (or section) then Order.
function itemsOn_(checklist, dateKey) {
  var day = dayIndex_(dateKey);
  var week = rotationWeek_(checklist, dateKey);
  var rank = {};
  if (checklist.perPosition) {
    requiredPositions_(checklist).forEach(function (p, i) { rank[p.key] = i; });
  } else {
    checklist.items.forEach(function (it) {
      if (rank[it.positionKey] === undefined) rank[it.positionKey] = Object.keys(rank).length;
    });
  }
  return checklist.items
    .filter(function (it) { return it.active && it.days[day] && (!it.weeks.length || it.weeks.indexOf(week) > -1); })
    .sort(function (a, b) {
      var ra = rank[a.positionKey] === undefined ? 999 : rank[a.positionKey];
      var rb = rank[b.positionKey] === undefined ? 999 : rank[b.positionKey];
      return ra - rb || a.order - b.order || a.row - b.row;
    });
}

// Short name for subjects: "Positional Closing Checklist" -> "Positional Closing"
function shortName_(checklist) {
  return checklist.name.replace(/\s+checklist$/i, '');
}
