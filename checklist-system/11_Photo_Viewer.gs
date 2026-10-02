/**
 * ============================================================
 * CHECKLIST SYSTEM - Photo Viewer (Checklists > View photos)
 * ============================================================
 * A pop-up inside the sheet for flipping through any day's photos.
 * For each checklist and position it shows the status, who
 * submitted it and when, tasks marked "Could not complete", notes,
 * and the photos (tap one for full size).
 *
 * It runs as whoever opens it, so they need view access to the
 * "Checklist Photos" Drive folder. The store account owns it; share
 * that folder with managers as Viewer. New photos inherit the
 * sharing. Only files listed on the Photos tab are ever opened.
 */

var VIEWER_THUMBS_PER_CALL = 30;

function menuViewPhotos() {
  var cfg = loadSettings_();
  var t = HtmlService.createTemplateFromFile('PhotoViewer');
  t.data = JSON.stringify({
    today: businessMoment_(new Date(), cfg).dateKey,
    keepDays: cfg.keepPhotosDays
  }).replace(/</g, '\\u003c');
  SpreadsheetApp.getUi().showModalDialog(t.evaluate().setWidth(1100).setHeight(800), 'Checklist photos');
}

// One row per photo on the Photos tab: { id, dateKey, checklistId, position, leader, time }
function photoRows_(cfg) {
  var tab = readTab_(TABS.photos, true);
  var c = {
    at: colOrThrow_(tab, 'Uploaded at'),
    date: colOrThrow_(tab, 'Business date'),
    id: colOrThrow_(tab, 'Checklist ID'),
    pos: colOrThrow_(tab, 'Position'),
    leader: colOrThrow_(tab, 'Leader'),
    file: colOrThrow_(tab, 'File')
  };
  var rows = [];
  tab.display.forEach(function (r, i) {
    var m = String(r[c.file]).match(/\/d\/([\w-]+)/);
    if (!m) return;
    var at = tab.values[i][c.at];
    rows.push({
      id: m[1],
      dateKey: toDateKey_(r[c.date]),
      checklistId: r[c.id],
      position: r[c.pos],
      leader: r[c.leader],
      time: at instanceof Date ? Utilities.formatDate(at, cfg.tz, 'h:mm a') : ''
    });
  });
  return rows;
}

// Everything for one day: a group per checklist + position, in Checklists / Positions order
function serverPhotoDay(dateKey) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(dateKey))) throw new Error('Pick a date.');
  var cfg = loadSettings_();
  var model = loadModel_();
  var groups = [];
  var byKey = {};

  // Non-positional checklists are one group, whatever section name their tasks use
  function group(checklistId, position) {
    var c = model.byId[checklistId];
    var pos = c && c.perPosition ? position : '';
    var k = checklistId + '|' + posKey_(pos);
    if (!byKey[k]) {
      byKey[k] = {
        name: c ? c.name : checklistId,
        position: pos,
        status: '',
        by: '',
        time: '',
        photoTasks: c ? itemsOn_(c, dateKey).filter(function (it) {
          return it.photo && (!c.perPosition || it.positionKey === posKey_(pos));
        }).length : 0,
        notDone: [],
        notes: [],
        photos: []
      };
      groups.push(byKey[k]);
    }
    return byKey[k];
  }

  var st = readTab_(TABS.status, true);
  var sc = statusCols_(st);
  var atCol = colOrThrow_(st, 'Submitted at');
  st.display.forEach(function (r, i) {
    if (toDateKey_(r[sc.date]) !== dateKey) return;
    var g = group(r[sc.checklist], r[sc.position]);
    var at = st.values[i][atCol];
    g.status = r[sc.status].trim();
    g.by = r[sc.by];
    g.time = at instanceof Date ? Utilities.formatDate(at, cfg.tz, 'h:mm a') : '';
  });

  var re = readTab_(TABS.results);
  var rc = {
    date: colOrThrow_(re, 'Business date'), id: colOrThrow_(re, 'Checklist ID'),
    pos: colOrThrow_(re, 'Position'), task: colOrThrow_(re, 'Task'), result: colOrThrow_(re, 'Result')
  };
  re.display.forEach(function (r) {
    if (toDateKey_(r[rc.date]) === dateKey && r[rc.result] === ANSWER_NOT_DONE) group(r[rc.id], r[rc.pos]).notDone.push(r[rc.task]);
  });

  var su = readTab_(TABS.submissions);
  var uc = {
    date: colOrThrow_(su, 'Business date'), id: colOrThrow_(su, 'Checklist ID'),
    pos: colOrThrow_(su, 'Position'), leader: colOrThrow_(su, 'Leader name'), notes: colOrThrow_(su, 'Notes')
  };
  su.display.forEach(function (r) {
    if (toDateKey_(r[uc.date]) === dateKey && r[uc.notes]) group(r[uc.id], r[uc.pos]).notes.push(r[uc.leader] + ': ' + r[uc.notes]);
  });

  var photos = cfg.photosOn ? photoRows_(cfg) : [];
  photos.forEach(function (p) {
    if (p.dateKey === dateKey) group(p.checklistId, p.position).photos.push({ id: p.id, leader: p.leader, time: p.time });
  });

  return { dateKey: dateKey, label: longLabel_(dateKey), groups: groups, canSeePhotos: canSeePhotos_() };
}

// Can the person viewing open the photo folder? (It belongs to the store account.)
function canSeePhotos_() {
  var id = PropertiesService.getScriptProperties().getProperty('PHOTO_ROOT_ID');
  if (!id) return true; // nothing uploaded yet
  try {
    DriveApp.getFolderById(id).getName();
    return true;
  } catch (err) {
    return false;
  }
}

// Small previews: Drive's thumbnail when it has one, else the photo itself
function serverPhotoThumbs(ids) {
  var allowed = {};
  photoRows_(loadSettings_()).forEach(function (p) { allowed[p.id] = true; });
  return (ids || []).slice(0, VIEWER_THUMBS_PER_CALL).map(function (id) {
    if (!allowed[id]) return { id: id, gone: true };
    try {
      var file = DriveApp.getFileById(id);
      if (file.isTrashed()) return { id: id, gone: true };
      var blob = file.getThumbnail() || file.getBlob();
      return { id: id, mime: blob.getContentType(), data: Utilities.base64Encode(blob.getBytes()) };
    } catch (err) {
      return { id: id, gone: true };
    }
  });
}

// Full-size photo for the click-to-view overlay
function serverPhotoFull(id) {
  var known = photoRows_(loadSettings_()).some(function (p) { return p.id === id; });
  if (!known) throw new Error('That photo is not on the Photos tab.');
  var file = DriveApp.getFileById(id);
  if (file.isTrashed()) throw new Error('That photo was deleted (older than the keep window).');
  var blob = file.getBlob();
  return { mime: blob.getContentType(), data: Utilities.base64Encode(blob.getBytes()) };
}
