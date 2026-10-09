/**
 * ============================================================
 * CHECKLIST SYSTEM - Photo Uploads
 * ============================================================
 * A no-sign-in upload page (web app) for checklist pictures.
 *
 * Deploy it once from the STORE account (Apps Script editor):
 *   Deploy > New deployment > Web app
 *   Execute as: Me    Who has access: Anyone
 * Paste the /exec link into Settings "Photo upload page".
 * After a code change to this page, the store account updates it:
 *   Deploy > Manage deployments > Edit > Version: New version.
 *
 * The page only works with the Settings "Photo upload key" in its
 * link (forms and the QR code carry it). Change the key to retire
 * old links.
 *
 * Photos land in the store account's Drive:
 *   Checklist Photos / 2026-10-06 / Positional Closing Checklist / Stocker /
 * with one row per photo on the Photos tab. Each morning, date
 * folders older than "Keep photos for (days)" go to the trash.
 *
 * Uploading is two steps so each photo is quick:
 *   serverStartPhotos - once per checklist/position/name: checks the link,
 *                       reads the sheet, finds or makes the Drive folder, and
 *                       remembers it all for 6 hours under a batch ID.
 *   serverSavePhoto   - once per photo: saves the file and logs one row.
 *                       No sheet reads and no lock, so several photos save at
 *                       once and form submissions never wait behind photos.
 */

var PHOTO_ROOT_NAME = 'Checklist Photos';
var PHOTO_MAX_BYTES = 15 * 1024 * 1024;
var PHOTO_BATCH_SECONDS = 6 * 60 * 60;   // CacheService's maximum
var MSG_OLD_LINK = 'This photo link is out of date. Ask a manager for the current link.';
var MSG_NOT_SAVED = 'The photo could not be saved. A manager has been notified. Try again in a minute.';

// -- Upload page ----------------------------------------------

function doGet(e) {
  var cfg = loadSettings_();
  var params = (e && e.parameter) || {};
  var page;
  if (!cfg.photoKey || params.k !== cfg.photoKey) {
    page = HtmlService.createHtmlOutput('<p style="font-family:Arial,sans-serif;font-size:18px;padding:24px">' +
      'This link is out of date. Ask a manager for the current link or QR code.</p>');
  } else if (params.station) {
    return stationPage_(cfg, params);   // a station's QR code (14_Station_QR.gs)
  } else if (params.qr) {
    return qrSheetPage_(cfg, params);   // printable station QR codes
  } else {
    var t = HtmlService.createTemplateFromFile('PhotoPage');
    t.data = JSON.stringify(photoPageData_(cfg, params.k, params.list || '')).replace(/</g, '\\u003c');
    page = t.evaluate();
  }
  return page.setTitle('Checklist photos').addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// Checklists for the current business day (after midnight that's still last night)
function photoPageData_(cfg, key, preselect) {
  var day = businessMoment_(new Date(), cfg).dateKey;
  return {
    key: key,
    dayLabel: longLabel_(day),
    preselect: preselect,
    checklists: loadModel_().checklists.filter(function (c) { return scheduledOn_(c, day, cfg); }).map(function (c) {
      return { id: c.id, name: c.name, positions: c.perPosition ? requiredPositions_(c).map(function (p) { return p.name; }) : [] };
    })
  };
}

// -- Uploading ------------------------------------------------

// Step 1, once per checklist/position/name. p = { key, checklistId, position, leader }.
// Returns a batch ID the page sends with each photo.
function serverStartPhotos(p) {
  var cfg = loadSettings_();
  if (!p || !cfg.photoKey || p.key !== cfg.photoKey) throw new Error(MSG_OLD_LINK);
  var checklist = loadModel_().byId[p.checklistId];
  if (!checklist || !checklist.active) throw new Error('Pick a checklist.');
  var position = '';
  if (checklist.perPosition) {
    var match = requiredPositions_(checklist).filter(function (x) { return x.name === p.position; })[0];
    if (!match) throw new Error('Pick your position.');
    position = match.name;
  }
  var leader = String(p.leader || '').replace(/\s+/g, ' ').trim().slice(0, 60);
  if (!leader) throw new Error('Enter your name.');

  // Column positions are looked up once here so each photo can log without reading the sheet
  var tab = openTab_(TABS.photos);
  var cols = {
    at: colOrThrow_(tab, 'Uploaded at'),
    date: colOrThrow_(tab, 'Business date'),
    checklist: colOrThrow_(tab, 'Checklist ID'),
    position: colOrThrow_(tab, 'Position'),
    leader: colOrThrow_(tab, 'Leader'),
    file: colOrThrow_(tab, 'File')
  };
  var dateKey = businessMoment_(new Date(), cfg).dateKey;
  var folderId;
  try {
    folderId = withLock_(function () { return photoFolder_(dateKey, checklist.name, position).getId(); });
  } catch (err) {
    reportError_('Photo upload', err);
    throw new Error(MSG_NOT_SAVED);
  }

  var batchId = Utilities.getUuid();
  CacheService.getScriptCache().put('photos:' + batchId, JSON.stringify({
    folderId: folderId,
    ssId: getSS_().getId(),
    tz: cfg.tz,
    dateKey: dateKey,
    checklistId: checklist.id,
    position: position,
    leader: leader,
    cols: cols,
    width: tab.width
  }), PHOTO_BATCH_SECONDS);
  return batchId;
}

// Step 2, once per photo. data = base64 JPEG, already shrunk on the phone; number = the photo's
// number on the page. Safe to repeat: a retry after a lost reply never saves or logs a photo twice.
function serverSavePhoto(batchId, number, data) {
  var cache = CacheService.getScriptCache();
  var raw = cache.get('photos:' + batchId);
  if (!raw) throw new Error('BATCH_EXPIRED'); // the page starts a new batch and tries again
  var batch = JSON.parse(raw);
  var stepKey = 'photos:' + batchId + ':' + number;
  var step = cache.get(stepKey) || '';   // '' = not saved yet, a file URL = saved but not logged, 'logged' = done
  if (step === 'logged') return { ok: true };

  try {
    var url = step;
    if (!url) {
      var bytes = decodeJpeg_(data);
      var now = new Date();
      var name = batch.leader + ' ' + Utilities.formatDate(now, batch.tz, 'h.mm a') + ' ' + (Number(number) || 1) + '.jpg';
      url = DriveApp.getFolderById(batch.folderId).createFile(Utilities.newBlob(bytes, 'image/jpeg', name)).getUrl();
      cache.put(stepKey, url, PHOTO_BATCH_SECONDS);
    }
    // appendRow adds the row in one step, so photos saving at the same time can't overwrite each other
    SpreadsheetApp.openById(batch.ssId).getSheetByName(TABS.photos).appendRow(photoRow_(batch, new Date(), url));
    cache.put(stepKey, 'logged', PHOTO_BATCH_SECONDS);
  } catch (err) {
    if (err.unreadable) throw new Error('That photo could not be read. Try again.');
    reportError_('Photo upload', err);
    throw new Error(MSG_NOT_SAVED);
  }
  return { ok: true };
}

// Pages opened before the two-step upload send one photo per call with everything in it
function serverUploadPhoto(p) {
  return serverSavePhoto(serverStartPhotos(p), p.index, p.data);
}

// Base64 -> bytes, or an error marked unreadable if it isn't a JPEG of a sane size
function decodeJpeg_(data) {
  var bytes = [];
  try {
    bytes = Utilities.base64Decode(String(data || ''));
  } catch (err) {}
  if (bytes.length > 2 && bytes.length <= PHOTO_MAX_BYTES && (bytes[0] & 0xFF) === 0xFF && (bytes[1] & 0xFF) === 0xD8) return bytes;
  var bad = new Error('Not a JPEG');
  bad.unreadable = true;
  throw bad;
}

// Photos tab row in the sheet's column order. Text gets a leading apostrophe so Sheets keeps it
// exactly as written (otherwise "2026-10-06" becomes a date and a name starting with "=" a formula).
function photoRow_(batch, when, url) {
  var row = [];
  for (var c = 0; c < batch.width; c++) row.push('');
  row[batch.cols.at] = when;
  row[batch.cols.date] = "'" + batch.dateKey;
  row[batch.cols.checklist] = batch.checklistId;
  row[batch.cols.position] = batch.position ? "'" + batch.position : '';
  row[batch.cols.leader] = "'" + batch.leader;
  row[batch.cols.file] = url;
  return row;
}

// Upload page link with the key; checklistId preselects the checklist.
// Don't name link parameters "c" or "sid": Google reserves them and the page fails to open.
function photoLink_(cfg, checklistId) {
  if (!cfg.photosOn) return '';
  return cfg.photoPage + (cfg.photoPage.indexOf('?') > -1 ? '&' : '?') + 'k=' + encodeURIComponent(cfg.photoKey) +
    (checklistId ? '&list=' + encodeURIComponent(checklistId) : '');
}

// -- Drive folders --------------------------------------------

function photoRoot_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('PHOTO_ROOT_ID');
  if (id) {
    try {
      var existing = DriveApp.getFolderById(id);
      if (!existing.isTrashed()) return existing;
    } catch (err) {}
  }
  var root = DriveApp.createFolder(PHOTO_ROOT_NAME);
  props.setProperty('PHOTO_ROOT_ID', root.getId());
  return root;
}

function childFolder_(parent, name) {
  var found = parent.getFoldersByName(name);
  return found.hasNext() ? found.next() : parent.createFolder(name);
}

function photoFolder_(dateKey, checklistName, position) {
  var folder = childFolder_(childFolder_(photoRoot_(), dateKey), checklistName);
  return position ? childFolder_(folder, position) : folder;
}

// Moves date folders older than "Keep photos for (days)" to the trash (recoverable for 30 days).
// Only folders named like 2026-10-06 directly inside Checklist Photos are touched.
function trashOldPhotos_(cfg, today) {
  if (!cfg.keepPhotosDays) return;
  var id = PropertiesService.getScriptProperties().getProperty('PHOTO_ROOT_ID');
  if (!id) return; // nothing uploaded yet
  var cutoff = addDays_(today, -cfg.keepPhotosDays);
  var root;
  try {
    root = DriveApp.getFolderById(id);
  } catch (err) {
    return; // folder deleted by hand; the next upload makes a new one
  }
  var folders = root.getFolders();
  while (folders.hasNext()) {
    var f = folders.next();
    var name = f.getName();
    if (/^\d{4}-\d{2}-\d{2}$/.test(name) && name < cutoff) f.setTrashed(true);
  }
}

// -- Photo checks for the morning email -----------------------

// noPhotos: submitted positions (or checklists) that had photo tasks that day but uploaded none.
// low:      checklists whose count is far below their usual: below mean - k x std dev, where k is
//           Settings "Photo alert sensitivity (std devs)". Needs 10+ days of photo history,
//           using the last 28 days the checklist was submitted.
// stats:    count / usual per checklist, for the email.
function photoIssues_(cfg, data, day) {
  var out = { noPhotos: [], low: [], stats: {} };
  if (!cfg.photosOn) return out;

  var perDay = {};   // checklistId -> dateKey -> photos
  var perSlot = {};  // 'checklistId|position' -> photos on `day`
  var first = {};    // checklistId -> first date with any photo
  data.photos.forEach(function (p) {
    var byDay = perDay[p.checklistId] = perDay[p.checklistId] || {};
    byDay[p.dateKey] = (byDay[p.dateKey] || 0) + 1;
    if (!first[p.checklistId] || p.dateKey < first[p.checklistId]) first[p.checklistId] = p.dateKey;
    if (p.dateKey === day) {
      var slot = p.checklistId + '|' + posKey_(p.position);
      perSlot[slot] = (perSlot[slot] || 0) + 1;
    }
  });

  var submittedDays = {}; // checklistId -> { dateKey: true } for days with at least one submission
  data.status.forEach(function (r) {
    if (!isSubmitted_(r)) return;
    (submittedDays[r.checklistId] = submittedDays[r.checklistId] || {})[r.dateKey] = true;
    if (r.dateKey !== day) return;
    var c = data.model.byId[r.checklistId];
    if (!c) return;
    var photoTasks = itemsOn_(c, day).filter(function (it) {
      return it.photo && (!c.perPosition || it.positionKey === posKey_(r.position));
    });
    if (photoTasks.length && !perSlot[r.checklistId + '|' + posKey_(r.position)]) {
      out.noPhotos.push(c.name + (c.perPosition ? ' – ' + r.position : ''));
    }
  });

  Object.keys(submittedDays).forEach(function (id) {
    if (!submittedDays[id][day]) return;
    var count = (perDay[id] || {})[day] || 0;
    var stat = { count: count, mean: null, sd: null };
    out.stats[id] = stat;
    if (!first[id]) return;
    var history = Object.keys(submittedDays[id])
      .filter(function (d) { return d < day && d >= first[id]; }).sort().slice(-28);
    if (history.length < 10) return;
    var counts = history.map(function (d) { return (perDay[id] || {})[d] || 0; });
    var mean = counts.reduce(function (a, b) { return a + b; }, 0) / counts.length;
    var sd = Math.sqrt(counts.reduce(function (a, b) { return a + (b - mean) * (b - mean); }, 0) / (counts.length - 1));
    stat.mean = mean;
    stat.sd = sd;
    // Floor the std dev at 1 so a very steady history doesn't flag a one-photo dip.
    // Zero photos is already reported per position above.
    if (count > 0 && count < mean - cfg.photoSigma * Math.max(sd, 1)) {
      out.low.push({ name: nameOf_(data.model, id), count: count, mean: mean, sd: sd });
    }
  });
  return out;
}
