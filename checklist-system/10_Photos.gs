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
 */

var PHOTO_ROOT_NAME = 'Checklist Photos';
var PHOTO_MAX_BYTES = 15 * 1024 * 1024;

// -- Upload page ----------------------------------------------

function doGet(e) {
  var cfg = loadSettings_();
  var params = (e && e.parameter) || {};
  var page;
  if (!cfg.photoKey || params.k !== cfg.photoKey) {
    page = HtmlService.createHtmlOutput('<p style="font-family:Arial,sans-serif;font-size:18px;padding:24px">' +
      'This photo link is out of date. Ask a manager for the current link or QR code.</p>');
  } else {
    var t = HtmlService.createTemplateFromFile('PhotoPage');
    t.data = JSON.stringify(photoPageData_(cfg, params.k, params.c || '')).replace(/</g, '\\u003c');
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

// Called by the upload page once per photo.
// p = { key, checklistId, position, leader, index, data (base64 JPEG, already shrunk on the phone) }
function serverUploadPhoto(p) {
  var cfg = loadSettings_();
  if (!cfg.photoKey || p.key !== cfg.photoKey) throw new Error('This photo link is out of date. Ask a manager for the current link.');
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

  var bytes = Utilities.base64Decode(String(p.data || ''));
  var isJpeg = bytes.length > 2 && (bytes[0] & 0xFF) === 0xFF && (bytes[1] & 0xFF) === 0xD8;
  if (!isJpeg || bytes.length > PHOTO_MAX_BYTES) throw new Error('That photo could not be read. Try again.');

  var now = new Date();
  var moment = businessMoment_(now, cfg);
  var name = leader + ' ' + Utilities.formatDate(now, cfg.tz, 'h.mm a') + ' ' + (Number(p.index) || 1) + '.jpg';

  // Lock only around folder creation and the log row, so photos from several phones don't queue up.
  // Saving problems (Drive permission, quota) email the admin; the team member gets a plain message.
  try {
    var folder = withLock_(function () { return photoFolder_(moment.dateKey, checklist.name, position); });
    var file = folder.createFile(Utilities.newBlob(bytes, 'image/jpeg', name));
    withLock_(function () {
      appendRows_(openTab_(TABS.photos), [{
        'Uploaded at': now,
        'Business date': moment.dateKey,
        'Checklist ID': checklist.id,
        'Position': position,
        'Leader': leader,
        'File': file.getUrl()
      }]);
      SpreadsheetApp.flush();
    });
  } catch (err) {
    reportError_('Photo upload', err);
    throw new Error('The photo could not be saved. A manager has been notified. Try again in a minute.');
  }
  return { ok: true };
}

// Upload page link with the key; checklistId preselects the checklist
function photoLink_(cfg, checklistId) {
  if (!cfg.photosOn) return '';
  return cfg.photoPage + (cfg.photoPage.indexOf('?') > -1 ? '&' : '?') + 'k=' + encodeURIComponent(cfg.photoKey) +
    (checklistId ? '&c=' + encodeURIComponent(checklistId) : '');
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
