/**
 * ============================================================
 * CHECKLIST SYSTEM - Station QR Codes
 * ============================================================
 * For checklists with "Station QR codes" = TRUE on the Checklists
 * tab (the Daily Facilities Walk). A QR code posted at each
 * station (position) opens today's form with the station picked
 * and the person's name filled in.
 *
 * Each QR carries a code unique to its station, and the form has
 * a "Station code" question the QR fills in. A submission whose
 * code doesn't match its station is logged but doesn't count
 * (06_Form_Submit.gs), so a station is only checked off by
 * someone who scanned it there.
 *
 * The QR links go to the photo web app, which never changes:
 *   ?k=<Photo upload key>&list=<Checklist ID>&station=<position>&t=<code>
 * The form's own question IDs change with every daily rebuild, so
 * the rebuild saves them for the station page (savePrefill_).
 * Print the codes from Checklists > Print station QR codes.
 * Changing "Photo upload key" retires them.
 */

var PREFILL_NAME = 'PREFILLNAME';  // placeholders used to find each question's entry ID
var PREFILL_CODE = 'PREFILLCODE';
var CODE_LETTERS = '23456789ABCDEFGHJKLMNPQRSTUVWXYZ'; // 32 characters, no 0/O/1/I

function menuStationQr() {
  var cfg = loadSettings_();
  var lists = loadModel_().checklists.filter(function (c) { return c.stationQr && c.perPosition; }); // print before going live
  var html;
  if (!cfg.photosOn) {
    html = '<p>Set "Photo upload page" and "Photo upload key" in Settings first. Station QR codes use the same web app.</p>';
  } else if (!lists.length) {
    html = '<p>No checklist has "Station QR codes" checked on the Checklists tab.</p>';
  } else {
    html = lists.map(function (c) {
      return '<p><b>' + esc_(c.name) + '</b>' + (c.active ? '' : ' (not active yet: check Active once these are posted)') + '<br><a href="' + esc_(qrSheetLink_(cfg, c.id)) + '" target="_blank">Open the printable QR codes</a></p>';
    }).join('') + '<p style="color:#666">Post each one at its station. Reprint after renaming, adding or removing a station. ' +
      'Keep this link to managers: it shows every station\'s code.</p>';
  }
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(
    '<div style="font-family:Arial,sans-serif;font-size:14px;line-height:1.5">' + html + '</div>').setWidth(480).setHeight(280), 'Station QR codes');
}

// -- Links and codes ------------------------------------------

function webAppLink_(cfg, params) {
  var query = Object.keys(params).map(function (k) { return k + '=' + encodeURIComponent(params[k]); }).join('&');
  return cfg.photoPage + (cfg.photoPage.indexOf('?') > -1 ? '&' : '?') + 'k=' + encodeURIComponent(cfg.photoKey) + '&' + query;
}

function stationLink_(cfg, checklistId, position) {
  return webAppLink_(cfg, { list: checklistId, station: position, t: stationCode_(checklistId, position) });
}

// The print page shows every station's code, so it needs a token only the menu hands out
function qrSheetLink_(cfg, checklistId) {
  return webAppLink_(cfg, { qr: checklistId, a: stationCode_(checklistId, '(print)') });
}

// Six letters and digits unique to one station, from a secret kept in Script Properties
function stationCode_(checklistId, position) {
  var sig = Utilities.computeHmacSha256Signature(checklistId + '|' + posKey_(position), stationSecret_());
  var code = '';
  for (var i = 0; i < 6; i++) code += CODE_LETTERS.charAt((sig[i] + 256) % 32);
  return code;
}

function stationSecret_() {
  var props = PropertiesService.getScriptProperties();
  var secret = props.getProperty('STATION_SECRET');
  if (secret) return secret;
  return withLock_(function () { // made once; two first-time callers must not make two
    secret = props.getProperty('STATION_SECRET');
    if (!secret) {
      secret = Utilities.getUuid();
      props.setProperty('STATION_SECRET', secret);
    }
    return secret;
  });
}

// -- Rebuild: today's entry IDs -------------------------------

// Remembers which entry IDs today's form uses for name, station and code, so the station
// page can open it filled in. If Google won't say, the page opens the plain form and shows
// the code to type instead.
function savePrefill_(form, checklist, leaderQ, positionQ, codeQ, firstPosition) {
  var props = PropertiesService.getScriptProperties();
  try {
    var url = form.createResponse()
      .withItemResponse(leaderQ.createResponse(PREFILL_NAME))
      .withItemResponse(positionQ.createResponse(firstPosition))
      .withItemResponse(codeQ.createResponse(PREFILL_CODE))
      .toPrefilledUrl();
    var ids = { base: url.split('?')[0] };
    url.replace(/(entry\.\d+)=([^&]*)/g, function (all, entry, value) {
      if (value === PREFILL_NAME) ids.leader = entry;
      else if (value === PREFILL_CODE) ids.code = entry;
      else ids.position = entry;
    });
    if (!ids.leader || !ids.position || !ids.code) throw new Error('Unexpected prefilled link: ' + url);
    props.setProperty('PREFILL_' + checklist.id, JSON.stringify(ids));
  } catch (err) {
    props.deleteProperty('PREFILL_' + checklist.id);
    reportError_('Station QR links', err);
  }
}

// -- Web app pages --------------------------------------------

// What a station's QR code opens
function stationPage_(cfg, params) {
  var checklist = loadModel_().byId[String(params.list || '')];
  var station = checklist && checklist.active && checklist.stationQr && checklist.perPosition &&
    requiredPositions_(checklist).filter(function (p) { return p.key === posKey_(params.station); })[0];
  if (!station || String(params.t || '').toUpperCase() !== stationCode_(checklist.id, station.name)) {
    return simplePage_('This QR code is out of date. Ask a manager to print new ones (Checklists > Print station QR codes).');
  }
  if (!checklist.formLink) {
    return simplePage_(checklist.name + ' doesn\'t have a form yet. It\'s made at the next morning rebuild.');
  }
  var t = HtmlService.createTemplateFromFile('StationPage');
  t.data = JSON.stringify({
    checklist: checklist.name,
    station: station.name,
    code: stationCode_(checklist.id, station.name),
    formLink: checklist.formLink,
    prefill: JSON.parse(PropertiesService.getScriptProperties().getProperty('PREFILL_' + checklist.id) || 'null')
  }).replace(/</g, '\\u003c');
  return t.evaluate().setTitle(station.name).addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// The printable sheet: one large QR code per station
function qrSheetPage_(cfg, params) {
  var checklist = loadModel_().byId[String(params.qr || '')];
  if (!checklist || !checklist.stationQr || !checklist.perPosition ||
      String(params.a || '') !== stationCode_(checklist.id, '(print)')) {
    return simplePage_('This print link is out of date. Open it again from Checklists > Print station QR codes.');
  }
  var t = HtmlService.createTemplateFromFile('QrSheet');
  t.data = JSON.stringify({
    checklist: checklist.name,
    stations: requiredPositions_(checklist).map(function (p) {
      return { name: p.name, url: stationLink_(cfg, checklist.id, p.name) };
    })
  }).replace(/</g, '\\u003c');
  return t.evaluate().setTitle(checklist.name + ' QR codes');
}

function simplePage_(message) {
  return HtmlService.createHtmlOutput('<p style="font-family:Arial,sans-serif;font-size:18px;padding:24px">' + esc_(message) + '</p>')
    .setTitle('Checklist').addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
