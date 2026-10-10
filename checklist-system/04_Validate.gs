/**
 * ============================================================
 * CHECKLIST SYSTEM - Validate Sheet
 * ============================================================
 * Runs before every rebuild. Blocking problems would publish a
 * broken form, so the rebuild keeps yesterday's form instead.
 * Warnings don't stop anything.
 */

function validateSheet_() {
  var out = { blocking: [], warnings: [] };
  var model, cfg;
  try {
    model = loadModel_();
    cfg = loadSettings_();
  } catch (err) {
    out.blocking.push(err.message);
    return out;
  }

  // Unique IDs
  var seenChecklist = {};
  model.checklists.forEach(function (c) {
    if (seenChecklist[c.id]) out.blocking.push('Checklists row ' + c.row + ': Checklist ID ' + c.id + ' is used more than once.');
    seenChecklist[c.id] = true;
  });

  var seenItem = {};
  model.items.forEach(function (it) {
    if (!it.id) {
      out.blocking.push('Items row ' + it.row + ': task has no Item ID.');
      return;
    }
    if (seenItem[it.id]) out.blocking.push('Items row ' + it.row + ': Item ID ' + it.id + ' is also on row ' + seenItem[it.id] + '. Item IDs must be unique.');
    else seenItem[it.id] = it.row;
    if (it.active && !it.task) out.blocking.push('Items row ' + it.row + ' (' + it.id + '): Task is blank.');

    var checklist = model.byId[it.checklistId];
    if (!checklist) {
      out.blocking.push('Items row ' + it.row + ' (' + it.id + '): Checklist ID "' + it.checklistId + '" is not on the Checklists tab.');
      return;
    }
    if (checklist.perPosition) {
      var known = checklist.positions.some(function (p) { return p.key === it.positionKey; });
      if (!known) out.blocking.push('Items row ' + it.row + ' (' + it.id + '): position "' + it.positionName + '" is not listed for ' + checklist.id + ' on the Positions tab.');
    }
    if (it.active && it.answerText && !it.typeIn) {
      out.warnings.push('Items row ' + it.row + ' (' + it.id + '): Answer "' + it.answerText + '" should be blank or "Type in". Using Complete / Could not complete.');
    }
    if (it.active && it.weeksText) {
      var n = checklist.rotationLength;
      if (!n) {
        out.blocking.push('Items row ' + it.row + ' (' + it.id + '): Rotation weeks is "' + it.weeksText + '", but ' + checklist.id + ' has no Rotation length on the Checklists tab.');
      } else if (it.weeksBad || it.weeks.some(function (w) { return w < 1 || w > n; })) {
        out.blocking.push('Items row ' + it.row + ' (' + it.id + '): Rotation weeks "' + it.weeksText + '" should be week numbers from 1 to ' + n + ', e.g. 2 or 1, 4.');
      }
    }
  });

  model.positionRows.forEach(function (p) {
    if (!model.byId[p.checklistId]) out.warnings.push('Positions row ' + p.row + ': Checklist ID "' + p.checklistId + '" is not on the Checklists tab.');
  });

  // Active checklists need times, days, and (if positional) positions
  var areas = {};
  model.checklists.forEach(function (c) {
    if (!c.active) return;
    areas[c.area] = true;
    var where = 'Checklists row ' + c.row + ' (' + c.id + '): ';
    if (c.dueMin === null) out.blocking.push(where + 'Due by "' + c.dueText + '" is not a time.');
    if (c.lateMin === null) out.blocking.push(where + 'Late after "' + c.lateText + '" is not a time.');
    if (c.dueMin !== null && c.lateMin !== null && bizMinutes_(c.lateMin, cfg) < bizMinutes_(c.dueMin, cfg)) {
      out.warnings.push(where + 'Late after (' + c.lateText + ') is earlier than Due by (' + c.dueText + ').');
    }
    if (!c.days.some(function (d) { return d; })) out.blocking.push(where + 'no days are checked.');
    if (c.perPosition && !requiredPositions_(c).length) out.blocking.push(where + 'set to one submission per position, but no required positions are on the Positions tab.');
    if (!c.area) out.warnings.push(where + 'Area is blank, so late alerts have nowhere to go.');
    if (c.rotationText || c.rotationStartText) checkRotation_(out, c, where);
    if (c.languageText && c.language === 'en' && !/^english$/i.test(c.languageText)) {
      out.warnings.push(where + 'Language "' + c.languageText + '" should be English or Spanish. Using English.');
    }
    if (c.language === 'es') {
      var untranslated = c.items.filter(function (it) { return it.active && !it.taskEs; }).map(function (it) { return it.id; });
      if (untranslated.length) out.warnings.push(where + 'the form is in Spanish, but ' + untranslated.join(', ') + ' have no "Spanish task", so they show in English.');
    }
    if (c.stationQr && !c.perPosition) out.warnings.push(where + '"Station QR codes" only works with "One submission per position?" set to TRUE.');
    if (c.stationQr && !cfg.photosOn) out.warnings.push(where + 'Station QR codes use the photo web app. Set "Photo upload page" and "Photo upload key" in Settings.');
  });

  // Settings times
  ['Business day ends at', 'Rebuild forms at', 'Daily summary at'].forEach(function (label) {
    if (parseTimeToMinutes_(cfg.get(label)) === null) out.blocking.push('Settings: "' + label + '" is "' + cfg.get(label) + '", which is not a time.');
  });
  if (cfg.rebuildMin < cfg.dayEndMin) out.warnings.push('Settings: "Rebuild forms at" is before "Business day ends at". Late closers would lose their form mid-close.');
  if (cfg.summaryMin < cfg.dayEndMin) out.warnings.push('Settings: "Daily summary at" is before "Business day ends at". The summary would go out before the day is closed.');
  if (cfg.get('Weekly scorecard day') && dayFromName_(cfg.get('Weekly scorecard day')) < 0) out.warnings.push('Settings: "Weekly scorecard day" should be a day like Tue. Using Tuesday.');

  // Photo uploads (optional): the page link and key go together
  var page = cfg.get('Photo upload page');
  var key = cfg.get('Photo upload key');
  if (page && !key) out.warnings.push('Settings: "Photo upload page" is set but "Photo upload key" is blank, so photo uploads are off.');
  if (page && !/^https:\/\/script\.google\.com\/.+\/exec$/.test(page)) out.warnings.push('Settings: "Photo upload page" should be the web app link ending in /exec.');
  if (cfg.get('Keep photos for (days)') && !(parseInt(cfg.get('Keep photos for (days)'), 10) > 0)) out.warnings.push('Settings: "Keep photos for (days)" isn\'t a number, so old photos are never deleted.');
  var sigma = cfg.get('Photo alert sensitivity (std devs)');
  if (sigma && !(parseFloat(sigma) > 0)) out.warnings.push('Settings: "Photo alert sensitivity (std devs)" isn\'t a number. Using 2.');
  if (!/^(live|digest only)$/i.test(cfg.get('Alerts mode'))) out.warnings.push('Settings: "Alerts mode" should be "Digest only" or "Live". Treating it as Digest only.');
  parseDateList_(cfg.get('Closed dates')).bad.forEach(function (s) {
    out.warnings.push('Settings: "Closed dates" has "' + s + '", which isn\'t a date. Use M/D/YYYY, e.g. 11/26/2026.');
  });
  var tzSetting = cfg.get('Time zone');
  if (tzSetting && tzSetting !== cfg.tz) out.warnings.push('Settings: Time zone says ' + tzSetting + ' but the spreadsheet is set to ' + cfg.tz + ' (File > Settings). The script uses the spreadsheet\'s.');

  // Emails
  Object.keys(areas).forEach(function (area) {
    if (!area) return;
    checkEmails_(out, area + ' escalation email', cfg.get(area + ' escalation email'), 'no one will get late alerts for ' + area + '.');
  });
  checkEmails_(out, 'Daily summary recipients', cfg.get('Daily summary recipients'), 'no daily summary will be sent.');

  return out;
}

// A rotating checklist needs a length and a start date, and every week should have tasks
function checkRotation_(out, c, where) {
  if (!c.rotationLength) out.blocking.push(where + 'Rotation length "' + c.rotationText + '" should be a whole number of weeks, e.g. 6.');
  if (!c.rotationStart) {
    out.blocking.push(where + 'Rotation start "' + c.rotationStartText + '" is not a date. Use M/D/YYYY, e.g. 10/11/2026.');
  } else if (!c.days[dayIndex_(c.rotationStart)]) {
    out.warnings.push(where + 'Rotation start (' + c.rotationStartText + ') is a ' + DAY_LONG[dayIndex_(c.rotationStart)] + ', a day this checklist doesn\'t run. Weeks still count from that date.');
  }
  for (var w = 1; w <= c.rotationLength; w++) {
    var any = c.items.some(function (it) { return it.active && (!it.weeks.length || it.weeks.indexOf(w) > -1); });
    if (!any) out.warnings.push(where + 'rotation week ' + w + ' has no active tasks, so that week\'s form would be empty.');
  }
}

function checkEmails_(out, label, text, blankMessage) {
  if (!text) {
    out.warnings.push('Settings: "' + label + '" is blank, so ' + blankMessage);
    return;
  }
  text.split(/[,;\s]+/).forEach(function (part) {
    if (part && !isEmail_(part)) out.warnings.push('Settings: "' + label + '" has "' + part + '", which doesn\'t look like an email address.');
  });
}

function showValidation_(result) {
  var html = '<div style="font-family:Arial,sans-serif;font-size:13px;line-height:1.5">';
  if (!result.blocking.length && !result.warnings.length) {
    html += '<p><b>No problems found.</b></p>';
  }
  if (result.blocking.length) {
    html += '<p style="color:#b3261e"><b>Blocking (forms will not rebuild until fixed):</b></p><ul>' +
      result.blocking.map(function (m) { return '<li>' + esc_(m) + '</li>'; }).join('') + '</ul>';
  }
  if (result.warnings.length) {
    html += '<p style="color:#8a5a00"><b>Warnings:</b></p><ul>' +
      result.warnings.map(function (m) { return '<li>' + esc_(m) + '</li>'; }).join('') + '</ul>';
  }
  html += '</div>';
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(620).setHeight(440), 'Validate sheet');
}
