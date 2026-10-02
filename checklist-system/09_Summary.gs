/**
 * ============================================================
 * CHECKLIST SYSTEM - Daily Summary and Weekly Scorecard
 * ============================================================
 * Checked once a day after "Daily summary at", in both alert modes.
 *   Daily  - covers yesterday, and ONLY goes out when something needs
 *            attention: a checklist or position Missed, a task marked
 *            "Could not complete", a note, or (when photo uploads are on)
 *            a position with no photos or a photo count far below usual.
 *            A clean night sends nothing.
 *   Weekly - on Settings "Weekly scorecard day" (Tuesday), covers the
 *            previous Mon-Sat. Sent every week.
 */

var STATUS_COLORS = { 'Complete': '#1e7e34', 'Completed late': '#a15c00', 'Late': '#b3261e', 'Missed': '#b3261e', 'Pending': '#555' };
var TABLE_STYLE = 'border-collapse:collapse;font-size:13px;margin:4px 0 16px';
var CELL_STYLE = 'border:1px solid #ddd;padding:4px 8px;text-align:left;vertical-align:top';

function sendDailySummary_(cfg, todayKey) {
  var to = emailList_(cfg.get('Daily summary recipients'));
  if (!to.length) return;

  var day = addDays_(todayKey, -1);
  var data = loadSummaryData_(cfg);
  var parts = [];
  var subjects = [];
  var types = [];

  var photo = photoIssues_(cfg, data, day);
  var issues = dayIssues_(data, day, photo);
  if (issues.count) {
    parts.push(dailyHtml_(data, day, photo));
    subjects.push('Checklist issues – ' + shortLabel_(day) + ' (' + issues.text + ')');
    types.push('Daily summary');
  }
  if (dayIndex_(todayKey) === cfg.scorecardDay) {
    var end = lastSaturday_(todayKey);
    var start = addDays_(end, -5);
    parts.push(weeklyHtml_(data, start, end));
    subjects.push('Weekly scorecard ' + shortLabel_(start) + '–' + shortLabel_(end));
    types.push('Weekly scorecard');
  }
  if (!parts.length) return;

  MailApp.sendEmail({
    to: to.join(','),
    subject: subjects.join(' + '),
    body: 'This summary is formatted for an email app that shows HTML.',
    htmlBody: '<div style="font-family:Arial,sans-serif;color:#222">' + parts.join('<hr>') +
      '<p style="font-size:12px;color:#777">Checklist Master: ' + esc_(getSS_().getUrl()) + '</p></div>'
  });
  types.forEach(function (t) { logAlert_(day, '', '', t, to.join(', ')); });
}

// Most recent Saturday before todayKey, where the last full Mon-Sat week ends
function lastSaturday_(todayKey) {
  var back = (dayIndex_(todayKey) + 1) % 7 || 7;
  return addDays_(todayKey, -back);
}

// What makes the daily email worth sending. Completed late still counts as done.
function dayIssues_(data, day, photo) {
  var missed = data.status.filter(function (r) { return r.dateKey === day && r.status === 'Missed'; }).length;
  var notDone = data.results.filter(function (r) { return r.dateKey === day && r.result === ANSWER_NOT_DONE; }).length;
  var notes = data.subs.filter(function (s) { return s.dateKey === day && s.notes; }).length;
  var bits = [];
  if (missed) bits.push(missed + ' missed');
  if (notDone) bits.push(notDone + ' not completed');
  if (notes) bits.push(notes + (notes === 1 ? ' note' : ' notes'));
  if (photo.noPhotos.length) bits.push(photo.noPhotos.length + ' no photos');
  if (photo.low.length) bits.push('low photos');
  return { count: missed + notDone + notes + photo.noPhotos.length + photo.low.length, text: bits.join(', ') };
}

function loadSummaryData_(cfg) {
  var st = readTab_(TABS.status, true);
  var sc = statusCols_(st);
  var atCol = colOrThrow_(st, 'Submitted at');
  var status = st.display.map(function (r, i) {
    var at = st.values[i][atCol];
    return {
      dateKey: toDateKey_(r[sc.date]),
      checklistId: r[sc.checklist],
      position: r[sc.position],
      status: r[sc.status].trim(),
      time: at instanceof Date ? Utilities.formatDate(at, cfg.tz, 'h:mm a') : r[atCol],
      by: r[sc.by]
    };
  }).filter(function (r) { return r.status !== 'Closed'; }); // closed days don't count either way

  var su = readTab_(TABS.submissions);
  var s = function (h) { return colOrThrow_(su, h); };
  var subs = su.display.map(function (r) {
    return {
      id: r[s('Submission ID')], dateKey: toDateKey_(r[s('Business date')]), checklistId: r[s('Checklist ID')],
      position: r[s('Position')], leader: r[s('Leader name')], notes: r[s('Notes')], onTime: r[s('On time?')],
      expected: Number(r[s('Items expected')]) || 0, complete: Number(r[s('Items complete')]) || 0
    };
  });

  var re = readTab_(TABS.results);
  var x = function (h) { return colOrThrow_(re, h); };
  var results = re.display.map(function (r) {
    return {
      submissionId: r[x('Submission ID')], dateKey: toDateKey_(r[x('Business date')]), checklistId: r[x('Checklist ID')],
      position: r[x('Position')], itemId: r[x('Item ID')], task: r[x('Task')], result: r[x('Result')]
    };
  });

  var photos = [];
  if (cfg.photosOn) {
    var ph = readTab_(TABS.photos);
    var y = function (h) { return colOrThrow_(ph, h); };
    photos = ph.display.map(function (r) {
      return { dateKey: toDateKey_(r[y('Business date')]), checklistId: r[y('Checklist ID')], position: r[y('Position')] };
    });
  }

  return { model: loadModel_(), status: status, subs: subs, results: results, photos: photos };
}

function dailyHtml_(data, day, photo) {
  var html = '<h2 style="margin-bottom:4px">Checklists – ' + esc_(longLabel_(day)) + '</h2>';
  var dayRows = data.status.filter(function (r) { return r.dateKey === day; });

  // Each checklist, in Checklists tab order
  orderedIds_(data.model, dayRows).forEach(function (id) {
    var checklist = data.model.byId[id];
    var rows = dayRows.filter(function (r) { return r.checklistId === id; });
    var inCount = rows.filter(isSubmitted_).length;
    var heading = (checklist ? checklist.name : id) +
      (checklist && checklist.perPosition ? ' – ' + inCount + ' of ' + rows.length + ' positions in' : '');
    html += '<h3 style="margin:12px 0 4px">' + esc_(heading) + '</h3>' + table_(['Position', 'Status', 'Submitted', 'Leader'],
      rows.map(function (r) {
        var color = STATUS_COLORS[r.status] || '#555';
        return [esc_(r.position || '–'), '<b style="color:' + color + '">' + esc_(r.status) + '</b>', esc_(r.time), esc_(r.by)];
      }));
  });

  // Could not complete + notes
  var leaderBySub = {};
  data.subs.forEach(function (s) { leaderBySub[s.id] = s.leader; });
  var cnc = data.results.filter(function (r) { return r.dateKey === day && r.result === ANSWER_NOT_DONE; });
  html += '<h3 style="margin:12px 0 4px">Marked "Could not complete" (' + cnc.length + ')</h3>';
  html += cnc.length ? table_(['Checklist', 'Position', 'Task', 'Leader'], cnc.map(function (r) {
    return [esc_(nameOf_(data.model, r.checklistId)), esc_(r.position), esc_(r.task), esc_(leaderBySub[r.submissionId])];
  })) : '<p>None.</p>';

  var noted = data.subs.filter(function (s) { return s.dateKey === day && s.notes; });
  html += '<h3 style="margin:12px 0 4px">Notes</h3>';
  html += noted.length ? table_(['Checklist', 'Position', 'Leader', 'Note'], noted.map(function (s) {
    return [esc_(nameOf_(data.model, s.checklistId)), esc_(s.position || '–'), esc_(s.leader), esc_(s.notes)];
  })) : '<p>None.</p>';

  html += photoHtml_(data, photo);
  html += '<h3 style="margin:12px 0 4px">Last 7 days</h3>' + ratesTable_(data, addDays_(day, -6), day);
  return html;
}

function weeklyHtml_(data, start, end) {
  var html = '<h2 style="margin-bottom:4px">Weekly scorecard – ' + esc_(shortLabel_(start) + ' to ' + shortLabel_(end)) + '</h2>';
  html += '<h3 style="margin:12px 0 4px">By checklist</h3>' + ratesTable_(data, start, end);

  // By leader. Names are free text: match trimmed and case-insensitive.
  var leaders = {};
  data.subs.forEach(function (s) {
    if (s.dateKey < start || s.dateKey > end) return;
    var key = String(s.leader).replace(/\s+/g, ' ').trim().toLowerCase() || '(no name)';
    var g = leaders[key] = leaders[key] || { name: titleCase_(key), subs: 0, onTime: 0, expected: 0, complete: 0 };
    g.subs++;
    if (s.onTime === 'Yes') g.onTime++;
    g.expected += s.expected;
    g.complete += s.complete;
  });
  var leaderRows = Object.keys(leaders).map(function (k) { return leaders[k]; })
    .sort(function (a, b) { return b.subs - a.subs || (a.name < b.name ? -1 : 1); });
  html += '<h3 style="margin:12px 0 4px">By leader</h3>';
  html += leaderRows.length ? table_(['Leader', 'Submissions', 'On time', 'Tasks complete'], leaderRows.map(function (g) {
    return [esc_(g.name), g.subs, pct_(g.onTime, g.subs), pct_(g.complete, g.expected)];
  })) : '<p>No submissions.</p>';

  // Tasks most often not completed
  var counts = {};
  data.results.forEach(function (r) {
    if (r.dateKey < start || r.dateKey > end || r.result !== ANSWER_NOT_DONE) return;
    var key = r.itemId || r.task;
    var g = counts[key] = counts[key] || { count: 0, task: r.task, checklistId: r.checklistId, position: r.position };
    g.count++;
  });
  var top = Object.keys(counts).map(function (k) { return counts[k]; })
    .sort(function (a, b) { return b.count - a.count; }).slice(0, 10);
  html += '<h3 style="margin:12px 0 4px">Most often "Could not complete"</h3>';
  html += top.length ? table_(['Times', 'Checklist', 'Position', 'Task'], top.map(function (g) {
    return [g.count, esc_(nameOf_(data.model, g.checklistId)), esc_(g.position), esc_(g.task)];
  })) : '<p>None.</p>';
  return html;
}

// Submitted = Complete or Completed late. On time = Complete.
function ratesTable_(data, start, end) {
  var by = {};
  var rows = data.status.filter(function (r) { return r.dateKey >= start && r.dateKey <= end; });
  rows.forEach(function (r) {
    var g = by[r.checklistId] = by[r.checklistId] || { expected: 0, submitted: 0, onTime: 0 };
    g.expected++;
    if (isSubmitted_(r)) g.submitted++;
    if (r.status === 'Complete') g.onTime++;
  });
  var ids = orderedIds_(data.model, rows);
  if (!ids.length) return '<p>No checklists scheduled.</p>';
  return table_(['Checklist', 'Submitted', 'On time'], ids.map(function (id) {
    var g = by[id];
    return [esc_(nameOf_(data.model, id)), g.submitted + ' of ' + g.expected + ' (' + pct_(g.submitted, g.expected) + ')', pct_(g.onTime, g.expected)];
  }));
}

// Photos section: who sent none, anything far below usual, and each checklist's count vs. usual
function photoHtml_(data, photo) {
  var ids = orderedIds_(data.model, Object.keys(photo.stats).map(function (id) { return { checklistId: id }; }));
  if (!ids.length && !photo.noPhotos.length) return '';
  var html = '<h3 style="margin:12px 0 4px">Photos</h3>';
  if (photo.noPhotos.length) html += '<p style="color:#b3261e"><b>No photos from:</b> ' + esc_(photo.noPhotos.join(', ')) + '</p>';
  photo.low.forEach(function (l) {
    html += '<p style="color:#b3261e"><b>Low:</b> ' + esc_(l.name) + ' had ' + l.count + ' photos; usually ' +
      Math.round(l.mean) + ' &plusmn; ' + Math.round(l.sd) + '.</p>';
  });
  return html + table_(['Checklist', 'Photos', 'Usual'], ids.map(function (id) {
    var s = photo.stats[id];
    var usual = s.mean === null ? 'building history (10 days)' : Math.round(s.mean) + ' &plusmn; ' + Math.round(s.sd);
    return [esc_(nameOf_(data.model, id)), s.count, usual];
  }));
}

function isSubmitted_(r) {
  return r.status === 'Complete' || r.status === 'Completed late';
}

// Checklist IDs that appear in rows, in Checklists tab order
function orderedIds_(model, rows) {
  var present = {};
  rows.forEach(function (r) { present[r.checklistId] = true; });
  var ids = model.checklists.map(function (c) { return c.id; }).filter(function (id) { return present[id]; });
  Object.keys(present).forEach(function (id) { if (ids.indexOf(id) < 0) ids.push(id); });
  return ids;
}

function nameOf_(model, id) {
  return model.byId[id] ? model.byId[id].name : id;
}

function pct_(a, b) {
  return b ? Math.round(100 * a / b) + '%' : '–';
}

function titleCase_(s) {
  return s.replace(/\b\w/g, function (ch) { return ch.toUpperCase(); });
}

// cells are already-escaped HTML
function table_(headers, rows) {
  return '<table style="' + TABLE_STYLE + '"><tr>' +
    headers.map(function (h) { return '<th style="' + CELL_STYLE + ';background:#f4f4f4">' + esc_(h) + '</th>'; }).join('') + '</tr>' +
    rows.map(function (r) {
      return '<tr>' + r.map(function (cell) { return '<td style="' + CELL_STYLE + '">' + cell + '</td>'; }).join('') + '</tr>';
    }).join('') + '</table>';
}
