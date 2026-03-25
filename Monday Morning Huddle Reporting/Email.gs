/**
 * Schedule Variance Analyzer — Two-Tier Email System
 *
 * Director Tier: full detail — all flags, employee roll-up, OT breakdown,
 *   missed clock-outs, chronic offenders with history.
 * Manager Tier: summary only — absences, chronic lateness, top OT offenders.
 *   Clean enough for a team meeting reference.
 */

function sendTwoTierEmails(ss, allRows, crossLocOT, chronicFlags, locationData, weeks, cfg) {
  var sheetUrl = ss.getUrl();
  var locNames = locationData.map(function(l) { return l.name; }).join(' + ');
  var weekLabels = weeks.map(function(w) { return w.label; }).join(', ');
  var subject = 'Schedule Variance — ' + locNames + ' — ' + weekLabels;

  if (cfg.DIRECTOR_EMAILS.length) {
    var dirHtml = buildDirectorEmail_(allRows, crossLocOT, chronicFlags, locNames, weekLabels, sheetUrl, cfg);
    cfg.DIRECTOR_EMAILS.forEach(function(email) {
      MailApp.sendEmail({ to: email, subject: subject, htmlBody: dirHtml });
    });
  }

  if (cfg.MANAGER_EMAILS.length) {
    var mgrHtml = buildManagerEmail_(allRows, crossLocOT, chronicFlags, locNames, weekLabels, cfg);
    cfg.MANAGER_EMAILS.forEach(function(email) {
      MailApp.sendEmail({ to: email, subject: subject + ' (Summary)', htmlBody: mgrHtml });
    });
  }
}

/* ── Director Email: Full Detail ── */

function buildDirectorEmail_(allRows, crossLocOT, chronicFlags, locNames, weekLabels, sheetUrl, cfg) {
  var totalEmp = allRows.length;
  var totalMatched = 0, totalAbsent = 0, totalMidnight = 0, netVar = 0, totalSched = 0;
  allRows.forEach(function(r) {
    totalMatched += r.matchedDays || r.workedDays || 0;
    totalAbsent  += r.absentDays || 0;
    totalMidnight += r.midnightCount || 0;
    netVar += r.totalVar || 0;
    totalSched += r.scheduledDays || 0;
  });
  var adherence = totalSched > 0 ? Math.round(totalMatched / totalSched * 1000) / 10 : 0;

  var h = emailWrap_();
  h += emailHeader_('Schedule Variance Report', locNames + ' — ' + weekLabels);

  // Quick stats
  h += '<table style="width:100%;border-collapse:collapse;margin:16px 0">';
  h += statRow_('Employees Analyzed', totalEmp);
  h += statRow_('Schedule Adherence', adherence + '%');
  h += statRow_('Absences', totalAbsent, totalAbsent > 0 ? '#DC2626' : '#059669');
  h += statRow_('Missed Clock-Outs', totalMidnight, totalMidnight > 0 ? '#D97706' : '#059669');
  h += statRow_('Net Variance', formatVariance(netVar));
  h += statRow_('Cross-Location OT', crossLocOT.totalOTHours + ' hrs');
  h += '</table>';

  // Missed clock-outs
  var midnighters = allRows.filter(function(r) { return (r.midnightCount || 0) > 0; });
  if (midnighters.length) {
    h += sectionHeader_('Missed Clock-Outs', '#D97706');
    h += '<ul style="font-size:13px;line-height:1.8">';
    midnighters.forEach(function(r) {
      var days = '';
      if (r.dayList) {
        days = ' — ' + r.dayList.filter(function(d) { return d.midnightFlag; })
          .map(function(d) { return formatFriendlyDate(d.date); }).join(', ');
      }
      h += '<li><strong>' + esc_(r.name) + '</strong>' + days + ' (' + r.midnightCount + 'x)</li>';
    });
    h += '</ul>';
  }

  // Absences
  var absentees = allRows.filter(function(r) { return (r.absentDays || 0) > 0; });
  if (absentees.length) {
    h += sectionHeader_('Absences', '#DC2626');
    h += '<ul style="font-size:13px;line-height:1.8">';
    absentees.forEach(function(r) {
      var days = '';
      if (r.dayList) {
        days = ' — ' + r.dayList.filter(function(d) { return d.status === 'absent'; })
          .map(function(d) { return formatFriendlyDate(d.date); }).join(', ');
      }
      h += '<li><strong>' + esc_(r.name) + '</strong>' + days + ' (' + r.absentDays + ' day(s))</li>';
    });
    h += '</ul>';
  }

  // Ongoing concerns (chronic)
  if (chronicFlags && chronicFlags.length) {
    h += sectionHeader_('Ongoing Concerns', '#7C3AED');
    h += '<p style="font-size:12px;color:#888;margin-bottom:8px">Flagged in ' + cfg.CHRONIC_TRIGGER + '+ of the last ' + cfg.CHRONIC_WINDOW + ' weeks</p>';
    h += '<ul style="font-size:13px;line-height:1.8">';
    chronicFlags.forEach(function(f) {
      h += '<li><strong>' + esc_(f.name) + '</strong> — ' + esc_(f.reason) + ' (' + f.weeksTriggered + ' weeks)</li>';
    });
    h += '</ul>';
  }

  // Chronic lateness
  var lateOnes = allRows.filter(function(r) {
    var pat = typeof r.pattern === 'string' ? r.pattern : (r.pattern && r.pattern.label || '');
    return pat === 'Late In' || pat === 'Works Less';
  });
  if (lateOnes.length) {
    h += sectionHeader_('Chronic Lateness', '#DC2626');
    h += '<ul style="font-size:13px;line-height:1.8">';
    lateOnes.forEach(function(r) {
      var pat = typeof r.pattern === 'string' ? r.pattern : (r.pattern && r.pattern.label || '');
      h += '<li><strong>' + esc_(r.name) + '</strong> — ' + pat + ', ' + formatVariance(r.lateIn) + ' total late</li>';
    });
    h += '</ul>';
  }

  // High variance
  var highVar = allRows.filter(function(r) { return Math.abs(r.totalVar || 0) >= cfg.HIGH_VARIANCE_THRESHOLD; })
    .sort(function(a, b) { return Math.abs(b.totalVar) - Math.abs(a.totalVar); });
  if (highVar.length) {
    h += sectionHeader_('High Variance (>' + cfg.HIGH_VARIANCE_THRESHOLD + ' min)', '#7C3AED');
    h += '<ul style="font-size:13px;line-height:1.8">';
    highVar.slice(0, 10).forEach(function(r) {
      h += '<li><strong>' + esc_(r.name) + '</strong> — ' + formatVariance(r.totalVar) + '</li>';
    });
    h += '</ul>';
  }

  // OT summary
  h += sectionHeader_('Overtime Summary', '#1a1a2e');
  h += '<table style="width:100%;border-collapse:collapse;margin:8px 0">';
  Object.keys(crossLocOT.byLocation).forEach(function(loc) {
    var d = crossLocOT.byLocation[loc];
    h += '<tr><td style="padding:4px 12px;font-weight:bold;font-size:13px">' + esc_(loc) + '</td>';
    h += '<td style="padding:4px 12px;font-size:13px">Sched: ' + d.schedHours + ' hrs</td>';
    h += '<td style="padding:4px 12px;font-size:13px">Actual: ' + d.actualHours + ' hrs</td></tr>';
  });
  h += '<tr style="border-top:2px solid #e5e7eb"><td style="padding:6px 12px;font-weight:bold;font-size:13px;color:#7C3AED">Reconciled OT Total</td>';
  h += '<td style="padding:6px 12px;font-size:13px;font-weight:bold">' + crossLocOT.totalOTHours + ' hrs actual</td>';
  h += '<td style="padding:6px 12px;font-size:13px">' + crossLocOT.totalScheduledOTHours + ' hrs scheduled</td></tr>';
  h += '</table>';

  // Top OT offenders
  var otList = [];
  Object.keys(crossLocOT.employees).forEach(function(name) {
    var e = crossLocOT.employees[name];
    if (e.otHours > 0) otList.push({ name: name, ot: e.otHours, total: e.totalHours, multi: e.isMultiLocation });
  });
  otList.sort(function(a, b) { return b.ot - a.ot; });
  if (otList.length) {
    h += '<p style="font-size:12px;font-weight:bold;margin:12px 0 6px;color:#555">Top OT Employees:</p>';
    h += '<ul style="font-size:13px;line-height:1.8">';
    otList.slice(0, 5).forEach(function(e) {
      h += '<li><strong>' + esc_(e.name) + '</strong> — ' + e.ot + ' hrs OT (' + e.total + ' hrs total)' +
        (e.multi ? ' <span style="color:#7C3AED">★ multi-location</span>' : '') + '</li>';
    });
    h += '</ul>';
  }

  // Positive callouts
  var stars = allRows.filter(function(r) {
    var pat = typeof r.pattern === 'string' ? r.pattern : (r.pattern && r.pattern.label || '');
    return pat === 'On Track' && (r.reliability || 0) >= 95;
  });
  if (stars.length) {
    h += sectionHeader_('Positive Callouts', '#059669');
    h += '<ul style="font-size:13px;line-height:1.8">';
    stars.slice(0, 5).forEach(function(r) {
      h += '<li><strong>' + esc_(r.name) + '</strong> — On Track, ' + r.reliability + '% reliability</li>';
    });
    h += '</ul>';
  }

  h += emailFooter_(sheetUrl);
  h += '</div>';
  return h;
}

/* ── Manager Email: Summary Only ── */

function buildManagerEmail_(allRows, crossLocOT, chronicFlags, locNames, weekLabels, cfg) {
  var totalEmp = allRows.length;
  var totalAbsent = 0, totalMidnight = 0, totalSched = 0, totalMatched = 0;
  allRows.forEach(function(r) {
    totalAbsent  += r.absentDays || 0;
    totalMidnight += r.midnightCount || 0;
    totalSched += r.scheduledDays || 0;
    totalMatched += r.matchedDays || r.workedDays || 0;
  });
  var adherence = totalSched > 0 ? Math.round(totalMatched / totalSched * 1000) / 10 : 0;

  var h = emailWrap_();
  h += emailHeader_('Schedule Variance Summary', locNames + ' — ' + weekLabels);

  // Quick stats
  h += '<table style="width:100%;border-collapse:collapse;margin:16px 0">';
  h += statRow_('Employees', totalEmp);
  h += statRow_('Adherence', adherence + '%');
  h += statRow_('Absences', totalAbsent, totalAbsent > 0 ? '#DC2626' : '#059669');
  h += statRow_('OT Hours', crossLocOT.totalOTHours + ' hrs');
  h += '</table>';

  // Absences
  var absentees = allRows.filter(function(r) { return (r.absentDays || 0) > 0; });
  if (absentees.length) {
    h += sectionHeader_('Absences', '#DC2626');
    h += '<ul style="font-size:13px;line-height:1.8">';
    absentees.forEach(function(r) {
      h += '<li><strong>' + esc_(r.name) + '</strong> — ' + r.absentDays + ' day(s) absent</li>';
    });
    h += '</ul>';
  }

  // Ongoing concerns
  if (chronicFlags && chronicFlags.length) {
    h += sectionHeader_('Ongoing Concerns', '#7C3AED');
    h += '<ul style="font-size:13px;line-height:1.8">';
    chronicFlags.forEach(function(f) {
      h += '<li><strong>' + esc_(f.name) + '</strong> — ' + esc_(f.reason) + '</li>';
    });
    h += '</ul>';
  }

  // Chronic lateness
  var lateOnes = allRows.filter(function(r) {
    var pat = typeof r.pattern === 'string' ? r.pattern : (r.pattern && r.pattern.label || '');
    return pat === 'Late In' || pat === 'Works Less';
  });
  if (lateOnes.length) {
    h += sectionHeader_('Chronically Late', '#DC2626');
    h += '<ul style="font-size:13px;line-height:1.8">';
    lateOnes.slice(0, 5).forEach(function(r) {
      var pat = typeof r.pattern === 'string' ? r.pattern : (r.pattern && r.pattern.label || '');
      h += '<li><strong>' + esc_(r.name) + '</strong> — ' + pat + '</li>';
    });
    h += '</ul>';
  }

  // Top OT offenders
  var otList = [];
  Object.keys(crossLocOT.employees).forEach(function(name) {
    var e = crossLocOT.employees[name];
    if (e.otHours > 0) otList.push({ name: name, ot: e.otHours });
  });
  otList.sort(function(a, b) { return b.ot - a.ot; });
  if (otList.length) {
    h += sectionHeader_('Top OT Offenders', '#D97706');
    h += '<ul style="font-size:13px;line-height:1.8">';
    otList.slice(0, 5).forEach(function(e) {
      h += '<li><strong>' + esc_(e.name) + '</strong> — ' + e.ot + ' hrs OT</li>';
    });
    h += '</ul>';
  }

  h += '<hr style="border:none;border-top:1px solid #e5e7eb;margin:24px 0">';
  h += '<p style="font-size:11px;color:#999">Full detail available in the Director report. Contact your Operations Director for the complete analysis.</p>';
  h += '</div>';
  return h;
}

/* ── Email HTML helpers ── */

function emailWrap_() {
  return '<div style="font-family:Arial,sans-serif;max-width:620px;margin:0 auto;color:#333;padding:20px">';
}

function emailHeader_(title, subtitle) {
  return '<h1 style="font-size:20px;color:#1a1a2e;border-bottom:3px solid #e3342f;padding-bottom:8px;margin-bottom:4px">' + title + '</h1>' +
    '<p style="color:#666;font-size:13px;margin-bottom:16px">' + subtitle + '</p>';
}

function emailFooter_(sheetUrl) {
  var h = '<hr style="border:none;border-top:1px solid #e5e7eb;margin:24px 0">';
  if (sheetUrl) {
    h += '<p style="font-size:12px;margin-bottom:8px"><a href="' + sheetUrl + '" style="color:#e3342f;text-decoration:none">View full report in Google Sheets →</a></p>';
  }
  h += '<p style="font-size:11px;color:#999">Generated automatically by Schedule Variance Analyzer</p>';
  return h;
}

function statRow_(label, value, color) {
  return '<tr>' +
    '<td style="padding:6px 12px;font-size:13px;font-weight:bold;color:#555;border-bottom:1px solid #eee">' + label + '</td>' +
    '<td style="padding:6px 12px;font-size:14px;font-weight:bold;border-bottom:1px solid #eee;' +
    (color ? 'color:' + color : '') + '">' + value + '</td></tr>';
}

function sectionHeader_(title, color) {
  return '<h3 style="font-size:14px;color:' + color + ';margin:20px 0 8px;border-left:3px solid ' + color + ';padding-left:10px">' + title + '</h3>';
}

function esc_(s) {
  return String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}
