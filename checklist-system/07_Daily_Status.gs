/**
 * ============================================================
 * CHECKLIST SYSTEM - Daily Status
 * ============================================================
 * One row per expected submission per day:
 *   Pending        - not submitted, not yet late
 *   Complete       - submitted before Late after
 *   Late           - past Late after, still not submitted
 *   Completed late - submitted after Late after
 *   Missed         - still not submitted when the business day ended
 *
 * Due by / Late after are copied onto each row as text when the
 * row is created, so a mid-day edit to Checklists changes
 * tomorrow, not today.
 */

// Creates any of today's rows that don't exist yet. Safe to re-run.
function ensureDailyStatus_(cfg, dateKey) {
  var model = loadModel_();
  withLock_(function () {
    var tab = readTab_(TABS.status);
    var c = statusCols_(tab);
    var have = {};
    tab.display.forEach(function (r) {
      have[toDateKey_(r[c.date]) + '|' + r[c.checklist] + '|' + posKey_(r[c.position])] = true;
    });

    var rows = [];
    model.checklists.forEach(function (checklist) {
      if (!scheduledOn_(checklist, dateKey)) return;
      if (checklist.dueMin === null || checklist.lateMin === null) return; // Validate reports this
      var positions = checklist.perPosition ? requiredPositions_(checklist).map(function (p) { return p.name; }) : [''];
      positions.forEach(function (pos) {
        if (have[dateKey + '|' + checklist.id + '|' + posKey_(pos)]) return;
        rows.push({
          'Date': dateKey,
          'Checklist ID': checklist.id,
          'Position': pos,
          'Due by': checklist.dueText,
          'Late after': checklist.lateText,
          'Status': 'Pending'
        });
      });
    });
    appendRows_(tab, rows);
    SpreadsheetApp.flush();
  });
}

function statusCols_(tab) {
  return {
    date: colOrThrow_(tab, 'Date'),
    checklist: colOrThrow_(tab, 'Checklist ID'),
    position: colOrThrow_(tab, 'Position'),
    due: colOrThrow_(tab, 'Due by'),
    late: colOrThrow_(tab, 'Late after'),
    status: colOrThrow_(tab, 'Status'),
    by: colOrThrow_(tab, 'Submitted by'),
    reminder: colOrThrow_(tab, 'Reminder sent'),
    escalation: colOrThrow_(tab, 'Escalation sent')
  };
}

// Caller must hold the lock. Non-positional checklists match on date + checklist only.
function findStatusRow_(dateKey, checklist, position) {
  var tab = readTab_(TABS.status);
  var c = statusCols_(tab);
  for (var i = 0; i < tab.display.length; i++) {
    var r = tab.display[i];
    if (r[c.checklist] !== checklist.id || toDateKey_(r[c.date]) !== dateKey) continue;
    if (checklist.perPosition && posKey_(r[c.position]) !== posKey_(position)) continue;
    return { tab: tab, index: i, status: r[c.status].trim(), lateText: r[c.late] };
  }
  return null;
}

// Every 15 minutes: close out past days as Missed, flip today's overdue rows to Late,
// and (Live mode only) send heads-ups and escalations.
function checkStatuses_(cfg) {
  var now = new Date();
  var current = businessMoment_(now, cfg);

  withLock_(function () {
    var tab = readTab_(TABS.status);
    var c = statusCols_(tab);
    var todayRows = [];

    tab.display.forEach(function (r, i) {
      var dateKey = toDateKey_(r[c.date]);
      var status = r[c.status].trim();
      if (!dateKey) return;

      // Business day is over: anything still open is Missed (no email; it goes in the summary)
      if (dateKey < current.dateKey) {
        if (status === 'Pending' || status === 'Late') setCell_(tab, i, 'Status', 'Missed');
        return;
      }
      if (dateKey !== current.dateKey) return;

      var row = {
        index: i,
        dateKey: dateKey,
        checklistId: r[c.checklist],
        position: r[c.position],
        dueText: r[c.due],
        lateText: r[c.late],
        dueMin: parseTimeToMinutes_(r[c.due]),
        lateMin: parseTimeToMinutes_(r[c.late]),
        status: status,
        reminderSent: r[c.reminder] !== '',
        escalationSent: r[c.escalation] !== ''
      };
      if (status === 'Pending' && row.lateMin !== null && current.minutes >= bizMinutes_(row.lateMin, cfg)) {
        setCell_(tab, i, 'Status', 'Late');
        row.status = 'Late';
      }
      todayRows.push(row);
    });

    if (cfg.live && todayRows.length) sendDueAlerts_(cfg, tab, todayRows, current);
    SpreadsheetApp.flush();
  });
}
