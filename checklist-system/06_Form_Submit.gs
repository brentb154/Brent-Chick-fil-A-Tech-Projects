/**
 * ============================================================
 * CHECKLIST SYSTEM - Logging Each Submission
 * ============================================================
 * Installed as an onFormSubmit trigger on every checklist form.
 * Forms are NOT linked to response sheets (the daily rebuild
 * would add new columns every day). This writes clean rows:
 *   Submissions  - one row per submission
 *   Item Results - one row per task
 *   Daily Status - marks the matching row Complete / Completed late
 *
 * The daily rebuild deletes questions, and Google deletes their
 * stored answers with them, so these tabs are the only record.
 */

function onChecklistSubmit(e) {
  if (!isTriggerOwner_()) return;
  try {
    logSubmission_(e);
  } catch (err) {
    reportError_('Form submission', err);
  }
}

function logSubmission_(e) {
  var cfg = loadSettings_();
  var formId = formIdFromEvent_(e);
  var model = loadModel_();
  var checklist = model.checklists.filter(function (c) { return c.formId === formId; })[0];
  if (!checklist) throw new Error('Got a submission from form ' + formId + ', but no row on Checklists has that Form ID.');

  var response = e.response;
  var submittedAt = response.getTimestamp();
  var moment = businessMoment_(submittedAt, cfg);

  // Sort answers into leader / position / notes / tasks
  var leader = '';
  var position = '';
  var notes = [];
  var taskAnswers = [];
  response.getItemResponses().forEach(function (ir) {
    var item;
    try {
      item = ir.getItem();
    } catch (err) {
      return; // question deleted by a rebuild after this person opened the form
    }
    var title = item.getTitle();
    var answer = String(ir.getResponse() || '').trim();
    if (title === Q_LEADER) leader = answer;
    else if (title === Q_POSITION) position = answer;
    else if (title === Q_NOTES) { if (answer) notes.push(answer); }
    else taskAnswers.push({ questionId: String(item.getId()), task: title, result: answer });
  });
  var answeredPosition = position;
  if (!checklist.perPosition) position = '';
  var complete = taskAnswers.filter(function (a) { return a.result === ANSWER_DONE; }).length;
  var notComplete = taskAnswers.filter(function (a) { return a.result === ANSWER_NOT_DONE; }).length;
  var submissionId = 'S-' + Utilities.getUuid().slice(0, 8).toUpperCase();

  // Form Map is read inside the lock so a rebuild that's rewriting it can't be caught half-done
  withLock_(function () {
    var formMap = loadFormMap_(checklist.id);
    // Unmapped questions (form opened before a rebuild) are kept by title with no Item ID
    var answers = taskAnswers.map(function (a) {
      var mapped = formMap.byQuestion[a.questionId];
      return {
        itemId: mapped ? mapped.itemId : '',
        position: mapped ? mapped.position : answeredPosition,
        date: mapped ? mapped.date : '',
        task: a.task,
        result: a.result
      };
    });

    // Expected = tasks on the form version they answered (for that position)
    var expected = answers.length;
    var builtFor = answers.filter(function (a) { return a.date; })[0];
    if (builtFor) {
      expected = formMap.rows.filter(function (r) {
        return r.date === builtFor.date && (!checklist.perPosition || posKey_(r.position) === posKey_(position));
      }).length;
    }

    var statusMatch = findStatusRow_(moment.dateKey, checklist, position);
    var lateText = statusMatch ? statusMatch.lateText : checklist.lateText;
    var lateMin = parseTimeToMinutes_(lateText);
    var onTime = lateMin === null ? '' : (moment.minutes < bizMinutes_(lateMin, cfg) ? 'Yes' : 'No');

    appendRows_(openTab_(TABS.submissions), [{
      'Submission ID': submissionId,
      'Submitted at': submittedAt,
      'Business date': moment.dateKey,
      'Checklist ID': checklist.id,
      'Position': position,
      'Leader name': leader,
      'Items expected': expected,
      'Items complete': complete,
      'Items not complete': notComplete,
      'Notes': notes.join(' | '),
      'On time?': onTime,
      'Form response ID': response.getId()
    }]);

    appendRows_(openTab_(TABS.results), answers.map(function (a) {
      return {
        'Submission ID': submissionId,
        'Business date': moment.dateKey,
        'Checklist ID': checklist.id,
        'Position': a.position,
        'Item ID': a.itemId,
        'Task': a.task,
        'Result': a.result
      };
    }));

    // First submission sets the status; later duplicates are logged above but don't change it.
    // Missed is included: a 3:59 AM submission can land just after the 4:00 AM sweep marked the row Missed.
    var open = ['Pending', 'Late', 'Missed'];
    if (statusMatch && open.indexOf(statusMatch.status) > -1) {
      var tab = statusMatch.tab;
      setCell_(tab, statusMatch.index, 'Status', onTime === 'No' ? 'Completed late' : 'Complete');
      setCell_(tab, statusMatch.index, 'Submitted at', submittedAt);
      setCell_(tab, statusMatch.index, 'Submitted by', leader.charAt(0) === '=' ? "'" + leader : leader);
    }
    SpreadsheetApp.flush();
  });
}

// The form that fired. e.source is the Form; the trigger lookup is a fallback.
function formIdFromEvent_(e) {
  try {
    if (e.source && e.source.getId) return e.source.getId();
  } catch (err) {}
  var match = ScriptApp.getProjectTriggers().filter(function (t) { return t.getUniqueId() === e.triggerUid; })[0];
  if (match) return match.getTriggerSourceId();
  throw new Error('Could not tell which form this submission came from.');
}
