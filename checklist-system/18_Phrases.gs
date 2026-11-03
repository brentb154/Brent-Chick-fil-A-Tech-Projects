/**
 * ============================================================
 * CHECKLIST SYSTEM - Slack Phrases
 * ============================================================
 * A task with [PHRASE] in it, e.g.
 *   Post "[PHRASE]" in the EOD Notes Slack channel.
 * shows a different phrase from the Phrases tab each day. If the
 * phrase isn't in Slack, the checklist wasn't read.
 *
 * Each checklist gets the phrase used longest ago. The Phrases tab
 * keeps the day and checklist it went to, so a rebuild later the
 * same day keeps the same phrase. With 3 checklists, 42 phrases
 * means none comes back within 14 days (Validate warns below that).
 *
 * On a one-submission checklist the phrase task also moves to a
 * different spot each day: never first or last, the same all day.
 * Run setupPhrasesTab() once to make the tab and its starter list.
 */

var PHRASE_TOKEN = /\[phrase\]/i;
var PHRASE_REPEAT_DAYS = 14;

var STARTER_PHRASES = [
  'Ooga booga', 'Hunga busta', 'My pleasure, but louder', 'The cows are watching', 'Eat mor chikin',
  'Nugget nation rises', 'Waffle fries are a personality', 'Sauce is a love language', 'The lemonade is sentient',
  'Polynesian or bust', 'Second mile, first lap', 'The milkshake machine lives', 'Two pickles, no regrets',
  'Spicy deluxe energy', 'Breading station champion', 'Kale crunch is a lifestyle', 'The cow costume needs PTO',
  'Chick-n-Minis before noon', 'The drive-thru never sleeps', 'Cows can\'t spell', 'Frosted lemonade forever',
  'Honey roasted BBQ believer', 'Cool wrap, cooler team', 'Hash brown scramble supremacy', 'Biscuit whisperer',
  'Ice dream team', 'Mac and cheese diplomacy', 'Lettuce be friends', 'Sauce packets are currency',
  'The nuggets have unionized', 'Bread the bird', 'Zesty buffalo vibes only', 'Sweet tea is a food group',
  'Cow appreciation day is every day', 'Just one more nugget', 'Pickles are a fruit now', 'The Chick-fil-A sauce knows',
  'Grilled cool wrap gang', 'Moo means thank you', 'Fry station royalty', 'Diet lemonade, full effort',
  'Napkin tower complete', 'Order of the golden nugget', 'Strips, not trips', 'Chicken biscuit, no notes',
  'The fryer and I are close', 'Sundays are for the cows', 'Mild sauce, wild heart', 'Extra pickles, extra joy',
  'Mooove with purpose'
];

function hasPhrase_(text) {
  return PHRASE_TOKEN.test(String(text || ''));
}

// Today's phrase for one checklist. The first ask each day marks the phrase used.
function todaysPhrase_(checklistId, dateKey) {
  var tab = readTab_(TABS.phrases);
  var c = { phrase: colOrThrow_(tab, 'Phrase'), last: colOrThrow_(tab, 'Last used'), on: colOrThrow_(tab, 'Used on') };
  var rows = tab.display.map(function (r, i) {
    return { index: i, phrase: r[c.phrase].trim(), last: toDateKey_(r[c.last]), on: r[c.on].trim() };
  }).filter(function (r) { return r.phrase; });

  var kept = rows.filter(function (r) { return r.last === dateKey && r.on === checklistId; })[0];
  if (kept) return kept.phrase;

  // Longest ago first (never used = blank, sorts first); skip ones another checklist has today
  var pick = rows.filter(function (r) { return r.last !== dateKey; }).sort(function (a, b) {
    return a.last < b.last ? -1 : a.last > b.last ? 1 : a.index - b.index;
  })[0];
  if (!pick) throw new Error('The Phrases tab has no phrase left for ' + checklistId + ' today. Add more phrases.');
  tab.sheet.getRange(pick.index + 2, c.last + 1).setNumberFormat('@').setValue(dateKey);
  setCell_(tab, pick.index, 'Used on', checklistId);
  SpreadsheetApp.flush();
  return pick.phrase;
}

// Moves the phrase task to a spot picked from the date: never first or last, and not
// yesterday's spot
function movePhrase_(tasks, checklistId, dateKey) {
  var at = -1;
  tasks.forEach(function (s, i) { if (s.phrase) at = i; });
  if (at < 0 || tasks.length < 3) return tasks;
  var others = tasks.slice(0, at).concat(tasks.slice(at + 1));
  var spots = others.length - 1; // spots 1..spots
  // Walk the last week forward, bumping any day that landed on the day before's spot
  var spot = 0;
  for (var back = 7; back >= 0; back--) {
    var prev = spot;
    spot = phraseSpot_(checklistId, addDays_(dateKey, -back), spots);
    if (spots > 1 && spot === prev) spot = spot % spots + 1;
  }
  others.splice(spot, 0, tasks[at]);
  return others;
}

function phraseSpot_(checklistId, dateKey, spots) {
  var d = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, dateKey + '|' + checklistId + '|phrase', Utilities.Charset.UTF_8);
  return 1 + ((d[0] & 255) * 256 + (d[1] & 255)) % spots;
}

// Validate: the Phrases tab is there and has enough phrases for every checklist using one
function checkPhrases_(out, lists) {
  var count = 0;
  var problem = '';
  try {
    var tab = readTab_(TABS.phrases);
    var col = colOrThrow_(tab, 'Phrase');
    colOrThrow_(tab, 'Last used');
    colOrThrow_(tab, 'Used on');
    count = tab.display.filter(function (r) { return r[col].trim(); }).length;
    if (!count) problem = 'the Phrases tab has no phrases.';
  } catch (err) {
    problem = err.message;
  }
  if (problem) {
    lists.forEach(function (c) { skip_(out, c, 'Checklists row ' + c.row + ' (' + c.id + '): has a [PHRASE] task, but ' + problem); });
    return;
  }
  var needed = lists.length * PHRASE_REPEAT_DAYS;
  if (count < needed) {
    out.warnings.push('Phrases: ' + count + ' phrases for ' + lists.length + ' checklists. Add ' + (needed - count) +
      ' more so no phrase comes back within ' + PHRASE_REPEAT_DAYS + ' days.');
  }
}

// Run once from the editor. Makes the Phrases tab if it's missing and adds any starter phrase
// it doesn't have. Safe to re-run: never changes or removes a row.
function setupPhrasesTab() {
  var ss = getSS_();
  var sheet = ss.getSheetByName(TABS.phrases);
  if (!sheet) {
    sheet = ss.insertSheet(TABS.phrases);
    sheet.getRange(1, 1, 1, 3).setValues([['Phrase', 'Last used', 'Used on']]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  var tab = readTab_(TABS.phrases);
  var col = colOrThrow_(tab, 'Phrase');
  var have = {};
  tab.display.forEach(function (r) { have[r[col].trim().toLowerCase()] = true; });
  var add = STARTER_PHRASES.filter(function (p) { return !have[p.toLowerCase()]; });
  appendRows_(tab, add.map(function (p) { return { 'Phrase': p }; }));
  SpreadsheetApp.flush();
}
