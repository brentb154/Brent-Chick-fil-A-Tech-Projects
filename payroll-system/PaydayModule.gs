/**
 * Payday Module - Single source of truth for the bi-weekly payday calendar.
 *
 * Every payday/pay-period calculation in the system flows through these helpers so
 * there is exactly ONE reference date and ONE set of date math. The reference date
 * is operator-editable via the `paydayReference` setting (Settings sheet); the
 * hardcoded fallback only applies if the setting is missing/invalid.
 *
 * Conventions (do not deviate):
 * - All Date objects are anchored at LOCAL noon so UTC vs local never shifts the day.
 * - All date strings are produced with formatDateISO() (local components), never toISOString().
 */

// Hard fallback if the setting and DEFAULT_SETTINGS are both unavailable.
// Friday, November 29, 2024 (a known Friday payday).
const PAYDAY_FALLBACK_REFERENCE = '2024-11-29';

/**
 * Returns the configured payday reference as a LOCAL-noon Date (always a Friday in
 * normal operation). Reads the `paydayReference` setting, falls back to DEFAULT_SETTINGS,
 * then to PAYDAY_FALLBACK_REFERENCE. Logs a warning if the resolved date isn't a Friday.
 * @returns {Date} Local-noon reference payday
 */
function getPaydayReferenceDate_() {
  let ref = null;

  try {
    const settings = getSettings();
    if (settings && settings.paydayReference) ref = settings.paydayReference;
  } catch (e) {
    // getSettings unavailable (e.g. very early init) - fall through to defaults
  }

  if (!ref && typeof DEFAULT_SETTINGS !== 'undefined' && DEFAULT_SETTINGS.paydayReference) {
    ref = DEFAULT_SETTINGS.paydayReference;
  }
  if (!ref) ref = PAYDAY_FALLBACK_REFERENCE;

  let parsed = null;
  if (ref instanceof Date) {
    parsed = new Date(ref.getFullYear(), ref.getMonth(), ref.getDate(), 12, 0, 0);
  } else {
    const parts = String(ref).trim().split('-');
    if (parts.length === 3) {
      // YYYY-MM-DD built with explicit local components (no UTC parse)
      parsed = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10), 12, 0, 0);
    }
  }

  if (!parsed || isNaN(parsed.getTime())) {
    parsed = new Date(2024, 10, 29, 12, 0, 0); // Nov 29 2024
  }

  if (parsed.getDay() !== 5) {
    console.warn('paydayReference is not a Friday: "' + ref + '" (day of week ' + parsed.getDay() + ')');
  }

  return parsed;
}

/**
 * Generates the canonical bi-weekly Friday payday series around today.
 * @param {number} historyCount - How many past paydays to include (includes the current/most-recent payday)
 * @param {number} futureCount - How many future paydays to include
 * @returns {Date[]} Local-noon Date objects, sorted chronologically (oldest first)
 */
function generatePaydaySeries_(historyCount, futureCount) {
  const referenceDate = getPaydayReferenceDate_();

  const today = new Date();
  today.setHours(12, 0, 0, 0); // noon to avoid edge cases

  // Most recent payday on or before today
  let mostRecentPayday = new Date(referenceDate.getTime());
  while (mostRecentPayday < today) {
    mostRecentPayday.setDate(mostRecentPayday.getDate() + 14);
  }
  if (mostRecentPayday > today) {
    mostRecentPayday.setDate(mostRecentPayday.getDate() - 14);
  }

  if (mostRecentPayday.getDay() !== 5) {
    console.error('WARNING: mostRecentPayday is not a Friday! Day of week: ' + mostRecentPayday.getDay());
  }

  const paydays = [];

  // History (includes most-recent at i = 0)
  for (let i = historyCount - 1; i >= 0; i--) {
    const d = new Date(mostRecentPayday.getTime());
    d.setDate(mostRecentPayday.getDate() - (i * 14));
    paydays.push(d);
  }

  // Future
  for (let i = 1; i <= futureCount; i++) {
    const d = new Date(mostRecentPayday.getTime());
    d.setDate(mostRecentPayday.getDate() + (i * 14));
    paydays.push(d);
  }

  paydays.sort((a, b) => a - b);
  return paydays;
}

/**
 * Returns the payday (Friday) for the pay period that contains the given date.
 * Pay period: Sunday..Saturday (14 days), paid the following Friday (6 days after period end).
 * @param {Date|string} date - Any date inside a pay period
 * @returns {Date} Local-noon payday Date
 */
function getPaydayForDate_(date) {
  const referencePayday = getPaydayReferenceDate_();

  const received = new Date(date);
  received.setHours(12, 0, 0, 0);

  const DAYS_BEFORE_PAY = 6; // period ends Saturday, 6 days before Friday payday
  const PERIOD_LENGTH = 14;

  let currentPayday = new Date(referencePayday.getTime());
  let periodEnd = new Date(currentPayday.getTime());
  periodEnd.setDate(periodEnd.getDate() - DAYS_BEFORE_PAY);
  let periodStart = new Date(periodEnd.getTime());
  periodStart.setDate(periodStart.getDate() - (PERIOD_LENGTH - 1));

  while (received < periodStart) {
    currentPayday.setDate(currentPayday.getDate() - 14);
    periodEnd.setDate(periodEnd.getDate() - 14);
    periodStart.setDate(periodStart.getDate() - 14);
  }
  while (received > periodEnd) {
    currentPayday.setDate(currentPayday.getDate() + 14);
    periodEnd.setDate(periodEnd.getDate() + 14);
    periodStart.setDate(periodStart.getDate() + 14);
  }

  return currentPayday;
}

/**
 * Returns the pay period boundaries for a given payday.
 * @param {Date|string} payday - The Friday payday
 * @returns {{periodStart: Date, periodEnd: Date}} Local-noon boundary dates
 */
function getPayPeriodForPayday_(payday) {
  let pd;
  if (payday instanceof Date) {
    pd = new Date(payday.getTime());
  } else {
    const s = String(payday);
    pd = new Date(s.indexOf('T') !== -1 ? s : s + 'T12:00:00');
  }
  pd.setHours(12, 0, 0, 0);

  const periodEnd = new Date(pd);
  periodEnd.setDate(periodEnd.getDate() - 6); // Saturday before Friday payday

  const periodStart = new Date(periodEnd);
  periodStart.setDate(periodStart.getDate() - 13); // 14-day period

  return { periodStart: periodStart, periodEnd: periodEnd };
}

/**
 * Returns the payday (Friday) for a given pay-period end (Saturday).
 * The exact inverse of getPayPeriodForPayday_: payday is 6 days after the period end.
 * @param {Date|string} periodEnd - The Saturday period-end (Date or 'YYYY-MM-DD')
 * @returns {Date} Local-noon payday Date
 */
function getPaydayForPeriodEnd_(periodEnd) {
  let pe;
  if (periodEnd instanceof Date) {
    pe = new Date(periodEnd.getFullYear(), periodEnd.getMonth(), periodEnd.getDate(), 12, 0, 0);
  } else {
    const s = String(periodEnd);
    pe = new Date(s.indexOf('T') !== -1 ? s : s + 'T12:00:00');
  }
  pe.setHours(12, 0, 0, 0);
  const payday = new Date(pe);
  payday.setDate(payday.getDate() + 6);
  return payday;
}
