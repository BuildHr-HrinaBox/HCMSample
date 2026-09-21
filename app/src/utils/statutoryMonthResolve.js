/**
 * Calendar month resolution for statutory Autofill / Excel export.
 *
 * Do not match with s.startsWith(month.slice(0, 3)):
 * "july".startsWith("jun") is true, so July was fetched and laid out as June.
 * May/July were the reported examples; the same prefix bug (and 31-day grids)
 * applies to every month. Resolve once, for all states and forms.
 */

export const STATUTORY_MONTH_NAMES = [
  'January',
  'February',
  'March',
  'April',
  'May',
  'June',
  'July',
  'August',
  'September',
  'October',
  'November',
  'December',
];

export const STATUTORY_MONTH_ABBR = [
  'Jan',
  'Feb',
  'Mar',
  'Apr',
  'May',
  'Jun',
  'Jul',
  'Aug',
  'Sep',
  'Oct',
  'Nov',
  'Dec',
];

function stripYearSuffix(s) {
  return String(s || '')
    .replace(/[\s,./_-]+(?:19|20)\d{2}\b/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function monthIndexFromIsoToken(s) {
  const iso = String(s || '').trim().match(/^(\d{4})-(\d{2})(?:-\d{2})?(?:[t\s].*)?$/i);
  if (!iso) return -1;
  const mi = parseInt(iso[2], 10);
  return mi >= 1 && mi <= 12 ? mi - 1 : -1;
}

/** Word-boundary hits, left-to-right. Longer names win at the same index (March over Mar). */
function monthHitsInText(s) {
  const text = String(s || '');
  const hits = [];
  for (let i = 0; i < STATUTORY_MONTH_NAMES.length; i += 1) {
    const full = STATUTORY_MONTH_NAMES[i];
    const ab = STATUTORY_MONTH_ABBR[i];
    const re = new RegExp(`\\b(?:${full}|${ab})\\b`, 'gi');
    let m = re.exec(text);
    while (m) {
      hits.push({ index: i, at: m.index, len: m[0].length });
      m = re.exec(text);
    }
  }
  hits.sort((a, b) => a.at - b.at || b.len - a.len);
  return hits;
}

/**
 * Map UI / payroll / wage-period month tokens to a canonical month name.
 * Accepts "July", "Jul", "july 2026", "2026-07", and unique prefixes ("sept").
 * Never maps July → June or May → March.
 */
export function resolveToFullMonthName(raw) {
  if (raw == null) return null;
  const original = String(raw).trim();
  if (!original) return null;
  const s = original.toLowerCase().replace(/\s+/g, ' ');

  const isoIdx = monthIndexFromIsoToken(s);
  if (isoIdx >= 0) return STATUTORY_MONTH_NAMES[isoIdx];

  const stripped = stripYearSuffix(s);
  if (!stripped) return null;

  for (let i = 0; i < STATUTORY_MONTH_NAMES.length; i += 1) {
    const full = STATUTORY_MONTH_NAMES[i].toLowerCase();
    const ab = STATUTORY_MONTH_ABBR[i].toLowerCase();
    if (stripped === full || stripped === ab) return STATUTORY_MONTH_NAMES[i];
  }

  if (stripped.length >= 3) {
    const prefixHits = [];
    for (let i = 0; i < STATUTORY_MONTH_NAMES.length; i += 1) {
      const full = STATUTORY_MONTH_NAMES[i].toLowerCase();
      if (full.startsWith(stripped)) prefixHits.push(i);
    }
    if (prefixHits.length === 1) return STATUTORY_MONTH_NAMES[prefixHits[0]];
  }

  const hits = monthHitsInText(original);
  if (hits.length === 1) return STATUTORY_MONTH_NAMES[hits[0].index];
  if (hits.length > 1) {
    const lastAt = hits[hits.length - 1].at;
    const atLast = hits.filter((h) => h.at === lastAt);
    atLast.sort((a, b) => b.len - a.len);
    return STATUTORY_MONTH_NAMES[atLast[0].index];
  }
  return null;
}

export function statutoryMonthIndex(raw) {
  const name = resolveToFullMonthName(raw);
  if (!name) return -1;
  return STATUTORY_MONTH_NAMES.indexOf(name);
}

export function daysInStatutoryMonth(monthToken, year = new Date().getFullYear()) {
  const idx = statutoryMonthIndex(monthToken);
  if (idx < 0) return 31;
  const y = Number(year);
  const yUse = Number.isFinite(y) && y >= 1900 ? y : new Date().getFullYear();
  return new Date(yUse, idx + 1, 0).getDate();
}

/** Expected calendar lengths for validation (not a hardcoded list of “wrong” months). */
export function statutoryMonthDayCountsForYear(year = new Date().getFullYear()) {
  return STATUTORY_MONTH_NAMES.map((name) => ({
    month: name,
    days: daysInStatutoryMonth(name, year),
  }));
}
