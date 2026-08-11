/** Rajasthan Form 15 — Notice of Weekly Holidays [Rule 22(4)]. */

/** Official Form 15 RJ columns — one each (Excel often merges each title across 2–3 cells). */
export const FORM_15_RJ_TABLE_HEADERS = [
  'Name of persons employed',
  'Designation / Department',
  'Day of the week',
];

const norm = (s) =>
  String(s || '')
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();

const stripDedupeSuffix = (h) =>
  String(h || '')
    .replace(/\s+\(\d+\)$/, '')
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

function buildContextBlob(formHeader, rowItem, fileName, sheetText, tableHeaders) {
  return [
    rowItem?.formName,
    rowItem?.FormName,
    rowItem?.description,
    rowItem?.Description,
    rowItem?.formFileName,
    rowItem?.FormFileName,
    rowItem?.siteState,
    rowItem?.SiteState,
    rowItem?.state,
    rowItem?.State,
    fileName,
    formHeader?.title,
    formHeader?.subtitle,
    formHeader?.reference,
    sheetText,
    Array.isArray(tableHeaders) ? tableHeaders.join(' ') : '',
  ]
    .filter((x) => x != null && String(x).trim() !== '')
    .join(' ')
    .toLowerCase();
}

/** Table looks like Notice of Weekly Holidays (not leave-with-wages Form 15). */
export function headersIndicateForm15RajasthanWeeklyHolidays(tableHeaders) {
  const list = (Array.isArray(tableHeaders) ? tableHeaders : [])
    .map((h) => stripDedupeSuffix(h))
    .filter(Boolean);
  if (list.length < 2) return false;
  const joined = list.map((h) => norm(h)).join(' | ');
  if (/earned\s+leave|leave\s+with\s+wages|medical\s+leave|maternity/.test(joined)) return false;
  const hasName = list.some((h) => /name\s+of\s+persons?\s+employed/.test(norm(h)));
  const hasDesignation = list.some((h) => /designation/.test(norm(h)) && /department/.test(norm(h)));
  const hasDay = list.some((h) => /day\s+of\s+the\s+week/.test(norm(h)));
  return hasName && (hasDesignation || hasDay);
}

/**
 * Rajasthan Shops Form 15 (Notice of Weekly Holidays).
 * Must not match TN Form 15 leave registers / Part 1–2.
 */
export function isForm15RajasthanContext(
  formHeader,
  rowItem,
  fileName,
  sheetText = '',
  tableHeaders = []
) {
  const parts = buildContextBlob(formHeader, rowItem, fileName, sheetText, tableHeaders);

  if (/tamil\s*nadu|\b_tn\b|form[\s._-]*15[\s._-]*tn/.test(parts)) return false;
  if (/leave\s+with\s+wages|earned\s+leave|part[\s_-]*[12]\b|parti\b|partii\b/.test(parts)) {
    if (!/weekly\s+holiday|day\s+of\s+the\s+week/.test(parts)) return false;
  }

  if (/form[\s._-]*15[\s._-]*rj|\bform_15_rj\b|\b15_rj\b/.test(parts)) return true;
  if (/rajasthan/.test(parts) && /notice\s+of\s+weekly\s+holidays/.test(parts)) return true;
  if (
    /rajasthan/.test(parts) &&
    /form[\s._-]*15\b/.test(parts) &&
    (/weekly\s+holiday|rule\s*22\s*\(\s*4\s*\)/.test(parts) ||
      headersIndicateForm15RajasthanWeeklyHolidays(tableHeaders))
  ) {
    return true;
  }
  if (headersIndicateForm15RajasthanWeeklyHolidays(tableHeaders) && /rajasthan|\b_rj\b/.test(parts)) {
    return true;
  }
  return false;
}

/**
 * Collapse Excel merge-expanded duplicates so each logical column appears once.
 * e.g. Name×3 + Designation×3 + Day×3 → three unique headers.
 */
export function collapseForm15RajasthanDuplicateHeaders(tableHeaders) {
  const src = Array.isArray(tableHeaders) ? tableHeaders : [];
  const out = [];
  const seen = new Set();
  for (let i = 0; i < src.length; i += 1) {
    const raw = stripDedupeSuffix(src[i]);
    if (!raw) continue;
    const key = norm(raw);
    if (seen.has(key)) continue;
    seen.add(key);
    out.push(raw);
  }
  return out;
}

/** Prefer template wording when the three Form 15 columns are present; else canonical. */
export function resolveForm15RajasthanTableHeaders(tableHeaders) {
  const collapsed = collapseForm15RajasthanDuplicateHeaders(tableHeaders);
  if (collapsed.length >= 2 && headersIndicateForm15RajasthanWeeklyHolidays(collapsed)) {
    // Keep template labels but ensure Day of the week exists when only name+designation survived.
    const hasDay = collapsed.some((h) => /day\s+of\s+the\s+week/.test(norm(h)));
    if (!hasDay && collapsed.length === 2) {
      return [...collapsed, FORM_15_RJ_TABLE_HEADERS[2]];
    }
    return collapsed.slice(0, 3);
  }
  return [...FORM_15_RJ_TABLE_HEADERS];
}

function classifyForm15RJHeader(header) {
  const s = norm(stripDedupeSuffix(header));
  if (/name\s+of\s+persons?\s+employed/.test(s) || (s.includes('name') && s.includes('employed'))) {
    return 'name';
  }
  if (s.includes('designation') && s.includes('department')) return 'designation';
  if (/day\s+of\s+the\s+week/.test(s) || (s.includes('day') && s.includes('week'))) return 'day';
  return '';
}

function readFirstNonEmpty(row, headers, className) {
  if (!row || typeof row !== 'object') return '';
  for (let i = 0; i < headers.length; i += 1) {
    const h = headers[i];
    if (classifyForm15RJHeader(h) !== className) continue;
    const v = row[h];
    if (v != null && String(v).trim() !== '') return v;
  }
  // Also scan object keys (dedupe suffixes / collapsed duplicates).
  for (const [k, v] of Object.entries(row)) {
    if (String(k).startsWith('__')) continue;
    if (classifyForm15RJHeader(k) !== className) continue;
    if (v != null && String(v).trim() !== '') return v;
  }
  return '';
}

export function remapForm15RajasthanRowsToHeaders(rows, sourceHeaders, targetHeaders) {
  const src = Array.isArray(sourceHeaders) ? sourceHeaders : [];
  const tgt = resolveForm15RajasthanTableHeaders(targetHeaders?.length ? targetHeaders : src);
  if (!Array.isArray(rows)) return [];
  return rows.map((row) => {
    if (!row || typeof row !== 'object') {
      const blank = {};
      tgt.forEach((h) => {
        blank[h] = '';
      });
      return blank;
    }
    const out = {};
    tgt.forEach((targetHeader) => {
      const cls = classifyForm15RJHeader(targetHeader);
      let val = '';
      if (Object.prototype.hasOwnProperty.call(row, targetHeader)) {
        val = row[targetHeader];
      } else if (cls) {
        val = readFirstNonEmpty(row, src, cls);
      }
      out[targetHeader] = val == null ? '' : val;
    });
    Object.keys(row).forEach((k) => {
      if (String(k).startsWith('__')) out[k] = row[k];
    });
    return out;
  });
}
