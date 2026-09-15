/** Gujarat Form P — wage / attendance register (Working hours, days 1–31, wage summary columns). */

import * as XLSX from 'xlsx';
import {
  flattenPayrollEarningColumns,
  readPayrollForm15WageAmounts,
  readPayrollScalar,
  readPayrollTextScalar,
} from '../../utils/payrollEarnings';
import { extractApprovedLeavePositiveDayKeys } from './formFKarnataka';
import {
  approvedLeaveRecordOverlapsFormOGJMonth,
  formOGJRecordMatchesWorker,
  readApprovedLeaveLeaveType,
} from './formOGJGujarat';

export function formPGJGujaratHeaderNorm(txt) {
  return String(txt || '')
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

const FORM_PGJ_MONTH_INDEX = {
  jan: 0,
  january: 0,
  feb: 1,
  february: 1,
  mar: 2,
  march: 2,
  apr: 3,
  april: 3,
  may: 4,
  jun: 5,
  june: 5,
  jul: 6,
  july: 6,
  aug: 7,
  august: 7,
  sep: 8,
  sept: 8,
  september: 8,
  oct: 9,
  october: 9,
  nov: 10,
  november: 10,
  dec: 11,
  december: 11,
};

export function clampFormPGJGujaratDaysInMonth(value, fallback = 31) {
  const n = Number(value);
  if (!Number.isFinite(n)) return fallback;
  return Math.min(Math.max(Math.round(n), 28), 31);
}

/** June → 30, July → 31, February → 28/29. */
export function resolveFormPGJGujaratDaysInMonth(monthText, yearText = '') {
  const blob = `${monthText || ''} ${yearText || ''}`
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
  if (!blob) return 31;
  const monthMatch = blob.match(
    /\b(january|february|march|april|may|june|july|august|september|october|november|december|jan|feb|mar|apr|jun|jul|aug|sept?|oct|nov|dec)\b/i
  );
  if (!monthMatch) return 31;
  const monthIdx = FORM_PGJ_MONTH_INDEX[String(monthMatch[1] || '').toLowerCase()];
  if (!Number.isFinite(monthIdx)) return 31;
  const yearMatch = blob.match(/\b(20\d{2}|19\d{2})\b/);
  const yearFromArg = Number(String(yearText || '').replace(/\D/g, '').slice(0, 4));
  const year = yearMatch
    ? Number(yearMatch[1])
    : Number.isFinite(yearFromArg) && yearFromArg >= 1900
      ? yearFromArg
      : new Date().getFullYear();
  return clampFormPGJGujaratDaysInMonth(new Date(year, monthIdx + 1, 0).getDate(), 31);
}

const stripLeadingNumber = (s) =>
  String(s || '')
    .replace(/^\(?\d+\)?\s*[\.\)]?\s*/i, '')
    .trim();

function normHeaderLabel(h) {
  return formPGJGujaratHeaderNorm(stripLeadingNumber(h))
    .replace(/[^a-z0-9 ]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

export function isFormPGJGujaratContext(
  formHeader,
  rowItem,
  fileName,
  sheetText = '',
  tableHeaders = []
) {
  const parts = [
    rowItem?.formName,
    rowItem?.FormName,
    rowItem?.formFileName,
    rowItem?.FormFileName,
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

  if (/\bform\s*q\b|\bform_q\b|form[\s._-]*q[\s._-]*gj/.test(parts)) return false;
  if (/karnataka/.test(parts) && /accumulated\s+leave/.test(parts)) return false;
  if (/form[\s._-]*p[\s._-]*ka\b/.test(parts) && /accumulated\s+leave/.test(parts)) return false;
  const hasGujarat = /gujarat|\b_gj\b|form[\s._-]*p[\s._-]*gj|form_p_gj/.test(parts);
  const hasFormP = /\bform[\s._-]*p\b/.test(parts);
  if (hasGujarat && hasFormP) return true;
  if (/\bform_p_gj\b/.test(parts)) return true;
  if (hasFormP && /muster[\s-]*roll/.test(parts) && /wage\s+register/.test(parts)) return true;
  // Original Form P template fingerprint (Age / Sex / Interval) — not Form Q MH.
  const joined = parts;
  if (
    /interval\s+for\s+rest/.test(joined) &&
    /\bage\b/.test(joined) &&
    /\bsex\b/.test(joined) &&
    /date\s+of\s+month/.test(joined) &&
    (/full\s+name\s+of\s+the\s+worker/.test(joined) || /name\s+of\s+the\s+worker/.test(joined))
  ) {
    return true;
  }
  // App UI headers for Gujarat Form P (Father's name + Wage Rate + Date of Month_n).
  return (
    hasGujarat &&
    /sr\.?\s*no/.test(joined) &&
    /date\s+of\s+month/.test(joined) &&
    (/father/.test(joined) || /wage\s+rate/.test(joined) || /name\s+of\s+the\s+worker/.test(joined))
  );
}

export function isFormPGJGujaratTableLayoutFormHeader(formHeader) {
  return !!formHeader?.formPGJGujaratTableLayout;
}

const FORM_PGJ_WAGE_TAIL_HEADERS = [
  'Total Days Worked',
  'Minimum Rate of Wages Rs.',
  'Total Production in case of piece rate Rs.',
  'Actual Wages Paid Rs.',
  'House Rent Allowance Rs.',
  'Dearness Allowance Rs.',
  'Gross Amount Payable Rs.',
  'Total hours of Overtime during the month',
  'Overtime Earnings Rs.',
  'Provident Fund Rs.',
  'Family Pension Rs.',
  'ESI Contribution Rs.',
  'Professional Tax Rs.',
  'Income Tax Rs.',
  'Loan & Interest Rs.',
  'Advances Rs.',
  'Other Deductions Rs.',
  'Total Deduction Rs.',
  'Net Payable Rs.',
  'Date of Payment',
  'Signature/Thumb Impression',
];

/** Form P — overtime columns always default to NIL (not fetched). */
export const FORM_PGJ_OVERTIME_DEFAULT = 'NIL';

/** Form P — Working hours From / To (fixed defaults). */
export const FORM_PGJ_WORKING_HOURS_FROM = '9 AM';
export const FORM_PGJ_WORKING_HOURS_TO = '5 PM';

/** Format Sample Payroll Pay date as DD-MM-YYYY when possible. */
export function formatFormPGJPayrollPayDate(payDateRaw) {
  const raw = String(payDateRaw || '').trim();
  if (!raw) return '';
  const iso = raw.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) return `${iso[3]}-${iso[2]}-${iso[1]}`;
  const dmy = raw.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})$/);
  if (dmy) {
    const dd = String(dmy[1]).padStart(2, '0');
    const mm = String(dmy[2]).padStart(2, '0');
    return `${dd}-${mm}-${dmy[3]}`;
  }
  return raw;
}

/** Full modal / export column list — worker cols, hours, days 1–31, wage summary. */
export function buildFormPGJGujaratCanonicalHeaders(daysInMonth = 31) {
  const headers = [
    'Sr. No.',
    'Name of the Worker',
    "Father's / Husband's Name",
    'Designation',
    'Date of entry into service',
    'Wage Rate',
    'Working hours_From',
    'Working hours_To',
  ];
  const days = Math.min(Math.max(Number(daysInMonth) || 31, 28), 31);
  for (let d = 1; d <= days; d += 1) {
    headers.push(`Date of Month_${d}`);
  }
  FORM_PGJ_WAGE_TAIL_HEADERS.forEach((h) => headers.push(h));
  return headers;
}

export function buildFormPGJSubColumnsFromHeaders(headers) {
  const subs = {};
  (headers || []).forEach((h) => {
    const m = String(h || '').match(/^(.+)_([\s\S]+)$/);
    if (!m) return;
    const parent = m[1].trim();
    const sub = m[2].trim();
    if (!subs[parent]) subs[parent] = [];
    if (!subs[parent].includes(sub)) subs[parent].push(sub);
  });
  return subs;
}

export function resolveFormPGJGujaratTableHeaders(_tableHeaders, daysInMonth = 31) {
  return buildFormPGJGujaratCanonicalHeaders(daysInMonth);
}

const isEmployerMetaLabel = (txt) => {
  const t = formPGJGujaratHeaderNorm(txt);
  return (
    /name\s+of\s+the\s+employer/.test(t) ||
    /^month\s*:?$/.test(t) ||
    /name\s+of\s+the\s+establishment/.test(t)
  );
};

const FORM_PGJ_DEFAULT_HEADER_FIELDS = [
  { label: 'Name of the employer', value: '', key: 'form_p_gj_employer' },
  { label: 'Month', value: '', key: 'form_p_gj_month' },
];

function ensureFormPGJGujaratHeaderFields(fields) {
  const list = Array.isArray(fields) ? [...fields] : [];
  const has = (re) =>
    list.some((f) => formPGJGujaratHeaderNorm(f?.label).replace(/:+$/, '').trim().match(re));
  if (!has(/name\s+of\s+(?:the\s+)?employer/)) {
    list.unshift({ label: 'Name of the employer', value: '', key: 'form_p_gj_employer' });
  }
  if (!has(/^month$/)) {
    list.push({ label: 'Month', value: '', key: 'form_p_gj_month' });
  }
  return list;
}

/** "VAYONA ENERGY PRIVATE LIMITED, address…" → company name for Name of the employer. */
export function formPGJEmployerDisplayName(text) {
  const raw = String(text || '').trim();
  if (!raw) return '';
  const comma = raw.indexOf(',');
  if (comma >= 8) return raw.slice(0, comma).trim();
  return raw;
}

export function applyFormPGJGujaratHeaderAutofill(headerData, siteContext = {}, options = {}) {
  const { onlyEmpty = false } = options;
  const out = headerData && typeof headerData === 'object' ? { ...headerData } : {};
  const monthText = String(siteContext.monthText || '').trim();
  const employerText = formPGJEmployerDisplayName(siteContext.employerText);
  const setIf = (key, value) => {
    if (!key || !value) return;
    const cur = String(out[key] ?? '').trim();
    if (!onlyEmpty || !cur || /^enter\b/i.test(cur)) out[key] = value;
  };
  if (monthText) {
    setIf('form_p_gj_month', monthText);
    setIf('form_x_month', monthText);
  }
  if (employerText) setIf('form_p_gj_employer', employerText);
  return out;
}

/** Excel section index (1–9) under worker parents — not a data sub-column. */
const isFormPGJSectionMarkerSub = (sub, main) => {
  const subText = String(sub || '').trim();
  if (!/^\d{1,2}$/.test(subText)) return false;
  const n = parseInt(subText, 10);
  if (n < 1 || n > 9) return false;
  const ml = formPGJGujaratHeaderNorm(main);
  if (ml.includes('date') && ml.includes('month')) return false;
  return (
    /sr\.?\s*no/.test(ml) ||
    /name\s+of\s+the\s+worker/.test(ml) ||
    /father|husband/.test(ml) ||
    /designation/.test(ml) ||
    /entry\s+into\s+service/.test(ml) ||
    /wage\s+rate/.test(ml) ||
    /working\s+hours/.test(ml)
  );
};

const isFormPGJBlankOrMetaTableHeader = (headerKey) => {
  const t = String(headerKey || '').trim();
  if (!t || t === ':') return true;
  if (/^column\s*\d+$/i.test(t)) return true;
  const m = t.match(/^(\d{1,2})$/);
  if (m) {
    const n = parseInt(m[1], 10);
    return n >= 1 && n <= 9;
  }
  const tl = t.toLowerCase().replace(/\s+/g, ' ');
  return (
    /name\s+of\s+the\s+employer/.test(tl) ||
    /name\s+of\s+the\s+establishment/.test(tl) ||
    /^month\s*:?$/.test(tl)
  );
};

const normalizeFormPGJHeaderKey = (headerKey) => {
  const h = String(headerKey || '').trim();
  if (!h) return h;
  const m = h.match(/^(.+)_([\s\S]+)$/);
  if (!m) return h;
  const parent = m[1].trim();
  const sub = String(m[2]).trim();
  if (/^(from|to)$/i.test(sub)) return h;
  if (isFormPGJSectionMarkerSub(sub, parent)) return parent;
  if (/^date\s+of\s+month$/i.test(parent) && /^\d{1,2}$/.test(sub)) {
    return `Date of Month_${sub}`;
  }
  return h;
};

export const stripAndNormalizeFormPGJTableHeaders = (headers) => {
  if (!Array.isArray(headers) || headers.length === 0) return [];
  const normalized = headers.map(normalizeFormPGJHeaderKey);
  let start = 0;
  while (start < normalized.length && isFormPGJBlankOrMetaTableHeader(normalized[start])) {
    start += 1;
  }
  const seen = new Set();
  const out = [];
  for (let i = start; i < normalized.length; i += 1) {
    const h = normalized[i];
    if (isFormPGJBlankOrMetaTableHeader(h)) continue;
    const key = h.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    out.push(h);
  }
  return out;
};

export function remapFormPGJRowsToHeaders(rows, oldHeaders, newHeaders) {
  if (!Array.isArray(rows) || !Array.isArray(newHeaders) || newHeaders.length === 0) {
    return rows || [];
  }
  const oldList = Array.isArray(oldHeaders) ? oldHeaders : [];
  // Punctuation-insensitive (Rs. vs Rs) so Actual Wages / OT values survive remap.
  const normOld = (h) => normHeaderLabel(h);
  return rows.map((row, rowIndex) => {
    const next = {};
    newHeaders.forEach((newH) => {
      if (row && row[newH] != null && String(row[newH]).trim() !== '') {
        next[newH] = row[newH];
        return;
      }
      const newNorm = normOld(newH);
      for (let i = 0; i < oldList.length; i += 1) {
        const oldH = oldList[i];
        if (normOld(oldH) !== newNorm) continue;
        const v = row?.[oldH];
        if (v != null && String(v).trim() !== '') {
          next[newH] = v;
          return;
        }
      }
      // Bucket fallback — e.g. Actual Wages Paid under a slightly different label.
      if (row && typeof row === 'object') {
        const newBucket = matchFormPGJWageTailBucket(newH) || (
          isFormPGJWageRateHeader(newH) ? 'wageRate' : ''
        );
        if (newBucket && newBucket !== 'skip') {
          for (const [k, v] of Object.entries(row)) {
            if (String(k).startsWith('__')) continue;
            if (matchFormPGJWageTailBucket(k) === newBucket && v != null && String(v).trim() !== '') {
              next[newH] = v;
              return;
            }
          }
        }
      }
      if (/^sr\.?\s*no/i.test(newH)) {
        next[newH] = String(rowIndex + 1);
      } else {
        next[newH] = '';
      }
    });
    if (row && typeof row === 'object') {
      Object.entries(row).forEach(([k, v]) => {
        if (String(k).startsWith('__') && next[k] == null) next[k] = v;
      });
    }
    return next;
  });
}

/**
 * Parse Form P GJ Excel header rows (parent + From/To/day subs + wage tail columns).
 */
export function rebuildFormPGJGujaratTableHeadersFromSheet({
  worksheet,
  jsonData,
  effectiveSheetCols,
  getMergedAwareCellText,
  merges,
  daysInMonth = 31,
}) {
  const norm = (txt) =>
    String(txt || '')
      .replace(/\r?\n/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  const normLower = (txt) => norm(txt).toLowerCase();

  const getMergeSpanEndCol = (r, c) => {
    for (let i = 0; i < merges.length; i += 1) {
      const m = merges[i];
      if (!m?.s || !m?.e) continue;
      if (r >= m.s.r && r <= m.e.r && c >= m.s.c && c <= m.e.c) {
        return m.e.c + 1;
      }
    }
    return c + 1;
  };

  let headerRowIndex = -1;
  for (let r = 0; r < Math.min(30, jsonData.length); r += 1) {
    const parts = [];
    for (let c = 0; c < Math.max(effectiveSheetCols, 20); c += 1) {
      const t = normLower(getMergedAwareCellText(r, c));
      if (t) parts.push(t);
    }
    const joined = parts.join(' ');
    if (/sr\.?\s*no/.test(joined) && /name\s+of\s+the\s+worker/.test(joined)) {
      headerRowIndex = r;
      break;
    }
    if (/working\s+hours/.test(joined) && /date\s+of\s+month/.test(joined)) {
      headerRowIndex = r;
      break;
    }
  }
  if (headerRowIndex < 0) return null;

  let startCol = -1;
  for (let c = 0; c < Math.max(effectiveSheetCols, 12); c += 1) {
    const t = normLower(getMergedAwareCellText(headerRowIndex, c));
    if (/^sr\.?\s*no\.?$/.test(t) || t === 'sr no' || t.startsWith('sr. no')) {
      startCol = c;
      break;
    }
  }
  if (startCol < 0) startCol = 0;

  const subRowIndex = headerRowIndex + 1;
  const flatHeaders = [];
  const subColumnsData = {};
  const mainHeaders = [];

  let c = startCol;
  const maxCols = Math.max(effectiveSheetCols, (jsonData[headerRowIndex] || []).length, 70);
  while (c < maxCols) {
    const mergeEnd = getMergeSpanEndCol(headerRowIndex, c);
    let endC = mergeEnd;
    const main = norm(getMergedAwareCellText(headerRowIndex, c));
    if (!main || isEmployerMetaLabel(main) || isFormPGJBlankOrMetaTableHeader(main)) {
      c = Math.max(c + 1, endC);
      continue;
    }
    while (endC < maxCols) {
      const nextMain = norm(getMergedAwareCellText(headerRowIndex, endC));
      if (!nextMain) break;
      if (normLower(nextMain) !== normLower(main)) break;
      endC = Math.max(endC, getMergeSpanEndCol(headerRowIndex, endC));
    }

    const mainLower = normLower(main);
    const subs = [];
    for (let sc = c; sc < endC; sc += 1) {
      const sub = norm(getMergedAwareCellText(subRowIndex, sc));
      if (!sub || sub.toLowerCase() === main.toLowerCase()) continue;
      if (sub.endsWith(':') && sub.length < 40) continue;
      if (isFormPGJSectionMarkerSub(sub, main)) continue;
      subs.push(sub);
    }

    mainHeaders.push(main);
    if (!subColumnsData[main]) subColumnsData[main] = [];

    if (/date\s+of\s+month/i.test(mainLower)) {
      const dayCount = Math.min(Math.max(Number(daysInMonth) || 31, 28), 31);
      for (let d = 1; d <= dayCount; d += 1) {
        const key = `Date of Month_${d}`;
        flatHeaders.push(key);
        subColumnsData[main].push(String(d));
      }
    } else if (/working\s+hours/i.test(mainLower)) {
      const fromTo = subs.filter((s) => /^(from|to)$/i.test(s));
      if (fromTo.length >= 2) {
        fromTo.forEach((sub) => {
          const key = `Working hours_${sub}`;
          flatHeaders.push(key);
          subColumnsData[main].push(sub);
        });
      } else {
        flatHeaders.push('Working hours_From', 'Working hours_To');
        subColumnsData[main].push('From', 'To');
      }
    } else if (subs.length === 0) {
      flatHeaders.push(main);
    } else {
      const allNumericDays = subs.every((s) => /^\d{1,2}$/.test(s));
      if (allNumericDays && subs.length >= 20) {
        subs.forEach((sub) => {
          const key = `Date of Month_${sub}`;
          flatHeaders.push(key);
          subColumnsData[main].push(sub);
        });
      } else {
        subs.forEach((sub) => {
          const key = `${main}_${sub}`;
          flatHeaders.push(key);
          subColumnsData[main].push(sub);
        });
      }
    }
    c = endC;
  }

  const canonical = buildFormPGJGujaratCanonicalHeaders(daysInMonth);
  const expandedHeaders = stripAndNormalizeFormPGJTableHeaders(
    flatHeaders.length >= 10 ? flatHeaders : canonical
  );
  const finalHeaders =
    expandedHeaders.length >= canonical.length - 4
      ? expandedHeaders
      : stripAndNormalizeFormPGJTableHeaders(canonical);

  const headerFields = [];
  for (let r = 0; r < headerRowIndex; r += 1) {
    for (let col = 0; col < Math.max(20, effectiveSheetCols); col += 1) {
      const raw = String(getMergedAwareCellText(r, col) || '').replace(/\s+/g, ' ').trim();
      if (!raw) continue;
      const labelOnly = raw.split(':')[0].trim();
      if (/^month\s*:?$/i.test(raw) || /^month$/i.test(labelOnly)) {
        let val = String(raw.split(':').slice(1).join(':') || '').trim();
        if (!val) {
          for (let nc = col + 1; nc < Math.min(col + 8, maxCols); nc += 1) {
            const v = String(getMergedAwareCellText(r, nc) || '').trim();
            if (v && !/^month\s*:?$/i.test(v)) {
              val = v;
              break;
            }
          }
        }
        if (!headerFields.some((f) => f.key === 'form_p_gj_month')) {
          headerFields.push({ label: 'Month', value: val, key: 'form_p_gj_month' });
        }
      }
      if (/name\s+of\s+(?:the\s+)?employer/i.test(labelOnly) && !/address/i.test(labelOnly)) {
        let val = String(raw.split(':').slice(1).join(':') || '').trim();
        if (!val) {
          for (let nc = col + 1; nc < Math.min(col + 8, maxCols); nc += 1) {
            const v = String(getMergedAwareCellText(r, nc) || '').trim();
            if (v && !/name\s+of\s+(?:the\s+)?employer/i.test(v)) {
              val = v;
              break;
            }
          }
        }
        if (!headerFields.some((f) => f.key === 'form_p_gj_employer')) {
          headerFields.push({
            label: 'Name of the employer',
            value: val,
            key: 'form_p_gj_employer',
          });
        }
      }
    }
  }

  return {
    headerRowIndex,
    startCol,
    headers: mainHeaders,
    expandedHeaders: finalHeaders,
    subColumnsData,
    headerFields: ensureFormPGJGujaratHeaderFields(headerFields),
  };
}

export function resolveFormPGJTableHeadersForWorkbook(workbook, daysInMonth = 31) {
  const canonical = stripAndNormalizeFormPGJTableHeaders(
    buildFormPGJGujaratCanonicalHeaders(daysInMonth)
  );
  if (!workbook?.SheetNames?.[0]) {
    return {
      headers: canonical,
      subColumns: buildFormPGJSubColumnsFromHeaders(canonical),
      headerRowIndex: 11,
      dataStartIndex: 14,
      headerFields: [...FORM_PGJ_DEFAULT_HEADER_FIELDS],
    };
  }
  const worksheet = workbook.Sheets[workbook.SheetNames[0]];
  const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
  const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1:A1');
  const maxSheetCols = range.e && typeof range.e.c === 'number' ? range.e.c + 1 : 0;
  const merges = worksheet['!merges'] || [];
  let mergeMaxCol = 0;
  merges.forEach((m) => {
    if (m?.e && typeof m.e.c === 'number') mergeMaxCol = Math.max(mergeMaxCol, m.e.c + 1);
  });
  const effectiveSheetCols = Math.max(maxSheetCols, mergeMaxCol);
  const getRawCellText = (r, c) => {
    if (r < 0 || c < 0) return '';
    const ref = XLSX.utils.encode_cell({ r, c });
    const cell = worksheet[ref];
    if (!cell || cell.v == null) return '';
    return String(cell.v).trim();
  };
  const getMergedAwareCellText = (r, c) => {
    const direct = getRawCellText(r, c);
    if (direct) return direct;
    for (let i = 0; i < merges.length; i += 1) {
      const m = merges[i];
      if (!m?.s || !m?.e) continue;
      if (r >= m.s.r && r <= m.e.r && c >= m.s.c && c <= m.e.c) {
        return getRawCellText(m.s.r, m.s.c);
      }
    }
    return '';
  };
  const rebuild = rebuildFormPGJGujaratTableHeadersFromSheet({
    worksheet,
    jsonData,
    effectiveSheetCols,
    getMergedAwareCellText,
    merges,
    daysInMonth,
  });
  const headers = rebuild?.expandedHeaders?.length >= 10 ? rebuild.expandedHeaders : canonical;
  let dataStartIndex = (rebuild?.headerRowIndex ?? 11) + 2;
  if (rebuild?.headerRowIndex != null && rebuild.headerRowIndex + 2 < jsonData.length) {
    let seqHits = 0;
    const sectionRow = rebuild.headerRowIndex + 2;
    const startCol = rebuild.startCol ?? 0;
    for (let hi = 0; hi < Math.min(9, headers.length); hi += 1) {
      const t = String(getMergedAwareCellText(sectionRow, startCol + hi) || '')
        .replace(/\s+/g, '')
        .trim();
      if (t === String(hi + 1)) seqHits += 1;
    }
    if (seqHits >= 4) dataStartIndex = sectionRow + 1;
  }
  return {
    headers: stripAndNormalizeFormPGJTableHeaders(headers),
    subColumns:
      rebuild?.subColumnsData &&
      Object.keys(rebuild.subColumnsData).some((k) => (rebuild.subColumnsData[k] || []).length > 0)
        ? rebuild.subColumnsData
        : buildFormPGJSubColumnsFromHeaders(headers),
    headerRowIndex: rebuild?.headerRowIndex ?? 11,
    dataStartIndex,
    headerFields: ensureFormPGJGujaratHeaderFields(rebuild?.headerFields),
  };
}

function unwrapEmployeeRecord(emp) {
  return emp?.Employee || emp?.employee || emp;
}

function pickScalarEmployeeValue(raw) {
  if (raw == null || raw === '') return '';
  if (typeof raw === 'object') {
    return String(
      raw.display_value ??
        raw.displayValue ??
        raw.name ??
        raw.Name ??
        raw.value ??
        raw.text ??
        ''
    ).trim();
  }
  return String(raw).trim();
}

function collectEmployeeSources(emp) {
  const sources = [];
  const seen = new Set();
  const push = (value) => {
    if (!value || typeof value !== 'object' || seen.has(value)) return;
    seen.add(value);
    sources.push(value);
  };
  push(emp);
  if (emp && typeof emp === 'object') {
    push(emp.Employee);
    push(emp.employee);
    push(unwrapEmployeeRecord(emp));
  }
  return sources;
}

function readScalarFromEmployee(emp, keys) {
  const sources = collectEmployeeSources(emp);
  for (let i = 0; i < sources.length; i += 1) {
    const src = sources[i];
    for (let k = 0; k < keys.length; k += 1) {
      const val = pickScalarEmployeeValue(src[keys[k]]);
      if (val && val !== '-') return val;
    }
    for (const [k, v] of Object.entries(src)) {
      const kn = String(k || '')
        .toLowerCase()
        .replace(/[^a-z0-9]/g, '');
      if (keys.some((key) => kn === String(key).toLowerCase().replace(/[^a-z0-9]/g, ''))) {
        const val = pickScalarEmployeeValue(v);
        if (val && val !== '-') return val;
      }
    }
  }
  return '';
}

export function readFormPGJShiftStartFromEmployee(emp) {
  return readScalarFromEmployee(emp, [
    'ShiftStartTime',
    'Shift_Start_Time',
    'shiftStartTime',
    'Shift Start Time',
  ]);
}

export function readFormPGJShiftEndFromEmployee(emp) {
  return readScalarFromEmployee(emp, [
    'ShiftEndTime',
    'Shift_End_Time',
    'shiftEndTime',
    'Shift End Time',
  ]);
}

export function readFormPGJShiftNameFromEmployee(emp) {
  return readScalarFromEmployee(emp, [
    'ShiftName',
    'Shift_Name',
    'shiftName',
    'Shift Name',
    'Shift',
  ]);
}

export function isFormPGJWorkingHoursFromHeader(header) {
  const n = normHeaderLabel(header);
  return n === 'working hours from' || (n.includes('working hours') && n.includes('from'));
}

export function isFormPGJWorkingHoursToHeader(header) {
  const n = normHeaderLabel(header);
  return n === 'working hours to' || (n.includes('working hours') && n.includes('to'));
}

export function isFormPGJDateOfMonthDayHeader(header) {
  return /^date\s+of\s+month_\d{1,2}$/i.test(String(header || '').trim());
}

export function isFormPGJWageTailHeader(header) {
  const n = normHeaderLabel(header);
  return FORM_PGJ_WAGE_TAIL_HEADERS.some((h) => normHeaderLabel(h) === n);
}

export function isFormPGJSignatureHeader(header) {
  const n = normHeaderLabel(header);
  return n.includes('signature') || n.includes('thumb impression');
}

function parsePayrollNumber(value) {
  const n = Number(String(value ?? '').replace(/[,₹]/g, '').trim());
  return Number.isFinite(n) ? n : NaN;
}

/** Prefer flatten-computed basic; also accept Basic Salary / basic_* component keys. */
function readFormPGJBasicAmount(payrollRow, flat) {
  const wages = readPayrollForm15WageAmounts(payrollRow);
  const candidates = [
    wages.basic,
    flat?.basic,
    flat?.earned_basic,
    flat?.basic_pay,
    flat?.basic_earnings,
    flat?.basic_salary,
    flat?.Basic,
    flat?.['Basic Salary'],
    flat?.['basic salary'],
  ];
  if (flat && typeof flat === 'object') {
    Object.entries(flat).forEach(([key, value]) => {
      const k = String(key || '')
        .toLowerCase()
        .replace(/[^a-z0-9]+/g, '_');
      if (!k) return;
      if (k.includes('arrear') || k.includes('overtime') || k.includes('_ot_') || k.endsWith('_ot')) return;
      if (
        k === 'basic' ||
        k === 'earned_basic' ||
        k === 'basic_pay' ||
        k === 'basic_earnings' ||
        k === 'basic_salary' ||
        /(^|_)basic($|_)/.test(k)
      ) {
        candidates.push(value);
      }
    });
  }
  for (let i = 0; i < candidates.length; i += 1) {
    const n = parsePayrollNumber(candidates[i]);
    if (Number.isFinite(n)) return n;
  }
  return '';
}

function readPayrollDeductionScalar(flat, payrollRow, keys, patterns = []) {
  const val = readPayrollScalar({ ...flat, ...payrollRow }, keys, patterns);
  return val !== '' && val != null ? val : '';
}

/** Prefer flatten() computed Sample Payroll / benefit columns first. */
function pickFormPGJAmount(...values) {
  for (let i = 0; i < values.length; i += 1) {
    const value = values[i];
    if (value === '' || value == null) continue;
    const n = parsePayrollNumber(value);
    if (Number.isFinite(n)) return n;
  }
  return '';
}

/** Form P Total Deduction Rs. ← gross_pay − net_pay */
export function computeFormPGJGujaratTotalDeduction(grossPay, netPay) {
  const g = parsePayrollNumber(grossPay);
  const n = parsePayrollNumber(netPay);
  if (!Number.isFinite(g) || !Number.isFinite(n)) return '';
  const diff = Math.round((g - n) * 100) / 100;
  return diff < 0 ? '' : diff;
}

/** Form P Other Deductions Rs. ← Sample Payroll Total Deduction − Professional Tax */
export function computeFormPGJGujaratOtherDeductions(payrollTotalDeduction, professionalTax) {
  if (payrollTotalDeduction === '' || payrollTotalDeduction == null) return '';
  const total = parsePayrollNumber(payrollTotalDeduction);
  if (!Number.isFinite(total)) return '';
  const ptRaw = professionalTax === '' || professionalTax == null ? 0 : parsePayrollNumber(professionalTax);
  const ptVal = Number.isFinite(ptRaw) ? ptRaw : 0;
  const diff = Math.round((total - ptVal) * 100) / 100;
  return diff < 0 ? 0 : diff;
}

export function resolveFormPGJGujaratPayrollFields(payrollRow, helpers = {}) {
  const monthEndDate = String(helpers.monthEndDate || '').trim();
  const payDateFromHelpers = () => {
    // Date of Payment ← month end (preferred); Pay date only as legacy fallback.
    if (monthEndDate) return monthEndDate;
    const raw = String(helpers.payDate || '').trim();
    if (!raw) return '';
    return (
      formatFormPGJPayrollPayDate(raw) ||
      (typeof helpers.formatStatutoryDateDisplay === 'function'
        ? helpers.formatStatutoryDateDisplay(raw) || raw
        : raw)
    );
  };
  const empty = {
    paidDays: '',
    basic: '',
    hra: '',
    grossPay: '',
    overtimeHours: FORM_PGJ_OVERTIME_DEFAULT,
    overtimeEarnings: FORM_PGJ_OVERTIME_DEFAULT,
    esi: '',
    pf: '',
    professionalTax: '',
    incomeTax: '',
    otherDeductions: '',
    totalDeduction: '',
    netPay: '',
    paymentDate: payDateFromHelpers(),
  };
  if (!payrollRow || payrollRow.fetch_error) return empty;

  // Flatten first — raw payrollRow.basic often overwrites computed basic when merged last.
  const flat = flattenPayrollEarningColumns(payrollRow);
  const wages = readPayrollForm15WageAmounts(payrollRow);
  const source = { ...flat, ...payrollRow };

  const paidDays = readPayrollScalar(
    flat,
    ['paid_days', 'Paid Days', 'days_worked', 'Days Worked', 'paidDays', 'no_of_days_worked'],
    [/^paid_days$/, /paiddays/, /daysworked/]
  );

  const basic = readFormPGJBasicAmount(payrollRow, flat);

  const hra =
    wages.hra !== '' && wages.hra != null
      ? wages.hra
      : readPayrollScalar(
          flat,
          ['hra_fbp', 'HRA FBP', 'hra (fbp)', 'hra_fp', 'HRA (FBP)', 'hra', 'HRA', 'house_rent_allowance'],
          [/^hra_fbp$/, /^hra_fp$/, /^hra$/, /house.*rent/]
        );

  const grossPay = readPayrollScalar(
    source,
    ['gross_pay', 'Gross Pay', 'grossPay', 'gross', 'Gross', 'total_earnings', 'Gross Amount Payable'],
    [/^gross_pay$/, /^gross$/, /^total_earnings$/]
  );

  // Form P: overtime columns are always NIL (not fetched from payroll).
  const overtimeHours = FORM_PGJ_OVERTIME_DEFAULT;
  const overtimeEarnings = FORM_PGJ_OVERTIME_DEFAULT;

  const esi = readPayrollDeductionScalar(
    flat,
    payrollRow,
    ['esi', 'ESI', 'esic', 'ESIC', 'esi_contribution', 'ESI Contribution'],
    [/^esi/, /employee state insurance/]
  );

  // Provident Fund Rs. ← Sample Payroll PF
  const pf = pickFormPGJAmount(
    flat.epf_contribution,
    flat.pf,
    flat.PF,
    flat.provident_fund,
    readPayrollScalar(
      source,
      [
        'pf',
        'PF',
        'epf_contribution',
        'EPF Contribution',
        'epf',
        'EPF',
        'employee_pf',
        'Employee PF',
        'provident_fund',
        'Provident Fund',
      ],
      [/^pf$/, /^epf(_contribution)?$/, /^provident_fund$/]
    )
  );

  const professionalTax = readPayrollDeductionScalar(
    flat,
    payrollRow,
    ['professional_tax', 'Professional Tax', 'ProfessionalTax', 'professionalTax', 'pt', 'PT'],
    [/^professional_tax$/, /^pt$/]
  );

  // Income Tax Rs. ← Sample Payroll Income Tax
  const incomeTax = pickFormPGJAmount(
    flat.income_tax,
    flat.IncomeTax,
    flat.incomeTax,
    flat.tds,
    readPayrollScalar(
      source,
      [
        'income_tax',
        'Income Tax',
        'IncomeTax',
        'incomeTax',
        'tds',
        'TDS',
        'tax_deducted_at_source',
      ],
      [/income.*tax/i, /^tds$/, /^incometax$/]
    )
  );

  // Sample Payroll "Total Deduction" column (not Form Total Deduction Rs.).
  const payrollTotalDeduction = readPayrollScalar(
    source,
    [
      'total_deductions',
      'Total Deductions',
      'Total Deduction',
      'totalDeductions',
      'totalDeduction',
      'TotalDeduction',
      'total_employee_deductions',
      'total_deduction',
    ],
    [/^total_deductions?$/, /^totaldeduction$/]
  );

  const netPay = readPayrollScalar(
    source,
    ['net_pay', 'Net Pay', 'netPay', 'netpay', 'Netpay', 'monthly_salary', 'Net Payable'],
    [/^net_pay$/, /^netpay$/]
  );

  // Total Deduction Rs. ← gross_pay − net_pay
  const totalDeduction = computeFormPGJGujaratTotalDeduction(grossPay, netPay);

  // Other Deductions Rs. ← Sample Payroll Total Deduction − Professional Tax
  const otherDeductions = computeFormPGJGujaratOtherDeductions(
    payrollTotalDeduction,
    professionalTax
  );

  // Date of Payment ← month end (Form P requirement).
  const paymentDate =
    monthEndDate ||
    (() => {
      const payDateRaw =
        readPayrollTextScalar(
          source,
          [
            'pay_date',
            'Pay Date',
            'payDate',
            'PayDate',
            'payment_date',
            'Payment Date',
            'paid_date',
            'date_of_payment',
            'Date of Payment',
          ],
          [/^pay_date$/, /^paydate$/, /payment.*date/i, /^paid_date$/, /date.*payment/i]
        ) ||
        String(payrollRow?.pay_date || '').trim() ||
        String(payrollRow?.payDate || '').trim() ||
        String(flat?.pay_date || '').trim() ||
        String(flat?.payDate || '').trim() ||
        String(helpers.payDate || '').trim();
      return (
        formatFormPGJPayrollPayDate(payDateRaw) ||
        (typeof helpers.formatStatutoryDateDisplay === 'function'
          ? helpers.formatStatutoryDateDisplay(payDateRaw) || payDateRaw
          : payDateRaw)
      );
    })();

  return {
    paidDays,
    basic,
    hra,
    grossPay,
    overtimeHours,
    overtimeEarnings,
    esi,
    pf,
    professionalTax,
    incomeTax,
    otherDeductions,
    totalDeduction,
    netPay,
    paymentDate,
  };
}

function formatCellValue(value, { allowZero = false } = {}) {
  if (value == null || value === '') return '';
  if (typeof value === 'string' && /^nil$/i.test(value.trim())) return 'NIL';
  const num = parsePayrollNumber(value);
  if (Number.isFinite(num) && (allowZero || num !== 0)) return num;
  return String(value).trim();
}

export function isFormPGJWageRateHeader(header) {
  const n = normHeaderLabel(header);
  return (
    n === 'wage rate' ||
    /^wage\s*rate$/.test(n) ||
    (n.includes('wage') && n.includes('rate') && !n.includes('minimum') && !n.includes('overtime'))
  );
}

export function isFormPGJActualWagesPaidHeader(header) {
  return /actual\s+wages\s+paid/.test(normHeaderLabel(header));
}

export function isFormPGJOvertimeHoursHeader(header) {
  return /total\s+hours?\s+of\s+overtime/.test(normHeaderLabel(header));
}

export function isFormPGJOvertimeEarningsHeader(header) {
  return /overtime\s+earnings?/.test(normHeaderLabel(header));
}

export function isFormPGJPaymentDateHeader(header) {
  const n = normHeaderLabel(header);
  return (n.includes('date') && n.includes('payment')) || n === 'pay date' || n === 'payment date';
}

/** Resolve Sample Payroll Pay date from cached table / meta / helpers. */
export function resolveFormPGJPayDateFromCache(cached, fallback = '') {
  return String(
    cached?.payDate ||
      cached?.meta?.payDate ||
      cached?.meta?.pay_date ||
      fallback ||
      ''
  ).trim();
}

function matchFormPGJWageTailBucket(header) {
  const n = normHeaderLabel(header);
  if (n.includes('total days worked')) return 'paidDays';
  if (n.includes('minimum rate of wages')) return 'skip';
  if (n.includes('total production') && n.includes('piece rate')) return 'skip';
  if (n.includes('actual wages paid')) return 'basic';
  if (n.includes('house rent allowance')) return 'hra';
  if (n.includes('dearness allowance')) return 'skip';
  if (n.includes('gross amount payable')) return 'grossPay';
  if (n.includes('total hours of overtime')) return 'overtimeHours';
  if (n.includes('overtime earnings')) return 'overtimeEarnings';
  if (n.includes('provident fund')) return 'pf';
  if (n.includes('family pension')) return 'skip';
  if (n.includes('esi contribution') || (n.includes('esi') && n.includes('contribution'))) return 'esi';
  if (n.includes('professional tax')) return 'professionalTax';
  if (n.includes('income tax')) return 'incomeTax';
  if (n.includes('loan') && n.includes('interest')) return 'skip';
  if (n.includes('advances')) return 'skip';
  if (n.includes('other deductions')) return 'otherDeductions';
  if (n.includes('total deduction')) return 'totalDeduction';
  if (n.includes('net payable')) return 'netPay';
  if (n.includes('date of payment')) return 'paymentDate';
  if (n.includes('signature') || n.includes('thumb impression')) return 'signature';
  return '';
}

function unwrapPGJEmployee(empItem) {
  return empItem?.Employee || empItem?.employee || empItem || {};
}

/** Build People match candidates — unwrap nested Employee + row lookup metadata from grid. */
export function buildFormPGJEmployeePayrollCandidates(empItem, row = null) {
  const emp = unwrapPGJEmployee(empItem);
  const candidates = [];
  const seen = new Set();
  const push = (value) => {
    if (!value || typeof value !== 'object' || seen.has(value)) return;
    seen.add(value);
    candidates.push(value);
  };
  push(empItem);
  push(emp);
  const lookupId = String(row?.__employeeLookupId || '').trim();
  if (lookupId) {
    push({
      ...emp,
      Zoho_ID: lookupId,
      EmployeeID: lookupId,
      employee_number: lookupId,
      Employee_Number: lookupId,
    });
  }
  const lookupName = String(row?.__employeeLookupName || '').trim();
  if (lookupName) {
    const parts = lookupName.split(/\s+/).filter(Boolean);
    push({
      ...emp,
      DisplayName: lookupName,
      EmployeeName: lookupName,
      FirstName: parts[0] || '',
      LastName: parts.slice(1).join(' ') || '',
    });
  }
  return candidates;
}

export function applyFormPGJGujaratPayrollToRow(row, payrollRow, headers, helpers = {}) {
  const hdrs = Array.isArray(headers) && headers.length > 0 ? headers : buildFormPGJGujaratCanonicalHeaders();
  const out = row && typeof row === 'object' ? row : {};
  const {
    sanitizeValue = (v) => String(v ?? '').trim(),
    formatStatutoryDateDisplay = (v) => String(v || '').trim(),
    payDate = '',
    monthEndDate = '',
    overwrite = true,
  } = helpers;
  const payroll = resolveFormPGJGujaratPayrollFields(
    payrollRow && !payrollRow.fetch_error ? payrollRow : null,
    { formatStatutoryDateDisplay, payDate, monthEndDate }
  );
  const cellIsEmpty = (header) => {
    const v = String(out[header] || '').trim();
    if (!v) return true;
    const s = v.toLowerCase();
    return /^enter\b/.test(s) || s.includes('enter ');
  };
  const setCell = (header, value, opts = {}) => {
    if (!header || value == null || value === '') return false;
    if (!overwrite && !cellIsEmpty(header)) return false;
    out[header] = sanitizeValue(formatCellValue(value, opts));
    return true;
  };
  let wrote = false;
  const stampAliases = (predicate, value) => {
    hdrs.forEach((header) => {
      if (predicate(header) && setCell(header, value, { allowZero: true })) wrote = true;
    });
    Object.keys(out).forEach((key) => {
      if (String(key).startsWith('__')) return;
      if (predicate(key) && setCell(key, value, { allowZero: true })) wrote = true;
    });
  };

  hdrs.forEach((header) => {
    if (isFormPGJSignatureHeader(header)) return;
    if (isFormPGJWageRateHeader(header)) {
      if (overwrite || cellIsEmpty(header)) out[header] = '';
      return;
    }
    const bucket = matchFormPGJWageTailBucket(header);
    if (!bucket || bucket === 'skip') return;
    if (bucket === 'overtimeHours' || bucket === 'overtimeEarnings') {
      if (setCell(header, FORM_PGJ_OVERTIME_DEFAULT, { allowZero: true })) wrote = true;
      return;
    }
    if (Object.prototype.hasOwnProperty.call(payroll, bucket)) {
      if (setCell(header, payroll[bucket], { allowZero: true })) wrote = true;
    }
  });

  // Force Actual Wages Paid ← basic (header text variants / alias keys).
  if (payroll.basic !== '' && payroll.basic != null) {
    stampAliases(isFormPGJActualWagesPaidHeader, payroll.basic);
  }
  stampAliases(isFormPGJOvertimeHoursHeader, FORM_PGJ_OVERTIME_DEFAULT);
  stampAliases(isFormPGJOvertimeEarningsHeader, FORM_PGJ_OVERTIME_DEFAULT);
  // Working hours From / To ← fixed 9 AM / 5 PM.
  stampAliases(isFormPGJWorkingHoursFromHeader, FORM_PGJ_WORKING_HOURS_FROM);
  stampAliases(isFormPGJWorkingHoursToHeader, FORM_PGJ_WORKING_HOURS_TO);
  // Date of Payment ← Sample Payroll Pay date (stamp all header aliases on the row).
  if (payroll.paymentDate !== '' && payroll.paymentDate != null) {
    stampAliases(isFormPGJPaymentDateHeader, payroll.paymentDate);
  }
  hdrs.forEach((header) => {
    if (isFormPGJWageRateHeader(header)) out[header] = '';
  });
  Object.keys(out).forEach((key) => {
    if (isFormPGJWageRateHeader(key)) out[key] = '';
  });

  return wrote;
}

/** Apply Form P static defaults (OT NIL, Working hours 9 AM–5 PM, clear Wage Rate). */
export function applyFormPGJGujaratStaticDefaults(mappedData, headers, { overwrite = true } = {}) {
  if (!Array.isArray(mappedData) || mappedData.length === 0) return 0;
  const hdrs = Array.isArray(headers) && headers.length > 0 ? headers : buildFormPGJGujaratCanonicalHeaders();
  let hits = 0;
  mappedData.forEach((row) => {
    if (!row || typeof row !== 'object') return;
    let touched = false;
    const setIf = (header, value) => {
      const cur = String(row[header] ?? '').trim();
      const empty = !cur || /^enter\b/i.test(cur) || cur.toLowerCase().includes('enter ');
      if (!overwrite && !empty) return;
      row[header] = value;
      touched = true;
    };
    hdrs.forEach((header) => {
      if (isFormPGJWageRateHeader(header)) setIf(header, '');
      if (isFormPGJOvertimeHoursHeader(header) || isFormPGJOvertimeEarningsHeader(header)) {
        setIf(header, FORM_PGJ_OVERTIME_DEFAULT);
      }
      if (isFormPGJWorkingHoursFromHeader(header)) setIf(header, FORM_PGJ_WORKING_HOURS_FROM);
      if (isFormPGJWorkingHoursToHeader(header)) setIf(header, FORM_PGJ_WORKING_HOURS_TO);
    });
    Object.keys(row).forEach((key) => {
      if (String(key).startsWith('__')) return;
      if (isFormPGJWageRateHeader(key)) setIf(key, '');
      if (isFormPGJOvertimeHoursHeader(key) || isFormPGJOvertimeEarningsHeader(key)) {
        setIf(key, FORM_PGJ_OVERTIME_DEFAULT);
      }
      if (isFormPGJWorkingHoursFromHeader(key)) setIf(key, FORM_PGJ_WORKING_HOURS_FROM);
      if (isFormPGJWorkingHoursToHeader(key)) setIf(key, FORM_PGJ_WORKING_HOURS_TO);
    });
    if (touched) hits += 1;
  });
  return hits;
}

const FORM_PGJ_LEAVE_MONTHS = [
  'jan',
  'feb',
  'mar',
  'apr',
  'may',
  'jun',
  'jul',
  'aug',
  'sep',
  'oct',
  'nov',
  'dec',
];

function parseZohoLeaveDateParts(dateStr) {
  const m = String(dateStr || '').match(/^(\d{1,2})-([A-Za-z]{3})-(\d{4})$/i);
  if (!m) return null;
  const month = FORM_PGJ_LEAVE_MONTHS.indexOf(m[2].toLowerCase());
  if (month < 0) return null;
  return { day: parseInt(m[1], 10), month, year: parseInt(m[3], 10) };
}

function formatZohoLeaveDate(date) {
  const mon = FORM_PGJ_LEAVE_MONTHS[date.getMonth()];
  if (!mon) return '';
  return `${String(date.getDate()).padStart(2, '0')}-${mon.charAt(0).toUpperCase()}${mon.slice(1)}-${date.getFullYear()}`;
}

function zohoLeaveDateInRange(dateStr, monthFrom, monthTo) {
  if (!monthFrom || !monthTo) return true;
  const parts = parseZohoLeaveDateParts(dateStr);
  const fromParts = parseZohoLeaveDateParts(monthFrom);
  const toParts = parseZohoLeaveDateParts(monthTo);
  if (!parts || !fromParts || !toParts) return true;
  const t = Date.UTC(parts.year, parts.month, parts.day);
  const start = Date.UTC(fromParts.year, fromParts.month, fromParts.day);
  const end = Date.UTC(toParts.year, toParts.month, toParts.day);
  return t >= start && t <= end;
}

export function formPGJLeaveCodeFromType(leaveType) {
  const t = String(leaveType || '')
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .trim();
  if (/contingency|casual|\bc\/?l\b|^cl$/.test(t)) return 'CL';
  if (/earned|privilege|annual|accumulated|\be\/?l\b|^el$/.test(t)) return 'EL';
  if (/sick|medical|\bs\/?l\b|^sl$/.test(t)) return 'SL';
  return 'L';
}

function collectFormPGJLeaveDayNumbers(record, monthFrom, monthTo) {
  const days = new Set();
  const addIfInRange = (dateStr) => {
    if (!zohoLeaveDateInRange(dateStr, monthFrom, monthTo)) return;
    const parts = parseZohoLeaveDateParts(dateStr);
    if (parts) days.add(parts.day);
  };

  const positive = extractApprovedLeavePositiveDayKeys(record?.Days ?? record?.days);
  if (positive.length > 0) {
    positive.forEach((d) => addIfInRange(d));
    return [...days];
  }

  const fromParts = parseZohoLeaveDateParts(record?.From ?? record?.from);
  const toParts = parseZohoLeaveDateParts(
    record?.To ?? record?.to ?? record?.From ?? record?.from
  );
  if (!fromParts) return [];
  const end = toParts || fromParts;
  const cursor = new Date(fromParts.year, fromParts.month, fromParts.day);
  const last = new Date(end.year, end.month, end.day);
  while (cursor.getTime() <= last.getTime()) {
    addIfInRange(formatZohoLeaveDate(cursor));
    cursor.setDate(cursor.getDate() + 1);
  }
  return [...days];
}

/**
 * Stamp approved-leave days onto Form P Date of Month_n cells (CL / EL / SL / L).
 * Uses the same worker match as Form O so 3-part names like Patel Chandrakant Virabhai attach.
 */
export function applyFormPGJGujaratApprovedLeaveToDayColumns(
  mappedData,
  employeesForMapping,
  tableHeaders,
  approvedLeaveRecords,
  options = {}
) {
  if (!Array.isArray(mappedData) || mappedData.length === 0) return 0;
  if (!Array.isArray(approvedLeaveRecords) || approvedLeaveRecords.length === 0) return 0;

  const monthFrom = options.monthFrom || '';
  const monthTo = options.monthTo || '';
  const hdrs =
    Array.isArray(tableHeaders) && tableHeaders.length > 0
      ? tableHeaders
      : buildFormPGJGujaratCanonicalHeaders();
  const dayHeaders = hdrs
    .map((header) => {
      const m = String(header || '')
        .trim()
        .match(/^Date of Month_(\d{1,2})$/i);
      if (!m) return null;
      const day = parseInt(m[1], 10);
      if (day < 1 || day > 31) return null;
      return { header, day };
    })
    .filter(Boolean);
  if (dayHeaders.length === 0) return 0;

  let hits = 0;
  mappedData.forEach((row, index) => {
    if (!row || typeof row !== 'object') return;
    const empItem = employeesForMapping?.[index];
    const emp = unwrapPGJEmployee(empItem);
    const workerName = String(row['Name of the Worker'] || row.__employeeLookupName || '').trim();

    const matches = approvedLeaveRecords.filter((rec) => {
      if (!rec) return false;
      if (monthFrom && monthTo && !approvedLeaveRecordOverlapsFormOGJMonth(rec, monthFrom, monthTo)) {
        return false;
      }
      return formOGJRecordMatchesWorker(rec, workerName, emp, {
        collectEmployeeIdCandidates: options.collectEmployeeIdCandidates,
      });
    });
    if (matches.length === 0) return;

    let wrote = false;
    matches.forEach((rec) => {
      const code = formPGJLeaveCodeFromType(readApprovedLeaveLeaveType(rec));
      collectFormPGJLeaveDayNumbers(rec, monthFrom, monthTo).forEach((day) => {
        const col = dayHeaders.find((h) => h.day === day);
        if (!col) return;
        row[col.header] = code;
        wrote = true;
      });
    });
    if (wrote) hits += 1;
  });
  return hits;
}

export function enrichFormPGJGujaratPayrollRows(mappedData, employees, headers, helpers = {}) {
  const hdrs = Array.isArray(headers) && headers.length > 0 ? headers : buildFormPGJGujaratCanonicalHeaders();
  const {
    resolvePayrollRow = null,
    sanitizeValue = (v) => String(v ?? '').trim(),
    formatStatutoryDateDisplay = (v) => String(v || '').trim(),
    payDate = '',
    monthEndDate = '',
    overwrite = true,
    rowIndexOffset = 0,
  } = helpers;
  if (!Array.isArray(mappedData) || mappedData.length === 0 || typeof resolvePayrollRow !== 'function') {
    return 0;
  }
  let hits = 0;
  mappedData.forEach((row, rowIndex) => {
    const globalRowIndex = rowIndexOffset + rowIndex;
    const empItem = employees?.[rowIndex];
    let payrollRow = null;
    if (typeof resolvePayrollRow === 'function') {
      payrollRow =
        resolvePayrollRow(empItem, row, globalRowIndex) ||
        resolvePayrollRow(unwrapPGJEmployee(empItem), row, globalRowIndex);
      if (!payrollRow || payrollRow.fetch_error) {
        const candidates = buildFormPGJEmployeePayrollCandidates(empItem, row);
        for (let ci = 0; ci < candidates.length; ci += 1) {
          const candidate = candidates[ci];
          const hit = resolvePayrollRow(candidate, row, globalRowIndex);
          if (hit && !hit.fetch_error) {
            payrollRow = hit;
            break;
          }
        }
      }
    }
    const wrote = applyFormPGJGujaratPayrollToRow(row, payrollRow, hdrs, {
      sanitizeValue,
      formatStatutoryDateDisplay,
      payDate,
      monthEndDate,
      overwrite,
    });
    applyFormPGJGujaratStaticDefaults([row], hdrs, { overwrite: true });
    if (wrote) hits += 1;
  });
  return hits;
}

export function enrichFormPGJGujaratDisplayHeader(formHeader, fileName, item, headers, sheetText = '') {
  if (!isFormPGJGujaratContext(formHeader, item, fileName, sheetText, headers)) return formHeader;
  return {
    ...(formHeader || {}),
    title: formHeader?.title || 'FORM P',
    formPGJGujaratTableLayout: true,
    fields: ensureFormPGJGujaratHeaderFields(
      Array.isArray(formHeader?.fields) && formHeader.fields.length > 0
        ? formHeader.fields
        : FORM_PGJ_DEFAULT_HEADER_FIELDS
    ),
  };
}

const excelCellValueToString = (val) => {
  if (val == null) return '';
  if (typeof val === 'string' || typeof val === 'number' || typeof val === 'boolean') return String(val);
  if (val instanceof Date) return val.toISOString();
  if (typeof val === 'object') {
    if (Array.isArray(val.richText)) return val.richText.map((rt) => rt?.text || '').join('');
    if (val.text != null) return String(val.text);
    if (val.result != null) return String(val.result);
  }
  return '';
};

function buildMergeTopLeftResolver(worksheet) {
  const cache = new Map();
  return (r, c) => {
    const key = `${r}:${c}`;
    if (cache.has(key)) return cache.get(key);
    let topLeft = { r, c };
    const merges = worksheet.model?.merges;
    if (Array.isArray(merges)) {
      for (let mi = 0; mi < merges.length; mi += 1) {
        const parts = String(merges[mi] || '').split(':');
        if (parts.length !== 2) continue;
        const tl = worksheet.getCell(parts[0]);
        const br = worksheet.getCell(parts[1]);
        if (!tl || !br) continue;
        const r1 = tl.fullAddress?.row ?? tl.row;
        const c1 = tl.fullAddress?.col ?? tl.col;
        const r2 = br.fullAddress?.row ?? br.row;
        const c2 = br.fullAddress?.col ?? br.col;
        if (r >= r1 && r <= r2 && c >= c1 && c <= c2) {
          topLeft = { r: r1, c: c1 };
          break;
        }
      }
    }
    cache.set(key, topLeft);
    return topLeft;
  };
}

function formPGJTemplateBucket(label) {
  const n = normHeaderLabel(label)
    .replace(/\b\d+\b/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  if (!n) return '';
  if (/^sr\.?\s*no/.test(n) || /^s\.?\s*no/.test(n)) return 'sno';
  if (/full\s+name\s+of\s+the\s+worker/.test(n) || /name\s+of\s+the\s+worker/.test(n)) return 'workerName';
  if (/^designation/.test(n) || (/designation/.test(n) && /nature\s+of\s+work/.test(n))) return 'designation';
  if (/^age$/.test(n)) return 'age';
  if (/^sex$/.test(n)) return 'sex';
  if (/entry\s+into\s+service/.test(n)) return 'dateOfEntry';
  if (/^working\s+hours$/.test(n) || (n.includes('working hours') && !/from|to/.test(n))) {
    return 'workingHours';
  }
  if (/working\s+hours/.test(n) && /\bfrom\b/.test(n)) return 'workingHoursFrom';
  if (/working\s+hours/.test(n) && /\bto\b/.test(n)) return 'workingHoursTo';
  if (/interval\s+for\s+rest/.test(n) && /\bfrom\b/.test(n)) return 'intervalFrom';
  if (/interval\s+for\s+rest/.test(n) && /\bto\b/.test(n)) return 'intervalTo';
  if (/^interval\s+for\s+rest$/.test(n) || (/interval\s+for\s+rest/.test(n) && !/from|to/.test(n))) {
    return 'intervalRest';
  }
  if (/^date\s+of\s+(the\s+)?month$/.test(n)) return 'dateBand';
  // Day numbers — use raw digit label (do not strip numbers above)
  {
    const rawDigit = formPGJGujaratHeaderNorm(label).replace(/[^0-9]/g, '');
    if (/^\d{1,2}$/.test(rawDigit) && /^\d{1,2}$/.test(formPGJGujaratHeaderNorm(label).trim())) {
      const d = parseInt(rawDigit, 10);
      if (d >= 1 && d <= 31) return `day:${d}`;
    }
  }
  if (/father|husband/.test(n)) return 'fatherName';
  if (/wage\s+rate/.test(n) && !/minimum/.test(n)) return 'wageRate';
  const wageBucket = matchFormPGJWageTailBucket(label);
  if (wageBucket && wageBucket !== 'skip') return wageBucket;
  if (isFormPGJWageTailHeader(label) || /minimum\s+rate\s+of\s+wages/.test(n)) {
    if (/minimum\s+rate/.test(n)) return 'skip';
    if (/total\s+production/.test(n)) return 'skip';
    if (/dearness/.test(n)) return 'skip';
    if (/family\s+pension|loan|advances/.test(n)) {
      return 'skip';
    }
    if (/signature|thumb/.test(n)) return 'signature';
  }
  if (/actual\s+wages\s+paid/.test(n)) return 'basic';
  if (/house\s+rent/.test(n)) return 'hra';
  if (/gross\s+amount/.test(n)) return 'grossPay';
  if (/total\s+hours?\s+of\s+overtime/.test(n)) return 'overtimeHours';
  if (/overtime\s+earnings?/.test(n)) return 'overtimeEarnings';
  if (/total\s+days?\s+worked/.test(n)) return 'paidDays';
  if (/net\s+payable/.test(n)) return 'netPay';
  if (/date\s+of\s+payment/.test(n)) return 'paymentDate';
  if (/esi/.test(n)) return 'esi';
  if (/provident\s+fund/.test(n)) return 'pf';
  if (/professional\s+tax/.test(n)) return 'professionalTax';
  if (/income\s+tax/.test(n)) return 'incomeTax';
  if (/other\s+deduction/.test(n)) return 'otherDeductions';
  if (/total\s+deduction/.test(n)) return 'totalDeduction';
  return '';
}

/** Detect original Form P template columns (Age/Sex/Working Hours/Interval/days/wage tail). */
export function detectFormPGJGujaratTableLayout(worksheet) {
  const getMergeTopLeft = buildMergeTopLeftResolver(worksheet);
  const getText = (r, c) => {
    const tl = getMergeTopLeft(r, c);
    return excelCellValueToString(worksheet.getCell(tl.r, tl.c)?.value).trim();
  };
  const norm = (txt) => formPGJGujaratHeaderNorm(txt);

  let headerRow = -1;
  let startCol = 1;
  const maxR = Math.max(40, worksheet.rowCount || 40);
  const maxC = Math.max(80, worksheet.columnCount || 0, worksheet.actualColumnCount || 0);

  for (let r = 1; r <= maxR; r += 1) {
    const parts = [];
    for (let c = 1; c <= maxC; c += 1) {
      const t = norm(getText(r, c));
      if (t) parts.push(t);
    }
    const joined = parts.join(' ');
    if (
      /sr\.?\s*no/.test(joined) &&
      (/full\s+name\s+of\s+the\s+worker/.test(joined) || /name\s+of\s+the\s+worker/.test(joined)) &&
      (/date\s+of\s+(the\s+)?month/.test(joined) || /working\s+hours/.test(joined))
    ) {
      headerRow = r;
      for (let c = 1; c <= maxC; c += 1) {
        const t = norm(getText(r, c));
        if (/^sr\.?\s*no/.test(t) || t === 'sr no') {
          startCol = c;
          break;
        }
      }
      break;
    }
  }
  if (headerRow < 1) return null;

  const templateCols = [];
  const seenBuckets = new Set();
  let day1Col = -1;
  const subRow = headerRow + 1;

  for (let c = startCol; c <= startCol + 70; c += 1) {
    const main = getText(headerRow, c);
    const sub = getText(subRow, c);
    const mainN = norm(main);
    const subN = norm(sub);

    // Day numbers on sub-row under Date of Month band
    if (/^\d{1,2}$/.test(subN) || /^\d{1,2}$/.test(mainN)) {
      const day = parseInt(subN || mainN, 10);
      if (day >= 1 && day <= 31) {
        const pastWorker =
          seenBuckets.has('workerName') ||
          seenBuckets.has('designation') ||
          seenBuckets.has('workingHours') ||
          seenBuckets.has('workingHoursFrom') ||
          day1Col > 0;
        if (pastWorker) {
          if (day1Col < 0) day1Col = c;
          const bucket = `day:${day}`;
          if (!seenBuckets.has(bucket)) {
            seenBuckets.add(bucket);
            templateCols.push({ col: c, bucket, label: `Date of Month_${day}` });
          }
          continue;
        }
      }
    }

    let label = main || sub;
    const sub2 = getText(headerRow + 2, c);
    // Prefer parent label when sub is From/To under Working hours / Interval
    if (/^(from|to)$/i.test(sub) && /working\s+hours|interval\s+for\s+rest/i.test(main)) {
      label = `${main}_${sub}`;
    } else if (
      // DEDUCTION is a merged parent — real labels sit on the sub-row (PF / PT / Income Tax / …).
      (/^deduction$/i.test(mainN) || /^deductions?$/i.test(mainN)) &&
      sub &&
      !/^(from|to|\d{1,2})$/i.test(sub)
    ) {
      label = sub;
    } else if (
      (/^deduction$/i.test(mainN) || /^deductions?$/i.test(mainN)) &&
      sub2 &&
      !/^(from|to|\d{1,2})$/i.test(sub2)
    ) {
      label = sub2;
    } else if (main && sub && !formPGJTemplateBucket(main) && formPGJTemplateBucket(sub)) {
      label = sub;
    } else if (main && sub2 && !formPGJTemplateBucket(main) && formPGJTemplateBucket(sub2)) {
      label = sub2;
    } else if (!main && sub) {
      label = sub;
    } else if (!main && !sub && sub2) {
      label = sub2;
    }

    const bucket = formPGJTemplateBucket(label);
    if (!bucket || bucket === 'skip' || bucket === 'dateBand') {
      // Last resort: try sub / sub2 even when main looked like a section title.
      const alt = [sub, sub2].find((t) => t && formPGJTemplateBucket(t));
      if (alt) {
        const altBucket = formPGJTemplateBucket(alt);
        if (altBucket && altBucket !== 'skip' && altBucket !== 'dateBand') {
          if (!(seenBuckets.has(altBucket) && !String(altBucket).startsWith('day:'))) {
            seenBuckets.add(altBucket);
            templateCols.push({ col: c, bucket: altBucket, label: alt || altBucket });
          }
        }
      }
      continue;
    }
    if (seenBuckets.has(bucket) && !String(bucket).startsWith('day:')) continue;
    seenBuckets.add(bucket);
    templateCols.push({ col: c, bucket, label: label || bucket });
  }

  // Recover wage-tail columns missed because of merged DEDUCTION / multi-row headers.
  const neededWageTail = [
    'pf',
    'professionalTax',
    'incomeTax',
    'otherDeductions',
    'totalDeduction',
    'paymentDate',
    'signature',
    'esi',
    'overtimeHours',
    'overtimeEarnings',
    'grossPay',
    'netPay',
  ];
  const missing = neededWageTail.filter((b) => !seenBuckets.has(b));
  if (missing.length > 0) {
    for (let r = Math.max(1, headerRow - 1); r <= headerRow + 3; r += 1) {
      for (let c = startCol; c <= startCol + 70; c += 1) {
        // Prefer direct cell text so merged "DEDUCTION" parents do not hide child labels.
        const direct = excelCellValueToString(worksheet.getCell(r, c)?.value).trim();
        const merged = getText(r, c);
        const candidates = [direct, merged].filter(Boolean);
        for (let i = 0; i < candidates.length; i += 1) {
          const bucket = formPGJTemplateBucket(candidates[i]);
          if (!bucket || !missing.includes(bucket)) continue;
          if (seenBuckets.has(bucket)) continue;
          seenBuckets.add(bucket);
          templateCols.push({ col: c, bucket, label: candidates[i] });
          const mi = missing.indexOf(bucket);
          if (mi >= 0) missing.splice(mi, 1);
          break;
        }
        if (missing.length === 0) break;
      }
      if (missing.length === 0) break;
    }
  }

  // Fill missing day columns from day1Col if only some were found
  if (day1Col > 0) {
    for (let d = 1; d <= 31; d += 1) {
      const bucket = `day:${d}`;
      if (seenBuckets.has(bucket)) continue;
      seenBuckets.add(bucket);
      templateCols.push({ col: day1Col + (d - 1), bucket, label: `Date of Month_${d}` });
    }
  }

  templateCols.sort((a, b) => a.col - b.col);

  let dataStartRow = headerRow + 2;
  for (let r = headerRow + 1; r <= headerRow + 4; r += 1) {
    let seqHits = 0;
    for (let hi = 0; hi < 9; hi += 1) {
      const t = String(getText(r, startCol + hi) || '').replace(/\s+/g, '').trim();
      if (t === String(hi + 1)) seqHits += 1;
    }
    if (seqHits >= 4) {
      dataStartRow = r + 1;
      break;
    }
    if (/^(from|to|\d{1,2})$/i.test(getText(r, startCol + 6)) || /^(from|to)$/i.test(getText(r, startCol + 7))) {
      dataStartRow = Math.max(dataStartRow, r + 1);
    }
  }

  if (templateCols.length < 6) return null;
  return { headerRow, dataStartRow, templateCols, startCol, day1Col };
}

/** Count Date of Month marks that count as days worked (present / paid leave). */
export function countFormPGJGujaratDaysWorkedFromAttendance(row) {
  if (!row || typeof row !== 'object') return '';
  let n = 0;
  for (let d = 1; d <= 31; d += 1) {
    const v = String(
      row[`Date of Month_${d}`] ||
        row[`Date of the Month_${d}`] ||
        row[`Date of Month_${Number(d)}`] ||
        ''
    )
      .trim()
      .toUpperCase();
    if (!v) continue;
    if (/^(A|AB|ABSENT|WO|W\/O|OFF|HO|NH|FH)$/.test(v)) continue;
    n += 1;
  }
  return n > 0 ? n : '';
}

function getFormPGJExportValueForBucket(row, bucket, idx = 0) {
  if (!row || typeof row !== 'object') return '';
  if (bucket === 'sno') {
    const direct =
      row['Sr. No.'] ?? row['Sr. No'] ?? row['Sr No'] ?? row.sno ?? '';
    if (direct != null && String(direct).trim() !== '') return String(direct).trim();
    return String(idx + 1);
  }
  if (bucket === 'workerName') {
    return (
      row['Name of the Worker'] ||
      row['Full Name of the Worker'] ||
      row['Full Name of the worker'] ||
      row['Name of Workers'] ||
      row.__employeeLookupName ||
      ''
    );
  }
  if (bucket === 'designation') {
    return row.Designation || row['Designation'] || '';
  }
  if (bucket === 'fatherName') {
    return row["Father's / Husband's Name"] || row['Father Name'] || '';
  }
  if (bucket === 'dateOfEntry') {
    return (
      row['Date of entry into service'] ||
      row['Date of Entry into service'] ||
      row['Date of Entry into Service'] ||
      ''
    );
  }
  if (bucket === 'workingHoursFrom') {
    return row['Working hours_From'] || row['Working Hours_From'] || '';
  }
  if (bucket === 'workingHoursTo') {
    return row['Working hours_To'] || row['Working Hours_To'] || '';
  }
  if (bucket === 'workingHours') {
    const from = String(row['Working hours_From'] || row['Working Hours_From'] || '').trim();
    const to = String(row['Working hours_To'] || row['Working Hours_To'] || '').trim();
    if (from && to) return `${from} - ${to}`;
    return from || to || row['Working Hours'] || row['Working hours'] || '';
  }
  if (bucket === 'intervalRest' || bucket === 'intervalFrom' || bucket === 'intervalTo') {
    return (
      row['Interval for Rest'] ||
      row['Interval for Rest_From'] ||
      row['Interval for Rest_To'] ||
      ''
    );
  }
  if (bucket === 'age' || bucket === 'sex' || bucket === 'wageRate') return '';
  if (String(bucket).startsWith('day:')) {
    const d = bucket.split(':')[1];
    return (
      row[`Date of Month_${d}`] ||
      row[`Date of the Month_${d}`] ||
      row[`Date of Month_${Number(d)}`] ||
      ''
    );
  }
  if (bucket === 'basic') {
    return (
      row['Actual Wages Paid Rs.'] ||
      row['Actual Wages Paid Rs'] ||
      row.basic ||
      ''
    );
  }
  if (bucket === 'overtimeHours') {
    return (
      row['Total hours of Overtime during the month'] ||
      row['Total hours of overtime worked during the month'] ||
      FORM_PGJ_OVERTIME_DEFAULT
    );
  }
  if (bucket === 'overtimeEarnings') {
    return row['Overtime Earnings Rs.'] || row['Overtime earnings Rs.'] || FORM_PGJ_OVERTIME_DEFAULT;
  }
  if (bucket === 'hra') {
    return row['House Rent Allowance Rs.'] || row['House Rent Allowance Paid Rs.'] || '';
  }
  if (bucket === 'grossPay') {
    return row['Gross Amount Payable Rs.'] || '';
  }
  if (bucket === 'paidDays') {
    const direct =
      row['Total Days Worked'] ||
      row['Total Days worked'] ||
      row.paidDays ||
      '';
    if (direct != null && String(direct).trim() !== '') return direct;
    return countFormPGJGujaratDaysWorkedFromAttendance(row);
  }
  if (bucket === 'netPay') return row['Net Payable Rs.'] || '';
  if (bucket === 'paymentDate') return row['Date of Payment'] || '';
  if (bucket === 'esi') return row['ESI Contribution Rs.'] || '';
  if (bucket === 'pf') return row['Provident Fund Rs.'] || '';
  if (bucket === 'professionalTax') return row['Professional Tax Rs.'] || '';
  if (bucket === 'incomeTax') return row['Income Tax Rs.'] || '';
  if (bucket === 'otherDeductions') return row['Other Deductions Rs.'] || '';
  if (bucket === 'totalDeduction') return row['Total Deduction Rs.'] || '';
  if (bucket === 'signature') return '';

  // Generic: find row key whose template bucket matches
  for (const [k, v] of Object.entries(row)) {
    if (String(k).startsWith('__')) continue;
    if (formPGJTemplateBucket(k) === bucket && v != null && String(v).trim() !== '') {
      return v;
    }
  }
  return '';
}

function findFormPGJGujaratSignatureColumn(worksheet, layout) {
  if (!worksheet || !layout) return -1;
  const fromTemplate = (layout.templateCols || []).find((c) => c?.bucket === 'signature');
  if (fromTemplate?.col > 0) return fromTemplate.col;
  const headerRow = layout.headerRow || 1;
  const startCol = layout.startCol || 1;
  const maxC = Math.max(80, worksheet.columnCount || 0, worksheet.actualColumnCount || 0);
  for (let r = headerRow; r <= headerRow + 3; r += 1) {
    for (let c = startCol; c <= startCol + 80 && c <= maxC; c += 1) {
      const t = formPGJGujaratHeaderNorm(excelCellValueToString(worksheet.getCell(r, c)?.value));
      if (t && (t.includes('signature') || t.includes('thumb impression'))) return c;
    }
  }
  return -1;
}

/**
 * Fill Form P header Name of the Establishment / employer / Month on Excel download.
 */
export function writeFormPGJGujaratHeaderMetaToExcelJsWorksheet(
  worksheet,
  headerFormData = {},
  { headerRowEnd = 12 } = {}
) {
  if (!worksheet || !headerFormData || typeof headerFormData !== 'object') return false;
  const monthVal = String(
    headerFormData.form_p_gj_month || headerFormData.form_x_month || ''
  ).trim();
  const employerVal = formPGJEmployerDisplayName(
    headerFormData.form_p_gj_employer ||
      headerFormData.statutory_principal_employer ||
      headerFormData.statutory_employer_name_address ||
      headerFormData.form_t_employer ||
      ''
  );
  const establishmentVal = String(
    headerFormData.form_p_gj_establishment ||
      headerFormData.statutory_establishment_name_address ||
      headerFormData.statutory_establishment_name ||
      headerFormData.form_t_establishment_name_address ||
      ''
  ).trim();
  if (!monthVal && !employerVal && !establishmentVal) return false;

  const maxR = Math.max(1, Math.min(Number(headerRowEnd) || 12, worksheet.rowCount || 12));
  const maxC = Math.min(20, Math.max(12, worksheet.columnCount || 12, worksheet.actualColumnCount || 12));
  let wroteEmployer = false;
  let wroteMonth = false;
  let wroteEstablishment = false;

  for (let r = 1; r <= maxR; r += 1) {
    for (let c = 1; c <= maxC; c += 1) {
      const raw = excelCellValueToString(worksheet.getCell(r, c)?.value).replace(/\s+/g, ' ').trim();
      if (!raw) continue;
      const labelOnly = raw.split(':')[0].trim();

      if (
        establishmentVal &&
        /name\s+of\s+(?:the\s+)?establishment/i.test(labelOnly) &&
        !/employer|already/i.test(labelOnly)
      ) {
        const cell = worksheet.getCell(r, c);
        cell.value = `Name of the Establishment : ${establishmentVal}`;
        wroteEstablishment = true;
        continue;
      }

      if (
        employerVal &&
        /name\s+of\s+(?:the\s+)?employer/i.test(labelOnly) &&
        !/address/i.test(labelOnly)
      ) {
        const cell = worksheet.getCell(r, c);
        cell.value = `Name of the employer: ${employerVal}`;
        wroteEmployer = true;
        continue;
      }

      if (monthVal && /^month\s*:?\s*$/i.test(raw)) {
        worksheet.getCell(r, c).value = `Month: ${monthVal}`;
        wroteMonth = true;
        continue;
      }
      if (monthVal && /^month\s*:/i.test(raw) && !/year/i.test(raw) && !/ending/i.test(raw)) {
        worksheet.getCell(r, c).value = `Month: ${monthVal}`;
        wroteMonth = true;
      }
    }
  }
  return wroteEmployer || wroteMonth || wroteEstablishment;
}

/** Show full Establishment / Employer text (wrap + wide merge). */
export function ensureFormPGJGujaratHeaderIdentityRowsVisible(
  worksheet,
  { headerRowEnd = 12, mergeToCol = 12 } = {}
) {
  if (!worksheet) return false;
  const maxR = Math.max(1, Math.min(Number(headerRowEnd) || 12, worksheet.rowCount || 12));
  const mergeEnd = Math.max(8, Number(mergeToCol) || 12);
  let applied = false;
  for (let r = 1; r <= maxR; r += 1) {
    for (let c = 1; c <= 8; c += 1) {
      const raw = excelCellValueToString(worksheet.getCell(r, c)?.value)
        .replace(/\s+/g, ' ')
        .trim();
      if (!raw) continue;
      const labelOnly = raw.split(':')[0].trim();
      const isEst =
        /name\s+of\s+(?:the\s+)?establishment/i.test(labelOnly) &&
        !/employer|already/i.test(labelOnly);
      const isEmp =
        /name\s+of\s+(?:the\s+)?employer/i.test(labelOnly) && !/address/i.test(labelOnly);
      if (!isEst && !isEmp) continue;
      formPGJUnmergeCovering(worksheet, r, c, r, mergeEnd);
      try {
        if (mergeEnd > c) worksheet.mergeCells(r, c, r, mergeEnd);
      } catch (_) {
        /* already merged */
      }
      const cell = worksheet.getCell(r, c);
      cell.alignment = {
        ...(cell.alignment || {}),
        wrapText: true,
        vertical: 'top',
        horizontal: 'left',
      };
      const lines = Math.max(2, Math.ceil(raw.length / 64));
      const wsRow = worksheet.getRow(r);
      if (wsRow) wsRow.height = Math.min(72, Math.max(Number(wsRow.height) || 0, lines * 16, 32));
      applied = true;
      break;
    }
  }
  return applied;
}

/** Form P table header + body font on Excel/PDF download. */
export const FORM_PGJ_TABLE_FONT_SIZE = 8;
export const FORM_PGJ_DATE_OF_MONTH_GROUP_LABEL = 'Date of Month (9)';

function formPGJIsDateOfMonthGroupLabel(text) {
  const t = formPGJGujaratHeaderNorm(text);
  if (!t) return false;
  if (/date\s+of\s+payment/.test(t) || /entry\s+into\s+service/.test(t)) return false;
  return /^date\s+of\s+(the\s+)?month/.test(t);
}

function formPGJUnmergeCovering(worksheet, r1, c1, r2, c2) {
  const merges = Array.isArray(worksheet.model?.merges) ? [...worksheet.model.merges] : [];
  merges.forEach((label) => {
    const parts = String(label || '').split(':');
    if (parts.length !== 2) return;
    try {
      const tl = worksheet.getCell(parts[0]);
      const br = worksheet.getCell(parts[1]);
      if (!tl || !br || br.row < r1 || tl.row > r2 || br.col < c1 || tl.col > c2) return;
      worksheet.unMergeCells(label);
    } catch (_) {
      /* ignore */
    }
  });
}

/**
 * Merge Excel "Date of Month (9)" across the day 1–daysInMonth columns, like the Gujarat template.
 */
export function ensureFormPGJGujaratDateOfMonthGroupMerge(worksheet, layout, daysInMonth = 31) {
  if (!worksheet || !layout?.headerRow) return false;
  const headerRow = layout.headerRow;
  const subRow = headerRow + 1;
  const getMergeTopLeft = buildMergeTopLeftResolver(worksheet);
  const getText = (r, c) => {
    const tl = getMergeTopLeft(r, c);
    return excelCellValueToString(worksheet.getCell(tl.r, tl.c)?.value).trim();
  };
  const subDirect = (c) => excelCellValueToString(worksheet.getCell(subRow, c)?.value).trim();
  const startCol = layout.startCol || 1;
  const maxC = Math.max(80, worksheet.columnCount || 0, worksheet.actualColumnCount || 0);
  const days = clampFormPGJGujaratDaysInMonth(daysInMonth);

  let bandStart = -1;
  for (let c = startCol; c <= startCol + 70 && c <= maxC; c += 1) {
    if (formPGJIsDateOfMonthGroupLabel(getText(headerRow, c))) {
      bandStart = c;
      break;
    }
  }
  if (bandStart < 0 && layout.day1Col > 0) bandStart = layout.day1Col;
  if (bandStart < 0) return false;

  formPGJUnmergeCovering(worksheet, headerRow, bandStart, headerRow, bandStart + 30);

  let bandEnd = bandStart;
  for (let c = bandStart; c <= bandStart + 30 && c <= maxC; c += 1) {
    const main = excelCellValueToString(worksheet.getCell(headerRow, c)?.value).trim();
    const sub = formPGJGujaratHeaderNorm(subDirect(c)).replace(/[()]/g, '');
    const isDateLabel = formPGJIsDateOfMonthGroupLabel(main) || formPGJIsDateOfMonthGroupLabel(getText(headerRow, c));
    const dayNum = parseInt(sub, 10);
    const isDay = /^\d{1,2}$/.test(sub) && dayNum >= 1 && dayNum <= 31;
    if (isDay && dayNum > days) break;
    if (isDay || (c === bandStart && isDateLabel)) {
      bandEnd = c;
      continue;
    }
    break;
  }
  if (!(bandEnd > bandStart)) return false;

  formPGJUnmergeCovering(worksheet, headerRow, bandStart, headerRow, bandEnd);
  for (let c = bandStart + 1; c <= bandEnd; c += 1) {
    try {
      worksheet.getCell(headerRow, c).value = null;
    } catch (_) {
      /* ignore */
    }
  }
  try {
    worksheet.mergeCells(headerRow, bandStart, headerRow, bandEnd);
  } catch (_) {
    /* already merged */
  }
  const cell = worksheet.getCell(headerRow, bandStart);
  cell.value = FORM_PGJ_DATE_OF_MONTH_GROUP_LABEL;
  cell.alignment = {
    ...(cell.alignment || {}),
    horizontal: 'center',
    vertical: 'middle',
    wrapText: true,
  };
  cell.font = {
    ...(cell.font || {}),
    bold: true,
    size: FORM_PGJ_TABLE_FONT_SIZE,
  };
  return true;
}

/**
 * Keep Date of Month day numbers 1..daysInMonth (June=30, July=31) and hide extras.
 */
export function applyFormPGJGujaratMonthDayColumns(worksheet, layout, daysInMonth = 31) {
  if (!worksheet || !layout?.headerRow) return false;
  const days = clampFormPGJGujaratDaysInMonth(daysInMonth);
  const headerRow = layout.headerRow;
  const subRow = headerRow + 1;
  const dataStart = layout.dataStartRow || headerRow + 2;
  const getMergeTopLeft = buildMergeTopLeftResolver(worksheet);
  const getText = (r, c) => {
    const tl = getMergeTopLeft(r, c);
    return excelCellValueToString(worksheet.getCell(tl.r, tl.c)?.value).trim();
  };
  let bandStart = layout.day1Col > 0 ? layout.day1Col : -1;
  if (bandStart < 0) {
    const startCol = layout.startCol || 1;
    const maxC = Math.max(80, worksheet.columnCount || 0, worksheet.actualColumnCount || 0);
    for (let c = startCol; c <= startCol + 70 && c <= maxC; c += 1) {
      if (formPGJIsDateOfMonthGroupLabel(getText(headerRow, c))) {
        bandStart = c;
        break;
      }
    }
  }
  if (!(bandStart > 0)) return false;

  let changed = false;
  for (let offset = 0; offset < 31; offset += 1) {
    const col = bandStart + offset;
    const sub = formPGJGujaratHeaderNorm(
      excelCellValueToString(worksheet.getCell(subRow, col)?.value)
    ).replace(/[()]/g, '');
    const dayNum = parseInt(sub, 10);
    const isDay = /^\d{1,2}$/.test(sub) && dayNum >= 1 && dayNum <= 31;
    if (!isDay) continue;
    if (dayNum <= days) {
      try {
        worksheet.getColumn(col).hidden = false;
      } catch (_) {
        /* ignore */
      }
      continue;
    }
    changed = true;
    try {
      worksheet.getColumn(col).hidden = true;
    } catch (_) {
      /* ignore */
    }
    for (let r = headerRow; r < dataStart; r += 1) {
      const raw = excelCellValueToString(worksheet.getCell(r, col)?.value).trim();
      if (/^\(?\s*\d{1,2}\s*\)?$/.test(raw) && parseInt(raw.replace(/[()]/g, ''), 10) === dayNum) {
        worksheet.getCell(r, col).value = null;
      }
    }
    const lastClear = Math.max(dataStart + 40, worksheet.rowCount || dataStart);
    for (let r = dataStart; r <= lastClear; r += 1) {
      try {
        worksheet.getCell(r, col).value = null;
      } catch (_) {
        /* ignore */
      }
    }
  }
  return changed;
}

function applyFormPGJGujaratTableFontSize(
  worksheet,
  { headerRow, lastRow, colFrom, colTo, size = FORM_PGJ_TABLE_FONT_SIZE } = {}
) {
  if (!worksheet || !(headerRow > 0) || !(lastRow >= headerRow) || !(colTo >= colFrom)) return;
  for (let r = headerRow; r <= lastRow; r += 1) {
    for (let c = colFrom; c <= colTo; c += 1) {
      const cell = worksheet.getCell(r, c);
      cell.font = { ...(cell.font || {}), size };
    }
  }
}

export async function buildFormPGJGujaratWorkbookWithTemplateStyles({
  templateArrayBuffer,
  mappedData,
  headersToUse,
  parsedFormHeader,
  formFileName,
  headerFormData,
  formatStatutoryDateDisplay = null,
  daysInMonth = null,
}) {
  const ExcelJS = (await import('exceljs')).default;
  const { applyExcelJSFullBoxBordersToRange, ensureExcelJSDataRowsWithBorders } = await import(
    '../../utils/excelTableBorders'
  );
  const { writeStatutoryHeaderFieldsToExcelJsWorksheet } = await import(
    '../../utils/statutorySiteCompanyHeaders'
  );

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(templateArrayBuffer);
  const worksheet = workbook.worksheets[0];
  if (!worksheet) throw new Error('Form P template worksheet not found.');

  const layout = detectFormPGJGujaratTableLayout(worksheet);
  if (!layout) throw new Error('Could not locate Gujarat Form P table header row.');

  const monthDays = clampFormPGJGujaratDaysInMonth(
    daysInMonth != null
      ? daysInMonth
      : resolveFormPGJGujaratDaysInMonth(
          headerFormData?.form_p_gj_month || headerFormData?.form_x_month || '',
          headerFormData?.form_x_year || headerFormData?.form_p_gj_year || ''
        )
  );

  const { dataStartRow, templateCols, startCol } = layout;
  const formatDate =
    typeof formatStatutoryDateDisplay === 'function'
      ? formatStatutoryDateDisplay
      : (v) => String(v || '').trim();

  const parsedForWrite = parsedFormHeader
    ? {
        ...parsedFormHeader,
        fields: ensureFormPGJGujaratHeaderFields(parsedFormHeader.fields),
      }
    : { fields: [...FORM_PGJ_DEFAULT_HEADER_FIELDS] };

  if (headerFormData && typeof headerFormData === 'object') {
    writeStatutoryHeaderFieldsToExcelJsWorksheet(worksheet, {
      headerFormData,
      parsedFormHeader: parsedForWrite,
      headerRowEnd: Math.max(1, layout.headerRow - 1),
      maxScanCols: 80,
    });
    writeFormPGJGujaratHeaderMetaToExcelJsWorksheet(worksheet, headerFormData, {
      headerRowEnd: Math.max(1, layout.headerRow - 1),
    });
    ensureFormPGJGujaratHeaderIdentityRowsVisible(worksheet, {
      headerRowEnd: Math.max(1, layout.headerRow - 1),
      mergeToCol: Math.max(12, (layout.startCol || 2) + 10),
    });
  }

  const rows = (Array.isArray(mappedData) ? mappedData : []).filter(
    (row) =>
      row &&
      typeof row === 'object' &&
      Object.entries(row).some(
        ([k, v]) => !String(k).startsWith('__') && v != null && String(v).trim() !== ''
      )
  );

  const signatureCol = findFormPGJGujaratSignatureColumn(worksheet, layout);
  const tableColMin = Math.min(...templateCols.map(({ col }) => col), startCol);
  const tableColMax = Math.max(
    ...templateCols.map(({ col }) => col),
    signatureCol > 0 ? signatureCol : 0
  );

  // Unmerge vertical body merges so each employee keeps its own row.
  const merges = Array.isArray(worksheet.model?.merges) ? [...worksheet.model.merges] : [];
  merges.forEach((label) => {
    const parts = String(label || '').split(':');
    if (parts.length !== 2 || typeof worksheet.unMergeCells !== 'function') return;
    try {
      const tl = worksheet.getCell(parts[0]);
      const br = worksheet.getCell(parts[1]);
      const r1 = tl.fullAddress?.row ?? tl.row;
      const c1 = tl.fullAddress?.col ?? tl.col;
      const r2 = br.fullAddress?.row ?? br.row;
      const c2 = br.fullAddress?.col ?? br.col;
      if (r2 > r1 && r1 <= dataStartRow + rows.length + 5 && r2 >= dataStartRow && c2 >= tableColMin && c1 <= tableColMax) {
        worksheet.unMergeCells(label);
      }
    } catch (_) {
      /* ignore */
    }
  });

  const clearTo = Math.max(rows.length + 15, 20);
  for (let i = 0; i < clearTo; i += 1) {
    for (let col = tableColMin; col <= tableColMax; col += 1) {
      worksheet.getCell(dataStartRow + i, col).value = '';
    }
  }

  rows.forEach((row, idx) => {
    const excelRow = dataStartRow + idx;
    worksheet.getRow(excelRow).height = 18;
    templateCols.forEach(({ col, bucket }) => {
      let val = getFormPGJExportValueForBucket(row, bucket, idx);
      if (bucket === 'sno') val = String(idx + 1);
      if (String(bucket).startsWith('day:')) {
        const dayNum = Number(String(bucket).split(':')[1]);
        if (Number.isFinite(dayNum) && dayNum > monthDays) val = '';
      }
      if (bucket === 'dateOfEntry' && val) val = formatDate(val);
      if (bucket === 'paymentDate' && val) val = formatDate(val);
      if (bucket === 'overtimeHours' || bucket === 'overtimeEarnings') {
        if (!val || String(val).trim() === '') val = FORM_PGJ_OVERTIME_DEFAULT;
      }
      if (bucket === 'wageRate' || bucket === 'age' || bucket === 'sex' || bucket === 'signature') {
        val = '';
      }
      if (val == null || val === '' || /^enter\s+/i.test(String(val).trim())) {
        worksheet.getCell(excelRow, col).value = '';
        return;
      }
      const cell = worksheet.getCell(excelRow, col);
      if (
        typeof val === 'number' ||
        (typeof val === 'string' && /^-?\d+(\.\d+)?$/.test(String(val).trim()))
      ) {
        cell.value = Number(val);
      } else {
        cell.value = String(val);
      }
      cell.alignment = { ...(cell.alignment || {}), vertical: 'middle', wrapText: true };
    });
  });

  if (rows.length > 0) {
    ensureExcelJSDataRowsWithBorders(worksheet, {
      dataStartRow,
      dataRowCount: rows.length,
      colFrom: tableColMin,
      colTo: tableColMax,
      templateRow: dataStartRow,
      templateBodyRows: 1,
    });
    // Full box on Sr.No (B) through Signature/Thumb Impression (BI) for every data row.
    applyExcelJSFullBoxBordersToRange(worksheet, {
      rowFrom: dataStartRow,
      rowTo: dataStartRow + rows.length - 1,
      colFrom: tableColMin,
      colTo: tableColMax,
      borderStyle: 'thin',
    });
  }

  applyFormPGJGujaratMonthDayColumns(worksheet, layout, monthDays);
  applyFormPGJGujaratTableFontSize(worksheet, {
    headerRow: layout.headerRow,
    lastRow: dataStartRow + Math.max(rows.length, 1) - 1,
    colFrom: tableColMin,
    colTo: tableColMax,
    size: FORM_PGJ_TABLE_FONT_SIZE,
  });
  ensureFormPGJGujaratDateOfMonthGroupMerge(worksheet, layout, monthDays);

  const out = await workbook.xlsx.writeBuffer();
  return {
    blob: new Blob([out], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    }),
    fileName:
      formFileName ||
      parsedFormHeader?.title?.replace(/[^a-zA-Z0-9]/g, '_') ||
      'Form_P_GJ_Gujarat.xlsx',
  };
}
