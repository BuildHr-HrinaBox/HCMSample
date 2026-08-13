/** Rajasthan Form 12 — Register where opening/closing hours are ordinarily uniform [Rule 22 sub-rule 1]. */

import * as XLSX from 'xlsx';
import { flattenPayrollEarningColumns, readPayrollScalar } from '../../utils/payrollEarnings';

export const FORM_12_RJ_DISPLAY_TITLE = 'FORM 12';
export const FORM_12_RJ_DISPLAY_SUBTITLE =
  'Where opening and closing hours are ordinarily uniform';
export const FORM_12_RJ_DISPLAY_REFERENCE = 'Rule 22 - sub rule 1';

/** Leaf columns (19) matching Form_12_RJ.xlsx (B…T). */
export const FORM_12_RJ_TABLE_HEADERS = [
  'Name of person employed',
  'Whether young person or not',
  'Time at which employment commences',
  'Time at which employment ceases',
  'Rest interval',
  'Hours worked on_1',
  'Hours worked on_2',
  'Hours worked on_3',
  'Hours worked on_4',
  'Hours worked on_5',
  'Hours worked on_6',
  'Hours worked on_7',
  'Hours worked on_8',
  'Hours worked on_9',
  'Hours worked on_10',
  'Total hours worked during the month',
  '*Days on which overtime work is done and extent of overtime on each occasion',
  'Extent of overtime worked during the month',
  'Extent of overtime worked previously during the year',
];

export const FORM_12_RJ_YOUNG_DEFAULT = 'Young';
export const FORM_12_RJ_COMMENCES_DEFAULT = '9 AM';
export const FORM_12_RJ_CEASES_DEFAULT = '5 PM';
export const FORM_12_RJ_OT_DEFAULT = 'NIL';

const FORM_12_RJ_HOURS_COUNT = 10;

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

/** Explicit Form 12 RJ filename / title markers. */
export function looksLikeForm12RajasthanFileName(text) {
  const s = String(text || '').toLowerCase();
  return /form[\s._-]*12[\s._-]*rj|\bform_12_rj\b|\b12_rj\b/.test(s);
}

/** Table looks like RJ Form 12 (Hours worked on 1–10), not Form 11 Days-of-Month band. */
export function headersIndicateForm12RajasthanEmployment(tableHeaders) {
  const list = (Array.isArray(tableHeaders) ? tableHeaders : [])
    .map((h) => stripDedupeSuffix(h))
    .filter(Boolean);
  if (list.length < 8) return false;
  const joined = list.map((h) => norm(h)).join(' | ');
  if (/accident\s+book|insurance\s+no|cause\s+of\s+injury|periods?\s+of\s+work|group\s+letter/.test(joined)) {
    return false;
  }
  // Form 11 markers — do not treat as Form 12.
  if (/days\s+of\s+month|during\s+the\s+quarter|register\s+of\s+employment/.test(joined)) {
    return false;
  }
  const hasName = list.some((h) => /name\s+of\s+persons?\s+employ/.test(norm(h)));
  const hasYoung = list.some((h) => /young\s+person/.test(norm(h)));
  const hasCommences = list.some((h) =>
    /employment\s+commences|time\s+at\s+which\s+employment\s+commence/.test(norm(h))
  );
  const hasRest = list.some((h) => /rest\s+interval/.test(norm(h)));
  const hasHoursWorkedOn = list.some((h) => /hours\s+worked\s+on/.test(norm(h)));
  const hasPrevYear = list.some((h) =>
    /previously\s+during\s+the\s+year|overtime\s+worked\s+previously/.test(norm(h))
  );
  const numberedHours = list.filter((h) => {
    const s = norm(h);
    return /^hours\s+worked\s+on[_\s-]?\d+$/.test(s) || (/^\d{1,2}$/.test(s) && Number(s) >= 1 && Number(s) <= 10);
  }).length;
  return (
    hasName &&
    (hasYoung || hasCommences) &&
    (hasHoursWorkedOn || hasPrevYear || numberedHours >= 5 || (hasRest && list.length >= 15))
  );
}

/**
 * Rajasthan Shops Form 12 (uniform opening/closing hours register).
 * Must not match Form 11 RJ Register of Employment.
 */
export function isForm12RajasthanContext(
  formHeader,
  rowItem,
  fileName,
  sheetText = '',
  tableHeaders = []
) {
  const parts = buildContextBlob(formHeader, rowItem, fileName, sheetText, tableHeaders);

  if (/tamil\s*nadu|\b_tn\b|accident\s+book|employees.? state insurance/.test(parts)) {
    return false;
  }
  // Never treat Form 11 RJ filename as Form 12.
  if (/form[\s._-]*11[\s._-]*rj|\bform_11_rj\b|\b11_rj\b/.test(parts) && !looksLikeForm12RajasthanFileName(parts)) {
    return false;
  }
  // Form 14 RJ (record of hours of work) must not be treated as Form 12.
  if (
    /form[\s._-]*14[\s._-]*rj|\bform_14_rj\b|\b14_rj\b/.test(parts) ||
    /record\s+of\s+(?:the\s+)?hours\s+of\s+work/.test(parts)
  ) {
    return false;
  }
  if (/register\s+of\s+employment/.test(parts) && !/opening\s+and\s+closing\s+hours|form[\s._-]*12/.test(parts)) {
    return false;
  }

  if (looksLikeForm12RajasthanFileName(parts)) return true;
  if (/opening\s+and\s+closing\s+hours\s+are\s+ordinarily\s+uniform/.test(parts)) return true;
  if (
    /rajasthan/.test(parts) &&
    /\bform[\s._-]*12\b/.test(parts) &&
    (/rule\s*22/.test(parts) ||
      /opening\s+and\s+closing/.test(parts) ||
      headersIndicateForm12RajasthanEmployment(tableHeaders))
  ) {
    return true;
  }
  if (
    headersIndicateForm12RajasthanEmployment(tableHeaders) &&
    /rajasthan|\b_rj\b|shops\s+and\s+establishments/.test(parts)
  ) {
    return true;
  }
  return false;
}

export function classifyForm12RJHeader(header) {
  const s = norm(stripDedupeSuffix(header));
  if (!s) return '';
  if (/name\s+of\s+persons?\s+employ/.test(s) || (s.includes('name') && s.includes('employ') && !/young/.test(s))) {
    return 'name';
  }
  if (/young\s+person/.test(s)) return 'young';
  if (/employment\s+commences|time\s+at\s+which\s+employment\s+commence/.test(s)) return 'commences';
  if (/employment\s+ceases|time\s+at\s+which\s+employment\s+cease/.test(s)) return 'ceases';
  if (/rest\s+interval/.test(s)) return 'rest';
  const hoursMatch = s.match(/^hours\s+worked\s+on[_\s-]?(\d+)$/) || s.match(/^(\d{1,2})$/);
  if (hoursMatch) {
    const n = Number(hoursMatch[1]);
    if (n >= 1 && n <= FORM_12_RJ_HOURS_COUNT) return `hours${n}`;
  }
  if (/total\s+hours\s+worked/.test(s)) return 'totalHours';
  if (/days\s+on\s+which\s+overtime|overtime\s+work\s*is\s*done|extentof\s+overtime\s+on\s+each/.test(s)) {
    return 'otDays';
  }
  if (/extent\s+of\s+overtime.*month|overtime\s+worked\s+during\s+the\s+month/.test(s)) return 'otMonth';
  if (
    /previously\s+during\s+the\s+year|overtime\s+worked\s+previously|extent\s+of\s+overtime.*year/.test(s)
  ) {
    return 'otPrevYear';
  }
  return '';
}

const ROLE_TO_CANONICAL = {
  name: FORM_12_RJ_TABLE_HEADERS[0],
  young: FORM_12_RJ_TABLE_HEADERS[1],
  commences: FORM_12_RJ_TABLE_HEADERS[2],
  ceases: FORM_12_RJ_TABLE_HEADERS[3],
  rest: FORM_12_RJ_TABLE_HEADERS[4],
  hours1: FORM_12_RJ_TABLE_HEADERS[5],
  hours2: FORM_12_RJ_TABLE_HEADERS[6],
  hours3: FORM_12_RJ_TABLE_HEADERS[7],
  hours4: FORM_12_RJ_TABLE_HEADERS[8],
  hours5: FORM_12_RJ_TABLE_HEADERS[9],
  hours6: FORM_12_RJ_TABLE_HEADERS[10],
  hours7: FORM_12_RJ_TABLE_HEADERS[11],
  hours8: FORM_12_RJ_TABLE_HEADERS[12],
  hours9: FORM_12_RJ_TABLE_HEADERS[13],
  hours10: FORM_12_RJ_TABLE_HEADERS[14],
  totalHours: FORM_12_RJ_TABLE_HEADERS[15],
  otDays: FORM_12_RJ_TABLE_HEADERS[16],
  otMonth: FORM_12_RJ_TABLE_HEADERS[17],
  otPrevYear: FORM_12_RJ_TABLE_HEADERS[18],
};

/** Always return the canonical 19 leaf headers (stable order + spelling). */
export function resolveForm12RajasthanTableHeaders(_tableHeaders) {
  return [...FORM_12_RJ_TABLE_HEADERS];
}

export function buildForm12RajasthanExportColMap(tableStartCol = 1, headers = FORM_12_RJ_TABLE_HEADERS) {
  const start =
    tableStartCol != null && Number.isFinite(Number(tableStartCol)) && Number(tableStartCol) >= 0
      ? Number(tableStartCol)
      : 1;
  const hdrs = resolveForm12RajasthanTableHeaders(headers);
  const headerToCol = new Map();
  hdrs.forEach((header, j) => {
    if (header) headerToCol.set(header, start + j);
  });
  return { headers: hdrs, headerToCol, tableStartCol: start };
}

export function exportForm12RajasthanRowValues(row, headers = FORM_12_RJ_TABLE_HEADERS) {
  const hdrs = resolveForm12RajasthanTableHeaders(headers);
  return hdrs.map((h) => getForm12RajasthanRowValueForHeader(row, h));
}

export function prepareForm12RajasthanExport(workbook, rows, sourceHeaders = [], hints = {}) {
  const repaired = workbook ? repairForm12RajasthanTableHeadersFromWorkbook(workbook, hints) : null;
  const headers = resolveForm12RajasthanTableHeaders(
    repaired?.headers?.length ? repaired.headers : sourceHeaders
  );
  const mappedRows = (Array.isArray(rows) ? rows : []).map((row) =>
    repairForm12RajasthanExportRow(row, headers)
  );
  applyForm12RajasthanAutofillDefaultsToRows(mappedRows, headers);
  const tableStartCol =
    repaired?.tableStartCol != null && Number.isFinite(Number(repaired.tableStartCol))
      ? Number(repaired.tableStartCol)
      : hints.tableStartCol != null && Number.isFinite(Number(hints.tableStartCol))
        ? Number(hints.tableStartCol)
        : 1;
  const { headerToCol } = buildForm12RajasthanExportColMap(tableStartCol, headers);
  return {
    headers,
    rows: mappedRows,
    headerToCol,
    tableStartCol,
    headerRowIndex:
      repaired?.headerRowIndex != null && Number.isFinite(Number(repaired.headerRowIndex))
        ? Number(repaired.headerRowIndex)
        : hints.headerRowIndex,
    dataStartIndex:
      repaired?.dataStartIndex != null && Number.isFinite(Number(repaired.dataStartIndex))
        ? Number(repaired.dataStartIndex)
        : hints.dataStartIndex,
  };
}

export function isForm12RajasthanEmployeeIdLikeValue(value) {
  const s = String(value ?? '').trim();
  if (!s) return false;
  return /^\d{1,6}$/.test(s);
}

export function looksLikeForm12RajasthanPersonName(value) {
  const s = String(value ?? '').trim();
  if (!s || s.length < 2) return false;
  if (/^nil$/i.test(s)) return false;
  if (/^young$/i.test(s)) return false;
  if (isForm12RajasthanEmployeeIdLikeValue(s)) return false;
  if (/^\d{1,2}([:.]?\d{2})?\s*(a\.?m\.?|p\.?m\.?)?$/i.test(s)) return false;
  if (/^\d+(\.\d+)?$/.test(s)) return false;
  return /[a-z]/i.test(s);
}

export function looksLikeForm12RajasthanTimeValue(value) {
  const s = String(value ?? '').trim();
  if (!s) return false;
  if (/^nil$/i.test(s)) return false;
  if (looksLikeForm12RajasthanPersonName(s) && !/\d/.test(s)) return false;
  return (
    /^\d{1,2}([:.]?\d{2})?\s*(a\.?m\.?|p\.?m\.?)?$/i.test(s) ||
    /^\d{1,2}\s*(a\.?m\.?|p\.?m\.?)$/i.test(s)
  );
}

function collectForm12RajasthanRoleCandidateValues(row, role) {
  if (!row || typeof row !== 'object' || !role) return [];
  const out = [];
  const seen = new Set();
  const push = (raw) => {
    const s = String(raw ?? '').trim();
    if (!s || s === 'false' || seen.has(s)) return;
    seen.add(s);
    out.push(s);
  };
  Object.entries(row).forEach(([k, v]) => {
    if (String(k).startsWith('__')) return;
    if (classifyForm12RJHeader(k) === role) push(v);
  });
  return out;
}

export function pickForm12RajasthanRoleValue(row, role) {
  if (!row || !role) return '';
  const roleVals = collectForm12RajasthanRoleCandidateValues(row, role);

  if (role === 'name') {
    const fromName = roleVals.find((v) => looksLikeForm12RajasthanPersonName(v));
    if (fromName) return fromName;
    const misplaced = collectForm12RajasthanRoleCandidateValues(row, 'ceases')
      .concat(collectForm12RajasthanRoleCandidateValues(row, 'commences'))
      .concat(collectForm12RajasthanRoleCandidateValues(row, 'young'))
      .find((v) => looksLikeForm12RajasthanPersonName(v));
    if (misplaced) return misplaced;
    for (const [k, v] of Object.entries(row)) {
      if (String(k).startsWith('__')) continue;
      if (/employee\s*id|^empid$/i.test(norm(k))) continue;
      if (/^employeeid$/i.test(norm(k).replace(/\s+/g, ''))) continue;
      if (looksLikeForm12RajasthanPersonName(v)) return String(v).trim();
    }
    return '';
  }

  if (role === 'ceases' || role === 'commences') {
    const timed = roleVals.find((v) => looksLikeForm12RajasthanTimeValue(v));
    if (timed) return timed;
    const nonName = roleVals.find(
      (v) =>
        v &&
        !/^nil$/i.test(v) &&
        !looksLikeForm12RajasthanPersonName(v) &&
        !isForm12RajasthanEmployeeIdLikeValue(v)
    );
    if (nonName) return nonName;
    if (role === 'commences') return FORM_12_RJ_COMMENCES_DEFAULT;
    if (role === 'ceases') return FORM_12_RJ_CEASES_DEFAULT;
    return '';
  }

  if (role === 'young') {
    const good = roleVals.find(
      (v) =>
        v &&
        !/^nil$/i.test(v) &&
        !looksLikeForm12RajasthanPersonName(v) &&
        !isForm12RajasthanEmployeeIdLikeValue(v)
    );
    if (good) return good;
    return FORM_12_RJ_YOUNG_DEFAULT;
  }

  if (role === 'otDays' || role === 'otMonth' || role === 'otPrevYear') {
    const ot = roleVals.find((v) => v && !looksLikeForm12RajasthanPersonName(v));
    return ot || FORM_12_RJ_OT_DEFAULT;
  }

  if (role === 'totalHours') {
    const hours = roleVals.find((v) => v && !isForm12RajasthanJunkTotalHoursValue(v));
    return hours ? sanitizeForm12RajasthanTotalHoursValue(hours) : '';
  }

  if (role === 'rest' || /^hours\d+$/.test(role)) {
    const rest = roleVals.find((v) => v && !/^nil$/i.test(v) && !looksLikeForm12RajasthanPersonName(v));
    return rest || '';
  }

  return roleVals[0] || '';
}

export function repairForm12RajasthanExportRow(row, targetHeaders = FORM_12_RJ_TABLE_HEADERS) {
  const tgt = resolveForm12RajasthanTableHeaders(targetHeaders);
  const out = {};
  tgt.forEach((header) => {
    const role = classifyForm12RJHeader(header);
    const picked = role ? pickForm12RajasthanRoleValue(row, role) : '';
    out[header] = sanitizeForm12RajasthanCellValue(picked, header);
  });
  if (row && typeof row === 'object') {
    Object.keys(row).forEach((k) => {
      if (String(k).startsWith('__')) out[k] = row[k];
    });
  }
  const nameHdr = tgt.find((h) => classifyForm12RJHeader(h) === 'name');
  if (nameHdr && isForm12RajasthanEmployeeIdLikeValue(out[nameHdr])) {
    const recovered = pickForm12RajasthanRoleValue(row, 'name');
    if (looksLikeForm12RajasthanPersonName(recovered)) {
      out[nameHdr] = sanitizeForm12RajasthanCellValue(recovered, nameHdr);
    } else {
      out[nameHdr] = '';
    }
  }
  return out;
}

export function sanitizeForm12RajasthanCellValue(value, header = '') {
  let s = String(value ?? '').trim();
  if (!s) return '';
  if (/<[^>]+>/.test(s) || /please\s+follow\s+the\s+steps/i.test(s)) return '';
  s = s.replace(/<[^>]*>/g, '').trim();
  if (header && isForm12RajasthanTotalHoursHeader(header)) {
    return sanitizeForm12RajasthanTotalHoursValue(s);
  }
  return s;
}

export function remapForm12RajasthanRowsToHeaders(rows, sourceHeaders, targetHeaders) {
  const tgt = resolveForm12RajasthanTableHeaders(targetHeaders?.length ? targetHeaders : sourceHeaders);
  if (!Array.isArray(rows)) return [];
  return rows.map((row) => {
    if (!row || typeof row !== 'object') {
      const blank = {};
      tgt.forEach((h) => {
        blank[h] = '';
      });
      return blank;
    }
    return repairForm12RajasthanExportRow(row, tgt);
  });
}

export function isForm12RajasthanNameHeader(header) {
  return classifyForm12RJHeader(header) === 'name';
}

export function isForm12RajasthanYoungPersonHeader(header) {
  return classifyForm12RJHeader(header) === 'young';
}

export function isForm12RajasthanCommencesHeader(header) {
  return classifyForm12RJHeader(header) === 'commences';
}

export function isForm12RajasthanCeasesHeader(header) {
  return classifyForm12RJHeader(header) === 'ceases';
}

export function isForm12RajasthanTotalHoursHeader(header) {
  return classifyForm12RJHeader(header) === 'totalHours';
}

export function isForm12RajasthanRestHeader(header) {
  return classifyForm12RJHeader(header) === 'rest';
}

export function isForm12RajasthanHoursWorkedOnHeader(header) {
  return /^hours\d+$/.test(classifyForm12RJHeader(header));
}

export function isForm12RajasthanOtHeader(header) {
  const role = classifyForm12RJHeader(header);
  return role === 'otDays' || role === 'otMonth' || role === 'otPrevYear';
}

/** Rest interval + daily hours columns stay blank during People autofill. */
export function isForm12RajasthanSkipAutofillHeader(header) {
  const role = classifyForm12RJHeader(header);
  return role === 'rest' || /^hours\d+$/.test(role);
}

export function resolveForm12RajasthanYoungPersonValue(_ageYears) {
  return FORM_12_RJ_YOUNG_DEFAULT;
}

export function isForm12RajasthanJunkTotalHoursValue(value) {
  const s = String(value ?? '').trim();
  if (!s) return false;
  if (/month|year|week|day\s*\(|experience|tenure/i.test(s) && !/^\d+(\.\d+)?$/.test(s)) {
    return true;
  }
  if (!/^\d+(\.\d+)?$/.test(s.replace(/,/g, ''))) return true;
  return false;
}

export function sanitizeForm12RajasthanTotalHoursValue(value) {
  const s = String(value ?? '').trim();
  if (!s || isForm12RajasthanJunkTotalHoursValue(s)) return '';
  const n = Number(s.replace(/,/g, ''));
  if (!Number.isFinite(n) || n < 0) return '';
  return String(Math.round(n * 100) / 100);
}

export function computeForm12RajasthanTotalHours(paidDays) {
  if (paidDays === '' || paidDays == null) return '';
  const raw = String(paidDays).trim();
  if (!raw || /month|year|week|experience|tenure/i.test(raw)) return '';
  const n = Number(raw.replace(/,/g, ''));
  if (!Number.isFinite(n) || n < 0) return '';
  return String(Math.round(n * 8 * 100) / 100);
}

export function readForm12RajasthanPaidDays(payrollRow) {
  if (!payrollRow || typeof payrollRow !== 'object' || payrollRow.fetch_error) return '';
  const flat = flattenPayrollEarningColumns(payrollRow);
  const keys = [
    'Paid_days',
    'paid_days',
    'Paid Days',
    'paidDays',
    'PaidDays',
    'days_worked',
    'Days Worked',
    'daysWorked',
    'no_of_days_worked',
    'effective_paid_days',
  ];
  const patterns = [
    /^paid_days$/,
    /^paiddays$/,
    /^days_worked$/,
    /^daysworked$/,
    /^no_of_days_worked$/,
    /^effective_paid_days$/,
  ];
  const fromFlat = readPayrollScalar({ ...flat, ...payrollRow }, keys, patterns);
  if (fromFlat !== '' && fromFlat != null) {
    const raw = String(fromFlat).trim();
    if (/month|year|week|experience|tenure/i.test(raw)) return '';
    const num = Number(raw.replace(/,/g, ''));
    if (Number.isFinite(num) && num >= 0) return String(num);
  }
  return '';
}

export function getForm12RajasthanRowValueForHeader(row, header) {
  if (!row || typeof row !== 'object') return '';
  const role = classifyForm12RJHeader(header);
  if (role) {
    const picked = pickForm12RajasthanRoleValue(row, role);
    if (picked != null && String(picked).trim() !== '') return String(picked).trim();
  }
  const direct = row[header];
  if (direct != null && String(direct).trim() !== '' && String(direct) !== 'false') {
    if (role === 'name' && isForm12RajasthanEmployeeIdLikeValue(direct)) return '';
    return String(direct);
  }
  return '';
}

function setForm12RajasthanRoleValue(row, headers, role, value) {
  if (!row || !role) return false;
  const hdrs = Array.isArray(headers) ? headers : [];
  let wrote = false;
  hdrs.forEach((h) => {
    if (classifyForm12RJHeader(h) === role) {
      row[h] = value;
      wrote = true;
    }
  });
  if (!wrote && ROLE_TO_CANONICAL[role]) {
    row[ROLE_TO_CANONICAL[role]] = value;
    wrote = true;
  }
  return wrote;
}

export function applyForm12RajasthanAutofillDefaultsToRows(mappedData, tableHeaders, options = {}) {
  if (!Array.isArray(mappedData) || mappedData.length === 0) return 0;
  const headers = resolveForm12RajasthanTableHeaders(tableHeaders);
  const sanitizeValue =
    typeof options.sanitizeValue === 'function'
      ? options.sanitizeValue
      : (v) => String(v ?? '').trim();

  let hits = 0;
  mappedData.forEach((row) => {
    if (!row || typeof row !== 'object') return;
    setForm12RajasthanRoleValue(row, headers, 'young', sanitizeValue(FORM_12_RJ_YOUNG_DEFAULT));
    setForm12RajasthanRoleValue(row, headers, 'commences', sanitizeValue(FORM_12_RJ_COMMENCES_DEFAULT));
    setForm12RajasthanRoleValue(row, headers, 'ceases', sanitizeValue(FORM_12_RJ_CEASES_DEFAULT));
    setForm12RajasthanRoleValue(row, headers, 'otDays', sanitizeValue(FORM_12_RJ_OT_DEFAULT));
    setForm12RajasthanRoleValue(row, headers, 'otMonth', sanitizeValue(FORM_12_RJ_OT_DEFAULT));
    setForm12RajasthanRoleValue(row, headers, 'otPrevYear', sanitizeValue(FORM_12_RJ_OT_DEFAULT));

    // Hours worked on 1–10 + Rest interval stay blank (no attendance fetch).
    setForm12RajasthanRoleValue(row, headers, 'rest', '');
    for (let d = 1; d <= FORM_12_RJ_HOURS_COUNT; d += 1) {
      setForm12RajasthanRoleValue(row, headers, `hours${d}`, '');
    }

    const hoursHeader = headers.find((h) => isForm12RajasthanTotalHoursHeader(h));
    if (hoursHeader) {
      const cur = String(row[hoursHeader] ?? '').trim();
      if (!cur || isForm12RajasthanJunkTotalHoursValue(cur)) {
        row[hoursHeader] = '';
      }
    }
    hits += 1;
  });
  return hits;
}

/** Clear Hours worked on 1–10 (and Rest interval) after any attendance enrich pass. */
export function clearForm12RajasthanAttendanceDayColumns(mappedData, tableHeaders) {
  if (!Array.isArray(mappedData) || mappedData.length === 0) return 0;
  const headers = resolveForm12RajasthanTableHeaders(tableHeaders);
  let hits = 0;
  mappedData.forEach((row) => {
    if (!row || typeof row !== 'object') return;
    setForm12RajasthanRoleValue(row, headers, 'rest', '');
    for (let d = 1; d <= FORM_12_RJ_HOURS_COUNT; d += 1) {
      setForm12RajasthanRoleValue(row, headers, `hours${d}`, '');
    }
    hits += 1;
  });
  return hits;
}

export function applyForm12RajasthanPaidDaysToRows(
  mappedData,
  employeesForMapping,
  tableHeaders,
  options = {}
) {
  if (!Array.isArray(mappedData) || mappedData.length === 0) return 0;
  const headers = Array.isArray(tableHeaders) ? tableHeaders : [];
  const hoursHeader = headers.find((h) => isForm12RajasthanTotalHoursHeader(h));
  if (!hoursHeader) return 0;

  const resolvePayrollRow =
    typeof options.resolvePayrollRow === 'function' ? options.resolvePayrollRow : () => null;
  const unwrapEmp =
    typeof options.unwrapEmp === 'function'
      ? options.unwrapEmp
      : (empItem) => empItem && (empItem.Employee || empItem.employee || empItem);
  const sanitizeValue =
    typeof options.sanitizeValue === 'function'
      ? options.sanitizeValue
      : (v) => String(v ?? '').trim();
  const overwrite = options.overwrite !== false;

  let hits = 0;
  mappedData.forEach((row, idx) => {
    if (!row || typeof row !== 'object') return;
    const emp = unwrapEmp(employeesForMapping?.[idx]);
    const payrollRow = resolvePayrollRow(emp, row, idx);
    const paidDays = readForm12RajasthanPaidDays(payrollRow);
    const hours = computeForm12RajasthanTotalHours(paidDays);
    const cur = String(row[hoursHeader] ?? '').trim();
    if (hours === '') {
      if (cur && (overwrite || isForm12RajasthanJunkTotalHoursValue(cur))) {
        row[hoursHeader] = '';
        hits += 1;
      }
      return;
    }
    if (!overwrite && cur && !isForm12RajasthanJunkTotalHoursValue(cur)) return;
    row[hoursHeader] = sanitizeValue(hours);
    hits += 1;
  });
  return hits;
}

export function enrichForm12RajasthanDisplayHeader(
  formHeader,
  fileName = '',
  item = null,
  tableHeaders = [],
  sheetText = ''
) {
  if (!isForm12RajasthanContext(formHeader, item, fileName, sheetText, tableHeaders)) {
    return formHeader;
  }
  const base = formHeader && typeof formHeader === 'object' ? { ...formHeader } : {};
  base.title = FORM_12_RJ_DISPLAY_TITLE;
  base.subtitle = FORM_12_RJ_DISPLAY_SUBTITLE;
  base.reference = FORM_12_RJ_DISPLAY_REFERENCE;
  base.form12RJEmploymentLayout = true;
  return base;
}

export function isForm12RajasthanEmploymentLayoutFormHeader(formHeader) {
  return !!formHeader?.form12RJEmploymentLayout;
}

/**
 * Locate header band + data start from Form 12 workbook
 * (Hours worked on leaf row 1…10, then column-number row 1…19).
 */
export function repairForm12RajasthanTableHeadersFromWorkbook(workbook, hints = {}) {
  if (!workbook || !workbook.SheetNames || !workbook.SheetNames.length) return null;
  const sheetName = hints.sheetName || workbook.SheetNames[0];
  const ws = workbook.Sheets[sheetName];
  if (!ws) return null;
  const merges = ws['!merges'] || [];
  const ref = ws['!ref'];
  if (!ref) return null;
  const range = XLSX.utils.decode_range(ref);

  const rawCell = (r, c) => {
    const cell = ws[XLSX.utils.encode_cell({ r, c })];
    return cell && cell.v != null ? String(cell.v).trim() : '';
  };
  const mergedAware = (r, c) => {
    const direct = rawCell(r, c);
    if (direct) return direct;
    for (let i = 0; i < merges.length; i += 1) {
      const m = merges[i];
      if (!m?.s || !m?.e) continue;
      if (r >= m.s.r && r <= m.e.r && c >= m.s.c && c <= m.e.c) {
        return rawCell(m.s.r, m.s.c);
      }
    }
    return '';
  };

  let headerRow = -1;
  let startCol = range.s.c;
  for (let r = range.s.r; r <= Math.min(range.e.r, range.s.r + 30); r += 1) {
    for (let c = range.s.c; c <= range.e.c; c += 1) {
      const t = norm(mergedAware(r, c));
      if (/name\s+of\s+persons?\s+employ/.test(t)) {
        headerRow = r;
        startCol = c;
        break;
      }
    }
    if (headerRow >= 0) break;
  }
  if (headerRow < 0) return null;

  // Confirm Form 12 shape (Hours worked on / previously during the year), not Form 11.
  let probe = '';
  for (let c = startCol; c <= Math.min(range.e.c, startCol + 20); c += 1) {
    probe += ` ${mergedAware(headerRow, c)} ${mergedAware(headerRow + 1, c)} ${mergedAware(headerRow + 2, c)}`;
  }
  const probeNorm = norm(probe);
  if (
    /days\s+of\s+month/.test(probeNorm) ||
    (/during\s+the\s+quarter/.test(probeNorm) && !/hours\s+worked\s+on/.test(probeNorm))
  ) {
    return null;
  }
  if (
    !/hours\s+worked\s+on/.test(probeNorm) &&
    !/previously\s+during\s+the\s+year|opening\s+and\s+closing/.test(probeNorm)
  ) {
    // Still accept when filename hints Form 12.
    if (!looksLikeForm12RajasthanFileName(hints.fileName || hints.formFileName || sheetName)) {
      return null;
    }
  }

  // Leaf day numbers sit on headerRow+2 under "Hours worked on".
  let leafRow = headerRow + 1;
  const dayProbe = [];
  for (let i = 0; i < FORM_12_RJ_HOURS_COUNT; i += 1) {
    dayProbe.push(String(mergedAware(headerRow + 2, startCol + 5 + i) || '').trim());
  }
  const dayHits = dayProbe.filter((t, i) => t === String(i + 1)).length;
  if (dayHits >= 6) leafRow = headerRow + 2;

  const headers = resolveForm12RajasthanTableHeaders();

  // Column-number row (1…19) sits under the leaf band.
  let dataStart = leafRow + 1;
  let numHits = 0;
  for (let i = 0; i < 19; i += 1) {
    const t = String(mergedAware(dataStart, startCol + i) || '').trim();
    if (t === String(i + 1)) numHits += 1;
  }
  if (numHits >= 12) dataStart += 1;

  return {
    headers,
    headerRowIndex: headerRow,
    dataStartIndex: dataStart,
    tableStartCol: startCol,
    stackedHeaderBandEnd: leafRow,
  };
}
