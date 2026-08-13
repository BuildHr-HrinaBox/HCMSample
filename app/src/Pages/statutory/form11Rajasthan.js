/** Rajasthan Form 11 — Register of Employment [Rule 22 sub-rule 1]. */

import ExcelJS from 'exceljs';
import * as XLSX from 'xlsx';
import { flattenPayrollEarningColumns, readPayrollScalar } from '../../utils/payrollEarnings';
import { ensureExcelJSDataRowsWithBorders } from '../../utils/excelTableBorders';

export const FORM_11_RJ_DISPLAY_TITLE = 'FORM 11';
export const FORM_11_RJ_DISPLAY_SUBTITLE = 'Register of Employment';
export const FORM_11_RJ_DISPLAY_REFERENCE = 'Rule 22 - sub rule 1';

/** Leaf columns (12) matching Form_11_RJ.xlsx after resolving the Days-of-Month band. */
export const FORM_11_RJ_TABLE_HEADERS = [
  'Name of Persons Employed',
  'Whether young person or not',
  'Time at which employment commences',
  'Time at which employment ceases',
  '1',
  '2 Rest interval',
  '3',
  'Total hours worked during the month',
  '*Days on which overtime work is done and extent of such overtime on each day',
  'Extent of overtime worked during the month',
  'Extent of overtime worked during the quarter',
  'Extent of overtime worked during the year',
];

export const FORM_11_RJ_YOUNG_DEFAULT = 'Young';
export const FORM_11_RJ_COMMENCES_DEFAULT = '9AM';
export const FORM_11_RJ_CEASES_DEFAULT = '5PM';
export const FORM_11_RJ_REST_BEFORE_DEFAULT = '9AM';
export const FORM_11_RJ_REST_AFTER_DEFAULT = '5PM';
export const FORM_11_RJ_OT_DEFAULT = 'NIL';

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

/** Table looks like RJ Register of Employment (not ESIC Accident Book / AP period of work / Form 12). */
export function headersIndicateForm11RajasthanEmployment(tableHeaders) {
  const list = (Array.isArray(tableHeaders) ? tableHeaders : [])
    .map((h) => stripDedupeSuffix(h))
    .filter(Boolean);
  if (list.length < 4) return false;
  const joined = list.map((h) => norm(h)).join(' | ');
  if (/accident\s+book|insurance\s+no|cause\s+of\s+injury|periods?\s+of\s+work|group\s+letter/.test(joined)) {
    return false;
  }
  // Form 12 RJ: Hours worked on 1–10 / previously during the year (not Form 11 Days-of-Month).
  if (
    /hours\s+worked\s+on/.test(joined) ||
    /previously\s+during\s+the\s+year/.test(joined) ||
    (list.length >= 15 && /rest\s+interval/.test(joined) && !/days\s+of\s+month|during\s+the\s+quarter/.test(joined))
  ) {
    return false;
  }
  const hasName = list.some((h) => /name\s+of\s+persons?\s+employ/.test(norm(h)));
  const hasYoung = list.some((h) => /young\s+person/.test(norm(h)));
  const hasCommences = list.some((h) => /employment\s+commences|time\s+at\s+which\s+employment\s+commence/.test(norm(h)));
  const hasRest = list.some((h) => /rest\s+interval/.test(norm(h)));
  const hasTotalHours = list.some((h) => /total\s+hours\s+worked/.test(norm(h)));
  const hasOtExtent = list.some((h) => /extent\s+of\s+overtime/.test(norm(h)));
  return hasName && (hasYoung || hasCommences) && (hasRest || hasTotalHours || hasOtExtent || hasCommences);
}

/**
 * Rajasthan Shops Form 11 (Register of Employment).
 * Must not match TN ESIC Form 11 Accident Book or AP Form 11 Period of Work.
 */
export function isForm11RajasthanContext(
  formHeader,
  rowItem,
  fileName,
  sheetText = '',
  tableHeaders = []
) {
  const parts = buildContextBlob(formHeader, rowItem, fileName, sheetText, tableHeaders);

  if (/tamil\s*nadu|\b_tn\b|accident\s+book|employees.? state insurance|regulation\s*66/.test(parts)) {
    return false;
  }
  if (/andhra|\b_ap\b|period\s+of\s+work|rule\s*79/.test(parts) && !/rajasthan|\b_rj\b/.test(parts)) {
    return false;
  }
  // Form 12 RJ (uniform opening/closing hours) must not be treated as Form 11.
  if (
    /form[\s._-]*12[\s._-]*rj|\bform_12_rj\b|\b12_rj\b/.test(parts) ||
    /opening\s+and\s+closing\s+hours\s+are\s+ordinarily\s+uniform/.test(parts)
  ) {
    return false;
  }
  // Form 14 RJ (record of hours of work) must not be treated as Form 11.
  if (
    /form[\s._-]*14[\s._-]*rj|\bform_14_rj\b|\b14_rj\b/.test(parts) ||
    /record\s+of\s+(?:the\s+)?hours\s+of\s+work/.test(parts)
  ) {
    return false;
  }
  {
    const hdrJoined = (Array.isArray(tableHeaders) ? tableHeaders : [])
      .map((h) => norm(stripDedupeSuffix(h)))
      .join(' | ');
    if (
      /hours\s+worked\s+on/.test(hdrJoined) ||
      /previously\s+during\s+the\s+year/.test(hdrJoined)
    ) {
      return false;
    }
  }

  if (/form[\s._-]*11[\s._-]*rj|\bform_11_rj\b|\b11_rj\b/.test(parts)) return true;
  if (
    /rajasthan/.test(parts) &&
    /\bform[\s._-]*11\b/.test(parts) &&
    (/register\s+of\s+employment|rule\s*22/.test(parts) ||
      headersIndicateForm11RajasthanEmployment(tableHeaders))
  ) {
    return true;
  }
  if (
    headersIndicateForm11RajasthanEmployment(tableHeaders) &&
    /rajasthan|\b_rj\b|shops\s+and\s+establishments/.test(parts)
  ) {
    return true;
  }
  return false;
}

export function classifyForm11RJHeader(header) {
  const s = norm(stripDedupeSuffix(header));
  if (!s) return '';
  if (/name\s+of\s+persons?\s+employ/.test(s) || (s.includes('name') && s.includes('employ') && !/young/.test(s))) {
    return 'name';
  }
  if (/young\s+person/.test(s)) return 'young';
  if (/employment\s+commences|time\s+at\s+which\s+employment\s+commence/.test(s)) return 'commences';
  if (/employment\s+ceases|time\s+at\s+which\s+employment\s+cease/.test(s)) return 'ceases';
  if (/^1$/.test(s) || /^day\s*1\b/.test(s)) return 'day1';
  if (/rest\s+interval/.test(s) || /^2(\s|$)/.test(s)) return 'rest';
  if (/^3$/.test(s) || /^day\s*3\b/.test(s)) return 'day3';
  if (/total\s+hours\s+worked/.test(s)) return 'totalHours';
  if (/days\s+on\s+which\s+overtime|overtime\s+work\s+is\s+done/.test(s)) return 'otDays';
  if (/extent\s+of\s+overtime.*month|overtime\s+worked\s+during\s+the\s+month/.test(s)) return 'otMonth';
  if (/extent\s+of\s+overtime.*quarter|overtime\s+worked\s+during\s+the\s+quarter/.test(s)) {
    return 'otQuarter';
  }
  if (/extent\s+of\s+overtime.*year|overtime\s+worked\s+during\s+the\s+year/.test(s)) return 'otYear';
  return '';
}

const ROLE_TO_CANONICAL = {
  name: FORM_11_RJ_TABLE_HEADERS[0],
  young: FORM_11_RJ_TABLE_HEADERS[1],
  commences: FORM_11_RJ_TABLE_HEADERS[2],
  ceases: FORM_11_RJ_TABLE_HEADERS[3],
  day1: FORM_11_RJ_TABLE_HEADERS[4],
  rest: FORM_11_RJ_TABLE_HEADERS[5],
  day3: FORM_11_RJ_TABLE_HEADERS[6],
  totalHours: FORM_11_RJ_TABLE_HEADERS[7],
  otDays: FORM_11_RJ_TABLE_HEADERS[8],
  otMonth: FORM_11_RJ_TABLE_HEADERS[9],
  otQuarter: FORM_11_RJ_TABLE_HEADERS[10],
  otYear: FORM_11_RJ_TABLE_HEADERS[11],
};

/**
 * Always return the canonical 12 leaf headers (stable order + spelling).
 * Template typo "Employeed" and Object.keys() reordering of bare "1"/"3" keys
 * otherwise shift Name into the wrong Excel column on export.
 */
export function resolveForm11RajasthanTableHeaders(_tableHeaders) {
  return [...FORM_11_RJ_TABLE_HEADERS];
}

/** Build SheetJS header→column map: one physical column per canonical leaf (B…). */
export function buildForm11RajasthanExportColMap(tableStartCol = 1, headers = FORM_11_RJ_TABLE_HEADERS) {
  const start =
    tableStartCol != null && Number.isFinite(Number(tableStartCol)) && Number(tableStartCol) >= 0
      ? Number(tableStartCol)
      : 1;
  const hdrs = resolveForm11RajasthanTableHeaders(headers);
  const headerToCol = new Map();
  hdrs.forEach((header, j) => {
    if (header) headerToCol.set(header, start + j);
  });
  return { headers: hdrs, headerToCol, tableStartCol: start };
}

/** Values for one export row in canonical header order (role-aware read). */
export function exportForm11RajasthanRowValues(row, headers = FORM_11_RJ_TABLE_HEADERS) {
  const hdrs = resolveForm11RajasthanTableHeaders(headers);
  return hdrs.map((h) => getForm11RajasthanRowValueForHeader(row, h));
}

/**
 * Prepare rows/headers/anchors for Form 11 RJ Excel export.
 * Prefer workbook repair for dataStart / tableStartCol; always remap onto canonical headers.
 */
export function prepareForm11RajasthanExport(workbook, rows, sourceHeaders = [], hints = {}) {
  const repaired = workbook ? repairForm11RajasthanTableHeadersFromWorkbook(workbook, hints) : null;
  const headers = resolveForm11RajasthanTableHeaders(
    repaired?.headers?.length ? repaired.headers : sourceHeaders
  );
  const mappedRows = (Array.isArray(rows) ? rows : []).map((row) =>
    repairForm11RajasthanExportRow(row, headers)
  );
  // Re-apply modal defaults so Young / times / OT NIL stay correct after shift repair.
  applyForm11RajasthanAutofillDefaultsToRows(mappedRows, headers);
  // Name of Persons is always column B (1) on Form_11_RJ.xlsx — never allow 0 (column A),
  // which shifts Name→A / Young→B / name-under-Ceases→D and matches the broken download.
  const rawStart =
    repaired?.tableStartCol != null && Number.isFinite(Number(repaired.tableStartCol))
      ? Number(repaired.tableStartCol)
      : hints.tableStartCol != null && Number.isFinite(Number(hints.tableStartCol))
        ? Number(hints.tableStartCol)
        : 1;
  const tableStartCol = rawStart >= 1 ? rawStart : 1;
  const { headerToCol } = buildForm11RajasthanExportColMap(tableStartCol, headers);
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

function readFirstNonEmpty(row, headers, className) {
  if (!row || typeof row !== 'object') return '';
  for (let i = 0; i < headers.length; i += 1) {
    const h = headers[i];
    if (classifyForm11RJHeader(h) !== className) continue;
    const v = row[h];
    if (v != null && String(v).trim() !== '') return v;
  }
  for (const [k, v] of Object.entries(row)) {
    if (String(k).startsWith('__')) continue;
    if (classifyForm11RJHeader(k) !== className) continue;
    if (v != null && String(v).trim() !== '') return v;
  }
  return '';
}

/** Pure employee-id / serial (e.g. "4") — not a person name. */
export function isForm11RajasthanEmployeeIdLikeValue(value) {
  const s = String(value ?? '').trim();
  if (!s) return false;
  return /^\d{1,6}$/.test(s);
}

/** Person name: has letters, not NIL / Young / a clock time. */
export function looksLikeForm11RajasthanPersonName(value) {
  const s = String(value ?? '').trim();
  if (!s || s.length < 2) return false;
  if (/^nil$/i.test(s)) return false;
  if (/^young$/i.test(s)) return false;
  if (isForm11RajasthanEmployeeIdLikeValue(s)) return false;
  if (/^\d{1,2}([:.]?\d{2})?\s*(a\.?m\.?|p\.?m\.?)?$/i.test(s)) return false;
  if (/^\d+(\.\d+)?$/.test(s)) return false;
  return /[a-z]/i.test(s);
}

/** Employment commence/cease / rest time text. */
export function looksLikeForm11RajasthanTimeValue(value) {
  const s = String(value ?? '').trim();
  if (!s) return false;
  if (/^nil$/i.test(s)) return false;
  if (looksLikeForm11RajasthanPersonName(s) && !/\d/.test(s)) return false;
  return (
    /^\d{1,2}([:.]?\d{2})?\s*(a\.?m\.?|p\.?m\.?)?$/i.test(s) ||
    /^\d{1,2}\s*(a\.?m\.?|p\.?m\.?)$/i.test(s)
  );
}

function collectForm11RajasthanRoleCandidateValues(row, role) {
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
    if (classifyForm11RJHeader(k) === role) push(v);
  });
  return out;
}

/**
 * Pick the best cell for a Form 11 RJ role.
 * Fixes shifted downloads where Name holds EmployeeID and Ceases holds the real name.
 */
export function pickForm11RajasthanRoleValue(row, role) {
  if (!row || !role) return '';
  const roleVals = collectForm11RajasthanRoleCandidateValues(row, role);

  if (role === 'name') {
    const fromName = roleVals.find((v) => looksLikeForm11RajasthanPersonName(v));
    if (fromName) return fromName;
    // Classic shift: name column has "4", ceases column has "virendra".
    const misplaced = collectForm11RajasthanRoleCandidateValues(row, 'ceases')
      .concat(collectForm11RajasthanRoleCandidateValues(row, 'commences'))
      .concat(collectForm11RajasthanRoleCandidateValues(row, 'young'))
      .find((v) => looksLikeForm11RajasthanPersonName(v));
    if (misplaced) return misplaced;
    // Last resort: any non-meta string value that looks like a name.
    for (const [k, v] of Object.entries(row)) {
      if (String(k).startsWith('__')) continue;
      if (/employee\s*id|^empid$/i.test(norm(k))) continue;
      if (/^employeeid$/i.test(norm(k).replace(/\s+/g, ''))) continue;
      if (looksLikeForm11RajasthanPersonName(v)) return String(v).trim();
    }
    return '';
  }

  if (role === 'ceases' || role === 'commences' || role === 'day1' || role === 'day3') {
    const timed = roleVals.find((v) => looksLikeForm11RajasthanTimeValue(v));
    if (timed) return timed;
    const nonName = roleVals.find(
      (v) => v && !/^nil$/i.test(v) && !looksLikeForm11RajasthanPersonName(v) && !isForm11RajasthanEmployeeIdLikeValue(v)
    );
    if (nonName) return nonName;
    if (role === 'commences') return FORM_11_RJ_COMMENCES_DEFAULT;
    if (role === 'ceases') return FORM_11_RJ_CEASES_DEFAULT;
    if (role === 'day1') return FORM_11_RJ_REST_BEFORE_DEFAULT;
    if (role === 'day3') return FORM_11_RJ_REST_AFTER_DEFAULT;
    return '';
  }

  if (role === 'young') {
    const good = roleVals.find(
      (v) => v && !/^nil$/i.test(v) && !looksLikeForm11RajasthanPersonName(v) && !isForm11RajasthanEmployeeIdLikeValue(v)
    );
    if (good) return good;
    return FORM_11_RJ_YOUNG_DEFAULT;
  }

  if (role === 'otDays' || role === 'otMonth' || role === 'otQuarter' || role === 'otYear') {
    const ot = roleVals.find((v) => v && !looksLikeForm11RajasthanPersonName(v));
    return ot || FORM_11_RJ_OT_DEFAULT;
  }

  if (role === 'totalHours') {
    const hours = roleVals.find((v) => v && !isForm11RajasthanJunkTotalHoursValue(v));
    return hours ? sanitizeForm11RajasthanTotalHoursValue(hours) : '';
  }

  if (role === 'rest') {
    const rest = roleVals.find((v) => v && !/^nil$/i.test(v) && !looksLikeForm11RajasthanPersonName(v));
    return rest || '';
  }

  return roleVals[0] || '';
}

/** Rebuild one row onto canonical headers, repairing ID/name column swaps. */
export function repairForm11RajasthanExportRow(row, targetHeaders = FORM_11_RJ_TABLE_HEADERS) {
  const tgt = resolveForm11RajasthanTableHeaders(targetHeaders);
  const out = {};
  tgt.forEach((header) => {
    const role = classifyForm11RJHeader(header);
    const picked = role ? pickForm11RajasthanRoleValue(row, role) : '';
    out[header] = sanitizeForm11RajasthanCellValue(picked, header);
  });
  if (row && typeof row === 'object') {
    Object.keys(row).forEach((k) => {
      if (String(k).startsWith('__')) out[k] = row[k];
    });
  }
  // Ensure name is never left as a bare employee id when a real name was found.
  const nameHdr = tgt.find((h) => classifyForm11RJHeader(h) === 'name');
  if (nameHdr && isForm11RajasthanEmployeeIdLikeValue(out[nameHdr])) {
    const recovered = pickForm11RajasthanRoleValue(row, 'name');
    if (looksLikeForm11RajasthanPersonName(recovered)) {
      out[nameHdr] = sanitizeForm11RajasthanCellValue(recovered, nameHdr);
    } else {
      out[nameHdr] = '';
    }
  }
  return out;
}

/** Strip HTML / help text that Zoho sometimes dumps into cells. */
export function sanitizeForm11RajasthanCellValue(value, header = '') {
  let s = String(value ?? '').trim();
  if (!s) return '';
  if (/<[^>]+>/.test(s) || /please\s+follow\s+the\s+steps/i.test(s)) return '';
  s = s.replace(/<[^>]*>/g, '').trim();
  if (header && isForm11RajasthanTotalHoursHeader(header)) {
    return sanitizeForm11RajasthanTotalHoursValue(s);
  }
  return s;
}

export function remapForm11RajasthanRowsToHeaders(rows, sourceHeaders, targetHeaders) {
  const tgt = resolveForm11RajasthanTableHeaders(targetHeaders?.length ? targetHeaders : sourceHeaders);
  if (!Array.isArray(rows)) return [];
  return rows.map((row) => {
    if (!row || typeof row !== 'object') {
      const blank = {};
      tgt.forEach((h) => {
        blank[h] = '';
      });
      return blank;
    }
    // Prefer role-aware repair (Name↔ID / name parked under Ceases).
    return repairForm11RajasthanExportRow(row, tgt);
  });
}

export function isForm11RajasthanNameHeader(header) {
  return classifyForm11RJHeader(header) === 'name';
}

export function isForm11RajasthanYoungPersonHeader(header) {
  return classifyForm11RJHeader(header) === 'young';
}

export function isForm11RajasthanCommencesHeader(header) {
  return classifyForm11RJHeader(header) === 'commences';
}

export function isForm11RajasthanCeasesHeader(header) {
  return classifyForm11RJHeader(header) === 'ceases';
}

export function isForm11RajasthanTotalHoursHeader(header) {
  return classifyForm11RJHeader(header) === 'totalHours';
}

/** Rest-interval "before" column (day 1). */
export function isForm11RajasthanRestBeforeHeader(header) {
  return classifyForm11RJHeader(header) === 'day1';
}

/** Rest-interval "after" column (day 3). */
export function isForm11RajasthanRestAfterHeader(header) {
  return classifyForm11RJHeader(header) === 'day3';
}

export function isForm11RajasthanOtHeader(header) {
  const role = classifyForm11RJHeader(header);
  return role === 'otDays' || role === 'otMonth' || role === 'otQuarter' || role === 'otYear';
}

/** Middle Rest interval column stays blank during People autofill. */
export function isForm11RajasthanSkipAutofillHeader(header) {
  return classifyForm11RJHeader(header) === 'rest';
}

/** Whether young person or not — default "Young". */
export function resolveForm11RajasthanYoungPersonValue(_ageYears) {
  return FORM_11_RJ_YOUNG_DEFAULT;
}

/** Reject tenure/experience text like "1 month(s)" from total-hours cells. */
export function isForm11RajasthanJunkTotalHoursValue(value) {
  const s = String(value ?? '').trim();
  if (!s) return false;
  // Times with AM/PM (or HH:MM) must never sit under Total hours — blank when no Paid_days.
  // Do not use looksLikeForm11RajasthanTimeValue here: it can match bare "184".
  if (/[ap]\.?m\.?$/i.test(s) || /^\d{1,2}:\d{2}\b/i.test(s)) return true;
  if (/^nil$/i.test(s)) return true;
  if (/month|year|week|day\s*\(|experience|tenure/i.test(s) && !/^\d+(\.\d+)?$/.test(s)) {
    return true;
  }
  if (!/^\d+(\.\d+)?$/.test(s.replace(/,/g, ''))) return true;
  return false;
}

/** Keep only numeric hours; blank when missing or junk. */
export function sanitizeForm11RajasthanTotalHoursValue(value) {
  const s = String(value ?? '').trim();
  if (!s || isForm11RajasthanJunkTotalHoursValue(s)) return '';
  const n = Number(s.replace(/,/g, ''));
  if (!Number.isFinite(n) || n < 0) return '';
  return String(Math.round(n * 100) / 100);
}

/** Total hours worked during the month = Paid_days × 8. */
export function computeForm11RajasthanTotalHours(paidDays) {
  if (paidDays === '' || paidDays == null) return '';
  const raw = String(paidDays).trim();
  if (!raw || /month|year|week|experience|tenure/i.test(raw)) return '';
  const n = Number(raw.replace(/,/g, ''));
  if (!Number.isFinite(n) || n < 0) return '';
  return String(Math.round(n * 8 * 100) / 100);
}

/** Read Paid_days (and aliases) from a Sample Payroll / Payroll table row. */
export function readForm11RajasthanPaidDays(payrollRow) {
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

/** Read a Form 11 RJ cell by exact header or matching column role (avoids key mismatch blanks). */
export function getForm11RajasthanRowValueForHeader(row, header) {
  if (!row || typeof row !== 'object') return '';
  const role = classifyForm11RJHeader(header);
  if (role) {
    const picked = pickForm11RajasthanRoleValue(row, role);
    if (picked != null && String(picked).trim() !== '') return String(picked).trim();
  }
  const direct = row[header];
  if (direct != null && String(direct).trim() !== '' && String(direct) !== 'false') {
    // Never surface a bare employee id as the Name column.
    if (role === 'name' && isForm11RajasthanEmployeeIdLikeValue(direct)) return '';
    return String(direct);
  }
  return '';
}

function setForm11RajasthanRoleValue(row, headers, role, value) {
  if (!row || !role) return false;
  const hdrs = Array.isArray(headers) ? headers : [];
  let wrote = false;
  hdrs.forEach((h) => {
    if (classifyForm11RJHeader(h) === role) {
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

/**
 * Force modal defaults onto every employee row:
 * Young / 9 AM / 5PM / rest before-after / OT → NIL.
 * Total hours left for Paid_days pass (junk cleared here).
 *
 * Always recovers Name first — SampleData often parks EmployeeID under Name and the
 * real name under Ceases; writing cease defaults before repair would wipe the name.
 * @returns {number} rows updated
 */
export function applyForm11RajasthanAutofillDefaultsToRows(mappedData, tableHeaders, options = {}) {
  if (!Array.isArray(mappedData) || mappedData.length === 0) return 0;
  const headers = resolveForm11RajasthanTableHeaders(tableHeaders);
  const sanitizeValue =
    typeof options.sanitizeValue === 'function'
      ? options.sanitizeValue
      : (v) => String(v ?? '').trim();

  let hits = 0;
  mappedData.forEach((row) => {
    if (!row || typeof row !== 'object') return;

    // Preserve identity / meta before rebuilding onto canonical leaf keys.
    const preservedMeta = {};
    let preservedEmployeeId;
    Object.keys(row).forEach((k) => {
      if (String(k).startsWith('__')) preservedMeta[k] = row[k];
      if (/^employeeid$/i.test(String(k).replace(/\s+/g, ''))) {
        if (preservedEmployeeId == null || String(preservedEmployeeId).trim() === '') {
          preservedEmployeeId = row[k];
        }
      }
    });

    // Repair ID/name column swaps onto canonical keys before overwriting times.
    const repaired = repairForm11RajasthanExportRow(row, headers);
    Object.keys(row).forEach((k) => {
      delete row[k];
    });
    headers.forEach((h) => {
      row[h] = repaired[h];
    });
    Object.keys(preservedMeta).forEach((k) => {
      row[k] = preservedMeta[k];
    });
    if (preservedEmployeeId != null && String(preservedEmployeeId).trim() !== '') {
      row.EmployeeID = preservedEmployeeId;
    }

    setForm11RajasthanRoleValue(row, headers, 'young', sanitizeValue(FORM_11_RJ_YOUNG_DEFAULT));
    setForm11RajasthanRoleValue(row, headers, 'commences', sanitizeValue(FORM_11_RJ_COMMENCES_DEFAULT));
    setForm11RajasthanRoleValue(row, headers, 'ceases', sanitizeValue(FORM_11_RJ_CEASES_DEFAULT));
    setForm11RajasthanRoleValue(row, headers, 'day1', sanitizeValue(FORM_11_RJ_REST_BEFORE_DEFAULT));
    setForm11RajasthanRoleValue(row, headers, 'day3', sanitizeValue(FORM_11_RJ_REST_AFTER_DEFAULT));
    setForm11RajasthanRoleValue(row, headers, 'otDays', sanitizeValue(FORM_11_RJ_OT_DEFAULT));
    setForm11RajasthanRoleValue(row, headers, 'otMonth', sanitizeValue(FORM_11_RJ_OT_DEFAULT));
    setForm11RajasthanRoleValue(row, headers, 'otQuarter', sanitizeValue(FORM_11_RJ_OT_DEFAULT));
    setForm11RajasthanRoleValue(row, headers, 'otYear', sanitizeValue(FORM_11_RJ_OT_DEFAULT));

    // Clear junk total-hours text (incl. leaked 5PM / 9AM); numeric Paid_days×8 is applied separately.
    const hoursHeader = headers.find((h) => isForm11RajasthanTotalHoursHeader(h));
    if (hoursHeader) {
      const cur = String(row[hoursHeader] ?? '').trim();
      if (!cur || isForm11RajasthanJunkTotalHoursValue(cur)) {
        row[hoursHeader] = '';
      } else {
        row[hoursHeader] = sanitizeForm11RajasthanTotalHoursValue(cur);
      }
    }
    hits += 1;
  });
  return hits;
}

/**
 * Fill Total hours worked during the month ← Paid_days × 8.
 * Clears junk like "1 month(s)" when Paid_days is missing.
 * @returns {number} rows updated
 */
export function applyForm11RajasthanPaidDaysToRows(
  mappedData,
  employeesForMapping,
  tableHeaders,
  options = {}
) {
  if (!Array.isArray(mappedData) || mappedData.length === 0) return 0;
  const headers = Array.isArray(tableHeaders) ? tableHeaders : [];
  const hoursHeader = headers.find((h) => isForm11RajasthanTotalHoursHeader(h));
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
    const paidDays = readForm11RajasthanPaidDays(payrollRow);
    const hours = computeForm11RajasthanTotalHours(paidDays);
    const cur = String(row[hoursHeader] ?? '').trim();
    if (hours === '') {
      // No Paid_days — blank the cell (never keep "1 month(s)" / non-numeric junk).
      if (cur && (overwrite || isForm11RajasthanJunkTotalHoursValue(cur))) {
        row[hoursHeader] = '';
        hits += 1;
      }
      return;
    }
    if (!overwrite && cur && !isForm11RajasthanJunkTotalHoursValue(cur)) return;
    row[hoursHeader] = sanitizeValue(hours);
    hits += 1;
  });
  return hits;
}

export function enrichForm11RajasthanDisplayHeader(formHeader, fileName = '', item = null, tableHeaders = [], sheetText = '') {
  if (!isForm11RajasthanContext(formHeader, item, fileName, sheetText, tableHeaders)) {
    return formHeader;
  }
  const base = formHeader && typeof formHeader === 'object' ? { ...formHeader } : {};
  base.title = FORM_11_RJ_DISPLAY_TITLE;
  base.subtitle = FORM_11_RJ_DISPLAY_SUBTITLE;
  base.reference = FORM_11_RJ_DISPLAY_REFERENCE;
  base.form11RJEmploymentLayout = true;

  let fields = Array.isArray(base.fields) ? [...base.fields] : [];
  const hasKey = (k) => fields.some((f) => f.key === k);
  if (!hasKey('form_11_rj_month')) {
    fields.push({ label: 'Month', value: '', key: 'form_11_rj_month' });
  }
  if (!hasKey('form_11_rj_year')) {
    fields.push({ label: 'Year', value: '', key: 'form_11_rj_year' });
  }
  base.fields = fields;
  return base;
}

/** Apply selected payroll month/year onto Form 11 RJ header keys. */
export function applyForm11RajasthanMonthYearToHeaderData(headerData, fullMonth, year) {
  const out = headerData && typeof headerData === 'object' ? { ...headerData } : {};
  const monthVal = String(fullMonth || '').trim();
  const yearVal = String(year || '').trim();
  if (monthVal) out.form_11_rj_month = monthVal;
  if (yearVal) out.form_11_rj_year = yearVal;
  return out;
}

/**
 * Write Month / Year into the Form_11_RJ template placeholders
 * (B7 "Month––––", H7 "Year––––") as "Month : August" / "Year : 2026".
 * Always also writes the canonical B7/H7 cells used by Form_11_RJ.xlsx.
 */
export function writeForm11RajasthanMonthYearToSheetJs(ws, headerFormData) {
  if (!ws || !headerFormData || typeof headerFormData !== 'object') return;
  const monthVal = String(
    headerFormData.form_11_rj_month ||
      headerFormData.form_x_month ||
      headerFormData.Month ||
      headerFormData.month ||
      ''
  ).trim();
  const yearVal = String(
    headerFormData.form_11_rj_year ||
      headerFormData.form_x_year ||
      headerFormData.Year ||
      headerFormData.year ||
      ''
  ).trim();
  if (!monthVal && !yearVal) return;

  const writeText = (ref, value) => {
    // Do not inherit template number formats (can turn "5PM" into "%PM").
    ws[ref] = { t: 's', v: String(value), z: '@' };
  };

  let range;
  try {
    range = ws['!ref'] ? XLSX.utils.decode_range(ws['!ref']) : null;
  } catch (_) {
    range = null;
  }

  let wroteMonth = false;
  let wroteYear = false;
  if (range) {
    for (let r = range.s.r; r <= Math.min(range.e.r, range.s.r + 12); r += 1) {
      for (let c = range.s.c; c <= range.e.c; c += 1) {
        const ref = XLSX.utils.encode_cell({ r, c });
        const cell = ws[ref];
        if (!cell || cell.v == null) continue;
        const text = String(cell.v).trim();
        if (!text) continue;
        if (monthVal && /^month\b/i.test(text) && !/\byear\b/i.test(text)) {
          const rest = text.replace(/^month/i, '').replace(/[-–—:_.\s]/g, '');
          if (!rest || /^month\s*[-–—:_.\s]*$/i.test(text) || !/[a-z]{3,}/i.test(rest)) {
            writeText(ref, `Month : ${monthVal}`);
            wroteMonth = true;
          }
        }
        if (yearVal && /^year\b/i.test(text) && !/\bmonth\b/i.test(text)) {
          const rest = text.replace(/^year/i, '').replace(/[-–—:_.\s]/g, '');
          if (!rest || /^year\s*[-–—:_.\s]*$/i.test(text) || !/\d{4}/.test(rest)) {
            writeText(ref, `Year : ${yearVal}`);
            wroteYear = true;
          }
        }
      }
    }
  }

  // Canonical Form_11_RJ.xlsx anchors (row 7 / B and H).
  if (monthVal && !wroteMonth) writeText('B7', `Month : ${monthVal}`);
  if (yearVal && !wroteYear) writeText('H7', `Year : ${yearVal}`);
  if (monthVal && wroteMonth) {
    const b7 = ws.B7;
    if (!b7 || !/month\s*:/i.test(String(b7.v || ''))) writeText('B7', `Month : ${monthVal}`);
  }
  if (yearVal && wroteYear) {
    const h7 = ws.H7;
    if (!h7 || !/year\s*:/i.test(String(h7.v || ''))) writeText('H7', `Year : ${yearVal}`);
  }
}

/** True when a SheetJS workbook looks like Rajasthan Form 11 Register of Employment. */
export function sheetJsWorkbookLooksLikeForm11Rajasthan(workbook) {
  if (!workbook?.SheetNames?.length) return false;
  const maxSheets = Math.min(workbook.SheetNames.length, 3);
  for (let si = 0; si < maxSheets; si += 1) {
    const ws = workbook.Sheets[workbook.SheetNames[si]];
    if (!ws?.['!ref']) continue;
    let range;
    try {
      range = XLSX.utils.decode_range(ws['!ref']);
    } catch (_) {
      continue;
    }
    let sawForm11 = false;
    let sawRegister = false;
    let sawNameEmployed = false;
    let sawRajasthan = false;
    for (let r = range.s.r; r <= Math.min(range.e.r, range.s.r + 14); r += 1) {
      for (let c = range.s.c; c <= Math.min(range.e.c, range.s.c + 16); c += 1) {
        const cell = ws[XLSX.utils.encode_cell({ r, c })];
        if (!cell || cell.v == null) continue;
        const t = String(cell.v).replace(/\r?\n/g, ' ').toLowerCase();
        if (/^form\s*11\b/.test(t.trim()) || /\bform\s*11\b/.test(t)) sawForm11 = true;
        if (/register\s+of\s+employment/.test(t)) sawRegister = true;
        if (/name\s+of\s+persons?\s+employ/.test(t)) sawNameEmployed = true;
        if (/rajasthan/.test(t)) sawRajasthan = true;
      }
    }
    if (sawRegister && (sawForm11 || sawNameEmployed || sawRajasthan)) return true;
    if (sawForm11 && sawNameEmployed) return true;
  }
  return false;
}

/**
 * After ExcelJS load/write, force Form 11 time cells back to plain text
 * (template % formats otherwise display "5PM" as "%PM").
 */
export function repairForm11RajasthanExcelJsWorksheet(worksheet) {
  if (!worksheet) return false;
  let isForm11 = false;
  const maxScanRow = Math.min(Number(worksheet.rowCount) || 20, 20);
  for (let r = 1; r <= maxScanRow; r += 1) {
    const row = worksheet.getRow(r);
    row.eachCell({ includeEmpty: false }, (cell) => {
      const t = String(
        cell?.value && typeof cell.value === 'object' && cell.value.text != null
          ? cell.value.text
          : cell?.value ?? ''
      )
        .replace(/\r?\n/g, ' ')
        .toLowerCase();
      if (/register\s+of\s+employment/.test(t) || /^form\s*11\b/.test(t.trim())) isForm11 = true;
      if (/name\s+of\s+persons?\s+employ/.test(t) && /rajasthan|rule\s*22/.test(t)) isForm11 = true;
    });
    if (isForm11) break;
  }
  if (!isForm11) {
    // Title band often split across rows: FORM 11 + Register of Employment.
    let sawForm11 = false;
    let sawRegister = false;
    for (let r = 1; r <= maxScanRow; r += 1) {
      const row = worksheet.getRow(r);
      row.eachCell({ includeEmpty: false }, (cell) => {
        const t = String(
          cell?.value && typeof cell.value === 'object' && cell.value.text != null
            ? cell.value.text
            : cell?.value ?? ''
        )
          .replace(/\r?\n/g, ' ')
          .toLowerCase();
        if (/^form\s*11\b/.test(t.trim()) || /\bform\s*11\b/.test(t)) sawForm11 = true;
        if (/register\s+of\s+employment/.test(t)) sawRegister = true;
      });
    }
    isForm11 = sawForm11 && sawRegister;
  }
  if (!isForm11) return false;

  const plain = (cell) => {
    const v = cell?.value;
    if (v == null) return '';
    if (typeof v === 'object') {
      if (v.text != null) return String(v.text);
      if (v.result != null) return String(v.result);
      if (v.richText) return v.richText.map((x) => x.text || '').join('');
    }
    return String(v);
  };

  const forceText = (cell, text) => {
    cell.value = String(text);
    cell.numFmt = '@';
  };

  // Month / Year row (typically 7): keep label text, do not leave bare dashes.
  for (let r = 5; r <= 9; r += 1) {
    const row = worksheet.getRow(r);
    row.eachCell({ includeEmpty: false }, (cell) => {
      const text = plain(cell).trim();
      if (/^month\b/i.test(text) && !/\byear\b/i.test(text)) {
        forceText(cell, text);
      }
      if (/^year\b/i.test(text) && !/\bmonth\b/i.test(text)) {
        forceText(cell, text);
      }
    });
  }

  // Data rows: repair mangled times and force @ format on B–M body cells.
  // Never put 5PM/9AM into the Total-hours column (blank when no numeric hours).
  const lastRow = Math.min(Number(worksheet.rowCount) || 40, 200);
  let totalHoursCol = 0;
  for (let r = 8; r <= 14; r += 1) {
    for (let c = 2; c <= 14; c += 1) {
      const t = plain(worksheet.getRow(r).getCell(c)).trim().toLowerCase();
      if (/total\s+hours\s+worked/.test(t)) {
        totalHoursCol = c;
        break;
      }
    }
    if (totalHoursCol) break;
  }
  for (let r = 13; r <= lastRow; r += 1) {
    const row = worksheet.getRow(r);
    for (let c = 2; c <= 13; c += 1) {
      const cell = row.getCell(c);
      const text = plain(cell).trim();
      if (!text) continue;
      if (totalHoursCol && c === totalHoursCol) {
        const hours = sanitizeForm11RajasthanTotalHoursValue(text);
        if (hours !== text) forceText(cell, hours);
        continue;
      }
      // Excel % format turns leading "5" of "5PM" into a percent sign.
      if (/^%[\s]*PM$/i.test(text)) {
        forceText(cell, '5PM');
        continue;
      }
      if (/^%[\s]*AM$/i.test(text)) {
        forceText(cell, '9AM');
        continue;
      }
      const ampm = text.match(/^(\d{1,2})\s*([AP]M)$/i);
      if (ampm) {
        forceText(cell, `${ampm[1]}${ampm[2].toUpperCase()}`);
        continue;
      }
      if (/^young$/i.test(text) || /^nil$/i.test(text) || looksLikeForm11RajasthanTimeValue(text)) {
        forceText(cell, text);
      }
    }
  }
  return true;
}

/** True when a Form 11 cell should be forced as Excel text (times like 5PM / 9AM). */
export function isForm11RajasthanForceTextCellValue(value, header = '') {
  const s = String(value ?? '').trim();
  if (!s) return false;
  const role = classifyForm11RJHeader(header);
  if (role === 'commences' || role === 'ceases' || role === 'day1' || role === 'day3' || role === 'young') {
    return true;
  }
  if (role === 'name' || role === 'otDays' || role === 'otMonth' || role === 'otQuarter' || role === 'otYear') {
    return true;
  }
  return looksLikeForm11RajasthanTimeValue(s) || /^nil$/i.test(s) || /^young$/i.test(s);
}

export function isForm11RajasthanEmploymentLayoutFormHeader(formHeader) {
  return !!formHeader?.form11RJEmploymentLayout;
}

/**
 * Locate header band + data start from workbook (skip column-number row 1…12).
 */
export function repairForm11RajasthanTableHeadersFromWorkbook(workbook, hints = {}) {
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

  const leafRow = headerRow + 1;
  const headers = [];
  for (let i = 0; i < 12; i += 1) {
    const c = startCol + i;
    const top = stripDedupeSuffix(mergedAware(headerRow, c));
    const leaf = stripDedupeSuffix(mergedAware(leafRow, c));
    let chosen = '';
    if (i === 0) chosen = top || FORM_11_RJ_TABLE_HEADERS[0];
    else if (i === 1) chosen = top || FORM_11_RJ_TABLE_HEADERS[1];
    else if (i === 2) chosen = leaf || FORM_11_RJ_TABLE_HEADERS[2];
    else if (i === 3) chosen = leaf || FORM_11_RJ_TABLE_HEADERS[3];
    else if (i === 4) {
      // Day 1 may sit on the top row (legacy template) or leaf row (Days-of-Month group).
      if (/^\d+$/.test(leaf)) chosen = leaf;
      else if (/^\d+$/.test(top) && !/days\s+of\s+month/i.test(top)) chosen = top;
      else chosen = FORM_11_RJ_TABLE_HEADERS[4];
    } else if (i === 5) {
      const dayTop = /^\d+$/.test(top) && !/days\s+of\s+month/i.test(top) ? top : '';
      const leafFirst = String(leaf || '').split(/\s+|\n/)[0] || '';
      const dayLeaf = /^\d+$/.test(leafFirst) ? leafFirst : '';
      const day = dayLeaf || dayTop || '2';
      chosen = `${day} Rest interval`.replace(/\s+/g, ' ').trim();
    } else if (i === 6) {
      if (/^\d+$/.test(leaf)) chosen = leaf;
      else if (/^\d+$/.test(top) && !/days\s+of\s+month/i.test(top)) chosen = top;
      else chosen = FORM_11_RJ_TABLE_HEADERS[6];
    } else if (i >= 7) chosen = top || FORM_11_RJ_TABLE_HEADERS[i];
    headers.push(chosen || FORM_11_RJ_TABLE_HEADERS[i]);
  }

  // Column-number row (1…12) sits under the leaf band.
  let dataStart = leafRow + 1;
  let numHits = 0;
  for (let i = 0; i < 12; i += 1) {
    const t = String(mergedAware(dataStart, startCol + i) || '').trim();
    if (t === String(i + 1)) numHits += 1;
  }
  if (numHits >= 8) dataStart += 1;

  return {
    headers: resolveForm11RajasthanTableHeaders(headers),
    headerRowIndex: headerRow,
    dataStartIndex: dataStart,
    tableStartCol: startCol,
    stackedHeaderBandEnd: leafRow,
  };
}

function excelCellValueToString(val) {
  if (val == null) return '';
  if (typeof val === 'string' || typeof val === 'number' || typeof val === 'boolean') return String(val);
  if (val instanceof Date) return val.toISOString();
  if (typeof val === 'object') {
    if (Array.isArray(val.richText)) return val.richText.map((rt) => rt?.text || '').join('');
    if (val.text != null) return String(val.text);
    if (val.result != null) return String(val.result);
  }
  return '';
}

function buildMergeTopLeftResolver(worksheet) {
  const merges = Array.isArray(worksheet?._merges)
    ? Object.values(worksheet._merges)
    : Array.isArray(worksheet?.model?.merges)
      ? worksheet.model.merges
      : [];
  const map = new Map();
  merges.forEach((m) => {
    const top = m?.top ?? m?.s?.r;
    const left = m?.left ?? m?.s?.c;
    const bottom = m?.bottom ?? m?.e?.r;
    const right = m?.right ?? m?.e?.c;
    if (top == null || left == null || bottom == null || right == null) return;
    // ExcelJS model merges may be 0-based (s/e) or 1-based (top/left).
    const isZeroBased = m?.s != null && m?.top == null;
    const t = isZeroBased ? top + 1 : top;
    const l = isZeroBased ? left + 1 : left;
    const b = isZeroBased ? bottom + 1 : bottom;
    const rgt = isZeroBased ? right + 1 : right;
    for (let r = t; r <= b; r += 1) {
      for (let c = l; c <= rgt; c += 1) {
        map.set(`${r}:${c}`, { r: t, c: l });
      }
    }
  });
  return (r, c) => map.get(`${r}:${c}`) || { r, c };
}

/**
 * Locate Form 11 RJ table (Name at column B…) + data start (skip 1…12 index row).
 * Returns 1-based ExcelJS coordinates.
 */
export function detectForm11RajasthanTableLayout(worksheet, hints = {}) {
  if (!worksheet) return null;
  const getMergeTopLeft = buildMergeTopLeftResolver(worksheet);
  const getMergedAwareCellText = (r, c) => {
    const tl = getMergeTopLeft(r, c);
    return excelCellValueToString(worksheet.getCell(tl.r, tl.c)?.value).trim();
  };

  let headerRow =
    hints.parsedHeaderRowIndex != null && hints.parsedHeaderRowIndex >= 0
      ? hints.parsedHeaderRowIndex + 1
      : -1;
  // Form_11_RJ.xlsx Name column is B (ExcelJS col 2). Never allow column A.
  let startCol =
    hints.parsedTableStartCol != null && Number(hints.parsedTableStartCol) >= 1
      ? Number(hints.parsedTableStartCol) + 1
      : 2;
  if (startCol < 2) startCol = 2;

  const maxScanRows = Math.max(40, (worksheet.rowCount || 0) + 5);
  if (headerRow < 1) {
    for (let r = 1; r <= maxScanRows; r += 1) {
      for (let c = 1; c <= 40; c += 1) {
        const t = norm(getMergedAwareCellText(r, c));
        if (/name\s+of\s+persons?\s+employ/.test(t)) {
          headerRow = r;
          startCol = Math.max(2, c);
          break;
        }
      }
      if (headerRow > 0) break;
    }
  }
  if (headerRow < 1) {
    for (let r = 1; r <= maxScanRows; r += 1) {
      for (let c = 1; c <= 40; c += 1) {
        const t = norm(getMergedAwareCellText(r, c));
        if (/young\s+person/.test(t) && /employment\s+commence|time\s+at\s+which/.test(
          norm(getMergedAwareCellText(r, c + 1) + ' ' + getMergedAwareCellText(r + 1, c + 1))
        )) {
          headerRow = r;
          startCol = Math.max(2, c - 1);
          break;
        }
      }
      if (headerRow > 0) break;
    }
  }
  if (headerRow < 1) return null;
  startCol = Math.max(2, startCol);

  // Leaf / sub-header row under the Days-of-Month band.
  let leafRow = headerRow;
  for (let r = headerRow; r <= Math.min(headerRow + 3, maxScanRows); r += 1) {
    let commenceHits = 0;
    for (let c = startCol; c <= startCol + 11; c += 1) {
      const t = norm(getMergedAwareCellText(r, c));
      if (/employment\s+commence|employment\s+cease|rest\s+interval/.test(t)) commenceHits += 1;
    }
    if (commenceHits >= 2) {
      leafRow = r;
      break;
    }
  }

  // Skip the 1…12 column-index row when present.
  let dataStart = leafRow + 1;
  let numHits = 0;
  for (let i = 0; i < 12; i += 1) {
    const t = String(getMergedAwareCellText(dataStart, startCol + i) || '').trim();
    if (t === String(i + 1)) numHits += 1;
  }
  if (numHits >= 8) dataStart += 1;

  if (
    hints.parsedDataStartIndex != null &&
    Number(hints.parsedDataStartIndex) >= 0
  ) {
    const hinted = Number(hints.parsedDataStartIndex) + 1;
    // Prefer the later of hinted vs scanned so we never write over the 1…12 index row.
    dataStart = Math.max(dataStart, hinted);
  }

  const templateCols = FORM_11_RJ_TABLE_HEADERS.map((label, idx) => ({
    col: startCol + idx,
    label,
    role: classifyForm11RJHeader(label),
  }));

  return {
    headerRow,
    leafRow,
    headerBandEnd: leafRow,
    dataStartRow: dataStart,
    startCol,
    templateCols,
  };
}

/**
 * Form_11_RJ.xlsx / 1st-image model:
 *   Top:  Days of Month (merge over commence+cease) | 1 | 2 | 3
 *   Leaf: Commences | Ceases | (blank) | Rest interval | (blank)
 * Rewrites any colspan-5 / flat leaf layout back to this band.
 */
export function ensureForm11RajasthanDaysOfMonthHeaderBand(worksheet, layout) {
  if (!worksheet || !layout) return false;
  const headerRow = Number(layout.headerRow) || 0;
  let leafRow = Number(layout.leafRow) || headerRow + 1;
  const startCol = Math.max(2, Number(layout.startCol) || 2);
  if (headerRow < 1) return false;
  if (leafRow <= headerRow) leafRow = headerRow + 1;

  const daysStart = startCol + 2; // Commences column (D when startCol=B)
  const day1Col = daysStart + 2;
  const day2Col = daysStart + 3;
  const day3Col = daysStart + 4;
  const daysEnd = day3Col;

  const colLettersToNumber = (letters) => {
    const s = String(letters || '').toUpperCase();
    let n = 0;
    for (let i = 0; i < s.length; i += 1) {
      n = n * 26 + (s.charCodeAt(i) - 64);
    }
    return n;
  };

  const unmergeOverlapping = (r1, c1, r2, c2) => {
    const merges = worksheet?.model?.merges;
    if (!Array.isArray(merges) || !merges.length) return;
    const toRemove = [];
    merges.forEach((range) => {
      const parts = String(range || '').split(':');
      if (parts.length !== 2) return;
      const start = parts[0].match(/^([A-Z]+)(\d+)$/i);
      const end = parts[1].match(/^([A-Z]+)(\d+)$/i);
      if (!start || !end) return;
      const sr = parseInt(start[2], 10);
      const er = parseInt(end[2], 10);
      const sc = colLettersToNumber(start[1]);
      const ec = colLettersToNumber(end[1]);
      if (Number.isNaN(sc) || Number.isNaN(ec)) return;
      const rowOverlap = sr <= r2 && er >= r1;
      const colOverlap = sc <= c2 && ec >= c1;
      if (rowOverlap && colOverlap) toRemove.push(String(range));
    });
    toRemove.forEach((range) => {
      try {
        worksheet.unMergeCells(range);
      } catch (_) {
        /* ignore */
      }
    });
  };

  // Clear any Days-of-Month merges across the 5-col band (including a prior D:H span).
  unmergeOverlapping(headerRow, daysStart, headerRow, daysEnd);

  const center = {
    horizontal: 'center',
    vertical: 'middle',
    wrapText: true,
  };
  const setHeaderCell = (r, c, text) => {
    const cell = worksheet.getCell(r, c);
    cell.value = text;
    cell.alignment = { ...center };
    cell.numFmt = '@';
  };

  // Top row: Days of Month | 1 | 2 | 3
  for (let c = daysStart; c <= daysEnd; c += 1) {
    setHeaderCell(headerRow, c, '');
  }
  setHeaderCell(headerRow, daysStart, 'Days of Month');
  try {
    worksheet.mergeCells(headerRow, daysStart, headerRow, daysStart + 1);
  } catch (_) {
    /* ignore */
  }
  setHeaderCell(headerRow, day1Col, '1');
  setHeaderCell(headerRow, day2Col, '2');
  setHeaderCell(headerRow, day3Col, '3');

  // Leaf row: Commences | Ceases | (blank under 1) | Rest interval | (blank under 3)
  setHeaderCell(leafRow, daysStart, FORM_11_RJ_TABLE_HEADERS[2]);
  setHeaderCell(leafRow, daysStart + 1, FORM_11_RJ_TABLE_HEADERS[3]);
  setHeaderCell(leafRow, day1Col, '');
  setHeaderCell(leafRow, day2Col, 'Rest interval');
  setHeaderCell(leafRow, day3Col, '');

  return true;
}

function writeForm11RajasthanMonthYearExcelJs(worksheet, headerFormData, headerRowEnd = 12) {
  if (!worksheet || !headerFormData) return;
  const monthVal = String(
    headerFormData.form_11_rj_month ||
      headerFormData.form_x_month ||
      headerFormData.Month ||
      headerFormData.month ||
      ''
  ).trim();
  const yearVal = String(
    headerFormData.form_11_rj_year ||
      headerFormData.form_x_year ||
      headerFormData.Year ||
      headerFormData.year ||
      ''
  ).trim();
  if (!monthVal && !yearVal) return;

  const forceText = (cell, text) => {
    cell.value = String(text);
    cell.numFmt = '@';
  };

  let wroteMonth = false;
  let wroteYear = false;
  const maxR = Math.max(1, Number(headerRowEnd) || 12);
  for (let r = 1; r <= maxR; r += 1) {
    const row = worksheet.getRow(r);
    row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
      const text = excelCellValueToString(cell?.value).trim();
      if (!text) return;
      if (monthVal && /^month\b/i.test(text) && !/\byear\b/i.test(text)) {
        const rest = text.replace(/^month/i, '').replace(/[-–—:_.\s]/g, '');
        if (!rest || /^month\s*[-–—:_.\s]*$/i.test(text) || !/[a-z]{3,}/i.test(rest)) {
          forceText(cell, `Month : ${monthVal}`);
          wroteMonth = true;
        }
      }
      if (yearVal && /^year\b/i.test(text) && !/\bmonth\b/i.test(text)) {
        const rest = text.replace(/^year/i, '').replace(/[-–—:_.\s]/g, '');
        if (!rest || /^year\s*[-–—:_.\s]*$/i.test(text) || !/\d{4}/.test(rest)) {
          forceText(cell, `Year : ${yearVal}`);
          wroteYear = true;
        }
      }
      void colNumber;
    });
  }
  // Canonical Form_11_RJ.xlsx anchors.
  if (monthVal && !wroteMonth) forceText(worksheet.getCell(7, 2), `Month : ${monthVal}`);
  if (yearVal && !wroteYear) forceText(worksheet.getCell(7, 8), `Year : ${yearVal}`);
}

/**
 * Paint Form 11 RJ onto the clean formmaster template with ExcelJS so
 * Name/Young/times land in the original bordered boxes (not SheetJS shift / merge boxes).
 */
export async function buildForm11RajasthanWorkbookWithTemplateStyles({
  templateArrayBuffer,
  mappedData,
  headersToUse,
  parsedFormHeader,
  formFileName,
  headerFormData,
  parsedHeaderRowIndex,
  parsedDataStartIndex,
  parsedTableStartCol,
}) {
  if (!templateArrayBuffer) {
    throw new Error('Form 11 RJ template buffer is required.');
  }
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(templateArrayBuffer);
  const worksheet = workbook.worksheets[0];
  if (!worksheet) throw new Error('Template worksheet not found.');

  const layout = detectForm11RajasthanTableLayout(worksheet, {
    parsedHeaderRowIndex,
    parsedDataStartIndex,
    parsedTableStartCol:
      parsedTableStartCol != null && Number(parsedTableStartCol) >= 1
        ? Number(parsedTableStartCol)
        : 1,
  });
  if (!layout) throw new Error('Could not locate Rajasthan Form 11 table header row.');

  ensureForm11RajasthanDaysOfMonthHeaderBand(worksheet, layout);

  const { dataStartRow, templateCols, headerBandEnd } = layout;
  const protectedHeaderEnd = Math.max(layout.headerRow || 1, headerBandEnd || layout.headerRow || 1);
  const tableColMin = Math.min(...templateCols.map(({ col }) => col));
  const tableColMax = Math.max(...templateCols.map(({ col }) => col));

  writeForm11RajasthanMonthYearExcelJs(
    worksheet,
    headerFormData,
    Math.max(1, protectedHeaderEnd)
  );

  const headers =
    Array.isArray(headersToUse) && headersToUse.length > 0
      ? resolveForm11RajasthanTableHeaders(headersToUse)
      : FORM_11_RJ_TABLE_HEADERS.slice();

  let rows = (Array.isArray(mappedData) ? mappedData : [])
    .filter((row) => row && typeof row === 'object')
    .map((row) => ({ ...row }));
  rows = remapForm11RajasthanRowsToHeaders(rows, headers, headers);
  applyForm11RajasthanAutofillDefaultsToRows(rows, headers);

  const safeDataStart = Math.max(dataStartRow, protectedHeaderEnd + 1);
  const clearToRow = Math.max(safeDataStart + rows.length + 4, safeDataStart + 10);

  // Unmerge body cells so each employee occupies one bordered row (fixes nested "box" look).
  {
    const merges = worksheet?.model?.merges;
    if (Array.isArray(merges) && merges.length > 0) {
      const toRemove = [];
      merges.forEach((range) => {
        const parts = String(range || '').split(':');
        if (parts.length !== 2) return;
        const start = parts[0].match(/^([A-Z]+)(\d+)$/i);
        const end = parts[1].match(/^([A-Z]+)(\d+)$/i);
        if (!start || !end) return;
        const r1 = parseInt(start[2], 10);
        const r2 = parseInt(end[2], 10);
        if (r1 <= protectedHeaderEnd) return;
        if (r2 >= safeDataStart && r1 <= clearToRow) toRemove.push(String(range));
      });
      toRemove.forEach((range) => {
        try {
          worksheet.unMergeCells(range);
        } catch (_) {
          /* ignore */
        }
      });
    }
  }

  // Clear column A leftovers (old drafts wrote EmployeeIDs there) + table band.
  for (let r = safeDataStart; r <= clearToRow; r += 1) {
    for (let c = 1; c <= tableColMax; c += 1) {
      const cell = worksheet.getCell(r, c);
      cell.value = '';
      if (c >= tableColMin) cell.numFmt = '@';
    }
  }

  rows.forEach((row, idx) => {
    const targetRowNum = safeDataStart + idx;
    const targetRow = worksheet.getRow(targetRowNum);
    if (targetRow) targetRow.height = Math.max(Number(targetRow.height) || 0, 22);
    templateCols.forEach(({ col, label, role }) => {
      let val = getForm11RajasthanRowValueForHeader(row, label);
      if ((!val || !String(val).trim()) && role) {
        for (const h of headers) {
          if (classifyForm11RJHeader(h) === role) {
            const viaHeader = getForm11RajasthanRowValueForHeader(row, h);
            if (viaHeader) {
              val = viaHeader;
              break;
            }
          }
        }
      }
      // Never write an EmployeeID into the Name box.
      if (role === 'name' && isForm11RajasthanEmployeeIdLikeValue(val)) {
        val = '';
      }
      // Total hours: numeric only — blank when no Paid_days (never leak 5PM / 9AM / NIL).
      if (role === 'totalHours') {
        val = sanitizeForm11RajasthanTotalHoursValue(val);
      }
      const display = val == null ? '' : String(val).trim();
      const cell = worksheet.getCell(targetRowNum, col);
      cell.value = display;
      cell.numFmt = '@';
      cell.alignment = {
        horizontal: 'left',
        vertical: 'middle',
        wrapText: false,
      };
    });
  });

  // Belt-and-suspenders: blank any non-numeric Total-hours cells in the written band.
  {
    const hoursCols = templateCols
      .filter((c) => c.role === 'totalHours')
      .map((c) => c.col);
    hoursCols.forEach((col) => {
      for (let i = 0; i < rows.length; i += 1) {
        const cell = worksheet.getCell(safeDataStart + i, col);
        const raw = excelCellValueToString(cell?.value).trim();
        const next = sanitizeForm11RajasthanTotalHoursValue(raw);
        if (raw && !next) {
          cell.value = '';
          cell.numFmt = '@';
        } else if (next && next !== raw) {
          cell.value = next;
          cell.numFmt = '@';
        }
      }
    });
  }
  if (rows.length > 0) {
    ensureExcelJSDataRowsWithBorders(worksheet, {
      dataStartRow: safeDataStart,
      dataRowCount: rows.length,
      colFrom: tableColMin,
      colTo: tableColMax,
      templateRow: safeDataStart,
      templateBodyRows: 1,
    });
  }

  // Readable column widths for the register boxes.
  templateCols.forEach(({ col, role }) => {
    const column = worksheet.getColumn(col);
    if (!column) return;
    let w = 14;
    if (role === 'name') w = 28;
    else if (role === 'young') w = 18;
    else if (role === 'commences' || role === 'ceases') w = 16;
    else if (role === 'totalHours') w = 18;
    else if (role === 'otDays' || role === 'otMonth' || role === 'otQuarter' || role === 'otYear') w = 16;
    column.width = Math.max(Number(column.width) || 0, w);
  });

  repairForm11RajasthanExcelJsWorksheet(worksheet);

  const out = await workbook.xlsx.writeBuffer();
  const fileName =
    formFileName ||
    parsedFormHeader?.title?.replace(/[^a-zA-Z0-9]/g, '_') ||
    'Form_11_RJ_-_Rajasthan.xlsx';
  return {
    blob: new Blob([out], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    }),
    fileName,
    layout: {
      dataStartRow: safeDataStart,
      startCol: tableColMin,
      rowCount: rows.length,
    },
  };
}

