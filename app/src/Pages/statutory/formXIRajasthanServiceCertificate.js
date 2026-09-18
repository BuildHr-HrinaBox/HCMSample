import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { matchesFormXIHint } from './formXAPRegisterOfFines';
import { formatWorkmanNameAndGuardian } from './formXIXAPWageSlip';
import {
  buildFormXIXMPPayrollRowResolver,
  resolveFormXIXMPPayrollRowForEmployee,
} from './formXIXMPWageSlip';
import {
  readFormXIVMPPayrollGrossPay,
  readFormXIVMPPayrollNetPay,
} from './formXIVMPEmploymentCard';
import { flattenPayrollEarningColumns, readPayrollScalar } from '../../utils/payrollEarnings';
import { resolveHeaderFieldExportValue } from '../../utils/statutorySiteCompanyHeaders';
import { applyExcelJSFullBoxBordersToRange } from '../../utils/excelTableBorders';

/** Rajasthan CLRA Form XI — Service Certificate [Rule 76]; per-employee ZIP export. */

export const FORM_XI_RJ_STACKED_VALUE_COL = 5;

export const FORM_XI_RJ_FIELD_GROUPS = [
  { id: 'header', title: 'Establishment particulars' },
  { id: 'workman', title: 'Workman particulars' },
  { id: 'table', title: 'Employment & wages' },
];

export const FORM_XI_RJ_SITE_HEADER_SPECS = [
  {
    key: 'form_xi_rj_contractor',
    label: 'Name and address of contractor',
    group: 'header',
    fieldType: 'textarea',
    match: /name\s+and\s+address\s+of\s+(?:the\s+)?contractor/i,
  },
  {
    key: 'form_xi_rj_nature_location',
    label: 'Nature and location of work',
    group: 'header',
    fieldType: 'textarea',
    match: /nature\s+and\s+location\s+of\s+work/i,
  },
  {
    key: 'form_xi_rj_establishment',
    label: 'Name and address of establishment in/under which contract is carried on',
    group: 'header',
    fieldType: 'textarea',
    match:
      /name\s+and\s+address\s+of\s+establishment[\s\S]{0,80}?contract\s+is\s+carried\s+on|establishment\s+in\s*\/\s*under\s+which(?:\s+contract)?/i,
  },
  {
    key: 'form_xi_rj_principal_employer',
    label: 'Name and address of principal employer',
    group: 'header',
    fieldType: 'textarea',
    match: /name\s+and\s+address\s+of\s+(?:the\s+)?principal\s+employer/i,
  },
];

export const FORM_XI_RJ_WORKMAN_HEADER_SPECS = [
  {
    key: 'form_xi_rj_workman',
    label: '1. Name and address of the workmen',
    group: 'workman',
    fieldType: 'textarea',
    ordinal: 1,
    match: /(?:^|\s)1[\.\):\-]?\s*name\s+and\s+address\s+of\s+the\s+workm[ae]n|name\s+and\s+address\s+of\s+the\s+workm[ae]n/i,
  },
  {
    key: 'form_xi_rj_age_or_dob',
    label: '2. Age or date of birth',
    group: 'workman',
    fieldType: 'text',
    ordinal: 2,
    match: /(?:^|\s)2[\.\):\-]?\s*age\s+or\s+date\s+of\s+birth|age\s+or\s+date\s+of\s+birth/i,
  },
  {
    key: 'form_xi_rj_identification_marks',
    label: '3. Identification marks',
    group: 'workman',
    fieldType: 'text',
    ordinal: 3,
    match: /(?:^|\s)3[\.\):\-]?\s*identification\s+marks|identification\s+marks/i,
  },
  {
    key: 'form_xi_rj_father_or_husband_name',
    label: "4. Father's / Husband's Name",
    group: 'workman',
    fieldType: 'text',
    ordinal: 4,
    match: /(?:^|\s)4[\.\):\-]?\s*father.*husband.*name|father.*husband.*name|husband.*father.*name/i,
  },
];

export const FORM_XI_RJ_HEADER_SPECS = [
  ...FORM_XI_RJ_SITE_HEADER_SPECS,
  ...FORM_XI_RJ_WORKMAN_HEADER_SPECS,
];

export const FORM_XI_RJ_TABLE_HEADERS = [
  'Sr.No.',
  'Total period for which employed_From',
  'Total period for which employed_To',
  'Actual No. of days worked',
  'Nature of work done',
  'Rate of wages (with particulars of piece work)',
  'Total wages earned by workman during the period',
  'Total deduction made, if any',
  'Total wages actually paid',
  'Remarks',
];

const PAID_DAYS_KEYS = [
  'paid_days',
  'Paid Days',
  'days_worked',
  'Days Worked',
  'paidDays',
  'no_of_days_worked',
  'present_days',
  'Effective Paid Days',
  'effective_paid_days',
];

const PAID_DAYS_PATTERNS = [/^paid_days$/, /paiddays/, /daysworked/, /present_days/, /effective_paid_days/];

export function formXIRJHeaderNorm(txt) {
  return String(txt || '')
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

const normalizeTableHeader = (h) =>
  String(h || '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

export function isFormXIRajasthanServiceCertificateContext(
  formHeader,
  rowItem,
  fileName,
  sheetText = ''
) {
  const identityBlob = [
    rowItem?.formName,
    rowItem?.FormName,
    rowItem?.description,
    rowItem?.Description,
    fileName,
  ]
    .filter((x) => x != null && String(x).trim() !== '')
    .join(' ')
    .toLowerCase();

  const parts = [
    identityBlob,
    formHeader?.title,
    formHeader?.subtitle,
    formHeader?.reference,
    sheetText,
  ]
    .filter((x) => x != null && String(x).trim() !== '')
    .join(' ')
    .toLowerCase();

  // Form XV (Rajasthan Wage Slip) must never open as Form XI Service Certificate.
  if (
    /form[\s._-]*xv(?![a-z])/i.test(identityBlob) ||
    /\bxv[\s._-]*rj\b/i.test(identityBlob) ||
    /form[\s._-]*xv(?![a-z])/i.test(parts)
  ) {
    return false;
  }
  if (/wage\s*slip|wages?\s+slip|rule\s*77\s*\(\s*2\s*\)\s*\(\s*b\s*\)/i.test(parts)) {
    return false;
  }

  if (!matchesFormXIHint(identityBlob || parts)) return false;
  if (/employment\s+card/i.test(identityBlob) && !/service\s+certificate/i.test(identityBlob)) {
    return false;
  }
  if (
    /rajasthan|\bxi[\s._-]*rj\b|form[\s._-]*xi[\s._-]*rj/i.test(parts) ||
    /\bxi_rj\b/i.test(parts) ||
    /maharashtra|\bxi[\s._-]*mh\b|form[\s._-]*xi[\s._-]*mh/i.test(parts) ||
    /\bxi_mh\b/i.test(parts)
  ) {
    return true;
  }
  if (/service\s+certificate/i.test(parts) && /see\s+rule\s+76|rule\s*76/i.test(parts)) {
    return true;
  }
  return matchesFormXIHint(identityBlob) && /service\s+certificate/i.test(parts);
}

export function isFormXIHeaderFieldLayoutFormHeader(formHeader) {
  return !!formHeader?.formXIHeaderFieldLayout;
}

export function isFormXIRJSerialHeader(h) {
  const n = normalizeTableHeader(h);
  return n.includes('sr') && n.includes('no');
}

export function isFormXIRJPeriodFromHeader(h) {
  const n = normalizeTableHeader(h);
  return n.includes('from') && (n.includes('period') || n.includes('employed'));
}

export function isFormXIRJPeriodToHeader(h) {
  const n = normalizeTableHeader(h);
  return n.includes('to') && !n.includes('from') && (n.includes('period') || n.includes('employed'));
}

export function isFormXIRJDaysWorkedHeader(h) {
  const n = normalizeTableHeader(h);
  return n.includes('days') && n.includes('worked');
}

export function isFormXIRJNatureOfWorkHeader(h) {
  const n = normalizeTableHeader(h);
  return n.includes('nature') && n.includes('work');
}

export function isFormXIRJRateOfWageHeader(h) {
  const n = normalizeTableHeader(h);
  return n.includes('rate') && n.includes('wage');
}

export function isFormXIRJGrossWagesHeader(h) {
  const n = normalizeTableHeader(h);
  return n.includes('wages') && n.includes('earned');
}

export function isFormXIRJDeductionHeader(h) {
  const n = normalizeTableHeader(h);
  return n.includes('deduction');
}

export function isFormXIRJNetWagesHeader(h) {
  const n = normalizeTableHeader(h);
  return (n.includes('wages') && n.includes('paid')) || (n.includes('actually') && n.includes('paid'));
}

export function isFormXIRJRemarksHeader(h) {
  return normalizeTableHeader(h).includes('remarks');
}

export function isFormXIRJSkipAutofillHeader(h) {
  return isFormXIRJRemarksHeader(h);
}

export function resolveFormXIRJTableHeaders(tableHeaders) {
  if (Array.isArray(tableHeaders) && tableHeaders.length >= 6) {
    const joined = tableHeaders.map((h) => normalizeTableHeader(h)).join(' ');
    if (joined.includes('days worked') || joined.includes('wages earned')) {
      return tableHeaders;
    }
  }
  return [...FORM_XI_RJ_TABLE_HEADERS];
}

const excelCellValueToString = (val) => {
  if (val == null) return '';
  if (typeof val === 'string' || typeof val === 'number' || typeof val === 'boolean') return String(val);
  if (val instanceof Date) {
    const d = val;
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
  }
  if (typeof val === 'object') {
    if (Array.isArray(val.richText)) return val.richText.map((rt) => rt?.text || '').join('');
    if (val.text != null) return String(val.text);
    if (val.result != null) return String(val.result);
  }
  return '';
};

const isPlaceholderCell = (text) => {
  const s = String(text ?? '').trim();
  if (!s) return true;
  if (/^[\s._\-…·]+$/u.test(s)) return true;
  if (/[.…_-]{4,}$/u.test(s) && s.replace(/[.…_\-\s]+/gu, '').length < 4) return true;
  return false;
};

const buildTemplateField = (spec, coords = {}) => ({
  label: spec.label,
  value: coords.value || '',
  key: spec.key,
  group: spec.group || 'header',
  fieldType: spec.fieldType || 'text',
  labelRow: coords.labelRow ?? null,
  labelCol: coords.labelCol ?? null,
  valueCol: coords.valueCol ?? null,
  valueRow: coords.valueRow ?? null,
});

const stripLeadingOrdinal = (text) =>
  String(text || '')
    .replace(/^\s*\d+[\.\):\-]?\s*/, '')
    .trim();

const labelMatchesSpec = (raw, spec) => {
  const cell = String(raw || '').trim();
  if (!cell) return false;
  const norm = formXIRJHeaderNorm(cell);
  const stripped = formXIRJHeaderNorm(stripLeadingOrdinal(cell));
  if (spec.match.test(cell) || spec.match.test(norm) || spec.match.test(stripped)) return true;
  if (spec.ordinal) {
    const ordRe = new RegExp(`^\\s*${spec.ordinal}[\\.\\):\\-]?\\s*`, 'i');
    if (ordRe.test(cell) && spec.match.test(stripped)) return true;
  }
  return false;
};

const findHeaderLabelCell = (getMergedAwareCellText, matchRe, effectiveSheetCols, maxRows = 45, spec = null) => {
  const maxC = Math.max(20, effectiveSheetCols || 0);
  for (let r = 0; r < maxRows; r += 1) {
    for (let c = 0; c < maxC; c += 1) {
      const raw = String(getMergedAwareCellText(r, c) || '').trim();
      if (!raw) continue;
      if (spec ? labelMatchesSpec(raw, spec) : matchRe.test(raw) || matchRe.test(formXIRJHeaderNorm(raw))) {
        return { labelRow: r, labelCol: c };
      }
    }
  }
  return null;
};

const resolveValueColumnBesideLabel = (getMergedAwareCellText, labelRow, labelCol, effectiveSheetCols) => {
  const maxC = Math.max(20, effectiveSheetCols || 0);
  for (let c = labelCol + 1; c < Math.min(labelCol + 12, maxC); c += 1) {
    const v = String(getMergedAwareCellText(labelRow, c) || '').trim();
    if (!v || isPlaceholderCell(v)) return c + 1;
    if (!/name\s+and\s+address|age\s+or\s+date|identification|father|husband|contractor|establishment|principal|nature\s+and\s+location/i.test(v)) {
      return c;
    }
  }
  return FORM_XI_RJ_STACKED_VALUE_COL - 1;
};

function buildSheetTextBlob(workbook, sheetName) {
  const ws = workbook?.Sheets?.[sheetName];
  if (!ws) return '';
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  return rows
    .slice(0, 45)
    .map((row) => (Array.isArray(row) ? row : []).map((c) => String(c || '')).join(' '))
    .join(' ');
}

function pickWorkbookSheet(workbook, hints = {}) {
  const names = Array.isArray(workbook?.SheetNames) ? workbook.SheetNames : [];
  if (!names.length) return null;
  const preferred = String(hints.preferredSheetName || '').trim();
  if (preferred && workbook.Sheets?.[preferred]) return preferred;
  if (names.length === 1) return names[0];
  let best = names[0];
  let bestScore = -Infinity;
  for (const name of names) {
    const sheetText = buildSheetTextBlob(workbook, name);
    const blob = `${name} ${sheetText}`.toLowerCase();
    let score = 0;
    if (matchesFormXIHint(blob)) score += 120;
    if (/service\s+certificate/i.test(blob)) score += 90;
    if (/actual\s+no\.?\s+of\s+days\s+worked/i.test(blob)) score += 80;
    if (score > bestScore) {
      bestScore = score;
      best = name;
    }
  }
  return best;
}

function buildWorkbookMergedCellAccessor(workbook, hints = {}) {
  if (!workbook?.SheetNames?.length) return null;
  const sheetName = pickWorkbookSheet(workbook, hints);
  const worksheet = sheetName ? workbook.Sheets[sheetName] : null;
  if (!worksheet) return null;

  const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1:A1');
  const maxSheetCols = range.e && typeof range.e.c === 'number' ? range.e.c + 1 : 0;
  const merges = worksheet['!merges'] || [];
  let mergeMaxCol = 0;
  merges.forEach((m) => {
    if (m?.e && typeof m.e.c === 'number') mergeMaxCol = Math.max(mergeMaxCol, m.e.c + 1);
  });
  const effectiveSheetCols = Math.max(maxSheetCols, mergeMaxCol, 20);
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
        const topLeft = getRawCellText(m.s.r, m.s.c);
        if (topLeft) return topLeft;
      }
    }
    return '';
  };

  return { getMergedAwareCellText, effectiveSheetCols, sheetName };
}

export function buildFormXIRJHeaderFields(getMergedAwareCellText = null, effectiveSheetCols = 20) {
  return FORM_XI_RJ_HEADER_SPECS.map((spec) => {
    let coords = {};
    if (typeof getMergedAwareCellText === 'function' && spec.match) {
      const found =
        findHeaderLabelCell(getMergedAwareCellText, spec.match, effectiveSheetCols, 45, spec) || {};
      if (found.labelRow != null) {
        const valueCol = resolveValueColumnBesideLabel(
          getMergedAwareCellText,
          found.labelRow,
          found.labelCol,
          effectiveSheetCols
        );
        coords = {
          labelRow: found.labelRow,
          labelCol: found.labelCol,
          valueRow: found.labelRow,
          valueCol,
        };
      }
    }
    return buildTemplateField(spec, coords);
  });
}

export function resolveFormXIHeaderFieldLayout(parsed, workbook, hints = {}) {
  const formHeader = hints.formHeader || parsed?.formHeader || null;
  const item = hints.item || null;
  const fileName = hints.fileName || hints.formFileName || '';
  const sheetText = hints.sheetText || '';
  if (!isFormXIRajasthanServiceCertificateContext(formHeader, item, fileName, sheetText)) return null;

  const accessor = buildWorkbookMergedCellAccessor(workbook, hints);
  const getMergedAwareCellText = accessor?.getMergedAwareCellText ?? null;
  const effectiveSheetCols = accessor?.effectiveSheetCols ?? 20;
  const fields = buildFormXIRJHeaderFields(getMergedAwareCellText, effectiveSheetCols);
  const tableHeaders = resolveFormXIRJTableHeaders(parsed?.headers || hints.tableHeaders || null);

  return {
    formHeader: {
      ...(formHeader || {}),
      formXIHeaderFieldLayout: true,
      title: formHeader?.title || 'FORM XI',
      subtitle: formHeader?.subtitle || 'Service Certificate',
      reference: formHeader?.reference || 'See Rule 76',
      fields,
      fieldGroups: FORM_XI_RJ_FIELD_GROUPS,
    },
    headers: tableHeaders,
    sheetName: accessor?.sheetName || hints.preferredSheetName || null,
  };
}

const setHeaderField = (headerData, key, value) => {
  const out = headerData && typeof headerData === 'object' ? { ...headerData } : {};
  const text = String(value ?? '').trim();
  if (text) out[key] = text;
  return out;
};

export function applyFormXIRJAutofillFromSite(headerData, siteContext = {}, options = {}) {
  const { onlyEmpty = false } = options;
  const {
    contractorText = '',
    establishmentText = '',
    natureLocationText = '',
    principalEmployerText = '',
  } = siteContext;
  let out = { ...(headerData || {}) };
  const assign = (key, value) => {
    if (onlyEmpty && String(out[key] ?? '').trim()) return;
    out = setHeaderField(out, key, value);
  };
  assign('form_xi_rj_contractor', contractorText);
  assign('form_xi_rj_establishment', establishmentText);
  assign('form_xi_rj_nature_location', natureLocationText);
  assign('form_xi_rj_principal_employer', principalEmployerText);
  return out;
}

export function applyFormXIRJEmployeeToHeader(headerData, emp = {}, helpers = {}) {
  const {
    buildWorkmanNameAndAddress = null,
    buildAgeOrDob = null,
    buildIdentificationMarks = null,
    buildFatherOrHusbandName = null,
  } = helpers;
  let out = { ...(headerData || {}) };
  const assign = (key, value) => {
    const text = String(value ?? '').trim();
    if (text) out[key] = text;
  };
  if (typeof buildWorkmanNameAndAddress === 'function') {
    assign('form_xi_rj_workman', buildWorkmanNameAndAddress(emp));
  }
  if (typeof buildAgeOrDob === 'function') {
    assign('form_xi_rj_age_or_dob', buildAgeOrDob(emp));
  }
  if (typeof buildIdentificationMarks === 'function') {
    assign('form_xi_rj_identification_marks', buildIdentificationMarks(emp));
  }
  if (typeof buildFatherOrHusbandName === 'function') {
    assign('form_xi_rj_father_or_husband_name', buildFatherOrHusbandName(emp));
  } else {
    const guardian = String(formatWorkmanNameAndGuardian(emp) || '')
      .split(/\r?\n/)
      .slice(1)
      .join('\n')
      .trim();
    if (guardian) assign('form_xi_rj_father_or_husband_name', guardian);
  }
  return out;
}

const parseFlexibleDate = (raw) => {
  const s = String(raw ?? '').trim();
  if (!s) return null;
  const iso = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (iso) return new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3]));
  const dmy = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
  if (dmy) {
    const y = Number(dmy[3].length === 2 ? `20${dmy[3]}` : dmy[3]);
    return new Date(y, Number(dmy[2]) - 1, Number(dmy[1]));
  }
  const d = new Date(s);
  return Number.isNaN(d.getTime()) ? null : d;
};

const formatDateDisplay = (raw) => {
  const d = parseFlexibleDate(raw);
  if (!d) return String(raw ?? '').trim();
  return `${String(d.getDate()).padStart(2, '0')}-${String(d.getMonth() + 1).padStart(2, '0')}-${d.getFullYear()}`;
};

const unwrapEmployee = (empItem) => empItem?.Employee || empItem?.employee || empItem || {};

const pickEmployeeString = (emp, keys) => {
  for (const key of keys) {
    const v = emp?.[key];
    if (v != null && String(v).trim() !== '') return String(v).trim();
  }
  return '';
};

const readEmployeeJoinDate = (empItem) => {
  const emp = unwrapEmployee(empItem);
  return pickEmployeeString(emp, [
    'Dateofjoining',
    'Date of Joining',
    'Date_of_Joining',
    'DateofJoining',
    'DOJ',
    'JoiningDate',
    'Joining Date',
  ]);
};

const readEmployeeExitDate = (empItem) => {
  const emp = unwrapEmployee(empItem);
  return pickEmployeeString(emp, [
    'Dateofexit',
    'Date of Exit',
    'Date_of_Exit',
    'DateofExit',
    'LastWorkingDate',
    'Last Working Date',
    'DOE',
  ]);
};

const readEmployeeDesignation = (empItem) => {
  const emp = unwrapEmployee(empItem);
  return pickEmployeeString(emp, [
    'Designation',
    'DesignationName',
    'Designation Name',
    'JobTitle',
    'Job Title',
    'Nature_of_employment',
    'Nature of employment',
  ]);
};

export function readFormXIRJPayrollPaidDays(payrollRow) {
  if (!payrollRow || payrollRow.fetch_error) return '';
  const flat = flattenPayrollEarningColumns(payrollRow);
  return readPayrollScalar(flat, PAID_DAYS_KEYS, PAID_DAYS_PATTERNS);
}

export function resolveFormXIRJRateOfWage(payrollRow) {
  // Daily rate from monthly gross ÷ 26. Blank only when gross pay is missing.
  const grossRaw = readFormXIVMPPayrollGrossPay(payrollRow);
  const gross = Number(String(grossRaw || '').replace(/,/g, '').trim());
  if (!Number.isFinite(gross) || gross <= 0) return '';
  const rate = Math.round((gross / 26) * 100) / 100;
  return String(rate);
}

export function resolveFormXIRJDeduction(payrollRow) {
  const grossRaw = readFormXIVMPPayrollGrossPay(payrollRow);
  const netRaw = readFormXIVMPPayrollNetPay(payrollRow);
  const gross = Number(String(grossRaw || '').replace(/,/g, '').trim());
  const net = Number(String(netRaw || '').replace(/,/g, '').trim());
  if (!Number.isFinite(gross) || !Number.isFinite(net)) return '';
  const deduction = Math.round((gross - net) * 100) / 100;
  return deduction >= 0 ? String(deduction) : '';
}

export function getFormXIRJRowValueForHeader(row, header) {
  if (!row || !header) return '';
  if (typeof row === 'object' && !Array.isArray(row)) {
    if (row[header] != null && String(row[header]).trim() !== '') return String(row[header]).trim();
    const target = normalizeTableHeader(header);
    for (const [key, value] of Object.entries(row)) {
      if (normalizeTableHeader(key) === target) return String(value ?? '').trim();
    }
  }
  return '';
}

export function mapFormXIRJRowsFromEmployees(employees, headers, helpers = {}) {
  const hdrs = resolveFormXIRJTableHeaders(headers);
  const list = Array.isArray(employees) ? employees : [];
  const {
    sanitizeValue = (v) => String(v ?? '').trim(),
    resolvePayrollRow = null,
    rowIndexOffset = 0,
  } = helpers;

  return list.map((empItem, rowIndex) => {
    const emp = unwrapEmployee(empItem);
    const payrollRow =
      typeof resolvePayrollRow === 'function' ? resolvePayrollRow(empItem, rowIndex) : null;
    const row = {};

    hdrs.forEach((header, colIndex) => {
      if (isFormXIRJSkipAutofillHeader(header)) {
        row[header] = '';
        return;
      }
      if (colIndex === 0 || isFormXIRJSerialHeader(header)) {
        row[header] = rowIndexOffset + rowIndex + 1;
        return;
      }
      if (isFormXIRJPeriodFromHeader(header)) {
        row[header] = sanitizeValue(formatDateDisplay(readEmployeeJoinDate(empItem)));
        return;
      }
      if (isFormXIRJPeriodToHeader(header)) {
        row[header] = sanitizeValue(formatDateDisplay(readEmployeeExitDate(empItem)));
        return;
      }
      if (isFormXIRJDaysWorkedHeader(header)) {
        row[header] = sanitizeValue(readFormXIRJPayrollPaidDays(payrollRow));
        return;
      }
      if (isFormXIRJNatureOfWorkHeader(header)) {
        row[header] = sanitizeValue(readEmployeeDesignation(empItem));
        return;
      }
      if (isFormXIRJRateOfWageHeader(header)) {
        row[header] = sanitizeValue(resolveFormXIRJRateOfWage(payrollRow));
        return;
      }
      if (isFormXIRJGrossWagesHeader(header)) {
        row[header] = sanitizeValue(readFormXIVMPPayrollGrossPay(payrollRow));
        return;
      }
      if (isFormXIRJDeductionHeader(header)) {
        row[header] = sanitizeValue(resolveFormXIRJDeduction(payrollRow));
        return;
      }
      if (isFormXIRJNetWagesHeader(header)) {
        row[header] = sanitizeValue(readFormXIVMPPayrollNetPay(payrollRow));
        return;
      }
      row[header] = '';
    });

    return row;
  });
}

export function enrichFormXIRJPayrollRows(mappedData, employees, headers, helpers = {}) {
  const hdrs = resolveFormXIRJTableHeaders(headers);
  const rows = Array.isArray(mappedData) ? mappedData : [];
  const emps = Array.isArray(employees) ? employees : [];
  const {
    sanitizeValue = (v) => String(v ?? '').trim(),
    resolvePayrollRow = null,
    overwrite = false,
    rowIndexOffset = 0,
  } = helpers;
  let hits = 0;

  rows.forEach((row, idx) => {
    const empItem = emps[idx + rowIndexOffset] || emps[idx] || null;
    const payrollRow =
      typeof resolvePayrollRow === 'function'
        ? resolvePayrollRow(empItem, idx)
        : empItem
          ? resolveFormXIXMPPayrollRowForEmployee(empItem, [])
          : null;
    if (!payrollRow || payrollRow.fetch_error) return;

    const apply = (header, value) => {
      const text = sanitizeValue(value);
      if (!text) return;
      const cur = getFormXIRJRowValueForHeader(row, header);
      if (!overwrite && cur) return;
      row[header] = text;
      hits += 1;
    };

    const daysHeader = hdrs.find(isFormXIRJDaysWorkedHeader);
    const rateHeader = hdrs.find(isFormXIRJRateOfWageHeader);
    const grossHeader = hdrs.find(isFormXIRJGrossWagesHeader);
    const deductionHeader = hdrs.find(isFormXIRJDeductionHeader);
    const netHeader = hdrs.find(isFormXIRJNetWagesHeader);

    if (daysHeader) apply(daysHeader, readFormXIRJPayrollPaidDays(payrollRow));
    if (rateHeader) apply(rateHeader, resolveFormXIRJRateOfWage(payrollRow));
    if (grossHeader) apply(grossHeader, readFormXIVMPPayrollGrossPay(payrollRow));
    if (deductionHeader) apply(deductionHeader, resolveFormXIRJDeduction(payrollRow));
    if (netHeader) apply(netHeader, readFormXIVMPPayrollNetPay(payrollRow));
  });

  return hits;
}

export function rowHasMeaningfulFormXIRJExportData(row, headers) {
  if (!row || typeof row !== 'object') return false;
  if (String(row.__employeeLookupName ?? '').trim()) return true;
  if (String(row.__employeeLookupId ?? '').trim()) return true;
  const hdrs = resolveFormXIRJTableHeaders(headers);
  return hdrs.some((header) => {
    if (isFormXIRJRemarksHeader(header)) return false;
    return getFormXIRJRowValueForHeader(row, header) !== '';
  });
}

const buildFormXIRJExportPairs = (mappedData, employeesOverride, hdrs) => {
  const rows = Array.isArray(mappedData) ? mappedData : [];
  const emps = Array.isArray(employeesOverride) ? employeesOverride : [];
  const exportAllRoster = emps.length > 0 && emps.length === rows.length;
  const pairs = [];
  for (let i = 0; i < rows.length; i += 1) {
    const row = rows[i];
    if (!exportAllRoster && !rowHasMeaningfulFormXIRJExportData(row, hdrs)) continue;
    pairs.push({ row, emp: emps[i] ?? null, index: i });
  }
  return pairs;
};

const resolveStackedValueColumn = () => FORM_XI_RJ_STACKED_VALUE_COL;

const setCellValue = (worksheet, row, col, value) => {
  const text = String(value ?? '').trim();
  if (!text || row < 1 || col < 1) return;
  const cell = worksheet.getCell(row, col);
  cell.value = text;
  const multiline = text.includes('\n');
  cell.alignment = {
    ...(cell.alignment || {}),
    horizontal: 'left',
    vertical: multiline ? 'top' : (cell.alignment?.vertical || 'top'),
    wrapText: multiline,
    shrinkToFit: false,
  };
};

const resolveBesideLabelCol = (worksheet, row, labelCol, defaultValueCol = FORM_XI_RJ_STACKED_VALUE_COL) => {
  let targetCol = defaultValueCol;
  for (let c = labelCol + 1; c <= Math.min(labelCol + 10, 12); c += 1) {
    const cellStr = excelCellValueToString(worksheet.getCell(row, c)?.value).trim();
    if (!cellStr || isPlaceholderCell(cellStr)) {
      targetCol = c;
      break;
    }
    if (!/name\s+and\s+address|age\s+or\s+date|identification|father|husband/i.test(cellStr)) {
      targetCol = c;
      break;
    }
  }
  return targetCol;
};

const writeHeaderFieldsToWorksheet = (worksheet, headerFormData = {}, parsedFields = []) => {
  if (!worksheet) return;
  const defaultValueCol = resolveStackedValueColumn();
  const parsedByKey = new Map((parsedFields || []).map((f) => [f.key, f]));

  const writeBesideLabel = (row, labelCol, value) => {
    setCellValue(worksheet, row, resolveBesideLabelCol(worksheet, row, labelCol, defaultValueCol), value);
  };

  FORM_XI_RJ_HEADER_SPECS.forEach((spec) => {
    let val = headerFormData?.[spec.key];
    if ((val == null || String(val).trim() === '') && spec.key === 'form_xi_rj_contractor') {
      val = resolveHeaderFieldExportValue(headerFormData, spec);
    }
    if (val == null || String(val).trim() === '') return;
    const parsedField = parsedByKey.get(spec.key);
    if (parsedField?.labelRow != null) {
      const labelExcelRow = parsedField.labelRow + 1;
      const labelExcelCol = (parsedField.labelCol ?? 0) + 1;
      const targetRow = (parsedField.valueRow ?? parsedField.labelRow) + 1;
      const targetCol =
        parsedField.valueCol != null ? parsedField.valueCol + 1 : defaultValueCol;
      if (parsedField.valueCol != null) {
        setCellValue(worksheet, targetRow, targetCol, val);
      } else {
        writeBesideLabel(labelExcelRow, labelExcelCol, val);
      }
      return;
    }
    for (let r = 1; r <= 30; r += 1) {
      for (let c = 1; c <= 8; c += 1) {
        const raw = excelCellValueToString(worksheet.getCell(r, c)?.value).trim();
        if (!raw || !labelMatchesSpec(raw, spec)) continue;
        writeBesideLabel(r, c, val);
        return;
      }
    }
  });
};

const resolveTableHeaderMatchCol = (worksheet, headerRow, header) => {
  const target = normalizeTableHeader(header);
  for (let c = 1; c <= 12; c += 1) {
    const cellNorm = normalizeTableHeader(excelCellValueToString(worksheet.getCell(headerRow, c)?.value));
    if (!cellNorm) continue;
    if (cellNorm === target || cellNorm.includes(target) || target.includes(cellNorm)) return c;
    if (isFormXIRJSerialHeader(header) && /serial|s\.?\s*no|sl\.?\s*no/.test(cellNorm)) return c;
    if (isFormXIRJPeriodFromHeader(header) && /\bfrom\b/.test(cellNorm)) return c;
    if (isFormXIRJPeriodToHeader(header) && /\bto\b/.test(cellNorm) && !/\bfrom\b/.test(cellNorm)) return c;
    if (isFormXIRJDaysWorkedHeader(header) && /days/.test(cellNorm) && /worked/.test(cellNorm)) return c;
    if (isFormXIRJNatureOfWorkHeader(header) && /nature/.test(cellNorm) && /work/.test(cellNorm)) return c;
    if (isFormXIRJRateOfWageHeader(header) && /rate/.test(cellNorm) && /wage/.test(cellNorm)) return c;
    if (isFormXIRJGrossWagesHeader(header) && /earned/.test(cellNorm)) return c;
    if (isFormXIRJDeductionHeader(header) && /deduction/.test(cellNorm)) return c;
    if (isFormXIRJNetWagesHeader(header) && /paid/.test(cellNorm) && /wage/.test(cellNorm)) return c;
    if (isFormXIRJRemarksHeader(header) && /remark/.test(cellNorm)) return c;
  }
  return -1;
};

const resolveTableExportLayout = (worksheet, hdrs) => {
  if (!worksheet) return null;
  const exportHdrs = resolveFormXIRJTableHeaders(hdrs);
  for (let r = 20; r <= 30; r += 1) {
    const columnByHeader = exportHdrs.map((header) => resolveTableHeaderMatchCol(worksheet, r, header));
    const hits = columnByHeader.filter((col) => col > 0).length;
    if (hits < 4) continue;
    let dataStartRow = r + 1;
    for (let dr = r + 1; dr <= r + 4; dr += 1) {
      const indexHits = columnByHeader.filter((col, i) => {
        if (col <= 0) return false;
        const raw = excelCellValueToString(worksheet.getCell(dr, col)?.value).trim();
        const num = Number(raw);
        return Number.isFinite(num) && num === i + 1;
      }).length;
      if (indexHits >= 3) {
        dataStartRow = dr + 1;
        break;
      }
    }
    return { headerRow: r, dataStartRow, columnByHeader, hdrs: exportHdrs };
  }
  return null;
};

const writeTableRowToWorksheet = (worksheet, row, layout) => {
  if (!worksheet || !row || !layout) return;
  const { dataStartRow, columnByHeader, hdrs } = layout;
  hdrs.forEach((header, j) => {
    const col = columnByHeader[j];
    if (!col || col < 1) return;
    const val = getFormXIRJRowValueForHeader(row, header);
    // Always clear Rate of wages / Remarks / empty To when missing so template
    // sample values do not leak into Excel → PDF.
    const forceBlank =
      !val &&
      (isFormXIRJRateOfWageHeader(header) ||
        isFormXIRJRemarksHeader(header) ||
        isFormXIRJPeriodToHeader(header));
    if (!val && !forceBlank) return;
    const cell = worksheet.getCell(dataStartRow, col);
    if (!val) {
      cell.value = '';
      return;
    }
    setCellValue(worksheet, dataStartRow, col, val);
  });
};

/**
 * Full thin box borders on the Form XI RJ employment table
 * (Serial No. … Remarks header band + index row + data / blank body box).
 */
const applyFormXIRJTableFullBorders = (worksheet, layout, dataRowCount = 1) => {
  if (!worksheet || !layout) return;
  const cols = (layout.columnByHeader || []).filter((c) => Number(c) > 0);
  if (!cols.length) return;
  const colFrom = Math.min(...cols);
  const colTo = Math.max(...cols);

  let rowFrom = Math.max(1, Number(layout.headerRow) || 1);
  for (let r = Math.max(1, rowFrom - 5); r < rowFrom; r += 1) {
    let hit = false;
    for (let c = colFrom; c <= colTo; c += 1) {
      const raw = excelCellValueToString(worksheet.getCell(r, c)?.value);
      if (
        /serial\s*no|total\s+period|nature\s+of\s+work|rate\s+of\s+wages|remarks|days\s+worked/i.test(
          raw
        )
      ) {
        hit = true;
        break;
      }
    }
    if (hit) {
      rowFrom = r;
      break;
    }
  }

  const dataStart = Math.max(rowFrom, Number(layout.dataStartRow) || rowFrom + 1);
  let rowTo = dataStart + Math.max(1, Number(dataRowCount) || 1) - 1;
  // Keep any pre-drawn empty body boxes under the first data row.
  for (let r = rowTo + 1; r <= dataStart + 8; r += 1) {
    let anyBorder = false;
    for (let c = colFrom; c <= colTo; c += 1) {
      const b = worksheet.getCell(r, c)?.border;
      if (b?.top?.style || b?.bottom?.style || b?.left?.style || b?.right?.style) {
        anyBorder = true;
        break;
      }
    }
    if (!anyBorder) break;
    rowTo = r;
  }
  // Always leave at least one blank bordered row under data for the certificate box.
  rowTo = Math.max(rowTo, dataStart + Math.max(1, Number(dataRowCount) || 1));

  applyExcelJSFullBoxBordersToRange(worksheet, {
    rowFrom,
    rowTo,
    colFrom,
    colTo,
    borderStyle: 'thin',
  });
};

export async function buildFormXIRJWorkbookWithTemplateStyles({
  templateArrayBuffer,
  mappedData,
  headersToUse,
  parsedFormHeader,
  formFileName,
  headerFormData,
}) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(templateArrayBuffer);
  const worksheet = workbook.worksheets[0];
  if (!worksheet) throw new Error('Template worksheet not found.');

  const hdrs = resolveFormXIRJTableHeaders(headersToUse);
  const sourceRows = (Array.isArray(mappedData) ? mappedData : []).filter((row) =>
    rowHasMeaningfulFormXIRJExportData(row, hdrs)
  );
  const employeeRow = sourceRows.length >= 1 ? sourceRows[0] : null;
  const parsedFields = parsedFormHeader?.fields || [];

  writeHeaderFieldsToWorksheet(worksheet, headerFormData, parsedFields);
  const tableLayout = resolveTableExportLayout(worksheet, hdrs);
  if (employeeRow && tableLayout) {
    writeTableRowToWorksheet(worksheet, employeeRow, tableLayout);
  }
  if (tableLayout) {
    applyFormXIRJTableFullBorders(worksheet, tableLayout, 1);
  }

  const out = await workbook.xlsx.writeBuffer();
  const fileName =
    formFileName ||
    parsedFormHeader?.title?.replace(/[^a-zA-Z0-9]+/g, '_') ||
    `Form_XI_RJ_${Date.now()}.xlsx`;
  const blob = new Blob([out], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  return { blob, fileName };
}

const FORM_XI_RJ_EMPLOYEE_HEADER_KEYS = new Set(
  FORM_XI_RJ_WORKMAN_HEADER_SPECS.map((spec) => spec.key)
);
const FORM_XI_RJ_FAST_ZIP_BATCH = 12;

const sanitizeDownloadBaseName = (name, fallback) => {
  const base = String(name || fallback || 'Employee')
    .replace(/[<>:"/\\|?*]+/g, '_')
    .replace(/\s+/g, '_')
    .replace(/_+/g, '_')
    .replace(/^_|_$/g, '');
  return base || 'Employee';
};

const allocateUniqueFileName = (baseName, usedNames, ext = 'xlsx') => {
  const root = sanitizeDownloadBaseName(baseName, 'Employee');
  const count = usedNames.get(root) || 0;
  usedNames.set(root, count + 1);
  if (count === 0) return `${root}.${ext}`;
  return `${root}_${count + 1}.${ext}`;
};

const resolveEmployeeDownloadBaseName = (row, hdrs, empItem, fallbackIndex) => {
  const fromEmp = empItem
    ? String(formatWorkmanNameAndGuardian(unwrapEmployee(empItem)) || '')
        .split(/\r?\n/)[0]
        .trim()
    : '';
  if (fromEmp) return fromEmp;
  const fromRow = String(row?.__employeeLookupName ?? '').trim();
  if (fromRow) return fromRow.split(/\r?\n/)[0].trim();
  const serialHeader = hdrs.find(isFormXIRJSerialHeader);
  const serial = serialHeader ? getFormXIRJRowValueForHeader(row, serialHeader) : '';
  return serial ? `Employee_${serial}` : `Employee_${fallbackIndex + 1}`;
};

const formXIRJSanitizeExportText = (value) =>
  String(value ?? '')
    .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '')
    .trim();

const formXIRJEscapeXml = (value) =>
  String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');

const formXIRJColToLetter = (col) => {
  let result = '';
  let n = col;
  while (n > 0) {
    const rem = (n - 1) % 26;
    result = String.fromCharCode(65 + rem) + result;
    n = Math.floor((n - 1) / 26);
  }
  return result;
};

const formXIRJToCellRef = (row, col) => `${formXIRJColToLetter(col)}${row}`;

const formXIRJUpsertInlineStrCell = (sheetXml, cellRef, value) => {
  const text = formXIRJEscapeXml(formXIRJSanitizeExportText(value));
  const cellRe = new RegExp(`<c\\s+r="${cellRef}"([^>/]*)(?:/>|>([\\s\\S]*?)</c>)`, 'i');
  const existing = sheetXml.match(cellRe);
  let styleAttr = '';
  if (existing) {
    const styleMatch = String(existing[1] || '').match(/\ss="(\d+)"/i);
    if (styleMatch) styleAttr = ` s="${styleMatch[1]}"`;
  }
  const cellXml = text
    ? `<c r="${cellRef}"${styleAttr} t="inlineStr"><is><t xml:space="preserve">${text}</t></is></c>`
    : `<c r="${cellRef}"${styleAttr}/>`;
  if (existing) {
    return sheetXml.replace(cellRe, cellXml);
  }
  const rowNum = cellRef.replace(/^[A-Z]+/i, '');
  const rowRe = new RegExp(`(<row\\s+r="${rowNum}"[^>]*>)([\\s\\S]*?)(</row>)`, 'i');
  if (!rowRe.test(sheetXml)) return sheetXml;
  return sheetXml.replace(rowRe, `$1$2${cellXml}$3`);
};

const resolveFormXIRJWorksheetEntry = (zipFiles) =>
  Object.keys(zipFiles)
    .filter((name) => /^xl\/worksheets\/sheet\d+\.xml$/i.test(name))
    .sort((a, b) => a.localeCompare(b, undefined, { numeric: true }))[0] || null;

const resolveFormXIRJEmployeeHeaderValueCol = (worksheet, labelRow, labelCol, parsedField) => {
  if (parsedField?.valueCol != null) return parsedField.valueCol + 1;
  return resolveBesideLabelCol(worksheet, labelRow, labelCol, FORM_XI_RJ_STACKED_VALUE_COL);
};

function resolveFormXIRJFastExportPositions(worksheet, parsedFormHeader, hdrs) {
  const employeeHeaderPositions = [];
  const seenHeaderKeys = new Set();
  const parsedFields = Array.isArray(parsedFormHeader?.fields) ? parsedFormHeader.fields : [];
  const parsedByKey = new Map(parsedFields.map((f) => [f.key, f]));

  const pushEmployeeHeader = (key, row, col) => {
    if (!key || row == null || col == null || seenHeaderKeys.has(key)) return;
    seenHeaderKeys.add(key);
    employeeHeaderPositions.push({ key, row, col, cellRef: formXIRJToCellRef(row, col) });
  };

  FORM_XI_RJ_WORKMAN_HEADER_SPECS.forEach((spec) => {
    const parsedField = parsedByKey.get(spec.key);
    if (parsedField?.labelRow != null) {
      const labelExcelRow = parsedField.labelRow + 1;
      const labelExcelCol = (parsedField.labelCol ?? 0) + 1;
      const targetRow = (parsedField.valueRow ?? parsedField.labelRow) + 1;
      const targetCol = resolveFormXIRJEmployeeHeaderValueCol(
        worksheet,
        labelExcelRow,
        labelExcelCol,
        parsedField
      );
      pushEmployeeHeader(spec.key, targetRow, targetCol);
      return;
    }
    if (!worksheet) return;
    for (let r = 1; r <= 30; r += 1) {
      for (let c = 1; c <= 8; c += 1) {
        const raw = excelCellValueToString(worksheet.getCell(r, c)?.value).trim();
        if (!raw || !labelMatchesSpec(raw, spec)) continue;
        pushEmployeeHeader(spec.key, r, resolveBesideLabelCol(worksheet, r, c));
        return;
      }
    }
  });

  const tableLayout = resolveTableExportLayout(worksheet, hdrs);
  const tablePositions = [];
  if (tableLayout) {
    tableLayout.hdrs.forEach((header, j) => {
      const col = tableLayout.columnByHeader[j];
      if (!col || col < 1) return;
      tablePositions.push({
        header,
        row: tableLayout.dataStartRow,
        col,
        cellRef: formXIRJToCellRef(tableLayout.dataStartRow, col),
      });
    });
  }

  return { employeeHeaderPositions, tablePositions };
}

const clearFormXIRJPerEmployeeValueCells = (worksheet, positions) => {
  if (!worksheet || !positions) return;
  const all = [...(positions.employeeHeaderPositions || []), ...(positions.tablePositions || [])];
  all.forEach((pos) => {
    if (pos?.row != null && pos?.col != null) {
      worksheet.getCell(pos.row, pos.col).value = '';
    }
  });
};

const patchFormXIRJFastSheetXml = (baseSheetXml, mergedHeaderData, exportRow, positions) => {
  let sheetXml = baseSheetXml;
  (positions.employeeHeaderPositions || []).forEach((pos) => {
    const value = mergedHeaderData?.[pos.key];
    if (value != null && String(value).trim() !== '') {
      sheetXml = formXIRJUpsertInlineStrCell(sheetXml, pos.cellRef, value);
    }
  });
  (positions.tablePositions || []).forEach((pos) => {
    const value = getFormXIRJRowValueForHeader(exportRow, pos.header);
    if (value !== '') {
      sheetXml = formXIRJUpsertInlineStrCell(sheetXml, pos.cellRef, value);
    }
  });
  return sheetXml;
};

const buildFormXIRJFastXlsxBytes = async (fastTemplate, mergedHeaderData, exportRow) => {
  const { sheetEntry, baseSheetXml, staticFiles, positions } = fastTemplate;
  const sheetXml = patchFormXIRJFastSheetXml(baseSheetXml, mergedHeaderData, exportRow, positions);
  const entryZip = new JSZip();
  Object.entries(staticFiles).forEach(([path, data]) => {
    entryZip.file(path, data);
  });
  entryZip.file(sheetEntry, sheetXml);
  return entryZip.generateAsync({ type: 'uint8array', compression: 'STORE' });
};

async function prepareFormXIRJFastZipTemplate({
  templateArrayBuffer,
  parsedFormHeader,
  headerFormData,
  headersToUse,
}) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(templateArrayBuffer);
  const worksheet = workbook.worksheets[0];
  if (!worksheet) throw new Error('Template worksheet not found.');

  const hdrs = resolveFormXIRJTableHeaders(headersToUse);
  const staticHeaderData = { ...(headerFormData || {}) };
  FORM_XI_RJ_EMPLOYEE_HEADER_KEYS.forEach((key) => {
    delete staticHeaderData[key];
  });
  writeHeaderFieldsToWorksheet(worksheet, staticHeaderData, parsedFormHeader?.fields || []);

  const positions = resolveFormXIRJFastExportPositions(worksheet, parsedFormHeader, hdrs);
  clearFormXIRJPerEmployeeValueCells(worksheet, positions);
  const tableLayout = resolveTableExportLayout(worksheet, hdrs);
  if (tableLayout) {
    applyFormXIRJTableFullBorders(worksheet, tableLayout, 1);
  }

  const preparedBuffer = await workbook.xlsx.writeBuffer();
  const templateZip = await JSZip.loadAsync(preparedBuffer);
  const sheetEntry = resolveFormXIRJWorksheetEntry(templateZip.files);
  if (!sheetEntry) throw new Error('Template worksheet XML not found.');
  const baseSheetXml = await templateZip.file(sheetEntry).async('string');
  const staticFiles = {};
  await Promise.all(
    Object.keys(templateZip.files).map(async (path) => {
      const file = templateZip.files[path];
      if (!file || file.dir || path === sheetEntry) return;
      staticFiles[path] = await file.async('uint8array');
    })
  );
  return { sheetEntry, baseSheetXml, staticFiles, positions };
}

async function buildFormXIRJFastZipDownload({
  exportPairs,
  fastTemplate,
  baseHeaderData,
  formFileName,
  parsedFormHeader,
  hdrs,
  employeeHeaderHelpers = null,
}) {
  const zip = new JSZip();
  const usedNames = new Map();
  for (let i = 0; i < exportPairs.length; i += FORM_XI_RJ_FAST_ZIP_BATCH) {
    const batch = exportPairs.slice(i, i + FORM_XI_RJ_FAST_ZIP_BATCH);
    const batchBytes = await Promise.all(
      batch.map(async (pair, batchIndex) => {
        const index = i + batchIndex;
        let mergedHeaderData = { ...(baseHeaderData || {}) };
        if (pair?.emp && employeeHeaderHelpers) {
          mergedHeaderData = applyFormXIRJEmployeeToHeader(
            mergedHeaderData,
            pair.emp,
            employeeHeaderHelpers
          );
        }
        const xlsxBytes = await buildFormXIRJFastXlsxBytes(fastTemplate, mergedHeaderData, pair.row);
        return { xlsxBytes, pair, index };
      })
    );
    batchBytes.forEach((entry) => {
      if (!entry) return;
      const { xlsxBytes, pair, index } = entry;
      const baseName = resolveEmployeeDownloadBaseName(pair.row, hdrs, pair.emp, index);
      zip.file(allocateUniqueFileName(baseName, usedNames), xlsxBytes);
    });
    if (i > 0) {
      await new Promise((resolve) => setTimeout(resolve, 0));
    }
  }
  const zipBase = String(formFileName || parsedFormHeader?.title || 'Form_XI_RJ')
    .replace(/\.xlsx?$/i, '')
    .replace(/[^a-zA-Z0-9._-]+/g, '_');
  return {
    blob: await zip.generateAsync({ type: 'blob', compression: 'STORE' }),
    fileName: `${zipBase}_Employees.zip`,
  };
}

export function triggerFormXIRJZipDownload(blob, fileName) {
  if (!blob || !fileName) return;
  const downloadUrl = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = downloadUrl;
  link.download = fileName;
  link.style.display = 'none';
  link.rel = 'noopener';
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  setTimeout(() => {
    try {
      URL.revokeObjectURL(downloadUrl);
    } catch (_) {
      /* ignore */
    }
  }, 60_000);
}

export async function buildFormXIRajasthanPerEmployeeDownload({
  templateArrayBuffer,
  mappedData,
  headersToUse,
  parsedFormHeader,
  formFileName,
  headerFormData,
  employeesOverride = null,
  employeeHeaderHelpers = null,
}) {
  const hdrs = resolveFormXIRJTableHeaders(headersToUse);
  const exportPairs = buildFormXIRJExportPairs(mappedData, employeesOverride, hdrs);
  const workbookArgs = {
    templateArrayBuffer,
    headersToUse: hdrs,
    parsedFormHeader,
    formFileName,
  };

  if (exportPairs.length <= 1) {
    const pair = exportPairs[0] || null;
    let headerForSingle = { ...(headerFormData || {}) };
    if (pair?.emp && employeeHeaderHelpers) {
      headerForSingle = applyFormXIRJEmployeeToHeader(headerForSingle, pair.emp, employeeHeaderHelpers);
    }
    return buildFormXIRJWorkbookWithTemplateStyles({
      ...workbookArgs,
      mappedData: pair ? [pair.row] : [],
      headerFormData: headerForSingle,
    });
  }

  try {
    const fastTemplate = await prepareFormXIRJFastZipTemplate({
      templateArrayBuffer,
      parsedFormHeader,
      headerFormData,
      headersToUse: hdrs,
    });
    const hasPatchTargets =
      (fastTemplate.positions?.employeeHeaderPositions?.length || 0) > 0 ||
      (fastTemplate.positions?.tablePositions?.length || 0) > 0;
    if (hasPatchTargets) {
      return buildFormXIRJFastZipDownload({
        exportPairs,
        fastTemplate,
        baseHeaderData: headerFormData || {},
        formFileName,
        parsedFormHeader,
        hdrs,
        employeeHeaderHelpers,
      });
    }
  } catch (_err) {
    // Fall back to ExcelJS-per-file path if fast template prep fails.
  }

  const zip = new JSZip();
  const usedNames = new Map();
  for (let i = 0; i < exportPairs.length; i += 1) {
    const { row, emp, index } = exportPairs[i];
    let perEmployeeHeader = { ...(headerFormData || {}) };
    if (emp && employeeHeaderHelpers) {
      perEmployeeHeader = applyFormXIRJEmployeeToHeader(perEmployeeHeader, emp, employeeHeaderHelpers);
    }
    const { blob } = await buildFormXIRJWorkbookWithTemplateStyles({
      ...workbookArgs,
      mappedData: [row],
      headerFormData: perEmployeeHeader,
    });
    const baseName = resolveEmployeeDownloadBaseName(row, hdrs, emp, index);
    zip.file(allocateUniqueFileName(baseName, usedNames), blob);
    if (i > 0 && i % 10 === 0) {
      await new Promise((resolve) => setTimeout(resolve, 0));
    }
  }

  const zipBase = String(formFileName || parsedFormHeader?.title || 'Form_XI_RJ')
    .replace(/\.xlsx?$/i, '')
    .replace(/[^a-zA-Z0-9._-]+/g, '_');
  return {
    blob: await zip.generateAsync({ type: 'blob', compression: 'STORE' }),
    fileName: `${zipBase}_Employees.zip`,
  };
}

export { buildFormXIXMPPayrollRowResolver, resolveFormXIXMPPayrollRowForEmployee };
