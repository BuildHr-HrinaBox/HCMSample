import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { flattenPayrollEarningColumns, readPayrollNetPayForStatutory } from '../../utils/payrollEarnings';
import { formatWorkmanNameAndGuardian } from './formXIXAPWageSlip';
import { readFormXIVMPPayrollGrossPay } from './formXIVMPEmploymentCard';

/** Rajasthan CLRA Form XV — Wage Slip [Rule 77(2)(b)]. */

export const FORM_XV_RJ_NIL = 'Nil';

export const FORM_XV_RJ_TABLE_HEADERS = [
  'No. of days worked',
  'Rate of daily wages/piece rate',
  'No. of units worked in case of piece rate workers',
  'Date on which overtime worked',
  'Overtime hours and amount of overtime wages',
  'Gross wages payable',
  'Deductions, if any',
  'Actual wages paid',
  'Signature of the contractor or his representative',
];

export const FORM_XV_RJ_HEADER_SPECS = [
  {
    key: 'form_xv_rj_contractor',
    label: 'Name and address of contractor',
    match: /name\s+and\s+address\s+of\s+(?:the\s+)?contractor/i,
  },
  {
    key: 'form_xv_rj_nature_location',
    label: 'Name and location of work',
    match: /(?:name|nature)\s+and\s+location\s+of\s+work/i,
  },
  {
    key: 'form_xv_rj_establishment',
    label: 'Name and address of establishment in/under which contract is carried on',
    match:
      /name\s+and\s+address\s+of\s+establishment[\s\S]{0,80}?contract\s+is\s+carried\s+on|establishment\s+in\s*\/\s*under\s+which/i,
  },
  {
    key: 'form_xv_rj_principal_employer',
    label: 'Name and address of principal employer',
    match: /name\s+and\s+address\s+of\s+(?:the\s+)?principal\s+employer/i,
  },
  {
    key: 'form_xv_rj_workman',
    label: "Name and Father's name of the workman",
    match: /name\s+and\s+father.*(name\s+of\s+the\s+)?workman|father.?s?\s+name\s+of\s+the\s+workman/i,
  },
  {
    key: 'form_xv_rj_sex_token',
    label: 'Sex and identification token/ticket No.',
    match: /sex\s+and\s+identification|identification\s+token|ticket\s+no/i,
  },
  {
    key: 'form_xv_rj_period',
    label: 'For the week/fortnight/month',
    match: /for\s+the\s+week|fortnight|week\/fortnight\/month/i,
  },
];

const norm = (s) =>
  String(s || '')
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();

/** Rajasthan Form XV wage slip — not TN/MP/GJ Service Certificate Form XV. */
export function isFormXVRJWageSlipContext(formHeader, rowItem, fileName, sheetText = '') {
  const blob = [
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
  ]
    .filter((x) => x != null && String(x).trim() !== '')
    .join(' ')
    .toLowerCase();

  if (/form[\s._-]*xvi(?![a-z])/i.test(blob) && !/form[\s._-]*xv(?![a-z])/i.test(blob)) {
    return false;
  }
  if (!/form[\s._-]*xv(?![a-z])/i.test(blob) && !/\bxv[\s._-]*rj\b/i.test(blob)) {
    return false;
  }
  // Never treat as Form XI Service Certificate.
  if (/form[\s._-]*xi(?![vix])/i.test(blob) && !/form[\s._-]*xv(?![a-z])/i.test(blob)) {
    return false;
  }
  if (/service\s+certificate/i.test(blob) && !/wage\s*slip|wages?\s+slip/i.test(blob)) {
    return false;
  }

  if (/form[\s._-]*xv[\s._-]*rj|\bxv[\s._-]*rj\b/i.test(blob)) return true;
  if (/rajasthan/i.test(blob) && /form[\s._-]*xv(?![a-z])/i.test(blob)) return true;
  if (
    /rajasthan/i.test(blob) &&
    /wage\s*slip|wages?\s+slip|rule\s*77\s*\(\s*2\s*\)\s*\(\s*b\s*\)/i.test(blob)
  ) {
    return true;
  }
  return false;
}

export function isFormXVRJHeaderFieldLayoutFormHeader(formHeader) {
  return !!formHeader?.formXVRJHeaderFieldLayout;
}

export function isFormXVRJTableLayoutFormHeader(formHeader) {
  return !!formHeader?.formXVRJTableLayout;
}

export function resolveFormXVRJTableHeaders(tableHeaders) {
  if (Array.isArray(tableHeaders) && tableHeaders.length >= 6) {
    const joined = tableHeaders.map((h) => norm(h)).join(' ');
    if (
      /no\.?\s*of\s+days\s+worked/.test(joined) &&
      (/gross\s+wages/.test(joined) || /actual\s+wages\s+paid/.test(joined))
    ) {
      // Prefer template order but keep meaningful non-blank labels from workbook when aligned.
      const cleaned = tableHeaders
        .map((h) => String(h || '').replace(/\r?\n/g, ' ').replace(/\s+/g, ' ').trim())
        .filter((h) => h && !/^column\s+/i.test(h));
      if (cleaned.length >= 8) return cleaned.slice(0, 9);
    }
  }
  return [...FORM_XV_RJ_TABLE_HEADERS];
}

function pickExistingFieldValue(fields, matchRe) {
  if (!Array.isArray(fields)) return '';
  for (let i = 0; i < fields.length; i += 1) {
    const f = fields[i];
    const label = String(f?.label || '');
    if (!matchRe.test(label) && !matchRe.test(norm(label))) continue;
    const v = String(f?.value ?? '').trim();
    if (v) return v;
  }
  return '';
}

function buildFormXVRJHeaderFields(existingFields) {
  return FORM_XV_RJ_HEADER_SPECS.map((spec) => ({
    key: spec.key,
    label: spec.label,
    value: pickExistingFieldValue(existingFields, spec.match) || '',
    fieldType: /workman|contractor|establishment|employer|location/i.test(spec.label)
      ? 'textarea'
      : 'text',
  }));
}

/**
 * Keep Excel wage-slip table + canonical header band for Form XV RJ.
 * Drops duplicate generic/numbered fields that the spreadsheet parser collects.
 */
export function resolveFormXVRJWageSlipLayout(parsed, _workbook, hints = {}) {
  const formHeader = hints.formHeader || parsed?.formHeader || null;
  const item = hints.item || null;
  const fileName = hints.fileName || hints.formFileName || '';
  const sheetText = hints.sheetText || '';
  if (!isFormXVRJWageSlipContext(formHeader, item, fileName, sheetText)) return null;

  const existingFields = Array.isArray(formHeader?.fields)
    ? formHeader.fields
    : Array.isArray(parsed?.formHeader?.fields)
      ? parsed.formHeader.fields
      : [];
  const fields = buildFormXVRJHeaderFields(existingFields);
  const headers = resolveFormXVRJTableHeaders(parsed?.headers || []);

  return {
    formHeader: {
      ...(formHeader || {}),
      title: 'FORM XV',
      subtitle: 'Wages Slip',
      reference: formHeader?.reference || '[See Rule 77(2)(b)]',
      formXVRJHeaderFieldLayout: true,
      formXVRJTableLayout: true,
      formXIXAPHeaderFieldLayout: false,
      formXIHeaderFieldLayout: false,
      formXIVMPHeaderFieldLayout: false,
      fields,
      textRows: [],
    },
    headers,
    tableData: Array.isArray(parsed?.tableData) ? parsed.tableData : [],
    headerRowIndex: parsed?.headerRowIndex ?? -1,
    dataStartIndex: parsed?.dataStartIndex ?? 0,
    tableStartCol: parsed?.tableStartCol ?? 0,
  };
}

export function isFormXVRJDaysWorkedHeader(h) {
  const s = norm(h);
  return /no\.?\s*of\s+days\s+worked/.test(s) || (s.includes('days') && s.includes('worked') && !/over/.test(s));
}

export function isFormXVRJDailyRateHeader(h) {
  const s = norm(h);
  return (
    /rate\s+of\s+daily\s+wages/.test(s) ||
    (s.includes('rate') && (s.includes('daily') || s.includes('piece')) && !/units/.test(s))
  );
}

export function isFormXVRJUnitsWorkedHeader(h) {
  const s = norm(h);
  return /no\.?\s*of\s+units\s+worked/.test(s) || (s.includes('units') && s.includes('piece'));
}

export function isFormXVRJOvertimeDateHeader(h) {
  const s = norm(h);
  if (/payment|paid|hours|amount|rate|wage/.test(s) && !/date/.test(s)) return false;
  return (
    /dates?\s+on\s+which\s+over[\s-]*time\s+worked/.test(s) ||
    (s.includes('date') && s.includes('overtime') && s.includes('worked'))
  );
}

export function isFormXVRJOvertimeHoursAmountHeader(h) {
  const s = norm(h);
  return (
    /over[\s-]*time\s+hours/.test(s) ||
    (s.includes('overtime') && s.includes('amount') && s.includes('wage')) ||
    /amount\s+of\s+over[\s-]*time\s+wages/.test(s)
  );
}

export function isFormXVRJGrossWagesHeader(h) {
  const s = norm(h);
  return /gross\s+wages?\s+payable/.test(s) || (s.includes('gross') && s.includes('wage'));
}

export function isFormXVRJDeductionsHeader(h) {
  const s = norm(h);
  return /deductions?\s*,?\s*if\s+any/.test(s) || (s.includes('deduction') && !/gross|net|actual/.test(s));
}

export function isFormXVRJActualWagesPaidHeader(h) {
  const s = norm(h);
  return /actual\s+wages?\s+paid/.test(s) || (s.includes('actual') && s.includes('paid'));
}

export function isFormXVRJSignatureHeader(h) {
  const s = norm(h);
  return s.includes('signature') && (s.includes('contractor') || s.includes('representative'));
}

export function isFormXVRJSkipPeopleAutofillHeader(h) {
  return (
    isFormXVRJDaysWorkedHeader(h) ||
    isFormXVRJDailyRateHeader(h) ||
    isFormXVRJUnitsWorkedHeader(h) ||
    isFormXVRJOvertimeDateHeader(h) ||
    isFormXVRJOvertimeHoursAmountHeader(h) ||
    isFormXVRJGrossWagesHeader(h) ||
    isFormXVRJDeductionsHeader(h) ||
    isFormXVRJActualWagesPaidHeader(h) ||
    isFormXVRJSignatureHeader(h)
  );
}

function readFormXVRJPaidDays(payrollRow) {
  if (!payrollRow || payrollRow.fetch_error) return '';
  let p = payrollRow;
  if (typeof payrollRow?.payroll_payload === 'string') {
    try {
      p = { ...payrollRow, ...JSON.parse(payrollRow.payroll_payload) };
    } catch (_) {
      p = payrollRow;
    }
  } else if (payrollRow?.payroll_payload && typeof payrollRow.payroll_payload === 'object') {
    p = { ...payrollRow, ...payrollRow.payroll_payload };
  }
  const flat = flattenPayrollEarningColumns(payrollRow);
  const raw =
    flat.paid_days ??
    flat['paid_days'] ??
    flat['Paid Days'] ??
    flat.paidDays ??
    flat.days_worked ??
    flat['days_worked'] ??
    p.paid_days ??
    p['paid_days'] ??
    p['Paid Days'] ??
    p.paidDays ??
    p.days_worked ??
    payrollRow.paid_days ??
    payrollRow['paid_days'] ??
    '';
  if (raw === '' || raw == null) return '';
  const n = Number(String(raw).replace(/,/g, '').trim());
  return Number.isFinite(n) && n >= 0 ? n : '';
}

function parseMoney(raw) {
  if (raw === '' || raw == null) return null;
  const n = Number(String(raw).replace(/,/g, '').trim());
  return Number.isFinite(n) ? n : null;
}

function formatMoney(n) {
  if (!Number.isFinite(n)) return '';
  if (Number.isInteger(n)) return String(n);
  return String(Math.round(n * 100) / 100);
}

export function resolveFormXVRJGrossPay(payrollRow) {
  const gross = readFormXIVMPPayrollGrossPay(payrollRow);
  return gross === '' || gross == null ? '' : String(gross);
}

export function resolveFormXVRJNetPay(payrollRow) {
  const net = readPayrollNetPayForStatutory(payrollRow);
  return net === '' || net == null ? '' : String(net);
}

export function resolveFormXVRJDailyRate(payrollRow) {
  const grossN = parseMoney(resolveFormXVRJGrossPay(payrollRow));
  if (grossN == null) return '';
  return formatMoney(grossN / 26);
}

export function resolveFormXVRJDeductions(payrollRow) {
  const grossN = parseMoney(resolveFormXVRJGrossPay(payrollRow));
  const netN = parseMoney(resolveFormXVRJNetPay(payrollRow));
  if (grossN == null || netN == null) return '';
  const ded = grossN - netN;
  return formatMoney(ded >= 0 ? ded : 0);
}

/**
 * Fill Form XV RJ wage-slip table columns from payroll.
 * days → paid_days; rate → gross/26; OT date/hours → Nil;
 * gross → gross_pay; deductions → gross-net; actual → net_pay.
 */
export function applyFormXVRJWageSlipAutofillToRow(
  row,
  _emp,
  headers,
  payrollRow = null,
  { overwrite = true, sanitizeValue = (v) => String(v ?? '').trim() } = {}
) {
  if (!row || !Array.isArray(headers) || headers.length === 0) return false;
  const cellEmpty = (header) => {
    const v = String(row[header] ?? '').trim();
    return !v || /^n\/?a$/i.test(v) || v === '-' || v === '—';
  };
  const setCell = (header, value) => {
    if (!header || value == null || value === '') return;
    if (!overwrite && !cellEmpty(header)) return;
    row[header] = sanitizeValue(value);
  };

  let changed = false;
  const paidDays = readFormXVRJPaidDays(payrollRow);
  const dailyRate = resolveFormXVRJDailyRate(payrollRow);
  const grossPay = resolveFormXVRJGrossPay(payrollRow);
  const netPay = resolveFormXVRJNetPay(payrollRow);
  const deductions = resolveFormXVRJDeductions(payrollRow);

  headers.forEach((header) => {
    if (isFormXVRJDaysWorkedHeader(header) && paidDays !== '') {
      setCell(header, String(paidDays));
      changed = true;
      return;
    }
    if (isFormXVRJDailyRateHeader(header) && dailyRate !== '') {
      setCell(header, dailyRate);
      changed = true;
      return;
    }
    if (isFormXVRJOvertimeDateHeader(header)) {
      setCell(header, FORM_XV_RJ_NIL);
      changed = true;
      return;
    }
    if (isFormXVRJOvertimeHoursAmountHeader(header)) {
      setCell(header, FORM_XV_RJ_NIL);
      changed = true;
      return;
    }
    if (isFormXVRJGrossWagesHeader(header) && grossPay !== '') {
      setCell(header, grossPay);
      changed = true;
      return;
    }
    if (isFormXVRJDeductionsHeader(header) && deductions !== '') {
      setCell(header, deductions);
      changed = true;
      return;
    }
    if (isFormXVRJActualWagesPaidHeader(header) && netPay !== '') {
      setCell(header, netPay);
      changed = true;
    }
  });

  return changed;
}

const setHeaderField = (headerData, key, value) => {
  const out = { ...(headerData || {}) };
  const text = String(value ?? '').trim();
  if (!text || !key) return out;
  const cur = String(out[key] ?? '').trim();
  if (!cur || /^enter\b/i.test(cur)) out[key] = text;
  return out;
};

export function applyFormXVRJAutofillFromSite(headerData, siteContext = {}, options = {}) {
  const { onlyEmpty = false } = options;
  const {
    contractorText = '',
    establishmentText = '',
    natureLocationText = '',
    principalEmployerText = '',
    periodText = '',
  } = siteContext;
  let out = { ...(headerData || {}) };
  const assign = (key, value) => {
    if (onlyEmpty && String(out[key] ?? '').trim()) return;
    out = setHeaderField(out, key, value);
  };
  assign('form_xv_rj_contractor', contractorText);
  assign('form_xv_rj_establishment', establishmentText);
  assign('form_xv_rj_nature_location', natureLocationText);
  assign('form_xv_rj_principal_employer', principalEmployerText);
  assign('form_xv_rj_period', periodText);
  return out;
}

export function resolveFormXVRJEmployeeToken(emp = {}) {
  return String(
    emp.EmployeeID ||
      emp['EmployeeID'] ||
      emp['Employee ID'] ||
      emp.Employee_Number ||
      emp['Employee Number'] ||
      emp.EmployeeCode ||
      emp['Employee Code'] ||
      emp.Zoho_ID ||
      emp['Zoho_ID'] ||
      ''
  ).trim();
}

export function formatFormXVRJSexAndIdentificationToken(emp = {}) {
  const sex = String(emp.Sex || emp.Gender || emp['Gender'] || emp['Sex'] || '').trim();
  const token = resolveFormXVRJEmployeeToken(emp);
  const marks = String(
    emp.Identification_Marks ||
      emp['Identification Marks'] ||
      emp.IdentificationMarks ||
      emp.identification_marks ||
      emp['Identification_marks'] ||
      ''
  ).trim();
  const tokenLine = token || marks;
  if (sex && tokenLine) return `${sex}\n${tokenLine}`;
  return sex || tokenLine;
}

export function applyFormXVRJEmployeeToHeader(headerData, emp = {}) {
  let out = { ...(headerData || {}) };
  const assign = (key, value) => {
    const text = String(value ?? '').trim();
    if (text) out[key] = text;
  };
  const workman = formatWorkmanNameAndGuardian(emp);
  if (workman) assign('form_xv_rj_workman', workman);
  const sexToken = formatFormXVRJSexAndIdentificationToken(emp);
  if (sexToken) assign('form_xv_rj_sex_token', sexToken);
  return out;
}

const FORM_XV_RJ_STACKED_VALUE_COL = 2;

const formXVRJHeaderNorm = (text) =>
  String(text || '')
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();

const normalizeTableHeader = (text) => formXVRJHeaderNorm(text);

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

const labelMatchesSpec = (raw, spec) => {
  const cell = String(raw || '').trim();
  if (!cell) return false;
  const norm = formXVRJHeaderNorm(cell);
  if (spec.match.test(cell) || spec.match.test(norm)) return true;
  if (spec.key === 'form_xv_rj_workman' && /name/.test(norm) && /father/.test(norm) && /workman/.test(norm)) {
    return true;
  }
  if (
    spec.key === 'form_xv_rj_sex_token' &&
    /sex/.test(norm) &&
    /(identification|token|ticket)/.test(norm)
  ) {
    return true;
  }
  return false;
};

export function getFormXVRJRowValueForHeader(row, header) {
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

export function rowHasMeaningfulFormXVRJExportData(row, headers) {
  if (!row || typeof row !== 'object') return false;
  if (String(row.__employeeLookupName ?? '').trim()) return true;
  if (String(row.__employeeLookupId ?? '').trim()) return true;
  const hdrs = resolveFormXVRJTableHeaders(headers);
  return hdrs.some((header) => {
    if (isFormXVRJSignatureHeader(header)) return false;
    return getFormXVRJRowValueForHeader(row, header) !== '';
  });
}

const buildFormXVRJExportPairs = (mappedData, employeesOverride, hdrs) => {
  const rows = Array.isArray(mappedData) ? mappedData : [];
  const emps = Array.isArray(employeesOverride) ? employeesOverride : [];
  const exportAllRoster = emps.length > 0 && emps.length === rows.length;
  const pairs = [];
  for (let i = 0; i < rows.length; i += 1) {
    const row = rows[i];
    if (!exportAllRoster && !rowHasMeaningfulFormXVRJExportData(row, hdrs)) continue;
    pairs.push({ row, emp: emps[i] ?? null, index: i });
  }
  return pairs;
};

const setCellValue = (worksheet, row, col, value) => {
  const text = String(value ?? '').trim();
  if (!text || row < 1 || col < 1) return;
  const cell = worksheet.getCell(row, col);
  cell.value = text;
  const multiline = text.includes('\n');
  cell.alignment = {
    ...(cell.alignment || {}),
    horizontal: 'left',
    vertical: multiline ? 'top' : cell.alignment?.vertical || 'top',
    wrapText: multiline,
    shrinkToFit: false,
  };
};

const writeBesideLabel = (worksheet, row, labelCol, value, defaultValueCol = FORM_XV_RJ_STACKED_VALUE_COL) => {
  let targetCol = defaultValueCol;
  for (let c = labelCol + 1; c <= Math.min(labelCol + 12, 14); c += 1) {
    const cellStr = excelCellValueToString(worksheet.getCell(row, c)?.value).trim();
    if (!cellStr || isPlaceholderCell(cellStr)) {
      targetCol = c;
      break;
    }
    if (!/name\s+and\s+address|father|workman|sex|identification|week|fortnight|month/i.test(cellStr)) {
      targetCol = c;
      break;
    }
  }
  setCellValue(worksheet, row, targetCol, value);
};

export function writeFormXVRJHeaderFieldsToWorksheet(worksheet, headerFormData = {}, parsedFormHeader = null) {
  if (!worksheet || !headerFormData || typeof headerFormData !== 'object') return;
  const parsedFields = Array.isArray(parsedFormHeader?.fields) ? parsedFormHeader.fields : [];
  const parsedByKey = new Map(parsedFields.map((f) => [f.key, f]));

  FORM_XV_RJ_HEADER_SPECS.forEach((spec) => {
    const val = headerFormData[spec.key];
    if (val == null || String(val).trim() === '') return;
    const parsedField = parsedByKey.get(spec.key);
    if (parsedField?.labelRow != null) {
      const labelExcelRow = parsedField.labelRow + 1;
      const labelExcelCol = (parsedField.labelCol ?? 0) + 1;
      const targetRow = (parsedField.valueRow ?? parsedField.labelRow) + 1;
      const targetCol =
        parsedField.valueCol != null ? parsedField.valueCol + 1 : FORM_XV_RJ_STACKED_VALUE_COL;
      if (parsedField.valueCol != null) {
        setCellValue(worksheet, targetRow, targetCol, val);
      } else {
        writeBesideLabel(worksheet, labelExcelRow, labelExcelCol, val);
      }
      return;
    }
    for (let r = 1; r <= 30; r += 1) {
      for (let c = 1; c <= 10; c += 1) {
        const raw = excelCellValueToString(worksheet.getCell(r, c)?.value).trim();
        if (!raw || !labelMatchesSpec(raw, spec)) continue;
        writeBesideLabel(worksheet, r, c, val);
        return;
      }
    }
  });
}

const resolveTableHeaderMatchCol = (worksheet, headerRow, header) => {
  const target = normalizeTableHeader(header);
  for (let c = 1; c <= 12; c += 1) {
    const cellNorm = normalizeTableHeader(excelCellValueToString(worksheet.getCell(headerRow, c)?.value));
    if (!cellNorm) continue;
    if (cellNorm === target || cellNorm.includes(target) || target.includes(cellNorm)) return c;
    if (isFormXVRJDaysWorkedHeader(header) && /days/.test(cellNorm) && /worked/.test(cellNorm)) return c;
    if (isFormXVRJDailyRateHeader(header) && /rate/.test(cellNorm) && /(daily|piece)/.test(cellNorm)) return c;
    if (isFormXVRJUnitsWorkedHeader(header) && /units/.test(cellNorm)) return c;
    if (isFormXVRJOvertimeDateHeader(header) && /date/.test(cellNorm) && /overtime/.test(cellNorm)) return c;
    if (isFormXVRJOvertimeHoursAmountHeader(header) && /overtime/.test(cellNorm) && /(hour|amount)/.test(cellNorm)) {
      return c;
    }
    if (isFormXVRJGrossWagesHeader(header) && /gross/.test(cellNorm) && /wage/.test(cellNorm)) return c;
    if (isFormXVRJDeductionsHeader(header) && /deduction/.test(cellNorm)) return c;
    if (isFormXVRJActualWagesPaidHeader(header) && /actual/.test(cellNorm) && /paid/.test(cellNorm)) return c;
    if (isFormXVRJSignatureHeader(header) && /signature/.test(cellNorm)) return c;
  }
  return -1;
};

const resolveFormXVRJTableExportLayout = (worksheet, hdrs) => {
  if (!worksheet) return null;
  const exportHdrs = resolveFormXVRJTableHeaders(hdrs);
  for (let r = 15; r <= 30; r += 1) {
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

const writeFormXVRJTableRowToWorksheet = (worksheet, row, layout) => {
  if (!worksheet || !row || !layout) return;
  const { dataStartRow, columnByHeader, hdrs } = layout;
  hdrs.forEach((header, j) => {
    const col = columnByHeader[j];
    if (!col || col < 1) return;
    const val = getFormXVRJRowValueForHeader(row, header);
    if (!val) return;
    setCellValue(worksheet, dataStartRow, col, val);
  });
};

export async function buildFormXVRJWorkbookWithTemplateStyles({
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

  const hdrs = resolveFormXVRJTableHeaders(headersToUse);
  const sourceRows = (Array.isArray(mappedData) ? mappedData : []).filter((row) =>
    rowHasMeaningfulFormXVRJExportData(row, hdrs)
  );
  const employeeRow = sourceRows.length >= 1 ? sourceRows[0] : null;

  writeFormXVRJHeaderFieldsToWorksheet(worksheet, headerFormData, parsedFormHeader);
  if (employeeRow) {
    const tableLayout = resolveFormXVRJTableExportLayout(worksheet, hdrs);
    if (tableLayout) {
      writeFormXVRJTableRowToWorksheet(worksheet, employeeRow, tableLayout);
    }
  }

  const out = await workbook.xlsx.writeBuffer();
  const fileName =
    formFileName ||
    parsedFormHeader?.title?.replace(/[^a-zA-Z0-9]/g, '_') ||
    `Form_XV_RJ_${Date.now()}.xlsx`;
  const blob = new Blob([out], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  return { blob, fileName };
}

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
  if (count === 0) return `Form_XV_RJ_${root}.${ext}`;
  return `Form_XV_RJ_${root}_${count + 1}.${ext}`;
};

const unwrapEmployee = (empItem) => empItem?.Employee || empItem?.employee || empItem || null;

const resolveEmployeeDownloadBaseName = (row, index, employeesOverride) => {
  const fromRow = String(row?.__employeeLookupName ?? '').trim();
  if (fromRow) return fromRow.split(/\r?\n/)[0].trim();
  const emp = Array.isArray(employeesOverride) ? employeesOverride[index] : null;
  const fromEmp = emp
    ? String(formatWorkmanNameAndGuardian(unwrapEmployee(emp)) || '')
        .split(/\r?\n/)[0]
        .trim()
    : '';
  return fromEmp || `Employee_${index + 1}`;
};

export function triggerFormXVRJZipDownload(blob, fileName) {
  if (!blob || !fileName) return;
  const downloadUrl = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = downloadUrl;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(downloadUrl);
}

export async function buildFormXVRJPerEmployeeDownload({
  templateArrayBuffer,
  mappedData,
  headersToUse,
  parsedFormHeader,
  formFileName,
  headerFormData,
  employeesOverride = null,
}) {
  const hdrs = resolveFormXVRJTableHeaders(headersToUse);
  const exportPairs = buildFormXVRJExportPairs(mappedData, employeesOverride, hdrs);
  const workbookArgs = {
    templateArrayBuffer,
    headersToUse: hdrs,
    parsedFormHeader,
    formFileName,
  };

  if (exportPairs.length <= 1) {
    const pair = exportPairs[0] || null;
    const emp = pair?.emp ? unwrapEmployee(pair.emp) : null;
    const mergedHeader = applyFormXVRJEmployeeToHeader(headerFormData || {}, emp || {});
    return buildFormXVRJWorkbookWithTemplateStyles({
      ...workbookArgs,
      mappedData: pair ? [pair.row] : [],
      headerFormData: mergedHeader,
    });
  }

  const zip = new JSZip();
  const usedNames = new Map();
  for (let i = 0; i < exportPairs.length; i += 1) {
    const { row, emp } = exportPairs[i];
    const mergedHeader = applyFormXVRJEmployeeToHeader(
      headerFormData || {},
      emp ? unwrapEmployee(emp) : {}
    );
    const { blob } = await buildFormXVRJWorkbookWithTemplateStyles({
      ...workbookArgs,
      mappedData: [row],
      headerFormData: mergedHeader,
    });
    const baseName = resolveEmployeeDownloadBaseName(row, i, employeesOverride);
    zip.file(allocateUniqueFileName(baseName, usedNames), blob);
    if (i > 0 && i % 10 === 0) {
      await new Promise((resolve) => setTimeout(resolve, 0));
    }
  }

  const zipBase = String(formFileName || parsedFormHeader?.title || 'Form_XV_RJ')
    .replace(/\.xlsx?$/i, '')
    .replace(/[^a-zA-Z0-9._-]+/g, '_');
  return {
    blob: await zip.generateAsync({ type: 'blob', compression: 'STORE' }),
    fileName: `${zipBase}_Employees.zip`,
  };
}
