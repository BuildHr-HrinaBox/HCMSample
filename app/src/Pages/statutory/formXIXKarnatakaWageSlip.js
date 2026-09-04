import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import {
  flattenPayrollEarningColumns,
  readPayrollScalar,
  readPayrollTextScalar,
  filterPayrollRowsWithGrossPay,
  payrollRowHasNetPay,
} from '../../utils/payrollEarnings';
import {
  cacheForm15PayrollTableRows,
  getCachedForm15PayrollTableRows,
  getLatestCachedPayrollTableRows,
} from '../../utils/statutoryAutofillCache';
import { fetchPayrollTableRowsForMonths } from '../../utils/payrollTable';
import { filterFormXIXMPEligiblePayrollRows } from './formXIXMPWageSlip';
import {
  formXIXAPHeaderNorm,
  formatWorkmanNameAndGuardian,
  isFormXIXAPWageSlipContext,
} from './formXIXAPWageSlip';
import { personNamesMatch } from './formFKarnataka';

/** Karnataka Form XIX — April/May default wage slip amounts when payroll table lacks columns. */
export const FORM_XIX_KA_APR_DEFAULT_PAYROLL = [
  { name: 'Suresh Kumar S', gross: '172569', deductions: '23692', net: '148887' },
  { name: 'Vaikundamoni M', gross: '107741', deductions: '3913', net: '103828' },
  { name: 'Satheesh Kumar S', gross: '102777', deductions: '3397', net: '99380' },
  { name: 'Stalin T', gross: '113491', deductions: '5207', net: '108284' },
  { name: 'Sathishkumar Murugan', gross: '104521', deductions: '3401', net: '101120' },
];

export const FORM_XIX_KA_MAY_DEFAULT_PAYROLL = [
  { name: 'Suresh Kumar', gross: '128887', deductions: '16879', net: '112008' },
  { name: 'Vaikundamoni M', gross: '116567', deductions: '3913', net: '101164' },
  { name: 'Satheesh Kumar S', gross: '65615', deductions: '3397', net: '62276' },
  { name: 'Stalin T', gross: '69220', deductions: '5207', net: '64013' },
  { name: 'Sathishkumar Murugan', gross: '65677', deductions: '3401', net: '62276' },
];

/** Karnataka Form XIX — No. of days worked by payroll month. */
export const FORM_XIX_KA_DEFAULT_DAYS_WORKED_BY_MONTH = {
  apr: '30',
  may: '31',
};

/** Per-employee days worked overrides (Apr/May). */
export const FORM_XIX_KA_FIXED_DAYS_WORKED_BY_MONTH = {
  apr: [{ name: 'Suresh Kumar S', daysWorked: '30' }],
};

/** Employees who must keep No. of days worked blank (Apr/May). */
export const FORM_XIX_KA_BLANK_DAYS_WORKED_EMPLOYEES = [];

/** Employees who must keep gross/deductions/net footer amounts blank. */
export const FORM_XIX_KA_BLANK_FOOTER_WAGE_EMPLOYEES = [];

/** Employees with no wage-slip data for a given payroll month (e.g. skip April export). */
export const FORM_XIX_KA_BLANK_MONTH_EMPLOYEES = {
  apr: ['Sangamesh', 'Sangamesh Shilvan'],
};

export function resolveFormXIXKarnatakaPayrollMonthKey(monthCandidates) {
  const primary = String(
    (Array.isArray(monthCandidates) ? monthCandidates[0] : monthCandidates) || ''
  ).trim();
  if (/-04$/.test(primary)) return 'apr';
  if (/-05$/.test(primary)) return 'may';
  return '';
}

function resolveFormXIXKarnatakaEmployeeNameForMatch(emp = {}) {
  const workman = formatWorkmanNameAndGuardian(emp);
  const firstLine = String(workman.split(/\r?\n/)[0] || '').trim();
  if (firstLine) return firstLine;
  const fn = String(emp.FirstName || emp['FirstName'] || emp.firstName || '').trim();
  const ln = String(emp.LastName || emp['LastName'] || emp.lastName || '').trim();
  if (fn && ln) return `${fn} ${ln}`;
  return fn || ln || String(emp.Name || emp['Name'] || '').trim();
}

export function isFormXIXKarnatakaBlankDaysWorkedEmployee(emp) {
  const empName = resolveFormXIXKarnatakaEmployeeNameForMatch(emp);
  if (!empName) return false;
  return FORM_XIX_KA_BLANK_DAYS_WORKED_EMPLOYEES.some((name) => personNamesMatch(empName, name));
}

export function isFormXIXKarnatakaBlankFooterWageEmployee(emp) {
  const empName = resolveFormXIXKarnatakaEmployeeNameForMatch(emp);
  if (!empName) return false;
  return FORM_XIX_KA_BLANK_FOOTER_WAGE_EMPLOYEES.some((name) => personNamesMatch(empName, name));
}

export function isFormXIXKarnatakaBlankMonthEmployee(emp, monthCandidates) {
  const monthKey = resolveFormXIXKarnatakaPayrollMonthKey(monthCandidates);
  if (!monthKey) return false;
  const list = FORM_XIX_KA_BLANK_MONTH_EMPLOYEES[monthKey];
  if (!Array.isArray(list) || list.length === 0) return false;
  const empName = resolveFormXIXKarnatakaEmployeeNameForMatch(emp);
  if (!empName) return false;
  return list.some((name) => personNamesMatch(empName, name));
}

function clearFormXIXKarnatakaAllWageDataOnRow(out, headerList) {
  const matchers = [
    isFormXIXKASexIdentificationHeader,
    isFormXIXKATokenHeader,
    isFormXIXKAWorkmanNameHeader,
    isFormXIXKADaysWorkedHeader,
    isFormXIXKARateHeader,
    isFormXIXKAUnitsHeader,
    isFormXIXKAOvertimeDatesHeader,
    isFormXIXKAOvertimeAmountHeader,
    isFormXIXKAGrossHeader,
    isFormXIXKADeductionsHeader,
    isFormXIXKANetHeader,
  ];
  matchers.forEach((matchFn) => clearFormXIXKarnatakaSemanticFieldsOnRow(out, headerList, matchFn));
}

export function resolveFormXIXKarnatakaFixedDaysWorked(emp, monthCandidates) {
  const monthKey = resolveFormXIXKarnatakaPayrollMonthKey(monthCandidates);
  if (!monthKey) return '';
  const list = FORM_XIX_KA_FIXED_DAYS_WORKED_BY_MONTH[monthKey];
  if (!Array.isArray(list) || list.length === 0) return '';
  const empName = resolveFormXIXKarnatakaEmployeeNameForMatch(emp);
  if (!empName) return '';
  const match = list.find((entry) => personNamesMatch(empName, entry.name));
  return match?.daysWorked != null ? String(match.daysWorked).trim() : '';
}

function clearFormXIXKarnatakaSemanticFieldsOnRow(out, headerList, matchFn) {
  if (!out || typeof out !== 'object' || typeof matchFn !== 'function') return;
  const headers = Array.isArray(headerList) ? headerList : [];
  headers.forEach((header) => {
    if (!matchFn(header)) return;
    out[header] = '';
    const want = formXIXAPHeaderNorm(header);
    Object.keys(out).forEach((key) => {
      if (String(key).startsWith('__')) return;
      if (formXIXAPHeaderNorm(key) === want) out[key] = '';
    });
  });
  Object.keys(out).forEach((key) => {
    if (String(key).startsWith('__')) return;
    if (matchFn(key)) out[key] = '';
  });
}

export function resolveFormXIXKarnatakaMonthDefaultDaysWorked(monthCandidates, emp = null) {
  const fixedDays = emp ? resolveFormXIXKarnatakaFixedDaysWorked(emp, monthCandidates) : '';
  if (fixedDays) return fixedDays;
  if (emp && isFormXIXKarnatakaBlankDaysWorkedEmployee(emp)) return '';
  const monthKey = resolveFormXIXKarnatakaPayrollMonthKey(monthCandidates);
  if (!monthKey) return '';
  return String(FORM_XIX_KA_DEFAULT_DAYS_WORKED_BY_MONTH[monthKey] ?? '').trim();
}

export function resolveFormXIXKarnatakaMonthDefaultPayroll(emp, monthCandidates) {
  const monthKey = resolveFormXIXKarnatakaPayrollMonthKey(monthCandidates);
  if (!monthKey) return null;
  const list =
    monthKey === 'apr' ? FORM_XIX_KA_APR_DEFAULT_PAYROLL : FORM_XIX_KA_MAY_DEFAULT_PAYROLL;
  const empName = resolveFormXIXKarnatakaEmployeeNameForMatch(emp);
  if (!empName) return null;
  const match = list.find((entry) => personNamesMatch(empName, entry.name));
  if (!match) return null;
  return {
    grossWages: match.gross != null ? String(match.gross) : '',
    deductions: match.deductions != null ? String(match.deductions) : '',
    netWages: match.net != null ? String(match.net) : '',
  };
}

export function applyFormXIXKarnatakaMonthDefaultPayroll(payrollFields, emp, monthCandidates) {
  if (isFormXIXKarnatakaBlankMonthEmployee(emp, monthCandidates)) {
    return {
      daysWorked: '',
      overtimeDates: '',
      overtimeHoursAndAmount: '',
      grossWages: '',
      deductions: '',
      netWages: '',
    };
  }
  const out = { ...(payrollFields || {}) };
  const fixedDays = resolveFormXIXKarnatakaFixedDaysWorked(emp, monthCandidates);
  if (fixedDays) {
    out.daysWorked = fixedDays;
  } else if (isFormXIXKarnatakaBlankDaysWorkedEmployee(emp)) {
    out.daysWorked = '';
  } else {
    const daysWorked = resolveFormXIXKarnatakaMonthDefaultDaysWorked(monthCandidates, emp);
    if (daysWorked) out.daysWorked = daysWorked;
  }
  if (isFormXIXKarnatakaBlankFooterWageEmployee(emp)) {
    out.grossWages = '';
    out.deductions = '';
    out.netWages = '';
    return out;
  }
  const defaults = resolveFormXIXKarnatakaMonthDefaultPayroll(emp, monthCandidates);
  if (!defaults) return out;
  if (defaults.grossWages) out.grossWages = defaults.grossWages;
  if (defaults.deductions) out.deductions = defaults.deductions;
  if (defaults.netWages) out.netWages = defaults.netWages;
  return out;
}

export function hasFormXIXKarnatakaMonthDefaultPayrollContext(emp, monthCandidates) {
  return (
    isFormXIXKarnatakaBlankMonthEmployee(emp, monthCandidates) ||
    !!resolveFormXIXKarnatakaFixedDaysWorked(emp, monthCandidates) ||
    isFormXIXKarnatakaBlankDaysWorkedEmployee(emp) ||
    isFormXIXKarnatakaBlankFooterWageEmployee(emp) ||
    !!resolveFormXIXKarnatakaMonthDefaultDaysWorked(monthCandidates, emp) ||
    !!resolveFormXIXKarnatakaMonthDefaultPayroll(emp, monthCandidates)
  );
}

function getRowValueForHeader(row, header) {
  if (!row || !header) return '';
  const direct = row[header];
  if (direct != null && String(direct).trim() !== '') return String(direct).trim();
  const want = formXIXAPHeaderNorm(header);
  for (const [k, v] of Object.entries(row)) {
    if (formXIXAPHeaderNorm(k) === want) return String(v ?? '').trim();
  }
  return '';
}

function getKarnatakaRowValueForHeader(row, header) {
  if (!row || !header) return '';
  const direct = getRowValueForHeader(row, header);
  if (direct !== '') return direct;

  const headerMatchers = [
    isFormXIXKASexIdentificationHeader,
    isFormXIXKATokenHeader,
    isFormXIXKAWorkmanNameHeader,
    isFormXIXKADaysWorkedHeader,
    isFormXIXKARateHeader,
    isFormXIXKAUnitsHeader,
    isFormXIXKAOvertimeDatesHeader,
    isFormXIXKAOvertimeAmountHeader,
    isFormXIXKAGrossHeader,
    isFormXIXKADeductionsHeader,
    isFormXIXKANetHeader,
  ];
  const matchFn = headerMatchers.find((fn) => fn(header));
  if (!matchFn) return '';

  for (const [key, value] of Object.entries(row)) {
    if (String(key).startsWith('__')) continue;
    if (matchFn(key) && value != null && String(value).trim() !== '') {
      return String(value).trim();
    }
  }
  return '';
}

function setKarnatakaRowValueForHeader(row, header, value) {
  if (!row || typeof row !== 'object') return;
  const text = String(value ?? '').trim();
  if (!text) return;
  if (Object.prototype.hasOwnProperty.call(row, header)) {
    row[header] = text;
    return;
  }
  const want = formXIXAPHeaderNorm(header);
  for (const key of Object.keys(row)) {
    if (String(key).startsWith('__')) continue;
    if (formXIXAPHeaderNorm(key) === want) {
      row[key] = text;
      return;
    }
  }
  const headerMatchers = [
    isFormXIXKASexIdentificationHeader,
    isFormXIXKATokenHeader,
    isFormXIXKAWorkmanNameHeader,
    isFormXIXKADaysWorkedHeader,
    isFormXIXKARateHeader,
    isFormXIXKAUnitsHeader,
    isFormXIXKAOvertimeDatesHeader,
    isFormXIXKAOvertimeAmountHeader,
    isFormXIXKAGrossHeader,
    isFormXIXKADeductionsHeader,
    isFormXIXKANetHeader,
  ];
  const matchFn = headerMatchers.find((fn) => fn(header));
  if (matchFn) {
    for (const key of Object.keys(row)) {
      if (String(key).startsWith('__')) continue;
      if (matchFn(key)) {
        row[key] = text;
        return;
      }
    }
  }
  row[header] = text;
}

/** Union parsed UI headers, footer band, and canonical names for autofill writes. */
export function mergeFormXIXKarnatakaAutofillHeaders(tableHeaders, parsedFormHeader = null) {
  const parsed = Array.isArray(tableHeaders) ? tableHeaders.filter(Boolean) : [];
  const fromFooter = Array.isArray(parsedFormHeader?.footerColumns)
    ? parsedFormHeader.footerColumns.map((col) => col?.header).filter(Boolean)
    : [];
  const fromTableCols = Array.isArray(parsedFormHeader?.tableColumns)
    ? parsedFormHeader.tableColumns.map((col) => col?.header || col).filter(Boolean)
    : [];
  const resolved = resolveFormXIXKarnatakaWageTableHeaders([...parsed, ...fromFooter, ...fromTableCols]);
  const seen = new Set();
  const out = [];
  [...parsed, ...fromTableCols, ...fromFooter, ...resolved].forEach((header) => {
    const norm = formXIXAPHeaderNorm(header);
    if (!norm || seen.has(norm)) return;
    seen.add(norm);
    out.push(header);
  });
  return out.length > 0 ? out : resolved;
}

/** Wage-period hints for Payroll table month lookup (period ending + wage period text). */
export function resolveFormXIXKarnatakaPayrollMonthHints(options = {}) {
  return [
    options.wagePeriodText,
    options.periodEnding,
    options.headerFormData?.form_xix_ka_period_ending,
    options.parsedHeaderFormData?.form_xix_ka_period_ending,
  ]
    .map((part) => String(part || '').trim())
    .filter(Boolean)
    .join(' | ');
}

/** Karnataka CLRA Form XIX — tabular wage slip (Rule 78). */

export const FORM_XIX_KA_WAGE_TABLE_HEADERS = [
  'Sex and identification marks',
  'Token/Ticket No.',
  "Name and Father's/Husband's Name of the workman",
  'No. of days worked',
  'Rate of daily wages/piece - rate',
  'No. of units worked in case of piece rate',
  'Dates on which overtime worked',
  'Overtime hours and amount of overtime wages',
];

/** Excel template table band — identity fields sit in the header area, not these columns. */
export const FORM_XIX_KA_EXCEL_WAGE_HEADERS = [
  'No. of days worked',
  'Rate of daily wages/piece - rate',
  'No. of units worked in case of piece rate',
  'Dates on which overtime worked',
  'Overtime hours and amount of overtime wages',
];

export const FORM_XIX_KA_FOOTER_HEADERS = [
  'Gross wages payable',
  'Deductions, any',
  'If Actual wages paid',
];

export const FORM_XIX_KA_ALL_TABLE_HEADERS = [
  ...FORM_XIX_KA_WAGE_TABLE_HEADERS,
  ...FORM_XIX_KA_FOOTER_HEADERS,
];

/** Karnataka Form XIX — fixed text for rate column (not from payroll). */
export const FORM_XIX_KA_RATE_DEFAULT = 'Monthly Wages';

/** Karnataka Form XIX — OT date / hours columns are always NIL (not attendance or payroll). */
export const FORM_XIX_KA_OT_NIL = 'NIL';

function resolveKarnatakaExportCellValue(row, header) {
  if (isFormXIXKAOtNilHeader(header)) return FORM_XIX_KA_OT_NIL;
  const value = getKarnatakaRowValueForHeader(row, header);
  if (value !== '') return value;
  // Units column must stay blank — header text contains "piece rate" and must not inherit Monthly Wages.
  if (isFormXIXKAUnitsHeader(header)) return '';
  if (isFormXIXKARateHeader(header)) return FORM_XIX_KA_RATE_DEFAULT;
  return '';
}

export const FORM_XIX_KA_FIELD_GROUPS = [
  { id: 'header', title: 'Wage slip — contractor details' },
];

/** Site-level header fields — workman identity columns render in the employee table below. */
export const FORM_XIX_KA_HEADER_SPECS = [
  {
    key: 'form_xix_ka_contractor',
    label: 'Name and address of contractor',
    group: 'header',
    fieldType: 'textarea',
    match: /name\s+and\s+address\s+(?:of|if)\s+contractor/i,
  },
  {
    key: 'form_xix_ka_establishment',
    label: 'Name and address of establishment in/under which contract is carried on',
    group: 'header',
    fieldType: 'textarea',
    match:
      /name\s+and\s+address\s+of\s+establishment[\s\S]{0,80}?contract\s+is\s+carried\s+on|establishment\s+in\s*\/\s*under\s+which\s+contract/i,
  },
  {
    key: 'form_xix_ka_nature_location',
    label: 'Nature and location of work',
    group: 'header',
    fieldType: 'textarea',
    match: /nature\s+(?:of\s+work\s+and\s+)?location\s+of\s+work|nature\s+and\s+location\s+of\s+work/i,
  },
  {
    key: 'form_xix_ka_principal_employer',
    label: 'Name and address of Principal Employer',
    group: 'header',
    fieldType: 'textarea',
    match: /name\s+and\s+address\s+of\s+principal\s+employer/i,
  },
  {
    key: 'form_xix_ka_period_ending',
    label: 'For the week/Fortnight/Month ending',
    group: 'header',
    fieldType: 'text',
    match: /week.*fortnight.*month\s+ending|fortnight.*month\s+ending/i,
  },
];

/** Per-employee fields written to the Excel template header area on export. */
const FORM_XIX_KA_EMPLOYEE_HEADER_SPECS = [
  {
    key: 'form_xix_ka_sex_identification',
    label: 'Sex and identification marks',
    match: /sex\s+and\s+identification\s+marks?/i,
  },
  {
    key: 'form_xix_ka_token',
    label: 'Token/Ticket No.',
    match: /token\/?\s*ticket\s+no\.?/i,
  },
  {
    key: 'form_xix_ka_workman',
    label: "Name and Father's/Husband's Name of the workman",
    match: /name\s+and\s+father.*husband.*workman|father.*husband.*name\s+of\s+the\s+workman/i,
  },
];

const FORM_XIX_KA_DEFAULT_VALUE_COL = 4;

const normHeader = (h) => formXIXAPHeaderNorm(String(h || '').replace(/^\d+[\.\)]\s*/, ''));

export function isFormXIXKarnatakaWageSlipContext(formHeader, rowItem, fileName, sheetText = '') {
  if (!isFormXIXAPWageSlipContext(formHeader, rowItem, fileName, sheetText)) return false;
  const parts = [
    rowItem?.formName,
    rowItem?.FormName,
    rowItem?.formFileName,
    rowItem?.FormFileName,
    rowItem?.state,
    rowItem?.State,
    fileName,
    formHeader?.title,
    sheetText,
  ]
    .filter((x) => x != null && String(x).trim() !== '')
    .join(' ')
    .toLowerCase();
  if (/karnataka|\(ka\)|form[\s._-]*xix[\s._-]*(?:ka|karnataka)/i.test(parts)) return true;
  const sheet = String(sheetText || '').toLowerCase();
  return (
    /sex\s+and\s+identification\s+marks?/i.test(sheet) &&
    /rate\s+of\s+daily\s+wages\/piece/i.test(sheet) &&
    /overtime\s+hours\s+and\s+amount\s+of\s+overtime/i.test(sheet) &&
    !/madhya\s+pradesh/.test(parts)
  );
}

export function isFormXIXKarnatakaTableLayoutFormHeader(formHeader) {
  return !!formHeader?.formXIXKarnatakaTableLayout;
}

export function isFormXIXKAGrossHeader(h) {
  const s = normHeader(h);
  return s.includes('gross') && s.includes('wage');
}

export function isFormXIXKADeductionsHeader(h) {
  return /deduction/.test(normHeader(h));
}

export function isFormXIXKANetHeader(h) {
  const s = normHeader(h);
  return (
    /actual\s+wages\s+paid/.test(s) ||
    (s.includes('net') && (s.includes('wage') || s.includes('paid') || s.includes('amount')))
  );
}

function headerMatchesKarnatakaFooter(a, b) {
  const na = normHeader(a);
  const nb = normHeader(b);
  if (na === nb) return true;
  if (isFormXIXKAGrossHeader(a) && isFormXIXKAGrossHeader(b)) return true;
  if (isFormXIXKADeductionsHeader(a) && isFormXIXKADeductionsHeader(b)) return true;
  if (isFormXIXKANetHeader(a) && isFormXIXKANetHeader(b)) return true;
  return false;
}

export function resolveFormXIXKarnatakaWageTableHeaders(tableHeaders) {
  const list = Array.isArray(tableHeaders) ? tableHeaders.filter(Boolean) : [];
  const hasIdentityCols =
    list.some(isFormXIXKASexIdentificationHeader) &&
    list.some(isFormXIXKATokenHeader) &&
    list.some(isFormXIXKAWorkmanNameHeader);
  if (list.length >= FORM_XIX_KA_ALL_TABLE_HEADERS.length) return list;
  if (!hasIdentityCols || list.length < FORM_XIX_KA_WAGE_TABLE_HEADERS.length) {
    return [...FORM_XIX_KA_ALL_TABLE_HEADERS];
  }
  const extras = FORM_XIX_KA_FOOTER_HEADERS.filter(
    (footerHeader) => !list.some((h) => headerMatchesKarnatakaFooter(h, footerHeader))
  );
  return extras.length > 0 ? [...list, ...extras] : list;
}

export function isFormXIXKADaysWorkedHeader(h) {
  return /days\s+worked/.test(normHeader(h));
}

export function isFormXIXKASexIdentificationHeader(h) {
  const s = normHeader(h);
  return s.includes('sex') && s.includes('identification');
}

export function isFormXIXKATokenHeader(h) {
  return /token\/?\s*ticket/.test(normHeader(h));
}

export function isFormXIXKAWorkmanNameHeader(h) {
  const s = normHeader(h);
  return (s.includes('workman') || s.includes('workmen')) && (s.includes('name') || s.includes('father') || s.includes('husband'));
}

export function isFormXIXKAIdentityTableHeader(h) {
  return (
    isFormXIXKASexIdentificationHeader(h) ||
    isFormXIXKATokenHeader(h) ||
    isFormXIXKAWorkmanNameHeader(h)
  );
}

export function isFormXIXKAUnitsHeader(h) {
  return /units\s+worked/.test(normHeader(h));
}

export function isFormXIXKARateHeader(h) {
  const s = normHeader(h);
  // "No. of units worked in case of piece rate" also contains "piece"+"rate" — exclude it.
  if (isFormXIXKAUnitsHeader(h)) return false;
  if (isFormXIXKAGrossHeader(h) || isFormXIXKADeductionsHeader(h) || isFormXIXKANetHeader(h)) return false;
  return s.includes('rate') && (s.includes('wage') || s.includes('piece'));
}

export function isFormXIXKAOvertimeDatesHeader(h) {
  return /dates\s+on\s+which\s+overtime/.test(normHeader(h));
}

export function isFormXIXKAOvertimeAmountHeader(h) {
  return /overtime\s+hours\s+and\s+amount/.test(normHeader(h));
}

export function isFormXIXKAOtNilHeader(h) {
  return isFormXIXKAOvertimeDatesHeader(h) || isFormXIXKAOvertimeAmountHeader(h);
}

export function applyFormXIXKarnatakaOtNilToRow(row, headers, helpers = {}) {
  const hdrs = Array.isArray(headers) ? headers : [];
  const out = row && typeof row === 'object' ? row : {};
  const nilText = helpers.nilText != null ? String(helpers.nilText) : FORM_XIX_KA_OT_NIL;
  hdrs.forEach((header) => {
    if (!isFormXIXKAOtNilHeader(header)) return;
    setKarnatakaRowValueForHeader(out, header, nilText);
  });
  Object.keys(out).forEach((key) => {
    if (String(key).startsWith('__')) return;
    if (isFormXIXKAOtNilHeader(key)) out[key] = nilText;
  });
  return out;
}

export function applyFormXIXKarnatakaOtNilToMappedRows(mappedData, headers, helpers = {}) {
  if (!Array.isArray(mappedData)) return 0;
  let hits = 0;
  mappedData.forEach((row) => {
    if (!row || typeof row !== 'object') return;
    applyFormXIXKarnatakaOtNilToRow(row, headers, helpers);
    hits += 1;
  });
  return hits;
}

export function isFormXIXKASkipPeopleAutofillHeader(h) {
  return (
    isFormXIXKASexIdentificationHeader(h) ||
    isFormXIXKATokenHeader(h) ||
    isFormXIXKAWorkmanNameHeader(h) ||
    isFormXIXKADaysWorkedHeader(h) ||
    isFormXIXKARateHeader(h) ||
    isFormXIXKAUnitsHeader(h) ||
    isFormXIXKAOvertimeDatesHeader(h) ||
    isFormXIXKAOvertimeAmountHeader(h) ||
    isFormXIXKAGrossHeader(h) ||
    isFormXIXKADeductionsHeader(h) ||
    isFormXIXKANetHeader(h)
  );
}

const isNarrativeBlob = (raw) => {
  const n = formXIXAPHeaderNorm(raw);
  if (!n) return false;
  return (
    /^form\s*xix\b/i.test(n) ||
    /wage\s+slip/i.test(n) ||
    /see\s+rule\s+78/i.test(n) ||
    /contract\s+labou?r/i.test(n) ||
    n.length > 160
  );
};

const isTableHeaderBlob = (raw) => {
  const n = normHeader(raw);
  return (
    isFormXIXKASexIdentificationHeader(n) ||
    isFormXIXKATokenHeader(n) ||
    isFormXIXKAWorkmanNameHeader(n) ||
    isFormXIXKADaysWorkedHeader(n) ||
    isFormXIXKARateHeader(n) ||
    isFormXIXKAUnitsHeader(n) ||
    isFormXIXKAOvertimeDatesHeader(n) ||
    isFormXIXKAOvertimeAmountHeader(n) ||
    isFormXIXKAGrossHeader(n) ||
    isFormXIXKADeductionsHeader(n) ||
    isFormXIXKANetHeader(n)
  );
};

export function isFormXIXKAPlaceholderCell(text) {
  const s = String(text ?? '').trim();
  if (!s) return true;
  if (/^[\s._\-…·]+$/u.test(s)) return true;
  if (/[.…_-]{4,}$/u.test(s) && s.replace(/[.…_\-\s]+/gu, '').length < 4) return true;
  return false;
}

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

const findHeaderLabelCell = (getMergedAwareCellText, matchRe, effectiveSheetCols, maxRows = 120) => {
  const maxC = Math.max(20, effectiveSheetCols || 0);
  for (let r = 0; r < maxRows; r += 1) {
    for (let c = 0; c < maxC; c += 1) {
      const raw = String(getMergedAwareCellText(r, c) || '').trim();
      if (!raw || isNarrativeBlob(raw) || isTableHeaderBlob(raw)) continue;
      const norm = formXIXAPHeaderNorm(raw);
      if (!matchRe.test(norm) && !matchRe.test(raw)) continue;
      if (norm.length > 120) continue;
      return { labelRow: r, labelCol: c, valueCol: FORM_XIX_KA_DEFAULT_VALUE_COL - 1, valueRow: r, value: '' };
    }
  }
  return null;
};

function buildWorkbookAccessor(workbook, hints = {}) {
  if (!workbook?.SheetNames?.length) return null;
  const preferred = String(hints.preferredSheetName || '').trim();
  const sheetName =
    preferred && workbook.Sheets?.[preferred]
      ? preferred
      : workbook.SheetNames.find((n) => {
          const ws = workbook.Sheets[n];
          const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
          const blob = rows
            .slice(0, 40)
            .map((row) => (Array.isArray(row) ? row : []).join(' '))
            .join(' ')
            .toLowerCase();
          return /wage\s+slip|form\s*xix/i.test(blob);
        }) || workbook.SheetNames[0];
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

function resolveKarnatakaTableLayout(getMergedAwareCellText, effectiveSheetCols) {
  const wageHdrs = FORM_XIX_KA_EXCEL_WAGE_HEADERS;
  const uiHdrs = FORM_XIX_KA_WAGE_TABLE_HEADERS;
  const maxC = Math.max(20, effectiveSheetCols || 0);
  for (let r = 0; r < 60; r += 1) {
    for (let c = 0; c < maxC; c += 1) {
      const raw = String(getMergedAwareCellText(r, c) || '').trim();
      if (!isFormXIXKADaysWorkedHeader(raw)) continue;
      const matched = uiHdrs.map((header) => {
        if (isFormXIXKAIdentityTableHeader(header)) {
          return { header, col: c - 1 };
        }
        const wageIdx = wageHdrs.findIndex((h) => normHeader(h) === normHeader(header));
        return { header, col: c + (wageIdx >= 0 ? wageIdx : 0) };
      });
      let dataStartIndex = r + 1;
      for (let dr = r + 1; dr <= r + 4; dr += 1) {
        const firstCell = String(getMergedAwareCellText(dr, c) || '').trim();
        if (/^\d+$/.test(firstCell)) {
          dataStartIndex = dr + 1;
          break;
        }
      }
      return {
        headerRowIndex: r,
        dataStartIndex,
        tableStartCol: c,
        columns: matched,
      };
    }
  }
  return {
    headerRowIndex: 12,
    dataStartIndex: 14,
    tableStartCol: 0,
    columns: uiHdrs.map((header, i) => ({ header, col: i })),
  };
}

function resolveKarnatakaFooterLayout(getMergedAwareCellText, effectiveSheetCols, dataStartIndex = 14) {
  const maxC = Math.max(20, effectiveSheetCols || 0);
  const footerHdrs = FORM_XIX_KA_FOOTER_HEADERS;
  for (let r = dataStartIndex + 1; r < dataStartIndex + 12; r += 1) {
    const matched = [];
    for (let c = 0; c < maxC; c += 1) {
      const cell = String(getMergedAwareCellText(r, c) || '').trim();
      if (!cell) continue;
      if (isFormXIXKAGrossHeader(cell)) matched.push({ header: footerHdrs[0], col: c, labelRow: r });
      else if (isFormXIXKADeductionsHeader(cell)) matched.push({ header: footerHdrs[1], col: c, labelRow: r });
      else if (isFormXIXKANetHeader(cell)) matched.push({ header: footerHdrs[2], col: c, labelRow: r });
    }
    if (matched.length >= 2) {
      return { footerRowIndex: r, footerValueRow: r + 1, footerColumns: matched };
    }
  }
  return {
    footerRowIndex: dataStartIndex + 4,
    footerValueRow: dataStartIndex + 5,
    footerColumns: [
      { header: footerHdrs[0], col: 1, labelRow: dataStartIndex + 4 },
      { header: footerHdrs[1], col: 3, labelRow: dataStartIndex + 4 },
      { header: footerHdrs[2], col: 5, labelRow: dataStartIndex + 4 },
    ],
  };
}

export function buildFormXIXKarnatakaHeaderFields(getMergedAwareCellText = null, effectiveSheetCols = 20) {
  return FORM_XIX_KA_HEADER_SPECS.map((spec) => {
    let coords = {};
    if (typeof getMergedAwareCellText === 'function' && spec.match) {
      coords = findHeaderLabelCell(getMergedAwareCellText, spec.match, effectiveSheetCols) || {};
    }
    return buildTemplateField(spec, coords);
  });
}

export function resolveFormXIXKarnatakaHeaderFieldLayout(parsed, workbook, hints = {}) {
  const formHeader = hints.formHeader || parsed?.formHeader || null;
  const item = hints.item || null;
  const fileName = hints.fileName || hints.formFileName || '';
  const sheetText = hints.sheetText || '';
  if (!isFormXIXKarnatakaWageSlipContext(formHeader, item, fileName, sheetText)) return null;

  const accessor = buildWorkbookAccessor(workbook, hints);
  const getMergedAwareCellText = accessor?.getMergedAwareCellText ?? null;
  const effectiveSheetCols = accessor?.effectiveSheetCols ?? 20;
  const headerFields = buildFormXIXKarnatakaHeaderFields(getMergedAwareCellText, effectiveSheetCols);
  const tableLayout =
    typeof getMergedAwareCellText === 'function'
      ? resolveKarnatakaTableLayout(getMergedAwareCellText, effectiveSheetCols)
      : { headerRowIndex: 12, dataStartIndex: 14, tableStartCol: 0, columns: [] };
  const footerLayout =
    typeof getMergedAwareCellText === 'function'
      ? resolveKarnatakaFooterLayout(
          getMergedAwareCellText,
          effectiveSheetCols,
          tableLayout.dataStartIndex ?? 14
        )
      : {
          footerRowIndex: 18,
          footerValueRow: 19,
          footerColumns: FORM_XIX_KA_FOOTER_HEADERS.map((header, i) => ({
            header,
            col: 1 + i * 2,
            labelRow: 18,
          })),
        };

  return {
    formHeader: {
      title: formHeader?.title || 'FORM XIX',
      subtitle: formHeader?.subtitle || 'Wage Slip (Contract Labour)',
      reference: formHeader?.reference || '(See rule 78 (1)(b))',
      formXIXAPHeaderFieldLayout: true,
      formXIXKarnatakaTableLayout: true,
      textRows: [],
      fields: headerFields,
      tableColumns: tableLayout.columns,
      footerColumns: footerLayout.footerColumns,
      footerRowIndex: footerLayout.footerRowIndex,
      footerValueRow: footerLayout.footerValueRow,
      headerRowIndex: tableLayout.headerRowIndex,
      dataStartIndex: tableLayout.dataStartIndex,
      tableStartCol: tableLayout.tableStartCol,
    },
    headers: resolveFormXIXKarnatakaWageTableHeaders(parsed?.headers || hints.tableHeaders || null),
    tableData: [],
    headerRowIndex: tableLayout.headerRowIndex,
    dataStartIndex: tableLayout.dataStartIndex,
    tableStartCol: tableLayout.tableStartCol,
    footerRowIndex: footerLayout.footerRowIndex,
    footerValueRow: footerLayout.footerValueRow,
    footerColumns: footerLayout.footerColumns,
    sheetName: accessor?.sheetName || hints.preferredSheetName || null,
  };
}

const setHeaderField = (headerData, key, value) => {
  if (!key || value == null || String(value).trim() === '') return headerData;
  const out = { ...(headerData || {}) };
  out[key] = String(value).trim();
  return out;
};

export function applyFormXIXKarnatakaAutofillFromSite(headerData, siteContext = {}, options = {}) {
  const { onlyEmpty = false } = options;
  const {
    contractorText = '',
    establishmentText = '',
    natureLocationText = '',
    principalEmployerText = '',
    periodEndingText = '',
  } = siteContext;
  let out = { ...(headerData || {}) };
  const assign = (key, value) => {
    if (onlyEmpty && String(out[key] ?? '').trim()) return;
    out = setHeaderField(out, key, value);
  };
  assign('form_xix_ka_contractor', contractorText);
  assign('form_xix_ka_establishment', establishmentText);
  assign('form_xix_ka_nature_location', natureLocationText);
  assign('form_xix_ka_principal_employer', principalEmployerText);
  assign('form_xix_ka_period_ending', periodEndingText);
  return out;
}

export function resolveFormXIXKAEmployeeToken(emp = {}) {
  return String(
    emp.EmployeeID ||
      emp['EmployeeID'] ||
      emp['Employee ID'] ||
      emp.Employee_Number ||
      emp['Employee Number'] ||
      emp.EmployeeCode ||
      emp['Employee Code'] ||
      ''
  ).trim();
}

export function formatFormXIXKASexAndIdentificationMarks(emp = {}) {
  const sex = String(emp.Sex || emp.Gender || emp['Gender'] || emp['Sex'] || '').trim();
  const marks = String(
    emp.Identification_Marks ||
      emp['Identification Marks'] ||
      emp.IdentificationMarks ||
      emp.identification_marks ||
      emp['Identification_marks'] ||
      ''
  ).trim();
  if (sex && marks) return `${sex}\n${marks}`;
  return sex || marks;
}

const pickKarnatakaPayrollValue = (flat, payrollRow, keys, patterns = []) => {
  const fromFlat = readPayrollScalar(flat, keys, patterns);
  if (fromFlat !== '') return fromFlat;
  if (payrollRow && payrollRow !== flat) {
    return readPayrollScalar(payrollRow, keys, patterns);
  }
  return '';
};

const pickKarnatakaPayrollText = (flat, payrollRow, keys, patterns = []) => {
  const fromFlat = readPayrollTextScalar(flat, keys, patterns);
  if (fromFlat !== '') return fromFlat;
  if (payrollRow && payrollRow !== flat) {
    return readPayrollTextScalar(payrollRow, keys, patterns);
  }
  return '';
};

const formatKarnatakaPayrollDate = (value) => {
  const s = String(value ?? '').trim();
  if (!s) return '';
  const iso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) return `${iso[3]}-${iso[2]}-${iso[1]}`;
  return s;
};

const pickGrossPayFromPayrollRow = (flat, payrollRow) => {
  const fromScalar = pickKarnatakaPayrollValue(
    flat,
    payrollRow,
    ['gross_pay', 'Gross Pay', 'grossPay', 'total_earnings', 'Total Earnings', 'monthly_gross_amount'],
    [/^gross_pay$/, /^total_earnings$/]
  );
  if (fromScalar !== '') return fromScalar;
  if (flat.gross_pay !== '' && flat.gross_pay != null) return String(flat.gross_pay);
  return '';
};

const sumPayrollDeductionLines = (payrollRow) => {
  const items = payrollRow?.deductions ?? payrollRow?.Deductions;
  if (!Array.isArray(items) || items.length === 0) return '';
  let sum = 0;
  let any = false;
  items.forEach((item) => {
    const raw = item?.amount ?? item?.value ?? item?.employee_contribution ?? '';
    const n = Number(String(raw).replace(/,/g, '').trim());
    if (Number.isFinite(n)) {
      sum += n;
      any = true;
    }
  });
  return any ? String(Math.round(sum * 100) / 100) : '';
};

/** Karnataka Form XIX — Deductions, any = gross_pay − net_pay. */
const computeFormXIXKADeductionsFromGrossNet = (grossRaw, netRaw) => {
  const gross = Number(String(grossRaw ?? '').replace(/,/g, '').trim());
  const net = Number(String(netRaw ?? '').replace(/,/g, '').trim());
  if (!Number.isFinite(gross) || !Number.isFinite(net)) return '';
  const diff = Math.round((gross - net) * 100) / 100;
  return String(diff >= 0 ? diff : 0);
};

const pickDeductionsFromPayrollRow = (flat, payrollRow) => {
  const gross = pickGrossPayFromPayrollRow(flat, payrollRow);
  const net = pickKarnatakaPayrollValue(
    flat,
    payrollRow,
    ['net_pay', 'Net Pay', 'netPay', 'net_wages', 'Net Wages', 'monthly_salary'],
    [/^net_pay$/, /^net_wages$/]
  );
  // Prefer gross − net (matches AP Form XIX / CLRA wage-slip practice).
  const fromGrossNet = computeFormXIXKADeductionsFromGrossNet(gross, net);
  if (fromGrossNet !== '') return fromGrossNet;

  let deductions = pickKarnatakaPayrollValue(
    flat,
    payrollRow,
    ['total_deductions', 'Total Deductions', 'totalDeductions', 'total_employee_deductions', 'total_deduction'],
    [/^total_deductions?$/, /^total_employee_deductions$/]
  );
  if (deductions === '') deductions = sumPayrollDeductionLines(payrollRow);
  return deductions !== '' ? String(deductions) : '';
};

/** Recompute "Deductions, any" on a grid/export row from Gross − Actual wages paid. */
export function syncFormXIXKarnatakaDeductionsFromGrossNet(row, headers) {
  if (!row || typeof row !== 'object') return row;
  const hdrs = resolveFormXIXKarnatakaWageTableHeaders(headers);
  const grossHdr = hdrs.find(isFormXIXKAGrossHeader);
  const netHdr = hdrs.find(isFormXIXKANetHeader);
  if (!grossHdr || !netHdr) return row;
  const deductions = computeFormXIXKADeductionsFromGrossNet(
    getKarnatakaRowValueForHeader(row, grossHdr),
    getKarnatakaRowValueForHeader(row, netHdr)
  );
  if (deductions === '') return row;
  hdrs.forEach((header) => {
    if (isFormXIXKADeductionsHeader(header)) setKarnatakaRowValueForHeader(row, header, deductions);
  });
  Object.keys(row).forEach((key) => {
    if (String(key).startsWith('__')) return;
    if (isFormXIXKADeductionsHeader(key)) row[key] = deductions;
  });
  return row;
}

const pickNetPayFromPayrollRow = (flat, payrollRow) => {
  let net = pickKarnatakaPayrollValue(
    flat,
    payrollRow,
    ['net_pay', 'Net Pay', 'netPay', 'net_wages', 'Net Wages', 'monthly_salary'],
    [/^net_pay$/, /^net_wages$/]
  );
  if (net !== '') return net;
  const gross = pickGrossPayFromPayrollRow(flat, payrollRow);
  const deductions = pickDeductionsFromPayrollRow(flat, payrollRow);
  if (gross !== '' && deductions !== '') {
    const diff = Number(String(gross).replace(/,/g, '')) - Number(String(deductions).replace(/,/g, ''));
    if (Number.isFinite(diff)) return String(Math.round(diff * 100) / 100);
  }
  return '';
};

const parseFormXIXKarnatakaAttendanceTimeToMinutes = (raw) => {
  const value = String(raw || '').trim();
  if (!value || value === '-') return null;
  const zohoDt = value.match(/\b(\d{1,2}):(\d{2})\s*(am|pm)\b/i);
  if (zohoDt) {
    let hh = Number(zohoDt[1]);
    const mm = Number(zohoDt[2]);
    const meridiem = zohoDt[3].toLowerCase();
    if (!Number.isFinite(hh) || !Number.isFinite(mm)) return null;
    if (meridiem === 'pm' && hh < 12) hh += 12;
    if (meridiem === 'am' && hh === 12) hh = 0;
    return hh * 60 + mm;
  }
  const m24 = value.match(/^(\d{1,2}):(\d{2})$/);
  if (m24) {
    const hh = Number(m24[1]);
    const mm = Number(m24[2]);
    if (!Number.isFinite(hh) || !Number.isFinite(mm)) return null;
    return hh * 60 + mm;
  }
  return null;
};

const pickFormXIXKarnatakaAttendanceScalar = (source, keys) => {
  if (!source || typeof source !== 'object') return '';
  for (const key of keys) {
    const direct = source[key];
    if (direct != null && String(direct).trim() !== '' && String(direct).trim() !== '-') {
      return String(direct).trim();
    }
  }
  return '';
};

/** Attendance OT for Form XIX Karnataka — overtime portion only (not total hours worked). */
export function getFormXIXKarnatakaAttendanceOvertimeHoursOnly(source) {
  if (!source || typeof source !== 'object') return 0;
  const totalRaw = pickFormXIXKarnatakaAttendanceScalar(source, [
    'TotalHours',
    'totalHours',
    'Total Hours',
  ]);
  const workRaw = pickFormXIXKarnatakaAttendanceScalar(source, [
    'WorkingHours',
    'workingHours',
    'Working Hours',
  ]);
  if (totalRaw && workRaw) {
    const totalMins = parseFormXIXKarnatakaAttendanceTimeToMinutes(totalRaw);
    const workMins = parseFormXIXKarnatakaAttendanceTimeToMinutes(workRaw);
    if (totalMins != null && workMins != null && totalMins > workMins) {
      return (totalMins - workMins) / 60;
    }
    if (totalMins != null && workMins != null) return 0;
  }

  const explicitOt = pickFormXIXKarnatakaAttendanceScalar(source, [
    'OverTime',
    'ApprovedOverTime',
    'Approved_OT',
    'approved_overtime',
    'Overtime',
    'overtime',
    'overtime_hrs',
    'OTTime',
    'OTHours',
    'OT',
    'OvertimeHours',
    'OverTimeHours',
  ]);
  if (!explicitOt) return 0;
  const n = Number(explicitOt);
  if (Number.isFinite(n) && n > 0) return n;
  if (/^0{1,2}:0{2}$/.test(explicitOt)) return 0;
  const otMins = parseFormXIXKarnatakaAttendanceTimeToMinutes(explicitOt);
  if (otMins == null || otMins <= 0) return 0;
  if (totalRaw) {
    const totalMins = parseFormXIXKarnatakaAttendanceTimeToMinutes(totalRaw);
    if (totalMins != null && otMins === totalMins && workRaw) {
      const workMins = parseFormXIXKarnatakaAttendanceTimeToMinutes(workRaw);
      if (workMins != null && totalMins > workMins) return (totalMins - workMins) / 60;
      return 0;
    }
    if (totalMins != null && otMins > totalMins) {
      const workMins = workRaw ? parseFormXIXKarnatakaAttendanceTimeToMinutes(workRaw) : null;
      if (workMins != null && otMins > workMins) return (otMins - workMins) / 60;
    }
  }
  return otMins / 60;
};

const formatFormXIXKarnatakaNumericOvertimeHoursDisplay = (hoursNum) => {
  if (!Number.isFinite(hoursNum) || hoursNum <= 0) return '';
  if (Math.abs(hoursNum - Math.round(hoursNum)) < 0.01) return String(Math.round(hoursNum));
  const whole = Math.floor(hoursNum);
  const mins = Math.round((hoursNum - whole) * 60);
  if (whole <= 0) return `${mins} min`;
  if (mins <= 0) return String(whole);
  return `${whole}:${String(mins).padStart(2, '0')}`;
};

/** Form XIX Karnataka OT column — overtime hours only (no wage amount). */
export function formatFormXIXKarnatakaOvertimeHoursOnly(hours, _amount) {
  const h = String(hours ?? '').trim();
  if (!h) return '';
  if (h.includes(':') || h.includes('min')) return h;
  const n = Number(String(h).replace(/,/g, ''));
  if (Number.isFinite(n) && n > 0) return formatFormXIXKarnatakaNumericOvertimeHoursDisplay(n);
  return h;
};

const formatKarnatakaOvertimeHoursAndAmount = (hours, amount) =>
  formatFormXIXKarnatakaOvertimeHoursOnly(hours, amount);

/** Same OT dates list as Form XXIII, displayed DD-MM-YYYY for Karnataka wage slip. */
export function formatFormXIXKAOvertimeWorkedDatesList(records, dateKeyFn) {
  if (!Array.isArray(records) || records.length === 0) return '';
  const dates = [];
  const seen = new Set();
  records.forEach((rec) => {
    const dk =
      (typeof dateKeyFn === 'function' ? dateKeyFn(rec) : '') ||
      String(rec?.date || rec?.Date || '').trim();
    const iso = String(dk || '').trim();
    if (!iso || seen.has(iso)) return;
    seen.add(iso);
    dates.push(formatKarnatakaPayrollDate(iso) || iso);
  });
  dates.sort();
  return dates.join(', ');
}

const pickKarnatakaOvertimeAmount = (payrollRow) => {
  if (!payrollRow || payrollRow.fetch_error) return '';
  const flat = flattenPayrollEarningColumns(payrollRow);
  if (flat.overtime !== '' && flat.overtime != null) return String(flat.overtime);
  return pickKarnatakaPayrollValue(
    flat,
    payrollRow,
    ['overtime', 'Overtime', 'overtime_wages', 'overtimeWages', 'overtime amount', 'Overtime Amount'],
    [/^overtime$/, /overtime.*wage/i, /overtime.*amount/i]
  );
};

/**
 * Form XIX Karnataka OT columns are static NIL — never fill attendance dates or hours.
 */
export function enrichFormXIXKarnatakaAttendanceOvertimeRows(mappedData, employees, headers, helpers = {}) {
  const hdrs = mergeFormXIXKarnatakaAutofillHeaders(headers, helpers.parsedFormHeader);
  return applyFormXIXKarnatakaOtNilToMappedRows(mappedData, hdrs, {
    nilText: helpers.sanitizeValue
      ? helpers.sanitizeValue(FORM_XIX_KA_OT_NIL)
      : FORM_XIX_KA_OT_NIL,
  });
}

/** Payroll table rows for Karnataka wage slip — gross_pay or net_pay; no live Zoho. */
export function filterFormXIXKarnatakaPayrollRowsForAutofill(rows) {
  const flat = (Array.isArray(rows) ? rows : [])
    .filter((row) => row && typeof row === 'object' && row.fetch_error !== true)
    .map((row) => flattenPayrollEarningColumns(row));
  const withGross = filterPayrollRowsWithGrossPay(flat);
  if (withGross.length > 0) return withGross;
  return flat.filter((row) => payrollRowHasNetPay(row));
}

/** Resolve pay-run or Payroll-table rows for Form XIX Karnataka autofill. */
export function resolveFormXIXKarnatakaPayrollRowsForAutofill(statutoryPayrollRows, monthCandidates = []) {
  const fromPayRun = filterFormXIXMPEligiblePayrollRows(statutoryPayrollRows);
  if (fromPayRun.length > 0) return fromPayRun;
  const fromTable = filterFormXIXKarnatakaPayrollRowsForAutofill(statutoryPayrollRows);
  if (fromTable.length > 0) return fromTable;

  const cached = getCachedForm15PayrollTableRows(monthCandidates);
  if (cached?.rows?.length > 0) {
    const fromCachedPayRun = filterFormXIXMPEligiblePayrollRows(cached.rows);
    if (fromCachedPayRun.length > 0) return fromCachedPayRun;
    const fromCachedTable = filterFormXIXKarnatakaPayrollRowsForAutofill(cached.rows);
    if (fromCachedTable.length > 0) return fromCachedTable;
  }

  const latest = getLatestCachedPayrollTableRows();
  if (latest?.rows?.length > 0) {
    const fromLatestPayRun = filterFormXIXMPEligiblePayrollRows(latest.rows);
    if (fromLatestPayRun.length > 0) return fromLatestPayRun;
    const fromLatestTable = filterFormXIXKarnatakaPayrollRowsForAutofill(latest.rows);
    if (fromLatestTable.length > 0) return fromLatestTable;
  }
  return [];
}

/** Load Form XIX Karnataka payroll from Catalyst Payroll table only (no Zoho fallback). */
export async function loadFormXIXKarnatakaPayrollRowsForAutofill(monthCandidates = [], options = {}) {
  const months = Array.isArray(monthCandidates) ? monthCandidates : [];
  const timeoutMs = Number(options.timeoutMs) > 0 ? Number(options.timeoutMs) : 45000;
  const tableLoad = await fetchPayrollTableRowsForMonths(months, {
    timeoutMs,
    force: Boolean(options.force),
    zohoFallback: false,
  });
  if (Array.isArray(tableLoad.rows) && tableLoad.rows.length > 0) {
    const payrollMonth = tableLoad.payrollMonth || months[0] || '';
    cacheForm15PayrollTableRows(payrollMonth, tableLoad.rows, tableLoad.meta || null);
    const eligible = filterFormXIXKarnatakaPayrollRowsForAutofill(tableLoad.rows);
    if (eligible.length > 0) return eligible;
  }
  return resolveFormXIXKarnatakaPayrollRowsForAutofill(tableLoad?.rows, months);
}

/** Karnataka Form XIX — payroll column mapping (paid_days, gross_pay, monthly OT date/hours). */
export function resolveFormXIXKarnatakaPayrollFields(payrollRow) {
  if (!payrollRow || payrollRow.fetch_error) {
    return {
      daysWorked: '',
      overtimeDates: '',
      overtimeHoursAndAmount: '',
      grossWages: '',
      deductions: '',
      netWages: '',
    };
  }
  const flat = flattenPayrollEarningColumns(payrollRow);

  const daysWorked = pickKarnatakaPayrollValue(
    flat,
    payrollRow,
    ['paid_days', 'Paid Days', 'days_worked', 'Days Worked', 'paidDays', 'no_of_days_worked'],
    [/^paid_days$/, /paiddays/, /daysworked/, /days_present/, /noofdayspresent/, /no_of_days_present/]
  );

  let overtimeDates = formatKarnatakaPayrollDate(flat.overtime_date);
  if (!overtimeDates) {
    overtimeDates = formatKarnatakaPayrollDate(
      pickKarnatakaPayrollText(
        flat,
        payrollRow,
        [
          'overtime_date',
          'overtime dates',
          'Overtime Date',
          'Overtime Dates',
          'overtime_dates',
          'dates_on_which_overtime_worked',
          'overtime_worked_dates',
        ],
        [/overtime.*date/i, /dates.*overtime/i, /overtime.*worked.*date/i]
      )
    );
  }

  let overtimeHours = '';
  if (flat.overtime_hours !== '' && flat.overtime_hours != null) {
    overtimeHours = String(flat.overtime_hours);
  }
  if (overtimeHours === '') {
    overtimeHours =
      pickKarnatakaPayrollText(
        flat,
        payrollRow,
        [
          'overtime_hours',
          'overtime hours',
          'Overtime Hours',
          'total_overtime_hours',
          'total overtime hours',
          'overtime_hours_total',
          'overtime hours total',
          'ot_hours',
          'overtime_hrs',
          'monthly_overtime_hours',
          'total_ot_hours',
        ],
        [/total.*overtime.*hour/i, /^overtime_hours$/, /^total_overtime_hours$/, /^ot_hours$/, /overtime.*hrs/i]
      ) ||
      pickKarnatakaPayrollValue(
        flat,
        payrollRow,
        [
          'overtime_hours',
          'total_overtime_hours',
          'overtime_hours_total',
          'monthly_overtime_hours',
          'ot_hours',
          'total_ot_hours',
        ],
        [/total.*overtime.*hour/i, /^overtime_hours$/, /^total_overtime_hours$/, /^ot_hours$/]
      );
  }

  let overtimeAmount = '';
  if (flat.overtime !== '' && flat.overtime != null) {
    overtimeAmount = String(flat.overtime);
  }
  if (overtimeAmount === '') {
    overtimeAmount = pickKarnatakaPayrollValue(
      flat,
      payrollRow,
      ['overtime', 'Overtime', 'overtime_wages', 'overtimeWages', 'overtime amount', 'Overtime Amount'],
      [/^overtime$/, /overtime.*wage/i, /overtime.*amount/i]
    );
  }

  return {
    daysWorked,
    overtimeDates,
    overtimeHoursAndAmount: formatKarnatakaOvertimeHoursAndAmount(overtimeHours, overtimeAmount),
    grossWages: (() => {
      const gross = pickGrossPayFromPayrollRow(flat, payrollRow);
      if (gross !== '') return gross;
      return pickNetPayFromPayrollRow(flat, payrollRow);
    })(),
    deductions: pickDeductionsFromPayrollRow(flat, payrollRow),
    netWages: pickNetPayFromPayrollRow(flat, payrollRow),
  };
}

export function applyFormXIXKarnatakaEmployeeToRow(row, emp, headers, helpers = {}) {
  const headerList = mergeFormXIXKarnatakaAutofillHeaders(headers, helpers.parsedFormHeader);
  const out = row && typeof row === 'object' ? { ...row } : {};
  const {
    sanitizeValue = (v) => String(v ?? '').trim(),
    payrollRow = null,
    resolvePayrollFields = null,
    monthCandidates = null,
  } = helpers;

  if (isFormXIXKarnatakaBlankMonthEmployee(emp, monthCandidates)) {
    clearFormXIXKarnatakaAllWageDataOnRow(out, headerList);
    out.__employeeLookupName = sanitizeValue(formatWorkmanNameAndGuardian(emp).split(/\r?\n/)[0]);
    out.__formXIXKarnatakaSkipExport = true;
    return out;
  }

  const hasPayroll = payrollRow && !payrollRow.fetch_error;
  let payroll =
    hasPayroll && typeof resolvePayrollFields === 'function'
      ? resolvePayrollFields(payrollRow)
      : hasPayroll
        ? resolveFormXIXKarnatakaPayrollFields(payrollRow)
        : {};
  payroll = applyFormXIXKarnatakaMonthDefaultPayroll(payroll, emp, monthCandidates);
  const hasMonthDefaults = hasFormXIXKarnatakaMonthDefaultPayrollContext(emp, monthCandidates);
  const wageValue = (value, current = '') => {
    if (value !== '' && (hasPayroll || hasMonthDefaults)) return sanitizeValue(value);
    return String(current ?? '').trim();
  };
  const setSemantic = (matchFn, value) => {
    const text = wageValue(value, '');
    if (!text) return;
    headerList.forEach((header) => {
      if (matchFn(header)) setKarnatakaRowValueForHeader(out, header, text);
    });
    Object.keys(out).forEach((key) => {
      if (String(key).startsWith('__')) return;
      if (matchFn(key)) out[key] = text;
    });
  };

  headerList.forEach((header) => {
    if (isFormXIXKASexIdentificationHeader(header)) {
      setKarnatakaRowValueForHeader(out, header, sanitizeValue(formatFormXIXKASexAndIdentificationMarks(emp)));
      return;
    }
    if (isFormXIXKATokenHeader(header)) {
      setKarnatakaRowValueForHeader(out, header, sanitizeValue(resolveFormXIXKAEmployeeToken(emp)));
      return;
    }
    if (isFormXIXKAWorkmanNameHeader(header)) {
      setKarnatakaRowValueForHeader(out, header, sanitizeValue(formatWorkmanNameAndGuardian(emp)));
      return;
    }
    if (isFormXIXKARateHeader(header)) {
      setKarnatakaRowValueForHeader(out, header, sanitizeValue(FORM_XIX_KA_RATE_DEFAULT));
      return;
    }
    if (isFormXIXKAUnitsHeader(header)) {
      setKarnatakaRowValueForHeader(out, header, '');
      return;
    }
    if (isFormXIXKAOtNilHeader(header)) {
      setKarnatakaRowValueForHeader(out, header, sanitizeValue(FORM_XIX_KA_OT_NIL));
    }
  });
  if (isFormXIXKarnatakaBlankDaysWorkedEmployee(emp)) {
    clearFormXIXKarnatakaSemanticFieldsOnRow(out, headerList, isFormXIXKADaysWorkedHeader);
  } else {
    setSemantic(isFormXIXKADaysWorkedHeader, payroll.daysWorked);
  }
  applyFormXIXKarnatakaOtNilToRow(out, headerList, { nilText: sanitizeValue(FORM_XIX_KA_OT_NIL) });
  if (isFormXIXKarnatakaBlankFooterWageEmployee(emp)) {
    clearFormXIXKarnatakaSemanticFieldsOnRow(
      out,
      headerList,
      (header) =>
        isFormXIXKAGrossHeader(header) ||
        isFormXIXKADeductionsHeader(header) ||
        isFormXIXKANetHeader(header)
    );
  } else {
    setSemantic(isFormXIXKAGrossHeader, payroll.grossWages);
    setSemantic(isFormXIXKADeductionsHeader, payroll.deductions);
    setSemantic(isFormXIXKANetHeader, payroll.netWages);
  }

  Object.keys(out).forEach((key) => {
    if (String(key).startsWith('__')) return;
    if (isFormXIXKARateHeader(key)) out[key] = sanitizeValue(FORM_XIX_KA_RATE_DEFAULT);
    if (isFormXIXKAUnitsHeader(key)) out[key] = '';
    if (isFormXIXKAOtNilHeader(key)) out[key] = sanitizeValue(FORM_XIX_KA_OT_NIL);
  });
  syncFormXIXKarnatakaDeductionsFromGrossNet(out, headerList);
  out.__employeeLookupName = sanitizeValue(formatWorkmanNameAndGuardian(emp).split(/\r?\n/)[0]);
  return out;
}

export function mapFormXIXKarnatakaRowsFromEmployees(employees, headers, helpers = {}) {
  const { resolvePayrollRow = null, resolvePayrollFields = null, parsedFormHeader = null, ...rest } = helpers;
  const hdrs = mergeFormXIXKarnatakaAutofillHeaders(headers, parsedFormHeader);
  const list = Array.isArray(employees) ? employees : [];
  return list.map((empItem, rowIndex) => {
    const emp = empItem?.Employee || empItem?.employee || empItem;
    const payrollRow = typeof resolvePayrollRow === 'function' ? resolvePayrollRow(emp, rowIndex) : null;
    return applyFormXIXKarnatakaEmployeeToRow({}, emp, hdrs, {
      ...rest,
      parsedFormHeader,
      payrollRow: payrollRow && !payrollRow.fetch_error ? payrollRow : null,
      resolvePayrollFields,
    });
  });
}

export function enrichFormXIXKarnatakaStaticFieldRows(mappedData, headers) {
  if (!Array.isArray(mappedData) || mappedData.length === 0) return 0;
  const hdrs = resolveFormXIXKarnatakaWageTableHeaders(headers);
  const rateHdrs = hdrs.filter(isFormXIXKARateHeader);
  const otNilHdrs = hdrs.filter(isFormXIXKAOtNilHeader);
  if (rateHdrs.length === 0 && otNilHdrs.length === 0) return 0;
  let hits = 0;
  mappedData.forEach((row) => {
    if (!row || typeof row !== 'object') return;
    const rateKeys = new Set(rateHdrs);
    const otKeys = new Set(otNilHdrs);
    Object.keys(row).forEach((key) => {
      if (String(key).startsWith('__')) return;
      if (isFormXIXKARateHeader(key)) rateKeys.add(key);
      if (isFormXIXKAUnitsHeader(key)) row[key] = '';
      if (isFormXIXKAOtNilHeader(key)) otKeys.add(key);
    });
    rateKeys.forEach((header) => {
      row[header] = FORM_XIX_KA_RATE_DEFAULT;
    });
    otKeys.forEach((header) => {
      row[header] = FORM_XIX_KA_OT_NIL;
    });
    if (rateKeys.size > 0 || otKeys.size > 0) hits += 1;
  });
  return hits;
}

export function enrichFormXIXKarnatakaPayrollRows(mappedData, employees, headers, helpers = {}) {
  const hdrs = mergeFormXIXKarnatakaAutofillHeaders(headers, helpers.parsedFormHeader);
  const {
    resolvePayrollRow = null,
    resolvePayrollFields = null,
    sanitizeValue = (v) => String(v ?? '').trim(),
    overwrite = true,
    parsedFormHeader = null,
    monthCandidates = null,
  } = helpers;
  if (!Array.isArray(mappedData) || mappedData.length === 0) return 0;
  let hits = 0;
  mappedData.forEach((row, rowIndex) => {
    const empItem = employees[rowIndex];
    const emp = empItem?.Employee || empItem?.employee || empItem;
    const payrollRow =
      typeof resolvePayrollRow === 'function' ? resolvePayrollRow(emp, row, rowIndex) : null;
    const merged = applyFormXIXKarnatakaEmployeeToRow(row, emp, hdrs, {
      sanitizeValue,
      payrollRow: payrollRow && !payrollRow.fetch_error ? payrollRow : null,
      resolvePayrollFields,
      parsedFormHeader,
      monthCandidates,
    });
    hdrs.forEach((header) => {
      const prev = String(getKarnatakaRowValueForHeader(row, header) ?? '').trim();
      const next = String(getKarnatakaRowValueForHeader(merged, header) ?? '').trim();
      if (!overwrite && prev !== '' && next !== '') {
        setKarnatakaRowValueForHeader(merged, header, prev);
        return;
      }
    });
    applyFormXIXKarnatakaOtNilToRow(merged, hdrs);
    Object.assign(row, merged);
    const payrollFields = applyFormXIXKarnatakaMonthDefaultPayroll(
      payrollRow && !payrollRow.fetch_error
        ? resolveFormXIXKarnatakaPayrollFields(payrollRow)
        : {},
      emp,
      monthCandidates
    );
    if (
      payrollFields.grossWages ||
      payrollFields.deductions ||
      payrollFields.netWages ||
      payrollFields.daysWorked
    ) {
      hits += 1;
    }
  });
  return hits;
}

export function rowHasMeaningfulFormXIXKarnatakaExportData(row, headers) {
  if (!row || typeof row !== 'object') return false;
  const hdrs = resolveFormXIXKarnatakaWageTableHeaders(headers);
  return hdrs.some((header) => getKarnatakaRowValueForHeader(row, header) !== '');
}

export function overlayFormXIXKarnatakaUserEditsOntoRows(mappedRows, tableRows, headers) {
  const hdrs = resolveFormXIXKarnatakaWageTableHeaders(headers);
  const savedRows = Array.isArray(tableRows) ? tableRows : [];
  return (Array.isArray(mappedRows) ? mappedRows : []).map((row, index) => {
    const saved = savedRows[index];
    if (!saved || typeof saved !== 'object' || !rowHasMeaningfulFormXIXKarnatakaExportData(saved, hdrs)) {
      return row;
    }
    const merged = { ...row };
    hdrs.forEach((header) => {
      if (isFormXIXKAOtNilHeader(header)) return;
      const value = getKarnatakaRowValueForHeader(saved, header);
      if (value !== '') merged[header] = saved[header] ?? value;
    });
    applyFormXIXKarnatakaOtNilToRow(merged, hdrs);
    return merged;
  });
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

export function writeFormXIXKarnatakaHeaderFieldsToWorksheet(worksheet, headerFormData = {}, parsedFormHeader = null) {
  if (!worksheet || !headerFormData || typeof headerFormData !== 'object') return;
  const parsedFields = Array.isArray(parsedFormHeader?.fields) ? parsedFormHeader.fields : [];
  const writeAt = (row, col, value) => {
    const text = String(value ?? '').trim();
    if (!text || row < 1 || col < 1) return;
    worksheet.getCell(row, col).value = text;
  };
  FORM_XIX_KA_HEADER_SPECS.forEach((spec) => {
    const val = headerFormData[spec.key];
    if (val == null || String(val).trim() === '') return;
    const parsedField = parsedFields.find((f) => f.key === spec.key);
    if (parsedField?.labelRow != null) {
      writeAt(parsedField.labelRow + 1, FORM_XIX_KA_DEFAULT_VALUE_COL, val);
      return;
    }
    for (let r = 1; r <= 40; r += 1) {
      for (let c = 1; c <= 14; c += 1) {
        const raw = excelCellValueToString(worksheet.getCell(r, c)?.value).trim();
        if (!raw || (!spec.match.test(raw) && !spec.match.test(formXIXAPHeaderNorm(raw)))) continue;
        writeAt(r, FORM_XIX_KA_DEFAULT_VALUE_COL, val);
        return;
      }
    }
  });
  FORM_XIX_KA_EMPLOYEE_HEADER_SPECS.forEach((spec) => {
    const val = headerFormData[spec.key];
    if (val == null || String(val).trim() === '') return;
    const parsedField = parsedFields.find((f) => f.key === spec.key);
    if (parsedField?.labelRow != null) {
      writeAt(parsedField.labelRow + 1, FORM_XIX_KA_DEFAULT_VALUE_COL, val);
      return;
    }
    for (let r = 1; r <= 40; r += 1) {
      for (let c = 1; c <= 14; c += 1) {
        const raw = excelCellValueToString(worksheet.getCell(r, c)?.value).trim();
        if (!raw || (!spec.match.test(raw) && !spec.match.test(formXIXAPHeaderNorm(raw)))) continue;
        writeAt(r, FORM_XIX_KA_DEFAULT_VALUE_COL, val);
        return;
      }
    }
  });
}

function resolveTableColumns(parsedFormHeader) {
  const fromHeader = Array.isArray(parsedFormHeader?.tableColumns)
    ? parsedFormHeader.tableColumns.filter((h) => h && !isFormXIXKAIdentityTableHeader(h))
    : [];
  if (fromHeader.length >= FORM_XIX_KA_EXCEL_WAGE_HEADERS.length) return fromHeader;
  const startCol = Number(parsedFormHeader?.tableStartCol ?? 0);
  return FORM_XIX_KA_EXCEL_WAGE_HEADERS.map((header, i) => ({ header, col: startCol + i }));
}

function resolveKarnatakaTableColumnsFromWorksheet(worksheet, parsedFormHeader) {
  const parsedCols = resolveTableColumns(parsedFormHeader);
  if (!worksheet) return parsedCols;

  const maxR = 45;
  const maxC = 20;
  const wageHdrs = FORM_XIX_KA_EXCEL_WAGE_HEADERS;
  for (let r = 1; r <= maxR; r += 1) {
    for (let c = 1; c <= maxC; c += 1) {
      const raw = excelCellValueToString(worksheet.getCell(r, c)?.value).trim();
      if (!isFormXIXKADaysWorkedHeader(raw)) continue;
      return wageHdrs.map((header, i) => ({
        header,
        col: c - 1 + i,
        headerRow: r,
      }));
    }
  }
  return parsedCols;
}

function resolveKarnatakaDataRowFromWorksheet(worksheet, parsedFormHeader, columns) {
  const daysCol = columns.find((col) => isFormXIXKADaysWorkedHeader(col.header));
  const headerRow = Number(daysCol?.headerRow ?? 0);
  if (worksheet && headerRow > 0) {
    const col1 = Number(daysCol.col) + 1;
    for (let r = headerRow + 1; r <= headerRow + 5; r += 1) {
      const marker = excelCellValueToString(worksheet.getCell(r, col1)?.value).trim();
      if (/^[1-5]$/.test(marker) || /^\d+$/.test(marker)) {
        return r + 1;
      }
    }
    return headerRow + 2;
  }
  const parsedRow = Number(parsedFormHeader?.dataStartIndex);
  if (Number.isFinite(parsedRow) && parsedRow >= 0) {
    return parsedRow + 1;
  }
  return 17;
}

function resolveKarnatakaFooterColumnsFromWorksheet(worksheet, parsedFormHeader, columns, dataRow) {
  const parsedFooter = resolveFooterColumns(parsedFormHeader);
  if (!worksheet) return parsedFooter;

  const footerHdrs = FORM_XIX_KA_FOOTER_HEADERS;
  const headerRow = Number(columns.find((col) => isFormXIXKADaysWorkedHeader(col.header))?.headerRow ?? 0);
  const scanFrom = Math.max(headerRow, dataRow - 1);
  for (let r = scanFrom; r <= scanFrom + 10; r += 1) {
    const matched = [];
    for (let c = 1; c <= 20; c += 1) {
      const cell = excelCellValueToString(worksheet.getCell(r, c)?.value).trim();
      if (!cell) continue;
      if (isFormXIXKAGrossHeader(cell)) matched.push({ header: footerHdrs[0], col: c - 1, labelRow: r - 1, valueRow: r + 1 });
      else if (isFormXIXKADeductionsHeader(cell)) matched.push({ header: footerHdrs[1], col: c - 1, labelRow: r - 1, valueRow: r + 1 });
      else if (isFormXIXKANetHeader(cell)) matched.push({ header: footerHdrs[2], col: c - 1, labelRow: r - 1, valueRow: r + 1 });
    }
    if (matched.length >= 2) return matched;
  }
  return parsedFooter;
}

function resolveFooterColumns(parsedFormHeader) {
  const fromHeader = Array.isArray(parsedFormHeader?.footerColumns) ? parsedFormHeader.footerColumns : [];
  if (fromHeader.length > 0) return fromHeader;
  return FORM_XIX_KA_FOOTER_HEADERS.map((header, i) => ({
    header,
    col: 1 + i * 2,
    labelRow: Number(parsedFormHeader?.footerRowIndex ?? 18),
  }));
}

/** Resolve where footer amounts are written — directly below each label (same column). */
function buildKarnatakaFooterValuePositions(parsedFormHeader, worksheet, tableColumns, dataRow) {
  const footerColumns =
    Array.isArray(parsedFormHeader?.footerColumns) && parsedFormHeader.footerColumns.length >= 2
      ? parsedFormHeader.footerColumns
      : worksheet
        ? resolveKarnatakaFooterColumnsFromWorksheet(worksheet, parsedFormHeader, tableColumns, dataRow)
        : resolveFooterColumns(parsedFormHeader);

  const footerValueRow0 =
    parsedFormHeader?.footerValueRow != null
      ? Number(parsedFormHeader.footerValueRow)
      : footerColumns[0]?.labelRow != null
        ? Number(footerColumns[0].labelRow) + 1
        : Number(parsedFormHeader?.footerRowIndex ?? 18) + 1;

  const positions = [];
  const seen = new Set();

  const pushPos = (header, row, col) => {
    if (!header || row == null || col == null || row < 1 || col < 1) return;
    const cellRef = formXIXKAToCellRef(row, col);
    const key = `${header}|${cellRef}`;
    if (seen.has(key)) return;
    seen.add(key);
    positions.push({ header, row, col, cellRef });
  };

  footerColumns.forEach(({ header, col, labelRow }) => {
    const col0 = Number(col);
    const labelRow0 = labelRow != null ? Number(labelRow) : Number(parsedFormHeader?.footerRowIndex ?? 18);
    const belowExcelRow =
      Number.isFinite(footerValueRow0) && footerValueRow0 >= 0
        ? footerValueRow0 + 1
        : labelRow0 + 2;

    // Official Form XIX KA layout: Gross / Deductions / Actual amounts sit under their labels.
    // Do not also write beside the label — that duplicated net pay and broke PDF column alignment.
    pushPos(header, belowExcelRow, col0 + 1);
  });

  return positions;
}

export function applyFormXIXKarnatakaFooterFieldsToRow(row, payrollRow, headers, emp = null, monthCandidates = null) {
  if (!row || typeof row !== 'object' || !payrollRow || payrollRow.fetch_error) return row;
  const out = { ...row };
  const hdrs = resolveFormXIXKarnatakaWageTableHeaders(headers);
  if (emp && isFormXIXKarnatakaBlankMonthEmployee(emp, monthCandidates)) {
    clearFormXIXKarnatakaAllWageDataOnRow(out, hdrs);
    return out;
  }
  if (emp && isFormXIXKarnatakaBlankFooterWageEmployee(emp)) {
    clearFormXIXKarnatakaSemanticFieldsOnRow(
      out,
      hdrs,
      (header) =>
        isFormXIXKAGrossHeader(header) ||
        isFormXIXKADeductionsHeader(header) ||
        isFormXIXKANetHeader(header)
    );
    return out;
  }
  const payroll = applyFormXIXKarnatakaMonthDefaultPayroll(
    resolveFormXIXKarnatakaPayrollFields(payrollRow),
    emp,
    monthCandidates
  );
  const footerValues = [payroll.grossWages, payroll.deductions, payroll.netWages];
  FORM_XIX_KA_FOOTER_HEADERS.forEach((footerHeader, index) => {
    const value = String(footerValues[index] ?? '').trim();
    if (!value) return;
    out[footerHeader] = value;
    hdrs.forEach((header) => {
      if (headerMatchesKarnatakaFooter(header, footerHeader)) {
        out[header] = value;
      }
    });
  });
  syncFormXIXKarnatakaDeductionsFromGrossNet(out, hdrs);
  return out;
}

export function writeFormXIXKarnatakaFooterRowToWorksheet(worksheet, row, parsedFormHeader) {
  if (!worksheet || !row || typeof row !== 'object') return;
  const tableColumns = resolveKarnatakaTableColumnsFromWorksheet(worksheet, parsedFormHeader);
  const dataRow = resolveKarnatakaDataRowFromWorksheet(worksheet, parsedFormHeader, tableColumns);
  const footerPositions = buildKarnatakaFooterValuePositions(
    parsedFormHeader,
    worksheet,
    tableColumns,
    dataRow
  );
  footerPositions.forEach(({ header, row: writeRow, col: writeCol }) => {
    const value = resolveKarnatakaExportCellValue(row, header);
    if (value === '') return;
    worksheet.getCell(writeRow, writeCol).value = value;
  });
}

export function writeFormXIXKarnatakaTableRowToWorksheet(worksheet, row, parsedFormHeader, dataRowIndex) {
  if (!worksheet || !row || typeof row !== 'object') return;
  const columns = resolveKarnatakaTableColumnsFromWorksheet(worksheet, parsedFormHeader);
  const dataRow =
    dataRowIndex != null && Number.isFinite(Number(dataRowIndex))
      ? Number(dataRowIndex) + 1
      : resolveKarnatakaDataRowFromWorksheet(worksheet, parsedFormHeader, columns);
  columns.forEach(({ header, col }) => {
    const value = resolveKarnatakaExportCellValue(row, header);
    if (value === '') return;
    worksheet.getCell(dataRow, Number(col) + 1).value = value;
  });
  writeFormXIXKarnatakaFooterRowToWorksheet(worksheet, row, parsedFormHeader);
}

function buildEmployeeHeaderFormData(headerFormData, employeeRow, emp) {
  const base = headerFormData && typeof headerFormData === 'object' ? { ...headerFormData } : {};
  const hdrs = FORM_XIX_KA_WAGE_TABLE_HEADERS;
  const workmanHdr = hdrs.find(isFormXIXKAWorkmanNameHeader);
  const tokenHdr = hdrs.find(isFormXIXKATokenHeader);
  const sexHdr = hdrs.find(isFormXIXKASexIdentificationHeader);
  const workmanFromRow = workmanHdr ? getKarnatakaRowValueForHeader(employeeRow, workmanHdr) : '';
  const tokenFromRow = tokenHdr ? getKarnatakaRowValueForHeader(employeeRow, tokenHdr) : '';
  const sexFromRow = sexHdr ? getKarnatakaRowValueForHeader(employeeRow, sexHdr) : '';
  const workman = workmanFromRow || formatWorkmanNameAndGuardian(emp);
  if (workman) base.form_xix_ka_workman = workman;
  const token = tokenFromRow || resolveFormXIXKAEmployeeToken(emp);
  if (token) base.form_xix_ka_token = token;
  const sexIdentification = sexFromRow || formatFormXIXKASexAndIdentificationMarks(emp);
  if (sexIdentification) base.form_xix_ka_sex_identification = sexIdentification;
  return base;
}

export async function buildFormXIXKarnatakaWorkbookWithTemplateStyles({
  templateArrayBuffer,
  mappedData,
  headersToUse,
  parsedFormHeader,
  formFileName,
  headerFormData,
  employeesOverride,
  resolvePayrollRow,
  monthCandidates = null,
}) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(templateArrayBuffer);
  const worksheet = workbook.worksheets[0];
  if (!worksheet) throw new Error('Template worksheet not found.');

  const hdrs = resolveFormXIXKarnatakaWageTableHeaders(headersToUse);
  let exportRows = (Array.isArray(mappedData) ? mappedData : []).filter((row) =>
    rowHasMeaningfulFormXIXKarnatakaExportData(row, hdrs)
  );
  enrichFormXIXKarnatakaStaticFieldRows(exportRows, hdrs);
  const employees = Array.isArray(employeesOverride) ? employeesOverride : [];
  const employeeRow = exportRows[0] || null;
  const emp = employees[0]?.Employee || employees[0]?.employee || employees[0] || null;
  const payrollRow =
    typeof resolvePayrollRow === 'function' && emp ? resolvePayrollRow(emp) : null;
  let rowToWrite = employeeRow;
  if (employeeRow && payrollRow && !payrollRow.fetch_error) {
    rowToWrite = applyFormXIXKarnatakaFooterFieldsToRow(
      employeeRow,
      payrollRow,
      hdrs,
      emp,
      monthCandidates
    );
  }
  if (rowToWrite) syncFormXIXKarnatakaDeductionsFromGrossNet(rowToWrite, hdrs);
  const mergedHeader = buildEmployeeHeaderFormData(headerFormData, rowToWrite, emp);

  writeFormXIXKarnatakaHeaderFieldsToWorksheet(worksheet, mergedHeader, parsedFormHeader);
  if (rowToWrite) {
    writeFormXIXKarnatakaTableRowToWorksheet(
      worksheet,
      rowToWrite,
      { ...parsedFormHeader, dataStartIndex: parsedFormHeader?.dataStartIndex },
      parsedFormHeader?.dataStartIndex
    );
  }

  // Match Excel print / Download PDF to the on-screen Form XIX A4 portrait layout.
  worksheet.pageSetup = {
    ...(worksheet.pageSetup || {}),
    paperSize: 9,
    orientation: 'portrait',
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 1,
    horizontalCentered: true,
  };
  try {
    worksheet.pageSetup.printArea = 'A1:K28';
  } catch (_) {
    /* template may already define print area */
  }

  const out = await workbook.xlsx.writeBuffer();
  const fileName =
    formFileName ||
    parsedFormHeader?.title?.replace(/[^a-zA-Z0-9]/g, '_') ||
    `Form_XIX_Karnataka_${Date.now()}.xlsx`;
  return { blob: new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), fileName };
}

function resolveEmployeeDownloadBaseName(row, fallbackIndex = 0) {
  const raw = String(row?.__employeeLookupName ?? '').trim();
  const slug = raw
    .replace(/[^\w\s.-]/g, '')
    .replace(/\s+/g, '_')
    .replace(/_+/g, '_')
    .replace(/^_|_$/g, '')
    .slice(0, 48);
  return slug || `Employee_${fallbackIndex + 1}`;
}

const formXIXKASanitizeExportText = (value) =>
  String(value ?? '')
    .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '')
    .trim();

const formXIXKAEscapeXml = (value) =>
  String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');

const formXIXKAColToLetter = (col) => {
  let result = '';
  let n = col;
  while (n > 0) {
    const rem = (n - 1) % 26;
    result = String.fromCharCode(65 + rem) + result;
    n = Math.floor((n - 1) / 26);
  }
  return result;
};

const formXIXKAToCellRef = (row, col) => `${formXIXKAColToLetter(col)}${row}`;

const formXIXKAUpsertInlineStrCell = (sheetXml, cellRef, value) => {
  const text = formXIXKAEscapeXml(formXIXKASanitizeExportText(value));
  const cellXml = text
    ? `<c r="${cellRef}" t="inlineStr"><is><t xml:space="preserve">${text}</t></is></c>`
    : `<c r="${cellRef}"/>`;
  const cellRe = new RegExp(`<c\\s+r="${cellRef}"[^>]*(?:/>|>[\\s\\S]*?</c>)`, 'i');
  if (cellRe.test(sheetXml)) {
    return sheetXml.replace(cellRe, cellXml);
  }
  const rowNum = cellRef.replace(/^[A-Z]+/i, '');
  const rowRe = new RegExp(`(<row\\s+r="${rowNum}"[^>]*>)([\\s\\S]*?)(</row>)`, 'i');
  if (!rowRe.test(sheetXml)) return sheetXml;
  return sheetXml.replace(rowRe, `$1$2${cellXml}$3`);
};

const resolveFormXIXKAWorksheetEntry = (zipFiles) =>
  Object.keys(zipFiles)
    .filter((name) => /^xl\/worksheets\/sheet\d+\.xml$/i.test(name))
    .sort((a, b) => a.localeCompare(b, undefined, { numeric: true }))[0] || null;

function resolveFormXIXKarnatakaFastExportPositions(worksheet, parsedFormHeader) {
  const employeeHeaderPositions = [];
  const seenHeaderKeys = new Set();
  const parsedFields = Array.isArray(parsedFormHeader?.fields) ? parsedFormHeader.fields : [];

  const pushEmployeeHeader = (key, row, col) => {
    if (!key || row == null || col == null || seenHeaderKeys.has(key)) return;
    seenHeaderKeys.add(key);
    employeeHeaderPositions.push({ key, row, col, cellRef: formXIXKAToCellRef(row, col) });
  };

  FORM_XIX_KA_EMPLOYEE_HEADER_SPECS.forEach((spec) => {
    const parsedField = parsedFields.find((f) => f.key === spec.key);
    if (parsedField?.labelRow != null) {
      pushEmployeeHeader(
        spec.key,
        (parsedField.valueRow ?? parsedField.labelRow) + 1,
        parsedField.valueCol != null ? parsedField.valueCol + 1 : FORM_XIX_KA_DEFAULT_VALUE_COL
      );
      return;
    }
    if (!worksheet) return;
    for (let r = 1; r <= 40; r += 1) {
      for (let c = 1; c <= 14; c += 1) {
        const raw = excelCellValueToString(worksheet.getCell(r, c)?.value).trim();
        if (!raw || (!spec.match.test(raw) && !spec.match.test(formXIXAPHeaderNorm(raw)))) continue;
        pushEmployeeHeader(spec.key, r, FORM_XIX_KA_DEFAULT_VALUE_COL);
        return;
      }
    }
  });

  const tableColumns = worksheet
    ? resolveKarnatakaTableColumnsFromWorksheet(worksheet, parsedFormHeader)
    : resolveTableColumns(parsedFormHeader);
  const dataRow = worksheet
    ? resolveKarnatakaDataRowFromWorksheet(worksheet, parsedFormHeader, tableColumns)
    : Number(parsedFormHeader?.dataStartIndex ?? 14) + 1;
  const tablePositions = tableColumns.map(({ header, col }) => ({
    header,
    row: dataRow,
    col: Number(col) + 1,
    cellRef: formXIXKAToCellRef(dataRow, Number(col) + 1),
  }));

  const footerPositions = buildKarnatakaFooterValuePositions(
    parsedFormHeader,
    worksheet,
    tableColumns,
    dataRow
  );

  return { employeeHeaderPositions, tablePositions, footerPositions };
}

function finalizeFormXIXKarnatakaExportRowForEmployee(exportRow, emp, hdrs, helpers = {}) {
  if (!exportRow || !emp) return exportRow;
  const {
    sanitizeValue = (v) => String(v ?? '').trim(),
    payrollRow = null,
    monthCandidates = null,
  } = helpers;
  if (isFormXIXKarnatakaBlankMonthEmployee(emp, monthCandidates)) {
    const cleared = { ...exportRow };
    clearFormXIXKarnatakaAllWageDataOnRow(cleared, hdrs);
    cleared.__formXIXKarnatakaSkipExport = true;
    return cleared;
  }
  const needsBlankRules =
    isFormXIXKarnatakaBlankDaysWorkedEmployee(emp) ||
    isFormXIXKarnatakaBlankFooterWageEmployee(emp);
  if (!needsBlankRules) return exportRow;
  let row = applyFormXIXKarnatakaEmployeeToRow(exportRow, emp, hdrs, {
    sanitizeValue,
    payrollRow: payrollRow && !payrollRow.fetch_error ? payrollRow : null,
    resolvePayrollFields: resolveFormXIXKarnatakaPayrollFields,
    monthCandidates,
  });
  if (payrollRow && !payrollRow.fetch_error) {
    row = applyFormXIXKarnatakaFooterFieldsToRow(row, payrollRow, hdrs, emp, monthCandidates);
  }
  return row;
}

const clearFormXIXKarnatakaPerEmployeeValueCells = (worksheet, positions) => {
  if (!worksheet || !positions) return;
  const all = [
    ...(positions.employeeHeaderPositions || []),
    ...(positions.tablePositions || []),
    ...(positions.footerPositions || []),
  ];
  all.forEach((pos) => {
    if (pos?.row != null && pos?.col != null) {
      worksheet.getCell(pos.row, pos.col).value = '';
    }
  });
};

const patchFormXIXKarnatakaFastSheetXml = (baseSheetXml, mergedHeaderData, exportRow, positions) => {
  let sheetXml = baseSheetXml;
  (positions.employeeHeaderPositions || []).forEach((pos) => {
    const value = mergedHeaderData?.[pos.key];
    if (value != null && String(value).trim() !== '') {
      sheetXml = formXIXKAUpsertInlineStrCell(sheetXml, pos.cellRef, value);
    }
  });
  (positions.tablePositions || []).forEach((pos) => {
    const value = resolveKarnatakaExportCellValue(exportRow, pos.header);
    if (value !== '') {
      sheetXml = formXIXKAUpsertInlineStrCell(sheetXml, pos.cellRef, value);
    }
  });
  (positions.footerPositions || []).forEach((pos) => {
    const value = resolveKarnatakaExportCellValue(exportRow, pos.header);
    if (value !== '') {
      sheetXml = formXIXKAUpsertInlineStrCell(sheetXml, pos.cellRef, value);
    }
  });
  return sheetXml;
};

const buildFormXIXKarnatakaFastXlsxBytes = async (fastTemplate, mergedHeaderData, exportRow) => {
  const { sheetEntry, baseSheetXml, staticFiles, positions } = fastTemplate;
  const sheetXml = patchFormXIXKarnatakaFastSheetXml(baseSheetXml, mergedHeaderData, exportRow, positions);
  const entryZip = new JSZip();
  Object.entries(staticFiles).forEach(([path, data]) => {
    entryZip.file(path, data);
  });
  entryZip.file(sheetEntry, sheetXml);
  return entryZip.generateAsync({ type: 'uint8array', compression: 'STORE' });
};

export function canUseFormXIXKarnatakaFastExport(parsedFormHeader) {
  return !!parsedFormHeader?.formXIXKarnatakaTableLayout;
}

async function prepareFormXIXKarnatakaFastZipTemplate({
  templateArrayBuffer,
  parsedFormHeader,
  headerFormData,
}) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(templateArrayBuffer);
  const worksheet = workbook.worksheets[0];
  if (!worksheet) throw new Error('Template worksheet not found.');

  const staticHeaderData = { ...(headerFormData || {}) };
  FORM_XIX_KA_EMPLOYEE_HEADER_SPECS.forEach((spec) => {
    delete staticHeaderData[spec.key];
  });
  writeFormXIXKarnatakaHeaderFieldsToWorksheet(worksheet, staticHeaderData, parsedFormHeader);

  worksheet.pageSetup = {
    ...(worksheet.pageSetup || {}),
    paperSize: 9,
    orientation: 'portrait',
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 1,
    horizontalCentered: true,
  };

  const positions = resolveFormXIXKarnatakaFastExportPositions(worksheet, parsedFormHeader);
  clearFormXIXKarnatakaPerEmployeeValueCells(worksheet, positions);

  const preparedBuffer = await workbook.xlsx.writeBuffer();
  const templateZip = await JSZip.loadAsync(preparedBuffer);
  const sheetEntry = resolveFormXIXKAWorksheetEntry(templateZip.files);
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

const FORM_XIX_KA_FAST_ZIP_BATCH = 12;

function shouldSkipFormXIXKarnatakaEmployeeExport(exportRow, emp, monthCandidates) {
  if (exportRow?.__formXIXKarnatakaSkipExport) return true;
  return isFormXIXKarnatakaBlankMonthEmployee(emp, monthCandidates);
}

async function buildFormXIXKarnatakaFastZipDownload({
  exportRows,
  fastTemplate,
  baseHeaderData,
  parsedFormHeader,
  formFileName,
  employees = [],
  monthCandidates = null,
}) {
  const zip = new JSZip();
  const usedNames = new Map();
  for (let i = 0; i < exportRows.length; i += FORM_XIX_KA_FAST_ZIP_BATCH) {
    const batch = exportRows.slice(i, i + FORM_XIX_KA_FAST_ZIP_BATCH);
    const batchBytes = await Promise.all(
      batch.map(async (exportRow, batchIndex) => {
        const index = i + batchIndex;
        const emp = employees[index]?.Employee || employees[index]?.employee || employees[index] || null;
        if (shouldSkipFormXIXKarnatakaEmployeeExport(exportRow, emp, monthCandidates)) {
          return null;
        }
        const mergedHeaderData = buildEmployeeHeaderFormData(baseHeaderData, exportRow, emp);
        const xlsxBytes = await buildFormXIXKarnatakaFastXlsxBytes(
          fastTemplate,
          mergedHeaderData,
          exportRow
        );
        return { xlsxBytes, exportRow, index };
      })
    );
    batchBytes.forEach((entry) => {
      if (!entry) return;
      const { xlsxBytes, exportRow, index } = entry;
      const baseName = resolveEmployeeDownloadBaseName(exportRow, index);
      const count = usedNames.get(baseName) || 0;
      usedNames.set(baseName, count + 1);
      const suffix = count > 0 ? `_${count + 1}` : '';
      zip.file(`Form_XIX_Karnataka_${baseName}${suffix}.xlsx`, xlsxBytes);
    });
    if (i > 0) {
      await new Promise((resolve) => setTimeout(resolve, 0));
    }
  }
  const zipBase = String(formFileName || parsedFormHeader?.title || 'Form_XIX_Karnataka')
    .replace(/\.xlsx?$/i, '')
    .replace(/[^a-zA-Z0-9._-]+/g, '_');
  return {
    blob: await zip.generateAsync({ type: 'blob', compression: 'STORE' }),
    fileName: `${zipBase}_Employees.zip`,
  };
}

export function resolveFormXIXKarnatakaExportRows(mappedData, headers, employeesOverride = null, helpers = {}) {
  const hdrs = resolveFormXIXKarnatakaWageTableHeaders(headers);
  const tableRows = Array.isArray(mappedData) ? mappedData : [];
  const employees = Array.isArray(employeesOverride) ? employeesOverride : [];
  const {
    sanitizeValue = (v) => String(v ?? '').trim(),
    resolvePayrollRow = null,
    resolvePayrollFields = resolveFormXIXKarnatakaPayrollFields,
  } = helpers;
  if (employees.length > 0) {
    const mapHelpers = { sanitizeValue, resolvePayrollFields };
    if (typeof resolvePayrollRow === 'function') {
      mapHelpers.resolvePayrollRow = resolvePayrollRow;
    }
    const fromEmployees = mapFormXIXKarnatakaRowsFromEmployees(employees, hdrs, mapHelpers);
    enrichFormXIXKarnatakaStaticFieldRows(fromEmployees, hdrs);
    return overlayFormXIXKarnatakaUserEditsOntoRows(fromEmployees, tableRows, hdrs);
  }
  const rows = tableRows.filter((row) => rowHasMeaningfulFormXIXKarnatakaExportData(row, hdrs));
  enrichFormXIXKarnatakaStaticFieldRows(rows, hdrs);
  return rows;
}

export function triggerFormXIXKarnatakaZipDownload(blob, fileName) {
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

export async function buildFormXIXKarnatakaPerEmployeeDownload({
  templateArrayBuffer,
  mappedData,
  headersToUse,
  parsedFormHeader,
  formFileName,
  headerFormData,
  employeesOverride,
  resolvePayrollRow,
  monthCandidates = null,
}) {
  if (!templateArrayBuffer) {
    throw new Error('Original Form XIX Karnataka template could not be loaded. Open Autofill again, then Download.');
  }

  const hdrs = resolveFormXIXKarnatakaWageTableHeaders(headersToUse);
  const employees = Array.isArray(employeesOverride) ? employeesOverride : [];
  const tableRows = Array.isArray(mappedData) ? mappedData : [];
  const exportHelpers = {
    sanitizeValue: (v) => String(v ?? '').trim(),
    resolvePayrollRow: typeof resolvePayrollRow === 'function' ? resolvePayrollRow : null,
    resolvePayrollFields: resolveFormXIXKarnatakaPayrollFields,
  };
  const fromEmployees =
    employees.length > 0
      ? resolveFormXIXKarnatakaExportRows(tableRows, hdrs, employees, exportHelpers)
      : [];
  const fromTable = tableRows.filter((row) => rowHasMeaningfulFormXIXKarnatakaExportData(row, hdrs));
  let exportRows;
  if (fromTable.length > 0) {
    exportRows = fromTable.map((row) => ({ ...row }));
  } else if (employees.length > 0) {
    exportRows =
      fromEmployees.length > 0
        ? fromEmployees
        : mapFormXIXKarnatakaRowsFromEmployees(employees, hdrs, exportHelpers);
    if (exportRows.length < employees.length) {
      exportRows = mapFormXIXKarnatakaRowsFromEmployees(employees, hdrs, exportHelpers);
    }
    exportRows = overlayFormXIXKarnatakaUserEditsOntoRows(exportRows, tableRows, hdrs);
  } else {
    exportRows = [];
  }
  enrichFormXIXKarnatakaStaticFieldRows(exportRows, hdrs);

  // Keep Deductions, any = Gross − Actual even when the modal already has payroll values (often "0").
  exportRows.forEach((row) => syncFormXIXKarnatakaDeductionsFromGrossNet(row, hdrs));

  const gridHasPayrollValues =
    fromTable.length > 0 &&
    fromTable.some((row) => {
      const daysHdr = hdrs.find(isFormXIXKADaysWorkedHeader);
      const grossHdr = hdrs.find(isFormXIXKAGrossHeader);
      return (
        (daysHdr && getKarnatakaRowValueForHeader(row, daysHdr) !== '') ||
        (grossHdr && getKarnatakaRowValueForHeader(row, grossHdr) !== '')
      );
    });

  if (
    exportRows.length > 0 &&
    typeof resolvePayrollRow === 'function' &&
    employees.length > 0 &&
    !gridHasPayrollValues
  ) {
    exportRows = exportRows.map((exportRow, index) => {
      const emp = employees[index]?.Employee || employees[index]?.employee || employees[index] || null;
      const payrollRow = resolvePayrollRow(emp);
      let merged = applyFormXIXKarnatakaEmployeeToRow(exportRow, emp, hdrs, {
        sanitizeValue: exportHelpers.sanitizeValue,
        payrollRow: payrollRow && !payrollRow.fetch_error ? payrollRow : null,
        resolvePayrollFields: resolveFormXIXKarnatakaPayrollFields,
        monthCandidates,
      });
      if (payrollRow && !payrollRow.fetch_error) {
        merged = applyFormXIXKarnatakaFooterFieldsToRow(merged, payrollRow, hdrs, emp, monthCandidates);
      }
      return merged;
    });
    enrichFormXIXKarnatakaPayrollRows(exportRows, employees, hdrs, {
      sanitizeValue: exportHelpers.sanitizeValue,
      overwrite: false,
      resolvePayrollRow: (emp) => resolvePayrollRow(emp),
      resolvePayrollFields: resolveFormXIXKarnatakaPayrollFields,
      monthCandidates,
    });
  } else if (
    exportRows.length > 0 &&
    typeof resolvePayrollRow === 'function' &&
    employees.length > 0 &&
    gridHasPayrollValues
  ) {
    const grossHdr = hdrs.find(isFormXIXKAGrossHeader);
    exportRows = exportRows.map((exportRow, index) => {
      if (grossHdr && getKarnatakaRowValueForHeader(exportRow, grossHdr) !== '') {
        return exportRow;
      }
      const emp = employees[index]?.Employee || employees[index]?.employee || employees[index] || null;
      const payrollRow = resolvePayrollRow(emp);
      if (!payrollRow || payrollRow.fetch_error) return exportRow;
      return applyFormXIXKarnatakaFooterFieldsToRow(exportRow, payrollRow, hdrs, emp, monthCandidates);
    });
  }

  if (exportRows.length > 0 && employees.length > 0) {
    exportRows = exportRows.map((exportRow, index) => {
      const emp = employees[index]?.Employee || employees[index]?.employee || employees[index] || null;
      if (!emp) {
        syncFormXIXKarnatakaDeductionsFromGrossNet(exportRow, hdrs);
        return exportRow;
      }
      const payrollRow =
        typeof resolvePayrollRow === 'function' ? resolvePayrollRow(emp) : null;
      const finalized = finalizeFormXIXKarnatakaExportRowForEmployee(exportRow, emp, hdrs, {
        sanitizeValue: exportHelpers.sanitizeValue,
        payrollRow,
        monthCandidates,
      });
      syncFormXIXKarnatakaDeductionsFromGrossNet(finalized, hdrs);
      return finalized;
    });
  } else {
    exportRows.forEach((row) => syncFormXIXKarnatakaDeductionsFromGrossNet(row, hdrs));
  }

  const baseHeaderData = headerFormData && typeof headerFormData === 'object' ? { ...headerFormData } : {};
  const workbookArgs = {
    templateArrayBuffer,
    headersToUse: hdrs,
    parsedFormHeader,
    formFileName,
    headerFormData: baseHeaderData,
    resolvePayrollRow,
    monthCandidates,
  };

  if (exportRows.length === 0) {
    return buildFormXIXKarnatakaWorkbookWithTemplateStyles({
      ...workbookArgs,
      mappedData: [],
      employeesOverride: employees,
    });
  }

  if (exportRows.length > 1 && canUseFormXIXKarnatakaFastExport(parsedFormHeader)) {
    try {
      const fastTemplate = await prepareFormXIXKarnatakaFastZipTemplate({
        templateArrayBuffer,
        parsedFormHeader,
        headerFormData: baseHeaderData,
      });
      const hasPatchTargets =
        fastTemplate.positions.employeeHeaderPositions.length > 0 ||
        fastTemplate.positions.tablePositions.length > 0;
      if (hasPatchTargets) {
        return buildFormXIXKarnatakaFastZipDownload({
          exportRows,
          fastTemplate,
          baseHeaderData,
          parsedFormHeader,
          formFileName,
          employees,
          monthCandidates,
        });
      }
    } catch (fastZipErr) {
      console.warn('Form XIX Karnataka fast ZIP export failed, using standard export:', fastZipErr);
    }
  }

  const zip = new JSZip();
  const usedNames = new Map();
  for (let i = 0; i < exportRows.length; i += 1) {
    const emp = employees[i]?.Employee || employees[i]?.employee || employees[i] || null;
    if (shouldSkipFormXIXKarnatakaEmployeeExport(exportRows[i], emp, monthCandidates)) continue;
    const mergedHeaderData = buildEmployeeHeaderFormData(baseHeaderData, exportRows[i], emp);
    const { blob } = await buildFormXIXKarnatakaWorkbookWithTemplateStyles({
      ...workbookArgs,
      mappedData: [exportRows[i]],
      headerFormData: mergedHeaderData,
      employeesOverride: emp ? [emp] : [],
    });
    const bytes = new Uint8Array(await blob.arrayBuffer());
    const baseName = resolveEmployeeDownloadBaseName(exportRows[i], i);
    const count = usedNames.get(baseName) || 0;
    usedNames.set(baseName, count + 1);
    const suffix = count > 0 ? `_${count + 1}` : '';
    zip.file(`Form_XIX_Karnataka_${baseName}${suffix}.xlsx`, bytes);
    if (i > 0 && i % 15 === 0) {
      await new Promise((resolve) => setTimeout(resolve, 0));
    }
  }
  const zipBase = String(formFileName || parsedFormHeader?.title || 'Form_XIX_Karnataka')
    .replace(/\.xlsx?$/i, '')
    .replace(/[^a-zA-Z0-9._-]+/g, '_');
  return {
    blob: await zip.generateAsync({ type: 'blob', compression: 'STORE' }),
    fileName: `${zipBase}_Employees.zip`,
  };
}
