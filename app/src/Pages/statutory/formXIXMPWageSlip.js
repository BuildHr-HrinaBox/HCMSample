import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import {
  FORM_XIX_AP_FIELD_GROUPS,
  FORM_XIX_AP_TEMPLATE_SPECS,
  applyFormXIXAPAutofillFromSiteAndPayroll,
  formXIXAPHeaderNorm,
  formatWorkmanNameAndGuardian,
  isFormXIXAPTableWageSlipContext,
  isFormXIXAPWageSlipContext,
  matchesFormXIXHint,
  resolveFormXIXAPHeaderFieldLayout as resolveFormXIXAPHeaderFieldLayoutInner,
  writeFormXIXAPFieldsToExcelJsWorksheet,
} from './formXIXAPWageSlip';
import {
  flattenPayrollEarningColumns,
  getEarningsArray,
  readPayrollForm15WageAmounts,
  readPayrollScalar,
  filterPayrollRowsWithGrossPay,
} from '../../utils/payrollEarnings';
import {
  cacheForm15PayrollTableRows,
  getCachedForm15PayrollTableRows,
  getLatestCachedPayrollTableRows,
} from '../../utils/statutoryAutofillCache';
import { fetchPayrollTableRowsForMonths } from '../../utils/payrollTable';
import { personNamesMatch } from './formFKarnataka';
import { isAprilPayrollMonthCandidates } from './formQKarnataka';
import { FORM_XIX_MP_APR_PAYRUN_DEFAULTS } from './formXIXMPPayrunDefaults';

/** MP CLRA Form XIX — Wage Slip: site header fields + tabular wage particulars + per-employee ZIP. */

/** MP Form XIX — default No. of days worked when payroll lacks paid days. */
export const FORM_XIX_MP_DEFAULT_DAYS_WORKED = '30';

const FORM_XIX_MP_PAID_DAYS_KEYS = [
  'paid_days',
  'Paid Days',
  'days_worked',
  'Days Worked',
  'paidDays',
  'no_of_days_worked',
  'no_of_days_present',
  'present_days',
  'Effective Paid Days',
  'effective_paid_days',
  'effectivePaidDays',
  'Base Days',
  'base_days',
];

const FORM_XIX_MP_PAID_DAYS_PATTERNS = [
  /^paid_days$/,
  /paiddays/,
  /daysworked/,
  /days_present/,
  /noofdayspresent/,
  /present_days/,
  /effective_paid_days/,
  /base_days/,
];

const FORM_XIX_MP_GROSS_KEYS = [
  'gross_pay',
  'Gross Pay',
  'grossPay',
  'total_earnings',
  'Total Earnings',
];

const FORM_XIX_MP_DEDUCTIONS_KEYS = [
  'total_deductions',
  'Total Deductions',
  'totalDeductions',
  'total_employee_deductions',
  'total_deduction',
];

const FORM_XIX_MP_NET_KEYS = ['net_pay', 'Net Pay', 'netPay', 'net_wages'];

export function formatFormXIXMPEmployeeName(emp = {}) {
  const workmanLine = String(formatWorkmanNameAndGuardian(emp) || '')
    .split(/\r?\n/)[0]
    .trim();
  if (workmanLine) return workmanLine;
  const fn = String(emp.FirstName || emp['FirstName'] || emp.firstName || emp['First Name'] || '').trim();
  const ln = String(emp.LastName || emp['LastName'] || emp.lastName || emp['Last Name'] || '').trim();
  if (fn && ln) return `${fn} ${ln}`;
  return fn || ln || String(emp.Name || emp['Name'] || emp.EmployeeName || emp['Employee Name'] || '').trim();
}

function shouldUseFormXIXMPPayrunDefaults(options = {}) {
  if (!options.madhyaPradeshPayrollRules) return false;
  const months = options.monthCandidates;
  if (months == null || (Array.isArray(months) && months.length === 0)) return true;
  return isAprilPayrollMonthCandidates(months);
}

export function resolveFormXIXMPPayrunAprilDefaultWages(emp) {
  const empName = formatFormXIXMPEmployeeName(emp);
  if (!empName) return null;
  const match = FORM_XIX_MP_APR_PAYRUN_DEFAULTS.find((entry) => personNamesMatch(empName, entry.name));
  if (!match) return null;
  return {
    daysWorked: match.days != null ? String(match.days).trim() : FORM_XIX_MP_DEFAULT_DAYS_WORKED,
    grossWages: match.gross != null ? String(match.gross).trim() : '',
    deductions: match.deductions != null ? String(match.deductions).trim() : '',
    netWages: match.net != null ? String(match.net).trim() : '',
  };
}

function applyFormXIXMPPayrunDefaults(fields, emp, options = {}) {
  if (!shouldUseFormXIXMPPayrunDefaults(options) || !emp) return fields;
  const defaults = resolveFormXIXMPPayrunAprilDefaultWages(emp);
  if (!defaults) return fields;
  const out = { ...fields };
  if (!out.daysWorked) out.daysWorked = defaults.daysWorked || FORM_XIX_MP_DEFAULT_DAYS_WORKED;
  if (!out.grossWages) out.grossWages = defaults.grossWages;
  if (!out.deductions) out.deductions = defaults.deductions;
  if (!out.netWages) out.netWages = defaults.netWages;
  return out;
}

export const FORM_XIX_MP_WAGE_TABLE_HEADERS = [
  "Name and Father's/Husband's Name of the workman",
  '1. No. of days worked',
  '2. No. of units worked in case of piece-rate Workers',
  '3. Rate of daily wages/piece-rate',
  '4. Amount of overtime wages',
  '5. Gross wages payable',
  '6. Deductions, if any',
  '7. Net amount of wages paid',
];

/** Header/footer only — wage particulars render in the employee table below. */
export const FORM_XIX_MP_FIELD_GROUPS = [
  { id: 'header', title: 'Wage slip — workman & contractor' },
  { id: 'footer', title: 'Certification' },
];

const FORM_XIX_MP_HEADER_SPECS = FORM_XIX_AP_TEMPLATE_SPECS.filter((spec) => spec.group !== 'wages');

const FORM_XIX_MP_WAGE_SPECS = FORM_XIX_AP_TEMPLATE_SPECS.filter((spec) => spec.group === 'wages');

const FORM_XIX_MP_WAGE_KEY_BY_HEADER = new Map(
  FORM_XIX_MP_WAGE_SPECS.map((spec) => [spec.label.replace(/\s*:+\s*$/, ''), spec.key])
);

/** MP stacked template — values align in column E beside labels in column B. */
export const FORM_XIX_MP_STACKED_VALUE_COL = 5;

/** Tamil Nadu Form XIX — fixed text for rate column (not from payroll). */
export const FORM_XIX_TN_RATE_DEFAULT = 'Monthly Wages';

/** Gujarat Form XIX — same fixed rate label as Tamil Nadu / Karnataka. */
export const FORM_XIX_GJ_RATE_DEFAULT = FORM_XIX_TN_RATE_DEFAULT;
const FORM_XIX_GJ_FIXED_THARUN_GROSS = '70404';
const FORM_XIX_GJ_FIXED_THARUN_NET = '67937';
const FORM_XIX_GJ_FIXED_EMPLOYEE_MATCHES = ['tharun', 'tarun'];

/** Gujarat template — separate "Name of the workman" row (employee name only). */
const FORM_XIX_GJ_WORKMAN_NAME_SPEC = {
  key: 'form_xix_gj_workman_name',
  label: 'Name of the workman',
  group: 'header',
  fieldType: 'text',
  match: /^name\s+of\s+the\s+workman/i,
};

const formatWorkmanNameOnly = (emp = {}) =>
  String(formatWorkmanNameAndGuardian(emp) || '')
    .split(/\r?\n/)[0]
    .trim();

function isFormXIXGJFixedEmployee(emp = {}) {
  const name = String(formatFormXIXMPEmployeeName(emp) || '').trim().toLowerCase();
  if (!name) return false;
  return FORM_XIX_GJ_FIXED_EMPLOYEE_MATCHES.some((token) => name.includes(token));
}

function applyFormXIXGJFixedTharunWages(fields, options = {}) {
  if (!options?.gujaratPayrollRules) return fields;
  if (!isFormXIXGJFixedEmployee(options?.emp || {})) return fields;
  return {
    ...fields,
    grossWages: FORM_XIX_GJ_FIXED_THARUN_GROSS,
    netWages: FORM_XIX_GJ_FIXED_THARUN_NET,
  };
}

const sumPayrollScalars = (values) => {
  let sum = 0;
  let any = false;
  (Array.isArray(values) ? values : []).forEach((value) => {
    const n = Number(String(value ?? '').replace(/,/g, '').trim());
    if (Number.isFinite(n)) {
      sum += n;
      any = true;
    }
  });
  return any ? Math.round(sum * 100) / 100 : '';
};

const resolveFormXIXMPHeaderSpecs = (formHeader, item, fileName, sheetText = '') => {
  const specs = [...FORM_XIX_MP_HEADER_SPECS];
  if (isFormXIXGJGujaratWageSlipContext(formHeader, item, fileName, sheetText)) {
    specs.push(FORM_XIX_GJ_WORKMAN_NAME_SPEC);
  }
  return specs;
};

export function resolveFormXIXMPPayrollHelpers(contextHints = {}) {
  const formHeader = contextHints.formHeader || null;
  const item = contextHints.item || null;
  const fileName = contextHints.fileName || contextHints.formFileName || '';
  const sheetText = contextHints.sheetText || '';
  const gujarat = isFormXIXGJGujaratWageSlipContext(formHeader, item, fileName, sheetText);
  const tamilNadu = isFormXIXTamilNaduWageSlipContext(formHeader, item, fileName, sheetText);
  const andhraPradesh =
    !!formHeader?.formXIXAPTableLayout ||
    isFormXIXAPTableWageSlipContext(formHeader, item, fileName, sheetText);
  const madhyaPradesh =
    isFormXIXMPWageSlipContext(formHeader, item, fileName, sheetText) &&
    !gujarat &&
    !tamilNadu &&
    !andhraPradesh;
  return {
    gujaratPayrollRules: gujarat,
    useMonthlyWageRateDefault: gujarat || tamilNadu,
    madhyaPradeshPayrollRules: madhyaPradesh,
    andhraPradeshPayrollRules: andhraPradesh,
  };
}

/** Madhya Pradesh Form XIX only — not Karnataka / Gujarat / Tamil Nadu. */
export function isFormXIXMadhyaPradeshWageSlipContext(formHeader, rowItem, fileName, sheetText = '') {
  if (!isFormXIXAPWageSlipContext(formHeader, rowItem, fileName, sheetText)) return false;
  if (isFormXIXTamilNaduWageSlipContext(formHeader, rowItem, fileName, sheetText)) return false;
  if (isFormXIXGJGujaratWageSlipContext(formHeader, rowItem, fileName, sheetText)) return false;
  const parts = formXIXMPContextParts(formHeader, rowItem, fileName, sheetText);
  if (/karnataka|form[\s._-]*xix[\s._-]*ka|form_xix[\s._-]*karnataka/.test(parts)) return false;
  return (
    /madhya\s+pradesh/.test(parts) ||
    /form[\s._-]*xix[\s._-]*mp/.test(parts) ||
    (/\bmp\b/.test(parts) && matchesFormXIXHint(parts))
  );
}

const formXIXMPContextParts = (formHeader, rowItem, fileName, sheetText = '') =>
  [
    rowItem?.formName,
    rowItem?.FormName,
    rowItem?.description,
    rowItem?.Description,
    rowItem?.act,
    rowItem?.Act,
    rowItem?.siteState,
    rowItem?.SiteState,
    rowItem?.state,
    rowItem?.State,
    rowItem?.formFileName,
    rowItem?.FormFileName,
    fileName,
    formHeader?.title,
    formHeader?.subtitle,
    sheetText,
  ]
    .filter((x) => x != null && String(x).trim() !== '')
    .join(' ')
    .toLowerCase();

export function isFormXIXTamilNaduWageSlipContext(formHeader, rowItem, fileName, sheetText = '') {
  if (!isFormXIXAPWageSlipContext(formHeader, rowItem, fileName, sheetText)) return false;
  const parts = formXIXMPContextParts(formHeader, rowItem, fileName, sheetText);
  return /tamil[\s._-]*nadu|tamilnadu|form[\s._-]*xix[\s._-]*tamil|form_xix[_\s-]*tamil/.test(parts);
}

export function isFormXIXGJGujaratWageSlipContext(formHeader, rowItem, fileName, sheetText = '') {
  if (!isFormXIXAPWageSlipContext(formHeader, rowItem, fileName, sheetText)) return false;
  const parts = formXIXMPContextParts(formHeader, rowItem, fileName, sheetText);
  return /gujarat/.test(parts) || /form[\s._-]*xix[\s._-]*gj/.test(parts);
}

export function isFormXIXMPWageSlipContext(formHeader, rowItem, fileName, sheetText = '') {
  if (!isFormXIXAPWageSlipContext(formHeader, rowItem, fileName, sheetText)) return false;
  if (isFormXIXTamilNaduWageSlipContext(formHeader, rowItem, fileName, sheetText)) return true;
  if (isFormXIXGJGujaratWageSlipContext(formHeader, rowItem, fileName, sheetText)) return true;
  const parts = formXIXMPContextParts(formHeader, rowItem, fileName, sheetText);
  return (
    /madhya\s+pradesh/.test(parts) ||
    /form[\s._-]*xix[\s._-]*mp/.test(parts) ||
    (/\bmp\b/.test(parts) && matchesFormXIXHint(parts))
  );
}

export function isFormXIXMPHeaderFieldLayoutFormHeader(formHeader) {
  return !!formHeader?.formXIXMPTableLayout;
}

export function isFormXIXMPTableLayoutFormHeader(formHeader) {
  return !!formHeader?.formXIXMPTableLayout;
}

export function resolveFormXIXMPWageTableHeaders(tableHeaders) {
  const list = Array.isArray(tableHeaders) ? tableHeaders.filter(Boolean) : [];
  if (list.length >= FORM_XIX_MP_WAGE_TABLE_HEADERS.length - 1) return list;
  return [...FORM_XIX_MP_WAGE_TABLE_HEADERS];
}

const normHeader = (h) => formXIXAPHeaderNorm(String(h || '').replace(/^\d+[\.\)]\s*/, ''));

export function isFormXIXMPWorkmanNameHeader(h) {
  const s = normHeader(h);
  return (s.includes('workman') || s.includes('workmen')) && (s.includes('name') || s.includes('father') || s.includes('husband'));
}

export function isFormXIXMPDaysWorkedHeader(h) {
  return /(?:no|number)\.?\s*of\s+days\s+worked|days\s+worked/.test(normHeader(h));
}

export function isFormXIXMPUnitsWorkedHeader(h) {
  return /units\s+worked/.test(normHeader(h));
}

export function isFormXIXMPRateHeader(h) {
  const s = normHeader(h);
  return s.includes('rate') && (s.includes('wage') || s.includes('piece'));
}

export function isFormXIXMPOvertimeHeader(h) {
  return /overtime/.test(normHeader(h));
}

export function isFormXIXMPGrossHeader(h) {
  return /gross\s+wages/.test(normHeader(h));
}

export function isFormXIXMPDeductionsHeader(h) {
  return /deduction/.test(normHeader(h));
}

export function isFormXIXMPNetHeader(h) {
  return /net\s+amount/.test(normHeader(h));
}

/** Wage/payroll columns must come from Payroll table — workman name from People FirstName/LastName/Father_s_Name. */
export function isFormXIXMPSkipPeopleAutofillHeader(h) {
  return (
    isFormXIXMPWorkmanNameHeader(h) ||
    isFormXIXMPDaysWorkedHeader(h) ||
    isFormXIXMPUnitsWorkedHeader(h) ||
    isFormXIXMPRateHeader(h) ||
    isFormXIXMPOvertimeHeader(h) ||
    isFormXIXMPGrossHeader(h) ||
    isFormXIXMPDeductionsHeader(h) ||
    isFormXIXMPNetHeader(h)
  );
}

const normPayrollEmployeeCode = (value) => String(value || '').trim().replace(/^0+/, '') || '0';

const normPersonName = (value) =>
  String(value || '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ');

export function normPayrollGid(value) {
  return String(value || '').trim().toUpperCase();
}

const PAYROLL_GID_FIELD_KEYS = [
  'gidNumber',
  'GIDNumber',
  'gid_number',
  'GID Number',
  'GID_Number',
  'employee_number',
  'Employee Number',
  'Employee_Number',
  'employee_code',
  'Employee Code',
];

export function payrollRowGidCandidates(row) {
  if (!row || typeof row !== 'object') return [];
  const gids = new Set();
  PAYROLL_GID_FIELD_KEYS.forEach((key) => {
    const value = row[key];
    const normalized = normPayrollGid(value);
    if (normalized) gids.add(normalized);
  });
  return Array.from(gids);
}

export function employeeGidCandidates(emp, row = null) {
  const codes = [
    emp?.gidNumber,
    emp?.GIDNumber,
    emp?.gid_number,
    emp?.['GID Number'],
    emp?.employee_number,
    emp?.['employee_number'],
    emp?.Employee_Number,
    emp?.['Employee Number'],
    emp?.EmployeeCode,
    emp?.['Employee Code'],
    emp?.employeeCode,
    emp?.EmpCode,
    emp?.WorkerIdentityNo,
    emp?.WorkerIdentityNumber,
    emp?.['Worker Identity No'],
    emp?.['Worker Identity Number'],
  ];
  if (row && typeof row === 'object') {
    Object.entries(row).forEach(([key, value]) => {
      if (value == null || String(value).trim() === '') return;
      if (isFormPayrollGidHeader(key)) codes.push(value);
    });
  }
  const uniq = new Set();
  codes.forEach((value) => {
    const normalized = normPayrollGid(value);
    if (normalized) uniq.add(normalized);
  });
  return Array.from(uniq);
}

export function isFormPayrollGidHeader(header) {
  const s = String(header || '')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9 ]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  if (!s) return false;
  if (s.includes('gid')) return true;
  if (s === 'emp id' || s === 'empid') return true;
  if (
    (s.includes('worker') || s.includes('employee') || s.includes('emp')) &&
    (s.includes('identity') || s.includes('identification'))
  ) {
    return true;
  }
  if (s.includes('employee') && s.includes('code')) return true;
  if (s.includes('worker') && s.includes('code')) return true;
  return false;
}

function payrollRowMatchesNameCandidates(nameCandidates, payrollRow) {
  const payrollName = payrollRowDisplayName(payrollRow);
  if (!payrollName) return false;
  const names = (Array.isArray(nameCandidates) ? nameCandidates : [nameCandidates])
    .map((name) => normPersonName(name))
    .filter(Boolean);
  return names.some(
    (candidate) =>
      payrollName === candidate ||
      personNamesMatch(candidate, payrollName) ||
      personNamesMatch(payrollName, candidate)
  );
}

function payrollRowMatchesGidCandidates(gidCandidates, payrollRow) {
  const gids = (Array.isArray(gidCandidates) ? gidCandidates : [gidCandidates])
    .map(normPayrollGid)
    .filter(Boolean);
  if (gids.length === 0) return true;
  const payrollGids = payrollRowGidCandidates(payrollRow);
  if (payrollGids.length === 0) return false;
  const gidSet = new Set([...gids, ...gids.map(normPayrollEmployeeCode)]);
  return payrollGids.some(
    (payrollGid) => gidSet.has(payrollGid) || gidSet.has(normPayrollEmployeeCode(payrollGid))
  );
}

export function collectEmployeeNameCandidates(emp) {
  if (!emp || typeof emp !== 'object') return [];
  const fn = String(emp.FirstName || emp['FirstName'] || '').trim();
  const ln = String(emp.LastName || emp['LastName'] || '').trim();
  const names = [];
  if (fn || ln) names.push(normPersonName(`${fn} ${ln}`.trim()));
  if (fn && ln) names.push(normPersonName(`${ln} ${fn}`.trim()));
  const displayName = normPersonName(
    emp.DisplayName ||
      emp['Display Name'] ||
      emp.displayName ||
      emp.EmployeeName ||
      emp['Employee Name'] ||
      emp.Employee_Name ||
      ''
  );
  if (displayName) names.push(displayName);
  const workmanLine = String(formatWorkmanNameAndGuardian(emp) || '')
    .split(/\r?\n/)[0]
    .trim();
  if (workmanLine) names.push(normPersonName(workmanLine));
  return [...new Set(names.filter(Boolean))];
}

/** Match Sample Payroll row only when Employee Name and GID Number both align. */
export function resolvePayrollRowByNameAndGid(nameCandidates, gidCandidates, payrollRows) {
  const names = (Array.isArray(nameCandidates) ? nameCandidates : [nameCandidates])
    .map((name) => normPersonName(name))
    .filter(Boolean);
  if (names.length === 0) return null;
  const gids = (Array.isArray(gidCandidates) ? gidCandidates : [gidCandidates])
    .map((gid) => String(gid || '').trim())
    .filter(Boolean);
  const rows = (Array.isArray(payrollRows) ? payrollRows : []).filter((row) => row && !row.fetch_error);
  const matches = rows.filter(
    (row) => payrollRowMatchesNameCandidates(names, row) && payrollRowMatchesGidCandidates(gids, row)
  );
  return matches.length > 0 ? matches[0] : null;
}

const payrollEmployeeCodes = (row) =>
  [
    row?.employee_id,
    row?.employee_number,
    row?.employee_code,
    row?.employee_no,
    row?.emp_id,
    row?.EmployeeID,
    row?.['Employee ID'],
    row?.['Employee No'],
    row?.['Employee Number'],
    row?.gidNumber,
    row?.GIDNumber,
    row?.gid_number,
    row?.['GID Number'],
  ]
    .map((value) => String(value || '').trim())
    .filter(Boolean);

const payrollRowDisplayName = (row) =>
  normPersonName(
    row?.employee_name ||
      row?.full_name ||
      row?.name ||
      row?.['Employee Name'] ||
      row?.employeeName ||
      ''
  );

const normalizeFormXIXMPDaysWorked = (value) => {
  const raw = String(value ?? '').trim();
  if (!raw) return '';
  const n = Number(raw.replace(/,/g, ''));
  if (Number.isFinite(n) && n > 0) return String(Math.round(n));
  return raw;
};

const payrollRowIdentityKey = (row) => {
  if (!row || typeof row !== 'object') return '';
  const gids = payrollRowGidCandidates(row);
  const name = payrollRowDisplayName(row);
  if (gids.length > 0 && name) return `gid:${gids[0]}|name:${name}`;
  if (gids.length > 0) return `gid:${gids[0]}`;
  const codes = payrollEmployeeCodes(row);
  if (codes.length > 0) return `id:${normPayrollEmployeeCode(codes[0])}`;
  return name ? `name:${name}` : '';
};

const employeePeopleCodes = (emp) => {
  const codes = [
    emp?.Zoho_ID,
    emp?.['Zoho_ID'],
    emp?.ZohoID,
    emp?.EmployeeID,
    emp?.['EmployeeID'],
    emp?.['Employee ID'],
    emp?.employee_number,
    emp?.['employee_number'],
    emp?.Employee_Number,
    emp?.['Employee Number'],
    emp?.EmployeeCode,
    emp?.['Employee Code'],
    emp?.gidNumber,
    emp?.GIDNumber,
    emp?.gid_number,
    emp?.['GID Number'],
    emp?.WorkerIdentityNo,
    emp?.WorkerIdentityNumber,
    emp?.['Worker Identity No'],
    emp?.['Worker Identity Number'],
  ];
  if (emp?.Role && typeof emp.Role === 'object' && !Array.isArray(emp.Role)) {
    codes.push(emp.Role.ID, emp.Role.Id, emp.Role.id);
  }
  return codes.map((value) => String(value || '').trim()).filter(Boolean);
};

/** Match Payroll table row to Zoho People employee (employee_id / employee_number / name). */
export function resolveFormXIXMPPayrollRowForEmployee(emp, payrollRows) {
  if (!emp || !Array.isArray(payrollRows) || payrollRows.length === 0) return null;

  const nameCandidates = collectEmployeeNameCandidates(emp);
  const gidCandidates = employeeGidCandidates(emp);
  if (nameCandidates.length > 0) {
    const compositeHit = resolvePayrollRowByNameAndGid(nameCandidates, gidCandidates, payrollRows);
    if (compositeHit) return compositeHit;
    if (gidCandidates.length > 0) return null;
  }

  const peopleCodes = employeePeopleCodes(emp);
  const normPeopleCodes = new Set(peopleCodes.map(normPayrollEmployeeCode));
  let hit =
    payrollRows.find((row) => {
      if (!row || row.fetch_error) return false;
      return payrollEmployeeCodes(row).some((code) => {
        const raw = String(code || '').trim();
        if (!raw) return false;
        return peopleCodes.includes(raw) || normPeopleCodes.has(normPayrollEmployeeCode(raw));
      });
    }) || null;
  if (hit) return hit;

  const email = String(emp.EmailID || emp.Email || emp.email || emp['Email ID'] || '')
    .trim()
    .toLowerCase();
  if (email) {
    hit =
      payrollRows.find(
        (row) =>
          !row?.fetch_error &&
          String(row.work_mail || row.email || row.work_email || '')
            .trim()
            .toLowerCase() === email
      ) || null;
    if (hit) return hit;
  }

  const fn = String(emp.FirstName || emp['FirstName'] || '').trim().toLowerCase();
  const ln = String(emp.LastName || emp['LastName'] || '').trim().toLowerCase();
  const combo = normPersonName(`${fn} ${ln}`);
  const displayName = normPersonName(
    emp.DisplayName ||
      emp['Display Name'] ||
      emp.displayName ||
      emp.EmployeeName ||
      emp['Employee Name'] ||
      ''
  );
  const legacyNameCandidates = [combo, displayName].filter(Boolean);
  if (fn && ln) {
    legacyNameCandidates.push(normPersonName(`${ln} ${fn}`));
  }
  if (legacyNameCandidates.length > 0) {
    hit =
      payrollRows.find((row) => {
        if (!row || row.fetch_error) return false;
        const payrollName = payrollRowDisplayName(row);
        if (!payrollName) return false;
        const rfn = String(row.first_name || row.firstName || row['First Name'] || '').trim().toLowerCase();
        const rln = String(row.last_name || row.lastName || row['Last Name'] || '').trim().toLowerCase();
        const rcombo = normPersonName(`${rfn} ${rln}`);
        return legacyNameCandidates.some(
          (candidate) =>
            payrollName === candidate ||
            rcombo === candidate ||
            personNamesMatch(candidate, payrollName)
        );
      }) || null;
    if (hit) return hit;
  }

  const workmanLine = String(formatWorkmanNameAndGuardian(emp) || '')
    .split(/\r?\n/)[0]
    .trim()
    .toLowerCase();
  if (workmanLine) {
    hit =
      payrollRows.find((row) => {
        if (!row || row.fetch_error) return false;
        const payrollName = payrollRowDisplayName(row);
        return (
          payrollName === normPersonName(workmanLine) ||
          personNamesMatch(workmanLine, payrollName) ||
          payrollName.includes(normPersonName(workmanLine)) ||
          normPersonName(workmanLine).includes(payrollName)
        );
      }) || null;
  }
  return hit;
}

/** Match SamplePayroll / pay-run row by display name (form row or People). */
export function resolvePayrollRowByDisplayName(displayName, payrollRows) {
  const name = normPersonName(displayName);
  if (!name || !Array.isArray(payrollRows) || payrollRows.length === 0) return null;
  const rows = payrollRows.filter((row) => row && !row.fetch_error);

  const findUnique = (target) => {
    if (!target) return null;
    const exact = rows.filter((row) => payrollRowDisplayName(row) === target);
    if (exact.length === 1) return exact[0];
    if (exact.length > 1) return exact[0];
    const fuzzy = rows.filter((row) => {
      const payrollName = payrollRowDisplayName(row);
      return payrollName && personNamesMatch(target, payrollName);
    });
    return fuzzy.length === 1 ? fuzzy[0] : null;
  };

  let hit = findUnique(name);
  if (hit) return hit;

  const parts = name.split(' ').filter(Boolean);
  if (parts.length >= 2) {
    hit = findUnique(normPersonName(parts.slice().reverse().join(' ')));
    if (hit) return hit;
  }
  return null;
}

export function resolvePayrollRowByFormTableName(formRow, headers, payrollRows, { isNameHeader } = {}) {
  if (!formRow || !Array.isArray(headers) || !Array.isArray(payrollRows)) return null;
  const nameValues = [];
  const gidValues = [];
  headers.forEach((header) => {
    const raw = String(getFormXIXMPRowValueForHeader(formRow, header) || '').trim();
    if (!raw || /^enter\b/i.test(raw)) return;
    if (typeof isNameHeader === 'function' && isNameHeader(header)) nameValues.push(raw);
    if (isFormPayrollGidHeader(header)) gidValues.push(raw);
  });
  const lookupName = String(formRow.__employeeLookupName || '').trim();
  if (lookupName) nameValues.push(lookupName);

  const compositeHit = resolvePayrollRowByNameAndGid(
    [...new Set(nameValues)],
    [...new Set(gidValues)],
    payrollRows
  );
  if (compositeHit) return compositeHit;
  if (gidValues.length > 0 && nameValues.length > 0) return null;

  const hdrs = headers.filter((header) =>
    typeof isNameHeader === 'function' ? isNameHeader(header) : false
  );
  for (let i = 0; i < hdrs.length; i += 1) {
    const raw = String(getFormXIXMPRowValueForHeader(formRow, hdrs[i]) || '').trim();
    if (!raw || /^enter\b/i.test(raw)) continue;
    const hit = resolvePayrollRowByDisplayName(raw, payrollRows);
    if (hit) return hit;
  }
  if (lookupName) {
    return resolvePayrollRowByDisplayName(lookupName, payrollRows);
  }
  return null;
}

export function getFormXIXMPRowValueForHeader(row, header) {
  if (!row || typeof row !== 'object') return '';
  if (Object.prototype.hasOwnProperty.call(row, header)) return row[header];
  const target = normHeader(header);
  const key = Object.keys(row).find((k) => {
    const n = normHeader(k);
    return n === target || n.includes(target) || target.includes(n);
  });
  return key ? row[key] : '';
}

export function rowHasMeaningfulFormXIXMPExportData(row, headers) {
  if (!row || typeof row !== 'object') return false;
  const hdrs = resolveFormXIXMPWageTableHeaders(headers);
  return hdrs.some((header) => String(getFormXIXMPRowValueForHeader(row, header) ?? '').trim() !== '');
}

const pickPayrollScalar = (payrollRow, keys, patterns = []) => {
  if (!payrollRow || payrollRow.fetch_error) return '';
  return readPayrollScalar(payrollRow, keys, patterns);
};

const pickPayrollField = (flatRow, payrollRow, keys, patterns = []) => {
  const primary = pickPayrollScalar(flatRow, keys, patterns);
  if (primary !== '') return primary;
  if (payrollRow && payrollRow !== flatRow) {
    return pickPayrollScalar(payrollRow, keys, patterns);
  }
  return '';
};

export function payrollRowHasFormXIXMPPayRunFields(payrollRow) {
  if (!payrollRow || payrollRow.fetch_error) return false;
  if (payrollRowLooksLikeBulkSalaryNotPayRun(payrollRow)) return false;
  const flat = flattenPayrollEarningColumns(payrollRow);
  const days = pickPayrollField(flat, payrollRow, FORM_XIX_MP_PAID_DAYS_KEYS, FORM_XIX_MP_PAID_DAYS_PATTERNS);
  const gross = pickPayrollField(flat, payrollRow, FORM_XIX_MP_GROSS_KEYS, [/^gross_pay$/, /^total_earnings$/]);
  const net = pickPayrollField(flat, payrollRow, FORM_XIX_MP_NET_KEYS, [/^net_pay$/]);
  const deductions = pickPayrollField(
    flat,
    payrollRow,
    FORM_XIX_MP_DEDUCTIONS_KEYS,
    [/^total_deductions?$/, /^total_employee_deductions$/]
  );
  const paymentStatus = String(
    pickPayrollField(
      flat,
      payrollRow,
      ['payment_status', 'Payment Status', 'paymentStatus'],
      [/payment_status/]
    ) || ''
  )
    .trim()
    .toLowerCase();
  if (days !== '' && (gross !== '' || net !== '')) return true;
  if (gross !== '' && net !== '' && deductions !== '') return true;
  return gross !== '' && net !== '' && (paymentStatus === 'paid' || paymentStatus === 'partially_paid');
}

/** Salary-template rows from list_employees — not processed pay-run amounts. */
export function payrollRowLooksLikeBulkSalaryNotPayRun(payrollRow) {
  if (!payrollRow || payrollRow.fetch_error) return false;
  const flat = flattenPayrollEarningColumns(payrollRow);
  const days = pickPayrollField(
    flat,
    payrollRow,
    FORM_XIX_MP_PAID_DAYS_KEYS,
    FORM_XIX_MP_PAID_DAYS_PATTERNS
  );
  if (days !== '') return false;
  const monthlySalary = readPayrollScalar(flat, ['monthly_salary', 'MonthlySalary', 'monthly_gross_amount']);
  if (monthlySalary !== '') return true;
  const paymentStatus = String(
    pickPayrollField(
      flat,
      payrollRow,
      ['payment_status', 'Payment Status', 'paymentStatus'],
      [/payment_status/]
    ) || ''
  )
    .trim()
    .toLowerCase();
  if (paymentStatus === 'paid' || paymentStatus === 'partially_paid') return false;
  const gross = pickPayrollField(
    flat,
    payrollRow,
    FORM_XIX_MP_GROSS_KEYS,
    [/^gross_pay$/, /^total_earnings$/]
  );
  const net = pickPayrollField(flat, payrollRow, FORM_XIX_MP_NET_KEYS, [/^net_pay$/]);
  return gross !== '' || net !== '';
}

const flattenFormXIXMPPayrollRows = (rows) =>
  (Array.isArray(rows) ? rows : [])
    .filter((row) => row && typeof row === 'object' && row.fetch_error !== true)
    .map((row) => flattenPayrollEarningColumns(row));

export function filterFormXIXMPEligiblePayrollRows(rows) {
  const flat = flattenFormXIXMPPayrollRows(rows);
  if (flat.length === 0) return [];
  const withoutBulk = flat.filter((row) => !payrollRowLooksLikeBulkSalaryNotPayRun(row));
  if (payrollRowsLookLikeFormXIXMPPayRunTable(withoutBulk)) return withoutBulk;
  if (payrollRowsLookLikeFormXIXMPPayRunTable(flat)) return flat;
  return [];
}

/** Pay-run / Payroll-table rows only — never all_salaries bulk. */
export function resolveFormXIXMPPayrollRowsForAutofill(statutoryPayrollRows, monthCandidates = []) {
  const fromStatutory = filterFormXIXMPEligiblePayrollRows(statutoryPayrollRows);
  if (fromStatutory.length > 0) return fromStatutory;

  const fromGross = filterPayrollRowsWithGrossPay(statutoryPayrollRows);
  if (fromGross.length > 0) return fromGross;

  const cached = getCachedForm15PayrollTableRows(monthCandidates);
  if (cached?.rows?.length > 0) {
    const fromCache = filterFormXIXMPEligiblePayrollRows(cached.rows);
    if (fromCache.length > 0) return fromCache;
    const fromCacheGross = filterPayrollRowsWithGrossPay(cached.rows);
    if (fromCacheGross.length > 0) return fromCacheGross;
  }

  const latestCached = getLatestCachedPayrollTableRows();
  if (latestCached?.rows?.length > 0) {
    const fromLatest = filterFormXIXMPEligiblePayrollRows(latestCached.rows);
    if (fromLatest.length > 0) return fromLatest;
    const fromLatestGross = filterPayrollRowsWithGrossPay(latestCached.rows);
    if (fromLatestGross.length > 0) return fromLatestGross;
  }
  return [];
}

/** Load pay-run rows from Catalyst Payroll table — real-time API, then stored table cache. */
export async function loadFormXIXMPPayrollRowsForAutofill(monthCandidates = [], options = {}) {
  const months = Array.isArray(monthCandidates) ? monthCandidates : [];
  const timeoutMs = Number(options.timeoutMs) > 0 ? Number(options.timeoutMs) : 45000;
  const tableLoad = await fetchPayrollTableRowsForMonths(months, {
    timeoutMs,
    force: Boolean(options.force),
  });
  if (Array.isArray(tableLoad.rows) && tableLoad.rows.length > 0) {
    const payrollMonth = tableLoad.payrollMonth || months[0] || '';
    cacheForm15PayrollTableRows(payrollMonth, tableLoad.rows, tableLoad.meta || null);
    const eligible = filterFormXIXMPEligiblePayrollRows(tableLoad.rows);
    if (eligible.length > 0) return eligible;
    const withGross = filterPayrollRowsWithGrossPay(tableLoad.rows);
    if (withGross.length > 0) return withGross;
  }
  return resolveFormXIXMPPayrollRowsForAutofill(tableLoad?.rows, months);
}

/** One payroll row per employee — avoids reusing the same pay-run row for many People records. */
export function buildFormXIXMPPayrollRowResolver(payrollRows) {
  const rows = Array.isArray(payrollRows) ? payrollRows : [];
  const usedKeys = new Set();
  return (emp) => {
    const available = rows.filter((row) => {
      const key = payrollRowIdentityKey(row);
      return !key || !usedKeys.has(key);
    });
    const hit = resolveFormXIXMPPayrollRowForEmployee(emp, available.length > 0 ? available : rows);
    const hitKey = payrollRowIdentityKey(hit);
    if (hit && hitKey) usedKeys.add(hitKey);
    return hit;
  };
}

export function payrollRowsLookLikeFormXIXMPPayRunTable(rows) {
  if (!Array.isArray(rows) || rows.length === 0) return false;
  return rows.some((row) => payrollRowHasFormXIXMPPayRunFields(row));
}

const payrollLineAmount = (item) => {
  if (!item || typeof item !== 'object') return NaN;
  const raw = item.amount ?? item.value ?? item.employee_contribution ?? item.employer_contribution ?? '';
  const n = Number(String(raw).replace(/,/g, '').trim());
  return Number.isFinite(n) ? n : NaN;
};

const sumPayrollDeductionLines = (payrollRow) => {
  const items = payrollRow?.deductions ?? payrollRow?.Deductions;
  if (!Array.isArray(items) || items.length === 0) return '';
  let sum = 0;
  let any = false;
  items.forEach((item) => {
    const n = payrollLineAmount(item);
    if (Number.isFinite(n)) {
      sum += n;
      any = true;
    }
  });
  return any ? Math.round(sum * 100) / 100 : '';
};

/** Basic for the pay period from earnings lines (not gross / paid_days). */
const pickPeriodBasicFromEarnings = (payrollRow) => {
  const earnings = getEarningsArray(payrollRow);
  for (let i = 0; i < earnings.length; i += 1) {
    const item = earnings[i];
    const name = String(item?.name ?? item?.salary_component_name ?? item?.component_name ?? '')
      .trim()
      .toLowerCase();
    const type = String(item?.type ?? item?.earning_type ?? item?.salary_component_type ?? '')
      .trim()
      .toLowerCase();
    const isBasic =
      type === 'basic' ||
      type === 'earned_basic' ||
      name === 'basic' ||
      name === 'earned basic' ||
      name === 'basic pay' ||
      name === 'basic earnings' ||
      name === 'basic wage';
    if (!isBasic) continue;
    const n = payrollLineAmount(item);
    if (Number.isFinite(n) && n > 0) return n;
  }
  return '';
};

/** Daily basic wage rate — must not be gross_pay ÷ paid_days. */
const resolveFormXIXMPDailyWageRate = (flatRow, payrollRow) => {
  const explicit = pickPayrollField(
    flatRow,
    payrollRow,
    ['daily_wage_rate', 'daily_rate', 'rate_of_daily_wages', 'normal_rate_of_pay', 'Normal Rate of Pay'],
    [/daily_wage_rate/, /daily_rate/, /rate_of_daily_wages/, /normal_rate_of_pay/]
  );
  if (explicit !== '') return explicit;

  const days = pickPayrollField(
    flatRow,
    payrollRow,
    ['paid_days', 'Paid Days', 'days_worked', 'Days Worked', 'paidDays', 'no_of_days_worked'],
    [/^paid_days$/, /paiddays/, /daysworked/, /days_present/, /noofdayspresent/, /no_of_days_present/]
  );
  if (days === '' || Number(days) <= 0) return '';
  const daysNum = Number(days);

  const gross = pickPayrollField(
    flatRow,
    payrollRow,
    ['gross_pay', 'Gross Pay', 'grossPay'],
    [/^gross_pay$/, /^total_earnings$/]
  );
  const grossNum = gross !== '' ? Number(gross) : NaN;
  const avgGrossPerDay =
    Number.isFinite(grossNum) && grossNum > 0 ? Math.round((grossNum / daysNum) * 100) / 100 : NaN;

  let periodBasic = pickPeriodBasicFromEarnings(payrollRow);
  if (periodBasic === '') {
    const { basic } = readPayrollForm15WageAmounts(payrollRow);
    if (basic !== '') periodBasic = Number(basic);
  }
  if (periodBasic === '') {
    periodBasic = pickPayrollField(
      flatRow,
      payrollRow,
      ['basic', 'Basic', 'basic_pay', 'Basic_Pay', 'basic_earnings', 'Basic Earnings', 'earned_basic', 'Earned Basic'],
      [/^basic$/, /^earned_basic$/, /^basic_pay$/]
    );
  }

  if (periodBasic !== '') {
    const basicNum = Number(periodBasic);
    const dailyBasic = Math.round((basicNum / daysNum) * 100) / 100;

    // Reject mis-tagged values that are really gross ÷ days (not contractual daily rate).
    if (Number.isFinite(avgGrossPerDay) && Math.abs(dailyBasic - avgGrossPerDay) < 0.05) {
      if (basicNum >= grossNum * 0.98) {
        /* fall through to gross/days below */
      } else {
        return dailyBasic;
      }
    } else {
      return dailyBasic;
    }
  }

  if (Number.isFinite(grossNum) && grossNum > 0 && daysNum > 0) {
    return Math.round((grossNum / daysNum) * 100) / 100;
  }

  return '';
};

export function resolveFormXIXMPPayrollFields(payrollRow, options = {}) {
  if (!payrollRow || payrollRow.fetch_error) {
    const base = applyFormXIXMPPayrunDefaults(
      {
        daysWorked: options.madhyaPradeshPayrollRules ? FORM_XIX_MP_DEFAULT_DAYS_WORKED : '',
        unitsWorked: '',
        rate: options.useMonthlyWageRateDefault ? FORM_XIX_TN_RATE_DEFAULT : '',
        overtimeWages: options.andhraPradeshPayrollRules ? 'NIL' : '',
        grossWages: '',
        deductions: '',
        netWages: '',
      },
      options.emp,
      options
    );
    return applyFormXIXGJFixedTharunWages(base, options);
  }
  const row = flattenPayrollEarningColumns(payrollRow);
  const grossWages = pickPayrollField(
    row,
    payrollRow,
    FORM_XIX_MP_GROSS_KEYS,
    [/^gross_pay$/, /^total_earnings$/]
  );
  const netWages = pickPayrollField(row, payrollRow, FORM_XIX_MP_NET_KEYS, [/^net_pay$/]);

  let deductions = pickPayrollField(
    row,
    payrollRow,
    FORM_XIX_MP_DEDUCTIONS_KEYS,
    [/^total_deductions?$/, /^total_employee_deductions$/]
  );
  if (options.andhraPradeshPayrollRules) {
    // AP Form XIX — deductions = gross_pay − net_pay
    if (grossWages !== '' && netWages !== '') {
      const diff = Number(String(grossWages).replace(/,/g, '')) - Number(String(netWages).replace(/,/g, ''));
      if (Number.isFinite(diff) && diff >= 0) {
        deductions = Math.round(diff * 100) / 100;
      }
    }
  } else {
    if (options.gujaratPayrollRules) {
      const totalBenefits = pickPayrollField(
        row,
        payrollRow,
        ['total_benefits', 'Total Benefits', 'totalBenefits'],
        [/^total_benefits$/]
      );
      const totalTaxes = pickPayrollField(
        row,
        payrollRow,
        ['total_taxes', 'Total Taxes', 'totalTaxes'],
        [/^total_taxes$/]
      );
      const summed = sumPayrollScalars([deductions, totalBenefits, totalTaxes]);
      if (summed !== '') deductions = summed;
    }
    if (deductions === '') deductions = sumPayrollDeductionLines(payrollRow);
    if (deductions === '') {
      if (grossWages !== '' && netWages !== '') {
        const diff = Number(grossWages) - Number(netWages);
        if (Number.isFinite(diff) && diff >= 0) deductions = Math.round(diff * 100) / 100;
      }
    }
  }

  let daysWorked = normalizeFormXIXMPDaysWorked(
    pickPayrollField(row, payrollRow, FORM_XIX_MP_PAID_DAYS_KEYS, FORM_XIX_MP_PAID_DAYS_PATTERNS)
  );
  if (!daysWorked && options.madhyaPradeshPayrollRules) {
    daysWorked = FORM_XIX_MP_DEFAULT_DAYS_WORKED;
  }

  const resolved = applyFormXIXMPPayrunDefaults(
    {
      daysWorked,
      unitsWorked: pickPayrollField(row, payrollRow, ['units_worked', 'Units Worked', 'piece_units'], [/units_worked/]),
      rate: options.useMonthlyWageRateDefault
        ? FORM_XIX_TN_RATE_DEFAULT
        : resolveFormXIXMPDailyWageRate(row, payrollRow),
      // AP Form XIX — overtime wages always NIL
      overtimeWages: options.andhraPradeshPayrollRules
        ? 'NIL'
        : pickPayrollField(
            row,
            payrollRow,
            ['overtime', 'Overtime', 'overtime_wages', 'overtimeWages', 'ot'],
            [/overtime/]
          ),
      grossWages,
      deductions,
      netWages,
    },
    options.emp,
    options
  );
  return applyFormXIXGJFixedTharunWages(resolved, options);
}

export function enrichFormXIXTamilNaduStaticFieldRows(mappedData, headers) {
  if (!Array.isArray(mappedData) || mappedData.length === 0) return 0;
  const hdrs = resolveFormXIXMPWageTableHeaders(headers);
  const rateHdrs = hdrs.filter(isFormXIXMPRateHeader);
  if (rateHdrs.length === 0) return 0;
  let hits = 0;
  mappedData.forEach((row) => {
    if (!row || typeof row !== 'object') return;
    const rateKeys = new Set(rateHdrs);
    Object.keys(row).forEach((key) => {
      if (String(key).startsWith('__')) return;
      if (isFormXIXMPRateHeader(key)) rateKeys.add(key);
    });
    rateKeys.forEach((header) => {
      row[header] = FORM_XIX_TN_RATE_DEFAULT;
    });
    if (rateKeys.size > 0) hits += 1;
  });
  return hits;
}

export function applyFormXIXMPEmployeeToRow(row, emp, headers, helpers = {}) {
  const hdrs = resolveFormXIXMPWageTableHeaders(headers);
  const out = row && typeof row === 'object' ? { ...row } : {};
  const {
    sanitizeValue = (v) => String(v ?? '').trim(),
    payrollRow = null,
    useMonthlyWageRateDefault = false,
    gujaratPayrollRules = false,
    madhyaPradeshPayrollRules = false,
    andhraPradeshPayrollRules = false,
    monthCandidates = null,
  } = helpers;
  const hasPayroll = payrollRow && !payrollRow.fetch_error;
  const payroll = resolveFormXIXMPPayrollFields(hasPayroll ? payrollRow : null, {
    emp,
    monthCandidates,
    gujaratPayrollRules,
    useMonthlyWageRateDefault,
    madhyaPradeshPayrollRules,
    andhraPradeshPayrollRules,
  });
  const wageValue = (value, current = '') => {
    if (hasPayroll) return value !== '' ? sanitizeValue(value) : '';
    if (madhyaPradeshPayrollRules && value !== '') return sanitizeValue(value);
    if (andhraPradeshPayrollRules && value !== '') return sanitizeValue(value);
    return value !== '' ? sanitizeValue(value) : String(current ?? '').trim();
  };

  hdrs.forEach((header) => {
    if (isFormXIXMPWorkmanNameHeader(header)) {
      out[header] = sanitizeValue(formatWorkmanNameAndGuardian(emp));
      return;
    }
    if (isFormXIXMPDaysWorkedHeader(header)) {
      const days =
        wageValue(payroll.daysWorked, out[header]) ||
        (madhyaPradeshPayrollRules ? FORM_XIX_MP_DEFAULT_DAYS_WORKED : '');
      out[header] = sanitizeValue(days);
      return;
    }
    if (isFormXIXMPUnitsWorkedHeader(header)) {
      out[header] = wageValue(payroll.unitsWorked, out[header]);
      return;
    }
    if (isFormXIXMPRateHeader(header)) {
      if (useMonthlyWageRateDefault) {
        out[header] = sanitizeValue(FORM_XIX_TN_RATE_DEFAULT);
        return;
      }
      out[header] = wageValue(payroll.rate, out[header]);
      return;
    }
    if (isFormXIXMPOvertimeHeader(header)) {
      if (andhraPradeshPayrollRules) {
        out[header] = 'NIL';
        return;
      }
      out[header] = wageValue(payroll.overtimeWages, out[header]);
      return;
    }
    if (isFormXIXMPGrossHeader(header)) {
      out[header] = wageValue(payroll.grossWages, out[header]);
      return;
    }
    if (isFormXIXMPDeductionsHeader(header)) {
      out[header] = wageValue(payroll.deductions, out[header]);
      return;
    }
    if (isFormXIXMPNetHeader(header)) {
      out[header] = wageValue(payroll.netWages, out[header]);
    }
  });
  if (andhraPradeshPayrollRules) {
    applyFormXIXAPWageParticularsToRow(out, hdrs);
  }
  return out;
}

/** AP Form XIX — force overtime NIL and deductions = gross − net from row cells. */
export function applyFormXIXAPWageParticularsToRow(row, headers) {
  if (!row || typeof row !== 'object') return row;
  const hdrs = Array.isArray(headers) ? headers : resolveFormXIXMPWageTableHeaders(headers);
  let gross = null;
  let net = null;
  hdrs.forEach((header) => {
    if (isFormXIXMPOvertimeHeader(header)) {
      row[header] = 'NIL';
      return;
    }
    if (isFormXIXMPGrossHeader(header)) {
      const n = Number(String(row[header] ?? '').replace(/,/g, '').trim());
      if (Number.isFinite(n)) gross = n;
      return;
    }
    if (isFormXIXMPNetHeader(header)) {
      const n = Number(String(row[header] ?? '').replace(/,/g, '').trim());
      if (Number.isFinite(n)) net = n;
    }
  });
  if (gross != null && net != null) {
    const diff = Math.round((gross - net) * 100) / 100;
    hdrs.forEach((header) => {
      if (isFormXIXMPDeductionsHeader(header)) {
        row[header] = diff >= 0 ? String(diff) : '0';
      }
    });
  }
  return row;
}

/** Gujarat Form XIX — fill separate "Name of the workman" header row on export/autofill. */
export function applyFormXIXGJGujaratHeaderExtras(headerData, employee = {}) {
  const out = headerData && typeof headerData === 'object' ? { ...headerData } : {};
  const name = formatWorkmanNameOnly(employee);
  if (name) out.form_xix_gj_workman_name = name;
  return out;
}

/** Push Payroll table values into header form keys (form_xix_ap_days_worked, gross, net, etc.). */
export function applyFormXIXMPPayrollToHeaderData(headerData, payrollRow, options = {}) {
  return applyFormXIXAPAutofillFromSiteAndPayroll(headerData, {
    payroll: resolveFormXIXMPPayrollFields(payrollRow, options),
  });
}

/** Apply Payroll table fields onto mapped grid rows. */
export function enrichFormXIXMPPayrollRows(mappedData, employees, headers, helpers = {}) {
  const hdrs = resolveFormXIXMPWageTableHeaders(headers);
  const {
    resolvePayrollRow = null,
    sanitizeValue = (v) => String(v ?? '').trim(),
    overwrite = true,
    gujaratPayrollRules = false,
    useMonthlyWageRateDefault = false,
    madhyaPradeshPayrollRules = false,
    andhraPradeshPayrollRules = false,
    monthCandidates = null,
  } = helpers;
  if (!Array.isArray(mappedData) || mappedData.length === 0) {
    return 0;
  }
  let hits = 0;
  mappedData.forEach((row, rowIndex) => {
    const empItem = employees[rowIndex];
    const emp = empItem?.Employee || empItem?.employee || empItem;
    const payrollRow =
      typeof resolvePayrollRow === 'function' ? resolvePayrollRow(emp, row, rowIndex) : null;
    if (!payrollRow && !madhyaPradeshPayrollRules && !andhraPradeshPayrollRules) return;
    const merged = applyFormXIXMPEmployeeToRow(row, emp, hdrs, {
      sanitizeValue,
      payrollRow: payrollRow && !payrollRow.fetch_error ? payrollRow : null,
      gujaratPayrollRules,
      useMonthlyWageRateDefault,
      madhyaPradeshPayrollRules,
      andhraPradeshPayrollRules,
      monthCandidates,
    });
    if (!overwrite) {
      hdrs.forEach((header) => {
        if (String(row[header] ?? '').trim() !== '' && String(merged[header] ?? '').trim() !== '') {
          merged[header] = row[header];
        }
      });
    }
    Object.assign(row, merged);
    if (andhraPradeshPayrollRules) {
      applyFormXIXAPWageParticularsToRow(row, hdrs);
    }
    const hasWageData = hdrs.some((header) => {
      if (
        !isFormXIXMPGrossHeader(header) &&
        !isFormXIXMPDeductionsHeader(header) &&
        !isFormXIXMPNetHeader(header) &&
        !isFormXIXMPDaysWorkedHeader(header) &&
        !(andhraPradeshPayrollRules && isFormXIXMPOvertimeHeader(header))
      ) {
        return false;
      }
      return String(row[header] ?? '').trim() !== '';
    });
    if (hasWageData) hits += 1;
  });
  return hits;
}

export function mapFormXIXMPRowsFromEmployees(employees, headers, helpers = {}) {
  const hdrs = resolveFormXIXMPWageTableHeaders(headers);
  const list = Array.isArray(employees) ? employees : [];
  const { resolvePayrollRow = null, ...rest } = helpers;
  return list.map((empItem, rowIndex) => {
    const emp = empItem?.Employee || empItem?.employee || empItem;
    const payrollRow = typeof resolvePayrollRow === 'function' ? resolvePayrollRow(emp, rowIndex) : null;
    return applyFormXIXMPEmployeeToRow({}, emp, hdrs, {
      ...rest,
      payrollRow: payrollRow && !payrollRow.fetch_error ? payrollRow : null,
    });
  });
}

export function overlayFormXIXMPUserEditsOntoRows(mappedRows, tableRows, headers) {
  const hdrs = resolveFormXIXMPWageTableHeaders(headers);
  const savedRows = Array.isArray(tableRows) ? tableRows : [];
  return (Array.isArray(mappedRows) ? mappedRows : []).map((row, index) => {
    const saved = savedRows[index];
    if (!saved || typeof saved !== 'object' || !rowHasMeaningfulFormXIXMPExportData(saved, hdrs)) {
      return row;
    }
    const merged = { ...row };
    hdrs.forEach((header) => {
      const value = getFormXIXMPRowValueForHeader(saved, header);
      if (String(value ?? '').trim() !== '') merged[header] = saved[header] ?? value;
    });
    return merged;
  });
}

export function resolveFormXIXMPExportRows(mappedData, headers, employeesOverride = null, helpers = {}) {
  const hdrs = resolveFormXIXMPWageTableHeaders(headers);
  const tableRows = Array.isArray(mappedData) ? mappedData : [];
  const employees = Array.isArray(employeesOverride) ? employeesOverride : [];
  const {
    sanitizeValue = (v) => String(v ?? '').trim(),
    resolvePayrollRow = null,
  } = helpers;
  if (employees.length > 0) {
    const mapHelpers = { sanitizeValue };
    if (typeof resolvePayrollRow === 'function') {
      mapHelpers.resolvePayrollRow = resolvePayrollRow;
    }
    const fromEmployees = mapFormXIXMPRowsFromEmployees(employees, hdrs, mapHelpers);
    return overlayFormXIXMPUserEditsOntoRows(fromEmployees, tableRows, hdrs);
  }
  return tableRows;
}

const isNarrativeBlob = (raw) => {
  const n = formXIXAPHeaderNorm(raw);
  if (!n) return false;
  return (
    /^form\s*xix\b/i.test(n) ||
    /wage\s+slip/i.test(n) ||
    /wage\s+particulars/i.test(n) ||
    /contract\s+labou?r/i.test(n) ||
    n.length > 160
  );
};

const isWageLabelBlob = (raw) => {
  const n = formXIXAPHeaderNorm(raw).replace(/^\d+[\.\)]\s*/, '');
  return (
    isFormXIXMPDaysWorkedHeader(n) ||
    isFormXIXMPUnitsWorkedHeader(n) ||
    isFormXIXMPRateHeader(n) ||
    isFormXIXMPOvertimeHeader(n) ||
    isFormXIXMPGrossHeader(n) ||
    isFormXIXMPDeductionsHeader(n) ||
    isFormXIXMPNetHeader(n)
  );
};

const buildTemplateField = (spec, coords = {}) => ({
  label: spec.label,
  value: spec.group === 'wages' ? '' : coords.value || '',
  key: spec.key,
  group: spec.group || 'header',
  fieldType: spec.fieldType || 'text',
  labelRow: coords.labelRow ?? null,
  labelCol: coords.labelCol ?? null,
  valueCol: coords.valueCol ?? null,
  valueRow: coords.valueRow ?? null,
});

const findMPStackedLabelCell = (
  getMergedAwareCellText,
  matchRe,
  effectiveSheetCols,
  maxRows = 120,
  usedRows = null
) => {
  const maxC = Math.max(20, effectiveSheetCols || 0);
  const skipRows = usedRows instanceof Set ? usedRows : null;
  for (let r = 0; r < maxRows; r += 1) {
    if (skipRows?.has(r)) continue;
    for (let c = 0; c < Math.min(maxC, 8); c += 1) {
      const raw = String(getMergedAwareCellText(r, c) || '').trim();
      if (!raw || isNarrativeBlob(raw)) continue;
      const norm = formXIXAPHeaderNorm(raw);
      if (!matchRe.test(norm) && !matchRe.test(raw)) continue;
      const value = String(getMergedAwareCellText(r, FORM_XIX_MP_STACKED_VALUE_COL - 1) || '').trim();
      return {
        labelRow: r,
        labelCol: c,
        valueCol: FORM_XIX_MP_STACKED_VALUE_COL - 1,
        valueRow: r,
        value: isNarrativeBlob(value) || isWageLabelBlob(value) ? '' : value,
      };
    }
  }
  return null;
};

const buildFormXIXMPWageFieldsFromSheet = (getMergedAwareCellText, effectiveSheetCols) => {
  const usedRows = new Set();
  return FORM_XIX_MP_WAGE_SPECS.map((spec) => {
    const coords =
      typeof getMergedAwareCellText === 'function' && spec.match
        ? findMPStackedLabelCell(getMergedAwareCellText, spec.match, effectiveSheetCols, 120, usedRows) || {}
        : {};
    if (coords.labelRow != null) usedRows.add(coords.labelRow);
    return buildTemplateField(spec, coords);
  });
};

export function buildFormXIXMPHeaderFields(
  getMergedAwareCellText = null,
  effectiveSheetCols = 20,
  contextHints = {}
) {
  const specs = resolveFormXIXMPHeaderSpecs(
    contextHints.formHeader,
    contextHints.item,
    contextHints.fileName || contextHints.formFileName,
    contextHints.sheetText || ''
  );
  return specs.map((spec) => {
    let coords = {};
    if (typeof getMergedAwareCellText === 'function' && spec.match) {
      coords = findMPStackedLabelCell(getMergedAwareCellText, spec.match, effectiveSheetCols) || {};
    }
    return buildTemplateField(spec, coords);
  });
}

export function resolveFormXIXMPHeaderFieldLayout(parsed, workbook, hints = {}) {
  const formHeader = hints.formHeader || parsed?.formHeader || null;
  const item = hints.item || null;
  const fileName = hints.fileName || hints.formFileName || '';
  const sheetText = hints.sheetText || '';
  if (!isFormXIXMPWageSlipContext(formHeader, item, fileName, sheetText)) return null;

  const accessor = buildFormXIXMPWorkbookAccessor(workbook, hints);
  const getMergedAwareCellText = accessor?.getMergedAwareCellText ?? null;
  const effectiveSheetCols = accessor?.effectiveSheetCols ?? 20;
  const layoutHints = { formHeader, item, fileName, formFileName: fileName, sheetText };
  const headerFields = buildFormXIXMPHeaderFields(getMergedAwareCellText, effectiveSheetCols, layoutHints);
  const wageFields = buildFormXIXMPWageFieldsFromSheet(getMergedAwareCellText, effectiveSheetCols);

  return {
    formHeader: {
      title: formHeader?.title || 'Form XIX – Wage Slip',
      subtitle: formHeader?.subtitle || 'Wage Slip (Contract Labour)',
      reference: formHeader?.reference || '',
      formXIXAPHeaderFieldLayout: true,
      formXIXMPTableLayout: true,
      textRows: [],
      fields: [...headerFields, ...wageFields],
    },
    headers: resolveFormXIXMPWageTableHeaders(parsed?.headers || hints.tableHeaders || null),
    tableData: [],
    headerRowIndex: -1,
    dataStartIndex: 0,
    tableStartCol: 0,
    sheetName: accessor?.sheetName || hints.preferredSheetName || null,
  };
}

function buildFormXIXMPWorkbookAccessor(workbook, hints = {}) {
  if (!workbook?.SheetNames?.length) return null;
  const sheetName =
    hints.preferredSheetName && workbook.Sheets?.[hints.preferredSheetName]
      ? hints.preferredSheetName
      : workbook.SheetNames[0];
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

const formXIXMPExcelCellValueToString = (val) => {
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

const formXIXMPSanitizeExportText = (value) =>
  String(value ?? '')
    .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '')
    .trim();

/** Site / period headers — written once in fast export template (not per employee). */
const FORM_XIX_MP_STATIC_HEADER_KEYS = new Set([
  'form_xix_ap_contractor',
  'form_xix_ap_nature_location',
  'form_xix_ap_period_ending',
  'form_xix_ap_initials',
]);

const formXIXMPEscapeXml = (value) =>
  String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');

const formXIXMPColToLetter = (col) => {
  let result = '';
  let n = col;
  while (n > 0) {
    const rem = (n - 1) % 26;
    result = String.fromCharCode(65 + rem) + result;
    n = Math.floor((n - 1) / 26);
  }
  return result;
};

const formXIXMPToCellRef = (row, col) => `${formXIXMPColToLetter(col)}${row}`;

const formXIXMPUpsertInlineStrCell = (sheetXml, cellRef, value) => {
  const text = formXIXMPEscapeXml(String(value ?? '').trim());
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

const resolveFormXIXMPWorksheetEntry = (zipFiles) =>
  Object.keys(zipFiles)
    .filter((name) => /^xl\/worksheets\/sheet\d+\.xml$/i.test(name))
    .sort((a, b) => a.localeCompare(b, undefined, { numeric: true }))[0] || null;

const pickStaticFormXIXMPHeaderData = (headerFormData) => {
  const out = {};
  if (!headerFormData || typeof headerFormData !== 'object') return out;
  Object.entries(headerFormData).forEach(([key, value]) => {
    if (!FORM_XIX_MP_STATIC_HEADER_KEYS.has(key)) return;
    const text = formXIXMPSanitizeExportText(value);
    if (text) out[key] = text;
  });
  return out;
};

export function canUseFormXIXMPFastExport(parsedFormHeader) {
  if (!parsedFormHeader?.formXIXMPTableLayout) return false;
  const fields = Array.isArray(parsedFormHeader?.fields) ? parsedFormHeader.fields : [];
  const positioned = fields.filter((field) => field?.key && field.labelRow != null).length;
  return positioned >= 4;
}

export function headerFormDataHasFormXIXMPWageValues(headerFormData) {
  if (!headerFormData || typeof headerFormData !== 'object') return false;
  return ['form_xix_ap_gross', 'form_xix_ap_net', 'form_xix_ap_days_worked', 'form_xix_ap_workman'].some(
    (key) => String(headerFormData[key] ?? '').trim() !== ''
  );
}

const resolveFormXIXMPFastExportCellPositions = (parsedFormHeader, worksheet = null) => {
  const positions = [];
  const seenRefs = new Set();
  const seenKeys = new Set();

  const pushPos = (key, row, col) => {
    if (!key || row == null || col == null) return;
    if (seenKeys.has(key)) return;
    const cellRef = formXIXMPToCellRef(row, col);
    if (seenRefs.has(cellRef)) return;
    seenKeys.add(key);
    seenRefs.add(cellRef);
    positions.push({ key, row, col, cellRef });
  };

  if (worksheet) {
    resolveFormXIXMPWageFieldPositions(worksheet, parsedFormHeader).forEach((pos) => {
      if (!pos?.key || FORM_XIX_MP_STATIC_HEADER_KEYS.has(pos.key)) return;
      pushPos(pos.key, pos.row, pos.col);
    });
    const workmanField = resolveFormXIXMPWorkmanHeaderField(parsedFormHeader);
    if (workmanField?.labelRow != null) {
      const row = (workmanField.valueRow ?? workmanField.labelRow) + 1;
      const col =
        workmanField.valueCol != null ? workmanField.valueCol + 1 : FORM_XIX_MP_STACKED_VALUE_COL;
      pushPos('form_xix_ap_workman', row, col);
    }
  }

  const fields = Array.isArray(parsedFormHeader?.fields) ? parsedFormHeader.fields : [];
  fields.forEach((field) => {
    if (!field?.key || FORM_XIX_MP_STATIC_HEADER_KEYS.has(field.key)) return;
    if (seenKeys.has(field.key)) return;
    if (field.labelRow == null) return;
    const row = (field.valueRow ?? field.labelRow) + 1;
    const col = field.valueCol != null ? field.valueCol + 1 : FORM_XIX_MP_STACKED_VALUE_COL;
    pushPos(field.key, row, col);
  });

  if (worksheet) {
    const maxScanRows = Math.max(120, worksheet.rowCount + 10);
    const maxScanCols = 12;
    FORM_XIX_MP_WAGE_SPECS.forEach((spec) => {
      if (seenKeys.has(spec.key)) return;
      for (let r = 1; r <= maxScanRows; r += 1) {
        for (let c = 1; c <= maxScanCols; c += 1) {
          const raw = formXIXMPExcelCellValueToString(worksheet.getCell(r, c)?.value);
          const norm = formXIXAPHeaderNorm(raw);
          if (!spec.match.test(norm) && !spec.match.test(raw)) continue;
          pushPos(spec.key, r, FORM_XIX_MP_STACKED_VALUE_COL);
          return;
        }
      }
    });
    if (!seenKeys.has('form_xix_ap_workman')) {
      for (let r = 1; r <= maxScanRows; r += 1) {
        for (let c = 1; c <= maxScanCols; c += 1) {
          const raw = formXIXMPExcelCellValueToString(worksheet.getCell(r, c)?.value);
          const norm = formXIXAPHeaderNorm(raw);
          if (!/name\s+and\s+father.*workman|father.*husband.*workman/i.test(norm)) continue;
          pushPos('form_xix_ap_workman', r, FORM_XIX_MP_STACKED_VALUE_COL);
          break;
        }
        if (seenKeys.has('form_xix_ap_workman')) break;
      }
    }
    if (!seenKeys.has('form_xix_gj_workman_name')) {
      for (let r = 1; r <= maxScanRows; r += 1) {
        for (let c = 1; c <= maxScanCols; c += 1) {
          const raw = formXIXMPExcelCellValueToString(worksheet.getCell(r, c)?.value);
          const norm = formXIXAPHeaderNorm(raw);
          if (!/^name\s+of\s+the\s+workman/i.test(norm)) continue;
          pushPos('form_xix_gj_workman_name', r, FORM_XIX_MP_STACKED_VALUE_COL);
          break;
        }
        if (seenKeys.has('form_xix_gj_workman_name')) break;
      }
    }
  }

  return positions;
};

const clearFormXIXMPPerEmployeeValueCells = (worksheet, positions) => {
  if (!worksheet || !Array.isArray(positions)) return;
  positions.forEach((pos) => {
    worksheet.getCell(pos.row, pos.col).value = '';
  });
};

const patchFormXIXMPFastSheetXml = (baseSheetXml, mergedHeaderData, positions) => {
  let sheetXml = baseSheetXml;
  positions.forEach((pos) => {
    const value = mergedHeaderData?.[pos.key];
    sheetXml = formXIXMPUpsertInlineStrCell(sheetXml, pos.cellRef, formXIXMPSanitizeExportText(value));
  });
  return sheetXml;
};

const buildFormXIXMPFastXlsxBytes = async (fastTemplate, mergedHeaderData) => {
  const { sheetEntry, baseSheetXml, staticFiles, positions } = fastTemplate;
  const sheetXml = patchFormXIXMPFastSheetXml(baseSheetXml, mergedHeaderData, positions);
  const entryZip = new JSZip();
  Object.entries(staticFiles).forEach(([path, data]) => {
    entryZip.file(path, data);
  });
  entryZip.file(sheetEntry, sheetXml);
  return entryZip.generateAsync({ type: 'uint8array', compression: 'STORE' });
};

export async function prepareFormXIXMPFastExportTemplate({
  templateArrayBuffer,
  parsedFormHeader,
  headerFormData,
  sheetNameHint,
}) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(templateArrayBuffer);
  const sheetCandidates = Array.isArray(workbook.worksheets) ? workbook.worksheets : [];
  const preferred = String(sheetNameHint || '').trim();
  const worksheet =
    (preferred && sheetCandidates.find((ws) => String(ws?.name || '').trim() === preferred)) ||
    sheetCandidates[0] ||
    null;
  if (!worksheet) throw new Error('Template worksheet not found.');

  const positions = resolveFormXIXMPFastExportCellPositions(parsedFormHeader, worksheet);
  clearFormXIXMPPerEmployeeValueCells(worksheet, positions);

  const staticHeaderData = pickStaticFormXIXMPHeaderData(headerFormData);
  if (Object.keys(staticHeaderData).length > 0) {
    writeFormXIXAPFieldsToExcelJsWorksheet(worksheet, staticHeaderData, parsedFormHeader, {
      excelCellValueToString: formXIXMPExcelCellValueToString,
    });
  }

  const preparedBuffer = await workbook.xlsx.writeBuffer();
  const templateZip = await JSZip.loadAsync(preparedBuffer);
  const sheetEntry = resolveFormXIXMPWorksheetEntry(templateZip.files);
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
};

const buildFormXIXMPFastZipDownload = async ({
  exportRows,
  fastTemplate,
  baseHeaderData,
  parsedFormHeader,
  hdrs,
  formFileName,
  employees = [],
  resolvePayrollRow = null,
}) => {
  const rowsForZip = Array.isArray(exportRows) && exportRows.length > 0 ? exportRows : [null];
  const zip = new JSZip();
  const usedNames = new Map();
  for (let i = 0; i < rowsForZip.length; i += 1) {
    const empItem = resolveFormXIXMPDownloadEmployeeForRow(rowsForZip[i], employees, i);
    const payrollRow =
      typeof resolvePayrollRow === 'function' && empItem ? resolvePayrollRow(empItem) : null;
    const mergedHeaderData = sanitizeFormXIXMPHeaderFormData(
      buildEmployeeHeaderFormData(baseHeaderData, rowsForZip[i], parsedFormHeader, payrollRow, empItem)
    );
    const xlsxBytes = await buildFormXIXMPFastXlsxBytes(fastTemplate, mergedHeaderData);
    const baseName = resolveFormXIXMPEmployeeDownloadBaseName(rowsForZip[i], hdrs, i);
    zip.file(allocateUniqueFormXIXMPDownloadFileName(baseName, usedNames), xlsxBytes);
    if (i > 0 && i % 25 === 0) {
      await new Promise((resolve) => setTimeout(resolve, 0));
    }
  }
  const zipBase = String(formFileName || parsedFormHeader?.title || 'Form_XIX_MP')
    .replace(/\.xlsx?$/i, '')
    .replace(/[^a-zA-Z0-9._-]+/g, '_');
  return {
    blob: await zip.generateAsync({ type: 'blob', compression: 'STORE' }),
    fileName: `${zipBase}_Employees.zip`,
  };
};

const resolveFormXIXMPDownloadExportRows = (mappedData, headers, employees, helpers = {}) => {
  const hdrs = resolveFormXIXMPWageTableHeaders(headers);
  const tableRows = Array.isArray(mappedData) ? mappedData : [];
  const fromTable = tableRows.filter((row) => rowHasMeaningfulFormXIXMPExportData(row, hdrs));
  if (fromTable.length > 0) return fromTable;

  const fromEmployees =
    employees.length > 0 ? resolveFormXIXMPExportRows(tableRows, hdrs, employees, helpers) : [];
  const meaningfulFromEmployees = fromEmployees.filter((row) =>
    rowHasMeaningfulFormXIXMPExportData(row, hdrs)
  );
  if (meaningfulFromEmployees.length > 0) return meaningfulFromEmployees;

  if (employees.length > 0) {
    return mapFormXIXMPRowsFromEmployees(employees, hdrs, helpers);
  }
  return [];
};

export function resolveFormXIXMPWageFieldPositions(worksheet, parsedFormHeader) {
  if (!worksheet) return [];
  const fields = Array.isArray(parsedFormHeader?.fields) ? parsedFormHeader.fields : [];
  const positions = [];
  const maxScanRows = Math.max(120, worksheet.rowCount + 10);
  const maxScanCols = 12;

  const pushFromField = (field, tableHeader) => {
    if (field.labelRow == null) return;
    const row = (field.valueRow ?? field.labelRow) + 1;
    const col = field.valueCol != null ? field.valueCol + 1 : FORM_XIX_MP_STACKED_VALUE_COL;
    positions.push({ key: field.key, tableHeader, row, col });
  };

  fields.forEach((field) => {
    if (field.group === 'wages') {
      const tableHeader =
        FORM_XIX_MP_WAGE_TABLE_HEADERS.find((h) => FORM_XIX_MP_WAGE_KEY_BY_HEADER.get(h.replace(/\s*:+\s*$/, '')) === field.key) ||
        field.label;
      pushFromField(field, tableHeader);
    }
  });

  if (positions.length >= FORM_XIX_MP_WAGE_SPECS.length) return positions;

  FORM_XIX_MP_WAGE_SPECS.forEach((spec) => {
    if (positions.some((p) => p.key === spec.key)) return;
    for (let r = 1; r <= maxScanRows; r += 1) {
      for (let c = 1; c <= maxScanCols; c += 1) {
        const raw = formXIXMPExcelCellValueToString(worksheet.getCell(r, c)?.value);
        const norm = formXIXAPHeaderNorm(raw);
        if (!spec.match.test(norm) && !spec.match.test(raw)) continue;
        const tableHeader =
          FORM_XIX_MP_WAGE_TABLE_HEADERS.find(
            (h) => FORM_XIX_MP_WAGE_KEY_BY_HEADER.get(h.replace(/\s*:+\s*$/, '')) === spec.key
          ) || spec.label;
        positions.push({ key: spec.key, tableHeader, row: r, col: FORM_XIX_MP_STACKED_VALUE_COL });
        return;
      }
    }
  });

  return positions;
}

export function resolveFormXIXMPWorkmanHeaderField(parsedFormHeader) {
  const fields = Array.isArray(parsedFormHeader?.fields) ? parsedFormHeader.fields : [];
  return fields.find((f) => f.key === 'form_xix_ap_workman') || null;
}

function buildEmployeeHeaderFormData(headerFormData, employeeRow, parsedFormHeader, payrollRow = null, empItem = null) {
  const base = headerFormData && typeof headerFormData === 'object' ? { ...headerFormData } : {};
  if (employeeRow && typeof employeeRow === 'object') {
    const workmanHeader = FORM_XIX_MP_WAGE_TABLE_HEADERS.find(isFormXIXMPWorkmanNameHeader);
    const workmanValue = workmanHeader ? getFormXIXMPRowValueForHeader(employeeRow, workmanHeader) : '';
    if (workmanValue) base.form_xix_ap_workman = workmanValue;

    FORM_XIX_MP_WAGE_SPECS.forEach((spec) => {
      const tableHeader =
        FORM_XIX_MP_WAGE_TABLE_HEADERS.find(
          (h) => FORM_XIX_MP_WAGE_KEY_BY_HEADER.get(h.replace(/\s*:+\s*$/, '')) === spec.key
        ) || spec.label;
      const value = getFormXIXMPRowValueForHeader(employeeRow, tableHeader);
      if (String(value ?? '').trim() !== '') base[spec.key] = value;
    });
  }

  const emp = unwrapFormXIXMPEmployeeItem(empItem);
  if (emp && typeof emp === 'object') {
    const workmanName = formatWorkmanNameOnly(emp);
    if (workmanName) base.form_xix_gj_workman_name = workmanName;
    const workmanFull = formatWorkmanNameAndGuardian(emp);
    if (workmanFull && !String(base.form_xix_ap_workman ?? '').trim()) {
      base.form_xix_ap_workman = workmanFull;
    }
  } else if (String(base.form_xix_gj_workman_name ?? '').trim() === '') {
    const workmanOnly = String(base.form_xix_ap_workman ?? '')
      .split(/\r?\n/)[0]
      .trim();
    if (workmanOnly) base.form_xix_gj_workman_name = workmanOnly;
  }

  if (payrollRow && !payrollRow.fetch_error) {
    const payrollOptions = resolveFormXIXMPPayrollHelpers({
      formHeader: parsedFormHeader,
    });
    const fromPayroll = applyFormXIXMPPayrollToHeaderData({}, payrollRow, payrollOptions);
    FORM_XIX_MP_WAGE_SPECS.forEach((spec) => {
      const payrollVal = String(fromPayroll[spec.key] ?? '').trim();
      if (payrollVal && String(base[spec.key] ?? '').trim() === '') {
        base[spec.key] = payrollVal;
      }
    });
  }

  return base;
}

const unwrapFormXIXMPEmployeeItem = (empItem) =>
  empItem?.Employee || empItem?.employee || empItem || null;

const resolveFormXIXMPDownloadEmployeeForRow = (row, employees, index) => {
  const list = Array.isArray(employees) ? employees : [];
  if (list.length === 0) return null;
  const lookupId = String(row?.__employeeLookupId ?? '').trim();
  if (lookupId) {
    const hit = list.find((empItem) => {
      const emp = unwrapFormXIXMPEmployeeItem(empItem);
      if (!emp || typeof emp !== 'object') return false;
      const codes = [emp.EmployeeID, emp['EmployeeID'], emp.Zoho_ID, emp['Zoho_ID']]
        .map((value) => String(value ?? '').trim())
        .filter(Boolean);
      return codes.some((code) => code === lookupId);
    });
    if (hit) return hit;
  }
  return list[index] ?? list[0] ?? null;
};

const enrichFormXIXMPExportRowForDownload = (row, empItem, hdrs, helpers = {}) => {
  const emp = unwrapFormXIXMPEmployeeItem(empItem);
  if (!emp || typeof emp !== 'object') return row && typeof row === 'object' ? { ...row } : {};
  const {
    resolvePayrollRow,
    sanitizeValue = (v) => String(v ?? '').trim(),
    gujaratPayrollRules = false,
    useMonthlyWageRateDefault = false,
  } = helpers;
  const payrollRow = typeof resolvePayrollRow === 'function' ? resolvePayrollRow(empItem) : null;
  return applyFormXIXMPEmployeeToRow(row && typeof row === 'object' ? { ...row } : {}, emp, hdrs, {
    sanitizeValue,
    payrollRow: payrollRow && !payrollRow.fetch_error ? payrollRow : null,
    gujaratPayrollRules,
    useMonthlyWageRateDefault,
  });
};

export async function buildFormXIXMPWorkbookWithTemplateStyles({
  templateArrayBuffer,
  mappedData,
  headersToUse,
  parsedFormHeader,
  formFileName,
  headerFormData,
  sheetNameHint,
  employeesOverride = null,
  resolvePayrollRow = null,
}) {
  if (!templateArrayBuffer) {
    throw new Error('Original form template buffer is required for Form XIX MP export.');
  }
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(templateArrayBuffer);
  const sheetCandidates = Array.isArray(workbook.worksheets) ? workbook.worksheets : [];
  const preferred = String(sheetNameHint || '').trim();
  const worksheet =
    (preferred && sheetCandidates.find((ws) => String(ws?.name || '').trim() === preferred)) ||
    sheetCandidates[0] ||
    null;
  if (!worksheet) throw new Error('Template worksheet not found.');

  const hdrs = resolveFormXIXMPWageTableHeaders(headersToUse);
  const sourceRows = (Array.isArray(mappedData) ? mappedData : []).filter((row) =>
    rowHasMeaningfulFormXIXMPExportData(row, hdrs)
  );
  const employeeRow = sourceRows.length === 1 ? sourceRows[0] : sourceRows[0] || null;
  const employees = Array.isArray(employeesOverride) ? employeesOverride : [];
  const empItem = employees.length > 0 ? resolveFormXIXMPDownloadEmployeeForRow(employeeRow, employees, 0) : null;
  const payrollRow =
    typeof resolvePayrollRow === 'function' && empItem ? resolvePayrollRow(empItem) : null;
  const mergedHeaderData = sanitizeFormXIXMPHeaderFormData(
    buildEmployeeHeaderFormData(headerFormData, employeeRow, parsedFormHeader, payrollRow, empItem)
  );

  writeFormXIXAPFieldsToExcelJsWorksheet(worksheet, mergedHeaderData, parsedFormHeader, {
    excelCellValueToString: formXIXMPExcelCellValueToString,
  });

  const out = await workbook.xlsx.writeBuffer();
  const fileName =
    formFileName ||
    parsedFormHeader?.title?.replace(/[^a-zA-Z0-9]/g, '_') ||
    `Form_XIX_MP_${Date.now()}.xlsx`;
  const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  return { blob, fileName };
}

function sanitizeFormXIXMPHeaderFormData(headerFormData) {
  if (!headerFormData || typeof headerFormData !== 'object') return {};
  const out = {};
  Object.entries(headerFormData).forEach(([key, value]) => {
    if (value == null) return;
    out[key] = formXIXMPSanitizeExportText(value);
  });
  return out;
}

export function resolveFormXIXMPEmployeeDownloadBaseName(row, headers, fallbackIndex = 0) {
  const hdrs = resolveFormXIXMPWageTableHeaders(headers);
  const workmanHeader = hdrs.find(isFormXIXMPWorkmanNameHeader);
  const raw = workmanHeader ? String(getFormXIXMPRowValueForHeader(row, workmanHeader) || '').trim() : '';
  const firstLine = raw.split(/\r?\n/)[0] || raw;
  const slug = firstLine
    .replace(/[^\w\s.-]/g, '')
    .replace(/\s+/g, '_')
    .replace(/_+/g, '_')
    .replace(/^_|_$/g, '')
    .slice(0, 48);
  return slug || `Employee_${fallbackIndex + 1}`;
}

export function allocateUniqueFormXIXMPDownloadFileName(baseName, usedNames) {
  const root = String(baseName || 'Employee').trim() || 'Employee';
  const count = usedNames.get(root) || 0;
  usedNames.set(root, count + 1);
  const suffix = count > 0 ? `_${count + 1}` : '';
  return `Form_XIX_MP_${root}${suffix}.xlsx`;
}

export function triggerFormXIXMPZipDownload(blob, fileName) {
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

export async function buildFormXIXMPPerEmployeeDownload({
  templateArrayBuffer,
  mappedData,
  headersToUse,
  parsedFormHeader,
  formFileName,
  headerFormData,
  sheetNameHint,
  employeesOverride = null,
  resolvePayrollRow = null,
}) {
  if (!templateArrayBuffer) {
    throw new Error('Original form template buffer is required for Form XIX MP export.');
  }

  const hdrs = resolveFormXIXMPWageTableHeaders(headersToUse);
  const employees = Array.isArray(employeesOverride) ? employeesOverride : [];
  const payrollHelpers = resolveFormXIXMPPayrollHelpers({ formHeader: parsedFormHeader });
  const exportHelpers = {
    sanitizeValue: (v) => String(v ?? '').trim(),
    resolvePayrollRow: typeof resolvePayrollRow === 'function' ? resolvePayrollRow : null,
    ...payrollHelpers,
  };
  let exportRows = resolveFormXIXMPDownloadExportRows(
    mappedData,
    hdrs,
    employees,
    exportHelpers
  ).filter((row) => rowHasMeaningfulFormXIXMPExportData(row, hdrs));
  exportRows = exportRows.map((row, index) =>
    enrichFormXIXMPExportRowForDownload(
      row,
      resolveFormXIXMPDownloadEmployeeForRow(row, employees, index),
      hdrs,
      exportHelpers
    )
  );
  const baseHeaderData = sanitizeFormXIXMPHeaderFormData(headerFormData);
  const workbookArgs = {
    templateArrayBuffer,
    headersToUse: hdrs,
    parsedFormHeader,
    formFileName,
    sheetNameHint,
  };

  const rowsForZip = exportRows.length > 0 ? exportRows : [null];

  const useFastExport = canUseFormXIXMPFastExport(parsedFormHeader);
  if (useFastExport) {
    const fastTemplate = await prepareFormXIXMPFastExportTemplate({
      templateArrayBuffer,
      parsedFormHeader,
      headerFormData: baseHeaderData,
      sheetNameHint,
    });

    return buildFormXIXMPFastZipDownload({
      exportRows: rowsForZip,
      fastTemplate,
      baseHeaderData,
      parsedFormHeader,
      hdrs,
      formFileName,
      employees,
      resolvePayrollRow: exportHelpers.resolvePayrollRow,
    });
  }

  const zip = new JSZip();
  const usedNames = new Map();
  for (let i = 0; i < rowsForZip.length; i += 1) {
    const empItem = resolveFormXIXMPDownloadEmployeeForRow(rowsForZip[i], employees, i);
    const payrollRow =
      typeof exportHelpers.resolvePayrollRow === 'function' && empItem
        ? exportHelpers.resolvePayrollRow(empItem)
        : null;
    const mergedHeaderData = sanitizeFormXIXMPHeaderFormData(
      buildEmployeeHeaderFormData(baseHeaderData, rowsForZip[i], parsedFormHeader, payrollRow, empItem)
    );
    const mappedRow = rowsForZip[i] && typeof rowsForZip[i] === 'object' ? [rowsForZip[i]] : [];
    const { blob } = await buildFormXIXMPWorkbookWithTemplateStyles({
      ...workbookArgs,
      mappedData: mappedRow,
      headerFormData: mergedHeaderData,
    });
    const xlsxBytes = new Uint8Array(await blob.arrayBuffer());
    const baseName = resolveFormXIXMPEmployeeDownloadBaseName(rowsForZip[i], hdrs, i);
    zip.file(allocateUniqueFormXIXMPDownloadFileName(baseName, usedNames), xlsxBytes);
    if (i > 0 && i % 15 === 0) {
      await new Promise((resolve) => setTimeout(resolve, 0));
    }
  }
  const zipBase = String(formFileName || parsedFormHeader?.title || 'Form_XIX_MP')
    .replace(/\.xlsx?$/i, '')
    .replace(/[^a-zA-Z0-9._-]+/g, '_');
  return {
    blob: await zip.generateAsync({ type: 'blob', compression: 'STORE' }),
    fileName: `${zipBase}_Employees.zip`,
  };
}

export { applyFormXIXAPAutofillFromSiteAndPayroll as applyFormXIXMPAutofillFromSiteAndPayroll };

export function resolveFormXIXHeaderFieldLayout(parsed, workbook, hints = {}) {
  return (
    resolveFormXIXMPHeaderFieldLayout(parsed, workbook, hints) ||
    resolveFormXIXAPHeaderFieldLayoutInner(parsed, workbook, hints)
  );
}
