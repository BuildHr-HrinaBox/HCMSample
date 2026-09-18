import ExcelJS from 'exceljs';
import {
  flattenPayrollEarningColumns,
  readPayrollForm15WageAmounts,
  readPayrollScalar,
  readPayrollTextScalar,
} from '../../utils/payrollEarnings';
import { ensureExcelJSDataRowsWithBorders } from '../../utils/excelTableBorders';
import { writeStatutoryHeaderFieldsToExcelJsWorksheet } from '../../utils/statutorySiteCompanyHeaders';
import {
  collectForm10RowNameParts,
  findForm10PayrollRowByFirstAndLastName,
} from './form10TamilNadu';
import {
  resolveFormXIXMPPayrollRowForEmployee,
  resolveFormXIXMPPayrollRowsForAutofill,
} from './formXIXMPWageSlip';

/** Gujarat Form B — Register of Wages (Shops & Establishments). */

/** True when a Sample Payroll / pay-run row carries PF, VPF, or Income Tax. */
export function formBGJPayrollRowHasSampleDeductionFields(payrollRow) {
  if (!payrollRow || payrollRow.fetch_error) return false;
  const fields = resolveFormBGJGujaratPayrollFields(payrollRow);
  return (
    (fields.pf !== '' && fields.pf != null) ||
    (fields.voluntaryProvidentFund !== '' && fields.voluntaryProvidentFund != null) ||
    (fields.incomeTax !== '' && fields.incomeTax != null)
  );
}

/**
 * Form B must prefer Sample Payroll table rows (PF / VPF / Income Tax columns).
 * Zoho pay-run / list rows often have HRA/Gross but empty deduction scalars —
 * never early-return those when a Sample Payroll snapshot with deductions exists.
 */
export function resolveFormBGJGujaratPayrollRowsForAutofill(
  statutoryPayrollRows,
  monthCandidates = [],
  helpers = {}
) {
  const {
    cachedSampleRows = null,
    bulkSampleRows = null,
    payDate = '',
  } = helpers;

  const stampPayDate = (rows) => {
    const date = String(payDate || '').trim();
    if (!date) return rows;
    return rows.map((row) => {
      if (!row || typeof row !== 'object') return row;
      if (String(row.pay_date || row.payDate || '').trim()) return row;
      return {
        ...row,
        pay_date: date,
        payDate: date,
        payment_date: date,
        date_of_payment: date,
      };
    });
  };

  const normalizeRows = (rows) =>
    (Array.isArray(rows) ? rows : [])
      .filter((row) => row && typeof row === 'object' && row.fetch_error !== true)
      .map((row) => flattenPayrollEarningColumns(row));

  const preferDeductionRows = (rows) => {
    const list = normalizeRows(rows);
    if (list.length === 0) return [];
    const withDeductions = list.filter((row) => formBGJPayrollRowHasSampleDeductionFields(row));
    return stampPayDate(withDeductions.length > 0 ? withDeductions : list);
  };

  const takeIfHasDeductions = (rows) => {
    const preferred = preferDeductionRows(rows);
    if (
      preferred.length > 0 &&
      preferred.some((row) => formBGJPayrollRowHasSampleDeductionFields(row))
    ) {
      return preferred;
    }
    return null;
  };

  const fromCached = takeIfHasDeductions(cachedSampleRows);
  if (fromCached) return fromCached;

  const fromBulk = takeIfHasDeductions(bulkSampleRows);
  if (fromBulk) return fromBulk;

  const fromStatutoryPreferred = takeIfHasDeductions(statutoryPayrollRows);
  if (fromStatutoryPreferred) return fromStatutoryPreferred;

  const xix = resolveFormXIXMPPayrollRowsForAutofill(statutoryPayrollRows, monthCandidates);
  const fromXix = takeIfHasDeductions(xix);
  if (fromXix) return fromXix;

  // Last resort (gross/HRA only) — callers should reload Sample Payroll for PF/VPF/IT.
  const fallbackCandidates = [
    cachedSampleRows,
    bulkSampleRows,
    xix,
    statutoryPayrollRows,
  ];
  for (let i = 0; i < fallbackCandidates.length; i += 1) {
    const fallback = preferDeductionRows(fallbackCandidates[i]);
    if (fallback.length > 0) return fallback;
  }
  return [];
}

export function formBGJGujaratHeaderNorm(txt) {
  return String(txt || '')
    .replace(/\r?\n/g, ' ')
    .replace(/['''`´]/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

const stripLeadingNumber = (s) =>
  String(s || '')
    .replace(/^\d+[\.\)]\s*/, '')
    .trim();

function unwrapEmployeeRecord(emp) {
  return emp?.Employee || emp?.employee || emp;
}

function pickEmployeeValue(emp, keys) {
  const src = unwrapEmployeeRecord(emp);
  if (!src || typeof src !== 'object') return '';
  for (let i = 0; i < keys.length; i += 1) {
    const raw = src[keys[i]];
    if (raw == null || raw === '') continue;
    if (typeof raw === 'object') {
      const nested = String(raw.displayValue ?? raw.name ?? raw.Name ?? '').trim();
      if (nested) return nested;
      continue;
    }
    const text = String(raw).trim();
    if (text) return text;
  }
  return '';
}

function readEmployeeFullName(emp) {
  const src = unwrapEmployeeRecord(emp);
  const fn = pickEmployeeValue(src, ['FirstName', 'First Name', 'firstName']);
  const ln = pickEmployeeValue(src, ['LastName', 'Last Name', 'lastName']);
  const combo = [fn, ln].filter(Boolean).join(' ').trim();
  if (combo) return combo;
  return pickEmployeeValue(src, [
    'DisplayName',
    'Display Name',
    'Employee_Name',
    'Employee Name',
    'full_name',
    'Full Name',
  ]);
}

function readBankAccountNumber(emp) {
  return pickEmployeeValue(emp, [
    'Bank_Account_Number',
    'Bank Account Number',
    'Account_Number',
    'Account Number',
    'AccountNumber',
    'accountNumber',
  ]);
}

function parseFormBGJHeaderLabel(h) {
  const raw = stripLeadingNumber(h);
  const m = String(raw || '').match(/^(.+)_([\s\S]+)$/);
  if (m) return String(m[2] || '').trim();
  return raw;
}

const normHeader = (h) =>
  formBGJGujaratHeaderNorm(parseFormBGJHeaderLabel(h))
    .replace(/[^a-z0-9 ]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

export function isFormBGJSerialHeader(h) {
  const s = normHeader(h);
  return /^sr\.?\s*no|^s\.?\s*no|serial|sl\.?\s*no|^no\.?$/.test(s);
}

export function isFormBGJNameHeader(h) {
  const s = normHeader(h);
  if (!/\bname\b/.test(s)) return false;
  if (/father|spouse|husband|employer|establishment|bank/.test(s)) return false;
  return true;
}

export function isFormBGJDaysWorkedHeader(h) {
  const s = normHeader(h);
  return (s.includes('day') || s.includes('days')) && s.includes('work');
}

export function isFormBGJOvertimeHoursHeader(h) {
  const s = normHeader(h);
  return /\bover[\s-]*time\b/.test(s) && (s.includes('hour') || s.includes('hrs'));
}

export function isFormBGJBasicHeader(h) {
  const s = normHeader(h);
  return s === 'basic' || /^basic\b/.test(s);
}

export function isFormBGJHraHeader(h) {
  const s = normHeader(h);
  return s === 'hra' || /\bhra\b/.test(s) || (s.includes('house') && s.includes('rent'));
}

export function isFormBGJVoluntaryPfHeader(h) {
  const s = normHeader(h);
  const compact = s.replace(/\s+/g, '');
  return (
    compact === 'vpf' ||
    (s.includes('voluntary') && (s.includes('provident') || s.includes('pf') || compact.includes('vpf'))) ||
    s.includes('voluntary provident fund')
  );
}

export function isFormBGJPfHeader(h) {
  const s = normHeader(h);
  if (isFormBGJVoluntaryPfHeader(h)) return false;
  const compact = s.replace(/\s+/g, '');
  return (
    compact === 'pf' ||
    compact === 'epf' ||
    s.includes('provident fund') ||
    s.includes('epf') ||
    /^p\.?\s*f\.?$/.test(s)
  );
}

export function isFormBGJIncomeTaxHeader(h) {
  const s = normHeader(h);
  const compact = s.replace(/\s+/g, '');
  return (
    (s.includes('income') && s.includes('tax')) ||
    compact === 'tds' ||
    compact === 'incometax' ||
    s.includes('tax deducted')
  );
}

export function isFormBGJNetPaymentHeader(h) {
  const s = normHeader(h);
  return s.includes('net') && (s.includes('payment') || s.includes('payable') || s.includes('paid'));
}

export function isFormBGJBankReceiptHeader(h) {
  const s = normHeader(h);
  return (
    (s.includes('receipt') && (s.includes('employee') || s.includes('workman') || s.includes('worker'))) ||
    (s.includes('bank') && (s.includes('transaction') || s.includes('account')))
  );
}

export function isFormBGJPaymentDateHeader(h) {
  const s = normHeader(h);
  return (
    (s.includes('date') && s.includes('payment')) ||
    s === 'pay date' ||
    s === 'date of payment' ||
    s === 'payment date'
  );
}

export function isFormBGJGrossTotalHeader(header, headers, headerIndex = null) {
  if (!isFormBGJTotalHeader(header)) return false;
  const list = Array.isArray(headers) ? headers : [];
  const idx =
    Number.isInteger(headerIndex) && headerIndex >= 0 ? headerIndex : list.indexOf(header);
  if (idx < 0) return false;
  const totalIndices = list
    .map((h, i) => (isFormBGJTotalHeader(h) ? i : -1))
    .filter((i) => i >= 0);
  // Two "Total" columns: first is earnings total (gross pay).
  if (totalIndices.length >= 2) return idx === totalIndices[0];
  // In single-total layouts, "HRA" followed by "Total" means earnings total.
  if (idx > 0 && isFormBGJHraHeader(list[idx - 1])) return true;
  const dedStart = list.findIndex(
    (h) =>
      isFormBGJPfHeader(h) ||
      isFormBGJVoluntaryPfHeader(h) ||
      isFormBGJIncomeTaxHeader(h) ||
      isFormBGJRecoveriesHeader(h) ||
      isFormBGJDeductionBandHeader(h)
  );
  if (dedStart >= 0 && idx >= dedStart) return false;
  return true;
}

export function isFormBGJDeductionsTotalHeader(header, headers, headerIndex = null) {
  if (!isFormBGJTotalHeader(header)) return false;
  const list = Array.isArray(headers) ? headers : [];
  const idx =
    Number.isInteger(headerIndex) && headerIndex >= 0 ? headerIndex : list.indexOf(header);
  if (idx < 0) return false;
  const totalIndices = list
    .map((h, i) => (isFormBGJTotalHeader(h) ? i : -1))
    .filter((i) => i >= 0);
  // Two "Total" columns: second (last) is deductions/recoveries total.
  if (totalIndices.length >= 2) return idx === totalIndices[totalIndices.length - 1];
  if (idx > 0 && isFormBGJHraHeader(list[idx - 1])) return false;
  const norm = normHeader(header);
  if (!norm.includes('deduction') && !norm.includes('recover')) return false;
  const dedStart = list.findIndex(
    (h) =>
      isFormBGJPfHeader(h) ||
      isFormBGJVoluntaryPfHeader(h) ||
      isFormBGJIncomeTaxHeader(h) ||
      isFormBGJRecoveriesHeader(h) ||
      isFormBGJDeductionBandHeader(h)
  );
  return dedStart >= 0 && idx >= dedStart;
}

function isFormBGJTotalHeader(h) {
  const s = normHeader(h);
  if (!s) return false;
  if (s.includes('deduction') && !/^total\b/.test(s)) return false;
  return s === 'total' || /^total\b/.test(s);
}

function isFormBGJDeductionBandHeader(h) {
  const s = normHeader(h);
  return (
    s.includes('deduction') ||
    s.includes('recover') ||
    isFormBGJPfHeader(h) ||
    isFormBGJVoluntaryPfHeader(h) ||
    isFormBGJIncomeTaxHeader(h) ||
    isFormBGJInsuranceHeader(h)
  );
}

export function isFormBGJRateOfWageHeader(h) {
  const s = normHeader(h);
  return s.includes('rate') && s.includes('wage');
}

export function isFormBGJOthersHeader(h) {
  const s = normHeader(h);
  return s === 'others' || /^others\b/.test(s);
}

/** Insurance sits after Others in the recoveries band — leave blank (manual). */
export function isFormBGJInsuranceHeader(h) {
  const s = normHeader(h);
  if (!s || s.includes('income')) return false;
  return s === 'insurance' || /^insurance\b/.test(s);
}

/** Recoveries / recovery total column (not individual PF/VPF/IT lines). */
export function isFormBGJRecoveriesHeader(h) {
  const s = normHeader(h);
  if (!s) return false;
  if (isFormBGJPfHeader(h) || isFormBGJVoluntaryPfHeader(h) || isFormBGJIncomeTaxHeader(h)) {
    return false;
  }
  if (isFormBGJInsuranceHeader(h) || isFormBGJOthersHeader(h)) return false;
  return (
    s === 'recoveries' ||
    s === 'recovery' ||
    /^recoveries\b/.test(s) ||
    /^recovery\b/.test(s) ||
    (s.includes('recover') && (s.includes('total') || s.includes('amount')))
  );
}

function formatBGJPayrollPayDate(payDateRaw) {
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

function resolveFormBGJRateOfWageFromGross(grossPay) {
  const gross = parsePayrollNumber(grossPay);
  if (!Number.isFinite(gross) || gross <= 0) return '';
  return Math.round((gross / 26) * 100) / 100;
}

export function isFormBGJSpecialBasicHeader(h) {
  const s = normHeader(h);
  return s.includes('special') && s.includes('basic');
}

export function isFormBGJDaHeader(h) {
  const s = normHeader(h);
  return s === 'da' || /\bda\b/.test(s) || s.includes('dearness');
}

/** Payroll / wage columns — skip generic People field mapping. Name and bank receipt come from People. */
export function isFormBGJSkipPeopleAutofillHeader(h) {
  return (
    isFormBGJOthersHeader(h) ||
    isFormBGJInsuranceHeader(h) ||
    isFormBGJRecoveriesHeader(h) ||
    isFormBGJRateOfWageHeader(h) ||
    isFormBGJDaysWorkedHeader(h) ||
    isFormBGJOvertimeHoursHeader(h) ||
    isFormBGJBasicHeader(h) ||
    isFormBGJSpecialBasicHeader(h) ||
    isFormBGJDaHeader(h) ||
    isFormBGJHraHeader(h) ||
    isFormBGJTotalHeader(h) ||
    isFormBGJPfHeader(h) ||
    isFormBGJVoluntaryPfHeader(h) ||
    isFormBGJIncomeTaxHeader(h) ||
    isFormBGJNetPaymentHeader(h) ||
    isFormBGJPaymentDateHeader(h)
  );
}

export function formBGJHeaderAliasBucket(norm) {
  const n = String(norm || '').trim();
  if (!n) return '';
  if (/^sr\.?\s*no|^s\.?\s*no|serial|sl\.?\s*no|^no\.?$/.test(n)) return 'sno';
  if (n === 'name' || /^name$/.test(n)) return 'name';
  if ((n.includes('day') || n.includes('days')) && n.includes('work')) return 'daysWorked';
  if (/\bover[\s-]*time\b/.test(n) && (n.includes('hour') || n.includes('hrs'))) return 'overtimeHours';
  if (n === 'basic' || /^basic\b/.test(n)) return 'basic';
  if (n === 'hra' || /\bhra\b/.test(n) || (n.includes('house') && n.includes('rent'))) return 'hra';
  if (
    n === 'vpf' ||
    (n.includes('voluntary') && (n.includes('provident') || n.includes('pf'))) ||
    n.includes('voluntary provident fund')
  ) {
    return 'voluntaryPf';
  }
  if (
    n.replace(/\s+/g, '') === 'pf' ||
    n.replace(/\s+/g, '') === 'epf' ||
    n.includes('provident fund') ||
    n.includes('epf') ||
    /^p\.?\s*f\.?$/.test(n)
  ) {
    return 'pf';
  }
  if ((n.includes('income') && n.includes('tax')) || n === 'tds' || n.replace(/\s+/g, '') === 'incometax') {
    return 'incomeTax';
  }
  if (n.includes('net') && (n.includes('payment') || n.includes('payable') || n.includes('paid'))) {
    return 'netPay';
  }
  if (
    (n.includes('receipt') && (n.includes('employee') || n.includes('workman') || n.includes('worker'))) ||
    (n.includes('bank') && (n.includes('transaction') || n.includes('account')))
  ) {
    return 'bankReceipt';
  }
  if (n.includes('date') && n.includes('payment')) return 'paymentDate';
  if (n.includes('rate') && n.includes('wage')) return 'rateOfWage';
  if (n === 'others' || /^others\b/.test(n)) return 'others';
  if (n === 'insurance' || /^insurance\b/.test(n)) return 'insurance';
  if (
    n === 'recoveries' ||
    n === 'recovery' ||
    /^recoveries\b/.test(n) ||
    /^recovery\b/.test(n) ||
    (n.includes('recover') && (n.includes('total') || n.includes('amount')))
  ) {
    return 'recoveries';
  }
  if (n === 'total' || /^total\b/.test(n)) return 'total';
  return n;
}

export function headersIndicateFormBGJGujaratTable(tableHeaders) {
  if (!Array.isArray(tableHeaders) || tableHeaders.length < 5) return false;
  const joined = tableHeaders.map((h) => formBGJGujaratHeaderNorm(h)).join('\n');
  const hasName = /\bname\b/.test(joined);
  const hasDaysWorked = /day/.test(joined) && /work/.test(joined);
  const hasBasic = /\bbasic\b/.test(joined);
  const hasWageRegister =
    (/rate/.test(joined) && /wage/.test(joined)) ||
    /\bhra\b/.test(joined) ||
    /gross/.test(joined) ||
    /net\s+payment/.test(joined);
  const looksLikeFormAEmployeeRegister =
    /education/.test(joined) &&
    (/type/.test(joined) && /employ/.test(joined)) &&
    (/national/.test(joined) || /aadha?ar/.test(joined));
  return hasName && hasDaysWorked && hasBasic && hasWageRegister && !looksLikeFormAEmployeeRegister;
}

export function isFormBGJGujaratContext(
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

  if (/tamil\s*nadu|labour\s*welfare|\blwf\b/.test(parts)) return false;

  const hasGujarat = /gujarat|\b_gj\b|form[\s._-]*b[\s._-]*gj|form_b_gj/.test(parts);
  const hasFormB = /\bform[\s._-]*b\b/.test(parts);

  if (hasGujarat && hasFormB) return true;
  if (headersIndicateFormBGJGujaratTable(tableHeaders) && hasGujarat && hasFormB) return true;
  if (headersIndicateFormBGJGujaratTable(tableHeaders) && hasGujarat) return true;
  if (headersIndicateFormBGJGujaratTable(tableHeaders) && hasFormB && hasGujarat) return true;

  return false;
}

export function resolveFormBGJGujaratTableHeaders(tableHeaders) {
  const parsed = Array.isArray(tableHeaders)
    ? tableHeaders
        .map((h) => String(h ?? ''))
        .filter((h) => String(h).trim() !== '')
    : [];
  if (parsed.length === 0) return [];

  // Keep duplicate captions addressable as separate object keys.
  // Example: "Total" + "Total" should become "Total" + "Total ".
  const seen = new Map();
  return parsed.map((header) => {
    const count = seen.get(header) || 0;
    seen.set(header, count + 1);
    return count > 0 ? `${header}${' '.repeat(count)}` : header;
  });
}

function parsePayrollNumber(value) {
  const n = Number(String(value ?? '').replace(/[,₹]/g, '').trim());
  return Number.isFinite(n) ? n : NaN;
}

function sumPayrollNumbers(values) {
  let sum = 0;
  let any = false;
  values.forEach((value) => {
    const n = parsePayrollNumber(value);
    if (Number.isFinite(n)) {
      sum += n;
      any = true;
    }
  });
  return any ? Math.round(sum * 100) / 100 : '';
}

export function resolveFormBGJGujaratPayrollFields(payrollRow, helpers = {}) {
  const empty = {
    paidDays: '',
    overtimeHours: '',
    basic: '',
    hra: '',
    grossPay: '',
    rateOfWage: '',
    professionalTax: '',
    pf: '',
    voluntaryProvidentFund: '',
    incomeTax: '',
    deductionsTotal: '',
    netPay: '',
    paymentDate: '',
  };
  if (!payrollRow || payrollRow.fetch_error) return empty;

  const flat = flattenPayrollEarningColumns(payrollRow);
  const source = { ...payrollRow, ...flat };
  const wageAmounts = readPayrollForm15WageAmounts(payrollRow);

  const paidDays = readPayrollScalar(
    source,
    ['paid_days', 'Paid Days', 'Paid_days', 'days_worked', 'Days Worked', 'paidDays', 'no_of_days_worked'],
    [/^paid_days$/, /paiddays/, /daysworked/]
  );

  let overtimeHours = readPayrollTextScalar(
    source,
    [
      'total_ot_hours',
      'total overtime hours',
      'overtime_hours',
      'Overtime Hours',
      'total_overtime_hours',
      'ot_hours',
      'overtime_hrs',
    ],
    [/total.*ot.*hour/i, /^total_ot_hours$/, /^overtime_hours$/, /^ot_hours$/]
  );
  if (overtimeHours === '') {
    overtimeHours = readPayrollScalar(
      source,
      ['total_ot_hours', 'overtime_hours', 'total_overtime_hours', 'ot_hours'],
      [/total.*ot.*hour/i, /^overtime_hours$/, /^ot_hours$/]
    );
  }

  const basic =
    readPayrollScalar(
      source,
      ['basic', 'Basic', 'earned_basic', 'Earned Basic', 'basic_pay', 'Basic Earnings'],
      [/^basic$/, /^earned_basic$/]
    ) || wageAmounts.basic;

  const hra = readPayrollScalar(
    source,
    ['hra', 'HRA', 'hra_fbp', 'HRA FBP', 'house_rent_allowance', 'House Rent Allowance'],
    [/^hra$/, /house.*rent/]
  );

  const grossPay = readPayrollScalar(
    source,
    ['gross_pay', 'Gross Pay', 'grossPay', 'gross', 'Gross', 'total_earnings'],
    [/^gross_pay$/, /^gross$/, /^total_earnings$/]
  );

  const professionalTax = readPayrollScalar(
    source,
    ['professional_tax', 'Professional Tax', 'ProfessionalTax', 'professionalTax', 'pt', 'PT'],
    [/^professional_tax$/, /^pt$/]
  );

  // Prefer flatten() computed Sample Payroll / benefit columns first.
  const pickAmount = (...values) => {
    for (let i = 0; i < values.length; i += 1) {
      const value = values[i];
      if (value === '' || value == null) continue;
      const n = parsePayrollNumber(value);
      if (Number.isFinite(n)) return n;
    }
    return '';
  };

  const pf = pickAmount(
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

  const voluntaryProvidentFund = pickAmount(
    flat.voluntary_provident_fund,
    flat.vpf,
    flat.VoluntaryProvidentFund,
    flat.voluntaryProvidentFund,
    readPayrollScalar(
      source,
      [
        'voluntary_provident_fund',
        'VoluntaryProvidentFund',
        'Voluntary Provident Fund',
        'voluntaryProvidentFund',
        'vpf',
        'VPF',
      ],
      [/^vpf$/, /voluntary.*provident/i]
    )
  );

  const incomeTax = pickAmount(
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

  const netPay = readPayrollScalar(
    source,
    ['net_pay', 'Net Pay', 'netPay', 'netpay', 'Netpay', 'monthly_salary'],
    [/^net_pay$/, /^netpay$/]
  );

  // Recoveries (after earnings Total / PF band) ← gross_pay − net_pay.
  const grossN = parsePayrollNumber(grossPay);
  const netN = parsePayrollNumber(netPay);
  let deductionsTotal = '';
  if (Number.isFinite(grossN) && Number.isFinite(netN)) {
    deductionsTotal = Math.round((grossN - netN) * 100) / 100;
    if (deductionsTotal < 0) deductionsTotal = '';
  }
  if (deductionsTotal === '') {
    const totalDeductions = readPayrollScalar(
      source,
      [
        'total_deductions',
        'Total Deductions',
        'totalDeductions',
        'TotalDeduction',
        'totalDeduction',
        'total_employee_deductions',
      ],
      [/^total_deductions?$/]
    );
    const totalBenefits = readPayrollScalar(
      source,
      ['total_benefits', 'Total Benefits', 'totalBenefits'],
      [/^total_benefits$/]
    );
    const totalTaxes = readPayrollScalar(
      source,
      ['total_taxes', 'Total Taxes', 'totalTaxes'],
      [/^total_taxes$/]
    );
    deductionsTotal = sumPayrollNumbers([totalDeductions, totalBenefits, totalTaxes]);
  }

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
    String(
      payrollRow?.pay_date ??
        payrollRow?.payDate ??
        flat?.pay_date ??
        flat?.payDate ??
        helpers.payDate ??
        ''
    ).trim();

  const paymentDate =
    formatBGJPayrollPayDate(payDateRaw) ||
    (typeof helpers.formatStatutoryDateDisplay === 'function'
      ? helpers.formatStatutoryDateDisplay(payDateRaw)
      : payDateRaw);

  const rateOfWage = resolveFormBGJRateOfWageFromGross(grossPay);

  return {
    paidDays,
    overtimeHours,
    basic,
    hra,
    grossPay,
    rateOfWage,
    professionalTax,
    pf,
    voluntaryProvidentFund,
    incomeTax,
    deductionsTotal,
    netPay,
    paymentDate,
  };
}

function formatCellValue(value, { allowZero = false } = {}) {
  if (value == null || value === '') return '';
  const num = parsePayrollNumber(value);
  if (Number.isFinite(num) && (allowZero || num !== 0)) return num;
  return String(value).trim();
}

export function enrichFormBGJGujaratPayrollRows(mappedData, employees, headers, helpers = {}) {
  const hdrs = resolveFormBGJGujaratTableHeaders(headers);
  const {
    resolvePayrollRow = null,
    payrollRows = null,
    sanitizeValue = (v) => String(v ?? '').trim(),
    formatStatutoryDateDisplay = (v) => String(v || '').trim(),
    payDate = '',
    overwrite = true,
    rowIndexOffset = 0,
  } = helpers;
  if (!Array.isArray(mappedData) || mappedData.length === 0 || typeof resolvePayrollRow !== 'function') {
    return 0;
  }
  let hits = 0;
  mappedData.forEach((row, rowIndex) => {
    const empItem = employees[rowIndex];
    const emp = unwrapEmployeeRecord(empItem);
    let payrollRow =
      typeof resolvePayrollRow === 'function' ? resolvePayrollRow(emp, row, rowIndex) : null;
    if ((!payrollRow || payrollRow.fetch_error) && Array.isArray(payrollRows) && payrollRows.length > 0) {
      const extraParts = collectForm10RowNameParts(row, hdrs);
      payrollRow = findForm10PayrollRowByFirstAndLastName(emp || row, payrollRows, extraParts);
    }
    const merged = applyFormBGJGujaratEmployeeToRow(row, emp, hdrs, {
      sanitizeValue,
      formatStatutoryDateDisplay,
      payDate,
      payrollRow: payrollRow && !payrollRow.fetch_error ? payrollRow : null,
      rowIndex: rowIndexOffset + rowIndex,
      overwrite,
    });
    Object.assign(row, merged);
    hits += 1;
  });
  return hits;
}

export function applyFormBGJGujaratEmployeeToRow(row, emp, headers, helpers = {}) {
  const hdrs = resolveFormBGJGujaratTableHeaders(headers);
  const out = row && typeof row === 'object' ? { ...row } : {};
  const {
    sanitizeValue = (v) => String(v ?? '').trim(),
    payrollRow = null,
    rowIndex = 0,
    formatStatutoryDateDisplay = (v) => String(v || '').trim(),
    payDate = '',
    overwrite = true,
  } = helpers;

  const cellIsEmpty = (header) => {
    const v = String(out[header] || '').trim();
    if (!v) return true;
    const s = v.toLowerCase();
    return /^enter\b/.test(s) || s.includes('enter ');
  };
  const setCell = (header, value, opts = {}) => {
    if (!header || value == null || value === '') return;
    if (!overwrite && !cellIsEmpty(header)) return;
    out[header] = sanitizeValue(formatCellValue(value, opts));
  };

  const payroll = resolveFormBGJGujaratPayrollFields(
    payrollRow && !payrollRow.fetch_error ? payrollRow : null,
    { formatStatutoryDateDisplay, payDate }
  );
  const fullName = readEmployeeFullName(emp);
  if (fullName) out.__employeeLookupName = fullName;
  const firstTotalHeaderIndex = hdrs.findIndex((h) => isFormBGJTotalHeader(h));

  hdrs.forEach((header, headerIndex) => {
    if (isFormBGJSerialHeader(header)) {
      setCell(header, String(rowIndex + 1), { allowZero: true });
      return;
    }
    if (isFormBGJNameHeader(header)) {
      setCell(header, fullName);
      return;
    }
    if (isFormBGJDaysWorkedHeader(header)) {
      setCell(header, payroll.paidDays, { allowZero: true });
      return;
    }
    if (isFormBGJOvertimeHoursHeader(header)) {
      setCell(header, payroll.overtimeHours, { allowZero: true });
      return;
    }
    if (isFormBGJBasicHeader(header)) {
      setCell(header, payroll.basic);
      return;
    }
    if (isFormBGJRateOfWageHeader(header)) {
      setCell(header, payroll.rateOfWage);
      return;
    }
    if (isFormBGJOthersHeader(header) || isFormBGJInsuranceHeader(header)) {
      // Insurance is after Others — do not put recoveries here.
      if (overwrite || cellIsEmpty(header)) out[header] = '';
      return;
    }
    if (isFormBGJHraHeader(header)) {
      setCell(header, payroll.hra);
      return;
    }
    if (isFormBGJGrossTotalHeader(header, hdrs, headerIndex)) {
      setCell(header, payroll.grossPay);
      return;
    }
    if (isFormBGJVoluntaryPfHeader(header)) {
      setCell(header, payroll.voluntaryProvidentFund, { allowZero: true });
      return;
    }
    if (isFormBGJPfHeader(header)) {
      // If PF appears before the first Total, use gross pay per template mapping.
      if (firstTotalHeaderIndex >= 0 && headerIndex < firstTotalHeaderIndex) {
        setCell(header, payroll.grossPay, { allowZero: true });
      } else {
        setCell(header, payroll.pf, { allowZero: true });
      }
      return;
    }
    if (isFormBGJIncomeTaxHeader(header)) {
      setCell(header, payroll.incomeTax, { allowZero: true });
      return;
    }
    if (isFormBGJRecoveriesHeader(header)) {
      if (overwrite || cellIsEmpty(header)) out[header] = '';
      return;
    }
    if (isFormBGJDeductionsTotalHeader(header, hdrs, headerIndex)) {
      setCell(header, payroll.deductionsTotal, { allowZero: true });
      return;
    }
    if (isFormBGJNetPaymentHeader(header)) {
      setCell(header, payroll.netPay);
      return;
    }
    if (isFormBGJBankReceiptHeader(header)) {
      setCell(header, readBankAccountNumber(emp));
      return;
    }
    if (isFormBGJPaymentDateHeader(header)) {
      setCell(header, payroll.paymentDate);
    }
  });

  return out;
}

export function getFormBGJGujaratRowValueForHeader(row, header, rowIndex = 0) {
  if (!row || !header) return '';
  const direct = row[header];
  if (direct != null && String(direct).trim() !== '') return String(direct).trim();
  const bucket = formBGJHeaderAliasBucket(formBGJGujaratHeaderNorm(header));
  for (const [k, v] of Object.entries(row)) {
    if (
      formBGJHeaderAliasBucket(formBGJGujaratHeaderNorm(k)) === bucket &&
      v != null &&
      String(v).trim() !== ''
    ) {
      return String(v).trim();
    }
  }
  if (bucket === 'sno') return String(rowIndex + 1);
  return '';
}

export function rowHasMeaningfulFormBGJGujaratExportData(row, headers) {
  const hdrs = resolveFormBGJGujaratTableHeaders(headers);
  return hdrs.some((header) => {
    const bucket = formBGJHeaderAliasBucket(formBGJGujaratHeaderNorm(header));
    if (bucket === 'sno') return false;
    return getFormBGJGujaratRowValueForHeader(row, header) !== '';
  });
}

export function remapFormBGJGujaratRowsToHeaders(rows, sourceHeaders, targetHeaders) {
  const src = Array.isArray(sourceHeaders) ? sourceHeaders : [];
  const tgt = resolveFormBGJGujaratTableHeaders(targetHeaders);
  if (!Array.isArray(rows)) return [];
  return rows.map((row, rowIndex) => {
    if (!row || typeof row !== 'object') return {};
    const out = {};
    tgt.forEach((targetHeader, colIdx) => {
      let val = '';
      if (Object.prototype.hasOwnProperty.call(row, targetHeader)) {
        val = row[targetHeader];
      } else if (src[colIdx] && Object.prototype.hasOwnProperty.call(row, src[colIdx])) {
        val = row[src[colIdx]];
      } else {
        const bucket = formBGJHeaderAliasBucket(formBGJGujaratHeaderNorm(targetHeader));
        for (const [k, v] of Object.entries(row)) {
          if (formBGJHeaderAliasBucket(formBGJGujaratHeaderNorm(k)) === bucket) {
            val = v;
            break;
          }
        }
      }
      if (
        (val == null || val === '') &&
        formBGJHeaderAliasBucket(formBGJGujaratHeaderNorm(targetHeader)) === 'sno'
      ) {
        val = rowIndex + 1;
      }
      out[targetHeader] = val == null ? '' : val;
    });
    return out;
  });
}

export function filterFormBGJGujaratExportRows(rows, headers) {
  const hdrs = resolveFormBGJGujaratTableHeaders(headers);
  return (Array.isArray(rows) ? rows : []).filter((row) =>
    rowHasMeaningfulFormBGJGujaratExportData(row, hdrs)
  );
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

function excelCellLooksLikeSerialHeader(text) {
  const t = formBGJGujaratHeaderNorm(text).replace(/[^a-z0-9 ]/g, ' ').replace(/\s+/g, ' ').trim();
  return /^sr\.?\s*no|^s\.?\s*no|serial|sl\.?\s*no/.test(t);
}

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
        if (r >= tl.row && r <= br.row && c >= tl.col && c <= br.col) {
          topLeft = { r: tl.row, c: tl.col };
          break;
        }
      }
    }
    cache.set(key, topLeft);
    return topLeft;
  };
}

function detectFormBGJGujaratTableLayout(worksheet, hints = {}) {
  const getMergeTopLeft = buildMergeTopLeftResolver(worksheet);
  const getMergedAwareCellText = (r, c) => {
    const tl = getMergeTopLeft(r, c);
    return excelCellValueToString(worksheet.getCell(tl.r, tl.c)?.value).trim();
  };

  let headerRow =
    hints.parsedHeaderRowIndex != null && hints.parsedHeaderRowIndex >= 0
      ? hints.parsedHeaderRowIndex + 1
      : -1;
  let startCol =
    hints.parsedTableStartCol != null && hints.parsedTableStartCol >= 0
      ? hints.parsedTableStartCol + 1
      : 1;

  if (headerRow < 1) {
    const maxScanRows = Math.max(35, worksheet.rowCount + 5);
    for (let r = 1; r <= maxScanRows; r += 1) {
      for (let c = 1; c <= 30; c += 1) {
        if (excelCellLooksLikeSerialHeader(getMergedAwareCellText(r, c))) {
          headerRow = r;
          startCol = c;
          break;
        }
      }
      if (headerRow > 0) break;
    }
  }
  if (headerRow < 1) return null;

  const templateCols = [];
  for (let c = startCol; c <= startCol + 40; c += 1) {
    const label = getMergedAwareCellText(headerRow, c);
    if (!label) {
      if (templateCols.length >= 8) break;
      continue;
    }
    const bucket = formBGJHeaderAliasBucket(formBGJGujaratHeaderNorm(label));
    templateCols.push({ col: c, label, bucket });
    if (templateCols.length >= 30) break;
  }
  if (templateCols.length < 4) return null;

  const dataStartRow =
    hints.parsedDataStartIndex != null && hints.parsedDataStartIndex >= 0
      ? hints.parsedDataStartIndex + 1
      : headerRow + 1;

  return { headerRow, dataStartRow, templateCols, startCol };
}

export async function buildFormBGJGujaratWorkbookWithTemplateStyles({
  templateArrayBuffer,
  mappedData,
  headersToUse,
  parsedFormHeader,
  formFileName,
  headerFormData,
  parsedHeaderRowIndex,
  parsedDataStartIndex,
  parsedTableStartCol,
  formatStatutoryDateDisplay = null,
}) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(templateArrayBuffer);
  const worksheet = workbook.worksheets[0];
  if (!worksheet) throw new Error('Template worksheet not found.');

  const layout = detectFormBGJGujaratTableLayout(worksheet, {
    parsedHeaderRowIndex,
    parsedDataStartIndex,
    parsedTableStartCol,
  });
  if (!layout) throw new Error('Could not locate Gujarat Form B table header row.');

  const { dataStartRow, templateCols } = layout;

  if (headerFormData && typeof headerFormData === 'object') {
    writeStatutoryHeaderFieldsToExcelJsWorksheet(worksheet, {
      headerFormData,
      parsedFormHeader,
      headerRowEnd: Math.max(1, layout.headerRow - 1),
      maxScanCols: 80,
    });
  }

  const templateHeaderLabels = templateCols.map((entry) => entry.label);
  const normalizedHeaders = resolveFormBGJGujaratTableHeaders(
    Array.isArray(headersToUse) && headersToUse.length >= templateHeaderLabels.length
      ? headersToUse
      : templateHeaderLabels.length > 0
        ? templateHeaderLabels
        : headersToUse
  );
  const rows = filterFormBGJGujaratExportRows(
    remapFormBGJGujaratRowsToHeaders(
      Array.isArray(mappedData) ? mappedData : [],
      normalizedHeaders,
      normalizedHeaders
    ),
    normalizedHeaders
  );

  const tableColMin = Math.min(...templateCols.map(({ col }) => col));
  const tableColMax = Math.max(...templateCols.map(({ col }) => col));
  const clearToRow = Math.max(dataStartRow + rows.length + 10, dataStartRow + 15);
  for (let r = dataStartRow; r <= clearToRow; r += 1) {
    for (let c = tableColMin; c <= tableColMax; c += 1) {
      worksheet.getCell(r, c).value = '';
    }
  }

  rows.forEach((row, idx) => {
    const targetRow = worksheet.getRow(dataStartRow + idx);
    if (targetRow) targetRow.height = 18;
    templateCols.forEach(({ col, bucket, label }) => {
      let val = getFormBGJGujaratRowValueForHeader(row, label, idx);
      if ((val == null || val === '') && bucket) {
        for (const [k, v] of Object.entries(row)) {
          if (
            formBGJHeaderAliasBucket(formBGJGujaratHeaderNorm(k)) === bucket &&
            v != null &&
            String(v).trim() !== ''
          ) {
            val = String(v).trim();
            break;
          }
        }
      }
      if ((val == null || val === '') && bucket === 'sno') val = String(idx + 1);
      const cell = worksheet.getCell(dataStartRow + idx, col);
      if (val == null || val === '') {
        cell.value = '';
        return;
      }
      if (bucket === 'sno') {
        const n = Number(String(val).replace(/[,]/g, '').trim());
        cell.value = Number.isFinite(n) ? n : String(val);
      } else if (bucket === 'bankReceipt') {
        const digits = String(val).replace(/\D/g, '');
        cell.value = digits || String(val);
        if (digits) cell.numFmt = '0';
      } else {
        const n = Number(String(val).replace(/[,₹]/g, '').trim());
        cell.value = Number.isFinite(n) ? n : String(val);
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
  }

  const out = await workbook.xlsx.writeBuffer();
  const fileName =
    formFileName ||
    parsedFormHeader?.title?.replace(/[^a-zA-Z0-9]/g, '_') ||
    'Form_B_GJ_Gujarat.xlsx';
  return {
    blob: new Blob([out], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    }),
    fileName,
  };
}

export { resolveFormXIXMPPayrollRowForEmployee as resolveFormBGJGujaratPayrollRowForEmployee };
