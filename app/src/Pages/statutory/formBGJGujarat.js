import ExcelJS from 'exceljs';
import {
  flattenPayrollEarningColumns,
  readPayrollForm15WageAmounts,
  readPayrollScalar,
  readPayrollTextScalar,
} from '../../utils/payrollEarnings';
import { ensureExcelJSDataRowsWithBorders } from '../../utils/excelTableBorders';
import { writeStatutoryHeaderFieldsToExcelJsWorksheet } from '../../utils/statutorySiteCompanyHeaders';
import { resolveFormXIXMPPayrollRowForEmployee } from './formXIXMPWageSlip';

/** Gujarat Form B — Register of Wages (Shops & Establishments). */

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

export function isFormBGJPfHeader(h) {
  const s = normHeader(h);
  return s === 'pf' || s.includes('provident fund');
}

export function isFormBGJIncomeTaxHeader(h) {
  const s = normHeader(h);
  return (s.includes('income') && s.includes('tax')) || s === 'tds' || s.includes('tax deducted');
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
  return s.includes('date') && s.includes('payment');
}

export function isFormBGJGrossTotalHeader(header, headers) {
  if (!isFormBGJTotalHeader(header)) return false;
  const list = Array.isArray(headers) ? headers : [];
  const idx = list.indexOf(header);
  const dedStart = list.findIndex(
    (h) => isFormBGJPfHeader(h) || isFormBGJIncomeTaxHeader(h) || isFormBGJDeductionBandHeader(h)
  );
  if (dedStart >= 0 && idx >= dedStart) return false;
  return true;
}

export function isFormBGJDeductionsTotalHeader(header, headers) {
  if (!isFormBGJTotalHeader(header)) return false;
  const list = Array.isArray(headers) ? headers : [];
  const idx = list.indexOf(header);
  const dedStart = list.findIndex(
    (h) => isFormBGJPfHeader(h) || isFormBGJIncomeTaxHeader(h) || isFormBGJDeductionBandHeader(h)
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
  return s.includes('deduction') || isFormBGJPfHeader(h) || isFormBGJIncomeTaxHeader(h);
}

export function isFormBGJRateOfWageHeader(h) {
  const s = normHeader(h);
  return s.includes('rate') && s.includes('wage');
}

export function isFormBGJOthersHeader(h) {
  const s = normHeader(h);
  return s === 'others' || /^others\b/.test(s);
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
    isFormBGJRateOfWageHeader(h) ||
    isFormBGJDaysWorkedHeader(h) ||
    isFormBGJOvertimeHoursHeader(h) ||
    isFormBGJBasicHeader(h) ||
    isFormBGJSpecialBasicHeader(h) ||
    isFormBGJDaHeader(h) ||
    isFormBGJHraHeader(h) ||
    isFormBGJTotalHeader(h) ||
    isFormBGJPfHeader(h) ||
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
  if (n === 'pf' || n.includes('provident fund')) return 'pf';
  if ((n.includes('income') && n.includes('tax')) || n === 'tds') return 'incomeTax';
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
    ? tableHeaders.map((h) => String(h || '').trim()).filter(Boolean)
    : [];
  return parsed.length > 0 ? parsed : [];
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
    incomeTax: '',
    deductionsTotal: '',
    netPay: '',
    paymentDate: '',
  };
  if (!payrollRow || payrollRow.fetch_error) return empty;

  const flat = flattenPayrollEarningColumns(payrollRow);
  const wageAmounts = readPayrollForm15WageAmounts(payrollRow);

  const paidDays = readPayrollScalar(
    { ...flat, ...payrollRow },
    ['paid_days', 'Paid Days', 'days_worked', 'Days Worked', 'paidDays', 'no_of_days_worked'],
    [/^paid_days$/, /paiddays/, /daysworked/]
  );

  let overtimeHours = readPayrollTextScalar(
    { ...flat, ...payrollRow },
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
      { ...flat, ...payrollRow },
      ['total_ot_hours', 'overtime_hours', 'total_overtime_hours', 'ot_hours'],
      [/total.*ot.*hour/i, /^overtime_hours$/, /^ot_hours$/]
    );
  }

  const basic =
    readPayrollScalar(
      { ...flat, ...payrollRow },
      ['basic', 'Basic', 'earned_basic', 'Earned Basic', 'basic_pay', 'Basic Earnings'],
      [/^basic$/, /^earned_basic$/]
    ) || wageAmounts.basic;

  const hra = readPayrollScalar(
    { ...flat, ...payrollRow },
    ['hra', 'HRA', 'hra_fbp', 'HRA FBP', 'house_rent_allowance', 'House Rent Allowance'],
    [/^hra$/, /house.*rent/]
  );

  const grossPay = readPayrollScalar(
    { ...flat, ...payrollRow },
    ['gross_pay', 'Gross Pay', 'grossPay', 'total_earnings'],
    [/^gross_pay$/, /^total_earnings$/]
  );

  const professionalTax = readPayrollScalar(
    { ...flat, ...payrollRow },
    ['professional_tax', 'Professional Tax', 'pt', 'PT'],
    [/^professional_tax$/, /^pt$/]
  );

  let incomeTax = readPayrollScalar(
    { ...flat, ...payrollRow },
    ['income_tax', 'Income Tax', 'tds', 'TDS', 'tax_deducted_at_source'],
    [/income.*tax/i, /^tds$/]
  );
  if (incomeTax === '') {
    incomeTax = readPayrollScalar(
      { ...flat, ...payrollRow },
      ['total_taxes', 'Total Taxes', 'totalTaxes'],
      [/^total_taxes$/]
    );
  }

  const totalDeductions = readPayrollScalar(
    { ...flat, ...payrollRow },
    ['total_deductions', 'Total Deductions', 'totalDeductions', 'total_employee_deductions'],
    [/^total_deductions?$/]
  );
  const totalBenefits = readPayrollScalar(
    { ...flat, ...payrollRow },
    ['total_benefits', 'Total Benefits', 'totalBenefits'],
    [/^total_benefits$/]
  );
  const totalTaxes = readPayrollScalar(
    { ...flat, ...payrollRow },
    ['total_taxes', 'Total Taxes', 'totalTaxes'],
    [/^total_taxes$/]
  );
  const deductionsTotal = sumPayrollNumbers([totalDeductions, totalBenefits, totalTaxes]);

  const netPay = readPayrollScalar(
    { ...flat, ...payrollRow },
    ['net_pay', 'Net Pay', 'netPay', 'monthly_salary'],
    [/^net_pay$/]
  );

  const payDateRaw =
    readPayrollTextScalar(
      { ...flat, ...payrollRow },
      ['pay_date', 'Pay Date', 'payment_date', 'Payment Date', 'paid_date', 'date_of_payment', 'Date of Payment'],
      [/^pay_date$/, /payment.*date/i, /^paid_date$/]
    ) ||
    String(payrollRow?.pay_date ?? flat?.pay_date ?? helpers.payDate ?? '').trim();

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
    const payrollRow =
      typeof resolvePayrollRow === 'function' ? resolvePayrollRow(emp, row, rowIndex) : null;
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

  hdrs.forEach((header) => {
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
    if (isFormBGJOthersHeader(header)) {
      if (overwrite || cellIsEmpty(header)) out[header] = '';
      return;
    }
    if (isFormBGJHraHeader(header)) {
      setCell(header, payroll.hra);
      return;
    }
    if (isFormBGJGrossTotalHeader(header, hdrs)) {
      setCell(header, payroll.grossPay);
      return;
    }
    if (isFormBGJPfHeader(header)) {
      setCell(header, payroll.professionalTax, { allowZero: true });
      return;
    }
    if (isFormBGJIncomeTaxHeader(header)) {
      setCell(header, payroll.incomeTax, { allowZero: true });
      return;
    }
    if (isFormBGJDeductionsTotalHeader(header, hdrs)) {
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
