import ExcelJS from 'exceljs';
import {
  flattenPayrollEarningColumns,
  readPayrollForm15WageAmounts,
  readPayrollScalar,
  readPayrollTextScalar,
} from '../../utils/payrollEarnings';
import { ensureExcelJSDataRowsWithBorders } from '../../utils/excelTableBorders';
import { writeStatutoryHeaderFieldsToExcelJsWorksheet } from '../../utils/statutorySiteCompanyHeaders';
import { resolveFormXIXMPPayrollRowForEmployee, resolvePayrollRowByFormTableName } from './formXIXMPWageSlip';

/** Rajasthan Form B — Register of Wages. */

export const FORM_B_RJ_NIL = 'Nil';

export function formBRajasthanHeaderNorm(txt) {
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

function readEmployeeId(emp) {
  return pickEmployeeValue(emp, [
    'EmployeeID',
    'Employee ID',
    'Employee_ID',
    'EmployeeId',
    'employeeId',
    'Employee.ID',
    'erecno',
    'Erecno',
    'Zoho_ID',
  ]);
}

function parseFormBRJHeaderLabel(h) {
  const raw = stripLeadingNumber(h);
  const m = String(raw || '').match(/^(.+)_([\s\S]+)$/);
  if (m) return String(m[2] || '').trim();
  return raw;
}

const normHeader = (h) =>
  formBRajasthanHeaderNorm(parseFormBRJHeaderLabel(h))
    .replace(/[^a-z0-9 ]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

export function isFormBRJSerialHeader(h) {
  const s = normHeader(h);
  return /^sr\.?\s*no|^s\.?\s*no|serial|sl\.?\s*no|^no\.?$/.test(s);
}

export function isFormBRJNameHeader(h) {
  const s = normHeader(h);
  if (!/\bname\b/.test(s)) return false;
  if (/father|spouse|husband|employer|establishment|bank/.test(s)) return false;
  return true;
}

export function isFormBRJEmployeeCodeHeader(h) {
  const s = normHeader(h);
  return (
    (s.includes('employee') && s.includes('code')) ||
    s === 'employee code' ||
    s === 'emp code'
  );
}

export function isFormBRJDaysWorkedHeader(h) {
  const s = normHeader(h);
  return (s.includes('day') || s.includes('days')) && s.includes('work');
}

export function isFormBRJOvertimeHoursHeader(h) {
  const s = normHeader(h);
  return /\bover[\s-]*time\b/.test(s) && (s.includes('hour') || s.includes('hrs') || s.includes('worked'));
}

export function isFormBRJPaymentsOvertimeHeader(h) {
  const s = normHeader(h);
  if (!/\bover[\s-]*time\b/.test(s)) return false;
  if (s.includes('hour') || s.includes('hrs') || s.includes('worked')) return false;
  return s.includes('payment') || s.includes('pay');
}

export function isFormBRJBasicHeader(h) {
  const s = normHeader(h);
  return (s === 'basic' || /^basic\b/.test(s)) && !s.includes('special');
}

export function isFormBRJHraHeader(h) {
  const s = normHeader(h);
  return s === 'hra' || /\bhra\b/.test(s) || (s.includes('house') && s.includes('rent'));
}

export function isFormBRJSpecialBasicHeader(h) {
  const s = normHeader(h);
  return s.includes('special') && s.includes('basic');
}

export function isFormBRJDaHeader(h) {
  const s = normHeader(h);
  return s === 'da' || /\bda\b/.test(s) || s.includes('dearness');
}

function isFormBRJPlainTotalHeader(h) {
  const s = normHeader(h);
  if (!s) return false;
  if (s.includes('deduction') || s.includes('deducation')) return false;
  return s === 'total' || /^total\b/.test(s);
}

export function isFormBRJOthersHeader(h) {
  const s = normHeader(h);
  return s === 'others' || /^others\b/.test(s);
}

/** Deduction total / deducation — gross_pay − net_pay (not gross_pay). */
export function isFormBRJDeductionTotalHeader(header, headers = [], headerIndex = null) {
  const s = normHeader(header);
  if (!s) return false;
  if (
    (s.includes('deduction') || s.includes('deducation')) &&
    (s.includes('total') || s === 'deduction' || s === 'deducation')
  ) {
    return true;
  }
  const list = Array.isArray(headers) ? headers : [];
  if (!isFormBRJPlainTotalHeader(header)) return false;
  const idx =
    Number.isInteger(headerIndex) && headerIndex >= 0 ? headerIndex : list.indexOf(header);
  if (idx < 0) return false;
  const totalIndices = list
    .map((h, i) => (isFormBRJPlainTotalHeader(h) ? i : -1))
    .filter((i) => i >= 0);
  // Last plain "Total" is Deduction → Total (duplicate labels break indexOf).
  return totalIndices.length >= 2 && idx === totalIndices[totalIndices.length - 1];
}

/** Earnings total — Basic + HRA only. */
export function isFormBRJEarningsTotalHeader(header, headers = [], headerIndex = null) {
  if (isFormBRJDeductionTotalHeader(header, headers, headerIndex)) return false;
  return isFormBRJPlainTotalHeader(header);
}

export function isFormBRJNetPaymentHeader(h) {
  const s = normHeader(h);
  return s.includes('net') && (s.includes('payment') || s.includes('payable') || s.includes('paid'));
}

export function isFormBRJBankReceiptHeader(h) {
  const s = normHeader(h);
  return (
    (s.includes('receipt') && (s.includes('employee') || s.includes('workman') || s.includes('worker'))) ||
    (s.includes('bank') && (s.includes('transaction') || s.includes('account')))
  );
}

export function isFormBRJPaymentDateHeader(h) {
  const s = normHeader(h);
  return s.includes('date') && s.includes('payment');
}

/** Payroll / wage columns — skip generic People field mapping. Name, code, and bank receipt come from People. */
export function isFormBRJSkipPeopleAutofillHeader(h) {
  return (
    isFormBRJEmployeeCodeHeader(h) ||
    isFormBRJDaysWorkedHeader(h) ||
    isFormBRJOvertimeHoursHeader(h) ||
    isFormBRJPaymentsOvertimeHeader(h) ||
    isFormBRJBasicHeader(h) ||
    isFormBRJSpecialBasicHeader(h) ||
    isFormBRJDaHeader(h) ||
    isFormBRJHraHeader(h) ||
    isFormBRJOthersHeader(h) ||
    isFormBRJEarningsTotalHeader(h) ||
    isFormBRJDeductionTotalHeader(h) ||
    isFormBRJNetPaymentHeader(h) ||
    isFormBRJPaymentDateHeader(h)
  );
}

export function formBRJHeaderAliasBucket(norm) {
  const n = String(norm || '').trim();
  if (!n) return '';
  if (/^sr\.?\s*no|^s\.?\s*no|serial|sl\.?\s*no|^no\.?$/.test(n)) return 'sno';
  if (n === 'name' || /^name$/.test(n)) return 'name';
  if ((n.includes('day') || n.includes('days')) && n.includes('work')) return 'daysWorked';
  if (/\bover[\s-]*time\b/.test(n) && (n.includes('hour') || n.includes('hrs') || n.includes('worked'))) {
    return 'overtimeHours';
  }
  if (/\bover[\s-]*time\b/.test(n) && (n.includes('payment') || n.includes('pay'))) return 'paymentsOvertime';
  if (n === 'basic' || /^basic\b/.test(n)) return 'basic';
  if (n === 'hra' || /\bhra\b/.test(n) || (n.includes('house') && n.includes('rent'))) return 'hra';
  if (n === 'others' || /^others\b/.test(n)) return 'others';
  if (
    ((n.includes('deduction') || n.includes('deducation')) && n.includes('total')) ||
    n === 'deduction' ||
    n === 'deducation'
  ) {
    return 'deductionTotal';
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
  if (n === 'total' || /^total\b/.test(n)) return 'earningsTotal';
  if (n.includes('employee') && n.includes('code')) return 'employeeCode';
  return n;
}

export function headersIndicateFormBRajasthanTable(tableHeaders) {
  if (!Array.isArray(tableHeaders) || tableHeaders.length < 5) return false;
  const joined = tableHeaders.map((h) => formBRajasthanHeaderNorm(h)).join('\n');
  const hasName = /\bname\b/.test(joined);
  const hasDaysWorked = /day/.test(joined) && /work/.test(joined);
  const hasBasic = /\bbasic\b/.test(joined);
  const hasWageRegister =
    /\bhra\b/.test(joined) ||
    /net\s+payment/.test(joined) ||
    /payment\s+overtime/.test(joined) ||
    /deduction|deducation/.test(joined);
  const looksLikeFormAEmployeeRegister =
    /education/.test(joined) &&
    (/type/.test(joined) && /employ/.test(joined)) &&
    (/national/.test(joined) || /aadha?ar/.test(joined));
  return hasName && hasDaysWorked && hasBasic && hasWageRegister && !looksLikeFormAEmployeeRegister;
}

export function isFormBRajasthanContext(
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

  if (/gujarat|\b_gj\b|form[\s._-]*b[\s._-]*gj|form_b_gj/.test(parts)) return false;

  const hasRajasthan =
    /rajasthan|\b_rj\b|form[\s._-]*b[\s._-]*rj|form_b_rj/.test(parts);
  const hasFormB = /\bform[\s._-]*b\b/.test(parts);

  if (hasRajasthan && hasFormB) return true;
  if (headersIndicateFormBRajasthanTable(tableHeaders) && hasRajasthan && hasFormB) return true;
  if (headersIndicateFormBRajasthanTable(tableHeaders) && hasRajasthan) return true;
  return false;
}

export function resolveFormBRajasthanTableHeaders(tableHeaders) {
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

function formatBRJPayrollPayDate(payDateRaw) {
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

export function resolveFormBRajasthanPayrollFields(payrollRow, helpers = {}) {
  const monthEndDate = String(helpers.monthEndDate || '').trim();
  const empty = {
    paidDays: '',
    overtimeHours: FORM_B_RJ_NIL,
    paymentsOvertime: FORM_B_RJ_NIL,
    basic: '',
    hra: '',
    earningsTotal: '',
    deductionsTotal: '',
    netPay: '',
    paymentDate: monthEndDate,
  };
  if (!payrollRow || payrollRow.fetch_error) return empty;

  const flat = flattenPayrollEarningColumns(payrollRow);
  const wageAmounts = readPayrollForm15WageAmounts(payrollRow);

  const paidDays = readPayrollScalar(
    { ...flat, ...payrollRow },
    ['paid_days', 'Paid_days', 'Paid Days', 'days_worked', 'Days Worked', 'paidDays', 'no_of_days_worked'],
    [/^paid_days$/, /paiddays/, /daysworked/]
  );

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
    ['gross_pay', 'Gross Pay', 'Gross_pay', 'grossPay', 'total_earnings'],
    [/^gross_pay$/, /^total_earnings$/]
  );

  const netPay = readPayrollScalar(
    { ...flat, ...payrollRow },
    ['net_pay', 'Net Pay', 'netPay', 'monthly_salary'],
    [/^net_pay$/]
  );

  const grossN = parsePayrollNumber(grossPay);
  const netN = parsePayrollNumber(netPay);
  let deductionsTotal = '';
  if (Number.isFinite(grossN) && Number.isFinite(netN)) {
    deductionsTotal = Math.round((grossN - netN) * 100) / 100;
    if (deductionsTotal < 0) deductionsTotal = '';
  }

  const earningsTotal = sumPayrollNumbers([basic, hra]);

  // Date of Payment → month end (selected wage month), not payroll pay_date.
  const paymentDate =
    monthEndDate ||
    formatBRJPayrollPayDate(
      readPayrollTextScalar(
        { ...flat, ...payrollRow },
        ['pay_date', 'Pay Date', 'payment_date', 'Payment Date', 'paid_date', 'date_of_payment', 'Date of Payment'],
        [/^pay_date$/, /payment.*date/i, /^paid_date$/]
      ) || String(payrollRow?.pay_date ?? flat?.pay_date ?? helpers.payDate ?? '').trim()
    );

  return {
    paidDays,
    overtimeHours: FORM_B_RJ_NIL,
    paymentsOvertime: FORM_B_RJ_NIL,
    basic,
    hra,
    earningsTotal,
    deductionsTotal,
    netPay,
    paymentDate,
  };
}

function formatCellValue(value, { allowZero = false, allowNil = false } = {}) {
  if (value == null || value === '') return '';
  if (allowNil && String(value).trim().toLowerCase() === 'nil') return FORM_B_RJ_NIL;
  const num = parsePayrollNumber(value);
  if (Number.isFinite(num) && (allowZero || num !== 0)) return num;
  return String(value).trim();
}

export function enrichFormBRajasthanPayrollRows(mappedData, employees, headers, helpers = {}) {
  const hdrs = resolveFormBRajasthanTableHeaders(headers);
  const {
    resolvePayrollRow = null,
    payrollRows = null,
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
    const empItem = employees[rowIndex];
    const emp = unwrapEmployeeRecord(empItem);
    let payrollRow =
      typeof resolvePayrollRow === 'function' ? resolvePayrollRow(emp, row, rowIndex) : null;
    if ((!payrollRow || payrollRow.fetch_error) && Array.isArray(payrollRows) && payrollRows.length > 0) {
      payrollRow = resolvePayrollRowByFormTableName(row, headers, payrollRows, {
        isNameHeader: isFormBRJNameHeader,
      });
    }
    const merged = applyFormBRajasthanEmployeeToRow(row, emp, hdrs, {
      sanitizeValue,
      formatStatutoryDateDisplay,
      payDate,
      monthEndDate,
      payrollRow: payrollRow && !payrollRow.fetch_error ? payrollRow : null,
      rowIndex: rowIndexOffset + rowIndex,
      overwrite,
    });
    Object.assign(row, merged);
    hits += 1;
  });
  return hits;
}

export function applyFormBRajasthanEmployeeToRow(row, emp, headers, helpers = {}) {
  const hdrs = resolveFormBRajasthanTableHeaders(headers);
  const out = row && typeof row === 'object' ? { ...row } : {};
  const {
    sanitizeValue = (v) => String(v ?? '').trim(),
    payrollRow = null,
    rowIndex = 0,
    formatStatutoryDateDisplay = (v) => String(v || '').trim(),
    payDate = '',
    monthEndDate = '',
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

  const payroll = resolveFormBRajasthanPayrollFields(
    payrollRow && !payrollRow.fetch_error ? payrollRow : null,
    { formatStatutoryDateDisplay, payDate, monthEndDate }
  );
  const fullName = readEmployeeFullName(emp);
  if (fullName) out.__employeeLookupName = fullName;

  hdrs.forEach((header, headerIndex) => {
    if (isFormBRJSerialHeader(header)) {
      setCell(header, String(rowIndex + 1), { allowZero: true });
      return;
    }
    if (isFormBRJNameHeader(header)) {
      setCell(header, fullName);
      return;
    }
    if (isFormBRJEmployeeCodeHeader(header)) {
      setCell(header, readEmployeeId(emp));
      return;
    }
    if (isFormBRJDaysWorkedHeader(header)) {
      setCell(header, payroll.paidDays, { allowZero: true });
      return;
    }
    if (isFormBRJOvertimeHoursHeader(header)) {
      setCell(header, payroll.overtimeHours, { allowNil: true });
      return;
    }
    if (isFormBRJPaymentsOvertimeHeader(header)) {
      setCell(header, payroll.paymentsOvertime, { allowNil: true });
      return;
    }
    if (isFormBRJBasicHeader(header)) {
      setCell(header, payroll.basic);
      return;
    }
    if (isFormBRJHraHeader(header)) {
      setCell(header, payroll.hra);
      return;
    }
    if (isFormBRJOthersHeader(header)) {
      // Deduction → Others stays blank (do not fetch payroll residual).
      if (overwrite || cellIsEmpty(header)) out[header] = '';
      return;
    }
    if (isFormBRJDeductionTotalHeader(header, hdrs, headerIndex)) {
      setCell(header, payroll.deductionsTotal, { allowZero: true });
      return;
    }
    if (isFormBRJEarningsTotalHeader(header, hdrs, headerIndex)) {
      setCell(header, payroll.earningsTotal);
      return;
    }
    if (isFormBRJNetPaymentHeader(header)) {
      setCell(header, payroll.netPay);
      return;
    }
    if (isFormBRJBankReceiptHeader(header)) {
      setCell(header, readBankAccountNumber(emp));
      return;
    }
    if (isFormBRJPaymentDateHeader(header)) {
      setCell(header, payroll.paymentDate);
    }
  });

  return out;
}

export function getFormBRajasthanRowValueForHeader(row, header, rowIndex = 0) {
  if (!row || !header) return '';
  const direct = row[header];
  if (direct != null && String(direct).trim() !== '') return String(direct).trim();
  const bucket = formBRJHeaderAliasBucket(formBRajasthanHeaderNorm(header));
  for (const [k, v] of Object.entries(row)) {
    if (
      formBRJHeaderAliasBucket(formBRajasthanHeaderNorm(k)) === bucket &&
      v != null &&
      String(v).trim() !== ''
    ) {
      return String(v).trim();
    }
  }
  if (bucket === 'sno') return String(rowIndex + 1);
  return '';
}

export function rowHasMeaningfulFormBRajasthanExportData(row, headers) {
  const hdrs = resolveFormBRajasthanTableHeaders(headers);
  return hdrs.some((header) => {
    const bucket = formBRJHeaderAliasBucket(formBRajasthanHeaderNorm(header));
    if (bucket === 'sno') return false;
    return getFormBRajasthanRowValueForHeader(row, header) !== '';
  });
}

export function remapFormBRajasthanRowsToHeaders(rows, sourceHeaders, targetHeaders) {
  const src = resolveFormBRajasthanTableHeaders(sourceHeaders);
  const tgt = resolveFormBRajasthanTableHeaders(targetHeaders);
  if (!Array.isArray(rows)) return [];
  return rows.map((row, rowIndex) => {
    const out = {};
    tgt.forEach((targetHeader) => {
      let val = '';
      if (row && typeof row === 'object') {
        if (row[targetHeader] != null && String(row[targetHeader]).trim() !== '') {
          val = String(row[targetHeader]).trim();
        } else {
          val = getFormBRajasthanRowValueForHeader(row, targetHeader, rowIndex);
        }
      }
      if (!val && src.length > 0) {
        const bucket = formBRJHeaderAliasBucket(formBRajasthanHeaderNorm(targetHeader));
        for (const [k, v] of Object.entries(row || {})) {
          if (formBRJHeaderAliasBucket(formBRajasthanHeaderNorm(k)) === bucket) {
            val = String(v ?? '').trim();
            break;
          }
        }
      }
      if (
        !val &&
        formBRJHeaderAliasBucket(formBRajasthanHeaderNorm(targetHeader)) === 'sno'
      ) {
        val = String(rowIndex + 1);
      }
      out[targetHeader] = val;
    });
    return out;
  });
}

export function filterFormBRajasthanExportRows(rows, headers) {
  const hdrs = resolveFormBRajasthanTableHeaders(headers);
  return (Array.isArray(rows) ? rows : []).filter((row) =>
    rowHasMeaningfulFormBRajasthanExportData(row, hdrs)
  );
}

const excelCellValueToString = (val) => {
  if (val == null) return '';
  if (typeof val === 'string' || typeof val === 'number' || typeof val === 'boolean') {
    return String(val);
  }
  if (typeof val === 'object') {
    if (val instanceof Date) return val.toISOString().slice(0, 10);
    if (Array.isArray(val.richText)) return val.richText.map((rt) => rt?.text || '').join('');
    if (val.text != null) return String(val.text);
    if (val.result != null) return String(val.result);
  }
  return '';
};

function excelCellLooksLikeSerialHeader(text) {
  const t = formBRajasthanHeaderNorm(text).replace(/[^a-z0-9 ]/g, ' ').replace(/\s+/g, ' ').trim();
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

function detectFormBRajasthanTableLayout(worksheet, hints = {}) {
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
    const bucket = formBRJHeaderAliasBucket(formBRajasthanHeaderNorm(label));
    templateCols.push({ col: c, label, bucket });
    if (templateCols.length >= 30) break;
  }
  if (templateCols.length < 4) return null;

  // Disambiguate duplicate plain "Total" columns: last → Deduction Total.
  const plainTotalIdxs = templateCols
    .map((entry, i) => (entry.bucket === 'earningsTotal' ? i : -1))
    .filter((i) => i >= 0);
  if (plainTotalIdxs.length >= 2) {
    const lastIdx = plainTotalIdxs[plainTotalIdxs.length - 1];
    templateCols[lastIdx] = { ...templateCols[lastIdx], bucket: 'deductionTotal' };
  }

  const dataStartRow =
    hints.parsedDataStartIndex != null && hints.parsedDataStartIndex >= 0
      ? hints.parsedDataStartIndex + 1
      : headerRow + 1;

  return { headerRow, dataStartRow, templateCols, startCol };
}

export async function buildFormBRajasthanWorkbookWithTemplateStyles({
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

  const layout = detectFormBRajasthanTableLayout(worksheet, {
    parsedHeaderRowIndex,
    parsedDataStartIndex,
    parsedTableStartCol,
  });
  if (!layout) throw new Error('Could not locate Rajasthan Form B table header row.');

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
  const normalizedHeaders = resolveFormBRajasthanTableHeaders(
    Array.isArray(headersToUse) && headersToUse.length >= templateHeaderLabels.length
      ? headersToUse
      : templateHeaderLabels.length > 0
        ? templateHeaderLabels
        : headersToUse
  );
  const rows = filterFormBRajasthanExportRows(
    remapFormBRajasthanRowsToHeaders(
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
      let val = getFormBRajasthanRowValueForHeader(row, label, idx);
      if ((val == null || val === '') && bucket) {
        for (const [k, v] of Object.entries(row)) {
          if (
            formBRJHeaderAliasBucket(formBRajasthanHeaderNorm(k)) === bucket &&
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
      } else if (
        bucket === 'overtimeHours' ||
        bucket === 'paymentsOvertime'
      ) {
        cell.value = String(val);
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
    'Form_B_RJ_Rajasthan.xlsx';
  return {
    blob: new Blob([out], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    }),
    fileName,
  };
}

export { resolveFormXIXMPPayrollRowForEmployee as resolveFormBRajasthanPayrollRowForEmployee };
