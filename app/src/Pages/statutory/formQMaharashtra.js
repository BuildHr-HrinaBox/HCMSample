/** Maharashtra Form Q — muster-roll cum wage register payroll mapping. */

import {
  flattenPayrollEarningColumns,
  readPayrollScalar,
} from '../../utils/payrollEarnings';
import { buildFormPGJEmployeePayrollCandidates } from './formPGJGujarat';

function normHeaderLabel(h) {
  return String(h || '')
    .replace(/^\(?\d+\)?\s*[\.\)]?\s*/i, '')
    .replace(/\r?\n/g, ' ')
    .replace(/[^a-z0-9 ]/gi, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

function parsePayrollNumber(value) {
  const n = Number(String(value ?? '').replace(/[,₹]/g, '').trim());
  return Number.isFinite(n) ? n : NaN;
}

function formatCellValue(value, { allowZero = false } = {}) {
  if (value == null || value === '') return '';
  if (typeof value === 'string' && /^nil$/i.test(value.trim())) return 'Nil';
  const num = parsePayrollNumber(value);
  if (Number.isFinite(num) && (allowZero || num !== 0)) return num;
  return String(value).trim();
}

export function matchFormQMaharashtraWageTailBucket(header) {
  const n = normHeaderLabel(header);
  if (!n) return '';
  if (/total\s+days?\s+worked/.test(n)) return 'paidDays';
  if (/minimum\s+rate\s+of\s+wages/.test(n)) return 'minimumRateWages';
  if (/total\s+production/.test(n) && /piece\s+rate/.test(n)) return 'skip';
  if (/actual\s+wages\s+paid/.test(n)) return 'basic';
  if (/house\s+rent\s+allowance/.test(n)) return 'hra';
  if (/dearness\s+allowance/.test(n)) return 'skip';
  if (/gross\s+amount\s+payable/.test(n)) return 'grossPay';
  if (/total\s+hours?\s+of\s+overtime/.test(n)) return 'overtimeHours';
  if (/overtime\s+earnings?/.test(n)) return 'overtimeEarnings';
  if (/provident\s+fund/.test(n)) return 'skip';
  if (/family\s+pension/.test(n)) return 'skip';
  if (/esi\s+contribution/.test(n)) return 'skip';
  if (/professional\s+tax/.test(n)) return 'skip';
  if (/income\s+tax/.test(n)) return 'skip';
  if (/loan.*interest/.test(n)) return 'skip';
  if (/^advances$/.test(n)) return 'skip';
  if (/other\s+deductions?/.test(n)) return 'skip';
  if (/total\s+deductions?/.test(n)) return 'totalDeduction';
  if (/net\s+payable/.test(n)) return 'netPay';
  if (/date\s+of\s+payment/.test(n)) return 'paymentDate';
  if (/signature|thumb\s+impression/.test(n)) return 'skip';
  return '';
}

export function resolveFormQMaharashtraPayrollFields(payrollRow, helpers = {}) {
  const monthEndDate = String(helpers.monthEndDate || '').trim();
  const empty = {
    paidDays: '',
    minimumRateWages: '',
    basic: '',
    hra: '',
    grossPay: '',
    overtimeHours: 'Nil',
    overtimeEarnings: 'Nil',
    totalDeduction: '',
    netPay: '',
    paymentDate: monthEndDate,
  };
  if (!payrollRow || payrollRow.fetch_error) return empty;

  const flat = flattenPayrollEarningColumns(payrollRow);
  const merged = { ...flat, ...payrollRow };

  const paidDays = readPayrollScalar(
    merged,
    ['paid_days', 'Paid_days', 'Paid Days', 'paidDays', 'days_worked', 'Days Worked', 'no_of_days_worked'],
    [/^paid_days$/, /^paiddays$/, /daysworked/]
  );

  const basic =
    readPayrollScalar(
      merged,
      ['basic', 'Basic', 'earned_basic', 'Earned Basic', 'basic_pay', 'Basic Earnings'],
      [/^basic$/, /^earned_basic$/, /^basic_pay$/]
    ) || (flat.basic != null && flat.basic !== '' ? flat.basic : '');

  const hra =
    readPayrollScalar(
      merged,
      ['hra_fbp', 'HRA FBP', 'hra (fbp)', 'hra_fp', 'HRA (FBP)'],
      [/^hra_fbp$/, /^hra_fp$/]
    ) ||
    readPayrollScalar(
      merged,
      ['hra', 'HRA', 'house_rent_allowance', 'House Rent Allowance'],
      [/^hra$/, /house.*rent/]
    ) ||
    (flat.hra_fbp != null && flat.hra_fbp !== '' ? flat.hra_fbp : flat.hra || '');

  const grossPay = readPayrollScalar(
    merged,
    ['gross_pay', 'Gross Pay', 'grossPay', 'total_earnings', 'Gross Amount Payable'],
    [/^gross_pay$/, /^total_earnings$/]
  );

  const netPay = readPayrollScalar(
    merged,
    ['net_pay', 'Net Pay', 'netPay', 'monthly_salary', 'Net Payable'],
    [/^net_pay$/]
  );

  const grossNum = parsePayrollNumber(grossPay);
  const netNum = parsePayrollNumber(netPay);

  let minimumRateWages = '';
  if (Number.isFinite(grossNum) && grossNum > 0) {
    minimumRateWages = Math.round((grossNum / 26) * 100) / 100;
  }

  let totalDeduction = '';
  if (Number.isFinite(grossNum) && Number.isFinite(netNum)) {
    totalDeduction = Math.round((grossNum - netNum) * 100) / 100;
    if (totalDeduction < 0) totalDeduction = '';
  }

  return {
    paidDays,
    minimumRateWages,
    basic,
    hra,
    grossPay,
    overtimeHours: 'Nil',
    overtimeEarnings: 'Nil',
    totalDeduction,
    netPay,
    paymentDate: monthEndDate,
  };
}

export function applyFormQMaharashtraPayrollToRow(row, payrollRow, headers, helpers = {}) {
  const hdrs = Array.isArray(headers) ? headers : [];
  const out = row && typeof row === 'object' ? row : {};
  const {
    sanitizeValue = (v) => String(v ?? '').trim(),
    monthEndDate = '',
    overwrite = true,
  } = helpers;
  const payroll = resolveFormQMaharashtraPayrollFields(
    payrollRow && !payrollRow.fetch_error ? payrollRow : null,
    { monthEndDate }
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
  hdrs.forEach((header) => {
    const bucket = matchFormQMaharashtraWageTailBucket(header);
    if (!bucket || bucket === 'skip') return;
    if (bucket === 'overtimeHours' || bucket === 'overtimeEarnings') {
      if (setCell(header, payroll[bucket], { allowZero: true })) wrote = true;
      return;
    }
    if (bucket === 'paymentDate') {
      if (setCell(header, payroll.paymentDate, { allowZero: true })) wrote = true;
      return;
    }
    if (Object.prototype.hasOwnProperty.call(payroll, bucket)) {
      if (setCell(header, payroll[bucket], { allowZero: true })) wrote = true;
    }
  });
  return wrote;
}

export function enrichFormQMaharashtraPayrollRows(mappedData, employees, headers, helpers = {}) {
  const hdrs = Array.isArray(headers) && headers.length > 0 ? headers : [];
  const {
    resolvePayrollRow = null,
    sanitizeValue = (v) => String(v ?? '').trim(),
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
      payrollRow = resolvePayrollRow(empItem, row, globalRowIndex);
      if (!payrollRow || payrollRow.fetch_error) {
        const candidates = buildFormPGJEmployeePayrollCandidates(empItem, row);
        for (let ci = 0; ci < candidates.length; ci += 1) {
          const hit = resolvePayrollRow(candidates[ci], row, globalRowIndex);
          if (hit && !hit.fetch_error) {
            payrollRow = hit;
            break;
          }
        }
      }
    }
    const wrote = applyFormQMaharashtraPayrollToRow(row, payrollRow, hdrs, {
      sanitizeValue,
      monthEndDate,
      overwrite,
    });
    if (wrote) hits += 1;
  });
  return hits;
}
