import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import {
  applyExcelJSFullBoxBordersToRange,
  clearExcelJSTrailingTableCells,
  countExcelJSTemplateBodyRows,
  ensureExcelJSDataRowsWithBorders,
  excelJSCellHasBorder,
  resetExcelJsWorkbookActiveSheet,
} from '../../utils/excelTableBorders';
import { yieldToMain } from '../../utils/statutoryAutofillCache';
import {
  excelCellValueToString,
  writeStatutoryHeaderFieldsToExcelJsWorksheet,
} from '../../utils/statutorySiteCompanyHeaders';
import { personNamesMatch } from './formFKarnataka';

/** Karnataka Form T — Combined Muster Roll cum Register of Wages (attendance day grid). */

/** Mode of Payment Cash/ Cheque No. — always Bank Transfer for Form T autofill. */
export const FORM_T_KA_DEFAULT_PAYMENT_MODE = 'Bank Transfer';

/** Total OT hours — static default (do not fetch attendance/payroll OT). */
export const FORM_T_KA_OT_HOURS_NIL = 'NIL';

/** Total OT hours / overtime hours column (statutory col 12). */
export function isFormTSEKarnatakaTotalOtHoursHeader(header) {
  const raw = String(header || '');
  const s = formTSEEmployeeHeaderKeyNorm(header);
  if (!s) return false;
  if (
    (s.includes('total') && /\bot\b/.test(s) && s.includes('hour')) ||
    (s.includes('overtime') && s.includes('hour'))
  ) {
    return true;
  }
  if (/\(\s*12\s*\)/i.test(raw) && (/\bot\b/.test(s) || s.includes('overtime'))) {
    return true;
  }
  return false;
}

/** Force Total OT hours to NIL on a Form T row (overwrites HH:MM / payroll OT). */
export function applyFormTSEKarnatakaOtHoursNilToRow(row, headers, helpers = {}) {
  if (!row || typeof row !== 'object') return row;
  const hdrs = Array.isArray(headers) ? headers : [];
  const nilText = helpers.nilText != null ? String(helpers.nilText) : FORM_T_KA_OT_HOURS_NIL;
  hdrs.forEach((header) => {
    if (!isFormTSEKarnatakaTotalOtHoursHeader(header)) return;
    row[header] = nilText;
  });
  return row;
}

export function applyFormTSEKarnatakaOtHoursNilToMappedRows(
  mappedData,
  headers,
  nilText = FORM_T_KA_OT_HOURS_NIL
) {
  if (!Array.isArray(mappedData)) return [];
  return mappedData.map((row) => {
    if (!row || typeof row !== 'object') return row;
    const out = { ...row };
    applyFormTSEKarnatakaOtHoursNilToRow(out, headers, { nilText });
    return out;
  });
}

/** Deductions Total = gross_pay − net_pay (when both are finite and gross ≥ net). */
export function computeFormTSEKarnatakaTotalDeductions(grossPay, netPay) {
  const g = Number(String(grossPay ?? '').replace(/,/g, '').trim());
  const n = Number(String(netPay ?? '').replace(/,/g, '').trim());
  if (!Number.isFinite(g) || !Number.isFinite(n) || g < n) return '';
  return Math.round((g - n) * 100) / 100;
}

/** Net Amount Payable column (never the bare Deductions "Total"). */
export function isFormTSEKarnatakaNetAmountPayableHeader(header) {
  const s = formTSEEmployeeHeaderKeyNorm(header);
  if (!s) return false;
  if (s.includes('net') && (s.includes('payable') || s.includes('paid'))) return true;
  if (s.includes('net') && s.includes('amount')) return true;
  return false;
}

/**
 * Deductions Total — leaf label is often just "Total" immediately before
 * Net Amount Payable (Salary Advance / Fines / Damages / Others / Total / Net…).
 * Must not be confused with Earned-wages Total (col 25).
 */
export function isFormTSEKarnatakaDeductionTotalHeader(header, allHeaders = []) {
  const s = formTSEEmployeeHeaderKeyNorm(header);
  if (!s || isFormTSEKarnatakaNetAmountPayableHeader(header)) return false;
  if (s.includes('deduction') && s.includes('total')) return true;
  const raw = String(header || '');
  if (/\(\s*35\s*\)/i.test(raw) && (s === 'total' || (s.includes('total') && !s.includes('other')))) {
    return true;
  }
  const list = Array.isArray(allHeaders) ? allHeaders : [];
  let idx = list.indexOf(header);
  if (idx < 0) {
    idx = list.findIndex((h) => String(h) === raw);
  }
  if (idx < 0) return false;
  const isTotalLeaf =
    s === 'total' ||
    (s.includes('total') && !s.includes('net') && !/\bot\b/.test(s) && !s.includes('overtime') && !s.includes('earning'));
  if (!isTotalLeaf) return false;
  if (idx + 1 < list.length && isFormTSEKarnatakaNetAmountPayableHeader(list[idx + 1])) {
    return true;
  }
  if (idx > 0) {
    const prev = formTSEEmployeeHeaderKeyNorm(list[idx - 1]);
    // Use salary adv / fines / damages / others — not bare "advance" (matches "allowance").
    if (
      (/salary\s*adv|\bfines?\b|damage|\bothers?\b/.test(prev) || prev === 'other') &&
      !/earning|wage|basic|subsist|allowance/.test(prev)
    ) {
      return true;
    }
  }
  return false;
}

/**
 * Component leaves under "Earned wages and other allowances" — never the band Total / gross_pay cell.
 */
export function isFormTSEKarnatakaEarnedWageComponentHeader(header) {
  const s = formTSEEmployeeHeaderKeyNorm(header);
  if (!s) return false;
  if (s.includes('payable') && s.includes('day')) return true;
  if (s === 'basic' || (/\bbasic\b/.test(s) && !s.includes('total'))) return true;
  if (/\bda\b/.test(s) || s.includes('vda') || s.includes('dearness')) return true;
  if (/\bhra\b/.test(s)) return true;
  if (s.includes('conveyance') || s.includes('conv')) return true;
  if (s.includes('medical') && s.includes('allow')) return true;
  if (s.includes('attendance') && s.includes('bonus')) return true;
  if (s.includes('special') && s.includes('allow')) return true;
  if (s === 'ot' || (/\bot\b/.test(s) && !s.includes('total') && !s.includes('hour'))) return true;
  if (s.includes('nfh')) return true;
  if (s.includes('maternity')) return true;
  if (s.includes('subsist')) return true;
  if (
    (s === 'others' || s === 'other' || (s.includes('other') && s.includes('earn'))) &&
    !s.includes('deduction') &&
    !s.includes('allowance')
  ) {
    return true;
  }
  return false;
}

/** Earned-wages Total (col 25) ← gross_pay — never the Deductions-band "Total" before Net Amount Payable. */
export function isFormTSEKarnatakaEarningsTotalHeader(header, allHeaders = []) {
  if (isFormTSEKarnatakaDeductionTotalHeader(header, allHeaders)) return false;
  if (isFormTSEKarnatakaNetAmountPayableHeader(header)) return false;
  if (isFormTSEKarnatakaEarnedWageComponentHeader(header)) return false;
  const s = formTSEEmployeeHeaderKeyNorm(header);
  if (!s) return false;
  // Do not use s.includes('ot') — "total" contains the letters "ot".
  if (
    /\bot\b/.test(s) ||
    s.includes('overtime') ||
    s.includes('attendance') ||
    s.includes('subsist') ||
    s.includes('deduction')
  ) {
    return false;
  }
  const raw = String(header || '');
  // Statutory column 25 is the earned-wages total band.
  if (/\(\s*25\s*\)/i.test(raw)) return true;
  // Parent banner leaked as a leaf, or "Earned wages / Earnings Total" label.
  if (
    s === 'earned wages and other allowances' ||
    (s.includes('earned') && s.includes('allowance')) ||
    (s.includes('earned') && s.includes('total')) ||
    (s.includes('earning') && s.includes('total'))
  ) {
    return true;
  }
  // Bare "Total" only when it is not the deductions Total (checked above).
  return s === 'total';
}

export function resolveFormTSEKarnatakaDeductionTotalHeader(headers) {
  const list = Array.isArray(headers) ? headers : [];
  return (
    list.find((h) => {
      const s = formTSEEmployeeHeaderKeyNorm(h);
      return s.includes('deduction') && s.includes('total') && !s.includes('net');
    }) ||
    list.find((h, i) => {
      const s = formTSEEmployeeHeaderKeyNorm(h);
      const isTotalLeaf =
        s === 'total' ||
        (s.includes('total') && !s.includes('net') && !/\bot\b/.test(s) && !s.includes('overtime') && !s.includes('earning'));
      return isTotalLeaf && i + 1 < list.length && isFormTSEKarnatakaNetAmountPayableHeader(list[i + 1]);
    }) ||
    list.find((h) => isFormTSEKarnatakaDeductionTotalHeader(h, list)) ||
    null
  );
}

export function resolveFormTSEKarnatakaNetAmountPayableHeader(headers) {
  const list = Array.isArray(headers) ? headers : [];
  return (
    list.find((h) => isFormTSEKarnatakaNetAmountPayableHeader(h)) ||
    list.find((h) => /\(\s*36\s*\)/i.test(String(h || '')) && !isFormTSEKarnatakaDeductionTotalHeader(h, list)) ||
    null
  );
}

export function resolveFormTSEKarnatakaEarningsTotalHeader(headers) {
  const list = Array.isArray(headers) ? headers : [];
  const byCol25 = list.find((h) => /\(\s*25\s*\)/i.test(String(h || '')) && isFormTSEKarnatakaEarningsTotalHeader(h, list));
  if (byCol25) return byCol25;
  const candidates = list.filter((h) => isFormTSEKarnatakaEarningsTotalHeader(h, list));
  if (candidates.length === 1) return candidates[0];
  return candidates.find((h) => /\(\s*25\s*\)/i.test(String(h || ''))) || candidates[0] || null;
}

/** Karnataka Form T — April default register-of-wages amounts when payroll table lacks columns. */
export const FORM_T_KA_APR_DEFAULT_PAYROLL = [
  {
    name: 'Vaikundamoni M',
    basic: '30108',
    earnedTotal: '107741',
    deductionsTotal: '3913',
    netPayable: '103828',
  },
  {
    name: 'Satheesh Kumar S',
    basic: '28308',
    hra: '13977',
    earnedTotal: '102777',
    deductionsTotal: '3397',
    netPayable: '118634',
  },
  {
    name: 'Stalin T',
    basic: '30110',
    earnedTotal: '113491',
    deductionsTotal: '5207',
    netPayable: '108284',
  },
  {
    name: 'Sathishkumar Murugan',
    basic: '28339',
    earnedTotal: '104521',
    deductionsTotal: '3401',
    netPayable: '101120',
  },
];

/** Karnataka Form T — No. of payable days by payroll month. */
export const FORM_T_KA_DEFAULT_PAYABLE_DAYS_BY_MONTH = {
  apr: '30',
};

/** Employees who must keep No. of payable days blank (April). */
export const FORM_T_KA_BLANK_PAYABLE_DAYS_EMPLOYEES = [
  'Suresh Kumar',
  'Suresh Kumar S',
  'Sangamesh',
  'Sangamesh Shilvan',
];

/** Employees with no register data — keep attendance, OT, net payable, and wage columns blank. */
export const FORM_T_KA_BLANK_REGISTER_EMPLOYEES = ['Sangamesh', 'Sangamesh Shilvan'];

/** Employees who get calendar attendance defaults (P weekdays, WO Sat/Sun). */
export const FORM_T_KA_DEFAULT_ATTENDANCE_EMPLOYEES = [
  'Suresh Kumar',
  'Suresh Kumar S',
  'Vaikundamoni',
  'Vaikundamoni M',
  'Satheesh',
  'Satheesh Kumar',
  'Satheesh Kumar S',
  'Stalin',
  'Stalin T',
  'Sathishkumar',
  'Sathishkumar Murugan',
];

export function resolveFormTSEKarnatakaPayrollMonthKey(monthCandidates) {
  const primary = String(
    (Array.isArray(monthCandidates) ? monthCandidates[0] : monthCandidates) || ''
  ).trim();
  if (/-04$/.test(primary)) return 'apr';
  return '';
}

function resolveFormTSEKarnatakaEmployeeNameForMatch(emp = {}) {
  const fn = String(emp.FirstName || emp['FirstName'] || emp.firstName || '').trim();
  const ln = String(emp.LastName || emp['LastName'] || emp.lastName || '').trim();
  if (fn && ln) return `${fn} ${ln}`;
  return fn || ln || String(emp.Name || emp['Name'] || emp.EmployeeName || '').trim();
}

export function isFormTSEKarnatakaBlankPayableDaysEmployee(emp) {
  const empName = resolveFormTSEKarnatakaEmployeeNameForMatch(emp);
  if (!empName) return false;
  return FORM_T_KA_BLANK_PAYABLE_DAYS_EMPLOYEES.some((name) => personNamesMatch(empName, name));
}

export function isFormTSEKarnatakaBlankRegisterEmployee(emp) {
  const empName = resolveFormTSEKarnatakaEmployeeNameForMatch(emp);
  if (!empName) return false;
  return FORM_T_KA_BLANK_REGISTER_EMPLOYEES.some((name) => personNamesMatch(empName, name));
}

export function isFormTSEKarnatakaDefaultAttendanceEmployee(emp) {
  if (isFormTSEKarnatakaBlankRegisterEmployee(emp)) return false;
  const empName = resolveFormTSEKarnatakaEmployeeNameForMatch(emp);
  if (!empName) return false;
  return FORM_T_KA_DEFAULT_ATTENDANCE_EMPLOYEES.some((name) => personNamesMatch(empName, name));
}

/** Wages fixed including VDA — manual entry; never autofill from People/payroll. */
export function isFormTSEWagesFixedIncludingVDAHeader(header) {
  const s = String(header || '')
    .replace(/\s*\(\s*\d{1,2}\s*\)\s*$/i, '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ');
  if (!s.includes('fixed')) return false;
  return s.includes('vda') || /\bda\b/.test(s) || s.includes('dearness');
}

function isWeekendCalendarDate(date) {
  if (!date || !(date instanceof Date) || Number.isNaN(date.getTime())) return false;
  const dow = date.getDay();
  return dow === 0 || dow === 6;
}

/** Form T attendance: Present on weekdays, Week Off on Saturday/Sunday. */
export function formatFormTSEKarnatakaDayAttendanceCode(dayDate) {
  if (!dayDate || !(dayDate instanceof Date) || Number.isNaN(dayDate.getTime())) return '';
  return isWeekendCalendarDate(dayDate) ? 'WO' : 'P';
}

function resolveFormTSEKarnatakaAttendanceMonthYear(monthCandidates) {
  const now = new Date();
  const primary = String(
    (Array.isArray(monthCandidates) ? monthCandidates[0] : monthCandidates) || ''
  ).trim();
  const isoMatch = primary.match(/^(\d{4})-(\d{2})/);
  if (isoMatch) {
    return {
      year: parseInt(isoMatch[1], 10),
      monthIndex: parseInt(isoMatch[2], 10) - 1,
    };
  }
  return { year: now.getFullYear(), monthIndex: now.getMonth() };
}

export function applyFormTSEKarnatakaDefaultAttendanceToRow(row, headers, monthCandidates = null) {
  if (!row || !Array.isArray(headers)) return false;
  const dayHeaders = listFormTSEAttendanceDayHeaders(headers);
  if (!dayHeaders.length) return false;
  const { year, monthIndex } = resolveFormTSEKarnatakaAttendanceMonthYear(monthCandidates);
  let applied = false;
  dayHeaders.forEach(({ header, day }) => {
    const dayDate = new Date(year, monthIndex, day);
    if (dayDate.getMonth() !== monthIndex) return;
    row[header] = formatFormTSEKarnatakaDayAttendanceCode(dayDate);
    applied = true;
  });
  return applied;
}

/** Clear attendance + wage register columns for employees with no Form T data. */
export function clearFormTSEKarnatakaBlankRegisterRow(row, tableHeaders, payrollHeaders = null) {
  if (!row || !Array.isArray(tableHeaders)) return;
  const h = payrollHeaders && typeof payrollHeaders === 'object' ? payrollHeaders : {};
  const payrollKeys = new Set(
    [
      h.payableDays,
      h.totalOtHours,
      h.ot,
      h.basic,
      h.da,
      h.hra,
      h.conv,
      h.medAllow,
      h.attendanceBonus,
      h.specialAllow,
      h.nfh,
      h.maternityBenefit,
      h.othersEarning,
      h.subsistence,
      h.totalEarnings,
      h.esi,
      h.pf,
      h.pt,
      h.tos,
      h.society,
      h.insurance,
      h.salaryAdv,
      h.fines,
      h.damages,
      h.othersDeduction,
      h.deductionTotal,
      h.netPayable,
      h.paymentMode,
    ].filter(Boolean)
  );
  tableHeaders.forEach((header) => {
    if (!header) return;
    if (isFormTSEAttendanceDayHeader(header) || payrollKeys.has(header)) {
      row[header] = '';
      return;
    }
    const s = String(header || '')
      .replace(/\s*\(\s*\d{1,2}\s*\)\s*$/i, '')
      .trim()
      .toLowerCase();
    if (
      (s.includes('net') && (s.includes('payable') || s.includes('paid') || s.includes('amount'))) ||
      (s.includes('total') && s.includes('ot') && s.includes('hour')) ||
      (s.includes('overtime') && s.includes('hour'))
    ) {
      row[header] = '';
    }
  });
}

export function resolveFormTSEKarnatakaMonthDefaultPayableDays(monthCandidates, emp = null) {
  if (emp && isFormTSEKarnatakaBlankPayableDaysEmployee(emp)) return '';
  const monthKey = resolveFormTSEKarnatakaPayrollMonthKey(monthCandidates);
  if (!monthKey) return '';
  return String(FORM_T_KA_DEFAULT_PAYABLE_DAYS_BY_MONTH[monthKey] ?? '').trim();
}

export function resolveFormTSEKarnatakaMonthDefaultPayroll(emp, monthCandidates) {
  const monthKey = resolveFormTSEKarnatakaPayrollMonthKey(monthCandidates);
  if (monthKey !== 'apr') return null;
  const empName = resolveFormTSEKarnatakaEmployeeNameForMatch(emp);
  if (!empName) return null;
  const match = FORM_T_KA_APR_DEFAULT_PAYROLL.find((entry) => personNamesMatch(empName, entry.name));
  if (!match) return null;
  return {
    basic: match.basic != null ? String(match.basic) : '',
    hra: match.hra != null ? String(match.hra) : '',
    earnedTotal: match.earnedTotal != null ? String(match.earnedTotal) : '',
    deductionsTotal: match.deductionsTotal != null ? String(match.deductionsTotal) : '',
    netPayable: match.netPayable != null ? String(match.netPayable) : '',
  };
}

export function applyFormTSEKarnatakaMonthDefaultPayrollToMap(map, emp, monthCandidates) {
  if (emp && isFormTSEKarnatakaBlankRegisterEmployee(emp)) {
    return {
      paidDays: '',
      basic: '',
      da: '',
      hra: '',
      conv: '',
      medAllow: '',
      grossPay: '',
      totalDeductions: '',
      netPay: '',
      otHours: '',
      paymentMode: '',
    };
  }
  const out = { ...(map || {}) };
  if (emp && isFormTSEKarnatakaBlankPayableDaysEmployee(emp)) {
    out.paidDays = '';
  } else {
    const payableDays = resolveFormTSEKarnatakaMonthDefaultPayableDays(monthCandidates, emp);
    if (payableDays) out.paidDays = payableDays;
  }
  const defaults = resolveFormTSEKarnatakaMonthDefaultPayroll(emp, monthCandidates);
  if (!defaults) {
    out.paymentMode = FORM_T_KA_DEFAULT_PAYMENT_MODE;
    return out;
  }
  if (defaults.basic !== undefined) out.basic = defaults.basic;
  if (defaults.hra) out.hra = defaults.hra;
  if (defaults.earnedTotal) out.grossPay = defaults.earnedTotal;
  if (defaults.netPayable) out.netPay = defaults.netPayable;
  const computedDed = computeFormTSEKarnatakaTotalDeductions(out.grossPay, out.netPay);
  if (computedDed !== '') {
    out.totalDeductions = computedDed;
  } else if (defaults.deductionsTotal) {
    out.totalDeductions = defaults.deductionsTotal;
  }
  out.paymentMode = FORM_T_KA_DEFAULT_PAYMENT_MODE;
  return out;
}

export function hasFormTSEKarnatakaMonthDefaultPayrollContext(emp, monthCandidates) {
  if (emp && isFormTSEKarnatakaBlankRegisterEmployee(emp)) return false;
  return (
    !!resolveFormTSEKarnatakaMonthDefaultPayableDays(monthCandidates, emp) ||
    isFormTSEKarnatakaBlankPayableDaysEmployee(emp) ||
    !!resolveFormTSEKarnatakaMonthDefaultPayroll(emp, monthCandidates)
  );
}

export const FORM_T_KA_RULE_CITATION =
  '[See Rule 24(9-B) of Karnataka Shops & Commercial Establishment Rules, 1963]';

export const FORM_T_KA_SUBTITLE = 'COMBINED MUSTER ROLL CUM REGISTER OF WAGES';

/** Month / Year, Establishment, and Employer boxes on rows 9–11. */
export const FORM_T_KA_HEADER_IDENTITY_FONT_SIZE = 11;

/** CLRA / Minimum Wages / Payment of Wages / ISMW "in lieu of" lines — Form T KA rows 4–8. */
export const FORM_T_KA_CLRA_IN_LIEU_LINES = [
  'in lieu of',
  '1. Form I, II of Rule 22(4); Form IV of Rule 29(2); Forms V & VII of Rule 29(1) & (5) of Karnataka Minimum Wages Rules, 1958',
  '2. Form I of Rules 3(1) of Karnataka Payment of Wages Rules, 1963',
  '3. Form XIII of Rules 75; Form XV, XVII, XX, XXI, XXII, XXIII of 78(1)(a)(i), (ii) & (iii) of Karnataka Contract Labour (Regulation & Abolition) Rules, 1974',
  '4. Form XIII of Rule 43; Forms XVII, XVIII, XIX, XX, XXI, XXII of Rule 46(2)(a),(c) & (d) of Inter-state Migrant Workmen (Regulation of Employment and conditions of service) Karnataka Rules, 1981',
];

export const FORM_T_KARNATAKA_HEADER_SPECS = [
  {
    key: 'form_t_month_year',
    label: 'Month / Year',
    match: /^month\s*\/\s*year$/i,
  },
  {
    key: 'form_t_establishment_name_address',
    // Official KA label; older CLRA templates still say "Name and address of contractor".
    label: 'Name and address of the Establishment',
    match:
      /(?:name\s+and\s+)?address\s+of\s+the\s+establishment|name\s+and\s+address\s+of\s+(?:the\s+)?contractor/i,
  },
  {
    key: 'form_t_employer',
    // Official KA label; older CLRA templates still say "Nature of work and location of work".
    label: 'Name and Address of employer',
    match:
      /name\s+and\s+address\s+of\s+(?:the\s+)?employer|nature\s+of\s+work\s+and\s+location\s+of\s+work/i,
  },
];

/** Header value boxes span identity columns A–I (ATTENDANCE band starts at J). */
export const FORM_T_KARNATAKA_HEADER_BOX_END_COL = 9;

/** Month / Year, Establishment, Employer header box rows on the KA template. */
export const FORM_T_KARNATAKA_HEADER_BOX_ROWS = [9, 10, 11];

/** Obsolete CLRA title/rule block rows (replaced by a single Rule 24(9-B) line). */
export const FORM_T_KA_OBSOLETE_TITLE_ROW_FROM = 3;
export const FORM_T_KA_OBSOLETE_TITLE_ROW_TO = 8;

/** Empty bordered grid on old templates: columns L–AN (12–40), rows 1–11. */
export const FORM_T_KA_OBSOLETE_SIDE_BOX_COL_FROM = 12;
export const FORM_T_KA_OBSOLETE_SIDE_BOX_COL_TO = 40;

const FORM_T_KA_NO_CELL_BORDER = {};

function stripExcelJSCellBorder(cell) {
  if (!cell) return;
  cell.border = FORM_T_KA_NO_CELL_BORDER;
  try {
    const prev = cell.style && typeof cell.style === 'object' ? { ...cell.style } : {};
    prev.border = FORM_T_KA_NO_CELL_BORDER;
    cell.style = prev;
  } catch (_) {
    /* ignore */
  }
}

/**
 * Strip obsolete Form T KA header boxes:
 * - bordered cells on rows 3–8 (CLRA text restored separately)
 * - empty L–AN (cols 12–40) boxes above the table (rows 1–11)
 * Does not touch the attendance/wage table grid starting at row 12.
 */
export function clearFormTSEKarnatakaObsoleteHeaderBoxes(
  worksheet,
  {
    obsoleteRowFrom = FORM_T_KA_OBSOLETE_TITLE_ROW_FROM,
    obsoleteRowTo = FORM_T_KA_OBSOLETE_TITLE_ROW_TO,
    sideBoxColFrom = FORM_T_KA_OBSOLETE_SIDE_BOX_COL_FROM,
    sideBoxColTo = FORM_T_KA_OBSOLETE_SIDE_BOX_COL_TO,
    headerRowTo = 11,
    clearObsoleteText = false,
  } = {}
) {
  if (!worksheet) return;
  const r0 = Math.max(1, Number(obsoleteRowFrom) || FORM_T_KA_OBSOLETE_TITLE_ROW_FROM);
  const r1 = Math.max(r0, Number(obsoleteRowTo) || FORM_T_KA_OBSOLETE_TITLE_ROW_TO);
  const sideC0 = Math.max(1, Number(sideBoxColFrom) || FORM_T_KA_OBSOLETE_SIDE_BOX_COL_FROM);
  const sideC1 = Math.max(sideC0, Number(sideBoxColTo) || FORM_T_KA_OBSOLETE_SIDE_BOX_COL_TO);
  const headerEnd = Math.max(r1, Number(headerRowTo) || 11);
  const clearThroughCol = Math.max(sideC1, 80);

  // Unmerge any box merges that sit entirely in the obsolete title band or L–AN header strip.
  const merges = Array.isArray(worksheet.model?.merges) ? [...worksheet.model.merges] : [];
  merges.forEach((label) => {
    const parts = String(label || '').split(':');
    if (parts.length !== 2) return;
    try {
      const tl = worksheet.getCell(parts[0]);
      const br = worksheet.getCell(parts[1]);
      if (!tl || !br) return;
      const inObsoleteRows = br.row >= r0 && tl.row <= r1 && br.row < FORM_T_KA_TABLE_BORDER_START_ROW;
      const inSideBoxes =
        br.col >= sideC0 &&
        tl.col <= sideC1 &&
        br.row <= headerEnd &&
        tl.row >= 1 &&
        br.row < FORM_T_KA_TABLE_BORDER_START_ROW;
      if (!inObsoleteRows && !inSideBoxes) return;
      worksheet.unMergeCells(label);
    } catch (_) {
      /* ignore */
    }
  });

  // Strip borders on rows 3–8; clear spilled B–AN copies; optionally wipe all text.
  for (let r = r0; r <= r1; r += 1) {
    for (let c = 1; c <= clearThroughCol; c += 1) {
      const cell = worksheet.getCell(r, c);
      stripExcelJSCellBorder(cell);
      if (clearObsoleteText && (r > r0 || c > 1)) {
        try {
          cell.value = null;
        } catch (_) {
          /* ignore */
        }
      } else if (c > 1) {
        try {
          cell.value = null;
        } catch (_) {
          /* ignore */
        }
      }
    }
  }

  // Clear L–AN header-side boxes on rows 1–11 (never the table from row 12).
  for (let r = 1; r <= headerEnd; r += 1) {
    for (let c = sideC0; c <= sideC1; c += 1) {
      const cell = worksheet.getCell(r, c);
      stripExcelJSCellBorder(cell);
      if (r < FORM_T_KA_TABLE_BORDER_START_ROW) {
        try {
          cell.value = null;
        } catch (_) {
          /* ignore */
        }
      }
    }
  }
}

/**
 * Clear borders on Month/Year, Establishment, Employer rows — table borders start at row 12.
 * Unmerges header-band cells first so hidden merge slaves do not keep template box lines.
 */
export function clearFormTSEKarnatakaHeaderBandBorders(
  worksheet,
  {
    rowFrom = 9,
    rowTo = 11,
    colFrom = 1,
    colTo = 80,
    remergeHeaderBoxes = true,
  } = {}
) {
  if (!worksheet) return;
  const r0 = Math.max(1, Number(rowFrom) || 9);
  const r1 = Math.max(r0, Number(rowTo) || 11);
  const c0 = Math.max(1, Number(colFrom) || 1);
  const c1 = Math.max(c0, Number(colTo) || c0, FORM_T_KARNATAKA_HEADER_BOX_END_COL);
  const mergeEnd = FORM_T_KARNATAKA_HEADER_BOX_END_COL;

  const savedRows = [];
  if (remergeHeaderBoxes) {
    for (let r = r0; r <= r1; r += 1) {
      const cell = worksheet.getCell(r, 1);
      savedRows.push({
        row: r,
        value: cell.value,
        alignment: cell.alignment ? { ...cell.alignment } : undefined,
        height: worksheet.getRow(r)?.height,
      });
    }
  }

  const merges = Array.isArray(worksheet.model?.merges) ? [...worksheet.model.merges] : [];
  merges.forEach((label) => {
    const parts = String(label || '').split(':');
    if (parts.length !== 2) return;
    try {
      const tl = worksheet.getCell(parts[0]);
      const br = worksheet.getCell(parts[1]);
      if (!tl || !br || br.row < r0 || tl.row > r1) return;
      worksheet.unMergeCells(label);
    } catch (_) {
      /* ignore */
    }
  });

  for (let r = r0; r <= r1; r += 1) {
    for (let c = c0; c <= c1; c += 1) {
      stripExcelJSCellBorder(worksheet.getCell(r, c));
    }
  }

  if (!remergeHeaderBoxes) return;

  savedRows.forEach(({ row, value, alignment, height }) => {
    try {
      worksheet.mergeCells(row, 1, row, mergeEnd);
    } catch (_) {
      /* already merged */
    }
    const cell = worksheet.getCell(row, 1);
    if (value !== undefined) cell.value = value;
    stripExcelJSCellBorder(cell);
    if (alignment) {
      cell.alignment = { ...alignment, wrapText: alignment.wrapText !== false };
    } else {
      cell.alignment = {
        ...(cell.alignment || {}),
        wrapText: true,
        vertical: 'middle',
        horizontal: 'left',
      };
    }
    if (height != null) {
      const wsRow = worksheet.getRow(row);
      if (wsRow) wsRow.height = height;
    }
  });
}

/** Thin full box on Month/Year, Establishment, and Employer header value rows (A–I). */
export function applyFormTSEKarnatakaHeaderValueBoxBorders(
  worksheet,
  endCol = FORM_T_KARNATAKA_HEADER_BOX_END_COL
) {
  if (!worksheet) return;
  const mergeEnd = Math.max(1, Number(endCol) || FORM_T_KARNATAKA_HEADER_BOX_END_COL);
  FORM_T_KARNATAKA_HEADER_BOX_ROWS.forEach((row) => {
    applyExcelJSFullBoxBordersToRange(worksheet, {
      rowFrom: row,
      rowTo: row,
      colFrom: 1,
      colTo: mergeEnd,
    });
  });
}

/**
 * Infer the bordered table rectangle from row 12 through the last body row and
 * rightmost populated column (attendance days, wage bands, merges).
 */
export function inferFormTSEKarnatakaExportBorderRange(worksheet, hints = {}) {
  const startRow = FORM_T_KA_TABLE_BORDER_START_ROW;
  let colTo = Math.max(Number(hints.colTo) || 0, FORM_T_KARNATAKA_HEADER_BOX_END_COL);
  let tableLastRow = Math.max(Number(hints.tableLastRow) || 0, startRow);

  if (!worksheet) {
    return {
      tableHeaderRow: startRow,
      tableLastRow: Math.max(tableLastRow, startRow + 2),
      colFrom: 1,
      colTo: Math.max(colTo, 40),
    };
  }

  const merges = Array.isArray(worksheet.model?.merges) ? worksheet.model.merges : [];
  merges.forEach((label) => {
    const parts = String(label || '').split(':');
    if (parts.length !== 2) return;
    try {
      const br = worksheet.getCell(parts[1]);
      if (br.row >= startRow) colTo = Math.max(colTo, br.col);
    } catch (_) {
      /* ignore bad merge ref */
    }
  });

  for (let r = startRow; r <= startRow + 4; r += 1) {
    for (let c = 1; c <= 80; c += 1) {
      const t = formTSEExcelJsCellText(worksheet, r, c);
      if (t) colTo = Math.max(colTo, c);
    }
  }

  let footerRow = 0;
  for (let r = startRow; r <= 150; r += 1) {
    const a = formTSEExcelJsCellText(worksheet, r, 1).toLowerCase();
    let line = a;
    for (let c = 2; c <= 4; c += 1) {
      const t = formTSEExcelJsCellText(worksheet, r, c).toLowerCase();
      if (t) line += ` ${t}`;
    }
    if (/^date\s*:?$/.test(a) || /^date\s*:/.test(a) || /system\s+generated\s+document/.test(line)) {
      footerRow = r;
      break;
    }
  }

  const bodyScanEnd = footerRow > startRow ? footerRow - 1 : 120;
  let lastContentRow = startRow;
  for (let r = startRow; r <= bodyScanEnd; r += 1) {
    for (let c = 1; c <= Math.max(colTo, 80); c += 1) {
      const t = formTSEExcelJsCellText(worksheet, r, c);
      if (t) {
        lastContentRow = Math.max(lastContentRow, r);
        colTo = Math.max(colTo, c);
      }
    }
  }

  tableLastRow = Math.max(tableLastRow, lastContentRow);
  tableLastRow = Math.max(tableLastRow, lastContentRow + 2, startRow + 2);
  if (footerRow > startRow) {
    tableLastRow = Math.min(tableLastRow, footerRow - 1);
  }

  const remarksCol = findFormTSEKarnatakaRemarksHeaderCol(worksheet);
  const lastTextCol = lastContentColForCap(worksheet, tableLastRow);
  const lastHeaderCol = findFormTSEKarnatakaLastHeaderContentCol(worksheet);
  colTo = Math.max(
    lastTextCol,
    lastHeaderCol,
    Number(hints.colTo) || 0,
    FORM_T_KARNATAKA_HEADER_BOX_END_COL
  );
  // Official Form T ends at Remarks (typically BR). Leftover template numbers/boxes
  // in BS–CB must not widen the bordered grid.
  if (remarksCol > 0) colTo = remarksCol;

  return {
    tableHeaderRow: startRow,
    tableLastRow,
    colFrom: 1,
    colTo,
  };
}

function lastContentColForCap(worksheet, tableLastRow) {
  let last = 0;
  const r1 = Math.max(FORM_T_KA_TABLE_BORDER_START_ROW, Number(tableLastRow) || FORM_T_KA_TABLE_BORDER_START_ROW);
  for (let r = FORM_T_KA_TABLE_BORDER_START_ROW; r <= r1; r += 1) {
    for (let c = 1; c <= 80; c += 1) {
      if (formTSEExcelJsCellText(worksheet, r, c)) last = Math.max(last, c);
    }
  }
  return last;
}

export function findFormTSEKarnatakaRemarksHeaderCol(worksheet) {
  if (!worksheet) return 0;
  let remarks = 0;
  for (let r = FORM_T_KA_TABLE_BORDER_START_ROW; r <= FORM_T_KA_TABLE_BORDER_START_ROW + 2; r += 1) {
    for (let c = 1; c <= 80; c += 1) {
      const t = formTSEExcelJsCellText(worksheet, r, c);
      if (/^remarks?$/i.test(t)) remarks = Math.max(remarks, c);
    }
  }
  return remarks;
}

export function findFormTSEKarnatakaLastHeaderContentCol(worksheet) {
  if (!worksheet) return 0;
  const remarks = findFormTSEKarnatakaRemarksHeaderCol(worksheet);
  if (remarks > 0) return remarks;
  let last = 0;
  for (let r = FORM_T_KA_TABLE_BORDER_START_ROW; r <= FORM_T_KA_TABLE_BORDER_START_ROW + 2; r += 1) {
    for (let c = 1; c <= 80; c += 1) {
      const t = formTSEExcelJsCellText(worksheet, r, c);
      if (t) last = Math.max(last, c);
    }
  }
  return last;
}

/** Strip leftover template boxes to the right of Remarks (BS–CB / cols 71–80). */
export function clearFormTSEKarnatakaTrailingEmptyTableBoxes(
  worksheet,
  afterCol,
  { rowFrom = FORM_T_KA_TABLE_BORDER_START_ROW, rowTo = 80, throughCol = 80 } = {}
) {
  if (!worksheet) return;
  const c0 = Math.max(1, Number(afterCol) || 0) + 1;
  const c1 = Math.max(c0, Number(throughCol) || 80);
  const r0 = Math.max(1, Number(rowFrom) || FORM_T_KA_TABLE_BORDER_START_ROW);
  const r1 = Math.max(r0, Number(rowTo) || r0);
  const merges = Array.isArray(worksheet.model?.merges) ? [...worksheet.model.merges] : [];
  merges.forEach((label) => {
    const parts = String(label || '').split(':');
    if (parts.length !== 2) return;
    try {
      const tl = worksheet.getCell(parts[0]);
      const br = worksheet.getCell(parts[1]);
      if (!tl || !br) return;
      if (tl.col >= c0 && br.row >= r0) worksheet.unMergeCells(label);
    } catch (_) {
      /* ignore */
    }
  });
  for (let r = r0; r <= r1; r += 1) {
    for (let c = c0; c <= c1; c += 1) {
      const cell = worksheet.getCell(r, c);
      try {
        cell.value = null;
      } catch (_) {
        /* ignore */
      }
      stripExcelJSCellBorder(cell);
    }
  }
}

const formTSELeafHeaderNorm = (text) =>
  String(text || '')
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .replace(/\s*\(\s*\d{1,2}\s*\)\s*$/, '')
    .trim()
    .toLowerCase();

function pickFormTSEKarnatakaHeaderLeafAt(worksheet, col) {
  for (let r = FORM_T_KA_TABLE_BORDER_START_ROW + 1; r >= FORM_T_KA_TABLE_BORDER_START_ROW; r -= 1) {
    const t = formTSEExcelJsCellText(worksheet, r, col);
    if (!t || /^\d{1,2}$/.test(t)) continue;
    if (isFormTWageParentBannerText(t)) continue;
    if (/attendance/i.test(t) && t.length > 10) continue;
    return t;
  }
  return '';
}

export const FORM_T_KA_ATTENDANCE_GROUP_LABEL =
  'ATTENDANCE (Please mention the date of suspension of employees, if any)';

function isFormTAttendanceBannerText(raw) {
  const t = String(raw || '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
  if (!t) return false;
  if (t.includes('suspension of employees')) return true;
  return /^attendance\b/.test(t) && t.length > 8;
}

function isFormTAttendanceDayNumber(raw) {
  const t = String(raw || '').trim();
  if (!/^\d{1,2}$/.test(t)) return false;
  const n = Number(t);
  return n >= 1 && n <= 31;
}

/**
 * Official Form T model: merge ATTENDANCE across calendar-day columns (J…, days 1–31)
 * on the group-header row, with day numbers on the leaf row.
 */
export function ensureFormTSEKarnatakaAttendanceGroupMerge(worksheet) {
  if (!worksheet) return;
  const groupRow = FORM_T_KA_TABLE_BORDER_START_ROW;
  const leafRow = groupRow + 1;
  const startCol = FORM_T_KA_ATTENDANCE_START_COL0 + 1;
  const dayCols = [];
  for (let c = startCol; c <= startCol + 30; c += 1) {
    const leaf = formTSEExcelJsCellText(worksheet, leafRow, c);
    const group = formTSEExcelJsCellText(worksheet, groupRow, c);
    if (isFormTAttendanceDayNumber(leaf) || isFormTAttendanceDayNumber(group)) {
      dayCols.push(c);
    }
  }
  let first = dayCols.length ? dayCols[0] : 0;
  let last = dayCols.length ? dayCols[dayCols.length - 1] : 0;
  if (!first) {
    for (let c = startCol; c <= startCol + 30; c += 1) {
      const leaf = formTSEExcelJsCellText(worksheet, leafRow, c);
      const group = formTSEExcelJsCellText(worksheet, groupRow, c);
      if (isFormTAttendanceBannerText(leaf) || isFormTAttendanceBannerText(group)) {
        if (!first) first = c;
        last = c;
      } else if (first) {
        break;
      }
    }
  }
  if (!first || last < first) return;

  formTSEUnmergeCovering(worksheet, groupRow, first, groupRow, last);
  for (let c = first; c <= last; c += 1) {
    const leaf = formTSEExcelJsCellText(worksheet, leafRow, c);
    const group = formTSEExcelJsCellText(worksheet, groupRow, c);
    if (!leaf && isFormTAttendanceDayNumber(group)) {
      try {
        worksheet.getCell(leafRow, c).value = Number(group);
      } catch (_) {
        /* ignore */
      }
    }
    if (c > first) {
      try {
        worksheet.getCell(groupRow, c).value = null;
      } catch (_) {
        /* ignore */
      }
    }
  }
  try {
    if (last > first) worksheet.mergeCells(groupRow, first, groupRow, last);
  } catch (_) {
    /* already merged */
  }
  const cell = worksheet.getCell(groupRow, first);
  cell.value = FORM_T_KA_ATTENDANCE_GROUP_LABEL;
  cell.alignment = {
    ...(cell.alignment || {}),
    horizontal: 'center',
    vertical: 'middle',
    wrapText: true,
  };
  cell.font = { ...(cell.font || {}), bold: true };
  try {
    const wsRow = worksheet.getRow(groupRow);
    if (wsRow) wsRow.height = Math.max(Number(wsRow.height) || 0, 28);
  } catch (_) {
    /* ignore */
  }
}

function ensureFormTSEKarnatakaTableGroupMerges(worksheet) {
  ensureFormTSEKarnatakaAttendanceGroupMerge(worksheet);
  ensureFormTSEKarnatakaWageGroupMerges(worksheet);
}

/**
 * Official Form T model: merge "Earned wages and other allowances" and "Deductions"
 * across their leaf columns on the group-header row (row 12).
 */
export function ensureFormTSEKarnatakaWageGroupMerges(worksheet) {
  if (!worksheet) return;
  const groupRow = FORM_T_KA_TABLE_BORDER_START_ROW;
  const leafRow = groupRow + 1;
  const leaves = [];
  for (let c = FORM_T_KA_ATTENDANCE_START_COL0 + 1; c <= 80; c += 1) {
    const leaf = pickFormTSEKarnatakaHeaderLeafAt(worksheet, c);
    if (!leaf) continue;
    const n = formTSELeafHeaderNorm(leaf);
    leaves.push({ col: c, leaf, n });
  }
  if (!leaves.length) return;

  const isEsi = (n) => /^(esi|e\.?s\.?i\.?)$/.test(n);
  const isNet = (n, leaf) =>
    /net\s+amount/.test(n) || isFormTSEKarnatakaNetAmountPayableHeader(leaf);
  const isPayableOrOt = (n, leaf) =>
    /payable\s+days/.test(n) ||
    /total\s+ot\s+hours|ot\s+hours/.test(n) ||
    isFormTSEKarnatakaTotalOtHoursHeader(leaf);
  const esi = leaves.find((x) => isEsi(x.n));
  const net = leaves.find((x) => isNet(x.n, x.leaf));
  const earnedCandidates = leaves.filter((x) => {
    if (isPayableOrOt(x.n, x.leaf) || isEsi(x.n) || isNet(x.n, x.leaf)) return false;
    if (/mode\s+of\s+payment|signature|thumb|remarks?/.test(x.n)) return false;
    if (net && x.col >= net.col) return false;
    if (esi && x.col >= esi.col) return false;
    return true;
  });
  const deductionCandidates = leaves.filter((x) => {
    if (!esi) return false;
    if (x.col < esi.col) return false;
    if (net && x.col >= net.col) return false;
    if (/mode\s+of\s+payment|signature|thumb|remarks?/.test(x.n)) return false;
    return true;
  });

  const mergeGroup = (startCol, endCol, label) => {
    if (!startCol || !endCol || endCol < startCol) return;
    formTSEUnmergeCovering(worksheet, groupRow, startCol, groupRow, endCol);
    for (let c = startCol; c <= endCol; c += 1) {
      const existing = formTSEExcelJsCellText(worksheet, leafRow, c);
      if (!existing) {
        const fromGroup = formTSEExcelJsCellText(worksheet, groupRow, c);
        if (fromGroup && !isFormTWageParentBannerText(fromGroup) && !/^\d{1,2}$/.test(fromGroup)) {
          try {
            worksheet.getCell(leafRow, c).value = fromGroup;
          } catch (_) {
            /* ignore */
          }
        }
      }
      if (c > startCol) {
        try {
          worksheet.getCell(groupRow, c).value = null;
        } catch (_) {
          /* ignore */
        }
      }
    }
    try {
      if (endCol > startCol) worksheet.mergeCells(groupRow, startCol, groupRow, endCol);
    } catch (_) {
      /* already merged */
    }
    const cell = worksheet.getCell(groupRow, startCol);
    cell.value = label;
    cell.alignment = {
      ...(cell.alignment || {}),
      horizontal: 'center',
      vertical: 'middle',
      wrapText: true,
    };
    cell.font = { ...(cell.font || {}), bold: true };
  };

  if (earnedCandidates.length >= 2) {
    mergeGroup(
      earnedCandidates[0].col,
      earnedCandidates[earnedCandidates.length - 1].col,
      'Earned wages and other allowances'
    );
  }
  if (deductionCandidates.length >= 2) {
    mergeGroup(
      deductionCandidates[0].col,
      deductionCandidates[deductionCandidates.length - 1].col,
      'Deductions'
    );
  }
}

/**
 * Full thin box borders on the export grid from row 12 — column headers, index strip,
 * data rows, and empty template body rows (matches the official Form T_KA model sheet).
 */
export function applyFormTSEKarnatakaFullExportBorders(
  worksheet,
  {
    tableHeaderRow = FORM_T_KA_TABLE_BORDER_START_ROW,
    tableLastRow = 18,
    colFrom = 1,
    colTo = 40,
    headerBoxEndCol = FORM_T_KARNATAKA_HEADER_BOX_END_COL,
    includeHeaderBoxes = false,
  } = {}
) {
  if (!worksheet) return;
  const r0 = Math.max(
    FORM_T_KA_TABLE_BORDER_START_ROW,
    Number(tableHeaderRow) || FORM_T_KA_TABLE_BORDER_START_ROW
  );
  const r1 = Math.max(r0, Number(tableLastRow) || r0);
  const c0 = Math.max(1, Number(colFrom) || 1);
  const c1 = Math.max(c0, Number(colTo) || c0);
  if (includeHeaderBoxes) {
    applyFormTSEKarnatakaHeaderValueBoxBorders(worksheet, headerBoxEndCol);
  } else {
    // Strip template/header-band borders so only row 12+ shows the table grid.
    clearFormTSEKarnatakaHeaderBandBorders(worksheet, { colTo: c1 });
  }
  applyExcelJSFullBoxBordersToRange(worksheet, {
    rowFrom: r0,
    rowTo: r1,
    colFrom: c0,
    colTo: c1,
  });
  if (!includeHeaderBoxes) {
    // Re-strip header band after table paint — template merge slaves can keep box lines.
    clearFormTSEKarnatakaHeaderBandBorders(worksheet, { colTo: c1, remergeHeaderBoxes: true });
  }
}

export const FORM_T_KA_TITLE_COL_FROM = 1;
/** Title/rule lines span A–K (same as Form F) — avoids breaking multi-row template merges. */
export const FORM_T_KA_TITLE_COL_TO = 11;
export const FORM_T_KA_SYSTEM_GENERATED_NOTE = 'This is a System Generated Document';

function formTSETitleLineNorm(text) {
  return String(text || '')
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function looksLikeFormTSEKarnatakaFormTitle(text) {
  return /^form\s*t\b/i.test(formTSETitleLineNorm(text));
}

function looksLikeFormTSEKarnatakaMainSubtitle(text) {
  const t = formTSETitleLineNorm(text).toLowerCase();
  return /combined\s+muster|muster\s+roll\s+cum\s+register|register\s+of\s+wages/.test(t);
}

function looksLikeFormTSEKarnatakaRuleCitation(text) {
  const t = formTSETitleLineNorm(text).toLowerCase();
  if (!t || t.length > 240) return false;
  return (
    /see\s+rule|\[see\s+rule|vide\s+rule|in\s+lieu\s+of|shops\s*&\s*commercial|minimum\s+wages|payment\s+of\s+wages|contract\s+labour|inter-?state\s+migrant/.test(
      t
    )
  );
}

function formTSEUnmergeCovering(worksheet, r1, c1, r2, c2) {
  const merges = Array.isArray(worksheet.model?.merges) ? [...worksheet.model.merges] : [];
  merges.forEach((label) => {
    const parts = String(label || '').split(':');
    if (parts.length !== 2) return;
    try {
      const tl = worksheet.getCell(parts[0]);
      const br = worksheet.getCell(parts[1]);
      if (!tl || !br || br.row < r1 || tl.row > r2 || br.col < c1 || tl.col > c2) return;
      worksheet.unMergeCells(label);
    } catch (_) {
      /* ignore */
    }
  });
}

/** Unmerge only single-row bands — never split vertical template merges on rows 1–8. */
function formTSEUnmergeHorizontalRowBand(worksheet, row, colFrom, colTo) {
  const merges = Array.isArray(worksheet.model?.merges) ? [...worksheet.model.merges] : [];
  merges.forEach((label) => {
    const parts = String(label || '').split(':');
    if (parts.length !== 2) return;
    try {
      const tl = worksheet.getCell(parts[0]);
      const br = worksheet.getCell(parts[1]);
      if (!tl || !br || tl.row !== br.row || tl.row !== row) return;
      if (br.col < colFrom || tl.col > colTo) return;
      worksheet.unMergeCells(label);
    } catch (_) {
      /* ignore */
    }
  });
}

function formTSECaptureRowTextDeep(worksheet, row, matcher, scanCols = 28) {
  let best = formTSECaptureRowText(worksheet, row, matcher, scanCols);
  const merges = Array.isArray(worksheet.model?.merges) ? worksheet.model.merges : [];
  merges.forEach((label) => {
    const parts = String(label || '').split(':');
    if (parts.length !== 2) return;
    try {
      const tl = worksheet.getCell(parts[0]);
      const br = worksheet.getCell(parts[1]);
      if (!tl || !br || tl.row > row || br.row < row) return;
      const t = formTSEExcelJsCellText(worksheet, tl.row, tl.col);
      if (matcher(t) && t.length > best.length) best = t;
    } catch (_) {
      /* ignore */
    }
  });
  return best;
}

function formTSECaptureRowText(worksheet, row, matcher, scanCols = 28) {
  let best = '';
  for (let c = 1; c <= scanCols; c += 1) {
    const t = formTSEExcelJsCellText(worksheet, row, c);
    if (!t) continue;
    if (matcher(t) && t.length >= best.length) best = t;
  }
  return best;
}

function formTSEWriteLeftTitleBand(
  worksheet,
  row,
  text,
  colFrom,
  colTo,
  { bold = false, size = 11 } = {}
) {
  if (!worksheet || row < 1 || !text) return;
  const c0 = Math.max(1, colFrom);
  const c1 = Math.max(c0, colTo);
  formTSEUnmergeHorizontalRowBand(worksheet, row, c0, c1);
  try {
    worksheet.mergeCells(row, c0, row, c1);
  } catch (_) {
    /* overlap — still write left cell */
  }
  const cell = worksheet.getCell(row, c0);
  cell.value = text;
  const alignment = {
    horizontal: 'left',
    vertical: 'middle',
    wrapText: true,
    indent: 0,
    textRotation: 0,
  };
  cell.alignment = alignment;
  try {
    const prev = cell.style && typeof cell.style === 'object' ? { ...cell.style } : {};
    cell.style = { ...prev, alignment: { ...(prev.alignment || {}), ...alignment } };
  } catch (_) {
    /* alignment property is enough on most ExcelJS builds */
  }
  cell.font = {
    ...(cell.font || {}),
    bold: bold ?? cell.font?.bold ?? false,
    size: cell.font?.size || size,
  };
  const wsRow = worksheet.getRow(row);
  if (wsRow) {
    const len = String(text).length;
    const needed = len > 120 ? 36 : len > 80 ? 28 : len > 40 ? 22 : 16;
    wsRow.height = needed;
  }
}

export function resolveFormTSEKarnatakaTitleSpanCol(worksheet, hints = {}) {
  void worksheet;
  void hints;
  return FORM_T_KA_TITLE_COL_TO;
}

/**
 * Write CLRA "in lieu of" rule lines into rows 4–8 (column A, left-aligned, no boxes).
 */
export function writeFormTSEKarnatakaClraInLieuLines(worksheet, options = {}) {
  if (!worksheet) return false;
  const colFrom = FORM_T_KA_TITLE_COL_FROM;
  const colTo = resolveFormTSEKarnatakaTitleSpanCol(worksheet, options);
  const lines = Array.isArray(options.lines) ? options.lines : FORM_T_KA_CLRA_IN_LIEU_LINES;
  lines.forEach((text, index) => {
    const row = 4 + index;
    if (row > FORM_T_KA_OBSOLETE_TITLE_ROW_TO) return;
    formTSEWriteLeftTitleBand(worksheet, row, text, colFrom, colTo, {
      bold: true,
      size: index === 0 ? 10 : 9,
    });
    try {
      const wsRow = worksheet.getRow(row);
      if (wsRow) {
        wsRow.hidden = false;
        const needed = String(text).length > 120 ? 36 : String(text).length > 80 ? 28 : 18;
        wsRow.height = Math.max(Number(wsRow.height) || 0, needed);
      }
    } catch (_) {
      /* ignore */
    }
  });
  return true;
}

/**
 * Ensure rows 4–8 are visible with the CLRA in-lieu text (not collapsed).
 */
export function expandFormTSEKarnatakaClraTitleRows(worksheet) {
  if (!worksheet) return;
  for (let r = 4; r <= FORM_T_KA_OBSOLETE_TITLE_ROW_TO; r += 1) {
    try {
      const wsRow = worksheet.getRow(r);
      if (!wsRow) continue;
      wsRow.hidden = false;
      wsRow.height = Math.max(Number(wsRow.height) || 0, 18);
    } catch (_) {
      /* ignore */
    }
  }
  [1, 2, 3].forEach((r) => {
    try {
      const wsRow = worksheet.getRow(r);
      if (!wsRow) return;
      wsRow.hidden = false;
      const text = formTSEExcelJsCellText(worksheet, r, 1);
      const needed = r === 3 ? 28 : r === 2 ? 22 : 18;
      wsRow.height = Math.max(Number(wsRow.height) || 0, text.length > 80 ? needed + 8 : needed);
    } catch (_) {
      /* ignore */
    }
  });
  FORM_T_KARNATAKA_HEADER_BOX_ROWS.forEach((r) => {
    try {
      const wsRow = worksheet.getRow(r);
      if (!wsRow) return;
      wsRow.hidden = false;
      wsRow.height = Math.max(Number(wsRow.height) || 0, 22);
    } catch (_) {
      /* ignore */
    }
  });
}

/** @deprecated CLRA rows 4–8 stay visible — delegates to expandFormTSEKarnatakaClraTitleRows. */
export function collapseFormTSEKarnatakaObsoleteTitleRows(worksheet) {
  expandFormTSEKarnatakaClraTitleRows(worksheet);
}

/**
 * Left-align FORM T / muster subtitle / Rule 24(9-B) / CLRA "in lieu of" lines.
 * Rows 3 = Rule citation; rows 4–8 = in-lieu legal references (no boxes). L–AN boxes cleared.
 */
export function ensureFormTSEKarnatakaTitleLayout(worksheet, options = {}) {
  if (!worksheet || !looksLikeFormTSEKarnatakaWorksheet(worksheet)) return false;
  const colFrom = FORM_T_KA_TITLE_COL_FROM;
  const colTo = resolveFormTSEKarnatakaTitleSpanCol(worksheet, options);

  const colA = worksheet.getColumn(1);
  colA.hidden = false;
  colA.width = Math.max(Number(colA.width) || 0, 18);

  // Capture title/subtitle before clearing the obsolete band.
  let formTitle = '';
  let subtitle = '';
  for (let r = 1; r <= 8; r += 1) {
    if (!formTitle) {
      formTitle = formTSECaptureRowTextDeep(worksheet, r, looksLikeFormTSEKarnatakaFormTitle);
    }
    if (!subtitle) {
      subtitle = formTSECaptureRowTextDeep(worksheet, r, looksLikeFormTSEKarnatakaMainSubtitle);
    }
  }
  formTitle = formTitle || 'FORM T';
  subtitle = subtitle || FORM_T_KA_SUBTITLE;

  clearFormTSEKarnatakaObsoleteHeaderBoxes(worksheet, { clearObsoleteText: true });

  formTSEWriteLeftTitleBand(worksheet, 1, formTitle, colFrom, colTo, { bold: true, size: 14 });
  formTSEWriteLeftTitleBand(worksheet, 2, subtitle, colFrom, colTo, { bold: true, size: 12 });
  formTSEWriteLeftTitleBand(worksheet, 3, FORM_T_KA_RULE_CITATION, colFrom, colTo, {
    bold: true,
    size: 10,
  });
  writeFormTSEKarnatakaClraInLieuLines(worksheet, { colTo });
  expandFormTSEKarnatakaClraTitleRows(worksheet);

  return true;
}

/** Center the footer note across the full Form T table width. */
export function ensureFormTSEKarnatakaSystemGeneratedNoteCentered(
  worksheet,
  colTo = FORM_T_KARNATAKA_HEADER_BOX_END_COL
) {
  if (!worksheet) return false;
  const spanTo = Math.max(
    FORM_T_KARNATAKA_HEADER_BOX_END_COL,
    Number(colTo) || FORM_T_KARNATAKA_HEADER_BOX_END_COL
  );
  let noteRow = 0;
  for (let r = 1; r <= 150; r += 1) {
    const serialA = formTSEExcelJsCellText(worksheet, r, 1);
    const nameB = formTSEExcelJsCellText(worksheet, r, 2);
    // Never treat an employee identity row as the footer note (wipes the last name).
    if (/^\d{1,4}$/.test(serialA) || looksLikeFormTSEEmployeeNameCell(nameB)) continue;
    for (let c = 1; c <= 8; c += 1) {
      const t = formTSEExcelJsCellText(worksheet, r, c);
      if (
        /this\s+is\s+a\s+system\s+generated\s+document|system\s+generated\s+document|stem\s+generated\s+document/i.test(
          t
        )
      ) {
        noteRow = r;
        break;
      }
    }
    if (noteRow > 0) break;
  }
  if (noteRow < 1) return false;

  for (let c = 1; c <= spanTo + 8; c += 1) {
    const t = formTSEExcelJsCellText(worksheet, noteRow, c);
    if (
      /this\s+is\s+a\s+system\s+generated\s+document|system\s+generated\s+document|stem\s+generated\s+document/i.test(
        t
      ) &&
      c !== 1
    ) {
      worksheet.getCell(noteRow, c).value = '';
    }
  }

  formTSEUnmergeCovering(worksheet, noteRow, 1, noteRow, spanTo);
  try {
    worksheet.mergeCells(noteRow, 1, noteRow, spanTo);
  } catch (_) {
    /* overlap — write A cell only */
  }

  const cell = worksheet.getCell(noteRow, 1);
  cell.value = FORM_T_KA_SYSTEM_GENERATED_NOTE;
  const alignment = { horizontal: 'center', vertical: 'middle', wrapText: false, indent: 0 };
  cell.alignment = alignment;
  try {
    const prev = cell.style && typeof cell.style === 'object' ? { ...cell.style } : {};
    cell.style = { ...prev, alignment: { ...(prev.alignment || {}), ...alignment } };
  } catch (_) {
    /* ignore */
  }
  return true;
}

/** True when the active sheet is Karnataka Form T (Combined Muster Roll). */
export function looksLikeFormTSEKarnatakaWorksheet(worksheet) {
  if (!worksheet || typeof worksheet.getCell !== 'function') return false;
  let hasFormT = false;
  let hasMuster = false;
  for (let r = 1; r <= 8; r += 1) {
    for (let c = 1; c <= 8; c += 1) {
      const t = formTSEExcelJsCellText(worksheet, r, c).toLowerCase();
      if (!t) continue;
      if (/^form\s*t\b/.test(t)) hasFormT = true;
      if (/combined\s+muster|muster\s+roll\s+cum\s+register/.test(t)) hasMuster = true;
      if (hasFormT && hasMuster) return true;
    }
  }
  const monthRow = formTSEExcelJsCellText(worksheet, 9, 1).toLowerCase();
  const nameHeader = formTSEExcelJsCellText(worksheet, 12, 2).toLowerCase();
  if (
    /month\s*\/?\s*year/.test(monthRow) &&
    (/name\s+of\s+employee/.test(nameHeader) || /principal\s+employer/.test(nameHeader))
  ) {
    return true;
  }
  return false;
}

/** Row-12+ table borders, title band, and centered footer — every Form T download pass. */
export function finalizeFormTSEKarnatakaWorksheetExportBorders(worksheet, hints = {}) {
  if (!worksheet || !looksLikeFormTSEKarnatakaWorksheet(worksheet)) return;
  const range = inferFormTSEKarnatakaExportBorderRange(worksheet, hints);
  clearFormTSEKarnatakaTrailingEmptyTableBoxes(worksheet, range.colTo, {
    rowFrom: 1,
    rowTo: Math.max(range.tableLastRow + 4, 80),
    throughCol: Math.max(80, Number(worksheet.columnCount) || 80),
  });
  ensureFormTSEKarnatakaTableGroupMerges(worksheet);
  ensureFormTSEKarnatakaTitleLayout(worksheet, { colTo: range.colTo });
  renameFormTSEKarnatakaHeaderLabels(worksheet);
  applyFormTSEKarnatakaFullExportBorders(worksheet, {
    ...range,
    includeHeaderBoxes: hints.includeHeaderBoxes === true,
  });
  // Template merge slaves can keep L–AN / rows 3–8 box lines after table paint.
  clearFormTSEKarnatakaObsoleteHeaderBoxes(worksheet);
  clearFormTSEKarnatakaTrailingEmptyTableBoxes(worksheet, range.colTo, {
    rowFrom: FORM_T_KA_TABLE_BORDER_START_ROW,
    rowTo: Math.max(range.tableLastRow + 4, 80),
    throughCol: Math.max(80, Number(worksheet.columnCount) || 80),
  });
  ensureFormTSEKarnatakaTableGroupMerges(worksheet);
  ensureFormTSEKarnatakaTitleLayout(worksheet, { colTo: range.colTo });
  renameFormTSEKarnatakaHeaderLabels(worksheet);
  expandFormTSEKarnatakaClraTitleRows(worksheet);
  ensureFormTSEKarnatakaSystemGeneratedNoteCentered(worksheet, range.colTo);
  applyFormTSEKarnatakaExportBold(worksheet, {
    rowTo: Math.max(range.tableLastRow + 8, 80),
    colTo: Math.max(range.colTo, 80),
  });
}

/** Bold every populated Form T Excel cell (title, headers, employee data, footer). */
export function applyFormTSEKarnatakaExportBold(worksheet, { rowTo = 80, colTo = 80 } = {}) {
  if (!worksheet) return;
  const r1 = Math.max(1, Number(rowTo) || 80);
  const c1 = Math.max(1, Number(colTo) || 80);
  try {
    worksheet.eachRow({ includeEmpty: false }, (row) => {
      if (row.number > r1) return;
      row.eachCell({ includeEmpty: false }, (cell) => {
        if (Number(cell.col) > c1) return;
        cell.font = { ...(cell.font || {}), bold: true };
      });
    });
  } catch (_) {
    for (let r = 1; r <= r1; r += 1) {
      for (let c = 1; c <= c1; c += 1) {
        if (!formTSEExcelJsCellText(worksheet, r, c)) continue;
        const cell = worksheet.getCell(r, c);
        cell.font = { ...(cell.font || {}), bold: true };
      }
    }
  }
}

/** Apply row-12+ full box borders to every worksheet in a Form T workbook buffer. */
export async function applyFormTSEKarnatakaExportBordersToBuffer(arrayBuffer, hints = {}) {
  if (!arrayBuffer || arrayBuffer.byteLength < 32) return arrayBuffer;
  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(arrayBuffer);
    const sheets = workbook.worksheets || [];
    if (sheets.length === 0) return arrayBuffer;
    const headerFormData =
      hints.headerFormData && typeof hints.headerFormData === 'object' ? hints.headerFormData : null;
    const parsedFormHeader = hints.parsedFormHeader || null;
    sheets.forEach((worksheet) => {
      if (!looksLikeFormTSEKarnatakaWorksheet(worksheet) && hints.force !== true) return;
      finalizeFormTSEKarnatakaWorksheetExportBorders(worksheet, hints);
      // Stamp Establishment/Employer onto rows 10–11 after title-band cleanup.
      if (headerFormData) {
        writeFormTSEHeaderFieldsToWorksheet(worksheet, headerFormData, parsedFormHeader);
        expandFormTSEHeaderValueBoxes(worksheet, FORM_T_KARNATAKA_HEADER_BOX_END_COL);
        writeFormTSEKarnatakaClraInLieuLines(worksheet);
        expandFormTSEKarnatakaClraTitleRows(worksheet);
      }
      restoreFormTSEKarnatakaEmployeeNameColumnHeader(worksheet);
      const stampRows = hints.mappedData || hints.rows;
      const stampHeaders = hints.headers || hints.headersToUse || [];
      if (Array.isArray(stampRows) && stampRows.length > 0) {
        stampFormTSEKarnatakaEmployeeIdentity(
          worksheet,
          stampRows,
          stampHeaders,
          hints.employees || []
        );
      }
    });
    resetExcelJsWorkbookActiveSheet(workbook);
    const written = await workbook.xlsx.writeBuffer();
    const stampHeaders = hints.headers || hints.headersToUse || [];
    const stampRows = overlayFormTSEPeopleNamesOntoRows(
      hints.mappedData || hints.rows || [],
      stampHeaders,
      hints.employees || []
    );
    if (Array.isArray(stampRows) && stampRows.length > 0) {
      return forceFormTSEIdentityCellsInSheetXml(
        written,
        stampRows,
        stampHeaders,
        hints.employees || []
      );
    }
    return stripFormTSEDataIdentityMergesFromXlsx(written);
  } catch (err) {
    console.warn('Form T Karnataka export borders failed:', err);
    const stampHeaders = hints.headers || hints.headersToUse || [];
    const stampRows = overlayFormTSEPeopleNamesOntoRows(
      hints.mappedData || hints.rows || [],
      stampHeaders,
      hints.employees || []
    );
    if (Array.isArray(stampRows) && stampRows.length > 0) {
      return forceFormTSEIdentityCellsInSheetXml(
        arrayBuffer,
        stampRows,
        stampHeaders,
        hints.employees || []
      );
    }
    return arrayBuffer;
  }
}

const formTSEHeaderNorm = (txt) =>
  String(txt || '')
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9 ]/g, '');

/** Strip statutory column suffix e.g. "Name of Employee (2)" → normalized label key. */
export function formTSEEmployeeHeaderKeyNorm(header) {
  return String(header || '')
    .replace(/\s*\(\s*\d{1,2}\s*\)\s*$/i, '')
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9 ]/g, '');
}

export function isFormTSESerialNumberHeader(h) {
  const s = formTSEEmployeeHeaderKeyNorm(h);
  return (
    /^sno$/.test(s.replace(/\s/g, '')) ||
    /^slno$/.test(s.replace(/\s/g, '')) ||
    /^serialno$/.test(s.replace(/\s/g, '')) ||
    (s.includes('serial') && !s.includes('register'))
  );
}

export function isFormTSECanonicalEmployeeNameHeader(h) {
  const s = formTSEEmployeeHeaderKeyNorm(h);
  if (!s.includes('name')) return false;
  return (
    s.includes('employee') ||
    s.includes('workman') ||
    s.includes('worker') ||
    /nameoftheemployee/.test(s.replace(/\s/g, ''))
  );
}

/** CLRA leftover: table column 2 labelled "Name and address of principal employer". */
export function isFormTSECorruptedClraEmployeeNameHeader(h) {
  const s = formTSEEmployeeHeaderKeyNorm(h);
  return s.includes('name') && /principal\s*employer/.test(s);
}

export function isFormTSEEmployeeNameHeader(h) {
  return isFormTSECanonicalEmployeeNameHeader(h) || isFormTSECorruptedClraEmployeeNameHeader(h);
}

/** True when rows contain real employee names (not template 1,2,3… index strips). */
export function formTSEDownloadHasSubstantiveRows(rows, headers = []) {
  if (!Array.isArray(rows) || rows.length === 0) return false;
  const hdrs =
    Array.isArray(headers) && headers.length > 0
      ? headers
      : rows[0] && typeof rows[0] === 'object' && !Array.isArray(rows[0])
        ? Object.keys(rows[0]).filter((k) => !String(k).startsWith('__'))
        : [];
  const nameHeaders = hdrs.filter((h) => isFormTSEEmployeeNameHeader(h));
  for (let i = 0; i < rows.length; i += 1) {
    const row = rows[i];
    if (!row || typeof row !== 'object') continue;
    if (nameHeaders.length > 0) {
      const hasName = nameHeaders.some((h) => {
        const v = String(getFormTSEKarnatakaRowValueForHeader(row, h) ?? '').trim();
        return looksLikeFormTSEEmployeeNameCell(v);
      });
      if (hasName) return true;
    } else {
      const vals = Object.values(row)
        .map((v) => String(v ?? '').trim())
        .filter(Boolean);
      if (vals.some((v) => looksLikeFormTSEEmployeeNameCell(v))) return true;
    }
  }
  return false;
}

/** Build Form T export rows from People employees when the modal grid is empty/stale. */
export function buildFormTSERowsFromEmployees(employees, headers, helpers = {}) {
  const hdrs = Array.isArray(headers) ? headers.filter((h) => String(h || '').trim()) : [];
  const list = Array.isArray(employees) ? employees : [];
  if (!hdrs.length || !list.length) return [];
  return list.map((empItem, index) => {
    const emp =
      empItem && typeof empItem === 'object'
        ? empItem.Employee || empItem.employee || empItem
        : null;
    const row = {};
    hdrs.forEach((h) => {
      row[h] = '';
    });
    if (emp) {
      applyFormTSEKarnatakaEmployeeToRow(row, emp, hdrs, {
        ...helpers,
        rowIndex: index,
      });
    }
    return row;
  });
}

/**
 * Prefer live modal rows; if names are missing, rebuild from People employees and
 * overlay any manual edits from the modal grid.
 */
export function resolveFormTSERowsForExport({
  liveRows = [],
  headers = [],
  employees = [],
  helpers = {},
} = {}) {
  const hdrs = Array.isArray(headers) ? headers.filter((h) => String(h || '').trim()) : [];
  const fromLive = Array.isArray(liveRows)
    ? liveRows
        .filter((row) => row && typeof row === 'object')
        .map((row) => ({ ...row }))
    : [];
  const overlaid = overlayFormTSEPeopleNamesOntoRows(fromLive, hdrs, employees, helpers);
  const namedPayroll = overlayFormTSEPayrollOntoRowsByName(overlaid, hdrs, fromLive);
  if (formTSEDownloadHasSubstantiveRows(namedPayroll, hdrs)) {
    return prepareFormTSEExportRows(namedPayroll, hdrs, hdrs);
  }
  const fromEmployees = buildFormTSERowsFromEmployees(employees, hdrs, helpers);
  if (!formTSEDownloadHasSubstantiveRows(fromEmployees, hdrs)) {
    return prepareFormTSEExportRows(fromLive, hdrs, hdrs);
  }
  // Overlay Autofill payroll/edits by FirstName + LastName (not row index).
  const overlay = fromLive.filter((row) => formTSEDownloadHasSubstantiveRows([row], hdrs));
  if (overlay.length === 0) {
    return prepareFormTSEExportRows(fromEmployees, hdrs, hdrs);
  }
  const findLiveByName = (row) => {
    const named = getFormTSEKarnatakaEmployeeNameFromRow(row, hdrs);
    if (!named) return null;
    return (
      overlay.find((live) =>
        personNamesMatch(
          named,
          getFormTSEKarnatakaEmployeeNameFromRow(live, hdrs) ||
            String(live.__employeeLookupName || live.__employeelookupname || '').trim()
        )
      ) || null
    );
  };
  const merged = fromEmployees.map((row, i) => {
    const edit = findLiveByName(row) || overlay[i];
    if (!edit) return row;
    const out = { ...row };
    hdrs.forEach((h) => {
      const v = getFormTSEKarnatakaRowValueForHeader(edit, h);
      if (v != null && String(v).trim() !== '') out[h] = v;
    });
    return out;
  });
  for (let i = 0; i < overlay.length; i += 1) {
    const named = getFormTSEKarnatakaEmployeeNameFromRow(overlay[i], hdrs);
    const already = merged.some((row) =>
      personNamesMatch(named, getFormTSEKarnatakaEmployeeNameFromRow(row, hdrs))
    );
    if (!already) merged.push({ ...overlay[i] });
  }
  return prepareFormTSEExportRows(
    overlayFormTSEPayrollOntoRowsByName(merged, hdrs, fromLive),
    hdrs,
    hdrs
  );
}

export function isFormTSEFatherHusbandHeader(h) {
  const s = formTSEEmployeeHeaderKeyNorm(h);
  return s.includes('father') || s.includes('husband') || s.includes('spouse');
}

export function isFormTSEGenderHeader(h) {
  const s = formTSEEmployeeHeaderKeyNorm(h);
  return s === 'gender' || s === 'sex' || (s.includes('male') && s.includes('female'));
}

export function isFormTSEDesignationHeader(h) {
  const s = formTSEEmployeeHeaderKeyNorm(h);
  return s.includes('designation') || s.includes('department');
}

export function isFormTSEDateOfJoiningHeader(h) {
  const s = formTSEEmployeeHeaderKeyNorm(h);
  return (
    (s.includes('date') && s.includes('join')) ||
    s.includes('dateofjoining') ||
    (s.includes('entry') && s.includes('employ'))
  );
}

/** ESIC registration number — not wage-register ESI deduction column (26). */
export function isFormTSEEsicRegistrationHeader(h) {
  const raw = String(h || '');
  const s = formTSEEmployeeHeaderKeyNorm(h);
  if (!s) return false;
  if (/\(\s*2[0-9]\s*\)/.test(raw) && !/\(\s*[1-9]\s*\)/.test(raw)) return false;
  if (s.includes('deduction')) return false;
  if (/employeesstateinsurance|stateinsurancecorporation/.test(s.replace(/\s/g, ''))) return true;
  if (/\besic\b/.test(s) && (/\bno\b/.test(s) || s.includes('number'))) return true;
  if (/\besi\b/.test(s) && (/\bno\b/.test(s) || s.includes('number'))) return true;
  return false;
}

/** UAN / EPF registration — not PF deduction column (27). */
export function isFormTSEUanRegistrationHeader(h) {
  const raw = String(h || '');
  const s = formTSEEmployeeHeaderKeyNorm(h);
  if (!s) return false;
  if (/\(\s*2[0-9]\s*\)/.test(raw) && !/\(\s*[1-9]\s*\)/.test(raw)) return false;
  if (s.includes('deduction')) return false;
  if (/\buan\b/.test(s) && (/\bno\b/.test(s) || s.includes('number'))) return true;
  if (/\bepf\b/.test(s) && (/\bno\b/.test(s) || s.includes('uan') || s.includes('number'))) return true;
  if (s.includes('provident') && (s.includes('uan') || s.includes('no'))) return true;
  return false;
}

export function isFormTSEEmployeeIdentityHeader(h) {
  return (
    isFormTSESerialNumberHeader(h) ||
    isFormTSEEmployeeNameHeader(h) ||
    isFormTSEFatherHusbandHeader(h) ||
    isFormTSEGenderHeader(h) ||
    isFormTSEDesignationHeader(h) ||
    isFormTSEDateOfJoiningHeader(h) ||
    isFormTSEEsicRegistrationHeader(h) ||
    isFormTSEUanRegistrationHeader(h)
  );
}

export function isFormTSEKarnatakaPayrollRegisterHeader(header, allHeaders = []) {
  if (!header) return false;
  if (isFormTSEAttendanceDayHeader(header) || isFormTSEEmployeeIdentityHeader(header)) return false;
  if (isFormTSEWagesFixedIncludingVDAHeader(header)) return false;
  if (isFormTSEKarnatakaTotalOtHoursHeader(header)) return true;
  if (isFormTSEKarnatakaEarnedWageComponentHeader(header)) return true;
  if (isFormTSEKarnatakaEarningsTotalHeader(header, allHeaders)) return true;
  if (isFormTSEKarnatakaDeductionTotalHeader(header, allHeaders)) return true;
  if (isFormTSEKarnatakaNetAmountPayableHeader(header)) return true;
  const s = formTSEEmployeeHeaderKeyNorm(header);
  if (!s) return false;
  if (s.includes('mode') && s.includes('payment')) return true;
  if (s === 'pf' || s === 'epf') return true;
  if ((/\bpf\b/.test(s) || s.includes('provident')) && !s.includes('uan') && !s.includes('registration')) {
    return true;
  }
  if (s === 'pt' || s.includes('professional tax') || (s.includes('profession') && s.includes('tax'))) {
    return true;
  }
  if (s.includes('esi') && !s.includes('no') && !s.includes('registration')) return true;
  if (s.includes('salary adv') || /\bfines?\b/.test(s) || s.includes('damage')) return true;
  if (s.includes('insurance') || s.includes('society') || s === 'tos') return true;
  if (s.includes('deduction')) return true;
  return false;
}

export function formatFormTSEKarnatakaEmployeeName(emp = {}) {
  const fn = String(emp.FirstName || emp['FirstName'] || emp.firstName || emp['First Name'] || '').trim();
  const ln = String(emp.LastName || emp['LastName'] || emp.lastName || emp['Last Name'] || '').trim();
  if (fn && ln) return `${fn} ${ln}`;
  return (
    fn ||
    ln ||
    String(
      emp.Name ||
        emp['Name'] ||
        emp.EmployeeName ||
        emp['Employee Name'] ||
        emp.Name1 ||
        emp['Name1'] ||
        ''
    ).trim()
  );
}

function pickFormTSEEmployeeScalar(emp, keys) {
  if (!emp || typeof emp !== 'object') return '';
  for (const key of keys) {
    const v = emp[key];
    if (v != null && String(v).trim() !== '') return String(v).trim();
  }
  return '';
}

/** Read row cell when header keys changed after template repair (suffix / rename). */
export function getFormTSEKarnatakaRowValueForHeader(row, header) {
  if (!row || typeof row !== 'object') return '';
  if (Object.prototype.hasOwnProperty.call(row, header)) {
    const direct = row[header];
    return direct == null ? '' : direct;
  }
  const target = formTSEEmployeeHeaderKeyNorm(header);
  if (!target) return '';
  const hit = Object.keys(row).find((k) => formTSEEmployeeHeaderKeyNorm(k) === target);
  if (!hit) return '';
  const v = row[hit];
  return v == null ? '' : v;
}

/** Employee name from a modal/export row — never Father / Husband's Name. */
export function getFormTSEKarnatakaEmployeeNameFromRow(row, headers = []) {
  if (!row || typeof row !== 'object') return '';
  const hdrs = Array.isArray(headers) ? headers : [];
  const fatherHeader =
    hdrs.find((h) => isFormTSEFatherHusbandHeader(h)) ||
    Object.keys(row).find((k) => isFormTSEFatherHusbandHeader(k)) ||
    '';
  const fatherText = fatherHeader
    ? String(getFormTSEKarnatakaRowValueForHeader(row, fatherHeader) ?? '').trim()
    : '';
  const accept = (text) => {
    const named = String(text ?? '').trim();
    if (!looksLikeFormTSEEmployeeNameCell(named)) return '';
    if (fatherText && named.toLowerCase() === fatherText.toLowerCase()) return '';
    return named;
  };
  const tryHeader = (header) => {
    if (!header) return '';
    return accept(getFormTSEKarnatakaRowValueForHeader(row, header));
  };
  const pickFrom = (list, predicate) => {
    for (let i = 0; i < list.length; i += 1) {
      if (!predicate(list[i])) continue;
      const named = tryHeader(list[i]);
      if (named) return named;
    }
    return '';
  };
  const keys = Object.keys(row).filter((k) => !String(k).startsWith('__'));
  const fromHeaders =
    pickFrom(hdrs, isFormTSECanonicalEmployeeNameHeader) ||
    pickFrom(keys, isFormTSECanonicalEmployeeNameHeader) ||
    pickFrom(hdrs, isFormTSEEmployeeNameHeader) ||
    pickFrom(keys, isFormTSEEmployeeNameHeader) ||
    '';
  const lookup = accept(row.__employeeLookupName || row.__employeelookupname);
  // Autofill shows People (__employeeLookupName) even when the Name cell leaked Father ("abc").
  if (lookup) {
    if (!fromHeaders) return lookup;
    if (fromHeaders.toLowerCase() === lookup.toLowerCase()) return lookup;
    if (fatherText && fromHeaders.toLowerCase() === fatherText.toLowerCase()) return lookup;
    if (fromHeaders.length <= 3 && lookup.length > fromHeaders.length) return lookup;
  }
  return fromHeaders || lookup || '';
}

/**
 * Fill Name when the live/export row is missing it (or still has Father "abc")
 * using People autofill rows. Autofill can show People names while the Name cell is empty.
 */
export function overlayFormTSEPeopleNamesOntoRows(rows, headers, employees, helpers = {}) {
  const list = Array.isArray(rows)
    ? rows.filter((row) => row && typeof row === 'object').map((row) => ({ ...row }))
    : [];
  let hdrs = Array.isArray(headers) ? headers.filter((h) => String(h || '').trim()) : [];
  if (!hdrs.length && list[0]) {
    hdrs = Object.keys(list[0]).filter((k) => !String(k).startsWith('__'));
  }
  const fromPeople = buildFormTSERowsFromEmployees(employees, hdrs, helpers);
  const nameHeaders = hdrs.filter((h) => isFormTSEEmployeeNameHeader(h));
  const fatherHeader = hdrs.find((h) => isFormTSEFatherHusbandHeader(h)) || '';
  const getFather = (row) =>
    fatherHeader ? String(getFormTSEKarnatakaRowValueForHeader(row, fatherHeader) ?? '').trim() : '';
  const getRawName = (row) => {
    for (let i = 0; i < nameHeaders.length; i += 1) {
      const v = String(getFormTSEKarnatakaRowValueForHeader(row, nameHeaders[i]) ?? '').trim();
      if (v) return v;
    }
    return String(row?.__employeeLookupName || row?.__employeelookupname || '').trim();
  };
  const applyPeopleName = (row, peopleRow) => {
    if (!peopleRow) return row;
    const peopleName = getFormTSEKarnatakaEmployeeNameFromRow(peopleRow, hdrs);
    if (!peopleName) return row;
    const out = { ...row };
    if (nameHeaders.length > 0) {
      nameHeaders.forEach((h) => {
        out[h] = peopleName;
      });
    } else {
      out['Name of Employee'] = peopleName;
    }
    out.__employeeLookupName = peopleName;
    const peopleFather = getFather(peopleRow);
    if (peopleFather && fatherHeader && !getFather(out)) out[fatherHeader] = peopleFather;
    return out;
  };
  const fatherValues = new Set();
  const collectFather = (row) => {
    const f = getFather(row);
    if (f) fatherValues.add(f.toLowerCase());
  };
  list.forEach(collectFather);
  fromPeople.forEach(collectFather);
  const outLen = Math.max(list.length, fromPeople.length);
  const out = [];
  for (let i = 0; i < outLen; i += 1) {
    let row = list[i] ? { ...list[i] } : fromPeople[i] ? { ...fromPeople[i] } : null;
    if (!row) continue;
    const prevFather = i > 0 ? getFather(out[i - 1] || list[i - 1] || {}) : '';
    const named = getFormTSEKarnatakaEmployeeNameFromRow(row, hdrs);
    const rawName = getRawName(row);
    const ownFather = getFather(row);
    const leaked =
      !named ||
      (prevFather && rawName && rawName.toLowerCase() === prevFather.toLowerCase()) ||
      (!ownFather && rawName && fatherValues.has(rawName.toLowerCase()));
    if (leaked) row = applyPeopleName(row, fromPeople[i]);
    out.push(row);
  }
  return out;
}

/**
 * Copy payable days / Basic / HRA / Total / PF / PT onto the row whose
 * FirstName + LastName match Autofill — never by column index.
 * Do not copy wages onto a row with no attendance for the month.
 */
export function overlayFormTSEPayrollOntoRowsByName(rows, headers, sourceRows = []) {
  const hdrs = Array.isArray(headers) ? headers.filter((h) => String(h || '').trim()) : [];
  const list = Array.isArray(rows)
    ? rows.map((row) => (row && typeof row === 'object' ? { ...row } : row))
    : [];
  const sources = (Array.isArray(sourceRows) && sourceRows.length > 0 ? sourceRows : list).filter(
    (row) => row && typeof row === 'object'
  );
  const payrollHeaders = hdrs.filter((h) => isFormTSEKarnatakaPayrollRegisterHeader(h, hdrs));
  if (!payrollHeaders.length || !list.length) return list;
  const hasAttendanceBand = listFormTSEAttendanceDayHeaders(hdrs).length > 0;
  const rowName = (row) =>
    String(
      getFormTSEKarnatakaEmployeeNameFromRow(row, hdrs) ||
        row?.__employeeLookupName ||
        row?.__employeelookupname ||
        ''
    ).trim();
  const findByName = (named) => {
    if (!named) return null;
    return sources.find((src) => {
      const srcName = rowName(src);
      return srcName && personNamesMatch(named, srcName);
    });
  };
  return list.map((row) => {
    if (!row || typeof row !== 'object') return row;
    const named = rowName(row);
    const src = findByName(named);
    const out = { ...row };
    const canCopyPayroll = src && (!hasAttendanceBand || formTSERowHasAttendanceMarks(row, hdrs));
    if (canCopyPayroll) {
      payrollHeaders.forEach((h) => {
        const v = getFormTSEKarnatakaRowValueForHeader(src, h);
        if (v == null || String(v).trim() === '') return;
        out[h] = v;
      });
    }
    return clearFormTSEKarnatakaPayrollIfNoAttendance(out, hdrs);
  });
}

export function applyFormTSEKarnatakaEmployeeToRow(row, emp, headers, helpers = {}) {
  if (!row || !emp || !Array.isArray(headers)) return row;
  const {
    sanitizeValue = (v) => String(v ?? '').trim(),
    rowIndex = 0,
    formatStatutoryDateDisplay = (v) => String(v ?? '').trim(),
    getEmployeeLookupName = null,
    getFallbackName = null,
    getFallbackFatherOrSpouse = null,
    onlyEmpty = false,
  } = helpers;

  const setCell = (header, value) => {
    if (!header) return;
    if (onlyEmpty) {
      const existing = String(getFormTSEKarnatakaRowValueForHeader(row, header) ?? '').trim();
      if (existing) {
        if (!isFormTSEEmployeeNameHeader(header)) return;
        const fatherHeader = headers.find((h) => isFormTSEFatherHusbandHeader(h));
        const fatherExisting = fatherHeader
          ? String(getFormTSEKarnatakaRowValueForHeader(row, fatherHeader) ?? '').trim()
          : '';
        const peopleFather = String(fatherName || '').trim();
        const leaked =
          (fatherExisting && existing.toLowerCase() === fatherExisting.toLowerCase()) ||
          (peopleFather && existing.toLowerCase() === peopleFather.toLowerCase());
        if (!leaked) return;
      }
    }
    row[header] = sanitizeValue(value);
  };

  const name =
    (typeof getFallbackName === 'function' && getFallbackName(emp)) ||
    formatFormTSEKarnatakaEmployeeName(emp) ||
    (typeof getEmployeeLookupName === 'function' && getEmployeeLookupName(emp)) ||
    '';

  const fatherName =
    (typeof getFallbackFatherOrSpouse === 'function' && getFallbackFatherOrSpouse(emp)) ||
    pickFormTSEEmployeeScalar(emp, [
      'Father_s_Name',
      'Father_s Name',
      'Father_Name',
      'Father Name',
      'Spouse_Name',
      'Spouse Name',
    ]);

  const genderRaw = pickFormTSEEmployeeScalar(emp, [
    'Sex',
    'Gender',
    'gender',
    'Gender.displayValue',
  ]);
  const gender =
    genderRaw &&
    (() => {
      const sl = genderRaw.toLowerCase();
      if (sl === 'm' || sl === 'male') return 'Male';
      if (sl === 'f' || sl === 'female') return 'Female';
      return genderRaw;
    })();

  const designation = pickFormTSEEmployeeScalar(emp, [
    'Designation',
    'designation',
    'Designation.displayValue',
    'Department',
    'department',
  ]);

  const dojRaw = pickFormTSEEmployeeScalar(emp, [
    'Dateofjoining',
    'DateofJoining',
    'Date of Joining',
    'Date_of_Joining',
    'Date_of_Joining',
  ]);
  const doj = dojRaw ? formatStatutoryDateDisplay(dojRaw) : '';

  const esi = pickFormTSEEmployeeScalar(emp, [
    'ESI_Number',
    'ESI Number',
    'ESIC_Number',
    'ESIC No.',
    "Employee's State Insurance Corporation No.",
  ]);

  const uan = pickFormTSEEmployeeScalar(emp, [
    'UAN_Number',
    'UAN Number',
    'UAN',
    'UAN No.',
    'UAN_Number',
  ]);

  headers.forEach((header) => {
    if (isFormTSESerialNumberHeader(header)) {
      setCell(header, String(rowIndex + 1));
    } else if (isFormTSEEmployeeNameHeader(header)) {
      setCell(header, name);
    } else if (isFormTSEFatherHusbandHeader(header)) {
      setCell(header, fatherName);
    } else if (isFormTSEGenderHeader(header)) {
      setCell(header, gender);
    } else if (isFormTSEDesignationHeader(header)) {
      setCell(header, designation);
    } else if (isFormTSEDateOfJoiningHeader(header)) {
      setCell(header, doj);
    } else if (isFormTSEEsicRegistrationHeader(header)) {
      setCell(header, esi);
    } else if (isFormTSEUanRegistrationHeader(header)) {
      setCell(header, uan);
    }
  });

  return row;
}

function resolveFormTHeaderExportValue(headerFormData, spec, parsedFields = []) {
  if (!headerFormData || typeof headerFormData !== 'object') return '';
  const tryKeys = (keys) => {
    for (const k of keys) {
      const v = headerFormData[k];
      if (v != null && String(v).trim() !== '') return String(v).trim();
    }
    return '';
  };
  if (spec.key === 'form_t_month_year') {
    const direct = tryKeys(['form_t_month_year', 'form_xviii_month_year', 'form_b_header_for_the_month_of']);
    if (direct) return direct;
  } else if (spec.key === 'form_t_establishment_name_address') {
    const direct = tryKeys([
      'form_t_establishment_name_address',
      'form_t_establishment',
      'statutory_establishment_name_address',
      'statutory_establishment_address',
      'statutory_establishment_name',
      'statutory_establishment_name_shop',
      'form_q_establishment',
      'form_f_establishment_name_address',
      'form_h_establishment_name_address',
      'form25_establishment',
      'form_xv_establishment_contract_carried',
      'form_xviii_establishment',
      // Legacy CLRA modal keys sometimes still hold the site establishment line.
      'form_xv_contractor',
      'form25_contractor',
      'form_xviii_contractor',
    ]);
    if (direct) return direct;
  } else if (spec.key === 'form_t_employer') {
    const direct = tryKeys([
      'form_t_employer',
      'form_q_ka_employer',
      'statutory_employer_name_address',
      'statutory_principal_employer',
      'form25_principal_employer',
      'form_xviii_principal_employer',
      'form_xvii_principal_employer',
      'form_xv_principal_employer',
    ]);
    if (direct) return direct;
  } else {
    const direct = tryKeys([spec.key]);
    if (direct) return direct;
  }
  for (const field of parsedFields) {
    if (!field?.key) continue;
    if (!formTLabelMatchesSpec(field.label, spec)) continue;
    const v = headerFormData[field.key];
    if (v != null && String(v).trim() !== '') return String(v).trim();
  }
  return '';
}

function formTLabelMatchesSpec(raw, spec) {
  const labelOnly = String(raw || '').split(':')[0].trim();
  if (!labelOnly) return false;
  // Row-3 CLRA leftover — not the Establishment / Employer value boxes on rows 10–11.
  if (/establ(?:ishment|ishemnt)\s+in\s*\/?\s*under\s+which\s+contract/i.test(labelOnly)) {
    return false;
  }
  if (spec.match.test(labelOnly)) return true;
  const norm = formTSEHeaderNorm(labelOnly);
  if (spec.key === 'form_t_month_year') return /^month\s+year$/.test(norm);
  if (spec.key === 'form_t_establishment_name_address') {
    return (
      /name\s+and\s+address\s+of\s+the\s+establishment/.test(norm) ||
      /^address\s+of\s+the\s+establishment$/.test(norm) ||
      /name\s+and\s+address\s+of\s+(?:the\s+)?contractor/.test(norm)
    );
  }
  if (spec.key === 'form_t_employer') {
    return (
      /name\s+and\s+address\s+of\s+(?:the\s+)?employer/.test(norm) ||
      /nature\s+of\s+work\s+and\s+location\s+of\s+work/.test(norm)
    );
  }
  return false;
}

function buildFormTSEMergeTopLeftResolver(worksheet) {
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

export function applyFormTSEKarnatakaAutofillFromSite(headerData, context = {}, options = {}) {
  const { onlyEmpty = false } = options;
  const { monthYearText = '', establishmentText = '', employerText = '' } = context;
  const out = { ...(headerData || {}) };
  const assign = (key, value) => {
    if (!key || value == null || String(value).trim() === '') return;
    if (onlyEmpty && String(out[key] ?? '').trim()) return;
    out[key] = String(value).trim();
  };
  assign('form_t_month_year', monthYearText);
  assign('form_t_establishment_name_address', establishmentText);
  assign('form_t_employer', employerText);

  // Lift from sibling statutory keys when site context did not supply a value.
  if (!String(out.form_t_establishment_name_address ?? '').trim()) {
    assign(
      'form_t_establishment_name_address',
      out.statutory_establishment_name_address ||
        out.form_q_establishment ||
        out.form25_establishment ||
        out.form_f_establishment_name_address ||
        out.form_xv_contractor ||
        out.form25_contractor ||
        ''
    );
  }
  if (!String(out.form_t_employer ?? '').trim()) {
    assign(
      'form_t_employer',
      out.statutory_employer_name_address ||
        out.statutory_principal_employer ||
        out.form_q_ka_employer ||
        out.form25_principal_employer ||
        out.form_xviii_principal_employer ||
        ''
    );
  }
  return out;
}

/** Merge modal / parsed header values into Form T keys before Excel export. */
export function prepareFormTSEDownloadHeaderData(headerFormData, parsedFormHeader = null, siteContext = {}) {
  const parsedFields = Array.isArray(parsedFormHeader?.fields) ? parsedFormHeader.fields : [];
  let out = headerFormData && typeof headerFormData === 'object' ? { ...headerFormData } : {};

  const copyFieldValue = (spec) => {
    const val = resolveFormTHeaderExportValue(out, spec, parsedFields);
    if (val) out[spec.key] = val;
  };
  FORM_T_KARNATAKA_HEADER_SPECS.forEach(copyFieldValue);

  for (const field of parsedFields) {
    if (!field?.key) continue;
    const v = out[field.key];
    if (v == null || String(v).trim() === '') continue;
    for (const spec of FORM_T_KARNATAKA_HEADER_SPECS) {
      if (formTLabelMatchesSpec(field.label, spec)) {
        if (!String(out[spec.key] ?? '').trim()) out[spec.key] = String(v).trim();
      }
    }
  }

  // Prefer values already on the Autofill form. Site context only fills blanks —
  // never overwrite Establishment / Employer the user can see in the modal.
  const preserved = {
    form_t_month_year: String(out.form_t_month_year ?? '').trim(),
    form_t_establishment_name_address: String(out.form_t_establishment_name_address ?? '').trim(),
    form_t_employer: String(out.form_t_employer ?? '').trim(),
  };
  out = applyFormTSEKarnatakaAutofillFromSite(out, siteContext, { onlyEmpty: true });
  if (preserved.form_t_month_year) out.form_t_month_year = preserved.form_t_month_year;
  if (preserved.form_t_establishment_name_address) {
    out.form_t_establishment_name_address = preserved.form_t_establishment_name_address;
  }
  if (preserved.form_t_employer) out.form_t_employer = preserved.form_t_employer;
  const establishment = String(out.form_t_establishment_name_address ?? '').trim();
  const employer = String(out.form_t_employer ?? '').trim();
  if (establishment && !String(out.statutory_establishment_name_address ?? '').trim()) {
    out.statutory_establishment_name_address = establishment;
  }
  if (employer && !String(out.statutory_employer_name_address ?? '').trim()) {
    out.statutory_employer_name_address = employer;
  }
  return out;
}

function fieldLabelMatchesCell(label, cellText) {
  const labelText = String(label || '').trim().replace(/:+$/, '');
  const cellStr = String(cellText || '').split(':')[0].trim().replace(/:+$/, '');
  if (!labelText || !cellStr) return false;
  if (cellStr === labelText || cellStr.startsWith(labelText) || labelText.startsWith(cellStr)) {
    return true;
  }
  const labelNorm = formTSEHeaderNorm(labelText);
  const cellNorm = formTSEHeaderNorm(cellStr);
  return labelNorm === cellNorm || labelNorm.includes(cellNorm) || cellNorm.includes(labelNorm);
}

const FORM_T_KARNATAKA_VALUE_COL = 2;
const FORM_T_KARNATAKA_FALLBACK_ROWS = [9, 10, 11];

function readFormTSEMergedCellText(worksheet, getMergeTopLeft, row, col) {
  const tl = getMergeTopLeft(row, col);
  return excelCellValueToString(worksheet.getCell(tl.r, tl.c)?.value).trim();
}

/**
 * Widen Form T header value boxes (rows 9–11) across A–I so Month/Year, Address,
 * and Employer text is not clipped in the narrow A:D template merges.
 */
export function expandFormTSEHeaderValueBoxes(
  worksheet,
  endCol = FORM_T_KARNATAKA_HEADER_BOX_END_COL
) {
  if (!worksheet) return;
  const mergeEnd = Math.max(2, Number(endCol) || FORM_T_KARNATAKA_HEADER_BOX_END_COL);
  const headerRows = FORM_T_KARNATAKA_FALLBACK_ROWS;

  headerRows.forEach((row) => {
    const merges = Array.isArray(worksheet.model?.merges) ? [...worksheet.model.merges] : [];
    merges.forEach((label) => {
      const parts = String(label || '').split(':');
      if (parts.length !== 2) return;
      const tl = worksheet.getCell(parts[0]);
      const br = worksheet.getCell(parts[1]);
      if (!tl || !br) return;
      // Only touch the Form T header band (rows 9–11), never table merges.
      if (tl.row < 9 || br.row > 11) return;
      if (row < tl.row || row > br.row) return;
      if (br.col < 1 || tl.col > mergeEnd) return;
      try {
        worksheet.unMergeCells(label);
      } catch (_) {
        /* ignore */
      }
    });

    try {
      worksheet.mergeCells(row, 1, row, mergeEnd);
    } catch (_) {
      /* template may already span this far */
    }

    const cell = worksheet.getCell(row, 1);
    const text = excelCellValueToString(cell.value).trim();
    cell.alignment = {
      ...(cell.alignment || {}),
      wrapText: true,
      vertical: 'middle',
      horizontal: 'left',
    };
    const wsRow = worksheet.getRow(row);
    if (wsRow) {
      // Long establishment / employer addresses need taller header boxes.
      const needed = text.length > 160 ? 55 : text.length > 100 ? 42 : text.length > 50 ? 32 : 22;
      wsRow.height = Math.max(Number(wsRow.height) || 0, needed);
    }
  });
  clearFormTSEKarnatakaHeaderBandBorders(worksheet, {
    colTo: mergeEnd,
    remergeHeaderBoxes: true,
  });
}

/** Rename legacy CLRA header labels on rows 9–11 to the official Form T KA wording. */
export function renameFormTSEKarnatakaHeaderLabels(worksheet) {
  if (!worksheet) return false;
  let touched = false;
  const getMergeTopLeft = buildFormTSEMergeTopLeftResolver(worksheet);
  for (let r = 9; r <= 11; r += 1) {
    for (let c = 1; c <= 8; c += 1) {
      const tl = getMergeTopLeft(r, c);
      if (tl.r !== r || tl.c !== c) continue;
      const raw = readFormTSEMergedCellText(worksheet, getMergeTopLeft, r, c);
      if (!raw) continue;
      const parts = raw.split(':');
      const labelOnly = parts[0].trim();
      const valuePart = parts.length > 1 ? parts.slice(1).join(':').trim() : '';
      let nextLabel = '';
      if (/name\s+and\s+address\s+of\s+(?:the\s+)?contractor/i.test(labelOnly)) {
        nextLabel = 'Name and address of the Establishment';
      } else if (/nature\s+of\s+work\s+and\s+location\s+of\s+work/i.test(labelOnly)) {
        nextLabel = 'Name and Address of employer';
      } else if (
        /address\s+of\s+the\s+establishment/i.test(labelOnly) &&
        !/name\s+and\s+address\s+of\s+the\s+establishment/i.test(labelOnly)
      ) {
        nextLabel = 'Name and address of the Establishment';
      }
      if (!nextLabel) continue;
      const cell = worksheet.getCell(tl.r, tl.c);
      cell.value = valuePart ? `${nextLabel} : ${valuePart}` : nextLabel;
      cell.alignment = {
        ...(cell.alignment || {}),
        wrapText: true,
        vertical: 'middle',
        horizontal: 'left',
      };
      touched = true;
      break;
    }
  }
  return touched;
}

/** Write Form T header values into the header boxes (rows 9–11), Label : value in the widened merge. */
export function writeFormTSEHeaderFieldsToWorksheet(worksheet, headerFormData = {}, parsedFormHeader = null) {
  if (!worksheet || !headerFormData || typeof headerFormData !== 'object') return;
  const parsedFields = Array.isArray(parsedFormHeader?.fields) ? parsedFormHeader.fields : [];
  // Widen boxes first so Label : value is not clipped in A:D.
  expandFormTSEHeaderValueBoxes(worksheet, FORM_T_KARNATAKA_HEADER_BOX_END_COL);
  renameFormTSEKarnatakaHeaderLabels(worksheet);
  const getMergeTopLeft = buildFormTSEMergeTopLeftResolver(worksheet);

  const styleHeaderCell = (cell, text = '') => {
    if (!cell) return;
    cell.alignment = {
      ...(cell.alignment || {}),
      wrapText: true,
      vertical: 'middle',
      horizontal: 'left',
    };
    cell.font = {
      ...(cell.font || {}),
      bold: true,
      size: FORM_T_KA_HEADER_IDENTITY_FONT_SIZE,
    };
    try {
      const r = Number(cell.row) || 0;
      if (r >= 9 && r <= 11) {
        const wsRow = worksheet.getRow(r);
        const len = String(text || '').length;
        const needed = len > 100 ? 40 : len > 50 ? 30 : 20;
        if (wsRow) wsRow.height = Math.max(Number(wsRow.height) || 0, needed);
      }
    } catch (_) {
      /* ignore */
    }
  };

  const writeAt = (row, col, value, preferredLabel = '') => {
    const text = String(value ?? '').trim();
    if (!text || row < 1 || col < 1) return;
    // Always stamp Label : value on the row-A merge master (A–I). Do not rely on adjacent B
    // cells — when the A:I merge is missing, B-only writes leave column A looking empty.
    try {
      worksheet.mergeCells(row, 1, row, FORM_T_KARNATAKA_HEADER_BOX_END_COL);
    } catch (_) {
      /* already merged / overlap */
    }
    const existing = excelCellValueToString(worksheet.getCell(row, 1)?.value).trim();
    const existingLabel = existing.split(':')[0].trim();
    const labelOnly = preferredLabel || existingLabel || '';
    const combined = labelOnly ? `${labelOnly} : ${text}` : text;
    const cell = worksheet.getCell(row, 1);
    cell.value = combined;
    styleHeaderCell(cell, combined);
  };

  const resolveValueCol = (parsedField, labelCol1Based = 1) => {
    if (parsedField?.valueCol != null && Number(parsedField.valueCol) >= 0) {
      return Number(parsedField.valueCol) + 1;
    }
    return Math.max(labelCol1Based + 1, FORM_T_KARNATAKA_VALUE_COL);
  };

  FORM_T_KARNATAKA_HEADER_SPECS.forEach((spec, specIndex) => {
    const val = resolveFormTHeaderExportValue(headerFormData, spec, parsedFields);
    if (!val) return;

    const parsedField = parsedFields.find(
      (field) => field?.key === spec.key || formTLabelMatchesSpec(field?.label, spec)
    );
    // Form T value boxes are only rows 9–11. Never write into the title/rule band (1–8).
    if (parsedField?.labelRow != null) {
      const valueRow =
        parsedField.valueRow != null ? Number(parsedField.valueRow) + 1 : Number(parsedField.labelRow) + 1;
      if (valueRow >= 9 && valueRow <= 11) {
        writeAt(valueRow, resolveValueCol(parsedField, (parsedField.labelCol ?? 0) + 1), val, spec.label);
      }
    } else {
      let matched = false;
      for (let r = 9; r <= 11; r += 1) {
        for (let c = 1; c <= 8; c += 1) {
          const tl = getMergeTopLeft(r, c);
          if (tl.r !== r || tl.c !== c) continue;
          const raw = readFormTSEMergedCellText(worksheet, getMergeTopLeft, r, c);
          if (!raw) continue;
          const rawLabel = raw.split(':')[0].trim();
          if (!formTLabelMatchesSpec(rawLabel, spec) && !fieldLabelMatchesCell(spec.label, rawLabel)) {
            continue;
          }
          writeAt(r, resolveValueCol(parsedField, c), val, spec.label);
          matched = true;
          break;
        }
        if (matched) break;
      }
    }

    // Always re-stamp the canonical fallback row so Establishment/Employer cannot stay label-only.
    const fallbackRow = FORM_T_KARNATAKA_FALLBACK_ROWS[specIndex];
    if (fallbackRow) {
      writeAt(fallbackRow, FORM_T_KARNATAKA_VALUE_COL, val, spec.label);
    }
  });
}

function normCell(txt) {
  return String(txt || '').replace(/\s+/g, ' ').trim();
}

function parseCalendarDay(text) {
  const t = normCell(text).replace(/[()]/g, '');
  if (!/^\d{1,2}$/.test(t)) return 0;
  const n = parseInt(t, 10);
  return n >= 1 && n <= 31 ? n : 0;
}

function parseStatutoryColumnNumber(text) {
  const t = normCell(text).replace(/[()]/g, '');
  if (!/^\d{1,2}$/.test(t)) return 0;
  const n = parseInt(t, 10);
  return n >= 1 && n <= 99 ? n : 0;
}

/** Form T statutory column index row: 1, 2, 3 … or 11, 12, 13 … under wage columns — not calendar days. */
function rowLooksLikeFormTColumnIndexRow(rowIndex, startCol, maxCol, getMergedAwareCellText, readRaw) {
  let total = 0;
  let sequential = 0;
  let numeric = 0;
  let ascending = 0;
  let lastNum = 0;
  let leadingSequential = 0;
  const leadN = Math.min(8, Math.max(0, maxCol - startCol));
  for (let i = 0; i < Math.min(50, maxCol - startCol); i += 1) {
    const c = startCol + i;
    const t = normCell(readRaw(rowIndex, c) || getMergedAwareCellText(rowIndex, c));
    if (!t) continue;
    total += 1;
    const n = parseStatutoryColumnNumber(t);
    if (n > 0) {
      numeric += 1;
      if (n === i + 1) sequential += 1;
      if (lastNum === 0 || n === lastNum + 1) ascending += 1;
      lastNum = n;
      if (i < leadN && n === i + 1) leadingSequential += 1;
    } else if (i < leadN) {
      // Identity columns on the day-number strip are text (S.NO, Name…) — not an index row.
      leadingSequential = 0;
    }
  }
  if (total < 4) return false;
  // Calendar day strip has 1..31 only under ATTENDANCE; statutory index has 1,2,3… from the left.
  if (leadingSequential < 4) return false;
  return (
    sequential >= Math.max(3, Math.floor(total * 0.55)) ||
    ascending >= Math.max(4, Math.floor(total * 0.65)) ||
    numeric / total >= 0.85
  );
}

function columnHasAttendanceBanner(mainRow, endRow, col, getMergedAwareCellText) {
  for (let r = mainRow; r <= endRow; r += 1) {
    const t = String(getMergedAwareCellText(r, col) || '').toLowerCase();
    if (t.includes('attendance')) return true;
  }
  return false;
}

function pickEmployeeColumnHeader(mainRow, col, getMergedAwareCellText, readRaw) {
  const t = normCell(readRaw(mainRow, col) || getMergedAwareCellText(mainRow, col));
  if (!t || /^\d+$/.test(t)) return '';
  if (isFormTWageParentBannerText(t)) return '';
  if (/attendance/i.test(t) && t.length > 10) return '';
  return t;
}

function isFormTWageParentBannerText(raw) {
  const t = normCell(raw).toLowerCase();
  if (!t) return false;
  if (t === 'earned' || /^earned\s+wages/.test(t) || /^wages?\s+earned/.test(t)) return true;
  if (t.includes('earned') && t.includes('allowance')) return true;
  if (t.includes('deduction')) return true;
  if (/^attendance\s+and/.test(t)) return true;
  return false;
}

/** Leaf wage label (Basic, HRA, No. of payable days) — skip merged parent banners and column numbers. */
function pickFormTWageColumnHeader(topRow, leafRow, col, getMergedAwareCellText, readRaw) {
  const endRow = Math.max(topRow, leafRow);
  for (let r = endRow; r >= topRow; r -= 1) {
    const raw = normCell(readRaw(r, col) || getMergedAwareCellText(r, col));
    if (!raw || /^\d{1,2}$/.test(raw)) continue;
    if (isFormTWageParentBannerText(raw)) continue;
    if (/attendance/i.test(raw) && raw.length > 14) continue;
    return raw;
  }
  return '';
}

function findFormTWageLeafRow(mainRow, scanEnd, startCol, maxCol, getMergedAwareCellText, readRaw, indexRow) {
  const leafHitRe =
    /basic|hra|payable|conveyance|medical|allowance|overtime|\bot\b|da\/vda|\bvda\b|special|nfh|maternity|subsist|attendance\s*bonus|net\s*amount|mode\s*of\s*payment|\bfines?\b|damage|salary\s*adv/;
  const scoreLeafRow = (r) => {
    let hits = 0;
    for (let c = startCol; c < maxCol; c += 1) {
      const raw = normCell(readRaw(r, c) || getMergedAwareCellText(r, c));
      if (!raw || isFormTWageParentBannerText(raw)) continue;
      if (leafHitRe.test(raw.toLowerCase())) hits += 1;
    }
    return hits;
  };
  if (indexRow > mainRow) {
    const aboveIndex = indexRow - 1;
    if (aboveIndex >= mainRow) {
      const hits = scoreLeafRow(aboveIndex);
      if (hits >= 2) return aboveIndex;
    }
  }
  let bestRow = mainRow;
  let bestHits = 0;
  for (let r = mainRow; r <= scanEnd; r += 1) {
    if (r === indexRow) continue;
    const hits = scoreLeafRow(r);
    if (hits > bestHits) {
      bestHits = hits;
      bestRow = r;
    }
  }
  return bestHits >= 2 ? bestRow : Math.max(mainRow, scanEnd);
}

function appendFormTStatutoryColumnSuffix(header, statCol) {
  const base = normCell(header);
  if (!base || !statCol) return base;
  if (new RegExp(`\\(\\s*${statCol}\\s*\\)`).test(base)) return base;
  return `${base} (${statCol})`;
}

function rowHasSampleData(jsonData, col, fromRow, maxRow) {
  const end = Math.min(maxRow, jsonData.length);
  for (let r = fromRow; r < end; r += 1) {
    if (normCell((jsonData[r] || [])[col])) return true;
  }
  return false;
}

function scoreAttendanceCalendarDayRow(rowIndex, mainRow, scanEnd, startCol, maxCol, getMergedAwareCellText, readRaw) {
  if (rowLooksLikeFormTColumnIndexRow(rowIndex, startCol, maxCol, getMergedAwareCellText, readRaw)) {
    return 0;
  }
  let hits = 0;
  for (let c = startCol; c < maxCol; c += 1) {
    if (!columnHasAttendanceBanner(mainRow, scanEnd, c, getMergedAwareCellText)) continue;
    const day = parseCalendarDay(readRaw(rowIndex, c) || getMergedAwareCellText(rowIndex, c));
    if (day > 0) hits += 1;
  }
  return hits;
}

function findAttendanceCalendarDayRow(mainRow, scanEnd, startCol, maxCol, getMergedAwareCellText, readRaw) {
  let bestDayRow = -1;
  let bestDayHits = 0;
  for (let r = mainRow + 1; r <= scanEnd; r += 1) {
    const hits = scoreAttendanceCalendarDayRow(
      r,
      mainRow,
      scanEnd,
      startCol,
      maxCol,
      getMergedAwareCellText,
      readRaw
    );
    if (hits > bestDayHits) {
      bestDayHits = hits;
      bestDayRow = r;
    }
  }
  return { bestDayRow, bestDayHits };
}

function findFormTColumnIndexRow(mainRow, scanEnd, startCol, maxCol, getMergedAwareCellText, readRaw) {
  for (let r = mainRow + 1; r <= scanEnd; r += 1) {
    if (rowLooksLikeFormTColumnIndexRow(r, startCol, maxCol, getMergedAwareCellText, readRaw)) {
      return r;
    }
  }
  return -1;
}

/** True only for ATTENDANCE_1 … ATTENDANCE_31 — never plain "1" (Form T column index). */
export function isFormTSEAttendanceDayHeader(header) {
  const raw = String(header || '').trim();
  const m = raw.match(/^ATTENDANCE_(\d{1,2})(?:\s*\(\s*\d{1,2}\s*\))?$/i);
  if (!m) return false;
  const day = parseInt(m[1], 10);
  return day >= 1 && day <= 31;
}

/** Autofill attendance marks only — reject payroll leakage (31, NIL, 18563). */
export function sanitizeFormTSEAttendanceMark(value) {
  const t = String(value ?? '').trim();
  if (!t) return '';
  if (/^enter\s*\d+/i.test(t)) return '';
  if (/^(nil|n\/a)$/i.test(t)) return '';
  if (/^-?\d+(\.\d+)?$/.test(t)) return '';
  return t;
}

export function listFormTSEAttendanceDayHeaders(headers) {
  const out = [];
  (Array.isArray(headers) ? headers : []).forEach((header) => {
    if (!isFormTSEAttendanceDayHeader(header)) return;
    const day = parseInt(String(header).match(/^ATTENDANCE_(\d{1,2})/i)[1], 10);
    out.push({ header, day });
  });
  return out.sort((a, b) => a.day - b.day);
}

/** True when any ATTENDANCE_1…31 cell has a real mark (P, WO, CL, …). */
export function formTSERowHasAttendanceMarks(row, headers = []) {
  if (!row || typeof row !== 'object') return false;
  const hdrs = Array.isArray(headers) && headers.length
    ? headers
    : Object.keys(row).filter((k) => !String(k).startsWith('__'));
  const dayHeaders = listFormTSEAttendanceDayHeaders(hdrs);
  if (dayHeaders.length) {
    return dayHeaders.some(({ header }) =>
      Boolean(sanitizeFormTSEAttendanceMark(getFormTSEKarnatakaRowValueForHeader(row, header)))
    );
  }
  return Object.keys(row).some(
    (k) => isFormTSEAttendanceDayHeader(k) && Boolean(sanitizeFormTSEAttendanceMark(row[k]))
  );
}

function isFormTSEKarnatakaKeptPayrollHeaderWhenAbsent(header) {
  if (isFormTSEKarnatakaTotalOtHoursHeader(header)) return true;
  const s = formTSEEmployeeHeaderKeyNorm(header);
  return s.includes('mode') && s.includes('payment');
}

/**
 * If the person did not attend this month, do not show payroll register
 * amounts (payable days, Basic, HRA, Total, PF, PT, …). Keep OT NIL.
 */
export function clearFormTSEKarnatakaPayrollIfNoAttendance(row, headers = []) {
  if (!row || typeof row !== 'object') return row;
  const hdrs = Array.isArray(headers) && headers.length
    ? headers
    : Object.keys(row).filter((k) => !String(k).startsWith('__'));
  if (!listFormTSEAttendanceDayHeaders(hdrs).length) return row;
  if (formTSERowHasAttendanceMarks(row, hdrs)) return row;
  hdrs.forEach((h) => {
    if (!isFormTSEKarnatakaPayrollRegisterHeader(h, hdrs)) return;
    if (isFormTSEKarnatakaKeptPayrollHeaderWhenAbsent(h)) return;
    row[h] = '';
  });
  return row;
}

export function applyFormTSEKarnatakaPayrollOnlyWhenAttended(mappedData, headers = []) {
  if (!Array.isArray(mappedData)) return [];
  return mappedData.map((row) => {
    if (!row || typeof row !== 'object') return row;
    const out = { ...row };
    clearFormTSEKarnatakaPayrollIfNoAttendance(out, headers);
    return out;
  });
}

/** Consecutive ATTENDANCE_1 … band inside parsed table headers (modal day grid). */
export function getFormTSEAttendanceGridGroupInfo(tableHeaders) {
  const list = Array.isArray(tableHeaders) ? tableHeaders : [];
  for (let i = 0; i < list.length; i += 1) {
    if (!isFormTSEAttendanceDayHeader(list[i])) continue;
    let j = i + 1;
    while (j < list.length && isFormTSEAttendanceDayHeader(list[j])) j += 1;
    return { startIndex: i, dayCount: j - i, parentLabel: 'ATTENDANCE' };
  }
  return null;
}

/** Hide calendar days above the selected month length (April → 30 columns, not 31). */
export function filterFormTSETableHeadersForMonth(headers, monthDayCount = 31) {
  const days = Math.min(Math.max(Number(monthDayCount) || 31, 28), 31);
  return (Array.isArray(headers) ? headers : []).filter((header) => {
    if (!isFormTSEAttendanceDayHeader(header)) return true;
    const day = parseInt(String(header).match(/^ATTENDANCE_(\d{1,2})/i)[1], 10);
    return day >= 1 && day <= days;
  });
}

/** Keep group labels aligned when attendance day columns are removed for shorter months. */
export function filterFormTSETableHeadersAndGroupsForMonth(headers, groupLabels, monthDayCount = 31) {
  const list = Array.isArray(headers) ? headers : [];
  const groups = Array.isArray(groupLabels) ? groupLabels : [];
  const days = Math.min(Math.max(Number(monthDayCount) || 31, 28), 31);
  const outHeaders = [];
  const outGroups = [];
  list.forEach((header, index) => {
    if (isFormTSEAttendanceDayHeader(header)) {
      const day = parseInt(String(header).match(/^ATTENDANCE_(\d{1,2})/i)[1], 10);
      if (day < 1 || day > days) return;
    }
    outHeaders.push(header);
    outGroups.push(groups.length === list.length ? groups[index] : '');
  });
  return { headers: outHeaders, groupLabels: outGroups };
}

/**
 * When modal headers omit day 31, map each displayed column to its original Excel index
 * so merged wage bands (Deductions, etc.) stay aligned with body cells.
 */
export function mapFormTSEDisplayHeadersToExcelCols(parsedHeaders, displayHeaders, tableStartCol = 0) {
  const parsed = Array.isArray(parsedHeaders) ? parsedHeaders : [];
  const display = Array.isArray(displayHeaders) ? displayHeaders : [];
  const start = Math.max(0, Number(tableStartCol) || 0);
  if (!parsed.length || !display.length || display.length > parsed.length) return null;
  if (display.length === parsed.length) return null;

  const norm = (h) =>
    String(h || '')
      .trim()
      .toLowerCase()
      .replace(/\s*\(\s*\d{1,2}\s*\)\s*$/, '');
  const cols = [];
  let pi = 0;
  for (let di = 0; di < display.length; di += 1) {
    const header = display[di];
    const target = norm(header);
    while (pi < parsed.length) {
      const p = parsed[pi];
      if (p === header || norm(p) === target) break;
      if (isFormTSEAttendanceDayHeader(p)) {
        pi += 1;
        continue;
      }
      const directIdx = parsed.indexOf(header, pi);
      if (directIdx >= 0) {
        pi = directIdx;
        break;
      }
      const normIdx = parsed.findIndex((ph, idx) => idx >= pi && norm(ph) === target);
      if (normIdx >= 0) {
        pi = normIdx;
        break;
      }
      return null;
    }
    if (pi >= parsed.length) return null;
    cols.push(start + pi);
    pi += 1;
  }
  return cols.length === display.length ? cols : null;
}

/** Parent banner labels (Earned wages, Deductions, …) aligned to filtered display headers. */
export function buildFormTSEGroupLabelsFromWorksheet(
  worksheet,
  {
    displayHeaders = [],
    parsedHeaders = [],
    tableStartCol = 0,
    headerRowIndex = -1,
    dataStartIndex = -1,
  } = {}
) {
  if (!worksheet || !Array.isArray(displayHeaders) || displayHeaders.length === 0) return null;
  const excelCols =
    resolveFormTSEVisibleModalExcelCols(parsedHeaders, displayHeaders, tableStartCol) ||
    [];
  if (!excelCols.length) return null;
  const merges = worksheet['!merges'] || [];
  const rawCell = (r, c) => {
    if (r < 0 || c < 0) return '';
    const ref = XLSX.utils.encode_cell({ r, c });
    const cell = worksheet[ref];
    return cell && cell.v != null ? String(cell.v).trim() : '';
  };
  const mergedAware = (r, c) => {
    const direct = rawCell(r, c);
    if (direct) return direct;
    for (let i = 0; i < merges.length; i += 1) {
      const m = merges[i];
      if (!m?.s || !m?.e) continue;
      if (r >= m.s.r && r <= m.e.r && c >= m.s.c && c <= m.e.c) {
        const topLeft = rawCell(m.s.r, m.s.c);
        if (topLeft) return topLeft;
      }
    }
    return '';
  };
  const norm = (txt) =>
    String(txt || '')
      .replace(/\r?\n/g, ' ')
      .replace(/\s+/g, ' ')
      .trim()
      .toLowerCase();
  const ds = Number(dataStartIndex);
  const leafRow =
    ds > 0 ? ds - 1 : Number(headerRowIndex) >= 0 ? Number(headerRowIndex) + 1 : -1;
  const topRow = Number(headerRowIndex) >= 0 ? Number(headerRowIndex) : Math.max(0, leafRow - 2);
  if (leafRow < 0) return null;

  return displayHeaders.map((header, i) => {
    if (isFormTSEAttendanceDayHeader(header)) return '';
    const ec = excelCols[i];
    const leaf = norm(mergedAware(leafRow, ec));
    for (let r = topRow; r < leafRow; r += 1) {
      const parent = String(mergedAware(r, ec) || '')
        .replace(/\r?\n/g, ' ')
        .replace(/\s+/g, ' ')
        .trim();
      if (!parent || /^\d{1,2}$/.test(parent)) continue;
      if (norm(parent) === leaf) continue;
      if (/attendance/i.test(parent) && parent.length > 12) continue;
      return parent;
    }
    return '';
  });
}

/** Visible modal columns for the selected month (identity + attendance days + wage cols). */
export function resolveFormTSEVisibleModalHeaders(parsedHeaders, monthDayCount = 31) {
  return filterFormTSETableHeadersForMonth(parsedHeaders, monthDayCount);
}

export function resolveFormTSEVisibleModalGroupLabels(parsedHeaders, groupLabels, monthDayCount = 31) {
  return filterFormTSETableHeadersAndGroupsForMonth(parsedHeaders, groupLabels, monthDayCount).groupLabels;
}

/** Excel column index per visible header — never assume contiguous columns after day 31 is removed. */
export function resolveFormTSEVisibleModalExcelCols(parsedHeaders, visibleHeaders, tableStartCol = 0) {
  const parsed = Array.isArray(parsedHeaders) ? parsedHeaders : [];
  const visible = Array.isArray(visibleHeaders) ? visibleHeaders : [];
  const start = Math.max(0, Number(tableStartCol) || 0);
  if (!parsed.length || !visible.length) return null;

  const mapped = mapFormTSEDisplayHeadersToExcelCols(parsed, visible, start);
  if (mapped?.length === visible.length) return mapped;

  const norm = (h) =>
    String(h || '')
      .trim()
      .toLowerCase()
      .replace(/\s*\(\s*\d{1,2}\s*\)\s*$/, '');
  const cols = [];
  let pi = 0;
  for (let vi = 0; vi < visible.length; vi += 1) {
    const header = visible[vi];
    const target = norm(header);
    while (pi < parsed.length) {
      const p = parsed[pi];
      if (p === header || norm(p) === target) break;
      if (isFormTSEAttendanceDayHeader(p)) {
        pi += 1;
        continue;
      }
      const directIdx = parsed.indexOf(header, pi);
      if (directIdx >= 0) {
        pi = directIdx;
        break;
      }
      const normIdx = parsed.findIndex((ph, idx) => idx >= pi && norm(ph) === target);
      if (normIdx >= 0) {
        pi = normIdx;
        break;
      }
      return null;
    }
    if (pi >= parsed.length) return null;
    cols.push(start + pi);
    pi += 1;
  }
  return cols.length === visible.length ? cols : null;
}

/** Strip statutory column index suffix e.g. "Basic (12)" → "Basic" for modal labels. */
export function formatFormTSEWageColumnHeaderLabel(header) {
  const raw = String(header || '').trim();
  if (!raw) return raw;
  if (isFormTSEAttendanceDayHeader(raw)) {
    const day = parseInt(String(raw).match(/^ATTENDANCE_(\d{1,2})/i)[1], 10);
    return day >= 1 ? String(day) : raw;
  }
  return raw.replace(/\s*\(\s*\d{1,2}\s*\)\s*$/, '').trim() || raw;
}

/** Fallback parent band when Excel group labels are missing (Deductions / Earned wages). */
export function inferFormTSEWageGroupLabelFromHeader(header) {
  const h = formatFormTSEWageColumnHeaderLabel(header).toLowerCase();
  if (
    /salary advance|salary advances|\bfines\b|damage|other deduction|total deduction|\besi\b|\bpf\b|\bpt\b|\btds\b|society|insurance/.test(
      h
    )
  ) {
    return 'Deductions';
  }
  if (
    /^basic$|da\/vda|\bhra\b|conveyance|medical allowance|attendance bonus|special allowance|^ot$|\bnfh\b|maternity|subsistence|^others?$|^total$|earned wages|earnings?\s*total/.test(
      h
    )
  ) {
    return 'Earned wages and other allowances';
  }
  return '';
}

/**
 * Re-read Form T headers: employee columns keep text labels; only the ATTENDANCE band
 * gets ATTENDANCE_1 … ATTENDANCE_31 keys (never the statutory column-index row 1…9).
 */
export function rebuildFormTSETableHeadersFromSheet({
  headerRowIndex = -1,
  getMergedAwareCellText,
  getRawCellText = null,
  jsonData,
  effectiveSheetCols = 40,
  tableStartCol = 0,
  dataStartIndex = -1,
}) {
  const readRaw = typeof getRawCellText === 'function' ? getRawCellText : getMergedAwareCellText;
  if (!Array.isArray(jsonData) || jsonData.length === 0 || headerRowIndex < 0) return null;

  const mainRow = headerRowIndex;
  // Always scan far enough to find calendar-day + statutory index strips.
  // A premature parsedDataStartIndex often points AT the index row (1,2,3…) — if we
  // cut the scan to dataStartIndex-1 we miss that strip and then overwrite it with employees.
  const scanEnd = Math.min(
    jsonData.length - 1,
    Math.max(
      mainRow + 4,
      dataStartIndex > mainRow ? dataStartIndex + 1 : mainRow + 12,
      mainRow + 12
    )
  );
  // Form T: A–I identity + 31 attendance days + wage/deduction cols (~11–38) ≈ 70+ columns.
  const maxCol = Math.max(Number(effectiveSheetCols) || 0, tableStartCol + 80, 80);

  let startCol = Math.max(0, Number(tableStartCol) || 0);
  for (let c = 0; c < Math.min(maxCol, 14); c += 1) {
    const t = String(getMergedAwareCellText(mainRow, c) || '').toLowerCase();
    // Match "S.NO" / "S. No" / "Sl. No" / "Sr No" / "Serial No" (Form T uses "S.NO").
    if (/s\.?\s*no|sl\.?\s*no|sr\.?\s*no|serial/.test(t) || isFormTSESerialNumberHeader(t)) {
      startCol = c;
      break;
    }
  }

  const indexRow = findFormTColumnIndexRow(mainRow, scanEnd, startCol, maxCol, getMergedAwareCellText, readRaw);
  const wageLeafRow = findFormTWageLeafRow(
    mainRow,
    scanEnd,
    startCol,
    maxCol,
    getMergedAwareCellText,
    readRaw,
    indexRow
  );
  const { bestDayRow, bestDayHits } = findAttendanceCalendarDayRow(
    mainRow,
    scanEnd,
    startCol,
    maxCol,
    getMergedAwareCellText,
    readRaw
  );

  const dataStartRow = Math.max(mainRow, indexRow, bestDayRow) + 1;

  const columnHasHeaderSignal = (col) => {
    if (columnHasAttendanceBanner(mainRow, scanEnd, col, getMergedAwareCellText)) return true;
    if (pickFormTWageColumnHeader(mainRow, wageLeafRow, col, getMergedAwareCellText, readRaw)) {
      return true;
    }
    if (pickEmployeeColumnHeader(mainRow, col, getMergedAwareCellText, readRaw)) return true;
    if (indexRow >= 0) {
      const statCol = parseStatutoryColumnNumber(
        readRaw(indexRow, col) || getMergedAwareCellText(indexRow, col)
      );
      if (statCol > 0) return true;
    }
    return false;
  };

  const headers = [];
  for (let c = startCol; c < maxCol; c += 1) {
    let header = '';
    const inAttendanceBand = columnHasAttendanceBanner(mainRow, scanEnd, c, getMergedAwareCellText);
    if (inAttendanceBand) {
      let day = 0;
      if (bestDayRow >= 0 && bestDayHits >= 3) {
        day = parseCalendarDay(readRaw(bestDayRow, c) || getMergedAwareCellText(bestDayRow, c));
      } else if (indexRow >= 0) {
        day = parseCalendarDay(readRaw(indexRow, c) || getMergedAwareCellText(indexRow, c));
      }
      if (day > 0) header = `ATTENDANCE_${day}`;
    }
    if (!header) {
      header = pickFormTWageColumnHeader(mainRow, wageLeafRow, c, getMergedAwareCellText, readRaw);
    }
    if (!header) {
      header = pickEmployeeColumnHeader(mainRow, c, getMergedAwareCellText, readRaw);
    }
    let statCol = 0;
    if (indexRow >= 0 && !/^ATTENDANCE_\d{1,2}$/i.test(header)) {
      statCol = parseStatutoryColumnNumber(
        readRaw(indexRow, c) || getMergedAwareCellText(indexRow, c)
      );
      if (header && statCol > 0) header = appendFormTStatutoryColumnSuffix(header, statCol);
    }
    // Keep wage/deduction columns even when the leaf label cell is empty under a merge —
    // statutory index row (11…38) is enough to reserve the column.
    if (!header && statCol > 0) {
      header = `Column ${statCol} (${statCol})`;
    }
    const hasSample = rowHasSampleData(jsonData, c, dataStartRow, dataStartRow + 15);
    if (!header && !hasSample && headers.length > 0) {
      let anyMore = false;
      for (let cc = c; cc < Math.min(c + 16, maxCol); cc += 1) {
        if (rowHasSampleData(jsonData, cc, dataStartRow, dataStartRow + 10) || columnHasHeaderSignal(cc)) {
          anyMore = true;
          break;
        }
      }
      if (!anyMore) break;
    }
    headers.push(header || (headers.length ? `Column ${headers.length + 1}` : 'Sl. No.'));
  }

  while (headers.length > 0 && !normCell(headers[headers.length - 1])) headers.pop();

  const attendanceCols = listFormTSEAttendanceDayHeaders(headers).length;
  const employeeCols = headers.filter(
    (h) => normCell(h) && !isFormTSEAttendanceDayHeader(h)
  ).length;
  const hasWageColumnMarkers = headers.some((h) =>
    /\(\s*1[1-9]\s*\)|\(\s*2[0-9]\s*\)|\(\s*3[0-8]\s*\)/.test(String(h || ''))
  );
  if (headers.length < 4 || employeeCols < 2) return null;
  if (attendanceCols < 3 && bestDayHits < 3 && !hasWageColumnMarkers) return null;

  return {
    headers,
    headerRowIndex: mainRow,
    dataStartIndex: dataStartRow,
    tableStartCol: startCol,
    attendanceDayRow: bestDayRow,
    attendanceColumnCount: attendanceCols,
    columnIndexRow: indexRow,
  };
}

/** When table headers are rebuilt, copy cell values by column index and matching label. */
export function remapRowsToRebuiltTableHeaders(rows, oldHeaders, newHeaders) {
  if (!Array.isArray(rows) || !Array.isArray(oldHeaders) || !Array.isArray(newHeaders)) {
    return Array.isArray(rows) ? rows : [];
  }
  const n = Math.max(oldHeaders.length, newHeaders.length);
  return rows.map((row) => {
    if (!row || typeof row !== 'object') return row;
    const out = { ...row };
    for (let i = 0; i < n; i += 1) {
      const oldH = oldHeaders[i];
      const newH = newHeaders[i];
      if (!newH || oldH === newH) continue;
      const oldVal = oldH != null ? row[oldH] : undefined;
      const isAttendanceCode =
        /^(P|A|WO|H|L|WOP|OD|SL|CL|EL)$/i.test(String(oldVal ?? '').trim());
      if (oldH != null && oldH in row && !(newH in out)) {
        if (!isFormTSEAttendanceDayHeader(newH) || isAttendanceCode) {
          out[newH] = oldVal;
        } else {
          out[newH] = '';
        }
      }
      if (oldH != null && oldH !== newH && oldH in out) {
        delete out[oldH];
      }
    }
    newHeaders.forEach((newH) => {
      if (!newH) return;
      if (String(out[newH] ?? '').trim() !== '') return;
      const target = formTSEEmployeeHeaderKeyNorm(newH);
      if (!target || isFormTSEAttendanceDayHeader(newH)) return;
      for (let i = 0; i < oldHeaders.length; i += 1) {
        const oldH = oldHeaders[i];
        if (!oldH || formTSEEmployeeHeaderKeyNorm(oldH) !== target) continue;
        const oldVal = row[oldH];
        if (oldVal == null || String(oldVal).trim() === '') continue;
        out[newH] = oldVal;
        break;
      }
    });
    return out;
  });
}

/** Keep ATTENDANCE_1…31 keys when a re-read would drop them (autofill must match modal headers). */
export function preserveFormTSEAttendanceHeadersIfNeeded(oldHeaders, newHeaders) {
  const oldList = Array.isArray(oldHeaders) ? oldHeaders : [];
  const newList = Array.isArray(newHeaders) ? newHeaders : [];
  if (!oldList.length || !newList.length) return newList;
  const oldAtt = listFormTSEAttendanceDayHeaders(oldList).length;
  const newAtt = listFormTSEAttendanceDayHeaders(newList).length;
  if (oldAtt < 3 || newAtt >= oldAtt) return newList;
  const n = Math.max(oldList.length, newList.length);
  const out = newList.slice();
  for (let i = 0; i < n; i += 1) {
    const oh = oldList[i];
    const nh = out[i];
    if (!isFormTSEAttendanceDayHeader(nh) && isFormTSEAttendanceDayHeader(oh)) {
      out[i] = oh;
    }
  }
  return out;
}

export function normalizeFormTSETableHeaders(headers) {
  const list = Array.isArray(headers) ? headers : [];
  return list.map((h) => {
    const raw = String(h || '').trim();
    const m = raw.match(/^(ATTENDANCE_\d{1,2})\s*\(\s*\d{1,2}\s*\)$/i);
    return m ? m[1] : h;
  });
}

export function formTSETableHeadersHaveAttendanceBand(headers) {
  return listFormTSEAttendanceDayHeaders(headers).length >= 3;
}

/** Align live modal row keys to export headers (column index) for Excel download. */
export function prepareFormTSEExportRows(liveRows, exportHeaders, sourceHeaders) {
  const hdrs = Array.isArray(exportHeaders) ? exportHeaders : [];
  let srcHdrs = Array.isArray(sourceHeaders) && sourceHeaders.length > 0 ? sourceHeaders : hdrs;
  const rows = Array.isArray(liveRows) ? liveRows : [];
  if (!hdrs.length) return rows.map((row) => ({ ...row }));

  const normalize = (txt) =>
    String(txt || '')
      .replace(/\r?\n/g, ' ')
      .replace(/\s+/g, ' ')
      .trim()
      .toLowerCase()
      .replace(/[^a-z0-9 ]/g, '');

  if (rows.length > 0 && rows[0] && typeof rows[0] === 'object' && !Array.isArray(rows[0])) {
    const rowKeys = Object.keys(rows[0]).filter((k) => !String(k).startsWith('__'));
    const keysMatchExport =
      rowKeys.length === hdrs.length && rowKeys.every((k, i) => k === hdrs[i]);
    if (!keysMatchExport && rowKeys.length === hdrs.length) {
      srcHdrs = rowKeys;
    }
  }

  const alignedRows =
    srcHdrs.length === hdrs.length && srcHdrs.some((h, i) => h !== hdrs[i])
      ? remapRowsToRebuiltTableHeaders(rows, srcHdrs, hdrs)
      : rows;

  const pickRowValue = (row, header, colIndex) => {
    if (!row || typeof row !== 'object') {
      return Array.isArray(row) && colIndex < row.length ? row[colIndex] : '';
    }
    if (Array.isArray(row)) {
      return colIndex < row.length ? row[colIndex] : '';
    }
    // Prefer Form T key-norm match (strips "(2)" suffixes) — same as the autofill modal UI.
    if (isFormTSEAttendanceDayHeader(header)) {
      return sanitizeFormTSEAttendanceMark(getFormTSEKarnatakaRowValueForHeader(row, header));
    }
    const byFormTKey = getFormTSEKarnatakaRowValueForHeader(row, header);
    if (byFormTKey != null && String(byFormTKey).trim() !== '') return byFormTKey;
    if (Object.prototype.hasOwnProperty.call(row, header)) return row[header];
    const srcKey = srcHdrs[colIndex];
    const sameField =
      srcKey && formTSEEmployeeHeaderKeyNorm(srcKey) === formTSEEmployeeHeaderKeyNorm(header);
    if (sameField && Object.prototype.hasOwnProperty.call(row, srcKey)) {
      const srcVal = row[srcKey];
      if (srcVal != null && String(srcVal).trim() !== '') return srcVal;
    }
    if (sameField) {
      const bySrcKey = getFormTSEKarnatakaRowValueForHeader(row, srcKey);
      if (bySrcKey != null && String(bySrcKey).trim() !== '') return bySrcKey;
    }
    const target = normalize(header);
    if (!target) return '';
    const rowKeys = Object.keys(row).filter((k) => !String(k || '').startsWith('__'));
    const exact = rowKeys.filter((k) => normalize(k) === target);
    if (exact.length === 1) return row[exact[0]];
    if (exact.length > 1) {
      const occur =
        hdrs.slice(0, colIndex + 1).filter((h) => normalize(h) === target).length - 1;
      return row[exact[Math.min(Math.max(occur, 0), exact.length - 1)]];
    }
    return '';
  };

  const prepared = alignedRows.map((row, rowIndex) => {
    const out = {};
    hdrs.forEach((header, colIndex) => {
      let value = pickRowValue(row, header, colIndex);
      if (
        (value == null || String(value).trim() === '') &&
        /s\.?\s*no|serial|sl\.?\s*no/i.test(String(header || ''))
      ) {
        value = rowIndex + 1;
      }
      // Total OT hours always exports as NIL (never attendance/payroll HH:MM).
      if (isFormTSEKarnatakaTotalOtHoursHeader(header)) {
        value = FORM_T_KA_OT_HOURS_NIL;
      }
      out[header] = value != null ? value : '';
    });
    if (row && typeof row === 'object' && !Array.isArray(row)) {
      const lookup = String(row.__employeeLookupName || row.__employeelookupname || '').trim();
      if (lookup) out.__employeeLookupName = lookup;
    }
    const named = getFormTSEKarnatakaEmployeeNameFromRow(
      { ...(row && typeof row === 'object' && !Array.isArray(row) ? row : {}), ...out },
      hdrs
    );
    if (named) {
      hdrs.forEach((header) => {
        if (isFormTSEEmployeeNameHeader(header)) out[header] = named;
      });
      out.__employeeLookupName = named;
    }
    return out;
  });
  return applyFormTSEKarnatakaPayrollOnlyWhenAttended(
    overlayFormTSEPayrollOntoRowsByName(prepared, hdrs, rows),
    hdrs
  );
}

export function repairFormTSETableHeadersFromWorkbook(workbook, parsed = {}) {
  if (!workbook || !parsed) return parsed;
  const sheetName = parsed.sheetName || workbook.SheetNames?.[0];
  const ws = sheetName ? workbook.Sheets[sheetName] : null;
  if (!ws) return parsed;

  const merges = ws['!merges'] || [];
  const rawCell = (r, c) => {
    if (r < 0 || c < 0) return '';
    const ref = XLSX.utils.encode_cell({ r, c });
    const cell = ws[ref];
    return cell && cell.v != null ? String(cell.v).trim() : '';
  };
  const getMergedAwareCellText = (r, c) => {
    const direct = rawCell(r, c);
    if (direct) return direct;
    for (let i = 0; i < merges.length; i += 1) {
      const m = merges[i];
      if (!m?.s || !m?.e) continue;
      if (r >= m.s.r && r <= m.e.r && c >= m.s.c && c <= m.e.c) {
        const topLeft = rawCell(m.s.r, m.s.c);
        if (topLeft) return topLeft;
      }
    }
    return '';
  };

  const jsonData = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  let mergeMaxCol = 0;
  merges.forEach((m) => {
    if (m?.e?.c != null) mergeMaxCol = Math.max(mergeMaxCol, Number(m.e.c) + 1);
  });
  const ref = ws['!ref'];
  let refCols = 0;
  if (ref) {
    try {
      const range = XLSX.utils.decode_range(ref);
      refCols = (range.e.c || 0) + 1;
    } catch (_) {
      refCols = 0;
    }
  }
  const effectiveSheetCols = Math.max(
    ...(jsonData || []).map((row) => (Array.isArray(row) ? row.length : 0)),
    mergeMaxCol,
    refCols,
    80
  );
  const headerRowIndex =
    parsed.originalHeaderRowIndex ?? parsed.headerRowIndex ?? -1;
  const dataStartIndex = parsed.dataStartIndex ?? -1;
  const tableStartCol = parsed.tableStartCol ?? 0;

  const rebuilt = rebuildFormTSETableHeadersFromSheet({
    headerRowIndex,
    getMergedAwareCellText,
    getRawCellText: rawCell,
    jsonData,
    effectiveSheetCols,
    tableStartCol,
    dataStartIndex,
  });
  if (!rebuilt?.headers?.length) return parsed;

  const oldHeaders = Array.isArray(parsed.headers) ? parsed.headers : [];
  const mergedHeaders = preserveFormTSEAttendanceHeadersIfNeeded(
    oldHeaders,
    normalizeFormTSETableHeaders(rebuilt.headers)
  );
  return {
    ...parsed,
    headers: mergedHeaders,
    headerRowIndex: rebuilt.headerRowIndex,
    dataStartIndex: rebuilt.dataStartIndex,
    tableStartCol: rebuilt.tableStartCol,
    rows: remapRowsToRebuiltTableHeaders(parsed.rows || [], oldHeaders, mergedHeaders),
  };
}

/** Standard Form T Karnataka: A–I identity, J+ attendance days (0-based). */
export const FORM_T_KA_IDENTITY_START_COL0 = 0;
export const FORM_T_KA_ATTENDANCE_START_COL0 = 9; // column J

/** Table grid borders start on Excel row 12 (column headers) on the KA template. */
export const FORM_T_KA_TABLE_BORDER_START_ROW = 12;

function formTSEExcelJsCellText(worksheet, r, c) {
  const cell = worksheet?.getCell(r, c);
  if (!cell) return '';
  const v = cell.value;
  if (v != null) {
    if (typeof v === 'string' || typeof v === 'number' || typeof v === 'boolean') return String(v).trim();
    if (typeof v === 'object') {
      if (Array.isArray(v.richText)) return v.richText.map((rt) => rt?.text || '').join('').trim();
      if (v.text != null) return String(v.text).trim();
      if (v.result != null) return String(v.result).trim();
    }
  }
  try {
    const text = cell.text;
    if (typeof text === 'string' && text.trim() && text !== '[object Object]') {
      return text.trim();
    }
  } catch (_) {
    /* merge slave cells throw when reading .text */
  }
  return '';
}

function formTSERowLooksLikeConsecutiveNumberStrip(worksheet, r, startCol, count = 8) {
  let ok = 0;
  for (let i = 0; i < count; i += 1) {
    const v = formTSEExcelJsCellText(worksheet, r, startCol + i);
    if (String(v).trim() === String(i + 1)) ok += 1;
  }
  return ok >= 6;
}

/** Last Save auto-download workbook. Survives Statutory remounts in the same page session. */
let rememberedFormTSELastExport = null;

export function rememberFormTSELastExport(snapshot) {
  if (snapshot) rememberedFormTSELastExport = snapshot;
}

export function getRememberedFormTSELastExport() {
  return rememberedFormTSELastExport;
}

/** True for a real employee name — not header labels that can bleed into data rows via merges. */
export function looksLikeFormTSEEmployeeNameCell(text) {
  const t = String(text || '').trim();
  if (t.length < 2 || !/[a-zA-Z]{2,}/.test(t)) return false;
  if (
    /name\s+of\s+employee|employee\s*name|father|husband|designation|department|date\s+of\s+joining|wages\s+fixed|attendance|s\.?\s*no|sl\.?\s*no|serial|gender|male\s*\/\s*female|esi\s*no|uan\s*no|minimum\s+wages|contract\s+labour|migrant\s+workmen|establishment|employer|address\s+of|month\s*\/?\s*year|karnataka|payment\s+of\s+wages|see\s+rule|combined\s+muster|name\s+of\s+the/i.test(
      t
    )
  ) {
    return false;
  }
  return true;
}

function formTSERowLooksIdentityShiftedToJ(worksheet, r) {
  const shift = FORM_T_KA_ATTENDANCE_START_COL0;
  // Day-number / column-index strips (1,2,3…) are not employee rows.
  if (
    formTSERowLooksLikeConsecutiveNumberStrip(worksheet, r, shift + 1) ||
    formTSERowLooksLikeConsecutiveNumberStrip(worksheet, r, 1)
  ) {
    return false;
  }
  const nameB = formTSEExcelJsCellText(worksheet, r, 2);
  const nameK = formTSEExcelJsCellText(worksheet, r, shift + 2);
  const serialJ = formTSEExcelJsCellText(worksheet, r, shift + 1);
  const serialA = formTSEExcelJsCellText(worksheet, r, 1);
  const genderM = formTSEExcelJsCellText(worksheet, r, shift + 4);
  const genderD = formTSEExcelJsCellText(worksheet, r, 4);
  const fatherL = formTSEExcelJsCellText(worksheet, r, shift + 3);
  const personAtK = looksLikeFormTSEEmployeeNameCell(nameK);
  const personAtB = looksLikeFormTSEEmployeeNameCell(nameB);
  if (personAtK && !personAtB) return true;
  if (
    /^\d{1,4}$/.test(serialJ) &&
    !/^\d{1,4}$/.test(serialA) &&
    /^(male|female)$/i.test(genderM) &&
    !/^(male|female)$/i.test(genderD)
  ) {
    return true;
  }
  if (
    /^\d{1,4}$/.test(serialJ) &&
    !personAtB &&
    (personAtK || /^(male|female)$/i.test(genderM) || looksLikeFormTSEEmployeeNameCell(fatherL))
  ) {
    return true;
  }
  return false;
}

/** Identity already in A–I but a prior bad write also dumped S.NO/Name under day-1 (J–R). */
function formTSERowHasDuplicateIdentityUnderJ(worksheet, r) {
  const shift = FORM_T_KA_ATTENDANCE_START_COL0;
  const nameB = formTSEExcelJsCellText(worksheet, r, 2);
  const nameK = formTSEExcelJsCellText(worksheet, r, shift + 2);
  const serialJ = formTSEExcelJsCellText(worksheet, r, shift + 1);
  const serialA = formTSEExcelJsCellText(worksheet, r, 1);
  const personAtB = looksLikeFormTSEEmployeeNameCell(nameB);
  const personAtK = looksLikeFormTSEEmployeeNameCell(nameK);
  if (!personAtB && !/^\d{1,4}$/.test(serialA)) return false;
  if (personAtK && personAtB) {
    const nb = nameB.toLowerCase();
    const nk = nameK.toLowerCase();
    if (nk.includes(nb.slice(0, Math.min(4, nb.length))) || nb.includes(nk.slice(0, Math.min(4, nk.length)))) {
      return true;
    }
  }
  if (/^\d{1,4}$/.test(serialJ) && /^\d{1,4}$/.test(serialA) && serialJ === serialA) return true;
  return false;
}

/** Wipe employee body cells so a dirty Draft/template cannot leave J-shifted leftovers. */
function clearFormTSEExcelJsDataRowBand(worksheet, rowFrom, rowTo, colThrough = 80) {
  if (!worksheet || rowFrom < 1 || rowTo < rowFrom) return;
  const maxC = Math.max(40, Number(colThrough) || 80);
  unmergeFormTSEWorksheetRows(worksheet, rowFrom, rowTo);
  for (let r = rowFrom; r <= rowTo; r += 1) {
    for (let c = 1; c <= maxC; c += 1) {
      worksheet.getCell(r, c).value = null;
    }
  }
}

function formTSEParseMergeRangeLabel(worksheet, label) {
  const parts = String(label || '').split(':');
  if (parts.length !== 2) return null;
  try {
    const tl = worksheet.getCell(parts[0]);
    const br = worksheet.getCell(parts[1]);
    if (!tl || !br) return null;
    return { top: tl.row, left: tl.col, bottom: br.row, right: br.col };
  } catch (_) {
    return null;
  }
}

function eachFormTSEWorksheetMerge(worksheet, visit) {
  if (!worksheet || typeof visit !== 'function') return;
  const seen = new Set();
  const add = (top, left, bottom, right) => {
    if (![top, left, bottom, right].every((n) => Number.isFinite(n) && n >= 1)) return;
    const key = `${top}:${left}:${bottom}:${right}`;
    if (seen.has(key)) return;
    seen.add(key);
    visit({ top, left, bottom, right });
  };
  const ingestMergeLike = (merge) => {
    if (!merge) return;
    if (typeof merge === 'string') {
      const parsed = formTSEParseMergeRangeLabel(worksheet, merge);
      if (parsed) add(parsed.top, parsed.left, parsed.bottom, parsed.right);
      return;
    }
    if (typeof merge !== 'object') return;
    const top = Number(merge.top ?? merge.s?.r ?? merge.tl?.row);
    const left = Number(merge.left ?? merge.s?.c ?? merge.tl?.col);
    const bottom = Number(merge.bottom ?? merge.e?.r ?? merge.br?.row);
    const right = Number(merge.right ?? merge.e?.c ?? merge.br?.col);
    if (Number.isFinite(top) && Number.isFinite(left)) {
      add(top, left, Number.isFinite(bottom) ? bottom : top, Number.isFinite(right) ? right : left);
    }
  };
  const ingestCollection = (obj) => {
    if (!obj) return;
    if (typeof obj.forEach === 'function') {
      try {
        obj.forEach((merge) => ingestMergeLike(merge));
      } catch (_) {
        /* ignore */
      }
    }
    if (Array.isArray(obj)) {
      obj.forEach((merge) => ingestMergeLike(merge));
      return;
    }
    if (typeof obj !== 'object') return;
    if (obj.merges && obj.merges !== obj) ingestCollection(obj.merges);
    Object.keys(obj).forEach((addr) => {
      if (addr === 'merges') return;
      ingestMergeLike(obj[addr]);
      const parsed = formTSEParseMergeRangeLabel(worksheet, addr);
      if (parsed) add(parsed.top, parsed.left, parsed.bottom, parsed.right);
    });
  };
  ingestCollection(worksheet._merges);
  ingestCollection(worksheet.model?.merges);
}

function unmergeFormTSEWorksheetRows(worksheet, rowFrom, rowTo) {
  if (!worksheet || rowFrom < 1 || rowTo < rowFrom) return;
  const toUnmerge = [];
  eachFormTSEWorksheetMerge(worksheet, (m) => {
    if (m.bottom < rowFrom || m.top > rowTo) return;
    toUnmerge.push(m);
  });
  toUnmerge.forEach((m) => {
    try {
      worksheet.unMergeCells(m.top, m.left, m.bottom, m.right);
    } catch (_) {
      /* already unmerged */
    }
  });
}

/** Drop every A–I merge on a data row, including leftover B:C (Name+Father). */
function stripFormTSEIdentityMergesOnRow(worksheet, excelRow) {
  if (!worksheet || excelRow < 1) return;
  formTSEUnmergeCovering(worksheet, excelRow, 1, excelRow, 9);
  unmergeFormTSEWorksheetRows(worksheet, excelRow, excelRow);
  for (let c = 1; c <= 9; c += 1) {
    try {
      const cell = worksheet.getCell(excelRow, c);
      const master = cell?.master;
      if (master?.address) worksheet.unMergeCells(master.address);
      else if (cell?.isMerged && cell.address) worksheet.unMergeCells(cell.address);
    } catch (_) {
      /* already unmerged */
    }
  }
  try {
    worksheet.unMergeCells(excelRow, 2, excelRow, 3);
  } catch (_) {
    /* no B:C merge */
  }
  try {
    worksheet.unMergeCells(excelRow, 1, excelRow, 9);
  } catch (_) {
    /* no A:I merge */
  }
  try {
    worksheet.unMergeCells(`B${excelRow}:C${excelRow}`);
  } catch (_) {
    /* no B:C address merge */
  }
  const merges = worksheet.model?.merges;
  if (Array.isArray(merges)) {
    worksheet.model.merges = merges.filter((label) => {
      const parsed = formTSEParseMergeRangeLabel(worksheet, label);
      if (!parsed) return true;
      if (parsed.bottom < excelRow || parsed.top > excelRow) return true;
      if (parsed.right < 1 || parsed.left > 9) return true;
      return false;
    });
  }
}

function unmergeAllFormTSEDataIdentityBands(worksheet) {
  if (!worksheet) return;
  const rowTo = Math.max(80, Number(worksheet.rowCount) || 40);
  for (let r = 14; r <= rowTo; r += 1) {
    stripFormTSEIdentityMergesOnRow(worksheet, r);
  }
}

function formTSEA1ColLettersToNumber(letters) {
  let n = 0;
  const s = String(letters || '').toUpperCase();
  for (let i = 0; i < s.length; i += 1) {
    const code = s.charCodeAt(i);
    if (code < 65 || code > 90) return 0;
    n = n * 26 + (code - 64);
  }
  return n;
}

function formTSEMergeRefHitsDataIdentity(ref) {
  const parts = String(ref || '').split(':');
  if (parts.length !== 2) return false;
  const parse = (addr) => {
    const m = String(addr || '').trim().match(/^([A-Z]+)(\d+)$/i);
    if (!m) return null;
    return { col: formTSEA1ColLettersToNumber(m[1]), row: Number(m[2]) };
  };
  const a = parse(parts[0]);
  const b = parse(parts[1]);
  if (!a || !b) return false;
  const top = Math.min(a.row, b.row);
  const bottom = Math.max(a.row, b.row);
  const left = Math.min(a.col, b.col);
  const right = Math.max(a.col, b.col);
  if (bottom < 15 || top > 120) return false;
  return left <= 3 && right >= 2;
}

function formTSEWorksheetXmlPaths(zip) {
  return Object.keys(zip.files || {}).filter((p) => /xl\/worksheets\/sheet[^/]*\.xml$/i.test(p));
}

/** Drop Name/Father data-row merges and keep mergeCells count in sync. */
function rewriteFormTSESheetIdentityMergesXml(xml) {
  if (!xml) return xml;
  return String(xml).replace(/<mergeCells\b[^>]*>[\s\S]*?<\/mergeCells>/gi, (block) => {
    const tags = [];
    const re = /<mergeCell\b[^>]*\/>|<mergeCell\b[^>]*>\s*<\/mergeCell>/gi;
    let m;
    while ((m = re.exec(block))) {
      const refM = m[0].match(/\bref\s*=\s*["']([^"']+)["']/i);
      if (!refM) continue;
      if (formTSEMergeRefHitsDataIdentity(refM[1])) continue;
      tags.push(m[0].replace(/>\s*<\/mergeCell>/i, '/>'));
    }
    if (tags.length === 0) return '';
    return `<mergeCells count="${tags.length}">${tags.join('')}</mergeCells>`;
  });
}

/** Official Form T_KA keeps B20:C20 merged in sheet XML even after ExcelJS unMergeCells. */
export async function stripFormTSEDataIdentityMergesFromXlsx(arrayBuffer) {
  if (!arrayBuffer || arrayBuffer.byteLength < 32) return arrayBuffer;
  try {
    const zip = await JSZip.loadAsync(arrayBuffer);
    const paths = formTSEWorksheetXmlPaths(zip);
    let touched = false;
    for (let i = 0; i < paths.length; i += 1) {
      const path = paths[i];
      const file = zip.file(path);
      if (!file) continue;
      const xml = await file.async('string');
      const next = rewriteFormTSESheetIdentityMergesXml(xml);
      if (next !== xml) {
        touched = true;
        zip.file(path, next);
      }
    }
    if (!touched) return arrayBuffer;
    return zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' });
  } catch (_) {
    return arrayBuffer;
  }
}

function formTSEXmlEscapeText(text) {
  return String(text ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function formTSESheetXmlCellStyleId(xml, ref) {
  if (!xml || !ref) return '';
  const cellRe = new RegExp(
    `<c\\b(?=[^>]*\\br=["']${ref}["'])[^>]*>[\\s\\S]*?</c>|<c\\b(?=[^>]*\\br=["']${ref}["'])[^>]*/>`,
    'i'
  );
  const existing = xml.match(cellRe);
  const style = existing && existing[0].match(/\bs="([^"]+)"/i);
  return style ? style[1] : '';
}

/** Row 20 Name/Father sat on the Date: merge — reuse a data-row xf so Father "abc" is centered. */
function formTSEIdentityXmlStyleId(xml, ref) {
  const col = String(ref || '').replace(/\d+$/i, '');
  const rowNum = Number(String(ref || '').replace(/^[A-Z]+/i, ''));
  const own = formTSESheetXmlCellStyleId(xml, ref);
  const preferNeighbor = rowNum === 20 || !own;
  if (!preferNeighbor) return own;
  const probes = [rowNum - 1, rowNum + 1, 19, 18, 16, 15, 21]
    .filter((r) => Number.isFinite(r) && r >= 15 && r !== rowNum)
    .map((r) => `${col}${r}`);
  for (let i = 0; i < probes.length; i += 1) {
    const id = formTSESheetXmlCellStyleId(xml, probes[i]);
    if (id) return id;
  }
  return own;
}

function formTSEReplaceSheetCellInline(xml, ref, text) {
  const escaped = formTSEXmlEscapeText(text);
  const cellRe = new RegExp(
    `<c\\b(?=[^>]*\\br=["']${ref}["'])[^>]*>[\\s\\S]*?</c>|<c\\b(?=[^>]*\\br=["']${ref}["'])[^>]*/>`,
    'i'
  );
  const existing = xml.match(cellRe);
  const sId = formTSEIdentityXmlStyleId(xml, ref);
  const sAttr = sId ? ` s="${sId}"` : '';
  const spaceAttr = /^\s|\s$/.test(String(text ?? '')) ? ' xml:space="preserve"' : '';
  const cell = `<c r="${ref}"${sAttr} t="inlineStr"><is><t${spaceAttr}>${escaped}</t></is></c>`;
  if (existing) return xml.replace(cellRe, cell);
  const rowNum = String(ref).replace(/^[A-Z]+/i, '');
  const rowOpen = new RegExp(`(<row\\b[^>]*\\br=["']${rowNum}["'][^>]*>)`, 'i');
  if (rowOpen.test(xml)) return xml.replace(rowOpen, `$1${cell}`);
  return xml;
}

function collectFormTSEIdentityXmlAssignments(list, hdrs, fatherHeader) {
  const assignments = [];
  const startRow = 15;
  const pushPair = (excelRow, row) => {
    if (!row || excelRow < 1) return;
    const named = getFormTSEKarnatakaEmployeeNameFromRow(row, hdrs);
    const father = fatherHeader
      ? String(getFormTSEKarnatakaRowValueForHeader(row, fatherHeader) ?? '').trim()
      : '';
    if (named) assignments.push({ ref: `B${excelRow}`, text: named });
    assignments.push({ ref: `C${excelRow}`, text: father });
  };
  list.forEach((row, i) => pushPair(startRow + i, row));
  if (list[5]) pushPair(20, list[5]);
  return assignments;
}

/**
 * Last-pass write of Name/Father into sheet XML. Must not ExcelJS-writeBuffer
 * afterwards — that recreates the official B20:C20 merge and puts Father ("abc")
 * in the Name cell that Excel actually opens.
 */
export async function forceFormTSEIdentityCellsInSheetXml(arrayBuffer, rows, headers = [], employees = []) {
  const hdrs = Array.isArray(headers) ? headers : [];
  const list = overlayFormTSEPeopleNamesOntoRows(
    (Array.isArray(rows) ? rows : []).filter((row) => row && typeof row === 'object'),
    hdrs,
    employees || []
  );
  if (!arrayBuffer || arrayBuffer.byteLength < 32 || list.length === 0) return arrayBuffer;
  const fatherHeader =
    hdrs.find((h) => isFormTSEFatherHusbandHeader(h)) ||
    Object.keys(list[0] || {}).find((k) => isFormTSEFatherHusbandHeader(k)) ||
    '';
  const assignments = collectFormTSEIdentityXmlAssignments(list, hdrs, fatherHeader);
  if (assignments.length === 0) return arrayBuffer;
  try {
    const zip = await JSZip.loadAsync(arrayBuffer);
    const paths = formTSEWorksheetXmlPaths(zip);
    for (let i = 0; i < paths.length; i += 1) {
      const file = zip.file(paths[i]);
      if (!file) continue;
      let xml = await file.async('string');
      xml = rewriteFormTSESheetIdentityMergesXml(xml);
      assignments.forEach(({ ref, text }) => {
        xml = formTSEReplaceSheetCellInline(xml, ref, text);
      });
      zip.file(paths[i], xml);
    }
    return zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' });
  } catch (err) {
    console.warn('Form T identity XML stamp failed:', err);
    return arrayBuffer;
  }
}

function unmergeFormTSEIdentityBandOnRow(worksheet, excelRow) {
  stripFormTSEIdentityMergesOnRow(worksheet, excelRow);
}

/**
 * Write Name (col B) and Father (col C) onto physical cells after unmerging.
 * Official Form T_KA puts a B:C merge on the Date: footer; employee 6 lands there
 * and Father ("abc") replaces Ashok unless this runs last.
 */
function applyFormTSEIdentityCellAlignment(cell, template) {
  if (!cell) return;
  const from =
    template && template.alignment && typeof template.alignment === 'object'
      ? template.alignment
      : {};
  const alignment = {
    ...(cell.alignment || {}),
    horizontal: from.horizontal || 'center',
    vertical: from.vertical || 'middle',
    wrapText: from.wrapText === true,
    indent: 0,
  };
  cell.alignment = alignment;
  try {
    const prev = cell.style && typeof cell.style === 'object' ? { ...cell.style } : {};
    cell.style = { ...prev, alignment: { ...(prev.alignment || {}), ...alignment } };
  } catch (_) {
    /* alignment property is enough */
  }
}

function writeFormTSEPhysicalNameAndFather(worksheet, excelRow, name, father) {
  if (!worksheet || excelRow < 1) return;
  const nameText = name == null ? '' : String(name).trim();
  const fatherText = father == null ? '' : String(father).trim();
  const templateRow = excelRow > 15 ? excelRow - 1 : excelRow + 1;
  const assign = () => {
    stripFormTSEIdentityMergesOnRow(worksheet, excelRow);
    const fatherCell = worksheet.getCell(excelRow, 3);
    const nameCell = worksheet.getCell(excelRow, 2);
    fatherCell.value = fatherText;
    nameCell.value = nameText;
    applyFormTSEIdentityCellAlignment(fatherCell, worksheet.getCell(templateRow, 3));
    applyFormTSEIdentityCellAlignment(nameCell, worksheet.getCell(templateRow, 2));
  };
  assign();
  const nameNow = formTSEExcelJsCellText(worksheet, excelRow, 2);
  const fatherNow = formTSEExcelJsCellText(worksheet, excelRow, 3);
  const nameLost = nameText && nameNow.toLowerCase() !== nameText.toLowerCase();
  const fatherLostInName =
    fatherText &&
    !fatherNow &&
    nameNow.toLowerCase() === fatherText.toLowerCase();
  if (nameLost || fatherLostInName) assign();
}

/** Date / signature / system-generated row that must stay below employee data. */
function findFormTSEKarnatakaSheetFooterRow(worksheet, fromRow = 14) {
  if (!worksheet) return 0;
  const maxR = Math.max(160, Number(worksheet.rowCount) || 0);
  for (let r = Math.max(1, fromRow); r <= maxR; r += 1) {
    const a = formTSEExcelJsCellText(worksheet, r, 1).toLowerCase();
    let line = a;
    for (let c = 2; c <= 5; c += 1) {
      const t = formTSEExcelJsCellText(worksheet, r, c).toLowerCase();
      if (t) line += ` ${t}`;
    }
    if (
      /^date\s*:?$/.test(a) ||
      /^date\s*:/.test(a) ||
      /\bdate\s*:/.test(line) ||
      /system\s+generated/.test(line) ||
      /authorised\s+signatory|authorized\s+signatory/.test(line) ||
      /signature\s+of\s+employer/.test(line)
    ) {
      return r;
    }
  }
  return 0;
}

/**
 * Make room for every employee row and flatten leftover template merges.
 * ExcelJS getCell() on a B:C merge writes Name and Father into the same cell, so
 * the last employee name becomes Father ("abc") in the downloaded Excel.
 */
function ensureFormTSEKarnatakaDataRowCapacity(worksheet, dataStartRow, rowCount, footerRowHint = 0) {
  if (!worksheet || rowCount < 1 || dataStartRow < 1) return Number(footerRowHint) || 0;
  const lastDataRow = dataStartRow + rowCount - 1;
  let footerRow =
    Number(footerRowHint) > 0
      ? Number(footerRowHint)
      : findFormTSEKarnatakaSheetFooterRow(worksheet, dataStartRow);
  const unmergeTo = Math.max(lastDataRow + 6, footerRow || lastDataRow);
  unmergeFormTSEWorksheetRows(worksheet, dataStartRow, unmergeTo);
  for (let r = dataStartRow; r <= unmergeTo; r += 1) {
    stripFormTSEIdentityMergesOnRow(worksheet, r);
  }
  footerRow = findFormTSEKarnatakaSheetFooterRow(worksheet, dataStartRow) || footerRow;
  // Official Form T_KA: Date: + B:C merge is always Excel row 20 (6th employee).
  stripFormTSEIdentityMergesOnRow(worksheet, 20);
  try {
    worksheet.unMergeCells('B20:C20');
  } catch (_) {
    /* no B20:C20 */
  }
  if (
    dataStartRow <= 20 &&
    lastDataRow >= 20 &&
    typeof worksheet.spliceRows === 'function'
  ) {
    const a20 = formTSEExcelJsCellText(worksheet, 20, 1);
    if (/^date\s*:/i.test(a20) || !/^\d{1,4}$/.test(a20)) {
      try {
        worksheet.spliceRows(20, 0, []);
        footerRow = footerRow >= 20 ? footerRow + 1 : footerRow;
      } catch (_) {
        /* splice can fail on leftover merges */
      }
      stripFormTSEIdentityMergesOnRow(worksheet, 20);
      stripFormTSEIdentityMergesOnRow(worksheet, 21);
    }
  }
  if (footerRow > dataStartRow && typeof worksheet.spliceRows === 'function') {
    const available = footerRow - dataStartRow;
    const need = rowCount - available;
    if (need > 0) {
      try {
        const blanks = Array.from({ length: need }, () => []);
        worksheet.spliceRows(footerRow, 0, ...blanks);
        footerRow += need;
      } catch (_) {
        /* ExcelJS splice can fail on leftover merges — move footer below data instead. */
      }
    }
  }
  const stillFooter = findFormTSEKarnatakaSheetFooterRow(worksheet, dataStartRow);
  if (stillFooter > 0 && stillFooter <= lastDataRow) {
    footerRow = moveFormTSEKarnatakaFooterBelowData(worksheet, lastDataRow, stillFooter);
  }
  unmergeFormTSEWorksheetRows(worksheet, dataStartRow, dataStartRow + rowCount - 1);
  return footerRow;
}

/** Copy Date / signature / system-generated text below the last employee row. */
function moveFormTSEKarnatakaFooterBelowData(worksheet, lastDataRow, footerRow) {
  if (!worksheet || lastDataRow < 1 || footerRow < 1) return footerRow;
  if (footerRow > lastDataRow) return footerRow;
  const dest = lastDataRow + 1;
  unmergeFormTSEWorksheetRows(worksheet, Math.min(footerRow, dest), dest + 2);
  const maxCol = Math.max(Number(worksheet.columnCount) || 0, 40);
  for (let c = 1; c <= maxCol; c += 1) {
    const src = worksheet.getCell(footerRow, c);
    worksheet.getCell(dest, c).value = src.value == null ? null : src.value;
    src.value = null;
  }
  return dest;
}

function findFormTSEKarnatakaEmployeeDataStartRow(worksheet) {
  if (!worksheet) return 15;
  const rowTo = Math.min(40, Math.max(20, Number(worksheet.rowCount) || 20));
  let headerRow = 0;
  for (let r = 1; r <= rowTo; r += 1) {
    const a = formTSEExcelJsCellText(worksheet, r, 1);
    const b = formTSEExcelJsCellText(worksheet, r, 2);
    if (/s\.?\s*no/i.test(a) && /name/i.test(b)) {
      headerRow = r;
      break;
    }
  }
  let start = headerRow > 0 ? headerRow + 1 : 15;
  for (let guard = 0; guard < 4; guard += 1) {
    if (!formTSERowLooksLikeConsecutiveNumberStrip(worksheet, start, 1)) break;
    start += 1;
  }
  return Math.max(start, 15);
}

/** Official Form T column 2 is Name of Employee — not CLRA "principal employer". */
export function restoreFormTSEKarnatakaEmployeeNameColumnHeader(worksheet) {
  if (!worksheet) return false;
  let touched = false;
  for (let r = 11; r <= 16; r += 1) {
    const a = formTSEExcelJsCellText(worksheet, r, 1);
    const b = formTSEExcelJsCellText(worksheet, r, 2);
    const c = formTSEExcelJsCellText(worksheet, r, 3);
    // Row 11 is the employer header box — only rewrite the table header row.
    const isTableHeader =
      /s\.?\s*no/i.test(a) ||
      /father\s*\/\s*husband|father.*name/i.test(c) ||
      (r === 12 && /principal\s+employer/i.test(b));
    if (!isTableHeader) continue;
    if (/name\s+of\s+employee/i.test(b) && !/principal\s+employer/i.test(b)) continue;
    if (
      /principal\s+employer/i.test(b) ||
      (/name\s+and\s+address/i.test(b) && /employer/i.test(b) && !/establishment/i.test(b))
    ) {
      unmergeFormTSEIdentityBandOnRow(worksheet, r);
      const cell = worksheet.getCell(r, 2);
      cell.value = 'Name of Employee';
      cell.alignment = {
        ...(cell.alignment || {}),
        wrapText: true,
        vertical: 'middle',
        horizontal: 'center',
      };
      touched = true;
    }
  }
  return touched;
}

/**
 * True when any employee serial row shows Father in the Name column (Father cell blank).
 * Official Form T_KA always lands the 6th person on Excel row 20 (Date: B:C merge),
 * even when more employees follow — so this is not only the last row.
 */
export function formTSEWorksheetHasNameFatherSwap(worksheet) {
  if (!worksheet) return false;
  const rowTo = Math.max(80, Number(worksheet.rowCount) || 0);
  const serialRows = [];
  const fathers = [];
  for (let r = 14; r <= rowTo; r += 1) {
    if (formTSERowLooksLikeConsecutiveNumberStrip(worksheet, r, 1)) continue;
    if (!/^\d{1,4}$/.test(formTSEExcelJsCellText(worksheet, r, 1))) continue;
    serialRows.push(r);
    const fatherC = formTSEExcelJsCellText(worksheet, r, 3);
    if (fatherC) fathers.push(fatherC.toLowerCase());
  }
  for (let i = 0; i < serialRows.length; i += 1) {
    const r = serialRows[i];
    const nameB = formTSEExcelJsCellText(worksheet, r, 2);
    const fatherC = formTSEExcelJsCellText(worksheet, r, 3);
    const genderD = formTSEExcelJsCellText(worksheet, r, 4);
    if (fatherC || !nameB) continue;
    if (!/^(male|female)$/i.test(genderD)) continue;
    const prevFather =
      i > 0 ? formTSEExcelJsCellText(worksheet, serialRows[i - 1], 3) : '';
    if (prevFather && nameB.toLowerCase() === prevFather.toLowerCase()) return true;
    if (fathers.includes(nameB.toLowerCase())) return true;
  }
  return false;
}

/** @deprecated use formTSEWorksheetHasNameFatherSwap — kept for existing callers. */
export function formTSEWorksheetLastIdentityLooksNameFatherSwap(worksheet) {
  return formTSEWorksheetHasNameFatherSwap(worksheet);
}

/**
 * Write S.NO / Name / Father from autofill rows onto A–C (unmerges first).
 * Used on Download Draft so a cached workbook cannot keep Father ("abc") in Name.
 */
export function stampFormTSEKarnatakaEmployeeIdentity(worksheet, rows, headers = [], employees = []) {
  const hdrs = Array.isArray(headers) ? headers : [];
  const list = overlayFormTSEPeopleNamesOntoRows(
    (Array.isArray(rows) ? rows : []).filter((row) => row && typeof row === 'object'),
    hdrs,
    employees
  );
  if (!worksheet || list.length === 0) return 0;
  restoreFormTSEKarnatakaEmployeeNameColumnHeader(worksheet);
  unmergeAllFormTSEDataIdentityBands(worksheet);
  const startRow = findFormTSEKarnatakaEmployeeDataStartRow(worksheet);
  const footer = findFormTSEKarnatakaSheetFooterRow(worksheet, startRow);
  ensureFormTSEKarnatakaDataRowCapacity(worksheet, startRow, list.length, footer);
  const fatherHeader =
    hdrs.find((h) => isFormTSEFatherHusbandHeader(h)) ||
    Object.keys(list[0] || {}).find((k) => isFormTSEFatherHusbandHeader(k)) ||
    '';
  const serialHeader =
    hdrs.find((h) => isFormTSESerialNumberHeader(h)) ||
    Object.keys(list[0] || {}).find((k) => isFormTSESerialNumberHeader(k)) ||
    'S.NO';
  let stamped = 0;
  const writeIdentityAt = (excelRow, row, serialFallback) => {
    unmergeFormTSEIdentityBandOnRow(worksheet, excelRow);
    const named = getFormTSEKarnatakaEmployeeNameFromRow(row, hdrs);
    const father = fatherHeader
      ? String(getFormTSEKarnatakaRowValueForHeader(row, fatherHeader) ?? '').trim()
      : '';
    let serial = getFormTSEKarnatakaRowValueForHeader(row, serialHeader);
    if (serial == null || String(serial).trim() === '') serial = serialFallback;
    const serialText = String(serial).trim();
    worksheet.getCell(excelRow, 1).value = /^\d+$/.test(serialText) ? Number(serialText) : serialText;
    writeFormTSEPhysicalNameAndFather(worksheet, excelRow, named, father);
    if (named) stamped += 1;
  };
  for (let i = 0; i < list.length; i += 1) {
    writeIdentityAt(startRow + i, list[i], i + 1);
  }
  // Re-stamp by S.NO so a blank template row cannot shift Ashok off serial 6.
  const rowTo = Math.max(40, Number(worksheet.rowCount) || 0);
  for (let i = 0; i < list.length; i += 1) {
    const row = list[i];
    let serial = getFormTSEKarnatakaRowValueForHeader(row, serialHeader);
    if (serial == null || String(serial).trim() === '') serial = i + 1;
    const serialText = String(serial).trim();
    if (!/^\d+$/.test(serialText)) continue;
    for (let r = startRow; r <= rowTo; r += 1) {
      if (formTSERowLooksLikeConsecutiveNumberStrip(worksheet, r, 1)) continue;
      if (formTSEExcelJsCellText(worksheet, r, 1) !== serialText) continue;
      const nameB = formTSEExcelJsCellText(worksheet, r, 2);
      if (/name\s+of\s+employee|father\s*\/\s*husband|designation/i.test(nameB)) continue;
      writeIdentityAt(r, row, serialText);
      break;
    }
  }
  return stamped;
}

/** Move S.NO/Name from J–R onto A–I when a write dumped identity under attendance. */
export function shiftFormTSEWorksheetIdentityFromJToA(worksheet) {
  if (!worksheet) return 0;
  const shift = FORM_T_KA_ATTENDANCE_START_COL0;
  const maxCol = Math.max(worksheet.columnCount || 0, 80);
  let headerRow = 0;
  const rowTo = Math.max(120, Number(worksheet.rowCount) || 0);
  for (let r = 1; r <= Math.min(40, rowTo); r += 1) {
    const b = formTSEExcelJsCellText(worksheet, r, 2);
    const a = formTSEExcelJsCellText(worksheet, r, 1);
    if (/name\s+of\s+employee/i.test(b) || (/s\.?\s*no/i.test(a) && /name/i.test(b))) {
      headerRow = r;
      break;
    }
  }
  const dataRowFrom = headerRow > 0 ? Math.max(headerRow + 2, 14) : 14;
  unmergeFormTSEWorksheetRows(worksheet, dataRowFrom, rowTo);
  let shiftedRows = 0;
  for (let r = dataRowFrom; r <= rowTo; r += 1) {
    if (formTSERowHasDuplicateIdentityUnderJ(worksheet, r)) {
      for (let c = shift + 1; c <= shift + 9; c += 1) {
        const v = formTSEExcelJsCellText(worksheet, r, c);
        if (/^(P|A|WO|H|L|CL|EL|SL|OD|WOP)$/i.test(v)) continue;
        worksheet.getCell(r, c).value = null;
      }
      shiftedRows += 1;
      continue;
    }
    if (!formTSERowLooksIdentityShiftedToJ(worksheet, r)) {
      if (shiftedRows > 0) {
        const serialA = formTSEExcelJsCellText(worksheet, r, 1);
        const nameB = formTSEExcelJsCellText(worksheet, r, 2);
        const serialJ = formTSEExcelJsCellText(worksheet, r, shift + 1);
        const nameK = formTSEExcelJsCellText(worksheet, r, shift + 2);
        if (!serialA && !nameB && !serialJ && !nameK) break;
      }
      continue;
    }
    const snapshot = [];
    for (let c = 1; c <= maxCol + shift; c += 1) {
      const cell = worksheet.getCell(r, c);
      snapshot[c] = { value: cell.value, numFmt: cell.numFmt };
    }
    for (let c = 1; c <= maxCol; c += 1) {
      const src = snapshot[c + shift];
      const cell = worksheet.getCell(r, c);
      cell.value = src && src.value !== undefined ? src.value : null;
      if (src?.numFmt) cell.numFmt = src.numFmt;
    }
    for (let c = maxCol + 1; c <= maxCol + shift; c += 1) {
      worksheet.getCell(r, c).value = null;
    }
    shiftedRows += 1;
  }
  return shiftedRows;
}

/**
 * True when employee Name/S.NO sit in columns A–B (correct), not only under day-1 at J–K.
 * Used so Download Draft File can stream the Save workbook when alignment is good.
 */
export async function formTSEWorkbookHasIdentityInColumnA(arrayBuffer) {
  if (!arrayBuffer || arrayBuffer.byteLength < 32) return false;
  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(arrayBuffer);
    const worksheet = workbook.worksheets[0];
    if (!worksheet) return false;
    let sawAligned = false;
    let sawShifted = false;
    const rowTo = Math.max(120, Number(worksheet.rowCount) || 0);
    // Ignore Month/Year / Establishment header boxes. Only count names under the table header.
    let headerRow = 0;
    for (let r = 1; r <= Math.min(40, rowTo); r += 1) {
      const a = formTSEExcelJsCellText(worksheet, r, 1);
      const b = formTSEExcelJsCellText(worksheet, r, 2);
      if (
        /name\s+of\s+employee/i.test(b) ||
        (/s\.?\s*no/i.test(a) && /name/i.test(b))
      ) {
        headerRow = r;
        break;
      }
    }
    const scanFrom = headerRow > 0 ? headerRow + 1 : 4;
    for (let r = scanFrom; r <= rowTo; r += 1) {
      if (formTSERowLooksIdentityShiftedToJ(worksheet, r)) {
        if (!sawAligned) {
          sawShifted = true;
          break;
        }
        continue;
      }
      const nameB = formTSEExcelJsCellText(worksheet, r, 2);
      const serialA = formTSEExcelJsCellText(worksheet, r, 1);
      if (looksLikeFormTSEEmployeeNameCell(nameB)) {
        sawAligned = true;
      } else if (/^\d{1,4}$/.test(serialA) && looksLikeFormTSEEmployeeNameCell(nameB)) {
        sawAligned = true;
      }
    }
    if (sawShifted) return false;
    if (sawAligned && formTSEWorksheetHasNameFatherSwap(worksheet)) {
      return false;
    }
    return sawAligned;
  } catch (_) {
    return false;
  }
}

/**
 * Form T_KA templates often freeze at column J (attendance) with topLeftCell J15,
 * so Excel opens scrolled to day 1 instead of S.NO. Always open at column A.
 */
export function resetFormTSEWorksheetOpenAtColumnA(worksheet) {
  if (!worksheet) return;
  for (let c = 1; c <= 9; c += 1) {
    const col = worksheet.getColumn(c);
    col.hidden = false;
    const width = Number(col.width);
    if (!Number.isFinite(width) || width < 4) {
      col.width = c === 1 ? 18 : c === 2 ? 24 : 14;
    } else if (c === 1 && width < 16) {
      col.width = 18;
    }
  }
  worksheet.views = [
    {
      state: 'normal',
      xSplit: 0,
      ySplit: 0,
      topLeftCell: 'A1',
      activeCell: 'A1',
      showGridLines: true,
      zoomScale: 100,
      zoomScaleNormal: 100,
      workbookViewId: 0,
    },
  ];
}

/** Load a Form T workbook, open the view at A1, and return new bytes. */
export async function applyFormTSEWorkbookOpenAtColumnA(arrayBuffer) {
  if (!arrayBuffer || arrayBuffer.byteLength < 32) return arrayBuffer;
  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(arrayBuffer);
    (workbook.worksheets || []).forEach((ws) => resetFormTSEWorksheetOpenAtColumnA(ws));
    resetExcelJsWorkbookActiveSheet(workbook);
    return workbook.xlsx.writeBuffer();
  } catch (_) {
    return arrayBuffer;
  }
}

/**
 * Shift J-dumped identity back to A–I when needed, then open the sheet at column A.
 */
export async function prepareFormTSEWorkbookForDownload(arrayBuffer, borderHints = {}) {
  if (!arrayBuffer || arrayBuffer.byteLength < 32) return arrayBuffer;
  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(arrayBuffer);
    const sheets = workbook.worksheets || [];
    if (sheets.length === 0) return arrayBuffer;
    const headers = borderHints.headers || borderHints.headersToUse || [];
    const rows = overlayFormTSEPeopleNamesOntoRows(
      borderHints.mappedData || borderHints.rows || [],
      headers,
      borderHints.employees || [],
      borderHints.helpers || {}
    );
    const headerFormData =
      borderHints.headerFormData && typeof borderHints.headerFormData === 'object'
        ? borderHints.headerFormData
        : null;
    const parsedFormHeader = borderHints.parsedFormHeader || null;
    sheets.forEach((worksheet) => {
      shiftFormTSEWorksheetIdentityFromJToA(worksheet);
      if (headerFormData) {
        writeFormTSEHeaderFieldsToWorksheet(worksheet, headerFormData, parsedFormHeader);
        expandFormTSEHeaderValueBoxes(worksheet, FORM_T_KARNATAKA_HEADER_BOX_END_COL);
        writeFormTSEHeaderFieldsToWorksheet(worksheet, headerFormData, parsedFormHeader);
      }
      restoreFormTSEKarnatakaEmployeeNameColumnHeader(worksheet);
      if (Array.isArray(rows) && rows.length > 0) {
        stampFormTSEKarnatakaEmployeeIdentity(
          worksheet,
          rows,
          headers,
          borderHints.employees || []
        );
      }
      resetFormTSEWorksheetOpenAtColumnA(worksheet);
    });
    resetExcelJsWorkbookActiveSheet(workbook);
    let repaired = await workbook.xlsx.writeBuffer();
    repaired = await stripFormTSEDataIdentityMergesFromXlsx(repaired);
    const verifyWb = new ExcelJS.Workbook();
    await verifyWb.xlsx.load(repaired);
    const needsRestamp = (verifyWb.worksheets || []).some((ws) =>
      formTSEWorksheetHasNameFatherSwap(ws)
    );
    if (needsRestamp && Array.isArray(rows) && rows.length > 0) {
      verifyWb.worksheets.forEach((worksheet) => {
        unmergeAllFormTSEDataIdentityBands(worksheet);
        stampFormTSEKarnatakaEmployeeIdentity(
          worksheet,
          rows,
          headers,
          borderHints.employees || []
        );
        resetFormTSEWorksheetOpenAtColumnA(worksheet);
      });
      resetExcelJsWorkbookActiveSheet(verifyWb);
      repaired = await verifyWb.xlsx.writeBuffer();
      repaired = await stripFormTSEDataIdentityMergesFromXlsx(repaired);
    }
    const opened = await applyFormTSEWorkbookOpenAtColumnA(repaired);
    const bordered = await applyFormTSEKarnatakaExportBordersToBuffer(opened, borderHints);
    if (Array.isArray(rows) && rows.length > 0) {
      return forceFormTSEIdentityCellsInSheetXml(
        bordered,
        rows,
        headers,
        borderHints.employees || []
      );
    }
    return bordered;
  } catch (err) {
    console.warn('Form T Karnataka download prepare failed:', err);
    const repaired = await repairFormTSEWorkbookColumnAlignment(arrayBuffer);
    const opened = await applyFormTSEWorkbookOpenAtColumnA(repaired);
    const bordered = await applyFormTSEKarnatakaExportBordersToBuffer(opened, borderHints);
    if (Array.isArray(borderHints.mappedData || borderHints.rows) &&
        (borderHints.mappedData || borderHints.rows).length > 0) {
      return forceFormTSEIdentityCellsInSheetXml(
        bordered,
        borderHints.mappedData || borderHints.rows,
        borderHints.headers || borderHints.headersToUse || [],
        borderHints.employees || []
      );
    }
    return bordered;
  }
}

/**
 * If employee rows were written under column J (attendance) instead of A, shift each
 * data row left by 9 columns so S.NO/Name land on A–B and attendance returns to J+.
 * No-op when already aligned. Returns a new ArrayBuffer.
 */
export async function repairFormTSEWorkbookColumnAlignment(arrayBuffer) {
  if (!arrayBuffer || arrayBuffer.byteLength < 32) return arrayBuffer;
  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(arrayBuffer);
    const sheets = workbook.worksheets || [];
    if (sheets.length === 0) return arrayBuffer;
    let shiftedRows = 0;
    sheets.forEach((worksheet) => {
      shiftedRows += shiftFormTSEWorksheetIdentityFromJToA(worksheet);
      resetFormTSEWorksheetOpenAtColumnA(worksheet);
    });
    resetExcelJsWorkbookActiveSheet(workbook);
    if (shiftedRows === 0) {
      // Still rewrite so view/unhide changes persist even when rows were already aligned.
      return workbook.xlsx.writeBuffer();
    }
    return workbook.xlsx.writeBuffer();
  } catch (err) {
    console.warn('Form T column alignment repair failed:', err);
    return arrayBuffer;
  }
}

/**
 * Locate the Form T identity table start (S.NO / Name) on the sheet.
 * Never trust a modal hint that points under the ATTENDANCE band (col J+).
 */
export function findFormTSEIdentityTableStartCol0({
  getMergedAwareCellText,
  headerRowIndex = -1,
  maxScanCols = 40,
  preferCol = 0,
} = {}) {
  const read = typeof getMergedAwareCellText === 'function' ? getMergedAwareCellText : () => '';
  const maxCol = Math.max(14, Math.min(Number(maxScanCols) || 40, 40));
  const rowsToScan = [];
  if (Number(headerRowIndex) >= 0) rowsToScan.push(Number(headerRowIndex));
  for (let r = 8; r <= 20; r += 1) {
    if (!rowsToScan.includes(r)) rowsToScan.push(r);
  }
  for (let ri = 0; ri < rowsToScan.length; ri += 1) {
    const r = rowsToScan[ri];
    // Only search A–I — never treat ATTENDANCE-band text as the identity start.
    for (let c = 0; c < Math.min(maxCol, FORM_T_KA_ATTENDANCE_START_COL0); c += 1) {
      const t = String(read(r, c) || '').toLowerCase();
      if (!t) continue;
      if (
        /s\.?\s*no|sl\.?\s*no|sr\.?\s*no|serial/.test(t) ||
        isFormTSESerialNumberHeader(t) ||
        isFormTSEEmployeeNameHeader(t)
      ) {
        return c;
      }
    }
  }
  const preferred = Number(preferCol);
  if (Number.isFinite(preferred) && preferred >= 0 && preferred < FORM_T_KA_ATTENDANCE_START_COL0) {
    return preferred;
  }
  return FORM_T_KA_IDENTITY_START_COL0;
}

/**
 * Map Form T write headers → 0-based Excel columns.
 * Identity always lands on A–I; ATTENDANCE_d always lands on column J+(d-1).
 * This prevents a wrong tableStartCol from dumping employee rows under day 1+.
 */
export function resolveFormTSEWriteExcelCols0(writeHeaders, options = {}) {
  const headers = Array.isArray(writeHeaders) ? writeHeaders : [];
  const identityStart = Math.min(
    FORM_T_KA_ATTENDANCE_START_COL0 - 1,
    Math.max(0, Number(options.identityStartCol0) || FORM_T_KA_IDENTITY_START_COL0)
  );
  const attendanceStart =
    Number(options.attendanceStartCol0) >= FORM_T_KA_ATTENDANCE_START_COL0
      ? Number(options.attendanceStartCol0)
      : FORM_T_KA_ATTENDANCE_START_COL0;

  const identitySlot = (header) => {
    if (isFormTSESerialNumberHeader(header)) return 0;
    if (isFormTSEEmployeeNameHeader(header)) return 1;
    if (isFormTSEFatherHusbandHeader(header)) return 2;
    if (isFormTSEGenderHeader(header)) return 3;
    if (isFormTSEDesignationHeader(header)) return 4;
    if (isFormTSEDateOfJoiningHeader(header)) return 5;
    if (isFormTSEEsicRegistrationHeader(header)) return 6;
    if (isFormTSEUanRegistrationHeader(header)) return 7;
    if (isFormTSEWagesFixedIncludingVDAHeader(header)) return 8;
    return -1;
  };

  let nextLooseIdentity = 0;
  let postAttendanceIdx = 0;
  const firstAttIdx = headers.findIndex((h) => isFormTSEAttendanceDayHeader(h));

  return headers.map((header, idx) => {
    if (isFormTSEAttendanceDayHeader(header)) {
      const day = parseInt(String(header).match(/^ATTENDANCE_(\d{1,2})/i)[1], 10);
      return attendanceStart + Math.max(0, day - 1);
    }
    const slot = identitySlot(header);
    if (slot >= 0) return identityStart + slot;
    // Leading non-attendance headers without a known slot (still before attendance band).
    if (firstAttIdx < 0 || idx < firstAttIdx) {
      while (
        nextLooseIdentity <= 8 &&
        headers.some(
          (h, hi) =>
            hi < idx &&
            !isFormTSEAttendanceDayHeader(h) &&
            identitySlot(h) === nextLooseIdentity
        )
      ) {
        nextLooseIdentity += 1;
      }
      const col = identityStart + Math.min(nextLooseIdentity, 8);
      nextLooseIdentity += 1;
      return col;
    }
    // Wage / deduction columns after the attendance day band → after day 31.
    const col = attendanceStart + 31 + postAttendanceIdx;
    postAttendanceIdx += 1;
    return col;
  });
}

/**
 * Prefer the ORIGINAL template header order for download column positions.
 * Modal headers (possibly truncated) supply values; template headers supply layout.
 */
export function resolveFormTSEDownloadWriteHeaders(inputHeaders, templateHeaders) {
  const input = (Array.isArray(inputHeaders) ? inputHeaders : []).filter((h) => String(h || '').trim());
  const template = (Array.isArray(templateHeaders) ? templateHeaders : []).filter((h) =>
    String(h || '').trim()
  );
  const hasIdentity = (hdrs) =>
    (Array.isArray(hdrs) ? hdrs : []).some(
      (h) => isFormTSESerialNumberHeader(h) || isFormTSEEmployeeNameHeader(h)
    );
  const attCount = (hdrs) => listFormTSEAttendanceDayHeaders(hdrs).length;
  // Use the fuller original template layout when it has identity + a real attendance band.
  if (
    template.length >= 4 &&
    hasIdentity(template) &&
    attCount(template) >= 3 &&
    (template.length >= input.length || !hasIdentity(input) || attCount(input) < 3)
  ) {
    return template;
  }
  if (hasIdentity(input)) return input;
  if (hasIdentity(template)) return template;
  return input.length >= 4 ? input : template;
}

/** Keep UAN / ESI as plain text so Excel does not show 1.01E+11. */
export function formatFormTSERegistrationIdForExcel(value) {
  if (value == null || value === '') return '';
  if (typeof value === 'number' && Number.isFinite(value)) {
    return Math.round(value).toLocaleString('en-US', {
      useGrouping: false,
      maximumFractionDigits: 0,
    });
  }
  let s = String(value).trim();
  if (!s) return '';
  if (/^\d+(\.\d+)?e[+\-]?\d+$/i.test(s)) {
    const n = Number(s);
    if (Number.isFinite(n)) {
      return Math.round(n).toLocaleString('en-US', {
        useGrouping: false,
        maximumFractionDigits: 0,
      });
    }
  }
  return s.replace(/\.0+$/, '');
}

/**
 * Map modal/autofill row values onto original template headers by label (never by
 * column index — attendance-first modal order must not dump "P" into S.NO).
 */
export function alignFormTSERowsToTemplateHeaders(rows, fromHeaders, toHeaders) {
  const from = Array.isArray(fromHeaders) ? fromHeaders : [];
  const to = Array.isArray(toHeaders) ? toHeaders : [];
  return (Array.isArray(rows) ? rows : []).map((row) => {
    if (!row || typeof row !== 'object') return row;
    const out = { ...row };
    to.forEach((toH) => {
      if (!toH) return;
      const existing = getFormTSEKarnatakaRowValueForHeader(row, toH);
      if (existing != null && String(existing).trim() !== '') {
        out[toH] = existing;
        return;
      }
      const target = formTSEEmployeeHeaderKeyNorm(toH);
      if (!target) return;
      const fromH = from.find((h) => formTSEEmployeeHeaderKeyNorm(h) === target);
      if (!fromH) return;
      const v = row[fromH];
      if (v != null && String(v).trim() !== '') out[toH] = v;
    });
    return out;
  });
}

/** Write Form T grid rows into the template at exact column positions (ExcelJS, 1-based). */
export async function buildFormTSEWorkbookWithTemplateStyles({
  templateArrayBuffer,
  mappedData,
  headersToUse,
  parsedHeaderRowIndex = -1,
  parsedDataStartIndex = -1,
  parsedTableStartCol = 0,
  parsedFormHeader,
  headerFormData = {},
  headerSiteContext = {},
  formFileName,
  currentItem = null,
  sheetNameHint = '',
  employees = [],
}) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(templateArrayBuffer);
  const worksheet =
    (sheetNameHint && workbook.getWorksheet(sheetNameHint)) || workbook.worksheets[0];
  if (!worksheet) throw new Error('Template worksheet not found.');

  const excelCellValueToString = (val) => {
    if (val == null) return '';
    if (typeof val === 'string' || typeof val === 'number' || typeof val === 'boolean') {
      return String(val);
    }
    if (val instanceof Date) return val.toISOString();
    if (typeof val === 'object') {
      if (Array.isArray(val.richText)) return val.richText.map((rt) => rt?.text || '').join('');
      if (val.text != null) return String(val.text);
      if (val.result != null) return String(val.result);
    }
    return '';
  };

  const mergeTopLeftCache = new Map();
  const getMergeTopLeft = (r, c) => {
    const key = `${r}:${c}`;
    if (mergeTopLeftCache.has(key)) return mergeTopLeftCache.get(key);
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
    mergeTopLeftCache.set(key, topLeft);
    return topLeft;
  };
  const getMergedAwareCellText = (r, c) => {
    const tl = getMergeTopLeft(r, c);
    return excelCellValueToString(worksheet.getCell(tl.r, tl.c)?.value).trim();
  };

  const maxScanRows = Math.max(
    (Number(parsedDataStartIndex) >= 0 ? Number(parsedDataStartIndex) + 30 : 0),
    (Number(parsedHeaderRowIndex) >= 0 ? Number(parsedHeaderRowIndex) + 18 : 0),
    50,
    worksheet.rowCount + 5
  );
  // Full Form T (identity + 31 days + wage/deduction band) needs ~70+ columns.
  const maxScanCols = Math.max(worksheet.columnCount + 5, 80);
  const jsonData = [];
  for (let r = 1; r <= maxScanRows; r += 1) {
    const row = [];
    for (let c = 1; c <= maxScanCols; c += 1) {
      row.push(getMergedAwareCellText(r, c));
    }
    jsonData.push(row);
  }
  const getMergedAwareCellText0 = (r0, c0) => getMergedAwareCellText(r0 + 1, c0 + 1);
  const rawCell0 = (r0, c0) => getMergedAwareCellText0(r0, c0);

  const headerRowIndex =
    Number(parsedHeaderRowIndex) >= 0
      ? Number(parsedHeaderRowIndex)
      : rebuildFormTSETableHeadersFromSheet({
          headerRowIndex: -1,
          getMergedAwareCellText: getMergedAwareCellText0,
          getRawCellText: rawCell0,
          jsonData,
          effectiveSheetCols: maxScanCols,
          tableStartCol: parsedTableStartCol,
          dataStartIndex: parsedDataStartIndex,
        })?.headerRowIndex ?? -1;

  const rebuilt = rebuildFormTSETableHeadersFromSheet({
    headerRowIndex,
    getMergedAwareCellText: getMergedAwareCellText0,
    getRawCellText: rawCell0,
    jsonData,
    effectiveSheetCols: maxScanCols,
    tableStartCol: parsedTableStartCol,
    dataStartIndex: parsedDataStartIndex,
  });

  const inputHeaders = Array.isArray(headersToUse) ? headersToUse.filter((h) => String(h || '').trim()) : [];
  const rebuiltHeaders =
    rebuilt?.headers?.length >= 4
      ? preserveFormTSEAttendanceHeadersIfNeeded(
          inputHeaders,
          normalizeFormTSETableHeaders(rebuilt.headers)
        )
      : [];
  const effectiveHeaders =
    rebuiltHeaders.length >= 4
      ? rebuiltHeaders
      : inputHeaders;
  if (effectiveHeaders.length < 4) {
    throw new Error('Could not locate Form T table columns.');
  }

  // Form T KA layout is fixed: A–I identity, J+ attendance. Never use a modal
  // tableStartCol under ATTENDANCE (that dumped S.NO/Name into column J).
  const tableStartCol = FORM_T_KA_IDENTITY_START_COL0;

  const headersHaveIdentity = (hdrs) =>
    (Array.isArray(hdrs) ? hdrs : []).some(
      (h) => isFormTSESerialNumberHeader(h) || isFormTSEEmployeeNameHeader(h)
    );

  // Prefer ORIGINAL template column layout so download values fit the sheet
  // (modal headers can be truncated after Medical Allowance).
  let writeHeaders = resolveFormTSEDownloadWriteHeaders(inputHeaders, effectiveHeaders);
  if (!headersHaveIdentity(writeHeaders) && headersHaveIdentity(inputHeaders)) {
    // Fallback: keep modal identity headers if template rebuild lost them.
    writeHeaders = [...inputHeaders];
  }

  const excelCols0 = resolveFormTSEWriteExcelCols0(writeHeaders, {
    identityStartCol0: FORM_T_KA_IDENTITY_START_COL0,
    attendanceStartCol0: FORM_T_KA_ATTENDANCE_START_COL0,
  });
  // Absolute lock: first identity fields always A–I (1–9), never J (10)+.
  const fieldCols = excelCols0.map((c0, idx) => {
    const header = writeHeaders[idx];
    if (isFormTSEAttendanceDayHeader(header)) {
      return Number(c0) + 1;
    }
    if (
      isFormTSESerialNumberHeader(header) ||
      isFormTSEEmployeeNameHeader(header) ||
      isFormTSEFatherHusbandHeader(header) ||
      isFormTSEGenderHeader(header) ||
      isFormTSEDesignationHeader(header) ||
      isFormTSEDateOfJoiningHeader(header) ||
      isFormTSEEsicRegistrationHeader(header) ||
      isFormTSEUanRegistrationHeader(header) ||
      isFormTSEWagesFixedIncludingVDAHeader(header)
    ) {
      const locked = Number(c0) + 1;
      return locked >= 1 && locked <= FORM_T_KA_ATTENDANCE_START_COL0 ? locked : idx + 1;
    }
    const col1 = Number(c0) + 1;
    // Non-attendance before day band must stay in A–I when index is in the identity range.
    if (idx < FORM_T_KA_ATTENDANCE_START_COL0 && col1 > FORM_T_KA_ATTENDANCE_START_COL0) {
      return idx + 1;
    }
    return col1;
  });

  let startRow =
    rebuilt?.dataStartIndex != null && rebuilt.dataStartIndex >= 0
      ? rebuilt.dataStartIndex + 1
      : Number(parsedDataStartIndex) >= 0
        ? Number(parsedDataStartIndex) + 1
        : headerRowIndex >= 0
          ? headerRowIndex + 3
          : 15;
  // Employee data must sit below header / calendar-day / statutory index strips.
  if (rebuilt?.columnIndexRow != null && rebuilt.columnIndexRow >= 0) {
    startRow = Math.max(startRow, rebuilt.columnIndexRow + 2);
  }
  if (rebuilt?.attendanceDayRow != null && rebuilt.attendanceDayRow >= 0) {
    startRow = Math.max(startRow, rebuilt.attendanceDayRow + 2);
  }
  if (
    Number(parsedDataStartIndex) >= 0 &&
    Number(parsedDataStartIndex) + 1 > startRow &&
    !(rebuilt?.columnIndexRow >= 0 && Number(parsedDataStartIndex) <= rebuilt.columnIndexRow)
  ) {
    startRow = Number(parsedDataStartIndex) + 1;
  }
  // Never write onto the statutory column-index strip (1,2,3…) under the header row.
  const sheetLooksLikeIndexRow = (excelRow1Based) => {
    const r0 = excelRow1Based - 1;
    if (r0 < 0) return false;
    return rowLooksLikeFormTColumnIndexRow(
      r0,
      tableStartCol,
      maxScanCols,
      getMergedAwareCellText0,
      rawCell0
    );
  };
  for (let guard = 0; guard < 5 && sheetLooksLikeIndexRow(startRow); guard += 1) {
    startRow += 1;
  }
  // Form T body is immediately under the 1,2,3… strip (Excel row 15 on the KA template).
  // A stale parsedDataStartIndex / "Date:" footer must not push names below the grid.
  const stripEndExcel = Math.max(
    rebuilt?.columnIndexRow != null && rebuilt.columnIndexRow >= 0 ? rebuilt.columnIndexRow + 2 : 0,
    rebuilt?.attendanceDayRow != null && rebuilt.attendanceDayRow >= 0
      ? rebuilt.attendanceDayRow + 2
      : 0,
    headerRowIndex >= 0 ? headerRowIndex + 3 : 0,
    14
  );
  if (startRow < stripEndExcel) startRow = stripEndExcel;
  let dateFooterRow = 0;
  for (let r = stripEndExcel; r <= Math.min(stripEndExcel + 16, maxScanRows); r += 1) {
    const t = String(getMergedAwareCellText(r, 1) || '')
      .replace(/\s+/g, ' ')
      .trim()
      .toLowerCase();
    if (/^date\s*:?$/.test(t) || /^date\s*:/.test(t)) {
      dateFooterRow = r;
      break;
    }
  }
  if (dateFooterRow > stripEndExcel && startRow >= dateFooterRow) {
    startRow = dateFooterRow - 1;
  }
  if (startRow > stripEndExcel + 2) {
    startRow = stripEndExcel;
  }
  if (sheetLooksLikeIndexRow(startRow)) startRow += 1;

  const normalize = (txt) =>
    String(txt || '')
      .replace(/\r?\n/g, ' ')
      .replace(/\s+/g, ' ')
      .trim()
      .toLowerCase()
      .replace(/[^a-z0-9 ]/g, '');

  const getRowValueForHeader = (row, header, headerIndex, allHeaders) => {
    if (!row || typeof row !== 'object' || Array.isArray(row)) {
      return Array.isArray(row) ? row[headerIndex] : '';
    }
    if (isFormTSEAttendanceDayHeader(header)) {
      return sanitizeFormTSEAttendanceMark(getFormTSEKarnatakaRowValueForHeader(row, header));
    }
    const byFormTKey = getFormTSEKarnatakaRowValueForHeader(row, header);
    if (byFormTKey != null && String(byFormTKey).trim() !== '') return byFormTKey;
    if (Object.prototype.hasOwnProperty.call(row, header)) return row[header];
    const target = normalize(header);
    if (!target) return '';
    const rowKeys = Object.keys(row).filter((k) => !String(k || '').startsWith('__'));
    const exact = rowKeys.filter((k) => normalize(k) === target);
    if (exact.length === 1) return row[exact[0]];
    if (exact.length > 1) {
      const occur =
        allHeaders.slice(0, headerIndex + 1).filter((h) => normalize(h) === target).length - 1;
      return row[exact[Math.min(Math.max(occur, 0), exact.length - 1)]];
    }
    return '';
  };

  const rowLooksMeaningful = (row) => {
    if (Array.isArray(row)) {
      return row.some((v) => v != null && String(v).trim() !== '');
    }
    if (row && typeof row === 'object') {
      return Object.values(row).some((v) => v != null && String(v).trim() !== '');
    }
    return false;
  };

  const sourceHeaders =
    inputHeaders.length > 0 ? inputHeaders : writeHeaders;
  let sourcePrimary = prepareFormTSEExportRows(
    // Align modal/autofill rows onto the original template header keys by label.
    alignFormTSERowsToTemplateHeaders(
      Array.isArray(mappedData) ? mappedData : [],
      sourceHeaders,
      writeHeaders
    ),
    writeHeaders,
    writeHeaders
  ).filter((row) => rowLooksMeaningful(row));

  if (
    sourcePrimary.length === 0 &&
    Array.isArray(mappedData) &&
    mappedData.length > 0 &&
    formTSEDownloadHasSubstantiveRows(mappedData, writeHeaders)
  ) {
    sourcePrimary = prepareFormTSEExportRows(mappedData, writeHeaders, writeHeaders).filter((row) =>
      rowLooksMeaningful(row)
    );
  }

  if (
    sourcePrimary.length === 0 &&
    Array.isArray(mappedData) &&
    mappedData.length > 0 &&
    formTSEDownloadHasSubstantiveRows(mappedData, sourceHeaders)
  ) {
    sourcePrimary = prepareFormTSEExportRows(
      alignFormTSERowsToTemplateHeaders(mappedData, sourceHeaders, writeHeaders),
      writeHeaders,
      writeHeaders
    ).filter((row) => rowLooksMeaningful(row));
  }

  // Last resort: modal/export headers may not match rebuilt sheet headers — still pull names by Form T key-norm.
  if (
    sourcePrimary.length === 0 &&
    Array.isArray(mappedData) &&
    mappedData.length > 0 &&
    formTSEDownloadHasSubstantiveRows(mappedData, [])
  ) {
    sourcePrimary = prepareFormTSEExportRows(mappedData, writeHeaders, writeHeaders).filter((row) =>
      rowLooksMeaningful(row)
    );
  }

  if (sourcePrimary.length === 0 && Array.isArray(mappedData) && mappedData.length > 0) {
    sourcePrimary = mappedData
      .filter((row) => row && typeof row === 'object')
      .map((row, rowIndex) => {
        const out = {};
        writeHeaders.forEach((header, colIndex) => {
          let value = getFormTSEKarnatakaRowValueForHeader(row, header);
          if (
            (value == null || String(value).trim() === '') &&
            /s\.?\s*no|serial|sl\.?\s*no/i.test(String(header || ''))
          ) {
            value = rowIndex + 1;
          }
          out[header] = value != null ? value : '';
        });
        return out;
      })
      .filter((row) => rowLooksMeaningful(row));
  }

  if (
    !formTSEDownloadHasSubstantiveRows(sourcePrimary, writeHeaders) &&
    Array.isArray(employees) &&
    employees.length > 0 &&
    writeHeaders.length > 0
  ) {
    sourcePrimary = resolveFormTSERowsForExport({
      liveRows: sourcePrimary,
      headers: writeHeaders,
      employees,
    }).filter((row) => rowLooksMeaningful(row));
  }

  // Insert rows above Date:/signature and unmerge B:C leftovers so the last
  // employee name is not overwritten by Father / Husband's Name.
  dateFooterRow = ensureFormTSEKarnatakaDataRowCapacity(
    worksheet,
    startRow,
    Math.max(sourcePrimary.length, 1),
    dateFooterRow
  );
  mergeTopLeftCache.clear();

  const tableColMin = fieldCols.length > 0 ? Math.min(...fieldCols) : 1;
  const tableColMax = fieldCols.length > 0 ? Math.max(...fieldCols) : writeHeaders.length;
  const templateBodyRows = countExcelJSTemplateBodyRows(worksheet, startRow, tableColMin, tableColMax);
  const bodyRowsToPaint = Math.max(sourcePrimary.length, templateBodyRows);
  // Wipe A…wage band on body rows first. Dirty Draft templates keep a prior J-shifted
  // write; without this, names stay under ATTENDANCE even after a correct A–I rewrite.
  const clearThroughCol = Math.max(tableColMax + 40, FORM_T_KA_ATTENDANCE_START_COL0 + 50, 80);
  const clearFromRow = Math.min(startRow, 14);
  let clearToRow = Math.max(
    startRow + Math.max(bodyRowsToPaint, sourcePrimary.length, 1) - 1,
    19
  );
  if (dateFooterRow > 14) {
    clearToRow = Math.min(clearToRow, dateFooterRow - 1);
  }
  clearFormTSEExcelJsDataRowBand(
    worksheet,
    clearFromRow,
    clearToRow,
    clearThroughCol
  );

  clearExcelJSTrailingTableCells(worksheet, {
    dataStartRow: startRow,
    dataRowCount: bodyRowsToPaint,
    afterCol: tableColMax,
    throughCol: tableColMax + 30,
  });

  const writeEmployeeRow = (row, i, cols1Based) => {
    const excelRow = startRow + i;
    const safeRow =
      row && typeof row === 'object' && !Array.isArray(row)
        ? clearFormTSEKarnatakaPayrollIfNoAttendance({ ...row }, writeHeaders)
        : row;
    unmergeFormTSEIdentityBandOnRow(worksheet, excelRow);
    for (let j = 0; j < writeHeaders.length; j += 1) {
      const header = writeHeaders[j];
      let value = getRowValueForHeader(safeRow, header, j, writeHeaders);
      if (isFormTSEAttendanceDayHeader(header)) {
        value = sanitizeFormTSEAttendanceMark(value);
      }
      if (isFormTSEEmployeeNameHeader(header)) {
        const named = getFormTSEKarnatakaEmployeeNameFromRow(safeRow, writeHeaders);
        if (named) value = named;
      }
      if (
        (value == null || String(value).trim() === '') &&
        /s\.?\s*no|serial|sl\.?\s*no/i.test(String(header || ''))
      ) {
        value = i + 1;
      }
      // Total OT hours — static NIL default on every export write.
      if (isFormTSEKarnatakaTotalOtHoursHeader(header)) {
        value = FORM_T_KA_OT_HOURS_NIL;
      }
      let targetCol = cols1Based[j];
      if (!targetCol || targetCol < 1) continue;
      if (isFormTSESerialNumberHeader(header)) targetCol = 1;
      else if (isFormTSEEmployeeNameHeader(header)) targetCol = 2;
      else if (isFormTSEFatherHusbandHeader(header)) targetCol = 3;
      else if (isFormTSEGenderHeader(header)) targetCol = 4;
      else if (isFormTSEDesignationHeader(header)) targetCol = 5;
      else if (isFormTSEDateOfJoiningHeader(header)) targetCol = 6;
      else if (isFormTSEEsicRegistrationHeader(header)) targetCol = 7;
      else if (isFormTSEUanRegistrationHeader(header)) targetCol = 8;
      else if (isFormTSEWagesFixedIncludingVDAHeader(header)) targetCol = 9;
      // Identity band must stay on A–I. Merged-cell top-left can otherwise redirect into J+.
      const isIdentityWrite =
        isFormTSESerialNumberHeader(header) ||
        isFormTSEEmployeeNameHeader(header) ||
        isFormTSEFatherHusbandHeader(header) ||
        isFormTSEGenderHeader(header) ||
        isFormTSEDesignationHeader(header) ||
        isFormTSEDateOfJoiningHeader(header) ||
        isFormTSEEsicRegistrationHeader(header) ||
        isFormTSEUanRegistrationHeader(header) ||
        isFormTSEWagesFixedIncludingVDAHeader(header) ||
        (j < FORM_T_KA_ATTENDANCE_START_COL0 && !isFormTSEAttendanceDayHeader(header));
      if (isIdentityWrite && targetCol > 9) {
        targetCol = Math.min(j + 1, 9);
      }
      let cell;
      if (isIdentityWrite) {
        // Write the physical A–I cell; do not follow merges into the attendance band.
        cell = worksheet.getCell(excelRow, targetCol);
      } else {
        const tl = getMergeTopLeft(excelRow, targetCol);
        // If merge master sits under attendance but we aimed left of J, keep the aimed column.
        if (tl.c >= FORM_T_KA_ATTENDANCE_START_COL0 + 1 && targetCol <= FORM_T_KA_ATTENDANCE_START_COL0) {
          cell = worksheet.getCell(excelRow, targetCol);
        } else {
          cell = worksheet.getCell(tl.r, tl.c);
        }
      }
      if (value == null || value === '') {
        cell.value = '';
        continue;
      }
      const asTextRaw = String(value).trim();
      // Keep long registration numbers as text (avoid Excel 1.01E+11 scientific notation).
      const forceText =
        isFormTSEUanRegistrationHeader(header) ||
        isFormTSEEsicRegistrationHeader(header) ||
        /^\d{10,}$/.test(asTextRaw) ||
        /^\d+(\.\d+)?e[+\-]?\d+$/i.test(asTextRaw);
      if (forceText) {
        const asText = formatFormTSERegistrationIdForExcel(value);
        cell.numFmt = '@';
        cell.value = asText;
      } else if (
        typeof value === 'number' ||
        (typeof value === 'string' && /^-?\d+(\.\d+)?$/.test(asTextRaw))
      ) {
        cell.value = Number(value);
      } else {
        cell.value = asTextRaw;
      }
      cell.font = { ...(cell.font || {}), bold: true };
    }
    const namedLast = getFormTSEKarnatakaEmployeeNameFromRow(row, writeHeaders);
    const fatherLast = writeHeaders.find((h) => isFormTSEFatherHusbandHeader(h));
    const fatherVal = fatherLast
      ? String(getRowValueForHeader(row, fatherLast, writeHeaders.indexOf(fatherLast), writeHeaders) ?? '').trim()
      : '';
    writeFormTSEPhysicalNameAndFather(worksheet, excelRow, namedLast, fatherVal);
  };

  for (let i = 0; i < sourcePrimary.length; i += 1) {
    if (i > 0 && i % 15 === 0) {
      await yieldToMain();
    }
    writeEmployeeRow(sourcePrimary[i], i, fieldCols);
  }

  // Final hard pass: copy identity values into A–I and clear mistaken identity dumps from J–R.
  if (sourcePrimary.length > 0) {
    const identityHeaders = [
      writeHeaders.find((h) => isFormTSESerialNumberHeader(h)),
      writeHeaders.find((h) => isFormTSEEmployeeNameHeader(h)),
      writeHeaders.find((h) => isFormTSEFatherHusbandHeader(h)),
      writeHeaders.find((h) => isFormTSEGenderHeader(h)),
      writeHeaders.find((h) => isFormTSEDesignationHeader(h)),
      writeHeaders.find((h) => isFormTSEDateOfJoiningHeader(h)),
      writeHeaders.find((h) => isFormTSEEsicRegistrationHeader(h)),
      writeHeaders.find((h) => isFormTSEUanRegistrationHeader(h)),
      writeHeaders.find((h) => isFormTSEWagesFixedIncludingVDAHeader(h)),
    ];
    for (let i = 0; i < sourcePrimary.length; i += 1) {
      const row = sourcePrimary[i];
      const excelRow = startRow + i;
      unmergeFormTSEIdentityBandOnRow(worksheet, excelRow);
      for (let slot = 0; slot < identityHeaders.length; slot += 1) {
        const header = identityHeaders[slot];
        let value = header
          ? getRowValueForHeader(row, header, writeHeaders.indexOf(header), writeHeaders)
          : '';
        if (
          (value == null || String(value).trim() === '') &&
          slot === 0
        ) {
          value = i + 1;
        }
        // Fallback by leading column order when header detectors miss modal labels.
        if ((value == null || String(value).trim() === '') && writeHeaders[slot] && slot < 9) {
          value = getRowValueForHeader(row, writeHeaders[slot], slot, writeHeaders);
        }
        if (slot === 1) {
          const named = getFormTSEKarnatakaEmployeeNameFromRow(row, writeHeaders);
          if (named) value = named;
        }
        const cell = worksheet.getCell(excelRow, slot + 1);
        if (value == null || String(value).trim() === '') {
          // keep blank
        } else {
          const asTextRaw = String(value).trim();
          const forceText =
            slot === 6 ||
            slot === 7 ||
            /^\d{10,}$/.test(asTextRaw) ||
            /^\d+(\.\d+)?e[+\-]?\d+$/i.test(asTextRaw);
          if (forceText) {
            const asText = formatFormTSERegistrationIdForExcel(value);
            cell.numFmt = '@';
            cell.value = asText;
          } else if (/^-?\d+(\.\d+)?$/.test(asTextRaw) && slot === 0) {
            cell.value = Number(asTextRaw);
          } else {
            cell.value = asTextRaw;
          }
          cell.font = { ...(cell.font || {}), bold: true };
        }
      }
      const namedLast = getFormTSEKarnatakaEmployeeNameFromRow(row, writeHeaders);
      const fatherLast = identityHeaders[2];
      const fatherVal = fatherLast
        ? String(getRowValueForHeader(row, fatherLast, writeHeaders.indexOf(fatherLast), writeHeaders) ?? '').trim()
        : '';
      writeFormTSEPhysicalNameAndFather(worksheet, excelRow, namedLast, fatherVal);
      // If serial/name were wrongly left under day columns, clear those cells when they
      // duplicate identity (numeric serial in J / person name in K).
      const nameInB = excelCellValueToString(worksheet.getCell(excelRow, 2)?.value).trim();
      const maybeNameInK = excelCellValueToString(
        worksheet.getCell(excelRow, FORM_T_KA_ATTENDANCE_START_COL0 + 2)?.value
      ).trim();
      const maybeSerialInJ = excelCellValueToString(
        worksheet.getCell(excelRow, FORM_T_KA_ATTENDANCE_START_COL0 + 1)?.value
      ).trim();
      if (
        /[a-zA-Z]{2,}/.test(nameInB) &&
        /[a-zA-Z]{2,}/.test(maybeNameInK) &&
        maybeNameInK.toLowerCase().includes(nameInB.toLowerCase().slice(0, 4))
      ) {
        for (let c = FORM_T_KA_ATTENDANCE_START_COL0 + 1; c <= FORM_T_KA_ATTENDANCE_START_COL0 + 9; c += 1) {
          const v = excelCellValueToString(worksheet.getCell(excelRow, c)?.value).trim();
          // Only clear cells that look like shifted identity (not P/A/WO attendance codes).
          if (/^(P|A|WO|H|L|CL|EL|SL|OD|WOP)$/i.test(v)) continue;
          if (/^\d{1,2}$/.test(v) && Number(v) === i + 1 && c === FORM_T_KA_ATTENDANCE_START_COL0 + 1) {
            worksheet.getCell(excelRow, c).value = '';
            continue;
          }
          if (/[a-zA-Z]{3,}/.test(v) && !/^(P|A|WO)/i.test(v)) {
            worksheet.getCell(excelRow, c).value = '';
          }
        }
      } else if (/^\d+$/.test(maybeSerialInJ) && !/[a-zA-Z]{2,}/.test(nameInB)) {
        // Identity still missing in A–I but present from J — shift J–R back to A–I.
        for (let slot = 0; slot < 9; slot += 1) {
          const fromCol = FORM_T_KA_ATTENDANCE_START_COL0 + 1 + slot;
          const raw = worksheet.getCell(excelRow, fromCol).value;
          worksheet.getCell(excelRow, slot + 1).value = raw;
          worksheet.getCell(excelRow, fromCol).value = '';
        }
      }
    }
  }

  if (sourcePrimary.length > 0 && fieldCols.length > 0) {
    ensureExcelJSDataRowsWithBorders(worksheet, {
      dataStartRow: startRow,
      dataRowCount: sourcePrimary.length,
      colFrom: tableColMin,
      colTo: tableColMax,
      templateRow: startRow,
      templateBodyRows: 1,
    });
    clearExcelJSTrailingTableCells(worksheet, {
      dataStartRow: startRow,
      dataRowCount: bodyRowsToPaint,
      afterCol: tableColMax,
      throughCol: tableColMax + 30,
    });
  }

  const headerValues = prepareFormTSEDownloadHeaderData(
    headerFormData && typeof headerFormData === 'object' ? headerFormData : {},
    parsedFormHeader,
    headerSiteContext && typeof headerSiteContext === 'object' ? headerSiteContext : {}
  );
  const headerScanEnd = 11;
  // Put Month/Year, Address, Employer into the widened A–I header boxes (Label : value).
  // Adjacent mode left long values clipped inside the narrow A:D template merges.
  writeStatutoryHeaderFieldsToExcelJsWorksheet(worksheet, {
    headerFormData: headerValues,
    parsedFormHeader,
    // Rows 1–8 are title / Rule 24 citation — never dump Establishment into row 3.
    // Rows 12+ are the table — never rewrite "Name of Employee" as principal employer.
    headerRowStart: 9,
    headerRowEnd: headerScanEnd,
    maxScanCols: 80,
    colRightBound: FORM_T_KARNATAKA_HEADER_BOX_END_COL,
    writeMode: 'combined',
  });
  writeFormTSEHeaderFieldsToWorksheet(worksheet, headerValues, parsedFormHeader);
  expandFormTSEHeaderValueBoxes(worksheet, FORM_T_KARNATAKA_HEADER_BOX_END_COL);
  // Re-apply after expand/finalize-bound merges so Label : value survives on rows 10–11.
  writeFormTSEHeaderFieldsToWorksheet(worksheet, headerValues, parsedFormHeader);
  restoreFormTSEKarnatakaEmployeeNameColumnHeader(worksheet);
  shiftFormTSEWorksheetIdentityFromJToA(worksheet);
  stampFormTSEKarnatakaEmployeeIdentity(worksheet, sourcePrimary, writeHeaders, employees);

  await yieldToMain();
  resetFormTSEWorksheetOpenAtColumnA(worksheet);
  resetExcelJsWorkbookActiveSheet(workbook);
  let out = await workbook.xlsx.writeBuffer();
  // Shift any leftover J-dumped identity back to A–I, open at column A, and paint row-12+ borders.
  const tableHeaderRow = FORM_T_KA_TABLE_BORDER_START_ROW;
  let tableLastRow =
    startRow + Math.max(bodyRowsToPaint, sourcePrimary.length, templateBodyRows, 1) - 1;
  if (dateFooterRow > tableHeaderRow) {
    tableLastRow = Math.min(tableLastRow, dateFooterRow - 1);
  }
  tableLastRow = Math.max(tableLastRow, startRow);
  out = await prepareFormTSEWorkbookForDownload(out, {
    colTo: tableColMax,
    tableLastRow,
    force: true,
    headerFormData: headerValues,
    parsedFormHeader,
    mappedData: sourcePrimary,
    headers: writeHeaders,
    employees,
  });
  const outName =
    formFileName ||
    currentItem?.formName?.replace(/[^a-zA-Z0-9]/g, '_') ||
    parsedFormHeader?.title?.replace(/[^a-zA-Z0-9]/g, '_') ||
    `Form_T_${Date.now()}.xlsx`;
  const blob = new Blob([out], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  return { blob, fileName: outName };
}
