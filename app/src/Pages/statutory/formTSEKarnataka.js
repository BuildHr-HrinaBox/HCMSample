import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import {
  clearExcelJSTrailingTableCells,
  countExcelJSTemplateBodyRows,
  ensureExcelJSDataRowsWithBorders,
} from '../../utils/excelTableBorders';
import { yieldToMain } from '../../utils/statutoryAutofillCache';
import {
  excelCellValueToString,
  writeStatutoryHeaderFieldsToExcelJsWorksheet,
} from '../../utils/statutorySiteCompanyHeaders';
import { personNamesMatch } from './formFKarnataka';

/** Karnataka Form T — Combined Muster Roll cum Register of Wages (attendance day grid). */

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
  if (!defaults) return out;
  if (defaults.basic !== undefined) out.basic = defaults.basic;
  if (defaults.hra) out.hra = defaults.hra;
  if (defaults.earnedTotal) out.grossPay = defaults.earnedTotal;
  if (defaults.deductionsTotal) out.totalDeductions = defaults.deductionsTotal;
  if (defaults.netPayable) out.netPay = defaults.netPayable;
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

export const FORM_T_KARNATAKA_HEADER_SPECS = [
  {
    key: 'form_t_month_year',
    label: 'Month / Year',
    match: /^month\s*\/\s*year$/i,
  },
  {
    key: 'form_t_establishment_name_address',
    label: 'Name and address of the Establishment',
    match: /name\s+and\s+address\s+of\s+the\s+establishment/i,
  },
  {
    key: 'form_t_employer',
    label: 'Name and Address of employer',
    match: /name\s+and\s+address\s+of\s+employer/i,
  },
];

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

export function isFormTSEEmployeeNameHeader(h) {
  const s = formTSEEmployeeHeaderKeyNorm(h);
  if (!s.includes('name')) return false;
  return (
    s.includes('employee') ||
    s.includes('workman') ||
    s.includes('worker') ||
    /nameoftheemployee/.test(s.replace(/\s/g, ''))
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
      const existing = getFormTSEKarnatakaRowValueForHeader(row, header);
      if (String(existing ?? '').trim() !== '') return;
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
      'statutory_establishment_name',
      'statutory_establishment_name_shop',
      'form_q_establishment',
      'form_f_establishment_name_address',
      'form_h_establishment_name_address',
      'form25_establishment',
    ]);
    if (direct) return direct;
  } else if (spec.key === 'form_t_employer') {
    const direct = tryKeys([
      'form_t_employer',
      'form_q_ka_employer',
      'statutory_principal_employer',
      'form25_principal_employer',
      'form_xviii_principal_employer',
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
  if (spec.match.test(labelOnly)) return true;
  const norm = formTSEHeaderNorm(labelOnly);
  if (spec.key === 'form_t_month_year') return /^month\s+year$/.test(norm);
  if (spec.key === 'form_t_establishment_name_address') {
    return /name\s+and\s+address\s+of\s+the\s+establishment/.test(norm);
  }
  if (spec.key === 'form_t_employer') {
    return /name\s+and\s+address\s+of\s+employer/.test(norm);
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

  return applyFormTSEKarnatakaAutofillFromSite(out, siteContext, { onlyEmpty: false });
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

/** Write Form T header values into column B beside labels in column A (Karnataka template rows 9–11). */
export function writeFormTSEHeaderFieldsToWorksheet(worksheet, headerFormData = {}, parsedFormHeader = null) {
  if (!worksheet || !headerFormData || typeof headerFormData !== 'object') return;
  const parsedFields = Array.isArray(parsedFormHeader?.fields) ? parsedFormHeader.fields : [];
  const getMergeTopLeft = buildFormTSEMergeTopLeftResolver(worksheet);

  const writeAt = (row, col, value) => {
    const text = String(value ?? '').trim();
    if (!text || row < 1 || col < 1) return;
    const valueTl = getMergeTopLeft(row, col);
    const labelTl = getMergeTopLeft(row, 1);
    if (valueTl.r === labelTl.r && valueTl.c === labelTl.c && col > 1) {
      const existing = readFormTSEMergedCellText(worksheet, getMergeTopLeft, row, 1);
      const labelOnly = existing.split(':')[0].trim();
      worksheet.getCell(valueTl.r, valueTl.c).value = labelOnly ? `${labelOnly} : ${text}` : text;
      return;
    }
    worksheet.getCell(valueTl.r, valueTl.c).value = text;
    if (valueTl.r !== row || valueTl.c !== col) {
      try {
        worksheet.getCell(row, col).value = text;
      } catch (_) {
        // ignore write to non-master merge cell
      }
    }
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
    if (parsedField?.labelRow != null) {
      const valueRow =
        parsedField.valueRow != null ? Number(parsedField.valueRow) + 1 : Number(parsedField.labelRow) + 1;
      const labelCol1Based =
        parsedField.labelCol != null ? Number(parsedField.labelCol) + 1 : FORM_T_KARNATAKA_VALUE_COL - 1;
      writeAt(valueRow, resolveValueCol(parsedField, labelCol1Based), val);
      return;
    }

    for (let r = 1; r <= 25; r += 1) {
      for (let c = 1; c <= 8; c += 1) {
        const tl = getMergeTopLeft(r, c);
        if (tl.r !== r || tl.c !== c) continue;
        const raw = readFormTSEMergedCellText(worksheet, getMergeTopLeft, r, c);
        if (!raw) continue;
        const rawLabel = raw.split(':')[0].trim();
        if (!formTLabelMatchesSpec(rawLabel, spec) && !fieldLabelMatchesCell(spec.label, rawLabel)) {
          continue;
        }
        writeAt(r, resolveValueCol(parsedField, c), val);
        return;
      }
    }

    const fallbackRow = FORM_T_KARNATAKA_FALLBACK_ROWS[specIndex];
    if (fallbackRow) {
      writeAt(fallbackRow, FORM_T_KARNATAKA_VALUE_COL, val);
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
    }
  }
  if (total < 4) return false;
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
  if (indexRow > mainRow) {
    const aboveIndex = indexRow - 1;
    if (aboveIndex >= mainRow) {
      let hits = 0;
      for (let c = startCol; c < maxCol; c += 1) {
        const t = normCell(readRaw(aboveIndex, c) || getMergedAwareCellText(aboveIndex, c)).toLowerCase();
        if (
          /basic|hra|payable|conveyance|medical|allowance|overtime|\bot\b|da\/vda|\bvda\b/.test(t)
        ) {
          hits += 1;
        }
      }
      if (hits >= 2) return aboveIndex;
    }
  }
  let bestRow = mainRow;
  let bestHits = 0;
  for (let r = mainRow; r <= scanEnd; r += 1) {
    if (r === indexRow) continue;
    let hits = 0;
    for (let c = startCol; c < maxCol; c += 1) {
      const t = normCell(readRaw(r, c) || getMergedAwareCellText(r, c)).toLowerCase();
      if (
        /basic|hra|payable|conveyance|medical|allowance|overtime|\bot\b|da\/vda|\bvda\b/.test(t)
      ) {
        hits += 1;
      }
    }
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

export function listFormTSEAttendanceDayHeaders(headers) {
  const out = [];
  (Array.isArray(headers) ? headers : []).forEach((header) => {
    if (!isFormTSEAttendanceDayHeader(header)) return;
    const day = parseInt(String(header).match(/^ATTENDANCE_(\d{1,2})/i)[1], 10);
    out.push({ header, day });
  });
  return out.sort((a, b) => a.day - b.day);
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
    /salary advance|salary advances|\bfines\b|damage|other deduction|total deduction/.test(h)
  ) {
    return 'Deductions';
  }
  if (
    /^basic$|da\/vda|\bhra\b|conveyance|medical allowance|attendance bonus|special allowance|^ot$/.test(
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
  const scanEnd =
    dataStartIndex > mainRow ? dataStartIndex - 1 : Math.min(mainRow + 12, jsonData.length - 1);
  const maxCol = Math.max(Number(effectiveSheetCols) || 0, tableStartCol + 55);

  let startCol = Math.max(0, Number(tableStartCol) || 0);
  for (let c = 0; c < Math.min(maxCol, 14); c += 1) {
    const t = String(getMergedAwareCellText(mainRow, c) || '').toLowerCase();
    if (/sl\.?\s*no|sr\.?\s*no|serial/.test(t)) {
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
    if (header && indexRow >= 0 && !/^ATTENDANCE_\d{1,2}$/i.test(header)) {
      const statCol = parseStatutoryColumnNumber(
        readRaw(indexRow, c) || getMergedAwareCellText(indexRow, c)
      );
      if (statCol > 0) header = appendFormTStatutoryColumnSuffix(header, statCol);
    }
    const hasSample = rowHasSampleData(jsonData, c, dataStartRow, dataStartRow + 15);
    if (!header && !hasSample && headers.length > 0) {
      let anyMore = false;
      for (let cc = c; cc < Math.min(c + 10, maxCol); cc += 1) {
        if (rowHasSampleData(jsonData, cc, dataStartRow, dataStartRow + 10)) {
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
    /\(\s*1[1-9]\s*\)/.test(String(h || ''))
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

  if (rows.length > 0 && rows[0] && typeof rows[0] === 'object' && !Array.isArray(rows[0])) {
    const rowKeys = Object.keys(rows[0]).filter((k) => !String(k).startsWith('__'));
    const keysMatchExport =
      rowKeys.length === hdrs.length && rowKeys.every((k, i) => k === hdrs[i]);
    if (!keysMatchExport && rowKeys.length >= Math.min(hdrs.length, 4)) {
      if (rowKeys.length === hdrs.length) {
        srcHdrs = rowKeys;
      } else if (srcHdrs.length !== hdrs.length) {
        srcHdrs = rowKeys;
      }
    }
  }

  const alignedRows =
    srcHdrs.length === hdrs.length && srcHdrs.some((h, i) => h !== hdrs[i])
      ? remapRowsToRebuiltTableHeaders(rows, srcHdrs, hdrs)
      : rows;

  return alignedRows.map((row, rowIndex) => {
    const out = {};
    hdrs.forEach((header, colIndex) => {
      let value = '';
      if (row && typeof row === 'object' && !Array.isArray(row)) {
        if (Object.prototype.hasOwnProperty.call(row, header)) {
          value = row[header];
        } else {
          const srcKey = srcHdrs[colIndex];
          if (srcKey && Object.prototype.hasOwnProperty.call(row, srcKey)) {
            value = row[srcKey];
          }
        }
      } else if (Array.isArray(row) && colIndex < row.length) {
        value = row[colIndex];
      }
      if (
        (value == null || String(value).trim() === '') &&
        /s\.?\s*no|serial|sl\.?\s*no/i.test(String(header || ''))
      ) {
        value = rowIndex + 1;
      }
      out[header] = value != null ? value : '';
    });
    return out;
  });
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
  const effectiveSheetCols = Math.max(
    ...(jsonData || []).map((row) => (Array.isArray(row) ? row.length : 0)),
    32
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
  const maxScanCols = Math.max(worksheet.columnCount + 5, 55);
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
  const effectiveHeaders =
    rebuilt?.headers?.length >= 4
      ? preserveFormTSEAttendanceHeadersIfNeeded(
          inputHeaders,
          normalizeFormTSETableHeaders(rebuilt.headers)
        )
      : inputHeaders;
  if (effectiveHeaders.length < 4) {
    throw new Error('Could not locate Form T table columns.');
  }

  const tableStartCol =
    rebuilt?.tableStartCol != null && rebuilt.tableStartCol >= 0
      ? rebuilt.tableStartCol
      : Math.max(0, Number(parsedTableStartCol) || 0);
  const fieldCols = effectiveHeaders.map((_, idx) => tableStartCol + idx + 1);
  const startRow =
    rebuilt?.dataStartIndex != null && rebuilt.dataStartIndex >= 0
      ? rebuilt.dataStartIndex + 1
      : Number(parsedDataStartIndex) >= 0
        ? Number(parsedDataStartIndex) + 1
        : headerRowIndex >= 0
          ? headerRowIndex + 3
          : 14;

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
    const fuzzy = rowKeys.find((k) => {
      const nk = normalize(k);
      return nk && (nk.includes(target) || target.includes(nk));
    });
    if (fuzzy) return row[fuzzy];
    if (headerIndex < rowKeys.length) return row[rowKeys[headerIndex]];
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
    inputHeaders.length === effectiveHeaders.length ? inputHeaders : effectiveHeaders;
  const sourcePrimary = prepareFormTSEExportRows(
    Array.isArray(mappedData) ? mappedData : [],
    effectiveHeaders,
    sourceHeaders
  ).filter((row) => rowLooksMeaningful(row));

  const tableColMin = fieldCols.length > 0 ? Math.min(...fieldCols) : 1;
  const tableColMax = fieldCols.length > 0 ? Math.max(...fieldCols) : effectiveHeaders.length;
  const templateBodyRows = countExcelJSTemplateBodyRows(worksheet, startRow, tableColMin, tableColMax);
  const bodyRowsToPaint = Math.max(sourcePrimary.length, templateBodyRows);

  clearExcelJSTrailingTableCells(worksheet, {
    dataStartRow: startRow,
    dataRowCount: bodyRowsToPaint,
    afterCol: tableColMax,
    throughCol: tableColMax + 30,
  });

  for (let i = 0; i < sourcePrimary.length; i += 1) {
    if (i > 0 && i % 15 === 0) {
      await yieldToMain();
    }
    const row = sourcePrimary[i];
    const excelRow = startRow + i;
    for (let j = 0; j < effectiveHeaders.length; j += 1) {
      const header = effectiveHeaders[j];
      let value = getRowValueForHeader(row, header, j, effectiveHeaders);
      if (
        (value == null || String(value).trim() === '') &&
        /s\.?\s*no|serial|sl\.?\s*no/i.test(String(header || ''))
      ) {
        value = i + 1;
      }
      const targetCol = fieldCols[j];
      if (!targetCol || targetCol < 1) continue;
      const tl = getMergeTopLeft(excelRow, targetCol);
      const cell = worksheet.getCell(tl.r, tl.c);
      if (value == null || value === '') {
        cell.value = '';
        continue;
      }
      if (
        typeof value === 'number' ||
        (typeof value === 'string' && /^-?\d+(\.\d+)?$/.test(String(value).trim()))
      ) {
        cell.value = Number(value);
      } else {
        cell.value = String(value);
      }
      cell.font = { ...(cell.font || {}), bold: false };
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
  const headerScanEnd = Math.max(25, startRow > 0 ? startRow - 1 : 25);
  writeStatutoryHeaderFieldsToExcelJsWorksheet(worksheet, {
    headerFormData: headerValues,
    parsedFormHeader,
    headerRowEnd: headerScanEnd,
    maxScanCols: 80,
    writeMode: 'both',
  });
  writeFormTSEHeaderFieldsToWorksheet(worksheet, headerValues, parsedFormHeader);

  await yieldToMain();
  const out = await workbook.xlsx.writeBuffer();
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
