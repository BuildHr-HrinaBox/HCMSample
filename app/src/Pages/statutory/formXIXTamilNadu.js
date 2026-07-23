/**
 * Tamil Nadu CLRA Form XIX — Wage Slip [See rule 78 (1) (b)].
 * Template Form_XIX_-_TamilNadu.xlsx uses a key-value layout (not MP/AP particulars 1–7).
 */

import * as XLSX from 'xlsx';
import {
  formXIXAPHeaderNorm,
  formatWorkmanNameAndGuardian,
  isFormXIXAPWageSlipContext,
} from './formXIXAPWageSlip';
import {
  flattenPayrollEarningColumns,
  getDeductionsArray,
} from '../../utils/payrollEarnings';
import {
  FORM_XIX_TN_RATE_DEFAULT,
  isFormXIXTamilNaduWageSlipContext,
  resolveFormXIXMPPayrollFields,
} from './formXIXMPWageSlip';

export { isFormXIXTamilNaduWageSlipContext, FORM_XIX_TN_RATE_DEFAULT };

/** Autofill grid columns — Excel workman band (rows 4–8) + wages computation. */
export const FORM_XIX_TN_TABLE_HEADERS = [
  'Workman Code',
  'Workman Name',
  "Father's Name",
  'UAN',
  'ESIC IP Number',
  'Date of Joining',
  'No. of. Working Days',
  'No. of. Overtime',
  'Rate of Daily Wages/Piece Rate',
  'Nature of Work',
  'Basic',
  'Dearness Allowance',
  'House Rent Allowance',
  'Leave with Wages Including Cash in Lieu of Kinds',
  'Other Allowances',
  'Gross Wages',
  'Employee Provident Fund',
  'ESIC',
  'Advance/Loan',
  'Labour Welfare Fund',
  'Professional Tax',
  'Total Wage Deductions',
  'Net Amount of Wages Paid',
];

export const FORM_XIX_TN_FIELD_GROUPS = [
  { id: 'header', title: 'Wage slip — workman particulars' },
  { id: 'wages', title: 'Wages computation' },
  { id: 'footer', title: 'Certification' },
];

/** Excel label → value cell mapping (0-based row/col for parsed layout; write uses 1-based). */
export const FORM_XIX_TN_TEMPLATE_SPECS = [
  {
    key: 'form_xix_tn_workman_code',
    label: 'Workman Code',
    group: 'header',
    match: /^workman\s+code/i,
    valueRow: 3,
    valueCol: 1,
  },
  {
    key: 'form_xix_tn_workman_name',
    label: 'Workman Name',
    group: 'header',
    match: /^workman\s+name/i,
    valueRow: 4,
    valueCol: 1,
  },
  {
    key: 'form_xix_tn_father_name',
    label: "Father's Name",
    group: 'header',
    match: /^father'?s?\s+name/i,
    valueRow: 5,
    valueCol: 1,
  },
  {
    key: 'form_xix_tn_uan',
    label: 'UAN',
    group: 'header',
    match: /^uan$/i,
    valueRow: 6,
    valueCol: 1,
  },
  {
    key: 'form_xix_tn_esic_ip',
    label: 'ESIC IP Number',
    group: 'header',
    match: /^esic\s*ip\s*number|^esi\s*c?\s*ip/i,
    valueRow: 7,
    valueCol: 1,
  },
  {
    key: 'form_xix_tn_date_of_joining',
    label: 'Date of Joining',
    group: 'header',
    match: /^date\s+of\s+joining/i,
    valueRow: 3,
    valueCol: 3,
  },
  {
    key: 'form_xix_tn_working_days',
    label: 'No. of. Working Days',
    group: 'header',
    match: /(?:no|number)\.?\s*of\.?\s*working\s+days/i,
    valueRow: 4,
    valueCol: 3,
  },
  {
    key: 'form_xix_tn_overtime',
    label: 'No. of. Overtime',
    group: 'header',
    match: /(?:no|number)\.?\s*of\.?\s*overtime/i,
    valueRow: 5,
    valueCol: 3,
  },
  {
    key: 'form_xix_tn_rate',
    label: 'Rate of Daily Wages/Piece Rate',
    group: 'header',
    match: /rate\s+of\s+daily\s+wages|piece\s*rate/i,
    valueRow: 6,
    valueCol: 3,
  },
  {
    key: 'form_xix_tn_nature_of_work',
    label: 'Nature of Work',
    group: 'header',
    match: /^nature\s+of\s+work/i,
    valueRow: 7,
    valueCol: 3,
  },
  {
    key: 'form_xix_tn_basic',
    label: 'Basic',
    group: 'wages',
    match: /^basic$/i,
    valueRow: 9,
    valueCol: 1,
  },
  {
    key: 'form_xix_tn_da',
    label: 'Dearness Allowance',
    group: 'wages',
    match: /^dearness\s+allowance/i,
    valueRow: 10,
    valueCol: 1,
  },
  {
    key: 'form_xix_tn_hra',
    label: 'House Rent Allowance',
    group: 'wages',
    match: /^house\s+rent\s+allowance|^hra$/i,
    valueRow: 11,
    valueCol: 1,
  },
  {
    key: 'form_xix_tn_leave_wages',
    label: 'Leave with Wages Including Cash in Lieu of Kinds',
    group: 'wages',
    match: /leave\s+with\s+wages/i,
    valueRow: 12,
    valueCol: 1,
  },
  {
    key: 'form_xix_tn_other_allowances',
    label: 'Other Allowances',
    group: 'wages',
    match: /^other\s+allowances?/i,
    valueRow: 13,
    valueCol: 1,
  },
  {
    key: 'form_xix_tn_gross',
    label: 'Gross Wages',
    group: 'wages',
    match: /^gross\s+wages/i,
    valueRow: 14,
    valueCol: 1,
  },
  {
    key: 'form_xix_tn_epf',
    label: 'Employee Provident Fund',
    group: 'wages',
    match: /employee\s+provident\s+fund|^epf$|^pf$/i,
    valueRow: 9,
    valueCol: 3,
  },
  {
    key: 'form_xix_tn_esic',
    label: 'ESIC',
    group: 'wages',
    match: /^esic$|^esi$/i,
    valueRow: 10,
    valueCol: 3,
  },
  {
    key: 'form_xix_tn_advance_loan',
    label: 'Advance/Loan',
    group: 'wages',
    match: /advance\s*\/?\s*loan|^advance$|^loan$/i,
    valueRow: 11,
    valueCol: 3,
  },
  {
    key: 'form_xix_tn_lwf',
    label: 'Labour Welfare Fund',
    group: 'wages',
    match: /labour\s+welfare\s+fund|^lwf$/i,
    valueRow: 12,
    valueCol: 3,
  },
  {
    key: 'form_xix_tn_pt',
    label: 'Professional Tax',
    group: 'wages',
    match: /^professional\s+tax|^pt$/i,
    valueRow: 13,
    valueCol: 3,
  },
  {
    key: 'form_xix_tn_total_deductions',
    label: 'Total Wage Deductions',
    group: 'wages',
    match: /total\s+wage\s+deductions|total\s+deductions/i,
    valueRow: 14,
    valueCol: 3,
  },
  {
    key: 'form_xix_tn_net',
    label: 'Net Amount of Wages Paid',
    group: 'wages',
    match: /net\s+amount\s+of\s+wages\s+paid|^net\s+wages|^net\s+amount/i,
    valueRow: 16,
    valueCol: 3,
  },
];

const FORM_XIX_TN_HEADER_KEY_BY_LABEL = new Map(
  FORM_XIX_TN_TEMPLATE_SPECS.map((spec) => [formXIXAPHeaderNorm(spec.label), spec.key])
);

export function formXIXTamilNaduHeaderNorm(txt) {
  return formXIXAPHeaderNorm(String(txt || '').replace(/^\d+[\.\)]\s*/, ''));
}

export function looksLikeFormXIXTamilNaduTableHeaders(headers) {
  const list = Array.isArray(headers) ? headers : [];
  if (list.length < 4) return false;
  const joined = list.map((h) => formXIXTamilNaduHeaderNorm(h)).join(' | ');
  return (
    /workman\s+code/.test(joined) &&
    /workman\s+name/.test(joined) &&
    (/father'?s?\s+name/.test(joined) || /uan/.test(joined) || /esic/.test(joined))
  );
}

export function resolveFormXIXTamilNaduTableHeaders(tableHeaders) {
  const canonical = FORM_XIX_TN_TABLE_HEADERS;
  const incoming = Array.isArray(tableHeaders)
    ? tableHeaders.map((h) => String(h || '').trim()).filter(Boolean)
    : [];
  if (!incoming.length) return [...canonical];
  const joined = incoming.map((h) => formXIXTamilNaduHeaderNorm(h)).join(' | ');
  const hasWageBand = /gross\s+wages|net\s+amount|employee\s+provident|total\s+wage\s+deductions/.test(
    joined
  );
  // Upgrade older workman-only grids to the full Excel wage-computation columns.
  if (looksLikeFormXIXTamilNaduTableHeaders(incoming) && !hasWageBand) {
    return [...canonical];
  }
  if (looksLikeFormXIXTamilNaduTableHeaders(incoming) && incoming.length >= canonical.length - 2) {
    return incoming.length >= canonical.length ? incoming.slice(0, incoming.length) : [...canonical];
  }
  return [...canonical];
}

export function isFormXIXTamilNaduHeaderFieldLayoutFormHeader(formHeader) {
  return !!formHeader?.formXIXTamilNaduHeaderFieldLayout || !!formHeader?.formXIXTamilNaduTableLayout;
}

export function isFormXIXTamilNaduTableLayoutFormHeader(formHeader) {
  return !!formHeader?.formXIXTamilNaduTableLayout;
}

export function isFormXIXTamilNaduWorkmanCodeHeader(h) {
  return /^workman\s+code$/.test(formXIXTamilNaduHeaderNorm(h));
}

export function isFormXIXTamilNaduWorkmanNameHeader(h) {
  const s = formXIXTamilNaduHeaderNorm(h);
  return /^workman\s+name$/.test(s) || (/workman/.test(s) && /name/.test(s) && !/code|father|husband/.test(s));
}

export function isFormXIXTamilNaduFatherNameHeader(h) {
  return /father'?s?\s+name/.test(formXIXTamilNaduHeaderNorm(h));
}

export function isFormXIXTamilNaduUanHeader(h) {
  return /^uan$|^uan\s*(number|no\.?)$/.test(formXIXTamilNaduHeaderNorm(h));
}

export function isFormXIXTamilNaduEsicIpHeader(h) {
  const s = formXIXTamilNaduHeaderNorm(h);
  return /esic\s*ip/.test(s) || (/esic|esi/.test(s) && /ip/.test(s) && /number|no/.test(s));
}

export function isFormXIXTamilNaduDateOfJoiningHeader(h) {
  return /date\s+of\s+joining/.test(formXIXTamilNaduHeaderNorm(h));
}

export function isFormXIXTamilNaduWorkingDaysHeader(h) {
  const s = formXIXTamilNaduHeaderNorm(h);
  return /(?:no|number)\.?\s*of\.?\s*working\s+days/.test(s) || /^working\s+days$/.test(s);
}

export function isFormXIXTamilNaduOvertimeHeader(h) {
  const s = formXIXTamilNaduHeaderNorm(h);
  return /(?:no|number)\.?\s*of\.?\s*overtime/.test(s) || /^overtime$/.test(s);
}

export function isFormXIXTamilNaduRateHeader(h) {
  const s = formXIXTamilNaduHeaderNorm(h);
  return /rate/.test(s) && (/wage|piece/.test(s) || /daily/.test(s));
}

export function isFormXIXTamilNaduNatureOfWorkHeader(h) {
  return /nature\s+of\s+work/.test(formXIXTamilNaduHeaderNorm(h));
}

export function isFormXIXTamilNaduBasicHeader(h) {
  return /^basic$/.test(formXIXTamilNaduHeaderNorm(h));
}

export function isFormXIXTamilNaduDearnessHeader(h) {
  return /dearness\s+allowance/.test(formXIXTamilNaduHeaderNorm(h));
}

export function isFormXIXTamilNaduHraHeader(h) {
  const s = formXIXTamilNaduHeaderNorm(h);
  return /house\s+rent\s+allowance/.test(s) || /^hra$/.test(s);
}

export function isFormXIXTamilNaduLeaveWagesHeader(h) {
  return /leave\s+with\s+wages/.test(formXIXTamilNaduHeaderNorm(h));
}

export function isFormXIXTamilNaduOtherAllowancesHeader(h) {
  return /^other\s+allowances?$/.test(formXIXTamilNaduHeaderNorm(h));
}

export function isFormXIXTamilNaduGrossHeader(h) {
  return /gross\s+wages/.test(formXIXTamilNaduHeaderNorm(h));
}

export function isFormXIXTamilNaduEpfHeader(h) {
  const s = formXIXTamilNaduHeaderNorm(h);
  return /employee\s+provident\s+fund/.test(s) || /^epf$/.test(s) || /^pf$/.test(s);
}

export function isFormXIXTamilNaduEsicDeductionHeader(h) {
  const s = formXIXTamilNaduHeaderNorm(h);
  if (isFormXIXTamilNaduEsicIpHeader(h)) return false;
  return /^esic$/.test(s) || /^esi$/.test(s);
}

export function isFormXIXTamilNaduAdvanceLoanHeader(h) {
  const s = formXIXTamilNaduHeaderNorm(h);
  return /advance\s*\/?\s*loan/.test(s) || /^advance$/.test(s) || /^loan$/.test(s);
}

export function isFormXIXTamilNaduLwfHeader(h) {
  const s = formXIXTamilNaduHeaderNorm(h);
  return /labour\s+welfare\s+fund/.test(s) || /^lwf$/.test(s);
}

export function isFormXIXTamilNaduProfessionalTaxHeader(h) {
  const s = formXIXTamilNaduHeaderNorm(h);
  return /professional\s+tax/.test(s) || /^pt$/.test(s);
}

export function isFormXIXTamilNaduTotalDeductionsHeader(h) {
  const s = formXIXTamilNaduHeaderNorm(h);
  return /total\s+wage\s+deductions/.test(s) || /total\s+deductions/.test(s);
}

export function isFormXIXTamilNaduNetHeader(h) {
  const s = formXIXTamilNaduHeaderNorm(h);
  return /net\s+amount\s+of\s+wages\s+paid/.test(s) || /^net\s+(amount|wages)$/.test(s);
}

/** People/payroll-owned columns — skip generic People string matching. */
export function isFormXIXTamilNaduSkipPeopleAutofillHeader(h) {
  return (
    isFormXIXTamilNaduWorkmanCodeHeader(h) ||
    isFormXIXTamilNaduWorkmanNameHeader(h) ||
    isFormXIXTamilNaduFatherNameHeader(h) ||
    isFormXIXTamilNaduUanHeader(h) ||
    isFormXIXTamilNaduEsicIpHeader(h) ||
    isFormXIXTamilNaduDateOfJoiningHeader(h) ||
    isFormXIXTamilNaduWorkingDaysHeader(h) ||
    isFormXIXTamilNaduOvertimeHeader(h) ||
    isFormXIXTamilNaduRateHeader(h) ||
    isFormXIXTamilNaduNatureOfWorkHeader(h) ||
    isFormXIXTamilNaduBasicHeader(h) ||
    isFormXIXTamilNaduDearnessHeader(h) ||
    isFormXIXTamilNaduHraHeader(h) ||
    isFormXIXTamilNaduLeaveWagesHeader(h) ||
    isFormXIXTamilNaduOtherAllowancesHeader(h) ||
    isFormXIXTamilNaduGrossHeader(h) ||
    isFormXIXTamilNaduEpfHeader(h) ||
    isFormXIXTamilNaduEsicDeductionHeader(h) ||
    isFormXIXTamilNaduAdvanceLoanHeader(h) ||
    isFormXIXTamilNaduLwfHeader(h) ||
    isFormXIXTamilNaduProfessionalTaxHeader(h) ||
    isFormXIXTamilNaduTotalDeductionsHeader(h) ||
    isFormXIXTamilNaduNetHeader(h)
  );
}

const moneyText = (value) => {
  if (value === '' || value == null) return '';
  const n = Number(String(value).replace(/,/g, '').trim());
  if (!Number.isFinite(n)) return String(value).trim();
  return String(Math.round(n * 100) / 100);
};

const findDeductionAmount = (deductions, matcher) => {
  const list = Array.isArray(deductions) ? deductions : [];
  for (let i = 0; i < list.length; i += 1) {
    const item = list[i] || {};
    const type = String(item.type || item.Type || item.component_type || '').toLowerCase();
    const name = String(
      item.name || item.Name || item.component_name || item.display_name || ''
    ).toLowerCase();
    if (!matcher(type, name)) continue;
    const amount = item.amount ?? item.Amount ?? item.value ?? item.Value;
    const text = moneyText(amount);
    if (text !== '') return text;
  }
  return '';
};

/** Map Payroll table / pay-run fields onto Excel wages-computation labels. */
export function resolveFormXIXTamilNaduWageComputationFields(payrollRow) {
  if (!payrollRow || payrollRow.fetch_error) {
    return {
      basic: '',
      dearnessAllowance: '',
      houseRentAllowance: '',
      leaveWithWages: 'NIL',
      otherAllowances: '',
      grossWages: '',
      epf: '',
      esic: '',
      advanceLoan: 'NIL',
      lwf: '',
      professionalTax: '',
      totalDeductions: '',
      netWages: '',
    };
  }
  const flat = flattenPayrollEarningColumns(payrollRow);
  const deductions = getDeductionsArray(flat);

  const basic = moneyText(flat.basic ?? flat.earned_basic);
  const dearnessAllowance = moneyText(flat.dearness_allowance);
  const houseRentAllowance = moneyText(
    flat.hra_fbp ?? flat.hra ?? flat.house_rent_allowance ?? flat['House Rent Allowance']
  );
  const grossWages = moneyText(flat.gross_pay);
  const netWages = moneyText(flat.net_pay);
  let otherAllowances = moneyText(flat.other_allowance);
  if (otherAllowances === '' && grossWages !== '') {
    const g = Number(grossWages);
    const b = Number(basic) || 0;
    const d = Number(dearnessAllowance) || 0;
    const h = Number(houseRentAllowance) || 0;
    if (Number.isFinite(g)) otherAllowances = moneyText(Math.max(0, g - b - d - h));
  }

  const leaveWithWages =
    moneyText(
      flat.leave_with_wages ||
        flat.leave_wages ||
        flat.Leave_with_Wages ||
        flat['Leave with Wages'] ||
        flat.leave_encashment
    ) || 'NIL';

  const epf =
    moneyText(flat.epf_contribution) ||
    findDeductionAmount(
      deductions,
      (t, n) =>
        t === 'epf_contribution' ||
        t === 'epf' ||
        t === 'pf' ||
        n.includes('provident fund') ||
        (n.includes('epf') && !n.includes('employer'))
    );
  const esic =
    moneyText(flat.esi || flat.esic || flat.esi_contribution) ||
    findDeductionAmount(
      deductions,
      (t, n) => t === 'esi' || t === 'esic' || n.includes('esic') || n.includes('employee state insurance')
    );
  const advanceLoan =
    moneyText(flat.advance || flat.loan || flat.advance_recovery) ||
    findDeductionAmount(
      deductions,
      (t, n) => n.includes('advance') || n.includes('loan') || t === 'advance' || t === 'loan'
    ) ||
    'NIL';
  const lwf =
    moneyText(flat.lwf || flat.labour_welfare_fund) ||
    findDeductionAmount(
      deductions,
      (t, n) => n.includes('labour welfare') || n.includes('labor welfare') || n.includes('lwf')
    );
  const professionalTax =
    moneyText(flat.professional_tax) ||
    findDeductionAmount(
      deductions,
      (t, n) => t === 'professional_tax' || t === 'pt' || n.includes('professional tax')
    );

  // Total Wage Deductions = gross_pay − net_pay (always; ignore payroll total_deductions).
  let totalDeductions = '';
  if (grossWages !== '' && netWages !== '') {
    const g = Number(grossWages);
    const n = Number(netWages);
    if (Number.isFinite(g) && Number.isFinite(n)) {
      totalDeductions = moneyText(Math.max(0, Math.round((g - n) * 100) / 100));
    }
  }

  return {
    basic,
    dearnessAllowance,
    houseRentAllowance,
    leaveWithWages,
    otherAllowances,
    grossWages,
    epf,
    esic,
    advanceLoan: advanceLoan || 'NIL',
    lwf,
    professionalTax,
    totalDeductions,
    netWages,
  };
}

const pickEmp = (emp, keys) => {
  for (const key of keys) {
    const v = emp?.[key];
    if (v != null && String(v).trim() !== '') return String(v).trim();
  }
  return '';
};

export function resolveFormXIXTamilNaduWorkmanCode(emp = {}) {
  return pickEmp(emp, [
    'EmployeeID',
    'Employee ID',
    'Employee_Number',
    'Employee Number',
    'EmployeeCode',
    'Employee Code',
    'Emp_Code',
    'Emp Code',
    'Zoho_ID',
  ]);
}

export function resolveFormXIXTamilNaduWorkmanName(emp = {}) {
  const full = String(formatWorkmanNameAndGuardian(emp) || '')
    .split(/\r?\n/)[0]
    .trim();
  if (full) return full;
  return pickEmp(emp, ['EmployeeName', 'Employee Name', 'Name', 'Full_Name', 'Full Name']);
}

export function resolveFormXIXTamilNaduFatherName(emp = {}) {
  return pickEmp(emp, [
    'Father_s_Name',
    'Father_s Name',
    'Father_Name',
    'Father Name',
    'FatherName',
    "Father's Name",
    'SpouseName',
    'Spouse Name',
    'HusbandName',
    'Husband Name',
  ]);
}

export function resolveFormXIXTamilNaduUan(emp = {}) {
  return pickEmp(emp, ['UAN_Number', 'UAN Number', 'UAN', 'uan', 'UAN No.', 'UAN_No']);
}

export function resolveFormXIXTamilNaduEsicIp(emp = {}) {
  return pickEmp(emp, [
    'ESIC_Number',
    'ESIC Number',
    'ESICNo',
    'ESIC_No',
    'ESIC IP Number',
    'ESIC IP',
    'ESI_Number',
    'ESI Number',
    'ESI No.',
    'IP_Number',
    'IP Number',
    "Employee's State Insurance Corporation No.",
  ]);
}

export function resolveFormXIXTamilNaduDateOfJoining(emp = {}, formatDate) {
  const raw = pickEmp(emp, [
    'DateOfJoining',
    'Date of Joining',
    'Dateofjoining',
    'DateofJoining',
    'Date_of_Joining',
    'Date of entry into service',
    'DateofentryintoService',
  ]);
  if (!raw) return '';
  if (typeof formatDate === 'function') {
    try {
      return String(formatDate(raw) || raw).trim();
    } catch {
      return raw;
    }
  }
  return raw;
}

export function resolveFormXIXTamilNaduNatureOfWork(emp = {}) {
  return pickEmp(emp, [
    'Designation',
    'Designation.displayValue',
    'Designation Name',
    'DesignationName',
    'JobTitle',
    'Job Title',
    'Nature_of_employment',
    'Nature of employment',
    'Nature of Work',
    'Department',
  ]);
}

export function applyFormXIXTamilNaduEmployeeToRow(row, emp, headers, helpers = {}) {
  const hdrs = resolveFormXIXTamilNaduTableHeaders(headers);
  const out = row && typeof row === 'object' ? { ...row } : {};
  const {
    sanitizeValue = (v) => String(v ?? '').trim(),
    payrollRow = null,
    formatStatutoryDateDisplay = null,
    monthCandidates = null,
  } = helpers;

  const hasPayroll = payrollRow && !payrollRow.fetch_error;
  const payroll = resolveFormXIXMPPayrollFields(hasPayroll ? payrollRow : null, {
    emp,
    monthCandidates,
    useMonthlyWageRateDefault: true,
    gujaratPayrollRules: false,
    madhyaPradeshPayrollRules: false,
    andhraPradeshPayrollRules: false,
  });
  const wages = resolveFormXIXTamilNaduWageComputationFields(hasPayroll ? payrollRow : null);

  hdrs.forEach((header) => {
    if (isFormXIXTamilNaduWorkmanCodeHeader(header)) {
      out[header] = sanitizeValue(resolveFormXIXTamilNaduWorkmanCode(emp));
      return;
    }
    if (isFormXIXTamilNaduWorkmanNameHeader(header)) {
      out[header] = sanitizeValue(resolveFormXIXTamilNaduWorkmanName(emp));
      return;
    }
    if (isFormXIXTamilNaduFatherNameHeader(header)) {
      out[header] = sanitizeValue(resolveFormXIXTamilNaduFatherName(emp));
      return;
    }
    if (isFormXIXTamilNaduUanHeader(header)) {
      out[header] = sanitizeValue(resolveFormXIXTamilNaduUan(emp));
      return;
    }
    if (isFormXIXTamilNaduEsicIpHeader(header)) {
      out[header] = sanitizeValue(resolveFormXIXTamilNaduEsicIp(emp));
      return;
    }
    if (isFormXIXTamilNaduDateOfJoiningHeader(header)) {
      out[header] = sanitizeValue(
        resolveFormXIXTamilNaduDateOfJoining(emp, formatStatutoryDateDisplay)
      );
      return;
    }
    if (isFormXIXTamilNaduWorkingDaysHeader(header)) {
      out[header] = sanitizeValue(payroll.daysWorked || '');
      return;
    }
    if (isFormXIXTamilNaduOvertimeHeader(header)) {
      const ot = payroll.overtimeWages;
      out[header] = sanitizeValue(ot !== '' && ot != null ? ot : 'NIL');
      return;
    }
    if (isFormXIXTamilNaduRateHeader(header)) {
      out[header] = sanitizeValue(FORM_XIX_TN_RATE_DEFAULT);
      return;
    }
    if (isFormXIXTamilNaduNatureOfWorkHeader(header)) {
      out[header] = sanitizeValue(resolveFormXIXTamilNaduNatureOfWork(emp));
      return;
    }
    if (isFormXIXTamilNaduBasicHeader(header)) {
      out[header] = sanitizeValue(wages.basic);
      return;
    }
    if (isFormXIXTamilNaduDearnessHeader(header)) {
      out[header] = sanitizeValue(wages.dearnessAllowance);
      return;
    }
    if (isFormXIXTamilNaduHraHeader(header)) {
      out[header] = sanitizeValue(wages.houseRentAllowance);
      return;
    }
    if (isFormXIXTamilNaduLeaveWagesHeader(header)) {
      out[header] = sanitizeValue(wages.leaveWithWages);
      return;
    }
    if (isFormXIXTamilNaduOtherAllowancesHeader(header)) {
      out[header] = sanitizeValue(wages.otherAllowances);
      return;
    }
    if (isFormXIXTamilNaduGrossHeader(header)) {
      out[header] = sanitizeValue(wages.grossWages);
      return;
    }
    if (isFormXIXTamilNaduEpfHeader(header)) {
      out[header] = sanitizeValue(wages.epf);
      return;
    }
    if (isFormXIXTamilNaduEsicDeductionHeader(header)) {
      out[header] = sanitizeValue(wages.esic);
      return;
    }
    if (isFormXIXTamilNaduAdvanceLoanHeader(header)) {
      out[header] = sanitizeValue(wages.advanceLoan);
      return;
    }
    if (isFormXIXTamilNaduLwfHeader(header)) {
      out[header] = sanitizeValue(wages.lwf);
      return;
    }
    if (isFormXIXTamilNaduProfessionalTaxHeader(header)) {
      out[header] = sanitizeValue(wages.professionalTax);
      return;
    }
    if (isFormXIXTamilNaduTotalDeductionsHeader(header)) {
      // Prefer gross − net from already-mapped row cells (or wage map).
      const grossRaw =
        wages.grossWages ||
        hdrs.reduce((acc, h) => (isFormXIXTamilNaduGrossHeader(h) ? out[h] || acc : acc), '');
      const netRaw =
        wages.netWages ||
        hdrs.reduce((acc, h) => (isFormXIXTamilNaduNetHeader(h) ? out[h] || acc : acc), '');
      const g = Number(String(grossRaw ?? '').replace(/,/g, '').trim());
      const n = Number(String(netRaw ?? '').replace(/,/g, '').trim());
      if (Number.isFinite(g) && Number.isFinite(n)) {
        out[header] = sanitizeValue(String(Math.max(0, Math.round((g - n) * 100) / 100)));
      } else {
        out[header] = sanitizeValue(wages.totalDeductions);
      }
      return;
    }
    if (isFormXIXTamilNaduNetHeader(header)) {
      out[header] = sanitizeValue(wages.netWages);
    }
  });
  return out;
}

export function enrichFormXIXTamilNaduPayrollRows(mappedData, employees, headers, helpers = {}) {
  const hdrs = resolveFormXIXTamilNaduTableHeaders(headers);
  const {
    resolvePayrollRow = null,
    sanitizeValue = (v) => String(v ?? '').trim(),
    formatStatutoryDateDisplay = null,
    monthCandidates = null,
  } = helpers;
  const list = Array.isArray(employees) ? employees : [];
  let hits = 0;
  (Array.isArray(mappedData) ? mappedData : []).forEach((row, index) => {
    if (!row || typeof row !== 'object') return;
    const empItem = list[index];
    const emp = empItem?.Employee || empItem?.employee || empItem;
    if (!emp || typeof emp !== 'object') return;
    const payrollRow = typeof resolvePayrollRow === 'function' ? resolvePayrollRow(empItem || emp) : null;
    Object.assign(
      row,
      applyFormXIXTamilNaduEmployeeToRow(row, emp, hdrs, {
        sanitizeValue,
        payrollRow: payrollRow && !payrollRow.fetch_error ? payrollRow : null,
        formatStatutoryDateDisplay,
        monthCandidates,
      })
    );
    hits += 1;
  });
  return hits;
}

export function mapFormXIXTamilNaduRowsFromEmployees(employees, headers, helpers = {}) {
  const hdrs = resolveFormXIXTamilNaduTableHeaders(headers);
  const list = Array.isArray(employees) ? employees : [];
  const { resolvePayrollRow = null, ...rest } = helpers;
  return list.map((empItem, rowIndex) => {
    const emp = empItem?.Employee || empItem?.employee || empItem;
    const payrollRow = typeof resolvePayrollRow === 'function' ? resolvePayrollRow(emp, rowIndex) : null;
    return applyFormXIXTamilNaduEmployeeToRow({}, emp, hdrs, {
      ...rest,
      payrollRow: payrollRow && !payrollRow.fetch_error ? payrollRow : null,
    });
  });
}

function buildTemplateField(spec, coords = {}) {
  return {
    key: spec.key,
    label: spec.label,
    group: spec.group || 'header',
    fieldType: spec.fieldType || 'text',
    labelRow: coords.labelRow ?? spec.labelRow ?? null,
    labelCol: coords.labelCol ?? spec.labelCol ?? null,
    valueRow: coords.valueRow ?? spec.valueRow ?? null,
    valueCol: coords.valueCol ?? spec.valueCol ?? null,
  };
}

function findLabelCell(getMergedAwareCellText, match, maxCols = 8, maxRows = 40) {
  if (typeof getMergedAwareCellText !== 'function' || !match) return null;
  for (let r = 0; r < maxRows; r += 1) {
    for (let c = 0; c < maxCols; c += 1) {
      const text = String(getMergedAwareCellText(r, c) || '').trim();
      if (!text) continue;
      const norm = formXIXTamilNaduHeaderNorm(text);
      if (match.test(norm) || match.test(text)) {
        return { labelRow: r, labelCol: c, valueRow: r, valueCol: c + 1 };
      }
    }
  }
  return null;
}

export function resolveFormXIXTamilNaduHeaderFieldLayout(parsed, workbook, hints = {}) {
  const formHeader = hints.formHeader || parsed?.formHeader || null;
  const item = hints.item || null;
  const fileName = hints.fileName || hints.formFileName || '';
  const sheetText = hints.sheetText || '';
  if (!isFormXIXTamilNaduWageSlipContext(formHeader, item, fileName, sheetText)) {
    if (!isFormXIXAPWageSlipContext(formHeader, item, fileName, sheetText)) return null;
    const parts = [fileName, formHeader?.title, item?.formFileName, sheetText]
      .filter(Boolean)
      .join(' ')
      .toLowerCase();
    if (!/tamil[\s._-]*nadu|tamilnadu|form[\s._-]*xix[\s._-]*tamil/.test(parts)) return null;
  }

  let getMergedAwareCellText = null;
  if (workbook?.SheetNames?.length) {
    const sheetName =
      hints.preferredSheetName && workbook.Sheets?.[hints.preferredSheetName]
        ? hints.preferredSheetName
        : workbook.SheetNames[0];
    const worksheet = sheetName ? workbook.Sheets[sheetName] : null;
    if (worksheet) {
      const merges = worksheet['!merges'] || [];
      const getRaw = (r, c) => {
        const ref = XLSX.utils.encode_cell({ r, c });
        const cell = worksheet[ref];
        if (!cell || cell.v == null) return '';
        return String(cell.v).trim();
      };
      getMergedAwareCellText = (r, c) => {
        const direct = getRaw(r, c);
        if (direct) return direct;
        for (let i = 0; i < merges.length; i += 1) {
          const m = merges[i];
          if (!m?.s || !m?.e) continue;
          if (r >= m.s.r && r <= m.e.r && c >= m.s.c && c <= m.e.c) {
            const topLeft = getRaw(m.s.r, m.s.c);
            if (topLeft) return topLeft;
          }
        }
        return '';
      };
    }
  }

  const fields = FORM_XIX_TN_TEMPLATE_SPECS.map((spec) => {
    const coords = getMergedAwareCellText
      ? findLabelCell(getMergedAwareCellText, spec.match) || {}
      : {};
    return buildTemplateField(spec, coords);
  });

  return {
    formHeader: {
      title: formHeader?.title || 'Form XIX',
      subtitle: formHeader?.subtitle || 'Wage Slip',
      reference: formHeader?.reference || '[See rule 78 (1) (b)]',
      formXIXAPHeaderFieldLayout: true,
      formXIXMPTableLayout: true,
      formXIXTamilNaduHeaderFieldLayout: true,
      formXIXTamilNaduTableLayout: true,
      formXIXAPTableLayout: false,
      textRows: [],
      fields,
      fieldGroups: FORM_XIX_TN_FIELD_GROUPS,
    },
    headers: resolveFormXIXTamilNaduTableHeaders(parsed?.headers || hints.tableHeaders || null),
    tableData: [],
    headerRowIndex: -1,
    dataStartIndex: 0,
    tableStartCol: 0,
    sheetName: hints.preferredSheetName || null,
  };
}

/** Map autofill row cells onto form_xix_tn_* keys for Excel export. */
export function mergeFormXIXTamilNaduTableRowIntoHeaderData(headerData, row, headers = FORM_XIX_TN_TABLE_HEADERS) {
  const out = headerData && typeof headerData === 'object' ? { ...headerData } : {};
  const hdrs = resolveFormXIXTamilNaduTableHeaders(headers);
  if (!row || typeof row !== 'object') return out;

  hdrs.forEach((header) => {
    const value = String(row[header] ?? '').trim();
    if (!value) return;
    const key = FORM_XIX_TN_HEADER_KEY_BY_LABEL.get(formXIXTamilNaduHeaderNorm(header));
    if (key) out[key] = value;
  });
  return out;
}

export function writeFormXIXTamilNaduFieldsToExcelJsWorksheet(
  worksheet,
  headerFormData,
  parsedFormHeader,
  helpers = {}
) {
  if (!worksheet) return;
  const excelCellValueToString =
    helpers.excelCellValueToString || ((v) => (v == null ? '' : String(v)));
  const fields = Array.isArray(parsedFormHeader?.fields) ? parsedFormHeader.fields : [];
  const data = headerFormData && typeof headerFormData === 'object' ? headerFormData : {};

  const writeBySpec = (spec, field) => {
    const value = data[spec.key];
    if (value == null || String(value).trim() === '') return;
    let row1 = field?.valueRow != null ? field.valueRow + 1 : null;
    let col1 = field?.valueCol != null ? field.valueCol + 1 : null;
    if (row1 == null || col1 == null) {
      // Fallback to known TN template coordinates (1-based).
      row1 = (spec.valueRow ?? 0) + 1;
      col1 = (spec.valueCol ?? 0) + 1;
    }
    const cell = worksheet.getCell(row1, col1);
    cell.value = excelCellValueToString(value);
  };

  FORM_XIX_TN_TEMPLATE_SPECS.forEach((spec) => {
    if (spec.key === 'form_xix_tn_net') return;
    const field = fields.find((f) => f.key === spec.key) || null;
    writeBySpec(spec, field);
  });

  writeFormXIXTamilNaduNetAmountRow(worksheet, data.form_xix_tn_net, excelCellValueToString);
  clearFormXIXTamilNaduExtraColumns(worksheet);
  applyFormXIXTamilNaduAllBorders(worksheet);
}

/**
 * Row 17: "Net Amount of Wages Paid" in A–C, amount in D17.
 * Unmerges full-row merges so the label is not replaced by a lone amount.
 */
export function writeFormXIXTamilNaduNetAmountRow(worksheet, netAmount, excelCellValueToString) {
  if (!worksheet) return;
  const toText =
    typeof excelCellValueToString === 'function'
      ? excelCellValueToString
      : (v) => (v == null ? '' : String(v));

  if (typeof worksheet.unMergeCells === 'function') {
    ['A17:D17', 'A17:C17', 'B17:C17', 'B17:D17', 'A17:D18', 'B17:C18'].forEach((ref) => {
      try {
        worksheet.unMergeCells(ref);
      } catch (_) {
        /* ignore */
      }
    });
  }

  for (let c = 1; c <= 4; c += 1) {
    worksheet.getCell(17, c).value = null;
  }

  worksheet.getCell(17, 1).value = 'Net Amount of Wages Paid';
  worksheet.getCell(17, 1).alignment = {
    horizontal: 'left',
    vertical: 'middle',
    wrapText: true,
  };
  try {
    if (typeof worksheet.mergeCells === 'function') {
      worksheet.mergeCells(17, 1, 17, 3);
    }
  } catch (_) {
    /* ignore */
  }

  const amountText = toText(netAmount);
  if (String(amountText || '').trim() !== '') {
    worksheet.getCell(17, 4).value = amountText;
  }
  worksheet.getCell(17, 4).alignment = {
    horizontal: 'right',
    vertical: 'middle',
  };
}

/** Clear columns E–J so AP/MP stacked leftovers do not appear beside the A–D form. */
export function clearFormXIXTamilNaduExtraColumns(worksheet, options = {}) {
  if (!worksheet) return;
  const startCol = options.startCol ?? 5; // E
  const endCol = options.endCol ?? 10; // J
  const maxRow = Math.max(worksheet.rowCount || 0, options.maxRow ?? 20);
  for (let r = 1; r <= maxRow; r += 1) {
    for (let c = startCol; c <= endCol; c += 1) {
      const cell = worksheet.getCell(r, c);
      cell.value = null;
      cell.border = undefined;
    }
  }
}

/** Excel form band A1:D17 — title through Net Amount of Wages Paid. */
export const FORM_XIX_TN_BORDER = {
  top: { style: 'thin', color: { argb: 'FF000000' } },
  left: { style: 'thin', color: { argb: 'FF000000' } },
  bottom: { style: 'thin', color: { argb: 'FF000000' } },
  right: { style: 'thin', color: { argb: 'FF000000' } },
};

export const FORM_XIX_TN_BORDER_RANGE = {
  startRow: 1,
  endRow: 17,
  startCol: 1,
  endCol: 4,
};

/** Apply thin all-borders to every cell in the TN Form XIX wage-slip grid. */
export function applyFormXIXTamilNaduAllBorders(worksheet, options = {}) {
  if (!worksheet) return;
  const startRow = options.startRow ?? FORM_XIX_TN_BORDER_RANGE.startRow;
  const endRow = options.endRow ?? FORM_XIX_TN_BORDER_RANGE.endRow;
  const startCol = options.startCol ?? FORM_XIX_TN_BORDER_RANGE.startCol;
  const endCol = options.endCol ?? FORM_XIX_TN_BORDER_RANGE.endCol;
  for (let r = startRow; r <= endRow; r += 1) {
    for (let c = startCol; c <= endCol; c += 1) {
      const cell = worksheet.getCell(r, c);
      cell.border = {
        top: { ...FORM_XIX_TN_BORDER.top },
        left: { ...FORM_XIX_TN_BORDER.left },
        bottom: { ...FORM_XIX_TN_BORDER.bottom },
        right: { ...FORM_XIX_TN_BORDER.right },
      };
    }
  }
}
