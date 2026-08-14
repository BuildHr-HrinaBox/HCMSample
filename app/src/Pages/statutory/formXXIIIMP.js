import { readPayrollNetPayForStatutory } from '../../utils/payrollEarnings';
import { isAprilPayrollMonthCandidates } from './formQKarnataka';
import { personNamesMatch } from './formFKarnataka';

/** Form XXIII MP — Register of Overtime (Madhya Pradesh). */

/** April default Normal rate of wages from Form XXIII_MP template when payroll lacks net pay. */
export const FORM_XXIII_MP_APR_DEFAULT_PAYROLL = [
  { name: 'Prem Singh Bhati', normalRate: '122603' },
  { name: 'Satyapriya Behera', normalRate: '92766' },
  { name: 'Dhavalkumar Hiteshbhai Gondaliya', normalRate: '137383' },
  { name: 'Rahul Kushwah', normalRate: '112053' },
  { name: 'Katariya Jitendra', normalRate: '118855' },
  { name: 'Shekhar Sharma', normalRate: '111112' },
  { name: 'Rajan Kumar Keshari', normalRate: '179788' },
  { name: 'Manoj Kumar Gurjar', normalRate: '57916' },
  { name: 'Nil Kamal Sahu', normalRate: '63398' },
  { name: 'Amit Kumar', normalRate: '143114' },
  { name: 'Niranjan Sahoo', normalRate: '98051' },
  { name: 'Rajanish Kumar Maurya', normalRate: '100079' },
  { name: 'Amartya Pal', normalRate: '112105' },
  { name: 'Sudhir Kumar Raula', normalRate: '64768' },
  { name: 'Prakash Hamirbhai Chudasama', normalRate: '113958' },
  { name: 'Anuj Rana', normalRate: '109046' },
  { name: 'Rakesh Behera', normalRate: '100893' },
  { name: 'Harshad Dk', normalRate: '109329' },
  { name: 'Anuj Kumar', normalRate: '100079' },
  { name: 'Vivek Pandey', normalRate: '113958' },
  { name: 'Hardik Nandaniya', normalRate: '111182' },
  { name: 'Manjunatha Yarekoppa', normalRate: '144760' },
  { name: 'Darshit Patel', normalRate: '92977' },
  { name: 'Pradeep Kumar Nayak', normalRate: '101311' },
  { name: 'Jagdish Vasan', normalRate: '14036' },
  { name: 'Pankaj Kushwah', normalRate: '98140' },
  { name: 'Mohalkhram B', normalRate: '94925' },
  { name: 'Narendra Yogi', normalRate: '96515' },
  { name: 'Susanta Rout', normalRate: '64901' },
  { name: 'Rohit Kumar', normalRate: '84650' },
  { name: 'Manikandan V', normalRate: '87217' },
  { name: 'Shrinath Shukla', normalRate: '110861' },
  { name: 'Aditya Sharma', normalRate: '90134' },
  { name: 'Harish Chandra Pandey Harish Chandra Pandey', normalRate: '101533' },
  { name: 'Rahul Pal', normalRate: '92317' },
  { name: 'Vikram Sharma', normalRate: '109405' },
  { name: 'Mannu Singh', normalRate: '103837' },
  { name: 'Sartanbhai Parmabhai Pagi', normalRate: '90454' },
  { name: 'Ashwin Vikramshi Suva', normalRate: '85085' },
  { name: 'Rahul Patil', normalRate: '99829' },
  { name: 'Jayprakash Patra', normalRate: '94343' },
  { name: 'Janak Mukeshbhai Bhaliya', normalRate: '97999' },
  { name: 'Himansu Sekhar Panda', normalRate: '93185' },
  { name: 'Karuppasamy Erulappan', normalRate: '68686' },
  { name: 'Sampat Ram Unkaram', normalRate: '64724' },
  { name: 'Pradeep Saini Durga Prasad', normalRate: '99157' },
  { name: 'Piyush Kanjariya Odhavajibhai', normalRate: '120728' },
  { name: 'Maheshbhai Meda Ratnabhai Kuberbhai', normalRate: '96849' },
  { name: 'Jaydeep Dineshbhai Dineshbhai', normalRate: '115928' },
  { name: 'Ajaykumar Mansingbhai', normalRate: '102165' },
  { name: 'Dharmendra Kumar Yadav Haripal', normalRate: '138836' },
  { name: 'Patel Bhavin Kodarbhai', normalRate: '89710' },
  { name: 'Patel Chandrakant Virabhai', normalRate: '98551' },
  { name: 'Chandrakant Parmar', normalRate: '114766' },
  { name: 'Sujit Kumar Shivanna', normalRate: '121170' },
  { name: 'Ravi Kumar Singh', normalRate: '106452' },
  { name: 'Rajeshkumar Jamabhai Nai', normalRate: '117217' },
  { name: 'Rajan Vijay Arote', normalRate: '147131' },
  { name: 'Sachin Sharma', normalRate: '158629' },
  { name: 'Pradeep Sampath', normalRate: '87314' },
  { name: 'Muthukumar Paramasivan', normalRate: '209018' },
  { name: 'Kolappan Madevan Pillai', normalRate: '266963' },
];

export function formatFormXXIIIMPEmployeeName(emp = {}) {
  const fn = String(emp.FirstName || emp['FirstName'] || emp.firstName || emp['First Name'] || '').trim();
  const ln = String(emp.LastName || emp['LastName'] || emp.lastName || emp['Last Name'] || '').trim();
  if (fn && ln) return `${fn} ${ln}`;
  return fn || ln || String(emp.Name || emp['Name'] || emp.EmployeeName || emp['Employee Name'] || '').trim();
}

export function resolveFormXXIIIMPAprilDefaultNormalRate(emp) {
  const empName = formatFormXXIIIMPEmployeeName(emp);
  if (!empName) return '';
  const match = FORM_XXIII_MP_APR_DEFAULT_PAYROLL.find((entry) => personNamesMatch(empName, entry.name));
  return match?.normalRate != null ? String(match.normalRate).trim() : '';
}

/** People record salary when Sample Payroll has no row for the employee. */
export function resolveFormXXIIIMPPeopleSalary(emp = {}) {
  const raw =
    emp?.monthly_salary ??
    emp?.['monthly_salary'] ??
    emp?.MonthlySalary ??
    emp?.['Monthly Salary'] ??
    emp?.BasicSalary ??
    emp?.['Basic Salary'] ??
    emp?.Basic ??
    emp?.['Basic'] ??
    emp?.CTC ??
    emp?.['CTC'] ??
    '';
  const s = String(raw ?? '').replace(/[,₹]/g, '').trim();
  if (!s) return '';
  const n = Number(s);
  if (Number.isFinite(n) && n > 0) return String(n);
  return s;
}

export function isFormXXIIIMPContext(formHeader, rowItem, fileName, sheetText = '') {
  const parts = [
    rowItem?.formName,
    rowItem?.FormName,
    rowItem?.state,
    rowItem?.State,
    rowItem?.siteState,
    rowItem?.SiteState,
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

  const hasMP =
    /madhya[\s._-]*pradesh/.test(parts) ||
    /\bform[\s._-]*xxiii[\s._-]*mp\b/.test(parts) ||
    /\bxxiii[\s._-]*mp\b/.test(parts);
  const hasXXIII =
    /\bform[\s._-]*xxiii/i.test(parts) || /register\s+of\s+overtime/.test(parts);
  return hasMP && hasXXIII;
}

/**
 * Normal rate of wages for Form XXIII MP:
 * 1) Sample Payroll net_pay (when present)
 * 2) People MonthlySalary / Basic / CTC (when payroll row is missing)
 * 3) April template defaults only as last resort
 *
 * Never prefer template dummy rates over live payroll or People salary.
 */
export function resolveFormXXIIIMPNormalRateForEmployee(emp, payrollRow = null, monthCandidates = null) {
  if (payrollRow && !payrollRow.fetch_error) {
    const fromPayroll = String(readPayrollNetPayForStatutory(payrollRow) ?? '').trim();
    if (fromPayroll) return fromPayroll;
  }
  const fromPeople = resolveFormXXIIIMPPeopleSalary(emp);
  if (fromPeople) return fromPeople;
  if (isAprilPayrollMonthCandidates(monthCandidates)) {
    const aprilDefault = resolveFormXXIIIMPAprilDefaultNormalRate(emp);
    if (aprilDefault) return aprilDefault;
  }
  return '';
}

/** Form XXIII MP overtime columns that always default to NIL (no attendance/payroll OT). */
export const FORM_XXIII_MP_OT_NIL = 'NIL';

const normFormXXIIIMPHeader = (h) =>
  String(h || '')
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();

export function isFormXXIIIMPNormalRateHeader(h) {
  const s = normFormXXIIIMPHeader(h);
  if (!s || /overtime\s+rate/.test(s)) return false;
  return s.includes('normal') && s.includes('rate') && (s.includes('wage') || s.includes('pay'));
}

const isFormXXIIIMPEmptyRateCell = (value) => {
  const s = String(value ?? '').trim();
  if (!s) return true;
  return /^enter\b/i.test(s) || s.toLowerCase().includes('enter ');
};

/**
 * Copy a Normal rate stored under an Excel alias key (newlines / extra spaces)
 * onto the visible header the autofill grid reads.
 */
export function copyFormXXIIIMPNormalRateAliasesToRow(row, headers) {
  const hdrs = Array.isArray(headers) ? headers : [];
  const out = row && typeof row === 'object' ? { ...row } : {};
  const canonical = hdrs.filter(isFormXXIIIMPNormalRateHeader);
  if (!canonical.length) return out;
  let found = '';
  Object.entries(out).forEach(([key, val]) => {
    if (found) return;
    const s = String(val ?? '').trim();
    if (!s || isFormXXIIIMPEmptyRateCell(s)) return;
    if (isFormXXIIIMPNormalRateHeader(key)) found = s;
  });
  if (!found) return out;
  canonical.forEach((header) => {
    if (isFormXXIIIMPEmptyRateCell(out[header])) out[header] = found;
  });
  return out;
}

/** Write Normal rate onto every matching table header (including template newline variants). */
export function applyFormXXIIIMPNormalRateToRow(
  row,
  headers,
  emp,
  payrollRow = null,
  monthCandidates = null,
  helpers = {}
) {
  const hdrs = Array.isArray(headers) ? headers : [];
  const out = copyFormXXIIIMPNormalRateAliasesToRow(row, hdrs);
  const { overwrite = true, sanitizeValue = (v) => v } = helpers;
  const rate = resolveFormXXIIIMPNormalRateForEmployee(emp, payrollRow, monthCandidates);
  if (!rate) return out;
  hdrs.forEach((header) => {
    if (!isFormXXIIIMPNormalRateHeader(header)) return;
    const existing = String(out[header] ?? '').trim();
    if (!overwrite && existing && !isFormXXIIIMPEmptyRateCell(existing)) return;
    out[header] = sanitizeValue(rate);
  });
  return out;
}

export function applyFormXXIIIMPNormalRateToMappedRows(
  mappedData,
  headers,
  employeesForMapping = [],
  payrollRowsByIndex = [],
  monthCandidates = null,
  helpers = {}
) {
  if (!Array.isArray(mappedData)) return [];
  return mappedData.map((row, index) => {
    const empItem = employeesForMapping[index];
    const emp = empItem && (empItem.Employee || empItem.employee || empItem);
    const payrollRow = Array.isArray(payrollRowsByIndex) ? payrollRowsByIndex[index] : null;
    return applyFormXXIIIMPNormalRateToRow(
      row,
      headers,
      emp,
      payrollRow,
      monthCandidates,
      helpers
    );
  });
}

/** Dates on which overtime worked */
export function isFormXXIIIMPOtWorkedDatesHeader(h) {
  const s = normFormXXIIIMPHeader(h);
  if (!s) return false;
  if (/payment|paid|rate|wage|earning/.test(s) && !/worked/.test(s)) return false;
  return (
    /dates?\s+on\s+which\s+over[\s-]*time\s+worked/.test(s) ||
    (s.includes('date') && s.includes('overtime') && s.includes('worked') && !/paid|payment/.test(s))
  );
}

/** Total overtime worked or production in case of piece rate */
export function isFormXXIIIMPTotalOvertimeWorkedHeader(h) {
  const s = normFormXXIIIMPHeader(h);
  if (!s) return false;
  return (
    (s.includes('total') && s.includes('overtime') && (s.includes('worked') || s.includes('production'))) ||
    (s.includes('overtime') && s.includes('piece') && s.includes('rate'))
  );
}

/** Overtime rate of wages */
export function isFormXXIIIMPOvertimeRateHeader(h) {
  const s = normFormXXIIIMPHeader(h);
  if (!s || /normal/.test(s)) return false;
  return s.includes('overtime') && s.includes('rate') && (s.includes('wage') || s.includes('pay'));
}

/** Overtime earnings */
export function isFormXXIIIMPOvertimeEarningsHeader(h) {
  const s = normFormXXIIIMPHeader(h);
  return s.includes('overtime') && s.includes('earning');
}

/** Date on which overtime wages paid */
export function isFormXXIIIMPOtWagesPaidDateHeader(h) {
  const s = normFormXXIIIMPHeader(h);
  return (
    (s.includes('overtime') && s.includes('paid')) ||
    /dates?\s+on\s+which\s+overtime\s+wage/.test(s)
  );
}

export function isFormXXIIIMPOtNilHeader(header) {
  return (
    isFormXXIIIMPOtWorkedDatesHeader(header) ||
    isFormXXIIIMPTotalOvertimeWorkedHeader(header) ||
    isFormXXIIIMPOvertimeRateHeader(header) ||
    isFormXXIIIMPOvertimeEarningsHeader(header) ||
    isFormXXIIIMPOtWagesPaidDateHeader(header)
  );
}

export function applyFormXXIIIMPOtNilToRow(row, headers, helpers = {}) {
  const hdrs = Array.isArray(headers) ? headers : [];
  const out = row && typeof row === 'object' ? { ...row } : {};
  const nilText = helpers.nilText != null ? String(helpers.nilText) : FORM_XXIII_MP_OT_NIL;
  const { overwrite = true } = helpers;

  hdrs.forEach((header) => {
    if (!isFormXXIIIMPOtNilHeader(header)) return;
    const existing = String(out[header] ?? '').trim();
    if (
      !overwrite &&
      existing &&
      !/^enter\b/i.test(existing) &&
      !/^nil+$/i.test(existing) &&
      existing.toLowerCase() !== 'n/a' &&
      existing !== '-' &&
      existing !== '—'
    ) {
      return;
    }
    out[header] = nilText;
  });
  return out;
}

export function applyFormXXIIIMPOtNilToMappedRows(
  mappedData,
  headers,
  nilText = FORM_XXIII_MP_OT_NIL,
  helpers = {}
) {
  if (!Array.isArray(mappedData)) return [];
  const { overwrite = true } = helpers;
  return mappedData.map((row) =>
    applyFormXXIIIMPOtNilToRow(row, headers, { nilText, overwrite })
  );
}
