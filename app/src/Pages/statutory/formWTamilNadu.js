import { flattenPayrollEarningColumns, readPayrollScalar } from '../../utils/payrollEarnings';
import { samplePayrollRowMatchesEmployeeId } from '../../utils/samplePayrollApi';
import {
  collectForm10RowNameParts,
  findForm10PayrollRowByFirstAndLastName,
} from './form10TamilNadu';
import { personNamesMatch } from './formFKarnataka';

/**
 * Filename Form_W_-_TamilNadu.xlsx: word-boundary after "w" fails because "_" is a word char.
 * Match Form W / Form_W / Form-W with non-alnum boundary after W.
 */
export function looksLikeFormWFilename(text) {
  const s = String(text || '').toLowerCase();
  return /(?:^|[^a-z0-9])form[\s._-]*w(?:[^a-z0-9]|$)/i.test(s) || /\bform\s+w\b/i.test(s);
}

export function looksLikeFormWTamilNaduText(text) {
  const s = String(text || '').toLowerCase();
  return (
    /tamil[\s._-]*nadu|tamilnadu/.test(s) ||
    /form[\s._-]*w[\s._-]*tamil|form_w[_\s-]*tamil/.test(s)
  );
}

export function buildFormWTamilNaduContextBlob(hints = {}) {
  const item = hints.item || hints;
  return [
    hints.fileName,
    hints.formFileName,
    hints.formHeaderTitle,
    hints.formHeaderSubtitle,
    hints.sheetText,
    item?.formName,
    item?.FormName,
    item?.state,
    item?.State,
    item?.description,
    item?.Description,
    item?.formFileName,
    item?.FormFileName,
    Array.isArray(hints.tableHeaders) ? hints.tableHeaders.join(' ') : '',
  ]
    .filter(Boolean)
    .join(' ')
    .toLowerCase();
}

/** True for Form_W_-_TamilNadu.xlsx Register of Wages autofill (must beat Form 15 Part 2 header heuristic). */
export function isFormWTamilNaduRegisterContext(hints = {}) {
  const blob = buildFormWTamilNaduContextBlob(hints);
  if (!blob) return false;
  const fileBlob = [hints.fileName, hints.formFileName, hints.item?.formName, hints.item?.FormName]
    .map((v) => String(v || '').toLowerCase())
    .join(' ');
  const isFormW =
    looksLikeFormWFilename(fileBlob) ||
    looksLikeFormWFilename(blob) ||
    /register\s+of\s+wages/.test(blob);
  if (!isFormW) return false;
  return looksLikeFormWTamilNaduText(blob) || looksLikeFormWTamilNaduText(fileBlob);
}

/** Form W Tamil Nadu — Register of Wages default Basic / HRA when payroll lacks values. */
export const FORM_W_TN_DEFAULT_PAYROLL = [
  {
    ids: ['VE0147'],
    names: ['Avudaiappan', 'Avudaiappa'],
    basic: '67729',
    hra: '32665',
  },
  {
    ids: ['VE0079', 'VF0079'],
    names: ['Vijayakumar', 'Vijayakumal'],
    basic: '52628',
    hra: '22476',
  },
  {
    ids: ['VE0054'],
    names: ['Ajikumar'],
    basic: '67064',
    hra: '32345',
  },
  {
    ids: ['VE0051'],
    names: ['Senthilkannan', 'Senthilkann'],
    basic: '48145',
    hra: '22647',
  },
  {
    ids: ['VE0042'],
    names: ['Vinu Monikandan', 'Vinu Monik', 'Vinu Monikandan Muruganatham'],
    basic: '43688',
    hra: '20704',
    // Number of days worked defaults when Sample Payroll Paid_days is missing.
    paidDaysByMonth: {
      '04': 30, // April
      '05': 31, // May
      '06': 30, // June
      '07': 31, // July
      '08': 31, // August
    },
  },
];

/** Month → default Number of days worked for listed Form W TN employees (e.g. VE0042). */
export function resolveFormWTamilNaduDefaultPaidDays(empOrRow, monthIso = '', extraParts = null) {
  const month = String(monthIso || '').trim();
  const mm = /^\d{4}-(\d{2})$/.test(month)
    ? month.slice(5, 7)
    : /^\d{2}$/.test(month)
      ? month
      : '';
  if (!mm) return '';

  const emp = empOrRow && typeof empOrRow === 'object' ? empOrRow : {};
  const empId = String(
    resolveFormWTamilNaduEmployeeId(emp) ||
      (extraParts && extraParts.employeeId) ||
      emp.EmployeeID ||
      emp['Employee Identification No.'] ||
      ''
  )
    .trim()
    .toUpperCase();
  const empName =
    formatFormWTamilNaduEmployeeName(emp) ||
    String((extraParts && extraParts.fullName) || emp['Name of the Employee'] || '').trim();

  const entry = FORM_W_TN_DEFAULT_PAYROLL.find((item) => {
    if (!item || !item.paidDaysByMonth) return false;
    const idHit =
      empId && (item.ids || []).some((id) => String(id).trim().toUpperCase() === empId);
    const nameHit =
      empName &&
      (item.names || []).some((name) => personNamesMatch(empName, name));
    return Boolean(idHit || nameHit);
  });
  if (!entry) return '';
  const days = entry.paidDaysByMonth[mm];
  if (days == null || days === '') return '';
  const num = Number(days);
  return Number.isFinite(num) && num >= 0 ? num : '';
}

function pickFormWTamilNaduNameField(src, keys) {
  if (!src || typeof src !== 'object') return '';
  for (let i = 0; i < keys.length; i += 1) {
    const raw = src[keys[i]];
    if (raw == null) continue;
    const text = String(raw).replace(/\r?\n/g, ' ').replace(/\s+/g, ' ').trim();
    if (text) return text;
  }
  return '';
}

const FORM_W_TN_FIRST_KEYS = [
  'FirstName',
  'First Name',
  'firstName',
  'first_name',
  'First_Name',
];
const FORM_W_TN_MIDDLE_KEYS = [
  'MiddleName',
  'Middle Name',
  'middleName',
  'middle_name',
  'Middle_Name',
];
const FORM_W_TN_LAST_KEYS = [
  'LastName',
  'Last Name',
  'lastName',
  'last_name',
  'Last_Name',
  'Surname',
  'surname',
];

/**
 * People / payroll name parts: FirstName + MiddleName + LastName.
 * Also splits multi-token FirstName ("Vinu Monikandan") into given + middle when
 * MiddleName is blank and LastName is present.
 */
export function readFormWTamilNaduPersonNameParts(personOrRow) {
  if (!personOrRow || typeof personOrRow !== 'object') {
    return { firstName: '', middleName: '', lastName: '', fullName: '' };
  }
  const payload =
    personOrRow.payroll_payload && typeof personOrRow.payroll_payload === 'object'
      ? personOrRow.payroll_payload
      : {};
  const employee =
    personOrRow.employee &&
    typeof personOrRow.employee === 'object' &&
    !Array.isArray(personOrRow.employee)
      ? personOrRow.employee
      : {};
  const pools = [personOrRow, payload, employee];
  let firstName = '';
  let middleName = '';
  let lastName = '';
  pools.forEach((pool) => {
    if (!firstName) firstName = pickFormWTamilNaduNameField(pool, FORM_W_TN_FIRST_KEYS);
    if (!middleName) middleName = pickFormWTamilNaduNameField(pool, FORM_W_TN_MIDDLE_KEYS);
    if (!lastName) lastName = pickFormWTamilNaduNameField(pool, FORM_W_TN_LAST_KEYS);
  });

  // FirstName holds given + middle (e.g. "Vinu Monikandan") with separate LastName.
  if (firstName && lastName && !middleName && /\s/.test(firstName)) {
    const rawTokens = String(firstName)
      .replace(/\r?\n/g, ' ')
      .replace(/\s+/g, ' ')
      .trim()
      .split(' ')
      .filter(Boolean);
    if (rawTokens.length >= 2) {
      firstName = rawTokens[0];
      middleName = rawTokens.slice(1).join(' ');
    }
  }

  const fullName = [firstName, middleName, lastName].filter(Boolean).join(' ').trim();
  return { firstName, middleName, lastName, fullName };
}

function formatFormWTamilNaduEmployeeName(emp = {}) {
  const parts = readFormWTamilNaduPersonNameParts(emp);
  if (parts.fullName) return parts.fullName;
  return String(emp.Name || emp['Name'] || emp.EmployeeName || emp['Employee Name'] || '').trim();
}

function resolveFormWTamilNaduEmployeeId(emp = {}) {
  return String(
    emp.EmployeeID ||
      emp['EmployeeID'] ||
      emp.Employee_ID ||
      emp['Employee ID'] ||
      emp.employeeId ||
      emp.EmployeeCode ||
      emp['EmployeeCode'] ||
      ''
  )
    .trim()
    .toUpperCase();
}

function parseFormWMoney(value) {
  if (value === '' || value == null) return NaN;
  const n = Number(String(value).replace(/,/g, '').trim());
  return Number.isFinite(n) ? n : NaN;
}

/** Other Allowances = gross_pay − basic − hra */
export function computeFormWTamilNaduOtherAllowances(grossPay, basic, hra) {
  const g = parseFormWMoney(grossPay);
  if (!Number.isFinite(g)) return '';
  const b = parseFormWMoney(basic);
  const h = parseFormWMoney(hra);
  const known = (Number.isFinite(b) ? b : 0) + (Number.isFinite(h) ? h : 0);
  const other = Math.round((g - known) * 100) / 100;
  return Number.isFinite(other) ? other : '';
}

/** Total Deductions = gross_pay − net_pay */
export function computeFormWTamilNaduTotalDeductions(grossPay, netPay) {
  const g = parseFormWMoney(grossPay);
  const n = parseFormWMoney(netPay);
  if (!Number.isFinite(g) || !Number.isFinite(n) || g < n) return '';
  return Math.round((g - n) * 100) / 100;
}

export function resolveFormWTamilNaduDefaultPayroll(emp) {
  if (!emp || typeof emp !== 'object') return null;
  const empId = resolveFormWTamilNaduEmployeeId(emp);
  if (empId) {
    const byId = FORM_W_TN_DEFAULT_PAYROLL.find((entry) =>
      (entry.ids || []).some((id) => String(id).trim().toUpperCase() === empId)
    );
    if (byId) {
      return {
        basic: byId.basic != null ? String(byId.basic) : '',
        hra: byId.hra != null ? String(byId.hra) : '',
      };
    }
  }
  const empName = formatFormWTamilNaduEmployeeName(emp);
  if (!empName) return null;
  const byName = FORM_W_TN_DEFAULT_PAYROLL.find((entry) =>
    (entry.names || []).some((name) => personNamesMatch(empName, name))
  );
  if (!byName) return null;
  return {
    basic: byName.basic != null ? String(byName.basic) : '',
    hra: byName.hra != null ? String(byName.hra) : '',
  };
}

/**
 * Apply Form W TN payroll map rules:
 * - listed employees → force Basic / HRA defaults
 * - Other Allowances → gross − basic − hra
 * - Total Deductions → gross − net
 */
export function applyFormWTamilNaduDefaultPayrollToMap(map, emp) {
  const out = { ...(map || {}) };
  const defaults = resolveFormWTamilNaduDefaultPayroll(emp);
  if (defaults) {
    // Always put the configured Basic / HRA for these employees.
    if (defaults.basic) out.basicWage = defaults.basic;
    if (defaults.hra) out.houseRentAllowance = defaults.hra;
  }

  const gross = out.grossPay ?? out.grossWages;
  const net = out.netPay ?? out.netWages;
  const other = computeFormWTamilNaduOtherAllowances(gross, out.basicWage, out.houseRentAllowance);
  if (other !== '') out.otherAllowances = other;

  const ded = computeFormWTamilNaduTotalDeductions(gross, net);
  if (ded !== '') {
    out.deductionsFromGrossNet = ded;
    out.totalDeductions = ded;
  }
  return out;
}

export function hasFormWTamilNaduDefaultPayrollContext(emp) {
  return !!resolveFormWTamilNaduDefaultPayroll(emp);
}

/** Header keys for the four gender classification boxes + total headcount. */
export const FORM_W_TN_GENDER_HEADER_KEYS = {
  totalPersons: 'form_w_tn_total_persons_employed',
  men: 'form_w_tn_men',
  women: 'form_w_tn_women',
  maleYoung: 'form_w_tn_male_young_person',
  femaleYoung: 'form_w_tn_female_young_person',
};

export const FORM_W_TN_GENDER_BOX_LABELS = [
  { key: FORM_W_TN_GENDER_HEADER_KEYS.men, label: 'Men' },
  { key: FORM_W_TN_GENDER_HEADER_KEYS.women, label: 'Women' },
  { key: FORM_W_TN_GENDER_HEADER_KEYS.maleYoung, label: 'Male young person' },
  { key: FORM_W_TN_GENDER_HEADER_KEYS.femaleYoung, label: 'Female young person' },
];

export function resolveFormWTamilNaduEmployeeGender(emp) {
  const raw = String(
    emp?.Sex ||
      emp?.Gender ||
      emp?.gender ||
      emp?.['Sex'] ||
      emp?.['Gender'] ||
      emp?.['Gender.displayValue'] ||
      ''
  )
    .trim()
    .toLowerCase();
  if (!raw) return '';
  if (raw === 'm' || raw === 'male' || raw.startsWith('male')) return 'male';
  if (raw === 'f' || raw === 'female' || raw.startsWith('female')) return 'female';
  return '';
}

function parseFormWTamilNaduDob(emp) {
  const raw = String(
    emp?.Date_of_birth ||
      emp?.['Date_of_birth'] ||
      emp?.Dateofbirth ||
      emp?.DOB ||
      emp?.dob ||
      emp?.['Date of Birth'] ||
      ''
  ).trim();
  if (!raw) return null;
  const iso = raw.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) return new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3]));
  const dmy = raw.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})$/);
  if (dmy) return new Date(Number(dmy[3]), Number(dmy[2]) - 1, Number(dmy[1]));
  const t = Date.parse(raw);
  if (Number.isFinite(t)) return new Date(t);
  return null;
}

/** Young person = age under 18 as of end of payroll month (or today). */
export function isFormWTamilNaduYoungPerson(emp, payrollMonthIso = '') {
  const dob = parseFormWTamilNaduDob(emp);
  if (!dob || Number.isNaN(dob.getTime())) return false;
  let asOf = new Date();
  const m = String(payrollMonthIso || '').trim();
  if (/^\d{4}-\d{2}$/.test(m)) {
    const year = Number(m.slice(0, 4));
    const month = Number(m.slice(5, 7));
    asOf = new Date(year, month, 0);
  }
  let age = asOf.getFullYear() - dob.getFullYear();
  const monthDiff = asOf.getMonth() - dob.getMonth();
  if (monthDiff < 0 || (monthDiff === 0 && asOf.getDate() < dob.getDate())) age -= 1;
  return age >= 0 && age < 18;
}

/**
 * Count Men / Women / Male young person / Female young person from Zoho employees.
 * Young persons are excluded from Men/Women adult counts.
 */
export function countFormWTamilNaduGenderBoxes(employees, payrollMonthIso = '') {
  let men = 0;
  let women = 0;
  let maleYoung = 0;
  let femaleYoung = 0;
  const list = Array.isArray(employees) ? employees : [];
  list.forEach((item) => {
    const emp = item?.Employee || item?.employee || item;
    if (!emp || typeof emp !== 'object') return;
    const gender = resolveFormWTamilNaduEmployeeGender(emp);
    const young = isFormWTamilNaduYoungPerson(emp, payrollMonthIso);
    if (young) {
      if (gender === 'female') femaleYoung += 1;
      else if (gender === 'male') maleYoung += 1;
      return;
    }
    if (gender === 'female') women += 1;
    else if (gender === 'male') men += 1;
  });
  return {
    men,
    women,
    maleYoung,
    femaleYoung,
    total: men + women + maleYoung + femaleYoung,
  };
}

/** Month-end payment date as DD-MM-YYYY. */
export function formatFormWTamilNaduMonthEndPaymentDate(payrollMonthIso) {
  const m = String(payrollMonthIso || '').trim();
  if (!/^\d{4}-\d{2}$/.test(m)) return '';
  const year = parseInt(m.slice(0, 4), 10);
  const month = parseInt(m.slice(5, 7), 10);
  if (!Number.isFinite(year) || !Number.isFinite(month) || month < 1 || month > 12) return '';
  const lastDay = new Date(year, month, 0).getDate();
  return `${String(lastDay).padStart(2, '0')}-${String(month).padStart(2, '0')}-${year}`;
}

export function applyFormWTamilNaduGenderCountsToHeaderFormData(headerFormData, counts = {}) {
  const out = { ...(headerFormData || {}) };
  const total =
    counts.total != null
      ? counts.total
      : (Number(counts.men) || 0) +
        (Number(counts.women) || 0) +
        (Number(counts.maleYoung) || 0) +
        (Number(counts.femaleYoung) || 0);
  out[FORM_W_TN_GENDER_HEADER_KEYS.totalPersons] = String(total);
  out[FORM_W_TN_GENDER_HEADER_KEYS.men] = String(counts.men ?? 0);
  out[FORM_W_TN_GENDER_HEADER_KEYS.women] = String(counts.women ?? 0);
  out[FORM_W_TN_GENDER_HEADER_KEYS.maleYoung] = String(counts.maleYoung ?? 0);
  out[FORM_W_TN_GENDER_HEADER_KEYS.femaleYoung] = String(counts.femaleYoung ?? 0);
  return out;
}

/** Ensure Men / Women labels + young-person fields exist in the modal header field list. */
export function ensureFormWTamilNaduGenderHeaderFields(fields) {
  const list = Array.isArray(fields) ? [...fields] : [];
  const hasKey = (key) => list.some((f) => String(f?.key || '') === key);
  const hasLabel = (label) =>
    list.some((f) => String(f?.label || '').replace(/:$/, '').trim().toLowerCase() === label.toLowerCase());

  if (!hasKey(FORM_W_TN_GENDER_HEADER_KEYS.totalPersons) && !hasLabel('Total number of persons employed')) {
    list.push({
      label: 'Total number of persons employed:',
      value: '',
      key: FORM_W_TN_GENDER_HEADER_KEYS.totalPersons,
    });
  }
  FORM_W_TN_GENDER_BOX_LABELS.forEach(({ key, label }) => {
    if (hasKey(key) || hasLabel(label)) return;
    list.push({ label: `${label}:`, value: '', key });
  });
  return list;
}

function excelCellToPlainText(val) {
  if (val == null) return '';
  if (typeof val === 'string' || typeof val === 'number' || typeof val === 'boolean') return String(val);
  if (typeof val === 'object') {
    if (Array.isArray(val.richText)) return val.richText.map((rt) => rt?.text || '').join('');
    if (val.text != null) return String(val.text);
    if (val.result != null) return String(val.result);
  }
  return '';
}

/**
 * Prefer an explicit header count (including 0). Do not treat 0 as missing via `||`.
 */
export function pickFormWTamilNaduGenderCount(headerVal, fallbackVal) {
  if (headerVal != null && String(headerVal).trim() !== '') {
    const n = Number(String(headerVal).replace(/,/g, '').trim());
    if (Number.isFinite(n) && n >= 0) return n;
  }
  const f = Number(fallbackVal);
  return Number.isFinite(f) && f >= 0 ? f : 0;
}

/**
 * Write Men / Women / young-person labels (row above counts) and gender-wise counts into Form W Excel.
 * Template layout: labels on one row (Men, Women, Male young person, Female young person), counts on the next.
 * Anchor under "Total number of persons employed" / "Male young person" — never write into Establishment cols A–F.
 */
export function writeFormWTamilNaduGenderBoxesToExcelJsWorksheet(worksheet, counts = {}) {
  if (!worksheet) return false;
  const maxRows = Math.min(20, worksheet.rowCount || 20);
  const maxCols = Math.min(40, worksheet.columnCount || 40);
  const minGenderCol = 7; // keep out of Establishment / Employer columns (A–F)
  let labelRow = -1;
  let startCol = -1;

  // 1) Prefer exact "Men" label already in the gender band.
  for (let r = 1; r <= maxRows; r += 1) {
    for (let c = minGenderCol; c <= maxCols; c += 1) {
      const t = excelCellToPlainText(worksheet.getCell(r, c)?.value)
        .replace(/\s+/g, ' ')
        .trim()
        .toLowerCase();
      if (t === 'men') {
        labelRow = r;
        startCol = c;
        break;
      }
    }
    if (labelRow > 0) break;
  }

  // 2) Anchor on "Male young person" → Men is two columns left.
  if (labelRow < 0) {
    for (let r = 1; r <= maxRows; r += 1) {
      for (let c = minGenderCol; c <= maxCols; c += 1) {
        const t = excelCellToPlainText(worksheet.getCell(r, c)?.value)
          .replace(/\s+/g, ' ')
          .trim()
          .toLowerCase();
        if (t === 'male young person' || t.includes('male young person')) {
          labelRow = r;
          startCol = Math.max(minGenderCol, c - 2);
          break;
        }
      }
      if (labelRow > 0) break;
    }
  }

  // 3) Fallback: directly under "Total number of persons employed"
  if (labelRow < 0) {
    for (let r = 1; r <= maxRows; r += 1) {
      for (let c = minGenderCol; c <= maxCols; c += 1) {
        const t = excelCellToPlainText(worksheet.getCell(r, c)?.value)
          .replace(/\s+/g, ' ')
          .trim()
          .toLowerCase();
        if (t.includes('total number of persons') && t.includes('employed')) {
          labelRow = r + 1;
          startCol = c;
          break;
        }
      }
      if (labelRow > 0) break;
    }
  }

  if (labelRow < 1 || startCol < minGenderCol) return false;

  const labels = ['Men', 'Women', 'Male young person', 'Female young person'];
  const values = [
    pickFormWTamilNaduGenderCount(counts.men, 0),
    pickFormWTamilNaduGenderCount(counts.women, 0),
    pickFormWTamilNaduGenderCount(counts.maleYoung, 0),
    pickFormWTamilNaduGenderCount(counts.femaleYoung, 0),
  ];
  for (let i = 0; i < 4; i += 1) {
    const col = startCol + i;
    if (col < minGenderCol) continue;
    worksheet.getCell(labelRow, col).value = labels[i];
    worksheet.getCell(labelRow + 1, col).value = values[i];
  }

  const total = pickFormWTamilNaduGenderCount(
    counts.total,
    values[0] + values[1] + values[2] + values[3]
  );
  for (let r = Math.max(1, labelRow - 2); r <= labelRow; r += 1) {
    for (let c = Math.max(minGenderCol, startCol - 2); c <= startCol + 4; c += 1) {
      const raw = excelCellToPlainText(worksheet.getCell(r, c)?.value);
      if (/total\s+number\s+of\s+persons?\s+employed/i.test(raw)) {
        const labelOnly = raw.split(':')[0].trim();
        worksheet.getCell(r, c).value = `${labelOnly}: ${total}`;
        return true;
      }
    }
  }
  return true;
}

const FORM_W_TN_MONTH_NAMES = [
  'January',
  'February',
  'March',
  'April',
  'May',
  'June',
  'July',
  'August',
  'September',
  'October',
  'November',
  'December',
];

function ordinalDayFormW(n) {
  const d = Number(n);
  if (!Number.isFinite(d)) return String(n ?? '');
  const j = d % 10;
  const k = d % 100;
  if (j === 1 && k !== 11) return `${d}st`;
  if (j === 2 && k !== 12) return `${d}nd`;
  if (j === 3 && k !== 13) return `${d}rd`;
  return `${d}th`;
}

/** Resolve { monthName, year, iso } from YYYY-MM or selected month + year. */
export function resolveFormWTamilNaduPeriodParts(payrollMonthIso = '', fallbackYear = null) {
  const m = String(payrollMonthIso || '').trim();
  if (/^\d{4}-\d{2}$/.test(m)) {
    const year = parseInt(m.slice(0, 4), 10);
    const monthIdx = parseInt(m.slice(5, 7), 10) - 1;
    if (monthIdx >= 0 && monthIdx < 12 && Number.isFinite(year)) {
      return {
        iso: m,
        monthName: FORM_W_TN_MONTH_NAMES[monthIdx],
        year: String(year),
        monthIdx,
      };
    }
  }
  return {
    iso: '',
    monthName: '',
    year: fallbackYear != null ? String(fallbackYear) : '',
    monthIdx: -1,
  };
}

/** "Wage Period from 1st April 2026 to 30th April 2026" */
export function buildFormWTamilNaduWagePeriodLine(payrollMonthIso) {
  const parts = resolveFormWTamilNaduPeriodParts(payrollMonthIso);
  if (!parts.monthName || !parts.year || parts.monthIdx < 0) return '';
  const year = Number(parts.year);
  const lastDay = new Date(year, parts.monthIdx + 1, 0).getDate();
  return `Wage Period from ${ordinalDayFormW(1)} ${parts.monthName} ${year} to ${ordinalDayFormW(lastDay)} ${parts.monthName} ${year}`;
}

export function applyFormWTamilNaduMonthYearToHeaderFormData(headerFormData, payrollMonthIso) {
  const out = { ...(headerFormData || {}) };
  const parts = resolveFormWTamilNaduPeriodParts(payrollMonthIso);
  if (parts.monthName) {
    out.form_x_month = parts.monthName;
    out.form_w_tn_month = parts.monthName;
  }
  if (parts.year) {
    out.form_x_year = parts.year;
    out.form_w_tn_year = parts.year;
  }
  return out;
}

export function ensureFormWTamilNaduMonthYearHeaderFields(fields) {
  const list = Array.isArray(fields) ? [...fields] : [];
  const hasMonth = list.some(
    (f) =>
      f?.key === 'form_x_month' ||
      f?.key === 'form_w_tn_month' ||
      /^month\s*:?\s*$/i.test(String(f?.label || '').trim())
  );
  const hasYear = list.some(
    (f) =>
      f?.key === 'form_x_year' ||
      f?.key === 'form_w_tn_year' ||
      /^year\s*:?\s*$/i.test(String(f?.label || '').trim())
  );
  if (!hasMonth) list.push({ label: 'Month:', value: '', key: 'form_x_month' });
  if (!hasYear) list.push({ label: 'Year:', value: '', key: 'form_x_year' });
  return list;
}

/**
 * On download: rewrite incomplete "Wage Period from 1st April 202 …" with full year,
 * and fill Month: / Year: adjacent value cells.
 */
export function writeFormWTamilNaduPeriodFieldsToExcelJsWorksheet(worksheet, payrollMonthIso) {
  if (!worksheet) return false;
  const parts = resolveFormWTamilNaduPeriodParts(payrollMonthIso);
  const wageLine = buildFormWTamilNaduWagePeriodLine(payrollMonthIso);
  if (!parts.monthName && !parts.year && !wageLine) return false;

  const maxRows = Math.min(25, worksheet.rowCount || 25);
  const maxCols = Math.min(40, worksheet.columnCount || 40);
  let wroteWage = false;
  let wroteMonth = false;
  let wroteYear = false;

  for (let r = 1; r <= maxRows; r += 1) {
    for (let c = 1; c <= maxCols; c += 1) {
      const raw = excelCellToPlainText(worksheet.getCell(r, c)?.value).replace(/\s+/g, ' ').trim();
      if (!raw) continue;
      const lower = raw.toLowerCase();

      if (wageLine && /wage\s*period\s*from/i.test(lower)) {
        worksheet.getCell(r, c).value = wageLine;
        wroteWage = true;
        continue;
      }

      if (parts.monthName && /^month\s*:?\s*$/i.test(raw)) {
        const adj = worksheet.getCell(r, c + 1);
        const adjText = excelCellToPlainText(adj?.value).trim();
        if (!adjText || /^enter\b/i.test(adjText)) {
          adj.value = parts.monthName;
        } else {
          worksheet.getCell(r, c).value = `Month: ${parts.monthName}`;
        }
        wroteMonth = true;
        continue;
      }
      if (parts.monthName && /^month\s*:/i.test(raw) && !/year/i.test(raw)) {
        worksheet.getCell(r, c).value = `Month: ${parts.monthName}`;
        wroteMonth = true;
        continue;
      }

      if (parts.year && /^year\s*:?\s*$/i.test(raw)) {
        const adj = worksheet.getCell(r, c + 1);
        const adjText = excelCellToPlainText(adj?.value).trim();
        if (!adjText || /^enter\b/i.test(adjText)) {
          adj.value = parts.year;
        } else {
          worksheet.getCell(r, c).value = `Year: ${parts.year}`;
        }
        wroteYear = true;
        continue;
      }
      if (parts.year && /^year\s*:/i.test(raw) && !/month/i.test(raw)) {
        worksheet.getCell(r, c).value = `Year: ${parts.year}`;
        wroteYear = true;
      }
    }
  }

  return wroteWage || wroteMonth || wroteYear;
}

export function formWTamilNaduHeaderNorm(header) {
  return String(header || '')
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

/** FORM - W col: "Number of days worked". */
export function isFormWTamilNaduDaysWorkedHeader(header) {
  const h = formWTamilNaduHeaderNorm(header);
  if (!h) return false;
  if (/overtime|\bot\b|hours?\s+worked/.test(h)) return false;
  if (/basic|dearness|allowance|wage|identification|employee\s*id|name/.test(h) && !/days/.test(h)) {
    return false;
  }
  return (
    (h.includes('number') && h.includes('days') && h.includes('worked')) ||
    (h.includes('no') && h.includes('days') && h.includes('worked')) ||
    (h.includes('days') && h.includes('worked')) ||
    (/number\s+of\s+days/.test(h) && !/hours/.test(h)) ||
    (/no\.?\s*of\s+days/.test(h) && !/hours/.test(h))
  );
}

/** Prefer named "Number of days worked"; else the column between Employee ID and Basic Wage. */
export function findFormWTamilNaduDaysWorkedHeaders(headers) {
  const hdrs = Array.isArray(headers) ? headers : [];
  const named = hdrs.filter((h) => isFormWTamilNaduDaysWorkedHeader(h));
  if (named.length > 0) return named;
  const isId = (h) => {
    const t = formWTamilNaduHeaderNorm(h);
    return (
      /identification|employee\s*id|emp\s*id|worker\s*id|employee\s*number/.test(t) &&
      !/zoho/.test(t)
    );
  };
  const isBasic = (h) => {
    const t = formWTamilNaduHeaderNorm(h);
    return t.includes('basic') && /wage|wag/.test(t);
  };
  const idIdx = hdrs.findIndex(isId);
  const basicIdx = hdrs.findIndex(isBasic);
  if (idIdx >= 0 && basicIdx === idIdx + 2) return [hdrs[idIdx + 1]];
  if (idIdx >= 0 && basicIdx > idIdx + 1 && basicIdx - idIdx === 2) return [hdrs[idIdx + 1]];
  return [];
}

function mergeFormWTamilNaduPayrollPayload(payrollRow) {
  if (!payrollRow || typeof payrollRow !== 'object') return payrollRow;
  const payload = payrollRow.payroll_payload;
  if (typeof payload === 'string') {
    try {
      return { ...payrollRow, ...JSON.parse(payload) };
    } catch {
      return payrollRow;
    }
  }
  if (payload && typeof payload === 'object') return { ...payrollRow, ...payload };
  return payrollRow;
}

/** Sample Payroll Paid_days (flat row, payload, or common aliases). */
export function resolveFormWTamilNaduPaidDays(payrollRow) {
  if (!payrollRow || typeof payrollRow !== 'object' || payrollRow.fetch_error) return '';
  const merged = mergeFormWTamilNaduPayrollPayload(payrollRow);
  const nested =
    merged.employee && typeof merged.employee === 'object' && !Array.isArray(merged.employee)
      ? merged.employee
      : null;
  const sources = [flattenPayrollEarningColumns(merged), merged];
  if (nested) sources.push(nested);
  const keys = [
    'Paid_days',
    'paid_days',
    'Paid Days',
    'paidDays',
    'PaidDays',
    'paid_days_in_month',
    'paidDaysInMonth',
    'days_worked',
    'Days Worked',
    'daysWorked',
    'no_of_days_worked',
    'effective_paid_days',
  ];
  const patterns = [
    /^paid_days$/,
    /^paiddays$/,
    /paid_days/,
    /paid_days_in_month/,
    /daysworked/,
    /days_present/,
    /noofdayspresent/,
    /effective_paid_days/,
  ];
  for (let i = 0; i < sources.length; i += 1) {
    const src = sources[i];
    if (!src || typeof src !== 'object') continue;
    const raw = readPayrollScalar(src, keys, patterns);
    if (raw === '' || raw == null) continue;
    const num = Number(String(raw).replace(/,/g, '').trim());
    if (Number.isFinite(num) && num >= 0) return num;
  }
  return '';
}

export function readFormWTamilNaduPayrollMonthIso(row) {
  if (!row || typeof row !== 'object') return '';
  const keys = [
    'payroll_month',
    'Payroll_Month',
    'payrollMonth',
    'salary_month',
    'Salary_Month',
    'yearmonth',
    'year_month',
    'YearMonth',
    'monthFilter',
    'MonthFilter',
    'monthfilter',
    'month',
    'Month',
  ];
  for (let i = 0; i < keys.length; i += 1) {
    const raw = String(row[keys[i]] ?? '').trim();
    if (!raw) continue;
    const iso = raw.match(/^(\d{4}-\d{2})/);
    if (iso) return iso[1];
  }
  return '';
}

export function filterFormWTamilNaduPayrollRowsForMonth(records, monthIso) {
  const rows = (Array.isArray(records) ? records : []).filter(
    (row) => row && typeof row === 'object' && row.fetch_error !== true
  );
  const want = String(monthIso || '').trim();
  if (!/^\d{4}-\d{2}$/.test(want)) return rows;
  const tagged = rows.filter((row) => readFormWTamilNaduPayrollMonthIso(row));
  if (tagged.length === 0) return rows;
  const matchedTagged = rows.filter((row) => readFormWTamilNaduPayrollMonthIso(row) === want);
  // Keep untagged rows too — Sample Payroll sometimes omits PayrollMonth on one
  // employee (e.g. VE0042 / Vinu) while the rest of the batch is tagged. Dropping
  // untagged rows left Number of days worked blank for that person only.
  const untagged = rows.filter((row) => !readFormWTamilNaduPayrollMonthIso(row));
  if (matchedTagged.length > 0) {
    if (untagged.length === 0) return matchedTagged;
    const seen = new Set(matchedTagged);
    return matchedTagged.concat(untagged.filter((row) => !seen.has(row)));
  }
  // Fetcher often returns a single already-scoped Sample Payroll month while the
  // form's primary month candidate differs (due date ±1). Keep that batch.
  const taggedMonths = Array.from(
    new Set(tagged.map((row) => readFormWTamilNaduPayrollMonthIso(row)).filter(Boolean))
  );
  if (taggedMonths.length === 1) return rows;
  return matchedTagged;
}

function formWTamilNaduNormName(value) {
  return String(value || '')
    .replace(/\r?\n/g, ' ')
    .replace(/[.,]/g, ' ')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ');
}

function formWTamilNaduSplitFirstLast(text) {
  const parts = formWTamilNaduNormName(text)
    .split(' ')
    .filter(Boolean);
  if (parts.length < 2) return { firstName: '', lastName: '' };
  return { firstName: parts[0], lastName: parts[parts.length - 1] };
}

/** Soft last-name match: exact, or one is a prefix of the other (Muruganath / Muruganatham). */
function formWTamilNaduLastNamesSoftMatch(left, right) {
  const a = formWTamilNaduNormName(left);
  const b = formWTamilNaduNormName(right);
  if (!a || !b) return false;
  if (a === b) return true;
  if (a.length >= 6 && b.length >= 6 && (a.startsWith(b) || b.startsWith(a))) return true;
  return false;
}

/**
 * Soft first-name match: exact, or same leading token.
 * People often store "Vinu Monikandan" while Sample Payroll splits to first_name "Vinu".
 */
function formWTamilNaduFirstNamesSoftMatch(left, right) {
  const a = formWTamilNaduNormName(left);
  const b = formWTamilNaduNormName(right);
  if (!a || !b) return false;
  if (a === b) return true;
  const a0 = a.split(' ').filter(Boolean)[0] || '';
  const b0 = b.split(' ').filter(Boolean)[0] || '';
  return Boolean(a0 && b0 && a0 === b0);
}

function formWTamilNaduEdgeNamesMatch(leftName, rightName) {
  const leftEdge = formWTamilNaduSplitFirstLast(leftName);
  const rightEdge = formWTamilNaduSplitFirstLast(rightName);
  if (!leftEdge.firstName || !leftEdge.lastName || !rightEdge.firstName || !rightEdge.lastName) {
    return false;
  }
  return (
    formWTamilNaduNormName(leftEdge.firstName) === formWTamilNaduNormName(rightEdge.firstName) &&
    formWTamilNaduLastNamesSoftMatch(leftEdge.lastName, rightEdge.lastName)
  );
}

/**
 * "Vinu Monikandan" matches "Vinu Monikandan Muruganatham" (shorter is a token prefix).
 * Requires 2+ tokens and the same first token — never first-name-only.
 */
function formWTamilNaduTokenPrefixMatch(leftName, rightName) {
  const aParts = formWTamilNaduNormName(leftName).split(' ').filter(Boolean);
  const bParts = formWTamilNaduNormName(rightName).split(' ').filter(Boolean);
  if (aParts.length < 2 || bParts.length < 2) return false;
  if (aParts[0] !== bParts[0]) return false;
  const shorter = aParts.length <= bParts.length ? aParts : bParts;
  const longer = aParts.length <= bParts.length ? bParts : aParts;
  return shorter.every((tok, i) => tok === longer[i]);
}

/**
 * Build FirstName + MiddleName + LastName match parts from People and/or form
 * "Name of the Employee". Prefer the form column name when present.
 */
export function resolveFormWTamilNaduNameMatchParts(employeeOrRow, extraParts = null) {
  const emp = readFormWTamilNaduPersonNameParts(employeeOrRow);
  const extra = extraParts && typeof extraParts === 'object' ? extraParts : {};

  let firstName = String(extra.firstName || emp.firstName || '').trim();
  let middleName = String(extra.middleName || emp.middleName || '').trim();
  let lastName = String(extra.lastName || emp.lastName || '').trim();
  // Form "Name of the Employee" wins as full display name.
  let fullName = String(extra.fullName || '').trim();
  if (!fullName) {
    fullName = [firstName, middleName, lastName].filter(Boolean).join(' ').trim() || emp.fullName;
  }

  // Full display name only in FirstName (blank Middle/Last) → split edge tokens.
  if (firstName && !lastName && !middleName && /\s/.test(firstName)) {
    const split = formWTamilNaduSplitFirstLast(firstName);
    if (split.firstName && split.lastName) {
      const tokens = String(firstName)
        .replace(/\r?\n/g, ' ')
        .replace(/\s+/g, ' ')
        .trim()
        .split(' ')
        .filter(Boolean);
      firstName = tokens[0] || split.firstName;
      lastName = tokens[tokens.length - 1] || split.lastName;
      if (tokens.length > 2) middleName = tokens.slice(1, -1).join(' ');
    }
  }
  if ((!firstName || !lastName) && fullName) {
    const tokens = String(fullName)
      .replace(/\r?\n/g, ' ')
      .replace(/\s+/g, ' ')
      .trim()
      .split(' ')
      .filter(Boolean);
    if (tokens.length >= 2) {
      if (!firstName) firstName = tokens[0];
      if (!lastName) lastName = tokens[tokens.length - 1];
      if (!middleName && tokens.length > 2) middleName = tokens.slice(1, -1).join(' ');
    }
  }
  // Multi-token FirstName with LastName but no MiddleName → peel middle out.
  if (firstName && lastName && !middleName && /\s/.test(firstName)) {
    const tokens = String(firstName)
      .replace(/\r?\n/g, ' ')
      .replace(/\s+/g, ' ')
      .trim()
      .split(' ')
      .filter(Boolean);
    if (tokens.length >= 2) {
      firstName = tokens[0];
      middleName = tokens.slice(1).join(' ');
    }
  }
  if (!fullName) {
    fullName = [firstName, middleName, lastName].filter(Boolean).join(' ').trim();
  }
  return { firstName, middleName, lastName, fullName };
}

/** Collect every plausible display name for matching (form + People + MiddleName). */
export function collectFormWTamilNaduNameCandidates(employeeOrRow, extraParts = null) {
  const names = [];
  const seen = new Set();
  const push = (raw) => {
    const n = formWTamilNaduNormName(raw);
    if (!n || seen.has(n)) return;
    seen.add(n);
    names.push(String(raw || '').replace(/\r?\n/g, ' ').replace(/\s+/g, ' ').trim());
  };

  const parts = resolveFormWTamilNaduNameMatchParts(employeeOrRow, extraParts);
  push(parts.fullName);
  if (parts.firstName && parts.lastName) {
    push(`${parts.firstName} ${parts.lastName}`);
    if (parts.middleName) {
      push(`${parts.firstName} ${parts.middleName} ${parts.lastName}`);
      push(`${parts.firstName} ${parts.middleName}`);
    }
  }

  const extra = extraParts && typeof extraParts === 'object' ? extraParts : {};
  push(extra.fullName);
  if (extra.firstName && extra.lastName) {
    push(`${extra.firstName} ${extra.lastName}`);
    if (extra.middleName) {
      push(`${extra.firstName} ${extra.middleName} ${extra.lastName}`);
    }
  }

  const emp = employeeOrRow && typeof employeeOrRow === 'object' ? employeeOrRow : {};
  const empParts = readFormWTamilNaduPersonNameParts(emp);
  push(empParts.fullName);
  if (empParts.firstName && empParts.lastName) {
    push(`${empParts.firstName} ${empParts.lastName}`);
    if (empParts.middleName) {
      push(`${empParts.firstName} ${empParts.middleName} ${empParts.lastName}`);
    }
  }
  push(formatFormWTamilNaduEmployeeName(emp));

  // Form W TN default-payroll aliases (e.g. "Vinu Monikandan" / "Vinu Monik") so Paid_days
  // can resolve the same people whose Basic is filled from FORM_W_TN_DEFAULT_PAYROLL.
  const empId = resolveFormWTamilNaduEmployeeId(emp);
  FORM_W_TN_DEFAULT_PAYROLL.forEach((entry) => {
    const idHit =
      empId &&
      (entry.ids || []).some((id) => String(id).trim().toUpperCase() === empId);
    const nameHit = (entry.names || []).some((name) =>
      personNamesMatch(formatFormWTamilNaduEmployeeName(emp) || empParts.fullName || '', name)
    );
    if (!idHit && !nameHit) return;
    (entry.names || []).forEach((name) => push(name));
  });
  return names;
}

/**
 * First + Middle + Last token match against a payroll display name.
 * Requires first AND last; when middle token(s) exist they should appear in payroll
 * (soft: also allow first+last when Sample Payroll omitted the middle).
 */
function formWTamilNaduFirstMiddleLastMatch(leftParts, rightName) {
  const peopleTokens = [
    ...formWTamilNaduNormName(leftParts.firstName).split(' '),
    ...formWTamilNaduNormName(leftParts.middleName).split(' '),
    ...formWTamilNaduNormName(leftParts.lastName).split(' '),
  ].filter(Boolean);
  if (peopleTokens.length < 2) return false;
  const leftFirst = peopleTokens[0];
  const leftLast = peopleTokens[peopleTokens.length - 1];
  const middleTokens = peopleTokens.slice(1, -1);
  const tokens = formWTamilNaduNormName(rightName).split(' ').filter(Boolean);
  if (tokens.length < 2) return false;
  const payFirst = tokens[0];
  const payLast = tokens[tokens.length - 1];
  if (payFirst !== leftFirst) return false;
  if (!formWTamilNaduLastNamesSoftMatch(payLast, leftLast)) return false;
  if (middleTokens.length === 0) return true;
  // Prefer middle present in payroll; still accept first+last when Sample omits middle.
  if (middleTokens.every((tok) => tokens.includes(tok))) return true;
  return tokens.length === 2;
}

/**
 * Form W TN name match (Paid_days) — FirstName + MiddleName + LastName / fullname.
 * Handles middle names and near-spellings (Muruganath / Muruganatham).
 */
export function formWTamilNaduNamesMatch(employeeOrRow, payrollRow, extraParts = null) {
  if (!payrollRow || typeof payrollRow !== 'object' || payrollRow.fetch_error) return false;
  const right = resolveFormWTamilNaduNameMatchParts(payrollRow);
  const payNames = [
    right.fullName,
    right.firstName && right.middleName && right.lastName
      ? `${right.firstName} ${right.middleName} ${right.lastName}`
      : '',
    right.firstName && right.lastName ? `${right.firstName} ${right.lastName}` : '',
    payrollRow.employee_name,
    payrollRow.EmployeeName,
    payrollRow.full_name,
    payrollRow.name,
  ]
    .map((v) => String(v || '').replace(/\r?\n/g, ' ').replace(/\s+/g, ' ').trim())
    .filter(Boolean);

  const candidates = collectFormWTamilNaduNameCandidates(employeeOrRow, extraParts);
  if (candidates.length === 0 || payNames.length === 0) return false;

  for (let i = 0; i < candidates.length; i += 1) {
    const leftFull = formWTamilNaduNormName(candidates[i]);
    for (let j = 0; j < payNames.length; j += 1) {
      const rightFull = formWTamilNaduNormName(payNames[j]);
      if (leftFull && rightFull && leftFull === rightFull) return true;
      if (formWTamilNaduEdgeNamesMatch(candidates[i], payNames[j])) return true;
      // People/default alias "Vinu Monikandan" ↔ Sample "Vinu Monikandan Muruganatham"
      if (formWTamilNaduTokenPrefixMatch(candidates[i], payNames[j])) return true;
    }
  }

  const left = resolveFormWTamilNaduNameMatchParts(employeeOrRow, extraParts);
  for (let j = 0; j < payNames.length; j += 1) {
    if (formWTamilNaduFirstMiddleLastMatch(left, payNames[j])) return true;
  }

  // Direct first + middle + last (soft first token / soft last):
  // People FirstName "Vinu Monikandan" or First+Middle "Vinu"+"Monikandan" + Last "Muruganatham"
  // ↔ Sample first_name "Vinu" + last_name "Muruganatham".
  if (
    left.firstName &&
    left.lastName &&
    right.firstName &&
    right.lastName &&
    formWTamilNaduFirstNamesSoftMatch(left.firstName, right.firstName) &&
    formWTamilNaduLastNamesSoftMatch(left.lastName, right.lastName)
  ) {
    const leftMiddle = formWTamilNaduNormName(left.middleName);
    const rightMiddle = formWTamilNaduNormName(right.middleName);
    // When both sides expose a middle name, require they agree (soft / contained).
    if (leftMiddle && rightMiddle) {
      const leftMids = leftMiddle.split(' ').filter(Boolean);
      const rightMids = rightMiddle.split(' ').filter(Boolean);
      if (
        leftMids.every((tok) => rightMids.includes(tok)) ||
        rightMids.every((tok) => leftMids.includes(tok))
      ) {
        return true;
      }
      return false;
    }
    return true;
  }
  return false;
}

/**
 * Sample Payroll match for Form W TN Paid_days.
 * Prefer FirstName + LastName / fullname; last resort = form Employee ID / GID
 * (same VE0042 path that already fills Basic for Vinu Monikandan Muruganatham).
 */
export function findFormWTamilNaduPayrollRowByFirstAndLastName(
  employeeOrRow,
  records,
  extraParts = null,
  options = {}
) {
  const monthIso = options.monthIso || options.payrollMonth || '';
  const allRows = (Array.isArray(records) ? records : []).filter(
    (row) => row && typeof row === 'object' && row.fetch_error !== true
  );
  let rows = filterFormWTamilNaduPayrollRowsForMonth(records, monthIso);
  if (rows.length === 0) rows = allRows;
  if (rows.length === 0) return null;

  const tryMatch = (list) => {
    // Prefer the form "Name of the Employee" first — avoids People FirstName truncation
    // (e.g. "Vinu Monikandan" without Muruganatham) missing Sample Payroll Paid_days.
    if (extraParts && String(extraParts.fullName || '').trim()) {
      const fromForm = list.find((row) => formWTamilNaduNamesMatch({}, row, extraParts));
      if (fromForm && !fromForm.fetch_error) return fromForm;
    }

    const parts = resolveFormWTamilNaduNameMatchParts(employeeOrRow, extraParts);
    if (parts.firstName && parts.lastName) {
      const named =
        list.find((row) => formWTamilNaduNamesMatch(employeeOrRow, row, extraParts)) ||
        findForm10PayrollRowByFirstAndLastName({}, list, parts) ||
        null;
      if (named && !named.fetch_error) return named;
    } else {
      const namedOnly = list.find((row) => formWTamilNaduNamesMatch(employeeOrRow, row, extraParts));
      if (namedOnly && !namedOnly.fetch_error) return namedOnly;
    }

    // Last resort: Employee Identification No. / People EmployeeID ↔ Sample Payroll GID.
    // Needed when name tokens diverge but Basic already resolved via VE0042 defaults.
    const defaultIds = [];
    const empForDefault =
      employeeOrRow && typeof employeeOrRow === 'object' ? employeeOrRow : {};
    const empIdForDefault = resolveFormWTamilNaduEmployeeId(empForDefault);
    const empNameForDefault = formatFormWTamilNaduEmployeeName(empForDefault);
    const partsForDefault = resolveFormWTamilNaduNameMatchParts(employeeOrRow, extraParts);
    FORM_W_TN_DEFAULT_PAYROLL.forEach((entry) => {
      const idHit =
        empIdForDefault &&
        (entry.ids || []).some((id) => String(id).trim().toUpperCase() === empIdForDefault);
      const nameHit = (entry.names || []).some((name) =>
        personNamesMatch(empNameForDefault || partsForDefault.fullName || '', name)
      );
      const formId = String(
        (options && options.employeeId) || (extraParts && extraParts.employeeId) || ''
      )
        .trim()
        .toUpperCase();
      const formIdHit =
        formId && (entry.ids || []).some((id) => String(id).trim().toUpperCase() === formId);
      if (!idHit && !nameHit && !formIdHit) return;
      (entry.ids || []).forEach((id) => {
        const v = String(id || '').trim();
        if (v) defaultIds.push(v);
      });
    });
    const idCandidates = Array.from(
      new Set(
        [
          options.employeeId,
          extraParts && extraParts.employeeId,
          resolveFormWTamilNaduEmployeeId(employeeOrRow),
          employeeOrRow && employeeOrRow.EmployeeID,
          employeeOrRow && employeeOrRow['Employee ID'],
          employeeOrRow && employeeOrRow.employee_id,
          employeeOrRow && employeeOrRow.employee_number,
          ...defaultIds,
        ]
          .map((v) => String(v || '').trim())
          .filter(Boolean)
      )
    );
    for (let i = 0; i < idCandidates.length; i += 1) {
      const hit = list.find((row) => samplePayrollRowMatchesEmployeeId(row, idCandidates[i]));
      if (hit && !hit.fetch_error) return hit;
    }
    return null;
  };

  const hit = tryMatch(rows);
  if (hit) {
    // Prefer a hit that actually has Paid_days; otherwise keep searching.
    if (resolveFormWTamilNaduPaidDays(hit) !== '') return hit;
  }

  // Month filter / wrong tagged month can hide the Paid_days row — retry unfiltered.
  if (monthIso && rows !== allRows && allRows.length > rows.length) {
    const unfiltered = tryMatch(allRows);
    if (unfiltered && resolveFormWTamilNaduPaidDays(unfiltered) !== '') return unfiltered;
    if (!hit && unfiltered) return unfiltered;
  }
  return hit;
}

/**
 * FORM - W "Number of days worked" ← Sample Payroll Paid_days.
 * Match by firstname + lastname first; fall back to Employee ID / GID when needed.
 */
export function applyFormWTamilNaduPaidDaysToMappedRows(
  rows,
  headers,
  payrollRows = [],
  options = {}
) {
  const hdrs = Array.isArray(headers) ? headers : [];
  let daysHeaders = findFormWTamilNaduDaysWorkedHeaders(hdrs);
  if (daysHeaders.length === 0 && Array.isArray(rows) && rows[0] && typeof rows[0] === 'object') {
    daysHeaders = findFormWTamilNaduDaysWorkedHeaders(Object.keys(rows[0]));
  }
  if (daysHeaders.length === 0 || !Array.isArray(rows) || rows.length === 0) {
    return { paidDaysHits: 0 };
  }

  const overwrite = options.overwrite !== false;
  const monthIso = options.monthIso || options.payrollMonth || '';
  const unwrapEmp =
    typeof options.unwrapEmp === 'function'
      ? options.unwrapEmp
      : (item) => (item && (item.Employee || item.employee || item)) || null;
  const employeesForMapping = Array.isArray(options.employeesForMapping)
    ? options.employeesForMapping
    : [];
  const resolvePayrollRow =
    typeof options.resolvePayrollRow === 'function' ? options.resolvePayrollRow : null;
  const payrollList = filterFormWTamilNaduPayrollRowsForMonth(payrollRows, monthIso);
  const idHeader = hdrs.find((h) => {
    const t = formWTamilNaduHeaderNorm(h);
    return (
      /identification|employee\s*id|emp\s*id|worker\s*id|employee\s*number/.test(t) &&
      !/zoho/.test(t)
    );
  });

  const cellEmpty = (value) => {
    const s = String(value ?? '').trim();
    if (!s) return true;
    return /^enter\b/i.test(s);
  };

  let paidDaysHits = 0;
  rows.forEach((row, index) => {
    if (!row || typeof row !== 'object') return;
    const emp = unwrapEmp(employeesForMapping[index] || null);
    const formEmployeeId = idHeader ? String(row[idHeader] ?? '').trim() : '';
    const extraParts = {
      ...collectForm10RowNameParts(row, hdrs),
      employeeId: formEmployeeId,
    };
    let payrollRow = resolvePayrollRow
      ? resolvePayrollRow(emp, row, index) || null
      : findFormWTamilNaduPayrollRowByFirstAndLastName(emp || row, payrollList, extraParts, {
          monthIso,
          employeeId: formEmployeeId,
        });
    let paid = payrollRow && !payrollRow.fetch_error ? resolveFormWTamilNaduPaidDays(payrollRow) : '';
    // Custom resolver / month filter can miss VE0042 — retry against full Sample Payroll.
    if (paid === '' || paid == null) {
      const fallback = findFormWTamilNaduPayrollRowByFirstAndLastName(
        emp || row,
        payrollRows,
        extraParts,
        { monthIso: '', employeeId: formEmployeeId }
      );
      const fallbackPaid =
        fallback && !fallback.fetch_error ? resolveFormWTamilNaduPaidDays(fallback) : '';
      if (fallbackPaid !== '' && fallbackPaid != null) {
        payrollRow = fallback;
        paid = fallbackPaid;
      }
    }
    // VE0042 / Vinu: month defaults (Apr 30, May 31, Jun 30, Jul 31, Aug 31) when payroll miss.
    if (paid === '' || paid == null) {
      const defaultPaid = resolveFormWTamilNaduDefaultPaidDays(emp || row, monthIso, {
        ...extraParts,
        employeeId: formEmployeeId || extraParts.employeeId,
      });
      if (defaultPaid !== '' && defaultPaid != null) paid = defaultPaid;
    }
    let filledRow = false;
    daysHeaders.forEach((header) => {
      const cur = String(row[header] ?? '').trim();
      if (!overwrite && !cellEmpty(cur)) return;
      if (paid !== '' && paid != null) {
        row[header] = String(paid);
        filledRow = true;
      }
    });
    if (filledRow) paidDaysHits += 1;
  });
  return { paidDaysHits };
}
