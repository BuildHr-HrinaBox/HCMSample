import { flattenPayrollEarningColumns, readPayrollScalar } from '../../utils/payrollEarnings';
import { samplePayrollRowMatchesEmployeeId } from '../../utils/samplePayrollApi';
import { collectForm10RowNameParts } from './form10TamilNadu';
import { findFormXVIIITamilNaduRecordByFirstAndLastName } from './formXVIIITamilNaduWagesMuster';
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
    names: ['Vinu Monikandan', 'Vinu Monik'],
    basic: '43688',
    hra: '20704',
  },
];

function formatFormWTamilNaduEmployeeName(emp = {}) {
  const fn = String(emp.FirstName || emp['FirstName'] || emp.firstName || '').trim();
  const ln = String(emp.LastName || emp['LastName'] || emp.lastName || '').trim();
  if (fn && ln) return `${fn} ${ln}`;
  return fn || ln || String(emp.Name || emp['Name'] || emp.EmployeeName || emp['Employee Name'] || '').trim();
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
  const matched = rows.filter((row) => readFormWTamilNaduPayrollMonthIso(row) === want);
  if (matched.length > 0) return matched;
  // Fetcher often returns a single already-scoped Sample Payroll month while the
  // form's primary month candidate differs (due date ±1). Keep that batch.
  const taggedMonths = Array.from(
    new Set(tagged.map((row) => readFormWTamilNaduPayrollMonthIso(row)).filter(Boolean))
  );
  if (taggedMonths.length === 1) return rows;
  return matched;
}

function formWTamilNaduEmployeeIdCandidates(employeeOrRow, extraParts = null) {
  const src = employeeOrRow && typeof employeeOrRow === 'object' ? employeeOrRow : {};
  const extra = extraParts && typeof extraParts === 'object' ? extraParts : {};
  return Array.from(
    new Set(
      [
        resolveFormWTamilNaduEmployeeId(src),
        extra.employeeId,
        extra.EmployeeID,
        src.EmployeeID,
        src['Employee ID'],
        src.employee_id,
        src.employee_number,
        src.Zoho_ID,
      ]
        .map((v) => String(v || '').trim())
        .filter(Boolean)
    )
  );
}

/** Sample Payroll match: fullname OR firstname + lastname (never first-name-only). */
export function findFormWTamilNaduPayrollRowByFirstAndLastName(
  employeeOrRow,
  records,
  extraParts = null,
  options = {}
) {
  const monthIso = options.monthIso || options.payrollMonth || '';
  const rows = filterFormWTamilNaduPayrollRowsForMonth(records, monthIso);
  if (rows.length === 0) return null;
  // Rows are already month-scoped (incl. single-batch fallback). Do not re-filter
  // with Form XVIII's stricter month matcher, which can drop the Paid_days batch.
  const named = findFormXVIIITamilNaduRecordByFirstAndLastName(
    employeeOrRow,
    rows,
    extraParts,
    { monthIso: '' }
  );
  if (named && !named.fetch_error) return named;
  const ids = formWTamilNaduEmployeeIdCandidates(employeeOrRow, extraParts);
  for (let i = 0; i < ids.length; i += 1) {
    const hit = rows.find((row) => samplePayrollRowMatchesEmployeeId(row, ids[i]));
    if (hit && !hit.fetch_error) return hit;
  }
  return null;
}

/**
 * FORM - W "Number of days worked" ← Sample Payroll Paid_days.
 * Match People / form row to payroll by firstname + lastname.
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
    const extraParts = {
      ...collectForm10RowNameParts(row, hdrs),
      employeeId: idHeader ? String(row[idHeader] ?? '').trim() : '',
    };
    const payrollRow = resolvePayrollRow
      ? resolvePayrollRow(emp, row, index) || null
      : findFormWTamilNaduPayrollRowByFirstAndLastName(emp || row, payrollList, extraParts, {
          monthIso,
        });
    const paid = payrollRow && !payrollRow.fetch_error ? resolveFormWTamilNaduPaidDays(payrollRow) : '';
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
