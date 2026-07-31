import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { labelMatchScore } from './form18APAccidentNotice';
import {
  flattenPayrollEarningColumns,
  mergePayrollRunEmployeePayload,
  payrollRowHasWageBreakdown,
  readForm10GrossPayAmount,
  readPayrollForm15WageAmounts,
  readPayrollScalar,
  readStrictPayrollNetPay,
} from '../../utils/payrollEarnings';
import { personNamesMatch } from './formFKarnataka';

/** Karnataka Form Q — establishment/employer header + employee particulars table. */

/** April default Basic/Total for Karnataka Form Q when payroll lacks wage columns. */
export const FORM_Q_KARNATAKA_APR_DEFAULT_PAYROLL = [
  { name: 'Suresh Kumar S', basic: '55443', total: '172569' },
  { name: 'Vaikundamoni M', basic: '30108', total: '103828' },
  { name: 'Satheesh Kumar S', basic: '28308', total: '102777' },
  { name: 'Stalin T', basic: '30110', total: '108284' },
  { name: 'Sathishkumar Murugan', basic: '28339', total: '101120' },
];

export function isAprilPayrollMonthCandidates(monthCandidates) {
  const primary = String(
    (Array.isArray(monthCandidates) ? monthCandidates[0] : monthCandidates) || ''
  ).trim();
  return /-04$/.test(primary);
}

export function resolveFormQKarnatakaAprilDefaultWages(emp) {
  const empName = formatFormQKarnatakaEmployeeName(emp);
  if (!empName) return null;
  const match = FORM_Q_KARNATAKA_APR_DEFAULT_PAYROLL.find((entry) =>
    personNamesMatch(empName, entry.name)
  );
  if (!match) return null;
  return {
    basic: match.basic != null ? String(match.basic) : '',
    vda: '',
    other: '',
    total: match.total != null ? String(match.total) : '',
  };
}

export function applyFormQKarnatakaAprilDefaultWages(wages, emp, monthCandidates) {
  if (!isAprilPayrollMonthCandidates(monthCandidates)) return wages;
  const defaults = resolveFormQKarnatakaAprilDefaultWages(emp);
  if (!defaults) return wages;
  const out = { ...(wages || { basic: '', vda: '', other: '', total: '' }) };
  if (defaults.basic !== undefined) out.basic = defaults.basic;
  if (defaults.total) out.total = defaults.total;
  return out;
}

/** Form Q Karnataka Total — used as Normal rate of wages on Form XXIII overtime register. */
export function resolveFormQKarnatakaTotalWageForEmployee(emp, payrollRow = null, monthCandidates = null) {
  const wages = applyFormQKarnatakaAprilDefaultWages(
    resolveFormQKarnatakaPayrollWages(payrollRow),
    emp,
    monthCandidates
  );
  return String(wages?.total ?? '').trim();
}

export function isFormXXIIIKarnatakaContext(formHeader, rowItem, fileName, sheetText = '') {
  const parts = [
    rowItem?.formName,
    rowItem?.FormName,
    rowItem?.state,
    rowItem?.State,
    fileName,
    formHeader?.title,
    formHeader?.subtitle,
    sheetText,
  ]
    .filter((x) => x != null && String(x).trim() !== '')
    .join(' ')
    .toLowerCase();
  const hasKarnataka = /karnataka/.test(parts);
  // No trailing \b — filenames like Form_XXIII_-_Karnataka.xlsx use underscores after xxiii.
  const hasXXIII =
    /\bform[\s._-]*xxiii/i.test(parts) || /register\s+of\s+overtime/.test(parts);
  return hasKarnataka && hasXXIII;
}

export function formQKarnatakaHeaderNorm(txt) {
  return String(txt || '')
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

export function isFormQKarnatakaContext(formHeader, rowItem, fileName, sheetText = '') {
  const parts = [
    rowItem?.formName,
    rowItem?.FormName,
    rowItem?.formFileName,
    rowItem?.FormFileName,
    rowItem?.state,
    rowItem?.State,
    rowItem?.act,
    rowItem?.Act,
    rowItem?.actName,
    rowItem?.ActName,
    fileName,
    formHeader?.title,
    formHeader?.subtitle,
    formHeader?.reference,
    sheetText,
  ]
    .filter((x) => x != null && String(x).trim() !== '')
    .join(' ')
    .toLowerCase();

  const hasFormQ = /\bform[\s._-]*q\b/i.test(parts);
  const hasKarnataka = /karnataka/.test(parts);
  // Karnataka Form Q is the employment / appointment-order particulars register.
  const hasAppointmentOrder = /appointment\s+order/.test(parts);

  const hasAttendanceGrid =
    /full\s+name\s+of\s+the\s+worker/.test(parts) &&
    (/date\s+of\s+the\s+month|working\s+hours|interval\s+for\s+rest/.test(parts));

  const hasEmploymentParticulars =
    /name\s+of\s+the\s+employee/.test(parts) &&
    (/his\/her\s+postal\s+address|father\s*\/\s*husband\s+name|rates\s+of\s+wages\s+payable/.test(
      parts
    ) ||
      /name\s*&\s*address\s+of\s+the\s+employer/.test(parts));

  if (hasAttendanceGrid && !hasEmploymentParticulars) return false;
  if (hasFormQ && hasKarnataka) return true;
  if (hasFormQ && hasAppointmentOrder && (hasKarnataka || hasEmploymentParticulars)) return true;
  if (hasFormQ && hasEmploymentParticulars) return true;
  if (hasEmploymentParticulars && /name\s*&\s*address\s+of\s+the\s+establishment/.test(parts)) {
    return true;
  }
  return false;
}

export function isFormQKarnatakaHeaderFieldLayoutFormHeader(formHeader) {
  return !!formHeader?.formQKarnatakaHeaderFieldLayout;
}

export function isFormQKarnatakaTableLayoutFormHeader(formHeader) {
  return !!formHeader?.formQKarnatakaTableLayout;
}

export const FORM_Q_KARNATAKA_CANONICAL_TABLE_HEADERS = [
  'Name of the Employee',
  'His/Her Postal Address',
  'His/Her Permanent Address',
  'Father/Husband Name',
  'Date of Birth',
  'Date of his/her entry into employment',
  'Designation',
  'Nature of work entrusted to him/her',
  'His/Her serial number in the Register of employment',
  'Basic',
  'VDA',
  'Other allowances if any',
  'Total',
];

export const FORM_Q_KARNATAKA_HEADER_SPECS = [
  {
    key: 'form_q_establishment',
    label: '1. Name & Address of the Establishment',
    match: /name\s*&\s*address\s+of\s+the\s+establishment/i,
    fieldType: 'textarea',
    group: 'header',
  },
  {
    key: 'form_q_ka_employer',
    label: '2. Name & Address of the Employer',
    match: /name\s*&\s*address\s+of\s+the\s+employer/i,
    fieldType: 'textarea',
    group: 'header',
  },
  {
    key: 'form_q_ka_place',
    label: 'Place',
    match: /^place\s*:?$/i,
    group: 'footer',
  },
  {
    key: 'form_q_ka_date',
    label: 'Date',
    match: /^date\s*:?$/i,
    group: 'footer',
  },
  {
    key: 'form_q_ka_signature',
    label: 'Signature of employer',
    match: /signature\s+of\s+employer/i,
    group: 'footer',
  },
];

const stripLeadingNumber = (s) =>
  String(s || '')
    .replace(/^\d+[\.\)]\s*/, '')
    .trim();

export function isFormQKarnatakaEmployeeNameHeader(h) {
  const s = formQKarnatakaHeaderNorm(stripLeadingNumber(h));
  return /name\s+of\s+the\s+employee/.test(s);
}

export function isFormQKarnatakaPostalAddressHeader(h) {
  const s = formQKarnatakaHeaderNorm(stripLeadingNumber(h));
  return /postal\s+address/.test(s);
}

export function isFormQKarnatakaPermanentAddressHeader(h) {
  const s = formQKarnatakaHeaderNorm(stripLeadingNumber(h));
  return /permanent\s+address/.test(s);
}

export function isFormQKarnatakaFatherHusbandHeader(h) {
  const s = formQKarnatakaHeaderNorm(stripLeadingNumber(h));
  return /father\s*\/\s*husband/.test(s);
}

export function isFormQKarnatakaDateOfBirthHeader(h) {
  const s = formQKarnatakaHeaderNorm(stripLeadingNumber(h));
  return /date\s+of\s+birth/.test(s);
}

export function isFormQKarnatakaEntryDateHeader(h) {
  const s = formQKarnatakaHeaderNorm(stripLeadingNumber(h));
  return /entry\s+into\s+employment/.test(s);
}

export function isFormQKarnatakaDesignationHeader(h) {
  const s = formQKarnatakaHeaderNorm(stripLeadingNumber(h));
  return s === 'designation' || /^designation\b/.test(s);
}

export function isFormQKarnatakaNatureOfWorkHeader(h) {
  const s = formQKarnatakaHeaderNorm(stripLeadingNumber(h));
  return /nature\s+of\s+work/.test(s);
}

export function isFormQKarnatakaSerialNumberHeader(h) {
  const s = formQKarnatakaHeaderNorm(stripLeadingNumber(h));
  return /serial\s+number\s+in\s+the\s+register/.test(s);
}

export function isFormQKarnatakaBasicHeader(h) {
  const s = formQKarnatakaHeaderNorm(stripLeadingNumber(h));
  return s === 'basic' || /^basic\b/.test(s);
}

export function isFormQKarnatakaVdaHeader(h) {
  const s = formQKarnatakaHeaderNorm(stripLeadingNumber(h));
  return s === 'vda' || /^vda\b/.test(s);
}

export function isFormQKarnatakaOtherAllowanceHeader(h) {
  const s = formQKarnatakaHeaderNorm(stripLeadingNumber(h));
  return /other\s+allowances?/.test(s);
}

export function isFormQKarnatakaTotalWageHeader(h) {
  const s = formQKarnatakaHeaderNorm(stripLeadingNumber(h));
  return s === 'total' || /^total\b/.test(s);
}

export function headersIndicateFormQKarnatakaTable(tableHeaders) {
  const joined = (Array.isArray(tableHeaders) ? tableHeaders : [])
    .map((h) => formQKarnatakaHeaderNorm(h))
    .join('\n');
  return (
    /name\s+of\s+the\s+employee/.test(joined) &&
    (/postal\s+address|permanent\s+address/.test(joined) || /father\s*\/\s*husband/.test(joined))
  );
}

export function resolveFormQKarnatakaTableHeaders(tableHeaders) {
  const headers = Array.isArray(tableHeaders) ? tableHeaders : [];
  if (headersIndicateFormQKarnatakaTable(headers)) {
    return [...headers];
  }
  return [...FORM_Q_KARNATAKA_CANONICAL_TABLE_HEADERS];
}

const isNarrativeBlob = (raw) => {
  const n = formQKarnatakaHeaderNorm(raw);
  if (!n) return false;
  return (
    /^form\s*q\b/i.test(n) ||
    /shops?\s*&?\s*commercial\s+establishment/i.test(n) ||
    /karnataka\s+shops/i.test(n) ||
    n.length > 140
  );
};

const isTableLabelBlob = (raw) => {
  const n = formQKarnatakaHeaderNorm(stripLeadingNumber(raw));
  return (
    isFormQKarnatakaEmployeeNameHeader(n) ||
    isFormQKarnatakaPostalAddressHeader(n) ||
    isFormQKarnatakaPermanentAddressHeader(n) ||
    isFormQKarnatakaFatherHusbandHeader(n) ||
    isFormQKarnatakaDateOfBirthHeader(n) ||
    isFormQKarnatakaEntryDateHeader(n) ||
    isFormQKarnatakaDesignationHeader(n) ||
    isFormQKarnatakaNatureOfWorkHeader(n) ||
    isFormQKarnatakaSerialNumberHeader(n) ||
    isFormQKarnatakaBasicHeader(n) ||
    isFormQKarnatakaVdaHeader(n) ||
    isFormQKarnatakaOtherAllowanceHeader(n) ||
    isFormQKarnatakaTotalWageHeader(n)
  );
};

const readValueBesideLabel = (getMergedAwareCellText, row, labelCol, effectiveSheetCols) => {
  const maxC = Math.max(12, effectiveSheetCols || 0);
  for (let c = labelCol + 1; c < Math.min(labelCol + 8, maxC); c += 1) {
    const v = String(getMergedAwareCellText(row, c) || '').trim();
    if (!v || v === ':') continue;
    if (isTableLabelBlob(v) || isNarrativeBlob(v)) break;
    return { value: v, valueCol: c, valueRow: row };
  }
  return { value: '', valueCol: labelCol + 1, valueRow: row };
};

const buildTemplateField = (spec, coords = {}) => ({
  label: spec.label,
  value: coords.value || '',
  key: spec.key,
  group: spec.group || 'header',
  fieldType: spec.fieldType || 'text',
  labelRow: coords.labelRow ?? null,
  labelCol: coords.labelCol ?? null,
  valueCol: coords.valueCol ?? null,
  valueRow: coords.valueRow ?? null,
});

const findHeaderLabelCell = (getMergedAwareCellText, matchRe, effectiveSheetCols, maxRows = 80) => {
  const maxC = Math.max(12, effectiveSheetCols || 0);
  for (let r = 0; r < maxRows; r += 1) {
    for (let c = 0; c < maxC; c += 1) {
      const raw = String(getMergedAwareCellText(r, c) || '').trim();
      if (!raw || isNarrativeBlob(raw) || isTableLabelBlob(raw)) continue;
      const norm = formQKarnatakaHeaderNorm(raw);
      if (!matchRe.test(norm) && !matchRe.test(raw)) continue;
      const read = readValueBesideLabel(getMergedAwareCellText, r, c, effectiveSheetCols);
      const value = isNarrativeBlob(read.value) ? '' : read.value;
      return {
        labelRow: r,
        labelCol: c,
        valueCol: read.valueCol,
        valueRow: read.valueRow,
        value,
      };
    }
  }
  return null;
};

export function buildFormQKarnatakaHeaderFields(getMergedAwareCellText = null, effectiveSheetCols = 20) {
  return FORM_Q_KARNATAKA_HEADER_SPECS.map((spec) => {
    let coords = {};
    if (typeof getMergedAwareCellText === 'function' && spec.match) {
      coords = findHeaderLabelCell(getMergedAwareCellText, spec.match, effectiveSheetCols) || {};
    }
    return buildTemplateField(spec, coords);
  });
}

function buildWorkbookMergedCellAccessor(workbook, hints = {}) {
  if (!workbook?.SheetNames?.length) return null;
  const names = workbook.SheetNames;
  const preferred = String(hints.preferredSheetName || '').trim();
  const sheetName =
    (preferred && workbook.Sheets?.[preferred] ? preferred : null) || names[0];
  const worksheet = sheetName ? workbook.Sheets[sheetName] : null;
  if (!worksheet) return null;

  const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1:A1');
  const maxSheetCols = range.e && typeof range.e.c === 'number' ? range.e.c + 1 : 0;
  const merges = worksheet['!merges'] || [];
  let mergeMaxCol = 0;
  merges.forEach((m) => {
    if (m?.e && typeof m.e.c === 'number') mergeMaxCol = Math.max(mergeMaxCol, m.e.c + 1);
  });
  const effectiveSheetCols = Math.max(maxSheetCols, mergeMaxCol, 12);
  const getRawCellText = (r, c) => {
    if (r < 0 || c < 0) return '';
    const ref = XLSX.utils.encode_cell({ r, c });
    const cell = worksheet[ref];
    if (!cell || cell.v == null) return '';
    return String(cell.v).trim();
  };
  const getMergedAwareCellText = (r, c) => {
    const direct = getRawCellText(r, c);
    if (direct) return direct;
    for (let i = 0; i < merges.length; i += 1) {
      const m = merges[i];
      if (!m?.s || !m?.e) continue;
      if (r >= m.s.r && r <= m.e.r && c >= m.s.c && c <= m.e.c) {
        const topLeft = getRawCellText(m.s.r, m.s.c);
        if (topLeft) return topLeft;
      }
    }
    return '';
  };

  return { getMergedAwareCellText, effectiveSheetCols, sheetName };
}

export function resolveFormQKarnatakaHeaderFieldLayout(parsed, workbook, hints = {}) {
  const formHeader = hints.formHeader || parsed?.formHeader || null;
  const item = hints.item || null;
  const fileName = hints.fileName || hints.formFileName || '';
  const sheetText = hints.sheetText || '';
  if (!isFormQKarnatakaContext(formHeader, item, fileName, sheetText)) return null;

  const accessor = buildWorkbookMergedCellAccessor(workbook, hints);
  const getMergedAwareCellText = accessor?.getMergedAwareCellText ?? null;
  const effectiveSheetCols = accessor?.effectiveSheetCols ?? 20;
  const finalFields = buildFormQKarnatakaHeaderFields(getMergedAwareCellText, effectiveSheetCols);
  const tableHeaders = resolveFormQKarnatakaTableHeaders(parsed?.headers || hints.tableHeaders || null);

  return {
    formHeader: {
      title: formHeader?.title || 'FORM Q',
      subtitle: formHeader?.subtitle || '',
      reference: formHeader?.reference || '',
      formQKarnatakaHeaderFieldLayout: true,
      formQKarnatakaTableLayout: true,
      textRows: [],
      fields: finalFields,
    },
    headers: tableHeaders,
    tableData: [],
    headerRowIndex: -1,
    dataStartIndex: 0,
    tableStartCol: 0,
    sheetName: accessor?.sheetName || hints.preferredSheetName || null,
  };
}

const setHeaderField = (headerData, key, value) => {
  if (!key || value == null || String(value).trim() === '') return headerData;
  const out = { ...(headerData || {}) };
  out[key] = String(value).trim();
  return out;
};

export function applyFormQKarnatakaAutofillFromSite(headerData, siteContext = {}, options = {}) {
  const { onlyEmpty = false } = options;
  const { establishmentText = '', employerText = '', companyEmployerText = '' } = siteContext;
  let out = { ...(headerData || {}) };
  const assign = (key, value) => {
    if (onlyEmpty && String(out[key] ?? '').trim()) return;
    out = setHeaderField(out, key, value);
  };
  assign('form_q_establishment', establishmentText);
  assign('form_q_ka_employer', companyEmployerText || employerText);
  return out;
};

export function formatFormQKarnatakaEmployeeName(emp = {}) {
  const fn = String(emp.FirstName || emp['FirstName'] || emp.firstName || emp['First Name'] || '').trim();
  const ln = String(emp.LastName || emp['LastName'] || emp.lastName || emp['Last Name'] || '').trim();
  if (fn && ln) return `${fn} ${ln}`;
  return fn || ln || String(emp.Name || emp['Name'] || emp.EmployeeName || emp['Employee Name'] || '').trim();
}

export function resolveFormQKarnatakaEmployeeSerialNumber(emp = {}, rowIndex = 0) {
  const fromEmp = String(
    emp.EmployeeID ||
      emp['EmployeeID'] ||
      emp['Employee ID'] ||
      emp.EmployeeCode ||
      emp['Employee Code'] ||
      ''
  ).trim();
  if (fromEmp) return fromEmp;
  return String(rowIndex + 1);
}

const pickEmployeeValue = (emp, keys) => {
  for (const key of keys) {
    const v = emp?.[key];
    if (v != null && String(v).trim() !== '') return String(v).trim();
  }
  return '';
};

const pickPayrollValue = (flat, payrollRow, keys, patterns = []) => {
  const fromFlat = readPayrollScalar(flat, keys, patterns);
  if (fromFlat !== '') return fromFlat;
  if (payrollRow && payrollRow !== flat) {
    return readPayrollScalar(payrollRow, keys, patterns);
  }
  return '';
};

function parseFormQKarnatakaMoney(value) {
  if (value === '' || value == null) return NaN;
  const n = Number(String(value).replace(/,/g, '').trim());
  return Number.isFinite(n) ? n : NaN;
}

/** Other allowances if any = gross_pay − basic − hra */
export function computeFormQKarnatakaOtherAllowances(grossPay, basic, hra) {
  const g = parseFormQKarnatakaMoney(grossPay);
  if (!Number.isFinite(g)) return '';
  const b = parseFormQKarnatakaMoney(basic);
  const h = parseFormQKarnatakaMoney(hra);
  const known = (Number.isFinite(b) ? b : 0) + (Number.isFinite(h) ? h : 0);
  const other = Math.round((g - known) * 100) / 100;
  return Number.isFinite(other) ? String(other) : '';
}

const formQPayrollAmountToString = (value) => {
  if (value === '' || value == null) return '';
  const n = parseFormQKarnatakaMoney(value);
  if (Number.isFinite(n)) return String(n);
  const s = String(value).trim();
  return s || '';
};

/** Prefer SamplePayroll / pay-run rows that carry Basic (or HRA) wage breakdown. */
export function filterFormQKarnatakaPayrollRowsWithBasic(rows) {
  return (Array.isArray(rows) ? rows : [])
    .filter((row) => row && typeof row === 'object' && row.fetch_error !== true)
    .map((row) => flattenPayrollEarningColumns(row))
    .filter((row) => payrollRowHasWageBreakdown(row));
}

export function preferFormQKarnatakaPayrollRowsWithBasic(resolvedRows, fallbackRows = []) {
  const fromResolved = filterFormQKarnatakaPayrollRowsWithBasic(resolvedRows);
  if (fromResolved.length > 0) return fromResolved;
  const fromFallback = filterFormQKarnatakaPayrollRowsWithBasic(fallbackRows);
  if (fromFallback.length > 0) return fromFallback;
  return Array.isArray(resolvedRows) ? resolvedRows : [];
}

export function resolveFormQKarnatakaPayrollWages(payrollRow) {
  if (!payrollRow || payrollRow.fetch_error) {
    return { basic: '', vda: '', other: '', total: '' };
  }
  const flattened = flattenPayrollEarningColumns(payrollRow);
  const merged = mergePayrollRunEmployeePayload(payrollRow);
  // Flattened wage columns win over empty SamplePayroll / list scalars.
  const source = { ...payrollRow, ...merged, ...flattened };
  const wageAmounts = readPayrollForm15WageAmounts(payrollRow);
  const basicRaw =
    readPayrollScalar(
      source,
      ['basic', 'Basic', 'earned_basic', 'Earned Basic', 'basic_pay', 'Basic Earnings', 'basic_earnings'],
      [/^basic$/, /^earned_basic$/, /^basic_pay$/, /^basic_earnings$/]
    ) ||
    wageAmounts.basic;
  const basic = formQPayrollAmountToString(basicRaw);
  const vda = formQPayrollAmountToString(
    pickPayrollValue(
      source,
      payrollRow,
      ['vda', 'VDA', 'dearness_allowance', 'Dearness Allowance'],
      [/vda|dearness/i]
    )
  );
  const hraRaw =
    wageAmounts.hra !== '' && wageAmounts.hra != null
      ? wageAmounts.hra
      : pickPayrollValue(
          source,
          payrollRow,
          ['hra', 'HRA', 'hra_fbp', 'house_rent_allowance', 'House Rent Allowance'],
          [/^hra(_fbp)?$/i, /house.*rent/i]
        );
  const hra = formQPayrollAmountToString(hraRaw);
  const grossRaw = readForm10GrossPayAmount(payrollRow);
  const gross = formQPayrollAmountToString(
    grossRaw !== ''
      ? grossRaw
      : pickPayrollValue(
          source,
          payrollRow,
          ['gross_pay', 'Gross Pay', 'grossPay', 'total_earnings', 'gross', 'Gross'],
          [/^gross_pay$/i, /^total_earnings$/i, /^gross$/i]
        )
  );
  const other = computeFormQKarnatakaOtherAllowances(gross, basic, hra);
  const netPay = readStrictPayrollNetPay(payrollRow);
  const total = formQPayrollAmountToString(
    netPay !== ''
      ? netPay
      : pickPayrollValue(
          source,
          payrollRow,
          ['net_pay', 'Net Pay', 'netPay', 'netpay', 'Netpay'],
          [/^net_pay$/i, /^netpay$/i]
        )
  );
  return { basic, vda, other, total };
}

export function applyFormQKarnatakaEmployeeToRow(row, emp, headers, helpers = {}) {
  if (!row || !emp || !Array.isArray(headers)) return row;
  const {
    sanitizeValue = (v) => String(v ?? '').trim(),
    rowIndex = 0,
    formatStatutoryDateDisplay = (v) => String(v || '').trim(),
    payrollRow = null,
    monthCandidates = null,
  } = helpers;
  const out = { ...row };
  const wages = applyFormQKarnatakaAprilDefaultWages(
    resolveFormQKarnatakaPayrollWages(payrollRow),
    emp,
    monthCandidates
  );
  const postal = pickEmployeeValue(emp, [
    'Present_Address',
    'Present Address',
    'PresentAddress',
    'Address',
  ]);
  const permanent =
    pickEmployeeValue(emp, [
      'Permanent_Address',
      'Permanent Address',
      'PermanentAddress',
    ]) || postal;
  const fatherName = pickEmployeeValue(emp, [
    'Father_s_Name',
    'Father_s Name',
    'FatherName',
    'Father Name',
    'SpouseName',
    'Spouse Name',
  ]);
  const dob = pickEmployeeValue(emp, [
    'Date_of_birth',
    'Date of Birth',
    'DateofBirth',
    'Dateofbirth',
    'DOB',
    'dob',
  ]);
  const doj = pickEmployeeValue(emp, [
    'Dateofjoining',
    'DateofJoining',
    'Date of Joining',
    'DateofEntryintoService',
    'Date of entry into service',
  ]);
  const designation = String(emp.Designation || emp['Designation'] || '').trim();
  const nature = String(
    emp.NatureOfWork || emp['Nature of work'] || designation || ''
  ).trim();

  headers.forEach((header) => {
    if (isFormQKarnatakaEmployeeNameHeader(header)) {
      out[header] = sanitizeValue(formatFormQKarnatakaEmployeeName(emp));
    } else if (isFormQKarnatakaPostalAddressHeader(header)) {
      out[header] = sanitizeValue(postal);
    } else if (isFormQKarnatakaPermanentAddressHeader(header)) {
      out[header] = sanitizeValue(permanent);
    } else if (isFormQKarnatakaFatherHusbandHeader(header)) {
      out[header] = sanitizeValue(fatherName);
    } else if (isFormQKarnatakaDateOfBirthHeader(header)) {
      out[header] = sanitizeValue(dob ? formatStatutoryDateDisplay(dob) : '');
    } else if (isFormQKarnatakaEntryDateHeader(header)) {
      out[header] = sanitizeValue(doj ? formatStatutoryDateDisplay(doj) : '');
    } else if (isFormQKarnatakaDesignationHeader(header)) {
      out[header] = sanitizeValue(designation);
    } else if (isFormQKarnatakaNatureOfWorkHeader(header)) {
      out[header] = sanitizeValue(nature);
    } else if (isFormQKarnatakaSerialNumberHeader(header)) {
      out[header] = sanitizeValue(resolveFormQKarnatakaEmployeeSerialNumber(emp, rowIndex));
    } else if (isFormQKarnatakaBasicHeader(header)) {
      out[header] = sanitizeValue(wages.basic);
    } else if (isFormQKarnatakaVdaHeader(header)) {
      out[header] = sanitizeValue(wages.vda);
    } else if (isFormQKarnatakaOtherAllowanceHeader(header)) {
      out[header] = sanitizeValue(wages.other);
    } else if (isFormQKarnatakaTotalWageHeader(header)) {
      out[header] = sanitizeValue(wages.total);
    }
  });
  return out;
}

export function mapFormQKarnatakaRowsFromEmployees(employees, headers, helpers = {}) {
  const hdrs = resolveFormQKarnatakaTableHeaders(headers);
  const list = Array.isArray(employees) ? employees : [];
  const {
    sanitizeValue = (v) => String(v ?? '').trim(),
    formatStatutoryDateDisplay = (v) => String(v || '').trim(),
    resolvePayrollRow = null,
    rowIndexOffset = 0,
    monthCandidates = null,
  } = helpers;
  return list.map((empItem, rowIndex) => {
    const emp = empItem?.Employee || empItem?.employee || empItem;
    const globalRowIndex = rowIndexOffset + rowIndex;
    const payrollRow =
      typeof resolvePayrollRow === 'function' ? resolvePayrollRow(emp, globalRowIndex) : null;
    return applyFormQKarnatakaEmployeeToRow({}, emp, hdrs, {
      sanitizeValue,
      rowIndex: globalRowIndex,
      formatStatutoryDateDisplay,
      payrollRow: payrollRow && !payrollRow.fetch_error ? payrollRow : null,
      monthCandidates,
    });
  });
}

export function enrichFormQKarnatakaPayrollRows(mappedData, employees, headers, helpers = {}) {
  const hdrs = resolveFormQKarnatakaTableHeaders(headers);
  const wageHeaders = hdrs.filter(
    (h) =>
      isFormQKarnatakaBasicHeader(h) ||
      isFormQKarnatakaVdaHeader(h) ||
      isFormQKarnatakaOtherAllowanceHeader(h) ||
      isFormQKarnatakaTotalWageHeader(h)
  );
  if (!wageHeaders.length) return 0;
  const {
    resolvePayrollRow = null,
    sanitizeValue = (v) => String(v ?? '').trim(),
    overwrite = false,
    monthCandidates = null,
  } = helpers;
  if (!Array.isArray(mappedData) || typeof resolvePayrollRow !== 'function') return 0;
  let hits = 0;
  mappedData.forEach((row, rowIndex) => {
    const empItem = employees[rowIndex];
    const emp = empItem?.Employee || empItem?.employee || empItem;
    const payrollRow = resolvePayrollRow(emp, rowIndex);
    const aprilMatch =
      isAprilPayrollMonthCandidates(monthCandidates) && resolveFormQKarnatakaAprilDefaultWages(emp);
    if ((!payrollRow || payrollRow.fetch_error) && !aprilMatch) return;
    const wages = applyFormQKarnatakaAprilDefaultWages(
      resolveFormQKarnatakaPayrollWages(payrollRow),
      emp,
      monthCandidates
    );
    const hasWage = wages.basic || wages.vda || wages.other || wages.total || aprilMatch;
    if (!hasWage) return;
    let wrote = false;
    const cellEmpty = (header) => !String(row?.[header] ?? '').trim();
    hdrs.forEach((header) => {
      if (isFormQKarnatakaBasicHeader(header)) {
        if ((aprilMatch || wages.basic) && (overwrite || aprilMatch || cellEmpty(header))) {
          row[header] = sanitizeValue(wages.basic);
          wrote = true;
        }
      } else if (isFormQKarnatakaVdaHeader(header) && wages.vda) {
        if (overwrite || cellEmpty(header)) {
          row[header] = sanitizeValue(wages.vda);
          wrote = true;
        }
      } else if (isFormQKarnatakaOtherAllowanceHeader(header) && wages.other) {
        if (overwrite || cellEmpty(header)) {
          row[header] = sanitizeValue(wages.other);
          wrote = true;
        }
      } else if (isFormQKarnatakaTotalWageHeader(header)) {
        if ((aprilMatch || wages.total) && (overwrite || aprilMatch || cellEmpty(header))) {
          row[header] = sanitizeValue(wages.total);
          wrote = true;
        }
      }
    });
    if (wrote) hits += 1;
  });
  return hits;
}

export function getFormQKarnatakaRowValueForHeader(row, header) {
  if (!row || !header) return '';
  const direct = row[header];
  if (direct != null && String(direct).trim() !== '') return String(direct).trim();
  const target = formQKarnatakaHeaderNorm(header);
  for (const [k, v] of Object.entries(row)) {
    if (formQKarnatakaHeaderNorm(k) === target && v != null && String(v).trim() !== '') {
      return String(v).trim();
    }
  }
  return '';
}

export function rowHasMeaningfulFormQKarnatakaExportData(row, headers) {
  const hdrs = resolveFormQKarnatakaTableHeaders(headers);
  return hdrs.some((header) => getFormQKarnatakaRowValueForHeader(row, header) !== '');
}

/** Form Q template keeps Basic/VDA/Other/Total inside one value cell (cols C–E). */
const FORM_Q_KARNATAKA_VALUE_COL_MAX = 5;

/** Accept only numeric wage amounts — reject values like "7 month(s)". */
export function sanitizeFormQKarnatakaWageAmount(value) {
  const s = String(value ?? '').trim();
  if (!s) return '';
  if (/month/i.test(s)) return '';
  const cleaned = s.replace(/,/g, '').replace(/^\u20b9\s*/, '');
  if (!/^-?\d+(\.\d+)?$/.test(cleaned)) return '';
  return cleaned.replace(/\.0+$/, '').replace(/\.$/, '');
}

/**
 * Always rebuild the rates box as four vertical lines (original template layout).
 * Amounts sit beside headings; never append "/ month(s)" as a Total value.
 */
export function buildFormQKarnatakaWageRatesCellText(wages = {}) {
  const basic = sanitizeFormQKarnatakaWageAmount(wages.basic);
  const vda = sanitizeFormQKarnatakaWageAmount(wages.vda);
  const other = sanitizeFormQKarnatakaWageAmount(wages.other);
  const total = sanitizeFormQKarnatakaWageAmount(wages.total);
  const line = (label, amount) => (amount ? `${label}  ${amount}` : label);
  return [
    line('1. Basic', basic),
    line('2. VDA', vda),
    line('3. Other allowances if any/', other),
    line('4. Total', total),
  ].join('\n');
}

/**
 * Insert wage amounts next to "1. Basic" / "4. Total" lines inside the rates cell.
 * Keeps values inside the bordered box instead of spilling into columns F/G.
 */
export function patchFormQKarnatakaWageRatesCellText(rawText, wages = {}) {
  const basic = sanitizeFormQKarnatakaWageAmount(wages.basic);
  const vda = sanitizeFormQKarnatakaWageAmount(wages.vda);
  const other = sanitizeFormQKarnatakaWageAmount(wages.other);
  const total = sanitizeFormQKarnatakaWageAmount(wages.total);
  if (!basic && !vda && !other && !total) {
    // Still normalize smashed / rich-text templates back to vertical lines.
    const raw = String(rawText ?? '');
    if (!raw.trim()) return raw;
    if (
      /\b1[\.\)]\s*Basic\b/i.test(raw) &&
      /\b4[\.\)]\s*Total\b/i.test(raw) &&
      !/\r?\n/.test(raw)
    ) {
      return buildFormQKarnatakaWageRatesCellText({});
    }
    return raw;
  }
  return buildFormQKarnatakaWageRatesCellText({ basic, vda, other, total });
}

function findFormQKarnatakaWageRatesListCell(worksheet, excelCellValueToString, maxScanRows, maxScanCols) {
  let best = null;
  let bestScore = 0;
  for (let r = 1; r <= maxScanRows; r += 1) {
    for (let c = 1; c <= Math.min(maxScanCols, FORM_Q_KARNATAKA_VALUE_COL_MAX); c += 1) {
      const raw = excelCellValueToString(worksheet.getCell(r, c)?.value);
      if (!raw || !String(raw).trim()) continue;
      const n = formQKarnatakaHeaderNorm(raw);
      // Prefer the value-side list cell (Basic+VDA/Total), not the left "12. Rates..." label alone.
      let score = 0;
      if (/\bbasic\b/.test(n)) score += 25;
      if (/\bvda\b/.test(n)) score += 20;
      if (/other\s+allowances?/.test(n)) score += 20;
      if (/\btotal\b/.test(n)) score += 20;
      if (/rates\s+of\s+wages/.test(n) && score < 40) score += 5;
      if (score >= 45 && score > bestScore) {
        bestScore = score;
        best = { row: r, col: c, raw: String(raw) };
      }
    }
  }
  return best;
}

const excelCellValueToStringFallback = (val) => {
  if (val == null) return '';
  if (typeof val === 'string' || typeof val === 'number' || typeof val === 'boolean') return String(val);
  if (val instanceof Date) return val.toISOString();
  if (typeof val === 'object') {
    if (Array.isArray(val.richText)) {
      // Preserve soft line-breaks between rich-text runs (Form Q rates list).
      return val.richText
        .map((rt) => String(rt?.text || ''))
        .join('')
        .replace(/\r\n/g, '\n');
    }
    if (val.text != null) return String(val.text);
    if (val.result != null) return String(val.result);
  }
  return '';
};

function writeFormQKarnatakaWageRatesIntoBox(worksheet, wages, helpers = {}) {
  const excelCellValueToString = helpers.excelCellValueToString || excelCellValueToStringFallback;
  const maxScanRows = Math.max(80, (worksheet?.rowCount || 0) + 10);
  const maxScanCols = 10;
  const cell = findFormQKarnatakaWageRatesListCell(
    worksheet,
    excelCellValueToString,
    maxScanRows,
    maxScanCols
  );
  if (!cell) return false;
  const patched = patchFormQKarnatakaWageRatesCellText(cell.raw, wages);
  if (patched === cell.raw) return false;
  const excelCell = worksheet.getCell(cell.row, cell.col);
  excelCell.value = patched;
  excelCell.alignment = {
    ...(excelCell.alignment && typeof excelCell.alignment === 'object' ? excelCell.alignment : {}),
    wrapText: true,
    vertical: 'top',
  };
  return true;
}

/** Write header + first employee row back to the Karnataka Form Q template. */
export function writeFormQKarnatakaFieldsToExcelJsWorksheet(
  worksheet,
  headerFormData,
  parsedFormHeader,
  helpers = {}
) {
  if (!worksheet || !headerFormData || typeof headerFormData !== 'object') return;
  const excelCellValueToString =
    helpers.excelCellValueToString || excelCellValueToStringFallback;
  const normalize = (txt) => formQKarnatakaHeaderNorm(txt);
  const fields = Array.isArray(parsedFormHeader?.fields) ? parsedFormHeader.fields : [];
  const tableRow = helpers.tableRow || null;
  const tableHeaders = resolveFormQKarnatakaTableHeaders(helpers.headers || []);
  const maxScanRows = Math.max(80, worksheet.rowCount + 10);
  const maxScanCols = 10;

  const writeAt = (row, col, value) => {
    const text = String(value ?? '').trim();
    if (!text || row == null || col == null) return;
    worksheet.getCell(row, col).value = text;
  };

  const writeByLabel = (label, value, opts = {}) => {
    const text = String(value ?? '').trim();
    if (!text) return;
    const labelNorm = normalize(String(label || '').replace(/:+$/, ''));
    const colLimit = Number(opts.maxValueCol) > 0 ? Number(opts.maxValueCol) : maxScanCols + 4;
    let bestRow = -1;
    let bestCol = -1;
    let bestScore = 0;
    for (let r = 1; r <= maxScanRows; r += 1) {
      for (let c = 1; c <= maxScanCols; c += 1) {
        const raw = excelCellValueToString(worksheet.getCell(r, c)?.value);
        const score = labelMatchScore(labelNorm, normalize(raw.replace(/:+$/, '')));
        if (score > bestScore) {
          bestScore = score;
          bestRow = r;
          bestCol = c;
        }
      }
    }
    if (bestRow > 0 && bestCol > 0 && bestScore >= 40) {
      const labelCellRaw = excelCellValueToString(worksheet.getCell(bestRow, bestCol)?.value);
      // Wage headings already live inside the value box (cols C–E) — append beside the label.
      if (opts.appendInLabelCell || opts.wageKind) {
        const patched = patchFormQKarnatakaWageRatesCellText(labelCellRaw, {
          basic: opts.wageKind === 'basic' ? text : '',
          vda: opts.wageKind === 'vda' ? text : '',
          other: opts.wageKind === 'other' ? text : '',
          total: opts.wageKind === 'total' ? text : '',
        });
        if (patched && patched !== labelCellRaw) {
          writeAt(bestRow, bestCol, patched);
          return;
        }
        if (opts.appendInLabelCell) {
          const base = String(labelCellRaw || '').trimEnd();
          writeAt(bestRow, bestCol, base ? `${base}  ${text}` : text);
          return;
        }
      }
      for (let nc = bestCol + 1; nc <= Math.min(bestCol + 6, colLimit); nc += 1) {
        const nt = normalize(excelCellValueToString(worksheet.getCell(bestRow, nc)?.value));
        if (!nt || nt === ':') {
          writeAt(bestRow, nc, text);
          return;
        }
      }
      const fallbackCol = Math.min(bestCol + 2, colLimit);
      if (fallbackCol > bestCol) writeAt(bestRow, fallbackCol, text);
    }
  };

  fields.forEach((field) => {
    const val = headerFormData[field.key];
    if (val == null || String(val).trim() === '') return;
    if (field.labelRow != null && field.valueCol != null) {
      writeAt((field.valueRow ?? field.labelRow) + 1, field.valueCol + 1, val);
      return;
    }
    writeByLabel(field.label, val);
  });

  if (tableRow && typeof tableRow === 'object') {
    const wageValues = {
      basic: '',
      vda: '',
      other: '',
      total: '',
    };
    tableHeaders.forEach((header) => {
      const value = getFormQKarnatakaRowValueForHeader(tableRow, header);
      if (!value) return;
      if (isFormQKarnatakaBasicHeader(header)) wageValues.basic = value;
      else if (isFormQKarnatakaVdaHeader(header)) wageValues.vda = value;
      else if (isFormQKarnatakaOtherAllowanceHeader(header)) wageValues.other = value;
      else if (isFormQKarnatakaTotalWageHeader(header)) wageValues.total = value;
    });
    wageValues.basic = sanitizeFormQKarnatakaWageAmount(wageValues.basic);
    wageValues.vda = sanitizeFormQKarnatakaWageAmount(wageValues.vda);
    wageValues.other = sanitizeFormQKarnatakaWageAmount(wageValues.other);
    wageValues.total = sanitizeFormQKarnatakaWageAmount(wageValues.total);
    const wroteWageRatesInBox = writeFormQKarnatakaWageRatesIntoBox(worksheet, wageValues, {
      excelCellValueToString,
    });

    const labelMap = [
      ['Name of the Employee', isFormQKarnatakaEmployeeNameHeader],
      ['His/Her Postal Address', isFormQKarnatakaPostalAddressHeader],
      ['His/Her Permanent Address', isFormQKarnatakaPermanentAddressHeader],
      ['Father/Husband Name', isFormQKarnatakaFatherHusbandHeader],
      ['Date of Birth', isFormQKarnatakaDateOfBirthHeader],
      ['Date of his/her entry into employment', isFormQKarnatakaEntryDateHeader],
      ['Designation', isFormQKarnatakaDesignationHeader],
      ['Nature of work entrusted to him/her', isFormQKarnatakaNatureOfWorkHeader],
      ['His/Her serial number in the Register of employment', isFormQKarnatakaSerialNumberHeader],
      ['Basic', isFormQKarnatakaBasicHeader],
      ['VDA', isFormQKarnatakaVdaHeader],
      ['Other allowances if any', isFormQKarnatakaOtherAllowanceHeader],
      ['Total', isFormQKarnatakaTotalWageHeader],
    ];
    tableHeaders.forEach((header) => {
      const value = getFormQKarnatakaRowValueForHeader(tableRow, header);
      if (!value) return;
      const isWageHeader =
        isFormQKarnatakaBasicHeader(header) ||
        isFormQKarnatakaVdaHeader(header) ||
        isFormQKarnatakaOtherAllowanceHeader(header) ||
        isFormQKarnatakaTotalWageHeader(header);
      // Already written beside Basic/Total headings inside the bordered rates cell.
      if (wroteWageRatesInBox && isWageHeader) return;
      const match = labelMap.find(([, test]) => test(header));
      let wageKind = '';
      if (isFormQKarnatakaBasicHeader(header)) wageKind = 'basic';
      else if (isFormQKarnatakaVdaHeader(header)) wageKind = 'vda';
      else if (isFormQKarnatakaOtherAllowanceHeader(header)) wageKind = 'other';
      else if (isFormQKarnatakaTotalWageHeader(header)) wageKind = 'total';
      writeByLabel(match ? match[0] : header, value, {
        // Never spill wage amounts outside the form box (cols F+).
        maxValueCol: isWageHeader ? FORM_Q_KARNATAKA_VALUE_COL_MAX : undefined,
        appendInLabelCell: isWageHeader,
        wageKind,
      });
    });
  }
}

const excelCellValueToString = excelCellValueToStringFallback;

export function overlayFormQKarnatakaUserEditsOntoRows(mappedRows, tableRows, headers) {
  const hdrs = resolveFormQKarnatakaTableHeaders(headers);
  const savedRows = Array.isArray(tableRows) ? tableRows : [];
  return (Array.isArray(mappedRows) ? mappedRows : []).map((row, index) => {
    const saved = savedRows[index];
    if (!saved || typeof saved !== 'object' || !rowHasMeaningfulFormQKarnatakaExportData(saved, hdrs)) {
      return row;
    }
    const merged = { ...row };
    hdrs.forEach((header) => {
      const value = getFormQKarnatakaRowValueForHeader(saved, header);
      if (value !== '') merged[header] = saved[header] ?? value;
    });
    if (saved.__employeeLookupName) merged.__employeeLookupName = saved.__employeeLookupName;
    return merged;
  });
}

export function resolveFormQKarnatakaExportRows(mappedData, headers, employeesOverride = null, helpers = {}) {
  const hdrs = resolveFormQKarnatakaTableHeaders(headers);
  const tableRows = Array.isArray(mappedData) ? mappedData : [];
  const employees = Array.isArray(employeesOverride) ? employeesOverride : [];
  const {
    sanitizeValue = (v) => String(v ?? '').trim(),
    formatStatutoryDateDisplay = (v) => String(v || '').trim(),
    resolvePayrollRow = null,
  } = helpers;
  if (employees.length > 0) {
    const mapHelpers = { sanitizeValue, formatStatutoryDateDisplay };
    if (typeof resolvePayrollRow === 'function') mapHelpers.resolvePayrollRow = resolvePayrollRow;
    const fromEmployees = mapFormQKarnatakaRowsFromEmployees(employees, hdrs, mapHelpers).map(
      (row, index) => {
        const emp = employees[index]?.Employee || employees[index]?.employee || employees[index];
        const name = formatFormQKarnatakaEmployeeName(emp);
        if (name) row.__employeeLookupName = name;
        return row;
      }
    );
    return overlayFormQKarnatakaUserEditsOntoRows(fromEmployees, tableRows, hdrs);
  }
  return tableRows.filter((row) => rowHasMeaningfulFormQKarnatakaExportData(row, hdrs));
}

function resolveFormQKarnatakaEmployeeDownloadBaseName(row, headers, fallbackIndex = 0) {
  const hdrs = resolveFormQKarnatakaTableHeaders(headers);
  const nameHeader = hdrs.find((h) => isFormQKarnatakaEmployeeNameHeader(h));
  const raw = String(
    row?.__employeeLookupName ??
      (nameHeader ? getFormQKarnatakaRowValueForHeader(row, nameHeader) : '')
  ).trim();
  const slug = raw
    .replace(/[^\w\s.-]/g, '')
    .replace(/\s+/g, '_')
    .replace(/_+/g, '_')
    .replace(/^_|_$/g, '')
    .slice(0, 48);
  return slug || `Employee_${fallbackIndex + 1}`;
}

function allocateUniqueFormQKarnatakaDownloadFileName(baseName, usedNames) {
  const count = usedNames.get(baseName) || 0;
  usedNames.set(baseName, count + 1);
  const suffix = count > 0 ? `_${count + 1}` : '';
  return `Form_Q_Karnataka_${baseName}${suffix}.xlsx`;
}

const FORM_Q_KA_FAST_ZIP_BATCH = 12;

const FORM_Q_KA_EMPLOYEE_FIELD_SPECS = [
  { label: 'Name of the Employee', match: isFormQKarnatakaEmployeeNameHeader, isWage: false },
  { label: 'His/Her Postal Address', match: isFormQKarnatakaPostalAddressHeader, isWage: false },
  { label: 'His/Her Permanent Address', match: isFormQKarnatakaPermanentAddressHeader, isWage: false },
  { label: 'Father/Husband Name', match: isFormQKarnatakaFatherHusbandHeader, isWage: false },
  { label: 'Date of Birth', match: isFormQKarnatakaDateOfBirthHeader, isWage: false },
  {
    label: 'Date of his/her entry into employment',
    match: isFormQKarnatakaEntryDateHeader,
    isWage: false,
  },
  { label: 'Designation', match: isFormQKarnatakaDesignationHeader, isWage: false },
  {
    label: 'Nature of work entrusted to him/her',
    match: isFormQKarnatakaNatureOfWorkHeader,
    isWage: false,
  },
  {
    label: 'His/Her serial number in the Register of employment',
    match: isFormQKarnatakaSerialNumberHeader,
    isWage: false,
  },
  { label: 'Basic', match: isFormQKarnatakaBasicHeader, isWage: true, wageKind: 'basic' },
  { label: 'VDA', match: isFormQKarnatakaVdaHeader, isWage: true, wageKind: 'vda' },
  {
    label: 'Other allowances if any',
    match: isFormQKarnatakaOtherAllowanceHeader,
    isWage: true,
    wageKind: 'other',
  },
  { label: 'Total', match: isFormQKarnatakaTotalWageHeader, isWage: true, wageKind: 'total' },
];

const formQKarnatakaSanitizeExportText = (value, { preserveNewlines = false } = {}) => {
  let s = String(value ?? '').replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '');
  if (preserveNewlines) {
    return s.replace(/\r\n/g, '\n').replace(/^\n+|\n+$/g, '');
  }
  return s.trim();
};

const formQKarnatakaEscapeXml = (value) =>
  String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');

const formQKarnatakaColToLetter = (col) => {
  let result = '';
  let n = col;
  while (n > 0) {
    const rem = (n - 1) % 26;
    result = String.fromCharCode(65 + rem) + result;
    n = Math.floor((n - 1) / 26);
  }
  return result;
};

const formQKarnatakaToCellRef = (row, col) => `${formQKarnatakaColToLetter(col)}${row}`;

export function formQKarnatakaUpsertInlineStrCell(sheetXml, cellRef, value, options = {}) {
  const { preserveStyle = true, preserveNewlines = false } = options;
  // [^>/]* must not consume the "/" of self-closing cells — otherwise
  // `<c r="C8"/>` matches through the next `</c>` and drops later rows.
  const cellRe = new RegExp(
    `<c\\s+r="${cellRef}"(?=[\\s/>])([^>/]*)(?:/>|>([\\s\\S]*?)</c>)`,
    'i'
  );
  const existing = sheetXml.match(cellRe);
  let styleAttr = '';
  if (preserveStyle && existing) {
    const styleMatch = String(existing[1] || '').match(/\ss="(\d+)"/i);
    if (styleMatch) styleAttr = ` s="${styleMatch[1]}"`;
  }
  const sanitized = formQKarnatakaSanitizeExportText(value, { preserveNewlines });
  // &#10; keeps vertical wage-rate lines inside Excel inlineStr cells.
  const text = formQKarnatakaEscapeXml(sanitized).replace(/\n/g, '&#10;');
  const cellXml = text
    ? `<c r="${cellRef}"${styleAttr} t="inlineStr"><is><t xml:space="preserve">${text}</t></is></c>`
    : `<c r="${cellRef}"${styleAttr}/>`;
  if (existing) {
    return sheetXml.replace(cellRe, () => cellXml);
  }
  const rowNum = cellRef.replace(/^[A-Z]+/i, '');
  const rowRe = new RegExp(`(<row\\s+r="${rowNum}"[^>]*>)([\\s\\S]*?)(</row>)`, 'i');
  if (!rowRe.test(sheetXml)) return sheetXml;
  return sheetXml.replace(rowRe, (_, a, b, c) => `${a}${b}${cellXml}${c}`);
}

const resolveFormQKarnatakaWorksheetEntry = (zipFiles) =>
  Object.keys(zipFiles)
    .filter((name) => /^xl\/worksheets\/sheet\d+\.xml$/i.test(name))
    .sort((a, b) => a.localeCompare(b, undefined, { numeric: true }))[0] || null;

function findFormQKarnatakaBestLabelCell(
  worksheet,
  label,
  excelCellValueToString,
  maxScanRows,
  maxScanCols
) {
  const normalize = (txt) => formQKarnatakaHeaderNorm(txt);
  const labelNorm = normalize(String(label || '').replace(/:+$/, ''));
  let bestRow = -1;
  let bestCol = -1;
  let bestScore = 0;
  for (let r = 1; r <= maxScanRows; r += 1) {
    for (let c = 1; c <= maxScanCols; c += 1) {
      const raw = excelCellValueToString(worksheet.getCell(r, c)?.value);
      const score = labelMatchScore(labelNorm, normalize(raw.replace(/:+$/, '')));
      if (score > bestScore) {
        bestScore = score;
        bestRow = r;
        bestCol = c;
      }
    }
  }
  if (bestRow > 0 && bestCol > 0 && bestScore >= 40) {
    return { row: bestRow, col: bestCol, score: bestScore };
  }
  return null;
}

function resolveFormQKarnatakaBesideLabelValueCol(
  worksheet,
  labelRow,
  labelCol,
  excelCellValueToString,
  colLimit
) {
  for (let nc = labelCol + 1; nc <= Math.min(labelCol + 6, colLimit); nc += 1) {
    const nt = formQKarnatakaHeaderNorm(
      excelCellValueToString(worksheet.getCell(labelRow, nc)?.value)
    );
    if (!nt || nt === ':') return nc;
  }
  return Math.min(labelCol + 2, colLimit);
}

function resolveFormQKarnatakaFastExportPositions(worksheet, headersToUse = []) {
  const hdrs = resolveFormQKarnatakaTableHeaders(headersToUse);
  const excelCellValueToString = excelCellValueToStringFallback;
  const maxScanRows = Math.max(80, (worksheet?.rowCount || 0) + 10);
  const maxScanCols = 10;
  const fieldPositions = [];
  const seen = new Set();

  const wageRatesCell = findFormQKarnatakaWageRatesListCell(
    worksheet,
    excelCellValueToString,
    maxScanRows,
    maxScanCols
  );
  let wageRatesPosition = null;
  if (wageRatesCell) {
    // Force canonical vertical list so ZIP patches never inherit smashed single-line text.
    const normalizedBase = buildFormQKarnatakaWageRatesCellText({});
    const excelCell = worksheet.getCell(wageRatesCell.row, wageRatesCell.col);
    excelCell.value = normalizedBase;
    excelCell.alignment = {
      ...(excelCell.alignment && typeof excelCell.alignment === 'object' ? excelCell.alignment : {}),
      wrapText: true,
      vertical: 'top',
    };
    wageRatesPosition = {
      row: wageRatesCell.row,
      col: wageRatesCell.col,
      cellRef: formQKarnatakaToCellRef(wageRatesCell.row, wageRatesCell.col),
      baseRaw: normalizedBase,
    };
  }

  FORM_Q_KA_EMPLOYEE_FIELD_SPECS.forEach((spec) => {
    if (wageRatesPosition && spec.isWage) return;
    const header = hdrs.find((h) => spec.match(h)) || spec.label;
    if (seen.has(header)) return;
    const labelCell = findFormQKarnatakaBestLabelCell(
      worksheet,
      spec.label,
      excelCellValueToString,
      maxScanRows,
      maxScanCols
    );
    if (!labelCell) return;
    if (spec.isWage) {
      const baseRaw = excelCellValueToString(
        worksheet.getCell(labelCell.row, labelCell.col)?.value
      );
      fieldPositions.push({
        header,
        row: labelCell.row,
        col: labelCell.col,
        cellRef: formQKarnatakaToCellRef(labelCell.row, labelCell.col),
        isWage: true,
        wageKind: spec.wageKind,
        baseRaw: String(baseRaw || ''),
        appendInLabelCell: true,
      });
      seen.add(header);
      return;
    }
    const colLimit = maxScanCols + 4;
    const valueCol = resolveFormQKarnatakaBesideLabelValueCol(
      worksheet,
      labelCell.row,
      labelCell.col,
      excelCellValueToString,
      colLimit
    );
    if (valueCol <= labelCell.col) return;
    fieldPositions.push({
      header,
      row: labelCell.row,
      col: valueCol,
      cellRef: formQKarnatakaToCellRef(labelCell.row, valueCol),
      isWage: false,
    });
    seen.add(header);
  });

  return { fieldPositions, wageRatesPosition };
}

function clearFormQKarnatakaPerEmployeeValueCells(worksheet, positions) {
  if (!worksheet || !positions) return;
  (positions.fieldPositions || []).forEach((pos) => {
    if (pos?.appendInLabelCell && pos.baseRaw != null) {
      worksheet.getCell(pos.row, pos.col).value = pos.baseRaw;
      return;
    }
    if (pos?.row != null && pos?.col != null) {
      worksheet.getCell(pos.row, pos.col).value = '';
    }
  });
  if (positions.wageRatesPosition?.row != null && positions.wageRatesPosition?.col != null) {
    const excelCell = worksheet.getCell(
      positions.wageRatesPosition.row,
      positions.wageRatesPosition.col
    );
    excelCell.value = positions.wageRatesPosition.baseRaw || '';
    excelCell.alignment = {
      ...(excelCell.alignment && typeof excelCell.alignment === 'object' ? excelCell.alignment : {}),
      wrapText: true,
      vertical: 'top',
    };
  }
}

function collectFormQKarnatakaExportWages(exportRow, hdrs) {
  const wages = { basic: '', vda: '', other: '', total: '' };
  (Array.isArray(hdrs) ? hdrs : []).forEach((header) => {
    const value = getFormQKarnatakaRowValueForHeader(exportRow, header);
    if (!value) return;
    if (isFormQKarnatakaBasicHeader(header)) wages.basic = value;
    else if (isFormQKarnatakaVdaHeader(header)) wages.vda = value;
    else if (isFormQKarnatakaOtherAllowanceHeader(header)) wages.other = value;
    else if (isFormQKarnatakaTotalWageHeader(header)) wages.total = value;
  });
  return {
    basic: sanitizeFormQKarnatakaWageAmount(wages.basic),
    vda: sanitizeFormQKarnatakaWageAmount(wages.vda),
    other: sanitizeFormQKarnatakaWageAmount(wages.other),
    total: sanitizeFormQKarnatakaWageAmount(wages.total),
  };
}

export function patchFormQKarnatakaFastSheetXml(baseSheetXml, exportRow, positions, headersToUse = []) {
  let sheetXml = baseSheetXml;
  const hdrs = resolveFormQKarnatakaTableHeaders(headersToUse);
  const wages = collectFormQKarnatakaExportWages(exportRow, hdrs);

  if (positions?.wageRatesPosition?.cellRef) {
    const patched = patchFormQKarnatakaWageRatesCellText(
      positions.wageRatesPosition.baseRaw || '',
      wages
    );
    sheetXml = formQKarnatakaUpsertInlineStrCell(
      sheetXml,
      positions.wageRatesPosition.cellRef,
      patched,
      { preserveStyle: true, preserveNewlines: true }
    );
  }

  (positions?.fieldPositions || []).forEach((pos) => {
    if (pos.isWage) {
      const single = {
        basic: pos.wageKind === 'basic' ? wages.basic : '',
        vda: pos.wageKind === 'vda' ? wages.vda : '',
        other: pos.wageKind === 'other' ? wages.other : '',
        total: pos.wageKind === 'total' ? wages.total : '',
      };
      const patched = patchFormQKarnatakaWageRatesCellText(pos.baseRaw || '', single);
      const value =
        patched && patched !== pos.baseRaw
          ? patched
          : (() => {
              const amount = sanitizeFormQKarnatakaWageAmount(single[pos.wageKind] || '');
              if (!amount) return pos.baseRaw || '';
              const base = String(pos.baseRaw || '').trimEnd();
              return base ? `${base}  ${amount}` : amount;
            })();
      sheetXml = formQKarnatakaUpsertInlineStrCell(sheetXml, pos.cellRef, value, {
        preserveStyle: true,
        preserveNewlines: true,
      });
      return;
    }
    const value = getFormQKarnatakaRowValueForHeader(exportRow, pos.header);
    if (value !== '') {
      sheetXml = formQKarnatakaUpsertInlineStrCell(sheetXml, pos.cellRef, value, {
        preserveStyle: true,
      });
    }
  });

  return sheetXml;
}

async function buildFormQKarnatakaFastXlsxBytes(fastTemplate, exportRow, headersToUse) {
  const { sheetEntry, baseSheetXml, staticFiles, positions } = fastTemplate;
  const sheetXml = patchFormQKarnatakaFastSheetXml(
    baseSheetXml,
    exportRow,
    positions,
    headersToUse
  );
  const entryZip = new JSZip();
  Object.entries(staticFiles).forEach(([path, data]) => {
    entryZip.file(path, data);
  });
  entryZip.file(sheetEntry, sheetXml);
  return entryZip.generateAsync({ type: 'uint8array', compression: 'STORE' });
}

async function prepareFormQKarnatakaFastZipTemplate({
  templateArrayBuffer,
  parsedFormHeader,
  headerFormData,
  sheetNameHint,
  headersToUse = [],
}) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(templateArrayBuffer);
  const sheetCandidates = Array.isArray(workbook.worksheets) ? workbook.worksheets : [];
  const preferred = String(sheetNameHint || '').trim();
  const worksheet =
    (preferred && sheetCandidates.find((ws) => String(ws?.name || '').trim() === preferred)) ||
    sheetCandidates[0] ||
    null;
  if (!worksheet) throw new Error('Template worksheet not found.');

  const hdrs = resolveFormQKarnatakaTableHeaders(headersToUse);
  // Static establishment/footer only — employee fields are patched per ZIP entry.
  writeFormQKarnatakaFieldsToExcelJsWorksheet(worksheet, headerFormData, parsedFormHeader, {
    excelCellValueToString,
  });

  const positions = resolveFormQKarnatakaFastExportPositions(worksheet, hdrs);
  clearFormQKarnatakaPerEmployeeValueCells(worksheet, positions);

  const preparedBuffer = await workbook.xlsx.writeBuffer();
  const templateZip = await JSZip.loadAsync(preparedBuffer);
  const sheetEntry = resolveFormQKarnatakaWorksheetEntry(templateZip.files);
  if (!sheetEntry) throw new Error('Template worksheet XML not found.');
  const baseSheetXml = await templateZip.file(sheetEntry).async('string');
  const staticFiles = {};
  await Promise.all(
    Object.keys(templateZip.files).map(async (path) => {
      const file = templateZip.files[path];
      if (!file || file.dir || path === sheetEntry) return;
      staticFiles[path] = await file.async('uint8array');
    })
  );
  return { sheetEntry, baseSheetXml, staticFiles, positions, hdrs };
}

async function buildFormQKarnatakaFastZipDownload({
  exportRows,
  fastTemplate,
  formFileName,
  parsedFormHeader,
  headersToUse,
}) {
  const zip = new JSZip();
  const usedNames = new Map();
  const hdrs = resolveFormQKarnatakaTableHeaders(headersToUse);
  for (let i = 0; i < exportRows.length; i += FORM_Q_KA_FAST_ZIP_BATCH) {
    const batch = exportRows.slice(i, i + FORM_Q_KA_FAST_ZIP_BATCH);
    const batchBytes = await Promise.all(
      batch.map(async (exportRow, batchIndex) => {
        const index = i + batchIndex;
        const xlsxBytes = await buildFormQKarnatakaFastXlsxBytes(
          fastTemplate,
          exportRow,
          hdrs
        );
        return { xlsxBytes, exportRow, index };
      })
    );
    batchBytes.forEach((entry) => {
      if (!entry) return;
      const { xlsxBytes, exportRow, index } = entry;
      const baseName = resolveFormQKarnatakaEmployeeDownloadBaseName(exportRow, hdrs, index);
      zip.file(allocateUniqueFormQKarnatakaDownloadFileName(baseName, usedNames), xlsxBytes);
    });
    if (i > 0) {
      await new Promise((resolve) => setTimeout(resolve, 0));
    }
  }
  const zipBase = String(formFileName || parsedFormHeader?.title || 'Form_Q_Karnataka')
    .replace(/\.xlsx?$/i, '')
    .replace(/[^a-zA-Z0-9._-]+/g, '_');
  return {
    blob: await zip.generateAsync({
      type: 'blob',
      mimeType: 'application/zip',
      compression: 'STORE',
    }),
    fileName: `${zipBase}_Employees.zip`,
  };
}

export async function buildFormQKarnatakaWorkbookWithTemplateStyles({
  templateArrayBuffer,
  parsedFormHeader,
  headerFormData,
  formFileName,
  sheetNameHint,
  mappedData = [],
  headersToUse = [],
}) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(templateArrayBuffer);
  const sheetCandidates = Array.isArray(workbook.worksheets) ? workbook.worksheets : [];
  const preferred = String(sheetNameHint || '').trim();
  const worksheet =
    (preferred && sheetCandidates.find((ws) => String(ws?.name || '').trim() === preferred)) ||
    sheetCandidates[0] ||
    null;
  if (!worksheet) throw new Error('Template worksheet not found.');

  const hdrs = resolveFormQKarnatakaTableHeaders(headersToUse || []);
  const tableRow = Array.isArray(mappedData)
    ? mappedData.find((row) => rowHasMeaningfulFormQKarnatakaExportData(row, hdrs))
    : null;
  writeFormQKarnatakaFieldsToExcelJsWorksheet(worksheet, headerFormData, parsedFormHeader, {
    excelCellValueToString,
    tableRow,
    headers: hdrs,
  });

  const out = await workbook.xlsx.writeBuffer();
  const fileName =
    formFileName ||
    parsedFormHeader?.title?.replace(/[^a-zA-Z0-9]/g, '_') ||
    'Form_Q_Karnataka.xlsx';
  return {
    blob: new Blob([out], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    }),
    fileName,
  };
}

export function triggerFormQKarnatakaZipDownload(blob, fileName) {
  if (!blob || !fileName) return;
  const downloadUrl = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = downloadUrl;
  link.download = fileName;
  link.style.display = 'none';
  link.rel = 'noopener';
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  setTimeout(() => {
    try {
      URL.revokeObjectURL(downloadUrl);
    } catch (_) {
      /* ignore */
    }
  }, 60_000);
}

export async function buildFormQKarnatakaPerEmployeeDownload({
  templateArrayBuffer,
  mappedData,
  headersToUse,
  parsedFormHeader,
  formFileName,
  headerFormData,
  sheetNameHint,
  employeesOverride = null,
  resolvePayrollRow = null,
}) {
  if (!templateArrayBuffer) {
    throw new Error(
      'Original Form Q Karnataka template could not be loaded. Open Autofill again, then Download.'
    );
  }

  const hdrs = resolveFormQKarnatakaTableHeaders(headersToUse);
  const employees = Array.isArray(employeesOverride) ? employeesOverride : [];
  const tableRows = Array.isArray(mappedData) ? mappedData : [];
  const payrollRowMemo = new WeakMap();
  const memoizedResolvePayrollRow =
    typeof resolvePayrollRow === 'function'
      ? (emp, rowIndex) => {
          if (!emp || typeof emp !== 'object') return resolvePayrollRow(emp, rowIndex);
          if (payrollRowMemo.has(emp)) return payrollRowMemo.get(emp);
          const hit = resolvePayrollRow(emp, rowIndex);
          payrollRowMemo.set(emp, hit);
          return hit;
        }
      : null;
  const exportHelpers = {
    sanitizeValue: (v) => String(v ?? '').trim(),
    resolvePayrollRow: memoizedResolvePayrollRow,
  };
  const buildRowsFromEmployees = () =>
    mapFormQKarnatakaRowsFromEmployees(employees, hdrs, exportHelpers).map((row, index) => {
      const emp = employees[index]?.Employee || employees[index]?.employee || employees[index];
      const name = formatFormQKarnatakaEmployeeName(emp);
      if (name) row.__employeeLookupName = name;
      return row;
    });
  let exportRows =
    employees.length > 0
      ? resolveFormQKarnatakaExportRows(tableRows, hdrs, employees, exportHelpers)
      : tableRows.filter((row) => rowHasMeaningfulFormQKarnatakaExportData(row, hdrs));
  if (exportRows.length === 0 && employees.length > 0) {
    exportRows = buildRowsFromEmployees();
  }
  if (employees.length > exportRows.length) {
    exportRows = buildRowsFromEmployees();
    exportRows = overlayFormQKarnatakaUserEditsOntoRows(exportRows, tableRows, hdrs);
  }

  const baseHeaderData = headerFormData && typeof headerFormData === 'object' ? { ...headerFormData } : {};
  // Copy bytes so ExcelJS/JSZip cannot detach the shared template ArrayBuffer across employees.
  const templateBytes =
    templateArrayBuffer instanceof ArrayBuffer
      ? templateArrayBuffer.slice(0)
      : templateArrayBuffer;
  const workbookArgs = {
    templateArrayBuffer: templateBytes,
    parsedFormHeader,
    formFileName,
    sheetNameHint,
    headersToUse: hdrs,
    headerFormData: baseHeaderData,
  };

  if (exportRows.length <= 1 && employees.length <= 1) {
    const rows = exportRows.length === 1 ? exportRows : [];
    return buildFormQKarnatakaWorkbookWithTemplateStyles({
      ...workbookArgs,
      mappedData: rows,
    });
  }

  try {
    const fastTemplate = await prepareFormQKarnatakaFastZipTemplate({
      templateArrayBuffer:
        templateBytes instanceof ArrayBuffer ? templateBytes.slice(0) : templateBytes,
      parsedFormHeader,
      headerFormData: baseHeaderData,
      sheetNameHint,
      headersToUse: hdrs,
    });
    const hasPatchTargets =
      (fastTemplate.positions?.fieldPositions?.length || 0) > 0 ||
      !!fastTemplate.positions?.wageRatesPosition;
    if (hasPatchTargets) {
      return buildFormQKarnatakaFastZipDownload({
        exportRows,
        fastTemplate,
        formFileName,
        parsedFormHeader,
        headersToUse: hdrs,
      });
    }
  } catch (fastZipErr) {
    console.warn('Form Q Karnataka fast ZIP export failed, using standard export:', fastZipErr);
  }

  const zip = new JSZip();
  const usedNames = new Map();
  for (let i = 0; i < exportRows.length; i += 1) {
    const { blob } = await buildFormQKarnatakaWorkbookWithTemplateStyles({
      ...workbookArgs,
      // Fresh copy per employee — avoids ArrayBuffer detach after workbook.xlsx.load.
      templateArrayBuffer:
        templateBytes instanceof ArrayBuffer ? templateBytes.slice(0) : templateBytes,
      mappedData: [exportRows[i]],
    });
    const xlsxBytes = new Uint8Array(await blob.arrayBuffer());
    const baseName = resolveFormQKarnatakaEmployeeDownloadBaseName(exportRows[i], hdrs, i);
    zip.file(allocateUniqueFormQKarnatakaDownloadFileName(baseName, usedNames), xlsxBytes);
    if (i > 0 && i % 5 === 0) {
      await new Promise((resolve) => setTimeout(resolve, 0));
    }
  }
  const zipBase = String(formFileName || parsedFormHeader?.title || 'Form_Q_Karnataka')
    .replace(/\.xlsx?$/i, '')
    .replace(/[^a-zA-Z0-9._-]+/g, '_');
  return {
    blob: await zip.generateAsync({
      type: 'blob',
      mimeType: 'application/zip',
      compression: 'STORE',
    }),
    fileName: `${zipBase}_Employees.zip`,
  };
}
