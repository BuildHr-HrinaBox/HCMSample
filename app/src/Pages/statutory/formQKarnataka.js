import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { labelMatchScore } from './form18APAccidentNotice';
import {
  mergePayrollRunEmployeePayload,
  readPayrollScalar,
  readStrictPayrollNetPay,
} from '../../utils/payrollEarnings';

/** Karnataka Form Q — establishment/employer header + employee particulars table. */

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
  if (headersIndicateFormQKarnatakaTable(tableHeaders)) {
    return [...tableHeaders];
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
  for (const key of keys) {
    const v = readPayrollScalar(flat, key) ?? readPayrollScalar(payrollRow, key);
    if (v != null && String(v).trim() !== '') return String(v).trim();
  }
  for (const pattern of patterns) {
    for (const [k, v] of Object.entries(flat || {})) {
      if (pattern.test(String(k)) && v != null && String(v).trim() !== '') return String(v).trim();
    }
  }
  return '';
};

export function resolveFormQKarnatakaPayrollWages(payrollRow) {
  if (!payrollRow || payrollRow.fetch_error) {
    return { basic: '', vda: '', other: '', total: '' };
  }
  const flat = mergePayrollRunEmployeePayload(payrollRow);
  const basic = pickPayrollValue(flat, payrollRow, ['basic', 'Basic', 'basic_pay'], [/^basic$/i]);
  const vda = pickPayrollValue(
    flat,
    payrollRow,
    ['vda', 'VDA', 'dearness_allowance', 'Dearness Allowance'],
    [/vda|dearness/i]
  );
  const other = pickPayrollValue(
    flat,
    payrollRow,
    ['other_allowance', 'Other Allowance', 'other allowance'],
    [/^other_allowance$/i]
  );
  const netPay = readStrictPayrollNetPay(payrollRow);
  const total =
    netPay !== ''
      ? String(netPay)
      : pickPayrollValue(flat, payrollRow, ['net_pay', 'Net Pay', 'netPay'], [/^net_pay$/i]);
  return { basic, vda, other, total };
}

export function applyFormQKarnatakaEmployeeToRow(row, emp, headers, helpers = {}) {
  if (!row || !emp || !Array.isArray(headers)) return row;
  const {
    sanitizeValue = (v) => String(v ?? '').trim(),
    rowIndex = 0,
    formatStatutoryDateDisplay = (v) => String(v || '').trim(),
    payrollRow = null,
  } = helpers;
  const out = { ...row };
  const wages = resolveFormQKarnatakaPayrollWages(payrollRow);
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
    'Dateofjoining',
    'DateofJoining',
    'Date of Joining',
    'DateofBirth',
    'Date of Birth',
    'DOB',
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
  } = helpers;
  return list.map((empItem, rowIndex) => {
    const emp = empItem?.Employee || empItem?.employee || empItem;
    const payrollRow =
      typeof resolvePayrollRow === 'function' ? resolvePayrollRow(emp, rowIndex) : null;
    return applyFormQKarnatakaEmployeeToRow({}, emp, hdrs, {
      sanitizeValue,
      rowIndex,
      formatStatutoryDateDisplay,
      payrollRow: payrollRow && !payrollRow.fetch_error ? payrollRow : null,
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
  } = helpers;
  if (!Array.isArray(mappedData) || typeof resolvePayrollRow !== 'function') return 0;
  let hits = 0;
  mappedData.forEach((row, rowIndex) => {
    const empItem = employees[rowIndex];
    const emp = empItem?.Employee || empItem?.employee || empItem;
    const payrollRow = resolvePayrollRow(emp, rowIndex);
    if (!payrollRow || payrollRow.fetch_error) return;
    const wages = resolveFormQKarnatakaPayrollWages(payrollRow);
    const hasWage = wages.basic || wages.vda || wages.other || wages.total;
    if (!hasWage) return;
    if (!overwrite && wageHeaders.some((h) => String(row?.[h] ?? '').trim())) return;
    hdrs.forEach((header) => {
      if (isFormQKarnatakaBasicHeader(header) && wages.basic) row[header] = sanitizeValue(wages.basic);
      if (isFormQKarnatakaVdaHeader(header) && wages.vda) row[header] = sanitizeValue(wages.vda);
      if (isFormQKarnatakaOtherAllowanceHeader(header) && wages.other) {
        row[header] = sanitizeValue(wages.other);
      }
      if (isFormQKarnatakaTotalWageHeader(header) && wages.total) {
        row[header] = sanitizeValue(wages.total);
      }
    });
    hits += 1;
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

/** Write header + first employee row back to the Karnataka Form Q template. */
export function writeFormQKarnatakaFieldsToExcelJsWorksheet(
  worksheet,
  headerFormData,
  parsedFormHeader,
  helpers = {}
) {
  if (!worksheet || !headerFormData || typeof headerFormData !== 'object') return;
  const excelCellValueToString =
    helpers.excelCellValueToString ||
    ((val) => {
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
    });
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

  const writeByLabel = (label, value) => {
    const text = String(value ?? '').trim();
    if (!text) return;
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
      for (let nc = bestCol + 1; nc <= Math.min(bestCol + 6, maxScanCols + 4); nc += 1) {
        const nt = normalize(excelCellValueToString(worksheet.getCell(bestRow, nc)?.value));
        if (!nt || nt === ':') {
          writeAt(bestRow, nc, text);
          return;
        }
      }
      writeAt(bestRow, Math.min(bestCol + 2, maxScanCols + 2), text);
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
      const match = labelMap.find(([, test]) => test(header));
      writeByLabel(match ? match[0] : header, value);
    });
  }
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
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(downloadUrl);
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
  const exportHelpers = {
    sanitizeValue: (v) => String(v ?? '').trim(),
    resolvePayrollRow: typeof resolvePayrollRow === 'function' ? resolvePayrollRow : null,
  };
  let exportRows =
    employees.length > 0
      ? resolveFormQKarnatakaExportRows(tableRows, hdrs, employees, exportHelpers)
      : tableRows.filter((row) => rowHasMeaningfulFormQKarnatakaExportData(row, hdrs));
  if (exportRows.length === 0 && employees.length > 0) {
    exportRows = mapFormQKarnatakaRowsFromEmployees(employees, hdrs, exportHelpers).map((row, index) => {
      const emp = employees[index]?.Employee || employees[index]?.employee || employees[index];
      const name = formatFormQKarnatakaEmployeeName(emp);
      if (name) row.__employeeLookupName = name;
      return row;
    });
  }
  if (employees.length > exportRows.length) {
    exportRows = mapFormQKarnatakaRowsFromEmployees(employees, hdrs, exportHelpers).map((row, index) => {
      const emp = employees[index]?.Employee || employees[index]?.employee || employees[index];
      const name = formatFormQKarnatakaEmployeeName(emp);
      if (name) row.__employeeLookupName = name;
      return row;
    });
    exportRows = overlayFormQKarnatakaUserEditsOntoRows(exportRows, tableRows, hdrs);
  }

  const baseHeaderData = headerFormData && typeof headerFormData === 'object' ? { ...headerFormData } : {};
  const workbookArgs = {
    templateArrayBuffer,
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

  const zip = new JSZip();
  const usedNames = new Map();
  for (let i = 0; i < exportRows.length; i += 1) {
    const { blob } = await buildFormQKarnatakaWorkbookWithTemplateStyles({
      ...workbookArgs,
      mappedData: [exportRows[i]],
    });
    const xlsxBytes = new Uint8Array(await blob.arrayBuffer());
    const baseName = resolveFormQKarnatakaEmployeeDownloadBaseName(exportRows[i], hdrs, i);
    zip.file(allocateUniqueFormQKarnatakaDownloadFileName(baseName, usedNames), xlsxBytes);
    if (i > 0 && i % 15 === 0) {
      await new Promise((resolve) => setTimeout(resolve, 0));
    }
  }
  const zipBase = String(formFileName || parsedFormHeader?.title || 'Form_Q_Karnataka')
    .replace(/\.xlsx?$/i, '')
    .replace(/[^a-zA-Z0-9._-]+/g, '_');
  return {
    blob: await zip.generateAsync({ type: 'blob', compression: 'STORE' }),
    fileName: `${zipBase}_Employees.zip`,
  };
}
