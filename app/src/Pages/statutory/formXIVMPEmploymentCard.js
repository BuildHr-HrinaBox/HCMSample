import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import {
  mergePayrollRunEmployeePayload,
  readPayrollScalar,
} from '../../utils/payrollEarnings';
import {
  blobIndicatesEmploymentCard,
  isFormXIVEmploymentCardContext,
  matchesFormXIVHint,
  resolveFormXIVWorkbookSheetName,
} from './formXAPRegisterOfFines';

/** MP / CLRA Form XIV — Employment Card (Rule 76): header fields + employee table. */

export const FORM_XIV_MP_CANONICAL_TABLE_HEADERS = [
  'Name of the workman',
  'Serial number in the register of workmen employed',
  'Nature of employment / designation',
  "Wage rate with particulars or unit, in case of piece of work",
  'Wage period',
  'Tenure of employment',
  'Remarks',
];

export function formXIVMPHeaderNorm(txt) {
  return String(txt || '')
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

export function isFormXIVMPHeaderFieldLayoutFormHeader(formHeader) {
  return !!formHeader?.formXIVMPHeaderFieldLayout;
}

export function isFormXIVMPEmploymentCardContext(formHeader, rowItem, fileName, sheetText = '') {
  return isFormXIVEmploymentCardContext(formHeader, rowItem, fileName, sheetText, null);
}

export function headersIndicateFormXIVMPTable(tableHeaders) {
  const joined = (Array.isArray(tableHeaders) ? tableHeaders : [])
    .map((h) => formXIVMPHeaderNorm(h))
    .join('\n');
  return (
    /name\s+of\s+the\s+workman/.test(joined) &&
    /serial\s+(?:no\.?|number)\s+in\s+the\s+register/.test(joined) &&
    (/nature\s+of\s+employ/.test(joined) || /designat/.test(joined))
  );
}

export function resolveFormXIVMPTableHeaders(tableHeaders) {
  if (headersIndicateFormXIVMPTable(tableHeaders)) {
    return [...tableHeaders];
  }
  return [...FORM_XIV_MP_CANONICAL_TABLE_HEADERS];
}

export function isFormXIVMPWorkmanNameHeader(h) {
  const s = formXIVMPHeaderNorm(h).replace(/^\d+[\.\)]\s*/, '');
  return /name\s+of\s+the\s+workman/.test(s) || (s.includes('name') && s.includes('workman'));
}

export function isFormXIVMPSerialNumberHeader(h) {
  const s = formXIVMPHeaderNorm(h).replace(/^\d+[\.\)]\s*/, '');
  return /serial\s+(?:no\.?|number)\s+in\s+the\s+register/.test(s);
}

export function isFormXIVMPNatureDesignationHeader(h) {
  const s = formXIVMPHeaderNorm(h).replace(/^\d+[\.\)]\s*/, '');
  return (
    /nature\s+of\s+employ/.test(s) ||
    (s.includes('nature') && s.includes('designat')) ||
    (s.includes('designat') && s.includes('employ'))
  );
}

export function isFormXIVMPWageRateHeader(h) {
  const s = formXIVMPHeaderNorm(h).replace(/^\d+[\.\)]\s*/, '');
  return (
    /wage\s*[''']?\s*rate/.test(s) ||
    (/wage/.test(s) && /piece/.test(s)) ||
    (/wage/.test(s) && /particular/.test(s) && !/period/.test(s))
  );
}

export function isFormXIVMPWagePeriodHeader(h) {
  const s = formXIVMPHeaderNorm(h).replace(/^\d+[\.\)]\s*/, '');
  return /wage\s+period/.test(s);
}

export function isFormXIVMPTenureHeader(h) {
  const s = formXIVMPHeaderNorm(h).replace(/^\d+[\.\)]\s*/, '');
  return /tenure\s+of\s+employ/.test(s);
}

export function isFormXIVMPRemarksHeader(h) {
  const s = formXIVMPHeaderNorm(h).replace(/^\d+[\.\)]\s*/, '');
  return s === 'remarks' || s === 'remark';
}

export function isFormXIVMPSkipAutofillHeader(h) {
  return isFormXIVMPRemarksHeader(h);
}

export const FORM_XIV_MP_FIELD_GROUPS = [
  { id: 'header', title: 'Establishment particulars' },
  { id: 'table', title: 'Workman particulars' },
];

/** Top-of-form fields only — workman rows are in the table below. */
export const FORM_XIV_MP_HEADER_SPECS = [
  {
    key: 'form_xiv_mp_contractor',
    label: 'Name and address of contractor',
    group: 'header',
    fieldType: 'textarea',
    match: /name\s+and\s+address\s+(?:of|if)\s+contractor/i,
  },
  {
    key: 'form_xiv_mp_establishment',
    label: 'Name and address of establishment in/under which contract is carried on',
    group: 'header',
    fieldType: 'textarea',
    match:
      /name\s+and\s+address\s+of\s+establishment[\s\S]{0,80}?contract\s+is\s+carried\s+on|establishment\s+in\s*\/\s*under\s+which\s+contract/i,
  },
  {
    key: 'form_xiv_mp_nature_location',
    label: 'Nature of work and location of work',
    group: 'header',
    fieldType: 'textarea',
    match: /nature\s+of\s+work\s+and\s+location\s+of\s+work|nature\s+and\s+location\s+of\s+work/i,
  },
  {
    key: 'form_xiv_mp_principal_employer',
    label: 'Name and address of principal employer',
    group: 'header',
    fieldType: 'textarea',
    match: /name\s+and\s+address\s+of\s+principal\s+employer/i,
  },
];

const isNarrativeBlob = (raw) => {
  const n = formXIVMPHeaderNorm(raw);
  if (!n) return false;
  return (
    /^form\s*xiv\b/i.test(n) ||
    /employment\s+card/i.test(n) ||
    /see\s+rule\s+76/i.test(n) ||
    /contract\s+labou?r/i.test(n) ||
    n.length > 160
  );
};

const isTableLabelBlob = (raw) => {
  const n = formXIVMPHeaderNorm(raw).replace(/^\d+[\.\)]\s*/, '');
  return (
    isFormXIVMPWorkmanNameHeader(n) ||
    isFormXIVMPSerialNumberHeader(n) ||
    isFormXIVMPNatureDesignationHeader(n) ||
    isFormXIVMPWageRateHeader(n) ||
    isFormXIVMPWagePeriodHeader(n) ||
    isFormXIVMPTenureHeader(n) ||
    isFormXIVMPRemarksHeader(n)
  );
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

const readValueBelowOrBesideLabel = (getMergedAwareCellText, labelRow, labelCol, effectiveSheetCols) => {
  const maxC = Math.max(20, effectiveSheetCols || 0);
  for (let c = labelCol + 1; c < Math.min(labelCol + 14, maxC); c += 1) {
    const v = String(getMergedAwareCellText(labelRow, c) || '').trim();
    if (v && !isNarrativeBlob(v) && !isTableLabelBlob(v)) {
      return { value: v, valueCol: c, valueRow: labelRow };
    }
  }
  for (let r = labelRow + 1; r <= Math.min(labelRow + 6, labelRow + 12); r += 1) {
    const v = String(getMergedAwareCellText(r, labelCol) || '').trim();
    if (v && !isNarrativeBlob(v) && !isTableLabelBlob(v)) {
      return { value: v, valueCol: labelCol, valueRow: r };
    }
  }
  return { value: '', valueCol: labelCol + 1, valueRow: labelRow };
};

const findHeaderLabelCell = (getMergedAwareCellText, matchRe, effectiveSheetCols, maxRows = 120) => {
  const maxC = Math.max(20, effectiveSheetCols || 0);
  for (let r = 0; r < maxRows; r += 1) {
    for (let c = 0; c < maxC; c += 1) {
      const raw = String(getMergedAwareCellText(r, c) || '').trim();
      if (!raw || isNarrativeBlob(raw) || isTableLabelBlob(raw)) continue;
      const norm = formXIVMPHeaderNorm(raw);
      if (!matchRe.test(norm) && !matchRe.test(raw)) continue;
      if (norm.length > 120) continue;
      const read = readValueBelowOrBesideLabel(getMergedAwareCellText, r, c, effectiveSheetCols);
      const value = isNarrativeBlob(read.value) || isTableLabelBlob(read.value) ? '' : read.value;
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

function buildSheetTextBlob(workbook, sheetName) {
  const ws = workbook?.Sheets?.[sheetName];
  if (!ws) return '';
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  return rows
    .slice(0, 45)
    .map((row) => (Array.isArray(row) ? row : []).map((c) => String(c || '')).join(' '))
    .join(' ');
}

function pickWorkbookSheet(workbook, hints = {}) {
  const names = Array.isArray(workbook?.SheetNames) ? workbook.SheetNames : [];
  if (!names.length) return null;
  const preferred = String(hints.preferredSheetName || '').trim();
  if (preferred && workbook.Sheets?.[preferred]) return preferred;
  const fromResolver = resolveFormXIVWorkbookSheetName(workbook, hints);
  if (fromResolver) return fromResolver;
  if (names.length === 1) return names[0];
  const blob = [hints.fileName, hints.formFileName, hints.item?.formName, hints.item?.FormName]
    .filter(Boolean)
    .join(' ')
    .toLowerCase();
  if (!matchesFormXIVHint(blob) && !blobIndicatesEmploymentCard(blob, null)) return names[0];
  let best = null;
  let bestScore = -Infinity;
  for (const name of names) {
    const sheetText = buildSheetTextBlob(workbook, name);
    const sheetBlob = `${name} ${sheetText}`.toLowerCase();
    let score = 0;
    if (matchesFormXIVHint(sheetBlob)) score += 120;
    if (blobIndicatesEmploymentCard(sheetBlob, null)) score += 90;
    if (/register\s+of\s+workmen\s+employed\s+by\s+contractor/i.test(sheetBlob) && !/employment\s+card/i.test(sheetBlob)) {
      score -= 100;
    }
    if (score > bestScore) {
      bestScore = score;
      best = name;
    }
  }
  return bestScore > 0 ? best : names[0];
}

function buildWorkbookMergedCellAccessor(workbook, hints = {}) {
  if (!workbook?.SheetNames?.length) return null;
  const sheetName = pickWorkbookSheet(workbook, hints);
  const worksheet = sheetName ? workbook.Sheets[sheetName] : null;
  if (!worksheet) return null;

  const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1:A1');
  const maxSheetCols = range.e && typeof range.e.c === 'number' ? range.e.c + 1 : 0;
  const merges = worksheet['!merges'] || [];
  let mergeMaxCol = 0;
  merges.forEach((m) => {
    if (m?.e && typeof m.e.c === 'number') mergeMaxCol = Math.max(mergeMaxCol, m.e.c + 1);
  });
  const effectiveSheetCols = Math.max(maxSheetCols, mergeMaxCol, 20);
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

export function buildFormXIVMPHeaderFields(getMergedAwareCellText = null, effectiveSheetCols = 20) {
  return FORM_XIV_MP_HEADER_SPECS.map((spec) => {
    let coords = {};
    if (typeof getMergedAwareCellText === 'function' && spec.match) {
      coords = findHeaderLabelCell(getMergedAwareCellText, spec.match, effectiveSheetCols) || {};
    }
    return buildTemplateField(spec, coords);
  });
}

export function resolveFormXIVMPHeaderFieldLayout(parsed, workbook, hints = {}) {
  const formHeader = hints.formHeader || parsed?.formHeader || null;
  const item = hints.item || null;
  const fileName = hints.fileName || hints.formFileName || '';
  const sheetText = hints.sheetText || '';
  if (!isFormXIVMPEmploymentCardContext(formHeader, item, fileName, sheetText)) return null;

  const accessor = buildWorkbookMergedCellAccessor(workbook, hints);
  const getMergedAwareCellText = accessor?.getMergedAwareCellText ?? null;
  const effectiveSheetCols = accessor?.effectiveSheetCols ?? 20;
  const finalFields = buildFormXIVMPHeaderFields(getMergedAwareCellText, effectiveSheetCols);
  const tableHeaders = resolveFormXIVMPTableHeaders(parsed?.headers || hints.tableHeaders || null);

  return {
    formHeader: {
      title: 'FORM XIV',
      subtitle: 'Employment Card',
      reference: '(See rule 76)',
      formXIVMPHeaderFieldLayout: true,
      formXIVMPTableLayout: true,
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

export function isFormXIVPlaceholderCell(text) {
  const s = String(text ?? '').trim();
  if (!s) return true;
  if (/^[\s._\-…·]+$/u.test(s)) return true;
  if (/[.…_-]{4,}$/u.test(s) && s.replace(/[.…_\-\s]+/gu, '').length < 4) return true;
  return false;
}

export function applyFormXIVMPAutofillFromSite(headerData, siteContext = {}, options = {}) {
  const { onlyEmpty = false } = options;
  const {
    contractorText = '',
    establishmentText = '',
    natureLocationText = '',
    principalEmployerText = '',
  } = siteContext;
  let out = { ...(headerData || {}) };
  const assign = (key, value) => {
    if (onlyEmpty && String(out[key] ?? '').trim()) return;
    out = setHeaderField(out, key, value);
  };
  assign('form_xiv_mp_contractor', contractorText);
  assign('form_xiv_mp_establishment', establishmentText);
  assign('form_xiv_mp_nature_location', natureLocationText);
  assign('form_xiv_mp_principal_employer', principalEmployerText);
  return out;
}

export function formatFormXIVMPWorkmanName(emp = {}) {
  const fn = String(emp.FirstName || emp['FirstName'] || emp.firstName || emp['First Name'] || '').trim();
  const ln = String(emp.LastName || emp['LastName'] || emp.lastName || emp['Last Name'] || '').trim();
  if (fn && ln) return `${fn} ${ln}`;
  return fn || ln || String(emp.Name || emp['Name'] || emp.EmployeeName || emp['Employee Name'] || '').trim();
}

export function resolveFormXIVMPEmployeeSerialNumber(emp = {}, rowIndex = 0) {
  const fromEmp = String(
    emp.EmployeeID ||
      emp['EmployeeID'] ||
      emp['Employee ID'] ||
      emp.Employee_Number ||
      emp['Employee Number'] ||
      emp.EmployeeCode ||
      emp['Employee Code'] ||
      ''
  ).trim();
  if (fromEmp) return fromEmp;
  return String(rowIndex + 1);
}

export function resolveFormXIVMPDesignation(emp = {}) {
  return String(
    emp.Designation ||
      emp['Designation'] ||
      emp['Designation.displayValue'] ||
      emp.DesignationName ||
      emp['Designation Name'] ||
      emp.JobTitle ||
      emp['Job Title'] ||
      emp.Nature_of_employment ||
      emp['Nature of employment'] ||
      ''
  ).trim();
}

export function readFormXIVMPPayrollGrossPay(payrollRow) {
  if (!payrollRow || payrollRow.fetch_error) return '';
  const flat = mergePayrollRunEmployeePayload(payrollRow);
  const gross = readPayrollScalar(
    flat,
    ['gross_pay', 'Gross Pay', 'grossPay', 'total_earnings'],
    [/^gross_pay$/, /^total_earnings$/]
  );
  if (gross !== '') return gross;
  if (payrollRow !== flat) {
    return readPayrollScalar(
      payrollRow,
      ['gross_pay', 'Gross Pay', 'grossPay', 'total_earnings'],
      [/^gross_pay$/, /^total_earnings$/]
    );
  }
  return '';
}

export function resolveFormXIVMPWageRate(emp = {}, payrollRow = null) {
  const fromPayroll = readFormXIVMPPayrollGrossPay(payrollRow);
  if (fromPayroll !== '') {
    const n = Number(String(fromPayroll).replace(/,/g, '').trim());
    if (Number.isFinite(n) && n > 0) return String(n);
    return String(fromPayroll).trim();
  }
  const fromEmp = emp.MonthlySalary || emp['Monthly Salary'] || emp.BasicSalary || emp['Basic Salary'] || '';
  if (fromEmp != null && String(fromEmp).trim() !== '') return String(fromEmp).trim();
  return '';
}

export function formatFormXIVMPWagePeriod(selectedMonthStr, item, wagePeriodLine = '') {
  if (wagePeriodLine && String(wagePeriodLine).trim()) return String(wagePeriodLine).trim();
  const monthNames = [
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
  const sel = String(selectedMonthStr || '').trim().toLowerCase();
  if (sel.length >= 2) {
    const monthHit = monthNames.find(
      (m) => m.toLowerCase().startsWith(sel) || m.toLowerCase() === sel
    );
    const dueRaw = item?.dueDate || item?.DueDate || item?.monthFilter || item?.MonthFilter || '';
    const yearFromDue = dueRaw ? new Date(dueRaw).getFullYear() : NaN;
    const year = Number.isFinite(yearFromDue) ? yearFromDue : new Date().getFullYear();
    if (monthHit) return `${monthHit} ${year}`;
  }
  return sel ? sel : '';
}

export function formatFormXIVMPTenure(emp = {}, formatDateFn = (v) => String(v || '').trim()) {
  const dojRaw =
    emp.Dateofjoining ||
    emp['Dateofjoining'] ||
    emp.Date_of_Joining ||
    emp['Date of Joining'] ||
    emp.DateOfJoining ||
    emp['DateOfJoining'] ||
    emp.DOJ ||
    emp['DOJ'] ||
    '';
  const doj = formatDateFn(dojRaw);
  if (!doj) return '';
  return `From ${doj}`;
}

export function applyFormXIVMPEmployeeToRow(row, emp, headers, helpers = {}) {
  if (!row || !emp || !Array.isArray(headers)) return row;
  const {
    sanitizeValue = (v) => String(v ?? '').trim(),
    rowIndex = 0,
    selectedMonth = '',
    wagePeriodLine = '',
    formatStatutoryDateDisplay = (v) => String(v || '').trim(),
    payrollRow = null,
    item = null,
  } = helpers;
  const out = { ...row };
  const wagePeriod = formatFormXIVMPWagePeriod(selectedMonth, item, wagePeriodLine);
  headers.forEach((header) => {
    if (isFormXIVMPSkipAutofillHeader(header)) {
      out[header] = '';
      return;
    }
    if (isFormXIVMPWorkmanNameHeader(header)) {
      out[header] = sanitizeValue(formatFormXIVMPWorkmanName(emp));
      return;
    }
    if (isFormXIVMPSerialNumberHeader(header)) {
      out[header] = sanitizeValue(resolveFormXIVMPEmployeeSerialNumber(emp, rowIndex));
      return;
    }
    if (isFormXIVMPNatureDesignationHeader(header)) {
      out[header] = sanitizeValue(resolveFormXIVMPDesignation(emp));
      return;
    }
    if (isFormXIVMPWageRateHeader(header)) {
      out[header] = sanitizeValue(resolveFormXIVMPWageRate(emp, payrollRow));
      return;
    }
    if (isFormXIVMPWagePeriodHeader(header)) {
      out[header] = sanitizeValue(wagePeriod);
      return;
    }
    if (isFormXIVMPTenureHeader(header)) {
      out[header] = sanitizeValue(formatFormXIVMPTenure(emp, formatStatutoryDateDisplay));
    }
  });
  return out;
}

/** Fill wage rate from pay-run gross_pay when rows were mapped without payroll. */
export function enrichFormXIVMPPayrollRows(mappedData, employees, headers, helpers = {}) {
  const hdrs = resolveFormXIVMPTableHeaders(headers);
  const wageRateHeader = hdrs.find(isFormXIVMPWageRateHeader);
  if (!wageRateHeader) return 0;
  const {
    resolvePayrollRow = null,
    sanitizeValue = (v) => String(v ?? '').trim(),
    overwrite = false,
  } = helpers;
  if (!Array.isArray(mappedData) || mappedData.length === 0 || typeof resolvePayrollRow !== 'function') {
    return 0;
  }
  let hits = 0;
  mappedData.forEach((row, rowIndex) => {
    if (!overwrite && String(row?.[wageRateHeader] ?? '').trim()) return;
    const empItem = employees[rowIndex];
    const emp = empItem?.Employee || empItem?.employee || empItem;
    const payrollRow = resolvePayrollRow(emp, row, rowIndex);
    if (!payrollRow || payrollRow.fetch_error) return;
    const rate = resolveFormXIVMPWageRate(emp, payrollRow);
    if (rate) {
      row[wageRateHeader] = sanitizeValue(rate);
      hits += 1;
    }
  });
  return hits;
}

/** One employment card per employee — vertical numbered workman fields (Rule 76). */
export const FORM_XIV_MP_WORKMAN_FIELD_SPECS = [
  {
    label: 'Name of the workman',
    match: /name\s+of\s+the\s+workman/i,
    rowTest: isFormXIVMPWorkmanNameHeader,
  },
  {
    label: 'Serial number in the register of workmen employed',
    match: /serial\s+(?:no\.?|number)\s+in\s+the\s+register/i,
    rowTest: isFormXIVMPSerialNumberHeader,
  },
  {
    label: 'Nature of employment / designation',
    match: /nature\s+of\s+employ|designat/i,
    rowTest: isFormXIVMPNatureDesignationHeader,
  },
  {
    label: "Wage rate with particulars or unit, in case of piece of work",
    match: /wage\s*['']?\s*rate|piece\s+of\s+work/i,
    rowTest: isFormXIVMPWageRateHeader,
  },
  {
    label: 'Wage period',
    match: /wage\s+period/i,
    rowTest: isFormXIVMPWagePeriodHeader,
  },
  {
    label: 'Tenure of employment',
    match: /tenure\s+of\s+employ/i,
    rowTest: isFormXIVMPTenureHeader,
  },
  {
    label: 'Remarks',
    match: /^remarks?$/i,
    rowTest: isFormXIVMPRemarksHeader,
  },
];

export function getFormXIVMPRowValueForHeader(row, header) {
  if (!row || !header) return '';
  const direct = row[header];
  if (direct != null && String(direct).trim() !== '') return String(direct).trim();
  const want = formXIVMPHeaderNorm(header);
  for (const [k, v] of Object.entries(row)) {
    if (formXIVMPHeaderNorm(k) === want) return String(v ?? '').trim();
  }
  return '';
}

export function buildFormXIVMPWorkmanFieldsFromRow(row, headers) {
  const hdrs = resolveFormXIVMPTableHeaders(headers);
  return FORM_XIV_MP_WORKMAN_FIELD_SPECS.map((spec, index) => {
    const headerKey = hdrs.find((h) => spec.rowTest(h)) || hdrs[index] || spec.label;
    return {
      label: spec.label,
      key: `form_xiv_mp_workman_${index + 1}`,
      headerKey,
      group: 'table',
      fieldType: 'text',
      value: getFormXIVMPRowValueForHeader(row, headerKey),
    };
  });
}

export function rowHasMeaningfulFormXIVMPExportData(row, headers) {
  if (!row || typeof row !== 'object') return false;
  const hdrs = resolveFormXIVMPTableHeaders(headers);
  return hdrs.some((header) => {
    if (isFormXIVMPRemarksHeader(header)) return false;
    return getFormXIVMPRowValueForHeader(row, header) !== '';
  });
}

export function resolveFormXIVMPEmployeeDownloadBaseName(row, headers, fallbackIndex = 0) {
  const hdrs = resolveFormXIVMPTableHeaders(headers);
  const nameHeader = hdrs.find(isFormXIVMPWorkmanNameHeader);
  const raw = nameHeader ? getFormXIVMPRowValueForHeader(row, nameHeader) : '';
  const label =
    raw.split(',')[0].trim() ||
    String(row?.__employeeLookupName ?? '').trim() ||
    `Employee_${fallbackIndex + 1}`;
  const slug = label
    .replace(/[^\w\s.-]/g, '')
    .replace(/\s+/g, '_')
    .replace(/_+/g, '_')
    .replace(/^_|_$/g, '')
    .slice(0, 48);
  return slug || `Employee_${fallbackIndex + 1}`;
}

export function allocateUniqueFormXIVMPDownloadFileName(baseName, usedNames) {
  const root = String(baseName || 'Employee').trim() || 'Employee';
  const count = usedNames.get(root) || 0;
  usedNames.set(root, count + 1);
  const suffix = count > 0 ? `_${count + 1}` : '';
  return `Form_XIV_MP_${root}${suffix}.xlsx`;
}

const formXIVMPExcelCellValueToString = (val) => {
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

const looksLikeFormXIVMPWorkmanLabelCell = (cellStr, spec, ordinal) => {
  const s = String(cellStr || '').trim();
  if (!s) return false;
  const norm = formXIVMPHeaderNorm(s).replace(/^\d+[\.\)]\s*/, '');
  if (spec.match.test(s) || spec.match.test(norm)) return true;
  if (spec.rowTest(s) || spec.rowTest(norm)) return true;
  if (ordinal && new RegExp(`^${ordinal}[\\.\\)]\\s*`, 'i').test(s) && spec.match.test(norm)) return true;
  return false;
};

const FORM_XIV_MP_DEFAULT_VALUE_COL = 4;

const resolveFormXIVMPWorkmanValueColumn = (worksheet, maxScanRow = 45) => {
  for (let r = 1; r <= maxScanRow; r += 1) {
    for (let c = 1; c <= 6; c += 1) {
      const cellStr = formXIVMPExcelCellValueToString(worksheet.getCell(r, c)?.value).trim();
      if (!cellStr || !/name\s+and\s+address\s+(?:of|if)\s+contractor/i.test(cellStr)) continue;
      for (let vc = FORM_XIV_MP_DEFAULT_VALUE_COL; vc <= 10; vc += 1) {
        const v = formXIVMPExcelCellValueToString(worksheet.getCell(r, vc)?.value).trim();
        if (v && !/^[_\s.-]+$/.test(v) && !/name\s+and\s+address/i.test(v)) return vc;
      }
      return FORM_XIV_MP_DEFAULT_VALUE_COL;
    }
  }
  return FORM_XIV_MP_DEFAULT_VALUE_COL;
};

const clearFormXIVMPWorkmanValueBand = (worksheet, fromRow = 12, toRow = 35, valueCol = FORM_XIV_MP_DEFAULT_VALUE_COL) => {
  for (let r = fromRow; r <= toRow; r += 1) {
    for (let c = Math.max(1, valueCol - 1); c <= valueCol + 4; c += 1) {
      const cellStr = formXIVMPExcelCellValueToString(worksheet.getCell(r, c)?.value).trim();
      if (!cellStr) continue;
      if (/^form\s*xiv|employment\s+card|see\s+rule\s+76/i.test(cellStr)) continue;
      if (FORM_XIV_MP_HEADER_SPECS.some((spec) => spec.match.test(cellStr))) continue;
      if (FORM_XIV_MP_WORKMAN_FIELD_SPECS.some((spec) => looksLikeFormXIVMPWorkmanLabelCell(cellStr, spec))) continue;
      worksheet.getCell(r, c).value = '';
    }
  }
};

export function writeFormXIVMPHeaderFieldsToWorksheet(worksheet, headerFormData = {}, parsedFormHeader = null) {
  if (!worksheet || !headerFormData || typeof headerFormData !== 'object') return;
  const defaultValueCol = resolveFormXIVMPWorkmanValueColumn(worksheet);
  const parsedFields = Array.isArray(parsedFormHeader?.fields) ? parsedFormHeader.fields : [];
  const writeAt = (row, col, value) => {
    const text = String(value ?? '').trim();
    if (!text || row < 1 || col < 1) return;
    worksheet.getCell(row, col).value = text;
  };
  /** Karnataka / MP templates place values in column D beside the label row (not col B). */
  const writeHeaderBesideLabel = (excelRow, value) => {
    if (excelRow < 1) return;
    writeAt(excelRow, defaultValueCol, value);
  };

  FORM_XIV_MP_HEADER_SPECS.forEach((spec) => {
    const val = headerFormData[spec.key];
    if (val == null || String(val).trim() === '') return;

    const parsedField = parsedFields.find((f) => f.key === spec.key);
    if (parsedField?.labelRow != null) {
      writeHeaderBesideLabel(parsedField.labelRow + 1, val);
      return;
    }

    for (let r = 1; r <= 40; r += 1) {
      for (let c = 1; c <= 14; c += 1) {
        const raw = formXIVMPExcelCellValueToString(worksheet.getCell(r, c)?.value).trim();
        if (!raw || (!spec.match.test(raw) && !spec.match.test(formXIVMPHeaderNorm(raw)))) continue;
        writeHeaderBesideLabel(r, val);
        return;
      }
    }
  });
}

export function resolveFormXIVMPWorkmanFieldPositions(worksheet, headers) {
  const hdrs = resolveFormXIVMPTableHeaders(headers);
  const valueCol = resolveFormXIVMPWorkmanValueColumn(worksheet);
  const positions = [];
  const filled = new Set();
  for (let r = 1; r <= 45; r += 1) {
    for (let c = 1; c <= 12; c += 1) {
      const cellStr = formXIVMPExcelCellValueToString(worksheet.getCell(r, c)?.value).trim();
      if (!cellStr) continue;
      const ordinalMatch = cellStr.match(/^(\d+)[\.\)]?\s*/);
      const ordinal = ordinalMatch ? Number(ordinalMatch[1]) : null;
      for (let si = 0; si < FORM_XIV_MP_WORKMAN_FIELD_SPECS.length; si += 1) {
        const spec = FORM_XIV_MP_WORKMAN_FIELD_SPECS[si];
        if (filled.has(si)) continue;
        if (!looksLikeFormXIVMPWorkmanLabelCell(cellStr, spec, ordinal || si + 1)) continue;
        const headerKey = hdrs.find((h) => spec.rowTest(h)) || hdrs[si];
        let targetCol = valueCol;
        if (ordinal) {
          targetCol = valueCol;
        } else {
          for (let vc = c + 1; vc <= Math.min(c + 8, 14); vc += 1) {
            const nt = formXIVMPExcelCellValueToString(worksheet.getCell(r, vc)?.value).trim();
            if (isFormXIVPlaceholderCell(nt)) {
              targetCol = Math.max(vc, valueCol);
              break;
            }
            if (!nt) {
              targetCol = vc;
              break;
            }
          }
        }
        positions.push({ headerKey, row: r, col: targetCol });
        filled.add(si);
        break;
      }
    }
  }
  return { valueCol, positions };
}

export function clearFormXIVMPWorkmanFieldPositions(worksheet, positions) {
  if (!worksheet || !Array.isArray(positions)) return;
  positions.forEach(({ row, col }) => {
    if (row >= 1 && col >= 1) worksheet.getCell(row, col).value = '';
  });
}

export function writeFormXIVMPWorkmanFieldsViaPositions(worksheet, row, positions) {
  if (!worksheet || !row || typeof row !== 'object' || !Array.isArray(positions)) return;
  positions.forEach(({ headerKey, row: targetRow, col }) => {
    const value = getFormXIVMPRowValueForHeader(row, headerKey);
    worksheet.getCell(targetRow, col).value = value || '';
  });
}

export function writeFormXIVMPWorkmanFieldsToWorksheet(worksheet, row, headers, cachedLayout = null) {
  if (!worksheet || !row || typeof row !== 'object') return;
  const layout = cachedLayout || resolveFormXIVMPWorkmanFieldPositions(worksheet, headers);
  if (!cachedLayout) {
    clearFormXIVMPWorkmanValueBand(worksheet, 12, 40, layout.valueCol);
  } else {
    clearFormXIVMPWorkmanFieldPositions(worksheet, layout.positions);
  }
  writeFormXIVMPWorkmanFieldsViaPositions(worksheet, row, layout.positions);
}

/** Preserve manual edits from the modal grid when rebuilding all employee rows for export. */
export function mapFormXIVMPRowsFromEmployees(employees, headers, helpers = {}) {
  const hdrs = resolveFormXIVMPTableHeaders(headers);
  const list = Array.isArray(employees) ? employees : [];
  const {
    sanitizeValue = (v) => String(v ?? '').trim(),
    selectedMonth = '',
    wagePeriodLine = '',
    formatStatutoryDateDisplay = (v) => String(v || '').trim(),
    item = null,
    resolvePayrollRow = null,
  } = helpers;
  return list.map((empItem, rowIndex) => {
    const emp = empItem?.Employee || empItem?.employee || empItem;
    const payrollRow =
      typeof resolvePayrollRow === 'function' ? resolvePayrollRow(emp, rowIndex) : null;
    return applyFormXIVMPEmployeeToRow({}, emp, hdrs, {
      sanitizeValue,
      rowIndex,
      selectedMonth,
      wagePeriodLine,
      formatStatutoryDateDisplay,
      payrollRow: payrollRow && !payrollRow.fetch_error ? payrollRow : null,
      item,
    });
  });
}

export function overlayFormXIVMPUserEditsOntoRows(mappedRows, tableRows, headers) {
  const hdrs = resolveFormXIVMPTableHeaders(headers);
  const savedRows = Array.isArray(tableRows) ? tableRows : [];
  return (Array.isArray(mappedRows) ? mappedRows : []).map((row, index) => {
    const saved = savedRows[index];
    if (!saved || typeof saved !== 'object' || !rowHasMeaningfulFormXIVMPExportData(saved, hdrs)) {
      return row;
    }
    const merged = { ...row };
    hdrs.forEach((header) => {
      const value = getFormXIVMPRowValueForHeader(saved, header);
      if (value !== '') merged[header] = saved[header] ?? value;
    });
    return merged;
  });
}

export async function buildFormXIVMPWorkbookWithTemplateStyles({
  templateArrayBuffer,
  mappedData,
  headersToUse,
  parsedFormHeader,
  formFileName,
  headerFormData,
}) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(templateArrayBuffer);
  const worksheet = workbook.worksheets[0];
  if (!worksheet) throw new Error('Template worksheet not found.');

  const hdrs = resolveFormXIVMPTableHeaders(headersToUse);
  const sourceRows = (Array.isArray(mappedData) ? mappedData : []).filter((row) =>
    rowHasMeaningfulFormXIVMPExportData(row, hdrs)
  );
  const employeeRow = sourceRows.length === 1 ? sourceRows[0] : sourceRows[0] || null;

  writeFormXIVMPHeaderFieldsToWorksheet(worksheet, headerFormData, parsedFormHeader);
  if (employeeRow) {
    writeFormXIVMPWorkmanFieldsToWorksheet(worksheet, employeeRow, hdrs);
  }

  const out = await workbook.xlsx.writeBuffer();
  const fileName =
    formFileName ||
    parsedFormHeader?.title?.replace(/[^a-zA-Z0-9]/g, '_') ||
    `Form_XIV_MP_${Date.now()}.xlsx`;
  const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  return { blob, fileName };
}

const formXIVMPEscapeXml = (value) =>
  String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');

const formXIVMPColToLetter = (col) => {
  let result = '';
  let n = col;
  while (n > 0) {
    const rem = (n - 1) % 26;
    result = String.fromCharCode(65 + rem) + result;
    n = Math.floor((n - 1) / 26);
  }
  return result;
};

const formXIVMPToCellRef = (row, col) => `${formXIVMPColToLetter(col)}${row}`;

const formXIVMPUpsertInlineStrCell = (sheetXml, cellRef, value) => {
  const text = formXIVMPEscapeXml(String(value ?? '').trim());
  const cellXml = text
    ? `<c r="${cellRef}" t="inlineStr"><is><t xml:space="preserve">${text}</t></is></c>`
    : `<c r="${cellRef}"/>`;
  const cellRe = new RegExp(`<c\\s+r="${cellRef}"[^>]*(?:/>|>[\\s\\S]*?</c>)`, 'i');
  if (cellRe.test(sheetXml)) {
    return sheetXml.replace(cellRe, cellXml);
  }
  const rowNum = cellRef.replace(/^[A-Z]+/i, '');
  const rowRe = new RegExp(`(<row\\s+r="${rowNum}"[^>]*>)([\\s\\S]*?)(</row>)`, 'i');
  if (!rowRe.test(sheetXml)) return sheetXml;
  return sheetXml.replace(rowRe, `$1$2${cellXml}$3`);
};

const resolveFormXIVMPWorksheetEntry = (zipFiles) =>
  Object.keys(zipFiles)
    .filter((name) => /^xl\/worksheets\/sheet\d+\.xml$/i.test(name))
    .sort((a, b) => a.localeCompare(b, undefined, { numeric: true }))[0] || null;

async function prepareFormXIVMPFastZipTemplate({
  templateArrayBuffer,
  headersToUse,
  parsedFormHeader,
  headerFormData,
}) {
  const hdrs = resolveFormXIVMPTableHeaders(headersToUse);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(templateArrayBuffer);
  const worksheet = workbook.worksheets[0];
  if (!worksheet) throw new Error('Template worksheet not found.');
  writeFormXIVMPHeaderFieldsToWorksheet(worksheet, headerFormData, parsedFormHeader);
  const workmanLayout = resolveFormXIVMPWorkmanFieldPositions(worksheet, hdrs);
  clearFormXIVMPWorkmanFieldPositions(worksheet, workmanLayout.positions);
  const preparedBuffer = await workbook.xlsx.writeBuffer();
  const templateZip = await JSZip.loadAsync(preparedBuffer);
  const sheetEntry = resolveFormXIVMPWorksheetEntry(templateZip.files);
  if (!sheetEntry) throw new Error('Template worksheet XML not found.');
  const baseSheetXml = await templateZip.file(sheetEntry).async('string');
  const positions = workmanLayout.positions.map((pos) => ({
    ...pos,
    cellRef: formXIVMPToCellRef(pos.row, pos.col),
  }));
  const staticFiles = {};
  await Promise.all(
    Object.keys(templateZip.files).map(async (path) => {
      const file = templateZip.files[path];
      if (!file || file.dir || path === sheetEntry) return;
      staticFiles[path] = await file.async('uint8array');
    })
  );
  return { sheetEntry, baseSheetXml, positions, staticFiles, hdrs };
}

export function triggerFormXIVMPZipDownload(blob, fileName) {
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

export async function buildFormXIVMPPerEmployeeDownload({
  templateArrayBuffer,
  mappedData,
  headersToUse,
  parsedFormHeader,
  formFileName,
  headerFormData,
}) {
  const hdrs = resolveFormXIVMPTableHeaders(headersToUse);
  const exportRows = (Array.isArray(mappedData) ? mappedData : []).filter((row) =>
    rowHasMeaningfulFormXIVMPExportData(row, hdrs)
  );
  const workbookArgs = {
    templateArrayBuffer,
    headersToUse: hdrs,
    parsedFormHeader,
    formFileName,
    headerFormData,
  };
  const xlsxMime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

  if (exportRows.length <= 1) {
    const rows = exportRows.length === 1 ? exportRows : [];
    return buildFormXIVMPWorkbookWithTemplateStyles({ ...workbookArgs, mappedData: rows });
  }

  const { sheetEntry, baseSheetXml, positions, staticFiles } = await prepareFormXIVMPFastZipTemplate({
    templateArrayBuffer,
    headersToUse: hdrs,
    parsedFormHeader,
    headerFormData,
  });

  const zip = new JSZip();
  const usedNames = new Map();
  for (let i = 0; i < exportRows.length; i += 1) {
    let sheetXml = baseSheetXml;
    for (let pi = 0; pi < positions.length; pi += 1) {
      const pos = positions[pi];
      sheetXml = formXIVMPUpsertInlineStrCell(
        sheetXml,
        pos.cellRef,
        getFormXIVMPRowValueForHeader(exportRows[i], pos.headerKey)
      );
    }
    const entryZip = new JSZip();
    Object.entries(staticFiles).forEach(([path, data]) => {
      entryZip.file(path, data);
    });
    entryZip.file(sheetEntry, sheetXml);
    const xlsxBytes = await entryZip.generateAsync({ type: 'uint8array', compression: 'STORE' });
    const baseName = resolveFormXIVMPEmployeeDownloadBaseName(exportRows[i], hdrs, i);
    zip.file(allocateUniqueFormXIVMPDownloadFileName(baseName, usedNames), xlsxBytes);
  }
  const zipBase = String(formFileName || parsedFormHeader?.title || 'Form_XIV_MP')
    .replace(/\.xlsx?$/i, '')
    .replace(/[^a-zA-Z0-9._-]+/g, '_');
  return {
    blob: await zip.generateAsync({ type: 'blob', compression: 'STORE' }),
    fileName: `${zipBase}_Employees.zip`,
  };
}
