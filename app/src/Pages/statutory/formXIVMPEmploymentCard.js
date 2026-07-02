import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import {
  mergePayrollRunEmployeePayload,
  readPayrollNetPayForStatutory,
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

/** Gujarat CLRA Form XIV — vertical numbered workman fields (Rule 76, Central & Gujarat Rules). */
export const FORM_XIV_GJ_CANONICAL_TABLE_HEADERS = [
  'Name of the Workman',
  'S.No. in the Register of Workmen Employed',
  'Nature of Employment /Designation',
  'Date of entry into service',
  'Wage rate (With particulars of unit in case of piece - work)',
  'Wage period',
  'Tenure of Employment',
];

export function formXIVMPHeaderNorm(txt) {
  return String(txt || '')
    .replace(/\r?\n/g, ' ')
    .replace(/['''`´]/g, '')
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

export function headersIndicateFormXIVGJTable(tableHeaders) {
  const joined = (Array.isArray(tableHeaders) ? tableHeaders : [])
    .map((h) => formXIVMPHeaderNorm(h))
    .join('\n');
  return (
    /name\s+of\s+the\s+workman/.test(joined) &&
    (/s\.?\s*no\.?\s+in\s+the\s+register/.test(joined) ||
      /serial\s+(?:no\.?|number)\s+in\s+the\s+register/.test(joined)) &&
    /entry\s+into\s+service/.test(joined)
  );
}

export function headersIndicateFormXIVMPTable(tableHeaders) {
  const joined = (Array.isArray(tableHeaders) ? tableHeaders : [])
    .map((h) => formXIVMPHeaderNorm(h))
    .join('\n');
  if (headersIndicateFormXIVGJTable(tableHeaders)) return false;
  return (
    /name\s+of\s+the\s+workman/.test(joined) &&
    (/serial\s+(?:no\.?|number)\s+in\s+the\s+register/.test(joined) ||
      /s\.?\s*no\.?\s+in\s+the\s+register/.test(joined)) &&
    (/nature\s+of\s+employ/.test(joined) || /designat/.test(joined))
  );
}

export function resolveFormXIVVariant(formHeader, rowItem, fileName, sheetText = '', tableHeaders = null) {
  const parts = [
    rowItem?.formName,
    rowItem?.FormName,
    rowItem?.description,
    rowItem?.Description,
    fileName,
    formHeader?.title,
    formHeader?.subtitle,
    formHeader?.reference,
    sheetText,
  ]
    .filter((x) => x != null && String(x).trim() !== '')
    .join(' ')
    .toLowerCase();
  if (
    /gujarat/.test(parts) ||
    /xiv[\s._-]*gj/i.test(parts) ||
    /central\s*&\s*gujarat/.test(parts) ||
    /gujarat\s+rules/.test(parts)
  ) {
    return 'gj';
  }
  if (headersIndicateFormXIVGJTable(tableHeaders)) return 'gj';
  if (/entry\s+into\s+service/.test(parts) && /employment\s+card/.test(parts)) return 'gj';
  return 'mp';
}

export function isFormXIVGJGujaratContext(formHeader, rowItem, fileName, sheetText = '', tableHeaders = null) {
  if (!isFormXIVMPEmploymentCardContext(formHeader, rowItem, fileName, sheetText)) return false;
  return resolveFormXIVVariant(formHeader, rowItem, fileName, sheetText, tableHeaders) === 'gj';
}

export function resolveFormXIVMPTableHeaders(tableHeaders, hints = {}) {
  const headers = Array.isArray(tableHeaders) ? tableHeaders : [];
  const trimmed = headers.map((h) => String(h || '').trim()).filter(Boolean);
  if (headersIndicateFormXIVGJTable(trimmed) || headersIndicateFormXIVMPTable(trimmed)) {
    return [...trimmed];
  }
  const variant = resolveFormXIVVariant(
    hints.formHeader || null,
    hints.item || null,
    hints.fileName || hints.formFileName || '',
    hints.sheetText || '',
    trimmed
  );
  return variant === 'gj' ? [...FORM_XIV_GJ_CANONICAL_TABLE_HEADERS] : [...FORM_XIV_MP_CANONICAL_TABLE_HEADERS];
}

export function isFormXIVMPWorkmanNameHeader(h) {
  const s = formXIVMPHeaderNorm(h).replace(/^\d+[\.\)]\s*/, '');
  return /name\s+of\s+the\s+workman/.test(s) || (s.includes('name') && s.includes('workman'));
}

export function isFormXIVMPSerialNumberHeader(h) {
  const s = formXIVMPHeaderNorm(h).replace(/^\d+[\.\)]\s*/, '');
  return (
    /serial\s+(?:no\.?|number)\s+in\s+the\s+register/.test(s) ||
    /s\.?\s*no\.?\s+in\s+the\s+register/.test(s) ||
    (/register\s+of\s+workmen\s+employed/.test(s) && /s\.?\s*no/.test(s))
  );
}

export function isFormXIVMPEntryDateHeader(h) {
  const s = formXIVMPHeaderNorm(h).replace(/^\d+[\.\)]\s*/, '');
  return /date\s+of\s+entry\s+into\s+service/.test(s) || /entry\s+into\s+service/.test(s);
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
    isFormXIVMPEntryDateHeader(n) ||
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

const readValueBelowOrBesideLabel = (
  getMergedAwareCellText,
  labelRow,
  labelCol,
  effectiveSheetCols,
  preferBelow = false
) => {
  const maxC = Math.max(20, effectiveSheetCols || 0);
  const readBelow = () => {
    for (let r = labelRow + 1; r <= Math.min(labelRow + 8, labelRow + 12); r += 1) {
      const v = String(getMergedAwareCellText(r, labelCol) || '').trim();
      if (!v || isFormXIVPlaceholderCell(v)) {
        return { value: '', valueCol: labelCol, valueRow: r };
      }
      if (!isNarrativeBlob(v) && !isTableLabelBlob(v) && !isFormXIVHeaderLabelBlob(v)) {
        return { value: v, valueCol: labelCol, valueRow: r };
      }
    }
    return { value: '', valueCol: labelCol, valueRow: labelRow + 1 };
  };
  const readBeside = () => {
    for (let c = labelCol + 1; c < Math.min(labelCol + 14, maxC); c += 1) {
      const v = String(getMergedAwareCellText(labelRow, c) || '').trim();
      if (v && !isNarrativeBlob(v) && !isTableLabelBlob(v) && !isFormXIVHeaderLabelBlob(v)) {
        return { value: v, valueCol: c, valueRow: labelRow };
      }
    }
    return null;
  };
  if (preferBelow) {
    const below = readBelow();
    if (below.value) return below;
    const beside = readBeside();
    if (beside) return beside;
    return below;
  }
  const beside = readBeside();
  if (beside) return beside;
  return readBelow();
};

const isFormXIVHeaderLabelBlob = (raw) => {
  const n = formXIVMPHeaderNorm(raw);
  return FORM_XIV_MP_HEADER_SPECS.some((spec) => spec.match.test(n) || spec.match.test(raw));
};

const findHeaderLabelCell = (
  getMergedAwareCellText,
  matchRe,
  effectiveSheetCols,
  maxRows = 120,
  preferBelow = false
) => {
  const maxC = Math.max(20, effectiveSheetCols || 0);
  for (let r = 0; r < maxRows; r += 1) {
    for (let c = 0; c < maxC; c += 1) {
      const raw = String(getMergedAwareCellText(r, c) || '').trim();
      if (!raw || isNarrativeBlob(raw) || isTableLabelBlob(raw)) continue;
      const norm = formXIVMPHeaderNorm(raw);
      if (!matchRe.test(norm) && !matchRe.test(raw)) continue;
      if (norm.length > 120) continue;
      const read = readValueBelowOrBesideLabel(
        getMergedAwareCellText,
        r,
        c,
        effectiveSheetCols,
        preferBelow
      );
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

export function buildFormXIVMPHeaderFields(getMergedAwareCellText = null, effectiveSheetCols = 20, preferBelow = false) {
  return FORM_XIV_MP_HEADER_SPECS.map((spec) => {
    let coords = {};
    if (typeof getMergedAwareCellText === 'function' && spec.match) {
      coords =
        findHeaderLabelCell(getMergedAwareCellText, spec.match, effectiveSheetCols, 120, preferBelow) ||
        {};
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
  const resolvedSheetText =
    sheetText || (accessor?.sheetName ? buildSheetTextBlob(workbook, accessor.sheetName) : '');
  const repaired =
    workbook && getMergedAwareCellText
      ? repairFormXIVTableHeadersFromWorkbook(workbook, {
          ...hints,
          formHeader,
          item,
          fileName,
          sheetText: resolvedSheetText,
        })
      : null;
  const tableHeaders = resolveFormXIVMPTableHeaders(
    repaired?.headers || parsed?.headers || hints.tableHeaders || null,
    {
      formHeader,
      item,
      fileName,
      sheetText: resolvedSheetText,
    }
  );
  const layoutVariant =
    repaired?.variant ||
    resolveFormXIVVariant(formHeader, item, fileName, resolvedSheetText, tableHeaders);
  const finalFields = buildFormXIVMPHeaderFields(
    getMergedAwareCellText,
    effectiveSheetCols,
    layoutVariant === 'gj'
  );

  return {
    formHeader: {
      title: 'FORM XIV',
      subtitle: 'Employment Card',
      reference: '(See rule 76)',
      formXIVMPHeaderFieldLayout: true,
      formXIVMPTableLayout: true,
      formXIVVariant: layoutVariant,
      textRows: [],
      fields: finalFields,
    },
    headers: tableHeaders,
    tableData: [],
    headerRowIndex: -1,
    dataStartIndex: 0,
    tableStartCol: 0,
    sheetName: repaired?.sheetName || accessor?.sheetName || hints.preferredSheetName || null,
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

export function readFormXIVMPPayrollNetPay(payrollRow) {
  if (!payrollRow || payrollRow.fetch_error) return '';
  const net = readPayrollNetPayForStatutory(payrollRow);
  if (net !== '') return String(net);
  const flat = mergePayrollRunEmployeePayload(payrollRow);
  const fromFlat = readPayrollScalar(
    flat,
    ['net_pay', 'Net Pay', 'netPay', 'netWages', 'monthly_salary', 'MonthlySalary'],
    [/^net_pay$/, /^netwages$/, /^monthly_salary$/]
  );
  if (fromFlat !== '') return fromFlat;
  if (payrollRow !== flat) {
    return readPayrollScalar(
      payrollRow,
      ['net_pay', 'Net Pay', 'netPay', 'netWages', 'monthly_salary', 'MonthlySalary'],
      [/^net_pay$/, /^netwages$/, /^monthly_salary$/]
    );
  }
  return '';
}

export function resolveFormXIVMPWageRate(emp = {}, payrollRow = null, options = {}) {
  const preferNetPay = options.variant === 'gj' || options.preferNetPay === true;
  if (preferNetPay && payrollRow) {
    const fromNet = readFormXIVMPPayrollNetPay(payrollRow);
    if (fromNet !== '') {
      const n = Number(String(fromNet).replace(/,/g, '').trim());
      if (Number.isFinite(n) && n > 0) return String(n);
      return String(fromNet).trim();
    }
  }
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

const FORM_XIV_MP_MONTH_NAMES = [
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
const FORM_XIV_MP_MONTH_ABBR = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

function resolveFormXIVMPFullMonthName(raw) {
  if (!raw || typeof raw !== 'string') return null;
  const s = raw.trim().toLowerCase();
  if (!s) return null;
  for (let i = 0; i < FORM_XIV_MP_MONTH_NAMES.length; i++) {
    const full = FORM_XIV_MP_MONTH_NAMES[i].toLowerCase();
    const ab = FORM_XIV_MP_MONTH_ABBR[i].toLowerCase();
    if (full === s || ab === s) return FORM_XIV_MP_MONTH_NAMES[i];
    if (full.startsWith(s) || s.startsWith(full.slice(0, 3))) return FORM_XIV_MP_MONTH_NAMES[i];
    if (s.startsWith(ab.slice(0, 3))) return FORM_XIV_MP_MONTH_NAMES[i];
  }
  return null;
}

function getFormXIVMPMonthFromDueDate(dueDateStr) {
  if (dueDateStr == null || dueDateStr === '') return null;
  if (typeof dueDateStr === 'number') return null;
  const s = String(dueDateStr).trim().toLowerCase();
  if (!s || s.includes('monthly basis') || /^\d{1,2}$/.test(s)) return null;
  for (let i = 0; i < FORM_XIV_MP_MONTH_NAMES.length; i++) {
    if (
      s.includes(FORM_XIV_MP_MONTH_NAMES[i].toLowerCase()) ||
      s.includes(FORM_XIV_MP_MONTH_ABBR[i].toLowerCase())
    ) {
      return FORM_XIV_MP_MONTH_NAMES[i];
    }
  }
  return null;
}

function extractFormXIVMPYearFromDueDate(dueDateStr) {
  if (dueDateStr == null || dueDateStr === '') return null;
  const m4 = String(dueDateStr).match(/\b(19|20)\d{2}\b/);
  if (m4) return parseInt(m4[0], 10);
  return null;
}

/** Wage period uses the Statutory month filter plus a calendar year (not template sample dates). */
export function formatFormXIVMPWagePeriod(selectedMonthStr, item, wagePeriodLine = '') {
  if (wagePeriodLine && String(wagePeriodLine).trim()) return String(wagePeriodLine).trim();

  const fullMonth =
    resolveFormXIVMPFullMonthName(String(selectedMonthStr || '').trim()) ||
    resolveFormXIVMPFullMonthName(String(item?.monthFilter || item?.MonthFilter || '').trim()) ||
    getFormXIVMPMonthFromDueDate(item?.dueDate || item?.DueDate || '') ||
    FORM_XIV_MP_MONTH_NAMES[new Date().getMonth()];

  const year =
    extractFormXIVMPYearFromDueDate(item?.dueDate || item?.DueDate || '') ||
    new Date().getFullYear();

  return `${fullMonth} ${year}`;
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

export function formatFormXIVMPEntryDate(emp = {}, formatDateFn = (v) => String(v || '').trim()) {
  const raw =
    emp.Dateofjoining ||
    emp['Dateofjoining'] ||
    emp.Date_of_Joining ||
    emp['Date of Joining'] ||
    emp.DateOfJoining ||
    emp['DateOfJoining'] ||
    emp.DateofEntryintoService ||
    emp['DateofEntryintoService'] ||
    emp['Date of entry into service'] ||
    emp.DOJ ||
    emp['DOJ'] ||
    '';
  return formatDateFn(raw);
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
  const variant =
    helpers.variant ||
    (headersIndicateFormXIVGJTable(headers) || headers.some(isFormXIVMPEntryDateHeader) ? 'gj' : 'mp');
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
    if (isFormXIVMPEntryDateHeader(header)) {
      out[header] = sanitizeValue(formatFormXIVMPEntryDate(emp, formatStatutoryDateDisplay));
      return;
    }
    if (isFormXIVMPWageRateHeader(header)) {
      out[header] = sanitizeValue(resolveFormXIVMPWageRate(emp, payrollRow, { variant }));
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
  const variant =
    helpers.variant ||
    (headersIndicateFormXIVGJTable(hdrs) || hdrs.some(isFormXIVMPEntryDateHeader) ? 'gj' : 'mp');
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
    const rate = resolveFormXIVMPWageRate(emp, payrollRow, { variant });
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
    match: /serial\s+(?:no\.?|number)\s+in\s+the\s+register|s\.?\s*no\.?\s+in\s+the\s+register/i,
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

/** Gujarat Form XIV — numbered fields 1–7 (includes date of entry into service). */
export const FORM_XIV_GJ_WORKMAN_FIELD_SPECS = [
  {
    label: 'Name of the Workman',
    match: /name\s+of\s+the\s+workman/i,
    rowTest: isFormXIVMPWorkmanNameHeader,
  },
  {
    label: 'S.No. in the Register of Workmen Employed',
    match: /s\.?\s*no\.?\s+in\s+the\s+register|serial\s+(?:no\.?|number)\s+in\s+the\s+register/i,
    rowTest: isFormXIVMPSerialNumberHeader,
  },
  {
    label: 'Nature of Employment /Designation',
    match: /nature\s+of\s+employ|designat/i,
    rowTest: isFormXIVMPNatureDesignationHeader,
  },
  {
    label: 'Date of entry into service',
    match: /date\s+of\s+entry\s+into\s+service|entry\s+into\s+service/i,
    rowTest: isFormXIVMPEntryDateHeader,
  },
  {
    label: 'Wage rate (With particulars of unit in case of piece - work)',
    match: /wage\s*['']?\s*rate|particulars?\s+of\s+unit|piece\s*-\s*work/i,
    rowTest: isFormXIVMPWageRateHeader,
  },
  {
    label: 'Wage period',
    match: /wage\s+period/i,
    rowTest: isFormXIVMPWagePeriodHeader,
  },
  {
    label: 'Tenure of Employment',
    match: /tenure\s+of\s+employ/i,
    rowTest: isFormXIVMPTenureHeader,
  },
];

export function resolveFormXIVWorkmanFieldSpecsFromHeaders(headers) {
  const hdrs = resolveFormXIVMPTableHeaders(headers);
  if (headersIndicateFormXIVGJTable(hdrs) || hdrs.some(isFormXIVMPEntryDateHeader)) {
    return FORM_XIV_GJ_WORKMAN_FIELD_SPECS;
  }
  return FORM_XIV_MP_WORKMAN_FIELD_SPECS;
}

function scanWorkmanLabelsFromSheet(getMergedAwareCellText, specs) {
  if (typeof getMergedAwareCellText !== 'function' || !Array.isArray(specs) || specs.length === 0) {
    return null;
  }
  const byOrdinal = Array(specs.length).fill(null);
  for (let r = 0; r < 80; r += 1) {
    for (let c = 0; c < 6; c += 1) {
      const raw = String(getMergedAwareCellText(r, c) || '').trim();
      if (!raw) continue;
      const ordMatch = raw.match(/^(\d+)[\.\)]\s*/);
      if (!ordMatch) continue;
      const ordinal = Number(ordMatch[1]);
      if (ordinal < 1 || ordinal > specs.length) continue;
      const si = ordinal - 1;
      if (byOrdinal[si]) continue;
      const spec = specs[si];
      if (looksLikeFormXIVMPWorkmanLabelCell(raw, spec, ordinal)) {
        byOrdinal[si] = raw;
      }
    }
  }
  const hitCount = byOrdinal.filter(Boolean).length;
  if (hitCount < 3) return null;
  return byOrdinal.map((label, i) => label || specs[i].label);
}

export function repairFormXIVTableHeadersFromWorkbook(workbook, hints = {}) {
  if (!workbook) return null;
  const accessor = buildWorkbookMergedCellAccessor(workbook, hints);
  if (!accessor?.getMergedAwareCellText) return null;
  const sheetText =
    hints.sheetText ||
    (accessor.sheetName ? buildSheetTextBlob(workbook, accessor.sheetName) : '');
  const gjSpecs = FORM_XIV_GJ_WORKMAN_FIELD_SPECS;
  const mpSpecs = FORM_XIV_MP_WORKMAN_FIELD_SPECS;
  const gjScanned = scanWorkmanLabelsFromSheet(accessor.getMergedAwareCellText, gjSpecs);
  const mpScanned = scanWorkmanLabelsFromSheet(accessor.getMergedAwareCellText, mpSpecs);
  const hintedVariant = resolveFormXIVVariant(
    hints.formHeader || null,
    hints.item || null,
    hints.fileName || hints.formFileName || '',
    sheetText,
    hints.tableHeaders || gjScanned || mpScanned
  );
  let variant = hintedVariant;
  let headers = variant === 'gj' ? gjScanned : mpScanned;
  if (!headers && gjScanned && mpScanned) {
    const gjHits = gjScanned.filter(Boolean).length;
    const mpHits = mpScanned.filter(Boolean).length;
    if (gjHits >= mpHits) {
      variant = 'gj';
      headers = gjScanned;
    } else {
      variant = 'mp';
      headers = mpScanned;
    }
  } else if (!headers) {
    headers = variant === 'gj' ? gjScanned : mpScanned;
  }
  if (!headers) return null;
  return {
    headers,
    variant,
    sheetName: accessor.sheetName,
  };
}

export function getFormXIVMPRowValueForHeader(row, header) {
  if (!row || !header) return '';
  const readValue = (v) => {
    if (v == null) return '';
    const text = String(v).trim();
    return text;
  };
  const direct = readValue(row[header]);
  if (direct !== '') return direct;
  const want = formXIVMPHeaderNorm(header);
  for (const [k, v] of Object.entries(row)) {
    if (k.startsWith('__')) continue;
    if (formXIVMPHeaderNorm(k) === want) {
      const hit = readValue(v);
      if (hit !== '') return hit;
    }
  }
  const targetSpec = [
    ...FORM_XIV_GJ_WORKMAN_FIELD_SPECS,
    ...FORM_XIV_MP_WORKMAN_FIELD_SPECS,
  ].find(
    (spec) => spec.rowTest(header) || formXIVMPHeaderNorm(spec.label) === want
  );
  if (targetSpec) {
    for (const [k, v] of Object.entries(row)) {
      if (k.startsWith('__')) continue;
      if (!targetSpec.rowTest(k)) continue;
      const hit = readValue(v);
      if (hit !== '') return hit;
    }
  }
  return '';
}

export function buildFormXIVMPWorkmanFieldsFromRow(row, headers) {
  const hdrs = resolveFormXIVMPTableHeaders(headers);
  const specs = resolveFormXIVWorkmanFieldSpecsFromHeaders(hdrs);
  return specs.map((spec, index) => {
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

const normalizeFormXIVMPLabelText = (text) =>
  String(text || '')
    .replace(/\r?\n/g, ' ')
    .replace(/['''`´]/g, "'")
    .replace(/\s+/g, ' ')
    .trim();

const looksLikeFormXIVMPWorkmanLabelCell = (cellStr, spec, ordinal) => {
  const s = normalizeFormXIVMPLabelText(cellStr);
  if (!s) return false;
  const norm = formXIVMPHeaderNorm(s).replace(/^\d+[\.\)]\s*/, '');
  if (spec.match.test(s) || spec.match.test(norm)) return true;
  if (spec.rowTest(s) || spec.rowTest(norm)) return true;
  if (ordinal && new RegExp(`^${ordinal}[\\.\\)]\\s*`, 'i').test(s) && (spec.match.test(norm) || spec.rowTest(norm))) {
    return true;
  }
  return false;
};

const FORM_XIV_MP_DEFAULT_VALUE_COL = 4;
const FORM_XIV_GJ_DEFAULT_VALUE_COL = 2;
/** MP / Karnataka stacked employment card — labels in A–D, values in column E. */
const FORM_XIV_MP_STACKED_VALUE_COL = 5;
const FORM_XIV_MP_STACKED_VALUE_COL_TO = 12;

const isFormXIVLabelContinuationCell = (cellStr, workmanSpecs = null) => {
  const s = normalizeFormXIVMPLabelText(cellStr);
  if (!s) return false;
  if (/^form\s*xiv|employment\s+card|see\s+rule\s+76|vide\s+rule\s+76/i.test(s)) return false;
  if (FORM_XIV_MP_HEADER_SPECS.some((spec) => spec.match.test(s))) return true;
  const specs = workmanSpecs || [...FORM_XIV_GJ_WORKMAN_FIELD_SPECS, ...FORM_XIV_MP_WORKMAN_FIELD_SPECS];
  return specs.some((spec) => looksLikeFormXIVMPWorkmanLabelCell(s, spec));
};

const findFormXIVValueColumnBesideLabel = (worksheet, row, labelCol, fallbackCol, workmanSpecs = null) => {
  if (!worksheet || row < 1 || labelCol < 1) return fallbackCol;
  for (let vc = labelCol + 1; vc <= 14; vc += 1) {
    const nt = formXIVMPExcelCellValueToString(worksheet.getCell(row, vc)?.value).trim();
    if (!nt) return vc;
    if (isFormXIVPlaceholderCell(nt)) return vc;
    if (isFormXIVLabelContinuationCell(nt, workmanSpecs)) continue;
    if (/^[_\s.-]+$/.test(nt)) return vc;
  }
  return Math.max(fallbackCol, labelCol + 1);
};

const detectFormXIVStackedWorkmanLayout = (worksheet, maxScanRow = 45) => {
  if (!worksheet) return false;
  let numberedLabelHits = 0;
  for (let r = 1; r <= maxScanRow; r += 1) {
    for (let c = 1; c <= 4; c += 1) {
      const cellStr = formXIVMPExcelCellValueToString(worksheet.getCell(r, c)?.value).trim();
      if (!cellStr || !/^\d+[\.\)]\s*/.test(cellStr)) continue;
      const norm = formXIVMPHeaderNorm(cellStr);
      if (
        /workman|serial|s\.?\s*no|designat|employ|wage|tenure|remarks|entry\s+into\s+service/.test(norm)
      ) {
        numberedLabelHits += 1;
        break;
      }
    }
  }
  if (numberedLabelHits < 3) return false;
  for (let r = 1; r <= maxScanRow; r += 1) {
    const labelA = formXIVMPExcelCellValueToString(worksheet.getCell(r, 1)?.value).trim();
    if (!labelA || !/^1[\.\)]\s*/.test(labelA) || !/workman/i.test(labelA)) continue;
    const colE = formXIVMPExcelCellValueToString(
      worksheet.getCell(r, FORM_XIV_MP_STACKED_VALUE_COL)?.value
    ).trim();
    if (!colE || isFormXIVPlaceholderCell(colE)) return true;
    const colB = formXIVMPExcelCellValueToString(worksheet.getCell(r, FORM_XIV_GJ_DEFAULT_VALUE_COL)?.value).trim();
    if (!colB || isFormXIVPlaceholderCell(colB)) return false;
  }
  return true;
};

const setFormXIVMPWorkmanCellValue = (worksheet, row, col, value) => {
  const text = String(value ?? '').trim();
  const cell = worksheet.getCell(row, col);
  if (!text) {
    cell.value = '';
    return;
  }
  const normalized = text.replace(/,/g, '');
  const asNum = Number(normalized);
  cell.value =
    Number.isFinite(asNum) && /^-?\d+(\.\d+)?$/.test(normalized) ? asNum : text;
  cell.alignment = {
    ...(cell.alignment || {}),
    horizontal: 'left',
    vertical: 'top',
    wrapText: String(cell.value).length > 36 || String(cell.value).includes('\n'),
    shrinkToFit: false,
  };
};

const resolveFormXIVMPWorkmanVariantFromSpecs = (workmanSpecs) =>
  workmanSpecs === FORM_XIV_GJ_WORKMAN_FIELD_SPECS ? 'gj' : 'mp';

const resolveFormXIVMPWorkmanValueColumn = (worksheet, maxScanRow = 45, workmanSpecs = FORM_XIV_MP_WORKMAN_FIELD_SPECS) => {
  const variant = resolveFormXIVMPWorkmanVariantFromSpecs(workmanSpecs);

  if (variant === 'gj') {
    for (let r = 1; r <= maxScanRow; r += 1) {
      for (let c = 1; c <= 4; c += 1) {
        const cellStr = formXIVMPExcelCellValueToString(worksheet.getCell(r, c)?.value).trim();
        if (!cellStr || !/^1[\.\)]\s*/.test(cellStr)) continue;
        if (!/workman/i.test(cellStr) && !isFormXIVMPWorkmanNameHeader(cellStr)) continue;
        return findFormXIVValueColumnBesideLabel(
          worksheet,
          r,
          c,
          FORM_XIV_GJ_DEFAULT_VALUE_COL,
          workmanSpecs
        );
      }
    }
    return FORM_XIV_GJ_DEFAULT_VALUE_COL;
  }

  if (detectFormXIVStackedWorkmanLayout(worksheet, maxScanRow)) {
    return FORM_XIV_MP_STACKED_VALUE_COL;
  }

  for (let r = 1; r <= maxScanRow; r += 1) {
    const labelCell = formXIVMPExcelCellValueToString(worksheet.getCell(r, 1)?.value).trim();
    if (!labelCell || !/^1[\.\)]\s*/.test(labelCell)) continue;
    if (!workmanSpecs[0]?.match?.test(labelCell) && !isFormXIVMPWorkmanNameHeader(labelCell)) continue;
    const colD = formXIVMPExcelCellValueToString(worksheet.getCell(r, FORM_XIV_MP_DEFAULT_VALUE_COL)?.value).trim();
    if (colD && !isFormXIVPlaceholderCell(colD)) return FORM_XIV_MP_DEFAULT_VALUE_COL;
    return findFormXIVValueColumnBesideLabel(
      worksheet,
      r,
      1,
      FORM_XIV_MP_DEFAULT_VALUE_COL,
      workmanSpecs
    );
  }

  for (let r = 1; r <= maxScanRow; r += 1) {
    for (let c = 1; c <= 6; c += 1) {
      const cellStr = formXIVMPExcelCellValueToString(worksheet.getCell(r, c)?.value).trim();
      if (!cellStr || !/name\s+and\s+address\s+(?:of|if)\s+contractor/i.test(cellStr)) continue;
      return findFormXIVValueColumnBesideLabel(
        worksheet,
        r,
        c,
        FORM_XIV_MP_DEFAULT_VALUE_COL,
        workmanSpecs
      );
    }
  }
  return FORM_XIV_MP_DEFAULT_VALUE_COL;
};

const clearFormXIVMPWorkmanValueBand = (
  worksheet,
  fromRow = 12,
  toRow = 35,
  valueCol = FORM_XIV_MP_DEFAULT_VALUE_COL,
  workmanSpecs = FORM_XIV_MP_WORKMAN_FIELD_SPECS
) => {
  for (let r = fromRow; r <= toRow; r += 1) {
    for (let c = Math.max(1, valueCol - 1); c <= valueCol + 4; c += 1) {
      const cellStr = formXIVMPExcelCellValueToString(worksheet.getCell(r, c)?.value).trim();
      if (!cellStr) continue;
      if (/^form\s*xiv|employment\s+card|see\s+rule\s+76|vide\s+rule\s+76/i.test(cellStr)) continue;
      if (FORM_XIV_MP_HEADER_SPECS.some((spec) => spec.match.test(cellStr))) continue;
      if (workmanSpecs.some((spec) => looksLikeFormXIVMPWorkmanLabelCell(cellStr, spec))) continue;
      worksheet.getCell(r, c).value = '';
    }
  }
};

const clearFormXIVMPStackedWorkmanValueBand = (
  worksheet,
  fromRow = 12,
  toRow = 45,
  workmanSpecs = FORM_XIV_MP_WORKMAN_FIELD_SPECS
) => {
  if (!worksheet) return;
  for (let r = fromRow; r <= toRow; r += 1) {
    for (let c = FORM_XIV_GJ_DEFAULT_VALUE_COL; c <= FORM_XIV_MP_STACKED_VALUE_COL_TO; c += 1) {
      if (c >= FORM_XIV_MP_STACKED_VALUE_COL) {
        worksheet.getCell(r, c).value = '';
        continue;
      }
      const cellStr = formXIVMPExcelCellValueToString(worksheet.getCell(r, c)?.value).trim();
      if (!cellStr) continue;
      if (workmanSpecs.some((spec) => looksLikeFormXIVMPWorkmanLabelCell(cellStr, spec))) continue;
      if (FORM_XIV_MP_HEADER_SPECS.some((spec) => spec.match.test(cellStr))) continue;
      if (isFormXIVLabelContinuationCell(cellStr, workmanSpecs)) continue;
      worksheet.getCell(r, c).value = '';
    }
  }
};

const detectFormXIVGJBoxedHeaderLayout = (worksheet) => {
  if (!worksheet) return false;
  let contractorAt = null;
  let establishmentAt = null;
  for (let r = 1; r <= 25; r += 1) {
    for (let c = 1; c <= 12; c += 1) {
      const raw = formXIVMPExcelCellValueToString(worksheet.getCell(r, c)?.value).trim();
      if (!raw) continue;
      if (/name\s+and\s+address\s+(?:of|if)\s+contractor/i.test(raw)) contractorAt = { r, c };
      if (/establishment[\s\S]{0,60}?contract\s+is\s+carried/i.test(raw)) establishmentAt = { r, c };
    }
  }
  if (!contractorAt) return false;
  const beside = formXIVMPExcelCellValueToString(
    worksheet.getCell(contractorAt.r, contractorAt.c + 1)?.value
  ).trim();
  const below = formXIVMPExcelCellValueToString(
    worksheet.getCell(contractorAt.r + 1, contractorAt.c)?.value
  ).trim();
  if (establishmentAt && establishmentAt.r === contractorAt.r && establishmentAt.c > contractorAt.c) {
    return true;
  }
  if (!beside || isFormXIVPlaceholderCell(beside)) return true;
  if (below && (isFormXIVPlaceholderCell(below) || !below)) return true;
  return /nature\s+of\s+work/i.test(below) || /principal\s+employer/i.test(below);
};

const findFormXIVGJHeaderBoxValuePosition = (worksheet, labelRow, labelCol) => {
  if (!worksheet || labelRow < 1 || labelCol < 1) {
    return { row: labelRow + 1, col: labelCol };
  }
  for (let r = labelRow + 1; r <= Math.min(labelRow + 8, 40); r += 1) {
    const rowProbe = formXIVMPExcelCellValueToString(worksheet.getCell(r, labelCol)?.value).trim();
    if (rowProbe && isFormXIVHeaderLabelBlob(rowProbe)) break;
    if (rowProbe && /^\d+[\.\)]\s*/.test(rowProbe)) break;
    const cellStr = formXIVMPExcelCellValueToString(worksheet.getCell(r, labelCol)?.value).trim();
    if (!cellStr || isFormXIVPlaceholderCell(cellStr) || /^[_\\s.-]+$/.test(cellStr)) {
      return { row: r, col: labelCol };
    }
  }
  return { row: labelRow + 1, col: labelCol };
};

const clearFormXIVGJHeaderRowSpill = (worksheet, labelRow, labelCol, boxRow) => {
  if (!worksheet || labelRow < 1 || labelCol < 1) return;
  const isRightHalf = labelCol >= 4;
  const spillEnd = isRightHalf ? Math.min(labelCol + 6, 14) : Math.min(labelCol + 5, 7);
  for (let c = labelCol + 1; c <= spillEnd; c += 1) {
    if (boxRow === labelRow && c === labelCol) continue;
    const cellStr = formXIVMPExcelCellValueToString(worksheet.getCell(labelRow, c)?.value).trim();
    if (!cellStr) continue;
    if (isFormXIVHeaderLabelBlob(cellStr)) continue;
    worksheet.getCell(labelRow, c).value = '';
  }
};

const writeFormXIVGJBoxedHeaderValue = (worksheet, labelRow, labelCol, value) => {
  const pos = findFormXIVGJHeaderBoxValuePosition(worksheet, labelRow, labelCol);
  clearFormXIVGJHeaderRowSpill(worksheet, labelRow, labelCol, pos.row);
  setFormXIVMPWorkmanCellValue(worksheet, pos.row, pos.col, value);
  const cell = worksheet.getCell(pos.row, pos.col);
  cell.alignment = {
    ...(cell.alignment || {}),
    horizontal: 'left',
    vertical: 'top',
    wrapText: true,
    shrinkToFit: false,
  };
};

const readFormXIVMPLabelWithContinuation = (worksheet, row, col) => {
  let text = normalizeFormXIVMPLabelText(
    formXIVMPExcelCellValueToString(worksheet.getCell(row, col)?.value)
  );
  if (!text) return '';
  const next = normalizeFormXIVMPLabelText(
    formXIVMPExcelCellValueToString(worksheet.getCell(row + 1, col)?.value)
  );
  if (
    next &&
    !/^\d+[\.\)]\s*/.test(next) &&
    !isFormXIVHeaderLabelBlob(next) &&
    (/(?:^case\s+of|piece|particular|work\)|^of\s+)/i.test(next) ||
      (text.includes('(') && !text.includes(')')))
  ) {
    text = `${text} ${next}`;
  }
  return text;
};

export function writeFormXIVMPHeaderFieldsToWorksheet(worksheet, headerFormData = {}, parsedFormHeader = null) {
  if (!worksheet || !headerFormData || typeof headerFormData !== 'object') return;
  const variant = parsedFormHeader?.formXIVVariant || 'mp';
  const workmanSpecs =
    variant === 'gj' ? FORM_XIV_GJ_WORKMAN_FIELD_SPECS : FORM_XIV_MP_WORKMAN_FIELD_SPECS;
  const gjBoxedLayout = variant === 'gj' && detectFormXIVGJBoxedHeaderLayout(worksheet);
  const stackedLayout = variant !== 'gj' && detectFormXIVStackedWorkmanLayout(worksheet);
  const defaultValueCol = stackedLayout
    ? FORM_XIV_MP_STACKED_VALUE_COL
    : resolveFormXIVMPWorkmanValueColumn(worksheet, 45, workmanSpecs);
  const parsedFields = Array.isArray(parsedFormHeader?.fields) ? parsedFormHeader.fields : [];
  const writeAt = (row, col, value, wrapInBox = false) => {
    const text = String(value ?? '').trim();
    if (!text || row < 1 || col < 1) return;
    if (wrapInBox) {
      setFormXIVMPWorkmanCellValue(worksheet, row, col, text);
      const cell = worksheet.getCell(row, col);
      cell.alignment = {
        ...(cell.alignment || {}),
        horizontal: 'left',
        vertical: 'top',
        wrapText: true,
        shrinkToFit: false,
      };
      return;
    }
    setFormXIVMPWorkmanCellValue(worksheet, row, col, text);
  };
  const writeBesideLabel = (excelRow, labelCol, value) => {
    if (excelRow < 1) return;
    if (gjBoxedLayout) {
      writeFormXIVGJBoxedHeaderValue(worksheet, excelRow, labelCol, value);
      return;
    }
    const targetCol =
      labelCol != null && labelCol >= 1
        ? findFormXIVValueColumnBesideLabel(worksheet, excelRow, labelCol, defaultValueCol, workmanSpecs)
        : defaultValueCol;
    writeAt(excelRow, targetCol, value);
  };

  FORM_XIV_MP_HEADER_SPECS.forEach((spec) => {
    const val = headerFormData[spec.key];
    if (val == null || String(val).trim() === '') return;

    const parsedField = parsedFields.find((f) => f.key === spec.key);
    if (parsedField?.labelRow != null) {
      const labelExcelRow = parsedField.labelRow + 1;
      const labelExcelCol = (parsedField.labelCol ?? 0) + 1;
      if (gjBoxedLayout) {
        writeFormXIVGJBoxedHeaderValue(worksheet, labelExcelRow, labelExcelCol, val);
        return;
      }
      const targetRow = (parsedField.valueRow ?? parsedField.labelRow) + 1;
      const targetCol =
        parsedField.valueCol != null
          ? parsedField.valueCol + 1
          : findFormXIVValueColumnBesideLabel(
              worksheet,
              targetRow,
              labelExcelCol,
              defaultValueCol,
              workmanSpecs
            );
      writeAt(targetRow, targetCol, val, gjBoxedLayout);
      return;
    }

    for (let r = 1; r <= 40; r += 1) {
      for (let c = 1; c <= 14; c += 1) {
        const raw = formXIVMPExcelCellValueToString(worksheet.getCell(r, c)?.value).trim();
        if (!raw || (!spec.match.test(raw) && !spec.match.test(formXIVMPHeaderNorm(raw)))) continue;
        writeBesideLabel(r, c, val);
        return;
      }
    }
  });
}

export function resolveFormXIVMPWorkmanFieldPositions(worksheet, headers) {
  const hdrs = resolveFormXIVMPTableHeaders(headers);
  const workmanSpecs = resolveFormXIVWorkmanFieldSpecsFromHeaders(hdrs);
  const variant = resolveFormXIVMPWorkmanVariantFromSpecs(workmanSpecs);
  const stackedLayout = variant === 'mp' && detectFormXIVStackedWorkmanLayout(worksheet);
  const valueCol = resolveFormXIVMPWorkmanValueColumn(worksheet, 45, workmanSpecs);
  const positions = [];
  const filled = new Set();

  const resolveTargetCol = (r, c) => {
    if (stackedLayout) return FORM_XIV_MP_STACKED_VALUE_COL;
    let targetCol = findFormXIVValueColumnBesideLabel(worksheet, r, c, valueCol, workmanSpecs);
    if (!stackedLayout && variant !== 'gj') {
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
    } else {
      targetCol = Math.max(targetCol, valueCol);
    }
    return targetCol;
  };

  for (let r = 1; r <= 45; r += 1) {
    for (let c = 1; c <= 12; c += 1) {
      const cellStr = readFormXIVMPLabelWithContinuation(worksheet, r, c);
      if (!cellStr) continue;
      const ordinalMatch = cellStr.match(/^(\d+)[\.\)]?\s*/);
      const ordinal = ordinalMatch ? Number(ordinalMatch[1]) : null;
      for (let si = 0; si < workmanSpecs.length; si += 1) {
        const spec = workmanSpecs[si];
        if (filled.has(si)) continue;
        if (!looksLikeFormXIVMPWorkmanLabelCell(cellStr, spec, ordinal || si + 1)) continue;
        const headerKey = hdrs.find((h) => spec.rowTest(h)) || hdrs[si];
        let targetRow = r;
        let targetCol = resolveTargetCol(r, c);
        if (variant === 'gj' && ordinal === 5 && isFormXIVMPWageRateHeader(cellStr)) {
          const nextRowCol = resolveTargetCol(r + 1, c);
          const thisRowCol = resolveTargetCol(r, c);
          const nextHasValue = formXIVMPExcelCellValueToString(worksheet.getCell(r + 1, nextRowCol)?.value).trim();
          if (!nextHasValue || isFormXIVPlaceholderCell(nextHasValue)) {
            targetRow = r + 1;
            targetCol = nextRowCol;
          } else {
            targetRow = r;
            targetCol = thisRowCol;
          }
        }
        positions.push({ headerKey, row: targetRow, col: targetCol, labelCol: c });
        filled.add(si);
        break;
      }
    }
  }
  for (let si = 0; si < workmanSpecs.length; si += 1) {
    if (filled.has(si)) continue;
    const spec = workmanSpecs[si];
    const ordinal = si + 1;
    const headerKey = hdrs.find((h) => spec.rowTest(h)) || hdrs[si] || spec.label;
    for (let r = 1; r <= 45; r += 1) {
      let matched = false;
      for (let c = 1; c <= 12; c += 1) {
        const cellStr = readFormXIVMPLabelWithContinuation(worksheet, r, c);
        if (!cellStr || !new RegExp(`^${ordinal}[\\.\\)\\s]`, 'i').test(cellStr)) continue;
        if (!looksLikeFormXIVMPWorkmanLabelCell(cellStr, spec, ordinal) && !spec.rowTest(cellStr)) continue;
        const targetCol = resolveTargetCol(r, c);
        positions.push({ headerKey, row: r, col: Math.max(targetCol, valueCol), labelCol: c });
        filled.add(si);
        matched = true;
        break;
      }
      if (matched) break;
    }
  }
  return { valueCol, positions, workmanSpecs, stackedLayout, variant };
}

export function clearFormXIVMPWorkmanFieldPositions(worksheet, positions) {
  if (!worksheet || !Array.isArray(positions)) return;
  positions.forEach(({ row, col }) => {
    if (row >= 1 && col >= 1) worksheet.getCell(row, col).value = '';
  });
}

export function writeFormXIVMPWorkmanFieldsViaPositions(worksheet, row, positions) {
  if (!worksheet || !row || typeof row !== 'object' || !Array.isArray(positions)) return;
  positions.forEach(({ headerKey, row: targetRow, col, labelCol }) => {
    const value = getFormXIVMPRowValueForHeader(row, headerKey);
    if (value === '') {
      worksheet.getCell(targetRow, col).value = '';
      return;
    }
    if (labelCol != null && labelCol >= 1 && labelCol < col) {
      for (let c = labelCol + 1; c < col; c += 1) {
        const between = formXIVMPExcelCellValueToString(worksheet.getCell(targetRow, c)?.value).trim();
        if (between && !isFormXIVPlaceholderCell(between) && !isFormXIVLabelContinuationCell(between)) {
          continue;
        }
        worksheet.getCell(targetRow, c).value = '';
      }
    }
    setFormXIVMPWorkmanCellValue(worksheet, targetRow, col, value);
  });
}

export function writeFormXIVMPWorkmanFieldsToWorksheet(worksheet, row, headers, cachedLayout = null) {
  if (!worksheet || !row || typeof row !== 'object') return;
  const layout = cachedLayout || resolveFormXIVMPWorkmanFieldPositions(worksheet, headers);
  const workmanSpecs = layout.workmanSpecs || resolveFormXIVWorkmanFieldSpecsFromHeaders(headers);
  if (!cachedLayout) {
    if (layout.stackedLayout) {
      clearFormXIVMPStackedWorkmanValueBand(worksheet, 12, 45, workmanSpecs);
    } else {
      clearFormXIVMPWorkmanValueBand(worksheet, 12, 40, layout.valueCol, workmanSpecs);
    }
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

  const hdrs = resolveFormXIVMPTableHeaders(headersToUse, {
    formHeader: parsedFormHeader,
    fileName: formFileName,
  });
  const parsedFormHeaderWithVariant = {
    ...(parsedFormHeader || {}),
    formXIVVariant:
      parsedFormHeader?.formXIVVariant ||
      resolveFormXIVVariant(parsedFormHeader, null, formFileName, '', hdrs),
  };
  const sourceRows = (Array.isArray(mappedData) ? mappedData : []).filter((row) =>
    rowHasMeaningfulFormXIVMPExportData(row, hdrs)
  );
  const employeeRow = sourceRows.length === 1 ? sourceRows[0] : sourceRows[0] || null;

  writeFormXIVMPHeaderFieldsToWorksheet(worksheet, headerFormData, parsedFormHeaderWithVariant);
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
  formFileName = '',
}) {
  const hdrs = resolveFormXIVMPTableHeaders(headersToUse, {
    formHeader: parsedFormHeader,
    fileName: formFileName,
  });
  const parsedFormHeaderWithVariant = {
    ...(parsedFormHeader || {}),
    formXIVVariant:
      parsedFormHeader?.formXIVVariant ||
      resolveFormXIVVariant(parsedFormHeader, null, formFileName, '', hdrs),
  };
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(templateArrayBuffer);
  const worksheet = workbook.worksheets[0];
  if (!worksheet) throw new Error('Template worksheet not found.');
  writeFormXIVMPHeaderFieldsToWorksheet(worksheet, headerFormData, parsedFormHeaderWithVariant);
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
    formFileName,
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
