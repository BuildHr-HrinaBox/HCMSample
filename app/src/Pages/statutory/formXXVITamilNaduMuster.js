import ExcelJS from 'exceljs';
import * as XLSX from 'xlsx';
import { writeStatutoryHeaderFieldsToExcelJsWorksheet } from '../../utils/statutorySiteCompanyHeaders';
import { ensureExcelJSDataRowsWithBorders } from '../../utils/excelTableBorders';
import { flattenPayrollEarningColumns, readPayrollScalar } from '../../utils/payrollEarnings';

/** Form XXVI (Tamil Nadu CLRA) — Muster with daily hours 1–31 under column band (10). */

export const FORM_XXVI_TN_DAY_BAND_PARENT = '10';

export const FORM_XXVI_TN_TITLE = 'FORM XXVI';
export const FORM_XXVI_TN_SUBTITLE = 'Register of Employment of Contract Labour';
export const FORM_XXVI_TN_REFERENCE = '[See Rule 75]';

/** Identity columns left of the daily-hours band (official Form XXVI cols 1–9). */
export const FORM_XXVI_TN_LEFT_HEADERS = [
  'Serial Number',
  'Name of the Workman',
  'Age and Sex',
  'Permanent Home Address',
  'Local address',
  'Designation (Nature of Work)',
  "Father's / Husband's Name",
  'Date of Entry into Service',
  'Rate of Wages'
];

/** Columns right of the daily-hours band (official Form XXVI cols 11–14). */
export const FORM_XXVI_TN_RIGHT_HEADERS = [
  'Number of Days Worked',
  'Signature or Thumb impression of the Workman',
  'Date of Termination of Employment',
  'Signature of Contractor/Representative'
];

/** Ordered header fields — 2-column grid reads left-to-right, top-to-bottom. */
export const FORM_XXVI_TN_HEADER_SPECS = [
  {
    key: 'form_xxvi_principal_employer',
    label: 'Name and Address of the Principal Employer',
    fieldType: 'textarea',
    match: /name\s+and\s+address\s+of\s+(?:the\s+)?principal\s+employer/i
  },
  {
    key: 'form_xxvi_contractor',
    label: 'Name and Address of the Contractor',
    fieldType: 'textarea',
    match: /name\s+and\s+address\s+of\s+(?:the\s+)?contractor/i
  },
  {
    key: 'form_xxvi_worksite',
    label: 'Name and Location of Worksite',
    fieldType: 'textarea',
    match:
      /name\s+and\s+location\s+of\s+(?:the\s+)?work\s*site|name\s+and\s+location\s+of\s+worksite/i
  },
  {
    key: 'form_x_month',
    label: 'Month:',
    match: /^month\s*:?\s*$/i
  },
  {
    key: 'form_x_year',
    label: 'Year:',
    match: /^year\s*:?\s*$/i
  }
];

const FORM_XXVI_TN_HEADER_KEYS = new Set(FORM_XXVI_TN_HEADER_SPECS.map((s) => s.key));

export function formXXVITamilNaduHeaderNorm(txt) {
  return String(txt || '')
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

/** Header-only fields that must never appear as autofill table columns. */
export function isFormXXVITamilNaduNonTableHeaderField(header) {
  const n = formXXVITamilNaduHeaderNorm(header).replace(/:+$/, '').trim();
  if (!n) return false;
  if (/name\s+and\s+location\s+of\s+(?:the\s+)?work\s*site/.test(n)) return true;
  if (/name\s+and\s+location\s+of\s+worksite/.test(n)) return true;
  if (/name\s+and\s+address\s+of\s+(?:the\s+)?principal\s+employer/.test(n)) return true;
  if (/name\s+and\s+address\s+of\s+(?:the\s+)?contractor/.test(n) && !/workman/.test(n)) return true;
  if (/^month$/.test(n) || /^year$/.test(n) || /^month\s*\/\s*year$/.test(n)) return true;
  return false;
}

export function isFormXXVITamilNaduNumberOfDaysWorkedHeader(header) {
  const n = formXXVITamilNaduHeaderNorm(header).replace(/:+$/, '').trim();
  if (!n) return false;
  if (/number\s+of\s+days?\s+worked/.test(n)) return true;
  if (/^no\.?\s*of\s+days?\s+worked$/.test(n)) return true;
  if (/days?\s+worked/.test(n) && !/daily|hours|overtime|entry|termination/.test(n)) return true;
  return false;
}

export function isFormXXVITamilNaduRateOfWagesHeader(header) {
  const n = formXXVITamilNaduHeaderNorm(header).replace(/:+$/, '').trim();
  if (!n) return false;
  if (/overtime/.test(n)) return false;
  if (/rate\s+of\s+wages?/.test(n)) return true;
  if (/^rate\s+of\s+w/.test(n)) return true;
  if (/^wages?\s+rate$/.test(n)) return true;
  return false;
}

/** Drop worksite / principal / contractor header labels from the data grid. */
export function stripFormXXVITamilNaduNonTableHeaders(headers) {
  const list = Array.isArray(headers) ? headers : [];
  return list.filter((h) => !isFormXXVITamilNaduNonTableHeaderField(h));
}

export function findFormXXVITamilNaduNumberOfDaysWorkedHeader(headers) {
  const list = Array.isArray(headers) ? headers : [];
  return list.find((h) => isFormXXVITamilNaduNumberOfDaysWorkedHeader(h)) || null;
}

export function findFormXXVITamilNaduRateOfWagesHeader(headers) {
  const list = Array.isArray(headers) ? headers : [];
  return list.find((h) => isFormXXVITamilNaduRateOfWagesHeader(h)) || null;
}

/** Worksite header value: site Name + Location name. */
export function buildFormXXVITamilNaduWorksiteText(site = {}, extras = {}) {
  const name = String(
    extras.siteName ??
      extras.EmployeeName ??
      extras.employeeName ??
      site?.siteName ??
      site?.SiteName ??
      site?.EmployeeName ??
      site?.establishmentName ??
      ''
  ).trim();
  const location = String(
    extras.locationName ??
      extras.LocationName ??
      extras.location ??
      site?.location ??
      site?.Location ??
      site?.LocationName ??
      site?.locationName ??
      ''
  ).trim();
  return [name, location].filter(Boolean).join(', ');
}

export function getFormXXVITamilNaduEmployeeName(emp = {}) {
  const candidates = [
    emp.EmployeeName,
    emp['Employee Name'],
    emp.employeeName,
    emp.Name,
    emp.name,
    emp.Full_Name,
    emp['Full Name'],
    emp.fullName
  ];
  for (let i = 0; i < candidates.length; i += 1) {
    const v = String(candidates[i] ?? '').trim();
    if (v) return v;
  }
  return '';
}

export function matchesFormXXVIHint(blob) {
  const parts = String(blob || '').toLowerCase();
  if (/form[\s._-]*xxvii(?![a-z])/i.test(parts)) return false;
  // Factories Form 26-A — not CLRA Form XXVI.
  if (/\bform\s*26[\s._-]*a\b|\bform_26[\s._-]*a\b|\bform\s*26a\b/i.test(parts)) return false;
  return (
    /form[\s._-]*xxvi(?![a-z])/i.test(parts) ||
    /form[\s._-]*26(?!\d)/i.test(parts) ||
    /(?:^|[\s._-])xxvi(?=[\s._\W-]|$)/i.test(parts)
  );
}

export function isFormXXVITamilNaduContext(partsOrBlob) {
  const p = String(partsOrBlob || '').toLowerCase();
  return (
    /tamil[\s._-]*nadu|tamilnadu|form_xxvi[_\s-]*tamil|form[\s._-]*xxvi[_\s-]*tamil/.test(p)
  );
}

function buildFormXXVIContextBlob(formHeader, rowItem, fileName, tableHeaders, sheetText) {
  return [
    rowItem?.formName,
    rowItem?.FormName,
    rowItem?.formFileName,
    rowItem?.FormFileName,
    rowItem?.state,
    rowItem?.State,
    fileName,
    formHeader?.title,
    formHeader?.subtitle,
    formHeader?.reference,
    sheetText,
    Array.isArray(tableHeaders) ? tableHeaders.join(' ') : tableHeaders
  ]
    .join(' ')
    .toLowerCase();
}

/** Tamil Nadu Form XXVI — register with daily hours 1–31 (CLRA multi-tab workbook). */
export function isFormXXVITamilNaduClraContext(
  formHeader,
  rowItem,
  fileName,
  tableHeaders = '',
  sheetText = ''
) {
  const parts = buildFormXXVIContextBlob(formHeader, rowItem, fileName, tableHeaders, sheetText);
  if (!matchesFormXXVIHint(parts)) return false;
  if (!isFormXXVITamilNaduContext(parts)) return false;
  // Factories Form 26 Register of Accidents — not CLRA muster / employment register.
  if (
    /register\s+of\s+accident/i.test(parts) &&
    !/muster\s+roll|register\s+of\s+employment|daily\s+hours\s+of\s+work|rule\s*75/i.test(parts)
  ) {
    return false;
  }
  const joinedHeaders = Array.isArray(tableHeaders)
    ? tableHeaders.map((h) => String(h || '').toLowerCase()).join('\n')
    : '';
  if (
    joinedHeaders &&
    /nature.*extent.*injury|nature.*location.*injury|date.*time.*accident|exact\s+place.*accident/i.test(
      joinedHeaders
    )
  ) {
    return false;
  }
  return true;
}

export function isFormXXVITamilNaduHeaderFieldLayoutFormHeader(formHeader) {
  return !!formHeader?.formXXVITamilNaduHeaderFieldLayout;
}

const GENERIC_SITE_HEADER_KEY_RE =
  /^(statutory_establishment_name|statutory_establishment_address|statutory_establishment_name_shop|form25_establishment)$/;

const isExcludedXxviHeaderField = (field) => {
  const key = String(field?.key || '');
  const label = formXXVITamilNaduHeaderNorm(field?.label);
  if (GENERIC_SITE_HEADER_KEY_RE.test(key)) return true;
  if (/^name\s+of\s+the\s+establishment/.test(label)) return true;
  if (/^address\s+of\s+the\s+establishment/.test(label)) return true;
  if (/^name\s+of\s+establishment\s*\/\s*shop/.test(label)) return true;
  return false;
};

const isMeaningfulHeaderValue = (value, label) => {
  const raw = String(value || '').trim();
  if (!raw) return false;
  const normalizedRaw = formXXVITamilNaduHeaderNorm(raw);
  const normalizedLabel = formXXVITamilNaduHeaderNorm(label);
  if (!normalizedRaw || normalizedRaw === normalizedLabel) return false;
  if (
    normalizedRaw ===
    formXXVITamilNaduHeaderNorm(`Enter ${String(label || '').replace(/:+$/, '').trim()}`)
  ) {
    return false;
  }
  return true;
};

const labelMatchesSpec = (fieldLabel, specLabel) => {
  const a = formXXVITamilNaduHeaderNorm(fieldLabel).replace(/^\d+\.\s*/, '');
  const b = formXXVITamilNaduHeaderNorm(specLabel).replace(/^\d+\.\s*/, '');
  if (!a || !b) return false;
  if (a === b) return true;
  if (a.includes(b) || b.includes(a)) return true;
  return false;
};

const pickValueForSpec = (spec, existingFields = []) => {
  const byKey = existingFields.find((f) => f?.key === spec.key);
  if (byKey && isMeaningfulHeaderValue(byKey.value, spec.label)) {
    return String(byKey.value).trim();
  }
  for (const field of existingFields) {
    if (!field) continue;
    if (field.key === spec.key) continue;
    if (labelMatchesSpec(field.label, spec.label) && isMeaningfulHeaderValue(field.value, spec.label)) {
      return String(field.value).trim();
    }
    if (
      spec.match.test(formXXVITamilNaduHeaderNorm(field.label)) &&
      isMeaningfulHeaderValue(field.value, spec.label)
    ) {
      return String(field.value).trim();
    }
  }
  return '';
};

export function finalizeFormXXVITamilNaduHeaderFields(existingFields = []) {
  const fields = Array.isArray(existingFields)
    ? existingFields.filter((f) => !isExcludedXxviHeaderField(f))
    : [];
  return FORM_XXVI_TN_HEADER_SPECS.map((spec) => ({
    label: spec.label,
    key: spec.key,
    fieldType: spec.fieldType || 'text',
    value: pickValueForSpec(spec, fields)
  }));
}

function buildWorkbookMergedCellAccessor(workbook, hints = {}) {
  if (!workbook?.SheetNames?.length) return null;
  const preferred = hints.preferredSheetName || hints.sheetName;
  const sheetName =
    (preferred && workbook.SheetNames.includes(preferred) && preferred) ||
    workbook.SheetNames.find((n) => /form[\s._-]*xxvi|form[\s._-]*26(?!\d)/i.test(String(n))) ||
    workbook.SheetNames[0];
  const ws = workbook.Sheets[sheetName];
  if (!ws) return null;
  const merges = ws['!merges'] || [];
  const ref = ws['!ref'];
  let effectiveSheetCols = 20;
  if (ref) {
    try {
      const range = typeof XLSX !== 'undefined' ? XLSX.utils.decode_range(ref) : null;
      if (range) effectiveSheetCols = Math.max(20, range.e.c + 1);
    } catch {
      effectiveSheetCols = 20;
    }
  }
  const rawCell = (r, c) => {
    if (r < 0 || c < 0) return '';
    const cellRef = typeof XLSX !== 'undefined' ? XLSX.utils.encode_cell({ r, c }) : '';
    const cell = cellRef ? ws[cellRef] : null;
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
  return { getMergedAwareCellText, effectiveSheetCols, sheetName };
}

const readValueBesideLabel = (getMergedAwareCellText, labelRow, labelCol, effectiveSheetCols) => {
  const maxC = Math.max(20, effectiveSheetCols || 0);
  for (let c = labelCol + 1; c < Math.min(labelCol + 16, maxC); c += 1) {
    const v = String(getMergedAwareCellText(labelRow, c) || '').trim();
    if (v && !/^enter\b/i.test(v)) return v;
  }
  for (let r = labelRow + 1; r <= labelRow + 4; r += 1) {
    const v = String(getMergedAwareCellText(r, labelCol) || '').trim();
    if (v && !/^enter\b/i.test(v)) return v;
  }
  return '';
};

const scanSheetValueForSpec = (getMergedAwareCellText, spec, effectiveSheetCols, maxRows = 80) => {
  if (!getMergedAwareCellText) return '';
  const maxC = Math.max(20, effectiveSheetCols || 0);
  for (let r = 0; r < maxRows; r += 1) {
    for (let c = 0; c < maxC; c += 1) {
      const raw = String(getMergedAwareCellText(r, c) || '').trim();
      if (!raw) continue;
      const norm = formXXVITamilNaduHeaderNorm(raw);
      if (!spec.match.test(norm) && !spec.match.test(raw)) continue;
      const value = readValueBesideLabel(getMergedAwareCellText, r, c, effectiveSheetCols);
      if (value) return value;
    }
  }
  return '';
};

export function buildFormXXVITamilNaduTemplateFields(
  getMergedAwareCellText,
  effectiveSheetCols,
  existingFields = []
) {
  return FORM_XXVI_TN_HEADER_SPECS.map((spec) => {
    const fromExisting = pickValueForSpec(spec, existingFields);
    const fromSheet = fromExisting || scanSheetValueForSpec(getMergedAwareCellText, spec, effectiveSheetCols);
    return {
      label: spec.label,
      key: spec.key,
      fieldType: spec.fieldType || 'text',
      value: fromSheet
    };
  });
}

/** Force correct Form XXVI TN heading when Excel still carries Form XVIII title text. */
export function enrichFormXXVITamilNaduDisplayHeader(
  formHeader,
  item = null,
  fileName = '',
  tableHeaders = [],
  sheetText = ''
) {
  if (!isFormXXVITamilNaduClraContext(formHeader, item, fileName, tableHeaders, sheetText)) {
    return formHeader;
  }
  const base = formHeader && typeof formHeader === 'object' ? { ...formHeader } : {};
  const title = String(base.title || '');
  const needsTitleFix =
    !title.trim() ||
    /form\s*xviii|wages[\s-]*cum[\s-]*muster/i.test(title) ||
    !/form\s*xxvi|form[\s._-]*26(?!\d)|muster\s+roll|register\s+of\s+employment/i.test(title);
  if (needsTitleFix) {
    base.title = FORM_XXVI_TN_TITLE;
  }
  if (
    !base.subtitle ||
    /wages[\s-]*cum[\s-]*muster|form\s+of\s+register\s+of\s+wages|muster\s+roll/i.test(
      String(base.subtitle || '')
    ) ||
    !/register\s+of\s+employment/i.test(String(base.subtitle || ''))
  ) {
    base.subtitle = FORM_XXVI_TN_SUBTITLE;
  }
  if (!base.reference || /\brule\s*78/i.test(String(base.reference || ''))) {
    base.reference = FORM_XXVI_TN_REFERENCE;
  }
  base.formXXVITamilNaduHeaderFieldLayout = true;
  base.formXVIIITamilNaduHeaderFieldLayout = false;
  base.formXXVIAPHeaderFieldLayout = false;
  const existing = Array.isArray(base.fields) ? base.fields : [];
  base.fields = finalizeFormXXVITamilNaduHeaderFields(existing);
  return base;
}

export function resolveFormXXVITamilNaduHeaderFieldLayout(parsed, workbook, hints = {}) {
  const formHeader = hints.formHeader || parsed?.formHeader || null;
  const item = hints.item || null;
  const fileName = hints.fileName || hints.formFileName || '';
  const sheetText = hints.sheetText || '';
  const tableHeaders = hints.tableHeaders || parsed?.headers || [];
  if (!isFormXXVITamilNaduClraContext(formHeader, item, fileName, tableHeaders, sheetText)) {
    return null;
  }

  let getMergedAwareCellText = null;
  let effectiveSheetCols = 20;
  const accessor = buildWorkbookMergedCellAccessor(workbook, hints);
  if (accessor) {
    getMergedAwareCellText = accessor.getMergedAwareCellText;
    effectiveSheetCols = accessor.effectiveSheetCols;
  }

  const existingFields = Array.isArray(formHeader?.fields) ? formHeader.fields : [];
  const finalFields = buildFormXXVITamilNaduTemplateFields(
    getMergedAwareCellText,
    effectiveSheetCols,
    existingFields
  );

  return {
    formHeader: {
      ...formHeader,
      title: FORM_XXVI_TN_TITLE,
      subtitle: /register\s+of\s+employment/i.test(String(formHeader?.subtitle || ''))
        ? formHeader.subtitle
        : FORM_XXVI_TN_SUBTITLE,
      reference: formHeader?.reference || FORM_XXVI_TN_REFERENCE,
      formXXVITamilNaduHeaderFieldLayout: true,
      formXVIIITamilNaduHeaderFieldLayout: false,
      formXXVIAPHeaderFieldLayout: false,
      fields: finalFields
    }
  };
}

export function applyFormXXVITamilNaduAutofillFromSite(headerData, context = {}) {
  const out = headerData && typeof headerData === 'object' ? { ...headerData } : {};
  const fill = (key, value) => {
    const text = String(value ?? '').trim();
    if (!text || !key) return;
    const cur = String(out[key] ?? '').trim();
    if (!cur || /^enter\b/i.test(cur)) out[key] = text;
  };
  const worksiteText =
    String(context.worksiteText || '').trim() ||
    buildFormXXVITamilNaduWorksiteText(
      {},
      {
        siteName: context.siteName || context.establishmentName || '',
        EmployeeName: context.EmployeeName || context.employeeName || '',
        locationName: context.locationName || context.LocationName || context.natureLocationText || '',
        location: context.location || ''
      }
    );
  fill('form_xxvi_principal_employer', context.principalEmployerText || '');
  fill('form_xxvi_contractor', context.contractorText || '');
  fill('form_xxvi_worksite', worksiteText);
  fill('form_x_month', context.monthName || '');
  fill('form_x_year', context.year || '');
  return out;
}

function pickFormXXVITamilNaduPaidDaysValue(payrollRow) {
  if (!payrollRow || typeof payrollRow !== 'object' || payrollRow.fetch_error) return '';
  const flat = flattenPayrollEarningColumns(payrollRow);
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
    'effective_paid_days'
  ];
  const patterns = [
    /^paid_days$/,
    /^paiddays$/,
    /paid_days/,
    /daysworked/,
    /days_present/,
    /noofdayspresent/,
    /no_of_days_present/,
    /effective_paid_days/
  ];
  const fromFlat = readPayrollScalar(flat, keys, patterns);
  if (fromFlat !== '' && fromFlat != null) {
    const num = Number(String(fromFlat).replace(/,/g, '').trim());
    if (Number.isFinite(num) && num >= 0) return num;
  }
  if (payrollRow !== flat) {
    const fromRow = readPayrollScalar(payrollRow, keys, patterns);
    if (fromRow !== '' && fromRow != null) {
      const num = Number(String(fromRow).replace(/,/g, '').trim());
      if (Number.isFinite(num) && num >= 0) return num;
    }
  }
  return '';
}

/** Rate of Wages ← SamplePayroll gross_pay. */
function pickFormXXVITamilNaduRateOfWagesValue(payrollRow) {
  if (!payrollRow || typeof payrollRow !== 'object' || payrollRow.fetch_error) return '';
  const flat = flattenPayrollEarningColumns(payrollRow);
  const keys = [
    'gross_pay',
    'Gross_pay',
    'grossPay',
    'GrossPay',
    'gross',
    'Gross',
    'total_earnings',
    'Total_earnings'
  ];
  const patterns = [
    /^gross_pay$/,
    /^grosspay$/,
    /^gross$/,
    /gross_pay/,
    /total_earnings/
  ];
  const fromFlat = readPayrollScalar(flat, keys, patterns);
  if (fromFlat !== '' && fromFlat != null) {
    const num = Number(String(fromFlat).replace(/,/g, '').trim());
    if (Number.isFinite(num) && num > 0) return num;
  }
  if (payrollRow !== flat) {
    const fromRow = readPayrollScalar(payrollRow, keys, patterns);
    if (fromRow !== '' && fromRow != null) {
      const num = Number(String(fromRow).replace(/,/g, '').trim());
      if (Number.isFinite(num) && num > 0) return num;
    }
  }
  return '';
}

/** First / last name from People emp or Sample Payroll (both required for Rate of Wages). */
export function readFormXXVITamilNaduPersonNameParts(empOrPayroll = null, payrollRow = null) {
  const pick = (src) => {
    if (!src || typeof src !== 'object') return { firstName: '', lastName: '' };
    const flat =
      src.employee_name || src.paid_days != null || src.gross_pay != null || src.payroll_payload
        ? { ...flattenPayrollEarningColumns(src), ...src }
        : src;
    const firstName = String(
      flat.FirstName ??
        flat.firstName ??
        flat.first_name ??
        flat.firstname ??
        flat['First Name'] ??
        flat.First_Name ??
        ''
    ).trim();
    const lastName = String(
      flat.LastName ??
        flat.lastName ??
        flat.last_name ??
        flat.lastname ??
        flat['Last Name'] ??
        flat.Last_Name ??
        flat.Surname ??
        flat.surname ??
        ''
    ).trim();
    return { firstName, lastName };
  };

  const fromEmp = pick(empOrPayroll);
  if (fromEmp.firstName && fromEmp.lastName) return fromEmp;
  const fromPayroll = pick(payrollRow || (empOrPayroll !== payrollRow ? empOrPayroll : null));
  if (fromPayroll.firstName && fromPayroll.lastName) return fromPayroll;
  return {
    firstName: fromEmp.firstName || fromPayroll.firstName || '',
    lastName: fromEmp.lastName || fromPayroll.lastName || ''
  };
}

export function hasFormXXVITamilNaduPersonNameParts(empOrPayroll = null, payrollRow = null) {
  const { firstName, lastName } = readFormXXVITamilNaduPersonNameParts(empOrPayroll, payrollRow);
  return Boolean(firstName && lastName);
}

/**
 * Rate of Wages ← gross_pay only when the person has both firstname and lastname
 * (People or Sample Payroll).
 */
export function resolveFormXXVITamilNaduRateOfWages(payrollRow, emp = null) {
  if (!hasFormXXVITamilNaduPersonNameParts(emp, payrollRow)) return '';
  const rate = pickFormXXVITamilNaduRateOfWagesValue(payrollRow);
  return rate === '' ? '' : String(rate);
}

function normalizeFormXXVITamilNaduMatchKey(value) {
  return String(value ?? '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ');
}

function formXXVITamilNaduEmployeeMatchKeys(empOrRow = {}) {
  const src = empOrRow && typeof empOrRow === 'object' ? empOrRow : {};
  const flat =
    src.employee_name || src.paid_days != null || src.payroll_payload
      ? flattenPayrollEarningColumns(src)
      : src;
  const vals = [
    flat.EmployeeName,
    flat['Employee Name'],
    flat.employeeName,
    flat.employee_name,
    flat.full_name,
    flat.FullName,
    flat['Full Name'],
    flat.Name,
    flat.name,
    flat['Name of the Workman'],
    flat['Name of the Worker'],
    flat.EmployeeID,
    flat.EmployeeId,
    flat.employeeId,
    flat.employee_id,
    flat.Employee_ID,
    flat['Employee ID'],
    flat.employee_number,
    flat.Employee_Number,
    flat.Zoho_ID,
    flat.ZohoID,
    flat.email,
    flat.Email,
    flat.work_email
  ]
    .map((v) => normalizeFormXXVITamilNaduMatchKey(v))
    .filter(Boolean);
  // Also pick any obvious name-like keys from the row (Excel leaf headers vary).
  Object.keys(flat).forEach((key) => {
    if (String(key).startsWith('__')) return;
    const kn = normalizeFormXXVITamilNaduMatchKey(key);
    if (
      /name\s+of\s+(?:the\s+)?workm[ae]n|name\s+of\s+(?:the\s+)?worker|employee\s*name|^name$/.test(
        kn
      )
    ) {
      const v = normalizeFormXXVITamilNaduMatchKey(flat[key]);
      if (v) vals.push(v);
    }
  });
  return Array.from(new Set(vals));
}

function formXXVITamilNaduKeysLooselyMatch(a, b) {
  const left = normalizeFormXXVITamilNaduMatchKey(a);
  const right = normalizeFormXXVITamilNaduMatchKey(b);
  if (!left || !right) return false;
  if (left === right) return true;
  // Avoid substring theft ("raja" → "rajeshkumar"). Only allow full-token equality
  // after stripping spaces/punctuation already done by normalize.
  return false;
}

/** Rate of Wages / Number of Days Worked — SamplePayroll only (never People fuzzy-fill). */
export function isFormXXVITamilNaduPayrollOnlyHeader(header) {
  return (
    isFormXXVITamilNaduNumberOfDaysWorkedHeader(header) ||
    isFormXXVITamilNaduRateOfWagesHeader(header)
  );
}

/**
 * Form XXVI payroll columns from SamplePayroll:
 * - Number of Days Worked ← Paid_days
 * - Rate of Wages ← gross_pay (only when firstname + lastname present)
 * Prefer resolvePayrollRow(emp, row, index) when provided (same path as Form V).
 * Mutates rows in place; returns { paidDaysHits, rateHits }.
 * When a person has no SamplePayroll match (or empty Paid_days / gross_pay), those
 * columns are cleared so Autofill never shows another employee's leftover values.
 */
export function applyFormXXVITamilNaduPaidDaysToMappedRows(
  rows,
  headers,
  payrollRows = [],
  options = {}
) {
  const daysHeader = findFormXXVITamilNaduNumberOfDaysWorkedHeader(headers);
  const rateHeader = findFormXXVITamilNaduRateOfWagesHeader(headers);
  if ((!daysHeader && !rateHeader) || !Array.isArray(rows) || rows.length === 0) {
    return { paidDaysHits: 0, rateHits: 0 };
  }

  const overwrite = options.overwrite !== false;
  const resolvePayrollRow =
    typeof options.resolvePayrollRow === 'function' ? options.resolvePayrollRow : null;
  const unwrapEmp =
    typeof options.unwrapEmp === 'function'
      ? options.unwrapEmp
      : (item) => (item && (item.Employee || item.employee || item)) || null;
  const employeesForMapping = Array.isArray(options.employeesForMapping)
    ? options.employeesForMapping
    : [];
  const payrollList = Array.isArray(payrollRows) ? payrollRows : [];

  const clearPayrollOnlyColumns = (row) => {
    if (!overwrite || !row || typeof row !== 'object') return;
    if (daysHeader) row[daysHeader] = '';
    if (rateHeader) row[rateHeader] = '';
  };

  const payrollByKey = new Map();
  payrollList.forEach((pr) => {
    const paid = pickFormXXVITamilNaduPaidDaysValue(pr);
    const rate = pickFormXXVITamilNaduRateOfWagesValue(pr);
    if (paid === '' && rate === '') return;
    formXXVITamilNaduEmployeeMatchKeys(pr).forEach((k) => {
      if (!payrollByKey.has(k)) payrollByKey.set(k, { paid, rate, payrollRow: pr });
    });
  });

  const resolvePayrollForRow = (row, index) => {
    if (resolvePayrollRow) {
      const emp = unwrapEmp(employeesForMapping[index] || null);
      const payrollRow = resolvePayrollRow(emp, row, index);
      // Resolver owns matching — never loosely steal another employee's payroll.
      return payrollRow || null;
    }
    const keys = formXXVITamilNaduEmployeeMatchKeys(row);
    for (let i = 0; i < keys.length; i += 1) {
      if (payrollByKey.has(keys[i])) {
        const cached = payrollByKey.get(keys[i]);
        return {
          paid_days: cached.paid,
          gross_pay: cached.rate,
          ...(cached.payrollRow && typeof cached.payrollRow === 'object' ? cached.payrollRow : {}),
          __fromKeyCache: true
        };
      }
    }
    if (keys.length > 0 && payrollByKey.size > 0) {
      for (const [pk, cached] of payrollByKey.entries()) {
        if (keys.some((k) => formXXVITamilNaduKeysLooselyMatch(k, pk))) {
          return {
            paid_days: cached.paid,
            gross_pay: cached.rate,
            ...(cached.payrollRow && typeof cached.payrollRow === 'object' ? cached.payrollRow : {}),
            __fromKeyCache: true
          };
        }
      }
    }
    return null;
  };

  let paidDaysHits = 0;
  let rateHits = 0;
  rows.forEach((row, index) => {
    if (!row || typeof row !== 'object') return;
    const emp = unwrapEmp(employeesForMapping[index] || null);
    const payrollRow = resolvePayrollForRow(row, index);
    if (!payrollRow || payrollRow.fetch_error) {
      // No SamplePayroll row for this person — leave Rate / Days Worked blank.
      clearPayrollOnlyColumns(row);
      return;
    }

    if (daysHeader) {
      const cur = String(row[daysHeader] ?? '').trim();
      if (overwrite || !cur) {
        const paid = pickFormXXVITamilNaduPaidDaysValue(payrollRow);
        if (paid !== '') {
          row[daysHeader] = String(paid);
          paidDaysHits += 1;
        } else if (overwrite) {
          row[daysHeader] = '';
        }
      }
    }
    if (rateHeader) {
      const cur = String(row[rateHeader] ?? '').trim();
      if (overwrite || !cur) {
        const rate = resolveFormXXVITamilNaduRateOfWages(payrollRow, emp);
        if (rate !== '') {
          row[rateHeader] = rate;
          rateHits += 1;
        } else if (overwrite) {
          // Missing firstname/lastname or empty gross_pay — never keep leftover values.
          row[rateHeader] = '';
        }
      }
    }
  });
  return { paidDaysHits, rateHits };
}

export { FORM_XXVI_TN_HEADER_KEYS };

export function isFormXXVITamilNaduDayHeaderKey(headerKey) {
  const h = String(headerKey || '').trim();
  // Bare 1–31 only count as day keys when they are clearly day markers (not column index leftovers alone).
  if (/^10_(\d{1,2})$/.test(h)) {
    const n = Number(h.match(/^10_(\d{1,2})$/)[1]);
    return n >= 1 && n <= 31;
  }
  if (/^daily\s+hours\s+of\s+work_(\d{1,2})$/i.test(h)) {
    const n = Number(h.match(/^daily\s+hours\s+of\s+work_(\d{1,2})$/i)[1]);
    return n >= 1 && n <= 31;
  }
  if (/^(?:day|dates?|attendance)[_\s-]*(\d{1,2})$/i.test(h)) {
    const n = Number(h.match(/(\d{1,2})$/)[1]);
    return n >= 1 && n <= 31;
  }
  // Bare numeric headers are day columns only in a contiguous day band context — handled by callers.
  if (/^\d{1,2}$/.test(h)) {
    const n = Number(h);
    return n >= 1 && n <= 31;
  }
  const m = h.match(/_(\d{1,2})$/);
  if (!m) return false;
  const n = Number(m[1]);
  if (n < 1 || n > 31) return false;
  // Avoid treating "Column 9" / identity leftovers as day keys.
  if (/column|serial|rate|wage|address|name|age|sex|father|husband|designation|entry|local|permanent/i.test(h)) {
    return false;
  }
  return true;
}

/** True when header is Rate of Wages (must stay left of the day band — never a day column). */
export function isFormXXVITamilNaduRateOfWagesHeaderKey(header) {
  return isFormXXVITamilNaduRateOfWagesHeader(header);
}

export function resolveFormXXVITamilNaduDayNumberFromHeader(header) {
  const h = String(header || '').trim();
  const m1 = h.match(/^10_(\d{1,2})$/i);
  if (m1) return Number(m1[1]);
  const m2 = h.match(/^daily\s+hours\s+of\s+work_(\d{1,2})$/i);
  if (m2) return Number(m2[1]);
  const m3 = h.match(/_(\d{1,2})$/);
  if (m3) return Number(m3[1]);
  if (/^\d{1,2}$/.test(h)) return Number(h);
  return 0;
}

export function listFormXXVITamilNaduDayHeaders(headers) {
  const out = [];
  (Array.isArray(headers) ? headers : []).forEach((header) => {
    const day = resolveFormXXVITamilNaduDayNumberFromHeader(header);
    if (day >= 1 && day <= 31) out.push({ header, day });
  });
  return out;
}

function formXXVITamilNaduDayKeyStyle(existingDayHeader) {
  const h = String(existingDayHeader || '').trim();
  if (/^10_(\d{1,2})$/i.test(h)) return { kind: 'band', parent: '10' };
  const daily = h.match(/^(daily\s+hours\s+of\s+work)_(\d{1,2})$/i);
  if (daily) return { kind: 'band', parent: daily[1] };
  if (/^\d{1,2}$/.test(h)) return { kind: 'bare' };
  const m = h.match(/^(.*_)(\d{1,2})$/);
  if (m) return { kind: 'band', parent: m[1].replace(/_$/, '') };
  return { kind: 'band', parent: FORM_XXVI_TN_DAY_BAND_PARENT };
}

function synthesizeFormXXVITamilNaduDayHeader(day, style) {
  if (style?.kind === 'bare') return String(day);
  const parent = style?.parent || FORM_XXVI_TN_DAY_BAND_PARENT;
  return `${parent}_${day}`;
}

/**
 * When parsed headers only capture a partial day band (often 1–14 from the form
 * column-index row), synthesize missing days through the end of the month so the
 * Autofill UI shows the full Register of Employment grid.
 */
export function ensureFormXXVITamilNaduDayColumnHeaders(headers, daysInMonth = 31) {
  const monthDays = Math.min(Math.max(Number(daysInMonth) || 31, 28), 31);
  const base = Array.isArray(headers) ? [...headers] : [];
  while (base.length > 0 && !String(base[base.length - 1] || '').trim()) base.pop();

  const dayByNum = new Map();
  const prefix = [];
  const suffix = [];
  let seenDay = false;

  base.forEach((header) => {
    if (isFormXXVITamilNaduNonTableHeaderField(header)) return;
    if (isFormXXVITamilNaduRateOfWagesHeaderKey(header)) {
      if (!seenDay) prefix.push(header);
      else suffix.push(header);
      return;
    }
    if (isFormXXVITamilNaduNumberOfDaysWorkedHeader(header)) {
      seenDay = true;
      suffix.push(header);
      return;
    }
    const day = resolveFormXXVITamilNaduDayNumberFromHeader(header);
    const looksDay =
      day >= 1 &&
      day <= 31 &&
      isFormXXVITamilNaduDayHeaderKey(header) &&
      !/serial|name|age|sex|address|designation|father|husband|entry|termination|signature|thumb|contractor|representative|wages|rate/i.test(
        String(header || '')
      );
    if (looksDay) {
      seenDay = true;
      if (!dayByNum.has(day)) dayByNum.set(day, header);
      return;
    }
    if (!seenDay) prefix.push(header);
    else suffix.push(header);
  });

  if (dayByNum.size === 0) return base;
  if (dayByNum.size >= monthDays) {
    const days = [...dayByNum.entries()]
      .sort((a, b) => a[0] - b[0])
      .map(([, header]) => header);
    return [...prefix, ...days, ...suffix];
  }

  // Partial band (e.g. only 1–14) — pad missing days using the same key style.
  const sample = dayByNum.values().next().value;
  const style = formXXVITamilNaduDayKeyStyle(sample);
  for (let day = 1; day <= monthDays; day += 1) {
    if (!dayByNum.has(day)) {
      dayByNum.set(day, synthesizeFormXXVITamilNaduDayHeader(day, style));
    }
  }
  const days = [...dayByNum.entries()]
    .sort((a, b) => a[0] - b[0])
    .map(([, header]) => header);
  return [...prefix, ...days, ...suffix];
}

/** Rebuild group-label row after day-column expansion (Daily hours of work band). */
export function buildFormXXVITamilNaduColumnGroupLabels(headers, parentLabel = 'Daily hours of work') {
  const list = Array.isArray(headers) ? headers : [];
  const labels = new Array(list.length).fill('');
  const parent = String(parentLabel || 'Daily hours of work').trim() || 'Daily hours of work';
  list.forEach((h, i) => {
    const day = resolveFormXXVITamilNaduDayNumberFromHeader(h);
    if (
      day >= 1 &&
      day <= 31 &&
      isFormXXVITamilNaduDayHeaderKey(h) &&
      !isFormXXVITamilNaduRateOfWagesHeaderKey(h) &&
      !isFormXXVITamilNaduNumberOfDaysWorkedHeader(h)
    ) {
      labels[i] = parent;
    } else {
      labels[i] = String(h || '').trim();
    }
  });
  return labels;
}

const headerAliasBucket = (key) => {
  const n = formXXVITamilNaduHeaderNorm(key).replace(/[^a-z0-9]/g, '');
  if (/^srno|serialno|serialnumber|sno/.test(n)) return 'srno';
  if (/nameoftheworkman|nameoftheworker|workmanname|workername|employeename/.test(n)) return 'name';
  if (/agesex|ageandsex/.test(n)) return 'agesex';
  if (/permanenthomeaddress|permanentaddress/.test(n)) return 'permaddr';
  if (/localaddress/.test(n)) return 'localaddr';
  if (/designation|natureofwork/.test(n)) return 'designation';
  if (/fathersname|husbandsname|fatherhusband/.test(n)) return 'fatherhusband';
  if (/dateofentry|dateofjoining|entryintoservice/.test(n)) return 'doj';
  if (/rateofwages|^wagesrate$/.test(n) && !/overtime/.test(n)) return 'wages';
  // Avoid matching generic "wages" keys to Rate of Wages (attendance/export bleed).
  if (/^wages$/.test(n)) return 'wages';
  return '';
};

export function readFormXXVITamilNaduCellValue(row, header) {
  if (!row || typeof row !== 'object') return '';
  const direct = row[header];
  if (direct != null && String(direct).trim() !== '') return direct;

  const day = resolveFormXXVITamilNaduDayNumberFromHeader(header);
  if (day >= 1 && day <= 31) {
    const aliases = [
      `${FORM_XXVI_TN_DAY_BAND_PARENT}_${day}`,
      `Daily hours of work_${day}`,
      `daily hours of work_${day}`,
      String(day),
      `Day ${day}`,
      `DATES_${day}`
    ];
    for (let i = 0; i < aliases.length; i += 1) {
      const val = row[aliases[i]];
      if (val != null && String(val).trim() !== '') return val;
    }
  }

  const bucket = headerAliasBucket(header);
  if (bucket) {
    for (const [key, val] of Object.entries(row)) {
      if (String(key).startsWith('__')) continue;
      if (headerAliasBucket(key) === bucket && val != null && String(val).trim() !== '') {
        return val;
      }
    }
  }

  const target = formXXVITamilNaduHeaderNorm(header);
  if (target) {
    const keys = Object.keys(row);
    const exact = keys.filter((k) => formXXVITamilNaduHeaderNorm(k) === target);
    if (exact.length === 1) return row[exact[0]];
    // For Rate of Wages, never fuzzy-match short keys (avoids picking day codes like "A").
    if (isFormXXVITamilNaduRateOfWagesHeaderKey(header)) {
      const wageKey = keys.find((k) => isFormXXVITamilNaduRateOfWagesHeaderKey(k));
      if (wageKey && row[wageKey] != null && String(row[wageKey]).trim() !== '') return row[wageKey];
      return '';
    }
    const fuzzy = keys.find((k) => {
      const n = formXXVITamilNaduHeaderNorm(k);
      return n && n.length >= 4 && (n.includes(target) || target.includes(n));
    });
    if (fuzzy) return row[fuzzy];
  }

  return '';
}

export function detectFormXXVITamilNaduDayColumnMap(getCell, headerRow, startCol, maxCols) {
  const dayCols = new Map();
  const scanFrom = Math.max(startCol, 1);
  const scanTo = Math.max(maxCols, scanFrom + 45);

  const scoreMap = (cols) => {
    if (!cols || cols.size < 5) return -1;
    const day1 = cols.get(1) ?? cols.get(2);
    if (day1 == null) return -1;
    // Contiguity: how many successive days sit on successive columns.
    let run = 0;
    for (let d = 1; d <= 31; d += 1) {
      if (!cols.has(d)) break;
      if (d > 1 && cols.get(d) !== cols.get(d - 1) + 1) break;
      run += 1;
    }
    // Column-index rows are typically (1)–(14) starting at the table edge — deprioritize.
    const looksLikeColIndexRow = cols.size <= 16 && day1 <= scanFrom + 1;
    if (looksLikeColIndexRow && run < 20) return run * 0.25;
    return run * 10 + cols.size;
  };

  let bestMap = new Map();
  let bestScore = -1;
  for (let r = headerRow; r <= headerRow + 5; r += 1) {
    // Last occurrence wins so a left-side "(1)" cannot override true day-1 under the hours band.
    const cols = new Map();
    for (let c = scanFrom; c < scanTo; c += 1) {
      const raw = String(getCell(r, c) || '')
        .trim()
        .replace(/[()]/g, '');
      if (!/^\d{1,2}$/.test(raw)) continue;
      const n = Number(raw);
      if (n >= 1 && n <= 31) cols.set(n, c);
    }
    const score = scoreMap(cols);
    if (score > bestScore) {
      bestScore = score;
      bestMap = cols;
    }
  }
  if (bestScore < 5 || bestMap.size < 5) return dayCols;
  if (!bestMap.has(1) && !bestMap.has(2)) return dayCols;
  // Reject maps that still look like identity column numbers (1)–(9)/(1)–(14).
  const day1Col = bestMap.get(1) ?? bestMap.get(2);
  if (bestMap.size <= 16 && day1Col != null && day1Col <= scanFrom + 1) return dayCols;
  return bestMap;
}

function excelCellValueToString(val) {
  if (val == null) return '';
  if (typeof val === 'string' || typeof val === 'number' || typeof val === 'boolean') return String(val);
  if (val instanceof Date) return val.toISOString();
  if (typeof val === 'object') {
    if (Array.isArray(val.richText)) return val.richText.map((rt) => rt?.text || '').join('');
    if (val.text != null) return String(val.text);
    if (val.result != null) return String(val.result);
  }
  return '';
}

function pickExportCell(v) {
  if (v == null || String(v).trim() === '' || /^enter\s+/i.test(String(v).trim())) return '';
  return v;
}

function rowLooksMeaningful(row) {
  const hasAttendanceCode = (vals) =>
    vals.some((v) => /^(P|A|WO|H|L|WOP|OD|SL|CL|EL)$/i.test(String(v).trim()));
  if (Array.isArray(row)) {
    const vals = row.map((v) => String(v ?? '').trim()).filter(Boolean);
    if (vals.length === 0) return false;
    return (
      hasAttendanceCode(vals) ||
      vals.some((v) => /[a-z]/i.test(v)) ||
      !vals.every((v) => /^-?\d+(\.\d+)?$/.test(v))
    );
  }
  if (row && typeof row === 'object') {
    const vals = Object.values(row).map((v) => String(v ?? '').trim()).filter(Boolean);
    if (vals.length === 0) return false;
    return (
      hasAttendanceCode(vals) ||
      vals.some((v) => /[a-z]/i.test(v)) ||
      !vals.every((v) => /^-?\d+(\.\d+)?$/.test(v))
    );
  }
  return false;
}

function resolveWorksheet(workbook, sheetNameHint) {
  const hint = String(sheetNameHint || '').trim();
  if (hint) {
    const exact = workbook.worksheets.find((ws) => String(ws?.name || '') === hint);
    if (exact) return exact;
    const fuzzy = workbook.worksheets.find((ws) =>
      /xxvi|form[\s._-]*26(?!\d)/i.test(String(ws?.name || ''))
    );
    if (fuzzy) return fuzzy;
  }
  return (
    workbook.worksheets.find((ws) => /xxvi|form[\s._-]*26(?!\d)/i.test(String(ws?.name || ''))) ||
    workbook.worksheets[0]
  );
}

/** Form XXVI Tamil Nadu Excel export — map prefix columns + day grid (10_1..10_31) to template positions. */
export async function buildFormXXVITamilNaduWorkbookWithTemplateStyles({
  templateArrayBuffer,
  mappedData,
  mappedRowMatrix,
  headersToUse,
  parsedHeaderRowIndex,
  parsedDataStartIndex,
  parsedTableStartCol,
  parsedFormHeader,
  headerFormData,
  formFileName,
  sheetNameHint
}) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(templateArrayBuffer);
  const worksheet = resolveWorksheet(workbook, sheetNameHint);
  if (!worksheet) throw new Error('Template worksheet not found.');

  const normalize = formXXVITamilNaduHeaderNorm;
  const maxScanRows = Math.max(120, worksheet.rowCount + 10);
  const maxScanCols = 300;

  let headerRow = parsedHeaderRowIndex != null && parsedHeaderRowIndex >= 0 ? parsedHeaderRowIndex + 1 : -1;
  if (headerRow < 1) {
    for (let r = 1; r <= maxScanRows; r += 1) {
      let hits = 0;
      for (let c = 1; c <= maxScanCols; c += 1) {
        const t = normalize(excelCellValueToString(worksheet.getCell(r, c)?.value));
        if (!t) continue;
        if (
          /sr\.?\s*no|serial|s\.?\s*no/.test(t) ||
          /name\s+of\s+the\s+workman|name\s+of\s+the\s+worker/.test(t) ||
          /daily\s+hours/.test(t)
        ) {
          hits += 1;
        }
      }
      if (hits >= 2) {
        headerRow = r;
        break;
      }
    }
  }
  if (headerRow < 1) throw new Error('Could not locate Form XXVI header row.');

  let startCol = 1;
  for (let c = 1; c <= maxScanCols; c += 1) {
    const t = normalize(excelCellValueToString(worksheet.getCell(headerRow, c)?.value));
    if (t) {
      startCol = c;
      break;
    }
  }
  if (parsedTableStartCol != null && Number(parsedTableStartCol) >= 0) {
    startCol = Number(parsedTableStartCol) + 1;
  }

  const getCell = (r, c) => excelCellValueToString(worksheet.getCell(r, c)?.value);
  const dayColumnMap = detectFormXXVITamilNaduDayColumnMap(getCell, headerRow, startCol, maxScanCols);
  const markerCols = [...dayColumnMap.entries()]
    .sort((a, b) => a[0] - b[0])
    .map(([, col]) => col);

  let markerRow = -1;
  if (markerCols.length >= 5) {
    for (let r = headerRow; r <= Math.min(headerRow + 5, maxScanRows); r += 1) {
      let hits = 0;
      for (let i = 0; i < markerCols.length; i += 1) {
        const t = normalize(getCell(r, markerCols[i]));
        if (/^\d{1,2}$/.test(t) && Number(t) >= 1 && Number(t) <= 31) hits += 1;
      }
      if (hits >= 5) {
        markerRow = r;
        break;
      }
    }
  }

  let dataStartRow =
    parsedDataStartIndex != null && parsedDataStartIndex >= 0
      ? parsedDataStartIndex + 1
      : markerRow > 0
        ? markerRow + 1
        : headerRow + 3;
  if (markerRow > 0) {
    dataStartRow = Math.max(dataStartRow, markerRow + 1);
  }

  const sourceHeaders = Array.isArray(headersToUse) ? [...headersToUse] : [];
  // Prefer 10_N / Daily hours_* keys so bare column numbers (1)–(9) never become dayStartIdx.
  const dayStartIdx = (() => {
    const preferred = sourceHeaders.findIndex(
      (h) =>
        /^10_(\d{1,2})$/i.test(String(h || '').trim()) ||
        /^daily\s+hours\s+of\s+work_(\d{1,2})$/i.test(String(h || '').trim())
    );
    if (preferred >= 0) return preferred;
    return sourceHeaders.findIndex((h) => {
      if (isFormXXVITamilNaduRateOfWagesHeaderKey(h)) return false;
      return isFormXXVITamilNaduDayHeaderKey(h);
    });
  })();
  const prefixCount = dayStartIdx >= 0 ? dayStartIdx : sourceHeaders.length;
  const firstDayCol = markerCols.length > 0 ? markerCols[0] : Math.max(1, startCol + prefixCount);
  const lastDayCol =
    markerCols.length > 0
      ? markerCols[markerCols.length - 1]
      : firstDayCol + Math.max(markerCols.length, 31) - 1;

  const buildOrderedCols = () => {
    if (markerCols.length >= 5 && dayStartIdx >= 0) {
      const cols = [];
      // Anchor identity columns immediately left of day 1 (Rate of Wages = firstDayCol - 1).
      // Fall back to contiguous startCol when day-1 was mis-detected at the table edge
      // (otherwise firstDayCol - prefixCount can be negative, e.g. 1 - 9 = -8).
      let prefixStart = firstDayCol - prefixCount;
      if (prefixStart < 1 || firstDayCol <= prefixCount) {
        prefixStart = Math.max(1, startCol);
      }
      for (let j = 0; j < prefixCount; j += 1) {
        cols.push(prefixStart + j);
      }
      // Contiguous day headers only (stop at Number of Days Worked / signatures).
      let dayCount = 0;
      for (let i = dayStartIdx; i < sourceHeaders.length; i += 1) {
        const h = sourceHeaders[i];
        if (isFormXXVITamilNaduRateOfWagesHeaderKey(h)) break;
        if (!isFormXXVITamilNaduDayHeaderKey(h)) break;
        dayCount += 1;
      }
      if (dayCount < 5) dayCount = Math.min(markerCols.length, 31);
      dayCount = Math.min(dayCount, markerCols.length, 31);
      for (let d = 0; d < dayCount; d += 1) {
        const h = sourceHeaders[dayStartIdx + d];
        const dayNum = h != null ? resolveFormXXVITamilNaduDayNumberFromHeader(h) : 0;
        const mapped =
          dayNum >= 1 && dayNum <= 31 && dayColumnMap.has(dayNum)
            ? dayColumnMap.get(dayNum)
            : markerCols[d];
        cols.push(Math.max(1, Number(mapped) || prefixStart + prefixCount + d));
      }
      const rightStartIdx = dayStartIdx + dayCount;
      const rightCount = Math.max(0, sourceHeaders.length - rightStartIdx);
      if (rightCount > 0) {
        const rightStart = Math.max(1, lastDayCol + 1);
        for (let r = 0; r < rightCount; r += 1) cols.push(rightStart + r);
      }
      return cols.filter((c) => Number.isFinite(c) && c >= 1 && c <= 16384);
    }
    return Array.from({ length: Math.max(12, sourceHeaders.length) }, (_, i) => Math.max(1, startCol + i));
  };
  const orderedCols = buildOrderedCols();
  const hdrs = sourceHeaders;

  const getRowValueForHeader = (rowObj, header, headerIndex) => {
    if (!rowObj || typeof rowObj !== 'object' || Array.isArray(rowObj)) return '';
    const picked = pickExportCell(readFormXXVITamilNaduCellValue(rowObj, header));
    if (picked !== '') return picked;
    if (headerIndex != null && hdrs[headerIndex] && hdrs[headerIndex] !== header) {
      const alt = pickExportCell(readFormXXVITamilNaduCellValue(rowObj, hdrs[headerIndex]));
      if (alt !== '') return alt;
    }
    // Do NOT fall back to Object.values()[index] — key order ≠ column order and
    // shifts day-1 attendance ("A") into Rate of Wages.
    return '';
  };

  const toOrderedValues = (row, rowIndex) => {
    if (Array.isArray(row)) {
      const out = [...row];
      if (out[0] == null || String(out[0]).trim() === '') out[0] = rowIndex + 1;
      const colCount = Math.max(orderedCols.length, hdrs.length, out.length);
      return Array.from({ length: colCount }, (_, idx) => out[idx] ?? '');
    }
    const colCount = Math.max(orderedCols.length, hdrs.length);
    const out = Array.from({ length: colCount }, (_, idx) => {
      const header = hdrs[idx] || '';
      return header ? getRowValueForHeader(row, header, idx) : '';
    });
    if (out[0] == null || String(out[0]).trim() === '') out[0] = rowIndex + 1;
    return out;
  };

  const headerValues = headerFormData && typeof headerFormData === 'object' ? headerFormData : {};
  writeStatutoryHeaderFieldsToExcelJsWorksheet(worksheet, {
    headerFormData: headerValues,
    parsedFormHeader,
    headerRowEnd: headerRow,
    maxScanCols: 80
  });

  const sourcePrimary =
    Array.isArray(mappedData) && mappedData.length > 0
      ? mappedData
      : Array.isArray(mappedRowMatrix)
        ? mappedRowMatrix
        : [];
  const sourceRows = sourcePrimary
    .filter((row) => rowLooksMeaningful(row))
    .map((row, idx) => toOrderedValues(row, idx));

  const isNumericMarkerTemplateRow = (r) => {
    if (!markerCols || markerCols.length < 5) return false;
    let numericHits = 0;
    for (let i = 0; i < markerCols.length; i += 1) {
      const v = normalize(getCell(r, markerCols[i]));
      if (/^\d+$/.test(v)) numericHits += 1;
    }
    return numericHits >= Math.max(5, Math.floor(markerCols.length * 0.6));
  };
  const isHeaderBandRow = (r) => r >= headerRow && r < dataStartRow;
  const clearFromRow = Math.max(1, dataStartRow);
  const clearToRow = Math.min(maxScanRows, Math.max(dataStartRow + sourceRows.length + 30, clearFromRow + 30));
  for (let r = clearFromRow; r <= clearToRow; r += 1) {
    if (isNumericMarkerTemplateRow(r) || isHeaderBandRow(r)) continue;
    for (let j = 0; j < orderedCols.length; j += 1) {
      worksheet.getCell(r, orderedCols[j]).value = '';
    }
  }

  for (let i = 0; i < sourceRows.length; i += 1) {
    const row = sourceRows[i] || [];
    for (let j = 0; j < orderedCols.length; j += 1) {
      const value = row[j];
      if (value == null || value === '') continue;
      const cell = worksheet.getCell(dataStartRow + i, orderedCols[j]);
      if (
        j === 0 &&
        (typeof value === 'number' || (typeof value === 'string' && /^-?\d+(\.\d+)?$/.test(String(value).trim())))
      ) {
        cell.value = Number(value);
      } else {
        cell.value = String(value);
      }
    }
  }

  if (sourceRows.length > 0) {
    const tableColMin = orderedCols.length > 0 ? Math.min(...orderedCols) : startCol;
    const tableColMax = orderedCols.length > 0 ? Math.max(...orderedCols) : startCol + orderedCols.length - 1;
    ensureExcelJSDataRowsWithBorders(worksheet, {
      dataStartRow,
      dataRowCount: sourceRows.length,
      colFrom: tableColMin,
      colTo: tableColMax,
      templateRow: dataStartRow,
      templateBodyRows: 1
    });
  }

  const targetSheetName = worksheet.name;
  workbook.worksheets
    .filter((ws) => ws.name !== targetSheetName)
    .forEach((ws) => {
      workbook.removeWorksheet(ws.id);
    });

  const out = await workbook.xlsx.writeBuffer();
  const fileName =
    formFileName ||
    parsedFormHeader?.title?.replace(/[^a-zA-Z0-9]/g, '_') ||
    `Form_XXVI_TamilNadu_${Date.now()}.xlsx`;
  const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  return { blob, fileName };
}
