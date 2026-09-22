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

/** Column (2) — Name of the Workman (not contractor / employer / father). */
export function isFormXXVITamilNaduWorkmanNameHeader(header) {
  const n = formXXVITamilNaduHeaderNorm(header).replace(/:+$/, '').trim();
  if (!n) return false;
  if (/father|husband|contractor|employer|establishment|worksite|location of/.test(n)) {
    return false;
  }
  return (
    /name\s+of\s+the\s+workman/.test(n) ||
    /name\s+of\s+the\s+worker/.test(n) ||
    /workman\s*name/.test(n) ||
    (/name/.test(n) && /workm[ae]n|worker/.test(n))
  );
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
  const { firstName, lastName } = readFormXXVITamilNaduPersonNameParts(emp);
  const composed = [firstName, lastName].filter(Boolean).join(' ').trim();
  // Prefer FirstName + LastName so "Name of the Workman" is the full name.
  if (firstName && lastName) return composed;
  const candidates = [
    emp.EmployeeName,
    emp['Employee Name'],
    emp.employeeName,
    emp.Name,
    emp.name,
    emp.Full_Name,
    emp['Full Name'],
    emp.fullName,
    emp.employee_name,
    composed
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
  const { firstName, lastName } = readFormXXVITamilNaduPersonNameParts(src);
  const composedName = [firstName, lastName].filter(Boolean).join(' ').trim();
  const vals = [
    composedName,
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
  const nameHeader = (Array.isArray(headers) ? headers : []).find((h) =>
    isFormXXVITamilNaduWorkmanNameHeader(h)
  );
  if ((!daysHeader && !rateHeader && !nameHeader) || !Array.isArray(rows) || rows.length === 0) {
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
    if (nameHeader && overwrite) {
      const fromEmp = emp ? getFormXXVITamilNaduEmployeeName(emp) : '';
      const fromPayroll = payrollRow ? getFormXXVITamilNaduEmployeeName(payrollRow) : '';
      const fullName = fromEmp || fromPayroll;
      if (fullName) row[nameHeader] = fullName;
    }
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

  // Inject Rate of Wages before the day band when the parsed prefix omitted it
  // (short Serial/Name/Age prefixes used to shift Age into the wages Excel column).
  const ensuredPrefix = [...prefix];
  if (!ensuredPrefix.some((h) => isFormXXVITamilNaduRateOfWagesHeaderKey(h))) {
    ensuredPrefix.push(
      FORM_XXVI_TN_LEFT_HEADERS.find((h) => isFormXXVITamilNaduRateOfWagesHeaderKey(h)) ||
        'Rate of Wages'
    );
  } else {
    const wagesIdx = ensuredPrefix.findIndex((h) => isFormXXVITamilNaduRateOfWagesHeaderKey(h));
    if (wagesIdx >= 0 && wagesIdx !== ensuredPrefix.length - 1) {
      const [wagesHdr] = ensuredPrefix.splice(wagesIdx, 1);
      ensuredPrefix.push(wagesHdr);
    }
  }

  let ensuredSuffix = [...suffix];
  if (!ensuredSuffix.some((h) => isFormXXVITamilNaduNumberOfDaysWorkedHeader(h))) {
    ensuredSuffix = [FORM_XXVI_TN_RIGHT_HEADERS[0], ...ensuredSuffix];
  }

  if (dayByNum.size >= monthDays) {
    const days = [...dayByNum.entries()]
      .sort((a, b) => a[0] - b[0])
      .map(([, header]) => header);
    return [...ensuredPrefix, ...days, ...ensuredSuffix];
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
  return [...ensuredPrefix, ...days, ...ensuredSuffix];
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

export function isFormXXVITamilNaduTableTextHeaderRow(getCell, row, maxCols = 80) {
  let hits = 0;
  for (let c = 1; c <= maxCols; c += 1) {
    const t = formXXVITamilNaduHeaderNorm(getCell(row, c));
    if (!t) continue;
    if (/sr\.?\s*no|serial|s\.?\s*no/.test(t)) hits += 1;
    if (/name\s+of\s+the\s+work(?:er|man)/.test(t)) hits += 1;
    if (/daily\s+hours|rate\s+of\s+wages|permanent\s+home/.test(t)) hits += 1;
  }
  return hits >= 2;
}

/** Resolve the descriptive table header row (Sr. No / Name / Daily hours), not the 1–14 or 1–31 numeric rows. */
export function resolveFormXXVITamilNaduTableTextHeaderRow(
  getCell,
  parsedHeaderRowIndex,
  maxScanRows = 120,
  maxCols = 300
) {
  const isTextRow = (r) => r >= 1 && isFormXXVITamilNaduTableTextHeaderRow(getCell, r, maxCols);
  let parsed = parsedHeaderRowIndex != null && parsedHeaderRowIndex >= 0 ? parsedHeaderRowIndex + 1 : -1;
  if (isTextRow(parsed)) {
    if (
      isFormXXVITamilNaduColumnIndexRow(getCell, parsed, maxCols) &&
      parsed > 1 &&
      isTextRow(parsed - 1)
    ) {
      return parsed - 1;
    }
    return parsed;
  }
  if (
    parsed >= 1 &&
    isFormXXVITamilNaduColumnIndexRow(getCell, parsed, maxCols) &&
    parsed > 1 &&
    isTextRow(parsed - 1)
  ) {
    return parsed - 1;
  }
  for (let r = 1; r <= maxScanRows; r += 1) {
    if (isTextRow(r)) return r;
  }
  return parsed;
}

const FORM_XXVI_TN_MONTH_YEAR_SCAN_COLS = 120;

const plainExcelCell = (val) => {
  if (val == null) return '';
  if (typeof val === 'string' || typeof val === 'number') return String(val).trim();
  if (typeof val === 'object' && val.text != null) return String(val.text).trim();
  return String(val).trim();
};

/** Write Month / Date (year) into the top-right template cells (adjacent to labels). */
export function writeFormXXVITamilNaduMonthYearToExcelJsWorksheet(worksheet, headerFormData = {}) {
  if (!worksheet || !headerFormData || typeof headerFormData !== 'object') return false;
  const month = String(headerFormData.form_x_month ?? '').trim();
  const year = String(headerFormData.form_x_year ?? '').trim();
  if (!month && !year) return false;

  const maxRows = Math.min(25, Math.max(15, worksheet.rowCount || 15));
  const maxCols = FORM_XXVI_TN_MONTH_YEAR_SCAN_COLS;
  let wroteMonth = false;
  let wroteYear = false;

  const writeValueBesideLabel = (r, labelCol, value) => {
    const val = String(value ?? '').trim();
    if (!val) return false;
    for (let ac = labelCol + 1; ac <= Math.min(labelCol + 4, maxCols); ac += 1) {
      const adjText = plainExcelCell(worksheet.getCell(r, ac)?.value);
      if (!adjText || /^enter\b/i.test(adjText) || /^month\s*:?\s*$/i.test(adjText)) {
        worksheet.getCell(r, ac).value = val;
        return true;
      }
    }
    worksheet.getCell(r, labelCol + 1).value = val;
    return true;
  };

  for (let r = 1; r <= maxRows; r += 1) {
    for (let c = 1; c <= maxCols; c += 1) {
      const raw = plainExcelCell(worksheet.getCell(r, c)?.value).replace(/\s+/g, ' ').trim();
      if (!raw) continue;
      const lower = raw.toLowerCase();
      if (month && /^month\s*:?\s*$/.test(lower)) {
        worksheet.getCell(r, c).value = 'Month';
        if (writeValueBesideLabel(r, c, month)) wroteMonth = true;
        continue;
      }
      if (year && (/^date\s*:?\s*$/.test(lower) || /^year\s*:?\s*$/.test(lower))) {
        if (writeValueBesideLabel(r, c, year)) wroteYear = true;
        continue;
      }
      if (month && /^month\s*:/i.test(raw) && !/\b(january|february|march|april|may|june|july|august|september|october|november|december)\b/i.test(raw)) {
        worksheet.getCell(r, c).value = `Month: ${month}`;
        wroteMonth = true;
      }
      if (year && (/^date\s*:/i.test(raw) || /^year\s*:/i.test(raw)) && !/\d{4}/.test(raw)) {
        worksheet.getCell(r, c).value = `${raw.split(':')[0].trim()}: ${year}`;
        wroteYear = true;
      }
    }
  }

  // Far-right band (cols AO–AQ): template labels are often outside worksheet.columnCount until scanned.
  if (!wroteMonth || !wroteYear) {
    for (let r = 3; r <= 8; r += 1) {
      for (let c = 35; c <= 55; c += 1) {
        const lower = plainExcelCell(worksheet.getCell(r, c)?.value).toLowerCase();
        if (!wroteMonth && month && /^month\s*:?\s*$/.test(lower)) {
          worksheet.getCell(r, c).value = 'Month';
          wroteMonth = writeValueBesideLabel(r, c, month);
        }
        if (!wroteYear && year && (/^date\s*:?\s*$/.test(lower) || /^year\s*:?\s*$/.test(lower))) {
          wroteYear = writeValueBesideLabel(r, c, year);
        }
      }
    }
  }

  const monthNameOnly = new RegExp(`^${month.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}$`, 'i');
  if (month) {
    for (let r = 3; r <= 8; r += 1) {
      for (let c = 36; c <= 50; c += 1) {
        const raw = plainExcelCell(worksheet.getCell(r, c)?.value);
        if (!monthNameOnly.test(raw)) continue;
        const left = plainExcelCell(worksheet.getCell(r, c - 1)?.value);
        if (!left || /^enter\b/i.test(left)) {
          worksheet.getCell(r, c - 1).value = 'Month';
          wroteMonth = true;
        }
      }
    }
  }

  return wroteMonth || wroteYear;
}

export function resolveFormXXVITamilNaduDownloadContext(
  formHeader,
  rowItem,
  fileName,
  tableHeaders = '',
  sheetText = ''
) {
  if (isFormXXVITamilNaduClraContext(formHeader, rowItem, fileName, tableHeaders, sheetText)) {
    return true;
  }
  const parts = buildFormXXVIContextBlob(formHeader, rowItem, fileName, tableHeaders, sheetText);
  if (!/register\s+of\s+employment\s+of\s+contractual\s*labo/i.test(parts)) return false;
  if (!/tamil\s*nadu|rule\s*75|daily\s+hours\s+of\s+work|form_xxvi|form_26.*tamil/i.test(parts)) {
    return false;
  }
  if (/register\s+of\s+accident/i.test(parts) && !/contractual\s*labo/i.test(parts)) return false;
  return true;
}

const MONTH_NAME_TO_INDEX = {
  january: 0,
  february: 1,
  march: 2,
  april: 3,
  may: 4,
  june: 5,
  july: 6,
  august: 7,
  september: 8,
  october: 9,
  november: 10,
  december: 11
};

function parseFormXXVITamilNaduJoinDate(raw) {
  if (raw instanceof Date && !Number.isNaN(raw.getTime())) {
    return new Date(raw.getFullYear(), raw.getMonth(), raw.getDate());
  }
  const t = String(raw ?? '').trim();
  if (!t) return null;
  if (/^\d+(\.\d+)?$/.test(t)) {
    const serial = Number(t);
    if (serial > 20000 && serial < 80000) {
      const utcMs = Math.round((serial - 25569) * 86400 * 1000);
      const d = new Date(utcMs);
      if (!Number.isNaN(d.getTime())) {
        return new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate());
      }
    }
  }
  const direct = new Date(t);
  if (!Number.isNaN(direct.getTime()) && direct.getFullYear() >= 1970) {
    return new Date(direct.getFullYear(), direct.getMonth(), direct.getDate());
  }
  const dmy = t.match(/(\d{1,2})[\s./-]+([a-z]{3,9})[\s./-]+(\d{4})/i);
  if (dmy) {
    const monKey = dmy[2].toLowerCase().slice(0, 3);
    const monFull = Object.keys(MONTH_NAME_TO_INDEX).find((k) => k.startsWith(monKey));
    const mi = monFull != null ? MONTH_NAME_TO_INDEX[monFull] : -1;
    if (mi >= 0) return new Date(Number(dmy[3]), mi, Number(dmy[1]));
  }
  const iso = t.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) {
    return new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3]));
  }
  const dmyNum = t.match(/^(\d{1,2})[\s./-]+(\d{1,2})[\s./-]+(\d{2,4})$/);
  if (dmyNum) {
    const dd = Number(dmyNum[1]);
    const mm = Number(dmyNum[2]) - 1;
    const yyyy = Number(dmyNum[3].length === 2 ? `20${dmyNum[3]}` : dmyNum[3]);
    const d = new Date(yyyy, mm, dd);
    if (!Number.isNaN(d.getTime())) return d;
  }
  return null;
}

function coerceFormXXVIEmployeeDateRaw(value) {
  if (value == null || value === '') return '';
  if (typeof value === 'object' && !Array.isArray(value)) {
    const nested =
      value.display_value ??
      value.displayValue ??
      value.date ??
      value.Date ??
      value.value ??
      '';
    return nested != null && String(nested).trim() !== '' ? String(nested).trim() : '';
  }
  return String(value).trim();
}

function getFormXXVITamilNaduEmployeeJoinRaw(emp) {
  const e = emp && typeof emp === 'object' ? emp : {};
  const keys = [
    'Dateofjoining',
    'Date of Joining',
    'DateofJoining',
    'dateOfJoining',
    'Date of joining',
    'Date_of_Joining',
    'Date_of_joining'
  ];
  for (let i = 0; i < keys.length; i += 1) {
    const v = coerceFormXXVIEmployeeDateRaw(e[keys[i]]);
    if (v) return v;
  }
  return '';
}

/** Same precedence as Autofill UI: Zoho DOJ first, then row Date of Entry into Service. */
export function resolveFormXXVITamilNaduEmployeeJoinDate(emp, row, headers) {
  const fromEmp = parseFormXXVITamilNaduJoinDate(getFormXXVITamilNaduEmployeeJoinRaw(emp));
  if (fromEmp) return fromEmp;
  const joinHeaderCandidates = [
    'Date of Entry into Service',
    'Date of entry into the service',
    'Date of Entry Into Service'
  ];
  if (row && typeof row === 'object') {
    for (let i = 0; i < joinHeaderCandidates.length; i += 1) {
      const d = parseFormXXVITamilNaduJoinDate(readFormXXVITamilNaduCellValue(row, joinHeaderCandidates[i]));
      if (d) return d;
    }
    const hit = (headers || []).find((h) =>
      /date\s+of\s+entry\s+into\s+(?:the\s+)?service/i.test(String(h || ''))
    );
    if (hit) {
      const d = parseFormXXVITamilNaduJoinDate(readFormXXVITamilNaduCellValue(row, hit));
      if (d) return d;
    }
  }
  return null;
}

function unwrapFormXXVITamilNaduExportEmployee(empItem) {
  if (!empItem || typeof empItem !== 'object') return null;
  return empItem.Employee || empItem.employee || empItem;
}

function resolveFormXXVITamilNaduPayrollMonthParts(selectedMonthIso, headerFormData = {}) {
  const iso = String(selectedMonthIso || '').trim();
  const isoMatch = iso.match(/^(\d{4})-(\d{2})/);
  if (isoMatch) {
    return { year: Number(isoMatch[1]), monthIndex: Number(isoMatch[2]) - 1 };
  }
  const monthName = String(headerFormData.form_x_month || '').trim().toLowerCase();
  const yearNum = Number(String(headerFormData.form_x_year || '').trim());
  if (monthName && MONTH_NAME_TO_INDEX[monthName] != null && Number.isFinite(yearNum)) {
    return { year: yearNum, monthIndex: MONTH_NAME_TO_INDEX[monthName] };
  }
  const now = new Date();
  return { year: now.getFullYear(), monthIndex: now.getMonth() };
}

/** Clear day-band marks before join date (matches Autofill UI — employee DOJ when present). */
export function scrubFormXXVITamilNaduPreJoinDayMarks(
  rows,
  headers,
  selectedMonthIso,
  headerFormData,
  options = {}
) {
  if (!Array.isArray(rows) || rows.length === 0) return;
  const dayHeaders = listFormXXVITamilNaduDayHeaders(headers);
  if (dayHeaders.length === 0) return;
  const { year, monthIndex } = resolveFormXXVITamilNaduPayrollMonthParts(selectedMonthIso, headerFormData);
  const employees = Array.isArray(options.employees) ? options.employees : [];
  const rowIndexOffset = Number(options.rowIndexOffset) || 0;
  rows.forEach((row, rowIndex) => {
    if (!row || typeof row !== 'object') return;
    const empItem = employees[rowIndex + rowIndexOffset] ?? employees[rowIndex];
    const emp = unwrapFormXXVITamilNaduExportEmployee(empItem);
    const joinDate = resolveFormXXVITamilNaduEmployeeJoinDate(emp, row, headers);
    if (!joinDate) return;
    dayHeaders.forEach(({ header, day }) => {
      const dayDate = new Date(year, monthIndex, day);
      if (dayDate.getMonth() !== monthIndex) return;
      if (dayDate.getTime() < joinDate.getTime()) row[header] = '';
    });
  });
}

/** First table column (Serial / Sr. No) — ignore modal tableStartCol when it shifts the day grid. */
export function resolveFormXXVITamilNaduExportTableStartCol(getCell, headerRow, maxCols, parsedTableStartCol) {
  for (let c = 1; c <= maxCols; c += 1) {
    const t = formXXVITamilNaduHeaderNorm(getCell(headerRow, c));
    if (/sr\.?\s*no|serial|s\.?\s*no/.test(t)) return c;
  }
  for (let c = 1; c <= maxCols; c += 1) {
    const t = formXXVITamilNaduHeaderNorm(getCell(headerRow, c));
    if (t) return c;
  }
  if (parsedTableStartCol != null && Number(parsedTableStartCol) >= 0) {
    return Number(parsedTableStartCol) + 1;
  }
  return 1;
}

/** Day 1 of the attendance band sits immediately after Rate of Wages (col 9 → day 1 at col 10). */
export function resolveFormXXVITamilNaduExportDay1Column(getCell, headerRow, startCol, identityCount, maxCols) {
  const scanFrom = 1;
  for (let c = scanFrom; c <= maxCols; c += 1) {
    const t = formXXVITamilNaduHeaderNorm(getCell(headerRow, c));
    if (/rate\s+of\s+wages/.test(t)) return c + 1;
  }
  for (let c = scanFrom; c <= maxCols; c += 1) {
    const t = formXXVITamilNaduHeaderNorm(getCell(headerRow, c));
    if (/daily\s+hours\s+of\s+work/.test(t)) return c;
  }
  return Math.max(1, (startCol > 0 ? startCol : 1) + identityCount);
}

function rebuildFormXXVITamilNaduOrderedColsFromDayRow(
  day1Col,
  identityCount,
  dayHeaderCount,
  rightCount
) {
  const safeDay1 = Math.max(1, Number(day1Col) || 1);
  const prefixStart = Math.max(1, safeDay1 - identityCount);
  const cols = [];
  for (let j = 0; j < identityCount; j += 1) cols.push(prefixStart + j);
  const dayCount = Math.min(Math.max(dayHeaderCount, 31), 31);
  for (let d = 0; d < dayCount; d += 1) cols.push(safeDay1 + d);
  const rightStart = safeDay1 + dayCount;
  for (let r = 0; r < rightCount; r += 1) cols.push(rightStart + r);
  return cols.filter((c) => Number.isFinite(c) && c >= 1);
}

function clearFormXXVITamilNaduPreJoinOnWorksheet(
  worksheet,
  {
    dataStartRow,
    rowCount,
    orderedCols,
    identityCount,
    day1Col,
    sourceRows,
    sourceObjects,
    selectedMonthIso,
    headerFormData,
    exportHeaders,
    employees,
    rowIndexOffset = 0
  }
) {
  if (!worksheet || rowCount < 1 || !orderedCols?.length) return;
  const { year, monthIndex } = resolveFormXXVITamilNaduPayrollMonthParts(selectedMonthIso, headerFormData);
  const safeDay1 = Math.max(1, Number(day1Col) || orderedCols[identityCount] || 1);
  for (let i = 0; i < rowCount; i += 1) {
    const rowObj = sourceObjects?.[i];
    if (!rowObj || typeof rowObj !== 'object') continue;
    const empItem = Array.isArray(employees) ? employees[i + rowIndexOffset] ?? employees[i] : null;
    const emp = unwrapFormXXVITamilNaduExportEmployee(empItem);
    const joinDate = resolveFormXXVITamilNaduEmployeeJoinDate(emp, rowObj, exportHeaders);
    if (!joinDate) continue;
    for (let d = 1; d <= 31; d += 1) {
      const dayDate = new Date(year, monthIndex, d);
      if (dayDate.getMonth() !== monthIndex) continue;
      if (dayDate.getTime() >= joinDate.getTime()) continue;
      const col = safeDay1 + d - 1;
      worksheet.getCell(dataStartRow + i, col).value = '';
      const j = identityCount + d - 1;
      if (sourceRows?.[i]) sourceRows[i][j] = '';
    }
  }
}

function writeFormXXVITamilNaduDayBandToWorksheet(
  worksheet,
  { dataStartRow, rowCount, day1Col, identityCount, sourceRows }
) {
  if (!worksheet || rowCount < 1) return;
  const safeDay1 = Math.max(1, Number(day1Col) || 1);
  for (let i = 0; i < rowCount; i += 1) {
    const row = sourceRows[i] || [];
    for (let d = 1; d <= 31; d += 1) {
      const j = identityCount + d - 1;
      const value = row[j];
      const cell = worksheet.getCell(dataStartRow + i, safeDay1 + d - 1);
      if (value == null || value === '') {
        cell.value = '';
      } else {
        cell.value = String(value);
      }
    }
  }
}

/** Stamp Month/Date and employer headers onto a draft workbook before PDF conversion. */
export async function stampFormXXVITamilNaduWorkbookHeaderForPdf(
  arrayBuffer,
  { headerFormData, parsedFormHeader, sheetNameHint } = {}
) {
  if (!arrayBuffer || arrayBuffer.byteLength < 32) return arrayBuffer;
  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(arrayBuffer);
    const worksheet = resolveWorksheet(workbook, sheetNameHint);
    if (!worksheet) return arrayBuffer;
    const headerValues =
      headerFormData && typeof headerFormData === 'object' ? { ...headerFormData } : {};
    writeStatutoryHeaderFieldsToExcelJsWorksheet(worksheet, {
      headerFormData: headerValues,
      parsedFormHeader,
      headerRowStart: 1,
      headerRowEnd: 35,
      maxScanCols: 120,
      writeMode: 'both'
    });
    writeFormXXVITamilNaduMonthYearToExcelJsWorksheet(worksheet, headerValues);
    return await workbook.xlsx.writeBuffer();
  } catch (err) {
    console.warn('Form XXVI TN PDF header stamp skipped:', err);
    return arrayBuffer;
  }
}

export function isFormXXVITamilNaduColumnIndexRow(getCell, row, maxCols = 80) {
  let nums = 0;
  let total = 0;
  for (let c = 1; c <= maxCols; c += 1) {
    const raw = String(getCell(row, c) || '')
      .trim()
      .replace(/[()]/g, '');
    if (!raw) continue;
    total += 1;
    if (/^\d{1,2}$/.test(raw)) nums += 1;
  }
  return total >= 8 && nums / total >= 0.75;
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

  const scanRows = [];
  for (const offset of [2, 3, 1, 4, 5, 0]) {
    const r = headerRow + offset;
    if (r >= 1 && !scanRows.includes(r)) scanRows.push(r);
  }

  let bestMap = new Map();
  let bestScore = -1;
  for (let ri = 0; ri < scanRows.length; ri += 1) {
    const r = scanRows[ri];
    const cols = new Map();
    for (let c = scanFrom; c < scanTo; c += 1) {
      const raw = String(getCell(r, c) || '')
        .trim()
        .replace(/[()]/g, '');
      if (!/^\d{1,2}$/.test(raw)) continue;
      const n = Number(raw);
      if (n >= 1 && n <= 31) cols.set(n, c);
    }
    let score = scoreMap(cols);
    // Prefer the dedicated 1–31 day band row (headerRow + 2) over 1–14 index rows.
    if (ri === 0 && score > 0) score += 500;
    if (score > bestScore) {
      bestScore = score;
      bestMap = cols;
    }
  }
  if (bestScore < 50 || bestMap.size < 5) return dayCols;
  if (!bestMap.has(1) && !bestMap.has(2)) return dayCols;
  // Reject maps that still look like identity column numbers (1)–(9)/(1)–(14).
  const day1Col = bestMap.get(1) ?? bestMap.get(2);
  if (bestMap.size <= 16 && day1Col != null && day1Col <= scanFrom + 1) return dayCols;
  let run = 0;
  for (let d = 1; d <= 31; d += 1) {
    if (!bestMap.has(d)) break;
    if (d > 1 && bestMap.get(d) !== bestMap.get(d - 1) + 1) break;
    run += 1;
  }
  if (run < 10) return dayCols;
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
  sheetNameHint,
  selectedMonthIso = '',
  employeesForExport = [],
  employeeRowIndexOffset = 0
}) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(templateArrayBuffer);
  const worksheet = resolveWorksheet(workbook, sheetNameHint);
  if (!worksheet) throw new Error('Template worksheet not found.');

  const normalize = formXXVITamilNaduHeaderNorm;
  const maxScanRows = Math.max(120, worksheet.rowCount + 10);
  const maxScanCols = 300;

  const getCellEarly = (r, c) => excelCellValueToString(worksheet.getCell(r, c)?.value);
  let headerRow = resolveFormXXVITamilNaduTableTextHeaderRow(
    getCellEarly,
    parsedHeaderRowIndex,
    maxScanRows,
    maxScanCols
  );
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

  const startCol = resolveFormXXVITamilNaduExportTableStartCol(
    getCellEarly,
    headerRow,
    maxScanCols,
    parsedTableStartCol
  );

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

  // Canonical export layout: 9 identity cols | days 1–31 | right cols (11–14).
  // Never let a short parsed prefix (e.g. only Serial/Name/Age) shift Age into Rate of Wages
  // or leave Daily hours / Number of Days Worked blank on download.
  const sourceHeaders = ensureFormXXVITamilNaduDayColumnHeaders(
    stripFormXXVITamilNaduNonTableHeaders(Array.isArray(headersToUse) ? headersToUse : [])
  );
  const identityCount = FORM_XXVI_TN_LEFT_HEADERS.length;
  const dayHeaderEntries = listFormXXVITamilNaduDayHeaders(sourceHeaders).filter(
    ({ header }) =>
      !isFormXXVITamilNaduRateOfWagesHeaderKey(header) &&
      !isFormXXVITamilNaduNumberOfDaysWorkedHeader(header)
  );
  const dayHeaders =
    dayHeaderEntries.length >= 5
      ? dayHeaderEntries
          .sort((a, b) => a.day - b.day)
          .map(({ header }) => header)
      : Array.from({ length: Math.min(markerCols.length || 31, 31) }, (_, i) => `10_${i + 1}`);
  const isLeftIdentityHeader = (h) => {
    const bucket = headerAliasBucket(h);
    if (bucket && FORM_XXVI_TN_LEFT_HEADERS.some((left) => headerAliasBucket(left) === bucket)) {
      return true;
    }
    return FORM_XXVI_TN_LEFT_HEADERS.some(
      (left) => formXXVITamilNaduHeaderNorm(left) === formXXVITamilNaduHeaderNorm(h)
    );
  };
  const rightHeaders = sourceHeaders.filter(
    (h) =>
      !isFormXXVITamilNaduRateOfWagesHeaderKey(h) &&
      !isFormXXVITamilNaduDayHeaderKey(h) &&
      !isLeftIdentityHeader(h) &&
      (isFormXXVITamilNaduNumberOfDaysWorkedHeader(h) ||
        /signature|thumb|termination|contractor|representative/i.test(String(h || '')))
  );
  const exportRightHeaders =
    rightHeaders.length > 0
      ? rightHeaders
      : FORM_XXVI_TN_RIGHT_HEADERS.slice(0, Math.min(4, FORM_XXVI_TN_RIGHT_HEADERS.length));
  // Prefer canonical left labels so alias lookup fills Age / Rate of Wages / etc. from row keys.
  const exportLeftHeaders = FORM_XXVI_TN_LEFT_HEADERS.map((canonical) => {
    const hit = sourceHeaders.find((h) => {
      const bucket = headerAliasBucket(canonical);
      return bucket && headerAliasBucket(h) === bucket;
    });
    return hit || canonical;
  });
  const hdrs = [...exportLeftHeaders, ...dayHeaders, ...exportRightHeaders];

  const templateDay1Col = resolveFormXXVITamilNaduExportDay1Column(
    getCell,
    headerRow,
    startCol,
    identityCount,
    maxScanCols
  );
  const firstDayCol =
    templateDay1Col > 0
      ? templateDay1Col
      : markerCols.length > 0
        ? markerCols[0]
        : Math.max(1, startCol + identityCount);

  const buildOrderedCols = () => {
    if (markerCols.length >= 5 || templateDay1Col > 0) {
      const cols = [];
      const expectedDay1Col = firstDayCol;
      let prefixStart = expectedDay1Col - identityCount;
      if (prefixStart < 1) prefixStart = Math.max(1, startCol);
      for (let j = 0; j < identityCount; j += 1) {
        cols.push(prefixStart + j);
      }
      const dayCount = Math.min(Math.max(dayHeaders.length, 31), 31);
      for (let d = 0; d < dayCount; d += 1) {
        cols.push(expectedDay1Col + d);
      }
      const rightCount = exportRightHeaders.length;
      if (rightCount > 0) {
        const rightStart = expectedDay1Col + dayCount;
        for (let r = 0; r < rightCount; r += 1) cols.push(rightStart + r);
      }
      return cols.filter((c) => Number.isFinite(c) && c >= 1 && c <= 16384);
    }
    return Array.from({ length: Math.max(12, hdrs.length) }, (_, i) => Math.max(1, startCol + i));
  };
  let orderedCols = buildOrderedCols();

  const applyFormXXVITamilNaduTemplateHeaderBand = () => {
    const indexRow = headerRow + 1;
    const dayMarkRow = headerRow + 2;
    const identityCountLocal = exportLeftHeaders.length;
    const dayCountLocal = dayHeaders.length;
    const rightCountLocal = exportRightHeaders.length;
    const tableClearTo = Math.max(
      maxScanCols,
      orderedCols.length > 0 ? Math.max(...orderedCols) + 4 : 50
    );
    for (let c = 1; c <= tableClearTo; c += 1) {
      worksheet.getCell(indexRow, c).value = null;
      worksheet.getCell(dayMarkRow, c).value = null;
    }
    for (let j = 0; j < hdrs.length; j += 1) {
      const col = orderedCols[j];
      if (!col) continue;
      const label = String(hdrs[j] || '').trim();
      if (!label) continue;
      const cell = worksheet.getCell(headerRow, col);
      if (j === identityCountLocal) {
        cell.value = 'Daily hours of work';
      } else if (j > identityCountLocal && j < identityCountLocal + dayCountLocal) {
        cell.value = '';
      } else {
        cell.value = label;
      }
    }
    if (dayCountLocal > 1) {
      const dayColStart = orderedCols[identityCountLocal];
      const dayColEnd = orderedCols[identityCountLocal + dayCountLocal - 1];
      if (dayColStart && dayColEnd && dayColEnd > dayColStart) {
        try {
          worksheet.mergeCells(headerRow, dayColStart, headerRow, dayColEnd);
        } catch (_) {
          /* template may already merge */
        }
      }
    }
    for (let j = 0; j < identityCountLocal; j += 1) {
      if (orderedCols[j]) worksheet.getCell(indexRow, orderedCols[j]).value = j + 1;
    }
    const firstDayJ = identityCountLocal;
    if (orderedCols[firstDayJ]) worksheet.getCell(indexRow, orderedCols[firstDayJ]).value = 10;
    for (let j = firstDayJ + 1; j < firstDayJ + dayCountLocal; j += 1) {
      if (orderedCols[j]) worksheet.getCell(indexRow, orderedCols[j]).value = null;
    }
    if (dayCountLocal > 1) {
      const dayColStart = orderedCols[firstDayJ];
      const dayColEnd = orderedCols[firstDayJ + dayCountLocal - 1];
      if (dayColStart && dayColEnd && dayColEnd > dayColStart) {
        try {
          worksheet.mergeCells(indexRow, dayColStart, indexRow, dayColEnd);
        } catch (_) {
          /* already merged */
        }
      }
    }
    for (let r = 0; r < rightCountLocal; r += 1) {
      const j = identityCountLocal + dayCountLocal + r;
      if (orderedCols[j]) worksheet.getCell(indexRow, orderedCols[j]).value = 11 + r;
    }
    const bandLeft =
      orderedCols.length > 0 ? Math.min(...orderedCols.filter(Boolean)) : Math.max(1, startCol);
    const bandRight =
      orderedCols.length > 0 ? Math.max(...orderedCols.filter(Boolean)) : bandLeft + hdrs.length;
    for (let c = bandLeft; c <= bandRight; c += 1) {
      worksheet.getCell(dayMarkRow, c).value = null;
    }
    for (let d = 0; d < dayCountLocal; d += 1) {
      const j = identityCountLocal + d;
      if (orderedCols[j]) worksheet.getCell(dayMarkRow, orderedCols[j]).value = d + 1;
    }
    for (let j = 0; j < orderedCols.length; j += 1) {
      const col = orderedCols[j];
      const h = formXXVITamilNaduHeaderNorm(hdrs[j] || '');
      let w = 10;
      if (j === 0) w = 4.5;
      else if (j < identityCountLocal) {
        if (/name of the work/.test(h)) w = 11;
        else if (/permanent/.test(h) && /address/.test(h)) w = 13;
        else if (/local address/.test(h)) w = 12;
        else if (/age/.test(h)) w = 7;
        else w = 9;
      } else if (j < identityCountLocal + dayCountLocal) w = 2.85;
      else w = 10;
      worksheet.getColumn(col).width = w;
    }
  };
  applyFormXXVITamilNaduTemplateHeaderBand();
  orderedCols = rebuildFormXXVITamilNaduOrderedColsFromDayRow(
    firstDayCol,
    exportLeftHeaders.length,
    dayHeaders.length,
    exportRightHeaders.length
  );
  dataStartRow = Math.max(dataStartRow, headerRow + 3);
  markerRow = headerRow + 2;

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

  const headerValues = headerFormData && typeof headerFormData === 'object' ? { ...headerFormData } : {};
  if (!String(headerValues.form_x_month ?? '').trim() && parsedFormHeader?.wagePeriodText) {
    const wp = String(parsedFormHeader.wagePeriodText || '').trim();
    const monthMatch = wp.match(
      /\b(january|february|march|april|may|june|july|august|september|october|november|december)\b/i
    );
    if (monthMatch) headerValues.form_x_month = monthMatch[1];
  }
  writeStatutoryHeaderFieldsToExcelJsWorksheet(worksheet, {
    headerFormData: headerValues,
    parsedFormHeader,
    headerRowStart: 1,
    headerRowEnd: Math.max(35, headerRow),
    maxScanCols: 120,
    writeMode: 'both'
  });
  writeFormXXVITamilNaduMonthYearToExcelJsWorksheet(worksheet, headerValues);

  const sourcePrimary =
    Array.isArray(mappedData) && mappedData.length > 0
      ? mappedData.map((row) =>
          row && typeof row === 'object' && !Array.isArray(row) ? { ...row } : row
        )
      : Array.isArray(mappedRowMatrix)
        ? mappedRowMatrix.map((row) =>
            row && typeof row === 'object' && !Array.isArray(row) ? { ...row } : row
          )
        : [];
  const exportHeadersForScrub = ensureFormXXVITamilNaduDayColumnHeaders(
    stripFormXXVITamilNaduNonTableHeaders(headersToUse || [])
  );
  scrubFormXXVITamilNaduPreJoinDayMarks(
    sourcePrimary,
    exportHeadersForScrub,
    selectedMonthIso,
    headerValues,
    { employees: employeesForExport, rowIndexOffset: employeeRowIndexOffset }
  );
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

  const meaningfulSourceRows = sourcePrimary.filter((row) => rowLooksMeaningful(row));
  const identityCountExport = exportLeftHeaders.length;
  const dayBandEndIndex = identityCountExport + 31;
  for (let i = 0; i < sourceRows.length; i += 1) {
    const row = sourceRows[i] || [];
    for (let j = 0; j < orderedCols.length; j += 1) {
      if (j >= identityCountExport && j < dayBandEndIndex) continue;
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
  writeFormXXVITamilNaduDayBandToWorksheet(worksheet, {
    dataStartRow,
    rowCount: sourceRows.length,
    day1Col: firstDayCol,
    identityCount: identityCountExport,
    sourceRows
  });

  clearFormXXVITamilNaduPreJoinOnWorksheet(worksheet, {
    dataStartRow,
    rowCount: sourceRows.length,
    orderedCols,
    identityCount: identityCountExport,
    day1Col: firstDayCol,
    sourceRows,
    sourceObjects: meaningfulSourceRows,
    selectedMonthIso,
    headerFormData: headerValues,
    exportHeaders: exportHeadersForScrub,
    employees: employeesForExport,
    rowIndexOffset: employeeRowIndexOffset
  });

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
    // Full table span for centered "System Generated" box below the grid.
    worksheet._formXxviTableColMin = tableColMin;
    worksheet._formXxviTableColMax = tableColMax;
  }

  writeFormXXVITamilNaduMonthYearToExcelJsWorksheet(worksheet, headerValues);

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
