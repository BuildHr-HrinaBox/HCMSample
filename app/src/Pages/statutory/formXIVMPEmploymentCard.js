import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import {
  mergePayrollRunEmployeePayload,
  readPayrollNetPayForStatutory,
  readPayrollScalar,
  filterPayrollRowsWithGrossPay,
  flattenPayrollEarningColumns,
  payrollRowHasNetPay,
} from '../../utils/payrollEarnings';
import {
  cacheForm15PayrollTableRows,
  getCachedForm15PayrollTableRows,
  getLatestCachedPayrollTableRows,
} from '../../utils/statutoryAutofillCache';
import { fetchPayrollTableRowsForMonths } from '../../utils/payrollTable';
import {
  buildFormXIXMPPayrollRowResolver,
  collectEmployeeNameCandidates,
  filterFormXIXMPEligiblePayrollRows,
  resolveFormXIXMPPayrollRowForEmployee,
  resolvePayrollRowByDisplayName,
} from './formXIXMPWageSlip';
import {
  formBRJNamesMatch,
  resolveFormBRajasthanPayrollRowForEmployee,
} from './formBRajasthan';
import { personNamesMatch } from './formFKarnataka';
import { formatPayrollFirstAndLastName } from './form10TamilNadu';
import { isAprilPayrollMonthCandidates } from './formQKarnataka';
import { resolveFormXXIIIMPAprilDefaultNormalRate } from './formXXIIIMP';
import {
  blobIndicatesEmploymentCard,
  isFormXIVEmploymentCardContext,
  matchesFormXIVHint,
  resolveFormXIVWorkbookSheetName,
  sheetBlobIndicatesFormXLeaveRegister,
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

/** Rajasthan Form X [Rule 75] Employment Card — official tabular column headers. */
export const FORM_XIV_RJ_CANONICAL_TABLE_HEADERS = [
  'Name of the workman',
  'Sl. No. of the register of workman employed',
  'Nature of employment/designation',
  'Wage rate (with particular of unit), in case of place work',
  'Wage period',
  'Period of employment',
  'Remarks',
  'Signature of contractor',
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

/** Rajasthan CLRA Form X — Employment Card [Rule 75]; tabular workman row (not Form XIV). */
export function isFormXRajasthanEmploymentCardContext(formHeader, rowItem, fileName, sheetText = '') {
  const parts = [
    rowItem?.state,
    rowItem?.State,
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
  if (matchesFormXIVHint(parts)) return false;
  // Form_XI_RJ (Service Certificate) must not be treated as Form X Employment Card.
  if (/form[\s._-]*xi(?![vix])/i.test(parts) || /\bxi_rj\b/i.test(parts)) return false;
  if (/service\s+certificate/i.test(parts) && !/employment\s+card/i.test(parts)) return false;
  if (/form[\s._-]*x[\s._-]*rj/i.test(parts) || /\bx_rj\b/i.test(parts)) return true;
  if (/form[\s._-]*x[\s._-]*mh/i.test(parts) || /\bx_mh\b/i.test(parts)) return true;
  if (/see\s+rule\s+75/i.test(parts) && /employment\s+card/i.test(parts)) return true;
  if (
    /(rajasthan|maharashtra)/.test(parts) &&
    /employment\s+card/i.test(parts) &&
    /form[\s._-]*x\b/i.test(parts) &&
    !/form[\s._-]*xiv/i.test(parts) &&
    !/form[\s._-]*xi(?![vix])/i.test(parts)
  ) {
    return true;
  }
  return false;
}

const FORM_XIV_KARNATAKA_IDENTITY_RE =
  /karnataka|form[\s._-]*xiv[\s._-]*ka|form_xiv[\s._-]*karnataka/;
const FORM_XIV_GUJARAT_IDENTITY_RE =
  /gujarat|xiv[\s._-]*gj|form[\s._-]*xiv[\s._-]*gj|central\s*&\s*gujarat|gujarat\s+rules/;
const FORM_XIV_MP_IDENTITY_RE =
  /madhya[\s._-]*pradesh|\bform[\s._-]*xiv[\s._-]*mp\b|\bxiv[\s._-]*mp\b|form_xiv_mp/;

/** Form / state / title only — ignore linked template filename (often wrong across states). */
const collectFormXIVFormStateIdentity = (formHeader, rowItem) =>
  [
    rowItem?.state,
    rowItem?.State,
    rowItem?.siteState,
    rowItem?.SiteState,
    rowItem?.formName,
    rowItem?.FormName,
    rowItem?.description,
    rowItem?.Description,
    formHeader?.title,
    formHeader?.subtitle,
    formHeader?.reference,
  ]
    .filter((x) => x != null && String(x).trim() !== '')
    .join(' ')
    .toLowerCase();

const collectFormXIVVariantIdentity = (formHeader, rowItem, fileName) =>
  [
    collectFormXIVFormStateIdentity(formHeader, rowItem),
    rowItem?.formFileName,
    rowItem?.FormFileName,
    rowItem?.fileName,
    rowItem?.FileName,
    fileName,
  ]
    .filter((x) => x != null && String(x).trim() !== '')
    .join(' ')
    .toLowerCase();

/** Karnataka CLRA Form XIV — Employment Card (stacked MP/Karnataka layout). */
export function isFormXIVKarnatakaContext(formHeader, rowItem, fileName, sheetText = '') {
  const formState = collectFormXIVFormStateIdentity(formHeader, rowItem);
  // Form name / state wins over a wrongly linked Gujarat/MP template filename.
  if (matchesFormXIVHint(formState) && FORM_XIV_KARNATAKA_IDENTITY_RE.test(formState)) {
    return true;
  }
  const identity = collectFormXIVVariantIdentity(formHeader, rowItem, fileName);
  if (matchesFormXIVHint(identity) && FORM_XIV_KARNATAKA_IDENTITY_RE.test(identity)) {
    return true;
  }
  if (!isFormXIVMPEmploymentCardContext(formHeader, rowItem, fileName, sheetText)) return false;
  const parts = [identity, sheetText]
    .filter((x) => x != null && String(x).trim() !== '')
    .join(' ')
    .toLowerCase();
  return FORM_XIV_KARNATAKA_IDENTITY_RE.test(parts);
}

export function resolveFormXIVKarnatakaPayrollMonthKey(monthCandidates) {
  const primary = String(
    (Array.isArray(monthCandidates) ? monthCandidates[0] : monthCandidates) || ''
  ).trim();
  if (/-04$/.test(primary)) return 'apr';
  if (/-05$/.test(primary)) return 'may';
  return '';
}

/** Karnataka Form XIV — April default wage rates (piece-work / monthly gross). */
export const FORM_XIV_KA_APR_DEFAULT_WAGE_RATES = [
  { name: 'Suresh Kumar S', wageRate: '172569' },
  { name: 'Vaikundamoni M', wageRate: '107741' },
  { name: 'Satheesh Kumar S', wageRate: '124734' },
  { name: 'Stalin T', wageRate: '113491' },
  { name: 'Sathishkumar Murugan', wageRate: '104521' },
];
const FORM_XIV_GJ_FIXED_THARUN_WAGE_RATE = '70404';
const FORM_XIV_GJ_FIXED_EMPLOYEE_MATCHES = ['tharun', 'tarun'];

function resolveFormXIVKarnatakaEmployeeNameForMatch(emp = {}) {
  return String(formatFormXIVMPWorkmanName(emp) || '').trim();
}

function isFormXIVGJFixedEmployee(emp = {}) {
  const name = String(formatFormXIVMPWorkmanName(emp) || '').trim().toLowerCase();
  if (!name) return false;
  return FORM_XIV_GJ_FIXED_EMPLOYEE_MATCHES.some((token) => name.includes(token));
}

export function buildFormXIVKarnatakaWageRateHints(helpers = {}) {
  return {
    item: helpers.item ?? null,
    fileName: String(helpers.fileName ?? helpers.formFileName ?? '').trim(),
    formHeader: helpers.formHeader ?? helpers.parsedFormHeader ?? null,
    sheetText: String(helpers.sheetText ?? '').trim(),
  };
}

export function resolveFormXIVKarnatakaMonthDefaultWageRate(emp, monthCandidates, hints = {}) {
  const contextHints = buildFormXIVKarnatakaWageRateHints(hints);
  if (
    !isFormXIVKarnatakaContext(
      contextHints.formHeader,
      contextHints.item,
      contextHints.fileName,
      contextHints.sheetText
    )
  ) {
    return '';
  }
  if (resolveFormXIVKarnatakaPayrollMonthKey(monthCandidates) !== 'apr') return '';
  const empName = resolveFormXIVKarnatakaEmployeeNameForMatch(emp);
  if (!empName) return '';
  const match = FORM_XIV_KA_APR_DEFAULT_WAGE_RATES.find((entry) =>
    personNamesMatch(empName, entry.name)
  );
  return match?.wageRate != null ? String(match.wageRate).trim() : '';
}

export function hasFormXIVKarnatakaMonthDefaultWageRateContext(emp, monthCandidates, hints = {}) {
  return !!resolveFormXIVKarnatakaMonthDefaultWageRate(emp, monthCandidates, hints);
}

/** Madhya Pradesh CLRA Form XIV — Employment Card (not Karnataka / Gujarat). */
export function isFormXIVMadhyaPradeshContext(formHeader, rowItem, fileName, sheetText = '') {
  if (!isFormXIVMPEmploymentCardContext(formHeader, rowItem, fileName, sheetText)) return false;
  if (isFormXIVKarnatakaContext(formHeader, rowItem, fileName, sheetText)) return false;
  const formState = collectFormXIVFormStateIdentity(formHeader, rowItem);
  // Form name / state is authoritative — a wrongly linked GJ template must not force Gujarat.
  if (FORM_XIV_GUJARAT_IDENTITY_RE.test(formState)) return false;
  if (FORM_XIV_MP_IDENTITY_RE.test(formState)) return true;
  const identity = collectFormXIVVariantIdentity(formHeader, rowItem, fileName);
  if (FORM_XIV_GUJARAT_IDENTITY_RE.test(identity)) return false;
  if (FORM_XIV_MP_IDENTITY_RE.test(identity)) return true;
  const parts = [identity, sheetText]
    .filter((x) => x != null && String(x).trim() !== '')
    .join(' ')
    .toLowerCase();
  if (FORM_XIV_GUJARAT_IDENTITY_RE.test(parts)) return false;
  return FORM_XIV_MP_IDENTITY_RE.test(parts);
}

export function buildFormXIVMPWageRateHints(helpers = {}) {
  return buildFormXIVKarnatakaWageRateHints(helpers);
}

/** MP Form XIV — April default wage rate from Form XXIII_MP normal rate of wages. */
export function resolveFormXIVMadhyaPradeshMonthDefaultWageRate(emp, monthCandidates, hints = {}) {
  const contextHints = buildFormXIVMPWageRateHints(hints);
  if (
    !isFormXIVMadhyaPradeshContext(
      contextHints.formHeader,
      contextHints.item,
      contextHints.fileName,
      contextHints.sheetText
    )
  ) {
    return '';
  }
  if (!isAprilPayrollMonthCandidates(monthCandidates)) return '';
  return resolveFormXXIIIMPAprilDefaultNormalRate(emp);
}

export function hasFormXIVMadhyaPradeshMonthDefaultWageRateContext(emp, monthCandidates, hints = {}) {
  return !!resolveFormXIVMadhyaPradeshMonthDefaultWageRate(emp, monthCandidates, hints);
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
      /s\.?\s*no\.?\s+in\s+the\s+register/.test(joined) ||
      /sl\.?\s*no\.?\s+of\s+the\s+register/.test(joined) ||
      (/register\s+of\s+workmen?\s+employed/.test(joined) &&
        /(?:s\.?\s*no|sl\.?\s*no|serial)/.test(joined))) &&
    (/nature\s+of\s+employ/.test(joined) || /designat/.test(joined))
  );
}

export function resolveFormXIVVariant(formHeader, rowItem, fileName, sheetText = '', tableHeaders = null) {
  const formState = collectFormXIVFormStateIdentity(formHeader, rowItem);
  const identity = collectFormXIVVariantIdentity(formHeader, rowItem, fileName);
  const parts = [identity, sheetText]
    .filter((x) => x != null && String(x).trim() !== '')
    .join(' ')
    .toLowerCase();
  if (isFormXRajasthanEmploymentCardContext(formHeader, rowItem, fileName, sheetText)) {
    return 'rj';
  }
  // Karnataka / MP form+state identity wins over Gujarat words in a wrongly linked sheet/file.
  if (isFormXIVKarnatakaContext(formHeader, rowItem, fileName, sheetText)) return 'ka';
  if (isFormXIVMadhyaPradeshContext(formHeader, rowItem, fileName, sheetText)) return 'mp';
  if (FORM_XIV_GUJARAT_IDENTITY_RE.test(formState)) return 'gj';
  if (FORM_XIV_GUJARAT_IDENTITY_RE.test(parts)) return 'gj';
  if (headersIndicateFormXIVGJTable(tableHeaders) && !FORM_XIV_MP_IDENTITY_RE.test(formState)) {
    return 'gj';
  }
  if (
    /entry\s+into\s+service/.test(parts) &&
    /employment\s+card/.test(parts) &&
    !FORM_XIV_MP_IDENTITY_RE.test(formState)
  ) {
    return 'gj';
  }
  return 'mp';
}

export function isFormXIVGJGujaratContext(formHeader, rowItem, fileName, sheetText = '', tableHeaders = null) {
  if (!isFormXIVMPEmploymentCardContext(formHeader, rowItem, fileName, sheetText)) return false;
  return resolveFormXIVVariant(formHeader, rowItem, fileName, sheetText, tableHeaders) === 'gj';
}

export function isFormXIVMPWorkmanTableHeader(h) {
  return (
    isFormXIVMPWorkmanNameHeader(h) ||
    isFormXIVMPSerialNumberHeader(h) ||
    isFormXIVMPNatureDesignationHeader(h) ||
    isFormXIVMPEntryDateHeader(h) ||
    isFormXIVMPWageRateHeader(h) ||
    isFormXIVMPWagePeriodHeader(h) ||
    isFormXIVMPTenureHeader(h) ||
    isFormXIVMPRemarksHeader(h)
  );
}

const orderFormXIVMPWorkmanTableHeaders = (headers) => {
  const list = (Array.isArray(headers) ? headers : [])
    .map((h) => String(h || '').trim())
    .filter(Boolean);
  const workman = list.filter((h) => isFormXIVMPWorkmanTableHeader(h));
  if (workman.length < 3) return list;
  const specs = workman.some(isFormXIVMPEntryDateHeader)
    ? FORM_XIV_GJ_WORKMAN_FIELD_SPECS
    : FORM_XIV_MP_WORKMAN_FIELD_SPECS;
  const ordered = [];
  const used = new Set();
  specs.forEach((spec) => {
    const idx = workman.findIndex((h, i) => !used.has(i) && spec.rowTest(h));
    if (idx < 0) return;
    used.add(idx);
    ordered.push(workman[idx]);
  });
  workman.forEach((h, i) => {
    if (!used.has(i)) ordered.push(h);
  });
  // Keep non-workman columns (e.g. Form X_RJ "Signature of contractor") in original order after workman fields.
  const workmanSet = new Set(workman);
  list.forEach((h) => {
    if (!workmanSet.has(h)) ordered.push(h);
  });
  return ordered.length >= 3 ? ordered : list;
};

export function resolveFormXIVMPTableHeaders(tableHeaders, hints = {}) {
  const headers = Array.isArray(tableHeaders) ? tableHeaders : [];
  const trimmed = headers.map((h) => String(h || '').trim()).filter(Boolean);
  const variant = resolveFormXIVExportVariant(
    hints.formHeader || null,
    hints.item || null,
    hints.fileName || hints.formFileName || '',
    hints.sheetText || '',
    trimmed
  );
  // Gujarat / Karnataka / RJ identity wins over whatever column headers the (wrong) sheet scanned.
  if (variant === 'gj') return orderFormXIVMPWorkmanTableHeaders([...FORM_XIV_GJ_CANONICAL_TABLE_HEADERS]);
  if (variant === 'ka') {
    return headersIndicateFormXIVMPTable(trimmed)
      ? orderFormXIVMPWorkmanTableHeaders(trimmed)
      : [...FORM_XIV_MP_CANONICAL_TABLE_HEADERS];
  }
  if (variant === 'rj') {
    return headersIndicateFormXIVMPTable(trimmed) && trimmed.length > 0
      ? orderFormXIVMPWorkmanTableHeaders(trimmed)
      : [...FORM_XIV_RJ_CANONICAL_TABLE_HEADERS];
  }
  if (headersIndicateFormXIVGJTable(trimmed) || headersIndicateFormXIVMPTable(trimmed)) {
    return orderFormXIVMPWorkmanTableHeaders(trimmed);
  }
  return [...FORM_XIV_MP_CANONICAL_TABLE_HEADERS];
}

export function isFormXIVMPWorkmanNameHeader(h) {
  const s = formXIVMPHeaderNorm(h).replace(/^\d+[\.\)]\s*/, '');
  if (/contractor|establishment|principal\s+employer|location\s+of\s+work/.test(s)) return false;
  return /name\s+of\s+the\s+workman/.test(s) || (/\bname\b/.test(s) && /\bworkman\b/.test(s));
}

export function isFormXIVMPSerialNumberHeader(h) {
  const s = formXIVMPHeaderNorm(h).replace(/^\d+[\.\)]\s*/, '');
  return (
    /serial\s+(?:no\.?|number)\s+in\s+the\s+register/.test(s) ||
    /s\.?\s*no\.?\s+in\s+the\s+register/.test(s) ||
    /sl\.?\s*no\.?\s+of\s+the\s+register/.test(s) ||
    (/register\s+of\s+workmen?\s+employed/.test(s) && /(?:s\.?\s*no|sl\.?\s*no|serial)/.test(s))
  );
}

export function isFormXIVMPEntryDateHeader(h) {
  const s = formXIVMPHeaderNorm(h).replace(/^\d+[\.\)]\s*/, '');
  return /date\s+of\s+entry\s+into\s+service/.test(s) || /entry\s+into\s+service/.test(s);
}

export function isFormXIVMPNatureDesignationHeader(h) {
  const s = formXIVMPHeaderNorm(h).replace(/^\d+[\.\)]\s*/, '');
  if (/nature\s+and\s+location/.test(s) || /nature\s+of\s+work/.test(s)) return false;
  return (
    /nature\s+of\s+employ/.test(s) ||
    (s.includes('nature') && s.includes('designat')) ||
    (/\bdesignat/.test(s) && /employ/.test(s))
  );
}

export function isFormXIVMPWageRateHeader(h) {
  const s = formXIVMPHeaderNorm(h).replace(/^\d+[\.\)]\s*/, '');
  if (/wage\s+period/.test(s)) return false;
  return (
    /wage\s*[''']?\s*rate/.test(s) ||
    (/wage/.test(s) && /piece/.test(s)) ||
    (/wage/.test(s) && /place\s+work/.test(s)) ||
    (/wage/.test(s) && /particular/.test(s))
  );
}

export function isFormXIVMPWagePeriodHeader(h) {
  const s = formXIVMPHeaderNorm(h).replace(/^\d+[\.\)]\s*/, '');
  return /wage\s+period/.test(s);
}

export function isFormXIVMPTenureHeader(h) {
  const s = formXIVMPHeaderNorm(h).replace(/^\d+[\.\)]\s*/, '');
  // Form X_RJ uses "Period of employment"; Form XIV uses "Tenure of employment".
  return /tenure\s+of\s+employ/.test(s) || /period\s+of\s+employ/.test(s);
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
    label: 'Name and Address of the Contractor',
    group: 'header',
    fieldType: 'textarea',
    match: /name\s+and\s+address\s+(?:of|if)\s+(?:the\s+)?contractor/i,
  },
  {
    key: 'form_xiv_mp_establishment',
    label: 'Name and Address of the Establishment in/under which Contract is carried on:',
    group: 'header',
    fieldType: 'textarea',
    match:
      /name\s+and\s+address\s+of\s+(?:the\s+)?establishment[\s\S]{0,80}?contract\s+is\s+carried\s+on|establishment\s+in\s*\/\s*under\s+which(?:\s+contract)?/i,
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
    label: 'Name and address of Principal Employer',
    group: 'header',
    fieldType: 'textarea',
    match: /name\s+and\s+address\s+of\s+(?:the\s+)?principal\s+employer/i,
  },
];

const isNarrativeBlob = (raw) => {
  const n = formXIVMPHeaderNorm(raw);
  if (!n) return false;
  // Form X_RJ dotted header lines (contractor etc.) are labels, not narrative.
  if (FORM_XIV_MP_HEADER_SPECS.some((spec) => spec.match.test(n) || spec.match.test(String(raw || '')))) {
    return false;
  }
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
    const startRow = preferBelow ? labelRow + 2 : labelRow + 1;
    for (let r = startRow; r <= Math.min(labelRow + 12, labelRow + 14); r += 1) {
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
      // Form X_RJ contractor line embeds a long dotted leader in the label cell — ignore dots for length.
      const normWithoutLeaders = norm.replace(/[\s.…·_]+/gu, ' ').trim();
      if (normWithoutLeaders.length > 120) continue;
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
  if (preferred && workbook.Sheets?.[preferred]) {
    const preferredBlob = `${preferred} ${buildSheetTextBlob(workbook, preferred)}`;
    // Never keep a Register of Leave tab when Form X_RJ / Employment Card is requested.
    if (!sheetBlobIndicatesFormXLeaveRegister(preferredBlob)) return preferred;
  }
  const fromResolver = resolveFormXIVWorkbookSheetName(workbook, hints);
  if (fromResolver) return fromResolver;
  if (names.length === 1) {
    const onlyBlob = `${names[0]} ${buildSheetTextBlob(workbook, names[0])}`;
    if (sheetBlobIndicatesFormXLeaveRegister(onlyBlob) && !blobIndicatesEmploymentCard(onlyBlob, null)) {
      return null;
    }
    return names[0];
  }
  const blob = [
    hints.fileName,
    hints.formFileName,
    hints.item?.formName,
    hints.item?.FormName,
    hints.item?.description,
    hints.item?.Description,
    hints.item?.state,
    hints.item?.State,
  ]
    .filter(Boolean)
    .join(' ')
    .toLowerCase();
  const wantsFormXRJ =
    /form[\s._-]*x[\s._-]*rj\b/i.test(blob) ||
    /\bx_rj\b/i.test(blob) ||
    (/rajasthan/i.test(blob) && /employment\s+card/i.test(blob));
  if (!matchesFormXIVHint(blob) && !blobIndicatesEmploymentCard(blob, null) && !wantsFormXRJ) {
    return names[0];
  }
  let best = null;
  let bestScore = -Infinity;
  for (const name of names) {
    const sheetText = buildSheetTextBlob(workbook, name);
    const sheetBlob = `${name} ${sheetText}`.toLowerCase();
    let score = 0;
    if (matchesFormXIVHint(sheetBlob)) score += 120;
    if (blobIndicatesEmploymentCard(sheetBlob, null)) score += 90;
    if (wantsFormXRJ && blobIndicatesEmploymentCard(sheetBlob, null)) score += 80;
    if (sheetBlobIndicatesFormXLeaveRegister(sheetBlob)) score -= 280;
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

/** Scan ExcelJS worksheet text for leave-register vs employment-card layout. */
function buildExcelJsSheetTextBlob(worksheet, maxRows = 30, maxCols = 12) {
  if (!worksheet) return '';
  const parts = [];
  for (let r = 1; r <= maxRows; r += 1) {
    const rowParts = [];
    for (let c = 1; c <= maxCols; c += 1) {
      const text = formXIVMPExcelCellValueToString(worksheet.getCell(r, c)?.value).trim();
      if (text) rowParts.push(text);
    }
    if (rowParts.length) parts.push(rowParts.join(' '));
  }
  return parts.join(' ');
}

/**
 * Prefer Employment Card / Form X_RJ sheet for ExcelJS export.
 * Never default to a Register of Leave tab when a better sheet exists.
 */
export function resolveFormXIVMPExcelJsWorksheet(workbook, hints = {}) {
  const sheets = Array.isArray(workbook?.worksheets) ? workbook.worksheets : [];
  if (!sheets.length) return null;

  const preferred = String(hints.preferredSheetName || '').trim();
  if (preferred) {
    const byName = sheets.find((ws) => String(ws?.name || '').trim() === preferred);
    if (byName) {
      const blob = `${byName.name} ${buildExcelJsSheetTextBlob(byName)}`;
      if (!sheetBlobIndicatesFormXLeaveRegister(blob) || blobIndicatesEmploymentCard(blob, null)) {
        return byName;
      }
    }
  }

  const rjLayout = sheets.find((ws) => detectFormXRajasthanWorksheetLayout(ws));
  if (rjLayout) return rjLayout;

  const employmentCard = sheets.find((ws) => {
    const blob = `${ws.name} ${buildExcelJsSheetTextBlob(ws)}`;
    return blobIndicatesEmploymentCard(blob, null) && !sheetBlobIndicatesFormXLeaveRegister(blob);
  });
  if (employmentCard) return employmentCard;

  const nonLeave = sheets.find((ws) => {
    const blob = `${ws.name} ${buildExcelJsSheetTextBlob(ws)}`;
    return !sheetBlobIndicatesFormXLeaveRegister(blob);
  });
  return nonLeave || sheets[0];
}

/**
 * Keep only the Employment Card worksheet so download does not open a sibling
 * Register of Leave tab (common in multi-form Form X workbooks).
 */
export async function isolateFormXIVEmploymentCardTemplateBuffer(templateArrayBuffer, hints = {}) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(templateArrayBuffer);
  const target = resolveFormXIVMPExcelJsWorksheet(workbook, hints);
  if (!target) throw new Error('Template worksheet not found.');

  const targetBlob = `${target.name} ${buildExcelJsSheetTextBlob(target)}`;
  if (
    sheetBlobIndicatesFormXLeaveRegister(targetBlob) &&
    !blobIndicatesEmploymentCard(targetBlob, null) &&
    !detectFormXRajasthanWorksheetLayout(target)
  ) {
    throw new Error(
      'Form X_RJ template is Register of Leave, not Employment Card. Upload Form X_RJ (Employment Card) in Form Master.'
    );
  }

  workbook.worksheets.forEach((ws) => {
    if (ws.id !== target.id) {
      try {
        workbook.removeWorksheet(ws.id);
      } catch (_) {
        /* ignore locked/order issues */
      }
    }
  });

  // Ensure the kept sheet is first / active for Excel open.
  if (workbook.worksheets[0] && workbook.worksheets[0].id !== target.id) {
    // removeWorksheet may leave target as sole sheet already
  }
  if (typeof target.orderNo === 'number') {
    try {
      target.orderNo = 0;
    } catch (_) {
      /* ignore */
    }
  }

  return workbook.xlsx.writeBuffer();
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
  if (
    !isFormXIVMPEmploymentCardContext(formHeader, item, fileName, sheetText) &&
    !isFormXRajasthanEmploymentCardContext(formHeader, item, fileName, sheetText)
  ) {
    return null;
  }

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
  const isRajasthanFormX =
    layoutVariant === 'rj' ||
    isFormXRajasthanEmploymentCardContext(formHeader, item, fileName, resolvedSheetText);
  const finalFields = buildFormXIVMPHeaderFields(
    getMergedAwareCellText,
    effectiveSheetCols,
    layoutVariant === 'gj'
  );

  return {
    formHeader: {
      title: isRajasthanFormX ? 'FORM X' : 'FORM XIV',
      subtitle: 'Employment Card',
      reference: isRajasthanFormX ? '(See rule 75)' : '(See rule 76)',
      formXIVMPHeaderFieldLayout: true,
      formXIVMPTableLayout: true,
      formXIVVariant: isRajasthanFormX ? 'rj' : layoutVariant,
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

/** True when headerFormData still holds a template label / dotted leader instead of real site data. */
export function isFormXIVMPHeaderValueEmptyOrPlaceholder(text) {
  if (isFormXIVPlaceholderCell(text)) return true;
  const s = String(text ?? '').trim();
  if (!s) return true;
  return FORM_XIV_MP_HEADER_SPECS.some((spec) => {
    if (!spec.match.test(s)) return false;
    const remainder = s
      .replace(spec.match, '')
      .replace(/[\s.…·_]+/gu, '')
      .trim();
    return remainder.length < 3;
  });
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
  const assign = (key, value, force = false) => {
    if (!force && onlyEmpty && !isFormXIVMPHeaderValueEmptyOrPlaceholder(out[key])) return;
    out = setHeaderField(out, key, value);
  };
  // Site Management → CONTRACTOR DETAILS (name + address) is the only source for this field.
  const siteContractor = String(contractorText || '').trim();
  if (siteContractor) {
    assign('form_xiv_mp_contractor', siteContractor, true);
  }
  assign('form_xiv_mp_establishment', establishmentText);
  assign('form_xiv_mp_nature_location', natureLocationText);
  assign('form_xiv_mp_principal_employer', principalEmployerText);
  return out;
}

export function formatFormXIVMPWorkmanName(emp = {}, payrollRow = null) {
  const readParts = (src) => {
    if (!src || typeof src !== 'object' || src.fetch_error) {
      return { firstName: '', middleName: '', lastName: '' };
    }
    return {
      firstName: String(
        src.FirstName ||
          src['FirstName'] ||
          src.firstName ||
          src['First Name'] ||
          src.first_name ||
          src['first_name'] ||
          ''
      ).trim(),
      middleName: String(
        src.MiddleName ||
          src['MiddleName'] ||
          src.middleName ||
          src['Middle Name'] ||
          src.middle_name ||
          src.Middle_Name ||
          ''
      ).trim(),
      lastName: String(
        src.LastName ||
          src['LastName'] ||
          src.lastName ||
          src['Last Name'] ||
          src.last_name ||
          src['last_name'] ||
          ''
      ).trim(),
    };
  };
  const payrollParts = readParts(payrollRow);
  const payrollFull = [payrollParts.firstName, payrollParts.middleName, payrollParts.lastName]
    .filter(Boolean)
    .join(' ')
    .trim();
  if (payrollFull) return payrollFull;
  const payrollName = formatPayrollFirstAndLastName(payrollRow);
  if (payrollName) return payrollName;
  const payrollEmployeeName = String(
    payrollRow?.employee_name ||
      payrollRow?.EmployeeName ||
      payrollRow?.['Employee Name'] ||
      payrollRow?.full_name ||
      ''
  ).trim();
  if (payrollEmployeeName) return payrollEmployeeName;
  const empParts = readParts(emp);
  const peopleFull = [empParts.firstName, empParts.middleName, empParts.lastName]
    .filter(Boolean)
    .join(' ')
    .trim();
  if (peopleFull) return peopleFull;
  return String(
    emp.employee_name ||
      emp.EmployeeName ||
      emp['Employee Name'] ||
      emp.full_name ||
      emp.DisplayName ||
      emp['Display Name'] ||
      emp.Name ||
      emp['Name'] ||
      ''
  ).trim();
}

export function resolveFormXIVMPEmployeeSerialNumber(emp = {}, rowIndex = 0) {
  const fromEmp = String(
    emp.EmployeeID ||
      emp['EmployeeID'] ||
      emp['Employee ID'] ||
      emp.employee_id ||
      emp['employee_id'] ||
      emp.Employee_ID ||
      emp.Employee_Number ||
      emp['Employee Number'] ||
      emp.EmployeeNumber ||
      emp.EmployeeCode ||
      emp['Employee Code'] ||
      emp.employee_number ||
      emp.GIDNumber ||
      emp.gid_number ||
      emp['GID Number'] ||
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

/** Gujarat Nature-of-work box: site location + workman designation (e.g. GJ-Amreli / Junior Engineer). */
export function resolveFormXIVDesignationFromMappedRow(row = null, headers = null) {
  if (!row || typeof row !== 'object') return '';
  if (Array.isArray(headers)) {
    const designationHeader = headers.find(isFormXIVMPNatureDesignationHeader);
    if (designationHeader) {
      const fromHeader = String(row[designationHeader] ?? '').trim();
      if (fromHeader) return fromHeader;
    }
  }
  return resolveFormXIVMPDesignation(row);
}

export function buildFormXIVGJNatureLocationWithDesignation(
  natureLocationText = '',
  employeeRow = null,
  headers = null
) {
  const location = String(natureLocationText || '').trim();
  const designation = resolveFormXIVDesignationFromMappedRow(employeeRow, headers);
  if (!designation) return location;
  if (!location) return designation;
  const locNorm = location.replace(/\s+/g, ' ').toLowerCase();
  const desNorm = designation.replace(/\s+/g, ' ').toLowerCase();
  if (locNorm.includes(desNorm)) return location;
  return `${location}\n${designation}`;
}

export function withFormXIVGJNatureLocationDesignation(
  headerFormData = {},
  employeeRow = null,
  headers = null,
  variant = ''
) {
  if (String(variant || '').toLowerCase() !== 'gj') return headerFormData || {};
  const base = headerFormData && typeof headerFormData === 'object' ? headerFormData : {};
  return {
    ...base,
    form_xiv_mp_nature_location: buildFormXIVGJNatureLocationWithDesignation(
      base.form_xiv_mp_nature_location,
      employeeRow,
      headers
    ),
  };
}

const FORM_XIV_MP_GROSS_PAY_KEYS = [
  'gross_pay',
  'Gross Pay',
  'grossPay',
  'total_earnings',
  'monthly_gross_amount',
  'Monthly Gross',
  'monthlyGross',
];
const FORM_XIV_MP_GROSS_PAY_PATTERNS = [/^gross_pay$/, /^total_earnings$/, /^monthly_gross_amount$/];

export function readFormXIVMPPayrollGrossPay(payrollRow) {
  if (!payrollRow || payrollRow.fetch_error) return '';
  const merged = mergePayrollRunEmployeePayload(payrollRow);
  const flat = flattenPayrollEarningColumns(merged);
  const fromFlat = readPayrollScalar(flat, FORM_XIV_MP_GROSS_PAY_KEYS, FORM_XIV_MP_GROSS_PAY_PATTERNS);
  if (fromFlat !== '') return fromFlat;
  if (flat !== merged) {
    const fromMerged = readPayrollScalar(merged, FORM_XIV_MP_GROSS_PAY_KEYS, FORM_XIV_MP_GROSS_PAY_PATTERNS);
    if (fromMerged !== '') return fromMerged;
  }
  if (payrollRow !== merged) {
    return readPayrollScalar(payrollRow, FORM_XIV_MP_GROSS_PAY_KEYS, FORM_XIV_MP_GROSS_PAY_PATTERNS);
  }
  return '';
}

/** People master salary when Sample Payroll has no gross/net for the workman. */
export function resolveFormXIVMPPeopleSalary(emp = {}) {
  const raw =
    emp?.monthly_salary ??
    emp?.['monthly_salary'] ??
    emp?.MonthlySalary ??
    emp?.['Monthly Salary'] ??
    emp?.monthlySalary ??
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

/** Payroll table rows usable for Form XIV wage rate — gross_pay or net_pay; no live Zoho. */
export function filterFormXIVMPPayrollRowsForWageRate(rows) {
  const flat = (Array.isArray(rows) ? rows : [])
    .filter((row) => row && typeof row === 'object' && row.fetch_error !== true)
    .map((row) => flattenPayrollEarningColumns(row));
  const withGross = filterPayrollRowsWithGrossPay(flat);
  if (withGross.length > 0) return withGross;
  return flat.filter((row) => payrollRowHasNetPay(row));
}

/** Resolve Form XIV wage-rate rows from Payroll table snapshots (pay-run shape or gross/net). */
export function resolveFormXIVMPPayrollRowsForAutofill(statutoryPayrollRows, monthCandidates = []) {
  const fromPayRun = filterFormXIXMPEligiblePayrollRows(statutoryPayrollRows);
  if (fromPayRun.length > 0) return fromPayRun;
  const fromTable = filterFormXIVMPPayrollRowsForWageRate(statutoryPayrollRows);
  if (fromTable.length > 0) return fromTable;

  const cached = getCachedForm15PayrollTableRows(monthCandidates);
  if (cached?.rows?.length > 0) {
    const fromCachedPayRun = filterFormXIXMPEligiblePayrollRows(cached.rows);
    if (fromCachedPayRun.length > 0) return fromCachedPayRun;
    const fromCachedTable = filterFormXIVMPPayrollRowsForWageRate(cached.rows);
    if (fromCachedTable.length > 0) return fromCachedTable;
  }

  const latest = getLatestCachedPayrollTableRows();
  if (latest?.rows?.length > 0) {
    const fromLatestPayRun = filterFormXIXMPEligiblePayrollRows(latest.rows);
    if (fromLatestPayRun.length > 0) return fromLatestPayRun;
    const fromLatestTable = filterFormXIVMPPayrollRowsForWageRate(latest.rows);
    if (fromLatestTable.length > 0) return fromLatestTable;
  }
  return [];
}

/** Load Form XIV wage-rate rows from Catalyst Payroll table only (no live Zoho fallback). */
export async function loadFormXIVMPPayrollRowsForAutofill(monthCandidates = [], options = {}) {
  const months = Array.isArray(monthCandidates) ? monthCandidates : [];
  const timeoutMs = Number(options.timeoutMs) > 0 ? Number(options.timeoutMs) : 45000;
  const tableLoad = await fetchPayrollTableRowsForMonths(months, {
    timeoutMs,
    force: Boolean(options.force),
    zohoFallback: false,
  });
  if (Array.isArray(tableLoad.rows) && tableLoad.rows.length > 0) {
    const payrollMonth = tableLoad.payrollMonth || months[0] || '';
    cacheForm15PayrollTableRows(payrollMonth, tableLoad.rows, tableLoad.meta || null);
    const eligible = filterFormXIVMPPayrollRowsForWageRate(tableLoad.rows);
    if (eligible.length > 0) return eligible;
  }
  return resolveFormXIVMPPayrollRowsForAutofill(tableLoad?.rows, months);
}

export function resolveFormXIVMPWageRate(emp = {}, payrollRow = null, options = {}) {
  if ((options.variant === 'gj' || options.preferNetPay === true) && isFormXIVGJFixedEmployee(emp)) {
    return FORM_XIV_GJ_FIXED_THARUN_WAGE_RATE;
  }

  const preferNetPay = options.variant === 'gj' || options.preferNetPay === true;
  if (payrollRow && !payrollRow.fetch_error) {
    if (preferNetPay) {
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
    const fromNet = readFormXIVMPPayrollNetPay(payrollRow);
    if (fromNet !== '') {
      const n = Number(String(fromNet).replace(/,/g, '').trim());
      if (Number.isFinite(n) && n > 0) return String(n);
      return String(fromNet).trim();
    }
  }

  const fromPeople = resolveFormXIVMPPeopleSalary(emp);
  if (fromPeople) return fromPeople;

  const monthDefault = resolveFormXIVKarnatakaMonthDefaultWageRate(
    emp,
    options.monthCandidates,
    options
  );
  if (monthDefault) return monthDefault;

  const mpMonthDefault = resolveFormXIVMadhyaPradeshMonthDefaultWageRate(
    emp,
    options.monthCandidates,
    options
  );
  if (mpMonthDefault) return mpMonthDefault;

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

const resolveFormXIVMPMappingHeaders = (headers, helpers = {}) => {
  const source = (Array.isArray(headers) ? headers : [])
    .map((h) => String(h || '').trim())
    .filter(Boolean);
  if (source.some(isFormXIVMPWageRateHeader) || headersIndicateFormXIVMPTable(source)) {
    return source;
  }
  return resolveFormXIVMPTableHeaders(headers, {
    item: helpers.item ?? null,
    fileName: helpers.fileName ?? helpers.formFileName ?? '',
    formHeader: helpers.formHeader ?? helpers.parsedFormHeader ?? null,
    sheetText: helpers.sheetText ?? '',
  });
};

const writeFormXIVMPWageRateOnRow = (row, wageRateHeader, rate, allHeaders = []) => {
  if (!row || !wageRateHeader || rate == null || String(rate).trim() === '') return;
  const value = String(rate).trim();
  row[wageRateHeader] = value;
  const aliasHeaders = [
    ...(Array.isArray(allHeaders) ? allHeaders : []),
    FORM_XIV_MP_CANONICAL_TABLE_HEADERS.find(isFormXIVMPWageRateHeader),
    FORM_XIV_RJ_CANONICAL_TABLE_HEADERS.find(isFormXIVMPWageRateHeader),
    'Wage rate (with particular of unit), in case of place work',
    'Wage rate with particulars or unit, in case of piece of work',
  ].filter(Boolean);
  aliasHeaders.forEach((header) => {
    if (header && isFormXIVMPWageRateHeader(header)) {
      row[header] = value;
    }
  });
};

/**
 * Form XIV / Form X_RJ/MH wage-rate payroll match.
 * Same-name people (same or other location) use Form B rules: ID first, then unique location.
 */
export function resolveFormXIVMPPayrollRowForWageRate(emp, payrollRows, formRow = null) {
  const rows = Array.isArray(payrollRows) ? payrollRows : [];
  if (!emp || rows.length === 0) return null;

  const byFullName = resolveFormBRajasthanPayrollRowForEmployee(emp, rows);
  if (byFullName && !byFullName.fetch_error) return byFullName;

  const primary = resolveFormXIXMPPayrollRowForEmployee(emp, rows);
  if (primary && !primary.fetch_error) {
    const sameName = rows.filter((row) => row && !row.fetch_error && formBRJNamesMatch(emp, row));
    if (sameName.length <= 1) return primary;
    const picked = resolveFormBRajasthanPayrollRowForEmployee(emp, sameName);
    if (picked) return picked;
  }

  const nameCandidates = [
    ...collectEmployeeNameCandidates(emp),
    formatFormXIVMPWorkmanName(emp),
    formRow?.__employeeLookupName,
  ]
    .map((n) => String(n || '').trim())
    .filter(Boolean);

  if (formRow && typeof formRow === 'object') {
    Object.keys(formRow).forEach((key) => {
      if (!isFormXIVMPWorkmanNameHeader(key)) return;
      const n = String(formRow[key] || '').trim();
      if (n && !/^enter\b/i.test(n)) nameCandidates.push(n);
    });
  }

  const uniqueNames = [...new Set(nameCandidates)];
  for (let i = 0; i < uniqueNames.length; i += 1) {
    const hit = resolvePayrollRowByDisplayName(uniqueNames[i], rows);
    if (!hit || hit.fetch_error) continue;
    const sameName = rows.filter((row) => row && !row.fetch_error && formBRJNamesMatch(emp, row));
    if (sameName.length <= 1 && formBRJNamesMatch(emp, hit)) return hit;
    const picked = resolveFormBRajasthanPayrollRowForEmployee(emp, sameName.length ? sameName : rows);
    if (picked) return picked;
  }

  for (let i = 0; i < uniqueNames.length; i += 1) {
    const candidate = uniqueNames[i];
    const fuzzy = rows.find((row) => {
      if (!row || row.fetch_error) return false;
      if (formBRJNamesMatch(emp, row)) return true;
      const payrollName =
        String(row.employee_name || row['employee_name'] || '').trim() ||
        [
          row.first_name || row.FirstName || '',
          row.middle_name || row.MiddleName || '',
          row.last_name || row.LastName || '',
        ]
          .map((v) => String(v || '').trim())
          .filter(Boolean)
          .join(' ') ||
        formatPayrollFirstAndLastName(row);
      return payrollName && personNamesMatch(candidate, payrollName);
    });
    if (!fuzzy) continue;
    const sameName = rows.filter((row) => row && !row.fetch_error && formBRJNamesMatch(emp, row));
    if (sameName.length <= 1) return fuzzy;
    const picked = resolveFormBRajasthanPayrollRowForEmployee(emp, sameName);
    if (picked) return picked;
  }
  return null;
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
    resolveFormXIVExportVariant(
      helpers.formHeader ?? helpers.parsedFormHeader ?? null,
      item,
      helpers.fileName ?? helpers.formFileName ?? '',
      helpers.sheetText ?? '',
      headers
    ) ||
    (headersIndicateFormXIVGJTable(headers) || headers.some(isFormXIVMPEntryDateHeader) ? 'gj' : 'mp');
  headers.forEach((header) => {
    if (isFormXIVMPSkipAutofillHeader(header)) {
      out[header] = '';
      return;
    }
    if (isFormXIVMPWorkmanNameHeader(header)) {
      out[header] = sanitizeValue(formatFormXIVMPWorkmanName(emp, payrollRow));
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
      const rate = sanitizeValue(
        resolveFormXIVMPWageRate(emp, payrollRow, {
          variant,
          monthCandidates: helpers.monthCandidates,
          item,
          fileName: helpers.fileName ?? helpers.formFileName ?? '',
          formHeader: helpers.formHeader ?? helpers.parsedFormHeader ?? null,
          sheetText: helpers.sheetText ?? '',
        })
      );
      writeFormXIVMPWageRateOnRow(out, header, rate, headers);
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

/** One payroll row per workman for Form XIV / Form X_RJ/MH wage-rate autofill. */
export function buildFormXIVMPPayrollRowResolver(payrollRows) {
  const rows = Array.isArray(payrollRows) ? payrollRows : [];
  const primary = buildFormXIXMPPayrollRowResolver(rows);
  const usedFallback = new Set();
  const rowKey = (row) =>
    String(
      row?.employee_id ||
        row?.employee_number ||
        row?.EmployeeID ||
        `${row?.first_name || ''}|${row?.middle_name || ''}|${row?.last_name || ''}|${row?.employee_name || ''}|${row?.gross_pay || ''}`
    ).trim();

  return (emp, formRow) => {
    const byFullName = resolveFormBRajasthanPayrollRowForEmployee(emp, rows);
    if (byFullName && !byFullName.fetch_error) {
      const key = rowKey(byFullName);
      if (key) usedFallback.add(key);
      return byFullName;
    }

    const hit = primary(emp);
    if (hit && !hit.fetch_error) {
      const sameName = rows.filter((row) => row && !row.fetch_error && formBRJNamesMatch(emp, row));
      if (sameName.length > 1) {
        const picked = resolveFormBRajasthanPayrollRowForEmployee(emp, sameName);
        if (picked) {
          const key = rowKey(picked);
          if (key) usedFallback.add(key);
          return picked;
        }
      } else {
        const key = rowKey(hit);
        if (key) usedFallback.add(key);
        return hit;
      }
    }

    const available = rows.filter((row) => {
      const key = rowKey(row);
      return !key || !usedFallback.has(key);
    });
    const fallback = resolveFormXIVMPPayrollRowForWageRate(
      emp,
      available.length > 0 ? available : rows,
      formRow
    );
    if (fallback && !fallback.fetch_error) {
      const key = rowKey(fallback);
      if (key) usedFallback.add(key);
      return fallback;
    }
    return null;
  };
}

/** Fill wage rate from pay-run gross_pay when rows were mapped without payroll. */
export function enrichFormXIVMPPayrollRows(mappedData, employees, headers, helpers = {}) {
  const hdrs = resolveFormXIVMPMappingHeaders(headers, helpers);
  const wageRateHeader =
    (Array.isArray(headers) ? headers : []).find(isFormXIVMPWageRateHeader) ||
    hdrs.find(isFormXIVMPWageRateHeader);
  if (!wageRateHeader) return 0;
  const variant =
    helpers.variant ||
    resolveFormXIVExportVariant(
      helpers.formHeader ?? helpers.parsedFormHeader ?? null,
      helpers.item ?? null,
      helpers.fileName ?? helpers.formFileName ?? '',
      helpers.sheetText ?? '',
      hdrs
    ) ||
    (headersIndicateFormXIVGJTable(hdrs) || hdrs.some(isFormXIVMPEntryDateHeader) ? 'gj' : 'mp');
  const {
    resolvePayrollRow = null,
    sanitizeValue = (v) => String(v ?? '').trim(),
    overwrite = false,
    monthCandidates = null,
    item = null,
    fileName = '',
    formHeader = null,
    parsedFormHeader = null,
    sheetText = '',
    payrollRows = null,
    // Gujarat / Karnataka: FirstName + LastName only — never ID / loose personNamesMatch.
    strictFirstLastMatch = false,
  } = helpers;
  const wageRateHints = buildFormXIVKarnatakaWageRateHints({
    item,
    fileName,
    formHeader: formHeader ?? parsedFormHeader,
    sheetText,
  });
  if (!Array.isArray(mappedData) || mappedData.length === 0) return 0;
  const useStrictFirstLast =
    Boolean(strictFirstLastMatch) || variant === 'gj' || variant === 'ka';
  const fallbackResolver =
    !useStrictFirstLast && Array.isArray(payrollRows) && payrollRows.length > 0
      ? buildFormXIVMPPayrollRowResolver(payrollRows)
      : null;
  let hits = 0;
  mappedData.forEach((row, rowIndex) => {
    const empItem = employees[rowIndex];
    const emp = empItem?.Employee || empItem?.employee || empItem;
    const existing = getFormXIVMPRowValueForHeader(row, wageRateHeader);
    if (!overwrite && existing) return;
    let payrollRow =
      typeof resolvePayrollRow === 'function' ? resolvePayrollRow(emp, row, rowIndex) : null;
    if ((!payrollRow || payrollRow.fetch_error) && fallbackResolver) {
      payrollRow = fallbackResolver(emp, row);
    }
    if (
      !useStrictFirstLast &&
      (!payrollRow || payrollRow.fetch_error) &&
      Array.isArray(payrollRows) &&
      payrollRows.length > 0
    ) {
      payrollRow = resolveFormXIVMPPayrollRowForWageRate(emp, payrollRows, row);
    }
    const rate = resolveFormXIVMPWageRate(emp, payrollRow && !payrollRow.fetch_error ? payrollRow : null, {
      variant,
      monthCandidates,
      ...wageRateHints,
    });
    if (rate) {
      writeFormXIVMPWageRateOnRow(row, wageRateHeader, sanitizeValue(rate), [
        ...hdrs,
        ...(Array.isArray(headers) ? headers : []),
      ]);
      hits += 1;
    } else if (existing && overwrite) {
      // Keep any alias-key wage rate visible under the UI column header.
      writeFormXIVMPWageRateOnRow(row, wageRateHeader, existing, [
        ...hdrs,
        ...(Array.isArray(headers) ? headers : []),
      ]);
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
    match: /serial\s+(?:no\.?|number)\s+in\s+the\s+register|s\.?\s*no\.?\s+in\s+the\s+register|sl\.?\s*no\.?\s+of\s+the\s+register/i,
    rowTest: isFormXIVMPSerialNumberHeader,
  },
  {
    label: 'Nature of employment / designation',
    match: /nature\s+of\s+employ|designat/i,
    rowTest: isFormXIVMPNatureDesignationHeader,
  },
  {
    label: "Wage rate with particulars or unit, in case of piece of work",
    match: /wage\s*['']?\s*rate|piece\s+of\s+work|place\s+work|particular/i,
    rowTest: isFormXIVMPWageRateHeader,
  },
  {
    label: 'Wage period',
    match: /wage\s+period/i,
    rowTest: isFormXIVMPWagePeriodHeader,
  },
  {
    label: 'Tenure of employment',
    match: /tenure\s+of\s+employ|period\s+of\s+employ/i,
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
  if (variant === 'rj') {
    const tableHdrs = Array.isArray(hints.tableHeaders) ? hints.tableHeaders : [];
    const headers =
      headersIndicateFormXIVMPTable(tableHdrs) && tableHdrs.length > 0
        ? [...tableHdrs]
        : [...FORM_XIV_MP_CANONICAL_TABLE_HEADERS];
    return {
      headers,
      variant,
      sheetName: accessor.sheetName,
    };
  }
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

export function allocateUniqueFormXIVMPDownloadFileName(baseName, usedNames, variant = 'mp') {
  const root = String(baseName || 'Employee').trim() || 'Employee';
  const count = usedNames.get(root) || 0;
  usedNames.set(root, count + 1);
  const suffix = count > 0 ? `_${count + 1}` : '';
  const v = String(variant || 'mp').toLowerCase();
  const prefix =
    v === 'rj'
      ? 'Form_X_RJ'
      : v === 'ka'
        ? 'Form_XIV_KA'
        : v === 'gj'
          ? 'Form_XIV_GJ'
          : 'Form_XIV_MP';
  return `${prefix}_${root}${suffix}.xlsx`;
}

/**
 * Prefer live state/form identity (gj/ka/rj/mp) over a stale formXIVVariant stamp
 * (e.g. callers that only branched on Karnataka / Rajasthan, or a prior GJ session).
 */
export function resolveFormXIVExportVariant(
  formHeader,
  rowItem,
  fileName,
  sheetText = '',
  tableHeaders = null
) {
  const resolved = resolveFormXIVVariant(formHeader, rowItem, fileName, sheetText, tableHeaders);
  const existing = String(formHeader?.formXIVVariant || '').toLowerCase();
  if (
    resolved === 'rj' ||
    existing === 'rj' ||
    isFormXRajasthanEmploymentCardContext(formHeader, rowItem, fileName, sheetText)
  ) {
    return 'rj';
  }
  // Live ka/gj always wins over a stale 'mp' stamp.
  if (resolved === 'ka' || resolved === 'gj') return resolved;
  // Live Madhya Pradesh identity wins over a stale 'gj'/'ka' stamp from a wrong template link.
  if (
    resolved === 'mp' &&
    isFormXIVMadhyaPradeshContext(formHeader, rowItem, fileName, sheetText)
  ) {
    return 'mp';
  }
  if (existing === 'ka' || existing === 'gj') return existing;
  return resolved || existing || 'mp';
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
/** Rajasthan Form X — header labels in column B, values in bordered column C. */
const FORM_XIV_RJ_HEADER_VALUE_COL = 3;
/** MP / Karnataka stacked employment card — labels in A–D, values in column E. */
const FORM_XIV_MP_STACKED_VALUE_COL = 5;
const FORM_XIV_MP_STACKED_VALUE_COL_TO = 12;
/** Karnataka Form XIV — official labels + dotted fill-in + value on one line. */
const FORM_XIV_KA_VALUE_COL = 6;
const FORM_XIV_KA_LABEL_COL = 2;
const FORM_XIV_KA_TITLE_COL = 3;
const FORM_XIV_KA_TITLE_MERGE_END = 5;
/** Merge label+value through column E so long lines stay visible; keep F for signature. */
const FORM_XIV_KA_LINE_MERGE_END = 5;
const FORM_XIV_KA_DOT_LEADER = '.'.repeat(40);
const FORM_XIV_KA_HEADER_OFFICIAL = [
  {
    match: /name\s+and\s+address\s+(?:of|if)\s+(?:the\s+)?contractor/i,
    text: 'Name and address of contractor',
  },
  {
    match: /establishment/i,
    text: 'Name and address of Establishment in/under which contract is carried on :',
  },
  {
    match: /nature\s+and\s+location\s+of\s+work|nature\s+of\s+work\s+and\s+location/i,
    text: 'Nature and location of work:',
  },
  {
    match: /principal\s+employer/i,
    text: 'Name and address of Principal Employer:',
  },
];
const FORM_XIV_KA_WORKMAN_OFFICIAL = [
  '1 Name of the Workman',
  '2 Serial No. in the Register of workmen employed',
  '3 Nature of employment/Designation',
  '4 Wage rate (with particulars of unit in case of piece-work)',
  '5 Wage period',
  '6 Tenure of employment',
  '7 Remarks',
];

const detectFormXRajasthanWorksheetLayout = (worksheet, maxScanRow = 24) => {
  if (!worksheet) return false;
  let contractorLabelInB = false;
  let tableHeaderRow = false;
  for (let r = 1; r <= maxScanRow; r += 1) {
    const labelB = formXIVMPExcelCellValueToString(worksheet.getCell(r, 2)?.value).trim();
    if (/name\s+and\s+address\s+.*contractor/i.test(labelB)) contractorLabelInB = true;
    // Form X_RJ is tabular: several workman column headers on one row (not stacked 1./2./3. labels).
    let headerHits = 0;
    for (let c = 1; c <= 12; c += 1) {
      const cellNorm = formXIVMPHeaderNorm(
        formXIVMPExcelCellValueToString(worksheet.getCell(r, c)?.value)
      );
      if (!cellNorm) continue;
      if (/name\s+of\s+the\s+workman/.test(cellNorm)) headerHits += 1;
      if (/serial|s\.?\s*no|sl\.?\s*no/.test(cellNorm) && /register|workman/.test(cellNorm)) {
        headerHits += 1;
      }
      if (/nature/.test(cellNorm) && /employ|designat/.test(cellNorm)) headerHits += 1;
      if (/wage\s+period/.test(cellNorm)) headerHits += 1;
      if (/tenure|period\s+of\s+employ/.test(cellNorm)) headerHits += 1;
      if (/wage/.test(cellNorm) && /rate|piece|place/.test(cellNorm)) headerHits += 1;
      if (/remark/.test(cellNorm)) headerHits += 1;
    }
    if (headerHits >= 3) tableHeaderRow = true;
  }
  return contractorLabelInB && tableHeaderRow;
};

const FORM_XIV_RJ_DOT_LEADER_RE = /(?:\.{3,}|…{2,}|[.…·_]{6,})/u;

const stripFormXIVRJHeaderLeader = (text) =>
  String(text || '')
    .replace(/\s*(?:\.{3,}|…{2,}|[.…·_]{6,}).*$/u, '')
    .replace(/[\s.…·_]+$/u, '')
    .trim();

/** Form X_RJ: label + dotted fill-in share one (often wide-merged) cell — e.g. contractor line. */
const isFormXIVRJInlineDottedHeaderLabel = (raw) => {
  const s = String(raw || '').trim();
  if (!s || !FORM_XIV_RJ_DOT_LEADER_RE.test(s)) return false;
  return FORM_XIV_MP_HEADER_SPECS.some((spec) => spec.match.test(s) || spec.match.test(formXIVMPHeaderNorm(s)));
};

/** True when column C (value col) sits inside the label merge — cannot hold a separate value. */
const formXIVRJLabelMergeCoversValueCol = (worksheet, labelRow, labelCol) => {
  const ranges = parseExcelJsMergeRanges(worksheet);
  const labelMerge = ranges.find(
    (m) =>
      labelRow >= m.top &&
      labelRow <= m.bottom &&
      labelCol >= m.left &&
      labelCol <= m.right
  );
  if (!labelMerge) return false;
  return (
    FORM_XIV_RJ_HEADER_VALUE_COL >= labelMerge.left &&
    FORM_XIV_RJ_HEADER_VALUE_COL <= labelMerge.right
  );
};

const writeFormXIVRJInlineHeaderValue = (worksheet, labelRow, labelCol, value) => {
  const text = String(value ?? '').trim();
  if (!text || labelRow < 1 || labelCol < 1) return;
  if (isFormXIVMPHeaderValueEmptyOrPlaceholder(text)) return;
  const topLeft = resolveGJMergeTopLeft(worksheet, labelRow, labelCol);
  const raw = formXIVMPExcelCellValueToString(
    worksheet.getCell(topLeft.row, topLeft.col)?.value
  ).trim();
  const matchedSpec = FORM_XIV_MP_HEADER_SPECS.find(
    (spec) => spec.match.test(raw) || spec.match.test(formXIVMPHeaderNorm(raw))
  );
  const clean =
    stripFormXIVRJHeaderLeader(raw) ||
    matchedSpec?.label ||
    'Name and address of contractor';
  // Keep the template's dotted leader length so alignment stays the same.
  const leaderMatch = raw.match(FORM_XIV_RJ_DOT_LEADER_RE);
  const leader = leaderMatch ? leaderMatch[0] : '.'.repeat(40);
  const cell = worksheet.getCell(topLeft.row, topLeft.col);
  cell.value = `${clean}${leader} ${text}`;
  cell.alignment = {
    ...(cell.alignment || {}),
    horizontal: 'left',
    vertical: 'middle',
    wrapText: true,
    shrinkToFit: false,
  };
};

const resolveFormXRJHeaderValueCell = (worksheet, labelRow, labelCol) => {
  // Prefer the official RJ value column (C) when it is free of the label merge.
  if (!formXIVRJLabelMergeCoversValueCol(worksheet, labelRow, labelCol)) {
    return resolveGJMergeTopLeft(worksheet, labelRow, FORM_XIV_RJ_HEADER_VALUE_COL);
  }
  const targetCol = findFormXIVValueColumnBesideLabel(
    worksheet,
    labelRow,
    labelCol,
    FORM_XIV_RJ_HEADER_VALUE_COL,
    FORM_XIV_MP_WORKMAN_FIELD_SPECS
  );
  return resolveGJMergeTopLeft(worksheet, labelRow, targetCol);
};

const clearFormXRJHeaderValueMerge = (worksheet, labelRow, labelCol) => {
  const { row, col } = resolveFormXRJHeaderValueCell(worksheet, labelRow, labelCol);
  const ranges = parseExcelJsMergeRanges(worksheet);
  const merge = ranges.find(
    (m) => row >= m.top && row <= m.bottom && col >= m.left && col <= m.right
  );
  if (!merge) {
    const cellStr = formXIVMPExcelCellValueToString(worksheet.getCell(row, col)?.value).trim();
    if (isFormXIVPlaceholderCell(cellStr) || isFormXIVMPHeaderValueEmptyOrPlaceholder(cellStr)) {
      worksheet.getCell(row, col).value = '';
    }
    return;
  }
  for (let r = merge.top; r <= merge.bottom; r += 1) {
    for (let c = merge.left; c <= merge.right; c += 1) {
      const cellStr = formXIVMPExcelCellValueToString(worksheet.getCell(r, c)?.value).trim();
      if (!cellStr || isFormXIVPlaceholderCell(cellStr) || isFormXIVMPHeaderValueEmptyOrPlaceholder(cellStr)) {
        worksheet.getCell(r, c).value = '';
      }
    }
  }
};

const writeFormXIVRJBoxedHeaderValue = (worksheet, labelRow, labelCol, value) => {
  const text = String(value ?? '').trim();
  if (!text || labelRow < 1) return;
  if (isFormXIVMPHeaderValueEmptyOrPlaceholder(text)) return;
  const labelRaw = formXIVMPExcelCellValueToString(worksheet.getCell(labelRow, labelCol)?.value).trim();
  const dottedContractor =
    isFormXIVRJInlineDottedHeaderLabel(labelRaw) ||
    (/name\s+and\s+address\s+.*contractor/i.test(labelRaw) && FORM_XIV_RJ_DOT_LEADER_RE.test(labelRaw));
  // Official Form X_RJ: keep dotted contractor label in B untouched; put value in C like other headers.
  // Only fall back to inline when a wide merge covers column C (value would land off-form).
  if (dottedContractor && formXIVRJLabelMergeCoversValueCol(worksheet, labelRow, labelCol)) {
    writeFormXIVRJInlineHeaderValue(worksheet, labelRow, labelCol, text);
    return;
  }
  clearFormXRJHeaderValueMerge(worksheet, labelRow, labelCol);
  const { row, col } = resolveFormXRJHeaderValueCell(worksheet, labelRow, labelCol);
  const bounds = resolveFormXRJDataBoxBounds(worksheet, row, col);
  for (let r = bounds.top; r <= bounds.bottom; r += 1) {
    const sheetRow = worksheet.getRow(r);
    if (!sheetRow.height || sheetRow.height < 18) sheetRow.height = Math.max(18, text.split(/\n/).length * 15);
  }
  setFormXIVMPWorkmanCellValue(worksheet, row, col, text, { variant: 'mp', inBox: true });
  const cell = worksheet.getCell(row, col);
  cell.alignment = {
    ...(cell.alignment || {}),
    horizontal: 'left',
    vertical: 'top',
    wrapText: true,
    shrinkToFit: false,
  };
};

const resolveFormXRJTableHeaderMatchCol = (worksheet, headerRow, header) => {
  const target = formXIVMPHeaderNorm(header);
  for (let c = 1; c <= 12; c += 1) {
    const cellNorm = formXIVMPHeaderNorm(
      formXIVMPExcelCellValueToString(worksheet.getCell(headerRow, c)?.value)
    );
    if (!cellNorm) continue;
    if (cellNorm === target || cellNorm.includes(target) || target.includes(cellNorm)) return c;
    if (isFormXIVMPWorkmanNameHeader(header) && /name\s+of\s+the\s+workman/.test(cellNorm)) return c;
    if (isFormXIVMPSerialNumberHeader(header) && /serial|s\.?\s*no|sl\.?\s*no/.test(cellNorm) && /register/.test(cellNorm)) {
      return c;
    }
    if (isFormXIVMPNatureDesignationHeader(header) && /nature/.test(cellNorm) && /employ|designat/.test(cellNorm)) {
      return c;
    }
    if (
      isFormXIVMPWageRateHeader(header) &&
      /wage/.test(cellNorm) &&
      /rate|piece|place|particular/.test(cellNorm) &&
      !/wage\s+period/.test(cellNorm)
    ) {
      return c;
    }
    if (isFormXIVMPWagePeriodHeader(header) && /wage\s+period/.test(cellNorm)) return c;
    if (isFormXIVMPTenureHeader(header) && /tenure|period\s+of\s+employ/.test(cellNorm)) return c;
    if (isFormXIVMPRemarksHeader(header) && /remark/.test(cellNorm)) return c;
    if (/signature/.test(formXIVMPHeaderNorm(header)) && /signature/.test(cellNorm)) return c;
  }
  return -1;
};

const resolveFormXRJTableHeaderColumns = (worksheet, headerRow, hdrs) => {
  const exportHdrs = resolveFormXIVMPTableHeaders(hdrs);
  const tryRows = [headerRow, headerRow + 1].filter((r) => r >= 1);
  let bestCols = null;
  let bestHits = 0;
  tryRows.forEach((r) => {
    const columnByHeader = exportHdrs.map((header) => resolveFormXRJTableHeaderMatchCol(worksheet, r, header));
    const hits = columnByHeader.filter((col) => col > 0).length;
    if (hits > bestHits) {
      bestHits = hits;
      bestCols = columnByHeader;
    }
  });
  return bestCols;
};

const isFormXRJColumnIndexRow = (worksheet, row, columnByHeader) => {
  const usedCols = [...new Set(columnByHeader.filter((col) => col > 0))].sort((a, b) => a - b);
  if (usedCols.length < 3) return false;
  let indexHits = 0;
  for (let i = 0; i < usedCols.length; i += 1) {
    const col = usedCols[i];
    const raw = formXIVMPExcelCellValueToString(worksheet.getCell(row, col)?.value).trim();
    const asNum = Number(raw);
    if (raw && Number.isFinite(asNum) && (asNum === i + 1 || asNum === col - usedCols[0] + 1)) {
      indexHits += 1;
    }
  }
  return indexHits >= 3;
};

const findFormXRJDataBoxStartRow = (worksheet, headerRow, columnByHeader, maxScanRows = 30) => {
  for (let r = headerRow + 1; r <= Math.min(headerRow + 6, maxScanRows); r += 1) {
    if (isFormXRJColumnIndexRow(worksheet, r, columnByHeader)) {
      return r + 1;
    }
  }
  return headerRow + 3;
};

const resolveFormXRJDataBoxBounds = (worksheet, dataStartRow, col) => {
  const topLeft = resolveGJMergeTopLeft(worksheet, dataStartRow, col);
  const ranges = parseExcelJsMergeRanges(worksheet);
  const merge = ranges.find(
    (m) =>
      topLeft.row >= m.top &&
      topLeft.row <= m.bottom &&
      topLeft.col >= m.left &&
      topLeft.col <= m.right
  );
  if (merge) {
    // Keep vertical merges for tall boxes, but never spill across table columns
    // (Nature of employment must not paint into Wage rate).
    return { top: merge.top, bottom: merge.bottom, left: topLeft.col, right: topLeft.col };
  }
  return { top: topLeft.row, bottom: topLeft.row + 3, left: topLeft.col, right: topLeft.col };
};

const isFormXRJNumericWageRateValue = (raw) => {
  const s = String(raw ?? '')
    .replace(/[,₹]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  if (!s) return false;
  if (/[a-z]/i.test(s) && !/^(rs\.?|inr)\b/i.test(s)) {
    const stripped = s.replace(/^(rs\.?|inr)\s*/i, '').trim();
    if (/[a-z]/i.test(stripped)) return false;
  }
  const num = Number(String(s).replace(/[^\d.]/g, ''));
  return Number.isFinite(num) && num > 0;
};

const setFormXRJTableBoxCellValue = (worksheet, dataStartRow, col, value) => {
  const text = String(value ?? '').trim();
  if (!text || dataStartRow < 1 || col < 1) return;
  const writePos = resolveGJMergeTopLeft(worksheet, dataStartRow, col);
  const bounds = resolveFormXRJDataBoxBounds(worksheet, dataStartRow, col);
  for (let r = bounds.top; r <= bounds.bottom; r += 1) {
    const row = worksheet.getRow(r);
    if (!row.height || row.height < 22) row.height = 36;
  }
  setFormXIVMPWorkmanCellValue(worksheet, writePos.row, writePos.col, text, {
    variant: 'mp',
    inBox: true,
  });
  const cell = worksheet.getCell(writePos.row, writePos.col);
  cell.alignment = {
    ...(cell.alignment || {}),
    horizontal: 'center',
    vertical: 'middle',
    wrapText: true,
    shrinkToFit: false,
  };
};

const ensureFormXRJDataBoxRowHeights = (worksheet, layout) => {
  if (!worksheet || !layout) return;
  const { dataStartRow, columnByHeader } = layout;
  const usedCols = [...new Set(columnByHeader.filter((col) => col > 0))];
  let maxBottom = dataStartRow;
  usedCols.forEach((col) => {
    const bounds = resolveFormXRJDataBoxBounds(worksheet, dataStartRow, col);
    maxBottom = Math.max(maxBottom, bounds.bottom);
  });
  for (let r = dataStartRow; r <= maxBottom; r += 1) {
    const row = worksheet.getRow(r);
    if (!row.height || row.height < 22) row.height = 36;
  }
};

export function resolveFormXRJTableExportLayout(worksheet, hdrs) {
  if (!worksheet) return null;
  const exportHdrs = resolveFormXIVMPTableHeaders(hdrs);
  const maxScanRows = 30;
  for (let r = 1; r <= maxScanRows; r += 1) {
    const columnByHeader = resolveFormXRJTableHeaderColumns(worksheet, r, hdrs) || exportHdrs.map((header) =>
      resolveFormXRJTableHeaderMatchCol(worksheet, r, header)
    );
    const hits = columnByHeader.filter((col) => col > 0).length;
    if (hits < 3) continue;
    const dataStartRow = findFormXRJDataBoxStartRow(worksheet, r, columnByHeader, maxScanRows);
    const remappedCols = columnByHeader.map((col) => {
      if (!col || col < 1) return col;
      return resolveGJMergeTopLeft(worksheet, dataStartRow, col).col;
    });
    return { headerRow: r, dataStartRow, columnByHeader: remappedCols, hdrs: exportHdrs };
  }
  return null;
};

const writeFormXIVRJTableRowToWorksheet = (worksheet, row, layout) => {
  if (!worksheet || !row || !layout) return;
  ensureFormXRJDataBoxRowHeights(worksheet, layout);
  const { dataStartRow, columnByHeader, hdrs } = layout;
  hdrs.forEach((header, j) => {
    const col = columnByHeader[j];
    if (!col || col < 1) return;
    let val = getFormXIVMPRowValueForHeader(row, header);
    // Autofill may store wage rate under the template "place work" key or MP "piece of work" key.
    if (!val && isFormXIVMPWageRateHeader(header)) {
      for (const [k, v] of Object.entries(row)) {
        if (String(k || '').startsWith('__')) continue;
        if (!isFormXIVMPWageRateHeader(k)) continue;
        const hit = String(v ?? '').trim();
        if (hit) {
          val = hit;
          break;
        }
      }
    }
    if (!val) return;
    // Wage rate: only write real amounts — leave blank when missing (never designation text).
    if (isFormXIVMPWageRateHeader(header) && !isFormXRJNumericWageRateValue(val)) return;
    setFormXRJTableBoxCellValue(worksheet, dataStartRow, col, val);
  });
};

const clearFormXRJTableDataRows = (worksheet, layout) => {
  if (!worksheet || !layout) return;
  const { dataStartRow, columnByHeader } = layout;
  const usedCols = [...new Set(columnByHeader.filter((col) => col > 0))];
  // Break horizontal merges that span multiple table columns (Nature must not fill Wage rate).
  const ranges = parseExcelJsMergeRanges(worksheet);
  ranges.forEach((m) => {
    if (!m || m.right <= m.left) return;
    if (m.bottom < dataStartRow || m.top > dataStartRow + 8) return;
    let hits = 0;
    usedCols.forEach((col) => {
      if (col >= m.left && col <= m.right) hits += 1;
    });
    if (hits < 2) return;
    try {
      worksheet.unMergeCells(m.top, m.left, m.bottom, m.right);
    } catch (_) {
      /* ignore already-unmerged */
    }
  });
  usedCols.forEach((col) => {
    const bounds = resolveFormXRJDataBoxBounds(worksheet, dataStartRow, col);
    for (let r = bounds.top; r <= bounds.bottom; r += 1) {
      for (let c = bounds.left; c <= bounds.right; c += 1) {
        worksheet.getCell(r, c).value = '';
      }
    }
  });
};

const excelColLettersToNumber = (letters) => {
  let n = 0;
  for (const ch of String(letters || '').toUpperCase()) {
    n = n * 26 + (ch.charCodeAt(0) - 64);
  }
  return n;
};

const parseExcelJsMergeRanges = (worksheet) => {
  const raw = worksheet?.model?.merges;
  if (!Array.isArray(raw) || raw.length === 0) return [];
  return raw
    .map((range) => {
      const parts = String(range || '').split(':');
      if (parts.length !== 2) return null;
      const parseRef = (ref) => {
        const m = String(ref).match(/^([A-Z]+)(\d+)$/i);
        if (!m) return null;
        return { col: excelColLettersToNumber(m[1]), row: Number(m[2]) };
      };
      const start = parseRef(parts[0]);
      const end = parseRef(parts[1]);
      if (!start || !end) return null;
      return {
        top: Math.min(start.row, end.row),
        left: Math.min(start.col, end.col),
        bottom: Math.max(start.row, end.row),
        right: Math.max(start.col, end.col),
      };
    })
    .filter(Boolean);
};

const resolveGJMergeTopLeft = (worksheet, row, col) => {
  if (!worksheet || row < 1 || col < 1) return { row, col };
  const ranges = parseExcelJsMergeRanges(worksheet);
  for (let i = 0; i < ranges.length; i += 1) {
    const m = ranges[i];
    if (row >= m.top && row <= m.bottom && col >= m.left && col <= m.right) {
      return { row: m.top, col: m.left };
    }
  }
  return { row, col };
};

const findGJWorkmanValueMergeOnRow = (worksheet, row, labelCol) => {
  const ranges = parseExcelJsMergeRanges(worksheet);
  let best = null;
  ranges.forEach((m) => {
    if (row < m.top || row > m.bottom) return;
    if (m.left <= labelCol) return;
    if (m.right - m.left < 1) return;
    if (!best || m.left < best.col) best = { row, col: m.left };
  });
  if (best) return best;
  const col = findFormXIVValueColumnBesideLabel(
    worksheet,
    row,
    labelCol,
    FORM_XIV_GJ_DEFAULT_VALUE_COL,
    FORM_XIV_GJ_WORKMAN_FIELD_SPECS
  );
  return { row, col: Math.min(Math.max(col, FORM_XIV_GJ_DEFAULT_VALUE_COL), 12) };
};

const resolveGJWorkmanValueBounds = (worksheet, row, labelCol) => {
  if (!worksheet || row < 1) {
    return { top: row, bottom: row, left: FORM_XIV_GJ_DEFAULT_VALUE_COL, right: 6 };
  }
  const { col: valueCol } = findGJWorkmanValueMergeOnRow(worksheet, row, labelCol ?? 1);
  const topLeft = resolveGJMergeTopLeft(worksheet, row, valueCol);
  const ranges = parseExcelJsMergeRanges(worksheet);
  const merge = ranges.find(
    (m) =>
      topLeft.row >= m.top &&
      topLeft.row <= m.bottom &&
      topLeft.col >= m.left &&
      topLeft.col <= m.right
  );
  if (merge) {
    return { top: merge.top, bottom: merge.bottom, left: merge.left, right: merge.right };
  }

  const left = Math.max(topLeft.col, FORM_XIV_GJ_DEFAULT_VALUE_COL);
  let right = left;
  for (let c = left + 1; c <= 12; c += 1) {
    const cellStr = formXIVMPExcelCellValueToString(worksheet.getCell(row, c)?.value).trim();
    if (!cellStr || isFormXIVPlaceholderCell(cellStr) || /^[_\s.-]+$/.test(cellStr)) {
      right = c;
      continue;
    }
    if (cellStr.length <= 4 && /^[_\-.…\s]+$/u.test(cellStr)) {
      right = c;
      continue;
    }
    break;
  }
  if (right === left && (labelCol == null || labelCol <= 1)) {
    right = 6;
  }
  return { top: row, bottom: row, left, right };
};

const normalizeFormXIVGJWorkmanValueBorders = (worksheet, bounds) => {
  if (!worksheet || !bounds) return;
  for (let r = bounds.top; r <= bounds.bottom; r += 1) {
    for (let c = bounds.left; c <= bounds.right; c += 1) {
      const cell = worksheet.getCell(r, c);
      const existing = cell.border || {};
      const onTop = r === bounds.top;
      const onBottom = r === bounds.bottom;
      const onLeft = c === bounds.left;
      const onRight = c === bounds.right;
      const nextBorder = {};
      if (onTop && existing.top) nextBorder.top = existing.top;
      if (onBottom && existing.bottom) nextBorder.bottom = existing.bottom;
      if (onLeft && existing.left) nextBorder.left = existing.left;
      if (onRight && existing.right) nextBorder.right = existing.right;
      cell.border = nextBorder;
    }
  }
};

const clearFormXIVGJWorkmanValueBand = (worksheet, row, labelCol) => {
  if (!worksheet || row < 1) return;
  const bounds = resolveGJWorkmanValueBounds(worksheet, row, labelCol);
  for (let r = bounds.top; r <= bounds.bottom; r += 1) {
    for (let c = bounds.left; c <= bounds.right; c += 1) {
      worksheet.getCell(r, c).value = '';
    }
  }
  normalizeFormXIVGJWorkmanValueBorders(worksheet, bounds);
};

const isFormXIVLabelContinuationCell = (cellStr, workmanSpecs = null) => {
  const s = normalizeFormXIVMPLabelText(cellStr);
  if (!s) return false;
  if (/^form\s*xiv|employment\s+card|see\s+rule\s+76|vide\s+rule\s+76/i.test(s)) return false;
  if (FORM_XIV_MP_HEADER_SPECS.some((spec) => spec.match.test(s))) return true;
  const specs = workmanSpecs || [...FORM_XIV_GJ_WORKMAN_FIELD_SPECS, ...FORM_XIV_MP_WORKMAN_FIELD_SPECS];
  return specs.some((spec) => looksLikeFormXIVMPWorkmanLabelCell(s, spec));
};

const isFormXIVStackedWorkmanKeyword = (norm) =>
  /workman|serial|s\.?\s*no|designat|employ|wage|tenure|remarks|entry\s+into\s+service/.test(
    String(norm || '')
  );

/** True when col A is only an ordinal (1 / 1.) and the workman label sits in B–D. */
const isFormXIVSplitOrdinalLabelRow = (worksheet, row) => {
  if (!worksheet || row < 1) return false;
  const ordinalCell = formXIVMPExcelCellValueToString(worksheet.getCell(row, 1)?.value).trim();
  if (!/^\d+[\.\)]?$/.test(ordinalCell)) return false;
  for (let c = 2; c <= 4; c += 1) {
    const labelStr = formXIVMPExcelCellValueToString(worksheet.getCell(row, c)?.value).trim();
    if (!labelStr) continue;
    if (isFormXIVStackedWorkmanKeyword(formXIVMPHeaderNorm(labelStr))) return true;
  }
  return false;
};

const findFormXIVValueColumnBesideLabel = (worksheet, row, labelCol, fallbackCol, workmanSpecs = null) => {
  if (!worksheet || row < 1 || labelCol < 1) return fallbackCol;
  // Never write into empty cells that are part of the label's merge (A–D band).
  const ranges = parseExcelJsMergeRanges(worksheet);
  let skipUntil = labelCol;
  for (let i = 0; i < ranges.length; i += 1) {
    const m = ranges[i];
    if (row >= m.top && row <= m.bottom && labelCol >= m.left && labelCol <= m.right) {
      skipUntil = Math.max(skipUntil, m.right);
      break;
    }
  }
  for (let vc = skipUntil + 1; vc <= 14; vc += 1) {
    const nt = formXIVMPExcelCellValueToString(worksheet.getCell(row, vc)?.value).trim();
    if (!nt) return vc;
    if (isFormXIVPlaceholderCell(nt)) return vc;
    if (isFormXIVLabelContinuationCell(nt, workmanSpecs)) continue;
    if (/^[_\s.-]+$/.test(nt)) return vc;
  }
  return Math.max(fallbackCol, skipUntil + 1);
};

/** Karnataka Form XIV — header labels in B, numbered workman lines in B or C (often with dots). */
export function detectFormXIVKarnatakaWorksheetLayout(worksheet, maxScanRow = 28) {
  if (!worksheet) return false;
  let headerInB = 0;
  let numberedWorkman = 0;
  let dottedLeaders = 0;
  let kaWording = 0;
  let formXivTitle = false;
  let consecutiveHeaders = 0;
  let maxConsecutiveHeaders = 0;
  for (let r = 1; r <= maxScanRow; r += 1) {
    const colA = formXIVMPExcelCellValueToString(worksheet.getCell(r, 1)?.value).trim();
    const colB = formXIVMPExcelCellValueToString(worksheet.getCell(r, 2)?.value).trim();
    const colC = formXIVMPExcelCellValueToString(worksheet.getCell(r, 3)?.value).trim();
    const colD = formXIVMPExcelCellValueToString(worksheet.getCell(r, 4)?.value).trim();
    if (
      /form\s*xiv/i.test(`${colB} ${colC} ${colD}`) ||
      /employment\s+card/i.test(`${colB} ${colC} ${colD}`)
    ) {
      formXivTitle = true;
    }
    if (FORM_XIV_MP_HEADER_SPECS.some((spec) => spec.match.test(colB))) headerInB += 1;
    const headerLike =
      !/^\d+[\.\)]?\s+/.test(colB) &&
      (FORM_XIV_KA_HEADER_OFFICIAL.some((spec) => spec.match.test(colB)) ||
        /^(name|nature)\b/i.test(colB));
    if (headerLike) {
      consecutiveHeaders += 1;
      maxConsecutiveHeaders = Math.max(maxConsecutiveHeaders, consecutiveHeaders);
    } else {
      consecutiveHeaders = 0;
    }
    if (/^\d+[\.\)]?$/.test(colA) && /\.{4,}/.test(colB)) numberedWorkman += 1;
    [colB, colC].forEach((cell) => {
      const ord = cell.match(/^(\d+)[\.\)]?\s+/);
      if (ord && Number(ord[1]) >= 1 && Number(ord[1]) <= 7) {
        numberedWorkman += 1;
      }
    });
    for (let c = 1; c <= 8; c += 1) {
      const cell = formXIVMPExcelCellValueToString(worksheet.getCell(r, c)?.value).trim();
      if (/\.{4,}/.test(cell)) dottedLeaders += 1;
    }
    if (
      /nature\s+and\s+location\s+of\s+work/i.test(colB) ||
      /name\s+and\s+address\s+if\s+contractor/i.test(colB) ||
      /serial\s+no\.?\s+in\s+the\s+register\s+of\s+workmen/i.test(`${colB} ${colC}`)
    ) {
      kaWording += 1;
    }
  }
  // Official MP card: FORM XIV titles live in column F, not B–D. Do not treat it as Karnataka.
  let titleInF = false;
  for (let r = 1; r <= 5; r += 1) {
    const colF = formXIVMPExcelCellValueToString(worksheet.getCell(r, 6)?.value).trim();
    if (isFormXIVMpTitleText(colF)) titleInF = true;
  }
  if (titleInF && !formXivTitle) return false;
  if (numberedWorkman >= 3 && (dottedLeaders >= 1 || kaWording >= 1)) return true;
  // Truncated KA drafts: FORM XIV title in B–D plus consecutive Name/Nature header lines.
  return (
    formXivTitle &&
    maxConsecutiveHeaders >= 3 &&
    (dottedLeaders >= 1 || kaWording >= 1 || numberedWorkman >= 1 || headerInB >= 1)
  );
}

const isFormXIVStackedVariant = (variant) => {
  const v = String(variant || '').toLowerCase();
  return v === 'mp' || v === 'ka';
};

export function isFormXIVKarnatakaExport(worksheet, parsedFormHeader = null) {
  const variant = String(parsedFormHeader?.formXIVVariant || '').toLowerCase();
  if (variant === 'ka') return true;
  if (variant === 'mp' || variant === 'gj' || variant === 'rj') return false;
  return detectFormXIVKarnatakaWorksheetLayout(worksheet);
}

const resolveFormXIVStackedValueCol = (worksheet, parsedFormHeader = null) =>
  isFormXIVKarnatakaExport(worksheet, parsedFormHeader)
    ? FORM_XIV_KA_VALUE_COL
    : FORM_XIV_MP_STACKED_VALUE_COL;

const unmergeFormXIVRange = (worksheet, r1, c1, r2, c2) => {
  if (!worksheet) return;
  // ExcelJS API is unMergeCells (camel-case M). unmergeCells is not a method.
  try {
    if (typeof worksheet.unMergeCells === 'function') {
      worksheet.unMergeCells(r1, c1, r2, c2);
    }
  } catch (_err) {
    /* ignore already-unmerged */
  }
  parseExcelJsMergeRanges(worksheet).forEach((m) => {
    if (m.bottom < r1 || m.top > r2 || m.right < c1 || m.left > c2) return;
    try {
      if (typeof worksheet.unMergeCells === 'function') {
        worksheet.unMergeCells(m.top, m.left, m.bottom, m.right);
      } else if (typeof worksheet.unmergeCells === 'function') {
        worksheet.unmergeCells(m.top, m.left, m.bottom, m.right);
      }
    } catch (_err) {
      /* ignore already-unmerged */
    }
  });
};

const stripFormXIVKALeader = (text) =>
  String(text || '')
    .replace(/\s+\.{4,}.*$/u, '')
    .replace(/[\s.·…_]+$/u, '')
    .trim();

const stripFormXIVKAOrdinal = (text) =>
  String(text || '')
    .replace(/^\d+[\.\)]?\s*/, '')
    .trim();

const formatFormXIVKAFilledLine = (label, value, withDots = true) => {
  // Keep the official ordinal in the label cell (original KA template: "1 Name of the Workman ....").
  const clean = stripFormXIVKALeader(label);
  const val = String(value || '').trim();
  if (!withDots) return val ? `${clean}  ${val}` : clean;
  if (!val) return `${clean} ${FORM_XIV_KA_DOT_LEADER}`;
  return `${clean} ${FORM_XIV_KA_DOT_LEADER} ${val}`;
};

const FORM_XIV_KA_BODY_FONT = {
  name: 'Calibri',
  size: 11,
  bold: false,
  italic: false,
  underline: false,
};

const FORM_XIV_KA_BODY_ALIGNMENT = {
  horizontal: 'left',
  vertical: 'middle',
  wrapText: false,
  shrinkToFit: false,
  indent: 0,
  readingOrder: 'ltr',
};

const styleFormXIVKABodyCell = (cell) => {
  if (!cell) return;
  // Replace the whole style so ExcelJS does not keep the draft/template xf (bold + center).
  cell.style = {
    font: { ...FORM_XIV_KA_BODY_FONT },
    alignment: { ...FORM_XIV_KA_BODY_ALIGNMENT },
  };
  cell.font = { ...FORM_XIV_KA_BODY_FONT };
  cell.alignment = { ...FORM_XIV_KA_BODY_ALIGNMENT };
};

const styleFormXIVKABodyRow = (worksheet, row) => {
  if (!worksheet || row < 1) return;
  unmergeFormXIVRange(worksheet, row, 1, row, 16);
  const excelRow = worksheet.getRow(row);
  excelRow.font = { ...FORM_XIV_KA_BODY_FONT };
  excelRow.alignment = { ...FORM_XIV_KA_BODY_ALIGNMENT };
  for (let c = 1; c <= 8; c += 1) {
    styleFormXIVKABodyCell(worksheet.getCell(row, c));
  }
};

const extractFormXIVKAInlineValue = (raw, official = '') => {
  const text = String(raw || '').trim();
  if (!text) return '';
  const dotted = text.match(/\.{4,}\s+(.+)$/);
  if (dotted) return String(dotted[1] || '').trim();

  const officialBare = stripFormXIVKAOrdinal(String(official || '').replace(/:+\s*$/, '')).trim();
  if (officialBare) {
    const escaped = officialBare.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const afterOfficial = text
      .replace(new RegExp(`^${escaped}\\s*:?\\s+`, 'i'), '')
      .replace(/^[.\s]+/, '')
      .trim();
    if (afterOfficial && afterOfficial.toLowerCase() !== text.toLowerCase()) {
      const officialNorm = officialBare.toLowerCase();
      if (!(officialNorm.startsWith(afterOfficial.toLowerCase()) && afterOfficial.length < officialNorm.length)) {
        return afterOfficial;
      }
    }
  }

  const twoSpace = text.match(/^.{12,}?\s{2,}(\S.*)$/);
  if (twoSpace) {
    const maybe = String(twoSpace[1] || '').replace(/^[.\s]+/, '').trim();
    if (maybe && !FORM_XIV_KA_HEADER_OFFICIAL.some((spec) => spec.match.test(maybe))) {
      return maybe;
    }
  }
  return '';
};

const resolveFormXIVKAOfficialLabel = (raw) => {
  const text = String(raw || '').trim();
  if (!text) return { official: '', withDots: false };
  const headerSpec = FORM_XIV_KA_HEADER_OFFICIAL.find(
    (spec) => spec.match.test(text) || spec.match.test(formXIVMPHeaderNorm(text))
  );
  if (headerSpec) {
    return { official: headerSpec.text, withDots: false };
  }
  const ordMatch = text.match(/^(\d+)[\.\)]?\s*/);
  const ordinal = ordMatch ? Number(ordMatch[1]) : 0;
  if (ordinal >= 1 && ordinal <= FORM_XIV_KA_WORKMAN_OFFICIAL.length) {
    return {
      official: FORM_XIV_KA_WORKMAN_OFFICIAL[ordinal - 1],
      withDots: true,
      ordinal,
    };
  }
  const bare = stripFormXIVKAOrdinal(stripFormXIVKALeader(text)).toLowerCase();
  const byBare = FORM_XIV_KA_WORKMAN_OFFICIAL.findIndex((label) => {
    const officialBare = stripFormXIVKAOrdinal(label).toLowerCase();
    return bare === officialBare || bare.startsWith(officialBare) || officialBare.startsWith(bare);
  });
  if (byBare >= 0 && bare.length >= 6) {
    return {
      official: FORM_XIV_KA_WORKMAN_OFFICIAL[byBare],
      withDots: true,
      ordinal: byBare + 1,
    };
  }
  return { official: '', withDots: false, ordinal: 0 };
};

const findFormXIVKAHeaderLabelRows = (worksheet) => {
  const rows = [];
  if (!worksheet) return rows;
  for (let r = 1; r <= 20 && rows.length < 4; r += 1) {
    for (let c = 1; c <= 8; c += 1) {
      const raw = formXIVMPExcelCellValueToString(worksheet.getCell(r, c)?.value).trim();
      if (!raw || /^\d+[\.\)]?\s*/.test(raw)) continue;
      if (/form\s*xiv|employment\s+card|see\s+rule\s*76|signature/i.test(raw)) continue;
      const isHeader =
        FORM_XIV_KA_HEADER_OFFICIAL.some(
          (spec) => spec.match.test(raw) || spec.match.test(formXIVMPHeaderNorm(raw))
        ) || /^(name|nature)\b/i.test(stripFormXIVKALeader(raw));
      if (!isHeader) continue;
      rows.push({ row: r, col: c, raw });
      break;
    }
  }
  return rows;
};

const isFormXIVMergeSlaveCell = (worksheet, row, col) => {
  if (!worksheet || row < 1 || col < 1) return false;
  return parseExcelJsMergeRanges(worksheet).some(
    (m) =>
      row >= m.top &&
      row <= m.bottom &&
      col >= m.left &&
      col <= m.right &&
      (row !== m.top || col !== m.left)
  );
};

const mergeFormXIVKALabelLine = (worksheet, row, labelCol = FORM_XIV_KA_LABEL_COL) => {
  if (!worksheet || row < 1 || labelCol < 1) return;
  unmergeFormXIVRange(worksheet, row, 1, row, FORM_XIV_KA_LINE_MERGE_END);
  if (labelCol >= FORM_XIV_KA_LINE_MERGE_END) return;
  try {
    worksheet.mergeCells(row, labelCol, row, FORM_XIV_KA_LINE_MERGE_END);
  } catch (_err) {
    /* ignore */
  }
};

/** Official KA templates put labels in A or B — never leave them right-aligned in a 3-wide column A. */
const resolveFormXIVKABodyLabelCol = (worksheet) => {
  const found = findFormXIVKAHeaderLabelRows(worksheet);
  const counts = { 1: 0, 2: 0, 3: 0 };
  found.forEach((pos) => {
    if (pos.col >= 1 && pos.col <= 3) counts[pos.col] += 1;
  });
  if (counts[1] >= 2 && counts[1] >= counts[2]) return 1;
  if (counts[2] >= 1) return 2;
  for (let r = 1; r <= 20; r += 1) {
    const colA = formXIVMPExcelCellValueToString(worksheet.getCell(r, 1)?.value).trim();
    const colB = formXIVMPExcelCellValueToString(worksheet.getCell(r, 2)?.value).trim();
    if (/^\d+[\.\)]?\s+/.test(colA) && /workman|serial|nature|wage|tenure|remark/i.test(colA)) {
      return 1;
    }
    if (/^\d+[\.\)]?\s+/.test(colB) && /workman|serial|nature|wage|tenure|remark/i.test(colB)) {
      return 2;
    }
    if (
      FORM_XIV_KA_HEADER_OFFICIAL.some((spec) => spec.match.test(colA)) ||
      /^(name|nature)\b/i.test(colA)
    ) {
      return 1;
    }
  }
  return FORM_XIV_KA_LABEL_COL;
};

const isFormXIVKAStrayHeaderValue = (raw) => {
  const text = String(raw || '').trim();
  if (!text) return false;
  if (/^m\/s\b/i.test(text) && /pvt|ltd|karnataka|village|dist|substation/i.test(text)) return true;
  if (text.length > 60 && /karnataka|pvt|ltd|village|dist/i.test(text)) return true;
  return false;
};

const writeFormXIVKASerialAndLabel = (
  worksheet,
  row,
  official,
  value,
  ordinal = 0,
  withDots = true,
  labelCol = FORM_XIV_KA_LABEL_COL
) => {
  const col = labelCol >= 1 && labelCol <= 3 ? labelCol : FORM_XIV_KA_LABEL_COL;
  unmergeFormXIVRange(worksheet, row, 1, row, 16);
  for (let c = 1; c <= 8; c += 1) {
    const other = formXIVMPExcelCellValueToString(worksheet.getCell(row, c)?.value).trim();
    if (
      /signature/i.test(other) ||
      /^form\s*xiv/i.test(other) ||
      /see\s+rule\s*76/i.test(other) ||
      /^\s*employment\s+card\s*$/i.test(other)
    ) {
      continue;
    }
    worksheet.getCell(row, c).value = '';
  }
  const labelCell = worksheet.getCell(row, col);
  labelCell.value = formatFormXIVKAFilledLine(official, value, withDots);
  styleFormXIVKABodyRow(worksheet, row);
  styleFormXIVKABodyCell(labelCell);
};

const writeFormXIVKAInlineOnLabelRow = (worksheet, row, labelCol, value, officialFallback = '') => {
  if (!worksheet || row < 1 || labelCol < 1) return;
  const raw = formXIVMPExcelCellValueToString(worksheet.getCell(row, labelCol)?.value).trim();
  const resolved = resolveFormXIVKAOfficialLabel(raw || officialFallback);
  const official = resolved.official || officialFallback || stripFormXIVKALeader(raw);
  if (!official) return;
  const ordinal =
    resolved.ordinal ||
    Number(String(officialFallback || official).match(/^(\d+)/)?.[1] || 0);
  writeFormXIVKASerialAndLabel(worksheet, row, official, value, ordinal, ordinal >= 1, labelCol);
};

const applyFormXIVKATitleAlignment = (worksheet) => {
  const titleSpecs = [
    { match: /^\s*form\s*xiv\s*$/i, bold: true, italic: false },
    { match: /see\s+rule\s*76/i, bold: false, italic: true },
    { match: /^\s*employment\s+card\s*$/i, bold: true, italic: false },
  ];
  titleSpecs.forEach((spec) => {
    for (let r = 1; r <= 8; r += 1) {
      for (let c = 1; c <= 8; c += 1) {
        const raw = formXIVMPExcelCellValueToString(worksheet.getCell(r, c)?.value).trim();
        if (!raw || !spec.match.test(raw)) continue;
        unmergeFormXIVRange(worksheet, r, 1, r, FORM_XIV_KA_VALUE_COL);
        for (let clearCol = 1; clearCol <= FORM_XIV_KA_VALUE_COL; clearCol += 1) {
          if (clearCol === FORM_XIV_KA_TITLE_COL) continue;
          const other = formXIVMPExcelCellValueToString(
            worksheet.getCell(r, clearCol)?.value
          ).trim();
          if (other && !spec.match.test(other) && clearCol !== c) continue;
          worksheet.getCell(r, clearCol).value = '';
        }
        const titleCell = worksheet.getCell(r, FORM_XIV_KA_TITLE_COL);
        titleCell.value = raw;
        if (c !== FORM_XIV_KA_TITLE_COL) worksheet.getCell(r, c).value = '';
        titleCell.style = {
          font: {
            name: 'Calibri',
            size: 12,
            bold: spec.bold,
            italic: spec.italic,
          },
          alignment: {
            horizontal: 'center',
            vertical: 'middle',
            wrapText: false,
          },
        };
        titleCell.font = {
          name: 'Calibri',
          size: 12,
          bold: spec.bold,
          italic: spec.italic,
        };
        titleCell.alignment = {
          horizontal: 'center',
          vertical: 'middle',
          wrapText: false,
        };
        try {
          worksheet.mergeCells(r, FORM_XIV_KA_TITLE_COL, r, FORM_XIV_KA_TITLE_MERGE_END);
        } catch (_err) {
          /* ignore */
        }
        return;
      }
    }
  });
};

const applyFormXIVKASignatureAlignment = (worksheet) => {
  for (let r = 16; r <= 32; r += 1) {
    for (let c = 1; c <= 12; c += 1) {
      const raw = formXIVMPExcelCellValueToString(worksheet.getCell(r, c)?.value).trim();
      if (!/signature\s+of\s+contractor/i.test(raw)) continue;
      if (c !== FORM_XIV_KA_VALUE_COL) {
        worksheet.getCell(r, FORM_XIV_KA_VALUE_COL).value = raw;
        worksheet.getCell(r, c).value = '';
      }
      worksheet.getCell(r, FORM_XIV_KA_VALUE_COL).alignment = {
        horizontal: 'center',
        vertical: 'top',
        wrapText: false,
      };
      const prevRaw = formXIVMPExcelCellValueToString(
        worksheet.getCell(r - 1, c)?.value
      ).trim();
      const prevIsLeader =
        isFormXIVPlaceholderCell(prevRaw) || /^\.{5,}/.test(prevRaw) || /^_{5,}/.test(prevRaw);
      if (prevIsLeader) {
        if (c !== FORM_XIV_KA_VALUE_COL) {
          worksheet.getCell(r - 1, FORM_XIV_KA_VALUE_COL).value = prevRaw || '.'.repeat(28);
          worksheet.getCell(r - 1, c).value = '';
        }
        worksheet.getCell(r - 1, FORM_XIV_KA_VALUE_COL).alignment = {
          horizontal: 'center',
          vertical: 'bottom',
          wrapText: false,
        };
      }
      return;
    }
  }
};

const applyFormXIVKAFilledLineToRow = (worksheet, row, labelCol, official, withDots, ordinal = 0) => {
  if (!worksheet || row < 1 || !official) return;
  const col = labelCol >= 1 && labelCol <= 3 ? labelCol : FORM_XIV_KA_LABEL_COL;
  unmergeFormXIVRange(worksheet, row, 1, row, 16);
  let movedValue = '';
  for (let c = 1; c <= 8; c += 1) {
    const raw = formXIVMPExcelCellValueToString(worksheet.getCell(row, c)?.value).trim();
    if (!raw) continue;
    if (/^\d+[\.\)]?$/.test(raw)) continue;
    if (/signature/i.test(raw) || /^form\s*xiv/i.test(raw) || /employment\s+card/i.test(raw)) continue;
    if (isFormXIVPlaceholderCell(raw)) continue;
    // When labels live in B/C, leftover addresses in column A are stray values.
    if (ordinal >= 1 && col !== 1 && (c === 1 || isFormXIVKAStrayHeaderValue(raw))) continue;
    const extracted = extractFormXIVKAInlineValue(raw, official);
    if (extracted) {
      if (ordinal >= 1 && isFormXIVKAStrayHeaderValue(extracted)) continue;
      movedValue = extracted;
      continue;
    }
    if (isFormXIVLabelContinuationCell(raw) || resolveFormXIVKAOfficialLabel(raw).official) continue;
    if (ordinal >= 1 && isFormXIVKAStrayHeaderValue(raw)) continue;
    movedValue = raw;
  }
  const resolvedOrdinal =
    ordinal ||
    Number(String(official).match(/^(\d+)/)?.[1] || 0);
  writeFormXIVKASerialAndLabel(
    worksheet,
    row,
    official,
    movedValue,
    resolvedOrdinal,
    resolvedOrdinal >= 1 || withDots,
    col
  );
  const excelRow = worksheet.getRow(row);
  excelRow.height = Math.max(Number(excelRow.height) || 0, 18);
};

const findFormXIVKATitleBottomRow = (worksheet) => {
  let bottom = 4;
  if (!worksheet) return bottom;
  for (let r = 1; r <= 8; r += 1) {
    for (let c = 1; c <= 8; c += 1) {
      const raw = formXIVMPExcelCellValueToString(worksheet.getCell(r, c)?.value).trim();
      if (
        /^\s*form\s*xiv\s*$/i.test(raw) ||
        /^\s*employment\s+card\s*$/i.test(raw) ||
        /see\s+rule\s*76/i.test(raw)
      ) {
        bottom = Math.max(bottom, r);
      }
    }
  }
  return bottom;
};

const resolveFormXIVKAHeaderTargetRows = (worksheet) => {
  const found = findFormXIVKAHeaderLabelRows(worksheet)
    .map((pos) => pos.row)
    .filter((row, index, all) => all.indexOf(row) === index)
    .slice(0, 4);
  if (found.length >= 4) return found;
  const start = findFormXIVKATitleBottomRow(worksheet) + 2;
  return [start, start + 1, start + 2, start + 3];
};

const applyFormXIVKADefaultSignature = (worksheet, afterRow) => {
  let hasSignature = false;
  for (let r = 16; r <= 32; r += 1) {
    for (let c = 1; c <= 12; c += 1) {
      const raw = formXIVMPExcelCellValueToString(worksheet.getCell(r, c)?.value).trim();
      if (/signature\s+of\s+contractor/i.test(raw)) {
        hasSignature = true;
        break;
      }
    }
    if (hasSignature) break;
  }
  if (hasSignature) return;
  const sigRow = Math.max(19, Number(afterRow) + 2);
  worksheet.getCell(sigRow, FORM_XIV_KA_VALUE_COL).value = '.'.repeat(28);
  worksheet.getCell(sigRow + 1, FORM_XIV_KA_VALUE_COL).value = 'Signature of Contractor';
};

/** Restore official KA template: full labels left-aligned, titles in C, signature in F. */
const ensureFormXIVKarnatakaColumnAlignment = (worksheet) => {
  if (!worksheet) return;
  const labelCol = resolveFormXIVKABodyLabelCol(worksheet);
  const minWidths = {
    1: labelCol === 1 ? 100 : 5,
    2: labelCol === 2 ? 100 : 14,
    3: 14,
    4: 14,
    5: 14,
    6: 22,
    7: 10,
    8: 10,
  };
  Object.entries(minWidths).forEach(([col, minW]) => {
    const column = worksheet.getColumn(Number(col));
    column.width = Math.max(Number(column.width) || 0, minW);
  });
  worksheet.getColumn(labelCol).width = Math.max(Number(worksheet.getColumn(labelCol).width) || 0, 100);

  applyFormXIVKATitleAlignment(worksheet);

  const headerTargetRows = resolveFormXIVKAHeaderTargetRows(worksheet);
  headerTargetRows.forEach((row, index) => {
    const official = FORM_XIV_KA_HEADER_OFFICIAL[index]?.text;
    if (!official) return;
    applyFormXIVKAFilledLineToRow(worksheet, row, labelCol, official, false, 0);
  });

  const lastHeader = headerTargetRows[headerTargetRows.length - 1] || 9;
  const workmanStart = lastHeader + 2;
  FORM_XIV_KA_WORKMAN_OFFICIAL.forEach((official, index) => {
    applyFormXIVKAFilledLineToRow(
      worksheet,
      workmanStart + index,
      labelCol,
      official,
      true,
      index + 1
    );
  });

  applyFormXIVKADefaultSignature(worksheet, workmanStart + FORM_XIV_KA_WORKMAN_OFFICIAL.length - 1);
  applyFormXIVKASignatureAlignment(worksheet);

  const keepLabelRows = new Set([
    ...headerTargetRows,
    ...FORM_XIV_KA_WORKMAN_OFFICIAL.map((_, i) => workmanStart + i),
  ]);
  for (let r = 5; r <= 22; r += 1) {
    let skipTitle = false;
    for (let c = 1; c <= 8; c += 1) {
      const raw = formXIVMPExcelCellValueToString(worksheet.getCell(r, c)?.value).trim();
      if (
        /^\s*form\s*xiv\s*$/i.test(raw) ||
        /^\s*employment\s+card\s*$/i.test(raw) ||
        /see\s+rule\s*76/i.test(raw)
      ) {
        skipTitle = true;
        break;
      }
    }
    if (skipTitle) continue;
    if (keepLabelRows.has(r)) {
      styleFormXIVKABodyRow(worksheet, r);
      styleFormXIVKABodyCell(worksheet.getCell(r, labelCol));
      continue;
    }
    const label = formXIVMPExcelCellValueToString(worksheet.getCell(r, labelCol)?.value).trim();
    const serial = formXIVMPExcelCellValueToString(worksheet.getCell(r, 1)?.value).trim();
    if (!label && !serial) continue;
    if (/signature/i.test(label)) continue;
    styleFormXIVKABodyRow(worksheet, r);
  }
  // Titles last so body left-align cannot pull FORM XIV / Employment Card back to column B.
  applyFormXIVKATitleAlignment(worksheet);
};

/** Re-apply KA Employment Card layout after download finalize (system-note pass). */
export function finalizeFormXIVKarnatakaWorksheet(worksheet, parsedFormHeader = null) {
  if (!worksheet) return false;
  if (!isFormXIVKarnatakaExport(worksheet, parsedFormHeader)) return false;
  ensureFormXIVKarnatakaColumnAlignment(worksheet);
  return true;
}

export const detectFormXIVStackedWorkmanLayout = (worksheet, maxScanRow = 45) => {
  if (!worksheet) return false;
  if (detectFormXIVKarnatakaWorksheetLayout(worksheet, maxScanRow)) return true;
  let numberedLabelHits = 0;
  let splitOrdinalHits = 0;
  for (let r = 1; r <= maxScanRow; r += 1) {
    for (let c = 1; c <= 4; c += 1) {
      const cellStr = formXIVMPExcelCellValueToString(worksheet.getCell(r, c)?.value).trim();
      if (!cellStr || !/^\d+[\.\)]?\s+/.test(cellStr)) continue;
      const norm = formXIVMPHeaderNorm(cellStr);
      if (isFormXIVStackedWorkmanKeyword(norm)) {
        numberedLabelHits += 1;
        break;
      }
    }
    if (isFormXIVSplitOrdinalLabelRow(worksheet, r)) {
      splitOrdinalHits += 1;
    }
  }
  if (Math.max(numberedLabelHits, splitOrdinalHits) < 3) return false;

  for (let r = 1; r <= maxScanRow; r += 1) {
    const labelA = formXIVMPExcelCellValueToString(worksheet.getCell(r, 1)?.value).trim();
    const labelB = formXIVMPExcelCellValueToString(worksheet.getCell(r, 2)?.value).trim();
    const labelC = formXIVMPExcelCellValueToString(worksheet.getCell(r, 3)?.value).trim();
    const combinedInA = /^1[\.\)]\s*/.test(labelA) && /workman/i.test(labelA);
    const splitRow1 = /^1[\.\)]?$/.test(labelA) && /workman/i.test(labelB);
    const combinedInB = /^1[\.\)]\s*/.test(labelB) && /workman/i.test(labelB);
    const combinedInC = /^1[\.\)]\s*/.test(labelC) && /workman/i.test(labelC);
    if (!combinedInA && !splitRow1 && !combinedInB && !combinedInC) continue;

    const colE = formXIVMPExcelCellValueToString(
      worksheet.getCell(r, FORM_XIV_MP_STACKED_VALUE_COL)?.value
    ).trim();
    if (!colE || isFormXIVPlaceholderCell(colE)) return true;

    // Split ordinal templates keep labels in B–D; values belong in E even when C looks empty.
    if (splitOrdinalHits >= 3) return true;

    const colB = formXIVMPExcelCellValueToString(
      worksheet.getCell(r, FORM_XIV_GJ_DEFAULT_VALUE_COL)?.value
    ).trim();
    if (!colB || isFormXIVPlaceholderCell(colB)) return false;
  }
  return true;
};

const FORM_XIV_MP_SINGLE_LINE_ALIGNMENT = {
  horizontal: 'left',
  vertical: 'middle',
  wrapText: false,
  shrinkToFit: false,
};

/** Official MP Form XIV (Rule 76) labels — one line, matching the blank template. */
const FORM_XIV_MP_HEADER_OFFICIAL = [
  {
    match: /name\s+and\s+address\s+(?:of|if)\s+(?:the\s+)?contractor/i,
    text: 'Name and address of contractor',
  },
  {
    match: /establishment/i,
    text: 'Name and address of establishment in / under which contract is carried on',
  },
  {
    match: /nature\s+(?:of\s+work\s+and\s+location|and\s+location\s+of\s+work)/i,
    text: 'Nature of work and location of work',
  },
  {
    match: /principal\s+employer/i,
    text: 'Name and address of principal employer',
  },
];

const isFormXIVMpTitleText = (text) =>
  /^\s*form\s*xiv\s*$/i.test(text) ||
  /^\s*employment\s+card\s*$/i.test(text) ||
  /see\s+rule\s*76|vide\s+rule\s*76/i.test(text);

const FORM_XIV_MP_BODY_FONT = {
  name: 'Calibri',
  size: 12,
  bold: false,
  italic: false,
};

const styleFormXIVMPSingleLineLabelCell = (cell) => {
  if (!cell) return;
  const nextStyle = {
    font: { ...FORM_XIV_MP_BODY_FONT, ...(cell.font || {}), bold: false, italic: false },
    alignment: { ...FORM_XIV_MP_SINGLE_LINE_ALIGNMENT },
  };
  if (cell.border) nextStyle.border = { ...cell.border };
  if (cell.fill) nextStyle.fill = { ...cell.fill };
  if (cell.numFmt) nextStyle.numFmt = cell.numFmt;
  // Replace xf so template wrapText / bold cannot survive ExcelJS rewrite.
  cell.style = nextStyle;
  cell.font = { ...nextStyle.font };
  cell.alignment = { ...FORM_XIV_MP_SINGLE_LINE_ALIGNMENT };
};

const styleFormXIVMPBodyValueCell = (cell, wrapText = false) => {
  if (!cell) return;
  const alignment = {
    horizontal: 'left',
    vertical: 'middle',
    wrapText,
    shrinkToFit: false,
  };
  const font = { ...FORM_XIV_MP_BODY_FONT, ...(cell.font || {}), bold: false, italic: false };
  const nextStyle = { font, alignment };
  if (cell.border) nextStyle.border = { ...cell.border };
  if (cell.fill) nextStyle.fill = { ...cell.fill };
  if (cell.numFmt) nextStyle.numFmt = cell.numFmt;
  cell.style = nextStyle;
  cell.font = { ...font };
  cell.alignment = { ...alignment };
};

const resolveFormXIVMPOfficialLabelText = (raw, ordinal = 0) => {
  const text = normalizeFormXIVMPLabelText(raw);
  if (!text || isFormXIVMpTitleText(text)) return text;
  const hasOrdinalPrefix = /^\d+[\.\)]\s*/.test(text);
  const bare = text.replace(/^\d+[\.\)]\s*/, '');
  const resolvedOrdinal =
    ordinal || Number(String(text).match(/^(\d+)[\.\)]/)?.[1] || 0);
  if (resolvedOrdinal >= 1 && resolvedOrdinal <= FORM_XIV_MP_WORKMAN_FIELD_SPECS.length) {
    const official = FORM_XIV_MP_WORKMAN_FIELD_SPECS[resolvedOrdinal - 1].label;
    const next = hasOrdinalPrefix ? `${resolvedOrdinal} ${official}` : official;
    if (bare.length + 8 < official.length) return next;
    return hasOrdinalPrefix ? `${resolvedOrdinal} ${normalizeFormXIVMPLabelText(bare)}` : text;
  }
  const headerSpec = FORM_XIV_MP_HEADER_OFFICIAL.find((spec) => spec.match.test(text));
  if (headerSpec && (bare.length + 12 < headerSpec.text.length || /[\r\n]/.test(String(raw || '')))) {
    return headerSpec.text;
  }
  return text;
};

const isFormXIVMPHeaderLabelCandidate = (raw) => {
  const text = normalizeFormXIVMPLabelText(raw);
  if (!text || isFormXIVMpTitleText(text)) return false;
  if (/^\d+[\.\)]\s*/.test(text)) return false;
  return (
    FORM_XIV_MP_HEADER_OFFICIAL.some((spec) => spec.match.test(text)) ||
    FORM_XIV_MP_HEADER_SPECS.some((spec) => spec.match.test(text)) ||
    /^name\b/i.test(text) ||
    /^nature\b/i.test(text)
  );
};

const findFormXIVMPWorkmanOrdinalOnRow = (worksheet, row) => {
  if (!worksheet || row < 1) return 0;
  const colA = formXIVMPExcelCellValueToString(worksheet.getCell(row, 1)?.value).trim();
  if (/^\d+[\.\)]?$/.test(colA)) return Number(colA.replace(/[.)]/g, ''));
  for (let c = 2; c <= 4; c += 1) {
    const raw = formXIVMPExcelCellValueToString(worksheet.getCell(row, c)?.value).trim();
    const match = raw.match(/^(\d+)[\.\)]?\s+/);
    if (match) return Number(match[1]);
  }
  return 0;
};

const findFormXIVMPHeaderLabelRows = (worksheet) => {
  const rows = [];
  if (!worksheet) return rows;
  for (let r = 1; r <= 40; r += 1) {
    if (findFormXIVMPWorkmanOrdinalOnRow(worksheet, r) >= 1) continue;
    for (let c = 1; c <= 4; c += 1) {
      const raw = formXIVMPExcelCellValueToString(worksheet.getCell(r, c)?.value).trim();
      if (!isFormXIVMPHeaderLabelCandidate(raw)) continue;
      rows.push({ row: r, col: c });
      break;
    }
  }
  return rows.slice(0, FORM_XIV_MP_HEADER_OFFICIAL.length);
};

const findFormXIVMPWorkmanLabelRows = (worksheet) => {
  const rows = [];
  if (!worksheet) return rows;
  for (let r = 1; r <= 45; r += 1) {
    const ordinal = findFormXIVMPWorkmanOrdinalOnRow(worksheet, r);
    if (ordinal < 1 || ordinal > FORM_XIV_MP_WORKMAN_FIELD_SPECS.length) continue;
    const splitOrdinal = /^\d+[\.\)]?$/.test(
      formXIVMPExcelCellValueToString(worksheet.getCell(r, 1)?.value).trim()
    );
    let labelCol = 2;
    let labelRaw = '';
    for (let c = 2; c <= 4; c += 1) {
      const raw = formXIVMPExcelCellValueToString(worksheet.getCell(r, c)?.value).trim();
      if (raw && !/^\d+[\.\)]?$/.test(raw)) {
        labelCol = c;
        labelRaw = raw;
        break;
      }
    }
    const combinedA = formXIVMPExcelCellValueToString(worksheet.getCell(r, 1)?.value).trim();
    const labelForKeyword = labelRaw || combinedA;
    // Form A / other tabular registers also number rows 1–7 in column A.
    // Only treat a row as an Employment Card field when the label is a workman keyword.
    if (!isFormXIVStackedWorkmanKeyword(formXIVMPHeaderNorm(labelForKeyword))) continue;
    rows.push({ row: r, col: labelCol, ordinal, splitOrdinal });
  }
  return rows;
};

/** Rajasthan / Gujarat Form A employee register — must not be rewritten as MP Form XIV. */
const worksheetLooksLikeFormAEmployeeRegister = (worksheet) => {
  if (!worksheet) return false;
  const maxRow = Math.min(10, Number(worksheet.rowCount) || 10);
  let sawFormA = false;
  let sawEmployeeRegister = false;
  let sawRule21 = false;
  let sawSurname = false;
  let sawDob = false;
  let sawEmployeeCode = false;
  for (let r = 1; r <= maxRow; r += 1) {
    for (let c = 1; c <= 12; c += 1) {
      const raw = formXIVMPExcelCellValueToString(worksheet.getCell(r, c)?.value).trim();
      if (!raw) continue;
      if (/see\s+rule\s*2\s*\(\s*1\s*\)/i.test(raw)) sawRule21 = true;
      if (/format\s+of\s+employee/i.test(raw)) sawEmployeeRegister = true;
      if (/^\s*form\s*a\s*$/i.test(raw)) sawFormA = true;
      const n = formXIVMPHeaderNorm(raw);
      if (/surname/.test(n)) sawSurname = true;
      if (/date/.test(n) && (/birth/.test(n) || /dob/.test(n))) sawDob = true;
      if (/employee/.test(n) && (/code/.test(n) || /\bid\b/.test(n))) sawEmployeeCode = true;
    }
  }
  if (sawFormA || sawEmployeeRegister || sawRule21) return true;
  return sawSurname && sawDob && sawEmployeeCode;
};

const applyFormXIVMPOfficialTemplateLabels = (worksheet) => {
  if (!worksheet) return;
  if (worksheetLooksLikeFormAEmployeeRegister(worksheet)) return;
  const assignedRows = new Set();

  FORM_XIV_MP_HEADER_OFFICIAL.forEach((headerSpec) => {
    for (let r = 1; r <= 40; r += 1) {
      if (assignedRows.has(r) || findFormXIVMPWorkmanOrdinalOnRow(worksheet, r) >= 1) continue;
      for (let c = 1; c <= 4; c += 1) {
        const raw = formXIVMPExcelCellValueToString(worksheet.getCell(r, c)?.value).trim();
        if (!raw || isFormXIVMpTitleText(raw)) continue;
        if (!headerSpec.match.test(raw) && !headerSpec.match.test(formXIVMPHeaderNorm(raw))) continue;
        const cell = worksheet.getCell(r, c);
        cell.value = headerSpec.text;
        styleFormXIVMPSingleLineLabelCell(cell);
        assignedRows.add(r);
        return;
      }
    }
  });

  const fallbackHeaderRows = findFormXIVMPHeaderLabelRows(worksheet).filter(
    (pos) => !assignedRows.has(pos.row)
  );
  let fallbackIndex = 0;
  FORM_XIV_MP_HEADER_OFFICIAL.forEach((headerSpec) => {
    const alreadyAssigned = [...assignedRows].some((row) => {
      for (let c = 1; c <= 4; c += 1) {
        const raw = formXIVMPExcelCellValueToString(worksheet.getCell(row, c)?.value).trim();
        if (raw === headerSpec.text) return true;
      }
      return false;
    });
    if (alreadyAssigned) return;
    const pos = fallbackHeaderRows[fallbackIndex];
    if (!pos) return;
    fallbackIndex += 1;
    const cell = worksheet.getCell(pos.row, pos.col);
    cell.value = headerSpec.text;
    styleFormXIVMPSingleLineLabelCell(cell);
    assignedRows.add(pos.row);
  });

  findFormXIVMPWorkmanLabelRows(worksheet).forEach(({ row, col, ordinal, splitOrdinal }) => {
    const spec = FORM_XIV_MP_WORKMAN_FIELD_SPECS[ordinal - 1];
    if (!spec) return;
    if (splitOrdinal) {
      const ordinalCell = worksheet.getCell(row, 1);
      ordinalCell.value = ordinal;
      styleFormXIVMPSingleLineLabelCell(ordinalCell);
      const labelCell = worksheet.getCell(row, col);
      labelCell.value = spec.label;
      styleFormXIVMPSingleLineLabelCell(labelCell);
      return;
    }
    const labelCell = worksheet.getCell(row, col);
    labelCell.value = `${ordinal} ${spec.label}`;
    styleFormXIVMPSingleLineLabelCell(labelCell);
  });
};

const shouldApplyFormXIVMPExportAlignment = (worksheet, parsedFormHeader = null) => {
  if (!worksheet) return false;
  if (worksheetLooksLikeFormAEmployeeRegister(worksheet)) return false;
  if (isFormXIVKarnatakaExport(worksheet, parsedFormHeader)) return false;
  if (detectFormXRajasthanWorksheetLayout(worksheet)) return false;
  if (detectFormXIVGJEmploymentCardWorksheet(worksheet)) return false;
  const variant = String(parsedFormHeader?.formXIVVariant || '').toLowerCase();
  if (variant === 'ka' || variant === 'gj' || variant === 'rj') return false;
  // Only rewrite when this is explicitly MP Form XIV, or the sheet is a stacked Employment Card.
  // Do not default to MP — download finalize runs on every statutory Excel (including Form A RJ).
  if (variant === 'mp') return true;
  return (
    detectFormXIVStackedWorkmanLayout(worksheet) ||
    findFormXIVMPWorkmanLabelRows(worksheet).length >= 2
  );
};

const findFormXIVMPStackedLabelCol = (worksheet, row) => {
  if (!worksheet || row < 1) return 0;
  for (let c = 1; c <= 4; c += 1) {
    const raw = formXIVMPExcelCellValueToString(worksheet.getCell(row, c)?.value).trim();
    if (!raw || /^\d+[\.\)]?$/.test(raw) || isFormXIVMpTitleText(raw)) continue;
    if (
      isFormXIVLabelContinuationCell(raw) ||
      /^\d+[\.\)]\s*/.test(raw) ||
      FORM_XIV_MP_HEADER_SPECS.some((spec) => spec.match.test(raw)) ||
      isFormXIVStackedWorkmanKeyword(formXIVMPHeaderNorm(raw)) ||
      /^name\b/i.test(raw) ||
      /^nature\b/i.test(raw) ||
      /^serial\b/i.test(raw) ||
      /^wage/i.test(raw) ||
      /^tenure\b/i.test(raw) ||
      /^remark/i.test(raw)
    ) {
      return c;
    }
  }
  return 0;
};

/**
 * MP Form XIV only — keep original-template labels on one fully visible line.
 * Official even-row card: titles in F, all labels in B, values in C,
 * blank row between each line. Does not run for Karnataka / GJ / RJ.
 */
const FORM_XIV_MP_TITLE_COL = 3;
const FORM_XIV_MP_HEADER_LABEL_COL = 2;
const FORM_XIV_MP_OFFICIAL_WORKMAN_COL = 2;
const FORM_XIV_MP_OFFICIAL_VALUE_COL = 3;
const FORM_XIV_MP_OFFICIAL_HEADER_ROWS = [6, 8, 10, 12];
const FORM_XIV_MP_OFFICIAL_WORKMAN_ROWS = [14, 16, 18, 20, 22, 24, 26];

const detectFormXIVMPColumnCWorkmanLayout = (worksheet) => {
  if (!worksheet) return false;
  let splitOrdinalHits = 0;
  let colCHits = 0;
  for (let r = 12; r <= 28; r += 1) {
    if (isFormXIVSplitOrdinalLabelRow(worksheet, r)) splitOrdinalHits += 1;
    const colC = formXIVMPExcelCellValueToString(worksheet.getCell(r, 3)?.value).trim();
    if (
      /^\d+[\.\)]?\s+/.test(colC) &&
      isFormXIVStackedWorkmanKeyword(formXIVMPHeaderNorm(colC))
    ) {
      colCHits += 1;
    }
  }
  return colCHits >= 2 && splitOrdinalHits < 3;
};

/** Gujarat Form XIV boxed card — must never be rewritten as MP even-row stacked layout. */
const detectFormXIVGJEmploymentCardWorksheet = (worksheet, maxScanRow = 40) => {
  if (!worksheet) return false;
  let gujaratMarker = false;
  let entryIntoService = false;
  let signatureContractor = false;
  let boxedWorkmanOrdinal = false;
  for (let r = 1; r <= maxScanRow; r += 1) {
    for (let c = 1; c <= 14; c += 1) {
      const raw = formXIVMPExcelCellValueToString(worksheet.getCell(r, c)?.value).trim();
      if (!raw) continue;
      const n = formXIVMPHeaderNorm(raw);
      if (/central\s*&\s*gujarat|gujarat\s+rules|xiv[\s._-]*gj/.test(n)) gujaratMarker = true;
      if (/vide\s+rule\s+76/.test(n) && /(?:gujarat|contract\s+labou?r)/.test(n)) {
        gujaratMarker = true;
      }
      if (/entry\s+into\s+service/.test(n)) entryIntoService = true;
      if (/signature\s+of\s+the\s+contractor/.test(n)) signatureContractor = true;
      if (/^\d+[\.\)]\s+name\s+of\s+the\s+workman/.test(n)) boxedWorkmanOrdinal = true;
    }
  }
  // GJ template: Vide Rule 76 Central & Gujarat Rules and/or field "Date of entry into service".
  return (
    gujaratMarker ||
    entryIntoService ||
    (signatureContractor && boxedWorkmanOrdinal)
  );
};

const FORM_XIV_GJ_OFFICIAL_THIN_BORDER = { style: 'thin', color: { argb: 'FF000000' } };
const FORM_XIV_GJ_OFFICIAL_WORKMAN_LABELS = [
  '1. Name of the Workman',
  '2. S.No. in the Register of Workmen Employed',
  '3. Nature of Employment /Designation',
  '4. Date of entry into service',
  '5. Wage rate (With particulars of unit in case of piece - work)',
  '6. Wage period',
  '7. Tenure of Employment',
  '8. Remarks',
];

const applyFormXIVGJBoxBorders = (worksheet, top, left, bottom, right) => {
  for (let r = top; r <= bottom; r += 1) {
    for (let c = left; c <= right; c += 1) {
      const cell = worksheet.getCell(r, c);
      cell.border = {
        top: r === top ? FORM_XIV_GJ_OFFICIAL_THIN_BORDER : undefined,
        bottom: r === bottom ? FORM_XIV_GJ_OFFICIAL_THIN_BORDER : undefined,
        left: c === left ? FORM_XIV_GJ_OFFICIAL_THIN_BORDER : undefined,
        right: c === right ? FORM_XIV_GJ_OFFICIAL_THIN_BORDER : undefined,
      };
    }
  }
};

/**
 * Rebuild official Gujarat Form XIV Employment Card (boxed 2×2 header + numbered 1–8 fields).
 * Used when Form XIV GJ is requested but the linked template is an MP stacked Sheet1 card.
 */
export function stampFormXIVGJOfficialBoxedLayout(worksheet) {
  if (!worksheet) return false;
  try {
    if (typeof worksheet.name === 'string' && worksheet.name !== 'FORM XIV') {
      worksheet.name = 'FORM XIV';
    }
  } catch (_) {
    /* ignore locked sheet name */
  }

  // Clear MP stacked residue (titles in F, truncated B labels, values in C/E).
  for (let r = 1; r <= 40; r += 1) {
    for (let c = 1; c <= 14; c += 1) {
      const cell = worksheet.getCell(r, c);
      cell.value = null;
      cell.border = {};
      cell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
      cell.font = { name: 'Calibri', size: 11, bold: false, italic: false };
    }
  }
  parseExcelJsMergeRanges(worksheet).forEach((m) => {
    unmergeFormXIVRange(worksheet, m.top, m.left, m.bottom, m.right);
  });

  const centerTitle = (row, text, opts = {}) => {
    try {
      worksheet.mergeCells(row, 1, row, 13);
    } catch (_) {
      /* ignore */
    }
    const cell = worksheet.getCell(row, 1);
    cell.value = text;
    cell.font = {
      name: 'Calibri',
      size: opts.size || 12,
      bold: !!opts.bold,
      italic: !!opts.italic,
    };
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  };

  centerTitle(2, 'FORM XIV', { bold: true, size: 14 });
  centerTitle(3, 'EMPLOYMENT CARD', { bold: true, size: 13 });
  centerTitle(4, '(Vide Rule 76 of Contract Labour (R&A) Central & Gujarat Rules)', {
    size: 10,
  });

  const headerBoxes = [
    {
      labelRow: 7,
      labelCol: 1,
      label: 'Name and Address of the Contractor',
      box: { top: 9, left: 1, bottom: 11, right: 6 },
    },
    {
      labelRow: 7,
      labelCol: 7,
      label: 'Name and Address of the Establishment in/under which Contract is carried on:',
      box: { top: 9, left: 7, bottom: 11, right: 13 },
    },
    {
      labelRow: 13,
      labelCol: 1,
      label: 'Nature of work and location of work:',
      box: { top: 15, left: 1, bottom: 17, right: 6 },
    },
    {
      labelRow: 13,
      labelCol: 7,
      label: 'Name and address of Principal Employer:',
      box: { top: 15, left: 7, bottom: 17, right: 13 },
    },
  ];
  headerBoxes.forEach((item) => {
    worksheet.getCell(item.labelRow, item.labelCol).value = item.label;
    worksheet.getCell(item.labelRow, item.labelCol).font = {
      name: 'Calibri',
      size: 11,
      bold: false,
    };
    worksheet.getCell(item.labelRow, item.labelCol).alignment = {
      horizontal: 'left',
      vertical: 'middle',
      wrapText: true,
    };
    try {
      worksheet.mergeCells(item.box.top, item.box.left, item.box.bottom, item.box.right);
    } catch (_) {
      /* ignore */
    }
    applyFormXIVGJBoxBorders(
      worksheet,
      item.box.top,
      item.box.left,
      item.box.bottom,
      item.box.right
    );
  });

  FORM_XIV_GJ_OFFICIAL_WORKMAN_LABELS.forEach((label, i) => {
    const row = 21 + i;
    worksheet.getCell(row, 1).value = label;
    worksheet.getCell(row, 1).font = { name: 'Calibri', size: 11, bold: false };
    worksheet.getCell(row, 1).alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
    try {
      worksheet.mergeCells(row, 2, row, 6);
    } catch (_) {
      /* ignore */
    }
    applyFormXIVGJBoxBorders(worksheet, row, 2, row, 6);
  });

  worksheet.getCell(32, 7).value = 'Signature of the Contractor';
  worksheet.getCell(32, 7).alignment = { horizontal: 'center', vertical: 'middle' };
  for (let c = 1; c <= 13; c += 1) {
    worksheet.getColumn(c).width = c === 1 ? 42 : 12;
  }
  return true;
};

/**
 * Rebuild official MP Form XIV stacked card when Form Master linked a Gujarat boxed template.
 */
export function stampFormXIVMPOfficialStackedLayout(worksheet) {
  if (!worksheet) return false;

  for (let r = 1; r <= 40; r += 1) {
    for (let c = 1; c <= 14; c += 1) {
      const cell = worksheet.getCell(r, c);
      cell.value = null;
      cell.border = {};
      cell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: false };
      cell.font = { name: 'Calibri', size: 12, bold: false, italic: false };
    }
  }
  parseExcelJsMergeRanges(worksheet).forEach((m) => {
    unmergeFormXIVRange(worksheet, m.top, m.left, m.bottom, m.right);
  });

  [
    { row: 2, text: 'FORM XIV', italic: false },
    { row: 3, text: '(See rule 76)', italic: true },
    { row: 4, text: 'Employment Card', italic: false },
  ].forEach((spec) => {
    const cell = worksheet.getCell(spec.row, FORM_XIV_MP_TITLE_COL);
    cell.value = spec.text;
    cell.font = { name: 'Calibri', size: 12, bold: false, italic: !!spec.italic };
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: false };
  });

  FORM_XIV_MP_HEADER_OFFICIAL.forEach((spec, index) => {
    const row = FORM_XIV_MP_OFFICIAL_HEADER_ROWS[index];
    worksheet.getCell(row, FORM_XIV_MP_HEADER_LABEL_COL).value = spec.text;
    styleFormXIVMPSingleLineLabelCell(worksheet.getCell(row, FORM_XIV_MP_HEADER_LABEL_COL));
  });

  FORM_XIV_MP_WORKMAN_FIELD_SPECS.forEach((spec, index) => {
    const row = FORM_XIV_MP_OFFICIAL_WORKMAN_ROWS[index];
    worksheet.getCell(row, FORM_XIV_MP_OFFICIAL_WORKMAN_COL).value = `${index + 1} ${spec.label}`;
    styleFormXIVMPSingleLineLabelCell(worksheet.getCell(row, FORM_XIV_MP_OFFICIAL_WORKMAN_COL));
  });

  applyFormXIVMPOfficialColumnWidths(worksheet, false);
  return true;
}

/** When Form XIV MP is requested but the linked file is a Gujarat boxed card, rebuild MP. */
const ensureFormXIVMPStackedWorksheetLayout = (worksheet, parsedFormHeader = null) => {
  const variant = String(parsedFormHeader?.formXIVVariant || '').toLowerCase();
  if (variant !== 'mp') return false;
  if (detectFormXRajasthanWorksheetLayout(worksheet)) return false;
  if (isFormXIVKarnatakaExport(worksheet, parsedFormHeader)) return false;
  if (!detectFormXIVGJEmploymentCardWorksheet(worksheet)) return false;
  return stampFormXIVMPOfficialStackedLayout(worksheet);
};

const ensureFormXIVGJBoxedWorksheetLayout = (worksheet, parsedFormHeader = null) => {
  const variant = String(parsedFormHeader?.formXIVVariant || '').toLowerCase();
  if (variant !== 'gj') return false;
  if (detectFormXIVGJEmploymentCardWorksheet(worksheet)) return false;
  return stampFormXIVGJOfficialBoxedLayout(worksheet);
};

/** Official MP card: titles in F and even-row header labels, unless ordinals are split in A. */
const detectFormXIVMPOfficialEvenRowCard = (worksheet) => {
  if (!worksheet) return false;
  // Never treat Gujarat boxed Employment Card as MP official even-row layout.
  if (detectFormXIVGJEmploymentCardWorksheet(worksheet)) return false;
  if (detectFormXIVMPColumnCWorkmanLayout(worksheet)) return true;
  let splitOrdinalHits = 0;
  for (let r = 12; r <= 28; r += 1) {
    if (isFormXIVSplitOrdinalLabelRow(worksheet, r)) splitOrdinalHits += 1;
  }
  if (splitOrdinalHits >= 3) return false;
  let titleHit = false;
  for (let r = 1; r <= 5; r += 1) {
    for (let c = 1; c <= 8; c += 1) {
      const raw = formXIVMPExcelCellValueToString(worksheet.getCell(r, c)?.value).trim();
      if (isFormXIVMpTitleText(raw)) titleHit = true;
    }
  }
  let headerHits = 0;
  FORM_XIV_MP_OFFICIAL_HEADER_ROWS.forEach((row) => {
    for (let c = 1; c <= 2; c += 1) {
      const raw = formXIVMPExcelCellValueToString(worksheet.getCell(row, c)?.value).trim();
      if (!isFormXIVMPHeaderLabelCandidate(raw)) continue;
      headerHits += 1;
      break;
    }
  });
  return titleHit && headerHits >= 2;
};

const isFormXIVMPKeepableLabelText = (raw) => {
  const text = normalizeFormXIVMPLabelText(raw);
  if (!text || isFormXIVMpTitleText(text) || isFormXIVPlaceholderCell(text)) return false;
  if (/^\d+[\.\)]?$/.test(text)) {
    const n = Number(String(text).replace(/[.)]/g, ''));
    return n >= 1 && n <= 7;
  }
  if (/^\d+[\.\)]?\s+/.test(text) && isFormXIVStackedWorkmanKeyword(formXIVMPHeaderNorm(text))) {
    return true;
  }
  return isFormXIVMPHeaderLabelCandidate(text);
};

const collectFormXIVMPRowValue = (worksheet, row, keepCols) => {
  let value = '';
  const preferredCols = [FORM_XIV_MP_STACKED_VALUE_COL, FORM_XIV_MP_OFFICIAL_VALUE_COL];
  preferredCols.forEach((col) => {
    if (value || keepCols.has(col)) return;
    const preferred = formXIVMPExcelCellValueToString(worksheet.getCell(row, col)?.value).trim();
    if (preferred && !isFormXIVMPKeepableLabelText(preferred) && !isFormXIVMpTitleText(preferred)) {
      value = preferred;
    }
  });
  for (let c = 1; c <= 8; c += 1) {
    if (keepCols.has(c) || c === FORM_XIV_MP_STACKED_VALUE_COL || c === FORM_XIV_MP_OFFICIAL_VALUE_COL) {
      continue;
    }
    const raw = formXIVMPExcelCellValueToString(worksheet.getCell(row, c)?.value).trim();
    if (!raw || isFormXIVMpTitleText(raw) || isFormXIVPlaceholderCell(raw)) continue;
    if (isFormXIVMPKeepableLabelText(raw)) continue;
    if (!value) value = raw;
    worksheet.getCell(row, c).value = '';
  }
  return value;
};

const clearFormXIVMPOfficialValueBand = (worksheet, row) => {
  if (!worksheet || row < 1) return;
  [1, 4, 5].forEach((col) => {
    if (col === FORM_XIV_MP_TITLE_COL) return;
    worksheet.getCell(row, col).value = '';
  });
};

const stampFormXIVMPOfficialEvenRowLayout = (worksheet) => {
  if (!worksheet) return;
  [
    { row: 2, text: 'FORM XIV', bold: false, italic: false },
    { row: 3, text: '(See rule 76)', bold: false, italic: true },
    { row: 4, text: 'Employment Card', bold: false, italic: false },
  ].forEach((spec) => {
    for (let c = 1; c <= 8; c += 1) {
      const raw = formXIVMPExcelCellValueToString(worksheet.getCell(spec.row, c)?.value).trim();
      if (!raw || !isFormXIVMpTitleText(raw)) continue;
      if (c !== FORM_XIV_MP_TITLE_COL) {
        worksheet.getCell(spec.row, FORM_XIV_MP_TITLE_COL).value = spec.text;
        worksheet.getCell(spec.row, c).value = '';
      } else {
        worksheet.getCell(spec.row, FORM_XIV_MP_TITLE_COL).value = spec.text;
      }
      const titleCell = worksheet.getCell(spec.row, FORM_XIV_MP_TITLE_COL);
      titleCell.font = { name: 'Calibri', size: 12, bold: spec.bold, italic: spec.italic };
      titleCell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: false };
      return;
    }
    const titleCell = worksheet.getCell(spec.row, FORM_XIV_MP_TITLE_COL);
    if (!formXIVMPExcelCellValueToString(titleCell.value).trim()) {
      titleCell.value = spec.text;
      titleCell.font = { name: 'Calibri', size: 12, bold: spec.bold, italic: spec.italic };
      titleCell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: false };
    }
  });

  FORM_XIV_MP_HEADER_OFFICIAL.forEach((spec, index) => {
    const row = FORM_XIV_MP_OFFICIAL_HEADER_ROWS[index];
    unmergeFormXIVRange(worksheet, row, 1, row, 8);
    const value = collectFormXIVMPRowValue(worksheet, row, new Set([FORM_XIV_MP_HEADER_LABEL_COL]));
    clearFormXIVMPOfficialValueBand(worksheet, row);
    worksheet.getCell(row, FORM_XIV_MP_HEADER_LABEL_COL).value = spec.text;
    styleFormXIVMPSingleLineLabelCell(worksheet.getCell(row, FORM_XIV_MP_HEADER_LABEL_COL));
    worksheet.getCell(row, FORM_XIV_MP_OFFICIAL_VALUE_COL).value = value || '';
    if (value) {
      styleFormXIVMPBodyValueCell(
        worksheet.getCell(row, FORM_XIV_MP_OFFICIAL_VALUE_COL),
        String(value).length > 36
      );
    }
  });

  FORM_XIV_MP_WORKMAN_FIELD_SPECS.forEach((spec, index) => {
    const row = FORM_XIV_MP_OFFICIAL_WORKMAN_ROWS[index];
    unmergeFormXIVRange(worksheet, row, 1, row, 8);
    const value = collectFormXIVMPRowValue(
      worksheet,
      row,
      new Set([FORM_XIV_MP_OFFICIAL_WORKMAN_COL])
    );
    clearFormXIVMPOfficialValueBand(worksheet, row);
    worksheet.getCell(row, FORM_XIV_MP_OFFICIAL_WORKMAN_COL).value = `${index + 1} ${spec.label}`;
    styleFormXIVMPSingleLineLabelCell(worksheet.getCell(row, FORM_XIV_MP_OFFICIAL_WORKMAN_COL));
    worksheet.getCell(row, FORM_XIV_MP_OFFICIAL_VALUE_COL).value = value || '';
    if (value) {
      styleFormXIVMPBodyValueCell(
        worksheet.getCell(row, FORM_XIV_MP_OFFICIAL_VALUE_COL),
        false
      );
    }
  });
};

const applyFormXIVMPOfficialColumnWidths = (worksheet, workmanInC = false) => {
  if (!worksheet) return;
  const minWidths = {
    1: 5,
    2: 92,
    3: workmanInC ? 72 : 36,
    4: 16,
    5: 18,
    6: 18,
    7: 14,
    8: 14,
  };
  Object.entries(minWidths).forEach(([col, minW]) => {
    const column = worksheet.getColumn(Number(col));
    column.hidden = false;
    column.width = minW;
  });
};

const writeFormXIVMPOfficialEvenRowFieldValues = (worksheet, headerFormData = {}, workmanRow = null, headers = null) => {
  if (!worksheet) return;
  FORM_XIV_MP_HEADER_SPECS.forEach((spec, index) => {
    const row = FORM_XIV_MP_OFFICIAL_HEADER_ROWS[index];
    const text = String(headerFormData?.[spec.key] ?? '').trim();
    if (!row || !text) return;
    worksheet.getCell(row, FORM_XIV_MP_STACKED_VALUE_COL).value = '';
    const cell = worksheet.getCell(row, FORM_XIV_MP_OFFICIAL_VALUE_COL);
    cell.value = text;
    styleFormXIVMPBodyValueCell(cell, String(text).length > 36);
  });
  if (!workmanRow || typeof workmanRow !== 'object') return;
  const hdrs = resolveFormXIVMPTableHeaders(headers);
  FORM_XIV_MP_WORKMAN_FIELD_SPECS.forEach((spec, index) => {
    const row = FORM_XIV_MP_OFFICIAL_WORKMAN_ROWS[index];
    const value = getFormXIVKAWorkmanFieldValue(workmanRow, hdrs, index);
    if (!row) return;
    worksheet.getCell(row, FORM_XIV_MP_STACKED_VALUE_COL).value = '';
    const cell = worksheet.getCell(row, FORM_XIV_MP_OFFICIAL_VALUE_COL);
    if (value) {
      setFormXIVMPWorkmanCellValue(worksheet, row, FORM_XIV_MP_OFFICIAL_VALUE_COL, value, {
        variant: 'mp',
        inBox: false,
      });
    }
    styleFormXIVMPBodyValueCell(cell, false);
  });
};

const ensureFormXIVMadhyaPradeshColumnAlignment = (worksheet) => {
  if (!worksheet) return;
  if (worksheetLooksLikeFormAEmployeeRegister(worksheet)) return;
  const officialCard = detectFormXIVMPOfficialEvenRowCard(worksheet);
  if (officialCard) {
    stampFormXIVMPOfficialEvenRowLayout(worksheet);
  } else {
    applyFormXIVMPOfficialTemplateLabels(worksheet);
  }

  const workmanInC =
    !officialCard &&
    findFormXIVMPWorkmanLabelRows(worksheet).some((pos) => pos.col === 3);
  applyFormXIVMPOfficialColumnWidths(worksheet, workmanInC);

  for (let r = 1; r <= 45; r += 1) {
    let skipTitle = false;
    for (let c = 1; c <= 8; c += 1) {
      const titleText = formXIVMPExcelCellValueToString(worksheet.getCell(r, c)?.value).trim();
      if (isFormXIVMpTitleText(titleText)) {
        skipTitle = true;
        break;
      }
    }
    if (skipTitle) {
      for (let c = 1; c <= 8; c += 1) {
        const titleCell = worksheet.getCell(r, c);
        const titleText = formXIVMPExcelCellValueToString(titleCell.value).trim();
        if (!isFormXIVMpTitleText(titleText)) continue;
        titleCell.font = {
          name: 'Calibri',
          size: 12,
          bold: false,
          italic: /see\s+rule\s*76|vide\s+rule\s*76/i.test(titleText),
        };
      }
      continue;
    }

    const labelCol = findFormXIVMPStackedLabelCol(worksheet, r);
    const excelRow = worksheet.getRow(r);
    excelRow.font = { ...(excelRow.font || {}), bold: false };
    excelRow.alignment = {
      ...(excelRow.alignment || {}),
      wrapText: false,
      shrinkToFit: false,
      vertical: 'middle',
    };
    if (labelCol >= 1 && Number(excelRow.height) > 22) excelRow.height = 18;

    for (let c = 1; c <= 8; c += 1) {
      const cell = worksheet.getCell(r, c);
      const text = formXIVMPExcelCellValueToString(cell.value).trim();
      if (!text) continue;
      const isLabelCol =
        c <= 4 && c !== FORM_XIV_MP_STACKED_VALUE_COL && c !== FORM_XIV_MP_OFFICIAL_VALUE_COL;
      if (isLabelCol) {
        styleFormXIVMPSingleLineLabelCell(cell);
        continue;
      }
      styleFormXIVMPBodyValueCell(
        cell,
        String(text).length > 36 || String(text).includes('\n')
      );
    }
  }
  applyFormXIVMPOfficialColumnWidths(worksheet, workmanInC);
};

/** Keep stacked Employment Card label/value columns readable after ExcelJS rewrite. */
const ensureFormXIVStackedColumnAlignment = (worksheet, parsedFormHeader = null) => {
  if (!worksheet) return;
  if (isFormXIVKarnatakaExport(worksheet, parsedFormHeader)) {
    ensureFormXIVKarnatakaColumnAlignment(worksheet);
    return;
  }
  if (shouldApplyFormXIVMPExportAlignment(worksheet, parsedFormHeader)) {
    ensureFormXIVMadhyaPradeshColumnAlignment(worksheet);
  }
};

/** Re-apply MP Employment Card layout after download finalize (system-note pass). */
export function finalizeFormXIVMadhyaPradeshWorksheet(worksheet, parsedFormHeader = null) {
  if (!worksheet) return false;
  if (worksheetLooksLikeFormAEmployeeRegister(worksheet)) return false;
  const variant = String(parsedFormHeader?.formXIVVariant || '').toLowerCase();
  if (variant === 'gj' || variant === 'ka' || variant === 'rj') return false;
  if (detectFormXIVGJEmploymentCardWorksheet(worksheet)) return false;
  if (detectFormXIVMPOfficialEvenRowCard(worksheet)) {
    ensureFormXIVMadhyaPradeshColumnAlignment(worksheet);
    applyFormXIVMPOfficialColumnWidths(worksheet, false);
    return true;
  }
  if (isFormXIVKarnatakaExport(worksheet, parsedFormHeader)) return false;
  if (!shouldApplyFormXIVMPExportAlignment(worksheet, parsedFormHeader)) return false;
  ensureFormXIVMadhyaPradeshColumnAlignment(worksheet);
  return true;
}

const setFormXIVMPWorkmanCellValue = (worksheet, row, col, value, options = {}) => {
  const { variant = 'mp', inBox = false } = options;
  const writePos =
    variant === 'gj' ? resolveGJMergeTopLeft(worksheet, row, col) : { row, col };
  const text = String(value ?? '').trim();
  const cell = worksheet.getCell(writePos.row, writePos.col);
  if (!text) {
    cell.value = '';
    return;
  }
  const normalized = text.replace(/,/g, '');
  const asNum = Number(normalized);
  cell.value =
    Number.isFinite(asNum) && /^-?\d+(\.\d+)?$/.test(normalized) ? asNum : text;
  const isGjBox = variant === 'gj' && inBox;
  const isGjWorkman = variant === 'gj' && !inBox;
  const isRjBox = inBox && variant === 'mp';
  cell.alignment = {
    ...(cell.alignment || {}),
    horizontal: isRjBox ? 'center' : 'left',
    vertical: isGjWorkman || isRjBox ? 'middle' : 'top',
    wrapText: isGjBox || isRjBox || String(cell.value).length > 36 || String(cell.value).includes('\n'),
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
    return resolveFormXIVStackedValueCol(worksheet);
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
  const valueFrom = resolveFormXIVStackedValueCol(worksheet);
  for (let r = fromRow; r <= toRow; r += 1) {
    const workmanOrdinal = findFormXIVMPWorkmanOrdinalOnRow(worksheet, r);
    for (let c = FORM_XIV_GJ_DEFAULT_VALUE_COL; c <= FORM_XIV_MP_STACKED_VALUE_COL_TO; c += 1) {
      if (c >= valueFrom) {
        worksheet.getCell(r, c).value = '';
        continue;
      }
      if (c <= 4 && workmanOrdinal >= 1) continue;
      const cellStr = formXIVMPExcelCellValueToString(worksheet.getCell(r, c)?.value).trim();
      if (!cellStr) continue;
      if (c <= 4 && isFormXIVMPHeaderLabelCandidate(cellStr)) continue;
      if (workmanSpecs.some((spec) => looksLikeFormXIVMPWorkmanLabelCell(cellStr, spec))) continue;
      if (FORM_XIV_MP_HEADER_SPECS.some((spec) => spec.match.test(cellStr))) continue;
      if (isFormXIVLabelContinuationCell(cellStr, workmanSpecs)) continue;
      if (c <= 4 && /^\d+[\.\)]?\s+/.test(cellStr)) continue;
      worksheet.getCell(r, c).value = '';
    }
  }
};

const isGJHeaderLabelContinuation = (raw) => {
  const n = formXIVMPHeaderNorm(raw);
  if (!n || n.length > 120) return false;
  if (isFormXIVHeaderLabelBlob(raw)) return true;
  if (/contract\s+is\s+carried/.test(n)) return true;
  if (/in\s*\/\s*under\s+which/.test(n)) return true;
  if (/^of\s+(?:the\s+)?(?:contractor|establishment|principal)/.test(n)) return true;
  if (/^the\s+establishment\b/.test(n)) return true;
  return FORM_XIV_MP_HEADER_SPECS.some(
    (spec) => spec.match.test(raw) || spec.match.test(n)
  );
};

const matchesFormXIVHeaderSpecLabel = (spec, raw) => {
  const text = String(raw || '').trim();
  if (!text || !spec?.match) return false;
  const n = formXIVMPHeaderNorm(text);
  return spec.match.test(text) || spec.match.test(n);
};

const resolveFormXIVGJHeaderHalfBounds = (labelCol) => {
  const isRightHalf = labelCol >= 6;
  return {
    isRightHalf,
    halfStart: isRightHalf ? 7 : 1,
    halfEnd: isRightHalf ? 12 : 6,
  };
};

const cellHasVisibleBorder = (border, side) => !!border?.[side]?.style;

/** GJ template: heading row, spacer row, bordered value box from labelRow + 2 downward. */
const resolveFormXIVGJHeaderBoxArea = (worksheet, labelRow, labelCol) => {
  const { halfStart, halfEnd } = resolveFormXIVGJHeaderHalfBounds(labelCol);
  const minBoxRow = labelRow + 2;
  const ranges = parseExcelJsMergeRanges(worksheet);

  let bestMerge = null;
  ranges.forEach((m) => {
    if (m.top < minBoxRow) return;
    if (m.top > labelRow + 16) return;
    if (m.bottom - m.top < 1) return;
    const overlapLeft = Math.max(m.left, halfStart);
    const overlapRight = Math.min(m.right, halfEnd);
    if (overlapRight - overlapLeft < 1) return;
    if (!bestMerge || m.top < bestMerge.top || (m.top === bestMerge.top && m.left < bestMerge.left)) {
      bestMerge = m;
    }
  });

  if (bestMerge) {
    return {
      valueRow: bestMerge.top,
      valueCol: Math.max(bestMerge.left, halfStart),
      top: bestMerge.top,
      bottom: bestMerge.bottom,
      left: Math.max(bestMerge.left, halfStart),
      right: Math.min(bestMerge.right, halfEnd),
    };
  }

  for (let r = minBoxRow; r <= labelRow + 14; r += 1) {
    for (let c = halfStart; c <= halfEnd; c += 1) {
      const b = worksheet.getCell(r, c)?.border || {};
      if (!cellHasVisibleBorder(b, 'top') && !cellHasVisibleBorder(b, 'left')) continue;
      let bottom = r;
      let right = c;
      for (let br = r; br <= r + 12; br += 1) {
        const bb = worksheet.getCell(br, c)?.border || {};
        if (cellHasVisibleBorder(bb, 'left') || cellHasVisibleBorder(bb, 'right')) bottom = br;
        else if (br > r + 1) break;
      }
      for (let bc = c; bc <= halfEnd; bc += 1) {
        const bb = worksheet.getCell(r, bc)?.border || {};
        if (cellHasVisibleBorder(bb, 'top') || cellHasVisibleBorder(bb, 'bottom')) right = bc;
        else if (bc > c + 1) break;
      }
      return {
        valueRow: r,
        valueCol: c,
        top: r,
        bottom,
        left: c,
        right,
      };
    }
  }

  return {
    valueRow: minBoxRow,
    valueCol: halfStart,
    top: minBoxRow,
    bottom: minBoxRow + 4,
    left: halfStart,
    right: halfEnd,
  };
};

const findFormXIVGJHeaderBoxValuePosition = (worksheet, labelRow, labelCol) => {
  const area = resolveFormXIVGJHeaderBoxArea(worksheet, labelRow, labelCol);
  return { row: area.valueRow, col: area.valueCol };
};

const findFormXIVGJHeaderLabelPosition = (worksheet, spec, parsedFields = []) => {
  if (!worksheet || !spec) return null;
  const parsed = Array.isArray(parsedFields)
    ? parsedFields.find((f) => f.key === spec.key)
    : null;
  if (parsed?.labelRow != null) {
    return {
      row: parsed.labelRow + 1,
      col: (parsed.labelCol ?? 0) + 1,
    };
  }
  for (let r = 1; r <= 30; r += 1) {
    for (let c = 1; c <= 12; c += 1) {
      const raw = formXIVMPExcelCellValueToString(worksheet.getCell(r, c)?.value).trim();
      if (!raw || !matchesFormXIVHeaderSpecLabel(spec, raw)) continue;
      return { row: r, col: c };
    }
  }
  const canonicalByKey = {
    form_xiv_mp_contractor: { row: 7, col: 1 },
    form_xiv_mp_establishment: { row: 7, col: 7 },
    form_xiv_mp_nature_location: { row: 13, col: 1 },
    form_xiv_mp_principal_employer: { row: 13, col: 7 },
  };
  return canonicalByKey[spec.key] || null;
};

const getFormXIVGJHeaderLabelLines = (worksheet, spec, snapshots, parsedFields) => {
  const snap = snapshots.find((s) => s.spec?.key === spec.key);
  if (snap?.lines?.length) return snap.lines;
  const pos = findFormXIVGJHeaderLabelPosition(worksheet, spec, parsedFields);
  if (!pos) return [{ row: 5, col: 1, text: spec.label }];
  const lines = [{ row: pos.row, col: pos.col, text: spec.label }];
  const nextRaw = formXIVMPExcelCellValueToString(worksheet.getCell(pos.row + 1, pos.col)?.value).trim();
  if (nextRaw && isGJHeaderLabelContinuation(nextRaw) && !matchesFormXIVHeaderSpecLabel(spec, nextRaw)) {
    lines.push({ row: pos.row + 1, col: pos.col, text: nextRaw });
  }
  return lines;
};

const clearFormXIVGJHeaderLabelAreaData = (worksheet, labelRow, labelCol, area, spec = null) => {
  if (!worksheet || labelRow < 1) return;
  const { halfStart, halfEnd } = resolveFormXIVGJHeaderHalfBounds(labelCol);
  const boxTop = area?.top ?? labelRow + 2;
  for (let r = labelRow; r < boxTop; r += 1) {
    for (let c = halfStart; c <= halfEnd; c += 1) {
      const raw = formXIVMPExcelCellValueToString(worksheet.getCell(r, c)?.value).trim();
      if (!raw) continue;
      if (spec && (matchesFormXIVHeaderSpecLabel(spec, raw) || isGJHeaderLabelContinuation(raw))) {
        continue;
      }
      if (!spec && (isFormXIVHeaderLabelBlob(raw) || isGJHeaderLabelContinuation(raw))) continue;
      worksheet.getCell(r, c).value = '';
    }
  }
};

const snapshotFormXIVGJHeaderLabels = (worksheet) => {
  if (!worksheet) return [];
  const snapshots = [];
  FORM_XIV_MP_HEADER_SPECS.forEach((spec) => {
    for (let r = 1; r <= 25; r += 1) {
      for (let c = 1; c <= 12; c += 1) {
        const raw = formXIVMPExcelCellValueToString(worksheet.getCell(r, c)?.value).trim();
        if (!raw || !matchesFormXIVHeaderSpecLabel(spec, raw)) continue;
        const lines = [{ row: r, col: c, text: raw }];
        const nextRaw = formXIVMPExcelCellValueToString(worksheet.getCell(r + 1, c)?.value).trim();
        if (nextRaw && isGJHeaderLabelContinuation(nextRaw) && !matchesFormXIVHeaderSpecLabel(spec, nextRaw)) {
          lines.push({ row: r + 1, col: c, text: nextRaw });
        }
        snapshots.push({ spec, lines });
        return;
      }
    }
  });
  return snapshots;
};

const restoreFormXIVGJHeaderLabels = (worksheet, snapshots = [], parsedFields = []) => {
  if (!worksheet) return;
  FORM_XIV_MP_HEADER_SPECS.forEach((spec) => {
    const lines = getFormXIVGJHeaderLabelLines(worksheet, spec, snapshots, parsedFields);
    lines.forEach(({ row, col, text }) => {
      if (row < 1 || col < 1 || !text) return;
      const cell = worksheet.getCell(row, col);
      const current = formXIVMPExcelCellValueToString(cell.value).trim();
      if (current && (matchesFormXIVHeaderSpecLabel(spec, current) || isGJHeaderLabelContinuation(current))) {
        return;
      }
      cell.value = text;
      const existingFont = cell.font || {};
      if (!existingFont.bold) {
        cell.font = { ...existingFont, bold: true };
      }
    });
  });
};

const isGJHeaderLabelRow = (worksheet, row, labelCol) => {
  if (!worksheet || row < 1 || labelCol < 1) return false;
  const raw = formXIVMPExcelCellValueToString(worksheet.getCell(row, labelCol)?.value).trim();
  if (!raw || isFormXIVPlaceholderCell(raw)) return false;
  return isGJHeaderLabelContinuation(raw);
};

const clearFormXIVGJHeaderStrayAboveBox = (worksheet, labelRow, labelCol, boxRow) => {
  if (!worksheet || labelRow < 1 || boxRow <= labelRow) return;
  const isRightHalf = labelCol >= 6;
  const colStart = isRightHalf ? 6 : 1;
  const colEnd = isRightHalf ? 12 : 6;
  for (let r = labelRow; r < boxRow; r += 1) {
    for (let c = colStart; c <= colEnd; c += 1) {
      if (r === labelRow && c === labelCol) continue;
      const raw = formXIVMPExcelCellValueToString(worksheet.getCell(r, c)?.value).trim();
      if (!raw) continue;
      if (c === labelCol && (isFormXIVHeaderLabelBlob(raw) || isGJHeaderLabelContinuation(raw))) {
        continue;
      }
      if (r === labelRow && isFormXIVHeaderLabelBlob(raw)) continue;
      worksheet.getCell(r, c).value = '';
    }
  }
};

const clearFormXIVGJHeaderRowSpill = (worksheet, labelRow, labelCol, boxRow) => {
  if (!worksheet || labelRow < 1 || labelCol < 1) return;
  const spillEnd = 14;
  for (let c = labelCol + 1; c <= spillEnd; c += 1) {
    if (boxRow === labelRow && c === labelCol) continue;
    const cellStr = formXIVMPExcelCellValueToString(worksheet.getCell(labelRow, c)?.value).trim();
    if (!cellStr) continue;
    if (isFormXIVHeaderLabelBlob(cellStr)) continue;
    worksheet.getCell(labelRow, c).value = '';
  }
};

/** Remove wrongly beside-written header values (e.g. address spill into column M). */
const clearFormXIVGJHeaderSpillZones = (worksheet) => {
  if (!worksheet) return;
  for (let r = 1; r <= 22; r += 1) {
    for (let c = 9; c <= 14; c += 1) {
      const raw = formXIVMPExcelCellValueToString(worksheet.getCell(r, c)?.value).trim();
      if (!raw) continue;
      if (isFormXIVPlaceholderCell(raw)) continue;
      if (isFormXIVHeaderLabelBlob(raw)) continue;
      if (isTableLabelBlob(raw)) continue;
      if (/^form\s*xiv|employment\s+card|see\s+rule\s+76|vide\s+rule\s+76/i.test(raw)) continue;
      worksheet.getCell(r, c).value = '';
    }
  }
  FORM_XIV_MP_HEADER_SPECS.forEach((spec) => {
    for (let r = 1; r <= 22; r += 1) {
      for (let c = 1; c <= 12; c += 1) {
        const raw = formXIVMPExcelCellValueToString(worksheet.getCell(r, c)?.value).trim();
        if (!raw || (!spec.match.test(raw) && !spec.match.test(formXIVMPHeaderNorm(raw)))) continue;
        for (let spillCol = c + 1; spillCol <= 14; spillCol += 1) {
          const spill = formXIVMPExcelCellValueToString(worksheet.getCell(r, spillCol)?.value).trim();
          if (!spill) continue;
          if (isFormXIVHeaderLabelBlob(spill)) break;
          worksheet.getCell(r, spillCol).value = '';
        }
        return;
      }
    }
  });
};

const clearFormXIVGJHeaderBoxBand = (worksheet, labelRow, labelCol, spec = null) => {
  if (!worksheet || labelRow < 1 || labelCol < 1) return;
  const area = resolveFormXIVGJHeaderBoxArea(worksheet, labelRow, labelCol);
  let startRow = area.top;
  const endRow = area.bottom;
  const startCol = area.left;
  const endCol = area.right;
  if (startRow <= labelRow) startRow = labelRow + 2;
  while (startRow <= endRow && isGJHeaderLabelRow(worksheet, startRow, labelCol)) {
    startRow += 1;
  }
  if (startRow > endRow) return;
  for (let r = startRow; r <= endRow; r += 1) {
    if (isGJHeaderLabelRow(worksheet, r, labelCol)) continue;
    for (let c = startCol; c <= endCol; c += 1) {
      const raw = formXIVMPExcelCellValueToString(worksheet.getCell(r, c)?.value).trim();
      if (
        raw &&
        c === labelCol &&
        (isFormXIVHeaderLabelBlob(raw) ||
          isGJHeaderLabelContinuation(raw) ||
          (spec && matchesFormXIVHeaderSpecLabel(spec, raw)))
      ) {
        continue;
      }
      worksheet.getCell(r, c).value = '';
    }
  }
};

const writeFormXIVGJBoxedHeaderValue = (worksheet, labelRow, labelCol, value, spec = null) => {
  const area = resolveFormXIVGJHeaderBoxArea(worksheet, labelRow, labelCol);
  clearFormXIVGJHeaderLabelAreaData(worksheet, labelRow, labelCol, area, spec);
  clearFormXIVGJHeaderBoxBand(worksheet, labelRow, labelCol, spec);
  clearFormXIVGJHeaderStrayAboveBox(worksheet, labelRow, labelCol, area.top);
  clearFormXIVGJHeaderRowSpill(worksheet, labelRow, labelCol, area.top);
  setFormXIVMPWorkmanCellValue(worksheet, area.valueRow, area.valueCol, value, {
    variant: 'gj',
    inBox: true,
  });
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
  const gjBoxedLayout = variant === 'gj';
  if (gjBoxedLayout) {
    ensureFormXIVGJBoxedWorksheetLayout(worksheet, parsedFormHeader);
  } else if (variant === 'mp') {
    ensureFormXIVMPStackedWorksheetLayout(worksheet, parsedFormHeader);
  }
  const rjBoxedLayout = variant === 'rj' || detectFormXRajasthanWorksheetLayout(worksheet);
  // Stamp official MP card before filling so truncated B labels cannot pull values into C.
  if (
    variant === 'mp' &&
    !gjBoxedLayout &&
    !rjBoxedLayout &&
    !isFormXIVKarnatakaExport(worksheet, parsedFormHeader) &&
    detectFormXIVMPOfficialEvenRowCard(worksheet)
  ) {
    stampFormXIVMPOfficialEvenRowLayout(worksheet);
    applyFormXIVMPOfficialColumnWidths(worksheet, false);
  }
  // MP/KA Employment Card is always stacked label|value (KA values in F, official MP in C).
  const stackedLayout =
    !gjBoxedLayout &&
    !rjBoxedLayout &&
    (isFormXIVStackedVariant(variant) || detectFormXIVStackedWorkmanLayout(worksheet));
  const kaLayout = isFormXIVKarnatakaExport(worksheet, parsedFormHeader);
  const mpOfficialCard =
    variant === 'mp' &&
    !gjBoxedLayout &&
    !rjBoxedLayout &&
    !kaLayout &&
    detectFormXIVMPOfficialEvenRowCard(worksheet);
  const stackedValueCol = mpOfficialCard
    ? FORM_XIV_MP_OFFICIAL_VALUE_COL
    : resolveFormXIVStackedValueCol(worksheet, parsedFormHeader);
  const defaultValueCol = stackedLayout
    ? stackedValueCol
    : rjBoxedLayout
      ? FORM_XIV_RJ_HEADER_VALUE_COL
      : resolveFormXIVMPWorkmanValueColumn(worksheet, 45, workmanSpecs);
  const parsedFields = Array.isArray(parsedFormHeader?.fields) ? parsedFormHeader.fields : [];
  const gjHeaderLabelSnapshots = gjBoxedLayout ? snapshotFormXIVGJHeaderLabels(worksheet) : [];

  if (gjBoxedLayout) {
    FORM_XIV_MP_HEADER_SPECS.forEach((spec) => {
      const labelPos = findFormXIVGJHeaderLabelPosition(worksheet, spec, parsedFields);
      if (!labelPos) return;
      const area = resolveFormXIVGJHeaderBoxArea(worksheet, labelPos.row, labelPos.col);
      clearFormXIVGJHeaderLabelAreaData(worksheet, labelPos.row, labelPos.col, area, spec);
    });
    clearFormXIVGJHeaderSpillZones(worksheet);
    FORM_XIV_MP_HEADER_SPECS.forEach((spec) => {
      const labelPos = findFormXIVGJHeaderLabelPosition(worksheet, spec, parsedFields);
      if (!labelPos) return;
      clearFormXIVGJHeaderBoxBand(worksheet, labelPos.row, labelPos.col, spec);
    });
  }

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
  const writeBesideLabel = (excelRow, labelCol, value, spec = null) => {
    if (excelRow < 1) return;
    if (gjBoxedLayout) {
      writeFormXIVGJBoxedHeaderValue(worksheet, excelRow, labelCol, value, spec);
      return;
    }
    if (rjBoxedLayout) {
      writeFormXIVRJBoxedHeaderValue(worksheet, excelRow, labelCol, value);
      return;
    }
    // Karnataka: keep label + dotted fill-in + value on the same line.
    if (kaLayout) {
      writeFormXIVKAInlineOnLabelRow(worksheet, excelRow, labelCol, value);
      return;
    }
    // Stacked MP cards: labels occupy A–D; values stay in column E.
    if (stackedLayout) {
      writeAt(excelRow, stackedValueCol, value);
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
    // Never write template dotted labels back onto the sheet as "values".
    if (isFormXIVMPHeaderValueEmptyOrPlaceholder(val)) return;

    const labelPos = gjBoxedLayout
      ? findFormXIVGJHeaderLabelPosition(worksheet, spec, parsedFields)
      : null;
    const parsedField = parsedFields.find((f) => f.key === spec.key);
    if (labelPos || parsedField?.labelRow != null) {
      const labelExcelRow = labelPos?.row ?? parsedField.labelRow + 1;
      const labelExcelCol = labelPos?.col ?? (parsedField.labelCol ?? 0) + 1;
      if (gjBoxedLayout) {
        clearFormXIVGJHeaderBoxBand(worksheet, labelExcelRow, labelExcelCol, spec);
        writeFormXIVGJBoxedHeaderValue(worksheet, labelExcelRow, labelExcelCol, val, spec);
        return;
      }
      if (rjBoxedLayout) {
        writeFormXIVRJBoxedHeaderValue(worksheet, labelExcelRow, labelExcelCol, val);
        return;
      }
      if (kaLayout) {
        writeFormXIVKAInlineOnLabelRow(worksheet, labelExcelRow, labelExcelCol, val);
        return;
      }
      if (stackedLayout) {
        const targetRow = (parsedField?.valueRow ?? parsedField?.labelRow ?? labelExcelRow - 1) + 1;
        writeAt(targetRow, stackedValueCol, val);
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
      writeAt(targetRow, targetCol, val);
      return;
    }

    for (let r = 1; r <= 40; r += 1) {
      for (let c = 1; c <= 14; c += 1) {
        const raw = formXIVMPExcelCellValueToString(worksheet.getCell(r, c)?.value).trim();
        if (!raw || (!spec.match.test(raw) && !spec.match.test(formXIVMPHeaderNorm(raw)))) continue;
        writeBesideLabel(r, c, val, spec);
        return;
      }
    }
    if (kaLayout) {
      const specIndex = FORM_XIV_MP_HEADER_SPECS.findIndex((item) => item.key === spec.key);
      const headerRows = findFormXIVKAHeaderLabelRows(worksheet);
      const pos = specIndex >= 0 ? headerRows[specIndex] : null;
      if (pos) {
        writeFormXIVKAInlineOnLabelRow(
          worksheet,
          pos.row,
          pos.col,
          val,
          FORM_XIV_KA_HEADER_OFFICIAL[specIndex]?.text || ''
        );
      }
    }
  });
  if (gjBoxedLayout) {
    restoreFormXIVGJHeaderLabels(worksheet, gjHeaderLabelSnapshots, parsedFields);
  }
  if (stackedLayout || kaLayout) {
    ensureFormXIVStackedColumnAlignment(worksheet, {
      ...(parsedFormHeader || {}),
      formXIVVariant: kaLayout ? 'ka' : parsedFormHeader?.formXIVVariant || 'mp',
    });
  }
  if (
    variant === 'mp' &&
    !gjBoxedLayout &&
    !rjBoxedLayout &&
    !kaLayout &&
    detectFormXIVMPOfficialEvenRowCard(worksheet)
  ) {
    writeFormXIVMPOfficialEvenRowFieldValues(worksheet, headerFormData);
    applyFormXIVMPOfficialColumnWidths(worksheet, false);
  }
}

export function resolveFormXIVMPWorkmanFieldPositions(worksheet, headers, parsedFormHeader = null) {
  const hdrs = resolveFormXIVMPTableHeaders(headers, {
    formHeader: parsedFormHeader,
    fileName: parsedFormHeader?.title || '',
  });
  const explicitVariant = String(parsedFormHeader?.formXIVVariant || '').toLowerCase();
  const workmanSpecs =
    explicitVariant === 'gj'
      ? FORM_XIV_GJ_WORKMAN_FIELD_SPECS
      : resolveFormXIVWorkmanFieldSpecsFromHeaders(hdrs);
  const variant =
    explicitVariant === 'gj' || explicitVariant === 'rj' || explicitVariant === 'ka'
      ? explicitVariant
      : resolveFormXIVMPWorkmanVariantFromSpecs(workmanSpecs);
  // MP/KA cards always use stacked value columns (E for MP, F for Karnataka).
  // Gujarat boxed card must never fall into stacked MP layout.
  const stackedLayout =
    variant !== 'gj' &&
    variant !== 'rj' &&
    isFormXIVStackedVariant(variant) &&
    (detectFormXIVStackedWorkmanLayout(worksheet) || !detectFormXRajasthanWorksheetLayout(worksheet));
  const kaInline = variant !== 'gj' && isFormXIVKarnatakaExport(worksheet, parsedFormHeader);
  const stackedValueCol = resolveFormXIVStackedValueCol(worksheet, parsedFormHeader);
  const valueCol = stackedLayout
    ? stackedValueCol
    : resolveFormXIVMPWorkmanValueColumn(worksheet, 45, workmanSpecs);
  const positions = [];
  const filled = new Set();

  const resolveTargetCol = (r, c) => {
    if (kaInline) return c;
    if (stackedLayout) return stackedValueCol;
    if (variant === 'gj') {
      return findGJWorkmanValueMergeOnRow(worksheet, r, c).col;
    }
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
      if (isFormXIVMergeSlaveCell(worksheet, r, c)) continue;
      const cellStr = readFormXIVMPLabelWithContinuation(worksheet, r, c);
      if (!cellStr) continue;
      const ordinalMatch = cellStr.match(/^(\d+)[\.\)]?\s*/);
      const ordinal = ordinalMatch ? Number(ordinalMatch[1]) : null;
      for (let si = 0; si < workmanSpecs.length; si += 1) {
        const spec = workmanSpecs[si];
        if (filled.has(si)) continue;
        if (ordinal && ordinal !== si + 1) continue;
        if (!looksLikeFormXIVMPWorkmanLabelCell(cellStr, spec, ordinal || si + 1)) continue;
        const headerKey = hdrs.find((h) => spec.rowTest(h)) || hdrs[si];
        let targetRow = r;
        let targetCol = resolveTargetCol(r, c);
        // GJ wage-rate label often wraps to a continuation line ("case of piece - work)").
        // Prefer the value box on the "5." label row; only use the next row when that
        // line is a continuation AND the ordinal row has no box to the right.
        if (variant === 'gj' && ordinal === 5 && isFormXIVMPWageRateHeader(cellStr)) {
          const thisPos = findGJWorkmanValueMergeOnRow(worksheet, r, c);
          const nextRaw = formXIVMPExcelCellValueToString(worksheet.getCell(r + 1, c)?.value).trim();
          const nextIsContinuation =
            !!nextRaw &&
            !/^\d+[\.\)]\s*/.test(nextRaw) &&
            !isFormXIVMPWagePeriodHeader(nextRaw) &&
            !isFormXIVMPTenureHeader(nextRaw) &&
            (/(?:^case\s+of|piece|particular|work\))/i.test(nextRaw) ||
              (cellStr.includes('(') && !String(cellStr).includes(')')));
          const thisHasBox = thisPos.col > c;
          if (thisHasBox) {
            targetRow = r;
            targetCol = thisPos.col;
          } else if (nextIsContinuation) {
            const nextPos = findGJWorkmanValueMergeOnRow(worksheet, r + 1, c);
            if (nextPos.col > c) {
              targetRow = r + 1;
              targetCol = nextPos.col;
            }
          }
        }
        positions.push({
          headerKey,
          row: targetRow,
          col: targetCol,
          labelCol: c,
          officialLabel: FORM_XIV_KA_WORKMAN_OFFICIAL[(ordinal || si + 1) - 1] || '',
        });
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
        if (isFormXIVMergeSlaveCell(worksheet, r, c)) continue;
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
  return { valueCol, positions, workmanSpecs, stackedLayout, variant, kaInline };
}

export function clearFormXIVMPWorkmanFieldPositions(worksheet, positions, options = {}) {
  if (!worksheet || !Array.isArray(positions)) return;
  const variant = options.variant || 'mp';
  const kaInline = options.kaInline || detectFormXIVKarnatakaWorksheetLayout(worksheet);
  positions.forEach(({ row, col, labelCol }) => {
    if (row < 1 || col < 1) return;
    if (variant === 'gj') {
      clearFormXIVGJWorkmanValueBand(worksheet, row, labelCol ?? 1);
      return;
    }
    if (kaInline) {
      // Label + dotted value share one merged line; clearing slaves would wipe the label.
      return;
    }
    worksheet.getCell(row, col).value = '';
  });
}

export function writeFormXIVMPWorkmanFieldsViaPositions(worksheet, row, positions, options = {}) {
  if (!worksheet || !row || typeof row !== 'object' || !Array.isArray(positions)) return;
  const variant = options.variant || 'mp';
  const kaInline = options.kaInline || detectFormXIVKarnatakaWorksheetLayout(worksheet);
  positions.forEach(({ headerKey, row: targetRow, col, labelCol, officialLabel }) => {
    const value = getFormXIVMPRowValueForHeader(row, headerKey);
    let writeRow = targetRow;
    let writeCol = col;
    if (variant === 'gj') {
      if (labelCol != null) {
        clearFormXIVGJWorkmanValueBand(worksheet, targetRow, labelCol);
      }
      const merged = resolveGJMergeTopLeft(worksheet, targetRow, col);
      writeRow = merged.row;
      writeCol = merged.col;
    }
    if (kaInline) {
      writeFormXIVKAInlineOnLabelRow(
        worksheet,
        writeRow,
        labelCol || writeCol,
        value,
        officialLabel
      );
      return;
    }
    if (value === '') {
      worksheet.getCell(writeRow, writeCol).value = '';
      return;
    }
    if (labelCol != null && labelCol >= 1 && labelCol < writeCol && variant !== 'gj') {
      for (let c = labelCol + 1; c < writeCol; c += 1) {
        const between = formXIVMPExcelCellValueToString(worksheet.getCell(writeRow, c)?.value).trim();
        if (!between) continue;
        if (!isFormXIVPlaceholderCell(between) && !isFormXIVLabelContinuationCell(between)) {
          continue;
        }
        worksheet.getCell(writeRow, c).value = null;
      }
    }
    setFormXIVMPWorkmanCellValue(worksheet, writeRow, writeCol, value, { variant, inBox: false });
  });
}

const getFormXIVKAWorkmanFieldValue = (row, headers, specIndex) => {
  const spec = FORM_XIV_MP_WORKMAN_FIELD_SPECS[specIndex];
  if (!spec) return '';
  const hdrs = Array.isArray(headers) ? headers : [];
  const header = hdrs.find((h) => spec.rowTest(h)) || spec.label;
  let value = getFormXIVMPRowValueForHeader(row, header);
  if (specIndex === 0 && !value) {
    value = String(row?.__employeeLookupName || '').trim();
  }
  if (specIndex === 2 && value) {
    const nameSpec = FORM_XIV_MP_WORKMAN_FIELD_SPECS[0];
    const nameHeader = hdrs.find((h) => nameSpec.rowTest(h)) || nameSpec.label;
    const nameValue =
      getFormXIVMPRowValueForHeader(row, nameHeader) ||
      String(row?.__employeeLookupName || '').trim();
    if (nameValue && personNamesMatch(value, nameValue)) {
      return '';
    }
  }
  return value;
};

export function writeFormXIVMPWorkmanFieldsToWorksheet(
  worksheet,
  row,
  headers,
  cachedLayout = null,
  parsedFormHeader = null
) {
  if (!worksheet || !row || typeof row !== 'object') return;
  const explicitVariant = String(parsedFormHeader?.formXIVVariant || '').toLowerCase();
  if (explicitVariant === 'gj') {
    ensureFormXIVGJBoxedWorksheetLayout(worksheet, parsedFormHeader);
  } else if (explicitVariant === 'mp') {
    ensureFormXIVMPStackedWorksheetLayout(worksheet, parsedFormHeader);
  }
  const layout =
    cachedLayout || resolveFormXIVMPWorkmanFieldPositions(worksheet, headers, parsedFormHeader);
  const kaInline =
    explicitVariant !== 'mp' &&
    explicitVariant !== 'gj' &&
    (layout.kaInline || isFormXIVKarnatakaExport(worksheet, parsedFormHeader));
  const mpOfficialCard =
    explicitVariant !== 'ka' &&
    explicitVariant !== 'gj' &&
    explicitVariant !== 'rj' &&
    detectFormXIVMPOfficialEvenRowCard(worksheet);
  if (mpOfficialCard) {
    ensureFormXIVMadhyaPradeshColumnAlignment(worksheet);
    writeFormXIVMPOfficialEvenRowFieldValues(worksheet, {}, row, headers);
    applyFormXIVMPOfficialColumnWidths(worksheet, false);
    return;
  }
  if (kaInline) {
    ensureFormXIVKarnatakaColumnAlignment(worksheet);
    const headerRows = findFormXIVKAHeaderLabelRows(worksheet);
    const start = (headerRows[headerRows.length - 1]?.row || 9) + 2;
    const labelCol = resolveFormXIVKABodyLabelCol(worksheet);
    const hdrs = resolveFormXIVMPTableHeaders(headers, {
      formHeader: parsedFormHeader,
    });
    FORM_XIV_KA_WORKMAN_OFFICIAL.forEach((official, i) => {
      writeFormXIVKASerialAndLabel(
        worksheet,
        start + i,
        official,
        getFormXIVKAWorkmanFieldValue(row, hdrs, i),
        i + 1,
        true,
        labelCol
      );
    });
    return;
  }
  const workmanSpecs = layout.workmanSpecs || resolveFormXIVWorkmanFieldSpecsFromHeaders(headers);
  const effectiveVariant = explicitVariant === 'gj' ? 'gj' : layout.variant || 'mp';
  if (!cachedLayout && effectiveVariant !== 'gj' && effectiveVariant !== 'rj') {
    applyFormXIVMPOfficialTemplateLabels(worksheet);
  }
  if (!cachedLayout) {
    if (effectiveVariant === 'gj') {
      layout.positions.forEach(({ row: posRow, labelCol }) => {
        if (posRow >= 1) clearFormXIVGJWorkmanValueBand(worksheet, posRow, labelCol ?? 1);
      });
    } else if (layout.kaInline) {
      // Label + dotted value stay in the same cell; do not wipe the merged line.
    } else if (layout.stackedLayout) {
      clearFormXIVMPStackedWorkmanValueBand(worksheet, 12, 45, workmanSpecs);
    } else {
      clearFormXIVMPWorkmanValueBand(worksheet, 12, 40, layout.valueCol, workmanSpecs);
    }
  } else {
    clearFormXIVMPWorkmanFieldPositions(worksheet, layout.positions, {
      variant: effectiveVariant,
      kaInline: layout.kaInline,
    });
  }
  writeFormXIVMPWorkmanFieldsViaPositions(worksheet, row, layout.positions, {
    variant: effectiveVariant,
    kaInline: layout.kaInline,
  });
  if (layout.kaInline) {
    ensureFormXIVStackedColumnAlignment(worksheet, { formXIVVariant: 'ka' });
  } else if (effectiveVariant !== 'gj' && effectiveVariant !== 'rj') {
    ensureFormXIVStackedColumnAlignment(worksheet, { formXIVVariant: 'mp' });
  }
}

/** Preserve manual edits from the modal grid when rebuilding all employee rows for export. */
export function mapFormXIVMPRowsFromEmployees(employees, headers, helpers = {}) {
  const hdrs = resolveFormXIVMPMappingHeaders(headers, helpers);
  const list = Array.isArray(employees) ? employees : [];
  const {
    sanitizeValue = (v) => String(v ?? '').trim(),
    selectedMonth = '',
    wagePeriodLine = '',
    formatStatutoryDateDisplay = (v) => String(v || '').trim(),
    item = null,
    resolvePayrollRow = null,
    monthCandidates = null,
    fileName = '',
    formHeader = null,
    parsedFormHeader = null,
    sheetText = '',
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
      monthCandidates,
      fileName,
      formHeader: formHeader ?? parsedFormHeader,
      sheetText,
    });
  });
}

export function overlayFormXIVMPUserEditsOntoRows(mappedRows, tableRows, headers) {
  const hdrs = resolveFormXIVMPTableHeaders(headers);
  const savedRows = Array.isArray(tableRows) ? tableRows : [];
  const nameHdr = hdrs.find(isFormXIVMPWorkmanNameHeader);
  const natureHdr = hdrs.find(isFormXIVMPNatureDesignationHeader);
  const savedNameOf = (saved) =>
    nameHdr ? getFormXIVMPRowValueForHeader(saved, nameHdr) : '';
  return (Array.isArray(mappedRows) ? mappedRows : []).map((row, index) => {
    const mappedName = nameHdr ? getFormXIVMPRowValueForHeader(row, nameHdr) : '';
    const savedByName =
      mappedName &&
      savedRows.find((saved) => {
        const savedName = savedNameOf(saved);
        return savedName && personNamesMatch(savedName, mappedName);
      });
    const saved = savedByName || savedRows[index];
    if (!saved || typeof saved !== 'object' || !rowHasMeaningfulFormXIVMPExportData(saved, hdrs)) {
      return row;
    }
    const merged = { ...row };
    hdrs.forEach((header) => {
      const value = getFormXIVMPRowValueForHeader(saved, header);
      if (value === '') return;
      if (
        natureHdr &&
        header === natureHdr &&
        mappedName &&
        personNamesMatch(value, mappedName)
      ) {
        return;
      }
      merged[header] = saved[header] ?? value;
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
  preferredSheetName = '',
  rowItem = null,
  sheetText = '',
}) {
  const hdrs = resolveFormXIVMPTableHeaders(headersToUse, {
    formHeader: parsedFormHeader,
    fileName: formFileName,
    item: rowItem,
    sheetText,
  });
  const resolvedVariant = resolveFormXIVExportVariant(
    parsedFormHeader,
    rowItem,
    formFileName,
    sheetText,
    hdrs
  );
  const isRajasthanTableLayout = resolvedVariant === 'rj';
  const parsedFormHeaderWithVariant = {
    ...(parsedFormHeader || {}),
    formXIVVariant: resolvedVariant,
  };

  const isolatedBuffer = await isolateFormXIVEmploymentCardTemplateBuffer(templateArrayBuffer, {
    preferredSheetName,
    fileName: formFileName,
    formFileName,
    formName: parsedFormHeaderWithVariant?.title,
    formHeader: parsedFormHeaderWithVariant,
    item: rowItem,
  });

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(isolatedBuffer);
  const worksheet = resolveFormXIVMPExcelJsWorksheet(workbook, {
    preferredSheetName,
    fileName: formFileName,
    formFileName,
    formHeader: parsedFormHeaderWithVariant,
    item: rowItem,
  });
  if (!worksheet) throw new Error('Template worksheet not found.');

  const sourceRows = (Array.isArray(mappedData) ? mappedData : []).filter((row) =>
    rowHasMeaningfulFormXIVMPExportData(row, hdrs)
  );
  const employeeRow = sourceRows.length === 1 ? sourceRows[0] : sourceRows[0] || null;

  const headerFormDataForWrite = withFormXIVGJNatureLocationDesignation(
    headerFormData,
    employeeRow,
    hdrs,
    resolvedVariant
  );
  writeFormXIVMPHeaderFieldsToWorksheet(worksheet, headerFormDataForWrite, parsedFormHeaderWithVariant);
  if (employeeRow) {
    if (isRajasthanTableLayout) {
      const tableLayout = resolveFormXRJTableExportLayout(worksheet, hdrs);
      if (tableLayout) {
        clearFormXRJTableDataRows(worksheet, tableLayout);
        writeFormXIVRJTableRowToWorksheet(worksheet, employeeRow, tableLayout);
      }
    } else {
      writeFormXIVMPWorkmanFieldsToWorksheet(
        worksheet,
        employeeRow,
        hdrs,
        null,
        parsedFormHeaderWithVariant
      );
    }
  }
  ensureFormXIVStackedColumnAlignment(worksheet, parsedFormHeaderWithVariant);

  const out = await workbook.xlsx.writeBuffer();
  const variantTag =
    resolvedVariant === 'rj'
      ? 'Form_X_RJ'
      : resolvedVariant === 'ka'
        ? 'Form_XIV_KA'
        : resolvedVariant === 'gj'
          ? 'Form_XIV_GJ'
          : 'Form_XIV_MP';
  const fileName =
    formFileName ||
    parsedFormHeader?.title?.replace(/[^a-zA-Z0-9]/g, '_') ||
    `${variantTag}_${Date.now()}.xlsx`;
  const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  return { blob, fileName, buffer: out };
}

const FORM_XIV_MP_FAST_ZIP_BATCH = 24;

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

const formXIVMPUpsertInlineStrCell = (sheetXml, cellRef, value, options = {}) => {
  const { preserveStyle = false } = options;
  const text = formXIVMPEscapeXml(String(value ?? '').trim());
  const cellRe = new RegExp(`<c\\s+r="${cellRef}"([^>/]*)(?:/>|>([\\s\\S]*?)</c>)`, 'i');
  const existing = sheetXml.match(cellRe);
  let styleAttr = '';
  if (preserveStyle && existing) {
    const styleMatch = String(existing[1] || '').match(/\ss="(\d+)"/i);
    if (styleMatch) styleAttr = ` s="${styleMatch[1]}"`;
  }
  const cellXml = text
    ? `<c r="${cellRef}"${styleAttr} t="inlineStr"><is><t xml:space="preserve">${text}</t></is></c>`
    : `<c r="${cellRef}"${styleAttr}/>`;
  if (existing) {
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
  preferredSheetName = '',
  rowItem = null,
  sheetText = '',
}) {
  const hdrs = resolveFormXIVMPTableHeaders(headersToUse, {
    formHeader: parsedFormHeader,
    fileName: formFileName,
    item: rowItem,
    sheetText,
  });
  const resolvedVariant = resolveFormXIVExportVariant(
    parsedFormHeader,
    rowItem,
    formFileName,
    sheetText,
    hdrs
  );
  const isRajasthanTableLayout = resolvedVariant === 'rj';
  const parsedFormHeaderWithVariant = {
    ...(parsedFormHeader || {}),
    formXIVVariant: resolvedVariant,
  };

  const isolatedBuffer = await isolateFormXIVEmploymentCardTemplateBuffer(templateArrayBuffer, {
    preferredSheetName,
    fileName: formFileName,
    formFileName,
    formHeader: parsedFormHeaderWithVariant,
    item: rowItem,
  });

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(isolatedBuffer);
  const worksheet = resolveFormXIVMPExcelJsWorksheet(workbook, {
    preferredSheetName,
    fileName: formFileName,
    formFileName,
    formHeader: parsedFormHeaderWithVariant,
    item: rowItem,
  });
  if (!worksheet) throw new Error('Template worksheet not found.');
  writeFormXIVMPHeaderFieldsToWorksheet(worksheet, headerFormData, parsedFormHeaderWithVariant);
  let positions = [];
  const isKarnatakaFastLayout =
    parsedFormHeaderWithVariant.formXIVVariant === 'ka' ||
    isFormXIVKarnatakaExport(worksheet, parsedFormHeaderWithVariant);
  if (isRajasthanTableLayout) {
    const tableLayout = resolveFormXRJTableExportLayout(worksheet, hdrs);
    if (tableLayout) {
      clearFormXRJTableDataRows(worksheet, tableLayout);
      positions = tableLayout.hdrs.map((headerKey, j) => {
        const col = tableLayout.columnByHeader[j];
        if (!col || col < 1) return null;
        const writePos = resolveGJMergeTopLeft(worksheet, tableLayout.dataStartRow, col);
        return {
          headerKey,
          row: writePos.row,
          col: writePos.col,
          cellRef: formXIVMPToCellRef(writePos.row, writePos.col),
        };
      }).filter(Boolean);
    }
  } else if (isKarnatakaFastLayout) {
    // Stamp official KA labels once, then inject per-employee values into those cells.
    ensureFormXIVKarnatakaColumnAlignment(worksheet);
    const headerRows = findFormXIVKAHeaderLabelRows(worksheet);
    const start = (headerRows[headerRows.length - 1]?.row || 9) + 2;
    const labelCol = resolveFormXIVKABodyLabelCol(worksheet);
    positions = FORM_XIV_KA_WORKMAN_OFFICIAL.map((official, i) => {
      const spec = FORM_XIV_MP_WORKMAN_FIELD_SPECS[i];
      return {
        specIndex: i,
        headerKey: hdrs.find((h) => spec.rowTest(h)) || spec.label,
        row: start + i,
        col: labelCol,
        cellRef: formXIVMPToCellRef(start + i, labelCol),
        kaInline: true,
        inlineLabel: official,
      };
    });
  } else if (String(parsedFormHeaderWithVariant.formXIVVariant || 'mp').toLowerCase() === 'mp') {
    ensureFormXIVMadhyaPradeshColumnAlignment(worksheet);
    positions = FORM_XIV_MP_WORKMAN_FIELD_SPECS.map((spec, i) => ({
      specIndex: i,
      headerKey: hdrs.find((h) => spec.rowTest(h)) || spec.label,
      row: FORM_XIV_MP_OFFICIAL_WORKMAN_ROWS[i],
      col: FORM_XIV_MP_OFFICIAL_VALUE_COL,
      cellRef: formXIVMPToCellRef(
        FORM_XIV_MP_OFFICIAL_WORKMAN_ROWS[i],
        FORM_XIV_MP_OFFICIAL_VALUE_COL
      ),
      kaInline: false,
    }));
  } else {
    const workmanLayout = resolveFormXIVMPWorkmanFieldPositions(
      worksheet,
      hdrs,
      parsedFormHeaderWithVariant
    );
    clearFormXIVMPWorkmanFieldPositions(worksheet, workmanLayout.positions, {
      variant: workmanLayout.variant || 'mp',
    });
    positions = workmanLayout.positions.map((pos) => {
      const writePos =
        workmanLayout.variant === 'gj'
          ? resolveGJMergeTopLeft(worksheet, pos.row, pos.col)
          : { row: pos.row, col: pos.col };
      const labelCol = pos.labelCol || writePos.col;
      const labelRaw = formXIVMPExcelCellValueToString(
        worksheet.getCell(writePos.row, labelCol)?.value
      );
      const resolved = resolveFormXIVKAOfficialLabel(labelRaw);
      return {
        ...pos,
        row: writePos.row,
        col: writePos.col,
        cellRef: formXIVMPToCellRef(writePos.row, writePos.col),
        kaInline: !!workmanLayout.kaInline,
        inlineLabel: resolved.official || stripFormXIVKALeader(labelRaw),
      };
    });
    ensureFormXIVStackedColumnAlignment(worksheet, parsedFormHeaderWithVariant);
  }

  // Gujarat: reinject Nature + Designation per employee (site location alone is on the static sheet).
  if (String(parsedFormHeaderWithVariant.formXIVVariant || '').toLowerCase() === 'gj') {
    const natureSpec = FORM_XIV_MP_HEADER_SPECS.find((s) => s.key === 'form_xiv_mp_nature_location');
    if (natureSpec) {
      const labelPos = findFormXIVGJHeaderLabelPosition(
        worksheet,
        natureSpec,
        Array.isArray(parsedFormHeaderWithVariant?.fields)
          ? parsedFormHeaderWithVariant.fields
          : []
      );
      if (labelPos) {
        const area = resolveFormXIVGJHeaderBoxArea(worksheet, labelPos.row, labelPos.col);
        positions = [
          ...positions,
          {
            cellRef: formXIVMPToCellRef(area.valueRow, area.valueCol),
            gjNatureLocation: true,
            baseNatureText: String(headerFormData?.form_xiv_mp_nature_location || '').trim(),
          },
        ];
      }
    }
  }

  const preparedBuffer = await workbook.xlsx.writeBuffer();
  const templateZip = await JSZip.loadAsync(preparedBuffer);
  const sheetEntry = resolveFormXIVMPWorksheetEntry(templateZip.files);
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
  return { sheetEntry, baseSheetXml, positions, staticFiles, hdrs };
}

export function triggerFormXIVMPZipDownload(blob, fileName) {
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

export async function buildFormXIVMPPerEmployeeDownload({
  templateArrayBuffer,
  mappedData,
  headersToUse,
  parsedFormHeader,
  formFileName,
  headerFormData,
  rowItem = null,
  sheetText = '',
  preferredSheetName = '',
}) {
  const hdrs = resolveFormXIVMPTableHeaders(headersToUse, {
    formHeader: parsedFormHeader,
    fileName: formFileName,
    item: rowItem,
    sheetText,
  });
  const resolvedVariant = resolveFormXIVExportVariant(
    parsedFormHeader,
    rowItem,
    formFileName,
    sheetText,
    hdrs
  );
  const isRajasthanTableLayout = resolvedVariant === 'rj';
  const parsedFormHeaderWithVariant = {
    ...(parsedFormHeader || {}),
    formXIVVariant: resolvedVariant,
  };
  const exportRows = (Array.isArray(mappedData) ? mappedData : []).filter((row) =>
    rowHasMeaningfulFormXIVMPExportData(row, hdrs)
  );
  const workbookArgs = {
    templateArrayBuffer,
    headersToUse: hdrs,
    parsedFormHeader: parsedFormHeaderWithVariant,
    formFileName,
    headerFormData,
    preferredSheetName,
    rowItem,
    sheetText,
  };
  const isKarnatakaCardLayout = parsedFormHeaderWithVariant.formXIVVariant === 'ka';
  const isGujaratCardLayout = parsedFormHeaderWithVariant.formXIVVariant === 'gj';
  const zipFilePrefix = isRajasthanTableLayout
    ? 'Form_X_RJ'
    : isKarnatakaCardLayout
      ? 'Form_XIV_KA'
      : isGujaratCardLayout
        ? 'Form_XIV_GJ'
        : 'Form_XIV_MP';

  // Form X RJ and Karnataka Form XIV: always ZIP (one employment card per employee).
  // MP/GJ keep a single .xlsx only when there is at most one workman row.
  if (!isRajasthanTableLayout && !isKarnatakaCardLayout && exportRows.length <= 1) {
    const rows = exportRows.length === 1 ? exportRows : [];
    return buildFormXIVMPWorkbookWithTemplateStyles({ ...workbookArgs, mappedData: rows });
  }

  const zipExportRows =
    (isRajasthanTableLayout || isKarnatakaCardLayout) && exportRows.length === 0
      ? [{}]
      : exportRows;

  const buildSlowZipDownload = async () => {
    const zip = new JSZip();
    const usedNames = new Map();
    for (let i = 0; i < zipExportRows.length; i += 1) {
      const { blob, buffer } = await buildFormXIVMPWorkbookWithTemplateStyles({
        ...workbookArgs,
        mappedData: [zipExportRows[i]],
      });
      const xlsxBytes =
        buffer != null
          ? buffer
          : typeof blob.arrayBuffer === 'function'
            ? new Uint8Array(await blob.arrayBuffer())
            : blob;
      const baseName = resolveFormXIVMPEmployeeDownloadBaseName(zipExportRows[i], hdrs, i);
      zip.file(
        allocateUniqueFormXIVMPDownloadFileName(
          baseName,
          usedNames,
          parsedFormHeaderWithVariant.formXIVVariant
        ),
        xlsxBytes
      );
      if (i > 0 && i % 10 === 0) {
        await new Promise((resolve) => setTimeout(resolve, 0));
      }
    }
    const zipBase = String(formFileName || parsedFormHeader?.title || zipFilePrefix)
      .replace(/\.xlsx?$/i, '')
      .replace(/[^a-zA-Z0-9._-]+/g, '_');
    return {
      blob: await zip.generateAsync({ type: 'blob', compression: 'STORE' }),
      fileName: `${zipBase}_Employees.zip`,
    };
  };

  try {
    const { sheetEntry, baseSheetXml, positions, staticFiles } = await prepareFormXIVMPFastZipTemplate({
      templateArrayBuffer,
      headersToUse: hdrs,
      parsedFormHeader: parsedFormHeaderWithVariant,
      headerFormData,
      formFileName,
      preferredSheetName,
      rowItem,
      sheetText,
    });
    if (!Array.isArray(positions) || positions.length === 0) {
      return buildSlowZipDownload();
    }

    const zip = new JSZip();
    const usedNames = new Map();
    for (let i = 0; i < zipExportRows.length; i += FORM_XIV_MP_FAST_ZIP_BATCH) {
      const batch = zipExportRows.slice(i, i + FORM_XIV_MP_FAST_ZIP_BATCH);
      const batchBytes = await Promise.all(
        batch.map(async (exportRow, batchIndex) => {
          const index = i + batchIndex;
          let sheetXml = baseSheetXml;
          for (let pi = 0; pi < positions.length; pi += 1) {
            const pos = positions[pi];
            let injectValue = '';
            if (pos.gjNatureLocation) {
              injectValue = buildFormXIVGJNatureLocationWithDesignation(
                pos.baseNatureText,
                exportRow,
                hdrs
              );
            } else {
              const rawValue = Number.isInteger(pos.specIndex)
                ? getFormXIVKAWorkmanFieldValue(exportRow, hdrs, pos.specIndex)
                : getFormXIVMPRowValueForHeader(exportRow, pos.headerKey);
              injectValue = pos.kaInline
                ? formatFormXIVKAFilledLine(pos.inlineLabel || '', rawValue, true)
                : rawValue;
            }
            sheetXml = formXIVMPUpsertInlineStrCell(
              sheetXml,
              pos.cellRef,
              injectValue,
              { preserveStyle: true }
            );
          }
          const entryZip = new JSZip();
          Object.entries(staticFiles).forEach(([path, data]) => {
            entryZip.file(path, data);
          });
          entryZip.file(sheetEntry, sheetXml);
          const xlsxBytes = await entryZip.generateAsync({ type: 'uint8array', compression: 'STORE' });
          return { xlsxBytes, exportRow, index };
        })
      );
      batchBytes.forEach((entry) => {
        if (!entry) return;
        const { xlsxBytes, exportRow, index } = entry;
        const baseName = resolveFormXIVMPEmployeeDownloadBaseName(exportRow, hdrs, index);
        zip.file(
          allocateUniqueFormXIVMPDownloadFileName(
            baseName,
            usedNames,
            parsedFormHeaderWithVariant.formXIVVariant
          ),
          xlsxBytes
        );
      });
      if (i > 0) {
        await new Promise((resolve) => setTimeout(resolve, 0));
      }
    }
    const zipBase = String(formFileName || parsedFormHeader?.title || zipFilePrefix)
      .replace(/\.xlsx?$/i, '')
      .replace(/[^a-zA-Z0-9._-]+/g, '_');
    return {
      blob: await zip.generateAsync({ type: 'blob', compression: 'STORE' }),
      fileName: `${zipBase}_Employees.zip`,
    };
  } catch (_err) {
    return buildSlowZipDownload();
  }
}
