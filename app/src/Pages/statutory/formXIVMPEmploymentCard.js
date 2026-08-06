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
import { filterFormXIXMPEligiblePayrollRows } from './formXIXMPWageSlip';
import { personNamesMatch } from './formFKarnataka';
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
  if (/see\s+rule\s+75/i.test(parts) && /employment\s+card/i.test(parts)) return true;
  if (
    /rajasthan/.test(parts) &&
    /employment\s+card/i.test(parts) &&
    /form[\s._-]*x\b/i.test(parts) &&
    !/form[\s._-]*xiv/i.test(parts) &&
    !/form[\s._-]*xi(?![vix])/i.test(parts)
  ) {
    return true;
  }
  return false;
}

/** Karnataka CLRA Form XIV — Employment Card (stacked MP/Karnataka layout). */
export function isFormXIVKarnatakaContext(formHeader, rowItem, fileName, sheetText = '') {
  if (!isFormXIVMPEmploymentCardContext(formHeader, rowItem, fileName, sheetText)) return false;
  const parts = [
    rowItem?.state,
    rowItem?.State,
    rowItem?.formName,
    rowItem?.FormName,
    fileName,
    formHeader?.title,
    formHeader?.subtitle,
    sheetText,
  ]
    .filter((x) => x != null && String(x).trim() !== '')
    .join(' ')
    .toLowerCase();
  return /karnataka|form[\s._-]*xiv[\s._-]*ka|form_xiv[\s._-]*karnataka/.test(parts);
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
  const parts = [
    rowItem?.state,
    rowItem?.State,
    rowItem?.siteState,
    rowItem?.SiteState,
    rowItem?.formName,
    rowItem?.FormName,
    fileName,
    formHeader?.title,
    formHeader?.subtitle,
    sheetText,
  ]
    .filter((x) => x != null && String(x).trim() !== '')
    .join(' ')
    .toLowerCase();
  if (/gujarat|xiv[\s._-]*gj/i.test(parts)) return false;
  return (
    /madhya[\s._-]*pradesh/.test(parts) ||
    /\bform[\s._-]*xiv[\s._-]*mp\b/.test(parts) ||
    /\bxiv[\s._-]*mp\b/.test(parts) ||
    /form_xiv_mp/.test(parts)
  );
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
      /s\.?\s*no\.?\s+in\s+the\s+register/.test(joined)) &&
    (/nature\s+of\s+employ/.test(joined) || /designat/.test(joined))
  );
}

export function resolveFormXIVVariant(formHeader, rowItem, fileName, sheetText = '', tableHeaders = null) {
  const parts = [
    rowItem?.state,
    rowItem?.State,
    rowItem?.siteState,
    rowItem?.SiteState,
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
  if (isFormXRajasthanEmploymentCardContext(formHeader, rowItem, fileName, sheetText)) {
    return 'rj';
  }
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
  if (payrollRow) {
    const fromNet = readFormXIVMPPayrollNetPay(payrollRow);
    if (fromNet !== '') {
      const n = Number(String(fromNet).replace(/,/g, '').trim());
      if (Number.isFinite(n) && n > 0) return String(n);
      return String(fromNet).trim();
    }
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
      out[header] = sanitizeValue(
        resolveFormXIVMPWageRate(emp, payrollRow, {
          variant,
          monthCandidates: helpers.monthCandidates,
          item,
          fileName: helpers.fileName ?? helpers.formFileName ?? '',
          formHeader: helpers.formHeader ?? helpers.parsedFormHeader ?? null,
          sheetText: helpers.sheetText ?? '',
        })
      );
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
    monthCandidates = null,
    item = null,
    fileName = '',
    formHeader = null,
    parsedFormHeader = null,
    sheetText = '',
  } = helpers;
  const wageRateHints = buildFormXIVKarnatakaWageRateHints({
    item,
    fileName,
    formHeader: formHeader ?? parsedFormHeader,
    sheetText,
  });
  if (!Array.isArray(mappedData) || mappedData.length === 0) return 0;
  let hits = 0;
  mappedData.forEach((row, rowIndex) => {
    const empItem = employees[rowIndex];
    const emp = empItem?.Employee || empItem?.employee || empItem;
    const monthDefault =
      resolveFormXIVKarnatakaMonthDefaultWageRate(emp, monthCandidates, wageRateHints) ||
      resolveFormXIVMadhyaPradeshMonthDefaultWageRate(emp, monthCandidates, wageRateHints);
    if (monthDefault && (overwrite || !String(row?.[wageRateHeader] ?? '').trim())) {
      row[wageRateHeader] = sanitizeValue(monthDefault);
      hits += 1;
      return;
    }
    if (!overwrite && String(row?.[wageRateHeader] ?? '').trim()) return;
    if (typeof resolvePayrollRow !== 'function') return;
    const payrollRow = resolvePayrollRow(emp, row, rowIndex);
    if (!payrollRow || payrollRow.fetch_error) return;
    const rate = resolveFormXIVMPWageRate(emp, payrollRow, {
      variant,
      monthCandidates,
      ...wageRateHints,
    });
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
  const prefix = variant === 'rj' ? 'Form_X_RJ' : 'Form_XIV_MP';
  return `${prefix}_${root}${suffix}.xlsx`;
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

const detectFormXRajasthanWorksheetLayout = (worksheet, maxScanRow = 24) => {
  if (!worksheet) return false;
  let contractorLabelInB = false;
  let tableHeaderRow = false;
  for (let r = 1; r <= maxScanRow; r += 1) {
    const labelB = formXIVMPExcelCellValueToString(worksheet.getCell(r, 2)?.value).trim();
    if (/name\s+and\s+address\s+.*contractor/i.test(labelB)) contractorLabelInB = true;
    for (let c = 1; c <= 10; c += 1) {
      const cellNorm = formXIVMPHeaderNorm(
        formXIVMPExcelCellValueToString(worksheet.getCell(r, c)?.value)
      );
      if (/name\s+of\s+the\s+workman/.test(cellNorm)) {
        tableHeaderRow = true;
        break;
      }
    }
  }
  return contractorLabelInB && tableHeaderRow;
};

const resolveFormXRJHeaderValueCell = (worksheet, labelRow, labelCol) => {
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
    if (isFormXIVPlaceholderCell(cellStr)) worksheet.getCell(row, col).value = '';
    return;
  }
  for (let r = merge.top; r <= merge.bottom; r += 1) {
    for (let c = merge.left; c <= merge.right; c += 1) {
      const cellStr = formXIVMPExcelCellValueToString(worksheet.getCell(r, c)?.value).trim();
      if (!cellStr || isFormXIVPlaceholderCell(cellStr)) worksheet.getCell(r, c).value = '';
    }
  }
};

const writeFormXIVRJBoxedHeaderValue = (worksheet, labelRow, labelCol, value) => {
  const text = String(value ?? '').trim();
  if (!text || labelRow < 1) return;
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
    if (isFormXIVMPWageRateHeader(header) && /wage/.test(cellNorm) && /rate|piece|place/.test(cellNorm)) return c;
    if (isFormXIVMPWagePeriodHeader(header) && /wage\s+period/.test(cellNorm)) return c;
    if (isFormXIVMPTenureHeader(header) && /tenure|period\s+of\s+employ/.test(cellNorm)) return c;
    if (isFormXIVMPRemarksHeader(header) && /remark/.test(cellNorm)) return c;
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
    return { top: merge.top, bottom: merge.bottom, left: merge.left, right: merge.right };
  }
  return { top: topLeft.row, bottom: topLeft.row + 3, left: topLeft.col, right: topLeft.col };
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
    const val = getFormXIVMPRowValueForHeader(row, header);
    if (!val) return;
    setFormXRJTableBoxCellValue(worksheet, dataStartRow, col, val);
  });
};

const clearFormXRJTableDataRows = (worksheet, layout) => {
  if (!worksheet || !layout) return;
  const { dataStartRow, columnByHeader } = layout;
  const usedCols = [...new Set(columnByHeader.filter((col) => col > 0))];
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
    form_xiv_mp_contractor: { row: 5, col: 1 },
    form_xiv_mp_establishment: { row: 5, col: 8 },
    form_xiv_mp_nature_location: { row: 14, col: 1 },
    form_xiv_mp_principal_employer: { row: 14, col: 8 },
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
  const rjBoxedLayout = variant === 'rj' || detectFormXRajasthanWorksheetLayout(worksheet);
  const stackedLayout = !gjBoxedLayout && !rjBoxedLayout && detectFormXIVStackedWorkmanLayout(worksheet);
  const defaultValueCol = stackedLayout
    ? FORM_XIV_MP_STACKED_VALUE_COL
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
    const targetCol =
      labelCol != null && labelCol >= 1
        ? findFormXIVValueColumnBesideLabel(worksheet, excelRow, labelCol, defaultValueCol, workmanSpecs)
        : defaultValueCol;
    writeAt(excelRow, targetCol, value);
  };

  FORM_XIV_MP_HEADER_SPECS.forEach((spec) => {
    const val = headerFormData[spec.key];
    if (val == null || String(val).trim() === '') return;

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
  });
  if (gjBoxedLayout) {
    restoreFormXIVGJHeaderLabels(worksheet, gjHeaderLabelSnapshots, parsedFields);
  }
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

export function clearFormXIVMPWorkmanFieldPositions(worksheet, positions, options = {}) {
  if (!worksheet || !Array.isArray(positions)) return;
  const variant = options.variant || 'mp';
  positions.forEach(({ row, col, labelCol }) => {
    if (row < 1 || col < 1) return;
    if (variant === 'gj') {
      clearFormXIVGJWorkmanValueBand(worksheet, row, labelCol ?? 1);
      return;
    }
    worksheet.getCell(row, col).value = '';
  });
}

export function writeFormXIVMPWorkmanFieldsViaPositions(worksheet, row, positions, options = {}) {
  if (!worksheet || !row || typeof row !== 'object' || !Array.isArray(positions)) return;
  const variant = options.variant || 'mp';
  positions.forEach(({ headerKey, row: targetRow, col, labelCol }) => {
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
    if (value === '') {
      worksheet.getCell(writeRow, writeCol).value = '';
      return;
    }
    if (labelCol != null && labelCol >= 1 && labelCol < writeCol && variant !== 'gj') {
      for (let c = labelCol + 1; c < writeCol; c += 1) {
        const between = formXIVMPExcelCellValueToString(worksheet.getCell(writeRow, c)?.value).trim();
        if (between && !isFormXIVPlaceholderCell(between) && !isFormXIVLabelContinuationCell(between)) {
          continue;
        }
        worksheet.getCell(writeRow, c).value = '';
      }
    }
    setFormXIVMPWorkmanCellValue(worksheet, writeRow, writeCol, value, { variant, inBox: false });
  });
}

export function writeFormXIVMPWorkmanFieldsToWorksheet(worksheet, row, headers, cachedLayout = null) {
  if (!worksheet || !row || typeof row !== 'object') return;
  const layout = cachedLayout || resolveFormXIVMPWorkmanFieldPositions(worksheet, headers);
  const workmanSpecs = layout.workmanSpecs || resolveFormXIVWorkmanFieldSpecsFromHeaders(headers);
  if (!cachedLayout) {
    if (layout.variant === 'gj') {
      layout.positions.forEach(({ row, labelCol }) => {
        if (row >= 1) clearFormXIVGJWorkmanValueBand(worksheet, row, labelCol ?? 1);
      });
    } else if (layout.stackedLayout) {
      clearFormXIVMPStackedWorkmanValueBand(worksheet, 12, 45, workmanSpecs);
    } else {
      clearFormXIVMPWorkmanValueBand(worksheet, 12, 40, layout.valueCol, workmanSpecs);
    }
  } else {
    clearFormXIVMPWorkmanFieldPositions(worksheet, layout.positions, {
      variant: layout.variant || 'mp',
    });
  }
  writeFormXIVMPWorkmanFieldsViaPositions(worksheet, row, layout.positions, {
    variant: layout.variant || 'mp',
  });
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
  preferredSheetName = '',
  rowItem = null,
  sheetText = '',
}) {
  const hdrs = resolveFormXIVMPTableHeaders(headersToUse, {
    formHeader: parsedFormHeader,
    fileName: formFileName,
  });
  const resolvedVariant = resolveFormXIVVariant(
    parsedFormHeader,
    rowItem,
    formFileName,
    sheetText,
    hdrs
  );
  const isRajasthanTableLayout =
    String(parsedFormHeader?.formXIVVariant || '').toLowerCase() === 'rj' ||
    resolvedVariant === 'rj' ||
    isFormXRajasthanEmploymentCardContext(parsedFormHeader, rowItem, formFileName, sheetText);
  const parsedFormHeaderWithVariant = {
    ...(parsedFormHeader || {}),
    formXIVVariant: isRajasthanTableLayout ? 'rj' : parsedFormHeader?.formXIVVariant || resolvedVariant,
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

  writeFormXIVMPHeaderFieldsToWorksheet(worksheet, headerFormData, parsedFormHeaderWithVariant);
  if (employeeRow) {
    if (isRajasthanTableLayout) {
      const tableLayout = resolveFormXRJTableExportLayout(worksheet, hdrs);
      if (tableLayout) {
        clearFormXRJTableDataRows(worksheet, tableLayout);
        writeFormXIVRJTableRowToWorksheet(worksheet, employeeRow, tableLayout);
      }
    } else {
      writeFormXIVMPWorkmanFieldsToWorksheet(worksheet, employeeRow, hdrs);
    }
  }

  const out = await workbook.xlsx.writeBuffer();
  const fileName =
    formFileName ||
    parsedFormHeader?.title?.replace(/[^a-zA-Z0-9]/g, '_') ||
    (isRajasthanTableLayout ? `Form_X_RJ_${Date.now()}.xlsx` : `Form_XIV_MP_${Date.now()}.xlsx`);
  const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  return { blob, fileName };
}

const FORM_XIV_MP_FAST_ZIP_BATCH = 12;

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
  });
  const resolvedVariant = resolveFormXIVVariant(
    parsedFormHeader,
    rowItem,
    formFileName,
    sheetText,
    hdrs
  );
  const isRajasthanTableLayout =
    String(parsedFormHeader?.formXIVVariant || '').toLowerCase() === 'rj' ||
    resolvedVariant === 'rj' ||
    isFormXRajasthanEmploymentCardContext(parsedFormHeader, rowItem, formFileName, sheetText);
  const parsedFormHeaderWithVariant = {
    ...(parsedFormHeader || {}),
    formXIVVariant: isRajasthanTableLayout ? 'rj' : parsedFormHeader?.formXIVVariant || resolvedVariant,
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
  } else {
    const workmanLayout = resolveFormXIVMPWorkmanFieldPositions(worksheet, hdrs);
    clearFormXIVMPWorkmanFieldPositions(worksheet, workmanLayout.positions, {
      variant: workmanLayout.variant || 'mp',
    });
    positions = workmanLayout.positions.map((pos) => {
      const writePos =
        workmanLayout.variant === 'gj'
          ? resolveGJMergeTopLeft(worksheet, pos.row, pos.col)
          : { row: pos.row, col: pos.col };
      return {
        ...pos,
        row: writePos.row,
        col: writePos.col,
        cellRef: formXIVMPToCellRef(writePos.row, writePos.col),
      };
    });
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
  });
  const resolvedVariant = resolveFormXIVVariant(
    parsedFormHeader,
    rowItem,
    formFileName,
    sheetText,
    hdrs
  );
  // Prefer live RJ detection over a stale formXIVVariant (e.g. cached 'mp' on Form_X_RJ).
  const isRajasthanTableLayout =
    String(parsedFormHeader?.formXIVVariant || '').toLowerCase() === 'rj' ||
    resolvedVariant === 'rj' ||
    isFormXRajasthanEmploymentCardContext(parsedFormHeader, rowItem, formFileName, sheetText);
  const parsedFormHeaderWithVariant = {
    ...(parsedFormHeader || {}),
    formXIVVariant: isRajasthanTableLayout ? 'rj' : parsedFormHeader?.formXIVVariant || resolvedVariant,
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
  const zipFilePrefix = isRajasthanTableLayout ? 'Form_X_RJ' : 'Form_XIV_MP';

  // Form X RJ: always ZIP (even for 0–1 employees). MP/GJ keep single .xlsx for ≤1.
  if (!isRajasthanTableLayout && exportRows.length <= 1) {
    const rows = exportRows.length === 1 ? exportRows : [];
    return buildFormXIVMPWorkbookWithTemplateStyles({ ...workbookArgs, mappedData: rows });
  }

  const zipExportRows =
    isRajasthanTableLayout && exportRows.length === 0 ? [{}] : exportRows;

  const buildSlowZipDownload = async () => {
    const zip = new JSZip();
    const usedNames = new Map();
    for (let i = 0; i < zipExportRows.length; i += 1) {
      const { blob } = await buildFormXIVMPWorkbookWithTemplateStyles({
        ...workbookArgs,
        mappedData: [zipExportRows[i]],
      });
      const xlsxBytes = new Uint8Array(await blob.arrayBuffer());
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
            sheetXml = formXIVMPUpsertInlineStrCell(
              sheetXml,
              pos.cellRef,
              getFormXIVMPRowValueForHeader(exportRow, pos.headerKey),
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
