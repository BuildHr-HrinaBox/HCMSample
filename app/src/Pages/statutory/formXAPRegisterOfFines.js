import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import {
  applyExcelJSDataRowBorders,
  applyExcelJSFullBoxBordersToRange,
  countExcelJSTemplateBodyRows,
  ensureExcelJSDataRowsWithBorders,
  excelJSCellHasBorder,
  excelJSCellHasFullBoxBorder,
  findExcelJSRemarksOrLastHeaderCol,
} from '../../utils/excelTableBorders';
import {
  formatStatutoryHeaderLabelValueExport,
  normalizeStatutoryHeaderLabel,
  resolveHeaderFieldExportValue,
  statutoryHeaderLabelMatchKey,
  writeStatutoryHeaderFieldsToExcelJsWorksheet,
} from '../../utils/statutorySiteCompanyHeaders';
export {
  FORM_XXI_AP_FINE_COLUMN_NIL_TEXT,
  applyFormXXIAPFinesNilToMappedRows,
  applyFormXXIAPFinesNilToRow,
  isFormXXIAPFineNilHeader,
  isFormXXIAPSkipAutofillHeader,
} from './formXXIAPFinesNil';

/** AP Shops Form X — Register of Fines (Rules under Payment of Wages / Minimum Wages / S&E). */

export const FORM_X_AP_FINES_NIL_OF_MONTH_TEXT = 'Nill of the month';

export const FORM_XX_AP_DEDUCTIONS_NIL_OF_MONTH_TEXT = 'Nill of the month';

/** Default cell value for Form XX damage / recovery columns when no deduction case exists. */
export const FORM_XX_AP_DEDUCTION_COLUMN_NIL_TEXT = 'NIL';

export const FORM_X_AP_TEMPLATE_MISMATCH_MESSAGE =
  'This row is Form X (Shops & Establishment — Register of Fines) but the template file is Form XXI (Contract Labour — Register of Fines). Upload the correct Form X template in Form Master.';

/** Must not match the "name" inside "surname of workmen" (Form XIII false positive for Form XXI). */
export const CLRA_WORKMEN_NAME_HEADER_RE = /\bname\s+of\s+(?:the\s+)?(?:workmen|wokmen)\b/i;

export const REGISTER_OF_WORKMEN_SURNAME_HEADER_RE = /name\s+and\s+surname\s+of\s+workmen/i;

const MULTI_X_FORM_RE =
  /form[\s._-]*(xii|xiii|xiv|xv|xvi|xvii|xviii|xix|xx|xxi|xxii|xxiii|xxiv|xxv|xxvi|xxvii)\b/i;

const FORM_XXI_COLLISION_ROMAN_RE =
  /form[\s._-]*(xiii|xiv|xv|xvi|xvii|xviii|xix|xx|xxii|xxiii|xxiv|xxv|xxvi|xxvii)(?![a-z])/i;

const FORM_XX_COLLISION_ROMAN_RE =
  /form[\s._-]*(xix|xxi|xxii|xxiii|xxiv|xxv|xxvi|xxvii|xxviii|xxix|xxx)(?![a-z])/i;

function blobIndicatesFormXIXHint(blob) {
  const parts = String(blob || '').toLowerCase();
  if (/form[\s._-]*xxix(?![a-z])/i.test(parts)) return false;
  return (
    /form[\s._-]*xix(?![a-z])/i.test(parts) ||
    /form[\s._-]*19(?!\d)/i.test(parts) ||
    /(?:^|[\s._-])xix(?=[\s._\W-]|$)/i.test(parts)
  );
}

export function blobIndicatesFormXXIXHint(blob) {
  const parts = String(blob || '').toLowerCase();
  return (
    /form[\s._-]*xxix(?![a-z])/i.test(parts) ||
    /form[\s._-]*29(?!\d)/i.test(parts) ||
    /(?:^|[\s._-])xxix(?=[\s._\W-]|$)/i.test(parts)
  );
}

export function matchesFormXXHint(blob) {
  const parts = String(blob || '').toLowerCase();
  if (FORM_XX_COLLISION_ROMAN_RE.test(parts)) return false;
  return (
    /form[\s._-]*xx(?![a-z])/i.test(parts) ||
    /form[\s._-]*20(?!\d)/i.test(parts) ||
    /(?:^|[\s._-])xx(?=[\s._\W-]|$)/i.test(parts)
  );
}

export function blobIndicatesRegisterOfDeductionsForDamage(blob, tableHeaders = null) {
  const text = String(blob || '').toLowerCase();
  if (/register\s+of\s+deductions/i.test(text)) return true;
  if (/deductions?\s+(for\s+)?damage|damage\s+or\s+loss/i.test(text)) return true;
  if (/particulars\s+of\s+damage|date\s+of\s+damage/i.test(text)) return true;
  if (Array.isArray(tableHeaders) && tableHeaders.length > 0) {
    const hdr = tableHeaders.map((h) => String(h || '').toLowerCase()).join('\n');
    if (/register\s+of\s+deductions/i.test(hdr)) return true;
    if (/particulars\s+of\s+damage|date\s+of\s+damage|deductions?\s+for\s+damage/i.test(hdr)) {
      return true;
    }
  }
  return false;
}

/** AP CLRA Form XX — Register of Deductions for Damage or Loss (not Form XXI Register of Fines). */
export function isFormXXAPRegisterOfDeductionsContext(
  formHeader,
  rowItem,
  fileName,
  sheetText = '',
  tableHeaders = null
) {
  const parts = [
    rowItem?.formName,
    rowItem?.FormName,
    rowItem?.description,
    rowItem?.Description,
    rowItem?.act,
    rowItem?.Act,
    rowItem?.formFileName,
    rowItem?.FormFileName,
    fileName,
    formHeader?.title,
    formHeader?.subtitle,
    formHeader?.reference,
    sheetText
  ]
    .filter((x) => x != null && String(x).trim() !== '')
    .join(' ')
    .toLowerCase();

  if (blobIndicatesFormXIXHint(parts)) return false;
  if (blobIndicatesFormXXIXHint(parts)) return false;

  if (matchesFormXXHint(parts)) return true;
  return blobIndicatesRegisterOfDeductionsForDamage(parts, tableHeaders);
}

function formXXAPHeaderBare(header) {
  return String(header || '')
    .toLowerCase()
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .replace(/^\(?\d+\)?\s*/, '')
    .replace(/\s*\(\d+\)\s*$/, '')
    .trim();
}

/**
 * Form XX damage / cause / recovery columns that default to NIL when no case exists
 * (Particulars, Date of damage, showed cause, explanation witness, amount, instalments, remarks).
 * Never People-autofill these — download Excel/PDF must show NIL, not employee names.
 */
export function isFormXXAPDeductionNilHeader(header) {
  const bare = formXXAPHeaderBare(header);
  if (!bare) return false;
  // Never treat workmen / father columns as NIL targets.
  if (
    /\bname\s+of\s+(?:the\s+)?workmen\b|\bname\s+of\s+workman\b/.test(bare) ||
    (/^name\s+of\b/.test(bare) &&
      /workmen|workman|worker|employee/.test(bare) &&
      !/presence|explanation/.test(bare))
  ) {
    return false;
  }
  if (/father|husband/.test(bare) && !/presence|explanation|showed\s+cause/.test(bare)) {
    return false;
  }
  if (/particulars\s+of\s+damage|date\s+of\s+damage|amount\s+of\s+deduction/.test(bare)) return true;
  if (/date\s+of\s+recovery/.test(bare)) return true;
  if (
    /show\s+cause|showed\s+cause|whether\s+work\s*man|whether\s+workman|against\s+deduction|total\s+amount/.test(
      bare
    )
  ) {
    return true;
  }
  // instalment (UK) and installment (US)
  if (/install?ments?/.test(bare)) return true;
  if (bare === 'first' || bare === 'last') return true;
  if (/\bremarks?\b/.test(bare)) return true;
  if (
    /name\s+of\s+person.*presence|presence.*explanation|explanation\s+was\s+heard|whose\s+presence/.test(
      bare
    )
  ) {
    return true;
  }
  if (/^first\s+install?ment|^last\s+install?ment|^no\.?\s*of\s+install?ment/.test(bare)) return true;
  return false;
}

/** Skip People / payroll autofill for Form XX NIL columns. */
export function isFormXXAPSkipAutofillHeader(header) {
  return isFormXXAPDeductionNilHeader(header);
}

export function applyFormXXAPDeductionsNilToRow(row, headers, helpers = {}) {
  const hdrs = Array.isArray(headers) ? headers : [];
  const out = row && typeof row === 'object' ? { ...row } : {};
  const nilText =
    helpers.nilText != null ? String(helpers.nilText) : FORM_XX_AP_DEDUCTION_COLUMN_NIL_TEXT;
  const { overwrite = false } = helpers;

  hdrs.forEach((header) => {
    if (!isFormXXAPDeductionNilHeader(header)) return;
    const existing = String(out[header] ?? '').trim();
    if (!overwrite && existing && !/^enter\b/i.test(existing) && !/^nil+$/i.test(existing)) return;
    out[header] = nilText;
  });
  return out;
}

export function applyFormXXAPDeductionsNilToMappedRows(
  mappedData,
  headers,
  nilText = FORM_XX_AP_DEDUCTION_COLUMN_NIL_TEXT,
  helpers = {}
) {
  if (!Array.isArray(mappedData)) return [];
  const { overwrite = false } = helpers;
  return mappedData.map((row) =>
    applyFormXXAPDeductionsNilToRow(row, headers, { nilText, overwrite })
  );
}

export function matchesFormXIIIHint(blob) {
  const parts = String(blob || '').toLowerCase();
  if (/form[\s._-]*xxiii(?![a-z])/i.test(parts)) return false;
  if (/(?:^|[\s._-])xxiii(?=[\s._\W-]|$)/i.test(parts)) return false;
  if (matchesFormXIVHint(parts)) return false;
  return (
    /form[\s._-]*xiii(?![a-z])/i.test(parts) ||
    /form[\s._-]*13(?!\d)/i.test(parts) ||
    /(?:^|[\s._-])xiii(?=[\s._\W-]|$)/i.test(parts)
  );
}

/** CLRA Form XIV — Employment Card (Roman XIV; distinct from AP Form 14 child workers). */
export function matchesFormXIVHint(blob) {
  const parts = String(blob || '').toLowerCase();
  if (/form[\s._-]*xxiv(?![a-z])/i.test(parts)) return false;
  if (/(?:^|[\s._-])xxiv(?=[\s._\W-]|$)/i.test(parts)) return false;
  return (
    /form[\s._-]*xiv(?![a-z])/i.test(parts) ||
    /(?:^|[\s._-])xiv(?=[\s._\W-]|$)/i.test(parts)
  );
}

/**
 * CLRA Form XI only (e.g. Form_XI_RJ Service Certificate).
 * Must not match XII / XIII / XIV / XV / XVI / XVII / XVIII / XIX.
 */
export function matchesFormXIHint(blob) {
  const parts = String(blob || '').toLowerCase();
  if (matchesFormXIVHint(parts)) return false;
  if (/form[\s._-]*x(?:i{2,}|[vx])/i.test(parts) && !/form[\s._-]*xi(?![vix])/i.test(parts)) {
    return false;
  }
  return (
    /form[\s._-]*xi(?![vix])/i.test(parts) ||
    /\bxi_rj\b/i.test(parts) ||
    /(?:^|[\s._-])xi(?![vix])(?=[\s._\W-]|$)/i.test(parts)
  );
}

export function blobIndicatesEmploymentCard(blob, tableHeaders = null) {
  const text = String(blob || '').toLowerCase();
  if (/employment\s+card/i.test(text)) return true;
  if (/see\s+rule\s+76/i.test(text)) return true;
  if (/serial\s+number\s+in\s+the\s+register\s+of\s+workmen/i.test(text)) return true;
  if (Array.isArray(tableHeaders) && tableHeaders.length > 0) {
    const hdr = tableHeaders.map((h) => String(h || '').toLowerCase()).join('\n');
    if (/employment\s+card/i.test(hdr)) return true;
    if (/serial\s+number\s+in\s+the\s+register/.test(hdr)) return true;
  }
  return false;
}

/** CLRA Form XIV — Employment Card (not Form XIII Register of Workmen). */
export function isFormXIVEmploymentCardContext(
  formHeader,
  rowItem,
  fileName,
  sheetText = '',
  tableHeaders = null
) {
  // Row identity (form name / file / description) wins over sheet body — Form_XI_RJ must not open as XIV.
  const identityBlob = [
    rowItem?.formName,
    rowItem?.FormName,
    rowItem?.description,
    rowItem?.Description,
    fileName,
  ]
    .filter((x) => x != null && String(x).trim() !== '')
    .join(' ')
    .toLowerCase();

  const parts = [
    identityBlob,
    formHeader?.title,
    formHeader?.subtitle,
    formHeader?.reference,
    sheetText,
  ]
    .filter((x) => x != null && String(x).trim() !== '')
    .join(' ')
    .toLowerCase();

  if (/form[\s._-]*xxiv(?![a-z])/i.test(parts)) return false;

  // Explicit Form XI / Service Certificate row — never force Employment Card (Form XIV) layout.
  if (!matchesFormXIVHint(identityBlob)) {
    if (matchesFormXIHint(identityBlob)) return false;
    if (/service\s+certificate/i.test(identityBlob) && !/employment\s+card/i.test(identityBlob)) {
      return false;
    }
  }

  if (matchesFormXIVHint(identityBlob) || matchesFormXIVHint(parts)) return true;
  return blobIndicatesEmploymentCard(parts, tableHeaders);
}

export function blobIndicatesRegisterOfWorkmen(blob, tableHeaders = null) {
  const text = String(blob || '').toLowerCase();
  if (matchesFormXIVHint(text) || blobIndicatesEmploymentCard(text, tableHeaders)) return false;
  if (/register\s+of\s+workmen/i.test(text)) return true;
  if (Array.isArray(tableHeaders) && tableHeaders.length > 0) {
    const hdr = tableHeaders.map((h) => String(h || '').toLowerCase()).join('\n');
    if (REGISTER_OF_WORKMEN_SURNAME_HEADER_RE.test(hdr)) return true;
    if (/age\s+and\s+sex/.test(hdr) && /local\s+address/.test(hdr) && /permanent\s+house\s+address/.test(hdr)) {
      return true;
    }
  }
  return false;
}

/** AP CLRA Form XIII — Register of Workmen (not Form XXI Register of Fines). */
export function isFormXIIIRegisterOfWorkmenContext(
  formHeader,
  rowItem,
  fileName,
  sheetText = '',
  tableHeaders = null
) {
  // Row identity (form name / file) wins over sheet body — Form_14_RJ must not open as XIII.
  const identityBlob = [
    rowItem?.formName,
    rowItem?.FormName,
    rowItem?.description,
    rowItem?.Description,
    rowItem?.formFileName,
    rowItem?.FormFileName,
    fileName,
  ]
    .filter((x) => x != null && String(x).trim() !== '')
    .join(' ')
    .toLowerCase();

  const parts = [
    identityBlob,
    formHeader?.title,
    formHeader?.subtitle,
    formHeader?.reference,
    sheetText
  ]
    .filter((x) => x != null && String(x).trim() !== '')
    .join(' ')
    .toLowerCase();

  // Tamil Nadu Form I/1 is the Register of Workmen under the Conferment of
  // Permanent Status Act. Do not misclassify it as CLRA Form XIII merely
  // because both forms contain "Register of Workmen".
  if (
    /tamil\s*nadu|tamilnadu/i.test(parts) &&
    (/form[\s._-]*(?:i|1)(?:[\s._-]|$)/i.test(parts) ||
      /conferment\s+of\s+permanent\s+status/i.test(parts))
  ) {
    return false;
  }
  // Rajasthan Shops Forms 11/12/14/15 — never CLRA Form XIII (even if sheet text is wrong).
  if (
    /form[\s._-]*1[1245][\s._-]*rj|\bform_1[1245]_rj\b|\b1[1245]_rj\b/.test(identityBlob) ||
    /record\s+of\s+(?:the\s+)?hours\s+of\s+work/.test(identityBlob)
  ) {
    return false;
  }
  if (/form[\s._-]*xxiii(?![a-z])/i.test(parts)) return false;
  if (/register\s+of\s+overtime|overtime\s+register/i.test(parts)) return false;
  if (/register[\s._-]*of[\s._-]*wages[\s._-]*cum[\s._-]*muster/i.test(parts)) return false;
  if (/form[\s._-]*xviii(?![a-z])/i.test(parts)) return false;
  if (matchesFormXIVHint(parts)) return false;
  if (isFormXIVEmploymentCardContext(formHeader, rowItem, fileName, sheetText, tableHeaders)) {
    return false;
  }
  if (matchesFormXIIIHint(parts)) return true;
  if (blobIndicatesRegisterOfWorkmen(parts, tableHeaders) && /contract\s+labou?r/i.test(parts)) {
    return true;
  }
  return blobIndicatesRegisterOfWorkmen(parts, tableHeaders);
}

/**
 * AP/TN Form X (leave / fines) only — not Rajasthan CLRA Form X_RJ Employment Card.
 * Form_X_RJ must not match, or View File / sheet-repick swap in Register of Leave.
 */
export function matchesFormXHint(blob) {
  const text = String(blob || '').toLowerCase();
  if (MULTI_X_FORM_RE.test(text)) return false;
  // Form X_RJ / Form_X_RJ — Rajasthan Employment Card [Rule 75], not AP Shops Form X.
  if (/form[\s._-]*x[\s._-]*rj\b/i.test(text) || /\bx_rj\b/i.test(text)) return false;
  if (
    /rajasthan/i.test(text) &&
    /employment\s+card/i.test(text) &&
    /form[\s._-]*x\b/i.test(text) &&
    !/form[\s._-]*xiv/i.test(text)
  ) {
    return false;
  }
  if (/form[\s._-]*x(?:[\s._\-]|$)/i.test(text)) return true;
  if (/form___x\b|form_-_x\b/i.test(text)) return true;
  return false;
}

/** Form 21 / XXI — must not collide with Form X (substring "x" inside "xxi") or Form XIII. */
export function matchesFormXXIHint(blob) {
  const parts = String(blob || '').toLowerCase();
  if (FORM_XXI_COLLISION_ROMAN_RE.test(parts)) {
    return false;
  }
  return (
    /form[\s._-]*xxi(?![a-z])/i.test(parts) ||
    /form[\s._-]*21(?!\d)/i.test(parts) ||
    /(?:^|[\s._-])xxi(?=[\s._\W-]|$)/i.test(parts) ||
    /xxi[\s._-]*fines/i.test(parts)
  );
}

/** CLRA Form XXI — Register of Fines (Contract Labour), not AP Shops Form X or Form XX. */
export function sheetBlobIndicatesFormXXIRegisterOfFines(blob) {
  const text = String(blob || '').toLowerCase();
  if (matchesFormXXHint(text) || blobIndicatesRegisterOfDeductionsForDamage(text)) return false;
  if (matchesFormXIIIHint(text) || /register\s+of\s+workmen/i.test(text)) return false;
  if (matchesFormXXIHint(text)) return true;
  if (!/register\s+of\s+fines/i.test(text)) return false;
  return (
    /contract\s+labou?r/i.test(text) ||
    /principal\s+employer/i.test(text) ||
    /name\s+and\s+address\s+of\s+(the\s+)?contractor/i.test(text) ||
    (CLRA_WORKMEN_NAME_HEADER_RE.test(text) &&
      (/act\/omission|father'?s\/husband'?s|nature\s+of\s+employment/i.test(text) ||
        /vide\s+rule\s+78/i.test(text)))
  );
}

/** AP Shops Form X — Register of Leave and Social Security Benefits (not Register of Fines). */
export function sheetBlobIndicatesFormXLeaveRegister(blob) {
  const text = String(blob || '').toLowerCase();
  if (/register\s+of\s+fines/i.test(text)) return false;
  return (
    /register\s+of\s+leave/i.test(text) ||
    /leave\s+and\s+social\s+security/i.test(text) ||
    /leave\s+with\s+wages/i.test(text) ||
    (/earned\s+leave/i.test(text) &&
      (/medical\s+leave/i.test(text) || /leave\s+at\s+the\s+beginning/i.test(text)))
  );
}

/**
 * PW / Tamil Nadu Form I Register of Fines — must not be treated as AP Shops Form X.
 * (Form_I_-_TamilNadu.xlsx often cites Payment/Minimum Wages Act, which previously
 * tripped the soft Form X heuristic.)
 */
export function blobIndicatesFormIRegisterOfFinesNotFormX(blob) {
  const text = String(blob || '').toLowerCase();
  if (!text) return false;
  // Explicit multi-X Romans / Form X / XXI always win.
  if (MULTI_X_FORM_RE.test(text)) return false;
  if (matchesFormXXIHint(text)) return false;
  if (matchesFormXHint(text)) return false;

  const formITamilFile =
    /form[_\s.-]*i[_\s.-]*tamil/i.test(text) ||
    /form_i_-_tamil/i.test(text) ||
    (/form[_\s.-]*1[_\s.-]*tamil/i.test(text) && !/form[_\s.-]*1[0-9]/i.test(text));
  const formILetter =
    /\bform\s*[-–]?\s*i\b/i.test(text) ||
    /(?:^|[^a-z0-9])form[_\s.-]*i(?:[_\s.-]|$)/i.test(text) ||
    /pw\s+form\s*i\b/i.test(text);
  const formIDigit =
    /(?:^|[^a-z0-9])form[_\s.-]*1(?:[_\s.-]|$)/i.test(text) &&
    !/form[_\s.-]*1[0-9]/i.test(text);

  if (!(formITamilFile || formILetter || formIDigit)) return false;
  return (
    formITamilFile ||
    /register\s+of\s+fines/i.test(text) ||
    /tamil\s*nadu|tamilnadu/i.test(text)
  );
}

/** AP Shops & Establishment Form X — Register of Fines. */
export function sheetBlobIndicatesFormXAPRegisterOfFines(blob) {
  const text = String(blob || '').toLowerCase();
  if (sheetBlobIndicatesFormXXIRegisterOfFines(text)) return false;
  if (matchesFormXXIHint(text)) return false;
  if (blobIndicatesFormIRegisterOfFinesNotFormX(text)) return false;
  if (matchesFormXHint(text) && /register\s+of\s+fines/i.test(text)) return true;
  // Do not use bare "payment/minimum wages" — that also matches PW Form I Register of Fines.
  return (
    /register\s+of\s+fines/i.test(text) &&
    (/nature\s*&\s*date\s+of\s+offence|show\s+cause/i.test(text) ||
      /name\s+of\s+the\s+worker/i.test(text) ||
      /shops\s*(?:&|and)\s*establishment/i.test(text))
  );
}

function matchesForm27CHintFromParts(parts) {
  const text = String(parts || '').toLowerCase();
  return (
    /\bform[\s._-]*no\.?\s*27[\s._-]*c\b/.test(text) ||
    /\bform[\s._-]*27[\s._-]*c\b/.test(text) ||
    (/\b27[\s._-]*c\b/.test(text) && (/health\s+register/.test(text) || /\bform\b/.test(text)))
  );
}

export function isFormXXIAPRegisterOfFinesContext(
  formHeader,
  rowItem,
  fileName,
  sheetText = '',
  tableHeaders = null
) {
  const parts = [
    rowItem?.formName,
    rowItem?.FormName,
    rowItem?.description,
    rowItem?.Description,
    rowItem?.act,
    rowItem?.Act,
    fileName,
    formHeader?.title,
    formHeader?.subtitle,
    formHeader?.reference,
    sheetText,
    Array.isArray(tableHeaders) ? tableHeaders.join(' ') : ''
  ]
    .filter((x) => x != null && String(x).trim() !== '')
    .join(' ')
    .toLowerCase();

  if (matchesForm27CHintFromParts(parts)) return false;
  if (
    Array.isArray(tableHeaders) &&
    tableHeaders.length > 0 &&
    /name\s+of\s+worker/.test(tableHeaders.map((h) => String(h || '').toLowerCase()).join('\n')) &&
    /department/.test(tableHeaders.map((h) => String(h || '').toLowerCase()).join('\n')) &&
    /health|medical\s+examin|fit\s*\/\s*unfit/.test(tableHeaders.map((h) => String(h || '').toLowerCase()).join('\n'))
  ) {
    return false;
  }

  if (isFormXIIIRegisterOfWorkmenContext(formHeader, rowItem, fileName, sheetText, tableHeaders)) {
    return false;
  }
  if (isFormXXAPRegisterOfDeductionsContext(formHeader, rowItem, fileName, sheetText, tableHeaders)) {
    return false;
  }
  if (blobIndicatesRegisterOfWorkmen(parts, tableHeaders)) return false;
  if (sheetBlobIndicatesFormXXIRegisterOfFines(parts)) return true;
  if (matchesFormXXIHint(parts) && /register\s+of\s+fines/i.test(parts)) return true;

  if (Array.isArray(tableHeaders) && tableHeaders.length > 0) {
    const hdr = tableHeaders.map((h) => String(h || '').toLowerCase()).join('\n');
    if (blobIndicatesRegisterOfDeductionsForDamage('', tableHeaders)) return false;
    if (REGISTER_OF_WORKMEN_SURNAME_HEADER_RE.test(hdr)) return false;
    if (
      CLRA_WORKMEN_NAME_HEADER_RE.test(hdr) &&
      (/act\/omission|father'?s\/husband'?s|nature\s+of\s+employment/i.test(hdr) ||
        /principal\s+employer|contractor/i.test(hdr))
    ) {
      return true;
    }
  }
  return false;
}

function rowMetadataWantsFormXAP(rowItem, fileName) {
  const rowBlob = [
    rowItem?.formName,
    rowItem?.FormName,
    rowItem?.description,
    rowItem?.Description,
    rowItem?.formFileName,
    rowItem?.FormFileName,
    fileName
  ]
    .filter((x) => x != null && String(x).trim() !== '')
    .join(' ')
    .toLowerCase();
  if (matchesFormXXIHint(rowBlob)) return false;
  if (/contract\s+labou?r/i.test(rowBlob) && /register\s+of\s+fines/i.test(rowBlob)) return false;
  // matchesFormXHint already excludes Form X_RJ; keep explicit guard for employment-card rows.
  if (/form[\s._-]*x[\s._-]*rj\b/i.test(rowBlob) || /\bx_rj\b/i.test(rowBlob)) return false;
  if (/employment\s+card/i.test(rowBlob) && /form[\s._-]*x\b/i.test(rowBlob) && !/form[\s._-]*xiv/i.test(rowBlob)) {
    return false;
  }
  return matchesFormXHint(rowBlob);
}

export function isFormXAPRegisterOfFinesContext(
  formHeader,
  rowItem,
  fileName,
  sheetText = '',
  tableHeaders = null
) {
  const parts = [
    rowItem?.formName,
    rowItem?.FormName,
    rowItem?.description,
    rowItem?.Description,
    rowItem?.act,
    rowItem?.Act,
    rowItem?.state,
    rowItem?.State,
    rowItem?.formFileName,
    rowItem?.FormFileName,
    fileName,
    formHeader?.title,
    formHeader?.subtitle,
    formHeader?.reference,
    sheetText,
    Array.isArray(tableHeaders) ? tableHeaders.join(' ') : ''
  ]
    .filter((x) => x != null && String(x).trim() !== '')
    .join(' ')
    .toLowerCase();

  if (blobIndicatesFormIRegisterOfFinesNotFormX(parts)) return false;
  if (sheetBlobIndicatesFormXXIRegisterOfFines(parts)) return false;
  if (isFormXXIAPRegisterOfFinesContext(formHeader, rowItem, fileName, sheetText, tableHeaders)) {
    return false;
  }

  if (sheetBlobIndicatesFormXAPRegisterOfFines(parts)) return true;

  if (rowMetadataWantsFormXAP(rowItem, fileName) && /register\s+of\s+fines/i.test(parts)) {
    return true;
  }

  if (!rowMetadataWantsFormXAP(rowItem, fileName)) return false;

  if (Array.isArray(tableHeaders) && tableHeaders.length > 0) {
    const hdr = tableHeaders.map((h) => String(h || '').toLowerCase()).join('\n');
    if (
      CLRA_WORKMEN_NAME_HEADER_RE.test(hdr) ||
      REGISTER_OF_WORKMEN_SURNAME_HEADER_RE.test(hdr) ||
      /contractor|principal\s+employer|father'?s\/husband'?s/i.test(hdr)
    ) {
      return false;
    }
    if (
      /name\s+of\s+the\s+worker|name\s+of\s+the\s+employee/.test(hdr) &&
      /nature\s*&\s*date\s+of\s+offence|fine\s+imposed|show\s+cause/.test(hdr)
    ) {
      return true;
    }
  }

  if (sheetText && sheetBlobIndicatesFormXXIRegisterOfFines(sheetText)) return false;
  if (formHeader?.title && sheetBlobIndicatesFormXXIRegisterOfFines(formHeader.title)) return false;

  return false;
}

function buildSheetTextBlob(workbook, sheetName) {
  const ws = workbook?.Sheets?.[sheetName];
  if (!ws) return '';
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  return rows
    .slice(0, 40)
    .map((row) => (Array.isArray(row) ? row : []).map((c) => String(c || '')).join(' '))
    .join(' ')
    .toLowerCase();
}

function scoreFormXAPSheet(sheetName, sheetText, hintsBlob) {
  const sheetLower = String(sheetName || '').toLowerCase();
  const sheetBlob = `${sheetLower} ${sheetText}`;
  let score = 0;

  if (sheetBlobIndicatesFormXXIRegisterOfFines(sheetBlob)) score -= 300;
  if (matchesFormXXIHint(sheetBlob)) score -= 250;
  if (/contract\s+labou?r/i.test(sheetBlob) && /register\s+of\s+fines/i.test(sheetBlob)) score -= 220;
  if (sheetBlobIndicatesFormXLeaveRegister(sheetBlob)) score -= 280;

  if (sheetBlobIndicatesFormXAPRegisterOfFines(sheetBlob)) score += 180;
  if (matchesFormXHint(sheetLower) && !matchesFormXXIHint(sheetLower)) score += 120;
  if (/register\s+of\s+fines/i.test(sheetBlob) && !sheetBlobIndicatesFormXXIRegisterOfFines(sheetBlob)) {
    score += 60;
  }
  if (/name\s+of\s+the\s+worker|name\s+of\s+the\s+employee/i.test(sheetText)) score += 35;
  if (/nature\s*&\s*date\s+of\s+offence|show\s+cause/i.test(sheetText)) score += 45;
  if (/fine\s+imposed/i.test(sheetText) && !/act\/omission/i.test(sheetText)) score += 20;
  if (/rate\s+of\s+wages/i.test(sheetText)) score += 20;
  if (/appointment|notice\s+of\s+change|establishment\s+already\s+registered/i.test(sheetBlob)) {
    score -= 150;
  }
  if (/advance\s+of\s+wages|register\s+of\s+advances/i.test(sheetBlob)) score -= 120;
  if (matchesFormXHint(hintsBlob) && matchesFormXHint(sheetLower)) score += 30;
  if (/andhra|a\.p\./i.test(hintsBlob) && /andhra|a\.p\./i.test(sheetLower)) score += 10;

  return score;
}

export function resolveFormXWorkbookSheetName(workbook, hints = {}) {
  const names = Array.isArray(workbook?.SheetNames) ? workbook.SheetNames : [];
  if (names.length <= 1) return names[0] || null;

  const blob = [
    hints.fileName,
    hints.formFileName,
    hints.formName,
    hints.item?.formName,
    hints.item?.FormName,
    hints.formHeaderTitle,
    hints.formHeader?.title
  ]
    .filter(Boolean)
    .join(' ')
    .toLowerCase();

  let best = null;
  let bestScore = -Infinity;
  for (const name of names) {
    const sheetText = buildSheetTextBlob(workbook, name);
    const score = scoreFormXAPSheet(name, sheetText, blob);
    if (score > bestScore) {
      bestScore = score;
      best = name;
    }
  }
  if (bestScore > 0 && best) return best;

  const apFinesSheet = names.find((n) => {
    const sheetText = buildSheetTextBlob(workbook, n);
    return sheetBlobIndicatesFormXAPRegisterOfFines(`${n} ${sheetText}`);
  });
  return apFinesSheet || null;
}

export function repickFormXWorkbookSheetIfNeeded(workbook, hints, currentSheetName, parsedFormHeader) {
  const names = Array.isArray(workbook?.SheetNames) ? workbook.SheetNames : [];
  if (names.length <= 1) return null;

  const wantsFormX =
    isFormXAPRegisterOfFinesContext(
      parsedFormHeader,
      hints.item,
      hints.fileName || hints.formFileName
    ) ||
    (rowMetadataWantsFormXAP(hints.item, hints.fileName || hints.formFileName) &&
      /register\s+of\s+fines/i.test(
        [
          hints.item?.formName,
          hints.item?.FormName,
          hints.item?.description,
          hints.fileName,
          hints.formFileName
        ]
          .filter(Boolean)
          .join(' ')
          .toLowerCase()
      ));
  if (!wantsFormX) return null;

  const currentText = currentSheetName ? buildSheetTextBlob(workbook, currentSheetName) : '';
  const currentBlob = `${parsedFormHeader?.title || ''} ${currentSheetName || ''} ${currentText}`.toLowerCase();
  if (sheetBlobIndicatesFormXLeaveRegister(currentBlob)) {
    return resolveFormXWorkbookSheetName(workbook, {
      ...hints,
      formHeader: parsedFormHeader
    });
  }
  if (
    sheetBlobIndicatesFormXAPRegisterOfFines(currentBlob) &&
    !sheetBlobIndicatesFormXXIRegisterOfFines(currentBlob)
  ) {
    return null;
  }

  return resolveFormXWorkbookSheetName(workbook, {
    ...hints,
    formHeader: parsedFormHeader
  });
}

/** When checklist row is Form X but workbook only has Form XXI fines sheet. */
export function validateFormXAPTemplateAgainstRow(workbook, hints = {}) {
  if (!rowMetadataWantsFormXAP(hints.item, hints.fileName || hints.formFileName)) return null;
  if (!workbook?.SheetNames?.length) return null;

  const names = workbook.SheetNames;
  const hasFormXAPSheet = names.some((n) => {
    const sheetText = buildSheetTextBlob(workbook, n);
    const blob = `${n} ${sheetText}`;
    return sheetBlobIndicatesFormXAPRegisterOfFines(blob) && !sheetBlobIndicatesFormXXIRegisterOfFines(blob);
  });
  if (hasFormXAPSheet) return null;

  const hasFormXXISheet = names.some((n) => {
    const sheetText = buildSheetTextBlob(workbook, n);
    return sheetBlobIndicatesFormXXIRegisterOfFines(`${n} ${sheetText}`);
  });
  if (hasFormXXISheet) return FORM_X_AP_TEMPLATE_MISMATCH_MESSAGE;
  return null;
}

function getSheetMergedCellText(rows, merges, r, c) {
  const direct = String((rows[r] || [])[c] ?? '').trim();
  if (direct) return direct;
  for (const m of merges || []) {
    if (!m?.s || !m?.e) continue;
    if (r >= m.s.r && r <= m.e.r && c >= m.s.c && c <= m.e.c) {
      return String((rows[m.s.r] || [])[m.s.c] ?? '').trim();
    }
  }
  return '';
}

function normStatutoryHeaderLabel(h) {
  return String(h || '')
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

/** True when merged Excel headers produced duplicate / one mega-label for many columns. */
export function statutoryTableHeadersLookMergedDuplicate(headers) {
  const list = (Array.isArray(headers) ? headers : [])
    .map((h) => String(h || '').replace(/\s+/g, ' ').trim())
    .filter(Boolean);
  if (list.length < 2) return false;
  const norms = list.map(normStatutoryHeaderLabel);
  const unique = new Set(norms);
  if (unique.size <= 1) return true;
  if (unique.size < Math.ceil(list.length * 0.6)) return true;
  const longest = list.reduce((a, b) => (String(a).length >= String(b).length ? a : b), '');
  if (String(longest).length > 48 && list.filter((h) => h === longest).length >= 2) return true;
  if (norms.filter((n) => n === 'date of recovery').length >= 2) return true;
  return false;
}

function formXXAPNormalizeHeaderCell(text) {
  return String(text || '')
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function formXXAPRowCellText(rows, merges, rowIdx, colIdx) {
  const direct = formXXAPNormalizeHeaderCell((rows[rowIdx] || [])[colIdx]);
  if (direct) return direct;
  for (const m of merges || []) {
    if (!m?.s || !m?.e) continue;
    if (rowIdx >= m.s.r && rowIdx <= m.e.r && colIdx >= m.s.c && colIdx <= m.e.c) {
      return formXXAPNormalizeHeaderCell((rows[m.s.r] || [])[m.s.c]);
    }
  }
  return '';
}

/** Expand short Date-of-recovery child labels (First / Last) to instalment leaf headers. */
export function formXXAPExpandRecoveryChildLabel(text) {
  const bare = formXXAPNormalizeHeaderCell(text).toLowerCase();
  if (bare === 'first') return 'First instalment';
  if (bare === 'last') return 'Last instalment';
  if (/^first\s+install?ment$/i.test(bare)) return 'First instalment';
  if (/^last\s+install?ment$/i.test(bare)) return 'Last instalment';
  return formXXAPNormalizeHeaderCell(text);
}

/**
 * Resolve leaf table headers for Form XX AP when "Date of recovery" spans instalment sub-columns.
 * `headerRowIndex` / `startCol` are 0-based indices into `rows`.
 */
export function resolveFormXXAPLeafHeaders(rows, merges, headerRowIndex, startCol, maxCols) {
  const parentRow = Math.max(0, Number(headerRowIndex) || 0);
  const childRow = parentRow + 1;
  const c0 = Math.max(0, Number(startCol) || 0);
  const span = Math.max(1, Number(maxCols) || 1);
  const labels = [];

  for (let c = c0; c < c0 + span; c += 1) {
    const parent = formXXAPRowCellText(rows, merges, parentRow, c);
    const child = formXXAPRowCellText(rows, merges, childRow, c);
    const parentNorm = parent.toLowerCase();

    if (child) {
      labels.push(formXXAPExpandRecoveryChildLabel(child));
      continue;
    }
    if (/^date\s+of\s+recovery$/i.test(parentNorm)) continue;
    if (parent) {
      labels.push(parent);
    } else if (labels.length > 0) {
      break;
    }
  }
  return labels;
}

/** Leaf headers with 1-based Excel column numbers for template write-back. */
function resolveFormXXAPTemplateLeafCols(rows, merges, headerRowIndex, startCol, maxCols) {
  const parentRow = Math.max(0, Number(headerRowIndex) || 0);
  const childRow = parentRow + 1;
  const c0 = Math.max(0, Number(startCol) || 0);
  const span = Math.max(1, Number(maxCols) || 1);
  const cols = [];

  for (let c = c0; c < c0 + span; c += 1) {
    const parent = formXXAPRowCellText(rows, merges, parentRow, c);
    const child = formXXAPRowCellText(rows, merges, childRow, c);
    const parentNorm = parent.toLowerCase();

    if (child) {
      cols.push({ col: c + 1, label: formXXAPExpandRecoveryChildLabel(child) });
      continue;
    }
    if (/^date\s+of\s+recovery$/i.test(parentNorm)) continue;
    if (parent) {
      cols.push({ col: c + 1, label: parent });
    } else if (cols.length > 0) {
      break;
    }
  }
  return cols;
}

function buildFormXXAPWorksheetHeaderMatrix(
  worksheet,
  headerRowIndex,
  startCol,
  colSpan,
  excelCellValueToString
) {
  const rows = [];
  for (let r = headerRowIndex; r <= headerRowIndex + 2; r += 1) {
    const row = [];
    for (let c = startCol; c < startCol + colSpan; c += 1) {
      row.push(formXXAPNormalizeHeaderCell(excelCellValueToString(worksheet.getCell(r, c)?.value)));
    }
    rows.push(row);
  }
  return rows;
}

function resolveFormXXAPHeaderBandRows(worksheet, columnHeaderRow, startCol, colSpan, excelCellValueToString) {
  for (let r = Math.max(1, columnHeaderRow); r <= columnHeaderRow + 2; r += 1) {
    let childHits = 0;
    for (let c = startCol; c < startCol + colSpan; c += 1) {
      const t = formXXAPNormalizeHeaderCell(
        excelCellValueToString(worksheet.getCell(r, c)?.value)
      ).toLowerCase();
      if (/no\.?\s*of\s+install?ments?|^first$|^last$|first\s+install?ment|last\s+install?ment/.test(t)) {
        childHits += 1;
      }
    }
    if (childHits >= 2) {
      return { parentRow: Math.max(1, r - 1), childRow: r };
    }
  }
  return { parentRow: columnHeaderRow, childRow: columnHeaderRow + 1 };
}

function resolveFormXXAPTemplateLeafColsFromWorksheet(
  worksheet,
  columnHeaderRow,
  startCol,
  colSpan,
  excelCellValueToString
) {
  const { parentRow } = resolveFormXXAPHeaderBandRows(
    worksheet,
    columnHeaderRow,
    startCol,
    colSpan,
    excelCellValueToString
  );
  const matrix = buildFormXXAPWorksheetHeaderMatrix(
    worksheet,
    parentRow,
    startCol,
    colSpan,
    excelCellValueToString
  );
  const merges = Array.isArray(worksheet?.model?.merges)
    ? worksheet.model.merges
        .map((range) => {
          const m = String(range || '').match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/i);
          if (!m) return null;
          const col = (letters) =>
            String(letters || '')
              .toUpperCase()
              .split('')
              .reduce((n, ch) => n * 26 + (ch.charCodeAt(0) - 64), 0) - 1;
          return {
            s: { r: parseInt(m[2], 10) - 1, c: col(m[1]) },
            e: { r: parseInt(m[4], 10) - 1, c: col(m[3]) },
          };
        })
        .filter(Boolean)
    : [];
  return resolveFormXXAPTemplateLeafCols(matrix, merges, 0, 0, colSpan);
}

function mapFormXXAPHeaderNormKey(header) {
  return String(header || '')
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9 ]/gi, '')
    .replace(/\s+/g, ' ');
}

/** UI may concatenate duplicate labels: "Father's/Husband's Name_Father's/Husband's Name". */
export function dedupeFormXXAPConcatenatedHeader(header) {
  const raw = String(header || '')
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  if (!raw) return '';
  if (raw.includes('_')) {
    const parts = raw
      .split('_')
      .map((p) => p.trim())
      .filter(Boolean);
    const seen = new Set();
    for (const part of parts) {
      const key = mapFormXXAPHeaderNormKey(part);
      if (key && !seen.has(key)) {
        seen.add(key);
        return part;
      }
    }
    return parts[0] || raw;
  }
  return raw;
}

export function isFormXXAPWorkmenNameHeader(header) {
  const bare = formXXAPHeaderBare(dedupeFormXXAPConcatenatedHeader(header));
  return (
    /\bname\s+of\s+(?:the\s+)?(?:workmen|wokmen|workman)\b|\bname\s+of\s+workman\b/.test(bare) ||
    (/\bname\s+of\b/.test(bare) && /workmen|workman|worker/.test(bare) && !/presence|explanation/.test(bare))
  );
}

export function isFormXXAPFatherHusbandHeader(header) {
  const bare = formXXAPHeaderBare(dedupeFormXXAPConcatenatedHeader(header));
  return /father|husband/.test(bare) && !/presence|explanation|showed\s+cause/.test(bare);
}

export function isFormXXAPNatureOfEmploymentHeader(header) {
  const bare = formXXAPHeaderBare(dedupeFormXXAPConcatenatedHeader(header));
  return /nature\s+of\s+employ/.test(bare) || (/designat/.test(bare) && /employ/.test(bare));
}

export function normalizeFormXXAPExportHeaders(headers) {
  return (Array.isArray(headers) ? headers : [])
    .map((h) => dedupeFormXXAPConcatenatedHeader(h))
    .filter(Boolean);
}

/** Prefer modal headers that carry Father / Designation when template layout strips or duplicates them. */
export function pickFormXXAPExportHeaders(uiHeaders, layoutHeaders) {
  const ui = normalizeFormXXAPExportHeaders(uiHeaders);
  const layout = normalizeFormXXAPExportHeaders(layoutHeaders);
  const uiHasIdentity =
    ui.some((h) => isFormXXAPFatherHusbandHeader(h)) &&
    ui.some((h) => isFormXXAPNatureOfEmploymentHeader(h));
  const layoutHasIdentity =
    layout.some((h) => isFormXXAPFatherHusbandHeader(h)) &&
    layout.some((h) => isFormXXAPNatureOfEmploymentHeader(h));
  if (ui.length >= 3 && uiHasIdentity) return ui;
  if (layout.length >= 3 && layoutHasIdentity && !uiHasIdentity) return layout;
  if (ui.length >= 3 && !statutoryTableHeadersLookMergedDuplicate(ui)) return ui;
  if (layout.length >= 3 && !statutoryTableHeadersLookMergedDuplicate(layout)) return layout;
  return ui.length >= layout.length ? ui : layout;
}

function formXXAPFindRowValueByHeaderKind(row, matcher) {
  if (!row || typeof row !== 'object') return '';
  for (const key of Object.keys(row)) {
    if (String(key || '').startsWith('__')) continue;
    if (!matcher(key)) continue;
    const value = row[key];
    if (value != null && String(value).trim() !== '') return value;
  }
  return '';
}

function formXXAPResolveTemplateColForHeader(label, templateLeafCols, usedCols, startCol, idx) {
  const clean = dedupeFormXXAPConcatenatedHeader(label);
  const key = mapFormXXAPHeaderNormKey(clean);
  const pick = (matcher) =>
    templateLeafCols.find(({ col, label: tplLabel }) => {
      if (usedCols.has(col)) return false;
      return matcher(tplLabel) || matcher(clean);
    })?.col;

  if (isFormXXAPFatherHusbandHeader(clean)) {
    const col = pick(isFormXXAPFatherHusbandHeader);
    if (col != null) return col;
  }
  if (isFormXXAPNatureOfEmploymentHeader(clean)) {
    const col = pick(isFormXXAPNatureOfEmploymentHeader);
    if (col != null) return col;
  }
  if (isFormXXAPWorkmenNameHeader(clean)) {
    const col = pick(isFormXXAPWorkmenNameHeader);
    if (col != null) return col;
  }

  const direct = templateLeafCols.find(({ col, label: tplLabel }) => {
    if (usedCols.has(col)) return false;
    return mapFormXXAPHeaderNormKey(tplLabel) === key;
  });
  if (direct) return direct.col;

  const fuzzy = templateLeafCols.find(({ col, label: tplLabel }) => {
    if (usedCols.has(col)) return false;
    const tplKey = mapFormXXAPHeaderNormKey(tplLabel);
    return tplKey.includes(key) || key.includes(tplKey);
  });
  return fuzzy?.col ?? startCol + idx;
}

function mapFormXXAPWritableCols(sourceHeaders, templateLeafCols, startCol) {
  if (!Array.isArray(templateLeafCols) || templateLeafCols.length === 0) {
    return normalizeFormXXAPExportHeaders(sourceHeaders).map((label, idx) => ({
      col: startCol + idx,
      label: String(label || ''),
    }));
  }
  if (!Array.isArray(sourceHeaders) || sourceHeaders.length === 0) {
    return templateLeafCols.map(({ col, label }) => ({ col, label: String(label || '') }));
  }

  const usedCols = new Set();
  const mapped = normalizeFormXXAPExportHeaders(sourceHeaders).map((label, idx) => {
    const col = formXXAPResolveTemplateColForHeader(label, templateLeafCols, usedCols, startCol, idx);
    usedCols.add(col);
    return { col, label: String(label || '') };
  });

  templateLeafCols.forEach(({ col, label }) => {
    if (usedCols.has(col)) return;
    mapped.push({ col, label: String(label || '') });
  });

  mapped.sort((a, b) => a.col - b.col);
  return mapped;
}

function rowValueForFormXXAPHeader(row, header, headers) {
  const label = dedupeFormXXAPConcatenatedHeader(header);
  const hdrs = (Array.isArray(headers) ? headers : []).map((h) => dedupeFormXXAPConcatenatedHeader(h));

  if (isFormXXAPFatherHusbandHeader(label)) {
    const hit = formXXAPFindRowValueByHeaderKind(row, isFormXXAPFatherHusbandHeader);
    if (hit !== '') return hit;
  }
  if (isFormXXAPNatureOfEmploymentHeader(label)) {
    const hit = formXXAPFindRowValueByHeaderKind(row, isFormXXAPNatureOfEmploymentHeader);
    if (hit !== '') return hit;
  }
  if (isFormXXAPWorkmenNameHeader(label)) {
    const hit = formXXAPFindRowValueByHeaderKind(row, isFormXXAPWorkmenNameHeader);
    if (hit !== '') return hit;
  }

  const values = exportFormXXAPRowValuesByHeaders(row, hdrs.length > 0 ? hdrs : [label]);
  const idx = hdrs.indexOf(label);
  if (idx >= 0 && values[idx] != null && String(values[idx]).trim() !== '') return values[idx];

  if (!row || typeof row !== 'object') return '';
  const direct = row[header];
  if (direct != null && String(direct).trim() !== '') return direct;
  if (row[label] != null && String(row[label]).trim() !== '') return row[label];

  const target = mapFormXXAPHeaderNormKey(label);
  const key = Object.keys(row).find((k) => {
    if (String(k).startsWith('__')) return false;
    const nk = mapFormXXAPHeaderNormKey(dedupeFormXXAPConcatenatedHeader(k));
    return nk === target || nk.includes(target) || target.includes(nk);
  });
  return key && row[key] != null && String(row[key]).trim() !== '' ? row[key] : '';
}

function mergeMasterCol(merges, r, c) {
  for (const m of merges || []) {
    if (!m?.s || !m?.e) continue;
    if (r >= m.s.r && r <= m.e.r && c >= m.s.c && c <= m.e.c) return m.s.c;
  }
  return c;
}

function readMergeAwareHeaderLabels(rows, merges, rowIdx, startCol, maxCols) {
  const labels = [];
  let c = startCol;
  while (c < maxCols) {
    const master = mergeMasterCol(merges, rowIdx, c);
    if (master !== c) {
      c += 1;
      continue;
    }
    const text = getSheetMergedCellText(rows, merges, rowIdx, c).replace(/\s+/g, ' ').trim();
    if (!text) {
      if (labels.length > 0) break;
      c += 1;
      continue;
    }
    let endCol = c;
    for (const m of merges || []) {
      if (m?.s?.r === rowIdx && m?.s?.c === c) {
        endCol = m.e.c;
        break;
      }
    }
    labels.push(text);
    c = endCol + 1;
  }
  return labels;
}

function scoreFormXXIClraHeaderLabelRow(rows, merges, rowIdx, startCol, maxCols) {
  const labels = readMergeAwareHeaderLabels(rows, merges, rowIdx, startCol, maxCols);
  if (labels.length < 3) return { score: -1, labels: [] };
  if (statutoryTableHeadersLookMergedDuplicate(labels)) return { score: -1, labels: [] };
  let score = labels.length * 12;
  const joined = labels.map(normStatutoryHeaderLabel).join(' ');
  if (/serial|sl no|sr no|s no/.test(joined)) score += 35;
  if (/name\s+of\s+workman|name\s+of\s+workmen/.test(joined)) score += 35;
  if (/father|husband/.test(joined)) score += 25;
  if (/designat|nature\s+of\s+employ/.test(joined)) score += 25;
  if (/act\/omission|fine\s+imposed|date\s+of\s+offence|show\s+cause/.test(joined)) score += 20;
  for (const lab of labels) {
    if (String(lab).length > 55) score -= 50;
  }
  return { score, labels };
}

function resolveFormXXIClraHeaderBand(rows, merges, anchorRow, startCol, maxCols) {
  let best = { score: -1, labels: [], rowIdx: anchorRow };
  for (let r = Math.max(0, anchorRow - 2); r <= anchorRow + 4; r += 1) {
    const hit = scoreFormXXIClraHeaderLabelRow(rows, merges, r, startCol, maxCols);
    if (hit.score > best.score) best = { ...hit, rowIdx: r };
  }
  return best;
}

/** Prefer UI/modal headers when template layout headers are merged duplicates. */
export function pickFormXXIAPExportHeaders(uiHeaders, layoutHeaders) {
  const ui = (Array.isArray(uiHeaders) ? uiHeaders : [])
    .map((h) => String(h || '').replace(/\s+/g, ' ').trim())
    .filter(Boolean);
  const layout = (Array.isArray(layoutHeaders) ? layoutHeaders : [])
    .map((h) => String(h || '').replace(/\s+/g, ' ').trim())
    .filter(Boolean);
  if (ui.length >= 3 && !statutoryTableHeadersLookMergedDuplicate(ui)) return ui;
  if (layout.length >= 3 && !statutoryTableHeadersLookMergedDuplicate(layout)) return layout;
  return ui.length >= layout.length ? ui : layout;
}

export function exportFormXXAPRowValuesByHeaders(row, headers) {
  const list = (Array.isArray(headers) ? headers : []).map((h) => dedupeFormXXAPConcatenatedHeader(h));
  const normalize = (h) => mapFormXXAPHeaderNormKey(h);
  return list.map((hdr, idx) => {
    if (Array.isArray(row)) {
      return row[idx] != null && String(row[idx]).trim() !== '' ? row[idx] : '';
    }
    if (!row || typeof row !== 'object') return '';
    if (isFormXXAPFatherHusbandHeader(hdr)) {
      const hit = formXXAPFindRowValueByHeaderKind(row, isFormXXAPFatherHusbandHeader);
      if (hit !== '') return hit;
    }
    if (isFormXXAPNatureOfEmploymentHeader(hdr)) {
      const hit = formXXAPFindRowValueByHeaderKind(row, isFormXXAPNatureOfEmploymentHeader);
      if (hit !== '') return hit;
    }
    if (isFormXXAPWorkmenNameHeader(hdr)) {
      const hit = formXXAPFindRowValueByHeaderKind(row, isFormXXAPWorkmenNameHeader);
      if (hit !== '') return hit;
    }
    const direct = row[hdr];
    if (direct != null && String(direct).trim() !== '') return direct;
    const target = normalize(hdr);
    if (!target) return '';
    const key = Object.keys(row).find((k) => {
      if (String(k).startsWith('__')) return false;
      const nk = normalize(dedupeFormXXAPConcatenatedHeader(k));
      return nk === target || nk.includes(target) || target.includes(nk);
    });
    if (key && row[key] != null && String(row[key]).trim() !== '') return row[key];
    return '';
  });
}

function unmergeExcelJSRowsInRange(worksheet, startRow, endRow, colFrom, colTo) {
  if (!worksheet) return;
  for (let pass = 0; pass < 12; pass += 1) {
    let removed = 0;
    const merges = Array.isArray(worksheet?.model?.merges) ? [...worksheet.model.merges] : [];
    merges.forEach((range) => {
      const parts = String(range || '').split(':');
      if (parts.length !== 2) return;
      const start = parts[0].match(/^([A-Z]+)(\d+)$/i);
      const end = parts[1].match(/^([A-Z]+)(\d+)$/i);
      if (!start || !end) return;
      const r1 = parseInt(start[2], 10);
      const r2 = parseInt(end[2], 10);
      const c1 = XLSX.utils.decode_col(start[1].toUpperCase()) + 1;
      const c2 = XLSX.utils.decode_col(end[1].toUpperCase()) + 1;
      const rowOverlap = r2 >= startRow && r1 <= endRow;
      const colOverlap = c2 >= colFrom && c1 <= colTo;
      if (!rowOverlap || !colOverlap) return;
      try {
        worksheet.unMergeCells(range);
        removed += 1;
      } catch (_) {
        // ignore invalid merge ranges
      }
    });
    for (let r = startRow; r <= endRow; r += 1) {
      for (let c = colFrom; c <= colTo; c += 1) {
        const cell = worksheet.getCell(r, c);
        if (!cell?.isMerged) continue;
        try {
          const master = cell.master || cell;
          const addr = master.address || master.$col$row;
          if (addr) {
            worksheet.unMergeCells(String(addr));
            removed += 1;
          }
        } catch (_) {
          /* ignore */
        }
      }
    }
    if (!removed) break;
  }
}

/** Plain text from ExcelJS cell values (richText / formula results included). */
function formXXAPExcelCellText(value) {
  if (value == null || value === '') return '';
  if (typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') {
    return String(value).trim();
  }
  if (value instanceof Date) {
    try {
      return value.toLocaleDateString('en-IN');
    } catch (_) {
      return value.toISOString().slice(0, 10);
    }
  }
  if (typeof value === 'object') {
    if (Array.isArray(value.richText)) {
      return value.richText
        .map((p) => p?.text || '')
        .join('')
        .trim();
    }
    if (value.text != null) return String(value.text).trim();
    if (value.result != null) return formXXAPExcelCellText(value.result);
    if (value.hyperlink != null && value.text != null) return String(value.text).trim();
    if (value.v != null) return formXXAPExcelCellText(value.v);
  }
  return String(value).trim();
}

function formXXAPRowLooksLikeTableHeaderBand(worksheet, row, startCol, colTo) {
  let joined = '';
  for (let c = startCol; c <= colTo; c += 1) {
    joined += ` ${formXXAPExcelCellText(worksheet.getCell(row, c)?.value).replace(/\s+/g, ' ')}`;
  }
  const lower = joined.toLowerCase();
  const first = formXXAPExcelCellText(worksheet.getCell(row, startCol)?.value);
  return (
    (/^s\.?\s*no/i.test(first) && /name\s+of\s+workmen|name\s+of\s+workman/i.test(lower)) ||
    /no\.?\s*of\s+install?ments?|first\s+install?ment|last\s+install?ment|particulars\s+of\s+damage|date\s+of\s+recovery|father|husband|nature\s+of\s+employ|date\s+of\s+damage|whether\s+workman|amount\s+of\s+deduction|remarks?/i.test(
      lower
    )
  );
}

function locateFormXXAPDataStartRowFromWorksheet(worksheet, headerBandEnd, startCol, colTo) {
  let r = Math.max(1, Number(headerBandEnd) || 1) + 1;
  while (r <= headerBandEnd + 3 && formXXAPRowLooksLikeTableHeaderBand(worksheet, r, startCol, colTo)) {
    r += 1;
  }
  return r;
}

/** Match 2nd-image model: S.No/NIL centered; names & text left; vertical middle. */
function applyFormXXAPDataRowAlignment(
  worksheet,
  { dataStartRow, dataRowCount, colFrom, colTo } = {}
) {
  if (!worksheet || !dataRowCount || dataRowCount < 1) return;
  const r0 = Math.max(1, Number(dataStartRow) || 1);
  const r1 = r0 + Math.max(1, Number(dataRowCount) || 1) - 1;
  const c0 = Math.max(1, Number(colFrom) || 1);
  const c1 = Math.max(c0, Number(colTo) || c0);
  for (let r = r0; r <= r1; r += 1) {
    for (let c = c0; c <= c1; c += 1) {
      const cell = worksheet.getCell(r, c);
      const text = formXXAPExcelCellText(cell.value);
      const isNil = /^n+i+l+\.?$/i.test(text);
      const isSerial =
        c === c0 && (text === '' || /^-?\d+(\.\d+)?$/.test(text));
      const horizontal = isNil || isSerial ? 'center' : 'left';
      const nextAlign = {
        ...(cell.alignment || {}),
        horizontal,
        vertical: 'middle',
        wrapText: !!(cell.alignment && cell.alignment.wrapText),
      };
      cell.alignment = nextAlign;
      try {
        const prev = cell.style && typeof cell.style === 'object' ? { ...cell.style } : {};
        cell.style = { ...prev, alignment: { ...(prev.alignment || {}), ...nextAlign } };
      } catch (_) {
        /* alignment property above is enough */
      }
    }
  }
}

function findFormXXAPTemplateBodyBorderRow(worksheet, dataStartRow, colFrom, colTo, preferRow = 0) {
  if (preferRow >= dataStartRow) {
    let hits = 0;
    for (let c = colFrom; c <= colTo; c += 1) {
      if (excelJSCellHasBorder(worksheet.getCell(preferRow, c))) hits += 1;
    }
    if (hits >= Math.max(3, colTo - colFrom)) return preferRow;
  }
  for (let r = dataStartRow; r < dataStartRow + 40; r += 1) {
    let boxHits = 0;
    let borderHits = 0;
    for (let c = colFrom; c <= colTo; c += 1) {
      const cell = worksheet.getCell(r, c);
      if (excelJSCellHasFullBoxBorder(cell)) boxHits += 1;
      if (excelJSCellHasBorder(cell)) borderHits += 1;
    }
    const need = colTo - colFrom + 1;
    if (boxHits >= need || borderHits >= need) return r;
  }
  return Math.max(dataStartRow, preferRow || dataStartRow);
}

function countFormXXAPTemplateBodyRows(worksheet, dataStartRow, colFrom, colTo) {
  return countExcelJSTemplateBodyRows(worksheet, dataStartRow, colFrom, colTo, 120);
}

function resolveFormXXAPTableColMax(worksheet, columnHeaderRow, startCol, templateLeafCols, writableCols) {
  const remarksCol = findExcelJSRemarksOrLastHeaderCol(worksheet, {
    headerRow: columnHeaderRow,
    startCol,
    scanCols: 24,
  });
  const fromLeaf =
    Array.isArray(templateLeafCols) && templateLeafCols.length > 0
      ? templateLeafCols[templateLeafCols.length - 1].col
      : 0;
  const fromWritable =
    Array.isArray(writableCols) && writableCols.length > 0
      ? writableCols[writableCols.length - 1].col
      : 0;
  return Math.max(remarksCol || 0, fromLeaf, fromWritable, startCol + 12);
}

/** Form XX AP data rows: unmerge body merges then paint thin full box on every cell. */
function ensureFormXXAPTableDataBorders(
  worksheet,
  { dataStartRow, dataRowCount, colFrom, colTo, templateBodyRow = null } = {}
) {
  if (!worksheet || !dataRowCount || dataRowCount < 1) return;
  const r0 = Math.max(1, Number(dataStartRow) || 1);
  const r1 = r0 + Math.max(1, Number(dataRowCount) || 1) - 1;
  const c0 = Math.max(1, Number(colFrom) || 1);
  const c1 = Math.max(c0, Number(colTo) || c0);
  unmergeExcelJSRowsInRange(worksheet, r0, r1, c0, c1);
  for (let r = r0; r <= r1; r += 1) {
    for (let c = c0; c <= c1; c += 1) {
      const cell = worksheet.getCell(r, c);
      if (cell.value == null) cell.value = '';
      // Detach shared style refs so borders stick on writeBuffer.
      try {
        const prev = cell.style && typeof cell.style === 'object' ? { ...cell.style } : {};
        delete prev.border;
        cell.style = prev;
      } catch (_) {
        cell.style = {};
      }
    }
  }
  applyExcelJSFullBoxBordersToRange(worksheet, {
    rowFrom: r0,
    rowTo: r1,
    colFrom: c0,
    colTo: c1,
    borderStyle: 'thin',
  });
  applyFormXXAPDataRowAlignment(worksheet, {
    dataStartRow: r0,
    dataRowCount: r1 - r0 + 1,
    colFrom: c0,
    colTo: c1,
  });
  // Second paint — alignment/style writes can drop borders on some ExcelJS versions.
  applyExcelJSFullBoxBordersToRange(worksheet, {
    rowFrom: r0,
    rowTo: r1,
    colFrom: c0,
    colTo: c1,
    borderStyle: 'thin',
  });
  void templateBodyRow;
}

function worksheetLooksLikeFormXXAPDeductionsRegister(worksheet) {
  if (!worksheet || typeof worksheet.getCell !== 'function') return false;
  const parts = [];
  const maxR = Math.min(Number(worksheet.rowCount) || 30, 30);
  for (let r = 1; r <= maxR; r += 1) {
    for (let c = 1; c <= 16; c += 1) {
      const t = formXXAPExcelCellText(worksheet.getCell(r, c)?.value);
      if (t) parts.push(t);
    }
  }
  const blob = parts.join(' ').toLowerCase();
  if (/form\s*[-._ ]*xxi\b/.test(blob) && /register\s+of\s+fines/.test(blob)) return false;
  if (/form\s*[-._ ]*xx\b/.test(blob) && /register\s+of\s+deductions|deductions?\s+for\s+damage/.test(blob)) {
    return true;
  }
  if (
    /particulars\s+of\s+damage/.test(blob) &&
    /date\s+of\s+recovery|no\.?\s*of\s+install?ments?/.test(blob)
  ) {
    return true;
  }
  return false;
}

/** Last-pass on downloaded workbook — restore full grid borders on Form XX table (header + body). */
export function reapplyFormXXAPDownloadTableBordersFromWorksheet(worksheet) {
  if (!worksheet || typeof worksheet.getCell !== 'function') return;
  if (!worksheetLooksLikeFormXXAPDeductionsRegister(worksheet)) return;

  let headerBandStart = 0;
  let headerBandEnd = 0;
  for (let r = 1; r <= Math.min(Number(worksheet.rowCount) || 40, 40); r += 1) {
    if (formXXAPRowLooksLikeTableHeaderBand(worksheet, r, 1, 16)) {
      if (!headerBandStart) headerBandStart = r;
      headerBandEnd = r;
    } else if (headerBandEnd && r > headerBandEnd + 1) {
      break;
    }
  }
  if (!headerBandEnd) {
    // Fallback: S.No + Name of Workmen on the same row.
    for (let r = 1; r <= Math.min(Number(worksheet.rowCount) || 40, 40); r += 1) {
      let joined = '';
      for (let c = 1; c <= 16; c += 1) {
        joined += ` ${formXXAPExcelCellText(worksheet.getCell(r, c)?.value)}`;
      }
      if (/s\.?\s*no/i.test(joined) && /name\s+of\s+workmen|name\s+of\s+workman/i.test(joined)) {
        headerBandStart = r;
        headerBandEnd = r;
        if (formXXAPRowLooksLikeTableHeaderBand(worksheet, r + 1, 1, 16)) {
          headerBandEnd = r + 1;
        }
        break;
      }
    }
  }
  if (!headerBandEnd) return;
  if (!headerBandStart) headerBandStart = headerBandEnd;

  const startCol = 1;
  const tableColMax = resolveFormXXAPTableColMax(worksheet, headerBandEnd, startCol, [], []);
  const dataStartRow = locateFormXXAPDataStartRowFromWorksheet(
    worksheet,
    headerBandEnd,
    startCol,
    tableColMax
  );
  let lastPopulatedRow = dataStartRow - 1;
  for (let r = dataStartRow; r <= Math.min(Number(worksheet.rowCount) || 120, dataStartRow + 80); r += 1) {
    let hit = false;
    for (let c = startCol; c <= tableColMax; c += 1) {
      const cell = worksheet.getCell(r, c);
      const text = formXXAPExcelCellText(cell?.value);
      if (text) hit = true;
      if (excelJSCellHasBorder(cell)) hit = true;
    }
    if (hit) lastPopulatedRow = r;
    else if (lastPopulatedRow >= dataStartRow && r > lastPopulatedRow + 2) break;
  }
  if (lastPopulatedRow < dataStartRow) {
    lastPopulatedRow = dataStartRow + 1;
  }

  const bodyRowCount = Math.max(
    lastPopulatedRow - dataStartRow + 1,
    countFormXXAPTemplateBodyRows(worksheet, dataStartRow, startCol, tableColMax),
    2
  );
  const bodyEndRow = dataStartRow + bodyRowCount - 1;

  // Keep header merges (Date of recovery group); only force borders on the band.
  applyExcelJSFullBoxBordersToRange(worksheet, {
    rowFrom: headerBandStart,
    rowTo: headerBandEnd,
    colFrom: startCol,
    colTo: tableColMax,
    borderStyle: 'thin',
  });

  ensureFormXXAPTableDataBorders(worksheet, {
    dataStartRow,
    dataRowCount: bodyRowCount,
    colFrom: startCol,
    colTo: tableColMax,
  });

  // One more continuous paint so header+body share the same closed box (2nd image model).
  applyExcelJSFullBoxBordersToRange(worksheet, {
    rowFrom: headerBandStart,
    rowTo: bodyEndRow,
    colFrom: startCol,
    colTo: tableColMax,
    borderStyle: 'thin',
  });
}

function clearExcelJSWorksheetDataRange(worksheet, startRow, rowCount, colFrom, colTo) {
  if (!worksheet || rowCount < 1) return;
  unmergeExcelJSRowsInRange(worksheet, startRow, startRow + rowCount - 1, colFrom, colTo);
  for (let r = startRow; r < startRow + rowCount; r += 1) {
    for (let c = colFrom; c <= colTo; c += 1) {
      worksheet.getCell(r, c).value = null;
    }
  }
}

/** Locate the fines register table on the correct AP workbook tab (SheetJS). */
export function resolveFormXAPFinesTableLayout(workbook, hints = {}) {
  const names = Array.isArray(workbook?.SheetNames) ? workbook.SheetNames : [];
  if (!names.length) return null;

  const preferred =
    hints.preferredSheetName && workbook.Sheets?.[hints.preferredSheetName]
      ? hints.preferredSheetName
      : null;
  let sheetName = preferred;
  if (sheetName) {
    const prefText = buildSheetTextBlob(workbook, sheetName);
    const prefBlob = `${sheetName} ${prefText}`;
    if (sheetBlobIndicatesFormXXIRegisterOfFines(prefBlob)) {
      sheetName = null;
    }
    if (sheetBlobIndicatesFormXLeaveRegister(prefBlob)) {
      sheetName = null;
    }
  }
  if (!sheetName) {
    sheetName =
      resolveFormXWorkbookSheetName(workbook, hints) ||
      names.find((n) => {
        const sheetText = buildSheetTextBlob(workbook, n);
        const blob = `${n} ${sheetText}`;
        return (
          sheetBlobIndicatesFormXAPRegisterOfFines(blob) &&
          !sheetBlobIndicatesFormXLeaveRegister(blob)
        );
      }) ||
      null;
  }
  if (!sheetName) return null;

  const ws = workbook?.Sheets?.[sheetName];
  if (!ws) return null;

  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  const merges = ws['!merges'] || [];
  const maxScan = Math.min(rows.length, 90);
  let best = null;
  let bestScore = -1;

  for (let r = 0; r < maxScan; r += 1) {
    const cells = [];
    const maxCols = Math.max((rows[r] || []).length, (rows[r + 1] || []).length, 24);
    for (let c = 0; c < maxCols; c += 1) {
      const text = getSheetMergedCellText(rows, merges, r, c);
      if (text) cells.push({ c, text, lower: text.toLowerCase() });
    }
    if (cells.length < 2) continue;

    const nextRowCells = [];
    for (let c = 0; c < maxCols; c += 1) {
      const text = getSheetMergedCellText(rows, merges, r + 1, c);
      if (text) nextRowCells.push(text.toLowerCase());
    }
    const rowText = cells.map((x) => x.lower).join(' ');
    const combinedText = `${rowText} ${nextRowCells.join(' ')}`.trim();
    if (sheetBlobIndicatesFormXLeaveRegister(combinedText)) continue;
    if (sheetBlobIndicatesFormXXIRegisterOfFines(combinedText) && !/nature\s*&\s*date\s+of\s+offence/i.test(combinedText)) {
      continue;
    }
    let score = 0;
    if (/register\s+of\s+fines/.test(combinedText)) score += 50;
    if (/show\s+cause/.test(combinedText)) score += 35;
    if (/fine\s+imposed|nature\s*&\s*date\s+of\s+offence|act\s+or\s+omission/.test(combinedText)) score += 30;
    if (/name\s+of\s+the\s+worker|name\s+of\s+worker|name\s+of\s+the\s+employee/.test(combinedText)) score += 25;
    if (
      CLRA_WORKMEN_NAME_HEADER_RE.test(combinedText) ||
      REGISTER_OF_WORKMEN_SURNAME_HEADER_RE.test(combinedText) ||
      /father'?s\/husband'?s/i.test(combinedText)
    ) {
      score -= 40;
    }
    if (/rate\s+of\s+wages|total\s+wages|wages\s+payable/.test(combinedText)) score += 20;
    if (/s\.?\s*no|serial|sl\.?\s*no/.test(combinedText)) score += 15;
    score += cells.length * 2;

    if (score < 15) continue;

    let startCol = cells[0].c;
    for (const cell of cells) {
      if (/^s\.?\s*no\.?$|^sl\.?\s*no\.?$|serial/i.test(cell.lower)) {
        startCol = cell.c;
        break;
      }
    }

    const headers = [];
    for (let c = startCol; c < maxCols; c += 1) {
      let h = getSheetMergedCellText(rows, merges, r, c);
      if (!h) h = getSheetMergedCellText(rows, merges, r + 1, c);
      if (!h && r > 0) h = getSheetMergedCellText(rows, merges, r - 1, c);
      h = String(h || '').replace(/\s+/g, ' ').trim();
      if (!h) {
        if (headers.length > 0) break;
        continue;
      }
      if (/^\d+$/.test(h) && headers.length === 0) continue;
      headers.push(h);
    }
    if (headers.length < 3) continue;

    let dataStart = r + 1;
    const nextRow = rows[r + 1] || [];
    let seqHits = 0;
    for (let i = 0; i < Math.min(headers.length, 12); i += 1) {
      const t = String(nextRow[startCol + i] ?? '').trim();
      if (t === String(i + 1)) seqHits += 1;
    }
    if (seqHits >= 4) dataStart = r + 2;

    const candidate = {
      sheetName,
      headerRowIndex: r,
      dataStartIndex: dataStart,
      tableStartCol: startCol,
      headers
    };
    if (score > bestScore) {
      bestScore = score;
      best = candidate;
    }
  }

  return best;
}

function scoreFormXXIAPSheet(sheetName, sheetText, hintsBlob) {
  const sheetLower = String(sheetName || '').toLowerCase();
  const sheetBlob = `${sheetLower} ${sheetText}`;
  let score = 0;

  if (sheetBlobIndicatesFormXXIRegisterOfFines(sheetBlob)) score += 200;
  if (matchesFormXXIHint(sheetBlob)) score += 120;
  if (/register\s+of\s+fines/i.test(sheetBlob)) score += 80;
  if (/act\/omission|fine\s+imposed|show\s+cause|rate\s+of\s+wages/i.test(sheetText)) score += 60;
  if (
    CLRA_WORKMEN_NAME_HEADER_RE.test(sheetText) ||
    /father'?s\/husband'?s|nature\s+of\s+employ/.test(sheetText)
  ) {
    score += 40;
  }
  if (sheetBlobIndicatesFormXXAPRegisterOfDeductions(sheetBlob)) score -= 220;
  if (blobIndicatesRegisterOfDeductionsForDamage(sheetBlob)) score -= 200;
  if (sheetBlobIndicatesFormXAPRegisterOfFines(sheetBlob)) score -= 150;
  if (matchesFormXXHint(sheetBlob) && !blobIndicatesRegisterOfDeductionsForDamage(sheetBlob)) score -= 180;
  if (matchesFormXXIHint(hintsBlob) && matchesFormXXIHint(sheetLower)) score += 30;

  return score;
}

function extractClraSheetFormRomanToken(text) {
  const parts = String(text || '').toLowerCase();
  const formMatch = parts.match(
    /form[\s._-]*(xxiii|xxii|xxi|xviii|xvii|xvi|xv|xiv|xiii)(?![a-z])/i
  );
  return formMatch ? String(formMatch[1] || '').toLowerCase() : '';
}

function scoreFormXIVSheet(sheetName, sheetText, blob) {
  const sheetBlob = `${sheetName} ${sheetText}`.toLowerCase();
  let score = 0;
  const sheetToken = extractClraSheetFormRomanToken(sheetBlob);
  if (sheetToken === 'xiv') score += 120;
  if (sheetToken === 'xiii') score -= 220;
  if (matchesFormXIVHint(blob)) score += 40;
  if (blobIndicatesEmploymentCard(sheetBlob, null)) score += 90;
  // Form X_RJ Employment Card — prefer over Register of Leave sheets in multi-tab workbooks.
  if (/form[\s._-]*x[\s._-]*rj\b|\bx_rj\b/i.test(blob) && blobIndicatesEmploymentCard(sheetBlob, null)) {
    score += 80;
  }
  if (sheetBlobIndicatesFormXLeaveRegister(sheetBlob)) score -= 280;
  if (
    /register\s+of\s+workmen\s+employed\s+by\s+contractor/i.test(sheetBlob) &&
    !/employment\s+card/i.test(sheetBlob)
  ) {
    score -= 100;
  }
  if (/name\s+and\s+surname\s+of\s+workmen/i.test(sheetBlob) && /local\s+address/i.test(sheetBlob)) {
    score -= 80;
  }
  return score;
}

/** Pick CLRA Form XIV / Form X_RJ Employment Card worksheet (not Form XIII or leave register). */
export function resolveFormXIVWorkbookSheetName(workbook, hints = {}) {
  const names = Array.isArray(workbook?.SheetNames) ? workbook.SheetNames : [];
  if (!names.length) return null;

  const blob = [
    hints.fileName,
    hints.formFileName,
    hints.formName,
    hints.item?.formName,
    hints.item?.FormName,
    hints.item?.description,
    hints.item?.Description,
    hints.item?.state,
    hints.item?.State,
    hints.formHeaderTitle,
    hints.formHeader?.title,
    hints.formHeaderSubtitle,
    hints.formHeader?.subtitle
  ]
    .filter(Boolean)
    .join(' ')
    .toLowerCase();

  const wantsFormXRJ =
    /form[\s._-]*x[\s._-]*rj\b/i.test(blob) ||
    /\bx_rj\b/i.test(blob) ||
    (/rajasthan/i.test(blob) && /employment\s+card/i.test(blob) && /form[\s._-]*x\b/i.test(blob));

  if (names.length === 1) {
    const only = names[0];
    const sheetBlob = `${only} ${buildSheetTextBlob(workbook, only)}`;
    // Single leave-register sheet is not a usable Employment Card template.
    if (
      (wantsFormXRJ || matchesFormXIVHint(blob) || blobIndicatesEmploymentCard(blob, null)) &&
      sheetBlobIndicatesFormXLeaveRegister(sheetBlob) &&
      !blobIndicatesEmploymentCard(sheetBlob, null)
    ) {
      return null;
    }
    return only;
  }

  if (!matchesFormXIVHint(blob) && !blobIndicatesEmploymentCard(blob, null) && !wantsFormXRJ) {
    return null;
  }

  let best = null;
  let bestScore = -Infinity;
  for (const name of names) {
    const sheetText = buildSheetTextBlob(workbook, name);
    const score = scoreFormXIVSheet(name, sheetText, blob);
    if (score > bestScore) {
      bestScore = score;
      best = name;
    }
  }
  if (bestScore > 0 && best) return best;

  const xivSheet = names.find((n) => {
    const sheetText = buildSheetTextBlob(workbook, n);
    const sheetBlob = `${n} ${sheetText}`;
    if (sheetBlobIndicatesFormXLeaveRegister(sheetBlob)) return false;
    return blobIndicatesEmploymentCard(sheetBlob, null) || matchesFormXIVHint(sheetBlob);
  });
  return xivSheet || null;
}

/** Pick the CLRA Form XXI fines worksheet from a multi-tab AP workbook. */
export function resolveFormXXIWorkbookSheetName(workbook, hints = {}) {
  const names = Array.isArray(workbook?.SheetNames) ? workbook.SheetNames : [];
  if (names.length <= 1) return names[0] || null;

  const blob = [
    hints.fileName,
    hints.formFileName,
    hints.formName,
    hints.item?.formName,
    hints.item?.FormName,
    hints.formHeaderTitle,
    hints.formHeader?.title
  ]
    .filter(Boolean)
    .join(' ')
    .toLowerCase();

  let best = null;
  let bestScore = -Infinity;
  for (const name of names) {
    const sheetText = buildSheetTextBlob(workbook, name);
    const score = scoreFormXXIAPSheet(name, sheetText, blob);
    if (score > bestScore) {
      bestScore = score;
      best = name;
    }
  }
  if (bestScore > 0 && best) return best;

  const xxiSheet = names.find((n) => {
    const sheetText = buildSheetTextBlob(workbook, n);
    return sheetBlobIndicatesFormXXIRegisterOfFines(`${n} ${sheetText}`);
  });
  return xxiSheet || null;
}

/** Locate the fines register table on the correct CLRA Form XXI workbook tab (SheetJS). */
export function resolveFormXXIAPFinesTableLayout(workbook, hints = {}) {
  const names = Array.isArray(workbook?.SheetNames) ? workbook.SheetNames : [];
  if (!names.length) return null;

  const preferred =
    hints.preferredSheetName && workbook.Sheets?.[hints.preferredSheetName]
      ? hints.preferredSheetName
      : null;
  let sheetName = preferred;
  if (sheetName) {
    const prefText = buildSheetTextBlob(workbook, sheetName);
    const prefBlob = `${sheetName} ${prefText}`;
    if (blobIndicatesRegisterOfDeductionsForDamage(prefBlob)) {
      sheetName = null;
    }
    if (sheetBlobIndicatesFormXAPRegisterOfFines(prefBlob) && !sheetBlobIndicatesFormXXIRegisterOfFines(prefBlob)) {
      sheetName = null;
    }
  }
  if (!sheetName) {
    sheetName =
      resolveFormXXIWorkbookSheetName(workbook, hints) ||
      names.find((n) => {
        const sheetText = buildSheetTextBlob(workbook, n);
        return sheetBlobIndicatesFormXXIRegisterOfFines(`${n} ${sheetText}`);
      }) ||
      null;
  }
  if (!sheetName) return null;

  const ws = workbook?.Sheets?.[sheetName];
  if (!ws) return null;

  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  const merges = ws['!merges'] || [];
  const maxScan = Math.min(rows.length, 90);
  let best = null;
  let bestScore = -1;

  for (let r = 0; r < maxScan; r += 1) {
    const cells = [];
    const maxCols = Math.max((rows[r] || []).length, (rows[r + 1] || []).length, 24);
    for (let c = 0; c < maxCols; c += 1) {
      const text = getSheetMergedCellText(rows, merges, r, c);
      if (text) cells.push({ c, text, lower: text.toLowerCase() });
    }
    if (cells.length < 2) continue;

    const nextRowCells = [];
    for (let c = 0; c < maxCols; c += 1) {
      const text = getSheetMergedCellText(rows, merges, r + 1, c);
      if (text) nextRowCells.push(text.toLowerCase());
    }
    const rowText = cells.map((x) => x.lower).join(' ');
    const combinedText = `${rowText} ${nextRowCells.join(' ')}`.trim();
    if (blobIndicatesRegisterOfDeductionsForDamage(combinedText)) continue;
    if (
      sheetBlobIndicatesFormXAPRegisterOfFines(combinedText) &&
      !sheetBlobIndicatesFormXXIRegisterOfFines(combinedText)
    ) {
      continue;
    }

    let score = 0;
    if (/register\s+of\s+fines/i.test(combinedText)) score += 60;
    if (/act\/omission|fine\s+imposed|show\s+cause/i.test(combinedText)) score += 50;
    if (/rate\s+of\s+wages/i.test(combinedText)) score += 25;
    if (
      CLRA_WORKMEN_NAME_HEADER_RE.test(combinedText) ||
      /name\s+of\s+workmen|name\s+of\s+workman/.test(combinedText)
    ) {
      score += 35;
    }
    if (/father'?s\/husband'?s|nature\s+of\s+employ/.test(combinedText)) score += 25;
    if (/particulars\s+of\s+damage|date\s+of\s+damage|amount\s+of\s+deduction/.test(combinedText)) {
      score -= 45;
    }
    if (/s\.?\s*no|serial|sl\.?\s*no/.test(combinedText)) score += 15;
    score += cells.length * 2;

    if (score < 15) continue;

    let startCol = cells[0].c;
    for (const cell of cells) {
      if (/^s\.?\s*no\.?$|^sl\.?\s*no\.?$|serial/i.test(cell.lower)) {
        startCol = cell.c;
        break;
      }
    }

    const headers = [];
    for (let c = startCol; c < maxCols; c += 1) {
      let h = getSheetMergedCellText(rows, merges, r, c);
      if (!h) h = getSheetMergedCellText(rows, merges, r + 1, c);
      if (!h && r > 0) h = getSheetMergedCellText(rows, merges, r - 1, c);
      h = String(h || '').replace(/\s+/g, ' ').trim();
      if (!h) {
        if (headers.length > 0) break;
        continue;
      }
      if (/^\d+$/.test(h) && headers.length === 0) continue;
      headers.push(h);
    }
    if (headers.length < 3) continue;

    const rowHasSerialHeader = (rowIdx) => {
      for (let c = startCol; c < Math.min(startCol + 14, maxCols); c += 1) {
        const h = String(getSheetMergedCellText(rows, merges, rowIdx, c) || '')
          .toLowerCase()
          .replace(/\s+/g, ' ')
          .trim();
        if (/^s\.?\s*no\.?$|^sl\.?\s*no\.?$|serial(\s+number)?/.test(h)) return true;
      }
      return false;
    };

    let headerRowIndex = r;
    if (!rowHasSerialHeader(headerRowIndex) && rowHasSerialHeader(headerRowIndex + 1)) {
      headerRowIndex = r + 1;
    }

    const headerBand = resolveFormXXIClraHeaderBand(rows, merges, headerRowIndex, startCol, maxCols);
    let resolvedHeaders =
      headerBand.labels.length >= 3 ? headerBand.labels : [];
    if (headerBand.rowIdx >= 0 && headerBand.labels.length >= 3) {
      headerRowIndex = headerBand.rowIdx;
    }
    if (resolvedHeaders.length < 3) {
      resolvedHeaders = [];
      for (let c = startCol; c < maxCols; c += 1) {
        let h = getSheetMergedCellText(rows, merges, headerRowIndex, c);
        if (!h) h = getSheetMergedCellText(rows, merges, headerRowIndex + 1, c);
        if (!h && headerRowIndex > 0) h = getSheetMergedCellText(rows, merges, headerRowIndex - 1, c);
        h = String(h || '').replace(/\s+/g, ' ').trim();
        if (!h) {
          if (resolvedHeaders.length > 0) break;
          continue;
        }
        if (/^\d+$/.test(h) && resolvedHeaders.length === 0) continue;
        resolvedHeaders.push(h);
      }
    }
    if (resolvedHeaders.length < 3) continue;

    let dataStart = headerRowIndex + 1;
    const probeRow = rows[dataStart] || [];
    let seqHits = 0;
    for (let i = 0; i < Math.min(resolvedHeaders.length, 12); i += 1) {
      const t = String(probeRow[startCol + i] ?? '').trim();
      if (t === String(i + 1)) seqHits += 1;
    }
    if (seqHits >= 3) {
      dataStart = headerRowIndex + 2;
    } else {
      const probeText = probeRow
        .slice(startCol, startCol + 6)
        .map((x) => String(x || '').toLowerCase())
        .join(' ');
      if (
        /name\s+of|father|husband|nature\s+of|act\/omission|fine|show\s+cause|rate\s+of/.test(probeText) &&
        !/^\d+$/.test(String(probeRow[startCol] ?? '').trim())
      ) {
        dataStart = headerRowIndex + 2;
      }
    }

    const candidate = {
      sheetName,
      headerRowIndex,
      dataStartIndex: dataStart,
      tableStartCol: startCol,
      headers: resolvedHeaders
    };
    if (score > bestScore) {
      bestScore = score;
      best = candidate;
    }
  }

  return best;
}

/** CLRA Form XX — Register of Deductions for Damage or Loss. */
export function sheetBlobIndicatesFormXXAPRegisterOfDeductions(blob) {
  const text = String(blob || '').toLowerCase();
  if (sheetBlobIndicatesFormXAPRegisterOfFines(text) && !blobIndicatesRegisterOfDeductionsForDamage(text)) {
    return false;
  }
  if (
    matchesFormXXHint(text) &&
    /register\s+of\s+deductions|deductions?\s+for\s+damage|damage\s+or\s+loss/i.test(text)
  ) {
    return true;
  }
  return blobIndicatesRegisterOfDeductionsForDamage(text);
}

function scoreFormXXAPSheet(sheetName, sheetText, hintsBlob) {
  const sheetLower = String(sheetName || '').toLowerCase();
  const sheetBlob = `${sheetLower} ${sheetText}`;
  let score = 0;

  if (sheetBlobIndicatesFormXXAPRegisterOfDeductions(sheetBlob)) score += 200;
  if (matchesFormXXHint(sheetBlob)) score += 120;
  if (/register\s+of\s+deductions/i.test(sheetBlob)) score += 80;
  if (/particulars\s+of\s+damage|date\s+of\s+damage|amount\s+of\s+deduction/i.test(sheetText)) score += 60;
  if (CLRA_WORKMEN_NAME_HEADER_RE.test(sheetText) || /father'?s\/husband'?s/i.test(sheetText)) score += 40;
  if (sheetBlobIndicatesFormXXIRegisterOfFines(sheetBlob) && !blobIndicatesRegisterOfDeductionsForDamage(sheetBlob)) {
    score -= 220;
  }
  if (sheetBlobIndicatesFormXAPRegisterOfFines(sheetBlob)) score -= 150;
  if (matchesFormXXIHint(sheetBlob) && !blobIndicatesRegisterOfDeductionsForDamage(sheetBlob)) score -= 180;
  if (matchesFormXXHint(hintsBlob) && matchesFormXXHint(sheetLower)) score += 30;

  return score;
}

export function resolveFormXXWorkbookSheetName(workbook, hints = {}) {
  const names = Array.isArray(workbook?.SheetNames) ? workbook.SheetNames : [];
  if (names.length <= 1) return names[0] || null;

  const blob = [
    hints.fileName,
    hints.formFileName,
    hints.formName,
    hints.item?.formName,
    hints.item?.FormName,
    hints.formHeaderTitle,
    hints.formHeader?.title
  ]
    .filter(Boolean)
    .join(' ')
    .toLowerCase();

  let best = null;
  let bestScore = -Infinity;
  for (const name of names) {
    const sheetText = buildSheetTextBlob(workbook, name);
    const score = scoreFormXXAPSheet(name, sheetText, blob);
    if (score > bestScore) {
      bestScore = score;
      best = name;
    }
  }
  if (bestScore > 0 && best) return best;

  const xxSheet = names.find((n) => {
    const sheetText = buildSheetTextBlob(workbook, n);
    return sheetBlobIndicatesFormXXAPRegisterOfDeductions(`${n} ${sheetText}`);
  });
  return xxSheet || null;
}

/** Locate the deductions register table on the correct CLRA Form XX workbook tab (SheetJS). */
export function resolveFormXXAPDeductionsTableLayout(workbook, hints = {}) {
  const names = Array.isArray(workbook?.SheetNames) ? workbook.SheetNames : [];
  if (!names.length) return null;

  const preferred =
    hints.preferredSheetName && workbook.Sheets?.[hints.preferredSheetName]
      ? hints.preferredSheetName
      : null;
  let sheetName = preferred;
  if (sheetName) {
    const prefText = buildSheetTextBlob(workbook, sheetName);
    const prefBlob = `${sheetName} ${prefText}`;
    if (
      sheetBlobIndicatesFormXXIRegisterOfFines(prefBlob) &&
      !blobIndicatesRegisterOfDeductionsForDamage(prefBlob)
    ) {
      sheetName = null;
    }
    if (sheetBlobIndicatesFormXAPRegisterOfFines(prefBlob) && !blobIndicatesRegisterOfDeductionsForDamage(prefBlob)) {
      sheetName = null;
    }
  }
  if (!sheetName) {
    sheetName =
      resolveFormXXWorkbookSheetName(workbook, hints) ||
      names.find((n) => {
        const sheetText = buildSheetTextBlob(workbook, n);
        return sheetBlobIndicatesFormXXAPRegisterOfDeductions(`${n} ${sheetText}`);
      }) ||
      null;
  }
  if (!sheetName) return null;

  const ws = workbook?.Sheets?.[sheetName];
  if (!ws) return null;

  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  const merges = ws['!merges'] || [];
  const maxScan = Math.min(rows.length, 90);
  let best = null;
  let bestScore = -1;

  for (let r = 0; r < maxScan; r += 1) {
    const cells = [];
    const maxCols = Math.max((rows[r] || []).length, (rows[r + 1] || []).length, 24);
    for (let c = 0; c < maxCols; c += 1) {
      const text = getSheetMergedCellText(rows, merges, r, c);
      if (text) cells.push({ c, text, lower: text.toLowerCase() });
    }
    if (cells.length < 2) continue;

    const nextRowCells = [];
    for (let c = 0; c < maxCols; c += 1) {
      const text = getSheetMergedCellText(rows, merges, r + 1, c);
      if (text) nextRowCells.push(text.toLowerCase());
    }
    const rowText = cells.map((x) => x.lower).join(' ');
    const combinedText = `${rowText} ${nextRowCells.join(' ')}`.trim();
    if (
      sheetBlobIndicatesFormXXIRegisterOfFines(combinedText) &&
      !blobIndicatesRegisterOfDeductionsForDamage(combinedText)
    ) {
      continue;
    }
    if (
      sheetBlobIndicatesFormXAPRegisterOfFines(combinedText) &&
      !blobIndicatesRegisterOfDeductionsForDamage(combinedText)
    ) {
      continue;
    }

    let score = 0;
    if (/register\s+of\s+deductions/.test(combinedText)) score += 60;
    if (/deductions?\s+for\s+damage|damage\s+or\s+loss/.test(combinedText)) score += 50;
    if (/particulars\s+of\s+damage|date\s+of\s+damage|amount\s+of\s+deduction/.test(combinedText)) score += 40;
    if (
      CLRA_WORKMEN_NAME_HEADER_RE.test(combinedText) ||
      /name\s+of\s+workmen|name\s+of\s+workman/.test(combinedText)
    ) {
      score += 35;
    }
    if (/father'?s\/husband'?s|nature\s+of\s+employ/.test(combinedText)) score += 25;
    if (/show\s+cause|fine\s+imposed|nature\s*&\s*date\s+of\s+offence/.test(combinedText) && !/damage/.test(combinedText)) {
      score -= 45;
    }
    if (/surname/.test(combinedText) && /\bgender\b/.test(combinedText) && /date\s+of\s+birth|nationality|education/.test(combinedText)) {
      score -= 120;
    }
    if (/employees?\s*\/\s*workme|worker\s+code/.test(combinedText) && /surname/.test(combinedText)) {
      score -= 100;
    }
    if (/s\.?\s*no|serial|sl\.?\s*no/.test(combinedText)) score += 15;
    score += cells.length * 2;

    if (score < 15) continue;

    let startCol = cells[0].c;
    for (const cell of cells) {
      if (/^s\.?\s*no\.?$|^sl\.?\s*no\.?$|serial/i.test(cell.lower)) {
        startCol = cell.c;
        break;
      }
    }

    const headers = [];
    for (let c = startCol; c < maxCols; c += 1) {
      let h = getSheetMergedCellText(rows, merges, r, c);
      if (!h) h = getSheetMergedCellText(rows, merges, r + 1, c);
      if (!h && r > 0) h = getSheetMergedCellText(rows, merges, r - 1, c);
      h = String(h || '').replace(/\s+/g, ' ').trim();
      if (!h) {
        if (headers.length > 0) break;
        continue;
      }
      if (/^\d+$/.test(h) && headers.length === 0) continue;
      headers.push(h);
    }
    if (headers.length < 3) continue;

    const rowHasSerialHeader = (rowIdx) => {
      for (let c = startCol; c < Math.min(startCol + 14, maxCols); c += 1) {
        const h = String(getSheetMergedCellText(rows, merges, rowIdx, c) || '')
          .toLowerCase()
          .replace(/\s+/g, ' ')
          .trim();
        if (/^s\.?\s*no\.?$|^sl\.?\s*no\.?$|serial(\s+number)?/.test(h)) return true;
      }
      return false;
    };

    let headerRowIndex = r;
    if (!rowHasSerialHeader(headerRowIndex) && rowHasSerialHeader(headerRowIndex + 1)) {
      headerRowIndex = r + 1;
    }

    const headerBand = resolveFormXXIClraHeaderBand(rows, merges, headerRowIndex, startCol, maxCols);
    let resolvedHeaders =
      headerBand.labels.length >= 3 ? headerBand.labels : [];
    if (headerBand.rowIdx >= 0 && headerBand.labels.length >= 3) {
      headerRowIndex = headerBand.rowIdx;
    }
    if (resolvedHeaders.length < 3) {
      resolvedHeaders = [];
      for (let c = startCol; c < maxCols; c += 1) {
        let h = getSheetMergedCellText(rows, merges, headerRowIndex, c);
        if (!h) h = getSheetMergedCellText(rows, merges, headerRowIndex + 1, c);
        if (!h && headerRowIndex > 0) h = getSheetMergedCellText(rows, merges, headerRowIndex - 1, c);
        h = String(h || '').replace(/\s+/g, ' ').trim();
        if (!h) {
          if (resolvedHeaders.length > 0) break;
          continue;
        }
        if (/^\d+$/.test(h) && resolvedHeaders.length === 0) continue;
        resolvedHeaders.push(h);
      }
    }
    if (resolvedHeaders.length < 3) continue;

    const leafHeaderRows = [];
    for (let r = headerRowIndex; r <= headerRowIndex + 2; r += 1) {
      const band = [];
      for (let c = startCol; c < maxCols; c += 1) {
        band.push(getSheetMergedCellText(rows, merges, r, c));
      }
      leafHeaderRows.push(band);
    }
    const leafFromBand = resolveFormXXAPLeafHeaders(
      leafHeaderRows,
      merges,
      0,
      0,
      maxCols - startCol
    );
    if (
      leafFromBand.length >= 3 &&
      (statutoryTableHeadersLookMergedDuplicate(resolvedHeaders) ||
        leafFromBand.length >= resolvedHeaders.length)
    ) {
      resolvedHeaders = leafFromBand;
    }

    let dataStart = headerRowIndex + 1;
    const probeRow = rows[dataStart] || [];
    let seqHits = 0;
    for (let i = 0; i < Math.min(resolvedHeaders.length, 12); i += 1) {
      const t = String(probeRow[startCol + i] ?? '').trim();
      if (t === String(i + 1)) seqHits += 1;
    }
    if (seqHits >= 3) {
      dataStart = headerRowIndex + 2;
    } else {
      const probeText = probeRow
        .slice(startCol, startCol + 6)
        .map((x) => String(x || '').toLowerCase())
        .join(' ');
      if (
        /name\s+of|particulars|father|husband|nature\s+of|damage|deduction|amount\s+of/.test(probeText) &&
        !/^\d+$/.test(String(probeRow[startCol] ?? '').trim())
      ) {
        dataStart = headerRowIndex + 2;
      }
    }

    const candidate = {
      sheetName,
      headerRowIndex,
      dataStartIndex: dataStart,
      tableStartCol: startCol,
      headers: resolvedHeaders
    };
    if (score > bestScore) {
      bestScore = score;
      best = candidate;
    }
  }

  return best;
}

function excelCellLooksLikeSerialHeader(rawText) {
  const t = String(rawText || '')
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
  if (!t) return false;
  return (
    t === 's.no' ||
    t === 's no' ||
    t === 'sl.no' ||
    t === 'sl no' ||
    /^s\.?\s*no\.?$/.test(t) ||
    /^sl\.?\s*no\.?$/.test(t) ||
    /^serial(\s+number)?$/.test(t)
  );
}

function pickFormXXAPWorksheet(workbook, sheetNameHint = '') {
  let worksheet = workbook.worksheets[0];
  if (!sheetNameHint) return worksheet;
  const hintNorm = String(sheetNameHint).trim().toLowerCase();
  const hinted =
    workbook.getWorksheet(sheetNameHint) ||
    workbook.worksheets.find((ws) => String(ws?.name || '') === String(sheetNameHint)) ||
    workbook.worksheets.find((ws) => String(ws?.name || '').trim().toLowerCase() === hintNorm) ||
    workbook.worksheets.find((ws) => {
      const n = String(ws?.name || '').trim().toLowerCase();
      return hintNorm && (n.includes(hintNorm) || hintNorm.includes(n));
    });
  return hinted || worksheet;
}

function rowTextLooksLikeFormXXTableHeader(normalize, excelCellValueToString, worksheet, row, startCol, colCount = 8) {
  const parts = [];
  for (let j = 0; j < colCount; j += 1) {
    parts.push(normalize(excelCellValueToString(worksheet.getCell(row, startCol + j)?.value)));
  }
  const joined = parts.join(' ');
  if (!joined.trim()) return false;
  return (
    excelCellLooksLikeSerialHeader(parts[0]) ||
    /name\s+of\s+(?:the\s+)?workmen|name\s+of\s+(?:the\s+)?workman|name\s+of\s+(?:the\s+)?worker|name\s+of\s+(?:the\s+)?employee/.test(
      joined
    ) ||
    /particulars\s+of\s+damage|father|husband|nature\s+of\s+employ|date\s+of\s+damage|amount\s+of\s+deduction|date\s+of\s+recovery|act\/omission|fine\s+imposed|show\s+cause|rate\s+of\s+wages|wage\s*period.*wages?\s+payable|total\s+wages/.test(
      joined
    )
  );
}

function locateFormXXColumnHeaderRow(worksheet, hintHeaderRow, startCol, normalize, excelCellValueToString) {
  let bestRow = Math.max(1, hintHeaderRow);
  let bestScore = -1;
  for (let r = Math.max(1, hintHeaderRow - 2); r <= hintHeaderRow + 5; r += 1) {
    let score = 0;
    for (let c = startCol; c <= startCol + 14; c += 1) {
      const raw = excelCellValueToString(worksheet.getCell(r, c)?.value);
      if (excelCellLooksLikeSerialHeader(raw)) score += 50;
    }
    const joined = [];
    for (let c = startCol; c <= startCol + 14; c += 1) {
      joined.push(normalize(excelCellValueToString(worksheet.getCell(r, c)?.value)));
    }
    const text = joined.join(' ');
    if (/name\s+of\s+workmen|name\s+of\s+workman|name\s+of\s+(?:the\s+)?worker|name\s+of\s+(?:the\s+)?employee/.test(text)) {
      score += 40;
    }
    if (/particulars\s+of\s+damage|date\s+of\s+damage|amount\s+of\s+deduction/.test(text)) score += 30;
    if (/act\/omission|fine\s+imposed|show\s+cause|rate\s+of\s+wages|wage\s*period.*wages?\s+payable/.test(text)) {
      score += 30;
    }
    if (/father|husband|nature\s+of\s+employ/.test(text)) score += 20;
    if (/date\s+of\s+recovery/.test(text) && !/name\s+of\s+workmen|name\s+of\s+workman/.test(text)) score -= 25;
    if (score > bestScore) {
      bestScore = score;
      bestRow = r;
    }
  }
  return bestRow;
}

function locateFormXXDataStartRow(worksheet, columnHeaderRow, startCol, hdrCount, normalize, excelCellValueToString) {
  let dataStartRow = columnHeaderRow + 1;
  for (let guard = 0; guard < 3; guard += 1) {
    if (rowTextLooksLikeFormXXTableHeader(normalize, excelCellValueToString, worksheet, dataStartRow, startCol)) {
      dataStartRow += 1;
      continue;
    }
    let seqHits = 0;
    for (let i = 0; i < Math.min(hdrCount, 12); i += 1) {
      const t = String(excelCellValueToString(worksheet.getCell(dataStartRow, startCol + i)?.value) || '').trim();
      if (t === String(i + 1)) seqHits += 1;
    }
    if (seqHits >= 3) {
      dataStartRow += 1;
      continue;
    }
    break;
  }
  return dataStartRow;
}

/** Preserve original Form XX CLRA template styling (ExcelJS) while writing table rows. */
export async function buildFormXXAPWorkbookWithTemplateStyles({
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
  sheetNameHint = '',
  headerWriteMode = 'combined',
  formCRajasthanTitleAnchorCol = null,
}) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(templateArrayBuffer);
  const worksheet = pickFormXXAPWorksheet(workbook, sheetNameHint);
  if (!worksheet) throw new Error('Template worksheet not found.');

  const normalize = (txt) =>
    String(txt || '')
      .replace(/\r?\n/g, ' ')
      .replace(/\s+/g, ' ')
      .trim()
      .toLowerCase();
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

  if (Number(formCRajasthanTitleAnchorCol) > 0) {
    const establishment = String(
      headerFormData?.statutory_establishment_name ||
        headerFormData?.form_c_rj_establishment ||
        ''
    ).trim();
    const monthYear = String(headerFormData?.form_c_rj_month_year || '').trim();
    if (establishment || monthYear) {
      for (let row = 1; row <= 25; row += 1) {
        for (let col = 1; col <= 12; col += 1) {
          const cell = worksheet.getCell(row, col);
          const raw = excelCellValueToString(cell.value).replace(/\r\n/g, '\n');
          if (!/name\s+of\s+establishment/i.test(raw) || !/month\s*\/\s*year/i.test(raw)) continue;
          cell.value = [
            `Name of Establishment:- ${establishment}`.trimEnd(),
            `Month/Year: ${monthYear}`.trimEnd(),
          ].join('\n');
          cell.alignment = {
            ...(cell.alignment || {}),
            wrapText: true,
            vertical: 'middle',
            horizontal: 'left',
          };
          break;
        }
      }
    }
  }

  if (Number(formCRajasthanTitleAnchorCol) > 0) {
    const targetCol = Math.max(1, Number(formCRajasthanTitleAnchorCol));
    for (let row = 1; row <= 3; row += 1) {
      let sourceCol = 0;
      for (let col = 1; col <= 12; col += 1) {
        const value = excelCellValueToString(worksheet.getCell(row, col)?.value).trim();
        if (/^form\s*c$/i.test(value)) {
          sourceCol = col;
          break;
        }
      }
      if (!sourceCol || sourceCol === targetCol) continue;
      const source = worksheet.getCell(row, sourceCol);
      const target = worksheet.getCell(row, targetCol);
      if (target.isMerged) {
        const merges = Array.isArray(worksheet?.model?.merges) ? worksheet.model.merges : [];
        merges.forEach((range) => {
          const match = String(range || '').match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/i);
          if (!match || Number(match[2]) !== row || Number(match[4]) !== row) return;
          const toNumber = (letters) =>
            String(letters)
              .toUpperCase()
              .split('')
              .reduce((total, letter) => total * 26 + letter.charCodeAt(0) - 64, 0);
          const start = toNumber(match[1]);
          const end = toNumber(match[3]);
          if (sourceCol >= start && sourceCol <= end || targetCol >= start && targetCol <= end) {
            try {
              worksheet.unMergeCells(range);
            } catch (_) {
              // Keep the template merge when ExcelJS cannot remove it.
            }
          }
        });
      }
      const moved = worksheet.getCell(row, targetCol);
      moved.value = source.value;
      moved.style = { ...source.style };
      source.value = '';
      break;
    }
  }

  if (parsedHeaderRowIndex == null || parsedHeaderRowIndex < 0) {
    let autoHeaderRow = 0;
    let autoScore = -1;
    for (let r = 1; r <= Math.min(worksheet.rowCount || 90, 90); r += 1) {
      if (!rowTextLooksLikeFormXXTableHeader(normalize, excelCellValueToString, worksheet, r, 1, 12)) {
        continue;
      }
      let score = 0;
      for (let c = 1; c <= 12; c += 1) {
        const raw = excelCellValueToString(worksheet.getCell(r, c)?.value);
        if (excelCellLooksLikeSerialHeader(raw)) score += 50;
      }
      const joined = [];
      for (let c = 1; c <= 12; c += 1) {
        joined.push(normalize(excelCellValueToString(worksheet.getCell(r, c)?.value)));
      }
      const text = joined.join(' ');
      if (/name\s+of\s+(?:the\s+)?worker|name\s+of\s+(?:the\s+)?employee/.test(text)) score += 40;
      if (/rate\s+of\s+wages|fine\s+imposed|show\s+cause/.test(text)) score += 30;
      if (score > autoScore) {
        autoScore = score;
        autoHeaderRow = r;
      }
    }
    if (autoHeaderRow > 0) {
      parsedHeaderRowIndex = autoHeaderRow - 1;
    }
  }
  if (parsedHeaderRowIndex == null || parsedHeaderRowIndex < 0) {
    throw new Error('Could not locate Form XX header row.');
  }

  const hintHeaderRow = parsedHeaderRowIndex + 1;
  const hdrCount = Math.max(13, Array.isArray(headersToUse) ? headersToUse.length : 13);

  let startCol =
    parsedTableStartCol != null && parsedTableStartCol >= 0 ? parsedTableStartCol + 1 : 1;
  if (parsedTableStartCol == null || parsedTableStartCol < 0) {
    for (let c = 1; c <= 280; c += 1) {
      const raw = excelCellValueToString(worksheet.getCell(hintHeaderRow, c)?.value);
      if (excelCellLooksLikeSerialHeader(raw)) {
        startCol = c;
        break;
      }
    }
    if (startCol === 1) {
      for (let c = 1; c <= 280; c += 1) {
        const t = normalize(excelCellValueToString(worksheet.getCell(hintHeaderRow, c)?.value));
        if (t) {
          startCol = c;
          break;
        }
      }
    }
  }

  const columnHeaderRow = locateFormXXColumnHeaderRow(
    worksheet,
    hintHeaderRow,
    startCol,
    normalize,
    excelCellValueToString
  );

  let dataStartRow = locateFormXXDataStartRow(
    worksheet,
    columnHeaderRow,
    startCol,
    hdrCount,
    normalize,
    excelCellValueToString
  );

  const parsedDataStartRow =
    parsedDataStartIndex != null && parsedDataStartIndex >= 0 ? parsedDataStartIndex + 1 : 0;
  if (parsedDataStartRow > dataStartRow && parsedDataStartRow > columnHeaderRow) {
    dataStartRow = parsedDataStartRow;
  }
  if (dataStartRow <= columnHeaderRow) {
    dataStartRow = columnHeaderRow + 1;
  }
  while (
    dataStartRow <= columnHeaderRow + 2 &&
    rowTextLooksLikeFormXXTableHeader(normalize, excelCellValueToString, worksheet, dataStartRow, startCol)
  ) {
    dataStartRow += 1;
  }

  const headerValues = headerFormData && typeof headerFormData === 'object' ? headerFormData : {};
  writeStatutoryHeaderFieldsToExcelJsWorksheet(worksheet, {
    headerFormData: headerValues,
    parsedFormHeader,
    headerRowEnd: columnHeaderRow,
    maxScanCols: 24,
    writeMode: headerWriteMode,
  });

  const sourceHeaders = normalizeFormXXAPExportHeaders(
    Array.isArray(headersToUse) ? headersToUse.filter(Boolean) : []
  );
  const headerNormKey = (h) => normalize(String(h || '').replace(/[^a-z0-9]/gi, ' '));
  const templateColSpan = Math.max(hdrCount, sourceHeaders.length, 13) + 6;
  const templateLeafCols = resolveFormXXAPTemplateLeafColsFromWorksheet(
    worksheet,
    columnHeaderRow,
    startCol,
    templateColSpan,
    excelCellValueToString
  );
  const writableCols = mapFormXXAPWritableCols(sourceHeaders, templateLeafCols, startCol);

  const rowLooksMeaningful = (row) => {
    if (Array.isArray(row)) return row.some((v) => String(v ?? '').trim() !== '');
    if (row && typeof row === 'object') {
      return Object.entries(row).some(([key, v]) => {
        if (String(key || '').startsWith('__')) return false;
        return v != null && String(v).trim() !== '';
      });
    }
    return false;
  };

  const sourcePrimary =
    Array.isArray(mappedData) && mappedData.length > 0
      ? mappedData
      : Array.isArray(mappedRowMatrix)
        ? mappedRowMatrix
        : [];
  const sourceRows = sourcePrimary.filter((row) => rowLooksMeaningful(row));

  const tableColMax = resolveFormXXAPTableColMax(
    worksheet,
    columnHeaderRow,
    startCol,
    templateLeafCols,
    writableCols
  );
  const bodyRowCount = Math.max(
    sourceRows.length,
    countFormXXAPTemplateBodyRows(worksheet, dataStartRow, startCol, tableColMax)
  );
  const templateBodyBorderRow = findFormXXAPTemplateBodyBorderRow(
    worksheet,
    dataStartRow,
    startCol,
    tableColMax,
    dataStartRow + Math.max(bodyRowCount, 1) - 1
  );

  if (sourceRows.length > 0) {
    clearExcelJSWorksheetDataRange(
      worksheet,
      dataStartRow,
      bodyRowCount,
      startCol,
      tableColMax
    );
  }

  // Write only data cells — never clear header rows (preserves merged column headers).
  // Force NIL on damage / cause / installment / explanation-presence columns (no People names).
  for (let i = 0; i < sourceRows.length; i += 1) {
    const row = sourceRows[i];
    for (let j = 0; j < writableCols.length; j += 1) {
      const { col: targetCol, label } = writableCols[j];
      let value = rowValueForFormXXAPHeader(row, label, sourceHeaders);
      if (isFormXXAPDeductionNilHeader(label)) {
        value = FORM_XX_AP_DEDUCTION_COLUMN_NIL_TEXT;
      }
      const cell = worksheet.getCell(dataStartRow + i, targetCol);
      if (
        !isFormXXAPDeductionNilHeader(label) &&
        (value == null || value === '')
      ) {
        cell.value = '';
        continue;
      }
      if (
        mapFormXXAPHeaderNormKey(label).match(/^(s no|sl no|serial number)$/) &&
        (typeof value === 'number' ||
          (typeof value === 'string' && /^-?\d+(\.\d+)?$/.test(String(value).trim())))
      ) {
        cell.value = Number(value);
      } else {
        cell.value = String(value);
      }
    }
    for (let c = startCol; c <= tableColMax; c += 1) {
      const cell = worksheet.getCell(dataStartRow + i, c);
      if (cell.value == null) cell.value = '';
    }
  }

  if (sourceRows.length > 0) {
    ensureFormXXAPTableDataBorders(worksheet, {
      dataStartRow,
      dataRowCount: bodyRowCount,
      colFrom: startCol,
      colTo: tableColMax,
      templateBodyRow: templateBodyBorderRow,
    });
  }

  // Rajasthan Form C keeps the legal note in the template body; it belongs only in PDF.
  if (Number(formCRajasthanTitleAnchorCol) > 0) {
    worksheet.eachRow({ includeEmpty: false }, (row) => {
      row.eachCell({ includeEmpty: false }, (cell) => {
        const text = excelCellValueToString(cell.value);
        if (/\*?\s*applicable\s+only\s+in\s+case\s+of\s+damage\s*\/\s*loss\s*\/\s*fine/i.test(text)) {
          cell.value = '';
        }
      });
    });
  }

  const out = await workbook.xlsx.writeBuffer();
  const fileName =
    formFileName ||
    parsedFormHeader?.title?.replace(/[^a-zA-Z0-9]/g, '_') ||
    `Form_XX_${Date.now()}.xlsx`;
  const blob = new Blob([out], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  });
  return { blob, fileName };
}

/** Form XXI CLRA fines register — same Excel write path as Form XX deductions. */
export const buildFormXXIAPWorkbookWithTemplateStyles = buildFormXXAPWorkbookWithTemplateStyles;

function scoreFormXIIISheet(sheetName, sheetText, blob) {
  const sheetBlob = `${sheetName} ${sheetText}`.toLowerCase();
  let score = 0;
  const sheetToken = extractClraSheetFormRomanToken(sheetBlob);
  if (sheetToken === 'xiii') score += 140;
  if (sheetToken === 'xiv') score -= 200;
  if (sheetToken === 'xxi') score -= 180;
  if (matchesFormXIIIHint(blob) || matchesFormXIIIHint(sheetBlob)) score += 50;
  if (blobIndicatesRegisterOfWorkmen(sheetBlob, null)) score += 80;
  if (/register\s+of\s+workmen\s+employed\s+by\s+contractor/i.test(sheetBlob)) score += 60;
  if (REGISTER_OF_WORKMEN_SURNAME_HEADER_RE.test(sheetBlob)) score += 50;
  if (/age\s+and\s+sex/i.test(sheetBlob) && /local\s+address/i.test(sheetBlob)) score += 40;
  if (/employment\s+card/i.test(sheetBlob)) score -= 120;
  if (/register\s+of\s+fines|register\s+of\s+deductions/i.test(sheetBlob)) score -= 100;
  return score;
}

/** Pick CLRA Form XIII Register of Workmen worksheet from a multi-tab AP workbook. */
export function resolveFormXIIIWorkbookSheetName(workbook, hints = {}) {
  const names = Array.isArray(workbook?.SheetNames) ? workbook.SheetNames : [];
  if (names.length <= 1) return names[0] || null;

  const blob = [
    hints.fileName,
    hints.formFileName,
    hints.formName,
    hints.item?.formName,
    hints.item?.FormName,
    hints.formHeaderTitle,
    hints.formHeader?.title,
    hints.formHeaderSubtitle,
    hints.formHeader?.subtitle
  ]
    .filter(Boolean)
    .join(' ')
    .toLowerCase();

  let best = null;
  let bestScore = -Infinity;
  for (const name of names) {
    const sheetText = buildSheetTextBlob(workbook, name);
    const score = scoreFormXIIISheet(name, sheetText, blob);
    if (score > bestScore) {
      bestScore = score;
      best = name;
    }
  }
  if (bestScore > 0 && best) return best;

  const xiiiSheet = names.find((n) => {
    const sheetText = buildSheetTextBlob(workbook, n);
    return (
      matchesFormXIIIHint(`${n} ${sheetText}`) ||
      blobIndicatesRegisterOfWorkmen(`${n} ${sheetText}`, null)
    );
  });
  return xiiiSheet || null;
}

function formXIIIExcelCellText(val) {
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

function isFormXIIITitleBandText(text) {
  const s = String(text || '').replace(/\r?\n/g, ' ').replace(/\s+/g, ' ').trim();
  if (!s) return false;
  return (
    /form\s*[-–]?\s*xiii\b/i.test(s) ||
    /register\s+of\s+workmen\s+employed\s+by\s+contractor/i.test(s) ||
    (/vide\s+rule\s*75/i.test(s) && /contract\s+labou?r/i.test(s))
  );
}

/** AP CLRA Form XIII — canonical A–K layout (11 columns). */
export const FORM_XIII_AP_TABLE_COLS = 11;
export const FORM_XIII_AP_COL_WIDTHS = [6, 22, 10, 18, 18, 28, 18, 14, 14, 14, 16];
export const FORM_XIII_AP_MERGES = [
  'A3:K3',
  'A4:K4',
  'A5:A8',
  'B5:G8',
  'I5:I8',
  'J5:K8',
  'A9:A12',
  'B9:G12',
  'I9:I12',
  'J9:K12',
  'A13:A14',
  'B13:B14',
  'C13:C14',
  'D13:D14',
  'E13:E14',
  'F13:F14',
  'G13:G14',
  'H13:H14',
  'I13:I14',
  'J13:J14',
  'K13:K14',
];
export const FORM_XIII_AP_ROW_HEIGHTS = { 3: 30, 4: 18, 13: 42, 14: 42 };

function normalizeFormXIIITitleText(raw) {
  return String(raw || '')
    .replace(/_x000d_/gi, '\n')
    .replace(/\u000d/g, '\n')
    .replace(/\r\n/g, '\n')
    .replace(/\r/g, '\n')
    .replace(/\s+/g, ' ')
    .trim();
}

function cloneFormXIIICellStyle(style) {
  if (!style) return null;
  try {
    const cloned = JSON.parse(JSON.stringify(style));
    if (cloned.font?.color && cloned.font.color.indexed != null) {
      delete cloned.font.color.indexed;
      if (!cloned.font.color.argb && !cloned.font.color.theme) {
        cloned.font.color = { argb: 'FF000000' };
      }
    }
    return cloned;
  } catch (_) {
    return null;
  }
}

function copyFormXIIIExcelCellValue(raw) {
  if (raw == null) return null;
  if (typeof raw === 'string') return normalizeFormXIIITitleText(raw);
  if (typeof raw === 'number' || typeof raw === 'boolean') return raw;
  if (raw instanceof Date) return raw;
  if (typeof raw === 'object' && Array.isArray(raw.richText)) {
    return {
      richText: raw.richText.map((rt) => ({
        text: normalizeFormXIIITitleText(rt?.text || ''),
        font: rt?.font ? JSON.parse(JSON.stringify(rt.font)) : undefined,
      })),
    };
  }
  return raw;
}

function dedupeFormXIIITitleRow(worksheet, row, maxCol = FORM_XIII_AP_TABLE_COLS) {
  let masterCol = 0;
  let masterText = '';
  for (let c = 1; c <= maxCol; c += 1) {
    const text = formXIIIExcelCellText(worksheet.getCell(row, c)?.value).trim();
    if (!isFormXIIITitleBandText(text)) continue;
    if (!masterCol || text.length > masterText.length) {
      masterCol = c;
      masterText = text;
    }
  }
  if (!masterCol) return;
  for (let c = 1; c <= maxCol; c += 1) {
    const cell = worksheet.getCell(row, c);
    if (c === masterCol) {
      cell.value = masterText;
      continue;
    }
    const text = formXIIIExcelCellText(cell?.value).trim();
    if (isFormXIIITitleBandText(text)) cell.value = null;
  }
}

/**
 * Rebuild a clean A–K workbook so Excel does not repair phantom columns / broken merges
 * from the Form Master template (same strategy as TN Form I workmen register).
 */
export function cloneFormXIIIAPWorksheetClean(sourceWs, sheetName = 'FORM XIII') {
  const outWb = new ExcelJS.Workbook();
  const safeName = String(sheetName || 'FORM XIII')
    .replace(/[\\/*?[\]:]/g, ' ')
    .trim()
    .slice(0, 31) || 'FORM XIII';
  // Prefer a stable Form XIII sheet name — never keep a misleading "FORM 1" tab label.
  const outName =
    /form\s*1\b/i.test(safeName) && !matchesFormXIIIHint(safeName)
      ? 'FORM XIII'
      : matchesFormXIIIHint(safeName) || /workmen/i.test(safeName)
        ? safeName
        : 'FORM XIII';
  const outWs = outWb.addWorksheet(outName);
  const maxCol = FORM_XIII_AP_TABLE_COLS;
  const maxRow = Math.min(
    Math.max(sourceWs?.actualRowCount || 0, sourceWs?.rowCount || 0, 30),
    200
  );

  const FORM_XIII_TITLE =
    'FORM - XIII REGISTER OF WORKMEN EMPLOYED BY CONTRACTOR';
  const FORM_XIII_RULE =
    '[Vide Rule 75 Contract Labour (Regulation and Abolition) Central/A.P Rules)';

  // Copy only body/meta cells — title + stacked header masters are rewritten below.
  for (let r = 1; r <= maxRow; r += 1) {
    const srcRow = sourceWs.getRow(r);
    const dstRow = outWs.getRow(r);
    if (srcRow.height != null) dstRow.height = srcRow.height;
    for (let c = 1; c <= maxCol; c += 1) {
      const src = sourceWs.getCell(r, c);
      const dst = outWs.getCell(r, c);
      // Skip title band + stacked header slave row — rewritten canonically below.
      if (r === 3 || r === 4 || r === 14) continue;
      const copied = copyFormXIIIExcelCellValue(src.value);
      if (copied != null) {
        // Drop duplicated title fragments that leaked into contractor/header rows.
        if (isFormXIIITitleBandText(formXIIIExcelCellText(copied))) continue;
        dst.value = copied;
      }
      const cloned = cloneFormXIIICellStyle(src.style);
      if (cloned) dst.style = cloned;
    }
  }

  FORM_XIII_AP_COL_WIDTHS.forEach((w, i) => {
    outWs.getColumn(i + 1).width = w;
  });
  Object.entries(FORM_XIII_AP_ROW_HEIGHTS).forEach(([row, height]) => {
    outWs.getRow(Number(row)).height = height;
  });

  // Absolute single-title band: clear every cell then write once in A3 / A4.
  for (let r = 3; r <= 4; r += 1) {
    for (let c = 1; c <= maxCol; c += 1) {
      outWs.getCell(r, c).value = null;
    }
  }
  let titleText = '';
  let ruleText = '';
  for (let c = 1; c <= Math.max(maxCol, 16); c += 1) {
    const t3 = formXIIIExcelCellText(sourceWs.getCell(3, c)?.value).trim();
    const t4 = formXIIIExcelCellText(sourceWs.getCell(4, c)?.value).trim();
    if (isFormXIIITitleBandText(t3) && t3.length > titleText.length) titleText = t3;
    if (isFormXIIITitleBandText(t4) && t4.length > ruleText.length) ruleText = t4;
    // Some templates put title+rule in one cell on row 3.
    if (/form\s*[-–]?\s*xiii/i.test(t3) && /vide\s+rule\s*75/i.test(t3) && !ruleText) {
      const parts = t3.split(/\[Vide/i);
      if (parts.length >= 2) {
        titleText = parts[0].trim();
        ruleText = `[Vide${parts.slice(1).join('[Vide')}`.trim();
      }
    }
  }
  if (!titleText) titleText = FORM_XIII_TITLE;
  if (!ruleText) ruleText = FORM_XIII_RULE;
  // Prefer short title without embedded rule when both were glued together.
  if (/form\s*[-–]?\s*xiii/i.test(titleText) && /vide\s+rule\s*75/i.test(titleText) && ruleText) {
    titleText = titleText.replace(/\s*\[Vide[\s\S]*$/i, '').trim() || FORM_XIII_TITLE;
  }

  outWs.getCell(3, 1).value = titleText;
  outWs.getCell(3, 1).alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  outWs.getCell(3, 1).font = { bold: true, size: 12 };
  outWs.getCell(4, 1).value = ruleText;
  outWs.getCell(4, 1).alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  outWs.getCell(4, 1).font = { bold: false, size: 10 };

  // Stacked column headers: keep values only on row 13 (merge master); clear row 14.
  for (let c = 1; c <= maxCol; c += 1) {
    const headerLabel =
      formXIIIExcelCellText(sourceWs.getCell(13, c)?.value).trim() ||
      formXIIIExcelCellText(sourceWs.getCell(14, c)?.value).trim();
    outWs.getCell(13, c).value = headerLabel || null;
    outWs.getCell(14, c).value = null;
    const headerStyle =
      cloneFormXIIICellStyle(sourceWs.getCell(13, c)?.style) ||
      cloneFormXIIICellStyle(sourceWs.getCell(14, c)?.style);
    if (headerStyle) {
      outWs.getCell(13, c).style = headerStyle;
      outWs.getCell(14, c).style = headerStyle;
    }
    outWs.getCell(13, c).alignment = {
      wrapText: true,
      vertical: 'middle',
      horizontal: 'center',
    };
    outWs.getCell(13, c).font = {
      ...(outWs.getCell(13, c).font || {}),
      bold: true,
    };
  }

  FORM_XIII_AP_MERGES.forEach((range) => {
    try {
      outWs.mergeCells(range);
    } catch (_) {
      /* ignore overlapping template merges */
    }
  });

  // Template / prior drafts often copy "Label : value" into both A and B (and I and J).
  dedupeFormXIIIMetaHeaderBand(outWs);

  return { workbook: outWb, worksheet: outWs };
}

function labelMatchesFormXIIIHeaderField(raw, label) {
  const rawLabelOnly = String(raw).split(':')[0].trim();
  const rawKey = statutoryHeaderLabelMatchKey(rawLabelOnly);
  const labelKey = statutoryHeaderLabelMatchKey(label);
  if (rawKey && labelKey && rawKey === labelKey) return true;
  const rawNorm = normalizeStatutoryHeaderLabel(String(rawLabelOnly).replace(/:+$/, ''));
  const labelNorm = normalizeStatutoryHeaderLabel(String(label).replace(/:+$/, ''));
  if (!rawNorm || !labelNorm) return false;
  if (rawNorm === labelNorm) return true;
  return rawNorm.includes(labelNorm) || labelNorm.includes(rawNorm);
}

function splitFormXIIILabelAndValue(raw) {
  const text = String(raw || '').trim();
  if (!text) return { label: '', value: '' };
  const idx = text.indexOf(':');
  if (idx < 0) return { label: text.replace(/\.+$/, '').trim(), value: '' };
  return {
    label: text.slice(0, idx).replace(/\.+$/, '').trim(),
    value: text.slice(idx + 1).trim(),
  };
}

function isFormXIIIMetaHeaderLabelText(text) {
  const s = normalizeStatutoryHeaderLabel(String(text || '').split(':')[0]);
  if (!s) return false;
  return (
    /name\s+and\s+address\s+of\s+contractor/.test(s) ||
    /(?:name|nature)\s+and\s+location\s+of\s+work/.test(s) ||
    /establ(?:ishment|ishemnt)\s+in.*under\s+which\s+contract/.test(s) ||
    /name\s+and\s+address\s+of\s+principal\s+employer/.test(s) ||
    /name\s+and\s+addr/.test(s)
  );
}

/**
 * Meta band (rows 5–12): keep one "Label : value" line in value columns B / J only.
 * Clear duplicate combined text from narrow label columns A / I (user-visible double display).
 */
function dedupeFormXIIIMetaHeaderBand(worksheet) {
  if (!worksheet) return;

  const blocks = [
    { labelRow: 5, labelCol: 1, valueCol: 2 }, // Contractor — A / B
    { labelRow: 9, labelCol: 1, valueCol: 2 }, // Nature/location — A / B
    { labelRow: 5, labelCol: 9, valueCol: 10 }, // Establishment — I / J
    { labelRow: 9, labelCol: 9, valueCol: 10, boldValue: true }, // Principal employer — I / J
  ];

  for (const block of blocks) {
    const labelCell = worksheet.getCell(block.labelRow, block.labelCol);
    const valueCell = worksheet.getCell(block.labelRow, block.valueCol);
    const labelRaw = formXIIIExcelCellText(labelCell?.value).trim();
    const valueRaw = formXIIIExcelCellText(valueCell?.value).trim();

    let bestLabel = '';
    let bestValue = '';

    const fromLabel = splitFormXIIILabelAndValue(labelRaw);
    const fromValue = splitFormXIIILabelAndValue(valueRaw);

    if (isFormXIIIMetaHeaderLabelText(fromLabel.label) || isFormXIIIMetaHeaderLabelText(labelRaw)) {
      bestLabel = fromLabel.label || String(labelRaw).split(':')[0].trim();
      if (fromLabel.value) bestValue = fromLabel.value;
    }
    if (isFormXIIIMetaHeaderLabelText(fromValue.label) || isFormXIIIMetaHeaderLabelText(valueRaw)) {
      if (!bestLabel) bestLabel = fromValue.label || String(valueRaw).split(':')[0].trim();
      if (fromValue.value) bestValue = fromValue.value;
      else if (!bestValue && valueRaw && !isFormXIIIMetaHeaderLabelText(valueRaw)) {
        bestValue = valueRaw;
      }
    } else if (valueRaw && !bestValue) {
      bestValue = valueRaw;
    }

    // Always clear narrow label columns A / I — combined text belongs only in B / J.
    labelCell.value = null;

    if (bestLabel || bestValue) {
      const isPrincipalEmployer =
        block.boldValue === true ||
        /principal\s+employer/i.test(bestLabel) ||
        /principal\s+employer/i.test(valueRaw);
      if (bestValue && isPrincipalEmployer) {
        // Label normal, Principal Employer value bold (matches Form XIII Excel layout).
        valueCell.value = {
          richText: [
            { text: `${bestLabel || 'Name and address of Principal Employer'} : `, font: { bold: false } },
            { text: bestValue, font: { bold: true } },
          ],
        };
      } else {
        valueCell.value = bestValue
          ? `${bestLabel || 'Details'} : ${bestValue}`
          : bestLabel
            ? `${bestLabel} :`
            : null;
      }
      valueCell.alignment = {
        ...(valueCell.alignment || {}),
        wrapText: true,
        vertical: 'top',
        horizontal: 'left',
      };
      if (isPrincipalEmployer && bestValue) {
        valueCell.font = { ...(valueCell.font || {}), bold: true };
      }
    }
  }

  // Clear any leftover duplicate fragments in A/I across the full meta band.
  for (let r = 5; r <= 12; r += 1) {
    for (const c of [1, 9]) {
      const cell = worksheet.getCell(r, c);
      const text = formXIIIExcelCellText(cell?.value).trim();
      if (!text) continue;
      if (isFormXIIIMetaHeaderLabelText(text) || /:/i.test(text)) {
        cell.value = null;
      }
    }
  }
}

/** Write contractor / establishment values once into columns B / J (never into A / I). */
function writeFormXIIIAPHeaderFieldsToWorksheet(
  worksheet,
  { headerFormData, parsedFormHeader, headerRowEnd = 12 } = {}
) {
  if (!worksheet || !headerFormData || typeof headerFormData !== 'object') return;
  const fields = Array.isArray(parsedFormHeader?.fields) ? parsedFormHeader.fields : [];
  const written = new Set();

  const writeAt = (r, c, label, raw, val) => {
    if (!val) return;
    const labelOnly = String(raw).split(':')[0].trim().replace(/\.+$/, '');
    // Official layout: labels sit in A/I merges; values belong in B/J.
    // Never write combined text into A (1) or I (9) — that duplicates B/J.
    const isLabelCol = c === 1 || c === 9;
    const valueCol = isLabelCol ? c + 1 : c;
    const labelCol = isLabelCol ? c : c === 2 || c === 10 ? c - 1 : 0;

    if (labelCol === 1 || labelCol === 9) {
      worksheet.getCell(r, labelCol).value = null;
    }
    const valueCell = worksheet.getCell(r, valueCol);
    const isPrincipalEmployer = /principal\s+employer/i.test(labelOnly) || /principal\s+employer/i.test(label);
    if (isPrincipalEmployer) {
      // Label normal, Principal Employer name/address bold.
      valueCell.value = {
        richText: [
          {
            text: `${String(label || labelOnly).replace(/:+\s*$/, '').trim()} : `,
            font: { bold: false },
          },
          { text: String(val), font: { bold: true } },
        ],
      };
      valueCell.font = { ...(valueCell.font || {}), bold: true };
    } else {
      valueCell.value = formatStatutoryHeaderLabelValueExport(label, labelOnly, val);
    }
    valueCell.alignment = {
      ...(valueCell.alignment || {}),
      wrapText: true,
      vertical: 'top',
      horizontal: 'left',
    };
  };

  for (let r = 1; r <= headerRowEnd; r += 1) {
    for (let c = 1; c <= FORM_XIII_AP_TABLE_COLS; c += 1) {
      const raw = formXIIIExcelCellText(worksheet.getCell(r, c)?.value).trim();
      if (!raw || isFormXIIITitleBandText(raw)) continue;

      let wrote = false;
      for (const field of fields) {
        const key = String(field?.key || '').trim();
        const label = String(field?.label || '').trim();
        if (!key || !label || written.has(key)) continue;
        if (!labelMatchesFormXIIIHeaderField(raw, label)) continue;
        const val = resolveHeaderFieldExportValue(headerFormData, field);
        if (!val) continue;
        writeAt(r, c, label, raw, val);
        written.add(key);
        wrote = true;
        break;
      }
      if (wrote) continue;

      // Spec fallback so Establishemnt typo / missing parsed fields still get company value.
      const labelOnly = String(raw).split(':')[0].trim();
      const matchKey = statutoryHeaderLabelMatchKey(labelOnly);
      if (matchKey === 'statutory_establishment_contract' && !written.has(matchKey)) {
        const val = resolveHeaderFieldExportValue(headerFormData, {
          key: 'form_xiii_establishment_contract_carried',
          label: labelOnly,
        });
        if (val) {
          writeAt(r, c, labelOnly, raw, val);
          written.add(matchKey);
        }
      }
    }
  }

  dedupeFormXIIIMetaHeaderBand(worksheet);
}

function rowLooksLikeFormXIIIIndexRow(worksheet, row, startCol, colCount = FORM_XIII_AP_TABLE_COLS) {
  let hits = 0;
  for (let i = 0; i < colCount; i += 1) {
    const t = String(formXIIIExcelCellText(worksheet.getCell(row, startCol + i)?.value) || '').trim();
    if (t === String(i + 1)) hits += 1;
  }
  return hits >= 6;
}

function resolveFormXIIIAPTableLayout(worksheet, parsedHeaderRowIndex, parsedDataStartIndex, parsedTableStartCol) {
  const startCol =
    parsedTableStartCol != null && parsedTableStartCol >= 0 ? parsedTableStartCol + 1 : 1;
  let columnHeaderRow = parsedHeaderRowIndex != null && parsedHeaderRowIndex >= 0
    ? parsedHeaderRowIndex + 1
    : 13;
  let dataStartRow =
    parsedDataStartIndex != null && parsedDataStartIndex >= 0 ? parsedDataStartIndex + 1 : 16;

  const titleText = formXIIIExcelCellText(worksheet.getCell(3, 1)?.value);
  if (isFormXIIITitleBandText(titleText) && excelCellLooksLikeSerialHeader(formXIIIExcelCellText(worksheet.getCell(13, 1)?.value))) {
    columnHeaderRow = 13;
    dataStartRow = rowLooksLikeFormXIIIIndexRow(worksheet, 15, startCol) ? 16 : 15;
    while (
      dataStartRow <= columnHeaderRow + 3 &&
      (rowTextLooksLikeFormXIIITableHeader(
        (txt) => String(txt || '').replace(/\s+/g, ' ').trim().toLowerCase(),
        worksheet,
        dataStartRow,
        startCol
      ) ||
        rowLooksLikeFormXIIIIndexRow(worksheet, dataStartRow, startCol))
    ) {
      dataStartRow += 1;
    }
    return { startCol, columnHeaderRow, dataStartRow };
  }

  return null;
}

function rowTextLooksLikeFormXIIITableHeader(normalize, worksheet, row, startCol, colCount = 11) {
  const parts = [];
  for (let j = 0; j < colCount; j += 1) {
    parts.push(normalize(formXIIIExcelCellText(worksheet.getCell(row, startCol + j)?.value)));
  }
  const joined = parts.join(' ');
  if (!joined.trim()) return false;
  return (
    excelCellLooksLikeSerialHeader(parts[0]) ||
    REGISTER_OF_WORKMEN_SURNAME_HEADER_RE.test(joined) ||
    (/age\s+and\s+sex/.test(joined) && /local\s+address/.test(joined)) ||
    (/father|husband/.test(joined) && /nature\s+of\s+employ/.test(joined)) ||
    (/permanent\s+house\s+address/.test(joined) && /date\s+of\s+commencement/.test(joined))
  );
}

function locateFormXIIIColumnHeaderRow(worksheet, hintHeaderRow, startCol, normalize) {
  let bestRow = Math.max(1, hintHeaderRow);
  let bestScore = -1;
  for (let r = Math.max(1, hintHeaderRow - 2); r <= hintHeaderRow + 6; r += 1) {
    let score = 0;
    for (let c = startCol; c <= startCol + 14; c += 1) {
      const raw = formXIIIExcelCellText(worksheet.getCell(r, c)?.value);
      if (excelCellLooksLikeSerialHeader(raw)) score += 50;
    }
    const joined = [];
    for (let c = startCol; c <= startCol + 14; c += 1) {
      joined.push(normalize(formXIIIExcelCellText(worksheet.getCell(r, c)?.value)));
    }
    const text = joined.join(' ');
    if (REGISTER_OF_WORKMEN_SURNAME_HEADER_RE.test(text)) score += 50;
    if (/age\s+and\s+sex/.test(text)) score += 25;
    if (/father|husband/.test(text)) score += 20;
    if (/nature\s+of\s+employ|designation/.test(text)) score += 20;
    if (/permanent\s+house\s+address/.test(text)) score += 20;
    if (/local\s+address/.test(text)) score += 15;
    if (/date\s+of\s+commencement/.test(text)) score += 15;
    if (/signature|thumb/.test(text)) score += 10;
    if (/reasons?\s+for\s+termination/.test(text)) score += 10;
    if (score > bestScore) {
      bestScore = score;
      bestRow = r;
    }
  }
  return bestRow;
}

function locateFormXIIIDataStartRow(worksheet, columnHeaderRow, startCol, hdrCount, normalize) {
  let dataStartRow = columnHeaderRow + 1;
  for (let guard = 0; guard < 4; guard += 1) {
    if (rowTextLooksLikeFormXIIITableHeader(normalize, worksheet, dataStartRow, startCol)) {
      dataStartRow += 1;
      continue;
    }
    let seqHits = 0;
    for (let i = 0; i < Math.min(hdrCount, 12); i += 1) {
      const t = String(formXIIIExcelCellText(worksheet.getCell(dataStartRow, startCol + i)?.value) || '').trim();
      if (t === String(i + 1)) seqHits += 1;
    }
    if (seqHits >= 3) {
      dataStartRow += 1;
      continue;
    }
    break;
  }
  return dataStartRow;
}

/** Canonical Form XIII column widths so wrapped headers stay readable. */
function applyFormXIIIColumnLayout(worksheet, startCol, headerRow, dataStartRow) {
  if (!worksheet) return;
  for (let i = 0; i < FORM_XIII_AP_COL_WIDTHS.length; i += 1) {
    const col = worksheet.getColumn(startCol + i);
    const next = FORM_XIII_AP_COL_WIDTHS[i];
    const cur = Number(col.width) || 0;
    if (cur < next * 0.85) col.width = next;
  }
  for (let r = headerRow; r < dataStartRow; r += 1) {
    const rowObj = worksheet.getRow(r);
    const curH = Number(rowObj.height) || 15;
    if (curH < 36) rowObj.height = 42;
    for (let c = startCol; c < startCol + FORM_XIII_AP_COL_WIDTHS.length; c += 1) {
      const cell = worksheet.getCell(r, c);
      cell.alignment = {
        ...(cell.alignment || {}),
        wrapText: true,
        vertical: 'middle',
        horizontal: 'center',
      };
    }
  }
}

function clearFormXIIIWorksheetDataRange(worksheet, startRow, rowCount, colFrom, colTo) {
  if (!worksheet || rowCount < 1) return;
  for (let r = startRow; r < startRow + rowCount; r += 1) {
    for (let c = colFrom; c <= colTo; c += 1) {
      worksheet.getCell(r, c).value = null;
    }
  }
}

function pickFormXIIIAPSourceWorksheet(workbook, sheetNameHint = '') {
  let worksheet = pickFormXXAPWorksheet(workbook, sheetNameHint);
  if (!workbook.worksheets?.length) return worksheet;
  if (workbook.worksheets.length === 1) return worksheet || workbook.worksheets[0];

  let best = worksheet || workbook.worksheets[0];
  let bestScore = -1;
  for (const ws of workbook.worksheets) {
    let score = 0;
    const name = String(ws?.name || '');
    if (sheetNameHint && name === String(sheetNameHint)) score += 200;
    if (matchesFormXIIIHint(name)) score += 40;
    if (/xiv|employment\s+card/i.test(name)) score -= 80;
    for (let r = 1; r <= 18; r += 1) {
      for (let c = 1; c <= FORM_XIII_AP_TABLE_COLS; c += 1) {
        const t = formXIIIExcelCellText(ws.getCell(r, c)?.value);
        if (isFormXIIITitleBandText(t)) score += 30;
        if (REGISTER_OF_WORKMEN_SURNAME_HEADER_RE.test(t)) score += 40;
        if (excelCellLooksLikeSerialHeader(t)) score += 10;
      }
    }
    if (score > bestScore) {
      bestScore = score;
      best = ws;
    }
  }
  return best;
}

/**
 * Preserve Form XIII AP template styling (ExcelJS) while writing header fields + workmen rows.
 * Rebuilds a clean A–K sheet so Excel does not repair phantom columns / broken merges.
 */
export async function buildFormXIIIAPWorkbookWithTemplateStyles({
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
  sheetNameHint = '',
}) {
  const templateWb = new ExcelJS.Workbook();
  await templateWb.xlsx.load(templateArrayBuffer);
  const sourceWs = pickFormXIIIAPSourceWorksheet(templateWb, sheetNameHint);
  if (!sourceWs) throw new Error('Template worksheet not found.');

  const cleaned = cloneFormXIIIAPWorksheetClean(sourceWs, sourceWs.name || sheetNameHint || 'FORM XIII');
  const workbook = cleaned.workbook;
  const worksheet = cleaned.worksheet;

  const normalize = (txt) =>
    String(txt || '')
      .replace(/\r?\n/g, ' ')
      .replace(/\s+/g, ' ')
      .trim()
      .toLowerCase();

  const hdrCount = Math.max(FORM_XIII_AP_TABLE_COLS, Array.isArray(headersToUse) ? headersToUse.length : 0);
  const canonicalLayout = resolveFormXIIIAPTableLayout(
    worksheet,
    parsedHeaderRowIndex,
    parsedDataStartIndex,
    parsedTableStartCol
  );

  let startCol = 1;
  let columnHeaderRow = 13;
  let dataStartRow = 16;

  if (canonicalLayout) {
    ({ startCol, columnHeaderRow, dataStartRow } = canonicalLayout);
  } else {
    let hintHeaderRow =
      parsedHeaderRowIndex != null && parsedHeaderRowIndex >= 0 ? parsedHeaderRowIndex + 1 : 13;
    startCol =
      parsedTableStartCol != null && parsedTableStartCol >= 0 ? parsedTableStartCol + 1 : 1;
    if (parsedTableStartCol == null || parsedTableStartCol < 0) {
      for (let c = 1; c <= FORM_XIII_AP_TABLE_COLS; c += 1) {
        const raw = formXIIIExcelCellText(worksheet.getCell(hintHeaderRow, c)?.value);
        if (excelCellLooksLikeSerialHeader(raw)) {
          startCol = c;
          break;
        }
      }
    }
    columnHeaderRow = locateFormXIIIColumnHeaderRow(worksheet, hintHeaderRow, startCol, normalize);
    dataStartRow = locateFormXIIIDataStartRow(
      worksheet,
      columnHeaderRow,
      startCol,
      hdrCount,
      normalize
    );
    const parsedDataStartRow =
      parsedDataStartIndex != null && parsedDataStartIndex >= 0 ? parsedDataStartIndex + 1 : 0;
    if (parsedDataStartRow > dataStartRow && parsedDataStartRow > columnHeaderRow) {
      dataStartRow = parsedDataStartRow;
    }
    if (dataStartRow <= columnHeaderRow) dataStartRow = columnHeaderRow + 1;
    while (
      dataStartRow <= columnHeaderRow + 3 &&
      (rowTextLooksLikeFormXIIITableHeader(normalize, worksheet, dataStartRow, startCol) ||
        rowLooksLikeFormXIIIIndexRow(worksheet, dataStartRow, startCol))
    ) {
      dataStartRow += 1;
    }
  }

  writeFormXIIIAPHeaderFieldsToWorksheet(worksheet, {
    headerFormData,
    parsedFormHeader,
    headerRowEnd: columnHeaderRow - 1,
  });

  const sourceHeaders = Array.isArray(headersToUse) ? headersToUse.filter(Boolean) : [];
  const writableCols =
    sourceHeaders.length > 0
      ? sourceHeaders.map((label, idx) => ({
          col: startCol + idx,
          label: String(label || ''),
        }))
      : Array.from({ length: hdrCount }, (_, idx) => ({
          col: startCol + idx,
          label: formXIIIExcelCellText(worksheet.getCell(columnHeaderRow, startCol + idx)?.value),
        }));

  const rowValuesForExport = (row) => exportFormXXAPRowValuesByHeaders(row, sourceHeaders);

  const rowLooksMeaningful = (row) => {
    if (Array.isArray(row)) return row.some((v) => String(v ?? '').trim() !== '');
    if (row && typeof row === 'object') {
      return Object.entries(row).some(([key, v]) => {
        if (String(key || '').startsWith('__')) return false;
        return v != null && String(v).trim() !== '';
      });
    }
    return false;
  };

  const sourcePrimary =
    Array.isArray(mappedData) && mappedData.length > 0
      ? mappedData
      : Array.isArray(mappedRowMatrix)
        ? mappedRowMatrix
        : [];
  const sourceRows = sourcePrimary.filter((row) => rowLooksMeaningful(row));

  const tableColMax =
    writableCols.length > 0 ? writableCols[writableCols.length - 1].col : startCol + hdrCount - 1;
  const countTemplateBodyRows = () => {
    let rows = 0;
    for (let r = dataStartRow; r < dataStartRow + 120; r += 1) {
      let hasBorder = false;
      for (let c = startCol; c <= tableColMax; c += 1) {
        const b = worksheet.getCell(r, c)?.border;
        if (b && (b.top?.style || b.bottom?.style || b.left?.style || b.right?.style)) {
          hasBorder = true;
          break;
        }
      }
      if (hasBorder) rows += 1;
      else if (rows > 0) break;
    }
    return Math.max(rows, 1);
  };

  if (sourceRows.length > 0) {
    clearFormXIIIWorksheetDataRange(
      worksheet,
      dataStartRow,
      Math.max(sourceRows.length, countTemplateBodyRows()),
      startCol,
      tableColMax
    );
    ensureExcelJSDataRowsWithBorders(worksheet, {
      dataStartRow,
      dataRowCount: sourceRows.length,
      colFrom: startCol,
      colTo: tableColMax,
      templateRow: dataStartRow,
      templateBodyRows: countTemplateBodyRows(),
    });
  }

  for (let i = 0; i < sourceRows.length; i += 1) {
    const row = sourceRows[i];
    const rowValues = rowValuesForExport(row);
    for (let j = 0; j < writableCols.length; j += 1) {
      const { col: targetCol } = writableCols[j];
      let value = rowValues[j];
      if (
        (value == null || value === '') &&
        j === 0 &&
        excelCellLooksLikeSerialHeader(writableCols[j]?.label || sourceHeaders[0])
      ) {
        value = i + 1;
      }
      if (value == null || value === '') continue;
      const cell = worksheet.getCell(dataStartRow + i, targetCol);
      if (
        j === 0 &&
        (typeof value === 'number' ||
          (typeof value === 'string' && /^-?\d+(\.\d+)?$/.test(String(value).trim())))
      ) {
        cell.value = Number(value);
      } else {
        cell.value = String(value);
      }
      cell.alignment = {
        ...(cell.alignment || {}),
        wrapText: true,
        vertical: 'top',
        horizontal: j === 0 ? 'center' : 'left',
      };
    }
  }

  applyFormXIIIColumnLayout(worksheet, startCol, columnHeaderRow, dataStartRow);

  const out = await workbook.xlsx.writeBuffer();
  const fileName =
    formFileName ||
    parsedFormHeader?.title?.replace(/[^a-zA-Z0-9]/g, '_') ||
    `Form_XIII_${Date.now()}.xlsx`;
  const blob = new Blob([out], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  return { blob, fileName };
}
