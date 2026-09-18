import {
  blobIndicatesRegisterOfDeductionsForDamage,
  resolveFormXXAPDeductionsTableLayout,
} from './formXAPRegisterOfFines';
import {
  enrichFormCGJGujaratDisplayHeader,
  headersIndicateFormCGJDeductionsRegisterTable,
  isFormCGJFirstMonthYearHeader,
  isFormCGJLastMonthYearHeader,
  isFormCGJNameHeader,
  isFormCGJRecoveryTypeHeader,
  normCGJHeader,
  readFormCGJEmployeeFullName,
} from './formCGJGujarat';
import {
  enrichEstablishmentPrincipalEmployerHeaderFields,
  FORM_RJ_CONTRACTOR_HEADER_KEYS,
} from '../../utils/statutorySiteCompanyHeaders';

/** Rajasthan Form C — Register of Deductions for Damage or Loss (Contract Labour). */

export const FORM_C_RJ_DEDUCTION_NIL_TEXT = 'NIL';

export function isFormCRajasthanContext(
  formHeader,
  rowItem,
  fileName,
  sheetText = '',
  tableHeaders = []
) {
  const parts = [
    rowItem?.formName,
    rowItem?.FormName,
    rowItem?.formFileName,
    rowItem?.FormFileName,
    rowItem?.siteState,
    rowItem?.SiteState,
    rowItem?.state,
    rowItem?.State,
    fileName,
    formHeader?.title,
    formHeader?.subtitle,
    formHeader?.reference,
    sheetText,
    Array.isArray(tableHeaders) ? tableHeaders.join(' ') : '',
  ]
    .filter((x) => x != null && String(x).trim() !== '')
    .join(' ')
    .toLowerCase();

  if (/gujarat|\b_gj\b|form[\s._-]*c[\s._-]*gj|form_c_gj/.test(parts)) return false;

  // Form_C_MH / Form_C_RJ — underscore after C breaks \bform…c\b.
  if (/form[\s._-]*c[\s._-]*(rj|mh)\b|form_c_(rj|mh)\b/.test(parts)) return true;

  const hasRajasthan =
    /rajasthan|\b_rj\b|form[\s._-]*c[\s._-]*rj|form_c_rj/.test(parts);
  const hasMaharashtra =
    /maharashtra|\b_mh\b|form[\s._-]*c[\s._-]*mh|form_c_mh/.test(parts);
  const hasFormC =
    /form[\s._-]*c(?=[\s._-]|$)/.test(parts) || /\bform[\s._-]*c\b/.test(parts);

  if ((hasRajasthan || hasMaharashtra) && hasFormC) return true;
  if (
    (hasRajasthan || hasMaharashtra) &&
    blobIndicatesRegisterOfDeductionsForDamage(parts, tableHeaders) &&
    !/\bform[\s._-]*a\b/.test(parts)
  ) {
    return true;
  }
  return false;
}

/** Deduction / recovery columns that must default to NIL when no case exists. */
export function isFormCRJDeductionNilHeader(header) {
  if (isFormCGJRecoveryTypeHeader(header)) return true;
  if (isFormCGJFirstMonthYearHeader(header)) return true;
  if (isFormCGJLastMonthYearHeader(header)) return true;

  const s = normCGJHeader(header);
  if (!s) return false;

  if (/^particulars\b/.test(s) || (s.includes('particular') && !/employ/.test(s))) return true;
  if (/date\s+of\s+damage|damage\s*\/\s*loss|damage\s+or\s+loss/.test(s)) return true;
  if (s === 'amount' || /^amount\b/.test(s)) return true;
  if (/show\s+cause|showed\s+cause/.test(s)) return true;
  if (/explanation/.test(s) && (/presence|heard/.test(s) || /person/.test(s))) return true;
  if (/number\s+of\s+instalment|no\.?\s*of\s+instalment|instalments?/.test(s)) return true;
  if (/date\s+of\s+complete\s+recovery|complete\s+recovery/.test(s)) return true;
  if (/recovery\s*type|type\s+of\s+recovery/.test(s)) return true;
  if (/first\s*month/.test(s) && /year/.test(s)) return true;
  if (/last\s*month/.test(s) && /year/.test(s)) return true;

  return false;
}

export function repairFormCRajasthanTableHeadersFromWorkbook(workbook, hints = {}) {
  if (!workbook) return null;
  const layout = resolveFormXXAPDeductionsTableLayout(workbook, {
    preferredSheetName: hints.preferredSheetName || hints.sheetName || '',
    fileName: hints.fileName || '',
    formFileName: hints.formFileName || '',
    formName: hints.formName || '',
    item: hints.item || null,
    formHeader: hints.formHeader || hints.parsedFormHeader || null,
    formHeaderTitle: hints.formHeaderTitle || hints.formHeader?.title || '',
  });
  if (!layout?.headers?.length) return null;
  return layout;
}

function isFormCRajasthanEstablishmentHeaderLabel(label) {
  const s = String(label || '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  if (!s) return false;
  if (/principal|employer|contractor|address\s+of/.test(s)) return false;
  return /^name\s+of\s+(?:the\s+)?establishment/.test(s);
}

export function enrichFormCRajasthanDisplayHeader(formHeader, fileName, item, tableHeaders, sheetText) {
  const base = enrichFormCGJGujaratDisplayHeader(formHeader, fileName, item, tableHeaders, sheetText);
  let fields = enrichEstablishmentPrincipalEmployerHeaderFields(base.fields || []);
  const hasEstablishment = fields.some((field) => isFormCRajasthanEstablishmentHeaderLabel(field?.label));
  if (!hasEstablishment) {
    fields = [
      ...fields,
      { label: 'Name of Establishment:-', value: '', key: 'statutory_establishment_name' },
    ];
  }
  const hasMonthYear = fields.some((field) => /month\s*\/\s*year/i.test(String(field?.label || '')));
  if (!hasMonthYear) {
    fields = [...fields, { label: 'Month/Year:', value: '', key: 'form_c_rj_month_year' }];
  }
  return { ...base, fields };
}

export function prepareFormCRajasthanDownloadHeaderData(
  headerFormData,
  parsedFormHeader = null,
  siteContext = {}
) {
  const out = headerFormData && typeof headerFormData === 'object' ? { ...headerFormData } : {};
  const { establishmentText = '', periodText = '', contractorText = '' } = siteContext;
  if (establishmentText) {
    out.statutory_establishment_name = establishmentText;
    out.form_c_rj_establishment = establishmentText;
  }
  if (periodText) out.form_c_rj_month_year = periodText;
  // Same Site Management contractor line Form C_RJ uses — copy onto every RJ form key.
  if (contractorText) {
    FORM_RJ_CONTRACTOR_HEADER_KEYS.forEach((key) => {
      out[key] = contractorText;
    });
  }
  const fields = Array.isArray(parsedFormHeader?.fields) ? parsedFormHeader.fields : [];
  fields.forEach((field) => {
    const key = field?.key;
    if (!key) return;
    if (!String(out[key] ?? '').trim()) {
      if (isFormCRajasthanEstablishmentHeaderLabel(field.label) && establishmentText) {
        out[key] = establishmentText;
      }
      if (/month\s*\/\s*year/i.test(String(field.label || '')) && periodText) {
        out[key] = periodText;
      }
    }
    if (contractorText && /name\s+and\s+address\s+of\s+(?:the\s+)?contractor/i.test(String(field.label || '')) &&
        !/principal/i.test(String(field.label || ''))) {
      out[key] = contractorText;
    }
  });
  return out;
}

export function headersIndicateFormCRajasthanDeductionsTable(tableHeaders) {
  return headersIndicateFormCGJDeductionsRegisterTable(tableHeaders);
}

export function applyFormCRajasthanNilToRow(row, headers, helpers = {}) {
  const hdrs = Array.isArray(headers) ? headers : [];
  const out = row && typeof row === 'object' ? { ...row } : {};
  const nilText = helpers.nilText != null ? String(helpers.nilText) : FORM_C_RJ_DEDUCTION_NIL_TEXT;
  const { overwrite = true } = helpers;

  hdrs.forEach((header) => {
    if (!isFormCRJDeductionNilHeader(header)) return;
    const existing = String(out[header] ?? '').trim();
    if (!overwrite && existing && !/^enter\b/i.test(existing)) return;
    out[header] = nilText;
  });
  return out;
}

export function applyFormCRajasthanEmployeeToRow(row, emp, headers, helpers = {}) {
  const hdrs = Array.isArray(headers) ? headers : [];
  const out = row && typeof row === 'object' ? { ...row } : {};
  const {
    sanitizeValue = (v) => String(v ?? '').trim(),
    overwrite = true,
    nilText = FORM_C_RJ_DEDUCTION_NIL_TEXT,
  } = helpers;

  const cellIsEmpty = (header) => {
    const v = String(out[header] || '').trim();
    if (!v) return true;
    const s = v.toLowerCase();
    return /^enter\b/.test(s) || s.includes('enter ');
  };
  const setCell = (header, value) => {
    if (!header || value == null || value === '') return;
    if (!overwrite && !cellIsEmpty(header)) return;
    out[header] = sanitizeValue(value);
  };

  const fullName = readFormCGJEmployeeFullName(emp);
  if (fullName) out.__employeeLookupName = fullName;

  hdrs.forEach((header) => {
    if (isFormCRJDeductionNilHeader(header)) {
      setCell(header, nilText);
      return;
    }
    if (isFormCGJNameHeader(header)) {
      setCell(header, fullName);
    }
  });

  return out;
}

export function applyFormCRajasthanNilToMappedRows(mappedData, headers, nilText = FORM_C_RJ_DEDUCTION_NIL_TEXT) {
  if (!Array.isArray(mappedData)) return [];
  return mappedData.map((row) => applyFormCRajasthanNilToRow(row, headers, { nilText, overwrite: true }));
}
