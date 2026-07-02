import {
  blobIndicatesRegisterOfDeductionsForDamage,
  isFormXXAPRegisterOfDeductionsContext,
  resolveFormXXAPDeductionsTableLayout,
} from './formXAPRegisterOfFines';

/** Gujarat Form C — Register of Deductions for Damage or Loss (CLRA-style). */

export function formCGJGujaratHeaderNorm(txt) {
  return String(txt || '')
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

export function isFormCGJGujaratContext(
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

  const hasGujarat = /gujarat|\b_gj\b|form[\s._-]*c[\s._-]*gj|form_c_gj/.test(parts);
  const hasFormC = /\bform[\s._-]*c\b/.test(parts);

  if (hasGujarat && hasFormC) return true;
  if (hasFormC && blobIndicatesRegisterOfDeductionsForDamage(parts, tableHeaders)) return true;
  if (
    hasGujarat &&
    blobIndicatesRegisterOfDeductionsForDamage(parts, tableHeaders) &&
    !/\bform[\s._-]*a\b/.test(parts) &&
    !/\bform[\s._-]*b\b/.test(parts)
  ) {
    return true;
  }
  return false;
}

/** Employee master register (Form A) — must not be used for Form C deductions table. */
export function headersIndicateFormAGJEmployeeRegisterTable(tableHeaders) {
  if (!Array.isArray(tableHeaders) || tableHeaders.length < 6) return false;
  const joined = tableHeaders.map((h) => formCGJGujaratHeaderNorm(h)).join('\n');
  const hasEmployeeCode =
    /employee/.test(joined) &&
    (/workme|workman|worker|code/.test(joined) || /workmen/.test(joined));
  const hasSurname = /surname/.test(joined);
  const hasGender = /\bgender\b|^sex\b/.test(joined);
  const hasDob = /date/.test(joined) && (/birth|bi\b|dob/.test(joined));
  const hasNationality = /national/.test(joined);
  const hasEducation = /education/.test(joined);
  return hasEmployeeCode && hasSurname && (hasGender || hasDob || hasNationality || hasEducation);
}

export function headersIndicateFormCGJDeductionsRegisterTable(tableHeaders) {
  if (!Array.isArray(tableHeaders) || tableHeaders.length < 3) return false;
  const joined = tableHeaders.map((h) => formCGJGujaratHeaderNorm(h)).join('\n');
  if (/particulars\s+of\s+damage|date\s+of\s+damage|amount\s+of\s+deduction/.test(joined)) {
    return true;
  }
  if (/recovery\s*type|first\s*month.*year|last\s*month.*year/.test(joined)) {
    return true;
  }
  if (/deductions?\s+for\s+damage|damage\s+or\s+loss/.test(joined)) return true;
  if (
    /name\s+of\s+(?:the\s+)?workmen|name\s+of\s+workman/.test(joined) &&
    (/father|husband/.test(joined) || /nature\s+of\s+employ/.test(joined)) &&
    !/surname/.test(joined) &&
    !/gender/.test(joined)
  ) {
    return true;
  }
  return false;
}

export function isFormCGJGujaratUsingWrongEmployeeRegisterTable(tableHeaders) {
  return (
    headersIndicateFormAGJEmployeeRegisterTable(tableHeaders) &&
    !headersIndicateFormCGJDeductionsRegisterTable(tableHeaders)
  );
}

export function isFormCGJGujaratDeductionsContext(
  formHeader,
  rowItem,
  fileName,
  sheetText = '',
  tableHeaders = []
) {
  return (
    isFormCGJGujaratContext(formHeader, rowItem, fileName, sheetText, tableHeaders) ||
    (isFormXXAPRegisterOfDeductionsContext(
      formHeader,
      rowItem,
      fileName,
      sheetText,
      tableHeaders
    ) &&
      /gujarat|\b_gj\b|form[\s._-]*c[\s._-]*gj|form_c_gj/.test(
        [
          rowItem?.formName,
          rowItem?.FormName,
          fileName,
          formHeader?.title,
          sheetText,
        ]
          .filter(Boolean)
          .join(' ')
          .toLowerCase()
      ))
  );
}

export function repairFormCGJGujaratTableHeadersFromWorkbook(workbook, hints = {}) {
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

export function enrichFormCGJGujaratDisplayHeader(formHeader, fileName, item, tableHeaders, sheetText) {
  const base = formHeader && typeof formHeader === 'object' ? { ...formHeader } : {};
  const parts = [item?.formName, item?.FormName, fileName, base.title, base.subtitle, sheetText]
    .filter(Boolean)
    .join(' ')
    .toLowerCase();
  const isFormC =
    isFormCGJGujaratContext(base, item, fileName, sheetText, tableHeaders) ||
    /\bform[\s._-]*c\b/.test(parts);

  if (!isFormC) return base;

  return {
    ...base,
    formVIIAPHeaderFieldLayout: false,
    formXXVIAPHeaderFieldLayout: false,
    formXIXAPHeaderFieldLayout: false,
    title:
      /\bform[\s._-]*c\b/i.test(String(base.title || '')) ||
      /form[\s._-]*c[\s._-]*gj/i.test(parts)
        ? base.title || 'Form C – Register of Deductions'
        : 'Form C – Register of Deductions',
    subtitle:
      /register\s+of\s+deductions|damage\s+or\s+loss/i.test(String(base.subtitle || ''))
        ? base.subtitle
        : 'Register of Deductions for Damage or Loss (Contract Labour)',
  };
}

const stripLeadingNumber = (s) =>
  String(s || '')
    .replace(/^\(?\d+\)?[\.\)]\s*/, '')
    .trim();

export function parseFormCGJHeaderLabel(h) {
  const raw = stripLeadingNumber(h);
  const m = String(raw || '').match(/^(.+)_([\s\S]+)$/);
  if (m) return String(m[2] || '').trim();
  return raw;
}

export function normCGJHeader(h) {
  return formCGJGujaratHeaderNorm(parseFormCGJHeaderLabel(h))
    .replace(/[^a-z0-9 ]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

export function readFormCGJEmployeeFullName(emp) {
  const fn = String(
    emp?.FirstName || emp?.['FirstName'] || emp?.firstName || emp?.['First Name'] || ''
  ).trim();
  const ln = String(
    emp?.LastName || emp?.['LastName'] || emp?.lastName || emp?.['Last Name'] || ''
  ).trim();
  if (fn && ln) return `${fn} ${ln}`;
  return fn || ln || '';
}

export function isFormCGJNameHeader(h) {
  const s = normCGJHeader(h);
  if (!/\bname\b/.test(s)) return false;
  if (/father|spouse|husband|employer|establishment|person|presence|explanation|bank/.test(s)) {
    return false;
  }
  return (
    s === 'name' ||
    /^name\b/.test(s) ||
    /\bname\s+of\s+(?:the\s+)?(?:workmen|workman|employee|worker)\b/.test(s)
  );
}

export function isFormCGJRecoveryTypeHeader(h) {
  const s = normCGJHeader(h);
  return (
    /recovery\s*type/.test(s) ||
    (s.includes('recovery') && s.includes('type')) ||
    /type\s+of\s+recovery/.test(s)
  );
}

export function isFormCGJFirstMonthYearHeader(h) {
  const s = normCGJHeader(h);
  return (
    /first\s*month\s*\/?\s*year/.test(s) ||
    (s.includes('first') && s.includes('month') && s.includes('year'))
  );
}

export function isFormCGJLastMonthYearHeader(h) {
  const s = normCGJHeader(h);
  return (
    /last\s*month\s*\/?\s*year/.test(s) ||
    (s.includes('last') && s.includes('month') && s.includes('year'))
  );
}

export function isFormCGJSkipAutofillHeader(h) {
  return (
    isFormCGJRecoveryTypeHeader(h) ||
    isFormCGJFirstMonthYearHeader(h) ||
    isFormCGJLastMonthYearHeader(h)
  );
}

export function applyFormCGJGujaratEmployeeToRow(row, emp, headers, helpers = {}) {
  const hdrs = Array.isArray(headers) ? headers : [];
  const out = row && typeof row === 'object' ? { ...row } : {};
  const {
    sanitizeValue = (v) => String(v ?? '').trim(),
    overwrite = true,
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
    if (isFormCGJSkipAutofillHeader(header)) {
      setCell(header, '');
      return;
    }
    if (isFormCGJNameHeader(header)) {
      setCell(header, fullName);
    }
  });

  return out;
}
