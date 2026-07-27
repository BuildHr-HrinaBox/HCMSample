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

/** Sr. Number in Employee / Workman / Worker Register ← People EmployeeID. */
export function isFormCGJSrNumberRegisterHeader(h) {
  const s = normCGJHeader(h);
  if (!s) return false;
  if (isFormCGJNameHeader(h)) return false;
  return (
    /sr\.?\s*(number|no).*employee.*register/.test(s) ||
    /sr\.?\s*(number|no).*workm[ae]n.*register/.test(s) ||
    /sr\.?\s*(number|no).*worker.*register/.test(s) ||
    /serial\s*(number|no).*employee.*register/.test(s) ||
    /serial\s*(number|no).*register\s+of\s+workm/.test(s) ||
    (/register/.test(s) &&
      /(sr|serial|sl)\s*(number|no)/.test(s) &&
      /(employee|workm[ae]n|worker)/.test(s))
  );
}

export function isFormCGJRecoveryTypeHeader(h) {
  const s = normCGJHeader(h);
  return (
    /recovery\s*type/.test(s) ||
    (s.includes('recovery') && s.includes('type')) ||
    /type\s+of\s+recovery/.test(s) ||
    /damage\s*\/\s*loss\s*\/\s*fine|advance\s*\/\s*loans\s*\/\s*absence/.test(s)
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

export const FORM_C_GJ_DEDUCTION_NIL_TEXT = 'NIL';

/**
 * Deduction / recovery columns that default to NIL when no case exists.
 * Matches Form_C_GJ template columns (Recovery type … Date of Complete Recovery).
 */
export function isFormCGJDeductionNilHeader(header) {
  if (isFormCGJSrNumberRegisterHeader(header)) return false;
  if (isFormCGJNameHeader(header)) return false;
  if (isFormCGJRecoveryTypeHeader(header)) return true;
  if (isFormCGJFirstMonthYearHeader(header)) return true;
  if (isFormCGJLastMonthYearHeader(header)) return true;

  const s = normCGJHeader(header);
  if (!s) return false;

  if (/^particulars\b/.test(s) || (s.includes('particular') && !/employ/.test(s))) return true;
  if (/date\s+of\s+damage|damage\s*\/\s*loss|damage\s+or\s+loss|loss\s*\/\s*absence/.test(s)) {
    return true;
  }
  if (s === 'amount' || /^amount\b/.test(s)) return true;
  if (/show\s+cause|showed\s+cause|whether\s+show/.test(s)) return true;
  if (/explanation/.test(s) && (/presence|heard/.test(s) || /person/.test(s))) return true;
  if (/number\s+of\s+instalment|no\.?\s*of\s+instalment|instalments?/.test(s)) return true;
  if (/date\s+of\s+complete\s+recovery|complete\s+recovery/.test(s)) return true;
  if (/recovery\s*type|type\s+of\s+recovery/.test(s)) return true;
  if (/first\s*month/.test(s) && /year/.test(s)) return true;
  if (/last\s*month/.test(s) && /year/.test(s)) return true;

  return false;
}

/** @deprecated Prefer isFormCGJDeductionNilHeader — kept for callers that clear/skip these columns. */
export function isFormCGJSkipAutofillHeader(h) {
  return isFormCGJDeductionNilHeader(h);
}

function unwrapFormCGJEmployee(emp) {
  return emp?.Employee || emp?.employee || emp || null;
}

export function readFormCGJEmployeeId(emp) {
  const src = unwrapFormCGJEmployee(emp);
  if (!src || typeof src !== 'object') return '';
  const keys = [
    'EmployeeID',
    'Employee ID',
    'Employee_ID',
    'EmployeeId',
    'employeeId',
    'employee_id',
    'Employee.ID',
    'erecno',
    'Erecno',
    'Zoho_ID',
    'ZohoID',
    'zoho_id',
  ];
  for (let i = 0; i < keys.length; i += 1) {
    const raw = src[keys[i]];
    if (raw == null || raw === '') continue;
    if (typeof raw === 'object') {
      const nested = String(raw.displayValue ?? raw.name ?? raw.Name ?? raw.ID ?? '').trim();
      if (nested) return nested;
      continue;
    }
    const text = String(raw).trim();
    if (text) return text;
  }
  return '';
}

export function applyFormCGJGujaratNilToRow(row, headers, helpers = {}) {
  const hdrs = Array.isArray(headers) ? headers : [];
  const out = row && typeof row === 'object' ? { ...row } : {};
  const nilText = helpers.nilText != null ? String(helpers.nilText) : FORM_C_GJ_DEDUCTION_NIL_TEXT;
  const { overwrite = true } = helpers;

  hdrs.forEach((header) => {
    if (!isFormCGJDeductionNilHeader(header)) return;
    const existing = String(out[header] ?? '').trim();
    if (!overwrite && existing && !/^enter\b/i.test(existing) && !/^nil+$/i.test(existing)) return;
    out[header] = nilText;
  });
  return out;
}

export function applyFormCGJGujaratNilToMappedRows(
  mappedData,
  headers,
  nilText = FORM_C_GJ_DEDUCTION_NIL_TEXT
) {
  if (!Array.isArray(mappedData)) return [];
  return mappedData.map((row) =>
    applyFormCGJGujaratNilToRow(row, headers, { nilText, overwrite: true })
  );
}

export function applyFormCGJGujaratEmployeeToRow(row, emp, headers, helpers = {}) {
  const hdrs = Array.isArray(headers) ? headers : [];
  const out = row && typeof row === 'object' ? { ...row } : {};
  const {
    sanitizeValue = (v) => String(v ?? '').trim(),
    overwrite = true,
    nilText = FORM_C_GJ_DEDUCTION_NIL_TEXT,
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

  const fullName = readFormCGJEmployeeFullName(unwrapFormCGJEmployee(emp) || emp);
  const employeeId = readFormCGJEmployeeId(emp);
  if (fullName) out.__employeeLookupName = fullName;
  if (employeeId) out.__employeeLookupId = employeeId;

  hdrs.forEach((header) => {
    if (isFormCGJDeductionNilHeader(header)) {
      setCell(header, nilText);
      return;
    }
    if (isFormCGJSrNumberRegisterHeader(header)) {
      setCell(header, employeeId);
      return;
    }
    if (isFormCGJNameHeader(header)) {
      setCell(header, fullName);
    }
  });

  return out;
}
