/** Tamil Nadu CLRA Form XXIX — Register of Advances, Deductions for Damage/Loss and Fines [Rule 78(1)(d)]. */

export const FORM_XXIX_TN_TITLE = 'FORM XXIX';
export const FORM_XXIX_TN_SUBTITLE =
  'Register of Advances, Deductions for Damage or Loss and Fines (Contract Labour)';
export const FORM_XXIX_TN_REFERENCE =
  '[See Rule 78 (1) (d) of the Tamil Nadu Contract Labour (Regulation and Abolition) Rules ,1975]';

/**
 * Official Form XXIX leaf columns (1)–(20).
 * Bands: Advance Paid (6–9), Deduction for Damage/Loss (10–14), Fines (15–18).
 */
export const FORM_XXIX_TN_TABLE_HEADERS = [
  'Serial Number',
  'Name of the Workman',
  "Father/Husband's Name",
  'Employee Number',
  'Designation',
  'Date of Payment',
  'Amount Paid',
  'Number of Installments to be recovered',
  'Date on which recovery completed',
  'Date on which caused Damage or Loss',
  'Date of show cause Notice',
  'Total amount of deduction imposed',
  'Number of installments to be recovered',
  'Date on which deduction completed',
  'Act or Omission',
  'Date of Show Cause Notice',
  'Amount of fine imposed',
  'Date on which fine recovery completed',
  'Signature or Thumb Impression of the Workman',
  'Remarks',
];

export const FORM_XXIX_TN_GROUP_ADVANCE = 'Advance Paid';
export const FORM_XXIX_TN_GROUP_DAMAGE = 'Deduction for Damage/Loss';
export const FORM_XXIX_TN_GROUP_FINES = 'Fines';

/** Parent band labels aligned to FORM_XXIX_TN_TABLE_HEADERS indices. */
export function buildFormXXIXTamilNaduColumnGroupLabels(headerCount = FORM_XXIX_TN_TABLE_HEADERS.length) {
  const n = Math.max(0, Number(headerCount) || 0);
  const labels = Array.from({ length: n }, () => '');
  for (let i = 0; i < n; i += 1) {
    if (i >= 5 && i <= 8) labels[i] = FORM_XXIX_TN_GROUP_ADVANCE;
    else if (i >= 9 && i <= 13) labels[i] = FORM_XXIX_TN_GROUP_DAMAGE;
    else if (i >= 14 && i <= 17) labels[i] = FORM_XXIX_TN_GROUP_FINES;
  }
  return labels;
}

const isGenericOrEmptyLeafHeader = (header) => {
  const raw = String(header || '').trim();
  if (!raw) return true;
  const s = formXXIXTamilNaduHeaderNorm(raw).replace(/:+$/, '').trim();
  if (!s) return true;
  if (/^column\s*\d+$/i.test(raw)) return true;
  if (/^\(?\d{1,2}\)?$/.test(s)) return true;
  if (/^month$/.test(s) || /^year$/.test(s) || /^month\s*\/\s*year$/.test(s)) return true;
  // Parent band titles wrongly used as leaf headers
  if (/^advance\s+paid(\s*\(\d+\))?$/.test(s)) return true;
  if (/^deduction\s+for\s+damage/.test(s) && !/date|amount|notice|instal|install|completed|caused/.test(s)) {
    return true;
  }
  if (/^fines?$/.test(s)) return true;
  if (/^loss(\s*\(\d+\))?$/.test(s)) return true;
  return false;
};

export function looksLikeFormXXIXTamilNaduTableHeaders(headers) {
  const list = Array.isArray(headers) ? headers : [];
  if (list.length < 6) return false;
  const joined = list.map((h) => formXXIXTamilNaduHeaderNorm(h)).join(' | ');
  // Form XVIII wages-cum-muster / wage register — not Register of Advances.
  if (
    /net\s+amount\s+paid|daily\s+attendance|total\s+attendance|units\s+of\s+work|amount\s+of\s+wages\s+earned|dearness\s+allowance|daily\s+rate\s+of\s+wages|other\s+cash\s+payments|initial\s+of\s+contractor/.test(
      joined
    )
  ) {
    return false;
  }
  const hasIdentity =
    /name of (?:the )?workm[ae]n|father|husband|employee\s+number|designation/.test(joined);
  // "Net Amount Paid" must not count as advance "Amount Paid".
  const hasAdvanceBand =
    /advance\s+paid|(?:^|\|)\s*amount\s+paid\s*(?:\||$)|installment|instalment|date\s+of\s+payment|recovery\s+completed/.test(
      joined
    );
  const hasDamageOrFinesOrSignature =
    /damage|show\s*cause|fine|act\s+or\s+omission|signature|thumb|column\s*1[5-8]/.test(joined) ||
    (/deduction/.test(joined) && /damage|loss|fine/.test(joined));
  return hasIdentity && hasAdvanceBand && (hasDamageOrFinesOrSignature || list.length >= 12);
}

/**
 * Align leaf headers to the official Form XXIX (1)–(20) model.
 * Replaces blank / "Column N" / parent-band leftovers with canonical titles.
 */
export function resolveFormXXIXTamilNaduTableHeaders(headers) {
  const canonical = FORM_XXIX_TN_TABLE_HEADERS;
  const incoming = stripFormXXIXTamilNaduMonthYearFromTableHeaders(
    Array.isArray(headers) ? headers.map((h) => String(h || '').trim()) : []
  );
  if (!incoming.length) return [...canonical];

  const out = [];
  const max = Math.max(canonical.length, Math.min(incoming.length, canonical.length + 2));
  for (let i = 0; i < max && i < canonical.length; i += 1) {
    const cur = incoming[i] || '';
    if (isGenericOrEmptyLeafHeader(cur)) {
      out.push(canonical[i]);
      continue;
    }
    // Keep a meaningful template label when it already describes the leaf.
    out.push(cur);
  }
  // Ensure trailing Signature / Remarks exist when template truncated early.
  while (out.length < canonical.length) {
    out.push(canonical[out.length]);
  }
  return out.slice(0, canonical.length);
}

/** True when a group-band cell is Month:/Year: (must not appear in table thead). */
export function isFormXXIXTamilNaduMonthYearGroupLabel(label) {
  const s = formXXIXTamilNaduHeaderNorm(label).replace(/:+$/, '').trim();
  return /^month$/.test(s) || /^year$/.test(s) || /^month\s*\/\s*year$/.test(s);
}

/**
 * Replace YEAR:/Month: (and empty) group labels with the official Advance/Damage/Fines bands.
 */
export function sanitizeFormXXIXTamilNaduColumnGroupLabels(groupLabels, headerCount) {
  const n =
    Number(headerCount) > 0
      ? Number(headerCount)
      : Array.isArray(groupLabels)
        ? groupLabels.length
        : FORM_XXIX_TN_TABLE_HEADERS.length;
  const canonical = buildFormXXIXTamilNaduColumnGroupLabels(n);
  if (!Array.isArray(groupLabels) || groupLabels.length === 0) return canonical;

  const out = canonical.map((fallback, i) => {
    const raw = String(groupLabels[i] ?? '').trim();
    if (!raw || isFormXXIXTamilNaduMonthYearGroupLabel(raw)) return fallback;
    const s = formXXIXTamilNaduHeaderNorm(raw);
    if (/advance/.test(s) && /paid/.test(s)) return FORM_XXIX_TN_GROUP_ADVANCE;
    if (/damage|loss/.test(s) && /deduction|damage/.test(s)) return FORM_XXIX_TN_GROUP_DAMAGE;
    if (/^fines?$/.test(s) || (/\bfine\b/.test(s) && !/damage|deduction|amount/.test(s))) {
      return FORM_XXIX_TN_GROUP_FINES;
    }
    // Identity columns should stay blank in the group row.
    if (i < 5 || i >= 18) return '';
    return fallback || raw;
  });
  return out;
}

export function formXXIXTamilNaduHeaderNorm(txt) {
  return String(txt || '')
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

export function matchesFormXXIXHint(blob) {
  const parts = String(blob || '').toLowerCase();
  return (
    /form[\s._-]*xxix(?![a-z])/i.test(parts) ||
    /form[\s._-]*29(?!\d)/i.test(parts) ||
    /(?:^|[\s._-])xxix(?=[\s._\W-]|$)/i.test(parts)
  );
}

/** Form XXII advances register — must not be treated as Form XXIX (xxii ≠ xxix/xxiii). */
export function matchesFormXXIIHint(blob) {
  const parts = String(blob || '').toLowerCase();
  return (
    /form[\s._-]*xxii(?![a-z])/i.test(parts) ||
    /form[\s._-]*22(?!\d)/i.test(parts) ||
    /(?:^|[\s._-])xxii(?=[\s._\W-]|$)/i.test(parts)
  );
}

function buildFormXXIXContextBlob(formHeader, rowItem, fileName, tableHeaders, sheetText) {
  return [
    rowItem?.formName,
    rowItem?.FormName,
    rowItem?.formFileName,
    rowItem?.FormFileName,
    rowItem?.state,
    rowItem?.State,
    rowItem?.siteState,
    rowItem?.SiteState,
    fileName,
    formHeader?.title,
    formHeader?.subtitle,
    formHeader?.reference,
    sheetText,
    Array.isArray(tableHeaders) ? tableHeaders.join(' ') : tableHeaders,
  ]
    .filter((x) => x != null && String(x).trim() !== '')
    .join(' ')
    .toLowerCase();
}

export function isFormXXIXTamilNaduContext(
  formHeader,
  rowItem,
  fileName,
  tableHeaders = [],
  sheetText = ''
) {
  const parts = buildFormXXIXContextBlob(formHeader, rowItem, fileName, tableHeaders, sheetText);
  const fileBlob = String(fileName || rowItem?.formFileName || rowItem?.FormFileName || '')
    .toLowerCase();
  const hasTN =
    /tamil[\s._-]*nadu|tamilnadu|\(tn\)|\btn\b/.test(parts) ||
    /\bform[\s._-]*xxix[\s._-]*tamil/.test(parts) ||
    /\bxxix[\s._-]*tamil/.test(parts);
  const hasXXIX = matchesFormXXIXHint(parts);
  const hasXXIXInFile = matchesFormXXIXHint(fileBlob);
  const hasXXIIInFile = matchesFormXXIIHint(fileBlob);
  const hasXVIIIInFile =
    /form[\s._-]*xviii(?![a-z])/i.test(fileBlob) || /form[\s._-]*18(?!\d)/i.test(fileBlob);
  const hasAdvancesTitle = /register\s+of\s+advances/i.test(parts);
  const headersLookXxix = looksLikeFormXXIXTamilNaduTableHeaders(tableHeaders);

  // Form_XXII_-_TamilNadu.xlsx is Register of Advances, not Form XXIX.
  // Filename wins when it clearly names XXII and not XXIX (sheet may share advances headers).
  if (hasXXIIInFile && !hasXXIXInFile) return false;

  // Form_XVIII_-_TamilNadu.xlsx is wages-cum-muster — never rewrite as Form XXIX.
  if (hasXVIIIInFile && !hasXXIXInFile) return false;
  if (
    /form[\s._-]*xviii(?![a-z])/i.test(parts) &&
    /wages[\s-]*cum[\s-]*muster|register\s+of\s+wages/i.test(parts) &&
    !hasXXIXInFile
  ) {
    return false;
  }

  // Filename Form_XXIX_-_TamilNadu.xlsx is enough even if item.state is blank.
  if (hasXXIX && (hasTN || headersLookXxix || hasAdvancesTitle)) return true;
  // TN + advances title alone is not enough — Form XXII TN is also Register of Advances.
  if (hasTN && hasXXIX) return true;
  if (hasTN && headersLookXxix && !matchesFormXXIIHint(parts) && !hasXVIIIInFile) return true;
  // Avoid wage-register / Form XVIII confusion when filename is not XXIX.
  if (!hasXXIX && /wages[\s-]*cum[\s-]*muster|register\s+of\s+wages/i.test(parts)) {
    return false;
  }
  return false;
}

/** People EmployeeID (+ common aliases) for Form XXIX Employee Number column. */
export function resolveFormXXIXTamilNaduEmployeeId(emp) {
  if (!emp || typeof emp !== 'object') return '';
  const pick = (...vals) => {
    for (let i = 0; i < vals.length; i += 1) {
      const v = vals[i];
      if (v != null && String(v).trim() !== '') return String(v).trim();
    }
    return '';
  };
  return pick(
    emp.EmployeeID,
    emp['EmployeeID'],
    emp['Employee ID'],
    emp['Employee Id'],
    emp.Employee_ID,
    emp['Employee_ID'],
    emp.employeeId,
    emp.EmployeeId,
    emp.employee_id,
    emp['employee_id'],
    emp.EmpID,
    emp.EmpId,
    emp.empId,
    emp['Emp ID'],
    emp.EmployeeCode,
    emp['Employee Code'],
    emp.employeeCode,
    emp.Employee_Number,
    emp['Employee Number'],
    emp.employee_number,
    emp.Zoho_ID,
    emp['Zoho_ID'],
    emp.ZohoID,
    emp['Role.ID'],
    emp.Role && typeof emp.Role === 'object' ? emp.Role.ID : ''
  );
}

export function isFormXXIXTamilNaduHeaderFieldLayoutFormHeader(formHeader) {
  return !!formHeader?.formXXIXTamilNaduHeaderFieldLayout;
}

/** Ordered header fields — 2-column grid: left-to-right, top-to-bottom. */
export const FORM_XXIX_TN_HEADER_SPECS = [
  {
    key: 'form_xxix_contractor',
    label: '1. Name and Address of Contractor.',
    fieldType: 'textarea',
    match: /^1\.?\s*name\s+and\s+address\s+of\s+contractor|name\s+and\s+address\s+of\s+contractor/i,
  },
  {
    key: 'form_xxix_establishment_contract_carried',
    label: '2. Name and address of establishment in/under which contract is carried on.',
    fieldType: 'textarea',
    match:
      /^2\.?\s*name\s+and\s+address\s+of\s+establishment|establishment[\s\S]*contract\s+is\s+carried\s+on/i,
  },
  {
    key: 'form_xxix_nature_location_work',
    label: '3. Nature and location of work.',
    fieldType: 'textarea',
    match: /^3\.?\s*nature\s+and\s+location\s+of\s+work|nature\s+and\s+location\s+of\s+work/i,
  },
  {
    key: 'form_xxix_principal_employer',
    label: '4. Name and address of Principal Employer.',
    fieldType: 'textarea',
    match: /^4\.?\s*name\s+and\s+address\s+of\s+principal\s+employer|principal\s+employer/i,
  },
  {
    key: 'form_x_month',
    label: 'Month:',
    match: /^month\s*:?\s*$/i,
  },
  {
    key: 'form_x_year',
    label: 'Year:',
    match: /^year\s*:?\s*$/i,
  },
];

const FORM_XXIX_TN_HEADER_KEYS = new Set(FORM_XXIX_TN_HEADER_SPECS.map((s) => s.key));

const GENERIC_SITE_HEADER_KEY_RE =
  /^(statutory_establishment_name|statutory_establishment_address|statutory_establishment_name_shop|form25_establishment)$/;

const isExcludedXxixHeaderField = (field) => {
  const key = String(field?.key || '');
  const label = formXXIXTamilNaduHeaderNorm(field?.label);
  if (GENERIC_SITE_HEADER_KEY_RE.test(key)) return true;
  if (/^name\s+of\s+the\s+establishment/.test(label)) return true;
  if (/^address\s+of\s+the\s+establishment/.test(label)) return true;
  if (/^name\s+of\s+establishment\s*\/\s*shop/.test(label)) return true;
  return false;
};

const isMeaningfulHeaderValue = (value, label) => {
  const raw = String(value || '').trim();
  if (!raw) return false;
  const normalizedRaw = formXXIXTamilNaduHeaderNorm(raw);
  const normalizedLabel = formXXIXTamilNaduHeaderNorm(label);
  if (!normalizedRaw || normalizedRaw === normalizedLabel) return false;
  if (
    normalizedRaw ===
    formXXIXTamilNaduHeaderNorm(`Enter ${String(label || '').replace(/:+$/, '').trim()}`)
  ) {
    return false;
  }
  return true;
};

const labelMatchesSpec = (fieldLabel, specLabel) => {
  const a = formXXIXTamilNaduHeaderNorm(fieldLabel).replace(/^\d+\.\s*/, '');
  const b = formXXIXTamilNaduHeaderNorm(specLabel).replace(/^\d+\.\s*/, '');
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
      spec.match.test(formXXIXTamilNaduHeaderNorm(field.label)) &&
      isMeaningfulHeaderValue(field.value, spec.label)
    ) {
      return String(field.value).trim();
    }
  }
  return '';
};

export function finalizeFormXXIXTamilNaduHeaderFields(existingFields = []) {
  const fields = Array.isArray(existingFields)
    ? existingFields.filter((f) => !isExcludedXxixHeaderField(f))
    : [];
  return FORM_XXIX_TN_HEADER_SPECS.map((spec) => ({
    label: spec.label,
    key: spec.key,
    fieldType: spec.fieldType || 'text',
    value: pickValueForSpec(spec, fields),
  }));
}

export function ensureFormXXIXTamilNaduMonthYearHeaderFields(formHeader) {
  if (!formHeader || typeof formHeader !== 'object') return formHeader;
  const fields = Array.isArray(formHeader.fields) ? [...formHeader.fields] : [];
  const hasMonth = fields.some(
    (f) =>
      f?.key === 'form_x_month' ||
      f?.key === 'form_xxix_month' ||
      /^month\s*:?\s*$/i.test(String(f?.label || '').trim())
  );
  const hasYear = fields.some(
    (f) =>
      f?.key === 'form_x_year' ||
      f?.key === 'form_xxix_year' ||
      /^year\s*:?\s*$/i.test(String(f?.label || '').trim())
  );
  if (!hasMonth) fields.push({ label: 'Month:', value: '', key: 'form_x_month' });
  if (!hasYear) fields.push({ label: 'Year:', value: '', key: 'form_x_year' });
  return { ...formHeader, fields };
}

/** Drop Month:/Year: labels that leaked into the table header band. */
export function stripFormXXIXTamilNaduMonthYearFromTableHeaders(headers) {
  if (!Array.isArray(headers)) return headers;
  return headers.filter((h) => {
    const s = formXXIXTamilNaduHeaderNorm(h).replace(/:+$/, '').trim();
    if (/^month$/.test(s) || /^year$/.test(s)) return false;
    if (/^month\s*\/\s*year$/.test(s)) return false;
    return true;
  });
}

export function enrichFormXXIXTamilNaduDisplayHeader(
  formHeader,
  item = null,
  fileName = '',
  tableHeaders = [],
  sheetText = ''
) {
  if (!isFormXXIXTamilNaduContext(formHeader, item, fileName, tableHeaders, sheetText)) {
    return formHeader;
  }
  const base = formHeader && typeof formHeader === 'object' ? { ...formHeader } : {};
  const existing = Array.isArray(base.fields) ? base.fields : [];
  base.title = FORM_XXIX_TN_TITLE;
  if (!/register\s+of\s+advances/i.test(String(base.subtitle || ''))) {
    base.subtitle = FORM_XXIX_TN_SUBTITLE;
  }
  if (!/\brule\s*78/i.test(String(base.reference || ''))) {
    base.reference = FORM_XXIX_TN_REFERENCE;
  }
  base.formXXIXTamilNaduHeaderFieldLayout = true;
  base.formVIIAPHeaderFieldLayout = false;
  base.formXXVIAPHeaderFieldLayout = false;
  base.formXIXAPHeaderFieldLayout = false;
  base.formXVIIITamilNaduHeaderFieldLayout = false;
  base.fields = finalizeFormXXIXTamilNaduHeaderFields(existing);
  return base;
}

export function applyFormXXIXTamilNaduAutofillFromSite(headerData, context = {}) {
  const out = headerData && typeof headerData === 'object' ? { ...headerData } : {};
  const fill = (key, value) => {
    const text = String(value ?? '').trim();
    if (!text || !key) return;
    const cur = String(out[key] ?? '').trim();
    if (!cur || /^enter\b/i.test(cur)) out[key] = text;
  };
  fill('form_xxix_contractor', context.contractorText || '');
  fill('form_xxix_establishment_contract_carried', context.establishmentText || '');
  fill('form_xxix_nature_location_work', context.natureLocationText || '');
  fill('form_xxix_principal_employer', context.principalEmployerText || '');
  fill('form_x_month', context.monthName || '');
  fill('form_x_year', context.year != null ? String(context.year) : '');
  return out;
}

/**
 * Write Month: / Year: into the Excel template (label cell or adjacent value cell).
 */
export function writeFormXXIXTamilNaduMonthYearToExcelJsWorksheet(worksheet, headerFormData = {}) {
  if (!worksheet) return false;
  const monthVal = String(headerFormData.form_x_month ?? headerFormData.form_xxix_month ?? '').trim();
  const yearVal = String(headerFormData.form_x_year ?? headerFormData.form_xxix_year ?? '').trim();
  if (!monthVal && !yearVal) return false;

  const excelCellToPlainText = (value) => {
    if (value == null) return '';
    if (typeof value === 'string' || typeof value === 'number') return String(value);
    if (typeof value === 'object') {
      if (value.text != null) return String(value.text);
      if (value.result != null) return String(value.result);
      if (Array.isArray(value.richText)) {
        return value.richText.map((p) => p?.text || '').join('');
      }
    }
    return String(value);
  };

  const maxRows = Math.min(40, worksheet.rowCount || 40);
  const maxCols = Math.min(40, worksheet.columnCount || 40);
  let wroteMonth = false;
  let wroteYear = false;

  for (let r = 1; r <= maxRows; r += 1) {
    for (let c = 1; c <= maxCols; c += 1) {
      const raw = excelCellToPlainText(worksheet.getCell(r, c)?.value).replace(/\s+/g, ' ').trim();
      if (!raw) continue;

      if (monthVal && /^month\s*:?\s*$/i.test(raw)) {
        const adj = worksheet.getCell(r, c + 1);
        const adjText = excelCellToPlainText(adj?.value).trim();
        if (!adjText || /^enter\b/i.test(adjText)) {
          adj.value = monthVal;
        } else {
          worksheet.getCell(r, c).value = `Month: ${monthVal}`;
        }
        wroteMonth = true;
        continue;
      }
      if (monthVal && /^month\s*:/i.test(raw) && !/year/i.test(raw)) {
        worksheet.getCell(r, c).value = `Month: ${monthVal}`;
        wroteMonth = true;
        continue;
      }

      if (yearVal && /^year\s*:?\s*$/i.test(raw)) {
        const adj = worksheet.getCell(r, c + 1);
        const adjText = excelCellToPlainText(adj?.value).trim();
        if (!adjText || /^enter\b/i.test(adjText)) {
          adj.value = yearVal;
        } else {
          worksheet.getCell(r, c).value = `Year: ${yearVal}`;
        }
        wroteYear = true;
        continue;
      }
      if (yearVal && /^year\s*:/i.test(raw) && !/month/i.test(raw)) {
        worksheet.getCell(r, c).value = `Year: ${yearVal}`;
        wroteYear = true;
      }
    }
  }

  return wroteMonth || wroteYear;
}

/** Default for Advance / Damage / Fines / Signature leaf columns (no advances/deductions/fines). */
export const FORM_XXIX_TN_NIL = 'NILL';

const normFormXXIXTamilNaduNilHeader = (h) =>
  String(h || '')
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();

/** Employee Number column — filled from People EmployeeID. */
export function isFormXXIXTamilNaduEmployeeNumberHeader(header) {
  const s = normFormXXIXTamilNaduNilHeader(header)
    .replace(/^\(?\d+\)?\s*/, '')
    .replace(/^\d+[\).:\-]\s*/, '')
    .replace(/\s*\(\d+\)\s*$/, '')
    .replace(/\.+$/, '')
    .trim();
  return (
    /^employee\s+number$/.test(s) ||
    /^employee\s+no\.?$/.test(s) ||
    /^employee\s+id$/.test(s) ||
    /^emp\.?\s*id$/.test(s) ||
    s === 'employee number' ||
    s === 'employee id' ||
    s === 'employeeid' ||
    s === 'empid'
  );
}

/**
 * Columns that default to NILL (manual / not fetched from payroll).
 * Excludes identity cols and Date of Payment / Remarks.
 */
export function isFormXXIXTamilNaduNilDefaultHeader(header) {
  const s = normFormXXIXTamilNaduNilHeader(header);
  if (!s) return false;
  const bare = s.replace(/^\(?\d+\)?\s*/, '').replace(/\s*\(\d+\)\s*$/, '').trim();
  if (/^remarks?$/.test(bare)) return false;
  if (/^serial(\s+number)?$|^s\.?\s*no\.?$/.test(bare)) return false;
  if (/^name of (?:the )?workm[ae]n$|^name$/.test(bare)) return false;
  if (/father|husband/.test(bare) && /name/.test(bare)) return false;
  if (/^employee\s+(number|no\.?|id)$/.test(bare)) return false;
  if (/^designation$/.test(bare)) return false;
  if (/^date\s+of\s+payment$/.test(bare)) return false;

  if (/^amount\s+paid$/.test(bare) || (/amount\s+paid/.test(bare) && !/deduction|fine/.test(bare))) {
    return true;
  }
  if (/number\s+of\s+install?ments?\s+to\s+be\s+recovered/.test(bare)) return true;
  if (/date\s+on\s+which\s+recovery\s+completed/.test(bare)) return true;
  if (/caused\s+damage|damage\s+or\s*\/?\s*loss|date\s+on\s+which\s+caused/.test(bare)) return true;
  if (/show\s*cause\s*notice/.test(bare)) return true;
  if (/total\s+amount\s+of\s+deduction/.test(bare)) return true;
  if (/date\s+on\s+which\s+deduction\s+completed/.test(bare)) return true;
  if (/act\s+or\s+omission/.test(bare)) return true;
  if (/amount\s+of\s+fine/.test(bare)) return true;
  if (/fine\s+recovery\s+completed|date\s+on\s+which\s+fine/.test(bare)) return true;
  if (/signature|thumb\s*impress/.test(bare)) return true;
  return false;
}

export function applyFormXXIXTamilNaduNilToRow(row, headers, helpers = {}) {
  const hdrs = Array.isArray(headers) ? headers : [];
  const out = row && typeof row === 'object' ? { ...row } : {};
  const nilText = helpers.nilText != null ? String(helpers.nilText) : FORM_XXIX_TN_NIL;
  const { overwrite = true } = helpers;

  hdrs.forEach((header) => {
    if (!isFormXXIXTamilNaduNilDefaultHeader(header)) return;
    const existing = String(out[header] ?? '').trim();
    if (
      !overwrite &&
      existing &&
      !/^enter\b/i.test(existing) &&
      !/^nil+l?$/i.test(existing) &&
      existing.toLowerCase() !== 'n/a' &&
      existing !== '-' &&
      existing !== '—'
    ) {
      return;
    }
    out[header] = nilText;
  });
  return out;
}

export function applyFormXXIXTamilNaduNilToMappedRows(
  mappedData,
  headers,
  nilText = FORM_XXIX_TN_NIL,
  helpers = {}
) {
  if (!Array.isArray(mappedData)) return [];
  const { overwrite = true } = helpers;
  return mappedData.map((row) =>
    applyFormXXIXTamilNaduNilToRow(row, headers, { nilText, overwrite })
  );
}

export { FORM_XXIX_TN_HEADER_KEYS };
