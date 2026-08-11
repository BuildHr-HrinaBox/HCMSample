/** Tamil Nadu CLRA Form XXIX — Register of Advances, Deductions for Damage/Loss and Fines [Rule 78(1)(d)]. */

export const FORM_XXIX_TN_TITLE = 'FORM XXIX';
export const FORM_XXIX_TN_SUBTITLE =
  'Register of Advances, Deductions for Damage or Loss and Fines (Contract Labour)';
/** Exact template heading line (preserves template spelling REGSITER / DAME). */
export const FORM_XXIX_TN_REGISTER_HEADING_LINE =
  'REGSITER OF ADVANCES, DEDUCTIONS FOR DAME OR LOSS AND FINES';
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
  // Form I TN Register of Subsistence Allowance — father + "date of payment" + signature
  // must not look like Form XXIX advances/damage/fines bands.
  if (
    /subsistence\s+allowance|date\s+of\s+suspension|kept\s+under\s+suspension|revocation\s+of\s+suspension|employees\s+placed\s+under\s+suspension/.test(
      joined
    )
  ) {
    return false;
  }
  const hasIdentity =
    /name of (?:the )?workm[ae]n|father|husband|employee\s+number|designation/.test(joined);
  // "Net Amount Paid" / subsistence "…paid and the date of payment" must not count as advance band.
  // Require leaf-level "Amount Paid" / "Date of Payment", or advance/instalment/recovery wording.
  const hasAdvanceBand =
    /advance\s+paid|(?:^|\|)\s*amount\s+paid\s*(?:\||$)|installment|instalment|(?:^|\|)\s*date\s+of\s+payment\s*(?:\||$)|recovery\s+completed/.test(
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

/** Form I / Form 1 TN Register of Subsistence Allowance — never Form XXIX. */
export function matchesFormITamilNaduSubsistenceHint(blob) {
  const parts = String(blob || '').toLowerCase();
  if (!parts) return false;
  if (
    /subsistence\s+allowance|date\s+of\s+suspension|kept\s+under\s+suspension|revocation\s+of\s+suspension|employees\s+placed\s+under\s+suspension|sa\s*form\s*1/.test(
      parts
    )
  ) {
    return true;
  }
  // Form_1_TN_-_TamilNadu.xlsx / Form_I_-_TamilNadu.xlsx (not Form XXIX).
  const looksLikeFormITnFile =
    /form[\s._-]*1[\s._-]*tn\b/i.test(parts) ||
    /form[\s._-]*i[\s._-]*tn\b/i.test(parts) ||
    (/form[\s._-]*i(?![a-z])/i.test(parts) && /tamil[\s._-]*nadu|tamilnadu/i.test(parts)) ||
    (/form[\s._-]*1(?!\d)/i.test(parts) && /tamil[\s._-]*nadu|tamilnadu/i.test(parts));
  return looksLikeFormITnFile && !matchesFormXXIXHint(parts);
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
  // Form_1_TN_-_TamilNadu.xlsx is Register of Subsistence Allowance — never rewrite as XXIX.
  if (matchesFormITamilNaduSubsistenceHint(parts) || matchesFormITamilNaduSubsistenceHint(fileBlob)) {
    return false;
  }
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

function formXXIXExcelJsCellText(val) {
  if (val == null) return '';
  if (typeof val === 'string' || typeof val === 'number' || typeof val === 'boolean') return String(val);
  if (val instanceof Date) return val.toISOString();
  if (typeof val === 'object') {
    if (Array.isArray(val.richText)) return val.richText.map((rt) => rt?.text || '').join('');
    if (val.text != null) return String(val.text);
    if (val.result != null) return String(val.result);
  }
  return String(val);
}

function parseFormXXIXMergeLabel(label) {
  const m = String(label || '')
    .trim()
    .match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/i);
  if (!m) return null;
  const colToNum = (col) => {
    let n = 0;
    const s = String(col || '').toUpperCase();
    for (let i = 0; i < s.length; i += 1) n = n * 26 + (s.charCodeAt(i) - 64);
    return n;
  };
  return {
    r1: Number(m[2]),
    c1: colToNum(m[1]),
    r2: Number(m[4]),
    c2: colToNum(m[3]),
  };
}

const looksLikeFormXXIXTitle = (t) => /^form\s*[-–]?\s*xxix\b/i.test(String(t || '').trim());
const looksLikeFormXXIXReference = (t) =>
  /see\s*rule\s*78|rule\s*78\s*\(\s*1\s*\)\s*\(\s*d\s*\)|contract\s+labour\s*\(regulation/i.test(
    String(t || '').trim()
  );
const looksLikeFormXXIXSubtitle = (t) => {
  const s = String(t || '').replace(/\s+/g, ' ').trim();
  if (!s || looksLikeFormXXIXTitle(s) || looksLikeFormXXIXReference(s)) return false;
  return (
    /regist[er]+ of advances/i.test(s) ||
    /regsiter\s+of\s+advances/i.test(s) ||
    (/advances/i.test(s) && /deduction|damage|dame|fines?/i.test(s)) ||
    /^regsiter\s+of\b/i.test(s) ||
    /^deductions?\s+for\s+dame/i.test(s) ||
    /dame\s+or\s+loss\s+and\s+fines?/i.test(s)
  );
};

const toSingleLineFormXXIXHeading = (txt) =>
  String(txt || '')
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

/**
 * Prefer the official template single-line register heading when the sheet already
 * has REGSITER / DAME wording (or fragmented wrap pieces of it).
 */
export function resolveFormXXIXTamilNaduRegisterHeadingLine(existingText = '') {
  const flat = toSingleLineFormXXIXHeading(existingText);
  if (/regsiter|dame\s+or\s+loss/i.test(flat) || /regist[er]+\s+of\s+advances/i.test(flat)) {
    return FORM_XXIX_TN_REGISTER_HEADING_LINE;
  }
  if (flat && looksLikeFormXXIXSubtitle(flat)) return flat;
  return FORM_XXIX_TN_REGISTER_HEADING_LINE;
}

/**
 * Excel download: place FORM XXIX / See Rule / Register description in column I only.
 * Do not merge heading cells (template often parks them right-aligned / merged into column T).
 */
export function ensureFormXXIXTamilNaduTitleLayout(worksheet, options = {}) {
  if (!worksheet) return false;
  // Column I = 9
  const titleCol = Math.max(1, Number(options.titleCol) || 9);
  const scanCols = Math.max(titleCol + 16, Number(options.colTo) || 30, 30);
  const scanTo = Math.max(3, Number(options.scanTo) || 8);

  const findHit = (matcher) => {
    for (let r = 1; r <= scanTo; r += 1) {
      for (let c = 1; c <= scanCols; c += 1) {
        const t = formXXIXExcelJsCellText(worksheet.getCell(r, c)?.value).replace(/\s+/g, ' ').trim();
        if (matcher(t)) return { row: r, col: c, text: t };
      }
    }
    return null;
  };

  const titleHit = findHit(looksLikeFormXXIXTitle);
  const refHit = findHit(looksLikeFormXXIXReference);
  // Join fragmented wrap pieces (e.g. "REGSITER OF" + "ADVANCES, DEDUCTIONS…") into one line.
  const subtitlePieces = [];
  for (let r = 1; r <= scanTo; r += 1) {
    for (let c = 1; c <= scanCols; c += 1) {
      const t = formXXIXExcelJsCellText(worksheet.getCell(r, c)?.value).replace(/\s+/g, ' ').trim();
      if (looksLikeFormXXIXSubtitle(t)) subtitlePieces.push({ row: r, col: c, text: t });
    }
  }
  const subHit = subtitlePieces[0]
    ? {
        row: Math.min(...subtitlePieces.map((p) => p.row)),
        col: subtitlePieces[0].col,
        text: toSingleLineFormXXIXHeading(subtitlePieces.map((p) => p.text).join(' ')),
      }
    : findHit(looksLikeFormXXIXSubtitle);

  // Always force rows 1–3 for this template when any XXIX heading exists (or force=true).
  const force = options.force === true;
  if (!titleHit && !refHit && !subHit && !force) return false;

  const titleText = toSingleLineFormXXIXHeading(titleHit?.text || FORM_XXIX_TN_TITLE);
  const refText = toSingleLineFormXXIXHeading(refHit?.text || FORM_XXIX_TN_REFERENCE);
  const subText = resolveFormXXIXTamilNaduRegisterHeadingLine(subHit?.text || '');

  // Keep headings on fixed template rows 1 / 2 / 3 (never spill Month/Year row).
  const titleRow = 1;
  const refRow = 2;
  const subRow = 3;

  const collectMergeRefsForRow = (row) => {
    const refs = new Set();
    const modelMerges = Array.isArray(worksheet.model?.merges) ? worksheet.model.merges : [];
    modelMerges.forEach((label) => {
      const m = parseFormXXIXMergeLabel(label);
      if (!m) return;
      if (row < m.r1 || row > m.r2) return;
      refs.add(String(label));
    });
    const internal = worksheet._merges || worksheet.merges || {};
    Object.keys(internal).forEach((key) => {
      const merge = internal[key];
      const top = Number(merge?.top ?? merge?.model?.top);
      const left = Number(merge?.left ?? merge?.model?.left);
      const bottom = Number(merge?.bottom ?? merge?.model?.bottom);
      const right = Number(merge?.right ?? merge?.model?.right);
      if (![top, left, bottom, right].every((n) => Number.isFinite(n))) return;
      if (row < top || row > bottom) return;
      refs.add(`${top},${left},${bottom},${right}`);
      // Also try Excel A1-style from key when present.
      if (key && /:/.test(key)) refs.add(key);
    });
    // Common template merges that park titles on the right (ending at col T = 20).
    ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S'].forEach(
      (start) => {
        refs.add(`${start}${row}:T${row}`);
        refs.add(`${start}${row}:U${row}`);
      }
    );
    return [...refs];
  };

  const unmergeRowHard = (row) => {
    collectMergeRefsForRow(row).forEach((ref) => {
      try {
        if (/^\d+,\d+,\d+,\d+$/.test(ref)) {
          const [t, l, b, r] = ref.split(',').map(Number);
          worksheet.unMergeCells(t, l, b, r);
        } else {
          worksheet.unMergeCells(ref);
        }
      } catch (_) {
        /* ignore */
      }
    });
    // Last resort: walk cells and unmerge via master address.
    for (let c = 1; c <= scanCols; c += 1) {
      try {
        const cell = worksheet.getCell(row, c);
        if (!cell?.isMerged) continue;
        const master = cell.master || cell;
        const addr = master.address || `${master.row || row},${master.col || c}`;
        try {
          worksheet.unMergeCells(addr);
        } catch (_) {
          /* ignore */
        }
      } catch (_) {
        /* ignore */
      }
    }
  };

  const clearHeadingCellsInRow = (row) => {
    for (let c = 1; c <= scanCols; c += 1) {
      const cell = worksheet.getCell(row, c);
      const t = formXXIXExcelJsCellText(cell?.value).replace(/\s+/g, ' ').trim();
      if (
        !t ||
        looksLikeFormXXIXTitle(t) ||
        looksLikeFormXXIXReference(t) ||
        looksLikeFormXXIXSubtitle(t)
      ) {
        // Clear every heading copy, including empty slaves after unmerge.
        if (
          looksLikeFormXXIXTitle(t) ||
          looksLikeFormXXIXReference(t) ||
          looksLikeFormXXIXSubtitle(t) ||
          c !== titleCol
        ) {
          if (
            looksLikeFormXXIXTitle(t) ||
            looksLikeFormXXIXReference(t) ||
            looksLikeFormXXIXSubtitle(t)
          ) {
            cell.value = null;
          }
        }
      }
      if (c !== titleCol) {
        const ct = formXXIXExcelJsCellText(cell?.value).replace(/\s+/g, ' ').trim();
        if (looksLikeFormXXIXTitle(ct) || looksLikeFormXXIXReference(ct) || looksLikeFormXXIXSubtitle(ct)) {
          cell.value = null;
        }
      }
      try {
        cell.alignment = {
          ...(cell.alignment || {}),
          horizontal: c === titleCol ? 'center' : 'general',
          textRotation: 0,
        };
      } catch (_) {
        /* ignore */
      }
    }
  };

  const writeTitleCell = (row, text, { bold = true, size } = {}) => {
    if (row < 1 || !text) return;
    unmergeRowHard(row);
    clearHeadingCellsInRow(row);
    // Explicitly wipe column T (and neighbors) so right-side ghosts cannot remain.
    for (let c = titleCol + 1; c <= scanCols; c += 1) {
      const cell = worksheet.getCell(row, c);
      const t = formXXIXExcelJsCellText(cell?.value).replace(/\s+/g, ' ').trim();
      if (
        !t ||
        looksLikeFormXXIXTitle(t) ||
        looksLikeFormXXIXReference(t) ||
        looksLikeFormXXIXSubtitle(t)
      ) {
        if (looksLikeFormXXIXTitle(t) || looksLikeFormXXIXReference(t) || looksLikeFormXXIXSubtitle(t)) {
          cell.value = null;
        }
      }
    }
    // Always clear T specifically.
    try {
      worksheet.getCell(row, 20).value = null;
    } catch (_) {
      /* ignore */
    }
    const cell = worksheet.getCell(row, titleCol);
    cell.value = toSingleLineFormXXIXHeading(text);
    cell.alignment = {
      ...(cell.alignment || {}),
      horizontal: 'center',
      vertical: 'middle',
      // Template shows one continuous line — do not wrap into stacked rows.
      wrapText: false,
      shrinkToFit: false,
      textRotation: 0,
    };
    cell.font = {
      ...(cell.font || {}),
      bold,
      ...(size ? { size } : {}),
    };
    try {
      const excelRow = worksheet.getRow(row);
      if (excelRow && (excelRow.height == null || excelRow.height < 18)) excelRow.height = 22;
    } catch (_) {
      /* ignore */
    }
  };

  // Widen column I so the long register heading can display as one line.
  try {
    const col = worksheet.getColumn(titleCol);
    const current = Number(col?.width) || 10;
    if (col) col.width = Math.max(current, 48);
  } catch (_) {
    /* ignore */
  }

  writeTitleCell(titleRow, titleText, { bold: true, size: 14 });
  writeTitleCell(refRow, refText, { bold: false, size: 10 });
  writeTitleCell(subRow, subText, { bold: true, size: 11 });

  // Clear leftover wrap fragments on rows 3–4 outside the single-line heading cell.
  for (let r = 3; r <= 4; r += 1) {
    for (let c = 1; c <= scanCols; c += 1) {
      if (r === subRow && c === titleCol) continue;
      const cell = worksheet.getCell(r, c);
      const t = formXXIXExcelJsCellText(cell?.value).replace(/\s+/g, ' ').trim();
      if (looksLikeFormXXIXSubtitle(t) || /^regsiter\b/i.test(t) || /^deductions?\s+for\b/i.test(t)) {
        cell.value = null;
      }
    }
  }

  // Final pass: ensure no heading text remains outside column I on rows 1–3.
  [titleRow, refRow, subRow].forEach((row) => {
    for (let c = 1; c <= scanCols; c += 1) {
      if (c === titleCol) continue;
      const cell = worksheet.getCell(row, c);
      const t = formXXIXExcelJsCellText(cell?.value).replace(/\s+/g, ' ').trim();
      if (looksLikeFormXXIXTitle(t) || looksLikeFormXXIXReference(t) || looksLikeFormXXIXSubtitle(t)) {
        cell.value = null;
      }
    }
  });

  // Title box ends at column T (20) — same right edge as the data table (do not use scan colTo).
  const borderColTo =
    Number(options.borderColTo) > 0 ? Math.min(Number(options.borderColTo), 20) : 20;
  const thin = { style: 'thin', color: { argb: 'FF000000' } };
  const clearPast = Math.max(scanCols, Number(options.colTo) || 0, borderColTo + 8);
  [titleRow, refRow, subRow].forEach((row) => {
    for (let c = 1; c <= clearPast; c += 1) {
      try {
        const cell = worksheet.getCell(row, c);
        if (c > borderColTo) {
          // Strip any leftover borders past column T.
          cell.border = {};
          continue;
        }
        const border = { top: thin, bottom: thin };
        if (c === 1) border.left = thin;
        if (c === borderColTo) border.right = thin;
        cell.border = border;
      } catch (_) {
        /* ignore */
      }
    }
  });

  return true;
}

/** Default for Advance / Damage / Fines / Signature leaf columns (no advances/deductions/fines). */
export const FORM_XXIX_TN_NIL = 'NILL';

/** Empty-register notice (Form XXIX / CLRA advances) — period appended separately. */
export const FORM_XXIX_TN_NIL_OF_MONTH_TEXT = 'NIL of the Month';

const FORM_XXIX_TN_MONTH_ABBR = [
  'Jan',
  'Feb',
  'Mar',
  'Apr',
  'May',
  'Jun',
  'Jul',
  'Aug',
  'Sep',
  'Oct',
  'Nov',
  'Dec'
];

/**
 * Build "NIL of the Month Aug 2026" from a month name/abbr + year.
 */
export function buildFormXXIXTamilNaduNilOfMonthLabel(monthNameOrAbbr, year) {
  const raw = String(monthNameOrAbbr || '').trim();
  const y = Number(year);
  if (!raw || !Number.isFinite(y) || y < 1900) return FORM_XXIX_TN_NIL_OF_MONTH_TEXT;
  const idx = FORM_XXIX_TN_MONTH_ABBR.findIndex(
    (m) => m.toLowerCase() === raw.slice(0, 3).toLowerCase()
  );
  const fromFull = [
    'january',
    'february',
    'march',
    'april',
    'may',
    'june',
    'july',
    'august',
    'september',
    'october',
    'november',
    'december'
  ].findIndex((m) => m === raw.toLowerCase());
  const abbr =
    idx >= 0
      ? FORM_XXIX_TN_MONTH_ABBR[idx]
      : fromFull >= 0
        ? FORM_XXIX_TN_MONTH_ABBR[fromFull]
        : raw.slice(0, 3).replace(/^./, (c) => c.toUpperCase());
  if (!abbr) return FORM_XXIX_TN_NIL_OF_MONTH_TEXT;
  return `${FORM_XXIX_TN_NIL_OF_MONTH_TEXT} ${abbr} ${y}`;
}

/** True when a cell is the Form XXIX nil default (NILL or NIL of the Month …). */
export function isFormXXIXTamilNaduNilCellValue(v) {
  const s = String(v ?? '')
    .replace(/\s+/g, ' ')
    .trim();
  if (!s) return false;
  if (/^nil+l?$/i.test(s)) return true;
  return /^nill?\s+of\s+the\s+month\b/i.test(s);
}

/**
 * Prefer Month:/Year: header fields; otherwise fall back to selected period.
 */
export function resolveFormXXIXTamilNaduNilOfMonthText({
  monthName,
  year,
  headerFormData,
  selectedMonthStr,
  item,
  resolvePeriod,
} = {}) {
  const hdr = headerFormData && typeof headerFormData === 'object' ? headerFormData : {};
  const hdrMonth = String(
    monthName ?? hdr.form_x_month ?? hdr.form_xxix_month ?? ''
  ).trim();
  const hdrYearRaw = year ?? hdr.form_x_year ?? hdr.form_xxix_year ?? '';
  const hdrYear = String(hdrYearRaw).trim();
  if (hdrMonth && hdrYear) {
    return buildFormXXIXTamilNaduNilOfMonthLabel(hdrMonth, hdrYear);
  }
  if (typeof resolvePeriod === 'function') {
    const period = resolvePeriod(selectedMonthStr, item) || {};
    return buildFormXXIXTamilNaduNilOfMonthLabel(
      hdrMonth || period.fullMonth || period.monthName || '',
      hdrYear || period.year || ''
    );
  }
  return buildFormXXIXTamilNaduNilOfMonthLabel(hdrMonth, hdrYear);
}

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
      !isFormXXIXTamilNaduNilCellValue(existing) &&
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

/**
 * Single empty-register row: "NIL of the Month May 2026" in the name column (no People rows).
 * Also accepts template typos like "Name of the Wokmen".
 */
export function buildFormXXIXTamilNaduNilTableRows(headers, nilOfMonthText) {
  const hdrs =
    Array.isArray(headers) && headers.length > 0 ? headers : FORM_XXIX_TN_TABLE_HEADERS;
  const row = {};
  hdrs.forEach((h) => {
    row[h] = '';
  });
  const text = String(nilOfMonthText || FORM_XXIX_TN_NIL_OF_MONTH_TEXT).trim();
  let nameHdr = null;
  let serialHdr = null;
  hdrs.forEach((h) => {
    const s = normFormXXIXTamilNaduNilHeader(h)
      .replace(/^\(?\d+\)?\s*/, '')
      .replace(/\s*\(\d+\)\s*$/, '')
      .trim();
    if (
      !nameHdr &&
      (/^name of (?:the )?workm[ae]n$/.test(s) ||
        /^name of (?:the )?wokm[ae]n$/.test(s) ||
        s === 'name')
    ) {
      nameHdr = h;
    }
    if (!serialHdr && (/^serial(\s+number)?$/.test(s) || /^s\.?\s*no\.?$/.test(s) || /^sl\.?\s*no\.?$/.test(s))) {
      serialHdr = h;
    }
  });
  // Prefer name column so SL.NO stays blank until Excel merge rewrites the band.
  const primary = nameHdr || hdrs.find((h) => h !== serialHdr) || serialHdr || hdrs[0];
  if (primary && text) row[primary] = text;
  return [row];
}

/** Detect Form XXIX empty-register nil-of-month cell text. */
export function isFormXXIXTamilNaduNilMonthValue(v) {
  return /^nill?\s+of\s+the\s+month\b/i.test(String(v ?? '').trim());
}

/**
 * Excel download: merge the nil-of-month data row across the full table band (single centered line).
 */
export function mergeFormXXIXTamilNaduNilRowInExcelJsWorksheet(
  worksheet,
  { dataStartRow, sourceRows, colFrom, colTo, nilDisplayText = '' } = {}
) {
  if (!worksheet || !Array.isArray(sourceRows) || sourceRows.length === 0) return false;
  const startCol = Number(colFrom);
  const endCol = Number(colTo);
  if (!Number.isFinite(startCol) || !Number.isFinite(endCol) || endCol < startCol) return false;
  let merged = false;
  for (let i = 0; i < sourceRows.length; i += 1) {
    const row = sourceRows[i] || [];
    const cells = Array.isArray(row)
      ? row.map((v) => String(v ?? '').trim())
      : Object.values(row || {}).map((v) => String(v ?? '').trim());
    const nilFromRow = cells.find((t) => isFormXXIXTamilNaduNilMonthValue(t));
    if (!nilFromRow && !String(nilDisplayText || '').trim()) continue;
    let display = String(nilDisplayText || '').trim() || nilFromRow;
    // Flatten multi-line / wrapped template spill into one line.
    display = display.replace(/\s+/g, ' ').trim();
    if (!isFormXXIXTamilNaduNilMonthValue(display)) {
      display = FORM_XXIX_TN_NIL_OF_MONTH_TEXT;
    }
    const targetRowNum = Number(dataStartRow) + i;
    if (!Number.isFinite(targetRowNum) || targetRowNum < 1) continue;
    for (let c = startCol; c <= endCol; c += 1) {
      try {
        worksheet.unMergeCells(targetRowNum, c, targetRowNum, c);
      } catch (_) {
        /* ignore */
      }
      const cell = worksheet.getCell(targetRowNum, c);
      cell.value = '';
    }
    try {
      worksheet.unMergeCells(targetRowNum, startCol, targetRowNum, endCol);
    } catch (_) {
      /* ignore */
    }
    try {
      worksheet.mergeCells(targetRowNum, startCol, targetRowNum, endCol);
    } catch (_) {
      /* ignore */
    }
    const nilCell = worksheet.getCell(targetRowNum, startCol);
    nilCell.value = display;
    nilCell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: false };
    nilCell.font = { ...(nilCell.font || {}), bold: false };
    try {
      worksheet.getRow(targetRowNum).height = 28;
    } catch (_) {
      /* ignore */
    }
    merged = true;
  }
  return merged;
}

export { FORM_XXIX_TN_HEADER_KEYS };
