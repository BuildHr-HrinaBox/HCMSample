import * as XLSX from 'xlsx';

/** Tamil Nadu CLRA Form XVIII — Register of Wages-cum-Muster Roll [Rule 78(1)(a)(i)]. */

export const FORM_XVIII_TN_TITLE = 'Form XVIII – Register of Wages-cum-Muster Roll';
export const FORM_XVIII_TN_SUBTITLE = 'Form of Register of Wages-cum-Muster Roll';
export const FORM_XVIII_TN_REFERENCE = '[See rule 78(1)(a)(i)]';

export function formXVIIITamilNaduHeaderNorm(txt) {
  return String(txt || '')
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

export function matchesFormXVIIIHint(blob) {
  const parts = String(blob || '').toLowerCase();
  if (/form[\s._-]*xxviii(?![a-z])/i.test(parts)) return false;
  return (
    /form[\s._-]*xviii(?![a-z])/i.test(parts) ||
    /form[\s._-]*18(?!\d)/i.test(parts) ||
    /(?:^|[\s._-])xviii(?=[\s._\W-]|$)/i.test(parts)
  );
}

export function isFormXVIIITamilNaduContext(partsOrBlob) {
  const p = String(partsOrBlob || '').toLowerCase();
  return /tamil[\s._-]*nadu|tamilnadu|form_xviii[_\s-]*tamil|form[\s._-]*xviii[_\s-]*tamil/.test(p);
}

function buildFormXVIIIContextBlob(formHeader, rowItem, fileName, tableHeaders, sheetText) {
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

export function isRegisterOfWagesCumMusterRollBlob(blob) {
  const parts = String(blob || '').toLowerCase();
  return (
    matchesFormXVIIIHint(parts) ||
    /register[\s._-]*of[\s._-]*wages[\s._-]*cum[\s._-]*muster/i.test(parts) ||
    /register\s+of\s+wages[\s-]*cum[\s-]*muster\s+roll/i.test(parts)
  );
}

/** Tamil Nadu CLRA Form XVIII — wages-cum-muster (distinct from generic/AP Form XVIII workbooks). */
export function isFormXVIIITamilNaduClraContext(
  formHeader,
  rowItem,
  fileName,
  tableHeaders = '',
  sheetText = ''
) {
  const parts = buildFormXVIIIContextBlob(formHeader, rowItem, fileName, tableHeaders, sheetText);
  if (!isRegisterOfWagesCumMusterRollBlob(parts)) return false;
  return isFormXVIIITamilNaduContext(parts);
}

export function isFormXVIIITamilNaduHeaderFieldLayoutFormHeader(formHeader) {
  return !!formHeader?.formXVIIITamilNaduHeaderFieldLayout;
}

/** Ordered header fields for TN Form XVIII — 2-column grid reads left-to-right, top-to-bottom. */
export const FORM_XVIII_TN_HEADER_SPECS = [
  {
    key: 'form_xviii_contractor',
    label: '1. Name and Address of Contractor.',
    fieldType: 'textarea',
    match: /^1\.?\s*name\s+and\s+address\s+of\s+contractor|name\s+and\s+address\s+of\s+contractor/i
  },
  {
    key: 'form_xviii_establishment_contract_carried',
    label: '2. Name and address of establishment in/under which contract is carried on.',
    fieldType: 'textarea',
    match: /^2\.?\s*name\s+and\s+address\s+of\s+establishment|establishment[\s\S]*contract\s+is\s+carried\s+on/i
  },
  {
    key: 'form_xviii_nature_location_work',
    label: '3. Nature and location of work.',
    fieldType: 'textarea',
    match: /^3\.?\s*nature\s+and\s+location\s+of\s+work|nature\s+and\s+location\s+of\s+work/i
  },
  {
    key: 'form_xviii_principal_employer',
    label: '4. Name and address of Principal Employer.',
    fieldType: 'textarea',
    match: /^4\.?\s*name\s+and\s+address\s+of\s+principal\s+employer|principal\s+employer/i
  },
  {
    key: 'form_xviii_wage_period',
    label: 'Wage period : Weekly/Fortnightly',
    match: /wage\s*period\s*:?\s*weekly\s*\/\s*fortnightly|wage\s*period/i
  },
  {
    key: 'form_xviii_month_year',
    label: 'Month/Year',
    match: /^month\s*\/\s*year$|^month\s+year$/i
  }
];

const FORM_XVIII_TN_HEADER_KEYS = new Set(FORM_XVIII_TN_HEADER_SPECS.map((s) => s.key));

const GENERIC_SITE_HEADER_KEY_RE =
  /^(statutory_establishment_name|statutory_establishment_address|statutory_establishment_name_shop|form25_establishment)$/;

const isExcludedTnHeaderField = (field) => {
  const key = String(field?.key || '');
  const label = formXVIIITamilNaduHeaderNorm(field?.label);
  if (GENERIC_SITE_HEADER_KEY_RE.test(key)) return true;
  if (/^name\s+of\s+the\s+establishment/.test(label)) return true;
  if (/^address\s+of\s+the\s+establishment/.test(label)) return true;
  if (/^name\s+of\s+establishment\s*\/\s*shop/.test(label)) return true;
  return false;
};

const isMeaningfulHeaderValue = (value, label) => {
  const raw = String(value || '').trim();
  if (!raw) return false;
  const normalizedRaw = formXVIIITamilNaduHeaderNorm(raw);
  const normalizedLabel = formXVIIITamilNaduHeaderNorm(label);
  if (!normalizedRaw || normalizedRaw === normalizedLabel) return false;
  if (normalizedRaw === formXVIIITamilNaduHeaderNorm(`Enter ${String(label || '').replace(/:+$/, '').trim()}`)) {
    return false;
  }
  return true;
};

const labelMatchesSpec = (fieldLabel, specLabel) => {
  const a = formXVIIITamilNaduHeaderNorm(fieldLabel).replace(/^\d+\.\s*/, '');
  const b = formXVIIITamilNaduHeaderNorm(specLabel).replace(/^\d+\.\s*/, '');
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
    if (spec.match.test(formXVIIITamilNaduHeaderNorm(field.label)) && isMeaningfulHeaderValue(field.value, spec.label)) {
      return String(field.value).trim();
    }
  }
  return '';
};

export function finalizeFormXVIIITamilNaduHeaderFields(existingFields = []) {
  const fields = Array.isArray(existingFields) ? existingFields.filter((f) => !isExcludedTnHeaderField(f)) : [];
  return FORM_XVIII_TN_HEADER_SPECS.map((spec) => ({
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
    workbook.SheetNames.find((n) => /form[\s._-]*xviii|wages[\s-]*cum[\s-]*muster/i.test(String(n))) ||
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
      const norm = formXVIIITamilNaduHeaderNorm(raw);
      if (!spec.match.test(norm) && !spec.match.test(raw)) continue;
      const value = readValueBesideLabel(getMergedAwareCellText, r, c, effectiveSheetCols);
      if (value) return value;
    }
  }
  return '';
};

export function buildFormXVIIITamilNaduTemplateFields(getMergedAwareCellText, effectiveSheetCols, existingFields = []) {
  return FORM_XVIII_TN_HEADER_SPECS.map((spec) => {
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

export function enrichFormXVIIITamilNaduDisplayHeader(
  formHeader,
  item = null,
  fileName = '',
  tableHeaders = [],
  sheetText = ''
) {
  if (!isFormXVIIITamilNaduClraContext(formHeader, item, fileName, tableHeaders, sheetText)) {
    return formHeader;
  }
  const base = formHeader && typeof formHeader === 'object' ? { ...formHeader } : {};
  const existing = Array.isArray(base.fields) ? base.fields : [];
  base.title = FORM_XVIII_TN_TITLE;
  if (!/wages[\s-]*cum[\s-]*muster/i.test(String(base.subtitle || ''))) {
    base.subtitle = FORM_XVIII_TN_SUBTITLE;
  }
  if (!/\brule\s*78/i.test(String(base.reference || ''))) {
    base.reference = base.reference || FORM_XVIII_TN_REFERENCE;
  }
  base.formXVIIITamilNaduHeaderFieldLayout = true;
  base.fields = finalizeFormXVIIITamilNaduHeaderFields(existing);
  return base;
}

export function resolveFormXVIIITamilNaduHeaderFieldLayout(parsed, workbook, hints = {}) {
  const formHeader = hints.formHeader || parsed?.formHeader || null;
  const item = hints.item || null;
  const fileName = hints.fileName || hints.formFileName || '';
  const sheetText = hints.sheetText || '';
  const tableHeaders = hints.tableHeaders || parsed?.headers || [];
  if (!isFormXVIIITamilNaduClraContext(formHeader, item, fileName, tableHeaders, sheetText)) {
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
  const finalFields = buildFormXVIIITamilNaduTemplateFields(
    getMergedAwareCellText,
    effectiveSheetCols,
    existingFields
  );

  return {
    formHeader: {
      ...formHeader,
      title: FORM_XVIII_TN_TITLE,
      subtitle: /wages[\s-]*cum[\s-]*muster/i.test(String(formHeader?.subtitle || ''))
        ? formHeader.subtitle
        : FORM_XVIII_TN_SUBTITLE,
      reference: formHeader?.reference || FORM_XVIII_TN_REFERENCE,
      formXVIIITamilNaduHeaderFieldLayout: true,
      fields: finalFields
    }
  };
}

export function applyFormXVIIITamilNaduAutofillFromSite(headerData, context = {}) {
  const out = headerData && typeof headerData === 'object' ? { ...headerData } : {};
  const fill = (key, value) => {
    const text = String(value ?? '').trim();
    if (!text || !key) return;
    const cur = String(out[key] ?? '').trim();
    if (!cur || /^enter\b/i.test(cur)) out[key] = text;
  };
  fill('form_xviii_contractor', context.contractorText || '');
  fill('form_xviii_establishment_contract_carried', context.establishmentText || '');
  fill('form_xviii_nature_location_work', context.natureLocationText || '');
  fill('form_xviii_principal_employer', context.principalEmployerText || '');
  fill('form_xviii_month_year', context.monthYearText || '');
  return out;
}

export { FORM_XVIII_TN_HEADER_KEYS };
