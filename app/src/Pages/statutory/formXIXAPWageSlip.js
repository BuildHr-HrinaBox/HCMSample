import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { labelMatchScore } from './form18APAccidentNotice';
import { blobIndicatesRegisterOfDeductionsForDamage } from './formXAPRegisterOfFines';

/** AP CLRA Form XIX — Wage Slip [Rule 78(1)(b)]. */

export function formXIXAPHeaderNorm(txt) {
  return String(txt || '')
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

export function matchesFormXIXHint(blob) {
  const parts = String(blob || '').toLowerCase();
  if (/form[\s._-]*xxix(?![a-z])/i.test(parts)) return false;
  return (
    /form[\s._-]*xix(?![a-z])/i.test(parts) ||
    /form[\s._-]*19(?!\d)/i.test(parts) ||
    /(?:^|[\s._-])xix(?=[\s._\W-]|$)/i.test(parts)
  );
}

/** Rajasthan CLRA Form XIX — Register of Overtime [Rule 77(2)(e)]; not a wage slip. */
export function isFormXIXRajasthanOvertimeRegisterContext(
  formHeader,
  rowItem,
  fileName,
  sheetText = ''
) {
  const parts = [
    rowItem?.formName,
    rowItem?.FormName,
    rowItem?.description,
    rowItem?.Description,
    rowItem?.act,
    rowItem?.Act,
    rowItem?.siteState,
    rowItem?.SiteState,
    rowItem?.state,
    rowItem?.State,
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

  // Form_XIX_RJ / Form_XIX_MH identity wins even if company/site text mentions another state.
  if (/form[\s._-]*xix[\s._-]*rj/i.test(parts) || /\bxix_rj\b/i.test(parts)) return true;
  if (/form[\s._-]*xix[\s._-]*mh/i.test(parts) || /\bxix_mh\b/i.test(parts)) return true;

  if (/wage\s+slip/i.test(parts)) return false;
  // Other states' Form XIX wage-slip files must not match via siteState=Rajasthan alone.
  if (
    /form[\s._-]*xix[\s._-]*(mp|ka|gj|ap|tn|tamil)/i.test(parts) ||
    /madhya\s+pradesh|karnataka|gujarat|andhra\s+pradesh|tamil[\s._-]*nadu/i.test(parts)
  ) {
    return false;
  }

  if (/register\s+of\s+over[\s-]*time|over[\s-]*time\s+register/i.test(parts) && matchesFormXIXHint(parts)) {
    return /rajasthan|maharashtra/i.test(parts) || /form[\s._-]*xix(?![a-z])/i.test(parts);
  }

  return /(rajasthan|maharashtra)/i.test(parts) && matchesFormXIXHint(parts);
}

export function isFormXIXAPWageSlipContext(formHeader, rowItem, fileName, sheetText = '') {
  const parts = [
    rowItem?.formName,
    rowItem?.FormName,
    rowItem?.description,
    rowItem?.Description,
    rowItem?.act,
    rowItem?.Act,
    rowItem?.siteState,
    rowItem?.SiteState,
    rowItem?.state,
    rowItem?.State,
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

  // Filename / catalog identity wins: Form XV (e.g. Form_XV_RJ) must never become Form XIX wage slip,
  // even when the sheet has wage-particular columns that look like an AP Form XIX template.
  const fileIdentity = [rowItem?.formFileName, rowItem?.FormFileName, fileName]
    .filter((x) => x != null && String(x).trim() !== '')
    .join(' ')
    .toLowerCase();
  if (
    /form[\s._-]*xv(?![a-z])/i.test(fileIdentity) ||
    /\bxv[\s._-]*rj\b/i.test(fileIdentity) ||
    /form[\s._-]*xv(?![a-z])/i.test(parts) ||
    /\bxv[\s._-]*rj\b/i.test(parts)
  ) {
    return false;
  }

  // Rajasthan Form XIX is Register of Overtime (Rule 77(2)(e)), not wage slip.
  if (isFormXIXRajasthanOvertimeRegisterContext(formHeader, rowItem, fileName, sheetText)) {
    return false;
  }
  if (/register\s+of\s+over[\s-]*time|over[\s-]*time\s+register/i.test(parts)) {
    return false;
  }

  if (/wage\s+slip/i.test(parts)) return true;
  if (matchesFormXIXHint(parts) && /wage\s+slip|rule\s+78\s*\(\s*1\s*\)\s*\(\s*b\s*\)/i.test(parts)) {
    return true;
  }
  if (
    matchesFormXIXHint(parts) &&
    !/register\s+of\s+(fines|advances|workmen|wages|employment|over[\s-]*time|deductions)/i.test(parts) &&
    !/deductions?\s+for\s+damage|damage\s+or\s+loss/i.test(parts)
  ) {
    const fileIdentity = [rowItem?.formFileName, rowItem?.FormFileName, fileName]
      .filter((x) => x != null && String(x).trim() !== '')
      .join(' ')
      .toLowerCase();
    if (
      /wage\s+slip/i.test(fileIdentity) ||
      /rule\s+78\s*\(\s*1\s*\)\s*\(\s*b\s*\)/i.test(fileIdentity) ||
      /andhra\s+pradesh|form[\s._-]*xix[\s._-]*ap/.test(fileIdentity) ||
      matchesFormXIXHint(fileIdentity)
    ) {
      return true;
    }
  }
  return false;
}

/**
 * Andhra Pradesh Form XIX — contractor/header fields + tabular wage particulars (1–7).
 * Distinct from MP/GJ/TN/KA state-specific wage-slip variants.
 */
export function isFormXIXAPTableWageSlipContext(formHeader, rowItem, fileName, sheetText = '') {
  if (!isFormXIXAPWageSlipContext(formHeader, rowItem, fileName, sheetText)) return false;
  if (formHeader?.formXIXMPTableLayout && !formHeader?.formXIXAPTableLayout) return false;
  if (formHeader?.formXIXKarnatakaTableLayout) return false;
  const parts = [
    rowItem?.formName,
    rowItem?.FormName,
    rowItem?.formFileName,
    rowItem?.FormFileName,
    rowItem?.state,
    rowItem?.State,
    fileName,
    formHeader?.title,
    sheetText,
  ]
    .filter((x) => x != null && String(x).trim() !== '')
    .join(' ')
    .toLowerCase();
  if (/madhya\s+pradesh|form[\s._-]*xix[\s._-]*mp/.test(parts)) return false;
  if (/karnataka|form[\s._-]*xix[\s._-]*(?:ka|karnataka)/.test(parts)) return false;
  if (/gujarat|form[\s._-]*xix[\s._-]*gj/.test(parts)) return false;
  if (/tamil[\s._-]*nadu|form[\s._-]*xix[\s._-]*tamil/.test(parts)) return false;
  return true;
}

/** @deprecated Use isFormXIXAPTableWageSlipContext — AP wage particulars are tabular. */
export function isFormXIXAPHeaderOnlyWageSlipContext(formHeader, rowItem, fileName, sheetText = '') {
  return isFormXIXAPTableWageSlipContext(formHeader, rowItem, fileName, sheetText);
}

export function isFormXIXAPHeaderFieldLayoutFormHeader(formHeader) {
  return !!formHeader?.formXIXAPHeaderFieldLayout;
}

export function isFormXIXAPTableLayoutFormHeader(formHeader) {
  return !!formHeader?.formXIXAPTableLayout;
}

/** Modal/export wage grid — workman + official Form XIX particulars 1–7. */
export const FORM_XIX_AP_WAGE_TABLE_HEADERS = [
  "Name and Father's/Husband's Name of the workman",
  '1. No. of days worked',
  '2. No. of units worked in case of piece-rate Workers',
  '3. Rate of daily wages/piece-rate',
  '4. Amount of overtime wages',
  '5. Gross wages payable',
  '6. Deductions, if any',
  '7. Net amount of wages paid',
];

export function resolveFormXIXAPWageTableHeaders(tableHeaders) {
  const list = Array.isArray(tableHeaders) ? tableHeaders.filter(Boolean) : [];
  if (list.length >= FORM_XIX_AP_WAGE_TABLE_HEADERS.length - 1) return list;
  return [...FORM_XIX_AP_WAGE_TABLE_HEADERS];
}

/** Header/footer only — wage particulars 1–7 render in the employee table. */
export const FORM_XIX_AP_FIELD_GROUPS = [
  { id: 'header', title: 'Wage slip — workman & contractor' },
  { id: 'footer', title: 'Certification' }
];

/**
 * Andhra Pradesh printed Form XIX — fixed value columns beside labels (1-based Excel).
 * Nature/location → D (same row as label), period ending → L, workman name → N.
 */
export const FORM_XIX_AP_HEADER_VALUE_COLS = {
  form_xix_ap_nature_location: 4, // D
  form_xix_ap_period_ending: 12, // L
  form_xix_ap_workman: 14, // N
};

export const FORM_XIX_AP_TEMPLATE_SPECS = [
  {
    key: 'form_xix_ap_contractor',
    label: 'Name and address of contractor:',
    group: 'header',
    fieldType: 'textarea',
    // Gujarat Form XIX wording is "if contractor"; other states use "of contractor".
    match: /name\s+and\s+address\s+(?:of|if)\s+contractor/i
  },
  {
    key: 'form_xix_ap_workman',
    label: "Name and Father's/Husband's Name of the workman:",
    group: 'header',
    fieldType: 'textarea',
    match: /name\s+and\s+father.*husband.*workman|father.*husband.*name\s+of\s+the\s+workman/i
  },
  {
    key: 'form_xix_ap_nature_location',
    label: 'Nature and location of work:',
    group: 'header',
    fieldType: 'textarea',
    match: /nature\s+and\s+location\s+of\s+work/i
  },
  {
    key: 'form_xix_ap_period_ending',
    label: 'For the week/Fortnight/Month ending:',
    group: 'header',
    match: /week.*fortnight.*month\s+ending|fortnight.*month\s+ending/i
  },
  {
    key: 'form_xix_ap_days_worked',
    label: '1. No. of days worked',
    group: 'wages',
    match: /(?:no|number)\.?\s*of\s+days\s+worked|days\s+worked|^\s*1[\.\)]\s+.*days\s+worked/i
  },
  {
    key: 'form_xix_ap_units_worked',
    label: '2. No. of units worked in case of piece-rate Workers',
    group: 'wages',
    match: /no\.?\s*of\s+units\s+worked|piece.?rate\s+workers|^\s*2[\.\)]\s+.*units/i
  },
  {
    key: 'form_xix_ap_rate',
    label: '3. Rate of daily wages/piece-rate',
    group: 'wages',
    match: /rate\s+of\s+(?:daily\s+)?wages|piece[\s-]?rate|daily\s+wages[\s/]*piece|^\s*3[\.\)]\s+.*rate/i
  },
  {
    key: 'form_xix_ap_overtime',
    label: '4. Amount of overtime wages',
    group: 'wages',
    match: /amount\s+of\s+overtime\s+wages|^\s*4[\.\)]\s+.*overtime/i
  },
  {
    key: 'form_xix_ap_gross',
    label: '5. Gross wages payable',
    group: 'wages',
    match: /gross\s+wages\s+payable|^\s*5[\.\)]\s+.*gross/i
  },
  {
    key: 'form_xix_ap_deductions',
    label: '6. Deductions, if any',
    group: 'wages',
    match: /deductions?,?\s*if\s+any|^\s*6[\.\)]\s+.*deduction/i
  },
  {
    key: 'form_xix_ap_net',
    label: '7. Net amount of wages paid',
    group: 'wages',
    match: /net\s+amount\s+of\s+wages\s+paid|^\s*7[\.\)]\s+.*net/i
  },
  {
    key: 'form_xix_ap_initials',
    label: 'Initials of the contractor or his representative',
    group: 'footer',
    match: /initials\s+of\s+the\s+contractor|contractor\s+or\s+his\s+representative/i
  }
];

const isNarrativeBlob = (raw) => {
  const n = formXIXAPHeaderNorm(raw);
  if (!n) return false;
  return (
    /^form\s*xix\b/i.test(n) ||
    /wage\s+slip/i.test(n) ||
    /rule\s+78\s*\(\s*1\s*\)\s*\(\s*b\s*\)/i.test(n) ||
    /contract\s+labou?r/i.test(n) ||
    n.length > 160
  );
};

const isDottedPlaceholderText = (txt) => {
  const t = String(txt || '').trim();
  return t.length > 0 && /^[.\u2026…_\-\s]+$/.test(t);
};

/** True when cell text is another Form XIX heading (not a fill-in value). */
export function isFormXIXAPTemplateLabelText(raw, excludeKey = null) {
  const text = String(raw || '').trim();
  if (!text || isDottedPlaceholderText(text)) return false;
  const stripped = text.replace(/[:.\u2026…_]+$/g, '').replace(/\s+/g, ' ').trim();
  const norm = formXIXAPHeaderNorm(stripped);
  if (!norm) return false;
  return FORM_XIX_AP_TEMPLATE_SPECS.some((spec) => {
    if (excludeKey && spec.key === excludeKey) return false;
    if (spec.match && (spec.match.test(text) || spec.match.test(norm) || spec.match.test(stripped))) {
      return true;
    }
    const labelNorm = formXIXAPHeaderNorm(String(spec.label || '').replace(/:+$/, ''));
    return !!labelNorm && (norm === labelNorm || (norm.length >= 18 && labelNorm.length >= 18 && norm === labelNorm));
  });
}

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
  sectionRow: coords.sectionRow ?? null
});

const readValueBelowOrBesideLabel = (getMergedAwareCellText, labelRow, labelCol, effectiveSheetCols) => {
  const maxC = Math.max(20, effectiveSheetCols || 0);
  let dottedBeside = null;
  for (let c = labelCol + 1; c < Math.min(labelCol + 14, maxC); c += 1) {
    const v = String(getMergedAwareCellText(labelRow, c) || '').trim();
    if (!v || isNarrativeBlob(v) || isFormXIXAPTemplateLabelText(v)) continue;
    if (isDottedPlaceholderText(v)) {
      if (!dottedBeside) dottedBeside = { value: '', valueCol: c, valueRow: labelRow };
      continue;
    }
    return { value: v, valueCol: c, valueRow: labelRow };
  }
  if (dottedBeside) return dottedBeside;
  for (let r = labelRow + 1; r <= Math.min(labelRow + 6, labelRow + 12); r += 1) {
    const v = String(getMergedAwareCellText(r, labelCol) || '').trim();
    if (!v || isNarrativeBlob(v) || isFormXIXAPTemplateLabelText(v)) continue;
    if (isDottedPlaceholderText(v)) {
      return { value: '', valueCol: labelCol, valueRow: r };
    }
    return { value: v, valueCol: labelCol, valueRow: r };
  }
  // Original AP template is two-column: neighboring H-column cells are other headings.
  // Default the value cell under the label so export does not overwrite those headings.
  return { value: '', valueCol: labelCol, valueRow: labelRow + 1 };
};

const findHeaderLabelCell = (getMergedAwareCellText, matchRe, effectiveSheetCols, maxRows = 120) => {
  const maxC = Math.max(20, effectiveSheetCols || 0);
  for (let r = 0; r < maxRows; r += 1) {
    for (let c = 0; c < maxC; c += 1) {
      const raw = String(getMergedAwareCellText(r, c) || '').trim();
      if (!raw || isNarrativeBlob(raw)) continue;
      const norm = formXIXAPHeaderNorm(raw);
      if (!matchRe.test(norm) && !matchRe.test(raw)) continue;
      if (norm.length > 120) continue;
      const read = readValueBelowOrBesideLabel(getMergedAwareCellText, r, c, effectiveSheetCols);
      const value = isNarrativeBlob(read.value) ? '' : read.value;
      return {
        labelRow: r,
        labelCol: c,
        valueCol: read.valueCol,
        valueRow: read.valueRow,
        value
      };
    }
  }
  return null;
};

const findNumberedLabelCell = (getMergedAwareCellText, matchRe, effectiveSheetCols, maxRows = 120) => {
  const maxC = Math.max(20, effectiveSheetCols || 0);
  for (let r = 0; r < maxRows; r += 1) {
    for (let c = 0; c < maxC; c += 1) {
      const raw = String(getMergedAwareCellText(r, c) || '').trim();
      if (!raw || isNarrativeBlob(raw)) continue;
      const norm = formXIXAPHeaderNorm(raw);
      if (!/^\s*\d+[\.\)]/.test(raw) && !matchRe.test(norm)) continue;
      if (!matchRe.test(norm) && !matchRe.test(raw)) continue;
      const read = readValueBelowOrBesideLabel(getMergedAwareCellText, r, c, effectiveSheetCols);
      return {
        labelRow: r,
        labelCol: c,
        valueCol: read.valueCol,
        valueRow: read.valueRow,
        value: isNarrativeBlob(read.value) ? '' : read.value
      };
    }
  }
  return findHeaderLabelCell(getMergedAwareCellText, matchRe, effectiveSheetCols, maxRows);
};

function buildSheetTextBlob(workbook, sheetName) {
  const ws = workbook?.Sheets?.[sheetName];
  if (!ws) return '';
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  return rows
    .slice(0, 35)
    .map((row) => (Array.isArray(row) ? row : []).map((c) => String(c || '')).join(' '))
    .join(' ');
}

function scoreFormXIXSheet(sheetName, sheetText, hintsBlob) {
  const sheetLower = String(sheetName || '').toLowerCase();
  const sheetBlob = `${sheetLower} ${sheetText}`;
  let score = 0;

  if (matchesFormXIXHint(sheetBlob) || /wage\s+slip/i.test(sheetBlob)) score += 200;
  if (/rule\s+78\s*\(\s*1\s*\)\s*\(\s*b\s*\)/i.test(sheetText)) score += 80;
  if (/gross\s+wages\s+payable|net\s+amount\s+of\s+wages/i.test(sheetText)) score += 60;
  if (/no\.?\s*of\s+days\s+worked/i.test(sheetText)) score += 40;
  if (blobIndicatesRegisterOfDeductionsForDamage(sheetBlob)) score -= 220;
  if (/register\s+of\s+deductions/i.test(sheetBlob)) score -= 180;
  if (matchesFormXIXHint(hintsBlob)) score += 30;

  return score;
}

function pickWorkbookSheet(workbook, hints = {}) {
  const names = Array.isArray(workbook?.SheetNames) ? workbook.SheetNames : [];
  if (!names.length) return null;
  const preferred = String(hints.preferredSheetName || '').trim();
  if (preferred && workbook.Sheets?.[preferred]) return preferred;
  const blob = [hints.fileName, hints.formFileName, hints.item?.formName, hints.item?.FormName]
    .filter(Boolean)
    .join(' ')
    .toLowerCase();

  if (names.length > 1) {
    let best = null;
    let bestScore = -Infinity;
    for (const name of names) {
      const sheetText = buildSheetTextBlob(workbook, name);
      const score = scoreFormXIXSheet(name, sheetText, blob);
      if (score > bestScore) {
        bestScore = score;
        best = name;
      }
    }
    if (bestScore > 0 && best) return best;
  }

  let best = null;
  let bestScore = -Infinity;
  for (const name of names) {
    const sheetLower = String(name || '').toLowerCase();
    let score = 0;
    if (matchesFormXIXHint(sheetLower) || /wage\s+slip/i.test(sheetLower)) score += 120;
    if (/register\s+of|appointment|notice\s+of\s+change/i.test(sheetLower)) score -= 80;
    if (score > bestScore) {
      bestScore = score;
      best = name;
    }
  }
  if (bestScore > 0 && best) return best;
  const slipSheet = names.find((n) => /wage\s+slip|form\s*xix/i.test(String(n || '')));
  if (slipSheet) return slipSheet;
  if (matchesFormXIXHint(blob)) {
    const xixSheet = names.find((n) => {
      const sheetText = buildSheetTextBlob(workbook, n);
      return (
        /wage\s+slip/i.test(sheetText) ||
        (!blobIndicatesRegisterOfDeductionsForDamage(`${n} ${sheetText}`) && matchesFormXIXHint(sheetText))
      );
    });
    return xixSheet || names.find((n) => matchesFormXIXHint(String(n || ''))) || names[0];
  }
  return names[0];
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

export function buildFormXIXAPTemplateFields(getMergedAwareCellText = null, effectiveSheetCols = 20) {
  return FORM_XIX_AP_TEMPLATE_SPECS.map((spec) => {
    let coords = {};
    if (typeof getMergedAwareCellText === 'function' && spec.match) {
      coords =
        findNumberedLabelCell(getMergedAwareCellText, spec.match, effectiveSheetCols) ||
        findHeaderLabelCell(getMergedAwareCellText, spec.match, effectiveSheetCols) ||
        {};
    }
    // AP printed slip: pin nature → D, period ending → L, workman → N (0-based valueCol).
    const fixedExcelCol = FORM_XIX_AP_HEADER_VALUE_COLS[spec.key];
    if (fixedExcelCol != null && coords.labelRow != null) {
      coords = {
        ...coords,
        valueCol: fixedExcelCol - 1,
        valueRow: coords.labelRow,
      };
    }
    return buildTemplateField(spec, coords);
  });
}

export function resolveFormXIXAPHeaderFieldLayout(parsed, workbook, hints = {}) {
  const formHeader = hints.formHeader || parsed?.formHeader || null;
  const item = hints.item || null;
  const fileName = hints.fileName || hints.formFileName || '';
  const sheetText = hints.sheetText || '';
  if (!isFormXIXAPWageSlipContext(formHeader, item, fileName, sheetText)) return null;
  // MP/GJ/TN/KA use their own tabular resolvers first via resolveFormXIXCombinedHeaderFieldLayout.
  if (!isFormXIXAPTableWageSlipContext(formHeader, item, fileName, sheetText)) return null;

  const accessor = buildWorkbookMergedCellAccessor(workbook, hints);
  const getMergedAwareCellText = accessor?.getMergedAwareCellText ?? null;
  const effectiveSheetCols = accessor?.effectiveSheetCols ?? 20;
  const allFields = buildFormXIXAPTemplateFields(getMergedAwareCellText, effectiveSheetCols);
  // Wage particulars live in the table — keep contractor / period / certification in header fields.
  const finalFields = allFields.filter((f) => f.group !== 'wages');
  const title =
    formHeader?.title ||
    (String(sheetText).toLowerCase().includes('wage slip') ? 'Form XIX – Wage Slip' : 'Form XIX');

  return {
    formHeader: {
      title,
      subtitle: formHeader?.subtitle || 'Wage Slip (Contract Labour)',
      reference: formHeader?.reference || '[Rule 78(1)(b)]',
      formXIXAPHeaderFieldLayout: true,
      formXIXAPTableLayout: true,
      textRows: [],
      fields: finalFields
    },
    headers: resolveFormXIXAPWageTableHeaders(parsed?.headers || hints.tableHeaders || null),
    tableData: [],
    headerRowIndex: -1,
    dataStartIndex: 0,
    tableStartCol: 0,
    sheetName: accessor?.sheetName || hints.preferredSheetName || null
  };
}

export function resolveFormXIXWorkbookSheetName(workbook, hints = {}) {
  return pickWorkbookSheet(workbook, hints);
}

export const formatWorkmanNameAndGuardian = (emp = {}) => {
  const fn = String(emp.FirstName || emp['FirstName'] || emp.firstName || emp['First Name'] || '').trim();
  const mn = String(
    emp.MiddleName ||
      emp['MiddleName'] ||
      emp.middleName ||
      emp['Middle Name'] ||
      emp.middle_name ||
      emp.Middle_Name ||
      ''
  ).trim();
  const ln = String(emp.LastName || emp['LastName'] || emp.lastName || emp['Last Name'] || '').trim();
  const name =
    [fn, mn, ln].filter(Boolean).join(' ').trim() ||
    String(emp.Name || emp['Name'] || emp.EmployeeName || emp['Employee Name'] || '').trim();
  const guardian = String(
    emp.Father_s_Name ||
      emp['Father_s_Name'] ||
      emp.Father_Name ||
      emp['Father_Name'] ||
      emp.FatherName ||
      emp['Father Name'] ||
      emp.Father_SpouseName ||
      emp['Father/Husband Name'] ||
      emp.Spouse_Name ||
      emp['Spouse_Name'] ||
      emp.SpouseName ||
      emp['Spouse Name'] ||
      emp.HusbandName ||
      emp['Husband Name'] ||
      ''
  ).trim();
  if (name && guardian) return `${name}\n${guardian}`;
  return name || guardian || '';
};

export function applyFormXIXAPAutofillFromSiteAndPayroll(headerData, context = {}) {
  const out = headerData && typeof headerData === 'object' ? { ...headerData } : {};
  const emp = context.employee || {};
  const payroll = context.payroll || {};
  const fill = (key, value) => {
    const text = String(value ?? '').trim();
    if (!text || !key) return;
    const cur = String(out[key] ?? '').trim();
    if (!cur || /^enter\b/i.test(cur)) out[key] = text;
  };

  fill('form_xix_ap_contractor', context.contractorText || '');
  fill('form_xix_ap_nature_location', context.natureLocationText || '');
  fill('form_xix_ap_period_ending', context.periodEndingText || '');
  fill('form_xix_ap_workman', formatWorkmanNameAndGuardian(emp));

  fill('form_xix_ap_days_worked', payroll.daysWorked ?? '');
  fill('form_xix_ap_units_worked', payroll.unitsWorked ?? '');
  fill('form_xix_ap_rate', payroll.rate ?? '');
  fill('form_xix_ap_overtime', payroll.overtimeWages ?? '');
  fill('form_xix_ap_gross', payroll.grossWages ?? '');
  fill('form_xix_ap_deductions', payroll.deductions ?? '');
  fill('form_xix_ap_net', payroll.netWages ?? '');

  return out;
}

/** Map modal wage-table row cells onto form_xix_ap_* header keys for Excel export. */
export function mergeFormXIXAPWageTableRowIntoHeaderData(headerData, row, headers = FORM_XIX_AP_WAGE_TABLE_HEADERS) {
  const out = headerData && typeof headerData === 'object' ? { ...headerData } : {};
  const hdrs = resolveFormXIXAPWageTableHeaders(headers);
  const read = (pred) => {
    const header = hdrs.find(pred);
    if (!header || !row || typeof row !== 'object') return '';
    return String(row[header] ?? '').trim();
  };
  const setIf = (key, value) => {
    const text = String(value ?? '').trim();
    if (!text || !key) return;
    out[key] = text;
  };
  const norm = (h) =>
    String(h || '')
      .replace(/^\d+[\.\)]\s*/, '')
      .replace(/\s+/g, ' ')
      .trim()
      .toLowerCase();
  setIf(
    'form_xix_ap_workman',
    read((h) => {
      const s = norm(h);
      return (s.includes('workman') || s.includes('workmen')) && (s.includes('name') || s.includes('father'));
    })
  );
  setIf('form_xix_ap_days_worked', read((h) => /days\s+worked/.test(norm(h))));
  setIf('form_xix_ap_units_worked', read((h) => /units\s+worked/.test(norm(h))));
  setIf('form_xix_ap_rate', read((h) => {
    const s = norm(h);
    // Do not match "piece-rate Workers" on the units column.
    if (/units\s+worked/.test(s)) return false;
    return /rate\s+of\s+(?:daily\s+)?wages|daily\s+wages|piece[\s/-]*rate/.test(s);
  }));
  setIf('form_xix_ap_overtime', read((h) => /overtime/.test(norm(h))));
  setIf('form_xix_ap_gross', read((h) => /gross\s+wages/.test(norm(h))));
  setIf('form_xix_ap_deductions', read((h) => /deduction/.test(norm(h))));
  setIf('form_xix_ap_net', read((h) => /net\s+amount/.test(norm(h))));
  return out;
}

export function importFormXIXAPHeaderFieldsFromWorkbook(workbook, hints = {}) {
  const accessor = buildWorkbookMergedCellAccessor(workbook, hints);
  if (!accessor) return null;
  const { getMergedAwareCellText, effectiveSheetCols } = accessor;
  const fields = buildFormXIXAPTemplateFields(getMergedAwareCellText, effectiveSheetCols);
  const out = {};
  fields.forEach((field) => {
    let val = '';
    if (field.labelRow != null && field.valueCol != null) {
      val = String(
        getMergedAwareCellText(field.valueRow ?? field.labelRow, field.valueCol) || field.value || ''
      ).trim();
      if (isNarrativeBlob(val)) val = '';
    } else if (field.value) {
      val = String(field.value).trim();
    }
    if (val) out[field.key] = val;
  });
  return out;
}

export function writeFormXIXAPFieldsToExcelJsWorksheet(worksheet, headerFormData, parsedFormHeader, helpers = {}) {
  if (!worksheet || !headerFormData || typeof headerFormData !== 'object') return;
  const excelCellValueToString =
    helpers.excelCellValueToString ||
    ((val) => {
      if (val == null) return '';
      if (typeof val === 'string' || typeof val === 'number' || typeof val === 'boolean') return String(val);
      if (val instanceof Date) return val.toISOString();
      if (typeof val === 'object') {
        if (Array.isArray(val.richText)) return val.richText.map((rt) => rt?.text || '').join('');
        if (val.text != null) return String(val.text);
        if (val.result != null) return String(val.result);
      }
      return '';
    });
  const normalize = (txt) => formXIXAPHeaderNorm(txt);
  const isPlaceholderCell = (txt) => {
    const t = String(txt || '').trim();
    if (!t) return true;
    if (t === ':') return true;
    // Template dotted lines (.….) must be overwritten with wage values.
    if (/^[.\u2026…_\-\s]+$/.test(t)) return true;
    if (/^enter\b/i.test(t)) return true;
    return false;
  };
  const isDottedPlaceholder = (txt) => {
    const t = String(txt || '').trim();
    return t.length > 0 && /^[.\u2026…_\-\s]+$/.test(t);
  };
  // Always include wage particulars 1–7 for export — modal layout may omit them from fields[].
  const parsedFields = Array.isArray(parsedFormHeader?.fields) ? parsedFormHeader.fields : [];
  const fieldsByKey = new Map();
  FORM_XIX_AP_TEMPLATE_SPECS.forEach((spec) => {
    fieldsByKey.set(spec.key, {
      key: spec.key,
      label: spec.label,
      group: spec.group,
      fieldType: spec.fieldType || 'text',
      match: spec.match,
    });
  });
  parsedFields.forEach((field) => {
    if (!field?.key) return;
    const prev = fieldsByKey.get(field.key) || {};
    fieldsByKey.set(field.key, { ...prev, ...field });
  });
  const fields = Array.from(fieldsByKey.values());
  const maxScanRows = Math.max(200, worksheet.rowCount + 10);
  const maxScanCols = Math.max(24, worksheet.columnCount || 0, 16);

  const writeAt = (row, col, value) => {
    const text = String(value ?? '').trim();
    if (!text || row == null || col == null) return;
    worksheet.getCell(row, col).value = text;
  };

  const cellMaster = (row, col) => {
    try {
      const cell = worksheet.getCell(row, col);
      const master = cell?.master;
      if (master && typeof master.row === 'number' && typeof master.col === 'number') {
        return { row: master.row, col: master.col };
      }
    } catch (_err) {
      /* ignore */
    }
    return { row, col };
  };

  const isSameMergedCell = (r1, c1, r2, c2) => {
    const a = cellMaster(r1, c1);
    const b = cellMaster(r2, c2);
    return a.row === b.row && a.col === b.col;
  };

  const cellWouldClobberLabel = (row, col, fieldKey) => {
    const existing = excelCellValueToString(worksheet.getCell(row, col)?.value);
    if (isFormXIXAPTemplateLabelText(existing, fieldKey)) return true;
    const master = cellMaster(row, col);
    if (master.row !== row || master.col !== col) {
      const masterText = excelCellValueToString(worksheet.getCell(master.row, master.col)?.value);
      if (isFormXIXAPTemplateLabelText(masterText, fieldKey)) return true;
    }
    return false;
  };

  const fillPeriodEndingOnLabel = (labelRow, labelCol, value) => {
    const raw = excelCellValueToString(worksheet.getCell(labelRow, labelCol)?.value);
    if (!/week.*fortnight.*month\s+ending/i.test(raw)) return false;
    const heading = String(raw)
      .replace(/[.\u2026…]+/g, ' ')
      .replace(/\s+/g, ' ')
      .replace(/:+$/, '')
      .trim();
    worksheet.getCell(labelRow, labelCol).value = `${heading} ${value}`.replace(/\s+/g, ' ').trim();
    return true;
  };

  // AP printed template: nature → D, period ending → L, workman → N (same row as label).
  const isApPrintedLayout =
    !!(parsedFormHeader?.formXIXAPTableLayout || parsedFormHeader?.formXIXAPHeaderFieldLayout) &&
    !(parsedFormHeader?.formXIXMPTableLayout && !parsedFormHeader?.formXIXAPTableLayout);

  const writeApFixedHeaderValue = (labelRow, fieldKey, value) => {
    if (!isApPrintedLayout || labelRow == null) return false;
    const col = FORM_XIX_AP_HEADER_VALUE_COLS[fieldKey];
    if (!col) return false;
    if (cellWouldClobberLabel(labelRow, col, fieldKey)) return false;
    writeAt(labelRow, col, value);
    return true;
  };

  /** Keep printed headings; put values under/beside the label (two-column AP wage slip). */
  const writePreservingLabel = (labelRow, labelCol, value, fieldKey) => {
    if (writeApFixedHeaderValue(labelRow, fieldKey, value)) return true;
    // Non-AP / fallback: append period ending onto the dotted label text.
    if (
      !isApPrintedLayout &&
      fieldKey === 'form_xix_ap_period_ending' &&
      fillPeriodEndingOnLabel(labelRow, labelCol, value)
    ) {
      return true;
    }
    // Prefer same-row cells to the right before dropping values under the label column
    // (under-column writes put nature/location into A14 on the AP template).
    for (let c = labelCol + 1; c <= Math.min(labelCol + 14, maxScanCols); c += 1) {
      if (isSameMergedCell(labelRow, c, labelRow, labelCol)) continue;
      if (cellWouldClobberLabel(labelRow, c, fieldKey)) continue;
      const existing = excelCellValueToString(worksheet.getCell(labelRow, c)?.value);
      if (isPlaceholderCell(existing)) {
        writeAt(labelRow, c, value);
        return true;
      }
    }
    for (let r = labelRow + 1; r <= Math.min(labelRow + 3, maxScanRows); r += 1) {
      if (isSameMergedCell(r, labelCol, labelRow, labelCol)) continue;
      if (cellWouldClobberLabel(r, labelCol, fieldKey)) continue;
      const existing = excelCellValueToString(worksheet.getCell(r, labelCol)?.value);
      if (isPlaceholderCell(existing)) {
        writeAt(r, labelCol, value);
        return true;
      }
    }
    if (fieldKey === 'form_xix_ap_workman') {
      const raw = excelCellValueToString(worksheet.getCell(labelRow, labelCol)?.value).replace(/:+$/, '');
      if (/workman/i.test(raw)) {
        worksheet.getCell(labelRow, labelCol).value = `${raw}\n${value}`.trim();
        return true;
      }
    }
    return false;
  };

  /** Template stores wage values on dotted lines in cols H–L; avoid writing into empty B–G. */
  const pickValueCol = (row, labelCol, preferValueBand, fieldKey = null) => {
    let bestCol = null;
    let bestScore = -1;
    const scanEnd = Math.min(labelCol + 14, Math.max(maxScanCols + 4, 14));
    for (let nc = labelCol + 1; nc <= scanEnd; nc += 1) {
      if (isSameMergedCell(row, nc, row, labelCol)) continue;
      if (cellWouldClobberLabel(row, nc, fieldKey)) continue;
      const neighbor = excelCellValueToString(worksheet.getCell(row, nc)?.value);
      if (!isPlaceholderCell(neighbor)) continue;
      let score = 10;
      if (isDottedPlaceholder(neighbor)) score += 100;
      if (preferValueBand && nc >= 8 && nc <= 14) score += 80;
      if (preferValueBand && nc < 8) score -= 60;
      // Header fields: do not prefer the right-hand heading column (workman / period ending).
      if (!preferValueBand && nc >= 8) score -= 20;
      if (score > bestScore) {
        bestScore = score;
        bestCol = nc;
      }
    }
    return bestCol;
  };

  // MP/Gujarat stacked wage slip — header values (incl. contractor) align in column E, not H–L.
  const stackedHeaderValueCol =
    parsedFormHeader?.formXIXMPTableLayout && !parsedFormHeader?.formXIXAPTableLayout ? 5 : null;

  fields.forEach((field) => {
    const val = headerFormData[field.key];
    const valText = String(val ?? '').trim();
    if (!valText || isDottedPlaceholderText(valText)) return;

    const spec = FORM_XIX_AP_TEMPLATE_SPECS.find((s) => s.key === field.key);
    const isWageField = (field.group || spec?.group) === 'wages';

    // Header/footer fields may use parsed coords; wage particulars always locate dotted value band.
    if (!isWageField && field.labelRow != null) {
      const labelExcelRow = field.labelRow + 1;
      const labelExcelCol = field.labelCol != null ? field.labelCol + 1 : 1;
      // AP printed slip: nature → D, period → L, workman → N on the label row.
      if (writeApFixedHeaderValue(labelExcelRow, field.key, valText)) return;
      if (field.valueCol != null) {
        const targetRow = (field.valueRow ?? field.labelRow) + 1;
        // Gujarat/MP stacked: keep contractor (and other headers) in column E.
        const col =
          stackedHeaderValueCol != null && field.key === 'form_xix_ap_contractor'
            ? stackedHeaderValueCol
            : field.valueCol + 1;
        if (stackedHeaderValueCol != null && field.key === 'form_xix_ap_contractor') {
          writeAt(targetRow, col, valText);
          return;
        }
        if (!cellWouldClobberLabel(targetRow, col, field.key)) {
          writeAt(targetRow, col, valText);
          return;
        }
        if (writePreservingLabel(labelExcelRow, labelExcelCol, valText, field.key)) return;
      }
    }

    const labelNorm = normalize(String(field.label || spec?.label || '').replace(/:+$/, ''));
    if (!labelNorm) return;
    const matchRe = spec?.match || null;
    for (let r = 1; r <= maxScanRows; r += 1) {
      for (let c = 1; c <= maxScanCols; c += 1) {
        const raw = excelCellValueToString(worksheet.getCell(r, c)?.value);
        const rawNorm = normalize(raw.replace(/:+$/, ''));
        const score = labelMatchScore(labelNorm, rawNorm);
        const regexHit = matchRe && (matchRe.test(raw) || matchRe.test(rawNorm));
        if (score < 45 && !regexHit) continue;
        if (writeApFixedHeaderValue(r, field.key, valText)) return;
        // Stacked MP/GJ: do not let pickValueCol prefer the AP H–L band for header fields.
        if (stackedHeaderValueCol != null && !isWageField) {
          writeAt(r, stackedHeaderValueCol, valText);
          return;
        }
        const col = pickValueCol(r, c, isWageField, field.key);
        if (col != null && !cellWouldClobberLabel(r, col, field.key)) {
          writeAt(r, col, valText);
          return;
        }
        if (!isWageField && writePreservingLabel(r, c, valText, field.key)) return;
        // Fall back: first placeholder in the H–L value band on this / next rows.
        for (let r2 = r; r2 <= Math.min(r + 2, maxScanRows); r2 += 1) {
          const bandCol = pickValueCol(r2, 7, true, field.key);
          if (bandCol != null && bandCol >= 8 && !cellWouldClobberLabel(r2, bandCol, field.key)) {
            writeAt(r2, bandCol, valText);
            return;
          }
        }
        if (isWageField) {
          writeAt(r, Math.min(Math.max(c + 7, 8), maxScanCols), valText);
        }
        return;
      }
    }
  });
}

function formXIXAPRowHasExportData(row, headers) {
  if (!row || typeof row !== 'object') return false;
  const hdrs = resolveFormXIXAPWageTableHeaders(headers);
  return hdrs.some((h) => String(row[h] ?? '').trim() !== '');
}

function sanitizeFormXIXAPHeaderFormData(headerFormData) {
  const out = headerFormData && typeof headerFormData === 'object' ? { ...headerFormData } : {};
  Object.keys(out).forEach((key) => {
    if (out[key] == null) return;
    out[key] = String(out[key]).trim();
  });
  return out;
}

function applyFormXIXAPExportWageRulesToHeaderData(headerData) {
  const out = sanitizeFormXIXAPHeaderFormData(headerData);
  out.form_xix_ap_overtime = 'NIL';
  const gross = Number(String(out.form_xix_ap_gross ?? '').replace(/,/g, '').trim());
  const net = Number(String(out.form_xix_ap_net ?? '').replace(/,/g, '').trim());
  if (Number.isFinite(gross) && Number.isFinite(net)) {
    const diff = Math.round((gross - net) * 100) / 100;
    out.form_xix_ap_deductions = String(diff >= 0 ? diff : 0);
  }
  return out;
}

export { applyFormXIXAPExportWageRulesToHeaderData };

function resolveFormXIXAPEmployeeDownloadBaseName(row, headers, index) {
  const hdrs = resolveFormXIXAPWageTableHeaders(headers);
  const workmanHeader = hdrs.find((h) => {
    const s = formXIXAPHeaderNorm(h);
    return (s.includes('workman') || s.includes('workmen')) && (s.includes('name') || s.includes('father'));
  });
  const raw = workmanHeader ? String(row?.[workmanHeader] ?? '').split(/\r?\n/)[0].trim() : '';
  const safe = String(raw || `Employee_${index + 1}`)
    .replace(/[^a-zA-Z0-9._-]+/g, '_')
    .replace(/^_+|_+$/g, '')
    .slice(0, 60);
  return safe || `Employee_${index + 1}`;
}

function allocateUniqueFormXIXAPDownloadFileName(baseName, usedNames) {
  const root = String(baseName || 'Employee').replace(/\.xlsx?$/i, '');
  const count = usedNames.get(root) || 0;
  usedNames.set(root, count + 1);
  const suffix = count > 0 ? `_${count + 1}` : '';
  return `Form_XIX_AP_${root}${suffix}.xlsx`;
}

export async function buildFormXIXAPWorkbookWithTemplateStyles({
  templateArrayBuffer,
  parsedFormHeader,
  headerFormData,
  formFileName,
  sheetNameHint,
}) {
  if (!templateArrayBuffer) {
    throw new Error('Original form template buffer is required for Form XIX AP export.');
  }
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(templateArrayBuffer);
  const sheetCandidates = Array.isArray(workbook.worksheets) ? workbook.worksheets : [];
  const preferred = String(sheetNameHint || '').trim();
  const worksheet =
    (preferred && sheetCandidates.find((ws) => String(ws?.name || '').trim() === preferred)) ||
    sheetCandidates.find((ws) => /wage\s+slip|form\s*xix/i.test(String(ws?.name || ''))) ||
    sheetCandidates.find((ws) => /xix/i.test(String(ws?.name || ''))) ||
    sheetCandidates[0] ||
    null;
  if (!worksheet) throw new Error('Template worksheet not found.');

  writeFormXIXAPFieldsToExcelJsWorksheet(
    worksheet,
    applyFormXIXAPExportWageRulesToHeaderData(headerFormData),
    parsedFormHeader
  );

  const out = await workbook.xlsx.writeBuffer();
  const fileName =
    formFileName ||
    parsedFormHeader?.title?.replace(/[^a-zA-Z0-9]/g, '_') ||
    'Form_XIX_AP.xlsx';
  return {
    blob: new Blob([out], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    }),
    fileName,
  };
}

export function triggerFormXIXAPZipDownload(blob, fileName) {
  if (!blob || !fileName) return;
  const downloadUrl = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = downloadUrl;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(downloadUrl);
}

/** One wage-slip XLSX per employee, packaged as ZIP (same pattern as Form XIX MP). */
export async function buildFormXIXAPPerEmployeeDownload({
  templateArrayBuffer,
  mappedData,
  headersToUse,
  parsedFormHeader,
  formFileName,
  headerFormData,
  sheetNameHint,
}) {
  if (!templateArrayBuffer) {
    throw new Error('Original form template buffer is required for Form XIX AP ZIP export.');
  }

  const hdrs = resolveFormXIXAPWageTableHeaders(headersToUse);
  const exportRows = (Array.isArray(mappedData) ? mappedData : []).filter((row) =>
    formXIXAPRowHasExportData(row, hdrs)
  );
  const rowsForZip = exportRows.length > 0 ? exportRows : [null];
  const baseHeaderData = sanitizeFormXIXAPHeaderFormData(headerFormData);
  // Export must include wage particulars 1–7 (modal UI keeps them only in the table).
  const formXIXFields = FORM_XIX_AP_TEMPLATE_SPECS.map((spec) => {
    const fromParsed = Array.isArray(parsedFormHeader?.fields)
      ? parsedFormHeader.fields.find((f) => f?.key === spec.key)
      : null;
    return {
      key: spec.key,
      label: spec.label,
      group: spec.group,
      fieldType: spec.fieldType || 'text',
      ...(fromParsed || {}),
      key: spec.key,
      label: fromParsed?.label || spec.label,
      group: spec.group,
    };
  });
  const headerForWrite = {
    ...(parsedFormHeader || {}),
    fields: formXIXFields,
    formXIXAPHeaderFieldLayout: true,
    formXIXAPTableLayout: true,
  };

  const zip = new JSZip();
  const usedNames = new Map();
  for (let i = 0; i < rowsForZip.length; i += 1) {
    const row = rowsForZip[i];
    const mergedHeaderData = applyFormXIXAPExportWageRulesToHeaderData(
      mergeFormXIXAPWageTableRowIntoHeaderData(baseHeaderData, row, hdrs)
    );
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(templateArrayBuffer);
    const sheetCandidates = Array.isArray(workbook.worksheets) ? workbook.worksheets : [];
    const preferred = String(sheetNameHint || '').trim();
    const worksheet =
      (preferred && sheetCandidates.find((ws) => String(ws?.name || '').trim() === preferred)) ||
      sheetCandidates.find((ws) => /wage\s+slip|form\s*xix/i.test(String(ws?.name || ''))) ||
      sheetCandidates.find((ws) => /xix/i.test(String(ws?.name || ''))) ||
      sheetCandidates[0] ||
      null;
    if (!worksheet) throw new Error('Template worksheet not found.');
    writeFormXIXAPFieldsToExcelJsWorksheet(worksheet, mergedHeaderData, headerForWrite);
    const out = await workbook.xlsx.writeBuffer();
    const xlsxBytes = out instanceof Uint8Array ? out : new Uint8Array(out);
    const baseName = resolveFormXIXAPEmployeeDownloadBaseName(row, hdrs, i);
    zip.file(allocateUniqueFormXIXAPDownloadFileName(baseName, usedNames), xlsxBytes);
    if (i > 0 && i % 15 === 0) {
      await new Promise((resolve) => setTimeout(resolve, 0));
    }
  }

  const zipBase = String(formFileName || parsedFormHeader?.title || 'Form_XIX_AP')
    .replace(/\.xlsx?$/i, '')
    .replace(/[^a-zA-Z0-9._-]+/g, '_');
  return {
    blob: await zip.generateAsync({ type: 'blob', compression: 'STORE' }),
    fileName: `${zipBase}_Employees.zip`,
  };
}
