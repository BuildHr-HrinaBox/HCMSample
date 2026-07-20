/**
 * Tamil Nadu CLRA Form XXVII — Register of Wages [Rule 78(1)(a)].
 * Template Form_XXVII_-_TamilNadu.xlsx multi-tier header (from Excel):
 *
 *   WAGES EARNED                          | GROSS | DEDUCTIONS
 *     … Basic, DA,
 *     ALLOWANCES/CASH PAYMENT NATURE… (13–17)
 *       WASH ALLOW | HRA | STB | WAGES INCLUDING… | OTHER ALLOWANCES, ECCA
 *     GROSS WAGES (18)
 *     PROVIDEND FUND | ESI |
 *     OTHER (21–22): PT | UNIFORM DEPOSITS
 *     FINES | OTHER DEDUCTIONS | TOTAL DEDUCTIONS
 */

export const FORM_XXVII_TN_TITLE = 'FORM XXVII';
export const FORM_XXVII_TN_SUBTITLE = 'Register of Wages';
export const FORM_XXVII_TN_REFERENCE =
  '[See Rule 78 (1) (a) of the Tamil Nadu Contract Labour (Regulation and Abolition) Rules, 1975]';

/** Mid-tier band over columns 13–17 (Excel merged header). */
export const FORM_XXVII_TN_GROUP_OTHER_ALLOWANCES =
  'OTHER ALLOWANCES/CASH PAYMENT NATURE TO BE SPECIFIED';
export const FORM_XXVII_TN_GROUP_WAGES_EARNED = 'WAGES EARNED';
export const FORM_XXVII_TN_GROUP_DEDUCTIONS = 'DEDUCTIONS';
/** Mid-tier band over columns 21–22 under Deductions. */
export const FORM_XXVII_TN_GROUP_OTHER_DEDUCTIONS = 'OTHER';

/** Official Excel leaf titles for columns 13–17 under ALLOWANCES/CASH PAYMENT. */
export const FORM_XXVII_TN_OTHER_ALLOWANCE_LEAVES = [
  'WASH ALLOW',
  'HRA',
  'STB',
  'WAGES INCLUDING CASH IN LIEU OF KINDS',
  'OTHER ALLOWANCES, ECCA',
];

/** Official Excel leaf titles for columns 21–22 under OTHER (deductions). */
export const FORM_XXVII_TN_OTHER_DEDUCTION_LEAVES = ['PT', 'UNIFORM DEPOSITS'];

export const FORM_XXVII_TN_OTHER_ALLOWANCE_COL_NOS = ['13', '14', '15', '16', '17'];

/** Columns that must stay blank on Autofill (manual entry only). */
export function isFormXXVIITamilNaduSkipAutofillHeader(header) {
  const s = formXXVIITamilNaduHeaderNorm(header);
  if (!s) return false;
  // TOTAL UNPAID AMOUNT ACCUMULATED (template spelling varies: unpaid / umpaid)
  if (/unpaid|umpaid/.test(s) && /accumul/.test(s)) return true;
  if (/total/.test(s) && /unpaid|umpaid/.test(s) && /amount/.test(s)) return true;
  // WAGES INCLUDING CASH IN LIEU OF KIND(S) — do not fetch payroll / days
  if (/cash\s+in\s+lieu/.test(s)) return true;
  if (/wages?\s+including/.test(s) && /lieu/.test(s) && /kind/.test(s)) return true;
  return false;
}

const FORM_XXVII_TN_MONTH_NAMES = [
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

/** Resolve { monthName, year } from YYYY-MM. */
export function resolveFormXXVIITamilNaduPeriodParts(payrollMonthIso = '') {
  const m = String(payrollMonthIso || '').trim();
  if (/^\d{4}-\d{2}$/.test(m)) {
    const year = parseInt(m.slice(0, 4), 10);
    const monthIdx = parseInt(m.slice(5, 7), 10) - 1;
    if (monthIdx >= 0 && monthIdx < 12 && Number.isFinite(year)) {
      return {
        iso: m,
        monthName: FORM_XXVII_TN_MONTH_NAMES[monthIdx],
        year: String(year),
        monthIdx,
      };
    }
  }
  return { iso: '', monthName: '', year: '', monthIdx: -1 };
}

/**
 * Excel template uses "Wage Period :" (Z4) + "Year: 2026" (Z5).
 * Banner / cell text: "Wage Period : April"
 */
export function buildFormXXVIITamilNaduWagePeriodLine(payrollMonthIso) {
  const parts = resolveFormXXVIITamilNaduPeriodParts(payrollMonthIso);
  if (!parts.monthName) return '';
  return `Wage Period : ${parts.monthName}`;
}

function excelCellToPlainText(value) {
  if (value == null) return '';
  if (typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') {
    return String(value);
  }
  if (typeof value === 'object') {
    if (value.text != null) return String(value.text);
    if (value.result != null) return String(value.result);
    if (value.richText && Array.isArray(value.richText)) {
      return value.richText.map((p) => String(p?.text || '')).join('');
    }
  }
  return String(value);
}

/**
 * Write Wage Period : <Month> and Year: <YYYY> into Form_XXVII_-_TamilNadu.xlsx cells.
 */
export function writeFormXXVIITamilNaduPeriodFieldsToExcelJsWorksheet(worksheet, payrollMonthIso) {
  if (!worksheet) return false;
  const parts = resolveFormXXVIITamilNaduPeriodParts(payrollMonthIso);
  const wageLine = buildFormXXVIITamilNaduWagePeriodLine(payrollMonthIso);
  if (!parts.monthName && !parts.year) return false;

  const maxRows = Math.min(30, worksheet.rowCount || 30);
  const maxCols = Math.min(50, worksheet.columnCount || 50);
  let wroteWage = false;
  let wroteYear = false;

  for (let r = 1; r <= maxRows; r += 1) {
    for (let c = 1; c <= maxCols; c += 1) {
      const raw = excelCellToPlainText(worksheet.getCell(r, c)?.value).replace(/\s+/g, ' ').trim();
      if (!raw) continue;
      const lower = raw.toLowerCase();

      // "Wage Period :" or "Wage Period : April" (not Form W "Wage Period from …")
      if (
        wageLine &&
        /wage\s*period\s*:/i.test(lower) &&
        !/wage\s*period\s*from/i.test(lower)
      ) {
        worksheet.getCell(r, c).value = wageLine;
        wroteWage = true;
        continue;
      }
      if (wageLine && /^wage\s*period\s*:?\s*$/i.test(raw)) {
        worksheet.getCell(r, c).value = wageLine;
        wroteWage = true;
        continue;
      }

      if (parts.year && /^year\s*:?\s*$/i.test(raw)) {
        const adj = worksheet.getCell(r, c + 1);
        const adjText = excelCellToPlainText(adj?.value).trim();
        if (!adjText || /^enter\b/i.test(adjText) || /^\d{4}$/.test(adjText)) {
          adj.value = parts.year;
        } else {
          worksheet.getCell(r, c).value = `Year: ${parts.year}`;
        }
        wroteYear = true;
        continue;
      }
      if (parts.year && /^year\s*:/i.test(raw) && !/month|wage/i.test(raw)) {
        worksheet.getCell(r, c).value = `Year: ${parts.year}`;
        wroteYear = true;
      }
    }
  }

  return wroteWage || wroteYear;
}

export function formXXVIITamilNaduHeaderNorm(txt) {
  return String(txt || '')
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

/** Form 27 / XXVII — underscore filenames break \\b word boundaries. */
export function matchesFormXXVIIHint(blob) {
  const parts = String(blob || '').toLowerCase();
  if (/form[\s._-]*27[\s._-]*c\b|form[\s._-]*xxvii[\s._-]*c\b/i.test(parts)) return false;
  return (
    /form[\s._-]*xxvii(?![a-z])/i.test(parts) ||
    /form[\s._-]*27(?!\d)/i.test(parts) ||
    /(?:^|[\s._-])xxvii(?=[\s._\W-]|$)/i.test(parts)
  );
}

export function looksLikeFormXXVIITamilNaduText(text) {
  const s = String(text || '').toLowerCase();
  return (
    /tamil[\s._-]*nadu|tamilnadu/.test(s) ||
    /form[\s._-]*xxvii[\s._-]*tamil|form_xxvii[_\s-]*tamil|form[\s._-]*27[\s._-]*tamil/.test(s)
  );
}

function buildFormXXVIIContextBlob(formHeader, rowItem, fileName, tableHeaders, sheetText) {
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

/** AP Shops Form XXVII Quarterly Returns — not TN CLRA wage register. */
export function isFormXXVIIQuarterlyReturnsBlob(blob) {
  const s = String(blob || '').toLowerCase();
  return (
    /quarterly\s+return/.test(s) ||
    (/shops?\s+and\s+establishments?/.test(s) && /quarter/.test(s)) ||
    /form[\s._-]*xxvii[\s._-]*quarter|form[\s._-]*27[\s._-]*quarter/.test(s)
  );
}

export function isFormXXVIITamilNaduContext(
  formHeader,
  rowItem,
  fileName,
  tableHeaders = '',
  sheetText = ''
) {
  const blob = buildFormXXVIIContextBlob(formHeader, rowItem, fileName, tableHeaders, sheetText);
  if (!blob) return false;
  if (isFormXXVIIQuarterlyReturnsBlob(blob)) return false;
  if (!matchesFormXXVIIHint(blob) && !looksLikeFormXXVIITamilNaduWageHeaders(tableHeaders)) {
    return false;
  }
  const fileBlob = [fileName, rowItem?.formFileName, rowItem?.FormFileName, rowItem?.formName, rowItem?.FormName]
    .map((v) => String(v || '').toLowerCase())
    .join(' ');
  return (
    looksLikeFormXXVIITamilNaduText(blob) ||
    looksLikeFormXXVIITamilNaduText(fileBlob) ||
    (matchesFormXXVIIHint(fileBlob) && looksLikeFormXXVIITamilNaduWageHeaders(tableHeaders))
  );
}

export function looksLikeFormXXVIITamilNaduWageHeaders(headers) {
  const list = Array.isArray(headers) ? headers : [];
  if (list.length < 8) return false;
  const joined = list.map((h) => formXXVIITamilNaduHeaderNorm(h)).join(' | ');
  const hasBasic = /basic\s+wage/.test(joined);
  const hasDa = /dearness\s+allowance|\bda\b/.test(joined);
  const hasGross = /gross/.test(joined);
  const hasOtherBand =
    /other\s+allowances?\/?\s*cash|allowances?\/?\s*cash\s+payment|wash\s+allow|\bhra\b|\bstb\b|column\s*1[3-7]|(?:^|\|)\s*1[3-7]\s*(?:\||$)/.test(
      joined
    );
  const hasOtRate = /overtime\s+rate/.test(joined);
  return hasBasic && hasDa && hasGross && (hasOtherBand || hasOtRate || list.length >= 18);
}

export function isFormXXVIITamilNaduOtherAllowancesGroupLabel(label) {
  const s = formXXVIITamilNaduHeaderNorm(label);
  if (!s) return false;
  if (/other\s+allowances?\/?\s*cash\s+payment/.test(s)) return true;
  if (/allowances?\/?\s*cash\s+payment/.test(s)) return true;
  if (/other\s+allowances?/.test(s) && /cash|nature|specified/.test(s)) return true;
  if (/^other\s+allowances?$/.test(s)) return true;
  return false;
}

export function isFormXXVIITamilNaduGenericLeafHeader(header) {
  const raw = String(header || '').trim();
  if (!raw) return true;
  const s = formXXVIITamilNaduHeaderNorm(raw).replace(/:+$/, '').trim();
  if (!s) return true;
  if (/^column\s*\d+$/i.test(raw)) return true;
  if (/^\(?\d{1,2}\)?$/.test(s)) return true;
  if (isFormXXVIITamilNaduOtherAllowancesGroupLabel(raw)) return true;
  if (/^other$/.test(s)) return true;
  if (/^deductions?$/.test(s) || /^wages?\s+earned$/.test(s)) return true;
  return false;
}

function isDearnessAllowanceHeader(header) {
  const s = formXXVIITamilNaduHeaderNorm(header);
  return /dearness\s+allowance/.test(s) || /^da$/.test(s) || /^d\.?a\.?$/.test(s);
}

function isGrossWagesHeader(header) {
  const s = formXXVIITamilNaduHeaderNorm(header);
  return /^gross\b/.test(s) || /gross\s+wages?/.test(s) || /gross\s+total/.test(s);
}

function isProvidentFundHeader(header) {
  const s = formXXVIITamilNaduHeaderNorm(header);
  return /providen[dt]\s+fund|\bpf\b/.test(s);
}

function isEsiHeader(header) {
  const s = formXXVIITamilNaduHeaderNorm(header);
  return /^esi$/.test(s) || /^e\.?s\.?i\.?$/.test(s);
}

function isFinesHeader(header) {
  const s = formXXVIITamilNaduHeaderNorm(header);
  return /\bfines?\b/.test(s);
}

function isTotalDeductionsHeader(header) {
  const s = formXXVIITamilNaduHeaderNorm(header);
  return /total\s+deductions?/.test(s);
}

function isOtherDeductionsLeafHeader(header) {
  const s = formXXVIITamilNaduHeaderNorm(header);
  // Column 24 standalone "OTHER DEDUCTIONS" — not the OTHER (21–22) band.
  return /other\s+deductions?/.test(s) && !/indicating|nature/.test(s);
}

function findDearnessAndGrossIndices(headers) {
  let daIdx = -1;
  let grossIdx = -1;
  for (let i = 0; i < headers.length; i += 1) {
    if (daIdx < 0 && isDearnessAllowanceHeader(headers[i])) daIdx = i;
    if (grossIdx < 0 && isGrossWagesHeader(headers[i])) grossIdx = i;
  }
  return { daIdx, grossIdx };
}

function findEsiAndFinesIndices(headers) {
  let esiIdx = -1;
  let finesIdx = -1;
  for (let i = 0; i < headers.length; i += 1) {
    if (esiIdx < 0 && isEsiHeader(headers[i])) esiIdx = i;
    if (finesIdx < 0 && isFinesHeader(headers[i])) finesIdx = i;
  }
  return { esiIdx, finesIdx };
}

function isCanonicalOtherAllowanceLeaf(header) {
  const s = formXXVIITamilNaduHeaderNorm(header);
  return FORM_XXVII_TN_OTHER_ALLOWANCE_LEAVES.some((leaf) => formXXVIITamilNaduHeaderNorm(leaf) === s);
}

function isCanonicalOtherDeductionLeaf(header) {
  const s = formXXVIITamilNaduHeaderNorm(header);
  return FORM_XXVII_TN_OTHER_DEDUCTION_LEAVES.some((leaf) => formXXVIITamilNaduHeaderNorm(leaf) === s);
}

function isOtherAllowancesLeafHeader(header, index, headers) {
  const raw = String(header || '').trim();
  if (isFormXXVIITamilNaduOtherAllowancesGroupLabel(raw)) return true;
  if (isCanonicalOtherAllowanceLeaf(raw)) return true;
  if (/^column\s*1[3-7]$/i.test(raw) || /^1[3-7]$/.test(raw)) return true;
  if (/wash\s+allow|\bhra\b|\bstb\b|cash\s+in\s+lieu|other\s+allowances?\s*,?\s*ecca/.test(
    formXXVIITamilNaduHeaderNorm(raw)
  )) {
    return true;
  }
  if (!isFormXXVIITamilNaduGenericLeafHeader(raw)) return false;
  const { daIdx, grossIdx } = findDearnessAndGrossIndices(headers);
  if (daIdx >= 0 && grossIdx > daIdx) {
    return index > daIdx && index < grossIdx;
  }
  return false;
}

function isOtherDeductionsBandLeafHeader(header, index, headers) {
  const raw = String(header || '').trim();
  if (isCanonicalOtherDeductionLeaf(raw)) return true;
  if (/^pt$/.test(formXXVIITamilNaduHeaderNorm(raw))) return true;
  if (/uniform\s+deposits?/.test(formXXVIITamilNaduHeaderNorm(raw))) return true;
  if (!isFormXXVIITamilNaduGenericLeafHeader(raw) && !/^column\s*2[12]$/i.test(raw) && !/^2[12]$/.test(raw)) {
    return false;
  }
  const { esiIdx, finesIdx } = findEsiAndFinesIndices(headers);
  if (esiIdx >= 0 && finesIdx > esiIdx) {
    return index > esiIdx && index < finesIdx;
  }
  return /^column\s*2[12]$/i.test(raw) || /^2[12]$/.test(raw);
}

function shouldReplaceOtherAllowanceLeaf(header) {
  const raw = String(header || '').trim();
  if (!raw) return true;
  if (isFormXXVIITamilNaduGenericLeafHeader(raw)) return true;
  if (/^1[3-7]$/.test(raw) || /^column\s*1[3-7]$/i.test(raw)) return true;
  // Keep a meaningful Excel leaf if it already matches the official set (or close).
  if (isCanonicalOtherAllowanceLeaf(raw)) return false;
  if (/wash\s+allow|\bhra\b|\bstb\b|cash\s+in\s+lieu|ecca/.test(formXXVIITamilNaduHeaderNorm(raw))) {
    return false;
  }
  return true;
}

function shouldReplaceOtherDeductionLeaf(header) {
  const raw = String(header || '').trim();
  if (!raw) return true;
  if (isFormXXVIITamilNaduGenericLeafHeader(raw)) return true;
  if (/^2[12]$/.test(raw) || /^column\s*2[12]$/i.test(raw)) return true;
  if (isCanonicalOtherDeductionLeaf(raw)) return false;
  if (/^pt$/.test(formXXVIITamilNaduHeaderNorm(raw))) return false;
  if (/uniform\s+deposits?/.test(formXXVIITamilNaduHeaderNorm(raw))) return false;
  // Mis-parsed long "OTHER DEDUCTIONS (INDICATING…)" must not sit in the OTHER band.
  if (/other\s+deductions?/.test(formXXVIITamilNaduHeaderNorm(raw))) return true;
  return true;
}

/**
 * Align leaf headers to Excel Form XXVII:
 * cols 13–17 → WASH ALLOW / HRA / STB / WAGES INCLUDING… / OTHER ALLOWANCES, ECCA
 * cols 21–22 → PT / UNIFORM DEPOSITS
 */
export function resolveFormXXVIITamilNaduTableHeaders(headers = []) {
  const list = Array.isArray(headers) ? headers.map((h) => String(h || '').trim()) : [];
  if (!list.length) return list;

  const out = [...list];
  const { daIdx, grossIdx } = findDearnessAndGrossIndices(out);

  if (daIdx >= 0 && grossIdx > daIdx + 1) {
    let slot = 0;
    for (let i = daIdx + 1; i < grossIdx; i += 1) {
      if (!shouldReplaceOtherAllowanceLeaf(out[i]) && isCanonicalOtherAllowanceLeaf(out[i])) {
        slot += 1;
        continue;
      }
      if (!shouldReplaceOtherAllowanceLeaf(out[i])) {
        slot += 1;
        continue;
      }
      const canonical = FORM_XXVII_TN_OTHER_ALLOWANCE_LEAVES[slot];
      if (canonical) out[i] = canonical;
      slot += 1;
    }
  }

  const { esiIdx, finesIdx } = findEsiAndFinesIndices(out);
  if (esiIdx >= 0 && finesIdx > esiIdx + 1) {
    let slot = 0;
    for (let i = esiIdx + 1; i < finesIdx; i += 1) {
      if (!shouldReplaceOtherDeductionLeaf(out[i])) {
        slot += 1;
        continue;
      }
      const canonical = FORM_XXVII_TN_OTHER_DEDUCTION_LEAVES[slot];
      if (canonical) out[i] = canonical;
      slot += 1;
    }
  }

  return dedupeFormXXVIILeafHeaders(out);
}

function dedupeFormXXVIILeafHeaders(headers) {
  const seen = new Map();
  return headers.map((h) => {
    const base = String(h || '').trim();
    if (!base) return '';
    const key = base.toLowerCase();
    const n = seen.get(key) || 0;
    seen.set(key, n + 1);
    if (n === 0) return base;
    return `${base} (${n + 1})`;
  });
}

/**
 * Parent band labels for thead: OTHER ALLOWANCES (13–17) and OTHER (21–22).
 * Standalone leaves get label === header (rowspan=2).
 */
export function buildFormXXVIITamilNaduColumnGroupLabels(headers = []) {
  const list = Array.isArray(headers) ? headers : [];
  const n = list.length;
  const labels = Array.from({ length: n }, () => '');
  for (let i = 0; i < n; i += 1) {
    const h = list[i];
    if (isOtherAllowancesLeafHeader(h, i, list)) {
      labels[i] = FORM_XXVII_TN_GROUP_OTHER_ALLOWANCES;
      continue;
    }
    if (isOtherDeductionsBandLeafHeader(h, i, list)) {
      labels[i] = FORM_XXVII_TN_GROUP_OTHER_DEDUCTIONS;
      continue;
    }
    const s = formXXVIITamilNaduHeaderNorm(h);
    if (
      /overtime\s+rate/.test(s) ||
      /basic\s+wage/.test(s) ||
      isDearnessAllowanceHeader(h) ||
      isGrossWagesHeader(h) ||
      isProvidentFundHeader(h) ||
      isEsiHeader(h) ||
      isFinesHeader(h) ||
      isOtherDeductionsLeafHeader(h) ||
      isTotalDeductionsHeader(h)
    ) {
      labels[i] = String(h || '').trim();
    }
  }
  return labels;
}

/**
 * Merge workbook-parsed group labels with official OTHER ALLOWANCES / OTHER bands.
 */
export function sanitizeFormXXVIITamilNaduColumnGroupLabels(groupLabels, headers) {
  const list = Array.isArray(headers) ? headers : [];
  const canonical = buildFormXXVIITamilNaduColumnGroupLabels(list);
  if (!Array.isArray(groupLabels) || groupLabels.length === 0) return canonical;

  return canonical.map((fallback, i) => {
    const raw = String(groupLabels[i] ?? '').trim();
    if (!raw) return fallback;
    if (isFormXXVIITamilNaduOtherAllowancesGroupLabel(raw)) {
      return FORM_XXVII_TN_GROUP_OTHER_ALLOWANCES;
    }
    const s = formXXVIITamilNaduHeaderNorm(raw);
    if (/wages?\s+earned/.test(s)) return FORM_XXVII_TN_GROUP_WAGES_EARNED;
    if (/^deductions?$/.test(s)) return FORM_XXVII_TN_GROUP_DEDUCTIONS;
    if (/^other$/.test(s) && fallback === FORM_XXVII_TN_GROUP_OTHER_DEDUCTIONS) {
      return FORM_XXVII_TN_GROUP_OTHER_DEDUCTIONS;
    }
    if (fallback === FORM_XXVII_TN_GROUP_OTHER_ALLOWANCES) return fallback;
    if (fallback === FORM_XXVII_TN_GROUP_OTHER_DEDUCTIONS) return fallback;
    return raw;
  });
}

/** True when leaf headers need the explicit group thead (not Excel-merge-only). */
export function formXXVIITamilNaduNeedsOtherAllowancesGroupThead(headers) {
  const list = Array.isArray(headers) ? headers : [];
  if (list.length < 8) return false;
  let otherLeaves = 0;
  let hasSynthetic = false;
  for (let i = 0; i < list.length; i += 1) {
    if (isOtherAllowancesLeafHeader(list[i], i, list)) otherLeaves += 1;
    const t = String(list[i] || '').trim();
    if (/^column\s*\d+$/i.test(t) || /^\d{1,2}$/.test(t)) hasSynthetic = true;
  }
  return otherLeaves >= 2 || hasSynthetic;
}
