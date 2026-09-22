import {
  flattenPayrollEarningColumns,
  readForm10GrossPayAmount,
  readForm10NetPayAmount,
  readPayrollForm15WageAmounts,
} from '../../utils/payrollEarnings';

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
 *
 * Autofill (Sample Payroll):
 *   TOTAL / UNITS days worked ← paid_days (firstname + lastname, or full employee_name)
 *   DAILY/PIECE/MONTHLY RATED (type) ← "Monthly" (never fetch payroll)
 *   DAILY RATED WAGES / PIECE RATES (amount) ← gross_pay
 *   WAGE PERIOD ← "Monthly"
 *   OVERTIME RATE ← NIL (never fetch payroll)
 *   HRA ← hra
 *   OTHER ALLOWANCES, ECCA ← gross − basic − hra
 *   PT ← professional_tax
 *   OTHER DEDUCTIONS ← (gross − net) − PT − ESI − PF
 *   TOTAL DEDUCTIONS ← gross_pay − net_pay
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

/** Column-number strip (1–28) under Form XXVII headings — always horizontal. */
export function isFormXXVIITamilNaduColumnIndexHeader(text) {
  return /^\(?\s*\d{1,2}\s*\)?$/.test(String(text || '').trim());
}

export function isFormXXVIITamilNaduColumnIndexRow(row, expectedCols = 28) {
  if (!Array.isArray(row) || row.length === 0) return false;
  const limit = Math.min(Math.max(1, expectedCols), row.length);
  let consecutive = 0;
  for (let c = 0; c < limit; c += 1) {
    const raw = String(row[c] ?? '').trim();
    if (!isFormXXVIITamilNaduColumnIndexHeader(raw)) break;
    const n = Number(String(raw).replace(/^\(|\)$/g, ''));
    if (n !== c + 1) break;
    consecutive += 1;
  }
  const minRequired = Math.min(20, limit);
  return consecutive >= minRequired;
}

/**
 * Official Form XXVII model: Sr.No / Name / Sex / Designation stay horizontal
 * in wide identity cells. Father/husband and employee number follow the same band.
 */
export function isFormXXVIITamilNaduIdentityHeaderText(text) {
  const h = formXXVIITamilNaduHeaderNorm(text);
  if (!h) return false;
  if (/^(?:s|sr|si|sl)\.?\s*no\b|^serial\s*(?:no|number)/.test(h)) return true;
  if (/name of(?:\s+the)?\s+(?:work(?:man|er)|employee)/.test(h)) return true;
  if (/^sex$/.test(h) || /^gender$/.test(h)) return true;
  if (/designation|nature of work/.test(h)) return true;
  if (/father|husband/.test(h)) return true;
  if (/employee\s*(?:number|id|code)/.test(h)) return true;
  return false;
}

/**
 * Official Form XXVII model: wage / allowance / deduction leaf titles read
 * bottom→top (vertical). Group banners and identity columns stay horizontal.
 */
export function isFormXXVIITamilNaduVerticalHeaderText(text) {
  const raw = String(text || '')
    .replace(/\s+/g, ' ')
    .trim();
  if (!raw) return false;
  if (isFormXXVIITamilNaduColumnIndexHeader(raw)) return false;
  if (isFormXXVIITamilNaduIdentityHeaderText(raw)) return false;
  const n = formXXVIITamilNaduHeaderNorm(raw);
  if (/^wages?\s+earned$/.test(n)) return false;
  if (/^deductions?$/.test(n)) return false;
  if (/other\s+allowances?\s*\/\s*cash\s+payment|cash\s+payment\s+nature/.test(n)) return false;
  if (/^other$/.test(n)) return false;
  return true;
}

/**
 * Split long vertical Form XXVII headings so they fit the header-box height
 * (Excel textRotation 90 / PDF angle 90). Short labels stay one line.
 */
export function splitFormXXVIITamilNaduVerticalHeaderLines(text, maxLines = 4) {
  const label = String(text || '')
    .replace(/\s+/g, ' ')
    .trim();
  if (!label) return [];
  const limit = Math.max(1, Math.min(4, Number(maxLines) || 4));
  if (limit === 1) return [label];

  const preferred = [
    [
      /^daily\s+rated\s*\/\s*piece\s+rate/i,
      ['DAILY RATED/', 'PIECE RATE/', 'MONTHLY RATED']
    ],
    [/^wage\s*period/i, ['WAGE PERIOD', 'WEEKLY/FN/MONTHLY']],
    [
      /^total\s+number\s+of\s+days\s+worked/i,
      ['TOTAL NUMBER OF DAYS', 'WORKED DURING THE', 'WEEK/FN/MONTH']
    ],
    [/^units\s+of\s+work/i, ['UNITS OF', 'WORK DONE']],
    [/^daily\s+rated\s+wages?/i, ['DAILY RATED WAGES/', 'PIECE RATES']],
    [/^over\s*-?\s*time\s+rate/i, ['OVER TIME', 'RATE']],
    [/^basic\s+wage/i, ['BASIC', 'WAGE']],
    [/^dearness\s+allow/i, ['DEARNESS', 'ALLOWANCE']],
    [/^wash\s+allow/i, ['WASH', 'ALLOW']],
    [
      /^wages?\s+including\s+cash\s+in\s+lieu/i,
      ['WAGES INCLUDING', 'CASH IN LIEU', 'OF KINDS']
    ],
    [/^leave\s+with\s+wages/i, ['LEAVE WITH WAGES', 'INCLUDING CASH', 'IN LIEU OF KIND']],
    [/^other\s+allowances?\s*,?\s*ecca/i, ['OTHER ALLOWANCES,', 'ECCA']],
    [/^providen/i, ['PROVIDENT', 'FUND']],
    [/^uniform\s+deposits?/i, ['UNIFORM', 'DEPOSITS']],
    [/^fines?/i, ['FINES', '(IF ANY)']],
    [/^other\s+deductions?/i, ['OTHER', 'DEDUCTIONS']],
    [/^total\s+deductions?/i, ['TOTAL', 'DEDUCTIONS']],
    [/^gross\s+wages?/i, ['GROSS', 'WAGES']],
    [/^net\s+wages?/i, ['NET', 'WAGES']],
    [
      /^signature/i,
      ['SIGNATURE / THUMB', 'IMPRESSION', 'CHEQUE No. & DATE', '/ BANK']
    ],
    [/^total\s+(?:un|um)paid/i, ['TOTAL UNPAID', 'AMOUNT ACCUMULATED']]
  ];
  for (const [re, parts] of preferred) {
    if (re.test(label)) return parts.slice(0, Math.max(limit, parts.length));
  }

  if (label.length <= 18 || !/[\s/]/.test(label)) return [label];

  const words = label
    .replace(/\s*\/\s*/g, ' / ')
    .split(/\s+/)
    .filter(Boolean);
  if (words.length <= 2) return [label];

  if (limit <= 2) {
    const total = label.length;
    let best = 1;
    let bestScore = Infinity;
    for (let i = 1; i < words.length; i += 1) {
      const left = words.slice(0, i).join(' ');
      const right = words.slice(i).join(' ');
      if (!right) continue;
      const score = Math.abs(left.length - right.length) + Math.abs(left.length - total / 2) * 0.25;
      if (score < bestScore) {
        bestScore = score;
        best = i;
      }
    }
    return [words.slice(0, best).join(' '), words.slice(best).join(' ')].filter(Boolean);
  }

  const chunk = Math.ceil(words.length / Math.min(limit, 3));
  const lines = [];
  for (let i = 0; i < words.length && lines.length < limit; i += chunk) {
    lines.push(words.slice(i, i + chunk).join(' '));
  }
  return lines.filter(Boolean);
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

/** Column default for "WAGE PERIOD — WEEKLY/FN/MONTHLY". */
export const FORM_XXVII_TN_WAGE_PERIOD_DEFAULT = 'Monthly';
/** DAILY RATED / PIECE RATED / MONTHLY RATED type column — never fetch payroll. */
export const FORM_XXVII_TN_DAILY_RATED_TYPE_DEFAULT = 'Monthly';
/** OVERTIME RATE is never fetched from Sample Payroll. */
export const FORM_XXVII_TN_OVERTIME_RATE_DEFAULT = 'NIL';

function parseFormXXVIITamilNaduMoney(value) {
  if (value == null || value === '') return NaN;
  const n = Number(String(value).replace(/,/g, '').trim());
  return Number.isFinite(n) ? n : NaN;
}

function moneyTextFormXXVIITamilNadu(value) {
  const n = parseFormXXVIITamilNaduMoney(value);
  if (!Number.isFinite(n)) return '';
  return String(Math.round(n * 100) / 100);
}

function firstPresentFormXXVII(...vals) {
  for (let i = 0; i < vals.length; i += 1) {
    const v = vals[i];
    if (v === null || v === undefined) continue;
    if (typeof v === 'string' && v.trim() === '') continue;
    return v;
  }
  return '';
}

/** Type column: DAILY RATED / PIECE RATED / MONTHLY RATED ← Monthly (no payroll). */
export function isFormXXVIITamilNaduDailyRatedTypeHeader(header) {
  const s = formXXVIITamilNaduHeaderNorm(header);
  if (!s) return false;
  if (/over[\s-]*time|wage\s*period|units?\s+of\s+work|days?\s+worked|basic\s+wage/.test(s)) {
    return false;
  }
  if (/daily\s+rated\s+wages?/.test(s) && !/monthly\s+rated/.test(s)) return false;
  return /monthly\s+rated/.test(s) || (/daily\s+rated/.test(s) && /piece\s+rated/.test(s));
}

/** Amount column: DAILY RATED WAGES / PIECE RATES ← gross_pay */
export function isFormXXVIITamilNaduDailyRatedHeader(header) {
  const s = formXXVIITamilNaduHeaderNorm(header);
  if (!s) return false;
  if (isFormXXVIITamilNaduDailyRatedTypeHeader(header)) return false;
  if (/over[\s-]*time|wage\s*period|units?\s+of\s+work|days?\s+worked|basic\s+wage/.test(s)) {
    return false;
  }
  return (
    /daily\s+rated/.test(s) ||
    /piece\s+rated/.test(s) ||
    /daily\s+rate/.test(s) ||
    (/piece\s+rate/.test(s) && !/over[\s-]*time/.test(s))
  );
}

/**
 * TOTAL NUMBER OF DAYS WORKED… / UNITS OF WORK DONE/NUMBER OF DAYS WORKED
 * ← Sample Payroll paid_days only.
 */
export function isFormXXVIITamilNaduDaysWorkedHeader(header) {
  const s = formXXVIITamilNaduHeaderNorm(header);
  if (!s) return false;
  if (/daily\s+rated|piece\s+rated|monthly\s+rated|over[\s-]*time|wage\s*period/.test(s)) {
    return false;
  }
  if (/units?\s+of\s+work/.test(s) && /days?\s+worked/.test(s)) return true;
  if (/number\s+of\s+days\s+worked|no\.?\s*of\s+days\s+worked/.test(s)) return true;
  if (/total\s+number\s+of\s+days/.test(s) && /worked/.test(s)) return true;
  if (/days?\s+worked/.test(s) && /week|fortnight|fn|month/.test(s)) return true;
  return /days?\s+worked/.test(s);
}

/** Read firstname / lastname from a Sample Payroll row (required for days-worked autofill). */
export function readFormXXVIITamilNaduSamplePayrollNameParts(payrollRow) {
  if (!payrollRow || typeof payrollRow !== 'object' || payrollRow.fetch_error) {
    return { firstName: '', lastName: '', fullName: '' };
  }
  const flat = flattenPayrollEarningColumns(payrollRow);
  let firstName = String(
    firstPresentFormXXVII(
      flat.first_name,
      flat.firstname,
      flat.FirstName,
      flat['First Name'],
      flat.firstName,
      payrollRow.first_name,
      payrollRow.firstname,
      payrollRow.FirstName,
      payrollRow['First Name'],
      payrollRow.firstName
    ) || ''
  ).trim();
  let lastName = String(
    firstPresentFormXXVII(
      flat.last_name,
      flat.lastname,
      flat.LastName,
      flat['Last Name'],
      flat.lastName,
      payrollRow.last_name,
      payrollRow.lastname,
      payrollRow.LastName,
      payrollRow['Last Name'],
      payrollRow.lastName
    ) || ''
  ).trim();
  const fullName = String(
    firstPresentFormXXVII(
      flat.employee_name,
      flat.EmployeeName,
      flat.full_name,
      payrollRow.employee_name,
      payrollRow.EmployeeName,
      payrollRow.full_name,
      payrollRow.name
    ) || ''
  ).trim();
  if ((!firstName || !lastName) && fullName) {
    const parts = fullName.split(/\s+/).filter(Boolean);
    if (parts.length >= 2) {
      if (!firstName) firstName = parts[0];
      if (!lastName) lastName = parts[parts.length - 1];
    }
  }
  return { firstName, lastName, fullName };
}

export function hasFormXXVIITamilNaduSamplePayrollNameParts(payrollRow) {
  const { firstName, lastName, fullName } = readFormXXVIITamilNaduSamplePayrollNameParts(payrollRow);
  if (firstName && lastName) return true;
  return String(fullName || '')
    .split(/\s+/)
    .filter(Boolean).length >= 2;
}

/** Read paid_days from Sample Payroll (empty when missing). */
export function readFormXXVIITamilNaduPaidDays(payrollRow) {
  if (!payrollRow || typeof payrollRow !== 'object' || payrollRow.fetch_error) return '';
  const flat = flattenPayrollEarningColumns(payrollRow);
  const raw = firstPresentFormXXVII(
    flat.paid_days,
    flat.Paid_days,
    flat['paid_days'],
    flat['Paid_days'],
    flat.paidDays,
    flat['Paid Days'],
    flat.days_worked,
    flat['days_worked'],
    flat.no_of_days_worked,
    payrollRow.paid_days,
    payrollRow.Paid_days,
    payrollRow['paid_days'],
    payrollRow['Paid_days'],
    payrollRow.paidDays,
    payrollRow['Paid Days'],
    payrollRow.days_worked,
    payrollRow.no_of_days_worked
  );
  if (raw === '' || raw == null) return '';
  const n = Number(String(raw).replace(/,/g, '').trim());
  if (!Number.isFinite(n) || n < 0) return '';
  return String(n);
}

/**
 * Days-worked columns ← paid_days only when Sample Payroll has
 * firstname + lastname + paid_days (never Form 25 / calendar defaults).
 */
export function resolveFormXXVIITamilNaduDaysWorked(payrollRow) {
  if (!hasFormXXVIITamilNaduSamplePayrollNameParts(payrollRow)) return '';
  return readFormXXVIITamilNaduPaidDays(payrollRow);
}

/** Body column "WAGE PERIOD — WEEKLY/FN/MONTHLY" (not banner "Wage Period : April"). */
export function isFormXXVIITamilNaduWagePeriodColumnHeader(header) {
  const s = formXXVIITamilNaduHeaderNorm(header);
  if (!s) return false;
  if (!/wage\s*period/.test(s)) return false;
  // Banner line already filled as "Wage Period : <Month>"
  if (/^wage\s*period\s*:/.test(s) && !/weekly|fortnight|fn|monthly/.test(s)) return false;
  return true;
}

/** OVERTIME RATE — never fetch Sample Payroll; always NIL. */
export function isFormXXVIITamilNaduOvertimeRateHeader(header) {
  const s = formXXVIITamilNaduHeaderNorm(header);
  if (!s || /earning|wages?\s+earned|normal/.test(s)) return false;
  return /over[\s-]*time/.test(s) && /rate/.test(s);
}

/** HRA leaf under ALLOWANCES/CASH PAYMENT */
export function isFormXXVIITamilNaduHraHeader(header) {
  const s = formXXVIITamilNaduHeaderNorm(header);
  if (!s) return false;
  return /^hra$/.test(s) || (/\bhra\b/.test(s) && !/dearness|house\s+rent/.test(s));
}

/** OTHER ALLOWANCES, ECCA ← gross − basic − hra */
export function isFormXXVIITamilNaduOtherAllowancesEccaHeader(header) {
  const s = formXXVIITamilNaduHeaderNorm(header);
  if (!s) return false;
  if (/deduction|cash\s+in\s+lieu|wash\s+allow|\bstb\b|^hra$/.test(s)) return false;
  return (
    (/other\s+allowances?/.test(s) && /ecca/.test(s)) ||
    /^other\s+allowances?\s*,?\s*ecca$/.test(s) ||
    (s.includes('other') && s.includes('allowance') && s.includes('ecca'))
  );
}

/** PT under OTHER (deductions) ← Professional Tax */
export function isFormXXVIITamilNaduPtHeader(header) {
  const s = formXXVIITamilNaduHeaderNorm(header);
  if (!s) return false;
  return /^pt$/.test(s) || /professional\s+tax/.test(s);
}

/** Standalone OTHER DEDUCTIONS ← total − PT − ESI − PF */
export function isFormXXVIITamilNaduOtherDeductionsHeader(header) {
  const s = formXXVIITamilNaduHeaderNorm(header);
  if (!s) return false;
  if (/total\s+deductions?/.test(s)) return false;
  if (/^other$/.test(s)) return false;
  return /other\s+deductions?/.test(s);
}

/**
 * Excel merges sometimes bleed "OTHER DEDUCTIONS" into cols 21–22 (PT / Uniform).
 * Keep that label only on the official col 24 leaf before PDF paint.
 */
export function normalizeFormXXVIITamilNaduRegisterPdfHeaderRows(
  rows,
  colCount = 28,
  headerBandEnd = -1,
  tableStart = 0
) {
  if (!Array.isArray(rows) || !rows.length) return rows;
  const out = rows.map((r) => (Array.isArray(r) ? [...r] : []));
  const start = Math.max(0, Number(tableStart) || 0);
  const end =
    Number.isFinite(headerBandEnd) && headerBandEnd >= start
      ? headerBandEnd
      : Math.min(out.length - 1, start + 10);

  let otherDedCol = -1;
  for (let r = end; r >= start; r -= 1) {
    if (isFormXXVIITamilNaduColumnIndexRow(out[r], colCount)) continue;
    const row = out[r] || [];
    for (let c = 0; c < Math.min(colCount, row.length); c += 1) {
      const t = String(row[c] ?? '')
        .replace(/\s+/g, ' ')
        .trim();
      if (isFormXXVIITamilNaduOtherDeductionsHeader(t)) {
        otherDedCol = c;
        break;
      }
    }
    if (otherDedCol >= 0) break;
  }
  if (otherDedCol < 0) otherDedCol = Math.min(23, colCount - 1);

  for (let r = start; r <= end && r < out.length; r += 1) {
    if (isFormXXVIITamilNaduColumnIndexRow(out[r], colCount)) continue;
    const row = out[r];
    for (let c = 0; c < Math.min(colCount, row.length); c += 1) {
      const t = String(row[c] ?? '')
        .replace(/\s+/g, ' ')
        .trim();
      if (!t) continue;
      if (!isFormXXVIITamilNaduOtherDeductionsHeader(t)) continue;
      if (c === otherDedCol) continue;
      if (c >= 20 && c <= 21) {
        row[c] = c === 20 && !/^other$/i.test(t) ? 'OTHER' : '';
      } else {
        row[c] = '';
      }
    }
  }
  return out;
}

/** TOTAL DEDUCTIONS ← gross_pay − net_pay */
export function isFormXXVIITamilNaduTotalDeductionsHeader(header) {
  const s = formXXVIITamilNaduHeaderNorm(header);
  if (!s) return false;
  if (/other\s+deductions?/.test(s)) return false;
  return /total\s+deductions?/.test(s) || /^deductions?$/.test(s);
}

/** TOTAL DEDUCTIONS = gross_pay − net_pay */
export function computeFormXXVIITamilNaduTotalDeductions(grossPay, netPay) {
  const g = parseFormXXVIITamilNaduMoney(grossPay);
  const n = parseFormXXVIITamilNaduMoney(netPay);
  if (!Number.isFinite(g) || !Number.isFinite(n) || g < n) return '';
  return Math.round((g - n) * 100) / 100;
}

export function resolveFormXXVIITamilNaduTotalDeductions(payrollRow, payrollMap = null) {
  if (!payrollRow || payrollRow.fetch_error) return '';
  const map = payrollMap && typeof payrollMap === 'object' ? payrollMap : {};
  const fromMap = firstPresentFormXXVII(map.deductionsFromGrossNet);
  if (fromMap !== '' && fromMap != null) {
    const n = parseFormXXVIITamilNaduMoney(fromMap);
    if (Number.isFinite(n)) return Math.round(n * 100) / 100;
  }
  const gross = firstPresentFormXXVII(
    map.grossPay,
    map.grossWages,
    readForm10GrossPayAmount(payrollRow)
  );
  const net = firstPresentFormXXVII(
    map.netPay,
    map.netWages,
    readForm10NetPayAmount(payrollRow)
  );
  return computeFormXXVIITamilNaduTotalDeductions(gross, net);
}

/** Type column is always Monthly — do not fetch payroll. */
export function resolveFormXXVIITamilNaduDailyRatedType(_payrollRow) {
  return FORM_XXVII_TN_DAILY_RATED_TYPE_DEFAULT;
}

/** DAILY RATED WAGES / PIECE RATES ← gross_pay */
export function resolveFormXXVIITamilNaduDailyRated(payrollRow) {
  if (!payrollRow || payrollRow.fetch_error) return '';
  const gross = readForm10GrossPayAmount(payrollRow);
  return gross === '' || gross == null ? '' : String(gross);
}

/** OVERTIME RATE is always NIL — do not compute or fetch payroll OT. */
export function resolveFormXXVIITamilNaduOvertimeRate(_payrollRow) {
  return FORM_XXVII_TN_OVERTIME_RATE_DEFAULT;
}

/** HRA ← Sample Payroll hra / hra_fbp */
export function resolveFormXXVIITamilNaduHra(payrollRow) {
  if (!payrollRow || payrollRow.fetch_error) return '';
  const flat = flattenPayrollEarningColumns(payrollRow);
  const wages = readPayrollForm15WageAmounts(payrollRow);
  return moneyTextFormXXVIITamilNadu(
    firstPresentFormXXVII(
      wages.hra_fbp,
      wages.hra,
      flat.hra_fbp,
      flat.hra,
      payrollRow.hra_fbp,
      payrollRow.hra
    )
  );
}

/** OTHER ALLOWANCES, ECCA = gross_pay − basic − hra */
export function computeFormXXVIITamilNaduOtherAllowancesEcca(grossPay, basic, hra) {
  const g = parseFormXXVIITamilNaduMoney(grossPay);
  if (!Number.isFinite(g)) return '';
  const b = parseFormXXVIITamilNaduMoney(basic);
  const h = parseFormXXVIITamilNaduMoney(hra);
  const known = (Number.isFinite(b) ? b : 0) + (Number.isFinite(h) ? h : 0);
  const other = Math.round((g - known) * 100) / 100;
  return Number.isFinite(other) ? other : '';
}

export function resolveFormXXVIITamilNaduOtherAllowancesEcca(payrollRow) {
  if (!payrollRow || payrollRow.fetch_error) return '';
  const gross = readForm10GrossPayAmount(payrollRow);
  const wages = readPayrollForm15WageAmounts(payrollRow);
  const hra = resolveFormXXVIITamilNaduHra(payrollRow);
  return computeFormXXVIITamilNaduOtherAllowancesEcca(gross, wages.basic, hra);
}

/** PT ← Professional Tax from Sample Payroll */
export function resolveFormXXVIITamilNaduPt(payrollRow) {
  if (!payrollRow || payrollRow.fetch_error) return '';
  const flat = flattenPayrollEarningColumns(payrollRow);
  return moneyTextFormXXVIITamilNadu(
    firstPresentFormXXVII(
      flat.professional_tax,
      payrollRow.professional_tax,
      payrollRow['Professional Tax'],
      payrollRow.pt,
      payrollRow.PT
    )
  );
}

/**
 * OTHER DEDUCTIONS = TOTAL DEDUCTIONS − Professional Tax − ESI − PROVIDENT FUND
 * TOTAL DEDUCTIONS here is gross_pay − net_pay.
 */
export function computeFormXXVIITamilNaduOtherDeductions(totalDeductions, professionalTax, esi, providentFund) {
  const total = parseFormXXVIITamilNaduMoney(totalDeductions);
  if (!Number.isFinite(total)) return '';
  const pt = parseFormXXVIITamilNaduMoney(professionalTax);
  const e = parseFormXXVIITamilNaduMoney(esi);
  const pf = parseFormXXVIITamilNaduMoney(providentFund);
  const subtract = (Number.isFinite(pt) ? pt : 0) + (Number.isFinite(e) ? e : 0) + (Number.isFinite(pf) ? pf : 0);
  const other = Math.round((total - subtract) * 100) / 100;
  if (!Number.isFinite(other)) return '';
  return other < 0 ? 0 : other;
}

export function resolveFormXXVIITamilNaduOtherDeductions(payrollRow, payrollMap = null) {
  if (!payrollRow || payrollRow.fetch_error) return '';
  const flat = flattenPayrollEarningColumns(payrollRow);
  const map = payrollMap && typeof payrollMap === 'object' ? payrollMap : {};
  // Prefer gross − net as TOTAL DEDUCTIONS (never raw payroll total_deductions alone).
  const total = firstPresentFormXXVII(
    resolveFormXXVIITamilNaduTotalDeductions(payrollRow, map),
    map.deductionsFromGrossNet
  );
  const pt = firstPresentFormXXVII(map.pt, resolveFormXXVIITamilNaduPt(payrollRow));
  const esi = firstPresentFormXXVII(
    map.esi,
    flat.esi,
    flat.esic,
    flat.esi_contribution,
    payrollRow.esi,
    payrollRow.esic
  );
  const pf = firstPresentFormXXVII(
    map.providentFund,
    flat.epf_contribution,
    flat.pf,
    payrollRow.epf_contribution,
    payrollRow.pf
  );
  return computeFormXXVIITamilNaduOtherDeductions(total, pt, esi, pf);
}

/**
 * Apply Form XXVII TN Sample Payroll mappings onto a register row.
 * Returns true when at least one mapped cell was written.
 */
export function applyFormXXVIITamilNaduPayrollToRow(row, payrollRow, headers, helpers = {}) {
  if (!row || !Array.isArray(headers) || headers.length === 0) return false;
  const { overwrite = false, sanitizeValue = (v) => v, payrollMap = null } = helpers;
  const cellIsEmpty = (header) => {
    const v = String(row[header] ?? '').trim();
    return (
      !v ||
      /^enter\b/i.test(v) ||
      /^nil+$/i.test(v) ||
      v.toLowerCase() === 'n/a' ||
      v === '-' ||
      v === '—'
    );
  };
  const setCell = (header, value) => {
    if (!header || value === '' || value == null) return false;
    if (isFormXXVIITamilNaduSkipAutofillHeader(header)) return false;
    if (!overwrite && !cellIsEmpty(header)) return false;
    row[header] = sanitizeValue(value);
    return true;
  };

  let hit = false;
  const daysWorked = resolveFormXXVIITamilNaduDaysWorked(payrollRow);
  const dailyRated = resolveFormXXVIITamilNaduDailyRated(payrollRow);
  const hra = resolveFormXXVIITamilNaduHra(payrollRow);
  const otherAllow = resolveFormXXVIITamilNaduOtherAllowancesEcca(payrollRow);
  const pt = resolveFormXXVIITamilNaduPt(payrollRow);
  const totalDed = resolveFormXXVIITamilNaduTotalDeductions(payrollRow, payrollMap);
  const otherDed = resolveFormXXVIITamilNaduOtherDeductions(payrollRow, {
    ...(payrollMap && typeof payrollMap === 'object' ? payrollMap : {}),
    deductionsFromGrossNet: totalDed !== '' ? totalDed : payrollMap?.deductionsFromGrossNet,
  });

  headers.forEach((header) => {
    if (isFormXXVIITamilNaduWagePeriodColumnHeader(header)) {
      if (setCell(header, FORM_XXVII_TN_WAGE_PERIOD_DEFAULT)) hit = true;
      return;
    }
    if (isFormXXVIITamilNaduOvertimeRateHeader(header)) {
      if (setCell(header, FORM_XXVII_TN_OVERTIME_RATE_DEFAULT)) hit = true;
      return;
    }
    if (isFormXXVIITamilNaduDailyRatedTypeHeader(header)) {
      if (setCell(header, FORM_XXVII_TN_DAILY_RATED_TYPE_DEFAULT)) hit = true;
      return;
    }
    // Days worked: paid_days only when firstname + lastname + paid_days exist; else clear
    // Form 25 / calendar leftovers.
    if (isFormXXVIITamilNaduDaysWorkedHeader(header)) {
      if (daysWorked !== '') {
        if (setCell(header, daysWorked)) hit = true;
      } else if (overwrite) {
        const cur = String(row[header] ?? '').trim();
        if (cur) {
          row[header] = '';
          hit = true;
        }
      }
      return;
    }
    if (!payrollRow || payrollRow.fetch_error) return;
    if (isFormXXVIITamilNaduDailyRatedHeader(header)) {
      if (setCell(header, dailyRated)) hit = true;
      return;
    }
    if (isFormXXVIITamilNaduHraHeader(header)) {
      if (setCell(header, hra)) hit = true;
      return;
    }
    if (isFormXXVIITamilNaduOtherAllowancesEccaHeader(header)) {
      if (setCell(header, otherAllow)) hit = true;
      return;
    }
    if (isFormXXVIITamilNaduPtHeader(header)) {
      if (setCell(header, pt)) hit = true;
      return;
    }
    if (isFormXXVIITamilNaduTotalDeductionsHeader(header)) {
      if (setCell(header, totalDed)) hit = true;
      return;
    }
    if (isFormXXVIITamilNaduOtherDeductionsHeader(header)) {
      if (setCell(header, otherDed)) hit = true;
    }
  });
  return hit;
}
