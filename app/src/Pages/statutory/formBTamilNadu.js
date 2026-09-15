import { applyExcelJSFullBoxBordersToRange } from '../../utils/excelTableBorders';
import { resolveFormWTamilNaduDefaultPayroll } from './formWTamilNadu';

/**
 * Tamil Nadu Form B — Register of Wages (LWF monthly summary, one row).
 * Form_B_-_TamilNadu.xlsx payroll column mapping:
 * - Total emoluments … → sum of corresponding employees' gross_pay
 * - Other deductions → sum of (gross_pay − basic − hra)
 * - Amount actually paid during the month → sum of net_pay
 * - Balance due to the employees → 0 when (gross_pay − net_pay) equals net_pay (wages paid)
 */

/** True when Form B TN should show company (not site) as establishment. */
export function isFormBTamilNaduEstablishmentFromCompanyContext({
  formHeader,
  rowItem,
  fileName,
  tableHeaders,
} = {}) {
  const parts = [
    rowItem?.formName || '',
    rowItem?.FormName || '',
    rowItem?.fileName || '',
    rowItem?.FileName || '',
    rowItem?.displayFileName || '',
    rowItem?.formFileName || '',
    rowItem?.FormFileName || '',
    rowItem?.templateFileName || '',
    rowItem?.DocumentName || '',
    rowItem?.documentName || '',
    rowItem?.title || '',
    rowItem?.state || '',
    rowItem?.State || '',
    fileName || '',
    formHeader?.title || '',
    formHeader?.subtitle || '',
    formHeader?.reference || '',
    Array.isArray(tableHeaders) ? tableHeaders.join(' ') : '',
  ]
    .join(' ')
    .toLowerCase();

  if (
    /\bform\s*["']?\s*b\b/.test(parts) &&
    /labour\s*welfare/.test(parts) &&
    /register\s+of\s+wages/.test(parts)
  ) {
    return true;
  }
  if (
    /form[_\s-]*b(?:[_\s-]|$)/i.test(parts) &&
    (/tamil/i.test(parts) || /register\s+of\s+wages/i.test(parts)) &&
    !/form[_\s-]*b[_\s-]*(gj|gujarat|rj|rajasthan)/i.test(parts)
  ) {
    if (
      /total\s+emoluments/i.test(parts) ||
      /total\s+number\s+of\s+empl/i.test(parts) ||
      /amounts?\s+deducted/i.test(parts) ||
      /labour\s*welfare|\blwf\b/i.test(parts) ||
      /form[_\s-]*b[_\s-]*tamil/i.test(parts) ||
      /for\s+the\s+month\s+of/i.test(parts)
    ) {
      return true;
    }
  }
  return false;
}

/** Force "Name and Address of the Establishment" from company text (never site). */
export function applyFormBTamilNaduEstablishmentFromCompany(headerData, companyText) {
  const text = String(companyText || '').trim();
  if (!headerData || typeof headerData !== 'object' || !text) return headerData;
  return {
    ...headerData,
    statutory_establishment_name_address: text,
    statutory_establishment_name: text,
    statutory_establishment_name_shop: text,
    form_t_establishment_name_address: text,
    form_h_establishment_name_address: text,
    form_f_establishment_name_address: text,
    form_a_establishment_name: text,
    form_q_establishment: text,
    form25_establishment: text,
  };
}

export function formBTamilNaduHeaderNorm(txt) {
  return String(txt || '')
    .toLowerCase()
    .replace(/\r?\n/g, ' ')
    .replace(/[_-]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

export function parseFormBTamilNaduMoney(value) {
  if (value === '' || value == null) return NaN;
  const n = Number(String(value).replace(/,/g, '').trim());
  return Number.isFinite(n) ? n : NaN;
}

function roundFormBMoney(n) {
  return Math.round(n * 100) / 100;
}

/** basic + hra for one employee (missing parts treated as 0). */
export function computeFormBTamilNaduBasicPlusHra(basic, hra) {
  const b = parseFormBTamilNaduMoney(basic);
  const h = parseFormBTamilNaduMoney(hra);
  const known = (Number.isFinite(b) ? Math.max(0, b) : 0) + (Number.isFinite(h) ? Math.max(0, h) : 0);
  return roundFormBMoney(known);
}

/**
 * Resolve basic + HRA for one location employee.
 * Force Form W TN defaults when the employee is known (same as Form W), so emoluments
 * never becomes a single employee's payroll-basic-only value (e.g. 67729 without HRA).
 */
export function resolveFormBTamilNaduBasicAndHra({ basic, hra, emp } = {}) {
  let nextBasic = basic != null && basic !== '' ? basic : '';
  let nextHra = hra != null && hra !== '' ? hra : '';
  const defaults = emp ? resolveFormWTamilNaduDefaultPayroll(emp) : null;
  if (defaults) {
    if (defaults.basic) nextBasic = defaults.basic;
    if (defaults.hra) nextHra = defaults.hra;
  }
  return { basic: nextBasic, hra: nextHra };
}

/** Other deductions = gross_pay − basic − hra (location employees). Never exceed gross. */
export function computeFormBTamilNaduOtherDeductions(grossPay, basic, hra) {
  const g = parseFormBTamilNaduMoney(grossPay);
  if (!Number.isFinite(g)) return '';
  const known = computeFormBTamilNaduBasicPlusHra(basic, hra);
  const other = roundFormBMoney(g - known);
  if (!Number.isFinite(other)) return '';
  // Guard against missing basic/hra accidentally dumping full gross as "other", and negatives.
  if (other < 0) return 0;
  return other;
}

/**
 * Balance due to the employees.
 * Rule: if (gross_pay − net_pay) equals net_pay, put 0.
 * When net_pay is present (amount actually paid), remaining due is also 0.
 */
export function computeFormBTamilNaduBalanceDue(totalGrossPay, totalNetPay) {
  const g = parseFormBTamilNaduMoney(totalGrossPay);
  const n = parseFormBTamilNaduMoney(totalNetPay);
  if (!Number.isFinite(g) && !Number.isFinite(n)) return '';
  const gross = Number.isFinite(g) ? g : 0;
  const net = Number.isFinite(n) ? n : 0;
  const diff = roundFormBMoney(gross - net);
  if (diff === roundFormBMoney(net)) return 0;
  // Wages paid via net_pay → nothing remains due to employees.
  if (Number.isFinite(n)) return 0;
  return diff > 0 ? diff : 0;
}

/** Total emoluments payable … including basic wages, D.A, O.T., and bonus */
export function isFormBTamilNaduTotalEmolumentsHeader(h) {
  const s = formBTamilNaduHeaderNorm(h);
  if (s.includes('emolument')) {
    return s.includes('payable') || s.includes('including') || s.includes('basic wage');
  }
  // Merged Excel titles sometimes omit the word "emoluments" in the stored key fragment.
  return (
    s.includes('basic wage') &&
    (s.includes('bonus') || /\bd\.?a\.?\b/.test(s) || s.includes('o.t')) &&
    (s.includes('payable') || s.includes('including') || s.includes('during the month'))
  );
}

export const FORM_B_TN_AMOUNTS_DEDUCTED_GROUP_LABEL = 'Amounts deducted during the month';

/** Parent banner over Fine + Other deductions (not a leaf column). */
export function isFormBTamilNaduAmountsDeductedGroupLabel(text) {
  const s = formBTamilNaduHeaderNorm(text);
  if (!s) return false;
  if (/other\s+deductions?/.test(s) && !/amounts?\s+deducted/.test(s)) return false;
  return /amounts?\s+deducted\s+during\s+the\s+month/.test(s);
}

function formBTamilNaduPdfLeafAt(rows, fromRow, toRow, col) {
  const list = Array.isArray(rows) ? rows : [];
  for (let r = fromRow; r <= toRow && r < list.length; r += 1) {
    const t = String(list[r]?.[col] || '')
      .replace(/\s+/g, ' ')
      .trim();
    if (!t) continue;
    if (isFormBTamilNaduAmountsDeductedGroupLabel(t)) continue;
    return t;
  }
  return '';
}

/**
 * Re-span the Excel-merged "Amounts deducted during the month" banner across
 * Fine + Other deductions so the full heading stays visible in PDF.
 */
export function detectFormBTamilNaduPdfGroupBands(rows, tableStart, headerBandEnd, colCount) {
  const list = Array.isArray(rows) ? rows : [];
  const start = Math.max(0, Number(tableStart) || 0);
  const end = Math.max(start, Number(headerBandEnd) || start);
  const cols = Math.max(0, Number(colCount) || 0);
  if (!list.length || cols < 3) return [];

  let labelRow = -1;
  let labelCol = -1;
  let label = FORM_B_TN_AMOUNTS_DEDUCTED_GROUP_LABEL;
  for (let r = start; r <= end && r < list.length; r += 1) {
    for (let c = 0; c < cols; c += 1) {
      const t = String(list[r]?.[c] || '')
        .replace(/\s+/g, ' ')
        .trim();
      if (!isFormBTamilNaduAmountsDeductedGroupLabel(t)) continue;
      labelRow = r;
      labelCol = c;
      label = t;
    }
  }
  if (labelRow < 0 || labelCol < 0) return [];

  let leafEnd = labelCol;
  for (let c = labelCol; c < cols; c += 1) {
    const sameRow = String(list[labelRow]?.[c] || '')
      .replace(/\s+/g, ' ')
      .trim();
    if (c > labelCol && sameRow && !isFormBTamilNaduAmountsDeductedGroupLabel(sameRow)) {
      break;
    }
    const leaf = formBTamilNaduPdfLeafAt(list, labelRow + 1, end, c);
    if (
      c > labelCol &&
      /amount\s+actually\s+paid|balance\s+due|total\s+emolument|total\s+number/i.test(leaf)
    ) {
      break;
    }
    if (
      c === labelCol ||
      !leaf ||
      /\bfine\b/i.test(leaf) ||
      /other\s+deductions?/i.test(leaf) ||
      /^\(?\s*[ab]\s*\)?$/i.test(leaf)
    ) {
      leafEnd = c;
      continue;
    }
    break;
  }
  if (leafEnd <= labelCol) leafEnd = Math.min(labelCol + 1, cols - 1);
  return [
    {
      labelRow,
      start: labelCol,
      end: leafEnd,
      label: FORM_B_TN_AMOUNTS_DEDUCTED_GROUP_LABEL,
    },
  ];
}

/**
 * Keep Fine blank and Other deductions in the Other column.
 * Collapsed Excel merges / skipped empty Fine cells can slide 458439 into Fine.
 */
export function normalizeFormBTamilNaduPdfMatrix(rows, colCount, tableStart = 0, metaLines = []) {
  const list = (Array.isArray(rows) ? rows : []).map((row) => (Array.isArray(row) ? [...row] : row));
  const cols = Math.max(
    Number(colCount) || 0,
    ...list.map((row) => (Array.isArray(row) ? row.length : 0)),
    6
  );
  const start = Math.max(0, Number(tableStart) || 0);

  let fineCol = -1;
  let otherCol = -1;
  let paidCol = -1;
  let headerEnd = start;
  for (let r = start; r < Math.min(list.length, start + 6); r += 1) {
    const row = list[r] || [];
    for (let c = 0; c < cols; c += 1) {
      const t = String(row[c] || '').replace(/\s+/g, ' ').trim();
      if (!t) continue;
      if (isFormBTamilNaduFineHeader(t) && !isFormBTamilNaduOtherDeductionsHeader(t)) {
        fineCol = c;
        headerEnd = Math.max(headerEnd, r);
      }
      if (isFormBTamilNaduOtherDeductionsHeader(t)) {
        otherCol = c;
        headerEnd = Math.max(headerEnd, r);
      }
      if (isFormBTamilNaduAmountActuallyPaidHeader(t)) {
        paidCol = c;
        headerEnd = Math.max(headerEnd, r);
      }
      if (isFormBTamilNaduAmountsDeductedGroupLabel(t) && fineCol < 0) {
        fineCol = c;
        if (otherCol < 0) otherCol = c + 1;
        headerEnd = Math.max(headerEnd, r);
      }
    }
  }
  if (fineCol < 0) fineCol = 2;
  if (otherCol < 0) otherCol = fineCol + 1;
  if (paidCol < 0) paidCol = otherCol + 1;

  const isHeaderishRow = (row) => {
    const filled = (row || []).map((c) => String(c || '').trim()).filter(Boolean);
    if (!filled.length) return false;
    const blob = filled.join(' ').toLowerCase();
    if (/total\s+number\s+of\s+empl|emolument|amounts?\s+deducted|actually\s+paid|balance\s+due/.test(blob)) {
      return true;
    }
    if (filled.every((t) => /^(?:\(?\s*[ab]\s*\)?|\(?\s*\d{1,2}\s*\)?|fine|other\s+deductions?)$/i.test(t))) {
      return true;
    }
    return false;
  };

  for (let r = headerEnd + 1; r < list.length; r += 1) {
    const row = list[r];
    if (!Array.isArray(row)) continue;
    if (isFormBTamilNaduPdfSignatoryRow(row)) continue;
    if (isHeaderishRow(row)) continue;
    while (row.length < cols) row.push('');

    const fineVal = String(row[fineCol] ?? '').trim();
    const otherVal = String(row[otherCol] ?? '').trim();
    const paidVal = String(row[paidCol] ?? '').trim();
    const fineNum = parseFormBTamilNaduMoney(fineVal);
    const otherNum = parseFormBTamilNaduMoney(otherVal);
    const paidNum = parseFormBTamilNaduMoney(paidVal);
    const hasAmounts = [fineNum, otherNum, paidNum].some((n) => Number.isFinite(n));
    if (!hasAmounts && !fineVal && !otherVal) continue;

    // Empty Fine collapsed → Other holds Fine's slot and Amount paid slid left.
    if (
      (!otherVal || otherNum === 0) &&
      Number.isFinite(fineNum) &&
      fineNum !== 0 &&
      Number.isFinite(paidNum)
    ) {
      row[otherCol] = fineVal;
      row[fineCol] = '';
    } else if ((!otherVal || otherNum === 0) && Number.isFinite(fineNum) && fineNum !== 0) {
      row[otherCol] = fineVal;
      row[fineCol] = '';
    }

    const nextFine = String(row[fineCol] ?? '').trim();
    const nextFineNum = parseFormBTamilNaduMoney(nextFine);
    if (nextFine === '0' || nextFineNum === 0) {
      row[fineCol] = '';
    }
  }

  return {
    rows: list,
    colCount: cols,
    tableStartRow: start,
    metaLines,
  };
}

/** Keep Fine / Other wide enough for the parent "Amounts deducted during the month" banner. */
export function formBTamilNaduPdfColumnWeight(headerText) {
  const h = formBTamilNaduHeaderNorm(headerText);
  if (/total\s+number\s+of\s+empl/.test(h)) return 7.5;
  if (/emolument|basic\s+wage/.test(h)) return 13;
  if (/\bfine\b/.test(h) && !/other/.test(h)) return 10;
  if (/other\s+deductions?/.test(h)) return 10;
  if (/amounts?\s+deducted/.test(h)) return 10;
  if (/actually\s+paid/.test(h)) return 10.5;
  if (/balance\s+due/.test(h)) return 9.5;
  return 8;
}

/** "Amounts deducted during the month" → Other deductions sub-column only. */
export function isFormBTamilNaduOtherDeductionsHeader(h) {
  const s = formBTamilNaduHeaderNorm(h);
  return /other\s+deductions?/.test(s) || /deducted.*other\s+deductions?/.test(s);
}

/** Amount actually paid during the month */
export function isFormBTamilNaduAmountActuallyPaidHeader(h) {
  const s = formBTamilNaduHeaderNorm(h);
  if (/amount\s+actually\s+paid/.test(s)) return true;
  if (s.includes('actually paid') && s.includes('month')) return true;
  return s.includes('actually paid') && !s.includes('balance');
}

/** Balance due to the employees */
export function isFormBTamilNaduBalanceDueHeader(h) {
  const s = formBTamilNaduHeaderNorm(h);
  return (
    /balance\s+due\s+to\s+the\s+employees?/.test(s) ||
    (s.includes('balance') && s.includes('due') && s.includes('employee'))
  );
}

/** Payroll-driven Form B summary columns — never restore from statutory overlay. */
export function isFormBTamilNaduPayrollSummaryHeader(h) {
  return (
    isFormBTamilNaduTotalEmolumentsHeader(h) ||
    isFormBTamilNaduOtherDeductionsHeader(h) ||
    isFormBTamilNaduAmountActuallyPaidHeader(h) ||
    isFormBTamilNaduBalanceDueHeader(h) ||
    isFormBTamilNaduFineHeader(h) ||
    isFormBTamilNaduTotalEmployeeCountHeader(h)
  );
}

/** Total number of employees */
export function isFormBTamilNaduTotalEmployeeCountHeader(h) {
  const s = formBTamilNaduHeaderNorm(h);
  if (/total\s+number\s+of\s+empl/.test(s)) return true;
  if (/total\s+number\s+of\s+employees?/.test(s)) return true;
  return s.includes('total') && s.includes('number') && (s.includes('employee') || s.includes('empl'));
}

/** Fine sub-column under Amounts deducted (never filled from payroll summary). */
export function isFormBTamilNaduFineHeader(h) {
  const s = formBTamilNaduHeaderNorm(h);
  if (/other\s+deductions?/.test(s)) return false;
  return /(^|[\s(_-])fine(\s|$|[)_-])/.test(s) || /^fine\b/.test(s) || /\bfine\s*\(?\s*a\s*\)?/.test(s);
}

/**
 * Pull Form B export values from an autofill summary row (numeric-only OK).
 * Maps by header semantics so Parent_Sub / display labels both work.
 */
export function extractFormBTamilNaduSummaryExportValues(row, headers = []) {
  const empty = {
    employeeCount: '',
    totalEmoluments: '',
    fine: '',
    otherDeductions: '',
    amountActuallyPaid: '',
    balanceDue: '',
  };
  if (!row || typeof row !== 'object') return empty;
  const hdrs =
    Array.isArray(headers) && headers.length > 0 ? headers : Object.keys(row);

  const pickBy = (matcher) => {
    for (let i = 0; i < hdrs.length; i += 1) {
      const h = hdrs[i];
      if (!matcher(h)) continue;
      const v = row[h];
      if (v == null || String(v).trim() === '') continue;
      return String(v).trim();
    }
    // Fallback: scan all own keys (row may use different header strings than headers[]).
    for (const h of Object.keys(row)) {
      if (!matcher(h)) continue;
      const v = row[h];
      if (v == null || String(v).trim() === '') continue;
      return String(v).trim();
    }
    return '';
  };

  return {
    employeeCount: pickBy(isFormBTamilNaduTotalEmployeeCountHeader),
    totalEmoluments: pickBy(isFormBTamilNaduTotalEmolumentsHeader),
    fine: pickBy(isFormBTamilNaduFineHeader),
    otherDeductions: pickBy(isFormBTamilNaduOtherDeductionsHeader),
    amountActuallyPaid: pickBy(isFormBTamilNaduAmountActuallyPaidHeader),
    balanceDue: pickBy(isFormBTamilNaduBalanceDueHeader),
  };
}

export function formBTamilNaduSummaryHasExportAmounts(values) {
  if (!values || typeof values !== 'object') return false;
  return [values.totalEmoluments, values.otherDeductions, values.amountActuallyPaid, values.balanceDue].some(
    (v) => v != null && String(v).trim() !== ''
  );
}

/**
 * Aggregate payroll amounts for the Form B summary row.
 * Total emoluments = sum(gross_pay); Amount actually paid = sum(net_pay);
 * Balance due = 0 when (gross − net) equals net (or wages were paid).
 * @param {Array<{ gross?: *, basic?: *, hra?: *, net?: * }>} amounts
 */
export function summarizeFormBTamilNaduPayrollAmounts(amounts) {
  let totalGrossPay = 0;
  let totalBasicPlusHra = 0;
  let totalOtherDeductions = 0;
  let totalNetPay = 0;
  let matchedCount = 0;
  let emolumentsEmployeeCount = 0;

  (Array.isArray(amounts) ? amounts : []).forEach((a) => {
    if (!a || typeof a !== 'object') return;
    const gross = parseFormBTamilNaduMoney(a.gross);
    const basicNum = parseFormBTamilNaduMoney(a.basic);
    const hraNum = parseFormBTamilNaduMoney(a.hra);
    const hasBasicOrHra = Number.isFinite(basicNum) || Number.isFinite(hraNum);

    // Keep basic+hra for Other deductions (gross − basic − hra).
    if (hasBasicOrHra) {
      totalBasicPlusHra += computeFormBTamilNaduBasicPlusHra(a.basic, a.hra);
    }

    if (!Number.isFinite(gross)) return;
    matchedCount += 1;
    emolumentsEmployeeCount += 1;
    totalGrossPay += gross;
    const other = computeFormBTamilNaduOtherDeductions(a.gross, a.basic, a.hra);
    if (other !== '') totalOtherDeductions += other;
    const net = parseFormBTamilNaduMoney(a.net);
    if (Number.isFinite(net)) totalNetPay += net;
  });

  const grossTotal = roundFormBMoney(totalGrossPay);
  const netTotal = roundFormBMoney(totalNetPay);
  const basicHraTotal = roundFormBMoney(totalBasicPlusHra);
  return {
    matchedCount,
    payrollMatchedCount: matchedCount,
    emolumentsEmployeeCount,
    totalGrossPay: grossTotal,
    totalBasicPlusHra: basicHraTotal,
    /** Total emoluments = sum(gross_pay) for corresponding employees. */
    totalEmoluments: grossTotal,
    totalOtherDeductions: roundFormBMoney(totalOtherDeductions),
    totalNetPay: netTotal,
    /** Amount actually paid = sum(net_pay). */
    totalAmountActuallyPaid: netTotal,
    /** Balance due — 0 when paid / when (gross − net) equals net. */
    totalBalanceDue: computeFormBTamilNaduBalanceDue(grossTotal, netTotal),
  };
}

export const FORM_B_TN_ESTABLISHMENT_LABEL = 'Name of the Establishment';
export const FORM_B_TN_MONTH_LABEL = 'For the Month of';

const FORM_B_TN_MONTH_NAME_RE =
  /^(january|february|march|april|may|june|july|august|september|october|november|december)$/i;

function formBTamilNaduExcelCellText(val) {
  if (val == null) return '';
  if (typeof val === 'string' || typeof val === 'number' || typeof val === 'boolean') {
    return String(val).replace(/\r?\n/g, ' ').replace(/\s+/g, ' ').trim();
  }
  if (val instanceof Date) return val.toISOString();
  if (typeof val === 'object') {
    if (Array.isArray(val.richText)) {
      return val.richText
        .map((rt) => rt?.text || '')
        .join('')
        .replace(/\r?\n/g, ' ')
        .replace(/\s+/g, ' ')
        .trim();
    }
    if (val.text != null) return String(val.text).replace(/\r?\n/g, ' ').replace(/\s+/g, ' ').trim();
    if (val.result != null) return formBTamilNaduExcelCellText(val.result);
  }
  return '';
}

function titleCaseMonthName(name) {
  const s = String(name || '').trim();
  if (!s) return '';
  return `${s.charAt(0).toUpperCase()}${s.slice(1).toLowerCase()}`;
}

/** "July2026" / "july 2026" → "July 2026". */
export function formatFormBTamilNaduMonthDisplay(value) {
  const raw = String(value || '')
    .replace(/\s+/g, ' ')
    .trim();
  if (!raw) return '';
  const glued = raw.match(
    /^(january|february|march|april|may|june|july|august|september|october|november|december)[-/.\s]*(\d{4})$/i
  );
  if (glued) return `${titleCaseMonthName(glued[1])} ${glued[2]}`;
  return raw;
}

export function formatFormBTamilNaduEstablishmentLine(value) {
  const val = String(value || '').replace(/\s+/g, ' ').trim();
  if (!val) return `${FORM_B_TN_ESTABLISHMENT_LABEL} :`;
  if (/^name\s+of\s+the\s+establishment\s*:/i.test(val)) {
    const rest = val.replace(/^name\s+of\s+the\s+establishment\s*:?\s*/i, '').trim();
    return rest ? `${FORM_B_TN_ESTABLISHMENT_LABEL} : ${rest}` : `${FORM_B_TN_ESTABLISHMENT_LABEL} :`;
  }
  return `${FORM_B_TN_ESTABLISHMENT_LABEL} : ${val}`;
}

export function formatFormBTamilNaduMonthLine(value) {
  const month = formatFormBTamilNaduMonthDisplay(
    String(value || '').replace(/^for\s+the\s+month\s+of\s*:?\s*/i, '')
  );
  if (!month) return `${FORM_B_TN_MONTH_LABEL} :`;
  return `${FORM_B_TN_MONTH_LABEL} : ${month}`;
}

function looksLikeFormBTamilNaduMonthValue(value) {
  const s = formatFormBTamilNaduMonthDisplay(value);
  if (!s) return false;
  if (FORM_B_TN_MONTH_NAME_RE.test(s)) return true;
  return /^(january|february|march|april|may|june|july|august|september|october|november|december)\s+\d{4}$/i.test(
    s
  );
}

/** Pull establishment + month from headerFormData without grabbing wage/table keys. */
export function resolveFormBTamilNaduHeaderExportValues(headerFormData) {
  const data = headerFormData && typeof headerFormData === 'object' ? headerFormData : {};
  const pick = (keys) => {
    for (let i = 0; i < keys.length; i += 1) {
      const v = data[keys[i]];
      if (v != null && String(v).trim() !== '') return String(v).trim();
    }
    return '';
  };
  let establishment = pick([
    'statutory_establishment_name_address',
    'statutory_establishment_name',
    'statutory_establishment_name_shop',
    'form_q_establishment',
    'form_a_establishment_name',
    'form25_establishment',
  ]);
  if (!establishment) {
    const entries = Object.entries(data);
    for (let i = 0; i < entries.length; i += 1) {
      const [k, v] = entries[i];
      if (!/establishment/i.test(String(k || ''))) continue;
      if (v == null || String(v).trim() === '') continue;
      establishment = String(v).trim();
      break;
    }
  }
  let month = pick(['form_b_header_for_the_month_of', 'form_xvi_for_the_month_of']);
  if (!month) {
    const entries = Object.entries(data);
    for (let i = 0; i < entries.length; i += 1) {
      const [k, v] = entries[i];
      if (!/for[_ ]the[_ ]month[_ ]of/i.test(String(k || ''))) continue;
      if (v == null || String(v).trim() === '') continue;
      month = String(v).trim();
      break;
    }
  }
  if (!month) {
    const entries = Object.entries(data);
    for (let i = 0; i < entries.length; i += 1) {
      const [k, v] = entries[i];
      if (!/month/i.test(String(k || ''))) continue;
      if (/amount|emolument|wage|header_row|table/i.test(String(k || ''))) continue;
      if (!looksLikeFormBTamilNaduMonthValue(v)) continue;
      month = String(v).trim();
      break;
    }
  }
  return {
    establishment,
    month: formatFormBTamilNaduMonthDisplay(month),
  };
}

/**
 * Write "Name of the Establishment :" and "For the Month of :" onto the template
 * label cells (same cell — never overwrite a merged label with the value alone).
 */
export function writeFormBTamilNaduHeaderFieldsToWorksheet(worksheet, headerFormData) {
  if (!worksheet || typeof worksheet.getCell !== 'function') return false;
  const { establishment, month } = resolveFormBTamilNaduHeaderExportValues(headerFormData);
  if (!establishment && !month) return false;

  const maxRows = Math.min(12, Math.max(8, Number(worksheet.rowCount) || 8));
  const maxCols = Math.min(12, Math.max(6, Number(worksheet.columnCount) || 6));
  let wroteEst = false;
  let wroteMonth = false;

  const parseMerge = (label) => {
    const parts = String(label || '').split(':');
    if (parts.length !== 2) return null;
    const start = parts[0].match(/^([A-Z]+)(\d+)$/i);
    const end = parts[1].match(/^([A-Z]+)(\d+)$/i);
    if (!start || !end) return null;
    const colToNum = (letters) => {
      let n = 0;
      const s = String(letters || '').toUpperCase();
      for (let i = 0; i < s.length; i += 1) n = n * 26 + (s.charCodeAt(i) - 64);
      return n;
    };
    return {
      r1: parseInt(start[2], 10),
      r2: parseInt(end[2], 10),
      c1: colToNum(start[1]),
      c2: colToNum(end[1]),
    };
  };

  const mergeMaster = (row, col) => {
    const merges = Array.isArray(worksheet.model?.merges) ? worksheet.model.merges : [];
    for (let i = 0; i < merges.length; i += 1) {
      const m = parseMerge(merges[i]);
      if (!m) continue;
      if (row >= m.r1 && row <= m.r2 && col >= m.c1 && col <= m.c2) {
        return { row: m.r1, col: m.c1, c2: Math.max(m.c2, 6) };
      }
    }
    return { row, col, c2: Math.max(col, 6) };
  };

  const writeLine = (row, col, text) => {
    const master = mergeMaster(row, col);
    const endCol = Math.max(master.c2, 6);
    // Unmerge the real range first (A:D), then remesh across the full table (A:F)
    // so "Name of the Establishment : Vayona Energy Pvt Ltd, …" is not clipped.
    if (master.c2 > master.col) {
      try {
        worksheet.unMergeCells(master.row, master.col, master.row, master.c2);
      } catch (_) {
        /* ignore */
      }
    }
    if (endCol > master.col) {
      try {
        worksheet.mergeCells(master.row, master.col, master.row, endCol);
      } catch (_) {
        /* ignore */
      }
    }
    const cell = worksheet.getCell(master.row, master.col);
    cell.value = text;
    cell.alignment = {
      ...(cell.alignment || {}),
      wrapText: true,
      vertical: 'middle',
      horizontal: 'left',
    };
    const wsRow = worksheet.getRow(master.row);
    if (wsRow) {
      const colSpan = Math.max(1, endCol - master.col + 1);
      const charsPerLine = Math.max(28, colSpan * 14);
      const lines = Math.max(2, Math.ceil(String(text).length / charsPerLine));
      wsRow.height = Math.max(Number(wsRow.height) || 0, lines * 18, 48);
    }
    applyExcelJSFullBoxBordersToRange(worksheet, {
      rowFrom: master.row,
      rowTo: master.row,
      colFrom: master.col,
      colTo: endCol,
    });
  };

  for (let r = 1; r <= maxRows; r += 1) {
    for (let c = 1; c <= maxCols; c += 1) {
      const raw = formBTamilNaduExcelCellText(worksheet.getCell(r, c)?.value);
      if (!raw) continue;
      if (!wroteEst && establishment && /name\s+of\s+the\s+establishment/i.test(raw)) {
        writeLine(r, c, formatFormBTamilNaduEstablishmentLine(establishment));
        wroteEst = true;
        continue;
      }
      if (!wroteMonth && month && /for\s+the\s+month\s+of/i.test(raw)) {
        writeLine(r, c, formatFormBTamilNaduMonthLine(month));
        wroteMonth = true;
      }
    }
  }

  // Re-export of a filled sheet may have already dropped the labels.
  if (establishment && !wroteEst) {
    for (let r = 4; r <= Math.min(7, maxRows); r += 1) {
      const raw = formBTamilNaduExcelCellText(worksheet.getCell(r, 1)?.value);
      if (!raw) continue;
      if (/form\s*b|see\s+rule|register\s+of\s+wages|labour\s+welfare/i.test(raw)) continue;
      if (/for\s+the\s+month\s+of/i.test(raw)) continue;
      writeLine(r, 1, formatFormBTamilNaduEstablishmentLine(establishment));
      wroteEst = true;
      break;
    }
  }
  if (month && !wroteMonth) {
    for (let r = 5; r <= Math.min(8, maxRows); r += 1) {
      const raw = formBTamilNaduExcelCellText(worksheet.getCell(r, 1)?.value);
      if (raw && /name\s+of\s+the\s+establishment/i.test(raw)) continue;
      if (raw && /form\s*b|see\s+rule|register\s+of\s+wages|labour\s+welfare/i.test(raw)) continue;
      if (!raw || looksLikeFormBTamilNaduMonthValue(raw) || /for\s+the\s+month\s+of/i.test(raw)) {
        writeLine(r, 1, formatFormBTamilNaduMonthLine(month));
        wroteMonth = true;
        break;
      }
    }
  }
  // Close the register header stack so remeshed establishment/month rows stay boxed.
  if (wroteEst || wroteMonth) {
    applyExcelJSFullBoxBordersToRange(worksheet, {
      rowFrom: 1,
      rowTo: 6,
      colFrom: 1,
      colTo: 6,
    });
  }
  return wroteEst || wroteMonth;
}

export function looksLikeFormBTamilNaduPdfContext(metaLines, rows, sheetName = '', fileName = '') {
  const blob = [...(metaLines || []), ...(rows || []).slice(0, 18).flat(), sheetName || '', fileName || '']
    .join(' ')
    .toLowerCase();
  if (/form[_\s-]*b[_\s-]*(gj|gujarat|rj|rajasthan)/i.test(blob)) return false;
  if (/rate\s+of\s+minimum\s+wages/.test(blob) && /gujarat|rajasthan/.test(blob)) return false;
  const hasFormB = /\bform\s*b\b|form[_\s-]*b(?:[_\s-]|$)/i.test(blob);
  const hasTamilLwf =
    /tamil\s*nadu\s+labour\s+welfare/.test(blob) ||
    (/labour\s+welfare/.test(blob) && /tamil/.test(blob)) ||
    /form[_\s-]*b[_\s-]*tamil/.test(blob);
  const hasRegister = /register\s+of\s+wages/.test(blob);
  const hasSummary =
    /total\s+emoluments/.test(blob) ||
    /total\s+number\s+of\s+empl/.test(blob) ||
    /amounts?\s+deducted/.test(blob);
  if (hasFormB && hasTamilLwf && (hasRegister || hasSummary)) return true;
  if (/form[_\s-]*b[_\s-]*tamil/.test(blob)) return true;
  return false;
}

export function isFormBTamilNaduPdfSignatoryText(text) {
  const flat = String(text || '')
    .replace(/_x000d_/gi, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  if (!flat) return false;
  if (/^for\s*\(/i.test(flat)) return true;
  if (/authori[sz]ed\s+signatory/i.test(flat)) return true;
  if (/signature\s+of\s+employer/i.test(flat)) return true;
  if (/manager\s*\/\s*authori[sz]ed\s+person/i.test(flat)) return true;
  return false;
}

export function isFormBTamilNaduPdfSignatoryRow(row) {
  if (!Array.isArray(row)) return isFormBTamilNaduPdfSignatoryText(row);
  const filled = row.map((c) => String(c || '').replace(/\s+/g, ' ').trim()).filter(Boolean);
  if (!filled.length) return false;
  return filled.every((t) => isFormBTamilNaduPdfSignatoryText(t));
}

function stripFormBTamilNaduFieldValue(line, labelRe) {
  const text = String(line || '').replace(/\s+/g, ' ').trim();
  if (!text) return '';
  if (labelRe.test(text) && /:/.test(text)) {
    return text.replace(labelRe, '').replace(/^[:\s]+/, '').trim();
  }
  return text;
}

/**
 * Keep Form B TN titles (Form B / See rule 29 / LWF Rules / Register of Wages)
 * and force establishment + month field labels. Drop signatory footer lines.
 */
export function rewriteFormBTamilNaduPdfHeader(titles, fields) {
  const titleList = Array.isArray(titles) ? titles : [];
  const fieldList = Array.isArray(fields) ? fields : [];
  const nextTitles = [];
  const leftover = [];
  let establishment = '';
  let month = '';

  const consider = (raw, asTitle) => {
    const line = String(raw || '').replace(/\s+/g, ' ').trim();
    if (!line) return;
    if (isFormBTamilNaduPdfSignatoryText(line)) return;
    if (/name\s+of\s+the\s+establishment/i.test(line)) {
      const val = stripFormBTamilNaduFieldValue(line, /^name\s+of\s+the\s+establishment\s*:?\s*/i);
      if (val && !/^name\s+of\s+the\s+establishment$/i.test(val)) establishment = val;
      return;
    }
    if (/for\s+the\s+month\s+of/i.test(line) || looksLikeFormBTamilNaduMonthValue(line)) {
      const val = formatFormBTamilNaduMonthDisplay(
        stripFormBTamilNaduFieldValue(line, /^for\s+the\s+month\s+of\s*:?\s*/i)
      );
      if (val) month = val;
      return;
    }
    if (asTitle) {
      if (/pvt\.?\s*ltd|private\s+limited|energy|road|district|taluk/i.test(line)) {
        leftover.push(line);
        return;
      }
      nextTitles.push(line);
      return;
    }
    leftover.push(line);
  };

  titleList.forEach((t) => consider(t, true));
  fieldList.forEach((f) => consider(f, false));

  if (!establishment) {
    const companyLike = leftover.find((l) =>
      /pvt\.?\s*ltd|private\s+limited|energy|road|district|taluk/i.test(l)
    );
    if (companyLike) {
      establishment = companyLike;
    }
  }

  const nextFields = [
    formatFormBTamilNaduEstablishmentLine(establishment),
    formatFormBTamilNaduMonthLine(month),
  ];
  leftover.forEach((l) => {
    if (/pvt\.?\s*ltd|private\s+limited|energy|road|district|taluk/i.test(l) && establishment) {
      return;
    }
    if (isFormBTamilNaduPdfSignatoryText(l)) return;
    if (looksLikeFormBTamilNaduMonthValue(l)) return;
    nextFields.push(l);
  });

  const seen = new Set();
  const uniqueFields = nextFields.filter((f) => {
    const key = String(f || '').toLowerCase();
    if (!key || seen.has(key)) return false;
    seen.add(key);
    return true;
  });

  return { titles: nextTitles, fields: uniqueFields };
}

/**
 * Write summary totals onto the single Form B row using table headers.
 */
export function applyFormBTamilNaduSummaryTotals(summaryRow, headers, totals, options = {}) {
  if (!summaryRow || typeof summaryRow !== 'object') return summaryRow;
  const list = Array.isArray(headers) ? headers : [];
  const t = totals || {};
  const sanitize =
    typeof options.sanitizeValue === 'function' ? options.sanitizeValue : (v) => String(v ?? '');
  const hasMatch = (t.matchedCount || 0) > 0;

  const emolHeader = list.find((h) => isFormBTamilNaduTotalEmolumentsHeader(h));
  const otherDedHeader = list.find((h) => isFormBTamilNaduOtherDeductionsHeader(h));
  const paidHeader = list.find((h) => isFormBTamilNaduAmountActuallyPaidHeader(h));
  const balanceHeader = list.find((h) => isFormBTamilNaduBalanceDueHeader(h));
  const emolumentsTotal =
    t.totalEmoluments != null && t.totalEmoluments !== ''
      ? t.totalEmoluments
      : t.totalGrossPay;
  const paidTotal =
    t.totalAmountActuallyPaid != null && t.totalAmountActuallyPaid !== ''
      ? t.totalAmountActuallyPaid
      : t.totalNetPay;
  const balanceTotal =
    t.totalBalanceDue != null && t.totalBalanceDue !== ''
      ? t.totalBalanceDue
      : computeFormBTamilNaduBalanceDue(emolumentsTotal, paidTotal);

  if (emolHeader && (hasMatch || emolumentsTotal > 0 || t.totalGrossPay > 0)) {
    // Total emoluments = sum of gross_pay.
    summaryRow[emolHeader] = sanitize(emolumentsTotal);
  }
  if (otherDedHeader && (hasMatch || t.totalOtherDeductions > 0 || t.totalOtherDeductions === 0)) {
    summaryRow[otherDedHeader] = sanitize(t.totalOtherDeductions);
  }
  if (paidHeader && (hasMatch || paidTotal > 0)) {
    // Amount actually paid = sum of net_pay.
    summaryRow[paidHeader] = sanitize(paidTotal);
  }
  if (balanceHeader && (hasMatch || paidTotal > 0 || emolumentsTotal > 0 || balanceTotal === 0)) {
    summaryRow[balanceHeader] = sanitize(balanceTotal === '' ? 0 : balanceTotal);
  }
  return summaryRow;
}
