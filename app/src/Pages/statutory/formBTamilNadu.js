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
