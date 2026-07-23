import {
  readForm10GrossPayAmount,
  readPayrollForm15WageAmounts,
} from '../../utils/payrollEarnings';

/** Form XXIII Tamil Nadu — Register of Overtime [See Rule 78(1)(a)(iii)]. */

/** Form XXIII TN overtime columns that always default to NIL (no attendance/payroll OT). */
export const FORM_XXIII_TN_OT_NIL = 'NIL';

const normFormXXIIITamilNaduHeader = (h) =>
  String(h || '')
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();

export function isFormXXIIITamilNaduContext(formHeader, rowItem, fileName, sheetText = '') {
  const parts = [
    rowItem?.formName,
    rowItem?.FormName,
    rowItem?.state,
    rowItem?.State,
    rowItem?.siteState,
    rowItem?.SiteState,
    rowItem?.formFileName,
    rowItem?.FormFileName,
    fileName,
    formHeader?.title,
    formHeader?.subtitle,
    formHeader?.reference,
    sheetText,
  ]
    .filter((x) => x != null && String(x).trim() !== '')
    .join(' ')
    .toLowerCase();

  const hasTN =
    /tamil[\s._-]*nadu|tamilnadu|\(tn\)|\btn\b/.test(parts) ||
    /\bform[\s._-]*xxiii[\s._-]*tamil/.test(parts) ||
    /\bxxiii[\s._-]*tamil/.test(parts);
  const hasXXIII =
    /\bform[\s._-]*xxiii/i.test(parts) || /register\s+of\s+overtime/.test(parts);
  return hasTN && hasXXIII;
}

/** Dates on which overtime worked */
export function isFormXXIIITamilNaduOtWorkedDatesHeader(h) {
  const s = normFormXXIIITamilNaduHeader(h);
  if (!s) return false;
  if (/payment|paid|rate|wage|earning/.test(s) && !/worked/.test(s)) return false;
  return (
    /dates?\s+on\s+which\s+over[\s-]*time\s+worked/.test(s) ||
    (s.includes('date') && s.includes('overtime') && s.includes('worked') && !/paid|payment/.test(s))
  );
}

/** Total overtime worked or production in case of piece rate */
export function isFormXXIIITamilNaduTotalOvertimeWorkedHeader(h) {
  const s = normFormXXIIITamilNaduHeader(h);
  if (!s) return false;
  return (
    (s.includes('total') && s.includes('overtime') && (s.includes('worked') || s.includes('production'))) ||
    (s.includes('overtime') && s.includes('piece') && s.includes('rate'))
  );
}

/** Normal rate of wages ← SamplePayroll gross_pay */
export function isFormXXIIITamilNaduNormalRateHeader(h) {
  const s = normFormXXIIITamilNaduHeader(h);
  if (!s || /over[\s-]*time/.test(s)) return false;
  return s.includes('normal') && s.includes('rate') && (s.includes('wage') || s.includes('pay'));
}

/** Overtime rate of wages ← (Basic / 26 / 8) * 2 */
export function isFormXXIIITamilNaduOvertimeRateHeader(h) {
  const s = normFormXXIIITamilNaduHeader(h);
  if (!s || /normal/.test(s)) return false;
  return s.includes('overtime') && s.includes('rate') && (s.includes('wage') || s.includes('pay'));
}

/** Overtime earnings */
export function isFormXXIIITamilNaduOvertimeEarningsHeader(h) {
  const s = normFormXXIIITamilNaduHeader(h);
  return s.includes('overtime') && s.includes('earning');
}

/** Date on which overtime wages paid */
export function isFormXXIIITamilNaduOtWagesPaidDateHeader(h) {
  const s = normFormXXIIITamilNaduHeader(h);
  return (
    (s.includes('overtime') && s.includes('paid')) ||
    /dates?\s+on\s+which\s+overtime\s+wage/.test(s)
  );
}

export function isFormXXIIITamilNaduOtNilHeader(header) {
  return (
    isFormXXIIITamilNaduOtWorkedDatesHeader(header) ||
    isFormXXIIITamilNaduTotalOvertimeWorkedHeader(header) ||
    isFormXXIIITamilNaduOvertimeEarningsHeader(header) ||
    isFormXXIIITamilNaduOtWagesPaidDateHeader(header)
  );
}

/** Normal rate of wages ← gross_pay. */
export function resolveFormXXIIITamilNaduNormalRate(payrollRow) {
  if (!payrollRow || payrollRow.fetch_error) return '';
  const gross = readForm10GrossPayAmount(payrollRow);
  return gross === '' || gross == null ? '' : String(gross);
}

/**
 * Overtime rate of wages = (Basic / 26 / 8) * 2
 * Basic from payroll wage breakdown (basic / earned_basic).
 */
export function resolveFormXXIIITamilNaduOvertimeRate(payrollRow) {
  if (!payrollRow || payrollRow.fetch_error) return '';
  const { basic } = readPayrollForm15WageAmounts(payrollRow);
  const basicNum = Number(basic);
  if (!Number.isFinite(basicNum) || basicNum <= 0) return '';
  const rate = (basicNum / 26 / 8) * 2;
  if (!Number.isFinite(rate) || rate <= 0) return '';
  return String(Math.round(rate * 100) / 100);
}

export function applyFormXXIIITamilNaduOtNilToRow(row, headers, helpers = {}) {
  const hdrs = Array.isArray(headers) ? headers : [];
  const out = row && typeof row === 'object' ? { ...row } : {};
  const nilText = helpers.nilText != null ? String(helpers.nilText) : FORM_XXIII_TN_OT_NIL;
  const { overwrite = true } = helpers;

  hdrs.forEach((header) => {
    if (!isFormXXIIITamilNaduOtNilHeader(header)) return;
    const existing = String(out[header] ?? '').trim();
    if (
      !overwrite &&
      existing &&
      !/^enter\b/i.test(existing) &&
      !/^nil+$/i.test(existing) &&
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

export function applyFormXXIIITamilNaduOtNilToMappedRows(
  mappedData,
  headers,
  nilText = FORM_XXIII_TN_OT_NIL,
  helpers = {}
) {
  if (!Array.isArray(mappedData)) return [];
  const { overwrite = true } = helpers;
  return mappedData.map((row) =>
    applyFormXXIIITamilNaduOtNilToRow(row, headers, { nilText, overwrite })
  );
}
