/** Form XXIII GJ — Register of Overtime (Gujarat). */

export function isFormXXIIIGJContext(formHeader, rowItem, fileName, sheetText = '') {
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
    sheetText,
  ]
    .filter((x) => x != null && String(x).trim() !== '')
    .join(' ')
    .toLowerCase();

  const hasGujarat =
    /gujarat/.test(parts) ||
    /\bform[\s._-]*xxiii[\s._-]*gj\b/.test(parts) ||
    /\bxxiii[\s._-]*gj\b/.test(parts) ||
    /form_xxiii_gj/.test(parts);
  const hasXXIII =
    /\bform[\s._-]*xxiii/i.test(parts) || /register\s+of\s+overtime/.test(parts);
  return hasGujarat && hasXXIII;
}

/** Form XXIII GJ overtime columns that always default to NIL (no attendance/payroll OT). */
export const FORM_XXIII_GJ_OT_NIL = 'NIL';

const normFormXXIIIGJHeader = (h) =>
  String(h || '')
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();

/** Dates on which overtime worked */
export function isFormXXIIIGJOtWorkedDatesHeader(h) {
  const s = normFormXXIIIGJHeader(h);
  if (!s) return false;
  if (/payment|paid|rate|wage|earning/.test(s) && !/worked/.test(s)) return false;
  return (
    /dates?\s+on\s+which\s+over[\s-]*time\s+worked/.test(s) ||
    (s.includes('date') && s.includes('overtime') && s.includes('worked') && !/paid|payment/.test(s))
  );
}

/** Total overtime worked or production in case of piece rate */
export function isFormXXIIIGJTotalOvertimeWorkedHeader(h) {
  const s = normFormXXIIIGJHeader(h);
  if (!s) return false;
  return (
    (s.includes('total') && s.includes('overtime') && (s.includes('worked') || s.includes('production'))) ||
    (s.includes('overtime') && s.includes('piece') && s.includes('rate'))
  );
}

/** Overtime rate of wages */
export function isFormXXIIIGJOvertimeRateHeader(h) {
  const s = normFormXXIIIGJHeader(h);
  if (!s || /normal/.test(s)) return false;
  return s.includes('overtime') && s.includes('rate') && (s.includes('wage') || s.includes('pay'));
}

/** Overtime earnings */
export function isFormXXIIIGJOvertimeEarningsHeader(h) {
  const s = normFormXXIIIGJHeader(h);
  return s.includes('overtime') && s.includes('earning');
}

/** Date on which overtime wages paid */
export function isFormXXIIIGJOtWagesPaidDateHeader(h) {
  const s = normFormXXIIIGJHeader(h);
  return (
    (s.includes('overtime') && s.includes('paid')) ||
    /dates?\s+on\s+which\s+overtime\s+wage/.test(s)
  );
}

export function isFormXXIIIGJOtNilHeader(header) {
  return (
    isFormXXIIIGJOtWorkedDatesHeader(header) ||
    isFormXXIIIGJTotalOvertimeWorkedHeader(header) ||
    isFormXXIIIGJOvertimeRateHeader(header) ||
    isFormXXIIIGJOvertimeEarningsHeader(header) ||
    isFormXXIIIGJOtWagesPaidDateHeader(header)
  );
}

export function applyFormXXIIIGJOtNilToRow(row, headers, helpers = {}) {
  const hdrs = Array.isArray(headers) ? headers : [];
  const out = row && typeof row === 'object' ? { ...row } : {};
  const nilText = helpers.nilText != null ? String(helpers.nilText) : FORM_XXIII_GJ_OT_NIL;
  const { overwrite = true } = helpers;

  hdrs.forEach((header) => {
    if (!isFormXXIIIGJOtNilHeader(header)) return;
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

export function applyFormXXIIIGJOtNilToMappedRows(
  mappedData,
  headers,
  nilText = FORM_XXIII_GJ_OT_NIL,
  helpers = {}
) {
  if (!Array.isArray(mappedData)) return [];
  const { overwrite = true } = helpers;
  return mappedData.map((row) =>
    applyFormXXIIIGJOtNilToRow(row, headers, { nilText, overwrite })
  );
}
