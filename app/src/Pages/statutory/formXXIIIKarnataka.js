/** Form XXIII Karnataka — Register of Overtime [See rule 78(1)(a)(iii)]. */

/** Form XXIII KA overtime columns that always default to NIL (no attendance/payroll OT). */
export const FORM_XXIII_KA_OT_NIL = 'NIL';

const normFormXXIIIKarnatakaHeader = (h) =>
  String(h || '')
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();

/** Total overtime worked or production in case of piece rates */
export function isFormXXIIIKarnatakaTotalOvertimeWorkedHeader(h) {
  const s = normFormXXIIIKarnatakaHeader(h);
  if (!s) return false;
  return (
    (s.includes('total') && s.includes('overtime') && (s.includes('worked') || s.includes('production'))) ||
    (s.includes('overtime') && s.includes('piece') && s.includes('rate'))
  );
}

/** Overtime rate of wages */
export function isFormXXIIIKarnatakaOvertimeRateHeader(h) {
  const s = normFormXXIIIKarnatakaHeader(h);
  if (!s || /normal/.test(s)) return false;
  return s.includes('overtime') && s.includes('rate') && (s.includes('wage') || s.includes('pay'));
}

/** Overtime earnings */
export function isFormXXIIIKarnatakaOvertimeEarningsHeader(h) {
  const s = normFormXXIIIKarnatakaHeader(h);
  return s.includes('overtime') && s.includes('earning');
}

/** Date on which overtime wages paid */
export function isFormXXIIIKarnatakaOtWagesPaidDateHeader(h) {
  const s = normFormXXIIIKarnatakaHeader(h);
  return (
    (s.includes('overtime') && s.includes('paid')) ||
    /dates?\s+on\s+which\s+overtime\s+wage/.test(s)
  );
}

export function isFormXXIIIKarnatakaOtNilHeader(header) {
  return (
    isFormXXIIIKarnatakaTotalOvertimeWorkedHeader(header) ||
    isFormXXIIIKarnatakaOvertimeRateHeader(header) ||
    isFormXXIIIKarnatakaOvertimeEarningsHeader(header) ||
    isFormXXIIIKarnatakaOtWagesPaidDateHeader(header)
  );
}

export function applyFormXXIIIKarnatakaOtNilToRow(row, headers, helpers = {}) {
  const hdrs = Array.isArray(headers) ? headers : [];
  const out = row && typeof row === 'object' ? { ...row } : {};
  const nilText = helpers.nilText != null ? String(helpers.nilText) : FORM_XXIII_KA_OT_NIL;
  const { overwrite = true } = helpers;

  hdrs.forEach((header) => {
    if (!isFormXXIIIKarnatakaOtNilHeader(header)) return;
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

export function applyFormXXIIIKarnatakaOtNilToMappedRows(
  mappedData,
  headers,
  nilText = FORM_XXIII_KA_OT_NIL,
  helpers = {}
) {
  if (!Array.isArray(mappedData)) return [];
  const { overwrite = true } = helpers;
  return mappedData.map((row) =>
    applyFormXXIIIKarnatakaOtNilToRow(row, headers, { nilText, overwrite })
  );
}
