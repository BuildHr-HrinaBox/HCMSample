/** Tamil Nadu CLRA Form XXII — Register of Advances. */

export const FORM_XXII_TN_NIL = 'NILL';

const normFormXXIITamilNaduHeader = (h) =>
  String(h || '')
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();

const bareFormXXIITamilNaduHeader = (h) =>
  normFormXXIITamilNaduHeader(h)
    .replace(/^\(?\d+\)?\s*/, '')
    .replace(/^\d+[\).:\-]\s*/, '')
    .replace(/\s*\(\d+\)\s*$/, '')
    .replace(/\.+$/, '')
    .trim();

export function matchesFormXXIITamilNaduFileHint(blob) {
  const parts = String(blob || '').toLowerCase();
  if (!parts) return false;
  if (/form[\s._-]*xxix(?![a-z])/i.test(parts)) return false;
  const hasXXII =
    /form[\s._-]*xxii(?![a-z])/i.test(parts) ||
    /(?:^|[\s._-])xxii(?=[\s._\W-]|$)/i.test(parts);
  const hasTN =
    /tamil[\s._-]*nadu|tamilnadu|\(tn\)|\btn\b/.test(parts) ||
    /\bform[\s._-]*xxii[\s._-]*tamil/.test(parts) ||
    /\bxxii[\s._-]*tamil/.test(parts);
  return hasXXII && hasTN;
}

export function isFormXXIITamilNaduContext(
  formHeader,
  rowItem,
  fileName,
  tableHeaders = [],
  sheetText = ''
) {
  const fileBlob = String(fileName || rowItem?.formFileName || rowItem?.FormFileName || '').toLowerCase();
  if (matchesFormXXIITamilNaduFileHint(fileBlob)) return true;

  const parts = [
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

  if (!parts) return false;
  if (/form[\s._-]*xxix(?![a-z])/i.test(parts)) return false;
  if (/register\s+of\s+employment|shops\s*(?:&|and)\s*establishment/i.test(parts)) {
    return false;
  }

  const hasTN =
    /tamil[\s._-]*nadu|tamilnadu|\(tn\)|\btn\b/.test(parts) ||
    /\bform[\s._-]*xxii[\s._-]*tamil/.test(parts) ||
    /\bxxii[\s._-]*tamil/.test(parts);
  const hasXXII =
    /form[\s._-]*xxii(?![a-z])/i.test(parts) ||
    /(?:^|[\s._-])xxii(?=[\s._\W-]|$)/i.test(parts);
  const hasAdvances =
    /register\s+of\s+advances/i.test(parts) ||
    /date\s+and\s+amount\s+of\s+advance/i.test(parts) ||
    /wage\s+period\s+and\s+wages\s+payable/i.test(parts);

  return hasTN && hasXXII && hasAdvances;
}

/**
 * Form XXII TN advance columns that default to NILL (manual / no advance entry).
 * Wage period and wages payable
 * Date and amount of advance given
 * Purpose (s) for which advance made
 * No. of installments by which advance to be repaid
 * Date and amount of each installment repaid
 * Date on which last installment was repaid
 */
export function isFormXXIITamilNaduNilDefaultHeader(header) {
  const bare = bareFormXXIITamilNaduHeader(header);
  if (!bare) return false;

  if (/wage\s+period/.test(bare) && /wages?\s+payable/.test(bare)) return true;
  if (/date\s+and\s+amount/.test(bare) && /advance\s+given/.test(bare)) return true;
  if (/purpose/.test(bare) && /advance/.test(bare)) return true;
  if (
    /(?:no\.?|number)\s+of\s+install?ments?/.test(bare) &&
    /(?:to\s+be\s+)?repaid|by\s+which\s+advance/.test(bare)
  ) {
    return true;
  }
  if (
    /date\s+and\s+amount/.test(bare) &&
    /(?:each\s+)?install?ment/.test(bare) &&
    /repaid/.test(bare)
  ) {
    return true;
  }
  if (/last\s+install?ment/.test(bare) && /repaid/.test(bare)) return true;

  return false;
}

export function applyFormXXIITamilNaduNilToRow(row, headers, helpers = {}) {
  const hdrs = Array.isArray(headers) ? headers : [];
  const out = row && typeof row === 'object' ? { ...row } : {};
  const nilText = helpers.nilText != null ? String(helpers.nilText) : FORM_XXII_TN_NIL;
  const { overwrite = true } = helpers;

  hdrs.forEach((header) => {
    if (!isFormXXIITamilNaduNilDefaultHeader(header)) return;
    const existing = String(out[header] ?? '').trim();
    if (
      !overwrite &&
      existing &&
      !/^enter\b/i.test(existing) &&
      !/^nil+l?$/i.test(existing) &&
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

export function applyFormXXIITamilNaduNilToMappedRows(
  mappedData,
  headers,
  nilText = FORM_XXII_TN_NIL,
  helpers = {}
) {
  if (!Array.isArray(mappedData)) return [];
  const { overwrite = true } = helpers;
  return mappedData.map((row) =>
    applyFormXXIITamilNaduNilToRow(row, headers, { nilText, overwrite })
  );
}
