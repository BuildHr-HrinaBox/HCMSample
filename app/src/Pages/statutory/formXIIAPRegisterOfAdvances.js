/** Andhra Pradesh Form XII — Register of Advances of Wages (Rule 18(4) A.P. Shops & Establishments). */

export const FORM_XII_AP_NIL = 'NIL';

const normFormXIIAPHeader = (h) =>
  String(h || '')
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();

const bareFormXIIAPHeader = (h) =>
  normFormXIIAPHeader(h)
    .replace(/^\(?\d+\)?\s*/, '')
    .replace(/^\d+[\).:\-]\s*/, '')
    .replace(/\s*\(\d+\)\s*$/, '')
    .replace(/\.+$/, '')
    .trim();

export function matchesFormXIIAPFileHint(blob) {
  const parts = String(blob || '').toLowerCase();
  if (!parts) return false;
  // Do not treat Form XXII / XIII / XIV / XXIX as Form XII (xxii contains "xii").
  if (/form[\s._-]*xxi{2,3}(?![a-z])/i.test(parts)) return false;
  if (/form[\s._-]*xiii(?![a-z])/i.test(parts)) return false;
  if (/form[\s._-]*xxix(?![a-z])/i.test(parts)) return false;
  const hasXII =
    /form[\s._-]*(?<![x])xii(?![a-zivx])/i.test(parts) ||
    /form[\s._-]*12(?!\d)/i.test(parts) ||
    /(?:^|[\s._-])(?<![x])xii(?=[\s._\W-]|$)/i.test(parts);
  const hasAP =
    /andhra[\s._-]*pradesh|a\.?\s*p\.?\s*shops|\(ap\)|\bap\b/.test(parts) ||
    /\bform[\s._-]*xii[\s._-]*andhra/.test(parts) ||
    /\bxii[\s._-]*andhra/.test(parts);
  return hasXII && hasAP;
}

export function isFormXIIAPRegisterOfAdvancesContext(
  formHeader,
  rowItem,
  fileName,
  tableHeaders = [],
  sheetText = ''
) {
  const fileBlob = String(fileName || rowItem?.formFileName || rowItem?.FormFileName || '').toLowerCase();
  if (matchesFormXIIAPFileHint(fileBlob)) return true;

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
  if (/form[\s._-]*xxi{2,3}(?![a-z])/i.test(parts)) return false;
  if (/form[\s._-]*xiii(?![a-z])/i.test(parts)) return false;
  if (/form[\s._-]*xxix(?![a-z])/i.test(parts)) return false;
  if (/register\s+of\s+fines|register\s+of\s+employment|register\s+of\s+overtime/i.test(parts)) {
    return false;
  }

  const hasAP =
    /andhra[\s._-]*pradesh|a\.?\s*p\.?\s*shops|\(ap\)|\bap\b/.test(parts) ||
    /\bform[\s._-]*xii[\s._-]*andhra/.test(parts);
  const hasXII =
    /form[\s._-]*(?<![x])xii(?![a-zivx])/i.test(parts) ||
    /form[\s._-]*12(?!\d)/i.test(parts) ||
    /(?:^|[\s._-])(?<![x])xii(?=[\s._\W-]|$)/i.test(parts);
  const hasAdvances =
    /register\s+of\s+advances(?:\s+of\s+wages)?/i.test(parts) ||
    /amount\s+of\s+advance\s+given/i.test(parts) ||
    /postponement\s+granted/i.test(parts);

  return hasAP && hasXII && hasAdvances;
}

/**
 * Form XII AP advance columns that default to NIL when no advance was given.
 * Amount of Advance Given
 * Date on which Advance was Given
 * Purpose(s) for which Advance was Given
 * No. of Instalments by which Advance to be Recovered
 * Postponement Granted
 * Date on which Total Amount is Recovered
 */
export function isFormXIIAPNilDefaultHeader(header) {
  const bare = bareFormXIIAPHeader(header);
  if (!bare) return false;

  if (/amount\s+of\s+advance/.test(bare) && /given/.test(bare)) return true;
  if (/date\s+on\s+which\s+advance/.test(bare) && /given/.test(bare)) return true;
  if (/purpose/.test(bare) && /advance/.test(bare) && /given|made/.test(bare)) return true;
  if (
    /(?:no\.?|number)\s+of\s+install?ments?/.test(bare) &&
    /(?:to\s+be\s+)?recover|by\s+which\s+advance/.test(bare)
  ) {
    return true;
  }
  if (/postponement\s+granted/.test(bare)) return true;
  if (/date\s+on\s+which\s+total\s+amount/.test(bare) && /recover/.test(bare)) return true;

  return false;
}

export function applyFormXIIAPNilToRow(row, headers, helpers = {}) {
  const hdrs = Array.isArray(headers) ? headers : [];
  const out = row && typeof row === 'object' ? { ...row } : {};
  const nilText = helpers.nilText != null ? String(helpers.nilText) : FORM_XII_AP_NIL;
  const { overwrite = true } = helpers;

  hdrs.forEach((header) => {
    if (!isFormXIIAPNilDefaultHeader(header)) return;
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

export function applyFormXIIAPNilToMappedRows(
  mappedData,
  headers,
  nilText = FORM_XII_AP_NIL,
  helpers = {}
) {
  if (!Array.isArray(mappedData)) return [];
  const { overwrite = true } = helpers;
  return mappedData.map((row) =>
    applyFormXIIAPNilToRow(row, headers, { nilText, overwrite })
  );
}
