/** Tamil Nadu CLRA Form XXI — Register of Fines [See rule 78 (1) a) (ii)]. */

export const FORM_XXI_TN_NIL = 'NIL';

const normFormXXITamilNaduHeader = (h) =>
  String(h || '')
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();

const bareFormXXITamilNaduHeader = (h) =>
  normFormXXITamilNaduHeader(h)
    .replace(/^\(?\d+\)?\s*/, '')
    .replace(/^\d+[\).:\-]\s*/, '')
    .replace(/\s*\(\d+\)\s*$/, '')
    .replace(/\.+$/, '')
    .trim();

/** True for Form_XXI_-_TamilNadu.xlsx style filenames (not XXII / XXIII / XXIX). */
export function matchesFormXXITamilNaduFileHint(blob) {
  const parts = String(blob || '').toLowerCase();
  if (!parts) return false;
  if (/form[\s._-]*xxi[iixv]/i.test(parts)) return false;
  const hasXXI =
    /form[\s._-]*xxi(?![a-z])/i.test(parts) ||
    /(?:^|[\s._-])xxi(?=[\s._\W-]|$)/i.test(parts);
  const hasTN =
    /tamil[\s._-]*nadu|tamilnadu|\(tn\)|\btn\b/.test(parts) ||
    /\bform[\s._-]*xxi[\s._-]*tamil/.test(parts) ||
    /\bxxi[\s._-]*tamil/.test(parts);
  return hasXXI && hasTN;
}

export function isFormXXITamilNaduContext(
  formHeader,
  rowItem,
  fileName,
  tableHeaders = [],
  sheetText = ''
) {
  const fileBlob = String(fileName || rowItem?.formFileName || rowItem?.FormFileName || '').toLowerCase();
  if (matchesFormXXITamilNaduFileHint(fileBlob)) return true;

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
  if (/form[\s._-]*xxi[iixv]/i.test(parts)) return false;

  const hasTN =
    /tamil[\s._-]*nadu|tamilnadu|\(tn\)|\btn\b/.test(parts) ||
    /\bform[\s._-]*xxi[\s._-]*tamil/.test(parts) ||
    /\bxxi[\s._-]*tamil/.test(parts);
  const hasXXI =
    /form[\s._-]*xxi(?![a-z])/i.test(parts) ||
    /(?:^|[\s._-])xxi(?=[\s._\W-]|$)/i.test(parts);
  const hasFines =
    /register\s+of\s+fines/i.test(parts) ||
    /act\s*\/?\s*omission/i.test(parts) ||
    /fine\s+imposed/i.test(parts);

  return hasTN && hasXXI && hasFines;
}

/**
 * Form XXI TN fine columns that default to NIL (manual / no fine imposed).
 * Act/Omission for which fine imposed
 * Date of offence
 * Whether workman showed cause against fine
 * Name of person in whose presence employee's explanation was heard
 * Wage periods and wages payable
 * Amount of fine Imposed
 * Date on which fine realized
 */
export function isFormXXITamilNaduNilDefaultHeader(header) {
  const bare = bareFormXXITamilNaduHeader(header);
  if (!bare) return false;

  if (/^(s\.?\s*no\.?|sl\.?\s*no\.?|serial(\s+number)?)$/.test(bare)) return false;
  if (/name\s+and\s+address\s+of\s+the\s+workman/.test(bare)) return false;
  if (/father|husband|spouse/.test(bare)) return false;
  if (/nature\s+of\s+employment/.test(bare)) return false;
  if (/^remarks?$/.test(bare)) return false;

  if (/act\s*\/?\s*omission/.test(bare) && /fine/.test(bare)) return true;
  if (/act\s*\/?\s*omission/.test(bare)) return true;
  if (/date\s+of\s+offence/.test(bare)) return true;
  if (/whether\s+work\s*man/.test(bare) && /cause|fine/.test(bare)) return true;
  if (/showed\s+cause/.test(bare) || /show\s+cause\s+against\s+fine/.test(bare)) return true;
  if (
    /name\s+of\s+person/.test(bare) &&
    (/whose|presence|explanation/.test(bare) || /in\s+whose/.test(bare))
  ) {
    return true;
  }
  if (/wage\s+period/.test(bare) && /wages?\s+payable/.test(bare)) return true;
  if (/amount\s+of\s+fine/.test(bare) || (/fine\s+imposed/.test(bare) && /amount/.test(bare))) {
    return true;
  }
  if (/fine\s+imposed/.test(bare) && !/act|omission|cause|show/.test(bare)) return true;
  if (
    /fine\s+reali[sz]/.test(bare) ||
    (/date\s+on\s+which/.test(bare) && /fine/.test(bare) && /reali[sz]/.test(bare))
  ) {
    return true;
  }

  return false;
}

export function applyFormXXITamilNaduNilToRow(row, headers, helpers = {}) {
  const hdrs = Array.isArray(headers) ? headers : [];
  const out = row && typeof row === 'object' ? { ...row } : {};
  const nilText = helpers.nilText != null ? String(helpers.nilText) : FORM_XXI_TN_NIL;
  const { overwrite = true } = helpers;

  hdrs.forEach((header) => {
    if (!isFormXXITamilNaduNilDefaultHeader(header)) return;
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

export function applyFormXXITamilNaduNilToMappedRows(
  mappedData,
  headers,
  nilText = FORM_XXI_TN_NIL,
  helpers = {}
) {
  if (!Array.isArray(mappedData)) return [];
  const { overwrite = true } = helpers;
  return mappedData.map((row) =>
    applyFormXXITamilNaduNilToRow(row, headers, { nilText, overwrite })
  );
}
