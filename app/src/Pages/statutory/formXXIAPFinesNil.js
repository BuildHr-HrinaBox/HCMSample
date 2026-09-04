/** AP CLRA Form XXI — Register of Fines NIL helpers (standalone for CRA production named exports). */

export const FORM_XXI_AP_FINE_COLUMN_NIL_TEXT = 'NIL';

function formXXIAPHeaderBare(header) {
  return String(header || '')
    .toLowerCase()
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .replace(/^\(?\d+\)?\s*/, '')
    .replace(/\s*\(\d+\)\s*$/, '')
    .trim();
}

export function isFormXXIAPFineNilHeader(header) {
  const bare = formXXIAPHeaderBare(header);
  if (!bare) return false;
  if (/^(s\.?\s*no\.?|sl\.?\s*no\.?|serial(\s+number)?)$/.test(bare)) return false;
  if (
    /\bname\s+of\s+(?:the\s+)?workmen\b|\bname\s+of\s+workman\b/.test(bare) ||
    (/^name\s+of\b/.test(bare) &&
      /workmen|workman|worker|employee/.test(bare) &&
      !/presence|explanation/.test(bare))
  ) {
    return false;
  }
  if (/father|husband/.test(bare) && !/presence|explanation|showed\s+cause/.test(bare)) {
    return false;
  }
  if (/nature\s+of\s+employ/.test(bare)) return false;
  if (/^remarks?$/.test(bare)) return false;

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

export function isFormXXIAPSkipAutofillHeader(header) {
  return isFormXXIAPFineNilHeader(header);
}

export function applyFormXXIAPFinesNilToRow(row, headers, helpers = {}) {
  const hdrs = Array.isArray(headers) ? headers : [];
  const out = row && typeof row === 'object' ? { ...row } : {};
  const nilText =
    helpers.nilText != null ? String(helpers.nilText) : FORM_XXI_AP_FINE_COLUMN_NIL_TEXT;
  const { overwrite = false } = helpers;

  hdrs.forEach((header) => {
    if (!isFormXXIAPFineNilHeader(header)) return;
    const existing = String(out[header] ?? '').trim();
    if (!overwrite && existing && !/^enter\b/i.test(existing) && !/^nil+$/i.test(existing)) return;
    out[header] = nilText;
  });
  return out;
}

export function applyFormXXIAPFinesNilToMappedRows(
  mappedData,
  headers,
  nilText = FORM_XXI_AP_FINE_COLUMN_NIL_TEXT,
  helpers = {}
) {
  if (!Array.isArray(mappedData)) return [];
  const { overwrite = false } = helpers;
  return mappedData.map((row) =>
    applyFormXXIAPFinesNilToRow(row, headers, { nilText, overwrite })
  );
}
