/**
 * Tamil Nadu Form XIX — Wage Slip PDF helpers.
 * Key/value A–D layout (Workman Code … Net Amount of Wages Paid).
 * Does not change Karnataka / Gujarat / MP / AP / Rajasthan Form XIX.
 */

const normalizeFormXIXTNText = (text) =>
  String(text || '')
    .replace(/\s+/g, ' ')
    .trim();

const lowerFormXIXTNText = (text) => normalizeFormXIXTNText(text).toLowerCase();

const formXIXTNContextBlob = (metaLines, rows, sheetName = '', fileName = '') =>
  [...(metaLines || []), ...(rows || []).slice(0, 22).flat(), sheetName || '', fileName || '']
    .join(' ')
    .toLowerCase();

/** True when context is clearly another state's Form XIX — never steal those layouts. */
const isOtherStateFormXIXPdfContext = (blob) => {
  if (/tamil[\s._-]*nadu|tamilnadu|form[\s._-]*xix[\s._-]*tamil/.test(blob)) return false;
  return (
    /gujarat|form[\s._-]*xix[\s._-]*gj|xix_gj/.test(blob) ||
    /karnataka|form[\s._-]*xix[\s._-]*ka\b|xix_ka/.test(blob) ||
    /madhya\s+pradesh|form[\s._-]*xix[\s._-]*mp|xix_mp/.test(blob) ||
    /andhra\s+pradesh|form[\s._-]*xix[\s._-]*ap\b|xix_ap/.test(blob) ||
    /rajasthan|form[\s._-]*xix[\s._-]*rj|xix_rj|register\s+of\s+overtime/.test(blob)
  );
};

/**
 * Tamil Nadu CLRA Form XIX wage slip (Workman Code key/value grid).
 */
export const looksLikeFormXIXTamilNaduPdfContext = (
  metaLines,
  rows,
  sheetName = '',
  fileName = ''
) => {
  const blob = formXIXTNContextBlob(metaLines, rows, sheetName, fileName);
  if (!blob) return false;
  if (isOtherStateFormXIXPdfContext(blob)) return false;

  const isXix =
    /form\s*xix\b/.test(blob) ||
    /form[\s._-]*xix[\s._-]/.test(blob) ||
    (/wage\s*s?\s*lip/.test(blob) && /rule\s*78/.test(blob));
  if (!isXix) return false;

  const isTnName = /tamil[\s._-]*nadu|tamilnadu|form[\s._-]*xix[\s._-]*tamil|xix[_\s-]*tn\b/.test(
    blob
  );
  const looksLikeTnLayout =
    /workman\s+code/.test(blob) &&
    /net\s+amount\s+of\s+wages\s+paid/.test(blob) &&
    (/wages?\s+computation/.test(blob) ||
      /no\.?\s*of\.?\s*working\s+days/.test(blob) ||
      /father'?s?\s+name/.test(blob));

  return isTnName || looksLikeTnLayout;
};

/**
 * After Excel merge collapse, Net Amount is often:
 * ['Net Amount of Wages Paid', '', '', '130787']
 * PDF must re-group label columns into one box (A–C) + amount (D).
 */
export const resolveFormXIXTamilNaduNetAmountPdfSpan = (row, colCount) => {
  const n = Math.max(0, Number(colCount) || 0);
  if (!Array.isArray(row) || n < 2) return null;

  let labelCol = -1;
  let amountCol = -1;
  let label = '';
  let amount = '';

  for (let c = 0; c < n; c += 1) {
    const t = normalizeFormXIXTNText(row[c]);
    if (!t) continue;
    if (/net\s+amount\s+of\s+wages\s+paid/i.test(t)) {
      labelCol = c;
      label = t;
      continue;
    }
    if (amountCol < 0) {
      amountCol = c;
      amount = t;
    }
  }

  if (labelCol < 0) return null;

  if (amountCol < 0) {
    return {
      labelStart: labelCol,
      labelEnd: n - 1,
      amountCol: -1,
      label,
      amount: '',
    };
  }

  // Prefer amount in the rightmost filled column when both sit on the row.
  if (amountCol <= labelCol) {
    amountCol = -1;
    amount = '';
    for (let c = n - 1; c > labelCol; c -= 1) {
      const t = normalizeFormXIXTNText(row[c]);
      if (!t) continue;
      amountCol = c;
      amount = t;
      break;
    }
  }

  if (amountCol < 0) {
    return {
      labelStart: labelCol,
      labelEnd: n - 1,
      amountCol: -1,
      label,
      amount: '',
    };
  }

  return {
    labelStart: labelCol,
    labelEnd: Math.max(labelCol, amountCol - 1),
    amountCol,
    label,
    amount,
  };
};

export const isFormXIXTamilNaduNetAmountLabelText = (text) =>
  /net\s+amount\s+of\s+wages\s+paid/i.test(lowerFormXIXTNText(text));
