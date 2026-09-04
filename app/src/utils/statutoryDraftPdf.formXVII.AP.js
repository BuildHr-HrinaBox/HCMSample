import {
  collectUniqueApHeaderLines,
  pickExclusiveApHeaderTitles
} from './statutoryDraftPdf.apHeaderTitles';

const normalizeFormXVIIText = (text) =>
  String(text || '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();

export const looksLikeFormXVIIAPPdfContext = (metaLines = [], rows = [], sheetName = '') => {
  const blob = [
    ...(metaLines || []),
    ...(rows || []).slice(0, 20).flat(),
    sheetName || ''
  ]
    .map(normalizeFormXVIIText)
    .join(' ');
  return (
    /form[\s._-]*xvii\b/.test(blob) &&
    /register\s+of\s+wages/.test(blob) &&
    !/form[\s._-]*xviii\b|wages?[- ]cum[- ]muster/.test(blob) &&
    (/andhra\s+pradesh|central\s*&\s*a\.?p\.?|central\/a\.?p\.?|a\.?p\.?\s+rules/.test(blob) ||
      /name\s+and\s+address\s+of\s+contractor/.test(blob))
  );
};

export const isFormXVIITableHeaderRow = (row = []) => {
  const blob = (Array.isArray(row) ? row : [])
    .map(normalizeFormXVIIText)
    .filter(Boolean)
    .join(' ');
  return (
    /(?:s\.?|sr\.?|sl\.?)\s*no/.test(blob) &&
    /register\s+of\s+workmen|name\s+of\s+workman|name\s+of\s+the\s+workman/.test(blob) &&
    /day(?:s)?\s+worked|wage|deduction/.test(blob)
  );
};

export const isFormXVIIAdministrativeRow = (row = []) => {
  const blob = (Array.isArray(row) ? row : [])
    .map(normalizeFormXVIIText)
    .filter(Boolean)
    .join(' ');
  return /contractor|nature\s+and\s+location\s+of\s+work|wage\s+period|establishment|principal\s+employer/.test(blob);
};

/**
 * Form XVII AP header labels wrap onto many lines. Keep the block in the
 * upper part of the cell so the last line never sits on the bottom border.
 * Other statutory forms must not use this metric.
 */
export const formXVIIAPHeaderTextMetrics = ({
  lineCount = 1,
  fontSize = 8,
  cellHeight = 82,
  lineHeightFactor = 1.12,
  padTop = 5,
  padBottom = 8,
  vAlign = 'top'
} = {}) => {
  const lines = Math.max(1, Number(lineCount) || 1);
  const size = Math.max(5, Number(fontSize) || 8);
  const height = Math.max(size + 4, Number(cellHeight) || 0);
  const lineH = size * lineHeightFactor;
  const textBlockH = size + Math.max(0, lines - 1) * lineH;
  let firstBaseline =
    vAlign === 'middle'
      ? height / 2 - textBlockH / 2 + size
      : padTop + size;
  const lastBaseline = firstBaseline + Math.max(0, lines - 1) * lineH;
  const maxLast = height - padBottom - size * 0.22;
  if (lastBaseline > maxLast) firstBaseline -= lastBaseline - maxLast;
  firstBaseline = Math.max(firstBaseline, size * 0.92);
  return { firstBaseline, textBlockH, lastBaseline: firstBaseline + Math.max(0, lines - 1) * lineH };
};

export const getFormXVIIAPHeaderTitles = (metaLines = [], rows = []) => {
  return pickExclusiveApHeaderTitles(collectUniqueApHeaderLines(metaLines, rows), {
    form: /form[\s._-]*xvii\b/i,
    register: /register\s+of\s+wages/i,
    rule: /vide\s+rule\s+78/i,
    formFallback: 'FORM XVII',
    registerFallback: 'REGISTER OF WAGES',
    ruleFallback:
      '[Vide Rule 78 (1) (a) (i) of the Contract Labour (Regulation and Abolition) Central Rules, 1971 / A.P. Rules]',
    ruleExtract: (text) => String(text || '').match(/\[?Vide\s+Rule\s+78[^\]]*\]?/i)?.[0] || ''
  });
};