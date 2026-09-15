import {
  collectUniqueApHeaderLines,
  pickExclusiveApHeaderTitles
} from './statutoryDraftPdf.apHeaderTitles';

const normalizeFormXVIText = (text) =>
  String(text || '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();

const FORM_XVI_AP_LEFTOVER_HEADER_RE =
  /\b(?:Form\s*XVI\b|Muster\s+Roll|Address\s+of\s+the\s+Establishment|Name\s+and\s+Address\s+of\s+the\s+Establishment|Name\s+and\s+Address\s+of\s+Contractor|Name\s+and\s+address\s+of\s+Principal\s+Employer|\(?\s*Vide\s+rule)/i;

const FORM_XVI_AP_NATURE_CUT_RE =
  /\b(?:For\s+the\s+Month\s+of|Form\s*XVI\b|Address\s+of\s+the\s+Establishment|Name\s+and\s+Address\s+of\s+the\s+Establishment|Name\s+and\s+Address\s+of\s+Contractor|Name\s+and\s+address\s+of\s+Principal\s+Employer|\(?\s*Vide\s+rule)/i;

const FORM_XVI_AP_MONTH_YEAR_RE =
  /\b(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{4}\b/i;

export const sanitizeFormXVIAPMonthValue = (value) => {
  const text = String(value || '').replace(/\s+/g, ' ').trim();
  if (!text) return '';
  const cut = text.split(FORM_XVI_AP_LEFTOVER_HEADER_RE)[0].trim();
  const monthYear = cut.match(FORM_XVI_AP_MONTH_YEAR_RE) || text.match(FORM_XVI_AP_MONTH_YEAR_RE);
  if (monthYear) return monthYear[0];
  return cut.replace(/^[:\s]+|[:\s]+$/g, '').trim();
};

export const sanitizeFormXVIAPNatureValue = (value) => {
  const text = String(value || '').replace(/\s+/g, ' ').trim();
  if (!text) return '';
  return text
    .split(FORM_XVI_AP_NATURE_CUT_RE)[0]
    .replace(/^[:\s]+|[:\s]+$/g, '')
    .trim();
};

export const looksLikeFormXVIAPPdfContext = (metaLines = [], rows = [], sheetName = '') => {
  const blob = [...(metaLines || []), ...(rows || []).slice(0, 24).flat(), sheetName || '']
    .map(normalizeFormXVIText)
    .join(' ');
  return (
    /form[\s._-]*xvi\b/.test(blob) &&
    (/muster\s+roll|dates?\s+1\s+2\s+3|nature\s+and\s+location\s+of\s+work/.test(blob)) &&
    !/form[\s._-]*xvii\b|wages?\s+cum\s+muster/.test(blob)
  );
};

export const isFormXVIAPTableHeaderRow = (row = []) => {
  const blob = (Array.isArray(row) ? row : []).map(normalizeFormXVIText).join(' ');
  return /name\s+of\s+(?:the\s+)?workm[ae]n|serial\s+no|s\.?\s*no/.test(blob) &&
    (/dates?\b|attendance|muster\s+roll|workman/.test(blob));
};

export const isFormXVIIDateNumberRow = (row = []) => {
  const values = (Array.isArray(row) ? row : [])
    .map((cell) => String(cell || '').trim())
    .filter(Boolean);
  const numbers = values.filter((value) => /^\d{1,2}$/.test(value)).map(Number);
  return numbers.length >= 20 && numbers.includes(1) && numbers.includes(31);
};

export const getFormXVIAPHeaderTitles = (metaLines = [], rows = []) => {
  return pickExclusiveApHeaderTitles(collectUniqueApHeaderLines(metaLines, rows), {
    form: /form[\s._-]*xvi\b/i,
    register: /muster\s+roll/i,
    rule: /vide\s+rule\s+78/i,
    formFallback: 'FORM XVI',
    registerFallback: 'MUSTER ROLL',
    ruleFallback: '[Vide Rule 78 (1) (a) (i) of the Contract Labour (Regulation and Abolition) Central / A.P. Rules]',
    ruleExtract: (text) => {
      const ruleMatch = String(text || '').match(/Vide\s+Rule\s+78[^:]*?(?=\s+Name\s+and\s+Address|$)/i);
      return ruleMatch ? ruleMatch[0].replace(/\s+$/g, '') : '';
    }
  });
};