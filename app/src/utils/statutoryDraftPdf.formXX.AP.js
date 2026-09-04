// statutoryDraftPdf.formXX.AP.js

import {
  collectUniqueApHeaderLines,
  pickExclusiveApHeaderTitles
} from './statutoryDraftPdf.apHeaderTitles';

const normalizeFormXXText = (text) =>
  String(text || '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();

export const isFormXXDateOfRecoveryGroupLabel = (text) => {
  return normalizeFormXXText(text) === 'date of recovery';
};

export const isFormXXAPTableHeaderRow = (row = []) => {
  const blob = (Array.isArray(row) ? row : [])
    .map((cell) => normalizeFormXXText(cell))
    .filter(Boolean)
    .join(' ');
  return (
    /(?:^|\s)s\.?\s*no\b/.test(blob) &&
    /name of workmen/.test(blob) &&
    /date of recovery/.test(blob) &&
    /remarks?/.test(blob)
  );
};

export const isFormXXAPAdministrativeRow = (row = []) => {
  const blob = (Array.isArray(row) ? row : [])
    .map((cell) => normalizeFormXXText(cell))
    .filter(Boolean)
    .join(' ');
  return (
    /contractor/.test(blob) ||
    /establishment\s+in\b/.test(blob) ||
    /nature\s+and\s+location\s+of\s+work/.test(blob) ||
    /principal\s+employer/.test(blob)
  );
};

export const getFormXXAPHeaderTitles = (metaLines = [], rows = []) => {
  return pickExclusiveApHeaderTitles(collectUniqueApHeaderLines(metaLines, rows), {
    form: /form\s*[-._ ]*xx\b/i,
    register: /register\s+of\s+deductions\s+for\s+damage\s+or\s+loss/i,
    rule: /vide\s+rule\s+78/i,
    formFallback: 'FORM - XX',
    registerFallback: 'REGISTER OF DEDUCTIONS FOR DAMAGE OR LOSS',
    ruleFallback:
      '[Vide Rule 78 (1) (a) (ii) of Contract Labour (Reg. & Abolition) Central & A.P.Rules]',
    ruleExtract: (text) =>
      String(text || '').match(
        /\[?Vide\s+Rule\s+78\s*\(\s*1\s*\)\s*\(\s*a\s*\)\s*\(\s*ii\s*\)[^\]]*\]?/i
      )?.[0] || ''
  });
};

export const looksLikeFormXXAPPdfContext = (
  metaLines = [],
  rows = [],
  sheetName = ''
) => {
  const blob = [
    ...(metaLines || []),
    ...(rows || []).slice(0, 20).flat(),
    sheetName || ''
  ]
    .join(' ')
    .replace(/\s+/g, ' ')
    .toLowerCase();

  return (
    /form\s*[-._ ]*xx\b/.test(blob) &&
    /register\s+of\s+deductions\s+for\s+damage\s+or\s+loss/.test(blob) &&
    /date\s+of\s+recovery/.test(blob)
  );
};