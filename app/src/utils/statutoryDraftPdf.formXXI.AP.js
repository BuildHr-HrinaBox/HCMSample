import {
	collectUniqueApHeaderLines,
	pickExclusiveApHeaderTitles
  } from './statutoryDraftPdf.apHeaderTitles';
  
  const normalizeFormXXIText = (text) =>
	  String(text || '')
		  .replace(/\s+/g, ' ')
		  .trim()
		  .toLowerCase();
  
  export const looksLikeFormXXIAPPdfContext = (
	  metaLines = [],
	  rows = [],
	  sheetName = ''
  ) => {
	  const blob = [
		  ...(metaLines || []),
		  ...(rows || []).slice(0, 20).flat(),
		  sheetName || ''
	  ]
		  .map(normalizeFormXXIText)
		  .join(' ');
	  return (
		  /form\s*[-._ ]*xxi\b/.test(blob) &&
		  /register\s+of\s+fines/.test(blob) &&
		  (/andhra\s+pradesh|central\s*&\s*a\.?p\.?|central\/a\.?p\.?|a\.?p\.?\s+rules/.test(blob) ||
			  /act\s*\/\s*omission\s+for\s+which\s+fine\s+imposed/.test(blob))
	  );
  };
  
  export const getFormXXIAPHeaderTitles = (metaLines = [], rows = []) => {
	  return pickExclusiveApHeaderTitles(collectUniqueApHeaderLines(metaLines, rows), {
		  form: /form\s*[-._ ]*xxi\b/i,
		  register: /register\s+of\s+fines/i,
		  rule: /vide\s+rule\s+78/i,
		  formFallback: 'FORM - XXI',
		  registerFallback: 'REGISTER OF FINES',
		  ruleFallback:
			  '[Vide Rule 78 (1) (a) (ii) of the Contract Labour (Regulation and Abolition) Central Rules, 1971 / A.P. Rules]',
		  ruleExtract: (text) => String(text || '').match(/\[?Vide\s+Rule\s+78[^\]]*\]?/i)?.[0] || ''
	  });
  };