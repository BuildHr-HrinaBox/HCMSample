import {
    collectUniqueApHeaderLines,
    pickExclusiveApHeaderTitles
  } from './statutoryDraftPdf.apHeaderTitles';
  
  const normalizeFormXIIText = (text) =>
    String(text || '')
      .replace(/\s+/g, ' ')
      .trim();
  
  const lowerFormXIIText = (text) => normalizeFormXIIText(text).toLowerCase();
  
  const collectFormXIIAdminLines = (metaLines = [], rows = [], tableStart = 0) => {
    const lines = [];
    const push = (raw) => {
      const text = normalizeFormXIIText(raw);
      if (text) lines.push(text);
    };
    (metaLines || []).forEach(push);
    const limited = (rows || []).slice(0, Math.max(0, Number(tableStart) || 0));
    (limited.length ? limited : (rows || []).slice(0, 12)).forEach((row) => {
      if (Array.isArray(row)) row.forEach(push);
      else push(row);
    });
    return lines;
  };
  
  const collectFormXIISourceLines = (metaLines = [], rows = [], tableStart = 0) => {
    const beforeTable = (rows || []).slice(0, Math.max(0, Number(tableStart) || 0));
    return collectUniqueApHeaderLines(
      metaLines,
      beforeTable.length ? beforeTable : (rows || []).slice(0, 12)
    );
  };
  
  export const looksLikeFormXIIAPSEPdfContext = (metaLines = [], rows = [], sheetName = '') => {
    const blob = [...(metaLines || []), ...(rows || []).slice(0, 24).flat(), sheetName || '']
      .map(lowerFormXIIText)
      .join(' ');
    if (!blob) return false;
    if (/form[\s._-]*xxi{2,3}\b/.test(blob)) return false;
    if (/form[\s._-]*xiii\b/.test(blob)) return false;
    if (/form[\s._-]*xxix\b/.test(blob)) return false;
    if (/form[\s._-]*xiv\b/.test(blob) && !/form[\s._-]*xii(?![ivx])/.test(blob)) return false;
    if (/register\s+of\s+overtime|register\s+of\s+wages(?!\s+of)|register\s+of\s+workmen/.test(blob)) {
      if (!/register\s+of\s+advances/.test(blob)) return false;
    }
    const hasXII =
      /form[\s._-]*xii(?![ivx])/.test(blob) ||
      /form[\s._-]*12(?!\d)/.test(blob);
    const hasAdvances =
      /register\s+of\s+advances/.test(blob) ||
      /amount\s+of\s+advance\s+given/.test(blob) ||
      /postponement\s+granted/.test(blob);
    const hasAP =
      /vide\s+rule\s*18|shops\s*&\s*establishment|a\.?\s*p\.?\s+shops/.test(blob) ||
      /name\s+of\s+establishment/.test(blob);
    return hasXII && hasAdvances && hasAP;
  };
  
  export const isFormXIIAPSETitleRow = (row = []) => {
    const blob = (Array.isArray(row) ? row : []).map(normalizeFormXIIText).filter(Boolean).join(' ');
    return (
      /form[\s._-]*xii(?![ivx])/i.test(blob) ||
      /^register\s+of\s+advances\b/i.test(blob) ||
      /vide\s+rule\s*18/i.test(blob)
    );
  };
  
  export const isFormXIIAPSEAdminRow = (row = []) => {
    const blob = lowerFormXIIText((Array.isArray(row) ? row : []).join(' '));
    return /name\s+of\s+establishment/.test(blob);
  };
  
  export const isFormXIIAPSETableHeaderRow = (row = []) => {
    const blob = lowerFormXIIText((Array.isArray(row) ? row : []).join(' '));
    return (
      /(?:^|\s)s\.?\s*no\b/.test(blob) &&
      /name\s+of\s+(?:the\s+)?employee/.test(blob) &&
      /amount\s+of\s+advance|father'?s?\s*\/?\s*husband|postponement\s+granted/.test(blob)
    );
  };
  
  export const formatFormXIIAPSEFormTitle = (text = '') => {
    const line = normalizeFormXIIText(text);
    if (!/form[\s._-]*xii(?![ivx])/i.test(line)) return line;
    if (/register\s+of\s+advances|vide\s+rule/i.test(line)) return line;
    return 'Form XII';
  };
  
  export const getFormXIIAPSEHeaderTitles = (metaLines = [], rows = [], tableStart = 0) =>
    pickExclusiveApHeaderTitles(collectFormXIISourceLines(metaLines, rows, tableStart), {
      form: /form[\s._-]*xii(?![ivx])/i,
      register: /register\s+of\s+advances/i,
      rule: /vide\s+rule\s*18/i,
      formFallback: 'Form XII',
      registerFallback: 'Register of Advances of Wages',
      ruleFallback: '[Vide Rule 18(4) of A.P. Shops & Establishments Rules, 1990]',
      ruleExtract: (text) =>
        String(text || '').match(
          /\[?\s*Vide\s+Rule\s*18\s*\(\s*4\s*\)\s*of\s+A\.?\s*P\.?\s+Shops\s*&\s*Establishments?\s+Rules,?\s*1990\s*\]?/i
        )?.[0] || ''
    }).map((line) => formatFormXIIAPSEFormTitle(line));
  
  const splitInlineValue = (text, pattern) => {
    const raw = normalizeFormXIIText(text);
    const match = raw.match(pattern);
    if (!match) return '';
    return normalizeFormXIIText(raw.slice(match.index + match[0].length).replace(/^[:.\-\s]+/, ''));
  };
  
  export const extractFormXIIAPSEAdminLayout = (metaLines = [], rows = [], tableStart = 0) => {
    const lines = collectFormXIIAdminLines(metaLines, rows, tableStart);
    const layout = {
      establishmentLabel: 'Name of Establishment :',
      establishmentValue: '',
      establishmentLine: ''
    };
    lines.forEach((line, index) => {
      if (/form[\s._-]*xii(?![ivx])|register\s+of\s+advances|vide\s+rule/i.test(line)) return;
      if (/name\s+of\s+establishment/i.test(line)) {
        const inline = splitInlineValue(line, /name\s+of\s+establishment/i);
        if (inline) {
          layout.establishmentValue = inline;
          layout.establishmentLine = line;
        } else {
          layout.establishmentValue = normalizeFormXIIText(lines[index + 1] || '');
          layout.establishmentLine = layout.establishmentValue
            ? `${line} ${layout.establishmentValue}`
            : line;
        }
      }
    });
    return layout;
  };
  
  export const isFormXIIAPSENameHeader = (headerText = '') => {
    const lower = lowerFormXIIText(headerText);
    return /name\s+of\s+(?:the\s+)?employee/.test(lower) || /father'?s?\s*\/?\s*husband/.test(lower);
  };
  
  export const isFormXIIAPSERightValueHeader = (headerText = '') => {
    const lower = lowerFormXIIText(headerText);
    return /(?:^|\s)s\.?\s*no\.?\s*$/.test(lower) || /^s\.?\s*no\b/.test(lower);
  };
  
  export const isFormXIIAPSECenterValueHeader = (headerText = '') => {
    const lower = lowerFormXIIText(headerText);
    if (isFormXIIAPSENameHeader(headerText)) return false;
    if (/^remarks?$/.test(lower)) return false;
    return (
      /(?:^|\s)s\.?\s*no\.?\s*$/.test(lower) ||
      /^s\.?\s*no\b/.test(lower) ||
      /amount\s+of\s+advance/.test(lower) ||
      /date\s+on\s+which\s+advance/.test(lower) ||
      /purpose/.test(lower) ||
      /install?ments?/.test(lower) ||
      /postponement\s+granted/.test(lower) ||
      /date\s+on\s+which\s+total\s+amount/.test(lower)
    );
  };
  
  export const collapseFormXIIAPSERepeatedDisplayText = (raw) => {
    const text = normalizeFormXIIText(raw);
    if (!text) return text;
    const words = text.split(' ');
    if (words.length >= 2 && words.length % 2 === 0) {
      const half = words.length / 2;
      const left = words.slice(0, half).join(' ');
      const right = words.slice(half).join(' ');
      if (left && left === right) return left;
    }
    return text;
  };
  
  export const formXIIAPSEDisplayCellText = (primary, duplicate) => {
    const first = normalizeFormXIIText(primary);
    const second = normalizeFormXIIText(duplicate);
    if (!first) return collapseFormXIIAPSERepeatedDisplayText(second);
    if (!second || second === first) return collapseFormXIIAPSERepeatedDisplayText(first);
    if (first.startsWith(second) || second.startsWith(first)) {
      return collapseFormXIIAPSERepeatedDisplayText(first.length >= second.length ? first : second);
    }
    return collapseFormXIIAPSERepeatedDisplayText(first);
  };
  
  const headerTextAt = (rows, tableStart, headerBandEnd, column) => {
    const parts = [];
    for (let r = Math.max(0, Number(tableStart) || 0); r <= headerBandEnd && r < rows.length; r += 1) {
      const text = normalizeFormXIIText(rows[r]?.[column]);
      if (text) parts.push(text);
    }
    return parts.join(' ');
  };
  
  export const findFormXIIAPSEDuplicateColumnSpans = (
    rows = [],
    tableStart = 0,
    headerBandEnd = 0,
    colCount = 0
  ) => {
    const count = Math.max(0, Number(colCount) || 0);
    const headerEnd = Math.max(Number(tableStart) || 0, Number(headerBandEnd) || 0);
    const spans = [];
    for (let c = 0; c < count - 1; c += 1) {
      const header = headerTextAt(rows, tableStart, headerEnd, c);
      if (!isFormXIIAPSENameHeader(header)) continue;
      const nextHeader = headerTextAt(rows, tableStart, headerEnd, c + 1);
      if (nextHeader) continue;
      let keepSpan = true;
      for (let r = headerEnd + 1; r < rows.length; r += 1) {
        const left = normalizeFormXIIText(rows[r]?.[c]);
        const right = normalizeFormXIIText(rows[r]?.[c + 1]);
        if (!right) continue;
        if (right === left || left.includes(right) || right.includes(left)) continue;
        keepSpan = false;
        break;
      }
      if (!keepSpan) continue;
      spans.push({ start: c, end: c + 1, header });
      c += 1;
    }
    return spans;
  };
  
  export const formXIIAPSEHeaderWrapLines = (headerText = '') => {
    const text = normalizeFormXIIText(headerText);
    if (/name\s+of\s+(?:the\s+)?employee/i.test(text)) return ['Name of', 'the Employee'];
    if (/father'?s?\s*\/?\s*husband/i.test(text)) return ["Father's /", "Husband's Name"];
    return text ? [text] : [];
  };
  
  export const formXIIAPSEPairedCellValue = (row = [], span = null) => {
    if (!span) return '';
    return formXIIAPSEDisplayCellText(row[span.start], row[span.end]);
  };
  
  export const formXIIAPSESpanAt = (spans = [], column) =>
    (spans || []).find((span) => column >= span.start && column <= span.end) || null;
  
  export const ensureFormXIIAPSEPairedNameColumns = (
    rows = [],
    tableStart = 0,
    headerBandEnd = 0,
    colCount = 0
  ) => {
    const count = Math.max(0, Number(colCount) || 0);
    if (!count) return { rows, colCount: 0 };
    const existing = findFormXIIAPSEDuplicateColumnSpans(rows, tableStart, headerBandEnd, count);
    if (existing.length) return { rows, colCount: count };
    const headerEnd = Math.max(Number(tableStart) || 0, Number(headerBandEnd) || 0);
    const insertAfter = [];
    for (let c = 0; c < count; c += 1) {
      if (isFormXIIAPSENameHeader(headerTextAt(rows, tableStart, headerEnd, c))) insertAfter.push(c);
    }
    if (!insertAfter.length) return { rows, colCount: count };
    let nextRows = (rows || []).map((row) => (Array.isArray(row) ? [...row] : []));
    let nextCount = count;
    insertAfter
      .slice()
      .reverse()
      .forEach((index) => {
        nextRows = nextRows.map((row, r) => {
          const copy = Array.from({ length: Math.max(row.length, index + 1) }, (_, i) => row[i] ?? '');
          const isHeaderRow = r >= (Number(tableStart) || 0) && r <= headerEnd;
          copy.splice(index + 1, 0, isHeaderRow ? '' : copy[index]);
          return copy;
        });
        nextCount += 1;
      });
    return { rows: nextRows, colCount: nextCount };
  };
  
  export const formXIIAPSEColumnWeight = (headerText = '') => {
    const lower = lowerFormXIIText(headerText);
    if (/(?:^|\s)s\.?\s*no\b/.test(lower)) return 4;
    if (/name\s+of\s+(?:the\s+)?employee/.test(lower)) return 16;
    if (/father'?s?\s*\/?\s*husband/.test(lower)) return 16;
    if (/amount\s+of\s+advance/.test(lower)) return 8;
    if (/date\s+on\s+which\s+advance/.test(lower)) return 8.5;
    if (/purpose/.test(lower)) return 9;
    if (/install?ments?/.test(lower)) return 10;
    if (/postponement\s+granted/.test(lower)) return 8;
    if (/date\s+on\s+which\s+total\s+amount/.test(lower)) return 9.5;
    if (/^remarks?$/.test(lower)) return 7;
    return 6.5;
  };