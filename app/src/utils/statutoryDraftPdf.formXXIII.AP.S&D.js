import {
    collectUniqueApHeaderLines,
    pickExclusiveApHeaderTitles
  } from './statutoryDraftPdf.apHeaderTitles';
  
  const normalizeFormXXIIIText = (text) =>
    String(text || '')
      .replace(/\s+/g, ' ')
      .trim();
  
  const lowerFormXXIIIText = (text) => normalizeFormXXIIIText(text).toLowerCase();
  
  const collectFormXXIIIAdminLines = (metaLines = [], rows = [], tableStart = 0) => {
    const lines = [];
    const push = (raw) => {
      const text = normalizeFormXXIIIText(raw);
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
  
  const collectFormXXIIISourceLines = (metaLines = [], rows = [], tableStart = 0) => {
    const beforeTable = (rows || []).slice(0, Math.max(0, Number(tableStart) || 0));
    return collectUniqueApHeaderLines(metaLines, beforeTable.length ? beforeTable : (rows || []).slice(0, 12));
  };
  
  export const looksLikeFormXXIIIAPSEPdfContext = (metaLines = [], rows = [], sheetName = '') => {
    const blob = [...(metaLines || []), ...(rows || []).slice(0, 24).flat(), sheetName || '']
      .map(lowerFormXXIIIText)
      .join(' ');
    if (!blob) return false;
    if (/register\s+of\s+overtime/.test(blob)) return false;
    if (/form[\s._-]*xxii\b/.test(blob) && !/form[\s._-]*xxiii\b/.test(blob)) return false;
    return (
      /form[\s._-]*xxiii\b/.test(blob) &&
      /register\s+of\s+wages/.test(blob) &&
      (/vide\s+rule\s*29|shops\s*&\s*establishment|a\.?\s*p\.?\s+shops/.test(blob) ||
        /name\s+of\s+establishment\s*\/?\s*shop/.test(blob))
    );
  };
  
  export const isFormXXIIIAPSETitleRow = (row = []) => {
    const blob = (Array.isArray(row) ? row : []).map(normalizeFormXXIIIText).filter(Boolean).join(' ');
    return (
      /form[\s._-]*xxiii\b/i.test(blob) ||
      /^register\s+of\s+wages\b/i.test(blob) ||
      /vide\s+rule\s*29/i.test(blob)
    );
  };
  
  export const isFormXXIIIAPSEAdminRow = (row = []) => {
    const blob = lowerFormXXIIIText((Array.isArray(row) ? row : []).join(' '));
    return (
      /name\s+of\s+establishment\s*\/?\s*shop/.test(blob) ||
      /address\s+of\s+(?:the\s+)?establishment/.test(blob) ||
      /registration\s+no/.test(blob) ||
      (/wage\s+period/.test(blob) && /from|to/.test(blob))
    );
  };
  
  export const isFormXXIIIAPSETableHeaderRow = (row = []) => {
    const blob = lowerFormXXIIIText((Array.isArray(row) ? row : []).join(' '));
    return (
      /(?:^|\s)s\.?\s*no\b/.test(blob) &&
      /name\s+of\s+(?:the\s+)?employee/.test(blob) &&
      (/rate\s+of\s+wages|gross\s+wages|actual\s+wages/.test(blob) || /deductions?\s+if\s+any/.test(blob))
    );
  };
  
  export const isFormXXIIIAPSEDeductionsGroupLabel = (text) => {
    const n = lowerFormXXIIIText(text);
    return /deductions?\s+if\s+any/.test(n) && /reasons?\s+thereof/.test(n);
  };
  
  export const getFormXXIIIAPSEHeaderTitles = (metaLines = [], rows = [], tableStart = 0) => {
    const lines = collectFormXXIIISourceLines(metaLines, rows, tableStart);
    const combined = lines.find(
      (line) => /form[\s._-]*xxiii\b/i.test(line) && /register\s+of\s+wages/i.test(line)
    );
    const exclusive = pickExclusiveApHeaderTitles(lines, {
      form: /form[\s._-]*xxiii\b/i,
      register: /register\s+of\s+wages/i,
      rule: /vide\s+rule\s*29/i,
      formFallback: 'Form XXIII – Register of Wages',
      registerFallback: '',
      ruleFallback: '(Vide Rule 29(2) of A.P. Shops & Establishment Rules, 1990)',
      ruleExtract: (text) =>
        String(text || '').match(
          /\(?\s*Vide\s+Rule\s*29\s*\(\s*2\s*\)\s*of\s+A\.?\s*P\.?\s+Shops\s*&\s*Establishment\s+Rules,?\s*1990\s*\)?/i
        )?.[0] || ''
    });
    if (!combined) return exclusive;
    const rule = exclusive.find((line) => /vide\s+rule/i.test(line));
    return [combined, rule].filter(Boolean);
  };
  
  const splitInlineValue = (text, pattern) => {
    const raw = normalizeFormXXIIIText(text);
    const match = raw.match(pattern);
    if (!match) return '';
    return normalizeFormXXIIIText(raw.slice(match.index + match[0].length).replace(/^[:.\-\s]+/, ''));
  };
  
  const isAdminLabelLine = (line) =>
    /name\s+of\s+establishment|address\s+of\s+(?:the\s+)?establishment|registration\s+no|^wage\s+period\b|^from\s*:|^to\s*:/i.test(
      line
    );
  
  export const extractFormXXIIIAPSEAdminLayout = (metaLines = [], rows = [], tableStart = 0) => {
    const lines = collectFormXXIIIAdminLines(metaLines, rows, tableStart);
    const layout = {
      establishmentLabel: 'Name of Establishment / Shop:',
      establishmentValue: '',
      registrationLabel: 'Registration No.:',
      registrationValue: '',
      addressLabel: 'Address of the Establishment:',
      addressValue: '',
      wagePeriodLabel: 'Wage Period',
      fromLabel: 'From:',
      fromValue: '',
      toLabel: 'To:',
      toValue: ''
    };
  
    let pending = '';
    lines.forEach((line) => {
      if (/form[\s._-]*xxiii\b|register\s+of\s+wages|vide\s+rule/i.test(line)) return;
      if (/name\s+of\s+establishment\s*\/?\s*shop/i.test(line)) {
        pending = 'establishmentValue';
        const inline = splitInlineValue(line, /name\s+of\s+establishment\s*\/?\s*shop/i);
        if (inline) {
          layout.establishmentValue = inline;
          pending = '';
        }
        return;
      }
      if (/registration\s+no/i.test(line)) {
        pending = 'registrationValue';
        const inline = splitInlineValue(line, /registration\s+no\.?/i);
        if (inline) {
          layout.registrationValue = inline;
          pending = '';
        }
        return;
      }
      if (/address\s+of\s+(?:the\s+)?establishment/i.test(line)) {
        pending = 'addressValue';
        const inline = splitInlineValue(line, /address\s+of\s+(?:the\s+)?establishment/i);
        if (inline) {
          layout.addressValue = inline;
          pending = '';
        }
        return;
      }
      if (/^wage\s+period\b/i.test(line)) {
        const fromInline = line.match(/from\s*:?\s*([^,]+?)(?:\s+to\s*:|\s*$)/i);
        const toInline = line.match(/to\s*:?\s*(.+)$/i);
        if (fromInline?.[1] && !/^from$/i.test(fromInline[1].trim())) {
          layout.fromValue = normalizeFormXXIIIText(fromInline[1]);
        }
        if (toInline?.[1] && !/^to$/i.test(toInline[1].trim())) {
          layout.toValue = normalizeFormXXIIIText(toInline[1]);
        }
        pending = layout.fromValue ? (layout.toValue ? '' : 'toValue') : 'fromValue';
        return;
      }
      if (/^from\s*:?/i.test(line)) {
        pending = 'fromValue';
        const inline = splitInlineValue(line, /^from/i);
        if (inline) {
          layout.fromValue = inline;
          pending = 'toValue';
        }
        return;
      }
      if (/^to\s*:?/i.test(line)) {
        pending = 'toValue';
        const inline = splitInlineValue(line, /^to/i);
        if (inline) {
          layout.toValue = inline;
          pending = '';
        }
        return;
      }
      if (pending && !isAdminLabelLine(line)) {
        layout[pending] = layout[pending] ? `${layout[pending]} ${line}` : line;
        pending = pending === 'fromValue' ? 'toValue' : '';
      }
    });
  
    return layout;
  };
  
  export const findFormXXIIIAPSEDeductionsBand = (rows = [], tableStart = 0, colCount = 0) => {
    const startScan = Math.max(0, Number(tableStart) || 0);
    const endScan = Math.min(rows.length - 1, startScan + 6);
    for (let r = startScan; r <= endScan; r += 1) {
      const row = rows[r] || [];
      const start = row.findIndex((cell) => isFormXXIIIAPSEDeductionsGroupLabel(cell));
      if (start < 0) continue;
      const childRow = rows[r + 1] || [];
      const amountAt = childRow.findIndex((cell) => /^amount\s*$/i.test(normalizeFormXXIIIText(cell)));
      const reasonAt = childRow.findIndex((cell) => /^reasons?\s*$/i.test(normalizeFormXXIIIText(cell)));
      const childStart = amountAt >= 0 ? amountAt : start;
      const childEnd =
        reasonAt >= 0 ? reasonAt : Math.min((colCount || row.length) - 1, childStart + 1);
      return {
        parentRow: r,
        childRow: r + 1,
        start: Math.min(start, childStart),
        end: Math.max(start + 1, childEnd)
      };
    }
    return null;
  };
  
  export const formXXIIIAPSEColumnWeight = (headerText = '') => {
    const lower = lowerFormXXIIIText(headerText);
    if (/(?:^|\s)s\.?\s*no\b/.test(lower)) return 4.2;
    if (/name\s+of\s+(?:the\s+)?employee/.test(lower)) return 14;
    if (/signature|thumb\s+impression/.test(lower)) return 11;
    if (/date\s+of\s+appointment|date\s+of\s+payment/.test(lower)) return 7.5;
    if (/wages\s+earned\s+for\s+overtime|normal\s+wages\s+earned|gross\s+wages|actual\s+wages|rate\s+of\s+wages/.test(lower)) {
      return 7.2;
    }
    if (/^wages\s+earned$/.test(lower)) return 6.5;
    if (/^amount$/.test(lower)) return 6.2;
    if (/^reasons?$/.test(lower) || /deductions?\s+if\s+any/.test(lower)) return 7.4;
    if (/^remarks?$/.test(lower)) return 7;
    return 6.5;
  };
  
  export const FORM_XXIII_APSE_ADMIN_WIDTHS = {
    label: 0.18,
    value: 0.5,
    rightLabel: 0.14,
    rightValue: 0.18
  };
  
  export const isFormXXIIIAPSECenterValueHeader = (headerText = '') => {
    const lower = lowerFormXXIIIText(headerText);
    return (
      /(?:^|\s)s\.?\s*no\.?\s*$/.test(lower) ||
      /^s\.?\s*no\b/.test(lower) ||
      /date\s+of\s+appointment/.test(lower) ||
      /date\s+of\s+payment/.test(lower)
    );
  };
  
  export const isFormXXIIIAPSEMoneyHeader = (headerText = '') => {
    const lower = lowerFormXXIIIText(headerText);
    if (/^reasons?$/.test(lower)) return false;
    return (
      /rate\s+of\s+wages/.test(lower) ||
      /normal\s+wages\s+earned/.test(lower) ||
      /^wages\s+earned$/.test(lower) ||
      /wages\s+earned\s+for\s+overtime/.test(lower) ||
      /gross\s+wages\s+payable/.test(lower) ||
      /actual\s+wages\s+paid/.test(lower) ||
      /^amount$/.test(lower)
    );
  };
  
  export const formatFormXXIIIAPSEMoneyPdfValue = (raw) => {
    const text = String(raw ?? '').trim();
    if (!text || /^nill?$/i.test(text)) return text;
    const normalized = text.replace(/,/g, '');
    if (!/^-?\d+(\.\d+)?$/.test(normalized)) return text;
    const amount = Number(normalized);
    if (!Number.isFinite(amount)) return text;
    const decimals = normalized.includes('.') ? Math.min(normalized.split('.')[1].length, 2) : 0;
    return amount.toLocaleString('en-IN', {
      maximumFractionDigits: decimals,
      minimumFractionDigits: decimals
    });
  };
  
  export const trimFormXXIIIAPSEEmptyPdfColumns = (
    rows = [],
    tableStart = 0,
    headerBandEnd = 0,
    colCount = 0
  ) => {
    const count = Math.max(0, Number(colCount) || 0);
    if (!count) return { rows, colCount: 0 };
    const headerEnd = Math.max(Number(tableStart) || 0, Number(headerBandEnd) || 0);
    const headerOf = (column) => {
      const parts = [];
      for (let r = Math.max(0, Number(tableStart) || 0); r <= headerEnd && r < rows.length; r += 1) {
        const text = String(rows[r]?.[column] || '').trim();
        if (text) parts.push(text);
      }
      return parts.join(' ');
    };
    const hasBody = (column) => {
      for (let r = headerEnd + 1; r < rows.length; r += 1) {
        if (String(rows[r]?.[column] || '').trim()) return true;
      }
      return false;
    };
    let signatureAt = -1;
    let remarksAt = -1;
    for (let c = 0; c < count; c += 1) {
      const header = lowerFormXXIIIText(headerOf(c));
      if (signatureAt < 0 && /signature|thumb\s+impression/.test(header)) signatureAt = c;
      if (/^remarks?$/.test(header)) remarksAt = c;
    }
    const keep = [];
    for (let c = 0; c < count; c += 1) {
      const empty = !headerOf(c) && !hasBody(c);
      const betweenSignatureAndRemarks =
        signatureAt >= 0 && remarksAt > signatureAt && c > signatureAt && c < remarksAt;
      const trailingAfterRemarks = remarksAt >= 0 && c > remarksAt;
      if (empty && (betweenSignatureAndRemarks || trailingAfterRemarks)) continue;
      keep.push(c);
    }
    if (!keep.length || keep.length === count) return { rows, colCount: count };
    return {
      rows: rows.map((row) => keep.map((index) => (Array.isArray(row) ? row[index] : ''))),
      colCount: keep.length
    };
  };