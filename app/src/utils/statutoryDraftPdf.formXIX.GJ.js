/**
 * Gujarat Form XIX — Wage Slip PDF layout.
 * Excel rows 8–12 (header fields) stay above the wage box;
 * numbered wage particulars 1–7 stay inside the box;
 * "Initial of the contract…" stays below the outer box.
 * Does not change Madhya Pradesh / AP / other Form XIX layouts.
 */

const normalizeFormXIXGJText = (text) =>
  String(text || '')
    .replace(/\s+/g, ' ')
    .trim();

const lowerFormXIXGJText = (text) => normalizeFormXIXGJText(text).toLowerCase();

const FORM_XIX_GJ_HEADER_LABELS = [
  'Name and address of contractor',
  "Name and father's/husband's name of the workman",
  'Nature and location of work',
  'For the week/fortnight/month ending'
];

const FORM_XIX_GJ_WAGE_LABELS = [
  'Number of days worked',
  'Number of units worked in case of piece rate workers',
  'Rate of daily wages/piece rate',
  'Amount of overtime wages',
  'Gross wages payable',
  'Deductions, if any',
  'Net amount of wages paid'
];

const FORM_XIX_GJ_TITLES = ['FORM XIX', '[See rule 78 (2)(b)]', 'Wage Slip'];
const FORM_XIX_GJ_FOOTER = 'Initial of the Contractor or his representative';

/** Excel templates sometimes say "address if contractor" (typo for "of"). */
const FORM_XIX_GJ_HEADER_PATTERNS = [
  /name\s+and\s+address\s+(?:of|if)\s+(?:the\s+)?contractor/i,
  /name\s+and\s+(?:father.*husband|husband.*father).+workman|name\s+of\s+(?:the\s+)?workman/i,
  /nature\s+and\s+location\s+of\s+work/i,
  /for\s+the\s+(?:week|fortnight|fought|month)/i
];

const FORM_XIX_GJ_WAGE_PATTERNS = [
  /(?:^|\b)(?:1[.)]?\s*)?(?:number|no\.?)\s+of\s+days\s+worked/i,
  /(?:^|\b)(?:2[.)]?\s*)?(?:number|no\.?)\s+of\s+units\s+worked(?:\s+in\s+case\s+of\s+piece\s+rate\s+workers?)?/i,
  /(?:^|\b)(?:3[.)]?\s*)?rate\s+of\s+daily\s+wages?\s*\/?\s*piece\s*-?\s*rate/i,
  /(?:^|\b)(?:4[.)]?\s*)?amount\s+of\s+overtime\s+wages/i,
  /(?:^|\b)(?:5[.)]?\s*)?gross\s+wages\s+payable/i,
  /(?:^|\b)(?:6[.)]?\s*)?deductions?\s*,?\s*if\s+any/i,
  /(?:^|\b)(?:7[.)]?\s*)?net\s+amount\s+of\s+wages\s+paid/i
];

const compactFormXIXGJCandidates = (row) =>
  (Array.isArray(row) ? row : [])
    .map((cell) => normalizeFormXIXGJText(cell))
    .filter(Boolean);

const looksLikeFormXIXGJTitle = (text) => /^form\s*xix\b/i.test(normalizeFormXIXGJText(text));
const looksLikeFormXIXGJRule78 = (text) => /see\s+rule\s*78/i.test(lowerFormXIXGJText(text));
const looksLikeFormXIXGJWageSlip = (text) => /^wage\s*s?lip$/i.test(normalizeFormXIXGJText(text));
const looksLikeFormXIXGJOrdinalOnly = (text) => /^[1-7][.)]?$/.test(normalizeFormXIXGJText(text));

const stripFormXIXGJLabelPrefix = (raw) => {
  const text = normalizeFormXIXGJText(raw);
  if (!text) return '';
  const isDotsOnly = (s) => !s || /^\.+$/.test(s);

  // "Label............. value" or "Label: value"
  const dotted = text.match(/^(.*?)(?:\.{2,}|\s*:\s*|\s+-\s+)\s*(.*)$/);
  if (dotted) {
    const rest = normalizeFormXIXGJText(dotted[2]);
    if (rest && !isDotsOnly(rest)) return rest;
  }
  return '';
};

const isFormXIXGJBlockedValue = (text, blockedPatterns = []) => {
  const raw = normalizeFormXIXGJText(text);
  if (!raw) return true;
  if (looksLikeFormXIXGJTitle(raw) || looksLikeFormXIXGJRule78(raw) || looksLikeFormXIXGJWageSlip(raw)) {
    return true;
  }
  if (looksLikeFormXIXGJOrdinalOnly(raw)) return true;
  if (/initial\s+of\s+the\s+contract/i.test(raw)) return true;
  return blockedPatterns.some((re) => re.test(raw));
};

const findFormXIXGJPatternValue = (rows = [], patterns = [], blockedPatterns = patterns) => {
  for (let r = 0; r < rows.length; r += 1) {
    const row = Array.isArray(rows[r]) ? rows[r] : [];
    for (let c = 0; c < row.length; c += 1) {
      const raw = normalizeFormXIXGJText(row[c]);
      if (!raw) continue;

      for (let p = 0; p < patterns.length; p += 1) {
        if (!patterns[p].test(raw)) continue;

        const inline = stripFormXIXGJLabelPrefix(raw);
        if (inline && !isFormXIXGJBlockedValue(inline, blockedPatterns)) {
          return { index: p, value: inline, row: r, col: c };
        }

        const rightCells = row
          .slice(c + 1)
          .map((cell) => normalizeFormXIXGJText(cell))
          .filter(Boolean);
        // Excel often stores "1" | "Number of days worked........ 31"
        if (rightCells.length) {
          const joined = rightCells.join(' ');
          if (patterns[p].test(joined)) {
            const fromRight = stripFormXIXGJLabelPrefix(joined);
            if (fromRight && !isFormXIXGJBlockedValue(fromRight, blockedPatterns)) {
              return { index: p, value: fromRight, row: r, col: c };
            }
          }
          if (!isFormXIXGJBlockedValue(joined, blockedPatterns)) {
            return { index: p, value: joined, row: r, col: c };
          }
        }

        // Value on the next row (same band) — common for contractor / workman name.
        for (let nr = r + 1; nr <= Math.min(rows.length - 1, r + 2); nr += 1) {
          const nextCells = compactFormXIXGJCandidates(rows[nr]);
          if (!nextCells.length) continue;
          if (isFormXIXGJBlockedValue(nextCells[0], blockedPatterns)) break;
          // Prefer a single non-label value cell.
          const valueCells = nextCells.filter((cell) => !isFormXIXGJBlockedValue(cell, blockedPatterns));
          if (!valueCells.length) break;
          return { index: p, value: valueCells.join(' '), row: r, col: c };
        }
      }
    }
  }
  return null;
};

const collectFormXIXGJMappedValues = (rows = [], patterns = [], blockedPatterns = patterns) => {
  const values = patterns.map(() => '');
  const workingRows = Array.isArray(rows) ? rows.map((row) => (Array.isArray(row) ? [...row] : row)) : [];
  for (let i = 0; i < patterns.length; i += 1) {
    const one = findFormXIXGJPatternValue(workingRows, [patterns[i]], blockedPatterns);
    if (one && one.value) values[i] = one.value;
  }
  return values;
};

const sanitizeFormXIXGJWageValues = (values = []) => {
  const source = Array.isArray(values) ? values : [];
  return source.map((rawValue, currentIndex) => {
    let value = normalizeFormXIXGJText(rawValue);
    if (!value) return '';

    for (let wageIndex = 0; wageIndex < FORM_XIX_GJ_WAGE_PATTERNS.length; wageIndex += 1) {
      if (wageIndex === currentIndex) continue;
      const pattern = FORM_XIX_GJ_WAGE_PATTERNS[wageIndex];
      const match = value.match(pattern);
      if (!match) continue;
      const at = typeof match.index === 'number' ? match.index : -1;
      if (at <= 0) {
        value = '';
        break;
      }
      value = normalizeFormXIXGJText(value.slice(0, at));
    }
    return value;
  });
};

const formXIXGJContextBlob = (metaLines = [], rows = [], sheetName = '', fileName = '') =>
  [...(metaLines || []), ...(rows || []).slice(0, 50).flat(), sheetName || '', fileName || '']
    .map(lowerFormXIXGJText)
    .join(' ');

export const looksLikeFormXIXGJWageSlipPdfContext = (
  metaLines = [],
  rows = [],
  sheetName = '',
  fileName = ''
) => {
  const blob = formXIXGJContextBlob(metaLines, rows, sheetName, fileName);
  if (!blob) return false;
  const isXix = /form\s*xix\b|form[\s._-]*xix[\s._-]/.test(blob);
  if (!isXix) return false;
  // Never steal Madhya Pradesh / AP / Karnataka / Tamil Nadu XIX.
  if (
    /madhya\s+pradesh|form[\s._-]*xix[\s._-]*mp\b|xix_mp|andhra\s+pradesh|form[\s._-]*xix[\s._-]*ap\b|xix_ap|karnataka|form[\s._-]*xix[\s._-]*ka\b|tamil[\s._-]*nadu|form[\s._-]*xix[\s._-]*tamil/.test(
      blob
    )
  ) {
    return false;
  }
  const hasGj =
    /gujarat|form[\s._-]*xix[\s._-]*gj|xix_gj|[-_\s]gj\b|\bgj[-_\s]|\bgj\b/.test(blob);
  if (!hasGj) return false;
  const hasWageSlip = /wage\s*s?lip/.test(blob);
  const hasRule = /rule\s*78/.test(blob);
  const hasCoreRows =
    /name\s+and\s+address\s+(?:of|if)\s+(?:the\s+)?contractor/.test(blob) &&
    /nature\s+and\s+location\s+of\s+work/.test(blob) &&
    (/net\s+amount\s+of\s+wages|gross\s+wages\s+payable/.test(blob) ||
      /(?:number|no\.?)\s+of\s+days\s+worked/.test(blob) ||
      /rate\s+of\s+daily\s+wages/.test(blob));
  // Filename Form_XIX_GJ alone is enough when wage-slip markers exist.
  const namedGj = /form[\s._-]*xix[\s._-]*gj|xix_gj/.test(blob);
  if (namedGj && (hasWageSlip || hasRule || hasCoreRows)) return true;
  return (hasWageSlip || hasRule) && hasCoreRows;
};

export const normalizeFormXIXGJWageSlipPdfMatrix = (
  rows = [],
  colCount = 0,
  tableStartRow = 0,
  metaLines = []
) => {
  const sourceRows = Array.isArray(rows) ? rows : [];
  const sourceMeta = Array.isArray(metaLines) ? metaLines : [];
  const titleLines = [];

  sourceMeta.forEach((line) => {
    const text = normalizeFormXIXGJText(line);
    if (!text) return;
    if (looksLikeFormXIXGJTitle(text) || looksLikeFormXIXGJRule78(text) || looksLikeFormXIXGJWageSlip(text)) {
      titleLines.push(text);
    }
  });

  for (let r = 0; r < Math.min(sourceRows.length, Math.max(20, tableStartRow + 10)); r += 1) {
    const cells = compactFormXIXGJCandidates(sourceRows[r]);
    cells.forEach((text) => {
      if (looksLikeFormXIXGJTitle(text) || looksLikeFormXIXGJRule78(text) || looksLikeFormXIXGJWageSlip(text)) {
        titleLines.push(text);
      }
    });
  }

  const hasTitle = titleLines.some((line) => looksLikeFormXIXGJTitle(line));
  const hasRule = titleLines.some((line) => looksLikeFormXIXGJRule78(line));
  const hasWage = titleLines.some((line) => looksLikeFormXIXGJWageSlip(line));
  const orderedTitles = [];
  if (hasTitle) orderedTitles.push('FORM XIX');
  if (hasRule) {
    const ruleLine = titleLines.find((line) => looksLikeFormXIXGJRule78(line));
    orderedTitles.push(
      /78\s*\(\s*2\s*\)/.test(String(ruleLine || ''))
        ? '[See rule 78 (2)(b)]'
        : /78\s*\(\s*1\s*\)/.test(String(ruleLine || ''))
          ? normalizeFormXIXGJText(ruleLine)
          : '[See rule 78 (2)(b)]'
    );
  }
  if (hasWage) orderedTitles.push('Wage Slip');
  if (!orderedTitles.length) orderedTitles.push(...FORM_XIX_GJ_TITLES);

  const headerBlocked = [...FORM_XIX_GJ_HEADER_PATTERNS, ...FORM_XIX_GJ_WAGE_PATTERNS];
  const headerValues = collectFormXIXGJMappedValues(
    sourceRows,
    FORM_XIX_GJ_HEADER_PATTERNS,
    headerBlocked
  );

  const wageValues = sanitizeFormXIXGJWageValues(
    collectFormXIXGJMappedValues(sourceRows, FORM_XIX_GJ_WAGE_PATTERNS, FORM_XIX_GJ_WAGE_PATTERNS)
  );

  // MP-style box model: header fields + wage 1–7 all inside the bordered table.
  // Titles stay in meta (bordered title bands that sit with the table as one form box).
  // Initials stay outside below the box.
  const normalizedRows = [];
  FORM_XIX_GJ_HEADER_LABELS.forEach((label, idx) => {
    normalizedRows.push([label, headerValues[idx] || '']);
  });
  normalizedRows.push(['', '']);
  FORM_XIX_GJ_WAGE_LABELS.forEach((label, idx) => {
    normalizedRows.push([`${idx + 1} ${label}`, wageValues[idx] || '']);
  });

  return {
    rows: normalizedRows,
    colCount: 2,
    tableStartRow: 0,
    metaLines: orderedTitles,
    formXIXGJHeaderLines: [],
    // Outside / below the bordered form box — same placement as MP Form XIX.
    formXIXAPFooterLines: [FORM_XIX_GJ_FOOTER],
    formXIXAPLayout: true,
    // Keep GJ identity for detectors, but use the shared AP/MP bordered box painter.
    formXIXGJLayout: false,
    formXIXGJBoxedLayout: true,
    sourceColCount: Math.max(1, Number(colCount) || 1)
  };
};
