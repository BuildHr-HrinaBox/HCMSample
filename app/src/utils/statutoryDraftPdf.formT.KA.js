/**
 * Karnataka Form T — Combined Muster Roll cum Register of Wages.
 * Citation lines mention "Form XIII of Rules 75", which must never route this
 * sheet through the AP Form XIII Register of Workmen PDF layout.
 */

export const FORM_T_KA_PDF_TITLE = 'FORM T';
export const FORM_T_KA_PDF_SUBTITLE = 'COMBINED MUSTER ROLL CUM REGISTER OF WAGES';
export const FORM_T_KA_PDF_RULE =
  '[See Rule 24(9-B) of Karnataka Shops & Commercial Establishment Rules, 1963]';
export const FORM_T_KA_PDF_IN_LIEU_LINES = [
  'in lieu of',
  '1. Form I, II of Rule 22(4); Form IV of Rule 29(2); Forms V & VII of Rule 29(1) & (5) of Karnataka Minimum Wages Rules, 1958',
  '2. Form I of Rules 3(1) of Karnataka Payment of Wages Rules, 1963',
  '3. Form XIII of Rules 75; Form XV, XVII, XX, XXI, XXII, XXIII of 78(1)(a)(i), (ii) & (iii) of Karnataka Contract Labour (Regulation & Abolition) Rules, 1974',
  '4. Form XIII of Rule 43; Forms XVII, XVIII, XIX, XX, XXI, XXII of Rule 46(2)(a),(c) & (d) of Inter-state Migrant Workmen (Regulation of Employment and conditions of service) Karnataka Rules, 1981',
];

const normalizeText = (value) =>
  String(value || '')
    .replace(/\s+/g, ' ')
    .trim();

const blobFrom = (metaLines, rows, sheetName = '', fileName = '') =>
  [...(metaLines || []), ...(rows || []).slice(0, 18).flat(), sheetName || '', fileName || '']
    .map(normalizeText)
    .join(' ')
    .toLowerCase();

const isFormTFilename = (sheetName = '', fileName = '') => {
  const name = `${sheetName || ''} ${fileName || ''}`.toLowerCase();
  return (
    /form[\s._-]*t[\s._-]*ka\b/.test(name) ||
    /form[\s._-]*t[\s._-]*karnataka/.test(name) ||
    /form_{1,}t(?:_|-|$)/.test(name) ||
    /form[\s._-]+t[\s._-]+s[\s._-]*e\b/.test(name)
  );
};

export const looksLikeFormTSEKarnatakaPdfContext = (
  metaLines = [],
  rows = [],
  sheetName = '',
  fileName = ''
) => {
  if (isFormTFilename(sheetName, fileName)) return true;
  const blob = blobFrom(metaLines, rows, sheetName, fileName);
  if (!blob) return false;
  const isFormT = /(?:^|[^a-z0-9])form[\s._-]*t(?:[^a-z0-9]|$)/.test(blob);
  const isCombinedMuster = /combined\s+muster\s+roll\s+cum\s+register\s+of\s+wages/.test(blob);
  const isRule249B = /rule\s+24\s*\(\s*9[\s-]*b/.test(blob);
  if (isCombinedMuster && (isFormT || /karnataka/.test(blob) || isRule249B)) return true;
  if (isFormT && (isCombinedMuster || isRule249B)) return true;
  return false;
};

export const isFormTSEKarnatakaTableHeaderRow = (row = []) => {
  const blob = (Array.isArray(row) ? row : []).map(normalizeText).join(' ').toLowerCase();
  if (!blob) return false;
  const hasSerial = /(?:^|[^a-z])s\.?\s*no(?:[^a-z]|$)/.test(blob);
  const hasIdentity =
    /name of(?:\s+the)?\s+employee/.test(blob) ||
    /father/.test(blob) ||
    /principal\s+employer/.test(blob) ||
    /designation/.test(blob) ||
    /attendance/.test(blob);
  return hasSerial && hasIdentity;
};

export const isFormTSEKarnatakaAdminRowBlob = (blob) => {
  const t = String(blob || '').toLowerCase();
  if (!t) return false;
  if (isFormTSEKarnatakaTableHeaderRow([t])) return false;
  return (
    /month\s*\/\s*year/.test(t) ||
    /name\s+and\s+address\s+of\s+the\s+establishment/.test(t) ||
    /name\s+and\s+address\s+of\s+(?:the\s+)?employer/.test(t) ||
    (/address\s+of\s+the\s+establishment/.test(t) && !/principal/.test(t)) ||
    (/name\s+and\s+address\s+of\s+(?:the\s+)?contractor/.test(t) && !/principal/.test(t)) ||
    /nature\s+of\s+work\s+and\s+location/.test(t)
  );
};

const collectLines = (metaLines, rows, tableStart) => {
  const fromRows = (rows || [])
    .slice(0, Math.min(Math.max(Number(tableStart) || 0, 12), (rows || []).length))
    .map((row) =>
      (Array.isArray(row) ? row : [])
        .map(normalizeText)
        .filter(Boolean)
        .join(' ')
    )
    .filter(Boolean);
  const fromMeta = (metaLines || []).map(normalizeText).filter(Boolean);
  const seen = new Set();
  const out = [];
  [...fromMeta, ...fromRows].forEach((line) => {
    const key = line.toLowerCase();
    if (!key || seen.has(key)) return;
    seen.add(key);
    out.push(line);
  });
  return out;
};

const pickLine = (lines, pattern, fallback = '') => {
  const found = (lines || []).find((line) => pattern.test(line));
  return found ? normalizeText(found) : fallback;
};

export const getFormTSEKarnatakaHeaderTitles = (metaLines = [], rows = [], tableStart = 0) => {
  const lines = collectLines(metaLines, rows, tableStart);
  const formTitle = pickLine(lines, /^form\s*t\b/i, FORM_T_KA_PDF_TITLE);
  const subtitle = pickLine(
    lines,
    /combined\s+muster\s+roll|muster\s+roll\s+cum\s+register/i,
    FORM_T_KA_PDF_SUBTITLE
  );
  const rule = pickLine(lines, /rule\s+24\s*\(\s*9[\s-]*b/i, FORM_T_KA_PDF_RULE);
  const inLieu = pickLine(lines, /^in\s+lieu\s+of$/i, FORM_T_KA_PDF_IN_LIEU_LINES[0]);
  const citation1 = pickLine(
    lines,
    /^\s*1\.\s*form\b/i,
    FORM_T_KA_PDF_IN_LIEU_LINES[1]
  );
  const citation2 = pickLine(
    lines,
    /^\s*2\.\s*form\b/i,
    FORM_T_KA_PDF_IN_LIEU_LINES[2]
  );
  const citation3 = pickLine(
    lines,
    /^\s*3\.\s*form\b/i,
    FORM_T_KA_PDF_IN_LIEU_LINES[3]
  );
  const citation4 = pickLine(
    lines,
    /^\s*4\.\s*form\b/i,
    FORM_T_KA_PDF_IN_LIEU_LINES[4]
  );
  return [
    /^form\s*t\b/i.test(formTitle) ? FORM_T_KA_PDF_TITLE : formTitle,
    /combined\s+muster|muster\s+roll\s+cum\s+register/i.test(subtitle)
      ? FORM_T_KA_PDF_SUBTITLE
      : subtitle,
    rule,
    inLieu,
    citation1,
    citation2,
    citation3,
    citation4,
  ].filter(Boolean);
};

const FIELD_SPECS = [
  {
    key: 'monthYear',
    label: 'Month / Year',
    pattern: /month\s*\/\s*year/i,
  },
  {
    key: 'establishment',
    label: 'Name and address of the Establishment',
    pattern:
      /(?:name\s+and\s+)?address\s+of\s+the\s+establishment|name\s+and\s+address\s+of\s+(?:the\s+)?contractor/i,
  },
  {
    key: 'employer',
    label: 'Name and Address of employer',
    pattern:
      /name\s+and\s+address\s+of\s+(?:the\s+)?employer|nature\s+of\s+work\s+and\s+location\s+of\s+work/i,
  },
];

const splitFieldValue = (text, spec) => {
  const src = normalizeText(text);
  if (!src || !spec.pattern.test(src)) return '';
  const cut = src.replace(spec.pattern, '').replace(/^\s*[:.\-]+\s*/, '').trim();
  if (!cut || (spec.pattern.test(cut) && cut.length < 8)) return '';
  return cut;
};

export const extractFormTSEKarnatakaHeaderFields = (
  metaLines = [],
  rows = [],
  tableStart = 0
) => {
  const lines = collectLines(metaLines, rows, tableStart);
  const values = { monthYear: '', establishment: '', employer: '' };
  lines.forEach((line) => {
    FIELD_SPECS.forEach((spec) => {
      if (values[spec.key]) return;
      if (!spec.pattern.test(line)) return;
      const inline = splitFieldValue(line, spec);
      if (inline) values[spec.key] = inline;
    });
  });
  return FIELD_SPECS.map((spec) =>
    values[spec.key] ? `${spec.label} : ${values[spec.key]}` : `${spec.label} :`
  );
};

const isTitleOrCitationLine = (text) => {
  const t = normalizeText(text);
  if (!t) return false;
  if (/^form\s*t\b/i.test(t)) return true;
  if (/combined\s+muster\s+roll|muster\s+roll\s+cum\s+register\s+of\s+wages/i.test(t)) return true;
  if (/rule\s+24\s*\(\s*9[\s-]*b/i.test(t)) return true;
  if (/^in\s+lieu\s+of$/i.test(t)) return true;
  if (/^\s*[1-4]\.\s*form\b/i.test(t)) return true;
  return false;
};

const isLeakedFormXiiiTitle = (text) => {
  const t = normalizeText(text).toLowerCase();
  if (!t) return false;
  if (/register\s+of\s+workmen\s+employed\s+by\s+contractor/.test(t) && !/in\s+lieu/.test(t)) {
    return true;
  }
  if (/vide\s+rule\s*75/.test(t) && /central\s*\/\s*a\.?p/i.test(t)) return true;
  return false;
};

const isLoneDateFooter = (row) => {
  const filled = (Array.isArray(row) ? row : []).map(normalizeText).filter(Boolean);
  return filled.length === 1 && /^date\s*:?\s*$/i.test(filled[0]);
};

/** Official Form T: A–I identity, attendance starts at column J (0-based 9). */
export const FORM_T_KA_PDF_ATTENDANCE_START_COL0 = 9;

const cellAt = (row, col) => normalizeText(Array.isArray(row) ? row[col] : '');

const looksLikePdfEmployeeName = (text) => {
  const t = normalizeText(text);
  if (!t || t.length < 2 || t.length > 80) return false;
  if (/^\d+([./-]\d+)*$/.test(t)) return false;
  if (/^(male|female|p|a|wo|h|l|cl|el|sl|od|wop|nil|date:?)$/i.test(t)) return false;
  if (/s\.?\s*no|name of|father|husband|designation|attendance|wages fixed|esi no|uan no|principal employer/i.test(t)) {
    return false;
  }
  return /[a-z]/i.test(t);
};

const isFormTPdfHeaderOrIndexRow = (row) => {
  if (isFormTSEKarnatakaTableHeaderRow(row)) return true;
  const filled = (Array.isArray(row) ? row : []).map(normalizeText).filter(Boolean);
  if (filled.filter((c) => /^\d{1,2}$/.test(c)).length >= 8) return true;
  return false;
};

export const formTPdfRowLooksIdentityShiftedToJ = (
  row,
  shift = FORM_T_KA_PDF_ATTENDANCE_START_COL0
) => {
  if (!Array.isArray(row) || isFormTPdfHeaderOrIndexRow(row)) return false;
  const serialA = cellAt(row, 0);
  const nameB = cellAt(row, 1);
  const genderD = cellAt(row, 3);
  const serialJ = cellAt(row, shift);
  const nameK = cellAt(row, shift + 1);
  const genderM = cellAt(row, shift + 3);
  const personAtK = looksLikePdfEmployeeName(nameK);
  const personAtB = looksLikePdfEmployeeName(nameB);
  if (personAtK && !personAtB) return true;
  if (
    /^\d{1,4}$/.test(serialJ) &&
    !/^\d{1,4}$/.test(serialA) &&
    /^(male|female)$/i.test(genderM) &&
    !/^(male|female)$/i.test(genderD)
  ) {
    return true;
  }
  if (
    /^\d{1,4}$/.test(serialJ) &&
    !personAtB &&
    (personAtK || /^(male|female)$/i.test(genderM))
  ) {
    return true;
  }
  return false;
};

const shiftPdfRowFromJToA = (row, shift = FORM_T_KA_PDF_ATTENDANCE_START_COL0) => {
  const src = Array.isArray(row) ? row : [];
  const out = [];
  for (let c = 0; c < src.length; c += 1) {
    out[c] = src[c + shift] != null ? src[c + shift] : '';
  }
  return out;
};

/** Move S.NO/Name from column J onto column A when the Excel dump landed under attendance. */
export const shiftFormTSEKarnatakaPdfIdentityFromJToA = (
  rows,
  shift = FORM_T_KA_PDF_ATTENDANCE_START_COL0
) => {
  const src = Array.isArray(rows) ? rows : [];
  if (!src.length) return src;
  return src.map((row) =>
    formTPdfRowLooksIdentityShiftedToJ(row, shift) ? shiftPdfRowFromJToA(row, shift) : row
  );
};

const findTableStart = (rows) => {
  const src = Array.isArray(rows) ? rows : [];
  for (let r = 0; r < src.length; r += 1) {
    if (isFormTSEKarnatakaTableHeaderRow(src[r])) return r;
  }
  for (let r = 0; r < src.length; r += 1) {
    const filled = (src[r] || []).map((c) => String(c || '').trim()).filter(Boolean);
    if (filled.filter((c) => /^\d{1,2}$/.test(c)).length >= 20) {
      return Math.max(0, r - 1);
    }
  }
  return 0;
};

export const normalizeFormTSEKarnatakaPdfMatrix = (
  rows,
  colCount,
  tableStartRow = 0,
  incomingMetaLines = []
) => {
  const src = Array.isArray(rows) ? rows : [];
  const detectedStart = findTableStart(src);
  const start = detectedStart > 0 ? detectedStart : Math.max(0, Number(tableStartRow) || 0);
  const preRows = src.slice(0, start);
  const tableRows = shiftFormTSEKarnatakaPdfIdentityFromJToA(
    src.slice(start).filter((row) => !isLoneDateFooter(row))
  );
  const metaFromRows = preRows
    .map((row) =>
      (Array.isArray(row) ? row : [])
        .map(normalizeText)
        .filter(Boolean)
        .join(' ')
    )
    .filter(Boolean)
    .filter((line) => !isLeakedFormXiiiTitle(line));
  const incoming = (incomingMetaLines || [])
    .map(normalizeText)
    .filter(Boolean)
    .filter((line) => !isLeakedFormXiiiTitle(line));
  const metaLines = [];
  const seen = new Set();
  [...incoming, ...metaFromRows].forEach((line) => {
    const key = line.toLowerCase();
    if (!key || seen.has(key)) return;
    if (isLeakedFormXiiiTitle(line)) return;
    seen.add(key);
    metaLines.push(line);
  });
  const titles = getFormTSEKarnatakaHeaderTitles(metaLines, src, start);
  const fields = extractFormTSEKarnatakaHeaderFields(metaLines, src, start);
  const orderedMeta = [];
  [...titles, ...fields].forEach((line) => {
    const key = normalizeText(line).toLowerCase();
    if (!key || orderedMeta.some((x) => x.toLowerCase() === key)) return;
    orderedMeta.push(line);
  });
  metaLines.forEach((line) => {
    if (isTitleOrCitationLine(line) || isFormTSEKarnatakaAdminRowBlob(line)) return;
    const key = line.toLowerCase();
    if (orderedMeta.some((x) => x.toLowerCase() === key)) return;
    orderedMeta.push(line);
  });

  let maxCol = Math.max(1, Number(colCount) || 0);
  tableRows.forEach((row) => {
    if (!Array.isArray(row)) return;
    maxCol = Math.max(maxCol, row.length);
  });

  return {
    rows: tableRows.length ? tableRows : src,
    colCount: maxCol,
    tableStartRow: 0,
    metaLines: orderedMeta,
    formTKALayout: true,
  };
};

export const applyFormTSEKarnatakaPdfNormalization = (target, normalized) => {
  if (!target || !normalized) return target;
  target.rows = normalized.rows;
  target.colCount = normalized.colCount;
  target.tableStartRow = normalized.tableStartRow;
  if (Array.isArray(normalized.metaLines) && normalized.metaLines.length) {
    target.metaLines = normalized.metaLines;
  }
  target.formTKALayout = true;
  return target;
};
