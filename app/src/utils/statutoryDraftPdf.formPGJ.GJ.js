/**
 * Gujarat Form P — Muster-Roll cum Wage Register PDF (Excel model).
 * Matches the downloaded Excel: titles, establishment / employer / month,
 * then Sr. No through Signature with Date of Month days 1–31.
 */

function blobText(metaLines, rows, sheetName, fileName) {
  return [
    ...(metaLines || []),
    ...(rows || []).slice(0, 24).flat(),
    sheetName || '',
    fileName || '',
  ]
    .join(' ')
    .toLowerCase();
}

export function looksLikeFormPGJGujaratPdfContext(
  metaLines,
  rows,
  sheetName = '',
  fileName = ''
) {
  const blob = blobText(metaLines, rows, sheetName, fileName);
  if (/\bform\s*q\b|form[\s._-]*q[\s._-]|form_q_/.test(blob) && !/form[\s._-]*p/.test(blob)) {
    return false;
  }
  if (/format\s+of\s+wage\s+register/.test(blob) && /\bform\s*b\b/.test(blob)) return false;
  if (/see\s+rule\s*20/.test(blob) && /notice of maximum leave accumulated/.test(blob)) return false;
  if (/karnataka/.test(blob) && /shri\s*\/?\s*smt/.test(blob) && !/muster/.test(blob)) return false;
  const hasFormP =
    /\bform\s*[-–.]?\s*p\b/.test(blob) || /form[\s._-]*p[\s._-]*gj|form_p_gj/.test(blob);
  const hasMuster = /muster[\s-]*roll\s+cum\s+wage\s+register/.test(blob);
  const hasGujarat = /gujarat|\b_gj\b|form[\s._-]*p[\s._-]*gj|form_p_gj/.test(blob);
  const hasFingerprint =
    (/interval\s+for\s+rest/.test(blob) || /date\s+of\s+(the\s+)?month/.test(blob)) &&
    (/full\s+name\s+of\s+the\s+worker/.test(blob) || /name\s+of\s+the\s+worker/.test(blob));
  if (hasFormP && (hasMuster || hasFingerprint)) return true;
  if (hasGujarat && hasMuster && hasFingerprint) return true;
  return Boolean(hasFingerprint && /see\s+rules?\s*26/.test(blob));
}

export const FORM_PGJ_DATE_OF_MONTH_GROUP_LABEL = 'Date of Month (9)';

export function isFormPGJDateOfMonthGroupLabel(text) {
  const t = String(text || '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
  if (!t) return false;
  if (/date\s+of\s+payment/.test(t) || /entry\s+into\s+service/.test(t)) return false;
  // Excel: "Date of Month (9)"; UI: "Date of Month_1" … "Date of Month_31"
  return /^date\s+of\s+(the\s+)?month(?:\s*\(\s*9\s*\))?(?:[_\s]+\d{1,2})?$/.test(t);
}

export function isFormPGJSerialHeaderText(text) {
  const t = String(text || '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
  return /^(?:sr\.?\s*no|s\.?\s*no|serial\s*(?:no|number))/.test(t);
}

function isFormPGJWageTailHeaderText(text) {
  const t = String(text || '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
  if (!t) return false;
  return (
    /total\s+days?\s+worked/.test(t) ||
    /minimum\s+rate\s+of\s+wages/.test(t) ||
    /total\s+production/.test(t) ||
    /actual\s+wages\s+paid/.test(t) ||
    /house\s+rent/.test(t) ||
    /dearness/.test(t) ||
    /gross\s+amount/.test(t) ||
    /overtime/.test(t) ||
    /provident\s+fund/.test(t) ||
    /net\s+payable/.test(t) ||
    /date\s+of\s+payment/.test(t) ||
    /signature|thumb/.test(t)
  );
}

export function isFormPGJMinimumRateOfWagesHeader(text) {
  return /minimum\s+rate\s+of\s+wages/.test(
    String(text || '')
      .replace(/\s+/g, ' ')
      .trim()
      .toLowerCase()
  );
}

export function isFormPGJTotalDaysWorkedHeader(text) {
  return /total\s+days?\s+worked/.test(
    String(text || '')
      .replace(/\s+/g, ' ')
      .trim()
      .toLowerCase()
  );
}

function formPGJPdfSkipWageHeaderMeta(text) {
  const t = String(text || '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
  if (isFormPGJMinimumRateOfWagesHeader(t)) return { keepRe: /^\(?\s*11\s*\)?$/ };
  if (/total\s+production/.test(t) && /piece/.test(t)) return { keepRe: /^\(?\s*12\s*\)?$/ };
  return null;
}

const FORM_PGJ_PDF_MONTH_INDEX = {
  jan: 0,
  january: 0,
  feb: 1,
  february: 1,
  mar: 2,
  march: 2,
  apr: 3,
  april: 3,
  may: 4,
  jun: 5,
  june: 5,
  jul: 6,
  july: 6,
  aug: 7,
  august: 7,
  sep: 8,
  sept: 8,
  september: 8,
  oct: 9,
  october: 9,
  nov: 10,
  november: 10,
  dec: 11,
  december: 11,
};

export function resolveFormPGJGujaratPdfDaysInMonth(metaLines, rows) {
  const blob = [...(metaLines || []), ...(rows || []).slice(0, 16).flat()]
    .join(' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
  const monthMatch = blob.match(
    /month\s*:?\s*(january|february|march|april|may|june|july|august|september|october|november|december|jan|feb|mar|apr|jun|jul|aug|sept?|oct|nov|dec)\b/i
  );
  const token = monthMatch
    ? monthMatch[1]
    : (blob.match(
        /\b(january|february|march|april|may|june|july|august|september|october|november|december)\b/i
      ) || [])[1];
  if (!token) return 31;
  const monthIdx = FORM_PGJ_PDF_MONTH_INDEX[String(token).toLowerCase()];
  if (!Number.isFinite(monthIdx)) return 31;
  const yearMatch = blob.match(/\b(20\d{2}|19\d{2})\b/);
  const year = yearMatch ? Number(yearMatch[1]) : new Date().getFullYear();
  const days = new Date(year, monthIdx + 1, 0).getDate();
  return Math.min(Math.max(days, 28), 31);
}

function isDayNumberLeaf(text) {
  const t = String(text || '')
    .replace(/\s+/g, ' ')
    .trim();
  if (!/^\(?\s*\d{1,2}\s*\)?$/.test(t)) return false;
  const n = parseInt(t.replace(/[()]/g, ''), 10);
  return n >= 1 && n <= 31;
}

function readHeaderLeaf(rows, fromRow, toRow, col) {
  for (let r = fromRow; r <= toRow && r < rows.length; r += 1) {
    const t = String(rows[r]?.[col] || '')
      .replace(/\s+/g, ' ')
      .trim();
    if (!t) continue;
    if (isFormPGJDateOfMonthGroupLabel(t)) continue;
    return t;
  }
  return '';
}

/**
 * Excel "Date of Month (9)" merge across days 1–31 (stop before wage-tail columns).
 */
export function detectFormPGJGujaratPdfGroupBands(rows, tableStart, headerBandEnd, colCount) {
  if (!Array.isArray(rows) || colCount < 8) return [];
  const start = Math.max(0, Number(tableStart) || 0);
  const end = Math.min(rows.length - 1, Math.max(start, Number(headerBandEnd) || 0));
  let labelRow = -1;
  let bandStart = -1;
  for (let r = start; r <= end; r += 1) {
    for (let c = 0; c < colCount; c += 1) {
      if (!isFormPGJDateOfMonthGroupLabel(rows[r]?.[c])) continue;
      labelRow = r;
      bandStart = c;
      break;
    }
    if (labelRow >= 0) break;
  }
  if (labelRow < 0) return [];

  let bandEnd = bandStart;
  for (let c = bandStart + 1; c < colCount; c += 1) {
    const group = String(rows[labelRow]?.[c] || '')
      .replace(/\s+/g, ' ')
      .trim();
    if (
      group &&
      !isFormPGJDateOfMonthGroupLabel(group) &&
      !isDayNumberLeaf(group)
    ) {
      break;
    }
    const leaf = readHeaderLeaf(rows, labelRow, end, c);
    if (isFormPGJWageTailHeaderText(leaf) || isFormPGJWageTailHeaderText(group)) break;
    if (isFormPGJSerialHeaderText(leaf) || isFormPGJSerialHeaderText(group)) break;
    if (
      leaf &&
      !isDayNumberLeaf(leaf) &&
      !isFormPGJDateOfMonthGroupLabel(leaf)
    ) {
      break;
    }
    bandEnd = c;
  }
  if (!(bandEnd > bandStart)) return [];
  return [
    {
      labelRow,
      start: bandStart,
      end: bandEnd,
      label: FORM_PGJ_DATE_OF_MONTH_GROUP_LABEL,
    },
  ];
}

/** Keep name/designation readable; day cells stay narrow. */
export function formPGJGujaratColumnWeight(headerText, maxDataLen = 0) {
  const h = String(headerText || '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
  const dataLen = Math.max(0, Number(maxDataLen) || 0);
  if (isFormPGJSerialHeaderText(h)) return 3.4;
  if (/full\s+name\s+of\s+the\s+worker|name\s+of\s+the\s+worker/.test(h)) return 14;
  if (/^designation/.test(h)) return 8;
  if (/^age$/.test(h) || /^sex$/.test(h)) return 3.2;
  if (/entry\s+into\s+service/.test(h)) return 8;
  if (/working\s+hours/.test(h)) return 8;
  if (/interval\s+for\s+rest/.test(h)) return 5.5;
  if (isDayNumberLeaf(h) || isFormPGJDateOfMonthGroupLabel(h)) {
    return Math.max(2, Math.min(2.4, 2 + Math.min(dataLen, 4) * 0.1));
  }
  if (/signature|thumb/.test(h)) return 8;
  if (/date\s+of\s+payment/.test(h)) return 7;
  if (isFormPGJWageTailHeaderText(h)) {
    return Math.max(5.5, Math.min(8, 5 + Math.min(dataLen, 10) * 0.2));
  }
  return Math.max(4.5, Math.min(8, 4 + Math.min(dataLen, 12) * 0.2));
}

function fieldRank(text) {
  const t = String(text || '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
  if (/name\s+of\s+(?:the\s+)?establishment/.test(t)) return 0;
  if (/name\s+of\s+(?:the\s+)?employer/.test(t) && !/address/.test(t)) return 1;
  if (/^month\b/.test(t)) return 2;
  return 50;
}

function isFormPGJTitleLine(text) {
  const t = String(text || '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
  if (!t) return false;
  if (/^form\s*[-–.]?\s*p\b/.test(t) && t.length < 40) return true;
  if (/see\s+rules?\s*26/.test(t)) return true;
  if (/muster[\s-]*roll\s+cum\s+wage\s+register/.test(t)) return true;
  return false;
}

function titleRank(text) {
  const t = String(text || '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
  if (/^form\s*[-–.]?\s*p\b/.test(t)) return 0;
  if (/see\s+rules?\s*26/.test(t)) return 1;
  if (/muster[\s-]*roll/.test(t)) return 2;
  return 9;
}

/**
 * Excel header order: FORM P / See rules 26 / MUSTER-ROLL … then
 * Name of the Establishment, Name of the employer, Month (left fields, not a right box).
 */
export function rewriteFormPGJGujaratPdfHeader(titles, fields) {
  const pool = [...(Array.isArray(titles) ? titles : []), ...(Array.isArray(fields) ? fields : [])]
    .map((line) => String(line || '').replace(/\s+/g, ' ').trim())
    .filter(Boolean);

  const nextTitles = [];
  const nextFields = [];
  const seen = new Set();
  pool.forEach((line) => {
    const key = line.toLowerCase();
    if (seen.has(key)) return;
    seen.add(key);
    if (isFormPGJTitleLine(line)) {
      nextTitles.push(line);
      return;
    }
    if (/rate\s+of\s+minimum\s+wages|highly\s+skilled|minimum\s+basic/.test(key) && !/minimum\s+rate\s+of\s+wages/.test(key)) {
      return;
    }
    nextFields.push(line);
  });
  nextTitles.sort((a, b) => titleRank(a) - titleRank(b));
  nextFields.sort((a, b) => fieldRank(a) - fieldRank(b));
  return { titles: nextTitles, fields: nextFields };
}

export function trimFormPGJGujaratLeadingBlankPdfColumns(rows, colCount, tableStartRow = 0) {
  if (!Array.isArray(rows) || rows.length === 0 || colCount <= 1) {
    return { rows, colCount };
  }
  const startRow = Math.max(0, Number(tableStartRow) || 0);
  let lead = 0;
  outer: for (let r = startRow; r < Math.min(rows.length, startRow + 8); r += 1) {
    const row = rows[r] || [];
    for (let c = 0; c < colCount; c += 1) {
      if (isFormPGJSerialHeaderText(row[c])) {
        lead = c;
        break outer;
      }
    }
  }
  if (lead <= 0) return { rows, colCount };
  const nextCount = Math.max(1, colCount - lead);
  const trimmed = rows.map((row) => {
    const src = Array.isArray(row) ? row : [];
    return Array.from({ length: nextCount }, (_, index) => String(src[lead + index] ?? ''));
  });
  return { rows: trimmed, colCount: nextCount };
}

function dropFormPGJPdfColumns(rows, colCount, dropCols) {
  const drop = new Set((dropCols || []).filter((c) => Number.isFinite(c) && c >= 0 && c < colCount));
  if (drop.size === 0) return { rows, colCount };
  const keep = [];
  for (let c = 0; c < colCount; c += 1) {
    if (!drop.has(c)) keep.push(c);
  }
  return {
    rows: (rows || []).map((row) => {
      const src = Array.isArray(row) ? row : [];
      return keep.map((c) => String(src[c] ?? ''));
    }),
    colCount: keep.length,
  };
}

function readFormPGJPdfHeaderLeaf(rows, fromRow, toRow, col) {
  for (let r = fromRow; r <= toRow && r < rows.length; r += 1) {
    const t = String(rows[r]?.[col] || '')
      .replace(/\s+/g, ' ')
      .trim();
    if (!t) continue;
    if (isFormPGJDateOfMonthGroupLabel(t)) continue;
    return t;
  }
  return '';
}

/**
 * Drop Date of Month columns after the selected month's last day (June=30, July=31).
 */
export function dropFormPGJGujaratExtraDayPdfColumns(
  rows,
  colCount,
  tableStartRow = 0,
  daysInMonth = 31
) {
  if (!Array.isArray(rows) || colCount < 8) return { rows, colCount };
  const days = Math.min(Math.max(Number(daysInMonth) || 31, 28), 31);
  if (days >= 31) return { rows, colCount };
  const start = Math.max(0, Number(tableStartRow) || 0);
  const headerEnd = Math.min(rows.length - 1, start + 6);
  const bands = detectFormPGJGujaratPdfGroupBands(rows, start, headerEnd, colCount);
  let bandStart = bands[0]?.start;
  let bandEnd = bands[0]?.end;
  if (!(bandEnd > bandStart)) {
    for (let r = start; r <= headerEnd; r += 1) {
      const row = rows[r] || [];
      const idx = row.findIndex((cell) => String(cell || '').trim() === '1');
      if (idx < 0) continue;
      let end = idx;
      for (let n = 1; n <= 31 && idx + n - 1 < colCount; n += 1) {
        if (String(row[idx + n - 1] || '').trim() !== String(n)) break;
        end = idx + n - 1;
      }
      if (end - idx + 1 >= 20) {
        bandStart = idx;
        bandEnd = end;
        break;
      }
    }
  }
  if (!(bandEnd > bandStart)) return { rows, colCount };

  const drop = [];
  let lastKeep = -1;
  for (let c = bandStart; c <= bandEnd; c += 1) {
    const leaf = readFormPGJPdfHeaderLeaf(rows, start, headerEnd, c);
    const n = parseInt(String(leaf).replace(/[()]/g, ''), 10);
    if (/^\(?\s*\d{1,2}\s*\)?$/.test(leaf) && n >= 1 && n <= 31) {
      if (n > days) drop.push(c);
      else lastKeep = c;
    }
  }
  if (lastKeep >= 0) {
    for (let c = lastKeep + 1; c <= bandEnd; c += 1) {
      if (!drop.includes(c)) drop.push(c);
    }
  }
  return dropFormPGJPdfColumns(rows, colCount, drop);
}

/** Keep skip-column headings, clear employee values (Min wages / piece-rate production). */
export function blankFormPGJGujaratMinimumRateOfWagesPdfCells(rows, tableStartRow = 0, colCount = 0) {
  if (!Array.isArray(rows) || colCount < 1) return rows;
  const start = Math.max(0, Number(tableStartRow) || 0);
  const headerEnd = Math.min(rows.length - 1, start + 6);
  const wageCols = [];
  for (let c = 0; c < colCount; c += 1) {
    for (let r = start; r <= headerEnd; r += 1) {
      const meta = formPGJPdfSkipWageHeaderMeta(rows[r]?.[c]);
      if (!meta) continue;
      wageCols.push({ c, keepRe: meta.keepRe, isHeader: (t) => Boolean(formPGJPdfSkipWageHeaderMeta(t)) });
      break;
    }
  }
  if (!wageCols.length) return rows;
  wageCols.forEach(({ c, keepRe, isHeader }) => {
    for (let r = start; r < rows.length; r += 1) {
      const t = String(rows[r]?.[c] ?? '').trim();
      if (!t) continue;
      if (isHeader(t)) continue;
      if (keepRe.test(t)) continue;
      rows[r][c] = '';
    }
  });
  return rows;
}

export function normalizeFormPGJGujaratPdfMatrix(rows, colCount, tableStartRow = 0, metaLines = []) {
  const lead = trimFormPGJGujaratLeadingBlankPdfColumns(rows, colCount, tableStartRow);
  const days = resolveFormPGJGujaratPdfDaysInMonth(metaLines, lead.rows);
  const dropped = dropFormPGJGujaratExtraDayPdfColumns(
    lead.rows,
    lead.colCount,
    tableStartRow,
    days
  );
  blankFormPGJGujaratMinimumRateOfWagesPdfCells(dropped.rows, tableStartRow, dropped.colCount);
  return dropped;
}
