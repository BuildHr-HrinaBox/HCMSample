/**
 * Gujarat Form O — Notice of Maximum Leave Accumulated (See rule 18).
 * Excel prints Address → Authorized person → Notice → legal text → Details of Workers,
 * then the worker table starting at Sr. No. Templates keep an empty column A;
 * PDF download should drop that spacer and keep the Excel field order.
 */

export const FORM_OGJ_GJ_PDF_TITLE = 'FORM - O';
export const FORM_OGJ_GJ_PDF_RULE = '(See rule 18)';
export const FORM_OGJ_GJ_PDF_SUBTITLE = 'NOTICE OF MAXIMUM LEAVE ACCUMULATED';

const norm = (text) =>
  String(text || '')
    .replace(/_x000d_/gi, ' ')
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

const lower = (text) => norm(text).toLowerCase();

export function looksLikeFormOGJGujaratPdfContext(
  metaLines,
  rows,
  sheetName = '',
  fileName = ''
) {
  const blob = [
    ...(metaLines || []),
    ...(rows || []).slice(0, 22).flat(),
    sheetName || '',
    fileName || '',
  ]
    .join(' ')
    .toLowerCase();
  if (!blob) return false;
  if (/\bform\s*q\b|\bform_q\b|form[\s._-]*q[\s._-]*gj/.test(blob)) return false;
  if (/\bform\s*p\b|\bform_p\b|form[\s._-]*p[\s._-]*gj/.test(blob)) return false;
  if (/\bform\s*n\b/.test(blob) && /leave\s+book|see\s+rule\s+17/.test(blob)) return false;
  if (/\bform\s*k\b/.test(blob) && /weekly\s+holiday|see\s+rule\s+12/.test(blob)) return false;

  const hasGujarat = /gujarat|\b_gj\b|form[\s._-]*o[\s._-]*gj|form_o_gj/.test(blob);
  const hasFormO =
    /\bform[\s._-]*o\b/.test(blob) || /form[\s._-]*o[\s._-]*gj/.test(blob) || /form_o_gj/.test(blob);
  const hasNotice = /notice of maximum leave accumulated/.test(blob);
  const hasRule18 = /see\s+rule\s+18/.test(blob);
  if (hasGujarat && hasFormO) return true;
  if (hasFormO && (hasNotice || hasRule18)) return true;
  return (
    /sr\.?\s*no/.test(blob) &&
    /name\s+of\s+workers?/.test(blob) &&
    /accumulated\s+leave/.test(blob) &&
    (hasNotice || hasRule18)
  );
}

function classifyFormOGJHeaderLine(text) {
  const t = lower(text);
  if (!t) return 'other';
  if (/^form\s*-?\s*o\b/.test(t) && t.length < 40) return 'title';
  if (/see\s+rule\s*18/.test(t)) return 'rule';
  if (/notice of maximum leave accumulated/.test(t)) return 'subtitle';
  if (/address of the establishment/.test(t)) return 'address';
  if (/name of the authorized|authorized person/.test(t)) return 'authorized';
  if (/^notice$/.test(t)) return 'notice';
  if (/details of workers/.test(t)) return 'details';
  if (
    /as per section\s*18/.test(t) ||
    /maximum leave that can be accumulated/.test(t) ||
    /gujarat shops and establishments/.test(t)
  ) {
    return 'legal';
  }
  return 'other';
}

/**
 * Excel order below the title band:
 * Address, Authorized person, Notice, legal paragraph, Details of Workers.
 */
export function rewriteFormOGJGujaratPdfHeader(titles, fields) {
  const pool = [...(Array.isArray(titles) ? titles : []), ...(Array.isArray(fields) ? fields : [])]
    .map((line) => norm(line))
    .filter(Boolean);

  const nextTitles = [];
  const buckets = {
    address: [],
    authorized: [],
    notice: [],
    legal: [],
    details: [],
    other: [],
  };
  const seen = new Set();

  pool.forEach((line) => {
    const key = lower(line);
    if (seen.has(key)) return;
    seen.add(key);
    const kind = classifyFormOGJHeaderLine(line);
    if (kind === 'title' || kind === 'rule' || kind === 'subtitle') {
      nextTitles.push(line);
      return;
    }
    if (buckets[kind]) buckets[kind].push(line);
    else buckets.other.push(line);
  });

  const titleRank = (line) => {
    const kind = classifyFormOGJHeaderLine(line);
    if (kind === 'title') return 0;
    if (kind === 'rule') return 1;
    if (kind === 'subtitle') return 2;
    return 9;
  };
  nextTitles.sort((a, b) => titleRank(a) - titleRank(b));

  return {
    titles: nextTitles,
    fields: [
      ...buckets.address,
      ...buckets.authorized,
      ...buckets.notice,
      ...buckets.legal,
      ...buckets.other,
      ...buckets.details,
    ],
  };
}

function isFormOGJSerialHeaderText(text) {
  const t = lower(text);
  if (!t) return false;
  return /^(?:sr\.?\s*no|s\.?\s*no|serial\s*(?:no|number)|sl\.?\s*no)/.test(t);
}

/**
 * Gujarat Form O templates keep an empty column A left of Sr. No.
 * Drop leading spacer columns so the PDF table starts at Sr. No.
 */
export function trimFormOGJGujaratLeadingBlankPdfColumns(rows, colCount, tableStartRow = 0) {
  if (!Array.isArray(rows) || rows.length === 0 || colCount <= 1) {
    return { rows, colCount };
  }
  const startRow = Math.max(0, Number(tableStartRow) || 0);
  let lead = 0;
  outer: for (let r = startRow; r < Math.min(rows.length, startRow + 8); r += 1) {
    const row = rows[r] || [];
    for (let c = 0; c < colCount; c += 1) {
      if (isFormOGJSerialHeaderText(row[c])) {
        lead = c;
        break outer;
      }
    }
  }
  if (lead <= 0) {
    for (let c = 0; c < colCount; c += 1) {
      let hasContent = false;
      for (let r = startRow; r < rows.length; r += 1) {
        const row = rows[r];
        if (!row) continue;
        if (String(row[c] || '').trim()) {
          hasContent = true;
          break;
        }
      }
      if (hasContent) {
        lead = c;
        break;
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
