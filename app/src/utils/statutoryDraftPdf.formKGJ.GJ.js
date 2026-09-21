/**
 * Gujarat Form K — Notice of Weekly Holiday PDF/Excel caption.
 * Excel wrapText splits the statutory sentence into a title band under
 * "NOTICE OF WEEKLY HOLIDAY" plus a short caption above the worker table.
 * PDF download should keep only the full sentence above the table.
 */

export const FORM_KGJ_GJ_WEEKLY_HOLIDAY_INTRO =
  'All the workers in the establishment are hereby informed that the weekly holiday for each worker is given below:';

export function normalizeFormKGJNoticeText(text) {
  return String(text || '')
    .replace(/_x000d_/gi, ' ')
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .replace(/[:.]+$/g, '')
    .toLowerCase();
}

/** Truncated title under NOTICE OF WEEKLY HOLIDAY (stops at "the weekly"). */
export function isFormKGJTruncatedWeeklyIntroText(text) {
  const t = normalizeFormKGJNoticeText(text);
  if (!t) return false;
  if (/holiday for each worker/.test(t)) return false;
  return /all the workers in the establishment are hereby informed that the weekly/.test(t);
}

/** Short caption currently printed above the worker table. */
export function isFormKGJHolidayCaptionText(text) {
  const t = normalizeFormKGJNoticeText(text);
  if (!t) return false;
  if (/all the workers/.test(t)) return false;
  return /holiday for each worker is given below/.test(t);
}

export function isFormKGJFullWeeklyIntroText(text) {
  const t = normalizeFormKGJNoticeText(text);
  return /all the workers in the establishment are hereby informed that the weekly holiday for each worker/.test(
    t
  );
}

export function looksLikeFormKGJGujaratPdfContext(
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
  const hasGujarat = /gujarat|\b_gj\b|form[\s._-]*k[\s._-]*gj|form_k_gj/.test(blob);
  const hasFormK = /\bform[\s._-]*k\b|form[\s._-]*k[\s._-]*gj|form_k_gj/.test(blob);
  const hasWeeklyNotice =
    /notice\s+of\s+weekly\s+holiday/.test(blob) ||
    /holiday for each worker is given below/.test(blob);
  return Boolean(hasFormK && (hasGujarat || hasWeeklyNotice));
}

export function rewriteFormKGJGujaratPdfHeader(titles, fields) {
  const nextTitles = (Array.isArray(titles) ? titles : []).filter(
    (t) =>
      !isFormKGJTruncatedWeeklyIntroText(t) && !isFormKGJFullWeeklyIntroText(t)
  );

  let replaced = false;
  const nextFields = [];
  (Array.isArray(fields) ? fields : []).forEach((raw) => {
    const line = String(raw || '').replace(/\s+/g, ' ').trim();
    if (!line) return;
    if (
      isFormKGJHolidayCaptionText(line) ||
      isFormKGJTruncatedWeeklyIntroText(line) ||
      isFormKGJFullWeeklyIntroText(line)
    ) {
      if (!replaced) {
        nextFields.push(FORM_KGJ_GJ_WEEKLY_HOLIDAY_INTRO);
        replaced = true;
      }
      return;
    }
    nextFields.push(line);
  });

  if (!replaced) nextFields.push(FORM_KGJ_GJ_WEEKLY_HOLIDAY_INTRO);
  return { titles: nextTitles, fields: nextFields };
}

function isFormKGJSerialHeaderText(text) {
  const t = String(text || '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
  if (!t) return false;
  return /^(?:sr\.?\s*no|s\.?\s*no|serial\s*(?:no|number)|sl\.?\s*no)/.test(t);
}

/**
 * Gujarat Form K templates keep an empty column A left of Sr. No.
 * Drop leading spacer columns so the PDF table starts at Sr. No.
 */
const FORM_KGJ_GJ_WEEKLY_HOLIDAY_DEFAULT = 'Saturday & Sunday';

function isFormKGJHoursOfWorkHeaderText(text) {
  const s = String(text || '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9 ]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  return /hours?\s+of\s+work/.test(s) || (/\bhours?\b/.test(s) && /\bwork\b/.test(s));
}

/**
 * Template sample rows sometimes copy weekly holiday into the Hours column.
 * Clear that placeholder so PDF shows fetched shift hours (or blank).
 */
export function sanitizeFormKGJGujaratPdfDataRows(
  rows,
  colCount,
  tableStartRow = 0,
  weeklyHolidayDefault = FORM_KGJ_GJ_WEEKLY_HOLIDAY_DEFAULT
) {
  if (!Array.isArray(rows) || rows.length === 0 || colCount <= 0) return rows;
  const startRow = Math.max(0, Number(tableStartRow) || 0);
  let hoursCol = -1;
  for (let r = startRow; r < Math.min(rows.length, startRow + 6); r += 1) {
    const row = rows[r] || [];
    for (let c = 0; c < colCount; c += 1) {
      if (isFormKGJHoursOfWorkHeaderText(row[c])) {
        hoursCol = c;
        break;
      }
    }
    if (hoursCol >= 0) break;
  }
  if (hoursCol < 0) return rows;
  const holidayNorm = String(weeklyHolidayDefault || '').trim().toLowerCase();
  return rows.map((row, rowIndex) => {
    if (rowIndex <= startRow || !Array.isArray(row)) return row;
    const cell = String(row[hoursCol] ?? '').trim();
    if (!cell) return row;
    if (cell.toLowerCase() === holidayNorm) {
      const next = [...row];
      next[hoursCol] = '';
      return next;
    }
    return row;
  });
}

export function trimFormKGJGujaratLeadingBlankPdfColumns(rows, colCount, tableStartRow = 0) {
  if (!Array.isArray(rows) || rows.length === 0 || colCount <= 1) {
    return { rows, colCount };
  }
  const startRow = Math.max(0, Number(tableStartRow) || 0);
  let lead = 0;
  outer: for (let r = startRow; r < Math.min(rows.length, startRow + 8); r += 1) {
    const row = rows[r] || [];
    for (let c = 0; c < colCount; c += 1) {
      if (isFormKGJSerialHeaderText(row[c])) {
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
