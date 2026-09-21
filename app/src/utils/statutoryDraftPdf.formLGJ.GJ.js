/**
 * Gujarat Form L — List of Workers Engaged in Shift (See rule 14).
 * Excel templates keep an empty column A left of Sr. No.; PDF download
 * should start the worker table at Sr. No.
 */

export function looksLikeFormLGJGujaratPdfContext(
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
  const hasGujarat = /gujarat|\b_gj\b|form[\s._-]*l[\s._-]*gj|form_l_gj/.test(blob);
  const hasFormL = /\bform[\s._-]*l\b|form[\s._-]*l[\s._-]*gj|form_l_gj/.test(blob);
  const hasShiftList =
    /list of workers engaged in shift/.test(blob) || /workers engaged in shift/.test(blob);
  const hasRule14 = /see\s+rule\s+14/.test(blob);
  return Boolean(hasFormL && (hasGujarat || hasShiftList || hasRule14));
}

function isFormLGJSerialHeaderText(text) {
  const t = String(text || '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
  if (!t) return false;
  return /^(?:sr\.?\s*no|s\.?\s*no|serial\s*(?:no|number)|sl\.?\s*no)/.test(t);
}

export function isFormLGJDateOfMonthGroupLabel(text) {
  const t = String(text || '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
  return /^date\s+of\s+the\s+month$/.test(t) || /^date\s+of\s+month$/.test(t);
}

export function isFormLGJShiftGroupLabel(text) {
  return /^\d+(?:st|nd|rd|th)\s+shift$/i.test(
    String(text || '')
      .replace(/\s+/g, ' ')
      .trim()
  );
}

export function isFormLGJWeeklyHolidayHeaderText(text) {
  return /weekly\s+holiday/.test(
    String(text || '')
      .replace(/\s+/g, ' ')
      .trim()
      .toLowerCase()
  );
}

function isFormLGJIdentityHeaderText(text) {
  const t = String(text || '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
  return (
    isFormLGJSerialHeaderText(t) ||
    /name\s+of\s+the\s+worker/.test(t) ||
    /^designation$/.test(t) ||
    isFormLGJWeeklyHolidayHeaderText(t)
  );
}

function readFormLGJHeaderLeafAtCol(rows, fromRow, toRow, col) {
  for (let r = fromRow; r <= toRow; r += 1) {
    const t = String(rows[r]?.[col] || '')
      .replace(/\s+/g, ' ')
      .trim();
    if (!t || isFormLGJDateOfMonthGroupLabel(t)) continue;
    return t;
  }
  return '';
}

/**
 * One wide "Date of the Month" heading across 1st/2nd/3rd Shift columns
 * (Excel repeats the label per shift column). Never swallow Weekly holiday.
 */
export function detectFormLGJGujaratPdfGroupBands(rows, tableStart, headerBandEnd, colCount) {
  if (!Array.isArray(rows) || colCount < 2) return [];
  const start = Math.max(0, Number(tableStart) || 0);
  const end = Math.min(rows.length - 1, Math.max(start, Number(headerBandEnd) || 0));
  const dateCols = [];
  let dateRow = -1;
  for (let r = start; r <= end; r += 1) {
    for (let c = 0; c < colCount; c += 1) {
      if (!isFormLGJDateOfMonthGroupLabel(rows[r]?.[c])) continue;
      if (dateRow < 0) dateRow = r;
      if (r === dateRow) dateCols.push(c);
    }
    if (dateRow >= 0) break;
  }
  if (dateRow < 0) return [];

  let bandStart = dateCols[0];
  let bandEnd = dateCols[0];
  for (let i = 1; i < dateCols.length; i += 1) {
    if (dateCols[i] === bandEnd + 1) {
      bandEnd = dateCols[i];
    } else {
      break;
    }
  }
  if (dateCols.length === 1) {
    bandEnd = dateCols[0];
    for (let c = dateCols[0] + 1; c < colCount; c += 1) {
      const group = String(rows[dateRow]?.[c] || '')
        .replace(/\s+/g, ' ')
        .trim();
      if (group && !isFormLGJDateOfMonthGroupLabel(group)) break;
      const leaf = readFormLGJHeaderLeafAtCol(rows, dateRow, end, c);
      if (isFormLGJWeeklyHolidayHeaderText(group) || isFormLGJWeeklyHolidayHeaderText(leaf)) break;
      if (leaf && isFormLGJIdentityHeaderText(leaf) && !isFormLGJShiftGroupLabel(leaf)) break;
      bandEnd = c;
    }
  }
  // Trim trailing Weekly holiday / identity columns if Date-of-Month text leaked into them.
  while (bandEnd > bandStart) {
    const group = String(rows[dateRow]?.[bandEnd] || '')
      .replace(/\s+/g, ' ')
      .trim();
    const leaf = readFormLGJHeaderLeafAtCol(rows, dateRow, end, bandEnd);
    if (
      isFormLGJWeeklyHolidayHeaderText(group) ||
      isFormLGJWeeklyHolidayHeaderText(leaf) ||
      (leaf && isFormLGJIdentityHeaderText(leaf) && !isFormLGJShiftGroupLabel(leaf) && !isFormLGJDateOfMonthGroupLabel(leaf))
    ) {
      bandEnd -= 1;
      continue;
    }
    break;
  }
  if (!(bandEnd > bandStart)) return [];
  return [
    {
      labelRow: dateRow,
      start: bandStart,
      end: bandEnd,
      label: 'Date of the Month',
    },
  ];
}

/**
 * Excel often repeats "Weekly holiday day" down its merged column, so the PDF
 * paints it again on the 1st/2nd/3rd Shift row (looks like it sits under 3rd Shift).
 * Keep the label only on the top header row of that column; never inside shift cols.
 */
export function sanitizeFormLGJGujaratPdfHeaderRows(rows, colCount, tableStartRow = 0) {
  if (!Array.isArray(rows) || rows.length === 0 || colCount <= 0) return rows;
  const start = Math.max(0, Number(tableStartRow) || 0);
  const headerEnd = Math.min(rows.length - 1, start + 5);
  let weeklyHolidayCol = -1;
  const shiftCols = new Set();
  const shiftLabelRows = new Set();
  const fromToRows = new Set();

  for (let r = start; r <= headerEnd; r += 1) {
    const row = rows[r] || [];
    for (let c = 0; c < colCount; c += 1) {
      const t = String(row[c] || '')
        .replace(/\s+/g, ' ')
        .trim();
      if (!t) continue;
      if (isFormLGJWeeklyHolidayHeaderText(t) && weeklyHolidayCol < 0) weeklyHolidayCol = c;
      if (isFormLGJShiftGroupLabel(t)) {
        shiftCols.add(c);
        shiftLabelRows.add(r);
      }
      if (/from\s*[-–—]\s*to/i.test(t) || /^(from|to)(\s*[-–—]\s*(from|to)?)?$/i.test(t)) {
        fromToRows.add(r);
        shiftCols.add(c);
      }
    }
  }
  if (weeklyHolidayCol < 0 && shiftCols.size === 0) return rows;

  let weeklyHolidayTopRow = -1;
  if (weeklyHolidayCol >= 0) {
    for (let r = start; r <= headerEnd; r += 1) {
      if (isFormLGJWeeklyHolidayHeaderText(rows[r]?.[weeklyHolidayCol])) {
        weeklyHolidayTopRow = r;
        break;
      }
    }
  }

  return rows.map((row, rowIndex) => {
    if (rowIndex < start || rowIndex > headerEnd || !Array.isArray(row)) return row;
    const next = [...row];
    let changed = false;
    for (let c = 0; c < colCount; c += 1) {
      if (!isFormLGJWeeklyHolidayHeaderText(next[c])) continue;
      const inShiftCol = shiftCols.has(c);
      const isHolidayCol = c === weeklyHolidayCol;
      const onShiftOrFromToRow = shiftLabelRows.has(rowIndex) || fromToRows.has(rowIndex);
      const duplicateHoliday =
        isHolidayCol &&
        weeklyHolidayTopRow >= 0 &&
        rowIndex > weeklyHolidayTopRow;
      if (inShiftCol || (isHolidayCol && (onShiftOrFromToRow || duplicateHoliday))) {
        next[c] = '';
        changed = true;
      }
    }
    return changed ? next : row;
  });
}

/** Keep Name / shift / weekly-holiday heading boxes wide enough to read. */
export function formLGJGujaratColumnWeight(headerText, maxDataLen = 0) {
  const h = String(headerText || '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
  const dataLen = Math.max(0, Number(maxDataLen) || 0);
  if (isFormLGJSerialHeaderText(h)) return 3.8;
  if (/name\s+of\s+the\s+worker/.test(h)) return 16;
  if (/^designation$/.test(h)) return 10;
  if (isFormLGJWeeklyHolidayHeaderText(h)) return 12;
  if (
    isFormLGJDateOfMonthGroupLabel(h) ||
    isFormLGJShiftGroupLabel(h) ||
    /^(from|to)(\s*[-–—]\s*(from|to)?)?$/.test(h) ||
    /from\s*[-–—]\s*to/.test(h)
  ) {
    return Math.max(9, Math.min(11, 8 + Math.min(dataLen, 8) * 0.2));
  }
  return Math.max(7, Math.min(10, 6 + Math.min(dataLen, 10) * 0.25));
}

/**
 * Gujarat Form L templates keep an empty column A left of Sr. No.
 * Drop leading spacer columns so the PDF table starts at Sr. No.
 */
export function trimFormLGJGujaratLeadingBlankPdfColumns(rows, colCount, tableStartRow = 0) {
  if (!Array.isArray(rows) || rows.length === 0 || colCount <= 1) {
    return { rows, colCount };
  }
  const startRow = Math.max(0, Number(tableStartRow) || 0);
  let lead = 0;
  outer: for (let r = startRow; r < Math.min(rows.length, startRow + 8); r += 1) {
    const row = rows[r] || [];
    for (let c = 0; c < colCount; c += 1) {
      if (isFormLGJSerialHeaderText(row[c])) {
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
