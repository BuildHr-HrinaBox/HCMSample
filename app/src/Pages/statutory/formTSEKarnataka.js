import * as XLSX from 'xlsx';

/** Karnataka Form T — Combined Muster Roll cum Register of Wages (attendance day grid). */

function normCell(txt) {
  return String(txt || '').replace(/\s+/g, ' ').trim();
}

function parseCalendarDay(text) {
  const t = normCell(text).replace(/[()]/g, '');
  if (!/^\d{1,2}$/.test(t)) return 0;
  const n = parseInt(t, 10);
  return n >= 1 && n <= 31 ? n : 0;
}

/** Form T statutory column index row: 1, 2, 3 … under employee / wage columns — not calendar days. */
function rowLooksLikeFormTColumnIndexRow(rowIndex, startCol, maxCol, getMergedAwareCellText, readRaw) {
  let total = 0;
  let sequential = 0;
  let numeric = 0;
  for (let i = 0; i < Math.min(50, maxCol - startCol); i += 1) {
    const c = startCol + i;
    const t = normCell(readRaw(rowIndex, c) || getMergedAwareCellText(rowIndex, c));
    if (!t) continue;
    total += 1;
    const n = parseCalendarDay(t);
    if (n > 0) {
      numeric += 1;
      if (n === i + 1) sequential += 1;
    }
  }
  if (total < 4) return false;
  return sequential >= Math.max(3, Math.floor(total * 0.55)) || numeric / total >= 0.85;
}

function columnHasAttendanceBanner(mainRow, endRow, col, getMergedAwareCellText) {
  for (let r = mainRow; r <= endRow; r += 1) {
    const t = String(getMergedAwareCellText(r, col) || '').toLowerCase();
    if (t.includes('attendance')) return true;
  }
  return false;
}

function pickEmployeeColumnHeader(mainRow, col, getMergedAwareCellText, readRaw) {
  const t = normCell(readRaw(mainRow, col) || getMergedAwareCellText(mainRow, col));
  if (!t || /^\d+$/.test(t)) return '';
  if (/attendance/i.test(t) && t.length > 10) return '';
  return t;
}

function rowHasSampleData(jsonData, col, fromRow, maxRow) {
  const end = Math.min(maxRow, jsonData.length);
  for (let r = fromRow; r < end; r += 1) {
    if (normCell((jsonData[r] || [])[col])) return true;
  }
  return false;
}

function scoreAttendanceCalendarDayRow(rowIndex, mainRow, scanEnd, startCol, maxCol, getMergedAwareCellText, readRaw) {
  if (rowLooksLikeFormTColumnIndexRow(rowIndex, startCol, maxCol, getMergedAwareCellText, readRaw)) {
    return 0;
  }
  let hits = 0;
  for (let c = startCol; c < maxCol; c += 1) {
    if (!columnHasAttendanceBanner(mainRow, scanEnd, c, getMergedAwareCellText)) continue;
    const day = parseCalendarDay(readRaw(rowIndex, c) || getMergedAwareCellText(rowIndex, c));
    if (day > 0) hits += 1;
  }
  return hits;
}

function findAttendanceCalendarDayRow(mainRow, scanEnd, startCol, maxCol, getMergedAwareCellText, readRaw) {
  let bestDayRow = -1;
  let bestDayHits = 0;
  for (let r = mainRow + 1; r <= scanEnd; r += 1) {
    const hits = scoreAttendanceCalendarDayRow(
      r,
      mainRow,
      scanEnd,
      startCol,
      maxCol,
      getMergedAwareCellText,
      readRaw
    );
    if (hits > bestDayHits) {
      bestDayHits = hits;
      bestDayRow = r;
    }
  }
  return { bestDayRow, bestDayHits };
}

function findFormTColumnIndexRow(mainRow, scanEnd, startCol, maxCol, getMergedAwareCellText, readRaw) {
  for (let r = mainRow + 1; r <= scanEnd; r += 1) {
    if (rowLooksLikeFormTColumnIndexRow(r, startCol, maxCol, getMergedAwareCellText, readRaw)) {
      return r;
    }
  }
  return -1;
}

/** True only for ATTENDANCE_1 … ATTENDANCE_31 — never plain "1" (Form T column index). */
export function isFormTSEAttendanceDayHeader(header) {
  const m = String(header || '')
    .trim()
    .match(/^ATTENDANCE_(\d{1,2})$/i);
  if (!m) return false;
  const day = parseInt(m[1], 10);
  return day >= 1 && day <= 31;
}

export function listFormTSEAttendanceDayHeaders(headers) {
  const out = [];
  (Array.isArray(headers) ? headers : []).forEach((header) => {
    if (!isFormTSEAttendanceDayHeader(header)) return;
    const day = parseInt(String(header).match(/^ATTENDANCE_(\d{1,2})$/i)[1], 10);
    out.push({ header, day });
  });
  return out.sort((a, b) => a.day - b.day);
}

/**
 * Re-read Form T headers: employee columns keep text labels; only the ATTENDANCE band
 * gets ATTENDANCE_1 … ATTENDANCE_31 keys (never the statutory column-index row 1…9).
 */
export function rebuildFormTSETableHeadersFromSheet({
  headerRowIndex = -1,
  getMergedAwareCellText,
  getRawCellText = null,
  jsonData,
  effectiveSheetCols = 40,
  tableStartCol = 0,
  dataStartIndex = -1,
}) {
  const readRaw = typeof getRawCellText === 'function' ? getRawCellText : getMergedAwareCellText;
  if (!Array.isArray(jsonData) || jsonData.length === 0 || headerRowIndex < 0) return null;

  const mainRow = headerRowIndex;
  const scanEnd =
    dataStartIndex > mainRow ? dataStartIndex - 1 : Math.min(mainRow + 12, jsonData.length - 1);
  const maxCol = Math.max(Number(effectiveSheetCols) || 0, tableStartCol + 55);

  let startCol = Math.max(0, Number(tableStartCol) || 0);
  for (let c = 0; c < Math.min(maxCol, 14); c += 1) {
    const t = String(getMergedAwareCellText(mainRow, c) || '').toLowerCase();
    if (/sl\.?\s*no|sr\.?\s*no|serial/.test(t)) {
      startCol = c;
      break;
    }
  }

  const indexRow = findFormTColumnIndexRow(mainRow, scanEnd, startCol, maxCol, getMergedAwareCellText, readRaw);
  const { bestDayRow, bestDayHits } = findAttendanceCalendarDayRow(
    mainRow,
    scanEnd,
    startCol,
    maxCol,
    getMergedAwareCellText,
    readRaw
  );

  const dataStartRow = Math.max(mainRow, indexRow, bestDayRow) + 1;

  const headers = [];
  for (let c = startCol; c < maxCol; c += 1) {
    let header = '';
    const inAttendanceBand = columnHasAttendanceBanner(mainRow, scanEnd, c, getMergedAwareCellText);
    if (inAttendanceBand && bestDayRow >= 0 && bestDayHits >= 3) {
      const day = parseCalendarDay(readRaw(bestDayRow, c) || getMergedAwareCellText(bestDayRow, c));
      if (day > 0) header = `ATTENDANCE_${day}`;
    }
    if (!header) {
      header = pickEmployeeColumnHeader(mainRow, c, getMergedAwareCellText, readRaw);
    }
    const hasSample = rowHasSampleData(jsonData, c, dataStartRow, dataStartRow + 15);
    if (!header && !hasSample && headers.length > 0) {
      let anyMore = false;
      for (let cc = c; cc < Math.min(c + 10, maxCol); cc += 1) {
        if (rowHasSampleData(jsonData, cc, dataStartRow, dataStartRow + 10)) {
          anyMore = true;
          break;
        }
      }
      if (!anyMore) break;
    }
    headers.push(header || (headers.length ? `Column ${headers.length + 1}` : 'Sl. No.'));
  }

  while (headers.length > 0 && !normCell(headers[headers.length - 1])) headers.pop();

  const attendanceCols = listFormTSEAttendanceDayHeaders(headers).length;
  const employeeCols = headers.filter(
    (h) => normCell(h) && !isFormTSEAttendanceDayHeader(h)
  ).length;
  if (headers.length < 4 || employeeCols < 2) return null;
  if (attendanceCols < 3 && bestDayHits < 3) return null;

  return {
    headers,
    headerRowIndex: mainRow,
    dataStartIndex: dataStartRow,
    tableStartCol: startCol,
    attendanceDayRow: bestDayRow,
    attendanceColumnCount: attendanceCols,
    columnIndexRow: indexRow,
  };
}

/** When table headers are rebuilt, copy cell values by column index. */
export function remapRowsToRebuiltTableHeaders(rows, oldHeaders, newHeaders) {
  if (!Array.isArray(rows) || !Array.isArray(oldHeaders) || !Array.isArray(newHeaders)) {
    return Array.isArray(rows) ? rows : [];
  }
  const n = Math.max(oldHeaders.length, newHeaders.length);
  return rows.map((row) => {
    if (!row || typeof row !== 'object') return row;
    const out = { ...row };
    for (let i = 0; i < n; i += 1) {
      const oldH = oldHeaders[i];
      const newH = newHeaders[i];
      if (!newH || oldH === newH) continue;
      const oldVal = oldH != null ? row[oldH] : undefined;
      const isAttendanceCode =
        /^(P|A|WO|H|L|WOP|OD|SL|CL|EL)$/i.test(String(oldVal ?? '').trim());
      if (oldH != null && oldH in row && !(newH in out)) {
        if (!isFormTSEAttendanceDayHeader(newH) || isAttendanceCode) {
          out[newH] = oldVal;
        } else {
          out[newH] = '';
        }
      }
      if (oldH != null && oldH !== newH && oldH in out) {
        delete out[oldH];
      }
    }
    return out;
  });
}

export function repairFormTSETableHeadersFromWorkbook(workbook, parsed = {}) {
  if (!workbook || !parsed) return parsed;
  const sheetName = parsed.sheetName || workbook.SheetNames?.[0];
  const ws = sheetName ? workbook.Sheets[sheetName] : null;
  if (!ws) return parsed;

  const merges = ws['!merges'] || [];
  const rawCell = (r, c) => {
    if (r < 0 || c < 0) return '';
    const ref = XLSX.utils.encode_cell({ r, c });
    const cell = ws[ref];
    return cell && cell.v != null ? String(cell.v).trim() : '';
  };
  const getMergedAwareCellText = (r, c) => {
    const direct = rawCell(r, c);
    if (direct) return direct;
    for (let i = 0; i < merges.length; i += 1) {
      const m = merges[i];
      if (!m?.s || !m?.e) continue;
      if (r >= m.s.r && r <= m.e.r && c >= m.s.c && c <= m.e.c) {
        const topLeft = rawCell(m.s.r, m.s.c);
        if (topLeft) return topLeft;
      }
    }
    return '';
  };

  const jsonData = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  const effectiveSheetCols = Math.max(
    ...(jsonData || []).map((row) => (Array.isArray(row) ? row.length : 0)),
    32
  );
  const headerRowIndex =
    parsed.originalHeaderRowIndex ?? parsed.headerRowIndex ?? -1;
  const dataStartIndex = parsed.dataStartIndex ?? -1;
  const tableStartCol = parsed.tableStartCol ?? 0;

  const rebuilt = rebuildFormTSETableHeadersFromSheet({
    headerRowIndex,
    getMergedAwareCellText,
    getRawCellText: rawCell,
    jsonData,
    effectiveSheetCols,
    tableStartCol,
    dataStartIndex,
  });
  if (!rebuilt?.headers?.length) return parsed;

  const oldHeaders = Array.isArray(parsed.headers) ? parsed.headers : [];
  return {
    ...parsed,
    headers: rebuilt.headers,
    headerRowIndex: rebuilt.headerRowIndex,
    dataStartIndex: rebuilt.dataStartIndex,
    tableStartCol: rebuilt.tableStartCol,
    rows: remapRowsToRebuiltTableHeaders(parsed.rows || [], oldHeaders, rebuilt.headers),
  };
}
