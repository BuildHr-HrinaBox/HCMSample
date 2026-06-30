import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import {
  clearExcelJSTrailingTableCells,
  countExcelJSTemplateBodyRows,
  ensureExcelJSDataRowsWithBorders,
} from '../../utils/excelTableBorders';
import { yieldToMain } from '../../utils/statutoryAutofillCache';
import { writeStatutoryHeaderFieldsToExcelJsWorksheet } from '../../utils/statutorySiteCompanyHeaders';

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

function parseStatutoryColumnNumber(text) {
  const t = normCell(text).replace(/[()]/g, '');
  if (!/^\d{1,2}$/.test(t)) return 0;
  const n = parseInt(t, 10);
  return n >= 1 && n <= 99 ? n : 0;
}

/** Form T statutory column index row: 1, 2, 3 … or 11, 12, 13 … under wage columns — not calendar days. */
function rowLooksLikeFormTColumnIndexRow(rowIndex, startCol, maxCol, getMergedAwareCellText, readRaw) {
  let total = 0;
  let sequential = 0;
  let numeric = 0;
  let ascending = 0;
  let lastNum = 0;
  for (let i = 0; i < Math.min(50, maxCol - startCol); i += 1) {
    const c = startCol + i;
    const t = normCell(readRaw(rowIndex, c) || getMergedAwareCellText(rowIndex, c));
    if (!t) continue;
    total += 1;
    const n = parseStatutoryColumnNumber(t);
    if (n > 0) {
      numeric += 1;
      if (n === i + 1) sequential += 1;
      if (lastNum === 0 || n === lastNum + 1) ascending += 1;
      lastNum = n;
    }
  }
  if (total < 4) return false;
  return (
    sequential >= Math.max(3, Math.floor(total * 0.55)) ||
    ascending >= Math.max(4, Math.floor(total * 0.65)) ||
    numeric / total >= 0.85
  );
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

function isFormTWageParentBannerText(raw) {
  const t = normCell(raw).toLowerCase();
  if (!t) return false;
  if (t === 'earned' || /^earned\s+wages/.test(t) || /^wages?\s+earned/.test(t)) return true;
  if (t.includes('earned') && t.includes('allowance')) return true;
  if (t.includes('deduction')) return true;
  if (/^attendance\s+and/.test(t)) return true;
  return false;
}

/** Leaf wage label (Basic, HRA, No. of payable days) — skip merged parent banners and column numbers. */
function pickFormTWageColumnHeader(topRow, leafRow, col, getMergedAwareCellText, readRaw) {
  const endRow = Math.max(topRow, leafRow);
  for (let r = endRow; r >= topRow; r -= 1) {
    const raw = normCell(readRaw(r, col) || getMergedAwareCellText(r, col));
    if (!raw || /^\d{1,2}$/.test(raw)) continue;
    if (isFormTWageParentBannerText(raw)) continue;
    if (/attendance/i.test(raw) && raw.length > 14) continue;
    return raw;
  }
  return '';
}

function findFormTWageLeafRow(mainRow, scanEnd, startCol, maxCol, getMergedAwareCellText, readRaw, indexRow) {
  if (indexRow > mainRow) {
    const aboveIndex = indexRow - 1;
    if (aboveIndex >= mainRow) {
      let hits = 0;
      for (let c = startCol; c < maxCol; c += 1) {
        const t = normCell(readRaw(aboveIndex, c) || getMergedAwareCellText(aboveIndex, c)).toLowerCase();
        if (
          /basic|hra|payable|conveyance|medical|allowance|overtime|\bot\b|da\/vda|\bvda\b/.test(t)
        ) {
          hits += 1;
        }
      }
      if (hits >= 2) return aboveIndex;
    }
  }
  let bestRow = mainRow;
  let bestHits = 0;
  for (let r = mainRow; r <= scanEnd; r += 1) {
    if (r === indexRow) continue;
    let hits = 0;
    for (let c = startCol; c < maxCol; c += 1) {
      const t = normCell(readRaw(r, c) || getMergedAwareCellText(r, c)).toLowerCase();
      if (
        /basic|hra|payable|conveyance|medical|allowance|overtime|\bot\b|da\/vda|\bvda\b/.test(t)
      ) {
        hits += 1;
      }
    }
    if (hits > bestHits) {
      bestHits = hits;
      bestRow = r;
    }
  }
  return bestHits >= 2 ? bestRow : Math.max(mainRow, scanEnd);
}

function appendFormTStatutoryColumnSuffix(header, statCol) {
  const base = normCell(header);
  if (!base || !statCol) return base;
  if (new RegExp(`\\(\\s*${statCol}\\s*\\)`).test(base)) return base;
  return `${base} (${statCol})`;
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
  const raw = String(header || '').trim();
  const m = raw.match(/^ATTENDANCE_(\d{1,2})(?:\s*\(\s*\d{1,2}\s*\))?$/i);
  if (!m) return false;
  const day = parseInt(m[1], 10);
  return day >= 1 && day <= 31;
}

export function listFormTSEAttendanceDayHeaders(headers) {
  const out = [];
  (Array.isArray(headers) ? headers : []).forEach((header) => {
    if (!isFormTSEAttendanceDayHeader(header)) return;
    const day = parseInt(String(header).match(/^ATTENDANCE_(\d{1,2})/i)[1], 10);
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
  const wageLeafRow = findFormTWageLeafRow(
    mainRow,
    scanEnd,
    startCol,
    maxCol,
    getMergedAwareCellText,
    readRaw,
    indexRow
  );
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
    if (inAttendanceBand) {
      let day = 0;
      if (bestDayRow >= 0 && bestDayHits >= 3) {
        day = parseCalendarDay(readRaw(bestDayRow, c) || getMergedAwareCellText(bestDayRow, c));
      } else if (indexRow >= 0) {
        day = parseCalendarDay(readRaw(indexRow, c) || getMergedAwareCellText(indexRow, c));
      }
      if (day > 0) header = `ATTENDANCE_${day}`;
    }
    if (!header) {
      header = pickFormTWageColumnHeader(mainRow, wageLeafRow, c, getMergedAwareCellText, readRaw);
    }
    if (!header) {
      header = pickEmployeeColumnHeader(mainRow, c, getMergedAwareCellText, readRaw);
    }
    if (header && indexRow >= 0 && !/^ATTENDANCE_\d{1,2}$/i.test(header)) {
      const statCol = parseStatutoryColumnNumber(
        readRaw(indexRow, c) || getMergedAwareCellText(indexRow, c)
      );
      if (statCol > 0) header = appendFormTStatutoryColumnSuffix(header, statCol);
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
  const hasWageColumnMarkers = headers.some((h) =>
    /\(\s*1[1-9]\s*\)/.test(String(h || ''))
  );
  if (headers.length < 4 || employeeCols < 2) return null;
  if (attendanceCols < 3 && bestDayHits < 3 && !hasWageColumnMarkers) return null;

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

/** Keep ATTENDANCE_1…31 keys when a re-read would drop them (autofill must match modal headers). */
export function preserveFormTSEAttendanceHeadersIfNeeded(oldHeaders, newHeaders) {
  const oldList = Array.isArray(oldHeaders) ? oldHeaders : [];
  const newList = Array.isArray(newHeaders) ? newHeaders : [];
  if (!oldList.length || !newList.length) return newList;
  const oldAtt = listFormTSEAttendanceDayHeaders(oldList).length;
  const newAtt = listFormTSEAttendanceDayHeaders(newList).length;
  if (oldAtt < 3 || newAtt >= oldAtt) return newList;
  const n = Math.max(oldList.length, newList.length);
  const out = newList.slice();
  for (let i = 0; i < n; i += 1) {
    const oh = oldList[i];
    const nh = out[i];
    if (!isFormTSEAttendanceDayHeader(nh) && isFormTSEAttendanceDayHeader(oh)) {
      out[i] = oh;
    }
  }
  return out;
}

export function normalizeFormTSETableHeaders(headers) {
  const list = Array.isArray(headers) ? headers : [];
  return list.map((h) => {
    const raw = String(h || '').trim();
    const m = raw.match(/^(ATTENDANCE_\d{1,2})\s*\(\s*\d{1,2}\s*\)$/i);
    return m ? m[1] : h;
  });
}

export function formTSETableHeadersHaveAttendanceBand(headers) {
  return listFormTSEAttendanceDayHeaders(headers).length >= 3;
}

/** Align live modal row keys to export headers (column index) for Excel download. */
export function prepareFormTSEExportRows(liveRows, exportHeaders, sourceHeaders) {
  const hdrs = Array.isArray(exportHeaders) ? exportHeaders : [];
  let srcHdrs = Array.isArray(sourceHeaders) && sourceHeaders.length > 0 ? sourceHeaders : hdrs;
  const rows = Array.isArray(liveRows) ? liveRows : [];
  if (!hdrs.length) return rows.map((row) => ({ ...row }));

  if (rows.length > 0 && rows[0] && typeof rows[0] === 'object' && !Array.isArray(rows[0])) {
    const rowKeys = Object.keys(rows[0]).filter((k) => !String(k).startsWith('__'));
    const keysMatchExport =
      rowKeys.length === hdrs.length && rowKeys.every((k, i) => k === hdrs[i]);
    if (!keysMatchExport && rowKeys.length >= Math.min(hdrs.length, 4)) {
      if (rowKeys.length === hdrs.length) {
        srcHdrs = rowKeys;
      } else if (srcHdrs.length !== hdrs.length) {
        srcHdrs = rowKeys;
      }
    }
  }

  const alignedRows =
    srcHdrs.length === hdrs.length && srcHdrs.some((h, i) => h !== hdrs[i])
      ? remapRowsToRebuiltTableHeaders(rows, srcHdrs, hdrs)
      : rows;

  return alignedRows.map((row, rowIndex) => {
    const out = {};
    hdrs.forEach((header, colIndex) => {
      let value = '';
      if (row && typeof row === 'object' && !Array.isArray(row)) {
        if (Object.prototype.hasOwnProperty.call(row, header)) {
          value = row[header];
        } else {
          const srcKey = srcHdrs[colIndex];
          if (srcKey && Object.prototype.hasOwnProperty.call(row, srcKey)) {
            value = row[srcKey];
          }
        }
      } else if (Array.isArray(row) && colIndex < row.length) {
        value = row[colIndex];
      }
      if (
        (value == null || String(value).trim() === '') &&
        /s\.?\s*no|serial|sl\.?\s*no/i.test(String(header || ''))
      ) {
        value = rowIndex + 1;
      }
      out[header] = value != null ? value : '';
    });
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
  const mergedHeaders = preserveFormTSEAttendanceHeadersIfNeeded(
    oldHeaders,
    normalizeFormTSETableHeaders(rebuilt.headers)
  );
  return {
    ...parsed,
    headers: mergedHeaders,
    headerRowIndex: rebuilt.headerRowIndex,
    dataStartIndex: rebuilt.dataStartIndex,
    tableStartCol: rebuilt.tableStartCol,
    rows: remapRowsToRebuiltTableHeaders(parsed.rows || [], oldHeaders, mergedHeaders),
  };
}

/** Write Form T grid rows into the template at exact column positions (ExcelJS, 1-based). */
export async function buildFormTSEWorkbookWithTemplateStyles({
  templateArrayBuffer,
  mappedData,
  headersToUse,
  parsedHeaderRowIndex = -1,
  parsedDataStartIndex = -1,
  parsedTableStartCol = 0,
  parsedFormHeader,
  headerFormData = {},
  formFileName,
  currentItem = null,
  sheetNameHint = '',
}) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(templateArrayBuffer);
  const worksheet =
    (sheetNameHint && workbook.getWorksheet(sheetNameHint)) || workbook.worksheets[0];
  if (!worksheet) throw new Error('Template worksheet not found.');

  const excelCellValueToString = (val) => {
    if (val == null) return '';
    if (typeof val === 'string' || typeof val === 'number' || typeof val === 'boolean') {
      return String(val);
    }
    if (val instanceof Date) return val.toISOString();
    if (typeof val === 'object') {
      if (Array.isArray(val.richText)) return val.richText.map((rt) => rt?.text || '').join('');
      if (val.text != null) return String(val.text);
      if (val.result != null) return String(val.result);
    }
    return '';
  };

  const mergeTopLeftCache = new Map();
  const getMergeTopLeft = (r, c) => {
    const key = `${r}:${c}`;
    if (mergeTopLeftCache.has(key)) return mergeTopLeftCache.get(key);
    let topLeft = { r, c };
    const merges = worksheet.model?.merges;
    if (Array.isArray(merges)) {
      for (let mi = 0; mi < merges.length; mi += 1) {
        const parts = String(merges[mi] || '').split(':');
        if (parts.length !== 2) continue;
        const tl = worksheet.getCell(parts[0]);
        const br = worksheet.getCell(parts[1]);
        if (!tl || !br) continue;
        if (r >= tl.row && r <= br.row && c >= tl.col && c <= br.col) {
          topLeft = { r: tl.row, c: tl.col };
          break;
        }
      }
    }
    mergeTopLeftCache.set(key, topLeft);
    return topLeft;
  };
  const getMergedAwareCellText = (r, c) => {
    const tl = getMergeTopLeft(r, c);
    return excelCellValueToString(worksheet.getCell(tl.r, tl.c)?.value).trim();
  };

  const maxScanRows = Math.max(
    (Number(parsedDataStartIndex) >= 0 ? Number(parsedDataStartIndex) + 30 : 0),
    (Number(parsedHeaderRowIndex) >= 0 ? Number(parsedHeaderRowIndex) + 18 : 0),
    50,
    worksheet.rowCount + 5
  );
  const maxScanCols = Math.max(worksheet.columnCount + 5, 55);
  const jsonData = [];
  for (let r = 1; r <= maxScanRows; r += 1) {
    const row = [];
    for (let c = 1; c <= maxScanCols; c += 1) {
      row.push(getMergedAwareCellText(r, c));
    }
    jsonData.push(row);
  }
  const getMergedAwareCellText0 = (r0, c0) => getMergedAwareCellText(r0 + 1, c0 + 1);
  const rawCell0 = (r0, c0) => getMergedAwareCellText0(r0, c0);

  const headerRowIndex =
    Number(parsedHeaderRowIndex) >= 0
      ? Number(parsedHeaderRowIndex)
      : rebuildFormTSETableHeadersFromSheet({
          headerRowIndex: -1,
          getMergedAwareCellText: getMergedAwareCellText0,
          getRawCellText: rawCell0,
          jsonData,
          effectiveSheetCols: maxScanCols,
          tableStartCol: parsedTableStartCol,
          dataStartIndex: parsedDataStartIndex,
        })?.headerRowIndex ?? -1;

  const rebuilt = rebuildFormTSETableHeadersFromSheet({
    headerRowIndex,
    getMergedAwareCellText: getMergedAwareCellText0,
    getRawCellText: rawCell0,
    jsonData,
    effectiveSheetCols: maxScanCols,
    tableStartCol: parsedTableStartCol,
    dataStartIndex: parsedDataStartIndex,
  });

  const inputHeaders = Array.isArray(headersToUse) ? headersToUse.filter((h) => String(h || '').trim()) : [];
  const effectiveHeaders =
    rebuilt?.headers?.length >= 4
      ? preserveFormTSEAttendanceHeadersIfNeeded(
          inputHeaders,
          normalizeFormTSETableHeaders(rebuilt.headers)
        )
      : inputHeaders;
  if (effectiveHeaders.length < 4) {
    throw new Error('Could not locate Form T table columns.');
  }

  const tableStartCol =
    rebuilt?.tableStartCol != null && rebuilt.tableStartCol >= 0
      ? rebuilt.tableStartCol
      : Math.max(0, Number(parsedTableStartCol) || 0);
  const fieldCols = effectiveHeaders.map((_, idx) => tableStartCol + idx + 1);
  const startRow =
    rebuilt?.dataStartIndex != null && rebuilt.dataStartIndex >= 0
      ? rebuilt.dataStartIndex + 1
      : Number(parsedDataStartIndex) >= 0
        ? Number(parsedDataStartIndex) + 1
        : headerRowIndex >= 0
          ? headerRowIndex + 3
          : 14;

  writeStatutoryHeaderFieldsToExcelJsWorksheet(worksheet, {
    headerFormData: headerFormData && typeof headerFormData === 'object' ? headerFormData : {},
    parsedFormHeader,
    headerRowEnd: Math.max(1, startRow - 1),
    maxScanCols: 80,
  });

  const normalize = (txt) =>
    String(txt || '')
      .replace(/\r?\n/g, ' ')
      .replace(/\s+/g, ' ')
      .trim()
      .toLowerCase()
      .replace(/[^a-z0-9 ]/g, '');

  const getRowValueForHeader = (row, header, headerIndex, allHeaders) => {
    if (!row || typeof row !== 'object' || Array.isArray(row)) {
      return Array.isArray(row) ? row[headerIndex] : '';
    }
    if (Object.prototype.hasOwnProperty.call(row, header)) return row[header];
    const target = normalize(header);
    if (!target) return '';
    const rowKeys = Object.keys(row).filter((k) => !String(k || '').startsWith('__'));
    const exact = rowKeys.filter((k) => normalize(k) === target);
    if (exact.length === 1) return row[exact[0]];
    if (exact.length > 1) {
      const occur =
        allHeaders.slice(0, headerIndex + 1).filter((h) => normalize(h) === target).length - 1;
      return row[exact[Math.min(Math.max(occur, 0), exact.length - 1)]];
    }
    const fuzzy = rowKeys.find((k) => {
      const nk = normalize(k);
      return nk && (nk.includes(target) || target.includes(nk));
    });
    if (fuzzy) return row[fuzzy];
    if (headerIndex < rowKeys.length) return row[rowKeys[headerIndex]];
    return '';
  };

  const rowLooksMeaningful = (row) => {
    if (Array.isArray(row)) {
      return row.some((v) => v != null && String(v).trim() !== '');
    }
    if (row && typeof row === 'object') {
      return Object.values(row).some((v) => v != null && String(v).trim() !== '');
    }
    return false;
  };

  const sourceHeaders =
    inputHeaders.length === effectiveHeaders.length ? inputHeaders : effectiveHeaders;
  const sourcePrimary = prepareFormTSEExportRows(
    Array.isArray(mappedData) ? mappedData : [],
    effectiveHeaders,
    sourceHeaders
  ).filter((row) => rowLooksMeaningful(row));

  const tableColMin = fieldCols.length > 0 ? Math.min(...fieldCols) : 1;
  const tableColMax = fieldCols.length > 0 ? Math.max(...fieldCols) : effectiveHeaders.length;
  const templateBodyRows = countExcelJSTemplateBodyRows(worksheet, startRow, tableColMin, tableColMax);
  const bodyRowsToPaint = Math.max(sourcePrimary.length, templateBodyRows);

  clearExcelJSTrailingTableCells(worksheet, {
    dataStartRow: startRow,
    dataRowCount: bodyRowsToPaint,
    afterCol: tableColMax,
    throughCol: tableColMax + 30,
  });

  for (let i = 0; i < sourcePrimary.length; i += 1) {
    if (i > 0 && i % 15 === 0) {
      await yieldToMain();
    }
    const row = sourcePrimary[i];
    const excelRow = startRow + i;
    for (let j = 0; j < effectiveHeaders.length; j += 1) {
      const header = effectiveHeaders[j];
      let value = getRowValueForHeader(row, header, j, effectiveHeaders);
      if (
        (value == null || String(value).trim() === '') &&
        /s\.?\s*no|serial|sl\.?\s*no/i.test(String(header || ''))
      ) {
        value = i + 1;
      }
      const targetCol = fieldCols[j];
      if (!targetCol || targetCol < 1) continue;
      const tl = getMergeTopLeft(excelRow, targetCol);
      const cell = worksheet.getCell(tl.r, tl.c);
      if (value == null || value === '') {
        cell.value = '';
        continue;
      }
      if (
        typeof value === 'number' ||
        (typeof value === 'string' && /^-?\d+(\.\d+)?$/.test(String(value).trim()))
      ) {
        cell.value = Number(value);
      } else {
        cell.value = String(value);
      }
      cell.font = { ...(cell.font || {}), bold: false };
    }
  }

  if (sourcePrimary.length > 0 && fieldCols.length > 0) {
    ensureExcelJSDataRowsWithBorders(worksheet, {
      dataStartRow: startRow,
      dataRowCount: sourcePrimary.length,
      colFrom: tableColMin,
      colTo: tableColMax,
      templateRow: startRow,
      templateBodyRows: 1,
    });
    clearExcelJSTrailingTableCells(worksheet, {
      dataStartRow: startRow,
      dataRowCount: bodyRowsToPaint,
      afterCol: tableColMax,
      throughCol: tableColMax + 30,
    });
  }

  await yieldToMain();
  const out = await workbook.xlsx.writeBuffer();
  const outName =
    formFileName ||
    currentItem?.formName?.replace(/[^a-zA-Z0-9]/g, '_') ||
    parsedFormHeader?.title?.replace(/[^a-zA-Z0-9]/g, '_') ||
    `Form_T_${Date.now()}.xlsx`;
  const blob = new Blob([out], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  return { blob, fileName: outName };
}
