/** Tamil Nadu Factories Form 26 — Register of Accidents (Rule 104). */

import ExcelJS from 'exceljs';
import { writeStatutoryHeaderFieldsToExcelJsWorksheet } from '../../utils/statutorySiteCompanyHeaders';
import { ensureExcelJSDataRowsWithBorders } from '../../utils/excelTableBorders';

export const FORM26_ACCIDENTS_NIL_OF_MONTH_TEXT = 'Nill of the month';

export function isForm26AccidentsNilMonthValue(v) {
  const text = String(v ?? '').trim();
  return /^nill?\s+of\s+the\s+month/i.test(text) || /^nil\s+for\s+the\s+month/i.test(text);
}

export function findForm26NilPrimaryHeader(headers) {
  for (const h of headers || []) {
    const s = String(h || '')
      .toLowerCase()
      .replace(/\s+/g, ' ')
      .trim();
    const bare = s.replace(/^\(?\d+\)?\s*/, '').trim();
    if (/\(1\)|^1\)/.test(s) && /sl\.?\s*no|serial|calendar\s+year|running/.test(bare)) return h;
    if (/^sl\.?\s*no/.test(bare) || /^serial(\s+number)?/.test(bare)) return h;
    if (/calendar\s+year/.test(bare) && /running|sl\.?\s*no|serial/.test(bare)) return h;
  }
  for (const h of headers || []) {
    const s = String(h || '')
      .toLowerCase()
      .replace(/\s+/g, ' ')
      .trim();
    const bare = s.replace(/^\(?\d+\)?\s*/, '').trim();
    if (/\(7\)|^7\)/.test(s)) return h;
    if (/date\s+of\s+despatch|date\s+of\s+dispatch/.test(bare) && /form\s*18/.test(bare)) return h;
  }
  return (headers && headers[0]) || null;
}

/** Merge "Nill of the month" across data columns up to (but not including) Remarks. */
export function resolveForm26NilSpanInfo(headers) {
  const primary = findForm26NilPrimaryHeader(headers);
  if (!primary) return null;
  const startIdx = (headers || []).indexOf(primary);
  if (startIdx < 0) return null;
  let span = 1;
  for (let i = startIdx + 1; i < headers.length; i += 1) {
    const h = headers[i];
    const s = String(h || '')
      .toLowerCase()
      .replace(/\s+/g, ' ')
      .trim();
    const bare = s.replace(/^\(?\d+\)?\s*/, '').trim();
    if (
      /\(14\)|^14\)/.test(s) ||
      (/remarks/.test(bare) && (/initials/.test(bare) || /manager/.test(bare) || /manage/.test(bare)))
    ) {
      break;
    }
    span += 1;
  }
  // Single-cell nil is hard to see on a 14-col register — span the full table band.
  if (span <= 1 && headers.length - startIdx > 1) {
    span = headers.length - startIdx;
  }
  return { startIdx, span, primaryHeader: primary };
}

export function isForm26AccidentDataHeader(h) {
  const s = String(h || '')
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .trim();
  const bare = s.replace(/^\(?\d+\)?\s*/, '').trim();
  if (/\(2\)|^2\)/.test(s) && /accident/.test(bare)) return true;
  if (/\(3\)|^3\)/.test(s) && /injured|designation|person/.test(bare)) return true;
  if (/\(4\)|^4\)/.test(s) && /place/.test(bare) && /accident|factory|branch|dept|machine/.test(bare)) {
    return true;
  }
  if (/\(5\)|^5\)/.test(s) && /description/.test(bare)) return true;
  if (/\(6\)|^6\)/.test(s) && /injury|nature|extent|location/.test(bare)) return true;
  if (/date.*hour.*accident/.test(bare)) return true;
  if (/name.*designation.*injured/.test(bare)) return true;
  if (/exact\s+place.*accident/.test(bare)) return true;
  if (/full\s+description.*accident/.test(bare)) return true;
  if (/nature.*extent.*injury|nature.*location.*injury/.test(bare)) return true;
  return false;
}

export function form26RowHasAccidentEntry(row, headers) {
  if (!row || typeof row !== 'object') return false;
  return (headers || []).some((h) => {
    if (!isForm26AccidentDataHeader(h)) return false;
    const v = String(row[h] ?? '').trim();
    return v && !isForm26AccidentsNilMonthValue(v);
  });
}

export function buildForm26NilTableRows(headers, nilLine = FORM26_ACCIDENTS_NIL_OF_MONTH_TEXT) {
  const nilPrimary = findForm26NilPrimaryHeader(headers);
  const row = {};
  (headers || []).forEach((h) => {
    row[h] = '';
  });
  if (nilPrimary && nilLine) row[nilPrimary] = nilLine;
  return [row];
}

export function applyForm26NilTableRows(headers, existingRows) {
  const rows = Array.isArray(existingRows) ? existingRows : [];
  if (rows.some((r) => form26RowHasAccidentEntry(r, headers))) return rows;
  return buildForm26NilTableRows(headers, FORM26_ACCIDENTS_NIL_OF_MONTH_TEXT);
}

/** True when headers look like Factories Form 26 Accidents (not Form 26-A). */
export function tableHeadersLookLikeForm26Accidents(tableHeaders) {
  if (!Array.isArray(tableHeaders) || tableHeaders.length === 0) return false;
  const joined = tableHeaders.map((h) => String(h || '').toLowerCase()).join('\n');
  if (/dangerous\s+occurrence|date.*time.*occurrence/.test(joined)) return false;
  if (
    /date\s*[&and]*\s*hour\s+of\s+accident|name.*designation.*(?:person\s+)?injured|person\s+injured/.test(
      joined
    )
  ) {
    return true;
  }
  if (/nature.*extent.*injury|nature.*location.*injury/.test(joined) && /remarks.*initials/.test(joined)) {
    return true;
  }
  return false;
}

function excelCellValueToString(val) {
  if (val == null) return '';
  if (typeof val === 'string' || typeof val === 'number' || typeof val === 'boolean') return String(val);
  if (val instanceof Date) return val.toISOString();
  if (typeof val === 'object') {
    if (Array.isArray(val.richText)) return val.richText.map((rt) => rt?.text || '').join('');
    if (val.text != null) return String(val.text);
    if (val.result != null) return String(val.result);
  }
  return '';
}

function sanitizeForm26ExportCell(value, header) {
  const text = String(value ?? '').trim();
  if (isForm26AccidentsNilMonthValue(text)) return text;
  if (isForm26AccidentDataHeader(header)) {
    if (text && !/^enter\s+/i.test(text) && !isForm26AccidentsNilMonthValue(text)) return text;
    return '';
  }
  if (value == null || text === '' || /^enter\s+/i.test(text)) return '';
  return text;
}

/** Ensure header wrap + readable widths so 14 Form 26 columns fit the register layout. */
export function applyForm26AccidentsColumnLayout(worksheet, orderedCols, headerRow) {
  if (!worksheet || !Array.isArray(orderedCols) || orderedCols.length === 0) return;
  const widthByOffset = [
    10, // (1) Sl. No.
    14, // (2) Date & Hour
    18, // (3) Name & Designation
    16, // (4) Exact Place
    18, // (5) Full Description
    16, // (6) Nature / Injury
    12, // (7)
    12, // (8)
    12, // (9)
    12, // (10)
    10, // (11)
    10, // (12)
    14, // (13)
    16, // (14) Remarks
  ];
  for (let i = 0; i < orderedCols.length; i += 1) {
    const c = orderedCols[i];
    const col = worksheet.getColumn(c);
    const current = Number(col.width) || 0;
    const target = widthByOffset[i] != null ? widthByOffset[i] : 12;
    if (current < target * 0.85) col.width = target;
    if (headerRow > 0) {
      const cell = worksheet.getCell(headerRow, c);
      cell.alignment = {
        ...(cell.alignment || {}),
        wrapText: true,
        vertical: 'middle',
        horizontal: 'center',
      };
    }
  }
  if (headerRow > 0) {
    const row = worksheet.getRow(headerRow);
    const h = Number(row.height) || 0;
    if (h < 45) row.height = 60;
  }
}

/** Merge title / subtitle across the table band when SheetJS left them in A1/A2 only. */
export function ensureForm26AccidentsTitleBand(worksheet, orderedCols) {
  if (!worksheet || !Array.isArray(orderedCols) || orderedCols.length < 2) return;
  const colFrom = Math.min(...orderedCols);
  const colTo = Math.max(...orderedCols);
  const normalize = (txt) =>
    String(txt || '')
      .replace(/\r?\n/g, ' ')
      .replace(/\s+/g, ' ')
      .trim()
      .toLowerCase();

  for (let r = 1; r <= 3; r += 1) {
    const text = excelCellValueToString(worksheet.getCell(r, colFrom)?.value).trim();
    if (!text) continue;
    const n = normalize(text);
    const looksLikeTitle =
      /form\s*26|factories\s+rules|register\s+of\s+accident|prescribed\s+under\s+rule/.test(n);
    if (!looksLikeTitle) continue;
    // Skip if another column already holds distinct title text (template already laid out).
    let otherFilled = false;
    for (let c = colFrom + 1; c <= colTo; c += 1) {
      const t = excelCellValueToString(worksheet.getCell(r, c)?.value).trim();
      if (t && normalize(t) !== n) {
        otherFilled = true;
        break;
      }
    }
    if (otherFilled) continue;
    const range = `${worksheet.getCell(r, colFrom).address}:${worksheet.getCell(r, colTo).address}`;
    try {
      worksheet.unMergeCells(range);
    } catch (_) {
      /* ignore */
    }
    try {
      worksheet.mergeCells(r, colFrom, r, colTo);
    } catch (_) {
      /* ignore */
    }
    const cell = worksheet.getCell(r, colFrom);
    cell.value = text;
    cell.alignment = {
      ...(cell.alignment || {}),
      horizontal: 'center',
      vertical: 'middle',
      wrapText: true,
    };
  }
}

/**
 * Form 26 Accidents Excel export — preserve template columns, widths, borders, and nil-month merge.
 * Avoids SheetJS buildDraftWorkbook which collapses column widths so headers no longer fit.
 */
export async function buildForm26WorkbookWithTemplateStyles({
  templateArrayBuffer,
  mappedData,
  mappedRowMatrix,
  headersToUse,
  parsedHeaderRowIndex,
  parsedDataStartIndex,
  parsedFormHeader,
  headerFormData,
  formFileName,
}) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(templateArrayBuffer);
  const worksheet =
    workbook.worksheets.find((ws) => {
      const name = String(ws?.name || '');
      if (/26[\s-]*a|dangerous/i.test(name)) return false;
      return /form\s*26|accident/i.test(name);
    }) || workbook.worksheets[0];
  if (!worksheet) throw new Error('Template worksheet not found.');

  const normalize = (txt) =>
    String(txt || '')
      .replace(/\r?\n/g, ' ')
      .replace(/\s+/g, ' ')
      .trim()
      .toLowerCase();

  const maxScanRows = Math.max(120, worksheet.rowCount + 10);
  const maxScanCols = 24;
  let headerRow =
    parsedHeaderRowIndex != null && parsedHeaderRowIndex >= 0 ? parsedHeaderRowIndex + 1 : -1;
  if (headerRow < 1) {
    for (let r = 1; r <= Math.min(20, maxScanRows); r += 1) {
      const parts = [];
      for (let c = 1; c <= Math.min(16, maxScanCols); c += 1) {
        const t = normalize(excelCellValueToString(worksheet.getCell(r, c)?.value));
        if (t) parts.push(t);
      }
      const joined = parts.join(' ');
      if (
        (/sl\.?\s*no|serial/.test(joined) && /accident|injured|injury/.test(joined)) ||
        (/date.*hour.*accident/.test(joined) && /exact\s+place|nature/.test(joined))
      ) {
        headerRow = r;
        break;
      }
    }
  }
  if (headerRow < 1) throw new Error('Could not locate Form 26 header row from parser.');

  let startCol = 1;
  for (let c = 1; c <= maxScanCols; c += 1) {
    const t = normalize(excelCellValueToString(worksheet.getCell(headerRow, c)?.value));
    if (t) {
      startCol = c;
      break;
    }
  }

  let markerRow = parsedDataStartIndex != null && parsedDataStartIndex > 0 ? parsedDataStartIndex : -1;
  let markerCols = [];
  for (
    let r = Math.max(headerRow, markerRow > 0 ? markerRow - 1 : headerRow);
    r <= Math.min(headerRow + 6, maxScanRows);
    r += 1
  ) {
    const cols = [];
    for (let c = startCol; c <= maxScanCols; c += 1) {
      const t = normalize(excelCellValueToString(worksheet.getCell(r, c)?.value));
      if (/^\d+$/.test(t) || /^\(\d+\)$/.test(t)) cols.push(c);
    }
    if (cols.length > markerCols.length) {
      markerCols = cols;
      markerRow = r;
    }
  }

  const dataStartRow =
    parsedDataStartIndex != null && parsedDataStartIndex >= 0
      ? parsedDataStartIndex + 1
      : markerCols.length >= 3 && markerRow > 0
        ? markerRow + 1
        : headerRow + 1;

  const hdrs = Array.isArray(headersToUse) ? headersToUse : [];
  // Form 26 Accidents headers are "(1)…(14)…" in one row — map sequentially from the first header col.
  // Only use a separate numeric marker row when it is below the label row and covers the full band.
  const sequentialCols = Array.from(
    { length: Math.max(hdrs.length || 14, 14) },
    (_, i) => startCol + i
  );
  const markerIsSeparateBand =
    markerCols.length >= Math.max(8, Math.min(hdrs.length || 14, 14)) &&
    markerRow > headerRow &&
    markerCols[0] === startCol;
  const cols = markerIsSeparateBand ? markerCols : sequentialCols;

  const nilSpanInfo = resolveForm26NilSpanInfo(hdrs);

  const data = headerFormData && typeof headerFormData === 'object' ? headerFormData : {};
  writeStatutoryHeaderFieldsToExcelJsWorksheet(worksheet, {
    headerFormData: data,
    parsedFormHeader,
    headerRowEnd: headerRow,
    maxScanCols: 24,
  });

  ensureForm26AccidentsTitleBand(worksheet, cols);
  applyForm26AccidentsColumnLayout(worksheet, cols, headerRow);

  const sourceRows = (
    Array.isArray(mappedData) && mappedData.length > 0
      ? mappedData
      : Array.isArray(mappedRowMatrix)
        ? mappedRowMatrix
        : []
  ).filter((row) => {
    if (Array.isArray(row)) return row.some((v) => String(v ?? '').trim());
    if (row && typeof row === 'object') return Object.values(row).some((v) => String(v ?? '').trim());
    return false;
  });

  const isNumericMarkerTemplateRow = (r) => {
    if (!markerCols || markerCols.length < 3) return false;
    if (r === headerRow) return false;
    let numericHits = 0;
    for (let i = 0; i < markerCols.length; i += 1) {
      const v = normalize(excelCellValueToString(worksheet.getCell(r, markerCols[i])?.value));
      if (/^\d+$/.test(v) || /^\(\d+\)$/.test(v)) numericHits += 1;
    }
    return numericHits >= Math.max(3, Math.floor(markerCols.length * 0.6));
  };

  const clearFromRow = Math.max(1, dataStartRow);
  const clearToRow = Math.min(
    maxScanRows,
    Math.max(dataStartRow + Math.max(sourceRows.length, 1) + 20, clearFromRow + 20)
  );
  for (let r = clearFromRow; r <= clearToRow; r += 1) {
    if (isNumericMarkerTemplateRow(r)) continue;
    // Unmerge leftover template merges in the data band so nil merge can apply cleanly.
    for (let j = 0; j < cols.length; j += 1) {
      const cell = worksheet.getCell(r, cols[j]);
      try {
        if (cell?.isMerged) {
          const master = cell.master || cell;
          const addr = master.address;
          // Best-effort: clear value; merge removal handled per nil write.
          void addr;
        }
      } catch (_) {
        /* ignore */
      }
      worksheet.getCell(r, cols[j]).value = '';
    }
  }

  for (let i = 0; i < sourceRows.length; i += 1) {
    const rowObj = sourceRows[i];
    const targetRowNum = dataStartRow + i;
    const rowValues = hdrs.map((h) => {
      if (Array.isArray(rowObj)) {
        const idx = hdrs.indexOf(h);
        return idx >= 0 ? rowObj[idx] : '';
      }
      return rowObj?.[h];
    });
    const nilPrimary = nilSpanInfo?.primaryHeader;
    const nilValue = nilPrimary
      ? sanitizeForm26ExportCell(
          Array.isArray(rowObj) ? rowValues[nilSpanInfo.startIdx] : rowObj?.[nilPrimary],
          nilPrimary
        )
      : '';

    if (nilSpanInfo && isForm26AccidentsNilMonthValue(nilValue)) {
      const startColIdx = cols[nilSpanInfo.startIdx];
      const endColIdx = cols[Math.min(nilSpanInfo.startIdx + nilSpanInfo.span - 1, cols.length - 1)];
      if (startColIdx != null && endColIdx != null && endColIdx >= startColIdx) {
        for (let c = startColIdx; c <= endColIdx; c += 1) {
          worksheet.getCell(targetRowNum, c).value = '';
        }
        try {
          worksheet.unMergeCells(targetRowNum, startColIdx, targetRowNum, endColIdx);
        } catch (_) {
          /* ignore */
        }
        try {
          worksheet.mergeCells(targetRowNum, startColIdx, targetRowNum, endColIdx);
        } catch (_) {
          /* ignore */
        }
        const nilCell = worksheet.getCell(targetRowNum, startColIdx);
        nilCell.value = String(nilValue);
        nilCell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        continue;
      }
    }

    for (let j = 0; j < cols.length; j += 1) {
      const header = hdrs[j] || '';
      const rawValue = Array.isArray(rowObj) ? rowValues[j] : rowObj?.[header];
      const value = sanitizeForm26ExportCell(rawValue, header);
      if (value == null || value === '') continue;
      const cell = worksheet.getCell(targetRowNum, cols[j]);
      if (/sl\.?\s*no|serial|\(1\)/i.test(header) && /^-?\d+$/.test(String(value))) {
        cell.value = Number(value);
      } else {
        cell.value = String(value);
        cell.alignment = { ...(cell.alignment || {}), wrapText: true, vertical: 'top' };
      }
    }
  }

  if (sourceRows.length > 0) {
    const tableColMin = cols.length > 0 ? Math.min(...cols) : startCol;
    const tableColMax = cols.length > 0 ? Math.max(...cols) : startCol + cols.length - 1;
    ensureExcelJSDataRowsWithBorders(worksheet, {
      dataStartRow,
      dataRowCount: sourceRows.length,
      colFrom: tableColMin,
      colTo: tableColMax,
      templateRow: dataStartRow,
      templateBodyRows: 1,
    });
  }

  // Re-apply layout after data write (borders helpers can leave widths untouched).
  applyForm26AccidentsColumnLayout(worksheet, cols, headerRow);

  const out = await workbook.xlsx.writeBuffer();
  const fileName =
    formFileName ||
    parsedFormHeader?.title?.replace(/[^a-zA-Z0-9]/g, '_') ||
    `Form_26_Register_of_Accidents_${Date.now()}.xlsx`;
  const blob = new Blob([out], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  return { blob, fileName };
}
