import ExcelJS from 'exceljs';
import { writeStatutoryHeaderFieldsToExcelJsWorksheet } from '../../utils/statutorySiteCompanyHeaders';
import { ensureExcelJSDataRowsWithBorders } from '../../utils/excelTableBorders';

export function isForm26NilMonthLineValue(v) {
  return /^nil\s+for\s+the\s+month/i.test(String(v ?? '').trim());
}

export function findForm26ACalendarYearHeader(headers) {
  for (const h of headers || []) {
    const s = String(h || '').toLowerCase().replace(/\s+/g, ' ').trim();
    const bare = s.replace(/^\(?\d+\)?\s*/, '').trim();
    if (/\(1\)|^1\)/.test(s) && /calendar\s+year/.test(bare)) return h;
    if (/^calendar\s+year/.test(bare)) return h;
  }
  return null;
}

export function findForm26ANilPrimaryHeader(headers) {
  for (const h of headers || []) {
    const s = String(h || '').toLowerCase().replace(/\s+/g, ' ').trim();
    const bare = s.replace(/^\(?\d+\)?\s*/, '').trim();
    if (/\(3\)|^3\)/.test(s) && /date|time|occurrence/.test(bare)) return h;
    if (/date.*time.*occurrence/.test(bare)) return h;
  }
  for (const h of headers || []) {
    const s = String(h || '').toLowerCase().replace(/\s+/g, ' ').trim();
    const bare = s.replace(/^\(?\d+\)?\s*/, '').trim();
    if (/\(4\)|^4\)/.test(s) && /date.*report/.test(bare)) return h;
  }
  return null;
}

export function resolveForm26ANilSpanInfo(headers) {
  const primary = findForm26ANilPrimaryHeader(headers);
  if (!primary) return null;
  const startIdx = (headers || []).indexOf(primary);
  if (startIdx < 0) return null;
  let span = 1;
  for (let i = startIdx + 1; i < headers.length; i += 1) {
    const h = headers[i];
    const s = String(h || '').toLowerCase().replace(/\s+/g, ' ').trim();
    const bare = s.replace(/^\(?\d+\)?\s*/, '').trim();
    if (/\(8\)|^8\)/.test(s) || (/remarks/.test(bare) && (/initials/.test(bare) || /manager/.test(bare)))) break;
    span += 1;
  }
  return { startIdx, span, primaryHeader: primary };
}

function isForm26AOccurrenceDataHeader(h) {
  const s = String(h || '').toLowerCase().replace(/\s+/g, ' ').trim();
  const bare = s.replace(/^\(?\d+\)?\s*/, '').trim();
  if (/\(2\)|^2\)/.test(s) && /serial|running/.test(bare)) return true;
  if (/\(3\)|^3\)/.test(s) && /date|time|occurrence/.test(bare)) return true;
  if (/\(4\)|^4\)/.test(s) && /date.*report/.test(bare)) return true;
  if (/\(5\)|^5\)/.test(s) && /place/.test(bare)) return true;
  if (/\(6\)|^6\)/.test(s) && /description/.test(bare)) return true;
  if (/\(7\)|^7\)/.test(s) && (/action\s+taken|details/.test(bare))) return true;
  return /date.*time.*occurrence|date.*report|exact\s+place|full\s+description|action\s+taken|running.*serial/.test(
    bare
  );
}

function sanitizeForm26AExportCell(value, header, calendarHeader) {
  const text = String(value ?? '').trim();
  if (isForm26NilMonthLineValue(text)) return text;
  if (calendarHeader && header === calendarHeader && /^\d{4}$/.test(text)) return text;
  if (isForm26AOccurrenceDataHeader(header)) {
    if (text && !/^enter\s+/i.test(text) && !isForm26NilMonthLineValue(text)) return text;
    return '';
  }
  if (value == null || text === '' || /^enter\s+/i.test(text)) return '';
  return text;
}

/** Form 26-A Excel export — preserve template columns and merged nil-month row. */
export async function buildForm26AWorkbookWithTemplateStyles({
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
    workbook.worksheets.find((ws) => /26[\s-]*a|dangerous/i.test(String(ws?.name || ''))) ||
    workbook.worksheets[0];
  if (!worksheet) throw new Error('Template worksheet not found.');

  const normalize = (txt) =>
    String(txt || '')
      .replace(/\r?\n/g, ' ')
      .replace(/\s+/g, ' ')
      .trim()
      .toLowerCase();
  const excelCellValueToString = (val) => {
    if (val == null) return '';
    if (typeof val === 'string' || typeof val === 'number' || typeof val === 'boolean') return String(val);
    if (val instanceof Date) return val.toISOString();
    if (typeof val === 'object') {
      if (Array.isArray(val.richText)) return val.richText.map((rt) => rt?.text || '').join('');
      if (val.text != null) return String(val.text);
      if (val.result != null) return String(val.result);
    }
    return '';
  };

  const maxScanRows = Math.max(120, worksheet.rowCount + 10);
  const maxScanCols = 20;
  let headerRow = parsedHeaderRowIndex != null && parsedHeaderRowIndex >= 0 ? parsedHeaderRowIndex + 1 : -1;
  if (headerRow < 1) throw new Error('Could not locate Form 26-A header row from parser.');
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
  for (let r = Math.max(headerRow, markerRow > 0 ? markerRow - 1 : headerRow); r <= Math.min(headerRow + 6, maxScanRows); r += 1) {
    const cols = [];
    for (let c = startCol; c <= maxScanCols; c += 1) {
      const t = normalize(excelCellValueToString(worksheet.getCell(r, c)?.value));
      if (/^\d+$/.test(t)) cols.push(c);
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
  const orderedCols =
    markerCols.length >= 3
      ? markerCols
      : Array.from({ length: Math.max(8, Array.isArray(headersToUse) ? headersToUse.length : 8) }, (_, i) => startCol + i);
  const hdrs = Array.isArray(headersToUse) ? headersToUse : [];
  const nilSpanInfo = resolveForm26ANilSpanInfo(hdrs);
  const calendarHeader = findForm26ACalendarYearHeader(hdrs);

  const data = headerFormData && typeof headerFormData === 'object' ? headerFormData : {};
  writeStatutoryHeaderFieldsToExcelJsWorksheet(worksheet, {
    headerFormData: data,
    parsedFormHeader,
    headerRowEnd: headerRow,
    maxScanCols: 20
  });

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
    let numericHits = 0;
    for (let i = 0; i < markerCols.length; i += 1) {
      const v = normalize(excelCellValueToString(worksheet.getCell(r, markerCols[i])?.value));
      if (/^\d+$/.test(v)) numericHits += 1;
    }
    return numericHits >= Math.max(3, Math.floor(markerCols.length * 0.6));
  };

  const clearFromRow = Math.max(1, dataStartRow);
  const clearToRow = Math.min(maxScanRows, Math.max(dataStartRow + Math.max(sourceRows.length, 1) + 20, clearFromRow + 20));
  for (let r = clearFromRow; r <= clearToRow; r += 1) {
    if (isNumericMarkerTemplateRow(r)) continue;
    for (let j = 0; j < orderedCols.length; j += 1) {
      worksheet.getCell(r, orderedCols[j]).value = '';
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
      ? sanitizeForm26AExportCell(
          Array.isArray(rowObj) ? rowValues[nilSpanInfo.startIdx] : rowObj?.[nilPrimary],
          nilPrimary,
          calendarHeader
        )
      : '';

    if (nilSpanInfo && isForm26NilMonthLineValue(nilValue)) {
      const startColIdx = orderedCols[nilSpanInfo.startIdx];
      const endColIdx =
        orderedCols[Math.min(nilSpanInfo.startIdx + nilSpanInfo.span - 1, orderedCols.length - 1)];
      if (startColIdx != null && endColIdx != null && endColIdx >= startColIdx) {
        if (calendarHeader) {
          const calIdx = hdrs.indexOf(calendarHeader);
          if (calIdx >= 0 && orderedCols[calIdx] != null) {
            const calVal = sanitizeForm26AExportCell(
              Array.isArray(rowObj) ? rowValues[calIdx] : rowObj?.[calendarHeader],
              calendarHeader,
              calendarHeader
            );
            if (calVal !== '') worksheet.getCell(targetRowNum, orderedCols[calIdx]).value = String(calVal);
          }
        }
        for (let c = startColIdx; c <= endColIdx; c += 1) {
          worksheet.getCell(targetRowNum, c).value = '';
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

    for (let j = 0; j < orderedCols.length; j += 1) {
      const header = hdrs[j] || '';
      const rawValue = Array.isArray(rowObj) ? rowValues[j] : rowObj?.[header];
      const value = sanitizeForm26AExportCell(rawValue, header, calendarHeader);
      if (value == null || value === '') continue;
      const cell = worksheet.getCell(targetRowNum, orderedCols[j]);
      if (/running.*serial|serial\s+number|\(2\)/i.test(header) && /^-?\d+$/.test(String(value))) {
        cell.value = Number(value);
      } else {
        cell.value = String(value);
        cell.alignment = { ...(cell.alignment || {}), wrapText: true, vertical: 'top' };
      }
    }
  }

  if (sourceRows.length > 0) {
    const tableColMin = orderedCols.length > 0 ? Math.min(...orderedCols) : startCol;
    const tableColMax = orderedCols.length > 0 ? Math.max(...orderedCols) : startCol + orderedCols.length - 1;
    ensureExcelJSDataRowsWithBorders(worksheet, {
      dataStartRow,
      dataRowCount: sourceRows.length,
      colFrom: tableColMin,
      colTo: tableColMax,
      templateRow: dataStartRow,
      templateBodyRows: 1
    });
  }

  const out = await workbook.xlsx.writeBuffer();
  const fileName =
    formFileName ||
    parsedFormHeader?.title?.replace(/[^a-zA-Z0-9]/g, '_') ||
    `Form_26A_Dangerous_Occurrences_${Date.now()}.xlsx`;
  const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  return { blob, fileName };
}
