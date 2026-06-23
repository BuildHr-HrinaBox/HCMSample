import ExcelJS from 'exceljs';

/** AP Form XXII — Register of Employment (Shops & Establishment; P / A / WO day grid). */

export function formXXIIAPHeaderNorm(txt) {
  return String(txt || '')
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

export function resolveFormXXIIAPDayNumberFromHeader(header) {
  const h = String(header || '').trim();
  if (/^\d{1,2}$/.test(h)) {
    const n = Number(h);
    return n >= 1 && n <= 31 ? n : 0;
  }
  const m =
    h.match(/(?:dates?|attendance)[^0-9]*(\d{1,2})$/i) ||
    h.match(/_(\d{1,2})$/);
  if (m) {
    const n = Number(m[1]);
    return n >= 1 && n <= 31 ? n : 0;
  }
  return 0;
}

export function isFormXXIIAPDayHeaderKey(header) {
  return resolveFormXXIIAPDayNumberFromHeader(header) >= 1;
}

export function listFormXXIIAPDayHeaders(headers) {
  const out = [];
  (Array.isArray(headers) ? headers : []).forEach((header) => {
    const day = resolveFormXXIIAPDayNumberFromHeader(header);
    if (day >= 1 && day <= 31) out.push({ header, day });
  });
  return out.sort((a, b) => a.day - b.day);
}

/** Merge template/modal headers with day keys present on export rows (SampleData may omit day columns). */
export function resolveFormXXIIExportHeaders(headersToUse, mappedData) {
  const base = Array.isArray(headersToUse) ? headersToUse.filter(Boolean) : [];
  const dayByNum = new Map();
  base.forEach((header) => {
    const day = resolveFormXXIIAPDayNumberFromHeader(header);
    if (day >= 1 && day <= 31 && !dayByNum.has(day)) dayByNum.set(day, header);
  });
  (Array.isArray(mappedData) ? mappedData : []).forEach((row) => {
    if (!row || typeof row !== 'object') return;
    Object.keys(row).forEach((key) => {
      if (String(key).startsWith('__')) return;
      const day = resolveFormXXIIAPDayNumberFromHeader(key);
      if (day >= 1 && day <= 31 && !dayByNum.has(day)) dayByNum.set(day, key);
    });
  });
  const prefix = base.filter((h) => !isFormXXIIAPDayHeaderKey(h));
  const dayHeaders = [...dayByNum.entries()]
    .sort((a, b) => a[0] - b[0])
    .map(([, header]) => header);
  if (dayHeaders.length === 0) return base;
  if (base.some((h) => isFormXXIIAPDayHeaderKey(h))) {
    const missing = dayHeaders.filter((h) => !base.includes(h));
    return missing.length > 0 ? [...base, ...missing] : base;
  }
  return [...prefix, ...dayHeaders];
}

export function readFormXXIIAPCellValue(row, header) {
  if (!row || typeof row !== 'object') return '';
  if (Object.prototype.hasOwnProperty.call(row, header)) {
    const direct = row[header];
    if (direct != null && String(direct).trim() !== '') return direct;
  }

  const day = resolveFormXXIIAPDayNumberFromHeader(header);
  if (day >= 1 && day <= 31) {
    const aliases = [
      `hours_${day}`,
      `Dates (Attendance: P = Present, A = Absent, L = Leave, W = Weekly Holiday)_${day}`,
      `Dates (Attendance: P = Present, A = Absent, L = Leave, H = Holiday, WO = Week Off)_${day}`,
      `Dates (Attendance: P = Present, A = Absent, L = Leave, WO = Week Off)_${day}`,
      `DATES_${day}`,
      `Dates_${day}`,
      String(day)
    ];
    for (let i = 0; i < aliases.length; i += 1) {
      const val = row[aliases[i]];
      if (val != null && String(val).trim() !== '') return val;
    }
    for (const [key, val] of Object.entries(row)) {
      if (val == null || String(val).trim() === '') continue;
      if (resolveFormXXIIAPDayNumberFromHeader(key) === day) return val;
    }
  }

  const target = formXXIIAPHeaderNorm(header);
  if (target) {
    for (const [key, val] of Object.entries(row)) {
      if (val == null || String(val).trim() === '') continue;
      const nk = formXXIIAPHeaderNorm(key);
      if (nk === target || (nk && (nk.includes(target) || target.includes(nk)))) return val;
    }
  }
  return '';
}

export function detectFormXXIIAPDayColumnMap(getCell, headerRow, maxScanRows, maxScanCols) {
  const dayCols = new Map();
  const scanTo = Math.max(maxScanCols, 45);
  const scanEndRow = Math.min(headerRow + 3, maxScanRows);
  for (let r = headerRow; r <= scanEndRow; r += 1) {
    for (let c = 1; c <= scanTo; c += 1) {
      const raw = String(getCell(r, c) || '')
        .trim()
        .replace(/[()]/g, '');
      if (!/^\d{1,2}$/.test(raw)) continue;
      const n = Number(raw);
      if (n >= 1 && n <= 31 && !dayCols.has(n)) {
        dayCols.set(n, c);
      }
    }
  }
  return dayCols;
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

function pickExportCell(v) {
  if (v == null || String(v).trim() === '' || /^enter\s+/i.test(String(v).trim())) return '';
  return v;
}

function rowLooksMeaningful(row) {
  const hasAttendanceCode = (vals) =>
    vals.some((v) => /^(P|A|WO|H|L|WOP|OD|SL|CL|EL)$/i.test(String(v).trim()));
  if (Array.isArray(row)) {
    const vals = row.map((v) => String(v ?? '').trim()).filter(Boolean);
    if (vals.length === 0) return false;
    return (
      hasAttendanceCode(vals) ||
      vals.some((v) => /[a-z]/i.test(v)) ||
      !vals.every((v) => /^-?\d+(\.\d+)?$/.test(v))
    );
  }
  if (row && typeof row === 'object') {
    const vals = Object.values(row).map((v) => String(v ?? '').trim()).filter(Boolean);
    if (vals.length === 0) return false;
    return (
      hasAttendanceCode(vals) ||
      vals.some((v) => /[a-z]/i.test(v)) ||
      !vals.every((v) => /^-?\d+(\.\d+)?$/.test(v))
    );
  }
  return false;
}

function resolveWorksheet(workbook, sheetNameHint) {
  const hint = String(sheetNameHint || '').trim();
  if (hint) {
    const exact = workbook.worksheets.find((ws) => String(ws?.name || '') === hint);
    if (exact) return exact;
    const hintNorm = hint.toLowerCase();
    const fuzzy = workbook.worksheets.find((ws) => {
      const n = String(ws?.name || '').toLowerCase();
      return n === hintNorm || n.includes(hintNorm) || hintNorm.includes(n);
    });
    if (fuzzy) return fuzzy;
  }
  return (
    workbook.worksheets.find((ws) =>
      /xxii|form[\s._-]*22(?!\d)|register\s+of\s+employment/i.test(String(ws?.name || ''))
    ) || workbook.worksheets[0]
  );
}

/**
 * Form XXII AP Excel export — same model as Form 25 TN: prefix columns + day marker columns (1–31).
 */
export async function buildFormXXIIAPWorkbookWithTemplateStyles({
  templateArrayBuffer,
  mappedData,
  mappedRowMatrix,
  headersToUse,
  parsedHeaderRowIndex,
  parsedDataStartIndex,
  parsedTableStartCol,
  parsedFormHeader,
  headerFormData,
  formFileName,
  sheetNameHint
}) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(templateArrayBuffer);
  const worksheet = resolveWorksheet(workbook, sheetNameHint);
  if (!worksheet) throw new Error('Template worksheet not found.');

  const normalize = formXXIIAPHeaderNorm;
  const maxScanRows = Math.max(120, worksheet.rowCount + 10);
  const maxScanCols = 300;

  let headerRow = parsedHeaderRowIndex != null && parsedHeaderRowIndex >= 0 ? parsedHeaderRowIndex + 1 : -1;
  if (headerRow < 1) {
    for (let r = 1; r <= maxScanRows; r += 1) {
      let hits = 0;
      for (let c = 1; c <= maxScanCols; c += 1) {
        const t = normalize(excelCellValueToString(worksheet.getCell(r, c)?.value));
        if (!t) continue;
        if (
          /s\.?\s*no|serial|sr\.?\s*no/.test(t) ||
          /name\s+of\s+the\s+employee/.test(t) ||
          /time\s+at\s+which|rest\s+interval|dates|attendance/.test(t)
        ) {
          hits += 1;
        }
      }
      if (hits >= 3) {
        headerRow = r;
        break;
      }
    }
  }
  if (headerRow < 1) throw new Error('Could not locate Form XXII header row.');

  const getCell = (r, c) => excelCellValueToString(worksheet.getCell(r, c)?.value);

  let nameCol = -1;
  for (let c = 1; c <= maxScanCols; c += 1) {
    const t = normalize(getCell(headerRow, c));
    if (/name\s+of\s+the\s+employee/.test(t)) {
      nameCol = c;
      break;
    }
  }

  const sourceHeadersPreview = Array.isArray(headersToUse) ? headersToUse : [];
  const firstHeaderIsSerial =
    sourceHeadersPreview.length > 0 &&
    /^(s\.?\s*no\.?|serial(\s+number)?|sr\.?\s*no\.?)$/i.test(
      String(sourceHeadersPreview[0] || '').trim()
    );

  let startCol = nameCol > 0 ? nameCol : 1;
  if (firstHeaderIsSerial) {
    startCol =
      parsedTableStartCol != null && Number(parsedTableStartCol) >= 0
        ? Number(parsedTableStartCol) + 1
        : 1;
  } else if (nameCol > 0) {
    startCol = nameCol;
  } else if (parsedTableStartCol != null && Number(parsedTableStartCol) >= 0) {
    startCol = Number(parsedTableStartCol) + 1;
  }

  let markerRow = -1;
  const dayColumnMap = detectFormXXIIAPDayColumnMap(getCell, headerRow, maxScanRows, maxScanCols);
  const markerCols = [...dayColumnMap.entries()]
    .sort((a, b) => a[0] - b[0])
    .map(([, col]) => col);
  if (markerCols.length >= 2) {
    for (let r = headerRow; r <= Math.min(headerRow + 3, maxScanRows); r += 1) {
      let hits = 0;
      for (let d = 1; d <= 31; d += 1) {
        const col = dayColumnMap.get(d);
        if (!col) continue;
        const t = String(getCell(r, col) || '').trim();
        if (t === String(d)) hits += 1;
      }
      if (hits >= Math.max(2, Math.min(5, markerCols.length))) {
        markerRow = r;
        break;
      }
    }
  }

  let dataStartRow =
    parsedDataStartIndex != null && parsedDataStartIndex >= 0
      ? parsedDataStartIndex + 1
      : markerCols.length >= 2 && markerRow > 0
        ? markerRow + 1
        : headerRow + 1;
  if (markerCols.length >= 2 && markerRow > 0) {
    dataStartRow = Math.max(dataStartRow, markerRow + 1);
  }

  const sourcePrimary =
    Array.isArray(mappedData) && mappedData.length > 0
      ? mappedData
      : Array.isArray(mappedRowMatrix)
        ? mappedRowMatrix
        : [];
  const sourceHeaders = resolveFormXXIIExportHeaders(headersToUse, sourcePrimary);
  const dayStartIdx = sourceHeaders.findIndex((h) => isFormXXIIAPDayHeaderKey(h));
  const prefixCount = dayStartIdx >= 0 ? dayStartIdx : sourceHeaders.length;

  const buildOrderedCols = () => {
    if (markerCols.length >= 2 && dayStartIdx >= 0) {
      const cols = [];
      for (let j = 0; j < prefixCount; j += 1) cols.push(startCol + j);
      const dayCount = Math.max(sourceHeaders.length - dayStartIdx, markerCols.length);
      for (let d = 0; d < dayCount; d += 1) {
        cols.push(markerCols[d] ?? startCol + prefixCount + d);
      }
      return cols;
    }
    if (markerCols.length >= 2) {
      const cols = [];
      for (let j = 0; j < prefixCount; j += 1) cols.push(startCol + j);
      for (let d = 0; d < markerCols.length; d += 1) cols.push(markerCols[d]);
      return cols;
    }
    return Array.from({ length: Math.max(12, sourceHeaders.length) }, (_, i) => startCol + i);
  };
  const orderedCols = buildOrderedCols();
  const hdrs = sourceHeaders;

  const getDayValueByNumber = (rowObj, dayNum) => {
    if (!rowObj || typeof rowObj !== 'object' || !dayNum) return '';
    for (const [key, val] of Object.entries(rowObj)) {
      if (String(key).startsWith('__')) continue;
      if (resolveFormXXIIAPDayNumberFromHeader(key) !== dayNum) continue;
      const picked = pickExportCell(val);
      if (picked !== '') return picked;
    }
    return '';
  };

  const getRowValueForHeader = (rowObj, header, headerIndex) => {
    if (!rowObj || typeof rowObj !== 'object' || Array.isArray(rowObj)) return '';
    const picked = pickExportCell(readFormXXIIAPCellValue(rowObj, header));
    if (picked !== '') return picked;
    const dayNum = resolveFormXXIIAPDayNumberFromHeader(header);
    if (dayNum >= 1 && dayNum <= 31) {
      const byDay = getDayValueByNumber(rowObj, dayNum);
      if (byDay !== '') return byDay;
    }
    if (headerIndex != null && hdrs[headerIndex]) {
      const alt = pickExportCell(readFormXXIIAPCellValue(rowObj, hdrs[headerIndex]));
      if (alt !== '') return alt;
    }
    if (headerIndex != null && headerIndex >= 0) {
      const vals = Object.values(rowObj);
      const indexed = pickExportCell(vals[headerIndex]);
      if (indexed !== '') return indexed;
    }
    return '';
  };

  const toOrderedValues = (row, rowIndex) => {
    if (Array.isArray(row)) {
      const out = [...row];
      const colCount = Math.max(orderedCols.length, hdrs.length, out.length);
      return Array.from({ length: colCount }, (_, idx) => out[idx] ?? '');
    }
    const colCount = Math.max(orderedCols.length, hdrs.length);
    return Array.from({ length: colCount }, (_, idx) => {
      const header = hdrs[idx] || '';
      if (header) return getRowValueForHeader(row, header, idx);
      if (dayStartIdx >= 0 && idx >= prefixCount) {
        const dayNum = idx - prefixCount + 1;
        return getDayValueByNumber(row, dayNum);
      }
      return '';
    });
  };

  let snoCol = -1;
  for (let c = 1; c <= 20; c += 1) {
    const t = normalize(getCell(headerRow, c));
    if (/^s\.?\s*no\.?$/.test(t) || t === 'sno' || /^sr\.?\s*no\.?$/.test(t)) {
      snoCol = c;
      break;
    }
  }

  const headerFields = parsedFormHeader?.fields;
  const headerValues = headerFormData && typeof headerFormData === 'object' ? headerFormData : {};
  if (Array.isArray(headerFields) && headerFields.length > 0) {
    for (let r = 1; r < headerRow; r += 1) {
      for (let c = 1; c <= 80; c += 1) {
        const cellStr = String(getCell(r, c) || '').trim();
        if (!cellStr) continue;
        for (const field of headerFields) {
          const label = String(field?.label || '').trim();
          if (!label) continue;
          if (cellStr === label || cellStr.startsWith(label) || label.startsWith(cellStr)) {
            const value = headerValues[field.key] ?? field.value ?? '';
            if (value != null && String(value).trim() !== '') {
              worksheet.getCell(r, c + 1).value = String(value);
            }
            break;
          }
        }
      }
    }
  }

  const filteredRows = sourcePrimary.filter((row) => rowLooksMeaningful(row));
  const sourceRows = filteredRows.map((row, idx) => toOrderedValues(row, idx));

  const writeDayAttendanceGrid = (rowObj, excelRowIdx) => {
    if (!rowObj || dayColumnMap.size === 0) return;
    for (let day = 1; day <= 31; day += 1) {
      const col = dayColumnMap.get(day);
      if (!col) continue;
      const val = getDayValueByNumber(rowObj, day);
      if (val == null || val === '') continue;
      worksheet.getCell(excelRowIdx, col).value = String(val);
    }
  };

  const isNumericMarkerTemplateRow = (r) => {
    if (!markerCols || markerCols.length < 2) return false;
    let numericHits = 0;
    for (let i = 0; i < markerCols.length; i += 1) {
      const v = normalize(getCell(r, markerCols[i]));
      if (/^\d+$/.test(v)) numericHits += 1;
    }
    return numericHits >= Math.max(2, Math.floor(markerCols.length * 0.5));
  };
  const isHeaderBandRow = (r) => r >= headerRow && r < dataStartRow;
  const clearFromRow = Math.max(1, dataStartRow);
  const clearToRow = Math.min(maxScanRows, Math.max(dataStartRow + sourceRows.length + 30, clearFromRow + 30));
  for (let r = clearFromRow; r <= clearToRow; r += 1) {
    if (isNumericMarkerTemplateRow(r) || isHeaderBandRow(r)) continue;
    for (let j = 0; j < orderedCols.length; j += 1) {
      worksheet.getCell(r, orderedCols[j]).value = '';
    }
  }

  for (let i = 0; i < sourceRows.length; i += 1) {
    const row = sourceRows[i] || [];
    const rowObj = filteredRows[i];
    const excelRowIdx = dataStartRow + i;
    if (snoCol > 0) {
      worksheet.getCell(excelRowIdx, snoCol).value = i + 1;
    }
    for (let j = 0; j < orderedCols.length; j += 1) {
      const value = row[j];
      if (value == null || value === '') continue;
      const hdr = hdrs[j] || '';
      if (isFormXXIIAPDayHeaderKey(hdr)) continue;
      const cell = worksheet.getCell(excelRowIdx, orderedCols[j]);
      if (
        typeof value === 'number' ||
        (typeof value === 'string' &&
          /^-?\d+(\.\d+)?$/.test(String(value).trim()) &&
          j > 0 &&
          !/^(P|A|WO|H|L)$/i.test(String(value).trim()))
      ) {
        if (/^age$/i.test(String(hdr).trim()) || resolveFormXXIIAPDayNumberFromHeader(hdr) === 0) {
          cell.value = Number(value);
        } else {
          cell.value = String(value);
        }
      } else {
        cell.value = String(value);
      }
    }
    writeDayAttendanceGrid(rowObj, excelRowIdx);
  }

  const out = await workbook.xlsx.writeBuffer();
  const fileName =
    formFileName ||
    parsedFormHeader?.title?.replace(/[^a-zA-Z0-9]/g, '_') ||
    `Form_XXII_AP_${Date.now()}.xlsx`;
  const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  return { blob, fileName };
}
