import ExcelJS from 'exceljs';

/** Form XXVI (Tamil Nadu CLRA) — Muster with daily hours 1–31 under column band (10). */

export const FORM_XXVI_TN_DAY_BAND_PARENT = '10';

export function formXXVITamilNaduHeaderNorm(txt) {
  return String(txt || '')
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

export function matchesFormXXVIHint(blob) {
  const parts = String(blob || '').toLowerCase();
  if (/form[\s._-]*xxvii(?![a-z])/i.test(parts)) return false;
  return (
    /form[\s._-]*xxvi(?![a-z])/i.test(parts) ||
    /form[\s._-]*26(?!\d)/i.test(parts) ||
    /(?:^|[\s._-])xxvi(?=[\s._\W-]|$)/i.test(parts)
  );
}

export function isFormXXVITamilNaduContext(partsOrBlob) {
  const p = String(partsOrBlob || '').toLowerCase();
  return (
    /tamil[\s._-]*nadu|tamilnadu|form_xxvi[_\s-]*tamil|form[\s._-]*xxvi[_\s-]*tamil/.test(p)
  );
}

function buildFormXXVIContextBlob(formHeader, rowItem, fileName, tableHeaders, sheetText) {
  return [
    rowItem?.formName,
    rowItem?.FormName,
    rowItem?.formFileName,
    rowItem?.FormFileName,
    rowItem?.state,
    rowItem?.State,
    fileName,
    formHeader?.title,
    formHeader?.subtitle,
    formHeader?.reference,
    sheetText,
    Array.isArray(tableHeaders) ? tableHeaders.join(' ') : tableHeaders
  ]
    .join(' ')
    .toLowerCase();
}

/** Tamil Nadu Form XXVI — register with daily hours 1–31 (CLRA multi-tab workbook). */
export function isFormXXVITamilNaduClraContext(
  formHeader,
  rowItem,
  fileName,
  tableHeaders = '',
  sheetText = ''
) {
  const parts = buildFormXXVIContextBlob(formHeader, rowItem, fileName, tableHeaders, sheetText);
  if (!matchesFormXXVIHint(parts)) return false;
  return isFormXXVITamilNaduContext(parts);
}

export function isFormXXVITamilNaduDayHeaderKey(headerKey) {
  const h = String(headerKey || '').trim();
  if (/^10_(\d{1,2})$/.test(h)) {
    const n = Number(h.match(/^10_(\d{1,2})$/)[1]);
    return n >= 1 && n <= 31;
  }
  if (/^daily\s+hours\s+of\s+work_(\d{1,2})$/i.test(h)) {
    const n = Number(h.match(/^daily\s+hours\s+of\s+work_(\d{1,2})$/i)[1]);
    return n >= 1 && n <= 31;
  }
  if (/^\d{1,2}$/.test(h)) {
    const n = Number(h);
    return n >= 1 && n <= 31;
  }
  const m = h.match(/_(\d{1,2})$/);
  if (!m) return false;
  const n = Number(m[1]);
  return n >= 1 && n <= 31;
}

export function resolveFormXXVITamilNaduDayNumberFromHeader(header) {
  const h = String(header || '').trim();
  const m1 = h.match(/^10_(\d{1,2})$/i);
  if (m1) return Number(m1[1]);
  const m2 = h.match(/^daily\s+hours\s+of\s+work_(\d{1,2})$/i);
  if (m2) return Number(m2[1]);
  const m3 = h.match(/_(\d{1,2})$/);
  if (m3) return Number(m3[1]);
  if (/^\d{1,2}$/.test(h)) return Number(h);
  return 0;
}

export function listFormXXVITamilNaduDayHeaders(headers) {
  const out = [];
  (Array.isArray(headers) ? headers : []).forEach((header) => {
    const day = resolveFormXXVITamilNaduDayNumberFromHeader(header);
    if (day >= 1 && day <= 31) out.push({ header, day });
  });
  return out;
}

const headerAliasBucket = (key) => {
  const n = formXXVITamilNaduHeaderNorm(key).replace(/[^a-z0-9]/g, '');
  if (/^srno|serialno|serialnumber|sno/.test(n)) return 'srno';
  if (/nameoftheworkman|nameoftheworker|workmanname|workername|employeename/.test(n)) return 'name';
  if (/agesex|ageandsex/.test(n)) return 'agesex';
  if (/permanenthomeaddress|permanentaddress/.test(n)) return 'permaddr';
  if (/localaddress/.test(n)) return 'localaddr';
  if (/designation|natureofwork/.test(n)) return 'designation';
  if (/fathersname|husbandsname|fatherhusband/.test(n)) return 'fatherhusband';
  if (/dateofentry|dateofjoining|entryintoservice/.test(n)) return 'doj';
  if (/rateofwages|wages/.test(n) && !/overtime/.test(n)) return 'wages';
  return '';
};

export function readFormXXVITamilNaduCellValue(row, header) {
  if (!row || typeof row !== 'object') return '';
  const direct = row[header];
  if (direct != null && String(direct).trim() !== '') return direct;

  const day = resolveFormXXVITamilNaduDayNumberFromHeader(header);
  if (day >= 1 && day <= 31) {
    const aliases = [
      `${FORM_XXVI_TN_DAY_BAND_PARENT}_${day}`,
      `Daily hours of work_${day}`,
      `daily hours of work_${day}`,
      String(day),
      `Day ${day}`,
      `DATES_${day}`
    ];
    for (let i = 0; i < aliases.length; i += 1) {
      const val = row[aliases[i]];
      if (val != null && String(val).trim() !== '') return val;
    }
  }

  const bucket = headerAliasBucket(header);
  if (bucket) {
    for (const [key, val] of Object.entries(row)) {
      if (String(key).startsWith('__')) continue;
      if (headerAliasBucket(key) === bucket && val != null && String(val).trim() !== '') {
        return val;
      }
    }
  }

  const target = formXXVITamilNaduHeaderNorm(header);
  if (target) {
    const keys = Object.keys(row);
    const exact = keys.filter((k) => formXXVITamilNaduHeaderNorm(k) === target);
    if (exact.length === 1) return row[exact[0]];
    const fuzzy = keys.find((k) => {
      const n = formXXVITamilNaduHeaderNorm(k);
      return n && (n.includes(target) || target.includes(n));
    });
    if (fuzzy) return row[fuzzy];
  }

  return '';
}

export function detectFormXXVITamilNaduDayColumnMap(getCell, headerRow, startCol, maxCols) {
  const dayCols = new Map();
  const scanFrom = Math.max(startCol, 1);
  const scanTo = Math.max(maxCols, scanFrom + 45);
  for (let r = headerRow; r <= headerRow + 4; r += 1) {
    for (let c = scanFrom; c < scanTo; c += 1) {
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
    const fuzzy = workbook.worksheets.find((ws) =>
      /xxvi|form[\s._-]*26(?!\d)/i.test(String(ws?.name || ''))
    );
    if (fuzzy) return fuzzy;
  }
  return (
    workbook.worksheets.find((ws) => /xxvi|form[\s._-]*26(?!\d)/i.test(String(ws?.name || ''))) ||
    workbook.worksheets[0]
  );
}

/** Form XXVI Tamil Nadu Excel export — map prefix columns + day grid (10_1..10_31) to template positions. */
export async function buildFormXXVITamilNaduWorkbookWithTemplateStyles({
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

  const normalize = formXXVITamilNaduHeaderNorm;
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
          /sr\.?\s*no|serial|s\.?\s*no/.test(t) ||
          /name\s+of\s+the\s+workman|name\s+of\s+the\s+worker/.test(t) ||
          /daily\s+hours/.test(t)
        ) {
          hits += 1;
        }
      }
      if (hits >= 2) {
        headerRow = r;
        break;
      }
    }
  }
  if (headerRow < 1) throw new Error('Could not locate Form XXVI header row.');

  let startCol = 1;
  for (let c = 1; c <= maxScanCols; c += 1) {
    const t = normalize(excelCellValueToString(worksheet.getCell(headerRow, c)?.value));
    if (t) {
      startCol = c;
      break;
    }
  }
  if (parsedTableStartCol != null && Number(parsedTableStartCol) >= 0) {
    startCol = Number(parsedTableStartCol) + 1;
  }

  const getCell = (r, c) => excelCellValueToString(worksheet.getCell(r, c)?.value);
  const dayColumnMap = detectFormXXVITamilNaduDayColumnMap(getCell, headerRow, startCol, maxScanCols);
  const markerCols = [...dayColumnMap.entries()]
    .sort((a, b) => a[0] - b[0])
    .map(([, col]) => col);

  let markerRow = -1;
  if (markerCols.length >= 5) {
    for (let r = headerRow; r <= Math.min(headerRow + 5, maxScanRows); r += 1) {
      let hits = 0;
      for (let i = 0; i < markerCols.length; i += 1) {
        const t = normalize(getCell(r, markerCols[i]));
        if (/^\d{1,2}$/.test(t) && Number(t) >= 1 && Number(t) <= 31) hits += 1;
      }
      if (hits >= 5) {
        markerRow = r;
        break;
      }
    }
  }

  let dataStartRow =
    parsedDataStartIndex != null && parsedDataStartIndex >= 0
      ? parsedDataStartIndex + 1
      : markerRow > 0
        ? markerRow + 1
        : headerRow + 3;
  if (markerRow > 0) {
    dataStartRow = Math.max(dataStartRow, markerRow + 1);
  }

  const sourceHeaders = Array.isArray(headersToUse) ? [...headersToUse] : [];
  const dayStartIdx = sourceHeaders.findIndex((h) => isFormXXVITamilNaduDayHeaderKey(h));
  const prefixCount = dayStartIdx >= 0 ? dayStartIdx : sourceHeaders.length;

  const buildOrderedCols = () => {
    if (markerCols.length >= 5 && dayStartIdx >= 0) {
      const cols = [];
      for (let j = 0; j < prefixCount; j += 1) cols.push(startCol + j);
      const dayCount = Math.max(sourceHeaders.length - dayStartIdx, markerCols.length);
      for (let d = 0; d < dayCount; d += 1) {
        cols.push(markerCols[d] ?? startCol + prefixCount + d);
      }
      const rightCount = sourceHeaders.length - dayStartIdx - dayCount;
      if (rightCount > 0) {
        const rightStart = (markerCols[markerCols.length - 1] ?? startCol + prefixCount) + 1;
        for (let r = 0; r < rightCount; r += 1) cols.push(rightStart + r);
      }
      return cols;
    }
    return Array.from({ length: Math.max(12, sourceHeaders.length) }, (_, i) => startCol + i);
  };
  const orderedCols = buildOrderedCols();
  const hdrs = sourceHeaders;

  const getRowValueForHeader = (rowObj, header, headerIndex) => {
    if (!rowObj || typeof rowObj !== 'object' || Array.isArray(rowObj)) return '';
    const picked = pickExportCell(readFormXXVITamilNaduCellValue(rowObj, header));
    if (picked !== '') return picked;
    if (headerIndex != null && hdrs[headerIndex]) {
      const alt = pickExportCell(readFormXXVITamilNaduCellValue(rowObj, hdrs[headerIndex]));
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
      if (out[0] == null || String(out[0]).trim() === '') out[0] = rowIndex + 1;
      const colCount = Math.max(orderedCols.length, hdrs.length, out.length);
      return Array.from({ length: colCount }, (_, idx) => out[idx] ?? '');
    }
    const colCount = Math.max(orderedCols.length, hdrs.length);
    const out = Array.from({ length: colCount }, (_, idx) => {
      const header = hdrs[idx] || '';
      return header ? getRowValueForHeader(row, header, idx) : '';
    });
    if (out[0] == null || String(out[0]).trim() === '') out[0] = rowIndex + 1;
    return out;
  };

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

  const sourcePrimary =
    Array.isArray(mappedData) && mappedData.length > 0
      ? mappedData
      : Array.isArray(mappedRowMatrix)
        ? mappedRowMatrix
        : [];
  const sourceRows = sourcePrimary
    .filter((row) => rowLooksMeaningful(row))
    .map((row, idx) => toOrderedValues(row, idx));

  const isNumericMarkerTemplateRow = (r) => {
    if (!markerCols || markerCols.length < 5) return false;
    let numericHits = 0;
    for (let i = 0; i < markerCols.length; i += 1) {
      const v = normalize(getCell(r, markerCols[i]));
      if (/^\d+$/.test(v)) numericHits += 1;
    }
    return numericHits >= Math.max(5, Math.floor(markerCols.length * 0.6));
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
    for (let j = 0; j < orderedCols.length; j += 1) {
      const value = row[j];
      if (value == null || value === '') continue;
      const cell = worksheet.getCell(dataStartRow + i, orderedCols[j]);
      if (
        j === 0 &&
        (typeof value === 'number' || (typeof value === 'string' && /^-?\d+(\.\d+)?$/.test(String(value).trim())))
      ) {
        cell.value = Number(value);
      } else {
        cell.value = String(value);
      }
    }
  }

  const targetSheetName = worksheet.name;
  workbook.worksheets
    .filter((ws) => ws.name !== targetSheetName)
    .forEach((ws) => {
      workbook.removeWorksheet(ws.id);
    });

  const out = await workbook.xlsx.writeBuffer();
  const fileName =
    formFileName ||
    parsedFormHeader?.title?.replace(/[^a-zA-Z0-9]/g, '_') ||
    `Form_XXVI_TamilNadu_${Date.now()}.xlsx`;
  const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  return { blob, fileName };
}
