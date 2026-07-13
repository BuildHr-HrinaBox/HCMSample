import ExcelJS from 'exceljs';
import { ensureExcelJSDataRowsWithBorders } from '../../utils/excelTableBorders';
import { writeStatutoryHeaderFieldsToExcelJsWorksheet } from '../../utils/statutorySiteCompanyHeaders';
import {
  enrichFormAGJGujaratRowsForExport,
  filterFormAGJGujaratExportRows,
  formAGJGujaratHeaderNorm,
  formAGJHeaderAliasBucket,
  getFormAGJGujaratRowValueForHeader,
  remapFormAGJGujaratRowsToHeaders,
  resolveFormAGJGujaratTableHeaders,
  rowHasMeaningfulFormAGJGujaratExportData,
} from './formAGJGujarat';

/** Rajasthan Form A — employee / workman register (Shops & Establishments). */

const FORM_A_RJ_HIDDEN_HEADER_BUCKETS = new Set([
  'employeeId',
  'gender',
  'categoryAddress',
  'bankAccount',
  'bankName',
  'bankBranchIfsc',
]);

/** All export fields on one consecutive row per employee (no blank spacer rows). */
const FORM_A_RJ_MAIN_ROW_OFFSET = 0;
const FORM_A_RJ_EXPORT_ROWS_PER_EMPLOYEE = 1;

export function isFormARajasthanHiddenTableHeader(header) {
  const bucket = formAGJHeaderAliasBucket(formAGJGujaratHeaderNorm(header));
  return FORM_A_RJ_HIDDEN_HEADER_BUCKETS.has(bucket);
}

export function isFormARajasthanContext(
  formHeader,
  rowItem,
  fileName,
  sheetText = '',
  tableHeaders = []
) {
  const parts = [
    rowItem?.formName,
    rowItem?.FormName,
    rowItem?.formFileName,
    rowItem?.FormFileName,
    rowItem?.siteState,
    rowItem?.SiteState,
    rowItem?.state,
    rowItem?.State,
    fileName,
    formHeader?.title,
    formHeader?.subtitle,
    formHeader?.reference,
    sheetText,
    Array.isArray(tableHeaders) ? tableHeaders.join(' ') : '',
  ]
    .filter((x) => x != null && String(x).trim() !== '')
    .join(' ')
    .toLowerCase();

  if (/gujarat|\b_gj\b|form[\s._-]*a[\s._-]*gj|form_a_gj/.test(parts)) return false;
  if (/maternity\s+benefit|register\s+of\s+muster\s+roll/.test(parts)) return false;

  const hasRajasthan =
    /rajasthan|\b_rj\b|form[\s._-]*a[\s._-]*rj|form_a_rj/.test(parts);
  const hasFormA = /\bform[\s._-]*a\b/.test(parts);
  const hasEmployeeWorkmanFormat =
    /format\s+of\s+employee/.test(parts) ||
    (/employee\s*\/\s*workman\s*\/\s*worker/.test(parts) &&
      !/register\s+of\s+wages|rate\s+of\s+wage/.test(parts));

  if (hasRajasthan && (hasFormA || hasEmployeeWorkmanFormat)) return true;
  if (hasRajasthan && /\bform_a_rj\b/.test(parts)) return true;

  return false;
}

export function resolveFormARajasthanTableHeaders(tableHeaders) {
  const parsed = Array.isArray(tableHeaders)
    ? tableHeaders.map((h) => String(h || '').trim()).filter(Boolean)
    : [];
  return parsed.filter((h) => !isFormARajasthanHiddenTableHeader(h));
}

/** Full template headers for Excel export (includes bank / ID columns hidden in autofill UI). */
export function resolveFormARajasthanExportHeaders(tableHeaders) {
  return resolveFormAGJGujaratTableHeaders(tableHeaders);
}

export function remapFormARajasthanRowsToHeaders(rows, sourceHeaders, targetHeaders) {
  const src = Array.isArray(sourceHeaders) ? sourceHeaders : [];
  const tgt = resolveFormARajasthanTableHeaders(targetHeaders);
  if (!Array.isArray(rows)) return [];
  return rows.map((row, rowIndex) => {
    if (!row || typeof row !== 'object') return {};
    const out = {};
    tgt.forEach((targetHeader, colIdx) => {
      let val = '';
      if (Object.prototype.hasOwnProperty.call(row, targetHeader)) {
        val = row[targetHeader];
      } else if (src[colIdx] && Object.prototype.hasOwnProperty.call(row, src[colIdx])) {
        val = row[src[colIdx]];
      } else {
        const bucket = formAGJHeaderAliasBucket(formAGJGujaratHeaderNorm(targetHeader));
        for (const [k, v] of Object.entries(row)) {
          if (formAGJHeaderAliasBucket(formAGJGujaratHeaderNorm(k)) === bucket) {
            val = v;
            break;
          }
        }
      }
      if (
        (val == null || val === '') &&
        formAGJHeaderAliasBucket(formAGJGujaratHeaderNorm(targetHeader)) === 'sno'
      ) {
        val = rowIndex + 1;
      }
      out[targetHeader] = val == null ? '' : val;
    });
    return out;
  });
}

export function filterFormARajasthanExportRows(rows, headers) {
  const exportHdrs = resolveFormARajasthanExportHeaders(headers);
  return filterFormAGJGujaratExportRows(
    remapFormAGJGujaratRowsToHeaders(Array.isArray(rows) ? rows : [], exportHdrs, exportHdrs),
    exportHdrs
  );
}

export function rowHasMeaningfulFormARajasthanExportData(row, headers) {
  return rowHasMeaningfulFormAGJGujaratExportData(row, resolveFormARajasthanExportHeaders(headers));
}

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

function excelCellLooksLikeSerialHeader(text) {
  const t = formAGJGujaratHeaderNorm(text).replace(/[^a-z0-9 ]/g, ' ').replace(/\s+/g, ' ').trim();
  return /^sr\.?\s*no|^s\.?\s*no|serial|sl\.?\s*no/.test(t);
}

function buildMergeTopLeftResolver(worksheet) {
  const cache = new Map();
  return (r, c) => {
    const key = `${r}:${c}`;
    if (cache.has(key)) return cache.get(key);
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
    cache.set(key, topLeft);
    return topLeft;
  };
}

function detectFormARajasthanTableLayout(worksheet, hints = {}) {
  const getMergeTopLeft = buildMergeTopLeftResolver(worksheet);
  const getMergedAwareCellText = (r, c) => {
    const tl = getMergeTopLeft(r, c);
    return excelCellValueToString(worksheet.getCell(tl.r, tl.c)?.value).trim();
  };

  let headerRow =
    hints.parsedHeaderRowIndex != null && hints.parsedHeaderRowIndex >= 0
      ? hints.parsedHeaderRowIndex + 1
      : -1;
  let startCol =
    hints.parsedTableStartCol != null && hints.parsedTableStartCol >= 0
      ? hints.parsedTableStartCol + 1
      : 1;

  if (headerRow < 1) {
    const maxScanRows = Math.max(35, worksheet.rowCount + 5);
    for (let r = 1; r <= maxScanRows; r += 1) {
      for (let c = 1; c <= 25; c += 1) {
        if (excelCellLooksLikeSerialHeader(getMergedAwareCellText(r, c))) {
          headerRow = r;
          startCol = c;
          break;
        }
      }
      if (headerRow > 0) break;
    }
  }
  if (headerRow < 1) return null;

  const bucketPlaced = new Set();
  const templateCols = [];
  const scanColEnd = Math.max(startCol + 50, 55);
  for (let c = startCol; c < scanColEnd; c += 1) {
    const label = getMergedAwareCellText(headerRow, c);
    if (!label || /^\d+$/.test(label)) continue;
    const bucket = formAGJHeaderAliasBucket(formAGJGujaratHeaderNorm(label));
    if (!bucket || bucketPlaced.has(bucket)) continue;
    bucketPlaced.add(bucket);
    templateCols.push({
      col: c,
      label,
      bucket,
      subRowOffset: FORM_A_RJ_MAIN_ROW_OFFSET,
    });
    if (templateCols.length >= 45) break;
  }
  if (templateCols.length < 4) return null;

  let dataStartRow =
    hints.parsedDataStartIndex != null && hints.parsedDataStartIndex >= 0
      ? hints.parsedDataStartIndex + 1
      : headerRow + 1;

  let seqHits = 0;
  for (let i = 0; i < templateCols.length; i += 1) {
    const t = getMergedAwareCellText(dataStartRow, templateCols[i].col).replace(/\s+/g, '').trim();
    if (t === String(i + 1)) seqHits += 1;
  }
  if (seqHits >= Math.max(3, Math.floor(templateCols.length * 0.35))) {
    dataStartRow = headerRow + 2;
  }

  return {
    headerRow,
    dataStartRow,
    templateCols,
    rowsPerEmployee: FORM_A_RJ_EXPORT_ROWS_PER_EMPLOYEE,
  };
}

function writeFormARajasthanCellValue(cell, bucket, val) {
  if (val == null || val === '') {
    cell.value = '';
    return;
  }
  if (bucket === 'sno') {
    const n = Number(String(val).replace(/[,]/g, '').trim());
    cell.value = Number.isFinite(n) ? n : String(val);
  } else if (bucket === 'mobile' || bucket === 'uan' || bucket === 'bankAccount' || bucket === 'aadhaar' || bucket === 'esic') {
    const digits = String(val).replace(/\D/g, '');
    cell.value = digits || String(val);
    if (digits) cell.numFmt = '0';
  } else if (bucket === 'pan') {
    cell.value = String(val).trim().toUpperCase();
  } else {
    cell.value = String(val);
  }
  cell.alignment = { ...(cell.alignment || {}), vertical: 'middle', wrapText: true };
}

/** Rajasthan Form A Excel export — one consecutive row per employee, no empty spacer rows. */
export async function buildFormARajasthanWorkbookWithTemplateStyles({
  templateArrayBuffer,
  mappedData,
  headersToUse,
  parsedFormHeader,
  formFileName,
  headerFormData,
  parsedHeaderRowIndex,
  parsedDataStartIndex,
  parsedTableStartCol,
  employeesOverride = null,
  formatStatutoryDateDisplay = null,
}) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(templateArrayBuffer);
  const worksheet = workbook.worksheets[0];
  if (!worksheet) throw new Error('Template worksheet not found.');

  const layout = detectFormARajasthanTableLayout(worksheet, {
    parsedHeaderRowIndex,
    parsedDataStartIndex,
    parsedTableStartCol,
  });
  if (!layout) throw new Error('Could not locate Rajasthan Form A table header row.');

  const { dataStartRow, templateCols } = layout;

  if (headerFormData && typeof headerFormData === 'object') {
    writeStatutoryHeaderFieldsToExcelJsWorksheet(worksheet, {
      headerFormData,
      parsedFormHeader,
      headerRowEnd: Math.max(1, layout.headerRow - 1),
      maxScanCols: 80,
    });
  }

  const templateHeaderLabels = templateCols.map((entry) => entry.label);
  const normalizedHeaders = resolveFormARajasthanExportHeaders(
    Array.isArray(headersToUse) && headersToUse.length >= templateHeaderLabels.length
      ? headersToUse
      : templateHeaderLabels.length > 0
        ? templateHeaderLabels
        : headersToUse
  );
  const exportHelpers = {
    formatStatutoryDateDisplay:
      typeof formatStatutoryDateDisplay === 'function'
        ? formatStatutoryDateDisplay
        : (v) => String(v || '').trim(),
  };
  const rows = enrichFormAGJGujaratRowsForExport(
    mappedData,
    normalizedHeaders,
    employeesOverride,
    exportHelpers
  );

  const tableColMin = Math.min(...templateCols.map(({ col }) => col));
  const tableColMax = Math.max(...templateCols.map(({ col }) => col));
  const clearToRow = Math.max(dataStartRow + rows.length + 12, dataStartRow + 20);
  for (let r = dataStartRow; r <= clearToRow; r += 1) {
    for (let c = tableColMin; c <= tableColMax; c += 1) {
      worksheet.getCell(r, c).value = '';
    }
  }

  rows.forEach((row, idx) => {
    const targetRowNum = dataStartRow + idx;
    const targetRow = worksheet.getRow(targetRowNum);
    if (targetRow) targetRow.height = 18;
    templateCols.forEach(({ col, bucket, label }) => {
      let val = getFormAGJGujaratRowValueForHeader(row, label, idx);
      if ((val == null || val === '') && bucket) {
        for (const [k, v] of Object.entries(row)) {
          if (
            formAGJHeaderAliasBucket(formAGJGujaratHeaderNorm(k)) === bucket &&
            v != null &&
            String(v).trim() !== ''
          ) {
            val = String(v).trim();
            break;
          }
        }
      }
      if ((val == null || val === '') && bucket === 'sno') val = String(idx + 1);
      const cell = worksheet.getCell(targetRowNum, col);
      writeFormARajasthanCellValue(cell, bucket, val);
    });
  });

  if (rows.length > 0) {
    ensureExcelJSDataRowsWithBorders(worksheet, {
      dataStartRow,
      dataRowCount: rows.length,
      colFrom: tableColMin,
      colTo: tableColMax,
      templateRow: dataStartRow,
      templateBodyRows: 1,
    });
  }

  const out = await workbook.xlsx.writeBuffer();
  const fileName =
    formFileName ||
    parsedFormHeader?.title?.replace(/[^a-zA-Z0-9]/g, '_') ||
    'Form_A_RJ_Rajasthan.xlsx';
  return {
    blob: new Blob([out], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    }),
    fileName,
  };
}
