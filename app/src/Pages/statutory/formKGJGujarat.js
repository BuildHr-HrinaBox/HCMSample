import ExcelJS from 'exceljs';
import { ensureExcelJSDataRowsWithBorders } from '../../utils/excelTableBorders';
import { writeStatutoryHeaderFieldsToExcelJsWorksheet } from '../../utils/statutorySiteCompanyHeaders';
import {
  FORM_KGJ_GJ_WEEKLY_HOLIDAY_INTRO,
  isFormKGJHolidayCaptionText,
  isFormKGJTruncatedWeeklyIntroText,
} from '../../utils/statutoryDraftPdf.formKGJ.GJ';

/** Gujarat Form K — worker register (name, designation, weekly holiday, hours of work). */

export const FORM_KGJ_WEEKLY_HOLIDAY_DEFAULT = 'Saturday & Sunday';

export const FORM_KGJ_GJ_CANONICAL_TABLE_HEADERS = [
  'Sr. No. (1)',
  'Name of Worker (2)',
  'Designation (3)',
  'Day of Weekly Holiday (4)',
  'Hours of Work (5)',
];

export function formKGJGujaratHeaderNorm(txt) {
  return String(txt || '')
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

const stripLeadingNumber = (s) =>
  String(s || '')
    .replace(/^\(?\d+\)?\s*/, '')
    .trim();

function normHeaderLabel(h) {
  return formKGJGujaratHeaderNorm(stripLeadingNumber(h))
    .replace(/[^a-z0-9 ]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function unwrapEmployeeRecord(emp) {
  return emp?.Employee || emp?.employee || emp;
}

function readEmployeeFullName(emp) {
  const src = unwrapEmployeeRecord(emp);
  const fn = String(
    src?.FirstName || src?.['FirstName'] || src?.firstName || src?.['First Name'] || ''
  ).trim();
  const ln = String(
    src?.LastName || src?.['LastName'] || src?.lastName || src?.['Last Name'] || ''
  ).trim();
  if (fn && ln) return `${fn} ${ln}`;
  return (
    fn ||
    ln ||
    String(
      src?.DisplayName ||
        src?.['Display Name'] ||
        src?.Employee_Name ||
        src?.['Employee Name'] ||
        ''
    ).trim()
  );
}

function readEmployeeDesignation(emp) {
  const src = unwrapEmployeeRecord(emp);
  return String(
    src?.Designation ||
      src?.['Designation'] ||
      src?.designation ||
      src?.['Designation.displayValue'] ||
      ''
  ).trim();
}

export function isFormKGJSerialHeader(header) {
  const s = normHeaderLabel(header);
  return /^sr\.?\s*no|^s\.?\s*no|serial|^sl\.?\s*no|^no\.?$/.test(s);
}

export function isFormKGJNameHeader(header) {
  const s = normHeaderLabel(header);
  if (!/\bname\b/.test(s)) return false;
  if (/father|spouse|husband|employer|establishment|bank/.test(s)) return false;
  return /\bworker\b|\bworkman\b|\bemployee\b/.test(s) || s === 'name' || /^name\b/.test(s);
}

export function isFormKGJDesignationHeader(header) {
  const s = normHeaderLabel(header);
  return s === 'designation' || /^designation\b/.test(s);
}

export function isFormKGJWeeklyHolidayHeader(header) {
  const s = normHeaderLabel(header);
  return /weekly/.test(s) && /holiday/.test(s);
}

export function isFormKGJHoursOfWorkHeader(header) {
  const s = normHeaderLabel(header);
  return /hours?\s+of\s+work/.test(s) || (/\bhours?\b/.test(s) && /\bwork\b/.test(s));
}

/** Manual entry only — never map from People / payroll. */
export function isFormKGJSkipPeopleAutofillHeader(header) {
  return isFormKGJHoursOfWorkHeader(header);
}

export function headersIndicateFormKGJGujaratTable(tableHeaders) {
  if (!Array.isArray(tableHeaders) || tableHeaders.length < 4) return false;
  const joined = tableHeaders.map((h) => formKGJGujaratHeaderNorm(h)).join('\n');
  const hasWeeklyHoliday = /weekly/.test(joined) && /holiday/.test(joined);
  const hasWorkerName =
    /name\s+of\s+worker/.test(joined) ||
    (/name/.test(joined) && /worker/.test(joined));
  const hasDesignation = /designation/.test(joined);
  const hasHoursOfWork = /hours?\s+of\s+work/.test(joined);
  const looksLikeMaharashtraRegister =
    /register\s+of\s+employment/.test(joined) &&
    (/residential\s+hotel|restaurant|eating\s+house|theatre|public\s+amusement|entertainment/.test(joined) ||
      /hours\s+worked\s+on/.test(joined));
  return hasWeeklyHoliday && hasWorkerName && hasDesignation && hasHoursOfWork && !looksLikeMaharashtraRegister;
}

export function isFormKGJGujaratContext(
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

  const hasGujarat = /gujarat|\b_gj\b|form[\s._-]*k[\s._-]*gj|form_k_gj/.test(parts);
  const hasFormK = /\bform[\s._-]*k\b/.test(parts);

  if (
    /register\s+of\s+employment/.test(parts) &&
    /residential\s+hotel|restaurant|eating\s+house|theatre|public\s+amusement|entertainment/.test(parts)
  ) {
    return false;
  }

  if (hasGujarat && hasFormK) return true;
  if (headersIndicateFormKGJGujaratTable(tableHeaders) && hasGujarat) return true;
  if (headersIndicateFormKGJGujaratTable(tableHeaders) && hasFormK && hasGujarat) return true;

  return false;
}

export function resolveFormKGJGujaratTableHeaders(tableHeaders) {
  const parsed = Array.isArray(tableHeaders)
    ? tableHeaders.map((h) => String(h || '').trim()).filter(Boolean)
    : [];
  return parsed.length > 0 ? parsed : [...FORM_KGJ_GJ_CANONICAL_TABLE_HEADERS];
}

function cellIsEmpty(value) {
  const v = String(value ?? '').trim();
  if (!v) return true;
  const s = v.toLowerCase();
  return /^enter\b/.test(s) || s.includes('enter ');
}

export function applyFormKGJGujaratEmployeeToRow(row, emp, headers, helpers = {}) {
  const hdrs = resolveFormKGJGujaratTableHeaders(headers);
  const out = row && typeof row === 'object' ? { ...row } : {};
  const {
    sanitizeValue = (v) => String(v ?? '').trim(),
    overwrite = true,
    rowIndex = 0,
  } = helpers;

  const shouldSet = (header) => overwrite || cellIsEmpty(out[header]);
  const setCell = (header, value) => {
    if (!header || !shouldSet(header)) return;
    out[header] = sanitizeValue(value ?? '');
  };

  const fullName = readEmployeeFullName(emp);
  const designation = readEmployeeDesignation(emp);
  if (fullName) out.__employeeLookupName = fullName;

  hdrs.forEach((header) => {
    if (isFormKGJSkipPeopleAutofillHeader(header)) {
      setCell(header, '');
      return;
    }
    if (isFormKGJSerialHeader(header)) {
      setCell(header, String(rowIndex + 1));
      return;
    }
    if (isFormKGJNameHeader(header)) {
      setCell(header, fullName);
      return;
    }
    if (isFormKGJDesignationHeader(header)) {
      setCell(header, designation);
      return;
    }
    if (isFormKGJWeeklyHolidayHeader(header)) {
      setCell(header, FORM_KGJ_WEEKLY_HOLIDAY_DEFAULT);
    }
  });

  return out;
}

export function enrichFormKGJGujaratStaticFieldRows(mappedData, headers, { overwrite = true } = {}) {
  if (!Array.isArray(mappedData) || mappedData.length === 0) return 0;
  const hdrs = resolveFormKGJGujaratTableHeaders(headers);
  const holidayHdrs = hdrs.filter(isFormKGJWeeklyHolidayHeader);
  const hoursHdrs = hdrs.filter(isFormKGJHoursOfWorkHeader);
  if (holidayHdrs.length === 0 && hoursHdrs.length === 0) return 0;
  let hits = 0;
  mappedData.forEach((row) => {
    if (!row || typeof row !== 'object') return;
    const holidayKeys = new Set(holidayHdrs);
    const hoursKeys = new Set(hoursHdrs);
    Object.keys(row).forEach((key) => {
      if (String(key).startsWith('__')) return;
      if (isFormKGJWeeklyHolidayHeader(key)) holidayKeys.add(key);
      if (isFormKGJHoursOfWorkHeader(key)) hoursKeys.add(key);
    });
    holidayKeys.forEach((header) => {
      if (!overwrite && !cellIsEmpty(row[header])) return;
      row[header] = FORM_KGJ_WEEKLY_HOLIDAY_DEFAULT;
    });
    hoursKeys.forEach((header) => {
      if (!overwrite && !cellIsEmpty(row[header])) return;
      row[header] = '';
    });
    if (holidayKeys.size > 0 || hoursKeys.size > 0) hits += 1;
  });
  return hits;
}

export function formKGJHeaderAliasBucket(norm) {
  const n = String(norm || '').trim();
  if (!n) return '';
  if (/^sr\.?\s*no|^s\.?\s*no|serial|sl\.?\s*no|^no\.?$/.test(n)) return 'sno';
  if (/\bname\b/.test(n) && (/\bworker\b|\bworkman\b|\bemployee\b/.test(n) || n === 'name' || /^name\b/.test(n))) {
    return 'name';
  }
  if (n === 'designation' || /^designation\b/.test(n)) return 'designation';
  if (/weekly/.test(n) && /holiday/.test(n)) return 'weeklyHoliday';
  if (/hours?\s+of\s+work/.test(n) || (/\bhours?\b/.test(n) && /\bwork\b/.test(n))) return 'hoursOfWork';
  return n;
}

export function getFormKGJGujaratRowValueForHeader(row, header, rowIndex = 0) {
  if (!row || !header) return '';
  const direct = row[header];
  if (direct != null && String(direct).trim() !== '') return String(direct).trim();
  const bucket = formKGJHeaderAliasBucket(normHeaderLabel(header));
  for (const [k, v] of Object.entries(row)) {
    if (
      formKGJHeaderAliasBucket(normHeaderLabel(k)) === bucket &&
      v != null &&
      String(v).trim() !== ''
    ) {
      return String(v).trim();
    }
  }
  if (bucket === 'sno') return String(rowIndex + 1);
  if (bucket === 'weeklyHoliday') return FORM_KGJ_WEEKLY_HOLIDAY_DEFAULT;
  return '';
}

export function rowHasMeaningfulFormKGJGujaratExportData(row, headers) {
  const hdrs = resolveFormKGJGujaratTableHeaders(headers);
  return hdrs.some((header) => {
    const bucket = formKGJHeaderAliasBucket(normHeaderLabel(header));
    if (bucket === 'sno') return false;
    return getFormKGJGujaratRowValueForHeader(row, header) !== '';
  });
}

export function remapFormKGJGujaratRowsToHeaders(rows, sourceHeaders, targetHeaders) {
  const src = Array.isArray(sourceHeaders) ? sourceHeaders : [];
  const tgt = resolveFormKGJGujaratTableHeaders(targetHeaders);
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
        const bucket = formKGJHeaderAliasBucket(normHeaderLabel(targetHeader));
        for (const [k, v] of Object.entries(row)) {
          if (formKGJHeaderAliasBucket(normHeaderLabel(k)) === bucket) {
            val = v;
            break;
          }
        }
      }
      if (
        (val == null || val === '') &&
        formKGJHeaderAliasBucket(normHeaderLabel(targetHeader)) === 'sno'
      ) {
        val = rowIndex + 1;
      }
      out[targetHeader] = val == null ? '' : val;
    });
    return out;
  });
}

export function filterFormKGJGujaratExportRows(rows, headers) {
  const hdrs = resolveFormKGJGujaratTableHeaders(headers);
  return (Array.isArray(rows) ? rows : []).filter((row) =>
    rowHasMeaningfulFormKGJGujaratExportData(row, hdrs)
  );
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
  const t = formKGJGujaratHeaderNorm(text).replace(/[^a-z0-9 ]/g, ' ').replace(/\s+/g, ' ').trim();
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
        const r1 = tl.fullAddress?.row ?? tl.row;
        const c1 = tl.fullAddress?.col ?? tl.col;
        const r2 = br.fullAddress?.row ?? br.row;
        const c2 = br.fullAddress?.col ?? br.col;
        if (r >= r1 && r <= r2 && c >= c1 && c <= c2) {
          topLeft = { r: r1, c: c1 };
          break;
        }
      }
    }
    cache.set(key, topLeft);
    return topLeft;
  };
}

function detectFormKGJGujaratTableLayout(worksheet, hints = {}) {
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
      for (let c = 1; c <= 30; c += 1) {
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

  const templateCols = [];
  for (let c = startCol; c <= startCol + 12; c += 1) {
    const label = getMergedAwareCellText(headerRow, c);
    if (!label) {
      if (templateCols.length >= 5) break;
      continue;
    }
    const bucket = formKGJHeaderAliasBucket(normHeaderLabel(label));
    templateCols.push({ col: c, label, bucket });
    if (templateCols.length >= 8) break;
  }
  if (templateCols.length < 4) return null;

  const dataStartRow =
    hints.parsedDataStartIndex != null && hints.parsedDataStartIndex >= 0
      ? hints.parsedDataStartIndex + 1
      : headerRow + 1;

  return { headerRow, dataStartRow, templateCols, startCol };
}

function rewriteFormKGJGujaratNoticeCaptionCells(worksheet, headerRow) {
  const getMergeTopLeft = buildMergeTopLeftResolver(worksheet);
  const seen = new Set();
  const maxR = Math.max(1, Number(headerRow) - 1);
  for (let r = 1; r <= maxR; r += 1) {
    for (let c = 1; c <= 16; c += 1) {
      const tl = getMergeTopLeft(r, c);
      const key = `${tl.r}:${tl.c}`;
      if (seen.has(key)) continue;
      seen.add(key);
      const cell = worksheet.getCell(tl.r, tl.c);
      const raw = excelCellValueToString(cell.value);
      if (isFormKGJTruncatedWeeklyIntroText(raw)) {
        cell.value = '';
        continue;
      }
      if (isFormKGJHolidayCaptionText(raw)) {
        cell.value = FORM_KGJ_GJ_WEEKLY_HOLIDAY_INTRO;
        cell.alignment = {
          ...(cell.alignment || {}),
          wrapText: true,
          vertical: 'middle',
          horizontal: 'left',
        };
      }
    }
  }
}

export async function buildFormKGJGujaratWorkbookWithTemplateStyles({
  templateArrayBuffer,
  mappedData,
  headersToUse,
  parsedFormHeader,
  formFileName,
  headerFormData,
  parsedHeaderRowIndex,
  parsedDataStartIndex,
  parsedTableStartCol,
}) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(templateArrayBuffer);
  const worksheet = workbook.worksheets[0];
  if (!worksheet) throw new Error('Template worksheet not found.');

  const layout = detectFormKGJGujaratTableLayout(worksheet, {
    parsedHeaderRowIndex,
    parsedDataStartIndex,
    parsedTableStartCol,
  });
  if (!layout) throw new Error('Could not locate Gujarat Form K table header row.');

  const { dataStartRow, templateCols } = layout;

  if (headerFormData && typeof headerFormData === 'object') {
    writeStatutoryHeaderFieldsToExcelJsWorksheet(worksheet, {
      headerFormData,
      parsedFormHeader,
      headerRowEnd: Math.max(1, layout.headerRow - 1),
      maxScanCols: 80,
    });
  }

  rewriteFormKGJGujaratNoticeCaptionCells(worksheet, layout.headerRow);

  const templateHeaderLabels = templateCols.map((entry) => entry.label);
  const normalizedHeaders = resolveFormKGJGujaratTableHeaders(
    Array.isArray(headersToUse) && headersToUse.length >= templateHeaderLabels.length
      ? headersToUse
      : templateHeaderLabels.length > 0
        ? templateHeaderLabels
        : headersToUse
  );
  const rows = filterFormKGJGujaratExportRows(
    remapFormKGJGujaratRowsToHeaders(
      Array.isArray(mappedData) ? mappedData : [],
      normalizedHeaders,
      normalizedHeaders
    ),
    normalizedHeaders
  );

  const tableColMin = Math.min(...templateCols.map(({ col }) => col));
  const tableColMax = Math.max(...templateCols.map(({ col }) => col));
  const clearToRow = Math.max(dataStartRow + rows.length + 10, dataStartRow + 15);
  for (let r = dataStartRow; r <= clearToRow; r += 1) {
    for (let c = tableColMin; c <= tableColMax; c += 1) {
      worksheet.getCell(r, c).value = '';
    }
  }

  rows.forEach((row, idx) => {
    const targetRow = worksheet.getRow(dataStartRow + idx);
    if (targetRow) targetRow.height = 18;
    templateCols.forEach(({ col, bucket, label }) => {
      let val = getFormKGJGujaratRowValueForHeader(row, label, idx);
      if ((val == null || val === '') && bucket) {
        for (const [k, v] of Object.entries(row)) {
          if (
            formKGJHeaderAliasBucket(normHeaderLabel(k)) === bucket &&
            v != null &&
            String(v).trim() !== ''
          ) {
            val = String(v).trim();
            break;
          }
        }
      }
      if ((val == null || val === '') && bucket === 'sno') val = String(idx + 1);
      const cell = worksheet.getCell(dataStartRow + idx, col);
      if (val == null || val === '') {
        cell.value = '';
        return;
      }
      if (bucket === 'sno') {
        const n = Number(String(val).replace(/[,]/g, '').trim());
        cell.value = Number.isFinite(n) ? n : String(val);
      } else {
        cell.value = String(val);
      }
      cell.alignment = { ...(cell.alignment || {}), vertical: 'middle', wrapText: true };
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
    'Form_K_GJ_Gujarat.xlsx';
  return {
    blob: new Blob([out], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    }),
    fileName,
  };
}
