import ExcelJS from 'exceljs';
import { ensureExcelJSDataRowsWithBorders } from '../../utils/excelTableBorders';
import { writeStatutoryHeaderFieldsToExcelJsWorksheet } from '../../utils/statutorySiteCompanyHeaders';
import { sanitizeFormXVIAPMonthValue, sanitizeFormXVIAPNatureValue } from '../../utils/statutoryDraftPdf.formXVI.AP';

/** Excel columns AK / AL (1-based) — leftover overflow boxes on Form XVI AP Muster Roll. */
export const FORM_XVI_AP_COL_AK = 37;
export const FORM_XVI_AP_COL_AL = 38;
export const FORM_XVI_MONTH_HEADER_KEY = 'form_xvi_month';

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

/** When parsed headers omit day columns, synthesize 1…31 for AP Register of Employment grids. */
export function ensureFormXXIIAPDayColumnHeaders(headers) {
  const base = Array.isArray(headers) ? headers.filter(Boolean) : [];
  const dayByNum = new Map();
  base.forEach((header) => {
    const day = resolveFormXXIIAPDayNumberFromHeader(header);
    if (day >= 1 && day <= 31 && !dayByNum.has(day)) dayByNum.set(day, header);
  });
  if (dayByNum.size >= 3) return base;
  const prefix = base.filter((h) => !isFormXXIIAPDayHeaderKey(h));
  const joined = prefix.join(' ').toLowerCase();
  const looksLikeEmploymentRegister =
    /register\s+of\s+employment/.test(joined) ||
    (/time\s+at\s+which\s+employment/.test(joined) && /rest\s+interval/.test(joined)) ||
    (/\bsex\b/.test(joined) && /\bage\b/.test(joined) && /employee/.test(joined) && /dates|attendance/.test(joined));
  if (!looksLikeEmploymentRegister && dayByNum.size === 0) return base;
  for (let day = 1; day <= 31; day += 1) {
    if (!dayByNum.has(day)) dayByNum.set(day, String(day));
  }
  return [...prefix, ...[...dayByNum.entries()].sort((a, b) => a[0] - b[0]).map(([, header]) => header)];
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
  let merged = base;
  if (dayHeaders.length === 0) {
    merged = base;
  } else if (base.some((h) => isFormXXIIAPDayHeaderKey(h))) {
    const missing = dayHeaders.filter((h) => !base.includes(h));
    merged = missing.length > 0 ? [...base, ...missing] : base;
  } else {
    merged = [...prefix, ...dayHeaders];
  }
  return ensureFormXXIIAPDayColumnHeaders(merged);
}

export function isFormXXIIAPAttendanceCode(value) {
  return /^(P|A|WO|H|L|WOP|OD|SL|CL|EL)$/i.test(String(value ?? '').trim());
}

function rowHasFormXXIIAPEmployeeName(row) {
  if (!row || typeof row !== 'object' || Array.isArray(row)) return false;
  for (const [key, val] of Object.entries(row)) {
    if (String(key).startsWith('__')) continue;
    if (/name\s+of\s+the\s+(employee|workman|workmen)/i.test(String(key || ''))) {
      return String(val ?? '').trim().length > 0;
    }
  }
  return false;
}

function defaultFormXXIIAPAttendanceCode(day, year, monthIndex) {
  const dayDate = new Date(year, monthIndex, day);
  if (dayDate.getMonth() !== monthIndex) return '';
  const dow = dayDate.getDay();
  if (dow === 0 || dow === 6) return 'WO';
  return 'A';
}

export function rowHasAnyFormXXIIAPAttendance(row) {
  if (!row || typeof row !== 'object' || Array.isArray(row)) return false;
  for (const [key, val] of Object.entries(row)) {
    if (String(key).startsWith('__')) continue;
    if (resolveFormXXIIAPDayNumberFromHeader(key) >= 1 && isFormXXIIAPAttendanceCode(val)) return true;
  }
  return false;
}

/** Ensure every row has P/A/WO on day keys before Excel export (Zoho + calendar fallback). */
export function normalizeFormXXIIAPRowsForExport(rows, headers, options = {}) {
  const list = Array.isArray(rows) ? rows : [];
  const today = options.today instanceof Date ? options.today : new Date();
  const year = Number.isFinite(options.year) ? options.year : today.getFullYear();
  const monthIndex = Number.isFinite(options.monthIndex) ? options.monthIndex : today.getMonth();
  const fillFallback = options.fillFallback !== false;
  const onlyFillEmpty = options.onlyFillEmpty !== false;
  const isCurrentMonth = year === today.getFullYear() && monthIndex === today.getMonth();

  const resolvedHdrs = ensureFormXXIIAPDayColumnHeaders(
    resolveFormXXIIExportHeaders(headers, list)
  );
  const dayHdrs = listFormXXIIAPDayHeaders(resolvedHdrs);
  if (dayHdrs.length === 0) return list.map((row) => (row && typeof row === 'object' ? { ...row } : row));

  return list.map((row) => {
    if (!row || typeof row !== 'object' || Array.isArray(row)) return row;
    const out = { ...row };
    const hasName = rowHasFormXXIIAPEmployeeName(out);
    const hasAnyAttendance = rowHasAnyFormXXIIAPAttendance(out);

    dayHdrs.forEach(({ header, day }) => {
      const existingOnHeader = readFormXXIIAPCellValue(out, header);
      const existingOnDay = readFormXXIIAPCellValue(out, String(day));
      if (
        onlyFillEmpty &&
        (isFormXXIIAPAttendanceCode(existingOnHeader) || isFormXXIIAPAttendanceCode(existingOnDay))
      ) {
        if (isFormXXIIAPAttendanceCode(existingOnHeader)) {
          out[String(day)] = String(existingOnHeader).trim();
        } else if (isFormXXIIAPAttendanceCode(existingOnDay)) {
          out[header] = String(existingOnDay).trim();
        }
        return;
      }

      let val = isFormXXIIAPAttendanceCode(existingOnHeader)
        ? String(existingOnHeader).trim()
        : isFormXXIIAPAttendanceCode(existingOnDay)
          ? String(existingOnDay).trim()
          : '';
      if (!isFormXXIIAPAttendanceCode(val)) {
        for (const [key, cellVal] of Object.entries(out)) {
          if (String(key).startsWith('__')) continue;
          if (resolveFormXXIIAPDayNumberFromHeader(key) !== day) continue;
          if (isFormXXIIAPAttendanceCode(cellVal)) {
            val = String(cellVal).trim();
            break;
          }
        }
      }
      if (!isFormXXIIAPAttendanceCode(val) && fillFallback && hasName && !hasAnyAttendance) {
        if (!(isCurrentMonth && day > today.getDate())) {
          val = defaultFormXXIIAPAttendanceCode(day, year, monthIndex);
        }
      }
      if (!isFormXXIIAPAttendanceCode(val)) return;
      out[header] = val;
      out[String(day)] = val;
    });
    return out;
  });
}

export function readFormXXIIAPCellValue(row, header) {
  if (!row || typeof row !== 'object' || Array.isArray(row)) return '';
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
  const { dayColumnMap } = detectFormXXIIAPDayColumnMapWithMarker(
    getCell,
    headerRow,
    maxScanRows,
    maxScanCols
  );
  return dayColumnMap;
}

/** Scan header band (incl. row above label row) for 1–31 day markers — Form XVI two-row thead. */
export function detectFormXXIIAPDayColumnMapWithMarker(getCell, headerRow, maxScanRows, maxScanCols) {
  const scanTo = Math.max(maxScanCols, 60);
  const anchor = headerRow > 0 ? headerRow : 1;
  const scanStartRow = Math.max(1, anchor - 4);
  const scanEndRow = Math.min(anchor + 8, maxScanRows);
  let bestMap = new Map();
  let bestHits = 0;
  let markerRow = -1;

  for (let r = scanStartRow; r <= scanEndRow; r += 1) {
    const rowMap = new Map();
    let hits = 0;
    for (let c = 1; c <= scanTo; c += 1) {
      const raw = String(getCell(r, c) || '')
        .trim()
        .replace(/[()]/g, '');
      if (!/^\d{1,2}$/.test(raw)) continue;
      const n = Number(raw);
      if (n >= 1 && n <= 31 && !rowMap.has(n)) {
        rowMap.set(n, c);
        hits += 1;
      }
    }
    if (hits > bestHits) {
      bestHits = hits;
      bestMap = rowMap;
      markerRow = r;
    }
  }

  return { dayColumnMap: bestMap, markerRow, markerHits: bestHits };
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

function excelColLettersToNumber(letters) {
  let n = 0;
  const s = String(letters || '').toUpperCase();
  for (let i = 0; i < s.length; i += 1) {
    n = n * 26 + (s.charCodeAt(i) - 64);
  }
  return n;
}

function unmergeExcelJSIntersections(worksheet, rowFrom, rowTo, colFrom, colTo) {
  if (!worksheet) return;
  const merges = worksheet?.model?.merges;
  if (!Array.isArray(merges) || merges.length === 0) return;
  const r0 = Math.max(1, rowFrom);
  const r1 = Math.max(r0, rowTo);
  const c0 = Math.max(1, colFrom);
  const c1 = Math.max(c0, colTo);
  const toRemove = [];
  for (const range of merges) {
    const parts = String(range || '').split(':');
    if (parts.length !== 2) continue;
    const start = parts[0].match(/^([A-Z]+)(\d+)$/i);
    const end = parts[1].match(/^([A-Z]+)(\d+)$/i);
    if (!start || !end) continue;
    const mr1 = parseInt(start[2], 10);
    const mr2 = parseInt(end[2], 10);
    const mc1 = excelColLettersToNumber(start[1]);
    const mc2 = excelColLettersToNumber(end[1]);
    if (mr2 < r0 || mr1 > r1) continue;
    if (mc2 < c0 || mc1 > c1) continue;
    toRemove.push(range);
  }
  toRemove.forEach((range) => {
    try {
      worksheet.unMergeCells(range);
    } catch (_) {
      try {
        if (typeof worksheet.unmergeCells === 'function') worksheet.unmergeCells(range);
      } catch (__) {
        /* already unmerged */
      }
    }
  });
}

function applyThinBoxBorder(cell) {
  if (!cell) return;
  const box = {
    top: { style: 'thin', color: { argb: 'FF000000' } },
    left: { style: 'thin', color: { argb: 'FF000000' } },
    bottom: { style: 'thin', color: { argb: 'FF000000' } },
    right: { style: 'thin', color: { argb: 'FF000000' } }
  };
  cell.border = box;
  try {
    const prev = cell.style && typeof cell.style === 'object' ? { ...cell.style } : {};
    cell.style = { ...prev, border: box };
  } catch (_) {
    /* border property is enough */
  }
}

function clearExcelJSCellBoxAndData(cell) {
  if (!cell) return;
  cell.value = null;
  cell.border = undefined;
  try {
    const prev = cell.style && typeof cell.style === 'object' ? { ...cell.style } : {};
    delete prev.border;
    cell.style = prev;
  } catch (_) {
    /* ignore */
  }
}

export function clampFormXVIAPHeadersToRemarks(headers) {
  const list = Array.isArray(headers) ? headers.filter((h) => h != null && String(h).trim() !== '') : [];
  const remarksIdx = list.findIndex((h) => /^remarks?$/i.test(String(h || '').trim()));
  if (remarksIdx >= 0) return list.slice(0, remarksIdx + 1);
  return list;
}

export function resolveFormXVIAPLastKeepCol(dayColumnMap, remarksCol, orderedCols = []) {
  let lastDay = 0;
  if (dayColumnMap instanceof Map) {
    dayColumnMap.forEach((col) => {
      const n = Number(col);
      if (Number.isFinite(n)) lastDay = Math.max(lastDay, n);
    });
  }
  const orderedMax = Array.isArray(orderedCols)
    ? orderedCols.reduce((m, c) => Math.max(m, Number(c) || 0), 0)
    : 0;
  let lastKeep = Math.max(lastDay, 1);
  const remarks = Number(remarksCol) || 0;
  if (remarks > 0 && remarks <= lastDay + 2) lastKeep = Math.max(lastKeep, remarks);
  if (orderedMax > 0 && orderedMax <= lastKeep + 2) lastKeep = Math.max(lastKeep, orderedMax);
  // Official Form XVI band is S.No … Remarks (A–AJ). When day 31 is before AK, drop leftover AK/AL boxes.
  if (lastDay > 0 && lastDay < FORM_XVI_AP_COL_AK) {
    lastKeep = Math.min(lastKeep, FORM_XVI_AP_COL_AK - 1);
  }
  return lastKeep;
}

function findFormXVIAPRemarksCol(getCell, headerRow, maxScanRows, maxScanCols) {
  const r0 = Math.max(1, headerRow - 2);
  const r1 = Math.min(headerRow + 3, maxScanRows);
  for (let r = r0; r <= r1; r += 1) {
    for (let c = 1; c <= Math.min(50, maxScanCols); c += 1) {
      if (/^remarks?$/i.test(String(getCell(r, c) || '').trim())) return c;
    }
  }
  return -1;
}

function formatFormXVIAPMonthYearFromAttendance(year, monthIndex) {
  const y = Number(year);
  const m = Number(monthIndex);
  if (!Number.isFinite(y) || !Number.isFinite(m) || m < 0 || m > 11) return '';
  const names = [
    'January',
    'February',
    'March',
    'April',
    'May',
    'June',
    'July',
    'August',
    'September',
    'October',
    'November',
    'December'
  ];
  return `${names[m]} ${y}`;
}

function resolveFormXVIAPMonthYearText(headerFormData, monthYearText, attendanceYear, attendanceMonthIndex) {
  const direct = String(monthYearText || '').trim();
  if (direct) return sanitizeFormXVIAPMonthValue(direct) || direct;
  if (headerFormData && typeof headerFormData === 'object') {
    const keys = [
      FORM_XVI_MONTH_HEADER_KEY,
      'form_b_header_for_the_month_of',
      'form_t_month_year',
      'form_xviii_month_year'
    ];
    for (let i = 0; i < keys.length; i += 1) {
      const v = String(headerFormData[keys[i]] ?? '').trim();
      if (v) return sanitizeFormXVIAPMonthValue(v) || v;
    }
  }
  return formatFormXVIAPMonthYearFromAttendance(attendanceYear, attendanceMonthIndex);
}

function resolveFormXVIAPNatureLocationText(headerFormData) {
  if (!headerFormData || typeof headerFormData !== 'object') return '';
  const keys = [
    'form_xvi_nature_location_work',
    'form_xv_nature_location_work',
    'form_xvii_nature_location_work',
    'form_xviii_nature_location_work',
    'form_xxiii_nature_location_work'
  ];
  for (let i = 0; i < keys.length; i += 1) {
    const v = String(headerFormData[keys[i]] ?? '').trim();
    if (!v) continue;
    return sanitizeFormXVIAPNatureValue(v) || v;
  }
  return '';
}

function isFormXVIAPMonthLabelCell(raw) {
  const s = String(raw || '').replace(/\s+/g, ' ').trim();
  if (!s) return false;
  // Template often has only "For the" above the Dates band.
  return /^for\s+the(\s+month(\s+of)?)?\s*:?\s*/i.test(s);
}

function isFormXVIAPNatureLabelCell(raw) {
  const s = String(raw || '').replace(/\s+/g, ' ').trim();
  if (!s) return false;
  // Must start with the nature label — do not match concatenated header dumps.
  return /^nature\s+and\s+location\s+of\s+work/i.test(s);
}

/** Merge a Form XVI header label across columns and keep the full text on one line. */
function writeFormXVIAPSingleLineHeader(worksheet, row, startCol, endCol, text) {
  if (!worksheet || !text || startCol < 1) return;
  const mergeEnd = Math.max(startCol, Number(endCol) || startCol);
  unmergeExcelJSIntersections(worksheet, row, row, startCol, mergeEnd);
  for (let c = startCol; c <= mergeEnd; c += 1) {
    const cell = worksheet.getCell(row, c);
    if (c === startCol) continue;
    cell.value = null;
  }
  if (mergeEnd > startCol) {
    try {
      worksheet.mergeCells(row, startCol, row, mergeEnd);
    } catch (_) {
      /* template may already merge */
    }
  }
  const cell = worksheet.getCell(row, startCol);
  cell.value = text;
  cell.font = { ...(cell.font || {}), bold: true };
  cell.alignment = {
    ...(cell.alignment || {}),
    wrapText: false,
    shrinkToFit: false,
    vertical: 'middle',
    horizontal: 'left'
  };
  const wsRow = worksheet.getRow(row);
  if (wsRow) {
    wsRow.height = Math.max(Number(wsRow.height) || 0, 20);
  }
}

/** Form XVI AP — make every populated cell bold (headers, meta, attendance, footer). */
export function applyFormXVIAPBoldAllData(worksheet, { rowFrom = 1, rowTo = 40, colFrom = 1, colTo = 45 } = {}) {
  if (!worksheet) return;
  const r0 = Math.max(1, Number(rowFrom) || 1);
  const r1 = Math.max(r0, Number(rowTo) || r0);
  const c0 = Math.max(1, Number(colFrom) || 1);
  const c1 = Math.max(c0, Number(colTo) || c0);
  for (let r = r0; r <= r1; r += 1) {
    for (let c = c0; c <= c1; c += 1) {
      const cell = worksheet.getCell(r, c);
      const raw = excelCellValueToString(cell?.value).trim();
      if (!raw && cell.value == null) continue;
      const prev = cell.font && typeof cell.font === 'object' ? { ...cell.font } : {};
      cell.font = { ...prev, bold: true };
      // Keep rich-text runs bold when present.
      if (cell.value && typeof cell.value === 'object' && Array.isArray(cell.value.richText)) {
        cell.value = {
          richText: cell.value.richText.map((rt) => ({
            ...rt,
            font: { ...(rt?.font || {}), bold: true }
          }))
        };
      }
    }
  }
}

/**
 * Force Nature + For the Month of onto the row above Dates as two full single-line bands.
 * Handles templates that only have a "For the" fragment and/or omit Nature entirely.
 */
function writeFormXVIAPNatureAndMonthSingleLines(
  worksheet,
  {
    datesHeaderRow,
    firstDayCol,
    lastDayCol,
    natureText,
    monthYearText,
    headerScanEnd
  } = {}
) {
  if (!worksheet) return;
  const metaRow = Math.max(1, (Number(datesHeaderRow) || 9) - 1);
  const scanRows = Math.max(metaRow + 1, Number(headerScanEnd) || metaRow + 1, 12);
  const dayStart = Number(firstDayCol) > 0 ? Number(firstDayCol) : 9;
  const dayEnd = Number(lastDayCol) >= dayStart ? Number(lastDayCol) : dayStart + 14;

  let natureRow = 0;
  let natureCol = 0;
  let natureFromCell = '';
  let monthRow = 0;
  let monthCol = 0;
  let monthFromCell = '';

  for (let r = 1; r <= scanRows; r += 1) {
    for (let c = 1; c <= 60; c += 1) {
      const raw = excelCellValueToString(worksheet.getCell(r, c)?.value).trim();
      if (!raw) continue;
      if (!natureCol && isFormXVIAPNatureLabelCell(raw)) {
        natureRow = r;
        natureCol = c;
        natureFromCell = sanitizeFormXVIAPNatureValue(
          raw.replace(/^.*?nature\s+and\s+location\s+of\s+work\s*:?\s*/i, '')
        );
      }
      if (!monthCol && isFormXVIAPMonthLabelCell(raw)) {
        monthRow = r;
        monthCol = c;
        monthFromCell = sanitizeFormXVIAPMonthValue(
          raw.replace(/^.*?for\s+the(\s+month(\s+of)?)?\s*:?\s*/i, '')
        );
      }
    }
  }

  const resolvedNature = String(natureText || natureFromCell || '').trim();
  const resolvedMonth = String(monthYearText || monthFromCell || '').trim();
  // Always place both lines on the row directly above Dates (single-line layout).
  const targetRow = metaRow;

  // Clear leftover "For the" / nature fragments on the target row and nearby meta rows.
  for (let r = Math.max(1, targetRow - 1); r <= targetRow + 1; r += 1) {
    for (let c = 1; c <= Math.max(60, dayEnd + 2); c += 1) {
      const raw = excelCellValueToString(worksheet.getCell(r, c)?.value).trim();
      if (!raw) continue;
      if (isFormXVIAPNatureLabelCell(raw) || isFormXVIAPMonthLabelCell(raw) || /^for\s+the\b/i.test(raw)) {
        worksheet.getCell(r, c).value = null;
      }
    }
  }

  // Left band: Nature and Location of work (full single line).
  if (resolvedNature) {
    const natureStart = 1;
    const natureEnd = Math.max(natureStart + 6, dayStart - 1);
    writeFormXVIAPSingleLineHeader(
      worksheet,
      targetRow,
      natureStart,
      natureEnd,
      `Nature and Location of work : ${resolvedNature}`
    );
  }

  // Right band above Dates: For the Month of (full single line).
  if (resolvedMonth) {
    const monthStart = Math.max(dayStart, (resolvedNature ? dayStart : monthCol) || dayStart);
    const monthEnd = Math.max(monthStart + 12, Math.min(dayEnd, monthStart + 16));
    writeFormXVIAPSingleLineHeader(
      worksheet,
      targetRow,
      monthStart,
      monthEnd,
      `For the Month of : ${resolvedMonth}`
    );
  } else if (monthCol) {
    // Still expand the fragment label even when month value is missing.
    const monthStart = Math.max(dayStart, monthCol);
    writeFormXVIAPSingleLineHeader(
      worksheet,
      targetRow,
      monthStart,
      Math.max(monthStart + 12, Math.min(dayEnd, monthStart + 16)),
      'For the Month of :'
    );
  }
}

function clearFormXVIAPColumnBoxAndData(worksheet, col, rowFrom, rowTo) {
  if (!worksheet || !Number.isFinite(col) || col < 1) return;
  const r0 = Math.max(1, rowFrom);
  const r1 = Math.max(r0, rowTo);
  unmergeExcelJSIntersections(worksheet, r0, r1, col, col);
  for (let r = r0; r <= r1; r += 1) {
    clearExcelJSCellBoxAndData(worksheet.getCell(r, col));
  }
}

function mergeFormXVIAPDatesHeading(worksheet, datesHeaderRow, firstDayCol, lastDayCol) {
  if (!worksheet || firstDayCol < 1 || lastDayCol < firstDayCol) return;
  unmergeExcelJSIntersections(worksheet, datesHeaderRow, datesHeaderRow, firstDayCol, lastDayCol);
  for (let c = firstDayCol; c <= lastDayCol; c += 1) {
    const cell = worksheet.getCell(datesHeaderRow, c);
    cell.value = c === firstDayCol ? 'Dates' : null;
    applyThinBoxBorder(cell);
    cell.alignment = {
      ...(cell.alignment || {}),
      wrapText: false,
      vertical: 'middle',
      horizontal: 'center'
    };
  }
  if (lastDayCol > firstDayCol) {
    try {
      worksheet.mergeCells(datesHeaderRow, firstDayCol, datesHeaderRow, lastDayCol);
    } catch (_) {
      /* already merged */
    }
  }
  const master = worksheet.getCell(datesHeaderRow, firstDayCol);
  master.value = 'Dates';
  master.font = { ...(master.font || {}), bold: true };
  master.alignment = {
    ...(master.alignment || {}),
    wrapText: false,
    vertical: 'middle',
    horizontal: 'center'
  };
}

/**
 * Form XVI AP Muster Roll Excel layout:
 * - merge the "Dates" heading across day 1–31
 * - do not merge attendance data rows
 * - write Nature / For the Month of as full single-line header text
 * - strip leftover AK / AL boxes and data after Remarks
 */
export function applyFormXVIAPMusterLayoutFixes(
  worksheet,
  {
    headerRow,
    markerRow,
    dataStartRow,
    dayColumnMap,
    remarksCol,
    lastKeepCol,
    sourceRowCount,
    monthYearText,
    headerFormData,
    maxScanRows,
    attendanceYear,
    attendanceMonthIndex
  } = {}
) {
  if (!worksheet) return;
  const dayCols = [];
  if (dayColumnMap instanceof Map) {
    dayColumnMap.forEach((col) => {
      const n = Number(col);
      if (Number.isFinite(n) && n >= 1) dayCols.push(n);
    });
  }
  dayCols.sort((a, b) => a - b);
  const firstDayCol = dayCols[0] || 0;
  const lastDayCol = dayCols[dayCols.length - 1] || 0;
  const datesHeaderRow =
    Number(markerRow) > 1 ? Number(markerRow) - 1 : Math.max(1, Number(headerRow) || 9);
  const bodyStart = Math.max(1, Number(dataStartRow) || datesHeaderRow + 2);
  const bodyEnd = Math.max(
    bodyStart,
    bodyStart + Math.max(0, Number(sourceRowCount) || 0) + 8,
    Number(maxScanRows) || bodyStart
  );
  const keepCol = Math.max(1, Number(lastKeepCol) || lastDayCol || FORM_XVI_AP_COL_AK - 1);
  const headerScanEnd = Math.max(datesHeaderRow, Number(headerRow) || 9);

  if (firstDayCol > 0 && lastDayCol >= firstDayCol) {
    mergeFormXVIAPDatesHeading(worksheet, datesHeaderRow, firstDayCol, lastDayCol);
    unmergeExcelJSIntersections(worksheet, bodyStart, bodyStart, firstDayCol, lastDayCol);
    for (let c = firstDayCol; c <= lastDayCol; c += 1) {
      const cell = worksheet.getCell(bodyStart, c);
      cell.alignment = {
        ...(cell.alignment || {}),
        wrapText: false,
        vertical: 'middle',
        horizontal: 'center'
      };
    }
  }

  writeFormXVIAPNatureAndMonthSingleLines(worksheet, {
    datesHeaderRow,
    firstDayCol,
    lastDayCol,
    natureText: resolveFormXVIAPNatureLocationText(headerFormData),
    monthYearText: resolveFormXVIAPMonthYearText(
      headerFormData,
      monthYearText,
      attendanceYear,
      attendanceMonthIndex
    ),
    headerScanEnd
  });

  const extraCols = new Set();
  extraCols.add(FORM_XVI_AP_COL_AK);
  extraCols.add(FORM_XVI_AP_COL_AL);
  const scanTo = Math.max(keepCol + 8, FORM_XVI_AP_COL_AL + 4, 45);
  for (let c = keepCol + 1; c <= scanTo; c += 1) extraCols.add(c);
  if (dayCols.includes(FORM_XVI_AP_COL_AK)) extraCols.delete(FORM_XVI_AP_COL_AK);
  if (dayCols.includes(FORM_XVI_AP_COL_AL) || Number(remarksCol) === FORM_XVI_AP_COL_AL) {
    extraCols.delete(FORM_XVI_AP_COL_AL);
  }
  // Never clear the meta header row cells we just wrote (Nature / Month).
  extraCols.forEach((col) => {
    clearFormXVIAPColumnBoxAndData(worksheet, col, 1, bodyEnd);
    try {
      const excelCol = worksheet.getColumn(col);
      excelCol.hidden = true;
      excelCol.width = 0;
    } catch (_) {
      /* ignore */
    }
  });

  applyFormXVIAPBoldAllData(worksheet, {
    rowFrom: 1,
    rowTo: bodyEnd,
    colFrom: 1,
    colTo: keepCol
  });
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
    ) ||
    workbook.worksheets.find((ws) =>
      /form[\s._-]*xvi(?![a-z])|form[\s._-]*16(?!\d)|muster\s+roll/i.test(String(ws?.name || ''))
    ) ||
    workbook.worksheets[0]
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
  sheetNameHint,
  attendanceYear,
  attendanceMonthIndex,
  formXVIAPMusterFixes = false,
  monthYearText = ''
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
          /name\s+of\s+the\s+(employee|workman|workmen)/.test(t) ||
          /name\s+of\s+workman|name\s+of\s+workmen/.test(t) ||
          /time\s+at\s+which|rest\s+interval|dates|attendance|muster\s+roll/.test(t)
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
    if (
      /name\s+of\s+the\s+(employee|workman|workmen)/.test(t) ||
      /name\s+of\s+workman|name\s+of\s+workmen/.test(t)
    ) {
      nameCol = c;
      break;
    }
  }
  if (nameCol < 1) {
    for (let r = Math.max(1, headerRow - 5); r <= Math.min(headerRow + 2, maxScanRows); r += 1) {
      for (let c = 1; c <= maxScanCols; c += 1) {
        const t = normalize(getCell(r, c));
        if (
          /name\s+of\s+the\s+(employee|workman|workmen)/.test(t) ||
          /name\s+of\s+workman|name\s+of\s+workmen/.test(t)
        ) {
          nameCol = c;
          headerRow = r;
          break;
        }
      }
      if (nameCol > 0) break;
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
  const { dayColumnMap, markerRow: detectedMarkerRow, markerHits } =
    detectFormXXIIAPDayColumnMapWithMarker(getCell, headerRow, maxScanRows, maxScanCols);
  let markerCols = [...dayColumnMap.entries()]
    .sort((a, b) => a[0] - b[0])
    .map(([, col]) => col);
  if (markerCols.length < 2) {
    for (let r = headerRow; r <= Math.min(headerRow + 8, maxScanRows); r += 1) {
      const cols = [];
      for (let c = startCol; c <= maxScanCols; c += 1) {
        const t = String(getCell(r, c) || '').trim();
        if (/^\d{1,2}$/.test(t) && Number(t) >= 1 && Number(t) <= 31) cols.push(c);
      }
      if (cols.length > markerCols.length) {
        markerCols = cols;
        markerRow = r;
      }
    }
  }
  if (markerHits >= 2 && detectedMarkerRow > 0) {
    markerRow = detectedMarkerRow;
    markerCols = [...dayColumnMap.entries()]
      .sort((a, b) => a[0] - b[0])
      .map(([, col]) => col);
  } else if (markerCols.length >= 2) {
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

  const rawPrimary =
    Array.isArray(mappedData) && mappedData.length > 0
      ? mappedData
      : Array.isArray(mappedRowMatrix)
        ? mappedRowMatrix
        : [];
  const exportYear = Number.isFinite(attendanceYear) ? attendanceYear : new Date().getFullYear();
  const exportMonthIndex = Number.isFinite(attendanceMonthIndex)
    ? attendanceMonthIndex
    : new Date().getMonth();
  const headersForExport = formXVIAPMusterFixes
    ? clampFormXVIAPHeadersToRemarks(headersToUse)
    : headersToUse;
  const sourcePrimary = normalizeFormXXIIAPRowsForExport(rawPrimary, headersForExport, {
    year: exportYear,
    monthIndex: exportMonthIndex,
    fillFallback: !rawPrimary.some((row) => rowHasAnyFormXXIIAPAttendance(row)),
    onlyFillEmpty: true
  });
  const sourceHeaders = resolveFormXXIIExportHeaders(headersForExport, sourcePrimary);
  const dayStartIdx = sourceHeaders.findIndex((h) => isFormXXIIAPDayHeaderKey(h));
  const prefixCount = dayStartIdx >= 0 ? dayStartIdx : sourceHeaders.length;
  const firstDayCol = markerCols.length > 0 ? markerCols[0] : -1;

  const buildOrderedCols = () => {
    if (markerCols.length >= 2 && dayStartIdx >= 0) {
      const cols = [];
      const templatePrefixSlots =
        firstDayCol > startCol ? firstDayCol - startCol : prefixCount;
      for (let j = 0; j < prefixCount; j += 1) {
        cols.push(startCol + Math.min(j, Math.max(templatePrefixSlots, 1) - 1));
      }
      const dayCount = Math.max(sourceHeaders.length - dayStartIdx, markerCols.length);
      for (let d = 0; d < dayCount; d += 1) {
        cols.push(markerCols[d] ?? firstDayCol + d);
      }
      return cols;
    }
    if (markerCols.length >= 2) {
      const cols = [];
      const templatePrefixSlots =
        firstDayCol > startCol ? firstDayCol - startCol : prefixCount;
      for (let j = 0; j < prefixCount; j += 1) {
        cols.push(startCol + Math.min(j, Math.max(templatePrefixSlots, 1) - 1));
      }
      for (let d = 0; d < markerCols.length; d += 1) cols.push(markerCols[d]);
      return cols;
    }
    if (firstDayCol > startCol && dayStartIdx >= 0) {
      const cols = [];
      const templatePrefixSlots = firstDayCol - startCol;
      for (let j = 0; j < prefixCount; j += 1) {
        cols.push(startCol + Math.min(j, templatePrefixSlots - 1));
      }
      const dayCount = sourceHeaders.length - dayStartIdx;
      for (let d = 0; d < dayCount; d += 1) {
        cols.push(firstDayCol + d);
      }
      return cols;
    }
    return Array.from({ length: Math.max(12, sourceHeaders.length) }, (_, i) => startCol + i);
  };
  const orderedCols = buildOrderedCols();
  const hdrs = sourceHeaders;

  const getDayValueByNumber = (rowObj, dayNum) => {
    if (!rowObj || !dayNum) return '';
    if (Array.isArray(rowObj)) {
      if (dayStartIdx >= 0) {
        const idx = dayStartIdx + (dayNum - 1);
        const picked = pickExportCell(rowObj[idx]);
        if (picked !== '') return picked;
      }
      return '';
    }
    if (typeof rowObj !== 'object') return '';
    const directByDayHeader = pickExportCell(readFormXXIIAPCellValue(rowObj, String(dayNum)));
    if (directByDayHeader !== '') return directByDayHeader;
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
    const dayNum = resolveFormXXIIAPDayNumberFromHeader(header);
    if (dayNum >= 1 && dayNum <= 31) {
      const byDay = getDayValueByNumber(rowObj, dayNum);
      if (byDay !== '') return byDay;
    }
    const picked = pickExportCell(readFormXXIIAPCellValue(rowObj, header));
    if (picked !== '') {
      if (dayNum === 0 && isFormXXIIAPAttendanceCode(picked)) return '';
      return picked;
    }
    if (headerIndex != null && hdrs[headerIndex]) {
      const alt = pickExportCell(readFormXXIIAPCellValue(rowObj, hdrs[headerIndex]));
      if (alt !== '') {
        const altDay = resolveFormXXIIAPDayNumberFromHeader(hdrs[headerIndex]);
        if (altDay === 0 && isFormXXIIAPAttendanceCode(alt)) return '';
        return alt;
      }
    }
    return '';
  };

  const toOrderedValues = (row, rowIndex) => {
    if (Array.isArray(row)) {
      const out = [...row];
      const colCount = Math.max(orderedCols.length, hdrs.length, out.length);
      return Array.from({ length: colCount }, (_, idx) => pickExportCell(out[idx] ?? ''));
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

  const remarksCol = formXVIAPMusterFixes
    ? findFormXVIAPRemarksCol(getCell, headerRow, maxScanRows, maxScanCols)
    : -1;
  const lastKeepColForWrite = formXVIAPMusterFixes
    ? resolveFormXVIAPLastKeepCol(dayColumnMap, remarksCol, orderedCols)
    : 0;

  const headerValues = headerFormData && typeof headerFormData === 'object' ? { ...headerFormData } : {};
  if (formXVIAPMusterFixes) {
    const monthText = resolveFormXVIAPMonthYearText(
      headerValues,
      monthYearText,
      attendanceYear,
      attendanceMonthIndex
    );
    if (monthText) headerValues[FORM_XVI_MONTH_HEADER_KEY] = monthText;
  }
  writeStatutoryHeaderFieldsToExcelJsWorksheet(worksheet, {
    headerFormData: headerValues,
    parsedFormHeader,
    headerRowEnd: headerRow,
    maxScanRows,
    maxScanCols
  });

  const filteredRows = sourcePrimary.filter((row) => rowLooksMeaningful(row));
  const sourceRows = filteredRows.map((row, idx) => toOrderedValues(row, idx));

  const writeDayAttendanceGrid = (rowObj, excelRowIdx) => {
    if (!rowObj) return;
    for (let day = 1; day <= 31; day += 1) {
      let col = dayColumnMap.get(day);
      if (!col && dayStartIdx >= 0) {
        const idx = prefixCount + (day - 1);
        if (idx < orderedCols.length) col = orderedCols[idx];
      }
      if (!col) continue;
      if (formXVIAPMusterFixes && lastKeepColForWrite > 0 && col > lastKeepColForWrite) continue;
      const val = getDayValueByNumber(rowObj, day);
      if (val == null || val === '') continue;
      const dayCell = worksheet.getCell(excelRowIdx, col);
      dayCell.value = String(val);
      dayCell.font = { ...(dayCell.font || {}), bold: formXVIAPMusterFixes ? true : dayCell.font?.bold };
      dayCell.alignment = {
        ...(dayCell.alignment || {}),
        wrapText: false,
        vertical: 'middle',
        horizontal: 'center'
      };
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
  if (formXVIAPMusterFixes && markerCols.length >= 2) {
    const dayFrom = markerCols[0];
    const dayTo = markerCols[markerCols.length - 1];
    unmergeExcelJSIntersections(
      worksheet,
      dataStartRow,
      dataStartRow + Math.max(sourceRows.length, 1) + 2,
      startCol,
      Math.max(dayTo, lastKeepColForWrite || dayTo)
    );
  }
  for (let r = clearFromRow; r <= clearToRow; r += 1) {
    if (isNumericMarkerTemplateRow(r) || isHeaderBandRow(r)) continue;
    for (let j = 0; j < orderedCols.length; j += 1) {
      const col = orderedCols[j];
      if (formXVIAPMusterFixes && lastKeepColForWrite > 0 && col > lastKeepColForWrite) continue;
      worksheet.getCell(r, col).value = '';
    }
  }

  for (let i = 0; i < sourceRows.length; i += 1) {
    const row = sourceRows[i] || [];
    const rowObj = filteredRows[i];
    const excelRowIdx = dataStartRow + i;
    if (snoCol > 0) {
      const snoCell = worksheet.getCell(excelRowIdx, snoCol);
      snoCell.value = i + 1;
      if (formXVIAPMusterFixes) snoCell.font = { ...(snoCell.font || {}), bold: true };
    }
    for (let j = 0; j < orderedCols.length; j += 1) {
      const col = orderedCols[j];
      if (formXVIAPMusterFixes && lastKeepColForWrite > 0 && col > lastKeepColForWrite) continue;
      const value = row[j];
      if (value == null || value === '') continue;
      const hdr = hdrs[j] || '';
      const cell = worksheet.getCell(excelRowIdx, col);
      const strVal = String(value).trim();
      const isAttendanceCode = /^(P|A|WO|H|L|WOP|OD|SL|CL|EL)$/i.test(strVal);
      if (
        typeof value === 'number' ||
        (typeof value === 'string' &&
          /^-?\d+(\.\d+)?$/.test(strVal) &&
          j > 0 &&
          !isAttendanceCode &&
          !isFormXXIIAPDayHeaderKey(hdr))
      ) {
        if (/^age$/i.test(String(hdr).trim()) || resolveFormXXIIAPDayNumberFromHeader(hdr) === 0) {
          cell.value = Number(value);
        } else {
          cell.value = String(value);
        }
      } else {
        cell.value = String(value);
      }
      if (formXVIAPMusterFixes) {
        cell.font = { ...(cell.font || {}), bold: true };
      }
    }
    writeDayAttendanceGrid(rowObj, excelRowIdx);
  }

  if (sourceRows.length > 0) {
    const tableColMin = orderedCols.length > 0 ? Math.min(...orderedCols) : startCol;
    let tableColMax = orderedCols.length > 0 ? Math.max(...orderedCols) : startCol + 11;
    if (formXVIAPMusterFixes && lastKeepColForWrite > 0) {
      tableColMax = Math.min(tableColMax, lastKeepColForWrite);
    }
    ensureExcelJSDataRowsWithBorders(worksheet, {
      dataStartRow,
      dataRowCount: sourceRows.length,
      colFrom: tableColMin,
      colTo: tableColMax,
      templateRow: dataStartRow,
      templateBodyRows: 1
    });
  }

  if (formXVIAPMusterFixes) {
    applyFormXVIAPMusterLayoutFixes(worksheet, {
      headerRow,
      markerRow,
      dataStartRow,
      dayColumnMap,
      remarksCol,
      lastKeepCol: lastKeepColForWrite,
      sourceRowCount: sourceRows.length,
      monthYearText,
      headerFormData: headerValues,
      maxScanRows,
      attendanceYear,
      attendanceMonthIndex
    });
  }

  const out = await workbook.xlsx.writeBuffer();
  const fileName =
    formFileName ||
    parsedFormHeader?.title?.replace(/[^a-zA-Z0-9]/g, '_') ||
    `Form_XXII_AP_${Date.now()}.xlsx`;
  const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  return { blob, fileName };
}

/** Form XVI AP Muster Roll — day-grid export plus Dates-row / month / AK-AL cleanup. */
export async function buildFormXVIAPMusterWorkbookWithTemplateStyles(opts = {}) {
  return buildFormXXIIAPWorkbookWithTemplateStyles({
    ...opts,
    formXVIAPMusterFixes: true
  });
}
