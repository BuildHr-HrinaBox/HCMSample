import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import {
  flattenPayrollEarningColumns,
  readPayrollScalar,
} from '../../utils/payrollEarnings';
import { ensureExcelJSDataRowsWithBorders } from '../../utils/excelTableBorders';
import {
  enrichEstablishmentPrincipalEmployerHeaderFields,
  excelCellValueToString,
  writeStatutoryHeaderFieldsToExcelJsWorksheet,
} from '../../utils/statutorySiteCompanyHeaders';
import { normalizeAttendanceDateKey } from '../../utils/attendanceApi';

/** Rajasthan Form D — Format of Attendance Register (Place of work, Date IN/OUT, days 1–31). */

/** One grid/export row per employee; day cells hold both punch times as "IN / OUT". */
export const FORM_D_RJ_ROWS_PER_EMPLOYEE = 1;
export const FORM_D_RJ_ROW_TYPE_IN = 'IN';
export const FORM_D_RJ_ROW_TYPE_OUT = 'OUT';
export const FORM_D_RJ_ROW_TYPE_COMBINED = 'IN / OUT';

const MONTH_NAMES_DRJ = [
  'January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December',
];

export function formDRajasthanHeaderNorm(txt) {
  return String(txt || '')
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

const stripLeadingNumber = (s) =>
  String(s || '')
    .replace(/^\(?\d+\)?[\.\)]?\s*/i, '')
    .trim();

function normHeaderLabel(h) {
  return formDRajasthanHeaderNorm(stripLeadingNumber(h))
    .replace(/[^a-z0-9 ]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function getSheetMergedCellText(rows, merges, r, c) {
  const direct = String((rows[r] || [])[c] ?? '').trim();
  if (direct) return direct;
  for (const m of merges || []) {
    if (!m?.s || !m?.e) continue;
    if (r >= m.s.r && r <= m.e.r && c >= m.s.c && c <= m.e.c) {
      return String((rows[m.s.r] || [])[m.s.c] ?? '').trim();
    }
  }
  return '';
}

export function headersIndicateFormDRajasthanAttendanceTable(tableHeaders) {
  if (!Array.isArray(tableHeaders) || tableHeaders.length < 5) return false;
  const joined = tableHeaders.map((h) => normHeaderLabel(h)).join('\n');
  const dayCount = tableHeaders.filter((h) => /^\d{1,2}$/.test(String(h || '').trim())).length;
  const hasPlace = /place\s+of\s+work/.test(joined);
  const hasSummary = /summary.*days|no\.?\s*of\s+days/.test(joined);
  const hasDate = /\bdate\b/.test(joined);
  const hasName = /\bname\b/.test(joined);
  return (hasPlace || hasSummary) && dayCount >= 5 && (hasDate || hasName);
}

export function headersIndicateFormDRajasthanPeriodRowMisparse(tableHeaders) {
  if (!Array.isArray(tableHeaders) || tableHeaders.length < 2) return false;
  const periodLike = tableHeaders.filter((header) => {
    const s = formDRajasthanHeaderNorm(header);
    return /for\s+the\s+period\s+from/.test(s) || /^for\s+the\s+period\b/.test(s);
  }).length;
  return periodLike >= Math.max(2, Math.ceil(tableHeaders.length * 0.5));
}

export function isFormDRajasthanContext(
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

  if (/gujarat|\b_gj\b|form[\s._-]*d[\s._-]*gj|form_d_gj/.test(parts)) return false;

  const hasRajasthan =
    /rajasthan|\b_rj\b|form[\s._-]*d[\s._-]*rj|form_d_rj/.test(parts);
  const hasFormD = /\bform[\s._-]*d\b/.test(parts);

  if (hasRajasthan && hasFormD) return true;
  if (hasRajasthan && /attendance\s+register|format\s+of\s+attendance/.test(parts)) return true;
  if (hasRajasthan && headersIndicateFormDRajasthanAttendanceTable(tableHeaders)) return true;
  if (
    hasRajasthan &&
    hasFormD &&
    headersIndicateFormDRajasthanPeriodRowMisparse(tableHeaders)
  ) {
    return true;
  }
  return false;
}

export function isFormDRajasthanUsingWrongTable(tableHeaders) {
  return (
    headersIndicateFormDRajasthanPeriodRowMisparse(tableHeaders) ||
    !headersIndicateFormDRajasthanAttendanceTable(tableHeaders)
  );
}

export function isFormDRajasthanAttendanceTableLayoutFormHeader(formHeader) {
  return !!formHeader?.formDRJAttendanceTableLayout;
}

export function listFormDRajasthanDayColumnHeaders(tableHeaders) {
  if (!Array.isArray(tableHeaders)) return [];
  return tableHeaders
    .map((header) => {
      const t = String(header || '').trim();
      const direct = t.match(/^(\d{1,2})$/);
      if (direct) {
        const day = Number(direct[1]);
        if (day >= 1 && day <= 31) return { header, day };
      }
      const suffixed = t.match(/(?:^|_)(\d{1,2})$/);
      if (suffixed) {
        const day = Number(suffixed[1]);
        if (day >= 1 && day <= 31) return { header, day };
      }
      return null;
    })
    .filter(Boolean);
}

export function buildFormDRajasthanCanonicalHeaders(daysInMonth = 31) {
  const days = Math.min(Math.max(Number(daysInMonth) || 31, 28), 31);
  const headers = [
    'Sr. No. in Employee / Workman / Worker Register',
    'Name',
    'Place of work*',
    'Date',
  ];
  for (let d = 1; d <= days; d += 1) headers.push(String(d));
  headers.push('Summary No. of Days', 'Remarks No. of hours', 'Signature of Register Keeper*');
  return headers;
}

export function buildFormDRJSubColumnsFromHeaders(headers) {
  return { Date: [FORM_D_RJ_ROW_TYPE_COMBINED] };
}

/** Combine punch times for a single day cell, e.g. "9.00 / 17.00". */
export function formatFormDRJInOutCell(inTime, outTime) {
  const a = String(inTime || '').trim();
  const b = String(outTime || '').trim();
  if (a && b) return `${a} / ${b}`;
  return a || b || '';
}

/** Split a day cell into check-in / check-out parts. */
export function parseFormDRJInOutCell(text) {
  const s = String(text || '').trim();
  if (!s) return { inTime: '', outTime: '' };
  const parts = s.split(/\s*\/\s*/);
  if (parts.length >= 2) {
    return {
      inTime: String(parts[0] || '').trim(),
      outTime: parts
        .slice(1)
        .map((p) => String(p || '').trim())
        .filter(Boolean)
        .join(' / '),
    };
  }
  return { inTime: s, outTime: '' };
}

function formatFormDRJDayCellDisplay(text) {
  const s = String(text || '').trim();
  if (!s) return '';
  if (/\//.test(s)) {
    const { inTime, outTime } = parseFormDRJInOutCell(s);
    return formatFormDRJInOutCell(
      formatFormDRJAttendanceTime(inTime) || inTime,
      formatFormDRJAttendanceTime(outTime) || outTime
    );
  }
  return formatFormDRJAttendanceTime(s) || s;
}

function scoreFormDRajasthanSheet(sheetName, sheetText, hintsBlob) {
  const sheetLower = String(sheetName || '').toLowerCase().replace(/\s+/g, ' ').trim();
  const sheetBlob = `${sheetLower} ${sheetText}`;
  let score = 0;
  if (/\bform[\s._-]*d\b/.test(sheetLower)) score += 200;
  if (/attendance\s+register|format\s+of\s+attendance/.test(sheetText)) score += 180;
  if (/place\s+of\s+work/.test(sheetText)) score += 90;
  if (/summary\s+no\.?\s*of\s+days/.test(sheetText)) score += 70;
  if (/remarks.*hours|no\.?\s*of\s+hours/.test(sheetText)) score += 40;
  if (/register\s+keeper/.test(sheetText)) score += 30;
  if (/relay\s+or\s+set/.test(sheetText) && !/place\s+of\s+work/.test(sheetText)) score -= 80;
  if (/rajasthan|\b_rj\b|form_d_rj/.test(hintsBlob) && /rajasthan|\b_rj\b/.test(sheetBlob)) score += 20;
  return score;
}

function buildSheetTextBlob(workbook, sheetName) {
  const ws = workbook?.Sheets?.[sheetName];
  if (!ws) return '';
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  return rows
    .slice(0, 45)
    .map((row) => (Array.isArray(row) ? row : []).map((c) => String(c || '')).join(' '))
    .join(' ')
    .toLowerCase();
}

export function resolveFormDRajasthanWorkbookSheetName(workbook, hints = {}) {
  const names = Array.isArray(workbook?.SheetNames) ? workbook.SheetNames : [];
  if (!names.length) return null;
  if (names.length === 1) return names[0];

  const blob = [
    hints.fileName,
    hints.formFileName,
    hints.formName,
    hints.item?.formName,
    hints.formHeaderTitle,
    hints.formHeader?.title,
  ]
    .filter(Boolean)
    .join(' ')
    .toLowerCase();

  let best = hints.preferredSheetName && names.includes(hints.preferredSheetName)
    ? hints.preferredSheetName
    : names[0];
  let bestScore = scoreFormDRajasthanSheet(best, buildSheetTextBlob(workbook, best), blob);
  for (const name of names) {
    const score = scoreFormDRajasthanSheet(name, buildSheetTextBlob(workbook, name), blob);
    if (score > bestScore) {
      bestScore = score;
      best = name;
    }
  }
  return bestScore > 0 ? best : hints.preferredSheetName || names[0];
}

export function resolveFormDRajasthanTableLayout(workbook, hints = {}) {
  const names = Array.isArray(workbook?.SheetNames) ? workbook.SheetNames : [];
  if (!names.length) return null;

  const sheetName = resolveFormDRajasthanWorkbookSheetName(workbook, hints);
  if (!sheetName) return null;

  const ws = workbook?.Sheets?.[sheetName];
  if (!ws) return null;

  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  const merges = ws['!merges'] || [];
  const maxScan = Math.min(rows.length, 90);
  let best = null;
  let bestScore = -1;

  for (let r = 0; r < maxScan; r += 1) {
    const cells = [];
    const maxCols = Math.max((rows[r] || []).length, (rows[r + 1] || []).length, 40);
    for (let c = 0; c < maxCols; c += 1) {
      const text = getSheetMergedCellText(rows, merges, r, c);
      if (text) cells.push({ c, text, lower: text.toLowerCase() });
    }
    if (cells.length < 4) continue;

    const nextRowCells = [];
    for (let c = 0; c < maxCols; c += 1) {
      const text = getSheetMergedCellText(rows, merges, r + 1, c);
      if (text) nextRowCells.push(text.toLowerCase());
    }
    const rowText = cells.map((x) => x.lower).join(' ');
    const combinedText = `${rowText} ${nextRowCells.join(' ')}`.trim();

    let score = 0;
    if (/place\s+of\s+work/.test(combinedText)) score += 90;
    if (/summary\s+no\.?\s*of\s+days/.test(combinedText)) score += 70;
    if (/remarks.*hours|no\.?\s*of\s+hours/.test(combinedText)) score += 40;
    if (/register\s+keeper|signature/.test(combinedText)) score += 30;
    if (/\bdate\b/.test(combinedText)) score += 25;
    if (/\bname\b/.test(combinedText) && !/establishment|employer|owner/.test(combinedText)) score += 20;
    if (/sr\.?\s*no|serial|sl\.?\s*no/.test(combinedText)) score += 15;

    const dayNums = cells.filter(({ text }) => /^\d{1,2}$/.test(String(text).trim())).length;
    if (dayNums >= 5) score += 60 + dayNums;
    if (/for\s+the\s+period\s+from/.test(combinedText) && !/place\s+of\s+work/.test(combinedText)) {
      score -= 200;
    }
    if (/relay\s+or\s+set/.test(combinedText) && !/place\s+of\s+work/.test(combinedText)) score -= 100;

    score += cells.length;
    if (score < 20) continue;

    let startCol = cells[0].c;
    for (const cell of cells) {
      if (/^s\.?\s*no\.?$|^sl\.?\s*no\.?$|serial/i.test(cell.lower)) {
        startCol = cell.c;
        break;
      }
    }

    const headers = [];
    for (let c = startCol; c < maxCols; c += 1) {
      let h = getSheetMergedCellText(rows, merges, r, c);
      if (!h) h = getSheetMergedCellText(rows, merges, r + 1, c);
      h = String(h || '').replace(/\s+/g, ' ').trim();
      if (!h) {
        if (headers.length >= 8) break;
        continue;
      }
      if (/^\d+$/.test(h) && headers.length > 0 && headers.some((x) => /^\d{1,2}$/.test(x))) {
        headers.push(h);
        continue;
      }
      if (/^\d+$/.test(h) && headers.length === 0) continue;
      if (/^in$|^out$/i.test(h)) continue;
      headers.push(h);
    }

    const dayHeaders = headers.filter((h) => /^\d{1,2}$/.test(String(h).trim()));
    if (headers.length < 6 || dayHeaders.length < 5) continue;

    let dataStartIndex = r + 1;
    const probe = rows[dataStartIndex] || [];
    const probeText = probe.map((x) => String(x || '').toLowerCase()).join(' ');
    if (/^in$|^out$/i.test(probeText.trim()) || /\bin\b/.test(probeText) && /\bout\b/.test(probeText)) {
      dataStartIndex = r + 3;
    } else if (probe.some((x) => /^in$|^out$/i.test(String(x || '').trim()))) {
      dataStartIndex = r + 2;
    }

    const candidate = {
      sheetName,
      headerRowIndex: r,
      dataStartIndex,
      tableStartCol: startCol,
      headers,
    };
    if (score > bestScore) {
      bestScore = score;
      best = candidate;
    }
  }

  return best;
}

export function repairFormDRajasthanTableHeadersFromWorkbook(workbook, hints = {}) {
  if (!workbook) return null;
  return resolveFormDRajasthanTableLayout(workbook, {
    preferredSheetName: hints.preferredSheetName || hints.sheetName || '',
    fileName: hints.fileName || '',
    formFileName: hints.formFileName || '',
    formName: hints.formName || '',
    item: hints.item || null,
    formHeader: hints.formHeader || hints.parsedFormHeader || null,
    formHeaderTitle: hints.formHeaderTitle || hints.formHeader?.title || '',
  });
}

export function resolveFormDRajasthanTableHeaders(tableHeaders) {
  const parsed = Array.isArray(tableHeaders)
    ? tableHeaders.map((h) => String(h || '').trim()).filter(Boolean)
    : [];
  if (parsed.length >= 6 && headersIndicateFormDRajasthanAttendanceTable(parsed)) return parsed;
  const dayCount = listFormDRajasthanDayColumnHeaders(parsed).length;
  if (dayCount >= 5) return parsed;
  return buildFormDRajasthanCanonicalHeaders(31);
}

function unwrapEmployeeRecord(emp) {
  return emp?.Employee || emp?.employee || emp;
}

function pickEmployeeValue(emp, keys) {
  const src = unwrapEmployeeRecord(emp);
  if (!src || typeof src !== 'object') return '';
  for (let i = 0; i < keys.length; i += 1) {
    const raw = src[keys[i]];
    if (raw == null || raw === '') continue;
    if (typeof raw === 'object') {
      const nested = String(raw.displayValue ?? raw.name ?? raw.Name ?? '').trim();
      if (nested) return nested;
      continue;
    }
    const text = String(raw).trim();
    if (text) return text;
  }
  return '';
}

export function readFormDRJEmployeeFullName(emp) {
  const fn = pickEmployeeValue(emp, ['FirstName', 'First Name', 'firstName']);
  const ln = pickEmployeeValue(emp, ['LastName', 'Last Name', 'lastName']);
  if (fn && ln) return `${fn} ${ln}`;
  return fn || ln || '';
}

export function readFormDRJDesignation(emp) {
  return pickEmployeeValue(emp, [
    'Designation',
    'designation',
    'Designation.displayValue',
    'Department',
    'department',
  ]);
}

export function isFormDRJNameHeader(h) {
  const s = normHeaderLabel(h);
  if (!/\bname\b/.test(s)) return false;
  if (/father|spouse|husband|employer|establishment|owner|register\s+keeper|signature/.test(s)) {
    return false;
  }
  return s === 'name' || /^name\b/.test(s);
}

export function isFormDRJPlaceOfWorkHeader(h) {
  const s = normHeaderLabel(h);
  return /place\s+of\s+work/.test(s);
}

export function isFormDRJDateHeader(h) {
  const s = normHeaderLabel(h);
  return s === 'date' || /^date\b/.test(s);
}

export function isFormDRJSummaryDaysHeader(h) {
  const s = normHeaderLabel(h);
  return /summary.*days|no\.?\s*of\s+days/.test(s) && !/remarks/.test(s);
}

export function isFormDRJRemarksHoursHeader(h) {
  const s = normHeaderLabel(h);
  return /remarks.*hours|no\.?\s*of\s+hours/.test(s);
}

export function isFormDRJSerialHeader(h) {
  const s = normHeaderLabel(h);
  return (
    /^sr\.?\s*no|^s\.?\s*no|serial|sl\.?\s*no/.test(s) ||
    /sl\.?\s*number.*employee|serial.*employee\s*register|number.*employee\s*register/.test(s)
  );
}

export function isFormDRJPeopleFieldHeader(h) {
  return (
    isFormDRJSerialHeader(h) ||
    isFormDRJNameHeader(h) ||
    isFormDRJPlaceOfWorkHeader(h) ||
    isFormDRJDateHeader(h)
  );
}

function stripAttendanceDateFromTime(raw) {
  const s = String(raw || '').trim();
  if (!s || s === '-') return '';
  const zohoDt = s.match(/\b\d{1,2}-[A-Za-z]{3}-\d{4}\s+(.+)$/);
  if (zohoDt) return zohoDt[1].trim();
  const isoPrefix = s.match(/^\d{4}-\d{2}-\d{2}[T\s]+(.+)$/);
  if (isoPrefix) return isoPrefix[1].trim();
  const slashDate = s.match(/^\d{1,2}[\/-]\d{1,2}[\/-]\d{2,4}\s+(.+)$/);
  if (slashDate) return slashDate[1].trim();
  return s;
}

function parseTimeToMinutes(raw) {
  const s = stripAttendanceDateFromTime(raw);
  if (!s || s === '-') return null;
  const colon = s.match(/^(\d{1,2}):(\d{1,2})(?::\d{1,2})?\s*(am|pm)?$/i);
  if (colon) {
    let hh = Number(colon[1]);
    const mm = Number(colon[2]);
    const ap = (colon[3] || '').toLowerCase();
    if (ap === 'pm' && hh < 12) hh += 12;
    if (ap === 'am' && hh === 12) hh = 0;
    if (Number.isFinite(hh) && Number.isFinite(mm)) return hh * 60 + mm;
  }
  const dot = s.match(/^(\d{1,2})\.(\d{2})$/);
  if (dot) {
    const hh = Number(dot[1]);
    const mm = Number(dot[2]);
    if (Number.isFinite(hh) && Number.isFinite(mm)) return hh * 60 + mm;
  }
  const embedded = String(raw || '').match(/(\d{1,2}):(\d{1,2})\s*(am|pm)?/i);
  if (embedded) {
    let hh = Number(embedded[1]);
    const mm = Number(embedded[2]);
    const ap = (embedded[3] || '').toLowerCase();
    if (ap === 'pm' && hh < 12) hh += 12;
    if (ap === 'am' && hh === 12) hh = 0;
    if (Number.isFinite(hh) && Number.isFinite(mm)) return hh * 60 + mm;
  }
  return null;
}

/** Excel-style H.MM e.g. 8.00, 17.30 — never returns date prefix. */
export function formatFormDRJAttendanceTime(raw) {
  const s = stripAttendanceDateFromTime(raw);
  if (!s || s === '-') return '';
  if (/^\d{1,2}\.\d{2}$/.test(s)) return s;
  const mins = parseTimeToMinutes(raw);
  if (mins == null) return '';
  const h = Math.floor(mins / 60);
  const m = mins % 60;
  return `${h}.${String(m).padStart(2, '0')}`;
}

function readAttendanceFirstIn(rec) {
  if (!rec || typeof rec !== 'object') return '';
  const keys = [
    'FirstIn', 'firstIn', 'ShiftStartTime', 'shiftStartTime', 'InTime', 'inTime', 'CheckIn',
    'CheckInTime', 'checkInTime', 'FirstCheckIn', 'firstCheckIn', 'In', 'in',
    'LoginTime', 'loginTime', 'PunchIn', 'punchIn',
  ];
  for (const k of keys) {
    const v = String(rec[k] ?? '').trim();
    if (v && v !== '-') return v;
  }
  return '';
}

function readAttendanceLastOut(rec) {
  if (!rec || typeof rec !== 'object') return '';
  const keys = [
    'LastOut', 'lastOut', 'ShiftEndTime', 'shiftEndTime', 'OutTime', 'outTime', 'CheckOut',
    'CheckOutTime', 'checkOutTime', 'LastCheckOut', 'lastCheckOut', 'Out', 'out',
    'LogoutTime', 'logoutTime', 'PunchOut', 'punchOut',
  ];
  for (const k of keys) {
    const v = String(rec[k] ?? '').trim();
    if (v && v !== '-') return v;
  }
  return '';
}

function attendanceRecordDateKey(rec) {
  const dRaw = rec?.date ?? rec?.Date ?? '';
  return normalizeAttendanceDateKey(dRaw);
}

function getFormDRJRowType(row) {
  const explicit = String(row?.__formDRJRowType || '').trim().toUpperCase().replace(/\s+/g, ' ');
  if (explicit === FORM_D_RJ_ROW_TYPE_IN || explicit === FORM_D_RJ_ROW_TYPE_OUT) return explicit;
  if (explicit === 'IN / OUT' || explicit === 'IN/OUT') return FORM_D_RJ_ROW_TYPE_COMBINED;
  const dateHdr = Object.keys(row || {}).find((k) => isFormDRJDateHeader(k));
  if (dateHdr) {
    const v = String(row[dateHdr] || '').trim().toUpperCase().replace(/\s+/g, ' ');
    if (v === FORM_D_RJ_ROW_TYPE_IN || v === FORM_D_RJ_ROW_TYPE_OUT) return v;
    if (v === 'IN / OUT' || v === 'IN/OUT') return FORM_D_RJ_ROW_TYPE_COMBINED;
  }
  return '';
}

export function resolveFormDRJEmployeeIndex(row, rowIndex = 0) {
  const fromPair = Math.floor(rowIndex / FORM_D_RJ_ROWS_PER_EMPLOYEE);
  if (row?.__formDRJEmployeeIndex == null || row.__formDRJEmployeeIndex === '') {
    return fromPair;
  }
  const stored = Number(row.__formDRJEmployeeIndex);
  if (!Number.isFinite(stored) || stored < 0) return fromPair;
  if (getFormDRJRowType(row) && stored === rowIndex) return fromPair;
  return stored;
}

export function isFormDRJEmployeeSpanColumnHeader(h) {
  return (
    isFormDRJSerialHeader(h) ||
    isFormDRJNameHeader(h) ||
    isFormDRJPlaceOfWorkHeader(h) ||
    isFormDRJSummaryDaysHeader(h) ||
    isFormDRJRemarksHoursHeader(h)
  );
}

/**
 * Ensure one row per employee. Legacy IN+OUT pairs are merged into a single row
 * with day cells formatted as "IN / OUT" times.
 */
export function expandFormDRajasthanInOutRows(rows, headers) {
  if (!Array.isArray(rows) || rows.length === 0) return rows || [];
  const hdrs = Array.isArray(headers) ? headers : [];
  const dateHeader = hdrs.find((h) => isFormDRJDateHeader(h)) || 'Date';
  const dayHeaders = listFormDRajasthanDayColumnHeaders(hdrs);
  const out = [];

  for (let i = 0; i < rows.length; i += 1) {
    const row = rows[i];
    if (!row || typeof row !== 'object') continue;
    const existingType = getFormDRJRowType(row);

    if (existingType === FORM_D_RJ_ROW_TYPE_IN) {
      const next = rows[i + 1];
      const nextType = next ? getFormDRJRowType(next) : '';
      const merged = { ...row };
      if (nextType === FORM_D_RJ_ROW_TYPE_OUT) {
        dayHeaders.forEach(({ header }) => {
          const inRaw = String(row[header] ?? '').trim();
          const outRaw = String(next[header] ?? '').trim();
          if (/\//.test(inRaw) && !outRaw) {
            merged[header] = formatFormDRJDayCellDisplay(inRaw);
            return;
          }
          const inVal = formatFormDRJAttendanceTime(inRaw) || (/\//.test(inRaw) ? '' : inRaw);
          const outVal = formatFormDRJAttendanceTime(outRaw) || outRaw;
          if (/\//.test(inRaw)) {
            merged[header] = formatFormDRJDayCellDisplay(inRaw);
          } else {
            merged[header] = formatFormDRJInOutCell(inVal, outVal);
          }
        });
        if (next.__employeeLookupName && !merged.__employeeLookupName) {
          merged.__employeeLookupName = next.__employeeLookupName;
        }
        if (next.__employeeLookupId && !merged.__employeeLookupId) {
          merged.__employeeLookupId = next.__employeeLookupId;
        }
        hdrs.forEach((h) => {
          if (!isFormDRJEmployeeSpanColumnHeader(h) && !isFormDRJDateHeader(h)) return;
          const cur = String(merged[h] ?? '').trim();
          const fromOut = String(next[h] ?? '').trim();
          if (!cur && fromOut) merged[h] = fromOut;
        });
        i += 1;
      } else {
        dayHeaders.forEach(({ header }) => {
          const inRaw = String(row[header] ?? '').trim();
          merged[header] = /\//.test(inRaw)
            ? formatFormDRJDayCellDisplay(inRaw)
            : formatFormDRJInOutCell(formatFormDRJAttendanceTime(inRaw) || inRaw, '');
        });
      }
      merged[dateHeader] = FORM_D_RJ_ROW_TYPE_COMBINED;
      merged.__formDRJRowType = FORM_D_RJ_ROW_TYPE_COMBINED;
      merged.__formDRJEmployeeIndex = out.length;
      out.push(merged);
      continue;
    }

    if (existingType === FORM_D_RJ_ROW_TYPE_OUT) {
      const merged = { ...row };
      dayHeaders.forEach(({ header }) => {
        const outRaw = String(row[header] ?? '').trim();
        merged[header] = /\//.test(outRaw)
          ? formatFormDRJDayCellDisplay(outRaw)
          : formatFormDRJInOutCell('', formatFormDRJAttendanceTime(outRaw) || outRaw);
      });
      merged[dateHeader] = FORM_D_RJ_ROW_TYPE_COMBINED;
      merged.__formDRJRowType = FORM_D_RJ_ROW_TYPE_COMBINED;
      merged.__formDRJEmployeeIndex = out.length;
      out.push(merged);
      continue;
    }

    const employeeIndex = out.length;
    const base = { ...row, __formDRJEmployeeIndex: employeeIndex };
    hdrs.forEach((h) => {
      if (base[h] == null) base[h] = '';
    });
    dayHeaders.forEach(({ header }) => {
      const raw = String(base[header] ?? '').trim();
      if (raw) base[header] = formatFormDRJDayCellDisplay(raw);
    });
    base[dateHeader] = FORM_D_RJ_ROW_TYPE_COMBINED;
    base.__formDRJRowType = FORM_D_RJ_ROW_TYPE_COMBINED;
    out.push(base);
  }
  return out;
}

export function applyFormDRajasthanPeopleToRow(row, emp, headers, helpers = {}) {
  const hdrs = Array.isArray(headers) ? headers : [];
  const out = row && typeof row === 'object' ? { ...row } : {};
  const {
    sanitizeValue = (v) => String(v ?? '').trim(),
    overwrite = true,
    rowIndex = 0,
  } = helpers;

  const cellIsEmpty = (header) => {
    const v = String(out[header] || '').trim();
    if (!v) return true;
    return /^enter\b/i.test(v) || v.toLowerCase().includes('enter ');
  };
  const setCell = (header, value) => {
    if (!header || value == null || value === '') return;
    if (!overwrite && !cellIsEmpty(header)) return;
    out[header] = sanitizeValue(value);
  };

  const fullName = readFormDRJEmployeeFullName(emp);
  const designation = readFormDRJDesignation(emp);
  if (fullName) out.__employeeLookupName = fullName;

  hdrs.forEach((header) => {
    if (isFormDRJDateHeader(header)) {
      setCell(header, FORM_D_RJ_ROW_TYPE_COMBINED);
      return;
    }
    if (isFormDRJNameHeader(header)) {
      setCell(header, fullName);
      return;
    }
    if (isFormDRJPlaceOfWorkHeader(header)) {
      setCell(header, designation);
      return;
    }
    if (isFormDRJSerialHeader(header)) {
      const empIdx = resolveFormDRJEmployeeIndex(out, rowIndex);
      setCell(header, String(empIdx + 1));
    }
  });

  out.__formDRJRowType = FORM_D_RJ_ROW_TYPE_COMBINED;
  return out;
}

export function applyFormDRajasthanEmployeeToRow(row, emp, headers, helpers = {}) {
  return applyFormDRajasthanPeopleToRow(row, emp, headers, helpers);
}

function normalizeAttendanceNameKey(name) {
  return String(name || '').trim().toLowerCase().replace(/\s+/g, ' ');
}

function pickAttendanceRecordIds(rec) {
  if (!rec || typeof rec !== 'object') return [];
  const raw = [
    rec.erecno,
    rec.id,
    rec.employeeId,
    rec.EmployeeId,
    rec['Employee ID'],
    rec.EmployeeID,
    rec.Zoho_ID,
    rec.ZohoID,
    rec.zoho_id,
  ];
  return [...new Set(raw.map((v) => String(v || '').trim().toLowerCase()).filter(Boolean))];
}

function pickAttendanceRecordName(rec) {
  if (!rec || typeof rec !== 'object') return '';
  const first = String(
    rec.firstName || rec.FirstName || rec.First_Name || rec['First Name'] || ''
  ).trim();
  const last = String(
    rec.lastName || rec.LastName || rec.Last_Name || rec['Last Name'] || ''
  ).trim();
  const full = `${first} ${last}`.trim();
  return normalizeAttendanceNameKey(
    full ||
      rec.EmployeeName ||
      rec['Employee Name'] ||
      rec.EmployeeName ||
      rec.Name ||
      rec.name ||
      ''
  );
}

function pickEmployeeAttendanceIds(emp, row) {
  const src = unwrapEmployeeRecord(emp);
  const raw = [];
  if (row?.__employeeLookupId) raw.push(row.__employeeLookupId);
  if (src && typeof src === 'object') {
    raw.push(
      src.erecno,
      src.EmployeeId,
      src.EmployeeID,
      src['Employee ID'],
      src.id,
      src.Zoho_ID,
      src.ZohoID,
      src.zoho_id,
      src['Role.ID'],
      src.Role && typeof src.Role === 'object' ? src.Role.ID : ''
    );
  }
  return [...new Set(raw.map((v) => String(v || '').trim().toLowerCase()).filter(Boolean))];
}

function pickEmployeeAttendanceNames(emp, row) {
  const names = [];
  if (row?.__employeeLookupName) names.push(row.__employeeLookupName);
  const full = readFormDRJEmployeeFullName(emp);
  if (full) names.push(full);
  const src = unwrapEmployeeRecord(emp);
  if (src && typeof src === 'object') {
    const first = String(src.FirstName || src.firstName || src['First Name'] || '').trim();
    const last = String(src.LastName || src.lastName || src['Last Name'] || '').trim();
    if (first) names.push(first);
    if (last) names.push(last);
    if (src.Name) names.push(src.Name);
    if (src['Employee Name']) names.push(src['Employee Name']);
  }
  return [...new Set(names.map(normalizeAttendanceNameKey).filter(Boolean))];
}

function namesLooselyMatchAttendance(a, b) {
  const x = normalizeAttendanceNameKey(a);
  const y = normalizeAttendanceNameKey(b);
  if (!x || !y) return false;
  if (x === y) return true;
  if (x.includes(y) || y.includes(x)) return true;
  const xFirst = x.split(/\s+/)[0];
  const yFirst = y.split(/\s+/)[0];
  return Boolean(xFirst && yFirst && xFirst === yFirst);
}

function indexAttendanceByEmployee(attendanceRecords, emp, row) {
  const idKeys = pickEmployeeAttendanceIds(emp, row);
  const nameKeys = pickEmployeeAttendanceNames(emp, row);
  const out = [];
  const seen = new Set();
  (attendanceRecords || []).forEach((rec) => {
    if (!rec || typeof rec !== 'object' || seen.has(rec)) return;
    const recIds = pickAttendanceRecordIds(rec);
    const recName = pickAttendanceRecordName(rec);
    const idHit = idKeys.length > 0 && recIds.some((id) => idKeys.includes(id));
    const nameHit =
      recName && nameKeys.some((name) => namesLooselyMatchAttendance(name, recName));
    if (idHit || nameHit) {
      seen.add(rec);
      out.push(rec);
    }
  });
  return out;
}

export function enrichFormDRajasthanRowsFromAttendance(
  rows,
  headers,
  attendanceRecords,
  options = {}
) {
  if (!Array.isArray(rows) || rows.length === 0) return 0;
  const hdrs = Array.isArray(headers) ? headers : [];
  const dayHeaders = listFormDRajasthanDayColumnHeaders(hdrs);
  if (dayHeaders.length === 0) return 0;

  const {
    targetYear,
    targetMonthIndex,
    overwrite = true,
    employees = null,
  } = options;

  const year = Number(targetYear);
  const monthIdx = Number(targetMonthIndex);
  if (!Number.isFinite(year) || !Number.isFinite(monthIdx)) return 0;

  let filled = 0;
  rows.forEach((row, rowIndex) => {
    if (!row || typeof row !== 'object') return;
    const rowType = getFormDRJRowType(row);
    if (rowType === FORM_D_RJ_ROW_TYPE_OUT) return;

    let empIndex = resolveFormDRJEmployeeIndex(row, rowIndex);
    if (Array.isArray(employees) && employees.length > 0) {
      const stored = Number(row.__formDRJEmployeeIndex);
      if (!(Number.isFinite(stored) && stored >= 0 && stored < employees.length)) {
        empIndex = rowIndex;
      } else {
        empIndex = stored;
      }
    }
    const empItem = Array.isArray(employees) ? employees[empIndex] : null;
    const emp = empItem && (empItem.Employee || empItem.employee || empItem);

    const empRecs = indexAttendanceByEmployee(attendanceRecords, emp, row);
    const byDate = new Map();
    empRecs.forEach((rec) => {
      const dk = attendanceRecordDateKey(rec);
      if (dk) byDate.set(dk, rec);
    });

    dayHeaders.forEach(({ header, day }) => {
      const existing = String(row[header] ?? '').trim();
      const isAttendanceStatusCode = /^[PAWOHL]$/i.test(existing);
      const hasEmbeddedDate = /\d{1,2}-[A-Za-z]{3}-\d{4}|\d{4}-\d{2}-\d{2}[T\s]/.test(existing);
      if (hasEmbeddedDate) {
        const fixed = formatFormDRJDayCellDisplay(existing);
        if (fixed) {
          row[header] = fixed;
          filled += 1;
        }
        return;
      }
      if (!overwrite && existing && !/^enter\b/i.test(existing) && !isAttendanceStatusCode) return;

      const dayDate = new Date(year, monthIdx, day);
      if (dayDate.getMonth() !== monthIdx) return;
      const y = dayDate.getFullYear();
      const m = String(dayDate.getMonth() + 1).padStart(2, '0');
      const d = String(dayDate.getDate()).padStart(2, '0');
      const iso = `${y}-${m}-${d}`;
      const rec = byDate.get(iso);
      if (!rec) return;

      const inVal = formatFormDRJAttendanceTime(readAttendanceFirstIn(rec));
      const outVal = formatFormDRJAttendanceTime(readAttendanceLastOut(rec));
      const val = formatFormDRJInOutCell(inVal, outVal);
      if (val) {
        row[header] = val;
        filled += 1;
      }
    });
  });
  return filled;
}

export function readFormDRJPaidDays(payrollRow) {
  if (!payrollRow || payrollRow.fetch_error) return '';
  const flat = flattenPayrollEarningColumns(payrollRow);
  const merged = { ...flat, ...payrollRow };
  const val = readPayrollScalar(
    merged,
    ['paid_days', 'Paid_days', 'Paid Days', 'days_worked', 'Days Worked', 'paidDays', 'no_of_days_worked'],
    [/^paid_?days$/i, /paiddays/i, /daysworked/i]
  );
  if (val === '' || val == null) return '';
  return String(val).trim();
}

export function applyFormDRajasthanPayrollToRow(row, payrollRow, headers, helpers = {}) {
  if (!row || !payrollRow || payrollRow.fetch_error) return false;
  const rowType = getFormDRJRowType(row);
  if (rowType === FORM_D_RJ_ROW_TYPE_OUT) return false;

  const hdrs = Array.isArray(headers) ? headers : [];
  const paidDays = readFormDRJPaidDays(payrollRow);
  if (paidDays === '') return false;

  const { overwrite = true, sanitizeValue = (v) => String(v ?? '').trim() } = helpers;
  let changed = false;
  hdrs.forEach((header) => {
    if (!isFormDRJSummaryDaysHeader(header)) return;
    const existing = String(row[header] ?? '').trim();
    if (!overwrite && existing && !/^enter\b/i.test(existing)) return;
    row[header] = sanitizeValue(paidDays);
    changed = true;
  });
  return changed;
}

export function enrichFormDRajasthanPayrollRows(mappedData, employees, headers, helpers = {}) {
  const hdrs = resolveFormDRajasthanTableHeaders(headers);
  const {
    resolvePayrollRow = null,
    payrollRows = null,
    sanitizeValue = (v) => String(v ?? '').trim(),
    overwrite = true,
  } = helpers;
  if (!Array.isArray(mappedData) || mappedData.length === 0) return 0;

  let filled = 0;
  mappedData.forEach((row, rowIndex) => {
    if (!row || typeof row !== 'object') return;
    const rowType = getFormDRJRowType(row);
    if (rowType === FORM_D_RJ_ROW_TYPE_OUT) return;

    const empIndex = resolveFormDRJEmployeeIndex(row, rowIndex);
    const empItem = Array.isArray(employees) ? employees[empIndex] : null;
    const emp = empItem && (empItem.Employee || empItem.employee || empItem);
    let payrollRow =
      typeof resolvePayrollRow === 'function' ? resolvePayrollRow(emp || {}, row, rowIndex) : null;
    if ((!payrollRow || payrollRow.fetch_error) && Array.isArray(payrollRows) && payrollRows.length > 0) {
      payrollRow = payrollRows[empIndex] || null;
    }
    if (
      payrollRow &&
      !payrollRow.fetch_error &&
      applyFormDRajasthanPayrollToRow(row, payrollRow, hdrs, { overwrite, sanitizeValue })
    ) {
      filled += 1;
    }
  });
  return filled;
}

export function getFormDRajasthanRowValueForHeader(row, header, rowIndex = 0, allRows = null) {
  if (!row || !header) return '';
  const hdr = String(header || '').trim();
  let direct = row[header];
  if (direct == null || String(direct).trim() === '') {
    if (Object.prototype.hasOwnProperty.call(row, hdr)) {
      direct = row[hdr];
    } else if (/^\d{1,2}$/.test(hdr)) {
      const dayNum = Number(hdr);
      direct = row[hdr] ?? row[dayNum] ?? row[`Day ${hdr}`] ?? row[`day_${hdr}`];
    } else {
      const targetNorm = normHeaderLabel(hdr);
      for (const [key, val] of Object.entries(row)) {
        if (String(key).startsWith('__')) continue;
        if (normHeaderLabel(key) === targetNorm) {
          direct = val;
          break;
        }
      }
    }
  }
  if (direct != null && String(direct).trim() !== '') {
    const text = String(direct).trim();
    if (/^\d{1,2}$/.test(hdr) || listFormDRajasthanDayColumnHeaders([hdr]).length === 1) {
      return formatFormDRJDayCellDisplay(text);
    }
    return text;
  }
  if (isFormDRJSerialHeader(header)) {
    return String(resolveFormDRJEmployeeIndex(row, rowIndex) + 1);
  }
  return '';
}

export function rowHasMeaningfulFormDRajasthanExportData(row, headers) {
  const hdrs = resolveFormDRajasthanTableHeaders(headers);
  return hdrs.some((header) => {
    if (isFormDRJSerialHeader(header) || isFormDRJDateHeader(header)) return false;
    return getFormDRajasthanRowValueForHeader(row, header) !== '';
  });
}

export function filterFormDRajasthanExportRows(rows, headers) {
  const hdrs = resolveFormDRajasthanTableHeaders(headers);
  return (Array.isArray(rows) ? rows : []).filter((row) =>
    rowHasMeaningfulFormDRajasthanExportData(row, hdrs)
  );
}

export function remapFormDRajasthanRowsToHeaders(rows, sourceHeaders, targetHeaders) {
  const src = Array.isArray(sourceHeaders) ? sourceHeaders : [];
  const tgt = resolveFormDRajasthanTableHeaders(targetHeaders);
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
        for (const [k, v] of Object.entries(row)) {
          if (normHeaderLabel(k) === normHeaderLabel(targetHeader)) {
            val = v;
            break;
          }
        }
      }
      if ((val == null || val === '') && isFormDRJSerialHeader(targetHeader)) {
        val = rowIndex + 1;
      }
      out[targetHeader] = val == null ? '' : val;
      if (row.__formDRJRowType) out.__formDRJRowType = row.__formDRJRowType;
      if (row.__formDRJEmployeeIndex != null) out.__formDRJEmployeeIndex = row.__formDRJEmployeeIndex;
      if (row.__employeeLookupName) out.__employeeLookupName = row.__employeeLookupName;
      if (row.__employeeLookupId) out.__employeeLookupId = row.__employeeLookupId;
    });
    return out;
  });
}

export function buildFormDRajasthanPeriodLine(monthName, year) {
  let idx = MONTH_NAMES_DRJ.findIndex(
    (m) => m.toLowerCase() === String(monthName || '').toLowerCase().trim()
  );
  if (idx < 0) {
    const token = String(monthName || '').toLowerCase().trim().slice(0, 3);
    idx = MONTH_NAMES_DRJ.findIndex((m) => m.toLowerCase().startsWith(token));
  }
  if (idx < 0) idx = new Date().getMonth();
  const y = Number(year);
  if (!Number.isFinite(y) || y < 1900) return '';
  const m = idx + 1;
  const lastDay = new Date(y, idx + 1, 0).getDate();
  const pad = (n) => String(n).padStart(2, '0');
  return `For the period From ${pad(1)}-${pad(m)}-${y} to ${pad(lastDay)}-${pad(m)}-${y}`;
}

export function enrichFormDRajasthanDisplayHeader(formHeader, fileName, item, tableHeaders, sheetText) {
  const base = formHeader && typeof formHeader === 'object' ? { ...formHeader } : {};
  if (!isFormDRajasthanContext(base, item, fileName, sheetText, tableHeaders)) return base;

  const fields = enrichEstablishmentPrincipalEmployerHeaderFields(base.fields || []);
  const hasPeriod = fields.some((field) => /for\s+the\s+period\s+from/i.test(String(field?.label || '')));
  if (!hasPeriod) {
    fields.push({ label: 'For the period From', value: '', key: 'form_d_rj_period' });
  }

  return {
    ...base,
    fields,
    formDRJAttendanceTableLayout: true,
    formVIIAPHeaderFieldLayout: false,
    formXXVIAPHeaderFieldLayout: false,
    formXIXAPHeaderFieldLayout: false,
    title: /\bform[\s._-]*d\b/i.test(String(base.title || '')) ? base.title || 'FORM D' : 'FORM D',
    subtitle:
      /attendance\s+register|format\s+of\s+attendance/i.test(String(base.subtitle || ''))
        ? base.subtitle
        : 'Format of Attendance Register',
  };
}

export function prepareFormDRajasthanDownloadHeaderData(
  headerFormData,
  parsedFormHeader = null,
  siteContext = {}
) {
  const out = headerFormData && typeof headerFormData === 'object' ? { ...headerFormData } : {};
  const {
    establishmentText = '',
    principalEmployerText = '',
    periodText = '',
  } = siteContext;

  if (establishmentText) {
    out.statutory_establishment_name = establishmentText;
    out.form_d_rj_establishment = establishmentText;
  }
  if (principalEmployerText) {
    out.statutory_principal_employer = principalEmployerText;
    out.form_d_rj_principal_employer = principalEmployerText;
  }
  if (periodText) {
    out.form_d_rj_period = periodText;
    out.statutory_period_from = periodText;
  }

  const fields = Array.isArray(parsedFormHeader?.fields) ? parsedFormHeader.fields : [];
  fields.forEach((field) => {
    const key = field?.key;
    if (!key || String(out[key] ?? '').trim()) return;
    if (/establishment/i.test(String(field.label || '')) && !/principal|employer|contractor/.test(String(field.label || ''))) {
      if (establishmentText) out[key] = establishmentText;
    } else if (/principal\s+employer/i.test(String(field.label || ''))) {
      if (principalEmployerText) out[key] = principalEmployerText;
    } else if (/for\s+the\s+period\s+from/i.test(String(field.label || ''))) {
      if (periodText) out[key] = periodText;
    }
  });

  return out;
}

function writeFormDRajasthanPeriodCell(worksheet, periodText, headerRowEnd = 25) {
  const line = String(periodText || '').trim();
  if (!line || !worksheet) return;
  for (let r = 1; r <= headerRowEnd; r += 1) {
    for (let c = 1; c <= 24; c += 1) {
      const raw = excelCellValueToString(worksheet.getCell(r, c)?.value).trim();
      if (/for\s+the\s+period\s+from/i.test(raw)) {
        worksheet.getCell(r, c).value = line;
        return;
      }
    }
  }
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

/** Remove template IN/OUT pair merges so one employee row can own Sl/Name/Place cells. */
function unmergeFormDRJExcelJSRowsInRange(worksheet, startRow, endRow, colFrom, colTo) {
  const merges = worksheet?.model?.merges;
  if (!worksheet || !Array.isArray(merges) || merges.length === 0) return;
  const toRemove = [];
  for (const range of merges) {
    const parts = String(range || '').split(':');
    if (parts.length !== 2) continue;
    const start = parts[0].match(/^([A-Z]+)(\d+)$/i);
    const end = parts[1].match(/^([A-Z]+)(\d+)$/i);
    if (!start || !end) continue;
    const r1 = parseInt(start[2], 10);
    const r2 = parseInt(end[2], 10);
    const c1 = XLSX.utils.decode_col(start[1].toUpperCase()) + 1;
    const c2 = XLSX.utils.decode_col(end[1].toUpperCase()) + 1;
    const rowOverlap = r2 >= startRow && r1 <= endRow;
    const colOverlap = c2 >= colFrom && c1 <= colTo;
    if (rowOverlap && colOverlap) toRemove.push(range);
  }
  toRemove.forEach((range) => {
    try {
      worksheet.unMergeCells(range);
    } catch (_) {
      // ignore invalid merge ranges
    }
  });
}

function detectFormDRajasthanTableLayout(worksheet, hints = {}) {
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
    const maxScanRows = Math.max(40, worksheet.rowCount + 5);
    for (let r = 1; r <= maxScanRows; r += 1) {
      let found = false;
      for (let c = 1; c <= 25; c += 1) {
        const text = getMergedAwareCellText(r, c);
        if (/place\s+of\s+work/i.test(text) || /summary\s+no\.?\s*of\s+days/i.test(text)) {
          headerRow = r;
          startCol = c > 1 ? c - 1 : 1;
          found = true;
          break;
        }
      }
      if (found) break;
    }
  }

  if (headerRow < 1) return null;

  const templateCols = [];
  for (let c = startCol; c <= startCol + 45; c += 1) {
    const label = getMergedAwareCellText(headerRow, c);
    if (!label) {
      if (templateCols.length >= 10) break;
      continue;
    }
    if (/^in$|^out$/i.test(label)) continue;
    templateCols.push({ col: c, label });
    if (templateCols.length >= 40) break;
  }
  if (templateCols.length < 6) return null;

  let dataStartRow =
    hints.parsedDataStartIndex != null && hints.parsedDataStartIndex >= 0
      ? hints.parsedDataStartIndex + 1
      : headerRow + 1;

  const rowLooksLikeInOutSubHeader = (r) => {
    let sawIn = false;
    let sawOut = false;
    let otherText = 0;
    for (let c = startCol; c <= startCol + 12; c += 1) {
      const t = getMergedAwareCellText(r, c).toLowerCase().trim();
      if (!t) continue;
      if (t === 'in') sawIn = true;
      else if (t === 'out') sawOut = true;
      else if (!/^\d{1,2}$/.test(t)) otherText += 1;
    }
    if (otherText > 2) return false;
    return sawIn || sawOut;
  };

  let guard = 0;
  while (guard < 4 && rowLooksLikeInOutSubHeader(dataStartRow)) {
    dataStartRow += 1;
    guard += 1;
  }

  return { headerRow, dataStartRow, templateCols, startCol };
}

export async function buildFormDRajasthanWorkbookWithTemplateStyles({
  templateArrayBuffer,
  mappedData,
  headersToUse,
  parsedHeaderRowIndex,
  parsedDataStartIndex,
  parsedTableStartCol,
  parsedFormHeader,
  formFileName,
  headerFormData,
  preferredSheetName = '',
}) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(templateArrayBuffer);
  let worksheet = null;
  if (preferredSheetName) worksheet = workbook.getWorksheet(preferredSheetName);
  if (!worksheet) {
    for (const ws of workbook.worksheets || []) {
      if (/\bform[\s._-]*d\b/i.test(String(ws.name || ''))) {
        worksheet = ws;
        break;
      }
    }
  }
  if (!worksheet) worksheet = workbook.worksheets?.[0] || null;
  if (!worksheet) throw new Error('Template worksheet not found.');

  const layout = detectFormDRajasthanTableLayout(worksheet, {
    parsedHeaderRowIndex,
    parsedDataStartIndex,
    parsedTableStartCol,
  });
  if (!layout) throw new Error('Could not locate Rajasthan Form D table header row.');

  const { dataStartRow, templateCols } = layout;

  writeStatutoryHeaderFieldsToExcelJsWorksheet(worksheet, {
    headerFormData: headerFormData || {},
    parsedFormHeader,
    headerRowEnd: Math.max(1, layout.headerRow - 1),
    maxScanCols: 80,
    writeMode: 'both',
  });
  writeFormDRajasthanPeriodCell(
    worksheet,
    headerFormData?.form_d_rj_period || headerFormData?.statutory_period_from,
    Math.max(1, layout.headerRow - 1)
  );

  const templateHeaderLabels = templateCols.map((entry) => entry.label);
  const normalizedHeaders = resolveFormDRajasthanTableHeaders(
    Array.isArray(headersToUse) && headersToUse.length >= templateHeaderLabels.length
      ? headersToUse
      : templateHeaderLabels
  );
  const rows = filterFormDRajasthanExportRows(
    remapFormDRajasthanRowsToHeaders(
      expandFormDRajasthanInOutRows(Array.isArray(mappedData) ? mappedData : [], normalizedHeaders),
      normalizedHeaders,
      normalizedHeaders
    ),
    normalizedHeaders
  );

  const tableColMin = Math.min(...templateCols.map(({ col }) => col));
  const tableColMax = Math.max(...templateCols.map(({ col }) => col));
  const clearToRow = Math.max(dataStartRow + rows.length + 10, dataStartRow + 20);

  // Template still has vertical merges for old IN/OUT pairs — clear them so every
  // employee row can show Sl. No / Name / Place of work.
  unmergeFormDRJExcelJSRowsInRange(worksheet, dataStartRow, clearToRow, tableColMin, tableColMax);

  for (let r = dataStartRow; r <= clearToRow; r += 1) {
    for (let c = tableColMin; c <= tableColMax; c += 1) {
      worksheet.getCell(r, c).value = '';
    }
  }

  rows.forEach((row, idx) => {
    const excelRow = dataStartRow + idx;
    const targetRow = worksheet.getRow(excelRow);
    if (targetRow) targetRow.height = 28;
    templateCols.forEach(({ col, label }) => {
      let val = getFormDRajasthanRowValueForHeader(row, label, idx, rows);
      if (isFormDRJSerialHeader(label)) {
        val = String(idx + 1);
      }
      const cell = worksheet.getCell(excelRow, col);
      if (val == null || val === '') {
        cell.value = '';
        return;
      }
      if (isFormDRJSerialHeader(label)) {
        const n = Number(String(val).replace(/[,]/g, '').trim());
        cell.value = Number.isFinite(n) ? n : String(val);
      } else if (listFormDRajasthanDayColumnHeaders([label]).length === 1) {
        // Stack IN / OUT on separate lines for print-friendly Excel cells.
        cell.value = String(val).replace(/\s*\/\s*/g, '\n');
      } else {
        cell.value = String(val);
      }
      const isIdentity =
        isFormDRJSerialHeader(label) ||
        isFormDRJNameHeader(label) ||
        isFormDRJPlaceOfWorkHeader(label);
      cell.alignment = {
        ...(cell.alignment || {}),
        vertical: 'middle',
        horizontal: isIdentity && !isFormDRJSerialHeader(label) ? 'left' : 'center',
        wrapText: true,
      };
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
    'Form_D_RJ_Rajasthan.xlsx';
  return {
    blob: new Blob([out], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    }),
    fileName,
  };
}
