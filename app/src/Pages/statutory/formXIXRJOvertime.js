import * as XLSX from 'xlsx';
import {
  flattenPayrollEarningColumns,
  readForm10GrossPayAmount,
  readPayrollNetPayForStatutory,
} from '../../utils/payrollEarnings';
import { resolveHeaderFieldExportValue } from '../../utils/statutorySiteCompanyHeaders';
import { isFormXIXRajasthanOvertimeRegisterContext } from './formXIXAPWageSlip';

/** Rajasthan Form XIX — Register of Overtime [Rule 77(2)(e)] autofill helpers. */

const FORM_XIX_RJ_NIL = 'Nil';

/** Form_XIX_RJ.xlsx keeps an empty (often merged) column A. Serial No. is column B. */
export const FORM_XIX_RJ_DEFAULT_TABLE_START_COL0 = 1;

/** Official Form XIX RJ column order [See Rule 77 (2)(e)] — no blank spacer columns. */
export const FORM_XIX_RJ_TABLE_HEADERS = [
  'Serial No.',
  'Name of workman',
  "Father's/Husband's name",
  'Sex',
  'Designation and department',
  'Date on which overtime work was put in',
  'Wages of overtime on each occasion',
  'Total over time worked or production in case of piece rates',
  'Normal Hours',
  'Normal rate',
  'Overtime rate',
  'Normal earning',
  'Overtime earnings',
  'Total earnings',
  'Date on which overtime payment made',
];

const normHeader = (h) =>
  String(h || '')
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();

const isBlankOrGeneratedHeader = (h) => {
  const t = String(h || '').trim();
  if (!t) return true;
  return /^column\s+([a-z]+|\d+)$/i.test(t);
};

export function isFormXIXRJSerialHeader(h) {
  const s = normHeader(h);
  return (
    /^s\.?\s*no\.?$/.test(s) ||
    /^sl\.?\s*no\.?$/.test(s) ||
    /^sr\.?\s*no\.?$/.test(s) ||
    s === 'serial no' ||
    s === 'serial number' ||
    /^serial\s+no\.?$/.test(s)
  );
}

export function isFormXIXRJWorkmanNameHeader(h) {
  const s = normHeader(h);
  if (s.includes('father') || s.includes('husband')) return false;
  return (
    /name\s+of\s+workman/.test(s) ||
    (s.includes('name') && (s.includes('workman') || s.includes('workmen')))
  );
}

export function isFormXIXRJFatherHusbandHeader(h) {
  const s = normHeader(h);
  return s.includes('father') || s.includes('husband');
}

export function isFormXIXRJSexHeader(h) {
  const s = normHeader(h);
  return s === 'sex' || s.includes('male/female') || (s.includes('gender') && !s.includes('name'));
}

export function isFormXIXRJDesignationDepartmentHeader(h) {
  const s = normHeader(h);
  return s.includes('designation') && s.includes('department');
}

export function isFormXIXRJOvertimeWorkedDateHeader(h) {
  const s = normHeader(h);
  if (/payment|paid|rate|wage|earning|hours|normal/.test(s)) return false;
  return (
    /dates?\s+on\s+which\s+over[\s-]*time\s+work\s+was\s+put\s+in/.test(s) ||
    (s.includes('date') && s.includes('overtime') && (s.includes('put in') || s.includes('worked')))
  );
}

export function isFormXIXRJWagesOfOvertimeOccasionHeader(h) {
  const s = normHeader(h);
  return /wages?\s+of\s+over[\s-]*time\s+on\s+each\s+occasion/.test(s);
}

export function isFormXIXRJTotalOvertimeWorkedHeader(h) {
  const s = normHeader(h);
  return (
    (s.includes('total') && s.includes('overtime') && (s.includes('worked') || s.includes('production'))) ||
    (s.includes('overtime') && s.includes('piece') && s.includes('rate'))
  );
}

export function isFormXIXRJNormalHoursHeader(h) {
  const s = normHeader(h);
  return s.includes('normal') && s.includes('hour');
}

export function isFormXIXRJNormalRateHeader(h) {
  const s = normHeader(h);
  if (/overtime/.test(s)) return false;
  return s.includes('normal') && s.includes('rate');
}

export function isFormXIXRJOvertimeRateHeader(h) {
  const s = normHeader(h);
  if (/normal/.test(s)) return false;
  return s.includes('overtime') && s.includes('rate');
}

export function isFormXIXRJNormalEarningsHeader(h) {
  const s = normHeader(h);
  return s.includes('normal') && s.includes('earning');
}

export function isFormXIXRJOvertimeEarningsHeader(h) {
  const s = normHeader(h);
  return s.includes('overtime') && s.includes('earning');
}

export function isFormXIXRJTotalEarningsHeader(h) {
  const s = normHeader(h);
  if (s.includes('overtime')) return false;
  return (s.includes('total') && s.includes('earning')) || s === 'total earnings';
}

export function isFormXIXRJOvertimePaymentDateHeader(h) {
  const s = normHeader(h);
  return (
    /dates?\s+on\s+which\s+over[\s-]*time\s+payment/.test(s) ||
    (s.includes('overtime') && s.includes('payment') && s.includes('date')) ||
    (s.includes('overtime') && s.includes('paid') && s.includes('date'))
  );
}

const FORM_XIX_RJ_HEADER_MATCHERS = [
  isFormXIXRJSerialHeader,
  isFormXIXRJWorkmanNameHeader,
  isFormXIXRJFatherHusbandHeader,
  isFormXIXRJSexHeader,
  isFormXIXRJDesignationDepartmentHeader,
  isFormXIXRJOvertimeWorkedDateHeader,
  isFormXIXRJWagesOfOvertimeOccasionHeader,
  isFormXIXRJTotalOvertimeWorkedHeader,
  isFormXIXRJNormalHoursHeader,
  isFormXIXRJNormalRateHeader,
  isFormXIXRJOvertimeRateHeader,
  isFormXIXRJNormalEarningsHeader,
  isFormXIXRJOvertimeEarningsHeader,
  isFormXIXRJTotalEarningsHeader,
  isFormXIXRJOvertimePaymentDateHeader,
];

function classifyFormXIXRJHeaderIndex(h) {
  for (let i = 0; i < FORM_XIX_RJ_HEADER_MATCHERS.length; i += 1) {
    if (FORM_XIX_RJ_HEADER_MATCHERS[i](h)) return i;
  }
  return -1;
}

/** Drop blank/spacer Excel columns and force official Form XIX RJ order. */
export function resolveFormXIXRJTableHeaders(_tableHeaders) {
  // Always canonical — never keep a leading blank Excel spacer as column 1.
  return [...FORM_XIX_RJ_TABLE_HEADERS];
}

function combineFormXIXRJHeaderBandText(readCell, startRow, col, extraRows = 2) {
  const parts = [];
  for (let i = 0; i <= extraRows; i += 1) {
    const t = String(readCell(startRow + i, col) || '')
      .replace(/\r?\n/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
    if (t && parts[parts.length - 1] !== t) parts.push(t);
  }
  return parts.join(' ');
}

/**
 * Consecutive 1..15 numbering row under the table headers.
 * Form_XIX_RJ.xlsx uses this row; a gap (1 in A, 2 in C) is NOT the original layout.
 */
function findFormXIXRJConsecutiveNumberBand(readCell, maxRows, maxCols, colCount) {
  for (let r = 0; r < maxRows; r += 1) {
    for (let start = 0; start <= maxCols - 8; start += 1) {
      let hits = 0;
      for (let j = 0; j < colCount; j += 1) {
        const t = String(readCell(r, start + j) || '')
          .replace(/\s+/g, '')
          .trim();
        if (t === String(j + 1)) hits += 1;
      }
      if (hits >= Math.max(10, Math.floor(colCount * 0.7))) {
        return { indexRow: r, startCol: start };
      }
    }
  }
  return null;
}

function findFormXIXRJNameColumn(readCell, maxRows, maxCols) {
  for (let r = 0; r < maxRows; r += 1) {
    for (let c = 0; c < maxCols; c += 1) {
      const thisCell = String(readCell(r, c) || '')
        .replace(/\r?\n/g, ' ')
        .replace(/\s+/g, ' ')
        .trim();
      if (!thisCell) continue;
      const band = combineFormXIXRJHeaderBandText(readCell, r, c, 2);
      if (isFormXIXRJWorkmanNameHeader(thisCell) || isFormXIXRJWorkmanNameHeader(band)) {
        return { nameRow: r, nameCol: c };
      }
    }
  }
  return null;
}

/**
 * Sequential canonical map starting at column B unless a real Serial No. column is given.
 * Never starts at column A — Form_XIX_RJ.xlsx uses A as a left-margin spacer.
 */
export function buildFormXIXRJExportColMap(
  tableStartCol = FORM_XIX_RJ_DEFAULT_TABLE_START_COL0,
  headers = FORM_XIX_RJ_TABLE_HEADERS
) {
  const parsed = Number(tableStartCol);
  const start =
    Number.isFinite(parsed) && parsed >= 1 ? parsed : FORM_XIX_RJ_DEFAULT_TABLE_START_COL0;
  const hdrs = resolveFormXIXRJTableHeaders(headers);
  const headerToCol = new Map();
  hdrs.forEach((header, j) => {
    if (header) headerToCol.set(header, start + j);
  });
  return { headers: hdrs, headerToCol, tableStartCol: start };
}

function buildConsecutiveFormXIXRJColMap(startCol, headers, snoRow = -1, dataStartRow = -1) {
  const start =
    Number.isFinite(startCol) && startCol >= 0 ? startCol : FORM_XIX_RJ_DEFAULT_TABLE_START_COL0;
  const mapped = buildFormXIXRJExportColMap(Math.max(start, 0), headers);
  // buildFormXIXRJExportColMap bumps 0 → B; allow a true start at A only when Serial really sits there.
  if (start === 0) {
    const hdrs = resolveFormXIXRJTableHeaders(headers);
    const headerToCol = new Map();
    hdrs.forEach((header, j) => {
      if (header) headerToCol.set(header, j);
    });
    return { headers: hdrs, headerToCol, tableStartCol: 0, snoRow, dataStartRow };
  }
  return { ...mapped, snoRow, dataStartRow };
}

/**
 * Map canonical Form XIX RJ headers onto the original template's physical Excel columns.
 * Original Form_XIX_RJ.xlsx is 15 consecutive columns (Serial beside Name) — usually B–P.
 * Never insert a blank gap between Serial No. and Name of workman.
 */
export function mapFormXIXRJHeadersToTemplateCols(
  readCell,
  { headers = FORM_XIX_RJ_TABLE_HEADERS, maxRows = 45, maxCols = 32 } = {}
) {
  const canon = resolveFormXIXRJTableHeaders(headers);
  const fallback = buildFormXIXRJExportColMap(FORM_XIX_RJ_DEFAULT_TABLE_START_COL0, canon);
  if (typeof readCell !== 'function') {
    return { ...fallback, snoRow: -1, dataStartRow: -1 };
  }

  const numberBand = findFormXIXRJConsecutiveNumberBand(readCell, maxRows, maxCols, canon.length);
  if (numberBand) {
    const snoRow = Math.max(0, numberBand.indexRow - 1);
    return buildConsecutiveFormXIXRJColMap(
      numberBand.startCol,
      canon,
      snoRow,
      numberBand.indexRow + 1
    );
  }

  const nameHit = findFormXIXRJNameColumn(readCell, maxRows, maxCols);
  if (nameHit && nameHit.nameCol >= 1) {
    // Serial No. is the column immediately left of Name — do not skip a spacer in between.
    const startCol = nameHit.nameCol - 1;
    return buildConsecutiveFormXIXRJColMap(startCol, canon, nameHit.nameRow, nameHit.nameRow + 2);
  }

  return { ...fallback, snoRow: -1, dataStartRow: -1 };
}

/** SheetJS worksheet → original Form XIX RJ column map (0-based). */
export function mapFormXIXRJHeadersToSheetJsCols(worksheet, headers = FORM_XIX_RJ_TABLE_HEADERS) {
  if (!worksheet) return buildFormXIXRJExportColMap(FORM_XIX_RJ_DEFAULT_TABLE_START_COL0, headers);
  const merges = worksheet['!merges'] || [];
  const raw = (r, c) => {
    if (r < 0 || c < 0) return '';
    const ref = XLSX.utils.encode_cell({ r, c });
    const cell = worksheet[ref];
    if (cell && cell.v != null) return String(cell.v).trim();
    return '';
  };
  const readCell = (r, c) => {
    const direct = raw(r, c);
    if (direct) return direct;
    for (let i = 0; i < merges.length; i += 1) {
      const m = merges[i];
      if (!m?.s || !m?.e) continue;
      if (r >= m.s.r && r <= m.e.r && c >= m.s.c && c <= m.e.c) {
        const t = raw(m.s.r, m.s.c);
        if (t) return t;
      }
    }
    return '';
  };
  return mapFormXIXRJHeadersToTemplateCols(readCell, { headers });
}

/** Read one cell for Form XIX RJ modal grid (handles blank-key serial shift). */
export function readFormXIXRJRowCell(row, header, colIndex = -1, rowIndex = 0) {
  if (!row || typeof row !== 'object') return '';
  const blankSerial = String(row[''] ?? '').trim();
  const underSerial = String(row['Serial No.'] ?? row['Serial No'] ?? '').trim();
  // Only treat as left-shifted when Serial No. itself holds a non-numeric (name) value.
  // After remap, Serial No. is numeric and a leftover '' key must not re-shift the grid.
  const looksShifted =
    /^\d+$/.test(blankSerial) && underSerial !== '' && !/^\d+$/.test(underSerial);

  if (looksShifted) {
    if (colIndex <= 0 || isFormXIXRJSerialHeader(header)) return blankSerial;
    const prev = FORM_XIX_RJ_TABLE_HEADERS[colIndex - 1];
    const fromPrev = row[prev];
    if (fromPrev != null && String(fromPrev).trim() !== '' && String(fromPrev) !== 'false') {
      return String(fromPrev);
    }
  }

  const direct = row[header];
  if (direct != null && String(direct).trim() !== '' && String(direct) !== 'false') {
    if (colIndex === 0 || isFormXIXRJSerialHeader(header)) {
      const t = String(direct).trim();
      if (/^\d+$/.test(t)) return t;
      if (/^\d+$/.test(blankSerial)) return blankSerial;
      return String(rowIndex + 1);
    }
    return String(direct);
  }

  if (colIndex === 0 || isFormXIXRJSerialHeader(header)) {
    if (/^\d+$/.test(blankSerial)) return blankSerial;
    if (/^\d+$/.test(underSerial)) return underSerial;
    return String(rowIndex + 1);
  }
  return '';
}

/**
 * Remap row object keys onto canonical Form XIX RJ headers.
 * Handles leading blank Excel spacer columns (Serial values often stored under '').
 */
export function remapFormXIXRJRowsToHeaders(rows, oldHeaders, newHeaders) {
  if (!Array.isArray(rows) || !Array.isArray(newHeaders) || newHeaders.length === 0) return rows || [];
  const rawOld = Array.isArray(oldHeaders) ? oldHeaders.map((h) => (h == null ? '' : h)) : [];
  let leadingBlanks = 0;
  while (leadingBlanks < rawOld.length && isBlankOrGeneratedHeader(rawOld[leadingBlanks])) {
    leadingBlanks += 1;
  }
  const nonBlankOld = rawOld.filter((h) => !isBlankOrGeneratedHeader(h));
  const serialCanon = FORM_XIX_RJ_TABLE_HEADERS[0];
  const nameCanon = FORM_XIX_RJ_TABLE_HEADERS[1];

  return rows.map((row, rowIndex) => {
    if (!row || typeof row !== 'object' || Array.isArray(row)) return row;
    const next = {};
    Object.keys(row).forEach((k) => {
      if (String(k).startsWith('__')) next[k] = row[k];
    });

    const readOld = (oldH) => {
      if (oldH == null) return '';
      if (Object.prototype.hasOwnProperty.call(row, oldH)) return row[oldH];
      if (isBlankOrGeneratedHeader(oldH)) return row[''] ?? row[oldH] ?? '';
      return row[oldH] ?? '';
    };

    // Detect already-shifted rows: blank key holds serial, "Serial No." holds a name.
    const blankSerial = String(row[''] ?? '').trim();
    const underSerial = String(row[serialCanon] ?? row['Serial No'] ?? '').trim();
    const looksShiftedByBlankKey =
      /^\d+$/.test(blankSerial) && underSerial && !/^\d+$/.test(underSerial);

    newHeaders.forEach((newH, newIdx) => {
      let val = '';

      if (looksShiftedByBlankKey) {
        if (newIdx === 0) {
          val = blankSerial;
        } else {
          // Values are one column left of their labels (serial under '', name under Serial No., ...)
          const prevCanon = FORM_XIX_RJ_TABLE_HEADERS[newIdx - 1];
          val = row[prevCanon] ?? '';
          if ((val == null || String(val).trim() === '') && nonBlankOld[newIdx - 1] != null) {
            val = readOld(nonBlankOld[newIdx - 1]);
          }
        }
      } else if (leadingBlanks > 0 && rawOld.length > leadingBlanks) {
        const shiftedOld = rawOld[newIdx + leadingBlanks];
        if (shiftedOld !== undefined) val = readOld(shiftedOld);
      }

      if (val == null || String(val).trim() === '') {
        if (row[newH] != null && String(row[newH]).trim() !== '') {
          val = row[newH];
        } else {
          const newClass = classifyFormXIXRJHeaderIndex(newH);
          if (newClass >= 0) {
            for (let i = 0; i < nonBlankOld.length; i += 1) {
              const oldH = nonBlankOld[i];
              if (classifyFormXIXRJHeaderIndex(oldH) !== newClass) continue;
              const candidate = readOld(oldH);
              if (candidate != null && String(candidate).trim() !== '') {
                val = candidate;
                break;
              }
            }
          }
        }
      }

      if ((val == null || String(val).trim() === '') && newIdx === 0) {
        if (/^\d+$/.test(blankSerial)) val = blankSerial;
        else val = String(rowIndex + 1);
      }

      // Guard: don't leave a person name under Serial No.
      if (newIdx === 0 && val != null && String(val).trim() && !/^\d+$/.test(String(val).trim())) {
        if (/^\d+$/.test(blankSerial)) val = blankSerial;
        else val = String(rowIndex + 1);
      }

      next[newH] = val ?? '';
    });

    // If name column still empty but Serial No. had a non-numeric leftover, keep mapping intact.
    if (!String(next[nameCanon] ?? '').trim() && underSerial && !/^\d+$/.test(underSerial)) {
      next[nameCanon] = underSerial;
    }

    // Drop blank spacer key so later reads don't treat remapped rows as still shifted.
    if (Object.prototype.hasOwnProperty.call(next, '')) delete next[''];

    return next;
  });
}

function readEmpScalar(emp, keys) {
  for (let i = 0; i < keys.length; i += 1) {
    const key = keys[i];
    let raw = emp?.[key];
    if (raw && typeof raw === 'object') {
      raw = raw.displayValue ?? raw.name ?? raw.Name ?? raw.value ?? '';
    }
    const text = String(raw ?? '').trim();
    if (text) return text;
  }
  return '';
}

export function resolveFormXIXRJDesignationAndDepartment(emp = {}) {
  const designation = readEmpScalar(emp, [
    'Designation',
    'Designation.displayValue',
    'DesignationName',
    'Designation Name',
    'JobTitle',
    'Job Title',
  ]);
  const department = readEmpScalar(emp, [
    'Department',
    'Department.displayValue',
    'DepartmentName',
    'Department Name',
    'Dept',
  ]);
  if (designation && department) {
    if (designation.toLowerCase() === department.toLowerCase()) return designation;
    return `${department} / ${designation}`;
  }
  return designation || department || '';
}

function readFormXIXRJPaidDays(payrollRow) {
  if (!payrollRow || payrollRow.fetch_error) return '';
  let p = payrollRow;
  if (typeof payrollRow?.payroll_payload === 'string') {
    try {
      p = { ...payrollRow, ...JSON.parse(payrollRow.payroll_payload) };
    } catch (_) {
      p = payrollRow;
    }
  } else if (payrollRow?.payroll_payload && typeof payrollRow.payroll_payload === 'object') {
    p = { ...payrollRow, ...payrollRow.payroll_payload };
  }
  const flat = flattenPayrollEarningColumns(payrollRow);
  const raw =
    flat.paid_days ??
    flat['paid_days'] ??
    flat['Paid Days'] ??
    flat.paidDays ??
    flat.days_worked ??
    flat['days_worked'] ??
    p.paid_days ??
    p['paid_days'] ??
    p['Paid Days'] ??
    p.paidDays ??
    p.days_worked ??
    payrollRow.paid_days ??
    payrollRow['paid_days'] ??
    '';
  if (raw === '' || raw == null) return '';
  const n = Number(String(raw).replace(/,/g, '').trim());
  return Number.isFinite(n) && n >= 0 ? n : '';
}

export function resolveFormXIXRJNormalHours(payrollRow) {
  const paidDays = readFormXIXRJPaidDays(payrollRow);
  if (paidDays === '') return '';
  return String(paidDays * 8);
}

export function resolveFormXIXRJNetPay(payrollRow) {
  const net = readPayrollNetPayForStatutory(payrollRow);
  return net === '' || net == null ? '' : String(net);
}

/** Normal rate ← gross_pay from Payroll / Zoho pay run. */
export function resolveFormXIXRJGrossPay(payrollRow) {
  const gross = readForm10GrossPayAmount(payrollRow);
  return gross === '' || gross == null ? '' : String(gross);
}

/**
 * Fill Form XIX RJ overtime columns on one employee row.
 * Designation+Department from People; OT date/wages/rate/earnings/payment date → Nil;
 * Normal Hours → paid_days*8; Normal rate → gross_pay; Normal/Total earnings → net_pay.
 */
export function applyFormXIXRJOvertimeAutofillToRow(
  row,
  emp,
  headers,
  payrollRow = null,
  { overwrite = true, sanitizeValue = (v) => String(v ?? '').trim() } = {}
) {
  if (!row || !Array.isArray(headers) || headers.length === 0) return false;
  const cellEmpty = (header) => {
    const v = String(row[header] ?? '').trim();
    return !v || /^n\/?a$/i.test(v) || v === '-' || v === '—';
  };
  const setCell = (header, value) => {
    if (!header || value == null || value === '') return;
    if (!overwrite && !cellEmpty(header)) return;
    row[header] = sanitizeValue(value);
  };

  let changed = false;
  const designationDept = resolveFormXIXRJDesignationAndDepartment(emp || {});
  const normalHours = resolveFormXIXRJNormalHours(payrollRow);
  const grossPay = resolveFormXIXRJGrossPay(payrollRow);
  const netPay = resolveFormXIXRJNetPay(payrollRow);

  headers.forEach((header) => {
    if (isFormXIXRJDesignationDepartmentHeader(header) && designationDept) {
      setCell(header, designationDept);
      changed = true;
      return;
    }
    if (isFormXIXRJOvertimeWorkedDateHeader(header)) {
      setCell(header, FORM_XIX_RJ_NIL);
      changed = true;
      return;
    }
    if (isFormXIXRJWagesOfOvertimeOccasionHeader(header)) {
      setCell(header, FORM_XIX_RJ_NIL);
      changed = true;
      return;
    }
    if (isFormXIXRJNormalHoursHeader(header) && normalHours !== '') {
      setCell(header, normalHours);
      changed = true;
      return;
    }
    if (isFormXIXRJNormalRateHeader(header) && grossPay !== '') {
      setCell(header, grossPay);
      changed = true;
      return;
    }
    if (isFormXIXRJOvertimeRateHeader(header)) {
      setCell(header, FORM_XIX_RJ_NIL);
      changed = true;
      return;
    }
    if (isFormXIXRJNormalEarningsHeader(header) && netPay !== '') {
      setCell(header, netPay);
      changed = true;
      return;
    }
    if (isFormXIXRJOvertimeEarningsHeader(header)) {
      setCell(header, FORM_XIX_RJ_NIL);
      changed = true;
      return;
    }
    if (isFormXIXRJTotalEarningsHeader(header) && netPay !== '') {
      setCell(header, netPay);
      changed = true;
      return;
    }
    if (isFormXIXRJOvertimePaymentDateHeader(header)) {
      setCell(header, FORM_XIX_RJ_NIL);
      changed = true;
    }
  });

  return changed;
}

export function isFormXIXRJOvertimeAutofillContext(formHeader, rowItem, fileName, sheetText = '', tableHeaders = null) {
  if (isFormXIXRajasthanOvertimeRegisterContext(formHeader, rowItem, fileName, sheetText)) return true;
  if (!Array.isArray(tableHeaders) || tableHeaders.length === 0) return false;
  const joined = tableHeaders.map((h) => normHeader(h)).join('\n');
  const looksLikeRjOt =
    /designation/.test(joined) &&
    /department/.test(joined) &&
    /normal\s+hours/.test(joined) &&
    (/wages?\s+of\s+over[\s-]*time\s+on\s+each\s+occasion/.test(joined) ||
      /dates?\s+on\s+which\s+over[\s-]*time\s+work\s+was\s+put\s+in/.test(joined));
  if (!looksLikeRjOt) return false;
  return isFormXIXRajasthanOvertimeRegisterContext(formHeader, rowItem, fileName, joined);
}

/**
 * Flatten modal/grid rows onto canonical Form XIX RJ headers for Excel download.
 * Handles blank-spacer shift and renumbers Serial No. 1..N so export matches the autofill grid.
 */
export function prepareFormXIXRJOvertimeExportRows(rows, headers = null) {
  const canon = resolveFormXIXRJTableHeaders(headers);
  const remapped = remapFormXIXRJRowsToHeaders(
    Array.isArray(rows) ? rows : [],
    Array.isArray(headers) && headers.length > 0 ? headers : canon,
    canon
  );
  const out = [];
  remapped.forEach((row, rowIndex) => {
    if (!row || typeof row !== 'object' || Array.isArray(row)) return;
    const next = {};
    Object.keys(row).forEach((k) => {
      if (String(k).startsWith('__')) next[k] = row[k];
    });
    let hasName = false;
    canon.forEach((header, colIndex) => {
      const val = readFormXIXRJRowCell(row, header, colIndex, rowIndex);
      next[header] = val;
      if (colIndex === 1 && String(val || '').trim() && /[a-zA-Z]{2,}/.test(String(val))) {
        hasName = true;
      }
    });
    next[canon[0]] = String(out.length + 1);
    const meaningful = Object.values(next).some(
      (v) => v != null && String(v).trim() !== '' && !String(v).startsWith('__')
    );
    if (meaningful || hasName) out.push(next);
  });
  return out;
}

function excelJsPlainText(cell) {
  const v = cell?.value;
  if (v == null) return '';
  if (typeof v === 'string' || typeof v === 'number' || typeof v === 'boolean') return String(v).trim();
  if (typeof v === 'object') {
    if (Array.isArray(v.richText)) return v.richText.map((rt) => rt?.text || '').join('').trim();
    if (v.text != null) return String(v.text).trim();
    if (v.result != null) return String(v.result).trim();
  }
  return String(v).trim();
}

function unmergeExcelJsIntersectingRange(worksheet, rowFrom1, rowTo1, colFrom1, colTo1) {
  const merges = worksheet?.model?.merges;
  if (!Array.isArray(merges) || merges.length === 0) return;
  const toRemove = [];
  for (const range of merges) {
    const parts = String(range || '').split(':');
    if (parts.length !== 2) continue;
    const start = parts[0].match(/^([A-Z]+)(\d+)$/i);
    const end = parts[1].match(/^([A-Z]+)(\d+)$/i);
    if (!start || !end) continue;
    const mr1 = parseInt(start[2], 10);
    const mr2 = parseInt(end[2], 10);
    const colLetterTo1 = (letters) => {
      let n = 0;
      const s = String(letters || '').toUpperCase();
      for (let i = 0; i < s.length; i += 1) n = n * 26 + (s.charCodeAt(i) - 64);
      return n;
    };
    const mc1 = colLetterTo1(start[1]);
    const mc2 = colLetterTo1(end[1]);
    if (mr2 < rowFrom1 || mr1 > rowTo1) continue;
    if (mc2 < colFrom1 || mc1 > colTo1) continue;
    toRemove.push(range);
  }
  toRemove.forEach((range) => {
    try {
      worksheet.unMergeCells(range);
    } catch (_) {
      /* ignore */
    }
  });
}

function copyExcelJsCellValueLeftToRight(worksheet, row1, fromCol1, toCol1) {
  const src = worksheet.getCell(row1, fromCol1);
  const dst = worksheet.getCell(row1, toCol1);
  if (excelJsPlainText(dst)) return false;
  const srcText = excelJsPlainText(src);
  if (!srcText) return false;
  dst.value = src.value;
  if (src.alignment && typeof src.alignment === 'object') {
    dst.alignment = { ...src.alignment };
  }
  src.value = null;
  return true;
}

/**
 * Official Form_XIX_RJ.xlsx: empty column A, Serial No. in B, Name in C.
 * Broken downloads put Serial in A with a blank gap at B — slide A → B when B is empty.
 */
export function compactFormXIXRJGapColumnsOnExcelJs(worksheet, headerRow1, throughRow1) {
  if (!worksheet || headerRow1 < 1) return false;
  const a = excelJsPlainText(worksheet.getCell(headerRow1, 1));
  const b = excelJsPlainText(worksheet.getCell(headerRow1, 2));
  const c = excelJsPlainText(worksheet.getCell(headerRow1, 3));
  const gap =
    isFormXIXRJSerialHeader(a) &&
    !b &&
    (isFormXIXRJWorkmanNameHeader(c) || /name\s+of\s+workman/i.test(c));
  if (!gap) return false;
  const last = Math.max(headerRow1, Number(throughRow1) || headerRow1);
  unmergeExcelJsIntersectingRange(worksheet, headerRow1, last, 1, 2);
  for (let r = headerRow1; r <= last; r += 1) {
    copyExcelJsCellValueLeftToRight(worksheet, r, 1, 2);
  }
  return true;
}

export function findFormXIXRJExcelJsTableLayout(worksheet) {
  if (!worksheet) return null;
  let nameRow = -1;
  let nameCol = -1;
  let serialCol = -1;
  const maxR = Math.min(40, Math.max(Number(worksheet.rowCount) || 0, 20));
  const maxC = 20;
  for (let r = 1; r <= maxR; r += 1) {
    for (let c = 1; c <= maxC; c += 1) {
      const t = excelJsPlainText(worksheet.getCell(r, c));
      if (!t) continue;
      if (nameRow < 0 && isFormXIXRJWorkmanNameHeader(t)) {
        nameRow = r;
        nameCol = c;
      }
      if (serialCol < 0 && isFormXIXRJSerialHeader(t)) serialCol = c;
    }
    if (nameRow > 0 && serialCol > 0) break;
  }
  if (nameRow < 1) return null;
  const headerRow1 = nameRow;
  let tableStartCol1 =
    serialCol > 0 && serialCol < nameCol ? serialCol : Math.max(2, nameCol - 1);
  if (tableStartCol1 < 2) tableStartCol1 = 2;
  let numberRow1 = headerRow1 + 1;
  const colCount = FORM_XIX_RJ_TABLE_HEADERS.length;
  for (let r = headerRow1; r <= headerRow1 + 3; r += 1) {
    let hits = 0;
    for (let j = 0; j < colCount; j += 1) {
      const t = String(excelJsPlainText(worksheet.getCell(r, tableStartCol1 + j)) || '').replace(/\s+/g, '');
      if (t === String(j + 1)) hits += 1;
    }
    if (hits >= 8) {
      numberRow1 = r;
      break;
    }
  }
  const dataStartRow1 = Math.max(headerRow1 + 2, numberRow1 + 1);
  return {
    headerRow1,
    numberRow1,
    dataStartRow1,
    tableStartCol1,
    tableStartCol0: tableStartCol1 - 1,
    colTo1: tableStartCol1 + colCount - 1,
  };
}

export function restoreFormXIXRJExcelJsVerticalTableHeaders(worksheet, layout) {
  if (!worksheet || !layout?.headerRow1) return;
  const row = worksheet.getRow(layout.headerRow1);
  row.hidden = false;
  row.height = Math.max(Number(row.height) || 0, 78);
  const c0 = Math.max(2, layout.tableStartCol1 || 2);
  const c1 = Math.max(c0, layout.colTo1 || c0 + 14);
  for (let c = c0; c <= c1; c += 1) {
    const cell = worksheet.getCell(layout.headerRow1, c);
    if (!excelJsPlainText(cell) && c > c0 + 2) continue;
    const cur = cell.alignment && typeof cell.alignment === 'object' ? cell.alignment : {};
    cell.alignment = {
      ...cur,
      textRotation: 90,
      wrapText: true,
      horizontal: 'center',
      vertical: 'middle',
    };
  }
}

export function worksheetLooksLikeFormXIXRJOvertime(worksheet) {
  if (!worksheet) return false;
  let sawFormXix = false;
  let sawOvertime = false;
  let sawWorkman = false;
  const maxR = Math.min(20, Math.max(Number(worksheet.rowCount) || 0, 8));
  for (let r = 1; r <= maxR; r += 1) {
    for (let c = 1; c <= 16; c += 1) {
      const t = excelJsPlainText(worksheet.getCell(r, c)).toLowerCase();
      if (!t) continue;
      if (/^form\s*xix\b/.test(t) || t === 'form xix') sawFormXix = true;
      if (/register\s+of\s+over[\s-]*time/.test(t) || /see\s+rule\s*77/.test(t)) sawOvertime = true;
      if (/name\s+of\s+workman/.test(t)) sawWorkman = true;
    }
  }
  return sawFormXix && (sawOvertime || sawWorkman);
}

/** Keep official Form XIX RJ boxes: empty A, vertical B–P headers, no Serial/Name gap. */
export function applyFormXIXRJOriginalTemplateLayoutToExcelJs(worksheet) {
  if (!worksheetLooksLikeFormXIXRJOvertime(worksheet)) return null;
  let layout = findFormXIXRJExcelJsTableLayout(worksheet) || {
    headerRow1: 13,
    numberRow1: 14,
    dataStartRow1: 15,
    tableStartCol1: 2,
    tableStartCol0: 1,
    colTo1: 16,
  };
  const through = Math.max(
    layout.dataStartRow1 + 40,
    Number(worksheet.actualRowCount) || 0,
    Number(worksheet.rowCount) || 0,
    40
  );
  compactFormXIXRJGapColumnsOnExcelJs(worksheet, layout.headerRow1, through);
  layout = findFormXIXRJExcelJsTableLayout(worksheet) || layout;
  layout.tableStartCol1 = Math.max(2, layout.tableStartCol1 || 2);
  layout.tableStartCol0 = layout.tableStartCol1 - 1;
  layout.colTo1 = layout.tableStartCol1 + FORM_XIX_RJ_TABLE_HEADERS.length - 1;
  restoreFormXIXRJExcelJsVerticalTableHeaders(worksheet, layout);
  return layout;
}

const FORM_XIX_RJ_HEADER_VALUE_KINDS = [
  {
    test: (t) => /name\s+and\s+address\s+of\s+(?:the\s+)?contractor/i.test(t) && !/principal/.test(t),
    field: { key: 'form_xxiii_contractor', label: 'Name and address of the Contractor' },
  },
  {
    test: (t) => /(?:name|nature)\s+and\s+location\s+of\s+work/i.test(t),
    field: { key: 'form_xxiii_nature_location_work', label: 'Name and location of work' },
  },
  {
    test: (t) => /establishment.*contract\s+is\s+carried/i.test(t),
    field: { key: 'form_xxiii_establishment_contract_carried', label: 'Name and address of Establishment' },
  },
  {
    test: (t) => /principal\s+employer/i.test(t),
    field: { key: 'statutory_principal_employer', label: 'Name and address of Principal Employer' },
  },
];

function placeFormXIXRJValueOnOriginalLabel(cell, value) {
  const raw = excelJsPlainText(cell);
  const val = String(value || '').trim();
  if (!raw || !val) return;
  if (raw.includes(val)) return;
  const dots = raw.match(/^(.*?)(\.{3,})\s*(.*)$/);
  if (dots) {
    cell.value = `${dots[1]}${dots[2]} ${val}`;
    return;
  }
  if (/:/.test(raw)) {
    const label = raw.split(':')[0].replace(/\.+$/, '').trim();
    cell.value = `${label}: ${val}`;
    return;
  }
  cell.value = `${raw.replace(/\.+$/, '').trim()} ${val}`;
}

/**
 * Write contractor / work / establishment / employer values onto the original dotted
 * labels. Never replace official wording or merge those rows.
 */
export function writeFormXIXRJHeaderValuesOntoOriginalLabels(
  worksheet,
  headerFormData,
  parsedFormHeader,
  headerRowEnd1
) {
  if (!worksheet || !headerFormData || typeof headerFormData !== 'object') return;
  const rowEnd = Math.max(1, Number(headerRowEnd1) || 12);
  const parsedFields = Array.isArray(parsedFormHeader?.fields) ? parsedFormHeader.fields : [];
  const written = new Set();
  for (let r = 1; r <= rowEnd; r += 1) {
    for (let c = 1; c <= 16; c += 1) {
      const cell = worksheet.getCell(r, c);
      const raw = excelJsPlainText(cell);
      if (!raw) continue;
      for (const kind of FORM_XIX_RJ_HEADER_VALUE_KINDS) {
        if (!kind.test(raw) || written.has(kind.field.key)) continue;
        let val = resolveHeaderFieldExportValue(headerFormData, kind.field);
        if (!val) {
          const parsed = parsedFields.find((f) => kind.test(String(f?.label || '')));
          if (parsed) val = resolveHeaderFieldExportValue(headerFormData, parsed);
        }
        if (val) {
          placeFormXIXRJValueOnOriginalLabel(cell, val);
          written.add(kind.field.key);
        }
        break;
      }
    }
  }
}

export { FORM_XIX_RJ_NIL };
