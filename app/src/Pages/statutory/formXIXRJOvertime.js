import { flattenPayrollEarningColumns, readPayrollNetPayForStatutory } from '../../utils/payrollEarnings';
import { isFormXIXRajasthanOvertimeRegisterContext } from './formXIXAPWageSlip';

/** Rajasthan Form XIX — Register of Overtime [Rule 77(2)(e)] autofill helpers. */

const FORM_XIX_RJ_NIL = 'Nil';

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

/**
 * Fill Form XIX RJ overtime columns on one employee row.
 * Designation+Department from People; OT date/wages → Nil; Normal Hours → paid_days*8;
 * Normal/Total earnings → net_pay.
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
    if (isFormXIXRJNormalEarningsHeader(header) && netPay !== '') {
      setCell(header, netPay);
      changed = true;
      return;
    }
    if (isFormXIXRJTotalEarningsHeader(header) && netPay !== '') {
      setCell(header, netPay);
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

export { FORM_XIX_RJ_NIL };
