/** Karnataka Form F — Register of Leave with Wages (Rule 8): header fields + PART I earned leave table. */

import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import {
  countExcelJSTemplateBodyRows,
  ensureExcelJSDataRowsWithBorders,
} from '../../utils/excelTableBorders';
import { flattenPayrollEarningColumns, readPayrollScalar } from '../../utils/payrollEarnings';
import {
  buildLeaveRecordLookupMap,
  findLeaveRecordForFormRow,
  getForm15EarnedLeavePeriodMetrics,
} from '../../utils/leaveMetrics';
import { resolveFormXIXMPPayrollRowForEmployee } from './formXIXMPWageSlip';

export function formFKarnatakaHeaderNorm(txt) {
  return String(txt || '')
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

export function isFormFLeaveWithWagesContext(formHeader, rowItem, fileName, sheetText = '') {
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
  ]
    .filter((x) => x != null && String(x).trim() !== '')
    .join(' ')
    .toLowerCase();

  const hasFormF = /\bform[\s._-]*f\b/i.test(parts);
  const hasLeaveRegister =
    /leave\s+with\s+wages/i.test(parts) ||
    /register\s+of\s+leave/i.test(parts);
  if (hasFormF && hasLeaveRegister) return true;
  if (hasFormF && /karnataka/.test(parts)) return true;
  return false;
}

export function isFormFKarnatakaTableLayoutFormHeader(formHeader) {
  return !!formHeader?.formFKarnatakaTableLayout;
}

/** PART I — Earned Leave columns (Karnataka Form F template). */
export const FORM_F_KARNATAKA_PART_I_HEADERS = [
  'Days worked From',
  'Days worked To',
  'Total days worked',
  'Leave earned',
  'Leave at credit (incl balance if any)',
  'Leave availed From',
  'Leave availed To',
  'Leave availed No. of days',
  'Balance on return from leave',
  'Date on which wages for leave paid',
  'Remarks',
];

/** PART II — Sick/Accident Leave (with pay) columns. */
export const FORM_F_KARNATAKA_PART_II_TITLE = 'PART II - Sick/Accident Leave (with pay)';
export const FORM_F_KARNATAKA_PART_II_HEADERS = [
  'PART II Year',
  'PART II Sick/Accident leave of credit',
  'PART II Sick/Accident leave Availed',
  'PART II Balance at the end of the year',
];

export const FORM_F_KARNATAKA_PART_II_DISPLAY_LABELS = [
  'Year',
  'Sick/Accident leave of credit',
  'Availed',
  'Balance at the end of the year',
];

/** April PART I defaults when payroll/leave APIs lack values (Karnataka Form F). */
export const FORM_F_KA_APR_DEFAULT_PART_I = {
  totalDaysWorked: '30',
  leaveEarned: '0',
  leaveAtCredit: '0',
};

export function isFormFKarnatakaAprilMonthCandidates(monthCandidates) {
  const primary = String(
    (Array.isArray(monthCandidates) ? monthCandidates[0] : monthCandidates) || ''
  ).trim();
  return /-04$/.test(primary);
}

export function applyFormFKarnatakaAprilPartIDefaultsToRow(
  row,
  monthCandidates,
  { overwrite = true } = {}
) {
  if (!row || !isFormFKarnatakaAprilMonthCandidates(monthCandidates)) return 0;
  const canon = FORM_F_KARNATAKA_PART_I_HEADERS;
  const defaults = FORM_F_KA_APR_DEFAULT_PART_I;
  let applied = 0;
  const setByIndex = (idx, value) => {
    if (value == null || value === '') return;
    const key = canon[idx];
    if (!key) return;
    if (!overwrite && String(row[key] ?? '').trim() !== '') return;
    row[key] = String(value);
    applied += 1;
  };
  setByIndex(2, defaults.totalDaysWorked);
  setByIndex(3, defaults.leaveEarned);
  const credit =
    defaults.leaveAtCredit != null && String(defaults.leaveAtCredit).trim() !== ''
      ? String(defaults.leaveAtCredit).trim()
      : defaults.leaveEarned;
  setByIndex(4, credit);
  const balanceOnReturn = computeFormFKarnatakaBalanceOnReturn(defaults.leaveEarned, '0');
  if (balanceOnReturn !== '') setByIndex(8, balanceOnReturn);
  return applied;
}

const PART_I_SUBHEADER_FRAGMENTS = [
  /^from$/i,
  /^to$/i,
  /total\s+days\s+worked/i,
  /leave\s+earned/i,
  /leave\s+at\s+credit/i,
  /no\.?\s*of\s+days/i,
  /balance\s+on\s+return\s+from\s+leave/i,
  /date\s+on\s+which\s+wages\s+for\s+leave\s+paid/i,
  /^remarks?$/i,
];

export function headersIndicateFormFKarnatakaPartITable(tableHeaders) {
  const joined = (Array.isArray(tableHeaders) ? tableHeaders : [])
    .map((h) => formFKarnatakaHeaderNorm(h))
    .join('\n');
  if (!joined) return false;
  if (/sl\s*no\s+in\s+the\s+register\s+of\s+adult/.test(joined) && !/leave\s+earned/.test(joined)) {
    return false;
  }
  return (
    /leave\s+earned/.test(joined) &&
    /balance\s+on\s+return\s+from\s+leave/.test(joined) &&
    (/total\s+days\s+worked/.test(joined) || /days\s+worked/.test(joined))
  );
}

export function resolveFormFKarnatakaTableHeaders(tableHeaders) {
  // Always use unique canonical keys — Excel sub-headers repeat bare "From"/"To" and break row storage.
  return [...FORM_F_KARNATAKA_PART_I_HEADERS];
}

/** Remap a grid row from parsed Excel headers (possibly duplicate "From"/"To") to canonical Form F keys. */
export function remapFormFKarnatakaRowToCanonicalHeaders(row, priorHeaders) {
  if (!row || typeof row !== 'object') return row;
  const canon = FORM_F_KARNATAKA_PART_I_HEADERS;
  const prior = Array.isArray(priorHeaders) ? priorHeaders : [];
  const out = { ...row };
  const daysWorkedFrom = String(row[canon[0]] ?? row['Days worked From'] ?? '').trim();
  const daysWorkedTo = String(row[canon[1]] ?? row['Days worked To'] ?? '').trim();
  canon.forEach((key, idx) => {
    let val = row[key];
    if ((val == null || String(val).trim() === '') && prior[idx] != null) {
      const priorKey = prior[idx];
      const priorNorm = formFKarnatakaHeaderNorm(priorKey);
      const isLeaveAvailedIdx = idx >= 5 && idx <= 7;
      const priorIsDaysWorked =
        priorNorm.includes('days worked') ||
        (isLeaveAvailedIdx && (priorNorm === 'from' || priorNorm === 'to'));
      if (
        priorKey &&
        priorKey !== key &&
        !priorIsDaysWorked &&
        row[priorKey] != null &&
        String(row[priorKey]).trim() !== ''
      ) {
        val = row[priorKey];
      }
    }
    if (isFormFKarnatakaLeaveAvailedColumnIndex(idx) && val != null) {
      const text = String(val).trim();
      if (text && (text === daysWorkedFrom || text === daysWorkedTo)) {
        val = '';
      }
    }
    out[key] = val != null ? val : '';
  });
  return out;
}

export function remapFormFKarnatakaRowsToCanonicalHeaders(rows, priorHeaders) {
  if (!Array.isArray(rows)) return [];
  return rows.map((row) => remapFormFKarnatakaRowToCanonicalHeaders(row, priorHeaders));
}

const normalizeCell = (txt) =>
  String(txt || '')
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

const isPartIEarnedLeaveBanner = (text) => {
  const t = formFKarnatakaHeaderNorm(text);
  return /part\s*[-–]?\s*i/.test(t) && /earned\s+leave/.test(t);
};

const isEmployeeIdentityBannerRow = (rowCells) => {
  const joined = rowCells.map((c) => formFKarnatakaHeaderNorm(c)).join(' ');
  return (
    /sl\s*no\s+in\s+the\s+register\s+of\s+adult/.test(joined) &&
    (/name\s+of\s+the\s+person/.test(joined) || /date\s+of\s+entry\s+into\s+service/.test(joined))
  );
};

const rowLooksLikePartISubHeaderRow = (rowCells) => {
  const nonEmpty = rowCells.map((c) => normalizeCell(c)).filter(Boolean);
  if (nonEmpty.length < 4) return false;
  const joined = nonEmpty.join(' ').toLowerCase();
  if (isEmployeeIdentityBannerRow(nonEmpty)) return false;
  let hits = 0;
  nonEmpty.forEach((cell) => {
    if (PART_I_SUBHEADER_FRAGMENTS.some((re) => re.test(cell))) hits += 1;
  });
  return hits >= 3 && /leave\s+earned/.test(joined);
};

const rowLooksLikePartINumberedIndexRow = (rowCells) => {
  let nums = 0;
  let total = 0;
  rowCells.forEach((cell) => {
    const t = normalizeCell(cell);
    if (!t) return;
    total += 1;
    if (/^\d+$/.test(t)) nums += 1;
  });
  return total >= 6 && nums >= 6 && nums / total >= 0.55;
};

const combinePartIHeaderBand = (upperCells, lowerCells) => {
  const maxLen = Math.max(upperCells.length, lowerCells.length, 11);
  const headers = [];
  let currentGroup = '';
  for (let c = 0; c < maxLen; c += 1) {
    const upper = normalizeCell(upperCells[c]);
    const lower = normalizeCell(lowerCells[c]);
    if (upper && !/^\d+$/.test(upper)) {
      if (/no\s+of\s+days\s+worked/i.test(upper)) currentGroup = 'Days worked';
      else if (/leave\s+availed/i.test(upper)) currentGroup = 'Leave availed';
      else if (/leave\s+earned/i.test(upper)) currentGroup = 'Leave earned';
      else if (/leave\s+at\s+credit/i.test(upper)) currentGroup = 'Leave at credit';
      else if (/balance\s+on\s+return/i.test(upper)) currentGroup = 'Balance on return from leave';
      else if (/date\s+on\s+which\s+wages/i.test(upper)) currentGroup = 'Date on which wages for leave paid';
      else if (/^remarks?$/i.test(upper)) currentGroup = 'Remarks';
    }
    let label = lower || upper;
    if (/^\d+$/.test(label)) {
      const canonical = FORM_F_KARNATAKA_PART_I_HEADERS[parseInt(label, 10) - 1];
      if (canonical) label = canonical;
    }
    if (!label) continue;
    if (/^from$/i.test(label) && currentGroup === 'Days worked') label = 'Days worked From';
    else if (/^to$/i.test(label) && currentGroup === 'Days worked') label = 'Days worked To';
    else if (/^from$/i.test(label) && currentGroup === 'Leave availed') label = 'Leave availed From';
    else if (/^to$/i.test(label) && currentGroup === 'Leave availed') label = 'Leave availed To';
    else if (/no\.?\s*of\s+days/i.test(label) && currentGroup === 'Leave availed') {
      label = 'Leave availed No. of days';
    }
    headers.push(label);
  }
  return headers.filter(Boolean);
};

/**
 * Re-point parseExcelForm to PART I earned leave columns instead of the employee identity row.
 */
export function rebuildFormFKarnatakaPartIHeaders({
  jsonData,
  effectiveSheetCols,
  getMergedAwareCellText,
  headers,
  expandedHeaders,
  headerRowIndex,
  startIndex,
  tableData,
  firstSheetName,
  formHeaderInfo,
  parseHints = {},
}) {
  const fileName = parseHints?.fileName || parseHints?.formFileName || firstSheetName || '';
  const sheetTextBlob = (Array.isArray(jsonData) ? jsonData : [])
    .slice(0, 45)
    .map((row) => (Array.isArray(row) ? row : []).map((c) => String(c || '')).join(' '))
    .join(' ');
  if (!isFormFLeaveWithWagesContext(formHeaderInfo, parseHints?.item || null, fileName, sheetTextBlob)) {
    return null;
  }

  const maxCols = Math.max(effectiveSheetCols || 0, 20);
  const readRow = (r) => {
    const cells = [];
    for (let c = 0; c < maxCols; c += 1) {
      cells.push(normalizeCell(getMergedAwareCellText(r, c)));
    }
    return cells;
  };

  let partIRow = -1;
  for (let r = 0; r < Math.min(45, jsonData.length); r += 1) {
    const line = readRow(r).filter(Boolean).join(' ');
    if (isPartIEarnedLeaveBanner(line)) {
      partIRow = r;
      break;
    }
  }

  let subHeaderRow = -1;
  let groupHeaderRow = -1;
  const scanFrom = partIRow >= 0 ? partIRow + 1 : 0;
  const scanTo = Math.min(jsonData.length - 1, scanFrom + 12);
  for (let r = scanFrom; r <= scanTo; r += 1) {
    const cells = readRow(r);
    if (isEmployeeIdentityBannerRow(cells)) continue;
    if (rowLooksLikePartISubHeaderRow(cells)) {
      subHeaderRow = r;
      if (r > 0 && !rowLooksLikePartISubHeaderRow(readRow(r - 1))) {
        groupHeaderRow = r - 1;
      }
      break;
    }
  }

  if (subHeaderRow < 0) {
    for (let r = scanFrom; r <= scanTo; r += 1) {
      const cells = readRow(r);
      if (rowLooksLikePartINumberedIndexRow(cells)) {
        subHeaderRow = r;
        if (r > 0) groupHeaderRow = r - 1;
        break;
      }
    }
  }

  let rebuiltHeaders = [...FORM_F_KARNATAKA_PART_I_HEADERS];
  let newHeaderRowIndex = subHeaderRow;
  let newStartIndex = subHeaderRow >= 0 ? subHeaderRow + 1 : startIndex;

  if (subHeaderRow >= 0) {
    const lowerCells = readRow(subHeaderRow);
    const upperCells = groupHeaderRow >= 0 ? readRow(groupHeaderRow) : [];
    const combined = combinePartIHeaderBand(upperCells, lowerCells);
    if (combined.length >= 6 && headersIndicateFormFKarnatakaPartITable(combined)) {
      rebuiltHeaders = combined;
    }
    while (newStartIndex < jsonData.length) {
      const cells = readRow(newStartIndex);
      const nonEmpty = cells.filter(Boolean);
      // Do not skip empty bordered PART I body rows — data must start in the 15-row grid.
      if (rowLooksLikePartISubHeaderRow(cells) || rowLooksLikePartINumberedIndexRow(cells)) {
        newStartIndex += 1;
        continue;
      }
      if (
        nonEmpty.length > 0 &&
        (isPartIEarnedLeaveBanner(nonEmpty.join(' ')) ||
          /part\s*[-–]?\s*ii/i.test(nonEmpty.join(' ').toLowerCase()))
      ) {
        break;
      }
      break;
    }
  }

  const newTableData = [];
  for (let i = newStartIndex; i < jsonData.length; i += 1) {
    const cells = readRow(i);
    const line = cells.filter(Boolean).join(' ');
    if (!line) continue;
    if (/part\s*[-–]?\s*ii/i.test(formFKarnatakaHeaderNorm(line))) break;
    const rowData = {};
    FORM_F_KARNATAKA_PART_I_HEADERS.forEach((header, idx) => {
      rowData[header] = cells[idx] || '';
    });
    const hasData = Object.values(rowData).some((v) => String(v ?? '').trim() !== '');
    if (hasData) newTableData.push(rowData);
  }

  const canonicalHeaders = [...FORM_F_KARNATAKA_PART_I_HEADERS];

  return {
    headers: canonicalHeaders,
    expandedHeaders: canonicalHeaders,
    subColumnsData: {},
    headerRowIndex: newHeaderRowIndex >= 0 ? newHeaderRowIndex : headerRowIndex,
    startIndex: newStartIndex,
    tableData: newTableData,
    columnGroupLabels: null,
    originalHeaderRowIndex: newHeaderRowIndex >= 0 ? newHeaderRowIndex : headerRowIndex,
    formHeaderPatch: {
      title: formHeaderInfo?.title || 'FORM F',
      subtitle: formHeaderInfo?.subtitle || 'REGISTER OF LEAVE WITH WAGES',
      reference: formHeaderInfo?.reference || '(SEE RULE 8)',
      formFKarnatakaTableLayout: true,
      formFKarnatakaPartIiLayout: true,
      formFKarnatakaPartIiTitle: FORM_F_KARNATAKA_PART_II_TITLE,
      formFKarnatakaPartIiHeaders: [...FORM_F_KARNATAKA_PART_II_HEADERS],
    },
  };
}

export function isFormFKarnatakaLeaveEarnedHeader(header) {
  const s = formFKarnatakaHeaderNorm(header);
  return s.includes('leave') && s.includes('earned') && !s.includes('credit');
}

export function isFormFKarnatakaLeaveAvailedHeader(header) {
  const s = formFKarnatakaHeaderNorm(header);
  return s.includes('leave') && s.includes('avail');
}

export function isFormFKarnatakaLeaveBalanceHeader(header) {
  const s = formFKarnatakaHeaderNorm(header);
  return s.includes('balance') && s.includes('return') && s.includes('leave');
}

export function isFormFKarnatakaLeaveAtCreditHeader(header) {
  const s = formFKarnatakaHeaderNorm(header);
  return s.includes('leave') && s.includes('credit');
}

export function isFormFKarnatakaDaysWorkedFromHeader(header) {
  const s = formFKarnatakaHeaderNorm(header);
  return s.includes('days worked') && s.includes('from');
}

export function isFormFKarnatakaDaysWorkedToHeader(header) {
  const s = formFKarnatakaHeaderNorm(header);
  return s.includes('days worked') && s.includes('to') && !s.includes('from');
}

/**
 * Balance count / Balance on return from leave =
 * Leave earned during the Period − approved LeaveCount (0 when no leave availed).
 */
export function computeFormFKarnatakaLeaveAtCredit(earnedDuring, leaveCount) {
  const earned = Number(String(earnedDuring ?? '').replace(/,/g, '').trim());
  if (!Number.isFinite(earned)) return '';
  const countRaw = leaveCount == null || leaveCount === '' ? '0' : leaveCount;
  const count = Number(String(countRaw).replace(/,/g, '').trim());
  const availed = Number.isFinite(count) ? count : 0;
  const credit = Math.round((earned - availed) * 100) / 100;
  return String(credit);
}

/** Alias: Balance count for Form F col "Balance on return from leave". */
export function computeFormFKarnatakaBalanceOnReturn(earnedDuring, leaveCount) {
  return computeFormFKarnatakaLeaveAtCredit(earnedDuring, leaveCount);
}

export function applyFormFKarnatakaMonthDatesToRow(row, fromDate, toDate, { overwrite = true } = {}) {
  if (!row) return 0;
  const canon = FORM_F_KARNATAKA_PART_I_HEADERS;
  let applied = 0;
  const setByIndex = (idx, value) => {
    if (value == null || value === '') return;
    const key = canon[idx];
    if (!key) return;
    if (!overwrite && String(row[key] ?? '').trim() !== '') return;
    row[key] = String(value);
    applied += 1;
  };
  setByIndex(0, String(fromDate || '').trim());
  setByIndex(1, String(toDate || '').trim());
  return applied;
}

export function applyFormFKarnatakaEarnedLeaveToRow(
  row,
  leaveRecord,
  leaveTypeLabels,
  approvedRecord,
  { overwrite = true } = {}
) {
  if (!row) return 0;
  const canon = FORM_F_KARNATAKA_PART_I_HEADERS;
  let applied = 0;

  const setByIndex = (idx, value) => {
    if (value == null || value === '') return;
    const key = canon[idx];
    if (!key) return;
    if (!overwrite && String(row[key] ?? '').trim() !== '') return;
    row[key] = String(value);
    applied += 1;
  };

  // Leave earned ← Leave earned during the Period (Zoho paidBalance / LeaveData).
  let earnedDuring = String(row[canon[3]] ?? '').trim();
  if (leaveRecord) {
    const { periodBalance } = getForm15EarnedLeavePeriodMetrics(leaveRecord, leaveTypeLabels);
    if (periodBalance !== '') {
      earnedDuring = periodBalance;
      setByIndex(3, earnedDuring);
      // Leave at credit includes balance on last occasion → same Leave-earned figure.
      setByIndex(4, earnedDuring);
    }
  }

  // Balance on return from leave ← Balance count (earned − approved LeaveCount).
  if (earnedDuring) {
    const leaveCount = approvedRecord
      ? parseApprovedLeaveDaysCount(approvedRecord.Days ?? approvedRecord.days)
      : '';
    const balanceCount = computeFormFKarnatakaBalanceOnReturn(
      earnedDuring,
      leaveCount !== '' ? leaveCount : '0'
    );
    if (balanceCount !== '') setByIndex(8, balanceCount);
  }

  return applied;
}

export function applyFormFKarnatakaLeaveEarnedAutofill(
  mappedData,
  employeesForMapping,
  leaveRecords,
  leaveTypeLabels,
  tableHeaders,
  options = {}
) {
  if (!Array.isArray(mappedData) || mappedData.length === 0) return 0;

  const priorHeaders = Array.isArray(tableHeaders) ? tableHeaders : [];
  const canonicalHeaders = resolveFormFKarnatakaTableHeaders(tableHeaders);
  const leaveLookup = buildLeaveRecordLookupMap(leaveRecords);
  const approvedLeaveRecords = filterApprovedLeaveRecordsForFormFKarnataka(
    Array.isArray(options.approvedLeaveRecords) ? options.approvedLeaveRecords : [],
    options.monthFrom || options.fromDate || '',
    options.monthTo || options.toDate || ''
  );
  const approvedLookup =
    approvedLeaveRecords.length > 0 ? buildApprovedLeaveLookupMap(approvedLeaveRecords) : null;
  const { employeeNameHeader, employeeIdHeader } = options.resolveEmployeeHeaders
    ? options.resolveEmployeeHeaders(canonicalHeaders)
    : { employeeNameHeader: '', employeeIdHeader: '' };

  const { fromDate = '', toDate = '' } = options;

  let hits = 0;
  mappedData.forEach((row, index) => {
    const normalizedRow = remapFormFKarnatakaRowToCanonicalHeaders(row, priorHeaders);
    Object.assign(row, normalizedRow);

    if (fromDate && toDate) {
      applyFormFKarnatakaMonthDatesToRow(row, fromDate, toDate, {
        overwrite: options.overwrite !== false,
      });
    }

    const empItem = employeesForMapping[index];
    const emp = empItem && (empItem.Employee || empItem.employee || empItem);
    const leaveRecord = findLeaveRecordForFormRow(
      leaveLookup,
      row,
      employeeIdHeader,
      employeeNameHeader
    );
    let approvedRecord = null;
    // Only attach Leave availed when LeaveData shows booked days for Earned Leave.
    const { periodBooked } = leaveRecord
      ? getForm15EarnedLeavePeriodMetrics(leaveRecord, leaveTypeLabels)
      : { periodBooked: '' };
    const hasBookedLeave =
      periodBooked !== '' && Number(String(periodBooked).replace(/,/g, '')) > 0;
    if (approvedLookup && hasBookedLeave) {
      approvedRecord = findApprovedLeaveForEmployee(
        approvedLookup,
        emp,
        row,
        employeeIdHeader,
        employeeNameHeader,
        {
          approvedRecords: approvedLeaveRecords,
          monthFrom: options.monthFrom || fromDate,
          monthTo: options.monthTo || toDate,
          strictIdentity: true,
          collectEmployeeNameCandidates: options.collectEmployeeNameCandidates,
          collectEmployeeIdCandidates: options.collectEmployeeIdCandidates,
        }
      );
    }

    const applied = applyFormFKarnatakaEarnedLeaveToRow(
      row,
      leaveRecord,
      leaveTypeLabels,
      approvedRecord,
      { overwrite: options.overwrite !== false }
    );
    if (options.monthCandidates) {
      applyFormFKarnatakaAprilPartIDefaultsToRow(row, options.monthCandidates, {
        overwrite: options.overwrite !== false,
      });
    }
    if (applied > 0 || (fromDate && toDate)) hits += 1;
  });
  return hits;
}

function parseJsonMaybe(val) {
  if (val == null) return null;
  if (typeof val === 'object') return val;
  const s = String(val).trim();
  if (!s.startsWith('{') && !s.startsWith('[')) return null;
  try {
    return JSON.parse(s);
  } catch (_) {
    return null;
  }
}

/** Sum LeaveCount values from Zoho approved-leave `Days` field (object keyed by date). */
export function parseApprovedLeaveDaysCount(daysField) {
  const obj = parseJsonMaybe(daysField);
  if (!obj || typeof obj !== 'object' || Array.isArray(obj)) {
    return '';
  }
  let total = 0;
  let hasCount = false;
  Object.values(obj).forEach((entry) => {
    if (!entry || typeof entry !== 'object') return;
    const raw = entry.LeaveCount ?? entry.leaveCount ?? entry.count;
    if (raw == null || raw === '') return;
    const n = Number(raw);
    // Skip weekly-off / session placeholders with LeaveCount 0.0
    if (Number.isFinite(n) && n > 0) {
      total += n;
      hasCount = true;
    }
  });
  return hasCount ? String(total) : '';
}

function isZohoLeaveDateString(value) {
  return /^\d{1,2}-[A-Za-z]{3}-\d{4}$/i.test(String(value || '').trim());
}

function zohoLeaveDateKeyToTime(key) {
  const m = String(key || '')
    .trim()
    .match(/^(\d{1,2})-([A-Za-z]{3})-(\d{4})$/i);
  if (!m) return 0;
  const months = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];
  const mon = months.indexOf(m[2].toLowerCase());
  if (mon < 0) return 0;
  return new Date(parseInt(m[3], 10), mon, parseInt(m[1], 10)).getTime();
}

function zohoLeaveDateKeyIsWeekend(key) {
  const t = zohoLeaveDateKeyToTime(key);
  if (!t) return false;
  const day = new Date(t).getDay();
  return day === 0 || day === 6;
}

/**
 * Dates in Days JSON with LeaveCount > 0 (ignore LeaveCount 0.0 weekly offs).
 * Optional `excludeWeekends` drops Sat/Sun even when LeaveCount is 1
 * (Contingency booked often counts working days only).
 */
export function extractApprovedLeavePositiveDayKeys(daysField, { excludeWeekends = false } = {}) {
  const obj = parseJsonMaybe(daysField);
  if (!obj || typeof obj !== 'object' || Array.isArray(obj)) return [];
  return Object.keys(obj)
    .map((k) => String(k || '').trim())
    .filter((k) => {
      if (!isZohoLeaveDateString(k)) return false;
      const entry = obj[k];
      if (!entry || typeof entry !== 'object') return false;
      const n = Number(entry.LeaveCount ?? entry.leaveCount ?? entry.count);
      if (!Number.isFinite(n) || n <= 0) return false;
      if (excludeWeekends && zohoLeaveDateKeyIsWeekend(k)) return false;
      return true;
    })
    .sort((a, b) => zohoLeaveDateKeyToTime(a) - zohoLeaveDateKeyToTime(b));
}

/** Earliest / latest date keys from approved-leave `Days` JSON (e.g. "18-May-2026"). */
function extractLeaveDatesFromDaysField(daysField) {
  const positive = extractApprovedLeavePositiveDayKeys(daysField);
  if (positive.length > 0) {
    return { from: positive[0], to: positive[positive.length - 1] };
  }
  const obj = parseJsonMaybe(daysField);
  if (!obj || typeof obj !== 'object' || Array.isArray(obj)) {
    return { from: '', to: '' };
  }
  const dateKeys = Object.keys(obj)
    .map((k) => String(k || '').trim())
    .filter((k) => isZohoLeaveDateString(k));
  if (dateKeys.length === 0) return { from: '', to: '' };
  dateKeys.sort((a, b) => zohoLeaveDateKeyToTime(a) - zohoLeaveDateKeyToTime(b));
  return { from: dateKeys[0], to: dateKeys[dateKeys.length - 1] };
}

/**
 * Form N / Contingency: prefer From/To and day count from Days with LeaveCount > 0.
 * When that exceeds LeaveData booked, drop weekend days (Zoho often marks Sat as LeaveCount 1).
 */
export function resolveApprovedLeavePeriodForFormNGJ(record, bookedLimit = '') {
  const base = normalizeApprovedLeaveRecord(record || {});
  const daysField = record?.Days ?? record?.days;
  let positiveKeys = extractApprovedLeavePositiveDayKeys(daysField);
  let daysCount = parseApprovedLeaveDaysCount(daysField) || base.daysCount;

  const bookedNum = Number(String(bookedLimit || '').replace(/,/g, '').trim());
  const countNum = Number(String(daysCount || '').replace(/,/g, '').trim());
  if (
    Number.isFinite(bookedNum) &&
    bookedNum > 0 &&
    Number.isFinite(countNum) &&
    countNum > bookedNum
  ) {
    const weekdayKeys = extractApprovedLeavePositiveDayKeys(daysField, { excludeWeekends: true });
    if (weekdayKeys.length > 0) {
      positiveKeys = weekdayKeys;
      const weekdaySum = weekdayKeys.reduce((sum, key) => {
        const obj = parseJsonMaybe(daysField);
        const entry = obj?.[key];
        const n = Number(entry?.LeaveCount ?? entry?.leaveCount ?? entry?.count);
        return sum + (Number.isFinite(n) && n > 0 ? n : 0);
      }, 0);
      if (weekdaySum > 0) daysCount = String(weekdaySum);
    }
  }

  const from = positiveKeys[0] || base.from;
  const to = positiveKeys[positiveKeys.length - 1] || base.to || from;
  return {
    from,
    to,
    daysCount: daysCount || base.daysCount,
    wagesPaidDate: from && to ? (from === to ? from : `${from} to ${to}`) : base.wagesPaidDate,
  };
}

export function normalizeApprovedLeaveRecord(record) {
  if (!record || typeof record !== 'object') {
    return { from: '', to: '', daysCount: '', wagesPaidDate: '' };
  }
  const daysField = record.Days ?? record.days;
  let from = String(record.From ?? record.from ?? '').trim();
  let to = String(record.To ?? record.to ?? '').trim();

  if (!isZohoLeaveDateString(from) || !isZohoLeaveDateString(to)) {
    const fromDays = extractLeaveDatesFromDaysField(daysField);
    if (!isZohoLeaveDateString(from) && fromDays.from) from = fromDays.from;
    if (!isZohoLeaveDateString(to) && fromDays.to) to = fromDays.to;
  }

  const daysCount = parseApprovedLeaveDaysCount(daysField);
  let wagesPaidDate = from;
  if (from && to) {
    wagesPaidDate = from === to ? from : `${from} to ${to}`;
  } else if (!from && to) {
    wagesPaidDate = to;
  }
  return { from, to, daysCount, wagesPaidDate };
}

/** Canonical storage key for Form F grid column index (avoids duplicate "From"/"To" display labels). */
export function getFormFKarnatakaStorageHeader(colIndex, fallbackHeader = '') {
  const key = FORM_F_KARNATAKA_PART_I_HEADERS[colIndex];
  if (key) return key;
  return String(fallbackHeader || '').trim();
}

/** Read cell value using canonical key first, then legacy duplicate short headers. */
export function readFormFKarnatakaRowCell(row, colIndex, displayHeader) {
  if (!row || typeof row !== 'object') return '';
  const storageKey = getFormFKarnatakaStorageHeader(colIndex, displayHeader);
  const primary = row[storageKey];
  if (primary != null && String(primary).trim() !== '') return primary;
  const legacy = row[displayHeader];
  if (legacy != null && String(legacy).trim() !== '') return legacy;
  return '';
}

export function isFormFKarnatakaLeaveAvailedColumnIndex(colIndex) {
  return colIndex === 5 || colIndex === 6 || colIndex === 7;
}

/** StatutoryData overlay must not overwrite API-filled leave availed cells (row-index keys are unreliable). */
export function isFormFKarnatakaLeaveAutofillOverlaySkipHeader(header) {
  const n = formFKarnatakaHeaderNorm(header);
  if (!n) return false;
  if (isFormFKarnatakaLeaveEarnedHeader(header)) return true;
  if (isFormFKarnatakaLeaveAtCreditHeader(header)) return true;
  if (isFormFKarnatakaLeaveAvailedHeader(header)) return true;
  if (isFormFKarnatakaLeaveBalanceHeader(header)) return true;
  if (n.includes('wages') && n.includes('leave') && n.includes('paid')) return true;
  if (n === 'from' || n === 'to') return true;
  if (/no\.?\s*of\s+days/.test(n) && !n.includes('days worked')) return true;
  return false;
}

function zohoLeaveDateToTime(dateStr) {
  const m = String(dateStr || '').match(/^(\d{1,2})-([A-Za-z]{3})-(\d{4})$/i);
  if (!m) return 0;
  const months = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];
  const mon = months.indexOf(m[2].toLowerCase());
  if (mon < 0) return 0;
  return new Date(parseInt(m[3], 10), mon, parseInt(m[1], 10)).getTime();
}

/** True only when approved leave has availed days overlapping the statutory month. */
export function approvedLeaveRecordIsAvailedInMonth(record, monthFrom, monthTo) {
  if (!record || typeof record !== 'object') return false;
  const status = String(
    record.ApprovalStatus ?? record.approvalStatus ?? record.Status ?? record.status ?? 'APPROVED'
  )
    .trim()
    .toUpperCase();
  if (status && status !== 'APPROVED') return false;

  const metrics = normalizeApprovedLeaveRecord(record);
  const days = Number(String(metrics.daysCount || '').replace(/,/g, ''));
  if (!Number.isFinite(days) || days <= 0) return false;

  const rangeStart = zohoLeaveDateToTime(monthFrom);
  const rangeEnd = zohoLeaveDateToTime(monthTo);
  if (!rangeStart || !rangeEnd) return true;

  const leaveStart = zohoLeaveDateToTime(metrics.from);
  const leaveEnd = zohoLeaveDateToTime(metrics.to || metrics.from);
  if (!leaveStart) return false;
  const end = leaveEnd || leaveStart;
  return leaveStart <= rangeEnd && end >= rangeStart;
}

export function filterApprovedLeaveRecordsForMonth(records, monthFrom, monthTo) {
  return (Array.isArray(records) ? records : []).filter((record) =>
    approvedLeaveRecordIsAvailedInMonth(record, monthFrom, monthTo)
  );
}

/** Strict full-name match — never first-name-only (avoids wrong Leave availed dates). */
export function formFKarnatakaPersonNamesMatchStrict(candidate, recordName) {
  const a = normalizePersonNameKey(candidate);
  const b = normalizePersonNameKey(recordName);
  if (!a || !b) return false;
  if (a === b) return true;

  const aParts = a.split(' ').filter(Boolean);
  const bParts = b.split(' ').filter(Boolean);
  // Require at least two name parts on both sides — never match "Sakthivel" alone.
  if (aParts.length < 2 || bParts.length < 2) return false;
  if (aParts[0] !== bParts[0]) return false;

  const aLast = aParts[aParts.length - 1];
  const bLast = bParts[bParts.length - 1];
  if (aLast === bLast) return true;
  if (aLast.length === 1 && bLast.startsWith(aLast)) return true;
  if (bLast.length === 1 && aLast.startsWith(bLast)) return true;
  return false;
}

export function readFormFKarnatakaApprovedLeaveType(record) {
  if (!record || typeof record !== 'object') return '';
  const raw =
    record['Leave Type'] ??
    record['Leave type'] ??
    record.LeaveType ??
    record.leaveType ??
    record.Leavetype ??
    record['Leavetype'] ??
    record['Type of Leave'] ??
    record.TypeOfLeave ??
    record.leave_type ??
    record.LeaveName ??
    record['Leave Name'] ??
    '';
  if (raw != null && typeof raw === 'object') {
    return String(
      raw.name ?? raw.Name ?? raw.value ?? raw.label ?? raw.displayValue ?? raw.display_value ?? ''
    ).trim();
  }
  return String(raw || '').trim();
}

/** PART I Earned Leave — ignore Contingency / Casual / Sick approved leave. */
export function isFormFKarnatakaEarnedApprovedLeave(record) {
  const leaveType = readFormFKarnatakaApprovedLeaveType(record);
  if (!leaveType) {
    // Zoho sometimes omits type; keep only when Days look like a real earned leave booking.
    return true;
  }
  const n = String(leaveType)
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  if (/contingency|casual|sick|accident|maternity|paternity|unpaid|lop|loss of pay|comp\s*off|compensatory/.test(n)) {
    return false;
  }
  return /earned/.test(n) || /^el\b/.test(n) || n === 'el' || /privilege|annual/.test(n);
}

/** PART II Sick/Accident Leave — only Sick or Accident approved leave. */
export function isFormFKarnatakaSickAccidentApprovedLeave(record) {
  const leaveType = readFormFKarnatakaApprovedLeaveType(record);
  if (!leaveType) return false;
  const n = String(leaveType)
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  if (/contingency|casual|earned|privilege|annual|maternity|paternity|unpaid|lop|loss of pay/.test(n)) {
    return false;
  }
  return /sick|accident/.test(n);
}

export function filterApprovedLeaveRecordsForFormFKarnataka(records, monthFrom, monthTo) {
  return filterApprovedLeaveRecordsForMonth(records, monthFrom, monthTo).filter(
    isFormFKarnatakaEarnedApprovedLeave
  );
}

export function filterApprovedSickLeaveRecordsForFormFKarnataka(records, monthFrom, monthTo) {
  return filterApprovedLeaveRecordsForMonth(records, monthFrom, monthTo).filter(
    isFormFKarnatakaSickAccidentApprovedLeave
  );
}

function findFormFKarnatakaSickLeaveObject(leaveRecord, leaveTypeLabels = {}) {
  if (!leaveRecord || typeof leaveRecord !== 'object') return null;
  const labels = leaveTypeLabels && typeof leaveTypeLabels === 'object' ? leaveTypeLabels : {};
  const keys = Object.keys(leaveRecord);
  for (let i = 0; i < keys.length; i += 1) {
    const key = keys[i];
    if (/^employee$/i.test(key) || /totals|month/i.test(key)) continue;
    const label = labels[key] != null ? String(labels[key]) : key;
    const blob = `${key} ${label}`.toLowerCase();
    if (/contingency|casual|earned|privilege|maternity|paternity/.test(blob)) continue;
    if (!/sick|accident/.test(blob)) continue;
    const obj = leaveRecord[key];
    if (obj && typeof obj === 'object') return obj;
    if (obj != null && String(obj).trim() !== '') {
      return { paidBalance: String(obj).trim(), paidBooked: '' };
    }
  }
  // LeaveData / flattened keys
  const flatCredit =
    leaveRecord.SickLeave ??
    leaveRecord['Sick Leave'] ??
    leaveRecord.AccidentLeave ??
    leaveRecord['Accident Leave'] ??
    '';
  if (flatCredit != null && String(flatCredit).trim() !== '') {
    return { paidBalance: String(flatCredit).trim(), paidBooked: '' };
  }
  return null;
}

/** Resolve PART II Year / of credit / Availed / Balance from LeaveData + approved sick leave. */
export function resolveFormFKarnatakaPartIiValues({
  leaveRecord = null,
  leaveTypeLabels = {},
  approvedSickRecord = null,
  monthFrom = '',
  yearOverride = '',
} = {}) {
  const yearFromMonth = String(monthFrom || '')
    .trim()
    .match(/(20\d{2}|19\d{2})/);
  const year =
    String(yearOverride || '').trim() ||
    (yearFromMonth ? yearFromMonth[1] : '') ||
    String(new Date().getFullYear());

  const sickObj = findFormFKarnatakaSickLeaveObject(leaveRecord, leaveTypeLabels);
  let credit = '';
  let availedFromBalance = '';
  if (sickObj) {
    const bal = sickObj.paidBalance ?? sickObj.balance ?? sickObj.Balance;
    const booked = sickObj.paidBooked ?? sickObj.booked ?? sickObj.Booked;
    if (bal != null && String(bal).trim() !== '') credit = String(bal).trim();
    if (booked != null && String(booked).trim() !== '') availedFromBalance = String(booked).trim();
  }

  let availed = availedFromBalance;
  if (approvedSickRecord) {
    const count = parseApprovedLeaveDaysCount(
      approvedSickRecord.Days ?? approvedSickRecord.days
    );
    if (count !== '') availed = String(count);
  }

  let balance = '';
  const creditNum = Number(String(credit).replace(/,/g, '').trim());
  const availedNum = Number(String(availed || '0').replace(/,/g, '').trim());
  if (Number.isFinite(creditNum)) {
    const a = Number.isFinite(availedNum) ? availedNum : 0;
    balance = String(Math.round((creditNum - a) * 100) / 100);
  } else if (credit === '' && availed === '') {
    balance = '-';
  }

  return {
    [FORM_F_KARNATAKA_PART_II_HEADERS[0]]: year,
    [FORM_F_KARNATAKA_PART_II_HEADERS[1]]: credit,
    [FORM_F_KARNATAKA_PART_II_HEADERS[2]]: availed,
    [FORM_F_KARNATAKA_PART_II_HEADERS[3]]: balance,
  };
}

export function applyFormFKarnatakaPartIiToRow(row, partIiValues, { overwrite = true } = {}) {
  if (!row || !partIiValues || typeof partIiValues !== 'object') return 0;
  let applied = 0;
  FORM_F_KARNATAKA_PART_II_HEADERS.forEach((key) => {
    const value = partIiValues[key];
    if (value == null || String(value).trim() === '') return;
    if (!overwrite && String(row[key] ?? '').trim() !== '') return;
    row[key] = String(value).trim();
    applied += 1;
  });
  return applied;
}

export function applyFormFKarnatakaPartIiAutofill(
  mappedData,
  employeesForMapping,
  leaveRecords,
  leaveTypeLabels,
  approvedSickRecords,
  options = {}
) {
  if (!Array.isArray(mappedData) || mappedData.length === 0) return 0;
  const employees = Array.isArray(employeesForMapping) ? employeesForMapping : [];
  const leaveLookup = buildLeaveRecordLookupMap(leaveRecords || []);
  const approved = Array.isArray(approvedSickRecords) ? approvedSickRecords : [];
  const monthFrom = options.monthFrom || options.fromDate || '';
  const yearOverride = options.yearOverride || '';
  let hits = 0;

  mappedData.forEach((row, index) => {
    if (!row || typeof row !== 'object') return;
    const empItem = employees[index] || null;
    const emp = empItem?.Employee || empItem?.employee || empItem;
    const leaveRecord =
      findLeaveRecordForFormRow(leaveLookup, row, '', '') ||
      (emp
        ? findLeaveRecordForFormRow(
            leaveLookup,
            {
              ...row,
              __employeeLookupName:
                row.__employeeLookupName || formatFormFKarnatakaEmployeeName(emp),
              __employeeLookupId:
                row.__employeeLookupId ||
                String(emp.Employee_ID || emp['Employee ID'] || emp.EmployeeID || '').trim(),
            },
            '',
            ''
          )
        : null);

    const approvedHit = findApprovedLeaveForEmployee(
      null,
      emp,
      row,
      '',
      'Name of the person',
      {
        approvedRecords: approved,
        monthFrom: options.monthFrom,
        monthTo: options.monthTo,
        strictIdentity: true,
        leaveRecords: leaveRecords || [],
        leaveTypeLabels: leaveTypeLabels || {},
      }
    );

    const values = resolveFormFKarnatakaPartIiValues({
      leaveRecord,
      leaveTypeLabels,
      approvedSickRecord: approvedHit,
      monthFrom,
      yearOverride,
    });
    hits += applyFormFKarnatakaPartIiToRow(row, values, { overwrite: options.overwrite !== false });
  });
  return hits;
}

export function normalizePersonNameKey(name) {
  return String(name || '')
    .replace(/\([^)]*\)/g, ' ')
    .replace(/[^a-z0-9\s]/gi, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

function tokensAppearInOrder(needles, haystack) {
  if (!Array.isArray(needles) || needles.length === 0) return false;
  if (!Array.isArray(haystack) || haystack.length === 0) return false;
  let i = 0;
  for (let h = 0; h < haystack.length; h += 1) {
    if (haystack[h] === needles[i]) i += 1;
    if (i === needles.length) return true;
  }
  return false;
}

/**
 * Match worker vs approved-leave employee name.
 * Handles 3-part Gujarati names ("Patel Chandrakant Virabhai") when Zoho FirstName
 * includes a middle name or omits the community surname.
 * Still rejects last-initial shortcuts ("Harshad Dk" ↛ "Harshad Dineshkumar").
 */
export function personNamesMatchFirstLastStrict(workerName, recordName, firstName = '', lastName = '') {
  const recordNorm = normalizePersonNameKey(recordName);
  if (!recordNorm) return false;

  const workerNorm = normalizePersonNameKey(workerName);
  if (workerNorm && workerNorm === recordNorm) return true;

  const fn = normalizePersonNameKey(firstName);
  const ln = normalizePersonNameKey(lastName);
  if (fn && ln) {
    const combined = `${fn} ${ln}`.replace(/\s+/g, ' ').trim();
    if (combined === recordNorm) return true;

    const fnParts = fn.split(' ').filter(Boolean);
    const lnParts = ln.split(' ').filter(Boolean);
    const recordParts = recordNorm.split(' ').filter(Boolean);
    const combinedParts = combined.split(' ').filter(Boolean);
    if (recordParts.length < 2) return false;

    const lastToken = lnParts[lnParts.length - 1];
    const recordLast = recordParts[recordParts.length - 1];
    // Truncated last name / initial must not match a longer surname.
    if (lastToken !== recordLast) return false;

    if (
      recordParts.length >= fnParts.length + lnParts.length &&
      fnParts.every((p, i) => recordParts[i] === p) &&
      lnParts.every((p, i) => recordParts[recordParts.length - lnParts.length + i] === p)
    ) {
      return true;
    }

    if (
      combinedParts.length >= 2 &&
      (tokensAppearInOrder(combinedParts, recordParts) ||
        tokensAppearInOrder(recordParts, combinedParts))
    ) {
      return true;
    }

    return false;
  }

  if (!workerNorm) return false;
  const aParts = workerNorm.split(' ').filter(Boolean);
  const bParts = recordNorm.split(' ').filter(Boolean);
  if (aParts.length < 2 || bParts.length < 2) return false;
  return aParts[0] === bParts[0] && aParts[aParts.length - 1] === bParts[bParts.length - 1];
}

/** Strict person-name match — avoids Satheesh ↔ Sathishkumar substring false positives. */
export function personNamesMatch(candidate, recordName) {
  const a = normalizePersonNameKey(candidate);
  const b = normalizePersonNameKey(recordName);
  if (!a || !b) return false;
  if (a === b) return true;
  const aParts = a.split(' ');
  const bParts = b.split(' ');
  if (aParts.length === 1 && bParts.length >= 2 && aParts[0].length >= 4 && aParts[0] === bParts[0]) {
    return true;
  }
  if (bParts.length === 1 && aParts.length >= 2 && bParts[0].length >= 4 && bParts[0] === aParts[0]) {
    return true;
  }
  return false;
}

export function getApprovedLeaveRecordUniqueKey(record) {
  if (!record || typeof record !== 'object') return '';
  const id = String(
    record.recordId || record['Zoho.ID'] || record.ZohoID || record.id || ''
  )
    .trim()
    .toLowerCase();
  if (id) return id;
  const { names } = collectApprovedLeaveIdentityKeys(record);
  const from = String(record.From ?? record.from ?? '').trim();
  const to = String(record.To ?? record.to ?? '').trim();
  return `${names[0] || 'unknown'}|${from}|${to}`;
}

export function collectApprovedLeaveIdentityKeys(record) {
  const names = new Set();
  const ids = new Set();
  if (!record || typeof record !== 'object') {
    return { names: [], ids: [] };
  }
  const employee = record.Employee ?? record.employee;
  if (typeof employee === 'string' && employee.trim()) {
    names.add(normalizePersonNameKey(employee));
  } else if (employee && typeof employee === 'object') {
    const n = employee.name ?? employee.Name;
    if (n) names.add(normalizePersonNameKey(n));
    const id = employee.id ?? employee.ID ?? employee.erecno;
    if (id) ids.add(String(id).trim().toLowerCase());
  }
  ['Employee Name', 'EmployeeName', 'employeeName'].forEach((k) => {
    const v = record[k];
    if (v) names.add(normalizePersonNameKey(v));
  });
  // Also compose FirstName + LastName when Zoho nests them on Employee / record.
  const first =
    (employee && typeof employee === 'object'
      ? employee.FirstName || employee.firstName || employee['First Name']
      : null) ||
    record.FirstName ||
    record.firstName ||
    record['First Name'];
  const last =
    (employee && typeof employee === 'object'
      ? employee.LastName || employee.lastName || employee['Last Name']
      : null) ||
    record.LastName ||
    record.lastName ||
    record['Last Name'];
  const fn = String(first || '').trim();
  const ln = String(last || '').trim();
  if (fn && ln) names.add(normalizePersonNameKey(`${fn} ${ln}`));
  [
    'Employee.ID',
    'EmployeeID',
    'employeeId',
    'Employee Id',
    'erecno',
    'Erecno',
    'ZohoID',
    'Zoho_ID',
    'zoho_id',
    'Zoho.ID',
  ].forEach((k) => {
    const v = record[k];
    if (v != null && String(v).trim()) ids.add(String(v).trim().toLowerCase());
  });
  return { names: [...names].filter(Boolean), ids: [...ids].filter(Boolean) };
}

export function buildApprovedLeaveLookupMap(records) {
  const map = new Map();
  const add = (key, record) => {
    const k = String(key || '').trim().toLowerCase();
    if (!k || map.has(k)) return;
    map.set(k, record);
  };
  (Array.isArray(records) ? records : []).forEach((record) => {
    const { names, ids } = collectApprovedLeaveIdentityKeys(record);
    names.forEach((n) => add(n, record));
    ids.forEach((id) => add(id, record));
    const uniqueKey = getApprovedLeaveRecordUniqueKey(record);
    if (uniqueKey) add(uniqueKey, record);
  });
  return map;
}

export function resolveFormFKarnatakaApprovedLeaveHeaders(tableHeaders) {
  const headers = Array.isArray(tableHeaders) ? tableHeaders : [];
  const canonical = FORM_F_KARNATAKA_PART_I_HEADERS;

  const byCanonicalIndex = (canonicalIdx) => {
    const canonicalName = canonical[canonicalIdx];
    const exact = headers.find(
      (h) => formFKarnatakaHeaderNorm(h) === formFKarnatakaHeaderNorm(canonicalName)
    );
    if (exact) return exact;
    if (headers[canonicalIdx] != null && String(headers[canonicalIdx]).trim() !== '') {
      return headers[canonicalIdx];
    }
    return canonicalName;
  };

  if (headers.length >= 8) {
    return {
      leaveAvailedFrom: byCanonicalIndex(5),
      leaveAvailedTo: byCanonicalIndex(6),
      leaveAvailedDays: byCanonicalIndex(7),
      wagesPaidDate: byCanonicalIndex(9),
      columnIndices: { from: 5, to: 6, days: 7, wages: 9 },
    };
  }

  const pick = (...candidates) => {
    for (const cand of candidates) {
      const want = formFKarnatakaHeaderNorm(cand);
      const exact = headers.find((h) => formFKarnatakaHeaderNorm(h) === want);
      if (exact) return exact;
    }
    for (const cand of candidates) {
      const want = formFKarnatakaHeaderNorm(cand);
      const partial = headers.find((h) => {
        const n = formFKarnatakaHeaderNorm(h);
        return n.includes(want) || want.includes(n);
      });
      if (partial) return partial;
    }
    return '';
  };

  return {
    leaveAvailedFrom: pick('Leave availed From'),
    leaveAvailedTo: pick('Leave availed To'),
    leaveAvailedDays: pick('Leave availed No. of days', 'No. of days'),
    wagesPaidDate: pick('Date on which wages for leave paid'),
    columnIndices: null,
  };
}

export function findApprovedLeaveForFormRow(lookup, row, employeeIdHeader, employeeNameHeader) {
  if (!lookup || typeof lookup.get !== 'function') return null;
  const tryKeys = [];
  if (employeeNameHeader && row?.[employeeNameHeader]) {
    tryKeys.push(normalizePersonNameKey(row[employeeNameHeader]));
  }
  if (employeeIdHeader && row?.[employeeIdHeader]) {
    tryKeys.push(String(row[employeeIdHeader]).trim().toLowerCase());
  }
  for (const key of tryKeys) {
    if (key && lookup.has(key)) return lookup.get(key);
  }
  return null;
}

function collectEmployeeIdentityCandidates(emp, row, employeeIdHeader, employeeNameHeader, options = {}) {
  const nameCandidates = new Set();
  const idCandidates = new Set();
  const strictIdentity = options.strictIdentity === true;

  const addName = (value) => {
    const n = normalizePersonNameKey(value);
    if (!n) return;
    // Strict mode: skip single-token names so first-name-only cannot match another employee.
    if (strictIdentity && !n.includes(' ')) return;
    nameCandidates.add(n);
  };
  const addId = (value) => {
    const id = String(value || '').trim().toLowerCase();
    if (!id || id.length < 4) return;
    if (/^\d{1,3}$/.test(id)) return;
    idCandidates.add(id);
  };

  if (typeof options.collectEmployeeNameCandidates === 'function') {
    options.collectEmployeeNameCandidates(emp).forEach(addName);
  }
  if (typeof options.collectEmployeeIdCandidates === 'function') {
    options.collectEmployeeIdCandidates(emp, row).forEach(addId);
  }

  if (employeeNameHeader && row?.[employeeNameHeader]) addName(row[employeeNameHeader]);
  if (employeeIdHeader && row?.[employeeIdHeader]) addId(row[employeeIdHeader]);
  if (row?.__employeeLookupName) addName(row.__employeeLookupName);
  if (row?.__employeeLookupId) addId(row.__employeeLookupId);

  if (emp && typeof emp === 'object') {
    addName(emp.Employee_Name || emp['Employee Name'] || emp.employeeName || emp.name || emp.Name);
    addId(
      emp.Zoho_ID ||
        emp['Zoho_ID'] ||
        emp.ZohoID ||
        emp['ZohoID'] ||
        emp.erecno ||
        emp.Erecno ||
        emp['Erecno'] ||
        emp.Employee_ID ||
        emp['Employee ID'] ||
        emp.employeeId ||
        emp.EmployeeID ||
        emp.Emp_ID ||
        emp.empId
    );
    const first = String(emp.FirstName || emp['FirstName'] || emp.firstName || '').trim();
    const last = String(emp.LastName || emp['LastName'] || emp.lastName || '').trim();
    if (first && last) addName(`${first} ${last}`);
    // Never index bare first name for Form F — that caused wrong Leave availed dates.
    if (!strictIdentity && first) addName(first);
  }

  return { nameCandidates, idCandidates };
}

function approvedLeaveRecordMatchesEmployee(record, nameCandidates, idCandidates, options = {}) {
  const { names, ids } = collectApprovedLeaveIdentityKeys(record);
  if (ids.some((id) => idCandidates.has(id))) return true;
  const nameMatch =
    options.strictIdentity === true ? formFKarnatakaPersonNamesMatchStrict : personNamesMatch;
  for (const recordName of names) {
    for (const cand of nameCandidates) {
      if (nameMatch(cand, recordName)) return true;
    }
  }
  return false;
}

export function findApprovedLeaveForEmployee(
  lookup,
  emp,
  row,
  employeeIdHeader,
  employeeNameHeader,
  options = {}
) {
  const approvedRecords = Array.isArray(options.approvedRecords) ? options.approvedRecords : [];
  const usedRecordKeys =
    options.usedRecordKeys instanceof Set ? options.usedRecordKeys : new Set();
  const { nameCandidates, idCandidates } = collectEmployeeIdentityCandidates(
    emp,
    row,
    employeeIdHeader,
    employeeNameHeader,
    options
  );

  if (nameCandidates.size === 0 && idCandidates.size === 0) return null;

  const isAvailable = (record) => {
    const key = getApprovedLeaveRecordUniqueKey(record);
    return key && !usedRecordKeys.has(key);
  };

  const matches = (record) =>
    approvedLeaveRecordMatchesEmployee(record, nameCandidates, idCandidates, options);

  const pickFrom = (records) => {
    const monthFrom = options.monthFrom || '';
    const monthTo = options.monthTo || '';
    for (const record of records) {
      if (!record || !isAvailable(record)) continue;
      if (monthFrom && monthTo && !approvedLeaveRecordIsAvailedInMonth(record, monthFrom, monthTo)) {
        continue;
      }
      if (matches(record)) {
        return record;
      }
    }
    return null;
  };

  // Prefer employee id exact map hits before any name matching.
  if (lookup && typeof lookup.get === 'function') {
    const monthFrom = options.monthFrom || '';
    const monthTo = options.monthTo || '';
    const lookupHitValid = (hit) =>
      hit &&
      isAvailable(hit) &&
      (!monthFrom || !monthTo || approvedLeaveRecordIsAvailedInMonth(hit, monthFrom, monthTo)) &&
      matches(hit);

    for (const id of idCandidates) {
      const hit = lookup.get(id);
      if (lookupHitValid(hit)) return hit;
    }
    for (const name of nameCandidates) {
      const hit = lookup.get(name);
      if (lookupHitValid(hit)) return hit;
    }
  }

  return pickFrom(approvedRecords);
}

export function clearFormFKarnatakaApprovedLeaveOnRow(row) {
  if (!row) return;
  const canon = FORM_F_KARNATAKA_PART_I_HEADERS;
  [5, 6, 7, 9].forEach((idx) => {
    const key = canon[idx];
    if (key) row[key] = '';
  });
}

export function applyFormFKarnatakaApprovedLeaveToRow(
  row,
  approvedRecord,
  headers,
  tableHeaders,
  { overwrite = true, monthFrom = '', monthTo = '' } = {}
) {
  if (!row || !approvedRecord) return 0;
  if (monthFrom && monthTo && !approvedLeaveRecordIsAvailedInMonth(approvedRecord, monthFrom, monthTo)) {
    return 0;
  }
  const metrics = normalizeApprovedLeaveRecord(approvedRecord);
  if (!metrics.daysCount || Number(metrics.daysCount) <= 0) return 0;
  const canon = FORM_F_KARNATAKA_PART_I_HEADERS;
  let applied = 0;

  const setByIndex = (idx, value) => {
    if (value == null || value === '') return;
    const key = canon[idx];
    if (!key) return;
    if (!overwrite && String(row[key] ?? '').trim() !== '') return;
    row[key] = String(value);
    applied += 1;
  };

  setByIndex(5, metrics.from);
  setByIndex(6, metrics.to);
  setByIndex(7, metrics.daysCount);
  setByIndex(9, metrics.wagesPaidDate);

  const earnedDuring = String(row[canon[3]] ?? '').trim();
  if (earnedDuring) {
    // Leave at credit stays as Leave earned during the Period.
    if (overwrite || String(row[canon[4]] ?? '').trim() === '') {
      setByIndex(4, earnedDuring);
    }
    // Balance on return ← Balance count (earned − LeaveCount).
    setByIndex(8, computeFormFKarnatakaBalanceOnReturn(earnedDuring, metrics.daysCount || '0'));
  }
  return applied;
}

export function isFormFKarnatakaTotalDaysWorkedHeader(header) {
  const s = formFKarnatakaHeaderNorm(header);
  return s === 'total days worked' || (s.includes('total') && s.includes('days') && s.includes('worked'));
}

/** Total days worked (PART I col 3) ← Payroll `paid_days`. */
export function resolveFormFKarnatakaPayrollPaidDays(payrollRow) {
  if (!payrollRow || payrollRow.fetch_error) return '';
  const flat = flattenPayrollEarningColumns(payrollRow);
  const keys = ['paid_days', 'Paid Days', 'days_worked', 'Days Worked', 'paidDays', 'no_of_days_worked'];
  const patterns = [/^paid_days$/, /paiddays/, /daysworked/, /days_present/, /noofdayspresent/, /no_of_days_present/];
  const fromFlat = readPayrollScalar(flat, keys, patterns);
  if (fromFlat !== '') return String(fromFlat);
  if (payrollRow !== flat) {
    const fromRow = readPayrollScalar(payrollRow, keys, patterns);
    if (fromRow !== '') return String(fromRow);
  }
  return '';
}

export function applyFormFKarnatakaPayrollToRow(row, payrollRow, { overwrite = true } = {}) {
  if (!row || !payrollRow || payrollRow.fetch_error) return 0;
  const paidDays = resolveFormFKarnatakaPayrollPaidDays(payrollRow);
  if (paidDays === '') return 0;
  const key = FORM_F_KARNATAKA_PART_I_HEADERS[2];
  if (!key) return 0;
  if (!overwrite && String(row[key] ?? '').trim() !== '') return 0;
  row[key] = String(paidDays);
  return 1;
}

export function applyFormFKarnatakaPayrollAutofill(
  mappedData,
  employeesForMapping,
  payrollRows,
  tableHeaders,
  options = {}
) {
  if (!Array.isArray(mappedData) || mappedData.length === 0) return 0;
  if (!Array.isArray(payrollRows) || payrollRows.length === 0) return 0;

  const priorHeaders = Array.isArray(tableHeaders) ? tableHeaders : [];
  const resolvePayrollRow =
    typeof options.resolvePayrollRow === 'function'
      ? options.resolvePayrollRow
      : (emp) => resolveFormXIXMPPayrollRowForEmployee(emp, payrollRows);

  let hits = 0;
  mappedData.forEach((row, index) => {
    const normalizedRow = remapFormFKarnatakaRowToCanonicalHeaders(row, priorHeaders);
    Object.assign(row, normalizedRow);

    const empItem = employeesForMapping[index];
    const emp = empItem && (empItem.Employee || empItem.employee || empItem);
    const payrollRow = resolvePayrollRow(emp, row);
    if (!payrollRow || payrollRow.fetch_error) return;
    const applied = applyFormFKarnatakaPayrollToRow(row, payrollRow, {
      overwrite: options.overwrite !== false,
    });
    if (options.monthCandidates) {
      applyFormFKarnatakaAprilPartIDefaultsToRow(row, options.monthCandidates, {
        overwrite: options.overwrite !== false,
      });
    }
    if (applied > 0) hits += 1;
  });
  return hits;
}

export function applyFormFKarnatakaApprovedLeaveAutofill(
  mappedData,
  employeesForMapping,
  tableHeaders,
  approvedLeaveRecords,
  options = {}
) {
  if (!Array.isArray(mappedData) || mappedData.length === 0) return 0;
  if (!Array.isArray(approvedLeaveRecords) || approvedLeaveRecords.length === 0) return 0;

  const canonicalHeaders = resolveFormFKarnatakaTableHeaders(tableHeaders);
  const priorHeaders = Array.isArray(tableHeaders) ? tableHeaders : [];
  const monthFrom = options.monthFrom || '';
  const monthTo = options.monthTo || '';
  const earnedApprovedRecords = filterApprovedLeaveRecordsForFormFKarnataka(
    approvedLeaveRecords,
    monthFrom,
    monthTo
  );
  if (earnedApprovedRecords.length === 0) {
    mappedData.forEach((row) => {
      if (options.overwrite !== false) {
        const normalizedRow = remapFormFKarnatakaRowToCanonicalHeaders(row, priorHeaders);
        Object.assign(row, normalizedRow);
        clearFormFKarnatakaApprovedLeaveOnRow(row);
      }
    });
    return 0;
  }

  const lookup = buildApprovedLeaveLookupMap(earnedApprovedRecords);
  const leaveLookup =
    Array.isArray(options.leaveRecords) && options.leaveRecords.length > 0
      ? buildLeaveRecordLookupMap(options.leaveRecords)
      : null;
  const leaveTypeLabels = options.leaveTypeLabels || {};
  const { employeeNameHeader, employeeIdHeader } = options.resolveEmployeeHeaders
    ? options.resolveEmployeeHeaders(canonicalHeaders)
    : { employeeNameHeader: '', employeeIdHeader: '' };

  const usedRecordKeys = new Set();
  let hits = 0;
  mappedData.forEach((row, index) => {
    const normalizedRow = remapFormFKarnatakaRowToCanonicalHeaders(row, priorHeaders);
    Object.assign(row, normalizedRow);

    const empItem = employeesForMapping[index];
    const emp = empItem && (empItem.Employee || empItem.employee || empItem);

    // If LeaveData says Booked: 0 for Earned Leave, do not invent Leave availed dates.
    if (leaveLookup) {
      const leaveRecord = findLeaveRecordForFormRow(
        leaveLookup,
        row,
        employeeIdHeader,
        employeeNameHeader
      );
      if (leaveRecord) {
        const { periodBooked } = getForm15EarnedLeavePeriodMetrics(leaveRecord, leaveTypeLabels);
        const bookedNum = Number(String(periodBooked || '').replace(/,/g, ''));
        if (!Number.isFinite(bookedNum) || bookedNum <= 0) {
          if (options.overwrite !== false) clearFormFKarnatakaApprovedLeaveOnRow(row);
          return;
        }
      }
    }

    const approvedRecord = findApprovedLeaveForEmployee(
      lookup,
      emp,
      row,
      employeeIdHeader,
      employeeNameHeader,
      {
        approvedRecords: earnedApprovedRecords,
        usedRecordKeys,
        monthFrom,
        monthTo,
        strictIdentity: true,
        collectEmployeeNameCandidates: options.collectEmployeeNameCandidates,
        collectEmployeeIdCandidates: options.collectEmployeeIdCandidates,
      }
    );
    if (!approvedRecord) {
      if (options.overwrite !== false) clearFormFKarnatakaApprovedLeaveOnRow(row);
      return;
    }
    const recordKey = getApprovedLeaveRecordUniqueKey(approvedRecord);
    if (recordKey) usedRecordKeys.add(recordKey);
    const applied = applyFormFKarnatakaApprovedLeaveToRow(
      row,
      approvedRecord,
      null,
      canonicalHeaders,
      { overwrite: options.overwrite !== false, monthFrom, monthTo }
    );
    if (applied > 0) hits += 1;
    else if (options.overwrite !== false) clearFormFKarnatakaApprovedLeaveOnRow(row);
  });
  return hits;
}

export function extractApprovedLeaveRecordsFromApiResult(result) {
  if (!result || typeof result !== 'object') {
    return { records: [], meta: null };
  }
  if (Array.isArray(result.leaveRecords) && result.leaveRecords.length > 0) {
    return { records: result.leaveRecords, meta: result.meta || null };
  }
  if (result.records && typeof result.records === 'object' && !Array.isArray(result.records)) {
    const records = Object.entries(result.records).map(([id, row]) => ({
      recordId: String(id),
      ...(row && typeof row === 'object' ? row : {}),
    }));
    return { records, meta: result.meta || null };
  }
  return { records: [], meta: result.meta || null };
}

/** Form F PART I table spans columns A–K (11 cols). Titles sit centered across that band. */
export const FORM_F_KARNATAKA_TITLE_COL_FROM = 1;
export const FORM_F_KARNATAKA_TITLE_COL_TO = 11;
export const FORM_F_KARNATAKA_DEFAULT_TITLE = 'FORM F';
export const FORM_F_KARNATAKA_DEFAULT_REFERENCE = '(SEE RULE 8)';
export const FORM_F_KARNATAKA_DEFAULT_SUBTITLE = 'REGISTER OF LEAVE WITH WAGES';

function formFExcelJsCellText(value) {
  if (value == null) return '';
  if (typeof value === 'string' || typeof value === 'number') return String(value).trim();
  if (typeof value === 'object') {
    if (value.text != null) return String(value.text).trim();
    if (value.richText && Array.isArray(value.richText)) {
      return value.richText.map((p) => String(p?.text || '')).join('').trim();
    }
    if (value.result != null) return String(value.result).trim();
  }
  return String(value).trim();
}

function looksLikeFormFKarnatakaMainTitle(text) {
  const n = formFKarnatakaHeaderNorm(text).replace(/[^a-z0-9 ]/g, ' ').replace(/\s+/g, ' ').trim();
  // Exact "FORM F" only — never "FORM 15", "FORM XV", etc.
  return n === 'form f' || /^form\s*f$/.test(n);
}

function looksLikeFormFKarnatakaReference(text) {
  const n = formFKarnatakaHeaderNorm(text);
  // Karnataka Form F is Rule 8 only — do not match Form 15 "(SEE RULE …)" lines.
  return /see\s+rule\s*8\b/.test(n);
}

function looksLikeFormFKarnatakaSubtitle(text) {
  const n = formFKarnatakaHeaderNorm(text);
  return /register\s+of\s+leave\s+with\s+wages/.test(n);
}

function looksLikeFormFKarnatakaPartBanner(text) {
  const n = formFKarnatakaHeaderNorm(text);
  return /part\s*[-–]?\s*i\b/.test(n) && /earned\s+leave/.test(n);
}

function looksLikeFormFKarnatakaPartIiBanner(text) {
  const n = formFKarnatakaHeaderNorm(text);
  // Require Sick/Accident wording so Form 15 "PART II" wage registers are not rewritten.
  return /part\s*[-–]?\s*ii/.test(n) && /sick|accident/.test(n);
}

function looksLikeForm15LeaveRegisterSheet(worksheet) {
  if (!worksheet) return false;
  const name = String(worksheet.name || '').toLowerCase();
  if (/form[\s._-]*15\b/.test(name) || /form[\s._-]*xv\b/.test(name)) return true;
  for (let r = 1; r <= 8; r += 1) {
    for (let c = 1; c <= 16; c += 1) {
      const t = formFExcelJsCellText(worksheet.getCell(r, c)?.value);
      if (!t) continue;
      const n = formFKarnatakaHeaderNorm(t).replace(/[^a-z0-9 ]/g, ' ').replace(/\s+/g, ' ').trim();
      if (/^form\s*15\b/.test(n) || /^form\s*xv\b/.test(n)) return true;
    }
  }
  return false;
}

function worksheetLooksLikeFormFKarnataka(worksheet) {
  if (!worksheet) return false;
  // Form 15 Part I/II share "REGISTER OF LEAVE WITH WAGES" — never treat as Form F.
  if (looksLikeForm15LeaveRegisterSheet(worksheet)) return false;

  const name = String(worksheet.name || '').toLowerCase();
  const sheetNameIsFormF =
    (/\bform[\s._-]*f\b/.test(name) || /^form\s*f\b/.test(name)) &&
    !/form[\s._-]*15\b/.test(name) &&
    !/form[\s._-]*xv\b/.test(name);

  let hasMainTitle = false;
  let hasRule8 = false;
  let hasSubtitle = false;
  let hasPartI = false;
  for (let r = 1; r <= 12; r += 1) {
    for (let c = 1; c <= 16; c += 1) {
      const t = formFExcelJsCellText(worksheet.getCell(r, c)?.value);
      if (!t) continue;
      if (looksLikeFormFKarnatakaMainTitle(t)) hasMainTitle = true;
      if (looksLikeFormFKarnatakaReference(t)) hasRule8 = true;
      if (looksLikeFormFKarnatakaSubtitle(t)) hasSubtitle = true;
      if (looksLikeFormFKarnatakaPartBanner(t)) hasPartI = true;
    }
  }

  // Require an explicit FORM F title (or Form F sheet name) so finalize-download
  // cannot rewrite Form 15 / other leave registers into FORM F.
  if (hasMainTitle && (hasRule8 || hasSubtitle || hasPartI || sheetNameIsFormF)) return true;
  if (sheetNameIsFormF && (hasRule8 || hasSubtitle || hasPartI)) return true;
  return false;
}

function parseFormFMergeLabel(label) {
  const m = String(label || '').match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/i);
  if (!m) return null;
  const colToNum = (letters) => {
    let n = 0;
    const s = String(letters).toUpperCase();
    for (let i = 0; i < s.length; i += 1) n = n * 26 + (s.charCodeAt(i) - 64);
    return n;
  };
  return {
    r1: parseInt(m[2], 10),
    c1: colToNum(m[1]),
    r2: parseInt(m[4], 10),
    c2: colToNum(m[3]),
  };
}

/**
 * Restore Form F title band centered across A–K (matches original template layout).
 * SheetJS/ExcelJS round-trips often leave FORM F / SEE RULE 8 / REGISTER… stuck in column L.
 */
export function ensureFormFKarnatakaTitleLayout(worksheet, options = {}) {
  if (!worksheet || !worksheetLooksLikeFormFKarnataka(worksheet)) return false;

  const colFrom = Math.max(1, Number(options.colFrom) || FORM_F_KARNATAKA_TITLE_COL_FROM);
  const colTo = Math.max(colFrom, Number(options.colTo) || FORM_F_KARNATAKA_TITLE_COL_TO);
  const titleText = String(options.title || FORM_F_KARNATAKA_DEFAULT_TITLE).trim();
  const referenceText = String(options.reference || FORM_F_KARNATAKA_DEFAULT_REFERENCE).trim();
  const subtitleText = String(options.subtitle || FORM_F_KARNATAKA_DEFAULT_SUBTITLE).trim();

  const findRow = (matcher, scanTo = 14) => {
    for (let r = 1; r <= scanTo; r += 1) {
      for (let c = 1; c <= Math.max(colTo + 4, 16); c += 1) {
        const t = formFExcelJsCellText(worksheet.getCell(r, c)?.value);
        if (matcher(t)) return r;
      }
    }
    return -1;
  };

  // Only rewrite rows that already exist on the template — never invent FORM F
  // titles onto an unrelated leave/wage register (e.g. Form 15 Part II).
  const titleRow = findRow(looksLikeFormFKarnatakaMainTitle, 6);
  const referenceRow = findRow(looksLikeFormFKarnatakaReference, 6);
  const subtitleRow = findRow(looksLikeFormFKarnatakaSubtitle, 8);
  const partIRow = findRow(looksLikeFormFKarnatakaPartBanner, 16);
  const partIiRow = findRow(looksLikeFormFKarnatakaPartIiBanner, 40);
  if (titleRow < 1 && referenceRow < 1 && subtitleRow < 1) return false;

  const mergeLabels = Array.isArray(worksheet.model?.merges) ? [...worksheet.model.merges] : [];
  const unmergeCovering = (r1, c1, r2, c2) => {
    mergeLabels.forEach((label) => {
      const m = parseFormFMergeLabel(label);
      if (!m) return;
      const overlaps = !(m.r2 < r1 || m.r1 > r2 || m.c2 < c1 || m.c1 > c2);
      if (!overlaps) return;
      try {
        worksheet.unMergeCells(label);
      } catch (_) {
        /* ignore */
      }
    });
  };

  const writeCenteredBand = (row, text, { bold = true, size = 12 } = {}) => {
    if (row < 1 || !text) return;
    // Clear stray copies (e.g. column L) across a wide scan before re-centering.
    for (let c = 1; c <= Math.max(colTo + 6, 20); c += 1) {
      const cell = worksheet.getCell(row, c);
      const t = formFExcelJsCellText(cell?.value);
      if (
        looksLikeFormFKarnatakaMainTitle(t) ||
        looksLikeFormFKarnatakaReference(t) ||
        looksLikeFormFKarnatakaSubtitle(t) ||
        looksLikeFormFKarnatakaPartBanner(t) ||
        looksLikeFormFKarnatakaPartIiBanner(t)
      ) {
        cell.value = null;
      }
    }
    unmergeCovering(row, colFrom, row, colTo);
    try {
      worksheet.mergeCells(row, colFrom, row, colTo);
    } catch (_) {
      /* overlap — still write left cell */
    }
    const cell = worksheet.getCell(row, colFrom);
    cell.value = text;
    cell.alignment = {
      ...(cell.alignment || {}),
      horizontal: 'center',
      vertical: 'middle',
      wrapText: true,
    };
    cell.font = {
      ...(cell.font || {}),
      bold,
      size: cell.font?.size || size,
      name: cell.font?.name || 'Calibri',
    };
  };

  const captureRowText = (row, matcher) => {
    if (row < 1) return '';
    for (let c = 1; c <= 20; c += 1) {
      const t = formFExcelJsCellText(worksheet.getCell(row, c)?.value);
      if (matcher(t)) return t;
    }
    return '';
  };

  writeCenteredBand(titleRow, titleText, { bold: true, size: 14 });
  writeCenteredBand(referenceRow, referenceText, { bold: false, size: 11 });
  writeCenteredBand(subtitleRow, subtitleText, { bold: true, size: 12 });
  if (partIRow > 0) {
    const captured = captureRowText(partIRow, looksLikeFormFKarnatakaPartBanner);
    writeCenteredBand(partIRow, captured || 'PART I EARNED LEAVE', { bold: true, size: 11 });
  }
  if (!options.skipPartIi && partIiRow > 0) {
    const captured = captureRowText(partIiRow, looksLikeFormFKarnatakaPartIiBanner);
    writeCenteredBand(
      partIiRow,
      captured || 'PART II - Sick/Accident Leave (with pay)',
      { bold: true, size: 11 }
    );
  }

  return true;
}

/** Template PART I body holds 15 employee leave rows. */
export const FORM_F_KARNATAKA_TEMPLATE_BODY_ROWS = 15;

function formFSanitizeFolderName(name) {
  const s = String(name || 'Unknown_Location')
    .replace(/[\\/:*?"<>|]+/g, '_')
    .replace(/\s+/g, '_')
    .replace(/_+/g, '_')
    .replace(/^_|_$/g, '')
    .trim();
  return s || 'Unknown_Location';
}

function formFScalarText(value) {
  if (value == null || value === '') return '';
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value.toISOString().slice(0, 10);
  if (typeof value === 'object') {
    const nested =
      value.display_value ??
      value.displayValue ??
      value.zc_display_value ??
      value.name ??
      value.Name ??
      value.value ??
      value.label ??
      '';
    if (nested instanceof Date && !Number.isNaN(nested.getTime())) {
      return nested.toISOString().slice(0, 10);
    }
    const text = String(nested ?? '').trim();
    if (text && text !== '[object Object]') return text;
    return '';
  }
  const text = String(value).trim();
  return text === '[object Object]' ? '' : text;
}

const pickFormFEmployeeValue = (emp, keys) => {
  if (!emp || typeof emp !== 'object') return '';
  const src = emp.Employee && typeof emp.Employee === 'object' ? emp.Employee : emp.employee || emp;
  const bags = src !== emp ? [src, emp] : [src];
  for (const bag of bags) {
    if (!bag || typeof bag !== 'object') continue;
    for (const key of keys) {
      const text = formFScalarText(bag[key]);
      if (text) return text;
    }
  }
  const normalizeKey = (k) => String(k || '').toLowerCase().replace(/[^a-z0-9]/g, '');
  const wanted = new Set((keys || []).map(normalizeKey).filter(Boolean));
  for (const bag of bags) {
    if (!bag || typeof bag !== 'object') continue;
    const found = Object.keys(bag).find((k) => wanted.has(normalizeKey(k)) && formFScalarText(bag[k]));
    if (found) return formFScalarText(bag[found]);
  }
  return '';
};

const pickFormFEmployeeValueByKeyPattern = (emp, patterns) => {
  if (!emp || typeof emp !== 'object' || !Array.isArray(patterns)) return '';
  const src = emp.Employee && typeof emp.Employee === 'object' ? emp.Employee : emp.employee || emp;
  const bags = src !== emp ? [src, emp] : [src];
  for (const bag of bags) {
    if (!bag || typeof bag !== 'object') continue;
    const keys = Object.keys(bag);
    for (const re of patterns) {
      const hit = keys.find((k) => re.test(String(k || '')));
      if (!hit) continue;
      const text = formFScalarText(bag[hit]);
      if (text) return text;
    }
  }
  return '';
};

export function formatFormFKarnatakaEmployeeName(emp = {}) {
  const src = emp?.Employee || emp?.employee || emp || {};
  const fn = String(src.FirstName || src.firstName || src['First Name'] || '').trim();
  const ln = String(src.LastName || src.lastName || src['Last Name'] || '').trim();
  if (fn && ln) return `${fn} ${ln}`;
  return (
    fn ||
    ln ||
    String(
      src.Employee_Name ||
        src['Employee Name'] ||
        src.EmployeeName ||
        src.name ||
        src.Name ||
        ''
    ).trim()
  );
}

/** Header identity fields above PART I on Karnataka Form F. */
export const FORM_F_KARNATAKA_IDENTITY_FIELD_SPECS = [
  {
    key: 'slNo',
    label: 'Sl No in the Register of Adult/young person',
    match: (n) =>
      (/s[li]\s*no/.test(n) && /register/.test(n) && /(adult|young)/.test(n)) ||
      /^\d+\.?\s*s[li]\s*no\s+in\s+the/.test(n),
  },
  {
    key: 'dateOfEntry',
    label: 'Date of entry into service',
    match: (n) => /date(?:\s+of)?\s+entry\s+into\s+service/.test(n) || /entry\s+into\s+service/.test(n),
  },
  {
    key: 'personName',
    label: 'Name of the person',
    match: (n) => {
      if (/father/.test(n) || /husband/.test(n) || /establishment/.test(n)) return false;
      return /name\s+of\s+the\s+person/.test(n) || /^\d+\.?\s*name\s+of\s+the/.test(n);
    },
  },
  {
    key: 'fatherName',
    label: "Father's Name",
    match: (n) => (/father/.test(n) && /name/.test(n)) || /husband'?s\s+name/.test(n),
  },
];

/**
 * Resolve identity header values for one employee Form F register.
 */
export function resolveFormFKarnatakaEmployeeIdentityValues(empItem, options = {}) {
  const emp = empItem?.Employee || empItem?.employee || empItem || {};
  const row = options.row && typeof options.row === 'object' ? options.row : null;
  const formatDate =
    typeof options.formatStatutoryDateDisplay === 'function'
      ? options.formatStatutoryDateDisplay
      : (v) => String(v ?? '').trim();
  const rowIndex = Number(options.rowIndex) >= 0 ? Number(options.rowIndex) : 0;

  const personName =
    formatFormFKarnatakaEmployeeName(emp) ||
    String(
      row?.__employeeLookupName ||
        row?.['Name of the person'] ||
        row?.Name ||
        row?.['Employee Name'] ||
        ''
    ).trim();

  const fatherName =
    pickFormFEmployeeValue(emp, [
      'Father_s_Name',
      "Father's_Name",
      "Father's Name",
      'Father_s Name',
      'Father_Name',
      'FatherName',
      'Father Name',
      'FathersName',
      'fathersName',
      'Father',
      'Spouse_Name',
      'SpouseName',
      'Spouse Name',
    ]) ||
    pickFormFEmployeeValueByKeyPattern(emp, [/^father/i, /father.*name/i, /spouse.*name/i]) ||
    String(row?.["Father's Name"] || row?.FatherName || row?.Father_s_Name || '').trim();

  const dojRaw =
    pickFormFEmployeeValue(emp, [
      'Dateofjoining',
      'DateofJoining',
      'Date_of_Joining',
      'Date of Joining',
      'dateOfJoining',
      'DateofEntryintoService',
      'Date of entry into service',
      'JoiningDate',
      'joiningDate',
      'DOJ',
      'doj',
    ]) ||
    pickFormFEmployeeValueByKeyPattern(emp, [
      /dateofjoining/i,
      /date_of_joining/i,
      /joiningdate/i,
      /entry.*service/i,
      /^doj$/i,
    ]) ||
    String(row?.['Date of entry into service'] || '').trim();

  const slNo =
    pickFormFEmployeeValue(emp, [
      'Employee_ID',
      'Employee ID',
      'EmployeeID',
      'EmployeeId',
      'employeeId',
      'EmployeeCode',
      'Employee Code',
      'Sl_No',
      'Sl No',
      'SerialNo',
      'Serial Number',
    ]) ||
    String(row?.__employeeLookupId || row?.['Employee ID'] || '').trim() ||
    String(rowIndex + 1);

  return {
    slNo,
    dateOfEntry: dojRaw ? formatDate(dojRaw) : '',
    personName,
    fatherName,
  };
}

function formFIdentityLabelLooksLikeOtherField(text, currentKey) {
  const n = formFKarnatakaHeaderNorm(text);
  return FORM_F_KARNATAKA_IDENTITY_FIELD_SPECS.some(
    (spec) => spec.key !== currentKey && spec.match(n)
  );
}

function formFKarnatakaListedMerges(worksheet) {
  const labels = Array.isArray(worksheet?.model?.merges) ? worksheet.model.merges : [];
  return labels.map(parseFormFMergeLabel).filter(Boolean);
}

function formFKarnatakaMergeAt(merges, row, col) {
  return merges.find((m) => row >= m.r1 && row <= m.r2 && col >= m.c1 && col <= m.c2) || null;
}

function formFKarnatakaMergeMaster(merges, row, col) {
  const m = formFKarnatakaMergeAt(merges, row, col);
  if (!m) return { row, col };
  return { row: m.r1, col: m.c1 };
}

function formFIdentityCellLooksEmpty(text) {
  const t = String(text || '').trim();
  return !t || t === ':' || /^_+$/.test(t);
}

function formFIdentityRowLooksLikeBanner(worksheet, row, maxCols) {
  if (!worksheet || row < 1) return false;
  for (let c = 1; c <= maxCols; c += 1) {
    const t = formFExcelJsCellText(worksheet.getCell(row, c)?.value);
    if (
      looksLikeFormFKarnatakaMainTitle(t) ||
      looksLikeFormFKarnatakaReference(t) ||
      looksLikeFormFKarnatakaSubtitle(t) ||
      looksLikeFormFKarnatakaPartBanner(t) ||
      looksLikeFormFKarnatakaPartIiBanner(t)
    ) {
      return true;
    }
  }
  return false;
}

/**
 * Write Sl No / Date of entry / Name / Father's Name beside their template labels.
 * Right-side labels are often merged across H–K; never write into that merge (Excel hides it).
 */
export function writeFormFKarnatakaIdentityFieldsToWorksheet(worksheet, identityValues, options = {}) {
  if (!worksheet || !identityValues || typeof identityValues !== 'object') return 0;
  const maxScanRows = Math.min(
    Number(options.maxScanRows) || 20,
    Math.max(20, Number(worksheet.rowCount) || 20)
  );
  const maxScanCols = Math.min(20, Number(options.maxScanCols) || 16);
  const merges = formFKarnatakaListedMerges(worksheet);
  const formRightEdge = 11;
  let written = 0;

  const writeAt = (row, col, value) => {
    const text = String(value ?? '').trim();
    if (!text || row < 1 || col < 1) return false;
    const cell = worksheet.getCell(row, col);
    cell.value = text;
    cell.alignment = {
      ...(cell.alignment || {}),
      vertical: 'middle',
      wrapText: true,
      horizontal: 'left',
    };
    return true;
  };

  const isInsideLabelMerge = (row, col, labelRow, labelCol) => {
    const labelMerge = formFKarnatakaMergeAt(merges, labelRow, labelCol);
    if (!labelMerge) return false;
    return row >= labelMerge.r1 && row <= labelMerge.r2 && col >= labelMerge.c1 && col <= labelMerge.c2;
  };

  const tryWriteValueCell = (row, col, spec, labelRow, labelCol, value) => {
    if (row < 1 || col < 1) return false;
    if (isInsideLabelMerge(row, col, labelRow, labelCol)) return false;
    if (formFIdentityRowLooksLikeBanner(worksheet, row, maxScanCols)) return false;
    const master = formFKarnatakaMergeMaster(merges, row, col);
    if (isInsideLabelMerge(master.row, master.col, labelRow, labelCol)) return false;
    const existing = formFExcelJsCellText(worksheet.getCell(master.row, master.col)?.value);
    if (formFIdentityLabelLooksLikeOtherField(existing, spec.key)) return false;
    if (!formFIdentityCellLooksEmpty(existing)) return false;
    return writeAt(master.row, master.col, value);
  };

  FORM_F_KARNATAKA_IDENTITY_FIELD_SPECS.forEach((spec) => {
    const value = String(identityValues[spec.key] ?? '').trim();
    if (!value) return;

    let bestRow = -1;
    let bestCol = -1;
    let bestLen = Infinity;
    for (let r = 1; r <= maxScanRows; r += 1) {
      for (let c = 1; c <= maxScanCols; c += 1) {
        const mergeAt = formFKarnatakaMergeAt(merges, r, c);
        if (mergeAt && (r !== mergeAt.r1 || c !== mergeAt.c1)) continue;
        const raw = formFExcelJsCellText(worksheet.getCell(r, c)?.value);
        if (!raw) continue;
        const n = formFKarnatakaHeaderNorm(raw);
        if (!spec.match(n)) continue;
        // Prefer the shortest matching label cell (avoids long merged title hits).
        if (raw.length < bestLen) {
          bestLen = raw.length;
          bestRow = r;
          bestCol = c;
        }
      }
    }
    if (bestRow < 1 || bestCol < 1) return;

    const labelMerge = formFKarnatakaMergeAt(merges, bestRow, bestCol);
    const rightStart = labelMerge ? labelMerge.c2 + 1 : bestCol + 1;
    const rightEnd = Math.min(bestCol + 6, formRightEdge, maxScanCols + 2);

    // Prefer empty cell to the right of the label (skip the label's own merge).
    for (let nc = rightStart; nc <= rightEnd; nc += 1) {
      const master = formFKarnatakaMergeMaster(merges, bestRow, nc);
      const nt = formFExcelJsCellText(worksheet.getCell(master.row, master.col)?.value);
      if (formFIdentityLabelLooksLikeOtherField(nt, spec.key)) break;
      if (tryWriteValueCell(bestRow, nc, spec, bestRow, bestCol, value)) {
        written += 1;
        return;
      }
    }

    // Karnataka template: value boxes sit above numbered labels (H5 above "Date of entry").
    if (tryWriteValueCell(bestRow - 1, bestCol, spec, bestRow, bestCol, value)) {
      written += 1;
      return;
    }

    // Fallback: append after the label in the same cell (keeps "Label : Value").
    const labelRaw = formFExcelJsCellText(worksheet.getCell(bestRow, bestCol)?.value);
    const base = String(labelRaw || '').replace(/\s+$/, '');
    const withColon = /:\s*$/.test(base) ? `${base} ${value}` : `${base} : ${value}`;
    if (writeAt(bestRow, bestCol, withColon)) written += 1;
  });

  return written;
}

export function resolveFormFKarnatakaEmployeeLocation(empItem, row = null) {
  const emp =
    empItem && typeof empItem === 'object'
      ? empItem.Employee || empItem.employee || empItem
      : null;
  const fromRow = String(row?.__employeeLocation || row?.Location || row?.Work_location || '').trim();
  if (fromRow) return fromRow;
  if (!emp || typeof emp !== 'object') return 'Unknown_Location';
  const raw =
    emp.LocationName ??
    emp['Location Name'] ??
    emp.Work_location ??
    emp['Work Location'] ??
    emp.work_location ??
    emp.workLocation ??
    emp.Location ??
    emp.location ??
    '';
  if (raw && typeof raw === 'object') {
    return String(raw.display_value ?? raw.name ?? raw.Name ?? '').trim() || 'Unknown_Location';
  }
  return String(raw || '').trim() || 'Unknown_Location';
}

function formFKarnatakaEmployeeIdentity(empItem) {
  const emp = empItem?.Employee || empItem?.employee || empItem;
  if (!emp || typeof emp !== 'object') return { id: '', name: '' };
  const id = String(
    emp.Employee_ID || emp['Employee ID'] || emp.employeeId || emp.Zoho_ID || emp.ZohoID || ''
  )
    .trim()
    .toLowerCase();
  const first = String(emp.FirstName || emp.firstName || '').trim();
  const last = String(emp.LastName || emp.lastName || '').trim();
  const name = normalizePersonNameKey(
    emp.Employee_Name || emp['Employee Name'] || emp.name || `${first} ${last}`
  );
  return { id, name };
}

function formFKarnatakaRowIdentity(row) {
  if (!row || typeof row !== 'object') return { id: '', name: '' };
  return {
    id: String(row.__employeeLookupId || row['Employee ID'] || row.EmployeeID || '')
      .trim()
      .toLowerCase(),
    name: normalizePersonNameKey(
      row['Name of the person'] ||
        row.Name ||
        row['Employee Name'] ||
        row.__employeeLookupName ||
        ''
    ),
  };
}

function formFKarnatakaIdentitiesMatch(rowId, rowName, empId, empName) {
  if (rowId && empId && rowId === empId) return true;
  if (rowName && empName && formFKarnatakaPersonNamesMatchStrict(rowName, empName)) return true;
  return false;
}

/**
 * Keep only mapped rows that belong to `employees` (site/location filter safe).
 * When lengths match, treat autofill order as aligned and stamp location metadata.
 */
export function alignFormFKarnatakaMappedRowsToEmployees(mappedData, employeesForMapping = []) {
  const rows = Array.isArray(mappedData) ? mappedData.filter((r) => r && typeof r === 'object') : [];
  const employees = Array.isArray(employeesForMapping) ? employeesForMapping : [];
  if (employees.length === 0) return rows;

  const stamp = (row, empItem) => {
    const loc = resolveFormFKarnatakaEmployeeLocation(empItem, row);
    const { id, name } = formFKarnatakaEmployeeIdentity(empItem);
    return {
      ...row,
      __employeeLocation: loc,
      ...(name ? { __employeeLookupName: row.__employeeLookupName || name } : {}),
      ...(id ? { __employeeLookupId: row.__employeeLookupId || id } : {}),
    };
  };

  if (rows.length === employees.length) {
    return rows.map((row, i) => stamp(row, employees[i]));
  }

  const used = new Set();
  const aligned = [];
  employees.forEach((empItem) => {
    const { id: empId, name: empName } = formFKarnatakaEmployeeIdentity(empItem);
    let matchIdx = -1;
    for (let i = 0; i < rows.length; i += 1) {
      if (used.has(i)) continue;
      const { id: rowId, name: rowName } = formFKarnatakaRowIdentity(rows[i]);
      if (formFKarnatakaIdentitiesMatch(rowId, rowName, empId, empName)) {
        matchIdx = i;
        break;
      }
    }
    if (matchIdx >= 0) {
      used.add(matchIdx);
      aligned.push(stamp(rows[matchIdx], empItem));
    }
  });
  return aligned.length > 0 ? aligned : rows.map((row, i) => stamp(row, employees[i] || null));
}

/**
 * Group Form F export rows by employee work location (for ZIP folders).
 * Returns [{ location, rows: [{ row, emp, index }] }]
 */
export function groupFormFKarnatakaRowsByLocation(mappedData, employeesForMapping = []) {
  const employees = Array.isArray(employeesForMapping) ? employeesForMapping : [];
  const alignedRows = alignFormFKarnatakaMappedRowsToEmployees(mappedData, employees);
  const findEmpForRow = (row, index) => {
    if (employees.length === alignedRows.length && employees[index]) return employees[index];
    if (employees[index] && employees.length === (Array.isArray(mappedData) ? mappedData.length : 0)) {
      return employees[index];
    }
    const { id: rowId, name: rowName } = formFKarnatakaRowIdentity(row);
    if (!rowName && !rowId) return employees[index] || null;
    return (
      employees.find((empItem) => {
        const { id, name } = formFKarnatakaEmployeeIdentity(empItem);
        return formFKarnatakaIdentitiesMatch(rowId, rowName, id, name);
      }) ||
      employees[index] ||
      null
    );
  };

  const groups = new Map();
  alignedRows.forEach((row, index) => {
    if (!row || typeof row !== 'object') return;
    const empItem = findEmpForRow(row, index);
    const location = resolveFormFKarnatakaEmployeeLocation(empItem, row);
    const key = formFSanitizeFolderName(location);
    if (!groups.has(key)) groups.set(key, { location: location || key, rows: [] });
    groups.get(key).rows.push({ row, emp: empItem, index });
  });
  return [...groups.values()];
}

/** Split location rows into chunks of `size` (default 15 template body rows). */
export function chunkFormFKarnatakaRowsForTemplate(locationRows, size = FORM_F_KARNATAKA_TEMPLATE_BODY_ROWS) {
  const list = Array.isArray(locationRows) ? locationRows : [];
  const chunkSize = Math.max(1, Number(size) || FORM_F_KARNATAKA_TEMPLATE_BODY_ROWS);
  const chunks = [];
  for (let i = 0; i < list.length; i += chunkSize) {
    chunks.push(list.slice(i, i + chunkSize));
  }
  return chunks.length > 0 ? chunks : [[]];
}

/**
 * Locate PART I leave table on the Form F worksheet (data starts inside the bordered body).
 * bodyRows never extends into PART II — that section must survive download intact.
 */
export function detectFormFKarnatakaSheetLayout(worksheet) {
  if (!worksheet) return null;
  const maxScanRows = Math.min(80, Number(worksheet.rowCount) || 80);
  const maxScanCols = 20;

  let partIRow = -1;
  let partIiRow = -1;
  for (let r = 1; r <= maxScanRows; r += 1) {
    let line = '';
    for (let c = 1; c <= maxScanCols; c += 1) {
      const t = formFExcelJsCellText(worksheet.getCell(r, c)?.value);
      if (t) line += ` ${t}`;
    }
    const n = formFKarnatakaHeaderNorm(line);
    if (partIRow < 0 && /part\s*[-–]?\s*i\b/.test(n) && /earned\s+leave/.test(n)) partIRow = r;
    if (partIiRow < 0 && /part\s*[-–]?\s*ii/.test(n)) {
      partIiRow = r;
      break;
    }
  }

  const scanFrom = partIRow > 0 ? partIRow + 1 : 1;
  const scanTo = partIiRow > 0 ? partIiRow - 1 : Math.min(maxScanRows, scanFrom + 20);

  let leafHeaderRow = -1;
  let numberHeaderRow = -1;
  let colFrom = 1;
  let colTo = 11;

  for (let r = scanFrom; r <= scanTo; r += 1) {
    const cells = [];
    for (let c = 1; c <= maxScanCols; c += 1) {
      cells.push(formFExcelJsCellText(worksheet.getCell(r, c)?.value));
    }
    const digitCount = cells.filter((v) => /^\d{1,2}$/.test(String(v || '').trim())).length;
    if (digitCount >= 8) {
      numberHeaderRow = r;
      const firstDigitCol = cells.findIndex((v) => /^\d{1,2}$/.test(String(v || '').trim()));
      if (firstDigitCol >= 0) colFrom = firstDigitCol + 1;
      colTo = colFrom + FORM_F_KARNATAKA_PART_I_HEADERS.length - 1;
      break;
    }
    const joined = cells.join(' ').toLowerCase();
    if (
      /leave\s+earned/.test(joined) ||
      (/from/.test(joined) && /to/.test(joined) && /total\s+days\s+worked|leave\s+at\s+credit/.test(joined))
    ) {
      leafHeaderRow = r;
      const firstTextCol = cells.findIndex((v) => String(v || '').trim() !== '');
      if (firstTextCol >= 0) colFrom = firstTextCol + 1;
      colTo = colFrom + FORM_F_KARNATAKA_PART_I_HEADERS.length - 1;
    }
  }

  let dataStartRow = -1;
  if (numberHeaderRow > 0) dataStartRow = numberHeaderRow + 1;
  else if (leafHeaderRow > 0) dataStartRow = leafHeaderRow + 1;
  else if (partIRow > 0) dataStartRow = partIRow + 3;

  if (dataStartRow < 1) return null;

  // Hard stop before PART II title / table — never clear or border over Sick leave section.
  const maxBodyBeforePartIi =
    partIiRow > dataStartRow ? Math.max(1, partIiRow - dataStartRow) : FORM_F_KARNATAKA_TEMPLATE_BODY_ROWS;
  const maxBody = Math.min(maxBodyBeforePartIi, FORM_F_KARNATAKA_TEMPLATE_BODY_ROWS);
  let bodyRows = countExcelJSTemplateBodyRows(
    worksheet,
    dataStartRow,
    colFrom,
    colTo,
    maxBody
  );
  bodyRows = Math.min(Math.max(bodyRows, 1), maxBody);

  return {
    partIRow,
    partIiRow,
    leafHeaderRow,
    numberHeaderRow,
    dataStartRow,
    colFrom,
    colTo,
    bodyRows,
  };
}

function readFormFRowCellValue(row, header, colIndex) {
  if (!row || typeof row !== 'object') return '';
  const storageKey = getFormFKarnatakaStorageHeader(colIndex, header);
  const primary = row[storageKey];
  if (primary != null && String(primary).trim() !== '') return String(primary).trim();
  if (header && row[header] != null && String(row[header]).trim() !== '') {
    return String(row[header]).trim();
  }
  const canon = FORM_F_KARNATAKA_PART_I_HEADERS[colIndex];
  if (canon && row[canon] != null && String(row[canon]).trim() !== '') {
    return String(row[canon]).trim();
  }
  return '';
}

/** Capture PART II title + next few header/data rows so download cannot wipe Sick leave. */
function captureFormFKarnatakaPartIiSnapshot(worksheet, partIiRow) {
  if (!worksheet || !(partIiRow > 0)) return null;
  const rows = [];
  const maxCols = 12;
  const rowCount = 6; // title + group header + leaf header + a few body rows
  for (let r = 0; r < rowCount; r += 1) {
    const excelRow = partIiRow + r;
    const cells = [];
    for (let c = 1; c <= maxCols; c += 1) {
      const cell = worksheet.getCell(excelRow, c);
      cells.push({
        col: c,
        value: cell.value == null ? null : cell.value,
      });
    }
    rows.push({ row: excelRow, cells });
  }
  return { partIiRow, rows };
}

function restoreFormFKarnatakaPartIiSnapshot(worksheet, snapshot) {
  if (!worksheet || !snapshot?.rows?.length) return;
  snapshot.rows.forEach((rowSnap) => {
    (rowSnap.cells || []).forEach((cellSnap) => {
      worksheet.getCell(rowSnap.row, cellSnap.col).value = cellSnap.value;
    });
  });
}

/**
 * Write PART II Year / of credit / Availed / Balance into the Sick leave table
 * (first data row under PART II headers).
 */
export function writeFormFKarnatakaPartIiValuesToWorksheet(worksheet, partIiRow, partIiValues) {
  if (!worksheet || !(partIiRow > 0) || !partIiValues || typeof partIiValues !== 'object') return 0;
  const maxScanCols = 12;
  const scanTo = Math.min(partIiRow + 8, (Number(worksheet.rowCount) || partIiRow) + 8);

  let headerRow = -1;
  let yearCol = -1;
  let creditCol = -1;
  let availedCol = -1;
  let balanceCol = -1;

  for (let r = partIiRow; r <= scanTo; r += 1) {
    const cells = [];
    for (let c = 1; c <= maxScanCols; c += 1) {
      cells.push(formFExcelJsCellText(worksheet.getCell(r, c)?.value));
    }
    const joined = cells.join(' ').toLowerCase();
    const hasYear = cells.some((t) => /^year$/i.test(String(t || '').trim()));
    const hasCredit = /credit/.test(joined);
    const hasAvailed = /availed/.test(joined);
    const hasBalance = /balance/.test(joined);
    if (hasYear && (hasCredit || hasAvailed || hasBalance)) {
      headerRow = r;
      cells.forEach((t, idx) => {
        const n = formFKarnatakaHeaderNorm(t);
        if (/^year$/.test(n)) yearCol = idx + 1;
        else if (/credit/.test(n) && creditCol < 0) creditCol = idx + 1;
        else if (/^availed$/.test(n) || (/availed/.test(n) && !/leave/.test(n))) availedCol = idx + 1;
        else if (/balance/.test(n)) balanceCol = idx + 1;
      });
      // Template often has group "Sick/Accident leave" over "of credit" / "Availed"
      if (creditCol < 0) {
        const creditIdx = cells.findIndex((t) => /credit/i.test(t));
        if (creditIdx >= 0) creditCol = creditIdx + 1;
      }
      if (availedCol < 0) {
        const availedIdx = cells.findIndex((t) => /availed/i.test(t));
        if (availedIdx >= 0) availedCol = availedIdx + 1;
      }
      break;
    }
  }

  if (headerRow < 0) {
    // Fallback: title row + 2 → data (common Form F layout).
    headerRow = partIiRow + 1;
    yearCol = 2;
    creditCol = 3;
    availedCol = 4;
    balanceCol = 5;
  }

  const dataRow = headerRow + 1;
  const write = (col, value) => {
    if (!(col > 0)) return false;
    const text = String(value ?? '').trim();
    if (!text) return false;
    const cell = worksheet.getCell(dataRow, col);
    cell.value = text;
    cell.alignment = {
      ...(cell.alignment || {}),
      vertical: 'middle',
      horizontal: /^-?\d+(\.\d+)?$/.test(text) ? 'right' : 'left',
      wrapText: true,
    };
    return true;
  };

  let written = 0;
  if (write(yearCol, partIiValues[FORM_F_KARNATAKA_PART_II_HEADERS[0]])) written += 1;
  if (write(creditCol, partIiValues[FORM_F_KARNATAKA_PART_II_HEADERS[1]])) written += 1;
  if (write(availedCol, partIiValues[FORM_F_KARNATAKA_PART_II_HEADERS[2]])) written += 1;
  if (write(balanceCol, partIiValues[FORM_F_KARNATAKA_PART_II_HEADERS[3]])) written += 1;
  return written;
}

/**
 * Write PART I rows into the original Form F template (keeps borders/merges).
 * Stops before PART II so Sick/Accident Leave section is never overwritten.
 */
export async function buildFormFKarnatakaWorkbookWithTemplateStyles({
  templateArrayBuffer,
  mappedData = [],
  headersToUse = [],
  parsedFormHeader,
  formFileName,
  sheetNameHint,
  employee = null,
  identityValues = null,
  formatStatutoryDateDisplay = null,
  employeeRowIndex = 0,
}) {
  const workbook = new ExcelJS.Workbook();
  const bytes =
    templateArrayBuffer instanceof ArrayBuffer
      ? templateArrayBuffer.slice(0)
      : templateArrayBuffer;
  await workbook.xlsx.load(bytes);
  const preferred = String(sheetNameHint || '').trim();
  const worksheet =
    (preferred && workbook.worksheets.find((ws) => String(ws?.name || '').trim() === preferred)) ||
    workbook.worksheets.find((ws) => /form\s*f/i.test(String(ws?.name || ''))) ||
    workbook.worksheets[0] ||
    null;
  if (!worksheet) throw new Error('Form F template worksheet not found.');

  // Detect + snapshot PART II from the pristine template before any rewrite.
  const layout = detectFormFKarnatakaSheetLayout(worksheet);
  if (!layout) throw new Error('Could not locate Form F PART I leave table.');
  const partIiSnapshot = captureFormFKarnatakaPartIiSnapshot(worksheet, layout.partIiRow);

  ensureFormFKarnatakaTitleLayout(worksheet, {
    colFrom: layout.colFrom,
    colTo: layout.colTo,
    skipPartIi: true,
  });

  const priorHeaders = Array.isArray(headersToUse) ? headersToUse : [];
  const canonHeaders = resolveFormFKarnatakaTableHeaders(priorHeaders);
  const sourceRows = remapFormFKarnatakaRowsToCanonicalHeaders(
    Array.isArray(mappedData) ? mappedData : [],
    priorHeaders
  ).slice(0, layout.bodyRows);

  const resolvedIdentity =
    identityValues && typeof identityValues === 'object'
      ? identityValues
      : resolveFormFKarnatakaEmployeeIdentityValues(employee, {
          row: sourceRows[0] || null,
          rowIndex: employeeRowIndex,
          formatStatutoryDateDisplay:
            typeof formatStatutoryDateDisplay === 'function'
              ? formatStatutoryDateDisplay
              : (v) => String(v ?? '').trim(),
        });

  // Clear only PART I body — never touch PART II rows.
  for (let i = 0; i < layout.bodyRows; i += 1) {
    const excelRow = layout.dataStartRow + i;
    if (layout.partIiRow > 0 && excelRow >= layout.partIiRow) break;
    for (let c = layout.colFrom; c <= layout.colTo; c += 1) {
      worksheet.getCell(excelRow, c).value = null;
    }
  }

  sourceRows.forEach((row, i) => {
    const excelRow = layout.dataStartRow + i;
    if (layout.partIiRow > 0 && excelRow >= layout.partIiRow) return;
    canonHeaders.forEach((header, colIndex) => {
      const value = readFormFRowCellValue(row, header, colIndex);
      if (value === '') return;
      const cell = worksheet.getCell(excelRow, layout.colFrom + colIndex);
      cell.value = value;
      cell.alignment = {
        ...(cell.alignment || {}),
        vertical: 'middle',
        wrapText: true,
        horizontal: /^-?\d+(\.\d+)?$/.test(value) ? 'right' : 'left',
      };
    });
  });

  const safeBorderRows =
    layout.partIiRow > layout.dataStartRow
      ? Math.min(layout.bodyRows, layout.partIiRow - layout.dataStartRow)
      : layout.bodyRows;
  ensureExcelJSDataRowsWithBorders(worksheet, {
    dataStartRow: layout.dataStartRow,
    dataRowCount: safeBorderRows,
    colFrom: layout.colFrom,
    colTo: layout.colTo,
    templateRow: layout.dataStartRow,
    templateBodyRows: safeBorderRows,
  });

  ensureFormFKarnatakaTitleLayout(worksheet, {
    colFrom: layout.colFrom,
    colTo: layout.colTo,
    skipPartIi: true,
  });
  // Always put original PART II back last so Sick/Accident Leave survives download.
  restoreFormFKarnatakaPartIiSnapshot(worksheet, partIiSnapshot);
  // Re-apply identity after title layout (title pass must not clear these labels).
  writeFormFKarnatakaIdentityFieldsToWorksheet(worksheet, resolvedIdentity, {
    maxScanRows: layout.partIRow > 0 ? layout.partIRow : 20,
  });

  const partIiFromRow = sourceRows[0] || {};
  const partIiValues = {
    [FORM_F_KARNATAKA_PART_II_HEADERS[0]]: partIiFromRow[FORM_F_KARNATAKA_PART_II_HEADERS[0]] || '',
    [FORM_F_KARNATAKA_PART_II_HEADERS[1]]: partIiFromRow[FORM_F_KARNATAKA_PART_II_HEADERS[1]] || '',
    [FORM_F_KARNATAKA_PART_II_HEADERS[2]]: partIiFromRow[FORM_F_KARNATAKA_PART_II_HEADERS[2]] || '',
    [FORM_F_KARNATAKA_PART_II_HEADERS[3]]: partIiFromRow[FORM_F_KARNATAKA_PART_II_HEADERS[3]] || '',
  };
  if (Object.values(partIiValues).some((v) => String(v || '').trim() !== '')) {
    writeFormFKarnatakaPartIiValuesToWorksheet(worksheet, layout.partIiRow, partIiValues);
  }

  const out = await workbook.xlsx.writeBuffer();
  const fileName =
    formFileName ||
    parsedFormHeader?.title?.replace(/[^a-zA-Z0-9]/g, '_') ||
    'Form_F_Karnataka.xlsx';
  return {
    blob: new Blob([out], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    }),
    fileName,
    layout,
    rowCount: sourceRows.length,
    identityValues: resolvedIdentity,
  };
}

function resolveFormFKarnatakaEmployeeDownloadBaseName(row, empItem, fallbackIndex = 0) {
  const fromRow = String(
    row?.__employeeLookupName ||
      row?.['Name of the person'] ||
      row?.Name ||
      row?.['Employee Name'] ||
      ''
  ).trim();
  const { id, name } = formFKarnatakaEmployeeIdentity(empItem);
  const raw = fromRow || name || id || '';
  const slug = String(raw)
    .replace(/[^\w\s.-]/g, '')
    .replace(/\s+/g, '_')
    .replace(/_+/g, '_')
    .replace(/^_|_$/g, '')
    .slice(0, 48);
  return slug || `Employee_${fallbackIndex + 1}`;
}

function allocateUniqueFormFKarnatakaDownloadFileName(baseName, usedNames) {
  const count = usedNames.get(baseName) || 0;
  usedNames.set(baseName, count + 1);
  const suffix = count > 0 ? `_${count + 1}` : '';
  return `Form_F_Karnataka_${baseName}${suffix}.xlsx`;
}

/**
 * One Form F workbook per employee (PART I has that employee's leave row only).
 * Single employee → one .xlsx; multiple → ZIP of employee files.
 */
export async function buildFormFKarnatakaPerEmployeeDownload({
  templateArrayBuffer,
  mappedData = [],
  headersToUse = [],
  parsedFormHeader,
  formFileName,
  sheetNameHint,
  employeesOverride = [],
  formatStatutoryDateDisplay = null,
}) {
  if (!templateArrayBuffer) {
    throw new Error(
      'Original Form F Karnataka template could not be loaded. Open Autofill again, then Download.'
    );
  }

  const employees = Array.isArray(employeesOverride) ? employeesOverride : [];
  const alignedRows = alignFormFKarnatakaMappedRowsToEmployees(mappedData, employees);
  const exportEntries = [];
  if (alignedRows.length > 0) {
    alignedRows.forEach((row, index) => {
      if (!row || typeof row !== 'object') return;
      const emp =
        employees.length === alignedRows.length
          ? employees[index]
          : employees.find((empItem) => {
              const { id: rowId, name: rowName } = formFKarnatakaRowIdentity(row);
              const { id, name } = formFKarnatakaEmployeeIdentity(empItem);
              return formFKarnatakaIdentitiesMatch(rowId, rowName, id, name);
            }) ||
            employees[index] ||
            null;
      exportEntries.push({ row, emp, index });
    });
  } else if (employees.length > 0) {
    employees.forEach((emp, index) => {
      exportEntries.push({ row: {}, emp, index });
    });
  }

  const zipBase = String(formFileName || parsedFormHeader?.title || 'Form_F_Karnataka')
    .replace(/\.xlsx?$/i, '')
    .replace(/[^a-zA-Z0-9._-]+/g, '_');
  const templateBytes =
    templateArrayBuffer instanceof ArrayBuffer
      ? templateArrayBuffer.slice(0)
      : templateArrayBuffer;
  const workbookArgs = {
    headersToUse,
    parsedFormHeader,
    sheetNameHint,
    formatStatutoryDateDisplay,
  };

  if (exportEntries.length <= 1) {
    const entry = exportEntries[0];
    const rows = entry?.row ? [entry.row] : [];
    const baseName = resolveFormFKarnatakaEmployeeDownloadBaseName(
      entry?.row,
      entry?.emp,
      entry?.index || 0
    );
    return buildFormFKarnatakaWorkbookWithTemplateStyles({
      ...workbookArgs,
      templateArrayBuffer:
        templateBytes instanceof ArrayBuffer ? templateBytes.slice(0) : templateBytes,
      mappedData: rows,
      employee: entry?.emp || null,
      employeeRowIndex: entry?.index || 0,
      formFileName: allocateUniqueFormFKarnatakaDownloadFileName(baseName, new Map()),
    });
  }

  const zip = new JSZip();
  const usedNames = new Map();
  for (let i = 0; i < exportEntries.length; i += 1) {
    const entry = exportEntries[i];
    const baseName = resolveFormFKarnatakaEmployeeDownloadBaseName(entry.row, entry.emp, entry.index);
    const entryName = allocateUniqueFormFKarnatakaDownloadFileName(baseName, usedNames);
    const { blob } = await buildFormFKarnatakaWorkbookWithTemplateStyles({
      ...workbookArgs,
      // Fresh copy per employee — avoids ArrayBuffer detach after workbook.xlsx.load.
      templateArrayBuffer:
        templateBytes instanceof ArrayBuffer ? templateBytes.slice(0) : templateBytes,
      mappedData: entry.row ? [entry.row] : [],
      employee: entry.emp || null,
      employeeRowIndex: entry.index,
      formFileName: entryName,
    });
    const xlsxBytes =
      typeof blob?.arrayBuffer === 'function'
        ? new Uint8Array(await blob.arrayBuffer())
        : new Uint8Array(await new Response(blob).arrayBuffer());
    zip.file(entryName, xlsxBytes);
    if (i > 0 && i % 5 === 0) {
      await new Promise((resolve) => setTimeout(resolve, 0));
    }
  }

  return {
    blob: await zip.generateAsync({
      type: 'blob',
      mimeType: 'application/zip',
      compression: 'STORE',
    }),
    fileName: `${zipBase}_Employees.zip`,
  };
}

/**
 * @deprecated Prefer buildFormFKarnatakaPerEmployeeDownload — Form F is one register per employee.
 * Location-wise ZIP kept for callers that still group by Work_location (≤15 PART I rows per xlsx).
 */
export async function buildFormFKarnatakaLocationWiseZipDownload({
  templateArrayBuffer,
  mappedData = [],
  headersToUse = [],
  parsedFormHeader,
  formFileName,
  sheetNameHint,
  employeesOverride = [],
}) {
  const employees = Array.isArray(employeesOverride) ? employeesOverride : [];
  const groups = groupFormFKarnatakaRowsByLocation(mappedData, employees);
  const zip = new JSZip();
  const zipBase = String(formFileName || parsedFormHeader?.title || 'Form_F_Karnataka')
    .replace(/\.xlsx?$/i, '')
    .replace(/[^a-zA-Z0-9._-]+/g, '_');

  const exportGroups =
    groups.length > 0
      ? groups
      : [{ location: 'Unknown_Location', rows: [] }];

  let fileCount = 0;
  for (let g = 0; g < exportGroups.length; g += 1) {
    const group = exportGroups[g];
    const folderName = formFSanitizeFolderName(group.location);
    const folder = zip.folder(folderName);
    const chunks = chunkFormFKarnatakaRowsForTemplate(group.rows);
    for (let c = 0; c < chunks.length; c += 1) {
      const chunkRows = chunks[c].map((entry) => (entry && entry.row != null ? entry.row : entry));
      const partSuffix = chunks.length > 1 ? `_part${c + 1}` : '';
      const entryName = `${zipBase}_${folderName}${partSuffix}.xlsx`;
      const { blob } = await buildFormFKarnatakaWorkbookWithTemplateStyles({
        templateArrayBuffer:
          templateArrayBuffer instanceof ArrayBuffer
            ? templateArrayBuffer.slice(0)
            : templateArrayBuffer,
        mappedData: chunkRows.filter((r) => r && typeof r === 'object'),
        headersToUse,
        parsedFormHeader,
        formFileName: entryName,
        sheetNameHint,
      });
      const bytes = await blob.arrayBuffer();
      folder.file(entryName, bytes);
      fileCount += 1;
      if (fileCount % 8 === 0) {
        await new Promise((resolve) => setTimeout(resolve, 0));
      }
    }
  }

  return {
    blob: await zip.generateAsync({
      type: 'blob',
      mimeType: 'application/zip',
      compression: 'STORE',
    }),
    fileName: `${zipBase}_By_Location.zip`,
  };
}

