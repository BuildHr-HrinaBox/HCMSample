/** Karnataka Form F — Register of Leave with Wages (Rule 8): header fields + PART I earned leave table. */

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
      const cells = readRow(newStartIndex).filter(Boolean);
      if (cells.length === 0) {
        newStartIndex += 1;
        continue;
      }
      if (isPartIEarnedLeaveBanner(cells.join(' ')) || /part\s*[-–]?\s*ii/i.test(cells.join(' ').toLowerCase())) {
        break;
      }
      if (rowLooksLikePartISubHeaderRow(cells) || rowLooksLikePartINumberedIndexRow(cells)) {
        newStartIndex += 1;
        continue;
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

/** Leave at credit = Leave earned during the Period − approved LeaveCount (0 when no leave availed). */
export function computeFormFKarnatakaLeaveAtCredit(earnedDuring, leaveCount) {
  const earned = Number(String(earnedDuring ?? '').replace(/,/g, '').trim());
  if (!Number.isFinite(earned)) return '';
  const countRaw = leaveCount == null || leaveCount === '' ? '0' : leaveCount;
  const count = Number(String(countRaw).replace(/,/g, '').trim());
  const availed = Number.isFinite(count) ? count : 0;
  const credit = Math.round((earned - availed) * 100) / 100;
  return String(credit);
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

  let earnedDuring = String(row[canon[3]] ?? '').trim();
  if (leaveRecord) {
    const { periodBalance } = getForm15EarnedLeavePeriodMetrics(leaveRecord, leaveTypeLabels);
    if (periodBalance !== '') {
      earnedDuring = periodBalance;
      setByIndex(3, earnedDuring);
    }
  }

  if (earnedDuring) {
    const leaveCount = approvedRecord
      ? parseApprovedLeaveDaysCount(approvedRecord.Days ?? approvedRecord.days)
      : '';
    const credit = computeFormFKarnatakaLeaveAtCredit(earnedDuring, leaveCount !== '' ? leaveCount : '0');
    if (credit !== '') setByIndex(4, credit);
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
  const approvedLeaveRecords = Array.isArray(options.approvedLeaveRecords)
    ? options.approvedLeaveRecords
    : [];
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
    if (approvedLookup) {
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
    if (Number.isFinite(n)) {
      total += n;
      hasCount = true;
    }
  });
  return hasCount ? String(total) : '';
}

function isZohoLeaveDateString(value) {
  return /^\d{1,2}-[A-Za-z]{3}-\d{4}$/i.test(String(value || '').trim());
}

/** Earliest / latest date keys from approved-leave `Days` JSON (e.g. "18-May-2026"). */
function extractLeaveDatesFromDaysField(daysField) {
  const obj = parseJsonMaybe(daysField);
  if (!obj || typeof obj !== 'object' || Array.isArray(obj)) {
    return { from: '', to: '' };
  }
  const dateKeys = Object.keys(obj)
    .map((k) => String(k || '').trim())
    .filter((k) => isZohoLeaveDateString(k));
  if (dateKeys.length === 0) return { from: '', to: '' };

  const toTime = (key) => {
    const m = key.match(/^(\d{1,2})-([A-Za-z]{3})-(\d{4})$/i);
    if (!m) return 0;
    const months = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];
    const mon = months.indexOf(m[2].toLowerCase());
    if (mon < 0) return 0;
    return new Date(parseInt(m[3], 10), mon, parseInt(m[1], 10)).getTime();
  };

  dateKeys.sort((a, b) => toTime(a) - toTime(b));
  return { from: dateKeys[0], to: dateKeys[dateKeys.length - 1] };
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

export function normalizePersonNameKey(name) {
  return String(name || '')
    .replace(/\([^)]*\)/g, ' ')
    .replace(/[^a-z0-9\s]/gi, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
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
  ['Employee.ID', 'EmployeeID', 'employeeId', 'Employee Id', 'erecno', 'Erecno'].forEach((k) => {
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

  const addName = (value) => {
    const n = normalizePersonNameKey(value);
    if (n) nameCandidates.add(n);
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
        emp.employeeId
    );
    const first = String(emp.FirstName || emp['FirstName'] || emp.firstName || '').trim();
    const last = String(emp.LastName || emp['LastName'] || emp.lastName || '').trim();
    if (first && last) addName(`${first} ${last}`);
    if (first) addName(first);
  }

  return { nameCandidates, idCandidates };
}

function approvedLeaveRecordMatchesEmployee(record, nameCandidates, idCandidates) {
  const { names, ids } = collectApprovedLeaveIdentityKeys(record);
  if (ids.some((id) => idCandidates.has(id))) return true;
  for (const recordName of names) {
    for (const cand of nameCandidates) {
      if (personNamesMatch(cand, recordName)) return true;
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

  const pickFrom = (records) => {
    const monthFrom = options.monthFrom || '';
    const monthTo = options.monthTo || '';
    for (const record of records) {
      if (!record || !isAvailable(record)) continue;
      if (monthFrom && monthTo && !approvedLeaveRecordIsAvailedInMonth(record, monthFrom, monthTo)) {
        continue;
      }
      if (approvedLeaveRecordMatchesEmployee(record, nameCandidates, idCandidates)) {
        return record;
      }
    }
    return null;
  };

    // Exact map lookup by id / normalized name (no substring matching).
  if (lookup && typeof lookup.get === 'function') {
    const monthFrom = options.monthFrom || '';
    const monthTo = options.monthTo || '';
    const lookupHitValid = (hit) =>
      hit &&
      isAvailable(hit) &&
      (!monthFrom || !monthTo || approvedLeaveRecordIsAvailedInMonth(hit, monthFrom, monthTo)) &&
      approvedLeaveRecordMatchesEmployee(hit, nameCandidates, idCandidates);

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
    setByIndex(4, computeFormFKarnatakaLeaveAtCredit(earnedDuring, metrics.daysCount || '0'));
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

  const lookup = buildApprovedLeaveLookupMap(approvedLeaveRecords);
  const { employeeNameHeader, employeeIdHeader } = options.resolveEmployeeHeaders
    ? options.resolveEmployeeHeaders(canonicalHeaders)
    : { employeeNameHeader: '', employeeIdHeader: '' };
  const monthFrom = options.monthFrom || '';
  const monthTo = options.monthTo || '';

  const usedRecordKeys = new Set();
  let hits = 0;
  mappedData.forEach((row, index) => {
    const normalizedRow = remapFormFKarnatakaRowToCanonicalHeaders(row, priorHeaders);
    Object.assign(row, normalizedRow);

    const empItem = employeesForMapping[index];
    const emp = empItem && (empItem.Employee || empItem.employee || empItem);
    const approvedRecord = findApprovedLeaveForEmployee(
      lookup,
      emp,
      row,
      employeeIdHeader,
      employeeNameHeader,
      {
        approvedRecords: approvedLeaveRecords,
        usedRecordKeys,
        monthFrom,
        monthTo,
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
