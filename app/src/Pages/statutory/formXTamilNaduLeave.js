/**
 * Tamil Nadu Form X — Register of Leave and Social Security Benefits.
 *
 * Earned Leave  → Zoho Earned Leave balance + approved LeaveCount
 * Medical Leave → Contingency Leave
 * Other Leave   → Legacy Earned Leave
 *
 * Beginning of month = Leave API balance + approved LeaveCount for the month
 *   (so April balance 23 + 1 LeaveCount → beginning 24; May with 0 LeaveCount → 23).
 * Availed during month = sum of LeaveCount from approved leaves of that type.
 */

import {
  buildLeaveRecordLookupMap,
  findLeaveRecordForFormRow,
  getEarnedLeaveMetricsFromRecord,
  parseLeaveCellObject,
  sanitizeLeaveMetricDisplayValue,
} from '../../utils/leaveMetrics';
import {
  collectApprovedLeaveIdentityKeys,
  filterApprovedLeaveRecordsForMonth,
  normalizeApprovedLeaveRecord,
  parseApprovedLeaveDaysCount,
  personNamesMatch,
} from './formFKarnataka';
import { readApprovedLeaveLeaveType } from './formOGJGujarat';

export const FORM_X_EARNED_LEAVE_TYPE_ALIASES = [
  'earned leave',
  'earned leave (test)',
  'earned leave(test)',
  'el',
];

export const FORM_X_MEDICAL_LEAVE_TYPE_ALIASES = [
  'contingency leave',
  'contingency',
  'cl',
];

export const FORM_X_OTHER_LEAVE_TYPE_ALIASES = [
  'legacy earned leave',
  'legacy earned',
  'legacy',
];

function normalizeLeaveTypeLabel(label) {
  return String(label || '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

export function leaveTypeLabelMatchesAliases(label, aliases) {
  const normalized = normalizeLeaveTypeLabel(label);
  if (!normalized) return false;
  const normalizedAliases = (aliases || []).map(normalizeLeaveTypeLabel).filter(Boolean);
  const isLegacyEarnedLeave = normalized.includes('legacy') && normalized.includes('earned leave');
  const aliasesTargetLegacy = normalizedAliases.some((alias) => alias.includes('legacy'));
  if (isLegacyEarnedLeave && !aliasesTargetLegacy) return false;
  return normalizedAliases.some((a) => {
    if (!a) return false;
    if (normalized === a) return true;
    // Label contains alias as a phrase (e.g. "contingency leave" ⊇ "contingency").
    if (a.length >= 3 && (normalized === a || normalized.startsWith(`${a} `) || normalized.endsWith(` ${a}`) || normalized.includes(` ${a} `))) {
      return true;
    }
    if (normalized.includes(a) && a.length >= 4) {
      // Reject "earned leave" matching inside a more specific alias check done below.
      return true;
    }
    // Alias contains label only when alias is not more specific (e.g. do not let
    // "legacy earned leave" match label "earned leave").
    if (a.includes(normalized) && normalized.length >= 4) {
      const extra = a.replace(normalized, ' ').replace(/\s+/g, ' ').trim();
      return !extra;
    }
    return false;
  });
}

function toFiniteNumber(value) {
  const n = Number(String(value ?? '').replace(/,/g, '').trim());
  return Number.isFinite(n) ? n : null;
}

function formatLeaveNumber(value) {
  const n = toFiniteNumber(value);
  if (n == null) return '';
  return String(Math.round(n * 100) / 100);
}

/**
 * Reconstruct beginning-of-month from Leave API end balance + approved LeaveCount.
 * Example: balance 23 + LeaveCount 1 → beginning 24; next month balance 23 + 0 → 23.
 */
export function computeFormXLeaveBeginning(apiBalance, approvedLeaveCount) {
  const balance = toFiniteNumber(apiBalance);
  if (balance == null) return '';
  const availed = toFiniteNumber(approvedLeaveCount);
  const add = availed == null ? 0 : availed;
  return formatLeaveNumber(balance + add);
}

export function computeFormXLeaveEndBalance(beginning, availed) {
  const start = toFiniteNumber(beginning);
  if (start == null) return '';
  const used = toFiniteNumber(availed);
  const sub = used == null ? 0 : used;
  return formatLeaveNumber(start - sub);
}

function getLeaveTypeCell(leaveRecord, aliases, leaveTypeLabels = {}) {
  if (!leaveRecord || typeof leaveRecord !== 'object') return null;
  const skipKeys = new Set(['employee', 'employeeid', 'totals', 's.no', 'sno']);
  for (const key of Object.keys(leaveRecord)) {
    const keyNorm = normalizeLeaveTypeLabel(key);
    if (skipKeys.has(keyNorm)) continue;
    const labelsToCheck = [key, leaveTypeLabels[key]].filter(Boolean);
    if (labelsToCheck.some((label) => leaveTypeLabelMatchesAliases(label, aliases))) {
      return leaveRecord[key];
    }
  }
  return null;
}

function leaveMetricsFromTypeCell(raw) {
  const parsed = parseLeaveCellObject(raw);
  if (!parsed || typeof parsed !== 'object') {
    // Plain numeric cell (e.g. Legacy Earned Leave = 0)
    const asNum = toFiniteNumber(raw);
    if (asNum != null) {
      return { balance: formatLeaveNumber(asNum), booked: '', hasData: true };
    }
    return { balance: '', booked: '', hasData: false };
  }
  const balanceRaw = parsed.paidBalance ?? parsed.balance ?? parsed.Balance;
  const bookedRaw = parsed.paidBooked ?? parsed.booked ?? parsed.Booked;
  return {
    balance: sanitizeLeaveMetricDisplayValue(balanceRaw),
    booked: sanitizeLeaveMetricDisplayValue(bookedRaw),
    hasData:
      sanitizeLeaveMetricDisplayValue(balanceRaw) !== '' ||
      sanitizeLeaveMetricDisplayValue(bookedRaw) !== '',
  };
}

export function getFormXEarnedLeaveApiMetrics(leaveRecord, leaveTypeLabels = {}) {
  const primary = getEarnedLeaveMetricsFromRecord(leaveRecord, leaveTypeLabels);
  if (primary.balance !== '' || primary.availed !== '') {
    return {
      balance: primary.balance,
      booked: primary.availed,
      hasData: true,
    };
  }
  return leaveMetricsFromTypeCell(
    getLeaveTypeCell(leaveRecord, FORM_X_EARNED_LEAVE_TYPE_ALIASES, leaveTypeLabels)
  );
}

export function getFormXMedicalLeaveApiMetrics(leaveRecord, leaveTypeLabels = {}) {
  return leaveMetricsFromTypeCell(
    getLeaveTypeCell(leaveRecord, FORM_X_MEDICAL_LEAVE_TYPE_ALIASES, leaveTypeLabels)
  );
}

export function getFormXOtherLeaveApiMetrics(leaveRecord, leaveTypeLabels = {}) {
  return leaveMetricsFromTypeCell(
    getLeaveTypeCell(leaveRecord, FORM_X_OTHER_LEAVE_TYPE_ALIASES, leaveTypeLabels)
  );
}

export function approvedLeaveMatchesTypeAliases(record, aliases) {
  const leaveType = readApprovedLeaveLeaveType(record);
  if (!leaveType) return false;
  return leaveTypeLabelMatchesAliases(leaveType, aliases);
}

/**
 * Sum LeaveCount from approved leave records for one employee + leave-type aliases.
 * When aliases are for Earned Leave and a record has no Leave Type, count it as earned
 * (Zoho often omits type for standard EL).
 */
export function sumApprovedLeaveCountForEmployee(
  approvedRecords,
  emp,
  row,
  employeeIdHeader,
  employeeNameHeader,
  aliases,
  options = {}
) {
  const records = Array.isArray(approvedRecords) ? approvedRecords : [];
  if (records.length === 0) return '';

  const monthFrom = options.monthFrom || '';
  const monthTo = options.monthTo || '';
  const filtered =
    monthFrom && monthTo
      ? filterApprovedLeaveRecordsForMonth(records, monthFrom, monthTo)
      : records;

  const nameCandidates = new Set();
  const idCandidates = new Set();
  const addName = (v) => {
    const s = String(v || '')
      .replace(/\([^)]*\)/g, ' ')
      .replace(/[^a-z0-9\s]/gi, ' ')
      .replace(/\s+/g, ' ')
      .trim()
      .toLowerCase();
    if (s) nameCandidates.add(s);
  };
  const addId = (v) => {
    const s = String(v || '')
      .trim()
      .toLowerCase();
    if (s) idCandidates.add(s);
  };

  if (employeeNameHeader && row?.[employeeNameHeader]) addName(row[employeeNameHeader]);
  if (employeeIdHeader && row?.[employeeIdHeader]) addId(row[employeeIdHeader]);
  if (options.collectEmployeeNameCandidates) {
    (options.collectEmployeeNameCandidates(emp || {}) || []).forEach(addName);
  }
  if (options.collectEmployeeIdCandidates) {
    (options.collectEmployeeIdCandidates(emp || {}, row || {}) || []).forEach(addId);
  } else if (emp && typeof emp === 'object') {
    addId(emp.Zoho_ID || emp.ZohoID || emp.EmployeeID || emp.Employee_ID);
    const first = String(emp.FirstName || '').trim();
    const last = String(emp.LastName || '').trim();
    if (first && last) addName(`${first} ${last}`);
    if (first) addName(first);
  }

  if (nameCandidates.size === 0 && idCandidates.size === 0) return '';

  const isEarnedAliases =
    Array.isArray(aliases) &&
    aliases.some((a) => normalizeLeaveTypeLabel(a).includes('earned')) &&
    !aliases.some((a) => normalizeLeaveTypeLabel(a).includes('legacy'));

  let total = 0;
  let hasCount = false;

  filtered.forEach((record) => {
    if (!record || typeof record !== 'object') return;
    const { names, ids } = collectApprovedLeaveIdentityKeys(record);
    const idHit = ids.some((id) => idCandidates.has(id));
    const nameHit = names.some((rn) =>
      [...nameCandidates].some((cand) => personNamesMatch(cand, rn))
    );
    if (!idHit && !nameHit) return;

    const leaveType = readApprovedLeaveLeaveType(record);
    if (leaveType) {
      if (!leaveTypeLabelMatchesAliases(leaveType, aliases)) return;
    } else if (!isEarnedAliases) {
      // Untyped approvals only count toward Earned Leave.
      return;
    }

    const days = parseApprovedLeaveDaysCount(record.Days ?? record.days);
    const n = toFiniteNumber(days);
    if (n == null) return;
    total += n;
    hasCount = true;
  });

  return hasCount ? formatLeaveNumber(total) : '';
}

export function buildFormXLeaveSectionValues(leaveRecord, approvedAvailed, leaveTypeLabels, section) {
  let apiMetrics = { balance: '', booked: '', hasData: false };
  if (section === 'earned') {
    apiMetrics = getFormXEarnedLeaveApiMetrics(leaveRecord, leaveTypeLabels);
  } else if (section === 'medical') {
    apiMetrics = getFormXMedicalLeaveApiMetrics(leaveRecord, leaveTypeLabels);
  } else if (section === 'other') {
    apiMetrics = getFormXOtherLeaveApiMetrics(leaveRecord, leaveTypeLabels);
  }

  const availed =
    approvedAvailed !== '' && approvedAvailed != null
      ? formatLeaveNumber(approvedAvailed)
      : '0';
  const beginning = computeFormXLeaveBeginning(apiMetrics.balance, availed);
  const balance =
    apiMetrics.balance !== ''
      ? formatLeaveNumber(apiMetrics.balance)
      : computeFormXLeaveEndBalance(beginning, availed);
  const earnedDuring = apiMetrics.balance !== '' ? formatLeaveNumber(apiMetrics.balance) : '';

  return {
    beginning,
    earnedDuring,
    availed,
    balance,
    hasApiData: apiMetrics.hasData,
  };
}

function setRowCell(row, header, value, overwrite = true) {
  if (!header || value == null || value === '') return false;
  const cur = String(row[header] ?? '').trim();
  if (!overwrite && cur) return false;
  row[header] = String(value);
  return true;
}

/**
 * Apply Form X leave columns for one employee row.
 * @returns {number} cells written
 */
export function applyFormXTamilNaduLeaveToRow(
  row,
  leaveRecord,
  sectionHeaders,
  leaveTypeLabels,
  approvedLeaveRecords,
  options = {}
) {
  if (!row || !sectionHeaders) return 0;
  const {
    emp = null,
    employeeIdHeader = '',
    employeeNameHeader = '',
    overwrite = true,
    monthFrom = '',
    monthTo = '',
    collectEmployeeNameCandidates,
    collectEmployeeIdCandidates,
  } = options;

  const sumFor = (aliases) =>
    sumApprovedLeaveCountForEmployee(
      approvedLeaveRecords,
      emp,
      row,
      employeeIdHeader,
      employeeNameHeader,
      aliases,
      {
        monthFrom,
        monthTo,
        collectEmployeeNameCandidates,
        collectEmployeeIdCandidates,
      }
    );

  let applied = 0;
  const sections = [
    {
      key: 'earned',
      aliases: FORM_X_EARNED_LEAVE_TYPE_ALIASES,
      headers: {
        beginning: sectionHeaders.earnedBeginning,
        earnedDuring: sectionHeaders.earnedEarned,
        availed: sectionHeaders.earnedAvailed,
        balance: sectionHeaders.earnedBalance,
      },
    },
    {
      key: 'medical',
      aliases: FORM_X_MEDICAL_LEAVE_TYPE_ALIASES,
      headers: {
        beginning: sectionHeaders.medicalBeginning,
        availed: sectionHeaders.medicalAvailed,
        balance: sectionHeaders.medicalBalance,
      },
    },
    {
      key: 'other',
      aliases: FORM_X_OTHER_LEAVE_TYPE_ALIASES,
      headers: {
        beginning: sectionHeaders.otherBeginning,
        availed: sectionHeaders.otherAvailed,
        balance: sectionHeaders.otherBalance,
      },
    },
  ];

  sections.forEach((section) => {
    const approvedAvailed = sumFor(section.aliases);
    const values = buildFormXLeaveSectionValues(
      leaveRecord,
      approvedAvailed,
      leaveTypeLabels,
      section.key
    );
    // Always write availed (0 when none) when the column exists and we have API data or approvals.
    const hasAnything =
      values.hasApiData ||
      (approvedAvailed !== '' && approvedAvailed != null) ||
      leaveRecord;
    if (!hasAnything && !leaveRecord) return;

    if (setRowCell(row, section.headers.beginning, values.beginning, overwrite)) applied += 1;
    if (
      section.headers.earnedDuring &&
      setRowCell(row, section.headers.earnedDuring, values.earnedDuring, overwrite)
    ) {
      applied += 1;
    }
    if (setRowCell(row, section.headers.availed, values.availed, overwrite)) applied += 1;
    if (setRowCell(row, section.headers.balance, values.balance, overwrite)) applied += 1;
  });

  return applied;
}

export function applyFormXTamilNaduLeaveAutofill(
  mappedData,
  employeesForMapping,
  leaveRecords,
  leaveTypeLabels,
  tableHeaders,
  sectionHeaders,
  approvedLeaveRecords,
  options = {}
) {
  if (!Array.isArray(mappedData) || mappedData.length === 0) return 0;
  if (!sectionHeaders) return 0;

  const leaveLookup = buildLeaveRecordLookupMap(leaveRecords);
  const { employeeNameHeader = '', employeeIdHeader = '' } = options.resolveEmployeeHeaders
    ? options.resolveEmployeeHeaders(tableHeaders)
    : {};

  let hits = 0;
  mappedData.forEach((row, index) => {
    const empItem = employeesForMapping?.[index];
    const emp = empItem && (empItem.Employee || empItem.employee || empItem);
    const leaveRecord =
      findLeaveRecordForFormRow(leaveLookup, row, employeeIdHeader, employeeNameHeader) || null;
    const applied = applyFormXTamilNaduLeaveToRow(
      row,
      leaveRecord,
      sectionHeaders,
      leaveTypeLabels || {},
      approvedLeaveRecords || [],
      {
        emp,
        employeeIdHeader,
        employeeNameHeader,
        overwrite: options.overwrite !== false,
        monthFrom: options.monthFrom || options.fromDate || '',
        monthTo: options.monthTo || options.toDate || '',
        collectEmployeeNameCandidates: options.collectEmployeeNameCandidates,
        collectEmployeeIdCandidates: options.collectEmployeeIdCandidates,
      }
    );
    if (applied > 0) hits += 1;
  });
  return hits;
}

export { normalizeApprovedLeaveRecord, parseApprovedLeaveDaysCount };
