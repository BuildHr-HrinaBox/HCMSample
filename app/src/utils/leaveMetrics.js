/** Shared Earned Leave metrics parsing (Leave page + Statutory Form 15 Part 1). */

const EARNED_LEAVE_NAMES = new Set([
  'Earned Leave (Test)',
  'Earned Leave(Test)',
  'Earned Leave (test)',
  'Earned Leave',
  'Earned leave',
]);

export function parseLeaveCellObject(value) {
  if (value == null || value === '') return null;
  if (typeof value === 'object') return value;
  const raw = String(value).trim();
  if (!raw || raw === '{}') return null;
  if (!raw.startsWith('{') && !raw.startsWith('[')) return null;
  try {
    const parsed = JSON.parse(raw);
    return parsed && typeof parsed === 'object' ? parsed : null;
  } catch (_) {
    return null;
  }
}

export function isSuspiciousZohoRecordIdValue(value) {
  const s = String(value ?? '').trim();
  return /^\d{10,}$/.test(s);
}

export function sanitizeLeaveMetricDisplayValue(value) {
  if (value == null || value === '') return '';
  const s = String(value).trim();
  if (!s || isSuspiciousZohoRecordIdValue(s)) return '';
  return s;
}

function isEarnedLeaveLabelText(label) {
  const s = String(label || '').trim();
  if (!s || /legacy/i.test(s)) return false;
  if (EARNED_LEAVE_NAMES.has(s)) return true;
  return /^earned\s+leave(\s*\(test\))?$/i.test(s);
}

/** Column key for earned-leave breakout after backend rename, or raw leave type id. */
export function findEarnedLeaveKey(row, leaveTypeLabels = {}) {
  if (!row || typeof row !== 'object') return null;
  for (const key of Object.keys(row)) {
    if (isEarnedLeaveLabelText(key)) return key;
    const label = leaveTypeLabels && leaveTypeLabels[key];
    if (label && isEarnedLeaveLabelText(label)) return key;
  }
  return null;
}

export function getEarnedLeaveMetricsFromRecord(row, leaveTypeLabels = {}) {
  const earnedKey = findEarnedLeaveKey(row, leaveTypeLabels);
  if (!earnedKey) return { balance: '', availed: '', earnedKey: null };
  const raw = row[earnedKey];
  const obj = parseLeaveCellObject(raw) ?? (raw && typeof raw === 'object' ? raw : null);
  if (!obj || typeof obj !== 'object') return { balance: '', availed: '', earnedKey };
  const balanceRaw = obj.paidBalance ?? obj.balance ?? obj.Balance;
  const bookedRaw = obj.paidBooked ?? obj.booked ?? obj.Booked;
  return {
    balance: sanitizeLeaveMetricDisplayValue(balanceRaw),
    availed: sanitizeLeaveMetricDisplayValue(bookedRaw),
    earnedKey,
  };
}

/**
 * Leave-page earned metrics (paidBalance / paidBooked).
 * Tamil Nadu Form 15 Part 1 uses Form X formulas in form15TamilNaduLeave.js
 * (beginning = balance + booked, earnedDuring = 0, etc.).
 */
export function getForm15EarnedLeavePeriodMetrics(leaveRecord, leaveTypeLabels = {}) {
  const { balance, availed, earnedKey } = getEarnedLeaveMetricsFromRecord(leaveRecord, leaveTypeLabels);
  return {
    periodBalance: balance,
    periodBooked: availed !== '' ? availed : '0',
    earnedKey,
    hasData: balance !== '' || availed !== '',
  };
}

function extractParenIdentityToken(value) {
  const m = String(value ?? '').match(/\(([^)]+)\)\s*$/);
  return m ? m[1].toLowerCase().trim() : '';
}

/** Names + ids from a Zoho leave row (for Statutory employee matching). */
export function collectLeaveRecordIdentityKeys(leaveRecord) {
  const names = [];
  const ids = [];
  const addName = (v) => {
    if (v == null || !String(v).trim()) return;
    names.push(String(v).toLowerCase().trim());
  };
  const addId = (v) => {
    if (v == null || !String(v).trim()) return;
    const s = String(v).toLowerCase().trim();
    if (isSuspiciousZohoRecordIdValue(s)) return;
    ids.push(s);
  };
  const addNameWithParenId = (v) => {
    addName(v);
    const paren = extractParenIdentityToken(v);
    if (paren) addId(paren);
  };

  const emp = leaveRecord?.employee;
  if (emp && typeof emp === 'object') {
    addNameWithParenId(emp.name);
    addId(emp.id);
    addId(emp.employeeId);
  } else if (typeof emp === 'string') {
    addNameWithParenId(emp);
  }

  addNameWithParenId(
    leaveRecord?.EmployeeName ||
      leaveRecord?.employeeName ||
      leaveRecord?.['Employee Name'] ||
      leaveRecord?.name
  );
  addId(
    leaveRecord?.EmployeeID ||
      leaveRecord?.Employee_ID ||
      leaveRecord?.employee_id ||
      leaveRecord?.['Employee ID'] ||
      leaveRecord?.employeeId ||
      leaveRecord?.ZohoID ||
      leaveRecord?.['Zoho.ID']
  );

  return { names: [...new Set(names)], ids: [...new Set(ids)] };
}

/** Normalize leave API payloads (records map vs leaveRecords array). */
export function extractLeaveRecordsFromApiResult(leaveResult) {
  if (!leaveResult || typeof leaveResult !== 'object') {
    return { records: [], leaveTypeLabels: {} };
  }
  const leaveTypeLabels = leaveResult.leaveTypeLabels || {};
  if (Array.isArray(leaveResult.leaveRecords) && leaveResult.leaveRecords.length > 0) {
    return { records: leaveResult.leaveRecords, leaveTypeLabels };
  }
  if (
    leaveResult.records &&
    typeof leaveResult.records === 'object' &&
    !Array.isArray(leaveResult.records) &&
    Object.keys(leaveResult.records).length > 0
  ) {
    const records = Object.entries(leaveResult.records).map(([employeeId, row]) => ({
      employeeId: String(employeeId),
      ...(row && typeof row === 'object' ? row : {}),
    }));
    return { records, leaveTypeLabels };
  }
  return { records: [], leaveTypeLabels };
}

function addLookupKey(map, key, record) {
  const k = String(key || '').toLowerCase().trim();
  if (!k || k.length < 2 || isSuspiciousZohoRecordIdValue(k)) return;
  if (!map.has(k)) map.set(k, record);
}

/** Fast lookup of leave rows by worker code / display name. */
export function buildLeaveRecordLookupMap(leaveRecords) {
  const byId = new Map();
  const byName = new Map();
  (Array.isArray(leaveRecords) ? leaveRecords : []).forEach((record) => {
    if (!record || typeof record !== 'object') return;
    const { names, ids } = collectLeaveRecordIdentityKeys(record);
    ids.forEach((id) => addLookupKey(byId, id, record));
    names.forEach((name) => addLookupKey(byName, name, record));
  });
  return { byId, byName };
}

/** Match a statutory form row to a leave API record using Worker Identity No. / name. */
export function findLeaveRecordForFormRow(lookup, row, employeeIdHeader, employeeNameHeader) {
  if (!lookup || !row || typeof row !== 'object') return null;
  const rowId =
    employeeIdHeader && row[employeeIdHeader]
      ? String(row[employeeIdHeader]).toLowerCase().trim()
      : '';
  const rowName =
    employeeNameHeader && row[employeeNameHeader]
      ? String(row[employeeNameHeader]).toLowerCase().trim()
      : '';
  const rowNameParen = extractParenIdentityToken(rowName);
  const lookupId = String(row.__employeeLookupId || '').toLowerCase().trim();

  const idCandidates = [rowId, rowNameParen, lookupId].filter(Boolean);
  for (const id of idCandidates) {
    if (lookup.byId.has(id)) return lookup.byId.get(id);
  }

  for (const id of idCandidates) {
    for (const [name, record] of lookup.byName) {
      if (name.includes(`(${id})`) || name.endsWith(` ${id}`)) return record;
    }
  }

  if (rowName && lookup.byName.has(rowName)) return lookup.byName.get(rowName);

  const nameCandidates = [
    rowName,
    String(row.__employeeLookupName || '')
      .trim()
      .toLowerCase()
      .replace(/\([^)]*\)/g, ' ')
      .replace(/[^a-z0-9\s]/gi, ' ')
      .replace(/\s+/g, ' ')
      .trim(),
  ].filter(Boolean);

  for (const name of nameCandidates) {
    if (lookup.byName.has(name)) return lookup.byName.get(name);
    const nameParts = name.split(' ');
    for (const [leaveName, record] of lookup.byName) {
      if (leaveName === name) return record;
      const leaveParts = leaveName.split(' ');
      if (
        nameParts.length === 1 &&
        leaveParts.length >= 2 &&
        nameParts[0].length >= 4 &&
        nameParts[0] === leaveParts[0]
      ) {
        return record;
      }
      if (
        leaveParts.length === 1 &&
        nameParts.length >= 2 &&
        leaveParts[0].length >= 4 &&
        leaveParts[0] === nameParts[0]
      ) {
        return record;
      }
    }
  }

  return null;
}

/** Sum booked (availed) and balance across all leave types on one Zoho leave row. */
export function sumLeaveRecordBookedAndBalance(leaveRecord, leaveTypeLabels = {}) {
  if (!leaveRecord || typeof leaveRecord !== 'object') {
    return { booked: '', balance: '', categoryLabels: [] };
  }
  let bookedTotal = 0;
  let balanceTotal = 0;
  let hasBooked = false;
  let hasBalance = false;
  const categoryLabels = [];
  const skipKeys = new Set([
    's.no',
    'sno',
    'employeeid',
    'employee_id',
    'employee',
    'employeename',
    'employeename',
    'zoho.id',
    'zohoid',
    'zoho_id',
  ]);

  Object.entries(leaveRecord).forEach(([key, value]) => {
    const kl = String(key || '')
      .toLowerCase()
      .replace(/\s+/g, ' ')
      .trim();
    if (skipKeys.has(kl.replace(/\s+/g, '')) || skipKeys.has(kl)) return;
    const label = String(leaveTypeLabels?.[key] || key || '').trim();
    const obj = parseLeaveCellObject(value) ?? (value && typeof value === 'object' ? value : null);
    if (!obj || typeof obj !== 'object') return;
    const bookRaw = obj.paidBooked ?? obj.booked ?? obj.Booked;
    const balRaw = obj.paidBalance ?? obj.balance ?? obj.Balance;
    const bookNum = Number(String(bookRaw ?? '').trim());
    const balNum = Number(String(balRaw ?? '').trim());
    if (Number.isFinite(bookNum) && bookNum !== 0) {
      bookedTotal += bookNum;
      hasBooked = true;
    }
    if (Number.isFinite(balNum)) {
      balanceTotal += balNum;
      hasBalance = true;
    }
    if (label && !/legacy/i.test(label)) categoryLabels.push(label);
  });

  return {
    booked: hasBooked ? String(bookedTotal) : '',
    balance: hasBalance ? String(balanceTotal) : '',
    categoryLabels: [...new Set(categoryLabels)],
  };
}
