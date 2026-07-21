/**
 * Tamil Nadu Form 15 Part I — Register of Leave with Wages.
 *
 * Leave columns use the same Zoho Leave Fetch mapping as Tamil Nadu Form X
 * (Register of Leave and Social Security Benefits):
 *
 * Earned Leave:
 *   Leave at the beginning of the Month = Leave earned during the Period + Leave availed during the Period
 *   Leave earned during the Period      = 0
 *   Leave availed during the Month      = Leave availed during the Period (paidBooked)
 *   Leave balance at the end of the Month = Leave earned during the Period (paidBalance)
 *
 * Medical Leave (Contingency Leave):
 *   Leave at beginning of the Month = Balance + Booked
 *   Leave availed during the Month  = Booked
 *   Leave balance at end of the Month = Balance
 *
 * Other Leave (Legacy Earned Leave), when the template includes that band:
 *   Beginning / availed / end balance from Leave API + approved LeaveCount (Form X rules).
 */

import {
  buildLeaveRecordLookupMap,
  collectLeaveRecordIdentityKeys,
  findLeaveRecordForFormRow,
} from '../../utils/leaveMetrics';
import {
  applyFormXTamilNaduLeaveAutofill,
  applyFormXTamilNaduLeaveToRow,
  buildFormXLeaveSectionValues,
} from './formXTamilNaduLeave';

/** Form X Medical Leave leaf titles (same relative layout). */
export const FORM_15_PART1_MEDICAL_LEAVE_LEAVES = [
  'Leave at beginning of the Month',
  'Leave availed during the Month',
  'Leave balance at end of the Month',
];

/** Form X Other Leave leaf titles (same relative layout). */
export const FORM_15_PART1_OTHER_LEAVE_LEAVES = [
  'Leave at beginning of the Month',
  'Leave availed during the Month',
  'Leave Balance at end of the Month',
];

function normHeader(v) {
  return String(v || '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function dedupeLeafHeaders(headers) {
  const seen = new Map();
  return (Array.isArray(headers) ? headers : []).map((h) => {
    const base = String(h || '').trim();
    if (!base) return '';
    const key = base.toLowerCase();
    const n = seen.get(key) || 0;
    seen.set(key, n + 1);
    if (n === 0) return base;
    return `${base} (${n + 1})`;
  });
}

/** Adapt Form 15 Part 1 header keys → Form X sectionHeaders shape. */
export function toFormXSectionHeadersFromForm15(form15Headers) {
  if (!form15Headers || typeof form15Headers !== 'object') return null;
  return {
    earnedBeginning: form15Headers.earnedBeginning || null,
    earnedEarned: form15Headers.earnedDuring || form15Headers.earnedEarned || null,
    earnedAvailed: form15Headers.earnedAvailed || null,
    earnedBalance: form15Headers.earnedBalance || null,
    medicalBeginning: form15Headers.medicalBeginning || null,
    medicalAvailed: form15Headers.medicalAvailed || null,
    medicalBalance: form15Headers.medicalBalance || null,
    otherBeginning: form15Headers.otherBeginning || null,
    otherAvailed: form15Headers.otherAvailed || null,
    otherBalance: form15Headers.otherBalance || null,
  };
}

/**
 * Resolve Form 15 Part 1 leave leaf columns under merged Earned / Medical / Other groups.
 * Always falls back to leaf column order (same as Form X) when group labels are blank.
 */
export function resolveForm15Part1LeaveColumnHeaders(headers, groupLabels) {
  const headersArr = Array.isArray(headers) ? headers : [];
  const groups =
    Array.isArray(groupLabels) && groupLabels.length === headersArr.length ? groupLabels : null;
  const result = {
    earnedBeginning: null,
    earnedDuring: null,
    earnedAvailed: null,
    earnedBalance: null,
    medicalBeginning: null,
    medicalAvailed: null,
    medicalBalance: null,
    otherBeginning: null,
    otherAvailed: null,
    otherBalance: null,
  };
  const lists = { beginning: [], availed: [], balance: [], earned: [] };

  const metricType = (header) => {
    const s = normHeader(header);
    if (!s.includes('leave')) return null;
    if (s.includes('beginning') && s.includes('month')) return 'beginning';
    if (s.includes('earned') && (s.includes('period') || s.includes('during'))) return 'earned';
    if (s.includes('availed') && s.includes('month')) return 'availed';
    if (s.includes('balance') && s.includes('end') && s.includes('month')) return 'balance';
    return null;
  };

  headersArr.forEach((header, index) => {
    const metric = metricType(header);
    if (!metric) return;
    if (metric === 'earned') lists.earned.push(header);
    if (metric === 'beginning') lists.beginning.push(header);
    if (metric === 'availed') lists.availed.push(header);
    if (metric === 'balance') lists.balance.push(header);

    if (!groups) return;
    const group = normHeader(groups[index]);
    if (group.includes('earned leave')) {
      if (metric === 'beginning' && !result.earnedBeginning) result.earnedBeginning = header;
      if (metric === 'earned' && !result.earnedDuring) result.earnedDuring = header;
      if (metric === 'availed' && !result.earnedAvailed) result.earnedAvailed = header;
      if (metric === 'balance' && !result.earnedBalance) result.earnedBalance = header;
    }
    if (group.includes('medical leave') || (group.includes('medical') && group.includes('leave'))) {
      if (metric === 'beginning' && !result.medicalBeginning) result.medicalBeginning = header;
      if (metric === 'availed' && !result.medicalAvailed) result.medicalAvailed = header;
      if (metric === 'balance' && !result.medicalBalance) result.medicalBalance = header;
    }
    if (group.includes('other leave')) {
      if (metric === 'beginning' && !result.otherBeginning) result.otherBeginning = header;
      if (metric === 'availed' && !result.otherAvailed) result.otherAvailed = header;
      if (metric === 'balance' && !result.otherBalance) result.otherBalance = header;
    }
  });

  if (!result.earnedBeginning && lists.beginning.length > 0) result.earnedBeginning = lists.beginning[0];
  if (!result.earnedDuring && lists.earned.length > 0) result.earnedDuring = lists.earned[0];
  if (!result.earnedAvailed && lists.availed.length > 0) result.earnedAvailed = lists.availed[0];
  if (!result.earnedBalance && lists.balance.length > 0) result.earnedBalance = lists.balance[0];
  if (!result.medicalBeginning && lists.beginning.length >= 2) {
    result.medicalBeginning = lists.beginning[1];
  }
  if (!result.medicalAvailed && lists.availed.length >= 2) result.medicalAvailed = lists.availed[1];
  if (!result.medicalBalance && lists.balance.length >= 2) result.medicalBalance = lists.balance[1];
  if (!result.otherBeginning && lists.beginning.length >= 3) result.otherBeginning = lists.beginning[2];
  if (!result.otherAvailed && lists.availed.length >= 3) result.otherAvailed = lists.availed[2];
  if (!result.otherBalance && lists.balance.length >= 3) result.otherBalance = lists.balance[2];

  return result;
}

function hasMedicalLeaveBand(resolved) {
  return !!(resolved?.medicalBeginning || resolved?.medicalAvailed || resolved?.medicalBalance);
}

function hasOtherLeaveBand(resolved) {
  return !!(resolved?.otherBeginning || resolved?.otherAvailed || resolved?.otherBalance);
}

/** Index after the last Earned Leave leaf (or last identity/leave column) for inserts. */
function findForm15Part1LeaveInsertIndex(headers, groupLabels, resolved) {
  const headersArr = Array.isArray(headers) ? headers : [];
  const groups =
    Array.isArray(groupLabels) && groupLabels.length === headersArr.length ? groupLabels : null;

  const earnedKeys = [
    resolved?.earnedBeginning,
    resolved?.earnedDuring,
    resolved?.earnedAvailed,
    resolved?.earnedBalance,
  ].filter(Boolean);

  let lastEarnedIdx = -1;
  earnedKeys.forEach((key) => {
    const idx = headersArr.indexOf(key);
    if (idx > lastEarnedIdx) lastEarnedIdx = idx;
  });
  if (lastEarnedIdx >= 0) return lastEarnedIdx + 1;

  if (groups) {
    for (let i = headersArr.length - 1; i >= 0; i -= 1) {
      if (normHeader(groups[i]).includes('earned leave')) return i + 1;
    }
  }

  // After Worker Identity / Name columns when Earned Leave leaves are missing.
  for (let i = 0; i < headersArr.length; i += 1) {
    const h = normHeader(headersArr[i]);
    if (h.includes('identity') || h.includes('identification') || (h.includes('employee') && h.includes('id'))) {
      return i + 1;
    }
  }
  return headersArr.length;
}

/**
 * Ensure Form 15 Part 1 has Medical Leave + Other Leave bands in Form X order/layout.
 * Inserts missing leaf columns after Earned Leave; preserves existing cells and labels.
 *
 * @returns {{ headers, groupLabels, rows, insertedMedical, insertedOther, changed }}
 */
export function ensureForm15Part1LeaveColumnsLikeFormX(headers, groupLabels, rows = []) {
  const headersArr = Array.isArray(headers) ? [...headers] : [];
  let groupsArr =
    Array.isArray(groupLabels) && groupLabels.length === headersArr.length
      ? [...groupLabels]
      : headersArr.map(() => '');

  const resolved = resolveForm15Part1LeaveColumnHeaders(headersArr, groupsArr);
  const needMedical = !hasMedicalLeaveBand(resolved);
  const needOther = !hasOtherLeaveBand(resolved);

  // Always dedupe repeated leaf titles (Earned/Medical/Other share "Leave availed…")
  // even when bands already exist — otherwise JS row keys collide and leave cells stay blank.
  if (!needMedical && !needOther) {
    const dedupedOnly = dedupeLeafHeaders(headersArr);
    const dedupeChanged = dedupedOnly.some((h, i) => h !== headersArr[i]);
    if (!dedupeChanged) {
      return {
        headers: headersArr,
        groupLabels: groupsArr,
        rows: Array.isArray(rows) ? rows : [],
        insertedMedical: false,
        insertedOther: false,
        changed: false,
      };
    }
    const nextRows = (Array.isArray(rows) ? rows : []).map((row) => {
      if (!row || typeof row !== 'object') return row;
      const next = { ...row };
      headersArr.forEach((oldH, i) => {
        const newH = dedupedOnly[i];
        if (!oldH || !newH || oldH === newH) return;
        if (next[newH] == null || next[newH] === '') {
          next[newH] = next[oldH] ?? '';
        }
      });
      dedupedOnly.forEach((h) => {
        if (h && next[h] == null) next[h] = '';
      });
      return next;
    });
    return {
      headers: dedupedOnly,
      groupLabels: groupsArr,
      rows: nextRows,
      insertedMedical: false,
      insertedOther: false,
      changed: true,
    };
  }

  let insertAt = findForm15Part1LeaveInsertIndex(headersArr, groupsArr, resolved);
  const leavesToInsert = [];
  const groupsToInsert = [];

  if (needMedical) {
    FORM_15_PART1_MEDICAL_LEAVE_LEAVES.forEach((leaf) => {
      leavesToInsert.push(leaf);
      groupsToInsert.push('Medical Leave');
    });
  }
  if (needOther) {
    FORM_15_PART1_OTHER_LEAVE_LEAVES.forEach((leaf) => {
      leavesToInsert.push(leaf);
      groupsToInsert.push('Other Leave');
    });
  }

  headersArr.splice(insertAt, 0, ...leavesToInsert);
  groupsArr.splice(insertAt, 0, ...groupsToInsert);

  const dedupedHeaders = dedupeLeafHeaders(headersArr);
  const nextRows = (Array.isArray(rows) ? rows : []).map((row) => {
    if (!row || typeof row !== 'object') return row;
    const next = { ...row };
    dedupedHeaders.forEach((h) => {
      if (h && next[h] == null) next[h] = '';
    });
    return next;
  });

  return {
    headers: dedupedHeaders,
    groupLabels: groupsArr,
    rows: nextRows,
    insertedMedical: needMedical,
    insertedOther: needOther,
    changed: true,
  };
}

/** Build Earned / Medical / Other column values for one Form 15 Part 1 employee (Form X rules). */
export function buildForm15Part1LeaveSectionValues(
  leaveRecord,
  approvedAvailed,
  leaveTypeLabels,
  section
) {
  return buildFormXLeaveSectionValues(leaveRecord, approvedAvailed, leaveTypeLabels, section);
}

/**
 * Apply Form 15 Part 1 leave columns for one employee row using Form X mapping.
 * @returns {number} cells written
 */
export function applyForm15Part1TamilNaduLeaveToRow(
  row,
  leaveRecord,
  form15Headers,
  leaveTypeLabels,
  approvedLeaveRecords,
  options = {}
) {
  const sectionHeaders = toFormXSectionHeadersFromForm15(form15Headers);
  if (!sectionHeaders) return 0;
  return applyFormXTamilNaduLeaveToRow(
    row,
    leaveRecord,
    sectionHeaders,
    leaveTypeLabels,
    approvedLeaveRecords,
    options
  );
}

function countNameDuplicatesInLeaveRecords(leaveRecords) {
  const counts = new Map();
  const bump = (name) => {
    const n = String(name || '').toLowerCase().trim();
    if (!n || n.length < 2) return;
    counts.set(n, (counts.get(n) || 0) + 1);
  };
  (Array.isArray(leaveRecords) ? leaveRecords : []).forEach((record) => {
    const { names } = collectLeaveRecordIdentityKeys(record);
    const seenForRecord = new Set();
    names.forEach((name) => {
      const full = String(name || '').toLowerCase().trim();
      if (!full || seenForRecord.has(full)) return;
      seenForRecord.add(full);
      bump(full);
      // Also count bare display name without "(EMPID)" so Form 15 name-only rows detect dups.
      const bare = full.replace(/\([^)]*\)\s*$/, '').replace(/\s+/g, ' ').trim();
      if (bare && bare !== full && !seenForRecord.has(bare)) {
        seenForRecord.add(bare);
        bump(bare);
      }
    });
  });
  return counts;
}

function rowHasEmployeeId(row, employeeIdHeader) {
  if (!employeeIdHeader || !row) return false;
  return String(row[employeeIdHeader] ?? '').trim() !== '';
}

/**
 * Match Form 15 row → leave record with ID-first priority.
 * Ambiguous name-only matches (duplicate names, no unique ID) are reported, not applied.
 */
export function matchForm15Part1LeaveRecord(lookup, row, employeeIdHeader, employeeNameHeader, nameDupCounts) {
  const empty = { record: null, matchKey: null, ambiguous: false, reason: 'not_found' };
  if (!lookup || !row) return empty;

  const rowId = employeeIdHeader ? String(row[employeeIdHeader] ?? '').toLowerCase().trim() : '';
  const rowName = employeeNameHeader ? String(row[employeeNameHeader] ?? '').toLowerCase().trim() : '';
  const rowNameBare = rowName.replace(/\([^)]*\)\s*$/, '').replace(/\s+/g, ' ').trim();

  if (rowId && lookup.byId.has(rowId)) {
    return { record: lookup.byId.get(rowId), matchKey: 'id', ambiguous: false, reason: 'matched_id' };
  }

  // Prefer ID match via paren token in leave display names (e.g. "Rajesh (VE0447)").
  if (rowId) {
    for (const [name, record] of lookup.byName) {
      if (name.includes(`(${rowId})`) || name.endsWith(` ${rowId}`)) {
        return { record, matchKey: 'id', ambiguous: false, reason: 'matched_id_in_name' };
      }
    }
  }

  const dupCount = Math.max(
    nameDupCounts?.get(rowName) || 0,
    nameDupCounts?.get(rowNameBare) || 0
  );
  const ambiguousPayload = {
    record: null,
    matchKey: null,
    ambiguous: true,
    reason: 'ambiguous_name',
    employeeName: String(row[employeeNameHeader] ?? '').trim(),
    employeeId: String(row[employeeIdHeader] ?? '').trim(),
  };

  // Duplicate bare names without a unique Employee ID → do not guess.
  if (dupCount > 1 && !rowHasEmployeeId(row, employeeIdHeader)) {
    return ambiguousPayload;
  }

  if (rowName && lookup.byName.has(rowName)) {
    if (dupCount > 1 && rowHasEmployeeId(row, employeeIdHeader)) {
      return { ...ambiguousPayload, reason: 'ambiguous_name_id_mismatch' };
    }
    return {
      record: lookup.byName.get(rowName),
      matchKey: 'name',
      ambiguous: false,
      reason: 'matched_name',
    };
  }

  if (rowNameBare && lookup.byName.has(rowNameBare)) {
    if (dupCount > 1 && rowHasEmployeeId(row, employeeIdHeader)) {
      return { ...ambiguousPayload, reason: 'ambiguous_name_id_mismatch' };
    }
    return {
      record: lookup.byName.get(rowNameBare),
      matchKey: 'name',
      ambiguous: false,
      reason: 'matched_name',
    };
  }

  // Fall back to existing fuzzy finder only when name is unique in the leave set.
  const fuzzy = findLeaveRecordForFormRow(lookup, row, employeeIdHeader, employeeNameHeader);
  if (fuzzy) {
    if (dupCount > 1 && !rowHasEmployeeId(row, employeeIdHeader)) {
      return ambiguousPayload;
    }
    return { record: fuzzy, matchKey: 'fuzzy', ambiguous: false, reason: 'matched_fuzzy' };
  }

  return {
    ...empty,
    employeeName: String(row[employeeNameHeader] ?? '').trim(),
    employeeId: String(row[employeeIdHeader] ?? '').trim(),
  };
}

function emptyLeaveTransferSummary() {
  return {
    totalEmployees: 0,
    matched: 0,
    updated: 0,
    notFound: [],
    ambiguous: [],
    insertedMedical: false,
    insertedOther: false,
  };
}

/**
 * Autofill Form 15 Part 1 leave register rows from Tamil Nadu Form X leave data rules.
 * Ensures Medical / Other columns exist, matches employees ID-first, and returns a summary.
 *
 * @returns {{ hits: number, summary: object, headers: string[], groupLabels: string[], rows: object[] }}
 */
export function applyForm15Part1TamilNaduLeaveAutofillWithSummary(
  mappedData,
  employeesForMapping,
  leaveRecords,
  leaveTypeLabels,
  tableHeaders,
  form15Headers,
  approvedLeaveRecords,
  options = {}
) {
  const ensured = ensureForm15Part1LeaveColumnsLikeFormX(
    tableHeaders,
    options.groupLabels,
    mappedData
  );
  const headers = ensured.headers;
  const groupLabels = ensured.groupLabels;
  const rows = ensured.rows;

  const resolved =
    form15Headers && !ensured.changed
      ? form15Headers
      : resolveForm15Part1LeaveColumnHeaders(headers, groupLabels);
  const sectionHeaders = toFormXSectionHeadersFromForm15(resolved);

  const summary = emptyLeaveTransferSummary();
  summary.totalEmployees = Array.isArray(rows) ? rows.length : 0;
  summary.insertedMedical = ensured.insertedMedical;
  summary.insertedOther = ensured.insertedOther;

  if (!sectionHeaders || !Object.values(sectionHeaders).some(Boolean)) {
    return { hits: 0, summary, headers, groupLabels, rows };
  }

  const leaveLookup = buildLeaveRecordLookupMap(leaveRecords);
  const nameDupCounts = countNameDuplicatesInLeaveRecords(leaveRecords);
  const { employeeNameHeader = '', employeeIdHeader = '' } = options.resolveEmployeeHeaders
    ? options.resolveEmployeeHeaders(headers)
    : {};

  let hits = 0;
  rows.forEach((row, index) => {
    if (!row || typeof row !== 'object') return;
    const empItem = employeesForMapping?.[index];
    const emp = empItem && (empItem.Employee || empItem.employee || empItem);

    const match = matchForm15Part1LeaveRecord(
      leaveLookup,
      row,
      employeeIdHeader,
      employeeNameHeader,
      nameDupCounts
    );

    if (match.ambiguous) {
      summary.ambiguous.push({
        employeeName: match.employeeName || '',
        employeeId: match.employeeId || '',
        reason: match.reason,
      });
      return;
    }

    if (!match.record) {
      // Leave columns stay blank when employee is on Form 15 but not in leave source (Form X / Leave API).
      summary.notFound.push({
        employeeName: match.employeeName || String(row[employeeNameHeader] ?? '').trim(),
        employeeId: match.employeeId || String(row[employeeIdHeader] ?? '').trim(),
      });
      return;
    }

    summary.matched += 1;
    const applied = applyFormXTamilNaduLeaveToRow(
      row,
      match.record,
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
    if (applied > 0) {
      hits += 1;
      summary.updated += 1;
    }
  });

  return { hits, summary, headers, groupLabels, rows };
}

/**
 * Autofill Form 15 Part 1 leave register rows from Tamil Nadu Form X leave data rules.
 * @returns {number} rows with at least one leave cell written
 */
export function applyForm15Part1TamilNaduLeaveAutofill(
  mappedData,
  employeesForMapping,
  leaveRecords,
  leaveTypeLabels,
  tableHeaders,
  form15Headers,
  approvedLeaveRecords,
  options = {}
) {
  const result = applyForm15Part1TamilNaduLeaveAutofillWithSummary(
    mappedData,
    employeesForMapping,
    leaveRecords,
    leaveTypeLabels,
    tableHeaders,
    form15Headers,
    approvedLeaveRecords,
    options
  );
  // Keep caller rows in sync when columns were inserted.
  if (Array.isArray(mappedData) && Array.isArray(result.rows) && result.rows !== mappedData) {
    mappedData.length = 0;
    result.rows.forEach((r) => mappedData.push(r));
  }
  if (typeof options.onLayoutEnsured === 'function' && result.summary) {
    options.onLayoutEnsured({
      headers: result.headers,
      groupLabels: result.groupLabels,
      insertedMedical: result.summary.insertedMedical,
      insertedOther: result.summary.insertedOther,
      summary: result.summary,
    });
  }
  return result.hits;
}

/** Human-readable Form 15 ← Form X leave transfer summary line. */
export function formatForm15Part1LeaveTransferSummary(summary) {
  if (!summary || typeof summary !== 'object') return '';
  const parts = [
    `Matched ${summary.matched || 0}`,
    `updated ${summary.updated || 0}`,
    `not found ${Array.isArray(summary.notFound) ? summary.notFound.length : 0}`,
  ];
  if (Array.isArray(summary.ambiguous) && summary.ambiguous.length > 0) {
    parts.push(`ambiguous ${summary.ambiguous.length}`);
  }
  if (summary.insertedMedical || summary.insertedOther) {
    const bands = [
      summary.insertedMedical ? 'Medical Leave' : null,
      summary.insertedOther ? 'Other Leave' : null,
    ]
      .filter(Boolean)
      .join(' + ');
    parts.push(`added columns: ${bands}`);
  }
  return `Form 15 Part 1 leave (Form X mapping): ${parts.join(', ')}.`;
}

function leaveMetricTypeFromHeader(header) {
  const s = normHeader(header);
  if (!s.includes('leave')) return null;
  if (s.includes('beginning') && s.includes('month')) return 'beginning';
  if (s.includes('earned') && (s.includes('period') || s.includes('during'))) return 'earned';
  if (s.includes('availed') && s.includes('month')) return 'availed';
  if (s.includes('balance') && s.includes('end') && s.includes('month')) return 'balance';
  return null;
}

/**
 * Collect leave cell values from an autofill/modal row in stable column order.
 * Prefers `headers` order, then any extra leave keys on the row.
 * Includes "0" (earned-during default).
 *
 * @returns {string[]}
 */
export function buildForm15Part1LeaveExportColumnValues(row, headers = []) {
  if (!row || typeof row !== 'object') return [];
  const values = [];
  const seen = new Set();
  const pushKey = (key) => {
    const k = String(key || '').trim();
    if (!k || seen.has(k) || String(k).startsWith('__')) return;
    if (!leaveMetricTypeFromHeader(k)) return;
    seen.add(k);
    const raw = row[k];
    if (raw == null) {
      values.push('');
      return;
    }
    values.push(String(raw).trim());
  };
  (Array.isArray(headers) ? headers : []).forEach(pushKey);
  Object.keys(row).forEach(pushKey);
  return values;
}

/**
 * Map modal leave values onto Excel physical columns that sit after the worker-id column.
 * Used when merged group headers ("Earned Leave") hide leaf titles so metric classification fails.
 *
 * @returns {(string|number|null)[]} same length as physicalColCount (or leaveColIndexes length)
 */
export function mapForm15Part1LeaveValuesToExcelLeaveColumns(
  row,
  headers,
  leaveColumnCount
) {
  const values = buildForm15Part1LeaveExportColumnValues(row, headers);
  const count = Math.max(0, Number(leaveColumnCount) || 0);
  const out = [];
  for (let i = 0; i < count; i += 1) {
    out.push(i < values.length ? values[i] : '');
  }
  return out;
}
