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
  applyFormXTamilNaduLeaveAutofill,
  applyFormXTamilNaduLeaveToRow,
  buildFormXLeaveSectionValues,
} from './formXTamilNaduLeave';

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

  const norm = (v) =>
    String(v || '')
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();

  const metricType = (header) => {
    const s = norm(header);
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
    const group = norm(groups[index]);
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
  const sectionHeaders = toFormXSectionHeadersFromForm15(form15Headers);
  if (!sectionHeaders) return 0;
  return applyFormXTamilNaduLeaveAutofill(
    mappedData,
    employeesForMapping,
    leaveRecords,
    leaveTypeLabels,
    tableHeaders,
    sectionHeaders,
    approvedLeaveRecords,
    options
  );
}
