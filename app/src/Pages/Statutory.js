import React, { useState, useEffect, useRef, useMemo, useCallback } from 'react';
import { Link, useSearchParams } from 'react-router-dom';
import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import './Statutory.css';
import './SEMaster.css';
import {
  checklistStateMatchesSiteState,
  fetchAllowedActCategoriesFromSites,
  getActCategoryFromActSector,
  industryLabelToActCategory,
  normalizeEmail,
  sectorMatchesInchargeSiteIndustries,
  statesFieldMatchesInchargeSiteStates
} from '../utils/siteInchargeScope';
import { resolveLoginEmailString } from '../utils/resolveLoginEmail';

/** @returns {(number | 'ellipsis')[]} */
function buildPaginationItems(currentPage, totalPages) {
  if (totalPages <= 0) return [];
  if (totalPages <= 9) {
    return Array.from({ length: totalPages }, (_, i) => i + 1);
  }
  const set = new Set([1, totalPages, currentPage]);
  for (let i = currentPage - 2; i <= currentPage + 2; i += 1) {
    if (i >= 1 && i <= totalPages) set.add(i);
  }
  const sorted = [...set].sort((a, b) => a - b);
  const out = [];
  for (let i = 0; i < sorted.length; i += 1) {
    if (i > 0 && sorted[i] - sorted[i - 1] > 1) {
      out.push('ellipsis');
    }
    out.push(sorted[i]);
  }
  return out;
}

/** Same as checklistNotifications/baseFormName: "Form B - TN LWF..." → "form b" so sibling rows match for Sector/category. */
function baseFormNameKey(name) {
  const n = String(name || '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ');
  const m = n.match(/^form\s+[a-z0-9]+/);
  return m ? m[0] : n;
}

/** Align Statutory rows with ChecklistBulk master (Sector, State, Act, Form Name, Description). */
function checklistBulkRowIdentityKey(row) {
  const norm = (v) =>
    String(v ?? '')
      .trim()
      .toLowerCase()
      .replace(/\s+/g, ' ');
  return [
    norm(row?.formName ?? row?.FormName),
    norm(row?.sector ?? row?.Sector),
    norm(row?.state ?? row?.State),
    norm(row?.act ?? row?.Act),
    norm(row?.description ?? row?.Description)
  ].join('\x1f');
}

/** Split comma-joined Statutory Site column values for comparisons (multi-site per state). */
function parseResolvedSiteTokens(resolved) {
  return String(resolved || '')
    .trim()
    .toLowerCase()
    .split(/\s*,\s*/)
    .map((p) => p.trim())
    .filter(Boolean);
}

function resolvedSitesOverlap(resolvedA, resolvedB) {
  const a = parseResolvedSiteTokens(resolvedA);
  const b = parseResolvedSiteTokens(resolvedB);
  if (a.length === 0 || b.length === 0) return false;
  const setA = new Set(a);
  return b.some((x) => setA.has(x));
}

function resolvedSiteMatchesSingleTarget(resolved, targetSiteSingle) {
  const t = String(targetSiteSingle || '').trim().toLowerCase();
  if (!t) return false;
  return parseResolvedSiteTokens(resolved).includes(t);
}

function resolvedSiteAnyAllowed(resolved, allowedSitesLowerSet) {
  return parseResolvedSiteTokens(resolved).some((x) => allowedSitesLowerSet.has(x));
}

function hasStatutoryDraftFileRef(row) {
  const d = row?.draftFile ?? row?.DraftFile;
  return d != null && String(d).trim() !== '' && String(d).trim() !== 'null';
}

function isNumericStatutoryBackendId(id) {
  return /^\d+$/.test(String(id ?? '').trim());
}

/** When this user is logged in, Statutory is filtered by ?site= (Delphi site → sector); separate from Site Management industry scope */
const SITE_WISE_STATUTORY_EMAIL = 'afrindinu14@gmail.com';
/** localStorage key for persisting the statutory month filter (based on saved due date) */
const STATUTORY_MONTH_FILTER_KEY = 'statutory_month_filter';
/** Formmaster templates API – form names in Statutory come from here (not ChecklistBulk/SEMaster) */
const FORMMASTER_TEMPLATES_API = '/server/formmaster_function/templates';

/** Form file modal: Zoho autofill + Excel table header row (single accent). */
const STAT_FORM_FILE_ACCENT_BG = '#3b82f6';
const STAT_FORM_FILE_ACCENT_BORDER = '#2563eb';
const STAT_FORM_FILE_ACCENT_BG_HOVER = '#2563eb';

/** Form U "format 2" – full column headers (Name of the employee, Employee Identification No., etc.) so downloaded draft shows full labels and template alignment is preserved */
const FORM_U_FORMAT2_HEADERS = [
  'S.No',
  'Name of the employee',
  'Employee Identification No.',
  'Gender',
  'Father / Spouse Name',
  'Date of Birth',
  'Date of Joining',
  'Designation',
  'Present Address',
  'Permanent Address',
  'Employee Identification No.',
  'Worker Identity No.',
  'Aadhaar',
  'Date on which joined',
  'Date on which left',
  'Period of service',
  'Bank A/c',
  'Photo'
];

const FORM_12_ALIGNMENT_HEADERS = [
  'Serial Number',
  'Name of the Worker',
  'Worker Identity No',
  'Gender',
  'Father / Spouse Name',
  'Date of Birth',
  'Present Address',
  'Permanent Address',
  'Aadhaar'
];

const MONTH_NAMES = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const MONTH_ABBR = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

/** Zoho People `bookedAndBalance` expects `from` / `to` as DD-Mon-YYYY (aligned with Statutory month picker). */
function zohoBookedBalanceRangeForUiMonth(selectedMonthStr) {
  const now = new Date();
  let monthIdx = now.getMonth();
  const idx = MONTH_NAMES.findIndex((m) =>
    m.toLowerCase().startsWith(String(selectedMonthStr || '').toLowerCase().trim())
  );
  if (idx >= 0) monthIdx = idx;
  const year = now.getFullYear();
  const first = new Date(year, monthIdx, 1);
  const last = new Date(year, monthIdx + 1, 0);
  const dd = (n) => String(n).padStart(2, '0');
  const fromDate = `${dd(first.getDate())}-${MONTH_ABBR[first.getMonth()]}-${first.getFullYear()}`;
  const toDate = `${dd(last.getDate())}-${MONTH_ABBR[last.getMonth()]}-${last.getFullYear()}`;
  return { fromDate, toDate };
}

function normalizeLooseHeaderText(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

/** Metric type for Form X leave register sub-columns */
function getLeaveColumnMetricType(header) {
  const headerLower = normalizeLooseHeaderText(header);
  if (!headerLower.includes('leave')) return null;
  if (headerLower.includes('earned') && (headerLower.includes('period') || headerLower.includes('during'))) {
    return 'earned';
  }
  if (headerLower.includes('beginning') && headerLower.includes('month')) return 'beginning';
  if (headerLower.includes('availed') && headerLower.includes('month')) return 'availed';
  if (headerLower.includes('balance') && headerLower.includes('end') && headerLower.includes('month')) {
    return 'balance';
  }
  return null;
}

const OTHER_LEAVE_ZOHO_ALIASES = {
  privilegeleave: ['privilege leave', 'privilegeleave', 'privilege', 'pl'],
  leavewithoutpay: ['leave without pay', 'leavewithoutpay', 'loss of pay', 'lop', 'lwp'],
  sickleave: ['sick leave', 'sickleave', 'sick', 'sl'],
  weddingleave: ['wedding leave', 'weddingleave', 'marriage leave', 'wedding'],
};

function leaveLabelMatchesOtherType(label, aliases) {
  const normalized = normalizeLooseHeaderText(label);
  if (!normalized) return false;
  return aliases.some(
    (alias) =>
      normalized === alias ||
      normalized.includes(alias) ||
      alias.includes(normalized)
  );
}

function getOtherLeaveTypeCell(leaveRecord, aliases, leaveTypeLabels = {}) {
  if (!leaveRecord || typeof leaveRecord !== 'object') return null;
  const skipKeys = new Set(['employee', 'employeeid', 'totals', 's.no', 'sno']);
  for (const key of Object.keys(leaveRecord)) {
    if (skipKeys.has(normalizeLooseHeaderText(key))) continue;
    const labelsToCheck = [key, leaveTypeLabels[key]].filter(Boolean);
    const matched = labelsToCheck.some((label) => leaveLabelMatchesOtherType(label, aliases));
    if (matched) return leaveRecord[key];
  }
  return null;
}

function parseLeaveCellObject(value) {
  if (value == null || value === '') return null;
  if (typeof value === 'object') return value;
  const raw = String(value).trim();
  if (!raw || raw === '{}') return null;
  try {
    const parsed = JSON.parse(raw);
    return parsed && typeof parsed === 'object' ? parsed : null;
  } catch (_) {
    return null;
  }
}

function getOtherLeaveTypeMetrics(leaveRecord, aliases, leaveTypeLabels = {}) {
  const raw = getOtherLeaveTypeCell(leaveRecord, aliases, leaveTypeLabels);
  const parsed = parseLeaveCellObject(raw);
  if (!parsed || typeof parsed !== 'object') {
    return { beginning: 0, availed: 0, balance: 0, hasData: false };
  }
  const balanceRaw = parsed.paidBalance ?? parsed.balance ?? parsed.Balance;
  const availedRaw = parsed.paidBooked ?? parsed.booked ?? parsed.Booked;
  const hasBalance = balanceRaw != null && balanceRaw !== '';
  const hasAvailed = availedRaw != null && availedRaw !== '';
  const balanceNum = hasBalance ? Number(balanceRaw) : NaN;
  const availedNum = hasAvailed ? Number(availedRaw) : NaN;
  const beginningRaw =
    parsed.beginningBalance ??
    parsed.openingBalance ??
    parsed.opening ??
    parsed.beginning ??
    (Number.isFinite(balanceNum) && Number.isFinite(availedNum) ? balanceNum + availedNum : balanceNum);
  const beginningNum = beginningRaw != null && beginningRaw !== '' ? Number(beginningRaw) : NaN;
  return {
    beginning: Number.isFinite(beginningNum) ? beginningNum : 0,
    availed: Number.isFinite(availedNum) ? availedNum : 0,
    balance: Number.isFinite(balanceNum) ? balanceNum : 0,
    hasData: Number.isFinite(beginningNum) || Number.isFinite(availedNum) || Number.isFinite(balanceNum),
  };
}

function getCombinedOtherLeaveMetrics(leaveRecord, leaveTypeLabels = {}) {
  const totals = { beginning: 0, availed: 0, balance: 0 };
  let hasAnyValue = false;
  Object.values(OTHER_LEAVE_ZOHO_ALIASES).forEach((aliases) => {
    const metrics = getOtherLeaveTypeMetrics(leaveRecord, aliases, leaveTypeLabels);
    if (!metrics.hasData) return;
    totals.beginning += metrics.beginning;
    totals.availed += metrics.availed;
    totals.balance += metrics.balance;
    hasAnyValue = true;
  });
  if (!hasAnyValue) return { beginning: '', availed: '', balance: '' };
  return {
    beginning: String(totals.beginning),
    availed: String(totals.availed),
    balance: String(totals.balance),
  };
}

/**
 * Form X: map detail headers to Earned Leave vs Other Leave bands (merged row above headers).
 */
function resolveFormXLeaveSectionHeaders(headers, groupLabels) {
  const headersArr = Array.isArray(headers) ? headers : [];
  const groups =
    Array.isArray(groupLabels) && groupLabels.length === headersArr.length ? groupLabels : null;
  const result = {
    earnedEarned: null,
    earnedAvailed: null,
    otherBeginning: null,
    otherAvailed: null,
    otherBalance: null,
  };
  const lists = { beginning: [], availed: [], balance: [], earned: [] };

  headersArr.forEach((header, index) => {
    const metric = getLeaveColumnMetricType(header);
    if (!metric) return;
    if (metric === 'earned') lists.earned.push(header);
    if (metric === 'beginning') lists.beginning.push(header);
    if (metric === 'availed') lists.availed.push(header);
    if (metric === 'balance') lists.balance.push(header);

    if (!groups) return;
    const group = normalizeLooseHeaderText(groups[index]);
    if (group.includes('earned leave')) {
      if (metric === 'earned' && !result.earnedEarned) result.earnedEarned = header;
      if (metric === 'availed' && !result.earnedAvailed) result.earnedAvailed = header;
    }
    if (group.includes('other leave')) {
      if (metric === 'beginning' && !result.otherBeginning) result.otherBeginning = header;
      if (metric === 'availed' && !result.otherAvailed) result.otherAvailed = header;
      if (metric === 'balance' && !result.otherBalance) result.otherBalance = header;
    }
  });

  if (!result.earnedEarned && lists.earned.length > 0) result.earnedEarned = lists.earned[0];
  if (!result.earnedAvailed && lists.availed.length > 0) result.earnedAvailed = lists.availed[0];
  if (!result.otherBeginning && lists.beginning.length >= 3) result.otherBeginning = lists.beginning[2];
  if (!result.otherAvailed && lists.availed.length >= 3) result.otherAvailed = lists.availed[2];
  if (!result.otherBalance && lists.balance.length >= 3) result.otherBalance = lists.balance[2];

  return result;
}

/** Extract full month name from a due date string (e.g. "15-Mar", "March 2024", "15-03-2024") for saving into month filter */
const getMonthFromDueDate = (dueDateStr) => {
  if (!dueDateStr || typeof dueDateStr !== 'string') return null;
  const s = dueDateStr.trim().toLowerCase();
  if (s.includes('monthly basis')) return null; // Don't change filter for "Monthly Basis"
  for (let i = 0; i < MONTH_NAMES.length; i++) {
    if (s.includes(MONTH_NAMES[i].toLowerCase()) || s.includes(MONTH_ABBR[i].toLowerCase())) return MONTH_NAMES[i];
  }
  const mmMatch = s.match(/\b(0?[1-9]|1[0-2])\b/);
  if (mmMatch) {
    const monthNum = parseInt(mmMatch[1], 10);
    if (monthNum >= 1 && monthNum <= 12) return MONTH_NAMES[monthNum - 1];
  }
  return null;
};

/** Normalize month for backend: never send empty string, use null so Month Filter column gets a proper value or null */
const monthForBackend = (value) => {
  const v = value && String(value).trim();
  return v !== '' ? v : null;
};

/** Map month filter / partial names to full calendar month name */
const resolveToFullMonthName = (raw) => {
  if (!raw || typeof raw !== 'string') return null;
  const s = raw.trim().toLowerCase();
  if (!s) return null;
  for (let i = 0; i < MONTH_NAMES.length; i++) {
    const full = MONTH_NAMES[i].toLowerCase();
    const ab = MONTH_ABBR[i].toLowerCase();
    if (full === s || ab === s) return MONTH_NAMES[i];
    if (full.startsWith(s) || s.startsWith(full.slice(0, 3))) return MONTH_NAMES[i];
    if (s.startsWith(ab.slice(0, 3))) return MONTH_NAMES[i];
  }
  return null;
};

/** Align checklist bulk rows (often missing MonthFilter) with saved statutory rows after Autofill→Save. */
function statutoryDedupeMonthNorm(item, uiMonthFallback) {
  const raw = item?.monthFilter ?? item?.MonthFilter ?? item?.monthfilter;
  if (raw != null && String(raw).trim() !== '') {
    const full = resolveToFullMonthName(String(raw).trim());
    const eff = full || String(raw).trim();
    return eff.toLowerCase().substring(0, 3);
  }
  const fromDue = getMonthFromDueDate(item?.dueDate || item?.DueDate || '');
  if (fromDue) return String(fromDue).trim().toLowerCase().substring(0, 3);
  const ud = String(item?.dueDate || '').trim().toLowerCase();
  if (ud.includes('monthly basis') && uiMonthFallback && String(uiMonthFallback).trim()) {
    const full = resolveToFullMonthName(String(uiMonthFallback).trim());
    const eff = full || String(uiMonthFallback).trim();
    return eff.toLowerCase().substring(0, 3);
  }
  return 'nomonth';
}

function squashStatutoryKeyPart(v) {
  return String(v ?? '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ');
}

/** Prefer rows with draft file / proof / real Catalyst id when collapsing duplicates. */
function statutoryRowRichnessScore(r) {
  let s = 0;
  const df = r?.draftFile ?? r?.DraftFile;
  if (df && df !== 'null' && String(df).trim() !== '') s += 200;
  const pf = r?.proofSubmissionFile ?? r?.ProofSubmissionFile;
  if (pf != null && String(pf).trim() !== '' && String(pf).trim() !== 'null') s += 150;
  if (/^\d+$/.test(String(r?.id ?? '').trim())) s += 100;
  if (!r.isBulkImported) s += 40;
  const ff = r?.formFile ?? r?.FormFile;
  if (ff && ff !== 'null' && String(ff).trim() !== '') s += 25;
  return s;
}

/** Bulk rows often have no Site; saved statutory rows do — merge so dedupe keys align. */
function statutoryCollapseSiteKey(item, allRows, resolveSiteFn) {
  let explicit = '';
  let inferred = '';
  try {
    explicit = String(item.site ?? item.Site ?? '').trim();
  } catch (_) {}
  try {
    if (typeof resolveSiteFn === 'function') {
      inferred = String(resolveSiteFn(item, allRows) || '').trim();
    }
  } catch (_) {}
  const merged = explicit || inferred || '';
  return squashStatutoryKeyPart(merged) || 'nosite';
}

/** Keep master rows distinct during UI dedupe (bulk/checklist/formmaster placeholders). */
function statutoryUiStableRowSuffix(item) {
  const idStr = String(item?.id ?? '').trim();
  if (item?.isBulkImported && idStr) return idStr;
  if (item?.checklistId != null && String(item.checklistId).trim() !== '') {
    return `chk:${String(item.checklistId).trim()}`;
  }
  if ((item?.isFromChecklist || /^checklist_/i.test(idStr)) && idStr) return idStr;
  if ((item?.isFromFormmaster || /^formmaster_/i.test(idStr)) && idStr) return idStr;
  return '';
}

/**
 * After merge, checklist bulk + saved statutory rows can produce two lines that differ only by whitespace/casing
 * in description/sector — strict formKey buckets duplicate them. Collapse to one row (prefer draft + numeric ROWID).
 */
function collapseStatutoryDisplayDuplicates(rows, uiMonthFallback, resolveSiteFn) {
  if (!Array.isArray(rows) || rows.length <= 1) return rows;

  const looseKey = (item) => {
    const mn = statutoryDedupeMonthNorm(item, uiMonthFallback);
    const sk = statutoryCollapseSiteKey(item, rows, resolveSiteFn);
    const base = [
      squashStatutoryKeyPart(item.formName || item.FormName),
      squashStatutoryKeyPart(item.act || item.Act),
      squashStatutoryKeyPart(item.description || item.Description),
      squashStatutoryKeyPart(item.sector || item.Sector) || 'nosector',
      squashStatutoryKeyPart(item.state || item.State) || 'nostate',
      mn,
      sk
    ].join('\x1f');
    const stable = statutoryUiStableRowSuffix(item);
    return stable ? `${base}\x1f${stable}` : base;
  };

  const bestByKey = new Map();
  for (const row of rows) {
    const k = looseKey(row);
    const prev = bestByKey.get(k);
    if (!prev) {
      bestByKey.set(k, row);
      continue;
    }
    bestByKey.set(k, statutoryRowRichnessScore(row) >= statutoryRowRichnessScore(prev) ? row : prev);
  }
  return Array.from(bestByKey.values());
}

/** Prefer inferred site for grouping so bulk rows (no Site column) align with saved statutory rows that store Site. */
function statutoryUiSiteDedupeKey(item, allRows, resolveSiteFn) {
  let inferred = '';
  try {
    if (typeof resolveSiteFn === 'function') {
      inferred = String(resolveSiteFn(item, allRows) || '').trim();
    }
  } catch (_) {}
  const explicit = String(item.site ?? item.Site ?? '').trim();
  const merged = inferred || explicit || '';
  return squashStatutoryKeyPart(merged) || 'nosite';
}

/**
 * Loose UI bucket for Statutory grid: same visible Form label + month + site — ignores Act/Sector/State/Description
 * mismatches between ChecklistBulk placeholders and Catalyst rows (fixes duplicate Form N after Autofill save).
 * Uses full squashed form name so "Form 15 Part 1" vs "Part 2" stay distinct.
 */
function statutoryUiDuplicateLooseKey(item, allRows, uiMonthFallback, resolveSiteFn) {
  return [
    squashStatutoryKeyPart(item.formName || item.FormName),
    statutoryDedupeMonthNorm(item, uiMonthFallback),
    statutoryUiSiteDedupeKey(item, allRows, resolveSiteFn)
  ].join('\x1f');
}

/** Final pass on filtered rows: one row per loose bucket; prefer row that actually stores DraftFile. */
function dedupeStatutoryRowsForDisplay(rows, uiMonthFallback, resolveSiteFn) {
  if (!Array.isArray(rows) || rows.length <= 1) return rows;

  const key = (item) => statutoryUiDuplicateLooseKey(item, rows, uiMonthFallback, resolveSiteFn);
  const grouped = new Map();
  for (const row of rows) {
    const k = key(row);
    if (!grouped.has(k)) grouped.set(k, []);
    grouped.get(k).push(row);
  }

  const out = [];
  grouped.forEach((grp) => {
    if (grp.length === 1) {
      out.push(grp[0]);
      return;
    }
    const withDraft = grp.filter((r) => {
      const d = r?.draftFile ?? r?.DraftFile;
      return d && d !== 'null' && String(d).trim() !== '';
    });
    if (withDraft.length > 0) {
      withDraft.sort((a, b) => statutoryRowRichnessScore(b) - statutoryRowRichnessScore(a));
      out.push(withDraft[0]);
    } else {
      grp.sort((a, b) => statutoryRowRichnessScore(b) - statutoryRowRichnessScore(a));
      out.push(grp[0]);
    }
  });
  return out;
}

/** Persist MonthFilter in Catalyst (due-date-derived month, else UI month / row — needed when Due Date is "Monthly Basis"). */
const resolveMonthFilterForSave = (dueDate, uiSelectedMonth, rowMonthOptional) => {
  let candidate = getMonthFromDueDate(dueDate);
  const dueStr = dueDate != null ? String(dueDate).trim() : '';
  if (!candidate && dueStr.toLowerCase().includes('monthly basis')) {
    candidate = rowMonthOptional || uiSelectedMonth;
  }
  if (!candidate) {
    const rowTrim = rowMonthOptional != null ? String(rowMonthOptional).trim() : '';
    candidate = rowTrim || (uiSelectedMonth != null ? String(uiSelectedMonth).trim() : '') || null;
  }
  // Final safety fallback: always persist a month when row exists.
  if (!candidate) {
    candidate = MONTH_NAMES[new Date().getMonth()];
  }
  const str = candidate != null ? String(candidate).trim() : '';
  if (!str) return monthForBackend(MONTH_NAMES[new Date().getMonth()]);
  const full = resolveToFullMonthName(str);
  return monthForBackend(full || str || MONTH_NAMES[new Date().getMonth()]);
};

const extractYearFromDueDate = (dueDateStr) => {
  if (dueDateStr == null || dueDateStr === '') return null;
  const s = String(dueDateStr);
  const m4 = s.match(/\b(19|20)\d{2}\b/);
  if (m4) return parseInt(m4[0], 10);
  const m2 = s.match(/[-/](\d{2})\b/);
  if (m2) {
    const yy = parseInt(m2[1], 10);
    if (yy >= 0 && yy <= 99) return yy <= 30 ? 2000 + yy : 1900 + yy;
  }
  return null;
};

const ordinalDay = (n) => {
  const j = n % 10;
  const k = n % 100;
  if (k >= 11 && k <= 13) return `${n}th`;
  if (j === 1) return `${n}st`;
  if (j === 2) return `${n}nd`;
  if (j === 3) return `${n}rd`;
  return `${n}th`;
};

/** Form W style: "Wage Period from 1st April 2026 to 30th April 2026" */
const buildRegisterWagePeriodLine = (monthName, year) => {
  const idx = MONTH_NAMES.findIndex((m) => m.toLowerCase() === String(monthName || '').toLowerCase().trim());
  if (idx < 0) return '';
  const y = Number(year);
  if (!Number.isFinite(y) || y < 1900) return '';
  const lastDay = new Date(y, idx + 1, 0).getDate();
  const mon = MONTH_NAMES[idx];
  return `Wage Period from ${ordinalDay(1)} ${mon} ${y} to ${ordinalDay(lastDay)} ${mon} ${y}`;
};

/** Infer calendar month from an existing wage-period sentence in the template (e.g. "…April 202…"). */
const inferMonthFromWagePeriodLine = (line) => {
  if (!line || typeof line !== 'string') return null;
  const lower = line.toLowerCase().replace(/\s+/g, ' ');
  for (let i = 0; i < MONTH_NAMES.length; i++) {
    if (lower.includes(MONTH_NAMES[i].toLowerCase())) return MONTH_NAMES[i];
  }
  return null;
};

const extractYearFromWagePeriodLine = (line) => {
  if (!line || typeof line !== 'string') return null;
  const m = line.match(/\b(19|20)\d{2}\b/);
  if (m) return parseInt(m[0], 10);
  return null;
};

/**
 * Fix templates where Excel shows a truncated year (e.g. "202" instead of "2026") after the month name.
 */
const repairTruncatedYearsInWagePeriodLine = (line, year) => {
  const yy = Number(year);
  const yUse = Number.isFinite(yy) && yy >= 1900 ? yy : new Date().getFullYear();
  let s = String(line || '').replace(/\s+/g, ' ').trim();
  if (!s) return '';
  return s.replace(/\b(2[0-9]{2})(?!\d)/g, (tok) => {
    const n = parseInt(tok, 10);
    if (n >= 200 && n <= 299) return String(yUse);
    return tok;
  });
};

/** Form V etc.: show "1".."31" under merged daily-hours header instead of repeating the full parent key. */
const formatStatutoryTableHeaderLabel = (headerKey) => {
  const h = String(headerKey || '');
  const m = h.match(/^(.+)_(\d{1,2})$/);
  if (!m) return h;
  const day = parseInt(m[2], 10);
  if (day < 1 || day > 31) return h;
  const parent = m[1].toLowerCase();
  const looksDatesBand = /^dates$/i.test(parent.trim());
  const looksDailyHoursBand =
    (parent.includes('daily') && parent.includes('hour')) ||
    (parent.includes('hour') && parent.includes('work') && parent.includes('overtime')) ||
    (parent.includes('daily') && parent.includes('work'));
  if (looksDatesBand || looksDailyHoursBand) return String(day);
  const comp = h.match(/^(.+)_([\s\S]+)$/);
  if (comp) {
    const parentNorm = comp[1].toLowerCase().replace(/\s+/g, ' ').trim();
    if (parentNorm.includes('components') && parentNorm.includes('remuneration')) {
      return comp[2].trim();
    }
  }
  return h;
};

/**
 * If headers use `Parent_1` .. `Parent_30` for daily hours, return the first day column
 * index and the parent label for a merged top row (Form V over Excel’s day grid).
 */
const getDailyHoursGridGroupInfo = (tableHeaders) => {
  if (!Array.isArray(tableHeaders) || tableHeaders.length < 2) return null;
  for (let i = 0; i < tableHeaders.length; i++) {
    const m = String(tableHeaders[i] || '').match(/^(.+)_(\d{1,2})$/);
    if (!m) continue;
    const day = parseInt(m[2], 10);
    if (day < 1 || day > 31) continue;
    const parent = m[1].toLowerCase();
    const looksDatesBand = /^dates$/i.test(parent.trim());
    const looksDailyHoursBand =
      (parent.includes('daily') && parent.includes('hour')) ||
      (parent.includes('hour') && parent.includes('work') && parent.includes('overtime')) ||
      (parent.includes('daily') && parent.includes('work'));
    if (looksDatesBand || looksDailyHoursBand) {
      return { startIndex: i, parentLabel: m[1].trim() };
    }
  }
  return null;
};

/**
 * Form D: consecutive `Components of Remuneration_<sub>` headers → merged top row + sub row (Excel model).
 */
const getFormDRemunerationGroupInfo = (tableHeaders) => {
  if (!Array.isArray(tableHeaders) || tableHeaders.length < 2) return null;
  const splitRe = /^(.+)_([\s\S]+)$/;
  let startIndex = -1;
  for (let i = 0; i < tableHeaders.length; i++) {
    const h = String(tableHeaders[i] || '');
    const m = h.match(splitRe);
    if (!m) continue;
    const parentNorm = m[1].toLowerCase().replace(/\s+/g, ' ').trim();
    if (parentNorm.includes('components') && parentNorm.includes('remuneration')) {
      startIndex = i;
      break;
    }
  }
  if (startIndex < 0) return null;
  const firstMatch = String(tableHeaders[startIndex]).match(splitRe);
  if (!firstMatch) return null;
  const firstParent = firstMatch[1].trim();
  let endIndex = startIndex;
  while (endIndex < tableHeaders.length) {
    const h = String(tableHeaders[endIndex] || '');
    const m = h.match(splitRe);
    if (!m) break;
    if (m[1].trim().toLowerCase() !== firstParent.toLowerCase()) break;
    endIndex += 1;
  }
  const count = endIndex - startIndex;
  if (count < 2) return null;
  return { startIndex, endIndex, parentLabel: firstParent };
};

const isFormDRemunerationFormContext = (formHeader, rowItem, fileName, tableHeaders) => {
  const parts = [
    rowItem?.formName,
    rowItem?.FormName,
    fileName,
    formHeader?.title,
    formHeader?.subtitle
  ]
    .join(' ')
    .toLowerCase();
  if (!/\bform\s*d\b|\bform_d\b|\bform-d\b/i.test(parts)) return false;
  return getFormDRemunerationGroupInfo(tableHeaders || []) != null;
};

/**
 * Form 12: factory / contractor / registration are often one merged cell (newlines or one long line).
 * Returns three { label, value } parts or null.
 */
const splitForm12FactoryContractorRegBlock = (text) => {
  const t = String(text || '')
    .replace(/\r\n/g, '\n')
    .trim();
  if (t.length < 30) return null;
  if (!/name\s+and\s+address\s+of\s+the\s+factory/i.test(t)) return null;
  if (!/registration\s+no\.?/i.test(t)) return null;

  const lines = t.split('\n').map((l) => l.trim()).filter(Boolean);
  if (lines.length >= 2) {
    const out = [];
    for (const line of lines) {
      const m = line.match(
        /^(Name and Address of the Factory|Name and Address of the Contractor|Registration No\.?)\s*:\s*(.*)$/i
      );
      if (m) {
        out.push({
          label: `${m[1].replace(/\s+/g, ' ').trim()}:`,
          value: (m[2] || '').trim()
        });
      }
    }
    if (out.length >= 2) return out;
  }

  const reF = /Name and Address of the Factory\s*:\s*([\s\S]+?)(?=\s*Name and Address of the Contractor\s*:|$)/i;
  const reC = /Name and Address of the Contractor\s*:\s*([\s\S]*?)(?=\s*Registration No\.?\s*:|$)/i;
  const reR = /Registration No\.?\s*:\s*([\s\S]*)$/i;
  const mf = t.match(reF);
  const mc = t.match(reC);
  const mr = t.match(reR);
  if (mf && mr) {
    return [
      { label: 'Name and Address of the Factory:', value: (mf[1] || '').trim() },
      { label: 'Name and Address of the Contractor:', value: (mc ? mc[1] : '').trim() },
      { label: 'Registration No:', value: (mr[1] || '').trim() }
    ];
  }
  return null;
};

/**
 * Form 10 (Overtime Muster / Rule 78): factory and contractor in one cell — no "Registration No."
 * "Month ending" is usually a separate cell; handled in parseExcelForm.
 */
const splitForm10FactoryContractorBlock = (text) => {
  const t = String(text || '')
    .replace(/\r\n/g, '\n')
    .trim();
  if (t.length < 25) return null;
  if (!/name\s+and\s+address\s+of\s+the\s+factory/i.test(t)) return null;
  if (!/name\s+and\s+address\s+of\s+the\s+contractor/i.test(t)) return null;
  if (/registration\s+no\.?/i.test(t)) return null;

  const lines = t.split('\n').map((l) => l.trim()).filter(Boolean);
  if (lines.length >= 2) {
    const out = [];
    for (const line of lines) {
      const m = line.match(
        /^(Name and Address of the Factory|Name and Address of the Contractor)\s*:\s*(.*)$/i
      );
      if (m) {
        out.push({
          label: `${m[1].replace(/\s+/g, ' ').trim()}:`,
          value: (m[2] || '').trim()
        });
      }
    }
    if (out.length >= 2) return out;
  }

  const reF = /Name and Address of the Factory\s*:\s*([\s\S]+?)(?=\s*Name and Address of the Contractor\s*:|$)/i;
  const reC = /Name and Address of the Contractor\s*:\s*([\s\S]*?)(?=\s*Month\s*ending\s*:|$)/i;
  const mf = t.match(reF);
  const mc = t.match(reC);
  if (mf && mc) {
    return [
      { label: 'Name and Address of the Factory:', value: (mf[1] || '').trim() },
      { label: 'Name and Address of the Contractor:', value: (mc[1] || '').trim() }
    ];
  }
  return null;
};

/** Form I: merged heading block should render as separate lines (Form - I / REGISTER OF WORKMEN / legal lines). */
const splitFormIHeaderBlock = (text) => {
  const t = String(text || '')
    .replace(/\r\n/g, '\n')
    .replace(/\s+\n/g, '\n')
    .replace(/\n\s+/g, '\n')
    .trim();
  if (!t) return null;
  if (!/\bform\s*[-–]?\s*i\b/i.test(t)) return null;
  if (!/register\s+of\s+workmen/i.test(t)) return null;

  const lines = t
    .split('\n')
    .map((l) => l.replace(/\s+/g, ' ').trim())
    .filter(Boolean);
  if (lines.length === 0) return null;

  const titleLine = lines.find((l) => /\bform\s*[-–]?\s*i\b/i.test(l)) || 'Form - I';
  const subtitleLine = lines.find((l) => /register\s+of\s+workmen/i.test(l)) || 'REGISTER OF WORKMEN';
  const bracketRule = lines.find((l) => /^\[.*\]$/.test(l) || /see\s+sub-?rule/i.test(l)) || '';
  const actLine = lines.find((l) => /tamil\s+nadu\s+industrial\s+establishments/i.test(l)) || '';
  const reference = [bracketRule, actLine].filter(Boolean).join('\n').trim();

  return {
    title: titleLine,
    subtitle: subtitleLine,
    reference
  };
};

const escapeRegExp = (s) => String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

/** Remove echoed form title from merged-cell subtitle text (Form 10 / Form 12 templates repeat "FORM No. N" in the same block as REGISTER / prescribed lines). */
const stripSubtitleTitleEcho = (formTitleText, chunk) => {
  const title = String(formTitleText || '').trim();
  if (!title || chunk == null) return String(chunk || '').trim();
  const tn = title.replace(/\s+/g, ' ').trim().toLowerCase();
  const titleRe = new RegExp(`^\\s*${escapeRegExp(title)}\\s*`, 'i');
  return String(chunk)
    .replace(/\r\n/g, '\n')
    .split('\n')
    .map((line) => {
      let l = line.trim();
      if (!l) return '';
      if (l.replace(/\s+/g, ' ').trim().toLowerCase() === tn) return '';
      l = l.replace(titleRe, '').trim();
      return l;
    })
    .filter(Boolean)
    .join('\n')
    .trim();
};

/** Form 25 — contractor / establishment / principal employer (wide layouts; value may be in cells to the right). */
const FORM25_SHEET_HEADER_SPECS = [
  {
    key: 'form25_contractor',
    label: 'Name and Address of the Contractor:',
    res: [/^Name and Address of the Contractor\s*:\s*(.*)$/is]
  },
  {
    key: 'form25_establishment',
    label: 'Name and address of establishment in/ under which contract is carried on:',
    res: [
      /^Name and address of establishment\s+in\s*\/\s*under which contract is carried on\s*:\s*(.*)$/is,
      /^Name\s+and\s+address\s+of\s+establishment[\s\S]{0,220}?contract\s+is\s+carried\s+on\s*:\s*(.*)$/is
    ]
  },
  {
    key: 'form25_principal_employer',
    label: 'Name and address of Principal Employer:',
    res: [/^Name and address of Principal Employer\s*:\s*(.*)$/is]
  }
];

const scanForm25RowValueAfterLabelColumn = (row, labelCol, effectiveSheetCols, getMergedAwareCellText) => {
  const maxC = Math.max(80, effectiveSheetCols);
  for (let nc = labelCol + 1; nc < Math.min(labelCol + 45, maxC); nc++) {
    const t = String(getMergedAwareCellText(row, nc) || '').trim();
    if (!t) continue;
    const tl = t.toLowerCase();
    if (/^name\s+and\s+address/i.test(tl)) continue;
    if (/^nature\s+and\s+location/i.test(tl)) continue;
    if (/^for\s+the\s+month/i.test(tl)) continue;
    if (/^approved\s+festival/i.test(tl)) continue;
    return t;
  }
  return '';
};

const enrichForm25AddressHeadersFromSheet = (headerRowIndex, effectiveSheetCols, getMergedAwareCellText) => {
  const results = [];
  for (const spec of FORM25_SHEET_HEADER_SPECS) {
    const patterns = spec.res || [];
    outer: for (let r = 0; r < headerRowIndex; r++) {
      for (let c = 0; c < Math.max(80, effectiveSheetCols); c++) {
        const raw = String(getMergedAwareCellText(r, c) || '').trim();
        if (!raw) continue;
        let m = null;
        for (let pi = 0; pi < patterns.length; pi++) {
          m = raw.match(patterns[pi]);
          if (m) break;
        }
        if (!m) continue;
        let val = (m[1] || '').trim().replace(/\s+/g, ' ');
        if (!val) val = scanForm25RowValueAfterLabelColumn(r, c, effectiveSheetCols, getMergedAwareCellText);
        results.push({ label: spec.label, value: val, key: spec.key });
        break outer;
      }
    }
  }
  return results;
};

/** Tamil Nadu LWF Form C–style grids: fixed category rows × quarterly columns — not one Zoho row per employee. */
const isFixedRowAggregateComplianceTable = (headers) => {
  const list = Array.isArray(headers) ? headers.map((h) => String(h || '').toLowerCase()) : [];
  if (list.length < 2) return false;
  const hasQuarterCols = list.some((h) => /quarter\s+ending/.test(h));
  const hasFinesDetailsCol = list.some((h) =>
    /details\s+of\s+fines|unpaid\s+accumulations/.test(h)
  );
  const looksLikeEmployeeRegister =
    list.some((h) => /\bs\.?\s*no\b|serial\s+number/.test(h)) &&
    list.some((h) => /name\s+of\s+(the\s+)?employee/.test(h));
  return hasQuarterCols && hasFinesDetailsCol && !looksLikeEmployeeRegister;
};

/**
 * Vertically merged body cells (common on TN LWF Form C) repeat the same merged values for each
 * physical sheet row; the parser then emits duplicate logical rows. Keep one row per consecutive run.
 */
const dedupeConsecutiveFixedAggregateBodyRows = (rows) => {
  if (!Array.isArray(rows) || rows.length < 2) return rows;
  const signature = (row) => {
    if (Array.isArray(row)) return row.map((v) => String(v ?? '').trim()).join('\x1f');
    if (row && typeof row === 'object') return Object.values(row).map((v) => String(v ?? '').trim()).join('\x1f');
    return String(row ?? '').trim();
  };
  const out = [];
  let prevSig = null;
  for (const row of rows) {
    const sig = signature(row);
    if (sig === prevSig) continue;
    prevSig = sig;
    out.push(row);
  }
  return out;
};

/** LWF Form-C: “[See rule … Labour Welfare …]” is rendered on its own line under the title (not the generic reference block). */
const shouldAppendReferenceToFormCTitle = (formHeader) => {
  const title = String(formHeader?.title || '').trim();
  const ref = String(formHeader?.reference || '').trim();
  if (!title || !ref) return false;
  if (!/\bform\s*[-–]?\s*c\b/i.test(title)) return false;
  return /^\s*\[\s*see\s+rule\s+\d+/i.test(ref) && /labour\s+welfare/i.test(ref);
};

const shouldPlaceFormBReferenceBeforeSubtitle = (formHeader) => {
  const title = String(formHeader?.title || '').trim();
  const subtitle = String(formHeader?.subtitle || '').trim();
  const ref = String(formHeader?.reference || '').trim();
  if (!title || !ref) return false;
  if (!/\bform\s*["']?\s*b\b/i.test(title)) return false;
  if (subtitle && !/register\s+of\s+wages/i.test(subtitle)) return false;
  return /labou?r\s*welfare\s*fund\s*rules?/i.test(ref);
};

const shouldPlaceFormAReferenceBeforeSubtitle = (formHeader) => {
  const title = String(formHeader?.title || '').trim();
  const subtitle = String(formHeader?.subtitle || '').trim();
  const ref = String(formHeader?.reference || '').trim();
  if (!ref) return false;
  const looksLikeFormA = /\bform\s*[-"']?\s*a\b/i.test(title) || /maternity\s+benefit\s+rules?/i.test(ref);
  if (!looksLikeFormA) return false;
  if (subtitle && !/register\s+of\s+muster\s+roll/i.test(subtitle)) return false;
  return /\bsee\s+sub-?rule\b/i.test(ref) && /maternity\s+benefit\s+rules?/i.test(ref);
};

const isFormAMusterRollContext = (formHeader) => {
  const title = String(formHeader?.title || '').trim();
  const subtitle = String(formHeader?.subtitle || '').trim();
  const ref = String(formHeader?.reference || '').trim();
  const text = [title, subtitle, ref].join(' ').toLowerCase();
  return (
    (/\bform\s*[-"']?\s*a\b/.test(text) || /maternity\s+benefit\s+rules?/.test(text)) &&
    /register\s+of\s+muster\s+roll/.test(text)
  );
};

const shouldHideFormAHeaderField = (formHeader, field) => {
  if (!isFormAMusterRollContext(formHeader)) return false;
  const label = String(field?.label || '').replace(/\s+/g, ' ').trim().toLowerCase();
  return /^\(?10\)?\s*date\s+on\s+which\s+payment\s+was\s+made\s*:?$/.test(label);
};

const FORM_A_MARKING_LEGEND = [
  'H for holidays allowed',
  'W/D for work on double wages',
  'W/H for work with substituted holiday',
  'N/E if not eligible for the wages'
];

const isFormVIFestivalContext = (formHeader, rowItem, fileName, tableHeaders) => {
  const parts = [
    (rowItem?.formName || ''),
    (rowItem?.FormName || ''),
    fileName || '',
    (formHeader?.title || ''),
    (formHeader?.subtitle || ''),
    (formHeader?.reference || ''),
    Array.isArray(tableHeaders) ? tableHeaders.join(' ') : ''
  ]
    .join(' ')
    .toLowerCase();
  const isFormVI = /\bform\s*[-–]?\s*vi\b|\bform\s*6\b/.test(parts);
  if (!isFormVI) return false;
  if (/national/.test(parts) && /festival/.test(parts) && /holidays?/.test(parts)) return true;
  if (/(enter\s+days\s+date|enter\s+days\s+dates|days?\s+dates?\s+and\s+months?|festival\s+holidays?)/.test(parts)) return true;
  return false;
};

const isFormVILegendLine = (text) => {
  const t = String(text || '').replace(/\s+/g, ' ').trim().toLowerCase();
  if (!t) return false;
  if (t.includes('#ref!')) return false;
  if (/^to\s+be\s+marked\s+as\s+follows\s*:?\s*$/.test(t)) return true;
  if (/^(h|w\/d|w\/h|n\/e)\s*["']?\s*for\b/.test(t)) return true;
  return false;
};

const isFormBLabourWelfareContext = (formHeader, rowItem, fileName, tableHeaders) => {
  const parts = [
    (rowItem?.formName || ''),
    (rowItem?.FormName || ''),
    fileName || '',
    (formHeader?.title || ''),
    (formHeader?.subtitle || ''),
    (formHeader?.reference || ''),
    Array.isArray(tableHeaders) ? tableHeaders.join(' ') : ''
  ]
    .join(' ')
    .toLowerCase();
  return (
    /\bform\s*["']?\s*b\b/.test(parts) &&
    /labour\s*welfare/.test(parts) &&
    /register\s+of\s+wages/.test(parts)
  );
};

const isRegisterOfWagesFormContext = (formHeader, rowItem, fileName, tableHeaders) => {
  const parts = [
    (rowItem?.formName || ''),
    (rowItem?.FormName || ''),
    fileName || '',
    (formHeader?.title || ''),
    (formHeader?.subtitle || ''),
    (formHeader?.wagePeriodText || ''),
    Array.isArray(tableHeaders) ? tableHeaders.join(' ') : ''
  ]
    .join(' ')
    .toLowerCase();
  // Form 10 / m_10 (Overtime Muster Roll) — "Number in Register" triggers `register` + wage heuristics; never use Form W thead.
  if (
    /\bform\s*no\.?\s*10\b|\bform\s+10\b|m_?10\.xlsx|m_?10|overtime\s*muster|muster\s*roll.*exempted|rule\s*78.*form\s*10/.test(
      parts
    ) &&
    !/\bform\s*w\b|form-w|form_w|register\s+of\s+wages/i.test(parts)
  ) {
    return false;
  }
  // Form B (Tamil Nadu LWF Register of Wages) should not use Form W wage-period/thead behaviors.
  if (isFormBLabourWelfareContext(formHeader, rowItem, fileName, tableHeaders)) return false;
  // Form X (leave register) must not use Form W wage-register table header logic.
  if (/\bform[_\s-]*x\b|\bform-x\b|form_x/.test(parts)) return false;
  if (parts.includes('form w') || parts.includes('form-w') || parts.includes('form_w')) return true;
  if (parts.includes('register') && /\bwages?\b/.test(parts)) return true;
  if (/wage\s*period/.test(parts) && /\bwages?\b/.test(parts)) return true;
  return false;
};

/**
 * Pick Form W table header rows immediately above the first data row. Skips a leading
 * "Wage Period …" row if it fell inside the last-3-rows window, and expands upward if needed
 * so we still get three real header tiers when possible.
 */
const adjustRegisterOfWagesHeaderBand = (worksheet, dataStartIndex, numCols) => {
  const ds = dataStartIndex;
  if (ds == null || ds < 1 || !worksheet || numCols < 1) return null;
  const merges = worksheet['!merges'] || [];
  const rawCell = (r, c) => {
    if (r < 0 || c < 0) return '';
    const ref = XLSX.utils.encode_cell({ r, c });
    const cell = worksheet[ref];
    if (cell && cell.v != null) return String(cell.v).trim();
    return '';
  };
  const mergedAware = (r, c) => {
    const d = rawCell(r, c);
    if (d) return d;
    for (let i = 0; i < merges.length; i++) {
      const m = merges[i];
      if (!m || !m.s || !m.e) continue;
      if (r >= m.s.r && r <= m.e.r && c >= m.s.c && c <= m.e.c) {
        const t = rawCell(m.s.r, m.s.c);
        if (t) return t;
      }
    }
    return '';
  };
  const rowLooksLikeWagePeriod = (r) => {
    let acc = '';
    const lim = Math.min(numCols, 28);
    for (let c = 0; c < lim; c++) acc += ` ${mergedAware(r, c)}`;
    return /wage\s*period/i.test(acc);
  };

  const r1 = ds - 1;
  let r0 = Math.max(0, ds - 3);
  while (r0 < r1 && rowLooksLikeWagePeriod(r0)) r0 += 1;
  while (r1 - r0 < 2 && r0 > 0 && !rowLooksLikeWagePeriod(r0 - 1)) r0 -= 1;
  if (r0 > r1) return null;
  return { r0, r1 };
};

/**
 * Build up to three <tr> rows for Register of Wages / Form W table headers, using the same
 * rowspan/colspan merges as the Excel sheet (Deductions band, Advances, etc.).
 */
const buildRegisterOfWagesHeaderRows = (worksheet, r0, r1, numCols) => {
  if (!worksheet || numCols < 1 || r1 < r0) return null;
  const merges = worksheet['!merges'] || [];

  const getActiveMerge = (r, c) => {
    for (let i = 0; i < merges.length; i++) {
      const m = merges[i];
      if (!m || !m.s || !m.e) continue;
      if (r >= m.s.r && r <= m.e.r && c >= m.s.c && c <= m.e.c) return m;
    }
    return { s: { r, c }, e: { r, c } };
  };

  const getTopLeftValue = (m) => {
    const ref = XLSX.utils.encode_cell({ r: m.s.r, c: m.s.c });
    const cell = worksheet[ref];
    if (cell && cell.v != null) return String(cell.v).trim();
    return '';
  };

  const bandRows = r1 - r0 + 1;
  const skipped = Array.from({ length: bandRows }, () => Array(numCols).fill(false));
  const trs = [];

  const thStyle = {
    backgroundColor: STAT_FORM_FILE_ACCENT_BG,
    color: 'white',
    padding: '8px 6px',
    textAlign: 'center',
    fontWeight: '600',
    fontSize: '11px',
    lineHeight: 1.25,
    border: `1px solid ${STAT_FORM_FILE_ACCENT_BORDER}`,
    verticalAlign: 'middle',
    whiteSpace: 'normal',
    wordBreak: 'break-word',
    minWidth: '72px'
  };

  for (let r = r0; r <= r1; r++) {
    const rr = r - r0;
    const tds = [];
    for (let c = 0; c < numCols; c++) {
      if (skipped[rr][c]) continue;
      const m = getActiveMerge(r, c);
      const intersectSr = Math.max(m.s.r, r0);
      const intersectSc = Math.max(m.s.c, 0);
      const intersectEr = Math.min(m.e.r, r1);
      const intersectEc = Math.min(m.e.c, numCols - 1);
      if (intersectSr > intersectEr || intersectSc > intersectEc) {
        tds.push(
          <th key={`h-${r}-${c}-e`} style={thStyle} rowSpan={1} colSpan={1}>
            {'\u00a0'}
          </th>
        );
        skipped[rr][c] = true;
        continue;
      }
      if (r !== intersectSr || c !== intersectSc) continue;
      const rowspan = intersectEr - intersectSr + 1;
      const colspan = intersectEc - intersectSc + 1;
      const label = getTopLeftValue(m);
      tds.push(
        <th key={`h-${intersectSr}-${intersectSc}`} rowSpan={rowspan} colSpan={colspan} style={thStyle}>
          {label || '\u00a0'}
        </th>
      );
      for (let mr = intersectSr; mr <= intersectEr; mr++) {
        for (let mc = intersectSc; mc <= intersectEc; mc++) {
          if (mr >= r0 && mr <= r1 && mc >= 0 && mc < numCols) {
            skipped[mr - r0][mc] = true;
          }
        }
      }
    }
    trs.push(<tr key={`head-row-${r}`}>{tds}</tr>);
  }
  return trs.length ? trs : null;
};

/** Proof Submission column: only PDF / images (not draft Excel or Word). */
const PROOF_SUBMISSION_DOC_EXT = /\.(pdf|png|jpe?g|gif|webp|bmp|tiff?|heic)$/i;
const isProofSubmissionDocumentFileName = (name) => PROOF_SUBMISSION_DOC_EXT.test(String(name || '').trim());
const isRowStatusApproved = (row) => {
  const rawStatus = row?.status != null && row.status !== '' ? row.status : row?.Status;
  if (String(rawStatus || '').trim().toLowerCase() === 'approved') return true;
  const rawApproval = row?.approval != null && row.approval !== '' ? row.approval : row?.Approval;
  const aNorm = String(rawApproval || '').trim().toLowerCase();
  return aNorm === 'approved' || aNorm === 'approve';
};

const AUTH_SIGNATORY_LABEL_RE =
  /\b(authorised|authorized)\s+signator(y|ies)\b/i;

function statutoryExcelCellText(val) {
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

function detectImageExtensionFromArrayBuffer(ab) {
  const u = new Uint8Array(ab.byteLength ? ab.slice(0, 12) : []);
  if (u.length >= 4 && u[0] === 0x89 && u[1] === 0x50 && u[2] === 0x4e && u[3] === 0x47) return 'png';
  if (u.length >= 3 && u[0] === 0xff && u[1] === 0xd8 && u[2] === 0xff) return 'jpeg';
  if (u.length >= 6 && u[0] === 0x47 && u[1] === 0x49 && u[2] === 0x46) return 'gif';
  if (u.length >= 12 && u[8] === 0x57 && u[9] === 0x45 && u[10] === 0x42 && u[11] === 0x50) return 'webp';
  return 'png';
}

function looksLikeJsonErrorArrayBuffer(ab) {
  if (!ab || !ab.byteLength) return true;
  const head = new TextDecoder('utf-8', { fatal: false }).decode(new Uint8Array(ab.slice(0, 120))).trim();
  return head.startsWith('{') && (head.includes('"failure"') || head.includes('"status"'));
}

function normalizeFooterSearchText(txt) {
  return String(txt || '')
    .replace(/\u00a0/g, ' ')
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

function getFooterScanCellText(worksheet, r, c) {
  let cell;
  try {
    cell = worksheet.getCell(r, c);
  } catch {
    return '';
  }
  const fromValue = statutoryExcelCellText(cell.value).replace(/\u00a0/g, ' ').trim();
  if (fromValue) return fromValue;
  try {
    const t = String(cell.text || '').replace(/\u00a0/g, ' ').trim();
    if (t) return t;
  } catch {
    /* ignore */
  }
  return '';
}

function footerRowMatchesAuthorisedSignatory(normRow) {
  if (!normRow) return false;
  if (AUTH_SIGNATORY_LABEL_RE.test(normRow)) return true;
  if (normRow.includes('authorised signatory') || normRow.includes('authorized signatory')) return true;
  if (normRow.includes('authorised') && normRow.includes('signatory')) return true;
  if (normRow.includes('authorized') && normRow.includes('signatory')) return true;
  return false;
}

function excelColumnLetterFrom1Based(col1Based) {
  let n = Math.max(1, Math.floor(col1Based));
  let s = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s || 'A';
}

function findAuthorisedSignatoryCell(worksheet) {
  const maxRow = 520;
  const maxCol = 120;
  for (let r = 1; r <= maxRow; r += 1) {
    const rowChunks = [];
    for (let c = 1; c <= maxCol; c += 1) {
      const raw = getFooterScanCellText(worksheet, r, c);
      if (!raw) continue;
      rowChunks.push(normalizeFooterSearchText(raw));
    }
    const normRow = rowChunks.join(' ').trim();
    if (!footerRowMatchesAuthorisedSignatory(normRow)) continue;
    for (let c = 1; c <= maxCol; c += 1) {
      const raw = getFooterScanCellText(worksheet, r, c);
      const norm = normalizeFooterSearchText(raw);
      if (!norm) continue;
      if (
        AUTH_SIGNATORY_LABEL_RE.test(norm) ||
        norm.includes('authorised signatory') ||
        norm.includes('authorized signatory') ||
        (norm.includes('signatory') && (norm.includes('authorised') || norm.includes('authorized')))
      ) {
        return { row: r, col: c };
      }
    }
    for (let c = 1; c <= maxCol; c += 1) {
      if (String(getFooterScanCellText(worksheet, r, c) || '').trim()) return { row: r, col: c };
    }
    return { row: r, col: 21 };
  }
  return null;
}

async function resolveCompanyIdForStatutoryTemplateDownload(item) {
  let listResp;
  try {
    listResp = await fetch('/server/company_function/company');
  } catch {
    return null;
  }
  if (!listResp || !listResp.ok) return null;
  const j = await listResp.json().catch(() => null);
  const list = j?.data?.companyDetails;
  if (!Array.isArray(list) || list.length === 0) return null;
  const target = String(item?.companyName || item?.CompanyName || '')
    .trim()
    .toLowerCase();
  if (target) {
    const match = list.find((co) => String(co.companyName || '').trim().toLowerCase() === target);
    if (match?.id != null) return match.id;
  }
  return list[0]?.id ?? null;
}

async function fetchCompanyHeadHrImageArrayBuffers(companyId) {
  if (companyId == null || companyId === '') return { sign: null, seal: null };
  const base = `/server/company_function/company/${encodeURIComponent(String(companyId))}/file/`;
  const fetchOne = async (docType) => {
    try {
      const r = await fetch(`${base}${docType}`);
      if (!r.ok) return null;
      const buf = await r.arrayBuffer();
      if (!buf || !buf.byteLength || looksLikeJsonErrorArrayBuffer(buf)) return null;
      return buf;
    } catch {
      return null;
    }
  };
  const [sign, seal] = await Promise.all([fetchOne('HeadHRSign'), fetchOne('HeadHRSeal')]);
  return { sign, seal };
}

const STATUTORY_SIGNATORY_FOOTER_FONT = { name: 'Palatino Linotype', size: 12, bold: true, color: { argb: 'FF000000' } };

function clearWorksheetCellRect(worksheet, r1, r2, c1, c2) {
  const rStart = Math.min(r1, r2);
  const rEnd = Math.max(r1, r2);
  const cStart = Math.min(c1, c2);
  const cEnd = Math.max(c1, c2);
  for (let r = rStart; r <= rEnd; r += 1) {
    for (let c = cStart; c <= cEnd; c += 1) {
      try {
        worksheet.getCell(r, c).value = null;
      } catch {
        /* ignore */
      }
    }
  }
}

/** Template often leaves "Signature of Employer / ..." on rows that overlap seal images — remove all but `keepRow1Based`. */
function removeGhostSignatureEmployerFooterText(worksheet, keepRow1Based, scanR1, scanR2, maxCol = 70) {
  const rLo = Math.max(1, Math.min(scanR1, scanR2));
  const rHi = Math.min(520, Math.max(scanR1, scanR2));
  const needle = /signature\s+of\s+employer/i;
  for (let r = rLo; r <= rHi; r += 1) {
    if (r === keepRow1Based) continue;
    for (let c = 1; c <= maxCol; c += 1) {
      try {
        const raw = getFooterScanCellText(worksheet, r, c);
        if (!raw || !needle.test(String(raw))) continue;
        worksheet.getCell(r, c).value = null;
      } catch {
        /* ignore */
      }
    }
  }
}

function applyMergedFooterTextLine(worksheet, row1Based, centerCol1Based, text, mergeHalfSpanCols) {
  const l = Math.max(1, centerCol1Based - mergeHalfSpanCols);
  const r = Math.min(90, centerCol1Based + mergeHalfSpanCols);
  const leftL = excelColumnLetterFrom1Based(l);
  const rightL = excelColumnLetterFrom1Based(r);
  const rangeAddr = `${leftL}${row1Based}:${rightL}${row1Based}`;
  try {
    worksheet.unMergeCells(rangeAddr);
  } catch {
    /* no existing merge for this exact range */
  }
  try {
    worksheet.mergeCells(rangeAddr);
  } catch {
    const cell = worksheet.getCell(row1Based, centerCol1Based);
    cell.value = text;
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: false };
    cell.font = { ...STATUTORY_SIGNATORY_FOOTER_FONT };
    return;
  }
  const cell = worksheet.getCell(row1Based, l);
  cell.value = text;
  cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: false };
  cell.font = { ...STATUTORY_SIGNATORY_FOOTER_FONT };
}

function writeStatutoryAuthorisedSignatoryFooterBlock(worksheet, sigRow1Based, sigCol1Based, footerCompanyName) {
  const companyPart =
    footerCompanyName && String(footerCompanyName).trim() ? String(footerCompanyName).trim() : 'Company Name';
  const lineForCompany = `For (${companyPart})`;
  const lineAuthorised = 'Authorised Signatory';
  const lineSignature = 'Signature of Employer / Manager / Authorised Person';
  /** Layout: For (row-3) → HeadHRSign (rows-2..-1) → Authorised (row) → HeadHRSeal (rows+1..+2) → Signature (row+3). */
  const forRow = Math.max(1, sigRow1Based - 3);
  const signatureLineRow = sigRow1Based + 3;
  applyMergedFooterTextLine(worksheet, forRow, sigCol1Based, lineForCompany, 5);
  applyMergedFooterTextLine(worksheet, sigRow1Based, sigCol1Based, lineAuthorised, 5);
  applyMergedFooterTextLine(worksheet, signatureLineRow, sigCol1Based, lineSignature, 7);
}

async function injectHeadHrSignSealAboveAuthorisedSignatoryInBlob(blob, { signBuffer, sealBuffer, footerCompanyName }) {
  const validSign = signBuffer && signBuffer.byteLength && !looksLikeJsonErrorArrayBuffer(signBuffer);
  const validSeal = sealBuffer && sealBuffer.byteLength && !looksLikeJsonErrorArrayBuffer(sealBuffer);
  if (!validSign && !validSeal) return blob;
  const arrayBuffer = await blob.arrayBuffer();
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(arrayBuffer);
  let anchorWs = null;
  let anchor = null;
  const sheets = Array.from(workbook.worksheets || []).sort((a, b) => {
    const score = (name) => (/form\s*u\b/i.test(String(name || '')) ? 0 : 1);
    return score(a.name) - score(b.name);
  });
  for (let i = 0; i < sheets.length; i += 1) {
    const ws = sheets[i];
    const found = findAuthorisedSignatoryCell(ws);
    if (found) {
      anchorWs = ws;
      anchor = found;
      break;
    }
  }
  if (!anchorWs || !anchor) return blob;

  const { row: sigRow, col: sigCol } = anchor;
  const sigCol0 = sigCol - 1;
  const signImageTop1Based = Math.max(1, sigRow - 2);
  const signImageBot1Based = signImageTop1Based + 1;
  const sealImageTop1Based = sigRow + 1;
  const sealImageBot1Based = sealImageTop1Based + 1;
  /** ~half prior width: anchor across two columns centered on sigCol (was five). */
  const imgBandLeft1 = Math.max(1, sigCol - 1);
  const imgBandRight1 = Math.min(90, imgBandLeft1 + 1);
  [signImageTop1Based, signImageBot1Based, sealImageTop1Based, sealImageBot1Based].forEach((rn) => {
    const row = anchorWs.getRow(rn);
    row.height = Math.max(Number(row.height) || 0, 56);
  });

  const forRow = Math.max(1, sigRow - 3);
  const signatureLineRow = sigRow + 3;
  const clearC1 = Math.max(1, sigCol - 14);
  const clearC2 = Math.min(90, sigCol + 14);
  clearWorksheetCellRect(anchorWs, signImageTop1Based, signImageBot1Based, clearC1, clearC2);
  clearWorksheetCellRect(anchorWs, sealImageTop1Based, sealImageBot1Based, clearC1, clearC2);
  removeGhostSignatureEmployerFooterText(anchorWs, signatureLineRow, forRow - 2, signatureLineRow + 4, clearC2);

  const pushImage = (buf, rangeStr, tlFallback) => {
    if (!buf || !buf.byteLength || looksLikeJsonErrorArrayBuffer(buf)) return;
    const ext = detectImageExtensionFromArrayBuffer(buf);
    if (ext === 'webp') {
      console.warn('Statutory template: WebP Head HR image is not supported for Excel embedding; skip.');
      return;
    }
    const uint = buf instanceof Uint8Array ? buf : new Uint8Array(buf);
    let imageId;
    try {
      imageId = workbook.addImage({ buffer: uint, extension: ext });
    } catch (e) {
      console.warn('Statutory template: addImage buffer rejected:', e);
      return;
    }
    try {
      if (rangeStr) {
        anchorWs.addImage(imageId, rangeStr);
        return;
      }
    } catch (e) {
      console.warn('Statutory template: addImage range anchor failed, trying tl/ext:', e);
    }
    if (tlFallback) {
      try {
        anchorWs.addImage(imageId, tlFallback);
      } catch (e2) {
        console.warn('Statutory template: addImage tl/ext anchor failed:', e2);
      }
    }
  };

  const signRange = `${excelColumnLetterFrom1Based(imgBandLeft1)}${signImageTop1Based}:${excelColumnLetterFrom1Based(imgBandRight1)}${signImageBot1Based}`;
  const sealRange = `${excelColumnLetterFrom1Based(imgBandLeft1)}${sealImageTop1Based}:${excelColumnLetterFrom1Based(imgBandRight1)}${sealImageBot1Based}`;
  const signTlRow0 = signImageTop1Based - 1;
  const sealTlRow0 = sealImageTop1Based - 1;
  const signTl = {
    tl: { col: Math.max(0, sigCol0 - 1) + 0.05, row: signTlRow0 + 0.04 },
    ext: { width: 90, height: 56 }
  };
  const sealTl = {
    tl: { col: Math.max(0, sigCol0 - 1) + 0.05, row: sealTlRow0 + 0.04 },
    ext: { width: 90, height: 56 }
  };

  pushImage(validSign ? signBuffer : null, signRange, signTl);
  pushImage(validSeal ? sealBuffer : null, sealRange, sealTl);

  try {
    writeStatutoryAuthorisedSignatoryFooterBlock(anchorWs, sigRow, sigCol, footerCompanyName);
  } catch (e) {
    console.warn('Statutory template: signatory footer text write failed:', e);
  }
  removeGhostSignatureEmployerFooterText(anchorWs, signatureLineRow, forRow - 2, signatureLineRow + 4, clearC2);

  const outBuf = await workbook.xlsx.writeBuffer();
  return new Blob([outBuf], { type: blob.type || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}

async function appendApprovedStatutoryHeadHrSignSealToDownloadBlob(blob, item) {
  if (!isRowStatusApproved(item)) return blob;
  try {
    const companyId = await resolveCompanyIdForStatutoryTemplateDownload(item);
    const { sign, seal } = await fetchCompanyHeadHrImageArrayBuffers(companyId);
    const footerCompanyName =
      String(item?.companyName || item?.CompanyName || '').trim() || null;
    return await injectHeadHrSignSealAboveAuthorisedSignatoryInBlob(blob, {
      signBuffer: sign,
      sealBuffer: seal,
      footerCompanyName
    });
  } catch (err) {
    console.warn('Statutory template: Head HR sign/seal injection failed:', err);
    return blob;
  }
}

/** Latest activity time for table ordering (new saves/submissions surface first). */
const getStatutoryRowActivityTimeMs = (row) => {
  const raw =
    row?.modifiedTime ??
    row?.ModifiedTime ??
    row?.modified_time ??
    row?.createdTime ??
    row?.CreatedTime ??
    row?.created_time ??
    '';
  const parsed = Date.parse(String(raw || '').trim());
  if (Number.isFinite(parsed)) return parsed;
  const idStr = String(row?.id ?? '').trim();
  if (/^\d+$/.test(idStr)) {
    const n = Number.parseInt(idStr, 10);
    return Number.isFinite(n) ? n : 0;
  }
  return 0;
};

/** Draft saved (Autofill → Save) or workflow sent — surface these rows before lines still empty. */
const hasStatutoryDraftSubmittedForSort = (row) => {
  const draftVal = row?.draftFile ?? row?.DraftFile ?? null;
  const hasDraftFile =
    draftVal != null &&
    String(draftVal).trim() !== '' &&
    String(draftVal).trim().toLowerCase() !== 'null' &&
    String(draftVal).trim().toLowerCase() !== 'undefined';
  if (hasDraftFile) return true;
  const sfa = String(row?.sendForApproval ?? row?.SendForApproval ?? '').trim();
  return /^sent$/i.test(sfa);
};

const initialForm = {
  act: '',
  description: '',
  formFile: null,
  formFileName: null,
  formName: '',
  dueDate: '',
  sector: '',
  state: '',
  site: '',
  autofill: '',
  draft: '',
  proofSubmissionFile: null,
  proofSubmissionFileName: null,
  draftFile: null,
  draftFileName: null,
  approval: '',
  status: '',
  sendForApproval: '',
  remarks: '',
  originalMonthFilter: null
};

// For afrindinu14: site name -> act category (same as Audit page)
const getSiteWiseCategory = (siteName) => {
  const s = String(siteName || '').trim().toLowerCase();
  if (s === 'delphi kakinada') return 'clra';
  if (s === 'delphi') return 'shops_and_establishment';
  if (s === 'delphi oragadam') return 'factories';
  return null;
};

const Statutory = ({ userEmail, userRole }) => {
  const [searchParams] = useSearchParams();
  const siteFromUrl = searchParams.get('site') || '';
  const effectiveUserEmail = (userEmail || (typeof localStorage !== 'undefined' ? localStorage.getItem('userEmail') : null) || '').trim().toLowerCase();
  const isSiteWiseStatutoryUser = effectiveUserEmail === SITE_WISE_STATUTORY_EMAIL;
  const siteWiseCategory = isSiteWiseStatutoryUser && siteFromUrl ? getSiteWiseCategory(siteFromUrl) : null;
  /** Act categories from Site Management (Incharge email + Industry); drives Statutory row filter instead of hardcoded emails */
  const [allowedActCategoryList, setAllowedActCategoryList] = useState(null);
  const [siteNamesByActCategory, setSiteNamesByActCategory] = useState({});
  const [allowedSiteNameList, setAllowedSiteNameList] = useState(null);
  /** Incharge sites with SiteState + industry bucket — drives Statutory Site column when row has `state`. */
  const [inchargeSitesMeta, setInchargeSitesMeta] = useState([]);
  /** All org sites (Site Management) with SiteState — used for Site column when user has no Incharge rows. */
  const [organizationSitesMeta, setOrganizationSitesMeta] = useState([]);
  /** Site `SiteState` values from Site Management for rows where Incharge = login (filter Statutory `state`). */
  const [allowedInchargeStateLabels, setAllowedInchargeStateLabels] = useState(null);
  /** After first `fetchStatutoryData` run finishes (incl. silent mount refresh), Site column visibility can use scope lists — avoids cache-first paint before Site Management meta arrives. */
  const [transactionSiteMetaReady, setTransactionSiteMetaReady] = useState(false);
  const hasSiteBasedActScope = Array.isArray(allowedActCategoryList) && allowedActCategoryList.length > 0;

  const [form, setForm] = useState(initialForm);
  const [statutoryData, setStatutoryData] = useState([]);
  const [formmasterTemplates, setFormmasterTemplates] = useState([]); // Formmaster file list for Form File column lookup by form name
  const [loading, setLoading] = useState(false);
  const [submitting, setSubmitting] = useState(false);
  const [error, setError] = useState('');
  const [success, setSuccess] = useState('');
  const [isFormOpen, setIsFormOpen] = useState(false);
  const [editingStatutoryId, setEditingStatutoryId] = useState(null);
  const [selectedRows, setSelectedRows] = useState(new Set());
  const [proofUploadingRowId, setProofUploadingRowId] = useState(null);
  const [approvalUpdatingRowId, setApprovalUpdatingRowId] = useState(null);
  const [sendForApprovalUpdatingRowId, setSendForApprovalUpdatingRowId] = useState(null);
  const [isRejectRemarksModalOpen, setIsRejectRemarksModalOpen] = useState(false);
  const [rejectRemarksValue, setRejectRemarksValue] = useState('');
  const [rejectRemarksTarget, setRejectRemarksTarget] = useState(null);
  const [lastSyncedAt, setLastSyncedAt] = useState(null);
  const fileInputRef = useRef(null);
  const isFetchingStatutoryRef = useRef(false);
  const [isFormFileModalOpen, setIsFormFileModalOpen] = useState(false);
  const [formFileModalData, setFormFileModalData] = useState(null);
  const [formFileLoading, setFormFileLoading] = useState(false);
  const [formHeader, setFormHeader] = useState(null);
  const [headerFormData, setHeaderFormData] = useState({});
  const [formTableData, setFormTableData] = useState([]);
  const [tableHeaders, setTableHeaders] = useState([]);
  const [subColumns, setSubColumns] = useState(null);
  const [formFields, setFormFields] = useState([]);
  const [isAutofillMode, setIsAutofillMode] = useState(false);
  const [autofillItem, setAutofillItem] = useState(null);
  const formModalImportInputRef = useRef(null);
  const formTableDataRef = useRef([]);
 
  const getCurrentMonth = () => MONTH_NAMES[new Date().getMonth()];
  const getInitialMonth = () => {
    try {
      const saved = typeof localStorage !== 'undefined' ? localStorage.getItem(STATUTORY_MONTH_FILTER_KEY) : null;
      if (saved && String(saved).trim() !== '') return String(saved).trim();
    } catch (_) {}
    return getCurrentMonth();
  };
  const [selectedMonth, setSelectedMonth] = useState(getInitialMonth);
  const [tableSearch, setTableSearch] = useState('');
  const [selectedFormFilter, setSelectedFormFilter] = useState('all');
  const [currentPage, setCurrentPage] = useState(1);
  const PAGE_SIZE = 10;
  useEffect(() => {
    formTableDataRef.current = Array.isArray(formTableData) ? formTableData : [];
  }, [formTableData]);
  const persistMonthFilter = (month) => {
    try {
      if (typeof localStorage !== 'undefined' && month) localStorage.setItem(STATUTORY_MONTH_FILTER_KEY, month);
    } catch (_) {}
  };

  // Function to format monthly basis due dates
  // If due date is "Monthly Basis" and a month is selected, show as "15-[Month Name]"
  const formatMonthlyBasisDate = (dueDate) => {
    if (!dueDate) return '';
   
    const dueDateStr = String(dueDate).trim();
    const dueDateLower = dueDateStr.toLowerCase();
   
    // Check if it's "Monthly Basis" and a month is selected
    if (dueDateLower.includes('monthly basis') && selectedMonth && selectedMonth.trim() !== '') {
      // Get the month abbreviation
      const monthNames = ['January', 'February', 'March', 'April', 'May', 'June',
                          'July', 'August', 'September', 'October', 'November', 'December'];
      const monthAbbr = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
                         'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
     
      const selectedMonthLower = selectedMonth.toLowerCase();
      const monthIndex = monthNames.findIndex(m => m.toLowerCase().startsWith(selectedMonthLower));
     
      if (monthIndex !== -1) {
        // Return "15-[Month Abbr]" format (e.g., "15-Dec", "15-Nov")
        return `15-${monthAbbr[monthIndex]}`;
      }
     
      // Fallback: return "15-[Selected Month]" if month not found in list
      return `15-${selectedMonth.substring(0, 3)}`;
    }
   
    // Return the due date as-is for non-monthly basis items
    return dueDateStr;
  };

  // Collapse checklist bulk placeholders vs Statutory DB rows after save (canonical month + site + numeric ROW priority).
  const removeDuplicateStatutoryRecords = (mergedRows, resolveSiteFn, uiMonthFallback) => {
    if (!mergedRows || mergedRows.length === 0) {
      return mergedRows;
    }

    const duplicateMap = new Map();

    mergedRows.forEach((item) => {
      const normalizedFormName = squashStatutoryKeyPart(item.formName || item.FormName);
      const normalizedAct = squashStatutoryKeyPart(item.act || item.Act);
      const normalizedDescription = squashStatutoryKeyPart(item.description || item.Description);
      const normalizedSector = squashStatutoryKeyPart(item.sector || item.Sector) || 'nosector';
      const normalizedState = squashStatutoryKeyPart(item.state || item.State) || 'nostate';
      const monthNorm = statutoryDedupeMonthNorm(item, uiMonthFallback);
      const siteKey = statutoryCollapseSiteKey(item, mergedRows, resolveSiteFn);
      const key = `${normalizedFormName}_${normalizedAct}_${normalizedDescription}_${monthNorm}_${normalizedSector || 'nosector'}_${normalizedState || 'nostate'}_${siteKey}`;

      if (!duplicateMap.has(key)) {
        duplicateMap.set(key, []);
      }
      duplicateMap.get(key).push(item);
    });

    const deduplicated = [];

    duplicateMap.forEach((group) => {
      if (group.length === 1) {
        deduplicated.push(group[0]);
        return;
      }
      const best = [...group].sort((a, b) => {
        const sc = statutoryRowRichnessScore(b) - statutoryRowRichnessScore(a);
        if (sc !== 0) return sc;
        return String(b.id || '').localeCompare(String(a.id || ''));
      })[0];
      deduplicated.push(best);
      console.log(`🔍 Removed ${group.length - 1} duplicate(s) for "${group[0].formName}" — kept record ID ${best.id}`);
    });

    return deduplicated;
  };

  // Convert Formmaster template list to statutory row shape; include formFile/formFileName so Form File column shows link
  const formmasterTemplatesToRows = (templates) => {
    if (!templates || !Array.isArray(templates)) return [];
    const timestamp = Date.now();
    return templates.map((t, index) => {
      const rawName = (t.name || t.file_name || '').trim();
      const formName = rawName.replace(/\.xlsx$/i, '').trim() || rawName;
      const fileId = t.id != null ? String(t.id) : null;
      const fileName = rawName || 'template.xlsx';
      return {
        id: `formmaster_${fileId ?? timestamp}_${index}`,
        formName,
        sector: '',
        state: '',
        act: '',
        dueDate: '',
        description: '',
        concernedGovtDepartment: '',
        formFile: fileId,
        formFileName: fileName,
        autofill: '',
        draft: '',
        proofSubmissionFile: null,
        proofSubmissionFileName: null,
        isFromFormmaster: true
      };
    });
  };

  // Function to merge ChecklistBulk data with Statutory data
  // All ChecklistBulk rows are displayed in Statutory (no filtering by form name)
  const mergeBulkDataWithStatutory = (statutoryData, bulkData) => {
    if (!bulkData || bulkData.length === 0) {
      return statutoryData;
    }

    // Convert all bulk data to statutory format – every ChecklistBulk row is included
    const timestamp = Date.now();
    const convertedBulkData = bulkData.map((bulkItem, index) => ({
      id: `bulk_${bulkItem.id ?? timestamp}_${index}`,
      formName: bulkItem.formName || '',
      sector: bulkItem.sector || bulkItem.Sector || '',
      state: bulkItem.state || bulkItem.State || '',
      act: bulkItem.act || '',
      dueDate: bulkItem.dueDate || '',
      description: bulkItem.description || '',
      concernedGovtDepartment: bulkItem.concernedGovtDepartment || '',
      formFile: null,
      formFileName: null,
      autofill: '',
      draft: '',
      proofSubmissionFile: null,
      proofSubmissionFileName: null,
      isBulkImported: true
    }));

    // Merge: statutory data first, then ALL ChecklistBulk rows so every bulk record is displayed
    const merged = [...statutoryData, ...convertedBulkData];
    console.log(`✅ mergeBulkDataWithStatutory: ${merged.length} total entries (${statutoryData.length} statutory + ${convertedBulkData.length} from ChecklistBulk – all bulk data shown)`);
    return merged;
  };

  // Fill missing meta fields (act/description/dueDate) from bulk entries
  const enrichWithBulkMetadata = (data, bulkData) => {
    if (!bulkData || bulkData.length === 0) return data;

    const bulkMap = new Map();
    bulkData.forEach((bulkItem) => {
      const fn = String(bulkItem.formName || '').toLowerCase().trim();
      if (!fn) return;
      const sec = String(bulkItem.sector || '').trim().toLowerCase();
      bulkMap.set(`${fn}|${sec}`, bulkItem);
    });

    const pick = (primary, fallback) => {
      const primaryTrimmed = String(primary || '').trim();
      if (primaryTrimmed) return primary;
      const fallbackTrimmed = String(fallback || '').trim();
      return fallbackTrimmed ? fallback : primary;
    };

    return data.map((item) => {
      const fn = String(item.formName || '').toLowerCase().trim();
      if (!fn) return item;
      const sec = String(item.sector || item.Sector || '').trim().toLowerCase();
      let bulkItem = bulkMap.get(`${fn}|${sec}`);
      if (!bulkItem && !sec) {
        bulkItem = bulkData.find((b) => String(b.formName || '').toLowerCase().trim() === fn && !String(b.sector || '').trim());
      }
      if (!bulkItem) return item;

      return {
        ...item,
        act: pick(item.act ?? item.Act, bulkItem.act),
        description: pick(item.description ?? item.Description, bulkItem.description),
        dueDate: pick(item.dueDate ?? item.DueDate, bulkItem.dueDate),
        sector: pick(item.sector ?? item.Sector, bulkItem.sector ?? bulkItem.Sector),
        state: pick(item.state ?? item.State, bulkItem.state ?? bulkItem.State)
      };
    });
  };

  // Fetch statutory data
  const fetchStatutoryData = async (options = {}) => {
    const { useCacheFirst = false, silentRefresh = false, force = false } = options;

    if (isFetchingStatutoryRef.current && !force) {
      console.log('⏭️ Skipping duplicate statutory fetch (already in progress)');
      return;
    }
    isFetchingStatutoryRef.current = true;

    if (useCacheFirst) {
      try {
        const cached = localStorage.getItem('statutoryData');
        if (cached) {
          const parsed = JSON.parse(cached);
          if (Array.isArray(parsed) && parsed.length > 0) {
            setStatutoryData(parsed);
          }
        }
      } catch (_) {}
    }

    if (!silentRefresh) setLoading(true);
    setError('');
    setSuccess('');
    try {
      const siteScopePromise = fetchAllowedActCategoriesFromSites(userEmail)
        .then((sc) => sc)
        .catch((scopeErr) => {
          console.warn('Could not load statutory scope from Site Management:', scopeErr);
          return null;
        });
      const siteNamesPromise = (async () => {
        const loginResolved = normalizeEmail(await resolveLoginEmailString(userEmail));
        try {
          const resp = await fetch('/server/sitemanagement_function/sitemanagement', { cache: 'no-store' });
          if (!resp.ok)
            return {
              byCategory: {},
              allowedNames: null,
              inchargeStates: null,
              inchargeSites: [],
              organizationSites: []
            };
          const json = await resp.json().catch(() => ({}));
          const details = Array.isArray(json?.data?.siteDetails) ? json.data.siteDetails : [];
          const buckets = {
            factories: new Set(),
            shops_and_establishment: new Set(),
            clra: new Set()
          };
          const allowedNames = new Set();
          const inchargeStates = new Set();
          const inchargeSites = [];
          const organizationSites = [];
          details.forEach((s) => {
            const industry = String(s?.industry ?? s?.Industry ?? '').trim();
            const siteName = String(s?.siteName ?? s?.SiteName ?? '').trim();
            const st = String(s?.siteState ?? s?.SiteState ?? s?.state ?? s?.State ?? '').trim();
            if (siteName && st) {
              organizationSites.push({
                siteName,
                siteState: st,
                industry,
                actCategory: industryLabelToActCategory(industry)
              });
            }
            const inchargeEmail = normalizeEmail(
              String(s?.inchargeEmail ?? s?.InchargeEmail ?? s?.incharge_email ?? '')
            );
            if (!inchargeEmail || inchargeEmail !== loginResolved) return;
            if (siteName) allowedNames.add(siteName);
            if (st) inchargeStates.add(st);
            if (siteName) {
              inchargeSites.push({
                siteName,
                siteState: st,
                industry,
                actCategory: industryLabelToActCategory(industry)
              });
            }
            if (!industry || !siteName) return;
            const cat = industryLabelToActCategory(industry);
            if (!cat || !buckets[cat]) return;
            buckets[cat].add(siteName);
          });
          return {
            byCategory: {
              factories: Array.from(buckets.factories),
              shops_and_establishment: Array.from(buckets.shops_and_establishment),
              clra: Array.from(buckets.clra)
            },
            allowedNames: allowedNames.size > 0 ? Array.from(allowedNames) : null,
            inchargeStates: inchargeStates.size > 0 ? Array.from(inchargeStates) : null,
            inchargeSites,
            organizationSites
          };
        } catch {
          return { byCategory: {}, allowedNames: null, inchargeStates: null, inchargeSites: [], organizationSites: [] };
        }
      })();

      const response = await fetch('/server/statutoryreg_function/statutory', {
        cache: 'no-store'
      });
      let baseStatutoryData = [];
      let parsedBulkDataCache = null;
      let siteScopeCategories = null;

      if (response.ok) {
        const data = await response.json();
        const siteMeta = await siteNamesPromise;
        setSiteNamesByActCategory(siteMeta?.byCategory || {});
        setAllowedSiteNameList(siteMeta?.allowedNames || null);
        setInchargeSitesMeta(Array.isArray(siteMeta?.inchargeSites) ? siteMeta.inchargeSites : []);
        setOrganizationSitesMeta(Array.isArray(siteMeta?.organizationSites) ? siteMeta.organizationSites : []);
        setAllowedInchargeStateLabels(siteMeta?.inchargeStates ?? null);
        if (data.status === 'success' && data.data && data.data.statutoryData) {
          baseStatutoryData = data.data.statutoryData;
          // Normalize field names to handle both camelCase and PascalCase
          baseStatutoryData = baseStatutoryData.map(item => ({
            ...item,
            formFile: item.formFile || item.FormFile || null,
            formFileName: item.formFileName || item.FormFileName || null,
            sampleFile: item.sampleFile || item.SampleFile || item.formFile || item.FormFile || null,
            sampleFileName: item.sampleFileName || item.SampleFileName || item.formFileName || item.FormFileName || null,
            formName: item.formName || item.FormName || '',
            act: item.act || item.Act || '',
            sector: item.sector || item.Sector || '',
            state: item.state || item.State || '',
            site: item.site || item.Site || item.siteName || item.SiteName || '',
            description: item.description || item.Description || '',
            dueDate: item.dueDate || item.DueDate || '',
            monthFilter: item.monthFilter || item.MonthFilter || item.monthfilter || null,
            autofill: item.autofill || item.Autofill || '',
            draft: item.draft || item.Draft || '',
            proofSubmissionFile: item.proofSubmissionFile || item.ProofSubmissionFile || null,
            proofSubmissionFileName: item.proofSubmissionFileName || item.ProofSubmissionFileName || null,
            draftFile: item.draftFile || item.DraftFile || null,
            draftFileName: item.draftFileName || item.DraftFileName || null,
            approval: item.approval || item.Approval || '',
            status: item.status != null && item.status !== '' ? item.status : item.Status != null ? item.Status : '',
            sendForApproval: item.sendForApproval || item.SendForApproval || '',
            remarks: item.remarks || item.Remarks || ''
          }));

          // Keep rendering consistent: avoid partial first paint; show final merged dataset once ready.
          siteScopeCategories = await siteScopePromise;
          setAllowedActCategoryList(siteScopeCategories);
         
          // Log entries with files
          const entriesWithFiles = baseStatutoryData.filter(item =>
            item.formFile && item.formFile !== 'null' && String(item.formFile).trim() !== ''
          );
          console.log(`📋 Fetched ${baseStatutoryData.length} statutory items from backend`);
          console.log(`📁 ${entriesWithFiles.length} entries have form files:`, entriesWithFiles.map(item => ({
            id: item.id,
            formName: item.formName,
            formFile: item.formFile,
            formFileName: item.formFileName
          })));
         
          // Specifically check for Form U
          const formUFromDB = baseStatutoryData.filter(item => {
            const formName = (item.formName || '').toLowerCase();
            return formName.includes('form u') || formName === 'form u';
          });
          if (formUFromDB.length > 0) {
            console.log(`🔍 Form U entries from database:`, formUFromDB.map(item => ({
              id: item.id,
              formName: item.formName,
              formFile: item.formFile,
              formFileName: item.formFileName,
              hasFile: !!(item.formFile && item.formFile !== 'null' && String(item.formFile).trim() !== '')
            })));
          }
        } else {
          const localData = localStorage.getItem('statutoryData');
          if (localData) {
            baseStatutoryData = JSON.parse(localData);
          }
        }
      } else {
        throw new Error('Failed to fetch statutory data');
      }

      // Fetch Checklist data (where SEMaster saves files)
      let checklistDataWithFiles = [];
      try {
        const checklistResponse = await fetch('/server/checklist_function/checklist', {
          cache: 'no-store'
        });
        if (checklistResponse.ok) {
          const checklistResult = await checklistResponse.json();
          if (checklistResult.status === 'success' && checklistResult.data && checklistResult.data.checklistData) {
            checklistDataWithFiles = checklistResult.data.checklistData;
            console.log(`📋 Fetched ${checklistDataWithFiles.length} items from Checklist`);
           
            // Filter to only items with form files
            const checklistItemsWithFiles = checklistDataWithFiles.filter(item =>
              item.formFile &&
              item.formFile !== 'null' &&
              item.formFile !== null &&
              String(item.formFile).trim() !== '' &&
              !String(item.formFile).startsWith('STATUTORY_MASTER_FORM:') // Exclude StatutoryMaster markers
            );
            console.log(`📁 ${checklistItemsWithFiles.length} Checklist items have form files`);
           
            // Specifically log Form I entries from Checklist
            const checklistFormIEntries = checklistItemsWithFiles.filter(item => {
              const formName = String(item.formName || '').toLowerCase().trim();
              return formName === 'form i';
            });
            console.log(`📋 Found ${checklistFormIEntries.length} Form I entries in Checklist with files:`, checklistFormIEntries.map(item => ({
              id: item.id,
              formName: item.formName,
              formFile: item.formFile,
              formFileName: item.formFileName,
              act: item.act,
              description: item.description
            })));
           
            console.log(`📁 All Checklist items with files:`, checklistItemsWithFiles.map(item => ({
              id: item.id,
              formName: item.formName,
              formFile: item.formFile,
              formFileName: item.formFileName
            })));
           
            // Merge Checklist items with files into statutory data
            // Create a map of statutory data by form name for quick lookup
            // Note: This map only stores ONE entry per form name for lookup purposes
            // but ALL entries remain in baseStatutoryData
            const normalizeFormName = (name) => {
              return String(name || '')
                .toLowerCase()
                .replace(/\s*-\s*/g, ' ')
                .replace(/\s+/g, ' ')
                .trim();
            };

            const extractFormId = (name) => {
              return normalizeFormName(name).replace(/^form\s*/i, '').trim();
            };
            /** Base form id for matching (e.g. "Form U" and "Form U - Register" both -> "u") */
            const getBaseFormId = (name) => {
              const id = extractFormId(name);
              const first = (id || '').split(/\s+/)[0];
              return (first || id || '').toLowerCase();
            };

            const statutoryByFormName = new Map();
            baseStatutoryData.forEach(item => {
              const normalizedName = normalizeFormName(item.formName);
              if (normalizedName && !statutoryByFormName.has(normalizedName)) {
                // Only set if not already in map (first entry wins for lookup)
                statutoryByFormName.set(normalizedName, item);
              }
            });
           
            console.log(`📊 Before Checklist merge: ${baseStatutoryData.length} entries in baseStatutoryData`);

            const hasValidFormFile = (item) => {
              const formFile = item?.formFile;
              return !!(formFile && formFile !== 'null' && String(formFile).trim() !== '');
            };
           
            // Add all Checklist entries with files to statutory data
            // Each Checklist entry becomes a separate row, even if they have the same form name (e.g., 3 Form I entries)
            // This ensures that if SEMaster imports Form I 3 times, all 3 will display as separate rows
            checklistItemsWithFiles.forEach(checklistItem => {
              const normalizedFormName = normalizeFormName(checklistItem.formName);
              if (!normalizedFormName) return;
             
              // Check if this exact Checklist entry already exists in statutory data (by checklistId)
              const existingChecklistEntry = baseStatutoryData.find(
                item => item.checklistId === checklistItem.id && item.isFromChecklist
              );
             
              if (existingChecklistEntry) {
                // Update existing Checklist entry if it's already there
                existingChecklistEntry.formFile = checklistItem.formFile;
                existingChecklistEntry.formFileName = checklistItem.formFileName;
                existingChecklistEntry.act = checklistItem.act || existingChecklistEntry.act || '';
                existingChecklistEntry.description = checklistItem.description || existingChecklistEntry.description || '';
                existingChecklistEntry.dueDate = checklistItem.dueDate || existingChecklistEntry.dueDate || '';
                existingChecklistEntry.sector = checklistItem.sector || existingChecklistEntry.sector || '';
                existingChecklistEntry.state = checklistItem.state || existingChecklistEntry.state || '';
                existingChecklistEntry.isFromChecklist = true;
                existingChecklistEntry.checklistId = checklistItem.id;
                console.log(`✅ Updated existing Checklist entry "${checklistItem.formName}" (ID: ${checklistItem.id})`);
                return;
              } else {
                // If a base statutory row exists with the same form name (or same base form e.g. "Form U" vs "Form U - Register"), always update that same row with the checklist file (do not add a new row).
                // This prevents duplicate rows when user imports e.g. Form U in SEMaster: Statutory already has Form U; show the imported file in that same row.
                let baseStatutoryMatch = statutoryByFormName.get(normalizedFormName);
                if (!baseStatutoryMatch) {
                  const checklistFormId = extractFormId(checklistItem.formName);
                  const checklistBaseId = getBaseFormId(checklistItem.formName);
                  if (checklistFormId) {
                    baseStatutoryMatch = baseStatutoryData.find(item => {
                      const itemFormId = extractFormId(item.formName);
                      if (itemFormId && itemFormId === checklistFormId) return true;
                      // Match by base form id so "Form U" matches "Form U - Register" (both base "u")
                      return checklistBaseId && getBaseFormId(item.formName) === checklistBaseId;
                    });
                  }
                }
                if (baseStatutoryMatch) {
                  baseStatutoryMatch.formFile = checklistItem.formFile;
                  baseStatutoryMatch.formFileName = checklistItem.formFileName;
                  baseStatutoryMatch.act = checklistItem.act || baseStatutoryMatch.act || '';
                  baseStatutoryMatch.description = checklistItem.description || baseStatutoryMatch.description || '';
                  baseStatutoryMatch.dueDate = checklistItem.dueDate || baseStatutoryMatch.dueDate || '';
                  baseStatutoryMatch.sector = checklistItem.sector || baseStatutoryMatch.sector || '';
                  baseStatutoryMatch.state = checklistItem.state || baseStatutoryMatch.state || '';
                  baseStatutoryMatch.isFromChecklist = true;
                  baseStatutoryMatch.checklistId = checklistItem.id;
                  console.log(`🔗 Updated existing statutory row "${baseStatutoryMatch.formName}" with Checklist file (Checklist ID: ${checklistItem.id}) - no duplicate row`);
                  return;
                }

                // Only add a new row when there is no existing statutory row for this form name (e.g. form exists only in Checklist)
                // Otherwise add Checklist entries as new separate rows
                // This ensures all Checklist entries with files are preserved, even if they share the same form name
                // Each Checklist entry gets its own row (e.g., 3 Form I entries = 3 rows)
                const newEntry = {
                  id: `checklist_${checklistItem.id}`,
                  formName: checklistItem.formName || '',
                  act: checklistItem.act || '',
                  description: checklistItem.description || '',
                  dueDate: checklistItem.dueDate || '',
                  formFile: checklistItem.formFile,
                  formFileName: checklistItem.formFileName,
                  autofill: '',
                  draft: '',
                  proofSubmissionFile: checklistItem.proofSubmissionFile || null,
                  proofSubmissionFileName: checklistItem.proofSubmissionFileName || null,
                  sector: checklistItem.sector || '',
                  state: checklistItem.state || '',
                  isFromChecklist: true,
                  checklistId: checklistItem.id // Store original Checklist ID for file download
                };
                baseStatutoryData.push(newEntry);
                // Only set in map if it's the first entry with this form name (for lookup purposes only, doesn't prevent duplicates)
                if (!statutoryByFormName.has(normalizedFormName)) {
                  statutoryByFormName.set(normalizedFormName, newEntry);
                }
                console.log(`➕ Added new Checklist entry "${checklistItem.formName}" with file (ID: ${checklistItem.id})`);
              }
            });
           
            // Log Form I entries after merge for debugging
            const formIEntries = baseStatutoryData.filter(item => {
              const formName = String(item.formName || '').toLowerCase().trim();
              return formName === 'form i';
            });
            if (formIEntries.length > 0) {
              console.log(`📋 Found ${formIEntries.length} Form I entry/entries after Checklist merge:`, formIEntries.map(item => ({
                id: item.id,
                checklistId: item.checklistId,
                formName: item.formName,
                formFileName: item.formFileName,
                isFromChecklist: item.isFromChecklist || false,
                act: item.act,
                description: item.description
              })));
            }
           
            // Log total count after Checklist merge
            console.log(`📊 Total entries after Checklist merge: ${baseStatutoryData.length}`);
          }
        }
      } catch (checklistError) {
        console.error('Error fetching Checklist data:', checklistError);
      }

      // Load form names from Formmaster (not ChecklistBulk/SEMaster) – one row per template
      let mergedData = [...baseStatutoryData];
      console.log(`📊 Before Formmaster merge: ${mergedData.length} entries`);

      let formmasterTemplates = [];
      try {
        const fmRes = await fetch(FORMMASTER_TEMPLATES_API, { cache: 'no-store' });
        if (fmRes.ok) {
          const fmData = await fmRes.json();
          if (fmData.status === 'success' && Array.isArray(fmData.data)) {
            formmasterTemplates = fmData.data;
            console.log(`📋 Fetched ${formmasterTemplates.length} templates from Formmaster`);
          }
        }
      } catch (fmErr) {
        console.error('Error fetching Formmaster templates:', fmErr);
      }

      if (formmasterTemplates.length >= 0) {
        setFormmasterTemplates(formmasterTemplates || []);
      }
      if (formmasterTemplates.length > 0) {
        const existingFormNames = new Set(
          baseStatutoryData.map((item) => (item.formName || '').trim().toLowerCase()).filter(Boolean)
        );
        const formmasterRows = formmasterTemplatesToRows(formmasterTemplates).filter(
          (row) => !existingFormNames.has((row.formName || '').trim().toLowerCase())
        );
        if (formmasterRows.length > 0) {
          mergedData = [...baseStatutoryData, ...formmasterRows];
          console.log(`✅ Merged ${mergedData.length} total statutory items (${baseStatutoryData.length} statutory + ${formmasterRows.length} from Formmaster with form file links)`);
        }
      }

      // Fetch ChecklistBulk (all bulk records – e.g. 13 Shops and Establishment)
      let bulkFromApi = null;
      try {
        const bulkResponse = await fetch('/server/checklistbulk_function/checklistbulk?action=getAll', { cache: 'no-store' });
        if (bulkResponse.ok) {
          const bulkResult = await bulkResponse.json();
          if (bulkResult.status === 'success' && bulkResult.data && Array.isArray(bulkResult.data)) {
            bulkFromApi = bulkResult.data;
            localStorage.setItem('checklistBulkData', JSON.stringify(bulkFromApi));
          }
        }
      } catch (bulkError) {
        console.error('Error fetching ChecklistBulk data:', bulkError);
      }
      const bulkToMerge = bulkFromApi ?? (() => {
        try {
          const stored = localStorage.getItem('checklistBulkData');
          return stored ? JSON.parse(stored) : null;
        } catch {
          return null;
        }
      })();
      if (bulkToMerge && bulkToMerge.length > 0) {
        parsedBulkDataCache = bulkToMerge;
        // Merge ALL ChecklistBulk rows so every bulk record is displayed (e.g. 13 Shops and Establishment forms)
        mergedData = mergeBulkDataWithStatutory(mergedData, bulkToMerge);
      }

      // Email-based sector filtering (Shops & Establishment only) for specific users
      // IMPORTANT: apply this AFTER bulk metadata enrichment so that `sector` is populated
      // for checklist-derived entries (e.g., Form D).

      // Note: Checklist entries are now also filtered by sector/act to ensure proper categorization
      // Forms from one sector should not appear in another sector's section

      // Safety net: ensure Form U entries from backend are not lost after merge/filter
      const backendFormUEntries = (baseStatutoryData || []).filter(item => {
        const formName = String(item.formName || '').toLowerCase().trim();
        return formName === 'form u' || formName.includes('form u');
      });
      const mergedHasFormU = (mergedData || []).some(item => {
        const formName = String(item.formName || '').toLowerCase().trim();
        return formName === 'form u' || formName.includes('form u');
      });
      if (backendFormUEntries.length > 0 && !mergedHasFormU) {
        console.warn('⚠️ Form U present in backend data but missing after merge. Reinserting from backend copy.');
        mergedData = [...mergedData, ...backendFormUEntries];
      }

      // Enrich missing Act/Description/DueDate/Sector/State using bulk metadata (keeps Form U fully populated)
      mergedData = enrichWithBulkMetadata(mergedData, parsedBulkDataCache);

      // Statutory merges DB rows + SEMaster Checklist + Formmaster + appended bulk rows — count can exceed ChecklistBulk.
      // Business rule: when ChecklistBulk master data exists, Statutory must mirror it 1:1.
      // Keep exactly one displayed row per ChecklistBulk line and only overlay file/draft metadata from merged sources.
      if (parsedBulkDataCache && parsedBulkDataCache.length > 0) {
        const beforeAlign = mergedData.length;
        const canonicalBulkRows = mergeBulkDataWithStatutory([], parsedBulkDataCache);
        const byIdentity = new Map();
        const byFormName = new Map();
        mergedData.forEach((item) => {
          const identityKey = checklistBulkRowIdentityKey(item);
          if (identityKey) {
            const list = byIdentity.get(identityKey) || [];
            list.push(item);
            byIdentity.set(identityKey, list);
          }
          const formNameKey = baseFormNameKey(item.formName || item.FormName);
          if (formNameKey) {
            const list = byFormName.get(formNameKey) || [];
            list.push(item);
            byFormName.set(formNameKey, list);
          }
        });
        const donorScore = (row) => {
          let score = 0;
          if (hasStatutoryDraftFileRef(row)) score += 1000;
          if (row?.formFile && String(row.formFile).trim() !== '') score += 100;
          const sfa = String(row?.sendForApproval ?? row?.SendForApproval ?? '').trim().toLowerCase();
          if (sfa === 'sent') score += 90;
          const approval = String(row?.approval ?? row?.Approval ?? '').trim().toLowerCase();
          if (approval === 'approved' || approval === 'rejected' || approval === 'approve' || approval === 'reject') score += 30;
          if (isNumericStatutoryBackendId(row?.id)) score += 10;
          return score;
        };
        mergedData = canonicalBulkRows.map((bulkRow) => {
          const identityKey = checklistBulkRowIdentityKey(bulkRow);
          const formNameKey = baseFormNameKey(bulkRow.formName || bulkRow.FormName);
          const identityMatches = byIdentity.get(identityKey) || [];
          const formMatches = byFormName.get(formNameKey) || [];
          // Keep draft/proof/approval metadata strict to the same statutory line.
          // Allow a second-level fallback by (form + act + description) only, because
          // bulk rows can have slight sector/state differences across sources.
          const norm = (v) => String(v ?? '').trim().toLowerCase().replace(/\s+/g, ' ');
          const bulkFormActDescKey = `${norm(bulkRow.formName || bulkRow.FormName)}|${norm(bulkRow.act || bulkRow.Act)}|${norm(bulkRow.description || bulkRow.Description)}`;
          const formActDescMatches = formMatches.filter((row) => {
            const rowKey = `${norm(row.formName || row.FormName)}|${norm(row.act || row.Act)}|${norm(row.description || row.Description)}`;
            return rowKey === bulkFormActDescKey;
          });
          const strictDonorCandidates = [...identityMatches, ...formActDescMatches];
          const formFileDonorCandidates = [...identityMatches, ...formMatches];
          const bulkMonthNorm = statutoryDedupeMonthNorm(bulkRow, selectedMonth);
          const sameMonthDonors = strictDonorCandidates.filter(
            (row) => statutoryDedupeMonthNorm(row, selectedMonth) === bulkMonthNorm
          );
          const scoredSameMonth = [...sameMonthDonors].sort((a, b) => donorScore(b) - donorScore(a));
          const monthDonor = scoredSameMonth[0] || null;
          const scoredAllForFormFile = [...formFileDonorCandidates].sort((a, b) => donorScore(b) - donorScore(a));
          const hasFormRef = (r) => {
            const fid = r?.formFile ?? r?.FormFile;
            return fid != null && String(fid).trim() !== '' && String(fid).trim() !== 'null';
          };
          // Template Form File may exist on another month's statutory row — ok to reuse for Autofill.
          // Draft / proof / send-for-approval / stored MonthFilter must NEVER bleed across months.
          let formDonor = monthDonor;
          if (!formDonor || !hasFormRef(formDonor)) {
            formDonor = scoredAllForFormFile.find(hasFormRef) || formDonor;
          }
          const uiMonthOnlyLabel =
            resolveToFullMonthName(String(selectedMonth || '').trim()) ||
            String(selectedMonth || '').trim() ||
            null;
          if (!monthDonor) {
            return {
              ...bulkRow,
              id: bulkRow.id,
              formFile: formDonor?.formFile ?? formDonor?.FormFile ?? bulkRow.formFile ?? null,
              formFileName: formDonor?.formFileName ?? formDonor?.FormFileName ?? bulkRow.formFileName ?? null,
              draftFile: null,
              draftFileName: null,
              draft: '',
              proofSubmissionFile: null,
              proofSubmissionFileName: null,
              sendForApproval: '',
              SendForApproval: '',
              approval: '',
              Approval: '',
              status: '',
              Status: '',
              remarks: '',
              Remarks: '',
              monthFilter: uiMonthOnlyLabel,
              MonthFilter: uiMonthOnlyLabel,
              monthfilter: uiMonthOnlyLabel,
              draftStatutoryRowIdForFile: null,
              checklistId: bulkRow.checklistId ?? formDonor?.checklistId ?? null,
              isFromChecklist: !!formDonor?.isFromChecklist && !!monthDonor
            };
          }
          const donor = monthDonor;
          return {
            ...bulkRow,
            // Keep canonical bulk id so each ChecklistBulk row remains distinct in UI grouping/deduping.
            id: bulkRow.id,
            checklistId: donor?.checklistId ?? bulkRow.checklistId ?? null,
            isFromChecklist: donor?.isFromChecklist ?? false,
            formFile: formDonor?.formFile ?? formDonor?.FormFile ?? bulkRow.formFile,
            formFileName: formDonor?.formFileName ?? formDonor?.FormFileName ?? bulkRow.formFileName,
            draftFile: donor?.draftFile ?? donor?.DraftFile ?? bulkRow.draftFile,
            draftFileName: donor?.draftFileName ?? donor?.DraftFileName ?? bulkRow.draftFileName,
            draft: donor?.draft ?? donor?.Draft ?? bulkRow.draft,
            proofSubmissionFile: donor?.proofSubmissionFile ?? donor?.ProofSubmissionFile ?? bulkRow.proofSubmissionFile ?? null,
            proofSubmissionFileName: donor?.proofSubmissionFileName ?? donor?.ProofSubmissionFileName ?? bulkRow.proofSubmissionFileName ?? null,
            sendForApproval: donor?.sendForApproval ?? donor?.SendForApproval ?? bulkRow.sendForApproval ?? '',
            SendForApproval: donor?.sendForApproval ?? donor?.SendForApproval ?? bulkRow.SendForApproval ?? '',
            approval: donor?.approval ?? donor?.Approval ?? bulkRow.approval ?? '',
            Approval: donor?.approval ?? donor?.Approval ?? bulkRow.Approval ?? '',
            status: donor?.status ?? donor?.Status ?? bulkRow.status ?? '',
            Status: donor?.status ?? donor?.Status ?? bulkRow.Status ?? '',
            remarks: donor?.remarks ?? donor?.Remarks ?? bulkRow.remarks ?? '',
            Remarks: donor?.remarks ?? donor?.Remarks ?? bulkRow.Remarks ?? '',
            // Preserve backend-saved month mapping so month filter shows the record only in its stored month.
            monthFilter: donor?.monthFilter ?? donor?.MonthFilter ?? donor?.monthfilter ?? bulkRow.monthFilter ?? null,
            MonthFilter: donor?.monthFilter ?? donor?.MonthFilter ?? donor?.monthfilter ?? bulkRow.MonthFilter ?? null,
            monthfilter: donor?.monthFilter ?? donor?.MonthFilter ?? donor?.monthfilter ?? bulkRow.monthfilter ?? null,
            draftStatutoryRowIdForFile: donor?.id ?? null
          };
        });
        console.log(
          `📌 Statutory aligned 1:1 with ChecklistBulk (${parsedBulkDataCache.length} master rows): ${beforeAlign} -> ${mergedData.length} displayed rows`
        );
      }

      // Filter by Site Management: Incharge email = login → Industry → act category (Factories / Shops / CLRA)
      if (siteScopeCategories == null) {
        siteScopeCategories = await siteScopePromise;
        setAllowedActCategoryList(siteScopeCategories);
      }
      if (siteScopeCategories && siteScopeCategories.length > 0 && mergedData.length > 0) {
        const allow = new Set(siteScopeCategories);
        const beforeFilter = mergedData.length;
        mergedData = mergedData.filter((item) => {
          const cat = getActCategoryWithFormFallback(item, mergedData);
          return allow.has(cat) || (cat === 'other' && hasComplianceFiles(item));
        });
        console.log(
          `🔍 Statutory filtered by Site Incharge industry scope [${siteScopeCategories.join(', ')}]: ${beforeFilter} -> ${mergedData.length}`
        );
      }

      // Deduplicate only when ChecklistBulk master data is not present.
      // When bulk exists, keep strict 1:1 count with ChecklistBulk rows.
      if (parsedBulkDataCache && parsedBulkDataCache.length > 0) {
        console.log(`⏭️ Skipping deduplication to preserve 1:1 ChecklistBulk row count (${mergedData.length} items)`);
      } else {
        console.log(`🔄 Before deduplication: ${mergedData.length} items`);
        mergedData = removeDuplicateStatutoryRecords(mergedData, resolveSiteDisplayName, selectedMonth);
        console.log(`🔄 After deduplication: ${mergedData.length} items`);
      }

      // Sort so rows with a form file appear first
      const hasFormFile = (item) => {
        const formFile = item?.formFile;
        return !!(formFile && formFile !== 'null' && String(formFile).trim() !== '');
      };
      const mergedWithOrder = mergedData.map((item, idx) => ({ ...item, __order: idx }));
      mergedWithOrder.sort((a, b) => {
        const aHas = hasFormFile(a);
        const bHas = hasFormFile(b);
        if (aHas !== bHas) return bHas - aHas; // true first
        return a.__order - b.__order; // keep stable order otherwise
      });
      mergedData = mergedWithOrder.map(({ __order, ...rest }) => rest);

      // Log Form U entries for debugging
      const formUEntries = mergedData.filter(item => {
        const formName = String(item.formName || '').toLowerCase();
        return formName.includes('form u') || formName === 'form u';
      });
      if (formUEntries.length > 0) {
        console.log(`📋 Found ${formUEntries.length} Form U entry/entries:`, formUEntries.map(item => ({
          id: item.id,
          formName: item.formName,
          act: item.act,
          formFile: item.formFile,
          formFileName: item.formFileName,
          formFileType: typeof item.formFile,
          formFileValue: String(item.formFile),
          hasFormFile: !!(item.formFile && item.formFile !== 'null' && String(item.formFile).trim() !== ''),
          willDisplay: !!(item.formFile && item.formFile !== 'null' && String(item.formFile).trim() !== '')
        })));
      } else {
        console.log('⚠️ No Form U entries found in statutory data');
        // Log all form names to help debug
        console.log('Available form names:', mergedData.map(item => item.formName || 'NO NAME').slice(0, 20));
      }
     
      // Log all forms with their file status
      console.log(`📊 Total statutory items before final processing: ${mergedData.length}`);
      console.log(`📊 Detailed statutory items:`, mergedData.map(item => ({
        id: item.id,
        formName: item.formName,
        hasFormFile: !!(item.formFile && item.formFile !== 'null' && String(item.formFile).trim() !== ''),
        formFileName: item.formFileName,
        formFile: item.formFile,
        isFromChecklist: item.isFromChecklist || false,
        checklistId: item.checklistId || null
      })));
     
      // Specifically check for Form I
      const formIEntries = mergedData.filter(item => {
        const formName = String(item.formName || '').toLowerCase().trim();
        return formName === 'form i';
      });
      if (formIEntries.length > 0) {
        console.log(`📋 Found ${formIEntries.length} Form I entry/entries:`, formIEntries.map(item => ({
          id: item.id,
          checklistId: item.checklistId,
          formName: item.formName,
          act: item.act,
          formFile: item.formFile,
          formFileName: item.formFileName,
          isFromChecklist: item.isFromChecklist || false,
          hasFormFile: !!(item.formFile && item.formFile !== 'null' && String(item.formFile).trim() !== '')
        })));
      } else {
        console.warn('⚠️ No Form I entries found in statutory data after all merges');
        console.log('Available form names:', mergedData.map(item => item.formName || 'NO NAME'));
      }
     
      // Final count summary
      console.log(`📊 FINAL COUNT: ${mergedData.length} total statutory entries will be displayed`);
      const byFormName = {};
      mergedData.forEach(item => {
        const name = String(item.formName || '').toLowerCase().trim();
        if (!byFormName[name]) byFormName[name] = 0;
        byFormName[name]++;
      });
      console.log(`📊 Entries by form name:`, byFormName);
     
      setStatutoryData(mergedData);
      setLastSyncedAt(new Date());
      localStorage.setItem('statutoryData', JSON.stringify(mergedData));
    } catch (err) {
      console.error('Error fetching statutory data:', err);
      const localData = localStorage.getItem('statutoryData');
      if (localData) {
        try {
          const parsed = JSON.parse(localData);
          const sc = await fetchAllowedActCategoriesFromSites(userEmail);
          if (Array.isArray(parsed) && parsed.length > 0 && sc?.length) {
            const allow = new Set(sc);
            setStatutoryData(parsed.filter((item) => {
              const cat = getActCategoryWithFormFallback(item, parsed);
              return allow.has(cat) || (cat === 'other' && hasComplianceFiles(item));
            }));
          } else if (Array.isArray(parsed)) {
            setStatutoryData(parsed);
          } else {
            setStatutoryData([]);
          }
        } catch (_) {
          setStatutoryData([]);
        }
      } else {
        // When statutory API fails, still try to show ChecklistBulk data so page is not empty
        const applySiteScopeFilter = async (data) => {
          if (!data || data.length === 0) return data;
          const sc = await fetchAllowedActCategoriesFromSites(userEmail);
          if (!sc?.length) return data;
          const allow = new Set(sc);
          return data.filter((item) => {
            const cat = getActCategoryWithFormFallback(item, data);
            return allow.has(cat) || (cat === 'other' && hasComplianceFiles(item));
          });
        };
        try {
          const bulkResponse = await fetch('/server/checklistbulk_function/checklistbulk?action=getAll');
          if (bulkResponse.ok) {
            const bulkResult = await bulkResponse.json();
            if (bulkResult.status === 'success' && bulkResult.data && Array.isArray(bulkResult.data) && bulkResult.data.length > 0) {
              const merged = mergeBulkDataWithStatutory([], bulkResult.data);
              setStatutoryData(await applySiteScopeFilter(merged));
              localStorage.setItem('checklistBulkData', JSON.stringify(bulkResult.data));
              return;
            }
          }
        } catch (bulkErr) {
          console.error('Error fetching ChecklistBulk fallback:', bulkErr);
        }
        const fallbackBulk = localStorage.getItem('checklistBulkData');
        if (fallbackBulk) {
          try {
            const parsed = JSON.parse(fallbackBulk);
            if (Array.isArray(parsed) && parsed.length > 0) {
              const merged = mergeBulkDataWithStatutory([], parsed);
              setStatutoryData(await applySiteScopeFilter(merged));
              return;
            }
          } catch (_) {}
        }
        setError('Failed to fetch statutory data. Please try again.');
      }
    } finally {
      if (!silentRefresh) setLoading(false);
      isFetchingStatutoryRef.current = false;
      setTransactionSiteMetaReady(true);
    }
  };

  useEffect(() => {
    // On re-entering Statutory, show cached data immediately and refresh in background.
    fetchStatutoryData({ useCacheFirst: true, silentRefresh: true });
  }, [userEmail]);

  // Listen for storage changes to detect when checklistBulkData is updated
  useEffect(() => {
    const handleStorageChange = (e) => {
      if (e.key === 'checklistBulkData') {
        console.log('📦 checklistBulkData changed in localStorage, refreshing Statutory data...');
        fetchStatutoryData();
      }
    };

    // Listen for storage changes
    window.addEventListener('storage', handleStorageChange);

    // Also listen for custom event when ChecklistBulk clears data
    const handleBulkDataCleared = () => {
      console.log('📦 checklistBulkData cleared event received, refreshing Statutory data...');
      fetchStatutoryData();
    };

    window.addEventListener('checklistBulkDataCleared', handleBulkDataCleared);

    // Listen for custom event when statutory data is updated (e.g., from SEMaster)
    const handleStatutoryDataUpdated = (e) => {
      console.log('📋 statutoryDataUpdated event received, refreshing Statutory data...', e.detail);
      fetchStatutoryData();
    };

    window.addEventListener('statutoryDataUpdated', handleStatutoryDataUpdated);

    return () => {
      window.removeEventListener('storage', handleStorageChange);
      window.removeEventListener('checklistBulkDataCleared', handleBulkDataCleared);
      window.removeEventListener('statutoryDataUpdated', handleStatutoryDataUpdated);
    };
  }, []);

  const onChange = (e) => {
    const { name, value, files } = e.target;

    if (name === 'formFile' && files && files[0]) {
      const file = files[0];

      // Validate file size (max 5MB)
      if (file.size > 5 * 1024 * 1024) {
        setError('File size must be less than 5MB');
        return;
      }

      // Store the file and filename for upload
      setForm(prev => ({
        ...prev,
        [name]: file,
        formFileName: file.name
      }));
      setError('');
    } else if (name === 'proofSubmissionFile' && files && files[0]) {
      const file = files[0];

      // Validate file size (max 5MB)
      if (file.size > 5 * 1024 * 1024) {
        setError('File size must be less than 5MB');
        return;
      }
      if (!isProofSubmissionDocumentFileName(file.name)) {
        setError('Proof submission must be a PDF or image file.');
        return;
      }

      // Store the file and filename for upload
      setForm(prev => ({
        ...prev,
        [name]: file,
        proofSubmissionFileName: file.name
      }));
      setError('');
    } else if (name === 'draftFile' && files && files[0]) {
      const file = files[0];

      // Validate file size (max 5MB)
      if (file.size > 5 * 1024 * 1024) {
        setError('File size must be less than 5MB');
        return;
      }

      // Store the file and filename for upload
      setForm(prev => ({
        ...prev,
        [name]: file,
        draftFileName: file.name
      }));
      setError('');
    } else {
      setForm(prev => ({ ...prev, [name]: value }));
    }
    setError('');
    setSuccess('');
  };

  const onSubmit = async (e) => {
    e.preventDefault();

    // Prevent duplicate submissions
    if (submitting) {
      return;
    }

    // Validate required fields
    if (!form.formName.trim()) {
      setError('Form Name is required');
      return;
    }

    setSubmitting(true);

    try {
      setError('');
      setSuccess('');

      // Proof Submission is never filled from Draft; only from an uploaded PDF/image.

      // Upload files first if present
      let formFileId = null;
      let formFileName = null;

      // Check if formFile is a File object (new upload) or already a file ID (existing)
      const isNewFormFile = form.formFile instanceof File;

      // Upload form file if present and it's a new file
      if (isNewFormFile) {
        try {
          const formData = new FormData();
          formData.append('file', form.formFile);

          const uploadResp = await fetch('/server/statutoryreg_function/statutory/upload/Form', {
            method: 'POST',
            body: formData
          });

          let uploadData;
          const contentType = uploadResp.headers.get('content-type') || '';

          const responseClone = uploadResp.clone();

          if (contentType.includes('application/json')) {
            try {
              uploadData = await uploadResp.json();
            } catch (jsonErr) {
              const textResponse = await responseClone.text();
              console.error('Failed to parse JSON response:', textResponse);
              setError(`Form file upload failed: ${textResponse || 'Invalid server response'}`);
              setSubmitting(false);
              return;
            }
          } else {
            const textResponse = await uploadResp.text();
            console.error('Non-JSON response received:', textResponse);
            setError(`Form file upload failed: ${textResponse || 'Server error occurred'}`);
            setSubmitting(false);
            return;
          }

          if (uploadResp.ok && uploadData.status === 'success') {
            formFileId = uploadData.fileId;
            formFileName = uploadData.fileName || form.formFileName;
            console.log('Form file uploaded successfully:', formFileId);
          } else {
            const errorMessage = uploadData.message || 'Form file upload failed';
            console.error('Form file upload error:', errorMessage);
            setError(errorMessage);
            setSubmitting(false);
            return;
          }
        } catch (uploadErr) {
          console.error('Error uploading form file:', uploadErr);
          setError(`Form file upload failed: ${uploadErr.message}`);
          setSubmitting(false);
          return;
        }
      } else {
        // Use existing file ID
        formFileId = form.formFile;
        formFileName = form.formFileName;
      }

      // Upload proof submission file if present and it's a new file
      let proofSubmissionFileId = null;
      let proofSubmissionFileName = null;
      const isNewProofFile = form.proofSubmissionFile instanceof File;

      if (isNewProofFile) {
        if (!isProofSubmissionDocumentFileName(form.proofSubmissionFile?.name)) {
          setError('Proof submission must be a PDF or image file.');
          setSubmitting(false);
          return;
        }
        try {
          const proofFormData = new FormData();
          proofFormData.append('file', form.proofSubmissionFile);

          const proofUploadResp = await fetch('/server/statutoryreg_function/statutory/upload/ProofSubmission', {
            method: 'POST',
            body: proofFormData
          });

          let proofUploadData;
          const proofContentType = proofUploadResp.headers.get('content-type') || '';

          const proofResponseClone = proofUploadResp.clone();

          if (proofContentType.includes('application/json')) {
            try {
              proofUploadData = await proofUploadResp.json();
            } catch (jsonErr) {
              const textResponse = await proofResponseClone.text();
              console.error('Failed to parse JSON response:', textResponse);
              setError(`Proof submission file upload failed: ${textResponse || 'Invalid server response'}`);
              setSubmitting(false);
              return;
            }
          } else {
            const textResponse = await proofUploadResp.text();
            console.error('Non-JSON response received:', textResponse);
            setError(`Proof submission file upload failed: ${textResponse || 'Server error occurred'}`);
            setSubmitting(false);
            return;
          }

          if (proofUploadResp.ok && proofUploadData.status === 'success') {
            const uploadedProofId = proofUploadData.fileId ?? proofUploadData.data?.fileId ?? proofUploadData.data?.file_id;
            const uploadedProofName = proofUploadData.fileName ?? proofUploadData.data?.fileName ?? form.proofSubmissionFileName;
            if (uploadedProofId != null && String(uploadedProofId).trim() !== '') {
              proofSubmissionFileId = String(uploadedProofId).trim();
              proofSubmissionFileName = (uploadedProofName != null && String(uploadedProofName).trim() !== '') ? String(uploadedProofName).trim() : (form.proofSubmissionFileName || 'proof');
              console.log('Proof submission file uploaded successfully:', proofSubmissionFileId, proofSubmissionFileName);
            } else {
              console.error('Proof upload response missing fileId:', proofUploadData);
              setError('Proof submission upload succeeded but server did not return file ID.');
              setSubmitting(false);
              return;
            }
          } else {
            const errorMessage = proofUploadData.message || 'Proof submission file upload failed';
            console.error('Proof submission file upload error:', errorMessage);
            setError(errorMessage);
            setSubmitting(false);
            return;
          }
        } catch (uploadErr) {
          console.error('Error uploading proof submission file:', uploadErr);
          setError(`Proof submission file upload failed: ${uploadErr.message}`);
          setSubmitting(false);
          return;
        }
      } else {
        const existingProofId = form.proofSubmissionFile || null;
        const existingProofName = form.proofSubmissionFileName || null;
        const keepExisting =
          isProofSubmissionDocumentFileName(existingProofName) &&
          existingProofId != null &&
          String(existingProofId).trim() !== '' &&
          String(existingProofId).trim() !== 'null' &&
          String(existingProofId).trim() !== 'undefined';
        proofSubmissionFileId = keepExisting ? String(existingProofId).trim() : null;
        proofSubmissionFileName = keepExisting &&
          existingProofName != null &&
          String(existingProofName).trim() !== '' &&
          String(existingProofName).trim() !== 'null' &&
          String(existingProofName).trim() !== 'undefined'
          ? String(existingProofName).trim()
          : null;
      }

      // Upload draft file if present and it's a new file
      let draftFileId = null;
      let draftFileName = null;
      const isNewDraftFile = form.draftFile instanceof File;

      if (isNewDraftFile) {
        try {
          const draftFormData = new FormData();
          draftFormData.append('file', form.draftFile);

          const draftUploadResp = await fetch('/server/statutoryreg_function/statutory/upload/Draft', {
            method: 'POST',
            body: draftFormData
          });

          let draftUploadData;
          const draftContentType = draftUploadResp.headers.get('content-type') || '';

          const draftResponseClone = draftUploadResp.clone();

          if (draftContentType.includes('application/json')) {
            try {
              draftUploadData = await draftUploadResp.json();
            } catch (jsonErr) {
              const textResponse = await draftResponseClone.text();
              console.error('Failed to parse JSON response:', textResponse);
              setError(`Draft file upload failed: ${textResponse || 'Invalid server response'}`);
              setSubmitting(false);
              return;
            }
          } else {
            const textResponse = await draftUploadResp.text();
            console.error('Non-JSON response received:', textResponse);
            setError(`Draft file upload failed: ${textResponse || 'Server error occurred'}`);
            setSubmitting(false);
            return;
          }

          if (draftUploadResp.ok && draftUploadData.status === 'success') {
            draftFileId = draftUploadData.fileId;
            draftFileName = draftUploadData.fileName || form.draftFileName;
            console.log('Draft file uploaded successfully:', draftFileId);
          } else {
            const errorMessage = draftUploadData.message || 'Draft file upload failed';
            console.error('Draft file upload error:', errorMessage);
            setError(errorMessage);
            setSubmitting(false);
            return;
          }
        } catch (uploadErr) {
          console.error('Error uploading draft file:', uploadErr);
          setError(`Draft file upload failed: ${uploadErr.message}`);
          setSubmitting(false);
          return;
        }
      } else {
        // Use existing file ID
        draftFileId = form.draftFile;
        draftFileName = form.draftFileName;
      }

      // Prepare the payload (month filter saved so it displays in backend – e.g. Proof Submission save)
      const monthFilterValue = resolveMonthFilterForSave(form.dueDate, selectedMonth, form.originalMonthFilter);
      const proofFileForPayload = (proofSubmissionFileId != null && String(proofSubmissionFileId).trim() !== '' && String(proofSubmissionFileId).trim() !== 'null' && String(proofSubmissionFileId).trim() !== 'undefined') ? String(proofSubmissionFileId).trim() : null;
      const proofNameForPayload = (proofSubmissionFileName != null && String(proofSubmissionFileName).trim() !== '' && String(proofSubmissionFileName).trim() !== 'null' && String(proofSubmissionFileName).trim() !== 'undefined') ? String(proofSubmissionFileName).trim() : null;
      const payload = {
        act: form.act || null,
        description: form.description || null,
        sector: form.sector || null,
        state: form.state || null,
        site: form.site || siteFromUrl || null,
        formFile: formFileId != null ? String(formFileId) : null,
        formFileName: formFileName || null,
        formName: form.formName.trim(),
        dueDate: form.dueDate || null,
        monthFilter: monthFilterValue,
        MonthFilter: monthFilterValue,
        monthfilter: monthFilterValue,
        autofill: form.autofill || null,
        draft: form.draft || null,
        proofSubmissionFile: proofFileForPayload,
        proofSubmissionFileName: proofNameForPayload,
        draftFile: draftFileId != null ? String(draftFileId) : null,
        draftFileName: draftFileName || null,
        approval: form.approval != null && String(form.approval).trim() !== '' ? String(form.approval).trim() : null,
        status: form.status != null && String(form.status).trim() !== '' ? String(form.status).trim() : null,
        sendForApproval:
          form.sendForApproval != null && String(form.sendForApproval).trim() !== ''
            ? String(form.sendForApproval).trim()
            : null,
        remarks: form.remarks != null && String(form.remarks).trim() !== '' ? String(form.remarks).trim() : null
      };

      const existingMonthNorm = (form.originalMonthFilter || '').trim().toLowerCase().substring(0, 3);
      const savingMonthNorm = (monthFilterValue || '').trim().toLowerCase().substring(0, 3);
      const sameMonth = existingMonthNorm && savingMonthNorm && existingMonthNorm === savingMonthNorm;
      const createNewForMonth = editingStatutoryId && !sameMonth;
      const isPlaceholderEdit = editingStatutoryId && String(editingStatutoryId).startsWith('placeholder_');

      let response;
      if (editingStatutoryId && sameMonth && !isPlaceholderEdit) {
        response = await fetch(`/server/statutoryreg_function/statutory/${editingStatutoryId}`, {
          method: 'PUT',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(payload)
        });
      } else if (createNewForMonth || isPlaceholderEdit) {
        response = await fetch('/server/statutoryreg_function/statutory', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(payload)
        });
      } else {
        response = await fetch('/server/statutoryreg_function/statutory', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(payload)
        });
      }

      const data = await response.json();
      const didUpdate = editingStatutoryId && sameMonth && !isPlaceholderEdit;

      if (response.ok && data.status === 'success') {
        setSuccess(didUpdate ? 'Statutory updated successfully!' : 'Statutory created successfully!');
        setForm(initialForm);
        setEditingStatutoryId(null);
        setIsFormOpen(false);
        if (form.dueDate) {
          const monthFromDueDate = getMonthFromDueDate(form.dueDate);
          if (monthFromDueDate) {
            setSelectedMonth(monthFromDueDate);
            persistMonthFilter(monthFromDueDate);
          }
        }
        // When a new file was imported, clear old ChecklistBulk data automatically (like checklistbulk clear)
        if (isNewFormFile) {
          try {
            const bulkDeleteResp = await fetch('/server/checklistbulk_function/checklistbulk?action=bulkDelete', { method: 'DELETE' });
            const bulkDeleteData = await bulkDeleteResp.json().catch(() => ({}));
            if (bulkDeleteResp.ok && (bulkDeleteData.status === 'success' || bulkDeleteData.status === 'partial')) {
              localStorage.removeItem('checklistBulkData');
              window.dispatchEvent(new Event('checklistBulkDataCleared'));
            }
          } catch (bulkErr) {
            console.warn('Could not clear ChecklistBulk after import:', bulkErr);
          }
        }
        fetchStatutoryData();
        setTimeout(() => setSuccess(''), 3000);
      } else {
        const errorMessage = data.message || (didUpdate ? 'Failed to update statutory' : 'Failed to create statutory');
        setError(errorMessage);
      }
    } catch (err) {
      console.error('Error submitting statutory:', err);
      setError(`Failed to ${editingStatutoryId ? 'update' : 'create'} statutory: ${err.message}`);
    } finally {
      setSubmitting(false);
    }
  };

  const onReset = () => {
    setForm(initialForm);
    setEditingStatutoryId(null);
    setError('');
    setSuccess('');
  };

  const handleClose = () => {
    setIsFormOpen(false);
    setForm(initialForm);
    setEditingStatutoryId(null);
    setSubmitting(false);
    setError('');
    setSuccess('');
    // Optionally clear selected rows when closing form
    // setSelectedRows(new Set());
  };

  const handleEditStatutory = (statutoryId, rowItem) => {
    const statutoryItem = statutoryData.find(r => r.id === statutoryId);
    const source = statutoryItem || rowItem;
    if (source) {
      if (isRowStatusApproved(source)) {
        setError('Approved rows are locked and cannot be edited.');
        setTimeout(() => setError(''), 5000);
        return;
      }
      const rowMonth = source.monthFilter || source.MonthFilter || source.monthfilter || '';
      const rawProofId = source.proofSubmissionFile ?? source.ProofSubmissionFile ?? null;
      const rawProofName = source.proofSubmissionFileName ?? source.ProofSubmissionFileName ?? null;
      const proofOk = isProofSubmissionDocumentFileName(rawProofName);
      setForm({
        act: source.act || '',
        description: source.description || '',
        sector: source.sector || '',
        state: source.state || '',
        site: source.site || source.Site || source.siteName || source.SiteName || '',
        formFile: source.formFile ?? source.FormFile ?? null,
        formFileName: source.formFileName ?? source.FormFileName ?? null,
        formName: source.formName ?? source.FormName ?? '',
        dueDate: source.dueDate || '',
        autofill: source.autofill || '',
        draft: source.draft || '',
        proofSubmissionFile: proofOk ? rawProofId : null,
        proofSubmissionFileName: proofOk ? rawProofName : null,
        draftFile: source.draftFile ?? source.DraftFile ?? null,
        draftFileName: source.draftFileName ?? source.DraftFileName ?? null,
        approval: source.approval ?? source.Approval ?? '',
        status: source.status != null && source.status !== '' ? source.status : source.Status != null ? source.Status : '',
        sendForApproval: source.sendForApproval ?? source.SendForApproval ?? '',
        remarks: source.remarks ?? source.Remarks ?? '',
        originalMonthFilter: rowMonth || null
      });
      setEditingStatutoryId(source.id);
      setIsFormOpen(true);
    }
  };

  const handleDeleteStatutory = async (statutoryId) => {
    if (!window.confirm('Are you sure you want to delete this statutory record?')) {
      return;
    }

    try {
      setLoading(true);
      const response = await fetch(`/server/statutoryreg_function/statutory/${statutoryId}`, {
        method: 'DELETE'
      });

      const data = await response.json();

      if (response.ok && data.status === 'success') {
        setSuccess('Statutory deleted successfully!');
        // Remove from selected rows if it was selected
        setSelectedRows(prev => {
          const newSet = new Set(prev);
          newSet.delete(statutoryId);
          return newSet;
        });
        fetchStatutoryData();
        setTimeout(() => setSuccess(''), 3000);
      } else {
        setError(data.message || 'Failed to delete statutory');
      }
    } catch (err) {
      console.error('Error deleting statutory:', err);
      setError(`Failed to delete statutory: ${err.message}`);
    } finally {
      setLoading(false);
    }
  };

  const handleCheckboxChange = (statutoryId) => {
    setSelectedRows(prev => {
      const newSet = new Set(prev);
      if (newSet.has(statutoryId)) {
        newSet.delete(statutoryId);
      } else {
        newSet.add(statutoryId);
      }
      return newSet;
    });
  };

  const proofRowDedupKey = (row) => {
    const form = (row.formName || row.FormName || '').trim().toLowerCase();
    const act = (row.act || row.Act || '').trim().toLowerCase();
    const desc = (row.description || row.Description || '').trim().toLowerCase();
    const sector = String(row.sector || row.Sector || '').trim().toLowerCase();
    const monthNorm = (
      row.monthFilter ||
      row.MonthFilter ||
      row.monthfilter ||
      getMonthFromDueDate(row.dueDate) ||
      selectedMonth ||
      ''
    )
      .toString()
      .trim()
      .toLowerCase()
      .substring(0, 3) || 'nomonth';
    const state = String(row.state || row.State || '').trim().toLowerCase();
    return `${form}_${act}_${desc}_${monthNorm}_${sector || 'nosector'}_${state || 'nostate'}`;
  };

  const resolveNumericStatutoryIdForProofRow = (item) => {
    if (/^\d+$/.test(String(item.id ?? ''))) return String(item.id);
    if (/^\d+$/.test(String(item?.draftStatutoryRowIdForFile ?? ''))) {
      return String(item.draftStatutoryRowIdForFile);
    }
    const key = proofRowDedupKey(item);
    if (!key || key.startsWith('_')) return null;
    const match = (statutoryData || []).find((r) => /^\d+$/.test(String(r.id)) && proofRowDedupKey(r) === key);
    if (match) return String(match.id);

    // Fallback: strict key can fail when description/act text differs between bulk placeholder and saved statutory row.
    // Match by base form + month + sector/state and prefer rows that already have saved files.
    const monthNorm = String(
      item?.monthFilter ||
      item?.MonthFilter ||
      item?.monthfilter ||
      getMonthFromDueDate(item?.dueDate) ||
      selectedMonth ||
      ''
    )
      .trim()
      .toLowerCase()
      .substring(0, 3);
    const formNorm = baseFormNameKey(item?.formName || item?.FormName);
    const sectorNorm = String(item?.sector || item?.Sector || '').trim().toLowerCase();
    const stateNorm = String(item?.state || item?.State || '').trim().toLowerCase();
    const candidates = (statutoryData || []).filter((r) => {
      if (!/^\d+$/.test(String(r?.id ?? ''))) return false;
      const rMonthNorm = String(
        r?.monthFilter ||
        r?.MonthFilter ||
        r?.monthfilter ||
        getMonthFromDueDate(r?.dueDate) ||
        selectedMonth ||
        ''
      )
        .trim()
        .toLowerCase()
        .substring(0, 3);
      if (monthNorm && rMonthNorm && monthNorm !== rMonthNorm) return false;
      if (baseFormNameKey(r?.formName || r?.FormName) !== formNorm) return false;
      if (String(r?.sector || r?.Sector || '').trim().toLowerCase() !== sectorNorm) return false;
      if (String(r?.state || r?.State || '').trim().toLowerCase() !== stateNorm) return false;
      return true;
    });
    if (candidates.length === 0) return null;
    candidates.sort((a, b) => {
      const score = (row) => {
        let s = 0;
        const draft = row?.draftFile ?? row?.DraftFile;
        const proof = row?.proofSubmissionFile ?? row?.ProofSubmissionFile;
        const formFile = row?.formFile ?? row?.FormFile;
        if (draft && String(draft).trim() !== '' && String(draft).trim() !== 'null') s += 100;
        if (proof && String(proof).trim() !== '' && String(proof).trim() !== 'null') s += 80;
        if (formFile && String(formFile).trim() !== '' && String(formFile).trim() !== 'null') s += 20;
        return s;
      };
      return score(b) - score(a);
    });
    return String(candidates[0].id);
  };
  const buildMonthFilterPayload = (item) => {
    const monthValue = resolveMonthFilterForSave(
      item?.dueDate,
      selectedMonth,
      item?.monthFilter || item?.MonthFilter || item?.monthfilter
    );
    return {
      monthFilter: monthValue,
      MonthFilter: monthValue,
      monthfilter: monthValue
    };
  };

  const submitApprovalColumnAction = async (rowItem, decision, rejectRemarks) => {
    const targetId = resolveNumericStatutoryIdForProofRow(rowItem);
    if (!targetId) return false;
    const decisionNorm = String(decision || '').trim().toLowerCase();
    const isApproveDecision = decisionNorm === 'approved' || decisionNorm === 'approve';
    const isRejectDecision = decisionNorm === 'rejected' || decisionNorm === 'reject';
    setError('');
    setSuccess('');
    setApprovalUpdatingRowId(targetId);
    const payloadApproval = decision == null || decision === '' ? null : decision;
    const payloadStatus = isApproveDecision ? 'Approved' : isRejectDecision ? 'Rejected' : undefined;
    try {
      const resp = await fetch(`/server/statutoryreg_function/statutory/${targetId}`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          ...buildMonthFilterPayload(rowItem),
          approval: payloadApproval,
          ...(payloadStatus ? { status: payloadStatus } : {}),
          ...(isRejectDecision ? { remarks: rejectRemarks } : {}),
          ...(isApproveDecision ? { remarks: null } : {})
        })
      });
      const data = await resp.json().catch(() => ({}));
      if (!resp.ok || data.status !== 'success') {
        setError(data.message || 'Failed to update approval.');
        setTimeout(() => setError(''), 5000);
        return false;
      }
      setSuccess(payloadApproval == null ? 'Approval cleared.' : `Approval: ${payloadApproval}.`);
      fetchStatutoryData();
      window.dispatchEvent(new CustomEvent('statutoryDataUpdated', { detail: { id: targetId } }));
      setTimeout(() => setSuccess(''), 2500);
      return true;
    } catch (err) {
      setError(err?.message || 'Failed to update approval.');
      setTimeout(() => setError(''), 5000);
      return false;
    } finally {
      setApprovalUpdatingRowId(null);
    }
  };

  const closeRejectRemarksModal = () => {
    setIsRejectRemarksModalOpen(false);
    setRejectRemarksValue('');
    setRejectRemarksTarget(null);
  };

  const submitRejectRemarks = async () => {
    if (!rejectRemarksTarget?.rowItem || !rejectRemarksTarget?.decision) return;
    const trimmedRemarks = String(rejectRemarksValue || '').trim();
    if (!trimmedRemarks) {
      setError('Remarks are required when rejecting.');
      setTimeout(() => setError(''), 5000);
      return;
    }
    const ok = await submitApprovalColumnAction(
      rejectRemarksTarget.rowItem,
      rejectRemarksTarget.decision,
      trimmedRemarks
    );
    if (ok) closeRejectRemarksModal();
  };

  const handleApprovalColumnAction = async (rowItem, decision) => {
    const targetId = resolveNumericStatutoryIdForProofRow(rowItem);
    if (!targetId) {
      setError('Approval can only be set for saved statutory table rows.');
      setTimeout(() => setError(''), 5000);
      return;
    }
    const decisionNorm = String(decision || '').trim().toLowerCase();
    const isRejectDecision = decisionNorm === 'rejected' || decisionNorm === 'reject';
    if (isRejectDecision) {
      setRejectRemarksTarget({ rowItem, decision, targetId });
      setRejectRemarksValue(String(rowItem?.remarks ?? rowItem?.Remarks ?? ''));
      setIsRejectRemarksModalOpen(true);
      return;
    }
    await submitApprovalColumnAction(rowItem, decision);
  };

  const handleSendForApprovalColumnAction = async (rowItem) => {
    const targetId = resolveNumericStatutoryIdForProofRow(rowItem);
    if (!targetId) {
      setError('Send for approval is only available for saved statutory table rows.');
      setTimeout(() => setError(''), 5000);
      return;
    }
    setError('');
    setSuccess('');
    setSendForApprovalUpdatingRowId(targetId);
    try {
      const resp = await fetch(`/server/statutoryreg_function/statutory/${targetId}`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          ...buildMonthFilterPayload(rowItem),
          SendForApproval: 'Sent',
          sendForApproval: 'Sent',
          // Re-send flow: clear prior approver decision so site sees "Sent" after re-submit.
          approval: null,
          status: 'Pending',
          remarks: null
        })
      });
      const data = await resp.json().catch(() => ({}));
      if (!resp.ok || data.status !== 'success') {
        setError(data.message || 'Failed to update send for approval.');
        setTimeout(() => setError(''), 5000);
        return;
      }
      // Optimistic local update so UI reflects Sent even if another fetch is currently running.
      setStatutoryData((prev) =>
        Array.isArray(prev)
          ? (() => {
              const baseForm = baseFormNameKey(rowItem?.formName || rowItem?.FormName);
              const siteNorm = String(
                resolveSiteDisplayName(rowItem, prev) || rowItem?.site || rowItem?.Site || ''
              )
                .trim()
                .toLowerCase();
              return prev.map((row) => {
                const sameId = String(row?.id ?? '') === String(targetId);
                const sameBaseForm = baseFormNameKey(row?.formName || row?.FormName) === baseForm;
                const rowSiteNorm = String(
                  resolveSiteDisplayName(row, prev) || row?.site || row?.Site || ''
                )
                  .trim()
                  .toLowerCase();
                // Site-view fallback: when visible row is bulk/checklist sibling, mirror Sent to same form in same site.
                const sameSiteFormSibling =
                  !!siteNorm && sameBaseForm && resolvedSitesOverlap(siteNorm, rowSiteNorm);
                if (!sameId && !sameSiteFormSibling) return row;
                return {
                  ...row,
                  sendForApproval: 'Sent',
                  SendForApproval: 'Sent',
                  approval: '',
                  Approval: '',
                  status: 'Pending',
                  Status: 'Pending',
                  remarks: '',
                  Remarks: ''
                };
              });
            })()
          : prev
      );
      setSuccess('Marked as sent for approval.');
      fetchStatutoryData({ force: true });
      window.dispatchEvent(new CustomEvent('statutoryDataUpdated', { detail: { id: targetId } }));
      setTimeout(() => setSuccess(''), 2500);
    } catch (err) {
      setError(err?.message || 'Failed to update send for approval.');
      setTimeout(() => setError(''), 5000);
    } finally {
      setSendForApprovalUpdatingRowId(null);
    }
  };

  const handleInlineProofUpload = async (rowItem, file) => {
    if (!file || rowItem == null) return;
    const formNameTrim = (rowItem.formName || rowItem.FormName || '').trim();
    if (!formNameTrim) {
      setError('Form name is missing; cannot save proof for this row.');
      setTimeout(() => setError(''), 5000);
      return;
    }
    if (file.size > 5 * 1024 * 1024) {
      setError('Proof file size must be less than 5MB');
      setTimeout(() => setError(''), 5000);
      return;
    }
    if (!isProofSubmissionDocumentFileName(file.name)) {
      setError('Proof submission must be a PDF or image file.');
      setTimeout(() => setError(''), 5000);
      return;
    }
    setError('');
    setProofUploadingRowId(String(rowItem.id ?? ''));
    try {
      const proofFormData = new FormData();
      proofFormData.append('file', file);

      const proofUploadResp = await fetch('/server/statutoryreg_function/statutory/upload/ProofSubmission', {
        method: 'POST',
        body: proofFormData
      });

      const proofContentType = proofUploadResp.headers.get('content-type') || '';
      let proofUploadData;
      if (proofContentType.includes('application/json')) {
        proofUploadData = await proofUploadResp.json();
      } else {
        const textResponse = await proofUploadResp.text();
        setError(`Proof upload failed: ${textResponse || 'Invalid response'}`);
        setTimeout(() => setError(''), 5000);
        return;
      }

      if (!proofUploadResp.ok || proofUploadData.status !== 'success') {
        setError(proofUploadData.message || 'Proof submission file upload failed');
        setTimeout(() => setError(''), 5000);
        return;
      }

      const uploadedProofId = proofUploadData.fileId ?? proofUploadData.data?.fileId ?? proofUploadData.data?.file_id;
      const uploadedProofName = proofUploadData.fileName ?? proofUploadData.data?.fileName ?? file.name;
      if (uploadedProofId == null || String(uploadedProofId).trim() === '') {
        setError('Proof upload succeeded but server did not return file ID.');
        setTimeout(() => setError(''), 5000);
        return;
      }

      const proofSubmissionFileId = String(uploadedProofId).trim();
      const proofSubmissionFileName =
        uploadedProofName != null && String(uploadedProofName).trim() !== ''
          ? String(uploadedProofName).trim()
          : file.name;

      let targetId = resolveNumericStatutoryIdForProofRow(rowItem);

      if (targetId) {
        const putResp = await fetch(`/server/statutoryreg_function/statutory/${targetId}`, {
          method: 'PUT',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            ...buildMonthFilterPayload(rowItem),
            proofSubmissionFile: proofSubmissionFileId,
            proofSubmissionFileName
          })
        });
        const putData = await putResp.json().catch(() => ({}));
        if (!putResp.ok || putData.status !== 'success') {
          setError(putData.message || 'Failed to attach proof to statutory record.');
          setTimeout(() => setError(''), 5000);
          return;
        }
      } else {
        const monthFilterValue = resolveMonthFilterForSave(
          rowItem.dueDate,
          selectedMonth,
          rowItem.monthFilter || rowItem.MonthFilter || rowItem.monthfilter
        );
        const formFileId = rowItem.formFile ?? rowItem.FormFile ?? null;
        const formFileNameVal = rowItem.formFileName ?? rowItem.FormFileName ?? null;
        const draftFileId = rowItem.draftFile ?? rowItem.DraftFile ?? null;
        const draftFileNameVal = rowItem.draftFileName ?? rowItem.DraftFileName ?? null;
        const approvalVal = rowItem.approval ?? rowItem.Approval;
        const statusVal = rowItem.status != null && rowItem.status !== '' ? rowItem.status : rowItem.Status;
        const sfaVal = rowItem.sendForApproval ?? rowItem.SendForApproval;
        const remarksVal = rowItem.remarks ?? rowItem.Remarks;
        const payload = {
          act: rowItem.act || rowItem.Act || null,
          description: rowItem.description || rowItem.Description || null,
          sector: rowItem.sector || rowItem.Sector || null,
          state: rowItem.state || rowItem.State || null,
          site: rowItem.site || rowItem.Site || rowItem.siteName || rowItem.SiteName || null,
          formFile: formFileId != null && String(formFileId).trim() !== '' && String(formFileId).trim() !== 'null' ? String(formFileId).trim() : null,
          formFileName: formFileNameVal || null,
          formName: formNameTrim,
          dueDate: rowItem.dueDate || null,
          monthFilter: monthFilterValue,
          MonthFilter: monthFilterValue,
          monthfilter: monthFilterValue,
          autofill: rowItem.autofill || null,
          draft: rowItem.draft || null,
          proofSubmissionFile: proofSubmissionFileId,
          proofSubmissionFileName,
          draftFile: draftFileId != null && String(draftFileId).trim() !== '' && String(draftFileId).trim() !== 'null' ? String(draftFileId).trim() : null,
          draftFileName: draftFileNameVal || null,
          approval: approvalVal != null && String(approvalVal).trim() !== '' ? String(approvalVal).trim() : null,
          status: statusVal != null && String(statusVal).trim() !== '' ? String(statusVal).trim() : null,
          sendForApproval: sfaVal != null && String(sfaVal).trim() !== '' ? String(sfaVal).trim() : null,
          remarks: remarksVal != null && String(remarksVal).trim() !== '' ? String(remarksVal).trim() : null
        };

        const postResp = await fetch('/server/statutoryreg_function/statutory', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(payload)
        });
        const postData = await postResp.json().catch(() => ({}));
        if (!postResp.ok || postData.status !== 'success') {
          setError(postData.message || 'Failed to create statutory record with proof.');
          setTimeout(() => setError(''), 5000);
          return;
        }
        targetId = postData.data?.statutory?.id != null ? String(postData.data.statutory.id) : null;
      }

      setSuccess('Proof submission uploaded successfully.');
      fetchStatutoryData();
      window.dispatchEvent(new CustomEvent('statutoryDataUpdated', { detail: { id: targetId } }));
      setTimeout(() => setSuccess(''), 3000);
    } catch (err) {
      console.error('Inline proof upload error:', err);
      setError(`Proof upload failed: ${err.message}`);
      setTimeout(() => setError(''), 5000);
    } finally {
      setProofUploadingRowId(null);
    }
  };
  const handleAutofill = async (item, options = {}) => {
    // Open editable Excel-style modal on Autofill click
    setIsAutofillMode(true);
    setAutofillItem(item);
    await handleViewFormFile(item, true, options);
  };

  // Build draft workbook from formmaster template + table data (used by Save and by View Draft File generate-and-download)
  const buildDraftWorkbook = (opts) => {
    const {
      currentItem,
      templateWb,
      headersToUse,
      headerRowIndex,
      dataStartIndex,
      parsedFormHeader,
      headerFormData,
      formTableData,
      hasSubColumns,
      formFileName,
      allowTemplateFallback = true
    } = opts;
    const formHeader = parsedFormHeader || null;
    const tableHeaders = headersToUse || [];
    const normalizeHeaderLookup = (txt) =>
      String(txt || '')
        .replace(/\r?\n/g, ' ')
        .replace(/\s+/g, ' ')
        .trim()
        .toLowerCase()
        .replace(/[^a-z0-9 ]/g, '');
    const getRowValueForHeader = (row, header, headerIndex, allHeaders) => {
      if (!row || typeof row !== 'object') return '';
      if (Object.prototype.hasOwnProperty.call(row, header)) return row[header];

      const target = normalizeHeaderLookup(header);
      if (!target) return '';

      const rowKeys = Object.keys(row);
      const exactMatches = rowKeys.filter((k) => normalizeHeaderLookup(k) === target);
      if (exactMatches.length === 1) return row[exactMatches[0]];
      if (exactMatches.length > 1) {
        const sameHeaderOccurIndex =
          allHeaders.slice(0, headerIndex + 1).filter((h) => normalizeHeaderLookup(h) === target).length - 1;
        const pick = exactMatches[Math.min(Math.max(sameHeaderOccurIndex, 0), exactMatches.length - 1)];
        return row[pick];
      }

      const fuzzyMatch = rowKeys.find((k) => {
        const n = normalizeHeaderLookup(k);
        return n && (n.includes(target) || target.includes(n));
      });
      if (fuzzyMatch) return row[fuzzyMatch];
      return '';
    };
    const useFormFileTemplate = templateWb && headerRowIndex >= 0 && dataStartIndex >= 0 && headersToUse.length > 0;
    let effectiveDataStartIndex = dataStartIndex;
    let wb;
    if (useFormFileTemplate) {
      wb = templateWb;
      const firstSheetName = wb.SheetNames[0];
      const ws = wb.Sheets[firstSheetName];
      const writeCellPreserveStyle = (ref, rawValue) => {
        const prev = ws[ref] || {};
        if (rawValue == null || rawValue === '') {
          ws[ref] = { ...prev, t: 's', v: '' };
          return;
        }
        const isNum = typeof rawValue === 'number' || (typeof rawValue === 'string' && /^-?\d+(\.\d+)?$/.test(String(rawValue).trim()));
        if (isNum) {
          ws[ref] = { ...prev, t: 'n', v: Number(rawValue) };
        } else {
          ws[ref] = { ...prev, t: 's', v: String(rawValue) };
        }
      };
      const effectiveFormHeader = parsedFormHeader || formHeader;
      const isFormA = isFormAMusterRollContext(effectiveFormHeader) ||
        /\bform\s*[-"']?\s*a\b/i.test(String(currentItem?.formName || currentItem?.FormName || '')) ||
        /\bma\b.*\bform\b.*\ba\b/i.test(String(formFileName || ''));
      if (effectiveFormHeader?.fields && effectiveFormHeader.fields.length > 0) {
        for (let r = 0; r < headerRowIndex; r++) {
          for (let c = 0; c < 20; c++) {
            const ref = XLSX.utils.encode_cell({ r, c });
            const cell = ws[ref];
            if (!cell || cell.v == null) continue;
            const cellStr = String(cell.v).trim();
            for (const field of effectiveFormHeader.fields) {
              if (effectiveFormHeader.form12MergedHeaderCell && String(field.key || '').startsWith('form12_header_')) {
                continue;
              }
              if (effectiveFormHeader.form10MergedHeaderCell && String(field.key || '').startsWith('form10_header_')) {
                continue;
              }
              const label = (field.label || '').trim();
              if (!label) continue;
              if (cellStr === label || cellStr.startsWith(label) || label.startsWith(cellStr)) {
                const value = headerFormData[field.key] ?? field.value ?? '';
                const nextRef = XLSX.utils.encode_cell({ r, c: c + 1 });
                writeCellPreserveStyle(nextRef, String(value));
                break;
              }
            }
          }
        }
      }
      if (effectiveFormHeader?.form12MergedHeaderCell) {
        const { r, c } = effectiveFormHeader.form12MergedHeaderCell;
        const vF = String(headerFormData.form12_header_factory ?? '').trim();
        const vC = String(headerFormData.form12_header_contractor ?? '').trim();
        const vR = String(headerFormData.form12_header_registration ?? '').trim();
        const combined = `Name and Address of the Factory: ${vF}\nName and Address of the Contractor: ${vC}\nRegistration No: ${vR}`;
        const ref = XLSX.utils.encode_cell({ r, c });
        writeCellPreserveStyle(ref, combined);
      }
      if (effectiveFormHeader?.form10MergedHeaderCell) {
        const { r, c } = effectiveFormHeader.form10MergedHeaderCell;
        const vF = String(headerFormData.form10_header_factory ?? '').trim();
        const vC = String(headerFormData.form10_header_contractor ?? '').trim();
        const combined = `Name and Address of the Factory: ${vF}\nName and Address of the Contractor: ${vC}`;
        const ref = XLSX.utils.encode_cell({ r, c });
        writeCellPreserveStyle(ref, combined);
      }
      if (effectiveFormHeader?.form10MonthEndingCell) {
        const me = effectiveFormHeader.form10MonthEndingCell;
        const vM = String(headerFormData.form10_header_month_ending ?? '').trim();
        const ref = XLSX.utils.encode_cell({ r: me.r, c: me.c });
        const v =
          me.writeMode === 'full'
            ? (vM ? `Month ending: ${vM}` : 'Month ending: ')
            : vM;
        writeCellPreserveStyle(ref, String(v));
      }
      if (effectiveFormHeader?.festivalGrid?.approvalProceedings) {
        const ap = effectiveFormHeader.festivalGrid.approvalProceedings;
        const v = headerFormData[ap.key] != null && headerFormData[ap.key] !== ''
          ? headerFormData[ap.key]
          : (ap.initialValue != null ? ap.initialValue : '');
        const writeAt = (r, c) => {
          if (r == null || c == null || c < 0) return;
          const ref = XLSX.utils.encode_cell({ r, c });
          writeCellPreserveStyle(ref, String(v));
        };
        if (ap.labelRow != null && ap.valueCol != null) {
          writeAt(ap.labelRow, ap.valueCol);
        } else {
          searchAp: for (let r = 0; r < headerRowIndex; r++) {
            for (let c = 0; c < 80; c++) {
              const ref = XLSX.utils.encode_cell({ r, c });
              const cell = ws[ref];
              if (!cell || cell.v == null) continue;
              const cellStr = String(cell.v).trim();
              const apLabel = (ap.label || '').trim();
              if (!apLabel) continue;
              if (cellStr === apLabel || cellStr.startsWith(apLabel) || apLabel.startsWith(cellStr)) {
                writeAt(r, c + 1);
                break searchAp;
              }
            }
          }
        }
      }
      if (effectiveFormHeader?.festivalGrid?.keys && Array.isArray(effectiveFormHeader.festivalGrid.slotCols) && effectiveFormHeader.festivalGrid.dataRow != null) {
        const { dataRow, slotCols, keys, initialValues } = effectiveFormHeader.festivalGrid;
        keys.forEach((key, i) => {
          const c = slotCols[i];
          if (c == null || c < 0) return;
          const value =
            headerFormData[key] != null && headerFormData[key] !== ''
              ? headerFormData[key]
              : (Array.isArray(initialValues) && initialValues[i] != null ? initialValues[i] : '');
          const ref = XLSX.utils.encode_cell({ r: dataRow, c });
          writeCellPreserveStyle(ref, String(value));
        });
      }
      const sheetArr = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
      const headerToCol = new Map();
      const normalizeHeader = (txt) =>
        String(txt || '')
          .replace(/\r?\n/g, ' ')
          .replace(/\s+/g, ' ')
          .trim()
          .toLowerCase();
      const splitHeader = (txt) => {
        const normalized = normalizeHeader(txt);
        if (!normalized) return { main: '', sub: '' };
        if (normalized.includes('_')) {
          const firstUnderscore = normalized.indexOf('_');
          return {
            main: normalized.slice(0, firstUnderscore).trim(),
            sub: normalized.slice(firstUnderscore + 1).trim()
          };
        }
        return { main: normalized, sub: '' };
      };
      const merges = ws['!merges'] || [];
      const getRawCellText = (r, c) => {
        if (r < 0 || c < 0) return '';
        const ref = XLSX.utils.encode_cell({ r, c });
        const cell = ws[ref];
        if (!cell || cell.v == null) return '';
        return String(cell.v).trim();
      };
      const getMergedAwareCellText = (r, c) => {
        const direct = getRawCellText(r, c);
        if (direct) return direct;
        for (let i = 0; i < merges.length; i++) {
          const m = merges[i];
          if (!m || !m.s || !m.e) continue;
          if (r >= m.s.r && r <= m.e.r && c >= m.s.c && c <= m.e.c) {
            const topLeft = getRawCellText(m.s.r, m.s.c);
            if (topLeft) return topLeft;
          }
        }
        return '';
      };

      // Form A (Maternity Benefit Rules) templates use rotated/merged headers; generic header text mapping misaligns.
      // Anchor on the real "S.No" cell in the template and map headers sequentially across columns.
      if (isFormA) {
        let snoAnchor = null;
        for (let r = 0; r < Math.min(60, sheetArr.length); r++) {
          for (let c = 0; c < 120; c++) {
            const t = normalizeHeader(getMergedAwareCellText(r, c));
            if (t && (t === 's.no' || t === 's no' || t.includes('s.no') || t.includes('serial'))) {
              snoAnchor = { r, c };
              break;
            }
          }
          if (snoAnchor) break;
        }
        const startCol = snoAnchor?.c != null ? snoAnchor.c : 0;
        const forcedDataStart = snoAnchor?.r != null ? (snoAnchor.r + 1) : effectiveDataStartIndex;
        effectiveDataStartIndex = forcedDataStart;
        for (let j = 0; j < headersToUse.length; j++) {
          headerToCol.set(headersToUse[j], startCol + j);
        }
      }

      const maxCols = Math.max(
        (sheetArr[headerRowIndex] || []).length,
        (sheetArr[headerRowIndex + 1] || []).length,
        headersToUse.length,
        20
      );
      if (!isFormA) {
        // Prefer exact mapping from (main header + sub header) so grouped headers (Form X) keep the original column alignment.
        for (let c = 0; c < maxCols; c++) {
          const mainHeader = normalizeHeader(getMergedAwareCellText(headerRowIndex, c));
          const subHeader = normalizeHeader(getMergedAwareCellText(headerRowIndex + 1, c));
          const combinedHeader = mainHeader && subHeader ? `${mainHeader}_${subHeader}` : '';
          let bestHeader = null;
          let bestScore = 0;
          for (let j = 0; j < headersToUse.length; j++) {
            const header = normalizeHeader(headersToUse[j] || '');
            if (!header) continue;
            const { main, sub } = splitHeader(header);
            const isCombinedExact = combinedHeader && header === combinedHeader;
            const isMainSubMatch = mainHeader && main === mainHeader && (!sub || (subHeader && sub === subHeader));
            const isHeaderContains = (mainHeader && header.includes(mainHeader)) || (subHeader && header.includes(subHeader));
            if (isCombinedExact || isMainSubMatch || isHeaderContains) {
              const score = isCombinedExact ? 2000 : isMainSubMatch ? 1500 : Math.min(header.length, (combinedHeader || mainHeader || subHeader).length);
              if (score > bestScore) {
                bestScore = score;
                bestHeader = headersToUse[j];
              }
            }
          }
          if (bestHeader != null && !headerToCol.has(bestHeader)) {
            headerToCol.set(bestHeader, c);
          }
        }
        const mappedCols = Array.from(headerToCol.values());
        const fallbackStartCol = mappedCols.length > 0 ? Math.min(...mappedCols) : 0;
        for (let j = 0; j < headersToUse.length; j++) {
          const header = headersToUse[j];
          if (!headerToCol.has(header)) headerToCol.set(header, fallbackStartCol + j);
        }
      }
      for (let i = 0; i < formTableData.length; i++) {
        const row = formTableData[i];
        for (let j = 0; j < headersToUse.length; j++) {
          const header = headersToUse[j];
          const value = getRowValueForHeader(row, header, j, headersToUse);
          const colIdx = headerToCol.has(header) ? headerToCol.get(header) : j;
          const cellRef = XLSX.utils.encode_cell({ r: effectiveDataStartIndex + i, c: colIdx });
          writeCellPreserveStyle(cellRef, value);
        }
      }
      // Do not overwrite the header row – keep the original formmaster template layout exactly. Only data rows are filled above.
    } else {
      if (!allowTemplateFallback) {
        throw new Error('Original form template is required to generate draft.');
      }
      wb = XLSX.utils.book_new();
      const headerRows = [];
      if (formHeader?.title) headerRows.push([formHeader.title]);
      if (formHeader?.subtitle) headerRows.push([formHeader.subtitle]);
      if (formHeader?.reference) headerRows.push([formHeader.reference]);
      if (formHeader?.fields && formHeader.fields.length > 0) {
        headerRows.push([]);
        const headerFieldRow = [];
        const headerValueRow = [];
        formHeader.fields.forEach((field) => {
          headerFieldRow.push(field.label);
          headerValueRow.push(headerFormData[field.key] || field.value || '');
        });
        headerRows.push(headerFieldRow);
        headerRows.push(headerValueRow);
      }
      headerRows.push([]);
      const effectiveTableHeaders = headersToUse.length ? headersToUse : tableHeaders;
      const formNameLowerFallback = (currentItem?.formName || formHeader?.title || '').toString().toLowerCase();
      const isFormUFallback = formNameLowerFallback.includes('form u') || formNameLowerFallback.includes('form u ');
      const headersForRow = isFormUFallback && FORM_U_FORMAT2_HEADERS.length > 0
        ? effectiveTableHeaders.map((_, idx) => FORM_U_FORMAT2_HEADERS[idx] ?? effectiveTableHeaders[idx])
        : effectiveTableHeaders;
      headerRows.push(headersForRow);
      const tableRows = formTableData.map((row) =>
        effectiveTableHeaders.map((header, idx) => {
          const v = getRowValueForHeader(row, header, idx, effectiveTableHeaders);
          return v != null ? v : '';
        })
      );
      const allData = [...headerRows, ...tableRows];
      const ws = XLSX.utils.aoa_to_sheet(allData);
      XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    }
    const fileName = formFileName ||
      formHeader?.title?.replace(/[^a-zA-Z0-9]/g, '_') ||
      `Form_${Date.now()}.xlsx`;
    const excelBuffer = XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
    const blob = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    return { blob, fileName };
  };

  const extractMappedRowsFromSavedDraft = ({
    draftWorkbook,
    templateHeaders,
    templateHeaderRowIndex,
    templateDataStartIndex,
    hasSubColumns
  }) => {
    if (!draftWorkbook || !Array.isArray(templateHeaders) || templateHeaders.length === 0) return [];
    const parsedDraft = parseExcelForm(draftWorkbook);
    const draftHeaders = Array.isArray(parsedDraft.headers) ? parsedDraft.headers : [];
    if (!draftHeaders.length) return [];
    const draftSheetName = draftWorkbook.SheetNames[0];
    const draftWs = draftWorkbook.Sheets[draftSheetName];
    if (!draftWs) return [];
    const draftRowsRaw = XLSX.utils.sheet_to_json(draftWs, { header: 1, defval: '' });
    const draftDataStart = parsedDraft.dataStartIndex != null && parsedDraft.dataStartIndex >= 0
      ? parsedDraft.dataStartIndex
      : 0;
    const fallbackTargetStart = templateDataStartIndex != null && templateDataStartIndex >= 0
      ? templateDataStartIndex
      : (templateHeaderRowIndex >= 0 ? templateHeaderRowIndex + (hasSubColumns ? 2 : 1) : 0);
    const startIndex = Math.max(0, Math.min(draftDataStart, fallbackTargetStart));
    const draftDataRows = draftRowsRaw.slice(startIndex).filter((row) =>
      Array.isArray(row) && row.some((cell) => String(cell ?? '').trim() !== '')
    );
    if (!draftDataRows.length) return [];
    const rows = draftDataRows.map((rowArr) => {
      const rowObj = {};
      for (let i = 0; i < draftHeaders.length; i++) {
        rowObj[draftHeaders[i]] = rowArr[i] != null ? rowArr[i] : '';
      }
      return rowObj;
    });
    return rows;
  };

  const extractRowMatrixFromSavedDraft = ({
    draftWorkbook,
    templateHeaderRowIndex,
    templateDataStartIndex,
    hasSubColumns
  }) => {
    if (!draftWorkbook) return [];
    const parsedDraft = parseExcelForm(draftWorkbook);
    const draftSheetName = draftWorkbook.SheetNames[0];
    const draftWs = draftWorkbook.Sheets[draftSheetName];
    if (!draftWs) return [];
    const draftRowsRaw = XLSX.utils.sheet_to_json(draftWs, { header: 1, defval: '' });
    const draftDataStart = parsedDraft.dataStartIndex != null && parsedDraft.dataStartIndex >= 0
      ? parsedDraft.dataStartIndex
      : 0;
    const fallbackTargetStart = templateDataStartIndex != null && templateDataStartIndex >= 0
      ? templateDataStartIndex
      : (templateHeaderRowIndex >= 0 ? templateHeaderRowIndex + (hasSubColumns ? 2 : 1) : 0);
    const startIndex = Math.max(0, Math.min(draftDataStart, fallbackTargetStart));
    return draftRowsRaw.slice(startIndex).filter((row) =>
      Array.isArray(row) && row.some((cell) => String(cell ?? '').trim() !== '')
    );
  };

  const buildFormAWorkbookWithTemplateStyles = async ({
    templateArrayBuffer,
    headersToUse,
    mappedData,
    mappedRowMatrix,
    templateFileId,
    parsedFormHeader,
    formFileName
  }) => {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(templateArrayBuffer);
    const worksheet = workbook.worksheets[0];
    if (!worksheet) {
      throw new Error('Template worksheet not found.');
    }
    const normalize = (txt) =>
      String(txt || '')
        .replace(/\r?\n/g, ' ')
        .replace(/\s+/g, ' ')
        .trim()
        .toLowerCase();
    const excelCellValueToString = (val) => {
      if (val == null) return '';
      if (typeof val === 'string' || typeof val === 'number' || typeof val === 'boolean') return String(val);
      if (val instanceof Date) return val.toISOString();
      if (typeof val === 'object') {
        if (Array.isArray(val.richText)) return val.richText.map((rt) => rt?.text || '').join('');
        if (val.text != null) return String(val.text);
        if (val.result != null) return String(val.result);
        if (val.hyperlink != null && val.text != null) return String(val.text);
        return '';
      }
      return '';
    };
    const getRowValueForHeader = (row, header) => {
      if (!row || typeof row !== 'object') return '';
      if (Object.prototype.hasOwnProperty.call(row, header)) return row[header];
      const target = normalize(header);
      if (!target) return '';
      const rowKeys = Object.keys(row);
      const exact = rowKeys.find((k) => normalize(k) === target);
      if (exact) return row[exact];
      const fuzzy = rowKeys.find((k) => {
        const nk = normalize(k);
        return nk && (nk.includes(target) || target.includes(nk));
      });
      return fuzzy ? row[fuzzy] : '';
    };

    let snoAnchor = null;
    const maxScanRows = Math.max(60, worksheet.rowCount + 5);
    const maxScanCols = 140;
    for (let r = 1; r <= maxScanRows; r++) {
      for (let c = 1; c <= maxScanCols; c++) {
        const cell = worksheet.getCell(r, c);
        const raw = excelCellValueToString(cell?.value);
        const t = normalize(raw);
        if (t && (t === 's.no' || t === 's no' || t.includes('s.no') || t.includes('serial'))) {
          snoAnchor = { r, c };
          break;
        }
      }
      if (snoAnchor) break;
    }
    if (!snoAnchor) {
      throw new Error('Could not locate Form A table anchor (S.No) in original template.');
    }

    // Apply strict safeguards for all Form A templates (template ID can change after re-upload).
    const isStrictFormATemplate = true;
    const startRow = isStrictFormATemplate ? (snoAnchor.r + 3) : (snoAnchor.r + 1);
    const startCol = snoAnchor.c;
    const useMatrix = Array.isArray(mappedRowMatrix) && mappedRowMatrix.length > 0;
    const cleanedMatrixRows = useMatrix
      ? mappedRowMatrix.filter((row) => {
          if (!Array.isArray(row)) return false;
          const cells = row.map((v) => String(v ?? '').trim()).filter(Boolean);
          if (cells.length === 0) return false;
          if (!isStrictFormATemplate) return true;
          const hasAlpha = cells.some((v) => /[a-z]/i.test(v));
          return hasAlpha;
        })
      : [];
    const cleanedObjectRows = !useMatrix && Array.isArray(mappedData)
      ? mappedData.filter((row) => {
          if (!row || typeof row !== 'object') return false;
          const vals = Object.values(row).map((v) => String(v ?? '').trim()).filter(Boolean);
          if (vals.length === 0) return false;
          if (!isStrictFormATemplate) return true;
          const hasAlpha = vals.some((v) => /[a-z]/i.test(v));
          return hasAlpha;
        })
      : [];
    const totalRows = useMatrix ? cleanedMatrixRows.length : cleanedObjectRows.length;
    for (let i = 0; i < totalRows; i++) {
      const rowObj = cleanedObjectRows[i] || {};
      const rowArr = useMatrix ? (cleanedMatrixRows[i] || []) : null;
      const totalCols = useMatrix ? rowArr.length : headersToUse.length;
      for (let j = 0; j < totalCols; j++) {
        const header = headersToUse[j];
        const value = useMatrix ? rowArr[j] : getRowValueForHeader(rowObj, header);
        const cell = worksheet.getCell(startRow + i, startCol + j);
        if (value == null || value === '') {
          // For strict Form A template, never blank existing cells — keep original header text/layout untouched.
          if (!isStrictFormATemplate) {
            cell.value = '';
          }
        } else if (typeof value === 'number' || (typeof value === 'string' && /^-?\d+(\.\d+)?$/.test(String(value).trim()))) {
          cell.value = Number(value);
        } else {
          cell.value = String(value);
        }
      }
    }

    const out = await workbook.xlsx.writeBuffer();
    const fileName = formFileName ||
      parsedFormHeader?.title?.replace(/[^a-zA-Z0-9]/g, '_') ||
      `Form_${Date.now()}.xlsx`;
    const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    return { blob, fileName };
  };

  const buildFormUWorkbookWithTemplateStyles = async ({
    templateArrayBuffer,
    headersToUse,
    mappedData,
    mappedRowMatrix,
    parsedFormHeader,
    formFileName
  }) => {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(templateArrayBuffer);
    const worksheet = workbook.worksheets[0];
    if (!worksheet) throw new Error('Template worksheet not found.');

    const normalize = (txt) =>
      String(txt || '')
        .replace(/\r?\n/g, ' ')
        .replace(/\s+/g, ' ')
        .trim()
        .toLowerCase();
    const excelCellValueToString = (val) => {
      if (val == null) return '';
      if (typeof val === 'string' || typeof val === 'number' || typeof val === 'boolean') return String(val);
      if (val instanceof Date) return val.toISOString();
      if (typeof val === 'object') {
        if (Array.isArray(val.richText)) return val.richText.map((rt) => rt?.text || '').join('');
        if (val.text != null) return String(val.text);
        if (val.result != null) return String(val.result);
        return '';
      }
      return '';
    };
    const getRowValueForHeader = (row, header) => {
      if (!row || typeof row !== 'object') return '';
      if (Object.prototype.hasOwnProperty.call(row, header)) return row[header];
      const target = normalize(header);
      if (!target) return '';
      const rowKeys = Object.keys(row);
      const exact = rowKeys.find((k) => normalize(k) === target);
      if (exact) return row[exact];
      const fuzzy = rowKeys.find((k) => {
        const nk = normalize(k);
        return nk && (nk.includes(target) || target.includes(nk));
      });
      return fuzzy ? row[fuzzy] : '';
    };
    const pickObjectValue = (row, patterns, fallbackIndex = null) => {
      if (!row || typeof row !== 'object' || Array.isArray(row)) return '';
      const keys = Object.keys(row);
      const norm = (s) => normalize(String(s || '').replace(/[^a-z0-9]/gi, ' '));
      for (let pi = 0; pi < patterns.length; pi++) {
        const p = patterns[pi];
        const hit = keys.find((k) => p.test(norm(k)));
        if (hit) {
          const v = row[hit];
          if (v != null && String(v).trim() !== '') return v;
        }
      }
      if (fallbackIndex != null && fallbackIndex >= 0 && fallbackIndex < keys.length) {
        return row[keys[fallbackIndex]];
      }
      return '';
    };
    const toFormUOrderedValues = (row, rowIndex) => {
      if (Array.isArray(row)) return row;
      const ordered = [
        pickObjectValue(row, [/\bs\.?\s*no\b/, /serial/], 0),
        pickObjectValue(row, [/name.*employee/, /employee.*name/, /name of the employee/], 1),
        pickObjectValue(row, [/employee.*identification/, /employee.*id/, /emp.*id/], 2),
        pickObjectValue(row, [/\bgender\b/, /\bsex\b/], 3),
        pickObjectValue(row, [/father.*spouse/, /spouse.*father/, /father.*name/, /spouse.*name/], 4),
        pickObjectValue(row, [/date.*birth/, /\bdob\b/], 5),
        pickObjectValue(row, [/date.*joining/, /\bdoj\b/], 6),
        pickObjectValue(row, [/designation/, /position/], 7),
        pickObjectValue(row, [/present.*address/], 8),
        pickObjectValue(row, [/permanent.*address/], 9),
        pickObjectValue(row, [/employee.*aadha?r/, /aadha?r/], 10),
        pickObjectValue(row, [/date.*notice.*pregnancy/, /notice.*pregnancy/], 11),
        pickObjectValue(row, [/date.*notice.*delivery/, /notice.*delivery/], 12),
        pickObjectValue(row, [/date.*proof.*birth/, /proof.*birth/], 13),
        pickObjectValue(row, [/date.*proof.*death/, /proof.*death/], 14),
        pickObjectValue(row, [/period.*wages/, /wages.*period/], 15),
        pickObjectValue(row, [/bank.*a\/?c/, /bank.*account/], 16),
        pickObjectValue(row, [/\bphoto\b/], 17)
      ];
      if (ordered[0] == null || String(ordered[0]).trim() === '') ordered[0] = rowIndex + 1;
      return ordered;
    };

    let snoAnchor = null;
    const maxScanRows = Math.max(80, worksheet.rowCount + 5);
    for (let r = 1; r <= maxScanRows; r++) {
      for (let c = 1; c <= 180; c++) {
        const raw = excelCellValueToString(worksheet.getCell(r, c)?.value);
        const t = normalize(raw);
        if (t && (t === 's.no' || t === 's no' || t.includes('s.no') || t.includes('serial'))) {
          snoAnchor = { r, c };
          break;
        }
      }
      if (snoAnchor) break;
    }
    if (!snoAnchor) throw new Error('Could not locate Form U table anchor (S.No) in original template.');

    // Form U has one numbering row (1..n) below the main header, so actual data starts on next row.
    const startRow = snoAnchor.r + 2;
    const startCol = snoAnchor.c;
    const headerRow = snoAnchor.r;
    const headerColMap = new Map();
    for (let c = Math.max(1, startCol - 1); c <= 220; c++) {
      const txt = normalize(excelCellValueToString(worksheet.getCell(headerRow, c)?.value));
      if (txt) headerColMap.set(c, txt);
    }
    const resolveFormUColForHeader = (header) => {
      const h = normalize(header);
      if (!h) return null;
      const matchers = [
        { key: 'sno', re: /\bs\.?\s*no\b|serial/ },
        { key: 'name_employee', re: /name\s+of\s+the\s+employee|name\s+of\s+employee/ },
        { key: 'employee_id', re: /employee\s+identification|employee\s+id/ },
        { key: 'gender', re: /\bgender\b/ },
        { key: 'father_spouse', re: /father.*spouse|spouse.*father/ },
        { key: 'dob', re: /date\s+of\s+birth/ },
        { key: 'doj', re: /date\s+of\s+joining/ },
        { key: 'designation', re: /designation/ },
        { key: 'present_address', re: /present\s+address/ },
        { key: 'permanent_address', re: /permanent\s+address/ },
        { key: 'aadhaar', re: /aadha?r/ },
        { key: 'wages_period', re: /period\s+of\s+wages/ },
        { key: 'bank', re: /bank\s*a\/?c/ },
        { key: 'photo', re: /\bphoto\b/ }
      ];
      let targetMatcher = matchers.find((m) => m.re.test(h));
      if (!targetMatcher) return null;
      for (const [col, txt] of headerColMap.entries()) {
        if (targetMatcher.re.test(txt)) return col;
      }
      return null;
    };
    const formUFieldMatchers = [
      /\bs\.?\s*no\b|serial/,                              // 0
      /name\s+of\s+the\s+employee|name\s+of\s+employee/,  // 1
      /employee\s+identification|employee\s+id/,          // 2
      /\bgender\b/,                                        // 3
      /father.*spouse|spouse.*father/,                    // 4
      /date\s+of\s+birth/,                                 // 5
      /date\s+of\s+joining/,                               // 6
      /designation/,                                       // 7
      /present\s+address/,                                 // 8
      /permanent\s+address/,                               // 9
      /aadha?r/,                                           // 10
      /notice.*pregnancy|pregnancy.*notice/,              // 11
      /notice.*delivery|delivery.*notice/,                // 12
      /proof.*birth|birth.*proof/,                        // 13
      /proof.*death|death.*proof/,                        // 14
      /period\s+of\s+wages|wages\s+period/,               // 15
      /bank\s*a\/?c|bank\s+account/,                      // 16
      /\bphoto\b/                                          // 17
    ];
    const formUFieldCols = formUFieldMatchers.map((re, idx) => {
      for (const [col, txt] of headerColMap.entries()) {
        if (re.test(txt)) return col;
      }
      return startCol + idx;
    });
    const hasMeaningfulRow = (row) => {
      if (!Array.isArray(row)) return false;
      const cells = row.map((v) => String(v ?? '').trim()).filter(Boolean);
      if (cells.length === 0) return false;
      // Ignore pure serial/index-like rows (e.g. "1,2,3,4,5...").
      const alphaCount = cells.filter((v) => /[a-z]/i.test(v)).length;
      if (alphaCount > 0) return true;
      const numericOnly = cells.every((v) => /^-?\d+(\.\d+)?$/.test(v));
      return !numericOnly && cells.length >= 3;
    };
    const matrixRows = Array.isArray(mappedRowMatrix)
      ? mappedRowMatrix.filter((row) => hasMeaningfulRow(row))
      : [];
    const mappedArrayRows = Array.isArray(mappedData)
      ? mappedData
          .map((row, idx) => toFormUOrderedValues(row, idx))
          .filter((row) => hasMeaningfulRow(row))
      : [];
    const useMatrix = matrixRows.length > 0 || mappedArrayRows.length > 0;
    const isMeaningfulObjectRow = (row) => {
      if (!row || typeof row !== 'object' || Array.isArray(row)) return false;
      const vals = Object.values(row).map((v) => String(v ?? '').trim()).filter(Boolean);
      if (vals.length === 0) return false;
      const hasAlpha = vals.some((v) => /[a-z]/i.test(v));
      if (hasAlpha) return true;
      const numericOnly = vals.every((v) => /^-?\d+(\.\d+)?$/.test(v));
      return !numericOnly;
    };
    const matrixSource = matrixRows.length > 0 ? matrixRows : mappedArrayRows;
    const totalRows = matrixSource.length;

    for (let i = 0; i < totalRows; i++) {
      const rowArr = matrixSource[i] || [];
      const totalCols = rowArr.length;
      for (let j = 0; j < totalCols; j++) {
        let value = rowArr[j];
        if (value == null || value === '') continue; // keep template untouched for empty values
        const targetCol = j < formUFieldCols.length ? formUFieldCols[j] : (startCol + j);
        const cell = worksheet.getCell(startRow + i, targetCol);
        // Keep Form U output stable with template formatting: write as text to avoid Excel auto-date (e.g. 6 -> 01/06/1900).
        if ((targetCol === formUFieldCols[0] || j === 0) && (typeof value === 'number' || (typeof value === 'string' && /^-?\d+(\.\d+)?$/.test(String(value).trim())))) {
          cell.value = Number(value); // S.No can stay numeric
        } else {
          cell.value = String(value);
        }
      }
    }

    const out = await workbook.xlsx.writeBuffer();
    const fileName = formFileName ||
      parsedFormHeader?.title?.replace(/[^a-zA-Z0-9]/g, '_') ||
      `Form_${Date.now()}.xlsx`;
    const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    return { blob, fileName };
  };

  const buildForm12WorkbookWithTemplateStyles = async ({
    templateArrayBuffer,
    mappedData,
    mappedRowMatrix,
    headersToUse,
    parsedFormHeader,
    formFileName
  }) => {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(templateArrayBuffer);
    const worksheet = workbook.worksheets[0];
    if (!worksheet) throw new Error('Template worksheet not found.');

    const normalize = (txt) =>
      String(txt || '')
        .replace(/\r?\n/g, ' ')
        .replace(/\s+/g, ' ')
        .trim()
        .toLowerCase();
    const excelCellValueToString = (val) => {
      if (val == null) return '';
      if (typeof val === 'string' || typeof val === 'number' || typeof val === 'boolean') return String(val);
      if (val instanceof Date) return val.toISOString();
      if (typeof val === 'object') {
        if (Array.isArray(val.richText)) return val.richText.map((rt) => rt?.text || '').join('');
        if (val.text != null) return String(val.text);
        if (val.result != null) return String(val.result);
      }
      return '';
    };

    let snoAnchor = null;
    const maxScanRows = Math.max(80, worksheet.rowCount + 5);
    for (let r = 1; r <= maxScanRows; r++) {
      for (let c = 1; c <= 220; c++) {
        const t = normalize(excelCellValueToString(worksheet.getCell(r, c)?.value));
        if (t && (t === 's.no' || t === 's no' || t.includes('s.no') || t.includes('serial'))) {
          snoAnchor = { r, c };
          break;
        }
      }
      if (snoAnchor) break;
    }
    if (!snoAnchor) throw new Error('Could not locate Form 12 table anchor (S.No).');

    const headerRow = snoAnchor.r;
    const startCol = snoAnchor.c;
    const headerColMap = new Map();
    for (let c = Math.max(1, startCol - 1); c <= 240; c++) {
      const txt = normalize(excelCellValueToString(worksheet.getCell(headerRow, c)?.value));
      if (txt) headerColMap.set(c, txt);
    }
    const nextRowValues = [];
    for (let c = startCol; c <= startCol + 40; c++) {
      nextRowValues.push(normalize(excelCellValueToString(worksheet.getCell(headerRow + 1, c)?.value)));
    }
    const numericMarkers = nextRowValues.filter((v) => /^\d+$/.test(v)).length;
    const startRow = numericMarkers >= 4 ? (headerRow + 2) : (headerRow + 1);

    const fieldMatchers = [
      /\bs\.?\s*no\b|serial/,
      /name.*employee|employee.*name|worker.*name/,
      /father|husband|spouse/,
      /\bage\b|date.*birth|\bdob\b/,
      /sex|gender/,
      /designation|nature.*work|occupation/,
      /date.*joining|date.*employment/,
      /wage|salary|rate.*wage/,
      /attendance|days.*worked|days.*present/,
      /overtime|ot/,
      /remarks|signature/
    ];
    const fieldCols = fieldMatchers.map((re, idx) => {
      for (const [col, txt] of headerColMap.entries()) {
        if (re.test(txt)) return col;
      }
      return startCol + idx;
    });
    const pickObjectValue = (row, patterns, fallbackIndex = null) => {
      if (!row || typeof row !== 'object' || Array.isArray(row)) return '';
      const keys = Object.keys(row);
      const norm = (s) => normalize(String(s || '').replace(/[^a-z0-9]/gi, ' '));
      for (let pi = 0; pi < patterns.length; pi++) {
        const p = patterns[pi];
        const hit = keys.find((k) => p.test(norm(k)));
        if (hit) {
          const v = row[hit];
          if (v != null && String(v).trim() !== '') return v;
        }
      }
      if (fallbackIndex != null && fallbackIndex >= 0 && fallbackIndex < keys.length) {
        return row[keys[fallbackIndex]];
      }
      return '';
    };
    const pickValueFromHeaders = (row, headerList, patterns, fallbackIndex = null) => {
      if (!row || typeof row !== 'object' || Array.isArray(row)) return '';
      const headers = Array.isArray(headerList) ? headerList : [];
      for (let hi = 0; hi < headers.length; hi++) {
        const h = headers[hi];
        const nh = normalize(String(h || '').replace(/[^a-z0-9]/gi, ' '));
        const matched = patterns.some((p) => p.test(nh));
        if (!matched) continue;
        if (Object.prototype.hasOwnProperty.call(row, h)) {
          const v = row[h];
          if (v != null && String(v).trim() !== '') return v;
        }
      }
      return pickObjectValue(row, patterns, fallbackIndex);
    };
    const toForm12OrderedValues = (row, rowIndex) => {
      if (Array.isArray(row)) {
        const out = [...row];
        // Common bad parse: "2 rathnakumar" in first cell; split into serial + name.
        if (out[0] != null) {
          const s0 = String(out[0]).trim();
          const m = s0.match(/^(\d+)\s+(.+)$/);
          if (m && (!out[1] || String(out[1]).trim() === '')) {
            out[0] = m[1];
            out[1] = m[2];
          }
        }
        return [
          out[0] ?? '',
          out[1] ?? '',
          out[2] ?? '',
          out[3] ?? '',
          out[4] ?? '',
          out[5] ?? '',
          out[6] ?? '',
          out[7] ?? '',
          out[8] ?? ''
        ];
      }
      return [
        pickValueFromHeaders(row, headersToUse, [/\bserial\b/, /\bs\.?\s*no\b/], 0),
        pickValueFromHeaders(row, headersToUse, [/name.*worker/, /name.*employee/, /worker.*name/], 1),
        pickValueFromHeaders(row, headersToUse, [/worker.*identity/, /employee.*identification/, /employee.*id/], 2),
        pickValueFromHeaders(row, headersToUse, [/\bgender\b/, /\bsex\b/], 3),
        pickValueFromHeaders(row, headersToUse, [/father.*spouse/, /spouse.*name/, /father.*name/], 4),
        pickValueFromHeaders(row, headersToUse, [/date.*birth/, /\bdob\b/], 5),
        pickValueFromHeaders(row, headersToUse, [/present.*address/], 6),
        pickValueFromHeaders(row, headersToUse, [/permanent.*address/], 7),
        pickValueFromHeaders(row, headersToUse, [/aadha?r/], 8)
      ];
    };

    const rowLooksMeaningful = (row) => {
      if (Array.isArray(row)) {
        const cells = row.map((v) => String(v ?? '').trim()).filter(Boolean);
        if (cells.length === 0) return false;
        return cells.some((v) => /[a-z]/i.test(v)) || !cells.every((v) => /^-?\d+(\.\d+)?$/.test(v));
      }
      if (row && typeof row === 'object') {
        const vals = Object.values(row).map((v) => String(v ?? '').trim()).filter(Boolean);
        if (vals.length === 0) return false;
        return vals.some((v) => /[a-z]/i.test(v)) || !vals.every((v) => /^-?\d+(\.\d+)?$/.test(v));
      }
      return false;
    };
    const sourcePrimary = Array.isArray(mappedData) && mappedData.length > 0 ? mappedData : (Array.isArray(mappedRowMatrix) ? mappedRowMatrix : []);
    const sourceRows = sourcePrimary
      .filter((row) => rowLooksMeaningful(row))
      .map((row, idx) => {
        const out = toForm12OrderedValues(row, idx);
        if (out[0] == null || String(out[0]).trim() === '') out[0] = idx + 1;
        return out;
      });

    const writableCols = fieldCols.slice(0, 9);
    for (let i = 0; i < sourceRows.length; i++) {
      const row = sourceRows[i];
      // Clear existing template sample row values so every employee row follows the same model alignment.
      for (let wc = 0; wc < writableCols.length; wc++) {
        const clearCell = worksheet.getCell(startRow + i, writableCols[wc]);
        clearCell.value = '';
      }
      for (let j = 0; j < row.length; j++) {
        const value = row[j];
        if (value == null || value === '') continue;
        const targetCol = j < fieldCols.length ? fieldCols[j] : (startCol + j);
        const cell = worksheet.getCell(startRow + i, targetCol);
        if ((j === 0) && (typeof value === 'number' || (typeof value === 'string' && /^-?\d+(\.\d+)?$/.test(String(value).trim())))) {
          cell.value = Number(value);
        } else {
          cell.value = String(value);
        }
        // Keep Form 12 data rows normal-weight (template sample row can be bold).
        cell.font = { ...(cell.font || {}), bold: false };
        cell.alignment = { ...(cell.alignment || {}), horizontal: 'left' };
      }
    }

    const out = await workbook.xlsx.writeBuffer();
    const fileName = formFileName ||
      parsedFormHeader?.title?.replace(/[^a-zA-Z0-9]/g, '_') ||
      `Form_${Date.now()}.xlsx`;
    const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    return { blob, fileName };
  };

  const buildForm10WorkbookWithTemplateStyles = async ({
    templateArrayBuffer,
    mappedData,
    mappedRowMatrix,
    headersToUse,
    parsedHeaderRowIndex,
    parsedDataStartIndex,
    parsedFormHeader,
    formFileName
  }) => {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(templateArrayBuffer);
    const worksheet = workbook.worksheets[0];
    if (!worksheet) throw new Error('Template worksheet not found.');

    const normalize = (txt) =>
      String(txt || '')
        .replace(/\r?\n/g, ' ')
        .replace(/\s+/g, ' ')
        .trim()
        .toLowerCase();
    const excelCellValueToString = (val) => {
      if (val == null) return '';
      if (typeof val === 'string' || typeof val === 'number' || typeof val === 'boolean') return String(val);
      if (val instanceof Date) return val.toISOString();
      if (typeof val === 'object') {
        if (Array.isArray(val.richText)) return val.richText.map((rt) => rt?.text || '').join('');
        if (val.text != null) return String(val.text);
        if (val.result != null) return String(val.result);
      }
      return '';
    };

    const maxScanRows = Math.max(120, worksheet.rowCount + 10);
    const maxScanCols = 260;
    let headerRow = parsedHeaderRowIndex != null && parsedHeaderRowIndex >= 0 ? (parsedHeaderRowIndex + 1) : -1;
    if (headerRow < 1) throw new Error('Could not locate Form 10 header row from parser.');
    let startCol = 1;
    for (let c = 1; c <= maxScanCols; c++) {
      const t = normalize(excelCellValueToString(worksheet.getCell(headerRow, c)?.value));
      if (t) {
        startCol = c;
        break;
      }
    }

    // Prefer numeric marker row (1..N) near header for exact template column model.
    let markerRow = (parsedDataStartIndex != null && parsedDataStartIndex > 0) ? parsedDataStartIndex : -1;
    let markerCols = [];
    for (let r = Math.max(headerRow, markerRow > 0 ? markerRow - 1 : headerRow); r <= Math.min(headerRow + 5, maxScanRows); r++) {
      const cols = [];
      for (let c = startCol; c <= maxScanCols; c++) {
        const t = normalize(excelCellValueToString(worksheet.getCell(r, c)?.value));
        if (/^\d+$/.test(t)) cols.push(c);
      }
      if (cols.length > markerCols.length) {
        markerCols = cols;
        markerRow = r;
      }
    }
    const dataStartRow =
      parsedDataStartIndex != null && parsedDataStartIndex >= 0
        ? (parsedDataStartIndex + 1)
        : (markerCols.length >= 5 && markerRow > 0 ? (markerRow + 1) : (headerRow + 1));
    const orderedCols = markerCols.length >= 5
      ? markerCols
      : Array.from({ length: Math.max(12, Array.isArray(headersToUse) ? headersToUse.length : 12) }, (_, i) => startCol + i);

    const rowLooksMeaningful = (row) => {
      if (Array.isArray(row)) {
        const vals = row.map((v) => String(v ?? '').trim()).filter(Boolean);
        if (vals.length === 0) return false;
        return vals.some((v) => /[a-z]/i.test(v)) || !vals.every((v) => /^-?\d+(\.\d+)?$/.test(v));
      }
      if (row && typeof row === 'object') {
        const vals = Object.values(row).map((v) => String(v ?? '').trim()).filter(Boolean);
        if (vals.length === 0) return false;
        return vals.some((v) => /[a-z]/i.test(v)) || !vals.every((v) => /^-?\d+(\.\d+)?$/.test(v));
      }
      return false;
    };

    const sourcePrimary = Array.isArray(mappedData) && mappedData.length > 0
      ? mappedData
      : (Array.isArray(mappedRowMatrix) ? mappedRowMatrix : []);
    const sourceRows = sourcePrimary
      .filter((row) => rowLooksMeaningful(row))
      .map((row, idx) => {
        if (Array.isArray(row)) {
          const out = [...row];
          const first = String(out[0] ?? '').trim();
          const m = first.match(/^(\d+)\s+(.+)$/);
          if (m && (!out[1] || String(out[1]).trim() === '')) {
            out[0] = m[1];
            out[1] = m[2];
          }
          if (out[0] == null || String(out[0]).trim() === '') out[0] = idx + 1;
          return out;
        }
        const hdrs = Array.isArray(headersToUse) ? headersToUse : Object.keys(row || {});
        const arr = hdrs.map((h) => row?.[h] ?? '');
        if (arr[0] == null || String(arr[0]).trim() === '') arr[0] = idx + 1;
        return arr;
      });

    // Remove stale sample data above/below model row while keeping header and numeric marker row intact.
    const isNumericMarkerTemplateRow = (r) => {
      if (!markerCols || markerCols.length < 5) return false;
      let numericHits = 0;
      for (let i = 0; i < markerCols.length; i++) {
        const v = normalize(excelCellValueToString(worksheet.getCell(r, markerCols[i])?.value));
        if (/^\d+$/.test(v)) numericHits++;
      }
      return numericHits >= Math.max(5, Math.floor(markerCols.length * 0.6));
    };
    const clearFromRow = Math.max(1, dataStartRow);
    const clearToRow = Math.min(maxScanRows, Math.max(dataStartRow + sourceRows.length + 30, clearFromRow + 30));
    for (let r = clearFromRow; r <= clearToRow; r++) {
      if (isNumericMarkerTemplateRow(r)) continue;
      for (let j = 0; j < orderedCols.length; j++) {
        worksheet.getCell(r, orderedCols[j]).value = '';
      }
    }

    for (let i = 0; i < sourceRows.length; i++) {
      const row = sourceRows[i] || [];
      let maxVisualLines = 1;
      // Clear model row cells first so previous sample values do not leak into next rows.
      for (let j = 0; j < orderedCols.length; j++) {
        worksheet.getCell(dataStartRow + i, orderedCols[j]).value = '';
      }
      for (let j = 0; j < orderedCols.length; j++) {
        const value = row[j];
        if (value == null || value === '') continue;
        const cell = worksheet.getCell(dataStartRow + i, orderedCols[j]);
        if (j === 0 && (typeof value === 'number' || (typeof value === 'string' && /^-?\d+(\.\d+)?$/.test(String(value).trim())))) {
          cell.value = Number(value);
        } else {
          const text = String(value);
          cell.value = text;
          cell.alignment = { ...(cell.alignment || {}), wrapText: true, vertical: 'top' };
          const approxCharsPerLine = 22;
          const visualLines = Math.max(1, Math.ceil(text.length / approxCharsPerLine));
          if (visualLines > maxVisualLines) maxVisualLines = visualLines;
        }
      }
      // Auto-expand row height for long sentence cells so content is fully visible.
      const targetRow = worksheet.getRow(dataStartRow + i);
      const baseHeight = 18;
      targetRow.height = Math.max(baseHeight, Math.min(120, baseHeight * maxVisualLines));
    }

    const out = await workbook.xlsx.writeBuffer();
    const fileName = formFileName ||
      parsedFormHeader?.title?.replace(/[^a-zA-Z0-9]/g, '_') ||
      `Form_${Date.now()}.xlsx`;
    const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    return { blob, fileName };
  };

  const buildForm25WorkbookWithTemplateStyles = async ({
    templateArrayBuffer,
    mappedData,
    mappedRowMatrix,
    headersToUse,
    parsedFormHeader,
    formFileName
  }) => {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(templateArrayBuffer);
    const worksheet = workbook.worksheets[0];
    if (!worksheet) throw new Error('Template worksheet not found.');

    const normalize = (txt) =>
      String(txt || '')
        .replace(/\r?\n/g, ' ')
        .replace(/\s+/g, ' ')
        .trim()
        .toLowerCase();
    const excelCellValueToString = (val) => {
      if (val == null) return '';
      if (typeof val === 'string' || typeof val === 'number' || typeof val === 'boolean') return String(val);
      if (val instanceof Date) return val.toISOString();
      if (typeof val === 'object') {
        if (Array.isArray(val.richText)) return val.richText.map((rt) => rt?.text || '').join('');
        if (val.text != null) return String(val.text);
        if (val.result != null) return String(val.result);
      }
      return '';
    };

    let headerRow = -1;
    let startCol = 1;
    const maxScanRows = Math.max(120, worksheet.rowCount + 10);
    for (let r = 1; r <= maxScanRows; r++) {
      let hits = 0;
      let firstCol = null;
      for (let c = 1; c <= 260; c++) {
        const t = normalize(excelCellValueToString(worksheet.getCell(r, c)?.value));
        if (!t) continue;
        const looksHeader =
          /serial|name\s*of\s*worker|worker\s*id|time\s*at\s*which|overtime|rest\s*interval|dates/.test(t);
        if (looksHeader) {
          hits++;
          if (firstCol == null) firstCol = c;
        }
      }
      if (hits >= 3) {
        headerRow = r;
        startCol = firstCol || 1;
        break;
      }
    }
    if (headerRow < 0) throw new Error('Could not locate Form 25 header row.');

    // Numeric row for dates (1..31) is common in Form 25; keep it intact and write below.
    let markerRow = -1;
    let markerCols = [];
    for (let r = headerRow; r <= Math.min(headerRow + 4, maxScanRows); r++) {
      const cols = [];
      for (let c = startCol; c <= 300; c++) {
        const t = normalize(excelCellValueToString(worksheet.getCell(r, c)?.value));
        if (/^\d+$/.test(t)) cols.push(c);
      }
      if (cols.length > markerCols.length) {
        markerCols = cols;
        markerRow = r;
      }
    }
    const dataStartRow = markerCols.length >= 5 && markerRow > 0 ? (markerRow + 1) : (headerRow + 1);
    const orderedCols = Array.from(
      { length: Math.max(8, Array.isArray(headersToUse) ? headersToUse.length : 8) },
      (_, i) => startCol + i
    );

    const rowLooksMeaningful = (row) => {
      if (Array.isArray(row)) {
        const vals = row.map((v) => String(v ?? '').trim()).filter(Boolean);
        if (vals.length === 0) return false;
        return vals.some((v) => /[a-z]/i.test(v)) || !vals.every((v) => /^-?\d+(\.\d+)?$/.test(v));
      }
      if (row && typeof row === 'object') {
        const vals = Object.values(row).map((v) => String(v ?? '').trim()).filter(Boolean);
        if (vals.length === 0) return false;
        return vals.some((v) => /[a-z]/i.test(v)) || !vals.every((v) => /^-?\d+(\.\d+)?$/.test(v));
      }
      return false;
    };

    const sourcePrimary = Array.isArray(mappedData) && mappedData.length > 0
      ? mappedData
      : (Array.isArray(mappedRowMatrix) ? mappedRowMatrix : []);
    const sourceRows = sourcePrimary
      .filter((row) => rowLooksMeaningful(row))
      .map((row, idx) => {
        if (Array.isArray(row)) {
          const out = [...row];
          const first = String(out[0] ?? '').trim();
          const m = first.match(/^(\d+)\s+(.+)$/);
          if (m && (!out[1] || String(out[1]).trim() === '')) {
            out[0] = m[1];
            out[1] = m[2];
          }
          if (out[0] == null || String(out[0]).trim() === '') out[0] = idx + 1;
          return out;
        }
        const hdrs = Array.isArray(headersToUse) ? headersToUse : Object.keys(row || {});
        const arr = hdrs.map((h) => row?.[h] ?? '');
        if (arr[0] == null || String(arr[0]).trim() === '') arr[0] = idx + 1;
        return arr;
      });

    // Clear stale rows but keep header/marker rows.
    const clearFromRow = Math.max(1, headerRow + 1);
    const clearToRow = Math.min(maxScanRows, Math.max(dataStartRow + sourceRows.length + 30, clearFromRow + 30));
    for (let r = clearFromRow; r <= clearToRow; r++) {
      if (r === headerRow || r === markerRow) continue;
      for (let j = 0; j < orderedCols.length; j++) {
        worksheet.getCell(r, orderedCols[j]).value = '';
      }
    }

    for (let i = 0; i < sourceRows.length; i++) {
      const row = sourceRows[i] || [];
      for (let j = 0; j < orderedCols.length; j++) {
        const value = row[j];
        if (value == null || value === '') continue;
        const cell = worksheet.getCell(dataStartRow + i, orderedCols[j]);
        if (j === 0 && (typeof value === 'number' || (typeof value === 'string' && /^-?\d+(\.\d+)?$/.test(String(value).trim())))) {
          cell.value = Number(value);
        } else {
          cell.value = String(value);
        }
      }
    }

    const out = await workbook.xlsx.writeBuffer();
    const fileName = formFileName ||
      parsedFormHeader?.title?.replace(/[^a-zA-Z0-9]/g, '_') ||
      `Form_${Date.now()}.xlsx`;
    const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    return { blob, fileName };
  };

  const buildFormBWorkbookWithTemplateStyles = async ({
    templateArrayBuffer,
    mappedData,
    mappedRowMatrix,
    headersToUse,
    parsedHeaderRowIndex,
    parsedDataStartIndex,
    parsedFormHeader,
    formFileName
  }) => {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(templateArrayBuffer);
    const worksheet = workbook.worksheets[0];
    if (!worksheet) throw new Error('Template worksheet not found.');

    const normalize = (txt) =>
      String(txt || '')
        .replace(/\r?\n/g, ' ')
        .replace(/\s+/g, ' ')
        .trim()
        .toLowerCase();
    const excelCellValueToString = (val) => {
      if (val == null) return '';
      if (typeof val === 'string' || typeof val === 'number' || typeof val === 'boolean') return String(val);
      if (val instanceof Date) return val.toISOString();
      if (typeof val === 'object') {
        if (Array.isArray(val.richText)) return val.richText.map((rt) => rt?.text || '').join('');
        if (val.text != null) return String(val.text);
        if (val.result != null) return String(val.result);
      }
      return '';
    };

    const maxScanRows = Math.max(120, worksheet.rowCount + 10);
    const headerRow = parsedHeaderRowIndex != null && parsedHeaderRowIndex >= 0 ? (parsedHeaderRowIndex + 1) : -1;
    const dataStartRow = parsedDataStartIndex != null && parsedDataStartIndex >= 0 ? (parsedDataStartIndex + 1) : -1;
    if (headerRow < 1 || dataStartRow < 1) throw new Error('Could not locate Form B table region from template parser.');
    // Form B has multi-band header rows; protect title/month/formula rows by starting lower.
    const effectiveDataStartRow = Math.max(dataStartRow, headerRow + 3);

    let startCol = 1;
    for (let c = 1; c <= 280; c++) {
      const t = normalize(excelCellValueToString(worksheet.getCell(headerRow, c)?.value));
      if (t) {
        startCol = c;
        break;
      }
    }
    const orderedCols = Array.from(
      { length: Math.max(12, Array.isArray(headersToUse) ? headersToUse.length : 12) },
      (_, i) => startCol + i
    );

    const rowLooksMeaningful = (row) => {
      if (Array.isArray(row)) {
        const vals = row.map((v) => String(v ?? '').trim()).filter(Boolean);
        if (vals.length === 0) return false;
        return vals.some((v) => /[a-z]/i.test(v)) || !vals.every((v) => /^-?\d+(\.\d+)?$/.test(v));
      }
      if (row && typeof row === 'object') {
        const vals = Object.values(row).map((v) => String(v ?? '').trim()).filter(Boolean);
        if (vals.length === 0) return false;
        return vals.some((v) => /[a-z]/i.test(v)) || !vals.every((v) => /^-?\d+(\.\d+)?$/.test(v));
      }
      return false;
    };

    const sourcePrimary = Array.isArray(mappedData) && mappedData.length > 0
      ? mappedData
      : (Array.isArray(mappedRowMatrix) ? mappedRowMatrix : []);
    const sourceRows = sourcePrimary
      .filter((row) => rowLooksMeaningful(row))
      .map((row, idx) => {
        if (Array.isArray(row)) return row;
        const hdrs = Array.isArray(headersToUse) ? headersToUse : Object.keys(row || {});
        const arr = hdrs.map((h) => row?.[h] ?? '');
        if (arr[0] == null || String(arr[0]).trim() === '') arr[0] = idx + 1;
        return arr;
      });

    const clearFromRow = Math.max(1, effectiveDataStartRow);
    const clearToRow = Math.min(maxScanRows, Math.max(effectiveDataStartRow + sourceRows.length + 30, clearFromRow + 30));
    for (let r = clearFromRow; r <= clearToRow; r++) {
      for (let j = 0; j < orderedCols.length; j++) {
        const clearCell = worksheet.getCell(r, orderedCols[j]);
        const current = clearCell?.value;
        // Never remove formulas while clearing old data rows.
        if (current && typeof current === 'object' && Object.prototype.hasOwnProperty.call(current, 'formula')) continue;
        clearCell.value = '';
      }
    }

    for (let i = 0; i < sourceRows.length; i++) {
      const row = sourceRows[i] || [];
      for (let j = 0; j < orderedCols.length; j++) {
        const value = row[j];
        if (value == null || value === '') continue;
        const cell = worksheet.getCell(effectiveDataStartRow + i, orderedCols[j]);
        if (j === 0 && (typeof value === 'number' || (typeof value === 'string' && /^-?\d+(\.\d+)?$/.test(String(value).trim())))) {
          cell.value = Number(value);
        } else {
          cell.value = String(value);
        }
      }
    }

    const out = await workbook.xlsx.writeBuffer();
    const fileName = formFileName ||
      parsedFormHeader?.title?.replace(/[^a-zA-Z0-9]/g, '_') ||
      `Form_${Date.now()}.xlsx`;
    const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    return { blob, fileName };
  };

  const buildFormWWorkbookWithTemplateStyles = async ({
    templateArrayBuffer,
    mappedData,
    mappedRowMatrix,
    headersToUse,
    parsedHeaderRowIndex,
    parsedDataStartIndex,
    parsedFormHeader,
    formFileName
  }) => {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(templateArrayBuffer);
    const worksheet = workbook.worksheets[0];
    if (!worksheet) throw new Error('Template worksheet not found.');

    const normalize = (txt) =>
      String(txt || '')
        .replace(/\r?\n/g, ' ')
        .replace(/\s+/g, ' ')
        .trim()
        .toLowerCase();
    const excelCellValueToString = (val) => {
      if (val == null) return '';
      if (typeof val === 'string' || typeof val === 'number' || typeof val === 'boolean') return String(val);
      if (val instanceof Date) return val.toISOString();
      if (typeof val === 'object') {
        if (Array.isArray(val.richText)) return val.richText.map((rt) => rt?.text || '').join('');
        if (val.text != null) return String(val.text);
        if (val.result != null) return String(val.result);
      }
      return '';
    };

    const maxScanRows = Math.max(200, worksheet.rowCount + 20);
    const headerRow = parsedHeaderRowIndex != null && parsedHeaderRowIndex >= 0 ? (parsedHeaderRowIndex + 1) : -1;
    const dataStartRow = parsedDataStartIndex != null && parsedDataStartIndex >= 0 ? (parsedDataStartIndex + 1) : -1;
    if (headerRow < 1 || dataStartRow < 1) throw new Error('Could not locate Form W table region from template parser.');

    let startCol = 1;
    for (let c = 1; c <= 280; c++) {
      const t = normalize(excelCellValueToString(worksheet.getCell(headerRow, c)?.value));
      if (t) {
        startCol = c;
        break;
      }
    }
    const orderedCols = Array.from(
      { length: Math.max(18, Array.isArray(headersToUse) ? headersToUse.length : 18) },
      (_, i) => startCol + i
    );
    // Form W usually has multi-tier/merged headers; keep 2 rows below parser header before writing.
    const effectiveDataStartRow = Math.max(dataStartRow, headerRow + 3);

    const rowLooksMeaningful = (row) => {
      if (Array.isArray(row)) {
        const vals = row.map((v) => String(v ?? '').trim()).filter(Boolean);
        if (vals.length === 0) return false;
        return vals.some((v) => /[a-z]/i.test(v)) || !vals.every((v) => /^-?\d+(\.\d+)?$/.test(v));
      }
      if (row && typeof row === 'object') {
        const vals = Object.values(row).map((v) => String(v ?? '').trim()).filter(Boolean);
        if (vals.length === 0) return false;
        return vals.some((v) => /[a-z]/i.test(v)) || !vals.every((v) => /^-?\d+(\.\d+)?$/.test(v));
      }
      return false;
    };

    const sourcePrimary = Array.isArray(mappedData) && mappedData.length > 0
      ? mappedData
      : (Array.isArray(mappedRowMatrix) ? mappedRowMatrix : []);
    const sourceRows = sourcePrimary
      .filter((row) => rowLooksMeaningful(row))
      .map((row, idx) => {
        if (Array.isArray(row)) return row;
        const hdrs = Array.isArray(headersToUse) ? headersToUse : Object.keys(row || {});
        const arr = hdrs.map((h) => row?.[h] ?? '');
        if (arr[0] == null || String(arr[0]).trim() === '') arr[0] = idx + 1;
        return arr;
      });

    const clearFromRow = Math.max(1, effectiveDataStartRow);
    const clearToRow = Math.min(maxScanRows, Math.max(effectiveDataStartRow + sourceRows.length + 40, clearFromRow + 40));
    for (let r = clearFromRow; r <= clearToRow; r++) {
      for (let j = 0; j < orderedCols.length; j++) {
        const clearCell = worksheet.getCell(r, orderedCols[j]);
        const current = clearCell?.value;
        if (current && typeof current === 'object' && Object.prototype.hasOwnProperty.call(current, 'formula')) continue;
        clearCell.value = '';
      }
    }

    for (let i = 0; i < sourceRows.length; i++) {
      const row = sourceRows[i] || [];
      for (let j = 0; j < orderedCols.length; j++) {
        const value = row[j];
        if (value == null || value === '') continue;
        const cell = worksheet.getCell(effectiveDataStartRow + i, orderedCols[j]);
        if (j === 0 && (typeof value === 'number' || (typeof value === 'string' && /^-?\d+(\.\d+)?$/.test(String(value).trim())))) {
          cell.value = Number(value);
        } else {
          cell.value = String(value);
        }
      }
    }

    const out = await workbook.xlsx.writeBuffer();
    const fileName = formFileName ||
      parsedFormHeader?.title?.replace(/[^a-zA-Z0-9]/g, '_') ||
      `Form_${Date.now()}.xlsx`;
    const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    return { blob, fileName };
  };

  const buildFormXWorkbookWithTemplateStyles = async ({
    templateArrayBuffer,
    mappedData,
    mappedRowMatrix,
    headersToUse,
    parsedHeaderRowIndex,
    parsedDataStartIndex,
    parsedFormHeader,
    formFileName
  }) => {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(templateArrayBuffer);
    const worksheet = workbook.worksheets[0];
    if (!worksheet) throw new Error('Template worksheet not found.');

    const normalize = (txt) =>
      String(txt || '')
        .replace(/\r?\n/g, ' ')
        .replace(/\s+/g, ' ')
        .trim()
        .toLowerCase();
    const excelCellValueToString = (val) => {
      if (val == null) return '';
      if (typeof val === 'string' || typeof val === 'number' || typeof val === 'boolean') return String(val);
      if (val instanceof Date) return val.toISOString();
      if (typeof val === 'object') {
        if (Array.isArray(val.richText)) return val.richText.map((rt) => rt?.text || '').join('');
        if (val.text != null) return String(val.text);
        if (val.result != null) return String(val.result);
      }
      return '';
    };

    const maxScanRows = Math.max(200, worksheet.rowCount + 20);
    const headerRow = parsedHeaderRowIndex != null && parsedHeaderRowIndex >= 0 ? (parsedHeaderRowIndex + 1) : -1;
    const dataStartRow = parsedDataStartIndex != null && parsedDataStartIndex >= 0 ? (parsedDataStartIndex + 1) : -1;
    if (headerRow < 1 || dataStartRow < 1) throw new Error('Could not locate Form X table region from template parser.');

    let startCol = 1;
    for (let c = 1; c <= 280; c++) {
      const t = normalize(excelCellValueToString(worksheet.getCell(headerRow, c)?.value));
      if (t) {
        startCol = c;
        break;
      }
    }
    const orderedCols = Array.from(
      { length: Math.max(18, Array.isArray(headersToUse) ? headersToUse.length : 18) },
      (_, i) => startCol + i
    );
    const effectiveDataStartRow = Math.max(dataStartRow, headerRow + 3);

    const rowLooksMeaningful = (row) => {
      if (Array.isArray(row)) {
        const vals = row.map((v) => String(v ?? '').trim()).filter(Boolean);
        if (vals.length === 0) return false;
        return vals.some((v) => /[a-z]/i.test(v)) || !vals.every((v) => /^-?\d+(\.\d+)?$/.test(v));
      }
      if (row && typeof row === 'object') {
        const vals = Object.values(row).map((v) => String(v ?? '').trim()).filter(Boolean);
        if (vals.length === 0) return false;
        return vals.some((v) => /[a-z]/i.test(v)) || !vals.every((v) => /^-?\d+(\.\d+)?$/.test(v));
      }
      return false;
    };

    const sourcePrimary = Array.isArray(mappedData) && mappedData.length > 0
      ? mappedData
      : (Array.isArray(mappedRowMatrix) ? mappedRowMatrix : []);
    const sourceRows = sourcePrimary
      .filter((row) => rowLooksMeaningful(row))
      .map((row, idx) => {
        if (Array.isArray(row)) return row;
        const hdrs = Array.isArray(headersToUse) ? headersToUse : Object.keys(row || {});
        const arr = hdrs.map((h) => row?.[h] ?? '');
        if (arr[0] == null || String(arr[0]).trim() === '') arr[0] = idx + 1;
        return arr;
      });

    const clearFromRow = Math.max(1, effectiveDataStartRow);
    const clearToRow = Math.min(maxScanRows, Math.max(effectiveDataStartRow + sourceRows.length + 40, clearFromRow + 40));
    for (let r = clearFromRow; r <= clearToRow; r++) {
      for (let j = 0; j < orderedCols.length; j++) {
        const clearCell = worksheet.getCell(r, orderedCols[j]);
        const current = clearCell?.value;
        if (current && typeof current === 'object' && Object.prototype.hasOwnProperty.call(current, 'formula')) continue;
        clearCell.value = '';
      }
    }

    for (let i = 0; i < sourceRows.length; i++) {
      const row = sourceRows[i] || [];
      for (let j = 0; j < orderedCols.length; j++) {
        const value = row[j];
        if (value == null || value === '') continue;
        const cell = worksheet.getCell(effectiveDataStartRow + i, orderedCols[j]);
        if (j === 0 && (typeof value === 'number' || (typeof value === 'string' && /^-?\d+(\.\d+)?$/.test(String(value).trim())))) {
          cell.value = Number(value);
        } else {
          cell.value = String(value);
        }
      }
    }

    const out = await workbook.xlsx.writeBuffer();
    const fileName = formFileName ||
      parsedFormHeader?.title?.replace(/[^a-zA-Z0-9]/g, '_') ||
      `Form_${Date.now()}.xlsx`;
    const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    return { blob, fileName };
  };

  const buildFormCWorkbookWithTemplateStyles = async ({
    templateArrayBuffer,
    mappedData,
    mappedRowMatrix,
    headersToUse,
    parsedHeaderRowIndex,
    parsedDataStartIndex,
    parsedFormHeader,
    formFileName
  }) => {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(templateArrayBuffer);
    const worksheet = workbook.worksheets[0];
    if (!worksheet) throw new Error('Template worksheet not found.');

    const normalize = (txt) =>
      String(txt || '')
        .replace(/\r?\n/g, ' ')
        .replace(/\s+/g, ' ')
        .trim()
        .toLowerCase();
    const excelCellValueToString = (val) => {
      if (val == null) return '';
      if (typeof val === 'string' || typeof val === 'number' || typeof val === 'boolean') return String(val);
      if (val instanceof Date) return val.toISOString();
      if (typeof val === 'object') {
        if (Array.isArray(val.richText)) return val.richText.map((rt) => rt?.text || '').join('');
        if (val.text != null) return String(val.text);
        if (val.result != null) return String(val.result);
      }
      return '';
    };

    const headerRow = parsedHeaderRowIndex != null && parsedHeaderRowIndex >= 0 ? (parsedHeaderRowIndex + 1) : -1;
    const dataStartRow = parsedDataStartIndex != null && parsedDataStartIndex >= 0 ? (parsedDataStartIndex + 1) : -1;
    if (headerRow < 1 || dataStartRow < 1) throw new Error('Could not locate Form C table region from template parser.');

    let startCol = 1;
    for (let c = 1; c <= 280; c++) {
      const t = normalize(excelCellValueToString(worksheet.getCell(headerRow, c)?.value));
      if (t) {
        startCol = c;
        break;
      }
    }
    const orderedCols = Array.from(
      { length: Math.max(12, Array.isArray(headersToUse) ? headersToUse.length : 12) },
      (_, i) => startCol + i
    );

    const rowLooksMeaningful = (row) => {
      if (Array.isArray(row)) return row.some((v) => String(v ?? '').trim() !== '');
      if (row && typeof row === 'object') return Object.values(row).some((v) => String(v ?? '').trim() !== '');
      return false;
    };
    const sourcePrimary = Array.isArray(mappedData) && mappedData.length > 0
      ? mappedData
      : (Array.isArray(mappedRowMatrix) ? mappedRowMatrix : []);
    const sourceRows = sourcePrimary
      .filter((row) => rowLooksMeaningful(row))
      .map((row, idx) => {
        if (Array.isArray(row)) return row;
        const hdrs = Array.isArray(headersToUse) ? headersToUse : Object.keys(row || {});
        const arr = hdrs.map((h) => row?.[h] ?? '');
        if (arr[0] == null || String(arr[0]).trim() === '') arr[0] = idx + 1;
        return arr;
      });

    const maxScanRows = Math.max(120, worksheet.rowCount + 10);
    const clearToRow = Math.min(maxScanRows, Math.max(dataStartRow + sourceRows.length + 30, dataStartRow + 30));
    for (let r = dataStartRow; r <= clearToRow; r++) {
      for (let j = 0; j < orderedCols.length; j++) {
        worksheet.getCell(r, orderedCols[j]).value = '';
      }
    }

    for (let i = 0; i < sourceRows.length; i++) {
      const row = sourceRows[i] || [];
      for (let j = 0; j < orderedCols.length; j++) {
        const value = row[j];
        if (value == null || value === '') continue;
        const cell = worksheet.getCell(dataStartRow + i, orderedCols[j]);
        if (j === 0 && (typeof value === 'number' || (typeof value === 'string' && /^-?\d+(\.\d+)?$/.test(String(value).trim())))) {
          cell.value = Number(value);
        } else {
          cell.value = String(value);
        }
      }
    }

    const out = await workbook.xlsx.writeBuffer();
    const fileName = formFileName ||
      parsedFormHeader?.title?.replace(/[^a-zA-Z0-9]/g, '_') ||
      `Form_${Date.now()}.xlsx`;
    const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    return { blob, fileName };
  };

  const buildFormDWorkbookWithTemplateStyles = async ({
    templateArrayBuffer,
    mappedData,
    mappedRowMatrix,
    headersToUse,
    parsedHeaderRowIndex,
    parsedDataStartIndex,
    parsedFormHeader,
    formFileName
  }) => {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(templateArrayBuffer);
    const worksheet = workbook.worksheets[0];
    if (!worksheet) throw new Error('Template worksheet not found.');

    let headerRow = -1;
    let dataStartRow = -1;

    const normalize = (txt) =>
      String(txt || '')
        .replace(/\r?\n/g, ' ')
        .replace(/\s+/g, ' ')
        .trim()
        .toLowerCase();
    const excelCellValueToString = (val) => {
      if (val == null) return '';
      if (typeof val === 'string' || typeof val === 'number' || typeof val === 'boolean') return String(val);
      if (val instanceof Date) return val.toISOString();
      if (typeof val === 'object') {
        if (Array.isArray(val.richText)) return val.richText.map((rt) => rt?.text || '').join('');
        if (val.text != null) return String(val.text);
        if (val.result != null) return String(val.result);
      }
      return '';
    };

    const maxScanRows = Math.max(180, worksheet.rowCount + 20);
    const maxScanCols = 320;
    // Form D hard anchor: table header row containing "Category of workers".
    for (let r = 1; r <= maxScanRows; r++) {
      for (let c = 1; c <= maxScanCols; c++) {
        const t = normalize(excelCellValueToString(worksheet.getCell(r, c)?.value));
        if (!t) continue;
        if (t.includes('category of workers')) {
          headerRow = r;
          break;
        }
      }
      if (headerRow > 0) break;
    }
    if (headerRow < 1 && parsedHeaderRowIndex != null && parsedHeaderRowIndex >= 0) {
      headerRow = parsedHeaderRowIndex + 1;
    }
    if (headerRow < 1) throw new Error('Could not locate Form D header row.');

    let startCol = 1;
    for (let c = 1; c <= maxScanCols; c++) {
      const t = normalize(excelCellValueToString(worksheet.getCell(headerRow, c)?.value));
      if (t) {
        startCol = c;
        break;
      }
    }

    // Data starts below numeric marker row (1..N) when present; otherwise below header.
    let markerRow = -1;
    for (let r = headerRow; r <= Math.min(headerRow + 3, maxScanRows); r++) {
      let numericHits = 0;
      for (let c = startCol; c <= Math.min(startCol + 20, maxScanCols); c++) {
        const t = normalize(excelCellValueToString(worksheet.getCell(r, c)?.value));
        if (/^\d+$/.test(t)) numericHits++;
      }
      if (numericHits >= 3) {
        markerRow = r;
        break;
      }
    }
    dataStartRow = markerRow > 0 ? (markerRow + 1) : (headerRow + 1);
    if (parsedDataStartIndex != null && parsedDataStartIndex >= 0) {
      // Prefer deeper start row to avoid touching top template sections.
      dataStartRow = Math.max(dataStartRow, parsedDataStartIndex + 1);
    }
    const orderedCols = Array.from(
      { length: Math.max(14, Array.isArray(headersToUse) ? headersToUse.length : 14) },
      (_, i) => startCol + i
    );

    const rowLooksMeaningful = (row) => {
      if (Array.isArray(row)) return row.some((v) => String(v ?? '').trim() !== '');
      if (row && typeof row === 'object') return Object.values(row).some((v) => String(v ?? '').trim() !== '');
      return false;
    };
    const sourcePrimary = Array.isArray(mappedData) && mappedData.length > 0
      ? mappedData
      : (Array.isArray(mappedRowMatrix) ? mappedRowMatrix : []);
    const sourceRows = sourcePrimary
      .filter((row) => rowLooksMeaningful(row))
      .map((row, idx) => {
        if (Array.isArray(row)) return row;
        const hdrs = Array.isArray(headersToUse) ? headersToUse : Object.keys(row || {});
        const arr = hdrs.map((h) => row?.[h] ?? '');
        if (arr[0] == null || String(arr[0]).trim() === '') arr[0] = idx + 1;
        return arr;
      });

    const clearToRow = Math.min(maxScanRows, Math.max(dataStartRow + sourceRows.length + 40, dataStartRow + 40));
    for (let r = dataStartRow; r <= clearToRow; r++) {
      for (let j = 0; j < orderedCols.length; j++) {
        worksheet.getCell(r, orderedCols[j]).value = '';
      }
    }

    for (let i = 0; i < sourceRows.length; i++) {
      const row = sourceRows[i] || [];
      for (let j = 0; j < orderedCols.length; j++) {
        const value = row[j];
        if (value == null || value === '') continue;
        const cell = worksheet.getCell(dataStartRow + i, orderedCols[j]);
        if (j === 0 && (typeof value === 'number' || (typeof value === 'string' && /^-?\d+(\.\d+)?$/.test(String(value).trim())))) {
          cell.value = Number(value);
        } else {
          cell.value = String(value);
        }
      }
    }

    const out = await workbook.xlsx.writeBuffer();
    const fileName = formFileName ||
      parsedFormHeader?.title?.replace(/[^a-zA-Z0-9]/g, '_') ||
      `Form_${Date.now()}.xlsx`;
    const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    return { blob, fileName };
  };

  const buildFormIWorkbookWithTemplateStyles = async ({
    templateArrayBuffer,
    mappedData,
    mappedRowMatrix,
    headersToUse,
    parsedHeaderRowIndex,
    parsedDataStartIndex,
    parsedFormHeader,
    formFileName
  }) => {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(templateArrayBuffer);
    const worksheet = workbook.worksheets[0];
    if (!worksheet) throw new Error('Template worksheet not found.');

    const normalize = (txt) =>
      String(txt || '')
        .replace(/\r?\n/g, ' ')
        .replace(/\s+/g, ' ')
        .trim()
        .toLowerCase();
    const excelCellValueToString = (val) => {
      if (val == null) return '';
      if (typeof val === 'string' || typeof val === 'number' || typeof val === 'boolean') return String(val);
      if (val instanceof Date) return val.toISOString();
      if (typeof val === 'object') {
        if (Array.isArray(val.richText)) return val.richText.map((rt) => rt?.text || '').join('');
        if (val.text != null) return String(val.text);
        if (val.result != null) return String(val.result);
      }
      return '';
    };

    const maxScanRows = Math.max(220, worksheet.rowCount + 30);
    const maxScanCols = 340;
    if (parsedHeaderRowIndex == null || parsedHeaderRowIndex < 0) {
      throw new Error('Could not locate Form I header row.');
    }
    const headerRow = parsedHeaderRowIndex + 1;
    let startCol = 1;
    for (let c = 1; c <= maxScanCols; c++) {
      const t = normalize(excelCellValueToString(worksheet.getCell(headerRow, c)?.value));
      if (t) {
        startCol = c;
        break;
      }
    }

    let dataStartRow =
      parsedDataStartIndex != null && parsedDataStartIndex >= 0
        ? (parsedDataStartIndex + 1)
        : (headerRow + 1);

    const rowLooksMeaningful = (row) => {
      if (Array.isArray(row)) return row.some((v) => String(v ?? '').trim() !== '');
      if (row && typeof row === 'object') return Object.values(row).some((v) => String(v ?? '').trim() !== '');
      return false;
    };
    const sourcePrimary = Array.isArray(mappedData) && mappedData.length > 0
      ? mappedData
      : (Array.isArray(mappedRowMatrix) ? mappedRowMatrix : []);
    const sourceRows = sourcePrimary
      .filter((row) => rowLooksMeaningful(row))
      .map((row, idx) => {
        if (Array.isArray(row)) return row;
        const hdrs = Array.isArray(headersToUse) ? headersToUse : Object.keys(row || {});
        const arr = hdrs.map((h) => row?.[h] ?? '');
        if (arr[0] == null || String(arr[0]).trim() === '') arr[0] = idx + 1;
        return arr;
      });

    // Write only row data cells; do not clear wide regions to avoid touching merged title layout.
    const maxColsToWrite = Math.max(12, Array.isArray(headersToUse) ? headersToUse.length : 12);
    for (let i = 0; i < sourceRows.length; i++) {
      const row = sourceRows[i] || [];
      for (let j = 0; j < maxColsToWrite; j++) {
        const value = row[j];
        const cell = worksheet.getCell(dataStartRow + i, startCol + j);
        if (value == null || value === '') {
          cell.value = '';
          continue;
        }
        if (j === 0 && (typeof value === 'number' || (typeof value === 'string' && /^-?\d+(\.\d+)?$/.test(String(value).trim())))) {
          cell.value = Number(value);
        } else {
          cell.value = String(value);
        }
      }
    }

    const out = await workbook.xlsx.writeBuffer();
    const fileName = formFileName ||
      parsedFormHeader?.title?.replace(/[^a-zA-Z0-9]/g, '_') ||
      `Form_${Date.now()}.xlsx`;
    const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    return { blob, fileName };
  };

  const buildFormVWorkbookWithTemplateStyles = async ({
    templateArrayBuffer,
    mappedData,
    mappedRowMatrix,
    headersToUse,
    parsedHeaderRowIndex,
    parsedDataStartIndex,
    parsedFormHeader,
    formFileName
  }) => {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(templateArrayBuffer);
    const worksheet = workbook.worksheets[0];
    if (!worksheet) throw new Error('Template worksheet not found.');

    const normalize = (txt) =>
      String(txt || '')
        .replace(/\r?\n/g, ' ')
        .replace(/\s+/g, ' ')
        .trim()
        .toLowerCase();
    const excelCellValueToString = (val) => {
      if (val == null) return '';
      if (typeof val === 'string' || typeof val === 'number' || typeof val === 'boolean') return String(val);
      if (val instanceof Date) return val.toISOString();
      if (typeof val === 'object') {
        if (Array.isArray(val.richText)) return val.richText.map((rt) => rt?.text || '').join('');
        if (val.text != null) return String(val.text);
        if (val.result != null) return String(val.result);
      }
      return '';
    };

    const maxScanCols = 360;
    const headerRow = parsedHeaderRowIndex != null && parsedHeaderRowIndex >= 0 ? (parsedHeaderRowIndex + 1) : -1;
    const dataStartRow = parsedDataStartIndex != null && parsedDataStartIndex >= 0 ? (parsedDataStartIndex + 1) : -1;
    if (headerRow < 1 || dataStartRow < 1) {
      throw new Error('Could not locate Form V table region from template parser.');
    }

    // Find the day-number row (1..31). Usually directly below header row in Form V.
    let dayRow = -1;
    let dayColsByNumber = new Map();
    for (let r = headerRow; r <= headerRow + 3; r++) {
      const cols = new Map();
      for (let c = 1; c <= maxScanCols; c++) {
        const t = normalize(excelCellValueToString(worksheet.getCell(r, c)?.value));
        if (/^\d{1,2}$/.test(t)) {
          const n = parseInt(t, 10);
          if (n >= 1 && n <= 31) cols.set(n, c);
        }
      }
      if (cols.size > dayColsByNumber.size) {
        dayColsByNumber = cols;
        dayRow = r;
      }
    }

    // Default left-start column is first non-empty header cell.
    let startCol = 1;
    for (let c = 1; c <= maxScanCols; c++) {
      const t = normalize(excelCellValueToString(worksheet.getCell(headerRow, c)?.value));
      if (t) {
        startCol = c;
        break;
      }
    }

    const rowLooksMeaningful = (row) => {
      if (Array.isArray(row)) return row.some((v) => String(v ?? '').trim() !== '');
      if (row && typeof row === 'object') return Object.values(row).some((v) => String(v ?? '').trim() !== '');
      return false;
    };
    const sourcePrimary = Array.isArray(mappedData) && mappedData.length > 0
      ? mappedData
      : (Array.isArray(mappedRowMatrix) ? mappedRowMatrix : []);
    const sourceRows = sourcePrimary
      .filter((row) => rowLooksMeaningful(row))
      .map((row, idx) => {
        if (Array.isArray(row)) return row;
        const hdrs = Array.isArray(headersToUse) ? headersToUse : Object.keys(row || {});
        const arr = hdrs.map((h) => row?.[h] ?? '');
        if (arr[0] == null || String(arr[0]).trim() === '') arr[0] = idx + 1;
        return arr;
      });

    // Clear only the data grid rows we will write (avoid touching header + merged layout).
    const maxColsToWrite = Math.max(12, Array.isArray(headersToUse) ? headersToUse.length : 12);
    for (let i = 0; i < Math.min(sourceRows.length + 30, 200); i++) {
      for (let j = 0; j < maxColsToWrite; j++) {
        worksheet.getCell(dataStartRow + i, startCol + j).value = '';
      }
    }

    const headerIndex = new Map();
    if (Array.isArray(headersToUse)) headersToUse.forEach((h, i) => headerIndex.set(h, i));
    const dayHeaderIndices = [];
    if (Array.isArray(headersToUse)) {
      headersToUse.forEach((h, i) => {
        const m = String(h || '').match(/_(\d{1,2})$/);
        if (!m) return;
        const n = parseInt(m[1], 10);
        if (n >= 1 && n <= 31) dayHeaderIndices.push({ idx: i, day: n });
      });
    }

    for (let i = 0; i < sourceRows.length; i++) {
      const row = sourceRows[i] || [];

      // Write non-day columns sequentially into the left area.
      for (let j = 0; j < maxColsToWrite; j++) {
        // Skip day headers here; they'll be written by real day columns if we found them.
        const h = Array.isArray(headersToUse) ? headersToUse[j] : null;
        if (h && /_(\d{1,2})$/.test(String(h))) continue;
        const value = row[j];
        if (value == null || value === '') continue;
        const cell = worksheet.getCell(dataStartRow + i, startCol + j);
        if (j === 0 && (typeof value === 'number' || (typeof value === 'string' && /^-?\d+(\.\d+)?$/.test(String(value).trim())))) {
          cell.value = Number(value);
        } else {
          cell.value = String(value);
        }
      }

      // Write day-grid values into the exact day-number columns.
      if (dayRow > 0 && dayColsByNumber.size >= 10 && dayHeaderIndices.length > 0) {
        for (const { idx, day } of dayHeaderIndices) {
          const targetCol = dayColsByNumber.get(day);
          if (!targetCol) continue;
          const value = row[idx];
          if (value == null || value === '') continue;
          const cell = worksheet.getCell(dataStartRow + i, targetCol);
          cell.value = String(value);
        }
      }
    }

    const out = await workbook.xlsx.writeBuffer();
    const fileName = formFileName ||
      parsedFormHeader?.title?.replace(/[^a-zA-Z0-9]/g, '_') ||
      `Form_${Date.now()}.xlsx`;
    const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    return { blob, fileName };
  };

  const buildFormVIWorkbookWithTemplateStyles = async ({
    templateArrayBuffer,
    mappedData,
    mappedRowMatrix,
    headersToUse,
    parsedHeaderRowIndex,
    parsedDataStartIndex,
    parsedFormHeader,
    formFileName
  }) => {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(templateArrayBuffer);
    const worksheet = workbook.worksheets[0];
    if (!worksheet) throw new Error('Template worksheet not found.');

    const headerRow = parsedHeaderRowIndex != null && parsedHeaderRowIndex >= 0 ? (parsedHeaderRowIndex + 1) : -1;
    const dataStartRow = parsedDataStartIndex != null && parsedDataStartIndex >= 0 ? (parsedDataStartIndex + 1) : -1;
    if (headerRow < 1 || dataStartRow < 1) {
      throw new Error('Could not locate Form VI table region from template parser.');
    }

    let startCol = 1;
    const maxScanCols = 360;
    for (let c = 1; c <= maxScanCols; c++) {
      const v = worksheet.getCell(headerRow, c)?.value;
      if (v != null && String(v).trim() !== '') {
        startCol = c;
        break;
      }
    }

    const rowLooksMeaningful = (row) => {
      if (Array.isArray(row)) return row.some((v) => String(v ?? '').trim() !== '');
      if (row && typeof row === 'object') return Object.values(row).some((v) => String(v ?? '').trim() !== '');
      return false;
    };
    const sourcePrimary = Array.isArray(mappedData) && mappedData.length > 0
      ? mappedData
      : (Array.isArray(mappedRowMatrix) ? mappedRowMatrix : []);
    const sourceRows = sourcePrimary
      .filter((row) => rowLooksMeaningful(row))
      .map((row, idx) => {
        if (Array.isArray(row)) return row;
        const hdrs = Array.isArray(headersToUse) ? headersToUse : Object.keys(row || {});
        const arr = hdrs.map((h) => row?.[h] ?? '');
        if (arr[0] == null || String(arr[0]).trim() === '') arr[0] = idx + 1;
        return arr;
      });

    const maxColsToWrite = Math.max(12, Array.isArray(headersToUse) ? headersToUse.length : 12);
    const clearToRow = Math.min(Math.max(200, worksheet.rowCount + 20), dataStartRow + sourceRows.length + 40);
    for (let r = dataStartRow; r <= clearToRow; r++) {
      for (let j = 0; j < maxColsToWrite; j++) {
        worksheet.getCell(r, startCol + j).value = '';
      }
    }

    for (let i = 0; i < sourceRows.length; i++) {
      const row = sourceRows[i] || [];
      for (let j = 0; j < maxColsToWrite; j++) {
        const value = row[j];
        if (value == null || value === '') continue;
        const cell = worksheet.getCell(dataStartRow + i, startCol + j);
        if (j === 0 && (typeof value === 'number' || (typeof value === 'string' && /^-?\d+(\.\d+)?$/.test(String(value).trim())))) {
          cell.value = Number(value);
        } else {
          cell.value = String(value);
        }
      }
    }

    const out = await workbook.xlsx.writeBuffer();
    const fileName = formFileName ||
      parsedFormHeader?.title?.replace(/[^a-zA-Z0-9]/g, '_') ||
      `Form_${Date.now()}.xlsx`;
    const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    return { blob, fileName };
  };

  // View Draft File: generate formmaster template + autofill (Zoho People) data and download (no modal, no save)
  const handleViewDraftFileGenerate = async (item, resolvedFormFileItem) => {
    const rowFormNameKey = normalizeTemplateFormNameKey(
      item?.formName || item?.FormName || resolvedFormFileItem?.formName || resolvedFormFileItem?.FormName || ''
    );
    const formmasterTemplateForRow = rowFormNameKey ? formmasterFormFileByFormName.get(rowFormNameKey) : null;
    const templateMeta = formmasterTemplateForRow && formmasterTemplateForRow.formFile
      ? formmasterTemplateForRow
      : resolvedFormFileItem;
    const hasFormFile = templateMeta?.formFile && templateMeta.formFile !== 'null' && String(templateMeta.formFile).trim() !== '';
    if (!hasFormFile) {
      setError('No form template available for this row. Use the draft link to open the saved file.');
      return;
    }
    try {
      setFormFileLoading(true);
      setError('');
      setSuccess('Generating draft...');
      const fn = (templateMeta.formFileName || resolvedFormFileItem.formFileName || 'template.xlsx').replace(/[^a-zA-Z0-9._-]/g, '_');
      const fileUrl = `/server/formmaster_function/templates/download/${templateMeta.formFile}?fileName=${encodeURIComponent(fn)}`;
      const res = await fetch(fileUrl);
      if (!res.ok) {
        setError('Could not load formmaster template.');
        return;
      }
      const arrayBuffer = await res.arrayBuffer();
      const templateWb = XLSX.read(arrayBuffer, { type: 'array' });
      const parsed = parseExcelForm(templateWb);
      let headersToUse = parsed.headers?.length ? [...parsed.headers] : [];
      if (!headersToUse.length) {
        setError('Could not read template headers.');
        return;
      }
      let mappedData = null;
      let usedLiveModalGrid = false;
      let usedSnapshot = false;
      let usedSavedDraftFile = false;
      let savedDraftRowMatrix = null;
      const downloadReqTs = Date.now();
      const isFormADownload = isFormAMusterRollContext(parsed.formHeader) ||
        /\bform\s*[-"']?\s*a\b/i.test(String(item?.formName || item?.FormName || ''));
      const isFormUDownload =
        /\bform\s*[-"']?\s*u\b/i.test(String(item?.formName || item?.FormName || '')) ||
        /\bform\s*[-"']?\s*u\b/i.test(String(parsed?.formHeader?.title || ''));
      const isForm12Download =
        /\bform\s*[-"']?\s*12\b/i.test(String(item?.formName || item?.FormName || '')) ||
        /\bform\s*[-"']?\s*12\b/i.test(String(parsed?.formHeader?.title || ''));
      const isForm10Download =
        /\bform\s*[-"']?\s*10\b/i.test(String(item?.formName || item?.FormName || '')) ||
        /\bform\s*[-"']?\s*10\b/i.test(String(parsed?.formHeader?.title || ''));
      const isForm25Download =
        /\bform\s*[-"']?\s*25\b/i.test(String(item?.formName || item?.FormName || '')) ||
        /\bform\s*[-"']?\s*25\b/i.test(String(parsed?.formHeader?.title || ''));
      const isFormBDownload =
        /\bform\s*[-"']?\s*b\b/i.test(String(item?.formName || item?.FormName || '')) ||
        /\bform\s*[-"']?\s*b\b/i.test(String(parsed?.formHeader?.title || ''));
      const isFormWDownload =
        /\bform\s*[-"']?\s*w\b/i.test(String(item?.formName || item?.FormName || '')) ||
        /\bform\s*[-"']?\s*w\b/i.test(String(parsed?.formHeader?.title || '')) ||
        /register\s+of\s+wages/i.test(String(parsed?.formHeader?.title || ''));
      const isFormXDownload =
        /\bform\s*[-"']?\s*x\b/i.test(String(item?.formName || item?.FormName || '')) ||
        /\bform\s*[-"']?\s*x\b/i.test(String(parsed?.formHeader?.title || '')) ||
        /leave\s+with\s+wages/i.test(String(parsed?.formHeader?.title || ''));
      const isFormCDownload =
        /\bform\s*[-"']?\s*c\b/i.test(String(item?.formName || item?.FormName || '')) ||
        /\bform\s*[-"']?\s*c\b/i.test(String(parsed?.formHeader?.title || ''));
      const isFormDDownload =
        /\bform\s*[-"']?\s*d\b/i.test(String(item?.formName || item?.FormName || '')) ||
        /\bform\s*[-"']?\s*d\b/i.test(String(parsed?.formHeader?.title || ''));
      const isFormIDownload =
        /\bform\s*[-"']?\s*i\b/i.test(String(item?.formName || item?.FormName || '')) ||
        /\bform\s*[-"']?\s*i\b/i.test(String(parsed?.formHeader?.title || ''));
      const isFormVDownload =
        /\bform\s*[-"']?\s*v\b/i.test(String(item?.formName || item?.FormName || '')) ||
        /\bform\s*[-"']?\s*v\b/i.test(String(parsed?.formHeader?.title || '')) ||
        /\bform\s*[-"']?\s*5\b/i.test(String(item?.formName || item?.FormName || '')) ||
        /\bform\s*[-"']?\s*5\b/i.test(String(parsed?.formHeader?.title || ''));
      const isFormVIDownload =
        /\bform\s*[-"']?\s*vi\b/i.test(String(item?.formName || item?.FormName || '')) ||
        /\bform\s*[-"']?\s*6\b/i.test(String(item?.formName || item?.FormName || '')) ||
        /\bform\s*[-"']?\s*vi\b/i.test(String(parsed?.formHeader?.title || '')) ||
        /\bform\s*[-"']?\s*6\b/i.test(String(parsed?.formHeader?.title || ''));

      const countFilledRowCells = (row, headers) => {
        if (!row || typeof row !== 'object' || !Array.isArray(headers)) return 0;
        let n = 0;
        for (let hi = 0; hi < headers.length; hi++) {
          const h = headers[hi];
          const v = row[h];
          if (v != null && String(v).trim() !== '') n++;
        }
        return n;
      };

      // If Autofill modal is open for this statutory line, export the same rows as on screen (avoids stale SampleData).
      const modalRow = formFileModalData?.item || autofillItem;
      const nidDownload =
        resolveNumericStatutoryIdForProofRow(item) ||
        resolveNumericStatutoryIdForProofRow(resolvedFormFileItem);
      const nidModal = modalRow ? resolveNumericStatutoryIdForProofRow(modalRow) : null;
      const dedupD = proofRowDedupKey(item);
      const dedupM = modalRow ? proofRowDedupKey(modalRow) : '';
      const sameStatutoryLineAsModal =
        !!modalRow &&
        !!isFormFileModalOpen &&
        ((nidDownload && nidModal && String(nidDownload) === String(nidModal)) ||
          !!(dedupD && dedupM && dedupD === dedupM));

      if (
        sameStatutoryLineAsModal &&
        Array.isArray(formTableDataRef.current) &&
        formTableDataRef.current.length > 0
      ) {
        usedLiveModalGrid = true;
        mappedData = formTableDataRef.current.map((row) => ({ ...row }));
        const modalHdrs = formFileModalData?.parsedTableHeaders;
        if (Array.isArray(modalHdrs) && modalHdrs.length > 0) {
          headersToUse = [...modalHdrs];
        }
      }

      // Form U hard preference: if live Autofill grid has rows, use it directly.
      // Row-id matching can fail for merged/statutory placeholders and then stale draft gets downloaded.
      if (
        isFormUDownload &&
        Array.isArray(formTableDataRef.current) &&
        formTableDataRef.current.length > 0
      ) {
        usedLiveModalGrid = true;
        mappedData = formTableDataRef.current.map((row) => ({ ...row }));
        const modalHdrs = formFileModalData?.parsedTableHeaders;
        if (Array.isArray(modalHdrs) && modalHdrs.length > 0) {
          headersToUse = [...modalHdrs];
        } else if (Array.isArray(mappedData[0])) {
          headersToUse = Array.from({ length: mappedData[0].length }, (_, idx) => `Column ${idx + 1}`);
        } else if (mappedData[0] && typeof mappedData[0] === 'object') {
          headersToUse = Object.keys(mappedData[0]);
        }
        savedDraftRowMatrix = null;
      }

      // Form 12 hard preference: use live Autofill grid rows directly when available
      // so alignment matches the modal (avoid stale/shifted draft matrix rows).
      if (
        isForm12Download &&
        Array.isArray(formTableDataRef.current) &&
        formTableDataRef.current.length > 0
      ) {
        usedLiveModalGrid = true;
        mappedData = formTableDataRef.current.map((row) => ({ ...row }));
        const modalHdrs = formFileModalData?.parsedTableHeaders;
        if (Array.isArray(modalHdrs) && modalHdrs.length > 0) {
          headersToUse = [...modalHdrs];
        } else if (mappedData[0] && typeof mappedData[0] === 'object') {
          headersToUse = Object.keys(mappedData[0]);
        }
        savedDraftRowMatrix = null;
      }
      // Form 10 hard preference: use live Autofill grid rows directly when available.
      if (
        isForm10Download &&
        Array.isArray(formTableDataRef.current) &&
        formTableDataRef.current.length > 0
      ) {
        usedLiveModalGrid = true;
        mappedData = formTableDataRef.current.map((row) => ({ ...row }));
        const modalHdrs = formFileModalData?.parsedTableHeaders;
        if (Array.isArray(modalHdrs) && modalHdrs.length > 0) {
          headersToUse = [...modalHdrs];
        } else if (mappedData[0] && typeof mappedData[0] === 'object') {
          headersToUse = Object.keys(mappedData[0]);
        }
        savedDraftRowMatrix = null;
      }
      if (
        isForm25Download &&
        Array.isArray(formTableDataRef.current) &&
        formTableDataRef.current.length > 0
      ) {
        usedLiveModalGrid = true;
        mappedData = formTableDataRef.current.map((row) => ({ ...row }));
        const modalHdrs = formFileModalData?.parsedTableHeaders;
        if (Array.isArray(modalHdrs) && modalHdrs.length > 0) {
          headersToUse = [...modalHdrs];
        } else if (mappedData[0] && typeof mappedData[0] === 'object') {
          headersToUse = Object.keys(mappedData[0]);
        }
        savedDraftRowMatrix = null;
      }
      if (
        isFormBDownload &&
        Array.isArray(formTableDataRef.current) &&
        formTableDataRef.current.length > 0
      ) {
        usedLiveModalGrid = true;
        mappedData = formTableDataRef.current.map((row) => ({ ...row }));
        const modalHdrs = formFileModalData?.parsedTableHeaders;
        if (Array.isArray(modalHdrs) && modalHdrs.length > 0) {
          headersToUse = [...modalHdrs];
        } else if (mappedData[0] && typeof mappedData[0] === 'object') {
          headersToUse = Object.keys(mappedData[0]);
        }
        savedDraftRowMatrix = null;
      }
      if (
        isFormWDownload &&
        Array.isArray(formTableDataRef.current) &&
        formTableDataRef.current.length > 0
      ) {
        usedLiveModalGrid = true;
        mappedData = formTableDataRef.current.map((row) => ({ ...row }));
        const modalHdrs = formFileModalData?.parsedTableHeaders;
        if (Array.isArray(modalHdrs) && modalHdrs.length > 0) {
          headersToUse = [...modalHdrs];
        } else if (mappedData[0] && typeof mappedData[0] === 'object') {
          headersToUse = Object.keys(mappedData[0]);
        }
        savedDraftRowMatrix = null;
      }
      if (
        isFormXDownload &&
        Array.isArray(formTableDataRef.current) &&
        formTableDataRef.current.length > 0
      ) {
        usedLiveModalGrid = true;
        mappedData = formTableDataRef.current.map((row) => ({ ...row }));
        const modalHdrs = formFileModalData?.parsedTableHeaders;
        if (Array.isArray(modalHdrs) && modalHdrs.length > 0) {
          headersToUse = [...modalHdrs];
        } else if (mappedData[0] && typeof mappedData[0] === 'object') {
          headersToUse = Object.keys(mappedData[0]);
        }
        savedDraftRowMatrix = null;
      }
      if (
        isFormCDownload &&
        Array.isArray(formTableDataRef.current) &&
        formTableDataRef.current.length > 0
      ) {
        usedLiveModalGrid = true;
        mappedData = formTableDataRef.current.map((row) => ({ ...row }));
        const modalHdrs = formFileModalData?.parsedTableHeaders;
        if (Array.isArray(modalHdrs) && modalHdrs.length > 0) {
          headersToUse = [...modalHdrs];
        } else if (mappedData[0] && typeof mappedData[0] === 'object') {
          headersToUse = Object.keys(mappedData[0]);
        }
        savedDraftRowMatrix = null;
      }
      if (
        isFormDDownload &&
        Array.isArray(formTableDataRef.current) &&
        formTableDataRef.current.length > 0
      ) {
        usedLiveModalGrid = true;
        mappedData = formTableDataRef.current.map((row) => ({ ...row }));
        const modalHdrs = formFileModalData?.parsedTableHeaders;
        if (Array.isArray(modalHdrs) && modalHdrs.length > 0) {
          headersToUse = [...modalHdrs];
        } else if (mappedData[0] && typeof mappedData[0] === 'object') {
          headersToUse = Object.keys(mappedData[0]);
        }
        savedDraftRowMatrix = null;
      }
      if (
        isFormIDownload &&
        Array.isArray(formTableDataRef.current) &&
        formTableDataRef.current.length > 0
      ) {
        usedLiveModalGrid = true;
        mappedData = formTableDataRef.current.map((row) => ({ ...row }));
        const modalHdrs = formFileModalData?.parsedTableHeaders;
        if (Array.isArray(modalHdrs) && modalHdrs.length > 0) {
          headersToUse = [...modalHdrs];
        } else if (mappedData[0] && typeof mappedData[0] === 'object') {
          headersToUse = Object.keys(mappedData[0]);
        }
        savedDraftRowMatrix = null;
      }
      if (
        isFormVDownload &&
        Array.isArray(formTableDataRef.current) &&
        formTableDataRef.current.length > 0
      ) {
        usedLiveModalGrid = true;
        mappedData = formTableDataRef.current.map((row) => ({ ...row }));
        const modalHdrs = formFileModalData?.parsedTableHeaders;
        if (Array.isArray(modalHdrs) && modalHdrs.length > 0) {
          headersToUse = [...modalHdrs];
        } else if (mappedData[0] && typeof mappedData[0] === 'object') {
          headersToUse = Object.keys(mappedData[0]);
        }
        savedDraftRowMatrix = null;
      }
      if (
        isFormVIDownload &&
        Array.isArray(formTableDataRef.current) &&
        formTableDataRef.current.length > 0
      ) {
        usedLiveModalGrid = true;
        mappedData = formTableDataRef.current.map((row) => ({ ...row }));
        const modalHdrs = formFileModalData?.parsedTableHeaders;
        if (Array.isArray(modalHdrs) && modalHdrs.length > 0) {
          headersToUse = [...modalHdrs];
        } else if (mappedData[0] && typeof mappedData[0] === 'object') {
          headersToUse = Object.keys(mappedData[0]);
        }
        savedDraftRowMatrix = null;
      }

      // Prefer latest saved snapshot from SampleData for this statutory record (includes imported data).
      // Must resolve real Catalyst ROWID: merged ChecklistBulk rows keep bulk_* id while draft lives on draftStatutoryRowIdForFile.
      if (!usedLiveModalGrid) {
        const sampleStatutoryId =
          resolveNumericStatutoryIdForProofRow(item) ||
          resolveNumericStatutoryIdForProofRow(resolvedFormFileItem);
        if (sampleStatutoryId && isNumericStatutoryBackendId(sampleStatutoryId)) {
          try {
            const sampleResp = await fetch(`/server/statutoryreg_function/statutory/${sampleStatutoryId}/sampledata?_ts=${downloadReqTs}`);
            if (sampleResp.ok) {
              const sampleJson = await sampleResp.json().catch(() => null);
              const sampleBlock = sampleJson?.data?.sampleData;
              const sampleRows = sampleBlock?.rows;
              const snapshotHeaders = Array.isArray(sampleBlock?.headers) ? sampleBlock.headers : null;
              if (Array.isArray(sampleRows) && sampleRows.length > 0) {
                usedSnapshot = true;
                mappedData = sampleRows;
                // Rows were persisted under these headers; prefer them so keys align with the modal / Autofill save.
                if (snapshotHeaders && snapshotHeaders.length > 0) {
                  headersToUse = [...snapshotHeaders];
                }
              }
            }
          } catch (sampleErr) {
            console.warn('Could not fetch SampleData snapshot for draft generation:', sampleErr);
          }
        }
      }

      // Prefer saved Draft file rows when available (DraftFileName-backed data), then fit those rows into original template.
      if (!usedLiveModalGrid && (isFormADownload || isFormUDownload || isForm12Download || isForm10Download || isForm25Download || isFormBDownload || isFormWDownload || isFormXDownload || isFormCDownload || isFormDDownload || isFormIDownload || isFormVDownload || isFormVIDownload || !Array.isArray(mappedData) || mappedData.length === 0)) {
        const draftStatutoryId =
          resolveNumericStatutoryIdForProofRow(item) ||
          resolveNumericStatutoryIdForProofRow(resolvedFormFileItem);
        if (draftStatutoryId && isNumericStatutoryBackendId(draftStatutoryId)) {
          try {
            const draftFileResp = await fetch(`/server/statutoryreg_function/statutory/${draftStatutoryId}/file/Draft?_ts=${downloadReqTs}`);
            if (draftFileResp.ok) {
              const draftArrayBuffer = await draftFileResp.arrayBuffer();
              const draftWb = XLSX.read(draftArrayBuffer, { type: 'array' });
              const matrixRows = extractRowMatrixFromSavedDraft({
                draftWorkbook: draftWb,
                templateHeaderRowIndex: parsed.headerRowIndex,
                templateDataStartIndex: parsed.dataStartIndex,
                hasSubColumns: parsed.subColumns && Object.keys(parsed.subColumns || {}).length > 0
              });
              if (Array.isArray(matrixRows) && matrixRows.length > 0) {
                savedDraftRowMatrix = matrixRows;
              }
              const draftRows = extractMappedRowsFromSavedDraft({
                draftWorkbook: draftWb,
                templateHeaders: headersToUse,
                templateHeaderRowIndex: parsed.headerRowIndex,
                templateDataStartIndex: parsed.dataStartIndex,
                hasSubColumns: parsed.subColumns && Object.keys(parsed.subColumns || {}).length > 0
              });
              if (Array.isArray(draftRows) && draftRows.length > 0) {
                mappedData = draftRows;
                usedSavedDraftFile = true;
              }
            }
          } catch (draftErr) {
            console.warn('Could not read saved draft file data for template download:', draftErr);
          }
        }
      }

      // Form A requirement: include submitted rows from all matching Form A statutory submissions
      // (same base form + site + month context), not only the current row.
      if (isFormADownload && !usedLiveModalGrid && Array.isArray(statutoryData) && statutoryData.length > 0) {
        try {
          const currentFormKey = baseFormNameKey(item?.formName || item?.FormName);
          const currentSiteNorm = String(resolveSiteDisplayName(item, statutoryData) || item?.site || item?.Site || '')
            .trim()
            .toLowerCase();
          const currentMonthNorm = statutoryDedupeMonthNorm(item, selectedMonth);
          const candidateIds = statutoryData
            .filter((row) => {
              if (!isNumericStatutoryBackendId(row?.id)) return false;
              if (baseFormNameKey(row?.formName || row?.FormName) !== currentFormKey) return false;
              const rowSiteNorm = String(resolveSiteDisplayName(row, statutoryData) || row?.site || row?.Site || '')
                .trim()
                .toLowerCase();
              if (currentSiteNorm && rowSiteNorm && !resolvedSitesOverlap(currentSiteNorm, rowSiteNorm)) return false;
              const rowMonthNorm = statutoryDedupeMonthNorm(row, selectedMonth);
              if (currentMonthNorm && currentMonthNorm !== 'nomonth' && rowMonthNorm !== currentMonthNorm) return false;
              return true;
            })
            .map((row) => String(row.id).trim());

          if (candidateIds.length > 0) {
            const sampleBlocks = await Promise.all(
              candidateIds.map(async (id) => {
                try {
                  const r = await fetch(`/server/statutoryreg_function/statutory/${id}/sampledata`);
                  if (!r.ok) return null;
                  const j = await r.json().catch(() => null);
                  return j?.data?.sampleData || null;
                } catch {
                  return null;
                }
              })
            );
            const allRows = [];
            sampleBlocks.forEach((block) => {
              const rows = block?.rows;
              if (Array.isArray(rows) && rows.length > 0) {
                rows.forEach((rw) => allRows.push(rw));
                if (Array.isArray(block?.headers) && block.headers.length > 0 && (!headersToUse || headersToUse.length === 0)) {
                  headersToUse = [...block.headers];
                }
              }
            });
            if (allRows.length > 0) {
              mappedData = allRows;
              usedSnapshot = true;
            }
          }
        } catch (allRowsErr) {
          console.warn('Could not aggregate all submitted Form A rows:', allRowsErr);
        }
      }

      // Fallback to Zoho data if snapshot is not available (skip employee expansion for LWF Form C–style fixed grids).
      if (!Array.isArray(mappedData) || mappedData.length === 0) {
        if (isFixedRowAggregateComplianceTable(headersToUse)) {
          mappedData = parsed.tableData || [];
        } else {
          mappedData = await fetchAndPopulateEmployeeData(headersToUse, {
            returnMappedData: true,
            ...(isFormADownload ? { formAPeopleSkipManualColumns: true } : {})
          });
        }
      }

      // Old saves sometimes left a single sparse SampleData row (e.g. one employee column) while Zoho Autofill has many rows.
      // Refresh from Zoho only when the snapshot row is barely filled — avoids wiping intentional single-row drafts that filled many columns.
      if (
        usedSnapshot &&
        !usedLiveModalGrid &&
        !usedSavedDraftFile &&
        !hasStatutoryDraftFileRef(item) &&
        !hasStatutoryDraftFileRef(resolvedFormFileItem) &&
        Array.isArray(mappedData) &&
        mappedData.length === 1 &&
        !isFixedRowAggregateComplianceTable(headersToUse)
      ) {
        const filled = countFilledRowCells(mappedData[0], headersToUse);
        try {
          const zohoRows = await fetchAndPopulateEmployeeData(headersToUse, {
            returnMappedData: true,
            ...(isFormADownload ? { formAPeopleSkipManualColumns: true } : {})
          });
          if (Array.isArray(zohoRows) && zohoRows.length > mappedData.length && filled <= 2) {
            mappedData = zohoRows;
          }
        } catch (_) {
          /* keep snapshot */
        }
      }

      // Form U: snapshot/draft can contain placeholder numeric index row (1..17) without employee data.
      // If rows look numeric-only, force refresh from live employee mapping so template gets actual values.
      if (
        isFormUDownload &&
        Array.isArray(mappedData) &&
        mappedData.length > 0
      ) {
        const rowLooksPlaceholder = (row) => {
          if (Array.isArray(row)) {
            const cells = row.map((v) => String(v ?? '').trim()).filter(Boolean);
            if (cells.length === 0) return true;
            const hasAlpha = cells.some((v) => /[a-z]/i.test(v));
            if (hasAlpha) return false;
            return cells.every((v) => /^-?\d+(\.\d+)?$/.test(v));
          }
          if (row && typeof row === 'object') {
            const vals = Object.values(row).map((v) => String(v ?? '').trim()).filter(Boolean);
            if (vals.length === 0) return true;
            const hasAlpha = vals.some((v) => /[a-z]/i.test(v));
            if (hasAlpha) return false;
            return vals.every((v) => /^-?\d+(\.\d+)?$/.test(v));
          }
          return true;
        };
        const placeholderOnly = mappedData.every((row) => rowLooksPlaceholder(row));
        if (placeholderOnly) {
          try {
            const zohoRows = await fetchAndPopulateEmployeeData(headersToUse, { returnMappedData: true });
            if (Array.isArray(zohoRows) && zohoRows.length > 0) {
              mappedData = zohoRows;
              savedDraftRowMatrix = null;
              usedSnapshot = false;
              usedSavedDraftFile = false;
            }
          } catch (_) {
            // keep existing mappedData if refresh fails
          }
        }
      }

      // Form U hard fallback: if still blank/placeholder, remap with canonical full Form U headers.
      if (
        isFormUDownload &&
        (!Array.isArray(mappedData) || mappedData.length === 0)
      ) {
        try {
          const canonicalHeaders = Array.isArray(FORM_U_FORMAT2_HEADERS) && FORM_U_FORMAT2_HEADERS.length > 0
            ? [...FORM_U_FORMAT2_HEADERS]
            : headersToUse;
          const zohoRowsStrict = await fetchAndPopulateEmployeeData(canonicalHeaders, { returnMappedData: true });
          if (Array.isArray(zohoRowsStrict) && zohoRowsStrict.length > 0) {
            headersToUse = canonicalHeaders;
            mappedData = zohoRowsStrict;
            savedDraftRowMatrix = null;
            usedSnapshot = false;
            usedSavedDraftFile = false;
          }
        } catch (_) {
          // keep previous mappedData
        }
      }

      // Form 12 strict alignment rule:
      // Rebuild from fresh autofill mapping (same model row -> all rows) unless live modal rows are already in use.
      if (isForm12Download && !usedLiveModalGrid) {
        try {
          const canonicalForm12Headers = [...FORM_12_ALIGNMENT_HEADERS];
          const refreshedForm12Rows = await fetchAndPopulateEmployeeData(canonicalForm12Headers, { returnMappedData: true });
          if (Array.isArray(refreshedForm12Rows) && refreshedForm12Rows.length > 0) {
            headersToUse = canonicalForm12Headers;
            mappedData = refreshedForm12Rows;
            savedDraftRowMatrix = null;
            usedSnapshot = false;
            usedSavedDraftFile = false;
          }
        } catch (_) {
          // keep existing mappedData if refresh fails
        }
      }
      if (isForm10Download && !usedLiveModalGrid) {
        try {
          const refreshedForm10Rows = await fetchAndPopulateEmployeeData(headersToUse, { returnMappedData: true });
          if (Array.isArray(refreshedForm10Rows) && refreshedForm10Rows.length > 0) {
            mappedData = refreshedForm10Rows;
            savedDraftRowMatrix = null;
            usedSnapshot = false;
            usedSavedDraftFile = false;
          }
        } catch (_) {
          // keep existing mappedData if refresh fails
        }
      }
      if (isForm25Download && !usedLiveModalGrid) {
        try {
          const refreshedForm25Rows = await fetchAndPopulateEmployeeData(headersToUse, {
            returnMappedData: true,
            form25PeopleSkipManualColumns: true
          });
          if (Array.isArray(refreshedForm25Rows) && refreshedForm25Rows.length > 0) {
            mappedData = refreshedForm25Rows;
            savedDraftRowMatrix = null;
            usedSnapshot = false;
            usedSavedDraftFile = false;
          }
        } catch (_) {
          // keep existing mappedData if refresh fails
        }
      }
      if (isFormBDownload && !usedLiveModalGrid) {
        try {
          const refreshedFormBRows = await fetchAndPopulateEmployeeData(headersToUse, {
            returnMappedData: true,
            formBLabourWelfareAutofill: true
          });
          if (Array.isArray(refreshedFormBRows) && refreshedFormBRows.length > 0) {
            mappedData = refreshedFormBRows;
            savedDraftRowMatrix = null;
            usedSnapshot = false;
            usedSavedDraftFile = false;
          }
        } catch (_) {
          // keep existing mappedData if refresh fails
        }
      }
      if (isFormWDownload && !usedLiveModalGrid) {
        try {
          const refreshedFormWRows = await fetchAndPopulateEmployeeData(headersToUse, { returnMappedData: true });
          if (Array.isArray(refreshedFormWRows) && refreshedFormWRows.length > 0) {
            mappedData = refreshedFormWRows;
            savedDraftRowMatrix = null;
            usedSnapshot = false;
            usedSavedDraftFile = false;
          }
        } catch (_) {
          // keep existing mappedData if refresh fails
        }
      }
      if (isFormXDownload && !usedLiveModalGrid) {
        try {
          const refreshedFormXRows = await fetchAndPopulateEmployeeData(headersToUse, { returnMappedData: true });
          if (Array.isArray(refreshedFormXRows) && refreshedFormXRows.length > 0) {
            mappedData = refreshedFormXRows;
            savedDraftRowMatrix = null;
            usedSnapshot = false;
            usedSavedDraftFile = false;
          }
        } catch (_) {
          // keep existing mappedData if refresh fails
        }
      }
      if (isFormCDownload && !usedLiveModalGrid) {
        try {
          const refreshedFormCRows = await fetchAndPopulateEmployeeData(headersToUse, { returnMappedData: true });
          if (Array.isArray(refreshedFormCRows) && refreshedFormCRows.length > 0) {
            mappedData = refreshedFormCRows;
            savedDraftRowMatrix = null;
            usedSnapshot = false;
            usedSavedDraftFile = false;
          }
        } catch (_) {
          // keep existing mappedData if refresh fails
        }
      }
      if (isFormDDownload && !usedLiveModalGrid) {
        try {
          const refreshedFormDRows = await fetchAndPopulateEmployeeData(headersToUse, { returnMappedData: true });
          if (Array.isArray(refreshedFormDRows) && refreshedFormDRows.length > 0) {
            mappedData = refreshedFormDRows;
            savedDraftRowMatrix = null;
            usedSnapshot = false;
            usedSavedDraftFile = false;
          }
        } catch (_) {
          // keep existing mappedData if refresh fails
        }
      }
      if (isFormIDownload && !usedLiveModalGrid) {
        try {
          const refreshedFormIRows = await fetchAndPopulateEmployeeData(headersToUse, { returnMappedData: true });
          if (Array.isArray(refreshedFormIRows) && refreshedFormIRows.length > 0) {
            mappedData = refreshedFormIRows;
            savedDraftRowMatrix = null;
            usedSnapshot = false;
            usedSavedDraftFile = false;
          }
        } catch (_) {
          // keep existing mappedData if refresh fails
        }
      }
      if (isFormVDownload && !usedLiveModalGrid) {
        try {
          const refreshedFormVRows = await fetchAndPopulateEmployeeData(headersToUse, { returnMappedData: true });
          if (Array.isArray(refreshedFormVRows) && refreshedFormVRows.length > 0) {
            mappedData = refreshedFormVRows;
            savedDraftRowMatrix = null;
            usedSnapshot = false;
            usedSavedDraftFile = false;
          }
        } catch (_) {
          // keep existing mappedData if refresh fails
        }
      }
      if (isFormVIDownload && !usedLiveModalGrid) {
        try {
          const refreshedFormVIRows = await fetchAndPopulateEmployeeData(headersToUse, { returnMappedData: true });
          if (Array.isArray(refreshedFormVIRows) && refreshedFormVIRows.length > 0) {
            mappedData = refreshedFormVIRows;
            savedDraftRowMatrix = null;
            usedSnapshot = false;
            usedSavedDraftFile = false;
          }
        } catch (_) {
          // keep existing mappedData if refresh fails
        }
      }

      // Form U: if mapped rows are available from modal/sample/Zoho, do not let stale draft matrix override them.
      if (isFormUDownload && Array.isArray(mappedData) && mappedData.length > 0) {
        savedDraftRowMatrix = null;
      }
      // Form 12: same protection against stale draft-matrix row shape.
      if (isForm12Download && Array.isArray(mappedData) && mappedData.length > 0) {
        savedDraftRowMatrix = null;
      }
      if (isForm10Download && Array.isArray(mappedData) && mappedData.length > 0) {
        savedDraftRowMatrix = null;
      }
      if (isForm25Download && Array.isArray(mappedData) && mappedData.length > 0) {
        savedDraftRowMatrix = null;
      }
      if (isFormBDownload && Array.isArray(mappedData) && mappedData.length > 0) {
        savedDraftRowMatrix = null;
      }
      if (isFormWDownload && Array.isArray(mappedData) && mappedData.length > 0) {
        savedDraftRowMatrix = null;
      }
      if (isFormXDownload && Array.isArray(mappedData) && mappedData.length > 0) {
        savedDraftRowMatrix = null;
      }
      if (isFormCDownload && Array.isArray(mappedData) && mappedData.length > 0) {
        savedDraftRowMatrix = null;
      }
      if (isFormDDownload && Array.isArray(mappedData) && mappedData.length > 0) {
        savedDraftRowMatrix = null;
      }
      if (isFormIDownload && Array.isArray(mappedData) && mappedData.length > 0) {
        savedDraftRowMatrix = null;
      }
      if (isFormVDownload && Array.isArray(mappedData) && mappedData.length > 0) {
        savedDraftRowMatrix = null;
      }
      if (isFormVIDownload && Array.isArray(mappedData) && mappedData.length > 0) {
        savedDraftRowMatrix = null;
      }

      if (!mappedData || mappedData.length === 0) {
        setError('No data available to fill in the draft.');
        return;
      }
      let { blob, fileName } = isFormADownload
        ? await buildFormAWorkbookWithTemplateStyles({
            templateArrayBuffer: arrayBuffer,
            headersToUse,
            mappedData,
            mappedRowMatrix: savedDraftRowMatrix,
            templateFileId: templateMeta.formFile,
            parsedFormHeader: parsed.formHeader,
            formFileName: templateMeta.formFileName || resolvedFormFileItem.formFileName || 'form-draft.xlsx'
          })
        : isFormUDownload
          ? await buildFormUWorkbookWithTemplateStyles({
              templateArrayBuffer: arrayBuffer,
              headersToUse,
              mappedData,
              mappedRowMatrix: savedDraftRowMatrix,
              parsedFormHeader: parsed.formHeader,
              formFileName: templateMeta.formFileName || resolvedFormFileItem.formFileName || 'form-draft.xlsx'
            })
          : isForm12Download
            ? await buildForm12WorkbookWithTemplateStyles({
                templateArrayBuffer: arrayBuffer,
                mappedData,
                mappedRowMatrix: savedDraftRowMatrix,
                headersToUse,
                parsedFormHeader: parsed.formHeader,
                formFileName: templateMeta.formFileName || resolvedFormFileItem.formFileName || 'form-draft.xlsx'
              })
            : isForm10Download
              ? await buildForm10WorkbookWithTemplateStyles({
                  templateArrayBuffer: arrayBuffer,
                  mappedData,
                  mappedRowMatrix: savedDraftRowMatrix,
                  headersToUse,
                  parsedHeaderRowIndex: parsed.headerRowIndex,
                  parsedDataStartIndex: parsed.dataStartIndex,
                  parsedFormHeader: parsed.formHeader,
                  formFileName: templateMeta.formFileName || resolvedFormFileItem.formFileName || 'form-draft.xlsx'
                })
              : isForm25Download
                ? await buildForm25WorkbookWithTemplateStyles({
                    templateArrayBuffer: arrayBuffer,
                    mappedData,
                    mappedRowMatrix: savedDraftRowMatrix,
                    headersToUse,
                    parsedFormHeader: parsed.formHeader,
                    formFileName: templateMeta.formFileName || resolvedFormFileItem.formFileName || 'form-draft.xlsx'
                  })
                : isFormBDownload
                  ? await buildFormBWorkbookWithTemplateStyles({
                      templateArrayBuffer: arrayBuffer,
                      mappedData,
                      mappedRowMatrix: savedDraftRowMatrix,
                      headersToUse,
                      parsedHeaderRowIndex: parsed.headerRowIndex,
                      parsedDataStartIndex: parsed.dataStartIndex,
                      parsedFormHeader: parsed.formHeader,
                      formFileName: templateMeta.formFileName || resolvedFormFileItem.formFileName || 'form-draft.xlsx'
                    })
                  : isFormWDownload
                    ? await buildFormWWorkbookWithTemplateStyles({
                        templateArrayBuffer: arrayBuffer,
                        mappedData,
                        mappedRowMatrix: savedDraftRowMatrix,
                        headersToUse,
                        parsedHeaderRowIndex: parsed.headerRowIndex,
                        parsedDataStartIndex: parsed.dataStartIndex,
                        parsedFormHeader: parsed.formHeader,
                        formFileName: templateMeta.formFileName || resolvedFormFileItem.formFileName || 'form-draft.xlsx'
                      })
                  : isFormXDownload
                    ? await buildFormXWorkbookWithTemplateStyles({
                        templateArrayBuffer: arrayBuffer,
                        mappedData,
                        mappedRowMatrix: savedDraftRowMatrix,
                        headersToUse,
                        parsedHeaderRowIndex: parsed.headerRowIndex,
                        parsedDataStartIndex: parsed.dataStartIndex,
                        parsedFormHeader: parsed.formHeader,
                        formFileName: templateMeta.formFileName || resolvedFormFileItem.formFileName || 'form-draft.xlsx'
                      })
                  : isFormCDownload
                    ? await buildFormCWorkbookWithTemplateStyles({
                        templateArrayBuffer: arrayBuffer,
                        mappedData,
                        mappedRowMatrix: savedDraftRowMatrix,
                        headersToUse,
                        parsedHeaderRowIndex: parsed.headerRowIndex,
                        parsedDataStartIndex: parsed.dataStartIndex,
                        parsedFormHeader: parsed.formHeader,
                        formFileName: templateMeta.formFileName || resolvedFormFileItem.formFileName || 'form-draft.xlsx'
                      })
                    : isFormDDownload
                      ? await buildFormDWorkbookWithTemplateStyles({
                          templateArrayBuffer: arrayBuffer,
                          mappedData,
                          mappedRowMatrix: savedDraftRowMatrix,
                          headersToUse,
                          parsedHeaderRowIndex: parsed.headerRowIndex,
                          parsedDataStartIndex: parsed.dataStartIndex,
                          parsedFormHeader: parsed.formHeader,
                          formFileName: templateMeta.formFileName || resolvedFormFileItem.formFileName || 'form-draft.xlsx'
                        })
                      : isFormIDownload
                        ? await buildFormIWorkbookWithTemplateStyles({
                            templateArrayBuffer: arrayBuffer,
                            mappedData,
                            mappedRowMatrix: savedDraftRowMatrix,
                            headersToUse,
                            parsedHeaderRowIndex: parsed.headerRowIndex,
                            parsedDataStartIndex: parsed.dataStartIndex,
                            parsedFormHeader: parsed.formHeader,
                            formFileName: templateMeta.formFileName || resolvedFormFileItem.formFileName || 'form-draft.xlsx'
                          })
                        : isFormVDownload
                          ? await buildFormVWorkbookWithTemplateStyles({
                              templateArrayBuffer: arrayBuffer,
                              mappedData,
                              mappedRowMatrix: savedDraftRowMatrix,
                              headersToUse,
                              parsedHeaderRowIndex: parsed.headerRowIndex,
                              parsedDataStartIndex: parsed.dataStartIndex,
                              parsedFormHeader: parsed.formHeader,
                              formFileName: templateMeta.formFileName || resolvedFormFileItem.formFileName || 'form-draft.xlsx'
                            })
                          : isFormVIDownload
                            ? await buildFormVIWorkbookWithTemplateStyles({
                                templateArrayBuffer: arrayBuffer,
                                mappedData,
                                mappedRowMatrix: savedDraftRowMatrix,
                                headersToUse,
                                parsedHeaderRowIndex: parsed.headerRowIndex,
                                parsedDataStartIndex: parsed.dataStartIndex,
                                parsedFormHeader: parsed.formHeader,
                                formFileName: templateMeta.formFileName || resolvedFormFileItem.formFileName || 'form-draft.xlsx'
                              })
        : buildDraftWorkbook({
            currentItem: item,
            templateWb,
            headersToUse,
            headerRowIndex: parsed.headerRowIndex,
            dataStartIndex: parsed.dataStartIndex,
            parsedFormHeader: parsed.formHeader,
            headerFormData: {},
            formTableData: mappedData,
            hasSubColumns: parsed.subColumns && Object.keys(parsed.subColumns || {}).length > 0,
            formFileName: templateMeta.formFileName || resolvedFormFileItem.formFileName || 'form-draft.xlsx',
            allowTemplateFallback: false
          });
      blob = await appendApprovedStatutoryHeadHrSignSealToDownloadBlob(blob, item);
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = fileName;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(url);
      setSuccess(
        usedLiveModalGrid
          ? 'Draft downloaded from current Autofill grid.'
          : usedSavedDraftFile
            ? 'Template downloaded with saved draft data.'
          : 'Draft downloaded (includes saved/imported data).'
      );
      setTimeout(() => setSuccess(''), 4000);
    } catch (err) {
      console.error('View Draft File generate error:', err);
      setError(err.message || 'Failed to generate draft. Try opening the form and using Autofill.');
    } finally {
      setFormFileLoading(false);
    }
  };

  // Save form file data to Excel and upload as draft
  const handleSaveFormFile = async () => {
    try {
      setSubmitting(true);
      setError('');
      setSuccess('');
      const tableDataForSave = Array.isArray(formTableDataRef.current) ? formTableDataRef.current : [];

      if (!tableDataForSave || tableDataForSave.length === 0) {
        setError('No data to save');
        setSubmitting(false);
        return;
      }

      // Get the current item being viewed (if any)
      const currentItem = formFileModalData?.item || autofillItem;
      const isFromChecklist = currentItem?.isFromChecklist || (currentItem?.id && String(currentItem.id).startsWith('checklist_'));
     
      let headersToUse = formFileModalData?.parsedTableHeaders?.length ? formFileModalData.parsedTableHeaders : tableHeaders;
      let templateWb = formFileModalData?.rawData;
      let headerRowIndex = formFileModalData?.headerRowIndex;
      let dataStartIndex = formFileModalData?.dataStartIndex;
      let parsedFormHeaderForSave = formFileModalData?.parsedFormHeader || formHeader;
      const hasSubColumns = formFileModalData?.parsedSubColumns && Object.keys(formFileModalData.parsedSubColumns || {}).length > 0;

      // If we don't have the form file template in state, fetch it now so draft always uses form file format (not new Excel)
      if (!templateWb && currentItem && (currentItem.formFile || currentItem.formFileName) && !isFromChecklist) {
        try {
          let fileUrl;
          const isFromFormmaster = currentItem.isFromFormmaster && currentItem.formFile;
          if (isFromFormmaster) {
            const fn = (currentItem.formFileName || 'template.xlsx').replace(/[^a-zA-Z0-9._-]/g, '_');
            fileUrl = `/server/formmaster_function/templates/download/${currentItem.formFile}?fileName=${encodeURIComponent(fn)}`;
          } else if (currentItem.id && !String(currentItem.id).startsWith('formmaster_') && !String(currentItem.id).startsWith('bulk_')) {
            fileUrl = `/server/statutoryreg_function/statutory/${currentItem.id}/file/Form`;
          } else {
            const fn = (currentItem.formFileName || 'template.xlsx').replace(/[^a-zA-Z0-9._-]/g, '_');
            fileUrl = `/server/formmaster_function/templates/download/${currentItem.formFile}?fileName=${encodeURIComponent(fn)}`;
          }
          const res = await fetch(fileUrl);
          if (res.ok) {
            const arrayBuffer = await res.arrayBuffer();
            templateWb = XLSX.read(arrayBuffer, { type: 'array' });
            const parsed = parseExcelForm(templateWb);
            headerRowIndex = parsed.headerRowIndex;
            dataStartIndex = parsed.dataStartIndex;
            parsedFormHeaderForSave = parsed.formHeader || parsedFormHeaderForSave;
            if (!headersToUse.length && parsed.headers?.length) headersToUse = parsed.headers;
          }
        } catch (fetchErr) {
          console.warn('Could not fetch form file template for draft:', fetchErr);
        }
      }

      // When we have the form master template but parse didn't return indices, detect header/data row from sheet
      if (templateWb && headersToUse.length > 0 && (headerRowIndex == null || headerRowIndex < 0 || dataStartIndex == null || dataStartIndex < 0)) {
        const firstSheetName = templateWb.SheetNames[0];
        const templateWs = templateWb.Sheets[firstSheetName];
        const jsonData = XLSX.utils.sheet_to_json(templateWs, { header: 1, defval: '' });
        const firstHeader = (headersToUse[0] || '').trim().toLowerCase();
        for (let r = 0; r < Math.min(30, jsonData.length); r++) {
          const row = jsonData[r] || [];
          for (let c = 0; c < row.length; c++) {
            const cell = String(row[c] || '').trim().toLowerCase();
            const matchesFirstHeader = cell && (
              cell.includes(firstHeader) ||
              firstHeader.includes(cell) ||
              (firstHeader.includes('s.no') && (cell.includes('s.no') || cell.includes('serial')))
            );
            if (matchesFirstHeader) {
              headerRowIndex = r;
              dataStartIndex = hasSubColumns ? r + 2 : r + 1;
              break;
            }
          }
          if (headerRowIndex >= 0) break;
        }
      }
      if (!templateWb) {
        setError('Original statutory template not found. Save requires the existing form template.');
        setSubmitting(false);
        return;
      }
      const draftFileNameForSave = formFileModalData?.formFileName ||
        formHeader?.title?.replace(/[^a-zA-Z0-9]/g, '_') ||
        `Form_${Date.now()}.xlsx`;
      const { blob, fileName } = buildDraftWorkbook({
        currentItem,
        templateWb,
        headersToUse,
        headerRowIndex,
        dataStartIndex,
        parsedFormHeader: parsedFormHeaderForSave,
        headerFormData,
        formTableData: tableDataForSave,
        hasSubColumns,
        formFileName: draftFileNameForSave,
        allowTemplateFallback: false
      });
      const file = new File([blob], fileName, { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });

      // Upload draft file
      const draftFormData = new FormData();
      draftFormData.append('file', file);

      const draftUploadResp = await fetch('/server/statutoryreg_function/statutory/upload/Draft', {
        method: 'POST',
        body: draftFormData
      });

      let draftUploadData;
      const draftContentType = draftUploadResp.headers.get('content-type') || '';

      if (draftContentType.includes('application/json')) {
        try {
          draftUploadData = await draftUploadResp.json();
        } catch (jsonErr) {
          const textResponse = await draftUploadResp.text();
          console.error('Failed to parse JSON response:', textResponse);
          setError(`Draft file upload failed: ${textResponse || 'Invalid server response'}`);
          setSubmitting(false);
          return;
        }
      } else {
        const textResponse = await draftUploadResp.text();
        console.error('Non-JSON response received:', textResponse);
        setError(`Draft file upload failed: ${textResponse || 'Server error occurred'}`);
        setSubmitting(false);
        return;
      }

      if (!draftUploadResp.ok || draftUploadData.status !== 'success') {
        const errorMessage = draftUploadData.message || 'Draft file upload failed';
        console.error('Draft file upload error:', errorMessage);
        setError(errorMessage);
        setSubmitting(false);
        return;
      }

      const draftFileId = draftUploadData.fileId != null ? String(draftUploadData.fileId).trim() : null;
      const draftFileName = draftUploadData.fileName || fileName;

      // First save (Autofill → Save): draft only. Proof Submission is set only when user uploads PDF/image in that column or from Edit.

      // Prepare payload to save/update statutory record
      // Extract form name from title or use existing form name
      const formNameValue = currentItem?.formName ||
                           formHeader?.title?.split('\n')[0]?.trim() ||
                           fileName.replace('.xlsx', '').replace(/_/g, ' ');
     
      // Month filter saved so it displays in backend (Autofill → Save)
      const monthFilterValue = resolveMonthFilterForSave(
        currentItem?.dueDate,
        selectedMonth,
        currentItem?.monthFilter ?? currentItem?.MonthFilter ?? currentItem?.monthfilter
      );
      const normalizedCurrentFormName = String(currentItem?.formName || formNameValue || '')
        .toLowerCase()
        .replace(/\s+/g, ' ')
        .trim();
      // Fallback metadata source for site-scope logins:
      // when Autofill starts from formmaster/checklist placeholder rows, act/sector can be blank.
      // Then saved draft rows get filtered out in site view. Reuse metadata from any existing row of same form.
      const metadataSource = Array.isArray(statutoryData)
        ? (statutoryData.find((r) => {
            const rForm = String(r?.formName || r?.FormName || '')
              .toLowerCase()
              .replace(/\s+/g, ' ')
              .trim();
            if (!rForm || rForm !== normalizedCurrentFormName) return false;
            const hasAct = String(r?.act || r?.Act || '').trim() !== '';
            const hasSector = String(r?.sector || r?.Sector || '').trim() !== '';
            return hasAct || hasSector;
          }) || null)
        : null;
      // Only treat as update (PUT) when this id is an actual Statutory table record. statutoryData is merged (API + form master + bulk), so we must detect synthetic ids explicitly — backend only has numeric ROWIDs.
      const idStr = currentItem?.id != null ? String(currentItem.id).trim() : '';
      const isSyntheticId = idStr.startsWith('formmaster_') || idStr.startsWith('bulk_') || idStr.startsWith('checklist_') || idStr.startsWith('placeholder_');
      const isUpdate = !!currentItem?.id && !isFromChecklist && !isSyntheticId;

      const latestRecord = currentItem?.id && statutoryData?.length ? statutoryData.find(r => String(r.id) === String(currentItem.id)) : null;
      // Form file from SEMaster must never be touched by draft save: on PUT do not send formFile/formFileName at all
      const preservedFormFile = latestRecord?.formFile ?? latestRecord?.FormFile ?? currentItem?.formFile ?? currentItem?.FormFile ?? null;
      const preservedFormFileName = latestRecord?.formFileName ?? latestRecord?.FormFileName ?? currentItem?.formFileName ?? currentItem?.FormFileName ?? fileName;
      const hasFormFile = preservedFormFile && String(preservedFormFile).trim() !== '' && preservedFormFile !== 'null';

      // On update (PUT): never send formFile/formFileName so backend keeps existing form file. On create (POST): include when we have it.
      // draftOnly: true = Autofill save → only save draft; do not copy draft to proof. Proof is set only when user clicks Edit and saves.
      const payload = {
        ...(!isUpdate && hasFormFile && { formFile: preservedFormFile, formFileName: preservedFormFileName || fileName }),
        formName: currentItem?.formName || formNameValue,
        act: currentItem?.act || metadataSource?.act || metadataSource?.Act || formHeader?.subtitle || '',
        description: currentItem?.description || metadataSource?.description || metadataSource?.Description || '',
        sector: currentItem?.sector || metadataSource?.sector || metadataSource?.Sector || '',
        state: currentItem?.state || metadataSource?.state || metadataSource?.State || '',
        site:
          currentItem?.site ||
          currentItem?.Site ||
          metadataSource?.site ||
          metadataSource?.Site ||
          resolveSiteDisplayName(currentItem || metadataSource || {}, statutoryData) ||
          siteFromUrl ||
          '',
        dueDate: currentItem?.dueDate || metadataSource?.dueDate || metadataSource?.DueDate || '',
        monthFilter: monthFilterValue,
        MonthFilter: monthFilterValue,
        monthfilter: monthFilterValue,
        autofill: currentItem?.autofill || '',
        draft: 'Draft',
        draftOnly: true,
        draftFile: draftFileId,
        draftFileName: draftFileName || null,
        approval: (() => {
          const a = currentItem?.approval ?? currentItem?.Approval;
          return a != null && String(a).trim() !== '' ? String(a).trim() : null;
        })(),
        status: (() => {
          const s = currentItem?.status != null && currentItem.status !== '' ? currentItem.status : currentItem?.Status;
          return s != null && String(s).trim() !== '' ? String(s).trim() : null;
        })(),
        sendForApproval: (() => {
          const s = currentItem?.sendForApproval ?? currentItem?.SendForApproval;
          return s != null && String(s).trim() !== '' ? String(s).trim() : null;
        })(),
        remarks: (() => {
          const r = currentItem?.remarks ?? currentItem?.Remarks;
          return r != null && String(r).trim() !== '' ? String(r).trim() : null;
        })(),
        // Persist autofill/imported grid snapshot in SampleData table on backend save.
        sampleDataHeader: Array.isArray(headersToUse) ? headersToUse : [],
        sampleData: Array.isArray(tableDataForSave) ? tableDataForSave : []
      };

      // Autofill → Save: UPDATE if this is an existing Statutory record; otherwise POST to create (e.g. when opened from form master row).
      const sendDraftSaveRequest = async (requestPayload) => {
        if (isUpdate) {
          console.log('Updating existing record with draft file (no new record):', currentItem.id);
          const { formFile: _f, formFileName: _n, ...putPayload } = requestPayload;
          return fetch(`/server/statutoryreg_function/statutory/${currentItem.id}`, {
            method: 'PUT',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(putPayload)
          });
        }
        console.log('Creating new record with draft file (no existing statutory record)');
        return fetch('/server/statutoryreg_function/statutory', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(requestPayload)
        });
      };

      let response = await sendDraftSaveRequest(payload);

      let data;
      try {
        data = await response.json();
      } catch (jsonErr) {
        const textResponse = await response.text();
        console.error('Failed to parse response:', textResponse);
        setError(`Failed to save draft: ${textResponse || 'Invalid server response'}`);
        setSubmitting(false);
        return;
      }

      // Backend can return 400 for very large SampleData payloads.
      // Retry once without sampleData snapshot so Draft file save still succeeds.
      if (!response.ok && Number(response.status) === 400) {
        try {
          const { sampleDataHeader: _h, sampleData: _d, ...payloadWithoutSample } = payload;
          const retryResp = await sendDraftSaveRequest(payloadWithoutSample);
          if (retryResp.ok) {
            response = retryResp;
            data = await retryResp.json().catch(() => ({}));
          }
        } catch (retryErr) {
          console.warn('Retry without sampleData failed:', retryErr);
        }
      }

      if (response.ok && data.status === 'success') {
        setSuccess('Draft saved. To add Proof Submission, click the blue Edit button and save with a proof file.');
       
        // Delete stale Catalyst rows same form/act/month/site but no draft (common after POST from checklist bulk).
        const keeperStatutoryId =
          isUpdate && currentItem?.id != null
            ? String(currentItem.id).trim()
            : data?.data?.statutory?.id != null
              ? String(data.data.statutory.id).trim()
              : '';
        // Some backend paths with draftOnly=true may skip MonthFilter persistence.
        // Enforce month storage explicitly on the saved row so Jan/Feb filtering remains strict.
        if (/^\d+$/.test(keeperStatutoryId)) {
          try {
            await fetch(`/server/statutoryreg_function/statutory/${keeperStatutoryId}`, {
              method: 'PUT',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({
                monthFilter: monthFilterValue,
                MonthFilter: monthFilterValue,
                monthfilter: monthFilterValue
              })
            });
          } catch (monthPersistErr) {
            console.warn('MonthFilter follow-up persist failed (non-blocking):', monthPersistErr);
          }
        }
        if (/^\d+$/.test(keeperStatutoryId)) {
          try {
            const updatedData = await fetch('/server/statutoryreg_function/statutory').then(res => res.json());
            if (updatedData.status === 'success' && updatedData.data?.statutoryData) {
              const allRecords = updatedData.data.statutoryData;
             
              // Normalize form name for comparison
              const normalizedFormName = (formNameValue || '').toLowerCase().trim();
              const normalizedAct = (payload.act || '').toLowerCase().trim();
              const normalizedDescription = (payload.description || '').toLowerCase().trim();
              const normalizedSector = String(payload.sector || currentItem?.sector || currentItem?.Sector || '').trim().toLowerCase();
              const normalizedStateRow = String(payload.state || currentItem?.state || currentItem?.State || '').trim().toLowerCase();
             
              // Find duplicates only within the SAME logical month (match checklist bulk vs saved Statutory month norms).
              const mergeProbeCurrent = {
                ...currentItem,
                monthFilter: payload.monthfilter ?? payload.monthFilter ?? currentItem?.monthFilter,
                MonthFilter: payload.monthfilter ?? payload.monthFilter ?? currentItem?.MonthFilter,
                monthfilter: payload.monthfilter ?? payload.monthFilter ?? currentItem?.monthfilter,
                dueDate: payload.dueDate ?? currentItem?.dueDate
              };
              const currentMonthNorm = statutoryDedupeMonthNorm(mergeProbeCurrent, selectedMonth);
              const duplicates = allRecords.filter(item => {
                const itemFormName = (item.formName || '').toLowerCase().trim();
                const itemAct = (item.act || '').toLowerCase().trim();
                const itemDescription = (item.description || '').toLowerCase().trim();
                const itemSector = String(item.sector || item.Sector || '').trim().toLowerCase();
                const itemState = String(item.state || item.State || '').trim().toLowerCase();
                const itemMonthNorm = statutoryDedupeMonthNorm(item, selectedMonth);
                const itemFormFile = item.formFile || item.FormFile || null;
                const currentFormFile = currentItem?.formFile || currentItem?.FormFile || null;
               
                const isSameFormName = itemFormName === normalizedFormName;
                const isSameAct = itemAct === normalizedAct;
                const isSameDescription = itemDescription === normalizedDescription;
                const isSameSector = itemSector === normalizedSector;
                const isSameStateRow = itemState === normalizedStateRow;
                const isSameMonth = currentMonthNorm !== 'nomonth' && itemMonthNorm !== 'nomonth' && currentMonthNorm === itemMonthNorm;
                const isDifferentId = String(item.id) !== keeperStatutoryId;
                const df = item.draftFile ?? item.DraftFile;
                const hasNoDraftFile = !df || df === null || df === 'null' || String(df).trim() === '';
                const hasSameFormFile = currentFormFile && itemFormFile && String(itemFormFile) === String(currentFormFile);
               
                return isDifferentId &&
                       isSameFormName &&
                       isSameAct &&
                       isSameDescription &&
                       isSameSector &&
                       isSameStateRow &&
                       isSameMonth &&
                       hasNoDraftFile &&
                       (hasSameFormFile || !itemFormFile);
              });
             
              // Delete duplicates that don't have draft files
              if (duplicates.length > 0) {
                console.log(`Found ${duplicates.length} duplicate records to delete:`, duplicates.map(d => d.id));
                for (const duplicate of duplicates) {
                  try {
                    const deleteResponse = await fetch(`/server/statutoryreg_function/statutory/${duplicate.id}`, {
                      method: 'DELETE'
                    });
                    const deleteData = await deleteResponse.json();
                    if (deleteResponse.ok && deleteData.status === 'success') {
                      console.log(`Deleted duplicate record: ${duplicate.id}`);
                    } else {
                      console.warn(`Failed to delete duplicate record ${duplicate.id}:`, deleteData.message);
                    }
                  } catch (deleteErr) {
                    console.error(`Error deleting duplicate record ${duplicate.id}:`, deleteErr);
                  }
                }
              }
            }
          } catch (duplicateErr) {
            console.error('Error removing duplicates:', duplicateErr);
            // Don't fail the save operation if duplicate removal fails
          }
        }
       
        // Refresh data first, then close modal
        await fetchStatutoryData();
        setTimeout(() => {
          handleCloseFormFileModal();
        }, 500);
      } else {
        const errorMessage = data?.message || data?.error || `Failed to save draft (HTTP ${response.status})`;
        console.error('Save draft error:', errorMessage, data);
        setError(errorMessage);
      }
     
    } catch (err) {
      console.error('Error saving form file:', err);
      setError('Error saving form data: ' + err.message);
    } finally {
      setSubmitting(false);
    }
  };

  // Fetch employee data from Zoho People and populate form table
  // When options.returnMappedData is true, returns the mapped rows (for use by View Draft File generate-and-download)
  const fetchAndPopulateEmployeeData = async (headersToUse = null, options = {}) => {
    const returnMappedData = options.returnMappedData === true;
    const modalData = options.formFileModalData || formFileModalData;
    try {
      if (!returnMappedData) {
        setFormFileLoading(true);
        setError('');
        setSuccess('');
      }
     
      // Use provided headers or fall back to state
      const currentHeaders = headersToUse || (modalData?.parsedTableHeaders?.length ? modalData.parsedTableHeaders : tableHeaders);
     
      if (!currentHeaders || currentHeaders.length === 0) {
        console.error('⚠️ Table headers not available');
        if (!returnMappedData) setError('Table headers not available. Please try again.');
        if (!returnMappedData) setFormFileLoading(false);
        if (returnMappedData) return [];
        return;
      }

      // Form C / similar: template rows come from Excel — do not replace with one row per employee (that clears the grid).
      if (isFixedRowAggregateComplianceTable(currentHeaders)) {
        if (!returnMappedData) {
          setFormFileLoading(false);
          setSuccess(
            'This form uses fixed rows from your Excel template (not per-employee autofill). Edit values directly or Import.'
          );
          setTimeout(() => setSuccess(''), 5000);
        }
        if (returnMappedData) return [];
        return;
      }
     
      console.log('Using table headers:', currentHeaders);
     
      // Fetch employee data from peopledata_function
      const response = await fetch('/server/peopledata_function?form=employee&limit=100');
      const result = await response.json();
     
      console.log('Zoho People API Response:', result);
     
      if (!result.success) {
        throw new Error(result.error || 'Failed to fetch employee data');
      }
     
      if (!result.data) {
        throw new Error('No data received from Zoho People API');
      }
     
      // Extract employee records from Zoho People response
      // Structure: { response: { result: [{ "ID": [{ employee data }] }] } }
      let employees = [];
      const data = result.data;
     
      // Handle Zoho People nested structure
      if (data.response && data.response.result) {
        if (Array.isArray(data.response.result)) {
          // Flatten the nested structure: result is array of objects like { "ID": [employeeData] }
          data.response.result.forEach(resultItem => {
            if (typeof resultItem === 'object' && resultItem !== null) {
              // Each resultItem has keys that are IDs, values are arrays of employee objects
              Object.keys(resultItem).forEach(idKey => {
                const employeeArray = resultItem[idKey];
                if (Array.isArray(employeeArray)) {
                  // Add all employees from this array
                  employees.push(...employeeArray);
                } else if (employeeArray && typeof employeeArray === 'object') {
                  // Single employee object
                  employees.push(employeeArray);
                }
              });
            } else if (resultItem && typeof resultItem === 'object') {
              // Direct employee object
              employees.push(resultItem);
            }
          });
        } else if (data.response.result.Employee) {
          employees = [data.response.result];
        } else if (typeof data.response.result === 'object') {
          // Try to extract from object structure
          const keys = Object.keys(data.response.result);
          if (keys.length > 0) {
            const firstKey = keys[0];
            if (Array.isArray(data.response.result[firstKey])) {
              employees = data.response.result[firstKey];
            } else {
              employees = [data.response.result[firstKey]];
            }
          }
        }
      } else if (Array.isArray(data)) {
        employees = data;
      } else if (data.records && Array.isArray(data.records)) {
        employees = data.records;
      } else if (data.result) {
        if (Array.isArray(data.result)) {
          employees = data.result;
        } else if (data.result.Employee) {
          employees = [data.result];
        }
      }
     
      console.log('Extracted employees:', employees);
      console.log('Number of employees:', employees.length);
      console.log('Table headers:', currentHeaders);
     
      if (!Array.isArray(employees) || employees.length === 0) {
        console.error('No employees found. Full response:', JSON.stringify(result, null, 2));
        throw new Error('No employee data found in response. Please check the API response structure.');
      }

      /** Form B table: infer from column titles when modal title/act is incomplete (merged TN LWF headers). */
      const looksLikeFormBLabourWelfareEmployeeTable = () => {
        const joined = currentHeaders.map((h) => String(h || '').toLowerCase()).join('\n');
        if (/labour\s*welfare|tamil\s*nadu.*lwf|\blwf\b/i.test(joined) && /register\s+of\s+wages/i.test(joined)) {
          return true;
        }
        if (/\btotal\s+emoluments\b/.test(joined)) return true;
        if (/\bamounts?\s+deducted\b/.test(joined) && /\bother\s+deductions?\b/.test(joined)) return true;
        if (/total\s+number\s+of\s+empl/.test(joined)) return true;
        return false;
      };

      const formBAutofillContext =
        options.formBLabourWelfareAutofill === true ||
        looksLikeFormBLabourWelfareEmployeeTable() ||
        (isFormFileModalOpen &&
          isFormBLabourWelfareContext(
            formFileModalData?.parsedFormHeader || {},
            formFileModalData?.item,
            String(formFileModalData?.fileName || formFileModalData?.formFileName || ''),
            currentHeaders
          ));

      const isFormDRemunerationHeader = (h) => {
        const s = normalizeLooseHeaderText(h);
        return (
          s.includes('components of remuneration') ||
          s.includes('rate of remuneration') ||
          s.includes('basic wages') ||
          s.includes('basic wage') ||
          s.includes('wages or salary') ||
          (s.includes('basic') && s.includes('salary')) ||
          s.includes('dearness') ||
          /\bda\b/.test(s) ||
          s.includes('house rent allowance') ||
          s.includes('hra') ||
          (s.includes('house') && s.includes('rent')) ||
          s.includes('other allowance') ||
          (s.includes('other') && s.includes('allowance'))
        );
      };

      const isLikelyFormDCategoryHeader = (h) => {
        const s = normalizeLooseHeaderText(h);
        return s.includes('category of workers') || s.includes('category of worker');
      };

      const isLikelyFormDRowCountHeader = (h) => {
        const s = normalizeLooseHeaderText(h);
        return s.includes('noofmenemployed') || s.includes('numberofmenemployed') || s.includes('men employed') || s.includes('male employed') || s.includes('noofwomenemployed') || s.includes('numberofwomenemployed') || s.includes('women employed') || s.includes('female employed');
      };

      const formDAutofillContext = (() => {
        const hasRemunerationHeaders = currentHeaders.filter((h) => isFormDRemunerationHeader(h)).length >= 2;
        const hasFormDCounts =
          currentHeaders.some((h) => isLikelyFormDCategoryHeader(h)) ||
          currentHeaders.some((h) => isLikelyFormDRowCountHeader(h));
        const hasRateOfRemuneration = currentHeaders.some((h) => normalizeLooseHeaderText(h).includes('rate of remuneration'));
        return currentHeaders.some((h) => isFormDRemunerationHeader(h)) && (hasFormDCounts || hasRateOfRemuneration || hasRemunerationHeaders);
      })();

      const isFormVIFestivalManualHeader = (header) => {
        const s = String(header || '').toLowerCase().replace(/\s+/g, ' ').trim();
        if (!s) return false;
        if (s === 'remark' || s === 'remarks') return true;
        if (s.includes('enter days date')) return true;
        if (s.includes('enter days dates')) return true;
        if (s.includes('days dates') && s.includes('months')) return true;
        if (s.includes('national') && s.includes('festival') && s.includes('holiday')) return true;
        return /festival\s+holidays?/.test(s);
      };

      const formVIFestivalAutofillContext =
        isFormFileModalOpen &&
        isFormVIFestivalContext(
          formFileModalData?.parsedFormHeader || {},
          formFileModalData?.item,
          String(formFileModalData?.fileName || formFileModalData?.formFileName || ''),
          currentHeaders
        );

      const unwrapEmp = (empItem) => empItem.Employee || empItem.employee || empItem;

      const getFormBEmployeeFullName = (em) => {
        const fn = String(em.FirstName || em['FirstName'] || em.firstName || em['First Name'] || '').trim();
        const ln = String(em.LastName || em['LastName'] || em.lastName || em['Last Name'] || '').trim();
        if (fn && ln) return `${fn} ${ln}`;
        return fn || ln || '';
      };

      const isFormBTotalEmployeeCountHeader = (h) => {
        const s = String(h || '').toLowerCase().replace(/\s+/g, ' ').trim();
        if (/total\s+number\s+of\s+empl/.test(s)) return true;
        if (/total\s+number\s+of\s+employees?/.test(s)) return true;
        if (s.includes('total') && s.includes('number') && (s.includes('employee') || s.includes('empl'))) return true;
        return false;
      };

      /** Form B: TN LWF total emoluments column — fill from payroll earning "Basic Earnings". */
      const isFormBTotalEmolumentsBasicEarningsHeader = (h) => {
        const s = String(h || '').toLowerCase().replace(/\s+/g, ' ');
        if (!s.includes('emolument')) return false;
        return s.includes('payable') || s.includes('including') || s.includes('basic wage');
      };

      /** Form B: "Other deductions" sub-column — per product mapping, use payroll "Other Allowance" amount. */
      const isFormBOtherDeductionsOtherAllowanceHeader = (h) => {
        const s = String(h || '').toLowerCase().replace(/\s+/g, ' ');
        return /other\s+deductions?/.test(s) || /deducted.*other\s+deductions?/.test(s);
      };

      const isFormDBasicSalaryHeader = (h) => {
        const s = normalizeLooseHeaderText(h);
        return (
          s.includes('basic wages') ||
          s.includes('basic wage') ||
          s.includes('wages or salary') ||
          (s.includes('basic') && s.includes('salary'))
        );
      };

      const isFormDHouseRentAllowanceHeader = (h) => {
        const s = normalizeLooseHeaderText(h);
        return (
          s.includes('house rent allowance') ||
          s.includes('hra') ||
          (s.includes('house') && s.includes('rent'))
        );
      };

      const isFormDDearnessAllowanceHeader = (h) => {
        const s = normalizeLooseHeaderText(h);
        return s.includes('dearness') || /\bda\b/.test(s) || s.includes('dearness allowance');
      };

      const isFormDOtherAllowanceHeader = (h) => {
        const s = normalizeLooseHeaderText(h);
        return (
          s.includes('other allowance') ||
          (s.includes('other') && s.includes('allowance'))
        );
      };

      const isFormDNoMenEmployedHeader = (h) => {
        const s = normalizeLooseHeaderText(h);
        return (
          s.includes('men employed') ||
          s.includes('male employed') ||
          (s.includes('no') && s.includes('men') && s.includes('employed')) ||
          (s.includes('number') && s.includes('men') && s.includes('employed'))
        );
      };

      const isFormDNoWomenEmployedHeader = (h) => {
        const s = normalizeLooseHeaderText(h);
        return (
          s.includes('women employed') ||
          s.includes('female employed') ||
          s.includes('noofwomen') ||
          s.includes('no.of women') ||
          (s.includes('no') && s.includes('women') && s.includes('employed')) ||
          (s.includes('number') && s.includes('women') && s.includes('employed'))
        );
      };

      const isPlaceholderStatutoryCellValue = (value) => {
        const s = String(value || '').trim().toLowerCase();
        if (!s) return true;
        return /^enter\b/.test(s) || s.includes('enter ') || s.includes('select ') || s.includes('enter no');
      };

      /** Form B: avoid mapping People `employee_status` ("Active") into wage/balance columns. */
      const isFormBBalanceDueHeader = (h) => {
        const s = String(h || '').toLowerCase().replace(/\s+/g, ' ');
        return /balance\s+due\s+to\s+the\s+employees?/.test(s) || (s.includes('balance') && s.includes('due') && s.includes('employee'));
      };

      let formBTotalEmployeeCount = 0;
      if (formBAutofillContext) {
        formBTotalEmployeeCount = employees.filter((empItem) => {
          const em = unwrapEmp(empItem);
          const id = String(
            em.EmployeeID ||
              em['EmployeeID'] ||
              em['Employee ID'] ||
              em['Role.ID'] ||
              (em.Role && typeof em.Role === 'object' && em.Role.ID ? em.Role.ID : '') ||
              ''
          ).trim();
          const fn = String(em.FirstName || em['FirstName'] || '').trim();
          const ln = String(em.LastName || em['LastName'] || '').trim();
          return Boolean(id && fn && ln);
        }).length;
      }

      // Map employee data to form table structure
      // Helper function to sanitize values - ensure null/undefined/false become empty strings
      const sanitizeValue = (value) => {
        if (value === null || value === undefined || value === false) return '';
        if (typeof value === 'string') return value;
        // For boolean true, convert to string; for false, already handled above
        if (typeof value === 'boolean') return value ? 'true' : '';
        return String(value);
      };
     
      const genericHeaderRegex = /^column\s+([a-z]+|\d+)$/i;
      const genericHeadersCount = currentHeaders.filter((h) => genericHeaderRegex.test(String(h || '').trim())).length;
      const shouldUseGenericColumnFallback =
        currentHeaders.length > 0 && genericHeadersCount >= Math.max(3, Math.ceil(currentHeaders.length * 0.5));

      /** Headers like `17` or `DATES_17` — calendar day columns only attendance should fill. */
      const isLikelyDayOfMonthColumnHeader = (h) => {
        const t = String(h || '').trim();
        const direct = t.match(/^(\d{1,2})$/);
        if (direct) {
          const n = Number(direct[1]);
          return n >= 1 && n <= 31;
        }
        const suffixed = t.match(/(?:^|_)(\d{1,2})$/);
        if (suffixed) {
          const n = Number(suffixed[1]);
          return n >= 1 && n <= 31;
        }
        return false;
      };

      /** Form 25: these columns must stay manual / blank — People autofill mis-maps them via the generic header matcher. */
      const looksLikeForm25EmployeeTable = () => {
        const joined = currentHeaders.map((h) => String(h || '').toLowerCase()).join('\n');
        let score = 0;
        if (/name\s+of\s+the\s+worker/.test(joined)) score++;
        if (/worker\s+identity/.test(joined)) score++;
        if (/register\s+of\s+adult/.test(joined)) score++;
        if (/scheme\s+of\s+shifts/.test(joined)) score++;
        if (/time\s+at\s+which\s+work\s+commence/.test(joined)) score++;
        if (/rest\s+interval/.test(joined)) score++;
        if (/time\s+at\s+which\s+work\s+end/.test(joined)) score++;
        return score >= 2;
      };
      const isForm25LossOfPayDaysHeader = (h) => {
        const s = String(h || '').toLowerCase().replace(/\s+/g, ' ').trim();
        const n = s.replace(/[^a-z0-9]/g, '');
        if (!n.includes('lossofpay')) return false;
        return /\bday\b|\bdays\b|no\.?\s*of|number\s+of/.test(s) || n.includes('day') || n.includes('number');
      };
      const isForm25NationalHolidayBenefitHeader = (h) => {
        const s = String(h || '').toLowerCase().replace(/\s+/g, ' ').trim();
        if (!s.includes('national') || !s.includes('holiday')) return false;
        return s.includes('benefit') || s.includes('availed') || s.includes('working');
      };
      const form25SkipLossPayAndNationalHolidayBenefit =
        looksLikeForm25EmployeeTable() ||
        options.form25PeopleSkipManualColumns === true ||
        (isFormFileModalOpen &&
          /\bform\s*[-"']?\s*25\b/i.test(
            String(formFileModalData?.item?.formName || formFileModalData?.item?.FormName || '')
          ));

      /** Form A (Maternity Benefit muster): leave calendar / employment-day columns manual; fill name + age from People. */
      const looksLikeFormAMusterEmployeeTable = () => {
        const joined = currentHeaders.map((h) => String(h || '').toLowerCase()).join('\n');
        if (
          currentHeaders.some((h) => {
            const s = String(h || '').toLowerCase();
            return s.includes('woman') && /\bage\b/.test(s);
          })
        ) {
          return true;
        }
        if (/name\s+of\s+the\s+woman/.test(joined)) return true;
        if (/no\.?\s*of\s+days?\s+not\s+employed/.test(joined)) return true;
        if (/no\.?\s*of\s+days?\s+laid\s+off/.test(joined)) return true;
        if (/no\.?\s*of\s+days?\s+employed/.test(joined) && /(^|\n)month(\n|$)/.test(`\n${joined}\n`)) return true;
        return false;
      };
      const formAAutofillContext =
        looksLikeFormAMusterEmployeeTable() ||
        options.formAPeopleSkipManualColumns === true ||
        (isFormFileModalOpen &&
          (/\bform\s*[-"']?\s*a\b/i.test(String(formFileModalData?.item?.formName || formFileModalData?.item?.FormName || '')) ||
            /maternity\s+benefit/i.test(
              String(
                formFileModalData?.item?.Act ||
                  formFileModalData?.item?.act ||
                  formFileModalData?.parsedFormHeader?.title ||
                  formFileModalData?.parsedFormHeader?.reference ||
                  ''
              )
            )));
      const isFormAMonthOnlyColumn = (h) => {
        const s = String(h || '').toLowerCase().replace(/\s+/g, ' ').trim();
        const n = s.replace(/[^a-z0-9]/g, '');
        if (n === 'month') return true;
        return /^\(?\s*\d*\s*\)?\s*month\s*$/.test(s);
      };
      const isFormADaysEmployedColumn = (h) => {
        const s = String(h || '').toLowerCase();
        if (/not\s+employed/.test(s)) return false;
        return /no\.?\s*of\s+days?\s+employed/.test(s) || (/days?\s+of\s+employment/.test(s) && !/not\s+/.test(s));
      };
      const isFormALaidOffColumn = (h) => /no\.?\s*of\s+days?\s+laid\s+off/.test(String(h || '').toLowerCase()) ||
        (/\blaid\s+off\b/.test(String(h || '').toLowerCase()) && /\bday/.test(String(h || '').toLowerCase()));
      const isFormANotEmployedColumn = (h) => /no\.?\s*of\s+days?\s+not\s+employed/.test(String(h || '').toLowerCase()) ||
        (/\bnot\s+employed\b/.test(String(h || '').toLowerCase()) && /\bday/.test(String(h || '').toLowerCase()));
      /** Form A: one cell for both name and age (e.g. "(1) Name of the woman and age"). */
      const isFormACombinedWomanNameAndAgeHeader = (h) => {
        const s = String(h || '').toLowerCase();
        return s.includes('woman') && /\bage\b/.test(s);
      };

      /** Leave register / Form X style: statutory leave-wage columns are manual, not Zoho People. */
      const isLeaveWagesPerSection9Or10Header = (h) => {
        const s = String(h || '').toLowerCase().replace(/\s+/g, ' ').trim();
        if (!s.includes('leave') || !/wage/.test(s)) return false;
        return /section\s*9\b/.test(s) || /section\s*10\b/.test(s);
      };

      const isFormXLeaveMetricHeader = (h) => {
        const s = String(h || '').toLowerCase().replace(/\s+/g, ' ').trim();
        if (!s.includes('leave')) return false;
        if (s.includes('name of the employee') || s.includes('employee identification') || s.includes('gender')) return false;
        return (
          s.includes('beginning of the month') ||
          s.includes('earned during the period') ||
          s.includes('availed during the month') ||
          s.includes('balance at the end of the month')
        );
      };

      const getFallbackName = (emp) =>
        emp.FirstName ||
        emp['FirstName'] ||
        emp.firstName ||
        emp['First Name'] ||
        emp.First_Name ||
        emp.Name1 ||
        emp['Name1'] ||
        emp.Name ||
        emp['Employee Name'] ||
        '';

      const getFallbackEmployeeId = (emp) =>
        emp.EmployeeID ||
        emp['EmployeeID'] ||
        emp['Role.ID'] ||
        (emp.Role && typeof emp.Role === 'object' ? emp.Role.ID : '') ||
        emp['Employee ID'] ||
        '';

      const getFallbackWorkerId = (emp) =>
        emp.Zoho_ID ||
        emp['Zoho_ID'] ||
        emp.zoho_id ||
        emp['zoho_id'] ||
        emp.ZohoID ||
        emp['ZohoID'] ||
        '';

      const getEmployeeNameCandidates = (emp) => {
        if (!emp || typeof emp !== 'object') return [];
        const first = String(
          emp.FirstName ||
          emp['FirstName'] ||
          emp.firstName ||
          emp['First Name'] ||
          emp.First_Name ||
          emp['First_Name'] ||
          findValueByNormalizedKey(emp, 'firstname') ||
          ''
        ).trim();
        const last = String(
          emp.LastName ||
          emp['LastName'] ||
          emp.lastName ||
          emp['Last Name'] ||
          emp.Last_Name ||
          emp['Last_Name'] ||
          findValueByNormalizedKey(emp, 'lastname') ||
          ''
        ).trim();
        const full = `${first} ${last}`.trim();
        const raw = [
          full,
          getFallbackName(emp),
          emp.Name,
          emp['Name'],
          emp.EmployeeName,
          emp['Employee Name'],
          emp['EmployeeName'],
          emp.Nameoftheemployee,
          emp['Name of the employee'],
          emp['Nameoftheemployee'],
          first,
          last
        ];
        return Array.from(
          new Set(
            raw
              .map((v) => String(v || '').trim().toLowerCase())
              .filter(Boolean)
          )
        );
      };

      const getEmployeeLookupName = (emp) => {
        const candidates = getEmployeeNameCandidates(emp);
        return candidates.length > 0 ? candidates[0] : '';
      };

      const getEmployeeLookupIdCandidates = (emp, row = {}) => {
        const rowEntries = row && typeof row === 'object' ? Object.entries(row) : [];
        const rowIds = rowEntries
          .filter(([key, value]) => {
            if (value == null || String(value).trim() === '') return false;
            const header = String(key || '').toLowerCase();
            return (
              header === '__employeelookupid' ||
              ((header.includes('employee') || header.includes('worker') || header.includes('emp')) &&
                (header.includes('id') || header.includes('identification')))
            );
          })
          .map(([, value]) => String(value || '').trim());

        const empIds = [
          getPayrollEmployeeId(emp || {}),
          getFallbackEmployeeId(emp || {}),
          getFallbackWorkerId(emp || {}),
          emp?.employeeId,
          emp?.EmployeeId,
          emp?.['Employee Id'],
          emp?.workerId,
          emp?.WorkerId,
        ]
          .map((value) => String(value || '').trim())
          .filter(Boolean);

        return Array.from(new Set([...rowIds, ...empIds].filter(Boolean)));
      };

      const getRowPayrollLookupNameCandidates = (row = {}) => {
        if (!row || typeof row !== 'object') return [];
        const names = Object.entries(row)
          .filter(([key, value]) => {
            if (value == null || String(value).trim() === '') return false;
            const header = String(key || '').toLowerCase();
            return (
              header === '__employeelookupname' ||
              (header.includes('name') &&
                (header.includes('employee') || header.includes('worker') || header.includes('emp')))
            );
          })
          .map(([, value]) => normalizePayrollLookupValue(value))
          .filter(Boolean);
        return Array.from(new Set(names));
      };

      const getFallbackFatherOrSpouse = (emp) =>
        emp.Father_s_Name ||
        emp['Father_s_Name'] ||
        emp.Father_Name ||
        emp['Father_Name'] ||
        emp.Spouse_Name ||
        emp['Spouse_Name'] ||
        '';

      const getFallbackByColumnIndex = (emp, colIndex, rowIndex) => {
        const bankInfo = [emp.Account_Number, emp.Bank_Name, emp.Branch_Name].filter((v) => String(v || '').trim()).join(', ');
        const fallbackMap = {
          0: String(rowIndex + 1), // S.No
          1: getFallbackName(emp),
          2: getFallbackEmployeeId(emp),
          3: emp.Sex || emp.Gender || '',
          4: getFallbackFatherOrSpouse(emp),
          5: emp.Date_of_birth || emp['Date_of_birth'] || '',
          6: emp.Dateofjoining || emp['Dateofjoining'] || '',
          7: emp.Designation || '',
          8: emp.Address_Line_1 || '',
          9: emp.Address_Line_2 || '',
          10: getFallbackWorkerId(emp),
          11: emp.AadhaarNo || emp['Aadhaar No.'] || '',
          12: bankInfo
        };
        return sanitizeValue(fallbackMap[colIndex] || '');
      };

      const formWHeaderBlob = currentHeaders.map((h) => String(h || '').toLowerCase()).join(' | ');
      const formWFileHint = String(
        formFileModalData?.fileName ||
          formFileModalData?.formFileName ||
          formFileModalData?.item?.formName ||
          formFileModalData?.item?.FormName ||
          ''
      ).toLowerCase();
      const isLikelyFormW =
        (/form\s*[-_]?\s*w\b/.test(formWFileHint) || /\bform\s*w\b/.test(formWFileHint)) ||
        ((/basic\s*wag/.test(formWHeaderBlob) || formWHeaderBlob.includes('basic wage')) &&
          formWHeaderBlob.includes('dearness') &&
          formWHeaderBlob.includes('overtime')) ||
        (formWHeaderBlob.includes('register') &&
          formWHeaderBlob.includes('wage') &&
          (/basic\s*wag/.test(formWHeaderBlob) || formWHeaderBlob.includes('basic wage')));
      const isLikelyForm10 =
        currentHeaders.some((h) => /number\s+in\s+register|register\s+number/i.test(String(h || ''))) &&
        currentHeaders.some((h) => /overtime/i.test(String(h || '')));

      const toNumber = (val) => {
        const n = Number(val);
        return Number.isFinite(n) ? n : 0;
      };

      const firstPresent = (...vals) => {
        for (let i = 0; i < vals.length; i++) {
          const v = vals[i];
          if (v === null || v === undefined) continue;
          if (typeof v === 'string' && v.trim() === '') continue;
          return v;
        }
        return '';
      };

      const getPayrollPayloadObject = (data) => {
        const parseMaybeJson = (v) => {
          if (!v) return {};
          if (typeof v === 'object') return v;
          if (typeof v === 'string') {
            try {
              const parsed = JSON.parse(v);
              return parsed && typeof parsed === 'object' ? parsed : {};
            } catch (_) {
              return {};
            }
          }
          return {};
        };

        if (!data) return {};
        if (Array.isArray(data)) {
          const first = data[0];
          if (!first) return {};
          if (typeof first === 'string') return parseMaybeJson(first);
          return getPayrollPayloadObject(first);
        }
        if (typeof data !== 'object') return {};
        if (data.salary) return getPayrollPayloadObject(parseMaybeJson(data.salary));
        if (data.employee_salary) return getPayrollPayloadObject(parseMaybeJson(data.employee_salary));
        if (data.data) return getPayrollPayloadObject(parseMaybeJson(data.data));
        if (data.employee && typeof data.employee === 'object') {
          return { ...parseMaybeJson(data.employee), ...data };
        }
        return data;
      };

      const getPayrollLineAmount = (lineItem) => {
        if (!lineItem || typeof lineItem !== 'object') return NaN;
        return toNumber(
          lineItem.amount ??
            lineItem.value ??
            lineItem.earning_amount ??
            lineItem.monthly_amount ??
            lineItem.Amount
        );
      };

      const getEarningsArray = (payrollObj) => {
        const p = payrollObj && typeof payrollObj === 'object' ? payrollObj : {};
        const candidates = [p.earnings, p.earning, p.employee_earnings, p.salary_components];
        for (const c of candidates) {
          if (Array.isArray(c)) return c;
          if (typeof c === 'string') {
            try {
              const parsed = JSON.parse(c);
              if (Array.isArray(parsed)) return parsed;
            } catch (_) {
              /* ignore */
            }
          }
        }
        return [];
      };

      const findEarningAmount = (earnings, matcher) => {
        const hit = earnings.find((it) =>
          matcher(String(it?.type || '').toLowerCase(), String(it?.name || '').toLowerCase())
        );
        if (!hit) return '';
        const amount = getPayrollLineAmount(hit);
        return Number.isFinite(amount) ? amount : '';
      };

      const pickFormBBasicEarningsAmount = (earnings) => {
        const list = Array.isArray(earnings) ? earnings : [];
        const exact = list.find((it) => String(it?.name || '').trim().toLowerCase() === 'basic earnings');
        if (exact) {
          const a = getPayrollLineAmount(exact);
          if (Number.isFinite(a)) return a;
        }
        return findEarningAmount(list, (type, name) => name.includes('basic earnings') || type === 'basic');
      };

      const pickFormBOtherAllowanceAmount = (earnings) => {
        const list = Array.isArray(earnings) ? earnings : [];
        const exact = list.find((it) => String(it?.name || '').trim().toLowerCase() === 'other allowance');
        if (exact) {
          const a = getPayrollLineAmount(exact);
          if (Number.isFinite(a)) return a;
        }
        return findEarningAmount(list, (type, name) => name.includes('other allowance'));
      };

      const pickFormWHouseRentAllowanceAmount = (earnings) => {
        const list = Array.isArray(earnings) ? earnings : [];
        const exact = list.find((it) => String(it?.name || '').trim().toLowerCase() === 'house rent allowance');
        if (exact) {
          const a = getPayrollLineAmount(exact);
          if (Number.isFinite(a)) return a;
        }
        return findEarningAmount(
          list,
          (type, name) => type === 'hra' || name.includes('house rent') || name.includes('hra')
        );
      };

      const getDeductionsArray = (payrollObj) => {
        const p = payrollObj && typeof payrollObj === 'object' ? payrollObj : {};
        const candidates = [p.deductions, p.deduction, p.employee_deductions];
        for (const c of candidates) {
          if (Array.isArray(c)) return c;
          if (typeof c === 'string') {
            try {
              const parsed = JSON.parse(c);
              if (Array.isArray(parsed)) return parsed;
            } catch (_) {
              /* ignore */
            }
          }
        }
        return [];
      };

      const sumPayrollLineItems = (items) => {
        const list = Array.isArray(items) ? items : [];
        let total = 0;
        let hasAny = false;
        list.forEach((it) => {
          const amount = getPayrollLineAmount(it);
          if (Number.isFinite(amount)) {
            total += amount;
            hasAny = true;
          }
        });
        return hasAny ? total : '';
      };

      const normalizePayrollLookupValue = (value) =>
        String(value || '')
          .trim()
          .toLowerCase()
          .replace(/\s+/g, ' ');

      const buildPayrollLookupFromRows = (payrollRows) => {
        const byPayrollPayload = new Map();
        if (!Array.isArray(payrollRows)) return { byPayrollPayload };

        payrollRows.forEach((it) => {
          if (!it || typeof it !== 'object' || it.fetch_error) return;
          const payload = getPayrollPayloadObject(it);
          const payrollEmployeeId = String(
            it.employee_id || payload.employee_id || payload.employeeId || ''
          ).trim();
          const email = normalizePayrollLookupValue(
            it.work_mail || it.email || payload.work_mail || payload.email || payload.work_email || ''
          );
          const firstName = normalizePayrollLookupValue(
            it.first_name || payload.first_name || payload.firstName || ''
          );
          const nameCandidates = [
            `${it.first_name || ''} ${it.last_name || ''}`,
            `${payload.first_name || ''} ${payload.last_name || ''}`,
            it.employee_name,
            payload.employee_name,
            payload.employeeName,
            payload.name,
          ]
            .map(normalizePayrollLookupValue)
            .filter(Boolean);

          const register = (key) => {
            if (!key || byPayrollPayload.has(key)) return;
            byPayrollPayload.set(key, it);
          };

          if (payrollEmployeeId) register(payrollEmployeeId);
          if (email) register(`email:${email}`);
          nameCandidates.forEach((name) => register(`name:${name}`));
          if (firstName) register(`first:${firstName}`);
        });

        return { byPayrollPayload };
      };

      const resolvePayrollRowForEmployee = (byPayrollPayload, emp, row, rowNameCandidates = []) => {
        const payrollEmpIds = getEmployeeLookupIdCandidates(emp, row);
        const empEmail = normalizePayrollLookupValue(
          (emp && (emp.EmailID || emp.Email || emp.email || emp['Email ID'])) || ''
        );
        const empFirstName = normalizePayrollLookupValue(
          (emp && (emp.FirstName || emp['FirstName'] || '')) || ''
        );
        const empNameCandidates = [
          ...rowNameCandidates,
          ...getRowPayrollLookupNameCandidates(row),
          ...getEmployeeNameCandidates(emp || {}),
        ]
          .map(normalizePayrollLookupValue)
          .filter(Boolean);

        return (
          payrollEmpIds.map((id) => byPayrollPayload.get(id)).find(Boolean) ||
          (empEmail && byPayrollPayload.get(`email:${empEmail}`)) ||
          empNameCandidates.map((name) => byPayrollPayload.get(`name:${name}`)).find(Boolean) ||
          (empFirstName && byPayrollPayload.get(`first:${empFirstName}`)) ||
          null
        );
      };

      const resolveFormWTableHeaders = (headers) => ({
        basic: headers.find((h) => {
          const t = String(h || '').toLowerCase();
          return t.includes('basic') && (t.includes('wage') || t.includes('wag'));
        }),
        dearness: headers.find((h) => String(h || '').toLowerCase().includes('dearness')),
        houseRent: headers.find((h) => {
          const t = String(h || '').toLowerCase();
          return t.includes('house') && t.includes('rent');
        }),
        otherAllowance: headers.find((h) => {
          const t = String(h || '').toLowerCase();
          return t.includes('other') && t.includes('allowance');
        }),
        overtime: headers.find((h) => String(h || '').toLowerCase().includes('overtime')),
        gross: headers.find((h) => {
          const t = String(h || '').toLowerCase();
          return t.includes('gross') && (t.includes('wage') || t.includes('wag'));
        }),
        net: headers.find((h) => {
          const t = String(h || '').toLowerCase();
          return t.includes('net') && (t.includes('wage') || t.includes('wag'));
        }),
        totalDeductions: headers.find((h) => {
          const t = String(h || '').toLowerCase();
          return t.includes('total') && t.includes('deduction');
        }),
        daysWorked: headers.find((h) => {
          const hl = String(h || '').toLowerCase().replace(/\s+/g, ' ');
          return (
            (hl.includes('number') && hl.includes('days') && hl.includes('worked')) ||
            (hl.includes('no') && hl.includes('days') && hl.includes('worked')) ||
            (hl.includes('days') && hl.includes('worked'))
          );
        }),
        employeeName: headers.find((h) => {
          const t = String(h || '').toLowerCase();
          return (
            t.includes('name') &&
            (t.includes('employee') || t.includes('worker') || t.includes('emp'))
          );
        }),
        employeeId: headers.find((h) => {
          const t = String(h || '').toLowerCase();
          return (
            (t.includes('identification') ||
              (t.includes('employee') && t.includes('id')) ||
              (t.includes('worker') && t.includes('id'))) &&
            !t.includes('zoho')
          );
        }),
      });

      const applyFormWPayrollToRow = (row, formWPayrollMap, formWHeaders, { overwrite = false } = {}) => {
        if (!row || !formWPayrollMap || !formWHeaders) return;
        const setCell = (header, value) => {
          if (!header || value === '' || value == null) return;
          if (!overwrite && String(row[header] || '').trim()) return;
          row[header] = sanitizeValue(value);
        };
        setCell(formWHeaders.basic, formWPayrollMap.basicWage);
        setCell(formWHeaders.dearness, formWPayrollMap.dearnessAllowance);
        setCell(formWHeaders.houseRent, formWPayrollMap.houseRentAllowance);
        setCell(formWHeaders.otherAllowance, formWPayrollMap.otherAllowances);
        setCell(formWHeaders.overtime, formWPayrollMap.overtimeWages);
        setCell(formWHeaders.gross, formWPayrollMap.grossWages);
        setCell(formWHeaders.net, formWPayrollMap.netWages);
        setCell(formWHeaders.totalDeductions, formWPayrollMap.totalDeductions);
      };

      const pickForm25FieldValue = (record, keys) => {
        if (!record || typeof record !== 'object' || !Array.isArray(keys)) return '';
        for (let i = 0; i < keys.length; i += 1) {
          const key = keys[i];
          const direct = record[key];
          if (direct != null && String(direct).trim() !== '') return String(direct);
          const normalized = findValueByNormalizedKey(record, key);
          if (normalized != null && String(normalized).trim() !== '') return String(normalized);
        }
        return '';
      };

      const buildFormWPayrollMap = (payrollData) => {
        const p = getPayrollPayloadObject(payrollData);
        const earnings = getEarningsArray(p);
        const deductions = getDeductionsArray(p);

        const basic = pickFormBBasicEarningsAmount(earnings);
        const hra = pickFormWHouseRentAllowanceAmount(earnings);
        const dearness = findEarningAmount(earnings, (type, name) => type === 'da' || name.includes('dearness'));
        const overtime = findEarningAmount(earnings, (type, name) => type === 'overtime' || type === 'ot' || name.includes('overtime'));
        const monthlyGross = toNumber(p.monthly_gross_amount);
        const monthlySalary = toNumber(p.monthly_salary);
        const otherAllowance = pickFormBOtherAllowanceAmount(earnings);

        let totalDeductions = sumPayrollLineItems(deductions);
        if (totalDeductions === '' && monthlyGross > 0 && monthlySalary >= 0 && monthlyGross >= monthlySalary) {
          totalDeductions = monthlyGross - monthlySalary;
        }
        const totalDeductionsRaw = firstPresent(
          p.total_deductions,
          p['total_deductions'],
          p.totalDeductions,
          p['totalDeductions'],
          p.monthly_total_deductions,
          totalDeductions !== '' ? totalDeductions : ''
        );

        return {
          basicWage: basic,
          dearnessAllowance: dearness,
          houseRentAllowance: hra,
          otherAllowances: otherAllowance,
          overtimeWages: overtime,
          grossWages: Number.isFinite(monthlyGross) ? monthlyGross : '',
          netWages: Number.isFinite(monthlySalary) ? monthlySalary : '',
          totalDeductions: totalDeductionsRaw !== '' ? totalDeductionsRaw : '',
        };
      };

      const parseTimeToMinutes = (raw) => {
        const value = String(raw || '').trim();
        if (!value) return null;
        const m12 = value.match(/^(\d{1,2}):(\d{2})\s*(am|pm)$/i);
        if (m12) {
          let hh = Number(m12[1]);
          const mm = Number(m12[2]);
          const meridiem = m12[3].toLowerCase();
          if (!Number.isFinite(hh) || !Number.isFinite(mm)) return null;
          if (meridiem === 'pm' && hh < 12) hh += 12;
          if (meridiem === 'am' && hh === 12) hh = 0;
          return (hh * 60) + mm;
        }
        const m24 = value.match(/^(\d{1,2}):(\d{2})$/);
        if (m24) {
          const hh = Number(m24[1]);
          const mm = Number(m24[2]);
          if (!Number.isFinite(hh) || !Number.isFinite(mm)) return null;
          return (hh * 60) + mm;
        }
        return null;
      };

      const parseFlexibleDate = (raw) => {
        const value = String(raw || '').trim();
        if (!value) return null;

        const direct = new Date(value);
        if (!Number.isNaN(direct.getTime())) return direct;

        let m = value.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})$/);
        if (m) {
          const dd = Number(m[1]);
          const mm = Number(m[2]) - 1;
          const yyyy = Number(m[3].length === 2 ? `20${m[3]}` : m[3]);
          const d = new Date(yyyy, mm, dd);
          if (!Number.isNaN(d.getTime())) return d;
        }

        m = value.match(/^(\d{1,2})[-/\s]([A-Za-z]{3,})[-/\s](\d{2,4})$/);
        if (m) {
          const dd = Number(m[1]);
          const mon = String(m[2]).slice(0, 3).toLowerCase();
          const yyyy = Number(m[3].length === 2 ? `20${m[3]}` : m[3]);
          const months = {
            jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5,
            jul: 6, aug: 7, sep: 8, oct: 9, nov: 10, dec: 11
          };
          if (Object.prototype.hasOwnProperty.call(months, mon)) {
            const d = new Date(yyyy, months[mon], dd);
            if (!Number.isNaN(d.getTime())) return d;
          }
        }
        return null;
      };

      const computeAgeFromDobForStatutory = (emp) => {
        const dobRaw =
          emp.Date_of_birth ||
          emp['Date_of_birth'] ||
          emp.DateofBirth ||
          emp['Date of Birth'] ||
          emp['DateofBirth'] ||
          '';
        const dob = parseFlexibleDate(dobRaw);
        if (dob && !Number.isNaN(dob.getTime())) {
          const today = new Date();
          let age = today.getFullYear() - dob.getFullYear();
          const md = today.getMonth() - dob.getMonth();
          if (md < 0 || (md === 0 && today.getDate() < dob.getDate())) age -= 1;
          if (age >= 0 && age < 120) return String(age);
        }
        const direct =
          emp.Age ||
          emp['Age'] ||
          emp.age ||
          emp['Employee Age'] ||
          '';
        return direct !== '' && direct != null ? String(direct).trim() : '';
      };

      const formatLocalDateYYYYMMDD = (dateObj) => {
        if (!(dateObj instanceof Date) || Number.isNaN(dateObj.getTime())) return '';
        const y = dateObj.getFullYear();
        const m = String(dateObj.getMonth() + 1).padStart(2, '0');
        const d = String(dateObj.getDate()).padStart(2, '0');
        return `${y}-${m}-${d}`;
      };

      const findValueByNormalizedKey = (obj, targetNormalizedKey, depth = 0) => {
        if (!obj || typeof obj !== 'object' || depth > 5) return '';
        const entries = Object.entries(obj);
        for (let i = 0; i < entries.length; i++) {
          const [k, v] = entries[i];
          const nk = String(k || '').toLowerCase().replace(/[^a-z0-9]/g, '');
          if (nk === targetNormalizedKey && v != null && String(v).trim() !== '') {
            return v;
          }
        }
        for (let i = 0; i < entries.length; i++) {
          const [, v] = entries[i];
          if (v && typeof v === 'object') {
            const nested = findValueByNormalizedKey(v, targetNormalizedKey, depth + 1);
            if (nested !== '') return nested;
          }
        }
        return '';
      };

      const computeNormalHoursFromShift = (emp) => {
        const shiftStart =
          emp.ShiftStartTime ||
          emp['ShiftStartTime'] ||
          emp.shiftStartTime ||
          emp['Shift Start Time'] ||
          '';
        const shiftEnd =
          emp.ShiftEndTime ||
          emp['ShiftEndTime'] ||
          emp.shiftEndTime ||
          emp['Shift End Time'] ||
          '';
        const startMin = parseTimeToMinutes(shiftStart);
        const endMin = parseTimeToMinutes(shiftEnd);
        if (startMin == null || endMin == null) return '';
        let diff = endMin - startMin;
        if (diff < 0) diff += 24 * 60;
        const hours = Math.floor(diff / 60);
        const mins = diff % 60;
        return `${String(hours).padStart(2, '0')}:${String(mins).padStart(2, '0')}`;
      };

      const computeNormalHoursFromAttendanceRecord = (record) => {
        if (!record || typeof record !== 'object') return '';
        const shiftStart =
          record.ShiftStartTime ||
          record['ShiftStartTime'] ||
          record.shiftStartTime ||
          record['Shift Start Time'] ||
          findValueByNormalizedKey(record, 'shiftstarttime') ||
          '';
        const shiftEnd =
          record.ShiftEndTime ||
          record['ShiftEndTime'] ||
          record.shiftEndTime ||
          record['Shift End Time'] ||
          findValueByNormalizedKey(record, 'shiftendtime') ||
          '';
        const startMin = parseTimeToMinutes(shiftStart);
        const endMin = parseTimeToMinutes(shiftEnd);
        if (startMin == null || endMin == null) return '';
        let diff = endMin - startMin;
        if (diff < 0) diff += 24 * 60;
        const hours = Math.floor(diff / 60);
        const mins = diff % 60;
        return `${String(hours).padStart(2, '0')}:${String(mins).padStart(2, '0')}`;
      };

      const getOvertimeWorkedValue = (emp) => {
        const raw =
          emp.OverTime ||
          emp['OverTime'] ||
          emp.overtime ||
          emp['overtime'] ||
          emp.OTHours ||
          emp['OTHours'] ||
          emp.OvertimeHours ||
          emp['OvertimeHours'] ||
          emp['Total overtime worked or production in case of piece workers'] ||
          '';
        const n = Number(raw);
        if (Number.isFinite(n) && n > 0) return n;
        const s = String(raw || '').trim();
        return s || '';
      };

      const getOvertimeDateValue = (emp, hasOvertime) => {
        if (!hasOvertime) return '';
        return (
          emp.OvertimeDate ||
          emp['OvertimeDate'] ||
          emp.OverTimeDate ||
          emp['OverTimeDate'] ||
          emp.overtime_date ||
          emp['overtime_date'] ||
          emp.Date ||
          emp['Date'] ||
          new Date().toISOString().slice(0, 10)
        );
      };

      const buildForm10PayrollMap = (payrollData) => {
        const p = getPayrollPayloadObject(payrollData);
        const earnings = getEarningsArray(p);
        const overtimeAmount = findEarningAmount(
          earnings,
          (type, name) => type === 'overtime' || type === 'ot' || name.includes('overtime')
        );
        const monthlySalaryRaw = firstPresent(
          p.monthly_salary,
          p['monthly_salary'],
          p.MonthlySalary,
          p['MonthlySalary']
        );
        const totalEarnings = monthlySalaryRaw !== ''
          ? monthlySalaryRaw
          : (
            toNumber(p.monthly_gross_amount) ||
            toNumber(p.gross_amount) ||
            ''
          );
        return {
          totalEarnings,
          normalEarnings: totalEarnings,
          totalOvertimeWorked: overtimeAmount
        };
      };

      const pickFormDBasicEarningsAmount = (earnings) => {
        const list = Array.isArray(earnings) ? earnings : [];
        const exact = list.find((it) => String(it?.name || '').trim().toLowerCase() === 'basic earnings');
        if (exact) {
          const a = getPayrollLineAmount(exact);
          if (Number.isFinite(a)) return a;
        }
        return findEarningAmount(list, (type, name) => name.includes('basic earnings') || type === 'basic' || name.includes('basic'));
      };

      const buildFormDPayrollMap = (payrollData) => {
        const p = getPayrollPayloadObject(payrollData);
        const earnings = getEarningsArray(p);
        return {
          basicSalary: firstPresent(
            pickFormDBasicEarningsAmount(earnings),
            p.monthly_salary,
            p['monthly_salary'],
            p.MonthlySalary,
            p['MonthlySalary']
          ),
          houseRentAllowance: findEarningAmount(
            earnings,
            (type, name) => type === 'hra' || name.includes('house rent') || name.includes('hra')
          ),
          dearnessAllowance: findEarningAmount(
            earnings,
            (type, name) => type === 'da' || name.includes('dearness') || name.includes('dearness allowance')
          ),
          otherAllowance: findEarningAmount(
            earnings,
            (type, name) => name.includes('other allowance') || (name.includes('other') && name.includes('allowance'))
          )
        };
      };

      const getPayrollEmployeeId = (emp) =>
        emp.Zoho_ID ||
        emp['Zoho_ID'] ||
        emp.zoho_id ||
        emp['zoho_id'] ||
        emp.ZohoID ||
        emp['ZohoID'] ||
        emp.zohoId ||
        emp.EmployeeID ||
        emp['EmployeeID'] ||
        emp.employeeId ||
        emp['Employee ID'] ||
        '';

      const isFormWDaysWorkedHeader = (header) => {
        const h = String(header || '').toLowerCase().replace(/\s+/g, ' ');
        return (
          (h.includes('number') && h.includes('days') && h.includes('worked')) ||
          (h.includes('no') && h.includes('days') && h.includes('worked')) ||
          (h.includes('days') && h.includes('worked'))
        );
      };

      const isFormWWageOrDeductionHeader = (header) => {
        const h = String(header || '').toLowerCase();
        if (isFormWDaysWorkedHeader(header)) return false;
        return (
          (h.includes('basic') && (h.includes('wage') || h.includes('wag'))) ||
          h.includes('dearness') ||
          (h.includes('house') && h.includes('rent')) ||
          (h.includes('other') && h.includes('allowance')) ||
          h.includes('overtime') ||
          (h.includes('gross') && (h.includes('wage') || h.includes('wag'))) ||
          h.includes('provident') ||
          h.includes('insurance') ||
          (h.includes('labour') && h.includes('welfare')) ||
          h.includes('advance') ||
          h.includes('damage') ||
          h.includes('fine') ||
          (h.includes('total') && h.includes('deduction')) ||
          (h.includes('deduction') && !h.includes('nature')) ||
          (h.includes('net') && h.includes('wage')) ||
          h.includes('unpaid') ||
          h.includes('accumulation')
        );
      };

      let formWPayrollLookup = null;
      let formWHeadersResolved = null;
      if (isLikelyFormW) {
        formWHeadersResolved = resolveFormWTableHeaders(currentHeaders);
        try {
          const payrollOrgId =
            process.env.REACT_APP_ZOHO_PAYROLL_ORGANIZATION_ID || '60006183023';
          const allPayrollQs = new URLSearchParams({
            all_salaries: '1',
            organization_id: payrollOrgId,
          });
          const allPayrollRes = await fetch(`/server/payroll_function?${allPayrollQs.toString()}`);
          const allPayrollJson = await allPayrollRes.json();
          if (allPayrollRes.ok && allPayrollJson?.success && Array.isArray(allPayrollJson.data)) {
            formWPayrollLookup = buildPayrollLookupFromRows(allPayrollJson.data);
            console.log(
              `Form W: indexed ${allPayrollJson.data.length} payroll row(s) for wage autofill`
            );
          } else {
            console.warn('Form W payroll all_salaries preload failed:', allPayrollJson);
          }
        } catch (formWPayPreloadErr) {
          console.warn('Form W payroll preload skipped:', formWPayPreloadErr?.message || formWPayPreloadErr);
        }
      }

      const mappedData = await Promise.all(employees.map(async (empItem, index) => {
        // Initialize all headers with empty strings to ensure no undefined/null values
        const row = {};
        currentHeaders.forEach((header) => {
          row[header] = '';
        });
       
        // Extract employee object - handle both wrapped and direct structures
        const emp = empItem.Employee || empItem.employee || empItem;
        row.__employeeLookupName = getEmployeeLookupName(emp);
        row.__employeeLookupId = firstPresent(
          getPayrollEmployeeId(emp),
          getFallbackEmployeeId(emp),
          getFallbackWorkerId(emp)
        );
       
        console.log(`Processing employee ${index + 1}:`, emp);
        console.log(`Available fields in employee ${index + 1}:`, Object.keys(emp));
       
        // Map to table headers if they exist
        currentHeaders.forEach((header, colIndex) => {
          const headerLower = header.toLowerCase().trim();
          const normalizedHeader = headerLower.replace(/[^a-z0-9]/g, '');
          if (shouldUseGenericColumnFallback) {
            const fallbackValue = getFallbackByColumnIndex(emp, colIndex, index);
            row[header] = fallbackValue;
            return;
          }
         
          // For Form W, wage/deduction and days-worked columns come from payroll / Form 25 — not People.
          if (isLikelyFormW && (isFormWWageOrDeductionHeader(header) || isFormWDaysWorkedHeader(header))) {
            row[header] = '';
            return;
          }

          // Day-of-month grid (Form 25 / Form V): never fuzzy-map People fields; attendance fills these later (if any).
          if (isLikelyDayOfMonthColumnHeader(header)) {
            row[header] = '';
            return;
          }

          if (
            form25SkipLossPayAndNationalHolidayBenefit &&
            (isForm25LossOfPayDaysHeader(header) || isForm25NationalHolidayBenefitHeader(header))
          ) {
            row[header] = '';
            return;
          }

          if (formVIFestivalAutofillContext && isFormVIFestivalManualHeader(header)) {
            row[header] = '';
            return;
          }

          if (
            formAAutofillContext &&
            (isFormAMonthOnlyColumn(header) ||
              isFormADaysEmployedColumn(header) ||
              isFormALaidOffColumn(header) ||
              isFormANotEmployedColumn(header))
          ) {
            row[header] = '';
            return;
          }

          if (formBAutofillContext && isFormBBalanceDueHeader(header)) {
            row[header] = '';
            return;
          }

          if (formBAutofillContext && isFormBTotalEmolumentsBasicEarningsHeader(header)) {
            row[header] = '';
            return;
          }

          if (formBAutofillContext && isFormBOtherDeductionsOtherAllowanceHeader(header)) {
            row[header] = '';
            return;
          }

          if (formDAutofillContext && isFormDRemunerationHeader(header)) {
            row[header] = '';
            return;
          }

          if (formDAutofillContext && headerLower.includes('category of workers')) {
            row[header] =
              emp.Designation ||
              emp.designation ||
              emp['Designation.displayValue'] ||
              findValueByNormalizedKey(emp, 'designation') ||
              '';
            return;
          }

          if (isFormDNoMenEmployedHeader(header)) {
            const gender = String(
              emp.Sex ||
              emp.Gender ||
              emp.gender ||
              emp['Sex'] ||
              emp['Gender.displayValue'] ||
              findValueByNormalizedKey(emp, 'gender') ||
              ''
            ).trim().toLowerCase();
            const currentValue = String(row[header] || '').trim();
            if (!currentValue || isPlaceholderStatutoryCellValue(currentValue)) {
              row[header] = gender === 'male' || gender === 'm' ? '1' : '';
            }
            return;
          }

          if (isFormDNoWomenEmployedHeader(header)) {
            const gender = String(
              emp.Sex ||
              emp.Gender ||
              emp.gender ||
              emp['Sex'] ||
              emp['Gender.displayValue'] ||
              findValueByNormalizedKey(emp, 'gender') ||
              ''
            ).trim().toLowerCase();
            const currentValue = String(row[header] || '').trim();
            if (!currentValue || isPlaceholderStatutoryCellValue(currentValue)) {
              row[header] = gender === 'female' || gender === 'f' ? '1' : '';
            }
            return;
          }

          if (isLeaveWagesPerSection9Or10Header(header)) {
            row[header] = '';
            return;
          }

          if (isFormXLeaveMetricHeader(header)) {
            row[header] = '';
            return;
          }

          // S.No mapping
          if (
            headerLower.includes('serial') ||
            headerLower.includes('s.no') ||
            normalizedHeader === 'sno' ||
            normalizedHeader === 'serialno' ||
            normalizedHeader === 'serialnumber'
          ) {
            row[header] = String(index + 1);
          }
          // Form B (TN LWF Register of Wages): total employees = count with Employee ID + First + Last name in Zoho.
          else if (formBAutofillContext && isFormBTotalEmployeeCountHeader(header)) {
            row[header] = String(formBTotalEmployeeCount);
          }
          // Form A: "(1) Name of the woman and age" — single column; append computed age after name.
          else if (formAAutofillContext && isFormACombinedWomanNameAndAgeHeader(header)) {
            const namePart =
              emp.FirstName ||
              emp['FirstName'] ||
              emp.firstName ||
              emp['First Name'] ||
              emp.First_Name ||
              emp.Name1 ||
              emp['Name1'] ||
              emp.Nameoftheemployee ||
              emp['Name of the employee'] ||
              emp['Nameoftheemployee'] ||
              emp.name ||
              emp.employeeName ||
              emp.Name ||
              emp['Employee Name'] ||
              emp.Full_Name ||
              emp['Full Name'] ||
              emp.Employee_Name ||
              emp['Employee_Name'] ||
              '';
            const agePart = computeAgeFromDobForStatutory(emp);
            const nm = namePart != null ? String(namePart).trim() : '';
            row[header] = sanitizeValue(nm && agePart ? `${nm}, ${agePart}` : nm || agePart || '');
          }
          // Name of the employee/worker - try multiple field name variations
          else if (
            headerLower.includes('name') &&
            (headerLower.includes('employee') ||
              headerLower.includes('emp') ||
              headerLower.includes('worker') ||
              headerLower.includes('woman')) &&
            !(headerLower.includes('woman') && /\bage\b/.test(headerLower))
          ) {
            // Prioritize FirstName as it's the actual field in Zoho People
            const nameValue = emp.FirstName ||
                         emp['FirstName'] ||
                         emp.firstName ||
                         emp['First Name'] ||
                         emp.First_Name ||
                         emp.Name1 ||
                         emp['Name1'] ||
                         emp.Nameoftheemployee ||
                         emp['Name of the employee'] ||
                         emp['Nameoftheemployee'] ||
                         emp.name ||
                         emp.employeeName ||
                         emp.Name ||
                         emp['Employee Name'] ||
                         emp.Full_Name ||
                         emp['Full Name'] ||
                         emp.Employee_Name ||
                         emp['Employee_Name'] ||
                         emp['Employee Name'] ||
                         // Try to find any field with "name" in it
                         (() => {
                           // First try FirstName variations
                           const firstNameFields = Object.keys(emp).filter(key =>
                             /^FirstName$/i.test(key) || /^First_Name$/i.test(key) || /^First Name$/i.test(key)
                           );
                           
                           if (firstNameFields.length > 0) {
                             const value = emp[firstNameFields[0]];
                             if (value && value !== '') {
                               return value;
                             }
                           }
                           
                           // Then try numbered name fields (Name1, Name2, etc.)
                           const numberedNameFields = Object.keys(emp).filter(key =>
                             /^Name\d+$/i.test(key)
                           ).sort(); // Sort to get Name1 first
                           
                           if (numberedNameFields.length > 0) {
                             const value = emp[numberedNameFields[0]];
                             if (value && value !== '') {
                               return value;
                             }
                           }
                           
                           // Then try other name fields
                           const nameFields = Object.keys(emp).filter(key =>
                             key.toLowerCase().includes('name') &&
                             !key.toLowerCase().includes('father') &&
                             !key.toLowerCase().includes('spouse') &&
                             !key.toLowerCase().includes('emergency') &&
                             !key.toLowerCase().includes('dependent')
                           );
                           if (nameFields.length > 0) {
                             // Try the first matching field
                             const value = emp[nameFields[0]];
                             return value && value !== '' ? value : '';
                           }
                           return '';
                         })() ||
                         '';
           
            let displayName = nameValue;
            if (
              formBAutofillContext &&
              (headerLower.includes('employee') || headerLower.includes('emp')) &&
              !headerLower.includes('identification') &&
              !headerLower.includes('id number') &&
              !(headerLower.includes('emp') && headerLower.includes('id') && !headerLower.includes('name'))
            ) {
              const full = getFormBEmployeeFullName(emp);
              if (full) displayName = full;
            }
            row[header] = displayName;
            if (displayName) {
              const sourceField = Object.keys(emp).find(k => emp[k] === displayName) || 'unknown';
              console.log(`✓ Mapped name for employee ${index + 1}: ${displayName} (from field: ${sourceField})`);
            } else {
              console.warn(`✗ Could not find name field for employee ${index + 1}. Available name fields:`, Object.keys(emp).filter(k => k.toLowerCase().includes('name')));
            }
          }
          // Form A: Age from date of birth (or explicit Age field)
          else if (
            formAAutofillContext &&
            (normalizedHeader === 'age' || /^age\b/i.test(headerLower)) &&
            !headerLower.includes('wage') &&
            !headerLower.includes('average') &&
            !headerLower.includes('image') &&
            !headerLower.includes('language')
          ) {
            row[header] = sanitizeValue(computeAgeFromDobForStatutory(emp));
          }
          // Employee Identification No. - fetch from EmployeeID from People API (e.g. "EmployeeID":"10193")
          else if (headerLower.includes('identification') ||
                   (headerLower.includes('employee') && headerLower.includes('id')) ||
                   (headerLower.includes('emp') && headerLower.includes('code'))) {
            const idValue = emp.EmployeeID ||
                         emp['EmployeeID'] ||
                         emp['Role.ID'] ||
                         emp.Role?.ID ||
                         (emp.Role && typeof emp.Role === 'object' && emp.Role.ID ? emp.Role.ID : null) ||
                         emp.EmployeeIdentificationNo ||
                         emp['Employee Identification No.'] ||
                         emp['EmployeeIdentificationNo'] ||
                         emp.employeeCode ||
                         emp.employeeId ||
                         emp.EmployeeCode ||
                         emp['Employee Code'] ||
                         emp.Employee_ID ||
                         emp['Employee ID'] ||
                         emp.ID ||
                         // Try to find Role.ID in nested structures
                         (() => {
                           // Check if Role.ID exists directly
                           if (emp['Role.ID']) {
                             return emp['Role.ID'];
                           }
                           // Check if Role is an object with ID property
                           if (emp.Role && typeof emp.Role === 'object' && emp.Role.ID) {
                             return emp.Role.ID;
                           }
                           // Try to find any field with "code" or "id" in it (excluding specific fields)
                           const idFields = Object.keys(emp).filter(key => {
                             const keyLower = key.toLowerCase();
                             return (keyLower.includes('code') || keyLower.includes('id') || keyLower.includes('number')) &&
                                    !keyLower.includes('phone') &&
                                    !keyLower.includes('mobile') &&
                                    !keyLower.includes('aadhaar') &&
                                    !keyLower.includes('pf') &&
                                    !keyLower.includes('esic') &&
                                    !keyLower.includes('grade') &&
                                    !keyLower.includes('role.id'); // Already checked above
                           });
                           if (idFields.length > 0) {
                             const value = emp[idFields[0]];
                             return value && value !== '' ? value : '';
                           }
                           return '';
                         })() ||
                         '';
           
            row[header] = idValue;
            if (idValue) {
              console.log(`✓ Mapped Employee Identification No. for employee ${index + 1}: ${idValue}`);
            }
          }
          // Worker Identity No. - fetch from EmployeeID
          else if ((headerLower.includes('worker') && headerLower.includes('identity')) ||
                   headerLower === 'worker identity no.' ||
                   headerLower === 'worker identity no') {
            // Form 12 rule: Worker Identity No <- EmployeeID
            const workerIdValue = emp.EmployeeID ||
                         emp['EmployeeID'] ||
                         emp['Employee ID'] ||
                         emp.employeeId ||
                         emp.Employee_ID ||
                         emp['Role.ID'] ||
                         '';
           
            row[header] = workerIdValue;
            if (workerIdValue) {
              console.log(`✓ Mapped Worker Identity No. for employee ${index + 1}: ${workerIdValue} (from EmployeeID)`);
            } else {
              console.warn(`✗ Could not find EmployeeID field for employee ${index + 1}. Available ID fields:`,
                Object.keys(emp).filter(k => k.toLowerCase().includes('employee') || k.toLowerCase().includes('id')));
            }
          }
          // Emp ID - fetch from EmployeeID (fallback to Zoho_ID)
          else if (headerLower === 'emp id' ||
                   headerLower === 'empid' ||
                   (headerLower.includes('emp') && headerLower.includes('id') &&
                    !headerLower.includes('identification') &&
                    !headerLower.includes('code') &&
                    !headerLower.includes('number'))) {
            // Prioritize EmployeeID from Zoho People; fallback to Zoho_ID variants
            const empIdValue = emp.EmployeeID ||
                         emp['EmployeeID'] ||
                         emp['Role.ID'] ||
                         (emp.Role && typeof emp.Role === 'object' ? emp.Role.ID : '') ||
                         emp['Employee ID'] ||
                         emp.Zoho_ID ||
                         emp['Zoho_ID'] ||
                         emp.zoho_id ||
                         emp['zoho_id'] ||
                         emp.ZohoID ||
                         emp['ZohoID'] ||
                         emp.zohoId ||
                         '';
           
            row[header] = empIdValue;
            if (empIdValue) {
              console.log(`✓ Mapped Emp ID for employee ${index + 1}: ${empIdValue} (EmployeeID preferred)`);
            } else {
              console.warn(`✗ Could not find EmployeeID/Zoho_ID field for employee ${index + 1}. Available ID fields:`,
                Object.keys(emp).filter(k => k.toLowerCase().includes('zoho') || k.toLowerCase().includes('id')));
            }
          }
          // Gender
          else if (headerLower.includes('gender') || headerLower.includes('sex')) {
            row[header] = emp.Sex ||
                         emp.Gender ||
                         emp.gender ||
                         emp['Sex'] ||
                         '';
          }
          // Father/Spouse Name - prioritize Father_s_Name, fallback to Spouse_Name
          else if (headerLower.includes('father') || headerLower.includes('spouse')) {
            // First try to get Father_s_Name (or Father_Name variations)
            const fatherName = emp.Father_s_Name ||
                         emp['Father_s_Name'] ||
                         emp.Father_Name ||
                         emp['Father_Name'] ||
                         emp.FathersName ||
                         emp['Father/Spouse Name'] ||
                         emp['FathersName'] ||
                         emp.fathersName ||
                         emp['Father Name'] ||
                         '';
           
            // If father name exists, use it; otherwise try spouse name
            const spouseName = emp.Spouse_Name ||
                         emp['Spouse_Name'] ||
                         emp.spouseName ||
                         emp['Spouse Name'] ||
                         '';
           
            // Use father name if available, otherwise use spouse name
            const nameValue = (fatherName && fatherName.trim()) ? fatherName :
                            (spouseName && spouseName.trim()) ? spouseName : '';
           
            row[header] = nameValue;
            if (nameValue) {
              const sourceField = fatherName && fatherName.trim() ?
                (Object.keys(emp).find(k => emp[k] === fatherName) || 'Father_s_Name') :
                (Object.keys(emp).find(k => emp[k] === spouseName) || 'Spouse_Name');
              console.log(`✓ Mapped Father/Spouse Name for employee ${index + 1}: ${nameValue} (from field: ${sourceField})`);
            }
          }
          // Date of Birth
          else if (headerLower.includes('date') && headerLower.includes('birth')) {
            // Prioritize Date_of_birth as it's the actual field in Zoho People
            const birthDate = emp.Date_of_birth ||
                         emp['Date_of_birth'] ||
                         emp.DateofBirth ||
                         emp['Date of Birth'] ||
                         emp['DateofBirth'] ||
                         emp.dateOfBirth ||
                         emp['Date of birth'] ||
                         emp.Date_Of_Birth ||
                         emp['Date_Of_Birth'] ||
                         '';
           
            row[header] = birthDate;
            if (birthDate) {
              console.log(`✓ Mapped Date of Birth for employee ${index + 1}: ${birthDate}`);
            }
          }
          // Date of Entry into Service (Form 12) -> Dateofjoining
          else if (headerLower.includes('date') && headerLower.includes('entry') && headerLower.includes('service')) {
            const entryDate = emp.Dateofjoining ||
                         emp['Dateofjoining'] ||
                         emp.DateofJoining ||
                         emp['Date of Joining'] ||
                         emp['DateofJoining'] ||
                         emp.dateOfJoining ||
                         emp['Date of joining'] ||
                         emp.Date_of_Joining ||
                         emp['Date_of_Joining'] ||
                         '';

            row[header] = entryDate;
            if (entryDate) {
              console.log(`✓ Mapped Date of Entry into Service for employee ${index + 1}: ${entryDate} (from Dateofjoining)`);
            }
          }
          // Date of Joining
          else if (headerLower.includes('date') && headerLower.includes('join')) {
            // Prioritize Dateofjoining as it's the actual field in Zoho People
            const joiningDate = emp.Dateofjoining ||
                         emp['Dateofjoining'] ||
                         emp.DateofJoining ||
                         emp['Date of Joining'] ||
                         emp['DateofJoining'] ||
                         emp.dateOfJoining ||
                         emp['Date of joining'] ||
                         emp.Date_of_Joining ||
                         emp['Date_of_Joining'] ||
                         '';
           
            row[header] = joiningDate;
            if (joiningDate) {
              console.log(`✓ Mapped Date of Joining for employee ${index + 1}: ${joiningDate}`);
            }
          }
          // Date on which made permanent - fetch from DateofConfirmation
          else if (headerLower.includes('date') && headerLower.includes('permanent') && headerLower.includes('made')) {
            // Form 12 rule: Date on which made Permanent <- DateofConfirmation
            const permanentDate = emp.DateofConfirmation ||
                         emp['DateofConfirmation'] ||
                         emp['Date of Confirmation'] ||
                         emp.dateOfConfirmation ||
                         '';
           
            row[header] = permanentDate;
            if (permanentDate) {
              console.log(`✓ Mapped Date on which made permanent for employee ${index + 1}: ${permanentDate} (from DateofConfirmation)`);
            } else {
              console.warn(`✗ Could not find DateofConfirmation field for employee ${index + 1} to populate "Date on which made permanent"`);
            }
          }
          // Form 12: Date on which completion of 480 days of Service
          else if (headerLower.includes('480') && headerLower.includes('service')) {
            // Keep blank as requested.
            row[header] = '';
          }
          // Date of Exit - fetch from Dateofexit
          else if (headerLower.includes('date') && headerLower.includes('exit')) {
            row[header] = emp.Dateofexit ||
                         emp['Dateofexit'] ||
                         emp['Date of Exit'] ||
                         emp.dateOfExit ||
                         emp.Date_of_Exit ||
                         emp['Date_of_Exit'] ||
                         '';
          }
          // Form 10: Number in Register
          else if (headerLower.includes('number in register') || (headerLower.includes('register') && headerLower.includes('number'))) {
            row[header] =
              emp.EmployeeID ||
              emp['EmployeeID'] ||
              emp['Employee ID'] ||
              emp.Zoho_ID ||
              emp['Zoho_ID'] ||
              String(index + 1);
          }
          // Form 10: Date of which overtime has been worked
          else if (headerLower.includes('date of which overtime') || (headerLower.includes('date') && headerLower.includes('overtime'))) {
            const overtimeWorked = getOvertimeWorkedValue(emp);
            const hasOvertime = Number(overtimeWorked) > 0 || String(overtimeWorked || '').trim() !== '';
            row[header] = getOvertimeDateValue(emp, hasOvertime);
          }
          // Form 10: Total overtime worked or production in case of piece workers
          else if (
            (headerLower.includes('total') && headerLower.includes('overtime') && headerLower.includes('worked')) ||
            (headerLower.includes('overtime') && headerLower.includes('production'))
          ) {
            row[header] = getOvertimeWorkedValue(emp);
          }
          // Form 10: Normal hours from ShiftStartTime and ShiftEndTime
          else if (headerLower.includes('normal hours')) {
            row[header] =
              computeNormalHoursFromShift(emp) ||
              emp.NormalHours ||
              emp['NormalHours'] ||
              emp['Normal hours'] ||
              '';
          }
          // Form 10: Total earnings from monthly_salary
          else if (
            (headerLower.includes('total') && headerLower.includes('earning')) ||
            headerLower.includes('total earnings') ||
            (headerLower.includes('normal') && headerLower.includes('earning'))
          ) {
            row[header] =
              emp.monthly_salary ||
              emp['monthly_salary'] ||
              emp.MonthlySalary ||
              emp['MonthlySalary'] ||
              '';
          }
          // Designation
          else if (headerLower.includes('designation')) {
            row[header] = emp.Designation ||
                         emp.designation ||
                         '';
          }
          // Form I: Name and Address of the workman -> include employee name + address
          else if (headerLower.includes('name and address of the workman')) {
            const employeeName =
              emp.Name ||
              emp['Name'] ||
              emp.Name1 ||
              emp['Name1'] ||
              emp.EmployeeName ||
              emp['Employee Name'] ||
              `${emp.FirstName || emp['FirstName'] || ''} ${emp.LastName || emp['LastName'] || ''}`.trim() ||
              '';
            const addressLine1 = emp.Address_Line_1 || emp['Address_Line_1'] || emp.AddressLine1 || '';
            const addressLine2 = emp.Address_Line_2 || emp['Address_Line_2'] || emp.AddressLine2 || '';
            const city = emp.City || emp['City'] || emp.City1 || emp['City1'] || '';
            const country = emp.Country || emp['Country'] || emp.Country1 || emp['Country1'] || '';
            const addressParts = [addressLine1, addressLine2, city, country]
              .map((value) => String(value || '').trim())
              .filter(Boolean);
            const combinedAddress = addressParts.join(', ');
            row[header] = [String(employeeName || '').trim(), combinedAddress].filter(Boolean).join(', ');
          }
          // Present Address - combine Address_Line_1, Country, City1, and City
          else if (headerLower.includes('present') && headerLower.includes('address')) {
            // Get individual address components
            const addressLine1 = emp.Address_Line_1 || emp['Address_Line_1'] || emp.AddressLine1 || emp['Address Line 1'] || '';
            const country = emp.Country || emp['Country'] || '';
            const city1 = emp.City1 || emp['City1'] || '';
            const city = emp.City || emp['City'] || '';
           
            // Combine address components with proper formatting
            const addressParts = [];
            if (addressLine1 && addressLine1.trim()) {
              addressParts.push(addressLine1.trim());
            }
            // Add City1 if available
            if (city1 && city1.trim()) {
              addressParts.push(city1.trim());
            }
            // Add City if available (and different from City1)
            if (city && city.trim() && city.trim() !== city1.trim()) {
              addressParts.push(city.trim());
            }
            // If City1 is empty but City exists, use City
            if (!city1 && city && city.trim()) {
              addressParts.push(city.trim());
            }
            if (country && country.trim()) {
              addressParts.push(country.trim());
            }
           
            // Join with comma and space, or use fallback if no components found
            const combinedAddress = addressParts.length > 0
              ? addressParts.join(', ')
              : (emp.PresentAddress ||
                 emp['Present Address'] ||
                 emp['PresentAddress'] ||
                 emp.presentAddress ||
                 '');
           
            row[header] = combinedAddress;
            if (combinedAddress) {
              console.log(`✓ Mapped Present Address for employee ${index + 1}: ${combinedAddress}`);
            }
          }
          // Permanent address - combine Address_Line_1, Address_Line_2, Country1, City, and City2
          else if (headerLower.includes('permanent') && headerLower.includes('address')) {
            // Get individual address components - check all variations
            const addressLine1 = emp.Address_Line_1 ||
                               emp['Address_Line_1'] ||
                               emp.AddressLine1 ||
                               emp['Address Line 1'] ||
                               emp.Address_Line1 ||
                               '';
            const addressLine2 = emp.Address_Line_2 ||
                               emp['Address_Line_2'] ||
                               emp.AddressLine2 ||
                               emp['Address Line 2'] ||
                               emp['Address_Line_2'] || // Check with underscore variation
                               emp.Address_Line2 ||
                               '';
            const country1 = emp.Country1 || emp['Country1'] || '';
            const city = emp.City || emp['City'] || '';
            const city2 = emp.City2 || emp['City2'] || '';
           
            // Combine address components with proper formatting
            const addressParts = [];
            // Add Address_Line_1 first if available
            if (addressLine1 && addressLine1.trim()) {
              addressParts.push(addressLine1.trim());
            }
            // Add Address_Line_2 if available
            if (addressLine2 && addressLine2.trim()) {
              addressParts.push(addressLine2.trim());
            }
            // Add City if available
            if (city && city.trim()) {
              addressParts.push(city.trim());
            }
            // Add City2 if available (and different from City)
            if (city2 && city2.trim() && city2.trim() !== city.trim()) {
              addressParts.push(city2.trim());
            }
            // If City is empty but City2 exists, use City2
            if (!city && city2 && city2.trim()) {
              addressParts.push(city2.trim());
            }
            if (country1 && country1.trim()) {
              addressParts.push(country1.trim());
            }
           
            // Join with comma and space, or use fallback if no components found
            const combinedAddress = addressParts.length > 0
              ? addressParts.join(', ')
              : (emp.Permanentaddress ||
                 emp['Permanent address'] ||
                 emp['Permanentaddress'] ||
                 emp.permanentAddress ||
                 '');
           
            row[header] = combinedAddress;
            if (combinedAddress) {
              console.log(`✓ Mapped Permanent address for employee ${index + 1}: ${combinedAddress} (from: Address_Line_1=${addressLine1}, Address_Line_2=${addressLine2}, City=${city}, City2=${city2}, Country1=${country1})`);
            }
          }
          // EPF No / UAN No - fetch from UAN_Number
          else if (headerLower.includes('provident') || headerLower.includes('pf')) {
            const pfValue = emp.UAN_Number ||
                         emp['UAN_Number'] ||
                         emp.UANNumber ||
                         emp['UAN Number'] ||
                         emp.uanNumber ||
                         emp.UAN ||
                         emp['UAN'] ||
                         '';

            row[header] = pfValue && pfValue.toString().trim() ? pfValue.toString().trim() : '';

            if (row[header]) {
              console.log(`✓ Mapped EPF/UAN for employee ${index + 1}: ${row[header]} (from UAN_Number)`);
            } else {
              console.log(`ℹ EPF/UAN left blank for employee ${index + 1} (UAN_Number not available in Zoho People data)`);
            }
          }
          // ESIC No.
          else if (headerLower.includes('insurance') || headerLower.includes('esic')) {
            row[header] = emp['Employee\'s State Insurance Corporation No.'] ||
                         emp.ESICNo ||
                         emp.esicNo ||
                         '';
          }
          // Aadhaar No.
          else if (headerLower.includes('aadhaar') || headerLower.includes('aadhar')) {
            row[header] = emp.Aadhaar_Number ||
                         emp['Aadhaar_Number'] ||
                         emp['Aadhaar Number'] ||
                         emp.AadhaarNo ||
                         emp['Aadhaar No.'] ||
                         emp['AadhaarNo'] ||
                         emp.aadhaarNo ||
                         '';
          }
          // Email ID mapping
          else if (headerLower.includes('email')) {
            row[header] = emp.EmailID ||
                         emp['EmailID'] ||
                         emp.Email ||
                         emp.email ||
                         emp['Email ID'] ||
                         emp.emailId ||
                         '';
          }
          // Mobile Number - fetch from Mobile
          else if (headerLower.includes('mobile') || (headerLower.includes('phone') && !headerLower.includes('emergency'))) {
            // Prioritize Mobile as the source field
            const mobileValue = emp.Mobile ||
                         emp['Mobile'] ||
                         emp.mobile ||
                         emp.Mobile_Number ||
                         emp['Mobile_Number'] ||
                         emp['Mobile Number'] ||
                         emp.MobileNumber ||
                         emp['MobileNumber'] ||
                         emp.Phone ||
                         emp['Phone'] ||
                         emp.phone ||
                         emp.Phone_Number ||
                         emp['Phone_Number'] ||
                         emp['Phone Number'] ||
                         emp.PhoneNumber ||
                         emp['PhoneNumber'] ||
                         '';
           
            row[header] = mobileValue;
            if (mobileValue) {
              console.log(`✓ Mapped Mobile Number for employee ${index + 1}: ${mobileValue} (from Mobile)`);
            }
          }
          // Department mapping
          else if (headerLower.includes('department')) {
            row[header] = emp.Department ||
                         emp.department ||
                         '';
          }
          // Bank fields from People API: Bank_Name, Account_Number, IFSC_Code, Branch_Name
          // Combined Bank Header: "Bank A/c Number, Name of Bank, Branch (Indian Financial System Code) (IFSC Code)"
          else if ((headerLower.includes('bank') &&
                   (headerLower.includes('account') || headerLower.includes('ac')) &&
                   headerLower.includes('name') &&
                   headerLower.includes('branch')) ||
                   (headerLower.includes('bank') && headerLower.includes('ifsc')) ||
                   (headerLower.includes('indian financial system code') && headerLower.includes('ifsc'))) {
            // Fetch from People: Bank_Name, Account_Number, IFSC_Code, Branch_Name
            const accountNumber = emp.Account_Number ||
                         emp['Account_Number'] ||
                         emp.AccountNumber ||
                         emp['Account Number'] ||
                         emp.accountNumber ||
                         '';
           
            // Try multiple variations of Bank_Name field with fallback search
            const bankName = emp.Bank_Name ||
                         emp['Bank_Name'] ||
                         emp.BankName ||
                         emp['Bank Name'] ||
                         emp.bankName ||
                         emp['bankName'] ||
                         emp.BANK_NAME ||
                         emp['BANK_NAME'] ||
                         // Try to find any field with "bank" in the name (excluding account, branch, ifsc)
                         (() => {
                           const bankFields = Object.keys(emp).filter(key => {
                             const keyLower = key.toLowerCase();
                             return (keyLower.includes('bank') &&
                                    !keyLower.includes('account') &&
                                    !keyLower.includes('branch') &&
                                    !keyLower.includes('ifsc') &&
                                    !keyLower.includes('code'));
                           });
                           if (bankFields.length > 0) {
                             const value = emp[bankFields[0]];
                             return value && value.toString().trim() ? value.toString().trim() : '';
                           }
                           return '';
                         })() ||
                         '';
           
            const branchName = emp.Branch_Name ||
                         emp['Branch_Name'] ||
                         emp.BranchName ||
                         emp['Branch Name'] ||
                         emp.branchName ||
                         '';
           
            const ifscCode = emp.IFSC_Code ||
                         emp['IFSC_Code'] ||
                         emp.IFSCCode ||
                         emp['IFSC Code'] ||
                         emp.ifscCode ||
                         emp.IFSC ||
                         emp['IFSC'] ||
                         '';
           
            // Combine all bank information into one string
            // Format: AccountNumber, BankName, BranchName (IFSC Code)
            const bankParts = [];
            if (accountNumber && accountNumber.trim()) {
              bankParts.push(accountNumber.trim());
            }
            if (bankName && bankName.trim()) {
              bankParts.push(bankName.trim());
            }
            if (branchName && branchName.trim()) {
              bankParts.push(branchName.trim());
            }
           
            // Build the combined string
            let combinedBankInfo = bankParts.length > 0 ? bankParts.join(', ') : '';
           
            // Add IFSC code in parentheses if available
            if (ifscCode && ifscCode.trim()) {
              if (combinedBankInfo) {
                combinedBankInfo += ` (${ifscCode.trim()})`;
              } else {
                combinedBankInfo = ifscCode.trim();
              }
            }
           
            row[header] = combinedBankInfo;
           
            if (combinedBankInfo) {
              const bankNameSource = bankName ? Object.keys(emp).find(k => emp[k] === bankName) || 'Bank_Name' : 'not found';
              console.log(`✓ Mapped combined bank header for employee ${index + 1}: ${combinedBankInfo}`);
              console.log(`  - Account_Number: ${accountNumber || 'not found'}`);
              console.log(`  - Bank_Name: ${bankName || 'not found'} (from field: ${bankNameSource})`);
              console.log(`  - Branch_Name: ${branchName || 'not found'}`);
              console.log(`  - IFSC_Code: ${ifscCode || 'not found'}`);
            } else {
              // Log available fields for debugging
              const availableBankFields = Object.keys(emp).filter(k => {
                const kLower = k.toLowerCase();
                return kLower.includes('bank') || kLower.includes('account') || kLower.includes('branch');
              });
              console.log(`ℹ Combined bank header left blank for employee ${index + 1} (no bank data found in Zoho People)`);
              console.log(`  Available bank-related fields:`, availableBankFields);
              console.log(`  All employee fields:`, Object.keys(emp));
            }
          }
          // IFSC Code (Indian Financial System Code) - from People: IFSC_Code
          else if (headerLower.includes('ifsc') && !headerLower.includes('bank') && !headerLower.includes('branch')) {
            const ifscCode = emp.IFSC_Code ||
                         emp['IFSC_Code'] ||
                         emp.IFSCCode ||
                         emp['IFSC Code'] ||
                         emp.ifscCode ||
                         emp.IFSC ||
                         emp['IFSC'] ||
                         '';
           
            row[header] = ifscCode;
            if (ifscCode) {
              console.log(`✓ Mapped IFSC Code for employee ${index + 1}: ${ifscCode}`);
            }
          }
          // Bank A/c Number - from People: Account_Number
          else if ((headerLower.includes('bank') &&
                   (headerLower.includes('account') || headerLower.includes('ac') ||
                    (headerLower.includes('number') && !headerLower.includes('name')))) &&
                   !headerLower.includes('branch')) {
            const accountNumber = emp.Account_Number ||
                         emp['Account_Number'] ||
                         emp.AccountNumber ||
                         emp['Account Number'] ||
                         emp.accountNumber ||
                         '';
           
            row[header] = accountNumber;
            if (accountNumber) {
              console.log(`✓ Mapped "${header}" -> Bank A/c Number for employee ${index + 1}: ${accountNumber} (from Account_Number)`);
            } else {
              console.log(`ℹ Bank A/c Number left blank for employee ${index + 1} (Account_Number not found in Zoho People data)`);
            }
          }
          // Name of Bank - from People: Bank_Name
          else if ((headerLower.includes('name') && headerLower.includes('bank') &&
                   !headerLower.includes('account') && !headerLower.includes('number') &&
                   !headerLower.includes('ifsc') && !headerLower.includes('branch'))) {
            const bankName = emp.Bank_Name ||
                         emp['Bank_Name'] ||
                         emp.BankName ||
                         emp['Bank Name'] ||
                         emp.bankName ||
                         emp['bankName'] ||
                         emp.BANK_NAME ||
                         emp['BANK_NAME'] ||
                         // Try to find any field with "bank" in the name (excluding account, branch, ifsc)
                         (() => {
                           const bankFields = Object.keys(emp).filter(key => {
                             const keyLower = key.toLowerCase();
                             return (keyLower.includes('bank') &&
                                    !keyLower.includes('account') &&
                                    !keyLower.includes('branch') &&
                                    !keyLower.includes('ifsc') &&
                                    !keyLower.includes('code'));
                           });
                           if (bankFields.length > 0) {
                             const value = emp[bankFields[0]];
                             return value && value.toString().trim() ? value.toString().trim() : '';
                           }
                           return '';
                         })() ||
                         '';
           
            row[header] = bankName;
            if (bankName) {
              const sourceField = Object.keys(emp).find(k => emp[k] === bankName) || 'Bank_Name';
              console.log(`✓ Mapped "${header}" -> Name of Bank for employee ${index + 1}: ${bankName} (from ${sourceField})`);
            } else {
              // Log available fields for debugging
              const availableBankFields = Object.keys(emp).filter(k => k.toLowerCase().includes('bank'));
              console.log(`ℹ Name of Bank left blank for employee ${index + 1}`);
              console.log(`  Available bank-related fields:`, availableBankFields);
              console.log(`  All employee fields:`, Object.keys(emp));
            }
          }
          // Branch - from People: Branch_Name
          else if (headerLower.includes('branch') &&
                   !headerLower.includes('ifsc') &&
                   !(headerLower.includes('bank') && headerLower.includes('account'))) {
            const branchName = emp.Branch_Name ||
                         emp['Branch_Name'] ||
                         emp.BranchName ||
                         emp['Branch Name'] ||
                         emp.branchName ||
                         '';
           
            row[header] = branchName;
            if (branchName) {
              console.log(`✓ Mapped "${header}" -> Branch for employee ${index + 1}: ${branchName} (from Branch_Name)`);
            } else {
              console.log(`ℹ Branch left blank for employee ${index + 1} (Branch_Name not found in Zoho People data)`);
            }
          }
          // Reason for Exit - explicitly only use exit-related fields, never use Dateofjoining or other date fields
          else if (headerLower.includes('reason') && headerLower.includes('exit')) {
            // Only check for exit-related fields, do not use any date fields as fallback
            const exitReason = emp.Reason_for_Leaving ||
                         emp['Reason_for_Leaving'] ||
                         emp['Reason for Leaving'] ||
                         emp.ReasonForLeaving ||
                         emp.reasonForLeaving ||
                         emp.Reason_for_Exit ||
                         emp['Reason_for_Exit'] ||
                         emp.ReasonForExit ||
                         emp['ReasonForExit'] ||
                         emp['Reason for Exit'] ||
                         emp['Reason For Exit'] ||
                         emp.ExitReason ||
                         emp['ExitReason'] ||
                         emp['Exit Reason'] ||
                         emp.Reason ||
                         emp['Reason'] ||
                         '';
           
            // Ensure we don't use Dateofjoining or any date field - explicitly set to empty if no exit reason data
            row[header] = exitReason &&
                         !exitReason.toString().toLowerCase().includes('date') &&
                         !exitReason.toString().toLowerCase().includes('join') &&
                         !exitReason.toString().toLowerCase().includes('joining') ? exitReason : '';
           
            if (row[header]) {
              console.log(`✓ Mapped Reason for Exit for employee ${index + 1}: ${row[header]}`);
            }
          }
          // Default: try to find matching field by checking all emp properties
          else {
            // Extract key words from header (e.g., "Name of the employee" -> ["name", "employee"])
            const headerWords = headerLower.split(/\s+/).filter(word =>
              word.length > 2 &&
              !['the', 'of', 'a', 'an', 'and', 'or', 'no', 'no.'].includes(word)
            );
           
            // Try exact match first (case-insensitive, ignoring special chars)
            // Exclude Zoho_ID from matching to Provident Fund, ESIC, Aadhaar, and other specific number fields
            const isSpecificNumberField = headerLower.includes('provident') ||
                                        headerLower.includes('pf') ||
                                        headerLower.includes('esic') ||
                                        headerLower.includes('insurance') ||
                                        headerLower.includes('aadhaar') ||
                                        headerLower.includes('aadhar');
           
            // Exclude date fields from matching to Reason for Exit
            const isExitReasonField = headerLower.includes('reason') && headerLower.includes('exit');
           
            const normalizedHeader = headerLower.replace(/[^a-z0-9]/g, '');
            let headerKey = Object.keys(emp).find(key => {
              const keyLower = key.toLowerCase();
              // Don't match Zoho_ID to specific number fields
              if (isSpecificNumberField && (keyLower.includes('zoho') || key === 'Zoho_ID' || key === 'ZohoID')) {
                return false;
              }
              // Don't match date/joining fields to Reason for Exit
              if (isExitReasonField && (keyLower.includes('date') || keyLower.includes('join') || keyLower.includes('joining'))) {
                return false;
              }
              const normalizedKey = keyLower.replace(/[^a-z0-9]/g, '');
              return normalizedKey === normalizedHeader ||
                     normalizedKey.includes(normalizedHeader) ||
                     normalizedHeader.includes(normalizedKey);
            });
           
            // If no exact match, try matching by keywords
            if (!headerKey && headerWords.length > 0) {
              headerKey = Object.keys(emp).find(key => {
                const keyLower = key.toLowerCase();
                // Don't match Zoho_ID to specific number fields
                if (isSpecificNumberField && (keyLower.includes('zoho') || key === 'Zoho_ID' || key === 'ZohoID')) {
                  return false;
                }
                // Don't match date/joining fields to Reason for Exit
                if (isExitReasonField && (keyLower.includes('date') || keyLower.includes('join') || keyLower.includes('joining'))) {
                  return false;
                }
                // Check if all header words are present in the key
                return headerWords.every(word => keyLower.includes(word)) ||
                       headerWords.some(word => keyLower.includes(word) && keyLower.length < 50);
              });
            }
           
            // If still no match, try partial match
            if (!headerKey) {
              headerKey = Object.keys(emp).find(key => {
                const keyLower = key.toLowerCase();
                // Don't match Zoho_ID to specific number fields
                if (isSpecificNumberField && (keyLower.includes('zoho') || key === 'Zoho_ID' || key === 'ZohoID')) {
                  return false;
                }
                // Don't match date/joining fields to Reason for Exit
                if (isExitReasonField && (keyLower.includes('date') || keyLower.includes('join') || keyLower.includes('joining'))) {
                  return false;
                }
                return headerLower.includes(keyLower) ||
                       keyLower.includes(headerLower) ||
                       keyLower.replace(/[^a-z0-9]/g, '').includes(normalizedHeader) ||
                       normalizedHeader.includes(keyLower.replace(/[^a-z0-9]/g, ''));
              });
            }
           
            if (headerKey) {
              const value = emp[headerKey];
              // Handle displayValue fields (e.g., "EL_Balance.displayValue")
              if (value === '' || value === null || value === undefined) {
                const displayValueKey = `${headerKey}.displayValue`;
                if (emp[displayValueKey]) {
                  row[header] = sanitizeValue(emp[displayValueKey]);
                } else {
                  row[header] = '';
                }
              } else {
                row[header] = sanitizeValue(value);
              }
             
              if (row[header]) {
                console.log(`✓ Mapped "${header}" to "${headerKey}" = "${row[header]}"`);
              }
            } else {
              // Already initialized to empty string, but ensure it stays empty
              row[header] = '';
              console.warn(`✗ Could not map header "${header}". Available fields:`, Object.keys(emp).slice(0, 10));
            }
          }
        });
       
        // Final sanitization: ensure all values are strings (not null/undefined/false)
        currentHeaders.forEach((header) => {
          if (row[header] === null || row[header] === undefined || row[header] === false) {
            row[header] = '';
          } else if (typeof row[header] === 'boolean') {
            // Convert boolean true to string, false to empty string
            row[header] = row[header] ? 'true' : '';
          } else if (typeof row[header] !== 'string') {
            row[header] = String(row[header]);
          }
        });
       
        if (isLikelyFormW && formWPayrollLookup && formWHeadersResolved) {
          const rowNameCandidates =
            formWHeadersResolved.employeeName && row[formWHeadersResolved.employeeName]
              ? [String(row[formWHeadersResolved.employeeName]).trim().toLowerCase()]
              : [];
          const matchedPayrollRow = resolvePayrollRowForEmployee(
            formWPayrollLookup.byPayrollPayload,
            emp,
            row,
            rowNameCandidates
          );
          if (matchedPayrollRow) {
            applyFormWPayrollToRow(
              row,
              buildFormWPayrollMap(matchedPayrollRow),
              formWHeadersResolved
            );
          }
        }

        if (isLikelyFormW || isLikelyForm10 || formBAutofillContext || formDAutofillContext) {
          try {
            const payrollEmployeeId = getPayrollEmployeeId(emp);
            if (payrollEmployeeId) {
              const payrollOrgId =
                process.env.REACT_APP_ZOHO_PAYROLL_ORGANIZATION_ID || '60006183023';
              const payrollQs = new URLSearchParams({
                employee_id: String(payrollEmployeeId),
                organization_id: payrollOrgId,
              });
              const payrollRes = await fetch(`/server/payroll_function?${payrollQs.toString()}`);
              const payrollJson = await payrollRes.json();
              if (payrollRes.ok && payrollJson && payrollJson.success) {
                if (isLikelyFormW && formWHeadersResolved) {
                  applyFormWPayrollToRow(
                    row,
                    buildFormWPayrollMap(payrollJson.data || {}),
                    formWHeadersResolved
                  );
                }

                if (formBAutofillContext) {
                  const p = getPayrollPayloadObject(payrollJson.data || {});
                  const earnings = getEarningsArray(p);
                  const basicEarningsAmt = pickFormBBasicEarningsAmount(earnings);
                  const otherAllowanceAmt = pickFormBOtherAllowanceAmount(earnings);
                  currentHeaders.forEach((header) => {
                    if (isFormBTotalEmolumentsBasicEarningsHeader(header)) {
                      if (basicEarningsAmt !== '' && basicEarningsAmt !== null && basicEarningsAmt !== undefined) {
                        row[header] = sanitizeValue(basicEarningsAmt);
                      }
                    }
                    if (isFormBOtherDeductionsOtherAllowanceHeader(header)) {
                      if (otherAllowanceAmt !== '' && otherAllowanceAmt !== null && otherAllowanceAmt !== undefined) {
                        row[header] = sanitizeValue(otherAllowanceAmt);
                      }
                    }
                  });
                }

                if (formDAutofillContext) {
                  const formDPayrollMap = buildFormDPayrollMap(payrollJson.data || {});
                  currentHeaders.forEach((header) => {
                    if (isFormDBasicSalaryHeader(header)) {
                      if (formDPayrollMap.basicSalary !== '' && formDPayrollMap.basicSalary != null) {
                        row[header] = sanitizeValue(formDPayrollMap.basicSalary);
                      }
                    } else if (isFormDDearnessAllowanceHeader(header)) {
                      if (formDPayrollMap.dearnessAllowance !== '' && formDPayrollMap.dearnessAllowance != null) {
                        row[header] = sanitizeValue(formDPayrollMap.dearnessAllowance);
                      }
                    } else if (isFormDHouseRentAllowanceHeader(header)) {
                      if (formDPayrollMap.houseRentAllowance !== '' && formDPayrollMap.houseRentAllowance != null) {
                        row[header] = sanitizeValue(formDPayrollMap.houseRentAllowance);
                      }
                    } else if (isFormDOtherAllowanceHeader(header)) {
                      if (formDPayrollMap.otherAllowance !== '' && formDPayrollMap.otherAllowance != null) {
                        row[header] = sanitizeValue(formDPayrollMap.otherAllowance);
                      }
                    }
                  });
                }

                if (isLikelyForm10) {
                  const form10PayrollMap = buildForm10PayrollMap(payrollJson.data || {});
                  currentHeaders.forEach((header) => {
                    const headerLower = String(header || '').toLowerCase();
                    let payrollVal = '';
                    if (
                      (headerLower.includes('total') && headerLower.includes('earning')) ||
                      headerLower.includes('total earnings')
                    ) {
                      payrollVal = form10PayrollMap.totalEarnings;
                    } else if (headerLower.includes('normal') && headerLower.includes('earning')) {
                      payrollVal = form10PayrollMap.normalEarnings;
                    } else if (
                      (headerLower.includes('total') && headerLower.includes('overtime') && headerLower.includes('worked')) ||
                      (headerLower.includes('overtime') && headerLower.includes('production'))
                    ) {
                      payrollVal = form10PayrollMap.totalOvertimeWorked;
                    } else if (headerLower.includes('date of which overtime') || (headerLower.includes('date') && headerLower.includes('overtime'))) {
                      const hasOvertime = toNumber(form10PayrollMap.totalOvertimeWorked) > 0;
                      if (hasOvertime && !row[header]) {
                        payrollVal = new Date().toISOString().slice(0, 10);
                      }
                    }

                    if (payrollVal !== '' && payrollVal !== null && payrollVal !== undefined) {
                      row[header] = sanitizeValue(payrollVal);
                    }
                  });
                }
              }
            }
          } catch (payErr) {
            console.warn(`Payroll mapping skipped for employee row ${index + 1}:`, payErr?.message || payErr);
          }
        }

        return row;
      }));

      // Form B: Zoho Payroll `employee_id` often differs from People `EmployeeID` — merge earnings from all_salaries by mail/name.
      if (formBAutofillContext) {
        try {
          const allPayrollRes = await fetch('/server/payroll_function?all_salaries=1');
          const allPayrollJson = await allPayrollRes.json();
          if (allPayrollRes.ok && allPayrollJson && allPayrollJson.success && Array.isArray(allPayrollJson.data)) {
            const payrollRows = allPayrollJson.data;
            const emolHeader = currentHeaders.find((h) => isFormBTotalEmolumentsBasicEarningsHeader(h));
            const otherDedHeader = currentHeaders.find((h) => isFormBOtherDeductionsOtherAllowanceHeader(h));
            if ((emolHeader || otherDedHeader) && payrollRows.length > 0) {
              const resolvePayrollRowForPeople = (em) => {
                if (!em) return null;
                const zid = String(em.Zoho_ID || em['Zoho_ID'] || em.ZohoID || '').trim();
                const eid = String(em.EmployeeID || em['EmployeeID'] || em['Employee ID'] || '').trim();
                const email = String(em.EmailID || em.Email || em.email || em['Email ID'] || '').trim().toLowerCase();
                const fn = String(em.FirstName || em['FirstName'] || '').trim().toLowerCase();
                const ln = String(em.LastName || em['LastName'] || '').trim().toLowerCase();
                const combo = `${fn} ${ln}`.trim();
                let hit = payrollRows.find((r) => zid && String(r.employee_id || '') === zid);
                if (hit) return hit;
                hit = payrollRows.find((r) => eid && String(r.employee_id || '') === eid);
                if (hit) return hit;
                hit = payrollRows.find(
                  (r) => email && String(r.work_mail || r.email || r.work_email || '').trim().toLowerCase() === email
                );
                if (hit) return hit;
                return (
                  payrollRows.find((r) => {
                    const rfn = String(r.first_name || '').trim().toLowerCase();
                    const rln = String(r.last_name || '').trim().toLowerCase();
                    return combo && `${rfn} ${rln}`.trim() === combo;
                  }) || null
                );
              };
              mappedData.forEach((row, rowIndex) => {
                const empItem = employees[rowIndex];
                const em = unwrapEmp(empItem);
                const pr = resolvePayrollRowForPeople(em);
                if (!pr || pr.fetch_error) return;
                const earnings = getEarningsArray(getPayrollPayloadObject(pr));
                const basicAmt = pickFormBBasicEarningsAmount(earnings);
                const otherAmt = pickFormBOtherAllowanceAmount(earnings);
                if (emolHeader && basicAmt !== '' && basicAmt != null) {
                  row[emolHeader] = sanitizeValue(basicAmt);
                }
                if (otherDedHeader && otherAmt !== '' && otherAmt != null) {
                  row[otherDedHeader] = sanitizeValue(otherAmt);
                }
              });
            }
          }
        } catch (formBPayErr) {
          console.warn('Form B all_salaries payroll enrich skipped:', formBPayErr?.message || formBPayErr);
        }
      }

      // Form W / Form 10 payroll fallback: fill any cells still empty after per-row preload.
      if (isLikelyFormW || isLikelyForm10) {
        try {
          let byPayrollPayload = formWPayrollLookup?.byPayrollPayload;
          if (!byPayrollPayload || byPayrollPayload.size === 0) {
            const payrollOrgId =
              process.env.REACT_APP_ZOHO_PAYROLL_ORGANIZATION_ID || '60006183023';
            const allPayrollQs = new URLSearchParams({
              all_salaries: '1',
              organization_id: payrollOrgId,
            });
            const allPayrollRes = await fetch(`/server/payroll_function?${allPayrollQs.toString()}`);
            const allPayrollJson = await allPayrollRes.json();
            if (allPayrollRes.ok && allPayrollJson?.success && Array.isArray(allPayrollJson.data)) {
              byPayrollPayload = buildPayrollLookupFromRows(allPayrollJson.data).byPayrollPayload;
            }
          }

          const formWHeaders = formWHeadersResolved || (isLikelyFormW ? resolveFormWTableHeaders(currentHeaders) : null);

          const normalEarningsHeader = currentHeaders.find((h) => {
            const t = String(h || '').toLowerCase();
            return t.includes('normal') && t.includes('earning');
          });
          const totalEarningsHeader = currentHeaders.find((h) => {
            const t = String(h || '').toLowerCase();
            return t.includes('total') && t.includes('earning');
          });

          if (byPayrollPayload && ((isLikelyFormW && formWHeaders) || normalEarningsHeader || totalEarningsHeader)) {
            mappedData.forEach((row, rowIndex) => {
              const empItem = employees[rowIndex];
              const emp = empItem && (empItem.Employee || empItem.employee || empItem);
              const rowNameCandidates =
                formWHeaders?.employeeName && row[formWHeaders.employeeName]
                  ? [String(row[formWHeaders.employeeName]).trim().toLowerCase()]
                  : [];
              const matchedPayrollRow = resolvePayrollRowForEmployee(
                byPayrollPayload,
                emp,
                row,
                rowNameCandidates
              );

              if (isLikelyFormW && matchedPayrollRow && formWHeaders) {
                applyFormWPayrollToRow(row, buildFormWPayrollMap(matchedPayrollRow), formWHeaders);
              }

              if (isLikelyForm10 && matchedPayrollRow) {
                const form10PayrollMap = buildForm10PayrollMap(matchedPayrollRow);
                const monthlySalary = form10PayrollMap.totalEarnings;
                if (normalEarningsHeader && !String(row[normalEarningsHeader] || '').trim() && monthlySalary !== '') {
                  row[normalEarningsHeader] = sanitizeValue(monthlySalary);
                }
                if (totalEarningsHeader && !String(row[totalEarningsHeader] || '').trim() && monthlySalary !== '') {
                  row[totalEarningsHeader] = sanitizeValue(monthlySalary);
                }
              }
            });
          }
        } catch (allPayrollErr) {
          console.warn('Form W / Form-10 all_salaries fallback skipped:', allPayrollErr?.message || allPayrollErr);
        }
      }

      if (formDAutofillContext) {
        try {
          const payrollOrgId =
            process.env.REACT_APP_ZOHO_PAYROLL_ORGANIZATION_ID || '60006183023';
          const allPayrollQs = new URLSearchParams({
            all_salaries: '1',
            organization_id: payrollOrgId,
          });
          const allPayrollRes = await fetch(`/server/payroll_function?${allPayrollQs.toString()}`);
          const allPayrollJson = await allPayrollRes.json();
          if (allPayrollRes.ok && allPayrollJson?.success && Array.isArray(allPayrollJson.data)) {
            const byPayrollPayload = buildPayrollLookupFromRows(allPayrollJson.data).byPayrollPayload;
            if (byPayrollPayload && byPayrollPayload.size > 0) {
              mappedData.forEach((row, rowIndex) => {
                const empItem = employees[rowIndex];
                const emp = empItem && (empItem.Employee || empItem.employee || empItem);
                const matchedPayrollRow = resolvePayrollRowForEmployee(byPayrollPayload, emp, row, []);
                if (!matchedPayrollRow) return;

                const formDPayrollMap = buildFormDPayrollMap(matchedPayrollRow);
                currentHeaders.forEach((header) => {
                  if (
                    isFormDBasicSalaryHeader(header) &&
                    !String(row[header] || '').trim() &&
                    formDPayrollMap.basicSalary !== '' &&
                    formDPayrollMap.basicSalary != null
                  ) {
                    row[header] = sanitizeValue(formDPayrollMap.basicSalary);
                  } else if (
                    isFormDDearnessAllowanceHeader(header) &&
                    !String(row[header] || '').trim() &&
                    formDPayrollMap.dearnessAllowance !== '' &&
                    formDPayrollMap.dearnessAllowance != null
                  ) {
                    row[header] = sanitizeValue(formDPayrollMap.dearnessAllowance);
                  } else if (
                    isFormDHouseRentAllowanceHeader(header) &&
                    !String(row[header] || '').trim() &&
                    formDPayrollMap.houseRentAllowance !== '' &&
                    formDPayrollMap.houseRentAllowance != null
                  ) {
                    row[header] = sanitizeValue(formDPayrollMap.houseRentAllowance);
                  } else if (
                    isFormDOtherAllowanceHeader(header) &&
                    !String(row[header] || '').trim() &&
                    formDPayrollMap.otherAllowance !== '' &&
                    formDPayrollMap.otherAllowance != null
                  ) {
                    row[header] = sanitizeValue(formDPayrollMap.otherAllowance);
                  }
                });
              });
            }
          }
        } catch (formDPayErr) {
          console.warn('Form D all_salaries payroll fallback skipped:', formDPayErr?.message || formDPayErr);
        }
      }
     
      console.log('Mapped data:', mappedData);
      console.log('Sample mapped row:', mappedData[0]);

      // Fetch attendance data and populate Total Hours / Daily worked hours columns
      let attendanceHoursPopulated = 0;
      let hasTotalHoursColumn = false;
      let hasDailyHoursColumn = false;
      const attendanceAggByEmployeeKey = new Map();
      try {
        const now = new Date();
        let targetMonthIndex = now.getMonth();
        const selectedMonthIndex = MONTH_NAMES.findIndex(
          (month) => month.toLowerCase().startsWith(String(selectedMonth || '').toLowerCase().trim())
        );
        if (selectedMonthIndex >= 0) {
          targetMonthIndex = selectedMonthIndex;
        }

        const targetYear = now.getFullYear();
        const startDate = new Date(targetYear, targetMonthIndex, 1);
        const endDate = new Date(targetYear, targetMonthIndex + 1, 0);
        const sdate = formatLocalDateYYYYMMDD(startDate);
        const edate = formatLocalDateYYYYMMDD(endDate);

        const todayCalendarStart = new Date(now.getFullYear(), now.getMonth(), now.getDate());
        const attendanceMonthNotYetStarted = startDate > todayCalendarStart;

        if (attendanceMonthNotYetStarted) {
          console.log(
            `Skipping attendance merge: wage period ${sdate}..${edate} has not started yet (today=${formatLocalDateYYYYMMDD(todayCalendarStart)}).`
          );
        }

        const attendanceRes =
          attendanceMonthNotYetStarted
            ? { ok: false }
            : await fetch(`/server/attendance_function?limit=200&sdate=${encodeURIComponent(sdate)}&edate=${encodeURIComponent(edate)}`);
        const attendanceJson = attendanceMonthNotYetStarted ? {} : await attendanceRes.json();

        if (attendanceRes.ok && attendanceJson.success && attendanceJson.data) {
          const tryParseJson = (val) => {
            if (!val) return null;
            if (typeof val === 'object') return val;
            if (typeof val !== 'string') return null;
            const s = String(val).trim();
            if (!s) return null;
            try {
              return JSON.parse(s);
            } catch (_) {
              return null;
            }
          };

          const mergeEmployeeMeta = (attendanceRow, employeeMeta) => {
            if (!employeeMeta || typeof employeeMeta !== 'object') return attendanceRow;
            return { ...attendanceRow, ...employeeMeta };
          };

          const normalizeAttendance = (rawData) => {
            if (!rawData) return [];
            const rows = [];

            // Common shape from attendance API: { "yyyy-mm-dd": { ...attendance fields... } }
            const dateKeyRegex = /^\d{4}-\d{2}-\d{2}$/;
            const collectDateMapRows = (obj, employeeMeta) => {
              if (!obj || typeof obj !== 'object') return;
              const directDateMap = Object.keys(obj).filter((k) => dateKeyRegex.test(k));
              if (directDateMap.length === 0) return;
              directDateMap.forEach((date) => {
                const val = obj[date];
                if (Array.isArray(val)) {
                  val.forEach((item) => {
                    if (item && typeof item === 'object') rows.push(mergeEmployeeMeta({ date, ...item }, employeeMeta));
                  });
                  return;
                }
                if (val && typeof val === 'object') {
                  const nestedKeys = Object.keys(val);
                  const looksLikeSingleRecord =
                    ('ShiftStartTime' in val) ||
                    ('ShiftEndTime' in val) ||
                    ('TotalHours' in val) ||
                    ('WorkingHours' in val);
                  if (looksLikeSingleRecord) {
                    rows.push(mergeEmployeeMeta({ date, ...val }, employeeMeta));
                    return;
                  }
                  nestedKeys.forEach((k) => {
                    const nestedVal = val[k];
                    if (nestedVal && typeof nestedVal === 'object') {
                      rows.push(mergeEmployeeMeta({ date, ...nestedVal }, employeeMeta));
                    }
                  });
                }
              });
            };

            const visit = (node, employeeMeta = null) => {
              if (!node) return;
              if (Array.isArray(node)) {
                node.forEach((it) => visit(it, employeeMeta));
                return;
              }
              if (typeof node === 'string') {
                const parsed = tryParseJson(node);
                if (parsed) visit(parsed, employeeMeta);
                return;
              }
              if (typeof node !== 'object') return;

              // Handle row format: { attendanceDetails: "...", employeeDetails: "..." }
              const empMeta = tryParseJson(node.employeeDetails) || node.employeeDetails || employeeMeta;
              const attDetails = tryParseJson(node.attendanceDetails) || node.attendanceDetails;
              if (attDetails) {
                collectDateMapRows(attDetails, empMeta);
                visit(attDetails, empMeta);
              }

              collectDateMapRows(node, empMeta);

              Object.values(node).forEach((child) => {
                if (child && (typeof child === 'object' || typeof child === 'string')) {
                  visit(child, empMeta);
                }
              });
            };

            const candidate = rawData.response || rawData.result || rawData.data || rawData.records || rawData.record || rawData;
            visit(candidate, null);
            if (rows.length > 0) return rows;

            if (Array.isArray(candidate)) return candidate;
            if (candidate && typeof candidate === 'object') return [candidate];
            return [];
          };

          const recordDateKeyStable = (obj) => {
            const dRaw = obj?.date || obj?.Date || '';
            if (!dRaw) return '';
            const s = String(dRaw).trim();
            const isoPrefix = s.match(/^(\d{4}-\d{2}-\d{2})/);
            if (isoPrefix) return isoPrefix[1];
            const d = new Date(s);
            if (!Number.isNaN(d.getTime())) return formatLocalDateYYYYMMDD(d);
            return '';
          };

          let attendanceRecords = normalizeAttendance(attendanceJson.data).filter((rec) => {
            const k = recordDateKeyStable(rec);
            return Boolean(k && k >= sdate && k <= edate);
          });
          if (attendanceRecords.length > 0) {
            const sortedRecords = [...attendanceRecords].sort((a, b) =>
              String(recordDateKeyStable(b) || b.date || '').localeCompare(String(recordDateKeyStable(a) || a.date || ''))
            );

            const totalWorkedDaysHeader = currentHeaders.find((header) => {
              const h = String(header || '').toLowerCase().trim();
              if (h.includes('overtime') || h.includes('loss') || h.includes('lop') || h.includes('holiday')) return false;
              if (isLikelyDayOfMonthColumnHeader(header)) return false;
              return (
                (h.includes('total') && (h.includes('day') || h.includes('days')) && h.includes('work')) ||
                /total\s*days?\s*worked/i.test(String(header || '')) ||
                /^totaldaysworked$/i.test(h.replace(/\s/g, ''))
              );
            });

            const totalHoursWorkedAggHeader = currentHeaders.find((header) => {
              const h = String(header || '').toLowerCase().trim();
              if (h.includes('overtime') || h.includes('normal') || h.includes('daily')) return false;
              if (isLikelyDayOfMonthColumnHeader(header)) return false;
              return (
                (h.includes('total') && h.includes('hour') && h.includes('work')) ||
                /total\s*hours?\s*worked/i.test(String(header || '')) ||
                /^totalhoursworked$/i.test(h.replace(/\s/g, ''))
              );
            });

            const form25LossOfPayHeader = currentHeaders.find((header) => isForm25LossOfPayDaysHeader(header));

            const totalHoursHeader = currentHeaders.find((header) => {
              const h = String(header || '').toLowerCase().trim();
              if (header === totalHoursWorkedAggHeader || header === totalWorkedDaysHeader) return false;
              if (h.includes('overtime') || h.includes('normal') || h.includes('daily')) return false;
              if (isLikelyDayOfMonthColumnHeader(header)) return false;
              return h.includes('total') && h.includes('hour');
            });
            const normalHoursHeader = currentHeaders.find((header) => {
              const h = header.toLowerCase().trim();
              return h.includes('normal') && h.includes('hour');
            });

            const dailyWorkedHoursHeader = currentHeaders.find((header) => {
              const h = header.toLowerCase().trim();
              return h.includes('daily') &&
                h.includes('hour') &&
                (h.includes('work') || h.includes('worked'));
            });
            const dateDayHeaders = currentHeaders
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

            if (totalHoursHeader || totalHoursWorkedAggHeader || totalWorkedDaysHeader) {
              hasTotalHoursColumn = true;
            }
            if (normalHoursHeader) {
              hasTotalHoursColumn = true;
            }
            if (dailyWorkedHoursHeader) {
              hasDailyHoursColumn = true;
            }

            if (
              totalHoursHeader ||
              totalHoursWorkedAggHeader ||
              totalWorkedDaysHeader ||
              dailyWorkedHoursHeader ||
              normalHoursHeader ||
              dateDayHeaders.length > 0
            ) {
              const getAttendanceCandidateIds = (obj) => {
                if (!obj || typeof obj !== 'object') return [];
                const vals = [
                  obj.EmployeeID,
                  obj['EmployeeID'],
                  obj['Employee ID'],
                  obj.employee_id,
                  obj['employee_id'],
                  obj.Zoho_ID,
                  obj['Zoho_ID'],
                  obj.zoho_id,
                  obj['zoho_id'],
                  obj.ZohoID,
                  obj['ZohoID'],
                  obj.userId,
                  obj.UserID,
                  obj.UserId,
                  obj.employeeId,
                  obj.EmpID,
                  obj.EmpId,
                  findValueByNormalizedKey(obj, 'employeeid'),
                  findValueByNormalizedKey(obj, 'zohoid'),
                  findValueByNormalizedKey(obj, 'userid')
                ];
                return vals.map((v) => String(v || '').trim()).filter(Boolean);
              };

              const getAttendanceDisplayName = (obj) => {
                if (!obj || typeof obj !== 'object') return '';
                const first = String(
                  obj.firstName ||
                  obj.FirstName ||
                  obj.First_Name ||
                  obj['First_Name'] ||
                  obj['First Name'] ||
                  findValueByNormalizedKey(obj, 'firstname') ||
                  ''
                ).trim();
                const last = String(
                  obj.lastName ||
                  obj.LastName ||
                  obj.Last_Name ||
                  obj['Last_Name'] ||
                  obj['Last Name'] ||
                  findValueByNormalizedKey(obj, 'lastname') ||
                  ''
                ).trim();
                const full = `${first} ${last}`.trim();
                return String(
                  full ||
                  obj.EmployeeName ||
                  obj['Employee Name'] ||
                  obj.EmployeeName ||
                  obj['EmployeeName'] ||
                  obj.Name ||
                  obj['Name'] ||
                  obj.firstName ||
                  obj.FirstName ||
                  findValueByNormalizedKey(obj, 'name') ||
                  ''
                ).trim().toLowerCase();
              };
              /** Daily punch hours only — do not use TotalHours (often monthly / misleading per day). */
              const getExplicitDailyHoursRaw = (obj) =>
                String(
                  obj?.WorkingHours ||
                    obj?.workingHours ||
                    obj?.DailyWorkedHours ||
                    obj?.dailyWorkedHours ||
                    ''
                ).trim();

              const hasParsableShiftPair = (rec) => {
                if (!rec || typeof rec !== 'object') return false;
                const ss =
                  rec.ShiftStartTime ||
                  rec['ShiftStartTime'] ||
                  rec.shiftStartTime ||
                  findValueByNormalizedKey(rec, 'shiftstarttime') ||
                  '';
                const se =
                  rec.ShiftEndTime ||
                  rec['ShiftEndTime'] ||
                  rec.shiftEndTime ||
                  findValueByNormalizedKey(rec, 'shiftendtime') ||
                  '';
                const st = String(ss).trim();
                const et = String(se).trim();
                if (!st || !et) return false;
                return parseTimeToMinutes(st) != null && parseTimeToMinutes(et) != null;
              };

              const minutesFromExplicitDailyOrShift = (rec) => {
                if (!rec || typeof rec !== 'object') return null;
                const daily = getExplicitDailyHoursRaw(rec);
                if (daily) {
                  if (/^0{1,2}:0{2}$/.test(daily)) return null;
                  const m = parseTimeToMinutes(daily);
                  if (m != null && m > 0) return m;
                  const n = Number(String(daily).replace(/,/g, '').trim());
                  if (Number.isFinite(n) && n > 0 && n <= 24 && !daily.includes(':')) {
                    return Math.round(n * 60);
                  }
                }
                if (hasParsableShiftPair(rec)) {
                  const nh = computeNormalHoursFromAttendanceRecord(rec);
                  if (nh) {
                    const m = parseTimeToMinutes(nh);
                    if (m != null && m > 0) return m;
                  }
                }
                return null;
              };

              /** True only when explicit daily hours exist OR both shift ends are present (times required). */
              const hasValidAttendanceTime = (rec) => minutesFromExplicitDailyOrShift(rec) != null;

              const formatHoursForDayCell = (rec) => {
                const daily = getExplicitDailyHoursRaw(rec);
                if (daily) {
                  const dm = parseTimeToMinutes(daily);
                  if (dm != null && dm > 0) return daily;
                  const n = Number(String(daily).replace(/,/g, '').trim());
                  if (Number.isFinite(n) && n > 0 && n <= 24 && !daily.includes(':')) return daily;
                }
                if (hasParsableShiftPair(rec)) {
                  const nh = computeNormalHoursFromAttendanceRecord(rec);
                  return nh || '';
                }
                return '';
              };

              const recordDateKey = recordDateKeyStable;

              const attendanceById = new Map();
              const attendanceByName = new Map();
              const attendanceByFirstName = new Map();
              const attendanceByIdDate = new Map();
              const attendanceByNameDate = new Map();
              const attendanceByFirstNameDate = new Map();
              const workedMinutesFromAttendanceRecord = (rec) => minutesFromExplicitDailyOrShift(rec);

              const isCountableWorkedDay = (rec) => hasValidAttendanceTime(rec);

              const formatCumulativeHoursMM = (totalMins) => {
                if (!Number.isFinite(totalMins) || totalMins <= 0) return '';
                const h = Math.floor(totalMins / 60);
                const m = Math.round(totalMins % 60);
                return `${h}:${String(m).padStart(2, '0')}`;
              };

              const aggregateWorkedDaysAndMinutesForEmployee = (idCands, nameCand, firstCand) => {
                const byDate = new Map();
                sortedRecords.forEach((rec) => {
                  const dk = recordDateKeyStable(rec);
                  if (!dk || dk < sdate || dk > edate) return;
                  let match = false;
                  const idsRec = getAttendanceCandidateIds(rec);
                  if (idCands.some((id) => idsRec.includes(id))) match = true;
                  else if (nameCand && getAttendanceDisplayName(rec) === nameCand) match = true;
                  else if (
                    firstCand &&
                    String(
                      rec?.firstName ||
                        rec?.FirstName ||
                        rec?.First_Name ||
                        findValueByNormalizedKey(rec, 'firstname') ||
                        ''
                    )
                      .trim()
                      .toLowerCase() === firstCand
                  ) {
                    match = true;
                  }
                  if (!match) return;
                  const prev = byDate.get(dk);
                  if (!prev || (hasValidAttendanceTime(rec) && !hasValidAttendanceTime(prev))) byDate.set(dk, rec);
                });
                let days = 0;
                let totalMins = 0;
                byDate.forEach((rec) => {
                  if (!isCountableWorkedDay(rec)) return;
                  days += 1;
                  const m = workedMinutesFromAttendanceRecord(rec);
                  if (m != null) totalMins += m;
                });
                return { days, totalMins };
              };

              sortedRecords.forEach((rec) => {
                getAttendanceCandidateIds(rec).forEach((id) => {
                  if (!attendanceById.has(id)) attendanceById.set(id, rec);
                });
                const nm = getAttendanceDisplayName(rec);
                if (nm && !attendanceByName.has(nm)) attendanceByName.set(nm, rec);
                const firstNm = String(
                  rec?.firstName ||
                  rec?.FirstName ||
                  rec?.First_Name ||
                  rec?.['First_Name'] ||
                  findValueByNormalizedKey(rec, 'firstname') ||
                  ''
                ).trim().toLowerCase();
                if (firstNm && !attendanceByFirstName.has(firstNm)) attendanceByFirstName.set(firstNm, rec);
                const rDate = recordDateKey(rec);
                if (rDate && hasValidAttendanceTime(rec)) {
                  getAttendanceCandidateIds(rec).forEach((id) => {
                    const key = `${id}::${rDate}`;
                    if (!attendanceByIdDate.has(key)) attendanceByIdDate.set(key, rec);
                  });
                  if (nm) {
                    const nKey = `${nm}::${rDate}`;
                    if (!attendanceByNameDate.has(nKey)) attendanceByNameDate.set(nKey, rec);
                  }
                  if (firstNm) {
                    const fKey = `${firstNm}::${rDate}`;
                    if (!attendanceByFirstNameDate.has(fKey)) attendanceByFirstNameDate.set(fKey, rec);
                  }
                }
              });

              mappedData.forEach((row, rowIndex) => {
                const empItem = employees[rowIndex];
                const emp = empItem && (empItem.Employee || empItem.employee || empItem);
                const idCandidates = [
                  ...getAttendanceCandidateIds(emp || {}),
                  String(getPayrollEmployeeId(emp || {}) || '').trim()
                ].filter(Boolean);
                const nameCandidates = getEmployeeNameCandidates(emp || {});
                const nameCandidate = nameCandidates[0] || '';
                const firstNameCandidate = String(
                  (emp && (
                    emp.FirstName ||
                    emp['FirstName'] ||
                    emp.First_Name ||
                    emp['First_Name'] ||
                    ''
                  )) || ''
                ).trim().toLowerCase();

                let matchedAttendance = null;
                for (let i = 0; i < idCandidates.length; i++) {
                  const rec = attendanceById.get(idCandidates[i]);
                  if (rec) {
                    matchedAttendance = rec;
                    break;
                  }
                }
                if (!matchedAttendance && nameCandidates.length > 0) {
                  matchedAttendance = nameCandidates.map((candidate) => attendanceByName.get(candidate)).find(Boolean) || null;
                }
                if (!matchedAttendance && firstNameCandidate) {
                  matchedAttendance = attendanceByFirstName.get(firstNameCandidate) || null;
                }

                const agg = aggregateWorkedDaysAndMinutesForEmployee(
                  idCandidates,
                  nameCandidate,
                  firstNameCandidate
                );

                idCandidates.forEach((id) => {
                  const key = String(id || '').trim().toLowerCase();
                  if (key) attendanceAggByEmployeeKey.set(`id:${key}`, agg);
                });
                nameCandidates.forEach((candidate) => {
                  attendanceAggByEmployeeKey.set(`name:${candidate}`, agg);
                });
                if (firstNameCandidate) attendanceAggByEmployeeKey.set(`first:${firstNameCandidate}`, agg);

                if (totalWorkedDaysHeader) {
                  row[totalWorkedDaysHeader] = String(agg.days);
                }
                if (form25LossOfPayHeader) {
                  const daysInMonth = new Date(targetYear, targetMonthIndex + 1, 0).getDate();
                  row[form25LossOfPayHeader] = String(Math.max(0, daysInMonth - agg.days));
                }
                if (totalHoursWorkedAggHeader) {
                  row[totalHoursWorkedAggHeader] = formatCumulativeHoursMM(agg.totalMins);
                }

                const totalHoursValue = matchedAttendance
                  ? String(matchedAttendance.TotalHours || matchedAttendance.totalHours || '')
                  : '';
                const dailyWorkedHoursValue = matchedAttendance
                  ? String(
                      matchedAttendance.WorkingHours ||
                        matchedAttendance.workingHours ||
                        matchedAttendance.DailyWorkedHours ||
                        matchedAttendance.dailyWorkedHours ||
                        ''
                    )
                  : '';

                if (totalHoursHeader && totalHoursValue) row[totalHoursHeader] = totalHoursValue;
                if (dailyWorkedHoursHeader && dailyWorkedHoursValue) row[dailyWorkedHoursHeader] = dailyWorkedHoursValue;
                if (normalHoursHeader) {
                  let normalFromAttendance = computeNormalHoursFromAttendanceRecord(matchedAttendance);
                  if (!normalFromAttendance && matchedAttendance) {
                    const idsEmp = new Set(getAttendanceCandidateIds(matchedAttendance));
                    const nmEmp = getAttendanceDisplayName(matchedAttendance);
                    const sameEmpShiftRec = sortedRecords.find((rec) => {
                      const sameEmp =
                        (idsEmp.size && getAttendanceCandidateIds(rec).some((id) => idsEmp.has(id))) ||
                        Boolean(nmEmp && getAttendanceDisplayName(rec) === nmEmp);
                      return sameEmp && computeNormalHoursFromAttendanceRecord(rec);
                    });
                    normalFromAttendance = computeNormalHoursFromAttendanceRecord(sameEmpShiftRec);
                  }
                  if (normalFromAttendance) {
                    row[normalHoursHeader] = normalFromAttendance;
                  } else if (!String(row[normalHoursHeader] || '').trim() && (dailyWorkedHoursValue || totalHoursValue)) {
                    row[normalHoursHeader] = dailyWorkedHoursValue || totalHoursValue;
                  }
                }
                if (dateDayHeaders.length > 0) {
                  const today = new Date();
                  const isCurrentSelectedMonth =
                    targetYear === today.getFullYear() && targetMonthIndex === today.getMonth();
                  const isWeekendCalendarDay = (d) => {
                    const dow = d.getDay();
                    return dow === 0 || dow === 6;
                  };
                  dateDayHeaders.forEach(({ header, day }) => {
                    const dayDate = new Date(targetYear, targetMonthIndex, day);
                    if (dayDate.getMonth() !== targetMonthIndex) return; // skip impossible days for selected month
                    if (isCurrentSelectedMonth && day > today.getDate()) {
                      row[header] = '';
                      return;
                    }
                    const iso = formatLocalDateYYYYMMDD(dayDate);
                    const byId = idCandidates
                      .map((id) => attendanceByIdDate.get(`${id}::${iso}`))
                      .find(Boolean);
                    const byName = !byId && nameCandidate ? attendanceByNameDate.get(`${nameCandidate}::${iso}`) : null;
                    const byFirstName = !byId && !byName && firstNameCandidate
                      ? attendanceByFirstNameDate.get(`${firstNameCandidate}::${iso}`)
                      : null;
                    const rec = byId || byName || byFirstName;
                    if (rec && hasValidAttendanceTime(rec)) {
                      row[header] = formatHoursForDayCell(rec);
                    } else if (isWeekendCalendarDay(dayDate)) {
                      // Saturday/Sunday: prefer any formatted hours from attendance row, else shift from People
                      const fromAtt = rec ? formatHoursForDayCell(rec) : '';
                      row[header] = fromAtt || computeNormalHoursFromShift(emp || {}) || '';
                    } else {
                      row[header] = '';
                    }
                  });
                }
                attendanceHoursPopulated++;
              });
              console.log(`Attendance hours populated for ${sdate}..${edate}; rows=${mappedData.length}`);
            }
          }
        }
      } catch (attendanceErr) {
        console.error('Error fetching attendance data for hours columns:', attendanceErr);
        // Do not fail autofill when attendance endpoint is unavailable
      }

      // Form W: Number of days worked from Form 25 Total Worked Days
      if (isLikelyFormW) {
        try {
          console.log('Fetching Form 25 data for Form W days worked...');
          const form25Res = await fetch('/server/form25_function/form25?perPage=200');
          const form25Json = await form25Res.json();
          let form25Records = [];

          const form25Ok =
            form25Res.ok &&
            (form25Json.success === true || form25Json.status === 'success');
          if (form25Ok && Array.isArray(form25Json.data?.form25Data)) {
            form25Records = form25Json.data.form25Data;
          } else {
            console.warn('Form 25 fetch for Form W returned unexpected response:', form25Json);
          }

          const form25ByName = new Map();
          const form25ById = new Map();
          form25Records.forEach((record) => {
            const name = String(record.nameOfTheWorker || '').trim().toLowerCase();
            const id = String(record.workerIdentityNumber || '').trim().toLowerCase();
            const firstNameOnly = name ? name.split(/\s+/)[0] : '';
            if (name) form25ByName.set(name, record);
            if (firstNameOnly && !form25ByName.has(firstNameOnly)) form25ByName.set(firstNameOnly, record);
            if (id) form25ById.set(id, record);
          });

          const daysWorkedHeader =
            formWHeadersResolved?.daysWorked ||
            currentHeaders.find((header) => isFormWDaysWorkedHeader(header));
          const employeeNameHeader =
            formWHeadersResolved?.employeeName ||
            currentHeaders.find((header) => {
              const headerLower = header.toLowerCase().trim();
              return (
                headerLower.includes('name') &&
                (headerLower.includes('employee') || headerLower.includes('worker') || headerLower.includes('emp'))
              );
            });
          const employeeIdHeader =
            formWHeadersResolved?.employeeId ||
            currentHeaders.find((header) => {
              const headerLower = header.toLowerCase().trim();
              return (
                (headerLower.includes('identification') ||
                  (headerLower.includes('employee') && headerLower.includes('id')) ||
                  (headerLower.includes('worker') && headerLower.includes('id')) ||
                  (headerLower.includes('emp') && headerLower.includes('code'))) &&
                !headerLower.includes('zoho')
              );
            });

          if (daysWorkedHeader && form25Records.length > 0) {
            let daysWorkedPopulated = 0;
            mappedData.forEach((row, rowIndex) => {
              const empItem = employees[rowIndex];
              const emp = empItem && (empItem.Employee || empItem.employee || empItem);
              if (!emp) return;

              const empNameCandidates = getEmployeeNameCandidates(emp || {});
              const rowEmployeeName =
                employeeNameHeader && row[employeeNameHeader]
                  ? String(row[employeeNameHeader]).toLowerCase().trim()
                  : '';
              const allNameCandidates = Array.from(
                new Set([rowEmployeeName, ...empNameCandidates].filter(Boolean))
              );
              const empId = String(
                (employeeIdHeader && row[employeeIdHeader] ? row[employeeIdHeader] : '') ||
                  getFallbackWorkerId(emp || {}) ||
                  getFallbackEmployeeId(emp || '') ||
                  ''
              )
                .toLowerCase()
                .trim();

              let matchedForm25 = null;
              for (const name of allNameCandidates) {
                if (form25ByName.has(name)) {
                  matchedForm25 = form25ByName.get(name);
                  break;
                }
              }
              if (!matchedForm25 && empId && form25ById.has(empId)) {
                matchedForm25 = form25ById.get(empId);
              }
              if (!matchedForm25 && allNameCandidates.length > 0) {
                const probe = allNameCandidates[0];
                for (const [key, record] of form25ByName.entries()) {
                  if (key.includes(probe) || probe.includes(key)) {
                    matchedForm25 = record;
                    break;
                  }
                }
              }
              if (!matchedForm25) {
                matchedForm25 = form25Records[rowIndex] || null;
              }

              const daysWorked = pickForm25FieldValue(matchedForm25, [
                'totalDaysWorked',
                'TotalDaysWorked',
                'Total Worked Days',
                'total_days_worked',
              ]);
              if (daysWorked !== '') {
                row[daysWorkedHeader] = daysWorked;
                daysWorkedPopulated++;
              }
            });
            console.log(
              `Form W days worked (Form 25 Total Worked Days) populated for ${daysWorkedPopulated} out of ${mappedData.length} employees`
            );
          }
        } catch (form25FormWErr) {
          console.warn('Form 25 data fetch for Form W skipped:', form25FormWErr?.message || form25FormWErr);
        }
      }

      // Form A: Fetch data from Form 25 for specific fields
      if (formAAutofillContext) {
        try {
          console.log('Fetching Form 25 data for Form A mappings...');
          const form25Res = await fetch('/server/form25_function/form25?perPage=200');
          const form25Json = await form25Res.json();
          let form25Records = [];

          const form25Ok =
            form25Res.ok &&
            (form25Json.success === true || form25Json.status === 'success');
          if (form25Ok && Array.isArray(form25Json.data?.form25Data)) {
            form25Records = form25Json.data.form25Data;
          } else {
            console.warn('Form 25 fetch for Form A returned unexpected response:', form25Json);
          }

          // Create maps for Form 25 data by employee identifier
          const form25ByName = new Map();
          const form25ById = new Map();

          form25Records.forEach((record) => {
            const name = String(record.nameOfTheWorker || '').trim().toLowerCase();
            const id = String(record.workerIdentityNumber || '').trim().toLowerCase();
            const firstNameOnly = name ? name.split(/\s+/)[0] : '';

            if (name) {
              form25ByName.set(name, record);
            }
            if (firstNameOnly && !form25ByName.has(firstNameOnly)) {
              form25ByName.set(firstNameOnly, record);
            }
            if (id) {
              form25ById.set(id, record);
            }
          });

          // Find Form A specific headers using normalized matching because merged Excel headers
          // can come through as "(3) Nature of work", "Month No.of days employed", etc.
          const normalizeFormAHeader = (header) =>
            String(header || '')
              .toLowerCase()
              .replace(/\r?\n/g, ' ')
              .replace(/\s+/g, ' ')
              .trim();
          const compactFormAHeader = (header) => normalizeFormAHeader(header).replace(/[^a-z0-9]/g, '');
          const pickForm25DayValue = (record, keys) => {
            if (!record || typeof record !== 'object' || !Array.isArray(keys)) return '';
            for (let i = 0; i < keys.length; i += 1) {
              const key = keys[i];
              const direct = record[key];
              if (direct != null && String(direct).trim() !== '') return String(direct);
              const normalized = findValueByNormalizedKey(record, key);
              if (normalized != null && String(normalized).trim() !== '') return String(normalized);
            }
            return '';
          };

          const natureOfWorkHeader = currentHeaders.find((header) => {
            const norm = normalizeFormAHeader(header);
            return norm.includes('nature') && norm.includes('work');
          });

          const dateOnWhichEmployedHeader = currentHeaders.find((header) => {
            const norm = normalizeFormAHeader(header);
            return norm.includes('date') && norm.includes('employed');
          });

          const noOfDaysEmployedHeader = currentHeaders.find((header) => {
            const norm = normalizeFormAHeader(header);
            const compact = compactFormAHeader(header);
            if (norm.includes('laid off') || norm.includes('not employed')) return false;
            return compact.includes('noofdaysemployed') ||
              compact.includes('numberofdaysemployed') ||
              (norm.includes('days') && norm.includes('employed'));
          });

          const noOfDaysLaidOffHeader = currentHeaders.find((header) => {
            const norm = normalizeFormAHeader(header);
            const compact = compactFormAHeader(header);
            return compact.includes('noofdayslaidoff') ||
              compact.includes('numberofdayslaidoff') ||
              (norm.includes('laid') && norm.includes('off') && norm.includes('days'));
          });

          // Populate Form A fields from Form 25 data
          mappedData.forEach((row, rowIndex) => {
            const empItem = employees[rowIndex];
            const emp = empItem && (empItem.Employee || empItem.employee || empItem);

            if (!emp) return;

            const empNameCandidates = getEmployeeNameCandidates(emp);
            const empName = empNameCandidates[0] || '';
            const empFirstName = String(
              emp.FirstName ||
              emp['FirstName'] ||
              emp.First_Name ||
              emp['First_Name'] ||
              findValueByNormalizedKey(emp, 'firstname') ||
              ''
            ).trim().toLowerCase();

            const empId = String(
              emp.EmployeeID ||
              emp['EmployeeID'] ||
              emp['Employee ID'] ||
              emp.Zoho_ID ||
              emp['Zoho_ID'] ||
              findValueByNormalizedKey(emp, 'employeeid') ||
              findValueByNormalizedKey(emp, 'employee_id') ||
              ''
            ).trim().toLowerCase();

            let matchedForm25Record = null;
            const matchedAttendanceAgg =
              (empId && attendanceAggByEmployeeKey.get(`id:${empId}`)) ||
              (empName && attendanceAggByEmployeeKey.get(`name:${empName}`)) ||
              (empFirstName && attendanceAggByEmployeeKey.get(`first:${empFirstName}`)) ||
              null;

            if (empNameCandidates.length > 0) {
              matchedForm25Record = empNameCandidates.map((candidate) => form25ByName.get(candidate)).find(Boolean) || null;
            }

            if (!matchedForm25Record && empId && form25ById.has(empId)) {
              matchedForm25Record = form25ById.get(empId);
            }

            const fallbackForm25Record =
              matchedForm25Record ||
              form25Records[rowIndex] ||
              null;

            if (noOfDaysEmployedHeader) {
              if (matchedAttendanceAgg) {
                row[noOfDaysEmployedHeader] = String(matchedAttendanceAgg.days || 0);
              } else if (fallbackForm25Record) {
                row[noOfDaysEmployedHeader] = pickForm25DayValue(fallbackForm25Record, [
                  'totalDaysWorked',
                  'Total Worked Days',
                  'total_worked_days'
                ]);
              }
            }
            if (noOfDaysLaidOffHeader) {
              if (matchedAttendanceAgg) {
                const daysInMonth = new Date(new Date().getFullYear(), (MONTH_NAMES.findIndex(
                  (month) => month.toLowerCase().startsWith(String(selectedMonth || '').toLowerCase().trim())
                ) >= 0 ? MONTH_NAMES.findIndex(
                  (month) => month.toLowerCase().startsWith(String(selectedMonth || '').toLowerCase().trim())
                ) : new Date().getMonth()) + 1, 0).getDate();
                row[noOfDaysLaidOffHeader] = String(Math.max(0, daysInMonth - Number(matchedAttendanceAgg.days || 0)));
              } else if (fallbackForm25Record) {
                row[noOfDaysLaidOffHeader] = pickForm25DayValue(fallbackForm25Record, [
                  'numberOfDaysOnLossOfPay',
                  'No of days on loss of pay',
                  'loss_of_pay',
                  'lopDays'
                ]);
              }
            }

            if (natureOfWorkHeader) {
              row[natureOfWorkHeader] = String(
                emp.Designation ||
                emp.designation ||
                emp['Designation.displayValue'] ||
                findValueByNormalizedKey(emp, 'designation') ||
                ''
              );
            }

            if (dateOnWhichEmployedHeader) {
              row[dateOnWhichEmployedHeader] = String(
                emp.Dateofjoining ||
                emp['Dateofjoining'] ||
                emp.DateofJoining ||
                emp['Date of Joining'] ||
                emp['DateofJoining'] ||
                emp.dateOfJoining ||
                emp['dateOfJoining'] ||
                emp['Date of joining'] ||
                emp.Date_of_Joining ||
                emp['Date_of_Joining'] ||
                findValueByNormalizedKey(emp, 'dateofjoining') ||
                findValueByNormalizedKey(emp, 'dateofjoiningdisplayvalue') ||
                ''
              );
            }
          });

          console.log('Form 25 data populated for Form A fields');
        } catch (form25Err) {
          console.warn('Form 25 data fetch for Form A skipped:', form25Err?.message || form25Err);
        }
      }
     
      // Check if name field was populated
      const nameHeader = currentHeaders.find(h =>
        h.toLowerCase().includes('name') &&
        (h.toLowerCase().includes('employee') || h.toLowerCase().includes('emp'))
      );
      if (nameHeader) {
        const namesPopulated = mappedData.filter(row => row[nameHeader] && row[nameHeader].trim()).length;
        console.log(`Name field "${nameHeader}" populated for ${namesPopulated} out of ${mappedData.length} employees`);
        if (namesPopulated === 0) {
          console.warn('⚠️ No employee names were populated! Check field mapping.');
          console.log('Sample employee object keys:', Object.keys(employees[0] || {}));
        }
      }
     
      // Fetch leave data and populate "Leave Balance at end of the Month" column
      let leaveBalancePopulated = 0;
      let hasLeaveBalanceColumn = false;
      let leaveEarnedPopulated = 0;
      let hasLeaveEarnedColumn = false;
      let leaveAvailedPopulated = 0;
      let hasLeaveAvailedColumn = false;
      let otherLeavePopulated = 0;
      let hasOtherLeaveColumn = false;

      try {
        console.log('Fetching leave data...');
        const { fromDate, toDate } = zohoBookedBalanceRangeForUiMonth(selectedMonth);
        const unit = 'Day';
        const leaveUrl = `/server/leavedata_function?from=${encodeURIComponent(fromDate)}&to=${encodeURIComponent(toDate)}&unit=${encodeURIComponent(unit)}`;
       
        const leaveResponse = await fetch(leaveUrl);
        const leaveResult = await leaveResponse.json();
        const leaveTypeLabels = (leaveResult && leaveResult.leaveTypeLabels) || {};

        if (leaveResult && typeof leaveResult === 'object') {
          console.log('Leave data received:', leaveResult);

          // Extract leave records. Backend may send a populated `records` map while `leaveRecords` is []
          // (normalizeLeaveResponse does not always match Zoho bookedAndBalance shape). Empty array must not win.
          let leaveRecords = [];

          if (Array.isArray(leaveResult.leaveRecords) && leaveResult.leaveRecords.length > 0) {
            leaveRecords = leaveResult.leaveRecords;
          } else if (
            leaveResult.records &&
            typeof leaveResult.records === 'object' &&
            !Array.isArray(leaveResult.records) &&
            Object.keys(leaveResult.records).length > 0
          ) {
            leaveRecords = Object.entries(leaveResult.records).map(([employeeId, row]) => ({
              employeeId: String(employeeId),
              ...(row && typeof row === 'object' ? row : {}),
            }));
          } else {
            const leaveData =
              (leaveResult.data && typeof leaveResult.data === 'object' ? leaveResult.data : null) ||
              leaveResult.records ||
              leaveResult;

            if (Array.isArray(leaveData) && leaveData.length > 0) {
              leaveRecords = leaveData;
            } else if (leaveData && leaveData.response && Array.isArray(leaveData.response.result)) {
              leaveRecords = leaveData.response.result;
            } else if (leaveData && leaveData.response && Array.isArray(leaveData.response)) {
              leaveRecords = leaveData.response;
            } else if (leaveData && leaveData.result && Array.isArray(leaveData.result)) {
              leaveRecords = leaveData.result;
            } else if (leaveData && typeof leaveData === 'object' && !Array.isArray(leaveData)) {
              const objectRows = Object.entries(leaveData).map(([employeeId, row]) => ({
                employeeId: String(employeeId),
                ...(row && typeof row === 'object' ? row : {}),
              }));
              if (objectRows.length > 0) leaveRecords = objectRows;
            }
          }

          console.log('Extracted leave records:', leaveRecords);

          const normalizedCurrentHeaders = Array.isArray(currentHeaders) ? currentHeaders : [];
          const parsedGroupLabels =
            (Array.isArray(options.columnGroupLabels) &&
              options.columnGroupLabels.length === normalizedCurrentHeaders.length &&
              options.columnGroupLabels) ||
            (Array.isArray(formFileModalData?.parsedColumnGroupLabels) &&
              formFileModalData.parsedColumnGroupLabels.length === normalizedCurrentHeaders.length &&
              formFileModalData.parsedColumnGroupLabels) ||
            null;

          const formXLeaveSections = resolveFormXLeaveSectionHeaders(
            normalizedCurrentHeaders,
            parsedGroupLabels
          );
          const leaveEarnedHeader = formXLeaveSections.earnedEarned;
          const leaveAvailedHeader = formXLeaveSections.earnedAvailed;
          const otherLeaveHeaderMap = {
            beginning: formXLeaveSections.otherBeginning,
            availed: formXLeaveSections.otherAvailed,
            balance: formXLeaveSections.otherBalance,
          };

          const leaveBeginningHeader = normalizedCurrentHeaders.find((header) => {
            const headerLower = header.toLowerCase().trim();
            return (
              headerLower.includes('leave') &&
              headerLower.includes('beginning') &&
              headerLower.includes('month')
            );
          });

          const leaveBalanceHeader = normalizedCurrentHeaders.find((header) => {
            const headerLower = header.toLowerCase().trim();
            return (
              headerLower.includes('leave') &&
              headerLower.includes('balance') &&
              headerLower.includes('end') &&
              headerLower.includes('month')
            );
          });

          if (leaveBeginningHeader || leaveBalanceHeader) {
            hasLeaveBalanceColumn = true;
          }

          if (leaveEarnedHeader) {
            hasLeaveEarnedColumn = true;
          }

          if (leaveAvailedHeader) {
            hasLeaveAvailedColumn = true;
          }

          hasOtherLeaveColumn = !!(
            otherLeaveHeaderMap.beginning ||
            otherLeaveHeaderMap.availed ||
            otherLeaveHeaderMap.balance
          );

          const hasAnyLeaveTargetColumn =
            leaveBeginningHeader ||
            leaveBalanceHeader ||
            leaveEarnedHeader ||
            leaveAvailedHeader ||
            hasOtherLeaveColumn;

          if (hasAnyLeaveTargetColumn && leaveRecords.length > 0) {
            if (leaveBeginningHeader) {
              console.log(`Found leave beginning column: "${leaveBeginningHeader}"`);
            }
            if (leaveBalanceHeader) {
              console.log(`Found leave balance column: "${leaveBalanceHeader}"`);
            }
            if (leaveEarnedHeader) {
              console.log(`Found leave earned column: "${leaveEarnedHeader}"`);
            }
            if (leaveAvailedHeader) {
              console.log(`Found leave availed column: "${leaveAvailedHeader}"`);
            }
           
            const tryParseLeaveCell = (value) => {
              if (value == null || value === '') return null;
              if (typeof value === 'object') return value;
              const raw = String(value).trim();
              if (!raw || raw === '{}') return null;
              try {
                const parsed = JSON.parse(raw);
                return parsed && typeof parsed === 'object' ? parsed : null;
              } catch (_) {
                return null;
              }
            };

            const toFiniteNumber = (value) => {
              const n = Number(value);
              return Number.isFinite(n) ? n : 0;
            };

            const getLeaveRowAllBalance = (leaveRecord) => {
              let total = 0;
              Object.entries(leaveRecord || {}).forEach(([key, value]) => {
                const keyLower = String(key || '').toLowerCase();
                if (keyLower === 's.no' || keyLower === 'sno' || keyLower === 'employeeid' || keyLower === 'employee_id') return;
                const parsed = tryParseLeaveCell(value);
                if (!parsed) return;
                total += toFiniteNumber(parsed.balance);
              });
              return String(total);
            };

            const earnedLeaveTestKeyMatches = (keyOrLabel) => {
              const kl = String(keyOrLabel || '')
                .toLowerCase()
                .trim();
              return (
                kl === 'earned leave (test)' ||
                kl === 'earned leave(test)' ||
                (kl.includes('earned') && kl.includes('test'))
              );
            };

            const getEarnedLeaveTestCell = (leaveRecord) => {
              if (!leaveRecord || typeof leaveRecord !== 'object') return null;
              for (const k of Object.keys(leaveRecord)) {
                if (earnedLeaveTestKeyMatches(k)) return leaveRecord[k];
                const lab = leaveTypeLabels[k];
                if (lab && earnedLeaveTestKeyMatches(lab)) return leaveRecord[k];
              }
              return null;
            };

            const getEarnedLeaveTestBalanceBooked = (leaveRecord) => {
              const raw = getEarnedLeaveTestCell(leaveRecord);
              const parsed = tryParseLeaveCell(raw);
              if (!parsed || typeof parsed !== 'object') return { balance: '', booked: '' };
              const bal = parsed.paidBalance ?? parsed.balance;
              const book = parsed.paidBooked ?? parsed.booked;
              return {
                balance: bal != null && bal !== '' ? String(bal) : '',
                booked: book != null && book !== '' ? String(book) : '',
              };
            };

            if (otherLeaveHeaderMap.beginning) {
              console.log(`Found other leave beginning column: "${otherLeaveHeaderMap.beginning}"`);
            }
            if (otherLeaveHeaderMap.availed) {
              console.log(`Found other leave availed column: "${otherLeaveHeaderMap.availed}"`);
            }
            if (otherLeaveHeaderMap.balance) {
              console.log(`Found other leave balance column: "${otherLeaveHeaderMap.balance}"`);
            }

            const collectLeaveApiIdentityKeys = (leaveRecord) => {
              const names = [];
              const ids = [];
              const addName = (v) => {
                if (v == null || !String(v).trim()) return;
                names.push(String(v).toLowerCase().trim());
              };
              const addId = (v) => {
                if (v == null || !String(v).trim()) return;
                ids.push(String(v).toLowerCase().trim());
              };
              const emp = leaveRecord.employee;
              if (emp && typeof emp === 'object') {
                addName(emp.name);
                addId(emp.id);
              }
              addName(
                leaveRecord.EmployeeName ||
                  leaveRecord.employeeName ||
                  leaveRecord['Employee Name'] ||
                  leaveRecord.name
              );
              addId(
                leaveRecord.EmployeeID ||
                  leaveRecord.Employee_ID ||
                  leaveRecord.employee_id ||
                  leaveRecord['Employee ID'] ||
                  leaveRecord.employeeId ||
                  leaveRecord.ZohoID ||
                  leaveRecord['Zoho.ID']
              );
              return { names: [...new Set(names)], ids: [...new Set(ids)] };
            };

            // Create maps of leave data by employee identifier
            const leaveBeginningMap = new Map();
            const leaveBalanceMap = new Map();
            const leaveEarnedMap = new Map();
            const leaveAvailedMap = new Map();
            const otherLeaveMaps = {
              beginning: new Map(),
              availed: new Map(),
              balance: new Map()
            };
           
            leaveRecords.forEach((leaveRecord) => {
              const { names: nameKeys, ids: idKeys } = collectLeaveApiIdentityKeys(leaveRecord);
              const { balance: earnedBalanceStr, booked: earnedBookedStr } = getEarnedLeaveTestBalanceBooked(leaveRecord);
              const balance = getLeaveRowAllBalance(leaveRecord);
              const otherLeaveMetrics = getCombinedOtherLeaveMetrics(leaveRecord, leaveTypeLabels);

              const setIdentityLeaveMaps = (identityKey) => {
                if (!identityKey) return;
                if (otherLeaveMetrics.beginning !== '') otherLeaveMaps.beginning.set(identityKey, otherLeaveMetrics.beginning);
                if (otherLeaveMetrics.availed !== '') otherLeaveMaps.availed.set(identityKey, otherLeaveMetrics.availed);
                if (otherLeaveMetrics.balance !== '') otherLeaveMaps.balance.set(identityKey, otherLeaveMetrics.balance);
              };

              nameKeys.forEach((nameKey) => {
                if (leaveBeginningHeader) {
                  leaveBeginningMap.set(nameKey, balance);
                }
                if (leaveBalanceHeader) {
                  leaveBalanceMap.set(nameKey, balance);
                }
                if (leaveEarnedHeader && earnedBalanceStr !== '') {
                  leaveEarnedMap.set(nameKey, earnedBalanceStr);
                }
                if (leaveAvailedHeader && earnedBookedStr !== '') {
                  leaveAvailedMap.set(nameKey, earnedBookedStr);
                }
                setIdentityLeaveMaps(nameKey);
              });
              idKeys.forEach((idKey) => {
                if (leaveBeginningHeader) {
                  leaveBeginningMap.set(idKey, balance);
                }
                if (leaveBalanceHeader) {
                  leaveBalanceMap.set(idKey, balance);
                }
                if (leaveEarnedHeader && earnedBalanceStr !== '') {
                  leaveEarnedMap.set(idKey, earnedBalanceStr);
                }
                if (leaveAvailedHeader && earnedBookedStr !== '') {
                  leaveAvailedMap.set(idKey, earnedBookedStr);
                }
                setIdentityLeaveMaps(idKey);
              });
            });
           
            console.log('Leave beginning map:', Array.from(leaveBeginningMap.entries()));
            console.log('Leave balance map:', Array.from(leaveBalanceMap.entries()));
            console.log('Leave earned map:', Array.from(leaveEarnedMap.entries()));
            console.log('Leave availed map:', Array.from(leaveAvailedMap.entries()));
           
            // Find employee name column to match employees
            const employeeNameHeader = currentHeaders.find(header => {
              const headerLower = header.toLowerCase().trim();
              return headerLower.includes('name') &&
                     (headerLower.includes('employee') || headerLower.includes('emp'));
            });
           
            // Find employee ID column for matching
            const employeeIdHeader = currentHeaders.find(header => {
              const headerLower = header.toLowerCase().trim();
              return (headerLower.includes('identification') ||
                     (headerLower.includes('employee') && headerLower.includes('id')) ||
                     (headerLower.includes('emp') && headerLower.includes('code'))) &&
                     !headerLower.includes('zoho');
            });
           
            // Populate leave values for each employee
            mappedData.forEach((row, index) => {
              const empItem = employees[index];
              const emp = empItem && (empItem.Employee || empItem.employee || empItem);
              const empNameCandidates = getEmployeeNameCandidates(emp || {});
              const empWorkerId = String(getFallbackWorkerId(emp || {}) || '').toLowerCase().trim();
              const empVisibleId = String(getFallbackEmployeeId(emp || {}) || '').toLowerCase().trim();
              let matchedBeginning = null;
              let matchedBalance = null;
              let matchedEarned = null;
              let matchedAvailed = null;
              const matchedOtherLeave = { beginning: null, availed: null, balance: null };

              // Try to match by employee name first
              const rowEmployeeName = employeeNameHeader && row[employeeNameHeader]
                ? String(row[employeeNameHeader]).toLowerCase().trim()
                : '';
              const allNameCandidates = Array.from(new Set([rowEmployeeName, ...empNameCandidates].filter(Boolean)));
              const tryMatchOtherLeaveValues = (identityKey) => {
                if (matchedOtherLeave.beginning == null && otherLeaveMaps.beginning.has(identityKey)) {
                  matchedOtherLeave.beginning = otherLeaveMaps.beginning.get(identityKey);
                }
                if (matchedOtherLeave.availed == null && otherLeaveMaps.availed.has(identityKey)) {
                  matchedOtherLeave.availed = otherLeaveMaps.availed.get(identityKey);
                }
                if (matchedOtherLeave.balance == null && otherLeaveMaps.balance.has(identityKey)) {
                  matchedOtherLeave.balance = otherLeaveMaps.balance.get(identityKey);
                }
              };
              for (const employeeName of allNameCandidates) {
                if (leaveBeginningHeader && matchedBeginning == null && leaveBeginningMap.has(employeeName)) {
                  matchedBeginning = leaveBeginningMap.get(employeeName);
                }
                if (leaveBalanceHeader && matchedBalance == null && leaveBalanceMap.has(employeeName)) {
                  matchedBalance = leaveBalanceMap.get(employeeName);
                }
                if (leaveEarnedHeader && matchedEarned == null && leaveEarnedMap.has(employeeName)) {
                  matchedEarned = leaveEarnedMap.get(employeeName);
                }
                if (leaveAvailedHeader && matchedAvailed == null && leaveAvailedMap.has(employeeName)) {
                  matchedAvailed = leaveAvailedMap.get(employeeName);
                }
                tryMatchOtherLeaveValues(employeeName);
              }

              // If not matched by name, try by employee ID
              if ((!matchedBeginning && leaveBeginningHeader) || (!matchedBalance && leaveBalanceHeader) || (!matchedEarned && leaveEarnedHeader) || (!matchedAvailed && leaveAvailedHeader)) {
                const candidateIds = [
                  employeeIdHeader && row[employeeIdHeader] ? String(row[employeeIdHeader]).toLowerCase().trim() : '',
                  empVisibleId,
                  empWorkerId
                ].filter(Boolean);
                for (const employeeId of candidateIds) {
                  if (leaveBeginningHeader && !matchedBeginning && leaveBeginningMap.has(employeeId)) {
                    matchedBeginning = leaveBeginningMap.get(employeeId);
                  }
                  if (leaveBalanceHeader && !matchedBalance && leaveBalanceMap.has(employeeId)) {
                    matchedBalance = leaveBalanceMap.get(employeeId);
                  }
                  if (leaveEarnedHeader && !matchedEarned && leaveEarnedMap.has(employeeId)) {
                    matchedEarned = leaveEarnedMap.get(employeeId);
                  }
                  if (leaveAvailedHeader && !matchedAvailed && leaveAvailedMap.has(employeeId)) {
                    matchedAvailed = leaveAvailedMap.get(employeeId);
                  }
                  tryMatchOtherLeaveValues(employeeId);
                }
              }

              // If still not matched, try to find by partial name match
              if (allNameCandidates.length > 0) {
                const employeeName = allNameCandidates[0];
                if (leaveBeginningHeader && !matchedBeginning) {
                  for (const [key, value] of leaveBeginningMap.entries()) {
                    if (key.includes(employeeName) || employeeName.includes(key)) {
                      matchedBeginning = value;
                      break;
                    }
                  }
                }
                if (leaveBalanceHeader && !matchedBalance) {
                  for (const [key, value] of leaveBalanceMap.entries()) {
                    if (key.includes(employeeName) || employeeName.includes(key)) {
                      matchedBalance = value;
                      break;
                    }
                  }
                }
                if (leaveEarnedHeader && !matchedEarned) {
                  for (const [key, value] of leaveEarnedMap.entries()) {
                    if (key.includes(employeeName) || employeeName.includes(key)) {
                      matchedEarned = value;
                      break;
                    }
                  }
                }
                if (leaveAvailedHeader && !matchedAvailed) {
                  for (const [key, value] of leaveAvailedMap.entries()) {
                    if (key.includes(employeeName) || employeeName.includes(key)) {
                      matchedAvailed = value;
                      break;
                    }
                  }
                }
                if (matchedOtherLeave.beginning == null) {
                  for (const [key, value] of otherLeaveMaps.beginning.entries()) {
                    if (key.includes(employeeName) || employeeName.includes(key)) {
                      matchedOtherLeave.beginning = value;
                      break;
                    }
                  }
                }
                if (matchedOtherLeave.availed == null) {
                  for (const [key, value] of otherLeaveMaps.availed.entries()) {
                    if (key.includes(employeeName) || employeeName.includes(key)) {
                      matchedOtherLeave.availed = value;
                      break;
                    }
                  }
                }
                if (matchedOtherLeave.balance == null) {
                  for (const [key, value] of otherLeaveMaps.balance.entries()) {
                    if (key.includes(employeeName) || employeeName.includes(key)) {
                      matchedOtherLeave.balance = value;
                      break;
                    }
                  }
                }
              }

              if (matchedBeginning !== null && leaveBeginningHeader) {
                row[leaveBeginningHeader] = String(matchedBeginning);
                leaveBalancePopulated++;
              }

              // Populate the balance if found
              if (matchedBalance !== null && leaveBalanceHeader) {
                row[leaveBalanceHeader] = String(matchedBalance);
                leaveBalancePopulated++;
                console.log(`✓ Populated leave balance for employee ${index + 1}: ${matchedBalance}`);
              }

              // Earned Leave (Test) paidBalance → "Leave earned during the Period"
              if (matchedEarned != null && leaveEarnedHeader) {
                row[leaveEarnedHeader] = String(matchedEarned);
                leaveEarnedPopulated++;
                console.log(`✓ Populated leave earned (Earned Leave Test balance) for employee ${index + 1}: ${matchedEarned}`);
              }

              // Earned Leave (Test) paidBooked → "Leave availed during the Month"
              if (matchedAvailed != null && leaveAvailedHeader) {
                row[leaveAvailedHeader] = String(matchedAvailed);
                leaveAvailedPopulated++;
                console.log(`✓ Populated leave availed (Earned Leave Test booked) for employee ${index + 1}: ${matchedAvailed}`);
              }
              if (otherLeaveHeaderMap.beginning && matchedOtherLeave.beginning != null) {
                row[otherLeaveHeaderMap.beginning] = String(matchedOtherLeave.beginning);
                otherLeavePopulated++;
              }
              if (otherLeaveHeaderMap.availed && matchedOtherLeave.availed != null) {
                row[otherLeaveHeaderMap.availed] = String(matchedOtherLeave.availed);
                otherLeavePopulated++;
              }
              if (otherLeaveHeaderMap.balance && matchedOtherLeave.balance != null) {
                row[otherLeaveHeaderMap.balance] = String(matchedOtherLeave.balance);
                otherLeavePopulated++;
              }
            });

            console.log(`Leave balance populated for ${leaveBalancePopulated} out of ${mappedData.length} employees`);
            console.log(`Leave earned (Earned Leave Test balance) populated for ${leaveEarnedPopulated} out of ${mappedData.length} employees`);
            console.log(`Leave availed (Earned Leave Test booked) populated for ${leaveAvailedPopulated} out of ${mappedData.length} employees`);
            console.log(`Other leave (Privilege/LWP/Sick/Wedding totals) populated for ${otherLeavePopulated} cell(s)`);
          } else {
            if (!leaveBalanceHeader && !leaveEarnedHeader && !leaveAvailedHeader && !hasOtherLeaveColumn) {
              console.warn('⚠️ No leave columns matched (earned / other leave / balance at end).');
            }
            if (leaveRecords.length === 0) {
              console.warn('⚠️ No leave records found in response');
            }
          }
        } else {
          console.warn('⚠️ Failed to fetch leave data or no data received');
        }
      } catch (leaveErr) {
        console.error('Error fetching leave data:', leaveErr);
        // Don't throw error, just log it - employee data is still populated
      }
     
      // Update form table data with employee data (including leave balance and leave earned)
      if (!returnMappedData) setFormTableData(mappedData);
      let successMessage = `Successfully loaded ${mappedData.length} employee records`;
      const populatedFields = [];
      if (hasTotalHoursColumn && attendanceHoursPopulated > 0) {
        populatedFields.push(`total hours (${attendanceHoursPopulated} populated)`);
      }
      if (hasDailyHoursColumn && attendanceHoursPopulated > 0) {
        populatedFields.push(`daily worked hours (${attendanceHoursPopulated} populated)`);
      }
      if (hasLeaveBalanceColumn && leaveBalancePopulated > 0) {
        populatedFields.push(`leave balance (${leaveBalancePopulated} populated)`);
      }
      if (hasLeaveEarnedColumn && leaveEarnedPopulated > 0) {
        populatedFields.push(`leave earned (${leaveEarnedPopulated} populated)`);
      }
      if (hasLeaveAvailedColumn && leaveAvailedPopulated > 0) {
        populatedFields.push(`leave availed (${leaveAvailedPopulated} populated)`);
      }
      if (hasOtherLeaveColumn && otherLeavePopulated > 0) {
        populatedFields.push(`other leave (${otherLeavePopulated} populated)`);
      }
      if (populatedFields.length > 0) {
        successMessage += ` with ${populatedFields.join(' and ')}`;
      }
      if (!returnMappedData) {
        setSuccess(successMessage);
        setTimeout(() => setSuccess(''), 5000);
      }
      if (returnMappedData) return mappedData;
     
    } catch (err) {
      console.error('Error fetching employee data:', err);
      console.error('Error stack:', err.stack);
      if (!returnMappedData) setError(err.message || 'Failed to fetch employee data from Zoho People');
      if (returnMappedData) throw err;
    } finally {
      if (!returnMappedData) setFormFileLoading(false);
    }
  };

  // Parse Excel file similar to SEMaster (same logic so Form V, Form U, W, X display correct format)
  const parseExcelForm = (workbook) => {
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
    const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1:A1');
    const maxSheetCols = (range.e && typeof range.e.c === 'number') ? (range.e.c + 1) : 0;
    const merges = worksheet['!merges'] || [];
    // !ref often ends at the last stored cell; merged Form W headers can extend far beyond (e.g. Deductions → column Z).
    let mergeMaxCol = 0;
    for (let mi = 0; mi < merges.length; mi++) {
      const m = merges[mi];
      if (m && m.e && typeof m.e.c === 'number') mergeMaxCol = Math.max(mergeMaxCol, m.e.c + 1);
    }
    const effectiveSheetCols = Math.max(maxSheetCols, mergeMaxCol);

    const getRawCellText = (r, c) => {
      if (r < 0 || c < 0) return '';
      const ref = XLSX.utils.encode_cell({ r, c });
      const cell = worksheet[ref];
      if (!cell || cell.v == null) return '';
      return String(cell.v).trim();
    };

    const getMergedAwareCellText = (r, c) => {
      const direct = getRawCellText(r, c);
      if (direct) return direct;
      for (let i = 0; i < merges.length; i++) {
        const m = merges[i];
        if (!m || !m.s || !m.e) continue;
        if (r >= m.s.r && r <= m.e.r && c >= m.s.c && c <= m.e.c) {
          const topLeft = getRawCellText(m.s.r, m.s.c);
          if (topLeft) return topLeft;
        }
      }
      return '';
    };
   
    let headerRowIndex = -1;
    let headers = [];
    let subColumnRowIndex = -1;
    let subColumnsData = {};
   
    // Table header keywords (Form U, V, W, X and similar statutory forms)
    const tableHeaderKeywords = [
      's.no', 's.no.', 'serial', 'serial number',
      'name of the employee', 'employee name', 'name of employee',
      'employee identification', 'identification no', 'identification number',
      'gender', 'designation', 'father', 'spouse', 'father / spouse',
      'date of birth', 'date of joining', 'birth', 'joining', 'commencement',
      'name of the manager', 'manager', 'incharge', 'in charge',
      'basic wage', 'dearness allowance', 'house rent allowance',
      'overtime wages', 'gross wages', 'net wages', 'deductions',
      'number of days', 'days worked', 'provident fund', 'signature', 'remarks',
      'number in register', 'muster', 'exempted', 'extent of overtime', 'normal hours',
      'normal rate of pay', 'overtime rate', 'department', 'date of which overtime',
      // LWF Form C and similar (merged quarterly columns)
      'quarter', 'fines', 'unpaid', 'accumulations'
    ];
   
    // First pass: strongly prefer the real table row that contains S.No/Serial.
    // This avoids selecting top metadata rows like "Name and Address of the Employer:" as headers.
    let bestSNoHeaderRow = -1;
    let bestSNoHeaderScore = -1;
    for (let rowIdx = 0; rowIdx < Math.min(40, jsonData.length); rowIdx++) {
      const mergedRow = [];
      for (let colIdx = 0; colIdx < Math.max(effectiveSheetCols, 20); colIdx++) {
        mergedRow.push(String(getMergedAwareCellText(rowIdx, colIdx) || '').trim());
      }
      const nonEmpty = mergedRow.filter(Boolean);
      if (nonEmpty.length === 0) continue;
      const rowTextLower = nonEmpty.join(' ').toLowerCase();
      const hasSNo =
        rowTextLower.includes('s.no') ||
        rowTextLower.includes('s. no') ||
        rowTextLower.includes('serial') ||
        /number\s+in\s+register|register\s+number/i.test(rowTextLower);
      if (!hasSNo) continue;

      const keywordHits = [
        'employee',
        'wage',
        'allowance',
        'overtime',
        'deduction',
        'provident',
        'insurance',
        'gross',
        'net',
        'fund'
      ].reduce((acc, kw) => (rowTextLower.includes(kw) ? acc + 1 : acc), 0);

      // Prefer rows that look like a wide table header, not a small label row.
      const score = (nonEmpty.length * 10) + (keywordHits * 30);
      if (score > bestSNoHeaderScore) {
        bestSNoHeaderScore = score;
        bestSNoHeaderRow = rowIdx;
      }
    }

    if (bestSNoHeaderRow !== -1) {
      headerRowIndex = bestSNoHeaderRow;
      headers = [];
      const maxCols = Math.max(
        effectiveSheetCols,
        (jsonData[bestSNoHeaderRow] || []).length,
        20
      );
      const isDeductionsBanner = (txt) => {
        const t = String(txt || '')
          .replace(/\r?\n/g, ' ')
          .replace(/\s+/g, ' ')
          .trim()
          .toLowerCase();
        return t === 'deduction' || t === 'deductions';
      };
      const isWagePeriodish = (txt) => /wage\s*period/i.test(String(txt || ''));
      for (let colIdx = 0; colIdx < maxCols; colIdx++) {
        let h = String(getMergedAwareCellText(bestSNoHeaderRow, colIdx) || '').trim();
        if (!h) {
          for (let lookBack = 1; lookBack <= 4; lookBack++) {
            const r = bestSNoHeaderRow - lookBack;
            if (r < 0) break;
            const cand = String(getMergedAwareCellText(r, colIdx) || '').trim();
            if (!cand || isDeductionsBanner(cand) || isWagePeriodish(cand)) continue;
            if (/^\d+$/.test(cand)) continue;
            if (cand.length < 2) continue;
            h = cand;
            break;
          }
        }
        if (!h) {
          for (let lookFwd = 1; lookFwd <= 2; lookFwd++) {
            const cand = String(getMergedAwareCellText(bestSNoHeaderRow + lookFwd, colIdx) || '').trim();
            if (!cand || /^\d+$/.test(cand) || cand.length < 2) continue;
            h = cand;
            break;
          }
        }
        headers.push(h);
      }
      while (headers.length > 0 && !headers[headers.length - 1]) headers.pop();
    }

    for (let rowIdx = 0; rowIdx < Math.min(25, jsonData.length) && headerRowIndex === -1; rowIdx++) {
      const row = jsonData[rowIdx] || [];
      if (row.length === 0) continue;
     
      const rowText = row.map(cell => String(cell || '').toLowerCase().trim()).join(' ');
      let matchingHeaderCells = 0;
      for (let cellIdx = 0; cellIdx < row.length; cellIdx++) {
        const cellText = String(row[cellIdx] || '').toLowerCase().trim();
        if (cellText && tableHeaderKeywords.some(kw => cellText.includes(kw))) matchingHeaderCells++;
      }
      const textCellCount = row.filter(cell => {
        const str = String(cell || '').trim();
        return str && str.length > 0 && !str.match(/^\d+$/);
      }).length;
      const isSingleFieldRow = textCellCount === 1 && row.length <= 2;
      const hasSNo = rowText.includes('s.no') || rowText.includes('s.no.') || rowText.includes('serial');
      const isTableHeaderRow = hasSNo ||
        (matchingHeaderCells >= 2 && textCellCount >= 3) ||
        (textCellCount >= 4 && !isSingleFieldRow);
     
      if (isTableHeaderRow) {
        if (hasSNo && rowIdx + 1 < jsonData.length) {
          const nextRow = jsonData[rowIdx + 1] || [];
          const nextRowText = nextRow.map(cell => String(cell || '').toLowerCase().trim()).join(' ');
          const nextRowHasSNo = nextRowText.includes('s.no') || nextRowText.includes('s.no.') || nextRowText.includes('serial');
          const nextRowTextCount = nextRow.filter(cell => {
            const str = String(cell || '').trim();
            return str && str.length > 0 && !str.match(/^\d+$/);
          }).length;
          if (nextRowHasSNo && nextRowTextCount > textCellCount) continue;
        }
        headerRowIndex = rowIdx;
        const nextRows = jsonData.slice(rowIdx + 1, Math.min(rowIdx + 10, jsonData.length));
        const nextLengths = nextRows.map(function (r) { return (r && r.length) ? r.length : 0; });
        const maxCols = Math.max.apply(Math, [row.length].concat(nextLengths));
        const columnsWithData = new Set();
        for (let checkRow = rowIdx + 1; checkRow < Math.min(rowIdx + 20, jsonData.length); checkRow++) {
          const dataRow = jsonData[checkRow] || [];
          for (let colIdx = 0; colIdx < dataRow.length; colIdx++) {
            if (dataRow[colIdx] !== undefined && dataRow[colIdx] !== null && String(dataRow[colIdx]).trim()) columnsWithData.add(colIdx);
          }
        }
        headers = [];
        for (let colIdx = 0; colIdx < maxCols; colIdx++) {
          let header = getMergedAwareCellText(rowIdx, colIdx) || (row[colIdx] ? String(row[colIdx]).trim() : '');
          // Prefer nearby descriptive text over placeholder blanks
          if (!header) {
            for (let lookAhead = 1; lookAhead <= 2; lookAhead++) {
              const candidate = getMergedAwareCellText(rowIdx + lookAhead, colIdx);
              if (!candidate) continue;
              if (/^\d+$/.test(candidate)) continue;
              if (candidate.length < 2) continue;
              header = candidate;
              break;
            }
          }
          if (header && header.endsWith(':') && header.length > 20) {
            let hasOtherTableHeaders = false;
            for (let c = 0; c < row.length; c++) {
              if (c === colIdx) continue;
              const other = row[c] ? String(row[c]).trim() : '';
              if (other && (other.toLowerCase().includes('s.no') || other.toLowerCase().includes('name') || other.toLowerCase().includes('employee') || other.toLowerCase().includes('designation'))) {
                hasOtherTableHeaders = true;
                break;
              }
            }
            if (!hasOtherTableHeaders) continue;
          }
          if (header) headers.push(header);
          else if (columnsWithData.has(colIdx)) headers.push(`Column ${String.fromCharCode(65 + colIdx)}`);
          else headers.push('');
        }
        while (headers.length > 0 && !headers[headers.length - 1] && !columnsWithData.has(headers.length - 1)) headers.pop();
        break;
      }
    }
   
    // Fallback: if no keyword-based header found, use first row with 3+ short column-like cells (not a single long label)
    if (headerRowIndex === -1) {
      for (let rowIdx = 0; rowIdx < Math.min(25, jsonData.length); rowIdx++) {
        const row = jsonData[rowIdx] || [];
        const cells = row.map(cell => String(cell || '').trim()).filter(Boolean);
        const allShort = cells.length >= 3 && cells.every(c => c.length < 50);
        const notAllLabels = !cells.every(c => c.endsWith(':'));
        if (cells.length >= 3 && allShort && notAllLabels) {
          headerRowIndex = rowIdx;
          headers = row.map((h) => String(h || '').trim());
          break;
        }
      }
    }
   
    // If headers are mostly generic placeholders (Column B, Column C...) rebuild from worksheet rows
    // so Form W and similar templates display real Excel headings in the modal.
    const genericHeaderRegex = /^column\s+([a-z]+|\d+)$/i;
    const rebuildHeadersFromWorksheetRows = () => {
      if (headerRowIndex < 0) return;
      const rebuilt = [];
      const secondHeaderRowExists = headerRowIndex + 1 < jsonData.length;
      const maxColsFromRows = Math.max(
        headers.length,
        (jsonData[headerRowIndex] || []).length,
        secondHeaderRowExists ? (jsonData[headerRowIndex + 1] || []).length : 0,
        effectiveSheetCols
      );

      const normalizeHeaderText = (txt) =>
        String(txt || '')
          .replace(/\r?\n/g, ' ')
          .replace(/\s+/g, ' ')
          .trim();

      for (let colIdx = 0; colIdx < maxColsFromRows; colIdx++) {
        const mainHeader = normalizeHeaderText(getMergedAwareCellText(headerRowIndex, colIdx));
        const subHeader = secondHeaderRowExists
          ? normalizeHeaderText(getMergedAwareCellText(headerRowIndex + 1, colIdx))
          : '';

        let combined = mainHeader;
        const subIsUseful =
          subHeader &&
          !/^\d+$/.test(subHeader) &&
          subHeader.toLowerCase() !== mainHeader.toLowerCase() &&
          !subHeader.endsWith(':');

        if (subIsUseful && mainHeader) combined = `${mainHeader} ${subHeader}`.replace(/\s+/g, ' ').trim();
        else if (!combined && subIsUseful) combined = subHeader;

        rebuilt.push(combined);
      }

      while (rebuilt.length > 0 && !rebuilt[rebuilt.length - 1]) rebuilt.pop();
      if (rebuilt.some((h) => h && !genericHeaderRegex.test(h))) {
        headers = rebuilt;
      }
    };

    const genericHeadersCountBeforeRebuild = headers.filter((h) => genericHeaderRegex.test(String(h || '').trim())).length;
    const shouldRebuildHeaders =
      headers.length > 0 && genericHeadersCountBeforeRebuild >= Math.max(3, Math.ceil(headers.length * 0.5));
    if (shouldRebuildHeaders) rebuildHeadersFromWorksheetRows();

    // Detect subcolumns (e.g. Form V: "Daily Hours of work..." with subcolumns 1, 2, ... 18)
    let expandedHeaders = headers.length ? [...headers] : [];
    if (headerRowIndex >= 0 && headerRowIndex + 1 < jsonData.length) {
      const subRow = jsonData[headerRowIndex + 1] || [];
      const subCells = subRow.map((cell, idx) => {
        const mergedAware = getMergedAwareCellText(headerRowIndex + 1, idx);
        return String(mergedAware || cell || '').trim();
      });
      let subColumnCount = 0;
      let totalSubCells = 0;
      let descriptiveSubCells = 0;
      let dayNumberHits = 0;
      subCells.forEach((cell) => {
        if (cell) {
          totalSubCells++;
          const isNum = /^\d+$/.test(cell);
          const isDayNum = /^(\d{1,2})$/.test(cell) && (() => {
            const n = parseInt(cell, 10);
            return n >= 1 && n <= 31;
          })();
          if (isDayNum) dayNumberHits++;
          const isUsefulLabel = !isNum && !cell.endsWith(':') && cell.length > 1;
          if (isUsefulLabel) subColumnCount++;
          if (!isNum && !cell.endsWith(':') && cell.length > 1) descriptiveSubCells++;
        }
      });
      // Form V: row below main header is day numbers 1..30 under "Daily Hours…" — numbers are valid subcolumns.
      const numericDaySubHeaderRow =
        totalSubCells >= 8 &&
        dayNumberHits >= 8 &&
        dayNumberHits >= Math.max(5, Math.floor(totalSubCells * 0.25));
      // Only treat row+1 as a sub-header row when at least one main header spans
      // multiple adjacent columns (merged parent like Form D/Form V bands).
      let hasMergedLikeMainBand = false;
      if (headers.length > 1) {
        for (let colIdx = 0; colIdx < Math.max(headers.length, subCells.length, effectiveSheetCols) - 1; colIdx++) {
          const m1 = String(getMergedAwareCellText(headerRowIndex, colIdx) || headers[colIdx] || '').trim().toLowerCase();
          const m2 = String(getMergedAwareCellText(headerRowIndex, colIdx + 1) || headers[colIdx + 1] || '').trim().toLowerCase();
          if (!m1 || !m2) continue;
          if (m1 !== m2) continue;
          const s1 = String(subCells[colIdx] || '').trim();
          const s2 = String(subCells[colIdx + 1] || '').trim();
          if (s1 || s2) {
            hasMergedLikeMainBand = true;
            break;
          }
        }
      }
      // Form V: day 1..31 under a merged "Daily hours…" band. Form 10 (Rule 78) has 1..9 in the row
      // under nine distinct column titles — that must not use the day-grid path (no merged main band).
      // Ignore serial-number rows like "1..10"; treat as subcolumn row when descriptive labels exist OR day grid.
      const looksLikeSubColumnRow =
        ((totalSubCells >= 2 && descriptiveSubCells >= 2 && (subColumnCount / totalSubCells) > 0.3) && hasMergedLikeMainBand) ||
        (numericDaySubHeaderRow && hasMergedLikeMainBand);
      if (looksLikeSubColumnRow) {
        let currentMain = null;
        const maxCols = Math.max(headers.length, subCells.length, effectiveSheetCols);
        for (let colIdx = 0; colIdx < maxCols; colIdx++) {
          const mergedMain = getMergedAwareCellText(headerRowIndex, colIdx);
          const mainHeader = (headers[colIdx] || mergedMain || '').trim();
          const subCell = (subCells[colIdx] || '').trim();
          if (mainHeader) {
            currentMain = mainHeader;
            if (!subColumnsData[currentMain]) subColumnsData[currentMain] = [];
          }
          const isDaySub =
            numericDaySubHeaderRow &&
            /^(\d{1,2})$/.test(subCell) &&
            (() => {
              const n = parseInt(subCell, 10);
              return n >= 1 && n <= 31;
            })();
          const validSub =
            subCell &&
            !subCell.includes(':') &&
            currentMain &&
            subCell.toLowerCase() !== currentMain.toLowerCase() &&
            (isDaySub || (!/^\d+$/.test(subCell) && subCell.length > 1));
          if (currentMain && validSub) {
            const arr = subColumnsData[currentMain];
            if (
              isDaySub &&
              /^dates$/i.test(String(currentMain).trim()) &&
              arr.length > 0 &&
              arr[arr.length - 1] === subCell
            ) {
              /* Form 25: merged “DATES” band can surface the same day number twice for adjacent cols */
            } else {
              arr.push(subCell);
            }
          }
        }
        const validMains = Object.keys(subColumnsData).filter((m) => subColumnsData[m].length > 0);
        if (validMains.length > 0) {
          const orderedMains = [];
          let lastMain = null;
          for (let colIdx = 0; colIdx < headers.length; colIdx++) {
            const h = (headers[colIdx] || '').trim();
            if (h) lastMain = h;
            if (lastMain && !orderedMains.includes(lastMain)) orderedMains.push(lastMain);
          }
          if (orderedMains.length === 0) orderedMains.push(...validMains);
          expandedHeaders = [];
          orderedMains.forEach((main) => {
            const subs = subColumnsData[main] || [];
            if (subs.length > 0) subs.forEach((s) => expandedHeaders.push(`${main}_${s}`));
            else expandedHeaders.push(main);
          });
        }
      }
    }
   
    // Extract form header information
    let formHeaderInfo = null;
    let headerFields = [];
    let headerTextRows = [];
   
    if (headerRowIndex > 0) {
      const headerRows = [];
      for (let i = 0; i < headerRowIndex; i++) {
        const row = jsonData[i] || [];
        const rowText = row.map(cell => String(cell || '').trim()).filter(cell => cell).join(' ');
        if (rowText) {
          headerRows.push(row);
        }
      }
     
      if (headerRows.length > 0) {
        let formTitle = '';
        let formSubtitle = '';
        let formReference = '';
        let wagePeriodText = '';
        let form12MergedHeaderCell = null;
        let form10MergedHeaderCell = null;
        let form10MonthEndingCell = null;

        for (let rowIdx = 0; rowIdx < headerRows.length; rowIdx++) {
          const row = headerRows[rowIdx];
          const rowText = row.map(cell => String(cell || '').trim()).join(' ').toUpperCase();
          const fullRowText = row.map(cell => String(cell || '').trim()).join(' ');
          const rowCells = row.map(cell => String(cell || '').trim());
         
          if (isFormVILegendLine(fullRowText)) {
            const compact = fullRowText.replace(/\s+/g, ' ').trim();
            if (compact && !headerTextRows.includes(compact)) headerTextRows.push(compact);
          }

          // Check for FORM No. N (Form 10, etc.) — Excel uses "FORM No. 10", not "FORM - 10"
          // Some templates typo "FORM" as "FROM" (e.g. Form 25).
          if (!formTitle) {
            for (let colIdx = 0; colIdx < row.length; colIdx++) {
              const cell = rowCells[colIdx];
              if (cell && /\b(FORM|FROM)\s+NO\.?\s*\d+/i.test(cell)) {
                const m = cell.match(/\b(FORM|FROM)\s+NO\.?\s*\d+/i);
                if (m) {
                  formTitle = m[0].replace(/\s+/g, ' ').trim();
                  break;
                }
              }
            }
          }
          // Form letter title pattern, e.g. Form " B " (See rule 29)
          if (!formTitle) {
            for (let colIdx = 0; colIdx < row.length; colIdx++) {
              const cell = rowCells[colIdx];
              if (cell && /\bform\s*["']?\s*[a-z]\s*["']?\s*(?:\([^)]*\))?/i.test(cell)) {
                formTitle = cell.replace(/\s+/g, ' ').trim();
                break;
              }
            }
          }
          // Check for FORM - X pattern
          if (rowText.includes('FORM') && rowText.includes('-') && !formTitle) {
            for (let colIdx = 0; colIdx < row.length; colIdx++) {
              const cell = rowCells[colIdx];
              if (cell && cell.toUpperCase().includes('FORM') && cell.includes('-')) {
                formTitle = cell;
                break;
              }
            }
          }
          // Check for subtitle (REGISTER, etc.)
          else if ((rowText.includes('REGISTER') || rowText.includes('EMPLOYEE')) && !formSubtitle && !rowText.includes(':')) {
            for (let colIdx = 0; colIdx < row.length; colIdx++) {
              const cell = rowCells[colIdx];
              if (cell && (cell.toUpperCase().includes('REGISTER') || cell.toUpperCase().includes('EMPLOYEE'))) {
                if (!cell.includes(':')) {
                  formSubtitle = cell;
                  break;
                }
              }
            }
          }
          // Check for subtitle in bracketed legal text, e.g. "(Equal Remuneration rules, 1976)".
          else if (!formReference && /^\(.*\)$/.test(fullRowText.trim()) && /rules?/i.test(fullRowText) && !fullRowText.includes('[') && !fullRowText.includes(':')) {
            formReference = fullRowText.trim();
          }
          // Check for reference [See...]
          else if (fullRowText.includes('[') && fullRowText.includes(']') && !formReference) {
            for (let colIdx = 0; colIdx < row.length; colIdx++) {
              const cell = rowCells[colIdx];
              if (cell && cell.includes('[') && cell.includes(']')) {
                formReference = cell;
                break;
              }
            }
          }
          // Form B style legal line without brackets.
          else if (!formReference && /labou?r\s*welfare\s*fund\s*rules?/i.test(fullRowText)) {
            formReference = fullRowText.replace(/\s+/g, ' ').trim();
          }
          // Check for wage period line (Form W etc.)
          else if (!wagePeriodText && /wage\s*period/i.test(fullRowText) && /\bfrom\b/i.test(fullRowText)) {
            wagePeriodText = fullRowText.replace(/\s+/g, ' ').trim();
          }
          // Form 25 and similar: "For the Period from … To …" (not labeled "wage period")
          else if (
            !wagePeriodText &&
            /\bfor\s+the\s+period\s+from\b/i.test(fullRowText) &&
            /\bto\b/i.test(fullRowText)
          ) {
            wagePeriodText = fullRowText.replace(/\s+/g, ' ').trim();
          }
         
          // Extract header fields (labels with colons)
          for (let colIdx = 0; colIdx < row.length; colIdx++) {
            const cell = rowCells[colIdx];
            if (cell && cell.includes(':')) {
              let label = cell;
              let value = '';
             
              // Look for value in next cells
              for (let nextCol = colIdx + 1; nextCol < Math.min(colIdx + 4, row.length); nextCol++) {
                const nextCell = rowCells[nextCol];
                if (nextCell && nextCell.length > 0 && !nextCell.includes(':')) {
                  value = nextCell;
                  break;
                }
              }
             
              const existingField = headerFields.find(f => {
                const fLabel = f.label.toLowerCase().replace(/[^a-zA-Z0-9]/g, '');
                const newLabel = label.toLowerCase().replace(/[^a-zA-Z0-9]/g, '');
                return fLabel === newLabel;
              });
             
              if (!existingField && label.length > 2) {
                headerFields.push({
                  label: label,
                  value: value,
                  key: `header_${rowIdx}_${colIdx}_${label.replace(/[^a-zA-Z0-9]/g, '_').toLowerCase()}`
                });
              }
            }
          }
        }

        // Merged header rows: sparse json cells may omit title / period lines that live only in merge top-left.
        if (headerRowIndex > 0) {
          for (let r = 0; r < headerRowIndex; r++) {
            for (let c = 0; c < effectiveSheetCols; c++) {
              const txt = String(getMergedAwareCellText(r, c) || '').trim();
              if (!txt) continue;
              if (!formTitle) {
                const m = txt.match(/\b(FORM|FROM)\s+NO\.?\s*\d+/i);
                if (m) formTitle = m[0].replace(/\s+/g, ' ').trim();
                else if (/\bform\s*["']?\s*[a-z]\s*["']?\s*(?:\([^)]*\))?/i.test(txt)) {
                  formTitle = txt.replace(/\s+/g, ' ').trim();
                }
              }
              if (!formReference && /labou?r\s*welfare\s*fund\s*rules?/i.test(txt)) {
                formReference = txt.replace(/\s+/g, ' ').trim();
              }
              if (
                !formReference &&
                /^\(.*\)$/.test(txt) &&
                /\bsee\s+sub-?rule\b/i.test(txt) &&
                /maternity\s+benefit\s+rules?/i.test(txt)
              ) {
                formReference = txt.replace(/\s+/g, ' ').trim();
              }
              if (!formSubtitle && /register\s+of\s+muster\s+roll/i.test(txt)) {
                formSubtitle = txt.replace(/\s+/g, ' ').trim();
              }
              if (
                !wagePeriodText &&
                /\bfor\s+the\s+period\s+from\b/i.test(txt) &&
                /\bto\b/i.test(txt)
              ) {
                wagePeriodText = txt.replace(/\s+/g, ' ').trim();
              }
            }
          }
        }

        // Form I: split merged heading block so title/subtitle render like Excel (line 1: Form - I, line 2: REGISTER OF WORKMEN).
        const formIParsed = splitFormIHeaderBlock([formTitle, formSubtitle, formReference].filter(Boolean).join('\n'));
        if (formIParsed) {
          if (formIParsed.title) formTitle = formIParsed.title;
          if (formIParsed.subtitle) formSubtitle = formIParsed.subtitle;
          if (formIParsed.reference) formReference = formIParsed.reference;
        }

        if (isFormBLabourWelfareContext({ title: formTitle, subtitle: formSubtitle, reference: formReference }, null, '', headers)) {
          // Form B should not render wage period banner.
          wagePeriodText = '';
        }

        if (formTitle && formSubtitle) {
          formSubtitle = stripSubtitleTitleEcho(formTitle, formSubtitle);
        }

        if (formTitle && /\bform\s+no\.?\s*10\b/i.test(formTitle)) {
          const form10lines = [];
          for (let hr = 0; hr < headerRows.length; hr++) {
            const row = headerRows[hr];
            const cells = row.map((x) => String(x || '').trim()).filter(Boolean);
            for (let c = 0; c < cells.length; c++) {
              const cell = cells[c];
              if (/\(prescribed\s+under\s+rule/i.test(cell) && cell.length < 200) {
                const cleaned = stripSubtitleTitleEcho(formTitle, cell);
                if (cleaned && !form10lines.some((l) => l.includes('Prescribed under Rule'))) form10lines.push(cleaned);
              }
              if (/\bovertime\s+muster\s+roll\b/i.test(cell) && !/name\s+and\s+address/i.test(cell) && cell.length < 220) {
                const cleaned = stripSubtitleTitleEcho(formTitle, cell);
                if (cleaned && !form10lines.some((l) => l.includes('Overtime Muster'))) form10lines.push(cleaned);
              }
            }
          }
          if (form10lines.length) {
            let sub = stripSubtitleTitleEcho(formTitle, form10lines.join('\n'));
            if (!formSubtitle) {
              if (sub) formSubtitle = sub;
            } else {
              const add = form10lines.filter((line) => !formSubtitle.toLowerCase().includes(line.toLowerCase().slice(0, Math.min(48, line.length))));
              if (add.length) {
                const sub2 = stripSubtitleTitleEcho(formTitle, add.join('\n'));
                if (sub2) formSubtitle = [formSubtitle, sub2].filter(Boolean).join('\n');
              }
            }
          }
        }

        /** Form V / Form 25: "Approved Festival Holidays" with a numbered column row (1…n) and values on the row below. Form 25 uses up to 12 slots. */
        const normHeaderLabel = (s) => String(s || '').toLowerCase().replace(/\s+/g, ' ').trim();
        let festivalGrid = null;
        if (headerRowIndex > 0) {
          /** Longest left-to-right run 1,2,…,n in this row (merged-aware); supports Form V (5) and Form 25 (12). */
          const findBestConsecutiveOneBasedSlotRow = (nr) => {
            const maxC = Math.max(100, effectiveSheetCols);
            const cells = [];
            for (let cc = 0; cc < maxC; cc++) {
              const raw = String(getMergedAwareCellText(nr, cc) || '').trim().replace(/\s+/g, '');
              if (!/^\d{1,2}$/.test(raw)) continue;
              const n = parseInt(raw, 10);
              if (n < 1 || n > 31) continue;
              cells.push({ c: cc, n });
            }
            cells.sort((a, b) => a.c - b.c);
            let best = null;
            for (let i = 0; i < cells.length; i++) {
              if (cells[i].n !== 1) continue;
              const cols = [cells[i].c];
              let expect = 2;
              let prevC = cells[i].c;
              for (let j = i + 1; j < cells.length; j++) {
                if (cells[j].c <= prevC) continue;
                if (cells[j].n !== expect) continue;
                cols.push(cells[j].c);
                prevC = cells[j].c;
                expect += 1;
              }
              if (cols.length >= 2 && (!best || cols.length > best.length)) best = cols;
            }
            return best;
          };
          for (let r = 0; r < headerRowIndex; r++) {
            for (let c = 0; c < Math.max(45, effectiveSheetCols); c++) {
              const cell = normHeaderLabel(getMergedAwareCellText(r, c));
              if (!/approved festival holidays/.test(cell)) continue;
              for (let dr = 1; dr <= 3; dr++) {
                const nr = r + dr;
                if (nr >= headerRowIndex) break;
                const slotCols = findBestConsecutiveOneBasedSlotRow(nr);
                if (!slotCols || slotCols.length < 2) continue;
                const dataRow = nr + 1;
                if (dataRow >= headerRowIndex) continue;
                const initialValues = slotCols.map((col) =>
                  String(getMergedAwareCellText(dataRow, col) || (jsonData[dataRow] && jsonData[dataRow][col] != null ? jsonData[dataRow][col] : '') || '')
                    .trim()
                );
                festivalGrid = {
                  labelRow: r,
                  numberRow: nr,
                  dataRow,
                  slotCols,
                  keys: slotCols.map((_, idx) => `festival_approved_slot_${idx + 1}`),
                  initialValues
                };
                break;
              }
              if (festivalGrid) break;
            }
            if (festivalGrid) break;
          }
        }
        /** Form V: "Festival Holidays Approval Proceedings No. and Date:" — show with the same block as the 1–5 date grid. */
        const isFestivalApprovalProceedingsLabel = (label) => {
          const n = normHeaderLabel(label);
          return n.includes('approval proceedings') && (n.includes('festival') || n.includes('holidays'));
        };
        const findApprovalProceedingsInSheet = () => {
          for (let r = 0; r < headerRowIndex; r++) {
            for (let c = 0; c < Math.max(50, effectiveSheetCols); c++) {
              const raw = String(getMergedAwareCellText(r, c) || '').trim();
              if (!raw || !/festival\s+holidays?\s+approval\s+proceedings/i.test(normHeaderLabel(raw))) continue;
              let value = '';
              let valueCol = c + 1;
              for (let nc = c + 1; nc < Math.min(c + 45, Math.max(120, effectiveSheetCols)); nc++) {
                const t = String(getMergedAwareCellText(r, nc) || '').trim();
                if (!t) continue;
                if (/^approved\s+festival\s+holidays?/i.test(t)) break;
                if (t.includes(':') && t.length < 80 && (isFestivalApprovalProceedingsLabel(t) || /approved festival/i.test(normHeaderLabel(t)))) continue;
                value = t;
                valueCol = nc;
                break;
              }
              return {
                key: 'festival_approval_proceedings',
                label: raw.includes(':') ? raw : `${raw}:`,
                initialValue: value,
                labelRow: r,
                valueCol
              };
            }
          }
          return null;
        };
        if (festivalGrid) {
          const apSheet = findApprovalProceedingsInSheet();
          const fromFields = headerFields.find((f) => isFestivalApprovalProceedingsLabel(f.label));
          if (apSheet) {
            festivalGrid.approvalProceedings = apSheet;
          } else if (fromFields) {
            festivalGrid.approvalProceedings = {
              key: 'festival_approval_proceedings',
              label: (fromFields.label && String(fromFields.label).trim()) || 'Festival Holidays Approval Proceedings No. and Date:',
              initialValue: fromFields.value || '',
              labelRow: null,
              valueCol: null
            };
          } else {
            festivalGrid.approvalProceedings = {
              key: 'festival_approval_proceedings',
              label: 'Festival Holidays Approval Proceedings No. and Date:',
              initialValue: '',
              labelRow: null,
              valueCol: null
            };
          }
        }
        if (festivalGrid) {
          headerFields = headerFields.filter((f) => {
            if (normHeaderLabel(f.label).includes('approved festival holidays')) return false;
            if (isFestivalApprovalProceedingsLabel(f.label)) return false;
            return true;
          });
        }

        // Form 12: factory / contractor / registration in one merged cell — three lines.
        // Form 10 (Rule 78): factory + contractor in one cell; "Month ending" usually in another cell.
        if (headerRowIndex > 0) {
          for (let r = 0; r < headerRowIndex; r++) {
            for (let c = 0; c < Math.max(20, effectiveSheetCols); c++) {
              const cellText = String(getMergedAwareCellText(r, c) || '').trim();
              if (cellText.length < 30) continue;
              const parts12 = splitForm12FactoryContractorRegBlock(cellText);
              if (parts12 && parts12.length >= 2) {
                headerFields = headerFields.filter((f) => {
                  const lab = String(f.label || '').trim();
                  if (lab === cellText) return false;
                  if (splitForm12FactoryContractorRegBlock(lab)) return false;
                  return true;
                });
                const form12Keys = ['form12_header_factory', 'form12_header_contractor', 'form12_header_registration'];
                parts12.forEach((part, i) => {
                  headerFields.push({
                    label: part.label,
                    value: part.value,
                    key: form12Keys[i]
                  });
                });
                form12MergedHeaderCell = { r, c };
                break;
              }
              const parts10 = splitForm10FactoryContractorBlock(cellText);
              if (parts10 && parts10.length >= 2) {
                headerFields = headerFields.filter((f) => {
                  const lab = String(f.label || '').trim();
                  if (lab === cellText) return false;
                  if (splitForm10FactoryContractorBlock(lab)) return false;
                  if (splitForm12FactoryContractorRegBlock(lab)) return false;
                  return true;
                });
                parts10.forEach((part, i) => {
                  headerFields.push({
                    label: part.label,
                    value: part.value,
                    key: i === 0 ? 'form10_header_factory' : 'form10_header_contractor'
                  });
                });
                form10MergedHeaderCell = { r, c };
                break;
              }
            }
            if (form12MergedHeaderCell || form10MergedHeaderCell) break;
          }
        }

        if (form10MergedHeaderCell && headerRowIndex > 0) {
          const monthFieldFromParser = headerFields.find((f) => /^month\s*ending/i.test(String(f.label || '').trim()));
          let monthVal = monthFieldFromParser ? String(monthFieldFromParser.value || '').trim() : '';
          headerFields = headerFields.filter((f) => !/^month\s*ending/i.test(String(f.label || '').trim()));
          let foundRef = false;
          for (let r = 0; r < headerRowIndex; r++) {
            for (let c = 0; c < Math.max(30, effectiveSheetCols); c++) {
              const raw = String(getMergedAwareCellText(r, c) || '').trim();
              if (!raw || !/month\s*ending/i.test(raw)) continue;
              const inline = raw.match(/^Month ending\s*:\s*(.*)$/i);
              if (inline) {
                monthVal = (inline[1] || monthVal).trim() || monthVal;
                form10MonthEndingCell = { r, c, writeMode: 'full' };
                foundRef = true;
                break;
              }
            }
            if (foundRef) break;
          }
          if (!form10MonthEndingCell) {
            for (let r = 0; r < headerRowIndex; r++) {
              for (let c = 0; c < Math.max(30, effectiveSheetCols); c++) {
                const raw = String(getMergedAwareCellText(r, c) || '').trim();
                if (!/^Month ending\s*:?\s*$/i.test(raw)) continue;
                const v = String(getMergedAwareCellText(r, c + 1) || '').trim();
                monthVal = v || monthVal;
                form10MonthEndingCell = { r, c: c + 1, writeMode: 'value' };
                foundRef = true;
                break;
              }
              if (form10MonthEndingCell) break;
            }
          }
          const f10 = (k) => headerFields.find((f) => f.key === k);
          const rest = headerFields.filter(
            (f) => !['form10_header_factory', 'form10_header_contractor', 'form10_header_month_ending'].includes(f.key)
          );
          const orderedF10 = [f10('form10_header_factory'), f10('form10_header_contractor')].filter(Boolean);
          headerFields = [
            ...orderedF10,
            { label: 'Month ending:', value: monthVal, key: 'form10_header_month_ending' },
            ...rest
          ];
        }

        const enriched25 = enrichForm25AddressHeadersFromSheet(headerRowIndex, effectiveSheetCols, getMergedAwareCellText);
        const looksLikeForm25Sheet =
          /\b(form|from)\s*no\.?\s*25\b/i.test(formTitle || '') ||
          (festivalGrid && festivalGrid.slotCols && festivalGrid.slotCols.length >= 8 && enriched25.length > 0);
        if (looksLikeForm25Sheet && enriched25.length > 0) {
            const nf = (s) => String(s || '').toLowerCase().replace(/\s+/g, ' ').trim();
            headerFields = headerFields.filter((f) => {
              const nl = nf(f.label);
              if (nl.includes('name and address of the contractor') && !nl.includes('factory')) return false;
              if (nl.includes('establishment') && nl.includes('contract') && nl.includes('carried')) return false;
              if (nl.includes('principal employer')) return false;
              return true;
            });
            headerFields = [
              ...enriched25.map((h) => ({ label: h.label, value: h.value, key: h.key })),
              ...headerFields
            ];
        }

        // Form B (TN LWF Register of Wages): "For the Month of" is often split across cells without ":" on the label,
        // so the colon-based parser above misses it. Scan merged cells above the table row.
        const formBHeaderProbe = { title: formTitle, subtitle: formSubtitle, reference: formReference };
        if (
          isFormBLabourWelfareContext(formBHeaderProbe, null, '', headers) &&
          headerRowIndex > 0
        ) {
          const hasMonthField = headerFields.some((f) => /for\s+the\s+month\s+of/i.test(String(f.label || '')));
          if (!hasMonthField) {
            const maxC = Math.max(45, effectiveSheetCols);
            monthScan: for (let r = 0; r < headerRowIndex; r++) {
              for (let c = 0; c < maxC; c++) {
                const raw = String(getMergedAwareCellText(r, c) || '').trim();
                if (!raw) continue;
                const normalized = raw.replace(/\s*:?\s*$/, '').trim();
                if (!/^for\s+the\s+month\s+of\b/i.test(normalized)) continue;
                let label = `${normalized}:`;
                let value = '';
                for (let nc = c + 1; nc < Math.min(c + 10, maxC); nc++) {
                  const v = String(getMergedAwareCellText(r, nc) || '').trim();
                  if (!v) continue;
                  if (/^for\s+the\s+month\s+of/i.test(v)) continue;
                  if (/^name\s+of\s+the\s+establishment/i.test(v)) continue;
                  value = v;
                  break;
                }
                headerFields.push({
                  label,
                  value,
                  key: 'form_b_header_for_the_month_of'
                });
                break monthScan;
              }
            }
          }
        }

        if (formTitle || formSubtitle || formReference || headerFields.length > 0 || wagePeriodText || festivalGrid) {
          formHeaderInfo = {
            title: formTitle,
            subtitle: formSubtitle,
            reference: formReference,
            wagePeriodText: wagePeriodText,
            fields: headerFields,
            textRows: headerTextRows,
            festivalGrid: festivalGrid || undefined,
            form12MergedHeaderCell: form12MergedHeaderCell || undefined,
            form10MergedHeaderCell: form10MergedHeaderCell || undefined,
            form10MonthEndingCell: form10MonthEndingCell || undefined
          };
        }
      }
    }
   
    // Extract table data – use expanded headers when subcolumns exist so each column maps correctly
    const hasSubColumns = subColumnsData && Object.keys(subColumnsData).length > 0;
    const headersToUse = expandedHeaders.length > 0 ? expandedHeaders : headers;

    /** Bottom sheet row of any merge that contains (headerRowIndex, c) — multi-row merged thead (e.g. LWF Form C). */
    let firstRowAfterHeaderBlock =
      headerRowIndex >= 0 ? headerRowIndex + 1 : 0;
    if (headerRowIndex >= 0 && merges.length > 0 && headersToUse.length > 0) {
      let bottom = headerRowIndex;
      const scanCols = Math.min(Math.max(headersToUse.length, 1), Math.max(effectiveSheetCols, 1));
      for (let c = 0; c < scanCols; c++) {
        for (let mi = 0; mi < merges.length; mi++) {
          const m = merges[mi];
          if (!m || !m.s || !m.e) continue;
          if (headerRowIndex >= m.s.r && headerRowIndex <= m.e.r && c >= m.s.c && c <= m.e.c) {
            bottom = Math.max(bottom, m.e.r);
          }
        }
      }
      firstRowAfterHeaderBlock = bottom + 1;
    }

    /** Form No. 10 style: row of 1, 2, … n under column titles (not day subcolumns) — skip it for data rows */
    let extraRowsAfterHeader = 0;
    if (!hasSubColumns && headerRowIndex >= 0 && headersToUse.length >= 2 && firstRowAfterHeaderBlock < jsonData.length) {
      const r = firstRowAfterHeaderBlock;
      let seqOk = true;
      for (let c = 0; c < headersToUse.length; c++) {
        const t = String(getMergedAwareCellText(r, c) || '').replace(/\s+/g, '').trim();
        if (t !== String(c + 1)) {
          seqOk = false;
          break;
        }
      }
      if (seqOk) extraRowsAfterHeader = 1;
    }

    const startIndex =
      (headerRowIndex >= 0 && hasSubColumns)
        ? headerRowIndex + 2
        : (headerRowIndex >= 0 ? firstRowAfterHeaderBlock + extraRowsAfterHeader : 0);
    const tableData = [];
    const maxCols = headersToUse.length || (jsonData.length > startIndex ? Math.max(...jsonData.slice(startIndex).map(r => r ? r.length : 0)) : 0);
   
    if (hasSubColumns && expandedHeaders.length > 0) {
      const orderedMains = [];
      let curMain = null;
      for (let colIdx = 0; colIdx < headers.length; colIdx++) {
        const h = (headers[colIdx] || '').trim();
        if (h) curMain = h;
        if (curMain && orderedMains[orderedMains.length - 1] !== curMain) orderedMains.push(curMain);
      }
      if (orderedMains.length === 0) orderedMains.push(...Object.keys(subColumnsData));
      for (let i = startIndex; i < jsonData.length; i++) {
        const rowData = {};
        let excelColIdx = 0;
        orderedMains.forEach((main) => {
          const subs = subColumnsData[main] || [];
          if (subs.length > 0) {
            subs.forEach((sub) => {
              const key = `${main}_${sub}`;
              rowData[key] = String(getMergedAwareCellText(i, excelColIdx) || '').trim();
              excelColIdx++;
            });
          } else {
            rowData[main] = String(getMergedAwareCellText(i, excelColIdx) || '').trim();
            excelColIdx++;
          }
        });
        if (Object.values(rowData).some((val) => val)) tableData.push(rowData);
      }
    } else {
      for (let i = startIndex; i < jsonData.length; i++) {
        const rowData = {};
        for (let j = 0; j < maxCols; j++) {
          const header = headersToUse[j] || `Column ${j + 1}`;
          rowData[header] = String(getMergedAwareCellText(i, j) || '').trim();
        }
        if (Object.values(rowData).some((val) => val)) tableData.push(rowData);
      }
    }

    if (isFixedRowAggregateComplianceTable(headersToUse) && tableData.length > 1) {
      const deduped = dedupeConsecutiveFixedAggregateBodyRows(tableData);
      tableData.length = 0;
      deduped.forEach((r) => tableData.push(r));
    }

    /** Form X: merged row above detail headers (Earned Leave, Medical Leave, …) — not employer meta like Registration Certificate. */
    let columnGroupLabels = null;
    if (!hasSubColumns && headerRowIndex > 0 && headersToUse.length > 0) {
      const titleCompact = (formHeaderInfo?.title || '').toString().toUpperCase().replace(/\s+/g, '');
      const subUpper = (formHeaderInfo?.subtitle || '').toString().toUpperCase();
      const isFormX =
        titleCompact.includes('FORM-X') ||
        titleCompact.includes('FORMX') ||
        (subUpper.includes('REGISTER OF LEAVE') && subUpper.includes('SOCIAL SECURITY'));

      const normalizeBandCell = (txt) =>
        String(txt || '')
          .replace(/\r?\n/g, ' ')
          .replace(/\s+/g, ' ')
          .trim();

      const buildRawGroupsForRow = (r) => {
        const arr = [];
        for (let c = 0; c < headersToUse.length; c++) {
          arr.push(normalizeBandCell(getMergedAwareCellText(r, c)));
        }
        return arr;
      };

      const joinedLower = (arr) => arr.filter(Boolean).join(' ').toLowerCase();

      const isEmployerMetaOnlyRow = (jl) =>
        /registration\s*certificate/.test(jl) ||
        ((/employer|establishment|incharge|manager|factory|premises/.test(jl) ||
          /name\s+and\s+address/.test(jl)) &&
          !jl.includes('earned leave'));

      const isLeaveCategoryBannerRow = (jl) =>
        jl.includes('earned leave') ||
        jl.includes('medical') ||
        jl.includes('maternity') ||
        jl.includes('other leave');

      let rawGroups = buildRawGroupsForRow(headerRowIndex - 1);
      let joined = joinedLower(rawGroups);

      if ((isFormX || isEmployerMetaOnlyRow(joined)) && !isLeaveCategoryBannerRow(joined)) {
        let picked = false;
        for (let off = 2; off <= 8; off++) {
          const r = headerRowIndex - off;
          if (r < 0) break;
          const candidate = buildRawGroupsForRow(r);
          const jl = joinedLower(candidate);
          if (isEmployerMetaOnlyRow(jl) && !isLeaveCategoryBannerRow(jl)) continue;
          if (isLeaveCategoryBannerRow(jl)) {
            rawGroups = candidate;
            joined = jl;
            picked = true;
            break;
          }
        }
        if (!picked) {
          rawGroups = rawGroups.map((cell) =>
            /registration\s*certificate/i.test(String(cell || '')) ? '' : cell
          );
          joined = joinedLower(rawGroups);
        }
      }

      const rowLooksLikeLeaveRegister =
        joined.includes('earned leave') &&
        (joined.includes('medical') || joined.includes('maternity'));
      if (isFormX || rowLooksLikeLeaveRegister || isLeaveCategoryBannerRow(joined)) {
        columnGroupLabels = rawGroups;
        if (!columnGroupLabels.some((x) => String(x || '').trim())) columnGroupLabels = null;
      }
    }

    /** Form W / wage registers: merged parent row (e.g. "Deductions") sits above mid-tier headers; Form X logic only reads one row up. */
    if (columnGroupLabels == null && !hasSubColumns && headerRowIndex >= 2 && headersToUse.length > 0) {
      const hdrBlob = headersToUse.map((h) => String(h || '').toLowerCase()).join(' | ');
      const looksWageRegister =
        (hdrBlob.includes('gross wage') || hdrBlob.includes('basic wage')) &&
        (hdrBlob.includes('net wage') || hdrBlob.includes('deduction')) &&
        (hdrBlob.includes('provident') ||
          hdrBlob.includes('esi') ||
          hdrBlob.includes('state insurance') ||
          hdrBlob.includes('advance'));
      if (looksWageRegister) {
        const normCell = (s) =>
          String(s || '')
            .replace(/\r?\n/g, ' ')
            .replace(/\s+/g, ' ')
            .trim();
        const isDeductionBanner = (s) => {
          const t = normCell(s).toLowerCase();
          return t === 'deduction' || t === 'deductions';
        };
        let bestRow = -1;
        let bestDedBannerCols = 0;
        for (let offset = 1; offset <= 4; offset += 1) {
          const r = headerRowIndex - offset;
          if (r < 0) break;
          const raw = [];
          let dedBannerCols = 0;
          for (let c = 0; c < headersToUse.length; c++) {
            const t = normCell(getMergedAwareCellText(r, c));
            raw.push(t);
            if (isDeductionBanner(t)) dedBannerCols += 1;
          }
          const rowText = raw.filter(Boolean).join(' ').toLowerCase();
          if (/wage\s*period/i.test(rowText)) continue;
          if (dedBannerCols >= 2 && dedBannerCols >= bestDedBannerCols) {
            bestDedBannerCols = dedBannerCols;
            bestRow = r;
          }
        }
        if (bestRow >= 0) {
          const rawGroups = [];
          for (let c = 0; c < headersToUse.length; c++) {
            rawGroups.push(normCell(getMergedAwareCellText(bestRow, c)));
          }
          if (rawGroups.some((x) => x)) columnGroupLabels = rawGroups;
        }
      }
    }

    return {
      formHeader: formHeaderInfo,
      tableData: tableData,
      headers: headersToUse,
      subColumns: subColumnsData,
      headerRowIndex,
      dataStartIndex: startIndex,
      columnGroupLabels
    };
  };

  const handleHeaderFieldChange = (key, value) => {
    setHeaderFormData(prev => ({
      ...prev,
      [key]: value
    }));
  };

  const handleImportFormRows = async (file, headersForImport) => {
    try {
      if (!file) return;
      const normalizedHeaders = Array.isArray(headersForImport) ? headersForImport : [];
      if (normalizedHeaders.length === 0) {
        setError('No table headers available for import.');
        return;
      }

      const arrayBuffer = await file.arrayBuffer();
      const workbook = XLSX.read(arrayBuffer, { type: 'array' });
      const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
      if (!firstSheet) {
        setError('Selected file has no worksheet.');
        return;
      }

      const rows = XLSX.utils.sheet_to_json(firstSheet, { header: 1, defval: '' });
      if (!Array.isArray(rows) || rows.length === 0) {
        setError('Selected file is empty.');
        return;
      }

      const norm = (v) =>
        String(v || '')
          .trim()
          .toLowerCase()
          .replace(/[\r\n]+/g, ' ')
          .replace(/\s+/g, ' ')
          .replace(/[^a-z0-9 ]/g, '');

      // Detect which row in imported sheet is the actual header row by best match score.
      let detectedHeaderRowIndex = -1;
      let bestMatchCount = -1;
      rows.forEach((row, rowIndex) => {
        if (!Array.isArray(row)) return;
        const importRowNorm = row.map((cell) => norm(cell));
        const matchCount = normalizedHeaders.reduce((acc, header) => {
          const expected = norm(header);
          if (!expected) return acc;
          return acc + (importRowNorm.includes(expected) ? 1 : 0);
        }, 0);
        if (matchCount > bestMatchCount) {
          bestMatchCount = matchCount;
          detectedHeaderRowIndex = rowIndex;
        }
      });

      const hasDetectedHeaderRow = bestMatchCount >= Math.max(1, Math.ceil(normalizedHeaders.length * 0.35));
      const importHeaderRow = hasDetectedHeaderRow && Array.isArray(rows[detectedHeaderRowIndex])
        ? rows[detectedHeaderRowIndex]
        : [];

      // Build source-column lookup by header text so imported values fill corresponding target headers.
      const sourceIndicesByHeaderNorm = new Map();
      importHeaderRow.forEach((cell, idx) => {
        const key = norm(cell);
        if (!key) return;
        const prev = sourceIndicesByHeaderNorm.get(key) || [];
        prev.push(idx);
        sourceIndicesByHeaderNorm.set(key, prev);
      });

      const headerOccurrenceCounter = new Map();
      const sourceColumnIndexForTarget = normalizedHeaders.map((header, idx) => {
        const key = norm(header);
        const options = sourceIndicesByHeaderNorm.get(key) || [];
        const seen = headerOccurrenceCounter.get(key) || 0;
        headerOccurrenceCounter.set(key, seen + 1);
        if (options.length > seen) return options[seen];
        return idx; // fallback to positional mapping when exact header not found
      });

      // Keep modal headers unchanged; imported header row should never be injected as data.
      const dataStartIndex = (detectedHeaderRowIndex >= 0 && bestMatchCount > 0) ? (detectedHeaderRowIndex + 1) : 0;
      const candidateRows = rows.slice(dataStartIndex).filter((row) => {
        if (!Array.isArray(row)) return false;
        return row.some((cell) => String(cell ?? '').trim() !== '');
      });

      if (candidateRows.length === 0) {
        setError('No data rows found in the imported file.');
        return;
      }

      let importedRows = candidateRows.map((row, rowIndex) => {
        const obj = {};
        normalizedHeaders.forEach((header, idx) => {
          const sourceIdx = sourceColumnIndexForTarget[idx];
          const rawVal = sourceIdx >= 0 ? row[sourceIdx] : '';
          const headerNorm = norm(header);
          if (/^s\.?\s*no\.?$/.test(headerNorm) && String(rawVal ?? '').trim() === '') {
            obj[header] = String(rowIndex + 1);
          } else {
            obj[header] = rawVal != null ? String(rawVal) : '';
          }
        });
        return obj;
      });

      if (isFixedRowAggregateComplianceTable(normalizedHeaders) && importedRows.length > 1) {
        importedRows = dedupeConsecutiveFixedAggregateBodyRows(importedRows);
      }

      const existingRows = Array.isArray(formTableData) ? formTableData : [];
      const totalRows = Math.max(existingRows.length, importedRows.length);
      const mergedRows = Array.from({ length: totalRows }, (_, rowIdx) => {
        const existingRow = existingRows[rowIdx] || {};
        const importedRow = importedRows[rowIdx] || {};
        const merged = {};
        normalizedHeaders.forEach((header) => {
          const importedValue = importedRow[header];
          const existingValue = existingRow[header];
          const hasImportedValue = importedValue != null && String(importedValue).trim() !== '';
          merged[header] = hasImportedValue ? importedValue : (existingValue != null ? existingValue : '');
        });
        return merged;
      });

      setFormTableData(mergedRows);
      setSuccess(`Imported ${importedRows.length} rows and merged with existing data.`);
      setError('');
    } catch (err) {
      console.error('Import rows error:', err);
      setError(err?.message || 'Failed to import Excel rows.');
    }
  };

  const formFileNameLooksLikePlaceholder = (name) => {
    if (name == null || String(name).trim() === '') return true;
    const n = String(name).trim().toLowerCase();
    if (n === 'view file') return true;
    if (!String(name).includes('.')) return true;
    return false;
  };

  const defaultFormTemplateFileName = (item) => {
    const base = ((item?.formName || item?.FormName || 'form') + '').trim().replace(/[/\\?%*:|"<>]/g, '') || 'form';
    return `${base}.xlsx`;
  };

  const normalizeTemplateFormNameKey = (name) =>
    String(name || '')
      .replace(/\.xlsx$/i, '')
      .replace(/[^a-z0-9]+/gi, ' ')
      .replace(/\s+/g, ' ')
      .trim()
      .toLowerCase();

  const parseContentDispositionFilename = (header) => {
    if (!header || typeof header !== 'string') return null;
    const m = /filename\*=UTF-8''([^;\n]+)|filename\*=([^;\n]+)|filename="([^"]+)"|filename=([^;\n]+)/i.exec(header);
    if (!m) return null;
    const raw = (m[1] || m[2] || m[3] || m[4] || '').trim();
    if (!raw) return null;
    try {
      let s = raw.replace(/^"|"$/g, '');
      if (m[1] || m[2]) s = decodeURIComponent(s);
      return s;
    } catch {
      return raw.replace(/^"|"$/g, '');
    }
  };

  const arrayBufferLooksLikeExcel = (buffer) => {
    if (!buffer || buffer.byteLength < 4) return false;
    const u8 = new Uint8Array(buffer.slice(0, 8));
    if (u8[0] === 0x50 && u8[1] === 0x4b) return true;
    if (u8[0] === 0xd0 && u8[1] === 0xcf && u8[2] === 0x11 && u8[3] === 0xe0) return true;
    return false;
  };

  const arrayBufferLooksLikePdf = (buffer) => {
    if (!buffer || buffer.byteLength < 4) return false;
    const u8 = new Uint8Array(buffer.slice(0, 4));
    return u8[0] === 0x25 && u8[1] === 0x50 && u8[2] === 0x44 && u8[3] === 0x46;
  };

  // Build direct download URL for form file (so "View File" can open Excel in new tab)
  const getFormFileDownloadUrl = (item) => {
    if (!item || !item.formFile || item.formFile === 'null' || String(item.formFile).trim() === '') return null;
    const isFromFormmaster = item.isFromFormmaster && item.formFile;
    const isFromChecklist = item.isFromChecklist || (item.id && String(item.id).startsWith('checklist_'));
    if (isFromFormmaster) {
      const rawName = formFileNameLooksLikePlaceholder(item.formFileName)
        ? defaultFormTemplateFileName(item)
        : item.formFileName;
      const fileName = String(rawName).replace(/[^a-zA-Z0-9._-]/g, '_');
      return `${window.location.origin}/server/formmaster_function/templates/download/${item.formFile}?fileName=${encodeURIComponent(fileName)}&disposition=attachment`;
    }
    if (isFromChecklist) {
      const checklistId = item.checklistId || String(item.id).replace(/^checklist_/, '');
      return `${window.location.origin}/server/checklist_function/checklist/${checklistId}/file/Form`;
    }
    return `${window.location.origin}/server/statutoryreg_function/statutory/${item.id}/file/Form`;
  };

  /** Same resource as "View File" — Autofill must load this URL so the modal matches the Form File column */
  const getFormFileFetchUrl = (item) => {
    const full = getFormFileDownloadUrl(item);
    if (!full) return null;
    try {
      const url = new URL(full);
      url.searchParams.delete('disposition');
      return url.pathname + (url.search ? url.search : '');
    } catch {
      return null;
    }
  };

  /** Build candidate fetch URLs so Autofill can still open if one mapping is stale. */
  const getFormFileFetchUrlCandidates = (item) => {
    const urls = [];
    const primary = getFormFileFetchUrl(item);
    if (primary) urls.push(primary);
    if (item?.formFile && item.formFile !== 'null' && String(item.formFile).trim() !== '') {
      const rawName = formFileNameLooksLikePlaceholder(item.formFileName)
        ? defaultFormTemplateFileName(item)
        : (item.formFileName || defaultFormTemplateFileName(item));
      const fileName = String(rawName).replace(/[^a-zA-Z0-9._-]/g, '_');
      urls.push(`/server/formmaster_function/templates/download/${item.formFile}?fileName=${encodeURIComponent(fileName)}`);
    }
    return Array.from(new Set(urls.filter(Boolean)));
  };

  const handleViewFormFile = async (item, forceAutofill = false, options = {}) => {
    try {
      setFormFileLoading(true);
      setError('');
      setFormHeader(null);
      setHeaderFormData({});
      setFormTableData([]);
      setTableHeaders([]);
      setSubColumns(null);

      const fileUrls = getFormFileFetchUrlCandidates(item);
      if (!fileUrls.length) {
        throw new Error('No form file linked for this row');
      }
      let response = null;
      let lastFetchError = null;
      for (let i = 0; i < fileUrls.length; i++) {
        try {
          const candidateResp = await fetch(fileUrls[i]);
          if (candidateResp.ok) {
            response = candidateResp;
            break;
          }
          const errorData = await candidateResp.json().catch(() => ({ message: 'Failed to load file' }));
          lastFetchError = new Error(errorData.message || 'Failed to load form file');
        } catch (fetchErr) {
          lastFetchError = fetchErr instanceof Error ? fetchErr : new Error('Failed to load form file');
        }
      }
      if (!response) {
        throw lastFetchError || new Error('Failed to load form file');
      }

      const contentType = (response.headers.get('content-type') || '').toLowerCase();
      const cdName = parseContentDispositionFilename(response.headers.get('content-disposition'));
      const nameFromItem = formFileNameLooksLikePlaceholder(item.formFileName)
        ? defaultFormTemplateFileName(item)
        : (item.formFileName || '');
      let displayFileName = cdName || nameFromItem || defaultFormTemplateFileName(item);
      if (formFileNameLooksLikePlaceholder(displayFileName)) {
        displayFileName = defaultFormTemplateFileName(item);
      }

      const arrayBuffer = await response.arrayBuffer();
      const extFromName = displayFileName.includes('.') ? displayFileName.split('.').pop().toLowerCase() : '';
      const typeSaysExcel = contentType.includes('spreadsheetml') || contentType.includes('ms-excel') || contentType.includes('vnd.ms-excel');
      const isExcel = extFromName === 'xlsx' || extFromName === 'xls' || typeSaysExcel || arrayBufferLooksLikeExcel(arrayBuffer);
      const isPdf = extFromName === 'pdf' || contentType.includes('pdf') || arrayBufferLooksLikePdf(arrayBuffer);

      if (isExcel) {
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        const firstSheetName = workbook.SheetNames?.[0];
        const firstSheet = firstSheetName ? workbook.Sheets[firstSheetName] : null;
        const excelSheetHtml = firstSheet
          ? XLSX.utils.sheet_to_html(firstSheet, { editable: true })
          : '';

        const parsed = parseExcelForm(workbook);

        const resolvedFormFileName = item.formFileName && !formFileNameLooksLikePlaceholder(item.formFileName)
          ? item.formFileName
          : displayFileName;

        const templateWageLine = parsed.formHeader?.wagePeriodText || '';

        const rawMonthToken =
          (item.monthFilter || item.MonthFilter || item.monthfilter || '').trim() ||
          (selectedMonth && String(selectedMonth).trim()) ||
          '';
        const fullMonthForPeriod =
          resolveToFullMonthName(rawMonthToken) ||
          getMonthFromDueDate(item.dueDate || '') ||
          resolveToFullMonthName(String(selectedMonth || '').trim()) ||
          inferMonthFromWagePeriodLine(templateWageLine);

        const yearForPeriod =
          extractYearFromDueDate(item.dueDate || '') ||
          extractYearFromWagePeriodLine(templateWageLine) ||
          new Date().getFullYear();

        let formHeaderForModal = parsed.formHeader ? { ...parsed.formHeader } : null;
        if (isRegisterOfWagesFormContext(formHeaderForModal, item, displayFileName, parsed.headers)) {
          let wageLine = '';
          if (fullMonthForPeriod) {
            wageLine = buildRegisterWagePeriodLine(fullMonthForPeriod, yearForPeriod);
          }
          if (!wageLine && /wage\s*period/i.test(templateWageLine)) {
            wageLine = repairTruncatedYearsInWagePeriodLine(templateWageLine, yearForPeriod);
          }
          if (wageLine) {
            formHeaderForModal = {
              ...(formHeaderForModal || {
                title: '',
                subtitle: '',
                reference: '',
                fields: [],
                textRows: []
              }),
              wagePeriodText: wageLine
            };
          }
        }

        // Form B: show "For the Month of" — template may omit it or have #REF!; default from Month filter / due date.
        if (formHeaderForModal && isFormBLabourWelfareContext(formHeaderForModal, item, displayFileName, parsed.headers)) {
          const monthKey = 'form_b_header_for_the_month_of';
          const monthFromStatutory =
            fullMonthForPeriod && yearForPeriod
              ? `${fullMonthForPeriod} ${yearForPeriod}`
              : String(fullMonthForPeriod || '').trim();
          const fields = Array.isArray(formHeaderForModal.fields) ? [...formHeaderForModal.fields] : [];
          const idx = fields.findIndex(
            (f) => f.key === monthKey || /for\s+the\s+month\s+of/i.test(String(f.label || '').trim())
          );
          const valueBad = (v) => {
            const s = String(v ?? '').trim();
            return !s || /^#REF!/i.test(s);
          };
          if (idx === -1 && monthFromStatutory) {
            fields.push({
              label: 'For the Month of:',
              value: monthFromStatutory,
              key: monthKey
            });
          } else if (idx !== -1 && monthFromStatutory && valueBad(fields[idx].value)) {
            fields[idx] = { ...fields[idx], value: monthFromStatutory };
          }
          formHeaderForModal = { ...formHeaderForModal, fields };
        }

        setFormHeader(formHeaderForModal);
        setTableHeaders(parsed.headers);
        setSubColumns(parsed.subColumns);

        let initialHeaderFormData = {};
        if (formHeaderForModal) {
          if (formHeaderForModal.fields) {
            formHeaderForModal.fields.forEach((field) => {
              initialHeaderFormData[field.key] = field.value || '';
            });
          }
          const fg = formHeaderForModal.festivalGrid;
          if (fg && Array.isArray(fg.keys)) {
            fg.keys.forEach((key, i) => {
              initialHeaderFormData[key] =
                (Array.isArray(fg.initialValues) && fg.initialValues[i] != null && String(fg.initialValues[i]) !== ''
                  ? String(fg.initialValues[i])
                  : '') || '';
            });
          }
          if (fg && fg.approvalProceedings && fg.approvalProceedings.key) {
            const ap = fg.approvalProceedings;
            initialHeaderFormData[ap.key] =
              ap.initialValue != null && String(ap.initialValue) !== '' ? String(ap.initialValue) : '';
          }
          if (formHeaderForModal.fields || (fg && (fg.keys || fg.approvalProceedings))) {
            setHeaderFormData(initialHeaderFormData);
          }
        }

        setFormFileModalData({
          fileName: displayFileName,
          formFileName: resolvedFormFileName,
          fileType: 'excel-form',
          sheetHtml: excelSheetHtml,
          rawData: workbook,
          item: item,
          parsedFormHeader: formHeaderForModal,
          parsedTableHeaders: parsed.headers || [],
          parsedSubColumns: parsed.subColumns,
          parsedHeaderFormData: initialHeaderFormData,
          headerRowIndex: parsed.headerRowIndex,
          dataStartIndex: parsed.dataStartIndex,
          parsedColumnGroupLabels: parsed.columnGroupLabels ?? null
        });

        const shouldAutofill = forceAutofill || isAutofillMode;

        if (shouldAutofill) {
          console.log('🔄 Autofill: template from Form File – headers:', (parsed.headers || []).length);
          // Show the selected form file immediately on first click, then replace rows with Autofill data.
          setFormTableData(parsed.tableData || []);
          const shouldPreferSavedDraftData = !!options?.preferSavedDraftData;
          if (shouldPreferSavedDraftData) {
            let loadedSavedData = false;
            const modalReqTs = Date.now();
            try {
              const sampleRowId =
                resolveNumericStatutoryIdForProofRow(item) ||
                (isNumericStatutoryBackendId(item?.id) ? String(item.id).trim() : null);
              if (sampleRowId) {
                // Prefer the saved Draft Excel (same source as "Download form template")
                // so View Draft modal shows exactly what user saved.
                const draftResp = await fetch(`/server/statutoryreg_function/statutory/${sampleRowId}/file/Draft?_ts=${modalReqTs}`);
                if (draftResp.ok) {
                  const draftArrayBuffer = await draftResp.arrayBuffer();
                  const draftWb = XLSX.read(draftArrayBuffer, { type: 'array' });
                  const draftRows = extractMappedRowsFromSavedDraft({
                    draftWorkbook: draftWb,
                    templateHeaders: parsed.headers || [],
                    templateHeaderRowIndex: parsed.headerRowIndex,
                    templateDataStartIndex: parsed.dataStartIndex,
                    hasSubColumns: parsed.subColumns && Object.keys(parsed.subColumns || {}).length > 0
                  });
                  if (Array.isArray(draftRows) && draftRows.length > 0) {
                    setFormTableData(draftRows);
                    loadedSavedData = true;
                    setSuccess('Loaded saved draft data.');
                    setTimeout(() => setSuccess(''), 2500);
                  }
                  // Some forms (like Form 25) are easier to recover via matrix rows.
                  // If header-based extraction is empty, map matrix rows to current template headers.
                  if (!loadedSavedData) {
                    const matrixRows = extractRowMatrixFromSavedDraft({
                      draftWorkbook: draftWb,
                      templateHeaderRowIndex: parsed.headerRowIndex,
                      templateDataStartIndex: parsed.dataStartIndex,
                      hasSubColumns: parsed.subColumns && Object.keys(parsed.subColumns || {}).length > 0
                    });
                    if (Array.isArray(matrixRows) && matrixRows.length > 0) {
                      const headersForMatrix =
                        Array.isArray(parsed.headers) && parsed.headers.length > 0
                          ? parsed.headers
                          : Array.from({ length: matrixRows[0]?.length || 0 }, (_, idx) => `Column ${idx + 1}`);
                      const matrixAsObjects = matrixRows.map((rowArr, rowIdx) => {
                        const obj = {};
                        headersForMatrix.forEach((h, colIdx) => {
                          const v = Array.isArray(rowArr) ? rowArr[colIdx] : '';
                          obj[h] = v != null ? v : '';
                        });
                        if ((obj[headersForMatrix[0]] == null || String(obj[headersForMatrix[0]]).trim() === '') && headersForMatrix.length > 0) {
                          obj[headersForMatrix[0]] = rowIdx + 1;
                        }
                        return obj;
                      });
                      setFormTableData(matrixAsObjects);
                      loadedSavedData = true;
                      setSuccess('Loaded saved draft data.');
                      setTimeout(() => setSuccess(''), 2500);
                    }
                  }
                }
                if (!loadedSavedData) {
                  const sampleResp = await fetch(`/server/statutoryreg_function/statutory/${sampleRowId}/sampledata?_ts=${modalReqTs}`);
                  if (sampleResp.ok) {
                    const sampleJson = await sampleResp.json().catch(() => null);
                    const sampleRows = sampleJson?.data?.sampleData?.rows;
                    if (Array.isArray(sampleRows) && sampleRows.length > 0) {
                      setFormTableData(sampleRows);
                      loadedSavedData = true;
                      setSuccess('Loaded saved draft/imported data.');
                      setTimeout(() => setSuccess(''), 2500);
                    }
                  }
                }
              }
            } catch (sampleErr) {
              console.warn('Could not load saved sample data for View Draft modal:', sampleErr);
            }
            if (!loadedSavedData) {
              fetchAndPopulateEmployeeData(parsed.headers, {
                columnGroupLabels: parsed.columnGroupLabels ?? null,
                formFileModalData: {
                  fileName: displayFileName,
                  formFileName: resolvedFormFileName,
                  fileType: 'excel-form',
                  sheetHtml: excelSheetHtml,
                  rawData: workbook,
                  item: item,
                  parsedFormHeader: formHeaderForModal,
                  parsedTableHeaders: parsed.headers || [],
                  parsedSubColumns: parsed.subColumns,
                  parsedHeaderFormData: initialHeaderFormData,
                  headerRowIndex: parsed.headerRowIndex,
                  dataStartIndex: parsed.dataStartIndex,
                  parsedColumnGroupLabels: parsed.columnGroupLabels ?? null
                }
              });
            }
          } else {
            fetchAndPopulateEmployeeData(parsed.headers, {
              columnGroupLabels: parsed.columnGroupLabels ?? null,
            });
          }
        } else {
          console.log('📄 View File mode – using parsed Excel data');
          setFormTableData(parsed.tableData);
        }
      } else if (isPdf) {
        const blob = new Blob([arrayBuffer], { type: 'application/pdf' });
        const blobUrl = URL.createObjectURL(blob);
        setFormFileModalData({
          fileName: displayFileName,
          fileType: 'pdf',
          content: blobUrl
        });
      } else {
        const blob = new Blob([arrayBuffer]);
        const blobUrl = URL.createObjectURL(blob);
        setFormFileModalData({
          fileName: displayFileName,
          fileType: extFromName === 'png' || extFromName === 'jpg' || extFromName === 'jpeg' || extFromName === 'gif' ? 'image' : 'other',
          content: blobUrl
        });
      }

      setIsFormFileModalOpen(true);
    } catch (err) {
      console.error('Error loading form file:', err);
      setError(err.message || 'Failed to load form file');
    } finally {
      setFormFileLoading(false);
    }
  };

  const handleCloseFormFileModal = () => {
    setIsFormFileModalOpen(false);
    setIsAutofillMode(false);
    setAutofillItem(null);
    setError('');
    setSuccess('');
    // Clean up blob URLs if any
    if (formFileModalData?.content && formFileModalData?.fileType !== 'excel') {
      try {
        URL.revokeObjectURL(formFileModalData.content);
      } catch (e) {
        // Ignore errors
      }
    }
    setFormFileModalData(null);
  };

  // Helper function to determine which act category a form belongs to
  const getActCategory = useCallback((item) => {
    if (!item) return null;
   
    const normalizeString = (str) => {
      if (!str) return '';
      return String(str).trim().toLowerCase();
    };
   
    // Support all common field names from statutory API, ChecklistBulk, and backend.
    // Forms A/B/D often store context in description/formName while Act is a standalone statute (e.g. LWF, Payment of Wages).
    const act = normalizeString(item.act || item.Act || item.actName || item.ActName || '');
    const desc = normalizeString(item.description || item.Description || '');
    const fname = normalizeString(item.formName || item.FormName || '');
    const haystack = `${act} ${desc} ${fname}`.trim();
    const sector = normalizeString(item.sector || item.Sector || '');
   
    // Check for Factories Act (bare "factory" only on act text — avoids noise in description)
    const isFactoriesAct =
      haystack.includes('factories act') ||
      haystack.includes('factory act') ||
      act.includes('factories') ||
      act.includes('factory') ||
      sector === 'factories act' ||
      sector === 'factory act' ||
      sector.includes('factories act') ||
      sector.includes('factory act') ||
      sector.includes('factories') ||
      sector.includes('factory');
   
    if (isFactoriesAct) return 'factories';
   
    // Check for Shops and Establishment (incl. TN LWF / commercial establishment — common for Form B etc.)
    const isShopsAndEstablishment =
      haystack.includes('shops and establishments') ||
      haystack.includes('shops and establishment') ||
      haystack.includes('shop and establishment') ||
      haystack.includes('the shops and establishments act') ||
      haystack.includes('the shops and establishment act') ||
      haystack.includes('commercial establishments') ||
      haystack.includes('commercial establishment') ||
      haystack.includes('labour welfare fund') ||
      haystack.includes('labor welfare fund') ||
      /\blwf\b/.test(haystack) ||
      haystack.includes('equal remuneration') ||
      sector === 'shops and establishment' ||
      sector === 'shops and establishments' ||
      sector === 'shops and establishment act' ||
      sector === 'shop and establishment' ||
      sector.includes('shops and establishment') ||
      sector.includes('shop and establishment');
   
    if (isShopsAndEstablishment) return 'shops_and_establishment';
   
    // Check for CLRA
    const isCLRA =
      haystack.includes('clra') ||
      haystack.includes('contract labour') ||
      haystack.includes('contract labor') ||
      sector === 'clra' ||
      sector.includes('clra') ||
      sector.includes('contract labour') ||
      sector.includes('contract labor');
   
    if (isCLRA) return 'clra';
   
    return 'other';
  }, []);

  // When saved statutory rows miss Sector/State columns, infer category from same-form siblings.
  const getActCategoryWithFormFallback = useCallback((item, allRows = []) => {
    const direct = getActCategory(item);
    if (direct && direct !== 'other') return direct;
    const formKey = baseFormNameKey(item?.formName || item?.FormName || '');
    if (!formKey || !Array.isArray(allRows) || allRows.length === 0) return direct;
    for (let i = 0; i < allRows.length; i++) {
      const row = allRows[i];
      if (baseFormNameKey(row?.formName || row?.FormName || '') !== formKey) continue;
      const cat = getActCategory(row);
      if (cat && cat !== 'other') return cat;
    }
    return direct;
  }, [getActCategory]);

  const hasComplianceFiles = useCallback((item) => {
    const hasVal = (v) => v != null && String(v).trim() !== '' && String(v).trim() !== 'null' && String(v).trim() !== 'undefined';
    return (
      hasVal(item?.draftFile ?? item?.DraftFile) ||
      hasVal(item?.proofSubmissionFile ?? item?.ProofSubmissionFile) ||
      hasVal(item?.formFile ?? item?.FormFile)
    );
  }, []);

  const resolveSiteDisplayName = useCallback((item, allRows = []) => {
    const explicitSite = String(item?.site || item?.Site || item?.siteName || item?.SiteName || '').trim();
    if (explicitSite) return explicitSite;
    const rowState = String(item?.state ?? item?.State ?? '').trim();
    if (rowState) {
      const inchargePool = Array.isArray(inchargeSitesMeta) ? inchargeSitesMeta : [];
      const orgPool = Array.isArray(organizationSitesMeta) ? organizationSitesMeta : [];
      const pool = inchargePool.length > 0 ? inchargePool : orgPool;
      if (pool.length > 0) {
        let byState = pool.filter((rec) => checklistStateMatchesSiteState(rowState, rec.siteState));
        if (byState.length > 0) {
          const rowSector = String(item?.sector ?? item?.Sector ?? '').trim();
          const rowAct = String(item?.act ?? item?.Act ?? '').trim();
          const rowCat = getActCategoryFromActSector(rowAct, rowSector);
          if (rowCat && rowCat !== 'other') {
            const byBucket = byState.filter((rec) => {
              const siteCat = industryLabelToActCategory(rec.industry);
              if (siteCat) return siteCat === rowCat;
              const ind = String(rec.industry || '').trim();
              if (!ind) return false;
              return (
                sectorMatchesInchargeSiteIndustries(rowSector, [ind]) ||
                sectorMatchesInchargeSiteIndustries(rowAct, [ind])
              );
            });
            byState = byBucket;
          } else if (rowSector) {
            const byFuzzy = byState.filter((rec) => {
              const ind = String(rec.industry || '').trim();
              if (!ind) return false;
              return (
                sectorMatchesInchargeSiteIndustries(rowSector, [ind]) ||
                sectorMatchesInchargeSiteIndustries(rowAct, [ind])
              );
            });
            if (byFuzzy.length > 0) byState = byFuzzy;
          }
          const names = [...new Set(byState.map((r) => r.siteName).filter(Boolean))];
          if (names.length === 1) return names[0];
          if (names.length > 1) return names.sort((a, b) => a.localeCompare(b)).join(', ');
        }
        if (inchargePool.length > 0) return '';
      }
    }
    if (Array.isArray(allowedSiteNameList) && allowedSiteNameList.length === 1) {
      return String(allowedSiteNameList[0] || '').trim();
    }
    if (siteFromUrl && String(siteFromUrl).trim() !== '') return String(siteFromUrl).trim();
    const cat = getActCategoryWithFormFallback(item, allRows);
    const names = Array.isArray(siteNamesByActCategory?.[cat]) ? siteNamesByActCategory[cat] : [];
    if (names.length > 0) return names[0];
    // If row category is unclear, but site login scope has exactly one allowed category, use that site's name.
    if ((cat === 'other' || !cat) && Array.isArray(allowedActCategoryList) && allowedActCategoryList.length === 1) {
      const scopedNames = Array.isArray(siteNamesByActCategory?.[allowedActCategoryList[0]])
        ? siteNamesByActCategory[allowedActCategoryList[0]]
        : [];
      if (scopedNames.length > 0) return scopedNames[0];
    }
    return '';
  }, [
    siteFromUrl,
    getActCategoryWithFormFallback,
    siteNamesByActCategory,
    allowedActCategoryList,
    allowedSiteNameList,
    inchargeSitesMeta,
    organizationSitesMeta
  ]);

  // Filter statutory data by month and ensure forms are properly separated by act
  // Items with "Monthly Basis" due date should appear in every month (treated as 15th of that month)
  const filteredStatutoryData = useMemo(() => {
    let data = statutoryData;

    // For afrindinu14 with ?site=: show only the corresponding sector act for that site (Delphi → SE, Delphi Kakinada → CLRA, Delphi Oragadam → Factories)
    if (siteWiseCategory && data && data.length > 0) {
      data = data.filter((item) => getActCategoryWithFormFallback(item, data) === siteWiseCategory);
    }
    // Site Management: Incharge + Industry → act categories (same as merge-time filter)
    else if (hasSiteBasedActScope && data && data.length > 0) {
      const allow = new Set(allowedActCategoryList);
      data = data.filter((item) => {
        const cat = getActCategoryWithFormFallback(item, data);
        return allow.has(cat) || (cat === 'other' && hasComplianceFiles(item));
      });
    }

    // Site login: show only corresponding site rows.
    if (siteFromUrl && data && data.length > 0) {
      const targetSite = String(siteFromUrl).trim().toLowerCase();
      data = data.filter((item) =>
        resolvedSiteMatchesSingleTarget(resolveSiteDisplayName(item, data) || '', targetSite)
      );
    } else if (Array.isArray(allowedSiteNameList) && allowedSiteNameList.length > 0 && data && data.length > 0) {
      const allowedSites = new Set(allowedSiteNameList.map((s) => String(s || '').trim().toLowerCase()).filter(Boolean));
      data = data.filter((item) => {
        const rowResolved = resolveSiteDisplayName(item, data) || '';
        return resolvedSiteAnyAllowed(rowResolved, allowedSites);
      });
    }

    if (Array.isArray(allowedInchargeStateLabels) && allowedInchargeStateLabels.length > 0 && data && data.length > 0) {
      data = data.filter((item) =>
        statesFieldMatchesInchargeSiteStates(item.state || item.State || '', allowedInchargeStateLabels)
      );
    }

    // Show all form names in every month: for selected month show matching record or placeholder (so "10 of 13" becomes 13 rows).
    // But when Statutory is already aligned 1:1 to ChecklistBulk, do not collapse by form/month.
    const effectiveSelectedMonth = resolveToFullMonthName(selectedMonth) || getCurrentMonth();
    const isChecklistBulkMirrored = Array.isArray(data) && data.length > 0 && data.every((row) => row?.isBulkImported === true);
    if (effectiveSelectedMonth && data && data.length > 0 && !isChecklistBulkMirrored) {
      const selectedMonthNorm = String(effectiveSelectedMonth).trim().toLowerCase().substring(0, 3);
      const formKey = (item) => {
        const idStr = String(item?.id ?? '').trim();
        const checklistStableId = item?.checklistId != null ? `checklist:${String(item.checklistId).trim()}` : '';
        const sourceStableId = checklistStableId || (/^(bulk_|checklist_)/i.test(idStr) ? `source:${idStr}` : '');
        const base = `${squashStatutoryKeyPart(item.formName || item.FormName)}|${squashStatutoryKeyPart(item.act || item.Act)}|${squashStatutoryKeyPart(item.description || item.Description)}|${squashStatutoryKeyPart(item.sector || item.Sector) || 'nosector'}|${squashStatutoryKeyPart(item.state || item.State) || 'nostate'}`;
        return sourceStableId ? `${base}|${sourceStableId}` : base;
      };
      // When several rows match the same month (e.g. ChecklistBulk + saved Statutory after Autofill→Save), prefer the row that has the draft file and a real DB id so the Draft column shows the Excel.
      const pickBestRecordForMonth = (matches) => {
        if (!matches || matches.length === 0) return null;
        if (matches.length === 1) return matches[0];
        const score = (r) => {
          let s = 0;
          const idStr = String(r.id ?? '');
          const dRef = r.draftFile ?? r.DraftFile;
          const hasDraft = dRef && dRef !== 'null' && String(dRef).trim() !== '';
          if (hasDraft) s += 100;
          if (/^\d+$/.test(idStr)) s += 50;
          if (!idStr.startsWith('bulk_') && !idStr.startsWith('placeholder_') && !idStr.startsWith('formmaster_')) s += 25;
          return s;
        };
        return [...matches].sort((a, b) => score(b) - score(a))[0];
      };
      const monthMatches = (item) => {
        const storedMonth = item.monthFilter || item.MonthFilter || item.monthfilter || '';
        const storedMonthNorm = String(storedMonth).trim().toLowerCase().substring(0, 3);
        if (storedMonthNorm) return selectedMonthNorm === storedMonthNorm;
        if (!item.dueDate) return false;
        const dueDateLower = String(item.dueDate).toLowerCase().trim();
        if (dueDateLower.includes('monthly basis')) return true;
        const monthNames = ['january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september', 'october', 'november', 'december'];
        const monthAbbr = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];
        const selectedMonthLower = String(effectiveSelectedMonth).toLowerCase();
        const monthIndex = monthNames.findIndex(m => m.startsWith(selectedMonthLower));
        if (monthIndex !== -1) {
          const fullMonthName = monthNames[monthIndex];
          const monthAbbrName = monthAbbr[monthIndex];
          return dueDateLower.includes(fullMonthName) || dueDateLower.includes(monthAbbrName);
        }
        return false;
      };
      const byFormKey = new Map();
      data.forEach((item) => {
        const key = formKey(item);
        if (!byFormKey.has(key)) byFormKey.set(key, []);
        byFormKey.get(key).push(item);
      });
      const out = [];
      byFormKey.forEach((records) => {
        const forMonth = pickBestRecordForMonth(records.filter(monthMatches));
        const template = records[0];
        if (forMonth) {
          out.push(forMonth);
        } else {
          // Strict month isolation:
          // If selected month has no matching saved row, render placeholder only.
          // Do not reuse a saved row from another month (prevents Jan save from showing in Feb+).
          out.push({
            ...template,
            id: template.id ? `placeholder_${template.id}_${selectedMonthNorm}` : `placeholder_${selectedMonthNorm}_${formKey(template)}`,
            formFile: template.isFromFormmaster ? template.formFile : null,
            formFileName: template.isFromFormmaster ? template.formFileName : null,
            draftFile: null,
            draftFileName: null,
            proofSubmissionFile: null,
            proofSubmissionFileName: null,
            monthFilter: effectiveSelectedMonth,
            monthfilter: effectiveSelectedMonth,
            isFromFormmaster: template.isFromFormmaster || false
          });
        }
      });
      data = dedupeStatutoryRowsForDisplay(
        collapseStatutoryDisplayDuplicates(out, effectiveSelectedMonth, resolveSiteDisplayName),
        effectiveSelectedMonth,
        resolveSiteDisplayName
      );
    }

    return data;
  }, [statutoryData, selectedMonth, getActCategoryWithFormFallback, hasComplianceFiles, resolveSiteDisplayName, siteWiseCategory, siteFromUrl, allowedSiteNameList, allowedInchargeStateLabels, hasSiteBasedActScope, allowedActCategoryList]);

  const tableFormOptions = useMemo(() => {
    return Array.from(
      new Set(
        (filteredStatutoryData || [])
          .map((item) => String(item.formName || item.FormName || '').trim())
          .filter(Boolean)
      )
    );
  }, [filteredStatutoryData]);

  const transactionTableData = useMemo(() => {
    const searchValue = String(tableSearch || '').trim().toLowerCase();
    const filtered = (filteredStatutoryData || []).filter((item) => {
      const formName = String(item.formName || item.FormName || '').trim();
      const searchMatch = !searchValue || [
        item.act,
        item.description,
        item.formName,
        item.FormName,
        item.dueDate,
        item.approval,
        item.Approval,
        item.status,
        item.Status,
        item.sendForApproval,
        item.SendForApproval,
        item.remarks,
        item.Remarks
      ].some((value) => String(value || '').toLowerCase().includes(searchValue));
      const formMatch = selectedFormFilter === 'all' || formName === selectedFormFilter;
      return searchMatch && formMatch;
    });
    // Group site-wise; within a site: non-approved with draft/sent first, then other non-approved, then approved last.
    return [...filtered].sort((a, b) => {
      const siteA = String(resolveSiteDisplayName(a, filteredStatutoryData) || '').trim().toLowerCase();
      const siteB = String(resolveSiteDisplayName(b, filteredStatutoryData) || '').trim().toLowerCase();
      if (siteA !== siteB) {
        if (!siteA) return 1;
        if (!siteB) return -1;
        return siteA.localeCompare(siteB);
      }
      const tierA = isRowStatusApproved(a) ? 1 : 0;
      const tierB = isRowStatusApproved(b) ? 1 : 0;
      if (tierA !== tierB) return tierA - tierB;
      const draftRankA = !isRowStatusApproved(a) && hasStatutoryDraftSubmittedForSort(a) ? 0 : 1;
      const draftRankB = !isRowStatusApproved(b) && hasStatutoryDraftSubmittedForSort(b) ? 0 : 1;
      if (draftRankA !== draftRankB) return draftRankA - draftRankB;
      const timeA = getStatutoryRowActivityTimeMs(a);
      const timeB = getStatutoryRowActivityTimeMs(b);
      if (timeB !== timeA) return timeB - timeA;
      const formA = String(a.formName || a.FormName || '').trim().toLowerCase();
      const formB = String(b.formName || b.FormName || '').trim().toLowerCase();
      if (formA !== formB) return formA.localeCompare(formB);
      return String(a.id || '').localeCompare(String(b.id || ''));
    });
  }, [filteredStatutoryData, tableSearch, selectedFormFilter, resolveSiteDisplayName]);

  const totalPages = Math.max(1, Math.ceil(transactionTableData.length / PAGE_SIZE));
  const effectivePage = Math.min(currentPage, totalPages);
  const pagedTransactionTableData = useMemo(() => {
    const start = (effectivePage - 1) * PAGE_SIZE;
    return transactionTableData.slice(start, start + PAGE_SIZE);
  }, [transactionTableData, effectivePage]);
  const showRemarksColumn = useMemo(
    () =>
      // In site login, always show Remarks column.
      String(siteFromUrl || '').trim() !== '' ||
      hasSiteBasedActScope ||
      String(userRole || '').trim().toLowerCase().includes('site') ||
      String(userRole || '').trim().toLowerCase().includes('incharge') ||
      transactionTableData.some((item) => {
        const rawA = item?.approval ?? item?.Approval ?? '';
        const aNorm = String(rawA).trim().toLowerCase();
        return aNorm === 'rejected' || aNorm === 'reject';
      }),
    [transactionTableData, siteFromUrl, hasSiteBasedActScope, userRole]
  );
  const roleLower = String(userRole || '').trim().toLowerCase();
  /** Admin sees Approve/Reject column; drafts must still be visible for review before site sends for approval. */
  const isAppAdministrator = roleLower === 'app administrator';
  const isSiteLoginRole =
    roleLower.includes('site') ||
    roleLower.includes('incharge') ||
    roleLower.includes('site incharge');
  const isNonDefaultRoleLogin =
    roleLower !== '' &&
    roleLower !== 'app user' &&
    roleLower !== 'app administrator';
  const isSiteLoginEmail =
    isSiteWiseStatutoryUser ||
    effectiveUserEmail === 'afrindinu14@gmail.com';
  const isSiteView =
    String(siteFromUrl || '').trim() !== '' ||
    isSiteWiseStatutoryUser ||
    !!siteWiseCategory ||
    hasSiteBasedActScope ||
    isSiteLoginRole;
  const showSendForApprovalColumn =
    isSiteView ||
    isSiteLoginRole ||
    isSiteLoginEmail ||
    hasSiteBasedActScope ||
    isNonDefaultRoleLogin;
  const showApprovalColumn = !showSendForApprovalColumn;
  /** Single-site contexts (?site= or Site Management lists one site): Site column is redundant — hide it.
   *  Until the first statutory fetch completes, hide the Site column so `useCacheFirst` + `silentRefresh` cannot paint
   *  cached rows before `allowedSiteNameList` / `allowedActCategoryList` exist (that mismatch caused the column to flash).
   *  For site-scoped logins, also hide while `allowedSiteNameList` is still unknown (`null`). */
  const showSiteColumn = useMemo(() => {
    if (!transactionSiteMetaReady) return false;
    if (String(siteFromUrl || '').trim() !== '') return false;
    if (Array.isArray(allowedSiteNameList) && allowedSiteNameList.length === 1) return false;
    const siteScopedLogin =
      hasSiteBasedActScope || isSiteLoginRole || isSiteWiseStatutoryUser;
    if (allowedSiteNameList == null && siteScopedLogin) return false;
    return true;
  }, [
    transactionSiteMetaReady,
    siteFromUrl,
    allowedSiteNameList,
    hasSiteBasedActScope,
    isSiteLoginRole,
    isSiteWiseStatutoryUser
  ]);
  const tableColSpan =
    2 +
    (showSiteColumn ? 1 : 0) +
    8 +
    (showSendForApprovalColumn ? 1 : 0) +
    (showApprovalColumn ? 1 : 0) +
    1 +
    (showRemarksColumn ? 1 : 0);
  const renderRemarksDisplay = (rawRemarks) => {
    const text = String(rawRemarks || '');
    const lines = text
      .split(/\r?\n/)
      .map((line) => line.trim())
      .filter(Boolean);
    if (lines.length === 0) return '-';
    if (lines.length === 1) return lines[0];
    const isNumberedLine = (line) => /^\d+\s*[\.\)]\s*/.test(line);
    return (
      <div style={{ whiteSpace: 'normal', lineHeight: 1.35 }}>
        {lines.map((line, idx) => (
          <div key={`r-${idx}`}>{isNumberedLine(line) ? line : `${idx + 1}. ${line}`}</div>
        ))}
      </div>
    );
  };

  const paginationItems = useMemo(
    () => buildPaginationItems(effectivePage, totalPages),
    [effectivePage, totalPages]
  );

  useEffect(() => {
    setCurrentPage(1);
  }, [tableSearch, selectedFormFilter, selectedMonth]);

  useEffect(() => {
    if (currentPage > totalPages) setCurrentPage(totalPages);
  }, [currentPage, totalPages]);

  const clearTransactionFilters = () => {
    setTableSearch('');
    setSelectedFormFilter('all');
    const currentMonth = getCurrentMonth();
    setSelectedMonth(currentMonth);
    persistMonthFilter(currentMonth);
    setCurrentPage(1);
  };

  // Form file shared across months: once uploaded in SEMaster, show in every month by looking up from any record with same form name
  const formFileByFormKey = useMemo(() => {
    const map = new Map();
    if (!statutoryData || !Array.isArray(statutoryData)) return map;
    statutoryData.forEach((row) => {
      const formFile = row.formFile || row.FormFile || null;
      const hasFile = formFile && formFile !== 'null' && String(formFile).trim() !== '';
      if (!hasFile) return;
      const sec = String(row.sector || row.Sector || '').trim().toLowerCase();
      const st = String(row.state || row.State || '').trim().toLowerCase();
      const key = `${(row.formName || '').toLowerCase().trim()}|${(row.act || '').toLowerCase().trim()}|${sec}|${st}`;
      if (!map.has(key)) {
        map.set(key, {
          formFile,
          formFileName: row.formFileName || row.FormFileName || 'View File',
          id: row.id
        });
      }
    });
    return map;
  }, [statutoryData]);

  // Draft file lookup for site/month views: if current rendered row has no draft, reuse only strict matching row
  // (same form + act + description + month + sector + state) to avoid cross-row bleed.
  const draftRowByFormActMonthKey = useMemo(() => {
    const map = new Map();
    if (!Array.isArray(statutoryData) || statutoryData.length === 0) return map;
    statutoryData.forEach((row) => {
      const draftFile = row?.draftFile ?? row?.DraftFile ?? null;
      const hasDraft = draftFile && draftFile !== 'null' && String(draftFile).trim() !== '';
      if (!hasDraft) return;
      const monthNorm = String(row?.monthFilter || row?.MonthFilter || row?.monthfilter || '').trim().toLowerCase().substring(0, 3);
      const descNorm = String(row?.description || row?.Description || '').trim().toLowerCase();
      const secNorm = String(row?.sector || row?.Sector || '').trim().toLowerCase();
      const stateNorm = String(row?.state || row?.State || '').trim().toLowerCase();
      const key = `${String(row?.formName || '').toLowerCase().trim()}|${String(row?.act || '').toLowerCase().trim()}|${descNorm}|${monthNorm || 'nomonth'}|${secNorm}|${stateNorm}`;
      if (!map.has(key)) {
        map.set(key, row);
        return;
      }
      const prev = map.get(key);
      const prevId = String(prev?.id ?? '');
      const rowId = String(row?.id ?? '');
      const prevNumeric = /^\d+$/.test(prevId);
      const rowNumeric = /^\d+$/.test(rowId);
      if (!prevNumeric && rowNumeric) {
        map.set(key, row);
      }
    });
    return map;
  }, [statutoryData]);

  // Autofill/View modal: use format read from the View File (parsed form header + table headers) so each form opens with its own structure
  const displayFormHeader = useMemo(() =>
    (formFileModalData?.parsedFormHeader != null) ? formFileModalData.parsedFormHeader : formHeader,
  [formFileModalData?.parsedFormHeader, formHeader]);
  const displayTableHeaders = useMemo(() => {
    const parsed = formFileModalData?.parsedTableHeaders;
    const baseHeaders = Array.isArray(parsed) && parsed.length > 0 ? parsed : tableHeaders || [];
    const effectiveSelectedMonth = resolveToFullMonthName(selectedMonth) || '';
    const selectedMonthIndex = MONTH_NAMES.findIndex(
      (month) => month.toLowerCase().startsWith(String(effectiveSelectedMonth || '').toLowerCase().trim())
    );
    if (selectedMonthIndex < 0) {
      return baseHeaders;
    }

    const monthDayCount = new Date(new Date().getFullYear(), selectedMonthIndex + 1, 0).getDate();
    const isDayHeader = (header) => {
      if (header == null) return false;
      const text = String(header).trim();
      const numericMatch = text.match(/^(?:.+_)?(\d{1,2})$/);
      if (!numericMatch) return false;
      const day = Number(numericMatch[1]);
      return Number.isFinite(day) && day >= 1 && day <= 31;
    };

    const dayHeadersFound = baseHeaders.some((header) => {
      if (!isDayHeader(header)) return false;
      const day = Number(String(header).trim().replace(/^.+_/, ''));
      return day >= 1 && day <= 31;
    });
    if (!dayHeadersFound) {
      return baseHeaders;
    }

    return baseHeaders.filter((header) => {
      if (!isDayHeader(header)) return true;
      const day = Number(String(header).trim().replace(/^.+_/, ''));
      return day <= monthDayCount;
    });
  }, [formFileModalData?.parsedTableHeaders, tableHeaders, selectedMonth]);

  const displayColumnGroupLabels = useMemo(() => {
    const g = formFileModalData?.parsedColumnGroupLabels;
    if (!Array.isArray(g) || !displayTableHeaders.length) return null;
    if (g.length !== displayTableHeaders.length) return null;
    if (!g.some((x) => String(x || '').trim())) return null;
    return g;
  }, [formFileModalData?.parsedColumnGroupLabels, displayTableHeaders]);

  /** Employer / company line for the signature block below the form table (Form I and all excel-form modals). */
  const statutoryModalEmployerName = useMemo(() => {
    const item = formFileModalData?.item;
    const fromItem = [
      item?.companyName,
      item?.CompanyName,
      item?.establishmentName,
      item?.employerName,
      item?.nameOfEmployer,
      item?.siteName,
      item?.SiteName
    ]
      .map((x) => String(x || '').trim())
      .find(Boolean);
    if (fromItem) return fromItem;
    const site = String(siteFromUrl || '').trim();
    if (site) {
      return site
        .replace(/[-_]+/g, ' ')
        .replace(/\b\w/g, (ch) => ch.toUpperCase())
        .trim();
    }
    return '';
  }, [formFileModalData?.item, siteFromUrl]);

  /** Form W: mirror Excel’s multi-row merged header (Deductions, Advances, …) using sheet !merges. */
  const registerOfWagesMergedTheadRows = useMemo(() => {
    const wb = formFileModalData?.rawData;
    const ds = formFileModalData?.dataStartIndex;
    const hdrs = displayTableHeaders;
    const item = formFileModalData?.item;
    const fn = formFileModalData?.fileName || '';
    if (!wb || wb.SheetNames == null || !wb.SheetNames[0] || ds == null || ds < 1 || !hdrs?.length) return null;
    if (!isRegisterOfWagesFormContext(displayFormHeader, item, fn, hdrs)) return null;
    const ws = wb.Sheets[wb.SheetNames[0]];
    const band = adjustRegisterOfWagesHeaderBand(ws, ds, hdrs.length);
    if (!band) return null;
    const rows = buildRegisterOfWagesHeaderRows(ws, band.r0, band.r1, hdrs.length);
    if (!rows || !rows.length) return null;
    return rows;
  }, [
    formFileModalData?.rawData,
    formFileModalData?.dataStartIndex,
    formFileModalData?.item,
    formFileModalData?.fileName,
    displayTableHeaders,
    displayFormHeader
  ]);

  /** Form V / Form 25: two-row header — merged parent (“Daily Hours…” or “DATES”) over days 1…31, day numbers on row 2 (Excel layout). */
  const formVDailyHoursThead = useMemo(() => {
    const hdrs = displayTableHeaders;
    const info = getDailyHoursGridGroupInfo(hdrs);
    if (!info) return null;
    const { startIndex, parentLabel } = info;
    const dayCount = hdrs.length - startIndex;
    if (dayCount < 2) return null;

    const groupHeaderStyle = {
      backgroundColor: STAT_FORM_FILE_ACCENT_BG,
      color: 'white',
      padding: '10px 12px',
      textAlign: 'center',
      fontWeight: '700',
      fontSize: '13px',
      border: `1px solid ${STAT_FORM_FILE_ACCENT_BORDER}`,
      verticalAlign: 'middle'
    };
    const identityOrDetailHeaderStyle = {
      backgroundColor: STAT_FORM_FILE_ACCENT_BG,
      color: 'white',
      padding: '12px',
      textAlign: 'left',
      fontWeight: '600',
      fontSize: '14px',
      border: `1px solid ${STAT_FORM_FILE_ACCENT_BORDER}`,
      verticalAlign: 'middle'
    };
    const stickyDetailHeaderStyle = {
      ...identityOrDetailHeaderStyle,
      textAlign: 'center',
      fontSize: '13px',
      minWidth: '44px',
      position: 'sticky',
      top: 0,
      zIndex: 10
    };

    return (
      <>
        <tr>
          {hdrs.slice(0, startIndex).map((h, i) => (
            <th key={`formv-left-${i}`} rowSpan={2} style={identityOrDetailHeaderStyle} title={h}>
              {h}
            </th>
          ))}
          <th key="formv-dailyhours" colSpan={dayCount} style={groupHeaderStyle}>
            {parentLabel}
          </th>
        </tr>
        <tr>
          {hdrs.slice(startIndex).map((h, i) => (
            <th key={`formv-day-${i}`} style={stickyDetailHeaderStyle} title={h}>
              {formatStatutoryTableHeaderLabel(h)}
            </th>
          ))}
        </tr>
      </>
    );
  }, [displayTableHeaders]);

  /** Form D: two-row header — "Components of Remuneration" over Basic/Dearness/House/Other; left & right cols rowspan 2. */
  const formDRemunerationThead = useMemo(() => {
    const hdrs = displayTableHeaders;
    const item = formFileModalData?.item;
    const fn = formFileModalData?.fileName || '';
    if (!Array.isArray(hdrs) || hdrs.length < 2) return null;
    if (!isFormDRemunerationFormContext(displayFormHeader, item, fn, hdrs)) return null;
    const info = getFormDRemunerationGroupInfo(hdrs);
    if (!info) return null;
    const { startIndex, endIndex, parentLabel } = info;
    const left = hdrs.slice(0, startIndex);
    const subs = hdrs.slice(startIndex, endIndex);
    const right = hdrs.slice(endIndex);
    if (subs.length < 2) return null;

    const groupHeaderStyle = {
      backgroundColor: STAT_FORM_FILE_ACCENT_BG,
      color: 'white',
      padding: '10px 12px',
      textAlign: 'center',
      fontWeight: '700',
      fontSize: '13px',
      border: `1px solid ${STAT_FORM_FILE_ACCENT_BORDER}`,
      verticalAlign: 'middle'
    };
    const identityOrDetailHeaderStyle = {
      backgroundColor: STAT_FORM_FILE_ACCENT_BG,
      color: 'white',
      padding: '12px',
      textAlign: 'left',
      fontWeight: '600',
      fontSize: '14px',
      border: `1px solid ${STAT_FORM_FILE_ACCENT_BORDER}`,
      verticalAlign: 'middle'
    };
    const stickyDetailHeaderStyle = {
      ...identityOrDetailHeaderStyle,
      textAlign: 'center',
      fontSize: '13px',
      minWidth: '72px',
      position: 'sticky',
      top: 0,
      zIndex: 10
    };

    return (
      <>
        <tr>
          {left.map((h, i) => (
            <th key={`formd-left-${i}`} rowSpan={2} style={identityOrDetailHeaderStyle} title={h}>
              {h}
            </th>
          ))}
          <th key="formd-rem-group" colSpan={subs.length} style={groupHeaderStyle}>
            {parentLabel}
          </th>
          {right.map((h, i) => (
            <th key={`formd-right-${i}`} rowSpan={2} style={identityOrDetailHeaderStyle} title={h}>
              {h}
            </th>
          ))}
        </tr>
        <tr>
          {subs.map((h, i) => (
            <th key={`formd-sub-${i}`} style={stickyDetailHeaderStyle} title={h}>
              {formatStatutoryTableHeaderLabel(h)}
            </th>
          ))}
        </tr>
      </>
    );
  }, [displayTableHeaders, displayFormHeader, formFileModalData?.item, formFileModalData?.fileName]);

  /** Form B (TN LWF Register of Wages): render grouped parent/sub headers from `Parent_Sub` keys. */
  const formBLabourWelfareThead = useMemo(() => {
    const hdrs = displayTableHeaders;
    const item = formFileModalData?.item;
    const fn = formFileModalData?.fileName || '';
    if (!Array.isArray(hdrs) || hdrs.length < 2) return null;
    if (!isFormBLabourWelfareContext(displayFormHeader, item, fn, hdrs)) return null;

    const splitHeader = (h) => {
      const m = String(h || '').match(/^(.+)_([\s\S]+)$/);
      if (!m) return null;
      return {
        parent: m[1].trim(),
        sub: m[2].trim()
      };
    };

    let hasGroupWithSubs = false;
    const chunks = [];
    for (let i = 0; i < hdrs.length; i += 1) {
      const first = splitHeader(hdrs[i]);
      if (!first) {
        chunks.push({ kind: 'single', header: hdrs[i], start: i, end: i + 1 });
        continue;
      }
      let j = i + 1;
      while (j < hdrs.length) {
        const next = splitHeader(hdrs[j]);
        if (!next || next.parent.toLowerCase() !== first.parent.toLowerCase()) break;
        j += 1;
      }
      const span = j - i;
      if (span >= 2) {
        hasGroupWithSubs = true;
        chunks.push({ kind: 'group', parent: first.parent, start: i, end: j });
      } else {
        chunks.push({ kind: 'single', header: hdrs[i], start: i, end: i + 1 });
      }
      i = j - 1;
    }
    if (!hasGroupWithSubs) return null;

    const groupHeaderStyle = {
      backgroundColor: STAT_FORM_FILE_ACCENT_BG,
      color: 'white',
      padding: '10px 12px',
      textAlign: 'center',
      fontWeight: '700',
      fontSize: '13px',
      border: `1px solid ${STAT_FORM_FILE_ACCENT_BORDER}`,
      verticalAlign: 'middle'
    };
    const identityOrDetailHeaderStyle = {
      backgroundColor: STAT_FORM_FILE_ACCENT_BG,
      color: 'white',
      padding: '12px',
      textAlign: 'left',
      fontWeight: '600',
      fontSize: '14px',
      border: `1px solid ${STAT_FORM_FILE_ACCENT_BORDER}`,
      verticalAlign: 'middle'
    };
    const stickyDetailHeaderStyle = {
      ...identityOrDetailHeaderStyle,
      textAlign: 'center',
      fontSize: '13px',
      position: 'sticky',
      top: 0,
      zIndex: 10
    };

    const row1Cells = chunks.map((chunk, idx) => {
      if (chunk.kind === 'single') {
        return (
          <th key={`formb-single-${idx}`} rowSpan={2} style={identityOrDetailHeaderStyle} title={chunk.header}>
            {formatStatutoryTableHeaderLabel(chunk.header)}
          </th>
        );
      }
      return (
        <th key={`formb-group-${idx}`} colSpan={chunk.end - chunk.start} style={groupHeaderStyle} title={chunk.parent}>
          {chunk.parent}
        </th>
      );
    });

    const row2Cells = [];
    chunks.forEach((chunk, chunkIdx) => {
      if (chunk.kind !== 'group') return;
      for (let col = chunk.start; col < chunk.end; col += 1) {
        const split = splitHeader(hdrs[col]);
        const sub = split?.sub || formatStatutoryTableHeaderLabel(hdrs[col]);
        row2Cells.push(
          <th key={`formb-sub-${chunkIdx}-${col}`} style={stickyDetailHeaderStyle} title={sub}>
            {sub}
          </th>
        );
      }
    });

    if (row2Cells.length === 0) return null;
    return (
      <>
        <tr>{row1Cells}</tr>
        <tr>{row2Cells}</tr>
      </>
    );
  }, [displayTableHeaders, displayFormHeader, formFileModalData?.item, formFileModalData?.fileName]);

  /** Form VI: merged festival-holiday band over multiple date columns; preserve Excel-style grouped header. */
  const formVIFestivalThead = useMemo(() => {
    const hdrs = displayTableHeaders;
    const item = formFileModalData?.item;
    const fn = formFileModalData?.fileName || '';
    if (!Array.isArray(hdrs) || hdrs.length < 4) return null;
    if (!isFormVIFestivalContext(displayFormHeader, item, fn, hdrs)) return null;

    const norm = (v) => String(v || '').replace(/\s+/g, ' ').trim().toLowerCase();
    const looksFestivalBand = (h) => {
      const n = norm(h);
      return n.includes('national') && n.includes('festival') && n.includes('holidays');
    };

    let start = -1;
    let end = -1;
    for (let i = 0; i < hdrs.length; i += 1) {
      if (!looksFestivalBand(hdrs[i])) continue;
      if (start < 0) start = i;
      end = i + 1;
    }
    if (start < 0 || end <= start || end - start < 2) return null;

    const left = hdrs.slice(0, start);
    const groupCols = hdrs.slice(start, end);
    const right = hdrs.slice(end);

    const groupHeaderStyle = {
      backgroundColor: STAT_FORM_FILE_ACCENT_BG,
      color: 'white',
      padding: '10px 12px',
      textAlign: 'center',
      fontWeight: '700',
      fontSize: '13px',
      border: `1px solid ${STAT_FORM_FILE_ACCENT_BORDER}`,
      verticalAlign: 'middle'
    };
    const identityOrDetailHeaderStyle = {
      backgroundColor: STAT_FORM_FILE_ACCENT_BG,
      color: 'white',
      padding: '12px',
      textAlign: 'left',
      fontWeight: '600',
      fontSize: '14px',
      border: `1px solid ${STAT_FORM_FILE_ACCENT_BORDER}`,
      verticalAlign: 'middle'
    };
    const stickyDetailHeaderStyle = {
      ...identityOrDetailHeaderStyle,
      textAlign: 'center',
      fontSize: '13px',
      minWidth: '44px',
      position: 'sticky',
      top: 0,
      zIndex: 10
    };

    const uniqueFestivalHeaders = new Set(groupCols.map((h) => norm(h)).filter(Boolean));
    const showIndexedSubHeaders = uniqueFestivalHeaders.size <= 1;
    const parentText = String(groupCols[0] || '').trim();

    return (
      <>
        <tr>
          {left.map((h, i) => (
            <th key={`formvi-left-${i}`} rowSpan={2} style={identityOrDetailHeaderStyle} title={h}>
              {h}
            </th>
          ))}
          <th key="formvi-group" colSpan={groupCols.length} style={groupHeaderStyle} title={parentText}>
            {parentText}
          </th>
          {right.map((h, i) => (
            <th key={`formvi-right-${i}`} rowSpan={2} style={identityOrDetailHeaderStyle} title={h}>
              {h}
            </th>
          ))}
        </tr>
        <tr>
          {groupCols.map((h, i) => (
            <th key={`formvi-sub-${i}`} style={stickyDetailHeaderStyle} title={h}>
              {showIndexedSubHeaders ? String(i + 1) : formatStatutoryTableHeaderLabel(h)}
            </th>
          ))}
        </tr>
      </>
    );
  }, [displayTableHeaders, displayFormHeader, formFileModalData?.item, formFileModalData?.fileName]);

  // Formmaster templates by normalized form name: for rows without a form file (e.g. ChecklistBulk-only), show Formmaster file when form name matches
  const formmasterFormFileByFormName = useMemo(() => {
    const map = new Map();
    if (!formmasterTemplates || !Array.isArray(formmasterTemplates)) return map;
    const templateRecencyScore = (t) => {
      const createdRaw = t?.createdTime || t?.created_time || t?.createdAt || t?.created_at || '';
      const createdMs = createdRaw ? Date.parse(String(createdRaw)) : NaN;
      if (Number.isFinite(createdMs)) return createdMs;
      const idNum = Number.parseInt(String(t?.id ?? ''), 10);
      return Number.isFinite(idNum) ? idNum : 0;
    };
    formmasterTemplates.forEach((t) => {
      const rawName = (t.name || t.file_name || t.fileName || '').trim();
      const formName = rawName.replace(/\.xlsx$/i, '').trim() || rawName;
      const key = normalizeTemplateFormNameKey(formName);
      if (!key) return;
      const candidate = {
        formFile: t.id != null ? String(t.id) : null,
        formFileName: rawName || 'template.xlsx',
        isFromFormmaster: true,
        _score: templateRecencyScore(t)
      };
      const current = map.get(key);
      if (!current || candidate._score >= (current._score || 0)) {
        map.set(key, candidate);
      }
    });
    return map;
  }, [formmasterTemplates]);

  return (
    <div className="statutory-page">
      {!isFormOpen ? (
        <>
          <header className="statutory-page-heading">
            <h1 className="statutory-page-title">Statutory</h1>
            <nav className="statutory-breadcrumb" aria-label="Breadcrumb">
              <Link to="/hcm-dashboard" className="statutory-bc-link">
                Dashboard
              </Link>
              <span className="statutory-bc-sep" aria-hidden>
                &gt;
              </span>
              <span className="statutory-bc-muted">Transaction</span>
              <span className="statutory-bc-sep" aria-hidden>
                &gt;
              </span>
              <span className="statutory-bc-current">Statutory</span>
            </nav>
          </header>
          {/* Card Container */}
          <div className="statutory-card-container">
            {/* Header Section */}
            <div className="statutory-inner-toolbar">
              {/* Transaction Filters Row */}
              <div className="statutory-actions-card statutory-actions-card--table">
                <input
                  type="text"
                  value={tableSearch}
                  onChange={(e) => setTableSearch(e.target.value)}
                  className="statutory-table-top-search"
                  placeholder="Search by Act or Form name..."
                  title="Search"
                />
                <select
                  value={selectedFormFilter}
                  onChange={(e) => setSelectedFormFilter(e.target.value)}
                  className="statutory-table-top-select"
                  title="All Forms"
                >
                  <option value="all">All Forms</option>
                  {tableFormOptions.map((formName) => (
                    <option key={formName} value={formName}>{formName}</option>
                  ))}
                </select>
                <select
                  value={selectedMonth}
                  onChange={(e) => {
                    const v = e.target.value;
                    setSelectedMonth(v);
                    persistMonthFilter(v);
                  }}
                  className="statutory-table-top-select"
                  title="Filter by Month"
                >
                  <option value="January">Jan</option>
                  <option value="February">Feb</option>
                  <option value="March">Mar</option>
                  <option value="April">Apr</option>
                  <option value="May">May</option>
                  <option value="June">Jun</option>
                  <option value="July">Jul</option>
                  <option value="August">Aug</option>
                  <option value="September">Sep</option>
                  <option value="October">Oct</option>
                  <option value="November">Nov</option>
                  <option value="December">Dec</option>
                </select>
                <button
                  type="button"
                  className="statutory-table-top-btn"
                  onClick={clearTransactionFilters}
                  title="Clear Filters"
                >
                  Clear Filters
                </button>
              </div>
            </div>

            {/* Table Container */}
            <div className="statutory-table-container">
              {/* Error Messages */}
              {error && (
                <div style={{
                  margin: '10px 0',
                  padding: '12px',
                  backgroundColor: '#fef2f2',
                  border: '1px solid #fecaca',
                  borderRadius: '6px',
                  color: '#dc2626',
                  fontSize: '14px'
                }}>
                  <strong>Error:</strong> {error}
                </div>
              )}
              {success && (
                <div style={{
                  margin: '10px 0',
                  padding: '12px',
                  backgroundColor: '#f0fdf4',
                  border: '1px solid #bbf7d0',
                  borderRadius: '6px',
                  color: '#16a34a',
                  fontSize: '14px'
                }}>
                  <strong>Success:</strong> {success}
                </div>
              )}

              <div className="statutory-table-wrapper">
                <table className="statutory-table">
                  <thead>
                    <tr>
                      <th className="statutory-col-checkbox" style={{ width: '24px', textAlign: 'center', padding: '0 2px' }}>
                        <input
                          type="checkbox"
                          checked={pagedTransactionTableData.length > 0 && pagedTransactionTableData.every((row) => selectedRows.has(row.id))}
                          onChange={(e) => {
                            if (e.target.checked) {
                              setSelectedRows(new Set(pagedTransactionTableData.map((row) => row.id)));
                            } else {
                              setSelectedRows(new Set());
                            }
                          }}
                        />
                      </th>
                      <th className="statutory-col-sno" style={{ width: '48px', textAlign: 'center', padding: '0 4px' }}>S.No</th>
                      {showSiteColumn ? <th>Site</th> : null}
                      <th
                        className="table-col-act statutory-col-left"
                        style={{ textAlign: 'left', width: '200px', minWidth: '200px', maxWidth: '200px' }}
                      >
                        Act
                      </th>
                      <th
                        className="table-col-description"
                        style={{ width: '220px', minWidth: '220px', maxWidth: '220px' }}
                      >
                        Description
                      </th>
                      <th>Form Name</th>
                      <th>Form File</th>
                      <th>Due Date</th>
                      <th>Month Filter</th>
                      <th>Autofill</th>
                      <th>Draft</th>
                      {showSendForApprovalColumn ? <th className="statutory-col-send-for-approval">Send for approval</th> : null}
                      {showApprovalColumn ? <th className="statutory-col-approval">Approval</th> : null}
                      <th>Status</th>
                      {showRemarksColumn ? <th>Remarks</th> : null}
                    </tr>
                  </thead>
                  <tbody>
                    {pagedTransactionTableData.length === 0 ? (
                      <tr>
                        <td colSpan={String(tableColSpan)} style={{ textAlign: 'center', padding: '40px', color: '#6b7280' }}>
                          {statutoryData.length === 0
                            ? hasSiteBasedActScope
                                ? `No statutory records for your Site industry scope (${allowedActCategoryList.join(', ')}). Click "Add New Statutory" to create one.`
                                : siteWiseCategory
                                  ? `No statutory records found for this sector (${siteFromUrl}). Click "Add New Statutory" to create one.`
                                  : 'No statutory records found. Click "Add New Statutory" to create one.'
                            : hasSiteBasedActScope || siteWiseCategory
                              ? `No statutory records for this sector match the selected month (${selectedMonth}). Try another month.`
                              : 'No statutory records match the selected month filter.'}
                        </td>
                      </tr>
                    ) : (
                      pagedTransactionTableData.map((item, index) => {
                          const isSelected = selectedRows.has(item.id);
                          const rowLocked = isRowStatusApproved(item);
                          const rowMonth = item.monthFilter || item.MonthFilter || item.monthfilter || '';
                          const selectedMonthNorm = (selectedMonth || '').trim().toLowerCase().substring(0, 3);
                          const rowMonthNorm = String(rowMonth).trim().toLowerCase().substring(0, 3);
                          const rowDraftFile = item.draftFile ?? item.DraftFile ?? null;
                          const rowHasDraft = rowDraftFile && rowDraftFile !== 'null' && String(rowDraftFile).trim() !== '';
                          const secPart = String(item.sector || item.Sector || '').trim().toLowerCase();
                          const statePart = String(item.state || item.State || '').trim().toLowerCase();
                          const descPart = String(item.description || item.Description || '').trim().toLowerCase();
                          const draftLookupKey = `${String(item.formName || '').toLowerCase().trim()}|${String(item.act || '').toLowerCase().trim()}|${descPart}|${(rowMonthNorm || selectedMonthNorm || 'nomonth')}|${secPart}|${statePart}`;
                          let fallbackDraftRow = draftRowByFormActMonthKey.get(draftLookupKey) || null;
                          // Autofill → Save POST creates a numeric statutory row with DraftFile while the grid row may still be bulk_* / checklist_* — pick sibling by line identity.
                          if (!fallbackDraftRow && !rowHasDraft && Array.isArray(statutoryData)) {
                            const fn = (x) => String(x?.formName || x?.FormName || '').trim().toLowerCase().replace(/\s+/g, ' ');
                            const ac = (x) => String(x?.act || x?.Act || '').trim().toLowerCase();
                            const ds = (x) => String(x?.description || x?.Description || '').trim().toLowerCase();
                            const sc = (x) => String(x?.sector || x?.Sector || '').trim().toLowerCase();
                            const st = (x) => String(x?.state || x?.State || '').trim().toLowerCase();
                            fallbackDraftRow =
                              statutoryData.find((r) => {
                                if (!/^\d+$/.test(String(r?.id ?? ''))) return false;
                                const df = r?.draftFile ?? r?.DraftFile;
                                if (!df || df === 'null' || String(df).trim() === '') return false;
                                return (
                                  fn(r) === fn(item) &&
                                  ac(r) === ac(item) &&
                                  ds(r) === ds(item) &&
                                  sc(r) === sc(item) &&
                                  st(r) === st(item)
                                );
                              }) || null;
                          }
                          // Use the same statutory DB sibling key as Send for approval / proof — avoids drafts staying hidden
                          // for admin until metadata lines up (strict fn/act/description match above is often stricter).
                          if (!fallbackDraftRow && !rowHasDraft && Array.isArray(statutoryData)) {
                            const sid = resolveNumericStatutoryIdForProofRow(item);
                            if (sid) {
                              const r = statutoryData.find((row) => String(row.id) === sid);
                              const df = r?.draftFile ?? r?.DraftFile;
                              if (r && df && df !== 'null' && String(df).trim() !== '') {
                                fallbackDraftRow = r;
                              }
                            }
                          }
                          const draftSourceItem = rowHasDraft ? item : fallbackDraftRow;
                          const dsDraft = draftSourceItem?.draftFile ?? draftSourceItem?.DraftFile;
                          const draftSourceHasFile = !!(dsDraft && dsDraft !== 'null' && String(dsDraft).trim() !== '');
                          const draftStoredMonthRaw =
                            draftSourceItem?.monthFilter ?? draftSourceItem?.MonthFilter ?? draftSourceItem?.monthfilter;
                          const draftStoredMonthNorm =
                            draftStoredMonthRaw != null && String(draftStoredMonthRaw).trim() !== ''
                              ? (() => {
                                  const full = resolveToFullMonthName(String(draftStoredMonthRaw).trim());
                                  return (full || String(draftStoredMonthRaw).trim()).toLowerCase().substring(0, 3);
                                })()
                              : null;
                          const uiMonthSelNorm = statutoryDedupeMonthNorm(
                            { monthFilter: selectedMonth, MonthFilter: selectedMonth },
                            selectedMonth
                          );
                          const hideDraftByStoredMonth =
                            draftStoredMonthNorm != null &&
                            uiMonthSelNorm !== 'nomonth' &&
                            draftStoredMonthNorm !== uiMonthSelNorm;
                          const sendForApprovalEffective = (() => {
                            for (const r of [item, draftSourceItem, fallbackDraftRow].filter(Boolean)) {
                              const v = r.sendForApproval ?? r.SendForApproval;
                              if (v != null && String(v).trim() !== '') return String(v).trim();
                            }
                            const nid = resolveNumericStatutoryIdForProofRow(item);
                            if (nid && Array.isArray(statutoryData)) {
                              const nr = statutoryData.find((row) => String(row.id) === nid);
                              const v = nr?.sendForApproval ?? nr?.SendForApproval;
                              if (v != null && String(v).trim() !== '') return String(v).trim();
                            }
                            // Fallback for month-mismatch rows (e.g. UI month differs from saved row month):
                            // read SendForApproval from any numeric sibling of the same statutory line.
                            if (Array.isArray(statutoryData)) {
                              const formNorm = baseFormNameKey(item?.formName || item?.FormName);
                              const secNorm = String(item?.sector || item?.Sector || '').trim().toLowerCase();
                              const stateNorm = String(item?.state || item?.State || '').trim().toLowerCase();
                              const siblings = statutoryData.filter((row) => {
                                if (!isNumericStatutoryBackendId(row?.id)) return false;
                                if (baseFormNameKey(row?.formName || row?.FormName) !== formNorm) return false;
                                if (String(row?.sector || row?.Sector || '').trim().toLowerCase() !== secNorm) return false;
                                if (String(row?.state || row?.State || '').trim().toLowerCase() !== stateNorm) return false;
                                return true;
                              });
                              const sentSibling = siblings.find((row) => /^sent$/i.test(String(row?.sendForApproval ?? row?.SendForApproval ?? '').trim()));
                              if (sentSibling) return 'Sent';
                              const anySibling = siblings.find((row) => {
                                const v = row?.sendForApproval ?? row?.SendForApproval;
                                return v != null && String(v).trim() !== '';
                              });
                              if (anySibling) return String(anySibling.sendForApproval ?? anySibling.SendForApproval).trim();

                              // Site-view fallback: in site logins, sector/state may be blank or mismatched in bulk placeholders.
                              // Use same site + base form to mirror already-sent status.
                              const itemSiteNorm = String(resolveSiteDisplayName(item, statutoryData) || item?.site || item?.Site || '')
                                .trim()
                                .toLowerCase();
                              if (itemSiteNorm) {
                                const siteSiblings = statutoryData.filter((row) => {
                                  if (!isNumericStatutoryBackendId(row?.id)) return false;
                                  if (baseFormNameKey(row?.formName || row?.FormName) !== formNorm) return false;
                                  const rowSiteNorm = String(resolveSiteDisplayName(row, statutoryData) || row?.site || row?.Site || '')
                                    .trim()
                                    .toLowerCase();
                                  return resolvedSitesOverlap(rowSiteNorm, itemSiteNorm);
                                });
                                const siteSent = siteSiblings.find((row) =>
                                  /^sent$/i.test(String(row?.sendForApproval ?? row?.SendForApproval ?? '').trim())
                                );
                                if (siteSent) return 'Sent';
                              }
                            }
                            return '';
                          })();
                          // Approval-flow gating: in approver view, show draft only after site marks Send for Approval as Sent.
                          const hideDraftUntilSiteSent =
                            showApprovalColumn &&
                            !/^sent$/i.test(sendForApprovalEffective);
                          const showDraftForRow =
                            !!draftSourceHasFile && !hideDraftUntilSiteSent && !hideDraftByStoredMonth;
                          // Resolved form file item: same source as "View File" so Autofill uses the same Excel format
                          const rowFormFile = item.formFile || item.FormFile || null;
                          const rowFormFileName = item.formFileName || item.FormFileName || null;
                          const rowHasFormFile = rowFormFile && rowFormFile !== 'null' && rowFormFile !== null && String(rowFormFile).trim() !== '';
                          const rowFormKey = `${(item.formName || '').toLowerCase().trim()}|${(item.act || '').toLowerCase().trim()}|${secPart}|${statePart}`;
                          const rowLookup = formFileByFormKey.get(rowFormKey);
                          const rowNormalizedFormName = normalizeTemplateFormNameKey(item.formName || item.FormName || '');
                          const rowFormmasterFile = formmasterFormFileByFormName.get(rowNormalizedFormName);
                          let resolvedFormFileItem = rowHasFormFile
                            ? {
                              ...item,
                              formFile: rowFormFile,
                              formFileName: rowFormFileName || `${((item.formName || 'Form') + '').trim().replace(/[/\\?%*:|"<>]/g, '') || 'Form'}.xlsx`
                            }
                            : rowLookup
                              ? { ...item, formFile: rowLookup.formFile, formFileName: rowLookup.formFileName, id: rowLookup.id }
                              : rowFormmasterFile && rowFormmasterFile.formFile
                                ? { ...item, formFile: rowFormmasterFile.formFile, formFileName: rowFormmasterFile.formFileName, isFromFormmaster: true }
                                : item;
                          // Saved draft lives on API row (numeric id); merged grid row may still be bulk_* — wire draft + id for View Draft / Autofill.
                          if (
                            !hideDraftUntilSiteSent &&
                            !hideDraftByStoredMonth &&
                            fallbackDraftRow &&
                            draftSourceItem &&
                            String(draftSourceItem.id ?? '') !== String(item.id ?? '')
                          ) {
                            resolvedFormFileItem = {
                              ...resolvedFormFileItem,
                              id: draftSourceItem.id,
                              draftFile: draftSourceItem.draftFile ?? draftSourceItem.DraftFile ?? resolvedFormFileItem.draftFile,
                              draftFileName:
                                draftSourceItem.draftFileName ??
                                draftSourceItem.DraftFileName ??
                                resolvedFormFileItem.draftFileName
                            };
                          }
                          if (
                            !hideDraftUntilSiteSent &&
                            !hideDraftByStoredMonth &&
                            item.draftStatutoryRowIdForFile &&
                            !isNumericStatutoryBackendId(item?.id)
                          ) {
                            resolvedFormFileItem = {
                              ...resolvedFormFileItem,
                              id: item.draftStatutoryRowIdForFile,
                              draftFile: item.draftFile ?? item.DraftFile ?? resolvedFormFileItem.draftFile,
                              draftFileName:
                                item.draftFileName ??
                                item.DraftFileName ??
                                resolvedFormFileItem.draftFileName
                            };
                          }
                          return (
                            <tr key={item.id || index} className={rowLocked ? 'statutory-row-approved' : undefined}>
                              <td className="statutory-col-checkbox" style={{ textAlign: 'center', padding: '4px' }}>
                                <input
                                  type="checkbox"
                                  checked={isSelected}
                                  onChange={(e) => {
                                    e.stopPropagation();
                                    handleCheckboxChange(item.id);
                                  }}
                                  style={{ width: '14px', height: '14px', cursor: 'pointer' }}
                                />
                              </td>
                            <td className="statutory-col-sno" style={{ width: '48px', textAlign: 'center', fontWeight: '500', padding: '4px' }}>{((effectivePage - 1) * PAGE_SIZE) + index + 1}</td>
                            {showSiteColumn ? (
                              <td>{resolveSiteDisplayName(item, statutoryData) || '-'}</td>
                            ) : null}
                            <td
                              className="table-col-act statutory-col-left"
                              style={{
                                textAlign: 'left',
                                width: '200px',
                                minWidth: '200px',
                                maxWidth: '200px'
                              }}
                            >
                              <div className="statutory-act-clamp-text">
                                {item.act || item.Act || item.sector || item.Sector || '-'}
                              </div>
                            </td>
                            <td
                              className="table-col-description"
                              style={{ width: '220px', minWidth: '220px', maxWidth: '220px' }}
                            >
                              {item.description || item.Description || item.formName || '-'}
                            </td>
                            <td><span className="statutory-pill statutory-pill-form">{item.formName || '-'}</span></td>
                            <td>
                              {(() => {
                                if (resolvedFormFileItem.formFile && resolvedFormFileItem.formFile !== 'null' && String(resolvedFormFileItem.formFile).trim() !== '') {
                                  const formNameNorm = (item.formName || '').trim().toUpperCase().replace(/\s+/g, ' ');
                                  const fileNameNorm = (resolvedFormFileItem.formFileName || '').trim().toUpperCase().replace(/\s+/g, ' ');
                                  const fileNameRedundantWithFormName = formNameNorm && fileNameNorm && fileNameNorm.startsWith(formNameNorm);
                                  const formFileLabel = fileNameRedundantWithFormName ? 'View File' : (resolvedFormFileItem.formFileName || 'View File');
                                  const fileUrl = getFormFileDownloadUrl(resolvedFormFileItem);
                                  return (
                                    <a
                                      href={fileUrl || '#'}
                                      target="_blank"
                                      rel="noopener noreferrer"
                                      onClick={(e) => e.stopPropagation()}
                                      style={{
                                        color: '#3b82f6',
                                        textDecoration: 'underline',
                                        cursor: 'pointer',
                                        fontWeight: '500'
                                      }}
                                      title="Open Excel file in new tab"
                                    >
                                      <span className="statutory-file-link">View File</span>
                                    </a>
                                  );
                                }
                                return 'No File';
                              })()}
                            </td>
                            <td>{formatMonthlyBasisDate(item.dueDate) || '-'}</td>
                            <td style={{ textAlign: 'center', whiteSpace: 'nowrap' }}>
                              {(() => {
                                const m =
                                  selectedMonth ||
                                  item.monthFilter ||
                                  item.MonthFilter ||
                                  item.monthfilter ||
                                  getMonthFromDueDate(item.dueDate) ||
                                  '';
                                const abbrev = m ? (m.length > 3 ? m.substring(0, 3) : m) : null;
                                return abbrev != null && abbrev !== '' ? <span className="statutory-pill statutory-pill-month">{abbrev}</span> : '-';
                              })()}
                            </td>
                            <td style={{ textAlign: 'center', padding: '8px' }}>
                              {(() => {
                                const hasFormFileForRow = resolvedFormFileItem.formFile && resolvedFormFileItem.formFile !== 'null' && String(resolvedFormFileItem.formFile).trim() !== '';
                                return (
                                  <button
                                    className="statutory-autofill-btn-ui"
                                    onClick={(e) => {
                                      e.stopPropagation();
                                      if (hasFormFileForRow && !rowLocked) handleAutofill(resolvedFormFileItem);
                                    }}
                                    title={
                                      rowLocked
                                        ? 'Approved row is locked for editing'
                                        : hasFormFileForRow
                                          ? 'Autofill using same form format as View File'
                                          : 'No form file – upload or link a form to enable Autofill'
                                    }
                                    disabled={!hasFormFileForRow || rowLocked}
                                  >
                                    Autofill
                                  </button>
                                );
                              })()}
                            </td>
                            <td>
                              {(() => {
                                const draftTextValue = String(item?.draft ?? item?.Draft ?? '').trim();
                                const hasDraftTextMarker =
                                  draftTextValue !== '' &&
                                  draftTextValue !== '-' &&
                                  draftTextValue.toLowerCase() !== 'null';
                                const showDraftActions =
                                  !hideDraftUntilSiteSent &&
                                  !hideDraftByStoredMonth &&
                                  (showDraftForRow || (hasDraftTextMarker && isNumericStatutoryBackendId(item?.id)));
                                return showDraftActions ? (
                                (() => {
                                  const formNameNorm = (item.formName || '').trim().toUpperCase().replace(/\s+/g, ' ');
                                  const draftNameNorm = String(draftSourceItem?.draftFileName || draftSourceItem?.DraftFileName || '').trim().toUpperCase().replace(/\s+/g, ' ');
                                  const draftLabelRedundant = formNameNorm && draftNameNorm && draftNameNorm.startsWith(formNameNorm);
                                  const draftLinkLabel = draftLabelRedundant ? 'View Draft File' : (draftSourceItem?.draftFileName || 'View Draft File');
                                  const draftApiRowId =
                                    item.draftStatutoryRowIdForFile ??
                                    (isNumericStatutoryBackendId(draftSourceItem?.id) ? draftSourceItem.id : null) ??
                                    (isNumericStatutoryBackendId(item?.id) ? item.id : null);
                                  const draftDownloadUrl = draftApiRowId
                                    ? `/server/statutoryreg_function/statutory/${draftApiRowId}/file/Draft`
                                    : '#';
                                  const draftDownloadFileName =
                                    draftSourceItem?.draftFileName ||
                                    draftSourceItem?.DraftFileName ||
                                    resolvedFormFileItem.formFileName ||
                                    'form-template-with-data.xlsx';
                                  const hasFormFileForDraft = resolvedFormFileItem?.formFile && resolvedFormFileItem.formFile !== 'null' && String(resolvedFormFileItem.formFile).trim() !== '';
                                  const canOpenDraftEditModal = hasFormFileForDraft && !rowLocked;
                                  return (
                                    <span style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
                                      <a
                                        href="#"
                                        style={{ color: '#3b82f6', textDecoration: 'underline' }}
                                        title={canOpenDraftEditModal ? 'Open Autofill modal for this draft' : 'View Draft File'}
                                        onClick={(e) => {
                                          e.preventDefault();
                                          e.stopPropagation();
                                          if (canOpenDraftEditModal) handleAutofill(resolvedFormFileItem, { preferSavedDraftData: true });
                                        }}
                                      >
                                        {draftLinkLabel}
                                      </a>
                                      <a
                                        href={draftDownloadUrl}
                                        download={draftDownloadFileName}
                                        style={{ color: '#059669', textDecoration: 'underline', fontSize: '12px' }}
                                        title="Download with same data as View Draft File"
                                        onClick={(e) => {
                                          e.preventDefault();
                                          handleViewDraftFileGenerate(draftSourceItem || item, resolvedFormFileItem);
                                        }}
                                      >
                                        Download form template
                                      </a>
                                    </span>
                                  );
                                })()
                              ) : (
                                hideDraftUntilSiteSent || hideDraftByStoredMonth ? '-' : item.draft || '-'
                              );
                              })()}
                            </td>
                            {showSendForApprovalColumn ? (
                              <td className="statutory-col-send-for-approval" style={{ textAlign: 'center' }}>
                                {(() => {
                                  const sfaStr =
                                    sendForApprovalEffective != null && String(sendForApprovalEffective).trim() !== ''
                                      ? String(sendForApprovalEffective).trim()
                                      : '';
                                  const looksSent = /^sent$/i.test(sfaStr);
                                  // Rejected rows in site view must allow re-send for approval.
                                  const approvalNorm = String(item?.approval ?? item?.Approval ?? '').trim().toLowerCase();
                                  const statusNorm = String(item?.status ?? item?.Status ?? '').trim().toLowerCase();
                                  const rejectedByApprover =
                                    approvalNorm === 'rejected' ||
                                    approvalNorm === 'reject' ||
                                    statusNorm === 'rejected' ||
                                    statusNorm === 'reject';
                                  const showAsSent = looksSent && !rejectedByApprover;
                                  const numericId = resolveNumericStatutoryIdForProofRow(item);
                                  const sfaBusy = numericId && sendForApprovalUpdatingRowId === numericId;
                                  return (
                                    <div onClick={(e) => e.stopPropagation()}>
                                      <button
                                        type="button"
                                        className={`statutory-send-for-approval-btn${showAsSent ? ' statutory-send-for-approval-btn--done' : ''}`}
                                        disabled={!numericId || !!sfaBusy || loading || showAsSent || rowLocked}
                                        title={
                                          !numericId
                                            ? 'Only for rows stored in Statutory (numeric ID)'
                                            : rowLocked
                                              ? 'Approved row is locked for editing'
                                            : showAsSent
                                              ? 'Already sent for approval'
                                              : 'Set Send for approval to Sent'
                                        }
                                        onClick={(e) => {
                                          e.stopPropagation();
                                          if (!numericId || sfaBusy || showAsSent || rowLocked) return;
                                          handleSendForApprovalColumnAction(item);
                                        }}
                                      >
                                        {sfaBusy ? '…' : showAsSent ? 'Sent' : 'Send for approval'}
                                      </button>
                                      {sfaStr && !looksSent ? (
                                        <div
                                          className="statutory-sfa-inline-value"
                                          title={sfaStr}
                                        >
                                          {sfaStr.length > 18 ? `${sfaStr.slice(0, 18)}…` : sfaStr}
                                        </div>
                                      ) : null}
                                    </div>
                                  );
                                })()}
                              </td>
                            ) : null}
                            {showApprovalColumn ? (
                              <td className="statutory-col-approval" style={{ textAlign: 'center' }}>
                                {(() => {
                                  const rawA = item.approval ?? item.Approval;
                                  const aStr = rawA != null && String(rawA).trim() !== '' ? String(rawA).trim() : '';
                                  const aNorm = aStr.toLowerCase();
                                  const isApproved = aNorm === 'approved' || aNorm === 'approve';
                                  const isRejected = aNorm === 'rejected' || aNorm === 'reject';
                                  const hasOtherApproval = aStr && !isApproved && !isRejected;
                                  const numericId = resolveNumericStatutoryIdForProofRow(item);
                                  const rowBusy = numericId && approvalUpdatingRowId === numericId;
                                  return (
                                    <div
                                      className="statutory-approval-actions"
                                      onClick={(e) => e.stopPropagation()}
                                    >
                                      {hasOtherApproval ? (
                                        <div className="statutory-approval-other-hint" title={aStr}>
                                          {aStr.length > 24 ? `${aStr.slice(0, 24)}…` : aStr}
                                        </div>
                                      ) : null}
                                      <div className="statutory-approval-btns">
                                        <button
                                          type="button"
                                          className={`statutory-approval-btn statutory-approval-btn--approve${isApproved ? ' is-selected' : ''}`}
                                          disabled={!numericId || !!rowBusy || loading || isApproved || rowLocked}
                                          title={
                                            !numericId
                                              ? 'Only for rows stored in Statutory (numeric ID)'
                                              : rowLocked
                                                ? 'Approved row is locked for editing'
                                              : isApproved
                                                ? 'Already approved'
                                                : 'Set approval to Approved'
                                          }
                                          onClick={(e) => {
                                            e.stopPropagation();
                                            if (!numericId || rowBusy || rowLocked) return;
                                            handleApprovalColumnAction(item, 'Approved');
                                          }}
                                        >
                                          {rowBusy ? '…' : 'Approve'}
                                        </button>
                                        <button
                                          type="button"
                                          className={`statutory-approval-btn statutory-approval-btn--reject${isRejected ? ' is-selected' : ''}`}
                                          disabled={!numericId || !!rowBusy || loading || isRejected || rowLocked}
                                          title={
                                            !numericId
                                              ? 'Only for rows stored in Statutory (numeric ID)'
                                              : rowLocked
                                                ? 'Approved row is locked for editing'
                                              : isRejected
                                                ? 'Already rejected'
                                                : 'Set approval to Rejected'
                                          }
                                          onClick={(e) => {
                                            e.stopPropagation();
                                            if (!numericId || rowBusy || rowLocked) return;
                                            handleApprovalColumnAction(item, 'Rejected');
                                          }}
                                        >
                                          {rowBusy ? '…' : 'Reject'}
                                        </button>
                                      </div>
                                    </div>
                                  );
                                })()}
                              </td>
                            ) : null}
                            <td style={{ textAlign: 'center' }}>
                              {(() => {
                                const draftVal = item.draftFile ?? item.DraftFile ?? null;
                                const hasDraftFile =
                                  draftVal != null &&
                                  String(draftVal).trim() !== '' &&
                                  String(draftVal).trim() !== 'null' &&
                                  String(draftVal).trim() !== 'undefined';
                                if (!hasDraftFile) {
                                  return <span className="statutory-status-btn statutory-status-btn--yet-to-complete">Yet to Complete</span>;
                                }
                                const raw = item.status != null && item.status !== '' ? item.status : item.Status;
                                const norm = String(raw || '').trim();
                                const normalized = norm.toLowerCase();
                                if (norm === '' || norm === '-' || norm === '—') {
                                  return <span className="statutory-status-btn statutory-status-btn--pending">Pending</span>;
                                }
                                const statusClass =
                                  normalized === 'approved'
                                    ? 'statutory-status-btn--approved'
                                    : normalized === 'rejected' || normalized === 'reject'
                                      ? 'statutory-status-btn--rejected'
                                      : normalized === 'pending'
                                        ? 'statutory-status-btn--pending'
                                      : 'statutory-status-btn--default';
                                return <span className={`statutory-status-btn ${statusClass}`}>{norm}</span>;
                              })()}
                            </td>
                            {showRemarksColumn ? (
                              <td>{renderRemarksDisplay(item.remarks || item.Remarks || '')}</td>
                            ) : null}
                          </tr>
                          );
                      })
                    )}
                  </tbody>
                </table>
              </div>
              <div className="statutory-table-footer">
                <nav className="statutory-pagination" aria-label="Table pagination">
                  <button
                    type="button"
                    className="statutory-pagination-nav"
                    disabled={effectivePage <= 1}
                    onClick={() => setCurrentPage(1)}
                    title="First page"
                    aria-label="First page"
                  >
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden>
                      <polyline points="11 17 6 12 11 7" />
                      <polyline points="18 17 13 12 18 7" />
                    </svg>
                  </button>
                  <button
                    type="button"
                    className="statutory-pagination-nav"
                    disabled={effectivePage <= 1}
                    onClick={() => setCurrentPage((p) => Math.max(1, p - 1))}
                    title="Previous page"
                    aria-label="Previous page"
                  >
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden>
                      <polyline points="15 18 9 12 15 6" />
                    </svg>
                  </button>
                  <div className="statutory-pagination-pages">
                    {paginationItems.map((item, i) =>
                      item === 'ellipsis' ? (
                        <span key={`e-${i}`} className="statutory-pagination-ellipsis" aria-hidden>
                          ...
                        </span>
                      ) : (
                        <button
                          key={item}
                          type="button"
                          className={`statutory-pagination-page${item === effectivePage ? ' statutory-pagination-page--active' : ''}`}
                          onClick={() => setCurrentPage(item)}
                          aria-label={`Page ${item}`}
                          aria-current={item === effectivePage ? 'page' : undefined}
                        >
                          {item}
                        </button>
                      )
                    )}
                  </div>
                  <button
                    type="button"
                    className="statutory-pagination-nav"
                    disabled={effectivePage >= totalPages}
                    onClick={() => setCurrentPage((p) => Math.min(totalPages, p + 1))}
                    title="Next page"
                    aria-label="Next page"
                  >
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden>
                      <polyline points="9 18 15 12 9 6" />
                    </svg>
                  </button>
                  <button
                    type="button"
                    className="statutory-pagination-nav"
                    disabled={effectivePage >= totalPages}
                    onClick={() => setCurrentPage(totalPages)}
                    title="Last page"
                    aria-label="Last page"
                  >
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden>
                      <polyline points="13 17 18 12 13 7" />
                      <polyline points="6 17 11 12 6 7" />
                    </svg>
                  </button>
                </nav>
              </div>
            </div>
          </div>
          {isRejectRemarksModalOpen ? (
            <div
              style={{
                position: 'fixed',
                inset: 0,
                background: 'rgba(0,0,0,0.35)',
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center',
                zIndex: 1200
              }}
              onClick={closeRejectRemarksModal}
            >
              <div
                style={{
                  width: '520px',
                  maxWidth: 'calc(100vw - 32px)',
                  background: '#fff',
                  borderRadius: '12px',
                  boxShadow: '0 10px 30px rgba(0,0,0,0.2)',
                  padding: '16px'
                }}
                onClick={(e) => e.stopPropagation()}
              >
                <h3 style={{ margin: '0 0 10px 0', fontSize: '18px', color: '#111827' }}>Reject Remarks</h3>
                <p style={{ margin: '0 0 8px 0', fontSize: '13px', color: '#6b7280' }}>
                  Enter remarks. Press <b>Shift+Enter</b> for a new line. Press <b>Enter</b> to submit.
                </p>
                <textarea
                  value={rejectRemarksValue}
                  onChange={(e) => setRejectRemarksValue(e.target.value)}
                  onKeyDown={(e) => {
                    if (e.key === 'Enter' && !e.shiftKey) {
                      e.preventDefault();
                      submitRejectRemarks();
                    }
                  }}
                  autoFocus
                  rows={5}
                  placeholder="Type reject remarks..."
                  style={{
                    width: '100%',
                    border: '1px solid #d1d5db',
                    borderRadius: '8px',
                    padding: '10px 12px',
                    fontSize: '14px',
                    outline: 'none',
                    resize: 'vertical',
                    boxSizing: 'border-box'
                  }}
                />
                <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '8px', marginTop: '12px' }}>
                  <button
                    type="button"
                    onClick={closeRejectRemarksModal}
                    style={{
                      border: '1px solid #d1d5db',
                      background: '#fff',
                      color: '#374151',
                      borderRadius: '8px',
                      padding: '8px 14px',
                      cursor: 'pointer'
                    }}
                  >
                    Cancel
                  </button>
                  <button
                    type="button"
                    disabled={
                      !rejectRemarksTarget?.targetId ||
                      approvalUpdatingRowId === rejectRemarksTarget?.targetId
                    }
                    onClick={submitRejectRemarks}
                    style={{
                      border: 'none',
                      background: '#b45309',
                      color: '#fff',
                      borderRadius: '8px',
                      padding: '8px 14px',
                      cursor: 'pointer',
                      opacity:
                        !rejectRemarksTarget?.targetId ||
                        approvalUpdatingRowId === rejectRemarksTarget?.targetId
                          ? 0.65
                          : 1
                    }}
                  >
                    {approvalUpdatingRowId === rejectRemarksTarget?.targetId ? 'Saving...' : 'Save'}
                  </button>
                </div>
              </div>
            </div>
          ) : null}
        </>
      ) : (
        <div className="statutory-forms-container">
          <div className="statutory-card">
            <div className="statutory-card-header">
              <h2>{editingStatutoryId ? 'Edit Statutory' : 'Add Statutory'}</h2>
              <button className="statutory-close-btn" onClick={handleClose} title="Close">
                &times;
              </button>
            </div>
            <form className="statutory-form" onSubmit={onSubmit}>
              <div className="statutory-form-row">
                <div className="statutory-field">
                  <label htmlFor="act">Act</label>
                  <input
                    id="act"
                    name="act"
                    className="input"
                    value={form.act}
                    onChange={onChange}
                    placeholder="Enter act"
                  />
                </div>
                <div className="statutory-field">
                  <label htmlFor="formName">Form Name <span className="req">*</span></label>
                  <input
                    id="formName"
                    name="formName"
                    className="input"
                    value={form.formName}
                    onChange={onChange}
                    placeholder="Enter form name"
                  />
                </div>
              </div>
              <div className="statutory-row">
                <label htmlFor="description">Description</label>
                <textarea
                  id="description"
                  name="description"
                  className="input"
                  value={form.description}
                  onChange={onChange}
                  placeholder="Enter description"
                  rows="3"
                  style={{
                    height: 'auto',
                    minHeight: '80px',
                    resize: 'vertical',
                    padding: '12px 16px'
                  }}
                />
              </div>
              <div className="statutory-form-row">
                <div className="statutory-field">
                  <label htmlFor="sector">Sector</label>
                  <input
                    id="sector"
                    name="sector"
                    className="input"
                    value={form.sector}
                    onChange={onChange}
                    placeholder="Enter sector"
                  />
                </div>
                <div className="statutory-field">
                  <label htmlFor="state">State</label>
                  <input
                    id="state"
                    name="state"
                    className="input"
                    value={form.state}
                    onChange={onChange}
                    placeholder="Enter state"
                  />
                </div>
              </div>
              <div className="statutory-row">
                <label htmlFor="site">Site</label>
                <input
                  id="site"
                  name="site"
                  className="input"
                  value={form.site}
                  onChange={onChange}
                  placeholder="Enter site"
                />
              </div>
              <div className="statutory-form-row">
                <div className="statutory-field">
                  <label htmlFor="formFile">Form File</label>
                  {form.formFile && !(form.formFile instanceof File) && (
                    <div style={{
                      marginBottom: '8px',
                      padding: '8px 12px',
                      backgroundColor: '#f0f9ff',
                      border: '1px solid #0ea5e9',
                      borderRadius: '6px',
                      fontSize: '14px',
                      color: '#0369a1',
                      display: 'flex',
                      alignItems: 'center',
                      justifyContent: 'space-between'
                    }}>
                      <span>Current File: {form.formFileName || 'View File'}</span>
                      <a
                        href={`/server/statutoryreg_function/statutory/${editingStatutoryId || form.id}/file/Form`}
                        target="_blank"
                        rel="noopener noreferrer"
                        style={{ color: '#0ea5e9', textDecoration: 'underline', marginLeft: '8px' }}
                      >
                        View
                      </a>
                    </div>
                  )}
                  <input
                    id="formFile"
                    name="formFile"
                    type="file"
                    accept=".pdf,.doc,.docx,.jpg,.jpeg,.png"
                    onChange={onChange}
                    style={{
                      color: 'transparent',
                      '::file-selector-button': {
                        display: 'none'
                      }
                    }}
                  />
                  {form.formFile instanceof File && form.formFileName && (
                    <div style={{
                      marginTop: '8px',
                      padding: '8px 12px',
                      backgroundColor: '#f0f9ff',
                      border: '1px solid #0ea5e9',
                      borderRadius: '6px',
                      fontSize: '14px',
                      color: '#0369a1'
                    }}>
                      New File Selected: {form.formFileName}
                    </div>
                  )}
                  <small style={{ color: '#6b7280', fontSize: '12px', marginTop: '4px', display: 'block' }}>
                    Maximum file size: 5MB. Supported formats: PDF, DOC, DOCX, JPG, PNG
                  </small>
                </div>
                <div className="statutory-field">
                  <label htmlFor="dueDate">Due Date</label>
                  <input
                    id="dueDate"
                    name="dueDate"
                    type="date"
                    className="input"
                    value={form.dueDate}
                    onChange={onChange}
                  />
                </div>
              </div>
              <div className="statutory-form-row">
                <div className="statutory-field">
                  <label htmlFor="autofill">Autofill</label>
                  <input
                    id="autofill"
                    name="autofill"
                    className="input"
                    value={form.autofill}
                    onChange={onChange}
                    placeholder="Enter autofill"
                  />
                </div>
                <div className="statutory-field">
                  <label htmlFor="draft">Draft</label>
                  <input
                    id="draft"
                    name="draft"
                    className="input"
                    value={form.draft}
                    onChange={onChange}
                    placeholder="Enter draft status"
                  />
                </div>
              </div>
              <div className="statutory-form-row">
                <div className="statutory-field">
                  <label htmlFor="sendForApproval">Send for approval</label>
                  <input
                    id="sendForApproval"
                    name="sendForApproval"
                    className="input"
                    value={form.sendForApproval}
                    onChange={onChange}
                    placeholder="e.g. Yes, No, Sent, Pending"
                  />
                </div>
                <div className="statutory-field">
                  <label htmlFor="remarks">Remarks</label>
                  <input
                    id="remarks"
                    name="remarks"
                    className="input"
                    value={form.remarks}
                    onChange={onChange}
                    placeholder="Enter remarks"
                  />
                </div>
              </div>
              <div className="statutory-form-row">
                <div className="statutory-field">
                  <label htmlFor="approval">Approval</label>
                  <input
                    id="approval"
                    name="approval"
                    className="input"
                    value={form.approval}
                    onChange={onChange}
                    placeholder="e.g. Pending, Approved, Rejected"
                  />
                </div>
                <div className="statutory-field">
                  <label htmlFor="status">Status</label>
                  <input
                    id="status"
                    name="status"
                    className="input"
                    value={form.status}
                    onChange={onChange}
                    placeholder="e.g. Open, Submitted, Closed"
                  />
                </div>
              </div>
              <div className="statutory-form-row">
                <div className="statutory-field">
                  <label htmlFor="draftFile">Draft File</label>
                  {form.draftFile && !(form.draftFile instanceof File) && (
                    <div style={{
                      marginBottom: '8px',
                      padding: '8px 12px',
                      backgroundColor: '#f0f9ff',
                      border: '1px solid #0ea5e9',
                      borderRadius: '6px',
                      fontSize: '14px',
                      color: '#0369a1',
                      display: 'flex',
                      alignItems: 'center',
                      justifyContent: 'space-between'
                    }}>
                      <span>Current File: {form.draftFileName || 'View File'}</span>
                      <a
                        href={`/server/statutoryreg_function/statutory/${editingStatutoryId || form.id}/file/Draft`}
                        target="_blank"
                        rel="noopener noreferrer"
                        style={{ color: '#0ea5e9', textDecoration: 'underline', marginLeft: '8px' }}
                      >
                        View
                      </a>
                    </div>
                  )}
                  <input
                    id="draftFile"
                    name="draftFile"
                    type="file"
                    accept=".pdf,.doc,.docx,.xlsx,.xls,.jpg,.jpeg,.png"
                    onChange={onChange}
                    style={{
                      color: 'transparent',
                      '::file-selector-button': {
                        display: 'none'
                      }
                    }}
                  />
                  {form.draftFile instanceof File && form.draftFileName && (
                    <div style={{
                      marginTop: '8px',
                      padding: '8px 12px',
                      backgroundColor: '#f0f9ff',
                      border: '1px solid #0ea5e9',
                      borderRadius: '6px',
                      fontSize: '14px',
                      color: '#0369a1'
                    }}>
                      New File Selected: {form.draftFileName}
                    </div>
                  )}
                  <small style={{ color: '#6b7280', fontSize: '12px', marginTop: '4px', display: 'block' }}>
                    Maximum file size: 5MB. Supported formats: PDF, DOC, DOCX, XLSX, XLS, JPG, PNG
                  </small>
                </div>
              </div>
              {error && <div className="error">{error}</div>}
              {success && <div className="success">{success}</div>}
              <div className="statutory-actions">
                <button className="btn btn-primary" type="submit" disabled={submitting}>
                  {submitting
                    ? (editingStatutoryId ? 'Updating...' : 'Submitting...')
                    : (editingStatutoryId ? 'Update' : 'Submit')
                  }
                </button>
                <button className="btn btn-secondary" type="button" onClick={onReset} disabled={submitting}>Reset</button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* Form File View Modal */}
      {isFormFileModalOpen && (
        <div
          className="form-file-modal-overlay"
          onClick={handleCloseFormFileModal}
          style={{
            position: 'fixed',
            top: 0,
            left: 0,
            right: 0,
            bottom: 0,
            backgroundColor: 'rgba(0, 0, 0, 0.7)',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            zIndex: 10000,
            padding: '20px'
          }}
        >
          <div
            className="form-file-modal-content"
            onClick={(e) => e.stopPropagation()}
            style={{
              backgroundColor: 'white',
              borderRadius: '12px',
              width: '90%',
              maxWidth: registerOfWagesMergedTheadRows || formVDailyHoursThead || formDRemunerationThead ? 'min(98vw, 1900px)' : '1200px',
              maxHeight: '90vh',
              display: 'flex',
              flexDirection: 'column',
              boxShadow: '0 10px 40px rgba(0, 0, 0, 0.3)'
            }}
          >
            {/* Modal Header */}
            <div style={{
              padding: '20px 24px',
              borderBottom: '1px solid #e5e7eb',
              display: 'flex',
              justifyContent: 'space-between',
              alignItems: 'center'
            }}>
              <h3 style={{ margin: 0, fontSize: '20px', fontWeight: '600', color: '#1f2937' }}>
                {formFileModalData?.fileName || 'Form File'}
              </h3>
              <button
                onClick={handleCloseFormFileModal}
                style={{
                  background: 'none',
                  border: 'none',
                  fontSize: '28px',
                  cursor: 'pointer',
                  color: '#6b7280',
                  padding: '0',
                  width: '32px',
                  height: '32px',
                  display: 'flex',
                  alignItems: 'center',
                  justifyContent: 'center',
                  borderRadius: '4px',
                  transition: 'all 0.2s'
                }}
                onMouseEnter={(e) => {
                  e.target.style.backgroundColor = '#f3f4f6';
                  e.target.style.color = '#1f2937';
                }}
                onMouseLeave={(e) => {
                  e.target.style.backgroundColor = 'transparent';
                  e.target.style.color = '#6b7280';
                }}
                title="Close"
              >
                &times;
              </button>
            </div>

            {/* Modal Body */}
            <div style={{
              padding: '24px',
              overflow: 'auto',
              flex: 1,
              backgroundColor: '#f9fafb'
            }}>
              {formFileLoading ? (
                <div style={{
                  display: 'flex',
                  justifyContent: 'center',
                  alignItems: 'center',
                  minHeight: '200px'
                }}>
                  <div style={{ textAlign: 'center' }}>
                    <div style={{
                      border: '4px solid #f3f4f6',
                      borderTop: '4px solid #3b82f6',
                      borderRadius: '50%',
                      width: '40px',
                      height: '40px',
                      animation: 'spin 1s linear infinite',
                      margin: '0 auto 16px'
                    }}></div>
                    <p style={{ color: '#6b7280', margin: 0 }}>Loading form file...</p>
                  </div>
                </div>
              ) : formFileModalData?.fileType === 'excel-form' ? (
                <div style={{
                  backgroundColor: 'white',
                  padding: '32px',
                  borderRadius: '8px',
                  maxHeight: 'calc(90vh - 200px)',
                  overflow: 'auto'
                }}>
                  {/* Error/Success Messages */}
                  {error && (
                    <div style={{
                      padding: '12px 16px',
                      backgroundColor: '#fef2f2',
                      border: '1px solid #fecaca',
                      borderRadius: '6px',
                      color: '#991b1b',
                      marginBottom: '16px',
                      fontSize: '14px'
                    }}>
                      {error}
                    </div>
                  )}
                  {success && (
                    <div style={{
                      padding: '12px 16px',
                      backgroundColor: '#f0fdf4',
                      border: '1px solid #bbf7d0',
                      borderRadius: '6px',
                      color: '#166534',
                      marginBottom: '16px',
                      fontSize: '14px'
                    }}>
                      {success}
                    </div>
                  )}
                 
                  {/* Form Header Section – format from View File */}
                  {displayFormHeader && (
                    <div style={{
                      marginBottom: '24px',
                      paddingBottom: '24px',
                      borderBottom: '2px solid #e5e7eb'
                    }}>
                      {/* Form Title */}
                      {displayFormHeader.title && (
                        <div style={{
                          fontSize: '24px',
                          fontWeight: '700',
                          color: '#1f2937',
                          textAlign: 'center',
                          marginBottom: '8px'
                        }}>
                          {displayFormHeader.title}
                        </div>
                      )}
                      {/* Form C / Form B: legal citation as 2nd line (between title and subtitle) */}
                      {(shouldAppendReferenceToFormCTitle(displayFormHeader) || shouldPlaceFormBReferenceBeforeSubtitle(displayFormHeader) || shouldPlaceFormAReferenceBeforeSubtitle(displayFormHeader)) && displayFormHeader.reference && (
                        <div style={{
                          fontSize: '14px',
                          color: '#6b7280',
                          fontStyle: 'italic',
                          textAlign: 'center',
                          marginBottom: '8px',
                          lineHeight: 1.45
                        }}>
                          {displayFormHeader.reference}
                        </div>
                      )}
                      {/* Form Subtitle — hide if it only repeats the title (merged Excel cells) */}
                      {displayFormHeader.subtitle &&
                        String(displayFormHeader.subtitle).replace(/\s+/g, ' ').trim().toLowerCase() !==
                          String(displayFormHeader.title || '').replace(/\s+/g, ' ').trim().toLowerCase() && (
                        <div style={{
                          fontSize: '18px',
                          fontWeight: '600',
                          color: '#374151',
                          textAlign: 'center',
                          marginBottom: '8px',
                          whiteSpace: 'pre-line'
                        }}>
                          {displayFormHeader.subtitle}
                          {!!displayFormHeader.reference && /equal\s+remuneration/i.test(displayFormHeader.subtitle) && (
                            <div style={{ fontSize: '12px', color: '#6b7280', fontStyle: 'italic', marginTop: '4px' }}>
                              {displayFormHeader.reference}
                            </div>
                          )}
                        </div>
                      )}
                      {/* Form Reference */}
                      {displayFormHeader.reference &&
                        !(displayFormHeader.subtitle && /equal\s+remuneration/i.test(displayFormHeader.subtitle)) &&
                        !shouldAppendReferenceToFormCTitle(displayFormHeader) &&
                        !shouldPlaceFormBReferenceBeforeSubtitle(displayFormHeader) &&
                        !shouldPlaceFormAReferenceBeforeSubtitle(displayFormHeader) && (
                        <div style={{
                          fontSize: '12px',
                          color: '#6b7280',
                          textAlign: 'center',
                          marginBottom: '16px',
                          fontStyle: 'italic'
                        }}>
                          {displayFormHeader.reference}
                        </div>
                      )}
                     
                      {/* Wage Period line (Form W / wage registers) — full-width banner like Excel row above the grid */}
                      {displayFormHeader.wagePeriodText && !isFormAMusterRollContext(displayFormHeader) && (
                        <div
                          className="statutory-form-w-wage-period"
                          style={{
                            marginTop: '10px',
                            marginBottom: '16px',
                            padding: '12px 16px',
                            width: '100%',
                            boxSizing: 'border-box',
                            backgroundColor: '#f1f5f9',
                            border: '1px solid #cbd5e1',
                            borderRadius: '6px',
                            textAlign: 'center',
                            color: '#0f172a',
                            fontSize: '15px',
                            fontWeight: '600',
                            lineHeight: 1.45,
                            letterSpacing: '0.01em'
                          }}
                        >
                          {displayFormHeader.wagePeriodText}
                        </div>
                      )}

                      {Array.isArray(displayFormHeader.textRows) && displayFormHeader.textRows.length > 0 && (
                        <div style={{
                          marginTop: '10px',
                          marginBottom: '16px',
                          padding: '12px 16px',
                          width: '100%',
                          boxSizing: 'border-box',
                          backgroundColor: '#f8fafc',
                          border: '1px solid #e2e8f0',
                          borderRadius: '6px',
                          color: '#1f2937',
                          fontSize: '14px',
                          lineHeight: 1.5
                        }}>
                          {displayFormHeader.textRows.map((line, idx) => (
                            <div key={`fh-text-row-${idx}`}>{line}</div>
                          ))}
                        </div>
                      )}

                      {isFormAMusterRollContext(displayFormHeader) && (
                        <div style={{
                          marginTop: '10px',
                          marginBottom: '16px',
                          padding: '12px 16px',
                          width: '100%',
                          boxSizing: 'border-box',
                          backgroundColor: '#f8fafc',
                          border: '1px solid #e2e8f0',
                          borderRadius: '6px',
                          color: '#1f2937',
                          fontSize: '14px',
                          lineHeight: 1.5
                        }}>
                          {FORM_A_MARKING_LEGEND.map((line, idx) => (
                            <div key={`form-a-legend-${idx}`}>{line}</div>
                          ))}
                        </div>
                      )}

                      {/* Header Fields */}
                      {displayFormHeader.fields && displayFormHeader.fields.length > 0 && (
                        <div style={{
                          display: 'grid',
                          gridTemplateColumns: (displayFormHeader.form12MergedHeaderCell || displayFormHeader.form10MergedHeaderCell)
                            ? '1fr'
                            : 'repeat(2, 1fr)',
                          gap: '16px',
                          marginTop: '20px'
                        }}>
                          {displayFormHeader.fields.filter((field) => !shouldHideFormAHeaderField(displayFormHeader, field)).map((field, index) => {
                            const isForm12Field =
                              !!displayFormHeader.form12MergedHeaderCell &&
                              String(field.key || '').startsWith('form12_header_');
                            const isForm10Field =
                              !!displayFormHeader.form10MergedHeaderCell &&
                              String(field.key || '').startsWith('form10_header_');
                            const compactHeaderField = isForm12Field || isForm10Field;
                            const commonStyle = {
                              padding: compactHeaderField ? '6px 10px' : '10px 12px',
                              border: '1px solid #d1d5db',
                              borderRadius: '6px',
                              fontSize: compactHeaderField ? '13px' : '14px',
                              color: '#1f2937',
                              backgroundColor: '#fff',
                              transition: 'border-color 0.2s',
                              width: '100%',
                              boxSizing: 'border-box',
                              minHeight: compactHeaderField ? '32px' : undefined
                            };
                            return (
                            <div key={field.key || index} style={{
                              display: 'flex',
                              flexDirection: 'column'
                            }}>
                              <label style={{
                                fontSize: compactHeaderField ? '13px' : '14px',
                                fontWeight: '500',
                                color: '#374151',
                                marginBottom: compactHeaderField ? '4px' : '6px'
                              }}>
                                {field.label}
                              </label>
                              <input
                                type="text"
                                value={headerFormData[field.key] || ''}
                                onChange={(e) => handleHeaderFieldChange(field.key, e.target.value)}
                                placeholder={compactHeaderField
                                  ? `Enter ${field.label.replace(/:+$/, '').trim()}`
                                  : `Enter ${field.label.replace(':', '')}`}
                                style={commonStyle}
                                onFocus={(e) => { e.target.style.borderColor = '#3b82f6'; }}
                                onBlur={(e) => { e.target.style.borderColor = '#d1d5db'; }}
                              />
                            </div>
                            );
                          })}
                        </div>
                      )}

                      {displayFormHeader.festivalGrid &&
                        (displayFormHeader.festivalGrid.approvalProceedings ||
                          (displayFormHeader.festivalGrid.keys && displayFormHeader.festivalGrid.keys.length > 0)) && (
                        <div style={{
                          marginTop: '20px',
                          padding: '16px',
                          border: '1px solid #e5e7eb',
                          borderRadius: '8px',
                          backgroundColor: '#fafafa',
                          maxWidth: '100%',
                          boxSizing: 'border-box'
                        }}>
                          {displayFormHeader.festivalGrid.approvalProceedings && (
                            <div style={{ marginBottom: '18px' }}>
                              <label style={{
                                fontSize: '14px',
                                fontWeight: '500',
                                color: '#374151',
                                marginBottom: '8px',
                                display: 'block',
                                lineHeight: 1.4
                              }}>
                                {(displayFormHeader.festivalGrid.approvalProceedings.label || 'Festival Holidays Approval Proceedings No. and Date:')
                                  .replace(/\s*:\s*$/, '')}:
                              </label>
                              <input
                                type="text"
                                value={headerFormData[displayFormHeader.festivalGrid.approvalProceedings.key] || ''}
                                onChange={(e) => handleHeaderFieldChange(displayFormHeader.festivalGrid.approvalProceedings.key, e.target.value)}
                                placeholder="Festival holidays approval proceedings no. and date"
                                style={{
                                  width: '100%',
                                  maxWidth: '100%',
                                  boxSizing: 'border-box',
                                  padding: '10px 12px',
                                  border: '1px solid #d1d5db',
                                  borderRadius: '6px',
                                  fontSize: '14px',
                                  color: '#1f2937',
                                  backgroundColor: '#fff',
                                  transition: 'border-color 0.2s'
                                }}
                                onFocus={(e) => { e.target.style.borderColor = '#3b82f6'; }}
                                onBlur={(e) => { e.target.style.borderColor = '#d1d5db'; }}
                              />
                            </div>
                          )}

                          {displayFormHeader.festivalGrid.keys && displayFormHeader.festivalGrid.keys.length > 0 && (
                            <div>
                              <div style={{
                                fontSize: '14px',
                                fontWeight: '600',
                                color: '#374151',
                                marginBottom: '10px'
                              }}>
                                Approved Festival Holidays
                              </div>
                              <div style={{
                                display: 'grid',
                                gridTemplateColumns: `repeat(${Math.min(displayFormHeader.festivalGrid.keys.length, 24)}, minmax(0, 1fr))`,
                                gap: '12px',
                                maxWidth: '100%'
                              }}>
                                {displayFormHeader.festivalGrid.keys.map((slotKey, idx) => (
                                  <div key={slotKey} style={{ display: 'flex', flexDirection: 'column' }}>
                                    <label style={{
                                      fontSize: '12px',
                                      fontWeight: '500',
                                      color: '#6b7280',
                                      marginBottom: '4px'
                                    }}>
                                      {idx + 1}
                                    </label>
                                    <input
                                      type="text"
                                      value={headerFormData[slotKey] || ''}
                                      onChange={(e) => handleHeaderFieldChange(slotKey, e.target.value)}
                                      style={{
                                        padding: '10px 12px',
                                        border: '1px solid #d1d5db',
                                        borderRadius: '6px',
                                        fontSize: '14px',
                                        color: '#1f2937',
                                        backgroundColor: '#fff',
                                        transition: 'border-color 0.2s',
                                        minWidth: 0
                                      }}
                                      onFocus={(e) => { e.target.style.borderColor = '#3b82f6'; }}
                                      onBlur={(e) => { e.target.style.borderColor = '#d1d5db'; }}
                                    />
                                  </div>
                                ))}
                              </div>
                            </div>
                          )}
                        </div>
                      )}
                    </div>
                  )}

                  {/* Action Buttons */}
                  <div style={{
                    marginBottom: '16px',
                    display: 'flex',
                    justifyContent: 'flex-end',
                    gap: '12px',
                    flexWrap: 'wrap'
                  }}>
                    {/* Autofill button - Hide if draft file exists */}
                    {displayTableHeaders.length > 0 && !formFileModalData?.item?.draftFile && (
                      <button
                        onClick={fetchAndPopulateEmployeeData}
                        disabled={formFileLoading}
                        style={{
                          padding: '10px 20px',
                          backgroundColor: formFileLoading ? '#9ca3af' : STAT_FORM_FILE_ACCENT_BG,
                          color: 'white',
                          border: 'none',
                          borderRadius: '6px',
                          fontSize: '14px',
                          fontWeight: '500',
                          cursor: formFileLoading ? 'not-allowed' : 'pointer',
                          display: 'flex',
                          alignItems: 'center',
                          gap: '8px',
                          transition: 'all 0.2s ease'
                        }}
                        onMouseOver={(e) => {
                          if (!formFileLoading) {
                            e.currentTarget.style.backgroundColor = STAT_FORM_FILE_ACCENT_BG_HOVER;
                          }
                        }}
                        onMouseOut={(e) => {
                          if (!formFileLoading) {
                            e.currentTarget.style.backgroundColor = STAT_FORM_FILE_ACCENT_BG;
                          }
                        }}
                        title="Fetch employee data from Zoho People and populate the form"
                      >
                        {formFileLoading ? 'Loading...' : 'Autofill'}
                      </button>
                    )}
                    {/* Import button (after Autofill) */}
                    {displayTableHeaders.length > 0 && (
                      <>
                        <input
                          ref={formModalImportInputRef}
                          type="file"
                          accept=".xlsx,.xls"
                          style={{ display: 'none' }}
                          onChange={async (e) => {
                            const file = e.target.files?.[0];
                            await handleImportFormRows(file, displayTableHeaders);
                            e.target.value = '';
                          }}
                        />
                        <button
                          onClick={() => formModalImportInputRef.current?.click()}
                          style={{
                            padding: '10px 20px',
                            backgroundColor: '#14b8a6',
                            color: 'white',
                            border: 'none',
                            borderRadius: '6px',
                            fontSize: '14px',
                            fontWeight: '500',
                            cursor: 'pointer',
                            display: 'flex',
                            alignItems: 'center',
                            gap: '8px',
                            transition: 'all 0.2s ease'
                          }}
                          onMouseOver={(e) => {
                            e.currentTarget.style.backgroundColor = '#0d9488';
                          }}
                          onMouseOut={(e) => {
                            e.currentTarget.style.backgroundColor = '#14b8a6';
                          }}
                          title="Import rows from Excel file"
                        >
                          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                            <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
                            <polyline points="7 10 12 15 17 10"></polyline>
                            <line x1="12" y1="15" x2="12" y2="3"></line>
                          </svg>
                          Import
                        </button>
                      </>
                    )}
                    {/* Save Button */}
                    {formTableData.length > 0 && displayTableHeaders.length > 0 && (
                      <button
                        onClick={handleSaveFormFile}
                        disabled={submitting}
                        style={{
                          padding: '10px 20px',
                          backgroundColor: submitting ? '#9ca3af' : '#10b981',
                          color: 'white',
                          border: 'none',
                          borderRadius: '6px',
                          fontSize: '14px',
                          fontWeight: '500',
                          cursor: submitting ? 'not-allowed' : 'pointer',
                          display: 'flex',
                          alignItems: 'center',
                          gap: '8px',
                          transition: 'all 0.2s ease'
                        }}
                        onMouseOver={(e) => {
                          if (!submitting) {
                            e.currentTarget.style.backgroundColor = '#059669';
                          }
                        }}
                        onMouseOut={(e) => {
                          if (!submitting) {
                            e.currentTarget.style.backgroundColor = '#10b981';
                          }
                        }}
                        title="Save form data"
                      >
                        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                          <path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"></path>
                          <polyline points="17 21 17 13 7 13 7 21"></polyline>
                          <polyline points="7 3 7 8 15 8"></polyline>
                        </svg>
                        {submitting ? 'Saving...' : 'Save'}
                      </button>
                    )}
                    {/* Add Row Button */}
                    {formTableData.length > 0 && (
                      <button
                        onClick={() => {
                          // Add a new empty row
                          const newRow = {};
                          displayTableHeaders.forEach(header => {
                            newRow[header] = '';
                          });
                          setFormTableData([...formTableData, newRow]);
                        }}
                        style={{
                          padding: '10px 20px',
                          backgroundColor: '#10b981',
                          color: 'white',
                          border: 'none',
                          borderRadius: '6px',
                          fontSize: '14px',
                          fontWeight: '500',
                          cursor: 'pointer',
                          display: 'flex',
                          alignItems: 'center',
                          gap: '8px',
                          transition: 'all 0.2s ease'
                        }}
                        onMouseOver={(e) => {
                          e.currentTarget.style.backgroundColor = '#059669';
                        }}
                        onMouseOut={(e) => {
                          e.currentTarget.style.backgroundColor = '#10b981';
                        }}
                      >
                        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                          <line x1="12" y1="5" x2="12" y2="19"/>
                          <line x1="5" y1="12" x2="19" y2="12"/>
                        </svg>
                        Add Row
                      </button>
                    )}
                  </div>

                  {/* Form table + employer / signatory footer (Form I layout and all statutory excel forms) */}
                  {formTableData.length > 0 && displayTableHeaders.length > 0 && (
                    <>
                    <div className="form-file-modal-table-wrap" style={{
                      overflowX: 'auto',
                      border: '1px solid #e5e7eb',
                      borderRadius: '8px'
                    }}>
                      <table className="form-file-modal-table" style={{
                        width: '100%',
                        borderCollapse: 'collapse',
                        backgroundColor: 'white'
                      }}>
                        <thead>
                          {registerOfWagesMergedTheadRows ? (
                            registerOfWagesMergedTheadRows
                          ) : formVDailyHoursThead ? (
                            formVDailyHoursThead
                          ) : formVIFestivalThead ? (
                            formVIFestivalThead
                          ) : formDRemunerationThead ? (
                            formDRemunerationThead
                          ) : formBLabourWelfareThead ? (
                            formBLabourWelfareThead
                          ) : displayColumnGroupLabels ? (() => {
                            const labels = displayColumnGroupLabels;
                            const hdrs = displayTableHeaders;
                            const n = hdrs.length;
                            const norm = (s) => String(s ?? '').trim().toLowerCase().replace(/\s+/g, ' ');
                            const looksEmployeeBaseColumn = (headerText) => {
                              const x = norm(headerText);
                              if (!x) return false;
                              if (/^s\.?\s*no\.?$/.test(String(headerText).trim())) return true;
                              if (x.includes('name of the employee') || x === 'name of employee' || (x.includes('name') && x.includes('employee'))) return true;
                              if (x.includes('identification') && (x.includes('employee') || x.includes('no'))) return true;
                              return false;
                            };
                            const isMergedIdentityCol = (i) => {
                              const lab = String(labels[i] ?? '').trim();
                              const h = String(hdrs[i] ?? '').trim();
                              if (!h) return false;
                              if (lab && norm(lab) === norm(h)) return true;
                              if (!lab && looksEmployeeBaseColumn(h)) return true;
                              return false;
                            };
                            let split = 0;
                            while (split < n && isMergedIdentityCol(split)) split += 1;

                            const groupHeaderStyle = {
                              backgroundColor: STAT_FORM_FILE_ACCENT_BG,
                              color: 'white',
                              padding: '10px 12px',
                              textAlign: 'center',
                              fontWeight: '700',
                              fontSize: '13px',
                              border: `1px solid ${STAT_FORM_FILE_ACCENT_BORDER}`,
                              verticalAlign: 'middle'
                            };
                            const identityOrDetailHeaderStyle = {
                              backgroundColor: STAT_FORM_FILE_ACCENT_BG,
                              color: 'white',
                              padding: '12px',
                              textAlign: 'left',
                              fontWeight: '600',
                              fontSize: '14px',
                              border: `1px solid ${STAT_FORM_FILE_ACCENT_BORDER}`,
                              verticalAlign: 'middle'
                            };
                            const stickyDetailHeaderStyle = {
                              ...identityOrDetailHeaderStyle,
                              position: 'sticky',
                              top: 0,
                              zIndex: 10
                            };

                            const row1Cells = [];
                            for (let idx = 0; idx < split; idx += 1) {
                              row1Cells.push(
                                <th key={`form-x-id-${idx}`} rowSpan={2} style={identityOrDetailHeaderStyle}>
                                  {hdrs[idx]}
                                </th>
                              );
                            }
                            let i = split;
                            while (i < n) {
                              const lab = String(labels[i] ?? '').trim();
                              let j = i + 1;
                              while (j < n && String(labels[j] ?? '').trim() === lab) j += 1;
                              row1Cells.push(
                                <th key={`form-x-grp-${i}`} colSpan={j - i} style={groupHeaderStyle}>
                                  {lab || '\u00a0'}
                                </th>
                              );
                              i = j;
                            }

                            const row2Cells = [];
                            for (let idx = split; idx < n; idx += 1) {
                              row2Cells.push(
                                <th key={`form-x-sub-${idx}`} style={stickyDetailHeaderStyle}>
                                  {formatStatutoryTableHeaderLabel(hdrs[idx])}
                                </th>
                              );
                            }

                            return (
                              <>
                                <tr>{row1Cells}</tr>
                                {row2Cells.length > 0 ? <tr>{row2Cells}</tr> : null}
                              </>
                            );
                          })() : (
                            <tr>
                              {displayTableHeaders.map((header, index) => (
                                <th
                                  key={index}
                                  title={header}
                                  style={{
                                    backgroundColor: STAT_FORM_FILE_ACCENT_BG,
                                    color: 'white',
                                    padding: '12px',
                                    textAlign: 'left',
                                    fontWeight: '600',
                                    fontSize: '14px',
                                    border: `1px solid ${STAT_FORM_FILE_ACCENT_BORDER}`,
                                    position: 'sticky',
                                    top: 0,
                                    zIndex: 10
                                  }}
                                >
                                  {formatStatutoryTableHeaderLabel(header)}
                                </th>
                              ))}
                            </tr>
                          )}
                        </thead>
                        <tbody>
                          {formTableData.map((row, rowIndex) => (
                            <tr key={rowIndex} style={{
                              backgroundColor: rowIndex % 2 === 0 ? '#fff' : '#f9fafb'
                            }}>
                              {displayTableHeaders.map((header, colIndex) => (
                                <td key={colIndex} style={{
                                  padding: '10px 12px',
                                  border: '1px solid #e5e7eb',
                                  fontSize: '14px',
                                  color: '#374151'
                                }}>
                                  <input
                                    type="text"
                                    value={row[header] === false || row[header] === 'false' || !row[header] ? '' : row[header]}
                                    onChange={(e) => {
                                      const newData = [...formTableData];
                                      newData[rowIndex] = {
                                        ...newData[rowIndex],
                                        [header]: e.target.value
                                      };
                                      setFormTableData(newData);
                                    }}
                                    style={{
                                      width: '100%',
                                      padding: '6px 8px',
                                      border: '1px solid #d1d5db',
                                      borderRadius: '4px',
                                      fontSize: '14px',
                                      backgroundColor: 'white'
                                    }}
                                    placeholder={`Enter ${formatStatutoryTableHeaderLabel(header)}`}
                                  />
                                </td>
                              ))}
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                    <div className="statutory-form-signature-footer" aria-label="Employer signature block">
                      <div className="statutory-form-signature-footer__inner">
                        <div className="statutory-form-signature-footer__for">
                          For {statutoryModalEmployerName ? statutoryModalEmployerName : '(Company Name)'}
                        </div>
                        <div className="statutory-form-signature-footer__mid">Authorised Signatory</div>
                        <div className="statutory-form-signature-footer__sig">
                          Signature of Employer / Manager / Authorised Person
                        </div>
                      </div>
                    </div>
                    </>
                  )}
                </div>
              ) : formFileModalData?.fileType === 'pdf' ? (
                <iframe
                  src={formFileModalData.content}
                  style={{
                    width: '100%',
                    height: 'calc(90vh - 200px)',
                    border: 'none',
                    borderRadius: '8px'
                  }}
                  title={formFileModalData.fileName}
                />
              ) : formFileModalData?.fileType === 'image' ? (
                <div style={{ textAlign: 'center' }}>
                  <img
                    src={formFileModalData.content}
                    alt={formFileModalData.fileName}
                    style={{
                      maxWidth: '100%',
                      maxHeight: 'calc(90vh - 200px)',
                      borderRadius: '8px',
                      boxShadow: '0 4px 12px rgba(0, 0, 0, 0.1)'
                    }}
                  />
                </div>
              ) : (
                <iframe
                  src={formFileModalData?.content}
                  style={{
                    width: '100%',
                    height: 'calc(90vh - 200px)',
                    border: 'none',
                    borderRadius: '8px'
                  }}
                  title={formFileModalData?.fileName}
                />
              )}
            </div>
          </div>
        </div>
      )}

      {/* CSS Animation for Spinner */}
      <style>{`
        @keyframes spin {
          0% { transform: rotate(0deg); }
          100% { transform: rotate(360deg); }
        }
      `}</style>
    </div>
  );
};

export default Statutory;
