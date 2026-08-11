import React, { useEffect, useMemo, useRef, useState } from 'react';
import { Link } from 'react-router-dom';
import { ChevronDown } from 'lucide-react';
import { MapContainer, GeoJSON, Marker, Popup, useMap } from 'react-leaflet';
import L from 'leaflet';
import 'leaflet/dist/leaflet.css';
import {
  buildInchargeSiteScopeFromList,
  fetchInchargeDisplayScopeFromSites,
  filterSitesForLoginUser,
  getActCategoryFromActSector,
  industryLabelToActCategory,
  isOrgWideSiteViewer,
  normalizeStateCompareKey,
  sectorMatchesInchargeSiteIndustries,
  siteNameFromSiteRecord,
  siteStateFromRecord,
  siteIndustry,
  splitStateFieldTokens,
  statesFieldMatchesInchargeSiteStates,
} from '../utils/siteInchargeScope';
import { resolveLoginEmailString } from '../utils/resolveLoginEmail';
import INDIAN_STATES from '../utils/indianStates';
import './ComplianceHealthDashboard.css';

const STATUTORY_API = '/server/statutoryreg_function/statutory';
const CHECKLIST_BULK_API = '/server/checklistbulk_function/checklistbulk?action=getAll';
const SITE_MANAGEMENT_API = '/server/sitemanagement_function/sitemanagement';
const POLL_MS = 30000;
const TRANSACTION_PAGE = '/rule-book/statutory';

const STATUS_COLORS = {
  approved: '#22c55e',
  pending: '#eab308',
  yetToSubmit: '#3b82f6',
  returned: '#ef4444',
};

const STATUS_LABELS = {
  approved: 'Approved',
  pending: 'Pending',
  yetToSubmit: 'Yet to Submit',
  returned: 'Returned',
};

/** Location legend order (matches mockup). */
const LOCATION_STATUS_ORDER = ['yetToSubmit', 'pending', 'approved', 'returned'];
/** Industry legend / bar group order (matches mockup). */
const INDUSTRY_STATUS_ORDER = ['approved', 'pending', 'yetToSubmit', 'returned'];
/** Trend legend order (image-1 model): Approved → Pending → Yet to Submit → Returned. */
const TREND_STATUS_ORDER = ['approved', 'pending', 'yetToSubmit', 'returned'];

/** Site-login “Compliance by Act” legend (image model). */
const ACT_CHART_STATUS_ORDER = ['approved', 'pending', 'returned', 'yetToSubmit'];
const ACT_CHART_LABELS = {
  approved: 'Approved',
  pending: 'Pending',
  returned: 'Returned',
  yetToSubmit: 'Yet to Submit',
};

const MONTH_NAME_INDEX = {
  jan: 0,
  january: 0,
  feb: 1,
  february: 1,
  mar: 2,
  march: 2,
  apr: 3,
  april: 3,
  may: 4,
  jun: 5,
  june: 5,
  jul: 6,
  july: 6,
  aug: 7,
  august: 7,
  sep: 8,
  sept: 8,
  september: 8,
  oct: 9,
  october: 9,
  nov: 10,
  november: 10,
  dec: 11,
  december: 11,
};

const PERIOD_OPTIONS = [
  { value: 'thisMonth', label: 'This Month' },
  { value: 'lastMonth', label: 'Last Month' },
  { value: 'last6Months', label: 'Last 6 Months' },
  { value: 'thisYear', label: 'This Year' },
];

const EMPTY_COUNTS = { approved: 0, pending: 0, yetToSubmit: 0, returned: 0 };

const INDUSTRY_GROUPS = [
  { key: 'clra', label: 'CLRA', category: 'clra' },
  { key: 'factory', label: 'Factory', category: 'factories' },
  { key: 'se', label: 'S&E', category: 'shops_and_establishment' },
];

/** Approximate state centroids for India map markers. */
const STATE_CENTROIDS = {
  andhrapradesh: [15.9129, 79.74],
  arunachalpradesh: [28.218, 94.7278],
  assam: [26.2006, 92.9376],
  bihar: [25.0961, 85.3131],
  chhattisgarh: [21.2787, 81.8661],
  goa: [15.2993, 74.124],
  gujarat: [22.2587, 71.1924],
  haryana: [29.0588, 76.0856],
  himachalpradesh: [31.1048, 77.1734],
  jharkhand: [23.6102, 85.2799],
  karnataka: [15.3173, 75.7139],
  kerala: [10.8505, 76.2711],
  madhyapradesh: [22.9734, 78.6569],
  maharashtra: [19.7515, 75.7139],
  manipur: [24.6637, 93.9063],
  meghalaya: [25.467, 91.3662],
  mizoram: [23.1645, 92.9376],
  nagaland: [26.1584, 94.5624],
  odisha: [20.9517, 85.0985],
  punjab: [31.1471, 75.3412],
  rajasthan: [27.0238, 74.2179],
  sikkim: [27.533, 88.5122],
  tamilnadu: [11.1271, 78.6569],
  telangana: [18.1124, 79.0193],
  tripura: [23.9408, 91.9882],
  uttarpradesh: [26.8467, 80.9462],
  uttarakhand: [30.0668, 79.0193],
  westbengal: [22.9868, 87.855],
  delhi: [28.6139, 77.209],
  nctofdelhi: [28.6139, 77.209],
};

/** Fixed marker colors for known states on the India map. */
const STATE_MAP_COLORS = {
  tamilnadu: '#22c55e', // green
  karnataka: '#eab308', // yellow
  gujarat: '#8b5cf6', // purple
  rajasthan: '#ef4444', // red
  andhrapradesh: '#38bdf8', // skyblue
  madhyapradesh: '#ec4899', // pink
  maharashtra: '#f97316', // orange
};

/** Fallback palette for states without a fixed color. */
const STATE_MAP_PALETTE = [
  '#3b82f6', // blue
  '#06b6d4', // cyan
  '#84cc16', // lime
  '#14b8a6', // teal
  '#6366f1', // indigo
  '#e11d48', // rose
  '#f59e0b', // amber
];

function mapColorForState(stateName) {
  const key = normalizeStateCompareKey(stateName) || String(stateName || '');
  if (STATE_MAP_COLORS[key]) return STATE_MAP_COLORS[key];
  let hash = 0;
  for (let i = 0; i < key.length; i += 1) {
    hash = (hash * 31 + key.charCodeAt(i)) >>> 0;
  }
  return STATE_MAP_PALETTE[hash % STATE_MAP_PALETTE.length];
}

/** Same rules as Statutory Status column / Calendar KPIs. */
function getComplianceBucket(item) {
  const raw = item?.status != null && item.status !== '' ? item.status : item?.Status;
  const norm = String(raw || '').trim();
  const normalized = norm.toLowerCase();
  const rawApproval =
    item?.approval != null && item.approval !== '' ? item.approval : item?.Approval;
  const approvalNorm = String(rawApproval || '').trim().toLowerCase();

  if (
    approvalNorm === 'approved' ||
    approvalNorm === 'approve' ||
    normalized === 'approved' ||
    normalized === 'approve'
  ) {
    return 'approved';
  }
  if (
    approvalNorm === 'rejected' ||
    approvalNorm === 'reject' ||
    normalized === 'rejected' ||
    normalized === 'reject' ||
    normalized === 'returned'
  ) {
    return 'returned';
  }

  const sendForApproval = String(item?.sendForApproval ?? item?.SendForApproval ?? '')
    .trim()
    .toLowerCase();
  // Pending = sent for approval, or explicit Pending status (matches top KPI / Calendar).
  if (sendForApproval === 'sent' || normalized === 'pending') {
    return 'pending';
  }

  const draftVal = item?.draftFile ?? item?.DraftFile ?? null;
  const hasDraftFile =
    draftVal != null &&
    String(draftVal).trim() !== '' &&
    String(draftVal).trim() !== 'null' &&
    String(draftVal).trim() !== 'undefined';

  if (
    normalized === 'yet to complete' ||
    normalized === 'yet to comply' ||
    normalized === 'yet to submit'
  ) {
    return 'yetToSubmit';
  }

  // Draft uploaded but not sent → still awaiting submit/send (Yet to submit).
  if (!hasDraftFile) return 'yetToSubmit';
  if (norm === '' || norm === '-' || norm === '—') return 'yetToSubmit';

  return 'yetToSubmit';
}

/** Day-of-month due dates (1–31) are not timestamps — common on statutory rows. */
const isDayOnlyDueDate = (value) => {
  if (typeof value === 'number' && Number.isInteger(value) && value >= 1 && value <= 31) {
    return true;
  }
  const text = String(value ?? '').trim();
  if (!/^\d{1,2}$/.test(text)) return false;
  const n = Number(text);
  return n >= 1 && n <= 31;
};

const parseMonthIndex = (value) => {
  if (value == null || value === '') return null;
  if (typeof value === 'number' && Number.isInteger(value) && value >= 1 && value <= 12) {
    return value - 1;
  }
  const text = String(value).trim().toLowerCase();
  if (!text) return null;
  if (/^\d{1,2}$/.test(text)) {
    const n = Number(text);
    if (n >= 1 && n <= 12) return n - 1;
    return null;
  }
  // "Mar", "March", "Mar-2026", "15-Mar-2026", "March 2026"
  const token = text.match(/[a-z]+/)?.[0] || '';
  if (token && Object.prototype.hasOwnProperty.call(MONTH_NAME_INDEX, token)) {
    return MONTH_NAME_INDEX[token];
  }
  return null;
};

const parseDateValue = (value) => {
  if (value == null || value === '') return null;
  // Reject day-of-month due dates so they are not treated as epoch timestamps.
  if (isDayOnlyDueDate(value)) return null;
  if (typeof value === 'number' && Number.isFinite(value)) {
    // Real epoch ms / sec only (ignore tiny integers that are not dates).
    if (value > 0 && value < 1e10) {
      // 10-digit unix seconds, or reject values that are too small to be ms.
      if (value >= 1e9) return value * 1000;
      return null;
    }
    return value > 0 ? value : null;
  }
  const text = String(value).trim();
  if (!text) return null;
  if (/^monthly\s*basis$/i.test(text)) return null;
  if (/^\d{10,13}$/.test(text)) {
    const n = Number(text);
    return n > 0 ? (text.length === 10 ? n * 1000 : n) : null;
  }
  // dd/mm/yyyy or dd-mm-yyyy
  const dmy = text.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
  if (dmy) {
    const day = Number(dmy[1]);
    const month = Number(dmy[2]) - 1;
    let year = Number(dmy[3]);
    if (year < 100) year += 2000;
    const d = new Date(year, month, day);
    return Number.isNaN(d.getTime()) ? null : d.getTime();
  }
  // 15-Mar-2026 / Mar-2026 / 15 Mar 2026
  const monYear = text.match(/^(?:\d{1,2}[\s\-]*)?([a-z]{3,9})[\s\-\/]*(\d{2,4})$/i);
  if (monYear) {
    const monthIdx = parseMonthIndex(monYear[1]);
    if (monthIdx != null) {
      let year = Number(monYear[2]);
      if (year < 100) year += 2000;
      const d = new Date(year, monthIdx, 1);
      return Number.isNaN(d.getTime()) ? null : d.getTime();
    }
  }
  const parsed = Date.parse(text);
  return Number.isNaN(parsed) ? null : parsed;
};

const monthKey = (date) => `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;

const monthLabel = (date) =>
  date.toLocaleDateString('en-IN', { month: 'short' });

/** Normalize MonthFilter / due month so Mar vs April rows stay distinct. */
function rowMonthKeyPart(row) {
  const raw = row?.monthFilter ?? row?.MonthFilter ?? row?.monthfilter;
  if (raw != null && String(raw).trim() !== '') {
    const idx = parseMonthIndex(raw);
    if (idx != null) return `m${idx}`;
    return String(raw).trim().toLowerCase().slice(0, 3);
  }
  const dueRaw = row?.dueDate ?? row?.DueDate;
  if (dueRaw != null && String(dueRaw).trim() !== '') {
    if (!isDayOnlyDueDate(dueRaw) && !/^monthly\s*basis$/i.test(String(dueRaw).trim())) {
      const ts = parseDateValue(dueRaw);
      if (ts != null) {
        const d = new Date(ts);
        return monthKey(new Date(d.getFullYear(), d.getMonth(), 1));
      }
      const idx = parseMonthIndex(dueRaw);
      if (idx != null) return `m${idx}`;
    }
  }
  return 'nomonth';
}

function rowIdentityKey(row, { ignoreSite = false, ignoreMonth = false } = {}) {
  const form = String(row?.formName ?? row?.FormName ?? '')
    .trim()
    .toLowerCase();
  const act = String(row?.act ?? row?.Act ?? '')
    .trim()
    .toLowerCase();
  const desc = String(row?.description ?? row?.Description ?? '')
    .trim()
    .toLowerCase();
  const site = String(row?.site ?? row?.Site ?? row?.siteName ?? row?.SiteName ?? '')
    .trim()
    .toLowerCase();
  const month = ignoreMonth ? '*' : rowMonthKeyPart(row);
  if (!form && !act && !desc) return '';
  if (ignoreSite) return `${form}|${act}|${desc}|${month}`;
  return `${form}|${act}|${desc}|${site}|${month}`;
}

function workflowScore(row) {
  let score = 0;
  const bucket = getComplianceBucket(row);
  if (bucket === 'approved') score += 400;
  else if (bucket === 'returned') score += 350;
  else if (bucket === 'pending') score += 300;
  const draftVal = row?.draftFile ?? row?.DraftFile;
  if (draftVal != null && String(draftVal).trim() && String(draftVal).trim() !== 'null') score += 100;
  const formFile = row?.formFile ?? row?.FormFile;
  if (formFile != null && String(formFile).trim() && String(formFile).trim() !== 'null') score += 20;
  if (row?.id != null && String(row.id).trim() !== '') score += 10;
  return score;
}

function pickStrongerRow(a, b) {
  if (!a) return b;
  if (!b) return a;
  return workflowScore(b) >= workflowScore(a) ? b : a;
}

function overlayMasterOntoWorkflow(master, workflow) {
  return {
    ...master,
    ...workflow,
    act: workflow.act ?? workflow.Act ?? master.act ?? master.Act,
    description:
      workflow.description ?? workflow.Description ?? master.description ?? master.Description,
    formName: workflow.formName ?? workflow.FormName ?? master.formName ?? master.FormName,
    sector: workflow.sector ?? workflow.Sector ?? master.sector ?? master.Sector,
    states:
      workflow.states ??
      workflow.state ??
      workflow.State ??
      master.states ??
      master.state ??
      master.State,
    dueDate: workflow.dueDate ?? workflow.DueDate ?? master.dueDate ?? master.DueDate,
    site: workflow.site ?? workflow.Site ?? master.site ?? master.Site,
    monthFilter:
      workflow.monthFilter ??
      workflow.MonthFilter ??
      workflow.monthfilter ??
      master.monthFilter ??
      master.MonthFilter ??
      master.monthfilter,
    MonthFilter:
      workflow.MonthFilter ??
      workflow.monthFilter ??
      workflow.monthfilter ??
      master.MonthFilter ??
      master.monthFilter ??
      master.monthfilter,
    monthfilter:
      workflow.monthfilter ??
      workflow.monthFilter ??
      workflow.MonthFilter ??
      master.monthfilter ??
      master.monthFilter ??
      master.MonthFilter,
  };
}

/** Prefer statutory workflow rows; keep one row per form+site+month. */
function mergeStatutoryWithBulk(statutory, bulk, now = new Date()) {
  const byKey = new Map();
  /** form|act|desc → Map(monthPart → row) so month history is not collapsed. */
  const byFormActDescMonths = new Map();
  const currentMonthPart = `m${now.getMonth()}`;

  const rememberSoft = (row) => {
    const softBase = rowIdentityKey(row, { ignoreSite: true, ignoreMonth: true });
    if (!softBase) return;
    const monthPart = rowMonthKeyPart(row);
    if (!byFormActDescMonths.has(softBase)) byFormActDescMonths.set(softBase, new Map());
    const monthMap = byFormActDescMonths.get(softBase);
    monthMap.set(monthPart, pickStrongerRow(monthMap.get(monthPart), row));
  };

  const pickSoftMatch = (softBase, preferredMonth) => {
    const monthMap = byFormActDescMonths.get(softBase);
    if (!monthMap || monthMap.size === 0) return null;
    if (preferredMonth && monthMap.has(preferredMonth)) return monthMap.get(preferredMonth);
    if (monthMap.has(currentMonthPart)) return monthMap.get(currentMonthPart);
    if (monthMap.has('nomonth')) return monthMap.get('nomonth');
    let best = null;
    monthMap.forEach((row) => {
      best = pickStrongerRow(best, row);
    });
    return best;
  };

  (statutory || []).forEach((row) => {
    const fullKey = rowIdentityKey(row) || `stat-${byKey.size}-${String(row?.id ?? '')}`;
    const next = pickStrongerRow(byKey.get(fullKey), row);
    byKey.set(fullKey, next);
    rememberSoft(next);
  });

  (bulk || []).forEach((row) => {
    const softBase = rowIdentityKey(row, { ignoreSite: true, ignoreMonth: true });
    if (!softBase) return;

    const bulkMonth = rowMonthKeyPart(row);
    const fullKey = rowIdentityKey(row);
    const exact = fullKey ? byKey.get(fullKey) : null;
    const soft = pickSoftMatch(softBase, bulkMonth === 'nomonth' ? currentMonthPart : bulkMonth);

    if (exact) {
      const merged = overlayMasterOntoWorkflow(row, exact);
      byKey.set(fullKey, merged);
      rememberSoft(merged);
      return;
    }

    if (soft) {
      // Bulk often has no MonthFilter — attach onto the matching month's statutory row.
      const softFullKey =
        rowIdentityKey(soft) || `stat-soft-${String(soft?.id ?? softBase)}`;
      const merged = overlayMasterOntoWorkflow(row, soft);
      byKey.set(softFullKey, merged);
      rememberSoft(merged);
      return;
    }

    const insertKey = fullKey || `bulk-${byKey.size}-${softBase}`;
    byKey.set(insertKey, row);
    rememberSoft(row);
  });

  return [...byKey.values()];
}

function itemActivityTimestamp(item) {
  const dueRaw = item?.dueDate ?? item?.DueDate;
  const candidates = [
    // Skip day-only / "Monthly Basis" due dates — those are not activity timestamps.
    isDayOnlyDueDate(dueRaw) || /^monthly\s*basis$/i.test(String(dueRaw || '').trim())
      ? null
      : dueRaw,
    item?.approvedDate ?? item?.ApprovedDate,
    item?.submittedDate ?? item?.SubmittedDate,
    item?.modifiedTime ?? item?.ModifiedTime ?? item?.MODIFIEDTIME,
    item?.createdTime ?? item?.CreatedTime ?? item?.CREATEDTIME,
  ];
  for (let i = 0; i < candidates.length; i += 1) {
    const ts = parseDateValue(candidates[i]);
    if (ts != null) return ts;
  }
  return null;
}

/** Workflow event month only — never created/modified (those dump open rows into one month). */
function itemWorkflowMonthTimestamp(item) {
  const candidates = [
    item?.approvedDate ?? item?.ApprovedDate,
    item?.submittedDate ?? item?.SubmittedDate,
  ];
  for (let i = 0; i < candidates.length; i += 1) {
    const ts = parseDateValue(candidates[i]);
    if (ts != null) return ts;
  }
  return null;
}

/** Map a month name / index onto the last-6-months window key (YYYY-MM). */
function monthIndexToWindowKey(monthIdx, windowKeys) {
  if (monthIdx == null || monthIdx < 0 || monthIdx > 11) return null;
  for (let i = 0; i < windowKeys.length; i += 1) {
    const key = windowKeys[i];
    const monthPart = Number(String(key).split('-')[1]);
    if (monthPart - 1 === monthIdx) return key;
  }
  return null;
}

/** Monthly Basis / day-of-month dues recur every month — not a single bucket. */
function isMonthlyRecurringDue(item) {
  const dueRaw = item?.dueDate ?? item?.DueDate;
  if (dueRaw == null || String(dueRaw).trim() === '') return true;
  if (/^monthly\s*basis$/i.test(String(dueRaw).trim())) return true;
  if (isDayOnlyDueDate(dueRaw)) return true;
  return false;
}

/**
 * Statutory month for a row inside the trend window.
 * Uses MonthFilter (Statutory) → real due date → approved/submitted date.
 * Does NOT use created/modified time (that wrongly piled Yet to Submit into one month).
 */
function resolveTrendMonthKey(item, windowKeys, now = new Date()) {
  const windowSet = new Set(windowKeys);
  const currentKey = monthKey(new Date(now.getFullYear(), now.getMonth(), 1));

  // 1) Statutory MonthFilter is the source of truth for month history.
  const monthFilter = item?.monthFilter ?? item?.MonthFilter ?? item?.monthfilter;
  if (monthFilter != null && String(monthFilter).trim() !== '') {
    const fromFilter = monthIndexToWindowKey(parseMonthIndex(monthFilter), windowKeys);
    if (fromFilter) return fromFilter;
    return undefined;
  }

  // 2) Real calendar due date (not "Monthly Basis" / day-only).
  const dueRaw = item?.dueDate ?? item?.DueDate;
  if (
    dueRaw != null &&
    String(dueRaw).trim() !== '' &&
    !isMonthlyRecurringDue(item)
  ) {
    const dueTs = parseDateValue(dueRaw);
    if (dueTs != null) {
      const d = new Date(dueTs);
      const key = monthKey(new Date(d.getFullYear(), d.getMonth(), 1));
      if (windowSet.has(key)) return key;
      return undefined;
    }
    const fromDueName = monthIndexToWindowKey(parseMonthIndex(dueRaw), windowKeys);
    if (fromDueName) return fromDueName;
  }

  // 3) Approved / submitted event month (workflow only).
  const workflowTs = itemWorkflowMonthTimestamp(item);
  if (workflowTs != null) {
    const d = new Date(workflowTs);
    const key = monthKey(new Date(d.getFullYear(), d.getMonth(), 1));
    if (windowSet.has(key)) return key;
    return undefined;
  }

  // 4) Open / undated statutory+bulk lines → current month only.
  return windowSet.has(currentKey) ? currentKey : undefined;
}

/**
 * Last-6-months trend.
 * Every month uses the same obligation universe as Compliance Health / top KPI
 * cards: count every statutory row (no form+site dedupe) so totals match (e.g. 111).
 * Past months: apply MonthFilter / workflow history status when available for a
 * form+site; remaining rows count as Yet to Submit.
 */
function buildTrendSeries(items, now = new Date()) {
  const months = [];
  for (let i = 5; i >= 0; i -= 1) {
    const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
    months.push({ key: monthKey(d), label: monthLabel(d), date: d, ...EMPTY_COUNTS });
  }
  const windowKeys = months.map((m) => m.key);
  const currentKey = months[months.length - 1]?.key;
  const list = items || [];

  // Past-month history: form+site → best workflow row tagged for that month.
  const historyByMonth = new Map(windowKeys.map((k) => [k, new Map()]));
  list.forEach((row) => {
    const mk = resolveTrendMonthKey(row, windowKeys, now);
    if (!mk || mk === currentKey) return;
    const base = rowIdentityKey(row, { ignoreMonth: true });
    if (!base) return;
    const map = historyByMonth.get(mk);
    map.set(base, pickStrongerRow(map.get(base), row));
  });

  months.forEach((m) => {
    if (m.key === currentKey) {
      list.forEach((item) => {
        m[getComplianceBucket(item)] += 1;
      });
      return;
    }

    const histMap = historyByMonth.get(m.key) || new Map();
    const usedHistory = new Set();

    // Same row count as Health (111): history status when known, else Yet to Submit.
    list.forEach((item) => {
      const base = rowIdentityKey(item, { ignoreMonth: true });
      if (base && histMap.has(base) && !usedHistory.has(base)) {
        usedHistory.add(base);
        m[getComplianceBucket(histMap.get(base))] += 1;
        return;
      }
      m.yetToSubmit += 1;
    });
  });

  return months.map((m) => {
    const total = m.approved + m.pending + m.yetToSubmit + m.returned;
    return {
      label: m.label,
      approved: m.approved,
      pending: m.pending,
      yetToSubmit: m.yetToSubmit,
      returned: m.returned,
      total,
    };
  });
}

function statutoryRowMatchesSiteScope(row, scope) {
  const actCategories = scope?.actCategories;
  const hasActFilter = Array.isArray(actCategories) && actCategories.length > 0;
  const hasSiteNames = Array.isArray(scope?.siteNames) && scope.siteNames.length > 0;
  const hasStates = Array.isArray(scope?.stateLabels) && scope.stateLabels.length > 0;
  const hasIndustries = Array.isArray(scope?.industryLabels) && scope.industryLabels.length > 0;
  if (!hasActFilter && !hasSiteNames && !hasStates && !hasIndustries) return true;

  const act = row?.act ?? row?.Act ?? '';
  const sector = row?.sector ?? row?.Sector ?? '';

  if (hasActFilter) {
    const cat = getActCategoryFromActSector(act, sector);
    if (!actCategories.includes(cat)) return false;
  }

  if (String(sector || '').trim()) {
    const sectorCat =
      industryLabelToActCategory(sector) || getActCategoryFromActSector('', sector);
    if (hasActFilter && sectorCat && sectorCat !== 'other' && !actCategories.includes(sectorCat)) {
      return false;
    }
    if (
      hasIndustries &&
      sectorCat &&
      sectorCat !== 'other' &&
      !sectorMatchesInchargeSiteIndustries(sector, scope.industryLabels)
    ) {
      return false;
    }
  }

  if (hasStates) {
    // Same as Statutory: blank / All India / other states are excluded for site login.
    const stateField = row?.states ?? row?.state ?? row?.State ?? '';
    if (!statesFieldMatchesInchargeSiteStates(stateField, scope.stateLabels)) {
      return false;
    }
  }

  if (hasSiteNames) {
    const rowSite = String(row?.site ?? row?.Site ?? row?.siteName ?? row?.SiteName ?? '')
      .trim()
      .toLowerCase();
    if (rowSite) {
      const allowed = new Set(scope.siteNames.map((s) => String(s).trim().toLowerCase()).filter(Boolean));
      if (!allowed.has(rowSite)) return false;
    }
  }

  return true;
}

function hasUsableSiteScope(scope) {
  if (!scope || typeof scope !== 'object') return false;
  return (
    (Array.isArray(scope.actCategories) && scope.actCategories.length > 0) ||
    (Array.isArray(scope.siteNames) && scope.siteNames.length > 0) ||
    (Array.isArray(scope.stateLabels) && scope.stateLabels.length > 0) ||
    (Array.isArray(scope.industryLabels) && scope.industryLabels.length > 0)
  );
}

function rowMatchesSelectedSite(row, selectedSite, siteRecords) {
  if (!selectedSite || selectedSite === 'all') return true;
  const selected = String(selectedSite).trim().toLowerCase();
  const rowSite = String(row?.site ?? row?.Site ?? row?.siteName ?? row?.SiteName ?? '')
    .trim()
    .toLowerCase();
  if (rowSite && rowSite === selected) return true;

  const site = (siteRecords || []).find(
    (s) => String(siteNameFromSiteRecord(s) || '').trim().toLowerCase() === selected
  );
  if (!site) return rowSite === selected;

  const state = siteStateFromRecord(site);
  const industry = siteIndustry(site);
  const rowState = row?.states ?? row?.state ?? row?.State ?? '';
  const rowSector = row?.sector ?? row?.Sector ?? '';

  const stateOk =
    !state ||
    !String(rowState || '').trim() ||
    statesFieldMatchesInchargeSiteStates(rowState, [state]);
  const industryOk =
    !industry ||
    !String(rowSector || '').trim() ||
    sectorMatchesInchargeSiteIndustries(rowSector, [industry]);

  return stateOk && industryOk;
}

function periodWindowKeys(period, now = new Date()) {
  const y = now.getFullYear();
  const m = now.getMonth();
  if (period === 'thisMonth') {
    return [monthKey(new Date(y, m, 1))];
  }
  if (period === 'lastMonth') {
    return [monthKey(new Date(y, m - 1, 1))];
  }
  if (period === 'last6Months') {
    const keys = [];
    for (let i = 5; i >= 0; i -= 1) {
      keys.push(monthKey(new Date(y, m - i, 1)));
    }
    return keys;
  }
  if (period === 'thisYear') {
    const keys = [];
    for (let i = 0; i <= m; i += 1) {
      keys.push(monthKey(new Date(y, i, 1)));
    }
    return keys;
  }
  return [monthKey(new Date(y, m, 1))];
}

function itemInPeriod(item, period, now = new Date()) {
  const windowKeys = periodWindowKeys(period, now);
  const key = resolveTrendMonthKey(item, windowKeys, now);
  return Boolean(key);
}

function countBuckets(items) {
  const counts = { ...EMPTY_COUNTS };
  items.forEach((item) => {
    const bucket = getComplianceBucket(item);
    counts[bucket] += 1;
  });
  return counts;
}

function pct(part, total) {
  if (!total) return 0;
  return Math.round((part / total) * 100);
}

function scoreFromCounts(counts) {
  const total =
    counts.approved + counts.pending + counts.yetToSubmit + counts.returned;
  return pct(counts.approved, total);
}

const STATE_CANONICAL_BY_KEY = (() => {
  const map = new Map();
  INDIAN_STATES.forEach((name) => {
    map.set(normalizeStateCompareKey(name), name);
  });
  map.set(normalizeStateCompareKey('Delhi'), 'Delhi');
  map.set(normalizeStateCompareKey('NCT of Delhi'), 'Delhi');
  // Misspellings / short labels that should display as Madhya Pradesh
  ['mathiya', 'madhiya', 'mathya', 'madhya', 'm.p.', 'm.p', 'mp'].forEach((alias) => {
    map.set(normalizeStateCompareKey(alias), 'Madhya Pradesh');
  });
  return map;
})();

function canonicalStateName(raw) {
  const key = normalizeStateCompareKey(raw);
  if (!key) return '';
  if (STATE_CANONICAL_BY_KEY.has(key)) return STATE_CANONICAL_BY_KEY.get(key);
  for (const [canonKey, name] of STATE_CANONICAL_BY_KEY.entries()) {
    // Skip short aliases (mp, etc.) for fuzzy includes — exact match only above.
    if (canonKey.length < 5) continue;
    if (key.includes(canonKey) || (key.length >= 5 && canonKey.includes(key))) return name;
  }
  // "mathiya" / similar typos of Madhya Pradesh
  if (/^ma[dt]h+i?y+a/.test(key)) {
    return 'Madhya Pradesh';
  }
  const cleaned = String(raw || '').trim();
  return cleaned;
}

function rowPrimaryState(item, siteRecords) {
  const field = item?.states ?? item?.state ?? item?.State ?? '';
  const tokens = splitStateFieldTokens(field);
  const fromTokens = tokens.length > 0 ? canonicalStateName(tokens[0]) : '';
  const fromField = fromTokens || canonicalStateName(field);

  const rowSite = String(item?.site ?? item?.Site ?? item?.siteName ?? item?.SiteName ?? '')
    .trim()
    .toLowerCase();
  let fromSite = '';
  if (rowSite && Array.isArray(siteRecords)) {
    const site = siteRecords.find(
      (s) => String(siteNameFromSiteRecord(s) || '').trim().toLowerCase() === rowSite
    );
    if (site) {
      // Prefer Site Management state (e.g. Madhya Pradesh) over misspelled statutory labels (mathiya).
      fromSite = canonicalStateName(siteStateFromRecord(site));
      if (!fromSite) {
        fromSite = canonicalStateName(siteNameFromSiteRecord(site));
      }
    } else {
      // Site name itself may be Madhya Pradesh / mathiya when no site record match.
      fromSite = canonicalStateName(rowSite);
    }
  }

  // If statutory state is an unknown/misspelled label but site resolves to a known state, use site.
  const fromFieldKnown =
    fromField && STATE_CANONICAL_BY_KEY.has(normalizeStateCompareKey(fromField));
  if (fromSite && (!fromField || !fromFieldKnown)) return fromSite;
  if (fromField) return fromField;
  return fromSite || '';
}

function dominantStatus(counts) {
  let best = 'yetToSubmit';
  let bestVal = -1;
  LOCATION_STATUS_ORDER.forEach((key) => {
    const v = counts[key] || 0;
    if (v > bestVal) {
      bestVal = v;
      best = key;
    }
  });
  return best;
}

function buildLocationSeries(items, siteRecords) {
  const map = new Map();
  items.forEach((item) => {
    const state = rowPrimaryState(item, siteRecords);
    if (!state) return;
    if (!map.has(state)) map.set(state, { state, ...EMPTY_COUNTS });
    map.get(state)[getComplianceBucket(item)] += 1;
  });

  return [...map.values()]
    .map((row) => {
      const total = row.approved + row.pending + row.yetToSubmit + row.returned;
      const score = pct(row.approved, total);
      const coords = STATE_CENTROIDS[normalizeStateCompareKey(row.state)] || null;
      return {
        ...row,
        total,
        score,
        statusColor: STATUS_COLORS[dominantStatus(row)],
        lat: coords?.[0] ?? null,
        lng: coords?.[1] ?? null,
      };
    })
    .filter((row) => row.total > 0)
    .sort((a, b) => b.total - a.total || a.state.localeCompare(b.state))
    .slice(0, 8)
    .map((row) => ({
      ...row,
      // Fixed map pin color per state (bars still use approved green).
      color: mapColorForState(row.state),
    }));
}

function industryGroupLabel(item) {
  const act = item?.act ?? item?.Act ?? '';
  const sector = item?.sector ?? item?.Sector ?? '';
  const cat =
    industryLabelToActCategory(sector) || getActCategoryFromActSector(act, sector);
  if (cat === 'clra') return 'CLRA';
  if (cat === 'factories') return 'Factory';
  if (cat === 'shops_and_establishment') return 'S&E';
  return null;
}

/**
 * Build industry bar series. When `allowedCategories` is set (site login),
 * only those act buckets are shown — e.g. Factories Act login → Factory only.
 */
function buildIndustrySeries(items, allowedCategories = null) {
  const groups =
    Array.isArray(allowedCategories) && allowedCategories.length > 0
      ? INDUSTRY_GROUPS.filter((g) => allowedCategories.includes(g.category))
      : INDUSTRY_GROUPS;

  // No matching industry for this login — show nothing rather than all three buckets.
  if (groups.length === 0) {
    return [];
  }

  const byKey = Object.fromEntries(
    groups.map((g) => [g.label, { industry: g.label, ...EMPTY_COUNTS, total: 0 }])
  );

  items.forEach((item) => {
    const label = industryGroupLabel(item);
    if (!label || !byKey[label]) return;
    const bucket = getComplianceBucket(item);
    byKey[label][bucket] += 1;
    byKey[label].total += 1;
  });

  return groups.map((g) => byKey[g.label]);
}

/** Display label for an act row (site “Compliance by Act” chart). */
function actDisplayName(item) {
  const raw = String(item?.act ?? item?.Act ?? item?.actName ?? item?.ActName ?? '').trim();
  if (!raw) return 'Other Acts';
  const lower = raw.toLowerCase();
  if (/factor/.test(lower)) return 'Factories Act, 1948';
  if (/\besi\b|employees.?state.?insurance/.test(lower)) return 'ESI Act, 1948';
  if (/\bepf\b|provident.?fund|employees.?provident/.test(lower)) return 'EPF Act, 1952';
  if (/shops?\s*&?\s*est/.test(lower) || /shops and establishment/.test(lower)) {
    return 'Shops & Establishments Act';
  }
  if (/professional.?tax|profession.?tax/.test(lower)) return 'Professional Tax Act';
  if (/clra|contract.?labour/.test(lower)) return 'CLRA Act, 1970';
  if (/minimum.?wages/.test(lower)) return 'Minimum Wages Act';
  if (/payment.?of.?wages/.test(lower)) return 'Payment of Wages Act';
  if (/payment.?of.?bonus/.test(lower)) return 'Payment of Bonus Act';
  if (/maternity/.test(lower)) return 'Maternity Benefit Act';
  if (/industrial.?disputes/.test(lower)) return 'Industrial Disputes Act';
  return raw.length > 40 ? `${raw.slice(0, 38)}…` : raw;
}

/**
 * Horizontal stacked-bar series by Act (site login chart model).
 */
function buildActSeries(items) {
  const map = new Map();
  (items || []).forEach((item) => {
    const name = actDisplayName(item);
    if (!map.has(name)) {
      map.set(name, { act: name, ...EMPTY_COUNTS, total: 0 });
    }
    const row = map.get(name);
    row[getComplianceBucket(item)] += 1;
    row.total += 1;
  });

  return [...map.values()]
    .filter((row) => row.total > 0)
    .sort((a, b) => b.total - a.total || a.act.localeCompare(b.act))
    .slice(0, 8);
}

async function fetchSiteDetails() {
  try {
    const res = await fetch(SITE_MANAGEMENT_API, { cache: 'no-store' });
    if (!res.ok) return [];
    const json = await res.json();
    return Array.isArray(json?.data?.siteDetails) ? json.data.siteDetails : [];
  } catch {
    return [];
  }
}

function StatusLegend({ compact = false, order, square = false, labels = STATUS_LABELS }) {
  const keys = order || Object.keys(labels);
  return (
    <div className={`chd-legend${compact ? ' chd-legend--compact' : ''}`}>
      {keys.map((key) => (
        <span key={key} className="chd-legend-item">
          <span
            className={square ? 'chd-swatch' : 'chd-dot'}
            style={{ background: STATUS_COLORS[key] }}
          />
          {labels[key] || STATUS_LABELS[key]}
        </span>
      ))}
    </div>
  );
}

function MapInvalidateSize() {
  const map = useMap();
  useEffect(() => {
    const t = setTimeout(() => map.invalidateSize(), 80);
    return () => clearTimeout(t);
  }, [map]);
  return null;
}

/** Fit full India (incl. western states like Gujarat) inside the map box. */
function FitIndiaBounds({ geojson, locations }) {
  const map = useMap();
  useEffect(() => {
    if (!geojson) return undefined;

    const applyFit = () => {
      try {
        map.invalidateSize();
        const layer = L.geoJSON(geojson);
        let bounds = layer.getBounds();
        if (!bounds.isValid()) return;

        // Include marker positions so pins (e.g. Gujarat purple) are never clipped.
        (locations || []).forEach((loc) => {
          if (Number.isFinite(loc.lat) && Number.isFinite(loc.lng)) {
            bounds.extend([loc.lat, loc.lng]);
          }
        });

        // Light geographic padding so western pins stay inside the frame.
        bounds = bounds.pad(0.02);
        map.fitBounds(bounds, {
          paddingTopLeft: [10, 8],
          paddingBottomRight: [10, 8],
          maxZoom: 6,
          animate: false,
        });
        // Slight zoom-in so India fills the box without large empty margins.
        const nextZoom = Math.min((map.getZoom() || 4) + 0.35, 6);
        map.setZoom(nextZoom, { animate: false });
      } catch {
        /* ignore invalid geojson */
      }
    };

    applyFit();
    const t = setTimeout(applyFit, 120);
    return () => clearTimeout(t);
  }, [map, geojson, locations]);
  return null;
}

const INDIA_GEOJSON_URL = `${process.env.PUBLIC_URL || ''}/data/india-states.geojson`;

const INDIA_MAP_STYLE = {
  fillColor: '#d1d5db',
  fillOpacity: 1,
  color: '#9ca3af',
  weight: 0.8,
  opacity: 1,
};

/** Coloured teardrop pin for each state on the India map. */
function createStatePointerIcon(color) {
  const fill = color || STATUS_COLORS.yetToSubmit;
  return L.divIcon({
    className: 'chd-map-pointer',
    html: `
      <div class="chd-map-pointer-pin" style="background-color:${fill};border-color:${fill};">
        <span class="chd-map-pointer-dot"></span>
      </div>
    `,
    iconSize: [22, 22],
    iconAnchor: [11, 22],
    popupAnchor: [0, -20],
  });
}

function IndiaLocationMap({ locations }) {
  const [indiaGeo, setIndiaGeo] = useState(null);
  const markers = useMemo(
    () =>
      (locations || [])
        .filter((loc) => Number.isFinite(loc.lat) && Number.isFinite(loc.lng))
        .map((loc) => ({
          ...loc,
          icon: createStatePointerIcon(loc.color || mapColorForState(loc.state)),
        })),
    [locations]
  );

  useEffect(() => {
    let cancelled = false;
    fetch(INDIA_GEOJSON_URL)
      .then((res) => (res.ok ? res.json() : null))
      .then((json) => {
        if (!cancelled && json) setIndiaGeo(json);
      })
      .catch(() => {
        if (!cancelled) setIndiaGeo(null);
      });
    return () => {
      cancelled = true;
    };
  }, []);

  return (
    <MapContainer
      center={[22.5, 80]}
      zoom={4}
      minZoom={3}
      maxZoom={6}
      zoomSnap={0.25}
      zoomDelta={0.25}
      scrollWheelZoom={false}
      dragging={false}
      doubleClickZoom={false}
      touchZoom={false}
      boxZoom={false}
      keyboard={false}
      zoomControl={false}
      attributionControl={false}
      className="chd-india-map"
      style={{ height: '100%', width: '100%', background: '#f3f4f6', cursor: 'default' }}
      maxBounds={L.latLngBounds([5, 66], [38, 100])}
      maxBoundsViscosity={1}
    >
      <MapInvalidateSize />
      <FitIndiaBounds geojson={indiaGeo} locations={markers} />
      {indiaGeo ? <GeoJSON data={indiaGeo} style={() => INDIA_MAP_STYLE} /> : null}
      {markers.map((loc) => (
        <Marker key={loc.state} position={[loc.lat, loc.lng]} icon={loc.icon} zIndexOffset={600}>
          <Popup>
            <strong>{loc.state}</strong>
            <br />
            {loc.score}% · {loc.approved} of {loc.total} compliances
          </Popup>
        </Marker>
      ))}
    </MapContainer>
  );
}

function LocationBars({ rows }) {
  if (!rows.length) {
    return <div className="chd-empty">No location compliance data yet.</div>;
  }

  return (
    <div className="chd-loc-list">
      {rows.map((row) => {
        const approvedPct = row.total ? (row.approved / row.total) * 100 : 0;
        const stateColor = row.color || mapColorForState(row.state);
        return (
          <div key={row.state} className="chd-loc-row">
            <span className="chd-loc-name" title={row.state}>
              {row.state}
            </span>
            <span className="chd-loc-score" style={{ color: stateColor }}>
              {row.score}%
            </span>
            <div className="chd-loc-bar-track" aria-hidden>
              {approvedPct > 0 ? (
                <span
                  className="chd-loc-bar-seg"
                  style={{
                    width: `${approvedPct}%`,
                    background: stateColor,
                  }}
                  title={`Approved: ${row.approved}`}
                />
              ) : null}
            </div>
          </div>
        );
      })}
    </div>
  );
}

function IndustryBars({ series }) {
  const width = 420;
  const height = 260;
  const pad = { top: 28, right: 16, bottom: 40, left: 44 };
  const chartW = width - pad.left - pad.right;
  const chartH = height - pad.top - pad.bottom;
  const keys = INDUSTRY_STATUS_ORDER;
  const rows = (series || []).filter((row) => (Number(row.total) || 0) > 0);
  const rawMax = Math.max(
    0,
    ...rows.flatMap((row) => keys.map((key) => Number(row[key]) || 0))
  );
  // Scale to data — avoid a forced 0–30 axis that makes small current-month counts look wrong.
  const niceStep = rawMax <= 10 ? 2 : rawMax <= 25 ? 5 : 10;
  const maxY = Math.max(niceStep, Math.ceil(rawMax / niceStep) * niceStep);
  const ticks = [];
  for (let v = 0; v <= maxY; v += niceStep) ticks.push(v);

  if (!rows.length) {
    return <div className="chd-empty">No industry compliance data yet.</div>;
  }

  const groupCount = Math.max(rows.length, 1);
  const groupW = chartW / groupCount;
  const barGap = 4;
  const barW = Math.min(22, Math.max(10, (groupW - 24) / keys.length - barGap));
  const clusterW = keys.length * (barW + barGap) - barGap;

  const yAt = (v) => pad.top + chartH - (Math.min(v, maxY) / maxY) * chartH;

  return (
    <svg
      className="chd-industry-svg"
      viewBox={`0 0 ${width} ${height}`}
      role="img"
      aria-label="Compliance by industry"
    >
      <text
        x={12}
        y={pad.top + chartH / 2}
        textAnchor="middle"
        fontSize="10"
        fill="#9ca3af"
        transform={`rotate(-90 12 ${pad.top + chartH / 2})`}
      >
        No. of Compliances
      </text>
      {ticks.map((v) => (
        <g key={v}>
          <line
            x1={pad.left}
            y1={yAt(v)}
            x2={pad.left + chartW}
            y2={yAt(v)}
            stroke="#eef2f7"
            strokeWidth="1"
          />
          <text x={pad.left - 8} y={yAt(v) + 4} textAnchor="end" fontSize="10" fill="#9ca3af">
            {v}
          </text>
        </g>
      ))}
      {rows.map((row, gi) => {
        const groupX = pad.left + gi * groupW + (groupW - clusterW) / 2;
        return (
          <g key={row.industry}>
            {keys.map((key, ki) => {
              const value = Number(row[key]) || 0;
              if (value <= 0) return null;
              const x = groupX + ki * (barW + barGap);
              const y = yAt(value);
              const h = Math.max(2, pad.top + chartH - y);
              return (
                <g key={key}>
                  <rect
                    x={x}
                    y={y}
                    width={barW}
                    height={h}
                    fill={STATUS_COLORS[key]}
                    rx="2"
                  />
                  <text
                    x={x + barW / 2}
                    y={y - 6}
                    textAnchor="middle"
                    fontSize="10"
                    fontWeight="600"
                    fill="#374151"
                  >
                    {value}
                  </text>
                </g>
              );
            })}
            <text
              x={pad.left + gi * groupW + groupW / 2}
              y={height - 12}
              textAnchor="middle"
              fontSize="11"
              fill="#6b7280"
              fontWeight="600"
            >
              {row.industry}
            </text>
          </g>
        );
      })}
    </svg>
  );
}

/** Site-login horizontal stacked bars (Compliance by Act image model). */
function ActStackedBars({ series }) {
  const keys = ACT_CHART_STATUS_ORDER;
  const maxTotal = Math.max(1, ...(series || []).map((row) => Number(row.total) || 0));
  // Nice X-axis upper bound (10 / 20 / 30 …)
  const axisMax = Math.max(10, Math.ceil(maxTotal / 10) * 10);
  const ticks = [];
  const step = axisMax <= 20 ? 5 : 10;
  for (let v = 0; v <= axisMax; v += step) ticks.push(v);

  if (!series || series.length === 0) {
    return <div className="chd-empty">No act compliance data yet.</div>;
  }

  return (
    <div className="chd-act-chart" role="img" aria-label="Compliance by act">
      <div className="chd-act-rows">
        {series.map((row) => (
          <div key={row.act} className="chd-act-row">
            <div className="chd-act-name" title={row.act}>
              {row.act}
            </div>
            <div className="chd-act-bar-wrap">
              <div className="chd-act-bar" style={{ width: `${(row.total / axisMax) * 100}%` }}>
                {keys.map((key) => {
                  const value = Number(row[key]) || 0;
                  if (value <= 0) return null;
                  const pctW = row.total ? (value / row.total) * 100 : 0;
                  return (
                    <span
                      key={key}
                      className="chd-act-seg"
                      style={{
                        width: `${pctW}%`,
                        background: STATUS_COLORS[key],
                      }}
                      title={`${ACT_CHART_LABELS[key]}: ${value}`}
                    >
                      {value >= 1 ? value : ''}
                    </span>
                  );
                })}
              </div>
            </div>
            <div className="chd-act-total">{row.total}</div>
          </div>
        ))}
      </div>
      <div className="chd-act-axis" aria-hidden>
        <span className="chd-act-axis-spacer" />
        <div className="chd-act-axis-track">
          {ticks.map((v) => (
            <span
              key={v}
              className="chd-act-tick"
              style={{ left: `${(v / axisMax) * 100}%` }}
            >
              {v}
            </span>
          ))}
        </div>
        <span className="chd-act-axis-end" />
      </div>
      <div className="chd-act-axis-label">No. of Compliances</div>
    </div>
  );
}

function ComplianceGauge({ score, counts }) {
  const total =
    counts.approved + counts.pending + counts.yetToSubmit + counts.returned;
  const r = 70;
  const cx = 90;
  const cy = 88;
  const startAngle = Math.PI;
  const endAngle = 0;
  const sweep = Math.PI;

  const segments = [
    { key: 'approved', value: counts.approved },
    { key: 'pending', value: counts.pending },
    { key: 'yetToSubmit', value: counts.yetToSubmit },
    { key: 'returned', value: counts.returned },
  ].filter((s) => s.value > 0);

  let offset = 0;
  const arcs =
    total === 0
      ? [
          {
            key: 'empty',
            d: describeArc(cx, cy, r, startAngle, endAngle),
            color: '#e5e7eb',
          },
        ]
      : segments.map((seg) => {
          const frac = seg.value / total;
          const a0 = startAngle - offset * sweep;
          offset += frac;
          const a1 = startAngle - offset * sweep;
          return {
            key: seg.key,
            d: describeArc(cx, cy, r, a0, a1),
            color: STATUS_COLORS[seg.key],
          };
        });

  return (
    <div className="chd-gauge">
      <svg viewBox="0 0 180 110" className="chd-gauge-svg" aria-hidden>
        {arcs.map((arc) => (
          <path
            key={arc.key}
            d={arc.d}
            fill="none"
            stroke={arc.color}
            strokeWidth="16"
            strokeLinecap="butt"
          />
        ))}
      </svg>
      <div className="chd-gauge-center">
        <strong>{score}%</strong>
        <span>Overall Compliance Score</span>
      </div>
    </div>
  );
}

function polar(cx, cy, r, angle) {
  return { x: cx + r * Math.cos(angle), y: cy - r * Math.sin(angle) };
}

function describeArc(cx, cy, r, startAngle, endAngle) {
  const start = polar(cx, cy, r, startAngle);
  const end = polar(cx, cy, r, endAngle);
  const largeArc = Math.abs(startAngle - endAngle) > Math.PI ? 1 : 0;
  const sweepFlag = startAngle > endAngle ? 1 : 0;
  return `M ${start.x} ${start.y} A ${r} ${r} 0 ${largeArc} ${sweepFlag} ${end.x} ${end.y}`;
}

/** Image-1 line model: open inventory series. Click a month to pin its counts (local only). */
function TrendChart({ series }) {
  const wrapRef = useRef(null);
  const [hover, setHover] = useState(null);
  // Clicked month stays local — Health / Location / Industry keep current-month data.
  const [selectedIndex, setSelectedIndex] = useState(null);
  const width = 560;
  const height = 240;
  const pad = { top: 12, right: 18, bottom: 30, left: 40 };
  const chartW = width - pad.left - pad.right;
  const chartH = height - pad.top - pad.bottom;
  // Draw order: Yet to Submit / Returned first so Approved stays visible when overlapping.
  const keys = ['yetToSubmit', 'returned', 'pending', 'approved'];
  const n = Math.max(series.length, 1);
  const rawMax = Math.max(
    0,
    ...series.flatMap((row) => keys.map((key) => Number(row[key]) || 0))
  );
  // Image-1 model: 0–200 with steps of 50 when volume is above 100; else 0–100 / 25.
  const maxY = rawMax <= 100 ? 100 : Math.max(200, Math.ceil(rawMax / 50) * 50);
  const ticks = [];
  const step = maxY <= 100 ? 25 : 50;
  for (let v = 0; v <= maxY; v += step) ticks.push(v);

  const xAt = (i) => pad.left + (n === 1 ? chartW / 2 : (i / (n - 1)) * chartW);
  const yAt = (v) => pad.top + chartH - (Math.min(Math.max(v, 0), maxY) / maxY) * chartH;
  const hasData = series.some((s) => (s.total || 0) > 0);
  const bandW = n === 1 ? chartW : chartW / Math.max(n - 1, 1);

  const paths = keys.map((key) => {
    const pts = series.map((row, i) => ({
      x: xAt(i),
      y: yAt(Number(row[key]) || 0),
    }));
    const d = pts.length
      ? pts.map((p, i) => `${i === 0 ? 'M' : 'L'} ${p.x} ${p.y}`).join(' ')
      : '';
    return { key, d };
  });

  const scaleTip = (index, focusKey) => {
    const row = series[index];
    if (!row) return null;
    const rect = wrapRef.current?.getBoundingClientRect();
    const scaleX = rect ? rect.width / width : 1;
    const scaleY = rect ? rect.height / height : 1;
    const wrapW = rect?.width ?? width;
    const tipHalf = 84; // ~half of tooltip min-width + padding
    const edgePad = 8;
    const focusValue =
      focusKey != null
        ? Number(row[focusKey]) || 0
        : Math.max(
            Number(row.approved) || 0,
            Number(row.pending) || 0,
            Number(row.yetToSubmit) || 0,
            Number(row.returned) || 0
          );
    let x = xAt(index) * scaleX;
    let xAlign = 'center';
    // Keep tooltip inside the chart when hovering edge months (e.g. Aug).
    if (x + tipHalf > wrapW - edgePad) {
      xAlign = 'end';
      x = Math.min(x, wrapW - edgePad);
    } else if (x - tipHalf < edgePad) {
      xAlign = 'start';
      x = Math.max(x, edgePad);
    }
    return {
      index,
      focusKey,
      label: row.label,
      counts: {
        approved: Number(row.approved) || 0,
        pending: Number(row.pending) || 0,
        yetToSubmit: Number(row.yetToSubmit) || 0,
        returned: Number(row.returned) || 0,
      },
      x,
      y: yAt(focusValue) * scaleY,
      xAlign,
    };
  };

  const nearestKeyAtY = (row, svgY) => {
    let nearestKey = keys[0];
    let best = Infinity;
    keys.forEach((key) => {
      const dist = Math.abs(yAt(Number(row[key]) || 0) - svgY);
      if (dist < best) {
        best = dist;
        nearestKey = key;
      }
    });
    return nearestKey;
  };

  const focusFromEvent = (row, e) => {
    const svg = e.currentTarget.ownerSVGElement || e.currentTarget;
    const rect = svg.getBoundingClientRect();
    const svgY = ((e.clientY - rect.top) / rect.height) * height;
    return nearestKeyAtY(row, svgY);
  };

  const selectMonth = (index, focusKey) => {
    if (selectedIndex === index) {
      setSelectedIndex(null);
      setHover(null);
      return;
    }
    setSelectedIndex(index);
    setHover(scaleTip(index, focusKey));
  };

  const tipIndex = hover?.index ?? selectedIndex;
  const activeTip =
    tipIndex == null
      ? null
      : hover?.index === tipIndex
        ? hover
        : scaleTip(tipIndex, null);
  const guideIndex = tipIndex;

  return (
    <div
      ref={wrapRef}
      className="chd-trend-wrap"
      onMouseLeave={() => setHover(null)}
    >
      <svg className="chd-trend-svg" viewBox={`0 0 ${width} ${height}`} role="img" aria-label="Compliance trend">
        {ticks.map((v) => (
          <g key={v}>
            <line
              x1={pad.left}
              y1={yAt(v)}
              x2={pad.left + chartW}
              y2={yAt(v)}
              stroke="#eef2f7"
              strokeWidth="1"
            />
            <text x={pad.left - 10} y={yAt(v) + 4} textAnchor="end" fontSize="10" fill="#9ca3af">
              {v}
            </text>
          </g>
        ))}
        {guideIndex != null && (
          <line
            x1={xAt(guideIndex)}
            y1={pad.top}
            x2={xAt(guideIndex)}
            y2={pad.top + chartH}
            stroke={selectedIndex === guideIndex ? '#94a3b8' : '#cbd5e1'}
            strokeWidth={selectedIndex === guideIndex ? 1.5 : 1}
            strokeDasharray="4 3"
            pointerEvents="none"
          />
        )}
        {paths.map((p) => (
          <path
            key={p.key}
            d={p.d}
            fill="none"
            stroke={STATUS_COLORS[p.key]}
            strokeWidth="2"
            strokeLinejoin="round"
            strokeLinecap="round"
            strokeOpacity={hasData ? 1 : 0.35}
            pointerEvents="none"
          />
        ))}
        {keys.map((key) =>
          series.map((row, i) => {
            const value = Number(row[key]) || 0;
            const active =
              (hover?.focusKey === key && hover?.index === i) || selectedIndex === i;
            return (
              <circle
                key={`${key}-${i}`}
                cx={xAt(i)}
                cy={yAt(value)}
                r={active ? 5.5 : 4}
                fill={STATUS_COLORS[key]}
                stroke="#fff"
                strokeWidth={active ? 2.5 : 1.5}
                opacity={hasData ? 1 : 0.35}
                pointerEvents="none"
              />
            );
          })
        )}
        {/* Count labels for the clicked month only */}
        {selectedIndex != null &&
          series[selectedIndex] &&
          TREND_STATUS_ORDER.map((key, ki) => {
            const row = series[selectedIndex];
            const value = Number(row[key]) || 0;
            const x = xAt(selectedIndex);
            const y = yAt(value);
            const offsetX = (ki - 1.5) * 18;
            return (
              <text
                key={`count-${key}`}
                x={x + offsetX}
                y={y - 10}
                textAnchor="middle"
                fontSize="10"
                fontWeight="700"
                fill={STATUS_COLORS[key]}
                pointerEvents="none"
              >
                {value}
              </text>
            );
          })}
        {series.map((row, i) => (
          <text
            key={row.label}
            x={xAt(i)}
            y={height - 8}
            textAnchor="middle"
            fontSize="11"
            fontWeight={selectedIndex === i ? 700 : 400}
            fill={selectedIndex === i ? '#111827' : '#6b7280'}
            style={{ cursor: 'pointer' }}
            onClick={(e) => {
              e.stopPropagation();
              selectMonth(i, 'approved');
            }}
          >
            {row.label}
          </text>
        ))}
        {/* Hit bands: click a month to pin counts; hover still previews */}
        {series.map((row, i) => (
          <rect
            key={`band-${row.label}`}
            x={xAt(i) - bandW / 2}
            y={pad.top}
            width={bandW}
            height={chartH}
            fill="transparent"
            style={{ cursor: 'pointer' }}
            onMouseEnter={(e) => setHover(scaleTip(i, focusFromEvent(row, e)))}
            onMouseMove={(e) => setHover(scaleTip(i, focusFromEvent(row, e)))}
            onClick={(e) => selectMonth(i, focusFromEvent(row, e))}
          />
        ))}
      </svg>
      {activeTip && (
        <div
          className={`chd-trend-tooltip${activeTip.y < 72 ? ' chd-trend-tooltip--below' : ''}${
            activeTip.xAlign === 'start' ? ' chd-trend-tooltip--align-start' : ''
          }${activeTip.xAlign === 'end' ? ' chd-trend-tooltip--align-end' : ''}${
            selectedIndex === activeTip.index ? ' chd-trend-tooltip--pinned' : ''
          }`}
          style={{ left: activeTip.x, top: activeTip.y }}
          role="tooltip"
        >
          <div className="chd-trend-tooltip-month">{activeTip.label}</div>
          {TREND_STATUS_ORDER.map((key) => (
            <div
              key={key}
              className={`chd-trend-tooltip-row${
                activeTip.focusKey === key ? ' is-active' : ''
              }`}
            >
              <span className="chd-dot" style={{ background: STATUS_COLORS[key] }} />
              <span>{STATUS_LABELS[key]}</span>
              <strong>{activeTip.counts[key]}</strong>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

export default function ComplianceHealthDashboard({ userRole, userEmail, onCountsChange }) {
  const [loading, setLoading] = useState(true);
  const [period, setPeriod] = useState('thisMonth');
  const [selectedSite, setSelectedSite] = useState('all');
  const [siteOptions, setSiteOptions] = useState([]);
  const [isSiteScopedUser, setIsSiteScopedUser] = useState(false);
  /** Site login act categories (factories | clra | shops_and_establishment); null = show all. */
  const [allowedActCategories, setAllowedActCategories] = useState(null);
  const [rawItems, setRawItems] = useState([]);
  const [siteRecords, setSiteRecords] = useState([]);
  const [refreshAt, setRefreshAt] = useState(0);

  const isAdmin = isOrgWideSiteViewer(userRole);

  useEffect(() => {
    let cancelled = false;

    const load = async () => {
      try {
        const [statutoryRes, bulkRes, siteDetails, siteScope, loginEmail] = await Promise.all([
          fetch(STATUTORY_API, { cache: 'no-store' }),
          fetch(CHECKLIST_BULK_API, { cache: 'no-store' }),
          fetchSiteDetails(),
          fetchInchargeDisplayScopeFromSites(userEmail),
          resolveLoginEmailString(userEmail),
        ]);

        const hasSiteScope = hasUsableSiteScope(siteScope);
        const visibleSites = filterSitesForLoginUser(siteDetails, loginEmail, userRole);
        const inchargeScope = buildInchargeSiteScopeFromList(siteDetails, loginEmail);
        // Same lens as top KPI cards: if this login is a Site Incharge, always
        // scope compliance counts to that assignment (even when role is HR Admin).
        const preferInchargeLens = Boolean(
          hasSiteScope || (inchargeScope?.mine && inchargeScope.mine.length > 0)
        );

        const options = (preferInchargeLens ? visibleSites : siteDetails)
          .map((s) => {
            const name = siteNameFromSiteRecord(s);
            return name ? { value: name, label: name } : null;
          })
          .filter(Boolean);

        let statutory = [];
        if (statutoryRes.ok) {
          const json = await statutoryRes.json();
          statutory = Array.isArray(json?.data?.statutoryData) ? json.data.statutoryData : [];
        }

        let bulk = [];
        if (bulkRes.ok) {
          const json = await bulkRes.json();
          bulk = Array.isArray(json?.data) ? json.data : Array.isArray(json?.data?.checklistBulkData)
            ? json.data.checklistBulkData
            : [];
        }

        // Same idea as Statutory page: unique form lines with live approval workflow fields.
        const merged = mergeStatutoryWithBulk(statutory, bulk);
        // Site Incharge login: only Tamil Nadu Factories (etc.) — match Statutory list size.
        // Org-wide admin with no incharge assignment: full dataset.
        const scoped =
          preferInchargeLens && hasSiteScope
            ? merged.filter((item) => statutoryRowMatchesSiteScope(item, siteScope))
            : preferInchargeLens && inchargeScope
              ? merged.filter((item) =>
                  statutoryRowMatchesSiteScope(item, {
                    actCategories: [
                      ...new Set(
                        (inchargeScope.industryLabels || [])
                          .map((ind) => industryLabelToActCategory(ind))
                          .filter(Boolean)
                      ),
                    ],
                    industryLabels: inchargeScope.industryLabels || [],
                    stateLabels: inchargeScope.stateLabels || [],
                    siteNames: (inchargeScope.mine || [])
                      .map((s) => siteNameFromSiteRecord(s))
                      .filter(Boolean),
                  })
                )
              : merged;

        if (!cancelled) {
          setIsSiteScopedUser(preferInchargeLens);
          let actCats = null;
          if (preferInchargeLens) {
            if (Array.isArray(siteScope?.actCategories) && siteScope.actCategories.length > 0) {
              actCats = siteScope.actCategories;
            } else if (Array.isArray(siteScope?.industryLabels) && siteScope.industryLabels.length > 0) {
              const derived = [
                ...new Set(
                  siteScope.industryLabels
                    .map((ind) => industryLabelToActCategory(ind))
                    .filter(Boolean)
                ),
              ];
              actCats = derived.length > 0 ? derived : null;
            } else if (inchargeScope?.industryLabels?.length) {
              const derived = [
                ...new Set(
                  inchargeScope.industryLabels
                    .map((ind) => industryLabelToActCategory(ind))
                    .filter(Boolean)
                ),
              ];
              actCats = derived.length > 0 ? derived : null;
            }
          }
          setAllowedActCategories(actCats);
          setSiteOptions(options);
          setSiteRecords(preferInchargeLens ? visibleSites : siteDetails);
          if (preferInchargeLens && options.length === 1) {
            setSelectedSite(options[0].value);
          }
          setRawItems(scoped);
          setLoading(false);
        }
      } catch {
        if (!cancelled) {
          setRawItems([]);
          setSiteRecords([]);
          setAllowedActCategories(null);
          setLoading(false);
        }
      }
    };

    load();
    const timer = setInterval(() => {
      setRefreshAt((n) => n + 1);
    }, POLL_MS);

    const bump = () => setRefreshAt((n) => n + 1);
    const onStorage = (e) => {
      if (
        e?.key === 'statutoryDataRevision' ||
        e?.key === 'checklistBulkData' ||
        e?.key === 'siteManagementData'
      ) {
        bump();
      }
    };
    window.addEventListener('statutoryDataUpdated', bump);
    window.addEventListener('storage', onStorage);

    return () => {
      cancelled = true;
      clearInterval(timer);
      window.removeEventListener('statutoryDataUpdated', bump);
      window.removeEventListener('storage', onStorage);
    };
    // refreshAt triggers polling reloads
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [userEmail, userRole, refreshAt]);

  const siteScopedItems = useMemo(() => {
    return rawItems.filter((item) =>
      rowMatchesSelectedSite(item, selectedSite, siteRecords)
    );
  }, [rawItems, selectedSite, siteRecords]);

  // Health / Location / Industry: same model as top KPI cards — count every
  // statutory row (no form+site month dedupe). Trend uses buildTrendSeries.
  const items = useMemo(() => {
    if (period === 'thisMonth' || period === 'last6Months') {
      return siteScopedItems;
    }
    return siteScopedItems.filter((item) => itemInPeriod(item, period));
  }, [siteScopedItems, period]);

  const lastMonthItems = useMemo(() => {
    return siteScopedItems.filter((item) => itemInPeriod(item, 'lastMonth'));
  }, [siteScopedItems]);

  const counts = useMemo(() => countBuckets(items), [items]);
  const total =
    counts.approved + counts.pending + counts.yetToSubmit + counts.returned;
  const score = scoreFromCounts(counts);
  const lastMonthScore = scoreFromCounts(countBuckets(lastMonthItems));
  const delta = score - lastMonthScore;

  // Keep top Compliance Status scorecards in sync with Health (same inventory + buckets).
  useEffect(() => {
    if (typeof onCountsChange !== 'function') return;
    onCountsChange(counts);
  }, [counts, onCountsChange]);

  const trend = useMemo(() => buildTrendSeries(siteScopedItems), [siteScopedItems]);

  const locations = useMemo(
    () =>
      isAdmin && !isSiteScopedUser
        ? buildLocationSeries(items, siteRecords)
        : [],
    [isAdmin, isSiteScopedUser, items, siteRecords]
  );
  const industries = useMemo(() => {
    if (isSiteScopedUser) return [];
    let cats = null;

    // When a single site is selected, show only that site's industry bar(s).
    if (selectedSite && selectedSite !== 'all') {
      const site = (siteRecords || []).find(
        (s) =>
          String(siteNameFromSiteRecord(s) || '').trim().toLowerCase() ===
          String(selectedSite).trim().toLowerCase()
      );
      const cat = site ? industryLabelToActCategory(siteIndustry(site)) : null;
      if (cat) cats = [cat];
    }

    return buildIndustrySeries(items, cats);
  }, [items, isSiteScopedUser, selectedSite, siteRecords]);

  const acts = useMemo(
    () => (isSiteScopedUser ? buildActSeries(items) : []),
    [isSiteScopedUser, items]
  );

  const showSiteFilter = isSiteScopedUser
    ? siteOptions.length > 1
    : isAdmin || siteOptions.length > 1;
  const siteSelectDisabled = isSiteScopedUser && siteOptions.length <= 1;
  // Org-wide admin only — hide for site / incharge login (even if role is HR Admin).
  const showComplianceByLocation = !loading && isAdmin && !isSiteScopedUser;

  return (
    <div className="chd-root" aria-busy={loading || undefined}>
      <div className="chd-filters chd-filters--outside-trend">
        {showSiteFilter && (
          <label className="chd-select-wrap">
            <span className="chd-sr-only">Site</span>
            <select
              className="chd-select"
              value={selectedSite}
              disabled={siteSelectDisabled}
              onChange={(e) => setSelectedSite(e.target.value)}
            >
              {(isSiteScopedUser
                ? siteOptions.length > 1
                : isAdmin || siteOptions.length > 1) && (
                <option value="all">
                  {isSiteScopedUser ? 'My Sites' : 'All Sites'}
                </option>
              )}
              {siteOptions.map((opt) => (
                <option key={opt.value} value={opt.value}>
                  {opt.label}
                </option>
              ))}
            </select>
            <ChevronDown size={14} className="chd-select-icon" aria-hidden />
          </label>
        )}
        {!showSiteFilter && isSiteScopedUser && siteOptions[0] && (
          <span className="chd-site-lock" title="Site-scoped login">
            {siteOptions[0].label}
          </span>
        )}
        <label className="chd-select-wrap">
          <span className="chd-sr-only">Period</span>
          <select
            className="chd-select"
            value={period}
            onChange={(e) => setPeriod(e.target.value)}
          >
            {PERIOD_OPTIONS.map((opt) => (
              <option key={opt.value} value={opt.value}>
                {opt.label}
              </option>
            ))}
          </select>
          <ChevronDown size={14} className="chd-select-icon" aria-hidden />
        </label>
      </div>

      <div className="chd-grid chd-grid--health-trend">
        <section className="chd-card chd-card--health">
          <div className="chd-card-head">
            <div className="chd-card-head-left">
              <h3>Compliance Health</h3>
            </div>
          </div>
          <div className="chd-health-body">
            <ComplianceGauge score={score} counts={counts} />
            <ul className="chd-status-list">
              {Object.keys(STATUS_LABELS).map((key) => (
                <li key={key}>
                  <span className="chd-status-left">
                    <span className="chd-dot" style={{ background: STATUS_COLORS[key] }} />
                    {STATUS_LABELS[key]}
                  </span>
                  <span className="chd-status-right">
                    <strong>{counts[key]}</strong>
                    <em>({pct(counts[key], total)}%)</em>
                  </span>
                </li>
              ))}
            </ul>
          </div>
          <div className="chd-health-footer">
            <span>Overall Compliance Score</span>
            <div className="chd-health-score">
              <strong className="chd-score-value">{score}%</strong>
              <small className={delta >= 0 ? 'chd-delta-up' : 'chd-delta-down'}>
                {Math.abs(delta)}% vs last month
              </small>
            </div>
          </div>
        </section>

        <section className="chd-card chd-card--trend">
          <div className="chd-card-head">
            <h3>Compliance Trend (Last 6 Months)</h3>
            <Link to={TRANSACTION_PAGE} className="chd-link">
              View Report
            </Link>
          </div>
          <StatusLegend order={TREND_STATUS_ORDER} labels={STATUS_LABELS} />
          <TrendChart series={trend} />
        </section>
      </div>

      <div
        className={`chd-grid ${
          showComplianceByLocation
            ? 'chd-grid--location-industry'
            : 'chd-grid--industry'
        }`}
      >
        {showComplianceByLocation && (
          <section className="chd-card chd-card--location">
            <div className="chd-card-head">
              <h3>Compliance by Location</h3>
            </div>
            <div className="chd-legend chd-legend--compact">
              {locations.map((loc) => (
                <span key={loc.state} className="chd-legend-item">
                  <span
                    className="chd-dot"
                    style={{ background: loc.color || mapColorForState(loc.state) }}
                  />
                  {loc.state}
                </span>
              ))}
            </div>
            <div className="chd-location-body">
              <div className="chd-map-wrap">
                <IndiaLocationMap locations={locations} />
              </div>
              <LocationBars rows={locations} />
            </div>
          </section>
        )}

        <section
          className={`chd-card chd-card--industry${
            isSiteScopedUser ? ' chd-card--act' : ''
          }`}
        >
          <div className="chd-card-head">
            <h3>{isSiteScopedUser ? 'Compliance by Act' : 'Compliance by Industry'}</h3>
            <Link to={TRANSACTION_PAGE} className="chd-link">
              View All
            </Link>
          </div>
          {isSiteScopedUser ? (
            <>
              <StatusLegend
                order={ACT_CHART_STATUS_ORDER}
                square
                labels={ACT_CHART_LABELS}
              />
              <ActStackedBars series={acts} />
            </>
          ) : (
            <>
              <StatusLegend order={INDUSTRY_STATUS_ORDER} square />
              <IndustryBars series={industries} />
            </>
          )}
        </section>
      </div>
    </div>
  );
}
