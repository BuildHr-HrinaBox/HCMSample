import React, { useEffect, useMemo, useState } from 'react';
import { Link } from 'react-router-dom';
import { ChevronDown } from 'lucide-react';
import { MapContainer, TileLayer, GeoJSON, useMap } from 'react-leaflet';
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
/** Trend legend order: Approved → Pending → Returned → Yet to Submit. */
const TREND_STATUS_ORDER = ['approved', 'pending', 'returned', 'yetToSubmit'];

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

function rowIdentityKey(row, { ignoreSite = false } = {}) {
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
  if (!form && !act && !desc) return '';
  if (ignoreSite) return `${form}|${act}|${desc}`;
  return `${form}|${act}|${desc}|${site}`;
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
  };
}

/** Prefer statutory workflow rows; add unique checklist-bulk lines as yet-to-submit only once. */
function mergeStatutoryWithBulk(statutory, bulk) {
  const byKey = new Map();
  const byFormActDesc = new Map();

  (statutory || []).forEach((row) => {
    const fullKey = rowIdentityKey(row) || `stat-${byKey.size}-${String(row?.id ?? '')}`;
    const next = pickStrongerRow(byKey.get(fullKey), row);
    byKey.set(fullKey, next);

    const softKey = rowIdentityKey(row, { ignoreSite: true });
    if (softKey) {
      byFormActDesc.set(softKey, pickStrongerRow(byFormActDesc.get(softKey), next));
    }
  });

  (bulk || []).forEach((row) => {
    const softKey = rowIdentityKey(row, { ignoreSite: true });
    if (!softKey) return;

    const fullKey = rowIdentityKey(row);
    const exact = fullKey ? byKey.get(fullKey) : null;
    const soft = byFormActDesc.get(softKey);

    if (exact) {
      const merged = overlayMasterOntoWorkflow(row, exact);
      byKey.set(fullKey, merged);
      byFormActDesc.set(softKey, pickStrongerRow(byFormActDesc.get(softKey), merged));
      return;
    }

    if (soft) {
      // Bulk often has no site — attach metadata onto the best matching statutory row.
      const softFullKey =
        rowIdentityKey(soft) || `stat-soft-${String(soft?.id ?? softKey)}`;
      const merged = overlayMasterOntoWorkflow(row, soft);
      byKey.set(softFullKey, merged);
      byFormActDesc.set(softKey, merged);
      return;
    }

    const insertKey = fullKey || `bulk-${byKey.size}-${softKey}`;
    byKey.set(insertKey, row);
    byFormActDesc.set(softKey, row);
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

/**
 * Resolve which trend month a live statutory row belongs to.
 * Prefer MonthFilter (same as Statutory UI), then real due/activity dates.
 */
function resolveTrendMonthKey(item, windowKeys, now = new Date()) {
  const windowSet = new Set(windowKeys);
  const currentKey = monthKey(new Date(now.getFullYear(), now.getMonth(), 1));

  const monthFilter = item?.monthFilter ?? item?.MonthFilter ?? item?.monthfilter;
  if (monthFilter != null && String(monthFilter).trim() !== '') {
    const fromFilter = monthIndexToWindowKey(parseMonthIndex(monthFilter), windowKeys);
    if (fromFilter) return fromFilter;
  }

  const dueRaw = item?.dueDate ?? item?.DueDate;
  if (dueRaw != null && String(dueRaw).trim() !== '' && !/^monthly\s*basis$/i.test(String(dueRaw).trim())) {
    if (!isDayOnlyDueDate(dueRaw)) {
      const dueTs = parseDateValue(dueRaw);
      if (dueTs != null) {
        const d = new Date(dueTs);
        const key = monthKey(new Date(d.getFullYear(), d.getMonth(), 1));
        if (windowSet.has(key)) return key;
      }
    }
    const dueMonthIdx = parseMonthIndex(dueRaw);
    const fromDueName = monthIndexToWindowKey(dueMonthIdx, windowKeys);
    if (fromDueName) return fromDueName;
  }

  const activity = itemActivityTimestamp(item);
  if (activity != null) {
    const d = new Date(activity);
    const key = monthKey(new Date(d.getFullYear(), d.getMonth(), 1));
    if (windowSet.has(key)) return key;
  }

  // Undated / out-of-window live rows → current month (same universe as Compliance Health).
  return windowSet.has(currentKey) ? currentKey : windowKeys[windowKeys.length - 1] || currentKey;
}

function buildTrendSeries(items, now = new Date()) {
  const months = [];
  for (let i = 5; i >= 0; i -= 1) {
    const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
    months.push({ key: monthKey(d), label: monthLabel(d), date: d, ...EMPTY_COUNTS });
  }
  const byKey = Object.fromEntries(months.map((m) => [m.key, m]));
  const windowKeys = months.map((m) => m.key);

  items.forEach((item) => {
    const bucket = getComplianceBucket(item);
    const key = resolveTrendMonthKey(item, windowKeys, now);
    if (byKey[key]) byKey[key][bucket] += 1;
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
    const stateField = row?.states ?? row?.state ?? row?.State ?? '';
    if (String(stateField || '').trim() && !statesFieldMatchesInchargeSiteStates(stateField, scope.stateLabels)) {
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

function itemInPeriod(item, period, now = new Date()) {
  // Live snapshot — same universe as top KPI cards (Approved / Pending / Returned).
  // Period still filters Last Month / This Year for historical views.
  if (period === 'thisMonth' || period === 'last6Months') return true;

  const activity = itemActivityTimestamp(item);
  const y = now.getFullYear();
  const m = now.getMonth();

  if (period === 'thisYear') {
    if (activity == null) return true;
    return new Date(activity).getFullYear() === y;
  }

  if (period === 'lastMonth') {
    const prev = new Date(y, m - 1, 1);
    if (activity == null) return false;
    const d = new Date(activity);
    return d.getFullYear() === prev.getFullYear() && d.getMonth() === prev.getMonth();
  }

  return true;
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
        color: STATUS_COLORS[dominantStatus(row)],
        lat: coords?.[0] ?? null,
        lng: coords?.[1] ?? null,
      };
    })
    .filter((row) => row.total > 0)
    .sort((a, b) => b.total - a.total || a.state.localeCompare(b.state))
    .slice(0, 8);
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

function buildIndustrySeries(items) {
  const byKey = Object.fromEntries(
    INDUSTRY_GROUPS.map((g) => [g.label, { industry: g.label, ...EMPTY_COUNTS, total: 0 }])
  );

  items.forEach((item) => {
    const label = industryGroupLabel(item);
    if (!label || !byKey[label]) return;
    const bucket = getComplianceBucket(item);
    byKey[label][bucket] += 1;
    byKey[label].total += 1;
  });

  return INDUSTRY_GROUPS.map((g) => byKey[g.label]);
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

const INDIA_STATES_GEOJSON_URL = '/data/india-states.geojson';
const MAP_IDLE_FILL = '#e5e7eb';
const MAP_IDLE_STROKE = '#cbd5e1';

/** Normalize GeoJSON / app state labels so Pondicherry↔Puducherry etc. match. */
function geoStateKey(name) {
  const key = normalizeStateCompareKey(name);
  if (key === 'pondicherry') return 'puducherry';
  if (key === 'orissa') return 'odisha';
  if (key === 'nctofdelhi' || key === 'nctdelhi') return 'delhi';
  return key;
}

function MapInvalidateSize() {
  const map = useMap();
  useEffect(() => {
    const t = setTimeout(() => map.invalidateSize(), 80);
    return () => clearTimeout(t);
  }, [map]);
  return null;
}

function IndiaLocationMap({ locations }) {
  const [geoData, setGeoData] = useState(null);

  useEffect(() => {
    let cancelled = false;
    fetch(INDIA_STATES_GEOJSON_URL, { cache: 'force-cache' })
      .then((res) => (res.ok ? res.json() : null))
      .then((data) => {
        if (!cancelled) setGeoData(data);
      })
      .catch(() => {
        if (!cancelled) setGeoData(null);
      });
    return () => {
      cancelled = true;
    };
  }, []);

  const byStateKey = useMemo(() => {
    const map = new Map();
    (locations || []).forEach((loc) => {
      const key = geoStateKey(loc.state);
      if (key) map.set(key, loc);
    });
    return map;
  }, [locations]);

  const geoStyle = useMemo(
    () => (feature) => {
      const loc = byStateKey.get(geoStateKey(feature?.properties?.name));
      if (loc) {
        return {
          fillColor: loc.color || STATUS_COLORS.yetToSubmit,
          fillOpacity: 0.88,
          color: '#ffffff',
          weight: 1.25,
          opacity: 1,
        };
      }
      return {
        fillColor: MAP_IDLE_FILL,
        fillOpacity: 0.75,
        color: MAP_IDLE_STROKE,
        weight: 0.8,
        opacity: 1,
      };
    },
    [byStateKey]
  );

  const onEachFeature = useMemo(
    () => (feature, layer) => {
      const name = feature?.properties?.name || 'State';
      const loc = byStateKey.get(geoStateKey(name));
      if (loc) {
        layer.bindPopup(
          `<strong>${loc.state}</strong><br/>${loc.score}% overall · ${loc.total} compliances`
        );
        layer.on({
          mouseover: (e) => {
            e.target.setStyle({ weight: 2, fillOpacity: 0.98 });
            e.target.bringToFront();
          },
          mouseout: (e) => {
            e.target.setStyle(geoStyle(feature));
          },
        });
      } else {
        layer.bindPopup(`<strong>${name}</strong><br/>No compliance data`);
      }
    },
    [byStateKey, geoStyle]
  );

  const geoKey = useMemo(
    () =>
      (locations || [])
        .map((l) => `${geoStateKey(l.state)}:${l.color}:${l.score}`)
        .join('|'),
    [locations]
  );

  return (
    <MapContainer
      center={[22.5, 82]}
      zoom={4}
      minZoom={4}
      maxZoom={7}
      scrollWheelZoom={false}
      dragging
      zoomControl={false}
      attributionControl={false}
      className="chd-india-map"
      style={{ height: '100%', width: '100%' }}
      maxBounds={L.latLngBounds([6, 68], [37, 98])}
      maxBoundsViscosity={1}
    >
      <MapInvalidateSize />
      <TileLayer
        url="https://{s}.basemaps.cartocdn.com/light_nolabels/{z}/{x}/{y}{r}.png"
        subdomains="abcd"
      />
      {geoData ? (
        <GeoJSON key={geoKey} data={geoData} style={geoStyle} onEachFeature={onEachFeature} />
      ) : null}
    </MapContainer>
  );
}

function LocationBars({ rows }) {
  if (!rows.length) {
    return <div className="chd-empty">No location compliance data yet.</div>;
  }

  return (
    <div className="chd-loc-list">
      {rows.map((row) => (
        <div key={row.state} className="chd-loc-row">
          <span className="chd-loc-name" title={row.state}>
            {row.state}
          </span>
          <span className="chd-loc-score">{row.score}% overall</span>
          <div className="chd-loc-bar-track" aria-hidden>
            {LOCATION_STATUS_ORDER.map((key) => {
              const value = row[key] || 0;
              if (!value || !row.total) return null;
              const width = (value / row.total) * 100;
              return (
                <span
                  key={key}
                  className="chd-loc-bar-seg"
                  style={{ width: `${width}%`, background: STATUS_COLORS[key] }}
                  title={`${STATUS_LABELS[key]}: ${value}`}
                />
              );
            })}
          </div>
        </div>
      ))}
    </div>
  );
}

function IndustryBars({ series }) {
  const width = 420;
  const height = 260;
  const pad = { top: 28, right: 16, bottom: 36, left: 48 };
  const chartW = width - pad.left - pad.right;
  const chartH = height - pad.top - pad.bottom;
  const keys = INDUSTRY_STATUS_ORDER;
  const rawMax = Math.max(
    0,
    ...series.flatMap((row) => keys.map((key) => Number(row[key]) || 0))
  );
  const maxY = Math.max(30, Math.ceil(rawMax / 5) * 5 || 30);
  const ticks = [];
  for (let v = 0; v <= maxY; v += 5) ticks.push(v);

  const groupCount = Math.max(series.length, 1);
  const groupW = chartW / groupCount;
  const barGap = 3;
  const barW = Math.min(14, (groupW - 20) / keys.length - barGap);
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
        x={14}
        y={pad.top + chartH / 2}
        textAnchor="middle"
        fontSize="10"
        fill="#9ca3af"
        transform={`rotate(-90 14 ${pad.top + chartH / 2})`}
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
      {series.map((row, gi) => {
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
              y={height - 10}
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

/** Straight-line mockup model: linear segments between month points (no curves). */
function TrendChart({ series }) {
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
  // Match mockup: 0–100 with steps of 25; scale up in steps of 50 when volume is higher.
  const maxY = Math.max(100, Math.ceil(rawMax / 50) * 50 || 100);
  const ticks = [];
  const step = maxY <= 100 ? 25 : 50;
  for (let v = 0; v <= maxY; v += step) ticks.push(v);

  const xAt = (i) => pad.left + (n === 1 ? chartW / 2 : (i / (n - 1)) * chartW);
  const yAt = (v) => pad.top + chartH - (Math.min(Math.max(v, 0), maxY) / maxY) * chartH;
  const hasData = series.some((s) => (s.total || 0) > 0);

  // Explicit M/L polyline → sharp straight segments (no curve / spline / rounded joins).
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

  return (
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
      {series.map((_, i) => (
        <line
          key={`v-${i}`}
          x1={xAt(i)}
          y1={pad.top}
          x2={xAt(i)}
          y2={pad.top + chartH}
          stroke="#f3f4f6"
          strokeWidth="1"
        />
      ))}
      {paths.map((p) => (
        <path
          key={p.key}
          d={p.d}
          fill="none"
          stroke={STATUS_COLORS[p.key]}
          strokeWidth="2"
          strokeLinejoin="miter"
          strokeLinecap="butt"
          strokeOpacity={hasData ? 1 : 0.35}
        />
      ))}
      {keys.map((key) =>
        series.map((row, i) => {
          const value = Number(row[key]) || 0;
          return (
            <circle
              key={`${key}-${i}`}
              cx={xAt(i)}
              cy={yAt(value)}
              r="3.5"
              fill={STATUS_COLORS[key]}
              stroke="#fff"
              strokeWidth="1.5"
              opacity={hasData ? 1 : 0.35}
            />
          );
        })
      )}
      {series.map((row, i) => (
        <text
          key={row.label}
          x={xAt(i)}
          y={height - 8}
          textAnchor="middle"
          fontSize="11"
          fill="#6b7280"
        >
          {row.label}
        </text>
      ))}
    </svg>
  );
}

export default function ComplianceHealthDashboard({ userRole, userEmail }) {
  const [loading, setLoading] = useState(true);
  const [period, setPeriod] = useState('thisMonth');
  const [selectedSite, setSelectedSite] = useState('all');
  const [siteOptions, setSiteOptions] = useState([]);
  const [isSiteScopedUser, setIsSiteScopedUser] = useState(false);
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
        const scopedUser = Boolean(hasSiteScope || inchargeScope?.mine?.length);
        const adminViewer = isOrgWideSiteViewer(userRole);

        const options = (scopedUser && !adminViewer ? visibleSites : siteDetails)
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
        // Site login: only their Site Management assignment (industry/state/site).
        // Admin: full org dataset (optional All Sites dropdown filters further).
        const scoped =
          hasSiteScope && !adminViewer
            ? merged.filter((item) => statutoryRowMatchesSiteScope(item, siteScope))
            : merged;

        if (!cancelled) {
          setIsSiteScopedUser(scopedUser && !adminViewer);
          setSiteOptions(options);
          setSiteRecords(scopedUser && !adminViewer ? visibleSites : siteDetails);
          if (scopedUser && !adminViewer && options.length === 1) {
            setSelectedSite(options[0].value);
          }
          setRawItems(scoped);
          setLoading(false);
        }
      } catch {
        if (!cancelled) {
          setRawItems([]);
          setSiteRecords([]);
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

  const items = useMemo(() => {
    return rawItems
      .filter((item) => rowMatchesSelectedSite(item, selectedSite, siteRecords))
      .filter((item) => itemInPeriod(item, period));
  }, [rawItems, selectedSite, period, siteRecords]);

  const lastMonthItems = useMemo(() => {
    return rawItems
      .filter((item) => rowMatchesSelectedSite(item, selectedSite, siteRecords))
      .filter((item) => itemInPeriod(item, 'lastMonth'));
  }, [rawItems, selectedSite, siteRecords]);

  const counts = useMemo(() => countBuckets(items), [items]);
  const total =
    counts.approved + counts.pending + counts.yetToSubmit + counts.returned;
  const score = scoreFromCounts(counts);
  const lastMonthScore = scoreFromCounts(countBuckets(lastMonthItems));
  const delta = score - lastMonthScore;

  const trend = useMemo(() => {
    const siteScoped = rawItems.filter((item) =>
      rowMatchesSelectedSite(item, selectedSite, siteRecords)
    );
    return buildTrendSeries(siteScoped);
  }, [rawItems, selectedSite, siteRecords]);

  const locations = useMemo(
    () => buildLocationSeries(items, siteRecords),
    [items, siteRecords]
  );
  const industries = useMemo(() => buildIndustrySeries(items), [items]);

  const showSiteFilter = isAdmin || siteOptions.length > 1;
  const siteSelectDisabled = isSiteScopedUser && siteOptions.length <= 1;

  return (
    <div className="chd-root" aria-busy={loading || undefined}>
      <div className="chd-grid chd-grid--health-trend">
        <section className="chd-card chd-card--health">
          <div className="chd-card-head">
            <div className="chd-card-head-left">
              <h3>Compliance Health</h3>
            </div>
            <div className="chd-filters">
              {showSiteFilter && (
                <label className="chd-select-wrap">
                  <span className="chd-sr-only">Site</span>
                  <select
                    className="chd-select"
                    value={selectedSite}
                    disabled={siteSelectDisabled}
                    onChange={(e) => setSelectedSite(e.target.value)}
                  >
                    {(isAdmin || siteOptions.length > 1) && (
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

      <div className="chd-grid chd-grid--location-industry">
        <section className="chd-card chd-card--location">
          <div className="chd-card-head">
            <h3>Compliance by Location</h3>
          </div>
          <div className="chd-location-body">
            <div className="chd-map-wrap">
              <IndiaLocationMap locations={locations} />
            </div>
            <LocationBars rows={locations} />
          </div>
        </section>

        <section className="chd-card chd-card--industry">
          <div className="chd-card-head">
            <h3>Compliance by Industry</h3>
            <Link to={TRANSACTION_PAGE} className="chd-link">
              View All
            </Link>
          </div>
          <StatusLegend order={INDUSTRY_STATUS_ORDER} square />
          <IndustryBars series={industries} />
        </section>
      </div>
    </div>
  );
}
