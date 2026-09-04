'use strict';
const express = { json() { return function noopJson() {}; } };
function createStubApp() {
  const app = { use() { return app; }, get() { return app; }, post() { return app; }, put() { return app; }, delete() { return app; } };
  return app;
}
const app = createStubApp();
const archiver = () => ({ on() { return this; }, pipe() { return this; }, append() { return this; }, finalize() { return Promise.resolve(); } });
const catalystSDK = { initialize() { return {}; } };


/** Same Draft folder as statutoryreg_function (DraftFileName) */
const DRAFT_FOLDER_ID = '26741000000061720';

const MONTH_ORDER = [
  'January',
  'February',
  'March',
  'April',
  'May',
  'June',
  'July',
  'August',
  'September',
  'October',
  'November',
  'December'
];

const SHORT_MONTH = {
  jan: 0,
  feb: 1,
  mar: 2,
  apr: 3,
  may: 4,
  jun: 5,
  jul: 6,
  aug: 7,
  sep: 8,
  oct: 9,
  nov: 10,
  dec: 11
};

app.use((req, res, next) => {
  try {
    const catalyst = catalystSDK.initialize(req);
    res.locals.catalyst = catalyst;
    next();
  } catch (err) {
    res.status(500).json({ status: 'failure', message: 'Catalyst init failed' });
  }
});

app.get('/', (req, res) => {
  res.status(200).json({ status: 'success', message: 'mainreport_function ready' });
});

function normalizeMonth(m) {
  const s = String(m || '')
    .trim()
    .toLowerCase();
  if (!s) return '';
  const full = MONTH_ORDER.find((x) => x.toLowerCase() === s);
  if (full) return full;
  if (SHORT_MONTH[s] !== undefined) return MONTH_ORDER[SHORT_MONTH[s]];
  if (/^\d+$/.test(s)) {
    const n = parseInt(s, 10);
    if (n >= 1 && n <= 12) return MONTH_ORDER[n - 1];
  }
  return String(m).trim();
}

function monthsMatch(a, b) {
  return normalizeMonth(a) === normalizeMonth(b);
}

function yearMatch(yr, yParam) {
  const a = parseInt(String(yr || '').trim(), 10);
  const b = parseInt(String(yParam || '').trim(), 10);
  if (isNaN(a) || isNaN(b)) return String(yr || '').trim() === String(yParam || '').trim();
  return a === b;
}

function sortMonths(months) {
  const uniq = [...new Set(months.map((m) => normalizeMonth(m) || m))].filter(Boolean);
  return uniq.sort((a, b) => {
    const ia = MONTH_ORDER.indexOf(normalizeMonth(a) || a);
    const ib = MONTH_ORDER.indexOf(normalizeMonth(b) || b);
    const va = ia === -1 ? 100 : ia;
    const vb = ib === -1 ? 100 : ib;
    return va - vb;
  });
}

function normalizeEmail(value) {
  return String(value || '')
    .trim()
    .toLowerCase();
}

function getActCategoryFromStatutoryRow(row) {
  if (!row) return null;
  const normalizeString = (str) => {
    if (!str) return '';
    return String(str)
      .trim()
      .toLowerCase();
  };
  const act = normalizeString(row.Act || row.act || '');
  const sector = normalizeString(row.Sector || row.sector || '');
  const isFactoriesAct =
    act.includes('factories act') ||
    act.includes('factory act') ||
    act.includes('factories') ||
    act.includes('factory') ||
    sector === 'factories act' ||
    sector === 'factory act' ||
    sector.includes('factories act') ||
    sector.includes('factory act') ||
    sector.includes('factories') ||
    sector.includes('factory');
  if (isFactoriesAct) return 'factories';
  const isShopsAndEstablishment =
    act.includes('shops and establishments') ||
    act.includes('shops and establishment') ||
    act.includes('shop and establishment') ||
    act.includes('the shops and establishments act') ||
    act.includes('the shops and establishment act') ||
    sector === 'shops and establishment' ||
    sector === 'shops and establishments' ||
    sector === 'shops and establishment act' ||
    sector === 'shop and establishment' ||
    sector.includes('shops and establishment') ||
    sector.includes('shop and establishment');
  if (isShopsAndEstablishment) return 'shops_and_establishment';
  const isCLRA =
    act.includes('clra') ||
    act.includes('contract labour') ||
    act.includes('contract labor') ||
    sector === 'clra' ||
    sector.includes('clra') ||
    sector.includes('contract labour') ||
    sector.includes('contract labor');
  if (isCLRA) return 'clra';
  return 'other';
}

function getActCategoryFromIndustryLabel(industry) {
  return getActCategoryFromStatutoryRow({ Sector: industry || '' });
}

/** When one SiteName has multiple Site rows (CLRA + Shops), keep each row's forms separate. */
function reportRowMatchesSiteRowIndustry(row, siteRow, siteName) {
  const wantCat = getActCategoryForSiteRow(siteRow, siteName);
  if (!wantCat || wantCat === 'other') return true;
  const formCat = getActCategoryFromStatutoryRow(row);
  if (formCat && formCat !== 'other') return formCat === wantCat;
  const fromSector = getActCategoryFromIndustryLabel(getSectorFromAnyRow(row));
  if (fromSector && fromSector !== 'other') return fromSector === wantCat;
  return true;
}

/** Same site-name → act bucket as Statutory.js for Delphi URLs when Industry does not map. */
function getActCategoryFromSiteNameFallback(siteName) {
  const s = String(siteName || '')
    .trim()
    .toLowerCase();
  if (s === 'delphi kakinada') return 'clra';
  if (s === 'delphi') return 'shops_and_establishment';
  if (s === 'delphi oragadam') return 'factories';
  return null;
}

/**
 * Act category for a Site row: Industry first, then known site-name aliases.
 */
function getActCategoryForSiteRow(siteRow, siteNameParam) {
  const industry = siteRow
    ? String(siteRow.Industry || siteRow.industry || siteRow.Act || siteRow.act || '').trim()
    : '';
  let cat = getActCategoryFromIndustryLabel(industry);
  if (cat && cat !== 'other') return cat;
  const fallback = getActCategoryFromSiteNameFallback(siteNameParam);
  return fallback || cat || 'other';
}

function normStateToken(v) {
  return String(v || '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ');
}

function normalizeStateCompareKey(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '');
}

/** Match SiteState to checklist State (Andhra Pradesh ↔ AP, Tamil Nadu ↔ tamilnadu, etc.). */
function statesMatchLoose(a, b) {
  const ka = normalizeStateCompareKey(a);
  const kb = normalizeStateCompareKey(b);
  if (!ka || !kb) return false;
  if (ka === kb) return true;
  if (ka.length >= 3 && kb.length >= 3 && (ka.includes(kb) || kb.includes(ka))) return true;
  if ((ka === 'ap' && kb === 'andhrapradesh') || (kb === 'ap' && ka === 'andhrapradesh')) return true;
  if ((ka === 'mp' && kb === 'madhyapradesh') || (kb === 'mp' && ka === 'madhyapradesh')) return true;
  if ((ka === 'tn' && kb === 'tamilnadu') || (kb === 'tn' && ka === 'tamilnadu')) return true;
  return false;
}

/** Same rules as app `statesFieldMatchesInchargeSiteStates` (strict state when Incharge sites have SiteState). */
function rowStatesMatchInchargeScope(rowStateField, stateLabels) {
  if (!stateLabels || stateLabels.length === 0) return true;
  const raw = String(rowStateField ?? '').trim();
  if (!raw) return false;
  const blob = normStateToken(raw);
  if (/\b(all india|pan india|pan-india|national|central|all states|all state)\b/.test(blob)) {
    return false;
  }
  const tokens = String(raw)
    .split(/[,;/|]/)
    .map((t) => normStateToken(t))
    .filter((t) => t.length > 0);
  if (tokens.length === 0) return false;
  const allowed = stateLabels.map(normStateToken).filter(Boolean);
  return tokens.some((tok) => allowed.some((a) => statesMatchLoose(tok, a) || a === tok || tok.includes(a) || a.includes(tok)));
}

function filterRowsForInchargeLocation(rows, stateLabels, siteNamesLower) {
  if ((!stateLabels || stateLabels.length === 0) && (!siteNamesLower || siteNamesLower.size === 0)) {
    return rows;
  }
  if (!Array.isArray(rows)) return rows;
  return rows.filter((r) => {
    const site = String(r.Site || r.site || '').trim().toLowerCase();
    if (siteNamesLower && siteNamesLower.size > 0 && site && siteNamesLower.has(site)) {
      return true;
    }
    if (siteNamesLower && siteNamesLower.size > 0 && site && !siteNamesLower.has(site)) return false;
    const stateRaw = getStateFromAnyRow(r);
    /** ChecklistBulk master rows often have no Site; resolve State later from bulk/sector/site scope. */
    if (!site && !stateRaw) return true;
    return rowStatesMatchInchargeScope(stateRaw, stateLabels);
  });
}

/**
 * Site rows where Incharge email matches user: act categories, state labels, site names (lowercase).
 */
async function getInchargeReportScope(catalyst, userEmail) {
  const email = normalizeEmail(userEmail);
  const empty = { actCategories: null, stateLabels: null, siteNamesLower: null };
  if (!email) return empty;
  try {
    const siteRows = await catalyst.datastore().table('Site').getAllRows();
    const mine = siteRows.filter(
      (row) => normalizeEmail(row.InchargeEmail || row.inchargeEmail || row.incharge_email) === email
    );
    if (mine.length === 0) return empty;
    const cats = new Set();
    const states = new Set();
    const sites = new Set();
    mine.forEach((row) => {
      const sn = String(row.SiteName || row.siteName || '').trim();
      if (sn) sites.add(sn.toLowerCase());
      const st = String(row.SiteState || row.siteState || row.State || row.state || '').trim();
      if (st) states.add(st);
      const cat = getActCategoryFromIndustryLabel(row.Industry || row.industry || '');
      if (cat && cat !== 'other') cats.add(cat);
    });
    return {
      actCategories: cats.size ? [...cats] : null,
      stateLabels: states.size ? [...states] : null,
      siteNamesLower: sites.size ? sites : null
    };
  } catch (err) {
    console.warn('getInchargeReportScope:', err?.message || err);
    return empty;
  }
}

async function findSiteRowByName(catalyst, siteName) {
  const rows = await findAllSiteRowsByName(catalyst, siteName);
  return rows.length ? rows[0] : null;
}

async function findAllSiteRowsByName(catalyst, siteName) {
  const want = String(siteName || '')
    .trim()
    .toLowerCase();
  if (!want) return [];
  try {
    const rows = await fetchAllSiteTableRows(catalyst);
    return rows.filter(
      (r) =>
        String(r?.SiteName || r?.siteName || '')
          .trim()
          .toLowerCase() === want
    );
  } catch (err) {
    console.warn('findAllSiteRowsByName:', err?.message || err);
    return [];
  }
}

function collectActCategoriesForSiteRows(siteRows, siteNameParam) {
  const cats = new Set();
  for (const row of siteRows || []) {
    const cat = getActCategoryForSiteRow(row, siteNameParam);
    if (cat && cat !== 'other') cats.add(cat);
  }
  if (cats.size === 0) cats.add('other');
  return [...cats];
}

function filterRowsByActCategoryScope(rows, allowedActCategories) {
  if (!Array.isArray(allowedActCategories) || allowedActCategories.length === 0) return rows;
  const allow = new Set(allowedActCategories);
  return rows.filter((r) => allow.has(getActCategoryFromStatutoryRow(r)));
}

function isPanIndiaOrBlankState(stateRaw) {
  const raw = String(stateRaw ?? '').trim();
  if (!raw) return true;
  const blob = normStateToken(raw);
  return /\b(all india|pan india|pan-india|national|central|all states|all state)\b/.test(blob);
}

/** Checklist bulk row state vs selected site's SiteState (pan-India / blank master rows still match). */
function rowStateMatchesSiteState(row, siteState) {
  if (!siteState) return true;
  const stateRaw = getStateFromAnyRow(row);
  if (isPanIndiaOrBlankState(stateRaw)) return true;
  return rowStatesMatchInchargeScope(stateRaw, [siteState]);
}

function filterBulkRowsForSiteScope(bulkRows, siteRow, allowedActCategories, siteParam = '') {
  let scoped = filterRowsByActCategoryScope(bulkRows, allowedActCategories);
  const siteState = getSiteStateFromSiteRow(siteRow);
  if (siteState) {
    scoped = scoped.filter((r) => rowStateMatchesSiteState(r, siteState));
  }
  if (siteParam) {
    const want = String(siteParam).trim().toLowerCase();
    scoped = scoped.filter((r) => {
      const rowSite = String(r.Site || r.site || '')
        .trim()
        .toLowerCase();
      return !rowSite || rowSite === want;
    });
  }
  return scoped;
}

/** Bulk master rows for report — site filter, else incharge state scope, else all. */
function filterBulkForReportScope(
  bulkAll,
  { siteParam, siteRow, narrowActCategories, inchargeStateLabels }
) {
  if (siteParam && siteRow) {
    return filterBulkRowsForSiteScope(bulkAll, siteRow, narrowActCategories, siteParam);
  }
  let scoped = filterRowsByActCategoryScope(bulkAll, narrowActCategories);
  if (Array.isArray(inchargeStateLabels) && inchargeStateLabels.length > 0) {
    scoped = scoped.filter((r) => {
      const stateRaw = getStateFromAnyRow(r);
      if (!stateRaw) return false;
      if (isPanIndiaOrBlankState(stateRaw)) return true;
      return rowStatesMatchInchargeScope(stateRaw, inchargeStateLabels);
    });
  }
  return scoped;
}

function reportRowMatchesSelectedSiteScope(row, siteState, siteParam) {
  if (!siteParam || !siteState) return true;
  const rowState = String(row?.state || '').trim();
  if (!rowState || isPanIndiaOrBlankState(rowState)) return true;
  return rowStatesMatchInchargeScope(rowState, [siteState]);
}

function finalizeReportRowForSite(row, siteParam, siteState) {
  if (!siteParam) return row;
  const next = { ...row, site: siteParam };
  if (siteState) {
    const rowState = String(next.state || '').trim();
    if (!rowState || isPanIndiaOrBlankState(rowState)) {
      next.state = siteState;
    }
  }
  return next;
}

function filterStatutoryRowsForSite(rows, siteParam) {
  if (!siteParam) return rows;
  const want = String(siteParam).trim().toLowerCase();
  return rows.filter((r) => String(r.Site || r.site || '').trim().toLowerCase() === want);
}

/** Month filter as stored on Statutory / Catalyst rows (field names vary by table mapping). */
function getStatutoryStoredMonthRaw(stRow) {
  if (!stRow) return '';
  const v =
    stRow.MonthFilter ??
    stRow.monthfilter ??
    stRow.Monthfilter ??
    stRow.monthFilter ??
    stRow.MONTHFILTER ??
    stRow.OriginalMonthFilter ??
    stRow.originalMonthFilter ??
    '';
  return String(v ?? '').trim();
}

/** Same field as Statutory "Month Filter" column (`monthFilter` / `MonthFilter`). */
function getStatutoryMonthFilterDisplay(stRow, reportMonthFallback) {
  const raw = getStatutoryStoredMonthRaw(stRow);
  if (raw !== '') {
    const n = normalizeMonth(raw);
    return n || raw;
  }
  const fb = normalizeMonth(reportMonthFallback);
  return fb || String(reportMonthFallback || '').trim();
}

function statutoryRowMatchesMonth(stRow, selectedMonth) {
  const selectedNorm = normalizeMonth(selectedMonth);
  if (!selectedNorm) return false;
  const selectedPrefix = selectedNorm.toLowerCase().substring(0, 3);
  const storedNorm = getStatutoryStoredMonthRaw(stRow).toLowerCase().substring(0, 3);
  if (storedNorm) return storedNorm === selectedPrefix;
  const submitted = getDraftDateFromStatutoryRow(stRow);
  if (submitted && submittedDateMatchesMonth(submitted, selectedMonth)) return true;
  const dueDate = String(stRow.DueDate || stRow.dueDate || '')
    .toLowerCase()
    .trim();
  if (!dueDate) return false;
  const monthNames = MONTH_ORDER.map((m) => m.toLowerCase());
  const monthAbbr = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];
  const idx = MONTH_ORDER.indexOf(selectedNorm);
  if (idx === -1) return false;
  const fullMonthName = monthNames[idx];
  const monthAbbrName = monthAbbr[idx];
  /** Never treat "Monthly Basis" alone as matching every month — that showed May-only drafts under January Reports. */
  return dueDate.includes(fullMonthName) || dueDate.includes(monthAbbrName);
}

function submittedDateMatchesMonth(submittedRaw, selectedMonth) {
  const raw = String(submittedRaw || '').trim();
  if (!raw) return false;
  const selectedNorm = normalizeMonth(selectedMonth);
  if (!selectedNorm) return false;
  const selectedPrefix = selectedNorm.toLowerCase().substring(0, 3);
  const lower = raw.toLowerCase();
  if (lower.includes(selectedPrefix)) return true;
  const idx = MONTH_ORDER.indexOf(selectedNorm);
  if (idx === -1) return false;
  return lower.includes(MONTH_ORDER[idx].toLowerCase());
}

function statutoryRowMatchesYear(stRow, selectedYear) {
  const y = parseInt(String(selectedYear || '').trim(), 10);
  if (isNaN(y)) return true;
  const dateFields = [
    getDraftDateFromStatutoryRow(stRow),
    getApprovalDateFromStatutoryRow(stRow),
    String(stRow.DueDate || stRow.dueDate || '').trim()
  ];
  let sawYear = false;
  for (const raw of dateFields) {
    if (!raw) continue;
    const m = String(raw).match(/\b(20\d{2})\b/);
    if (m) {
      sawYear = true;
      if (parseInt(m[1], 10) === y) return true;
    }
  }
  return !sawYear;
}

/** Same as Statutory.js baseFormNameKey — "Form B - TN LWF..." → "form b". */
function baseFormNameKey(name) {
  const n = String(name || '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ');
  const m = n.match(/^form\s+[a-z0-9]+/);
  return m ? m[0] : n;
}

function hasStatutoryDraftFromRow(stRow) {
  if (!stRow) return false;
  const draft = stRow.DraftFile ?? stRow.draftFile ?? stRow.draftfile;
  return draft != null && String(draft).trim() !== '' && String(draft).toLowerCase() !== 'null';
}

function formKeysMatchForReport(meta, stRow, strictFormKey) {
  if (getMasterKeyFromAnyRow(stRow) === strictFormKey) return true;
  const bulkName = getFormNameFromAnyRow(meta);
  const stName = getFormNameFromAnyRow(stRow);
  if (normalizeFormKey(bulkName) === normalizeFormKey(stName)) return true;
  const baseBulk = baseFormNameKey(bulkName);
  const baseSt = baseFormNameKey(stName);
  return Boolean(baseBulk && baseBulk === baseSt);
}

/** Link checklist bulk row → Statutory transaction (loose form key + site + month/year). */
function findStatutoryMatchForReport(meta, formKey, statutoryRows, month, year, siteParam) {
  if (!Array.isArray(statutoryRows) || statutoryRows.length === 0) return null;
  const siteLower = String(siteParam || '').trim().toLowerCase();

  let candidates = statutoryRows.filter((r) => formKeysMatchForReport(meta, r, formKey));

  if (siteLower) {
    candidates = candidates.filter(
      (r) =>
        String(r.Site || r.site || '')
          .trim()
          .toLowerCase() === siteLower
    );
  }

  if (!candidates.length) return null;

  const monthYearMatches = candidates.filter(
    (r) => statutoryRowMatchesMonth(r, month) && statutoryRowMatchesYear(r, year)
  );
  if (monthYearMatches.length) return pickBestStatutoryRow(monthYearMatches);

  // Draft/submitted in this month when MonthFilter is blank (still this period only).
  const withDraftInPeriod = candidates.filter((r) => {
    if (!hasStatutoryDraftFromRow(r)) return false;
    const sub = getDraftDateFromStatutoryRow(r);
    if (sub && submittedDateMatchesMonth(sub, month) && statutoryRowMatchesYear(r, year)) return true;
    const appr = getApprovalDateFromStatutoryRow(r);
    return appr && submittedDateMatchesMonth(appr, month) && statutoryRowMatchesYear(r, year);
  });
  if (withDraftInPeriod.length) return pickBestStatutoryRow(withDraftInPeriod);

  // No statutoryreg row for this month → checklistbulk-only (Yet to Submit). Do not reuse prior months.
  return null;
}

function getStatutoryRowStatusRaw(stRow) {
  if (!stRow) return '';
  const direct =
    stRow.Status ??
    stRow.status ??
    stRow.STATUS ??
    stRow.TransactionStatus ??
    stRow.transactionStatus;
  if (direct != null && String(direct).trim() !== '') return String(direct).trim();
  for (const k of Object.keys(stRow)) {
    if (String(k).toLowerCase() === 'status') {
      const v = stRow[k];
      if (v != null && String(v).trim() !== '') return String(v).trim();
    }
  }
  return '';
}

/** Match Statutory.js STATUS column: transaction Status or Approval field Approved. */
function isStatutoryTransactionApproved(stRow) {
  const st = getStatutoryRowStatusRaw(stRow).toLowerCase();
  if (st === 'approved' || st === 'approve') return true;
  const appr = String(stRow.Approval || stRow.approval || '')
    .trim()
    .toLowerCase();
  return appr === 'approved' || appr === 'approve';
}

function statutorySendForApprovalIsSent(stRow) {
  if (!stRow) return false;
  return /^sent$/i.test(String(stRow.SendForApproval ?? stRow.sendForApproval ?? '').trim());
}

function getStatutoryRowTimestamp(stRow) {
  if (!stRow) return '';
  const mt = stRow.MODIFIEDTIME ?? stRow.modifiedTime ?? stRow.CREATEDTIME ?? stRow.createdTime;
  return mt != null && String(mt).trim() !== '' ? String(mt).trim() : '';
}

function firstNonEmptyStatutoryDate(...candidates) {
  for (const c of candidates) {
    if (c != null && String(c).trim() !== '' && String(c).trim().toLowerCase() !== 'null') {
      return String(c).trim();
    }
  }
  return '';
}

/** Match Statutory.js `resolveStatutorySubmittedDateStored` + draft/modified-time fallback. */
function getDraftDateFromStatutoryRow(stRow) {
  if (!stRow) return '';
  const direct = firstNonEmptyStatutoryDate(
    stRow.SubmittedDate,
    stRow.submittedDate,
    stRow.DraftDate,
    stRow.draftDate
  );
  if (direct) return direct;
  if (statutorySendForApprovalIsSent(stRow) || hasStatutoryDraftFromRow(stRow)) {
    return getStatutoryRowTimestamp(stRow);
  }
  return '';
}

/** Match Statutory.js `resolveStatutoryApprovedDateStored` with modified-time fallback when approved. */
function getApprovalDateFromStatutoryRow(stRow) {
  if (!stRow) return '';
  const direct = firstNonEmptyStatutoryDate(
    stRow.ApprovedDate,
    stRow.approvedDate,
    stRow.ApprovalDate,
    stRow.approvalDate
  );
  if (direct) return direct;
  if (isStatutoryTransactionApproved(stRow)) {
    return getStatutoryRowTimestamp(stRow);
  }
  return '';
}

function pickBestStatutoryRow(matches) {
  if (!matches || matches.length === 0) return null;
  if (matches.length === 1) return matches[0];
  const score = (r) => {
    let s = 0;
    if (isStatutoryTransactionApproved(r)) s += 1000;
    if (hasStatutoryDraftFromRow(r)) s += 100;
    if (r.ROWID != null) s += 50;
    const sub = getDraftDateFromStatutoryRow(r);
    if (sub) s += 10;
    const appr = getApprovalDateFromStatutoryRow(r);
    if (appr) s += 10;
    return s;
  };
  return [...matches].sort((a, b) => score(b) - score(a))[0];
}

function findStatutoryCandidatesForMainRow(mainRow, filteredStatutoryRows) {
  const raw = String(mainRow.FormName != null ? mainRow.FormName : '').trim();
  if (!raw) return [];
  if (/^\d+$/.test(raw)) {
    const byId = filteredStatutoryRows.filter((x) => String(x.ROWID) === raw);
    if (byId.length) return byId;
  }
  const lower = raw.toLowerCase();
  return filteredStatutoryRows.filter((x) => String(x.FormName || '').trim().toLowerCase() === lower);
}

function buildStatutoryDraftRow(stRow) {
  const id = stRow.ROWID;
  const formName = String(stRow.FormName || '').trim();
  const act = getActFromAnyRow(stRow);
  const description = getDescriptionFromAnyRow(stRow);
  const sector = getSectorFromAnyRow(stRow);
  const approved = isStatutoryTransactionApproved(stRow);
  const hasDraftStored = hasStatutoryDraftFromRow(stRow);
  const draftFile = hasDraftStored ? String(stRow.DraftFile).trim() : null;
  const draftFileName =
    stRow.DraftFileName != null && String(stRow.DraftFileName).trim() !== ''
      ? String(stRow.DraftFileName).trim()
      : 'draft';
  /** Reports: draft download only after transaction approval (same gate as statutory workflow). */
  const draftFileUrl =
    draftFile && approved
      ? `/server/statutoryreg_function/statutory/${id}/file/Draft?name=${encodeURIComponent(draftFileName)}`
      : null;
  const approval = String(stRow.Approval || stRow.approval || '').trim();
  const statutoryStatus = getStatutoryRowStatusRaw(stRow);
  const monthFilter = getStatutoryMonthFilterDisplay(stRow, '');
  return {
    rowId: id,
    statutoryRowId: id,
    formName,
    act,
    description,
    sector,
    site: String(stRow.Site || stRow.site || '').trim(),
    state: getStateFromAnyRow(stRow),
    monthFilter,
    draftFile: approved ? draftFile : null,
    draftFileId: hasDraftStored ? draftFile : null,
    draftFileName: hasDraftStored ? draftFileName : '',
    draftFileUrl,
    approval,
    statutoryStatus,
    hasStatutoryDraftStored: hasDraftStored,
    submittedDate: getDraftDateFromStatutoryRow(stRow),
    draftDate: getDraftDateFromStatutoryRow(stRow),
    approvedDate: getApprovalDateFromStatutoryRow(stRow),
    approvalDate: getApprovalDateFromStatutoryRow(stRow),
    sendForApproval: String(stRow.SendForApproval || stRow.sendForApproval || '').trim(),
    modifiedTime: getStatutoryRowTimestamp(stRow)
  };
}

function normalizeFormKey(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]/g, '');
}

function getFormNameFromAnyRow(row) {
  return String(row?.FormName || row?.formName || '').trim();
}

function getActFromAnyRow(row) {
  return String(row?.Act || row?.act || '').trim();
}

function getDescriptionFromAnyRow(row) {
  return String(row?.Description || row?.description || '').trim();
}

function getSectorFromAnyRow(row) {
  return String(row?.Sector || row?.sector || '').trim();
}

function getStateFromAnyRow(row) {
  if (!row) return '';
  const direct = String(row.State ?? row.state ?? '').trim();
  if (direct) return direct;
  for (const k of Object.keys(row)) {
    if (String(k).toLowerCase() === 'state') {
      const v = row[k];
      if (v != null && String(v).trim() !== '') return String(v).trim();
    }
  }
  return '';
}

function unwrapCatalystTableRow(row, tableName) {
  if (!row || typeof row !== 'object') return row;
  const nested = row[tableName];
  if (nested && typeof nested === 'object') {
    return { ...nested, ROWID: nested.ROWID ?? row.ROWID };
  }
  return row;
}

function normalizeChecklistBulkRowForReport(row) {
  const base = unwrapCatalystTableRow(row, 'checklistbulk');
  if (!base || typeof base !== 'object') return base;
  return {
    ...base,
    Sector: base.Sector ?? base.sector ?? '',
    State: base.State ?? base.state ?? '',
    Act: base.Act ?? base.act ?? '',
    FormName: base.FormName ?? base.formName ?? '',
    Description: base.Description ?? base.description ?? '',
    Site: base.Site ?? base.site ?? ''
  };
}

function upsertFormMeta(formsMap, key, row) {
  const name = getFormNameFromAnyRow(row);
  const state = getStateFromAnyRow(row);
  const sector = getSectorFromAnyRow(row);
  if (!formsMap.has(key)) {
    formsMap.set(key, {
      formName: name,
      act: getActFromAnyRow(row),
      description: getDescriptionFromAnyRow(row),
      sector,
      state
    });
    return;
  }
  const existing = formsMap.get(key);
  if (!existing.state && state) existing.state = state;
  if (!existing.sector && sector) existing.sector = sector;
  if (!existing.act) existing.act = getActFromAnyRow(row);
  if (!existing.description) existing.description = getDescriptionFromAnyRow(row);
}

function findStateForMasterKey(formKey, bulkRows, statutoryRows) {
  for (const r of bulkRows) {
    if (getMasterKeyFromAnyRow(r) === formKey) {
      const st = getStateFromAnyRow(r);
      if (st) return st;
    }
  }
  for (const r of statutoryRows) {
    if (getMasterKeyFromAnyRow(r) === formKey) {
      const st = getStateFromAnyRow(r);
      if (st) return st;
    }
  }
  return '';
}

function findStateByFormNameLoose(formName, bulkRows, statutoryRows) {
  const want = normalizeFormKey(formName);
  if (!want) return '';
  for (const r of bulkRows) {
    if (normalizeFormKey(getFormNameFromAnyRow(r)) === want) {
      const st = getStateFromAnyRow(r);
      if (st) return st;
    }
  }
  for (const r of statutoryRows) {
    if (normalizeFormKey(getFormNameFromAnyRow(r)) === want) {
      const st = getStateFromAnyRow(r);
      if (st) return st;
    }
  }
  return '';
}

function findSectorForMasterKey(formKey, bulkRows, statutoryRows) {
  for (const r of bulkRows) {
    if (getMasterKeyFromAnyRow(r) === formKey) {
      const sec = getSectorFromAnyRow(r);
      if (sec) return sec;
    }
  }
  for (const r of statutoryRows) {
    if (getMasterKeyFromAnyRow(r) === formKey) {
      const sec = getSectorFromAnyRow(r);
      if (sec) return sec;
    }
  }
  return '';
}

function findSectorByFormNameLoose(formName, bulkRows, statutoryRows) {
  const want = normalizeFormKey(formName);
  if (!want) return '';
  for (const r of bulkRows) {
    if (normalizeFormKey(getFormNameFromAnyRow(r)) === want) {
      const sec = getSectorFromAnyRow(r);
      if (sec) return sec;
    }
  }
  for (const r of statutoryRows) {
    if (normalizeFormKey(getFormNameFromAnyRow(r)) === want) {
      const sec = getSectorFromAnyRow(r);
      if (sec) return sec;
    }
  }
  return '';
}

/** Pick checklistbulk Sector for an approved site/form (state-aware; avoids first global bulk match). */
function findBulkSectorForSiteForm(bulkRows, reportRow, statutoryRow) {
  const source = statutoryRow || reportRow || {};
  const formKey = getMasterKeyFromAnyRow(source);
  const formName = getFormNameFromAnyRow(source);
  const siteState =
    getStateFromAnyRow(statutoryRow) ||
    String(reportRow?.state || '').trim() ||
    '';
  let candidates = (bulkRows || []).filter((r) => getMasterKeyFromAnyRow(r) === formKey);
  if (!candidates.length && formName) {
    const want = normalizeFormKey(formName);
    candidates = (bulkRows || []).filter((r) => normalizeFormKey(getFormNameFromAnyRow(r)) === want);
  }
  if (siteState && candidates.length > 1) {
    const stateMatches = candidates.filter((r) => rowStateMatchesSiteState(r, siteState));
    if (stateMatches.length) candidates = stateMatches;
  }
  if (candidates.length) {
    return getSectorFromAnyRow(candidates[0]);
  }
  return (
    String(reportRow?.sector || '').trim() ||
    getSectorFromAnyRow(statutoryRow) ||
    ''
  );
}

function findStatutoryMatchForReturned(row, statutoryRows) {
  if (!Array.isArray(statutoryRows) || statutoryRows.length === 0) return null;
  const formKey = getMasterKeyFromAnyRow(row);
  const siteLower = String(row?.Site || row?.site || '')
    .trim()
    .toLowerCase();
  let candidates = statutoryRows.filter((r) => getMasterKeyFromAnyRow(r) === formKey);
  if (!candidates.length) {
    const fn = normalizeFormKey(getFormNameFromAnyRow(row));
    if (fn) {
      candidates = statutoryRows.filter((r) => normalizeFormKey(getFormNameFromAnyRow(r)) === fn);
    }
  }
  if (siteLower) {
    const siteMatches = candidates.filter(
      (r) =>
        String(r.Site || r.site || '')
          .trim()
          .toLowerCase() === siteLower
    );
    if (siteMatches.length) candidates = siteMatches;
  }
  return candidates.length ? pickBestStatutoryRow(candidates) : null;
}

function buildSectorStateMap(bulkRows) {
  const map = new Map();
  for (const row of bulkRows) {
    const sec = getSectorFromAnyRow(row).toLowerCase();
    const st = getStateFromAnyRow(row);
    if (sec && st && !map.has(sec)) map.set(sec, st);
  }
  return map;
}

function normalizeSiteRowForReport(row) {
  const base = unwrapCatalystTableRow(row, 'Site');
  if (!base || typeof base !== 'object') return base;
  return {
    ...base,
    SiteName: base.SiteName ?? base.siteName ?? '',
    SiteState: base.SiteState ?? base.siteState ?? base.State ?? base.state ?? '',
    Industry: base.Industry ?? base.industry ?? base.Act ?? base.act ?? ''
  };
}

function getSiteStateFromSiteRow(siteRow) {
  if (!siteRow) return '';
  return String(siteRow.SiteState || siteRow.siteState || siteRow.State || siteRow.state || '').trim();
}

function getSectorFolderFromActOrIndustry(act, industry, sectorRaw = '') {
  const fromAct = getConsolidatedSectorFolder({ act, sector: sectorRaw });
  if (fromAct && fromAct !== 'Others') return fromAct;
  const cat = getActCategoryFromIndustryLabel(industry);
  if (cat === 'clra') return 'CLRA';
  if (cat === 'factories') return 'Factories Act';
  if (cat === 'shops_and_establishment') return 'Shops and Establishment';
  const sector = String(sectorRaw || '').trim();
  if (sector) return sector;
  const ind = String(industry || '').trim();
  return ind || 'Others';
}

async function buildSiteStateByNameMap(catalyst) {
  const map = new Map();
  try {
    const rows = await fetchAllSiteTableRows(catalyst);
    for (const r of rows) {
      const name = String(r?.SiteName || '')
        .trim()
        .toLowerCase();
      const st = getSiteStateFromSiteRow(r);
      if (name && st) map.set(name, st);
    }
  } catch (err) {
    console.warn('buildSiteStateByNameMap:', err?.message || err);
  }
  return map;
}

async function buildSiteInfoByNameMap(catalyst) {
  const map = new Map();
  try {
    const rows = await fetchAllSiteTableRows(catalyst);
    for (const r of rows) {
      const name = String(r?.SiteName || '')
        .trim()
        .toLowerCase();
      if (!name) continue;
      const state = getSiteStateFromSiteRow(r);
      const industry = String(r?.Industry || '').trim();
      map.set(name, {
        state,
        industry,
        sectorFolder: getSectorFolderFromActOrIndustry('', industry, '')
      });
    }
  } catch (err) {
    console.warn('buildSiteInfoByNameMap:', err?.message || err);
  }
  return map;
}

function resolveReportState(dataRow, metaState, siteStateByName, siteParam, options = {}) {
  const {
    masterState = '',
    sectorStateMap = null,
    sector = '',
    inchargeStateLabels = null
  } = options;
  const direct =
    getStateFromAnyRow(dataRow) ||
    String(metaState || '').trim() ||
    String(masterState || '').trim();
  if (direct) return direct;
  const secKey = String(sector || '').trim().toLowerCase();
  if (secKey && sectorStateMap && sectorStateMap.has(secKey)) {
    return sectorStateMap.get(secKey);
  }
  const siteName = String(dataRow?.Site || dataRow?.site || siteParam || '')
    .trim()
    .toLowerCase();
  if (siteName && siteStateByName.has(siteName)) return siteStateByName.get(siteName);
  if (siteParam) {
    const fromFilter = siteStateByName.get(String(siteParam).trim().toLowerCase());
    if (fromFilter) return fromFilter;
  }
  if (Array.isArray(inchargeStateLabels) && inchargeStateLabels.length === 1) {
    return String(inchargeStateLabels[0] || '').trim();
  }
  return '';
}

function getMasterKeyFromAnyRow(row) {
  const formKey = normalizeFormKey(getFormNameFromAnyRow(row));
  const actKey = normalizeFormKey(getActFromAnyRow(row));
  const descKey = normalizeFormKey(getDescriptionFromAnyRow(row));
  return `${formKey}|${actKey}|${descKey}`;
}

function buildEmptyFormRow(formKey, formName, act, description, sector, reportMonthNorm, state = '') {
  const mf = normalizeMonth(reportMonthNorm) || String(reportMonthNorm || '').trim();
  return {
    rowId: `master_${formKey}`,
    statutoryRowId: null,
    formName: formName || '',
    act: act || '',
    description: description || '',
    sector: sector || '',
    state: String(state || '').trim(),
    monthFilter: mf,
    draftFile: null,
    draftFileName: '',
    draftFileUrl: null,
    approval: '',
    statutoryStatus: '',
    hasStatutoryDraftStored: false,
    submittedDate: '',
    draftDate: '',
    approvedDate: '',
    approvalDate: ''
  };
}

async function fetchSiteManagementTotalCount(catalyst) {
  try {
    const zcql = catalyst.zcql();
    const countRows = await zcql.executeZCQLQuery('SELECT COUNT(ROWID) as count FROM Site');
    const countCell = countRows[0]?.Site || countRows[0]?.site || {};
    const total = parseInt(countCell.count ?? countCell.COUNT ?? countCell['COUNT(ROWID)'] ?? 0, 10);
    return Number.isFinite(total) ? total : 0;
  } catch (err) {
    console.warn('fetchSiteManagementTotalCount failed:', err?.message || err);
    return 0;
  }
}

async function fetchAllSiteTableRows(catalyst) {
  let raw = [];
  const expectedCount = await fetchSiteManagementTotalCount(catalyst);
  try {
    const zcql = catalyst.zcql();
    if (expectedCount > 0 && expectedCount <= 300) {
      const allRows = await zcql.executeZCQLQuery(
        'SELECT ROWID, SiteName, SiteState, Industry FROM Site ORDER BY ROWID DESC'
      );
      if (Array.isArray(allRows) && allRows.length) {
        allRows.forEach((row) => {
          const normalized = normalizeSiteRowForReport(row);
          if (normalized) raw.push(normalized);
        });
      }
    } else {
      const pageSize = 100;
      let offset = 0;
      while (offset < 5000) {
        const chunk = await zcql.executeZCQLQuery(
          `SELECT ROWID, SiteName, SiteState, Industry FROM Site ORDER BY ROWID DESC LIMIT ${offset},${pageSize}`
        );
        if (!Array.isArray(chunk) || chunk.length === 0) break;
        chunk.forEach((row) => {
          const normalized = normalizeSiteRowForReport(row);
          if (normalized) raw.push(normalized);
        });
        if (chunk.length < pageSize) break;
        offset += pageSize;
      }
    }
    console.log(
      `fetchAllSiteTableRows: loaded ${raw.length} site row(s) via ZCQL${expectedCount ? ` (Site Management total: ${expectedCount})` : ''}`
    );
    if (expectedCount && raw.length < expectedCount) {
      console.warn(`fetchAllSiteTableRows: expected ${expectedCount} site row(s), loaded ${raw.length}`);
    }
  } catch (err) {
    console.warn('fetchAllSiteTableRows ZCQL failed:', err?.message || err);
    raw = [];
  }
  if (!raw.length) {
    try {
      const rows = await catalyst.datastore().table('Site').getAllRows();
      raw = (Array.isArray(rows) ? rows : []).map(normalizeSiteRowForReport).filter(Boolean);
      console.log(`fetchAllSiteTableRows: loaded ${raw.length} site row(s) via getAllRows fallback`);
    } catch (err) {
      console.warn('fetchAllSiteTableRows getAllRows failed:', err?.message || err);
      raw = [];
    }
  }
  return raw;
}

function fetchActiveSiteRowsForReminder(catalyst) {
  return fetchAllSiteTableRows(catalyst).then((rows) =>
    rows.filter((row) => String(row.SiteName || row.siteName || '').trim())
  );
}

async function getAllSiteNames(catalyst, allowedActCategories, siteParam, inchargeSiteNamesLower) {
  try {
    const rows = await fetchAllSiteTableRows(catalyst);
    let scopedRows = rows;
    if (Array.isArray(allowedActCategories) && allowedActCategories.length > 0) {
      const allow = new Set(allowedActCategories);
      scopedRows = rows.filter((r) => allow.has(getActCategoryFromIndustryLabel(r.Industry || r.industry || '')));
    }
    if (inchargeSiteNamesLower && inchargeSiteNamesLower.size > 0) {
      scopedRows = scopedRows.filter((r) =>
        inchargeSiteNamesLower.has(String(r.SiteName || r.siteName || '').trim().toLowerCase())
      );
    }
    let names = scopedRows
      .map((r) => String(r.SiteName || r.siteName || '').trim())
      .filter(Boolean);

    if (siteParam) {
      names = names.filter((n) => n.toLowerCase() === String(siteParam).trim().toLowerCase());
    }

    return [...new Set(names)].sort((a, b) => a.localeCompare(b, undefined, { sensitivity: 'base' }));
  } catch (err) {
    console.warn('Could not fetch Site names:', err?.message || err);
    return [];
  }
}

function getContentType(filename) {
  const ext = String(filename || '')
    .toLowerCase()
    .split('.')
    .pop();
  const contentTypes = {
    pdf: 'application/pdf',
    doc: 'application/msword',
    docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    xls: 'application/vnd.ms-excel',
    xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    png: 'image/png',
    jpg: 'image/jpeg',
    jpeg: 'image/jpeg',
    gif: 'image/gif',
    txt: 'text/plain'
  };
  return contentTypes[ext] || 'application/octet-stream';
}

function sanitizeZipPathSegment(value) {
  const s = String(value || '')
    .trim()
    .replace(/[\\/:*?"<>|]/g, '_')
    .replace(/\s+/g, ' ')
    .trim();
  return s.slice(0, 80) || 'Unknown';
}

const CONSOLIDATED_SECTOR_FOLDERS = ['CLRA', 'Factories Act', 'Shops and Establishment', 'Others'];

function getConsolidatedStateFolder(row) {
  const state = String(row?.state || '').trim();
  return sanitizeZipPathSegment(state || 'Unknown');
}

function getConsolidatedSectorFolder(row) {
  const cat = getActCategoryFromStatutoryRow({
    Act: row?.act,
    act: row?.act,
    Sector: row?.sector,
    sector: row?.sector
  });
  if (cat === 'clra') return 'CLRA';
  if (cat === 'factories') return 'Factories Act';
  if (cat === 'shops_and_establishment') return 'Shops and Establishment';
  const s = String(row?.sector || '').trim();
  if (s) return sanitizeZipPathSegment(s);
  return 'Others';
}

function ensureUniqueZipName(baseName, usedNames) {
  let name = String(baseName || 'draft').trim() || 'draft';
  if (!usedNames.has(name)) {
    usedNames.add(name);
    return name;
  }
  const dot = name.lastIndexOf('.');
  const stem = dot > 0 ? name.slice(0, dot) : name;
  const ext = dot > 0 ? name.slice(dot) : '';
  let n = 2;
  while (usedNames.has(`${stem}_${n}${ext}`)) n += 1;
  name = `${stem}_${n}${ext}`;
  usedNames.add(name);
  return name;
}

/** Shared report rows builder for /entries and /consolidated-zip. */
async function buildMainReportEntries(catalyst, { year, month, userEmail, siteParam, siteRow: siteRowHint }) {
  const inchargeScope = await getInchargeReportScope(catalyst, userEmail);
  const allowedActCategories = inchargeScope.actCategories;

  let narrowActCategories = allowedActCategories;
  let siteRow = null;
  if (siteParam) {
    const siteRowsForParam = siteRowHint
      ? [siteRowHint]
      : await findAllSiteRowsByName(catalyst, siteParam);
    if (!siteRowsForParam.length) {
      const siteNames = await getAllSiteNames(catalyst, allowedActCategories, '', inchargeScope.siteNamesLower);
      return { rows: [], siteNames };
    }
    siteRow = siteRowsForParam[0];
    const siteCats = siteRowHint
      ? [getActCategoryForSiteRow(siteRow, siteParam)]
      : collectActCategoriesForSiteRows(siteRowsForParam, siteParam);
    if (Array.isArray(allowedActCategories) && allowedActCategories.length > 0) {
      const intersection = siteCats.filter((c) => allowedActCategories.includes(c));
      if (intersection.length === 0) {
        const siteNames = await getAllSiteNames(catalyst, allowedActCategories, '', inchargeScope.siteNamesLower);
        return { rows: [], siteNames };
      }
      narrowActCategories = intersection;
    } else {
      narrowActCategories = siteCats;
    }
  }

  const statutoryTable = catalyst.datastore().table('Statutory');
  const statutoryAll = await statutoryTable.getAllRows();
  let statutoryForUser = filterRowsByActCategoryScope(statutoryAll, narrowActCategories);
  statutoryForUser = filterRowsForInchargeLocation(
    statutoryForUser,
    inchargeScope.stateLabels,
    inchargeScope.siteNamesLower
  );

  const bulkTable = catalyst.datastore().table('checklistbulk');
  const bulkAllRaw = await bulkTable.getAllRows();
  const bulkAll = (Array.isArray(bulkAllRaw) ? bulkAllRaw : []).map(normalizeChecklistBulkRowForReport);

  const selectedSiteState = siteRow ? getSiteStateFromSiteRow(siteRow) : '';
  const bulkForReport = filterBulkForReportScope(bulkAll, {
    siteParam,
    siteRow,
    narrowActCategories,
    inchargeStateLabels: inchargeScope.stateLabels
  });

  const siteStateByName = await buildSiteStateByNameMap(catalyst);
  const sectorStateMap = buildSectorStateMap(bulkForReport);
  const reportMonthNorm = normalizeMonth(month);
  const resolveOptsBase = {
    siteStateByName,
    siteParam,
    sectorStateMap,
    inchargeStateLabels: inchargeScope.stateLabels
  };

  const statutoryForMatch = (() => {
    let pool = filterRowsByActCategoryScope(statutoryAll, narrowActCategories);
    if (siteParam) {
      const want = String(siteParam).trim().toLowerCase();
      pool = pool.filter((r) => {
        const site = String(r.Site || r.site || '')
          .trim()
          .toLowerCase();
        return site === want;
      });
    }
    return pool;
  })();

  const mergePickedStatutoryIntoReportRow = (picked, meta, listKey, stateOptions) => {
    const built = buildStatutoryDraftRow(picked);
    if (!built.act && meta?.act) built.act = meta.act;
    if (!built.description && meta?.description) built.description = meta.description;
    if (meta?.sector) built.sector = meta.sector;
    built.state = resolveReportState(picked, meta?.state, siteStateByName, siteParam, stateOptions);
    built.monthFilter = getStatutoryMonthFilterDisplay(picked, reportMonthNorm) || reportMonthNorm;
    built.bulkRowId = listKey;
    return finalizeReportRowForSite(built, siteParam, selectedSiteState);
  };

  const pushReportRowForBulkMeta = (listKey, meta, formKey) => {
    const masterState =
      getStateFromAnyRow(meta) ||
      findStateForMasterKey(formKey, bulkForReport, statutoryForMatch) ||
      findStateByFormNameLoose(meta?.formName, bulkForReport, statutoryForMatch) ||
      String(meta?.state || '').trim();
    const picked = findStatutoryMatchForReport(meta, formKey, statutoryForMatch, month, year, siteParam);
    const stateOptions = { ...resolveOptsBase, masterState, sector: meta?.sector || '' };
    if (picked) {
      return mergePickedStatutoryIntoReportRow(picked, meta, listKey, stateOptions);
    }
    const emptyState = resolveReportState(null, meta?.state, siteStateByName, siteParam, stateOptions);
    const empty = buildEmptyFormRow(
      listKey,
      meta?.formName,
      meta?.act,
      meta?.description,
      meta?.sector,
      reportMonthNorm,
      emptyState
    );
    empty.bulkRowId = listKey;
    return finalizeReportRowForSite(empty, siteParam, selectedSiteState);
  };

  const out = [];
  const bulkFormKeysSeen = new Set();

  for (const row of bulkForReport) {
    const name = getFormNameFromAnyRow(row);
    if (!normalizeFormKey(name)) continue;
    const formKey = getMasterKeyFromAnyRow(row);
    bulkFormKeysSeen.add(formKey);
    const listKey = row.ROWID != null ? `bulk_${row.ROWID}` : `bulk_${formKey}_${out.length}`;
    const meta = {
      formName: name,
      act: getActFromAnyRow(row),
      description: getDescriptionFromAnyRow(row),
      sector: getSectorFromAnyRow(row),
      state: getStateFromAnyRow(row)
    };
    out.push(pushReportRowForBulkMeta(listKey, meta, formKey));
  }

  const statutoryForSite = filterStatutoryRowsForSite(statutoryForUser, siteParam);

  for (const row of statutoryForSite) {
    const name = getFormNameFromAnyRow(row);
    if (!normalizeFormKey(name)) continue;
    const formKey = getMasterKeyFromAnyRow(row);
    if (bulkFormKeysSeen.has(formKey)) continue;
    bulkFormKeysSeen.add(formKey);
    const meta = {
      formName: name,
      act: getActFromAnyRow(row),
      description: getDescriptionFromAnyRow(row),
      sector: getSectorFromAnyRow(row),
      state: getStateFromAnyRow(row)
    };
    out.push(pushReportRowForBulkMeta(`statutory_${row.ROWID}`, meta, formKey));
  }

  out.sort((a, b) => String(a.formName || '').localeCompare(String(b.formName || ''), undefined, { sensitivity: 'base' }));

  for (let i = 0; i < out.length; i++) {
    if (out[i].statutoryRowId != null) continue;
    const meta = {
      formName: out[i].formName,
      act: out[i].act,
      description: out[i].description,
      sector: out[i].sector,
      state: out[i].state
    };
    const formKey = getMasterKeyFromAnyRow(meta);
    const masterState =
      getStateFromAnyRow(meta) ||
      findStateForMasterKey(formKey, bulkForReport, statutoryForMatch) ||
      findStateByFormNameLoose(meta?.formName, bulkForReport, statutoryForMatch) ||
      String(meta?.state || '').trim();
    const stateOptions = { ...resolveOptsBase, masterState, sector: meta?.sector || '' };
    const picked = findStatutoryMatchForReport(meta, formKey, statutoryForMatch, month, year, siteParam);
    if (picked) {
      out[i] = mergePickedStatutoryIntoReportRow(
        picked,
        meta,
        out[i].bulkRowId || out[i].rowId,
        stateOptions
      );
    }
  }

  let rows = out;
  if (siteParam && selectedSiteState) {
    rows = out.filter((row) => reportRowMatchesSelectedSiteScope(row, selectedSiteState, siteParam));
  }

  const siteNames = await getAllSiteNames(catalyst, allowedActCategories, '', inchargeScope.siteNamesLower);
  return { rows, siteNames, selectedSiteState: siteParam ? selectedSiteState : '' };
}

/**
 * Years = distinct years from MainReport plus the current calendar year (server time).
 * When the date rolls to 1 Jan of a new year, that year is included automatically.
 */
app.get('/mainreport/summary', async (req, res) => {
  try {
    res.set('Cache-Control', 'no-store');
    const { catalyst } = res.locals;
    const currentYear = new Date().getFullYear();
    const table = catalyst.datastore().table('MainReport');
    const rows = await table.getAllRows();

    const yearNums = new Set();
    yearNums.add(currentYear);

    const monthsByYear = {};

    for (const r of rows) {
      const y = parseInt(String(r.Year || '').trim(), 10);
      if (!isNaN(y)) {
        yearNums.add(y);
        const yk = String(y);
        if (!monthsByYear[yk]) monthsByYear[yk] = [];
        const m = r.Months != null ? String(r.Months).trim() : '';
        if (m) monthsByYear[yk].push(m);
      }
    }

    for (const k of Object.keys(monthsByYear)) {
      monthsByYear[k] = sortMonths(monthsByYear[k]);
    }

    const years = [...yearNums].sort((a, b) => b - a).map(String);

    res.status(200).json({
      status: 'success',
      data: {
        years,
        monthsByYear,
        serverYear: currentYear
      }
    });
  } catch (err) {
    console.error('mainreport/summary:', err);
    res.status(500).json({ status: 'failure', message: err.message || 'Failed to load main report summary.' });
  }
});

app.get('/mainreport/entries', async (req, res) => {
  try {
    res.set('Cache-Control', 'no-store');
    const year = req.query.year;
    const month = req.query.month;
    const userEmail = req.query.userEmail != null ? String(req.query.userEmail) : '';
    if (year == null || year === '' || month == null || month === '') {
      return res.status(400).json({ status: 'failure', message: 'year and month query parameters are required.' });
    }

    const siteParam = req.query.site != null ? String(req.query.site).trim() : '';
    const { catalyst } = res.locals;
    const { rows, siteNames, selectedSiteState } = await buildMainReportEntries(catalyst, {
      year,
      month,
      userEmail,
      siteParam
    });

    res.status(200).json({
      status: 'success',
      data: { rows, siteNames, selectedSiteState: selectedSiteState || '' }
    });
  } catch (err) {
    console.error('mainreport/entries:', err);
    res.status(500).json({ status: 'failure', message: err.message || 'Failed to load entries.' });
  }
});

/** Draft files: State → Sector folders. PDFs are built in the app. */
app.get('/mainreport/consolidated-zip', async (req, res) => {
  try {
    const year = req.query.year;
    const month = req.query.month;
    const userEmail = req.query.userEmail != null ? String(req.query.userEmail) : '';
    if (year == null || year === '' || month == null || month === '') {
      return res.status(400).json({ status: 'failure', message: 'year and month query parameters are required.' });
    }

    const siteParam = req.query.site != null ? String(req.query.site).trim() : '';
    const monthNorm = normalizeMonth(month) || String(month).trim();
    const { catalyst } = res.locals;
    const { rows } = await buildMainReportEntries(catalyst, { year, month, userEmail, siteParam });

    const zipBase = `Consolidated_Report_${year}_${monthNorm}_${siteParam || 'All'}`.replace(/\s+/g, '_');
    const zipFileName = `${zipBase}.zip`;

    res.setHeader('Content-Type', 'application/zip');
    res.setHeader('Content-Disposition', `attachment; filename="${zipFileName.replace(/"/g, '')}"`);

    const archive = archiver('zip', { zlib: { level: 6 } });
    archive.on('error', (err) => {
      console.error('mainreport/consolidated-zip archive:', err);
      if (!res.headersSent) {
        res.status(500).json({ status: 'failure', message: err.message || 'Failed to build ZIP.' });
      } else {
        res.end();
      }
    });
    archive.pipe(res);

    const folder = catalyst.filestore().folder(DRAFT_FOLDER_ID);
    const usedNamesByPath = new Map();
    let filesAdded = 0;

    for (const row of rows) {
      const fileId = row?.draftFileId || row?.draftFile;
      if (!fileId) continue;

      const stateDir = getConsolidatedStateFolder(row);
      const sectorDir = getConsolidatedSectorFolder(row);
      const pathKey = `${stateDir}/${sectorDir}`;
      if (!usedNamesByPath.has(pathKey)) usedNamesByPath.set(pathKey, new Set());
      const usedInPath = usedNamesByPath.get(pathKey);

      let fileName = String(row.draftFileName || row.formName || 'draft').trim() || 'draft';
      fileName = ensureUniqueZipName(fileName, usedInPath);

      try {
        const buffer = await folder.downloadFile(fileId);
        archive.append(buffer, { name: `${stateDir}/${sectorDir}/${fileName}` });
        filesAdded += 1;
      } catch (err) {
        console.warn('mainreport/consolidated-zip skip file:', fileId, err?.message || err);
      }
    }

    if (filesAdded === 0) {
      const note =
        'No draft files in ZIP. Use Consolidated Report in the app for State → Sector PDF folders.\n';
      archive.append(Buffer.from(note, 'utf8'), { name: 'README.txt' });
    }

    await archive.finalize();
  } catch (err) {
    console.error('mainreport/consolidated-zip:', err);
    if (!res.headersSent) {
      res.status(500).json({ status: 'failure', message: err.message || 'Failed to build consolidated ZIP.' });
    }
  }
});

function looksLikeInvalidReturnedSiteValue(value) {
  const s = String(value || '').trim();
  if (!s) return true;
  if (/^\d{1,2}:\d{2}:\d{2,}$/.test(s)) return true;
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return true;
  if (/^\d{10,}$/.test(s)) return true;
  return false;
}

function getReturnedRowYear(row) {
  const ct = row?.CREATEDTIME || row?.createdtime || row?.MODIFIEDTIME || row?.modifiedtime;
  if (ct) {
    const d = new Date(ct);
    if (!Number.isNaN(d.getTime())) return d.getFullYear();
  }
  const due = row?.DueDate || row?.dueDate;
  if (due) {
    const d = new Date(due);
    if (!Number.isNaN(d.getTime())) return d.getFullYear();
  }
  return null;
}

function returnedRowMatchesYear(row, yearParam) {
  const y = getReturnedRowYear(row);
  if (y == null) return true;
  return yearMatch(y, yearParam);
}

function returnedRowMatchesMonth(row, monthParam) {
  const mf = row?.MonthFilter || row?.monthFilter || '';
  if (!mf) return false;
  return monthsMatch(mf, monthParam);
}

function buildReturnedReportApiRow(row, context = {}) {
  const {
    siteStateByName = null,
    siteInfoByName = null,
    bulkRows = [],
    statutoryRows = [],
    sectorStateMap = null
  } = context;

  const siteRaw = String(row?.Site || row?.site || '').trim();
  const site = looksLikeInvalidReturnedSiteValue(siteRaw) ? '' : siteRaw;
  const siteLower = site.toLowerCase();
  const act = getActFromAnyRow(row);
  const formName = getFormNameFromAnyRow(row);
  const formKey = getMasterKeyFromAnyRow(row);
  const statutoryMatch = findStatutoryMatchForReturned(row, statutoryRows);
  const siteInfo = siteLower && siteInfoByName ? siteInfoByName.get(siteLower) : null;

  const masterState =
    findStateForMasterKey(formKey, bulkRows, statutoryRows) ||
    findStateByFormNameLoose(formName, bulkRows, statutoryRows);
  const masterSector =
    findSectorForMasterKey(formKey, bulkRows, statutoryRows) ||
    findSectorByFormNameLoose(formName, bulkRows, statutoryRows);

  const sectorHint =
    masterSector ||
    getSectorFromAnyRow(row) ||
    getSectorFromAnyRow(statutoryMatch) ||
    siteInfo?.sectorFolder ||
    '';

  let state = resolveReportState(statutoryMatch || row, '', siteStateByName, site, {
    masterState,
    sectorStateMap,
    sector: sectorHint
  });
  if (!state && siteInfo?.state) state = siteInfo.state;
  if (!state && statutoryMatch) state = getStateFromAnyRow(statutoryMatch);

  const sector = getSectorFolderFromActOrIndustry(
    act,
    siteInfo?.industry || '',
    sectorHint
  );

  return {
    rowId: row?.ROWID != null ? String(row.ROWID) : null,
    formName,
    act,
    description: getDescriptionFromAnyRow(row),
    site,
    state,
    sector,
    status: String(row?.Status || row?.status || 'Returned').trim() || 'Returned',
    remarks: String(row?.Remarks || row?.remarks || '').trim(),
    dueDate: String(row?.DueDate || row?.dueDate || '').trim(),
    returnedAt: row?.CREATEDTIME || row?.createdtime || row?.MODIFIEDTIME || null
  };
}

async function loadScopedReturnedRows(catalyst, userEmail, siteParam) {
  const inchargeScope = await getInchargeReportScope(catalyst, userEmail);
  const allowedActCategories = inchargeScope.actCategories;

  let narrowActCategories = allowedActCategories;
  if (siteParam) {
    const siteRow = await findSiteRowByName(catalyst, siteParam);
    if (!siteRow) return { rows: [], inchargeScope, allowedActCategories };
    const siteCat = getActCategoryForSiteRow(siteRow, siteParam);
    const siteCats = [siteCat];
    if (Array.isArray(allowedActCategories) && allowedActCategories.length > 0) {
      const intersection = siteCats.filter((c) => allowedActCategories.includes(c));
      if (intersection.length === 0) {
        return { rows: [], inchargeScope, allowedActCategories };
      }
      narrowActCategories = intersection;
    } else {
      narrowActCategories = siteCats;
    }
  }

  const returnedTable = catalyst.datastore().table('Returned');
  const all = await returnedTable.getAllRows();
  let scoped = filterRowsByActCategoryScope(all, narrowActCategories);

  if (inchargeScope.siteNamesLower && inchargeScope.siteNamesLower.size > 0) {
    scoped = scoped.filter((r) => {
      const site = String(r.Site || r.site || '')
        .trim()
        .toLowerCase();
      return site && inchargeScope.siteNamesLower.has(site);
    });
  }

  if (siteParam) {
    const want = String(siteParam).trim().toLowerCase();
    scoped = scoped.filter((r) => {
      const site = String(r.Site || r.site || '')
        .trim()
        .toLowerCase();
      return site === want;
    });
  }

  scoped = scoped.filter((r) => !looksLikeInvalidReturnedSiteValue(r.Site || r.site));

  return { rows: scoped, inchargeScope, allowedActCategories };
}

/** Years/months derived from Returned table row timestamps and MonthFilter. */
app.get('/mainreport/returned-summary', async (req, res) => {
  try {
    res.set('Cache-Control', 'no-store');
    const { catalyst } = res.locals;
    const currentYear = new Date().getFullYear();
    const yearNums = new Set([currentYear]);
    const monthsByYear = {};

    try {
      const all = await catalyst.datastore().table('Returned').getAllRows();
      for (const r of all) {
        const y = getReturnedRowYear(r);
        if (y == null || Number.isNaN(y)) continue;
        yearNums.add(y);
        const yk = String(y);
        const m = normalizeMonth(r.MonthFilter || r.monthFilter || '');
        if (!m) continue;
        if (!monthsByYear[yk]) monthsByYear[yk] = [];
        monthsByYear[yk].push(m);
      }
    } catch (err) {
      console.warn('returned-summary: Returned table read failed:', err?.message || err);
    }

    for (const k of Object.keys(monthsByYear)) {
      monthsByYear[k] = sortMonths(monthsByYear[k]);
    }

    const years = [...yearNums].sort((a, b) => b - a).map(String);
    res.status(200).json({
      status: 'success',
      data: { years, monthsByYear, serverYear: currentYear }
    });
  } catch (err) {
    console.error('mainreport/returned-summary:', err);
    res.status(500).json({ status: 'failure', message: err.message || 'Failed to load returned report summary.' });
  }
});

app.get('/mainreport/returned-entries', async (req, res) => {
  try {
    res.set('Cache-Control', 'no-store');
    const year = req.query.year;
    const month = req.query.month;
    const userEmail = req.query.userEmail != null ? String(req.query.userEmail) : '';
    if (year == null || year === '' || month == null || month === '') {
      return res.status(400).json({ status: 'failure', message: 'year and month query parameters are required.' });
    }

    const { catalyst } = res.locals;
    const siteParam = req.query.site != null ? String(req.query.site).trim() : '';

    const { rows: scoped, inchargeScope, allowedActCategories } = await loadScopedReturnedRows(
      catalyst,
      userEmail,
      siteParam
    );

    const siteStateByName = await buildSiteStateByNameMap(catalyst);
    const siteInfoByName = await buildSiteInfoByNameMap(catalyst);

    const bulkAllRaw = await catalyst.datastore().table('checklistbulk').getAllRows();
    const bulkAll = (Array.isArray(bulkAllRaw) ? bulkAllRaw : []).map(normalizeChecklistBulkRowForReport);
    const sectorStateMap = buildSectorStateMap(bulkAll);

    let statutoryAll = [];
    try {
      statutoryAll = await catalyst.datastore().table('Statutory').getAllRows();
    } catch (err) {
      console.warn('returned-entries: Statutory table read failed:', err?.message || err);
    }

    const enrichContext = {
      siteStateByName,
      siteInfoByName,
      bulkRows: bulkAll,
      statutoryRows: Array.isArray(statutoryAll) ? statutoryAll : [],
      sectorStateMap
    };

    const filtered = scoped
      .filter((r) => returnedRowMatchesMonth(r, month))
      .filter((r) => returnedRowMatchesYear(r, year))
      .map((r) => buildReturnedReportApiRow(r, enrichContext));

    filtered.sort((a, b) =>
      String(a.formName || '').localeCompare(String(b.formName || ''), undefined, { sensitivity: 'base' })
    );

    const siteNames = await getAllSiteNames(catalyst, allowedActCategories, '', inchargeScope.siteNamesLower);

    res.status(200).json({
      status: 'success',
      data: { rows: filtered, siteNames, source: 'returned' }
    });
  } catch (err) {
    console.error('mainreport/returned-entries:', err);
    res.status(500).json({ status: 'failure', message: err.message || 'Failed to load returned entries.' });
  }
});

app.get('/mainreport/file/:fileId', async (req, res) => {
  try {
    const { fileId } = req.params;
    const name = (req.query.name && String(req.query.name)) || 'download';
    const { catalyst } = res.locals;
    const folder = catalyst.filestore().folder(DRAFT_FOLDER_ID);
    const fileBuffer = await folder.downloadFile(fileId);
    const contentType = getContentType(name);
    const disposition = req.query.disposition === 'attachment' ? 'attachment' : 'inline';
    res.writeHead(200, {
      'Content-Type': contentType,
      'Content-Disposition': `${disposition}; filename="${name.replace(/"/g, '')}"`,
      'Content-Length': fileBuffer.length
    });
    res.end(fileBuffer);
  } catch (err) {
    console.error('mainreport/file:', err);
    res.status(404).json({ status: 'failure', message: err.message || 'File not found.' });
  }
});

function isYetToCompleteStatusText(stLower) {
  return (
    String(stLower || '').includes('yet to complete') ||
    String(stLower || '').includes('yet to submit') ||
    (String(stLower || '').includes('yet') && String(stLower || '').includes('complete')) ||
    (String(stLower || '').includes('yet') && String(stLower || '').includes('submit')) ||
    String(stLower || '').includes('not started') ||
    String(stLower || '').includes('nostart') ||
    String(stLower || '').includes('incomplete') ||
    String(stLower || '').includes('not complete')
  );
}

function rowHasStoredDraft(row) {
  if (row?.hasStatutoryDraftStored === true) return true;
  const id = row?.draftFileId ?? row?.draftFile;
  return id != null && String(id).trim() !== '' && String(id).toLowerCase() !== 'null';
}

function rowSendForApprovalIsSent(row) {
  return /^sent$/i.test(String(row?.sendForApproval ?? row?.SendForApproval ?? '').trim());
}

/**
 * Pending / Approved / Returned only when a statutoryreg_function row exists for the month.
 * Checklistbulk forms with no statutory match → Yet to Complete (dashboard: Yet to Submit).
 * Same rules as Statutory.js getStatutoryRowDisplayStatus and Compliance Health KPIs.
 */
function getMainReportStatusLabel(row) {
  if (row?.statutoryRowId == null) return 'Yet to Complete';

  const st = String(row?.statutoryStatus || row?.status || '').trim();
  const stLower = st.toLowerCase();
  const appr = String(row?.approval || '').trim().toLowerCase();

  if (appr === 'rejected' || appr === 'reject' || stLower.includes('reject')) return 'Rejected';
  if (stLower === 'returned' || appr === 'returned') return 'Rejected';
  if (stLower === 'approved' || stLower === 'approve' || appr === 'approved' || appr === 'approve') {
    return 'Approved';
  }
  if (rowSendForApprovalIsSent(row)) return 'Pending';

  const hasDraft = rowHasStoredDraft(row);
  if (!hasDraft) return 'Yet to Complete';
  if (st === '' || st === '-' || st === '—') {
    return rowSendForApprovalIsSent(row) ? 'Pending' : 'Yet to Complete';
  }
  if (stLower === 'pending') {
    return rowSendForApprovalIsSent(row) ? 'Pending' : 'Yet to Complete';
  }
  if (isYetToCompleteStatusText(stLower)) {
    return rowSendForApprovalIsSent(row) ? 'Pending' : 'Yet to Complete';
  }
  return st || 'Yet to Complete';
}

/**
 * CHRO email cards: same rows as GET /mainreport/entries (Reports page, all sites).
 * Returned card uses the Returned Report for that month.
 */
async function buildChroRollupFromReports(catalyst, year, month) {
  const { rows, siteNames } = await buildMainReportEntries(catalyst, {
    year,
    month,
    userEmail: '',
    siteParam: ''
  });

  let approvedCount = 0;
  let pendingCount = 0;
  let yetToCompleteCount = 0;
  let rejectedCount = 0;
  for (const row of rows) {
    const status = getMainReportStatusLabel(row);
    if (status === 'Approved') approvedCount += 1;
    else if (status === 'Pending') pendingCount += 1;
    else if (status === 'Yet to Complete') yetToCompleteCount += 1;
    else if (status === 'Rejected') rejectedCount += 1;
  }

  let returnedRows = [];
  try {
    const loaded = await loadScopedReturnedRows(catalyst, '', '');
    returnedRows = (loaded.rows || [])
      .filter((r) => returnedRowMatchesMonth(r, month))
      .filter((r) => returnedRowMatchesYear(r, year));
  } catch (err) {
    console.warn('buildChroRollupFromReports: Returned table:', err?.message || err);
  }
  const returnedCount = returnedRows.length + rejectedCount;

  const siteInfoByName = await buildSiteInfoByNameMap(catalyst);

  const approvedForms = [];
  const seenApprovedKeys = new Set();

  const addApprovedForm = ({ siteName, stateHint, industryHint, formNumber, formName, statutoryRowId }) => {
    const site = String(siteName || '').trim();
    const number = String(formNumber || '').trim();
    const name = String(formName || '').trim();
    if (!site || (!number && !name)) return;
    const dedupeKey =
      statutoryRowId != null
        ? `id:${statutoryRowId}`
        : `${site.toLowerCase()}|${number.toLowerCase()}|${name.toLowerCase()}`;
    if (seenApprovedKeys.has(dedupeKey)) return;
    seenApprovedKeys.add(dedupeKey);
    const info = siteInfoByName.get(site.toLowerCase());
    const industry = String(industryHint || '').trim();
    approvedForms.push({
      siteName: site,
      label: site,
      state: stateHint || (info && info.state) || '—',
      industry: industry || '—',
      formNumber: number || '—',
      formName: name || '—',
      statutoryRowId: statutoryRowId != null ? String(statutoryRowId) : null
    });
  };

  /** Form Number = Statutory FormName; Form Name = Statutory Description (same as Reports page). */
  const statutoryFormNumber = (stRow, reportRow) =>
    String(getFormNameFromAnyRow(stRow) || getFormNameFromAnyRow(reportRow) || '').trim();

  const statutoryFormTitle = (stRow, reportRow) => {
    const desc = getDescriptionFromAnyRow(stRow) || String(reportRow?.description || '').trim();
    if (desc) return desc;
    return statutoryFormNumber(stRow, reportRow);
  };

  /** Industry column: Statutory table Sector (same field as statutoryreg_function). */
  const statutoryIndustryLabel = (stRow, reportRow) =>
    String(getSectorFromAnyRow(stRow) || String(reportRow?.sector || '').trim() || '').trim();

  const statutoryById = new Map();
  try {
    const statutoryAll = await catalyst.datastore().table('Statutory').getAllRows();
    for (const raw of statutoryAll) {
      const st = unwrapCatalystTableRow(raw, 'Statutory') || raw;
      if (st?.ROWID != null) statutoryById.set(String(st.ROWID), st);
    }
  } catch (err) {
    console.warn('buildChroRollupFromReports: Statutory read failed:', err?.message || err);
  }

  for (const row of rows) {
    if (getMainReportStatusLabel(row) !== 'Approved') continue;
    const st = row.statutoryRowId != null ? statutoryById.get(String(row.statutoryRowId)) : null;
    const siteName = String(st?.Site || st?.site || row.site || row.Site || '').trim();
    const formNumber = statutoryFormNumber(st, row);
    const formName = statutoryFormTitle(st, row);
    if (!siteName || (!formNumber && !formName)) continue;
    addApprovedForm({
      siteName,
      stateHint: getStateFromAnyRow(st) || row.state,
      industryHint: statutoryIndustryLabel(st, row),
      formNumber,
      formName,
      statutoryRowId: row.statutoryRowId
    });
  }

  approvedForms.sort((a, b) => {
    const siteCmp = String(a.siteName).localeCompare(String(b.siteName), undefined, { sensitivity: 'base' });
    if (siteCmp !== 0) return siteCmp;
    return String(a.formName).localeCompare(String(b.formName), undefined, { sensitivity: 'base' });
  });

  const approvedSites = approvedForms;
  const activeCount = Array.isArray(siteNames) ? siteNames.length : 0;
  const totalForms = Array.isArray(rows) ? rows.length : 0;
  const complianceScore = totalForms ? Math.round((approvedCount / totalForms) * 100) : 0;

  console.log('buildChroRollupFromReports', {
    year,
    month,
    reportRows: rows.length,
    approvedCount,
    pendingCount,
    yetToCompleteCount,
    returnedCount,
    activeCount,
    approvedSites: approvedSites.map((s) => `${s.label}: ${s.formName}`)
  });

  return {
    approvedSites,
    approvedCount,
    pendingCount,
    returnedCount,
    activeCount,
    complianceScore,
    yetToCompleteCount,
    reportRowCount: totalForms
  };
}

/** Same field mapping as Reports page (`Mainreport.js`). */
function getReportFormNumber(row) {
  return String(row?.formName || row?.formNumber || '').trim() || '—';
}

function getReportFormName(row) {
  const desc = String(row?.description || '').trim();
  if (desc) return desc;
  return String(row?.formName || '').trim() || '—';
}

function getReportState(row) {
  return String(row?.state || row?.zState || '').trim() || '—';
}

function getReportIndustry(row) {
  const sector = String(row?.sector || '').trim();
  if (sector && sector !== '—') return sector;
  const act = String(row?.act || '').trim();
  if (act && act !== '—') return act;
  return '—';
}

function getReportSiteName(row) {
  return String(row?.site || row?.Site || '').trim();
}

/**
 * Pending / Yet to Complete rows from Reports — one pass per site (same as selecting each site on Reports page).
 */
async function buildPendingRollupFromReports(catalyst, year, month) {
  const siteRows = await fetchActiveSiteRowsForReminder(catalyst);
  const siteInfoByName = await buildSiteInfoByNameMap(catalyst);
  const siteSections = [];
  const pendingSites = [];
  const seenKeys = new Set();
  const reportRowsBySiteKey = new Map();

  let approvedCount = 0;
  let pendingCount = 0;
  let yetToCompleteCount = 0;
  let totalForms = 0;

  for (const siteRow of siteRows) {
    const siteName = String(siteRow.SiteName || siteRow.siteName || '').trim();
    if (!siteName) continue;

    const siteLookupKey = siteName.toLowerCase();
    const cacheKey = String(siteRow.ROWID || `${siteLookupKey}|${getActCategoryForSiteRow(siteRow, siteName)}`);
    let rows = reportRowsBySiteKey.get(cacheKey);
    if (!rows) {
      const loaded = await buildMainReportEntries(catalyst, {
        year,
        month,
        userEmail: '',
        siteParam: siteName,
        siteRow
      });
      rows = (loaded.rows || []).filter((row) => reportRowMatchesSiteRowIndustry(row, siteRow, siteName));
      reportRowsBySiteKey.set(cacheKey, rows);
    }

    totalForms += rows.length;
    const siteForms = [];
    const info = siteInfoByName.get(siteLookupKey);
    const siteState = getSiteStateFromSiteRow(siteRow) || info?.state || '—';

    for (const row of rows) {
      const status = getMainReportStatusLabel(row);
      if (status === 'Approved') approvedCount += 1;
      else if (status === 'Pending') pendingCount += 1;
      else if (status === 'Yet to Complete') yetToCompleteCount += 1;

      if (status !== 'Pending' && status !== 'Yet to Complete') continue;

      const formNumber = getReportFormNumber(row);
      const formName = getReportFormName(row);
      if (formNumber === '—' && formName === '—') continue;

      const dedupeKey = `${siteLookupKey}|${row.bulkRowId || ''}|${row.statutoryRowId || ''}|${formNumber.toLowerCase()}|${formName.toLowerCase()}|${status}`;
      if (seenKeys.has(dedupeKey)) continue;
      seenKeys.add(dedupeKey);

      const state = getReportState(row);
      const resolvedState = state !== '—' ? state : siteState || '—';
      const siteKey = String(siteRow.ROWID || siteLookupKey);
      const formRow = {
        siteName,
        siteKey,
        label: siteName,
        state: resolvedState,
        industry: getReportIndustry(row),
        formNumber,
        formName,
        status,
        reportRowId: row.statutoryRowId || row.bulkRowId || null
      };
      siteForms.push(formRow);
      pendingSites.push(formRow);
    }

    siteSections.push({
      siteRowId: siteRow.ROWID || null,
      siteName,
      siteKey: String(siteRow.ROWID || siteLookupKey),
      label: siteName,
      state: siteState,
      siteIndustry: String(siteRow.Industry || siteRow.industry || info?.industry || '').trim(),
      forms: siteForms
    });
  }

  pendingSites.sort((a, b) => {
    const siteCmp = String(a.siteName).localeCompare(String(b.siteName), undefined, { sensitivity: 'base' });
    if (siteCmp !== 0) return siteCmp;
    const statusCmp = String(a.status).localeCompare(String(b.status), undefined, { sensitivity: 'base' });
    if (statusCmp !== 0) return statusCmp;
    return String(a.formName).localeCompare(String(b.formName), undefined, { sensitivity: 'base' });
  });

  siteSections.sort((a, b) =>
    String(a.siteName).localeCompare(String(b.siteName), undefined, { sensitivity: 'base' })
  );

  const sitesWithPending = siteSections.filter((section) => section.forms.length > 0).length;
  const sitesAllComplete = siteSections.length - sitesWithPending;
  const activeCount = siteRows.length;
  const complianceScore = totalForms ? Math.round((approvedCount / totalForms) * 100) : 0;

  console.log('buildPendingRollupFromReports', {
    year,
    month,
    activeSites: activeCount,
    sitesWithPending,
    sitesAllComplete,
    reportRowsScanned: totalForms,
    pendingCount,
    yetToCompleteCount,
    pendingFormRows: pendingSites.length
  });

  return {
    siteSections,
    pendingSites,
    approvedCount,
    pendingCount,
    yetToCompleteCount,
    activeCount,
    sitesWithPending,
    sitesAllComplete,
    complianceScore,
    reportRowCount: totalForms
  };
}

app.get('/mainreport/chro-rollup', async (req, res) => {
  try {
    res.set('Cache-Control', 'no-store');
    const year = req.query.year;
    const month = req.query.month;
    if (year == null || year === '' || month == null || month === '') {
      return res.status(400).json({ status: 'failure', message: 'year and month query parameters are required.' });
    }
    const { catalyst } = res.locals;
    const data = await buildChroRollupFromReports(catalyst, year, month);
    res.status(200).json({ status: 'success', data });
  } catch (err) {
    console.error('mainreport/chro-rollup:', err);
    res.status(500).json({ status: 'failure', message: err.message || 'Failed to load CHRO rollup.' });
  }
});

module.exports = {
  buildChroRollupFromReports,
  buildPendingRollupFromReports,
  buildMainReportEntries,
  getMainReportStatusLabel,
  buildSiteInfoByNameMap
};
