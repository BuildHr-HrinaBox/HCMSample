import { fetchAttendanceAll } from './attendanceApi';
import { fetchSamplePayrollRecords } from './samplePayrollApi';
import {
  fetchClraData,
  flattenClraEmployeesForAutofill,
  getCachedClraData,
  prefetchClraData,
} from './CLRA';

const SITE_DETAILS_CACHE_KEY = 'statutorySiteDetails_v1';
/** Site Management page writes full rows (incl. contractor) here on fetch/save. */
const SITE_MANAGEMENT_UI_CACHE_KEY = 'siteManagementData';
const PEOPLE_CACHE_KEY = 'statutoryPeopleData_v3';
const PEOPLE_CACHE_TTL_MS = 5 * 60 * 1000;
const PEOPLE_PAGE_SIZE = 300;
const PEOPLE_PAGE_DELAY_MS = 0;
const PEOPLE_FAST_FIRST_PAGE_SIZE = 80;
const ATTENDANCE_CACHE_TTL_MS = 5 * 60 * 1000;
const LEAVE_CACHE_TTL_MS = 5 * 60 * 1000;
const PAYROLL_BULK_CACHE_TTL_MS = 5 * 60 * 1000;
const FORM_TEMPLATE_TTL_MS = 30 * 60 * 1000;

let siteDetailsMemory = null;
let siteDetailsInflight = null;
let peopleMemory = null;
let peopleMemoryTs = 0;
let peopleInflight = null;
const formTemplateCache = new Map();
const formTemplateInflight = new Map();

function siteRecordNameKey(site) {
  return String(site?.siteName ?? site?.SiteName ?? site?.name ?? site?.Name ?? '')
    .trim()
    .toLowerCase();
}

function siteRecordIdKey(site) {
  return String(site?.id ?? site?.ID ?? site?.ROWID ?? '').trim();
}

/** Latest Site Management UI rows (contractor name/address included after Edit Site save). */
export function readSiteManagementUiCache() {
  try {
    const cached = localStorage.getItem(SITE_MANAGEMENT_UI_CACHE_KEY);
    if (!cached) return [];
    const parsed = JSON.parse(cached);
    return Array.isArray(parsed) ? parsed : [];
  } catch (_) {
    return [];
  }
}

/**
 * Prefer Site Management CONTRACTOR DETAILS onto statutory site rows.
 * Statutory cache / bootstrap meta sometimes omits contractorName after list fetch.
 */
export function mergeSiteDetailsWithContractorFields(details, uiSites = null) {
  const local = Array.isArray(uiSites) ? uiSites : readSiteManagementUiCache();
  const base = Array.isArray(details) ? details : [];
  if (local.length === 0) return base;
  if (base.length === 0) return local;

  const locationKey = (site) =>
    String(site?.location ?? site?.Location ?? '')
      .trim()
      .toLowerCase();

  return base.map((site) => {
    if (!site || typeof site !== 'object') return site;
    const id = siteRecordIdKey(site);
    const name = siteRecordNameKey(site);
    const loc = locationKey(site);
    const hit = local.find((s) => {
      if (!s || typeof s !== 'object') return false;
      const hid = siteRecordIdKey(s);
      if (id && hid && id === hid) return true;
      const hname = siteRecordNameKey(s);
      if (name && hname && name === hname) return true;
      const hloc = locationKey(s);
      return Boolean(loc && hloc && loc === hloc);
    });
    if (!hit) return site;
    const pick = (a, b) => {
      const av = String(a ?? '').trim();
      const bv = String(b ?? '').trim();
      // UI cache wins when present — Edit Site is the source of truth for contractor.
      return bv || av || '';
    };
    return {
      ...site,
      contractorName: pick(site.contractorName ?? site.ContractorName, hit.contractorName ?? hit.ContractorName),
      contractorAddress: pick(
        site.contractorAddress ?? site.ContractorAddress,
        hit.contractorAddress ?? hit.ContractorAddress
      ),
      contractorCity: pick(site.contractorCity ?? site.ContractorCity, hit.contractorCity ?? hit.ContractorCity),
      contractorState: pick(
        site.contractorState ?? site.ContractorState,
        hit.contractorState ?? hit.ContractorState
      ),
      contractorEmail: pick(site.contractorEmail ?? site.ContractorEmail, hit.contractorEmail ?? hit.ContractorEmail),
      contractorPhone: pick(site.contractorPhone ?? site.ContractorPhone, hit.contractorPhone ?? hit.ContractorPhone),
    };
  });
}

export function readSiteDetailsCache() {
  if (Array.isArray(siteDetailsMemory) && siteDetailsMemory.length > 0) {
    return mergeSiteDetailsWithContractorFields(siteDetailsMemory);
  }
  try {
    const cached = localStorage.getItem(SITE_DETAILS_CACHE_KEY);
    if (!cached) {
      const fromUi = readSiteManagementUiCache();
      return fromUi.length > 0 ? fromUi : null;
    }
    const parsed = JSON.parse(cached);
    if (Array.isArray(parsed) && parsed.length > 0) {
      siteDetailsMemory = parsed;
      return mergeSiteDetailsWithContractorFields(parsed);
    }
  } catch (_) {
    /* ignore */
  }
  const fromUi = readSiteManagementUiCache();
  return fromUi.length > 0 ? fromUi : null;
}

export function writeSiteDetailsCache(details) {
  if (!Array.isArray(details)) return;
  siteDetailsMemory = details;
  try {
    localStorage.setItem(SITE_DETAILS_CACHE_KEY, JSON.stringify(details));
  } catch (_) {
    /* ignore */
  }
}

export function fetchSiteDetails(options = {}) {
  const { force = false } = options;
  const cached = readSiteDetailsCache();
  if (!force && cached) return Promise.resolve(cached);
  if (!force && siteDetailsInflight) return siteDetailsInflight;

  siteDetailsInflight = fetch('/server/sitemanagement_function/sitemanagement', { cache: 'no-store' })
    .then(async (resp) => {
      if (!resp.ok) return cached || mergeSiteDetailsWithContractorFields([]) || [];
      const json = await resp.json().catch(() => ({}));
      const details = Array.isArray(json?.data?.siteDetails) ? json.data.siteDetails : [];
      if (details.length > 0) writeSiteDetailsCache(details);
      const merged = mergeSiteDetailsWithContractorFields(details.length > 0 ? details : cached || []);
      return merged.length > 0 ? merged : cached || [];
    })
    .catch(() => cached || mergeSiteDetailsWithContractorFields([]) || [])
    .finally(() => {
      siteDetailsInflight = null;
    });

  return siteDetailsInflight;
}

export function prefetchSiteDetails() {
  return fetchSiteDetails().catch(() => null);
}

const COMPANY_DETAILS_CACHE_KEY = 'statutoryCompanyDetails_v1';
let companyDetailsMemory = null;
let companyDetailsInflight = null;

export function readCompanyDetailsCache() {
  if (Array.isArray(companyDetailsMemory) && companyDetailsMemory.length > 0) return companyDetailsMemory;
  try {
    const cached = localStorage.getItem(COMPANY_DETAILS_CACHE_KEY);
    if (!cached) return null;
    const parsed = JSON.parse(cached);
    if (Array.isArray(parsed) && parsed.length > 0) {
      companyDetailsMemory = parsed;
      return parsed;
    }
  } catch (_) {
    /* ignore */
  }
  return null;
}

export function writeCompanyDetailsCache(details) {
  if (!Array.isArray(details)) return;
  companyDetailsMemory = details;
  try {
    localStorage.setItem(COMPANY_DETAILS_CACHE_KEY, JSON.stringify(details));
  } catch (_) {
    /* ignore */
  }
}

export function fetchCompanyDetails(options = {}) {
  const { force = false } = options;
  const cached = readCompanyDetailsCache();
  if (!force && cached) return Promise.resolve(cached);
  if (!force && companyDetailsInflight) return companyDetailsInflight;

  companyDetailsInflight = fetch('/server/company_function/company', { cache: 'no-store' })
    .then(async (resp) => {
      if (!resp.ok) return cached || [];
      const json = await resp.json().catch(() => ({}));
      const details = Array.isArray(json?.data?.companyDetails) ? json.data.companyDetails : [];
      if (details.length > 0) writeCompanyDetailsCache(details);
      return details.length > 0 ? details : cached || [];
    })
    .catch(() => cached || [])
    .finally(() => {
      companyDetailsInflight = null;
    });

  return companyDetailsInflight;
}

export function prefetchCompanyDetails() {
  return fetchCompanyDetails().catch(() => null);
}

export async function ensureStatutoryCompanyDetailsList(existingList = null, { force = false } = {}) {
  if (!force && Array.isArray(existingList) && existingList.length > 0) return existingList;
  if (!force) {
    const cached = readCompanyDetailsCache();
    if (Array.isArray(cached) && cached.length > 0) return cached;
  }
  try {
    const fetched = await fetchCompanyDetails({ force: !!force });
    if (Array.isArray(fetched) && fetched.length > 0) {
      writeCompanyDetailsCache(fetched);
      return fetched;
    }
  } catch (_) {
    /* fall through */
  }
  const cached = readCompanyDetailsCache();
  if (Array.isArray(cached) && cached.length > 0) return cached;
  if (Array.isArray(existingList) && existingList.length > 0) return existingList;
  return [];
}

function readPeopleCacheRaw() {
  if (peopleMemory && Date.now() - peopleMemoryTs < PEOPLE_CACHE_TTL_MS) {
    return peopleMemory;
  }
  try {
    const cached = localStorage.getItem(PEOPLE_CACHE_KEY);
    if (!cached) return null;
    // Huge sync JSON.parse blocks Autofill / Gujarat CLRA form open — skip oversized caches.
    if (cached.length > 2_500_000) {
      localStorage.removeItem(PEOPLE_CACHE_KEY);
      return null;
    }
    const parsed = JSON.parse(cached);
    if (parsed?.ts && parsed?.data && Date.now() - parsed.ts < PEOPLE_CACHE_TTL_MS) {
      peopleMemory = parsed.data;
      peopleMemoryTs = parsed.ts;
      return parsed.data;
    }
  } catch (_) {
    /* ignore */
  }
  return null;
}

export function writePeopleCache(data) {
  peopleMemory = data;
  peopleMemoryTs = Date.now();
  const ts = peopleMemoryTs;
  const schedulePersist =
    typeof requestIdleCallback === 'function'
      ? (fn) => requestIdleCallback(fn, { timeout: 2500 })
      : (fn) => setTimeout(fn, 0);
  // Persist off the critical path — stringify of large People payloads freezes Autofill UI.
  schedulePersist(() => {
    try {
      localStorage.setItem(PEOPLE_CACHE_KEY, JSON.stringify({ ts, data }));
    } catch (_) {
      /* ignore */
    }
  });
}

export function clearPeopleCache() {
  peopleMemory = null;
  peopleMemoryTs = 0;
  try {
    localStorage.removeItem(PEOPLE_CACHE_KEY);
  } catch (_) {
    /* ignore */
  }
}

/** Flatten Zoho People employee API payloads into an array of employee objects. */
export function flattenZohoPeopleEmployees(apiResult) {
  if (!apiResult) return [];

  const data = apiResult.data !== undefined ? apiResult.data : apiResult;
  const employees = [];

  const pushEmployee = (value) => {
    if (!value || typeof value !== 'object' || Array.isArray(value)) return;
    employees.push(value);
  };

  const pushMany = (values) => {
    if (!Array.isArray(values)) return;
    values.forEach((item) => pushEmployee(item));
  };

  const flattenNode = (node) => {
    if (!node) return;
    if (Array.isArray(node)) {
      node.forEach((item) => {
        if (!item || typeof item !== 'object') return;
        const keys = Object.keys(item);
        const looksLikeIdMap = keys.some((k) => /^\d{6,}$/.test(String(k)));
        if (looksLikeIdMap) {
          keys.forEach((key) => {
            const val = item[key];
            if (Array.isArray(val)) pushMany(val);
            else pushEmployee(val);
          });
          return;
        }
        pushEmployee(item);
      });
      return;
    }
    if (typeof node !== 'object') return;

    const keys = Object.keys(node);
    const looksLikeIdMap = keys.length > 0 && keys.every((k) => /^\d{6,}$/.test(String(k)));
    if (looksLikeIdMap) {
      keys.forEach((key) => {
        const val = node[key];
        if (Array.isArray(val)) pushMany(val);
        else pushEmployee(val);
      });
      return;
    }

    if (node.Employee) {
      pushEmployee(node);
      return;
    }

    pushEmployee(node);
  };

  const candidates = [
    data?.response?.result,
    data?.result,
    data?.response?.result?.record,
    data?.response?.record,
    data?.records,
    data?.record,
    data,
  ];

  for (const candidate of candidates) {
    const before = employees.length;
    flattenNode(candidate);
    if (employees.length > before) return employees;
  }

  if (Array.isArray(data)) return data;
  return employees;
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function parsePeopleApiResponse(response) {
  const text = await response.text();
  let json = {};
  try {
    json = text ? JSON.parse(text) : {};
  } catch {
    if (/access denied/i.test(text)) {
      throw new Error(
        'Access Denied — Zoho People OAuth may be expired. Redeploy peopledata_function and clear stale ZOHO_ACCESS_TOKEN in Catalyst env.'
      );
    }
    throw new Error(
      `People server returned HTTP ${response.status} (not JSON). Redeploy peopledata_function and check Catalyst logs.`
    );
  }
  return json;
}

function formatPeopleFetchError(pageResult, response) {
  const raw = pageResult.error || pageResult.message || 'Failed to load employee data';
  if (/access denied/i.test(String(raw))) {
    return `${raw} — redeploy peopledata_function; remove expired ZOHO_ACCESS_TOKEN from Catalyst env if set.`;
  }
  return raw;
}

async function fetchPeopleDataPaginated(pageSize = PEOPLE_PAGE_SIZE) {
  const merged = [];
  let sIndex = 1;

  while (true) {
    const qs = new URLSearchParams({
      form: 'employee',
      limit: String(pageSize),
      sIndex: String(sIndex),
    });
    const response = await fetch(`/server/peopledata_function?${qs.toString()}`, {
      cache: 'no-store',
      credentials: 'include',
    });
    const pageResult = await parsePeopleApiResponse(response);

    if (!response.ok || !pageResult.success) {
      // If we already have rows, stop pagination instead of failing autofill (Zoho 400 past last page).
      if (merged.length > 0 && (response.status === 400 || pageResult.httpStatus === 400)) {
        break;
      }
      const err = new Error(formatPeopleFetchError(pageResult, response));
      err.httpStatus = pageResult.httpStatus || response.status;
      throw err;
    }

    const batch = flattenZohoPeopleEmployees({ data: pageResult.data });
    merged.push(...batch);
    if (batch.length < pageSize) break;
    sIndex += pageSize;
    await sleep(PEOPLE_PAGE_DELAY_MS);
  }

  if (merged.length === 0) {
    const err = new Error('No employee data received from Zoho People API');
    throw err;
  }

  return {
    success: true,
    data: { response: { result: merged, status: 0 } },
    meta: { total: merged.length, mode: 'paginated' },
  };
}

async function fetchPeopleDataFromApi(options = {}) {
  const pageSize = Math.min(PEOPLE_PAGE_SIZE, Math.max(1, parseInt(options.limit, 10) || PEOPLE_PAGE_SIZE));

  try {
    const allQs = new URLSearchParams({
      form: 'employee',
      fetch_all: '1',
      limit: String(pageSize),
    });
    const allResponse = await fetch(`/server/peopledata_function?${allQs.toString()}`, {
      cache: 'no-store',
      credentials: 'include',
    });
    const allResult = await parsePeopleApiResponse(allResponse);

    if (allResponse.ok && allResult.success && allResult.data) {
      const employees = flattenZohoPeopleEmployees({ data: allResult.data });
      if (employees.length > 0) {
        return {
          success: true,
          data: { response: { result: employees, status: 0 } },
          meta: { ...(allResult.meta || {}), total: employees.length, mode: 'fetch_all' },
        };
      }
    }
  } catch (fetchAllErr) {
    console.warn('People fetch_all failed, using paginated fetch:', fetchAllErr?.message || fetchAllErr);
  }

  return fetchPeopleDataPaginated(pageSize);
}

export function fetchPeopleData(options = {}) {
  const { force = false } = options;

  if (!force) {
    const cached = readPeopleCacheRaw();
    if (cached) return Promise.resolve(cached);
    if (peopleInflight) return peopleInflight;
  } else {
    clearPeopleCache();
    if (peopleInflight) {
      return peopleInflight
        .catch(() => null)
        .then(() => fetchPeopleData({ force: true }));
    }
  }

  peopleInflight = fetchPeopleDataFromApi(options)
    .then((result) => {
      writePeopleCache(result);
      return result;
    })
    .finally(() => {
      peopleInflight = null;
    });

  return peopleInflight;
}

/** True when cache is only the fast first page and must still be upgraded to the full list. */
export function isPartialPeopleCache(cached = readPeopleCacheRaw()) {
  if (!cached?.success) return false;
  return cached?.meta?.mode === 'fast_first_page';
}

/** Continue loading full employee list after a fast first-page autofill response. */
export function startPeopleDataBackgroundRefresh(options = {}) {
  const cached = readPeopleCacheRaw();
  // A fast_first_page cache is intentionally size==80 — do NOT treat that as complete,
  // or Autofill never upgrades and site filters (e.g. GJ-Amreli) miss employees past page 1.
  if (
    cached?.success &&
    !isPartialPeopleCache(cached) &&
    flattenZohoPeopleEmployees(cached).length > 0
  ) {
    return Promise.resolve(cached);
  }
  if (peopleInflight) return peopleInflight;
  // Force when upgrading a partial cache; otherwise fetchPeopleData would return the 80-row page again.
  return fetchPeopleData({ ...options, force: isPartialPeopleCache(cached) }).catch(() => null);
}

/** Await a complete People list (upgrades fast_first_page cache when needed). */
export async function ensureCompletePeopleData(options = {}) {
  const cached = readPeopleCacheRaw();
  if (cached?.success && !isPartialPeopleCache(cached)) {
    const count = flattenZohoPeopleEmployees(cached).length;
    if (count > 0) return cached;
  }
  if (peopleInflight) {
    try {
      const inflight = await peopleInflight;
      if (inflight?.success && !isPartialPeopleCache(inflight)) {
        return inflight;
      }
    } catch (_) {
      /* fall through */
    }
  }
  return fetchPeopleData({ ...options, force: true });
}

/**
 * Statutory Autofill: return cached or first Zoho page immediately (<1s target),
 * then refresh the full list in the background for pagination.
 * Never await the full fetch_all inflight — that hangs Gujarat CLRA / all Autofill until every employee arrives.
 */
export async function fetchPeopleDataForAutofillDisplay(options = {}) {
  const cached = readPeopleCacheRaw();
  if (cached?.success) {
    const count = flattenZohoPeopleEmployees(cached).length;
    if (count > 0) {
      // Keep returning quickly, but never strand Autofill on a first-page-only cache.
      if (isPartialPeopleCache(cached)) {
        startPeopleDataBackgroundRefresh(options);
      }
      return cached;
    }
  }

  const pageSize = Math.min(
    PEOPLE_FAST_FIRST_PAGE_SIZE,
    Math.max(1, parseInt(options.limit, 10) || PEOPLE_FAST_FIRST_PAGE_SIZE)
  );
  const qs = new URLSearchParams({
    form: 'employee',
    limit: String(pageSize),
    sIndex: '1',
  });
  const response = await fetch(`/server/peopledata_function?${qs.toString()}`, {
    cache: 'no-store',
    credentials: 'include',
  });
  const pageResult = await parsePeopleApiResponse(response);

  if (response.ok && pageResult.success) {
    const batch = flattenZohoPeopleEmployees({ data: pageResult.data });
    if (batch.length > 0) {
      // Keep memory warm for the next paint; full list continues in background.
      if (!peopleMemory) {
        peopleMemory = {
          success: true,
          data: { response: { result: batch, status: 0 } },
          meta: { total: batch.length, mode: 'fast_first_page' },
        };
        peopleMemoryTs = Date.now();
      }
      startPeopleDataBackgroundRefresh(options);
      return {
        success: true,
        data: { response: { result: batch, status: 0 } },
        meta: { total: batch.length, mode: 'fast_first_page' },
      };
    }
  }

  // Last resort: only wait on full fetch if first page failed and one is already running.
  if (peopleInflight) {
    try {
      return await peopleInflight;
    } catch (_) {
      /* fall through */
    }
  }
  return fetchPeopleData(options);
}

/** Minimum employees expected when cache is considered usable for statutory autofill. */
export function readPeopleCacheEmployeeCount() {
  const cached = readPeopleCacheRaw();
  if (!cached) return 0;
  return flattenZohoPeopleEmployees(cached).length;
}

/** Synchronous read — use before any network call so Statutory autofill can start instantly. */
export function getCachedPeopleData() {
  return readPeopleCacheRaw();
}

/** Payroll salary breakdown cache only (does not trigger all_salaries download). */
export function getCachedPayrollBulkRows() {
  return readPayrollBulkCacheRaw();
}

/** Most recent Payroll table snapshot held in session (any month). */
export function getLatestCachedPayrollTableRows() {
  let best = null;
  form15PayrollTableByMonth.forEach((entry, month) => {
    if (
      !entry ||
      entry.source !== 'sample_payroll' ||
      !Array.isArray(entry.rows) ||
      entry.rows.length === 0 ||
      Date.now() - entry.ts >= FORM15_PAYROLL_TABLE_TTL_MS
    ) {
      return;
    }
    if (!best || entry.ts > best.ts) {
      best = {
        payrollMonth: month,
        rows: entry.rows,
        meta: entry.meta || null,
        payDate: entry.payDate || '',
        ts: entry.ts,
        source: 'sample_payroll',
      };
    }
  });
  return best;
}

/** Session cache of SamplePayroll table rows — prefer table snapshots over legacy all_salaries bulk. */
export function getPayrollBulkRowsForAutofill() {
  const tableCached = getLatestCachedPayrollTableRows();
  if (tableCached?.rows?.length > 0 && tableCached.source === 'sample_payroll') {
    return tableCached.rows;
  }
  const bulk = readPayrollBulkCacheRaw();
  return Array.isArray(bulk) && bulk.length > 0 ? bulk : null;
}

/** Let the browser paint and process scroll/input before heavy autofill work continues. */
export function yieldToMain() {
  return new Promise((resolve) => {
    const sched =
      typeof window !== 'undefined' && window.scheduler && typeof window.scheduler.yield === 'function'
        ? window.scheduler
        : null;
    if (sched) {
      sched.yield().then(resolve, () => setTimeout(resolve, 0));
      return;
    }
    // Double-rAF + macrotask: lets the browser paint and handle wheel/touch before we resume.
    if (typeof requestAnimationFrame === 'function') {
      requestAnimationFrame(() => {
        requestAnimationFrame(() => setTimeout(resolve, 0));
      });
      return;
    }
    setTimeout(resolve, 0);
  });
}

/** Extra idle turn after paint so the form modal can scroll while autofill continues. */
export async function yieldForModalScroll(isUiBusy) {
  await yieldToMain();
  await new Promise((resolve) => setTimeout(resolve, 48));
  if (typeof isUiBusy === 'function') {
    await waitWhileUiBusy(isUiBusy);
  }
}

/** Wait while the UI is busy (e.g. user scrolling the autofill modal). */
export async function waitWhileUiBusy(isUiBusy, idleMs = 80) {
  if (typeof isUiBusy !== 'function') return;
  let spins = 0;
  while (isUiBusy() && spins < 120) {
    spins += 1;
    await new Promise((resolve) => setTimeout(resolve, idleMs));
  }
}

/** Run a synchronous per-item handler in chunks so scroll/input stay responsive. */
export async function processInChunks(items, chunkSize, handler, options = {}) {
  const list = Array.isArray(items) ? items : [];
  const size = Math.max(1, chunkSize || 1);
  const isUiBusy = typeof options.isUiBusy === 'function' ? options.isUiBusy : null;
  for (let i = 0; i < list.length; i += 1) {
    if (isUiBusy) await waitWhileUiBusy(isUiBusy);
    handler(list[i], i);
    if ((i + 1) % size === 0) {
      await yieldToMain();
      if (isUiBusy) await waitWhileUiBusy(isUiBusy);
    }
  }
  if (list.length % size !== 0) {
    await yieldToMain();
  }
}

export async function fetchJsonWithTimeout(url, options = {}, timeoutMs = 15000) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);
  try {
    const resp = await fetch(url, { ...options, signal: controller.signal });
    const text = await resp.text();
    let json = {};
    try {
      json = text ? JSON.parse(text) : {};
    } catch (_) {
      json = { error: text || 'Invalid JSON response' };
    }
    return { resp, json };
  } catch (err) {
    if (err && err.name === 'AbortError') {
      const e = new Error(`Request timed out after ${timeoutMs}ms`);
      e.code = 'TIMEOUT';
      throw e;
    }
    throw err;
  } finally {
    clearTimeout(timer);
  }
}

export function prefetchPeopleData() {
  if (readPeopleCacheRaw()) return Promise.resolve(readPeopleCacheRaw());
  return fetchPeopleData().catch(() => null);
}

/** Fast first-page people load for instant statutory form open (<1s target). */
export function prefetchPeopleDataFast() {
  const cached = readPeopleCacheRaw();
  if (cached?.success && flattenZohoPeopleEmployees(cached).length > 0) {
    return Promise.resolve(cached);
  }
  // Do not return peopleInflight (full fetch_all) — that defeats the fast path.
  return fetchPeopleDataForAutofillDisplay().catch(() => null);
}

const MONTH_ABBR = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

function formatLocalYYYYMMDD(d) {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${y}-${m}-${day}`;
}

export function getDefaultAttendanceDateRange(referenceDate = new Date()) {
  const now = referenceDate instanceof Date ? referenceDate : new Date();
  const start = new Date(now.getFullYear(), now.getMonth(), 1);
  const end = new Date(now.getFullYear(), now.getMonth() + 1, 0);
  return { sdate: formatLocalYYYYMMDD(start), edate: formatLocalYYYYMMDD(end) };
}

export function getDefaultLeaveDateRange(referenceDate = new Date()) {
  const now = referenceDate instanceof Date ? referenceDate : new Date();
  const monthIdx = now.getMonth();
  const year = now.getFullYear();
  const first = new Date(year, monthIdx, 1);
  const last = new Date(year, monthIdx + 1, 0);
  const dd = (n) => String(n).padStart(2, '0');
  return {
    fromDate: `${dd(first.getDate())}-${MONTH_ABBR[first.getMonth()]}-${first.getFullYear()}`,
    toDate: `${dd(last.getDate())}-${MONTH_ABBR[last.getMonth()]}-${last.getFullYear()}`,
  };
}

function attendanceCacheKey(sdate, edate) {
  return `${sdate}|${edate}`;
}

function leaveCacheKey(fromDate, toDate, unit) {
  return `${fromDate}|${toDate}|${unit || 'Day'}`;
}

let attendanceMemory = new Map();
let attendanceInflight = new Map();
let leaveMemory = new Map();
let leaveInflight = new Map();

function readAttendanceCacheRaw(sdate, edate) {
  const key = attendanceCacheKey(sdate, edate);
  const entry = attendanceMemory.get(key);
  if (entry && Date.now() - entry.ts < ATTENDANCE_CACHE_TTL_MS) return entry.data;
  return null;
}

/** Synchronous attendance read for instant statutory autofill when prefetch already ran. */
export function getCachedAttendanceData(sdate, edate) {
  return readAttendanceCacheRaw(sdate, edate);
}

function writeAttendanceCache(sdate, edate, data) {
  const key = attendanceCacheKey(sdate, edate);
  attendanceMemory.set(key, { ts: Date.now(), data });
}

export function fetchAttendanceData(options = {}) {
  const { force = false } = options;
  const { sdate, edate } = {
    ...getDefaultAttendanceDateRange(),
    ...(options.sdate && options.edate ? { sdate: options.sdate, edate: options.edate } : {}),
  };
  const cached = readAttendanceCacheRaw(sdate, edate);
  if (!force && cached) return Promise.resolve(cached);

  const inflightKey = attendanceCacheKey(sdate, edate);
  // Reuse in-flight even when force — avoids duplicate Zoho calls (Form 25 prefetch + autofill).
  if (attendanceInflight.has(inflightKey)) {
    return attendanceInflight.get(inflightKey);
  }

  const timeoutMs = Math.max(15000, parseInt(options.timeoutMs, 10) || 60000);
  const task = fetchAttendanceAll(sdate, edate, timeoutMs)
    .then((result) => {
      if (!result?.success) {
        throw new Error(result?.error || 'Failed to load attendance data');
      }
      writeAttendanceCache(sdate, edate, result);
      return result;
    })
    .finally(() => {
      attendanceInflight.delete(inflightKey);
    });

  attendanceInflight.set(inflightKey, task);
  return task;
}

export function prefetchAttendanceData(options = {}) {
  const { sdate, edate } = {
    ...getDefaultAttendanceDateRange(),
    ...(options.sdate && options.edate ? { sdate: options.sdate, edate: options.edate } : {}),
  };
  if (readAttendanceCacheRaw(sdate, edate)) {
    return Promise.resolve(readAttendanceCacheRaw(sdate, edate));
  }
  return fetchAttendanceData({ sdate, edate }).catch(() => null);
}

function readLeaveCacheRaw(fromDate, toDate, unit) {
  const key = leaveCacheKey(fromDate, toDate, unit);
  const entry = leaveMemory.get(key);
  if (entry && Date.now() - entry.ts < LEAVE_CACHE_TTL_MS) return entry.data;
  return null;
}

/** Synchronous leave read for instant statutory autofill when prefetch already ran. */
export function getCachedLeaveData(fromDate, toDate, unit = 'Day') {
  return readLeaveCacheRaw(fromDate, toDate, unit);
}

function writeLeaveCache(fromDate, toDate, unit, data) {
  const key = leaveCacheKey(fromDate, toDate, unit);
  leaveMemory.set(key, { ts: Date.now(), data });
}

export function fetchLeaveData(options = {}) {
  const { force = false } = options;
  const unit = options.unit || 'Day';
  const { fromDate, toDate } = {
    ...getDefaultLeaveDateRange(),
    ...(options.fromDate && options.toDate ? { fromDate: options.fromDate, toDate: options.toDate } : {}),
  };
  const cached = readLeaveCacheRaw(fromDate, toDate, unit);
  if (!force && cached) return Promise.resolve(cached);

  const inflightKey = leaveCacheKey(fromDate, toDate, unit);
  // Reuse in-flight even when force — avoids duplicate Zoho leave calls during autofill.
  if (leaveInflight.has(inflightKey)) {
    return leaveInflight.get(inflightKey);
  }

  const timeoutMs = Math.max(5000, parseInt(options.timeoutMs, 10) || 20000);
  const leaveUrl = `/server/leavedata_function?from=${encodeURIComponent(fromDate)}&to=${encodeURIComponent(toDate)}&unit=${encodeURIComponent(unit)}&fetch_all=1`;
  const task = fetchJsonWithTimeout(leaveUrl, { cache: 'no-store' }, timeoutMs)
    .then(({ resp, json: result }) => {
      if (!resp.ok) {
        throw new Error(result.error || result.message || 'Failed to load leave data');
      }
      writeLeaveCache(fromDate, toDate, unit, result);
      return result;
    })
    .finally(() => {
      leaveInflight.delete(inflightKey);
    });

  leaveInflight.set(inflightKey, task);
  return task;
}

export function prefetchLeaveData(options = {}) {
  const unit = options.unit || 'Day';
  const { fromDate, toDate } = {
    ...getDefaultLeaveDateRange(),
    ...(options.fromDate && options.toDate ? { fromDate: options.fromDate, toDate: options.toDate } : {}),
  };
  if (readLeaveCacheRaw(fromDate, toDate, unit)) {
    return Promise.resolve(readLeaveCacheRaw(fromDate, toDate, unit));
  }
  return fetchLeaveData({ fromDate, toDate, unit }).catch(() => null);
}

const STORED_LEAVE_CACHE_TTL_MS = 5 * 60 * 1000;
/** @type {Map<string, { ts: number, data: object }>} */
let storedLeaveMemory = new Map();
/** @type {Map<string, Promise<object>>} */
let storedLeaveInflight = new Map();

function storedLeaveCacheKey(monthWise = '') {
  return `LeaveData|${monthWise || 'all'}`;
}

function readStoredLeaveCacheRaw(monthWise = '') {
  const key = storedLeaveCacheKey(monthWise);
  const entry = storedLeaveMemory.get(key);
  if (entry && Date.now() - entry.ts < STORED_LEAVE_CACHE_TTL_MS) return entry.data;
  return null;
}

function writeStoredLeaveCache(monthWise, data) {
  storedLeaveMemory.set(storedLeaveCacheKey(monthWise || ''), { ts: Date.now(), data });
}

/** Synchronous LeaveData table read when prefetch already ran. */
export function getCachedStoredLeaveData(monthWise = '') {
  const specific = readStoredLeaveCacheRaw(monthWise);
  if (specific) return specific;
  // Do not fall back to another month's cache when a specific month was requested.
  if (String(monthWise || '').trim()) return null;
  return readStoredLeaveCacheRaw('');
}

/**
 * Fetch leave rows from Catalyst LeaveData table (no live Zoho People call).
 * Used by Tamil Nadu Form 15 Part 1 and Form X.
 */
export function fetchStoredLeaveDataForAutofill(options = {}) {
  const { force = false } = options;
  const monthWise =
    (options.monthWise && String(options.monthWise).trim()) ||
    (options.fromDate && options.toDate
      ? (() => {
          const m = String(options.fromDate || '')
            .trim()
            .match(/^(\d{1,2})-([A-Za-z]{3})-(\d{4})$/);
          if (!m) return '';
          const mon = MONTH_ABBR.findIndex((x) => x.toLowerCase() === m[2].toLowerCase());
          if (mon < 0) return '';
          return `${m[3]}-${String(mon + 1).padStart(2, '0')}`;
        })()
      : '');

  const cached = readStoredLeaveCacheRaw(monthWise) || (!monthWise ? null : readStoredLeaveCacheRaw(''));
  if (!force && cached) return Promise.resolve(cached);

  const inflightKey = storedLeaveCacheKey(monthWise);
  if (!force && storedLeaveInflight.has(inflightKey)) {
    return storedLeaveInflight.get(inflightKey);
  }

  const timeoutMs = Math.max(5000, parseInt(options.timeoutMs, 10) || 60000);
  const params = new URLSearchParams();
  if (monthWise) params.set('monthWise', monthWise);
  const qs = params.toString();
  const url = `/server/leavedata_function/stored${qs ? `?${qs}` : ''}`;

  const task = fetchJsonWithTimeout(
    url,
    { cache: 'no-store', credentials: 'include' },
    timeoutMs
  )
    .then(({ resp, json: result }) => {
      if (!resp.ok) {
        throw new Error(result.error || result.message || 'Failed to load LeaveData table');
      }
      writeStoredLeaveCache(monthWise, result);
      if (monthWise) writeStoredLeaveCache('', result);
      return result;
    })
    .finally(() => {
      storedLeaveInflight.delete(inflightKey);
    });

  storedLeaveInflight.set(inflightKey, task);
  return task;
}

export function prefetchStoredLeaveData(options = {}) {
  const monthWise = options.monthWise || '';
  if (getCachedStoredLeaveData(monthWise)) {
    return Promise.resolve(getCachedStoredLeaveData(monthWise));
  }
  return fetchStoredLeaveDataForAutofill(options).catch(() => null);
}

const APPROVED_LEAVE_CACHE_TTL_MS = 5 * 60 * 1000;

function approvedLeaveCacheKey(fromDate, toDate) {
  return `${fromDate}|${toDate}`;
}

/** @type {Map<string, { ts: number, data: object }>} */
let approvedLeaveMemory = new Map();
/** @type {Map<string, Promise<object>>} */
let approvedLeaveInflight = new Map();

function readApprovedLeaveCacheRaw(fromDate, toDate) {
  const key = approvedLeaveCacheKey(fromDate, toDate);
  const entry = approvedLeaveMemory.get(key);
  if (entry && Date.now() - entry.ts < APPROVED_LEAVE_CACHE_TTL_MS) return entry.data;
  return null;
}

export function getCachedApprovedLeaveData(fromDate, toDate) {
  return readApprovedLeaveCacheRaw(fromDate, toDate);
}

function writeApprovedLeaveCache(fromDate, toDate, data) {
  const key = approvedLeaveCacheKey(fromDate, toDate);
  approvedLeaveMemory.set(key, { ts: Date.now(), data });
}

export function fetchApprovedLeaveData(options = {}) {
  const { force = false } = options;
  const { fromDate, toDate } = {
    ...(options.fromDate && options.toDate ? { fromDate: options.fromDate, toDate: options.toDate } : {}),
  };
  if (!fromDate || !toDate) {
    return Promise.reject(new Error('fromDate and toDate are required for approved leave fetch'));
  }
  const cached = readApprovedLeaveCacheRaw(fromDate, toDate);
  if (!force && cached) return Promise.resolve(cached);

  const inflightKey = approvedLeaveCacheKey(fromDate, toDate);
  if (!force && approvedLeaveInflight.has(inflightKey)) {
    return approvedLeaveInflight.get(inflightKey);
  }

  const timeoutMs = Math.max(30000, parseInt(options.timeoutMs, 10) || 120000);
  const url = `/server/approve_leave_function?from=${encodeURIComponent(fromDate)}&to=${encodeURIComponent(toDate)}&fetch_all=1`;
  const task = fetchJsonWithTimeout(url, { cache: 'no-store' }, timeoutMs)
    .then(({ resp, json: result }) => {
      if (!resp.ok) {
        throw new Error(result.error || result.message || 'Failed to load approved leave data');
      }
      writeApprovedLeaveCache(fromDate, toDate, result);
      return result;
    })
    .finally(() => {
      approvedLeaveInflight.delete(inflightKey);
    });

  approvedLeaveInflight.set(inflightKey, task);
  return task;
}

export function prefetchApprovedLeaveData(options = {}) {
  const { fromDate, toDate } = options;
  if (!fromDate || !toDate) return Promise.resolve(null);
  if (readApprovedLeaveCacheRaw(fromDate, toDate)) {
    return Promise.resolve(readApprovedLeaveCacheRaw(fromDate, toDate));
  }
  return fetchApprovedLeaveData({ fromDate, toDate }).catch(() => null);
}

let payrollBulkMemory = null;
let payrollBulkMemoryTs = 0;
let payrollBulkInflight = null;

function readPayrollBulkCacheRaw() {
  if (payrollBulkMemory && Date.now() - payrollBulkMemoryTs < PAYROLL_BULK_CACHE_TTL_MS) {
    return payrollBulkMemory;
  }
  return null;
}

export function writePayrollBulkCache(rows) {
  if (!Array.isArray(rows)) return;
  payrollBulkMemory = rows;
  payrollBulkMemoryTs = Date.now();
}

const FORM15_PAYROLL_TABLE_TTL_MS = 30 * 60 * 1000;
/** @type {Map<string, { rows: unknown[], ts: number, meta: object|null, payDate: string }>} */
const form15PayrollTableByMonth = new Map();

/** Reuse SamplePayroll rows across modal reopen (keyed by YYYY-MM). */
export function getCachedForm15PayrollTableRows(monthCandidates) {
  const list = Array.isArray(monthCandidates) ? monthCandidates : [];
  for (let i = 0; i < list.length; i += 1) {
    const month = String(list[i] || '').trim();
    if (!/^\d{4}-\d{2}$/.test(month)) continue;
    const entry = form15PayrollTableByMonth.get(month);
    if (
      entry &&
      entry.source === 'sample_payroll' &&
      Date.now() - entry.ts < FORM15_PAYROLL_TABLE_TTL_MS &&
      Array.isArray(entry.rows) &&
      entry.rows.length > 0
    ) {
      return {
        payrollMonth: month,
        rows: entry.rows,
        meta: entry.meta || null,
        payDate: entry.payDate || '',
        source: 'sample_payroll',
      };
    }
  }
  return null;
}

export function cacheForm15PayrollTableRows(
  payrollMonth,
  rows,
  meta = null,
  payDate = '',
  source = 'sample_payroll'
) {
  const month = String(payrollMonth || '').trim();
  if (!/^\d{4}-\d{2}$/.test(month) || !Array.isArray(rows) || rows.length === 0) return;
  form15PayrollTableByMonth.set(month, {
    rows,
    ts: Date.now(),
    meta: meta && typeof meta === 'object' ? meta : null,
    payDate: String(payDate || '').trim(),
    source,
  });
}

async function fetchPayrollBulkRowsBatched() {
  const sampleLoad = await fetchSamplePayrollRecords('', { timeoutMs: 60000 });
  return Array.isArray(sampleLoad.records) ? sampleLoad.records : [];
}

export function fetchPayrollBulkRows(options = {}) {
  const { force = false } = options;
  const cached = readPayrollBulkCacheRaw();
  if (!force && cached) return Promise.resolve(cached);
  if (!force && payrollBulkInflight) return payrollBulkInflight;

  payrollBulkInflight = fetchPayrollBulkRowsBatched()
    .then((rows) => {
      if (Array.isArray(rows) && rows.length > 0) {
        writePayrollBulkCache(rows);
        return rows;
      }
      return cached || [];
    })
    .catch(() => cached || [])
    .finally(() => {
      payrollBulkInflight = null;
    });

  return payrollBulkInflight;
}

export function prefetchPayrollBulkRows() {
  if (readPayrollBulkCacheRaw()) return Promise.resolve(readPayrollBulkCacheRaw());
  return fetchPayrollBulkRows().catch(() => null);
}

export function prefetchStatutoryAutofillData(options = {}) {
  prefetchSiteDetails();
  prefetchCompanyDetails();
  // First page first — then full list. Starting both together made Autofill await fetch_all.
  void prefetchPeopleDataFast().then(() => startPeopleDataBackgroundRefresh());
  const attOpts =
    options.sdate && options.edate ? { sdate: options.sdate, edate: options.edate } : {};
  prefetchAttendanceData(attOpts);
  const leaveOpts =
    options.fromDate && options.toDate
      ? { fromDate: options.fromDate, toDate: options.toDate, unit: options.unit || 'Day' }
      : {};
  prefetchLeaveData(leaveOpts);
  // Do not prefetch all_salaries here — statutory forms read from SamplePayroll table.
}

export {
  fetchClraData,
  flattenClraEmployeesForAutofill,
  getCachedClraData,
  prefetchClraData,
};

export function getFormTemplateCacheKey(item) {
  const rowId = String(item?.formFetchRowId ?? item?.id ?? '').trim();
  const id = String(item?.formFile || item?.FormFile || '').trim();
  const lineKey = [
    String(item?.formName || item?.FormName || '').trim().toLowerCase(),
    String(item?.act || item?.Act || '').trim().toLowerCase(),
    String(item?.description || item?.Description || '').trim().toLowerCase(),
    String(item?.sector || item?.Sector || '').trim().toLowerCase(),
    String(item?.state || item?.State || '').trim().toLowerCase()
  ].join('|');
  return `${rowId}|${id}|${lineKey}`;
}

/** True when a PK zip buffer includes an end-of-central-directory record (not truncated). */
export function hasZipEndOfCentralDirectory(arrayBuffer) {
  if (!arrayBuffer || arrayBuffer.byteLength < 22) return false;
  const u8 = new Uint8Array(arrayBuffer);
  const minEOCDSize = 22;
  const searchStart = Math.max(0, u8.length - 65557);
  for (let i = u8.length - minEOCDSize; i >= searchStart; i -= 1) {
    if (u8[i] === 0x50 && u8[i + 1] === 0x4b && u8[i + 2] === 0x05 && u8[i + 3] === 0x06) {
      return true;
    }
  }
  return false;
}

/** True when buffer looks like xlsx (PK zip) or legacy xls (OLE), not JSON/HTML error bodies. */
export function isValidExcelArrayBuffer(arrayBuffer) {
  if (!arrayBuffer || arrayBuffer.byteLength < 4) return false;
  const u8 = new Uint8Array(arrayBuffer.slice(0, 8));
  if (u8[0] === 0x50 && u8[1] === 0x4b) {
    return arrayBuffer.byteLength >= 512 && hasZipEndOfCentralDirectory(arrayBuffer);
  }
  if (u8[0] === 0xd0 && u8[1] === 0xcf && u8[2] === 0x11 && u8[3] === 0xe0) return true;
  const head = new TextDecoder('utf-8', { fatal: false })
    .decode(new Uint8Array(arrayBuffer.slice(0, Math.min(120, arrayBuffer.byteLength))))
    .trim()
    .toLowerCase();
  if (head.startsWith('<!doctype') || head.startsWith('<html')) return false;
  if (
    head.startsWith('{') &&
    (head.includes('"failure"') || head.includes('"status"') || head.includes('"error"') || head.includes('"message"'))
  ) {
    return false;
  }
  return false;
}

export function invalidateFormTemplateCache(cacheKey) {
  if (cacheKey) formTemplateCache.delete(cacheKey);
}

export function readFormTemplateCache(cacheKey) {
  if (!cacheKey) return null;
  const entry = formTemplateCache.get(cacheKey);
  if (entry && Date.now() - entry.ts < FORM_TEMPLATE_TTL_MS) {
    if (entry.arrayBuffer && !isValidExcelArrayBuffer(entry.arrayBuffer)) {
      formTemplateCache.delete(cacheKey);
      return null;
    }
    return entry;
  }
  if (entry) formTemplateCache.delete(cacheKey);
  return null;
}

export function writeFormTemplateCache(cacheKey, arrayBuffer, meta = {}) {
  if (!cacheKey || !arrayBuffer) return;
  formTemplateCache.set(cacheKey, {
    arrayBuffer,
    ts: Date.now(),
    ...meta
  });
}

export async function fetchFormTemplateArrayBuffer(urls, cacheKey, options = {}) {
  const { force = false } = options;
  const list = Array.isArray(urls) ? urls.filter(Boolean) : [];
  if (!list.length) throw new Error('No form file linked for this row');

  if (!force && cacheKey) {
    const cached = readFormTemplateCache(cacheKey);
    if (cached?.arrayBuffer) return cached;
  }

  if (!force && cacheKey && formTemplateInflight.has(cacheKey)) {
    return formTemplateInflight.get(cacheKey);
  }

  const task = fetchFirstOkResponse(list)
    .then(async ({ resp, url }) => {
      const arrayBuffer = await resp.arrayBuffer();
      if (!isValidExcelArrayBuffer(arrayBuffer)) {
        if (cacheKey) invalidateFormTemplateCache(cacheKey);
        throw new Error(
          'Form template download returned an invalid Excel file. Refresh the page and try again, or re-upload the template in Form Master.'
        );
      }
      const contentType = (resp.headers.get('content-type') || '').toLowerCase();
      const cdName = resp.headers.get('content-disposition') || '';
      const entry = { arrayBuffer, ts: Date.now(), url, contentType, contentDisposition: cdName };
      if (cacheKey) writeFormTemplateCache(cacheKey, arrayBuffer, { url, contentType, contentDisposition: cdName });
      return entry;
    })
    .finally(() => {
      if (cacheKey) formTemplateInflight.delete(cacheKey);
    });

  if (cacheKey) formTemplateInflight.set(cacheKey, task);
  return task;
}

export function prefetchFormTemplate(urls, cacheKey) {
  if (!cacheKey || !urls?.length) return Promise.resolve(null);
  if (readFormTemplateCache(cacheKey)) return Promise.resolve(readFormTemplateCache(cacheKey));
  return fetchFormTemplateArrayBuffer(urls, cacheKey).catch(() => null);
}

/** Fetch with retry/backoff when Catalyst returns HTTP 429 (rate limit). */
export async function fetchHttpWithRetry(url, options = {}) {
  const retries = Number(options.retries) > 0 ? Number(options.retries) : 3;
  const baseDelayMs = Number(options.baseDelayMs) > 0 ? Number(options.baseDelayMs) : 1500;
  const fetchOptions = options.fetchOptions || { cache: 'no-store' };
  let lastResp = null;
  for (let attempt = 0; attempt < retries; attempt += 1) {
    const resp = await fetch(url, fetchOptions);
    lastResp = resp;
    if (resp.ok) return resp;
    if (resp.status !== 429 || attempt >= retries - 1) return resp;
    await new Promise((resolve) => setTimeout(resolve, baseDelayMs * (attempt + 1)));
  }
  return lastResp;
}

/** Try all form file URLs in parallel; return the first successful response. */
export async function fetchFirstOkResponse(urls) {
  const list = Array.isArray(urls) ? urls.filter(Boolean) : [];
  if (!list.length) throw new Error('No form file linked for this row');

  return new Promise((resolve, reject) => {
    let pending = list.length;
    let lastErr = null;
    let settled = false;

    list.forEach(async (url) => {
      try {
        const resp = await fetchHttpWithRetry(url);
        if (settled) return;
        if (resp.ok) {
          settled = true;
          resolve({ resp, url });
          return;
        }
        const errorData = await resp.json().catch(() => ({}));
        lastErr = new Error(
          String(errorData.message || errorData.error || '').trim() || `Failed to load form file (HTTP ${resp.status})`
        );
      } catch (fetchErr) {
        if (!settled) {
          lastErr = fetchErr instanceof Error ? fetchErr : new Error('Failed to load form file');
        }
      } finally {
        if (settled) return;
        pending -= 1;
        if (pending === 0) reject(lastErr || new Error('Failed to load form file'));
      }
    });
  });
}