import { fetchAttendanceAll } from './attendanceApi';
import { getPayrollOrganizationId } from './payrollOrgId';
import {
  fetchClraData,
  flattenClraEmployeesForAutofill,
  getCachedClraData,
  prefetchClraData,
} from './CLRA';

const SITE_DETAILS_CACHE_KEY = 'statutorySiteDetails_v1';
const PEOPLE_CACHE_KEY = 'statutoryPeopleData_v3';
const PEOPLE_CACHE_TTL_MS = 5 * 60 * 1000;
const PEOPLE_PAGE_SIZE = 300;
const PEOPLE_PAGE_DELAY_MS = 0;
const PEOPLE_FAST_FIRST_PAGE_SIZE = 300;
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

export function readSiteDetailsCache() {
  if (Array.isArray(siteDetailsMemory) && siteDetailsMemory.length > 0) return siteDetailsMemory;
  try {
    const cached = localStorage.getItem(SITE_DETAILS_CACHE_KEY);
    if (!cached) return null;
    const parsed = JSON.parse(cached);
    if (Array.isArray(parsed) && parsed.length > 0) {
      siteDetailsMemory = parsed;
      return parsed;
    }
  } catch (_) {
    /* ignore */
  }
  return null;
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
      if (!resp.ok) return cached || [];
      const json = await resp.json().catch(() => ({}));
      const details = Array.isArray(json?.data?.siteDetails) ? json.data.siteDetails : [];
      if (details.length > 0) writeSiteDetailsCache(details);
      return details.length > 0 ? details : cached || [];
    })
    .catch(() => cached || [])
    .finally(() => {
      siteDetailsInflight = null;
    });

  return siteDetailsInflight;
}

export function prefetchSiteDetails() {
  return fetchSiteDetails().catch(() => null);
}

function readPeopleCacheRaw() {
  if (peopleMemory && Date.now() - peopleMemoryTs < PEOPLE_CACHE_TTL_MS) {
    return peopleMemory;
  }
  try {
    const cached = localStorage.getItem(PEOPLE_CACHE_KEY);
    if (!cached) return null;
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
  try {
    localStorage.setItem(PEOPLE_CACHE_KEY, JSON.stringify({ ts: peopleMemoryTs, data }));
  } catch (_) {
    /* ignore */
  }
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

/** Continue loading full employee list after a fast first-page autofill response. */
export function startPeopleDataBackgroundRefresh(options = {}) {
  const cached = readPeopleCacheRaw();
  if (cached?.success && flattenZohoPeopleEmployees(cached).length >= PEOPLE_FAST_FIRST_PAGE_SIZE) {
    return Promise.resolve(cached);
  }
  if (peopleInflight) return peopleInflight;
  return fetchPeopleData({ ...options, force: false }).catch(() => null);
}

/**
 * Statutory Autofill: return cached or first Zoho page immediately (<1s target),
 * then refresh the full list in the background for pagination.
 */
export async function fetchPeopleDataForAutofillDisplay(options = {}) {
  const cached = readPeopleCacheRaw();
  if (cached?.success) {
    const count = flattenZohoPeopleEmployees(cached).length;
    if (count > 0) return cached;
  }
  if (peopleInflight) {
    try {
      return await peopleInflight;
    } catch (_) {
      /* fall through to first-page fetch */
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
      startPeopleDataBackgroundRefresh(options);
      return {
        success: true,
        data: { response: { result: batch, status: 0 } },
        meta: { total: batch.length, mode: 'fast_first_page' },
      };
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

/** During statutory autofill, use cache only — never await full all_salaries download. */
export function getPayrollBulkRowsForAutofill() {
  const cached = readPayrollBulkCacheRaw();
  return Array.isArray(cached) && cached.length > 0 ? cached : null;
}

/** Let the browser paint and process scroll/input before heavy autofill work continues. */
export function yieldToMain() {
  return new Promise((resolve) => {
    requestAnimationFrame(() => {
      requestAnimationFrame(resolve);
    });
  });
}

/** Run a synchronous per-item handler in chunks so scroll/input stay responsive. */
export async function processInChunks(items, chunkSize, handler) {
  const list = Array.isArray(items) ? items : [];
  const size = Math.max(1, chunkSize || 1);
  for (let i = 0; i < list.length; i += 1) {
    handler(list[i], i);
    if ((i + 1) % size === 0) {
      await yieldToMain();
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
  if (!force && attendanceInflight.has(inflightKey)) {
    return attendanceInflight.get(inflightKey);
  }

  const timeoutMs = Math.max(30000, parseInt(options.timeoutMs, 10) || 60000);
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
  if (!force && leaveInflight.has(inflightKey)) {
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

const PAYROLL_SALARY_BATCH_SIZE = 6;

async function fetchPayrollBulkRowsBatched(organizationId) {
  const merged = [];
  let offset = 0;

  while (true) {
    const qs = new URLSearchParams({
      all_salaries: '1',
      organization_id: organizationId,
      salary_offset: String(offset),
      salary_limit: String(PAYROLL_SALARY_BATCH_SIZE),
    });
    const resp = await fetch(`/server/payroll_function?${qs.toString()}`, { cache: 'no-store' });
    const json = await resp.json().catch(() => ({}));
    if (!resp.ok || !json?.success || !Array.isArray(json.data)) {
      break;
    }
    const batch = json.data;
    if (batch.length === 0) break;
    merged.push(...batch);
    const hasMore =
      json.meta?.has_more === true ||
      (json.meta?.has_more !== false && batch.length >= PAYROLL_SALARY_BATCH_SIZE);
    if (!hasMore) break;
    offset += batch.length;
  }

  return merged;
}

export function fetchPayrollBulkRows(options = {}) {
  const { force = false } = options;
  const cached = readPayrollBulkCacheRaw();
  if (!force && cached) return Promise.resolve(cached);
  if (!force && payrollBulkInflight) return payrollBulkInflight;

  const payrollOrgId = getPayrollOrganizationId();

  payrollBulkInflight = fetchPayrollBulkRowsBatched(payrollOrgId)
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
  if (options?.clraIndustry === true) {
    prefetchClraData();
  } else {
    prefetchPeopleData();
  }
  prefetchAttendanceData();
  prefetchLeaveData();
  // Do not prefetch all_salaries here — it blocks payroll_function for minutes on large orgs.
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

export function readFormTemplateCache(cacheKey) {
  if (!cacheKey) return null;
  const entry = formTemplateCache.get(cacheKey);
  if (entry && Date.now() - entry.ts < FORM_TEMPLATE_TTL_MS) return entry;
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
        const resp = await fetch(url);
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
