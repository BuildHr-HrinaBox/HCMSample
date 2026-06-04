const SITE_DETAILS_CACHE_KEY = 'statutorySiteDetails_v1';
const PEOPLE_CACHE_KEY = 'statutoryPeopleData_v1';
const PEOPLE_CACHE_TTL_MS = 5 * 60 * 1000;
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

export function fetchPeopleData(options = {}) {
  const { force = false } = options;
  const cached = readPeopleCacheRaw();
  if (!force && cached) return Promise.resolve(cached);
  if (!force && peopleInflight) return peopleInflight;

  peopleInflight = fetch('/server/peopledata_function?form=employee&limit=100', { cache: 'no-store' })
    .then(async (response) => {
      const result = await response.json().catch(() => ({}));
      if (!response.ok || !result.success) {
        throw new Error(result.error || 'Failed to load employee data');
      }
      writePeopleCache(result);
      return result;
    })
    .finally(() => {
      peopleInflight = null;
    });

  return peopleInflight;
}

export function prefetchPeopleData() {
  if (readPeopleCacheRaw()) return Promise.resolve(readPeopleCacheRaw());
  return fetchPeopleData().catch(() => null);
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

export function fetchPayrollBulkRows(options = {}) {
  const { force = false } = options;
  const cached = readPayrollBulkCacheRaw();
  if (!force && cached) return Promise.resolve(cached);
  if (!force && payrollBulkInflight) return payrollBulkInflight;

  const payrollOrgId = process.env.REACT_APP_ZOHO_PAYROLL_ORGANIZATION_ID || '60006183023';
  const allPayrollQs = new URLSearchParams({
    all_salaries: '1',
    organization_id: payrollOrgId,
  });

  payrollBulkInflight = fetch(`/server/payroll_function?${allPayrollQs.toString()}`, { cache: 'no-store' })
    .then(async (resp) => {
      const json = await resp.json().catch(() => ({}));
      if (resp.ok && json?.success && Array.isArray(json.data) && json.data.length > 0) {
        writePayrollBulkCache(json.data);
        return json.data;
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

export function prefetchStatutoryAutofillData() {
  prefetchSiteDetails();
  prefetchPeopleData();
  prefetchPayrollBulkRows();
}

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
