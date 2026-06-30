const API_BASE = '/server/leavedata_function';
export const NEW_LEAVE_API_BASE = '/server/newleave_function';

/** Zoho bookedAndBalance returns up to 100 employees per request (.in DC). */
export const LEAVE_PAGE_SIZE = 100;
/** Zoho bookedAndBalance threshold is low (~20–30 req/min); pace client pagination. */
const LEAVE_PAGE_DELAY_MS = 2500;
const LEAVE_CACHE_TTL_MS = 5 * 60 * 1000;
const LEAVE_RATE_LIMIT_LOCK_MS = 5 * 60 * 1000;
const LEAVE_CACHE_KEY_PREFIX = 'leave_report_cache:';
const LEAVE_RATE_LIMIT_KEY = 'zoho_leave_rate_limit_until';
const LEAVE_RATE_LIMIT_PAGE_RETRIES = 2;
const LEAVE_RATE_LIMIT_RETRY_MS = 90000;

const MONTH_ABBR = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function isRateLimitError(message) {
  return /7209|threshold limit|rate limit exceeded/i.test(String(message || ''));
}

function getRateLimitLockUntil() {
  try {
    const value = sessionStorage.getItem(LEAVE_RATE_LIMIT_KEY);
    return value ? parseInt(value, 10) || 0 : 0;
  } catch (_) {
    return 0;
  }
}

function setRateLimitLock() {
  try {
    sessionStorage.setItem(LEAVE_RATE_LIMIT_KEY, String(Date.now() + LEAVE_RATE_LIMIT_LOCK_MS));
  } catch (_) {
    /* ignore */
  }
}

function assertNotRateLimited() {
  const until = getRateLimitLockUntil();
  if (until > Date.now()) {
    const minutes = Math.max(1, Math.ceil((until - Date.now()) / 60000));
    throw new Error(
      `Zoho People API is temporarily locked (rate limit). Please wait about ${minutes} minute(s) before fetching again.`
    );
  }
}

function readLeaveCache(cacheKey) {
  try {
    const raw = sessionStorage.getItem(cacheKey);
    if (!raw) return null;
    const parsed = JSON.parse(raw);
    if (!parsed || Date.now() - parsed.ts > LEAVE_CACHE_TTL_MS) return null;
    return parsed.data || null;
  } catch (_) {
    return null;
  }
}

function writeLeaveCache(cacheKey, data) {
  try {
    sessionStorage.setItem(cacheKey, JSON.stringify({ ts: Date.now(), data }));
  } catch (_) {
    /* ignore */
  }
}

function leaveEmployeeKey(row) {
  if (!row || typeof row !== 'object') return '';
  const emp = row.employee || row.Employee;
  if (emp && typeof emp === 'object') {
    return String(emp.id || emp.erecno || emp.zoho_id || '').trim();
  }
  return String(row.employeeId || row['Employee.ID'] || row['Zoho.ID'] || row.id || '').trim();
}

function normalizeLeaveApiResponse(json) {
  const recordsFromMap =
    json.records && typeof json.records === 'object' && !Array.isArray(json.records)
      ? Object.entries(json.records).map(([id, row]) => ({
          employeeId: String(id),
          ...(row && typeof row === 'object' ? row : {}),
        }))
      : null;

  return {
    success: json.success !== false,
    raw: json.records ?? json.data ?? json,
    leaveTypeLabels: json.leaveTypeLabels || {},
    leaveRecords: recordsFromMap || (Array.isArray(json.leaveRecords) ? json.leaveRecords : null),
    meta: json.meta || null,
  };
}

/**
 * Zoho bookedAndBalance is leave-year sensitive. This org returns only empty Absent
 * data for calendar-year windows (e.g. Jan–Dec) but full balances on Apr–Mar FY.
 */
export function getDefaultLeaveReportRange(referenceDate = new Date()) {
  const now = referenceDate instanceof Date ? referenceDate : new Date();
  const month = now.getMonth();
  const year = now.getFullYear();
  const fyStartYear = month >= 3 ? year : year - 1;
  const fyEndYear = fyStartYear + 1;
  return {
    from: `01-Apr-${fyStartYear}`,
    to: `31-Mar-${fyEndYear}`,
  };
}

const defaultRange = getDefaultLeaveReportRange();
export const DEFAULT_LEAVE_REPORT_FROM = defaultRange.from;
export const DEFAULT_LEAVE_REPORT_TO = defaultRange.to;

/** Validate Zoho bookedAndBalance date (DD-Mon-YYYY). Returns error message or empty string. */
export function validateZohoLeaveDate(dateStr, label = 'Date') {
  const m = String(dateStr || '')
    .trim()
    .match(/^(\d{1,2})-([A-Za-z]{3})-(\d{4})$/);
  if (!m) return `${label} must be DD-Mon-YYYY (e.g. 30-Jun-2026).`;
  const day = parseInt(m[1], 10);
  const mon = MONTH_ABBR.findIndex((x) => x.toLowerCase() === m[2].toLowerCase());
  const year = parseInt(m[3], 10);
  if (mon < 0) return `${label} has invalid month "${m[2]}".`;
  const d = new Date(year, mon, day);
  if (d.getFullYear() !== year || d.getMonth() !== mon || d.getDate() !== day) {
    return `${label} "${dateStr}" is not a valid calendar date (e.g. June has 30 days).`;
  }
  return '';
}

async function fetchLeavePageRequest({
  from,
  to,
  unit,
  fetchAll,
  startIndex = 0,
  timeoutMs,
  apiBase,
}) {
  const params = new URLSearchParams({ from, to, unit, fetch_all: fetchAll ? '1' : '0' });
  if (!fetchAll) {
    params.set('startIndex', String(startIndex));
    params.set('limit', String(LEAVE_PAGE_SIZE));
  }

  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);

  try {
    const res = await fetch(`${apiBase}?${params.toString()}`, {
      cache: 'no-store',
      credentials: 'include',
      signal: controller.signal,
    });
    const text = await res.text();
    let json;
    try {
      json = text ? JSON.parse(text) : {};
    } catch (_) {
      const snippet = String(text || '').trim().slice(0, 240);
      const hint =
        /cannot find module|function not found|not deployed/i.test(snippet)
          ? ' Deploy newleave_function: catalyst deploy --only functions:newleave_function'
          : '';
      throw new Error(
        snippet
          ? `Leave server returned HTTP ${res.status} (not JSON): ${snippet}${hint}`
          : `Leave server returned HTTP ${res.status} with an empty response`
      );
    }
    if (!res.ok) {
      const msg = json.error || json.message || `Request failed (HTTP ${res.status})`;
      if (isRateLimitError(msg)) {
        setRateLimitLock();
      }
      if (res.status === 408) {
        throw new Error(
          'Leave request timed out on the server (HTTP 408). Data is being loaded page by page — please wait or try again.'
        );
      }
      throw new Error(msg);
    }
    if (!json || typeof json !== 'object') {
      throw new Error('Invalid response');
    }
    return normalizeLeaveApiResponse(json);
  } catch (err) {
    if (err.name === 'AbortError') {
      throw new Error('Leave request timed out');
    }
    throw err;
  } finally {
    clearTimeout(timer);
  }
}

function mergeLeavePageResults(accumulated, page) {
  const leaveTypeLabels = {
    ...(accumulated.leaveTypeLabels || {}),
    ...(page.leaveTypeLabels || {}),
  };

  const raw =
    accumulated.raw && typeof accumulated.raw === 'object' && !Array.isArray(accumulated.raw)
      ? { ...accumulated.raw }
      : {};
  if (page.raw && typeof page.raw === 'object' && !Array.isArray(page.raw)) {
    Object.assign(raw, page.raw);
  }

  const leaveRecords = [...(accumulated.leaveRecords || [])];
  if (Array.isArray(page.leaveRecords)) {
    leaveRecords.push(...page.leaveRecords);
  }

  return {
    success: true,
    raw,
    leaveTypeLabels,
    leaveRecords,
  };
}

async function fetchLeavePageWithRetry(params) {
  for (let attempt = 0; attempt <= LEAVE_RATE_LIMIT_PAGE_RETRIES; attempt += 1) {
    try {
      return await fetchLeavePageRequest(params);
    } catch (err) {
      if (!isRateLimitError(err?.message) || attempt >= LEAVE_RATE_LIMIT_PAGE_RETRIES) {
        if (isRateLimitError(err?.message)) setRateLimitLock();
        throw err;
      }
      await sleep(LEAVE_RATE_LIMIT_RETRY_MS * (attempt + 1));
    }
  }
  throw new Error('Zoho People API rate limit exceeded after retries.');
}

async function fetchLeaveReportPaginated({ from, to, unit, timeoutMs, apiBase, onProgress }) {
  let accumulated = {
    success: true,
    raw: {},
    leaveTypeLabels: {},
    leaveRecords: [],
  };
  let startIndex = 0;
  let pages = 0;
  let firstPageKey = '';

  while (pages < 50) {
    const page = await fetchLeavePageWithRetry({
      from,
      to,
      unit,
      fetchAll: false,
      startIndex,
      timeoutMs,
      apiBase,
    });

    const batch = Array.isArray(page.leaveRecords) ? page.leaveRecords : [];
    const count = page.meta?.count ?? batch.length;
    const hasMore = page.meta?.has_more ?? count >= LEAVE_PAGE_SIZE;

    if (count === 0) break;

    accumulated = mergeLeavePageResults(accumulated, page);

    const pageKey = leaveEmployeeKey(batch[0]);
    if (pages === 0) {
      firstPageKey = pageKey;
    } else if (pageKey && pageKey === firstPageKey) {
      break;
    }

    pages += 1;
    onProgress?.({
      pages,
      total: accumulated.leaveRecords?.length ?? 0,
      status: hasMore ? 'loading' : 'done',
    });

    if (!hasMore) break;

    startIndex += LEAVE_PAGE_SIZE;
    await sleep(LEAVE_PAGE_DELAY_MS);
  }

  const total =
    accumulated.leaveRecords?.length ??
    (accumulated.raw && typeof accumulated.raw === 'object' ? Object.keys(accumulated.raw).length : 0);

  return {
    ...accumulated,
    meta: { total, pages, pageSize: LEAVE_PAGE_SIZE, mode: 'paginated' },
  };
}

/**
 * Fetch leave booked/balance report from Zoho People via leavedata_function / newleave_function.
 * Token on server: ZOHOPEOPLE.leave.ALL (api_domain https://www.zohoapis.in).
 */
export async function fetchLeaveReport({
  from = DEFAULT_LEAVE_REPORT_FROM,
  to = DEFAULT_LEAVE_REPORT_TO,
  unit = 'Day',
  fetchAll = true,
  timeoutMs = 300000,
  apiBase = API_BASE,
  onProgress,
  useCache = true,
} = {}) {
  const fromErr = validateZohoLeaveDate(from, 'From date');
  if (fromErr) throw new Error(fromErr);
  const toErr = validateZohoLeaveDate(to, 'To date');
  if (toErr) throw new Error(toErr);

  assertNotRateLimited();

  const cacheKey = `${LEAVE_CACHE_KEY_PREFIX}${apiBase}|${from}|${to}|${unit}`;
  if (useCache) {
    const cached = readLeaveCache(cacheKey);
    if (cached) {
      onProgress?.({
        pages: cached.meta?.pages ?? 1,
        total: cached.leaveRecords?.length ?? 0,
        status: 'cached',
      });
      return cached;
    }
  }

  let result;
  if (!fetchAll) {
    result = await fetchLeavePageWithRetry({
      from,
      to,
      unit,
      fetchAll: false,
      startIndex: 0,
      timeoutMs,
      apiBase,
    });
  } else {
    // Client-side pagination only — avoids server fetch_all bursts and Catalyst 408 timeouts.
    result = await fetchLeaveReportPaginated({ from, to, unit, timeoutMs, apiBase, onProgress });
  }

  if (useCache && result?.success !== false) {
    writeLeaveCache(cacheKey, result);
  }
  return result;
}
