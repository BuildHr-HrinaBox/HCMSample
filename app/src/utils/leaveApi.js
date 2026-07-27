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

/** YYYY-MM-DD for HTML date inputs (local calendar day). */
export function toIsoDateInput(referenceDate = new Date()) {
  const d = referenceDate instanceof Date ? referenceDate : new Date(referenceDate);
  if (Number.isNaN(d.getTime())) return '';
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${y}-${m}-${day}`;
}

/** Convert HTML date input (YYYY-MM-DD) to Zoho leave date (DD-Mon-YYYY). */
export function isoToZohoLeaveDate(isoDate) {
  const m = String(isoDate || '')
    .trim()
    .match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return '';
  const year = parseInt(m[1], 10);
  const monthIndex = parseInt(m[2], 10) - 1;
  const day = parseInt(m[3], 10);
  if (monthIndex < 0 || monthIndex > 11) return '';
  const d = new Date(year, monthIndex, day);
  if (d.getFullYear() !== year || d.getMonth() !== monthIndex || d.getDate() !== day) return '';
  return `${String(day).padStart(2, '0')}-${MONTH_ABBR[monthIndex]}-${year}`;
}

/** Display helper: YYYY-MM-DD → DD/MM/YY (e.g. 24/07/26). */
export function isoToDisplayDate(isoDate) {
  const m = String(isoDate || '')
    .trim()
    .match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return '';
  return `${m[3]}/${m[2]}/${m[1].slice(-2)}`;
}

/** Leave page default: today → today (day-wise). Returns ISO dates for date pickers. */
export function getTodayLeaveDayRange(referenceDate = new Date()) {
  const today = toIsoDateInput(referenceDate);
  return { from: today, to: today };
}

/** Leave page default: first → last day of the current calendar month (ISO for date pickers). */
export function getCurrentMonthLeaveRange(referenceDate = new Date()) {
  const d = referenceDate instanceof Date ? referenceDate : new Date(referenceDate);
  if (Number.isNaN(d.getTime())) return { from: '', to: '' };
  const year = d.getFullYear();
  const month = d.getMonth();
  const from = toIsoDateInput(new Date(year, month, 1));
  const to = toIsoDateInput(new Date(year, month + 1, 0));
  return { from, to };
}

/** Last calendar day of the month containing an ISO date (YYYY-MM-DD). */
export function getMonthEndIsoFromIso(isoDate) {
  const m = String(isoDate || '')
    .trim()
    .match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return '';
  const year = parseInt(m[1], 10);
  const monthIndex = parseInt(m[2], 10) - 1;
  if (monthIndex < 0 || monthIndex > 11) return '';
  return toIsoDateInput(new Date(year, monthIndex + 1, 0));
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

function parseJsonMaybe(val) {
  if (val == null) return null;
  if (typeof val === 'object') return val;
  if (typeof val === 'string') {
    const s = val.trim();
    if (!s.startsWith('{') && !s.startsWith('[')) return null;
    try {
      return JSON.parse(s);
    } catch {
      return null;
    }
  }
  return null;
}

function formatEmployeeForStore(val) {
  const o = parseJsonMaybe(val) ?? (val && typeof val === 'object' ? val : null);
  if (o && typeof o === 'object') {
    const name = o.name != null ? String(o.name) : '';
    const id = o.id != null ? String(o.id) : '';
    if (name && id) return `${name} (${id})`;
    return name || id || '';
  }
  return val == null ? '' : String(val);
}

function formatLeaveCellForStore(val) {
  const obj = parseJsonMaybe(val) ?? (val && typeof val === 'object' ? val : null);
  if (!obj || typeof obj !== 'object') return val == null ? '' : String(val);
  if (Object.keys(obj).length === 0) return '';
  if ('paidBalance' in obj || 'paidBooked' in obj || 'unpaidBalance' in obj || 'unpaidBooked' in obj) {
    const b = obj.paidBalance ?? obj.balance ?? obj.unpaidBalance;
    const book = obj.paidBooked ?? obj.booked ?? obj.unpaidBooked;
    const parts = [];
    if (b != null && b !== '') parts.push(`Balance: ${b}`);
    if (book != null && book !== '') parts.push(`Booked: ${book}`);
    return parts.join(', ');
  }
  if (Object.keys(obj).length === 1 && 'balance' in obj) return String(obj.balance);
  if ('balance' in obj || 'booked' in obj) {
    const parts = [];
    if (obj.balance != null && obj.balance !== '') parts.push(`Balance: ${obj.balance}`);
    if (obj.booked != null && obj.booked !== '') parts.push(`Booked: ${obj.booked}`);
    return parts.join(', ');
  }
  try {
    return JSON.stringify(obj);
  } catch {
    return String(obj);
  }
}

const EARNED_LEAVE_LABELS = new Set([
  'Earned Leave (Test)',
  'Earned Leave(Test)',
  'Earned Leave (test)',
  'Earned Leave',
  'Earned leave',
]);

function findLeaveTypeKey(row, leaveTypeLabels, matcher) {
  if (!row || typeof row !== 'object') return null;
  for (const key of Object.keys(row)) {
    const label = leaveTypeLabels && leaveTypeLabels[key] ? String(leaveTypeLabels[key]) : key;
    if (matcher(String(key), String(label))) return key;
  }
  return null;
}

function getEarnedLeaveBreakout(row, leaveTypeLabels) {
  const earnedKey = findLeaveTypeKey(row, leaveTypeLabels, (key, label) => {
    if (/legacy/i.test(key) || /legacy/i.test(label)) return false;
    return EARNED_LEAVE_LABELS.has(key) || EARNED_LEAVE_LABELS.has(label) || /^earned\s+leave(\s*\(test\))?$/i.test(label);
  });
  if (!earnedKey) return { balance: '', booked: '', earnedKey: null };
  const raw = row[earnedKey];
  const obj = parseJsonMaybe(raw) ?? (raw && typeof raw === 'object' ? raw : null);
  if (!obj || typeof obj !== 'object') return { balance: '', booked: '', earnedKey };
  const balance = obj.paidBalance ?? obj.balance ?? obj.Balance;
  const booked = obj.paidBooked ?? obj.booked ?? obj.Booked;
  return {
    balance: balance != null && balance !== '' ? String(balance) : '',
    booked: booked != null && booked !== '' ? String(booked) : '',
    earnedKey,
  };
}

function pickLeaveTypeDisplay(row, leaveTypeLabels, matcher) {
  const key = findLeaveTypeKey(row, leaveTypeLabels, matcher);
  if (!key) return '';
  return formatLeaveCellForStore(row[key]);
}

/**
 * Derive MonthWise (YYYY-MM) from Zoho leave From date (DD-Mon-YYYY).
 * Falls back to To date, then empty string.
 */
export function getLeaveMonthWise(fromDate, toDate) {
  const parse = (dateStr) => {
    const m = String(dateStr || '')
      .trim()
      .match(/^(\d{1,2})-([A-Za-z]{3})-(\d{4})$/);
    if (!m) return '';
    const mon = MONTH_ABBR.findIndex((x) => x.toLowerCase() === m[2].toLowerCase());
    if (mon < 0) return '';
    return `${m[3]}-${String(mon + 1).padStart(2, '0')}`;
  };
  return parse(fromDate) || parse(toDate) || '';
}

/**
 * Map a Zoho leave report row to LeaveData datastore columns.
 */
export function mapLeaveRecordToLeaveDataRow(row, leaveTypeLabels = {}, options = {}) {
  if (!row || typeof row !== 'object') return null;

  const employee =
    formatEmployeeForStore(row.employee ?? row.Employee) ||
    (row.employeeId != null ? String(row.employeeId) : '');
  if (!employee) return null;

  const earned = getEarnedLeaveBreakout(row, leaveTypeLabels);
  const monthWise =
    options.monthWise != null && String(options.monthWise).trim() !== ''
      ? String(options.monthWise).trim()
      : getLeaveMonthWise(options.from, options.to);

  return {
    Employee: employee,
    LeaveearnedduringthePeriod: earned.balance,
    LeaveavailedduringthePeriod: earned.booked,
    Totals: formatLeaveCellForStore(row.totals ?? row.Totals),
    LegacyEarnedLeave: pickLeaveTypeDisplay(
      row,
      leaveTypeLabels,
      (key, label) => /legacy/i.test(key) || /legacy/i.test(label)
    ),
    PaternityLeave: pickLeaveTypeDisplay(
      row,
      leaveTypeLabels,
      (key, label) => /paternity/i.test(key) || /paternity/i.test(label)
    ),
    Absent: pickLeaveTypeDisplay(
      row,
      leaveTypeLabels,
      (key, label) => /^absent$/i.test(key.trim()) || /^absent$/i.test(label.trim())
    ),
    ContingencyLeave: pickLeaveTypeDisplay(
      row,
      leaveTypeLabels,
      (key, label) => /contingency/i.test(key) || /contingency/i.test(label)
    ),
    MaternityLeave: pickLeaveTypeDisplay(
      row,
      leaveTypeLabels,
      (key, label) => /maternity/i.test(key) || /maternity/i.test(label)
    ),
    Earnedleave: earned.earnedKey ? formatLeaveCellForStore(row[earned.earnedKey]) : '',
    MonthWise: monthWise,
  };
}

/**
 * Persist mapped leave rows into Catalyst LeaveData table via leavedata_function /save.
 */
export async function saveLeaveDataToBackend({
  records,
  from,
  to,
  monthWise,
  apiBase = API_BASE,
  replaceExisting = true,
  chunkSize = 80,
  onProgress,
} = {}) {
  const resolvedMonth =
    (monthWise && String(monthWise).trim()) || getLeaveMonthWise(from, to) || null;
  const list = Array.isArray(records)
    ? records
        .filter((r) => r && r.Employee)
        .map((r) => ({
          ...r,
          MonthWise: r.MonthWise || resolvedMonth || '',
        }))
    : [];
  if (list.length === 0) {
    throw new Error('No leave records to save');
  }

  let lastJson = null;
  for (let offset = 0; offset < list.length; offset += chunkSize) {
    const chunk = list.slice(offset, offset + chunkSize);
    const isFirst = offset === 0;
    onProgress?.(
      `Saving to LeaveData… ${Math.min(offset + chunk.length, list.length)} / ${list.length}`
    );

    const res = await fetch(`${apiBase}/save`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      credentials: 'include',
      body: JSON.stringify({
        from: from || null,
        to: to || null,
        monthWise: resolvedMonth,
        records: chunk,
        replaceExisting: isFirst ? replaceExisting : false,
      }),
    });

    const text = await res.text();
    let json = {};
    try {
      json = text ? JSON.parse(text) : {};
    } catch {
      throw new Error(
        `LeaveData save returned HTTP ${res.status} (not JSON). Redeploy leavedata_function and check Catalyst logs.`
      );
    }
    if (!res.ok || json.success === false) {
      throw new Error(json.error || json.message || `Save failed (HTTP ${res.status})`);
    }
    lastJson = json;
  }

  return lastJson;
}

/**
 * Rebuild leave-report-shaped record from a LeaveData datastore row.
 * Used by Form X / Form 15 Part 1 so they can skip the live Zoho leave API.
 */
export function mapLeaveDataRowToLeaveRecord(row) {
  if (!row || typeof row !== 'object') return null;
  const employeeRaw = row.Employee ?? row.employee;
  if (employeeRaw == null || String(employeeRaw).trim() === '') {
    // Already a leave-report record (e.g. from /stored leaveRecords)
    if (row.employee && typeof row.employee === 'object') return row;
    return null;
  }

  if (row.employee && typeof row.employee === 'object' && row['Earned Leave']) {
    return row;
  }

  const empStr = String(employeeRaw).trim();
  const paren = empStr.match(/^(.*)\(([^)]+)\)\s*$/);
  const employee = paren
    ? { name: paren[1].trim(), id: String(paren[2]).trim() }
    : { name: empStr };

  const earnedBalance = row.LeaveearnedduringthePeriod;
  const earnedBooked = row.LeaveavailedduringthePeriod;
  const record = {
    employee,
    Employee: empStr,
    Totals: row.Totals ?? '',
    MonthWise: row.MonthWise ?? row.monthWise ?? '',
    'Earned Leave': {
      paidBalance:
        earnedBalance != null && String(earnedBalance).trim() !== ''
          ? String(earnedBalance).trim()
          : '',
      paidBooked:
        earnedBooked != null && String(earnedBooked).trim() !== ''
          ? String(earnedBooked).trim()
          : '',
    },
  };

  if (row.ContingencyLeave != null && String(row.ContingencyLeave).trim() !== '') {
    record['Contingency Leave'] = row.ContingencyLeave;
  }
  if (row.LegacyEarnedLeave != null && String(row.LegacyEarnedLeave).trim() !== '') {
    record['Legacy Earned Leave'] = row.LegacyEarnedLeave;
  }
  if (row.PaternityLeave != null && String(row.PaternityLeave).trim() !== '') {
    record['Paternity Leave'] = row.PaternityLeave;
  }
  if (row.MaternityLeave != null && String(row.MaternityLeave).trim() !== '') {
    record['Maternity Leave'] = row.MaternityLeave;
  }
  if (row.Absent != null && String(row.Absent).trim() !== '') {
    record.Absent = row.Absent;
  }
  if (row.Earnedleave != null && String(row.Earnedleave).trim() !== '') {
    record['Earned leave'] = row.Earnedleave;
  }

  return record;
}

/**
 * Load leave rows from Catalyst LeaveData table (no live Zoho People call).
 * Prefer this for Form 15 Part 1 / Form X autofill after Leave page has saved data.
 */
export async function fetchStoredLeaveData({
  from,
  to,
  monthWise,
  apiBase = API_BASE,
  timeoutMs = 60000,
} = {}) {
  const resolvedMonth =
    (monthWise && String(monthWise).trim()) || getLeaveMonthWise(from, to) || '';
  const params = new URLSearchParams();
  if (resolvedMonth) params.set('monthWise', resolvedMonth);

  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), Math.max(5000, timeoutMs));
  try {
    const qs = params.toString();
    const res = await fetch(`${apiBase}/stored${qs ? `?${qs}` : ''}`, {
      cache: 'no-store',
      credentials: 'include',
      signal: controller.signal,
    });
    const text = await res.text();
    let json = {};
    try {
      json = text ? JSON.parse(text) : {};
    } catch {
      throw new Error(
        `LeaveData load returned HTTP ${res.status} (not JSON). Redeploy leavedata_function and check Catalyst logs.`
      );
    }
    if (!res.ok || json.success === false) {
      throw new Error(json.error || json.message || `LeaveData load failed (HTTP ${res.status})`);
    }

    const leaveRecords = (
      Array.isArray(json.leaveRecords)
        ? json.leaveRecords
        : Array.isArray(json.records)
          ? json.records
          : []
    )
      .map((row) => mapLeaveDataRowToLeaveRecord(row) || row)
      .filter((row) => row && typeof row === 'object');

    return {
      success: true,
      leaveRecords,
      records: leaveRecords,
      leaveTypeLabels: json.leaveTypeLabels || {},
      meta: {
        ...(json.meta && typeof json.meta === 'object' ? json.meta : {}),
        source: 'LeaveData',
        monthWise: resolvedMonth || null,
      },
    };
  } catch (err) {
    if (err.name === 'AbortError') {
      throw new Error('LeaveData table request timed out');
    }
    throw err;
  } finally {
    clearTimeout(timer);
  }
}
