/** Zoho People getUserReport returns at most 100 employees per request. */
export const ATTENDANCE_PAGE_SIZE = 100;
/** Short gap between Zoho pages — keeps Form Q / Form 25 / muster autofill responsive. */
const ATTENDANCE_PAGE_DELAY_MS = 80;
const API_BASE = '/server/attendance_function';

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function attendanceEmployeeKey(record) {
  if (!record || typeof record !== 'object') return '';
  const emp = record.employeeDetails;
  const parsedEmp = typeof emp === 'string' ? tryParseJson(emp) : emp;
  const meta = parsedEmp && typeof parsedEmp === 'object' ? parsedEmp : record;
  return String(
    meta.erecno ||
      meta.id ||
      meta['mail id'] ||
      meta.emailId ||
      meta.email ||
      ''
  ).trim();
}

function tryParseJson(val) {
  if (!val || typeof val !== 'object') {
    if (typeof val !== 'string') return null;
    try {
      return JSON.parse(val);
    } catch (_) {
      return null;
    }
  }
  return val;
}

/** Flatten one Zoho attendance page into employee rows. */
export function flattenAttendanceRecords(apiResult) {
  if (!apiResult) return [];
  const data = apiResult.data !== undefined ? apiResult.data : apiResult;
  if (Array.isArray(data)) return data;

  const candidates = [
    data?.result,
    data?.response?.result,
    data?.response?.record,
    data?.response?.records,
    data?.records,
    data?.record,
  ];

  for (const candidate of candidates) {
    if (Array.isArray(candidate)) return candidate;
  }

  if (data && typeof data === 'object') {
    const dateKeyRegex = /^\d{4}-\d{2}-\d{2}$/;
    if (Object.keys(data).some((k) => dateKeyRegex.test(k))) {
      return [data];
    }
  }

  return [];
}

/** Normalize Zoho attendance date keys (yyyy-mm-dd, dd-MMM-yyyy, etc.) to yyyy-mm-dd. */
export function normalizeAttendanceDateKey(key) {
  const s = String(key || '').trim();
  if (!s || s === '-') return '';
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  const d = new Date(s);
  if (!Number.isNaN(d.getTime())) {
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    return `${y}-${m}-${day}`;
  }
  return '';
}

/** Expand Zoho getUserReport rows into one record per employee per day (for statutory autofill). */
export function expandAttendanceToDailyRows(apiResult) {
  const employees = flattenAttendanceRecords(apiResult);
  const rows = [];

  employees.forEach((rec) => {
    if (!rec || typeof rec !== 'object') return;

    const empMeta =
      tryParseJson(rec.employeeDetails) ||
      tryParseJson(rec.EmployeeDetails) ||
      rec.employeeDetails ||
      rec.EmployeeDetails ||
      null;
    const attRaw =
      tryParseJson(rec.attendanceDetails) ||
      tryParseJson(rec.AttendanceDetails) ||
      rec.attendanceDetails ||
      rec.AttendanceDetails ||
      null;

    const mergeMeta = (date, dayVal) => {
      const base = dayVal && typeof dayVal === 'object' && !Array.isArray(dayVal) ? dayVal : {};
      const meta = empMeta && typeof empMeta === 'object' ? empMeta : {};
      return { date, ...meta, ...base };
    };

    if (attRaw && typeof attRaw === 'object' && !Array.isArray(attRaw)) {
      Object.keys(attRaw).forEach((rawDate) => {
        const date = normalizeAttendanceDateKey(rawDate);
        if (!date) return;
        const val = attRaw[rawDate];
        if (Array.isArray(val)) {
          val.forEach((item) => {
            if (item && typeof item === 'object') rows.push(mergeMeta(date, item));
          });
          return;
        }
        if (val && typeof val === 'object') rows.push(mergeMeta(date, val));
      });
      return;
    }

    Object.keys(rec).forEach((rawKey) => {
      const date = normalizeAttendanceDateKey(rawKey);
      if (!date) return;
      const val = rec[rawKey];
      if (val && typeof val === 'object' && !Array.isArray(val)) rows.push(mergeMeta(date, val));
    });
  });

  return rows;
}

function wrapAttendanceResponse(records) {
  return {
    result: records,
    message: 'Success',
    status: 0,
  };
}

async function fetchAttendancePage(sdate, edate, startIndex, timeoutMs) {
  const qs = new URLSearchParams({
    fetch_all: '0',
    startIndex: String(startIndex),
    sdate,
    edate,
  });
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);
  try {
    const resp = await fetch(`${API_BASE}?${qs.toString()}`, {
      cache: 'no-store',
      credentials: 'include',
      signal: controller.signal,
    });
    const text = await resp.text();
    let json = {};
    try {
      json = text ? JSON.parse(text) : {};
    } catch (_) {
      throw new Error(`Attendance server returned HTTP ${resp.status} (not JSON).`);
    }
    return { resp, json };
  } catch (err) {
    if (err?.name === 'AbortError') {
      const e = new Error(`Attendance request timed out after ${timeoutMs}ms`);
      e.code = 'TIMEOUT';
      throw e;
    }
    throw err;
  } finally {
    clearTimeout(timer);
  }
}

export async function fetchAttendancePaginated(sdate, edate, timeoutMs = 60000) {
  const merged = [];
  let startIndex = 0;
  let pages = 0;
  let firstPageKey = '';

  while (pages < 50) {
    const { resp, json } = await fetchAttendancePage(sdate, edate, startIndex, timeoutMs);
    if (!resp.ok || !json.success) {
      if (merged.length > 0 && resp.status === 400) break;
      throw new Error(json.error || 'Failed to load attendance data');
    }

    const batch = flattenAttendanceRecords(json);
    if (batch.length === 0) break;

    const pageKey = attendanceEmployeeKey(batch[0]);
    if (pages === 0) {
      firstPageKey = pageKey;
    } else if (pageKey && pageKey === firstPageKey) {
      // Old server build ignores startIndex and repeats page 1.
      break;
    }

    merged.push(...batch);
    pages += 1;

    const hasMore = json.meta?.has_more ?? batch.length >= ATTENDANCE_PAGE_SIZE;
    if (!hasMore) break;

    startIndex += ATTENDANCE_PAGE_SIZE;
    await sleep(ATTENDANCE_PAGE_DELAY_MS);
  }

  if (merged.length === 0) {
    throw new Error('No attendance data received from Zoho People API');
  }

  return {
    success: true,
    data: wrapAttendanceResponse(merged),
    meta: { total: merged.length, pages, pageSize: ATTENDANCE_PAGE_SIZE, mode: 'paginated' },
  };
}

export async function fetchAttendanceAll(sdate, edate, timeoutMs = 60000) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);
  let fetchAllResult = null;

  try {
    const qs = new URLSearchParams({
      fetch_all: '1',
      sdate,
      edate,
    });
    const resp = await fetch(`${API_BASE}?${qs.toString()}`, {
      cache: 'no-store',
      credentials: 'include',
      signal: controller.signal,
    });
    const text = await resp.text();
    let json = {};
    try {
      json = text ? JSON.parse(text) : {};
    } catch (_) {
      throw new Error(`Attendance server returned HTTP ${resp.status} (not JSON).`);
    }

    if (resp.ok && json.success && json.data) {
      fetchAllResult = json;
      const records = flattenAttendanceRecords(json);
      const pages = Number(json.meta?.pages) || 0;
      const total = Number(json.meta?.total) || records.length;

      if (records.length > ATTENDANCE_PAGE_SIZE || pages > 1 || total > ATTENDANCE_PAGE_SIZE) {
        return {
          success: true,
          data: wrapAttendanceResponse(records),
          meta: { ...(json.meta || {}), total: records.length, mode: 'fetch_all' },
        };
      }

      if (records.length < ATTENDANCE_PAGE_SIZE) {
        return {
          success: true,
          data: wrapAttendanceResponse(records),
          meta: { ...(json.meta || {}), total: records.length, mode: 'fetch_all' },
        };
      }
    }
  } catch (fetchAllErr) {
    if (fetchAllErr?.code !== 'TIMEOUT') {
      console.warn('Attendance fetch_all failed, using paginated fetch:', fetchAllErr?.message || fetchAllErr);
    } else {
      throw fetchAllErr;
    }
  } finally {
    clearTimeout(timer);
  }

  const paginated = await fetchAttendancePaginated(sdate, edate, timeoutMs);
  if (fetchAllResult) {
    const fetchAllCount = flattenAttendanceRecords(fetchAllResult).length;
    if (paginated.meta.total <= fetchAllCount) {
      const records = flattenAttendanceRecords(fetchAllResult);
      return {
        success: true,
        data: wrapAttendanceResponse(records),
        meta: { ...(fetchAllResult.meta || {}), total: records.length, mode: 'fetch_all' },
      };
    }
  }
  return paginated;
}
