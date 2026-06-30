import { validateZohoLeaveDate } from './leaveApi';

const API_BASE = '/server/approve_leave_function';
const PAGE_SIZE = 200;
const PAGE_DELAY_MS = 800;

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function leaveRecordKey(row) {
  if (!row || typeof row !== 'object') return '';
  return String(
    row['Zoho.ID'] ||
      row.ZohoID ||
      row.recordId ||
      row.id ||
      row['Employee.ID'] ||
      ''
  ).trim();
}

function normalizeApprovedLeaveResponse(json) {
  const recordsFromMap =
    json.records && typeof json.records === 'object' && !Array.isArray(json.records)
      ? Object.entries(json.records).map(([id, row]) => ({
          recordId: String(id),
          ...(row && typeof row === 'object' ? row : {}),
        }))
      : null;

  return {
    success: json.success !== false,
    leaveRecords: recordsFromMap || (Array.isArray(json.leaveRecords) ? json.leaveRecords : []),
    meta: json.meta || null,
  };
}

export function getDefaultApprovedLeaveRange(referenceDate = new Date()) {
  const now = referenceDate instanceof Date ? referenceDate : new Date();
  const year = now.getFullYear();
  return {
    from: `01-Jan-${year}`,
    to: `31-Dec-${year}`,
  };
}

async function fetchApprovedLeavePage({ from, to, startIndex = 0, timeoutMs = 300000 }) {
  const params = new URLSearchParams({
    from,
    to,
    fetch_all: '0',
    startIndex: String(startIndex),
    limit: String(PAGE_SIZE),
  });

  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);

  try {
    const res = await fetch(`${API_BASE}?${params.toString()}`, {
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
          ? ' Deploy approve_leave_function: catalyst deploy --only functions:approve_leave_function'
          : '';
      throw new Error(
        snippet
          ? `Approved leave server returned HTTP ${res.status} (not JSON): ${snippet}${hint}`
          : `Approved leave server returned HTTP ${res.status} with an empty response`
      );
    }
    if (!res.ok) {
      throw new Error(json.error || json.message || `Request failed (HTTP ${res.status})`);
    }
    if (!json || typeof json !== 'object') {
      throw new Error('Invalid response');
    }
    return normalizeApprovedLeaveResponse(json);
  } catch (err) {
    if (err.name === 'AbortError') {
      throw new Error('Approved leave request timed out');
    }
    throw err;
  } finally {
    clearTimeout(timer);
  }
}

/**
 * Fetch approved leave records from Zoho People via approve_leave_function.
 */
export async function fetchApprovedLeaves({
  from,
  to,
  timeoutMs = 300000,
  onProgress,
} = {}) {
  const range = getDefaultApprovedLeaveRange();
  const fromDate = from || range.from;
  const toDate = to || range.to;

  const fromErr = validateZohoLeaveDate(fromDate, 'From date');
  if (fromErr) throw new Error(fromErr);
  const toErr = validateZohoLeaveDate(toDate, 'To date');
  if (toErr) throw new Error(toErr);

  const leaveRecords = [];
  let startIndex = 0;
  let pages = 0;
  let firstPageKey = '';

  while (pages < 200) {
    const page = await fetchApprovedLeavePage({
      from: fromDate,
      to: toDate,
      startIndex,
      timeoutMs,
    });

    const batch = Array.isArray(page.leaveRecords) ? page.leaveRecords : [];
    const count = page.meta?.count ?? batch.length;
    const hasMore = page.meta?.has_more ?? count >= PAGE_SIZE;

    if (count === 0) break;

    const pageKey = leaveRecordKey(batch[0]);
    if (pages === 0) {
      firstPageKey = pageKey;
    } else if (pageKey && pageKey === firstPageKey) {
      break;
    }

    leaveRecords.push(...batch);
    pages += 1;
    onProgress?.({
      pages,
      total: leaveRecords.length,
      status: hasMore ? 'loading' : 'done',
    });

    if (!hasMore) break;

    startIndex += PAGE_SIZE;
    await sleep(PAGE_DELAY_MS);
  }

  return {
    success: true,
    leaveRecords,
    meta: {
      total: leaveRecords.length,
      pages,
      pageSize: PAGE_SIZE,
      mode: 'paginated',
      approvalStatus: 'APPROVED',
    },
  };
}
