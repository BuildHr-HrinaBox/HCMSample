const API_BASE = '/server/leavedata_function';

const MONTH_ABBR = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

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

/**
 * Fetch leave booked/balance report from Zoho People via leavedata_function.
 * Token on server: ZOHOPEOPLE.leave.ALL (api_domain https://www.zohoapis.in).
 */
export async function fetchLeaveReport({
  from = DEFAULT_LEAVE_REPORT_FROM,
  to = DEFAULT_LEAVE_REPORT_TO,
  unit = 'Day',
  fetchAll = true,
  timeoutMs = 300000,
} = {}) {
  const fromErr = validateZohoLeaveDate(from, 'From date');
  if (fromErr) throw new Error(fromErr);
  const toErr = validateZohoLeaveDate(to, 'To date');
  if (toErr) throw new Error(toErr);

  const params = new URLSearchParams({ from, to, unit, fetch_all: fetchAll ? '1' : '0' });
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);

  try {
    const res = await fetch(`${API_BASE}?${params.toString()}`, {
      signal: controller.signal,
    });
    const json = await res.json();
    if (!res.ok) {
      throw new Error(json.error || json.message || 'Request failed');
    }
    if (!json || typeof json !== 'object') {
      throw new Error('Invalid response');
    }

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
  } catch (err) {
    if (err.name === 'AbortError') {
      throw new Error('Leave request timed out');
    }
    throw err;
  } finally {
    clearTimeout(timer);
  }
}
