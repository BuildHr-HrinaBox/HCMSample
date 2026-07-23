'use strict';

const catalystSDK = require('zcatalyst-sdk-node');

const MONTH_ABBR = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

/**
 * Same FY window as the Leave UI (Apr–Mar). Calendar-year ranges often return empty Absent rows.
 */
function getDefaultLeaveReportRange(referenceDate = new Date()) {
  const now = referenceDate instanceof Date ? referenceDate : new Date();
  const month = now.getMonth();
  const year = now.getFullYear();
  const fyStartYear = month >= 3 ? year : year - 1;
  return {
    from: `01-Apr-${fyStartYear}`,
    to: `31-Mar-${fyStartYear + 1}`,
  };
}

function readCronParam(cronDetails, name) {
  try {
    if (typeof cronDetails.getCronParam === 'function') {
      const value = cronDetails.getCronParam(name);
      if (value != null && String(value).trim() !== '') return String(value).trim();
    }
  } catch (_) {
    /* ignore */
  }
  try {
    if (typeof cronDetails.getAllCronParam === 'function') {
      const all = cronDetails.getAllCronParam() || {};
      const value = all[name];
      if (value != null && String(value).trim() !== '') return String(value).trim();
    }
  } catch (_) {
    /* ignore */
  }
  return '';
}

function isValidZohoLeaveDate(dateStr) {
  const m = String(dateStr || '')
    .trim()
    .match(/^(\d{1,2})-([A-Za-z]{3})-(\d{4})$/);
  if (!m) return false;
  const day = parseInt(m[1], 10);
  const mon = MONTH_ABBR.findIndex((x) => x.toLowerCase() === m[2].toLowerCase());
  const year = parseInt(m[3], 10);
  if (mon < 0) return false;
  const d = new Date(year, mon, day);
  return d.getFullYear() === year && d.getMonth() === mon && d.getDate() === day;
}

function unwrapExecuteResult(raw) {
  if (raw == null) return null;
  if (typeof raw === 'string') {
    try {
      return JSON.parse(raw);
    } catch (_) {
      return { success: false, error: raw };
    }
  }
  if (typeof raw !== 'object') return { success: false, error: String(raw) };
  // Catalyst may wrap Advanced I/O JSON as { output } / { data } / body string.
  if (raw.output != null && (typeof raw.output === 'string' || typeof raw.output === 'object')) {
    return unwrapExecuteResult(raw.output);
  }
  if (raw.data != null && typeof raw.data === 'object' && raw.success == null && raw.synced == null) {
    return unwrapExecuteResult(raw.data);
  }
  if (typeof raw.body === 'string') {
    return unwrapExecuteResult(raw.body);
  }
  return raw;
}

/**
 * Cron job: automatically fetch Zoho leave (same as Leave page Fetch Data)
 * via leavedata_function sync and store into LeaveData.
 *
 * @param {import('./types/cron').CronDetails} cronDetails
 * @param {import('./types/cron').Context} context
 */
module.exports = async (cronDetails, context) => {
  const defaults = getDefaultLeaveReportRange();
  const fromParam = readCronParam(cronDetails, 'from');
  const toParam = readCronParam(cronDetails, 'to');
  const unitParam = readCronParam(cronDetails, 'unit');
  const monthWiseParam = readCronParam(cronDetails, 'monthWise');

  const from = isValidZohoLeaveDate(fromParam) ? fromParam : defaults.from;
  const to = isValidZohoLeaveDate(toParam) ? toParam : defaults.to;
  const unit = unitParam || 'Day';

  console.log('leave_job_function: starting leave sync', { from, to, unit, monthWise: monthWiseParam || null });

  try {
    const catalyst = catalystSDK.initialize(context);
    const args = {
      sync: '1',
      from,
      to,
      unit,
      fetch_all: '1',
    };
    if (monthWiseParam) args.monthWise = monthWiseParam;

    // Prefer POST sync body (works with functions.execute). Fallback to GET ?sync=1.
    let rawResult;
    try {
      rawResult = await catalyst.functions().execute('leavedata_function', {
        method: 'POST',
        data: {
          ...args,
          path: '/sync',
        },
      });
    } catch (postErr) {
      console.warn(
        'leave_job_function: POST sync failed, retrying GET ?sync=1',
        postErr && postErr.message ? postErr.message : postErr
      );
      rawResult = await catalyst.functions().execute('leavedata_function', {
        args,
        method: 'GET',
      });
    }

    const result = unwrapExecuteResult(rawResult);

    if (!result || result.success === false) {
      const message = (result && (result.error || result.message)) || 'leavedata_function sync failed';
      console.error('leave_job_function: sync failed', message, result);
      context.closeWithFailure();
      return;
    }

    const data = result.data || result;
    console.log('leave_job_function: sync completed', {
      from: data.from || from,
      to: data.to || to,
      monthWise: data.monthWise || null,
      fetched: data.fetched,
      inserted: data.inserted,
      deleted: data.deleted,
      saved: data.saved,
      meta: data.meta || result.meta || null,
    });

    context.closeWithSuccess();
  } catch (err) {
    console.error('leave_job_function error:', err && err.message ? err.message : err);
    if (err && err.response && err.response.data) {
      console.error('leave_job_function error details:', err.response.data);
    }
    context.closeWithFailure();
  }
};
