'use strict';

const catalystSDK = require('zcatalyst-sdk-node');
const {
  syncLeaveDataToStore,
  getDefaultLeaveReportRange,
  isValidZohoLeaveDate,
} = require('./leaveSync');

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

/**
 * Cron job: automatically fetch Zoho leave (same as Leave page Fetch Data)
 * and store into LeaveData.
 *
 * Important: leavedata_function is Advanced I/O. Catalyst functions().execute()
 * only supports Basic I/O ("execution not supported"). This job runs the shared
 * sync core directly so it can use the cron 15-minute timeout.
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

  console.log('leave_job_function: starting leave sync', {
    from,
    to,
    unit,
    monthWise: monthWiseParam || null,
  });

  try {
    const catalyst = catalystSDK.initialize(context);
    const result = await syncLeaveDataToStore(catalyst, {
      fromDate: from,
      toDate: to,
      unit,
      replaceExisting: true,
      monthWise: monthWiseParam || null,
    });

    if (!result || result.success === false) {
      const message = (result && (result.error || result.message)) || 'leave sync failed';
      console.error('leave_job_function: sync failed', message, result);
      context.closeWithFailure();
      return;
    }

    console.log('leave_job_function: sync completed', {
      from: result.from || from,
      to: result.to || to,
      monthWise: result.monthWise || null,
      fetched: result.fetched,
      inserted: result.inserted,
      deleted: result.deleted,
      saved: result.saved,
      reason: result.reason || null,
      meta: result.meta || null,
    });

    // Do not report Success when nothing was written to LeaveData.
    if (!result.saved || !(result.inserted > 0)) {
      console.error(
        'leave_job_function: no rows stored in LeaveData',
        result.reason || 'saved_false',
        {
          fetched: result.fetched,
          metaTotal: result.meta && result.meta.total,
        }
      );
      context.closeWithFailure();
      return;
    }

    context.closeWithSuccess();
  } catch (err) {
    console.error('leave_job_function error:', err && err.message ? err.message : err);
    if (err && err.response && err.response.data) {
      console.error('leave_job_function error details:', err.response.data);
    }
    context.closeWithFailure();
  }
};
