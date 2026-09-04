'use strict';

const catalystSDK = require('zcatalyst-sdk-node');
const { sendApprovedSitesReportToChro, CHRO_MONTHLY_SEND_DAY } = require('./approvedSitesMailer');

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
    const getter = cronDetails.getAllCronParam || cronDetails.getAllCronParams;
    if (typeof getter === 'function') {
      const all = getter.call(cronDetails) || {};
      const value = all[name];
      if (value != null && String(value).trim() !== '') return String(value).trim();
    }
  } catch (_) {
    /* ignore */
  }
  return '';
}

function isTruthyParam(value) {
  return /^(1|true|yes|y)$/i.test(String(value || '').trim());
}

/**
 * CHRO cron — point the Catalyst "CHRO" cron at this function.
 * Sends the Approved Sites Report to the Mail Id stored in the CHRO table.
 * Schedule monthly on the 28th (Asia/Kolkata). Execute Now also sends immediately.
 *
 * Cron params:
 *   dryRun=true  compile only; do not send mail
 */
module.exports = async (cronDetails, context) => {
  const dryRun = isTruthyParam(readCronParam(cronDetails, 'dryRun'));

  console.log('chro_job_function: starting CHRO approved-sites report', {
    sendDay: CHRO_MONTHLY_SEND_DAY,
    dryRun
  });

  try {
    const catalyst = catalystSDK.initialize(context);
    const result = await sendApprovedSitesReportToChro(catalyst, {
      force: true,
      dryRun,
      requireMonthlySendDay: false
    });

    console.log('chro_job_function: completed', {
      skipped: result.skipped,
      reason: result.reason || null,
      periodLabel: result.periodLabel,
      emailsSent: result.emailsSent,
      emailsFailed: result.emailsFailed,
      toEmails: result.toEmails || []
    });

    if (result.emailsSent > 0) {
      context.closeWithSuccess();
      return;
    }

    console.error(
      'chro_job_function: CHRO mail was not sent',
      result.reason || result.error || 'no mail id in CHRO table'
    );
    context.closeWithFailure();
  } catch (err) {
    console.error('chro_job_function error:', err && err.message ? err.message : err);
    context.closeWithFailure();
  }
};
