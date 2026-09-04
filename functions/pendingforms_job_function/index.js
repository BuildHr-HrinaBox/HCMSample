'use strict';

const catalystSDK = require('zcatalyst-sdk-node');
const { sendApprovedSitesReportToChro } = require('./approvedSitesMailer');

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
 * Monthly cron — schedule this job daily in Catalyst Console (Asia/Kolkata).
 * - 28th: emails CHRO Notification the Approved Sites Report
 *
 * Pending forms digest is handled by pending_function using Settings → Pending Notification dates.
 *
 * Cron params:
 *   force=true   send even if today is not the 28th
 *   dryRun=true  compile only; do not send mail
 *
 * @param {import('./types/cron').CronDetails} cronDetails
 * @param {import('./types/cron').Context} context
 */
module.exports = async (cronDetails, context) => {
  const force = isTruthyParam(readCronParam(cronDetails, 'force'));
  const dryRun = isTruthyParam(readCronParam(cronDetails, 'dryRun'));

  console.log('pendingforms_job_function: starting CHRO approved-sites report', { force, dryRun });

  try {
    const catalyst = catalystSDK.initialize(context);

    let chroResult = null;
    try {
      chroResult = await sendApprovedSitesReportToChro(catalyst, {
        force,
        dryRun,
        requireMonthlySendDay: true
      });
      console.log('pendingforms_job_function: CHRO approved-sites report', {
        skipped: chroResult.skipped,
        reason: chroResult.reason || null,
        emailsSent: chroResult.emailsSent,
        emailsFailed: chroResult.emailsFailed,
        toEmails: chroResult.toEmails || []
      });
    } catch (chroErr) {
      console.error('pendingforms_job_function: CHRO approved-sites report failed', chroErr?.message || chroErr);
      context.closeWithFailure();
      return;
    }

    if (chroResult?.emailsFailed) {
      console.error('pendingforms_job_function: CHRO approved-sites report send failed');
      context.closeWithFailure();
      return;
    }

    if (chroResult?.skipped) {
      if (chroResult.reason && /mail sender|from_email|Mail →|MAIL_FROM/i.test(chroResult.reason)) {
        console.error('pendingforms_job_function: mail configuration missing', chroResult.reason);
        context.closeWithFailure();
        return;
      }
      context.closeWithSuccess();
      return;
    }

    if (chroResult && chroResult.emailsSent > 0) {
      context.closeWithSuccess();
      return;
    }

    console.error('pendingforms_job_function: CHRO approved-sites report was not sent', chroResult?.reason || '');
    context.closeWithFailure();
  } catch (err) {
    console.error('pendingforms_job_function error:', err && err.message ? err.message : err);
    context.closeWithFailure();
  }
};
