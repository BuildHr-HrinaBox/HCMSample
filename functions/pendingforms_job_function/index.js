'use strict';

const catalystSDK = require('zcatalyst-sdk-node');
const { sendPendingFormsDigest } = require('./pendingFormsMailer');

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
 * Monthly cron (schedule in Catalyst Console for day 20, Asia/Kolkata).
 * Emails each Site Management In-Charge the pending statutory forms for their site.
 *
 * Cron params:
 *   force=true   send even if today is not the 20th
 *   dryRun=true  compile only; do not send mail
 *
 * @param {import('./types/cron').CronDetails} cronDetails
 * @param {import('./types/cron').Context} context
 */
module.exports = async (cronDetails, context) => {
  const force = isTruthyParam(readCronParam(cronDetails, 'force'));
  const dryRun = isTruthyParam(readCronParam(cronDetails, 'dryRun'));

  console.log('pendingforms_job_function: starting site-wise pending forms digest', { force, dryRun });

  try {
    const catalyst = catalystSDK.initialize(context);
    const result = await sendPendingFormsDigest(catalyst, {
      force,
      dryRun,
      requireMonthlySendDay: true
    });

    console.log('pendingforms_job_function: completed', {
      skipped: result.skipped,
      reason: result.reason || null,
      periodLabel: result.periodLabel,
      sitesConsidered: result.sitesConsidered,
      emailsSent: result.emailsSent,
      emailsFailed: result.emailsFailed,
      sitesSkippedNoEmail: result.sitesSkippedNoEmail,
      sitesSkippedNoPending: result.sitesSkippedNoPending
    });

    if (result.skipped) {
      if (result.reason && /mail sender|from_email|Mail →|MAIL_FROM/i.test(result.reason)) {
        console.error('pendingforms_job_function: mail configuration missing', result.reason);
        context.closeWithFailure();
        return;
      }
      context.closeWithSuccess();
      return;
    }

    if (result.emailsSent > 0) {
      context.closeWithSuccess();
      return;
    }

    console.error(
      'pendingforms_job_function: no emails sent',
      result.emailsFailed ? 'send failed' : 'no pending forms with In-Charge email'
    );
    context.closeWithFailure();
  } catch (err) {
    console.error('pendingforms_job_function error:', err && err.message ? err.message : err);
    context.closeWithFailure();
  }
};
