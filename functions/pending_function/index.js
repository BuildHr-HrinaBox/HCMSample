'use strict';

const catalystSDK = require('zcatalyst-sdk-node');
const { sendPendingFormsToRecipients, nowInKolkata } = require('./pendingFormsMailer');
const { fetchAllPendingNotificationRows, rowsDueToday } = require('./pendingNotificationConfig');

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

function uniqueEmails(rows) {
  const seen = new Set();
  const list = [];
  for (const row of rows || []) {
    const email = String(row?.email || '')
      .trim()
      .toLowerCase();
    if (!email || !email.includes('@') || seen.has(email)) continue;
    seen.add(email);
    list.push(email);
  }
  return list;
}

/**
 * Daily cron — sends pending statutory forms digest when today matches
 * a Notification Date from Settings → Pending Notification (Asia/Kolkata).
 * Email is sent ONLY to the configured Pending Notification address(es).
 * Site In-Charge / Site Management emails are never contacted.
 *
 * Cron params:
 *   force=true   send even if today does not match any configured date
 *   dryRun=true  compile only; do not send mail
 *
 * @param {import('./types/cron').CronDetails} cronDetails
 * @param {import('./types/cron').Context} context
 */
module.exports = async (cronDetails, context) => {
  const force = isTruthyParam(readCronParam(cronDetails, 'force'));
  const dryRun = isTruthyParam(readCronParam(cronDetails, 'dryRun'));
  const now = nowInKolkata();

  console.log('pending_function: starting pending forms notification check', { force, dryRun, today: now.day });

  try {
    const catalyst = catalystSDK.initialize(context);
    const configRows = await fetchAllPendingNotificationRows(catalyst);
    const dueRows = rowsDueToday(configRows, now);

    if (!force && !dueRows.length) {
      console.log('pending_function: skipped — no pending notification date matches today', {
        today: `${now.day}/${now.month}/${now.year}`,
        configuredCount: configRows.length,
      });
      context.closeWithSuccess();
      return;
    }

    if (!configRows.length && !force) {
      console.log('pending_function: skipped — no pending notification records configured');
      context.closeWithSuccess();
      return;
    }

    const recipientEmails = uniqueEmails(force ? configRows : dueRows);
    if (!recipientEmails.length) {
      console.log('pending_function: skipped — no valid recipient email configured');
      context.closeWithSuccess();
      return;
    }

    console.log('pending_function: sending pending forms to configured emails only', {
      dueCount: dueRows.length,
      recipientEmails,
    });

    const result = await sendPendingFormsToRecipients(catalyst, {
      recipientEmails,
      dryRun,
    });

    console.log('pending_function: completed', {
      skipped: result.skipped,
      reason: result.reason || null,
      periodLabel: result.periodLabel,
      sitesConsidered: result.sitesConsidered,
      sitesWithPending: result.sitesWithPending,
      totalPendingForms: result.totalPendingForms,
      emailsSent: result.emailsSent,
      emailsFailed: result.emailsFailed,
      recipientEmails: result.recipientEmails || recipientEmails,
    });

    if (result.skipped) {
      if (result.reason && /mail sender|from_email|Mail →|MAIL_FROM/i.test(result.reason)) {
        console.error('pending_function: mail configuration missing', result.reason);
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

    console.error('pending_function: email was not sent', result.error || result.reason || '');
    context.closeWithFailure();
  } catch (err) {
    console.error('pending_function error:', err && err.message ? err.message : err);
    context.closeWithFailure();
  }
};
