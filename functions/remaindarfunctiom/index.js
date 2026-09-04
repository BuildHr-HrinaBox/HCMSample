'use strict';

const catalystSDK = require('zcatalyst-sdk-node');
const {
  fetchAllReminderNotificationRows,
  rowsDueToday,
  nowInKolkata,
} = require('./reminderNotificationConfig');
const { sendReminderEmails } = require('./reminderMailer');

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
 * Daily cron — sends reminder notification emails when today matches
 * a configured Notification Date from Settings → Reminder Notification (Asia/Kolkata).
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

  console.log('remaindarfunctiom: starting reminder notification check', {
    force,
    dryRun,
    today: `${now.day}/${now.month}/${now.year}`,
  });

  try {
    const catalyst = catalystSDK.initialize(context);
    const configRows = await fetchAllReminderNotificationRows(catalyst);
    const dueRows = rowsDueToday(configRows, now);

    if (!force && !dueRows.length) {
      console.log('remaindarfunctiom: skipped — no reminder notification date matches today', {
        today: `${now.day}/${now.month}/${now.year}`,
        configuredCount: configRows.length,
      });
      context.closeWithSuccess();
      return;
    }

    if (!configRows.length && !force) {
      console.log('remaindarfunctiom: skipped — no reminder notification records configured');
      context.closeWithSuccess();
      return;
    }

    const rowsToSend = force ? configRows : dueRows;
    if (!rowsToSend.length) {
      console.log('remaindarfunctiom: skipped — no reminder rows to process');
      context.closeWithSuccess();
      return;
    }

    console.log('remaindarfunctiom: sending reminder notifications', {
      dueCount: dueRows.length,
      sendCount: rowsToSend.length,
      levels: rowsToSend.map((row) => row.level),
    });

    const result = await sendReminderEmails(catalyst, rowsToSend, now, { dryRun });

    console.log('remaindarfunctiom: completed', {
      skipped: result.skipped,
      reason: result.reason || null,
      periodLabel: result.periodLabel,
      pendingForms: result.pendingSites?.length || 0,
      sitesWithPending: result.sitesWithPending,
      sitesAllComplete: result.sitesAllComplete,
      activeSites: result.activeCount,
      pendingCount: result.pendingCount,
      yetToCompleteCount: result.yetToCompleteCount,
      emailsSent: result.emailsSent,
      emailsFailed: result.emailsFailed,
      details: result.details
    });

    if (result.skipped) {
      if (result.reason && /mail sender|from_email|Mail →|MAIL_FROM/i.test(result.reason)) {
        console.error('remaindarfunctiom: mail configuration missing', result.reason);
        context.closeWithFailure();
        return;
      }
      console.log('remaindarfunctiom: skipped', result.reason || 'no work to do');
      context.closeWithSuccess();
      return;
    }

    if (result.emailsSent > 0) {
      context.closeWithSuccess();
      return;
    }

    console.error('remaindarfunctiom: email was not sent', result.reason || '');
    context.closeWithFailure();
  } catch (err) {
    console.error('remaindarfunctiom error:', err && err.message ? err.message : err);
    context.closeWithFailure();
  }
};
