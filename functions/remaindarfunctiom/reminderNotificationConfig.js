'use strict';

const REMINDER_NOTIFICATION_TABLE = 'Pending';
const MAX_REMINDER_LEVELS = 2;

function trimText(value) {
  return String(value ?? '').trim();
}

function normalizeNotificationStatus(value) {
  const raw = trimText(value).toLowerCase();
  if (raw === 'paused' || raw === 'inactive' || raw === 'disabled') return 'Paused';
  return 'Active';
}

function isNotificationActive(status) {
  return normalizeNotificationStatus(status) === 'Active';
}

function pickReminderNotificationRow(rowObj) {
  if (!rowObj || typeof rowObj !== 'object' || Array.isArray(rowObj)) return null;
  const nested =
    rowObj.Pending ||
    rowObj.pending ||
    rowObj.ReminderNotification ||
    rowObj.reminderNotification ||
    null;
  if (nested && typeof nested === 'object' && !Array.isArray(nested)) return nested;
  const keys = Object.keys(rowObj);
  if (keys.length === 1) {
    const only = rowObj[keys[0]];
    if (only && typeof only === 'object' && !Array.isArray(only)) return only;
  }
  return rowObj;
}

function reminderNotificationField(data, names) {
  if (!data || typeof data !== 'object') return '';
  for (const name of names) {
    if (data[name] != null && trimText(data[name]) !== '') return trimText(data[name]);
  }
  const lookup = {};
  Object.keys(data).forEach((key) => {
    lookup[String(key).toLowerCase().replace(/[_\s-]/g, '')] = data[key];
  });
  for (const name of names) {
    const value = lookup[String(name).toLowerCase().replace(/[_\s-]/g, '')];
    if (value != null && trimText(value) !== '') return trimText(value);
  }
  return '';
}

function toReminderNotificationAppRow(row, fallbackLevel) {
  const data = pickReminderNotificationRow(row) || {};
  const levelRaw = reminderNotificationField(data, ['Level', 'level', 'NotificationLevel', 'notificationLevel']);
  const level = Number(levelRaw) || fallbackLevel || null;
  return {
    id:
      reminderNotificationField(data, ['ROWID', 'rowId', 'rowid']) ||
      reminderNotificationField(row || {}, ['ROWID', 'rowId', 'rowid']) ||
      null,
    level,
    email: reminderNotificationField(data, [
      'Email',
      'email',
      'EmailId',
      'emailId',
      'MailId',
      'MailID',
      'Mail',
      'mail',
    ]),
    notificationDate: reminderNotificationField(data, [
      'NotificationDate',
      'notificationDate',
      'Notification_Date',
      'notification_date',
    ]),
    status: normalizeNotificationStatus(
      reminderNotificationField(data, ['Status', 'status', 'NotificationStatus', 'notificationStatus'])
    ),
  };
}

async function fetchAllReminderNotificationRows(catalyst) {
  let raw = [];
  try {
    const zcql = catalyst.zcql();
    const pageSize = 100;
    let offset = 0;
    while (offset < 2000) {
      const chunk = await zcql.executeZCQLQuery(
        `SELECT * FROM ${REMINDER_NOTIFICATION_TABLE} LIMIT ${offset},${pageSize}`
      );
      if (!Array.isArray(chunk) || chunk.length === 0) break;
      chunk.forEach((row) => raw.push(row));
      if (chunk.length < pageSize) break;
      offset += pageSize;
    }
  } catch (_) {
    raw = [];
  }
  if (!raw.length) {
    try {
      raw = (await catalyst.datastore().table(REMINDER_NOTIFICATION_TABLE).getAllRows()) || [];
    } catch (_) {
      raw = [];
    }
  }
  const mapped = raw
    .map((row, index) => toReminderNotificationAppRow(row, index + 1))
    .filter((row) => row.email && row.notificationDate);
  mapped.sort((a, b) => {
    const levelA = Number(a.level) || 0;
    const levelB = Number(b.level) || 0;
    if (levelA && levelB && levelA !== levelB) return levelA - levelB;
    return String(a.id || '').localeCompare(String(b.id || ''), undefined, { numeric: true });
  });
  mapped.forEach((row, index) => {
    if (!row.level) row.level = index + 1;
  });
  return mapped
    .filter((row) => {
      const level = Number(row.level);
      return level >= 1 && level <= MAX_REMINDER_LEVELS;
    })
    .slice(0, MAX_REMINDER_LEVELS);
}

function parseNotificationDateParts(value) {
  const raw = trimText(value);
  if (!raw) return null;
  if (/^\d{1,2}$/.test(raw)) {
    const day = Number(raw);
    if (day >= 1 && day <= 31) return { day, month: null };
  }
  const isoMatch = raw.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (isoMatch) {
    return { day: Number(isoMatch[3]), month: Number(isoMatch[2]) };
  }
  const parsed = new Date(raw);
  if (!Number.isNaN(parsed.getTime())) {
    return { day: parsed.getDate(), month: parsed.getMonth() + 1 };
  }
  return null;
}

function isNotificationDueToday(notificationDate, now) {
  const parts = parseNotificationDateParts(notificationDate);
  if (!parts || !parts.day) return false;
  if (Number(now.day) !== parts.day) return false;
  if (parts.month != null && Number(now.month) !== parts.month) return false;
  return true;
}

function rowsDueToday(rows, now) {
  return (rows || []).filter(
    (row) => isNotificationActive(row.status) && isNotificationDueToday(row.notificationDate, now)
  );
}

function nowInKolkata() {
  const formatter = new Intl.DateTimeFormat('en-GB', {
    timeZone: 'Asia/Kolkata',
    year: 'numeric',
    month: 'numeric',
    day: 'numeric',
  });
  const parts = formatter.formatToParts(new Date());
  const lookup = {};
  parts.forEach((part) => {
    lookup[part.type] = part.value;
  });
  return {
    day: Number(lookup.day),
    month: Number(lookup.month),
    year: Number(lookup.year),
  };
}

module.exports = {
  fetchAllReminderNotificationRows,
  isNotificationDueToday,
  rowsDueToday,
  nowInKolkata,
  MAX_REMINDER_LEVELS,
};
