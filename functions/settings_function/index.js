'use strict';

const express = require('express');
const catalystSDK = require('zcatalyst-sdk-node');
const fileUpload = require('express-fileupload');
const fs = require('fs');
const os = require('os');
const path = require('path');

const app = express();
const SETTINGS_TABLE = 'Settings';
const CHRO_TABLE = 'CHRO';
const PENDING_NOTIFICATION_TABLE = 'PendingNotification';
const REMINDER_NOTIFICATION_TABLE = 'Pending';
const MAX_REMINDER_LEVELS = 2;
const COMPANY_LOGO_FOLDER_ID = '31459000001361379';

app.use(express.json());
app.use(fileUpload());

app.use((req, res, next) => {
  try {
    res.locals.catalyst = catalystSDK.initialize(req);
    next();
  } catch (err) {
    res.status(500).json({ status: 'failure', message: 'Catalyst init failed' });
  }
});

function pickSettingsRow(rowObj) {
  if (!rowObj || typeof rowObj !== 'object') return null;
  return rowObj.Settings || rowObj.settings || rowObj;
}

function trimText(value) {
  return String(value ?? '').trim();
}

const NOTIFICATION_STATUS_ACTIVE = 'Active';
const NOTIFICATION_STATUS_PAUSED = 'Paused';

function normalizeNotificationStatus(value) {
  const raw = trimText(value).toLowerCase();
  if (raw === 'paused' || raw === 'inactive' || raw === 'disabled') {
    return NOTIFICATION_STATUS_PAUSED;
  }
  return NOTIFICATION_STATUS_ACTIVE;
}

function isNotificationActive(status) {
  return normalizeNotificationStatus(status) === NOTIFICATION_STATUS_ACTIVE;
}

function readFieldStatus(data, fieldReader) {
  const raw = fieldReader(data, ['Status', 'status', 'NotificationStatus', 'notificationStatus']);
  return normalizeNotificationStatus(raw || NOTIFICATION_STATUS_ACTIVE);
}

async function fetchLatestSettingsRow(catalyst) {
  const zcql = catalyst.zcql();
  const rows = await zcql.executeZCQLQuery(
    `SELECT ROWID, CompanyName, Logo, DueDate, Status FROM ${SETTINGS_TABLE} ORDER BY ROWID DESC LIMIT 1`
  );
  if (!Array.isArray(rows) || rows.length === 0) return null;
  return pickSettingsRow(rows[0]);
}

async function resolveLogoName(catalyst, fileId) {
  if (!fileId) return '';
  try {
    const details = await catalyst.filestore().folder(COMPANY_LOGO_FOLDER_ID).getFileDetails(fileId);
    if (Array.isArray(details) && details[0]?.file_name) return String(details[0].file_name);
    if (details?.file_name) return String(details.file_name);
    if (details?.name) return String(details.name);
    return '';
  } catch (_) {
    return '';
  }
}

async function updateChecklistBulkDueDates(catalyst, dueDate) {
  const table = catalyst.datastore().table('checklistbulk');
  const rows = await table.getAllRows();
  let updatedCount = 0;

  for (const row of rows) {
    if (!row?.ROWID) continue;
    await table.updateRow({ ROWID: row.ROWID, DueDate: dueDate });
    updatedCount += 1;
  }

  return updatedCount;
}

app.get('/', (req, res) => {
  res.status(200).json({ status: 'success', message: 'settings_function ready' });
});

app.get('/settings', async (req, res) => {
  try {
    const { catalyst } = res.locals;
    const row = await fetchLatestSettingsRow(catalyst);
    const companyName = String(row?.CompanyName || '').trim();
    const logo = row?.Logo || null;
    const dueDate = String(row?.DueDate || '').trim();
    const notificationStatus = normalizeNotificationStatus(row?.Status || row?.status);
    const logoName = await resolveLogoName(catalyst, logo);
    res.status(200).json({
      status: 'success',
      data: {
        rowId: row?.ROWID || null,
        companyName,
        logo,
        logoName,
        dueDate,
        notificationStatus,
      },
    });
  } catch (err) {
    res.status(500).json({ status: 'failure', message: err.message || 'Failed to load settings.' });
  }
});

app.get('/settings/logo', async (req, res) => {
  try {
    const { catalyst } = res.locals;
    const row = await fetchLatestSettingsRow(catalyst);
    const logoFileId = row?.Logo || null;
    if (!logoFileId) {
      return res.status(404).json({ status: 'failure', message: 'Logo not found.' });
    }

    const fileBuffer = await catalyst.filestore().folder(COMPANY_LOGO_FOLDER_ID).downloadFile(logoFileId);
    const logoName = await resolveLogoName(catalyst, logoFileId);
    const ext = String(logoName || '').toLowerCase().split('.').pop();
    const mime =
      ext === 'png'
        ? 'image/png'
        : ext === 'webp'
          ? 'image/webp'
          : ext === 'svg'
            ? 'image/svg+xml'
            : 'image/jpeg';

    res.writeHead(200, {
      'Content-Type': mime,
      'Content-Length': fileBuffer.length,
      'Cache-Control': 'no-store'
    });
    res.end(fileBuffer);
  } catch (err) {
    res.status(500).json({ status: 'failure', message: err.message || 'Failed to fetch logo.' });
  }
});

app.put('/settings', async (req, res) => {
  try {
    const { catalyst } = res.locals;
    const existing = await fetchLatestSettingsRow(catalyst);
    const companyName = String(req.body?.companyName || existing?.CompanyName || '').trim();
    const dueDate = String(req.body?.dueDate || '').trim();
    const notificationStatus = normalizeNotificationStatus(
      req.body?.notificationStatus || req.body?.status || existing?.Status || existing?.status
    );

    const table = catalyst.datastore().table(SETTINGS_TABLE);
    let logoFileId = existing?.Logo || null;

    if (req.files?.logo) {
      const logoFile = req.files.logo;
      const tempPath = path.join(os.tmpdir(), `${Date.now()}_${logoFile.name}`);
      await logoFile.mv(tempPath);
      try {
        const uploadResp = await catalyst.filestore().folder(COMPANY_LOGO_FOLDER_ID).uploadFile({
          code: fs.createReadStream(tempPath),
          name: logoFile.name,
        });
        if (Array.isArray(uploadResp) && uploadResp[0]?.id) logoFileId = uploadResp[0].id;
        else if (uploadResp?.id) logoFileId = uploadResp.id;
        else if (uploadResp?.file_details?.[0]?.id) logoFileId = uploadResp.file_details[0].id;
      } finally {
        if (fs.existsSync(tempPath)) fs.unlinkSync(tempPath);
      }
    }

    let savedRow;
    if (existing?.ROWID) {
      savedRow = await table.updateRow({
        ROWID: existing.ROWID,
        CompanyName: companyName,
        Logo: logoFileId || null,
        DueDate: dueDate,
        Status: notificationStatus,
      });
    } else {
      savedRow = await table.insertRow({
        CompanyName: companyName,
        Logo: logoFileId || null,
        DueDate: dueDate,
        Status: notificationStatus,
      });
    }

    const rowId = savedRow?.ROWID || existing?.ROWID || null;
    const logoName = await resolveLogoName(catalyst, logoFileId);
    const checklistBulkUpdatedCount = await updateChecklistBulkDueDates(catalyst, dueDate);
    res.status(200).json({
      status: 'success',
      data: {
        rowId,
        companyName,
        logo: logoFileId || null,
        logoName,
        dueDate,
        notificationStatus,
        checklistBulkUpdatedCount,
      },
    });
  } catch (err) {
    res.status(500).json({ status: 'failure', message: err.message || 'Failed to save settings.' });
  }
});

app.delete('/settings', async (req, res) => {
  try {
    const { catalyst } = res.locals;
    const existing = await fetchLatestSettingsRow(catalyst);
    if (!existing?.ROWID) {
      return res.status(200).json({ status: 'success', message: 'No settings found.' });
    }

    const table = catalyst.datastore().table(SETTINGS_TABLE);
    await table.deleteRow(existing.ROWID);
    res.status(200).json({ status: 'success', message: 'Settings deleted successfully.' });
  } catch (err) {
    res.status(500).json({ status: 'failure', message: err.message || 'Failed to delete settings.' });
  }
});

function pickChroRow(rowObj) {
  if (!rowObj || typeof rowObj !== 'object' || Array.isArray(rowObj)) return null;
  const nested = rowObj.CHRO || rowObj.Chro || rowObj.chro || null;
  if (nested && typeof nested === 'object' && !Array.isArray(nested)) return nested;
  const keys = Object.keys(rowObj);
  if (keys.length === 1) {
    const only = rowObj[keys[0]];
    if (only && typeof only === 'object' && !Array.isArray(only)) return only;
  }
  return rowObj;
}

function chroField(data, names) {
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

function toChroAppRow(row) {
  const data = pickChroRow(row) || {};
  const id =
    chroField(data, ['ROWID', 'rowId', 'rowid']) ||
    chroField(row || {}, ['ROWID', 'rowId', 'rowid']);
  return {
    id: id !== '' ? id : null,
    name: chroField(data, ['Name', 'name', 'CHROName', 'chroName']),
    email: chroField(data, ['Email', 'email', 'EmailId', 'emailId', 'MailId', 'MailID', 'Mail', 'mail']),
    status: readFieldStatus(data, chroField),
  };
}

async function fetchAllChroRows(catalyst) {
  let raw = [];
  try {
    const zcql = catalyst.zcql();
    const pageSize = 100;
    let offset = 0;
    while (offset < 2000) {
      const chunk = await zcql.executeZCQLQuery(
        `SELECT * FROM ${CHRO_TABLE} LIMIT ${offset},${pageSize}`
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
      raw = (await catalyst.datastore().table(CHRO_TABLE).getAllRows()) || [];
    } catch (_) {
      raw = [];
    }
  }
  return raw
    .map(toChroAppRow)
    .filter((row) => row.id || row.name || row.email);
}

app.get('/chro', async (req, res) => {
  try {
    const { catalyst } = res.locals;
    const data = await fetchAllChroRows(catalyst);
    res.status(200).json({ status: 'success', data });
  } catch (err) {
    res.status(500).json({ status: 'failure', message: err.message || 'Failed to load CHRO records.' });
  }
});

app.put('/chro', async (req, res) => {
  try {
    const { catalyst } = res.locals;
    const rowId = trimText(req.body?.rowId || req.body?.id || req.body?.ROWID);
    const name = trimText(req.body?.name || req.body?.Name);
    const email = trimText(req.body?.email || req.body?.Email);
    const notificationStatus = normalizeNotificationStatus(req.body?.notificationStatus || req.body?.status);

    if (!name) {
      return res.status(400).json({ status: 'failure', message: 'CHRO Name is required.' });
    }
    if (!email) {
      return res.status(400).json({ status: 'failure', message: 'Email is required.' });
    }
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
      return res.status(400).json({ status: 'failure', message: 'Enter a valid email address.' });
    }

    const table = catalyst.datastore().table(CHRO_TABLE);
    const payload = { Name: name, Email: email, Status: notificationStatus };
    const saved = rowId
      ? await table.updateRow({ ROWID: rowId, ...payload })
      : await table.insertRow(payload);
    const appRow = toChroAppRow(saved);
    res.status(200).json({
      status: 'success',
      data: {
        id: appRow.id || rowId || null,
        name: appRow.name || name,
        email: appRow.email || email,
        status: appRow.status || notificationStatus,
      },
    });
  } catch (err) {
    res.status(500).json({ status: 'failure', message: err.message || 'Failed to save CHRO record.' });
  }
});

app.delete('/chro', async (req, res) => {
  try {
    const { catalyst } = res.locals;
    const rowId = trimText(req.query?.rowId || req.body?.rowId || req.body?.id || req.body?.ROWID);
    if (!rowId) {
      return res.status(400).json({ status: 'failure', message: 'rowId is required.' });
    }
    await catalyst.datastore().table(CHRO_TABLE).deleteRow(rowId);
    res.status(200).json({ status: 'success', message: 'CHRO record deleted successfully.' });
  } catch (err) {
    res.status(500).json({ status: 'failure', message: err.message || 'Failed to delete CHRO record.' });
  }
});

function pickPendingNotificationRow(rowObj) {
  if (!rowObj || typeof rowObj !== 'object' || Array.isArray(rowObj)) return null;
  const nested =
    rowObj.PendingNotification ||
    rowObj.pendingNotification ||
    rowObj.Pending_Notification ||
    null;
  if (nested && typeof nested === 'object' && !Array.isArray(nested)) return nested;
  const keys = Object.keys(rowObj);
  if (keys.length === 1) {
    const only = rowObj[keys[0]];
    if (only && typeof only === 'object' && !Array.isArray(only)) return only;
  }
  return rowObj;
}

function pendingNotificationField(data, names) {
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

function toPendingNotificationAppRow(row) {
  const data = pickPendingNotificationRow(row) || {};
  const id =
    pendingNotificationField(data, ['ROWID', 'rowId', 'rowid']) ||
    pendingNotificationField(row || {}, ['ROWID', 'rowId', 'rowid']);
  return {
    id: id !== '' ? id : null,
    email: pendingNotificationField(data, [
      'Email',
      'email',
      'EmailId',
      'emailId',
      'MailId',
      'MailID',
      'Mail',
      'mail',
    ]),
    notificationDate: pendingNotificationField(data, [
      'NotificationDate',
      'notificationDate',
      'Notification_Date',
      'notification_date',
    ]),
    status: readFieldStatus(data, pendingNotificationField),
  };
}

async function fetchAllPendingNotificationRows(catalyst) {
  let raw = [];
  try {
    const zcql = catalyst.zcql();
    const pageSize = 100;
    let offset = 0;
    while (offset < 2000) {
      const chunk = await zcql.executeZCQLQuery(
        `SELECT * FROM ${PENDING_NOTIFICATION_TABLE} LIMIT ${offset},${pageSize}`
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
      raw = (await catalyst.datastore().table(PENDING_NOTIFICATION_TABLE).getAllRows()) || [];
    } catch (_) {
      raw = [];
    }
  }
  return raw
    .map(toPendingNotificationAppRow)
    .filter((row) => row.id || row.email || row.notificationDate);
}

app.get('/pending-notification', async (req, res) => {
  try {
    const { catalyst } = res.locals;
    const data = await fetchAllPendingNotificationRows(catalyst);
    res.status(200).json({ status: 'success', data });
  } catch (err) {
    res.status(500).json({
      status: 'failure',
      message: err.message || 'Failed to load pending notification records.',
    });
  }
});

app.put('/pending-notification', async (req, res) => {
  try {
    const { catalyst } = res.locals;
    const rowId = trimText(req.body?.rowId || req.body?.id || req.body?.ROWID);
    const email = trimText(req.body?.email || req.body?.Email);
    const notificationDate = trimText(
      req.body?.notificationDate || req.body?.NotificationDate
    );
    const notificationStatus = normalizeNotificationStatus(req.body?.notificationStatus || req.body?.status);

    if (!email) {
      return res.status(400).json({ status: 'failure', message: 'Email is required.' });
    }
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
      return res.status(400).json({ status: 'failure', message: 'Enter a valid email address.' });
    }
    if (!notificationDate) {
      return res.status(400).json({ status: 'failure', message: 'Notification Date is required.' });
    }

    const table = catalyst.datastore().table(PENDING_NOTIFICATION_TABLE);
    const payload = { Email: email, NotificationDate: notificationDate, Status: notificationStatus };
    const saved = rowId
      ? await table.updateRow({ ROWID: rowId, ...payload })
      : await table.insertRow(payload);
    const appRow = toPendingNotificationAppRow(saved);
    res.status(200).json({
      status: 'success',
      data: {
        id: appRow.id || rowId || null,
        email: appRow.email || email,
        notificationDate: appRow.notificationDate || notificationDate,
        status: appRow.status || notificationStatus,
      },
    });
  } catch (err) {
    res.status(500).json({
      status: 'failure',
      message: err.message || 'Failed to save pending notification record.',
    });
  }
});

app.delete('/pending-notification', async (req, res) => {
  try {
    const { catalyst } = res.locals;
    const rowId = trimText(req.query?.rowId || req.body?.rowId || req.body?.id || req.body?.ROWID);
    if (!rowId) {
      return res.status(400).json({ status: 'failure', message: 'rowId is required.' });
    }
    await catalyst.datastore().table(PENDING_NOTIFICATION_TABLE).deleteRow(rowId);
    res.status(200).json({ status: 'success', message: 'Pending notification record deleted successfully.' });
  } catch (err) {
    res.status(500).json({
      status: 'failure',
      message: err.message || 'Failed to delete pending notification record.',
    });
  }
});

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
  const id =
    reminderNotificationField(data, ['ROWID', 'rowId', 'rowid']) ||
    reminderNotificationField(row || {}, ['ROWID', 'rowId', 'rowid']);
  const levelRaw = reminderNotificationField(data, ['Level', 'level', 'NotificationLevel', 'notificationLevel']);
  const level = Number(levelRaw) || fallbackLevel || null;
  return {
    id: id !== '' ? id : null,
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
    status: readFieldStatus(data, reminderNotificationField),
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
    .filter((row) => row.id || row.email || row.notificationDate);
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

function buildReminderLevelSlots(rows) {
  const slots = Array.from({ length: MAX_REMINDER_LEVELS }, (_, index) => ({
    level: index + 1,
    id: null,
    email: '',
    notificationDate: '',
    status: NOTIFICATION_STATUS_ACTIVE,
  }));
  const usedLevels = new Set();
  for (const row of rows || []) {
    let level = Number(row.level);
    if (!level || level < 1 || level > MAX_REMINDER_LEVELS || usedLevels.has(level)) {
      level = slots.find((slot) => !slot.id)?.level;
    }
    if (!level) continue;
    usedLevels.add(level);
    const slot = slots[level - 1];
    slot.id = row.id;
    slot.email = row.email || '';
    slot.notificationDate = row.notificationDate || '';
    slot.status = row.status || NOTIFICATION_STATUS_ACTIVE;
  }
  return slots;
}

app.get('/reminder-notification', async (req, res) => {
  try {
    const { catalyst } = res.locals;
    const rows = await fetchAllReminderNotificationRows(catalyst);
    res.status(200).json({ status: 'success', data: buildReminderLevelSlots(rows) });
  } catch (err) {
    res.status(500).json({
      status: 'failure',
      message: err.message || 'Failed to load reminder notification records.',
    });
  }
});

app.put('/reminder-notification', async (req, res) => {
  try {
    const { catalyst } = res.locals;
    const rowId = trimText(req.body?.rowId || req.body?.id || req.body?.ROWID);
    const level = Number(trimText(req.body?.level || req.body?.Level));
    const email = trimText(req.body?.email || req.body?.Email);
    const notificationDate = trimText(
      req.body?.notificationDate || req.body?.NotificationDate
    );
    const notificationStatus = normalizeNotificationStatus(req.body?.notificationStatus || req.body?.status);

    if (!level || level < 1 || level > MAX_REMINDER_LEVELS) {
      return res.status(400).json({
        status: 'failure',
        message: `Level must be between 1 and ${MAX_REMINDER_LEVELS}.`,
      });
    }
    if (!email) {
      return res.status(400).json({ status: 'failure', message: 'Email is required.' });
    }
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
      return res.status(400).json({ status: 'failure', message: 'Enter a valid email address.' });
    }
    if (!notificationDate) {
      return res.status(400).json({ status: 'failure', message: 'Notification Date is required.' });
    }

    const existingRows = await fetchAllReminderNotificationRows(catalyst);
    const levelRow = existingRows.find((row) => Number(row.level) === level);
    const table = catalyst.datastore().table(REMINDER_NOTIFICATION_TABLE);
    const payload = {
      Email: email,
      NotificationDate: notificationDate,
      Level: String(level),
      Status: notificationStatus,
    };

    if (!rowId && existingRows.length >= MAX_REMINDER_LEVELS && !levelRow) {
      return res.status(400).json({
        status: 'failure',
        message: `Maximum ${MAX_REMINDER_LEVELS} reminder notification levels are allowed.`,
      });
    }

    const saved = rowId || levelRow?.id
      ? await table.updateRow({ ROWID: rowId || levelRow.id, ...payload })
      : await table.insertRow(payload);
    const appRow = toReminderNotificationAppRow(saved, level);
    res.status(200).json({
      status: 'success',
      data: {
        id: appRow.id || rowId || levelRow?.id || null,
        level: appRow.level || level,
        email: appRow.email || email,
        notificationDate: appRow.notificationDate || notificationDate,
        status: appRow.status || notificationStatus,
      },
    });
  } catch (err) {
    res.status(500).json({
      status: 'failure',
      message: err.message || 'Failed to save reminder notification record.',
    });
  }
});

app.delete('/reminder-notification', async (req, res) => {
  try {
    const { catalyst } = res.locals;
    const rowId = trimText(req.query?.rowId || req.body?.rowId || req.body?.id || req.body?.ROWID);
    if (!rowId) {
      return res.status(400).json({ status: 'failure', message: 'rowId is required.' });
    }
    await catalyst.datastore().table(REMINDER_NOTIFICATION_TABLE).deleteRow(rowId);
    res.status(200).json({ status: 'success', message: 'Reminder notification record deleted successfully.' });
  } catch (err) {
    res.status(500).json({
      status: 'failure',
      message: err.message || 'Failed to delete reminder notification record.',
    });
  }
});

module.exports = app;
