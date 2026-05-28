'use strict';

const express = require('express');
const catalystSDK = require('zcatalyst-sdk-node');
const fileUpload = require('express-fileupload');
const fs = require('fs');
const os = require('os');
const path = require('path');

const app = express();
const SETTINGS_TABLE = 'Settings';
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

async function fetchLatestSettingsRow(catalyst) {
  const zcql = catalyst.zcql();
  const rows = await zcql.executeZCQLQuery(
    `SELECT ROWID, CompanyName, Logo FROM ${SETTINGS_TABLE} ORDER BY ROWID DESC LIMIT 1`
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

app.get('/', (req, res) => {
  res.status(200).json({ status: 'success', message: 'settings_function ready' });
});

app.get('/settings', async (req, res) => {
  try {
    const { catalyst } = res.locals;
    const row = await fetchLatestSettingsRow(catalyst);
    const companyName = String(row?.CompanyName || '').trim();
    const logo = row?.Logo || null;
    const logoName = await resolveLogoName(catalyst, logo);
    res.status(200).json({
      status: 'success',
      data: {
        rowId: row?.ROWID || null,
        companyName,
        logo,
        logoName,
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
    const companyName = String(req.body?.companyName || '').trim();
    if (!companyName) {
      return res.status(400).json({ status: 'failure', message: 'Company name is required.' });
    }

    const table = catalyst.datastore().table(SETTINGS_TABLE);
    const existing = await fetchLatestSettingsRow(catalyst);
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
      });
    } else {
      savedRow = await table.insertRow({
        CompanyName: companyName,
        Logo: logoFileId || null,
      });
    }

    const rowId = savedRow?.ROWID || existing?.ROWID || null;
    const logoName = await resolveLogoName(catalyst, logoFileId);
    res.status(200).json({
      status: 'success',
      data: {
        rowId,
        companyName,
        logo: logoFileId || null,
        logoName,
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

module.exports = app;
