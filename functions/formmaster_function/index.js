'use strict';

const express = require('express');
const catalystSDK = require('zcatalyst-sdk-node');
const catalystConstants = require('zcatalyst-sdk-node/lib/utils/constants').default;

const app = express();
app.use(express.json());

app.use((req, res, next) => {
  try {
    const catalyst = catalystSDK.initialize(req);
    res.locals.catalyst = catalyst;
    next();
  } catch (err) {
    res.status(500).json({ status: 'failure', message: 'Catalyst init failed' });
  }
});

// File Store Templates folder ID (same as in Catalyst console)
const TEMPLATES_FOLDER_ID = '31459000000501475';
const TABLE_NAME = 'FormMaster';

function formatSizeForStore(bytes) {
  if (bytes == null || typeof bytes !== 'number') return '';
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

function getContentType(filename) {
  const ext = String(filename || '').toLowerCase().split('.').pop();
  const map = {
    pdf: 'application/pdf',
    xls: 'application/vnd.ms-excel',
    xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    txt: 'text/plain'
  };
  return map[ext] || 'application/octet-stream';
}

function normalizeFileToItem(f) {
  const detail = typeof f === 'object' && f !== null ? f : {};
  const rawCreatedBy = detail.created_by || detail.createdBy;
  const createdByStr =
    rawCreatedBy == null
      ? null
      : typeof rawCreatedBy === 'string'
        ? rawCreatedBy.trim() || null
        : typeof rawCreatedBy === 'object' && rawCreatedBy !== null
          ? (rawCreatedBy.email_id || rawCreatedBy.email || rawCreatedBy.user_id || '').trim() || null
          : null;
  return {
    id: detail.id != null ? String(detail.id) : null,
    name: detail.file_name || detail.name || detail.fileName || 'Unknown',
    size: detail.file_size != null ? detail.file_size : (detail.size != null ? detail.size : null),
    createdTime: detail.created_time || detail.createdTime || null,
    createdBy: createdByStr
  };
}

app.get('/', (req, res) => {
  res.status(200).json({ status: 'success', message: 'formmaster_function ready' });
});

// List templates from File Store Templates folder (FORM 10.xlsx, Form U.xlsx, etc.)
app.get('/templates', async (req, res) => {
  try {
    const formName = (req.query.formName || '').trim();
    const { catalyst } = res.locals;
    const filestore = catalyst.filestore();
    let files = [];
    const folderObj = await filestore.getFolderDetails(TEMPLATES_FOLDER_ID);
    const folderJson = folderObj.toJSON ? folderObj.toJSON() : folderObj;
    if (Array.isArray(folderJson.file_details) && folderJson.file_details.length > 0) {
      files = folderJson.file_details;
    } else {
      const folder = filestore.folder(TEMPLATES_FOLDER_ID);
      const folderId = folder.toJSON().id;
      const request = {
        method: catalystConstants.REQ_METHOD.get,
        path: `/folder/${folderId}/file`,
        type: 'json',
        catalyst: true,
        track: true,
        user: catalystConstants.CREDENTIAL_USER.user
      };
      const resp = await folder.requester.send(request);
      const raw = resp.data && resp.data.data != null ? resp.data.data : (Array.isArray(resp.data) ? resp.data : []);
      files = Array.isArray(raw) ? raw : [];
    }
    let items = files.map(normalizeFileToItem).filter((item) => item.id && item.name);
    if (formName) {
      const lower = formName.toLowerCase();
      items = items.filter((item) => item.name.toLowerCase().includes(lower));
    }
    res.status(200).json({ status: 'success', data: items });
  } catch (err) {
    console.error('Formmaster templates list error:', err);
    res.status(500).json({
      status: 'failure',
      message: err.message || 'Failed to list templates.'
    });
  }
});

// Sync templates from File Store into FormMaster Data Store table (FormName, Size, Action)
app.post('/templates/sync', async (req, res) => {
  try {
    const { catalyst } = res.locals;
    const filestore = catalyst.filestore();
    let files = [];
    const folderObj = await filestore.getFolderDetails(TEMPLATES_FOLDER_ID);
    const folderJson = folderObj.toJSON ? folderObj.toJSON() : folderObj;
    if (Array.isArray(folderJson.file_details) && folderJson.file_details.length > 0) {
      files = folderJson.file_details;
    } else {
      const folder = filestore.folder(TEMPLATES_FOLDER_ID);
      const folderId = folder.toJSON().id;
      const request = {
        method: catalystConstants.REQ_METHOD.get,
        path: `/folder/${folderId}/file`,
        type: 'json',
        catalyst: true,
        track: true,
        user: catalystConstants.CREDENTIAL_USER.user
      };
      const resp = await folder.requester.send(request);
      const raw = resp.data && resp.data.data != null ? resp.data.data : (Array.isArray(resp.data) ? resp.data : []);
      files = Array.isArray(raw) ? raw : [];
    }
    const items = files.map(normalizeFileToItem).filter((item) => item.id && item.name);
    const table = catalyst.datastore().table(TABLE_NAME);
    const existingRows = await table.getAllRows();
    const existingActionIds = new Set(
      (existingRows || []).map((r) => (r.Action != null ? String(r.Action).trim() : '')).filter(Boolean)
    );
    let inserted = 0;
    for (const item of items) {
      if (!item.id || existingActionIds.has(item.id)) continue;
      const sizeStr = typeof item.size === 'number' ? formatSizeForStore(item.size) : (item.size || '');
      await table.insertRow({
        FormName: item.name || 'Unknown',
        Size: sizeStr,
        Action: item.id
      });
      existingActionIds.add(item.id);
      inserted++;
    }
    res.status(200).json({
      status: 'success',
      message: `Synced ${inserted} template(s) to FormMaster. Total files in folder: ${items.length}.`
    });
  } catch (err) {
    console.error('Formmaster templates sync error:', err);
    res.status(500).json({
      status: 'failure',
      message: err.message || 'Failed to sync templates to Data Store.'
    });
  }
});

// Download template file from File Store (Action column = file ID in Templates folder)
app.get('/templates/download/:fileId', async (req, res) => {
  try {
    const { fileId } = req.params;
    const { catalyst } = res.locals;
    const folder = catalyst.filestore().folder(TEMPLATES_FOLDER_ID);
    const fileBuffer = await folder.downloadFile(fileId);
    const disposition = req.query.disposition === 'inline' ? 'inline' : 'attachment';
    const fileName = (req.query.fileName || 'template.xlsx').replace(/[^a-zA-Z0-9._-]/g, '_');
    res.writeHead(200, {
      'Content-Type': getContentType(fileName),
      'Content-Disposition': `${disposition}; filename="${fileName}"`,
      'Content-Length': fileBuffer.length
    });
    res.end(fileBuffer);
  } catch (err) {
    console.error('Formmaster template download error:', err);
    res.status(500).json({ status: 'failure', message: err.message || 'Download failed.' });
  }
});

module.exports = app;
