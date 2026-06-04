'use strict';

const express = require('express');
const catalystSDK = require('zcatalyst-sdk-node');
const catalystConstants = require('zcatalyst-sdk-node/lib/utils/constants').default;
const fileUpload = require('express-fileupload');
const fs = require('fs');
const os = require('os');
const path = require('path');

const app = express();
app.use(express.json());
app.use(fileUpload());

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

function toStoreStr(val) {
  if (val === undefined || val === null) return '';
  return String(val).trim();
}

function squashKeyPart(v) {
  return String(v ?? '')
    .trim()
    .toLowerCase()
    .replace(/\.(xlsx|xls)$/i, '')
    .replace(/\s+/g, ' ');
}

/** Unique template per row: same form name + different act/description/sector/state => different file */
function buildFormMasterMatchKey(data) {
  return [
    squashKeyPart(data.formName),
    squashKeyPart(data.act),
    squashKeyPart(data.description),
    squashKeyPart(data.sector) || 'nosector',
    squashKeyPart(data.state) || 'nostate'
  ].join('|');
}

function createFormMasterRecord(data) {
  const record = {
    Act: toStoreStr(data.act),
    Description: toStoreStr(data.description),
    Sector: toStoreStr(data.sector),
    State: toStoreStr(data.state),
    FormName: toStoreStr(data.formName),
    TemplateFileName: toStoreStr(data.templateFileName),
    Action: toStoreStr(data.action),
    Size: toStoreStr(data.size)
  };
  return record;
}

function readRecordAction(record) {
  if (!record || typeof record !== 'object') return '';
  const keys = ['Action', 'action', 'FileId', 'fileId', 'FormFile', 'formFile'];
  for (let i = 0; i < keys.length; i += 1) {
    const k = keys[i];
    if (record[k] == null) continue;
    const v = String(record[k]).trim();
    if (v && v !== 'null') return v;
  }
  return '';
}

function convertRecordToAppFormat(record) {
  const action = readRecordAction(record);
  const formName = record.FormName || '';
  const templateFileName = buildDisplayTemplateFileName(record);
  return {
    id: record.ROWID,
    act: record.Act || '',
    description: record.Description || '',
    sector: record.Sector || '',
    state: record.State || '',
    formName,
    templateFileName,
    action: action || null,
    size: record.Size || '',
    fileId: action || null,
    fileName: templateFileName || null,
    matchKey: buildFormMasterMatchKey({
      formName,
      act: record.Act,
      description: record.Description,
      sector: record.Sector,
      state: record.State
    }),
    createdTime: record.CREATEDTIME,
    modifiedTime: record.MODIFIEDTIME
  };
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

async function listTemplateFiles(catalyst) {
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
  return files.map(normalizeFileToItem).filter((item) => item.id && item.name);
}

function extractUploadFileId(uploadResp) {
  if (Array.isArray(uploadResp) && uploadResp[0]?.id) return uploadResp[0].id;
  if (uploadResp && uploadResp.id) return uploadResp.id;
  if (uploadResp && uploadResp.file_details && uploadResp.file_details[0]?.id) {
    return uploadResp.file_details[0].id;
  }
  return null;
}

/** File Store name must be unique per row — same "Form U.xlsx" must not overwrite another state's file. */
function buildUniqueStorageFileName(row, originalName) {
  const raw = String(originalName || 'template.xlsx').trim();
  const ext = path.extname(raw) || '.xlsx';
  const base = path.basename(raw, ext).replace(/[^a-zA-Z0-9_-]/g, '_').slice(0, 24) || 'template';
  const slug = [
    squashKeyPart(row?.FormName),
    squashKeyPart(row?.State),
    squashKeyPart(row?.Sector)
  ]
    .filter((s) => s && s !== 'nosector' && s !== 'nostate')
    .join('_')
    .replace(/[^a-z0-9_]/gi, '_')
    .slice(0, 48);
  const rowId = String(row?.ROWID || row?.id || Date.now());
  return `${base}${slug ? `_${slug}` : ''}_r${rowId}${ext}`;
}

function buildDisplayTemplateFileName(record) {
  const original = String(record?.TemplateFileName || '').trim();
  const formName = String(record?.FormName || 'Form').trim();
  const state = String(record?.State || '').trim();
  const act = String(record?.Act || '').trim();
  const looksGeneric =
    !original ||
    /^form\s*[a-z0-9]*\.xlsx$/i.test(original) ||
    original.toLowerCase() === `${formName.toLowerCase()}.xlsx`;
  if (!looksGeneric) return original;
  const parts = [formName];
  if (state) parts.push(state);
  else if (act) parts.push(act.slice(0, 40));
  return `${parts.join(' - ')}${path.extname(original) || '.xlsx'}`.replace(/[/\\?%*:|"<>]/g, '');
}

async function uploadToTemplatesFolder(catalyst, file, row) {
  const tempDir = os.tmpdir();
  const storageName = buildUniqueStorageFileName(row, file.name);
  const tempPath = path.join(tempDir, storageName);
  await file.mv(tempPath);
  try {
    const uploadResp = await catalyst.filestore().folder(TEMPLATES_FOLDER_ID).uploadFile({
      code: fs.createReadStream(tempPath),
      name: storageName
    });
    const fileId = extractUploadFileId(uploadResp);
    if (!fileId) throw new Error('File upload failed or invalid response from Catalyst.');
    const displayName = buildDisplayTemplateFileName({
      ...row,
      TemplateFileName: file.name
    });
    return {
      fileId: String(fileId),
      fileName: displayName,
      storageFileName: storageName,
      size: file.size
    };
  } finally {
    if (fs.existsSync(tempPath)) fs.unlinkSync(tempPath);
  }
}

const getAllFormMasterRecords = async (catalyst) => {
  const table = catalyst.datastore().table(TABLE_NAME);
  const rows = await table.getAllRows();
  const data = (rows || []).map(convertRecordToAppFormat);
  return { status: 'success', data, count: data.length };
};

const bulkImportFormMasterRecords = async (catalyst, dataArray) => {
  const table = catalyst.datastore().table(TABLE_NAME);
  const records = dataArray.map((d) => createFormMasterRecord(d));
  const result = await table.insertRows(records);
  return {
    status: 'success',
    message: `Imported ${result.length} record(s).`,
    data: result.map((row, index) => ({
      id: row.ROWID,
      ...dataArray[index]
    }))
  };
};

const deleteAllFormMasterRecords = async (catalyst) => {
  const table = catalyst.datastore().table(TABLE_NAME);
  const allRows = await table.getAllRows();
  if (!allRows || !allRows.length) {
    return { status: 'success', message: 'No records to delete.', deletedCount: 0 };
  }
  let successCount = 0;
  for (const row of allRows) {
    if (!row.ROWID) continue;
    await table.deleteRow(row.ROWID);
    successCount++;
  }
  return {
    status: 'success',
    message: `Deleted ${successCount} record(s).`,
    deletedCount: successCount
  };
};

app.get('/', (req, res) => {
  res.status(200).json({ status: 'success', message: 'formmaster_function ready' });
});

// —— Data Store records (Act, Description, Sector, State + per-row template file) ——

app.get('/records', async (req, res) => {
  try {
    const action = (req.query.action || 'getAll').trim();
    const { catalyst } = res.locals;
    if (action !== 'getAll') {
      return res.status(400).json({ status: 'failure', message: `Unsupported action: ${action}` });
    }
    const result = await getAllFormMasterRecords(catalyst);
    res.status(200).json(result);
  } catch (err) {
    console.error('Formmaster records list error:', err);
    res.status(500).json({ status: 'failure', message: err.message || 'Failed to load records.' });
  }
});

app.post('/records', async (req, res) => {
  try {
    const action = (req.query.action || '').trim();
    const { catalyst } = res.locals;
    if (action === 'bulkImport') {
      const body = req.body;
      if (!Array.isArray(body) || !body.length) {
        return res.status(400).json({ status: 'failure', message: 'Request body must be a non-empty array.' });
      }
      const result = await bulkImportFormMasterRecords(catalyst, body);
      return res.status(200).json(result);
    }
    return res.status(400).json({ status: 'failure', message: 'Unsupported action. Use action=bulkImport.' });
  } catch (err) {
    console.error('Formmaster records POST error:', err);
    res.status(500).json({ status: 'failure', message: err.message || 'Operation failed.' });
  }
});

app.post('/records/:id/upload', async (req, res) => {
  try {
    const { id } = req.params;
    const { catalyst } = res.locals;
    if (!req.files || !req.files.file) {
      return res.status(400).json({ status: 'failure', message: 'No file uploaded.' });
    }
    const file = req.files.file;
    const name = (file.name || '').toLowerCase();
    if (!name.endsWith('.xlsx') && !name.endsWith('.xls')) {
      return res.status(400).json({ status: 'failure', message: 'Only .xlsx and .xls template files are allowed.' });
    }

    const table = catalyst.datastore().table(TABLE_NAME);
    const row = await table.getRow(id);
    if (!row) {
      return res.status(404).json({ status: 'failure', message: 'Record not found.' });
    }

    const { fileId, fileName, storageFileName, size } = await uploadToTemplatesFolder(catalyst, file, row);
    const basePayload = createFormMasterRecord({
      act: row.Act,
      description: row.Description,
      sector: row.Sector,
      state: row.State,
      formName: row.FormName,
      templateFileName: fileName,
      action: fileId,
      size: formatSizeForStore(size)
    });
    const updatePayload = { ...basePayload, ROWID: id };
    if (storageFileName) updatePayload.StorageFileName = storageFileName;

    const tryUpdate = async (payload) => table.updateRow(payload);

    try {
      await tryUpdate(updatePayload);
    } catch (updateErr) {
      const msg = String(updateErr?.message || updateErr || '');
      if (/TemplateFileName|StorageFileName|invalid|column|field/i.test(msg)) {
        const { TemplateFileName, StorageFileName, ...withoutOptional } = updatePayload;
        await tryUpdate(withoutOptional);
      } else {
        throw updateErr;
      }
    }

    const refreshed = await table.getRow(id);
    const savedAction = readRecordAction(refreshed);
    if (!savedAction) {
      throw new Error(
        'File was uploaded to File Store but the row still shows no template link. ' +
          'Ensure the FormMaster table Action column is Text and writable via API.'
      );
    }

    res.status(200).json({
      status: 'success',
      fileId: savedAction,
      fileName,
      storageFileName,
      record: convertRecordToAppFormat(refreshed),
      message: 'Template file uploaded for this row.'
    });
  } catch (err) {
    console.error('Formmaster row upload error:', err);
    res.status(500).json({ status: 'failure', message: err.message || 'Upload failed.' });
  }
});

app.delete('/records', async (req, res) => {
  try {
    const action = (req.query.action || '').trim();
    const { catalyst } = res.locals;
    if (action !== 'deleteAll') {
      return res.status(400).json({ status: 'failure', message: 'Use action=deleteAll.' });
    }
    const result = await deleteAllFormMasterRecords(catalyst);
    res.status(200).json(result);
  } catch (err) {
    console.error('Formmaster records delete error:', err);
    res.status(500).json({ status: 'failure', message: err.message || 'Delete failed.' });
  }
});

// —— File Store template list (used by Statutory / ToDoList) ——

app.get('/templates', async (req, res) => {
  try {
    const formNameFilter = (req.query.formName || '').trim().toLowerCase();
    const source = (req.query.source || 'records').trim().toLowerCase();
    const { catalyst } = res.locals;

    if (source === 'filestore') {
      let items = await listTemplateFiles(catalyst);
      if (formNameFilter) {
        items = items.filter((item) => item.name.toLowerCase().includes(formNameFilter));
      }
      return res.status(200).json({ status: 'success', data: items });
    }

    const table = catalyst.datastore().table(TABLE_NAME);
    const rows = await table.getAllRows();
    let items = (rows || [])
      .map((record) => {
        const app = convertRecordToAppFormat(record);
        const fileId = app.action;
        if (!fileId) return null;
        const templateFileName = app.templateFileName || buildDisplayTemplateFileName({
          FormName: app.formName,
          State: app.state,
          Act: app.act,
          TemplateFileName: app.templateFileName
        });
        const storageName = record.StorageFileName || templateFileName;
        return {
          id: fileId,
          name: storageName,
          templateFileName,
          formName: app.formName,
          act: app.act,
          description: app.description,
          sector: app.sector,
          state: app.state,
          templateFileName,
          recordId: app.id,
          matchKey: app.matchKey,
          size: app.size
        };
      })
      .filter(Boolean);

    if (formNameFilter) {
      items = items.filter((item) => {
        const fn = (item.formName || item.name || '').toLowerCase();
        return fn.includes(formNameFilter);
      });
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

app.post('/templates/sync', async (req, res) => {
  try {
    const { catalyst } = res.locals;
    const items = await listTemplateFiles(catalyst);
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
        FormName: (item.name || 'Unknown').replace(/\.(xlsx|xls)$/i, '').trim() || item.name || 'Unknown',
        TemplateFileName: item.name || 'Unknown',
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
