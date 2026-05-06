'use strict';

const express = require('express');
const catalystSDK = require('zcatalyst-sdk-node');
const fileUpload = require('express-fileupload');
const fs = require('fs');
const os = require('os');
const path = require('path');

const TABLE_NAME = 'ToDoList';
/** FormFile folder in Catalyst File Store (HCM project) */
const FORM_FILE_FOLDER_ID = '26741000000347859';
/** ProofSubmissionFile folder in Catalyst File Store (HCM project) */
const PROOF_SUBMISSION_FOLDER_ID = '26741000000347952';

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

app.get('/', (req, res) => {
  res.status(200).json({ status: 'success', message: 'todolist_function ready' });
});

function mapRow(row) {
  return {
    id: row.ROWID,
    act: row.Act || '',
    description: row.Description || '',
    dueDate: row.DueDate || '',
    formName: row.FormName || '',
    formFile: row.FormFile || null,
    formFileName: row.FormFileName || null,
    proofSubmissionFile: row.ProofSubmissionFile || null,
    proofSubmissionFileName: row.ProofSubmissionFileName || null,
    creatorId: row.CREATORID ?? null,
    createdTime: row.CREATEDTIME,
    modifiedTime: row.MODIFIEDTIME
  };
}

function getFolderIdForDocType(docType) {
  if (docType === 'ProofSubmissionFile') return PROOF_SUBMISSION_FOLDER_ID;
  return FORM_FILE_FOLDER_ID;
}

async function uploadToFolder(catalyst, folderId, file) {
  const tempPath = path.join(os.tmpdir(), file.name);
  try {
    await file.mv(tempPath);
    const uploadResp = await catalyst.filestore().folder(folderId).uploadFile({
      code: fs.createReadStream(tempPath),
      name: file.name
    });
    if (fs.existsSync(tempPath)) fs.unlinkSync(tempPath);

    let fileId;
    if (Array.isArray(uploadResp) && uploadResp[0]?.id) {
      fileId = uploadResp[0].id;
    } else if (uploadResp?.id) {
      fileId = uploadResp.id;
    } else if (uploadResp?.file_details?.[0]?.id) {
      fileId = uploadResp.file_details[0].id;
    } else {
      throw new Error('File upload failed or invalid response from Catalyst.');
    }
    return { fileId, fileName: file.name };
  } catch (uploadErr) {
    if (fs.existsSync(tempPath)) fs.unlinkSync(tempPath);
    throw uploadErr;
  }
}

app.post('/todos/upload/:docType', async (req, res) => {
  try {
    const { docType } = req.params;
    const folderId = getFolderIdForDocType(docType);
    const { catalyst } = res.locals;
    if (!req.files || !req.files.file) {
      return res.status(400).json({ status: 'failure', message: 'No file uploaded.' });
    }
    const file = req.files.file;
    const uploaded = await uploadToFolder(catalyst, folderId, file);
    res.status(200).json({ status: 'success', ...uploaded });
  } catch (err) {
    console.error('todolist upload:', err);
    res.status(500).json({ status: 'failure', message: err.message || 'File upload failed.' });
  }
});

// Backward-compatible: existing endpoint uploads to FormFile folder
app.post('/todos/upload', async (req, res) => {
  try {
    const { catalyst } = res.locals;
    if (!req.files || !req.files.file) {
      return res.status(400).json({ status: 'failure', message: 'No file uploaded.' });
    }
    const file = req.files.file;
    const uploaded = await uploadToFolder(catalyst, FORM_FILE_FOLDER_ID, file);
    res.status(200).json({ status: 'success', ...uploaded });
  } catch (err) {
    console.error('todolist upload:', err);
    res.status(500).json({ status: 'failure', message: err.message || 'File upload failed.' });
  }
});

app.get('/todos', async (req, res) => {
  try {
    const { catalyst } = res.locals;
    const table = catalyst.datastore().table(TABLE_NAME);
    const rows = await table.getAllRows();
    const todos = rows.map(mapRow);
    res.status(200).json({ status: 'success', data: { todos } });
  } catch (err) {
    console.error('todolist GET /todos:', err);
    res.status(500).json({ status: 'failure', message: err.message || 'Failed to load to-dos.' });
  }
});

app.get('/todos/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const { catalyst } = res.locals;
    const table = catalyst.datastore().table(TABLE_NAME);
    const row = await table.getRow(id);
    if (!row) {
      return res.status(404).json({ status: 'failure', message: 'To-do not found.' });
    }
    res.status(200).json({ status: 'success', data: { todo: mapRow(row) } });
  } catch (err) {
    console.error('todolist GET /todos/:id:', err);
    res.status(500).json({ status: 'failure', message: err.message || 'Failed to load to-do.' });
  }
});

app.post('/todos', async (req, res) => {
  try {
    const {
      act, description, dueDate, formName, formFile, formFileName,
      proofSubmissionFile, proofSubmissionFileName
    } = req.body;
    if (!act || !String(act).trim()) {
      return res.status(400).json({ status: 'failure', message: 'Act is required.' });
    }
    const { catalyst } = res.locals;
    const table = catalyst.datastore().table(TABLE_NAME);
    const insertData = {
      Act: act.trim(),
      Description: description != null && String(description).trim() ? String(description).trim() : null,
      DueDate: dueDate != null && String(dueDate).trim() ? String(dueDate).trim() : null,
      FormName: formName != null && String(formName).trim() ? String(formName).trim() : null,
      FormFile: formFile || null,
      FormFileName: formFileName || null,
      ProofSubmissionFile: proofSubmissionFile || null,
      ProofSubmissionFileName: proofSubmissionFileName || null
    };
    const insertResp = await table.insertRow(insertData);
    const created = await table.getRow(insertResp.ROWID);
    res.status(200).json({ status: 'success', data: { todo: mapRow(created) } });
  } catch (err) {
    console.error('todolist POST /todos:', err);
    res.status(400).json({ status: 'failure', message: err.message || 'Failed to create to-do.' });
  }
});

app.put('/todos/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const {
      act, description, dueDate, formName, formFile, formFileName,
      proofSubmissionFile, proofSubmissionFileName
    } = req.body;
    if (!act || !String(act).trim()) {
      return res.status(400).json({ status: 'failure', message: 'Act is required.' });
    }
    const { catalyst } = res.locals;
    const table = catalyst.datastore().table(TABLE_NAME);
    const existing = await table.getRow(id);
    if (!existing) {
      return res.status(404).json({ status: 'failure', message: 'To-do not found.' });
    }
    const nextFile = formFile !== undefined ? formFile : existing.FormFile;
    const nextFileName = formFileName !== undefined ? formFileName : existing.FormFileName;
    const nextProofFile = proofSubmissionFile !== undefined ? proofSubmissionFile : existing.ProofSubmissionFile;
    const nextProofFileName = proofSubmissionFileName !== undefined ? proofSubmissionFileName : existing.ProofSubmissionFileName;
    const updateData = {
      ROWID: id,
      Act: act.trim(),
      Description: description != null && String(description).trim() ? String(description).trim() : null,
      DueDate: dueDate != null && String(dueDate).trim() ? String(dueDate).trim() : null,
      FormName: formName != null && String(formName).trim() ? String(formName).trim() : null,
      FormFile: nextFile || null,
      FormFileName: nextFileName || null,
      ProofSubmissionFile: nextProofFile || null,
      ProofSubmissionFileName: nextProofFileName || null
    };
    await table.updateRow(updateData);
    const updated = await table.getRow(id);
    res.status(200).json({ status: 'success', data: { todo: mapRow(updated) } });
  } catch (err) {
    console.error('todolist PUT /todos/:id:', err);
    res.status(400).json({ status: 'failure', message: err.message || 'Failed to update to-do.' });
  }
});

app.delete('/todos/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const { catalyst } = res.locals;
    const table = catalyst.datastore().table(TABLE_NAME);
    let row;
    try {
      row = await table.getRow(id);
    } catch (_) {
      row = null;
    }
    if (row?.FormFile) {
      try {
        await catalyst.filestore().folder(FORM_FILE_FOLDER_ID).deleteFile(row.FormFile);
      } catch (fileErr) {
        console.warn('todolist: could not delete file for row', id, fileErr?.message);
      }
    }
    if (row?.ProofSubmissionFile) {
      try {
        await catalyst.filestore().folder(PROOF_SUBMISSION_FOLDER_ID).deleteFile(row.ProofSubmissionFile);
      } catch (fileErr) {
        console.warn('todolist: could not delete proof file for row', id, fileErr?.message);
      }
    }
    await table.deleteRow(id);
    res.status(200).json({ status: 'success', data: { id } });
  } catch (err) {
    console.error('todolist DELETE /todos/:id:', err);
    res.status(500).json({ status: 'failure', message: err.message || 'Failed to delete to-do.' });
  }
});

app.get('/todos/:id/file/:docType?', async (req, res) => {
  try {
    const { id } = req.params;
    const docType = req.params.docType || 'FormFile';
    const { catalyst } = res.locals;
    const table = catalyst.datastore().table(TABLE_NAME);
    let row;
    try {
      row = await table.getRow(id);
    } catch (getRowErr) {
      console.error(getRowErr);
      return res.status(404).json({ status: 'failure', message: 'To-do not found.' });
    }
    if (!row) {
      return res.status(404).json({ status: 'failure', message: 'To-do not found.' });
    }
    const isProof = docType === 'ProofSubmissionFile';
    const folderId = isProof ? PROOF_SUBMISSION_FOLDER_ID : FORM_FILE_FOLDER_ID;
    const fileId = isProof ? row.ProofSubmissionFile : row.FormFile;
    const fileName = isProof
      ? (row.ProofSubmissionFileName || 'proof_submission')
      : (row.FormFileName || 'attachment');
    if (!fileId) {
      return res.status(404).json({ status: 'failure', message: 'No file attached to this to-do.' });
    }
    let fileBuffer;
    try {
      fileBuffer = await catalyst.filestore().folder(folderId).downloadFile(fileId);
    } catch (downloadErr) {
      if (downloadErr.statusCode === 404 || downloadErr.code === 'INVALID_ID') {
        return res.status(404).json({
          status: 'failure',
          message: 'File no longer exists in storage. Re-upload if needed.'
        });
      }
      throw downloadErr;
    }
    const ext = String(fileName).toLowerCase().split('.').pop();
    const contentTypes = {
      pdf: 'application/pdf',
      jpg: 'image/jpeg',
      jpeg: 'image/jpeg',
      png: 'image/png',
      gif: 'image/gif',
      webp: 'image/webp',
      doc: 'application/msword',
      docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      xls: 'application/vnd.ms-excel',
      xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      txt: 'text/plain'
    };
    const contentType = contentTypes[ext] || 'application/octet-stream';
    const disposition = req.query.disposition === 'attachment' ? 'attachment' : 'inline';
    res.writeHead(200, {
      'Content-Type': contentType,
      'Content-Disposition': `${disposition}; filename="${fileName}"`,
      'Content-Length': fileBuffer.length
    });
    res.end(fileBuffer);
  } catch (err) {
    console.error('todolist file download:', err);
    res.status(500).json({ status: 'failure', message: err.message || 'File download failed.' });
  }
});

app.delete('/todos/:id/file/:docType?', async (req, res) => {
  try {
    const { id } = req.params;
    const docType = req.params.docType || 'FormFile';
    const isProof = docType === 'ProofSubmissionFile';
    const folderId = isProof ? PROOF_SUBMISSION_FOLDER_ID : FORM_FILE_FOLDER_ID;
    const { catalyst } = res.locals;
    const table = catalyst.datastore().table(TABLE_NAME);
    const row = await table.getRow(id);
    const targetFileId = isProof ? row?.ProofSubmissionFile : row?.FormFile;
    if (!targetFileId) {
      return res.status(404).json({ status: 'failure', message: 'No file attached.' });
    }
    await catalyst.filestore().folder(folderId).deleteFile(targetFileId);
    if (isProof) {
      await table.updateRow({ ROWID: id, ProofSubmissionFile: null, ProofSubmissionFileName: null });
    } else {
      await table.updateRow({ ROWID: id, FormFile: null, FormFileName: null });
    }
    res.status(200).json({ status: 'success' });
  } catch (err) {
    console.error('todolist file delete:', err);
    res.status(500).json({ status: 'failure', message: err.message || 'File delete failed.' });
  }
});

module.exports = app;
