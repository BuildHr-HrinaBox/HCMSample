'use strict';

const express = require('express');
const catalystSDK = require('zcatalyst-sdk-node');
const fileUpload = require('express-fileupload');
const fs = require('fs');
const os = require('os');
const path = require('path');

const app = express();

app.use(express.json());
app.use(fileUpload());

// Attach catalyst instance per request
app.use((req, res, next) => {
  try {
    const catalyst = catalystSDK.initialize(req);
    res.locals.catalyst = catalyst;
    next();
  } catch (err) {
    res.status(500).send({ status: 'failure', message: 'Catalyst init failed' });
  }
});

// Health check
app.get('/', (req, res) => {
  res.status(200).json({ status: 'success', message: 'checklist_function ready' });
});

// Map document types to folder IDs (your Catalyst Filestore folders)
const DOC_TYPE_TO_FOLDER_ID = {
  Form: '26741000000061796',
  ProofSubmission: '26741000000061777'
};

// Upload file endpoint for checklist (Form or ProofSubmission)
app.post('/checklist/upload/:docType', async (req, res) => {
  try {
    const { docType } = req.params;
    console.log('Checklist file upload request for docType:', docType);
    console.log('Request headers:', req.headers);
    console.log('Request files:', req.files);
    
    const { catalyst } = res.locals;
    const folderId = DOC_TYPE_TO_FOLDER_ID[docType];
    console.log('Folder ID for docType:', docType, 'is:', folderId);
    
    if (!folderId) {
      console.error('Invalid document type:', docType);
      return res.status(400).json({ status: 'failure', message: `Invalid document type: ${docType}. Supported types: ${Object.keys(DOC_TYPE_TO_FOLDER_ID).join(', ')}` });
    }
    
    if (!req.files || !req.files.file) {
      console.error('No file uploaded for docType:', docType);
      return res.status(400).json({ status: 'failure', message: 'No file uploaded.' });
    }
    
    const file = req.files.file;
    console.log('File details:', {
      name: file.name,
      size: file.size,
      mimetype: file.mimetype,
      tempFilePath: file.tempFilePath
    });
    
    const tempDir = os.tmpdir();
    const tempPath = path.join(tempDir, file.name);
    
    try {
      await file.mv(tempPath);
      console.log('File moved to temp path:', tempPath);
      
      const uploadResp = await catalyst.filestore().folder(folderId).uploadFile({
        code: fs.createReadStream(tempPath),
        name: file.name
      });

      // Clean up temporary file
      if (fs.existsSync(tempPath)) {
        fs.unlinkSync(tempPath);
      }

      // Debug logging for file uploads
      console.log('File upload response:', uploadResp);

      // Defensive: handle different possible response structures
      let fileId, fileName;
      if (Array.isArray(uploadResp) && uploadResp[0]?.id) {
        fileId = uploadResp[0].id;
        fileName = file.name;
      } else if (uploadResp && uploadResp.id) {
        fileId = uploadResp.id;
        fileName = file.name;
      } else if (uploadResp && uploadResp.file_details && uploadResp.file_details[0]?.id) {
        fileId = uploadResp.file_details[0].id;
        fileName = file.name;
      } else {
        console.error('Unexpected upload response:', uploadResp);
        return res.status(500).json({ status: 'failure', message: 'File upload failed or invalid response from Catalyst.' });
      }

      // Debug logging for fileId
      console.log('FileId to be saved:', fileId);

      res.status(200).json({ status: 'success', fileId, fileName });
    } catch (uploadErr) {
      console.error('Upload error during file processing:', uploadErr);
      // Clean up temporary file in case of error
      if (fs.existsSync(tempPath)) {
        fs.unlinkSync(tempPath);
      }
      return res.status(500).json({ status: 'failure', message: uploadErr.message || 'File upload failed during processing.' });
    }
  } catch (err) {
    console.error('Upload error:', err);
    res.status(500).json({ status: 'failure', message: err.message || 'File upload failed.' });
  }
});

// Get all checklist records
app.get('/checklist', async (req, res) => {
  try {
    const { catalyst } = res.locals;
    const table = catalyst.datastore().table('Checklist');
    
    // Use getAllRows() method to get all rows
    const rows = await table.getAllRows();
    
    const checklistData = rows.map(row => ({
      id: row.ROWID,
      formName: row.FormName || '',
      act: row.Act || '',
      consultgovtdep: row.Consultgovtdep || '',
      dueDate: row.DueDate || '',
      description: row.Description || '',
      sector: row.Sector || '',
      state: row.State || '',
      formFile: row.FormFile || null,
      formFileName: row.FormFileName || null,
      proofSubmissionFile: row.ProofSubmissionFile || null,
      proofSubmissionFileName: row.ProofSubmissionFileName || null,
      complianceStatus: row.ComplianceStatus || null,
      approvalStatus: row.ApprovalStatus || null,
      marks: row.Marks || null,
      createdTime: row.CREATEDTIME,
      modifiedTime: row.MODIFIEDTIME
    }));
    
    res.status(200).json({
      status: 'success',
      data: { checklistData }
    });
  } catch (err) {
    console.error('Error fetching checklist data:', err);
    res.status(500).json({ status: 'failure', message: err.message || 'Failed to fetch checklist data.' });
  }
});

// Get single checklist record
app.get('/checklist/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const { catalyst } = res.locals;
    const table = catalyst.datastore().table('Checklist');
    
    const row = await table.getRow(id);
    
    if (!row) {
      return res.status(404).json({ status: 'failure', message: 'Checklist record not found.' });
    }
    
    const checklist = {
      id: row.ROWID,
      formName: row.FormName || '',
      act: row.Act || '',
      consultgovtdep: row.Consultgovtdep || '',
      dueDate: row.DueDate || '',
      description: row.Description || '',
      sector: row.Sector || '',
      state: row.State || '',
      formFile: row.FormFile || null,
      formFileName: row.FormFileName || null,
      proofSubmissionFile: row.ProofSubmissionFile || null,
      proofSubmissionFileName: row.ProofSubmissionFileName || null,
      complianceStatus: row.ComplianceStatus || null,
      approvalStatus: row.ApprovalStatus || null,
      marks: row.Marks || null,
      createdTime: row.CREATEDTIME,
      modifiedTime: row.MODIFIEDTIME
    };
    
    res.status(200).json({
      status: 'success',
      data: { checklist }
    });
  } catch (err) {
    console.error('Error fetching checklist record:', err);
    res.status(500).json({ status: 'failure', message: err.message || 'Failed to fetch checklist record.' });
  }
});

// Create checklist record
app.post('/checklist', async (req, res) => {
  try {
    console.log('Received request body:', JSON.stringify(req.body, null, 2));
    
    const {
      formName,
      act,
      consultgovtdep,
      dueDate,
      description,
      sector,
      state,
      formFile,
      formFileName,
      proofSubmissionFile,
      proofSubmissionFileName,
      complianceStatus,
      approvalStatus,
      marks
    } = req.body;

    // Validate required fields
    if (!formName || !String(formName).trim()) {
      return res.status(400).json({ status: 'failure', message: 'Form Name is required.' });
    }

    const { catalyst } = res.locals;
    const table = catalyst.datastore().table('Checklist');
    
    const insertData = {
      FormName: formName.trim(),
      Act: act || null,
      Consultgovtdep: consultgovtdep || null,
      DueDate: dueDate || null,
      Description: description || null,
      Sector: sector || null,
      State: state || null,
      FormFile: formFile || null,
      FormFileName: formFileName || null,
      ProofSubmissionFile: proofSubmissionFile || null,
      ProofSubmissionFileName: proofSubmissionFileName || null,
      ComplianceStatus: complianceStatus || null,
      ApprovalStatus: approvalStatus || null,
      Marks: marks || null
    };
    
    console.log('Inserting data to Checklist table:', JSON.stringify(insertData, null, 2));
    
    const insertResp = await table.insertRow(insertData);
    console.log('Insert response:', JSON.stringify(insertResp, null, 2));
    
    const created = await table.getRow(insertResp.ROWID);
    console.log('Created record:', JSON.stringify(created, null, 2));
    
    // Map to frontend expected format
    const mappedRecord = {
      id: created.ROWID,
      formName: created.FormName || '',
      act: created.Act || '',
      consultgovtdep: created.Consultgovtdep || '',
      dueDate: created.DueDate || '',
      description: created.Description || '',
      sector: created.Sector || '',
      state: created.State || '',
      formFile: created.FormFile || null,
      formFileName: created.FormFileName || null,
      proofSubmissionFile: created.ProofSubmissionFile || null,
      proofSubmissionFileName: created.ProofSubmissionFileName || null,
      complianceStatus: created.ComplianceStatus || null,
      approvalStatus: created.ApprovalStatus || null,
      marks: created.Marks || null,
      createdTime: created.CREATEDTIME,
      modifiedTime: created.MODIFIEDTIME
    };
    
    res.status(200).json({ status: 'success', data: { checklist: mappedRecord } });
  } catch (err) {
    console.error('Error creating checklist record:', err);
    res.status(400).json({ status: 'failure', message: err.message || 'Failed to create checklist record.' });
  }
});

// Update checklist record
app.put('/checklist/:id', async (req, res) => {
  try {
    const { id } = req.params;
    console.log('Received update request for ID:', id);
    console.log('Request body:', JSON.stringify(req.body, null, 2));
    
    const {
      formName,
      act,
      consultgovtdep,
      dueDate,
      description,
      sector,
      state,
      formFile,
      formFileName,
      proofSubmissionFile,
      proofSubmissionFileName,
      complianceStatus,
      approvalStatus,
      marks
    } = req.body;

    // Validate required fields
    if (!formName || !String(formName).trim()) {
      return res.status(400).json({ status: 'failure', message: 'Form Name is required.' });
    }

    const { catalyst } = res.locals;
    const table = catalyst.datastore().table('Checklist');
    
    const updateData = {
      ROWID: id,
      FormName: formName.trim(),
      Act: act || null,
      Consultgovtdep: consultgovtdep || null,
      DueDate: dueDate || null,
      Description: description || null,
      Sector: sector || null,
      State: state || null,
      FormFile: formFile || null,
      FormFileName: formFileName || null,
      ProofSubmissionFile: proofSubmissionFile || null,
      ProofSubmissionFileName: proofSubmissionFileName || null,
      ComplianceStatus: complianceStatus || null,
      ApprovalStatus: approvalStatus || null,
      Marks: marks || null
    };
    
    console.log('Updating Checklist record:', JSON.stringify(updateData, null, 2));
    
    await table.updateRow(updateData);
    
    const updated = await table.getRow(id);
    console.log('Updated record:', JSON.stringify(updated, null, 2));
    
    // Map to frontend expected format
    const mappedRecord = {
      id: updated.ROWID,
      formName: updated.FormName || '',
      act: updated.Act || '',
      consultgovtdep: updated.Consultgovtdep || '',
      dueDate: updated.DueDate || '',
      description: updated.Description || '',
      sector: updated.Sector || '',
      state: updated.State || '',
      formFile: updated.FormFile || null,
      formFileName: updated.FormFileName || null,
      proofSubmissionFile: updated.ProofSubmissionFile || null,
      proofSubmissionFileName: updated.ProofSubmissionFileName || null,
      complianceStatus: updated.ComplianceStatus || null,
      approvalStatus: updated.ApprovalStatus || null,
      marks: updated.Marks || null,
      createdTime: updated.CREATEDTIME,
      modifiedTime: updated.MODIFIEDTIME
    };
    
    res.status(200).json({ status: 'success', data: { checklist: mappedRecord } });
  } catch (err) {
    console.error('Error updating checklist record:', err);
    res.status(400).json({ status: 'failure', message: err.message || 'Failed to update checklist record.' });
  }
});

// Delete checklist record
app.delete('/checklist/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const { catalyst } = res.locals;
    const table = catalyst.datastore().table('Checklist');
    
    await table.deleteRow(id);
    
    res.status(200).json({ status: 'success', data: { id } });
  } catch (err) {
    console.error('Error deleting checklist record:', err);
    res.status(500).json({ status: 'failure', message: err.message || 'Failed to delete checklist record.' });
  }
});

// Download file endpoint
app.get('/checklist/:id/file/:docType', async (req, res) => {
  try {
    const executionStartTime = Date.now();
    console.log('Execution started at:', executionStartTime);
    
    const { id, docType } = req.params;
    console.log('File download request:', { id, docType });
    const { catalyst } = res.locals;
    const folderId = DOC_TYPE_TO_FOLDER_ID[docType];
    console.log('Folder ID for docType:', docType, 'is:', folderId);
    
    if (!folderId) {
      console.error('Invalid document type:', docType);
      return res.status(400).json({ status: 'failure', message: 'Invalid document type.' });
    }
    
    const table = catalyst.datastore().table('Checklist');
    let checklist;
    try {
      checklist = await table.getRow(id);
    } catch (getRowErr) {
      console.error('Error fetching checklist record:', getRowErr);
      return res.status(404).json({ status: 'failure', message: 'Checklist record not found.' });
    }
    
    if (!checklist) {
      console.error('Checklist record not found for ID:', id);
      return res.status(404).json({ status: 'failure', message: 'Checklist record not found.' });
    }
    
    let fileId, fileName;
    if (docType === 'Form') {
      fileId = checklist.FormFile;
      fileName = checklist.FormFileName || 'form_file';
    } else if (docType === 'ProofSubmission') {
      fileId = checklist.ProofSubmissionFile;
      fileName = checklist.ProofSubmissionFileName || 'proof_submission_file';
    } else {
      return res.status(400).json({ status: 'failure', message: 'Invalid document type.' });
    }
    
    console.log('File info:', { fileId, fileName });
    
    if (!fileId) {
      console.error('No file ID found for checklist:', id, 'docType:', docType);
      return res.status(404).json({ status: 'failure', message: 'File not found for this checklist record.' });
    }
    
    let fileBuffer;
    try {
      fileBuffer = await catalyst.filestore().folder(folderId).downloadFile(fileId);
    } catch (downloadErr) {
      console.error('Error downloading file from filestore:', downloadErr);
      
      // Check if it's a file not found error
      if (downloadErr.statusCode === 404 || downloadErr.code === 'INVALID_ID' || 
          (downloadErr.message && downloadErr.message.includes('No such file'))) {
        console.error(`File with ID ${fileId} does not exist in folder ${folderId}`);
        return res.status(404).json({ 
          status: 'failure', 
          message: `File "${fileName}" no longer exists in storage. The file may have been deleted or moved. Please re-upload the file.` 
        });
      }
      
      // Re-throw other errors to be caught by outer catch
      throw downloadErr;
    }
    
    // Determine content type based on file extension
    const getContentType = (filename) => {
      const ext = filename.toLowerCase().split('.').pop();
      const contentTypes = {
        'pdf': 'application/pdf',
        'jpg': 'image/jpeg',
        'jpeg': 'image/jpeg',
        'png': 'image/png',
        'gif': 'image/gif',
        'webp': 'image/webp',
        'doc': 'application/msword',
        'docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        'xls': 'application/vnd.ms-excel',
        'xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'txt': 'text/plain'
      };
      return contentTypes[ext] || 'application/octet-stream';
    };
    
    const contentType = getContentType(fileName);
    
    // Determine disposition (inline or attachment) from query parameter, default to 'inline'
    const disposition = req.query.disposition === 'attachment' ? 'attachment' : 'inline';
    
    console.log('Sending file:', { fileName, contentType, fileSize: fileBuffer.length, disposition });
    res.writeHead(200, {
      'Content-Type': contentType,
      'Content-Disposition': `${disposition}; filename="${fileName}"`,
      'Content-Length': fileBuffer.length
    });
    res.end(fileBuffer);
  } catch (err) {
    console.error('Error downloading file:', err);
    
    // Check if it's a file not found error that wasn't caught earlier
    if (err.statusCode === 404 || err.code === 'INVALID_ID' || 
        (err.message && err.message.includes('No such file'))) {
      return res.status(404).json({ 
        status: 'failure', 
        message: `File no longer exists in storage. The file may have been deleted or moved. Please re-upload the file.` 
      });
    }
    
    res.status(500).json({ status: 'failure', message: err.message || 'File download failed.' });
  }
});

// Delete file endpoint for checklist and docType
app.delete('/checklist/:id/file/:docType', async (req, res) => {
  try {
    const { id, docType } = req.params;
    const { catalyst } = res.locals;
    const folderId = DOC_TYPE_TO_FOLDER_ID[docType];
    
    if (!folderId) {
      return res.status(400).json({ status: 'failure', message: 'Invalid document type.' });
    }
    
    const table = catalyst.datastore().table('Checklist');
    const checklist = await table.getRow(id);
    
    let fileId;
    if (docType === 'Form') {
      fileId = checklist.FormFile;
    } else if (docType === 'ProofSubmission') {
      fileId = checklist.ProofSubmissionFile;
    } else {
      return res.status(400).json({ status: 'failure', message: 'Invalid document type.' });
    }
    
    if (!fileId) {
      return res.status(404).json({ status: 'failure', message: 'File not found for this checklist record.' });
    }
    
    // Delete file from file store
    await catalyst.filestore().folder(folderId).deleteFile(fileId);
    
    // Clear the appropriate column in the database
    const updateData = { ROWID: id };
    if (docType === 'Form') {
      updateData.FormFile = null;
      updateData.FormFileName = null;
    } else if (docType === 'ProofSubmission') {
      updateData.ProofSubmissionFile = null;
      updateData.ProofSubmissionFileName = null;
    }
    await table.updateRow(updateData);
    
    res.status(200).json({ status: 'success' });
  } catch (err) {
    console.log(err);
    res.status(500).json({ status: 'failure', message: err.message || 'File delete failed.' });
  }
});

module.exports = app;

