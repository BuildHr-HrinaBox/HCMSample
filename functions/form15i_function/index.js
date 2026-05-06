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
  res.status(200).json({ status: 'success', message: 'form15I_function ready' });
});

// Test file upload endpoint
app.get('/test-upload', (req, res) => {
  res.status(200).json({ 
    status: 'success', 
    message: 'Upload endpoint is accessible',
    docTypes: Object.keys(DOC_TYPE_TO_FOLDER_ID),
    folderIds: DOC_TYPE_TO_FOLDER_ID
  });
});

// Test endpoint to check table access
app.get('/test-table', async (req, res) => {
  try {
    console.log('Testing table access...');
    const { catalyst } = res.locals;
    const table = catalyst.datastore().table('Form15I');
    
    // Try to get table info
    console.log('Table object:', table);
    
    // Try a simple query
    const zcql = catalyst.zcql();
    const testQuery = await zcql.executeZCQLQuery('SELECT COUNT(ROWID) as count FROM Form15I');
    console.log('Test query result:', testQuery);
    
    res.status(200).send({ 
      status: 'success', 
      message: 'Table access successful',
      tableInfo: {
        tableName: 'Form15I',
        recordCount: testQuery[0]?.Form15I?.count || 0
      }
    });
  } catch (err) {
    console.error('Table test error:', err);
    res.status(500).send({ 
      status: 'failure', 
      message: err.message || 'Failed to access table',
      error: err.toString()
    });
  }
});

// Test endpoint to check File Store access
app.get('/test-filestore', async (req, res) => {
  try {
    console.log('Testing File Store access...');
    const { catalyst } = res.locals;
    const filestore = catalyst.filestore();
    
    // Try to access the Photo folder
    const folder = filestore.folder('Photo');
    console.log('Folder object:', folder);
    
    // Try to list files in the folder
    const files = await folder.getAllFiles();
    console.log('Files in Photo folder:', files);
    
    res.status(200).send({ 
      status: 'success', 
      message: 'File Store access successful',
      fileStoreInfo: {
        folderName: 'Photo',
        fileCount: files.length || 0,
        files: files.map(f => ({ name: f.name, id: f.id }))
      }
    });
  } catch (err) {
    console.error('File Store test error:', err);
    res.status(500).send({ 
      status: 'failure', 
      message: err.message || 'Failed to access File Store',
      error: err.toString()
    });
  }
});

// Helper: parse integer-like strings to number or null
function parseNumeric(value) {
  if (value == null) return null;
  if (typeof value === 'number') return value;
  if (typeof value === 'string') {
    const trimmed = value.trim();
    if (trimmed === '') return null;
    const parsed = parseInt(trimmed, 10);
    return Number.isNaN(parsed) ? null : parsed;
  }
  const str = String(value).trim();
  if (str === '') return null;
  const parsed = parseInt(str, 10);
  return Number.isNaN(parsed) ? null : parsed;
}

// ===== Form15I CRUD =====
// Table: Form15I (columns: RegisterofAdultworkersandyoungPersons, NameoftheWorker, WorkerIdentityNumber, LeaveatthebeginningoftheMonth, LeaveearnedduringthePeriod, LeaveavailedduringtheMonth, LeavebalanceattheendoftheMonth, MedicalLeaveatbeginningoftheMonth, MedicalLeaveavailedduringtheMonth, LeavebalanceatendoftheaMonth, OtherLeaveatbeginningoftheMonth, OtherLeaveavailedduringtheMonth, OtherLeaveBalanceatendoftheMonth, Dateofgivingnoticeofpregnancy, AmountofMatermity, Subsequentpayment, Amountpaid, LeavewithWages, AmountpaidasGratuity, Remarks)

// Mapping of document types to Catalyst File Store folder IDs
const DOC_TYPE_TO_FOLDER_ID = {
  Photo: '28341000000052337' // Your Photo folder ID from the File Store
};

// Helper to get file columns for a docType
function getFileColumns(docType) {
  return {
    fileIdCol: `${docType}FileId`,
    fileNameCol: `${docType}FileName`
  };
}

// Helper function to upload file to File Store (for base64 data from frontend)
async function uploadFileToFileStore(catalyst, fileData, fileName) {
  try {
    if (!fileData || !fileName) {
      return null;
    }

    // Check if it's a base64 data URL
    if (fileData.startsWith('data:')) {
      const base64Data = fileData.split(',')[1];
      const buffer = Buffer.from(base64Data, 'base64');
      
      // Get file extension from data URL
      const mimeType = fileData.split(';')[0].split(':')[1];
      const extension = mimeType.split('/')[1] || 'png';
      const finalFileName = `${fileName}_${Date.now()}.${extension}`;
      
      const filestore = catalyst.filestore();
      const folder = filestore.folder('Photo'); // Use the Photo folder you created
      
      const uploadResponse = await folder.uploadFile({
        name: finalFileName,
        content: buffer
      });
      
      console.log('File upload response:', uploadResponse);
      return uploadResponse.downloadUrl || uploadResponse.fileId;
    }
    
    return fileData; // Return as-is if not base64
  } catch (err) {
    console.error('Error uploading file:', err);
    return null;
  }
}

// Create Form15I
app.post('/form15I', async (req, res) => {
  try {
    console.log('Received request body:', JSON.stringify(req.body, null, 2));
    
    const {
      registerOfAdultWorkersAndYoungPersons,
      nameOfWorker,
      workerIdentityNumber,
      leaveAtBeginningOfMonth,
      leaveEarnedDuringPeriod,
      leaveAvailedDuringMonth,
      leaveBalanceAtEndOfMonth,
      medicalLeaveAtBeginningOfMonth,
      medicalLeaveAvailedDuringMonth,
      otherLeaveAtBeginningOfMonth,
      otherLeaveAvailedDuringMonth,
      otherLeaveBalanceAtEndOfMonth,
      dateOfGivingNoticeOfPregnancy,
      amountOfMaternity,
      subsequentPayment,
      amountPaid,
      leaveWithWages,
      amountPaidAsGratuity,
      remarks
    } = req.body;

    // Validate required fields
    if (!registerOfAdultWorkersAndYoungPersons || !String(registerOfAdultWorkersAndYoungPersons).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Register of Adult Workers and Young Persons is required.' });
    }
    if (!nameOfWorker || !String(nameOfWorker).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Name of the Worker is required.' });
    }
    if (!workerIdentityNumber || !String(workerIdentityNumber).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Worker Identity Number is required.' });
    }

    const { catalyst } = res.locals;
    const table = catalyst.datastore().table('Form15I');
    
    const insertData = {
      RegisterofAdultworkersandyoungPersons: registerOfAdultWorkersAndYoungPersons,
      NameoftheWorker: nameOfWorker,
      WorkerIdentityNumber: workerIdentityNumber,
      LeaveatthebeginningoftheMonth: parseNumeric(leaveAtBeginningOfMonth) || null,
      LeaveearnedduringthePeriod: parseNumeric(leaveEarnedDuringPeriod) || null,
      LeaveavailedduringtheMonth: parseNumeric(leaveAvailedDuringMonth) || null,
      LeavebalanceattheendoftheMonth: parseNumeric(leaveBalanceAtEndOfMonth) || null,
      MedicalLeaveatbeginningoftheMonth: parseNumeric(medicalLeaveAtBeginningOfMonth) || null,
      MedicalLeaveavailedduringtheMonth: parseNumeric(medicalLeaveAvailedDuringMonth) || null,
      OtherLeaveatbeginningoftheMonth: parseNumeric(otherLeaveAtBeginningOfMonth) || null,
      OtherLeaveavailedduringtheMonth: parseNumeric(otherLeaveAvailedDuringMonth) || null,
      OtherLeaveBalanceatendoftheMonth: parseNumeric(otherLeaveBalanceAtEndOfMonth) || null,
      Dateofgivingnoticeofpregnancy: dateOfGivingNoticeOfPregnancy || null,
      AmountofMatermity: parseNumeric(amountOfMaternity) || null,
      Subsequentpayment: parseNumeric(subsequentPayment) || null,
      Amountpaid: parseNumeric(amountPaid) || null,
      LeavewithWages: leaveWithWages || null,
      AmountpaidasGratuity: parseNumeric(amountPaidAsGratuity) || null,
      Remarks: remarks || null
    };
    
    console.log('Inserting data to Form15I table:', JSON.stringify(insertData, null, 2));
    
    const insertResp = await table.insertRow(insertData);
    console.log('Insert response:', JSON.stringify(insertResp, null, 2));
    
    const created = await table.getRow(insertResp.ROWID);
    console.log('Created record:', JSON.stringify(created, null, 2));
    
    res.status(200).send({ status: 'success', data: { form15I: created } });
  } catch (err) {
    console.error('Error creating form15I record:', err);
    res.status(400).send({ status: 'failure', message: err.message || 'Failed to create form15I record.' });
  }
});

// Get single Form15I record with all details
app.get('/form15I/:ROWID', async (req, res) => {
  try {
    const { ROWID } = req.params;
    console.log('Fetching form15I record:', ROWID);
    const { catalyst } = res.locals;
    const table = catalyst.datastore().table('Form15I');
    
    const form15I = await table.getRow(ROWID);
    console.log('Retrieved form15I record:', JSON.stringify(form15I, null, 2));
    
    res.status(200).send({ status: 'success', data: { form15I } });
  } catch (err) {
    console.error('Error fetching form15I record:', err);
    res.status(500).send({ status: 'failure', message: err.message || 'Failed to fetch form15I record.' });
  }
});

// List Form15I with optional pagination
app.get('/form15I', async (req, res) => {
  try {
    console.log('Fetching form15I records...');
    const { catalyst } = res.locals;
    const zcql = catalyst.zcql();
    const page = parseInt(req.query.page, 10) || 1;
    const perPage = parseInt(req.query.perPage, 10) || 50;
    const returnAll = !req.query.page && !req.query.perPage;
    const limitClause = returnAll ? '' : `LIMIT ${(page - 1) * perPage + 1},${perPage}`;
    
    console.log('Executing count query...');
    const countRows = await zcql.executeZCQLQuery('SELECT COUNT(ROWID) as count FROM Form15I');
    const total = parseInt(countRows[0].Form15I.count, 10) || 0;
    console.log('Total records:', total);
    
    console.log('Executing data query...');
    // Select all columns for Form15I
    const rows = await zcql.executeZCQLQuery(`SELECT ROWID, RegisterofAdultworkersandyoungPersons, NameoftheWorker, WorkerIdentityNumber, LeaveatthebeginningoftheMonth, LeaveearnedduringthePeriod, LeaveavailedduringtheMonth, LeavebalanceattheendoftheMonth, MedicalLeaveatbeginningoftheMonth, MedicalLeaveavailedduringtheMonth, OtherLeaveatbeginningoftheMonth, OtherLeaveavailedduringtheMonth, OtherLeaveBalanceatendoftheMonth, Dateofgivingnoticeofpregnancy, AmountofMatermity, Subsequentpayment, Amountpaid, LeavewithWages, AmountpaidasGratuity, Remarks, CREATEDTIME FROM Form15I ORDER BY ROWID DESC ${limitClause}`);
    console.log('Raw rows from database:', JSON.stringify(rows, null, 2));
    
    const form15IData = rows.map(r => ({
      id: r.Form15I.ROWID,
      registerOfAdultWorkersAndYoungPersons: r.Form15I.RegisterofAdultworkersandyoungPersons,
      nameOfWorker: r.Form15I.NameoftheWorker,
      workerIdentityNumber: r.Form15I.WorkerIdentityNumber,
      leaveAtBeginningOfMonth: r.Form15I.LeaveatthebeginningoftheMonth,
      leaveEarnedDuringPeriod: r.Form15I.LeaveearnedduringthePeriod,
      leaveAvailedDuringMonth: r.Form15I.LeaveavailedduringtheMonth,
      leaveBalanceAtEndOfMonth: r.Form15I.LeavebalanceattheendoftheMonth,
      medicalLeaveAtBeginningOfMonth: r.Form15I.MedicalLeaveatbeginningoftheMonth,
      medicalLeaveAvailedDuringMonth: r.Form15I.MedicalLeaveavailedduringtheMonth,
      otherLeaveAtBeginningOfMonth: r.Form15I.OtherLeaveatbeginningoftheMonth,
      otherLeaveAvailedDuringMonth: r.Form15I.OtherLeaveavailedduringtheMonth,
      otherLeaveBalanceAtEndOfMonth: r.Form15I.OtherLeaveBalanceatendoftheMonth,
      dateOfGivingNoticeOfPregnancy: r.Form15I.Dateofgivingnoticeofpregnancy,
      amountOfMaternity: r.Form15I.AmountofMatermity,
      subsequentPayment: r.Form15I.Subsequentpayment,
      amountPaid: r.Form15I.Amountpaid,
      leaveWithWages: r.Form15I.LeavewithWages,
      amountPaidAsGratuity: r.Form15I.AmountpaidasGratuity,
      remarks: r.Form15I.Remarks,
      createdTime: r.Form15I.CREATEDTIME,
      modifiedTime: null // Not selected in query
    }));
    
    console.log('Processed form15I data:', JSON.stringify(form15IData, null, 2));
    res.status(200).send({ status: 'success', data: { form15IData, total, hasMore: returnAll ? false : page * perPage < total } });
  } catch (err) {
    console.error('Error fetching form15I records:', err);
    res.status(500).send({ status: 'failure', message: err.message || 'Failed to fetch form15I records.' });
  }
});

// Update Form15I
app.put('/form15I/:ROWID', async (req, res) => {
  try {
    const { ROWID } = req.params;
    const {
      registerOfAdultWorkersAndYoungPersons,
      nameOfWorker,
      workerIdentityNumber,
      leaveAtBeginningOfMonth,
      leaveEarnedDuringPeriod,
      leaveAvailedDuringMonth,
      leaveBalanceAtEndOfMonth,
      medicalLeaveAtBeginningOfMonth,
      medicalLeaveAvailedDuringMonth,
      otherLeaveAtBeginningOfMonth,
      otherLeaveAvailedDuringMonth,
      otherLeaveBalanceAtEndOfMonth,
      dateOfGivingNoticeOfPregnancy,
      amountOfMaternity,
      subsequentPayment,
      amountPaid,
      leaveWithWages,
      amountPaidAsGratuity,
      remarks
    } = req.body;

    // Validate required fields
    if (!registerOfAdultWorkersAndYoungPersons || !String(registerOfAdultWorkersAndYoungPersons).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Register of Adult Workers and Young Persons is required.' });
    }
    if (!nameOfWorker || !String(nameOfWorker).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Name of the Worker is required.' });
    }
    if (!workerIdentityNumber || !String(workerIdentityNumber).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Worker Identity Number is required.' });
    }

    const { catalyst } = res.locals;
    const table = catalyst.datastore().table('Form15I');
    await table.updateRow({
      ROWID,
      RegisterofAdultworkersandyoungPersons: registerOfAdultWorkersAndYoungPersons,
      NameoftheWorker: nameOfWorker,
      WorkerIdentityNumber: workerIdentityNumber,
      LeaveatthebeginningoftheMonth: parseNumeric(leaveAtBeginningOfMonth) || null,
      LeaveearnedduringthePeriod: parseNumeric(leaveEarnedDuringPeriod) || null,
      LeaveavailedduringtheMonth: parseNumeric(leaveAvailedDuringMonth) || null,
      LeavebalanceattheendoftheMonth: parseNumeric(leaveBalanceAtEndOfMonth) || null,
      MedicalLeaveatbeginningoftheMonth: parseNumeric(medicalLeaveAtBeginningOfMonth) || null,
      MedicalLeaveavailedduringtheMonth: parseNumeric(medicalLeaveAvailedDuringMonth) || null,
      OtherLeaveatbeginningoftheMonth: parseNumeric(otherLeaveAtBeginningOfMonth) || null,
      OtherLeaveavailedduringtheMonth: parseNumeric(otherLeaveAvailedDuringMonth) || null,
      OtherLeaveBalanceatendoftheMonth: parseNumeric(otherLeaveBalanceAtEndOfMonth) || null,
      Dateofgivingnoticeofpregnancy: dateOfGivingNoticeOfPregnancy || null,
      AmountofMatermity: parseNumeric(amountOfMaternity) || null,
      Subsequentpayment: parseNumeric(subsequentPayment) || null,
      Amountpaid: parseNumeric(amountPaid) || null,
      LeavewithWages: leaveWithWages || null,
      AmountpaidasGratuity: parseNumeric(amountPaidAsGratuity) || null,
      Remarks: remarks || null
    });
    const updated = await table.getRow(ROWID);
    res.status(200).send({ status: 'success', data: { form15I: updated } });
  } catch (err) {
    res.status(400).send({ status: 'failure', message: err.message || 'Failed to update form15I record.' });
  }
});

// Delete Form15I
app.delete('/form15I/:ROWID', async (req, res) => {
  try {
    const { ROWID } = req.params;
    const { catalyst } = res.locals;
    const table = catalyst.datastore().table('Form15I');
    await table.deleteRow(ROWID);
    res.status(200).send({ status: 'success', data: { id: ROWID } });
  } catch (err) {
    res.status(500).send({ status: 'failure', message: err.message || 'Failed to delete form15I record.' });
  }
});

// Upload file for new form15I (before creation)
app.post('/form15I/upload/:docType', async (req, res) => {
  try {
    const { docType } = req.params;
    console.log('New form15I file upload request for docType:', docType);
    console.log('Request headers:', req.headers);
    console.log('Request files:', req.files);
    
    const { catalyst } = res.locals;
    const folderId = DOC_TYPE_TO_FOLDER_ID[docType];
    console.log('Folder ID for docType:', docType, 'is:', folderId);
    
    if (!folderId) {
      console.error('Invalid document type:', docType);
      return res.status(400).json({ status: 'failure', message: 'Invalid document type.' });
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

      // Debug logging for Photo uploads
      if (docType === 'Photo') {
        console.log('Photo upload response (new form15I):', uploadResp);
      }

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
        console.error('Unexpected upload response (new form15I):', uploadResp);
        return res.status(500).json({ status: 'failure', message: 'File upload failed or invalid response from Catalyst.' });
      }

      // Debug logging for Photo fileId
      if (docType === 'Photo') {
        console.log('Photo fileId to be saved (new form15I):', fileId);
      }

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
    console.error('Upload error (new form15I):', err);
    res.status(500).json({ status: 'failure', message: err.message || 'File upload failed.' });
  }
});

// Upload file for existing form15I and docType
app.post('/form15I/:id/upload/:docType', async (req, res) => {
  try {
    const { id, docType } = req.params;
    console.log('Existing form15I file upload request for id:', id, 'docType:', docType);
    const { catalyst } = res.locals;
    const folderId = DOC_TYPE_TO_FOLDER_ID[docType];
    console.log('Folder ID for docType:', docType, 'is:', folderId);
    
    if (!folderId) {
      console.error('Invalid document type:', docType);
      return res.status(400).send({ status: 'failure', message: 'Invalid document type.' });
    }
    
    if (!req.files || !req.files.file) {
      console.error('No file uploaded for form15I id:', id, 'docType:', docType);
      return res.status(400).send({ status: 'failure', message: 'No file uploaded.' });
    }
    
    const file = req.files.file;
    const tempDir = os.tmpdir();
    const tempPath = path.join(tempDir, file.name);
    
    try {
      await file.mv(tempPath);
      const uploadResp = await catalyst.filestore().folder(folderId).uploadFile({
        code: fs.createReadStream(tempPath),
        name: file.name
      });

      // Clean up temporary file
      if (fs.existsSync(tempPath)) {
        fs.unlinkSync(tempPath);
      }

      // Debug logging for Photo uploads
      if (docType === 'Photo') {
        console.log('Photo upload response:', uploadResp);
      }

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
        return res.status(500).send({ status: 'failure', message: 'File upload failed or invalid response from Catalyst.' });
      }

      // Debug logging for Photo fileId
      if (docType === 'Photo') {
        console.log('Photo fileId to be saved:', fileId);
      }

      // Update Form15I row - store file ID in Photo column
      const table = catalyst.datastore().table('Form15I');
      await table.updateRow({
        ROWID: id,
        Photo: fileId // Store file ID in the existing Photo column
      });
      
      res.status(200).send({ status: 'success', fileId, fileName });
    } catch (uploadErr) {
      // Clean up temporary file in case of error
      if (fs.existsSync(tempPath)) {
        fs.unlinkSync(tempPath);
      }
      throw uploadErr;
    }
  } catch (err) {
    console.log(err);
    res.status(500).send({ status: 'failure', message: err.message || 'File upload failed.' });
  }
});

// Download file for form15I and docType
app.get('/form15I/:id/file/:docType', async (req, res) => {
  try {
    const { id, docType } = req.params;
    console.log('Download request:', { id, docType });
    const { catalyst } = res.locals;
    const folderId = DOC_TYPE_TO_FOLDER_ID[docType];
    console.log('Folder ID for docType:', docType, 'is:', folderId);
    
    if (!folderId) {
      console.error('Invalid document type:', docType);
      return res.status(400).send({ status: 'failure', message: 'Invalid document type.' });
    }
    
    const table = catalyst.datastore().table('Form15I');
    const form15I = await table.getRow(id);
    const fileId = form15I.Photo; // Get file ID from Photo column
    const fileName = 'photo_file'; // Default filename since we don't store it separately
    
    console.log('File info:', { fileId, fileName });
    
    if (!fileId) {
      console.error('No file ID found for form15I:', id, 'docType:', docType);
      return res.status(404).send({ status: 'failure', message: 'File not found for this form15I record.' });
    }
    
    const fileBuffer = await catalyst.filestore().folder(folderId).downloadFile(fileId);
    
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
        'txt': 'text/plain'
      };
      return contentTypes[ext] || 'application/octet-stream';
    };
    
    const contentType = getContentType(fileName);
    
    console.log('Sending file:', { fileName, contentType, fileSize: fileBuffer.length });
    res.writeHead(200, {
      'Content-Type': contentType,
      'Content-Length': fileBuffer.length,
      'Content-Disposition': `attachment; filename="${fileName}"`,
      'Cache-Control': 'no-cache',
      'Pragma': 'no-cache'
    });
    res.end(fileBuffer);
    console.log('File download completed for:', fileName);
  } catch (err) {
    console.log(err);
    res.status(500).send({ status: 'failure', message: err.message || 'File download failed.' });
  }
});

// Delete file for form15I and docType
app.delete('/form15I/:id/file/:docType', async (req, res) => {
  try {
    const { id, docType } = req.params;
    const { catalyst } = res.locals;
    const folderId = DOC_TYPE_TO_FOLDER_ID[docType];
    
    if (!folderId) {
      return res.status(400).send({ status: 'failure', message: 'Invalid document type.' });
    }
    
    const table = catalyst.datastore().table('Form15I');
    const form15I = await table.getRow(id);
    const fileId = form15I.Photo; // Get file ID from Photo column
    
    if (!fileId) {
      return res.status(404).send({ status: 'failure', message: 'File not found for this form15I record.' });
    }
    
    await catalyst.filestore().folder(folderId).deleteFile(fileId);
    
    // Clear Photo column
    await table.updateRow({
      ROWID: id,
      Photo: null
    });
    
    res.status(200).send({ status: 'success' });
  } catch (err) {
    console.log(err);
    res.status(500).send({ status: 'failure', message: err.message || 'File delete failed.' });
  }
});

module.exports = app;