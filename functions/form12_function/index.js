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
  res.status(200).json({ status: 'success', message: 'form12_function ready' });
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
    const table = catalyst.datastore().table('Form12');
    
    // Try to get table info
    console.log('Table object:', table);
    
    // Try a simple query
    const zcql = catalyst.zcql();
    const testQuery = await zcql.executeZCQLQuery('SELECT COUNT(ROWID) as count FROM Form12');
    console.log('Test query result:', testQuery);
    
    res.status(200).send({ 
      status: 'success', 
      message: 'Table access successful',
      tableInfo: {
        tableName: 'Form12',
        recordCount: testQuery[0]?.Form12?.count || 0
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

// ===== Form12 CRUD =====
// Table: Form12 (columns: NameoftheWorker, WorkerIdentityNo, Gender, FatherName, DateofBirth, PresentAddress, PresentCity, PresentState, PresentPostalCode, PresentCountry, PermanentAddress, PermanentCity, PermanentState, PermanentPostalCode, PermanentCountry, AadhaarNo, Dateofentryintoservice, Designation, EPFNo, UANNo, ESINo, Dateonwhichcompleted, Dateonwhichmadepermanent, PeriodofSuspension, BankAccountNumber, NameoftheBank, Branch, Photo, MobileNumber, Gmail, SpecimenSignature, DateofExit, ReasonforExit, Remarks)

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

// Create Form12
app.post('/form12', async (req, res) => {
  try {
    console.log('Received request body:', JSON.stringify(req.body, null, 2));
    
    const {
      nameOfWorker,
      workerIdentityNo,
      gender,
      fatherName,
      dateOfBirth,
      presentAddress,
      presentCity,
      presentState,
      presentPostalCode,
      presentCountry,
      permanentAddress,
      permanentCity,
      permanentState,
      permanentPostalCode,
      permanentCountry,
      aadhaarNo,
      dateOfEntryIntoService,
      designation,
      epfNo,
      uanNo,
      esiNo,
      dateOnWhichCompleted,
      dateOnWhichMadePermanent,
      periodOfSuspension,
      bankAccountNumber,
      nameOfBank,
      branch,
      photo,
      mobileNumber,
      gmail,
      specimenSignature,
      dateOfExit,
      reasonForExit,
      remarks
    } = req.body;

    // Validate required fields
    if (!nameOfWorker || !String(nameOfWorker).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Name of the Worker is required.' });
    }
    if (!workerIdentityNo || !String(workerIdentityNo).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Worker Identity No is required.' });
    }
    if (!gender || !String(gender).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Gender is required.' });
    }
    if (!fatherName || !String(fatherName).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Father Name is required.' });
    }
    if (!dateOfBirth || !String(dateOfBirth).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Date of Birth is required.' });
    }

    const { catalyst } = res.locals;
    
    // Handle photo upload to File Store
    let photoUrl = null;
    if (photo && photo.trim()) {
      console.log('Uploading photo to File Store...');
      photoUrl = await uploadFileToFileStore(catalyst, photo, `${nameOfWorker}_photo`);
      console.log('Photo upload result:', photoUrl);
    }
    
    const table = catalyst.datastore().table('Form12');
    
    const insertData = {
      NameoftheWorker: nameOfWorker,
      WorkerIdentityNo: workerIdentityNo,
      Gender: gender,
      FatherName: fatherName,
      DateofBirth: dateOfBirth,
      PresentAddress: presentAddress || null,
      PresentCity: presentCity || null,
      PresentState: presentState || null,
      PresentPostalCode: presentPostalCode || null,
      PresentCountry: presentCountry || null,
      PermanentAddress: permanentAddress || null,
      PermanentCity: permanentCity || null,
      PermanentState: permanentState || null,
      PermanentPostalCode: permanentPostalCode || null,
      PermanentCountry: permanentCountry || null,
      AadhaarNo: aadhaarNo || null,
      Dateofentryintoservice: dateOfEntryIntoService || null,
      Designation: designation || null,
      EPFNo: epfNo || null,
      UANNo: uanNo || null,
      ESINo: esiNo || null,
      Dateonwhichcompletionof480daysofservice: dateOnWhichCompleted || null,
      Dateonwhichmadepermanent: dateOnWhichMadePermanent || null,
      PeriodofSuspensionifany: periodOfSuspension || null,
      BankAccountNumber: bankAccountNumber || null,
      NameoftheBank: nameOfBank || null,
      Branch: branch || null,
      Photo: photoUrl, // Store the file ID or URL in the existing Photo column
      MobileNumber: mobileNumber || null,
      Gmail: gmail || null,
      SpecimenSignature: specimenSignature || null,
      DateofExit: dateOfExit || null,
      ReasonforExit: reasonForExit || null,
      Remarks: remarks || null
    };
    
    console.log('Inserting data to Form12 table:', JSON.stringify(insertData, null, 2));
    
    const insertResp = await table.insertRow(insertData);
    console.log('Insert response:', JSON.stringify(insertResp, null, 2));
    
    const created = await table.getRow(insertResp.ROWID);
    console.log('Created record:', JSON.stringify(created, null, 2));
    
    res.status(200).send({ status: 'success', data: { form12: created } });
  } catch (err) {
    console.error('Error creating form12 record:', err);
    res.status(400).send({ status: 'failure', message: err.message || 'Failed to create form12 record.' });
  }
});

// Get single Form12 record with all details
app.get('/form12/:ROWID', async (req, res) => {
  try {
    const { ROWID } = req.params;
    console.log('Fetching form12 record:', ROWID);
    const { catalyst } = res.locals;
    const table = catalyst.datastore().table('Form12');
    
    const form12 = await table.getRow(ROWID);
    console.log('Retrieved form12 record:', JSON.stringify(form12, null, 2));
    
    res.status(200).send({ status: 'success', data: { form12 } });
  } catch (err) {
    console.error('Error fetching form12 record:', err);
    res.status(500).send({ status: 'failure', message: err.message || 'Failed to fetch form12 record.' });
  }
});

// List Form12 with optional pagination
app.get('/form12', async (req, res) => {
  try {
    console.log('Fetching form12 records...');
    const { catalyst } = res.locals;
    const zcql = catalyst.zcql();
    const page = parseInt(req.query.page, 10) || 1;
    const perPage = parseInt(req.query.perPage, 10) || 50;
    const returnAll = !req.query.page && !req.query.perPage;
    const limitClause = returnAll ? '' : `LIMIT ${(page - 1) * perPage + 1},${perPage}`;
    
    console.log('Executing count query...');
    const countRows = await zcql.executeZCQLQuery('SELECT COUNT(ROWID) as count FROM Form12');
    const total = parseInt(countRows[0].Form12.count, 10) || 0;
    console.log('Total records:', total);
    
    console.log('Executing data query...');
    // Select only essential columns to avoid the 30-column limit
    const rows = await zcql.executeZCQLQuery(`SELECT ROWID, NameoftheWorker, WorkerIdentityNo, Gender, FatherName, DateofBirth, PresentAddress, PresentCity, PresentState, PresentPostalCode, PresentCountry, PermanentAddress, PermanentCity, PermanentState, PermanentPostalCode, PermanentCountry, AadhaarNo, Dateofentryintoservice, Designation, EPFNo, UANNo, ESINo, Dateonwhichcompletionof480daysofservice, Dateonwhichmadepermanent, PeriodofSuspensionifany, BankAccountNumber, NameoftheBank, Branch, Photo, MobileNumber, Gmail, SpecimenSignature, DateofExit, ReasonforExit, Remarks, CREATEDTIME FROM Form12 ORDER BY ROWID DESC ${limitClause}`);
    console.log('Raw rows from database:', JSON.stringify(rows, null, 2));
    
    const form12Data = rows.map(r => ({
      id: r.Form12.ROWID,
      nameOfWorker: r.Form12.NameoftheWorker,
      workerIdentityNo: r.Form12.WorkerIdentityNo,
      gender: r.Form12.Gender,
      fatherName: r.Form12.FatherName,
      dateOfBirth: r.Form12.DateofBirth,
      presentAddress: r.Form12.PresentAddress,
      presentCity: r.Form12.PresentCity,
      presentState: r.Form12.PresentState,
      presentPostalCode: r.Form12.PresentPostalCode,
      presentCountry: r.Form12.PresentCountry,
      permanentAddress: r.Form12.PermanentAddress,
      permanentCity: r.Form12.PermanentCity,
      permanentState: r.Form12.PermanentState,
      permanentPostalCode: r.Form12.PermanentPostalCode,
      permanentCountry: r.Form12.PermanentCountry,
      aadhaarNo: r.Form12.AadhaarNo,
      dateOfEntryIntoService: r.Form12.Dateofentryintoservice,
      designation: r.Form12.Designation,
      epfNo: r.Form12.EPFNo,
      uanNo: r.Form12.UANNo,
      esiNo: r.Form12.ESINo,
      dateOnWhichCompleted: r.Form12.Dateonwhichcompletionof480daysofservice,
      dateOnWhichMadePermanent: r.Form12.Dateonwhichmadepermanent,
      periodOfSuspension: r.Form12.PeriodofSuspensionifany,
      bankAccountNumber: r.Form12.BankAccountNumber,
      nameOfBank: r.Form12.NameoftheBank,
      branch: r.Form12.Branch,
      photo: r.Form12.Photo,
      mobileNumber: r.Form12.MobileNumber,
      gmail: r.Form12.Gmail,
      specimenSignature: r.Form12.SpecimenSignature,
      dateOfExit: r.Form12.DateofExit,
      reasonForExit: r.Form12.ReasonforExit,
      remarks: r.Form12.Remarks,
      createdTime: r.Form12.CREATEDTIME,
      modifiedTime: null // Not selected in query
    }));
    
    console.log('Processed form12 data:', JSON.stringify(form12Data, null, 2));
    res.status(200).send({ status: 'success', data: { form12Data, total, hasMore: returnAll ? false : page * perPage < total } });
  } catch (err) {
    console.error('Error fetching form12 records:', err);
    res.status(500).send({ status: 'failure', message: err.message || 'Failed to fetch form12 records.' });
  }
});

// Update Form12
app.put('/form12/:ROWID', async (req, res) => {
  try {
    const { ROWID } = req.params;
    const {
      nameOfWorker,
      workerIdentityNo,
      gender,
      fatherName,
      dateOfBirth,
      presentAddress,
      presentCity,
      presentState,
      presentPostalCode,
      presentCountry,
      permanentAddress,
      permanentCity,
      permanentState,
      permanentPostalCode,
      permanentCountry,
      aadhaarNo,
      dateOfEntryIntoService,
      designation,
      epfNo,
      uanNo,
      esiNo,
      dateOnWhichCompleted,
      dateOnWhichMadePermanent,
      periodOfSuspension,
      bankAccountNumber,
      nameOfBank,
      branch,
      photo,
      mobileNumber,
      gmail,
      specimenSignature,
      dateOfExit,
      reasonForExit,
      remarks
    } = req.body;

    // Validate required fields
    if (!nameOfWorker || !String(nameOfWorker).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Name of the Worker is required.' });
    }
    if (!workerIdentityNo || !String(workerIdentityNo).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Worker Identity No is required.' });
    }
    if (!gender || !String(gender).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Gender is required.' });
    }
    if (!fatherName || !String(fatherName).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Father Name is required.' });
    }
    if (!dateOfBirth || !String(dateOfBirth).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Date of Birth is required.' });
    }

    const { catalyst } = res.locals;
    const table = catalyst.datastore().table('Form12');
    await table.updateRow({
      ROWID,
      NameoftheWorker: nameOfWorker,
      WorkerIdentityNo: workerIdentityNo,
      Gender: gender,
      FatherName: fatherName,
      DateofBirth: dateOfBirth,
      PresentAddress: presentAddress || null,
      PresentCity: presentCity || null,
      PresentState: presentState || null,
      PresentPostalCode: presentPostalCode || null,
      PresentCountry: presentCountry || null,
      PermanentAddress: permanentAddress || null,
      PermanentCity: permanentCity || null,
      PermanentState: permanentState || null,
      PermanentPostalCode: permanentPostalCode || null,
      PermanentCountry: permanentCountry || null,
      AadhaarNo: aadhaarNo || null,
      Dateofentryintoservice: dateOfEntryIntoService || null,
      Designation: designation || null,
      EPFNo: epfNo || null,
      UANNo: uanNo || null,
      ESINo: esiNo || null,
      Dateonwhichcompletionof480daysofservice: dateOnWhichCompleted || null,
      Dateonwhichmadepermanent: dateOnWhichMadePermanent || null,
      PeriodofSuspensionifany: periodOfSuspension || null,
      BankAccountNumber: bankAccountNumber || null,
      NameoftheBank: nameOfBank || null,
      Branch: branch || null,
      Photo: photo || null,
      MobileNumber: mobileNumber || null,
      Gmail: gmail || null,
      SpecimenSignature: specimenSignature || null,
      DateofExit: dateOfExit || null,
      ReasonforExit: reasonForExit || null,
      Remarks: remarks || null
    });
    const updated = await table.getRow(ROWID);
    res.status(200).send({ status: 'success', data: { form12: updated } });
  } catch (err) {
    res.status(400).send({ status: 'failure', message: err.message || 'Failed to update form12 record.' });
  }
});

// Delete Form12
app.delete('/form12/:ROWID', async (req, res) => {
  try {
    const { ROWID } = req.params;
    const { catalyst } = res.locals;
    const table = catalyst.datastore().table('Form12');
    await table.deleteRow(ROWID);
    res.status(200).send({ status: 'success', data: { id: ROWID } });
  } catch (err) {
    res.status(500).send({ status: 'failure', message: err.message || 'Failed to delete form12 record.' });
  }
});

// Upload file for new form12 (before creation)
app.post('/form12/upload/:docType', async (req, res) => {
  try {
    const { docType } = req.params;
    console.log('New form12 file upload request for docType:', docType);
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
        console.log('Photo upload response (new form12):', uploadResp);
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
        console.error('Unexpected upload response (new form12):', uploadResp);
        return res.status(500).json({ status: 'failure', message: 'File upload failed or invalid response from Catalyst.' });
      }

      // Debug logging for Photo fileId
      if (docType === 'Photo') {
        console.log('Photo fileId to be saved (new form12):', fileId);
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
    console.error('Upload error (new form12):', err);
    res.status(500).json({ status: 'failure', message: err.message || 'File upload failed.' });
  }
});

// Upload file for existing form12 and docType
app.post('/form12/:id/upload/:docType', async (req, res) => {
  try {
    const { id, docType } = req.params;
    console.log('Existing form12 file upload request for id:', id, 'docType:', docType);
    const { catalyst } = res.locals;
    const folderId = DOC_TYPE_TO_FOLDER_ID[docType];
    console.log('Folder ID for docType:', docType, 'is:', folderId);
    
    if (!folderId) {
      console.error('Invalid document type:', docType);
      return res.status(400).send({ status: 'failure', message: 'Invalid document type.' });
    }
    
    if (!req.files || !req.files.file) {
      console.error('No file uploaded for form12 id:', id, 'docType:', docType);
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

      // Update Form12 row - store file ID in Photo column
      const table = catalyst.datastore().table('Form12');
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

// Download file for form12 and docType
app.get('/form12/:id/file/:docType', async (req, res) => {
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
    
    const table = catalyst.datastore().table('Form12');
    const form12 = await table.getRow(id);
    const fileId = form12.Photo; // Get file ID from Photo column
    const fileName = 'photo_file'; // Default filename since we don't store it separately
    
    console.log('File info:', { fileId, fileName });
    
    if (!fileId) {
      console.error('No file ID found for form12:', id, 'docType:', docType);
      return res.status(404).send({ status: 'failure', message: 'File not found for this form12 record.' });
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

// Delete file for form12 and docType
app.delete('/form12/:id/file/:docType', async (req, res) => {
  try {
    const { id, docType } = req.params;
    const { catalyst } = res.locals;
    const folderId = DOC_TYPE_TO_FOLDER_ID[docType];
    
    if (!folderId) {
      return res.status(400).send({ status: 'failure', message: 'Invalid document type.' });
    }
    
    const table = catalyst.datastore().table('Form12');
    const form12 = await table.getRow(id);
    const fileId = form12.Photo; // Get file ID from Photo column
    
    if (!fileId) {
      return res.status(404).send({ status: 'failure', message: 'File not found for this form12 record.' });
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