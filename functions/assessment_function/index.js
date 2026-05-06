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
  res.status(200).json({ status: 'success', message: 'assessment_function ready' });
});

// Mapping of document types to Catalyst File Store folder IDs
const DOC_TYPE_TO_FOLDER_ID = {
  Attachment1: '28341000000040001', // Attachment1 folder ID
  Attachment2: '28341000000040002', // Attachment2 folder ID
  Attachment3: '28341000000040003'  // Attachment3 folder ID
};

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
    const table = catalyst.datastore().table('AssessmentReport');
    
    // Try to get table info
    console.log('Table object:', table);
    
    // Try a simple query
    const zcql = catalyst.zcql();
    const testQuery = await zcql.executeZCQLQuery('SELECT COUNT(ROWID) as count FROM AssessmentReport');
    console.log('Test query result:', testQuery);
    
    res.status(200).send({ 
      status: 'success', 
      message: 'Table access successful',
      tableInfo: {
        tableName: 'AssessmentReport',
        recordCount: testQuery[0]?.AssessmentReport?.count || 0
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
    
    // Try to access the Attachment1 folder
    const folder = filestore.folder('Attachment1');
    console.log('Folder object:', folder);
    
    // Try to list files in the folder
    const files = await folder.getAllFiles();
    console.log('Files in Attachment1 folder:', files);
    
    res.status(200).send({ 
      status: 'success', 
      message: 'File Store access successful',
      fileStoreInfo: {
        folderName: 'Attachment1',
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

// ===== AssessmentReport CRUD =====
// Table: AssessmentReport (columns: Act, Description, MaximumMarks, Applicability, CheckforRecords, RecordCategory, ScoringCriteria, Level, MarksObtained, Remarks, Attachment1, Attachment2, Attachment3, RejectedRemarks, ListOfComplianceV2ID)

// Create AssessmentReport
app.post('/assessment-report', async (req, res) => {
  try {
    console.log('Received request body:', JSON.stringify(req.body, null, 2));
    
    const {
      act,
      description,
      maximumMarks,
      applicability,
      checkforRecords,
      recordCategory,
      scoringCriteria,
      level,
      marksObtained,
      remarks,
      attachment1,
      attachment2,
      attachment3,
      rejectedRemarks,
      listOfComplianceV2ID
    } = req.body;

    // Validate required fields
    if (!act || !String(act).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Act is required.' });
    }
    if (!description || !String(description).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Description is required.' });
    }
    if (!maximumMarks || !String(maximumMarks).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Maximum Marks is required.' });
    }
    if (!applicability || !String(applicability).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Applicability is required.' });
    }
    if (!checkforRecords || !String(checkforRecords).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Check for Records is required.' });
    }
    if (!recordCategory || !String(recordCategory).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Record Category is required.' });
    }
    if (!scoringCriteria || !String(scoringCriteria).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Scoring Criteria is required.' });
    }
    if (!level || !String(level).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Level is required.' });
    }
    if (!marksObtained || !String(marksObtained).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Marks Obtained is required.' });
    }

    const { catalyst } = res.locals;
    
    const table = catalyst.datastore().table('AssessmentReport');
    
    const insertData = {
      Act: act,
      Description: description,
      MaximumMarks: maximumMarks,
      Applicability: applicability,
      CheckforRecords: checkforRecords,
      RecordCategory: recordCategory,
      ScoringCriteria: scoringCriteria,
      Level: level,
      MarksObtained: marksObtained,
      Remarks: remarks || null,
      Attachment1: attachment1 || null,
      Attachment2: attachment2 || null,
      Attachment3: attachment3 || null,
      RejectedRemarks: rejectedRemarks || null,
      ListOfComplianceV2ID: listOfComplianceV2ID || null
    };
    
    console.log('Inserting data to AssessmentReport table:', JSON.stringify(insertData, null, 2));
    
    const insertResp = await table.insertRow(insertData);
    console.log('Insert response:', JSON.stringify(insertResp, null, 2));
    
    const created = await table.getRow(insertResp.ROWID);
    console.log('Created record:', JSON.stringify(created, null, 2));
    
    res.status(200).send({ status: 'success', data: { assessmentReport: created } });
  } catch (err) {
    console.error('Error creating assessment report:', err);
    res.status(400).send({ status: 'failure', message: err.message || 'Failed to create assessment report.' });
  }
});

// List AssessmentReport with optional pagination
app.get('/assessment-report', async (req, res) => {
  try {
    console.log('Fetching assessment reports...');
    const { catalyst } = res.locals;
    const zcql = catalyst.zcql();
    const page = parseInt(req.query.page, 10) || 1;
    const perPage = parseInt(req.query.perPage, 10) || 50;
    const returnAll = !req.query.page && !req.query.perPage;
    const limitClause = returnAll ? '' : `LIMIT ${(page - 1) * perPage + 1},${perPage}`;
    
    console.log('Executing count query...');
    const countRows = await zcql.executeZCQLQuery('SELECT COUNT(ROWID) as count FROM AssessmentReport');
    const total = parseInt(countRows[0].AssessmentReport.count, 10) || 0;
    console.log('Total records:', total);
    
    console.log('Executing data query...');
    const rows = await zcql.executeZCQLQuery(`SELECT ROWID, Act, Description, MaximumMarks, Applicability, CheckforRecords, RecordCategory, ScoringCriteria, Level, MarksObtained, Remarks, Attachment1, Attachment2, Attachment3, RejectedRemarks, ListOfComplianceV2ID, CREATEDTIME, MODIFIEDTIME FROM AssessmentReport ORDER BY ROWID DESC ${limitClause}`);
    console.log('Raw rows from database:', JSON.stringify(rows, null, 2));
    
    const assessmentReports = rows.map(r => ({
      id: r.AssessmentReport.ROWID,
      act: r.AssessmentReport.Act,
      description: r.AssessmentReport.Description,
      maximumMarks: r.AssessmentReport.MaximumMarks,
      applicability: r.AssessmentReport.Applicability,
      checkforRecords: r.AssessmentReport.CheckforRecords,
      recordCategory: r.AssessmentReport.RecordCategory,
      scoringCriteria: r.AssessmentReport.ScoringCriteria,
      level: r.AssessmentReport.Level,
      marksObtained: r.AssessmentReport.MarksObtained,
      remarks: r.AssessmentReport.Remarks,
      attachment1: r.AssessmentReport.Attachment1,
      attachment2: r.AssessmentReport.Attachment2,
      attachment3: r.AssessmentReport.Attachment3,
      rejectedRemarks: r.AssessmentReport.RejectedRemarks,
      listOfComplianceV2ID: r.AssessmentReport.ListOfComplianceV2ID,
      createdTime: r.AssessmentReport.CREATEDTIME,
      modifiedTime: r.AssessmentReport.MODIFIEDTIME
    }));
    
    console.log('Processed assessment reports:', JSON.stringify(assessmentReports, null, 2));
    res.status(200).send({ status: 'success', data: { assessmentReports, total, hasMore: returnAll ? false : page * perPage < total } });
  } catch (err) {
    console.error('Error fetching assessment reports:', err);
    res.status(500).send({ status: 'failure', message: err.message || 'Failed to fetch assessment reports.' });
  }
});

// Update AssessmentReport
app.put('/assessment-report/:ROWID', async (req, res) => {
  try {
    const { ROWID } = req.params;
    const {
      act,
      description,
      maximumMarks,
      applicability,
      checkforRecords,
      recordCategory,
      scoringCriteria,
      level,
      marksObtained,
      remarks,
      attachment1,
      attachment2,
      attachment3,
      rejectedRemarks,
      listOfComplianceV2ID
    } = req.body;

    // Validate required fields
    if (!act || !String(act).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Act is required.' });
    }
    if (!description || !String(description).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Description is required.' });
    }
    if (!maximumMarks || !String(maximumMarks).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Maximum Marks is required.' });
    }
    if (!applicability || !String(applicability).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Applicability is required.' });
    }
    if (!checkforRecords || !String(checkforRecords).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Check for Records is required.' });
    }
    if (!recordCategory || !String(recordCategory).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Record Category is required.' });
    }
    if (!scoringCriteria || !String(scoringCriteria).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Scoring Criteria is required.' });
    }
    if (!level || !String(level).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Level is required.' });
    }
    if (!marksObtained || !String(marksObtained).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Marks Obtained is required.' });
    }

    const { catalyst } = res.locals;
    const table = catalyst.datastore().table('AssessmentReport');
    await table.updateRow({
      ROWID,
      Act: act,
      Description: description,
      MaximumMarks: maximumMarks,
      Applicability: applicability,
      CheckforRecords: checkforRecords,
      RecordCategory: recordCategory,
      ScoringCriteria: scoringCriteria,
      Level: level,
      MarksObtained: marksObtained,
      Remarks: remarks || null,
      Attachment1: attachment1 || null,
      Attachment2: attachment2 || null,
      Attachment3: attachment3 || null,
      RejectedRemarks: rejectedRemarks || null,
      ListOfComplianceV2ID: listOfComplianceV2ID || null
    });
    const updated = await table.getRow(ROWID);
    res.status(200).send({ status: 'success', data: { assessmentReport: updated } });
  } catch (err) {
    res.status(400).send({ status: 'failure', message: err.message || 'Failed to update assessment report.' });
  }
});

// Delete AssessmentReport
app.delete('/assessment-report/:ROWID', async (req, res) => {
  try {
    const { ROWID } = req.params;
    const { catalyst } = res.locals;
    const table = catalyst.datastore().table('AssessmentReport');
    await table.deleteRow(ROWID);
    res.status(200).send({ status: 'success', data: { id: ROWID } });
  } catch (err) {
    res.status(500).send({ status: 'failure', message: err.message || 'Failed to delete assessment report.' });
  }
});

// Upload file for new assessment report (before creation)
app.post('/assessment-report/upload/:docType', async (req, res) => {
  try {
    const { docType } = req.params;
    console.log('New assessment report file upload request for docType:', docType);
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

      // Debug logging for file uploads
      console.log('File upload response (new assessment report):', uploadResp);

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
        console.error('Unexpected upload response (new assessment report):', uploadResp);
        return res.status(500).json({ status: 'failure', message: 'File upload failed or invalid response from Catalyst.' });
      }

      // Debug logging for fileId
      console.log('FileId to be saved (new assessment report):', fileId);

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
    console.error('Upload error (new assessment report):', err);
    res.status(500).json({ status: 'failure', message: err.message || 'File upload failed.' });
  }
});

// Upload file for existing assessment report and docType
app.post('/assessment-report/:id/upload/:docType', async (req, res) => {
  try {
    const { id, docType } = req.params;
    console.log('Existing assessment report file upload request for id:', id, 'docType:', docType);
    const { catalyst } = res.locals;
    const folderId = DOC_TYPE_TO_FOLDER_ID[docType];
    console.log('Folder ID for docType:', docType, 'is:', folderId);
    
    if (!folderId) {
      console.error('Invalid document type:', docType);
      return res.status(400).send({ status: 'failure', message: 'Invalid document type.' });
    }
    
    if (!req.files || !req.files.file) {
      console.error('No file uploaded for assessment report id:', id, 'docType:', docType);
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
        return res.status(500).send({ status: 'failure', message: 'File upload failed or invalid response from Catalyst.' });
      }

      // Debug logging for fileId
      console.log('FileId to be saved:', fileId);

      // Update AssessmentReport row - store file ID in appropriate column
      const table = catalyst.datastore().table('AssessmentReport');
      const updateData = { ROWID: id };
      updateData[docType] = fileId; // Store file ID in the appropriate column
      await table.updateRow(updateData);
      
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

// Download file for assessment report and docType
app.get('/assessment-report/:id/file/:docType', async (req, res) => {
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
    
    const table = catalyst.datastore().table('AssessmentReport');
    const assessmentReport = await table.getRow(id);
    const fileId = assessmentReport[docType]; // Get file ID from the appropriate column
    const fileName = `${docType}_file`; // Default filename since we don't store it separately
    
    console.log('File info:', { fileId, fileName });
    
    if (!fileId) {
      console.error('No file ID found for assessment report:', id, 'docType:', docType);
      return res.status(404).send({ status: 'failure', message: 'File not found for this assessment report.' });
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

// Delete file for assessment report and docType
app.delete('/assessment-report/:id/file/:docType', async (req, res) => {
  try {
    const { id, docType } = req.params;
    const { catalyst } = res.locals;
    const folderId = DOC_TYPE_TO_FOLDER_ID[docType];
    
    if (!folderId) {
      return res.status(400).send({ status: 'failure', message: 'Invalid document type.' });
    }
    
    const table = catalyst.datastore().table('AssessmentReport');
    const assessmentReport = await table.getRow(id);
    const fileId = assessmentReport[docType]; // Get file ID from the appropriate column
    
    if (!fileId) {
      return res.status(404).send({ status: 'failure', message: 'File not found for this assessment report.' });
    }
    
    await catalyst.filestore().folder(folderId).deleteFile(fileId);
    
    // Clear the appropriate column
    const updateData = { ROWID: id };
    updateData[docType] = null;
    await table.updateRow(updateData);
    
    res.status(200).send({ status: 'success' });
  } catch (err) {
    console.log(err);
    res.status(500).send({ status: 'failure', message: err.message || 'File delete failed.' });
  }
});

module.exports = app;