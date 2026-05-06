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
  res.status(200).json({ status: 'success', message: 'clientdetails_function ready' });
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
    const table = catalyst.datastore().table('ClientDetails');
    
    // Try to get table info
    console.log('Table object:', table);
    
    // Try a simple query
    const zcql = catalyst.zcql();
    const testQuery = await zcql.executeZCQLQuery('SELECT COUNT(ROWID) as count FROM ClientDetails');
    console.log('Test query result:', testQuery);
    
    res.status(200).send({ 
      status: 'success', 
      message: 'Table access successful',
      tableInfo: {
        tableName: 'ClientDetails',
        recordCount: testQuery[0]?.ClientDetails?.count || 0
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
    
    // Try to access the Logo folder
    const folder = filestore.folder('Logo');
    console.log('Folder object:', folder);
    
    // Try to list files in the folder
    const files = await folder.getAllFiles();
    console.log('Files in Logo folder:', files);
    
    res.status(200).send({ 
      status: 'success', 
      message: 'File Store access successful',
      fileStoreInfo: {
        folderName: 'Logo',
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

// ===== ClientDetails CRUD =====
// Table: ClientDetails (columns: CompanyName, PortalEmailID, AddressLine1, AddressLine2, City, State, Pincode, Logo, ClientCode)

// Mapping of document types to Catalyst File Store folder IDs
const DOC_TYPE_TO_FOLDER_ID = {
  Logo: '28341000000023268' // Your Logo folder ID from the File Store
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
      const folder = filestore.folder('Logo'); // Use the Logo folder you created
      
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

// Create ClientDetail
app.post('/clientdetails', async (req, res) => {
  try {
    console.log('Received request body:', JSON.stringify(req.body, null, 2));
    
    const {
      companyName,
      portalEmailID,
      addressLine1,
      addressLine2,
      city,
      state,
      pincode,
      logo,
      clientCode
    } = req.body;

    // Validate required fields
    if (!companyName || !String(companyName).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Company Name is required.' });
    }
    if (!portalEmailID || !String(portalEmailID).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Portal Email ID is required.' });
    }
    if (!clientCode || !String(clientCode).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Client Code is required.' });
    }
    if (!addressLine1 || !String(addressLine1).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Address Line 1 is required.' });
    }
    if (!city || !String(city).trim()) {
      return res.status(400).send({ status: 'failure', message: 'City is required.' });
    }
    if (!state || !String(state).trim()) {
      return res.status(400).send({ status: 'failure', message: 'State is required.' });
    }
    if (!pincode || !String(pincode).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Pincode is required.' });
    }

    const { catalyst } = res.locals;
    
    // Handle logo upload to File Store
    let logoUrl = null;
    if (logo && logo.trim()) {
      console.log('Uploading logo to File Store...');
      logoUrl = await uploadFileToFileStore(catalyst, logo, `${companyName}_logo`);
      console.log('Logo upload result:', logoUrl);
    }
    
    const table = catalyst.datastore().table('ClientDetails');
    
    const insertData = {
      CompanyName: companyName,
      PortalEmailID: portalEmailID,
      AddressLine1: addressLine1,
      AddressLine2: addressLine2 || null,
      City: city,
      State: state,
      Pincode: pincode,
      Logo: logoUrl, // Store the file ID or URL in the existing Logo column
      ClientCode: clientCode
    };
    
    console.log('Inserting data to ClientDetails table:', JSON.stringify(insertData, null, 2));
    
    const insertResp = await table.insertRow(insertData);
    console.log('Insert response:', JSON.stringify(insertResp, null, 2));
    
    const created = await table.getRow(insertResp.ROWID);
    console.log('Created record:', JSON.stringify(created, null, 2));
    
    res.status(200).send({ status: 'success', data: { clientDetail: created } });
  } catch (err) {
    console.error('Error creating client detail:', err);
    res.status(400).send({ status: 'failure', message: err.message || 'Failed to create client detail.' });
  }
});

// List ClientDetails with optional pagination
app.get('/clientdetails', async (req, res) => {
  try {
    console.log('Fetching client details...');
    const { catalyst } = res.locals;
    const zcql = catalyst.zcql();
    const page = parseInt(req.query.page, 10) || 1;
    const perPage = parseInt(req.query.perPage, 10) || 50;
    const returnAll = !req.query.page && !req.query.perPage;
    const limitClause = returnAll ? '' : `LIMIT ${(page - 1) * perPage + 1},${perPage}`;
    
    console.log('Executing count query...');
    const countRows = await zcql.executeZCQLQuery('SELECT COUNT(ROWID) as count FROM ClientDetails');
    const total = parseInt(countRows[0].ClientDetails.count, 10) || 0;
    console.log('Total records:', total);
    
    console.log('Executing data query...');
    const rows = await zcql.executeZCQLQuery(`SELECT ROWID, CompanyName, PortalEmailID, AddressLine1, AddressLine2, City, State, Pincode, Logo, ClientCode, CREATEDTIME, MODIFIEDTIME FROM ClientDetails ORDER BY ROWID DESC ${limitClause}`);
    console.log('Raw rows from database:', JSON.stringify(rows, null, 2));
    
    const clientDetails = rows.map(r => ({
      id: r.ClientDetails.ROWID,
      companyName: r.ClientDetails.CompanyName,
      portalEmailID: r.ClientDetails.PortalEmailID,
      addressLine1: r.ClientDetails.AddressLine1,
      addressLine2: r.ClientDetails.AddressLine2,
      city: r.ClientDetails.City,
      state: r.ClientDetails.State,
      pincode: r.ClientDetails.Pincode,
      logo: r.ClientDetails.Logo,
      clientCode: r.ClientDetails.ClientCode,
      createdTime: r.ClientDetails.CREATEDTIME,
      modifiedTime: r.ClientDetails.MODIFIEDTIME
    }));
    
    console.log('Processed client details:', JSON.stringify(clientDetails, null, 2));
    res.status(200).send({ status: 'success', data: { clientDetails, total, hasMore: returnAll ? false : page * perPage < total } });
  } catch (err) {
    console.error('Error fetching client details:', err);
    res.status(500).send({ status: 'failure', message: err.message || 'Failed to fetch client details.' });
  }
});

// Update ClientDetail
app.put('/clientdetails/:ROWID', async (req, res) => {
  try {
    const { ROWID } = req.params;
    const {
      companyName,
      portalEmailID,
      addressLine1,
      addressLine2,
      city,
      state,
      pincode,
      logo,
      clientCode
    } = req.body;

    // Validate required fields
    if (!companyName || !String(companyName).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Company Name is required.' });
    }
    if (!portalEmailID || !String(portalEmailID).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Portal Email ID is required.' });
    }
    if (!clientCode || !String(clientCode).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Client Code is required.' });
    }
    if (!addressLine1 || !String(addressLine1).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Address Line 1 is required.' });
    }
    if (!city || !String(city).trim()) {
      return res.status(400).send({ status: 'failure', message: 'City is required.' });
    }
    if (!state || !String(state).trim()) {
      return res.status(400).send({ status: 'failure', message: 'State is required.' });
    }
    if (!pincode || !String(pincode).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Pincode is required.' });
    }

    const { catalyst } = res.locals;
    const table = catalyst.datastore().table('ClientDetails');
    await table.updateRow({
      ROWID,
      CompanyName: companyName,
      PortalEmailID: portalEmailID,
      AddressLine1: addressLine1,
      AddressLine2: addressLine2 || null,
      City: city,
      State: state,
      Pincode: pincode,
      Logo: logo || null,
      ClientCode: clientCode
    });
    const updated = await table.getRow(ROWID);
    res.status(200).send({ status: 'success', data: { clientDetail: updated } });
  } catch (err) {
    res.status(400).send({ status: 'failure', message: err.message || 'Failed to update client detail.' });
  }
});

// Delete ClientDetail
app.delete('/clientdetails/:ROWID', async (req, res) => {
  try {
    const { ROWID } = req.params;
    const { catalyst } = res.locals;
    const table = catalyst.datastore().table('ClientDetails');
    await table.deleteRow(ROWID);
    res.status(200).send({ status: 'success', data: { id: ROWID } });
  } catch (err) {
    res.status(500).send({ status: 'failure', message: err.message || 'Failed to delete client detail.' });
  }
});

// Upload file for new client (before creation)
app.post('/clientdetails/upload/:docType', async (req, res) => {
  try {
    const { docType } = req.params;
    console.log('New client file upload request for docType:', docType);
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

      // Debug logging for Logo uploads
      if (docType === 'Logo') {
        console.log('Logo upload response (new client):', uploadResp);
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
        console.error('Unexpected upload response (new client):', uploadResp);
        return res.status(500).json({ status: 'failure', message: 'File upload failed or invalid response from Catalyst.' });
      }

      // Debug logging for Logo fileId
      if (docType === 'Logo') {
        console.log('Logo fileId to be saved (new client):', fileId);
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
    console.error('Upload error (new client):', err);
    res.status(500).json({ status: 'failure', message: err.message || 'File upload failed.' });
  }
});

// Upload file for existing client and docType
app.post('/clientdetails/:id/upload/:docType', async (req, res) => {
  try {
    const { id, docType } = req.params;
    console.log('Existing client file upload request for id:', id, 'docType:', docType);
    const { catalyst } = res.locals;
    const folderId = DOC_TYPE_TO_FOLDER_ID[docType];
    console.log('Folder ID for docType:', docType, 'is:', folderId);
    
    if (!folderId) {
      console.error('Invalid document type:', docType);
      return res.status(400).send({ status: 'failure', message: 'Invalid document type.' });
    }
    
    if (!req.files || !req.files.file) {
      console.error('No file uploaded for client id:', id, 'docType:', docType);
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

      // Debug logging for Logo uploads
      if (docType === 'Logo') {
        console.log('Logo upload response:', uploadResp);
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

      // Debug logging for Logo fileId
      if (docType === 'Logo') {
        console.log('Logo fileId to be saved:', fileId);
      }

      // Update ClientDetails row - store file ID in Logo column
      const table = catalyst.datastore().table('ClientDetails');
      await table.updateRow({
        ROWID: id,
        Logo: fileId // Store file ID in the existing Logo column
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

// Download file for client and docType
app.get('/clientdetails/:id/file/:docType', async (req, res) => {
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
    
    const table = catalyst.datastore().table('ClientDetails');
    const client = await table.getRow(id);
    const fileId = client.Logo; // Get file ID from Logo column
    const fileName = 'logo_file'; // Default filename since we don't store it separately
    
    console.log('File info:', { fileId, fileName });
    
    if (!fileId) {
      console.error('No file ID found for client:', id, 'docType:', docType);
      return res.status(404).send({ status: 'failure', message: 'File not found for this client.' });
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

// Delete file for client and docType
app.delete('/clientdetails/:id/file/:docType', async (req, res) => {
  try {
    const { id, docType } = req.params;
    const { catalyst } = res.locals;
    const folderId = DOC_TYPE_TO_FOLDER_ID[docType];
    
    if (!folderId) {
      return res.status(400).send({ status: 'failure', message: 'Invalid document type.' });
    }
    
    const table = catalyst.datastore().table('ClientDetails');
    const client = await table.getRow(id);
    const fileId = client.Logo; // Get file ID from Logo column
    
    if (!fileId) {
      return res.status(404).send({ status: 'failure', message: 'File not found for this client.' });
    }
    
    await catalyst.filestore().folder(folderId).deleteFile(fileId);
    
    // Clear Logo column
    await table.updateRow({
      ROWID: id,
      Logo: null
    });
    
    res.status(200).send({ status: 'success' });
  } catch (err) {
    console.log(err);
    res.status(500).send({ status: 'failure', message: err.message || 'File delete failed.' });
  }
});

module.exports = app;
