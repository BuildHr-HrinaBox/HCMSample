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
  res.status(200).json({ status: 'success', message: 'certificate_function ready' });
});

// Mapping of document types to Catalyst File Store folder IDs
const DOC_TYPE_TO_FOLDER_ID = {
  GSTFile: '28341000000074427', // GST File folder ID
  CompanyPANFile: '28341000000074446', // Company PAN File folder ID
  IncorporationFile: '28341000000074474', // Incorporation File folder ID
  MOAFile: '28341000000074493', // MOA File folder ID
  AOAFile: '28341000000074512', // AOA File folder ID
  EPFFile: '28341000000074531', // EPF File folder ID
  ESIFile: '28341000000074559', // ESI File folder ID
  FactoryLicenseFile: '28341000000074578', // Factory License File folder ID
  LAFFile: '	28341000000074597', // LAF File folder ID
  PTFile: '28341000000074625', // PT File folder ID
  Form3TNNFHFile: '28341000000074644' // Form 3TNNFH File folder ID
};

// Test endpoint to check table access
app.get('/test-table', async (req, res) => {
  try {
    console.log('Testing certificate table access...');
    const { catalyst } = res.locals;
    const table = catalyst.datastore().table('Certificate');
    
    // Try to get table info
    console.log('Table object:', table);
    
    // Try a simple query
    const zcql = catalyst.zcql();
    const testQuery = await zcql.executeZCQLQuery('SELECT COUNT(ROWID) as count FROM Certificate');
    console.log('Test query result:', testQuery);
    
    res.status(200).send({ 
      status: 'success', 
      message: 'Certificate table access successful',
      tableInfo: {
        tableName: 'Certificate',
        recordCount: testQuery[0]?.Certificate?.count || 0
      }
    });
  } catch (err) {
    console.error('Certificate table test error:', err);
    res.status(500).send({ 
      status: 'failure', 
      message: err.message || 'Failed to access certificate table',
      error: err.toString()
    });
  }
});

// Test endpoint to check File Store access
app.get('/test-filestore', async (req, res) => {
  try {
    console.log('Testing Certificate File Store access...');
    const { catalyst } = res.locals;
    const filestore = catalyst.filestore();
    
    // Try to access the GSTFile folder
    const folder = filestore.folder('GSTFile');
    console.log('Folder object:', folder);
    
    // Try to list files in the folder
    const files = await folder.getAllFiles();
    console.log('Files in GSTFile folder:', files);
    
    res.status(200).send({ 
      status: 'success', 
      message: 'Certificate File Store access successful',
      fileStoreInfo: {
        folderName: 'GSTFile',
        fileCount: files.length || 0,
        files: files.map(f => ({ name: f.name, id: f.id }))
      }
    });
  } catch (err) {
    console.error('Certificate File Store test error:', err);
    res.status(500).send({ 
      status: 'failure', 
      message: err.message || 'Failed to access Certificate File Store',
      error: err.toString()
    });
  }
});

// Create Certificate
app.post('/certificate', async (req, res) => {
  try {
    console.log('Received certificate request body:', JSON.stringify(req.body, null, 2));
    
    const {
      GSTNo,
      GSTDate,
      CompanyPANNo,
      CompanyPANDate,
      IncorporationNo,
      IncorporationDate,
      MOANo,
      MOADate,
      AOANo,
      AOADate,
      EPFNo,
      EPFDate,
      ESINo,
      ESIDate,
      FactoryLicenseNo,
      FactoryLicenseDate,
      LAFNo,
      LRFDate,
      PTNo,
      PTDate,
      GSTFile,
      CompanyPANFile,
      IncorporationFile,
      MOAFile,
      AOAFile,
      EPFFile,
      ESIFile,
      FactoryLicenseFile,
      LAFFile,
      PTFile,
      Form3TNNFHFile
    } = req.body;

    const { catalyst } = res.locals;
    const table = catalyst.datastore().table('Certificate');
    
    const insertData = {
      GSTNo: GSTNo || null,
      GSTDate: GSTDate || null,
      CompanyPANNo: CompanyPANNo || null,
      CompanyPANDate: CompanyPANDate || null,
      IncorporationNo: IncorporationNo || null,
      IncorporationDate: IncorporationDate || null,
      MOANo: MOANo || null,
      MOADate: MOADate || null,
      AOANo: AOANo || null,
      AOADate: AOADate || null,
      EPFNo: EPFNo || null,
      EPFDate: EPFDate || null,
      ESINo: ESINo || null,
      ESIDate: ESIDate || null,
      FactoryLicenseNo: FactoryLicenseNo || null,
      FactoryLicenseDate: FactoryLicenseDate || null,
      LAFNo: LAFNo || null,
      LRFDate: LRFDate || null,
      PTNo: PTNo || null,
      PTDate: PTDate || null,
      GSTFile: GSTFile || null,
      CompanyPANFile: CompanyPANFile || null,
      IncorporationFile: IncorporationFile || null,
      MOAFile: MOAFile || null,
      AOAFile: AOAFile || null,
      EPFFile: EPFFile || null,
      ESIFile: ESIFile || null,
      FactoryLicenseFile: FactoryLicenseFile || null,
      LAFFile: LAFFile || null,
      PTFile: PTFile || null,
      Form3TNNFHFile: Form3TNNFHFile || null
    };
    
    console.log('Inserting certificate data to Certificate table:', JSON.stringify(insertData, null, 2));
    
    const insertResp = await table.insertRow(insertData);
    console.log('Insert response:', JSON.stringify(insertResp, null, 2));
    
    const created = await table.getRow(insertResp.ROWID);
    console.log('Created certificate record:', JSON.stringify(created, null, 2));
    
    res.status(200).send({ status: 'success', data: { certificate: created } });
  } catch (err) {
    console.error('Error creating certificate:', err);
    res.status(400).send({ status: 'failure', message: err.message || 'Failed to create certificate.' });
  }
});

// List Certificates with optional pagination
app.get('/certificate', async (req, res) => {
  try {
    console.log('Fetching certificate details...');
    const { catalyst } = res.locals;
    const zcql = catalyst.zcql();
    const page = parseInt(req.query.page, 10) || 1;
    const perPage = parseInt(req.query.perPage, 10) || 50;
    const returnAll = !req.query.page && !req.query.perPage;
    const limitClause = returnAll ? '' : `LIMIT ${(page - 1) * perPage + 1},${perPage}`;
    
    console.log('Executing count query...');
    const countRows = await zcql.executeZCQLQuery('SELECT COUNT(ROWID) as count FROM Certificate');
    const total = parseInt(countRows[0].Certificate.count, 10) || 0;
    console.log('Total certificate records:', total);
    
    console.log('Executing data query...');
    // Split the query to avoid the 30 column limit
    const rows = await zcql.executeZCQLQuery(`SELECT ROWID, GSTNo, GSTDate, CompanyPANNo, CompanyPANDate, IncorporationNo, IncorporationDate, MOANo, MOADate, AOANo, AOADate, EPFNo, EPFDate, ESINo, ESIDate, FactoryLicenseNo, FactoryLicenseDate, LAFNo, LRFDate, PTNo, PTDate, CREATEDTIME, MODIFIEDTIME FROM Certificate ORDER BY ROWID DESC ${limitClause}`);
    console.log('Raw rows from database (basic info):', JSON.stringify(rows, null, 2));
    
    // Get file information separately if needed
    let fileRows = [];
    if (rows.length > 0) {
      const rowIds = rows.map(r => r.Certificate.ROWID).join(',');
      try {
        fileRows = await zcql.executeZCQLQuery(`SELECT ROWID, GSTFile, CompanyPANFile, IncorporationFile, MOAFile, AOAFile, EPFFile, ESIFile, FactoryLicenseFile, LAFFile, PTFile, Form3TNNFHFile FROM Certificate WHERE ROWID IN (${rowIds})`);
        console.log('File rows from database:', JSON.stringify(fileRows, null, 2));
      } catch (fileError) {
        console.error('Error fetching file data:', fileError);
        // Continue without file data
      }
    }
    
    const certificateDetails = rows.map(r => {
      // Find corresponding file data
      const fileData = fileRows.find(fr => fr.Certificate.ROWID === r.Certificate.ROWID);
      
      return {
        id: r.Certificate.ROWID,
        GSTNo: r.Certificate.GSTNo,
        GSTDate: r.Certificate.GSTDate,
        CompanyPANNo: r.Certificate.CompanyPANNo,
        CompanyPANDate: r.Certificate.CompanyPANDate,
        IncorporationNo: r.Certificate.IncorporationNo,
        IncorporationDate: r.Certificate.IncorporationDate,
        MOANo: r.Certificate.MOANo,
        MOADate: r.Certificate.MOADate,
        AOANo: r.Certificate.AOANo,
        AOADate: r.Certificate.AOADate,
        EPFNo: r.Certificate.EPFNo,
        EPFDate: r.Certificate.EPFDate,
        ESINo: r.Certificate.ESINo,
        ESIDate: r.Certificate.ESIDate,
        FactoryLicenseNo: r.Certificate.FactoryLicenseNo,
        FactoryLicenseDate: r.Certificate.FactoryLicenseDate,
        LAFNo: r.Certificate.LAFNo,
        LRFDate: r.Certificate.LRFDate,
        PTNo: r.Certificate.PTNo,
        PTDate: r.Certificate.PTDate,
        GSTFile: fileData ? fileData.Certificate.GSTFile : null,
        CompanyPANFile: fileData ? fileData.Certificate.CompanyPANFile : null,
        IncorporationFile: fileData ? fileData.Certificate.IncorporationFile : null,
        MOAFile: fileData ? fileData.Certificate.MOAFile : null,
        AOAFile: fileData ? fileData.Certificate.AOAFile : null,
        EPFFile: fileData ? fileData.Certificate.EPFFile : null,
        ESIFile: fileData ? fileData.Certificate.ESIFile : null,
        FactoryLicenseFile: fileData ? fileData.Certificate.FactoryLicenseFile : null,
        LAFFile: fileData ? fileData.Certificate.LAFFile : null,
        PTFile: fileData ? fileData.Certificate.PTFile : null,
        Form3TNNFHFile: fileData ? fileData.Certificate.Form3TNNFHFile : null,
        createdTime: r.Certificate.CREATEDTIME,
        modifiedTime: r.Certificate.MODIFIEDTIME
      };
    });
    
    console.log('Processed certificate details:', JSON.stringify(certificateDetails, null, 2));
    res.status(200).send({ status: 'success', data: { certificateDetails, total, hasMore: returnAll ? false : page * perPage < total } });
  } catch (err) {
    console.error('Error fetching certificate details:', err);
    res.status(500).send({ status: 'failure', message: err.message || 'Failed to fetch certificate details.' });
  }
});

// Update Certificate
app.put('/certificate/:ROWID', async (req, res) => {
  try {
    const { ROWID } = req.params;
    const {
      GSTNo,
      GSTDate,
      CompanyPANNo,
      CompanyPANDate,
      IncorporationNo,
      IncorporationDate,
      MOANo,
      MOADate,
      AOANo,
      AOADate,
      EPFNo,
      EPFDate,
      ESINo,
      ESIDate,
      FactoryLicenseNo,
      FactoryLicenseDate,
      LAFNo,
      LRFDate,
      PTNo,
      PTDate,
      GSTFile,
      CompanyPANFile,
      IncorporationFile,
      MOAFile,
      AOAFile,
      EPFFile,
      ESIFile,
      FactoryLicenseFile,
      LAFFile,
      PTFile,
      Form3TNNFHFile
    } = req.body;

    const { catalyst } = res.locals;
    const table = catalyst.datastore().table('Certificate');
    
    await table.updateRow({
      ROWID,
      GSTNo: GSTNo || null,
      GSTDate: GSTDate || null,
      CompanyPANNo: CompanyPANNo || null,
      CompanyPANDate: CompanyPANDate || null,
      IncorporationNo: IncorporationNo || null,
      IncorporationDate: IncorporationDate || null,
      MOANo: MOANo || null,
      MOADate: MOADate || null,
      AOANo: AOANo || null,
      AOADate: AOADate || null,
      EPFNo: EPFNo || null,
      EPFDate: EPFDate || null,
      ESINo: ESINo || null,
      ESIDate: ESIDate || null,
      FactoryLicenseNo: FactoryLicenseNo || null,
      FactoryLicenseDate: FactoryLicenseDate || null,
      LAFNo: LAFNo || null,
      LRFDate: LRFDate || null,
      PTNo: PTNo || null,
      PTDate: PTDate || null,
      GSTFile: GSTFile || null,
      CompanyPANFile: CompanyPANFile || null,
      IncorporationFile: IncorporationFile || null,
      MOAFile: MOAFile || null,
      AOAFile: AOAFile || null,
      EPFFile: EPFFile || null,
      ESIFile: ESIFile || null,
      FactoryLicenseFile: FactoryLicenseFile || null,
      LAFFile: LAFFile || null,
      PTFile: PTFile || null,
      Form3TNNFHFile: Form3TNNFHFile || null
    });
    
    const updated = await table.getRow(ROWID);
    res.status(200).send({ status: 'success', data: { certificate: updated } });
  } catch (err) {
    res.status(400).send({ status: 'failure', message: err.message || 'Failed to update certificate.' });
  }
});

// Delete Certificate
app.delete('/certificate/:ROWID', async (req, res) => {
  try {
    const { ROWID } = req.params;
    const { catalyst } = res.locals;
    const table = catalyst.datastore().table('Certificate');
    await table.deleteRow(ROWID);
    res.status(200).send({ status: 'success', data: { id: ROWID } });
  } catch (err) {
    res.status(500).send({ status: 'failure', message: err.message || 'Failed to delete certificate.' });
  }
});

// Upload file for new certificate (before creation)
app.post('/certificate/upload/:docType', async (req, res) => {
  try {
    const { docType } = req.params;
    console.log('New certificate file upload request for docType:', docType);
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
      console.log('File upload response (new certificate):', uploadResp);

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
        console.error('Unexpected upload response (new certificate):', uploadResp);
        return res.status(500).json({ status: 'failure', message: 'File upload failed or invalid response from Catalyst.' });
      }

      // Debug logging for fileId
      console.log('FileId to be saved (new certificate):', fileId);

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
    console.error('Upload error (new certificate):', err);
    res.status(500).json({ status: 'failure', message: err.message || 'File upload failed.' });
  }
});

// Upload file for existing certificate and docType
app.post('/certificate/:id/upload/:docType', async (req, res) => {
  try {
    const { id, docType } = req.params;
    console.log('Existing certificate file upload request for id:', id, 'docType:', docType);
    const { catalyst } = res.locals;
    const folderId = DOC_TYPE_TO_FOLDER_ID[docType];
    console.log('Folder ID for docType:', docType, 'is:', folderId);
    
    if (!folderId) {
      console.error('Invalid document type:', docType);
      return res.status(400).send({ status: 'failure', message: 'Invalid document type.' });
    }
    
    if (!req.files || !req.files.file) {
      console.error('No file uploaded for certificate id:', id, 'docType:', docType);
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

      // Update Certificate row - store file ID in appropriate column
      const table = catalyst.datastore().table('Certificate');
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

// Download file for certificate and docType
app.get('/certificate/:id/file/:docType', async (req, res) => {
  try {
    const { id, docType } = req.params;
    console.log('File request:', { id, docType });
    const { catalyst } = res.locals;
    const folderId = DOC_TYPE_TO_FOLDER_ID[docType];
    console.log('Folder ID for docType:', docType, 'is:', folderId);
    
    if (!folderId) {
      console.error('Invalid document type:', docType);
      return res.status(400).send({ status: 'failure', message: 'Invalid document type.' });
    }
    
    const table = catalyst.datastore().table('Certificate');
    const certificate = await table.getRow(id);
    const fileId = certificate[docType]; // Get file ID from the appropriate column
    const fileName = `${docType}_file`; // Default filename since we don't store it separately
    
    console.log('File info:', { fileId, fileName });
    
    if (!fileId) {
      console.error('No file ID found for certificate:', id, 'docType:', docType);
      return res.status(404).send({ status: 'failure', message: 'File not found for this certificate.' });
    }
    
    const fileBuffer = await catalyst.filestore().folder(folderId).downloadFile(fileId);
    
    // Check if this is a download request (convert to PDF) or view request (original format)
    const isDownloadRequest = req.query.download === 'true';
    
    let finalBuffer = fileBuffer;
    let finalFileName = fileName;
    let contentType = 'application/pdf';
    
    if (isDownloadRequest) {
      // For download requests, always convert to PDF format
      finalFileName = fileName.replace(/\.[^/.]+$/, '') + '.pdf';
      
      // Convert different file types to PDF
      const fileExtension = fileName.toLowerCase().split('.').pop();
      
      if (['jpg', 'jpeg', 'png', 'gif', 'webp'].includes(fileExtension)) {
        // Convert image to PDF
        finalBuffer = await convertImageToPDF(fileBuffer, fileExtension);
      } else if (['doc', 'docx'].includes(fileExtension)) {
        // For Word documents, create a PDF with extracted text content
        finalBuffer = await convertDocumentToPDF(fileBuffer, fileName);
      } else if (['txt'].includes(fileExtension)) {
        // Convert text to PDF
        finalBuffer = await convertTextToPDF(fileBuffer);
      } else if (fileExtension === 'pdf') {
        // Already PDF, use as-is
        finalBuffer = fileBuffer;
      } else {
        // For other formats, create a simple PDF wrapper
        finalBuffer = await createPDFWrapper(fileBuffer, fileName);
      }
    } else {
      // For view requests, determine content type based on original file extension
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
      
      contentType = getContentType(fileName);
    }
    
    console.log('Sending file:', { fileName: finalFileName, contentType, fileSize: finalBuffer.length, isDownloadRequest });
    
    const headers = {
      'Content-Type': contentType,
      'Content-Length': finalBuffer.length,
      'Cache-Control': 'no-cache',
      'Pragma': 'no-cache'
    };
    
    // Set appropriate Content-Disposition header
    if (isDownloadRequest) {
      // For downloading: attachment to force download
      headers['Content-Disposition'] = `attachment; filename="${finalFileName}"`;
    } else {
      // For viewing: inline to display in browser
      headers['Content-Disposition'] = `inline; filename="${finalFileName}"`;
    }
    
    res.writeHead(200, headers);
    res.end(finalBuffer);
    console.log('File request completed for:', finalFileName, isDownloadRequest ? '(download as PDF)' : '(view original)');
  } catch (err) {
    console.log(err);
    res.status(500).send({ status: 'failure', message: err.message || 'File request failed.' });
  }
});

// Helper function to convert image to PDF
async function convertImageToPDF(imageBuffer, imageType) {
  // Create a simple PDF with the image embedded
  const pdfContent = `%PDF-1.4
1 0 obj
<<
/Type /Catalog
/Pages 2 0 R
>>
endobj

2 0 obj
<<
/Type /Pages
/Kids [3 0 R]
/Count 1
>>
endobj

3 0 obj
<<
/Type /Page
/Parent 2 0 R
/MediaBox [0 0 612 792]
/Resources <<
/XObject <<
/Im1 4 0 R
>>
>>
/Contents 5 0 R
>>
endobj

4 0 obj
<<
/Type /XObject
/Subtype /Image
/Width 400
/Height 300
/ColorSpace /DeviceRGB
/BitsPerComponent 8
/Length ${imageBuffer.length}
/Filter /DCTDecode
>>
stream
${imageBuffer.toString('binary')}
endstream
endobj

5 0 obj
<<
/Length 44
>>
stream
q
400 0 0 300 0 0 cm
/Im1 Do
Q
endstream
endobj

xref
0 6
0000000000 65535 f 
0000000009 00000 n 
0000000058 00000 n 
0000000115 00000 n 
0000000204 00000 n 
0000000300 00000 n 
trailer
<<
/Size 6
/Root 1 0 R
>>
startxref
400
%%EOF`;

  return Buffer.from(pdfContent, 'binary');
}

// Helper function to convert document to PDF
async function convertDocumentToPDF(docBuffer, fileName) {
  try {
    // For Word documents, we'll create a more comprehensive PDF
    // that includes the document content as text
    
    // Extract text content from the document buffer
    let textContent = '';
    
    try {
      // For .docx files, we can extract some text content
      if (fileName.toLowerCase().endsWith('.docx')) {
        textContent = await extractTextFromDocx(docBuffer);
      } else if (fileName.toLowerCase().endsWith('.doc')) {
        textContent = await extractTextFromDoc(docBuffer);
      } else {
        textContent = 'Document content could not be extracted.';
      }
    } catch (extractError) {
      console.log('Text extraction failed:', extractError);
      textContent = 'Document content could not be extracted.';
    }
    
    // Split text into lines and create PDF content
    const lines = textContent.split('\n').slice(0, 50); // Limit to 50 lines
    let pdfContent = `%PDF-1.4
1 0 obj
<<
/Type /Catalog
/Pages 2 0 R
>>
endobj

2 0 obj
<<
/Type /Pages
/Kids [3 0 R]
/Count 1
>>
endobj

3 0 obj
<<
/Type /Page
/Parent 2 0 R
/MediaBox [0 0 612 792]
/Contents 4 0 R
>>
endobj

4 0 obj
<<
/Length ${textContent.length * 8 + 200}
>>
stream
BT
/F1 14 Tf
72 750 Td
(Document: ${fileName}) Tj
0 -25 Td
/F1 10 Tf
(Converted to PDF on: ${new Date().toISOString()}) Tj
0 -20 Td
(Original file size: ${docBuffer.length} bytes) Tj
0 -30 Td
(--- Document Content ---) Tj
0 -20 Td
`;

    // Add document content lines
    lines.forEach((line, index) => {
      if (index > 0) {
        pdfContent += '0 -12 Td\n';
      }
      // Escape special characters and limit line length
      const escapedLine = line.replace(/[()\\]/g, '\\$&').substring(0, 80);
      pdfContent += `(${escapedLine}) Tj\n`;
    });

    pdfContent += `0 -20 Td
(--- End of Document ---) Tj
ET
endstream
endobj

xref
0 5
0000000000 65535 f 
0000000009 00000 n 
0000000058 00000 n 
0000000115 00000 n 
0000000204 00000 n 
trailer
<<
/Size 5
/Root 1 0 R
>>
startxref
${pdfContent.length - 100}
%%EOF`;

    return Buffer.from(pdfContent);
  } catch (error) {
    console.error('Error converting document to PDF:', error);
    // Fallback to simple PDF
    return createSimpleDocumentPDF(fileName, docBuffer.length);
  }
}

// Helper function to extract text from DOCX files
async function extractTextFromDocx(docBuffer) {
  try {
    // For DOCX files, we can try to extract text from the XML structure
    // This is a simplified extraction - in production, you'd want to use a proper library
    
    const bufferString = docBuffer.toString('binary');
    
    // Look for text content in the DOCX structure
    // DOCX files are ZIP archives containing XML files
    // We'll do a simple text extraction from the document.xml
    
    // Simple regex to find text content (this is basic and may not work for all cases)
    const textMatches = bufferString.match(/<w:t[^>]*>([^<]*)<\/w:t>/g);
    
    if (textMatches && textMatches.length > 0) {
      let extractedText = '';
      textMatches.forEach(match => {
        const textContent = match.replace(/<w:t[^>]*>|<\/w:t>/g, '');
        if (textContent.trim()) {
          extractedText += textContent + ' ';
        }
      });
      return extractedText.trim() || 'Text content could not be extracted from DOCX file.';
    }
    
    return 'Text content could not be extracted from DOCX file.';
  } catch (error) {
    console.error('DOCX text extraction error:', error);
    return 'Error extracting text from DOCX file.';
  }
}

// Helper function to extract text from DOC files
async function extractTextFromDoc(docBuffer) {
  try {
    // For .doc files, text extraction is more complex
    // This is a simplified approach - in production, you'd want to use a proper library
    
    const bufferString = docBuffer.toString('binary');
    
    // Simple approach: look for readable text patterns
    // This is very basic and may not work reliably
    const textPattern = /[A-Za-z0-9\s.,!?;:'"()-]{10,}/g;
    const matches = bufferString.match(textPattern);
    
    if (matches && matches.length > 0) {
      // Join matches and clean up
      let extractedText = matches.join(' ').replace(/\s+/g, ' ').trim();
      
      // Limit length and clean up
      if (extractedText.length > 1000) {
        extractedText = extractedText.substring(0, 1000) + '...';
      }
      
      return extractedText || 'Text content could not be extracted from DOC file.';
    }
    
    return 'Text content could not be extracted from DOC file.';
  } catch (error) {
    console.error('DOC text extraction error:', error);
    return 'Error extracting text from DOC file.';
  }
}

// Helper function to create a simple document PDF as fallback
function createSimpleDocumentPDF(fileName, fileSize) {
  const pdfContent = `%PDF-1.4
1 0 obj
<<
/Type /Catalog
/Pages 2 0 R
>>
endobj

2 0 obj
<<
/Type /Pages
/Kids [3 0 R]
/Count 1
>>
endobj

3 0 obj
<<
/Type /Page
/Parent 2 0 R
/MediaBox [0 0 612 792]
/Contents 4 0 R
>>
endobj

4 0 obj
<<
/Length 300
>>
stream
BT
/F1 14 Tf
72 750 Td
(Document: ${fileName}) Tj
0 -25 Td
/F1 10 Tf
(This document was converted to PDF format) Tj
0 -20 Td
(Original file size: ${fileSize} bytes) Tj
0 -20 Td
(Conversion date: ${new Date().toISOString()}) Tj
0 -20 Td
(Note: This is a converted PDF version) Tj
0 -20 Td
(For full content, please view the original file) Tj
ET
endstream
endobj

xref
0 5
0000000000 65535 f 
0000000009 00000 n 
0000000058 00000 n 
0000000115 00000 n 
0000000204 00000 n 
trailer
<<
/Size 5
/Root 1 0 R
>>
startxref
500
%%EOF`;

  return Buffer.from(pdfContent);
}

// Helper function to convert text to PDF
async function convertTextToPDF(textBuffer) {
  const textContent = textBuffer.toString('utf8');
  const lines = textContent.split('\n');
  
  let pdfContent = `%PDF-1.4
1 0 obj
<<
/Type /Catalog
/Pages 2 0 R
>>
endobj

2 0 obj
<<
/Type /Pages
/Kids [3 0 R]
/Count 1
>>
endobj

3 0 obj
<<
/Type /Page
/Parent 2 0 R
/MediaBox [0 0 612 792]
/Contents 4 0 R
>>
endobj

4 0 obj
<<
/Length ${textContent.length * 10}
>>
stream
BT
/F1 10 Tf
72 720 Td
`;

  lines.forEach((line, index) => {
    if (index > 0) {
      pdfContent += '0 -12 Td\n';
    }
    pdfContent += `(${line.replace(/[()\\]/g, '\\$&')}) Tj\n`;
  });

  pdfContent += `ET
endstream
endobj

xref
0 5
0000000000 65535 f 
0000000009 00000 n 
0000000058 00000 n 
0000000115 00000 n 
0000000204 00000 n 
trailer
<<
/Size 5
/Root 1 0 R
>>
startxref
${pdfContent.length - 100}
%%EOF`;

  return Buffer.from(pdfContent);
}

// Helper function to create PDF wrapper for unknown formats
async function createPDFWrapper(fileBuffer, fileName) {
  const pdfContent = `%PDF-1.4
1 0 obj
<<
/Type /Catalog
/Pages 2 0 R
>>
endobj

2 0 obj
<<
/Type /Pages
/Kids [3 0 R]
/Count 1
>>
endobj

3 0 obj
<<
/Type /Page
/Parent 2 0 R
/MediaBox [0 0 612 792]
/Contents 4 0 R
>>
endobj

4 0 obj
<<
/Length 300
>>
stream
BT
/F1 12 Tf
72 720 Td
(File: ${fileName}) Tj
0 -20 Td
(File size: ${fileBuffer.length} bytes) Tj
0 -20 Td
(File type: ${fileName.split('.').pop()}) Tj
0 -20 Td
(Converted to PDF on: ${new Date().toISOString()}) Tj
0 -20 Td
(Note: This file was converted to PDF format) Tj
ET
endstream
endobj

xref
0 5
0000000000 65535 f 
0000000009 00000 n 
0000000058 00000 n 
0000000115 00000 n 
0000000204 00000 n 
trailer
<<
/Size 5
/Root 1 0 R
>>
startxref
500
%%EOF`;

  return Buffer.from(pdfContent);
}

// Delete file for certificate and docType
app.delete('/certificate/:id/file/:docType', async (req, res) => {
  try {
    const { id, docType } = req.params;
    const { catalyst } = res.locals;
    const folderId = DOC_TYPE_TO_FOLDER_ID[docType];
    
    if (!folderId) {
      return res.status(400).send({ status: 'failure', message: 'Invalid document type.' });
    }
    
    const table = catalyst.datastore().table('Certificate');
    const certificate = await table.getRow(id);
    const fileId = certificate[docType]; // Get file ID from the appropriate column
    
    if (!fileId) {
      return res.status(404).send({ status: 'failure', message: 'File not found for this certificate.' });
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