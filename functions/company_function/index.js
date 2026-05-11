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
  res.status(200).json({ status: 'success', message: 'company_function ready' });
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
    const table = catalyst.datastore().table('Company');
    
    // Try to get table info
    console.log('Table object:', table);
    
    // Try a simple query
    const zcql = catalyst.zcql();
    const testQuery = await zcql.executeZCQLQuery('SELECT COUNT(ROWID) as count FROM Company');
    console.log('Test query result:', testQuery);
    
    res.status(200).send({ 
      status: 'success', 
      message: 'Table access successful',
      tableInfo: {
        tableName: 'Company',
        recordCount: testQuery[0]?.Company?.count || 0
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
    
    // Try to access the SafetyOfficerAppoitnmentorder folder
    const folder = filestore.folder('SafetyOfficerAppoitnmentorder');
    console.log('Folder object:', folder);
    
    // Try to list files in the folder
    const files = await folder.getAllFiles();
    console.log('Files in SafetyOfficerAppoitnmentorder folder:', files);
    
    res.status(200).send({ 
      status: 'success', 
      message: 'File Store access successful',
      fileStoreInfo: {
        folderName: 'SafetyOfficerAppoitnmentorder',
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

/** Doctro* columns may be mandatory in Data Store even when the UI omits doctor fields. */
function normalizeDoctorTextField(value, placeholder = 'N/A') {
  const s = value == null ? '' : String(value).trim();
  return s === '' ? placeholder : s;
}

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

// ===== Company CRUD =====
// Table: Company (columns: CompanyName, CompanyPANNumber, OwnerName, OwnerPAN, OwnerAaadhar, OwnerDesignation, SafetyOfficerName, SafetyOfficerPhone, DoctroName, DoctroPhone, DoctroIndustry, SafetyOfficerAppoitnmentorder, SafetyOfficerApporvalcopy, DoctroAppoitnmentorder, DoctroApporvalcopy, HeadHRSign, HeadHRSeal)

// Mapping of document types to Catalyst File Store folder IDs
const DOC_TYPE_TO_FOLDER_ID = {
  DoctroApporvalcopy: '31459000000501570',
  DoctroAppoitnmentorder: '31459000000501551',
  SafetyOfficerApporvalcopy: '31459000000501532',
  SafetyOfficerAppoitnmentorder: '31459000000501513',
  HeadHRSign: '31459000000521148',
  HeadHRSeal: '31459000000521129'
};

const DOC_TYPE_TO_COMPANY_KEY = {
  SafetyOfficerAppoitnmentorder: 'safetyOfficerAppoitnmentorder',
  SafetyOfficerApporvalcopy: 'safetyOfficerApporvalcopy',
  DoctroAppoitnmentorder: 'doctroAppoitnmentorder',
  DoctroApporvalcopy: 'doctroApporvalcopy',
  HeadHRSign: 'headHRSign',
  HeadHRSeal: 'headHRSeal'
};

// Helper to get file columns for a docType
function getFileColumns(docType) {
  return {
    fileIdCol: `${docType}FileId`,
    fileNameCol: `${docType}FileName`
  };
}

/** Datastore getRow may use table column names (PascalCase) or the same keys as the REST list (camelCase). */
function pickFileIdFromCompanyRow(company, docType) {
  if (!company || !docType) return null;
  const camel = DOC_TYPE_TO_COMPANY_KEY[docType];
  const keys = [docType, camel].filter((k, i, a) => k != null && k !== '' && a.indexOf(k) === i);
  for (const k of keys) {
    const v = company[k];
    if (v != null && String(v).trim() !== '') return v;
  }
  return null;
}

function pickDisplayNameFromFileDetails(raw, depth = 0) {
  if (raw == null || depth > 5) return '';
  const details = Array.isArray(raw) ? raw[0] : raw;
  if (!details || typeof details !== 'object') return '';

  const keyVariants = [
    'file_name',
    'fileName',
    'File_Name',
    'name',
    'Name',
    'filename',
    'original_file_name',
    'originalFileName'
  ];

  const sources = [
    details,
    details.file && typeof details.file === 'object' ? details.file : null,
    details.data && typeof details.data === 'object' ? details.data : null,
    Array.isArray(details.file_details) && details.file_details[0] ? details.file_details[0] : null,
    Array.isArray(details.files) && details.files[0] ? details.files[0] : null
  ].filter(Boolean);

  for (const src of sources) {
    for (const k of keyVariants) {
      const s = String(src[k] || '').trim();
      if (s) return s;
    }
  }

  const recurseKeys = ['data', 'file', 'file_details', 'result'];
  for (const rk of recurseKeys) {
    const child = details[rk];
    if (child == null) continue;
    const nested = pickDisplayNameFromFileDetails(child, depth + 1);
    if (nested) return nested;
  }
  return '';
}

async function resolveUploadedFileName(catalyst, docType, fileId) {
  const folderId = DOC_TYPE_TO_FOLDER_ID[docType];
  if (!folderId || !fileId) return '';
  try {
    const folder = catalyst.filestore().folder(folderId);
    const targetId = String(fileId).trim();
    if (!targetId) return '';

    try {
      const details = await folder.getFileDetails(fileId);
      const byDetails = pickDisplayNameFromFileDetails(details);
      if (byDetails) return byDetails;
    } catch (_) {
      // Optional list fallback below
    }

    // Older code called getAllFiles(); the stock Node SDK Folder has no such method — calling it threw and hid real names.
    if (typeof folder.getAllFiles === 'function') {
      try {
        const files = await folder.getAllFiles();
        if (Array.isArray(files)) {
          const hit = files.find((f) => String(f?.id ?? f?.ID ?? f?.file_id ?? '').trim() === targetId);
          if (hit) {
            const byList = pickDisplayNameFromFileDetails(hit);
            if (byList) return byList;
          }
        }
      } catch (_) {
        // ignore
      }
    }
    return '';
  } catch (err) {
    console.warn('Could not resolve file name for docType:', docType, 'fileId:', fileId, err?.message || err);
    return '';
  }
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
      const folder = filestore.folder('SafetyOfficerAppoitnmentorder'); // Use the SafetyOfficerAppoitnmentorder folder
      
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

// Create Company
app.post('/company', async (req, res) => {
  try {
    console.log('Received request body:', JSON.stringify(req.body, null, 2));
    
    const {
      companyName,
      companyPANNumber,
      companyAddress,
      incorprationDate,
      incorporationNo,
      gstNo,
      pfNo,
      esiNo,
      companyMail,
      companyPhoneNumber,
      directorName,
      directorPhoneNumber,
      directorMail,
      directorAddress,
      ownerName,
      ownerPAN,
      ownerAaadhar,
      ownerDesignation,
      safetyOfficerName,
      safetyOfficerPhone,
      doctroName,
      doctroPhone,
      safetyOfficerAppoitnmentorder,
      safetyOfficerApporvalcopy,
      doctroAppoitnmentorder,
      doctroApporvalcopy,
      headHRSign,
      headHRSeal
    } = req.body;

    const { catalyst } = res.locals;
    
    const table = catalyst.datastore().table('Company');
    
    const insertData = {
      CompanyName: companyName,
      CompanyPANNumber: companyPANNumber,
      CompanyAddress: companyAddress,
      IncorprationDate: incorprationDate,
      IncorporationNo: incorporationNo,
      GSTNo: gstNo,
      PFNo: pfNo,
      ESINo: esiNo,
      CompanyMail: companyMail,
      CompanyPhoneNumber: companyPhoneNumber,
      DirectorName: directorName,
      DirectorPhoneNumber: directorPhoneNumber,
      DirectorMail: directorMail,
      DirectorAddress: directorAddress,
      OwnerName: ownerName,
      OwnerPAN: ownerPAN,
      OwnerAaadhar: ownerAaadhar,
      OwnerDesignation: ownerDesignation,
      SafetyOfficerName: safetyOfficerName,
      SafetyOfficerPhone: safetyOfficerPhone,
      DoctroName: normalizeDoctorTextField(doctroName),
      DoctroPhone: normalizeDoctorTextField(doctroPhone),
      SafetyOfficerAppoitnmentorder: safetyOfficerAppoitnmentorder || null,
      SafetyOfficerApporvalcopy: safetyOfficerApporvalcopy || null,
      DoctroAppoitnmentorder: doctroAppoitnmentorder || null,
      DoctroApporvalcopy: doctroApporvalcopy || null,
      HeadHRSign: headHRSign || null,
      HeadHRSeal: headHRSeal || null
    };
    
    console.log('Inserting data to Company table:', JSON.stringify(insertData, null, 2));
    
    const insertResp = await table.insertRow(insertData);
    console.log('Insert response:', JSON.stringify(insertResp, null, 2));
    
    const created = await table.getRow(insertResp.ROWID);
    console.log('Created record:', JSON.stringify(created, null, 2));
    
    res.status(200).send({ status: 'success', data: { company: created } });
  } catch (err) {
    console.error('Error creating company:', err);
    res.status(400).send({ status: 'failure', message: err.message || 'Failed to create company.' });
  }
});

// List Company with optional pagination
app.get('/company', async (req, res) => {
  try {
    console.log('Fetching company details...');
    const { catalyst } = res.locals;
    const zcql = catalyst.zcql();
    const page = parseInt(req.query.page, 10) || 1;
    const perPage = parseInt(req.query.perPage, 10) || 50;
    const returnAll = !req.query.page && !req.query.perPage;
    const limitClause = returnAll ? '' : `LIMIT ${(page - 1) * perPage + 1},${perPage}`;
    
    console.log('Executing count query...');
    const countRows = await zcql.executeZCQLQuery('SELECT COUNT(ROWID) as count FROM Company');
    const total = parseInt(countRows[0].Company.count, 10) || 0;
    console.log('Total records:', total);
    
    console.log('Executing data query...');
    let rows = [];
    try {
      // Preferred query (new schema with HeadHRSign/HeadHRSeal).
      rows = await zcql.executeZCQLQuery(
        `SELECT ROWID, CompanyName, CompanyPANNumber, CompanyAddress, IncorprationDate, IncorporationNo, GSTNo, PFNo, ESINo, CompanyMail, CompanyPhoneNumber, DirectorName, DirectorPhoneNumber, DirectorMail, DirectorAddress, OwnerName, OwnerPAN, OwnerAaadhar, OwnerDesignation, SafetyOfficerName, SafetyOfficerPhone, DoctroName, DoctroPhone, SafetyOfficerAppoitnmentorder, SafetyOfficerApporvalcopy, DoctroAppoitnmentorder, DoctroApporvalcopy, HeadHRSign, HeadHRSeal, CREATEDTIME, MODIFIEDTIME FROM Company ORDER BY ROWID DESC ${limitClause}`
      );
    } catch (schemaErr) {
      // Backward-compatible fallback for environments where new columns are not added yet.
      console.warn('Company list query fallback (legacy schema):', schemaErr.message || schemaErr);
      rows = await zcql.executeZCQLQuery(
        `SELECT ROWID, CompanyName, CompanyPANNumber, CompanyAddress, IncorprationDate, IncorporationNo, GSTNo, PFNo, ESINo, CompanyMail, CompanyPhoneNumber, DirectorName, DirectorPhoneNumber, DirectorMail, DirectorAddress, OwnerName, OwnerPAN, OwnerAaadhar, OwnerDesignation, SafetyOfficerName, SafetyOfficerPhone, DoctroName, DoctroPhone, SafetyOfficerAppoitnmentorder, SafetyOfficerApporvalcopy, DoctroAppoitnmentorder, DoctroApporvalcopy, CREATEDTIME, MODIFIEDTIME FROM Company ORDER BY ROWID DESC ${limitClause}`
      );
    }
    console.log('Raw rows from database:', JSON.stringify(rows, null, 2));
    
    let companyDetails = rows.map(r => ({
      id: r.Company.ROWID,
      companyName: r.Company.CompanyName,
      companyPANNumber: r.Company.CompanyPANNumber,
      companyAddress: r.Company.CompanyAddress,
      incorprationDate: r.Company.IncorprationDate,
      incorporationNo: r.Company.IncorporationNo,
      gstNo: r.Company.GSTNo,
      pfNo: r.Company.PFNo,
      esiNo: r.Company.ESINo,
      companyMail: r.Company.CompanyMail,
      companyPhoneNumber: r.Company.CompanyPhoneNumber,
      directorName: r.Company.DirectorName,
      directorPhoneNumber: r.Company.DirectorPhoneNumber,
      directorMail: r.Company.DirectorMail,
      directorAddress: r.Company.DirectorAddress,
      ownerName: r.Company.OwnerName,
      ownerPAN: r.Company.OwnerPAN,
      ownerAaadhar: r.Company.OwnerAaadhar,
      ownerDesignation: r.Company.OwnerDesignation,
      safetyOfficerName: r.Company.SafetyOfficerName,
      safetyOfficerPhone: r.Company.SafetyOfficerPhone,
      doctroName: r.Company.DoctroName,
      doctroPhone: r.Company.DoctroPhone,
      safetyOfficerAppoitnmentorder: r.Company.SafetyOfficerAppoitnmentorder,
      safetyOfficerApporvalcopy: r.Company.SafetyOfficerApporvalcopy,
      doctroAppoitnmentorder: r.Company.DoctroAppoitnmentorder,
      doctroApporvalcopy: r.Company.DoctroApporvalcopy,
      headHRSign: r.Company.HeadHRSign,
      headHRSeal: r.Company.HeadHRSeal,
      createdTime: r.Company.CREATEDTIME,
      modifiedTime: r.Company.MODIFIEDTIME
    }));

    const docTypes = Object.keys(DOC_TYPE_TO_FOLDER_ID);
    companyDetails = await Promise.all(
      companyDetails.map(async (company) => {
        const documentFileNames = {};
        await Promise.all(
          docTypes.map(async (docType) => {
            const fileId = company[DOC_TYPE_TO_COMPANY_KEY[docType] || docType];
            if (!fileId) return;
            const fileName = await resolveUploadedFileName(catalyst, docType, fileId);
            if (fileName) documentFileNames[docType] = fileName;
          })
        );
        return { ...company, documentFileNames };
      })
    );
    
    console.log('Processed company details:', JSON.stringify(companyDetails, null, 2));
    res.status(200).send({ status: 'success', data: { companyDetails, total, hasMore: returnAll ? false : page * perPage < total } });
  } catch (err) {
    console.error('Error fetching company details:', err);
    res.status(500).send({ status: 'failure', message: err.message || 'Failed to fetch company details.' });
  }
});

// Update Company
app.put('/company/:ROWID', async (req, res) => {
  try {
    const { ROWID } = req.params;
    const {
      companyName,
      companyPANNumber,
      companyAddress,
      incorprationDate,
      incorporationNo,
      gstNo,
      pfNo,
      esiNo,
      companyMail,
      companyPhoneNumber,
      directorName,
      directorPhoneNumber,
      directorMail,
      directorAddress,
      ownerName,
      ownerPAN,
      ownerAaadhar,
      ownerDesignation,
      safetyOfficerName,
      safetyOfficerPhone,
      doctroName,
      doctroPhone,
      safetyOfficerAppoitnmentorder,
      safetyOfficerApporvalcopy,
      doctroAppoitnmentorder,
      doctroApporvalcopy,
      headHRSign,
      headHRSeal
    } = req.body;

    const { catalyst } = res.locals;
    const table = catalyst.datastore().table('Company');
    const updateData = {
      ROWID,
      CompanyName: companyName,
      CompanyPANNumber: companyPANNumber,
      CompanyAddress: companyAddress,
      IncorprationDate: incorprationDate,
      IncorporationNo: incorporationNo,
      GSTNo: gstNo,
      PFNo: pfNo,
      ESINo: esiNo,
      CompanyMail: companyMail,
      CompanyPhoneNumber: companyPhoneNumber,
      DirectorName: directorName,
      DirectorPhoneNumber: directorPhoneNumber,
      DirectorMail: directorMail,
      DirectorAddress: directorAddress,
      OwnerName: ownerName,
      OwnerPAN: ownerPAN,
      OwnerAaadhar: ownerAaadhar,
      OwnerDesignation: ownerDesignation,
      SafetyOfficerName: safetyOfficerName,
      SafetyOfficerPhone: safetyOfficerPhone,
      DoctroName: normalizeDoctorTextField(doctroName),
      DoctroPhone: normalizeDoctorTextField(doctroPhone)
    };

    // Preserve existing uploaded files on normal form updates.
    // Only update a document column when client explicitly sends it.
    if (safetyOfficerAppoitnmentorder !== undefined) {
      updateData.SafetyOfficerAppoitnmentorder = safetyOfficerAppoitnmentorder || null;
    }
    if (safetyOfficerApporvalcopy !== undefined) {
      updateData.SafetyOfficerApporvalcopy = safetyOfficerApporvalcopy || null;
    }
    if (doctroAppoitnmentorder !== undefined) {
      updateData.DoctroAppoitnmentorder = doctroAppoitnmentorder || null;
    }
    if (doctroApporvalcopy !== undefined) {
      updateData.DoctroApporvalcopy = doctroApporvalcopy || null;
    }
    if (headHRSign !== undefined) {
      updateData.HeadHRSign = headHRSign || null;
    }
    if (headHRSeal !== undefined) {
      updateData.HeadHRSeal = headHRSeal || null;
    }

    await table.updateRow(updateData);
    const updated = await table.getRow(ROWID);
    res.status(200).send({ status: 'success', data: { company: updated } });
  } catch (err) {
    res.status(400).send({ status: 'failure', message: err.message || 'Failed to update company.' });
  }
});

// Delete Company
app.delete('/company/:ROWID', async (req, res) => {
  try {
    const { ROWID } = req.params;
    const { catalyst } = res.locals;
    const table = catalyst.datastore().table('Company');
    await table.deleteRow(ROWID);
    res.status(200).send({ status: 'success', data: { id: ROWID } });
  } catch (err) {
    res.status(500).send({ status: 'failure', message: err.message || 'Failed to delete company.' });
  }
});

// Upload file for new company (before creation)
app.post('/company/upload/:docType', async (req, res) => {
  try {
    const { docType } = req.params;
    console.log('New company file upload request for docType:', docType);
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
      console.log('File upload response (new company):', uploadResp);

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
        console.error('Unexpected upload response (new company):', uploadResp);
        return res.status(500).json({ status: 'failure', message: 'File upload failed or invalid response from Catalyst.' });
      }

      // Debug logging for fileId
      console.log('FileId to be saved (new company):', fileId);

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
    console.error('Upload error (new company):', err);
    res.status(500).json({ status: 'failure', message: err.message || 'File upload failed.' });
  }
});

// Upload file for existing company and docType
app.post('/company/:id/upload/:docType', async (req, res) => {
  try {
    const { id, docType } = req.params;
    console.log('Existing company file upload request for id:', id, 'docType:', docType);
    const { catalyst } = res.locals;
    const folderId = DOC_TYPE_TO_FOLDER_ID[docType];
    console.log('Folder ID for docType:', docType, 'is:', folderId);
    
    if (!folderId) {
      console.error('Invalid document type:', docType);
      return res.status(400).send({ status: 'failure', message: 'Invalid document type.' });
    }
    
    if (!req.files || !req.files.file) {
      console.error('No file uploaded for company id:', id, 'docType:', docType);
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

      // Update Company row - store file ID in appropriate column
      const table = catalyst.datastore().table('Company');
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

// Download file for company and docType
app.get('/company/:id/file/:docType', async (req, res) => {
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
    
    const table = catalyst.datastore().table('Company');
    const company = await table.getRow(id);
    const fileId = pickFileIdFromCompanyRow(company, docType);
    let fileName = `${docType}_file`;
    
    console.log('File info:', { fileId, fileName });
    
    if (!fileId) {
      console.error('No file ID found for company:', id, 'docType:', docType);
      return res.status(404).send({ status: 'failure', message: 'File not found for this company.' });
    }

    const resolvedName = await resolveUploadedFileName(catalyst, docType, fileId);
    if (resolvedName) fileName = resolvedName;
    
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

// Get file metadata (name + id) for company and docType
app.get('/company/:id/file-meta/:docType', async (req, res) => {
  try {
    const { id, docType } = req.params;
    const { catalyst } = res.locals;
    const folderId = DOC_TYPE_TO_FOLDER_ID[docType];
    if (!folderId) {
      return res.status(400).send({ status: 'failure', message: 'Invalid document type.' });
    }

    const table = catalyst.datastore().table('Company');
    const company = await table.getRow(id);
    const fileId = pickFileIdFromCompanyRow(company, docType);
    if (!fileId) {
      return res.status(200).send({ status: 'success', data: { fileId: null, fileName: '' } });
    }

    const fileName = await resolveUploadedFileName(catalyst, docType, fileId);
    res.status(200).send({ status: 'success', data: { fileId, fileName } });
  } catch (err) {
    res.status(500).send({ status: 'failure', message: err.message || 'Failed to fetch file metadata.' });
  }
});

// Delete file for company and docType
app.delete('/company/:id/file/:docType', async (req, res) => {
  try {
    const { id, docType } = req.params;
    const { catalyst } = res.locals;
    const folderId = DOC_TYPE_TO_FOLDER_ID[docType];
    
    if (!folderId) {
      return res.status(400).send({ status: 'failure', message: 'Invalid document type.' });
    }
    
    const table = catalyst.datastore().table('Company');
    const company = await table.getRow(id);
    const fileId = pickFileIdFromCompanyRow(company, docType);
    
    if (!fileId) {
      return res.status(404).send({ status: 'failure', message: 'File not found for this company.' });
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
