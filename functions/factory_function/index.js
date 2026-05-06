'use strict';

const express = require('express');
const catalystSDK = require('zcatalyst-sdk-node');
const fileUpload = require('express-fileupload');
const fs = require('fs');
const path = require('path');
const os = require('os');

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

// Table and folder configuration
const TABLE_NAME = 'Factory';
const CHECKLIST_TABLE_NAME = 'Checklist';
const FILE_STORE_FOLDER_ID = '28341000000071362'; // FactoryPDF folder ID
const CHECKLIST_FORM_FOLDER_ID = '28341000000040000'; // Checklist Form folder ID

// Factory data structure based on the Data Store schema
const createFactoryRecord = (data) => {
  return {
    Sector: data.sector || '',
    State: data.state || '',
    FormName: data.formName || ''
  };
};

// Convert Data Store record back to application format
const convertToAppFormat = (record) => {
  return {
    id: record.ROWID,
    sector: record.Sector || '',
    state: record.State || '',
    formName: record.FormName || '',
    formPdfFile: record.FormPdfFile || '',
    createdTime: record.CREATEDTIME,
    modifiedTime: record.MODIFIEDTIME,
    creatorId: record.CREATORID
  };
};

const FACTORY_SYNC_DESC_PREFIX = 'Auto-synced from Factory';
const getFactorySyncDescription = (rowId) => `${FACTORY_SYNC_DESC_PREFIX} (Row ID: ${rowId})`;
const isFactorySyncedDescription = (description = '') => String(description).startsWith(FACTORY_SYNC_DESC_PREFIX);

const buildChecklistUpdatePayload = (row, overrides = {}) => ({
  ROWID: row.ROWID,
  FormName: overrides.FormName ?? row.FormName ?? '',
  Act: overrides.Act ?? row.Act ?? null,
  Consultgovtdep: overrides.Consultgovtdep ?? row.Consultgovtdep ?? null,
  DueDate: overrides.DueDate ?? row.DueDate ?? null,
  Description: overrides.Description ?? row.Description ?? null,
  Sector: overrides.Sector ?? row.Sector ?? null,
  State: overrides.State ?? row.State ?? null,
  FormFile: overrides.FormFile ?? row.FormFile ?? null,
  FormFileName: overrides.FormFileName ?? row.FormFileName ?? null,
  ProofSubmissionFile: overrides.ProofSubmissionFile ?? row.ProofSubmissionFile ?? null,
  ProofSubmissionFileName: overrides.ProofSubmissionFileName ?? row.ProofSubmissionFileName ?? null,
  ComplianceStatus: overrides.ComplianceStatus ?? row.ComplianceStatus ?? null,
  ApprovalStatus: overrides.ApprovalStatus ?? row.ApprovalStatus ?? null,
  Marks: overrides.Marks ?? row.Marks ?? null
});

const removeFactoryFilesFromChecklist = async (catalyst, factoryRows) => {
  if (!factoryRows || factoryRows.length === 0) {
    return 0;
  }

  const nameSet = new Set(
    factoryRows
      .map(row => (row.FormName ? row.FormName.toLowerCase().trim() : ''))
      .filter(Boolean)
  );

  if (nameSet.size === 0) {
    return 0;
  }

  const dataStore = catalyst.datastore();
  const checklistTable = dataStore.table(CHECKLIST_TABLE_NAME);
  const checklistRows = await checklistTable.getAllRows();
  const filestore = catalyst.filestore();
  const checklistFolder = filestore.folder(CHECKLIST_FORM_FOLDER_ID);

  let removedCount = 0;
  for (const row of checklistRows) {
    if (!row.FormName || !row.FormFile) continue;
    const key = row.FormName.toLowerCase().trim();
    if (!nameSet.has(key)) continue;
    if (!isFactorySyncedDescription(row.Description)) continue;

    try {
      await checklistFolder.deleteFile(row.FormFile);
    } catch (err) {
      console.warn(`Failed to delete checklist file ${row.FormFile}:`, err.message || err);
    }

    const updatePayload = buildChecklistUpdatePayload(row, {
      FormFile: null,
      FormFileName: null,
      Description: `${FACTORY_SYNC_DESC_PREFIX} - File removed because Factory data was cleared`
    });

    await checklistTable.updateRow(updatePayload);
    removedCount += 1;
  }

  return removedCount;
};

// Helper to safely extract uploaded file ID from Catalyst response
const getUploadedFileId = (uploadResponse) => {
  if (!uploadResponse) {
    throw new Error('Empty response received from file upload');
  }

  if (Array.isArray(uploadResponse)) {
    if (uploadResponse[0]?.id) {
      return uploadResponse[0].id;
    }
    if (uploadResponse[0]?.file_details?.[0]?.id) {
      return uploadResponse[0].file_details[0].id;
    }
  }

  if (uploadResponse.id) {
    return uploadResponse.id;
  }

  if (uploadResponse.file_details && uploadResponse.file_details[0]?.id) {
    return uploadResponse.file_details[0].id;
  }

  throw new Error('Unable to determine uploaded file ID from response');
};

// Helper to sync a factory file into the Checklist module
const syncFactoryFileToChecklist = async (catalyst, rowId) => {
  const dataStore = catalyst.datastore();
  const factoryTable = dataStore.table(TABLE_NAME);
  const checklistTable = dataStore.table(CHECKLIST_TABLE_NAME);

  const factoryRow = await factoryTable.getRow(rowId);
  if (!factoryRow) {
    throw new Error(`Factory record with ID ${rowId} not found`);
  }

  const formName = factoryRow.FormName;
  if (!formName || !String(formName).trim()) {
    throw new Error('Factory record is missing Form Name, cannot sync');
  }

  const factoryFileId = factoryRow.FormPdfFile;
  if (!factoryFileId) {
    throw new Error('Factory record does not have an uploaded file to sync');
  }

  const filestore = catalyst.filestore();
  const factoryFolder = filestore.folder(FILE_STORE_FOLDER_ID);
  const checklistFolder = filestore.folder(CHECKLIST_FORM_FOLDER_ID);

  let originalFileName = `${formName}.pdf`;
  try {
    const fileDetails = await factoryFolder.getFileDetails(factoryFileId);
    originalFileName = fileDetails?.file_name || fileDetails?.name || originalFileName;
  } catch (metaErr) {
    console.warn(`Unable to fetch file metadata for factory file ${factoryFileId}:`, metaErr.message || metaErr);
  }

  const fileBuffer = await factoryFolder.downloadFile(factoryFileId);
  const tempDir = os.tmpdir();
  const safeFileName = (originalFileName && originalFileName.replace(/[^\w.\-]/g, '_')) || `factory_${rowId}.pdf`;
  const tempPath = path.join(tempDir, `factory_sync_${rowId}_${Date.now()}_${safeFileName}`);

  fs.writeFileSync(tempPath, fileBuffer);

  try {
    const checklistUploadResp = await checklistFolder.uploadFile({
      code: fs.createReadStream(tempPath),
      name: safeFileName
    });

    const checklistFileId = getUploadedFileId(checklistUploadResp);
    const factoryDescription = getFactorySyncDescription(rowId);

    const checklistRows = await checklistTable.getAllRows();
    const matchingRow = checklistRows.find(row =>
      row.FormName && row.FormName.toLowerCase() === formName.toLowerCase()
    );

    if (matchingRow) {
      const shouldReplaceDescription = isFactorySyncedDescription(matchingRow.Description) || !matchingRow.Description;
      const descriptionValue = shouldReplaceDescription ? factoryDescription : matchingRow.Description;

      const updatePayload = {
        ROWID: matchingRow.ROWID,
        FormName: matchingRow.FormName || formName,
        Act: matchingRow.Act || factoryRow.Sector || null,
        Consultgovtdep: matchingRow.Consultgovtdep || null,
        DueDate: matchingRow.DueDate || null,
        Description: descriptionValue,
        Sector: matchingRow.Sector || factoryRow.Sector || null,
        State: matchingRow.State || factoryRow.State || null,
        FormFile: checklistFileId,
        FormFileName: originalFileName,
        ProofSubmissionFile: matchingRow.ProofSubmissionFile || null,
        ProofSubmissionFileName: matchingRow.ProofSubmissionFileName || null,
        ComplianceStatus: matchingRow.ComplianceStatus || null,
        ApprovalStatus: matchingRow.ApprovalStatus || null,
        Marks: matchingRow.Marks || null
      };

      await checklistTable.updateRow(updatePayload);

      return {
        status: 'success',
        action: 'updated',
        checklistRowId: matchingRow.ROWID,
        fileId: checklistFileId
      };
    }

    const insertPayload = {
      FormName: formName,
      Act: factoryRow.Sector || null,
      Consultgovtdep: null,
      DueDate: null,
      Description: factoryDescription,
      Sector: factoryRow.Sector || null,
      State: factoryRow.State || null,
      FormFile: checklistFileId,
      FormFileName: originalFileName,
      ProofSubmissionFile: null,
      ProofSubmissionFileName: null,
      ComplianceStatus: null,
      ApprovalStatus: null,
      Marks: null
    };

    const insertResp = await checklistTable.insertRow(insertPayload);

    return {
      status: 'success',
      action: 'created',
      checklistRowId: insertResp.ROWID,
      fileId: checklistFileId
    };
  } finally {
    if (fs.existsSync(tempPath)) {
      fs.unlinkSync(tempPath);
    }
  }
};

// Upload file to File Store and update Factory record
const uploadFileToFactory = async (catalyst, file, rowId) => {
  try {
    console.log(`Uploading file for row ${rowId}:`, file.name);
    
    // Generate unique filename
    const timestamp = Date.now();
    const fileExtension = file.name.split('.').pop();
    const fileName = `factory_${rowId}_${timestamp}.${fileExtension}`;
    
    console.log('Uploading file to File Store:', fileName);
    console.log('File size:', file.size);
    console.log('File mimetype:', file.mimetype);
    console.log('Folder ID:', FILE_STORE_FOLDER_ID);

    // Create temporary file path
    const tempDir = os.tmpdir();
    const tempPath = path.join(tempDir, file.name);
    
    try {
      // Move uploaded file to temp location
      await file.mv(tempPath);
      console.log('File moved to temp path:', tempPath);
      
      // Upload file to File Store using the same pattern as CompanyDetails
      const filestore = catalyst.filestore();
      const folder = filestore.folder(FILE_STORE_FOLDER_ID);
      
      const uploadResult = await folder.uploadFile({
        code: fs.createReadStream(tempPath),
        name: fileName
      });
      
      console.log('File uploaded successfully:', uploadResult);
      console.log('Upload result ID:', uploadResult.id);
      console.log('Upload result keys:', Object.keys(uploadResult));
      
      // Update Factory record with file ID
      const dataStore = catalyst.datastore();
      const table = dataStore.table(TABLE_NAME);
      
      const updateData = {
        ROWID: rowId,
        FormPdfFile: uploadResult.id.toString()
      };
      console.log('Update data:', updateData);
      
      console.log('About to update row with ID:', rowId);
      console.log('Row ID type:', typeof rowId);
      
      const updateResult = await table.updateRow(updateData);
      
      console.log('Factory record updated with file ID:', updateResult);
      
      return {
        status: 'success',
        message: 'File uploaded and record updated successfully',
        fileId: uploadResult.id,
        fileName: fileName,
        originalName: file.name
      };
    } finally {
      // Clean up temporary file
      if (fs.existsSync(tempPath)) {
        fs.unlinkSync(tempPath);
        console.log('Temporary file cleaned up:', tempPath);
      }
    }
  } catch (error) {
    console.error('Error uploading file:', error);
    console.error('Error details:', error.message);
    console.error('Error code:', error.code);
    console.error('Error statusCode:', error.statusCode);
    return {
      status: 'error',
      message: 'Failed to upload file',
      error: error.message,
      code: error.code,
      statusCode: error.statusCode
    };
  }
};

// Get all factory data
const getAllFactoryData = async (catalyst) => {
  try {
    console.log('Fetching all factory data from Data Store...');
    
    const dataStore = catalyst.datastore();
    const table = dataStore.table(TABLE_NAME);
    
    // Use getAllRows method instead of select
    const result = await table.getAllRows();
    
    console.log(`Found ${result.length} factory records`);
    
    const formattedData = result.map(record => convertToAppFormat(record));
    
    return {
      status: 'success',
      message: 'Factory data retrieved successfully',
      data: formattedData,
      count: formattedData.length
    };
  } catch (error) {
    console.error('Error fetching factory data:', error);
    return {
      status: 'error',
      message: 'Failed to fetch factory data',
      error: error.message
    };
  }
};

// Get factory data by ID
const getFactoryById = async (catalyst, id) => {
  try {
    console.log(`Fetching factory data for ID: ${id}`);
    
    const dataStore = catalyst.datastore();
    const table = dataStore.table(TABLE_NAME);
    
    // Use getRow method instead of select
    const result = await table.getRow(id);
    
    if (!result) {
      return {
        status: 'error',
        message: 'Factory record not found'
      };
    }
    
    const formattedData = convertToAppFormat(result);
    
    return {
      status: 'success',
      message: 'Factory data retrieved successfully',
      data: formattedData
    };
  } catch (error) {
    console.error('Error fetching factory data by ID:', error);
    return {
      status: 'error',
      message: 'Failed to fetch factory data',
      error: error.message
    };
  }
};

// Add new factory data
const addFactoryData = async (catalyst, data) => {
  try {
    console.log('Adding new factory data:', data);
    console.log('Table name being used:', TABLE_NAME);
    
    const dataStore = catalyst.datastore();
    const table = dataStore.table(TABLE_NAME);
    const record = createFactoryRecord(data);
    
    console.log('Record to be inserted:', record);
    
    const result = await table.insertRow(record);
    
    console.log('Factory data added successfully with ROWID:', result.ROWID);
    console.log('Full result object:', result);
    
    return {
      status: 'success',
      message: 'Factory data added successfully',
      data: {
        id: result.ROWID,
        ...data
      }
    };
  } catch (error) {
    console.error('Error adding factory data:', error);
    console.error('Error details:', error.message);
    console.error('Error stack:', error.stack);
    return {
      status: 'error',
      message: 'Failed to add factory data',
      error: error.message
    };
  }
};

// Update factory data
const updateFactoryData = async (catalyst, id, data) => {
  try {
    console.log(`Updating factory data for ID: ${id}`, data);
    
    const dataStore = catalyst.datastore();
    const table = dataStore.table(TABLE_NAME);
    const record = createFactoryRecord(data);
    
    const result = await table.updateRow(id, record);
    
    console.log('Factory data updated successfully');
    
    return {
      status: 'success',
      message: 'Factory data updated successfully',
      data: {
        id: id,
        ...data
      }
    };
  } catch (error) {
    console.error('Error updating factory data:', error);
    return {
      status: 'error',
      message: 'Failed to update factory data',
      error: error.message
    };
  }
};

// Delete factory data
const deleteFactoryData = async (catalyst, id) => {
  try {
    console.log(`Deleting factory data for ID: ${id}`);
    
    const dataStore = catalyst.datastore();
    const table = dataStore.table(TABLE_NAME);
    await table.deleteRow(id);
    
    console.log('Factory data deleted successfully');
    
    return {
      status: 'success',
      message: 'Factory data deleted successfully'
    };
  } catch (error) {
    console.error('Error deleting factory data:', error);
    return {
      status: 'error',
      message: 'Failed to delete factory data',
      error: error.message
    };
  }
};

// Bulk import factory data
const bulkImportFactoryData = async (catalyst, dataArray) => {
  try {
    console.log(`Bulk importing ${dataArray.length} factory records`);
    console.log('Table name being used:', TABLE_NAME);
    
    const dataStore = catalyst.datastore();
    const table = dataStore.table(TABLE_NAME);
    const records = dataArray.map(data => createFactoryRecord(data));
    
    console.log('Records to be inserted:', records.slice(0, 2)); // Log first 2 records as sample
    
    const result = await table.insertRows(records);
    
    console.log(`Bulk import completed. ${result.length} records added`);
    console.log('Result details:', result.slice(0, 2)); // Log first 2 results as sample
    
    return {
      status: 'success',
      message: `Bulk import completed successfully. ${result.length} records added`,
      data: result.map((record, index) => ({
        id: record.ROWID,
        ...dataArray[index]
      }))
    };
  } catch (error) {
    console.error('Error in bulk import:', error);
    console.error('Error details:', error.message);
    console.error('Error stack:', error.stack);
    return {
      status: 'error',
      message: 'Failed to bulk import factory data',
      error: error.message
    };
  }
};

// Get factory data by sector
const getFactoryDataBySector = async (catalyst, sector) => {
  try {
    console.log(`Fetching factory data for sector: ${sector}`);
    
    const dataStore = catalyst.datastore();
    const table = dataStore.table(TABLE_NAME);
    
    // Get all rows and filter by sector
    const allRows = await table.getAllRows();
    const filteredRows = allRows.filter(record => 
      record.Sector && record.Sector.toLowerCase() === sector.toLowerCase()
    );
    
    console.log(`Found ${filteredRows.length} records for sector: ${sector}`);
    
    const formattedData = filteredRows.map(record => convertToAppFormat(record));
    
    return {
      status: 'success',
      message: `Factory data for sector '${sector}' retrieved successfully`,
      data: formattedData,
      count: formattedData.length
    };
  } catch (error) {
    console.error('Error fetching factory data by sector:', error);
    return {
      status: 'error',
      message: 'Failed to fetch factory data by sector',
      error: error.message
    };
  }
};

// Get factory data by state
const getFactoryDataByState = async (catalyst, state) => {
  try {
    console.log(`Fetching factory data for state: ${state}`);
    
    const dataStore = catalyst.datastore();
    const table = dataStore.table(TABLE_NAME);
    
    // Get all rows and filter by state
    const allRows = await table.getAllRows();
    const filteredRows = allRows.filter(record => 
      record.State && record.State.toLowerCase().includes(state.toLowerCase())
    );
    
    console.log(`Found ${filteredRows.length} records for state: ${state}`);
    
    const formattedData = filteredRows.map(record => convertToAppFormat(record));
    
    return {
      status: 'success',
      message: `Factory data for state '${state}' retrieved successfully`,
      data: formattedData,
      count: formattedData.length
    };
  } catch (error) {
    console.error('Error fetching factory data by state:', error);
    return {
      status: 'error',
      message: 'Failed to fetch factory data by state',
      error: error.message
    };
  }
};

// Get count of factory records
const getFactoryCount = async (catalyst) => {
  try {
    console.log('Getting factory data count...');
    
    const dataStore = catalyst.datastore();
    const table = dataStore.table(TABLE_NAME);
    
    // Use getAllRows and get length
    const result = await table.getAllRows();
    
    return {
      status: 'success',
      message: 'Factory data count retrieved successfully',
      count: result.length
    };
  } catch (error) {
    console.error('Error getting factory count:', error);
    return {
      status: 'error',
      message: 'Failed to get factory data count',
      error: error.message
    };
  }
};

// Bulk delete all factory data
const bulkDeleteFactoryData = async (catalyst) => {
  try {
    console.log('Bulk deleting all factory data...');
    
    const dataStore = catalyst.datastore();
    const table = dataStore.table(TABLE_NAME);
    
    // Get all rows first to get their IDs
    const allRows = await table.getAllRows();
    console.log(`Found ${allRows.length} records to delete`);
    
    if (allRows.length === 0) {
      return {
        status: 'success',
        message: 'No records found to delete',
        deletedCount: 0
      };
    }

    try {
      const removedCount = await removeFactoryFilesFromChecklist(catalyst, allRows);
      console.log(`Removed ${removedCount} factory-synced checklist file(s) before deleting factory data`);
    } catch (cleanupErr) {
      console.error('Failed to remove factory files from checklist before deletion:', cleanupErr);
    }
    
    // Extract ROWIDs for deletion
    const rowIds = allRows.map(record => record.ROWID);
    console.log('Row IDs to delete:', rowIds);
    
    // Delete all rows
    const deletePromises = rowIds.map(id => table.deleteRow(id));
    await Promise.all(deletePromises);
    
    console.log(`Successfully deleted ${rowIds.length} factory records`);
    
    return {
      status: 'success',
      message: `Successfully deleted ${rowIds.length} factory records`,
      deletedCount: rowIds.length
    };
  } catch (error) {
    console.error('Error in bulk delete:', error);
    console.error('Error details:', error.message);
    console.error('Error stack:', error.stack);
    return {
      status: 'error',
      message: 'Failed to bulk delete factory data',
      error: error.message
    };
  }
};

// Populate table with sample data
const populateTableWithSampleData = async (catalyst) => {
  try {
    console.log('Populating Factory table with sample data...');
    
    const sampleData = [
      {
        Sector: "Manufacturing",
        State: "Maharashtra",
        FormName: "Form 1 - Notice of Occupation"
      },
      {
        Sector: "Manufacturing",
        State: "Karnataka",
        FormName: "Form 2 - Notice of Change"
      },
      {
        Sector: "Manufacturing",
        State: "Tamil Nadu",
        FormName: "Form 3 - Notice of Change of Manager"
      },
      {
        Sector: "Healthcare",
        State: "Maharashtra",
        FormName: "Form 4 - Notice of Change of Name"
      },
      {
        Sector: "Healthcare",
        State: "Karnataka",
        FormName: "Form 5 - Notice of Change of Occupier"
      },
      {
        Sector: "Finance",
        State: "Tamil Nadu",
        FormName: "Form 6 - Notice of Change of Address"
      },
      {
        Sector: "Environment",
        State: "Maharashtra",
        FormName: "Form 7 - Notice of Change of Nature of Work"
      },
      {
        Sector: "Manufacturing",
        State: "Gujarat",
        FormName: "Form 8 - Notice of Change of Working Hours"
      }
    ];
    
    const dataStore = catalyst.datastore();
    const table = dataStore.table(TABLE_NAME);
    const result = await table.insertRows(sampleData);
    
    console.log(`Successfully populated table with ${result.length} sample records`);
    
    return {
      status: 'success',
      message: `Successfully populated Factory table with ${result.length} sample records`,
      data: result.map(record => ({
        id: record.ROWID,
        sector: record.Sector,
        state: record.State,
        formName: record.FormName
      }))
    };
  } catch (error) {
    console.error('Error populating table with sample data:', error);
    return {
      status: 'error',
      message: 'Failed to populate table with sample data',
      error: error.message
    };
  }
};

// Health check
app.get('/', (req, res) => {
  res.status(200).json({ status: 'success', message: 'factory_function ready' });
});

// Test endpoint to check table access and verify data
app.get('/test-table', async (req, res) => {
  try {
    console.log('Testing Factory table access...');
    const catalyst = res.locals.catalyst;
    const dataStore = catalyst.datastore();
    const table = dataStore.table(TABLE_NAME);
    
    console.log('Table name:', TABLE_NAME);
    console.log('Table object:', table);
    
    // Try to get all rows to see if table exists and has data
    const allRows = await table.getAllRows();
    console.log('Total rows in table:', allRows.length);
    console.log('Sample rows:', allRows.slice(0, 3));
    
    // Also try ZCQL query as alternative
    const zcql = catalyst.zcql();
    const countQuery = await zcql.executeZCQLQuery(`SELECT COUNT(ROWID) as count FROM ${TABLE_NAME}`);
    console.log('ZCQL count result:', countQuery);
    
    res.status(200).json({ 
      status: 'success', 
      message: 'Table access successful',
      tableInfo: {
        tableName: TABLE_NAME,
        recordCount: allRows.length,
        zcqlCount: countQuery[0]?.[TABLE_NAME]?.count || 0,
        sampleData: allRows.slice(0, 3)
      }
    });
  } catch (err) {
    console.error('Table test error:', err);
    res.status(500).json({ 
      status: 'failure', 
      message: err.message || 'Failed to access table',
      error: err.toString(),
      tableName: TABLE_NAME
    });
  }
});

// Test endpoint to add a single record for debugging
app.post('/test-add', async (req, res) => {
  try {
    console.log('Test add endpoint called');
    const catalyst = res.locals.catalyst;
    
    const testData = {
      sector: "Test Sector",
      state: "Test State",
      formName: "Test Form"
    };
    
    console.log('Test data to be added:', testData);
    
    const result = await addFactoryData(catalyst, testData);
    
    res.status(result.status === 'success' ? 200 : 400).json(result);
    
  } catch (error) {
    console.error('Test add error:', error);
    res.status(500).json({
      status: 'error',
      message: 'Test add failed',
      error: error.message
    });
  }
});

// Test endpoint to check File Store access
app.get('/test-filestore', async (req, res) => {
  try {
    console.log('Testing File Store access...');
    const catalyst = res.locals.catalyst;
    const filestore = catalyst.filestore();
    const folder = filestore.folder(FILE_STORE_FOLDER_ID);
    
    console.log('File Store folder ID:', FILE_STORE_FOLDER_ID);
    
    // Try to get folder info
    const folderInfo = await folder.getFolderDetails();
    console.log('Folder info:', folderInfo);
    
    res.status(200).json({ 
      status: 'success', 
      message: 'File Store access successful',
      folderInfo: folderInfo,
      folderId: FILE_STORE_FOLDER_ID
    });
  } catch (err) {
    console.error('File Store test error:', err);
    res.status(500).json({ 
      status: 'failure', 
      message: err.message || 'Failed to access File Store',
      error: err.toString(),
      folderId: FILE_STORE_FOLDER_ID
    });
  }
});

// File upload endpoint
app.post('/uploadFile', async (req, res) => {
  try {
    console.log('File upload endpoint called');
    const catalyst = res.locals.catalyst;
    const { rowId } = req.body;
    
    if (!req.files || !req.files.file) {
      return res.status(400).json({
        status: 'error',
        message: 'No file provided'
      });
    }
    
    if (!rowId) {
      return res.status(400).json({
        status: 'error',
        message: 'Row ID is required'
      });
    }
    
    console.log('File upload request:', {
      fileName: req.files.file.name,
      fileSize: req.files.file.size,
      rowId: rowId
    });
    
    const result = await uploadFileToFactory(catalyst, req.files.file, rowId);
    
    res.status(result.status === 'success' ? 200 : 400).json(result);
    
  } catch (error) {
    console.error('File upload error:', error);
    res.status(500).json({
      status: 'error',
      message: 'File upload failed',
      error: error.message
    });
  }
});

// Sync a factory file to Checklist
app.post('/syncToChecklist', async (req, res) => {
  try {
    const { rowId } = req.body;
    if (!rowId) {
      return res.status(400).json({
        status: 'error',
        message: 'rowId is required to sync with Checklist'
      });
    }

    const catalyst = res.locals.catalyst;
    const result = await syncFactoryFileToChecklist(catalyst, rowId);

    res.status(200).json(result);
  } catch (error) {
    console.error('Error syncing factory file to checklist:', error);
    res.status(500).json({
      status: 'error',
      message: error.message || 'Failed to sync factory file to checklist'
    });
  }
});

// Get all factory data
app.get('/factory', async (req, res) => {
  try {
    const catalyst = res.locals.catalyst;
    const { action, id, sector, state } = req.query;
    
    console.log(`Factory API Request - Action: ${action}, ID: ${id}, Sector: ${sector}, State: ${state}`);
    
    let result;
    
    switch (action) {
      case 'getAll':
        result = await getAllFactoryData(catalyst);
        break;
        
      case 'getById':
        if (!id) {
          result = {
            status: 'error',
            message: 'ID parameter is required for getById action'
          };
        } else {
          result = await getFactoryById(catalyst, id);
        }
        break;
        
      case 'getBySector':
        if (!sector) {
          result = {
            status: 'error',
            message: 'Sector parameter is required for getBySector action'
          };
        } else {
          result = await getFactoryDataBySector(catalyst, sector);
        }
        break;
        
      case 'getByState':
        if (!state) {
          result = {
            status: 'error',
            message: 'State parameter is required for getByState action'
          };
        } else {
          result = await getFactoryDataByState(catalyst, state);
        }
        break;
        
      case 'count':
        result = await getFactoryCount(catalyst);
        break;
        
      case 'populate':
        result = await populateTableWithSampleData(catalyst);
        break;
        
      default:
        result = {
          status: 'error',
          message: 'Invalid action. Supported actions: getAll, getById, getBySector, getByState, count, populate'
        };
    }
    
    res.status(result.status === 'success' ? 200 : 400).json(result);
    
  } catch (error) {
    console.error('Factory API Error:', error);
    res.status(500).json({
      status: 'error',
      message: 'Internal server error',
      error: error.message
    });
  }
});

// POST routes for data modification
app.post('/factory', async (req, res) => {
  try {
    const catalyst = res.locals.catalyst;
    const { action } = req.query;
    const body = req.body;
    
    console.log(`Factory API POST Request - Action: ${action}`);
    
    let result;
    
    switch (action) {
      case 'add':
        if (!body) {
          result = {
            status: 'error',
            message: 'Request body is required for add action'
          };
        } else {
          result = await addFactoryData(catalyst, body);
        }
        break;
        
      case 'bulkImport':
        if (!body || !Array.isArray(body)) {
          result = {
            status: 'error',
            message: 'Request body must be an array for bulkImport action'
          };
        } else {
          result = await bulkImportFactoryData(catalyst, body);
        }
        break;
        
      case 'bulkDelete':
        result = await bulkDeleteFactoryData(catalyst);
        break;
        
      default:
        result = {
          status: 'error',
          message: 'Invalid action. Supported actions: add, bulkImport, bulkDelete'
        };
    }
    
    res.status(result.status === 'success' ? 200 : 400).json(result);
    
  } catch (error) {
    console.error('Factory API Error:', error);
    res.status(500).json({
      status: 'error',
      message: 'Internal server error',
      error: error.message
    });
  }
});

// PUT route for updates
app.put('/factory', async (req, res) => {
  try {
    const catalyst = res.locals.catalyst;
    const { action, id } = req.query;
    const body = req.body;
    
    console.log(`Factory API PUT Request - Action: ${action}, ID: ${id}`);
    
    let result;
    
    switch (action) {
      case 'update':
        if (!id || !body) {
          result = {
            status: 'error',
            message: 'ID and request body are required for update action'
          };
        } else {
          result = await updateFactoryData(catalyst, id, body);
        }
        break;
        
      default:
        result = {
          status: 'error',
          message: 'Invalid action. Supported actions: update'
        };
    }
    
    res.status(result.status === 'success' ? 200 : 400).json(result);
    
  } catch (error) {
    console.error('Factory API Error:', error);
    res.status(500).json({
      status: 'error',
      message: 'Internal server error',
      error: error.message
    });
  }
});

// DELETE route for deletions
app.delete('/factory', async (req, res) => {
  try {
    const catalyst = res.locals.catalyst;
    const { action, id } = req.query;
    
    console.log(`Factory API DELETE Request - Action: ${action}, ID: ${id}`);
    
    let result;
    
    switch (action) {
      case 'delete':
        if (!id) {
          result = {
            status: 'error',
            message: 'ID parameter is required for delete action'
          };
        } else {
          result = await deleteFactoryData(catalyst, id);
        }
        break;
        
      default:
        result = {
          status: 'error',
          message: 'Invalid action. Supported actions: delete'
        };
    }
    
    res.status(result.status === 'success' ? 200 : 400).json(result);
    
  } catch (error) {
    console.error('Factory API Error:', error);
    res.status(500).json({
      status: 'error',
      message: 'Internal server error',
      error: error.message
    });
  }
});

module.exports = app;
