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
  res.status(200).json({ status: 'success', message: 'statutorymasterfactory_function ready' });
});

// Test endpoint to check table access
app.get('/test-table', async (req, res) => {
  try {
    console.log('Testing table access...');
    const { catalyst } = res.locals;
    const table = catalyst.datastore().table('Statutorymasterfactory');
    
    // Try to get table info
    console.log('Table object:', table);
    
    // Try a simple query
    const zcql = catalyst.zcql();
    const testQuery = await zcql.executeZCQLQuery('SELECT COUNT(ROWID) as count FROM Statutorymasterfactory');
    console.log('Test query result:', testQuery);
    
    res.status(200).send({ 
      status: 'success', 
      message: 'Table access successful',
      tableInfo: {
        tableName: 'Statutorymasterfactory',
        recordCount: testQuery[0]?.Statutorymasterfactory?.count || 0
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
    
    // Try to access the PDFFile folder
    const folder = filestore.folder(STATUTORYMASTERFACTORY_FOLDER_ID);
    console.log('Folder object:', folder);
    
    // Try to list files in the folder
    const files = await folder.getAllFiles();
    console.log('Files in PDFFile folder:', files);
    
    res.status(200).send({ 
      status: 'success', 
      message: 'File Store access successful',
      fileStoreInfo: {
        folderId: STATUTORYMASTERFACTORY_FOLDER_ID,
        folderName: 'PDFFile',
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

// ===== Statutorymasterfactory CRUD =====
// Table: Statutorymasterfactory (columns: Sector, State, FormFile, Act)

// File Store folder ID for Statutorymasterfactory files
const STATUTORYMASTERFACTORY_FOLDER_ID = '28341000000072668'; // PDFFile folder ID from your File Store

// Upload file for Statutorymasterfactory
app.post('/statutorymasterfactory/upload', async (req, res) => {
  try {
    console.log('File upload request received');
    console.log('Request files:', req.files);
    
    const { catalyst } = res.locals;
    
    if (!req.files || !req.files.formFile) {
      console.error('No file uploaded');
      return res.status(400).json({ status: 'failure', message: 'No file uploaded.' });
    }
    
    const file = req.files.formFile;
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
      
      const uploadResp = await catalyst.filestore().folder(STATUTORYMASTERFACTORY_FOLDER_ID).uploadFile({
        code: fs.createReadStream(tempPath),
        name: file.name
      });

      // Clean up temporary file
      if (fs.existsSync(tempPath)) {
        fs.unlinkSync(tempPath);
      }

      console.log('File upload response:', uploadResp);

      // Handle different possible response structures
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

// Create Statutorymasterfactory
app.post('/statutorymasterfactory', async (req, res) => {
  try {
    console.log('Received request body:', JSON.stringify(req.body, null, 2));
    
    const {
      sector,
      state,
      formFile,
      act
    } = req.body;

    // Validate required fields
    if (!sector || !String(sector).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Sector is required.' });
    }
    if (!state || !String(state).trim()) {
      return res.status(400).send({ status: 'failure', message: 'State is required.' });
    }
    if (!act || !String(act).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Act is required.' });
    }

    const { catalyst } = res.locals;
    const table = catalyst.datastore().table('Statutorymasterfactory');
    
    const insertData = {
      Sector: sector,
      State: state,
      FormFile: formFile || null,
      Act: act
    };
    
    console.log('Inserting data to Statutorymasterfactory table:', JSON.stringify(insertData, null, 2));
    
    const insertResp = await table.insertRow(insertData);
    console.log('Insert response:', JSON.stringify(insertResp, null, 2));
    
    const created = await table.getRow(insertResp.ROWID);
    console.log('Created record:', JSON.stringify(created, null, 2));
    
    // Map to frontend expected format
    const mappedRecord = {
      id: created.ROWID,
      sector: created.Sector || '',
      state: created.State || '',
      formFile: created.FormFile || '',
      act: created.Act || '',
      createdTime: created.CREATEDTIME,
      modifiedTime: created.MODIFIEDTIME
    };
    
    res.status(200).send({ status: 'success', data: { statutorymasterfactory: mappedRecord } });
  } catch (err) {
    console.error('Error creating statutorymasterfactory record:', err);
    res.status(400).send({ status: 'failure', message: err.message || 'Failed to create statutorymasterfactory record.' });
  }
});

// Get single Statutorymasterfactory record with all details
app.get('/statutorymasterfactory/:ROWID', async (req, res) => {
  try {
    const { ROWID } = req.params;
    console.log('Fetching statutorymasterfactory record:', ROWID);
    const { catalyst } = res.locals;
    const table = catalyst.datastore().table('Statutorymasterfactory');
    
    const statutorymasterfactory = await table.getRow(ROWID);
    console.log('Retrieved statutorymasterfactory record:', JSON.stringify(statutorymasterfactory, null, 2));
    
    // Map to frontend expected format
    const mappedRecord = {
      id: statutorymasterfactory.ROWID,
      sector: statutorymasterfactory.Sector || '',
      state: statutorymasterfactory.State || '',
      formFile: statutorymasterfactory.FormFile || '',
      act: statutorymasterfactory.Act || '',
      createdTime: statutorymasterfactory.CREATEDTIME,
      modifiedTime: statutorymasterfactory.MODIFIEDTIME
    };
    
    res.status(200).send({ status: 'success', data: { statutorymasterfactory: mappedRecord } });
  } catch (err) {
    console.error('Error fetching statutorymasterfactory record:', err);
    res.status(500).send({ status: 'failure', message: err.message || 'Failed to fetch statutorymasterfactory record.' });
  }
});

// List Statutorymasterfactory with optional pagination
app.get('/statutorymasterfactory', async (req, res) => {
  try {
    console.log('Fetching statutorymasterfactory records...');
    const { catalyst } = res.locals;
    const zcql = catalyst.zcql();
    const page = parseInt(req.query.page, 10) || 1;
    const perPage = parseInt(req.query.perPage, 10) || 50;
    const returnAll = !req.query.page && !req.query.perPage;
    const limitClause = returnAll ? '' : `LIMIT ${(page - 1) * perPage + 1},${perPage}`;
    
    console.log('Executing count query...');
    const countRows = await zcql.executeZCQLQuery('SELECT COUNT(ROWID) as count FROM Statutorymasterfactory');
    const total = parseInt(countRows[0].Statutorymasterfactory.count, 10) || 0;
    console.log('Total records:', total);
    
    console.log('Executing data query...');
    // Use SELECT * to get all columns dynamically
    const rows = await zcql.executeZCQLQuery(`SELECT * FROM Statutorymasterfactory ORDER BY ROWID DESC ${limitClause}`);
    console.log('Raw rows from database:', JSON.stringify(rows, null, 2));
    
    // Process data and map to frontend expected format
    const statutorymasterfactoryData = rows.map(r => {
      const record = {
        id: r.Statutorymasterfactory.ROWID,
        sector: r.Statutorymasterfactory.Sector || '',
        state: r.Statutorymasterfactory.State || '',
        formFile: r.Statutorymasterfactory.FormFile || '',
        act: r.Statutorymasterfactory.Act || '',
        createdTime: r.Statutorymasterfactory.CREATEDTIME,
        modifiedTime: r.Statutorymasterfactory.MODIFIEDTIME
      };
      
      return record;
    });
    
    console.log('Processed statutorymasterfactory data:', JSON.stringify(statutorymasterfactoryData, null, 2));
    res.status(200).send({ status: 'success', data: { statutorymasterfactoryData, total, hasMore: returnAll ? false : page * perPage < total } });
  } catch (err) {
    console.error('Error fetching statutorymasterfactory records:', err);
    res.status(500).send({ status: 'failure', message: err.message || 'Failed to fetch statutorymasterfactory records.' });
  }
});

// Update Statutorymasterfactory
app.put('/statutorymasterfactory/:ROWID', async (req, res) => {
  try {
    const { ROWID } = req.params;
    const {
      sector,
      state,
      formFile,
      act
    } = req.body;

    // Validate required fields
    if (!sector || !String(sector).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Sector is required.' });
    }
    if (!state || !String(state).trim()) {
      return res.status(400).send({ status: 'failure', message: 'State is required.' });
    }
    if (!act || !String(act).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Act is required.' });
    }

    const { catalyst } = res.locals;
    const table = catalyst.datastore().table('Statutorymasterfactory');
    await table.updateRow({
      ROWID,
      Sector: sector,
      State: state,
      FormFile: formFile || null,
      Act: act
    });
    const updated = await table.getRow(ROWID);
    
    // Map to frontend expected format
    const mappedRecord = {
      id: updated.ROWID,
      sector: updated.Sector || '',
      state: updated.State || '',
      formFile: updated.FormFile || '',
      act: updated.Act || '',
      createdTime: updated.CREATEDTIME,
      modifiedTime: updated.MODIFIEDTIME
    };
    
    res.status(200).send({ status: 'success', data: { statutorymasterfactory: mappedRecord } });
  } catch (err) {
    res.status(400).send({ status: 'failure', message: err.message || 'Failed to update statutorymasterfactory record.' });
  }
});

// Delete Statutorymasterfactory
app.delete('/statutorymasterfactory/:ROWID', async (req, res) => {
  try {
    const { ROWID } = req.params;
    const { catalyst } = res.locals;
    const table = catalyst.datastore().table('Statutorymasterfactory');
    await table.deleteRow(ROWID);
    res.status(200).send({ status: 'success', data: { id: ROWID } });
  } catch (err) {
    res.status(500).send({ status: 'failure', message: err.message || 'Failed to delete statutorymasterfactory record.' });
  }
});

// Helper: map Excel field names to datastore column names
function mapFieldNameToColumn(excelFieldName) {
  const fieldMapping = {
    // Statutorymasterfactory fields
    'Sector': 'Sector',
    'State': 'State',
    'Form File': 'FormFile',
    'Act': 'Act',
    
    // Handle variations and common Excel column names
    'FormFile': 'FormFile',
    'Form_File': 'FormFile',
    'Form-File': 'FormFile',
    
    // Handle __EMPTY_ fields as fallback
    '__EMPTY_1': 'Sector',
    '__EMPTY_2': 'State',
    '__EMPTY_3': 'FormFile',
    '__EMPTY_4': 'Act'
  };
  
  // Return mapped column name or original if no mapping found
  return fieldMapping[excelFieldName] || excelFieldName;
}

// Bulk insert endpoint for Excel data
app.post('/statutorymasterfactory/bulk', async (req, res) => {
  try {
    const { records } = req.body;
    
    if (!records || !Array.isArray(records)) {
      return res.status(400).json({ error: 'Records array is required' });
    }

    console.log(`Processing bulk insert for ${records.length} records`);
    const { catalyst } = res.locals;
    const table = catalyst.datastore().table('Statutorymasterfactory');
    const results = [];
    
    // Process each record
    for (const record of records) {
      try {
        // Remove the id field if it exists (we don't want to insert it)
        const { id, ...recordData } = record;
        
        // Clean the data - remove null/undefined values and convert to strings
        const cleanedData = {};
        Object.keys(recordData).forEach(key => {
          const value = recordData[key];
          if (value !== null && value !== undefined && value !== '') {
            // Map Excel field name to datastore column name
            const columnName = mapFieldNameToColumn(key);
            // Convert to string and trim whitespace
            cleanedData[columnName] = String(value).trim();
          }
        });

        // Only insert if we have some data
        if (Object.keys(cleanedData).length > 0) {
          // Use mapped column names that match the datastore schema
          console.log('Inserting data with mapped columns:', cleanedData);
          const insertResult = await table.insertRow(cleanedData);
          results.push({
            success: true,
            rowId: insertResult.ROWID,
            data: cleanedData
          });
        } else {
          results.push({
            success: false,
            error: 'No valid data to insert',
            data: record
          });
        }
        
      } catch (recordError) {
        console.error('Error inserting record:', recordError);
        results.push({
          success: false,
          error: recordError.message,
          data: record
        });
      }
    }

    const successCount = results.filter(r => r.success).length;
    const errorCount = results.filter(r => !r.success).length;

    console.log(`Bulk insert completed: ${successCount} successful, ${errorCount} failed`);

    res.status(200).json({
      message: `Bulk insert completed. ${successCount} successful, ${errorCount} failed.`,
      totalRecords: records.length,
      successCount,
      errorCount,
      results
    });

  } catch (error) {
    console.error('Bulk insert error:', error);
    res.status(500).json({ error: error.message });
  }
});

// Get all records (for frontend compatibility)
app.get('/statutorymasterfactory?action=getAll', async (req, res) => {
  try {
    console.log('Fetching all statutorymasterfactory records...');
    const { catalyst } = res.locals;
    const zcql = catalyst.zcql();
    
    console.log('Executing data query...');
    const rows = await zcql.executeZCQLQuery('SELECT * FROM Statutorymasterfactory ORDER BY ROWID DESC');
    console.log('Raw rows from database:', JSON.stringify(rows, null, 2));
    
    // Process data dynamically - keep all columns as they are
    const statutorymasterfactoryData = rows.map(r => {
      const record = { ROWID: r.Statutorymasterfactory.ROWID };
      
      // Copy all fields except system fields
      Object.keys(r.Statutorymasterfactory).forEach(key => {
        if (key !== 'ROWID' && key !== 'CREATORID' && key !== 'CREATEDTIME' && key !== 'MODIFIEDTIME') {
          record[key] = r.Statutorymasterfactory[key];
        }
      });
      
      return record;
    });
    
    console.log('Processed statutorymasterfactory data:', JSON.stringify(statutorymasterfactoryData, null, 2));
    res.status(200).send({ status: 'success', data: statutorymasterfactoryData });
  } catch (err) {
    console.error('Error fetching statutorymasterfactory records:', err);
    res.status(500).send({ status: 'failure', message: err.message || 'Failed to fetch statutorymasterfactory records.' });
  }
});

// Search/Filter endpoint
app.get('/statutorymasterfactory/search', async (req, res) => {
  try {
    const { sector, state, act, formFile } = req.query;
    console.log('Search parameters:', { sector, state, act, formFile });
    
    const { catalyst } = res.locals;
    const zcql = catalyst.zcql();
    
    // Build dynamic WHERE clause
    const conditions = [];
    const params = [];
    
    if (sector) {
      conditions.push('Sector LIKE ?');
      params.push(`%${sector}%`);
    }
    if (state) {
      conditions.push('State LIKE ?');
      params.push(`%${state}%`);
    }
    if (act) {
      conditions.push('Act LIKE ?');
      params.push(`%${act}%`);
    }
    if (formFile) {
      conditions.push('FormFile LIKE ?');
      params.push(`%${formFile}%`);
    }
    
    const whereClause = conditions.length > 0 ? `WHERE ${conditions.join(' AND ')}` : '';
    const query = `SELECT * FROM Statutorymasterfactory ${whereClause} ORDER BY ROWID DESC`;
    
    console.log('Search query:', query);
    console.log('Search params:', params);
    
    const rows = await zcql.executeZCQLQuery(query, params);
    console.log('Search results:', JSON.stringify(rows, null, 2));
    
    // Process data dynamically
    const statutorymasterfactoryData = rows.map(r => {
      const record = { ROWID: r.Statutorymasterfactory.ROWID };
      
      // Copy all fields except system fields
      Object.keys(r.Statutorymasterfactory).forEach(key => {
        if (key !== 'ROWID' && key !== 'CREATORID' && key !== 'CREATEDTIME' && key !== 'MODIFIEDTIME') {
          record[key] = r.Statutorymasterfactory[key];
        }
      });
      
      return record;
    });
    
    res.status(200).send({ status: 'success', data: statutorymasterfactoryData });
  } catch (err) {
    console.error('Error searching statutorymasterfactory records:', err);
    res.status(500).send({ status: 'failure', message: err.message || 'Failed to search statutorymasterfactory records.' });
  }
});

// Export data endpoint
app.get('/statutorymasterfactory/export', async (req, res) => {
  try {
    console.log('Exporting statutorymasterfactory data...');
    const { catalyst } = res.locals;
    const zcql = catalyst.zcql();
    
    const rows = await zcql.executeZCQLQuery('SELECT * FROM Statutorymasterfactory ORDER BY ROWID DESC');
    
    // Process data for export
    const exportData = rows.map(r => ({
      'Sector': r.Statutorymasterfactory.Sector || '',
      'State': r.Statutorymasterfactory.State || '',
      'Form File': r.Statutorymasterfactory.FormFile || '',
      'Act': r.Statutorymasterfactory.Act || '',
      'Created Time': r.Statutorymasterfactory.CREATEDTIME || '',
      'Modified Time': r.Statutorymasterfactory.MODIFIEDTIME || ''
    }));
    
    res.status(200).json({
      status: 'success',
      data: exportData,
      totalRecords: exportData.length
    });
  } catch (err) {
    console.error('Error exporting statutorymasterfactory data:', err);
    res.status(500).send({ status: 'failure', message: err.message || 'Failed to export statutorymasterfactory data.' });
  }
});

// Clear all data endpoint (use with caution)
app.delete('/statutorymasterfactory/clear', async (req, res) => {
  try {
    console.log('Clearing all statutorymasterfactory data...');
    const { catalyst } = res.locals;
    const zcql = catalyst.zcql();
    
    // Get all records first
    const rows = await zcql.executeZCQLQuery('SELECT ROWID FROM Statutorymasterfactory');
    const table = catalyst.datastore().table('Statutorymasterfactory');
    
    // Delete each record
    for (const row of rows) {
      await table.deleteRow(row.Statutorymasterfactory.ROWID);
    }
    
    res.status(200).send({ 
      status: 'success', 
      message: `Cleared ${rows.length} records from Statutorymasterfactory table.` 
    });
  } catch (err) {
    console.error('Error clearing statutorymasterfactory data:', err);
    res.status(500).send({ status: 'failure', message: err.message || 'Failed to clear statutorymasterfactory data.' });
  }
});

module.exports = app;