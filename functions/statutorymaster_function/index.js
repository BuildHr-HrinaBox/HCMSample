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
  res.status(200).json({ status: 'success', message: 'statutorymaster_function ready' });
});

// Test endpoint to check table access
app.get('/test-table', async (req, res) => {
  try {
    console.log('Testing table access...');
    const { catalyst } = res.locals;
    const table = catalyst.datastore().table('StatutoryMaster');
    
    // Try to get table info
    console.log('Table object:', table);
    
    // Try a simple query
    const zcql = catalyst.zcql();
    const testQuery = await zcql.executeZCQLQuery('SELECT COUNT(ROWID) as count FROM StatutoryMaster');
    console.log('Test query result:', testQuery);
    
    res.status(200).send({ 
      status: 'success', 
      message: 'Table access successful',
      tableInfo: {
        tableName: 'StatutoryMaster',
        recordCount: testQuery[0]?.StatutoryMaster?.count || 0
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

// ===== StatutoryMaster CRUD =====
// Table: StatutoryMaster (columns: NameofWorker, WorkerID, Gender, FatherName, DateofBirth, PresentAddress, PresentCity, PresentState, PresentPostalCode, PresentCountry, PermanentAddress, PermanentCity, PermanentState, PermanentPostalCode, PermanentCountry, AadhaarNo, DateofEntryintoService, Designation, EPFNo, UANNo, ESINo, DateonwhichCompletion, DateonwhichMadePayment, PeriodofSuspension, BankAccountNumber, NameofBank, Branch, MobileNumber, Gmail, DateofExit, ReasonforExit, Remarks)

// Create StatutoryMaster
app.post('/statutorymaster', async (req, res) => {
  try {
    console.log('Received request body:', JSON.stringify(req.body, null, 2));
    
    const {
      nameofWorker,
      workerID,
      gender,
      fatherName,
      dateofBirth,
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
      dateofEntryintoService,
      designation,
      epfNo,
      uanNo,
      esiNo,
      dateonwhichCompletion,
      dateonwhichMadePayment,
      periodofSuspension,
      bankAccountNumber,
      nameofBank,
      branch,
      mobileNumber,
      gmail,
      dateofExit,
      reasonforExit,
      remarks
    } = req.body;

    // Validate required fields
    if (!nameofWorker || !String(nameofWorker).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Name of Worker is required.' });
    }
    if (!workerID || !String(workerID).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Worker ID is required.' });
    }
    if (!fatherName || !String(fatherName).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Father Name is required.' });
    }

    const { catalyst } = res.locals;
    const table = catalyst.datastore().table('StatutoryMaster');
    
    const insertData = {
      NameofWorker: nameofWorker,
      WorkerID: workerID,
      Gender: gender || null,
      FatherName: fatherName,
      DateofBirth: dateofBirth || null,
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
      DateofEntryintoService: dateofEntryintoService || null,
      Designation: designation || null,
      EPFNo: epfNo || null,
      UANNo: uanNo || null,
      ESINo: esiNo || null,
      DateonwhichCompletion: dateonwhichCompletion || null,
      DateonwhichMadePayment: dateonwhichMadePayment || null,
      PeriodofSuspension: periodofSuspension || null,
      BankAccountNumber: bankAccountNumber || null,
      NameofBank: nameofBank || null,
      Branch: branch || null,
      MobileNumber: mobileNumber || null,
      Gmail: gmail || null,
      DateofExit: dateofExit || null,
      ReasonforExit: reasonforExit || null,
      Remarks: remarks || null
    };
    
    console.log('Inserting data to StatutoryMaster table:', JSON.stringify(insertData, null, 2));
    
    const insertResp = await table.insertRow(insertData);
    console.log('Insert response:', JSON.stringify(insertResp, null, 2));
    
    const created = await table.getRow(insertResp.ROWID);
    console.log('Created record:', JSON.stringify(created, null, 2));
    
    res.status(200).send({ status: 'success', data: { statutoryMaster: created } });
  } catch (err) {
    console.error('Error creating statutory master record:', err);
    res.status(400).send({ status: 'failure', message: err.message || 'Failed to create statutory master record.' });
  }
});

// Get single StatutoryMaster record with all details
app.get('/statutorymaster/:ROWID', async (req, res) => {
  try {
    const { ROWID } = req.params;
    console.log('Fetching statutory master record:', ROWID);
    const { catalyst } = res.locals;
    const table = catalyst.datastore().table('StatutoryMaster');
    
    const statutoryMaster = await table.getRow(ROWID);
    console.log('Retrieved statutory master record:', JSON.stringify(statutoryMaster, null, 2));
    
    // Return the record as-is with dynamic column names
    res.status(200).send({ status: 'success', data: { statutoryMaster } });
  } catch (err) {
    console.error('Error fetching statutory master record:', err);
    res.status(500).send({ status: 'failure', message: err.message || 'Failed to fetch statutory master record.' });
  }
});

// List StatutoryMaster with optional pagination
app.get('/statutorymaster', async (req, res) => {
  try {
    console.log('Fetching statutory master records...');
    const { catalyst } = res.locals;
    const zcql = catalyst.zcql();
    const page = parseInt(req.query.page, 10) || 1;
    const perPage = parseInt(req.query.perPage, 10) || 50;
    const returnAll = !req.query.page && !req.query.perPage;
    const limitClause = returnAll ? '' : `LIMIT ${(page - 1) * perPage + 1},${perPage}`;
    
    console.log('Executing count query...');
    const countRows = await zcql.executeZCQLQuery('SELECT COUNT(ROWID) as count FROM StatutoryMaster');
    const total = parseInt(countRows[0].StatutoryMaster.count, 10) || 0;
    console.log('Total records:', total);
    
    console.log('Executing data query...');
    // Use SELECT * to get all columns dynamically
    const rows = await zcql.executeZCQLQuery(`SELECT * FROM StatutoryMaster ORDER BY ROWID DESC ${limitClause}`);
    console.log('Raw rows from database:', JSON.stringify(rows, null, 2));
    
    // Process data dynamically - keep all columns as they are
    const statutoryMasterData = rows.map(r => {
      const record = { id: r.StatutoryMaster.ROWID };
      
      // Copy all fields except system fields
      Object.keys(r.StatutoryMaster).forEach(key => {
        if (key !== 'ROWID' && key !== 'CREATORID' && key !== 'CREATEDTIME' && key !== 'MODIFIEDTIME') {
          record[key] = r.StatutoryMaster[key];
        }
      });
      
      return record;
    });
    
    console.log('Processed statutory master data:', JSON.stringify(statutoryMasterData, null, 2));
    res.status(200).send({ status: 'success', data: { statutoryMasterData, total, hasMore: returnAll ? false : page * perPage < total } });
  } catch (err) {
    console.error('Error fetching statutory master records:', err);
    res.status(500).send({ status: 'failure', message: err.message || 'Failed to fetch statutory master records.' });
  }
});

// Update StatutoryMaster
app.put('/statutorymaster/:ROWID', async (req, res) => {
  try {
    const { ROWID } = req.params;
    const {
      nameofWorker,
      workerID,
      gender,
      fatherName,
      dateofBirth,
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
      dateofEntryintoService,
      designation,
      epfNo,
      uanNo,
      esiNo,
      dateonwhichCompletion,
      dateonwhichMadePayment,
      periodofSuspension,
      bankAccountNumber,
      nameofBank,
      branch,
      mobileNumber,
      gmail,
      dateofExit,
      reasonforExit,
      remarks
    } = req.body;

    // Validate required fields
    if (!nameofWorker || !String(nameofWorker).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Name of Worker is required.' });
    }
    if (!workerID || !String(workerID).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Worker ID is required.' });
    }
    if (!fatherName || !String(fatherName).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Father Name is required.' });
    }

    const { catalyst } = res.locals;
    const table = catalyst.datastore().table('StatutoryMaster');
    await table.updateRow({
      ROWID,
      NameofWorker: nameofWorker,
      WorkerID: workerID,
      Gender: gender || null,
      FatherName: fatherName,
      DateofBirth: dateofBirth || null,
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
      DateofEntryintoService: dateofEntryintoService || null,
      Designation: designation || null,
      EPFNo: epfNo || null,
      UANNo: uanNo || null,
      ESINo: esiNo || null,
      DateonwhichCompletion: dateonwhichCompletion || null,
      DateonwhichMadePayment: dateonwhichMadePayment || null,
      PeriodofSuspension: periodofSuspension || null,
      BankAccountNumber: bankAccountNumber || null,
      NameofBank: nameofBank || null,
      Branch: branch || null,
      MobileNumber: mobileNumber || null,
      Gmail: gmail || null,
      DateofExit: dateofExit || null,
      ReasonforExit: reasonforExit || null,
      Remarks: remarks || null
    });
    const updated = await table.getRow(ROWID);
    res.status(200).send({ status: 'success', data: { statutoryMaster: updated } });
  } catch (err) {
    res.status(400).send({ status: 'failure', message: err.message || 'Failed to update statutory master record.' });
  }
});

// Delete StatutoryMaster
app.delete('/statutorymaster/:ROWID', async (req, res) => {
  try {
    const { ROWID } = req.params;
    const { catalyst } = res.locals;
    const table = catalyst.datastore().table('StatutoryMaster');
    await table.deleteRow(ROWID);
    res.status(200).send({ status: 'success', data: { id: ROWID } });
  } catch (err) {
    res.status(500).send({ status: 'failure', message: err.message || 'Failed to delete statutory master record.' });
  }
});

// Helper: map Excel field names to datastore column names
function mapFieldNameToColumn(excelFieldName) {
  const fieldMapping = {
    // Personal Information
    'Name of the Worker': 'NameofWorker',
    'Worker Identity No.': 'WorkerID',
    'Gender': 'Gender',
    'Father / Spouse Name': 'FatherName',
    'Date of Birth': 'DateofBirth',
    
    // Present Address
    'Present Address': 'PresentAddress',
    'Present City': 'PresentCity',
    'Present State': 'PresentState',
    'Present Postal Code': 'PresentPostalCode',
    'Present Country': 'PresentCountry',
    
    // Permanent Address
    'Permanent address': 'PermanentAddress',
    'Permanent City': 'PermanentCity',
    'Permanent State': 'PermanentState',
    'Permanent Postal Code': 'PermanentPostalCode',
    'Permanent Country': 'PermanentCountry',
    
    // Employment Details
    'Aadhaar No.': 'AadhaarNo',
    'Date of entry into service': 'DateofEntryintoService',
    'Designation / Nature of work': 'Designation',
    'EPF No / UAN No.': 'EPFNo',
    'UAN No': 'UANNo',
    'ESI No': 'ESINo',
    
    // Service & Payment
    'Date on which completion of 480 days\r\nof service': 'DateonwhichCompletion',
    'Date on which made permanent': 'DateonwhichMadePayment',
    'Period of Suspension if any': 'PeriodofSuspension',
    
    // Banking & Contact
    'Bank A/c Number, Name of Bank,\r\nBranch (Indian Financial System Code)': 'BankAccountNumber',
    'Photo': 'Photo',
    'Mobile Number': 'MobileNumber',
    'e-mail I.D': 'Gmail',
    'Specimen Signature / Thump\r\nImpression': 'Signature',
    
    // Exit Information
    'Date of Exit': 'DateofExit',
    'Reason for Exit': 'ReasonforExit',
    'Remarks': 'Remarks',
    
    // Handle __EMPTY_ fields as fallback
    '__EMPTY_1': 'NameofWorker',
    '__EMPTY_2': 'WorkerID',
    '__EMPTY_3': 'Gender',
    '__EMPTY_4': 'FatherName',
    '__EMPTY_5': 'DateofBirth',
    '__EMPTY_6': 'PresentAddress',
    '__EMPTY_7': 'PermanentAddress',
    '__EMPTY_8': 'AadhaarNo',
    '__EMPTY_9': 'DateofEntryintoService',
    '__EMPTY_10': 'Designation',
    '__EMPTY_11': 'EPFNo',
    '__EMPTY_12': 'ESINo',
    '__EMPTY_13': 'DateonwhichCompletion',
    '__EMPTY_14': 'DateonwhichMadePayment',
    '__EMPTY_15': 'PeriodofSuspension',
    '__EMPTY_16': 'BankAccountNumber',
    '__EMPTY_17': 'Photo',
    '__EMPTY_18': 'MobileNumber',
    '__EMPTY_19': 'Gmail',
    '__EMPTY_20': 'Signature',
    '__EMPTY_21': 'DateofExit',
    '__EMPTY_22': 'ReasonforExit',
    '__EMPTY_23': 'Remarks'
  };
  
  // Return mapped column name or original if no mapping found
  return fieldMapping[excelFieldName] || excelFieldName;
}

// Bulk insert endpoint for Excel data
app.post('/statutorymaster/bulk', async (req, res) => {
  try {
    const { records } = req.body;
    
    if (!records || !Array.isArray(records)) {
      return res.status(400).json({ error: 'Records array is required' });
    }

    console.log(`Processing bulk insert for ${records.length} records`);
    const { catalyst } = res.locals;
    const table = catalyst.datastore().table('StatutoryMaster');
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

module.exports = app;