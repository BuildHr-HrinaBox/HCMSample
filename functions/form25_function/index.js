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
  res.status(200).json({ status: 'success', message: 'form25_function ready' });
});

// Test endpoint to check table access
app.get('/test-table', async (req, res) => {
  try {
    console.log('Testing table access...');
    const { catalyst } = res.locals;
    const table = catalyst.datastore().table('Form25');
    
    // Try to get table info
    console.log('Table object:', table);
    
    // Try a simple query
    const zcql = catalyst.zcql();
    const testQuery = await zcql.executeZCQLQuery('SELECT COUNT(ROWID) as count FROM Form25');
    console.log('Test query result:', testQuery);
    
    res.status(200).send({ 
      status: 'success', 
      message: 'Table access successful',
      tableInfo: {
        tableName: 'Form25',
        recordCount: testQuery[0]?.Form25?.count || 0
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

// ===== Form25 CRUD =====
// Table: Form25 (columns: RegisterofAdultworkersandyoungpersons, Nameoftheworker, WorkerIdentityNumber, Timeatwhichworkcommences, RestInterval, Timeatwhichworkends, SchemeofShifts, TotalDaysWorked, TotalHoursWorked, NumberofdaysonLossofPay, BenefitavailedforNationalHoliday, BenefitavailedFestivalHoliday, Remarks)

// Create Form25
app.post('/form25', async (req, res) => {
  try {
    console.log('Received request body:', JSON.stringify(req.body, null, 2));
    
    const {
      registerOfAdultWorkersAndYoungPersons,
      nameOfTheWorker,
      workerIdentityNumber,
      timeAtWhichWorkCommences,
      restInterval,
      timeAtWhichWorkEnds,
      schemeOfShifts,
      totalDaysWorked,
      totalHoursWorked,
      numberOfDaysOnLossOfPay,
      benefitAvailedForNationalHoliday,
      benefitAvailedFestivalHoliday,
      remarks
    } = req.body;

    // Validate required fields
    if (!nameOfTheWorker || !String(nameOfTheWorker).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Name of the Worker is required.' });
    }
    if (!workerIdentityNumber || !String(workerIdentityNumber).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Worker Identity Number is required.' });
    }
    if (!timeAtWhichWorkCommences || !String(timeAtWhichWorkCommences).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Time at which work commences is required.' });
    }
    if (!timeAtWhichWorkEnds || !String(timeAtWhichWorkEnds).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Time at which work ends is required.' });
    }

    const { catalyst } = res.locals;
    const table = catalyst.datastore().table('Form25');
    
    const insertData = {
      RegisterofAdultworkersandyoungpersons: registerOfAdultWorkersAndYoungPersons || null,
      Nameoftheworker: nameOfTheWorker,
      WorkerIdentityNumber: workerIdentityNumber,
      Timeatwhichworkcommences: timeAtWhichWorkCommences,
      RestInterval: restInterval || null,
      Timeatwhichworkends: timeAtWhichWorkEnds,
      SchemeofShifts: schemeOfShifts || null,
      TotalDaysWorked: parseNumeric(totalDaysWorked),
      TotalHoursWorked: parseNumeric(totalHoursWorked),
      NumberofdaysonLossofPay: parseNumeric(numberOfDaysOnLossOfPay),
      BenefitavailedforNationalHoliday: benefitAvailedForNationalHoliday || null,
      BenefitavailedFestivalHoliday: benefitAvailedFestivalHoliday || null,
      Remarks: remarks || null
    };
    
    console.log('Inserting data to Form25 table:', JSON.stringify(insertData, null, 2));
    
    const insertResp = await table.insertRow(insertData);
    console.log('Insert response:', JSON.stringify(insertResp, null, 2));
    
    const created = await table.getRow(insertResp.ROWID);
    console.log('Created record:', JSON.stringify(created, null, 2));
    
    res.status(200).send({ status: 'success', data: { form25: created } });
  } catch (err) {
    console.error('Error creating form25 record:', err);
    res.status(400).send({ status: 'failure', message: err.message || 'Failed to create form25 record.' });
  }
});

// Get single Form25 record with all details
app.get('/form25/:ROWID', async (req, res) => {
  try {
    const { ROWID } = req.params;
    console.log('Fetching form25 record:', ROWID);
    const { catalyst } = res.locals;
    const table = catalyst.datastore().table('Form25');
    
    const form25 = await table.getRow(ROWID);
    console.log('Retrieved form25 record:', JSON.stringify(form25, null, 2));
    
    res.status(200).send({ status: 'success', data: { form25 } });
  } catch (err) {
    console.error('Error fetching form25 record:', err);
    res.status(500).send({ status: 'failure', message: err.message || 'Failed to fetch form25 record.' });
  }
});

// List Form25 with optional pagination
app.get('/form25', async (req, res) => {
  try {
    console.log('Fetching form25 records...');
    const { catalyst } = res.locals;
    const zcql = catalyst.zcql();
    const page = parseInt(req.query.page, 10) || 1;
    const perPage = parseInt(req.query.perPage, 10) || 50;
    const returnAll = !req.query.page && !req.query.perPage;
    const limitClause = returnAll ? '' : `LIMIT ${(page - 1) * perPage + 1},${perPage}`;
    
    console.log('Executing count query...');
    const countRows = await zcql.executeZCQLQuery('SELECT COUNT(ROWID) as count FROM Form25');
    const total = parseInt(countRows[0].Form25.count, 10) || 0;
    console.log('Total records:', total);
    
    console.log('Executing data query...');
    const rows = await zcql.executeZCQLQuery(`SELECT ROWID, RegisterofAdultworkersandyoungpersons, Nameoftheworker, WorkerIdentityNumber, Timeatwhichworkcommences, RestInterval, Timeatwhichworkends, SchemeofShifts, TotalDaysWorked, TotalHoursWorked, NumberofdaysonLossofPay, BenefitavailedforNationalHoliday, BenefitavailedFestivalHoliday, Remarks, CREATEDTIME FROM Form25 ORDER BY ROWID DESC ${limitClause}`);
    console.log('Raw rows from database:', JSON.stringify(rows, null, 2));
    
    const form25Data = rows.map(r => ({
      id: r.Form25.ROWID,
      registerOfAdultWorkersAndYoungPersons: r.Form25.RegisterofAdultworkersandyoungpersons,
      nameOfTheWorker: r.Form25.Nameoftheworker,
      workerIdentityNumber: r.Form25.WorkerIdentityNumber,
      timeAtWhichWorkCommences: r.Form25.Timeatwhichworkcommences,
      restInterval: r.Form25.RestInterval,
      timeAtWhichWorkEnds: r.Form25.Timeatwhichworkends,
      schemeOfShifts: r.Form25.SchemeofShifts,
      totalDaysWorked: r.Form25.TotalDaysWorked,
      totalHoursWorked: r.Form25.TotalHoursWorked,
      numberOfDaysOnLossOfPay: r.Form25.NumberofdaysonLossofPay,
      benefitAvailedForNationalHoliday: r.Form25.BenefitavailedforNationalHoliday,
      benefitAvailedFestivalHoliday: r.Form25.BenefitavailedFestivalHoliday,
      remarks: r.Form25.Remarks,
      createdTime: r.Form25.CREATEDTIME,
      modifiedTime: null // Not selected in query
    }));
    
    console.log('Processed form25 data:', JSON.stringify(form25Data, null, 2));
    res.status(200).send({ status: 'success', data: { form25Data, total, hasMore: returnAll ? false : page * perPage < total } });
  } catch (err) {
    console.error('Error fetching form25 records:', err);
    res.status(500).send({ status: 'failure', message: err.message || 'Failed to fetch form25 records.' });
  }
});

// Update Form25
app.put('/form25/:ROWID', async (req, res) => {
  try {
    const { ROWID } = req.params;
    const {
      registerOfAdultWorkersAndYoungPersons,
      nameOfTheWorker,
      workerIdentityNumber,
      timeAtWhichWorkCommences,
      restInterval,
      timeAtWhichWorkEnds,
      schemeOfShifts,
      totalDaysWorked,
      totalHoursWorked,
      numberOfDaysOnLossOfPay,
      benefitAvailedForNationalHoliday,
      benefitAvailedFestivalHoliday,
      remarks
    } = req.body;

    // Validate required fields
    if (!nameOfTheWorker || !String(nameOfTheWorker).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Name of the Worker is required.' });
    }
    if (!workerIdentityNumber || !String(workerIdentityNumber).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Worker Identity Number is required.' });
    }
    if (!timeAtWhichWorkCommences || !String(timeAtWhichWorkCommences).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Time at which work commences is required.' });
    }
    if (!timeAtWhichWorkEnds || !String(timeAtWhichWorkEnds).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Time at which work ends is required.' });
    }

    const { catalyst } = res.locals;
    const table = catalyst.datastore().table('Form25');
    await table.updateRow({
      ROWID,
      RegisterofAdultworkersandyoungpersons: registerOfAdultWorkersAndYoungPersons || null,
      Nameoftheworker: nameOfTheWorker,
      WorkerIdentityNumber: workerIdentityNumber,
      Timeatwhichworkcommences: timeAtWhichWorkCommences,
      RestInterval: restInterval || null,
      Timeatwhichworkends: timeAtWhichWorkEnds,
      SchemeofShifts: schemeOfShifts || null,
      TotalDaysWorked: parseNumeric(totalDaysWorked),
      TotalHoursWorked: parseNumeric(totalHoursWorked),
      NumberofdaysonLossofPay: parseNumeric(numberOfDaysOnLossOfPay),
      BenefitavailedforNationalHoliday: benefitAvailedForNationalHoliday || null,
      BenefitavailedFestivalHoliday: benefitAvailedFestivalHoliday || null,
      Remarks: remarks || null
    });
    const updated = await table.getRow(ROWID);
    res.status(200).send({ status: 'success', data: { form25: updated } });
  } catch (err) {
    res.status(400).send({ status: 'failure', message: err.message || 'Failed to update form25 record.' });
  }
});

// Delete Form25
app.delete('/form25/:ROWID', async (req, res) => {
  try {
    const { ROWID } = req.params;
    const { catalyst } = res.locals;
    const table = catalyst.datastore().table('Form25');
    await table.deleteRow(ROWID);
    res.status(200).send({ status: 'success', data: { id: ROWID } });
  } catch (err) {
    res.status(500).send({ status: 'failure', message: err.message || 'Failed to delete form25 record.' });
  }
});

module.exports = app;