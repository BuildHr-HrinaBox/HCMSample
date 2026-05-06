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
  res.status(200).json({ status: 'success', message: 'form15II_function ready' });
});

// Test endpoint to check table access
app.get('/test-table', async (req, res) => {
  try {
    console.log('Testing table access...');
    const { catalyst } = res.locals;
    const table = catalyst.datastore().table('Form15II');
    
    // Try to get table info
    console.log('Table object:', table);
    
    // Try a simple query
    const zcql = catalyst.zcql();
    const testQuery = await zcql.executeZCQLQuery('SELECT COUNT(ROWID) as count FROM Form15II');
    console.log('Test query result:', testQuery);
    
    res.status(200).send({ 
      status: 'success', 
      message: 'Table access successful',
      tableInfo: {
        tableName: 'Form15II',
        recordCount: testQuery[0]?.Form15II?.count || 0
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

// ===== Form15II CRUD =====
// Table: Form15II (columns: RegisterofAdultworkersandYoungpersons, Nameoftheworker, WorkerIdentityNumber, Numberofdaysworked, BasicWage, DearnessAllowance, HouseRentAllowance, OtherAllowances, OvertimeWages, LeaveWages, GrossWages, ProvidentFund, EmployeesStateInsurance, LabourWelfareFund, AdvancePaid, Advancerecovery, AdvanceRecovered, PendingRecovery, Deductionimposed, Deductionrecoverypending, DeductionmadeonDamages, DamagesPendingRecovery, AnyotherDeductions, TotalDeductions, NetWages, Dateofpayment, Unpaidaccumulations, ReceiptbyEmployee, Remarks)

// Create Form15II
app.post('/form15II', async (req, res) => {
  try {
    console.log('Received request body:', JSON.stringify(req.body, null, 2));
    
    const {
      registerOfAdultWorkersAndYoungPersons,
      nameOfWorker,
      workerIdentityNumber,
      numberOfDaysWorked,
      basicWage,
      dearnessAllowance,
      houseRentAllowance,
      otherAllowances,
      overtimeWages,
      leaveWages,
      grossWages,
      providentFund,
      employeesStateInsurance,
      labourWelfareFund,
      advancePaid,
      advanceRecovery,
      advanceRecovered,
      pendingRecovery,
      deductionImposed,
      deductionRecoveryPending,
      deductionMadeOnDamages,
      damagesPendingRecovery,
      anyOtherDeductions,
      totalDeductions,
      netWages,
      dateOfPayment,
      unpaidAccumulations,
      receiptByEmployee,
      remarks
    } = req.body;

    // Validate required fields
    if (!nameOfWorker || !String(nameOfWorker).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Name of Worker is required.' });
    }
    if (!workerIdentityNumber || !String(workerIdentityNumber).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Worker Identity Number is required.' });
    }
    if (!basicWage || !String(basicWage).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Basic Wage is required.' });
    }
    if (!grossWages || !String(grossWages).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Gross Wages is required.' });
    }

    const { catalyst } = res.locals;
    const table = catalyst.datastore().table('Form15II');
    
    const insertData = {
      RegisterofAdultworkersandYoungpersons: registerOfAdultWorkersAndYoungPersons || null,
      Nameoftheworker: nameOfWorker,
      WorkerIdentityNumber: workerIdentityNumber,
      Numberofdaysworked: parseNumeric(numberOfDaysWorked) || null,
      BasicWage: parseNumeric(basicWage) || null,
      DearnessAllowance: parseNumeric(dearnessAllowance) || null,
      HouseRentAllowance: parseNumeric(houseRentAllowance) || null,
      OtherAllowances: parseNumeric(otherAllowances) || null,
      OvertimeWages: parseNumeric(overtimeWages) || null,
      LeaveWages: parseNumeric(leaveWages) || null,
      GrossWages: parseNumeric(grossWages) || null,
      ProvidentFund: parseNumeric(providentFund) || null,
      EmployeesStateInsurance: parseNumeric(employeesStateInsurance) || null,
      LabourWelfareFund: parseNumeric(labourWelfareFund) || null,
      AdvancePaid: parseNumeric(advancePaid) || null,
      Advancerecovery: parseNumeric(advanceRecovery) || null,
      AdvanceRecovered: parseNumeric(advanceRecovered) || null,
      PendingRecovery: parseNumeric(pendingRecovery) || null,
      Deductionimposed: parseNumeric(deductionImposed) || null,
      Deductionrecoverypending: parseNumeric(deductionRecoveryPending) || null,
      DeductionmadeonDamages: parseNumeric(deductionMadeOnDamages) || null,
      DamagesPendingRecovery: parseNumeric(damagesPendingRecovery) || null,
      AnyotherDeductions: parseNumeric(anyOtherDeductions) || null,
      TotalDeductions: parseNumeric(totalDeductions) || null,
      NetWages: parseNumeric(netWages) || null,
      Dateofpayment: dateOfPayment || null,
      Unpaidaccumulations: parseNumeric(unpaidAccumulations) || null,
      ReceiptbyEmployee: receiptByEmployee || null,
      Remarks: remarks || null
    };
    
    console.log('Inserting data to Form15II table:', JSON.stringify(insertData, null, 2));
    
    const insertResp = await table.insertRow(insertData);
    console.log('Insert response:', JSON.stringify(insertResp, null, 2));
    
    const created = await table.getRow(insertResp.ROWID);
    console.log('Created record:', JSON.stringify(created, null, 2));
    
    res.status(200).send({ status: 'success', data: { form15II: created } });
  } catch (err) {
    console.error('Error creating form15II record:', err);
    res.status(400).send({ status: 'failure', message: err.message || 'Failed to create form15II record.' });
  }
});

// Get single Form15II record with all details
app.get('/form15II/:ROWID', async (req, res) => {
  try {
    const { ROWID } = req.params;
    console.log('Fetching form15II record:', ROWID);
    const { catalyst } = res.locals;
    const table = catalyst.datastore().table('Form15II');
    
    const form15II = await table.getRow(ROWID);
    console.log('Retrieved form15II record:', JSON.stringify(form15II, null, 2));
    
    res.status(200).send({ status: 'success', data: { form15II } });
  } catch (err) {
    console.error('Error fetching form15II record:', err);
    res.status(500).send({ status: 'failure', message: err.message || 'Failed to fetch form15II record.' });
  }
});

// List Form15II with optional pagination
app.get('/form15II', async (req, res) => {
  try {
    console.log('Fetching form15II records...');
    const { catalyst } = res.locals;
    const zcql = catalyst.zcql();
    const page = parseInt(req.query.page, 10) || 1;
    const perPage = parseInt(req.query.perPage, 10) || 50;
    const returnAll = !req.query.page && !req.query.perPage;
    const limitClause = returnAll ? '' : `LIMIT ${(page - 1) * perPage + 1},${perPage}`;
    
    console.log('Executing count query...');
    const countRows = await zcql.executeZCQLQuery('SELECT COUNT(ROWID) as count FROM Form15II');
    const total = parseInt(countRows[0].Form15II.count, 10) || 0;
    console.log('Total records:', total);
    
    console.log('Executing data query...');
    // Select all columns for Form15II
    const rows = await zcql.executeZCQLQuery(`SELECT ROWID, RegisterofAdultworkersandYoungpersons, Nameoftheworker, WorkerIdentityNumber, Numberofdaysworked, BasicWage, DearnessAllowance, HouseRentAllowance, OtherAllowances, OvertimeWages, LeaveWages, GrossWages, ProvidentFund, EmployeesStateInsurance, LabourWelfareFund, AdvancePaid, Advancerecovery, AdvanceRecovered, PendingRecovery, Deductionimposed, Deductionrecoverypending, DeductionmadeonDamages, DamagesPendingRecovery, AnyotherDeductions, TotalDeductions, NetWages, Dateofpayment, Unpaidaccumulations, ReceiptbyEmployee, Remarks, CREATEDTIME FROM Form15II ORDER BY ROWID DESC ${limitClause}`);
    console.log('Raw rows from database:', JSON.stringify(rows, null, 2));
    
    const form15IIData = rows.map(r => ({
      id: r.Form15II.ROWID,
      registerOfAdultWorkersAndYoungPersons: r.Form15II.RegisterofAdultworkersandYoungpersons,
      nameOfWorker: r.Form15II.Nameoftheworker,
      workerIdentityNumber: r.Form15II.WorkerIdentityNumber,
      numberOfDaysWorked: r.Form15II.Numberofdaysworked,
      basicWage: r.Form15II.BasicWage,
      dearnessAllowance: r.Form15II.DearnessAllowance,
      houseRentAllowance: r.Form15II.HouseRentAllowance,
      otherAllowances: r.Form15II.OtherAllowances,
      overtimeWages: r.Form15II.OvertimeWages,
      leaveWages: r.Form15II.LeaveWages,
      grossWages: r.Form15II.GrossWages,
      providentFund: r.Form15II.ProvidentFund,
      employeesStateInsurance: r.Form15II.EmployeesStateInsurance,
      labourWelfareFund: r.Form15II.LabourWelfareFund,
      advancePaid: r.Form15II.AdvancePaid,
      advanceRecovery: r.Form15II.Advancerecovery,
      advanceRecovered: r.Form15II.AdvanceRecovered,
      pendingRecovery: r.Form15II.PendingRecovery,
      deductionImposed: r.Form15II.Deductionimposed,
      deductionRecoveryPending: r.Form15II.Deductionrecoverypending,
      deductionMadeOnDamages: r.Form15II.DeductionmadeonDamages,
      damagesPendingRecovery: r.Form15II.DamagesPendingRecovery,
      anyOtherDeductions: r.Form15II.AnyotherDeductions,
      totalDeductions: r.Form15II.TotalDeductions,
      netWages: r.Form15II.NetWages,
      dateOfPayment: r.Form15II.Dateofpayment,
      unpaidAccumulations: r.Form15II.Unpaidaccumulations,
      receiptByEmployee: r.Form15II.ReceiptbyEmployee,
      remarks: r.Form15II.Remarks,
      createdTime: r.Form15II.CREATEDTIME,
      modifiedTime: null // Not selected in query
    }));
    
    console.log('Processed form15II data:', JSON.stringify(form15IIData, null, 2));
    res.status(200).send({ status: 'success', data: { form15IIData, total, hasMore: returnAll ? false : page * perPage < total } });
  } catch (err) {
    console.error('Error fetching form15II records:', err);
    res.status(500).send({ status: 'failure', message: err.message || 'Failed to fetch form15II records.' });
  }
});

// Update Form15II
app.put('/form15II/:ROWID', async (req, res) => {
  try {
    const { ROWID } = req.params;
    const {
      registerOfAdultWorkersAndYoungPersons,
      nameOfWorker,
      workerIdentityNumber,
      numberOfDaysWorked,
      basicWage,
      dearnessAllowance,
      houseRentAllowance,
      otherAllowances,
      overtimeWages,
      leaveWages,
      grossWages,
      providentFund,
      employeesStateInsurance,
      labourWelfareFund,
      advancePaid,
      advanceRecovery,
      advanceRecovered,
      pendingRecovery,
      deductionImposed,
      deductionRecoveryPending,
      deductionMadeOnDamages,
      damagesPendingRecovery,
      anyOtherDeductions,
      totalDeductions,
      netWages,
      dateOfPayment,
      unpaidAccumulations,
      receiptByEmployee,
      remarks
    } = req.body;

    // Validate required fields
    if (!nameOfWorker || !String(nameOfWorker).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Name of Worker is required.' });
    }
    if (!workerIdentityNumber || !String(workerIdentityNumber).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Worker Identity Number is required.' });
    }
    if (!basicWage || !String(basicWage).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Basic Wage is required.' });
    }
    if (!grossWages || !String(grossWages).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Gross Wages is required.' });
    }

    const { catalyst } = res.locals;
    const table = catalyst.datastore().table('Form15II');
    await table.updateRow({
      ROWID,
      RegisterofAdultworkersandYoungpersons: registerOfAdultWorkersAndYoungPersons || null,
      Nameoftheworker: nameOfWorker,
      WorkerIdentityNumber: workerIdentityNumber,
      Numberofdaysworked: parseNumeric(numberOfDaysWorked) || null,
      BasicWage: parseNumeric(basicWage) || null,
      DearnessAllowance: parseNumeric(dearnessAllowance) || null,
      HouseRentAllowance: parseNumeric(houseRentAllowance) || null,
      OtherAllowances: parseNumeric(otherAllowances) || null,
      OvertimeWages: parseNumeric(overtimeWages) || null,
      LeaveWages: parseNumeric(leaveWages) || null,
      GrossWages: parseNumeric(grossWages) || null,
      ProvidentFund: parseNumeric(providentFund) || null,
      EmployeesStateInsurance: parseNumeric(employeesStateInsurance) || null,
      LabourWelfareFund: parseNumeric(labourWelfareFund) || null,
      AdvancePaid: parseNumeric(advancePaid) || null,
      Advancerecovery: parseNumeric(advanceRecovery) || null,
      AdvanceRecovered: parseNumeric(advanceRecovered) || null,
      PendingRecovery: parseNumeric(pendingRecovery) || null,
      Deductionimposed: parseNumeric(deductionImposed) || null,
      Deductionrecoverypending: parseNumeric(deductionRecoveryPending) || null,
      DeductionmadeonDamages: parseNumeric(deductionMadeOnDamages) || null,
      DamagesPendingRecovery: parseNumeric(damagesPendingRecovery) || null,
      AnyotherDeductions: parseNumeric(anyOtherDeductions) || null,
      TotalDeductions: parseNumeric(totalDeductions) || null,
      NetWages: parseNumeric(netWages) || null,
      Dateofpayment: dateOfPayment || null,
      Unpaidaccumulations: parseNumeric(unpaidAccumulations) || null,
      ReceiptbyEmployee: receiptByEmployee || null,
      Remarks: remarks || null
    });
    const updated = await table.getRow(ROWID);
    res.status(200).send({ status: 'success', data: { form15II: updated } });
  } catch (err) {
    res.status(400).send({ status: 'failure', message: err.message || 'Failed to update form15II record.' });
  }
});

// Delete Form15II
app.delete('/form15II/:ROWID', async (req, res) => {
  try {
    const { ROWID } = req.params;
    const { catalyst } = res.locals;
    const table = catalyst.datastore().table('Form15II');
    await table.deleteRow(ROWID);
    res.status(200).send({ status: 'success', data: { id: ROWID } });
  } catch (err) {
    res.status(500).send({ status: 'failure', message: err.message || 'Failed to delete form15II record.' });
  }
});

module.exports = app;
