'use strict';

const express = require('express');
const catalystSDK = require('zcatalyst-sdk-node');

const app = express();

app.use(express.json());

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
  res.status(200).send('compliance_function ready');
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

// ===== Compliance CRUD =====
// Table: Compliance (columns: DateofAuditFrom, PeriodforAuditFrom, ClientName, ClientEmailAddress, Act, DateofAuditTo, PeriodforAuditTo, Address, ClientDetails, RecordCategory, MaximumMarks, Applicability, MarkObtained)

// Create Compliance record
app.post('/compliance', async (req, res) => {
  try {
    const {
      DateofAuditFrom,
      PeriodforAuditFrom,
      ClientName,
      ClientEmailAddress,
      Act,
      DateofAuditTo,
      PeriodforAuditTo,
      Address,
      ClientDetails,
      RecordCategory,
      MaximumMarks,
      Applicability,
      MarkObtained,
      categoryScoreRows
    } = req.body;

    if (!ClientName || !String(ClientName).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Client Name is required.' });
    }
    if (!Act || !String(Act).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Act is required.' });
    }

    const { catalyst } = res.locals;
    const table = catalyst.datastore().table('Compliance');
    
    // Create main compliance record first
    const mainRecord = {
      DateofAuditFrom: DateofAuditFrom || null,
      PeriodforAuditFrom: PeriodforAuditFrom || null,
      ClientName: ClientName.trim(),
      ClientEmailAddress: ClientEmailAddress || null,
      Act: Act.trim(),
      DateofAuditTo: DateofAuditTo || null,
      PeriodforAuditTo: PeriodforAuditTo || null,
      Address: Address || null,
      ClientDetails: ClientDetails || null,
      RecordCategory: RecordCategory || null,
      MaximumMarks: MaximumMarks || null,
      Applicability: Applicability || null,
      MarkObtained: MarkObtained || null
    };
    
    const insertResp = await table.insertRow(mainRecord);
    const created = await table.getRow(insertResp.ROWID);
    
    // If categoryScoreRows exist, create separate records for each row
    if (categoryScoreRows && Array.isArray(categoryScoreRows) && categoryScoreRows.length > 0) {
      const categoryScoreRecords = [];
      
      for (const row of categoryScoreRows) {
        // Only create a record if the row has meaningful data
        if (row.ClientDetails || row.RecordCategory || row.MaximumMarks || row.Applicability || row.MarkObtained) {
          const categoryRecord = {
            DateofAuditFrom: DateofAuditFrom || null,
            PeriodforAuditFrom: PeriodforAuditFrom || null,
            ClientName: ClientName.trim(),
            ClientEmailAddress: ClientEmailAddress || null,
            Act: Act.trim(),
            DateofAuditTo: DateofAuditTo || null,
            PeriodforAuditTo: PeriodforAuditTo || null,
            Address: Address || null,
            ClientDetails: row.ClientDetails || null,
            RecordCategory: row.RecordCategory || null,
            MaximumMarks: row.MaximumMarks || null,
            Applicability: row.Applicability || null,
            MarkObtained: row.MarkObtained || null
          };
          
          const categoryInsertResp = await table.insertRow(categoryRecord);
          const categoryCreated = await table.getRow(categoryInsertResp.ROWID);
          categoryScoreRecords.push(categoryCreated);
        }
      }
      
      res.status(200).send({ 
        status: 'success', 
        data: { 
          compliance: created,
          categoryScoreRecords: categoryScoreRecords,
          totalRecords: 1 + categoryScoreRecords.length
        } 
      });
    } else {
      res.status(200).send({ status: 'success', data: { compliance: created } });
    }
  } catch (err) {
    res.status(400).send({ status: 'failure', message: err.message || 'Failed to create compliance record.' });
  }
});

// List Compliance records with optional pagination
app.get('/compliance', async (req, res) => {
  try {
    const { catalyst } = res.locals;
    const zcql = catalyst.zcql();
    const page = parseInt(req.query.page, 10) || 1;
    const perPage = parseInt(req.query.perPage, 10) || 50;
    const returnAll = !req.query.page && !req.query.perPage;
    const limitClause = returnAll ? '' : `LIMIT ${(page - 1) * perPage + 1},${perPage}`;
    const countRows = await zcql.executeZCQLQuery('SELECT COUNT(ROWID) as count FROM Compliance');
    const total = parseInt(countRows[0].Compliance.count, 10) || 0;
    const rows = await zcql.executeZCQLQuery("SELECT ROWID, DateofAuditFrom, PeriodforAuditFrom, ClientName, ClientEmailAddress, Act, DateofAuditTo, PeriodforAuditTo, Address, ClientDetails, RecordCategory, MaximumMarks, Applicability, MarkObtained, CREATEDTIME, MODIFIEDTIME FROM Compliance ORDER BY ROWID DESC " + limitClause);
    const complianceRecords = rows.map(r => ({
      id: r.Compliance.ROWID,
      DateofAuditFrom: r.Compliance.DateofAuditFrom,
      PeriodforAuditFrom: r.Compliance.PeriodforAuditFrom,
      ClientName: r.Compliance.ClientName,
      ClientEmailAddress: r.Compliance.ClientEmailAddress,
      Act: r.Compliance.Act,
      DateofAuditTo: r.Compliance.DateofAuditTo,
      PeriodforAuditTo: r.Compliance.PeriodforAuditTo,
      Address: r.Compliance.Address,
      ClientDetails: r.Compliance.ClientDetails,
      RecordCategory: r.Compliance.RecordCategory,
      MaximumMarks: r.Compliance.MaximumMarks,
      Applicability: r.Compliance.Applicability,
      MarkObtained: r.Compliance.MarkObtained,
      createdTime: r.Compliance.CREATEDTIME,
      modifiedTime: r.Compliance.MODIFIEDTIME
    }));
    res.status(200).send({ status: 'success', data: { complianceRecords, total, hasMore: returnAll ? false : page * perPage < total } });
  } catch (err) {
    res.status(500).send({ status: 'failure', message: err.message || 'Failed to fetch compliance records.' });
  }
});

// Update Compliance record
app.put('/compliance/:ROWID', async (req, res) => {
  try {
    const { ROWID } = req.params;
    const {
      DateofAuditFrom,
      PeriodforAuditFrom,
      ClientName,
      ClientEmailAddress,
      Act,
      DateofAuditTo,
      PeriodforAuditTo,
      Address,
      ClientDetails,
      RecordCategory,
      MaximumMarks,
      Applicability,
      MarkObtained
    } = req.body;

    if (!ClientName || !String(ClientName).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Client Name is required.' });
    }
    if (!Act || !String(Act).trim()) {
      return res.status(400).send({ status: 'failure', message: 'Act is required.' });
    }

    const { catalyst } = res.locals;
    const table = catalyst.datastore().table('Compliance');
    await table.updateRow({
      ROWID,
      DateofAuditFrom: DateofAuditFrom || null,
      PeriodforAuditFrom: PeriodforAuditFrom || null,
      ClientName: ClientName.trim(),
      ClientEmailAddress: ClientEmailAddress || null,
      Act: Act.trim(),
      DateofAuditTo: DateofAuditTo || null,
      PeriodforAuditTo: PeriodforAuditTo || null,
      Address: Address || null,
      ClientDetails: ClientDetails || null,
      RecordCategory: RecordCategory || null,
      MaximumMarks: MaximumMarks || null,
      Applicability: Applicability || null,
      MarkObtained: MarkObtained || null
    });
    const updated = await table.getRow(ROWID);
    res.status(200).send({ status: 'success', data: { compliance: updated } });
  } catch (err) {
    res.status(400).send({ status: 'failure', message: err.message || 'Failed to update compliance record.' });
  }
});

// Delete Compliance record
app.delete('/compliance/:ROWID', async (req, res) => {
  try {
    const { ROWID } = req.params;
    const { catalyst } = res.locals;
    const table = catalyst.datastore().table('Compliance');
    await table.deleteRow(ROWID);
    res.status(200).send({ status: 'success', data: { id: ROWID } });
  } catch (err) {
    res.status(500).send({ status: 'failure', message: err.message || 'Failed to delete compliance record.' });
  }
});

module.exports = app;
