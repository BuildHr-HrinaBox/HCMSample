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
  res.status(200).send('states_function ready');
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

// ===== States CRUD =====
// Table: States (columns: StateName)

// Create State
app.post('/states', async (req, res) => {
  try {
    const { stateName } = req.body;
    if (!stateName || !String(stateName).trim()) {
      return res.status(400).send({ status: 'failure', message: 'State name is required.' });
    }
    const { catalyst } = res.locals;
    const table = catalyst.datastore().table('States');
    const insertResp = await table.insertRow({
      StateName: stateName,
    });
    const created = await table.getRow(insertResp.ROWID);
    res.status(200).send({ status: 'success', data: { state: created } });
  } catch (err) {
    res.status(400).send({ status: 'failure', message: err.message || 'Failed to create state.' });
  }
});

// List States with optional pagination
app.get('/states', async (req, res) => {
  try {
    const { catalyst } = res.locals;
    const zcql = catalyst.zcql();
    const page = parseInt(req.query.page, 10) || 1;
    const perPage = parseInt(req.query.perPage, 10) || 50;
    const returnAll = !req.query.page && !req.query.perPage;
    const limitClause = returnAll ? '' : `LIMIT ${(page - 1) * perPage + 1},${perPage}`;
    const countRows = await zcql.executeZCQLQuery('SELECT COUNT(ROWID) as count FROM States');
    const total = parseInt(countRows[0].States.count, 10) || 0;
    const rows = await zcql.executeZCQLQuery(`SELECT ROWID, StateName, CREATEDTIME, MODIFIEDTIME FROM States ORDER BY ROWID DESC ${limitClause}`);
    const states = rows.map(r => ({
      id: r.States.ROWID,
      stateName: r.States.StateName,
      createdTime: r.States.CREATEDTIME,
      modifiedTime: r.States.MODIFIEDTIME
    }));
    res.status(200).send({ status: 'success', data: { states, total, hasMore: returnAll ? false : page * perPage < total } });
  } catch (err) {
    res.status(500).send({ status: 'failure', message: err.message || 'Failed to fetch states.' });
  }
});

// Update State
app.put('/states/:ROWID', async (req, res) => {
  try {
    const { ROWID } = req.params;
    const { stateName } = req.body;
    if (!stateName || !String(stateName).trim()) {
      return res.status(400).send({ status: 'failure', message: 'State name is required.' });
    }
    const { catalyst } = res.locals;
    const table = catalyst.datastore().table('States');
    await table.updateRow({
      ROWID,
      StateName: stateName,
    });
    const updated = await table.getRow(ROWID);
    res.status(200).send({ status: 'success', data: { state: updated } });
  } catch (err) {
    res.status(400).send({ status: 'failure', message: err.message || 'Failed to update state.' });
  }
});

// Delete State
app.delete('/states/:ROWID', async (req, res) => {
  try {
    const { ROWID } = req.params;
    const { catalyst } = res.locals;
    const table = catalyst.datastore().table('States');
    await table.deleteRow(ROWID);
    res.status(200).send({ status: 'success', data: { id: ROWID } });
  } catch (err) {
    res.status(500).send({ status: 'failure', message: err.message || 'Failed to delete state.' });
  }
});

module.exports = app;
