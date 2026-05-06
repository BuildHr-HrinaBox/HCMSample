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
	res.status(200).send('unique_function ready');
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

// ===== Unique IDs CRUD =====
// Table: Uniqueld (columns: ActUniqueId, ActId, ActCode, GeneratedCode)

// Create Unique ID
app.post('/unique-ids', async (req, res) => {
	try {
		const { actUniqueId, actId, actCode, generatedCode } = req.body;
		if (!actUniqueId || !String(actUniqueId).trim()) {
			return res.status(400).send({ status: 'failure', message: 'Act Unique ID is required.' });
		}
		const { catalyst } = res.locals;
		const table = catalyst.datastore().table('Uniqueld');
		const insertResp = await table.insertRow({
			ActUniqueId: actUniqueId,
			ActId: parseNumeric(actId),
			ActCode: actCode || null,
			GeneratedCode: generatedCode || null,
		});
		const created = await table.getRow(insertResp.ROWID);
		res.status(200).send({ status: 'success', data: { uniqueId: created } });
	} catch (err) {
		res.status(400).send({ status: 'failure', message: err.message || 'Failed to create unique ID.' });
	}
});

// List Unique IDs with optional pagination
app.get('/unique-ids', async (req, res) => {
	try {
		const { catalyst } = res.locals;
		const zcql = catalyst.zcql();
		const page = parseInt(req.query.page, 10) || 1;
		const perPage = parseInt(req.query.perPage, 10) || 50;
		const returnAll = !req.query.page && !req.query.perPage;
		const limitClause = returnAll ? '' : `LIMIT ${(page - 1) * perPage + 1},${perPage}`;
		const countRows = await zcql.executeZCQLQuery('SELECT COUNT(ROWID) as count FROM Uniqueld');
		const total = parseInt(countRows[0].Uniqueld.count, 10) || 0;
		const rows = await zcql.executeZCQLQuery(`SELECT ROWID, ActUniqueId, ActId, ActCode, GeneratedCode, CREATEDTIME, MODIFIEDTIME FROM Uniqueld ORDER BY ROWID DESC ${limitClause}`);
		const uniqueIds = rows.map(r => ({
			id: r.Uniqueld.ROWID,
			actUniqueId: r.Uniqueld.ActUniqueId,
			actId: r.Uniqueld.ActId,
			actCode: r.Uniqueld.ActCode,
			generatedCode: r.Uniqueld.GeneratedCode,
			createdTime: r.Uniqueld.CREATEDTIME,
			modifiedTime: r.Uniqueld.MODIFIEDTIME
		}));
		res.status(200).send({ status: 'success', data: { uniqueIds, total, hasMore: returnAll ? false : page * perPage < total } });
	} catch (err) {
		res.status(500).send({ status: 'failure', message: err.message || 'Failed to fetch unique IDs.' });
	}
});

// Update Unique ID
app.put('/unique-ids/:ROWID', async (req, res) => {
	try {
		const { ROWID } = req.params;
		const { actUniqueId, actId, actCode, generatedCode } = req.body;
		if (!actUniqueId || !String(actUniqueId).trim()) {
			return res.status(400).send({ status: 'failure', message: 'Act Unique ID is required.' });
		}
		const { catalyst } = res.locals;
		const table = catalyst.datastore().table('Uniqueld');
		await table.updateRow({
			ROWID,
			ActUniqueId: actUniqueId,
			ActId: parseNumeric(actId),
			ActCode: actCode || null,
			GeneratedCode: generatedCode || null,
		});
		const updated = await table.getRow(ROWID);
		res.status(200).send({ status: 'success', data: { uniqueId: updated } });
	} catch (err) {
		res.status(400).send({ status: 'failure', message: err.message || 'Failed to update unique ID.' });
	}
});

// Delete Unique ID
app.delete('/unique-ids/:ROWID', async (req, res) => {
	try {
		const { ROWID } = req.params;
		const { catalyst } = res.locals;
		const table = catalyst.datastore().table('Uniqueld');
		await table.deleteRow(ROWID);
		res.status(200).send({ status: 'success', data: { id: ROWID } });
	} catch (err) {
		res.status(500).send({ status: 'failure', message: err.message || 'Failed to delete unique ID.' });
	}
});

module.exports = app;
