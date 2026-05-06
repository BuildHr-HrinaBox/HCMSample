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
	res.status(200).send('satatutory_function ready');
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

// ===== Acts CRUD =====
// Table: Acts (columns: NameoftheAct, MaximumMarks, Applicability)

// Create Act
app.post('/acts', async (req, res) => {
	try {
		const { actName, maxMarks, applicability } = req.body;
		if (!actName || !String(actName).trim()) {
			return res.status(400).send({ status: 'failure', message: 'Act name is required.' });
		}
		const { catalyst } = res.locals;
		const table = catalyst.datastore().table('Acts');
		const insertResp = await table.insertRow({
			NameoftheAct: actName,
			MaximumMarks: parseNumeric(maxMarks),
			Applicability: applicability || null,
		});
		const created = await table.getRow(insertResp.ROWID);
		res.status(200).send({ status: 'success', data: { act: created } });
	} catch (err) {
		res.status(400).send({ status: 'failure', message: err.message || 'Failed to create act.' });
	}
});

// List Acts with optional pagination
app.get('/acts', async (req, res) => {
	try {
		const { catalyst } = res.locals;
		const zcql = catalyst.zcql();
		const page = parseInt(req.query.page, 10) || 1;
		const perPage = parseInt(req.query.perPage, 10) || 50;
		const returnAll = !req.query.page && !req.query.perPage;
		const limitClause = returnAll ? '' : `LIMIT ${(page - 1) * perPage + 1},${perPage}`;
		const countRows = await zcql.executeZCQLQuery('SELECT COUNT(ROWID) as count FROM Acts');
		const total = parseInt(countRows[0].Acts.count, 10) || 0;
		const rows = await zcql.executeZCQLQuery(`SELECT ROWID, NameoftheAct, MaximumMarks, Applicability, CREATEDTIME, MODIFIEDTIME FROM Acts ORDER BY ROWID DESC ${limitClause}`);
		const acts = rows.map(r => ({
			id: r.Acts.ROWID,
			actName: r.Acts.NameoftheAct,
			maxMarks: r.Acts.MaximumMarks,
			applicability: r.Acts.Applicability,
			createdTime: r.Acts.CREATEDTIME,
			modifiedTime: r.Acts.MODIFIEDTIME
		}));
		res.status(200).send({ status: 'success', data: { acts, total, hasMore: returnAll ? false : page * perPage < total } });
	} catch (err) {
		res.status(500).send({ status: 'failure', message: err.message || 'Failed to fetch acts.' });
	}
});

// Update Act
app.put('/acts/:ROWID', async (req, res) => {
	try {
		const { ROWID } = req.params;
		const { actName, maxMarks, applicability } = req.body;
		if (!actName || !String(actName).trim()) {
			return res.status(400).send({ status: 'failure', message: 'Act name is required.' });
		}
		const { catalyst } = res.locals;
		const table = catalyst.datastore().table('Acts');
		await table.updateRow({
			ROWID,
			NameoftheAct: actName,
			MaximumMarks: parseNumeric(maxMarks),
			Applicability: applicability || null,
		});
		const updated = await table.getRow(ROWID);
		res.status(200).send({ status: 'success', data: { act: updated } });
	} catch (err) {
		res.status(400).send({ status: 'failure', message: err.message || 'Failed to update act.' });
	}
});

// Delete Act
app.delete('/acts/:ROWID', async (req, res) => {
	try {
		const { ROWID } = req.params;
		const { catalyst } = res.locals;
		const table = catalyst.datastore().table('Acts');
		await table.deleteRow(ROWID);
		res.status(200).send({ status: 'success', data: { id: ROWID } });
	} catch (err) {
		res.status(500).send({ status: 'failure', message: err.message || 'Failed to delete act.' });
	}
});

// ===== Unique IDs CRUD =====
// Table: Uniqueld (columns: ActUniqueId, ActId, ActCode, GeneratedCode)

// Create Unique ID
app.post('/unique-ids', async (req, res) => {
	try {
		const { actUniqueId, actId, actCode, generatedCode } = req.body;
		console.log('Received unique ID data:', { actUniqueId, actId, actCode, generatedCode });
		
		if (!actUniqueId || !String(actUniqueId).trim()) {
			return res.status(400).json({ status: 'failure', message: 'Act Unique ID is required.' });
		}
		
		const { catalyst } = res.locals;
		const table = catalyst.datastore().table('Uniqueld');
		
		console.log('Attempting to insert into Uniqueld table...');
		const insertResp = await table.insertRow({
			ActUniqueId: actUniqueId,
			ActId: parseNumeric(actId),
			ActCode: actCode || null,
			GeneratedCode: generatedCode || null,
		});
		
		console.log('Insert response:', insertResp);
		
		// Try to get the row if ROWID exists, otherwise return the data directly
		let created;
		if (insertResp && insertResp.ROWID) {
			try {
				created = await table.getRow(insertResp.ROWID);
			} catch (getRowErr) {
				console.log('getRow failed, using direct data:', getRowErr.message);
				created = {
					ROWID: insertResp.ROWID,
					ActUniqueId: actUniqueId,
					ActId: parseNumeric(actId),
					ActCode: actCode,
					GeneratedCode: generatedCode
				};
			}
		} else {
			// If no ROWID in response, create a temporary response
			created = {
				ROWID: 'temp_' + Date.now(),
				ActUniqueId: actUniqueId,
				ActId: parseNumeric(actId),
				ActCode: actCode,
				GeneratedCode: generatedCode
			};
		}
		
		res.status(200).json({ status: 'success', data: { uniqueId: created } });
	} catch (err) {
		console.error('Error creating unique ID:', err);
		res.status(400).json({ status: 'failure', message: err.message || 'Failed to create unique ID.' });
	}
});

// List Unique IDs with optional pagination
app.get('/unique-ids', async (req, res) => {
	try {
		const { catalyst } = res.locals;
		const zcql = catalyst.zcql();
		
		console.log('Attempting to fetch unique IDs...');
		const page = parseInt(req.query.page, 10) || 1;
		const perPage = parseInt(req.query.perPage, 10) || 50;
		const returnAll = !req.query.page && !req.query.perPage;
		const limitClause = returnAll ? '' : `LIMIT ${(page - 1) * perPage + 1},${perPage}`;
		
		const countQuery = `SELECT COUNT(ROWID) as count FROM Uniqueld`;
		const countRows = await zcql.executeZCQLQuery(countQuery);
		const total = parseInt(countRows[0].Uniqueld.count, 10) || 0;
		
		const selectQuery = `SELECT ROWID, ActUniqueId, ActId, ActCode, GeneratedCode, CREATEDTIME FROM Uniqueld ORDER BY ROWID DESC ${limitClause}`;
		const rows = await zcql.executeZCQLQuery(selectQuery);
		
		const uniqueIds = rows.map(r => ({
			id: r.Uniqueld.ROWID,
			actUniqueId: r.Uniqueld.ActUniqueId,
			actId: r.Uniqueld.ActId,
			actCode: r.Uniqueld.ActCode,
			generatedCode: r.Uniqueld.GeneratedCode,
			createdTime: r.Uniqueld.CREATEDTIME
		}));
		
		res.status(200).json({ 
			status: 'success', 
			data: { 
				uniqueIds, 
				total, 
				hasMore: returnAll ? false : page * perPage < total 
			} 
		});
	} catch (err) {
		console.error('Error fetching unique IDs:', err);
		res.status(500).json({ status: 'failure', message: err.message || 'Failed to fetch unique IDs.' });
	}
});

// Get single Unique ID
app.get('/unique-ids/:ROWID', async (req, res) => {
	try {
		const { ROWID } = req.params;
		const { catalyst } = res.locals;
		const table = catalyst.datastore().table('Uniqueld');
		const uniqueId = await table.getRow(ROWID);
		res.status(200).json({ status: 'success', data: { uniqueId } });
	} catch (err) {
		console.error('Error fetching unique ID:', err);
		res.status(404).json({ status: 'failure', message: err.message || 'Unique ID not found.' });
	}
});

// Update Unique ID
app.put('/unique-ids/:ROWID', async (req, res) => {
	try {
		const { ROWID } = req.params;
		const { actUniqueId, actId, actCode, generatedCode } = req.body;
		if (!actUniqueId || !String(actUniqueId).trim()) {
			return res.status(400).json({ status: 'failure', message: 'Act Unique ID is required.' });
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
		
		// Return the updated data directly instead of fetching again
		const updated = {
			ROWID: ROWID,
			ActUniqueId: actUniqueId,
			ActId: parseNumeric(actId),
			ActCode: actCode,
			GeneratedCode: generatedCode
		};
		res.status(200).json({ status: 'success', data: { uniqueId: updated } });
	} catch (err) {
		console.error('Error updating unique ID:', err);
		res.status(400).json({ status: 'failure', message: err.message || 'Failed to update unique ID.' });
	}
});

// Delete Unique ID
app.delete('/unique-ids/:ROWID', async (req, res) => {
	try {
		const { ROWID } = req.params;
		const { catalyst } = res.locals;
		const table = catalyst.datastore().table('Uniqueld');
		await table.deleteRow(ROWID);
		res.status(200).json({ status: 'success', data: { id: ROWID } });
	} catch (err) {
		console.error('Error deleting unique ID:', err);
		res.status(500).json({ status: 'failure', message: err.message || 'Failed to delete unique ID.' });
	}
});

module.exports = app;
