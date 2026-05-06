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
	res.status(200).send('Regulations_function ready');
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

// ===== Regulations CRUD =====
// Table: Regulations (columns: Regulations, Description, Type, State, Sector)

// Create Regulation
app.post('/regulations', async (req, res) => {
	try {
		const { regulations, description, type, state, sector } = req.body;
		if (!regulations || !String(regulations).trim()) {
			return res.status(400).json({ status: 'failure', message: 'Regulation name is required.' });
		}
		const { catalyst } = res.locals;
		const table = catalyst.datastore().table('Regulations');
		const insertResp = await table.insertRow({
			Regulations: regulations,
			Description: description || null,
			Type: type || null,
			State: state || null,
			Sector: sector || null,
		});
		const created = await table.getRow(insertResp.ROWID);
		res.status(200).json({ status: 'success', data: { regulation: created } });
	} catch (err) {
		console.error('Error creating regulation:', err);
		res.status(400).json({ status: 'failure', message: err.message || 'Failed to create regulation.' });
	}
});

// List Regulations with optional pagination and filtering
app.get('/regulations', async (req, res) => {
	try {
		const { catalyst } = res.locals;
		const zcql = catalyst.zcql();
		const page = parseInt(req.query.page, 10) || 1;
		const perPage = parseInt(req.query.perPage, 10) || 50;
		const regulations = req.query.regulations;
		const returnAll = !req.query.page && !req.query.perPage;
		const limitClause = returnAll ? '' : `LIMIT ${(page - 1) * perPage + 1},${perPage}`;
		
		let whereClause = '';
		if (regulations) {
			whereClause = `WHERE Regulations LIKE '%${regulations}%'`;
		}
		
		const countQuery = `SELECT COUNT(ROWID) as count FROM Regulations ${whereClause}`;
		const countRows = await zcql.executeZCQLQuery(countQuery);
		const total = parseInt(countRows[0].Regulations.count, 10) || 0;
		
		const selectQuery = `SELECT ROWID, Regulations, Description, Type, State, Sector, CREATEDTIME, MODIFIEDTIME FROM Regulations ${whereClause} ORDER BY ROWID DESC ${limitClause}`;
		const rows = await zcql.executeZCQLQuery(selectQuery);
		
		const regulationsList = rows.map(r => ({
			id: r.Regulations.ROWID,
			regulations: r.Regulations.Regulations,
			description: r.Regulations.Description,
			type: r.Regulations.Type,
			state: r.Regulations.State,
			sector: r.Regulations.Sector,
			createdTime: r.Regulations.CREATEDTIME,
			modifiedTime: r.Regulations.MODIFIEDTIME
		}));
		
		res.status(200).json({ 
			status: 'success', 
			data: { 
				regulations: regulationsList, 
				total, 
				hasMore: returnAll ? false : page * perPage < total 
			} 
		});
	} catch (err) {
		console.error('Error fetching regulations:', err);
		res.status(500).json({ status: 'failure', message: err.message || 'Failed to fetch regulations.' });
	}
});

// Get single Regulation
app.get('/regulations/:ROWID', async (req, res) => {
	try {
		const { ROWID } = req.params;
		const { catalyst } = res.locals;
		const table = catalyst.datastore().table('Regulations');
		const regulation = await table.getRow(ROWID);
		res.status(200).json({ status: 'success', data: { regulation } });
	} catch (err) {
		console.error('Error fetching regulation:', err);
		res.status(404).json({ status: 'failure', message: err.message || 'Regulation not found.' });
	}
});

// Update Regulation
app.put('/regulations/:ROWID', async (req, res) => {
	try {
		const { ROWID } = req.params;
		const { regulations, description, type, state, sector } = req.body;
		if (!regulations || !String(regulations).trim()) {
			return res.status(400).json({ status: 'failure', message: 'Regulation name is required.' });
		}
		const { catalyst } = res.locals;
		const table = catalyst.datastore().table('Regulations');
		await table.updateRow({
			ROWID,
			Regulations: regulations,
			Description: description || null,
			Type: type || null,
			State: state || null,
			Sector: sector || null,
		});
		const updated = await table.getRow(ROWID);
		res.status(200).json({ status: 'success', data: { regulation: updated } });
	} catch (err) {
		console.error('Error updating regulation:', err);
		res.status(400).json({ status: 'failure', message: err.message || 'Failed to update regulation.' });
	}
});

// Delete Regulation
app.delete('/regulations/:ROWID', async (req, res) => {
	try {
		const { ROWID } = req.params;
		const { catalyst } = res.locals;
		const table = catalyst.datastore().table('Regulations');
		await table.deleteRow(ROWID);
		res.status(200).json({ status: 'success', data: { id: ROWID } });
	} catch (err) {
		console.error('Error deleting regulation:', err);
		res.status(500).json({ status: 'failure', message: err.message || 'Failed to delete regulation.' });
	}
});

module.exports = app;
