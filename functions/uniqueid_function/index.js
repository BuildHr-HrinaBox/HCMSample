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
	res.status(200).send('uniqueid_function ready');
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
// Table: Uniqueld (columns: Uniqueld)

// Create Unique ID
app.post('/unique-ids', async (req, res) => {
	try {
		const { actUniqueId } = req.body;
		console.log('Received unique ID data:', { actUniqueId });
		
		if (!actUniqueId || !String(actUniqueId).trim()) {
			return res.status(400).json({ status: 'failure', message: 'Act Unique ID is required.' });
		}
		
		const { catalyst } = res.locals;
		const table = catalyst.datastore().table('Uniqueld');
		
		console.log('Attempting to insert into Uniqueld table...');
		const insertResp = await table.insertRow({
			Uniqueld: actUniqueId.trim()
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
					Uniqueld: actUniqueId.trim()
				};
			}
		} else {
			// If no ROWID in response, create a temporary response
			created = {
				ROWID: 'temp_' + Date.now(),
				Uniqueld: actUniqueId.trim()
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
		
		const selectQuery = `SELECT ROWID, Uniqueld, CREATEDTIME FROM Uniqueld ORDER BY ROWID DESC ${limitClause}`;
		const rows = await zcql.executeZCQLQuery(selectQuery);
		
		const uniqueIds = rows.map(r => ({
			id: r.Uniqueld.ROWID,
			actUniqueId: r.Uniqueld.Uniqueld,
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
		const { actUniqueId } = req.body;
		if (!actUniqueId || !String(actUniqueId).trim()) {
			return res.status(400).json({ status: 'failure', message: 'Act Unique ID is required.' });
		}
		const { catalyst } = res.locals;
		const table = catalyst.datastore().table('Uniqueld');
		await table.updateRow({
			ROWID,
			Uniqueld: actUniqueId.trim()
		});
		
		// Return the updated data directly instead of fetching again
		const updated = {
			ROWID: ROWID,
			Uniqueld: actUniqueId.trim()
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