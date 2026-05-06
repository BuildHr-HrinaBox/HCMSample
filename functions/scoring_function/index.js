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
	res.status(200).send('scoring_function ready');
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

// ===== Scoring Criteria CRUD =====
// Table: ScoringCiteria (columns: Levels, Criteriaforscoring)

// Create Scoring Criteria
app.post('/scoring-criteria', async (req, res) => {
	try {
		const { levels, criteriaforscoring } = req.body;
		console.log('Received data:', { levels, criteriaforscoring });
		
		if (!criteriaforscoring || !String(criteriaforscoring).trim()) {
			return res.status(400).json({ status: 'failure', message: 'Criteria for scoring is required.' });
		}
		const { catalyst } = res.locals;
		const table = catalyst.datastore().table('ScoringCiteria');
		
		console.log('Attempting to insert into ScoringCiteria table...');
		const insertResp = await table.insertRow({
			Levels: levels || null,
			Criteriaforscoring: criteriaforscoring,
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
					Levels: levels,
					Criteriaforscoring: criteriaforscoring || null
				};
			}
		} else {
			// If no ROWID in response, create a temporary response
			created = {
				ROWID: 'temp_' + Date.now(),
				Levels: levels,
				Criteriaforscoring: criteriaforscoring || null
			};
		}
		
		res.status(200).json({ status: 'success', data: { scoringCriteria: created } });
	} catch (err) {
		console.error('Error creating scoring criteria:', err);
		res.status(400).json({ status: 'failure', message: err.message || 'Failed to create scoring criteria.' });
	}
});

// List Scoring Criteria with optional pagination and filtering
app.get('/scoring-criteria', async (req, res) => {
	try {
		const { catalyst } = res.locals;
		const zcql = catalyst.zcql();
		
		console.log('Attempting to fetch scoring criteria...');
		const page = parseInt(req.query.page, 10) || 1;
		const perPage = parseInt(req.query.perPage, 10) || 50;
		const levels = req.query.levels;
		const returnAll = !req.query.page && !req.query.perPage;
		const limitClause = returnAll ? '' : `LIMIT ${(page - 1) * perPage + 1},${perPage}`;
		
		let whereClause = '';
		if (levels) {
			whereClause = `WHERE Levels LIKE '%${levels}%'`;
		}
		
		const countQuery = `SELECT COUNT(ROWID) as count FROM ScoringCiteria ${whereClause}`;
		const countRows = await zcql.executeZCQLQuery(countQuery);
		const total = parseInt(countRows[0].ScoringCiteria.count, 10) || 0;
		
		const selectQuery = `SELECT ROWID, Levels, Criteriaforscoring FROM ScoringCiteria ${whereClause} ORDER BY ROWID DESC ${limitClause}`;
		const rows = await zcql.executeZCQLQuery(selectQuery);
		
		const scoringCriteria = rows.map(r => ({
			id: r.ScoringCiteria.ROWID,
			levels: r.ScoringCiteria.Levels,
			criteriaforscoring: r.ScoringCiteria.Criteriaforscoring
		}));
		
		res.status(200).json({ 
			status: 'success', 
			data: { 
				scoringCriteria, 
				total, 
				hasMore: returnAll ? false : page * perPage < total 
			} 
		});
	} catch (err) {
		console.error('Error fetching scoring criteria:', err);
		res.status(500).json({ status: 'failure', message: err.message || 'Failed to fetch scoring criteria.' });
	}
});

// Get single Scoring Criteria
app.get('/scoring-criteria/:ROWID', async (req, res) => {
	try {
		const { ROWID } = req.params;
		const { catalyst } = res.locals;
		const table = catalyst.datastore().table('ScoringCiteria');
		const scoringCriteria = await table.getRow(ROWID);
		res.status(200).json({ status: 'success', data: { scoringCriteria } });
	} catch (err) {
		console.error('Error fetching scoring criteria:', err);
		res.status(404).json({ status: 'failure', message: err.message || 'Scoring criteria not found.' });
	}
});

// Update Scoring Criteria
app.put('/scoring-criteria/:ROWID', async (req, res) => {
	try {
		const { ROWID } = req.params;
		const { levels, criteriaforscoring } = req.body;
		if (!criteriaforscoring || !String(criteriaforscoring).trim()) {
			return res.status(400).json({ status: 'failure', message: 'Criteria for scoring is required.' });
		}
		const { catalyst } = res.locals;
		const table = catalyst.datastore().table('ScoringCiteria');
		await table.updateRow({
			ROWID,
			Levels: levels,
			Criteriaforscoring: criteriaforscoring || null,
		});
		
		// Return the updated data directly instead of fetching again
		const updated = {
			ROWID: ROWID,
			Levels: levels,
			Criteriaforscoring: criteriaforscoring || null
		};
		res.status(200).json({ status: 'success', data: { scoringCriteria: updated } });
	} catch (err) {
		console.error('Error updating scoring criteria:', err);
		res.status(400).json({ status: 'failure', message: err.message || 'Failed to update scoring criteria.' });
	}
});

// Delete Scoring Criteria
app.delete('/scoring-criteria/:ROWID', async (req, res) => {
	try {
		const { ROWID } = req.params;
		const { catalyst } = res.locals;
		const table = catalyst.datastore().table('ScoringCiteria');
		await table.deleteRow(ROWID);
		res.status(200).json({ status: 'success', data: { id: ROWID } });
	} catch (err) {
		console.error('Error deleting scoring criteria:', err);
		res.status(500).json({ status: 'failure', message: err.message || 'Failed to delete scoring criteria.' });
	}
});

module.exports = app;
