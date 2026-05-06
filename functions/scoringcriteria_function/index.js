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
		res.status(500).json({ status: 'failure', message: 'Catalyst init failed' });
	}
});

// Health check
app.get('/', (req, res) => {
	res.status(200).send('scoringcriteria_function ready');
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
// Table: ScoringCriteria (columns: Levels, Criteriaforscoring)

// Helper function to check if table exists and create if needed
async function ensureTableExists(catalyst) {
	try {
		const table = catalyst.datastore().table('ScoringCriteria');
		// Try to get table info to check if it exists
		await table.getTableDetails();
		return table;
	} catch (err) {
		if (err.errorInfo && err.errorInfo.code === 'api/request_failure') {
			console.log('ScoringCriteria table does not exist. Please create it in Zoho Catalyst console.');
			throw new Error('ScoringCriteria table does not exist. Please create the table in Zoho Catalyst console with columns: Levels (Text), Criteriaforscoring (Text)');
		}
		throw err;
	}
}

// Create Scoring Criteria
app.post('/scoring-criteria', async (req, res) => {
	try {
		const { levels, criteriaforscoring } = req.body;
		if (!criteriaforscoring || !String(criteriaforscoring).trim()) {
			return res.status(400).json({ status: 'failure', message: 'Criteria for scoring is required.' });
		}
		const { catalyst } = res.locals;
		const table = await ensureTableExists(catalyst);
		const insertResp = await table.insertRow({
			Levels: levels || null,
			Criteriaforscoring: criteriaforscoring,
		});
		const created = await table.getRow(insertResp.ROWID);
		res.status(200).json({ status: 'success', data: { scoringCriteria: created } });
	} catch (err) {
		console.error('Error creating scoring criteria:', err);
		if (err.message && err.message.includes('table does not exist')) {
			res.status(400).json({ status: 'failure', message: err.message });
		} else {
			res.status(400).json({ status: 'failure', message: err.message || 'Failed to create scoring criteria.' });
		}
	}
});

// List Scoring Criteria with optional pagination and filtering
app.get('/scoring-criteria', async (req, res) => {
	try {
		const { catalyst } = res.locals;
		await ensureTableExists(catalyst);
		const zcql = catalyst.zcql();
		const page = parseInt(req.query.page, 10) || 1;
		const perPage = parseInt(req.query.perPage, 10) || 50;
		const levels = req.query.levels;
		const returnAll = !req.query.page && !req.query.perPage;
		const limitClause = returnAll ? '' : `LIMIT ${(page - 1) * perPage + 1},${perPage}`;
		
		let whereClause = '';
		if (levels) {
			whereClause = `WHERE Levels LIKE '%${levels}%'`;
		}
		
		const countQuery = `SELECT COUNT(ROWID) as count FROM ScoringCriteria ${whereClause}`;
		const countRows = await zcql.executeZCQLQuery(countQuery);
		const total = parseInt(countRows[0].ScoringCriteria.count, 10) || 0;
		
		const selectQuery = `SELECT ROWID, Levels, Criteriaforscoring, CREATEDTIME, MODIFIEDTIME FROM ScoringCriteria ${whereClause} ORDER BY ROWID DESC ${limitClause}`;
		const rows = await zcql.executeZCQLQuery(selectQuery);
		
		const scoringCriteria = rows.map(r => ({
			id: r.ScoringCriteria.ROWID,
			levels: r.ScoringCriteria.Levels,
			criteriaforscoring: r.ScoringCriteria.Criteriaforscoring,
			createdTime: r.ScoringCriteria.CREATEDTIME,
			modifiedTime: r.ScoringCriteria.MODIFIEDTIME
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
		if (err.message && err.message.includes('table does not exist')) {
			res.status(400).json({ status: 'failure', message: err.message });
		} else {
			res.status(500).json({ status: 'failure', message: err.message || 'Failed to fetch scoring criteria.' });
		}
	}
});

// Get single Scoring Criteria
app.get('/scoring-criteria/:ROWID', async (req, res) => {
	try {
		const { ROWID } = req.params;
		const { catalyst } = res.locals;
		const table = await ensureTableExists(catalyst);
		const scoringCriteria = await table.getRow(ROWID);
		res.status(200).json({ status: 'success', data: { scoringCriteria } });
	} catch (err) {
		console.error('Error fetching scoring criteria:', err);
		if (err.message && err.message.includes('table does not exist')) {
			res.status(400).json({ status: 'failure', message: err.message });
		} else {
			res.status(404).json({ status: 'failure', message: err.message || 'Scoring criteria not found.' });
		}
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
		const table = await ensureTableExists(catalyst);
		await table.updateRow({
			ROWID,
			Levels: levels || null,
			Criteriaforscoring: criteriaforscoring,
		});
		const updated = await table.getRow(ROWID);
		res.status(200).json({ status: 'success', data: { scoringCriteria: updated } });
	} catch (err) {
		console.error('Error updating scoring criteria:', err);
		if (err.message && err.message.includes('table does not exist')) {
			res.status(400).json({ status: 'failure', message: err.message });
		} else {
			res.status(400).json({ status: 'failure', message: err.message || 'Failed to update scoring criteria.' });
		}
	}
});

// Delete Scoring Criteria
app.delete('/scoring-criteria/:ROWID', async (req, res) => {
	try {
		const { ROWID } = req.params;
		const { catalyst } = res.locals;
		const table = await ensureTableExists(catalyst);
		await table.deleteRow(ROWID);
		res.status(200).json({ status: 'success', data: { id: ROWID } });
	} catch (err) {
		console.error('Error deleting scoring criteria:', err);
		if (err.message && err.message.includes('table does not exist')) {
			res.status(400).json({ status: 'failure', message: err.message });
		} else {
			res.status(500).json({ status: 'failure', message: err.message || 'Failed to delete scoring criteria.' });
		}
	}
});

module.exports = app;