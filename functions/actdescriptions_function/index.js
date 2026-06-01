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
	res.status(200).send('actdescriptions_function ready');
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

// ===== Act Descriptions CRUD =====
// Table: ActDescription – column names (all text): ActName, Description, Type, State, Sector, Applicability, KeyComplianceRequirements, DueDate, PenaltyforNonCompliance, Registers, Status, Checkbox

// Create Act Description
app.post('/act-descriptions', async (req, res) => {
	try {
		const { actName, description, type, state, sector, applicability, keyComplianceRequirements, dueDate, penaltyforNonCompliance, registers, status, checkbox } = req.body;
		if (!actName || !String(actName).trim()) {
			return res.status(400).json({ status: 'failure', message: 'Act name is required.' });
		}
		const { catalyst } = res.locals;
		const table = catalyst.datastore().table('ActDescription');
		const insertResp = await table.insertRow({
			ActName: actName,
			Description: description || null,
			Type: type || null,
			State: state || null,
			Sector: sector || null,
			Applicability: applicability || null,
			KeyComplianceRequirements: keyComplianceRequirements || null,
			DueDate: dueDate || null,
			PenaltyforNonCompliance: penaltyforNonCompliance || null,
			Registers: registers || null,
			Status: status || 'yet to start',
			Checkbox: checkbox != null ? String(checkbox) : null,
		});
		const created = await table.getRow(insertResp.ROWID);
		res.status(200).json({ status: 'success', data: { actDescription: created } });
	} catch (err) {
		console.error('Error creating act description:', err);
		// Handle Catalyst permission errors
		if (err.statusCode === 401 || err.code === 'NO_ACCESS') {
			return res.status(401).json({ 
				status: 'failure', 
				statusCode: err.statusCode || 401,
				code: err.code || 'NO_ACCESS',
				message: err.message || 'No privileges to perform this action. Please check function permissions in catalyst-config.json.'
			});
		}
		const statusCode = err.statusCode || 400;
		res.status(statusCode).json({ 
			status: 'failure', 
			message: err.message || 'Failed to create act description.',
			error: err.toString()
		});
	}
});

// List Act Descriptions with optional pagination and filtering
app.get('/act-descriptions', async (req, res) => {
	try {
		const { catalyst } = res.locals;
		const zcql = catalyst.zcql();
		const page = parseInt(req.query.page, 10) || 1;
		const perPage = parseInt(req.query.perPage, 10) || 50;
		const actName = req.query.actName;
		const returnAll = !req.query.page && !req.query.perPage;
		const limitClause = returnAll ? '' : `LIMIT ${(page - 1) * perPage + 1},${perPage}`;
		
		let whereClause = '';
		if (actName) {
			whereClause = `WHERE ActName LIKE '%${actName}%'`;
		}
		
		const countQuery = `SELECT COUNT(ROWID) as count FROM ActDescription ${whereClause}`;
		const countRows = await zcql.executeZCQLQuery(countQuery);
		const total = parseInt(countRows[0].ActDescription.count, 10) || 0;
		
		const selectQuery = `SELECT ROWID, ActName, Description, Type, State, Sector, Applicability, KeyComplianceRequirements, DueDate, PenaltyforNonCompliance, Registers, Status, Checkbox, CREATEDTIME, MODIFIEDTIME FROM ActDescription ${whereClause} ORDER BY ROWID DESC ${limitClause}`;
		const rows = await zcql.executeZCQLQuery(selectQuery);
		
		const actDescriptions = rows.map(r => ({
			id: r.ActDescription.ROWID,
			actName: r.ActDescription.ActName,
			description: r.ActDescription.Description,
			type: r.ActDescription.Type,
			state: r.ActDescription.State,
			sector: r.ActDescription.Sector,
			applicability: r.ActDescription.Applicability,
			keyComplianceRequirements: r.ActDescription.KeyComplianceRequirements,
			dueDate: r.ActDescription.DueDate,
			penaltyforNonCompliance: r.ActDescription.PenaltyforNonCompliance,
			registers: r.ActDescription.Registers,
			status: r.ActDescription.Status || 'yet to start',
			checkbox: r.ActDescription.Checkbox || '',
			createdTime: r.ActDescription.CREATEDTIME,
			modifiedTime: r.ActDescription.MODIFIEDTIME
		}));
		
		res.status(200).json({ 
			status: 'success', 
			data: { 
				actDescriptions, 
				total, 
				hasMore: returnAll ? false : page * perPage < total 
			} 
		});
	} catch (err) {
		console.error('Error fetching act descriptions:', err);
		res.status(500).json({ status: 'failure', message: err.message || 'Failed to fetch act descriptions.' });
	}
});

// Get single Act Description
app.get('/act-descriptions/:ROWID', async (req, res) => {
	try {
		const { ROWID } = req.params;
		const { catalyst } = res.locals;
		const table = catalyst.datastore().table('ActDescription');
		const actDescription = await table.getRow(ROWID);
		res.status(200).json({ status: 'success', data: { actDescription } });
	} catch (err) {
		console.error('Error fetching act description:', err);
		res.status(404).json({ status: 'failure', message: err.message || 'Act description not found.' });
	}
});

// Update Act Description
app.put('/act-descriptions/:ROWID', async (req, res) => {
	try {
		const { ROWID } = req.params;
		const { actName, description, type, state, sector, applicability, keyComplianceRequirements, dueDate, penaltyforNonCompliance, registers, status, checkbox } = req.body;
		if (!actName || !String(actName).trim()) {
			return res.status(400).json({ status: 'failure', message: 'Act name is required.' });
		}
		const { catalyst } = res.locals;
		const table = catalyst.datastore().table('ActDescription');
		await table.updateRow({
			ROWID,
			ActName: actName,
			Description: description || null,
			Type: type || null,
			State: state || null,
			Sector: sector || null,
			Applicability: applicability || null,
			KeyComplianceRequirements: keyComplianceRequirements || null,
			DueDate: dueDate || null,
			PenaltyforNonCompliance: penaltyforNonCompliance || null,
			Registers: registers || null,
			Status: status || 'yet to start',
			Checkbox: checkbox != null ? String(checkbox) : null,
		});
		const updated = await table.getRow(ROWID);
		res.status(200).json({ status: 'success', data: { actDescription: updated } });
	} catch (err) {
		console.error('Error updating act description:', err);
		// Handle Catalyst permission errors
		if (err.statusCode === 401 || err.code === 'NO_ACCESS') {
			return res.status(401).json({ 
				status: 'failure', 
				statusCode: err.statusCode || 401,
				code: err.code || 'NO_ACCESS',
				message: err.message || 'No privileges to perform this action. Please check function permissions in catalyst-config.json.'
			});
		}
		const statusCode = err.statusCode || 400;
		res.status(statusCode).json({ 
			status: 'failure', 
			message: err.message || 'Failed to update act description.',
			error: err.toString()
		});
	}
});

// Sync all Actsbulk rows into ActDescription table (so fetched Actsbulk data is stored in ActDescription)
app.post('/act-descriptions/sync-from-bulk', async (req, res) => {
	try {
		const { acts } = req.body;
		if (!Array.isArray(acts) || acts.length === 0) {
			return res.status(200).json({ status: 'success', synced: 0, message: 'No acts to sync.' });
		}
		const { catalyst } = res.locals;
		const table = catalyst.datastore().table('ActDescription');
		const zcql = catalyst.zcql();
		let synced = 0;
		for (const row of acts) {
			const actName = row.acts || row.actName || '';
			const name = String(actName).trim();
			if (!name) continue;
			const hasMeaningfulStatus = row.status === 'Completed' || row.status === 'Pending';
			const statusVal = row.status === 'Completed' ? 'completed' : row.status === 'Pending' ? 'in progress' : (row.status || 'yet to start').toString().toLowerCase().replace(/\s+/g, ' ');
			const rowData = {
				ActName: name,
				Description: row.description || null,
				Type: row.type || null,
				State: row.states || row.state || null,
				Sector: row.sector || null,
				Applicability: row.applicability || null,
				KeyComplianceRequirements: row.keyComplianceRequirements || null,
				DueDate: row.dueDate || null,
				PenaltyforNonCompliance: row.penaltyforNonCompliance || null,
				Registers: row.registers || null,
				Status: statusVal || 'yet to start',
			};
			try {
				const findQuery = `SELECT ROWID FROM ActDescription WHERE ActName = '${name.replace(/'/g, "''")}'`;
				const existing = await zcql.executeZCQLQuery(findQuery);
				if (existing && existing.length > 0) {
					const ROWID = existing[0].ActDescription.ROWID;
					const existingRow = await table.getRow(ROWID);
					const existingStatus = existingRow && (existingRow.Status || existingRow.status);
					const existingCheckbox = existingRow && (existingRow.Checkbox || existingRow.checkbox);
					if (existingStatus) {
						rowData.Status = existingStatus;
					} else if (!hasMeaningfulStatus) {
						rowData.Status = 'yet to start';
					}
					if (existingCheckbox != null && String(existingCheckbox).trim() !== '') {
						rowData.Checkbox = existingCheckbox;
					}
					await table.updateRow({ ROWID, ...rowData });
				} else {
					await table.insertRow(rowData);
				}
				synced++;
			} catch (rowErr) {
				console.error('Error syncing act:', name, rowErr.message);
			}
		}
		res.status(200).json({ status: 'success', synced, total: acts.length });
	} catch (err) {
		console.error('Error sync-from-bulk:', err);
		if (err.statusCode === 401 || err.code === 'NO_ACCESS') {
			return res.status(401).json({ status: 'failure', statusCode: err.statusCode || 401, code: err.code, message: err.message });
		}
		res.status(err.statusCode || 500).json({ status: 'failure', message: err.message || 'Failed to sync from bulk.', error: err.toString() });
	}
});

// Update Mark as Read – saves only the Checkbox column in ActDescription table
app.patch('/act-descriptions/mark-read', async (req, res) => {
	try {
		const { actName, checkbox, read } = req.body;
		const name = actName != null ? String(actName).trim() : '';
		if (!name) {
			return res.status(400).json({ status: 'failure', message: 'Act name is required.' });
		}
		let checkboxVal = '';
		if (checkbox !== undefined) {
			checkboxVal = checkbox != null ? String(checkbox) : '';
		} else if (read !== undefined) {
			checkboxVal = read ? 'read' : '';
		} else {
			return res.status(400).json({ status: 'failure', message: 'checkbox or read is required.' });
		}

		const { catalyst } = res.locals;
		const table = catalyst.datastore().table('ActDescription');
		const zcql = catalyst.zcql();

		let ROWID = null;
		const findQuery = `SELECT ROWID FROM ActDescription WHERE ActName = '${name.replace(/'/g, "''")}'`;
		const existing = await zcql.executeZCQLQuery(findQuery);
		if (existing && existing.length > 0) {
			ROWID = existing[0].ActDescription.ROWID;
		}
		if (!ROWID) {
			const allRows = await zcql.executeZCQLQuery('SELECT ROWID, ActName FROM ActDescription');
			const nameLower = name.toLowerCase();
			for (const r of allRows || []) {
				const stored = (r.ActDescription.ActName || '').trim().toLowerCase();
				if (stored === nameLower) {
					ROWID = r.ActDescription.ROWID;
					break;
				}
			}
		}

		if (ROWID) {
			await table.updateRow({ ROWID, Checkbox: checkboxVal });
			const updated = await table.getRow(ROWID);
			return res.status(200).json({
				status: 'success',
				data: { actDescription: updated, checkbox: updated.Checkbox || checkboxVal },
			});
		}

		const insertResp = await table.insertRow({
			ActName: name,
			Status: 'yet to start',
			Checkbox: checkboxVal,
		});
		const created = await table.getRow(insertResp.ROWID);
		res.status(200).json({
			status: 'success',
			data: { actDescription: created, checkbox: created.Checkbox || checkboxVal },
			created: true,
		});
	} catch (err) {
		console.error('Error saving mark as read:', err);
		if (err.statusCode === 401 || err.code === 'NO_ACCESS') {
			return res.status(401).json({
				status: 'failure',
				statusCode: err.statusCode || 401,
				code: err.code || 'NO_ACCESS',
				message: err.message || 'No privileges to perform this action.',
			});
		}
		const statusCode = err.statusCode || 400;
		res.status(statusCode).json({
			status: 'failure',
			message: err.message || 'Failed to save mark as read.',
			error: err.toString(),
		});
	}
});

// Upsert Act Description by act name (create or update) – so status change is saved in backend Data Store
app.post('/act-descriptions/upsert', async (req, res) => {
	try {
		const { actName, description, type, state, sector, applicability, keyComplianceRequirements, dueDate, penaltyforNonCompliance, registers, status, checkbox } = req.body;
		const name = actName != null ? String(actName).trim() : '';
		if (!name) {
			return res.status(400).json({ status: 'failure', message: 'Act name is required.' });
		}
		const { catalyst } = res.locals;
		const table = catalyst.datastore().table('ActDescription');
		const zcql = catalyst.zcql();
		const statusVal = status === 'Completed' ? 'completed' : status === 'Pending' ? 'in progress' : (status || 'yet to start').toLowerCase().replace(/\s+/g, ' ');

		const rowData = {
			ActName: name,
			Description: description || null,
			Type: type || null,
			State: state || null,
			Sector: sector || null,
			Applicability: applicability || null,
			KeyComplianceRequirements: keyComplianceRequirements || null,
			DueDate: dueDate || null,
			PenaltyforNonCompliance: penaltyforNonCompliance || null,
			Registers: registers || null,
			Status: statusVal || 'yet to start',
		};
		if (checkbox !== undefined) {
			rowData.Checkbox = checkbox != null ? String(checkbox) : '';
		}

		let ROWID = null;
		const findQuery = `SELECT ROWID FROM ActDescription WHERE ActName = '${name.replace(/'/g, "''")}'`;
		const existing = await zcql.executeZCQLQuery(findQuery);
		if (existing && existing.length > 0) {
			ROWID = existing[0].ActDescription.ROWID;
		}
		if (!ROWID) {
			const allRows = await zcql.executeZCQLQuery('SELECT ROWID, ActName FROM ActDescription');
			const nameLower = name.toLowerCase();
			for (const r of allRows || []) {
				const stored = (r.ActDescription.ActName || '').trim().toLowerCase();
				if (stored === nameLower) {
					ROWID = r.ActDescription.ROWID;
					break;
				}
			}
		}
		if (ROWID) {
			rowData.Status = statusVal || 'yet to start';
			if (checkbox === undefined) {
				const existingRow = await table.getRow(ROWID);
				const existingCheckbox = existingRow && (existingRow.Checkbox || existingRow.checkbox);
				if (existingCheckbox != null) rowData.Checkbox = existingCheckbox;
			}
			await table.updateRow({ ROWID, ...rowData });
			const updated = await table.getRow(ROWID);
			return res.status(200).json({ status: 'success', data: { actDescription: updated }, updated: true });
		}

		const insertResp = await table.insertRow(rowData);
		const created = await table.getRow(insertResp.ROWID);
		res.status(200).json({ status: 'success', data: { actDescription: created }, updated: false });
	} catch (err) {
		console.error('Error upserting act description:', err);
		if (err.statusCode === 401 || err.code === 'NO_ACCESS') {
			return res.status(401).json({
				status: 'failure',
				statusCode: err.statusCode || 401,
				code: err.code || 'NO_ACCESS',
				message: err.message || 'No privileges to perform this action.'
			});
		}
		const statusCode = err.statusCode || 400;
		res.status(statusCode).json({
			status: 'failure',
			message: err.message || 'Failed to upsert act description.',
			error: err.toString()
		});
	}
});

// Delete Act Description
app.delete('/act-descriptions/:ROWID', async (req, res) => {
	try {
		const { ROWID } = req.params;
		const { catalyst } = res.locals;
		const table = catalyst.datastore().table('ActDescription');
		await table.deleteRow(ROWID);
		res.status(200).json({ status: 'success', data: { id: ROWID } });
	} catch (err) {
		console.error('Error deleting act description:', err);
		// Handle Catalyst permission errors
		if (err.statusCode === 401 || err.code === 'NO_ACCESS') {
			return res.status(401).json({ 
				status: 'failure', 
				statusCode: err.statusCode || 401,
				code: err.code || 'NO_ACCESS',
				message: err.message || 'No privileges to perform this action. Please check function permissions in catalyst-config.json.'
			});
		}
		const statusCode = err.statusCode || 500;
		res.status(statusCode).json({ 
			status: 'failure', 
			message: err.message || 'Failed to delete act description.',
			error: err.toString()
		});
	}
});

// Delete all act descriptions
app.delete('/act-descriptions', async (req, res) => {
	try {
		const { catalyst } = res.locals;
		const zcql = catalyst.zcql();
		const table = catalyst.datastore().table('ActDescription');

		const rows = await zcql.executeZCQLQuery('SELECT ROWID FROM ActDescription');
		if (!rows || rows.length === 0) {
			return res.status(200).json({ status: 'success', deletedCount: 0, message: 'No act descriptions to delete' });
		}

		let deletedCount = 0;
		for (const row of rows) {
			const rowId = row.ActDescription.ROWID;
			await table.deleteRow(rowId);
			deletedCount++;
		}

		res.status(200).json({ status: 'success', deletedCount });
	} catch (err) {
		console.error('Error deleting all act descriptions:', err);
		const statusCode = err.statusCode || 500;
		res.status(statusCode).json({ 
			status: 'failure', 
			message: err.message || 'Failed to delete act descriptions.',
			error: err.toString()
		});
	}
});

module.exports = app;
