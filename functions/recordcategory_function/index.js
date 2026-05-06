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
	res.status(200).send('recordcategory_function ready');
});

// ===== Record Categories CRUD =====
// Table: RecordCategory (columns: CategoryName, Description)

// Create Record Category
app.post('/recordcategories', async (req, res) => {
	try {
		const { categoryName, description } = req.body;
		if (!categoryName || !String(categoryName).trim()) {
			return res.status(400).send({ status: 'failure', message: 'Category name is required.' });
		}
		const { catalyst } = res.locals;
		const table = catalyst.datastore().table('RecordCategory');
		const insertResp = await table.insertRow({
			RecordCategory: categoryName,
		});
		const created = await table.getRow(insertResp.ROWID);
		res.status(200).send({ status: 'success', data: { recordCategory: created } });
	} catch (err) {
		res.status(400).send({ status: 'failure', message: err.message || 'Failed to create record category.' });
	}
});

// List Record Categories with optional pagination
app.get('/recordcategories', async (req, res) => {
	try {
		const { catalyst } = res.locals;
		const zcql = catalyst.zcql();
		const page = parseInt(req.query.page, 10) || 1;
		const perPage = parseInt(req.query.perPage, 10) || 50;
		const returnAll = !req.query.page && !req.query.perPage;
		const limitClause = returnAll ? '' : `LIMIT ${(page - 1) * perPage + 1},${perPage}`;
		const countRows = await zcql.executeZCQLQuery('SELECT COUNT(ROWID) as count FROM RecordCategory');
		const total = parseInt(countRows[0].RecordCategory.count, 10) || 0;
		const rows = await zcql.executeZCQLQuery(`SELECT ROWID, RecordCategory, CREATEDTIME, MODIFIEDTIME FROM RecordCategory ORDER BY ROWID DESC ${limitClause}`);
		const recordCategories = rows.map(r => ({
			id: r.RecordCategory.ROWID,
			categoryName: r.RecordCategory.RecordCategory,
			createdTime: r.RecordCategory.CREATEDTIME,
			modifiedTime: r.RecordCategory.MODIFIEDTIME
		}));
		res.status(200).send({ status: 'success', data: { recordCategories, total, hasMore: returnAll ? false : page * perPage < total } });
	} catch (err) {
		res.status(500).send({ status: 'failure', message: err.message || 'Failed to fetch record categories.' });
	}
});

// Update Record Category
app.put('/recordcategories/:ROWID', async (req, res) => {
	try {
		const { ROWID } = req.params;
		const { categoryName, description } = req.body;
		if (!categoryName || !String(categoryName).trim()) {
			return res.status(400).send({ status: 'failure', message: 'Category name is required.' });
		}
		const { catalyst } = res.locals;
		const table = catalyst.datastore().table('RecordCategory');
		await table.updateRow({
			ROWID,
			RecordCategory: categoryName,
		});
		const updated = await table.getRow(ROWID);
		res.status(200).send({ status: 'success', data: { recordCategory: updated } });
	} catch (err) {
		res.status(400).send({ status: 'failure', message: err.message || 'Failed to update record category.' });
	}
});

// Delete Record Category
app.delete('/recordcategories/:ROWID', async (req, res) => {
	try {
		const { ROWID } = req.params;
		const { catalyst } = res.locals;
		const table = catalyst.datastore().table('RecordCategory');
		await table.deleteRow(ROWID);
		res.status(200).send({ status: 'success', data: { id: ROWID } });
	} catch (err) {
		res.status(500).send({ status: 'failure', message: err.message || 'Failed to delete record category.' });
	}
});

module.exports = app;
