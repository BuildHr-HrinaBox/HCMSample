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
	res.status(200).send('employeedata_function ready');
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

// ===== Employee Data CRUD =====
// Table: Employee (columns: Nameoftheemployee, Sex, FathersName, Designation, EmployeeIdentificationNo, DateofBirth, DateofJoining, PresentAddress, Permanentaddress, AadhaarNo, MobileNumber, EmailID, DateofExit)

// Helper function to convert empty strings to null
function toNullIfEmpty(value) {
	if (value == null || value === '') return null;
	const str = String(value).trim();
	return str === '' ? null : str;
}

// Create Employee
app.post('/employee', async (req, res) => {
	try {
		console.log('POST /employee - Request body:', req.body);
		
		const { 
			nameoftheemployee, 
			sex, 
			fathersname, 
			designation, 
			employeeidentificationno, 
			dateofbirth, 
			dateofjoining, 
			presentaddress, 
			permanentaddress, 
			aadhaarno, 
			mobilenumber, 
			emailid, 
			dateofexit 
		} = req.body;
		
		if (!nameoftheemployee || !String(nameoftheemployee).trim()) {
			return res.status(400).json({ status: 'failure', message: 'Employee name is required.' });
		}
		
		const { catalyst } = res.locals;
		const table = catalyst.datastore().table('Employee');
		
		const rowData = {
			Nameoftheemployee: String(nameoftheemployee).trim(),
			Sex: toNullIfEmpty(sex),
			FathersName: toNullIfEmpty(fathersname),
			Designation: toNullIfEmpty(designation),
			EmployeeIdentificationNo: toNullIfEmpty(employeeidentificationno),
			DateofBirth: toNullIfEmpty(dateofbirth),
			DateofJoining: toNullIfEmpty(dateofjoining),
			PresentAddress: toNullIfEmpty(presentaddress),
			Permanentaddress: toNullIfEmpty(permanentaddress),
			AadhaarNo: toNullIfEmpty(aadhaarno),
			MobileNumber: toNullIfEmpty(mobilenumber),
			EmailID: toNullIfEmpty(emailid),
			DateofExit: toNullIfEmpty(dateofexit),
		};
		
		console.log('Inserting row data:', rowData);
		
		const insertResp = await table.insertRow(rowData);
		console.log('Insert response:', insertResp);
		
		const created = await table.getRow(insertResp.ROWID);
		console.log('Created employee:', created);
		
		res.status(200).json({ status: 'success', data: { employee: created } });
	} catch (err) {
		console.error('Error creating employee:', err);
		console.error('Error stack:', err.stack);
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
			message: err.message || 'Failed to create employee.',
			error: err.toString()
		});
	}
});

// List Employees with optional pagination and filtering
app.get('/employee', async (req, res) => {
	try {
		const { catalyst } = res.locals;
		const zcql = catalyst.zcql();
		const page = parseInt(req.query.page, 10) || 1;
		const perPage = parseInt(req.query.perPage, 10) || 50;
		const search = req.query.search;
		const returnAll = !req.query.page && !req.query.perPage;
		const limitClause = returnAll ? '' : `LIMIT ${(page - 1) * perPage + 1},${perPage}`;
		
		let whereClause = '';
		if (search) {
			whereClause = `WHERE Nameoftheemployee LIKE '%${search}%' OR EmployeeIdentificationNo LIKE '%${search}%' OR EmailID LIKE '%${search}%' OR MobileNumber LIKE '%${search}%'`;
		}
		
		const countQuery = `SELECT COUNT(ROWID) as count FROM Employee ${whereClause}`;
		const countRows = await zcql.executeZCQLQuery(countQuery);
		const total = parseInt(countRows[0].Employee.count, 10) || 0;
		
		const selectQuery = `SELECT ROWID, Nameoftheemployee, Sex, FathersName, Designation, EmployeeIdentificationNo, DateofBirth, DateofJoining, PresentAddress, Permanentaddress, AadhaarNo, MobileNumber, EmailID, DateofExit, CREATEDTIME, MODIFIEDTIME FROM Employee ${whereClause} ORDER BY ROWID DESC ${limitClause}`;
		const rows = await zcql.executeZCQLQuery(selectQuery);
		
		const employees = rows.map(r => ({
			id: r.Employee.ROWID,
			nameoftheemployee: r.Employee.Nameoftheemployee,
			sex: r.Employee.Sex,
			fathersname: r.Employee.FathersName,
			designation: r.Employee.Designation,
			employeeidentificationno: r.Employee.EmployeeIdentificationNo,
			dateofbirth: r.Employee.DateofBirth,
			dateofjoining: r.Employee.DateofJoining,
			presentaddress: r.Employee.PresentAddress,
			permanentaddress: r.Employee.Permanentaddress,
			aadhaarno: r.Employee.AadhaarNo,
			mobilenumber: r.Employee.MobileNumber,
			emailid: r.Employee.EmailID,
			dateofexit: r.Employee.DateofExit,
			createdTime: r.Employee.CREATEDTIME,
			modifiedTime: r.Employee.MODIFIEDTIME
		}));
		
		res.status(200).json({ 
			status: 'success', 
			data: { 
				employees, 
				total, 
				hasMore: returnAll ? false : page * perPage < total 
			} 
		});
	} catch (err) {
		console.error('Error fetching employees:', err);
		res.status(500).json({ status: 'failure', message: err.message || 'Failed to fetch employees.' });
	}
});

// Get single Employee
app.get('/employee/:ROWID', async (req, res) => {
	try {
		const { ROWID } = req.params;
		const { catalyst } = res.locals;
		const table = catalyst.datastore().table('Employee');
		const employee = await table.getRow(ROWID);
		res.status(200).json({ status: 'success', data: { employee } });
	} catch (err) {
		console.error('Error fetching employee:', err);
		res.status(404).json({ status: 'failure', message: err.message || 'Employee not found.' });
	}
});

// Update Employee
app.put('/employee/:ROWID', async (req, res) => {
	try {
		const { ROWID } = req.params;
		const { 
			nameoftheemployee, 
			sex, 
			fathersname, 
			designation, 
			employeeidentificationno, 
			dateofbirth, 
			dateofjoining, 
			presentaddress, 
			permanentaddress, 
			aadhaarno, 
			mobilenumber, 
			emailid, 
			dateofexit 
		} = req.body;
		
		if (!nameoftheemployee || !String(nameoftheemployee).trim()) {
			return res.status(400).json({ status: 'failure', message: 'Employee name is required.' });
		}
		
		const { catalyst } = res.locals;
		const table = catalyst.datastore().table('Employee');
		await table.updateRow({
			ROWID,
			Nameoftheemployee: nameoftheemployee,
			Sex: sex || null,
			FathersName: fathersname || null,
			Designation: designation || null,
			EmployeeIdentificationNo: employeeidentificationno || null,
			DateofBirth: dateofbirth || null,
			DateofJoining: dateofjoining || null,
			PresentAddress: presentaddress || null,
			Permanentaddress: permanentaddress || null,
			AadhaarNo: aadhaarno || null,
			MobileNumber: mobilenumber || null,
			EmailID: emailid || null,
			DateofExit: dateofexit || null,
		});
		const updated = await table.getRow(ROWID);
		res.status(200).json({ status: 'success', data: { employee: updated } });
	} catch (err) {
		console.error('Error updating employee:', err);
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
			message: err.message || 'Failed to update employee.',
			error: err.toString()
		});
	}
});

// Delete Employee
app.delete('/employee/:ROWID', async (req, res) => {
	try {
		const { ROWID } = req.params;
		const { catalyst } = res.locals;
		const table = catalyst.datastore().table('Employee');
		await table.deleteRow(ROWID);
		res.status(200).json({ status: 'success', data: { id: ROWID } });
	} catch (err) {
		console.error('Error deleting employee:', err);
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
			message: err.message || 'Failed to delete employee.',
			error: err.toString()
		});
	}
});

module.exports = app;
