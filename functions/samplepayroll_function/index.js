'use strict';

const express = require('express');
const catalystSDK = require('zcatalyst-sdk-node');

const app = express();
const TABLE_NAME = 'SamplePayroll';
const ZCQL_MAX_ROWS = 300;
const SAMPLE_PAYROLL_SELECT_COLUMNS =
	'ROWID, EmployeeName, EmployeeID, GIDNumber, Email, DateofBirth, Paid_days, Basic, HRA, Gross, Netpay, TotalDeduction, IncomeTax, PF, VoluntaryProvidentFund, ProfessionalTax, PayrollMonth, CREATEDTIME, MODIFIEDTIME';

app.use(express.json({ limit: '50mb' }));

app.use((req, res, next) => {
	try {
		const catalyst = catalystSDK.initialize(req);
		res.locals.catalyst = catalyst;
		next();
	} catch (err) {
		res.status(500).send({ status: 'failure', message: 'Catalyst init failed' });
	}
});

app.get('/', (req, res) => {
	res.status(200).send('samplepayroll_function ready');
});

function toNullIfEmpty(value) {
	if (value == null || value === '') return null;
	const str = String(value).trim();
	return str === '' ? null : str;
}

function isValidPayrollMonth(value) {
	return typeof value === 'string' && /^\d{4}-\d{2}$/.test(value.trim());
}

function normalizePayrollMonth(value) {
	if (value == null || value === '') return '';
	const text = String(value).trim();
	const match = text.match(/^(\d{4})[-/](\d{1,2})$/);
	if (!match) return text;
	return `${match[1]}-${String(Number(match[2])).padStart(2, '0')}`;
}

function parseDatastoreJson(value) {
	let parsed = value;
	for (let i = 0; i < 3; i += 1) {
		if (parsed == null || parsed === '') return null;
		if (typeof parsed === 'object') return parsed;
		try {
			parsed = JSON.parse(String(parsed));
		} catch (_) {
			return null;
		}
	}
	return typeof parsed === 'object' ? parsed : null;
}

function pickPayrollDatastoreRow(entry) {
	if (!entry || typeof entry !== 'object') return null;
	return entry.Payroll || entry.payroll || entry;
}

function extractPayrollMonth(parsed) {
	if (!parsed || typeof parsed !== 'object') return '';
	return normalizePayrollMonth(
		parsed.payrollMonth ?? parsed.PayrollMonth ?? parsed.month ?? parsed.Month
	);
}

function extractPayrollRecords(parsed) {
	if (!parsed || typeof parsed !== 'object') return [];
	const records = parsed.records ?? parsed.Records ?? parsed.data ?? parsed.Data;
	return Array.isArray(records) ? records : [];
}

function readRowDataField(row) {
	if (!row || typeof row !== 'object') return null;
	return row.Data ?? row.data ?? row.DATA ?? null;
}

function readRowId(row) {
	if (!row || typeof row !== 'object') return null;
	return row.ROWID ?? row.rowid ?? row.RowId ?? row.id ?? null;
}

function rowHasUsablePayrollData(rawData) {
	const parsed = parseDatastoreJson(rawData);
	if (!parsed || typeof parsed !== 'object') {
		const raw = String(rawData || '');
		return /"payrollMonth"\s*:\s*"\d{4}-\d{2}"/i.test(raw);
	}
	if (extractPayrollMonth(parsed)) return true;
	return extractPayrollRecords(parsed).length > 0;
}

async function hydratePayrollDatastoreRow(table, row) {
	const base = pickPayrollDatastoreRow(row) || row;
	const rowId = readRowId(base);
	if (rowId == null) return null;

	const rawData = readRowDataField(base);
	if (rowHasUsablePayrollData(rawData)) {
		return base;
	}

	try {
		const full = await table.getRow(rowId);
		return full || base;
	} catch (err) {
		console.warn(`Payroll getRow(${rowId}) failed:`, err.message || err);
		return base;
	}
}

function rowMatchesPayrollMonth(rawData, month) {
	const parsed = parseDatastoreJson(rawData);
	if (parsed && typeof parsed === 'object') {
		const rowMonth = extractPayrollMonth(parsed);
		if (rowMonth === month) {
			return { match: true, parsed };
		}
	}

	const raw = String(rawData || '');
	const monthPatterns = [
		`"payrollMonth":"${month}"`,
		`"payrollMonth": "${month}"`,
		`"PayrollMonth":"${month}"`,
		`"PayrollMonth": "${month}"`,
		`"month":"${month}"`,
		`"month": "${month}"`,
	];
	if (monthPatterns.some((pattern) => raw.includes(pattern))) {
		const recovered = parsed && typeof parsed === 'object' ? parsed : parseDatastoreJson(rawData);
		if (recovered && typeof recovered === 'object') {
			return { match: true, parsed: recovered };
		}
	}

	return { match: false, parsed: null };
}

function stashPayrollRow(map, row) {
	const base = pickPayrollDatastoreRow(row) || row;
	const rowId = readRowId(base);
	if (rowId == null) return;
	map.set(String(rowId), base);
}

async function fetchAllPayrollTableRows(catalyst) {
	const table = catalyst.datastore().table('Payroll');
	const byId = new Map();

	try {
		if (typeof table.getIterableRows === 'function') {
			for await (const row of table.getIterableRows()) {
				stashPayrollRow(byId, row);
			}
		}
	} catch (err) {
		console.warn('Payroll getIterableRows failed:', err.message || err);
	}

	if (byId.size === 0) {
		try {
			const rows = await table.getAllRows();
			if (Array.isArray(rows)) {
				for (const row of rows) {
					stashPayrollRow(byId, row);
				}
			}
		} catch (err) {
			console.warn('Payroll getAllRows failed:', err.message || err);
		}
	}

	if (byId.size === 0) {
		try {
			const zcql = catalyst.zcql();
			const rows = await zcql.executeZCQLQuery(
				'SELECT ROWID, Data, CREATEDTIME, MODIFIEDTIME FROM Payroll'
			);
			const list = Array.isArray(rows) ? rows : [];
			for (const entry of list) {
				stashPayrollRow(byId, entry);
			}
		} catch (err) {
			console.warn('Payroll ZCQL select failed:', err.message || err);
		}
	}

	const hydrated = [];
	for (const row of byId.values()) {
		const full = await hydratePayrollDatastoreRow(table, row);
		if (full) hydrated.push(full);
	}
	return hydrated;
}

async function tryLoadSnapshotByMonthLike(catalyst, month) {
	const table = catalyst.datastore().table('Payroll');
	const zcql = catalyst.zcql();
	const likePatterns = [
		`%"payrollMonth":"${month}"%`,
		`%"payrollMonth": "${month}"%`,
		`%"PayrollMonth":"${month}"%`,
		`%"PayrollMonth": "${month}"%`,
	];

	for (const pattern of likePatterns) {
		try {
			const safe = pattern.replace(/'/g, "''");
			const rows = await zcql.executeZCQLQuery(
				`SELECT ROWID, Data, CREATEDTIME, MODIFIEDTIME FROM Payroll WHERE Data LIKE '${safe}'`
			);
			const list = Array.isArray(rows) ? rows : [];
			let best = null;
			for (const entry of list) {
				const row = await hydratePayrollDatastoreRow(
					table,
					pickPayrollDatastoreRow(entry) || entry
				);
				if (!row) continue;
				const rawData = readRowDataField(row);
				const match = rowMatchesPayrollMonth(rawData, month);
				if (!match.match || !match.parsed) continue;
				const modified = String(row.MODIFIEDTIME || row.CREATEDTIME || '');
				if (!best || modified > String(best.row.MODIFIEDTIME || best.row.CREATEDTIME || '')) {
					best = { row, parsed: match.parsed };
				}
			}
			if (best) return best;
		} catch (err) {
			console.warn('Payroll ZCQL month LIKE failed:', err.message || err);
		}
	}
	return null;
}

function collectAvailablePayrollMonths(rows) {
	const months = new Set();
	for (const row of rows) {
		const rawData = readRowDataField(row);
		const parsed = parseDatastoreJson(rawData);
		const month = extractPayrollMonth(parsed);
		if (month) months.add(month);
		if (!month) {
			const raw = String(rawData || '');
			const match = raw.match(/"payrollMonth"\s*:\s*"(\d{4}-\d{2})"/i);
			if (match) months.add(match[1]);
		}
	}
	return [...months].sort();
}

async function loadPayrollTableSnapshot(catalyst, payrollMonth) {
	const month = normalizePayrollMonth(payrollMonth);
	if (!isValidPayrollMonth(month)) {
		return { found: false, reason: 'invalid_month' };
	}

	let best = await tryLoadSnapshotByMonthLike(catalyst, month);
	const list = await fetchAllPayrollTableRows(catalyst);

	if (!best) {
		for (const entry of list) {
			const row = pickPayrollDatastoreRow(entry) || entry;
			if (!row || readRowId(row) == null) continue;

			const rawData = readRowDataField(row);
			const match = rowMatchesPayrollMonth(rawData, month);
			if (!match.match || !match.parsed) continue;

			const modified = String(row.MODIFIEDTIME || row.CREATEDTIME || '');
			if (!best || modified > String(best.row.MODIFIEDTIME || best.row.CREATEDTIME || '')) {
				best = { row, parsed: match.parsed, rowMonth: month };
			}
		}
	}

	if (!best) {
		return {
			found: false,
			reason: 'not_found',
			payrollMonth: month,
			availableMonths: collectAvailablePayrollMonths(list),
			rowCount: list.length,
		};
	}

	const records = extractPayrollRecords(best.parsed);
	return {
		found: true,
		payrollMonth: month,
		rowId: readRowId(best.row),
		records,
		meta: {
			hasBreakdown: best.parsed.hasBreakdown === true,
			loadedCount: best.parsed.loadedCount,
			totalExpected: best.parsed.totalExpected,
			breakdownComplete: best.parsed.breakdownComplete === true,
			savedAt: best.parsed.savedAt || null,
			payDate: best.parsed.runMeta?.pay_date || best.parsed.runMeta?.payDate || null,
			payroll_run_id: best.parsed.runMeta?.payroll_run_id || null,
			organizationId: best.parsed.organizationId || null,
		},
	};
}

function mapRow(row) {
	return {
		id: row.ROWID,
		employeeName: row.EmployeeName || '',
		employeeId: row.EmployeeID || '',
		gidNumber: row.GIDNumber || '',
		email: row.Email || '',
		dateofBirth: row.DateofBirth || '',
		paidDays: row.Paid_days || '',
		basic: row.Basic || '',
		hra: row.HRA || '',
		gross: row.Gross || '',
		netpay: row.Netpay || '',
		totalDeduction: row.TotalDeduction || '',
		incomeTax: row.IncomeTax || '',
		pf: row.PF || '',
		voluntaryProvidentFund: row.VoluntaryProvidentFund || '',
		professionalTax: row.ProfessionalTax || '',
		payrollMonth: row.PayrollMonth || '',
		createdTime: row.CREATEDTIME,
		modifiedTime: row.MODIFIEDTIME,
	};
}

async function fetchSamplePayrollRowsPaged(
	catalyst,
	{ whereClause = '', orderClause = 'ORDER BY ROWID ASC' } = {}
) {
	const zcql = catalyst.zcql();
	const collected = [];
	let offset = 1;

	while (true) {
		const rows = await zcql.executeZCQLQuery(
			`SELECT ${SAMPLE_PAYROLL_SELECT_COLUMNS} FROM ${TABLE_NAME} ${whereClause} ${orderClause} LIMIT ${offset}, ${ZCQL_MAX_ROWS}`
		);
		const batch = (Array.isArray(rows) ? rows : []).map((entry) =>
			mapRow(entry[TABLE_NAME] || entry)
		);
		collected.push(...batch);
		if (batch.length < ZCQL_MAX_ROWS) break;
		offset += ZCQL_MAX_ROWS;
	}

	return collected;
}

function pickField(record, keys) {
	for (const key of keys) {
		const value = record[key];
		if (value != null && String(value).trim() !== '') return String(value).trim();
	}
	return null;
}

function parseBenefitsList(record) {
	if (!record || typeof record !== 'object') return [];
	const raw =
		record.benefits ??
		record.Benefits ??
		record.benefit ??
		record.employee_benefits ??
		null;
	if (Array.isArray(raw)) return raw;
	if (typeof raw === 'string' && raw.trim()) {
		try {
			const parsed = JSON.parse(raw);
			return Array.isArray(parsed) ? parsed : [];
		} catch (_) {
			return [];
		}
	}
	return [];
}

function benefitName(item) {
	return String(item?.name ?? item?.benefit_name ?? item?.label ?? '')
		.trim()
		.toLowerCase();
}

function benefitPlan(item) {
	return String(item?.plan ?? item?.type ?? item?.benefit_type ?? '')
		.trim()
		.toLowerCase()
		.replace(/[^a-z0-9]+/g, '_');
}

function benefitAmount(item) {
	if (!item || typeof item !== 'object') return null;
	const value = item.amount ?? item.value ?? item.component_amount;
	if (value == null || value === '') return null;
	const num = Number(String(value).replace(/,/g, '').trim());
	return Number.isFinite(num) ? String(num) : null;
}

function pickBenefitAmount(record, matcher) {
	const list = parseBenefitsList(record);
	for (let i = 0; i < list.length; i += 1) {
		const item = list[i];
		if (!item || typeof item !== 'object') continue;
		if (matcher(benefitPlan(item), benefitName(item))) {
			const amount = benefitAmount(item);
			if (amount != null) return amount;
		}
	}
	return null;
}

function mapPayrollRecordToSampleRow(record, payrollMonth) {
	const gross = pickField(record, ['gross_pay', 'total_earnings', 'gross', 'Gross']);
	const totalDeduction = pickField(record, [
		'total_deductions',
		'totalDeduction',
		'TotalDeduction',
		'total_employee_deductions',
	]);
	let netpay = pickField(record, [
		'net_pay',
		'netPay',
		'Netpay',
		'netpay',
		'monthly_salary',
		'net_salary',
		'net_wages',
	]);
	if (!netpay && gross && totalDeduction) {
		const netNum = Number(gross) - Number(totalDeduction);
		if (Number.isFinite(netNum) && netNum >= 0) netpay = String(netNum);
	}
	const pfFromBenefits = pickBenefitAmount(
		record,
		(plan, name) =>
			plan === 'epf_contribution' ||
			plan === 'epf' ||
			name === 'epf contribution' ||
			(name.includes('epf') &&
				!name.includes('voluntary') &&
				!name.includes('employer') &&
				!name.includes('admin') &&
				!name.includes('edli'))
	);
	const vpfFromBenefits = pickBenefitAmount(
		record,
		(plan, name) =>
			plan === 'vpf' ||
			plan === 'voluntary_provident_fund' ||
			name.includes('voluntary provident') ||
			name === 'vpf'
	);
	return {
		EmployeeName: pickField(record, ['employee_name', 'full_name', 'employeeName', 'name']),
		EmployeeID: pickField(record, ['employee_id', 'employeeId', 'EmployeeID']),
		GIDNumber: pickField(record, ['employee_number', 'gidNumber', 'GIDNumber', 'gid_number']),
		Email: pickField(record, [
			'email',
			'work_email',
			'work_mail',
			'personal_email',
			'mail_id',
			'EmailID',
			'Email',
			'Work_Email',
			'Work Email',
			'Email ID',
		]),
		DateofBirth: pickField(record, [
			'date_of_birth',
			'Date_of_birth',
			'DateofBirth',
			'dateofBirth',
			'dateOfBirth',
			'DOB',
			'dob',
		]),
		Paid_days: pickField(record, ['paid_days', 'paidDays', 'Paid_days', 'paid_days_in_month']),
		Basic: pickField(record, ['basic', 'earned_basic', 'Basic', 'basic_pay']),
		HRA: pickField(record, ['hra', 'hra_fbp', 'HRA', 'house_rent_allowance']),
		Gross: gross,
		Netpay: netpay,
		TotalDeduction: totalDeduction,
		IncomeTax: pickField(record, [
			'income_tax',
			'IncomeTax',
			'incomeTax',
			'Income Tax',
			'tds',
			'TDS',
			'tax_deducted_at_source',
		]),
		PF:
			pickField(record, [
				'epf_contribution',
				'EPF Contribution',
				'PF',
				'pf',
				'employer_pf',
				'employer_epf',
			]) || pfFromBenefits,
		VoluntaryProvidentFund:
			pickField(record, [
				'voluntary_provident_fund',
				'VoluntaryProvidentFund',
				'Voluntary Provident Fund',
				'voluntaryProvidentFund',
				'vpf',
				'VPF',
			]) || vpfFromBenefits,
		ProfessionalTax: pickField(record, [
			'professional_tax',
			'ProfessionalTax',
			'professionalTax',
			'Professional Tax',
			'pt',
			'PT',
		]),
		PayrollMonth: payrollMonth,
	};
}

async function deleteSamplePayrollRowsForMonth(catalyst, payrollMonth) {
	const table = catalyst.datastore().table(TABLE_NAME);
	const zcql = catalyst.zcql();
	const rows = await zcql.executeZCQLQuery(
		`SELECT ROWID FROM ${TABLE_NAME} WHERE PayrollMonth = '${String(payrollMonth).trim()}'`
	);
	const ids = (Array.isArray(rows) ? rows : [])
		.map((entry) => {
			const cell = entry[TABLE_NAME] || entry;
			return cell?.ROWID ?? cell?.rowid ?? null;
		})
		.filter((id) => id != null);

	if (ids.length === 0) return 0;

	const deleteChunkSize = 200;
	for (let i = 0; i < ids.length; i += deleteChunkSize) {
		const chunk = ids.slice(i, i + deleteChunkSize);
		if (typeof table.deleteRows === 'function') {
			await table.deleteRows(chunk);
		} else {
			for (const id of chunk) {
				await table.deleteRow(id);
			}
		}
	}
	return ids.length;
}

async function appendMonthRecordsToSamplePayroll(
	catalyst,
	payrollMonth,
	records,
	{ replaceExisting = false } = {}
) {
	const month = normalizePayrollMonth(payrollMonth);
	if (!isValidPayrollMonth(month)) {
		return { synced: false, reason: 'invalid_month' };
	}

	const list = Array.isArray(records)
		? records.filter((row) => row && typeof row === 'object' && row.fetch_error !== true)
		: [];
	if (list.length === 0) {
		return { synced: false, reason: 'empty_records' };
	}

	let deleted = 0;
	if (replaceExisting) {
		deleted = await deleteSamplePayrollRowsForMonth(catalyst, month);
	}

	const table = catalyst.datastore().table(TABLE_NAME);
	const rowData = list
		.map((record) => mapPayrollRecordToSampleRow(record, month))
		.filter((row) => row.EmployeeName || row.EmployeeID);

	const chunkSize = 100;
	let inserted = 0;
	for (let i = 0; i < rowData.length; i += chunkSize) {
		const chunk = rowData.slice(i, i + chunkSize);
		if (chunk.length === 0) continue;
		await table.insertRows(chunk);
		inserted += chunk.length;
	}

	return {
		synced: true,
		payrollMonth: month,
		inserted,
		deleted,
		total: inserted,
	};
}

async function syncMonthRecordsToSamplePayroll(catalyst, payrollMonth, records) {
	return appendMonthRecordsToSamplePayroll(catalyst, payrollMonth, records, {
		replaceExisting: true,
	});
}

function buildRowData(body) {
	const {
		employeeName,
		employeeId,
		gidNumber,
		email,
		dateofBirth,
		paidDays,
		basic,
		hra,
		gross,
		netpay,
		totalDeduction,
		incomeTax,
		pf,
		voluntaryProvidentFund,
		professionalTax,
		payrollMonth,
	} = body;

	return {
		EmployeeName: toNullIfEmpty(employeeName),
		EmployeeID: toNullIfEmpty(employeeId),
		GIDNumber: toNullIfEmpty(gidNumber),
		Email: toNullIfEmpty(email),
		DateofBirth: toNullIfEmpty(dateofBirth),
		Paid_days: toNullIfEmpty(paidDays),
		Basic: toNullIfEmpty(basic),
		HRA: toNullIfEmpty(hra),
		Gross: toNullIfEmpty(gross),
		Netpay: toNullIfEmpty(netpay),
		TotalDeduction: toNullIfEmpty(totalDeduction),
		IncomeTax: toNullIfEmpty(incomeTax),
		PF: toNullIfEmpty(pf),
		VoluntaryProvidentFund: toNullIfEmpty(voluntaryProvidentFund),
		ProfessionalTax: toNullIfEmpty(professionalTax),
		PayrollMonth: isValidPayrollMonth(payrollMonth) ? payrollMonth.trim() : null,
	};
}

// Sync month-wise payroll records into SamplePayroll table (replace existing month rows)
app.post('/samplepayroll/sync-month', async (req, res) => {
	try {
		const body = req.body && typeof req.body === 'object' ? req.body : {};
		const payrollMonth = body.payrollMonth || body.payroll_month;
		const records = Array.isArray(body.records) ? body.records : [];
		const replaceExisting = body.replaceExisting !== false && body.replaceExisting !== '0';
		const finalize = body.finalize === true || body.finalize === '1';
		const { catalyst } = res.locals;
		const result = await appendMonthRecordsToSamplePayroll(catalyst, payrollMonth, records, {
			replaceExisting,
		});

		if (!result.synced) {
			const status = result.reason === 'invalid_month' ? 400 : 400;
			return res.status(status).json({
				status: 'failure',
				message:
					result.reason === 'invalid_month'
						? 'payrollMonth must be YYYY-MM'
						: 'No payroll records to store in SamplePayroll table.',
				reason: result.reason || 'unknown',
			});
		}

		if (!finalize) {
			return res.status(200).json({
				status: 'success',
				message: `Stored ${result.inserted} employee record(s) in this batch.`,
				data: {
					payrollMonth: result.payrollMonth,
					sync: result,
					finalized: false,
				},
			});
		}

		const month = normalizePayrollMonth(payrollMonth);
		const stored = await fetchSamplePayrollRowsPaged(catalyst, {
			whereClause: `WHERE PayrollMonth = '${month}'`,
			orderClause: 'ORDER BY ROWID ASC',
		});

		res.status(200).json({
			status: 'success',
			message: `Stored ${stored.length} employee record(s) in SamplePayroll table for ${month}.`,
			data: {
				payrollMonth: month,
				records: stored,
				sync: result,
				finalized: true,
			},
		});
	} catch (err) {
		console.error('Error syncing SamplePayroll month:', err);
		res.status(500).json({
			status: 'failure',
			message: err.message || 'Failed to store payroll data in SamplePayroll table.',
		});
	}
});

// Create SamplePayroll record
app.post('/samplepayroll', async (req, res) => {
	try {
		const { employeeName } = req.body;
		if (!employeeName || !String(employeeName).trim()) {
			return res.status(400).json({ status: 'failure', message: 'Employee name is required.' });
		}

		const { catalyst } = res.locals;
		const table = catalyst.datastore().table(TABLE_NAME);
		const rowData = buildRowData(req.body);
		const insertResp = await table.insertRow(rowData);
		const created = await table.getRow(insertResp.ROWID);

		res.status(200).json({ status: 'success', data: { record: mapRow(created) } });
	} catch (err) {
		console.error('Error creating sample payroll record:', err);
		if (err.statusCode === 401 || err.code === 'NO_ACCESS') {
			return res.status(401).json({
				status: 'failure',
				statusCode: err.statusCode || 401,
				code: err.code || 'NO_ACCESS',
				message: err.message || 'No privileges to perform this action.',
			});
		}
		res.status(err.statusCode || 400).json({
			status: 'failure',
			message: err.message || 'Failed to create sample payroll record.',
		});
	}
});

// Read month-wise payroll snapshot from Payroll table (JSON Data column)
app.get('/samplepayroll/payroll-table', async (req, res) => {
	try {
		const payrollMonth = req.query.payroll_month || req.query.payrollMonth;
		const { catalyst } = res.locals;
		const result = await loadPayrollTableSnapshot(catalyst, payrollMonth);

		if (!result.found) {
			const status = result.reason === 'invalid_month' ? 400 : 404;
			const available = Array.isArray(result.availableMonths) ? result.availableMonths : [];
			let message =
				result.reason === 'invalid_month'
					? 'payroll_month must be YYYY-MM'
					: `No payroll snapshot found for ${normalizePayrollMonth(payrollMonth) || payrollMonth}`;
			if (available.length > 0) {
				message += `. Available months in Payroll table: ${available.join(', ')}`;
			} else if (result.rowCount === 0) {
				message += '. Payroll table appears empty — fetch and save data on the Payroll page first.';
			}
			return res.status(status).json({
				status: 'failure',
				message,
				reason: result.reason || 'not_found',
				availableMonths: available,
				rowCount: result.rowCount ?? 0,
			});
		}

		res.status(200).json({
			status: 'success',
			data: {
				payrollMonth: result.payrollMonth,
				rowId: result.rowId,
				records: result.records,
				meta: result.meta,
				total: result.records.length,
			},
		});
	} catch (err) {
		console.error('Error loading Payroll table snapshot:', err);
		res.status(500).json({
			status: 'failure',
			message: err.message || 'Failed to load payroll data from Payroll table.',
		});
	}
});

// List SamplePayroll records
app.get('/samplepayroll', async (req, res) => {
	try {
		const { catalyst } = res.locals;
		const zcql = catalyst.zcql();
		const page = parseInt(req.query.page, 10) || 1;
		const perPage = parseInt(req.query.perPage, 10) || 50;
		const search = req.query.search;
		const payrollMonth = req.query.payroll_month || req.query.payrollMonth;
		const returnAll =
			req.query.fetch_all === '1' ||
			req.query.fetch_all === 'true' ||
			(!req.query.page && !req.query.perPage);
		const limitClause = returnAll ? '' : `LIMIT ${(page - 1) * perPage + 1},${perPage}`;

		const whereParts = [];
		if (isValidPayrollMonth(payrollMonth)) {
			whereParts.push(`PayrollMonth = '${String(payrollMonth).trim()}'`);
		}
		if (search) {
			const safe = String(search).replace(/'/g, "''");
			whereParts.push(
				`(EmployeeName LIKE '%${safe}%' OR EmployeeID LIKE '%${safe}%' OR Email LIKE '%${safe}%')`
			);
		}
		const whereClause = whereParts.length ? `WHERE ${whereParts.join(' AND ')}` : '';

		const countRows = await zcql.executeZCQLQuery(
			`SELECT COUNT(ROWID) as count FROM ${TABLE_NAME} ${whereClause}`
		);
		const total = parseInt(countRows[0][TABLE_NAME].count, 10) || 0;

		let records = [];
		if (returnAll) {
			records = await fetchSamplePayrollRowsPaged(catalyst, {
				whereClause,
				orderClause: 'ORDER BY ROWID DESC',
			});
		} else {
			const selectQuery = `SELECT ${SAMPLE_PAYROLL_SELECT_COLUMNS} FROM ${TABLE_NAME} ${whereClause} ORDER BY ROWID DESC ${limitClause}`;
			const rows = await zcql.executeZCQLQuery(selectQuery);
			records = rows.map((r) => mapRow(r[TABLE_NAME]));
		}

		res.status(200).json({
			status: 'success',
			data: {
				records,
				total,
				payrollMonth: isValidPayrollMonth(payrollMonth) ? payrollMonth.trim() : null,
				hasMore: returnAll ? false : page * perPage < total,
			},
		});
	} catch (err) {
		console.error('Error fetching sample payroll records:', err);
		res.status(500).json({ status: 'failure', message: err.message || 'Failed to fetch sample payroll records.' });
	}
});

// Get single record
app.get('/samplepayroll/:ROWID', async (req, res) => {
	try {
		const { ROWID } = req.params;
		const { catalyst } = res.locals;
		const table = catalyst.datastore().table(TABLE_NAME);
		const row = await table.getRow(ROWID);
		res.status(200).json({ status: 'success', data: { record: mapRow(row) } });
	} catch (err) {
		console.error('Error fetching sample payroll record:', err);
		res.status(404).json({ status: 'failure', message: err.message || 'Record not found.' });
	}
});

// Update record
app.put('/samplepayroll/:ROWID', async (req, res) => {
	try {
		const { ROWID } = req.params;
		const { employeeName } = req.body;
		if (!employeeName || !String(employeeName).trim()) {
			return res.status(400).json({ status: 'failure', message: 'Employee name is required.' });
		}

		const { catalyst } = res.locals;
		const table = catalyst.datastore().table(TABLE_NAME);
		await table.updateRow({ ROWID, ...buildRowData(req.body) });
		const updated = await table.getRow(ROWID);
		res.status(200).json({ status: 'success', data: { record: mapRow(updated) } });
	} catch (err) {
		console.error('Error updating sample payroll record:', err);
		if (err.statusCode === 401 || err.code === 'NO_ACCESS') {
			return res.status(401).json({
				status: 'failure',
				statusCode: err.statusCode || 401,
				code: err.code || 'NO_ACCESS',
				message: err.message || 'No privileges to perform this action.',
			});
		}
		res.status(err.statusCode || 400).json({
			status: 'failure',
			message: err.message || 'Failed to update sample payroll record.',
		});
	}
});

// Delete record
app.delete('/samplepayroll/:ROWID', async (req, res) => {
	try {
		const { ROWID } = req.params;
		const { catalyst } = res.locals;
		const table = catalyst.datastore().table(TABLE_NAME);
		await table.deleteRow(ROWID);
		res.status(200).json({ status: 'success', data: { id: ROWID } });
	} catch (err) {
		console.error('Error deleting sample payroll record:', err);
		if (err.statusCode === 401 || err.code === 'NO_ACCESS') {
			return res.status(401).json({
				status: 'failure',
				statusCode: err.statusCode || 401,
				code: err.code || 'NO_ACCESS',
				message: err.message || 'No privileges to perform this action.',
			});
		}
		res.status(err.statusCode || 500).json({
			status: 'failure',
			message: err.message || 'Failed to delete sample payroll record.',
		});
	}
});

module.exports = app;
