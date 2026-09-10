'use strict';

const axios = require('axios');
const express = require('express');
const catalystSDK = require('zcatalyst-sdk-node');
const {
  flattenPayrollEarningColumns,
  getEarningsArray,
  mergePayrollRunEmployeePayload,
  unwrapSalaryEmployeePayload,
} = require('./payrollEarnings');

const PAYROLL_TABLE = 'Payroll';

/**
 * Zoho Payroll (India DC) proxy — same shape as `peopledata_function` for the app.
 *
 * Query:
 *   - organization_id (required) — Zoho Payroll organisation ID
 *   - employee_id (optional) — fetch merged employee + salary for one employee
 *   - all_salaries=1 (optional) — list employees and attach each employee's salary (for Statutory autofill)
 *   - salary_offset / salary_limit (optional) — process a slice of employees per request (avoids HTTP 408)
 *   - list_employees=1 (optional) — fast employee list without per-employee salary calls
 *   - payroll_month_data=1 (optional) — pay-run employee summary for payroll_month=YYYY-MM (or year + month)
 *   - include_earnings_detail=0 (optional) — skip per-employee pay-run detail (faster; no Basic/HRA)
 *   - payrun_employee_detail=1 (optional) — earnings/deductions for one employee (payroll_run_id + employee_id)
 *     or a batch (payroll_run_id + employee_ids=id1,id2,... max 25) — returns an array when employee_ids is used
 *
 * OAuth (same pattern as peopledata_function / leavedata_function):
 *   - ZOHO_PAYROLL_ACCESS_TOKEN — optional direct bearer (refreshed hourly in Zoho)
 *   - ZOHO_PAYROLL_REFRESH_TOKEN — Payroll-scoped refresh token (not People token)
 *   - ZOHO_PAYROLL_CLIENT_ID / ZOHO_PAYROLL_CLIENT_SECRET
 * If Catalyst env ZOHO_PAYROLL_REFRESH_TOKEN is wrong, remove it so embedded defaults apply.
 * Optional:
 *   - ZOHO_PAYROLL_ACCOUNTS_URL (default https://accounts.zoho.in/oauth/v2/token)
 *   - ZOHO_PAYROLL_API_BASE (default https://www.zohoapis.in/payroll/v1)
 *   - ZOHO_PAYROLL_SALARY_CONCURRENCY (default 5 — fetch salary rows faster for payroll page)
 *   - ZOHO_PAYROLL_SALARY_DELAY_MS (default 0 — no extra pause unless overridden)
 *   - ZOHO_PAYROLL_HTTP_RETRIES (default 5 — retries on HTTP 429 for token and GETs)
 *   - ZOHO_PAYROLL_LIST_PAGE_DELAY_MS (default 120 — pause between paginated employee list requests)
 *
 * @param {IncomingMessage} req
 * @param {ServerResponse} res
 */
/** Catalyst / proxies may put query on `req.query` or only on `originalUrl`; `req.url` can be path-only. */
function readQueryParam(req, name) {
  const q = req && req.query;
  if (q && typeof q === 'object' && q[name] != null && String(q[name]).trim() !== '') {
    return String(q[name]).trim();
  }
  const raw = String(req.originalUrl || req.url || '');
  const qMark = raw.indexOf('?');
  if (qMark >= 0) {
    try {
      const v = new URLSearchParams(raw.slice(qMark + 1)).get(name);
      if (v && String(v).trim()) return String(v).trim();
    } catch (_) {
      /* ignore */
    }
  }
  try {
    const v = new URL(raw, 'http://localhost').searchParams.get(name);
    return v && String(v).trim() ? String(v).trim() : '';
  } catch (_) {
    return '';
  }
}

const DEFAULT_PAYROLL_ORGANIZATION_ID = '60065031807';
/** Wrong org IDs seen in Catalyst env / old builds — remap before calling Zoho. */
const LEGACY_WRONG_ORG_IDS = new Set(['60006183023']);

function resolveOrganizationId(req) {
  const fromEnv = envOr('ZOHO_PAYROLL_ORGANIZATION_ID', '');
  if (fromEnv) return fromEnv;
  const fromQuery = readQueryParam(req, 'organization_id');
  if (LEGACY_WRONG_ORG_IDS.has(fromQuery)) {
    console.warn(
      `payroll_function: remapping organization_id ${fromQuery} -> ${DEFAULT_PAYROLL_ORGANIZATION_ID}`
    );
    return DEFAULT_PAYROLL_ORGANIZATION_ID;
  }
  return fromQuery || DEFAULT_PAYROLL_ORGANIZATION_ID;
}

function safeJsonParse(value, fallback = null) {
  if (value == null || value === '') return fallback;
  if (typeof value === 'object') return value;
  try {
    return JSON.parse(String(value));
  } catch (_) {
    return fallback;
  }
}

function normalizePayrollMonth(value) {
  if (value == null || value === '') return '';
  const text = String(value).trim();
  const match = text.match(/^(\d{4})[-/](\d{1,2})$/);
  if (!match) return text;
  return `${match[1]}-${String(Number(match[2])).padStart(2, '0')}`;
}

function pickPayrollDatastoreRow(entry) {
  if (!entry || typeof entry !== 'object') return null;
  return entry.Payroll || entry.payroll || entry;
}

function readRowId(row) {
  if (!row || typeof row !== 'object') return null;
  return row.ROWID ?? row.rowid ?? row.RowId ?? row.id ?? null;
}

function readRowDataField(row) {
  if (!row || typeof row !== 'object') return null;
  return row.Data ?? row.data ?? row.DATA ?? null;
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

function extractPayrollMonthFromParsed(parsed) {
  if (!parsed || typeof parsed !== 'object') return '';
  return normalizePayrollMonth(
    parsed.payrollMonth ?? parsed.PayrollMonth ?? parsed.month ?? parsed.Month
  );
}

function rowHasUsablePayrollData(rawData) {
  const parsed = parseDatastoreJson(rawData);
  if (!parsed || typeof parsed !== 'object') {
    const raw = String(rawData || '');
    return /"payrollMonth"\s*:\s*"\d{4}-\d{2}"/i.test(raw);
  }
  if (extractPayrollMonthFromParsed(parsed)) return true;
  return extractPayrollRecordsFromParsed(parsed).length > 0;
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

function stashPayrollRow(map, row) {
  const base = pickPayrollDatastoreRow(row) || row;
  const rowId = readRowId(base);
  if (rowId == null) return;
  map.set(String(rowId), base);
}

async function fetchAllPayrollDatastoreRows(catalyst) {
  const table = catalyst.datastore().table(PAYROLL_TABLE);
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
    } catch (getAllErr) {
      console.warn('Payroll getAllRows failed:', getAllErr.message || getAllErr);
    }
  }

  if (byId.size === 0) {
    try {
      const zcql = catalyst.zcql();
      const rows = await zcql.executeZCQLQuery(
        `SELECT ROWID, Data, MODIFIEDTIME, CREATEDTIME FROM ${PAYROLL_TABLE}`
      );
      const zlist = Array.isArray(rows) ? rows : [];
      for (const entry of zlist) {
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

function collectAvailablePayrollMonths(rows) {
  const months = new Set();
  for (const row of rows) {
    const rawData = readRowDataField(row);
    const parsed = parseDatastoreJson(rawData);
    const month = extractPayrollMonthFromParsed(parsed);
    if (month) months.add(month);
    if (!month) {
      const raw = String(rawData || '');
      const match = raw.match(/"payrollMonth"\s*:\s*"(\d{4}-\d{2})"/i);
      if (match) months.add(normalizePayrollMonth(match[1]));
    }
  }
  return [...months].sort();
}

function matchPayrollRowToMonth(row, month) {
  const rawData = readRowDataField(row);
  let parsed = parseDatastoreJson(rawData);
  if (!parsed || typeof parsed !== 'object') {
    const raw = String(rawData || '');
    const match = raw.match(/"payrollMonth"\s*:\s*"(\d{4}-\d{2})"/i);
    if (!match || normalizePayrollMonth(match[1]) !== month) {
      return null;
    }
    parsed = parseDatastoreJson(rawData) || {};
  }
  const rowMonth = extractPayrollMonthFromParsed(parsed);
  if (rowMonth !== month) {
    const raw = String(rawData || '');
    const match = raw.match(/"payrollMonth"\s*:\s*"(\d{4}-\d{2})"/i);
    if (!match || normalizePayrollMonth(match[1]) !== month) return null;
  }
  if (!parsed || typeof parsed !== 'object') {
    parsed = parseDatastoreJson(rawData) || {};
  }
  return parsed;
}

async function tryLoadSnapshotByMonthLike(catalyst, month) {
  const table = catalyst.datastore().table(PAYROLL_TABLE);
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
        `SELECT ROWID, Data, MODIFIEDTIME, CREATEDTIME FROM ${PAYROLL_TABLE} WHERE Data LIKE '${safe}'`
      );
      const list = Array.isArray(rows) ? rows : [];
      let best = null;
      for (const entry of list) {
        const row = await hydratePayrollDatastoreRow(
          table,
          pickPayrollDatastoreRow(entry) || entry
        );
        if (!row) continue;
        const parsed = matchPayrollRowToMonth(row, month);
        if (!parsed) continue;
        const modified = String(row.MODIFIEDTIME || row.CREATEDTIME || '');
        if (!best || modified > String(best.row.MODIFIEDTIME || best.row.CREATEDTIME || '')) {
          best = { row, parsed };
        }
      }
      if (best) return best;
    } catch (err) {
      console.warn('Payroll ZCQL month LIKE failed:', err.message || err);
    }
  }
  return null;
}

async function persistPayrollSnapshot(catalyst, payload) {
  const payrollMonth = String(payload?.payrollMonth || '').trim();
  const records = Array.isArray(payload?.records) ? payload.records : [];
  if (!/^\d{4}-\d{2}$/.test(payrollMonth)) {
    return { saved: false, reason: 'invalid_month' };
  }
  if (records.length === 0) {
    return { saved: false, reason: 'empty_records' };
  }

  const flattenedRecords = records.map((row) => flattenPayrollEarningColumns(row));

  const totalExpected = Number.isFinite(Number(payload?.totalExpected))
    ? Number(payload.totalExpected)
    : records.length;

  const dataPayload = {
    payrollMonth,
    organizationId: payload?.organizationId || null,
    runMeta: payload?.runMeta && typeof payload.runMeta === 'object' ? payload.runMeta : null,
    records: flattenedRecords,
    hasBreakdown: payload?.hasBreakdown === true,
    loadedCount: records.length,
    totalExpected,
    breakdownComplete: payload?.breakdownComplete === true,
    savedAt: new Date().toISOString(),
  };

  const table = catalyst.datastore().table(PAYROLL_TABLE);
  const zcql = catalyst.zcql();
  let existingRowId = null;

  try {
    const rows = await zcql.executeZCQLQuery(`SELECT ROWID, Data FROM ${PAYROLL_TABLE}`);
    const list = Array.isArray(rows) ? rows : [];
    for (const entry of list) {
      const row = pickPayrollDatastoreRow(entry);
      if (!row || row.ROWID == null) continue;
      const parsed = safeJsonParse(row.Data, {});
      if (parsed && parsed.payrollMonth === payrollMonth) {
        existingRowId = String(row.ROWID);
        break;
      }
    }
  } catch (queryErr) {
    console.warn('Payroll save: month lookup failed, will insert:', queryErr.message);
  }

  const dataText = JSON.stringify(dataPayload);
  if (existingRowId) {
    await table.updateRow({ ROWID: existingRowId, Data: dataText });
    return {
      saved: true,
      rowId: existingRowId,
      updated: true,
      recordCount: records.length,
      totalExpected,
      payrollMonth,
    };
  }

  const insertResp = await table.insertRow({ Data: dataText });
  return {
    saved: true,
    rowId: insertResp?.ROWID || null,
    updated: false,
    recordCount: records.length,
    totalExpected,
    payrollMonth,
  };
}

function extractPayrollRecordsFromParsed(parsed) {
  if (!parsed || typeof parsed !== 'object') return [];
  const records = parsed.records ?? parsed.Records ?? parsed.data ?? parsed.Data;
  return Array.isArray(records) ? records : [];
}

function buildPayrollSnapshotResult(best, payrollMonth) {
  const records = extractPayrollRecordsFromParsed(best.parsed).map((row) =>
    flattenPayrollEarningColumns(row)
  );
  const runMeta =
    best.parsed.runMeta && typeof best.parsed.runMeta === 'object' ? best.parsed.runMeta : {};
  return {
    found: true,
    rowId: best.row.ROWID,
    payrollMonth,
    records,
    meta: {
      ...runMeta,
      hasBreakdown: best.parsed.hasBreakdown === true,
      loadedCount: best.parsed.loadedCount,
      totalExpected: best.parsed.totalExpected,
      breakdownComplete: best.parsed.breakdownComplete === true,
      savedAt: best.parsed.savedAt || null,
      payDate: runMeta.pay_date || runMeta.payDate || null,
      pay_date: runMeta.pay_date || runMeta.payDate || null,
      payroll_run_id: runMeta.payroll_run_id || null,
      organizationId: best.parsed.organizationId || null,
      source: best.parsed.source || null,
    },
  };
}

async function loadPayrollSnapshotFromTable(catalyst, payrollMonth) {
  const month = normalizePayrollMonth(payrollMonth);
  if (!/^\d{4}-\d{2}$/.test(month)) {
    return { found: false, reason: 'invalid_month' };
  }

  try {
    let best = await tryLoadSnapshotByMonthLike(catalyst, month);
    const list = await fetchAllPayrollDatastoreRows(catalyst);

    if (!best) {
      for (const row of list) {
        if (!row || readRowId(row) == null) continue;
        const parsed = matchPayrollRowToMonth(row, month);
        if (!parsed) continue;
        const modified = String(row.MODIFIEDTIME || row.CREATEDTIME || '');
        if (!best || modified > String(best.row.MODIFIEDTIME || best.row.CREATEDTIME || '')) {
          best = { row, parsed };
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
    return buildPayrollSnapshotResult(best, month);
  } catch (err) {
    console.error('Payroll load error:', err.message || err);
    return { found: false, reason: 'query_failed', error: err.message || String(err) };
  }
}

/** Most recently modified payroll snapshot in the Payroll table (any month). */
async function loadLatestPayrollSnapshotFromTable(catalyst) {
  const zcql = catalyst.zcql();
  try {
    const rows = await zcql.executeZCQLQuery(`SELECT ROWID, Data, MODIFIEDTIME FROM ${PAYROLL_TABLE}`);
    const list = Array.isArray(rows) ? rows : [];
    let best = null;
    for (const entry of list) {
      const row = pickPayrollDatastoreRow(entry);
      if (!row || row.ROWID == null) continue;
      const parsed = safeJsonParse(row.Data, null);
      if (typeof parsed === 'string') parsed = safeJsonParse(parsed, {});
      if (!parsed || typeof parsed !== 'object') continue;
      const payrollMonth = normalizePayrollMonth(
        parsed.payrollMonth ?? parsed.PayrollMonth ?? parsed.month ?? parsed.Month
      );
      if (!/^\d{4}-\d{2}$/.test(payrollMonth)) continue;
      if (!Array.isArray(parsed.records) || parsed.records.length === 0) continue;
      const modified = String(row.MODIFIEDTIME || row.CREATEDTIME || '');
      if (!best || modified > String(best.row.MODIFIEDTIME || best.row.CREATEDTIME || '')) {
        best = { row, parsed, payrollMonth };
      }
    }
    if (!best) {
      return { found: false, reason: 'not_found' };
    }
    return buildPayrollSnapshotResult(best, best.payrollMonth);
  } catch (err) {
    console.error('Payroll latest load error:', err.message || err);
    return { found: false, reason: 'query_failed', error: err.message || String(err) };
  }
}

async function handlePayrollFetch(req, res) {
  try {
    const organizationId = resolveOrganizationId(req);
    if (!organizationId) {
      res.writeHead(400, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ success: false, error: 'organization_id is required' }));
      return;
    }

    const employeeId = readQueryParam(req, 'employee_id');
    const payrollRunId = readQueryParam(req, 'payroll_run_id');
    const allSalaries = readQueryParam(req, 'all_salaries') === '1';
    const listEmployees = readQueryParam(req, 'list_employees') === '1';
    const payrollMonthData = readQueryParam(req, 'payroll_month_data') === '1';
    const payrunEmployeeDetail = readQueryParam(req, 'payrun_employee_detail') === '1';

    const accessToken = await getPayrollAccessToken();

    if (payrunEmployeeDetail) {
      const employeeIdsRaw = readQueryParam(req, 'employee_ids');
      const batchIds = employeeIdsRaw
        ? employeeIdsRaw
            .split(',')
            .map((id) => String(id || '').trim())
            .filter(Boolean)
            .slice(0, 25)
        : [];
      const singleId = employeeId || (batchIds.length === 1 ? batchIds[0] : '');

      if (!payrollRunId || (!singleId && batchIds.length === 0)) {
        res.writeHead(400, { 'Content-Type': 'application/json' });
        res.end(
          JSON.stringify({
            success: false,
            error:
              'payroll_run_id and employee_id (or employee_ids, max 25) are required for payrun_employee_detail',
          })
        );
        return;
      }

      // Batch: up to 25 employees per call (Load salary breakdown).
      if (batchIds.length > 1 || (employeeIdsRaw && batchIds.length >= 1)) {
        const rows = await enrichPayrollRunRowsWithEmployeeDetail({
          accessToken,
          organizationId,
          payrollRunId,
          rows: batchIds.map((id) => ({ employee_id: id })),
          concurrency: Math.max(
            1,
            Math.min(
              5,
              parseInt(process.env.ZOHO_PAYROLL_DETAIL_CONCURRENCY || '3', 10) || 3
            )
          ),
        });
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ success: true, data: rows }));
        return;
      }

      const data = flattenPayrollEarningColumns(
        await fetchPayrollRunEmployeeDetail({
          accessToken,
          organizationId,
          payrollRunId,
          employeeId: singleId,
        })
      );
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ success: true, data }));
      return;
    }

    if (payrollMonthData) {
      const monthParts = parsePayrollMonthParam(req);
      if (!monthParts) {
        res.writeHead(400, { 'Content-Type': 'application/json' });
        res.end(
          JSON.stringify({
            success: false,
            error: 'payroll_month=YYYY-MM (or year + month) is required for payroll_month_data',
          })
        );
        return;
      }
      const employeeOffset = Math.max(
        0,
        parseInt(readQueryParam(req, 'employee_offset') || '0', 10) || 0
      );
      const runType = readQueryParam(req, 'payroll_run_type') || 'regular';
      const includeEarningsDetail = readQueryParam(req, 'include_earnings_detail') !== '0';
      const employeeLimitRaw = readQueryParam(req, 'employee_limit');
      const employeeLimit = employeeLimitRaw
        ? Math.max(1, Math.min(50, parseInt(employeeLimitRaw, 10) || 1))
        : includeEarningsDetail
          ? 15
          : 50;
      const detailLimitRaw = readQueryParam(req, 'detail_limit');
      const detailLimit = detailLimitRaw
        ? Math.max(1, Math.min(50, parseInt(detailLimitRaw, 10) || 1))
        : null;

      const payload = await fetchPayrollMonthEmployeeBatch({
        accessToken,
        organizationId,
        year: monthParts.year,
        month: monthParts.month,
        offset: employeeOffset,
        limit: employeeLimit,
        runType,
        includeEarningsDetail,
        detailLimit,
      });

      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(
        JSON.stringify({
          success: true,
          data: payload.rows,
          meta: payload.meta,
        })
      );
      return;
    }

    if (employeeId && !allSalaries && !payrollRunId) {
      const data = flattenPayrollEarningColumns(
        await fetchMergedEmployeeSalary({
          accessToken,
          organizationId,
          employeeId,
        })
      );
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ success: true, data }));
      return;
    }

    if (allSalaries) {
      const salaryOffset = Math.max(
        0,
        parseInt(readQueryParam(req, 'salary_offset') || '0', 10) || 0
      );
      const salaryLimitRaw = readQueryParam(req, 'salary_limit');
      const salaryLimit = salaryLimitRaw
        ? Math.max(1, parseInt(salaryLimitRaw, 10) || 1)
        : null;

      let payload;
      try {
        payload = await fetchAllEmployeesWithSalary({
          accessToken,
          organizationId,
          offset: salaryOffset,
          limit: salaryLimit,
        });
      } catch (salaryBatchErr) {
        console.warn('payroll_function: all_salaries batch failed, returning employee list only:', salaryBatchErr.message);
        const employees = await fetchAllEmployeePages({ accessToken, organizationId });
        const slice =
          salaryLimit != null
            ? employees.slice(salaryOffset, salaryOffset + salaryLimit)
            : employees;
        payload = {
          rows: slice.map((emp) => ({
            ...emp,
            employee_id: getEmployeeRecordId(emp) || emp.employee_id,
            salary: null,
            fetch_error: true,
            error: salaryBatchErr.message,
          })),
          total: employees.length,
          offset: salaryOffset,
          limit: slice.length,
          has_more: salaryLimit != null && salaryOffset + slice.length < employees.length,
        };
      }

      const body = { success: true, data: payload.rows };
      if (salaryLimit != null) {
        body.meta = {
          total: payload.total,
          offset: payload.offset,
          limit: payload.limit,
          has_more: payload.has_more,
        };
      }
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify(body));
      return;
    }

    if (listEmployees) {
      const data = await fetchAllEmployeePages({ accessToken, organizationId });
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ success: true, data }));
      return;
    }

    res.writeHead(400, { 'Content-Type': 'application/json' });
    res.end(
      JSON.stringify({
        success: false,
        error:
          'Specify employee_id, all_salaries=1, list_employees=1, payroll_month_data=1, or payrun_employee_detail=1',
      })
    );
  } catch (error) {
    const z = error.response?.data;
    const zohoDetail =
      z && typeof z === 'object'
        ? z.error || z.message || (z.code != null ? `code ${z.code}` : null)
        : null;
    const msg = [zohoDetail, error.message].filter(Boolean).join(' — ') || 'Payroll request failed';
    console.error('payroll_function error:', z || error.message);
    res.writeHead(500, { 'Content-Type': 'application/json' });
    res.end(
      JSON.stringify({
        success: false,
        error: msg,
      })
    );
  }
};

/** In-code defaults (Zoho Payroll India, api_domain https://www.zohoapis.in).
 * Scopes: ZohoPayroll.salary.ALL ZohoPayroll.employee.ALL ZohoPayroll.settings.ALL ZohoPayroll.worklocation.ALL ZohoPayroll.payrollrun.ALL
 * Env vars override when set.
 */
const DEFAULT_ZOHO_PAYROLL_CLIENT_ID = '1000.VEO83G2Y7D16OXNQC7CN5WTK7DFHYA';
const DEFAULT_ZOHO_PAYROLL_CLIENT_SECRET = 'b6d3935145d59974934b981d6291b40af69a3ab150';
/** Default refresh_token from Zoho Payroll OAuth response (grant_type=refresh_token). */
const DEFAULT_ZOHO_PAYROLL_REFRESH_TOKEN =
  '1000.ab6fc119a4cba66c406d7e2de608ce69.32a06c8fa10a10f261099ac7ec537aa4';
/** Fallback if primary refresh is revoked — set ZOHO_PAYROLL_REFRESH_TOKEN in Catalyst to override. */
const DEFAULT_ZOHO_PAYROLL_REFRESH_TOKEN_ALT =
  '1000.5aac4ac693ef072d01fd9a7cfdb48164.d29c4e3270ef4b70a60fb90227799a59';

/** Reuse access token within one function invocation (avoids multiple refresh calls per request). */
let cachedPayrollAccessToken = null;
let cachedPayrollAccessTokenExpiresAt = 0;

/** Catalyst sometimes defines env keys with empty strings — treat those as unset. */
function envOr(name, fallback) {
  const v = process.env[name];
  if (v == null) return fallback;
  const t = String(v).trim();
  return t === '' ? fallback : t;
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function backoffMsFor429(headers, attempt) {
  const maxBackoff = Math.max(
    5000,
    parseInt(process.env.ZOHO_PAYROLL_429_MAX_BACKOFF_MS || '30000', 10) || 30000
  );
  const retryAfter = parseInt(headers?.['retry-after'], 10);
  if (Number.isFinite(retryAfter) && retryAfter > 0) {
    return Math.min(maxBackoff, retryAfter * 1000);
  }
  return Math.min(maxBackoff, 1500 * 2 ** (attempt - 1));
}

async function axiosRequestWith429Retry(requestFn, label) {
  const maxAttempts = Math.max(1, parseInt(process.env.ZOHO_PAYROLL_HTTP_RETRIES || '3', 10) || 3);
  let lastErr;
  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    try {
      return await requestFn();
    } catch (e) {
      lastErr = e;
      const status = e.response?.status;
      const code = e.response?.data?.code;
      const is429 = status === 429 || code === 43;
      if (!is429 || attempt === maxAttempts) throw e;
      const backoffMs = backoffMsFor429(e.response?.headers, attempt);
      console.warn(`${label}: HTTP 429, retry ${attempt}/${maxAttempts} after ${backoffMs}ms`);
      await sleep(backoffMs);
    }
  }
  throw lastErr;
}

/**
 * Same token model as peopledata_function / leavedata_function / attendance_function:
 * 1) ZOHO_PAYROLL_ACCESS_TOKEN (or ZOHO_ACCESS_TOKEN) when set
 * 2) refresh_token → access_token (cached for this invocation)
 */
async function getPayrollAccessToken() {
  const now = Date.now();
  if (cachedPayrollAccessToken && now < cachedPayrollAccessTokenExpiresAt - 60_000) {
    return cachedPayrollAccessToken;
  }

  const directAccess = envOr('ZOHO_PAYROLL_ACCESS_TOKEN', '') || envOr('ZOHO_ACCESS_TOKEN', '');
  if (directAccess.length > 10) {
    cachedPayrollAccessToken = directAccess;
    cachedPayrollAccessTokenExpiresAt = now + 55 * 60 * 1000;
    return directAccess;
  }

  const clientId = envOr('ZOHO_PAYROLL_CLIENT_ID', DEFAULT_ZOHO_PAYROLL_CLIENT_ID);
  const clientSecret = envOr('ZOHO_PAYROLL_CLIENT_SECRET', DEFAULT_ZOHO_PAYROLL_CLIENT_SECRET);
  if (!clientId || !clientSecret) {
    throw new Error(
      'Missing Zoho Payroll OAuth. Set ZOHO_PAYROLL_REFRESH_TOKEN, ZOHO_PAYROLL_CLIENT_ID, ZOHO_PAYROLL_CLIENT_SECRET, or ZOHO_PAYROLL_ACCESS_TOKEN.'
    );
  }

  const tokenUrl = envOr('ZOHO_PAYROLL_ACCOUNTS_URL', 'https://accounts.zoho.in/oauth/v2/token');

  const refreshOnce = async (rt) => {
    const params = new URLSearchParams({
      refresh_token: rt,
      client_id: clientId,
      client_secret: clientSecret,
      grant_type: 'refresh_token',
    });
    let response;
    try {
      response = await axios.post(tokenUrl, params, {
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      });
    } catch (e) {
      const bd = e.response?.data;
      const detail =
        bd && typeof bd === 'object'
          ? bd.error || bd.error_description || bd.message
          : e.message;
      throw new Error(`Zoho OAuth token HTTP ${e.response?.status || '?'}: ${detail || 'token request failed'}`);
    }
    const data = response.data || {};
    if (!data.access_token) {
      const detail = data.error || data.error_description || data.message || JSON.stringify(data);
      throw new Error(`Failed to obtain Zoho Payroll access token: ${detail}`);
    }
    cachedPayrollAccessToken = data.access_token;
    const expiresIn = parseInt(data.expires_in, 10) || 3600;
    cachedPayrollAccessTokenExpiresAt = now + expiresIn * 1000;
    return cachedPayrollAccessToken;
  };

  const envRt = envOr('ZOHO_PAYROLL_REFRESH_TOKEN', '');
  const tokenCandidates = [
    envRt,
    DEFAULT_ZOHO_PAYROLL_REFRESH_TOKEN,
    DEFAULT_ZOHO_PAYROLL_REFRESH_TOKEN_ALT,
  ].filter((t, i, arr) => t && arr.indexOf(t) === i);

  let lastErr;
  for (let i = 0; i < tokenCandidates.length; i++) {
    try {
      return await refreshOnce(tokenCandidates[i]);
    } catch (e) {
      lastErr = e;
      if (i < tokenCandidates.length - 1) {
        console.warn(`payroll_function: refresh attempt ${i + 1} failed, trying next token`);
      }
    }
  }
  throw lastErr || new Error('Failed to obtain Zoho Payroll access token.');
}

function getApiBase() {
  return envOr('ZOHO_PAYROLL_API_BASE', 'https://www.zohoapis.in/payroll/v1').replace(/\/+$/, '');
}

async function payrollAxiosGet(url, config, label) {
  return axiosRequestWith429Retry(() => axios.get(url, config), label || `GET ${url}`);
}

function extractListPayload(data) {
  if (!data || typeof data !== 'object') return [];
  const candidates = [
    data.employees,
    data.data,
    data.results,
    data.response,
    data.employee,
  ];
  for (const c of candidates) {
    if (Array.isArray(c)) return c;
    if (c && typeof c === 'object' && Array.isArray(c.data)) return c.data;
  }
  return [];
}

function getEmployeeRecordId(emp) {
  if (!emp || typeof emp !== 'object') return '';
  return String(
    emp.employee_id ?? emp.employeeId ?? emp.id ?? emp.emp_id ?? emp.employee_id_str ?? ''
  ).trim();
}

async function fetchAllEmployeePages({ accessToken, organizationId }) {
  const base = getApiBase();
  const collected = [];
  let page = 1;
  const perPage = Math.min(
    200,
    Math.max(1, parseInt(process.env.ZOHO_PAYROLL_PAGE_SIZE || '200', 10) || 200)
  );

  for (let guard = 0; guard < 500; guard++) {
    const { data } = await payrollAxiosGet(
      `${base}/employees`,
      {
        headers: { Authorization: `Bearer ${accessToken}` },
        params: {
          organization_id: organizationId,
          page,
          per_page: perPage,
        },
      },
      'Zoho employees list'
    );

    const chunk = extractListPayload(data);
    collected.push(...chunk);

    const ctx = data?.page_context || data?.pageContext || {};
    const hasMore =
      ctx.has_more_page === true ||
      ctx.has_more === true ||
      data?.has_more_page === true ||
      chunk.length >= perPage;

    if (!hasMore || chunk.length === 0) break;
    page += 1;
    await sleep(Math.max(0, parseInt(process.env.ZOHO_PAYROLL_LIST_PAGE_DELAY_MS || '120', 10) || 120));
  }

  return collected;
}

async function fetchEmployeeById({ accessToken, organizationId, employeeId }) {
  const base = getApiBase();
  const { data } = await payrollAxiosGet(
    `${base}/employees/${encodeURIComponent(employeeId)}`,
    {
      headers: { Authorization: `Bearer ${accessToken}` },
      params: { organization_id: organizationId },
    },
    'Zoho employee by id'
  );
  if (data && typeof data === 'object' && data.employee && typeof data.employee === 'object') {
    return data.employee;
  }
  return data;
}

async function fetchSalaryPayload({ accessToken, organizationId, employeeId }) {
  const base = getApiBase();
  const paths = ['/salary', '/salarydetails'];

  for (const p of paths) {
    try {
      const { data } = await payrollAxiosGet(
        `${base}/employees/${encodeURIComponent(employeeId)}${p}`,
        {
          headers: { Authorization: `Bearer ${accessToken}` },
          params: { organization_id: organizationId },
        },
        `Zoho salary ${employeeId}`
      );
      if (data !== undefined && data !== null) return { path: p, body: data };
    } catch (e) {
      const status = e.response?.status;
      const code = e.response?.data?.code;
      if (status === 404 || code === 5) continue;
      throw e;
    }
  }
  return null;
}

async function fetchMergedEmployeeSalary({ accessToken, organizationId, employeeId }) {
  let employee = null;
  try {
    employee = await fetchEmployeeById({ accessToken, organizationId, employeeId });
  } catch (e) {
    console.warn('payroll_function: employee fetch failed, continuing with id only', e.message);
  }

  const salaryWrap = await fetchSalaryPayload({ accessToken, organizationId, employeeId });
  const salaryEmployee = unwrapSalaryEmployeePayload(salaryWrap?.body);

  const baseRow =
    employee && typeof employee === 'object'
      ? { ...employee }
      : { employee_id: employeeId };

  return {
    ...baseRow,
    ...salaryEmployee,
    employee_id: getEmployeeRecordId(baseRow) || String(employeeId),
    salary: salaryWrap?.body ?? null,
  };
}

const PAYROLL_MONTH_NAMES = [
  'January',
  'February',
  'March',
  'April',
  'May',
  'June',
  'July',
  'August',
  'September',
  'October',
  'November',
  'December',
];

function parsePayrollMonthParam(req) {
  const payrollMonth = readQueryParam(req, 'payroll_month');
  if (/^\d{4}-\d{2}$/.test(payrollMonth)) {
    const [y, m] = payrollMonth.split('-');
    const year = parseInt(y, 10);
    const month = parseInt(m, 10);
    if (Number.isFinite(year) && month >= 1 && month <= 12) return { year, month };
  }
  const year = parseInt(readQueryParam(req, 'year') || '', 10);
  const month = parseInt(readQueryParam(req, 'month') || '', 10);
  if (Number.isFinite(year) && Number.isFinite(month) && month >= 1 && month <= 12) {
    return { year, month };
  }
  return null;
}

function findPayrollRunForMonth(runs, year, month, preferredType = 'regular') {
  if (!Array.isArray(runs) || runs.length === 0) return null;
  const monthStr = String(month).padStart(2, '0');
  const prefix = `${year}-${monthStr}`;
  const periodLabel = `${PAYROLL_MONTH_NAMES[month - 1]} ${year}`.toLowerCase();

  const matches = runs.filter((run) => {
    const start = String(run.pay_period_start_date || '');
    const period = String(run.processing_period || '').toLowerCase().trim();
    return start.startsWith(prefix) || period === periodLabel;
  });

  if (matches.length === 0) return null;

  const typeMatch = matches.find((run) => String(run.type || '') === preferredType);
  if (typeMatch) return typeMatch;

  const statusPriority = ['completed', 'paid', 'approved', 'draft'];
  const ranked = [...matches].sort((a, b) => {
    const ai = statusPriority.indexOf(String(a.status || '').toLowerCase());
    const bi = statusPriority.indexOf(String(b.status || '').toLowerCase());
    return (ai < 0 ? 99 : ai) - (bi < 0 ? 99 : bi);
  });
  return ranked[0];
}

async function fetchAllPayrollRunPages({ accessToken, organizationId }) {
  const base = getApiBase();
  const collected = [];
  let page = 1;
  const perPage = Math.min(
    200,
    Math.max(1, parseInt(process.env.ZOHO_PAYROLL_PAGE_SIZE || '200', 10) || 200)
  );

  for (let guard = 0; guard < 100; guard++) {
    const { data } = await payrollAxiosGet(
      `${base}/payrollruns`,
      {
        headers: { Authorization: `Bearer ${accessToken}` },
        params: {
          organization_id: organizationId,
          page,
          per_page: perPage,
        },
      },
      'Zoho payroll runs list'
    );

    const chunk = Array.isArray(data?.payroll_runs) ? data.payroll_runs : [];
    collected.push(...chunk);

    const ctx = data?.page_context || {};
    const hasMore = ctx.has_more_page === true || chunk.length >= perPage;
    if (!hasMore || chunk.length === 0) break;
    page += 1;
    await sleep(Math.max(0, parseInt(process.env.ZOHO_PAYROLL_LIST_PAGE_DELAY_MS || '120', 10) || 120));
  }

  return collected;
}

let cachedPayrollRunsList = null;
let cachedPayrollRunsOrgId = '';
let cachedPayrollRunsTs = 0;
const PAYROLL_RUNS_CACHE_TTL_MS = 5 * 60 * 1000;

async function fetchAllPayrollRunPagesCached({ accessToken, organizationId }) {
  const org = String(organizationId || '').trim();
  if (
    cachedPayrollRunsList &&
    cachedPayrollRunsOrgId === org &&
    Date.now() - cachedPayrollRunsTs < PAYROLL_RUNS_CACHE_TTL_MS
  ) {
    return cachedPayrollRunsList;
  }
  const runs = await fetchAllPayrollRunPages({ accessToken, organizationId });
  cachedPayrollRunsList = runs;
  cachedPayrollRunsOrgId = org;
  cachedPayrollRunsTs = Date.now();
  return runs;
}

async function fetchPayrollRunById({ accessToken, organizationId, payrollRunId }) {
  const base = getApiBase();
  const { data } = await payrollAxiosGet(
    `${base}/payrollruns/${encodeURIComponent(payrollRunId)}`,
    {
      headers: { Authorization: `Bearer ${accessToken}` },
      params: { organization_id: organizationId },
    },
    'Zoho payroll run by id'
  );
  return data?.payroll_run || data;
}

async function fetchPayrollRunEmployeesPage({
  accessToken,
  organizationId,
  payrollRunId,
  page = 1,
  perPage = 200,
}) {
  const base = getApiBase();
  const { data } = await payrollAxiosGet(
    `${base}/payrollruns/${encodeURIComponent(payrollRunId)}/employees`,
    {
      headers: { Authorization: `Bearer ${accessToken}` },
      params: {
        organization_id: organizationId,
        page,
        per_page: perPage,
      },
    },
    'Zoho payroll run employees'
  );
  const employees = Array.isArray(data?.employees) ? data.employees : [];
  const ctx = data?.page_context || {};
  return {
    employees,
    page: ctx.page || page,
    perPage: ctx.per_page || perPage,
    hasMore: ctx.has_more_page === true,
    total:
      parseInt(ctx.total ?? ctx.total_records ?? ctx.total_count ?? ctx.total_employees, 10) || null,
  };
}

async function mapWithConcurrency(items, concurrency, mapper) {
  const list = Array.isArray(items) ? items : [];
  if (list.length === 0) return [];
  const results = new Array(list.length);
  let nextIndex = 0;
  const workerCount = Math.max(1, Math.min(concurrency, list.length));

  async function worker() {
    while (nextIndex < list.length) {
      const current = nextIndex;
      nextIndex += 1;
      results[current] = await mapper(list[current], current);
    }
  }

  await Promise.all(Array.from({ length: workerCount }, () => worker()));
  return results;
}

/** Pay-run list rows only include net_pay / total_earnings; per-employee detail has Basic, HRA, etc. */
async function enrichPayrollRunRowsWithEmployeeDetail({
  accessToken,
  organizationId,
  payrollRunId,
  rows,
  concurrency = 5,
}) {
  if (!Array.isArray(rows) || rows.length === 0) return [];
  const runId = String(payrollRunId || '').trim();
  if (!runId) return rows.map((row) => flattenPayrollEarningColumns(row));

  const detailConcurrency = Math.max(
    1,
    concurrency ||
      parseInt(process.env.ZOHO_PAYROLL_DETAIL_CONCURRENCY || '5', 10) ||
      5
  );

  return mapWithConcurrency(rows, detailConcurrency, async (row) => {
    const employeeId = getEmployeeRecordId(row);
    if (!employeeId) return flattenPayrollEarningColumns(row);
    const listFlat = flattenPayrollEarningColumns(row);
    try {
      const detail = await fetchPayrollRunEmployeeDetail({
        accessToken,
        organizationId,
        payrollRunId: runId,
        employeeId,
      });
      const detailFlat = flattenPayrollEarningColumns(detail);
      const merged = { ...listFlat, ...detailFlat, employee_id: employeeId };
      const preserveKeys = ['net_pay', 'gross_pay', 'total_earnings', 'paid_days', 'total_deductions'];
      for (const key of preserveKeys) {
        const listVal = listFlat[key];
        const detailVal = detailFlat[key];
        const listNum = Number(listVal);
        const detailNum = Number(detailVal);
        if ((!Number.isFinite(detailNum) || detailNum === 0) && Number.isFinite(listNum) && listNum !== 0) {
          merged[key] = listVal;
        }
      }
      return flattenPayrollEarningColumns(merged);
    } catch (detailErr) {
      console.warn(
        `payroll_function: payrun employee detail failed for ${employeeId}:`,
        detailErr.message || detailErr
      );
      return flattenPayrollEarningColumns(row);
    }
  });
}

async function fetchPayrollRunEmployeeDetail({
  accessToken,
  organizationId,
  payrollRunId,
  employeeId,
}) {
  const base = getApiBase();
  const { data } = await payrollAxiosGet(
    `${base}/payrollruns/${encodeURIComponent(payrollRunId)}/employees/${encodeURIComponent(employeeId)}`,
    {
      headers: { Authorization: `Bearer ${accessToken}` },
      params: { organization_id: organizationId },
    },
    `Zoho payroll run employee ${employeeId}`
  );
  let employee = mergePayrollRunEmployeePayload(data);
  if (getEarningsArray(employee).length === 0) {
    try {
      const salaryWrap = await fetchSalaryPayload({ accessToken, organizationId, employeeId });
      const salaryEmployee = unwrapSalaryEmployeePayload(salaryWrap?.body);
      if (getEarningsArray(salaryEmployee).length > 0) {
        employee = {
          ...employee,
          ...salaryEmployee,
          earnings_source: 'salary_template',
        };
      }
    } catch (salaryErr) {
      console.warn(
        `payroll_function: salary template fallback for ${employeeId}:`,
        salaryErr.message
      );
    }
  }
  return employee;
}

async function fetchPayrollMonthEmployeeBatch({
  accessToken,
  organizationId,
  year,
  month,
  offset = 0,
  limit = 50,
  runType = 'regular',
  includeEarningsDetail = true,
  detailLimit = null,
}) {
  const runs = await fetchAllPayrollRunPagesCached({ accessToken, organizationId });
  const run = findPayrollRunForMonth(runs, year, month, runType);
  if (!run) {
    throw new Error(
      `No ${runType} pay run found for ${PAYROLL_MONTH_NAMES[month - 1]} ${year}. Process payroll in Zoho Payroll for that month first.`
    );
  }

  const payrollRunId = String(run.payroll_run_id || '').trim();
  const perPage = Math.min(50, Math.max(1, limit));
  const page = Math.floor(offset / perPage) + 1;
  const pageBatch = await fetchPayrollRunEmployeesPage({
    accessToken,
    organizationId,
    payrollRunId,
    page,
    perPage,
  });

  let rows = pageBatch.employees;
  const sliceStart = offset % perPage;
  if (sliceStart > 0 || rows.length > limit) {
    rows = rows.slice(sliceStart, sliceStart + limit);
  }

  let total = parseInt(run.no_of_employees, 10);
  if (Number.isFinite(pageBatch.total) && pageBatch.total > 0) {
    total = pageBatch.total;
  }
  if (!Number.isFinite(total) || total <= 0) {
    try {
      const runDetail = await fetchPayrollRunById({ accessToken, organizationId, payrollRunId });
      total = parseInt(runDetail?.no_of_employees, 10);
    } catch (_) {
      /* ignore */
    }
  }
  if (!Number.isFinite(total) || total <= 0) {
    total = offset + rows.length + (pageBatch.hasMore ? perPage : 0);
  }

  const sameMonthRuns = runs.filter((item) => {
    const monthStr = String(month).padStart(2, '0');
    const prefix = `${year}-${monthStr}`;
    const start = String(item.pay_period_start_date || '');
    const period = String(item.processing_period || '').toLowerCase().trim();
    return (
      start.startsWith(prefix) ||
      period === `${PAYROLL_MONTH_NAMES[month - 1]} ${year}`.toLowerCase()
    );
  });

  const maxDetailRows = Math.max(
    1,
    detailLimit ||
      parseInt(process.env.PAYROLL_MONTH_DETAIL_LIMIT || '12', 10) ||
      12
  );
  const shouldEnrichDetail =
    includeEarningsDetail && rows.length > 0 && rows.length <= maxDetailRows;

  const flattenedRows = shouldEnrichDetail
    ? await enrichPayrollRunRowsWithEmployeeDetail({
        accessToken,
        organizationId,
        payrollRunId,
        rows,
      })
    : rows.map((row) => flattenPayrollEarningColumns(row));

  return {
    rows: flattenedRows,
    meta: {
      payroll_run_id: payrollRunId,
      processing_period: run.processing_period || `${PAYROLL_MONTH_NAMES[month - 1]} ${year}`,
      pay_period_start_date: run.pay_period_start_date || null,
      pay_period_end_date: run.pay_period_end_date || null,
      pay_date: run.pay_date || null,
      status: run.status || null,
      type: run.type || runType,
      year,
      month,
      total,
      offset,
      limit: rows.length,
      has_more: offset + rows.length < total,
      available_runs: sameMonthRuns.map((item) => ({
        payroll_run_id: item.payroll_run_id,
        type: item.type,
        status: item.status,
        processing_period: item.processing_period,
      })),
    },
  };
}

async function fetchAllEmployeesWithSalary({
  accessToken,
  organizationId,
  offset = 0,
  limit = null,
}) {
  const employees = await fetchAllEmployeePages({ accessToken, organizationId });
  const total = employees.length;
  const start = Math.min(Math.max(0, offset), total);
  const batch =
    limit != null && limit > 0
      ? employees.slice(start, start + limit)
      : employees.slice(start);
  /** Sequential by default — parallel salary calls trigger Zoho HTTP 429. */
  const concurrency = Math.max(
    1,
    parseInt(process.env.ZOHO_PAYROLL_SALARY_CONCURRENCY || '1', 10) || 1
  );
  const delayMs = Math.max(
    500,
    parseInt(process.env.ZOHO_PAYROLL_SALARY_DELAY_MS || '700', 10) || 700
  );
  const mapOne = async (emp) => {
    const id = getEmployeeRecordId(emp);
    if (!id) {
      return { ...emp, fetch_error: true, error: 'missing_employee_id' };
    }
    try {
      const salaryWrap = await fetchSalaryPayload({ accessToken, organizationId, employeeId: id });
      const salaryObj =
        salaryWrap &&
        salaryWrap.body &&
        typeof salaryWrap.body === 'object' &&
        salaryWrap.body.salary
          ? salaryWrap.body.salary
          : salaryWrap?.body;

      return {
        ...emp,
        employee_id: id,
        salary: salaryObj != null ? salaryObj : salaryWrap?.body ?? null,
      };
    } catch (err) {
      return {
        ...emp,
        employee_id: id,
        fetch_error: true,
        error: err.response?.data?.message || err.message,
      };
    }
  };

  if (concurrency <= 1) {
    const out = [];
    for (let i = 0; i < batch.length; i++) {
      if (delayMs > 0 && i > 0) await sleep(delayMs);
      out.push(await mapOne(batch[i]));
    }
    return {
      rows: out,
      total,
      offset: start,
      limit: batch.length,
      has_more: start + batch.length < total,
    };
  }

  const results = new Array(batch.length);
  let idx = 0;
  async function worker() {
    while (idx < batch.length) {
      const my = idx++;
      if (delayMs > 0) await sleep(delayMs);
      results[my] = await mapOne(batch[my]);
    }
  }
  await Promise.all(
    Array.from({ length: Math.min(concurrency, batch.length) }, () => worker())
  );
  return {
    rows: results,
    total,
    offset: start,
    limit: batch.length,
    has_more: start + batch.length < total,
  };
}

const app = express();
app.use(express.json({ limit: '50mb' }));
app.use((req, res, next) => {
  try {
    res.locals.catalyst = catalystSDK.initialize(req);
    next();
  } catch (err) {
    res.status(500).json({ success: false, error: 'Catalyst init failed' });
  }
});

function sendPayrollTableJson(res, result) {
  if (!result.found) {
    const status =
      result.reason === 'invalid_month' ? 400 : result.reason === 'not_found' ? 404 : 500;
    const available = Array.isArray(result.availableMonths) ? result.availableMonths : [];
    let error =
      result.reason === 'invalid_month'
        ? 'payroll_month must be YYYY-MM'
        : result.reason === 'not_found'
          ? 'No payroll snapshot found for this month'
          : 'Could not load payroll snapshot';
    if (available.length > 0) {
      error += `. Available months in Payroll table: ${available.join(', ')}`;
    } else if (result.rowCount === 0) {
      error += '. Payroll table appears empty — fetch and save data on the Payroll page first.';
    }
    res.status(status).json({
      success: false,
      error,
      message: error,
      reason: result.reason || 'unknown',
      availableMonths: available,
      rowCount: result.rowCount ?? 0,
    });
    return;
  }
  res.status(200).json({
    success: true,
    data: {
      rowId: result.rowId,
      payrollMonth: result.payrollMonth,
      records: result.records,
      meta: result.meta,
    },
  });
}

async function handlePayrollTableGet(req, res) {
  const payrollTableLatest = readQueryParam(req, 'payroll_table_latest') === '1';
  const payrollMonth = readQueryParam(req, 'payroll_month') || readQueryParam(req, 'payrollMonth');
  const { catalyst } = res.locals;
  const result = payrollTableLatest
    ? await loadLatestPayrollSnapshotFromTable(catalyst)
    : await loadPayrollSnapshotFromTable(catalyst, payrollMonth);
  sendPayrollTableJson(res, result);
}

async function handlePayrollGet(req, res) {
  const payrollTableLatest = readQueryParam(req, 'payroll_table_latest') === '1';
  const payrollTable = readQueryParam(req, 'payroll_table') === '1' || payrollTableLatest;
  const payrollMonth = readQueryParam(req, 'payroll_month') || readQueryParam(req, 'payrollMonth');
  const employeeId = readQueryParam(req, 'employee_id');
  const allSalaries = readQueryParam(req, 'all_salaries') === '1';
  const listEmployees = readQueryParam(req, 'list_employees') === '1';
  const payrollMonthData = readQueryParam(req, 'payroll_month_data') === '1';
  const payrunEmployeeDetail = readQueryParam(req, 'payrun_employee_detail') === '1';
  const tableMonthOnly =
    payrollMonth &&
    /^\d{4}-\d{2}$/.test(String(payrollMonth).trim()) &&
    !employeeId &&
    !allSalaries &&
    !listEmployees &&
    !payrollMonthData &&
    !payrunEmployeeDetail &&
    !payrollTableLatest;

  if ((payrollTable && (payrollMonth || payrollTableLatest)) || tableMonthOnly) {
    try {
      await handlePayrollTableGet(req, res);
    } catch (err) {
      console.error('payroll_function payroll_table error:', err.message || err);
      res.status(500).json({
        success: false,
        error: err.message || 'Failed to load payroll data',
      });
    }
    return;
  }
  return handlePayrollFetch(req, res);
}

app.get('/', handlePayrollGet);

app.get('/payroll', async (req, res) => {
  try {
    await handlePayrollTableGet(req, res);
  } catch (err) {
    console.error('payroll_function GET /payroll error:', err.message || err);
    res.status(500).json({
      success: false,
      error: err.message || 'Failed to load payroll data',
    });
  }
});

app.get('/payroll/latest', async (req, res) => {
  try {
    req.query = { ...(req.query || {}), payroll_table_latest: '1' };
    await handlePayrollTableGet(req, res);
  } catch (err) {
    console.error('payroll_function GET /payroll/latest error:', err.message || err);
    res.status(500).json({
      success: false,
      error: err.message || 'Failed to load latest payroll data',
    });
  }
});

app.post('/save', async (req, res) => {
  try {
    const { catalyst } = res.locals;
    const body = req.body && typeof req.body === 'object' ? req.body : {};
    const result = await persistPayrollSnapshot(catalyst, body);
    if (!result.saved) {
      res.status(400).json({
        success: false,
        error:
          result.reason === 'invalid_month'
            ? 'payrollMonth must be YYYY-MM'
            : result.reason === 'empty_records'
              ? 'No payroll records to save'
              : 'Could not save payroll snapshot',
        reason: result.reason || 'unknown',
      });
      return;
    }
    res.status(200).json({
      success: true,
      message: result.updated
        ? 'Payroll data updated in Payroll table.'
        : 'Payroll data saved to Payroll table.',
      data: {
        rowId: result.rowId,
        payrollMonth: result.payrollMonth,
        recordCount: result.recordCount,
        totalExpected: result.totalExpected,
        updated: result.updated === true,
      },
    });
  } catch (err) {
    console.error('payroll_function save error:', err.message || err);
    res.status(500).json({
      success: false,
      error: err.message || 'Failed to save payroll data',
    });
  }
});

module.exports = app;