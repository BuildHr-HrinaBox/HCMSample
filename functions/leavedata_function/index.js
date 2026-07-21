'use strict';

const axios = require('axios');
const express = require('express');
const catalystSDK = require('zcatalyst-sdk-node');

/** Zoho bookedAndBalance returns at most 30 per call (docs); .in DC accepts up to 100. */
const ZOHO_LEAVE_PAGE_SIZE = 100;
const ZOHO_LEAVE_MAX_PAGES = 200;
const LEAVE_DATA_TABLE = 'LeaveData';
const LEAVE_DATA_INSERT_CHUNK = 100;
const LEAVE_DATA_DELETE_CHUNK = 200;

/**
 * Catalyst function to fetch Zoho People Leave data.
 *
 * Token model (ZOHOPEOPLE.leave.ALL, api_domain: https://www.zohoapis.in):
 *   - Prefer refresh_token (ZOHO_LEAVE_REFRESH_TOKEN + ZOHO_CLIENT_ID + ZOHO_CLIENT_SECRET).
 *   - Use ZOHO_LEAVE_ACCESS_TOKEN only as a temporary override; access tokens expire in about 1 hour.
 *
 * Query:
 *   ?from=01-Apr-2026&to=31-Mar-2027&unit=Day
 *   ?fetch_all=1 (default) — paginate until all employees are loaded
 *   ?startIndex=0&limit=100 — single page when fetch_all=0
 *
 * @param {IncomingMessage} req
 * @param {ServerResponse} res
 */
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

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function formatZohoApiError(data, fallback) {
  if (!data) return fallback;
  if (typeof data === 'string') return data;
  if (typeof data !== 'object') return fallback;
  const err = data.error;
  if (typeof err === 'string') return err;
  if (err && typeof err === 'object') {
    const msg = err.message || err.error_description || err.description;
    if (msg && err.code != null) return `[${err.code}] ${msg}`;
    if (msg) return String(msg);
  }
  return data.message || data.error_description || data.msg || fallback;
}

function clampLeavePageSize(raw) {
  const n = parseInt(raw, 10);
  if (!Number.isFinite(n) || n < 1) return ZOHO_LEAVE_PAGE_SIZE;
  return Math.min(ZOHO_LEAVE_PAGE_SIZE, n);
}

function leaveEmployeeKey(record) {
  if (!record || typeof record !== 'object') return '';
  const emp = record.employee;
  if (emp && typeof emp === 'object') {
    return String(emp.id || emp.erecno || emp.zoho_id || '').trim();
  }
  return String(record.employeeId || record['Employee.ID'] || record.id || '').trim();
}

function mergeLeavePages(pages) {
  const merged = { leavetypes: {}, report: {}, employees: [] };
  for (const page of pages) {
    if (!page || typeof page !== 'object') continue;
    if (page.leavetypes && typeof page.leavetypes === 'object') {
      Object.assign(merged.leavetypes, page.leavetypes);
    }
    if (page.report && typeof page.report === 'object') {
      Object.assign(merged.report, page.report);
    }
    if (Array.isArray(page.employees)) {
      merged.employees.push(...page.employees);
    }
  }
  return merged;
}

async function handleLeaveFetch(req, res) {
  try {
    const fromDate = readQueryParam(req, 'from') || '01-Jan-2025';
    const toDate = readQueryParam(req, 'to') || '31-Dec-2025';
    const unit = readQueryParam(req, 'unit') || 'Day';
    const fetchAll = readQueryParam(req, 'fetch_all') !== '0';
    const startIndex = Math.max(0, parseInt(readQueryParam(req, 'startIndex') || '0', 10) || 0);
    const limit = clampLeavePageSize(readQueryParam(req, 'limit') || String(ZOHO_LEAVE_PAGE_SIZE));

    const accessToken = await getAccessToken();
    console.log('Access token obtained, length:', accessToken ? accessToken.length : 0);

    let rawData;
    let meta;

    if (fetchAll) {
      const result = await fetchAllLeaveData({ accessToken, fromDate, toDate, unit });
      rawData = result.data;
      meta = result.meta;
    } else {
      rawData = await fetchLeavePage({ accessToken, fromDate, toDate, unit, startIndex, limit });
      const count = Object.keys(rawData.report || {}).length;
      meta = {
        startIndex,
        limit,
        count,
        has_more: count >= limit,
        pageSize: limit,
      };
    }

    const leaveTypeLabels = extractLeaveTypeLabels(rawData);
    const leaveRecords = renameLeaveRecordKeys(normalizeLeaveResponse(rawData), leaveTypeLabels);
    const records = toRecordsMap(rawData, leaveRecords);

    res.status(200).json({ success: true, records, leaveTypeLabels, leaveRecords, meta });
  } catch (error) {
    console.error('leavedata_function error:', error);
    const errorMessage = formatZohoApiError(error.response?.data, error.message || 'Unknown error occurred');

    if (error.response) {
      console.error('Error response status:', error.response.status);
      console.error('Error response data:', error.response.data);
      console.error('Error response headers:', error.response.headers);
    }

    res.status(500).json({
      success: false,
      error: errorMessage,
      details: process.env.NODE_ENV === 'development' ? {
        status: error.response?.status,
        statusText: error.response?.statusText,
        data: error.response?.data
      } : undefined
    });
  }
}

function toNullIfEmpty(value) {
  if (value == null) return null;
  const str = String(value).trim();
  return str === '' ? null : str;
}

function cellToStoreValue(value) {
  if (value == null || value === '') return null;
  if (typeof value === 'object') {
    try {
      return JSON.stringify(value);
    } catch (_) {
      return String(value);
    }
  }
  const str = String(value).trim();
  return str === '' ? null : str;
}

function mapIncomingLeaveDataRow(row) {
  if (!row || typeof row !== 'object') return null;
  const mapped = {
    Employee: toNullIfEmpty(row.Employee ?? row.employee),
    LeaveearnedduringthePeriod: cellToStoreValue(
      row.LeaveearnedduringthePeriod ?? row.leaveEarnedDuringthePeriod
    ),
    LeaveavailedduringthePeriod: cellToStoreValue(
      row.LeaveavailedduringthePeriod ?? row.leaveAvailedDuringthePeriod
    ),
    Totals: cellToStoreValue(row.Totals ?? row.totals),
    LegacyEarnedLeave: cellToStoreValue(row.LegacyEarnedLeave ?? row.legacyEarnedLeave),
    PaternityLeave: cellToStoreValue(row.PaternityLeave ?? row.paternityLeave),
    Absent: cellToStoreValue(row.Absent ?? row.absent),
    ContingencyLeave: cellToStoreValue(row.ContingencyLeave ?? row.contingencyLeave),
    MaternityLeave: cellToStoreValue(row.MaternityLeave ?? row.maternityLeave),
    Earnedleave: cellToStoreValue(row.Earnedleave ?? row.EarnedLeave ?? row.earnedleave),
    MonthWise: toNullIfEmpty(row.MonthWise ?? row.monthWise ?? row.month_wise),
  };
  if (!mapped.Employee) return null;
  return mapped;
}

function escapeZcqlLiteral(value) {
  return String(value || '').replace(/'/g, "''");
}

async function deleteLeaveDataRowsForMonth(catalyst, monthWise) {
  const month = String(monthWise || '').trim();
  if (!month) return deleteAllLeaveDataRows(catalyst);

  const table = catalyst.datastore().table(LEAVE_DATA_TABLE);
  const zcql = catalyst.zcql();
  let ids = [];

  try {
    const safe = escapeZcqlLiteral(month);
    const rows = await zcql.executeZCQLQuery(
      `SELECT ROWID FROM ${LEAVE_DATA_TABLE} WHERE MonthWise = '${safe}'`
    );
    ids = (Array.isArray(rows) ? rows : [])
      .map((entry) => {
        const cell = entry[LEAVE_DATA_TABLE] || entry;
        return cell?.ROWID ?? cell?.rowid ?? null;
      })
      .filter((id) => id != null);
  } catch (err) {
    console.warn('LeaveData month delete query failed:', err.message || err);
    return 0;
  }

  if (ids.length === 0) return 0;

  for (let i = 0; i < ids.length; i += LEAVE_DATA_DELETE_CHUNK) {
    const chunk = ids.slice(i, i + LEAVE_DATA_DELETE_CHUNK);
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

async function deleteAllLeaveDataRows(catalyst) {
  const table = catalyst.datastore().table(LEAVE_DATA_TABLE);
  const zcql = catalyst.zcql();
  let ids = [];

  try {
    const rows = await zcql.executeZCQLQuery(`SELECT ROWID FROM ${LEAVE_DATA_TABLE}`);
    ids = (Array.isArray(rows) ? rows : [])
      .map((entry) => {
        const cell = entry[LEAVE_DATA_TABLE] || entry;
        return cell?.ROWID ?? cell?.rowid ?? null;
      })
      .filter((id) => id != null);
  } catch (err) {
    console.warn('LeaveData ZCQL select failed, trying getAllRows:', err.message || err);
    try {
      const rows = await table.getAllRows();
      ids = (Array.isArray(rows) ? rows : [])
        .map((row) => row?.ROWID ?? row?.rowid ?? null)
        .filter((id) => id != null);
    } catch (getAllErr) {
      console.warn('LeaveData getAllRows failed:', getAllErr.message || getAllErr);
    }
  }

  if (ids.length === 0) return 0;

  for (let i = 0; i < ids.length; i += LEAVE_DATA_DELETE_CHUNK) {
    const chunk = ids.slice(i, i + LEAVE_DATA_DELETE_CHUNK);
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

async function persistLeaveDataRows(catalyst, records, { replaceExisting = true, monthWise = null } = {}) {
  const month = String(monthWise || '').trim() || null;
  const list = Array.isArray(records)
    ? records
        .map((row) => {
          const mapped = mapIncomingLeaveDataRow(row);
          if (!mapped) return null;
          if (!mapped.MonthWise && month) mapped.MonthWise = month;
          return mapped;
        })
        .filter(Boolean)
    : [];
  if (list.length === 0) {
    return { saved: false, reason: 'empty_records' };
  }

  let deleted = 0;
  if (replaceExisting) {
    deleted = month
      ? await deleteLeaveDataRowsForMonth(catalyst, month)
      : await deleteAllLeaveDataRows(catalyst);
  }

  const table = catalyst.datastore().table(LEAVE_DATA_TABLE);
  let inserted = 0;
  for (let i = 0; i < list.length; i += LEAVE_DATA_INSERT_CHUNK) {
    const chunk = list.slice(i, i + LEAVE_DATA_INSERT_CHUNK);
    if (chunk.length === 0) continue;
    await table.insertRows(chunk);
    inserted += chunk.length;
  }

  return {
    saved: true,
    inserted,
    deleted,
    total: inserted,
    monthWise: month,
  };
}

const DEFAULT_ZOHO_CLIENT_ID = '1000.VEO83G2Y7D16OXNQC7CN5WTK7DFHYA';
const DEFAULT_ZOHO_CLIENT_SECRET = 'b6d3935145d59974934b981d6291b40af69a3ab150';
const DEFAULT_LEAVE_REFRESH_TOKEN =
  '1000.6ed75bd62c19f8d1cc99fd351d25c6b8.9b2f018e582a8032e46d806cbf6d22e8';

let cachedAccessToken = null;
let cachedAccessTokenExpiresAt = 0;

function envOr(name, fallback = '') {
  const value = process.env[name];
  if (value != null && String(value).trim() !== '') return String(value).trim();
  return fallback;
}

function oauthCredentials() {
  const refreshToken =
    envOr('ZOHO_LEAVE_REFRESH_TOKEN') ||
    envOr('ZOHO_REFRESH_TOKEN') ||
    DEFAULT_LEAVE_REFRESH_TOKEN;
  const clientId = envOr('ZOHO_CLIENT_ID', DEFAULT_ZOHO_CLIENT_ID);
  const clientSecret = envOr('ZOHO_CLIENT_SECRET', DEFAULT_ZOHO_CLIENT_SECRET);
  return { refreshToken, clientId, clientSecret };
}

async function refreshAccessToken() {
  const now = Date.now();
  if (cachedAccessToken && now < cachedAccessTokenExpiresAt - 60_000) {
    return cachedAccessToken;
  }

  const { refreshToken, clientId, clientSecret } = oauthCredentials();
  if (!refreshToken || !clientId || !clientSecret) {
    throw new Error(
      'Missing Zoho OAuth env. Set ZOHO_LEAVE_REFRESH_TOKEN + ZOHO_CLIENT_ID + ZOHO_CLIENT_SECRET.'
    );
  }

  const params = new URLSearchParams({
    refresh_token: refreshToken,
    client_id: clientId,
    client_secret: clientSecret,
    grant_type: 'refresh_token',
  });

  let data;
  try {
    const res = await axios.post('https://accounts.zoho.in/oauth/v2/token', params, {
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    });
    data = res.data;
  } catch (err) {
    const zoho = err.response && err.response.data;
    if (zoho && (zoho.error === 'invalid_code' || zoho.error === 'invalid_token')) {
      throw new Error('Invalid or expired leave refresh token. Regenerate OAuth with ZOHOPEOPLE.leave.ALL scope in Zoho API Console (.in).');
    }
    throw new Error(
      zoho && (zoho.error || zoho.error_description)
        ? `${zoho.error} - ${zoho.error_description || ''}`
        : err.message
    );
  }

  if (data && data.access_token) {
    cachedAccessToken = data.access_token;
    const expiresIn = parseInt(data.expires_in, 10) || 3600;
    cachedAccessTokenExpiresAt = now + expiresIn * 1000;
    return cachedAccessToken;
  }
  const errorMsg = data?.error || data?.error_description || 'Failed to obtain access token from Zoho.';
  throw new Error(`Token refresh failed: ${errorMsg}`);
}

async function getAccessToken(options = {}) {
  const forceRefresh = options.forceRefresh === true;
  if (forceRefresh) {
    cachedAccessToken = null;
    cachedAccessTokenExpiresAt = 0;
    return refreshAccessToken();
  }

  const { refreshToken, clientId, clientSecret } = oauthCredentials();
  if (refreshToken && clientId && clientSecret) {
    return refreshAccessToken();
  }

  const directToken = envOr('ZOHO_LEAVE_ACCESS_TOKEN') || envOr('ZOHO_ACCESS_TOKEN');
  if (directToken.length > 10) {
    return directToken;
  }

  throw new Error(
    'Missing Zoho OAuth env. Set ZOHO_LEAVE_REFRESH_TOKEN + ZOHO_CLIENT_ID + ZOHO_CLIENT_SECRET.'
  );
}

function isAuthError(status, data) {
  if (status === 401 || status === 403) return true;
  const text = JSON.stringify(data || '').toLowerCase();
  return /7213|invalid oauth|invalid token|unauthorized|access denied/.test(text);
}

function zohoAuthHeader(accessToken) {
  return `Zoho-oauthtoken ${accessToken}`;
}

/**
 * Normalize Zoho Leave API response into a single array of record objects for the frontend.
 * Handles: array of IDs + details keyed by id, or nested leaveReport/userReport arrays.
 */
function normalizeLeaveResponse(raw) {
  if (!raw || typeof raw !== 'object') return [];

  const findArrayOfObjects = (obj, depth) => {
    if (depth > 4 || !obj) return null;
    if (Array.isArray(obj) && obj.length > 0) {
      const first = obj[0];
      if (first != null && typeof first === 'object' && !Array.isArray(first) && Object.keys(first).length > 0) {
        return obj;
      }
      return null;
    }
    if (typeof obj !== 'object') return null;
    for (const key of ['leaveReport', 'userReport', 'leaveDetails', 'report', 'records', 'result', 'details']) {
      const val = obj[key];
      if (!val) continue;
      if (Array.isArray(val) && val.length > 0 && val[0] != null && typeof val[0] === 'object' && Object.keys(val[0]).length > 0) {
        return val;
      }
      const found = findArrayOfObjects(val, depth + 1);
      if (found) return found;
    }
    const vals = Object.values(obj);
    for (const v of vals) {
      const found = findArrayOfObjects(v, depth + 1);
      if (found) return found;
    }
    return null;
  };

  const arr = findArrayOfObjects(raw, 0);
  if (arr && arr.length > 0) return arr;

  // Fallback: result might be array of IDs and details live in an object keyed by id
  const res = raw.response || raw.result || raw;
  const idList = Array.isArray(res) ? res : (res && Array.isArray(res.result) ? res.result : null);
  if (idList && idList.length > 0 && (typeof idList[0] === 'string' || typeof idList[0] === 'number')) {
    const parent = raw.response || raw.result || raw;
    const root = raw;

    const candidateMaps = [];
    const pushMap = (val) => {
      if (val && typeof val === 'object' && !Array.isArray(val)) candidateMaps.push(val);
    };

    // Common places where keyed details might exist
    pushMap(parent);
    if (parent && typeof parent === 'object') {
      pushMap(parent.report);
      pushMap(parent.leaveReport);
      pushMap(parent.details);
      pushMap(parent.leaveDetails);
      pushMap(parent.resultData);
      pushMap(parent.data);
      pushMap(parent.records);
    }
    if (root && typeof root === 'object') {
      pushMap(root.report);
      pushMap(root.leaveReport);
      pushMap(root.details);
      pushMap(root.leaveDetails);
      pushMap(root.resultData);
      pushMap(root.data);
      pushMap(root.records);
    }

    const getDetailsForId = (id) => {
      const key = String(id);
      for (const map of candidateMaps) {
        const direct = map[key] ?? map[id];
        if (direct && typeof direct === 'object' && !Array.isArray(direct)) return direct;
      }
      return null;
    };

    return idList.map((id) => ({
      employeeId: String(id),
      ...(getDetailsForId(id) || {}),
    }));
  }

  return [];
}

function extractLeaveTypeLabels(raw) {
  const out = {};
  const seen = new Set();

  const maybeStore = (id, label) => {
    const key = String(id || '').trim();
    const value = String(label || '').trim();
    if (!/^\d{6,}$/.test(key) || !value) return;
    if (!out[key]) out[key] = value;
  };

  const visit = (node, depth = 0) => {
    if (!node || depth > 6) return;
    if (Array.isArray(node)) {
      node.forEach((item) => visit(item, depth + 1));
      return;
    }
    if (typeof node !== 'object') return;
    if (seen.has(node)) return;
    seen.add(node);

    Object.entries(node).forEach(([key, value]) => {
      if (value && typeof value === 'object' && !Array.isArray(value)) {
        const label =
          value.leaveTypeName ||
          value.leave_type_name ||
          value.leavetype ||
          value.leaveType ||
          value.name ||
          value.label ||
          value.displayName;
        if (label) maybeStore(key, label);
      }
      if (Array.isArray(value)) {
        value.forEach((item) => {
          if (!item || typeof item !== 'object') return;
          const id =
            item.id ||
            item.leaveTypeId ||
            item.leave_type_id ||
            item.typeId ||
            item.TypeId;
          const label =
            item.leaveTypeName ||
            item.leave_type_name ||
            item.leavetype ||
            item.leaveType ||
            item.name ||
            item.label ||
            item.displayName;
          if (id && label) maybeStore(id, label);
        });
      }
      visit(value, depth + 1);
    });
  };

  visit(raw, 0);
  return out;
}

function renameLeaveRecordKeys(records, leaveTypeLabels) {
  const list = Array.isArray(records) ? records : [];
  return list.map((row) => {
    if (!row || typeof row !== 'object' || Array.isArray(row)) return row;
    const next = {};
    Object.entries(row).forEach(([key, value]) => {
      const mappedKey = leaveTypeLabels && leaveTypeLabels[key] ? leaveTypeLabels[key] : key;
      next[mappedKey] = value;
    });
    return next;
  });
}

function toRecordsMap(raw, leaveRecords) {
  const isObjectRecordMap = (obj) => {
    if (!obj || typeof obj !== 'object' || Array.isArray(obj)) return false;
    const values = Object.values(obj);
    if (values.length === 0) return false;
    return values.some((v) => v && typeof v === 'object' && !Array.isArray(v));
  };

  const directCandidates = [
    raw && raw.records,
    raw && raw.response && raw.response.records,
    raw && raw.result && raw.result.records,
    raw && raw.data && raw.data.records,
    raw && raw.report,
    raw && raw.leaveReport,
    raw && raw.details,
    raw && raw.leaveDetails,
  ];

  for (const candidate of directCandidates) {
    if (isObjectRecordMap(candidate)) return candidate;
  }

  const records = {};
  const arr = Array.isArray(leaveRecords) ? leaveRecords : [];
  arr.forEach((row, index) => {
    if (!row || typeof row !== 'object') return;
    const key =
      row['Zoho.ID'] ??
      row.ZohoID ??
      row.employeeId ??
      row['Employee.ID'] ??
      row.id ??
      `row_${index + 1}`;
    records[String(key)] = row;
  });
  return records;
}

async function fetchLeavePage({
  accessToken,
  fromDate,
  toDate,
  unit,
  startIndex = 0,
  limit = ZOHO_LEAVE_PAGE_SIZE,
  allowAuthRetry = true,
}) {
  const base = process.env.ZOHO_PEOPLE_BASE_URL || 'https://people.zoho.in/people/api';
  const endpoint = `${base}/v2/leavetracker/reports/bookedAndBalance`;

  try {
    const response = await axios.get(endpoint, {
      headers: {
        Authorization: zohoAuthHeader(accessToken),
        'Content-Type': 'application/json',
        'Accept': 'application/json',
      },
      params: {
        from: fromDate,
        to: toDate,
        unit: unit,
        startIndex: String(startIndex),
        limit: String(limit),
      },
      validateStatus: () => true,
    });
    
    // Check if response has error status
    if (response.status >= 400) {
      if (allowAuthRetry && isAuthError(response.status, response.data)) {
        const freshToken = await getAccessToken({ forceRefresh: true });
        if (freshToken && freshToken !== accessToken) {
          return fetchLeavePage({
            accessToken: freshToken,
            fromDate,
            toDate,
            unit,
            startIndex,
            limit,
            allowAuthRetry: false,
          });
        }
      }

      let errorMessage;
      
      // Try to extract error message from response
      if (typeof response.data === 'string') {
        // Response is a string (might be HTML or plain text)
        errorMessage = `HTTP ${response.status}: ${response.data.substring(0, 200)}`;
      } else if (response.data && typeof response.data === 'object') {
        errorMessage = formatZohoApiError(
          response.data,
          `HTTP ${response.status}: ${response.statusText}`
        );
      } else {
        errorMessage = `HTTP ${response.status}: ${response.statusText}`;
      }
      
      // Special handling for 401 errors
      if (response.status === 401) {
        console.error('401 Unauthorized - Access token may be invalid or missing required leave scope (e.g. ZOHOPEOPLE.leave.ALL)');
        console.error('Response data:', JSON.stringify(response.data));
        errorMessage = `Unauthorized (401): ${errorMessage}. Regenerate OAuth with ZOHOPEOPLE.leave.ALL scope at https://api-console.zoho.in/`;
      }
      
      throw new Error(`Zoho People API error: ${errorMessage}`);
    }
    
    // Check if response data is valid
    if (!response.data) {
      throw new Error('Empty response from Zoho People API');
    }
    
    // Log the response structure for debugging
    console.log('Zoho People API response status:', response.status);
    console.log('Zoho People API response data type:', typeof response.data);
    console.log('Zoho People API response data keys:', response.data && typeof response.data === 'object' ? Object.keys(response.data) : 'N/A');
    
    return response.data;
  } catch (error) {
    // If it's already our custom error, re-throw it
    if (error.message && error.message.startsWith('Zoho People API error:')) {
      throw error;
    }
    
    if (error.response) {
      // The request was made and the server responded with a status code
      // that falls out of the range of 2xx
      let errorMessage;
      
      if (typeof error.response.data === 'string') {
        errorMessage = `HTTP ${error.response.status}: ${error.response.data.substring(0, 200)}`;
      } else if (error.response.data && typeof error.response.data === 'object') {
        errorMessage = error.response.data.message || 
                      error.response.data.error || 
                      error.response.data.error_description ||
                      `HTTP ${error.response.status}: ${error.response.statusText}`;
      } else {
        errorMessage = `HTTP ${error.response.status}: ${error.response.statusText}`;
      }
      
      throw new Error(`Zoho People API error: ${errorMessage}`);
    } else if (error.request) {
      // The request was made but no response was received
      throw new Error('No response from Zoho People API');
    } else {
      // Something happened in setting up the request that triggered an Error
      throw new Error(`Request setup error: ${error.message}`);
    }
  }
}

async function fetchAllLeaveData({ accessToken, fromDate, toDate, unit }) {
  const pageSize = clampLeavePageSize(process.env.ZOHO_LEAVE_PAGE_SIZE || String(ZOHO_LEAVE_PAGE_SIZE));
  const delayMs = Math.max(0, parseInt(process.env.ZOHO_LEAVE_PAGE_DELAY_MS || '300', 10) || 300);
  const pages = [];
  let startIndex = 0;
  let firstPageKey = '';

  for (let guard = 0; guard < ZOHO_LEAVE_MAX_PAGES; guard += 1) {
    const page = await fetchLeavePage({
      accessToken,
      fromDate,
      toDate,
      unit,
      startIndex,
      limit: pageSize,
    });
    const batch = Object.values(page.report || {});
    if (batch.length === 0) break;

    const pageKey = leaveEmployeeKey(batch[0]);
    if (guard === 0) {
      firstPageKey = pageKey;
    } else if (pageKey && pageKey === firstPageKey) {
      console.warn(
        `leavedata_function: page at startIndex=${startIndex} repeated first employee; stopping pagination`
      );
      break;
    }

    pages.push(page);
    console.log(
      `leavedata_function: page ${pages.length} startIndex=${startIndex} batch=${batch.length} total=${pages.reduce(
        (sum, p) => sum + Object.keys(p.report || {}).length,
        0
      )}`
    );

    if (batch.length < pageSize) break;
    startIndex += pageSize;
    if (delayMs) await sleep(delayMs);
  }

  const data = mergeLeavePages(pages);
  const total = Object.keys(data.report || {}).length;
  return {
    data,
    meta: { total, pages: pages.length, pageSize, mode: 'fetch_all' },
  };
}

const app = express();
app.use(express.json({ limit: '50mb' }));
app.use((req, res, next) => {
  try {
    res.locals.catalyst = catalystSDK.initialize(req);
    next();
  } catch (err) {
    // GET fetch does not need Catalyst datastore; only fail hard for POST save.
    if (String(req.method || '').toUpperCase() === 'POST') {
      res.status(500).json({ success: false, error: 'Catalyst init failed' });
      return;
    }
    res.locals.catalyst = null;
    next();
  }
});

app.get('/', handleLeaveFetch);

app.post('/save', async (req, res) => {
  try {
    const { catalyst } = res.locals;
    if (!catalyst) {
      res.status(500).json({ success: false, error: 'Catalyst init failed' });
      return;
    }
    const body = req.body && typeof req.body === 'object' ? req.body : {};
    const records = Array.isArray(body.records) ? body.records : [];
    const replaceExisting = body.replaceExisting !== false && body.replaceExisting !== '0';
    const monthWise = body.monthWise || body.MonthWise || null;
    const result = await persistLeaveDataRows(catalyst, records, { replaceExisting, monthWise });

    if (!result.saved) {
      res.status(400).json({
        success: false,
        error: result.reason === 'empty_records' ? 'No leave records to save' : 'Could not save leave data',
        reason: result.reason || 'unknown',
      });
      return;
    }

    res.status(200).json({
      success: true,
      message: result.monthWise
        ? `Leave data saved to LeaveData table for ${result.monthWise}.`
        : 'Leave data saved to LeaveData table.',
      data: {
        inserted: result.inserted,
        deleted: result.deleted,
        total: result.total,
        monthWise: result.monthWise || null,
        from: body.from || null,
        to: body.to || null,
      },
    });
  } catch (err) {
    console.error('leavedata_function save error:', err.message || err);
    res.status(500).json({
      success: false,
      error: err.message || 'Failed to save leave data',
    });
  }
});

module.exports = app;
