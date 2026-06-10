'use strict';

const axios = require('axios');
const { IncomingMessage, ServerResponse } = require('http');

/** Zoho bookedAndBalance returns at most 30 per call (docs); .in DC accepts up to 100. */
const ZOHO_LEAVE_PAGE_SIZE = 100;
const ZOHO_LEAVE_MAX_PAGES = 200;

/**
 * Catalyst function to fetch Zoho People Leave data.
 *
 * Token model (ZOHOPEOPLE.leave.ALL, api_domain: https://www.zohoapis.in):
 *   - Use access_token when set (ZOHO_LEAVE_ACCESS_TOKEN or ZOHO_ACCESS_TOKEN).
 *   - Else use refresh_token (ZOHO_LEAVE_REFRESH_TOKEN + ZOHO_CLIENT_ID + ZOHO_CLIENT_SECRET).
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

module.exports = async (req, res) => {
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

    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ success: true, records, leaveTypeLabels, leaveRecords, meta }));
  } catch (error) {
    console.error('leavedata_function error:', error);
    const errorMessage = formatZohoApiError(error.response?.data, error.message || 'Unknown error occurred');
    
    // Log full error details for debugging
    if (error.response) {
      console.error('Error response status:', error.response.status);
      console.error('Error response data:', error.response.data);
      console.error('Error response headers:', error.response.headers);
    }
    
    res.writeHead(500, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ 
      success: false, 
      error: errorMessage,
      details: process.env.NODE_ENV === 'development' ? {
        status: error.response?.status,
        statusText: error.response?.statusText,
        data: error.response?.data
      } : undefined
    }));
  }
};

async function getAccessToken() {
  // 1) Short-lived access_token (ZOHOPEOPLE.leave.ALL)
  const envAccessToken = process.env.ZOHO_LEAVE_ACCESS_TOKEN || process.env.ZOHO_ACCESS_TOKEN;
  if (envAccessToken && String(envAccessToken).trim().length > 10) {
    return envAccessToken.trim();
  }

  // 2) Refresh token → access_token (same client as People / Attendance)
  const refreshToken =
    process.env.ZOHO_LEAVE_REFRESH_TOKEN ||
    process.env.ZOHO_REFRESH_TOKEN ||
    '1000.8fb87613a10e02dfb9c11c8b3d44309f.d6c54b4a484ad7022c70accb35c4cef8';
  const clientId = process.env.ZOHO_CLIENT_ID || '1000.VEO83G2Y7D16OXNQC7CN5WTK7DFHYA';
  const clientSecret = process.env.ZOHO_CLIENT_SECRET || 'b6d3935145d59974934b981d6291b40af69a3ab150';

  if (!refreshToken || !clientId || !clientSecret) {
    throw new Error(
      'Missing Zoho OAuth env. Set ZOHO_LEAVE_ACCESS_TOKEN (access_token only) or ZOHO_LEAVE_REFRESH_TOKEN + ZOHO_CLIENT_ID + ZOHO_CLIENT_SECRET.'
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
      throw new Error('Invalid or expired leave refresh token. Set ZOHO_LEAVE_ACCESS_TOKEN to your access_token instead.');
    }
    throw new Error(
      zoho && (zoho.error || zoho.error_description)
        ? `${zoho.error} - ${zoho.error_description || ''}`
        : err.message
    );
  }

  if (data && data.access_token) {
    return data.access_token;
  }
  const errorMsg = data?.error || data?.error_description || 'Failed to obtain access token from Zoho.';
  throw new Error(`Token refresh failed: ${errorMsg}`);
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

async function fetchLeavePage({ accessToken, fromDate, toDate, unit, startIndex = 0, limit = ZOHO_LEAVE_PAGE_SIZE }) {
  const base = process.env.ZOHO_PEOPLE_BASE_URL || 'https://people.zoho.in/people/api';
  const endpoint = `${base}/v2/leavetracker/reports/bookedAndBalance`;

  try {
    const response = await axios.get(endpoint, {
      headers: {
        Authorization: `Bearer ${accessToken}`,
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
        errorMessage = `Unauthorized (401): ${errorMessage}. Please check: 1) Access token is valid, 2) Token has ZOHO People leave scope, 3) Token is not expired.`;
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
