'use strict';

const axios = require('axios');
const { IncomingMessage, ServerResponse } = require('http');

/** Zoho leave records API allows up to 200 per call. */
const ZOHO_APPROVED_LEAVE_PAGE_SIZE = 200;
const ZOHO_APPROVED_LEAVE_MAX_PAGES = 200;
const ZOHO_PAGE_DELAY_MS_DEFAULT = 800;

/**
 * Catalyst function to fetch approved leave records from Zoho People.
 *
 * API: GET /v2/leavetracker/leaves/records
 * Query: ?from=01-Jan-2026&to=31-Jan-2026&approvalStatus=["APPROVED"]&dataSelect=ALL
 *
 * Token (env overrides built-in defaults):
 *   ZOHO_APPROVE_LEAVE_ACCESS_TOKEN or ZOHO_LEAVE_ACCESS_TOKEN or ZOHO_ACCESS_TOKEN
 *   ZOHO_APPROVE_LEAVE_REFRESH_TOKEN or ZOHO_LEAVE_REFRESH_TOKEN + ZOHO_CLIENT_ID + ZOHO_CLIENT_SECRET
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

function clampPageSize(raw) {
  const n = parseInt(raw, 10);
  if (!Number.isFinite(n) || n < 1) return ZOHO_APPROVED_LEAVE_PAGE_SIZE;
  return Math.min(ZOHO_APPROVED_LEAVE_PAGE_SIZE, n);
}

function leaveRecordKey(record) {
  if (!record || typeof record !== 'object') return '';
  return String(
    record['Zoho.ID'] ||
      record.ZohoID ||
      record.recordId ||
      record.id ||
      record['Employee.ID'] ||
      ''
  ).trim();
}

function normalizeApprovedLeaveResponse(raw) {
  if (!raw || typeof raw !== 'object') return [];

  const candidateMaps = [
    raw.records,
    raw.response && raw.response.records,
    raw.result && raw.result.records,
    raw.data && raw.data.records,
    raw.record && typeof raw.record === 'object' ? { single: raw.record } : null,
  ];

  for (const map of candidateMaps) {
    if (!map) continue;
    if (Array.isArray(map)) {
      return map.filter((row) => row && typeof row === 'object');
    }
    if (typeof map === 'object') {
      return Object.entries(map).map(([id, row]) => ({
        recordId: String(id),
        ...(row && typeof row === 'object' ? row : {}),
      }));
    }
  }

  return [];
}

function toRecordsMap(leaveRecords) {
  const records = {};
  const arr = Array.isArray(leaveRecords) ? leaveRecords : [];
  arr.forEach((row, index) => {
    if (!row || typeof row !== 'object') return;
    const key =
      row['Zoho.ID'] ??
      row.ZohoID ??
      row.recordId ??
      row['Employee.ID'] ??
      row.id ??
      `row_${index + 1}`;
    records[String(key)] = row;
  });
  return records;
}

const DEFAULT_ZOHO_CLIENT_ID = '1000.VEO83G2Y7D16OXNQC7CN5WTK7DFHYA';
const DEFAULT_ZOHO_CLIENT_SECRET = 'b6d3935145d59974934b981d6291b40af69a3ab150';
const DEFAULT_APPROVE_LEAVE_REFRESH_TOKEN =
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
    envOr('ZOHO_APPROVE_LEAVE_REFRESH_TOKEN') ||
    envOr('ZOHO_LEAVE_REFRESH_TOKEN') ||
    envOr('ZOHO_REFRESH_TOKEN') ||
    DEFAULT_APPROVE_LEAVE_REFRESH_TOKEN;
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
      'Missing Zoho OAuth credentials. Set ZOHO_APPROVE_LEAVE_REFRESH_TOKEN + ZOHO_CLIENT_ID + ZOHO_CLIENT_SECRET in Catalyst env.'
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
      throw new Error(
        'Invalid or expired approve-leave refresh token. Regenerate OAuth token with ZOHOPEOPLE.leave.ALL scope in Zoho API Console (.in).'
      );
    }
    throw new Error(
      zoho && (zoho.error || zoho.error_description)
        ? `${zoho.error} - ${zoho.error_description || ''}`
        : err.message
    );
  }

  if (!data?.access_token) {
    const errorMsg = data?.error || data?.error_description || 'Failed to obtain access token from Zoho.';
    throw new Error(`Token refresh failed: ${errorMsg}`);
  }

  cachedAccessToken = data.access_token;
  const expiresIn = parseInt(data.expires_in, 10) || 3600;
  cachedAccessTokenExpiresAt = now + expiresIn * 1000;
  return cachedAccessToken;
}

/** Prefer refresh token (access tokens expire in ~1 hour). Do not use shared ZOHO_ACCESS_TOKEN — often expired/wrong scope. */
async function getAccessToken(options = {}) {
  const forceRefresh = options.forceRefresh === true;
  if (forceRefresh) {
    cachedAccessToken = null;
    cachedAccessTokenExpiresAt = 0;
  } else {
    const directToken = envOr('ZOHO_APPROVE_LEAVE_ACCESS_TOKEN');
    if (directToken.length > 10) {
      return directToken;
    }
  }
  return refreshAccessToken();
}

function isAuthError(status, data) {
  if (status === 401 || status === 403) return true;
  const text = JSON.stringify(data || '').toLowerCase();
  return /7213|invalid oauth|invalid token|unauthorized|access denied/.test(text);
}

function zohoAuthHeader(accessToken) {
  return `Zoho-oauthtoken ${accessToken}`;
}

async function fetchApprovedLeavePageOnce({
  accessToken,
  fromDate,
  toDate,
  startIndex = 0,
  limit = ZOHO_APPROVED_LEAVE_PAGE_SIZE,
  allowAuthRetry = true,
}) {
  const base = process.env.ZOHO_PEOPLE_BASE_URL || 'https://people.zoho.in/people/api';
  const endpoint = `${base}/v2/leavetracker/leaves/records`;

  const response = await axios.get(endpoint, {
    headers: {
      Authorization: zohoAuthHeader(accessToken),
      'Content-Type': 'application/json',
      Accept: 'application/json',
    },
    params: {
      from: fromDate,
      to: toDate,
      approvalStatus: '["APPROVED"]',
      dataSelect: 'ALL',
      startIndex: String(startIndex),
      limit: String(limit),
    },
    validateStatus: () => true,
  });

  if (response.status >= 400) {
    if (allowAuthRetry && isAuthError(response.status, response.data)) {
      const freshToken = await getAccessToken({ forceRefresh: true });
      if (freshToken && freshToken !== accessToken) {
        return fetchApprovedLeavePageOnce({
          accessToken: freshToken,
          fromDate,
          toDate,
          startIndex,
          limit,
          allowAuthRetry: false,
        });
      }
    }

    let errorMessage;
    if (typeof response.data === 'string') {
      errorMessage = `HTTP ${response.status}: ${response.data.substring(0, 200)}`;
    } else if (response.data && typeof response.data === 'object') {
      errorMessage = formatZohoApiError(response.data, `HTTP ${response.status}: ${response.statusText}`);
    } else {
      errorMessage = `HTTP ${response.status}: ${response.statusText}`;
    }

    if (response.status === 401) {
      errorMessage = `Unauthorized (401): ${errorMessage}. Regenerate OAuth with ZOHOPEOPLE.leave.ALL (or leave READ) scope at https://api-console.zoho.in/`;
    }
    throw new Error(`Zoho People API error: ${errorMessage}`);
  }

  if (!response.data) {
    throw new Error('Empty response from Zoho People API');
  }

  return response.data;
}

async function fetchAllApprovedLeaveData({ accessToken, fromDate, toDate }) {
  const pageSize = clampPageSize(
    process.env.ZOHO_APPROVED_LEAVE_PAGE_SIZE || String(ZOHO_APPROVED_LEAVE_PAGE_SIZE)
  );
  const delayMs = Math.max(
    0,
    parseInt(process.env.ZOHO_LEAVE_PAGE_DELAY_MS || String(ZOHO_PAGE_DELAY_MS_DEFAULT), 10) ||
      ZOHO_PAGE_DELAY_MS_DEFAULT
  );

  const allRecords = [];
  let startIndex = 0;
  let firstPageKey = '';

  for (let guard = 0; guard < ZOHO_APPROVED_LEAVE_MAX_PAGES; guard += 1) {
    const pageRaw = await fetchApprovedLeavePageOnce({
      accessToken,
      fromDate,
      toDate,
      startIndex,
      limit: pageSize,
    });
    const batch = normalizeApprovedLeaveResponse(pageRaw);
    if (batch.length === 0) break;

    const pageKey = leaveRecordKey(batch[0]);
    if (guard === 0) {
      firstPageKey = pageKey;
    } else if (pageKey && pageKey === firstPageKey) {
      console.warn(
        `approve_leave_function: page at startIndex=${startIndex} repeated first record; stopping pagination`
      );
      break;
    }

    allRecords.push(...batch);
    console.log(
      `approve_leave_function: page ${guard + 1} startIndex=${startIndex} batch=${batch.length} total=${allRecords.length}`
    );

    if (batch.length < pageSize) break;
    startIndex += pageSize;
    if (delayMs) await sleep(delayMs);
  }

  return {
    leaveRecords: allRecords,
    meta: {
      total: allRecords.length,
      pages: Math.ceil(allRecords.length / pageSize) || (allRecords.length ? 1 : 0),
      pageSize,
      mode: 'fetch_all',
      approvalStatus: 'APPROVED',
    },
  };
}

module.exports = async (req, res) => {
  try {
    const fromDate = readQueryParam(req, 'from') || '01-Jan-2025';
    const toDate = readQueryParam(req, 'to') || '31-Dec-2025';
    const fetchAll = readQueryParam(req, 'fetch_all') !== '0';
    const startIndex = Math.max(0, parseInt(readQueryParam(req, 'startIndex') || '0', 10) || 0);
    const limit = clampPageSize(readQueryParam(req, 'limit') || String(ZOHO_APPROVED_LEAVE_PAGE_SIZE));

    const accessToken = await getAccessToken();

    let leaveRecords;
    let meta;

    if (fetchAll) {
      const result = await fetchAllApprovedLeaveData({ accessToken, fromDate, toDate });
      leaveRecords = result.leaveRecords;
      meta = result.meta;
    } else {
      const pageRaw = await fetchApprovedLeavePageOnce({
        accessToken,
        fromDate,
        toDate,
        startIndex,
        limit,
      });
      leaveRecords = normalizeApprovedLeaveResponse(pageRaw);
      meta = {
        startIndex,
        limit,
        count: leaveRecords.length,
        has_more: leaveRecords.length >= limit,
        pageSize: limit,
        approvalStatus: 'APPROVED',
      };
    }

    const records = toRecordsMap(leaveRecords);

    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(
      JSON.stringify({
        success: true,
        records,
        leaveRecords,
        leaveTypeLabels: {},
        meta,
      })
    );
  } catch (error) {
    console.error('approve_leave_function error:', error);
    const errorMessage = formatZohoApiError(error.response?.data, error.message || 'Unknown error occurred');

    if (error.response) {
      console.error('Error response status:', error.response.status);
      console.error('Error response data:', error.response.data);
    }

    res.writeHead(500, { 'Content-Type': 'application/json' });
    res.end(
      JSON.stringify({
        success: false,
        error: errorMessage,
      })
    );
  }
};
