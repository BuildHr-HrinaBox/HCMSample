'use strict';

const axios = require('axios');
const { IncomingMessage, ServerResponse } = require('http');

const DEFAULT_PAGE_SIZE = 200;
const MAX_PAGE_SIZE = 200;

/**
 * Catalyst function to fetch Zoho People data.
 *
 * Query:
 *  - form=employee (default)
 *  - limit=200 (max 200 per Zoho API call)
 *  - sIndex=1 (1-based start index for pagination)
 *  - fetch_all=1 (server loops until all records are loaded)
 *
 * @param {IncomingMessage} req
 * @param {ServerResponse} res
 */
module.exports = async (req, res) => {
  try {
    const formName = readQueryParam(req, 'form') || 'employee';
    const fetchAll = readQueryParam(req, 'fetch_all') === '1';
    const limit = clampPageSize(readQueryParam(req, 'limit') || String(DEFAULT_PAGE_SIZE));
    const sIndex = Math.max(1, parseInt(readQueryParam(req, 'sIndex') || '1', 10) || 1);

    const accessToken = await getAccessToken();

    if (fetchAll) {
      const { data, meta } = await fetchAllPeopleRecords({ accessToken, formName, pageSize: limit });
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ success: true, data, meta }));
      return;
    }

    const data = await fetchPeopleRecords({ accessToken, formName, limit, sIndex });
    const count = extractEmployeesFromZohoPage(data).length;

    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(
      JSON.stringify({
        success: true,
        data,
        meta: { sIndex, limit, count, has_more: count >= limit },
      })
    );
  } catch (error) {
    const httpStatus = error.response?.status;
    const detail = extractZohoError(error);
    console.error('peopledata_function error:', error.response?.data || error.message);
    res.writeHead(httpStatus && httpStatus >= 400 && httpStatus < 600 ? httpStatus : 500, {
      'Content-Type': 'application/json',
    });
    res.end(JSON.stringify({ success: false, error: detail, httpStatus: httpStatus || 500 }));
  }
};

/** Catalyst may expose query on req.query or only on originalUrl. */
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

function clampPageSize(raw) {
  const n = parseInt(raw, 10);
  if (!Number.isFinite(n) || n < 1) return DEFAULT_PAGE_SIZE;
  return Math.min(MAX_PAGE_SIZE, n);
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

let cachedPeopleAccessToken = null;
let cachedPeopleAccessTokenExpiresAt = 0;

const DEFAULT_PEOPLE_REFRESH = '1000.03fd78a2572d754881d44af913ae91ab.401c27b0dc0710de4395bbf9eb552bb6';
const DEFAULT_PEOPLE_CLIENT_ID = '1000.VEO83G2Y7D16OXNQC7CN5WTK7DFHYA';
const DEFAULT_PEOPLE_CLIENT_SECRET = 'b6d3935145d59974934b981d6291b40af69a3ab150';
/** Legacy Catalyst defaults (refresh + client must stay paired). */
const LEGACY_PEOPLE_REFRESH = '1000.c3023ed55e6a598bfecd433320d55941.f27c9108bf6d8e858213a4336bb345ff';
const LEGACY_PEOPLE_CLIENT_ID = '1000.ABC3VBH4REB9DC28WYZS3EY5AJD73B';
const LEGACY_PEOPLE_CLIENT_SECRET = 'f2fca57c9b0436dcc6fe68d0f922015569bba642a8';

function envOr(name, fallback = '') {
  const v = process.env[name];
  return v != null && String(v).trim() !== '' ? String(v).trim() : fallback;
}

function peopleTokenProfiles() {
  const envRefresh = envOr('ZOHO_PEOPLE_REFRESH_TOKEN', '') || envOr('ZOHO_REFRESH_TOKEN', '');
  const envClientId = envOr('ZOHO_CLIENT_ID', '');
  const envClientSecret = envOr('ZOHO_CLIENT_SECRET', '');
  const profiles = [];

  if (envRefresh && envClientId && envClientSecret) {
    profiles.push({ refreshToken: envRefresh, clientId: envClientId, clientSecret: envClientSecret });
  } else if (envRefresh) {
    profiles.push({
      refreshToken: envRefresh,
      clientId: envClientId || DEFAULT_PEOPLE_CLIENT_ID,
      clientSecret: envClientSecret || DEFAULT_PEOPLE_CLIENT_SECRET,
    });
    profiles.push({
      refreshToken: envRefresh,
      clientId: LEGACY_PEOPLE_CLIENT_ID,
      clientSecret: LEGACY_PEOPLE_CLIENT_SECRET,
    });
  }

  profiles.push({
    refreshToken: DEFAULT_PEOPLE_REFRESH,
    clientId: DEFAULT_PEOPLE_CLIENT_ID,
    clientSecret: DEFAULT_PEOPLE_CLIENT_SECRET,
  });
  profiles.push({
    refreshToken: LEGACY_PEOPLE_REFRESH,
    clientId: LEGACY_PEOPLE_CLIENT_ID,
    clientSecret: LEGACY_PEOPLE_CLIENT_SECRET,
  });

  const seen = new Set();
  return profiles.filter((p) => {
    const key = `${p.refreshToken}|${p.clientId}`;
    if (!p.refreshToken || !p.clientId || !p.clientSecret || seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

async function refreshPeopleAccessToken() {
  const now = Date.now();
  if (cachedPeopleAccessToken && now < cachedPeopleAccessTokenExpiresAt - 60_000) {
    return cachedPeopleAccessToken;
  }

  const profiles = peopleTokenProfiles();
  if (profiles.length === 0) {
    throw new Error('Missing Zoho People OAuth credentials.');
  }

  let lastErr = null;
  for (const profile of profiles) {
    try {
      const params = new URLSearchParams({
        refresh_token: profile.refreshToken,
        client_id: profile.clientId,
        client_secret: profile.clientSecret,
        grant_type: 'refresh_token',
      });
      const { data } = await axios.post('https://accounts.zoho.in/oauth/v2/token', params, {
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      });
      if (!data?.access_token) {
        const errorMsg = data?.error || data?.error_description || 'Failed to obtain access token from Zoho.';
        throw new Error(`Token refresh failed: ${errorMsg}`);
      }
      cachedPeopleAccessToken = data.access_token;
      const expiresIn = parseInt(data.expires_in, 10) || 3600;
      cachedPeopleAccessTokenExpiresAt = now + expiresIn * 1000;
      return cachedPeopleAccessToken;
    } catch (err) {
      lastErr = err;
      const detail =
        err.response?.data?.error ||
        err.response?.data?.error_description ||
        err.message ||
        'refresh failed';
      console.warn('peopledata_function: token profile failed:', detail);
    }
  }

  throw lastErr || new Error('Failed to refresh Zoho People access token.');
}

/** Prefer refresh; only use ZOHO_PEOPLE_ACCESS_TOKEN (not shared ZOHO_ACCESS_TOKEN — often expired/wrong scope). */
async function getAccessToken(options = {}) {
  const forceRefresh = options.forceRefresh === true;
  if (!forceRefresh) {
    const directAccessToken = envOr('ZOHO_PEOPLE_ACCESS_TOKEN', '');
    if (directAccessToken.length > 10) {
      return directAccessToken;
    }
  } else {
    cachedPeopleAccessToken = null;
    cachedPeopleAccessTokenExpiresAt = 0;
  }
  return refreshPeopleAccessToken();
}

function extractZohoError(error) {
  const zohoBody = error?.response?.data;
  if (typeof zohoBody === 'string' && zohoBody.trim()) return zohoBody.trim();
  if (zohoBody && typeof zohoBody === 'object') {
    return (
      zohoBody.message ||
      zohoBody.error ||
      zohoBody.response?.errors?.message ||
      JSON.stringify(zohoBody)
    );
  }
  return error?.message || 'Unknown Zoho error';
}

function isAuthError(error) {
  const status = error?.response?.status;
  const text = extractZohoError(error).toLowerCase();
  return status === 401 || status === 403 || /access denied|invalid oauth|invalid token|unauthorized/.test(text);
}

function isZohoNoMoreRecordsError(error, sIndex) {
  const status = error?.response?.status;
  const text = JSON.stringify(error?.response?.data || error?.message || '').toLowerCase();
  if (/7024|no records found|no record found|record.?not.?found/i.test(text)) return true;
  if (status === 400 && Number(sIndex) > 1) return true;
  return false;
}

async function fetchPeopleRecords({ accessToken, formName, limit, sIndex, allowAuthRetry = true }) {
  const base = process.env.ZOHO_PEOPLE_BASE_URL || 'https://people.zoho.in/people/api';
  const endpoint = `${base}/forms/${encodeURIComponent(formName)}/getRecords`;
  try {
    const { data } = await axios.get(endpoint, {
      headers: {
        Authorization: `Bearer ${accessToken}`,
      },
      params: {
        limit: String(limit),
        sIndex: String(sIndex),
      },
    });
    return data;
  } catch (error) {
    if (isZohoNoMoreRecordsError(error, sIndex)) {
      return wrapEmployeesAsZohoResponse([]);
    }
    if (allowAuthRetry && isAuthError(error)) {
      const freshToken = await getAccessToken({ forceRefresh: true });
      if (freshToken && freshToken !== accessToken) {
        return fetchPeopleRecords({
          accessToken: freshToken,
          formName,
          limit,
          sIndex,
          allowAuthRetry: false,
        });
      }
    }
    throw error;
  }
}

/** Flatten one Zoho People page into employee objects (mirrors app flattenZohoPeopleEmployees). */
function extractEmployeesFromZohoPage(apiResult) {
  if (!apiResult) return [];

  const data = apiResult.data !== undefined ? apiResult.data : apiResult;
  const employees = [];

  const pushEmployee = (value) => {
    if (!value || typeof value !== 'object' || Array.isArray(value)) return;
    employees.push(value);
  };

  const pushMany = (values) => {
    if (!Array.isArray(values)) return;
    values.forEach((item) => pushEmployee(item));
  };

  const flattenNode = (node) => {
    if (!node) return;
    if (Array.isArray(node)) {
      node.forEach((item) => {
        if (!item || typeof item !== 'object') return;
        const keys = Object.keys(item);
        const looksLikeIdMap = keys.some((k) => /^\d{6,}$/.test(String(k)));
        if (looksLikeIdMap) {
          keys.forEach((key) => {
            const val = item[key];
            if (Array.isArray(val)) pushMany(val);
            else pushEmployee(val);
          });
          return;
        }
        pushEmployee(item);
      });
      return;
    }
    if (typeof node !== 'object') return;

    const keys = Object.keys(node);
    const looksLikeIdMap = keys.length > 0 && keys.every((k) => /^\d{6,}$/.test(String(k)));
    if (looksLikeIdMap) {
      keys.forEach((key) => {
        const val = node[key];
        if (Array.isArray(val)) pushMany(val);
        else pushEmployee(val);
      });
      return;
    }

    if (node.Employee) {
      pushEmployee(node);
      return;
    }

    pushEmployee(node);
  };

  const candidates = [
    data?.response?.result,
    data?.result,
    data?.response?.result?.record,
    data?.response?.record,
    data?.records,
    data?.record,
    data,
  ];

  for (const candidate of candidates) {
    const before = employees.length;
    flattenNode(candidate);
    if (employees.length > before) return employees;
  }

  if (Array.isArray(data)) return data;
  return employees;
}

function wrapEmployeesAsZohoResponse(employees) {
  return {
    response: {
      result: employees,
      status: 0,
    },
  };
}

async function fetchAllPeopleRecords({ accessToken, formName, pageSize }) {
  const limit = clampPageSize(pageSize);
  const merged = [];
  let sIndex = 1;
  let pages = 0;
  const delayMs = Math.max(0, parseInt(process.env.ZOHO_PEOPLE_PAGE_DELAY_MS || '300', 10) || 300);

  for (let guard = 0; guard < 500; guard++) {
    const page = await fetchPeopleRecords({ accessToken, formName, limit, sIndex });
    const batch = extractEmployeesFromZohoPage(page);
    pages += 1;
    merged.push(...batch);

    if (batch.length < limit) break;
    sIndex += limit;
    if (delayMs) await sleep(delayMs);
  }

  return {
    data: wrapEmployeesAsZohoResponse(merged),
    meta: { total: merged.length, pages, pageSize: limit },
  };
}
