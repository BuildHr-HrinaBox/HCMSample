'use strict';

const axios = require('axios');
const { IncomingMessage, ServerResponse } = require('http');

/** Zoho Creator Employee Master report (CMS Contract Management System). */
const DEFAULT_CREATOR_OWNER = 'zohoadmin_vayonaenergy';
const DEFAULT_CREATOR_APP = 'cms-contract-management-system';
const DEFAULT_CREATOR_REPORT = 'Employee_Master_Report';

const DEFAULT_PAGE_SIZE = 1000;
const ALLOWED_PAGE_SIZES = [200, 500, 1000];

let cachedCreatorAccessToken = null;
let cachedCreatorAccessTokenExpiresAt = 0;

/**
 * Fetch Employee Master data from Zoho Creator (ZohoCreator.report.READ).
 * GET .../creator/v2.1/data/{owner}/{app}/report/Employee_Master_Report?max_records=1000
 *
 * @param {IncomingMessage} req
 * @param {ServerResponse} res
 */
module.exports = async (req, res) => {
  try {
    const fetchAll = readQueryParam(req, 'fetch_all') !== '0';
    const maxRecords = clampPageSize(readQueryParam(req, 'limit') || String(DEFAULT_PAGE_SIZE));
    const target = {
      owner: readQueryParam(req, 'owner') || envOr('ZOHO_CREATOR_OWNER', DEFAULT_CREATOR_OWNER),
      app: readQueryParam(req, 'app') || envOr('ZOHO_CREATOR_APP', DEFAULT_CREATOR_APP),
      report: readQueryParam(req, 'report') || envOr('ZOHO_CREATOR_REPORT', DEFAULT_CREATOR_REPORT),
    };

    const accessToken = await getCreatorAccessToken();
    const result = await fetchEmployeeMasterReport({
      accessToken,
      ...target,
      fetchAll,
      maxRecords,
    });

    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ success: true, ...result }));
  } catch (error) {
    const httpStatus = error.response?.status;
    const detail = extractZohoError(error);
    console.error('clra_function error:', error.response?.data || error.message);
    res.writeHead(httpStatus && httpStatus >= 400 && httpStatus < 600 ? httpStatus : 500, {
      'Content-Type': 'application/json',
    });
    res.end(JSON.stringify({ success: false, error: detail, httpStatus: httpStatus || 500 }));
  }
};

function envOr(name, fallback = '') {
  const v = process.env[name];
  return v != null && String(v).trim() !== '' ? String(v).trim() : fallback;
}

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
  if (ALLOWED_PAGE_SIZES.includes(n)) return n;
  if (!Number.isFinite(n) || n < 1) return DEFAULT_PAGE_SIZE;
  if (n <= 200) return 200;
  if (n <= 500) return 500;
  return 1000;
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function getCreatorAccessToken(options = {}) {
  const forceRefresh = options.forceRefresh === true;

  if (!forceRefresh) {
    const direct =
      envOr('ZOHO_CLRA_ACCESS_TOKEN', '') ||
      envOr('ZOHO_CREATOR_ACCESS_TOKEN', '') ||
      envOr('ZOHO_ACCESS_TOKEN', '');
    if (direct.length > 10) return direct;

    const now = Date.now();
    if (cachedCreatorAccessToken && now < cachedCreatorAccessTokenExpiresAt - 60_000) {
      return cachedCreatorAccessToken;
    }
  } else {
    cachedCreatorAccessToken = null;
    cachedCreatorAccessTokenExpiresAt = 0;
  }

  const refreshToken = envOr('ZOHO_CREATOR_REFRESH_TOKEN', '') || envOr('ZOHO_REFRESH_TOKEN', '');
  const clientId = envOr('ZOHO_CLIENT_ID', '');
  const clientSecret = envOr('ZOHO_CLIENT_SECRET', '');

  if (!refreshToken || !clientId || !clientSecret) {
    const reason = forceRefresh
      ? 'The Zoho Creator access token is expired or invalid.'
      : 'No Zoho Creator OAuth credentials are configured.';
    throw new Error(
      `${reason} Set ZOHO_CLRA_ACCESS_TOKEN (access_token with ZohoCreator.report.READ) in Catalyst env for clra_function, or add ZOHO_CREATOR_REFRESH_TOKEN + ZOHO_CLIENT_ID + ZOHO_CLIENT_SECRET for auto-refresh, then redeploy.`
    );
  }

  const params = new URLSearchParams({
    refresh_token: refreshToken,
    client_id: clientId,
    client_secret: clientSecret,
    grant_type: 'refresh_token',
  });

  const { data } = await axios.post('https://accounts.zoho.in/oauth/v2/token', params, {
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
  });

  if (!data?.access_token) {
    const errorMsg = data?.error || data?.error_description || 'Failed to obtain access token from Zoho.';
    throw new Error(`Token refresh failed: ${errorMsg}`);
  }

  cachedCreatorAccessToken = data.access_token;
  const expiresIn = parseInt(data.expires_in, 10) || 3600;
  cachedCreatorAccessTokenExpiresAt = Date.now() + expiresIn * 1000;
  return cachedCreatorAccessToken;
}

function creatorApiBase() {
  return envOr('ZOHO_CREATOR_API_DOMAIN', 'https://www.zohoapis.in').replace(/\/+$/, '');
}

function creatorAuthHeader(accessToken) {
  return `Bearer ${accessToken}`;
}

function extractZohoError(error) {
  const zohoBody = error?.response?.data;
  if (typeof zohoBody === 'string' && zohoBody.trim()) return zohoBody.trim();
  if (zohoBody && typeof zohoBody === 'object') {
    return (
      zohoBody.message ||
      zohoBody.error ||
      zohoBody.description ||
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

function isCreatorSuccessBody(body) {
  if (!body || typeof body !== 'object') return false;
  const code = Number(body.code);
  if (code === 3000 || code === 9280) return true;
  return Array.isArray(body.data);
}

function extractCreatorReportRecords(apiResult) {
  if (!apiResult) return [];
  const data = apiResult.data !== undefined ? apiResult.data : apiResult;

  const candidates = [data?.data, data?.records, data?.result, data?.response?.result, data];
  for (const candidate of candidates) {
    if (Array.isArray(candidate)) return candidate;
  }

  if (data && typeof data === 'object' && !Array.isArray(data)) {
    const values = Object.values(data);
    if (values.length > 0 && values.every((v) => v && typeof v === 'object')) return values;
  }

  return [];
}

function wrapCreatorReportResponse(records) {
  return { code: 3000, data: records };
}

async function fetchCreatorReportPage({
  accessToken,
  owner,
  app,
  report,
  maxRecords,
  recordCursor,
  allowAuthRetry = true,
}) {
  const base = creatorApiBase();
  const endpoint = `${base}/creator/v2.1/data/${encodeURIComponent(owner)}/${encodeURIComponent(app)}/report/${encodeURIComponent(report)}`;

  const headers = {
    Authorization: creatorAuthHeader(accessToken),
    accept: 'application/json',
  };
  if (recordCursor) headers.record_cursor = recordCursor;

  try {
    const { data, headers: respHeaders } = await axios.get(endpoint, {
      headers,
      params: { max_records: String(maxRecords) },
    });
    const nextCursor = respHeaders?.record_cursor || respHeaders?.['record-cursor'] || null;
    return { body: data, nextCursor };
  } catch (error) {
    const body = error?.response?.data;
    if (body && isCreatorSuccessBody(body)) {
      const nextCursor = error?.response?.headers?.record_cursor || null;
      return { body, nextCursor };
    }
    if (allowAuthRetry && isAuthError(error)) {
      const freshToken = await getCreatorAccessToken({ forceRefresh: true });
      if (freshToken && freshToken !== accessToken) {
        return fetchCreatorReportPage({
          accessToken: freshToken,
          owner,
          app,
          report,
          maxRecords,
          recordCursor,
          allowAuthRetry: false,
        });
      }
    }
    throw error;
  }
}

async function fetchEmployeeMasterReport({ accessToken, owner, app, report, fetchAll, maxRecords }) {
  const merged = [];
  let recordCursor = null;
  let pages = 0;
  let lastCode = null;

  for (let guard = 0; guard < 100; guard += 1) {
    const { body, nextCursor } = await fetchCreatorReportPage({
      accessToken,
      owner,
      app,
      report,
      maxRecords,
      recordCursor,
    });
    pages += 1;
    lastCode = Number(body?.code);

    if (lastCode !== 3000 && lastCode !== 9280 && !Array.isArray(body?.data)) {
      throw new Error(body?.message || body?.description || 'Failed to load Employee Master report from Zoho Creator.');
    }

    const batch = extractCreatorReportRecords(body);
    merged.push(...batch);

    if (!fetchAll || !nextCursor || batch.length === 0) break;
    recordCursor = nextCursor;
    await sleep(200);
  }

  if (merged.length === 0 && lastCode !== 9280) {
    throw new Error('No employee records received from Zoho Creator Employee_Master_Report.');
  }

  return {
    data: wrapCreatorReportResponse(merged),
    meta: {
      total: merged.length,
      pages,
      maxRecords,
      owner,
      app,
      report,
      mode: fetchAll ? 'fetch_all' : 'single',
      source: 'zoho_creator',
    },
  };
}
