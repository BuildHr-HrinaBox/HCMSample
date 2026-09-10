'use strict';

const axios = require('axios');
const { IncomingMessage, ServerResponse } = require('http');

/** Zoho People getUserReport returns at most 100 employees per request. */
const ZOHO_ATTENDANCE_PAGE_SIZE = 100;

/**
 * Catalyst function to fetch Zoho People Attendance data.
 * Same model as peopledata_function: getAccessToken then fetch from API.
 * Token model (ZOHOPEOPLE.attendance.ALL, api_domain: https://www.zohoapis.in):
 *   Always mint an access_token from the attendance refresh_token (access_token expires in 1 hour).
 * Query:
 *   ?sdate=yyyy-MM-dd&edate=yyyy-MM-dd
 *   ?fetch_all=1 (default) – paginate until all employees are loaded
 *   ?startIndex=0 – used when fetch_all=0 for a single page
 */
function formatLocalYYYYMMDD(d) {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${y}-${m}-${day}`;
}

function getDefaultDateRange() {
  const now = new Date();
  const edate = formatLocalYYYYMMDD(now);
  const start = new Date(now.getFullYear(), now.getMonth(), 1);
  const sdate = formatLocalYYYYMMDD(start);
  return { sdate, edate };
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

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

module.exports = async (req, res) => {
  try {
    const fetchAll = readQueryParam(req, 'fetch_all') !== '0';
    const startIndex = Math.max(0, parseInt(readQueryParam(req, 'startIndex') || '0', 10) || 0);
    const sdate = readQueryParam(req, 'sdate') || getDefaultDateRange().sdate;
    const edate = readQueryParam(req, 'edate') || getDefaultDateRange().edate;

    const accessToken = await getAccessToken();

    if (fetchAll) {
      const { data, meta } = await fetchAllAttendanceRecords({ accessToken, sdate, edate });
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ success: true, data, meta }));
      return;
    }

    const page = await fetchAttendancePage({ accessToken, sdate, edate, startIndex });
    const batch = extractAttendanceRecordsFromPage(page);
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(
      JSON.stringify({
        success: true,
        data: wrapAttendanceResponse(batch),
        meta: {
          startIndex,
          pageSize: ZOHO_ATTENDANCE_PAGE_SIZE,
          count: batch.length,
          has_more: batch.length >= ZOHO_ATTENDANCE_PAGE_SIZE,
        },
      })
    );
  } catch (error) {
    console.error('attendance_function error:', error.response?.data || error.message);
    res.writeHead(500, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ success: false, error: error.message }));
  }
};

let cachedAttendanceAccessToken = null;
let cachedAttendanceAccessTokenExpiresAt = 0;

const DEFAULT_ATTENDANCE_REFRESH =
  '1000.cd4fc2bc6a91d4c6b08d8fc9a477faf3.85968ede6278b79bc4a7e72d40034efb';
const DEFAULT_ATTENDANCE_CLIENT_ID = '1000.VEO83G2Y7D16OXNQC7CN5WTK7DFHYA';
const DEFAULT_ATTENDANCE_CLIENT_SECRET = 'b6d3935145d59974934b981d6291b40af69a3ab150';

function attendanceTokenProfile() {
  // Always use the latest attendance token (scope ZOHOPEOPLE.attendance.ALL).
  // Ignore Catalyst env — a stale access_token expires in 1 hour.
  return {
    refreshToken: DEFAULT_ATTENDANCE_REFRESH,
    clientId: DEFAULT_ATTENDANCE_CLIENT_ID,
    clientSecret: DEFAULT_ATTENDANCE_CLIENT_SECRET,
  };
}

async function getAccessToken(options = {}) {
  if (options.forceRefresh === true) {
    cachedAttendanceAccessToken = null;
    cachedAttendanceAccessTokenExpiresAt = 0;
  }

  const now = Date.now();
  if (cachedAttendanceAccessToken && now < cachedAttendanceAccessTokenExpiresAt - 60_000) {
    return cachedAttendanceAccessToken;
  }

  const profile = attendanceTokenProfile();
  if (!profile.refreshToken || !profile.clientId || !profile.clientSecret) {
    throw new Error('Missing Zoho Attendance OAuth credentials.');
  }

  const params = new URLSearchParams({
    refresh_token: profile.refreshToken,
    client_id: profile.clientId,
    client_secret: profile.clientSecret,
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
      throw new Error('Invalid or expired attendance refresh token. Generate a new Self Client token with scope ZOHOPEOPLE.attendance.ALL.');
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

  cachedAttendanceAccessToken = data.access_token;
  const expiresIn = parseInt(data.expires_in, 10) || 3600;
  cachedAttendanceAccessTokenExpiresAt = now + expiresIn * 1000;
  return cachedAttendanceAccessToken;
}

function isZohoNoMoreAttendanceError(error) {
  const text = JSON.stringify(error?.response?.data || error?.message || '').toLowerCase();
  return /7024|no records found|no record found|no more record/i.test(text);
}

function attendanceEmployeeKey(record) {
  if (!record || typeof record !== 'object') return '';
  const emp = record.employeeDetails;
  const meta = emp && typeof emp === 'object' ? emp : record;
  return String(
    meta.erecno ||
      meta.id ||
      meta['mail id'] ||
      meta.emailId ||
      meta.email ||
      ''
  ).trim();
}

function assertZohoAttendancePageOk(apiResult, startIndex) {
  const status = apiResult?.status ?? apiResult?.response?.status;
  if (status === 1 || status === '1') {
    const msg =
      apiResult?.message ||
      apiResult?.errors?.message ||
      apiResult?.response?.errors?.message ||
      'Zoho attendance API error';
    if (Number(startIndex) > 0 && /no record|7024/i.test(String(msg))) {
      return wrapAttendanceResponse([]);
    }
    throw new Error(String(msg));
  }
  return apiResult;
}

async function fetchAttendancePage({
  accessToken,
  sdate,
  edate,
  startIndex,
  allowEmptyOnEnd = true,
  allowAuthRetry = true,
}) {
  const base = process.env.ZOHO_ATTENDANCE_API_BASE || 'https://people.zoho.in';
  const endpoint = `${base}/people/api/attendance/getUserReport`;
  const params = {
    sdate,
    edate,
    dateFormat: 'yyyy-MM-dd',
    startIndex: String(startIndex),
  };

  const requestWithAuth = async (token, scheme) => {
    const { data } = await axios.get(endpoint, {
      headers: { Authorization: `${scheme} ${token}` },
      params,
    });
    return assertZohoAttendancePageOk(data, startIndex);
  };

  try {
    try {
      return await requestWithAuth(accessToken, 'Zoho-oauthtoken');
    } catch (oauthErr) {
      const status = oauthErr?.response?.status;
      if (status !== 401 && status !== 403) throw oauthErr;
      return await requestWithAuth(accessToken, 'Bearer');
    }
  } catch (error) {
    const status = error?.response?.status;
    if (allowAuthRetry && (status === 401 || status === 403)) {
      const freshToken = await getAccessToken({ forceRefresh: true });
      if (freshToken && freshToken !== accessToken) {
        return fetchAttendancePage({
          accessToken: freshToken,
          sdate,
          edate,
          startIndex,
          allowEmptyOnEnd,
          allowAuthRetry: false,
        });
      }
    }
    if (allowEmptyOnEnd && isZohoNoMoreAttendanceError(error)) {
      return wrapAttendanceResponse([]);
    }
    throw error;
  }
}

function extractAttendanceRecordsFromPage(apiResult) {
  if (!apiResult) return [];
  if (Array.isArray(apiResult)) return apiResult;

  const candidates = [
    apiResult.result,
    apiResult.response?.result,
    apiResult.response?.record,
    apiResult.response?.records,
    apiResult.records,
    apiResult.record,
    apiResult.data,
  ];

  for (const candidate of candidates) {
    if (Array.isArray(candidate)) return candidate;
  }

  if (apiResult && typeof apiResult === 'object') {
    const dateKeyRegex = /^\d{4}-\d{2}-\d{2}$/;
    const keys = Object.keys(apiResult);
    if (keys.some((k) => dateKeyRegex.test(k))) {
      return [apiResult];
    }
  }

  return [];
}

function wrapAttendanceResponse(records) {
  return {
    result: records,
    message: 'Success',
    status: 0,
  };
}

async function fetchAllAttendanceRecords({ accessToken, sdate, edate }) {
  const merged = [];
  let startIndex = 0;
  let pages = 0;
  let firstPageKey = '';
  const delayMs = Math.max(0, parseInt(process.env.ZOHO_ATTENDANCE_PAGE_DELAY_MS || '80', 10) || 80);

  for (let guard = 0; guard < 500; guard++) {
    const page = await fetchAttendancePage({ accessToken, sdate, edate, startIndex });
    const batch = extractAttendanceRecordsFromPage(page);
    if (batch.length === 0) break;

    const pageKey = attendanceEmployeeKey(batch[0]);
    if (pages === 0) {
      firstPageKey = pageKey;
    } else if (pageKey && pageKey === firstPageKey) {
      console.warn(
        `attendance_function: page at startIndex=${startIndex} repeated first employee; stopping pagination`
      );
      break;
    }

    pages += 1;
    merged.push(...batch);
    console.log(`attendance_function: page ${pages} startIndex=${startIndex} batch=${batch.length} total=${merged.length}`);

    if (batch.length < ZOHO_ATTENDANCE_PAGE_SIZE) break;
    startIndex += ZOHO_ATTENDANCE_PAGE_SIZE;
    if (delayMs) await sleep(delayMs);
  }

  return {
    data: wrapAttendanceResponse(merged),
    meta: { total: merged.length, pages, pageSize: ZOHO_ATTENDANCE_PAGE_SIZE },
  };
}
