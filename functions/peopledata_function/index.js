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

    const source = await resolveEmployeeSource({ accessToken, formName });
    const data = await fetchPeopleRecords({
      accessToken,
      formName: source.formName,
      viewName: source.viewName,
      limit,
      sIndex,
    });
    const apiBase = data.__peopleApiBase || peopleApiBase();
    const usedForm = data.__peopleForm || source.formName;
    const usedView = data.__peopleView || source.viewName || '';
    if (data && typeof data === 'object') {
      delete data.__peopleApiBase;
      delete data.__peopleForm;
      delete data.__peopleView;
    }
    const count = extractEmployeesFromZohoPage(data).length;

    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(
      JSON.stringify({
        success: true,
        data,
        meta: {
          sIndex,
          limit,
          count,
          has_more: count >= limit,
          apiBase,
          formName: usedForm,
          viewName: usedView,
          warning: smallOrgWarning(count),
        },
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

const DEFAULT_PEOPLE_REFRESH = '1000.704677e6ac4277dbfaeff78ae2204da0.da9592dac5ffa146fd5224f77e732276';
const DEFAULT_PEOPLE_CLIENT_ID = '1000.VEO83G2Y7D16OXNQC7CN5WTK7DFHYA';
const DEFAULT_PEOPLE_CLIENT_SECRET = 'b6d3935145d59974934b981d6291b40af69a3ab150';

function envOr(name, fallback = '') {
  const v = process.env[name];
  return v != null && String(v).trim() !== '' ? String(v).trim() : fallback;
}

function peopleTokenProfiles() {
  // Always use the latest People token. Ignore Catalyst env — it still holds the
  // sample 5-employee org refresh token and would override this default.
  return [
    {
      refreshToken: DEFAULT_PEOPLE_REFRESH,
      clientId: DEFAULT_PEOPLE_CLIENT_ID,
      clientSecret: DEFAULT_PEOPLE_CLIENT_SECRET,
    },
  ];
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

/** Always mint a People access token from the People refresh token (ignore stale env access tokens). */
async function getAccessToken(options = {}) {
  if (options.forceRefresh === true) {
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

/** Official People API path. Portal UI (`/hrmsvayonaenergy/`) is not an API base — token selects the org. */
const PEOPLE_API_BASE = 'https://people.zoho.in/people/api';

function peopleApiBase() {
  const envBase = envOr('ZOHO_PEOPLE_BASE_URL', '').replace(/\/+$/, '');
  if (envBase && !/hrmsvayonaenergy/i.test(envBase)) return envBase;
  return PEOPLE_API_BASE;
}

async function zohoPeopleGet({ accessToken, path, params, allowAuthRetry = true }) {
  const url = `${peopleApiBase()}${path.startsWith('/') ? path : `/${path}`}`;
  try {
    const { data } = await axios.get(url, {
      headers: { Authorization: `Zoho-oauthtoken ${accessToken}` },
      params,
    });
    return data;
  } catch (error) {
    if (isAuthError(error)) {
      try {
        const { data } = await axios.get(url, {
          headers: { Authorization: `Bearer ${accessToken}` },
          params,
        });
        return data;
      } catch (bearerErr) {
        if (allowAuthRetry && isAuthError(bearerErr)) {
          const freshToken = await getAccessToken({ forceRefresh: true });
          if (freshToken && freshToken !== accessToken) {
            return zohoPeopleGet({
              accessToken: freshToken,
              path,
              params,
              allowAuthRetry: false,
            });
          }
        }
        throw bearerErr;
      }
    }
    throw error;
  }
}

function unwrapZohoList(payload) {
  const data = payload?.data !== undefined ? payload.data : payload;
  const result = data?.response?.result ?? data?.result ?? data;
  if (Array.isArray(result)) return result;
  return [];
}

function formLinkOf(form) {
  if (!form || typeof form !== 'object') return '';
  return String(
    form.componentName ||
      form.formLinkName ||
      form.linkName ||
      form.sys_name ||
      form.formName ||
      ''
  ).trim();
}

function viewLinkOf(view) {
  if (!view || typeof view !== 'object') return '';
  return String(
    view.viewName || view.viewLinkName || view.sys_name || view.componentName || ''
  ).trim();
}

async function listPeopleForms(accessToken) {
  try {
    const payload = await zohoPeopleGet({ accessToken, path: '/forms' });
    return unwrapZohoList(payload);
  } catch (err) {
    console.warn('peopledata_function: list forms failed:', extractZohoError(err));
    return [];
  }
}

async function listPeopleViews(accessToken, formName) {
  try {
    const payload = await zohoPeopleGet({
      accessToken,
      path: `/forms/${encodeURIComponent(formName)}/views`,
    });
    return unwrapZohoList(payload);
  } catch (err) {
    console.warn('peopledata_function: list views failed:', formName, extractZohoError(err));
    return [];
  }
}

async function fetchPeopleRecords({ accessToken, formName, limit, sIndex, viewName }) {
  try {
    const path = viewName
      ? `/views/${encodeURIComponent(viewName)}/records`
      : `/forms/${encodeURIComponent(formName)}/getRecords`;
    const data = await zohoPeopleGet({
      accessToken,
      path,
      params: { limit: String(limit), sIndex: String(sIndex) },
    });
    if (data && typeof data === 'object') {
      data.__peopleApiBase = peopleApiBase();
      data.__peopleForm = formName;
      data.__peopleView = viewName || '';
    }
    return data;
  } catch (error) {
    if (isZohoNoMoreRecordsError(error, sIndex)) {
      return wrapEmployeesAsZohoResponse([]);
    }
    throw error;
  }
}

function employeeLikeForms(forms, requestedForm) {
  const requested = String(requestedForm || 'employee').trim() || 'employee';
  const named = [];
  const seen = new Set();
  const add = (link) => {
    if (!link || seen.has(link.toLowerCase())) return;
    seen.add(link.toLowerCase());
    named.push(link);
  };
  add(requested);
  add('employee');
  add('P_Employee');
  add('Employee');
  forms.forEach((form) => {
    const link = formLinkOf(form);
    const label = String(form.labelName || form.displayName || form.formName || '');
    if (/employee/i.test(`${link} ${label}`)) add(link);
  });
  return named;
}

function preferredViewName(views) {
  if (!Array.isArray(views) || views.length === 0) return '';
  const scored = views
    .map((view) => {
      const link = viewLinkOf(view);
      const label = String(view.displayName || view.labelName || link);
      const blob = `${link} ${label}`.toLowerCase();
      let score = 0;
      if (/all/.test(blob)) score += 8;
      if (/active/.test(blob)) score += 4;
      if (/employee/.test(blob)) score += 2;
      if (/team|my /.test(blob)) score -= 5;
      return { link, score };
    })
    .filter((v) => v.link)
    .sort((a, b) => b.score - a.score);
  return scored[0]?.link || '';
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

function smallOrgWarning(count) {
  if (count > 5) return '';
  return (
    `This Zoho OAuth token only returned ${count} employee(s). ` +
    'That is the People organisation tied to the API Console login, not Vayona Energy. ' +
    'Generate a new Self Client token while logged in as zohoadmin@vayonaenergy.com ' +
    '(the same account as people.zoho.in/hrmsvayonaenergy). ' +
    'The Z_people connection in People Developer Space is not used by this Catalyst function.'
  );
}

async function resolveEmployeeSource({ accessToken, formName }) {
  const forms = await listPeopleForms(accessToken);
  const candidates = employeeLikeForms(forms, formName);
  let best = { formName: candidates[0] || 'employee', viewName: '', count: -1 };

  for (const link of candidates) {
    const views = await listPeopleViews(accessToken, link);
    const viewName = preferredViewName(views);
    const probes = viewName ? [viewName, ''] : [''];
    for (const probeView of probes) {
      try {
        const page = await fetchPeopleRecords({
          accessToken,
          formName: link,
          viewName: probeView,
          limit: 200,
          sIndex: 1,
        });
        const count = extractEmployeesFromZohoPage(page).length;
        if (count > best.count) {
          best = { formName: link, viewName: probeView, count };
        }
        if (count >= 200) return best;
      } catch (err) {
        console.warn('peopledata_function: source probe failed:', link, probeView, extractZohoError(err));
      }
    }
  }

  return best;
}

async function fetchAllPeopleRecords({ accessToken, formName, pageSize }) {
  const limit = clampPageSize(pageSize);
  const source = await resolveEmployeeSource({ accessToken, formName });
  const merged = [];
  let sIndex = 1;
  let pages = 0;
  const delayMs = Math.max(0, parseInt(process.env.ZOHO_PEOPLE_PAGE_DELAY_MS || '300', 10) || 300);

  for (let guard = 0; guard < 500; guard++) {
    const page = await fetchPeopleRecords({
      accessToken,
      formName: source.formName,
      viewName: source.viewName,
      limit,
      sIndex,
    });
    const batch = extractEmployeesFromZohoPage(page);
    pages += 1;
    merged.push(...batch);

    if (batch.length < limit) break;
    sIndex += limit;
    if (delayMs) await sleep(delayMs);
  }

  return {
    data: wrapEmployeesAsZohoResponse(merged),
    meta: {
      total: merged.length,
      pages,
      pageSize: limit,
      apiBase: peopleApiBase(),
      formName: source.formName,
      viewName: source.viewName || '',
      warning: smallOrgWarning(merged.length),
    },
  };
}