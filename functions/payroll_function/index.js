'use strict';

const axios = require('axios');
const { IncomingMessage, ServerResponse } = require('http');

/**
 * Catalyst function to fetch Zoho Payroll data.
 *
 * Auth (Zoho Payroll OAuth — e.g. ZohoPayroll.salary.ALL):
 *   - ZOHO_ACCESS_TOKEN — optional; short-lived; testing only.
 *   - ZOHO_PAYROLL_REFRESH_TOKEN or ZOHO_REFRESH_TOKEN + ZOHO_CLIENT_ID + ZOHO_CLIENT_SECRET.
 * Prefer Catalyst env vars in production; rotate tokens if exposed.
 *
 * GET ?list_employees=1&organization_id=… — list Payroll employees (path uses Payroll employee_id, not Zoho User ID).
 * GET ?all_salaries=1&organization_id=… — list all employees (all pages) and fetch salary per employee; returns an array of merged rows.
 *
 * @param {IncomingMessage} req
 * @param {ServerResponse} res
 */
module.exports = async (req, res) => {
  try {
    const url = new URL(req.url, 'http://localhost');
    const organizationId =
      url.searchParams.get('organization_id') ||
      process.env.ZOHO_PAYROLL_ORGANIZATION_ID ||
      '60006183023';
    const apiDomain = url.searchParams.get('api_domain') || process.env.ZOHO_API_DOMAIN || 'https://www.zohoapis.in';
    const listEmployees =
      url.searchParams.get('list_employees') === '1' || url.searchParams.get('list') === '1';
    const allSalaries = url.searchParams.get('all_salaries') === '1';

    if (!organizationId) {
      throw new Error('Missing organization_id. Pass ?organization_id=... or set ZOHO_PAYROLL_ORGANIZATION_ID.');
    }

    const accessToken = await getAccessToken();

    if (listEmployees) {
      const page = url.searchParams.get('page') || '1';
      const perPage = url.searchParams.get('per_page') || '200';
      const data = await fetchPayrollEmployees({
        accessToken,
        organizationId,
        apiDomain,
        page,
        perPage,
      });
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ success: true, data }));
      return;
    }

    if (allSalaries) {
      const rows = await fetchAllEmployeesSalaries({
        accessToken,
        organizationId,
        apiDomain,
      });
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ success: true, data: rows }));
      return;
    }

    const employeeId =
      url.searchParams.get('employee_id') || process.env.ZOHO_PAYROLL_EMPLOYEE_ID || '';
    if (!employeeId) {
      throw new Error(
        'Missing Payroll employee_id (not Zoho User ID). Use ?list_employees=1 to list IDs, or pass ?employee_id=PAYROLL_EMPLOYEE_ID.'
      );
    }

    const data = await fetchPayrollSalary({
      accessToken,
      employeeId,
      organizationId,
      apiDomain,
    });

    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ success: true, data }));
  } catch (error) {
    const zoho = error.response?.data;
    const zohoMsg =
      (zoho && typeof zoho === 'object' && (zoho.message || zoho.error)) ||
      (typeof zoho === 'string' ? zoho : null);
    const status = error.response?.status;
    const detailedError = buildPayrollErrorMessage({
      status,
      zohoMsg,
      organizationId:
        (() => {
          try {
            const url = new URL(req.url, 'http://localhost');
            return (
              url.searchParams.get('organization_id') ||
              process.env.ZOHO_PAYROLL_ORGANIZATION_ID ||
              '60006183023'
            );
          } catch (_) {
            return process.env.ZOHO_PAYROLL_ORGANIZATION_ID || '60006183023';
          }
        })(),
    });
    console.error('payroll_function error:', {
      status,
      data: zoho,
      message: error.message,
    });
    res.writeHead(status && status < 600 ? status : 500, {
      'Content-Type': 'application/json',
    });
    res.end(
      JSON.stringify({
        success: false,
        error: detailedError || zohoMsg || error.message,
        zohoCode: zoho && typeof zoho === 'object' ? zoho.code : undefined,
      })
    );
  }
};

function buildPayrollErrorMessage({ status, zohoMsg, organizationId }) {
  const normalized = String(zohoMsg || '').toLowerCase().trim();
  const looksAccessDenied =
    status === 400 &&
    (normalized.includes('access denied') ||
      normalized.includes('permission') ||
      normalized.includes('unauthorized'));

  if (looksAccessDenied) {
    return `Zoho Payroll rejected this request with "Access Denied". Check that the OAuth token/refresh token has Payroll scopes, the account has access to Payroll organisation ${organizationId}, and the organisation ID belongs to the same Zoho Payroll account.`;
  }

  if (status === 400 && normalized) {
    return `Zoho Payroll request failed (${zohoMsg}). Verify payroll permissions and organisation ID ${organizationId}.`;
  }

  return zohoMsg || '';
}

async function getAccessToken() {
  const direct = process.env.ZOHO_ACCESS_TOKEN;
  if (direct) {
    return direct;
  }

  const refreshToken =
    process.env.ZOHO_PAYROLL_REFRESH_TOKEN ||
    process.env.ZOHO_REFRESH_TOKEN ||
    '1000.533b41de8339def950141ebc05d589e9.2cae3a6edbc21af027bcc42401d04d09';
  const clientId =
    process.env.ZOHO_CLIENT_ID || '1000.ABC3VBH4REB9DC28WYZS3EY5AJD73B';
  const clientSecret =
    process.env.ZOHO_CLIENT_SECRET ||
    'f2fca57c9b0436dcc6fe68d0f922015569bba642a8';

  if (!refreshToken || !clientId || !clientSecret) {
    throw new Error(
      'Missing Zoho OAuth credentials for payroll. Set ZOHO_PAYROLL_REFRESH_TOKEN (or ZOHO_REFRESH_TOKEN), ZOHO_CLIENT_ID, ZOHO_CLIENT_SECRET, or ZOHO_ACCESS_TOKEN.'
    );
  }

  const params = new URLSearchParams({
    refresh_token: refreshToken,
    client_id: clientId,
    client_secret: clientSecret,
    grant_type: 'refresh_token',
  });

  const { data } = await axios.post('https://accounts.zoho.in/oauth/v2/token', params);
  if (!data.access_token) {
    throw new Error('Failed to obtain access token from Zoho.');
  }
  return data.access_token;
}

async function fetchPayrollSalary({
  accessToken,
  employeeId,
  organizationId,
  apiDomain,
}) {
  const safeEmployeeId = encodeURIComponent(employeeId);
  const url = `${apiDomain}/payroll/v1/employees/${safeEmployeeId}/salary`;

  const { data } = await axios.get(url, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
    },
    params: {
      organization_id: organizationId,
    },
  });

  return data;
}

async function fetchPayrollEmployees({ accessToken, organizationId, apiDomain, page, perPage }) {
  const { data } = await axios.get(`${apiDomain}/payroll/v1/employees`, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
    },
    params: {
      organization_id: organizationId,
      page,
      per_page: perPage,
    },
  });

  return data;
}

function pageContextMeta(body) {
  const pc = body && body.page_context;
  if (Array.isArray(pc) && pc.length) return pc[0];
  if (pc && typeof pc === 'object') return pc;
  return {};
}

async function fetchAllEmployeesPaged({ accessToken, organizationId, apiDomain }) {
  const perPage = 200;
  let page = 1;
  const all = [];
  for (;;) {
    const body = await fetchPayrollEmployees({
      accessToken,
      organizationId,
      apiDomain,
      page: String(page),
      perPage: String(perPage),
    });
    const batch = Array.isArray(body.employees) ? body.employees : [];
    all.push(...batch);
    const meta = pageContextMeta(body);
    const hasMore = meta.has_more_page === true;
    if (!hasMore || batch.length === 0) break;
    page += 1;
  }
  return all;
}

async function fetchAllEmployeesSalaries({ accessToken, organizationId, apiDomain }) {
  const employees = await fetchAllEmployeesPaged({ accessToken, organizationId, apiDomain });
  const rows = [];
  for (const emp of employees) {
    const id = emp && emp.employee_id != null ? String(emp.employee_id) : '';
    if (!id) continue;
    try {
      const salary = await fetchPayrollSalary({
        accessToken,
        employeeId: id,
        organizationId,
        apiDomain,
      });
      const payload = salary && typeof salary === 'object' ? salary : { salary_value: salary };
      rows.push({
        ...payload,
        employee_id: id,
      });
    } catch (err) {
      const zoho = err.response && err.response.data;
      rows.push({
        employee_id: id,
        first_name: emp.first_name,
        last_name: emp.last_name,
        work_mail: emp.work_mail,
        fetch_error:
          (zoho && typeof zoho === 'object' && (zoho.message || zoho.error)) || err.message || 'Salary fetch failed',
      });
    }
  }
  return rows;
}
