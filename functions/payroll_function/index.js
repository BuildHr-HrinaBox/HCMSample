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
 * @param {IncomingMessage} req
 * @param {ServerResponse} res
 */
module.exports = async (req, res) => {
  try {
    const url = new URL(req.url, 'http://localhost');
    const employeeId = url.searchParams.get('employee_id') || process.env.ZOHO_PAYROLL_EMPLOYEE_ID || '';
    const organizationId = url.searchParams.get('organization_id') || process.env.ZOHO_PAYROLL_ORGANIZATION_ID || '';
    const apiDomain = url.searchParams.get('api_domain') || process.env.ZOHO_API_DOMAIN || 'https://www.zohoapis.in';

    if (!employeeId) {
      throw new Error('Missing employee_id. Pass ?employee_id=... or set ZOHO_PAYROLL_EMPLOYEE_ID.');
    }
    if (!organizationId) {
      throw new Error('Missing organization_id. Pass ?organization_id=... or set ZOHO_PAYROLL_ORGANIZATION_ID.');
    }

    const accessToken = await getAccessToken();
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
    console.error('payroll_function error:', {
      status: error.response?.status,
      data: zoho,
      message: error.message,
    });
    res.writeHead(error.response?.status && error.response.status < 600 ? error.response.status : 500, {
      'Content-Type': 'application/json',
    });
    res.end(
      JSON.stringify({
        success: false,
        error: zohoMsg || error.message,
        zohoCode: zoho && typeof zoho === 'object' ? zoho.code : undefined,
      })
    );
  }
};

async function getAccessToken() {
  const direct = process.env.ZOHO_ACCESS_TOKEN;
  if (direct) {
    return direct;
  }

  const refreshToken =
    process.env.ZOHO_PAYROLL_REFRESH_TOKEN ||
    process.env.ZOHO_REFRESH_TOKEN ||
    '1000.51ab555f91665626f3b3d6cad50f5e84.45016010e6caa5de13149010aa564496';
  const clientId =
    process.env.ZOHO_CLIENT_ID || '1000.0CXPBO5F75LZ8I0441HONSSQEFL27X';
  const clientSecret =
    process.env.ZOHO_CLIENT_SECRET ||
    '206cce66cc89c5076f253eff3f9cb0c74affffa4a2';

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
