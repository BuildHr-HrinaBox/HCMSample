'use strict';

const axios = require('axios');
const { IncomingMessage, ServerResponse } = require('http');

/**
 * Catalyst function to fetch Zoho People data.
 * Tokens are expected from environment variables:
 *  - ZOHO_REFRESH_TOKEN
 *  - ZOHO_CLIENT_ID
 *  - ZOHO_CLIENT_SECRET
 * Refresh token should allow ZOHOPEOPLE.forms.ALL and ZOHOPEOPLE.employee.ALL (api_domain typically https://www.zohoapis.in).
 * Optionally override form/endpoint via query: ?form=employee&limit=50
 * 
 * @param {IncomingMessage} req 
 * @param {ServerResponse} res 
 */
module.exports = async (req, res) => {
  try {
    const url = new URL(req.url, 'http://localhost');
    const formName = url.searchParams.get('form') || 'employee';
    const limit = url.searchParams.get('limit') || '50';

    const accessToken = await getAccessToken();
    const data = await fetchPeopleRecords({ accessToken, formName, limit });

    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ success: true, data }));
  } catch (error) {
    console.error('peopledata_function error:', error.response?.data || error.message);
    res.writeHead(500, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ success: false, error: error.message }));
  }
};

async function getAccessToken() {
  // Prefer env vars but fall back to provided credentials to keep the
  // function usable without additional configuration. Replace these with
  // secure storage before production deployment.
  const refreshToken = process.env.ZOHO_REFRESH_TOKEN || '1000.c3023ed55e6a598bfecd433320d55941.f27c9108bf6d8e858213a4336bb345ff';
  const clientId = process.env.ZOHO_CLIENT_ID || '1000.ABC3VBH4REB9DC28WYZS3EY5AJD73B';
  const clientSecret = process.env.ZOHO_CLIENT_SECRET || 'f2fca57c9b0436dcc6fe68d0f922015569bba642a8';

  if (!refreshToken || !clientId || !clientSecret) {
    throw new Error('Missing Zoho OAuth environment variables.');
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

async function fetchPeopleRecords({ accessToken, formName, limit }) {
  // Adjust domain if your account uses a different DC
  const base = process.env.ZOHO_PEOPLE_BASE_URL || 'https://people.zoho.in/people/api';
  const endpoint = `${base}/forms/${encodeURIComponent(formName)}/getRecords`;
  const { data } = await axios.get(endpoint, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
    },
    params: {
      limit,
    },
  });
  return data;
}
