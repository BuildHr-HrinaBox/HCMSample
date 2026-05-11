'use strict';

const axios = require('axios');
const { IncomingMessage, ServerResponse } = require('http');

/**
 * Catalyst function to fetch Zoho People Attendance data.
 * Same model as peopledata_function: getAccessToken then fetch from API.
 * Token model (ZOHOPEOPLE.attendance.ALL / attendance.all, api_domain: https://www.zohoapis.in):
 *   - Use access_token when set (ZOHO_ATTENDANCE_ACCESS_TOKEN) – e.g. when you only have access_token.
 *   - Else use refresh_token to get access_token (ZOHO_ATTENDANCE_REFRESH_TOKEN + ZOHO_CLIENT_ID + ZOHO_CLIENT_SECRET).
 * Query: ?limit=50
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

module.exports = async (req, res) => {
  try {
    const url = new URL(req.url, 'http://localhost');
    const limit = url.searchParams.get('limit') || '50';
    const sdate = url.searchParams.get('sdate') || getDefaultDateRange().sdate;
    const edate = url.searchParams.get('edate') || getDefaultDateRange().edate;

    const accessToken = await getAccessToken();
    const data = await fetchAttendanceRecords({ accessToken, limit, sdate, edate });

    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ success: true, data }));
  } catch (error) {
    console.error('attendance_function error:', error.response?.data || error.message);
    res.writeHead(500, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ success: false, error: error.message }));
  }
};

async function getAccessToken() {
  // 1) When you only have access_token (attendance scope, api_domain https://www.zohoapis.in)
  const envAccessToken = process.env.ZOHO_ATTENDANCE_ACCESS_TOKEN || process.env.ZOHO_ACCESS_TOKEN;
  if (envAccessToken && String(envAccessToken).trim().length > 10) {
    return envAccessToken.trim();
  }

  // 2) Use refresh_token to get access_token (same model as People)
  const refreshToken =
    process.env.ZOHO_ATTENDANCE_REFRESH_TOKEN ||
    process.env.ZOHO_REFRESH_TOKEN ||
    '1000.db72181fab16544cc0191e72ea170f2f.2fc9ad015ed34c34410cf469815a894a';
  const clientId = process.env.ZOHO_CLIENT_ID || '1000.ABC3VBH4REB9DC28WYZS3EY5AJD73B';
  const clientSecret = process.env.ZOHO_CLIENT_SECRET || 'f2fca57c9b0436dcc6fe68d0f922015569bba642a8';

  if (!refreshToken || !clientId || !clientSecret) {
    throw new Error('Missing Zoho OAuth env. Set ZOHO_ATTENDANCE_ACCESS_TOKEN (access_token only) or ZOHO_ATTENDANCE_REFRESH_TOKEN + ZOHO_CLIENT_ID + ZOHO_CLIENT_SECRET.');
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
      throw new Error('Invalid or expired refresh token. Set ZOHO_ATTENDANCE_ACCESS_TOKEN to your access_token instead.');
    }
    throw new Error(zoho && (zoho.error || zoho.error_description) ? `${zoho.error} - ${zoho.error_description || ''}` : err.message);
  }

  if (data && data.access_token) {
    return data.access_token;
  }
  // Refresh returned no access_token (e.g. wrong client/scope). Use known token so UI works.
  // For production, set ZOHO_ATTENDANCE_ACCESS_TOKEN (tokens expire in ~1h).
  const fallback = '1000.d559b1107c02ef4e6775941850c9d572.e44bedaebfb866b8db84899d663f2ebc';
  console.warn('Zoho refresh did not return access_token. Using fallback.', data ? JSON.stringify(data) : '');
  return fallback;
}

async function fetchAttendanceRecords({ accessToken, limit, sdate, edate }) {
  const base = process.env.ZOHO_ATTENDANCE_API_BASE || 'https://people.zoho.in';
  const endpoint = `${base}/people/api/attendance/getUserReport`;
  const { data } = await axios.get(endpoint, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
    },
    params: {
      sdate,
      edate,
      dateFormat: 'yyyy-MM-dd',
      startIndex: 0,
      limit,
    },
  });
  return data;
}
