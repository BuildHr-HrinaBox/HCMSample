'use strict';

const axios = require('axios');
const { IncomingMessage, ServerResponse } = require('http');

/**
 * Catalyst function to fetch Zoho People Leave data.
 * 
 * Authentication options:
 * 1. Direct Access Token (recommended for testing):
 *    - Set ZOHO_ACCESS_TOKEN environment variable with your bearer token
 *    - Token must have ZOHOPEOPLE.leave.READ or ZOHOPEOPLE.leave.ALL scope
 * 
 * 2. Refresh Token (for production):
 *    - Set ZOHO_REFRESH_TOKEN, ZOHO_CLIENT_ID, ZOHO_CLIENT_SECRET
 *    - Refresh token must generate access tokens with leave scope (e.g. ZOHOPEOPLE.leave.ALL)
 * 
 * Query parameters:
 *  - from: Start date (default: 01-Jan-2025)
 *  - to: End date (default: 31-Dec-2025)
 *  - unit: Time unit (default: Day)
 * 
 * Example: ?from=01-Jan-2025&to=31-Dec-2025&unit=Day
 * 
 * @param {IncomingMessage} req 
 * @param {ServerResponse} res 
 */
module.exports = async (req, res) => {
  try {
    const url = new URL(req.url, 'http://localhost');
    const fromDate = url.searchParams.get('from') || '01-Jan-2025';
    const toDate = url.searchParams.get('to') || '31-Dec-2025';
    const unit = url.searchParams.get('unit') || 'Day';

    // Try to get access token - first check for direct access token, then refresh token
    const accessToken = await getAccessToken();
    console.log('Access token obtained, length:', accessToken ? accessToken.length : 0);
    
    const rawData = await fetchLeaveData({ accessToken, fromDate, toDate, unit });
    const leaveRecords = normalizeLeaveResponse(rawData);
    const records = toRecordsMap(rawData, leaveRecords);

    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ records }));
  } catch (error) {
    console.error('leavedata_function error:', error);
    const errorMessage = error.response?.data?.message || 
                        error.response?.data?.error || 
                        error.message || 
                        'Unknown error occurred';
    
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
  // Direct access token only when set in env (short-lived; prefer refresh below).
  const directAccessToken = process.env.ZOHO_ACCESS_TOKEN && String(process.env.ZOHO_ACCESS_TOKEN).trim();
  if (directAccessToken) {
    console.log('Using direct access token from ZOHO_ACCESS_TOKEN');
    return directAccessToken;
  }

  // Otherwise, use refresh token to get a new access token
  // Prefer env vars but fall back to provided credentials to keep the
  // function usable without additional configuration. Replace these with
  // secure storage before production deployment.
  // NOTE: Refresh token must have leave scope (e.g. ZOHOPEOPLE.leave.ALL); api_domain typically https://www.zohoapis.in
  const refreshToken = process.env.ZOHO_REFRESH_TOKEN || '1000.57027fa862de594094ee23dc6605f7da.db624019552faf17386c5376896bf7e4';
  const clientId = process.env.ZOHO_CLIENT_ID || '1000.ABC3VBH4REB9DC28WYZS3EY5AJD73B';
  const clientSecret = process.env.ZOHO_CLIENT_SECRET || 'f2fca57c9b0436dcc6fe68d0f922015569bba642a8';

  if (!refreshToken || !clientId || !clientSecret) {
    throw new Error('Missing Zoho OAuth environment variables. Please set ZOHO_REFRESH_TOKEN, ZOHO_CLIENT_ID, and ZOHO_CLIENT_SECRET, or set ZOHO_ACCESS_TOKEN for direct access.');
  }

  try {
    console.log('Refreshing access token...');
    const params = new URLSearchParams({
      refresh_token: refreshToken,
      client_id: clientId,
      client_secret: clientSecret,
      grant_type: 'refresh_token',
    });

    const response = await axios.post('https://accounts.zoho.in/oauth/v2/token', params, {
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
      },
    });
    
    if (!response.data || !response.data.access_token) {
      const errorMsg = response.data?.error || response.data?.error_description || 'Failed to obtain access token from Zoho.';
      console.error('Token refresh failed:', errorMsg);
      throw new Error(`Token refresh failed: ${errorMsg}`);
    }
    
    console.log('Access token refreshed successfully');
    return response.data.access_token;
  } catch (error) {
    if (error.response) {
      const errorMsg = error.response.data?.error || error.response.data?.error_description || error.message;
      console.error('Token refresh error response:', error.response.status, errorMsg);
      throw new Error(`Token refresh error: ${errorMsg}`);
    }
    console.error('Token refresh error:', error.message);
    throw new Error(`Token refresh error: ${error.message}`);
  }
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

async function fetchLeaveData({ accessToken, fromDate, toDate, unit }) {
  // Adjust domain if your account uses a different DC
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
      },
      // Always return response, don't throw on HTTP errors
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
        // Response is an object
        errorMessage = response.data.message || 
                      response.data.error || 
                      response.data.error_description ||
                      response.data.code ||
                      `HTTP ${response.status}: ${response.statusText}`;
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
