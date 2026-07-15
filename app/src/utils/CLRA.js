import React, { useState } from 'react';
import './CLRA.css';

const API_BASE = '/server/clra_function';
const CLRA_CACHE_KEY = 'clraCreatorData_v2';
const CLRA_CACHE_TTL_MS = 5 * 60 * 1000;
const CLRA_PAGE_SIZE = 1000;

/** Columns shown on the CLRA Form Fetch page (Zoho Creator Employee_Master_Report). */
export const CLRA_DISPLAY_COLUMNS = [
  'Employee_ID',
  'Employee_Name',
  'Name_as_per_aadhaar',
  'Designation_Name',
  'Department_Name',
  'Email_ID',
  'Phone_Number',
  'Employee_Status',
  'DOJ',
  'DOE',
  'Location',
  'Gender',
  'Contractor_Name',
  'UAN',
  'PAN',
  'Aadhar_Number',
  'Bank_Name',
  'IFSC_Code',
  'Account_Holder_Name',
  'Basic',
  'HRA',
  'DA',
];

const PREFERRED_COLUMNS = [
  'Employee_ID',
  'Name_as_per_aadhaar',
  'Email_ID',
  'Phone_Number',
  'Employee_Status',
  'Designation_Name',
  'Gender',
  'DOB',
  'DOJ',
  'Father_Name',
  'UAN',
  'PAN',
  'Aadhar_Number',
  'Bank_Name',
  'IFSC_Code',
  'Account_Holder_Name',
  'Basic',
  'HRA',
  'DA',
  'Location',
  'Function_Process',
  'Present_Address',
  'Permanent_Address',
  'ID',
];

let clraMemory = null;
let clraMemoryTs = 0;
let clraInflight = null;

function readClraCacheRaw() {
  if (clraMemory && Date.now() - clraMemoryTs < CLRA_CACHE_TTL_MS) {
    return clraMemory;
  }
  try {
    const cached = localStorage.getItem(CLRA_CACHE_KEY);
    if (!cached) return null;
    const parsed = JSON.parse(cached);
    if (parsed?.ts && parsed?.data && Date.now() - parsed.ts < CLRA_CACHE_TTL_MS) {
      clraMemory = parsed.data;
      clraMemoryTs = parsed.ts;
      return parsed.data;
    }
  } catch (_) {
    /* ignore */
  }
  return null;
}

export function writeClraCache(data) {
  clraMemory = data;
  clraMemoryTs = Date.now();
  try {
    localStorage.setItem(CLRA_CACHE_KEY, JSON.stringify({ ts: clraMemoryTs, data }));
  } catch (_) {
    /* ignore */
  }
}

export function clearClraCache() {
  clraMemory = null;
  clraMemoryTs = 0;
  try {
    localStorage.removeItem(CLRA_CACHE_KEY);
  } catch (_) {
    /* ignore */
  }
}

/** Flatten Zoho Creator Employee_Master_Report payloads into row objects. */
export function flattenClraRecords(apiResult) {
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

function pickField(row, ...keys) {
  for (const key of keys) {
    const val = row?.[key];
    if (val == null || val === '') continue;
    if (typeof val === 'object' && val.zc_display_value != null) return val.zc_display_value;
    if (typeof val === 'object' && val.display_value != null) return val.display_value;
    if (typeof val === 'object' && val.Location_Name != null) return val.Location_Name;
    return val;
  }
  return '';
}

export function normalizeClraEmployeeRows(records) {
  if (!Array.isArray(records)) return [];
  return records.map((row) => {
    if (!row || typeof row !== 'object') return row;
    return {
      ...row,
      Employee_ID: String(pickField(row, 'Employee_ID', 'EmployeeID', 'Employee_Id') || '').trim(),
      Name_as_per_aadhaar: String(
        pickField(row, 'Name_as_per_aadhaar', 'Name', 'Employee_Name', 'Name_as_per_Aadhaar') || ''
      ).trim(),
      Email_ID: String(pickField(row, 'Email_ID', 'EmailID', 'Email', 'email') || '').trim(),
      Phone_Number: String(pickField(row, 'Phone_Number', 'Mobile', 'Phone', 'mobile') || '').trim(),
      Location: pickField(row, 'Location') || row.Location,
    };
  });
}

export function flattenClraResponse(result) {
  if (!result) return [];
  return normalizeClraEmployeeRows(flattenClraRecords(result.data));
}

/** Map one Creator Employee_Master_Report row to Zoho People–shaped fields for statutory autofill. */
export function clraRowToAutofillEmployee(row) {
  const src = normalizeClraEmployeeRows([row])[0] || row || {};
  const scalar = (value) => {
    if (value == null || value === '') return '';
    if (typeof value === 'object') return formatCellValue(value);
    return String(value).trim();
  };
  const employeeId = scalar(pickField(src, 'Employee_ID', 'EmployeeID', 'Employee_Id'));
  const employeeName = scalar(
    pickField(src, 'Name_as_per_aadhaar', 'Employee_Name', 'Name', 'Name_as_per_Aadhaar')
  );
  const designation = scalar(pickField(src, 'Designation_Name', 'Designation'));
  const department = scalar(pickField(src, 'Department_Name', 'Department'));
  const email = scalar(pickField(src, 'Email_ID', 'EmailID', 'Email'));
  const phone = scalar(pickField(src, 'Phone_Number', 'Mobile', 'Phone'));
  const location = scalar(pickField(src, 'Location'));
  const contractor = scalar(pickField(src, 'Contractor_Name', 'Contractor'));

  return {
    ...src,
    EmployeeID: employeeId,
    Employee_Id: employeeId,
    employeeId,
    'Employee ID': employeeId,
    Full_Name: employeeName,
    Name: employeeName,
    employeeName,
    'Employee Name': employeeName,
    'Name of employee': employeeName,
    'Name of Employee': employeeName,
    FirstName: employeeName,
    Designation: designation,
    Designation_Name: designation,
    'Designation/nature of work': designation,
    Department: department,
    Department_Name: department,
    EmailID: email,
    Email_ID: email,
    Mobile: phone,
    Phone: phone,
    Phone_Number: phone,
    LocationName: location,
    Location: location,
    Contractor_Name: contractor,
    UAN: scalar(pickField(src, 'UAN')),
    PAN: scalar(pickField(src, 'PAN')),
    Aadhar_Number: scalar(pickField(src, 'Aadhar_Number', 'Aadhaar_Number')),
    Bank_Name: scalar(pickField(src, 'Bank_Name')),
    IFSC_Code: scalar(pickField(src, 'IFSC_Code', 'IFSC')),
    Bank_Account_Number: scalar(pickField(src, 'Bank_Account_Number', 'Account_Number')),
    Bank_Address: scalar(pickField(src, 'Bank_Address')),
    Account_Holder_Name: scalar(pickField(src, 'Account_Holder_Name')),
    Account_Number: scalar(pickField(src, 'Account_Number', 'Bank_Account_Number')),
    Basic: scalar(pickField(src, 'Basic')),
    HRA: scalar(pickField(src, 'HRA')),
    DA: scalar(pickField(src, 'DA')),
    DOJ: scalar(pickField(src, 'DOJ')),
    DOE: scalar(pickField(src, 'DOE', 'Date_of_Exit', 'Dateofexit', 'Date of Exit')),
    Dateofexit: scalar(pickField(src, 'DOE', 'Date_of_Exit', 'Dateofexit', 'Date of Exit')),
    DateofExit: scalar(pickField(src, 'DOE', 'Date_of_Exit', 'Dateofexit', 'Date of Exit')),
    'Date of Exit': scalar(pickField(src, 'DOE', 'Date_of_Exit', 'Dateofexit', 'Date of Exit')),
    DOB: scalar(pickField(src, 'DOB')),
    Gender: scalar(pickField(src, 'Gender')),
    Employee_Status: scalar(pickField(src, 'Employee_Status')),
    Father_Name: scalar(pickField(src, 'Father_Name')),
    Present_Address: scalar(pickField(src, 'Present_Address')),
    Permanent_Address: scalar(pickField(src, 'Permanent_Address')),
    _clraCreatorSource: true,
  };
}

/** Flatten Creator API result into autofill-ready employee objects (not Zoho People). */
export function flattenClraEmployeesForAutofill(apiResult) {
  return flattenClraResponse(apiResult).map(clraRowToAutofillEmployee);
}

async function parseClraApiResponse(response) {
  const text = await response.text();
  let json = {};
  try {
    json = text ? JSON.parse(text) : {};
  } catch {
    if (/access denied/i.test(text)) {
      throw new Error(
        'Access Denied — Zoho Creator OAuth token expired. Set ZOHO_CLRA_ACCESS_TOKEN or ZOHO_CREATOR_REFRESH_TOKEN in Catalyst env for clra_function and redeploy.'
      );
    }
    throw new Error(
      `CLRA server returned HTTP ${response.status} (not JSON). Redeploy clra_function and check Catalyst logs.`
    );
  }
  return json;
}

function formatClraFetchError(pageResult) {
  return pageResult.error || pageResult.message || 'Failed to load Employee Master data from Zoho Creator';
}

function buildClraQuery(options = {}) {
  const qs = new URLSearchParams();
  qs.set('fetch_all', options.fetchAll === false ? '0' : '1');
  const limit = Math.min(CLRA_PAGE_SIZE, Math.max(200, parseInt(options.limit, 10) || CLRA_PAGE_SIZE));
  qs.set('limit', String(limit));
  return qs;
}

async function fetchClraDataFromApi(options = {}) {
  const qs = buildClraQuery({ ...options, fetchAll: true });
  const response = await fetch(`${API_BASE}?${qs.toString()}`, {
    cache: 'no-store',
    credentials: 'include',
  });
  const result = await parseClraApiResponse(response);

  if (!response.ok || !result.success) {
    const err = new Error(formatClraFetchError(result));
    err.httpStatus = result.httpStatus || response.status;
    throw err;
  }

  const records = flattenClraResponse(result);
  if (records.length === 0) {
    throw new Error('No employee records received from Zoho Creator Employee_Master_Report.');
  }

  return {
    success: true,
    data: result.data,
    meta: { ...(result.meta || {}), total: records.length },
  };
}

/** Fetch Employee Master from Zoho Creator via clra_function. */
export function fetchClraData(options = {}) {
  const { force = false } = options;

  if (!force) {
    const cached = readClraCacheRaw();
    if (cached) return Promise.resolve(cached);
    if (clraInflight) return clraInflight;
  } else {
    clearClraCache();
    if (clraInflight) {
      return clraInflight
        .catch(() => null)
        .then(() => fetchClraData({ ...options, force: true }));
    }
  }

  clraInflight = fetchClraDataFromApi(options)
    .then((result) => {
      writeClraCache(result);
      return result;
    })
    .finally(() => {
      clraInflight = null;
    });

  return clraInflight;
}

export function getCachedClraData() {
  return readClraCacheRaw();
}

export function prefetchClraData(options = {}) {
  if (getCachedClraData()) return Promise.resolve(getCachedClraData());
  return fetchClraData(options).catch(() => null);
}

function formatCellValue(value) {
  if (value == null || value === '') return '';
  if (typeof value !== 'object') return String(value);
  if (Array.isArray(value)) {
    return value.map((item) => formatCellValue(item)).filter(Boolean).join(', ');
  }
  if (value.Location_Name != null) return String(value.Location_Name);
  if (value.display_value != null) return String(value.display_value);
  if (value.zc_display_value != null) return String(value.zc_display_value);
  if (value.name != null) return String(value.name);
  if (value.ID != null) return String(value.ID);
  return JSON.stringify(value);
}

function collectColumnKeys(records) {
  const keySet = new Set();
  records.forEach((row) => {
    if (row && typeof row === 'object') {
      Object.keys(row).forEach((k) => {
        if (!/^_|^\./.test(k)) keySet.add(k);
      });
    }
  });
  const display = CLRA_DISPLAY_COLUMNS.filter((k) => keySet.has(k));
  if (display.length > 0) return display;
  const preferred = PREFERRED_COLUMNS.filter((k) => keySet.has(k));
  const rest = [...keySet].filter((k) => !preferred.includes(k)).sort();
  return [...preferred, ...rest];
}

const CLRA = () => {
  const [data, setData] = useState(null);
  const [meta, setMeta] = useState(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const [progress, setProgress] = useState('');

  const fetchData = async () => {
    setLoading(true);
    setError('');
    setProgress('Loading Employee Master from Zoho Creator…');
    setData(null);
    setMeta(null);
    try {
      const result = await fetchClraData({ force: true });
      const merged = flattenClraResponse(result);
      setData(result.data);
      setMeta(result.meta || null);
      setProgress(`Loaded ${merged.length} employee(s)`);
    } catch (err) {
      setError(err.message || 'Failed to fetch CLRA data');
      setData(null);
      setMeta(null);
    } finally {
      setLoading(false);
      setProgress('');
    }
  };

  const records = data ? flattenClraResponse({ data, meta }) : [];
  const keys = collectColumnKeys(records);
  const gridColumns = `56px repeat(${keys.length}, minmax(130px, max-content))`;

  return (
    <div className="clra-page">
      <header className="clra-header">
        <h1 className="clra-title">CLRA</h1>
        <p className="clra-subtitle">
          Fetch employee master data from Zoho Creator — Employee_Master_Report
        </p>
      </header>

      <div className="clra-actions">
        <button
          type="button"
          className="clra-fetch-btn"
          onClick={fetchData}
          disabled={loading}
        >
          {loading ? 'Fetching...' : 'Fetch Data'}
        </button>
      </div>

      {error && <div className="clra-error">{error}</div>}

      {loading && (
        <div className="clra-loading">{progress || 'Loading employee master data...'}</div>
      )}

      {!loading && data !== null && (
        <div className="clra-content">
          {records.length === 0 ? (
            <p className="clra-empty">No records in Employee_Master_Report.</p>
          ) : (
            <div className="clra-table-wrap">
              <div className="clra-data-grid" style={{ gridTemplateColumns: gridColumns }}>
                <div className="clra-data-cell clra-data-cell--header clra-data-cell--sn">S.No</div>
                {keys.map((k) => (
                  <div key={`h-${k}`} className="clra-data-cell clra-data-cell--header" title={k}>
                    {k}
                  </div>
                ))}
                {records.map((row, i) => (
                  <React.Fragment key={i}>
                    <div className="clra-data-cell clra-data-cell--sn" data-row={i}>
                      {i + 1}
                    </div>
                    {keys.map((k) => (
                      <div
                        key={`${i}-${k}`}
                        className="clra-data-cell"
                        data-row={i}
                        title={formatCellValue(row[k])}
                      >
                        {formatCellValue(row[k])}
                      </div>
                    ))}
                  </React.Fragment>
                ))}
              </div>
            </div>
          )}
          <p className="clra-meta">
            Fetched {records.length} record(s) from Zoho Creator Employee_Master_Report.
          </p>
        </div>
      )}
    </div>
  );
};

export default CLRA;
