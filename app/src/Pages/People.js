import React, { useState } from 'react';
import './People.css';
import { flattenZohoPeopleEmployees, fetchPeopleData } from '../utils/statutoryAutofillCache';


const PREFERRED_COLUMNS = [
  'EmployeeID',
  'FirstName',
  'LastName',
  'EmailID',
  'Mobile',
  'parent_department',
  'Dateofjoining',
  'Date_of_birth',
  'employee_status',
  'LocationName',
  'Work_location',
  'Pan_Number',
  'UAN_Number',
  'Bank_Name',
  'Account_Number',
  'IFSC_Code',
  'Branch_Name',
];

function formatCellValue(value) {
  if (value == null || value === '') return '';
  if (typeof value !== 'object') return String(value);
  if (Array.isArray(value)) {
    return value.map((item) => formatCellValue(item)).filter(Boolean).join(', ');
  }
  if (value.display_value != null) return String(value.display_value);
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
  const preferred = PREFERRED_COLUMNS.filter((k) => keySet.has(k));
  const rest = [...keySet].filter((k) => !preferred.includes(k)).sort();
  return [...preferred, ...rest];
}

const People = ({ userRole, userEmail }) => {
  const [data, setData] = useState(null);
  const [meta, setMeta] = useState(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const [progress, setProgress] = useState('');

  const fetchData = async () => {
    setLoading(true);
    setError('');
    setProgress('Loading employees…');
    setData(null);
    setMeta(null);
    try {
      const result = await fetchPeopleData({ force: true });
      const merged = flattenZohoPeopleEmployees({ data: result.data });
      setData(result.data);
      setMeta(result.meta || null);
      setProgress(`Loaded ${merged.length} employees`);
    } catch (err) {
      setError(err.message || 'Failed to fetch people data');
      setData(null);
      setMeta(null);
    } finally {
      setLoading(false);
      setProgress('');
    }
  };

  const records = data ? flattenZohoPeopleEmployees({ data }) : [];
  const keys = collectColumnKeys(records);
  const gridColumns = `56px repeat(${keys.length}, minmax(130px, max-content))`;

  return (
    <div className="people-page">
      <header className="people-header">
        <h1 className="people-title">People</h1>
        <p className="people-subtitle">Fetch and view people data from Zoho People API</p>
      </header>

      <div className="people-actions">
        <button
          type="button"
          className="people-fetch-btn"
          onClick={fetchData}
          disabled={loading}
        >
          {loading ? 'Fetching...' : 'Fetch Data'}
        </button>
      </div>

      {error && (
        <div className="people-error">
          {error}
        </div>
      )}

      {loading && (
        <div className="people-loading">
          {progress || 'Loading people data...'}
        </div>
      )}

      {!loading && data !== null && (
        <div className="people-content">
          {records.length === 0 ? (
            <p className="people-empty">No records in response. Raw data structure may differ.</p>
          ) : (
            <div className="people-table-wrap">
              <div className="people-data-grid" style={{ gridTemplateColumns: gridColumns }}>
                <div className="people-data-cell people-data-cell--header people-data-cell--sn">S.No</div>
                {keys.map((k) => (
                  <div key={`h-${k}`} className="people-data-cell people-data-cell--header" title={k}>
                    {k}
                  </div>
                ))}
                {records.map((row, i) => (
                  <React.Fragment key={i}>
                    <div className="people-data-cell people-data-cell--sn" data-row={i}>
                      {i + 1}
                    </div>
                    {keys.map((k) => (
                      <div
                        key={`${i}-${k}`}
                        className="people-data-cell"
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
          {meta?.warning ? <div className="people-error">{meta.warning}</div> : null}
          <p className="people-meta">
            Fetched {records.length} record(s)
            {meta?.formName ? ` from form ${meta.formName}` : ''}
            {meta?.viewName ? ` / view ${meta.viewName}` : ''}.
          </p>
        </div>
      )}
    </div>
  );
};

export default People;
