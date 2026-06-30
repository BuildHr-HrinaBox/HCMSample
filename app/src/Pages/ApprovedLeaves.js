import React, { useState } from 'react';
import './Leave.css';
import {
  getDefaultApprovedLeaveRange,
  fetchApprovedLeaves,
} from '../utils/approvedLeaveApi';

const PREFERRED_COLUMNS = [
  'Employee',
  'Employee.ID',
  'Employee Name',
  'Leave Type',
  'From',
  'To',
  'Days',
  'Day(s)',
  'Approval Status',
  'Reason',
  'Team Email ID',
  'Department',
  'Designation',
];

function formatCell(val) {
  if (val == null) return '';
  if (typeof val === 'object') {
    if (val.name != null) return String(val.name);
    return JSON.stringify(val);
  }
  return String(val);
}

function collectColumns(records) {
  const keySet = new Set();
  records.forEach((row) => {
    if (row && typeof row === 'object') {
      Object.keys(row).forEach((k) => {
        if (!/^_|^\./.test(k)) keySet.add(k);
      });
    }
  });
  const allKeys = Array.from(keySet);
  const ordered = PREFERRED_COLUMNS.filter((k) => keySet.has(k));
  const rest = allKeys.filter((k) => !ordered.includes(k)).sort();
  return [...ordered, ...rest];
}

const ApprovedLeaves = () => {
  const defaultRange = getDefaultApprovedLeaveRange();
  const [fromDate, setFromDate] = useState(defaultRange.from);
  const [toDate, setToDate] = useState(defaultRange.to);
  const [records, setRecords] = useState([]);
  const [meta, setMeta] = useState(null);
  const [loading, setLoading] = useState(false);
  const [loadStatus, setLoadStatus] = useState('');
  const [error, setError] = useState('');
  const [fetched, setFetched] = useState(false);

  const fetchData = async () => {
    setLoading(true);
    setError('');
    setLoadStatus('');
    setRecords([]);
    setMeta(null);
    setFetched(false);
    try {
      const result = await fetchApprovedLeaves({
        from: fromDate,
        to: toDate,
        onProgress: ({ pages, total, status }) => {
          if (status === 'loading') {
            setLoadStatus(`Loading page ${pages} (${total} records so far)...`);
          }
        },
      });
      if (!result.success) {
        throw new Error('Invalid response');
      }
      setRecords(Array.isArray(result.leaveRecords) ? result.leaveRecords : []);
      setMeta(result.meta || null);
      setFetched(true);
    } catch (err) {
      setError(err.message || 'Failed to fetch approved leave data');
      setRecords([]);
      setMeta(null);
    } finally {
      setLoading(false);
    }
  };

  const columns = collectColumns(records);

  return (
    <div className="leave-page">
      <header className="leave-header">
        <h1 className="leave-title">Approved Leaves</h1>
        <p className="leave-subtitle">
          Fetch approved leave applications from Zoho People Leave Tracker for the selected date range.
        </p>
      </header>

      <div className="leave-actions">
        <div style={{ display: 'flex', gap: 12, alignItems: 'center', flexWrap: 'wrap', marginBottom: 12 }}>
          <label className="leave-subtitle" style={{ margin: 0 }}>
            From
            <input
              type="text"
              value={fromDate}
              onChange={(e) => setFromDate(e.target.value)}
              placeholder="01-Jan-2026"
              style={{ marginLeft: 8, padding: '8px 10px', borderRadius: 6, border: '1px solid #d1d5db', width: 130 }}
            />
          </label>
          <label className="leave-subtitle" style={{ margin: 0 }}>
            To
            <input
              type="text"
              value={toDate}
              onChange={(e) => setToDate(e.target.value)}
              placeholder="31-Dec-2026"
              style={{ marginLeft: 8, padding: '8px 10px', borderRadius: 6, border: '1px solid #d1d5db', width: 130 }}
            />
          </label>
        </div>
        <button
          type="button"
          className="leave-fetch-btn"
          onClick={fetchData}
          disabled={loading}
        >
          {loading ? 'Fetching...' : 'Fetch Data'}
        </button>
      </div>

      {error && <div className="leave-error">{error}</div>}

      {loading && (
        <div className="leave-loading">
          {loadStatus || 'Loading approved leave records...'}
        </div>
      )}

      {!loading && records.length > 0 && (
        <div className="leave-content">
          <div className="leave-table-wrap">
            <table className="leave-table">
              <thead>
                <tr>
                  <th>S.No</th>
                  {columns.map((col) => (
                    <th key={col}>{col}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {records.map((row, i) => (
                  <tr key={row.recordId || row['Zoho.ID'] || i}>
                    <td>{i + 1}</td>
                    {columns.map((col) => (
                      <td key={col}>{formatCell(row[col])}</td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <p className="leave-meta">
            Fetched {records.length} approved leave record(s)
            {meta?.pages ? ` across ${meta.pages} page(s)` : ''}.
          </p>
        </div>
      )}

      {!loading && fetched && records.length === 0 && !error && (
        <p className="leave-empty">No approved leave records found for this date range.</p>
      )}
    </div>
  );
};

export default ApprovedLeaves;
