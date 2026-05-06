import React, { useState } from 'react';
import './Leave.css';

const API_BASE = '/server/leavedata_function';

const Leave = ({ userRole, userEmail }) => {
  const [data, setData] = useState(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');

  const fetchData = async () => {
    setLoading(true);
    setError('');
    setData(null);
    try {
      const params = new URLSearchParams({
        from: '01-Jan-2025',
        to: '31-Dec-2025',
        unit: 'Day',
      });
      const res = await fetch(`${API_BASE}?${params.toString()}`);
      const json = await res.json();
      if (!res.ok) {
        throw new Error(json.error || json.message || 'Request failed');
      }
      if (json && typeof json === 'object') {
        const recordsFromMap =
          json.records && typeof json.records === 'object' && !Array.isArray(json.records)
            ? Object.entries(json.records).map(([id, row]) => ({
                employeeId: String(id),
                ...(row && typeof row === 'object' ? row : {}),
              }))
            : null;
        setData({
          raw: json.records ?? json.data ?? json,
          leaveRecords: recordsFromMap || (Array.isArray(json.leaveRecords) ? json.leaveRecords : null),
        });
      } else {
        throw new Error(json.error || 'Invalid response');
      }
    } catch (err) {
      setError(err.message || 'Failed to fetch leave data');
      setData(null);
    } finally {
      setLoading(false);
    }
  };

  const rawData = data && data.raw !== undefined ? data.raw : data;
  const backendRecords = data && Array.isArray(data.leaveRecords) ? data.leaveRecords : null;

  // Find array of objects that look like leave records (have multiple keys), not just IDs
  const findLeaveRecordsArray = (obj, depth = 0) => {
    if (depth > 5 || !obj) return null;
    if (Array.isArray(obj)) {
      const allObjects = obj.every(i => i != null && typeof i === 'object' && !Array.isArray(i));
      const hasKeys = obj.length > 0 && obj.some(i => i != null && typeof i === 'object' && Object.keys(i).length > 0);
      if (allObjects && hasKeys) return obj;
      return null;
    }
    if (typeof obj !== 'object') return null;
    const prefer = ['leaveReport', 'userReport', 'leaveDetails', 'details', 'records', 'record', 'result', 'report'];
    for (const key of prefer) {
      const val = obj[key];
      if (!val) continue;
      if (Array.isArray(val)) {
        const allObjects = val.every(i => i != null && typeof i === 'object' && !Array.isArray(i));
        const hasKeys = val.length > 0 && val.some(i => i != null && typeof i === 'object' && Object.keys(i).length > 0);
        if (allObjects && hasKeys) return val;
      }
      const found = findLeaveRecordsArray(val, depth + 1);
      if (found) return found;
    }
    for (const key of Object.keys(obj)) {
      const found = findLeaveRecordsArray(obj[key], depth + 1);
      if (found) return found;
    }
    return null;
  };

  // Normalize API response to array of records (use backend-normalized first, then parse raw)
  const records = (() => {
    if (backendRecords && backendRecords.length > 0) return backendRecords;
    if (!rawData) return [];
    if (Array.isArray(rawData)) {
      const allObjects = rawData.every(i => i != null && typeof i === 'object');
      if (allObjects && rawData.some(i => i != null && typeof i === 'object' && Object.keys(i).length > 0)) return rawData;
    }
    const asObjects = findLeaveRecordsArray(rawData);
    if (asObjects && asObjects.length > 0) return asObjects;
    const res = rawData.response || rawData.result || rawData.data || rawData;
    if (Array.isArray(res)) return res;
    let rec = res?.leaveReport ?? res?.userReport ?? res?.leaveDetails ?? res?.details ?? res?.record ?? res?.records;
    if (rec && typeof rec === 'object' && !Array.isArray(rec)) rec = Object.values(rec);
    if (Array.isArray(rec) && rec.length > 0) return rec;
    if (res && typeof res === 'object') {
      const arr = Object.values(res).find(Array.isArray);
      if (arr) return arr;
    }
    return [];
  })();

  // Collect keys from all records so we show columns even if first record is empty
  const keys = (() => {
    const keySet = new Set();
    records.forEach(r => {
      if (r != null && typeof r === 'object') {
        const k = Array.isArray(r) ? r.map((_, i) => String(i)) : Object.keys(r).filter(k => !/^_|^\./.test(k));
        k.forEach(keySet.add, keySet);
      }
    });
    return Array.from(keySet);
  })();

  // If no keys (e.g. empty objects or primitives), show raw data in one column
  const showRaw = records.length > 0 && keys.length === 0;

  return (
    <div className="leave-page">
      <header className="leave-header">
        <h1 className="leave-title">Leave</h1>
        <p className="leave-subtitle">Fetch and view leave data from Zoho People Leave API</p>
      </header>

      <div className="leave-actions">
        <button
          type="button"
          className="leave-fetch-btn"
          onClick={fetchData}
          disabled={loading}
        >
          {loading ? 'Fetching...' : 'Fetch Data'}
        </button>
      </div>

      {error && (
        <div className="leave-error">
          {error}
        </div>
      )}

      {loading && (
        <div className="leave-loading">
          Loading leave data...
        </div>
      )}

      {!loading && data != null && (
        <div className="leave-content">
          {records.length === 0 ? (
            <p className="leave-empty">No records in response. Raw data structure may differ.</p>
          ) : showRaw ? (
            <div className="leave-table-wrap">
              <table className="leave-table">
                <thead>
                  <tr>
                    <th>S.No</th>
                    <th>Response</th>
                  </tr>
                </thead>
                <tbody>
                  {records.map((row, i) => (
                    <tr key={i}>
                      <td>{i + 1}</td>
                      <td className="leave-raw-cell">
                        <pre>{typeof row === 'object' ? JSON.stringify(row, null, 2) : String(row)}</pre>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          ) : (
            <div className="leave-table-wrap">
              <table className="leave-table">
                <thead>
                  <tr>
                    <th>S.No</th>
                    {keys.map(k => (
                      <th key={k}>{k}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {records.map((row, i) => (
                    <tr key={i}>
                      <td>{i + 1}</td>
                      {keys.map(k => (
                        <td key={k}>
                          {row[k] != null && typeof row[k] === 'object'
                            ? JSON.stringify(row[k])
                            : String(row[k] ?? '')}
                        </td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
          <p className="leave-meta">
            Fetched {records.length} record(s).
          </p>
        </div>
      )}
    </div>
  );
};

export default Leave;
