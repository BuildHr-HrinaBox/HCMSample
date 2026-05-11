import React, { useState } from 'react';
import './People.css';

const API_BASE = '/server/attendance_function';

const formatDate = (d) => d.toISOString().slice(0, 10);
const getMonthStart = () => {
  const now = new Date();
  return formatDate(new Date(now.getFullYear(), now.getMonth(), 1));
};

const Attendance = ({ userRole, userEmail }) => {
  const [data, setData] = useState(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const [sdate, setSdate] = useState(getMonthStart());
  const [edate, setEdate] = useState(formatDate(new Date()));

  const fetchData = async () => {
    if (sdate && edate && sdate > edate) {
      setError('From date cannot be after To date.');
      return;
    }

    setLoading(true);
    setError('');
    setData(null);
    try {
      const qs = new URLSearchParams({
        limit: '200',
        sdate,
        edate,
      });
      const res = await fetch(`${API_BASE}?${qs.toString()}`);
      const json = await res.json();
      if (!res.ok) {
        throw new Error(json.error || json.message || 'Request failed');
      }
      if (json.success && json.data !== undefined) {
        setData(json.data);
      } else {
        throw new Error(json.error || 'Invalid response');
      }
    } catch (err) {
      setError(err.message || 'Failed to fetch attendance data');
      setData(null);
    } finally {
      setLoading(false);
    }
  };

  // Same as People: Zoho API may return { response: { result: { record: [...] } } } or similar
  const records = (() => {
    if (!data) return [];
    if (Array.isArray(data)) return data;
    const res = data.response || data.result || data;
    if (Array.isArray(res)) return res;
    const rec = res?.record ?? res?.records ?? res?.data;
    if (Array.isArray(rec)) return rec;
    if (res && typeof res === 'object') return [res];
    return [];
  })();

  const firstRecord = records[0];
  const keys = firstRecord && typeof firstRecord === 'object'
    ? Object.keys(firstRecord).filter(k => !/^_|^\./.test(k))
    : [];

  return (
    <div className="people-page">
      <header className="people-header">
        <h1 className="people-title">Attendance</h1>
        <p className="people-subtitle">Fetch and view attendance data from Zoho People Attendance API</p>
      </header>

      <div className="people-actions">
        <div style={{ display: 'flex', gap: 12, alignItems: 'center', flexWrap: 'wrap', marginBottom: 12 }}>
          <label className="people-subtitle" style={{ margin: 0 }}>
            From
            <input
              type="date"
              value={sdate}
              onChange={(e) => setSdate(e.target.value)}
              style={{ marginLeft: 8, padding: '8px 10px', borderRadius: 6, border: '1px solid #d1d5db' }}
            />
          </label>
          <label className="people-subtitle" style={{ margin: 0 }}>
            To
            <input
              type="date"
              value={edate}
              onChange={(e) => setEdate(e.target.value)}
              style={{ marginLeft: 8, padding: '8px 10px', borderRadius: 6, border: '1px solid #d1d5db' }}
            />
          </label>
        </div>
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
          Loading attendance data...
        </div>
      )}

      {!loading && data !== null && (
        <div className="people-content">
          {records.length === 0 ? (
            <p className="people-empty">No records in response. Raw data structure may differ.</p>
          ) : (
            <div className="people-table-wrap">
              <table className="people-table">
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
          <p className="people-meta">
            Fetched {records.length} record(s).
          </p>
        </div>
      )}
    </div>
  );
};

export default Attendance;
