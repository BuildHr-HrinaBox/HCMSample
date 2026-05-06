import React, { useState } from 'react';
import './People.css';

const API_BASE = '/server/peopledata_function';

const People = ({ userRole, userEmail }) => {
  const [data, setData] = useState(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');

  const fetchData = async () => {
    setLoading(true);
    setError('');
    setData(null);
    try {
      const res = await fetch(`${API_BASE}?form=employee&limit=50`);
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
      setError(err.message || 'Failed to fetch people data');
      setData(null);
    } finally {
      setLoading(false);
    }
  };

  // Zoho People API may return { response: { result: { record: [...] } } } or similar
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
          Loading people data...
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

export default People;
