import React, { useState, useEffect, useCallback } from 'react';
import './People.css';

const API_BASE = '/server/payroll_function';

const getOrgId = () =>
  process.env.REACT_APP_ZOHO_PAYROLL_ORGANIZATION_ID || '60006183023';

/**
 * Salary details (Zoho Payroll) — same UX pattern as People / Attendance.
 * Loads all employees with salary when the page opens.
 */
const Payroll = ({ userRole, userEmail }) => {
  const [data, setData] = useState(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');

  const fetchSalaryDetails = useCallback(async () => {
    const organizationId = getOrgId();
    setLoading(true);
    setError('');
    setData(null);
    try {
      const qs = new URLSearchParams({
        all_salaries: '1',
        organization_id: organizationId,
      });
      const res = await fetch(`${API_BASE}?${qs.toString()}`);
      const text = await res.text();
      let json = {};
      try {
        json = text ? JSON.parse(text) : {};
      } catch {
        throw new Error(
          `Payroll server returned HTTP ${res.status} (response was not JSON). Check function logs.`
        );
      }
      if (!res.ok) {
        throw new Error(
          json.error || json.message || `Request failed with HTTP ${res.status}`
        );
      }
      if (json.success && json.data !== undefined) {
        setData(json.data);
      } else {
        throw new Error(json.error || 'Invalid response');
      }
    } catch (err) {
      setError(err.message || 'Failed to fetch salary details');
      setData(null);
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    fetchSalaryDetails();
  }, [fetchSalaryDetails]);

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
  const keys =
    firstRecord && typeof firstRecord === 'object'
      ? Object.keys(firstRecord).filter((k) => !/^_|^\./.test(k))
      : [];

  return (
    <div className="people-page">
      <header className="people-header">
        <h1 className="people-title">Salary details</h1>
        <p className="people-subtitle">
          Zoho Payroll salary data (organisation {getOrgId()})
        </p>
      </header>

      <div className="people-actions">
        <button
          type="button"
          className="people-fetch-btn"
          onClick={fetchSalaryDetails}
          disabled={loading}
        >
          {loading ? 'Loading...' : 'Refresh'}
        </button>
      </div>

      {error && <div className="people-error">{error}</div>}

      {loading && (
        <div className="people-loading">Loading salary details from Zoho Payroll...</div>
      )}

      {!loading && data !== null && (
        <div className="people-content">
          {records.length === 0 ? (
            <p className="people-empty">No salary rows returned.</p>
          ) : (
            <div className="people-table-wrap">
              <table className="people-table">
                <thead>
                  <tr>
                    <th>S.No</th>
                    {keys.map((k) => (
                      <th key={k}>{k}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {records.map((row, i) => (
                    <tr key={i}>
                      <td>{i + 1}</td>
                      {keys.map((k) => (
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
          <p className="people-meta">Showing {records.length} row(s).</p>
        </div>
      )}
    </div>
  );
};

export default Payroll;
