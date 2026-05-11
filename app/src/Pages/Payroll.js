import React, { useState, useMemo } from 'react';
import './People.css';

const API_BASE = '/server/payroll_function';

const PREFERRED_KEYS = ['employee_id', 'first_name', 'last_name', 'work_mail', 'fetch_error'];

function computeColumns(rows) {
  const keySet = new Set();
  for (const row of rows) {
    if (row && typeof row === 'object') {
      Object.keys(row).forEach((k) => {
        if (!/^_|^\./.test(k)) keySet.add(k);
      });
    }
  }
  const front = PREFERRED_KEYS.filter((k) => keySet.has(k));
  const rest = [...keySet].filter((k) => !PREFERRED_KEYS.includes(k)).sort();
  return [...front, ...rest];
}

const Payroll = () => {
  const [data, setData] = useState(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');

  const orgId = process.env.REACT_APP_ZOHO_PAYROLL_ORGANIZATION_ID || '60006183023';

  const fetchAllSalaries = async () => {
    setLoading(true);
    setError('');
    setData(null);
    try {
      const qs = new URLSearchParams({
        organization_id: orgId,
        all_salaries: '1',
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
      setError(err.message || 'Failed to fetch payroll data');
      setData(null);
    } finally {
      setLoading(false);
    }
  };

  const records = useMemo(() => {
    if (!data) return [];
    if (Array.isArray(data)) return data;
    if (data.employee && typeof data.employee === 'object') return [data.employee];
    if (Array.isArray(data.employees)) return data.employees;
    if (Array.isArray(data.salaries)) return data.salaries;
    if (Array.isArray(data.payrolls)) return data.payrolls;
    if (Array.isArray(data.data)) return data.data;
    const res = data.response || data.result || data;
    if (Array.isArray(res)) return res;
    if (res && typeof res === 'object') return [res];
    return [];
  }, [data]);

  const keys = useMemo(() => {
    if (records.length === 0) return [];
    return computeColumns(records);
  }, [records]);

  return (
    <div className="people-page">
      <header className="people-header">
        <h1 className="people-title">Payroll</h1>
        <p className="people-subtitle">
          Loads every Payroll employee in your organisation (all pages), then fetches salary for each. One click —
          results appear in the table below.
        </p>
      </header>

      <div className="people-actions">
        <button
          type="button"
          className="people-fetch-btn"
          onClick={fetchAllSalaries}
          disabled={loading}
        >
          {loading ? 'Fetching…' : 'Fetch all salaries'}
        </button>
      </div>

      {error && <div className="people-error">{error}</div>}

      {loading && <div className="people-loading">Loading salaries for all employees…</div>}

      {!loading && data !== null && (
        <div className="people-content">
          {records.length === 0 ? (
            <p className="people-empty">No employees or salary rows returned.</p>
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
                    <tr key={row.employee_id != null ? String(row.employee_id) : i}>
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
          <p className="people-meta">Fetched {records.length} employee salary row(s).</p>
        </div>
      )}
    </div>
  );
};

export default Payroll;
