import React, { useState, useEffect, useCallback } from 'react';
import { useLocation } from 'react-router-dom';
import './People.css';
import { getPayrollOrganizationId } from '../utils/payrollOrgId';

const API_BASE = '/server/payroll_function';
const SALARY_BATCH_SIZE = 6;

async function fetchPayrollJson(qs) {
  const res = await fetch(`${API_BASE}?${qs.toString()}`);
  const text = await res.text();
  let json = {};
  try {
    json = text ? JSON.parse(text) : {};
  } catch {
    throw new Error(
      `Payroll server returned HTTP ${res.status} (not JSON). Redeploy payroll_function and check Catalyst logs.`
    );
  }
  if (!res.ok) {
    throw new Error(json.error || json.message || `Request failed (HTTP ${res.status})`);
  }
  if (!json.success || json.data === undefined) {
    throw new Error(json.error || 'Invalid response');
  }
  return json;
}

/**
 * Zoho Payroll salary list — same UX as People / Leave (fetch on open + Refresh button).
 */
const Payroll = ({ userRole, userEmail }) => {
  const location = useLocation();
  const [data, setData] = useState(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const [loadMode, setLoadMode] = useState('list');
  const [progress, setProgress] = useState('');

  const fetchData = useCallback(async (mode = 'list') => {
    const organizationId = getPayrollOrganizationId();
    setLoading(true);
    setError('');
    setProgress('');
    if (mode !== 'full') {
      setData(null);
    }
    setLoadMode(mode);
    try {
      if (mode === 'full') {
        const merged = [];
        let offset = 0;
        let total = null;

        while (true) {
          const qs = new URLSearchParams({
            organization_id: organizationId,
            all_salaries: '1',
            salary_offset: String(offset),
            salary_limit: String(SALARY_BATCH_SIZE),
          });
          const json = await fetchPayrollJson(qs);
          const batch = Array.isArray(json.data) ? json.data : [];
          merged.push(...batch);

          const meta = json.meta || {};
          if (total == null && Number.isFinite(meta.total)) {
            total = meta.total;
          }
          const loaded = merged.length;
          const displayTotal = total ?? loaded;
          setProgress(`Loading salary breakdown… ${loaded} / ${displayTotal}`);
          setData([...merged]);

          const hasMore =
            meta.has_more === true ||
            (meta.has_more !== false && batch.length >= SALARY_BATCH_SIZE);
          if (!hasMore || batch.length === 0) break;
          offset += batch.length;
        }
      } else {
        const qs = new URLSearchParams({
          organization_id: organizationId,
          list_employees: '1',
        });
        const json = await fetchPayrollJson(qs);
        setData(json.data);
      }
    } catch (err) {
      setError(err.message || 'Failed to fetch payroll data');
      if (mode !== 'full') {
        setData(null);
      }
    } finally {
      setLoading(false);
      setProgress('');
    }
  }, []);

  useEffect(() => {
    fetchData('list');
  }, [fetchData, location.pathname]);

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
          Fetch employee payroll roster from Zoho Payroll (same flow as People). Use Load salary breakdown for full earnings per employee (slower).
        </p>
      </header>

      <div className="people-actions" style={{ display: 'flex', gap: 12, flexWrap: 'wrap' }}>
        <button
          type="button"
          className="people-fetch-btn"
          onClick={() => fetchData('list')}
          disabled={loading}
        >
          {loading && loadMode === 'list' ? 'Fetching...' : 'Fetch Data'}
        </button>
        <button
          type="button"
          className="people-fetch-btn"
          style={{ background: '#2563eb' }}
          onClick={() => fetchData('full')}
          disabled={loading}
        >
          {loading && loadMode === 'full' ? 'Loading salaries...' : 'Load salary breakdown'}
        </button>
      </div>

      {error && <div className="people-error">{error}</div>}

      {loading && (
        <div className="people-loading">
          {progress || 'Loading payroll data from Zoho Payroll...'}
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
          <p className="people-meta">Fetched {records.length} record(s).</p>
        </div>
      )}
    </div>
  );
};

export default Payroll;
