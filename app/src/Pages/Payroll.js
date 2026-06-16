import React, { useState, useEffect, useCallback, useRef } from 'react';
import { useLocation } from 'react-router-dom';
import './People.css';
import { getPayrollOrganizationId } from '../utils/payrollOrgId';
import {
  collectPayrollColumnKeys,
  flattenPayrollEarningColumns,
} from '../utils/payrollEarnings';

const API_BASE = '/server/payroll_function';
const EMPLOYEE_BATCH_SIZE = 200;
const DETAIL_BATCH_SIZE = 1;
const DETAIL_DELAY_MS = 800;

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

const getDefaultPayrollMonth = () => {
  const now = new Date();
  const d = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
};

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

async function savePayrollSnapshot({
  payrollMonth,
  organizationId,
  runMeta,
  records,
  hasBreakdown = false,
  totalExpected,
  breakdownComplete = false,
}) {
  const res = await fetch(`${API_BASE}/save`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      payrollMonth,
      organizationId,
      runMeta,
      records,
      hasBreakdown: hasBreakdown === true,
      totalExpected: totalExpected ?? records.length,
      breakdownComplete: breakdownComplete === true,
    }),
  });
  const text = await res.text();
  let json = {};
  try {
    json = text ? JSON.parse(text) : {};
  } catch {
    throw new Error(
      `Payroll save returned HTTP ${res.status} (not JSON). Redeploy payroll_function and check Catalyst logs.`
    );
  }
  if (!res.ok || !json.success) {
    throw new Error(json.error || json.message || `Save failed (HTTP ${res.status})`);
  }
  return json;
}

/**
 * Zoho Payroll pay-run data — month-wise employee summary from processed pay runs.
 */
const Payroll = ({ userRole, userEmail }) => {
  const location = useLocation();
  const [payrollMonth, setPayrollMonth] = useState(getDefaultPayrollMonth);
  const [data, setData] = useState(null);
  const [runMeta, setRunMeta] = useState(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const [loadMode, setLoadMode] = useState('list');
  const [progress, setProgress] = useState('');
  const [saveMessage, setSaveMessage] = useState('');
  const runMetaRef = useRef(null);
  const dataRef = useRef(null);

  useEffect(() => {
    runMetaRef.current = runMeta;
  }, [runMeta]);

  useEffect(() => {
    dataRef.current = data;
  }, [data]);

  const fetchData = useCallback(async (mode = 'list') => {
    const organizationId = getPayrollOrganizationId();
    if (!payrollMonth || !/^\d{4}-\d{2}$/.test(payrollMonth)) {
      setError('Select a valid month (YYYY-MM).');
      return;
    }

    setLoading(true);
    setError('');
    setProgress('');
    setSaveMessage('');
    if (mode !== 'full') {
      setData(null);
      setRunMeta(null);
    }
    setLoadMode(mode);

    try {
      if (mode === 'full') {
        const payrollRunId = runMetaRef.current?.payroll_run_id;
        const baseRows = Array.isArray(dataRef.current) ? dataRef.current : [];
        if (!payrollRunId || baseRows.length === 0) {
          throw new Error('Fetch monthly payroll first, then load the earnings breakdown.');
        }

        const merged = [];
        for (let i = 0; i < baseRows.length; i += DETAIL_BATCH_SIZE) {
          const slice = baseRows.slice(i, i + DETAIL_BATCH_SIZE);
          const details = await Promise.all(
            slice.map(async (row) => {
              const employeeId = row.employee_id || row.employeeId;
              if (!employeeId) return { ...row, fetch_error: true, error: 'missing_employee_id' };
              try {
                const qs = new URLSearchParams({
                  organization_id: organizationId,
                  payrun_employee_detail: '1',
                  payroll_run_id: String(payrollRunId),
                  employee_id: String(employeeId),
                });
                const json = await fetchPayrollJson(qs);
                return flattenPayrollEarningColumns({
                  ...row,
                  ...(json.data || {}),
                });
              } catch (detailErr) {
                return {
                  ...row,
                  fetch_error: true,
                  error: detailErr.message || 'detail_fetch_failed',
                };
              }
            })
          );
          merged.push(...details);
          if (i + DETAIL_BATCH_SIZE < baseRows.length) {
            await sleep(DETAIL_DELAY_MS);
          }
          setProgress(
            `Loading salary components… ${merged.length} / ${baseRows.length} · saving to Payroll table…`
          );
          setData([...merged]);
          const saveResult = await savePayrollSnapshot({
            payrollMonth,
            organizationId,
            runMeta: runMetaRef.current,
            records: merged,
            hasBreakdown: true,
            totalExpected: baseRows.length,
            breakdownComplete: merged.length >= baseRows.length,
          });
          const savedCount = saveResult?.data?.recordCount ?? merged.length;
          setSaveMessage(
            `Stored ${savedCount} / ${baseRows.length} employee record(s) in Payroll table.`
          );
        }
      } else {
        const merged = [];
        let offset = 0;
        let total = null;
        let meta = null;

        while (true) {
          const qs = new URLSearchParams({
            organization_id: organizationId,
            payroll_month_data: '1',
            payroll_month: payrollMonth,
            payroll_run_type: 'regular',
            employee_offset: String(offset),
            employee_limit: String(EMPLOYEE_BATCH_SIZE),
            include_earnings_detail: '1',
          });
          const json = await fetchPayrollJson(qs);
          const batch = (Array.isArray(json.data) ? json.data : []).map((row) =>
            flattenPayrollEarningColumns(row)
          );
          merged.push(...batch);
          meta = json.meta || meta;

          if (total == null && Number.isFinite(meta?.total)) {
            total = meta.total;
          }
          const loaded = merged.length;
          const displayTotal = total ?? loaded;
          const hasMore =
            json.meta?.has_more === true ||
            (json.meta?.has_more !== false && batch.length >= EMPLOYEE_BATCH_SIZE);
          setProgress(
            `Loading ${payrollMonth} payroll (Basic, HRA, allowances)… ${loaded} / ${displayTotal} · saving to Payroll table…`
          );
          setData([...merged]);
          setRunMeta(meta);

          const saveResult = await savePayrollSnapshot({
            payrollMonth,
            organizationId,
            runMeta: meta,
            records: merged,
            hasBreakdown: true,
            totalExpected: displayTotal,
            breakdownComplete: !hasMore || batch.length === 0,
          });
          const savedCount = saveResult?.data?.recordCount ?? loaded;
          setSaveMessage(
            `Stored ${savedCount} / ${displayTotal} employee record(s) in Payroll table.`
          );

          if (!hasMore || batch.length === 0) break;
          offset += batch.length;
        }
      }
    } catch (err) {
      setError(err.message || 'Failed to fetch payroll data');
      if (mode !== 'full') {
        setData(null);
        setRunMeta(null);
      }
    } finally {
      setLoading(false);
      setProgress('');
    }
  }, [payrollMonth]);

  useEffect(() => {
    fetchData('list');
  }, [fetchData, location.pathname]);

  const records = (() => {
    let raw = [];
    if (!data) return raw;
    if (Array.isArray(data)) raw = data;
    else {
      const res = data.response || data.result || data;
      if (Array.isArray(res)) raw = res;
      else {
        const rec = res?.record ?? res?.records ?? res?.data;
        if (Array.isArray(rec)) raw = rec;
        else if (res && typeof res === 'object') raw = [res];
      }
    }
    return raw.map((row) => flattenPayrollEarningColumns(row));
  })();

  const keys = records.length > 0 ? collectPayrollColumnKeys(records) : [];

  const formatMonthLabel = (value) => {
    if (!value || !/^\d{4}-\d{2}$/.test(value)) return value || '';
    const [y, m] = value.split('-');
    const date = new Date(Number(y), Number(m) - 1, 1);
    return date.toLocaleString('en-IN', { month: 'long', year: 'numeric' });
  };

  return (
    <div className="people-page">
      <header className="people-header">
        <h1 className="people-title">Salary details</h1>
        <p className="people-subtitle">
          Fetch month-wise payroll from Zoho Payroll pay runs. Basic, HRA, and allowances are loaded
          from each employee&apos;s pay-run earnings. Use Load salary breakdown to refresh rows
          already fetched.
        </p>
      </header>

      <div className="people-actions" style={{ display: 'flex', gap: 12, flexWrap: 'wrap' }}>
        <label className="people-subtitle" style={{ margin: 0, display: 'flex', alignItems: 'center' }}>
          Month
          <input
            type="month"
            value={payrollMonth}
            onChange={(e) => setPayrollMonth(e.target.value)}
            style={{
              marginLeft: 8,
              padding: '8px 10px',
              borderRadius: 6,
              border: '1px solid #d1d5db',
            }}
          />
        </label>
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
          disabled={loading || !records.length}
        >
          {loading && loadMode === 'full' ? 'Loading salaries...' : 'Load salary breakdown'}
        </button>
      </div>

      {runMeta && !loading && (
        <p className="people-subtitle" style={{ marginTop: 8 }}>
          Pay run: {runMeta.processing_period || formatMonthLabel(payrollMonth)}
          {runMeta.status ? ` · ${runMeta.status}` : ''}
          {runMeta.type ? ` · ${runMeta.type}` : ''}
          {runMeta.pay_date ? ` · Pay date ${runMeta.pay_date}` : ''}
        </p>
      )}

      {error && <div className="people-error">{error}</div>}

      {saveMessage && !loading && !error && (
        <p className="people-subtitle" style={{ marginTop: 8, color: '#047857' }}>
          {saveMessage}
        </p>
      )}

      {loading && (
        <div className="people-loading">
          {progress || `Loading payroll data for ${formatMonthLabel(payrollMonth)}...`}
        </div>
      )}

      {!loading && data !== null && (
        <div className="people-content">
          {records.length === 0 ? (
            <p className="people-empty">
              No pay run data for {formatMonthLabel(payrollMonth)}. If payroll is not processed yet
              in Zoho Payroll, run it there first or choose another month.
            </p>
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
          <p className="people-meta">
            Fetched {records.length} record(s) for {formatMonthLabel(payrollMonth)}.
          </p>
        </div>
      )}
    </div>
  );
};

export default Payroll;
