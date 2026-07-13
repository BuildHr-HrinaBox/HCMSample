import React, { useState, useEffect, useCallback, useRef } from 'react';
import './People.css';
import { getPayrollOrganizationId } from '../utils/payrollOrgId';
import {
  collectPayrollColumnKeys,
  flattenPayrollEarningColumns,
} from '../utils/payrollEarnings';

const API_BASE = '/server/payroll_function';
const EMPLOYEE_BATCH_SIZE = 25;
/** Load salary breakdown: 25 employees per API call, one call per minute. */
const DETAIL_BATCH_SIZE = 25;
const DETAIL_BATCH_INTERVAL_MS = 60_000;
const PAYROLL_FETCH_BATCH_DELAY_MS = 300;

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

function pickEmployeeId(row) {
  if (!row || typeof row !== 'object') return '';
  return String(
    row.employee_id ||
      row.employeeId ||
      row.Employee_ID ||
      row.id ||
      ''
  ).trim();
}

function mergeListAndDetailRow(listRow, detailRow) {
  const listFlat = flattenPayrollEarningColumns(listRow || {});
  const detailFlat = flattenPayrollEarningColumns(detailRow || {});
  const merged = { ...listFlat, ...detailFlat };
  const preserveKeys = ['net_pay', 'gross_pay', 'total_earnings', 'paid_days', 'total_deductions'];
  for (let i = 0; i < preserveKeys.length; i += 1) {
    const key = preserveKeys[i];
    const listVal = listFlat[key];
    const detailVal = detailFlat[key];
    const listNum = Number(listVal);
    const detailNum = Number(detailVal);
    if ((!Number.isFinite(detailNum) || detailNum === 0) && Number.isFinite(listNum) && listNum !== 0) {
      merged[key] = listVal;
    }
  }
  return flattenPayrollEarningColumns(merged);
}

const getDefaultPayrollMonth = () => {
  const now = new Date();
  const d = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
};

async function loadPayrollTableSnapshot(payrollMonth) {
  const attempts = [
    `${API_BASE}/payroll?${new URLSearchParams({ payroll_month: payrollMonth })}`,
    `${API_BASE}?${new URLSearchParams({ payroll_table: '1', payroll_month: payrollMonth })}`,
  ];
  for (let i = 0; i < attempts.length; i += 1) {
    try {
      const res = await fetch(attempts[i], { cache: 'no-store' });
      const json = await res.json().catch(() => ({}));
      if (res.ok && json?.success && Array.isArray(json.data?.records) && json.data.records.length > 0) {
        return {
          records: json.data.records.map((row) => flattenPayrollEarningColumns(row)),
          meta: json.data.meta || null,
          payrollMonth: json.data.payrollMonth || payrollMonth,
        };
      }
    } catch (_) {
      /* try next URL */
    }
  }
  return null;
}

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
    if (res.status === 408) {
      throw new Error(
        'Request timed out (HTTP 408). Payroll is fetched in small batches — wait for progress to finish, or try again. Use Load salary breakdown separately for Basic/HRA detail.'
      );
    }
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

        const updated = baseRows.map((row) => flattenPayrollEarningColumns(row));
        const total = updated.length;
        let enrichedCount = 0;

        for (let i = 0; i < total; i += DETAIL_BATCH_SIZE) {
          const batchStart = Date.now();
          const slice = updated.slice(i, i + DETAIL_BATCH_SIZE);
          const employeeIds = slice.map(pickEmployeeId).filter(Boolean);

          if (employeeIds.length === 0) {
            slice.forEach((row, idx) => {
              updated[i + idx] = {
                ...row,
                fetch_error: true,
                error: 'missing_employee_id',
              };
            });
          } else {
            try {
              const qs = new URLSearchParams({
                organization_id: organizationId,
                payrun_employee_detail: '1',
                payroll_run_id: String(payrollRunId),
                employee_ids: employeeIds.join(','),
              });
              const json = await fetchPayrollJson(qs);
              const details = Array.isArray(json.data)
                ? json.data
                : json.data && typeof json.data === 'object'
                  ? [json.data]
                  : [];
              const byId = new Map();
              details.forEach((detail) => {
                const id = pickEmployeeId(detail);
                if (id) byId.set(id, detail);
              });

              slice.forEach((row, idx) => {
                const id = pickEmployeeId(row);
                if (!id) {
                  updated[i + idx] = { ...row, fetch_error: true, error: 'missing_employee_id' };
                  return;
                }
                const detail = byId.get(id);
                if (!detail) {
                  updated[i + idx] = {
                    ...row,
                    fetch_error: true,
                    error: 'detail_not_returned',
                  };
                  return;
                }
                updated[i + idx] = mergeListAndDetailRow(row, detail);
                enrichedCount += 1;
              });
            } catch (detailErr) {
              slice.forEach((row, idx) => {
                updated[i + idx] = {
                  ...row,
                  fetch_error: true,
                  error: detailErr.message || 'detail_fetch_failed',
                };
              });
            }
          }

          const done = Math.min(i + DETAIL_BATCH_SIZE, total);
          setProgress(
            `Loading salary components… ${done} / ${total} (25 per call, 1 call/min) · saving…`
          );
          setData([...updated]);

          const saveResult = await savePayrollSnapshot({
            payrollMonth,
            organizationId,
            runMeta: runMetaRef.current,
            records: updated,
            hasBreakdown: true,
            totalExpected: total,
            breakdownComplete: done >= total,
          });
          const savedCount = saveResult?.data?.recordCount ?? updated.length;
          setSaveMessage(
            `Stored ${savedCount} / ${total} employee record(s) in Payroll table` +
              (enrichedCount > 0 ? ` · breakdown loaded for ${enrichedCount}.` : '.')
          );

          if (done < total) {
            const elapsed = Date.now() - batchStart;
            const waitMs = Math.max(0, DETAIL_BATCH_INTERVAL_MS - elapsed);
            if (waitMs > 0) {
              setProgress(
                `Loaded ${done} / ${total}. Waiting ${Math.ceil(waitMs / 1000)}s before next 25…`
              );
              await sleep(waitMs);
            }
          }
        }
      } else {
        const tableSnapshot = await loadPayrollTableSnapshot(payrollMonth);
        if (tableSnapshot?.records?.length > 0) {
          setData(tableSnapshot.records);
          setRunMeta(tableSnapshot.meta);
          setSaveMessage(
            `Loaded ${tableSnapshot.records.length} record(s) from Payroll table for ${formatMonthLabel(
              tableSnapshot.payrollMonth || payrollMonth
            )}. Click Fetch Data again to refresh from Zoho.`
          );
          return;
        }

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
            include_earnings_detail: '0',
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
            `Loading ${payrollMonth} payroll (gross pay, net pay)… ${loaded} / ${displayTotal} · saving to Payroll table…`
          );
          setData([...merged]);
          setRunMeta(meta);

          const saveResult = await savePayrollSnapshot({
            payrollMonth,
            organizationId,
            runMeta: meta,
            records: merged,
            hasBreakdown: false,
            totalExpected: displayTotal,
            breakdownComplete: !hasMore || batch.length === 0,
          });
          const savedCount = saveResult?.data?.recordCount ?? loaded;
          setSaveMessage(
            `Stored ${savedCount} / ${displayTotal} employee record(s) in Payroll table.`
          );

          if (!hasMore || batch.length === 0) break;
          offset += batch.length;
          await sleep(PAYROLL_FETCH_BATCH_DELAY_MS);
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
          already fetched (25 employees per call, one call per minute).
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