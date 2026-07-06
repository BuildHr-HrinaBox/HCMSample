import React, { useState, useCallback } from 'react';
import './People.css';
import { flattenPayrollEarningColumns } from '../utils/payrollEarnings';
import { fetchZohoPayrollRowsForMonth } from '../utils/payrollTable';

const API_BASE = '/server/samplepayroll_function';

const TABLE_COLUMNS = [
  { key: 'employeeName', label: 'Employee Name' },
  { key: 'employeeId', label: 'Employee ID' },
  { key: 'paidDays', label: 'Paid Days' },
  { key: 'basic', label: 'Basic' },
  { key: 'hra', label: 'HRA' },
  { key: 'gross', label: 'Gross' },
  { key: 'netpay', label: 'Net Pay' },
  { key: 'totalDeduction', label: 'Total Deduction' },
];

const getDefaultPayrollMonth = () => {
  const now = new Date();
  const d = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
};

const formatMonthLabel = (value) => {
  if (!value || !/^\d{4}-\d{2}$/.test(value)) return value || '';
  const [y, m] = value.split('-');
  const date = new Date(Number(y), Number(m) - 1, 1);
  return date.toLocaleString('en-IN', { month: 'long', year: 'numeric' });
};

function pickFirstValue(row, keys) {
  for (let i = 0; i < keys.length; i += 1) {
    const value = row[keys[i]];
    if (value != null && String(value).trim() !== '') return value;
  }
  return '';
}

function mapPayrollRowToTable(row) {
  const flat = flattenPayrollEarningColumns(row);
  const gross = pickFirstValue(flat, ['gross_pay', 'total_earnings', 'gross', 'Gross']);
  const totalDeduction = pickFirstValue(flat, [
    'total_deductions',
    'totalDeduction',
    'TotalDeduction',
    'total_employee_deductions',
  ]);
  let netpay = pickFirstValue(flat, [
    'net_pay',
    'netPay',
    'Netpay',
    'monthly_salary',
    'net_salary',
    'net_wages',
  ]);
  if (!netpay && gross && totalDeduction) {
    const netNum = Number(gross) - Number(totalDeduction);
    if (Number.isFinite(netNum) && netNum >= 0) netpay = String(netNum);
  }
  return {
    employeeName: pickFirstValue(flat, ['employee_name', 'full_name', 'employeeName', 'name']),
    employeeId: pickFirstValue(flat, ['employee_id', 'employee_number', 'employeeId', 'EmployeeID']),
    paidDays: pickFirstValue(flat, ['paid_days', 'paidDays', 'Paid_days', 'paid_days_in_month']),
    basic: pickFirstValue(flat, ['basic', 'earned_basic', 'Basic', 'basic_pay']),
    hra: pickFirstValue(flat, ['hra', 'hra_fbp', 'HRA', 'house_rent_allowance']),
    gross,
    netpay,
    totalDeduction,
  };
}

function mapSamplePayrollRecordToTable(record) {
  return {
    employeeName: record.employeeName || '',
    employeeId: record.employeeId || '',
    paidDays: record.paidDays || '',
    basic: record.basic || '',
    hra: record.hra || '',
    gross: record.gross || '',
    netpay: record.netpay || '',
    totalDeduction: record.totalDeduction || '',
  };
}

async function fetchAllPayrollRecordsForMonth(payrollMonth, { onProgress } = {}) {
  const zohoLoad = await fetchZohoPayrollRowsForMonth(payrollMonth, {
    timeoutMs: 120000,
    includeEarningsDetail: true,
    batchSize: 50,
    detailConcurrency: 5,
    onProgress,
  });

  const records = Array.isArray(zohoLoad.rows) ? zohoLoad.rows : [];
  if (records.length === 0) {
    throw new Error(
      `No payroll data found for ${formatMonthLabel(payrollMonth)}. Fetch payroll on the Payroll page first.`
    );
  }

  return {
    records: records.map((row) => flattenPayrollEarningColumns(row)),
    meta: zohoLoad.meta || null,
    payrollMonth: zohoLoad.meta?.payrollMonth || payrollMonth,
    source: 'zoho_live',
  };
}

async function syncRecordsToSamplePayrollTable(payrollMonth, records) {
  const slimRecords = records.map((row) => mapPayrollRowToTable(row));
  const chunkSize = 40;
  let lastJson = null;

  for (let offset = 0; offset < slimRecords.length; offset += chunkSize) {
    const chunk = slimRecords.slice(offset, offset + chunkSize);
    const isLast = offset + chunkSize >= slimRecords.length;
    const res = await fetch(`${API_BASE}/samplepayroll/sync-month`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        payrollMonth,
        records: chunk,
        replaceExisting: offset === 0,
        finalize: isLast,
      }),
    });
    const json = await res.json().catch(() => ({}));
    if (!res.ok || json.status !== 'success') {
      throw new Error(json.message || 'Failed to store payroll data in SamplePayroll table.');
    }
    lastJson = json;
  }

  return lastJson;
}

const SamplePayroll = () => {
  const [payrollMonth, setPayrollMonth] = useState(getDefaultPayrollMonth);
  const [data, setData] = useState(null);
  const [meta, setMeta] = useState(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const [source, setSource] = useState('');
  const [saveMessage, setSaveMessage] = useState('');
  const [progress, setProgress] = useState('');

  const fetchData = useCallback(async () => {
    if (!payrollMonth || !/^\d{4}-\d{2}$/.test(payrollMonth)) {
      setError('Select a valid month (YYYY-MM).');
      return;
    }

    setLoading(true);
    setError('');
    setSaveMessage('');
    setProgress('');
    setData(null);
    setMeta(null);
    setSource('');

    try {
      const result = await fetchAllPayrollRecordsForMonth(payrollMonth, {
        onProgress: (label) =>
          setProgress(`Loading payroll (Basic, HRA, Net Pay)… ${label} employee(s)`),
      });
      setProgress('Saving to SamplePayroll table…');
      const syncResult = await syncRecordsToSamplePayrollTable(
        result.payrollMonth || payrollMonth,
        result.records
      );

      const stored = Array.isArray(syncResult.data?.records)
        ? syncResult.data.records.map(mapSamplePayrollRecordToTable)
        : result.records.map(mapPayrollRowToTable);

      setData(stored);
      setMeta(result.meta || null);
      setSource(result.source || '');
      setSaveMessage(
        syncResult.message ||
          `Stored ${stored.length} employee record(s) in SamplePayroll table.`
      );

      if (stored.length === 0) {
        setError(
          `No payroll employee rows found for ${formatMonthLabel(payrollMonth)}. Re-fetch payroll on the Payroll page.`
        );
      }
    } catch (err) {
      setData([]);
      setError(
        err.message ||
          `No payroll data found for ${formatMonthLabel(payrollMonth)}. Fetch and save payroll on the Payroll page first.`
      );
    } finally {
      setLoading(false);
      setProgress('');
    }
  }, [payrollMonth]);

  const records = Array.isArray(data) ? data : [];

  return (
    <div className="people-page">
      <header className="people-header">
        <h1 className="people-title">Sample Payroll</h1>
        <p className="people-subtitle">
          Fetch month-wise payroll, store it in the SamplePayroll table, and view all employees below.
        </p>
      </header>

      <div className="people-actions" style={{ display: 'flex', gap: 12, flexWrap: 'wrap', alignItems: 'center' }}>
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
          onClick={fetchData}
          disabled={loading}
        >
          {loading ? 'Loading...' : 'Fetch Data'}
        </button>
      </div>

      {meta && !loading && records.length > 0 && (
        <p className="people-subtitle" style={{ marginTop: 8 }}>
          Pay run: {formatMonthLabel(payrollMonth)}
          {meta.payroll_run_id ? ` · Run ID ${meta.payroll_run_id}` : ''}
          {meta.payDate ? ` · Pay date ${meta.payDate}` : ''}
          {source ? ` · Source: ${source}` : ''}
        </p>
      )}

      {saveMessage && !loading && !error && (
        <p className="people-subtitle" style={{ marginTop: 8, color: '#047857' }}>
          {saveMessage}
        </p>
      )}

      {error && <div className="people-error">{error}</div>}

      {loading && (
        <div className="people-loading">
          {progress || `Loading all employees for ${formatMonthLabel(payrollMonth)}…`}
        </div>
      )}

      {!loading && data !== null && (
        <div className="people-content">
          {records.length === 0 && !error ? (
            <p className="people-empty">
              No payroll records for {formatMonthLabel(payrollMonth)} in the SamplePayroll table.
            </p>
          ) : records.length > 0 ? (
            <div className="people-table-wrap">
              <table className="people-table">
                <thead>
                  <tr>
                    <th>S.No</th>
                    {TABLE_COLUMNS.map((col) => (
                      <th key={col.key}>{col.label}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {records.map((row, i) => (
                    <tr key={`${row.employeeId || row.employeeName || 'row'}-${i}`}>
                      <td>{i + 1}</td>
                      {TABLE_COLUMNS.map((col) => (
                        <td key={col.key}>{row[col.key] || '—'}</td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          ) : null}
          <p className="people-meta">
            Showing {records.length} record(s) for {formatMonthLabel(payrollMonth)} from SamplePayroll table.
          </p>
        </div>
      )}
    </div>
  );
};

export default SamplePayroll;
