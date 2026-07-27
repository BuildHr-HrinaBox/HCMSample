import React, { useState, useCallback } from 'react';
import './People.css';
import { flattenPayrollEarningColumns } from '../utils/payrollEarnings';
import { fetchZohoPayrollRowsForMonth } from '../utils/payrollTable';
import { enrichPayrollRowsWithPeopleEmailFromApi } from '../utils/samplePayrollApi';

const API_BASE = '/server/samplepayroll_function';
const PAYROLL_API_BASE = '/server/payroll_function';
/** Same rate limit as Payroll "Load salary breakdown": 25 employees per call, one call per minute. */
const DETAIL_BATCH_SIZE = 25;
const DETAIL_BATCH_INTERVAL_MS = 60_000;

const TABLE_COLUMNS = [
  { key: 'employeeName', label: 'Employee Name' },
  { key: 'employeeId', label: 'Employee ID' },
  { key: 'gidNumber', label: 'GID Number' },
  { key: 'email', label: 'Email' },
  { key: 'dateofBirth', label: 'Date of Birth' },
  { key: 'paidDays', label: 'Paid Days' },
  { key: 'basic', label: 'Basic' },
  { key: 'hra', label: 'HRA' },
  { key: 'gross', label: 'Gross' },
  { key: 'netpay', label: 'Net Pay' },
  { key: 'totalDeduction', label: 'Total Deduction' },
  { key: 'incomeTax', label: 'Income Tax' },
  { key: 'pf', label: 'PF' },
  { key: 'voluntaryProvidentFund', label: 'Voluntary Provident Fund' },
  { key: 'professionalTax', label: 'Professional Tax' },
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

function amountPresent(value) {
  if (value == null || value === '') return false;
  const num = Number(String(value).replace(/,/g, '').trim());
  return Number.isFinite(num) && num !== 0;
}

function rowHasBasicHra(row) {
  const flat = flattenPayrollEarningColumns(row || {});
  const hasBasic = amountPresent(flat.basic) || amountPresent(flat.earned_basic);
  const hasHra = amountPresent(flat.hra) || amountPresent(flat.hra_fbp);
  return hasBasic || hasHra;
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
    employeeId: pickFirstValue(flat, ['employee_id', 'employeeId', 'EmployeeID']),
    gidNumber: pickFirstValue(flat, ['employee_number', 'gidNumber', 'GIDNumber', 'gid_number']),
    email: pickFirstValue(flat, [
      'email',
      'work_email',
      'work_mail',
      'personal_email',
      'mail_id',
      'EmailID',
      'Email',
      'Work_Email',
    ]),
    dateofBirth: pickFirstValue(flat, [
      'date_of_birth',
      'Date_of_birth',
      'DateofBirth',
      'dateofBirth',
      'dateOfBirth',
      'DOB',
      'dob',
    ]),
    paidDays: pickFirstValue(flat, ['paid_days', 'paidDays', 'Paid_days', 'paid_days_in_month']),
    basic: pickFirstValue(flat, ['basic', 'earned_basic', 'Basic', 'basic_pay']),
    hra: pickFirstValue(flat, ['hra', 'hra_fbp', 'HRA', 'house_rent_allowance']),
    gross,
    netpay,
    totalDeduction,
    incomeTax: pickFirstValue(flat, [
      'income_tax',
      'IncomeTax',
      'incomeTax',
      'Income Tax',
      'tds',
      'TDS',
      'tax_deducted_at_source',
    ]),
    pf: pickFirstValue(flat, [
      'epf_contribution',
      'EPF Contribution',
      'PF',
      'pf',
      'employer_pf',
      'employer_epf',
    ]),
    voluntaryProvidentFund: pickFirstValue(flat, [
      'voluntary_provident_fund',
      'VoluntaryProvidentFund',
      'Voluntary Provident Fund',
      'voluntaryProvidentFund',
      'vpf',
      'VPF',
    ]),
    professionalTax: pickFirstValue(flat, [
      'professional_tax',
      'ProfessionalTax',
      'professionalTax',
      'Professional Tax',
      'pt',
      'PT',
    ]),
  };
}

function mapSamplePayrollRecordToTable(record) {
  return {
    employeeName: record.employeeName || '',
    employeeId: record.employeeId || '',
    gidNumber: record.gidNumber || '',
    email: record.email || '',
    dateofBirth: record.dateofBirth || '',
    paidDays: record.paidDays || '',
    basic: record.basic || '',
    hra: record.hra || '',
    gross: record.gross || '',
    netpay: record.netpay || '',
    totalDeduction: record.totalDeduction || '',
    incomeTax: record.incomeTax || '',
    pf: record.pf || '',
    voluntaryProvidentFund: record.voluntaryProvidentFund || '',
    professionalTax: record.professionalTax || '',
  };
}

/** Prefer Payroll table when Load salary breakdown already filled Basic/HRA. */
async function loadPayrollTableWithBreakdown(payrollMonth) {
  const attempts = [
    `${PAYROLL_API_BASE}/payroll?${new URLSearchParams({ payroll_month: payrollMonth })}`,
    `${PAYROLL_API_BASE}?${new URLSearchParams({ payroll_table: '1', payroll_month: payrollMonth })}`,
  ];
  for (let i = 0; i < attempts.length; i += 1) {
    try {
      const res = await fetch(attempts[i], { cache: 'no-store' });
      const json = await res.json().catch(() => ({}));
      const records = Array.isArray(json?.data?.records) ? json.data.records : [];
      if (!res.ok || !json?.success || records.length === 0) continue;
      const withBreakdown = records.filter(rowHasBasicHra);
      // Use Payroll table only when most rows already have Basic/HRA.
      if (withBreakdown.length < Math.max(1, Math.floor(records.length * 0.5))) continue;
      return {
        records: records.map((row) => flattenPayrollEarningColumns(row)),
        meta: json.data.meta || null,
        payrollMonth: json.data.payrollMonth || payrollMonth,
        source: 'payroll_table',
      };
    } catch (_) {
      /* try next */
    }
  }
  return null;
}

async function fetchAllPayrollRecordsForMonth(payrollMonth, { onProgress, onRows } = {}) {
  const fromPayrollTable = await loadPayrollTableWithBreakdown(payrollMonth);
  if (fromPayrollTable) {
    if (typeof onProgress === 'function') {
      onProgress(
        `Using Payroll table (${fromPayrollTable.records.length} rows with Basic/HRA)… matching emails…`
      );
    }
    const withEmail = await enrichPayrollRowsWithPeopleEmailFromApi(fromPayrollTable.records);
    const records = withEmail.map((row) => flattenPayrollEarningColumns(row));
    if (typeof onRows === 'function') onRows(records);
    return {
      records,
      meta: fromPayrollTable.meta || null,
      payrollMonth: fromPayrollTable.payrollMonth || payrollMonth,
      source: fromPayrollTable.source,
    };
  }

  const zohoLoad = await fetchZohoPayrollRowsForMonth(payrollMonth, {
    timeoutMs: 300000,
    includeEarningsDetail: true,
    batchSize: DETAIL_BATCH_SIZE,
    detailBatchSize: DETAIL_BATCH_SIZE,
    detailBatchIntervalMs: DETAIL_BATCH_INTERVAL_MS,
    detailMode: 'payrun',
    onProgress,
    onDetailBatch: (rows) => {
      if (typeof onRows === 'function') {
        onRows(rows.map((row) => flattenPayrollEarningColumns(row)));
      }
    },
  });

  const records = Array.isArray(zohoLoad.rows) ? zohoLoad.rows : [];
  if (records.length === 0) {
    throw new Error(
      `No payroll data found for ${formatMonthLabel(payrollMonth)}. Fetch payroll on the Payroll page first.`
    );
  }

  if (typeof onProgress === 'function') {
    onProgress('matching emails from People…');
  }
  const withEmail = await enrichPayrollRowsWithPeopleEmailFromApi(records);
  const flattened = withEmail.map((row) => flattenPayrollEarningColumns(row));
  if (typeof onRows === 'function') onRows(flattened);

  return {
    records: flattened,
    meta: zohoLoad.meta || null,
    payrollMonth: zohoLoad.meta?.payrollMonth || payrollMonth,
    source: 'zoho_live',
  };
}

async function syncRecordsToSamplePayrollTable(payrollMonth, records, { onProgress } = {}) {
  const slimRecords = records.map((row) => mapPayrollRowToTable(row));
  const chunkSize = 40;
  let lastJson = null;

  for (let offset = 0; offset < slimRecords.length; offset += chunkSize) {
    const chunk = slimRecords.slice(offset, offset + chunkSize);
    const isLast = offset + chunkSize >= slimRecords.length;
    if (typeof onProgress === 'function') {
      onProgress(
        `Saving to SamplePayroll table… ${Math.min(offset + chunk.length, slimRecords.length)} / ${slimRecords.length}`
      );
    }
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
          setProgress(
            `Loading Basic, HRA… ${label} (25 per call, 1 call/min — keep this page open)`
          ),
        onRows: (rows) => {
          setData(rows.map(mapPayrollRowToTable));
        },
      });

      setProgress('Saving to SamplePayroll table…');
      const syncResult = await syncRecordsToSamplePayrollTable(
        result.payrollMonth || payrollMonth,
        result.records,
        { onProgress: setProgress }
      );

      const storedFromSync = Array.isArray(syncResult.data?.records)
        ? syncResult.data.records.map(mapSamplePayrollRecordToTable)
        : [];
      const storedFromFetch = result.records.map(mapPayrollRowToTable);
      const stored =
        storedFromSync.length >= storedFromFetch.length ? storedFromSync : storedFromFetch;

      const withBasic = stored.filter((row) => amountPresent(row.basic) || amountPresent(row.hra));

      setData(stored);
      setMeta(result.meta || null);
      setSource(result.source || '');
      setSaveMessage(
        (syncResult.message ||
          `Stored ${stored.length} employee record(s) in SamplePayroll table for ${
            result.payrollMonth || payrollMonth
          }.`) + (withBasic.length > 0 ? ` Basic/HRA on ${withBasic.length} row(s).` : '')
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
          Basic and HRA load in batches of 25 employees per call (one call per minute). If Payroll
          already has a breakdown for this month, those values are reused.
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
          {meta.payDate || meta.pay_date
            ? ` · Pay date ${meta.payDate || meta.pay_date}`
            : ''}
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
          {progress ||
            `Loading all employees for ${formatMonthLabel(payrollMonth)} (25 per call, 1 call/min)…`}
        </div>
      )}

      {data !== null && (
        <div className="people-content">
          {records.length === 0 && !error && !loading ? (
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
            Showing {records.length} record(s) for {formatMonthLabel(payrollMonth)}
            {loading ? ' (loading…)' : ' from SamplePayroll table'}.
          </p>
        </div>
      )}
    </div>
  );
};

export default SamplePayroll;
