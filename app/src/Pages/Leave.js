import React, { useState } from 'react';
import './Leave.css';
import {
  getDefaultLeaveReportRange,
  fetchLeaveReport,
} from '../utils/leaveApi';

/** Leave types used for earned / availed breakout columns */
const EARNED_LEAVE_NAMES = new Set([
  'Earned Leave (Test)',
  'Earned Leave(Test)',
  'Earned Leave (test)',
  'Earned Leave',
  'Earned leave',
]);

const HIDDEN_DETAIL_KEYS = new Set(['employeeId', 'totals']);

function parseJsonMaybe(val) {
  if (val == null) return null;
  if (typeof val === 'object') return val;
  if (typeof val === 'string') {
    const s = val.trim();
    if (!s.startsWith('{') && !s.startsWith('[')) return null;
    try {
      return JSON.parse(s);
    } catch {
      return null;
    }
  }
  return null;
}

/** Column key for earned-leave breakout after backend rename, or raw leave type id */
function findEarnedLeaveKey(row, leaveTypeLabels) {
  if (!row || typeof row !== 'object') return null;
  for (const key of Object.keys(row)) {
    if (EARNED_LEAVE_NAMES.has(key)) return key;
    const label = leaveTypeLabels && leaveTypeLabels[key];
    if (label && EARNED_LEAVE_NAMES.has(String(label).trim())) return key;
  }
  return null;
}

function getEarnedLeaveMetrics(row, earnedKey) {
  if (!earnedKey || !row || typeof row !== 'object') return { balance: null, booked: null };
  const raw = row[earnedKey];
  const obj = parseJsonMaybe(raw) ?? (raw && typeof raw === 'object' ? raw : null);
  if (!obj || typeof obj !== 'object') return { balance: null, booked: null };
  const balance = obj.paidBalance ?? obj.balance ?? obj.Balance;
  const booked = obj.paidBooked ?? obj.booked ?? obj.Booked;
  return {
    balance: balance != null && balance !== '' ? balance : null,
    booked: booked != null && booked !== '' ? booked : null,
  };
}

function formatEmployeeCell(val) {
  const o = parseJsonMaybe(val) ?? (val && typeof val === 'object' ? val : null);
  if (o && typeof o === 'object') {
    const name = o.name != null ? String(o.name) : '';
    const id = o.id != null ? String(o.id) : '';
    if (name && id) return `${name} (${id})`;
    return name || id || '';
  }
  return val == null ? '' : String(val);
}

/** Readable cell for other leave types (balance / booked) */
function formatLeaveTypeCell(val) {
  const obj = parseJsonMaybe(val) ?? (val && typeof val === 'object' ? val : null);
  if (!obj || typeof obj !== 'object') return val == null ? '' : String(val);
  if (Object.keys(obj).length === 0) return '—';
  if ('paidBalance' in obj || 'paidBooked' in obj || 'unpaidBalance' in obj || 'unpaidBooked' in obj) {
    const b = obj.paidBalance ?? obj.balance ?? obj.unpaidBalance;
    const book = obj.paidBooked ?? obj.booked ?? obj.unpaidBooked;
    const parts = [];
    if (b != null && b !== '') parts.push(`Balance: ${b}`);
    if (book != null && book !== '') parts.push(`Booked: ${book}`);
    return parts.length ? parts.join(', ') : '—';
  }
  if (Object.keys(obj).length === 1 && 'balance' in obj) return String(obj.balance);
  if ('balance' in obj || 'booked' in obj) {
    const parts = [];
    if (obj.balance != null && obj.balance !== '') parts.push(`Balance: ${obj.balance}`);
    if (obj.booked != null && obj.booked !== '') parts.push(`Booked: ${obj.booked}`);
    return parts.length ? parts.join(', ') : '—';
  }
  return JSON.stringify(obj);
}

const Leave = ({ userRole, userEmail }) => {
  const defaultRange = getDefaultLeaveReportRange();
  const [fromDate, setFromDate] = useState(defaultRange.from);
  const [toDate, setToDate] = useState(defaultRange.to);
  const [data, setData] = useState(null);
  const [meta, setMeta] = useState(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');

  const fetchData = async () => {
    setLoading(true);
    setError('');
    setData(null);
    setMeta(null);
    try {
      const result = await fetchLeaveReport({
        from: fromDate,
        to: toDate,
        unit: 'Day',
        fetchAll: true,
      });
      if (!result.success) {
        throw new Error('Invalid response');
      }
      setData({
        raw: result.raw,
        leaveTypeLabels: result.leaveTypeLabels,
        leaveRecords: result.leaveRecords,
      });
      setMeta(result.meta || null);
    } catch (err) {
      setError(err.message || 'Failed to fetch leave data');
      setData(null);
      setMeta(null);
    } finally {
      setLoading(false);
    }
  };

  const rawData = data && data.raw !== undefined ? data.raw : data;
  const backendRecords = data && Array.isArray(data.leaveRecords) ? data.leaveRecords : null;
  const leaveTypeLabels = data && data.leaveTypeLabels ? data.leaveTypeLabels : {};

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

  const earnedLeaveKey = records.reduce(
    (acc, r) => acc || findEarnedLeaveKey(r, leaveTypeLabels),
    null
  );
  const employeeKeyCandidates = ['employee', 'Employee', 'Employee.ID'];
  const employeeKey =
    employeeKeyCandidates.find((k) => keys.includes(k)) || (keys.includes('employee') ? 'employee' : null);

  const detailKeys = keys.filter(
    (k) => k !== earnedLeaveKey && k !== employeeKey && !HIDDEN_DETAIL_KEYS.has(k)
  );
  const showEarnedLeaveBreakout = earnedLeaveKey != null;

  const renderDetailCell = (row, k) => {
    const v = row[k];
    if (v != null && typeof v === 'object') return formatLeaveTypeCell(v);
    const parsed = parseJsonMaybe(v);
    if (parsed && typeof parsed === 'object') return formatLeaveTypeCell(parsed);
    return v == null ? '' : String(v);
  };

  // If no keys (e.g. empty objects or primitives), show raw data in one column
  const showRaw = records.length > 0 && keys.length === 0;

  return (
    <div className="leave-page">
      <header className="leave-header">
        <h1 className="leave-title">Leave</h1>
        <p className="leave-subtitle">
          Fetch leave booked/balance from Zoho People Leave API. Use the leave year window (typically Apr–Mar);
          calendar-year ranges often return only empty Absent rows. Earned Leave shows balance for the period
          and booked days for the same From–To window.
        </p>
      </header>

      <div className="leave-actions">
        <div style={{ display: 'flex', gap: 12, alignItems: 'center', flexWrap: 'wrap', marginBottom: 12 }}>
          <label className="leave-subtitle" style={{ margin: 0 }}>
            From
            <input
              type="text"
              value={fromDate}
              onChange={(e) => setFromDate(e.target.value)}
              placeholder="01-Apr-2026"
              style={{ marginLeft: 8, padding: '8px 10px', borderRadius: 6, border: '1px solid #d1d5db', width: 130 }}
            />
          </label>
          <label className="leave-subtitle" style={{ margin: 0 }}>
            To
            <input
              type="text"
              value={toDate}
              onChange={(e) => setToDate(e.target.value)}
              placeholder="31-Mar-2027"
              style={{ marginLeft: 8, padding: '8px 10px', borderRadius: 6, border: '1px solid #d1d5db', width: 130 }}
            />
          </label>
        </div>
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
                    <th>Employee</th>
                    {showEarnedLeaveBreakout && (
                      <>
                        <th title="Earned Leave paidBalance from Zoho booked/balance report">
                          Leave earned during the Period
                        </th>
                        <th title="Earned Leave paidBooked from Zoho booked/balance report">
                          Leave availed during the Period
                        </th>
                      </>
                    )}
                    <th title="Row totals from Zoho (paid + unpaid balance/booked)">Totals</th>
                    {detailKeys.map((k) => (
                      <th key={k}>{leaveTypeLabels[k] || k}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {records.map((row, i) => {
                    const earned = showEarnedLeaveBreakout
                      ? getEarnedLeaveMetrics(row, earnedLeaveKey)
                      : { balance: null, booked: null };
                    const employeeDisplay = employeeKey
                      ? formatEmployeeCell(row[employeeKey])
                      : formatEmployeeCell(row.employee) ||
                        (row.employeeId != null ? String(row.employeeId) : '');
                    return (
                      <tr key={i}>
                        <td>{i + 1}</td>
                        <td>{employeeDisplay}</td>
                        {showEarnedLeaveBreakout && (
                          <>
                            <td className="leave-num">{earned.balance != null ? earned.balance : ''}</td>
                            <td className="leave-num">{earned.booked != null ? earned.booked : ''}</td>
                          </>
                        )}
                        <td>{formatLeaveTypeCell(row.totals)}</td>
                        {detailKeys.map((k) => (
                          <td key={k}>{renderDetailCell(row, k)}</td>
                        ))}
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          )}
          <p className="leave-meta">
            Fetched {records.length} record(s)
            {meta?.pages ? ` across ${meta.pages} Zoho page(s)` : ''}.
          </p>
        </div>
      )}
    </div>
  );
};

export default Leave;
