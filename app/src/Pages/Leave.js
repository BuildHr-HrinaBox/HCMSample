import React, { useCallback, useRef, useState } from 'react';
import './Leave.css';
import {
  getCurrentMonthLeaveRange,
  getMonthEndIsoFromIso,
  isoToZohoLeaveDate,
  isoToDisplayDate,
  fetchLeaveReport,
  mapLeaveRecordToLeaveDataRow,
  saveLeaveDataToBackend,
  getLeaveMonthWise,
} from '../utils/leaveApi';
import { extractLeaveRecordsFromApiResult } from '../utils/leaveMetrics';

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

const Leave = ({ userRole, userEmail, apiBase, pageTitle = 'Leave' }) => {
  const defaultMonth = getCurrentMonthLeaveRange();
  const [fromDate, setFromDate] = useState(defaultMonth.from);
  const [toDate, setToDate] = useState(defaultMonth.to);
  const [data, setData] = useState(null);
  const [meta, setMeta] = useState(null);
  const [loading, setLoading] = useState(false);
  const [loadStatus, setLoadStatus] = useState('');
  const [error, setError] = useState('');
  const [saveMessage, setSaveMessage] = useState('');
  const fetchSeqRef = useRef(0);

  const fetchData = useCallback(async (overrideFrom, overrideTo) => {
    const fromIso = overrideFrom != null ? overrideFrom : fromDate;
    const toIso = overrideTo != null ? overrideTo : toDate;
    const fromZoho = isoToZohoLeaveDate(fromIso);
    const toZoho = isoToZohoLeaveDate(toIso);
    if (!fromZoho || !toZoho) {
      setError('Please pick a valid From and To date.');
      return;
    }
    if (fromIso > toIso) {
      setError('From date cannot be after To date.');
      return;
    }

    const seq = ++fetchSeqRef.current;
    setLoading(true);
    setError('');
    setLoadStatus(`Fetching leave for ${isoToDisplayDate(fromIso)} to ${isoToDisplayDate(toIso)}…`);
    setSaveMessage('');
    setData(null);
    setMeta(null);
    try {
      const result = await fetchLeaveReport({
        from: fromZoho,
        to: toZoho,
        unit: 'Day',
        fetchAll: true,
        apiBase,
        useCache: true,
        onProgress: ({ pages, total, status }) => {
          if (seq !== fetchSeqRef.current) return;
          if (status === 'cached') {
            setLoadStatus(`Loaded ${total} employees from cache`);
            return;
          }
          setLoadStatus(`Loading page ${pages} (${total} employees so far)...`);
        },
      });
      if (seq !== fetchSeqRef.current) return;
      if (!result.success) {
        throw new Error('Invalid response');
      }

      const leaveTypeLabels = result.leaveTypeLabels || {};
      const extracted = extractLeaveRecordsFromApiResult(result);
      const leaveRecords =
        Array.isArray(extracted.records) && extracted.records.length > 0
          ? extracted.records
          : Array.isArray(result.leaveRecords)
            ? result.leaveRecords
            : [];
      const labelsForMap = extracted.leaveTypeLabels || leaveTypeLabels;
      const monthWise = getLeaveMonthWise(fromZoho, toZoho);
      const mappedRows = leaveRecords
        .map((row) =>
          mapLeaveRecordToLeaveDataRow(row, labelsForMap, {
            from: fromZoho,
            to: toZoho,
            monthWise,
          })
        )
        .filter(Boolean);

      if (mappedRows.length > 0) {
        setLoadStatus(
          `Saving ${mappedRows.length} row(s) to LeaveData${monthWise ? ` (${monthWise})` : ''}…`
        );
        const saveResult = await saveLeaveDataToBackend({
          records: mappedRows,
          from: fromZoho,
          to: toZoho,
          monthWise,
          apiBase,
          onProgress: (msg) => {
            if (seq === fetchSeqRef.current) setLoadStatus(msg);
          },
        });
        if (seq !== fetchSeqRef.current) return;
        const inserted = saveResult?.data?.inserted ?? mappedRows.length;
        setSaveMessage(
          monthWise
            ? `Saved ${inserted} row(s) to LeaveData for ${monthWise}.`
            : `Saved ${inserted} row(s) to LeaveData table.`
        );
      } else {
        setSaveMessage('No leave rows to save to LeaveData.');
      }

      if (seq !== fetchSeqRef.current) return;
      setData({
        raw: result.raw,
        leaveTypeLabels: labelsForMap,
        leaveRecords,
      });
      setMeta(result.meta || null);
    } catch (err) {
      if (seq !== fetchSeqRef.current) return;
      setError(err.message || 'Failed to fetch leave data');
      setData(null);
      setMeta(null);
      setSaveMessage('');
    } finally {
      if (seq === fetchSeqRef.current) setLoading(false);
    }
  }, [fromDate, toDate, apiBase]);

  const onFromDateChange = (e) => {
    const next = e.target.value;
    setFromDate(next);
    // Keep range month-aligned: To becomes the last day of the From month.
    const monthEnd = getMonthEndIsoFromIso(next);
    if (monthEnd) setToDate(monthEnd);
  };

  const onToDateChange = (e) => {
    setToDate(e.target.value);
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
        <h1 className="leave-title">{pageTitle}</h1>
        <p className="leave-subtitle">
          Choose the month (From / To), then click Fetch. Leave data loads only when you fetch — nothing syncs
          automatically. Earned Leave shows balance and booked days for the selected window.
        </p>
      </header>

      <div className="leave-actions">
        <div className="leave-date-filters">
          <label className="leave-date-label">
            From
            <input
              type="date"
              className="leave-date-input"
              value={fromDate}
              onChange={onFromDateChange}
              aria-label="From date"
            />
          </label>
          <label className="leave-date-label">
            To
            <input
              type="date"
              className="leave-date-input"
              value={toDate}
              onChange={onToDateChange}
              aria-label="To date"
            />
          </label>
          <span className="leave-date-range-hint">
            {isoToDisplayDate(fromDate) || '—'} to {isoToDisplayDate(toDate) || '—'}
          </span>
        </div>
        <button
          type="button"
          className="leave-fetch-btn"
          onClick={() => fetchData()}
          disabled={loading}
        >
          {loading ? 'Fetching...' : 'Fetch'}
        </button>
      </div>

      {error && (
        <div className="leave-error">
          {error}
        </div>
      )}

      {saveMessage && !error && (
        <div className="leave-meta" style={{ color: '#047857', marginBottom: 12 }}>
          {saveMessage}
        </div>
      )}

      {loading && (
        <div className="leave-loading">
          {loadStatus || 'Loading leave data...'}
          <div style={{ fontSize: 13, marginTop: 8, color: '#6b7280' }}>
            Fetching the selected month. Please wait — do not change the dates until loading finishes.
          </div>
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
