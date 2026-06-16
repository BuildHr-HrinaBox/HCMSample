import { fetchJsonWithTimeout } from './statutoryAutofillCache';
import { flattenPayrollEarningColumns } from './payrollEarnings';

function normalizePayrollTableRecords(records) {
  if (!Array.isArray(records)) return [];
  return records
    .filter((row) => row && typeof row === 'object' && row.fetch_error !== true)
    .map((row) => flattenPayrollEarningColumns(row));
}

function parsePayrollTableResponse(json) {
  if (!json?.success) return { records: [], meta: null, payrollMonth: '' };
  const data = json.data;
  let records = [];
  let meta = null;
  let payrollMonth = '';
  if (data && typeof data === 'object' && !Array.isArray(data)) {
    records = Array.isArray(data.records) ? data.records : [];
    meta = data.meta || null;
    payrollMonth = String(data.payrollMonth || '').trim();
  } else if (Array.isArray(data)) {
    records = data;
  }
  return {
    records: normalizePayrollTableRecords(records),
    meta,
    payrollMonth,
  };
}

async function fetchPayrollTablePayloadOnce(payrollMonth, { timeoutMs = 45000 } = {}) {
  const month = String(payrollMonth || '').trim();
  if (!/^\d{4}-\d{2}$/.test(month)) return { records: [], meta: null, payrollMonth: '' };

  const attempts = [
    `/server/payroll_function/payroll?${new URLSearchParams({ payroll_month: month })}`,
    `/server/payroll_function?${new URLSearchParams({ payroll_table: '1', payroll_month: month })}`,
  ];

  for (let i = 0; i < attempts.length; i += 1) {
    try {
      const { resp, json } = await fetchJsonWithTimeout(attempts[i], { cache: 'no-store' }, timeoutMs);
      if (!resp.ok || !json?.success) continue;
      const parsed = parsePayrollTableResponse(json);
      if (parsed.records.length > 0) {
        return {
          records: parsed.records,
          meta: parsed.meta,
          payrollMonth: parsed.payrollMonth || month,
        };
      }
    } catch (_) {
      /* try next URL */
    }
  }
  return { records: [], meta: null, payrollMonth: month };
}

async function fetchLatestPayrollTablePayload({ timeoutMs = 45000 } = {}) {
  const attempts = [
    `/server/payroll_function/payroll?${new URLSearchParams({ payroll_table_latest: '1' })}`,
    `/server/payroll_function?${new URLSearchParams({ payroll_table_latest: '1' })}`,
  ];
  for (let i = 0; i < attempts.length; i += 1) {
    try {
      const { resp, json } = await fetchJsonWithTimeout(attempts[i], { cache: 'no-store' }, timeoutMs);
      if (!resp.ok || !json?.success) continue;
      const parsed = parsePayrollTableResponse(json);
      if (parsed.records.length > 0) {
        return {
          records: parsed.records,
          meta: parsed.meta,
          payrollMonth: parsed.payrollMonth,
          source: 'table_latest',
        };
      }
    } catch (_) {
      /* try next URL */
    }
  }
  return { records: [], meta: null, payrollMonth: '', source: 'none' };
}

/** Load payroll employee rows saved in the Catalyst Payroll table for a month (YYYY-MM). */
export async function fetchPayrollTableRows(payrollMonth, options = {}) {
  const payload = await fetchPayrollTablePayloadOnce(payrollMonth, options);
  return payload.records;
}

/** Try each YYYY-MM candidate, then fall back to the latest Payroll table snapshot. */
export async function fetchPayrollTableRowsForMonths(monthCandidates, options = {}) {
  const list = Array.isArray(monthCandidates) ? monthCandidates : [];
  for (let i = 0; i < list.length; i += 1) {
    const month = String(list[i] || '').trim();
    if (!/^\d{4}-\d{2}$/.test(month)) continue;
    const payload = await fetchPayrollTablePayloadOnce(month, options);
    if (payload.records.length > 0) {
      return {
        payrollMonth: payload.payrollMonth || month,
        rows: payload.records,
        meta: payload.meta,
        source: 'table',
      };
    }
  }

  const latest = await fetchLatestPayrollTablePayload(options);
  if (latest.records.length > 0) {
    return {
      payrollMonth: latest.payrollMonth || list[0] || '',
      rows: latest.records,
      meta: latest.meta,
      source: latest.source,
    };
  }

  return { payrollMonth: list[0] || '', rows: [], meta: null, source: 'none' };
}
