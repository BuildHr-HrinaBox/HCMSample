import {
  cacheForm15PayrollTableRows,
  fetchJsonWithTimeout,
  getCachedForm15PayrollTableRows,
} from './statutoryAutofillCache';
import { flattenPayrollEarningColumns } from './payrollEarnings';

const payrollTableInflight = new Map();

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

async function fetchPayrollTablePayloadOnce(payrollMonth, { timeoutMs = 45000, force = false } = {}) {
  const month = String(payrollMonth || '').trim();
  if (!/^\d{4}-\d{2}$/.test(month)) return { records: [], meta: null, payrollMonth: '' };

  if (!force) {
    const cached = getCachedForm15PayrollTableRows([month]);
    if (cached?.rows?.length > 0) {
      return {
        records: cached.rows,
        meta: cached.meta || null,
        payrollMonth: cached.payrollMonth || month,
        source: 'cache',
      };
    }
    if (payrollTableInflight.has(month)) {
      return payrollTableInflight.get(month);
    }
  }

  const attempts = [
    `/server/payroll_function/payroll?${new URLSearchParams({ payroll_month: month })}`,
    `/server/payroll_function?${new URLSearchParams({ payroll_table: '1', payroll_month: month })}`,
  ];

  const task = (async () => {
    for (let i = 0; i < attempts.length; i += 1) {
      try {
        const { resp, json } = await fetchJsonWithTimeout(attempts[i], { cache: 'no-store' }, timeoutMs);
        if (!resp.ok || !json?.success) continue;
        const parsed = parsePayrollTableResponse(json);
        if (parsed.records.length > 0) {
          const payrollMonth = parsed.payrollMonth || month;
          cacheForm15PayrollTableRows(payrollMonth, parsed.records, parsed.meta);
          return {
            records: parsed.records,
            meta: parsed.meta,
            payrollMonth,
          };
        }
      } catch (_) {
        /* try next URL */
      }
    }
    return { records: [], meta: null, payrollMonth: month };
  })();

  if (!force) payrollTableInflight.set(month, task);
  try {
    return await task;
  } finally {
    payrollTableInflight.delete(month);
  }
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

/** Warm payroll table cache for statutory autofill (non-blocking). */
export function prefetchPayrollTableRowsForMonths(monthCandidates, options = {}) {
  const list = Array.isArray(monthCandidates) ? monthCandidates : [];
  const uncached = list.filter((month) => {
    const m = String(month || '').trim();
    if (!/^\d{4}-\d{2}$/.test(m)) return false;
    const hit = getCachedForm15PayrollTableRows([m]);
    return !(hit?.rows?.length > 0);
  });
  if (uncached.length === 0) return Promise.resolve(null);
  return fetchPayrollTableRowsForMonths(uncached, { ...options, timeoutMs: options.timeoutMs || 30000 }).catch(
    () => null
  );
}
