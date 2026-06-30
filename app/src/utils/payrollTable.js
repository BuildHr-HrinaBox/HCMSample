import {
  cacheForm15PayrollTableRows,
  fetchJsonWithTimeout,
  getCachedForm15PayrollTableRows,
  getLatestCachedPayrollTableRows,
} from './statutoryAutofillCache';
import { getPayrollOrganizationId } from './payrollOrgId';
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

function readPayrollTableSessionCache(month) {
  const cached = getCachedForm15PayrollTableRows([month]);
  if (cached?.rows?.length > 0) {
    return {
      records: cached.rows,
      meta: cached.meta || null,
      payrollMonth: cached.payrollMonth || month,
      source: 'cache',
    };
  }
  return null;
}

async function fetchPayrollTableHttpPayload(url, month, timeoutMs) {
  try {
    const { resp, json } = await fetchJsonWithTimeout(url, { cache: 'no-store' }, timeoutMs);
    if (!resp.ok || !json?.success) return null;
    const parsed = parsePayrollTableResponse(json);
    if (parsed.records.length === 0) return null;
    const resolvedMonth = parsed.payrollMonth || month || '';
    if (resolvedMonth) {
      cacheForm15PayrollTableRows(resolvedMonth, parsed.records, parsed.meta);
    }
    return {
      records: parsed.records,
      meta: parsed.meta,
      payrollMonth: resolvedMonth,
      source: 'table',
    };
  } catch (_) {
    return null;
  }
}

/** Real-time read from Payroll table API; falls back to session cache of stored table data. */
async function fetchPayrollTablePayloadOnce(payrollMonth, { timeoutMs = 45000, force = false } = {}) {
  const month = String(payrollMonth || '').trim();
  if (!/^\d{4}-\d{2}$/.test(month)) {
    return { records: [], meta: null, payrollMonth: month, source: 'none' };
  }

  if (!force && payrollTableInflight.has(month)) {
    return payrollTableInflight.get(month);
  }

  const attempts = [
    `/server/payroll_function?${new URLSearchParams({ payroll_table: '1', payroll_month: month })}`,
    `/server/payroll_function/payroll?${new URLSearchParams({ payroll_month: month })}`,
  ];

  const task = (async () => {
    for (let i = 0; i < attempts.length; i += 1) {
      const hit = await fetchPayrollTableHttpPayload(attempts[i], month, timeoutMs);
      if (hit) return hit;
    }
    const cached = readPayrollTableSessionCache(month);
    if (cached) return cached;
    return { records: [], meta: null, payrollMonth: month, source: 'none' };
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
    `/server/payroll_function/payroll/latest`,
    `/server/payroll_function/payroll?${new URLSearchParams({ payroll_table_latest: '1' })}`,
    `/server/payroll_function?${new URLSearchParams({ payroll_table_latest: '1' })}`,
  ];
  for (let i = 0; i < attempts.length; i += 1) {
    const hit = await fetchPayrollTableHttpPayload(attempts[i], '', timeoutMs);
    if (hit) {
      return { ...hit, source: 'table_latest' };
    }
  }
  const latestCached = getLatestCachedPayrollTableRows();
  if (latestCached?.rows?.length > 0) {
    return {
      records: latestCached.rows,
      meta: latestCached.meta || null,
      payrollMonth: latestCached.payrollMonth || '',
      source: 'cache_latest',
    };
  }
  return { records: [], meta: null, payrollMonth: '', source: 'none' };
}

/** Live Zoho pay-run rows when the Catalyst Payroll table has no snapshot for the month yet. */
async function fetchZohoPayrollRowsForMonth(payrollMonth, { timeoutMs = 45000 } = {}) {
  const month = String(payrollMonth || '').trim();
  if (!/^\d{4}-\d{2}$/.test(month)) return { rows: [], meta: null };

  const organizationId = getPayrollOrganizationId();
  const batchSize = 25;
  const merged = [];
  let offset = 0;
  let meta = null;

  while (true) {
    const qs = new URLSearchParams({
      organization_id: organizationId,
      payroll_month_data: '1',
      payroll_month: month,
      payroll_run_type: 'regular',
      employee_offset: String(offset),
      employee_limit: String(batchSize),
      include_earnings_detail: '0',
    });
    try {
      const { resp, json } = await fetchJsonWithTimeout(
        `/server/payroll_function?${qs.toString()}`,
        { cache: 'no-store' },
        timeoutMs
      );
      if (!resp.ok || !json?.success) break;
      meta = json.meta || meta;
      const batch = Array.isArray(json.data) ? json.data : [];
      merged.push(...batch);
      const hasMore =
        json.meta?.has_more === true ||
        (json.meta?.has_more !== false && batch.length >= batchSize);
      if (!hasMore || batch.length === 0) break;
      offset += batch.length;
    } catch (_) {
      break;
    }
  }

  const rows = normalizePayrollTableRecords(merged);
  if (rows.length > 0) {
    cacheForm15PayrollTableRows(month, rows, meta);
  }
  return { rows, meta };
}

/** Load payroll employee rows saved in the Catalyst Payroll table for a month (YYYY-MM). */
export async function fetchPayrollTableRows(payrollMonth, options = {}) {
  const payload = await fetchPayrollTablePayloadOnce(payrollMonth, options);
  return payload.records;
}

/** Synchronous Payroll table read for instant statutory autofill paint (session cache only). */
export function getPayrollTableRowsForStatutoryAutofillSync(monthCandidates = []) {
  const cached = getCachedForm15PayrollTableRows(monthCandidates);
  if (cached?.rows?.length > 0) {
    return {
      payrollMonth: cached.payrollMonth || monthCandidates[0] || '',
      rows: cached.rows,
      meta: cached.meta || null,
      source: cached.source || 'cache',
    };
  }
  const latest = getLatestCachedPayrollTableRows();
  if (latest?.rows?.length > 0) {
    return {
      payrollMonth: latest.payrollMonth || '',
      rows: latest.rows,
      meta: latest.meta || null,
      source: 'cache_latest',
    };
  }
  return null;
}

/** Real-time Payroll table load for statutory autofill (API first, then stored table cache). */
export async function loadPayrollTableRowsForStatutoryAutofill(monthCandidates = [], options = {}) {
  return fetchPayrollTableRowsForMonths(monthCandidates, options);
}

/** Try each YYYY-MM candidate, latest table snapshot, then live Zoho pay-run as last resort. */
export async function fetchPayrollTableRowsForMonths(monthCandidates, options = {}) {
  const list = Array.isArray(monthCandidates) ? monthCandidates : [];
  const timeoutMs = Number(options.timeoutMs) > 0 ? Number(options.timeoutMs) : 45000;
  const zohoFallback = options.zohoFallback !== false;

  for (let i = 0; i < list.length; i += 1) {
    const month = String(list[i] || '').trim();
    if (!/^\d{4}-\d{2}$/.test(month)) continue;
    const payload = await fetchPayrollTablePayloadOnce(month, { ...options, timeoutMs });
    if (payload.records.length > 0) {
      return {
        payrollMonth: payload.payrollMonth || month,
        rows: payload.records,
        meta: payload.meta,
        source: payload.source || 'table',
      };
    }
  }

  const latest = await fetchLatestPayrollTablePayload({ timeoutMs });
  if (latest.records.length > 0) {
    return {
      payrollMonth: latest.payrollMonth || list[0] || '',
      rows: latest.records,
      meta: latest.meta,
      source: latest.source,
    };
  }

  if (zohoFallback && list.length > 0) {
    for (let i = 0; i < list.length; i += 1) {
      const month = String(list[i] || '').trim();
      if (!/^\d{4}-\d{2}$/.test(month)) continue;
      const zohoLoad = await fetchZohoPayrollRowsForMonth(month, { timeoutMs });
      if (zohoLoad.rows.length > 0) {
        console.info(
          `Payroll table empty for ${month}; using live Zoho pay-run (${zohoLoad.rows.length} row(s)). Save payroll on the Payroll page to store in the Payroll table.`
        );
        return {
          payrollMonth: month,
          rows: zohoLoad.rows,
          meta: zohoLoad.meta,
          source: 'zoho_live',
        };
      }
    }
  }

  return { payrollMonth: list[0] || '', rows: [], meta: null, source: 'none' };
}

/** Warm payroll table cache for statutory autofill (non-blocking, table only — no Zoho). */
export function prefetchPayrollTableRowsForMonths(monthCandidates, options = {}) {
  const list = Array.isArray(monthCandidates) ? monthCandidates : [];
  if (list.length === 0) return Promise.resolve(null);
  return fetchPayrollTableRowsForMonths(list, {
    ...options,
    timeoutMs: options.timeoutMs || 30000,
    zohoFallback: false,
  }).catch(() => null);
}
