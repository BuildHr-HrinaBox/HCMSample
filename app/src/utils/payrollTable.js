import {
  cacheForm15PayrollTableRows,
  fetchJsonWithTimeout,
  getCachedForm15PayrollTableRows,
  getLatestCachedPayrollTableRows,
} from './statutoryAutofillCache';
import { getPayrollOrganizationId } from './payrollOrgId';
import { flattenPayrollEarningColumns } from './payrollEarnings';
import {
  fetchLatestSamplePayrollRows,
  fetchSamplePayrollRowsForMonth,
} from './samplePayrollApi';

const payrollTableInflight = new Map();
const PAYROLL_DETAIL_BATCH_DELAY_MS = 350;
const PAYROLL_DETAIL_BATCH_RETRIES = 3;
/** Load salary breakdown / Sample Payroll: 25 employees per Zoho detail call, one call per minute. */
const PAYRUN_DETAIL_BATCH_SIZE = 25;
const PAYRUN_DETAIL_BATCH_INTERVAL_MS = 60_000;
/** Above this count, server-side month batches hit Catalyst timeouts — use per-employee detail. */
const PAYRUN_DETAIL_ONLY_THRESHOLD = 20;

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

function payrollAmountPresent(value) {
  if (value == null || value === '') return false;
  const num = Number(String(value).replace(/,/g, '').trim());
  return Number.isFinite(num) && num !== 0;
}

function payrollRowMissingSalaryBreakdown(row) {
  const flat = flattenPayrollEarningColumns(row);
  const hasBasic = payrollAmountPresent(flat.basic) || payrollAmountPresent(flat.earned_basic);
  const hasHra = payrollAmountPresent(flat.hra) || payrollAmountPresent(flat.hra_fbp);
  return !hasBasic || !hasHra;
}

function normalizePayrollTableRecords(records) {
  if (!Array.isArray(records)) return [];
  return records
    .filter((row) => row && typeof row === 'object' && row.fetch_error !== true)
    .map((row) => flattenPayrollEarningColumns(row));
}

function readPayrollTableSessionCache(month) {
  const cached = getCachedForm15PayrollTableRows([month]);
  if (cached?.rows?.length > 0 && cached.source === 'sample_payroll') {
    return {
      records: cached.rows,
      meta: cached.meta || null,
      payrollMonth: cached.payrollMonth || month,
      source: 'cache',
    };
  }
  return null;
}

async function fetchPayrollTablePayloadOnce(payrollMonth, { timeoutMs = 45000, force = false } = {}) {
  const month = String(payrollMonth || '').trim();
  if (!/^\d{4}-\d{2}$/.test(month)) {
    return { records: [], meta: null, payrollMonth: month, source: 'none' };
  }

  if (!force && payrollTableInflight.has(month)) {
    return payrollTableInflight.get(month);
  }

  const task = (async () => {
    const sampleLoad = await fetchSamplePayrollRowsForMonth(month, { timeoutMs });
    if (sampleLoad.records.length > 0) {
      cacheForm15PayrollTableRows(
        sampleLoad.payrollMonth || month,
        sampleLoad.records,
        sampleLoad.meta,
        '',
        'sample_payroll'
      );
      return {
        records: sampleLoad.records,
        meta: sampleLoad.meta,
        payrollMonth: sampleLoad.payrollMonth || month,
        source: sampleLoad.source || 'sample_payroll',
      };
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
  const sampleLatest = await fetchLatestSamplePayrollRows({ timeoutMs });
  if (sampleLatest.records.length > 0) {
    const resolvedMonth = sampleLatest.payrollMonth || '';
    if (resolvedMonth) {
      cacheForm15PayrollTableRows(
        resolvedMonth,
        sampleLatest.records,
        sampleLatest.meta,
        '',
        'sample_payroll'
      );
    }
    return {
      records: sampleLatest.records,
      meta: sampleLatest.meta,
      payrollMonth: resolvedMonth,
      source: sampleLatest.source || 'sample_payroll_latest',
    };
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

/** Fast pay-run list (no per-employee detail) — paginates through all employees. */
async function fetchZohoPayrollListRowsForMonth(
  payrollMonth,
  { timeoutMs = 90000, batchSize = 50, onProgress = null } = {}
) {
  const month = String(payrollMonth || '').trim();
  if (!/^\d{4}-\d{2}$/.test(month)) return { rows: [], meta: null };

  const organizationId = getPayrollOrganizationId();
  const perBatch = Math.max(1, Math.min(50, Number(batchSize) || 50));
  const merged = [];
  let offset = 0;
  let meta = null;
  let total = null;

  while (true) {
    const qs = new URLSearchParams({
      organization_id: organizationId,
      payroll_month_data: '1',
      payroll_month: month,
      payroll_run_type: 'regular',
      employee_offset: String(offset),
      employee_limit: String(perBatch),
      include_earnings_detail: '0',
    });

    if (typeof onProgress === 'function') {
      const label = Number.isFinite(total) ? `${merged.length} / ${total}` : String(merged.length);
      onProgress(label);
    }

    try {
      const { resp, json } = await fetchJsonWithTimeout(
        `/server/payroll_function?${qs.toString()}`,
        { cache: 'no-store' },
        timeoutMs
      );
      if (!resp.ok || !json?.success) break;
      meta = json.meta || meta;
      if (Number.isFinite(Number(json.meta?.total))) {
        total = Number(json.meta.total);
      }
      const batch = Array.isArray(json.data) ? json.data : [];
      merged.push(...batch);

      const loaded = offset + batch.length;
      let hasMore = json.meta?.has_more === true;
      if (!hasMore && Number.isFinite(total) && total > 0) {
        hasMore = loaded < total;
      } else if (!hasMore && batch.length >= perBatch) {
        hasMore = true;
      }

      if (!hasMore || batch.length === 0) break;
      offset = loaded;
    } catch (_) {
      break;
    }
  }

  const rows = normalizePayrollTableRecords(merged);
  if (rows.length > 0) {
    cacheForm15PayrollTableRows(month, rows, meta);
  }
  return { rows, meta: meta ? { ...meta, total: total ?? meta.total } : null };
}

function pickPayrollEmployeeId(row) {
  const flat = flattenPayrollEarningColumns(row);
  for (const key of ['employee_id', 'employee_number', 'employeeId', 'EmployeeID']) {
    const value = flat[key];
    if (value != null && String(value).trim() !== '') return String(value).trim();
  }
  return '';
}

function mergeListAndDetailPayrollRow(listRow, detailRow) {
  const listFlat =
    listRow && typeof listRow === 'object' && (listRow.basic != null || listRow.earned_basic != null)
      ? listRow
      : flattenPayrollEarningColumns(listRow);
  const detailFlat = flattenPayrollEarningColumns(detailRow);
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
  const flat = flattenPayrollEarningColumns(merged);
  if (!flat.basic && !flat.earned_basic) {
    flat.basic = detailFlat.basic || detailFlat.earned_basic || listFlat.basic || listFlat.earned_basic || '';
    flat.earned_basic = flat.basic;
  }
  if (!flat.hra && !flat.hra_fbp) {
    flat.hra = detailFlat.hra || detailFlat.hra_fbp || listFlat.hra || listFlat.hra_fbp || '';
    flat.hra_fbp = flat.hra;
  }
  return flat;
}

async function fetchPayrollMonthDetailBatch(
  payrollMonth,
  offset,
  perBatch,
  { timeoutMs = 180000, organizationId = getPayrollOrganizationId() } = {}
) {
  const qs = new URLSearchParams({
    organization_id: organizationId,
    payroll_month_data: '1',
    payroll_month: String(payrollMonth || '').trim(),
    payroll_run_type: 'regular',
    employee_offset: String(offset),
    employee_limit: String(perBatch),
    include_earnings_detail: '1',
    detail_limit: String(perBatch),
  });

  let lastError = null;
  for (let attempt = 0; attempt < PAYROLL_DETAIL_BATCH_RETRIES; attempt += 1) {
    try {
      const { resp, json } = await fetchJsonWithTimeout(
        `/server/payroll_function?${qs.toString()}`,
        { cache: 'no-store' },
        timeoutMs
      );
      if (resp.ok && json?.success) {
        return Array.isArray(json.data) ? json.data : [];
      }
      lastError = new Error(json?.error || json?.message || `HTTP ${resp.status}`);
    } catch (err) {
      lastError = err;
    }
    if (attempt + 1 < PAYROLL_DETAIL_BATCH_RETRIES) {
      await sleep(PAYROLL_DETAIL_BATCH_DELAY_MS * (attempt + 1));
    }
  }
  if (lastError) throw lastError;
  return [];
}

/** Load Basic / HRA via payroll_month_data batches (server-side pay-run detail enrichment). */
async function enrichPayrollRowsWithMonthDetailBatches(
  payrollMonth,
  listRows,
  { timeoutMs = 180000, batchSize = 10, onProgress = null } = {}
) {
  const list = Array.isArray(listRows) ? listRows : [];
  const perBatch = Math.max(1, Math.min(10, Number(batchSize) || 10));
  const byId = new Map();

  list.forEach((row) => {
    const flat = flattenPayrollEarningColumns(row);
    const id = pickPayrollEmployeeId(flat);
    if (id) byId.set(id, flat);
  });

  const total = list.length;
  let offset = 0;

  while (offset < total) {
    let batch = [];
    try {
      batch = await fetchPayrollMonthDetailBatch(payrollMonth, offset, perBatch, { timeoutMs });
    } catch (_) {
      offset += perBatch;
      if (typeof onProgress === 'function') {
        onProgress(`${Math.min(offset, total)} / ${total}`);
      }
      if (offset < total) await sleep(PAYROLL_DETAIL_BATCH_DELAY_MS);
      continue;
    }

    batch.forEach((row) => {
      const detailFlat = flattenPayrollEarningColumns(row);
      const id = pickPayrollEmployeeId(detailFlat);
      if (!id) return;
      const existing = byId.get(id) || {};
      byId.set(id, mergeListAndDetailPayrollRow(existing, detailFlat));
    });

    if (batch.length === 0) {
      offset += perBatch;
    } else {
      offset += batch.length;
    }

    if (typeof onProgress === 'function') {
      onProgress(`${Math.min(offset, total)} / ${total}`);
    }

    if (offset < total) {
      await sleep(PAYROLL_DETAIL_BATCH_DELAY_MS);
    }
  }

  return list.map((row) => {
    const id = pickPayrollEmployeeId(row);
    if (id && byId.has(id)) return byId.get(id);
    return flattenPayrollEarningColumns(row);
  });
}

/** Pay-run detail for Basic / HRA — 25 employees per API call, one call per minute. */
async function enrichPayrollRowsWithPayrunDetail(
  rows,
  payrollRunId,
  {
    timeoutMs = 180000,
    batchSize = PAYRUN_DETAIL_BATCH_SIZE,
    batchIntervalMs = PAYRUN_DETAIL_BATCH_INTERVAL_MS,
    onProgress = null,
    onBatch = null,
  } = {}
) {
  const list = Array.isArray(rows) ? rows : [];
  const runId = String(payrollRunId || '').trim();
  if (!runId || list.length === 0) return list;

  const organizationId = getPayrollOrganizationId();
  const perBatch = Math.max(1, Math.min(25, Number(batchSize) || PAYRUN_DETAIL_BATCH_SIZE));
  const intervalMs = Math.max(0, Number(batchIntervalMs) || PAYRUN_DETAIL_BATCH_INTERVAL_MS);
  const updated = list.map((row) => flattenPayrollEarningColumns(row));

  for (let i = 0; i < updated.length; i += perBatch) {
    const batchStart = Date.now();
    const slice = updated.slice(i, i + perBatch);
    const employeeIds = slice.map(pickPayrollEmployeeId).filter(Boolean);

    if (employeeIds.length > 0) {
      try {
        const qs = new URLSearchParams({
          organization_id: organizationId,
          payrun_employee_detail: '1',
          payroll_run_id: runId,
          employee_ids: employeeIds.join(','),
        });
        const { resp, json } = await fetchJsonWithTimeout(
          `/server/payroll_function?${qs.toString()}`,
          { cache: 'no-store' },
          timeoutMs
        );
        if (resp.ok && json?.success) {
          const details = Array.isArray(json.data)
            ? json.data
            : json.data && typeof json.data === 'object'
              ? [json.data]
              : [];
          const byId = new Map();
          details.forEach((detail) => {
            const id = pickPayrollEmployeeId(detail);
            if (id) byId.set(id, detail);
          });
          slice.forEach((row, idx) => {
            const id = pickPayrollEmployeeId(row);
            const detail = id ? byId.get(id) : null;
            updated[i + idx] = detail ? mergeListAndDetailPayrollRow(row, detail) : row;
          });
        }
      } catch (_) {
        /* keep list rows for this batch */
      }
    }

    const done = Math.min(i + perBatch, updated.length);
    if (typeof onProgress === 'function') {
      onProgress(`${done} / ${updated.length}`);
    }
    if (typeof onBatch === 'function') {
      onBatch([...updated], done, updated.length);
    }

    if (done < updated.length && intervalMs > 0) {
      const waitMs = Math.max(0, intervalMs - (Date.now() - batchStart));
      if (waitMs > 0) {
        if (typeof onProgress === 'function') {
          onProgress(
            `${done} / ${updated.length} — waiting ${Math.ceil(waitMs / 1000)}s before next 25`
          );
        }
        await sleep(waitMs);
      }
    }
  }

  return updated;
}

/** Live Zoho pay-run rows when the Catalyst Payroll table has no snapshot for the month yet. */
export async function fetchZohoPayrollRowsForMonth(
  payrollMonth,
  {
    timeoutMs = 90000,
    includeEarningsDetail = false,
    batchSize = null,
    onProgress = null,
    onDetailBatch = null,
    detailMode = null,
    detailBatchSize = PAYRUN_DETAIL_BATCH_SIZE,
    detailBatchIntervalMs = PAYRUN_DETAIL_BATCH_INTERVAL_MS,
  } = {}
) {
  const listLoad = await fetchZohoPayrollListRowsForMonth(payrollMonth, {
    timeoutMs,
    batchSize: batchSize || 50,
    onProgress: includeEarningsDetail
      ? (label) => {
          if (typeof onProgress === 'function') onProgress(`List ${label}`);
        }
      : onProgress,
  });

  if (!includeEarningsDetail || listLoad.rows.length === 0) {
    return listLoad;
  }

  const runId = listLoad.meta?.payroll_run_id;
  const rowCount = listLoad.rows.length;
  const usePayrunDetailOnly =
    detailMode === 'payrun' ||
    !runId ||
    rowCount > PAYRUN_DETAIL_ONLY_THRESHOLD;

  if (usePayrunDetailOnly) {
    if (!runId) return listLoad;
    const enriched = await enrichPayrollRowsWithPayrunDetail(listLoad.rows, runId, {
      timeoutMs: Math.max(timeoutMs, 180000),
      batchSize: detailBatchSize,
      batchIntervalMs: detailBatchIntervalMs,
      onProgress: (label) => {
        if (typeof onProgress === 'function') onProgress(`Salary detail ${label}`);
      },
      onBatch: onDetailBatch,
    });
    return { rows: enriched, meta: listLoad.meta };
  }

  let enriched = await enrichPayrollRowsWithMonthDetailBatches(
    payrollMonth,
    listLoad.rows,
    {
      timeoutMs: Math.max(timeoutMs, 90000),
      batchSize: 5,
      onProgress: (label) => {
        if (typeof onProgress === 'function') onProgress(`Salary detail ${label}`);
      },
    }
  );

  const missingBreakdown = enriched.filter(payrollRowMissingSalaryBreakdown);
  if (missingBreakdown.length > 0 && runId) {
    const fallbackEnriched = await enrichPayrollRowsWithPayrunDetail(missingBreakdown, runId, {
      timeoutMs: Math.max(timeoutMs, 180000),
      batchSize: detailBatchSize,
      batchIntervalMs: detailBatchIntervalMs,
      onProgress: (label) => {
        if (typeof onProgress === 'function') onProgress(`Salary detail (retry) ${label}`);
      },
      onBatch: onDetailBatch,
    });
    const fallbackById = new Map();
    fallbackEnriched.forEach((row) => {
      const id = pickPayrollEmployeeId(row);
      if (id) fallbackById.set(id, row);
    });
    enriched = enriched.map((row) => {
      const id = pickPayrollEmployeeId(row);
      if (id && fallbackById.has(id)) return fallbackById.get(id);
      return row;
    });
  }

  return { rows: enriched, meta: listLoad.meta };
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
