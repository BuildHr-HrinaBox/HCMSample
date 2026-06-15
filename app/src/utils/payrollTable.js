import { fetchJsonWithTimeout } from './statutoryAutofillCache';

async function fetchPayrollTablePayloadOnce(payrollMonth, { timeoutMs = 45000 } = {}) {
  const month = String(payrollMonth || '').trim();
  if (!/^\d{4}-\d{2}$/.test(month)) return { records: [], meta: null };

  const attempts = [
    `/server/payroll_function/payroll?${new URLSearchParams({ payroll_month: month })}`,
    `/server/payroll_function?${new URLSearchParams({ payroll_table: '1', payroll_month: month })}`,
  ];

  for (let i = 0; i < attempts.length; i += 1) {
    try {
      const { resp, json } = await fetchJsonWithTimeout(attempts[i], { cache: 'no-store' }, timeoutMs);
      if (!resp.ok || !json?.success) continue;
      const records = Array.isArray(json.data?.records) ? json.data.records : [];
      if (records.length > 0) {
        return { records, meta: json.data?.meta || null };
      }
    } catch (_) {
      /* try next URL */
    }
  }
  return { records: [], meta: null };
}

/** Load payroll employee rows saved in the Catalyst Payroll table for a month (YYYY-MM). */
export async function fetchPayrollTableRows(payrollMonth, options = {}) {
  const payload = await fetchPayrollTablePayloadOnce(payrollMonth, options);
  return payload.records;
}

/** Try each YYYY-MM candidate until payroll rows are found in the Payroll table. */
export async function fetchPayrollTableRowsForMonths(monthCandidates, options = {}) {
  const list = Array.isArray(monthCandidates) ? monthCandidates : [];
  for (let i = 0; i < list.length; i += 1) {
    const month = String(list[i] || '').trim();
    if (!/^\d{4}-\d{2}$/.test(month)) continue;
    const payload = await fetchPayrollTablePayloadOnce(month, options);
    if (payload.records.length > 0) {
      return {
        payrollMonth: month,
        rows: payload.records,
        meta: payload.meta,
        source: 'table',
      };
    }
  }
  return { payrollMonth: list[0] || '', rows: [], meta: null, source: 'none' };
}
