import {
  clearSamplePayrollMonthCache,
  fetchSamplePayrollEmployeeRow,
  fetchSamplePayrollRecords,
  fetchSamplePayrollRowsForMonthCandidates,
  samplePayrollRowMatchesEmployeeId,
} from './samplePayrollApi';

describe('samplePayrollApi caching', () => {
  const originalFetch = global.fetch;

  beforeEach(() => {
    clearSamplePayrollMonthCache();
    global.fetch = jest.fn();
  });

  afterEach(() => {
    global.fetch = originalFetch;
    clearSamplePayrollMonthCache();
  });

  function mockMonthPayload(month, records) {
    global.fetch.mockResolvedValue({
      ok: true,
      text: async () =>
        JSON.stringify({
          status: 'success',
          data: {
            payrollMonth: month,
            records,
            meta: { payDate: `${month}-28` },
          },
        }),
    });
  }

  test('fetchSamplePayrollRecords reuses month cache and does not refetch', async () => {
    mockMonthPayload('2026-07', [
      { employeeId: 'E1', employeeName: 'Ada', netpay: 1000, basic: 800, hra: 200 },
    ]);

    const first = await fetchSamplePayrollRecords('2026-07', { timeoutMs: 5000 });
    const second = await fetchSamplePayrollRecords('2026-07', { timeoutMs: 5000 });

    expect(first.records).toHaveLength(1);
    expect(second.records).toHaveLength(1);
    expect(global.fetch).toHaveBeenCalledTimes(1);
  });

  test('fetchSamplePayrollEmployeeRow does not N+1 fetch for the same month', async () => {
    mockMonthPayload('2026-07', [
      { employeeId: 'E1', employeeName: 'Ada', netpay: 1000, basic: 800, hra: 200 },
      { employeeId: 'E2', employeeName: 'Bob', netpay: 1200, basic: 900, hra: 300 },
    ]);

    const row1 = await fetchSamplePayrollEmployeeRow('E1', ['2026-07'], { timeoutMs: 5000 });
    const row2 = await fetchSamplePayrollEmployeeRow('E2', ['2026-07'], { timeoutMs: 5000 });

    expect(row1?.employee_id || row1?.employeeId).toBeTruthy();
    expect(samplePayrollRowMatchesEmployeeId(row1, 'E1')).toBe(true);
    expect(samplePayrollRowMatchesEmployeeId(row2, 'E2')).toBe(true);
    expect(global.fetch).toHaveBeenCalledTimes(1);
  });

  test('fetchSamplePayrollRowsForMonthCandidates returns first month with rows', async () => {
    global.fetch.mockImplementation(async (url) => {
      const href = String(url);
      if (href.includes('payroll_month=2026-06')) {
        return {
          ok: true,
          text: async () =>
            JSON.stringify({ status: 'success', data: { payrollMonth: '2026-06', records: [] } }),
        };
      }
      return {
        ok: true,
        text: async () =>
          JSON.stringify({
            status: 'success',
            data: {
              payrollMonth: '2026-07',
              records: [{ employeeId: 'E9', employeeName: 'Zed', netpay: 1, basic: 1, hra: 1 }],
            },
          }),
      };
    });

    const loaded = await fetchSamplePayrollRowsForMonthCandidates(['2026-06', '2026-07'], {
      timeoutMs: 5000,
    });
    expect(loaded.records).toHaveLength(1);
    expect(loaded.payrollMonth).toBe('2026-07');
  });
});
