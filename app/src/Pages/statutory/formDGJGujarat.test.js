import {
  applyFormDGJGujaratEmployeeToRow,
  applyFormDGJGujaratPaidDaysToRows,
  computeFormDGJGujaratRemarksHours,
  FORM_DGJ_GJ_CANONICAL_TABLE_HEADERS,
  isFormDGJRemarksNoOfHoursHeader,
  isFormDGJSummaryNoOfDaysHeader,
  readFormDGJGujaratPaidDays,
} from './formDGJGujarat';

describe('Form D Gujarat Summary No. of Days ← Paid_days', () => {
  test('detects Summary / Summery No. of Days header', () => {
    expect(isFormDGJSummaryNoOfDaysHeader('Summary No. of Days')).toBe(true);
    expect(isFormDGJSummaryNoOfDaysHeader('Summery No. of Days')).toBe(true);
    expect(isFormDGJSummaryNoOfDaysHeader('Remarks No. of Hours')).toBe(false);
    expect(isFormDGJSummaryNoOfDaysHeader('Name')).toBe(false);
    expect(isFormDGJSummaryNoOfDaysHeader('Relay or Set Work')).toBe(false);
  });

  test('detects Remarks No. of Hours and computes days × 8', () => {
    expect(isFormDGJRemarksNoOfHoursHeader('Remarks No. of Hours')).toBe(true);
    expect(isFormDGJRemarksNoOfHoursHeader('Remarks No Hours')).toBe(true);
    expect(isFormDGJRemarksNoOfHoursHeader('Summary No. of Days')).toBe(false);
    expect(computeFormDGJGujaratRemarksHours(30)).toBe('240');
    expect(computeFormDGJGujaratRemarksHours('26')).toBe('208');
  });

  test('reads Paid_days from Sample Payroll row shapes', () => {
    expect(readFormDGJGujaratPaidDays({ paid_days: 26 })).toBe('26');
    expect(readFormDGJGujaratPaidDays({ Paid_days: '28' })).toBe('28');
    expect(readFormDGJGujaratPaidDays({ fetch_error: true, paid_days: 26 })).toBe('');
  });

  test('applyFormDGJGujaratEmployeeToRow fills Summary and Remarks hours from payrollRow', () => {
    const row = applyFormDGJGujaratEmployeeToRow(
      {},
      { FirstName: 'Vasan', LastName: 'Jagad', Designation: 'Engineer' },
      FORM_DGJ_GJ_CANONICAL_TABLE_HEADERS,
      { payrollRow: { Paid_days: 30 } }
    );
    expect(row.Name).toBe('Vasan Jagad');
    expect(row['Relay or Set Work']).toBe('Engineer');
    expect(row['Summary No. of Days']).toBe('30');
    expect(row['Remarks No. of Hours']).toBe('240');
  });

  test('applyFormDGJGujaratPaidDaysToRows sets Summary days and Remarks hours', () => {
    const headers = [...FORM_DGJ_GJ_CANONICAL_TABLE_HEADERS];
    const rows = [
      { Name: 'Vasan Jagad', 'Summary No. of Days': '', 'Remarks No. of Hours': '' },
      { Name: 'MUTHUKRISHNAN P', 'Summary No. of Days': '', 'Remarks No. of Hours': '' },
    ];
    const hits = applyFormDGJGujaratPaidDaysToRows(
      rows,
      [{ EmployeeID: '1' }, { EmployeeID: '2' }],
      headers,
      {
        resolvePayrollRow: (_emp, _row, index) =>
          index === 0 ? { paid_days: 30 } : { Paid_days: 26 },
      }
    );
    expect(hits).toBe(2);
    expect(rows[0]['Summary No. of Days']).toBe('30');
    expect(rows[0]['Remarks No. of Hours']).toBe('240');
    expect(rows[1]['Summary No. of Days']).toBe('26');
    expect(rows[1]['Remarks No. of Hours']).toBe('208');
  });

  test('matches Summery spelling and falls back to payroll name when emp lookup misses', () => {
    const headers = [
      'Sr. No. in Employee / Workman / Worker Register',
      'Name',
      'Relay or Set Work',
      'Summery No. of Days',
      'Remarks No. of Hours',
    ];
    expect(isFormDGJSummaryNoOfDaysHeader('Summery No. of Days')).toBe(true);
    const rows = [{ Name: 'Vasan Jagad', 'Summery No. of Days': '', 'Remarks No. of Hours': '' }];
    const hits = applyFormDGJGujaratPaidDaysToRows(rows, [null], headers, {
      payrollRows: [{ employee_name: 'Vasan Jagad', Paid_days: 28 }],
      resolvePayrollRow: () => null,
    });
    expect(hits).toBe(1);
    expect(rows[0]['Summery No. of Days']).toBe('28');
    expect(rows[0]['Remarks No. of Hours']).toBe('224');
  });
});
