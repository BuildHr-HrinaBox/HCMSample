import {
  applyFormDRajasthanPayrollToRow,
  applyFormDRJRemarksHoursFromSummary,
  computeFormDRJRemarksHours,
  enrichFormDRajasthanRowsFromAttendance,
  FORM_D_RJ_E_FORM_NOTE,
  FORM_D_RJ_RELAY_MINES_NOTE,
  isFormDRajasthanContext,
  isFormDRJOuterFootnoteText,
  isFormDRJRemarksHoursHeader,
  isFormDRJSummaryDaysHeader,
  normalizeFormDRJOuterFootnoteText,
  readFormDRJEmployeeFullName,
} from './formDRajasthan';

describe('Form D Rajasthan attendance autofill (no shift defaults)', () => {
  test('builds Name from FirstName + MiddleName + LastName', () => {
    expect(
      readFormDRJEmployeeFullName({
        FirstName: 'Prem',
        MiddleName: 'Shankar',
        LastName: 'Menaria',
      })
    ).toBe('Prem Shankar Menaria');
  });

  test('detects Form D_MH context from filename', () => {
    expect(
      isFormDRajasthanContext(
        { title: 'FORM D', subtitle: 'FORMAT OF ATTENDANCE REGISTER' },
        { formFileName: 'Form_D_MH_-_Maharashtra.xlsx', formName: 'Format of Attendance Register' },
        'Form_D_MH_-_Maharashtra.xlsx',
        ''
      )
    ).toBe(true);
  });

  const headers = [
    'Name',
    'Place of work*',
    'Date',
    '1',
    '2',
    '3',
    'Summary No. of Days',
    'Remarks No. of hours',
  ];

  test('does not fill roster ShiftStartTime/ShiftEndTime when FirstIn/LastOut missing', () => {
    const rows = [
      {
        Name: 'Suryakanta',
        Date: 'IN / OUT',
        __formDRJEmployeeIndex: 0,
        __employeeLookupName: 'Suryakanta',
        1: '9.00 / 17.00',
        2: '9.00 / 17.00',
        3: '9.00 / 17.00',
      },
    ];
    const employees = [{ FirstName: 'Suryakanta', id: 'e1' }];
    const attendanceRecords = [
      {
        date: '2025-07-01',
        FirstName: 'Suryakanta',
        id: 'e1',
        Status: 'Present',
        ShiftStartTime: '09:00',
        ShiftEndTime: '17:00',
        FirstIn: '-',
        LastOut: '-',
      },
      {
        date: '2025-07-02',
        FirstName: 'Suryakanta',
        id: 'e1',
        Status: 'Absent',
        ShiftStartTime: '09:00',
        ShiftEndTime: '17:00',
      },
      {
        date: '2025-07-03',
        FirstName: 'Suryakanta',
        id: 'e1',
        Status: 'Present',
        FirstIn: '10:17 AM',
        LastOut: '08:50 PM',
        ShiftStartTime: '09:00',
        ShiftEndTime: '17:00',
      },
    ];

    enrichFormDRajasthanRowsFromAttendance(rows, headers, attendanceRecords, {
      targetYear: 2025,
      targetMonthIndex: 6,
      overwrite: true,
      employees,
    });

    expect(rows[0]['1']).toBe('');
    expect(rows[0]['2']).toBe('');
    expect(rows[0]['3']).toBe('10.17 / 20.50');
  });

  test('clears day cell when no attendance record for that date', () => {
    const rows = [{ Name: 'A', Date: 'IN / OUT', __formDRJEmployeeIndex: 0, 1: '9.00 / 17.00' }];
    enrichFormDRajasthanRowsFromAttendance(rows, headers, [], {
      targetYear: 2025,
      targetMonthIndex: 6,
      overwrite: true,
      employees: [{ FirstName: 'A' }],
    });
    expect(rows[0]['1']).toBe('');
  });
});

describe('Form D Rajasthan Remarks No. of hours', () => {
  test('header detectors', () => {
    expect(isFormDRJSummaryDaysHeader('Summary No. of Days')).toBe(true);
    expect(isFormDRJRemarksHoursHeader('Remarks No. of hours')).toBe(true);
    expect(isFormDRJRemarksHoursHeader('Summary No. of Days')).toBe(false);
  });

  test('computeFormDRJRemarksHours = Summary Days × 8', () => {
    expect(computeFormDRJRemarksHours(26)).toBe('208');
    expect(computeFormDRJRemarksHours('30')).toBe('240');
    expect(computeFormDRJRemarksHours('')).toBe('');
  });

  test('applyFormDRajasthanPayrollToRow fills Summary and Remarks hours', () => {
    const headers = ['Name', 'Summary No. of Days', 'Remarks No. of hours'];
    const row = { Name: 'A', 'Summary No. of Days': '', 'Remarks No. of hours': '' };
    const ok = applyFormDRajasthanPayrollToRow(row, { paid_days: 26 }, headers, {
      overwrite: true,
    });
    expect(ok).toBe(true);
    expect(row['Summary No. of Days']).toBe('26');
    expect(row['Remarks No. of hours']).toBe('208');
  });

  test('applyFormDRJRemarksHoursFromSummary uses existing Summary days', () => {
    const headers = ['Summary No. of Days', 'Remarks No. of hours'];
    const rows = [{ 'Summary No. of Days': '25', 'Remarks No. of hours': '' }];
    const n = applyFormDRJRemarksHoursFromSummary(rows, headers, { overwrite: true });
    expect(n).toBe(1);
    expect(rows[0]['Remarks No. of hours']).toBe('200');
  });
});

describe('Form D Rajasthan outer footnotes', () => {
  test('detects Mines / absence / E Form notes', () => {
    expect(
      isFormDRJOuterFootnoteText(
        '#Relay and *Place of Work in case of Mines only (Underground/Opencast/Surface)'
      )
    ).toBe(true);
    expect(
      isFormDRJOuterFootnoteText(
        'In case an employee is not present the following to be entered: (R for Rest/L for Paid Leave/A for absent/O for Weekly Off/C for Establishment Closed)'
      )
    ).toBe(true);
    expect(
      isFormDRJOuterFootnoteText('** Not necessary in case of E Form maintenance.')
    ).toBe(true);
    expect(isFormDRJOuterFootnoteText('Suryakanta')).toBe(false);
  });

  test('normalizes wrapped footnote fragments', () => {
    expect(
      normalizeFormDRJOuterFootnoteText(
        '#Relay and *Place of Work in case of Mines only (Underground/Open cast/Surface)'
      )
    ).toBe(FORM_D_RJ_RELAY_MINES_NOTE);
    expect(
      normalizeFormDRJOuterFootnoteText('**Not necessary in case of E Form maintenance')
    ).toBe(FORM_D_RJ_E_FORM_NOTE);
  });
});
