import {
  applyFormKGJGujaratEmployeeToRow,
  formatFormKGJHoursOfWorkFromEmployee,
  FORM_KGJ_GJ_CANONICAL_TABLE_HEADERS,
  FORM_KGJ_WEEKLY_HOLIDAY_DEFAULT,
} from './formKGJGujarat';

describe('Form K GJ hours of work', () => {
  it('formats shift start and end for Hours of Work', () => {
    expect(
      formatFormKGJHoursOfWorkFromEmployee({
        ShiftStartTime: '09:00',
        ShiftEndTime: '18:00',
      })
    ).toBe('09:00 to 18:00');
  });

  it('maps employee shift times into the Hours of Work column', () => {
    const row = applyFormKGJGujaratEmployeeToRow(
      {},
      { ShiftStartTime: '09:30', ShiftEndTime: '18:30' },
      FORM_KGJ_GJ_CANONICAL_TABLE_HEADERS,
      { rowIndex: 0 }
    );
    expect(row['Day of Weekly Holiday (4)']).toBe(FORM_KGJ_WEEKLY_HOLIDAY_DEFAULT);
    expect(row['Hours of Work (5)']).toBe('09:30 to 18:30');
  });
});
