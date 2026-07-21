import {
  applyFormXVIIITamilNaduOvertimeNilToMappedRows,
  computeFormXVIIITamilNaduOtherCashPayments,
  FORM_XVIII_TN_OVERTIME_NIL,
  formXVIIITamilNaduRowsRicherThan,
  isFormXVIIITamilNaduDailyAttendanceHeader,
  isFormXVIIITamilNaduDailyRateHeader,
  isFormXVIIITamilNaduOtherCashPaymentsHeader,
  isFormXVIIITamilNaduOvertimeHeader,
  isFormXVIIITamilNaduSerialRegisterHeader,
  isFormXVIIITamilNaduTotalAttendanceHeader,
} from './formXVIIITamilNaduWagesMuster';

describe('Form XVIII Tamil Nadu payroll column headers', () => {
  it('detects SL.No.in register of workmen (sl.no, not only serial)', () => {
    expect(isFormXVIIITamilNaduSerialRegisterHeader('SL.No.in register of workmen')).toBe(true);
    expect(
      isFormXVIIITamilNaduSerialRegisterHeader('Serial Number in the Register of Workmen')
    ).toBe(true);
    expect(isFormXVIIITamilNaduSerialRegisterHeader('Sl No.')).toBe(false);
    expect(isFormXVIIITamilNaduSerialRegisterHeader('Name of employee')).toBe(false);
  });

  it('defaults Overtime to Nil when empty and keeps real OT amounts', () => {
    expect(isFormXVIIITamilNaduOvertimeHeader('Overtime')).toBe(true);
    expect(isFormXVIIITamilNaduOvertimeHeader('Other cash payments')).toBe(false);
    const headers = ['Name of employee', 'Overtime', 'Basic wages'];
    const rows = applyFormXVIIITamilNaduOvertimeNilToMappedRows(
      [
        { 'Name of employee': 'rajesh', Overtime: '', 'Basic wages': 100 },
        { 'Name of employee': 'raja', Overtime: 2500, 'Basic wages': 200 },
      ],
      headers,
      FORM_XVIII_TN_OVERTIME_NIL,
      { overwrite: false }
    );
    expect(rows[0].Overtime).toBe('Nil');
    expect(rows[1].Overtime).toBe(2500);
  });

  it('distinguishes Daily attendance from Total attendance', () => {
    expect(
      isFormXVIIITamilNaduDailyAttendanceHeader('Daily attendance /units worked')
    ).toBe(true);
    expect(
      isFormXVIIITamilNaduTotalAttendanceHeader('Daily attendance /units worked')
    ).toBe(false);
    expect(
      isFormXVIIITamilNaduTotalAttendanceHeader('Total attendance/ units of work done')
    ).toBe(true);
    expect(
      isFormXVIIITamilNaduDailyAttendanceHeader('Total attendance/ units of work done')
    ).toBe(false);
  });

  it('detects daily rate and other cash payment headers', () => {
    expect(
      isFormXVIIITamilNaduDailyRateHeader('Daily rate of wages/piece-rate')
    ).toBe(true);
    expect(
      isFormXVIIITamilNaduOtherCashPaymentsHeader(
        'Other cash payments (nature of payment to be indicated)'
      )
    ).toBe(true);
  });

  it('computes Other cash payments as gross_pay − basic − hra', () => {
    expect(computeFormXVIIITamilNaduOtherCashPayments(100000, 40000, 20000)).toBe(40000);
    expect(computeFormXVIIITamilNaduOtherCashPayments('113708', '50000', '20000')).toBe(43708);
    expect(computeFormXVIIITamilNaduOtherCashPayments(70326, '', '')).toBe('');
    expect(computeFormXVIIITamilNaduOtherCashPayments(74132, 74132, 0)).toBe(0);
  });

  it('scores autofill rows richer than sparse name/basic SampleData', () => {
    const headers = [
      'Name of employee',
      'Designation/nature of work',
      'Daily attendance /units worked',
      'Total attendance/ units of work done',
      'Daily rate of wages/piece-rate',
      'Basic wages',
    ];
    const rich = [
      {
        'Name of employee': 'rajesh',
        'Designation/nature of work': 'Engineer',
        'Daily attendance /units worked': 31,
        'Total attendance/ units of work done': 31,
        'Daily rate of wages/piece-rate': 70562,
        'Basic wages': 26781,
      },
    ];
    const sparse = [{ 'Name of employee': 'rajesh', 'Basic wages': 26781 }];
    expect(formXVIIITamilNaduRowsRicherThan(rich, headers, sparse, headers)).toBe(true);
    expect(formXVIIITamilNaduRowsRicherThan(sparse, headers, rich, headers)).toBe(false);
  });
});
