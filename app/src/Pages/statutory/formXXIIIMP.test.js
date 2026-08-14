import {
  FORM_XXIII_MP_OT_NIL,
  applyFormXXIIIMPNormalRateToMappedRows,
  applyFormXXIIIMPOtNilToMappedRows,
  copyFormXXIIIMPNormalRateAliasesToRow,
  isFormXXIIIMPNormalRateHeader,
  isFormXXIIIMPOtNilHeader,
  isFormXXIIIMPOtWorkedDatesHeader,
  isFormXXIIIMPTotalOvertimeWorkedHeader,
  isFormXXIIIMPOvertimeRateHeader,
  isFormXXIIIMPOvertimeEarningsHeader,
  isFormXXIIIMPOtWagesPaidDateHeader,
  resolveFormXXIIIMPNormalRateForEmployee,
  resolveFormXXIIIMPPeopleSalary,
} from './formXXIIIMP';

describe('Form XXIII MP Normal rate of wages', () => {
  it('prefers payroll net_pay over April template defaults', () => {
    const emp = { FirstName: 'Prem', LastName: 'Singh Bhati' };
    const rate = resolveFormXXIIIMPNormalRateForEmployee(
      emp,
      { net_pay: 65303 },
      ['2026-04']
    );
    expect(rate).toBe('65303');
  });

  it('uses People MonthlySalary when payroll row is missing', () => {
    const emp = { FirstName: 'Tejpal', LastName: 'Singh', MonthlySalary: 72590 };
    expect(resolveFormXXIIIMPPeopleSalary(emp)).toBe('72590');
    expect(resolveFormXXIIIMPNormalRateForEmployee(emp, null, ['2026-05'])).toBe('72590');
  });

  it('falls back to April template default only when payroll and People salary are missing', () => {
    const emp = { FirstName: 'Prem', LastName: 'Singh Bhati' };
    expect(resolveFormXXIIIMPNormalRateForEmployee(emp, null, ['2026-04'])).toBe('122603');
    expect(resolveFormXXIIIMPNormalRateForEmployee(emp, null, ['2026-05'])).toBe('');
  });

  it('copies alias-key Normal rate onto the visible header used by the autofill grid', () => {
    expect(isFormXXIIIMPNormalRateHeader('Normal rate of wages')).toBe(true);
    const row = copyFormXXIIIMPNormalRateAliasesToRow(
      { 'Normal rate\nof wages': '65303', 'Normal rate of wages': '' },
      ['Normal rate of wages', 'Overtime rate of wages']
    );
    expect(row['Normal rate of wages']).toBe('65303');
    expect(row['Overtime rate of wages']).toBeUndefined();
  });

  it('fills empty autofill cells from payroll so the grid matches Excel download', () => {
    const headers = ['Name of workman', 'Normal rate of wages'];
    const rows = applyFormXXIIIMPNormalRateToMappedRows(
      [
        { 'Name of workman': 'Tejpal Singh', 'Normal rate of wages': '' },
        { 'Name of workman': 'Jeevan Parmar', 'Normal rate\nof wages': '62218' },
      ],
      headers,
      [
        { FirstName: 'Tejpal', LastName: 'Singh' },
        { FirstName: 'Jeevan', LastName: 'Parmar' },
      ],
      [{ net_pay: 65303 }, null],
      ['2026-05']
    );
    expect(rows[0]['Normal rate of wages']).toBe('65303');
    expect(rows[1]['Normal rate of wages']).toBe('62218');
  });
});

describe('Form XXIII MP overtime NIL columns', () => {
  const headers = [
    'Serial No.',
    'Name and surname of workmen',
    "Father's/Husband's name",
    'Sex',
    'Designation/ Nature of employment',
    'Dates on which overtime worked',
    'Total overtime worked or production in case of piece rates',
    'Normal rate of wages',
    'Overtime rate of wages',
    'Overtime earnings',
    'Date on which overtime wages paid',
    'Remarks',
  ];

  it('recognizes the five OT columns for NIL', () => {
    expect(isFormXXIIIMPOtWorkedDatesHeader('Dates on which overtime worked')).toBe(true);
    expect(
      isFormXXIIIMPTotalOvertimeWorkedHeader(
        'Total overtime worked or production in case of piece rates'
      )
    ).toBe(true);
    expect(isFormXXIIIMPOvertimeRateHeader('Overtime rate of wages')).toBe(true);
    expect(isFormXXIIIMPOvertimeEarningsHeader('Overtime earnings')).toBe(true);
    expect(isFormXXIIIMPOtWagesPaidDateHeader('Date on which overtime wages paid')).toBe(true);
    expect(isFormXXIIIMPOtNilHeader('Normal rate of wages')).toBe(false);
    expect(isFormXXIIIMPOtNilHeader('Name and surname of workmen')).toBe(false);
  });

  it('fills blank OT columns with NIL without touching normal rate', () => {
    const rows = applyFormXXIIIMPOtNilToMappedRows(
      [
        {
          'Name and surname of workmen': 'Ravi',
          'Normal rate of wages': '72590',
          'Dates on which overtime worked': '',
          'Overtime rate of wages': '',
        },
      ],
      headers
    );
    expect(rows[0]['Name and surname of workmen']).toBe('Ravi');
    expect(rows[0]['Normal rate of wages']).toBe('72590');
    expect(rows[0]['Dates on which overtime worked']).toBe(FORM_XXIII_MP_OT_NIL);
    expect(rows[0]['Total overtime worked or production in case of piece rates']).toBe(
      FORM_XXIII_MP_OT_NIL
    );
    expect(rows[0]['Overtime rate of wages']).toBe(FORM_XXIII_MP_OT_NIL);
    expect(rows[0]['Overtime earnings']).toBe(FORM_XXIII_MP_OT_NIL);
    expect(rows[0]['Date on which overtime wages paid']).toBe(FORM_XXIII_MP_OT_NIL);
  });

  it('overwrites attendance OT values with NIL when overwrite is true', () => {
    const rows = applyFormXXIIIMPOtNilToMappedRows(
      [
        {
          'Dates on which overtime worked': '2026-05-01',
          'Total overtime worked or production in case of piece rates': '22:10',
          'Overtime rate of wages': '100',
        },
      ],
      headers,
      FORM_XXIII_MP_OT_NIL,
      { overwrite: true }
    );
    expect(rows[0]['Dates on which overtime worked']).toBe(FORM_XXIII_MP_OT_NIL);
    expect(rows[0]['Total overtime worked or production in case of piece rates']).toBe(
      FORM_XXIII_MP_OT_NIL
    );
    expect(rows[0]['Overtime rate of wages']).toBe(FORM_XXIII_MP_OT_NIL);
  });
});
