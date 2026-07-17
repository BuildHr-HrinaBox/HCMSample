import {
  FORM_XXIII_MP_OT_NIL,
  applyFormXXIIIMPOtNilToMappedRows,
  isFormXXIIIMPOtNilHeader,
  isFormXXIIIMPOtWorkedDatesHeader,
  isFormXXIIIMPTotalOvertimeWorkedHeader,
  isFormXXIIIMPOvertimeRateHeader,
  isFormXXIIIMPOvertimeEarningsHeader,
  isFormXXIIIMPOtWagesPaidDateHeader,
} from './formXXIIIMP';

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
