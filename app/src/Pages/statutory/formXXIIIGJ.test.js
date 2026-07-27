import {
  FORM_XXIII_GJ_OT_NIL,
  applyFormXXIIIGJOtNilToMappedRows,
  isFormXXIIIGJContext,
  isFormXXIIIGJOtNilHeader,
  isFormXXIIIGJOtWorkedDatesHeader,
  isFormXXIIIGJTotalOvertimeWorkedHeader,
  isFormXXIIIGJOvertimeRateHeader,
  isFormXXIIIGJOvertimeEarningsHeader,
  isFormXXIIIGJOtWagesPaidDateHeader,
} from './formXXIIIGJ';

describe('Form XXIII GJ overtime NIL columns', () => {
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

  it('detects Gujarat Form XXIII from file name', () => {
    expect(
      isFormXXIIIGJContext({}, { formName: 'Register of Overtime' }, 'Form_XXIII_GJ_-_Gujarat.xlsx', '')
    ).toBe(true);
    expect(
      isFormXXIIIGJContext({}, { formName: 'Register of Overtime' }, 'Form_XXIII_MP.xlsx', '')
    ).toBe(false);
  });

  it('recognizes the five OT columns for NIL', () => {
    expect(isFormXXIIIGJOtWorkedDatesHeader('Dates on which overtime worked')).toBe(true);
    expect(
      isFormXXIIIGJTotalOvertimeWorkedHeader(
        'Total overtime worked or production in case of piece rates'
      )
    ).toBe(true);
    expect(isFormXXIIIGJOvertimeRateHeader('Overtime rate of wages')).toBe(true);
    expect(isFormXXIIIGJOvertimeEarningsHeader('Overtime earnings')).toBe(true);
    expect(isFormXXIIIGJOtWagesPaidDateHeader('Date on which overtime wages paid')).toBe(true);
    expect(isFormXXIIIGJOtNilHeader('Normal rate of wages')).toBe(false);
    expect(isFormXXIIIGJOtNilHeader('Name and surname of workmen')).toBe(false);
  });

  it('fills blank OT columns with NIL without touching normal rate', () => {
    const rows = applyFormXXIIIGJOtNilToMappedRows(
      [
        {
          'Name and surname of workmen': 'Ravi',
          'Normal rate of wages': '30211',
          'Dates on which overtime worked': '',
          'Overtime rate of wages': '',
        },
      ],
      headers
    );
    expect(rows[0]['Name and surname of workmen']).toBe('Ravi');
    expect(rows[0]['Normal rate of wages']).toBe('30211');
    expect(rows[0]['Dates on which overtime worked']).toBe(FORM_XXIII_GJ_OT_NIL);
    expect(rows[0]['Total overtime worked or production in case of piece rates']).toBe(
      FORM_XXIII_GJ_OT_NIL
    );
    expect(rows[0]['Overtime rate of wages']).toBe(FORM_XXIII_GJ_OT_NIL);
    expect(rows[0]['Overtime earnings']).toBe(FORM_XXIII_GJ_OT_NIL);
    expect(rows[0]['Date on which overtime wages paid']).toBe(FORM_XXIII_GJ_OT_NIL);
  });

  it('overwrites attendance OT values with NIL when overwrite is true', () => {
    const rows = applyFormXXIIIGJOtNilToMappedRows(
      [
        {
          'Dates on which overtime worked': '2026-05-09',
          'Total overtime worked or production in case of piece rates': '71:29',
          'Overtime rate of wages': '100',
        },
      ],
      headers,
      FORM_XXIII_GJ_OT_NIL,
      { overwrite: true }
    );
    expect(rows[0]['Dates on which overtime worked']).toBe(FORM_XXIII_GJ_OT_NIL);
    expect(rows[0]['Total overtime worked or production in case of piece rates']).toBe(
      FORM_XXIII_GJ_OT_NIL
    );
    expect(rows[0]['Overtime rate of wages']).toBe(FORM_XXIII_GJ_OT_NIL);
  });
});
