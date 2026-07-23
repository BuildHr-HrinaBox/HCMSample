import {
  FORM_XXIII_TN_OT_NIL,
  applyFormXXIIITamilNaduOtNilToMappedRows,
  isFormXXIIITamilNaduContext,
  isFormXXIIITamilNaduOtNilHeader,
  isFormXXIIITamilNaduOtWorkedDatesHeader,
  isFormXXIIITamilNaduTotalOvertimeWorkedHeader,
  isFormXXIIITamilNaduNormalRateHeader,
  isFormXXIIITamilNaduOvertimeRateHeader,
  isFormXXIIITamilNaduOvertimeEarningsHeader,
  isFormXXIIITamilNaduOtWagesPaidDateHeader,
  resolveFormXXIIITamilNaduNormalRate,
  resolveFormXXIIITamilNaduOvertimeRate,
} from './formXXIIITamilNadu';

describe('Form XXIII Tamil Nadu overtime NIL columns', () => {
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

  it('detects Form_XXIII_-_TamilNadu.xlsx context', () => {
    expect(
      isFormXXIIITamilNaduContext(
        { title: 'FORM XXIII', subtitle: 'Register of Overtime' },
        { state: 'Tamil Nadu', formName: 'Form XXIII' },
        'Form_XXIII_-_TamilNadu.xlsx'
      )
    ).toBe(true);
    expect(
      isFormXXIIITamilNaduContext(
        { title: 'FORM XXIII' },
        { state: 'Madhya Pradesh' },
        'Form_XXIII_MP.xlsx'
      )
    ).toBe(false);
  });

  it('recognizes OT NIL columns and wage-rate headers', () => {
    expect(isFormXXIIITamilNaduOtWorkedDatesHeader('Dates on which overtime worked')).toBe(true);
    expect(
      isFormXXIIITamilNaduTotalOvertimeWorkedHeader(
        'Total overtime worked or production in case of piece rates'
      )
    ).toBe(true);
    expect(isFormXXIIITamilNaduNormalRateHeader('Normal rate of wages')).toBe(true);
    expect(isFormXXIIITamilNaduOvertimeRateHeader('Overtime rate of wages')).toBe(true);
    expect(isFormXXIIITamilNaduOvertimeEarningsHeader('Overtime earnings')).toBe(true);
    expect(isFormXXIIITamilNaduOtWagesPaidDateHeader('Date on which overtime wages paid')).toBe(
      true
    );
    expect(isFormXXIIITamilNaduOtNilHeader('Normal rate of wages')).toBe(false);
    expect(isFormXXIIITamilNaduOtNilHeader('Overtime rate of wages')).toBe(false);
    expect(isFormXXIIITamilNaduOtNilHeader('Name and surname of workmen')).toBe(false);
  });

  it('fills blank OT columns with NIL without touching normal or overtime rate', () => {
    const rows = applyFormXXIIITamilNaduOtNilToMappedRows(
      [
        {
          'Name and surname of workmen': 'Ravi',
          'Normal rate of wages': '100851',
          'Overtime rate of wages': '505.08',
          'Dates on which overtime worked': '',
        },
      ],
      headers
    );
    expect(rows[0]['Name and surname of workmen']).toBe('Ravi');
    expect(rows[0]['Normal rate of wages']).toBe('100851');
    expect(rows[0]['Overtime rate of wages']).toBe('505.08');
    expect(rows[0]['Dates on which overtime worked']).toBe(FORM_XXIII_TN_OT_NIL);
    expect(rows[0]['Total overtime worked or production in case of piece rates']).toBe(
      FORM_XXIII_TN_OT_NIL
    );
    expect(rows[0]['Overtime earnings']).toBe(FORM_XXIII_TN_OT_NIL);
    expect(rows[0]['Date on which overtime wages paid']).toBe(FORM_XXIII_TN_OT_NIL);
  });

  it('overwrites attendance OT values with NIL when overwrite is true', () => {
    const rows = applyFormXXIIITamilNaduOtNilToMappedRows(
      [
        {
          'Dates on which overtime worked': '2026-04-01',
          'Total overtime worked or production in case of piece rates': '150:31',
          'Overtime rate of wages': '100',
        },
      ],
      headers,
      FORM_XXIII_TN_OT_NIL,
      { overwrite: true }
    );
    expect(rows[0]['Dates on which overtime worked']).toBe(FORM_XXIII_TN_OT_NIL);
    expect(rows[0]['Total overtime worked or production in case of piece rates']).toBe(
      FORM_XXIII_TN_OT_NIL
    );
    expect(rows[0]['Overtime rate of wages']).toBe('100');
  });
});

describe('Form XXIII Tamil Nadu wage rates', () => {
  it('resolves Normal rate of wages from gross_pay', () => {
    expect(resolveFormXXIIITamilNaduNormalRate({ gross_pay: 100851, net_pay: 90000 })).toBe(
      '100851'
    );
    expect(resolveFormXXIIITamilNaduNormalRate({ net_pay: 90000 })).toBe('');
    expect(resolveFormXXIIITamilNaduNormalRate({ fetch_error: true, gross_pay: 100851 })).toBe('');
  });

  it('resolves Overtime rate of wages as (Basic/26/8)*2', () => {
    // 52628 / 26 / 8 * 2 = 506.038… → 506.04
    expect(resolveFormXXIIITamilNaduOvertimeRate({ basic: 52628 })).toBe('506.04');
    expect(resolveFormXXIIITamilNaduOvertimeRate({ earned_basic: 26000 })).toBe('250');
    expect(resolveFormXXIIITamilNaduOvertimeRate({ gross_pay: 100851 })).toBe('');
    expect(resolveFormXXIIITamilNaduOvertimeRate(null)).toBe('');
  });
});
