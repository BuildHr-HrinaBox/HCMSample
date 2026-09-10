import {
  FORM_XXIII_KA_OT_NIL,
  applyFormXXIIIKarnatakaOtNilToMappedRows,
  isFormXXIIIKarnatakaOtNilHeader,
  isFormXXIIIKarnatakaTotalOvertimeWorkedHeader,
  isFormXXIIIKarnatakaOvertimeRateHeader,
  isFormXXIIIKarnatakaOvertimeEarningsHeader,
  isFormXXIIIKarnatakaOtWagesPaidDateHeader,
} from './formXXIIIKarnataka';
import { buildKarnatakaPayrollRowResolver } from './form10TamilNadu';

describe('Form XXIII Karnataka overtime NIL columns', () => {
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

  it('recognizes the four OT columns for NIL', () => {
    expect(
      isFormXXIIIKarnatakaTotalOvertimeWorkedHeader(
        'Total overtime worked or production in case of piece rates'
      )
    ).toBe(true);
    expect(isFormXXIIIKarnatakaOvertimeRateHeader('Overtime rate of wages')).toBe(true);
    expect(isFormXXIIIKarnatakaOvertimeEarningsHeader('Overtime earnings')).toBe(true);
    expect(isFormXXIIIKarnatakaOtWagesPaidDateHeader('Date on which overtime wages paid')).toBe(
      true
    );
    expect(isFormXXIIIKarnatakaOtNilHeader('Normal rate of wages')).toBe(false);
    expect(isFormXXIIIKarnatakaOtNilHeader('Dates on which overtime worked')).toBe(false);
    expect(isFormXXIIIKarnatakaOtNilHeader('Name and surname of workmen')).toBe(false);
  });

  it('fills blank OT columns with NIL without touching normal rate', () => {
    const rows = applyFormXXIIIKarnatakaOtNilToMappedRows(
      [
        {
          'Name and surname of workmen': 'Ravi',
          'Normal rate of wages': '30211',
          'Overtime rate of wages': '',
          'Overtime earnings': '',
        },
      ],
      headers
    );
    expect(rows[0]['Name and surname of workmen']).toBe('Ravi');
    expect(rows[0]['Normal rate of wages']).toBe('30211');
    expect(rows[0]['Total overtime worked or production in case of piece rates']).toBe(
      FORM_XXIII_KA_OT_NIL
    );
    expect(rows[0]['Overtime rate of wages']).toBe(FORM_XXIII_KA_OT_NIL);
    expect(rows[0]['Overtime earnings']).toBe(FORM_XXIII_KA_OT_NIL);
    expect(rows[0]['Date on which overtime wages paid']).toBe(FORM_XXIII_KA_OT_NIL);
    expect(rows[0]['Dates on which overtime worked']).toBeUndefined();
  });

  it('overwrites attendance OT values with NIL when overwrite is true', () => {
    const rows = applyFormXXIIIKarnatakaOtNilToMappedRows(
      [
        {
          'Total overtime worked or production in case of piece rates': '71:29',
          'Overtime rate of wages': '100',
          'Overtime earnings': '500',
          'Date on which overtime wages paid': '2026-05-09',
        },
      ],
      headers,
      FORM_XXIII_KA_OT_NIL,
      { overwrite: true }
    );
    expect(rows[0]['Total overtime worked or production in case of piece rates']).toBe(
      FORM_XXIII_KA_OT_NIL
    );
    expect(rows[0]['Overtime rate of wages']).toBe(FORM_XXIII_KA_OT_NIL);
    expect(rows[0]['Overtime earnings']).toBe(FORM_XXIII_KA_OT_NIL);
    expect(rows[0]['Date on which overtime wages paid']).toBe(FORM_XXIII_KA_OT_NIL);
  });

  it('matches payroll by first name and last name, including People FirstName-only records', () => {
    const payroll = [
      { first_name: 'Ashok', last_name: 'Kumar', net_pay: 100 },
      { first_name: 'Ashok', last_name: 'Jangamashetti', net_pay: 200 },
      { first_name: 'Naveen', last_name: 'K M', net_pay: 300 },
    ];
    const resolve = buildKarnatakaPayrollRowResolver(payroll);
    expect(resolve({ FirstName: 'Ashok', LastName: 'Jangamashetti' })).toEqual(payroll[1]);
    expect(resolve({ FirstName: 'Naveen K M' })).toEqual(payroll[2]);
    expect(resolve({ FirstName: 'Ashok' })).toBeNull();
  });
});
