import {
  FORM_XIX_KA_ALL_TABLE_HEADERS,
  FORM_XIX_KA_OT_NIL,
  FORM_XIX_KA_RATE_DEFAULT,
  applyFormXIXKarnatakaAutofillFromSite,
  applyFormXIXKarnatakaEmployeeToRow,
  isFormXIXKARateHeader,
  isFormXIXKAUnitsHeader,
  resolveFormXIXKarnatakaDaysWorked,
  resolveFormXIXKarnatakaPayrollFields,
  syncFormXIXKarnatakaDeductionsFromGrossNet,
} from './formXIXKarnatakaWageSlip';
import { buildKarnatakaPayrollRowResolver } from './form10TamilNadu';

describe('Form XIX Karnataka wage slip', () => {
  it('does not treat units column as the rate column', () => {
    expect(isFormXIXKAUnitsHeader('No. of units worked in case of piece rate')).toBe(true);
    expect(isFormXIXKARateHeader('No. of units worked in case of piece rate')).toBe(false);
    expect(isFormXIXKARateHeader('Rate of daily wages/piece - rate')).toBe(true);
  });

  it('computes Deductions, any as gross_pay − net_pay (not total_deductions)', () => {
    const fields = resolveFormXIXKarnatakaPayrollFields({
      gross_pay: 71392,
      net_pay: 67928,
      total_deductions: 0,
    });
    expect(String(fields.grossWages)).toBe('71392');
    expect(String(fields.netWages)).toBe('67928');
    expect(Number(fields.deductions)).toBe(3464);
  });

  it('syncs deductions on a grid row from Gross − Actual wages paid', () => {
    const row = {
      'Gross wages payable': '107870',
      'Deductions, any': '0',
      'If Actual wages paid': '102278',
    };
    syncFormXIXKarnatakaDeductionsFromGrossNet(row, FORM_XIX_KA_ALL_TABLE_HEADERS);
    expect(Number(row['Deductions, any'])).toBe(5592);
  });

  it('autofills Monthly Wages only on rate, leaves units blank, and sets gross−net deductions', () => {
    const row = applyFormXIXKarnatakaEmployeeToRow(
      {},
      { FirstName: 'Test', LastName: 'Worker', EmployeeID: 'VE0712', Sex: 'Male' },
      FORM_XIX_KA_ALL_TABLE_HEADERS,
      {
        payrollRow: {
          paid_days: 31,
          gross_pay: 71392,
          net_pay: 67928,
          total_deductions: 0,
        },
      }
    );
    expect(row['Rate of daily wages/piece - rate']).toBe(FORM_XIX_KA_RATE_DEFAULT);
    expect(String(row['No. of units worked in case of piece rate'] ?? '').trim()).toBe('');
    expect(Number(row['Deductions, any'])).toBe(3464);
    expect(String(row['If Actual wages paid'])).toBe('67928');
  });

  it('puts NIL under overtime dates and overtime hours/amount columns', () => {
    const row = applyFormXIXKarnatakaEmployeeToRow(
      {},
      { FirstName: 'Test', LastName: 'Worker', EmployeeID: 'VE0712' },
      FORM_XIX_KA_ALL_TABLE_HEADERS,
      {
        payrollRow: {
          paid_days: 31,
          overtime_date: '2026-05-12',
          overtime_hours: 4,
          overtime: 800,
          gross_pay: 71392,
          net_pay: 67928,
        },
      }
    );
    expect(row['Dates on which overtime worked']).toBe(FORM_XIX_KA_OT_NIL);
    expect(row['Overtime hours and amount of overtime wages']).toBe(FORM_XIX_KA_OT_NIL);
  });

  it('writes company name and address into the establishment header field', () => {
    const next = applyFormXIXKarnatakaAutofillFromSite(
      {},
      {
        contractorText: 'VAYONA ENERGY PRIVATE LIMITED',
        establishmentText: 'VAYONA ENERGY PRIVATE LIMITED, Company Street, Bengaluru, Karnataka',
        natureLocationText: 'KA-Bableshwar',
        principalEmployerText: 'M/s Clean Wind Power Bableshwar Pvt Ltd.',
      }
    );
    expect(next.form_xix_ka_establishment).toBe(
      'VAYONA ENERGY PRIVATE LIMITED, Company Street, Bengaluru, Karnataka'
    );
  });

  it('matches payroll by first name and last name, including People FirstName-only records', () => {
    const payroll = [
      { first_name: 'Selva', last_name: 'Kumar', gross_pay: 100, net_pay: 90 },
      { first_name: 'Umesh', last_name: 'S O', gross_pay: 200, net_pay: 180 },
    ];
    const resolve = buildKarnatakaPayrollRowResolver(payroll);
    expect(resolve({ FirstName: 'Umesh S O' })).toEqual(payroll[1]);
    expect(resolve({ FirstName: 'Selva', LastName: 'Kumar' })).toEqual(payroll[0]);
    expect(resolve({ FirstName: 'Selva' })).toBeNull();
  });

  it('keeps payroll paid_days instead of May calendar default and separates father name', () => {
    const row = applyFormXIXKarnatakaEmployeeToRow(
      {},
      { FirstName: 'Umesh S O', Father_s_Name: 'Shivaputrappa oli', EmployeeID: 'VE1239' },
      FORM_XIX_KA_ALL_TABLE_HEADERS,
      {
        monthCandidates: ['2026-05'],
        payrollRow: {
          first_name: 'Umesh',
          last_name: 'S O',
          paid_days: 26,
          gross_pay: 50000,
          net_pay: 45000,
        },
      }
    );
    expect(String(row['No. of days worked'])).toBe('26');
    expect(row["Name and Father's/Husband's Name of the workman"]).toBe(
      'Umesh S O / Shivaputrappa oli'
    );
    expect(String(row['Gross wages payable'])).toBe('50000');
    expect(String(row['If Actual wages paid'])).toBe('45000');
  });

  it('fills No. of days worked from paid_days only when payroll has first_name and last_name', () => {
    expect(
      resolveFormXIXKarnatakaDaysWorked({
        paid_days: 26,
        first_name: 'Umesh',
        last_name: 'S O',
      })
    ).toBe('26');
    expect(
      resolveFormXIXKarnatakaDaysWorked({
        paid_days: 26,
        first_name: 'Umesh',
      })
    ).toBe('');
    expect(resolveFormXIXKarnatakaDaysWorked({ paid_days: 31 })).toBe('');

    const unnamed = applyFormXIXKarnatakaEmployeeToRow(
      { 'No. of days worked': '31' },
      { FirstName: 'Umesh S O' },
      FORM_XIX_KA_ALL_TABLE_HEADERS,
      {
        monthCandidates: ['2026-05'],
        payrollRow: { paid_days: 26, gross_pay: 50000, net_pay: 45000 },
      }
    );
    expect(String(unnamed['No. of days worked'] ?? '').trim()).toBe('');

    const firstOnly = applyFormXIXKarnatakaEmployeeToRow(
      {},
      { FirstName: 'Selva' },
      FORM_XIX_KA_ALL_TABLE_HEADERS,
      {
        monthCandidates: ['2026-05'],
        payrollRow: { first_name: 'Selva', paid_days: 31, gross_pay: 100, net_pay: 90 },
      }
    );
    expect(String(firstOnly['No. of days worked'] ?? '').trim()).toBe('');
  });
});
