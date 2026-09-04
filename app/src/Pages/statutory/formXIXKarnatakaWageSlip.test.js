import {
  FORM_XIX_KA_ALL_TABLE_HEADERS,
  FORM_XIX_KA_OT_NIL,
  FORM_XIX_KA_RATE_DEFAULT,
  applyFormXIXKarnatakaAutofillFromSite,
  applyFormXIXKarnatakaEmployeeToRow,
  isFormXIXKARateHeader,
  isFormXIXKAUnitsHeader,
  resolveFormXIXKarnatakaPayrollFields,
  syncFormXIXKarnatakaDeductionsFromGrossNet,
} from './formXIXKarnatakaWageSlip';

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
});
