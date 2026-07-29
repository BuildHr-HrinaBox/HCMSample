import {
  applyFormBGJGujaratEmployeeToRow,
  formBGJPayrollRowHasSampleDeductionFields,
  isFormBGJIncomeTaxHeader,
  isFormBGJPaymentDateHeader,
  isFormBGJPfHeader,
  isFormBGJVoluntaryPfHeader,
  resolveFormBGJGujaratPayrollFields,
  resolveFormBGJGujaratPayrollRowsForAutofill,
} from './formBGJGujarat';

describe('Form B Gujarat Sample Payroll deductions', () => {
  test('distinguishes PF, Voluntary Provident Fund, and Income Tax headers', () => {
    expect(isFormBGJPfHeader('PF')).toBe(true);
    expect(isFormBGJPfHeader('Provident Fund')).toBe(true);
    expect(isFormBGJPfHeader('Voluntary Provident Fund')).toBe(false);
    expect(isFormBGJVoluntaryPfHeader('Voluntary Provident Fund')).toBe(true);
    expect(isFormBGJVoluntaryPfHeader('VPF')).toBe(true);
    expect(isFormBGJVoluntaryPfHeader('PF')).toBe(false);
    expect(isFormBGJIncomeTaxHeader('Income Tax')).toBe(true);
    expect(isFormBGJIncomeTaxHeader('TDS')).toBe(true);
  });

  test('resolveFormBGJGujaratPayrollFields reads SamplePayroll PF / VPF / Income Tax', () => {
    const fields = resolveFormBGJGujaratPayrollFields({
      employee_name: 'Test Emp',
      pf: 1800,
      voluntary_provident_fund: 500,
      income_tax: 1200,
      professional_tax: 200,
    });
    expect(fields.pf).toBe(1800);
    expect(fields.voluntaryProvidentFund).toBe(500);
    expect(fields.incomeTax).toBe(1200);
    expect(fields.professionalTax).toBe(200);
  });

  test('resolveFormBGJGujaratPayrollFields reads table column aliases', () => {
    const fields = resolveFormBGJGujaratPayrollFields({
      PF: '1800',
      VoluntaryProvidentFund: '750',
      IncomeTax: '900',
    });
    expect(fields.pf).toBe(1800);
    expect(fields.voluntaryProvidentFund).toBe(750);
    expect(fields.incomeTax).toBe(900);
  });

  test('applyFormBGJGujaratEmployeeToRow fills PF, VPF, and Income Tax from Sample Payroll', () => {
    const headers = ['Name', 'PF', 'Voluntary Provident Fund', 'Income Tax', 'Income Tax'];
    const row = applyFormBGJGujaratEmployeeToRow(
      {},
      { FirstName: 'Test', LastName: 'Emp' },
      headers,
      {
        payrollRow: {
          pf: 1800,
          voluntaryProvidentFund: 500,
          income_tax: 1200,
          professional_tax: 200,
        },
      }
    );
    expect(row.Name).toBe('Test Emp');
    expect(row.PF).toBe('1800');
    expect(row['Voluntary Provident Fund']).toBe('500');
    expect(row['Income Tax']).toBe('1200');
  });

  test('PF column does not receive Professional Tax', () => {
    const row = applyFormBGJGujaratEmployeeToRow(
      {},
      { FirstName: 'Test', LastName: 'Emp' },
      ['PF'],
      {
        payrollRow: {
          pf: 1800,
          professional_tax: 200,
        },
      }
    );
    expect(row.PF).toBe('1800');
    expect(row.PF).not.toBe('200');
  });

  test('resolveFormBGJGujaratPayrollFields prefers flatten epf_contribution / vpf / income_tax', () => {
    const fields = resolveFormBGJGujaratPayrollFields({
      employee_name: 'Test Emp',
      epf_contribution: 9211,
      voluntary_provident_fund: 18422,
      income_tax: 30640,
      hra: 5708,
      gross_pay: 22832,
    });
    expect(fields.pf).toBe(9211);
    expect(fields.voluntaryProvidentFund).toBe(18422);
    expect(fields.incomeTax).toBe(30640);
    expect(fields.hra).toBe(5708);
  });

  test('detects P.F. style PF header', () => {
    expect(isFormBGJPfHeader('P.F.')).toBe(true);
    expect(isFormBGJPfHeader('P F')).toBe(true);
  });

  test('prefer Sample Payroll rows that carry PF / VPF / Income Tax over HRA-only rows', () => {
    const zohoOnly = [{ employee_name: 'A', hra: 5700, gross_pay: 22832 }];
    const sample = [
      {
        employee_name: 'A',
        hra: 5700,
        gross_pay: 22832,
        pf: 9211,
        voluntary_provident_fund: 18422,
        income_tax: 30640,
      },
    ];
    expect(formBGJPayrollRowHasSampleDeductionFields(zohoOnly[0])).toBe(false);
    expect(formBGJPayrollRowHasSampleDeductionFields(sample[0])).toBe(true);
    const rows = resolveFormBGJGujaratPayrollRowsForAutofill(zohoOnly, ['2026-04'], {
      cachedSampleRows: sample,
    });
    expect(rows).toHaveLength(1);
    expect(rows[0].pf).toBe(9211);
    expect(rows[0].voluntary_provident_fund).toBe(18422);
    expect(rows[0].income_tax).toBe(30640);
  });

  test('Date of Payment header matches Sample Payroll Pay date', () => {
    expect(isFormBGJPaymentDateHeader('Date of Payment')).toBe(true);
    expect(isFormBGJPaymentDateHeader('Pay Date')).toBe(true);
    expect(isFormBGJPaymentDateHeader('Payment Date')).toBe(true);
  });

  test('Date of Payment fills from Sample Payroll Pay date meta', () => {
    const fields = resolveFormBGJGujaratPayrollFields(
      { employee_name: 'A', hra: 5700 },
      { payDate: '2026-04-30' }
    );
    expect(fields.paymentDate).toBe('30-04-2026');

    const stamped = resolveFormBGJGujaratPayrollRowsForAutofill(
      [{ employee_name: 'A', hra: 5700 }],
      ['2026-04'],
      {
        cachedSampleRows: [{ employee_name: 'A', hra: 5700, pf: 100 }],
        payDate: '2026-04-30',
      }
    );
    expect(stamped[0].pay_date).toBe('2026-04-30');

    const row = applyFormBGJGujaratEmployeeToRow(
      {},
      { FirstName: 'A', LastName: 'B' },
      ['Date of Payment'],
      {
        payrollRow: stamped[0],
        payDate: '2026-04-30',
      }
    );
    expect(row['Date of Payment']).toBe('30-04-2026');
  });
});
