import {
  applyFormBGJGujaratEmployeeToRow,
  formBGJPayrollRowHasSampleDeductionFields,
  isFormBGJIncomeTaxHeader,
  isFormBGJInsuranceHeader,
  isFormBGJPaymentDateHeader,
  isFormBGJPfHeader,
  isFormBGJRecoveriesHeader,
  isFormBGJVoluntaryPfHeader,
  resolveFormBGJGujaratPayrollFields,
  resolveFormBGJGujaratPayrollRowsForAutofill,
  resolveFormBGJGujaratTableHeaders,
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

  test('skips HRA-only cached rows when bulk Sample Payroll has PF / VPF / Income Tax', () => {
    const zohoCached = [{ employee_name: 'A', hra: 5700, gross_pay: 22832 }];
    const sampleBulk = [
      {
        employee_name: 'A',
        hra: 5700,
        gross_pay: 22832,
        PF: 1800,
        VoluntaryProvidentFund: 500,
        IncomeTax: 1200,
      },
    ];
    const rows = resolveFormBGJGujaratPayrollRowsForAutofill(zohoCached, ['2026-05'], {
      cachedSampleRows: zohoCached,
      bulkSampleRows: sampleBulk,
    });
    expect(rows).toHaveLength(1);
    expect(formBGJPayrollRowHasSampleDeductionFields(rows[0])).toBe(true);
    const fields = resolveFormBGJGujaratPayrollFields(rows[0]);
    expect(fields.pf).toBe(1800);
    expect(fields.voluntaryProvidentFund).toBe(500);
    expect(fields.incomeTax).toBe(1200);
  });

  test('Recoveries stays blank while second Total gets gross_pay − net_pay', () => {
    expect(isFormBGJRecoveriesHeader('Recoveries')).toBe(true);
    expect(isFormBGJInsuranceHeader('Insurance')).toBe(true);
    expect(isFormBGJInsuranceHeader('Income Tax')).toBe(false);

    const fields = resolveFormBGJGujaratPayrollFields({
      gross_pay: 63622,
      net_pay: 50000,
      pf: 1800,
      income_tax: 1200,
    });
    expect(fields.deductionsTotal).toBe(13622);

    const headers = [
      'Name',
      'Total',
      'PF',
      'Voluntary Provident Fund',
      'Income Tax',
      'Insurance',
      'Others',
      'Recoveries',
      'Total ',
    ];
    const row = applyFormBGJGujaratEmployeeToRow(
      {},
      { FirstName: 'Test', LastName: 'Emp' },
      headers,
      {
        payrollRow: {
          gross_pay: 63622,
          net_pay: 50000,
          pf: 1800,
          income_tax: 1200,
        },
      }
    );
    expect(row.Total).toBe('63622');
    expect(row['Total ']).toBe('13622');
    expect(row.PF).toBe('1800');
    expect(row.Insurance).toBe('');
    expect(row.Others).toBe('');
    expect(row.Recoveries).toBe('');
  });

  test('PF before Total uses gross_pay', () => {
    const row = applyFormBGJGujaratEmployeeToRow(
      {},
      { FirstName: 'Test', LastName: 'Emp' },
      ['Name', 'PF', 'Total', 'Net Payment'],
      {
        payrollRow: {
          gross_pay: 63622,
          net_pay: 50000,
          pf: 1800,
        },
      }
    );
    expect(row.PF).toBe('63622');
    expect(row['Net Payment']).toBe('50000');
  });

  test('Total right after HRA uses gross_pay (not deductions)', () => {
    const row = applyFormBGJGujaratEmployeeToRow(
      {},
      { FirstName: 'Test', LastName: 'Emp' },
      ['HRA', 'Total', 'PF'],
      {
        payrollRow: {
          hra: 10935,
          gross_pay: 63622,
          net_pay: 60625,
          pf: 2797,
        },
      }
    );
    expect(row.HRA).toBe('10935');
    expect(row.Total).toBe('63622');
    expect(row.PF).toBe('2797');
  });

  test('with two Total columns: first is gross, second is gross minus net', () => {
    const headers = ['HRA', 'Total', 'PF', 'Recoveries', 'Total ', 'Net Payment'];
    const row = applyFormBGJGujaratEmployeeToRow(
      {},
      { FirstName: 'Test', LastName: 'Emp' },
      headers,
      {
        payrollRow: {
          hra: 10935,
          gross_pay: 63622,
          net_pay: 60625,
          pf: 2797,
        },
      }
    );
    expect(row.HRA).toBe('10935');
    expect(row.PF).toBe('2797');
    expect(row.Recoveries).toBe('');
    expect(row.Total).toBe('63622');
    expect(row['Total ']).toBe('2997');
    expect(row['Net Payment']).toBe('60625');
  });

  test('with duplicate "Total" captions, values stay in separate columns', () => {
    const rawHeaders = ['HRA', 'Total', 'PF', 'Recoveries', 'Total', 'Net Payment'];
    const headers = resolveFormBGJGujaratTableHeaders(rawHeaders);
    const firstTotal = headers[1];
    const secondTotal = headers[4];
    expect(firstTotal).toBe('Total');
    expect(secondTotal).toBe('Total ');

    const row = applyFormBGJGujaratEmployeeToRow(
      {},
      { FirstName: 'Test', LastName: 'Emp' },
      rawHeaders,
      {
        payrollRow: {
          hra: 10935,
          gross_pay: 63622,
          net_pay: 60625,
          pf: 2797,
        },
      }
    );

    expect(row.HRA).toBe('10935');
    expect(row.PF).toBe('2797');
    expect(row[firstTotal]).toBe('63622');
    expect(row[secondTotal]).toBe('2997');
    expect(row['Net Payment']).toBe('60625');
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
