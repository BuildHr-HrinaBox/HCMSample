import {
  applyFormPGJGujaratPayrollToRow,
  applyFormPGJGujaratStaticDefaults,
  computeFormPGJGujaratOtherDeductions,
  computeFormPGJGujaratTotalDeduction,
  FORM_PGJ_WORKING_HOURS_FROM,
  FORM_PGJ_WORKING_HOURS_TO,
  formatFormPGJPayrollPayDate,
  resolveFormPGJGujaratPayrollFields,
  resolveFormPGJPayDateFromCache,
} from './formPGJGujarat';

describe('Form P Gujarat Sample Payroll deductions', () => {
  test('Total Deduction Rs. = gross_pay − net_pay', () => {
    expect(computeFormPGJGujaratTotalDeduction(236887, 177014)).toBe(59873);
    expect(computeFormPGJGujaratTotalDeduction('95316', '89026')).toBe(6290);
    expect(computeFormPGJGujaratTotalDeduction('', 100)).toBe('');
  });

  test('Other Deductions Rs. = Sample Payroll Total Deduction − Professional Tax', () => {
    expect(computeFormPGJGujaratOtherDeductions(1600, 200)).toBe(1400);
    expect(computeFormPGJGujaratOtherDeductions(700, '')).toBe(700);
    expect(computeFormPGJGujaratOtherDeductions('', 200)).toBe('');
  });

  test('resolveFormPGJGujaratPayrollFields reads PF / PT / Income Tax from Sample Payroll', () => {
    const fields = resolveFormPGJGujaratPayrollFields({
      employee_name: 'Test Emp',
      pf: 9211,
      professional_tax: 200,
      income_tax: 30640,
      gross_pay: 236887,
      net_pay: 177014,
      total_deductions: 1600,
    });
    expect(fields.pf).toBe(9211);
    expect(fields.professionalTax).toBe(200);
    expect(fields.incomeTax).toBe(30640);
    expect(fields.totalDeduction).toBe(59873);
    expect(fields.otherDeductions).toBe(1400);
  });

  test('resolveFormPGJGujaratPayrollFields prefers flatten epf_contribution / income_tax', () => {
    const fields = resolveFormPGJGujaratPayrollFields({
      employee_name: 'Test Emp',
      epf_contribution: 9211,
      income_tax: 30640,
      professional_tax: 200,
      Gross: 22832,
      'Net Pay': 20000,
      'Total Deduction': 700,
    });
    expect(fields.pf).toBe(9211);
    expect(fields.incomeTax).toBe(30640);
    expect(fields.grossPay).toBe(22832);
    expect(fields.totalDeduction).toBe(2832);
    expect(fields.otherDeductions).toBe(500);
  });

  test('applyFormPGJGujaratPayrollToRow fills PF, PT, Income Tax, Other / Total Deduction', () => {
    const headers = [
      'Provident Fund Rs.',
      'Professional Tax Rs.',
      'Income Tax Rs.',
      'Other Deductions Rs.',
      'Total Deduction Rs.',
    ];
    const row = {};
    applyFormPGJGujaratPayrollToRow(
      row,
      {
        pf: 1800,
        professional_tax: 200,
        income_tax: 1200,
        gross_pay: 50000,
        net_pay: 45000,
        total_deductions: 1600,
      },
      headers
    );
    expect(row['Provident Fund Rs.']).toBe('1800');
    expect(row['Professional Tax Rs.']).toBe('200');
    expect(row['Income Tax Rs.']).toBe('1200');
    expect(row['Other Deductions Rs.']).toBe('1400');
    expect(row['Total Deduction Rs.']).toBe('5000');
  });

  test('PF column does not receive Professional Tax', () => {
    const row = {};
    applyFormPGJGujaratPayrollToRow(
      row,
      { pf: 1800, professional_tax: 200 },
      ['Provident Fund Rs.']
    );
    expect(row['Provident Fund Rs.']).toBe('1800');
    expect(row['Provident Fund Rs.']).not.toBe('200');
  });

  test('Working hours From / To default to 9 AM / 5 PM', () => {
    expect(FORM_PGJ_WORKING_HOURS_FROM).toBe('9 AM');
    expect(FORM_PGJ_WORKING_HOURS_TO).toBe('5 PM');

    const row = {};
    applyFormPGJGujaratPayrollToRow(row, { pf: 100 }, [
      'Working hours_From',
      'Working hours_To',
    ]);
    expect(row['Working hours_From']).toBe('9 AM');
    expect(row['Working hours_To']).toBe('5 PM');

    const staticRow = {};
    applyFormPGJGujaratStaticDefaults([staticRow], [
      'Working hours_From',
      'Working hours_To',
    ]);
    expect(staticRow['Working hours_From']).toBe('9 AM');
    expect(staticRow['Working hours_To']).toBe('5 PM');
  });

  test('Date of Payment uses month end when provided', () => {
    const fields = resolveFormPGJGujaratPayrollFields(
      { employee_name: 'A', pf: 100, pay_date: '2026-04-15' },
      { monthEndDate: '30-04-2026', payDate: '2026-04-15' }
    );
    expect(fields.paymentDate).toBe('30-04-2026');

    const row = {};
    applyFormPGJGujaratPayrollToRow(
      row,
      { pf: 100, pay_date: '2026-04-15' },
      ['Date of Payment'],
      { monthEndDate: '30-04-2026' }
    );
    expect(row['Date of Payment']).toBe('30-04-2026');
  });

  test('Date of Payment fills from Sample Payroll Pay date', () => {
    expect(formatFormPGJPayrollPayDate('2026-04-30')).toBe('30-04-2026');

    const fields = resolveFormPGJGujaratPayrollFields(
      { employee_name: 'A', pf: 100 },
      { payDate: '2026-04-30' }
    );
    expect(fields.paymentDate).toBe('30-04-2026');

    const row = {};
    applyFormPGJGujaratPayrollToRow(
      row,
      { pf: 100, pay_date: '2026-04-30' },
      ['Date of Payment']
    );
    expect(row['Date of Payment']).toBe('30-04-2026');

    const metaOnly = {};
    applyFormPGJGujaratPayrollToRow(metaOnly, null, ['Date of Payment'], {
      payDate: '2026-04-30',
    });
    expect(metaOnly['Date of Payment']).toBe('30-04-2026');
  });

  test('empty row pay_date still uses helpers Pay date', () => {
    const fields = resolveFormPGJGujaratPayrollFields(
      { employee_name: 'A', pf: 100, pay_date: '', payDate: '' },
      { payDate: '2026-04-30' }
    );
    expect(fields.paymentDate).toBe('30-04-2026');

    const row = { 'Date of Payment': '' };
    applyFormPGJGujaratPayrollToRow(
      row,
      { pf: 100, net_pay: 90000, gross_pay: 95000, pay_date: '' },
      ['Date of Payment', 'Net Payable Rs.'],
      { payDate: '2026-04-15' }
    );
    expect(row['Date of Payment']).toBe('15-04-2026');
  });

  test('formPGJTemplateBucket maps DEDUCTION child labels', () => {
    // Exercised via apply + export value helpers — headers under DEDUCTION parent.
    const row = {};
    applyFormPGJGujaratPayrollToRow(
      row,
      {
        pf: 1800,
        professional_tax: 200,
        income_tax: 1200,
        gross_pay: 50000,
        net_pay: 45000,
        total_deductions: 1600,
      },
      [
        'Provident Fund Rs.',
        'Professional Tax Rs.',
        'Income Tax Rs.',
        'Other Deductions Rs.',
        'Total Deduction Rs.',
        'Date of Payment',
      ],
      { monthEndDate: '30-04-2026' }
    );
    expect(row['Provident Fund Rs.']).toBe('1800');
    expect(row['Professional Tax Rs.']).toBe('200');
    expect(row['Income Tax Rs.']).toBe('1200');
    expect(row['Other Deductions Rs.']).toBe('1400');
    expect(row['Total Deduction Rs.']).toBe('5000');
    expect(row['Date of Payment']).toBe('30-04-2026');
  });

  test('resolveFormPGJPayDateFromCache reads top-level payDate', () => {
    expect(
      resolveFormPGJPayDateFromCache({
        payDate: '2026-04-30',
        meta: {},
      })
    ).toBe('2026-04-30');
    expect(
      resolveFormPGJPayDateFromCache({
        payDate: '',
        meta: { pay_date: '2026-05-01' },
      })
    ).toBe('2026-05-01');
  });
});
