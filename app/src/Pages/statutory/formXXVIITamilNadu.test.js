import {
  FORM_XXVII_TN_GROUP_OTHER_ALLOWANCES,
  FORM_XXVII_TN_GROUP_OTHER_DEDUCTIONS,
  FORM_XXVII_TN_OTHER_ALLOWANCE_LEAVES,
  FORM_XXVII_TN_OTHER_DEDUCTION_LEAVES,
  FORM_XXVII_TN_WAGE_PERIOD_DEFAULT,
  FORM_XXVII_TN_DAILY_RATED_TYPE_DEFAULT,
  FORM_XXVII_TN_OVERTIME_RATE_DEFAULT,
  applyFormXXVIITamilNaduPayrollToRow,
  buildFormXXVIITamilNaduColumnGroupLabels,
  buildFormXXVIITamilNaduWagePeriodLine,
  computeFormXXVIITamilNaduOtherAllowancesEcca,
  computeFormXXVIITamilNaduOtherDeductions,
  computeFormXXVIITamilNaduTotalDeductions,
  formXXVIITamilNaduNeedsOtherAllowancesGroupThead,
  hasFormXXVIITamilNaduSamplePayrollNameParts,
  isFormXXVIITamilNaduColumnIndexHeader,
  isFormXXVIITamilNaduColumnIndexRow,
  isFormXXVIITamilNaduIdentityHeaderText,
  isFormXXVIITamilNaduVerticalHeaderText,
  isFormXXVIITamilNaduContext,
  isFormXXVIITamilNaduDailyRatedHeader,
  isFormXXVIITamilNaduDailyRatedTypeHeader,
  isFormXXVIITamilNaduDaysWorkedHeader,
  isFormXXVIITamilNaduOtherAllowancesEccaHeader,
  isFormXXVIITamilNaduOtherDeductionsHeader,
  isFormXXVIITamilNaduOvertimeRateHeader,
  isFormXXVIITamilNaduPtHeader,
  isFormXXVIITamilNaduSkipAutofillHeader,
  isFormXXVIITamilNaduWagePeriodColumnHeader,
  looksLikeFormXXVIITamilNaduWageHeaders,
  readFormXXVIITamilNaduPaidDays,
  resolveFormXXVIITamilNaduDailyRated,
  resolveFormXXVIITamilNaduDaysWorked,
  resolveFormXXVIITamilNaduHra,
  resolveFormXXVIITamilNaduOtherAllowancesEcca,
  resolveFormXXVIITamilNaduOtherDeductions,
  resolveFormXXVIITamilNaduOvertimeRate,
  resolveFormXXVIITamilNaduPeriodParts,
  resolveFormXXVIITamilNaduPt,
  resolveFormXXVIITamilNaduTableHeaders,
  resolveFormXXVIITamilNaduTotalDeductions,
  sanitizeFormXXVIITamilNaduColumnGroupLabels,
  splitFormXXVIITamilNaduVerticalHeaderLines,
} from './formXXVIITamilNadu';

describe('formXXVIITamilNadu column grouping', () => {
  const headers = [
    'S.No',
    'Name of Workman',
    "Father's / Husband's Name",
    'Designation',
    'Employee Number',
    'Days Worked',
    'Units of Work',
    'Daily Rate',
    'Wage Period',
    'OVERTIME RATE',
    'BASIC WAGE',
    'DEARNESS ALLOWANCE',
    'Column 13',
    'Column 14',
    'Column 15',
    'Column 16',
    'Column 17',
    'GROSS WAGES',
    'PROVIDEND FUND',
    'ESI',
    'Column 21',
    'Column 22',
    'FINES (IF ANY)',
    'OTHER DEDUCTIONS',
    'TOTAL DEDUCTIONS',
  ];

  test('detects Form XXVII Tamil Nadu wage headers', () => {
    expect(looksLikeFormXXVIITamilNaduWageHeaders(headers)).toBe(true);
    expect(
      isFormXXVIITamilNaduContext(
        { title: 'FORM XXVII' },
        { formName: 'Form XXVII', state: 'Tamil Nadu' },
        'Form_XXVII_-_TamilNadu.xlsx',
        headers
      )
    ).toBe(true);
  });

  test('resolves cols 13–17 to Excel leaf titles (WASH ALLOW, HRA, …)', () => {
    const resolved = resolveFormXXVIITamilNaduTableHeaders(headers);
    expect(resolved.slice(12, 17)).toEqual([...FORM_XXVII_TN_OTHER_ALLOWANCE_LEAVES]);
    expect(resolved[11]).toBe('DEARNESS ALLOWANCE');
    expect(resolved[17]).toBe('GROSS WAGES');
  });

  test('resolves cols 21–22 to PT and UNIFORM DEPOSITS', () => {
    const resolved = resolveFormXXVIITamilNaduTableHeaders(headers);
    expect(resolved.slice(20, 22)).toEqual([...FORM_XXVII_TN_OTHER_DEDUCTION_LEAVES]);
    expect(resolved[19]).toBe('ESI');
    expect(resolved[22]).toBe('FINES (IF ANY)');
  });

  test('groups allowance and OTHER deduction bands', () => {
    const resolved = resolveFormXXVIITamilNaduTableHeaders(headers);
    const groups = buildFormXXVIITamilNaduColumnGroupLabels(resolved);
    expect(groups[10]).toBe('BASIC WAGE');
    expect(groups[11]).toBe('DEARNESS ALLOWANCE');
    expect(groups[12]).toBe(FORM_XXVII_TN_GROUP_OTHER_ALLOWANCES);
    expect(groups[16]).toBe(FORM_XXVII_TN_GROUP_OTHER_ALLOWANCES);
    expect(groups[17]).toBe('GROSS WAGES');
    expect(groups[20]).toBe(FORM_XXVII_TN_GROUP_OTHER_DEDUCTIONS);
    expect(groups[21]).toBe(FORM_XXVII_TN_GROUP_OTHER_DEDUCTIONS);
    expect(groups[22]).toBe('FINES (IF ANY)');
    expect(formXXVIITamilNaduNeedsOtherAllowancesGroupThead(resolved)).toBe(true);
  });

  test('sanitize prefers official allowance / OTHER bands', () => {
    const resolved = resolveFormXXVIITamilNaduTableHeaders(headers);
    const incoming = resolved.map((_, i) => {
      if (i >= 12 && i <= 16) return 'Other';
      if (i >= 20 && i <= 21) return 'OTHER';
      return '';
    });
    const sanitized = sanitizeFormXXVIITamilNaduColumnGroupLabels(incoming, resolved);
    expect(sanitized[12]).toBe(FORM_XXVII_TN_GROUP_OTHER_ALLOWANCES);
    expect(sanitized[20]).toBe(FORM_XXVII_TN_GROUP_OTHER_DEDUCTIONS);
    expect(sanitized[10]).toBe('BASIC WAGE');
  });

  test('does not treat AP quarterly returns as TN wage register', () => {
    expect(
      isFormXXVIITamilNaduContext(
        { title: 'Form XXVII', subtitle: 'Quarterly Returns' },
        { formName: 'Form XXVII Quarterly Returns', state: 'Andhra Pradesh' },
        'Form_XXVII_-_Andhra_Pradesh.xlsx',
        ['Particulars', 'Month']
      )
    ).toBe(false);
  });

  test('official model: identity headings stay horizontal; wage leaves rotate vertical', () => {
    expect(isFormXXVIITamilNaduIdentityHeaderText('Sr. No.')).toBe(true);
    expect(isFormXXVIITamilNaduIdentityHeaderText('Name of the Workman')).toBe(true);
    expect(isFormXXVIITamilNaduIdentityHeaderText('Sex')).toBe(true);
    expect(isFormXXVIITamilNaduIdentityHeaderText('Designation (Nature of work)')).toBe(true);
    expect(isFormXXVIITamilNaduVerticalHeaderText('Sr. No.')).toBe(false);
    expect(isFormXXVIITamilNaduVerticalHeaderText('WAGES EARNED')).toBe(false);
    expect(isFormXXVIITamilNaduVerticalHeaderText('DEDUCTIONS')).toBe(false);
    expect(
      isFormXXVIITamilNaduVerticalHeaderText('OTHER ALLOWANCES/CASH PAYMENT NATURE TO BE SPECIFIED')
    ).toBe(false);
    expect(isFormXXVIITamilNaduVerticalHeaderText('DAILY RATED/PIECE RATE/MONTHLY RATED')).toBe(
      true
    );
    expect(isFormXXVIITamilNaduVerticalHeaderText('BASIC WAGE')).toBe(true);
    expect(isFormXXVIITamilNaduVerticalHeaderText('HRA')).toBe(true);
    expect(isFormXXVIITamilNaduVerticalHeaderText('GROSS WAGES')).toBe(true);
    expect(isFormXXVIITamilNaduVerticalHeaderText('NET WAGES')).toBe(true);
    expect(isFormXXVIITamilNaduColumnIndexHeader('18')).toBe(true);
    expect(isFormXXVIITamilNaduColumnIndexRow(['1', '2', '3', '4', '5', '6', '7', '8'])).toBe(true);
    expect(splitFormXXVIITamilNaduVerticalHeaderLines('BASIC WAGE')).toEqual(['BASIC', 'WAGE']);
    expect(splitFormXXVIITamilNaduVerticalHeaderLines('HRA')).toEqual(['HRA']);
    expect(
      splitFormXXVIITamilNaduVerticalHeaderLines(
        'SIGNATURE / THUMB IMPRESSION CHEQUE No. & DATE / BANK'
      )
    ).toEqual(['SIGNATURE / THUMB', 'IMPRESSION', 'CHEQUE No. & DATE', '/ BANK']);
  });

  test('skips autofill for cash-in-lieu and unpaid accumulated columns', () => {
    expect(
      isFormXXVIITamilNaduSkipAutofillHeader('WAGES INCLUDING CASH IN LIEU OF KINDS')
    ).toBe(true);
    expect(isFormXXVIITamilNaduSkipAutofillHeader('WAGES INCLUDING CASH IN LIEU OF KIND')).toBe(
      true
    );
    expect(isFormXXVIITamilNaduSkipAutofillHeader('TOTAL UNPAID AMOUNT ACCUMULATED')).toBe(true);
    expect(isFormXXVIITamilNaduSkipAutofillHeader('TOTAL UMPAID AMOUNT ACCUMULATED')).toBe(true);
    expect(isFormXXVIITamilNaduSkipAutofillHeader('HRA')).toBe(false);
    expect(isFormXXVIITamilNaduSkipAutofillHeader('GROSS WAGES')).toBe(false);
  });

  test('builds Wage Period : <Month> from payroll month iso', () => {
    expect(buildFormXXVIITamilNaduWagePeriodLine('2026-04')).toBe('Wage Period : April');
    expect(buildFormXXVIITamilNaduWagePeriodLine('2026-07')).toBe('Wage Period : July');
    expect(resolveFormXXVIITamilNaduPeriodParts('2026-04')).toEqual(
      expect.objectContaining({ monthName: 'April', year: '2026' })
    );
  });
});

describe('formXXVIITamilNadu Sample Payroll autofill', () => {
  const payrollRow = {
    gross_pay: 104065,
    basic: 26781,
    hra: 12000,
    professional_tax: 200,
    total_deductions: 5000,
    epf_contribution: 3000,
    esi: 500,
    net_pay: 99065,
  };

  test('detects daily rated / wage period / overtime rate / ECCA / PT / other deductions headers', () => {
    expect(isFormXXVIITamilNaduDailyRatedHeader('DAILY RATED WAGES/PIECE RATES')).toBe(true);
    expect(isFormXXVIITamilNaduDailyRatedTypeHeader('DAILY RATED/ PIECE RATED/MONTHLY RATED')).toBe(
      true
    );
    expect(isFormXXVIITamilNaduDailyRatedHeader('DAILY RATED/ PIECE RATED/MONTHLY RATED')).toBe(
      false
    );
    expect(isFormXXVIITamilNaduWagePeriodColumnHeader('WAGE PERIOD- WEEKLY/FN/MONTHLY')).toBe(
      true
    );
    expect(isFormXXVIITamilNaduOvertimeRateHeader('OVERTIME RATE')).toBe(true);
    expect(isFormXXVIITamilNaduOtherAllowancesEccaHeader('OTHER ALLOWANCES, ECCA')).toBe(true);
    expect(isFormXXVIITamilNaduPtHeader('PT')).toBe(true);
    expect(isFormXXVIITamilNaduOtherDeductionsHeader('OTHER DEDUCTIONS')).toBe(true);
  });

  test('DAILY RATED type column defaults to Monthly (never fetch payroll)', () => {
    expect(resolveFormXXVIITamilNaduDailyRated(payrollRow)).toBe('104065');
    const typeHdr = 'DAILY RATED/ PIECE RATED/MONTHLY RATED';
    const row = { [typeHdr]: '104065' };
    applyFormXXVIITamilNaduPayrollToRow(row, payrollRow, [typeHdr], { overwrite: true });
    expect(row[typeHdr]).toBe(FORM_XXVII_TN_DAILY_RATED_TYPE_DEFAULT);
    applyFormXXVIITamilNaduPayrollToRow(row, null, [typeHdr], { overwrite: true });
    expect(row[typeHdr]).toBe(FORM_XXVII_TN_DAILY_RATED_TYPE_DEFAULT);
  });

  test('OVERTIME RATE is always NIL (never fetch payroll)', () => {
    expect(resolveFormXXVIITamilNaduOvertimeRate(payrollRow)).toBe(FORM_XXVII_TN_OVERTIME_RATE_DEFAULT);
    expect(resolveFormXXVIITamilNaduOvertimeRate(null)).toBe(FORM_XXVII_TN_OVERTIME_RATE_DEFAULT);
  });

  test('HRA ← Sample Payroll hra', () => {
    expect(resolveFormXXVIITamilNaduHra(payrollRow)).toBe('12000');
  });

  test('OTHER ALLOWANCES, ECCA ← gross − basic − hra', () => {
    expect(computeFormXXVIITamilNaduOtherAllowancesEcca(104065, 26781, 12000)).toBe(65284);
    expect(resolveFormXXVIITamilNaduOtherAllowancesEcca(payrollRow)).toBe(65284);
  });

  test('PT ← Professional Tax', () => {
    expect(resolveFormXXVIITamilNaduPt(payrollRow)).toBe('200');
  });

  test('OTHER DEDUCTIONS ← TOTAL − PT − ESI − PF', () => {
    expect(computeFormXXVIITamilNaduOtherDeductions(5000, 200, 500, 3000)).toBe(1300);
    // TOTAL = gross − net = 104065 − 99065 = 5000; OTHER = 5000 − 200 − 500 − 3000 = 1300
    expect(resolveFormXXVIITamilNaduOtherDeductions(payrollRow)).toBe(1300);
  });

  test('TOTAL DEDUCTIONS ← gross_pay − net_pay', () => {
    expect(computeFormXXVIITamilNaduTotalDeductions(104065, 99065)).toBe(5000);
    expect(resolveFormXXVIITamilNaduTotalDeductions(payrollRow)).toBe(5000);
  });

  test('TOTAL / UNITS days worked ← paid_days only when firstname + lastname present', () => {
    const totalHdr = 'TOTAL NUMBER OF DAYS WORKED DURING THE WEEK/FN/MONTH';
    const unitsHdr = 'UNITS OF WORK DONE/NUMBER OF DAYS WORKED';
    expect(isFormXXVIITamilNaduDaysWorkedHeader(totalHdr)).toBe(true);
    expect(isFormXXVIITamilNaduDaysWorkedHeader(unitsHdr)).toBe(true);

    expect(readFormXXVIITamilNaduPaidDays({ paid_days: 26 })).toBe('26');
    expect(hasFormXXVIITamilNaduSamplePayrollNameParts({ first_name: 'Ram', last_name: 'Kumar' })).toBe(
      true
    );
    expect(hasFormXXVIITamilNaduSamplePayrollNameParts({ first_name: 'Ram' })).toBe(false);
    expect(
      resolveFormXXVIITamilNaduDaysWorked({
        paid_days: 26,
        first_name: 'Ram',
        last_name: 'Kumar',
      })
    ).toBe('26');
    // Missing paid_days → blank (do not invent calendar days)
    expect(
      resolveFormXXVIITamilNaduDaysWorked({
        first_name: 'Ram',
        last_name: 'Kumar',
        gross_pay: 1000,
      })
    ).toBe('');
    // Missing last name → blank even if paid_days exists
    expect(
      resolveFormXXVIITamilNaduDaysWorked({
        paid_days: 31,
        first_name: 'Ram',
      })
    ).toBe('');

    const hdrs = [totalHdr, unitsHdr];
    const row = { [totalHdr]: '31', [unitsHdr]: '31' };
    applyFormXXVIITamilNaduPayrollToRow(
      row,
      { first_name: 'Ram', last_name: 'Kumar' },
      hdrs,
      { overwrite: true }
    );
    expect(row[totalHdr]).toBe('');
    expect(row[unitsHdr]).toBe('');

    applyFormXXVIITamilNaduPayrollToRow(
      row,
      { paid_days: 26, firstName: 'Ram', lastName: 'Kumar' },
      hdrs,
      { overwrite: true }
    );
    expect(row[totalHdr]).toBe('26');
    expect(row[unitsHdr]).toBe('26');

    applyFormXXVIITamilNaduPayrollToRow(
      row,
      {
        employee_name: 'Pal Pandian V',
        paid_days: 15,
      },
      hdrs,
      { overwrite: true }
    );
    expect(row[totalHdr]).toBe('15');
    expect(row[unitsHdr]).toBe('15');

    expect(
      resolveFormXXVIITamilNaduDaysWorked({
        employee_name: 'Vinu Monikandan Muruganatha',
        paid_days: 31,
      })
    ).toBe('31');
  });

  test('applyFormXXVIITamilNaduPayrollToRow fills mapped columns', () => {
    const hdrs = [
      'DAILY RATED WAGES/PIECE RATES',
      'WAGE PERIOD- WEEKLY/FN/MONTHLY',
      'OVERTIME RATE',
      'HRA',
      'OTHER ALLOWANCES, ECCA',
      'PT',
      'OTHER DEDUCTIONS',
      'TOTAL DEDUCTIONS',
      'WAGES INCLUDING CASH IN LIEU OF KINDS',
    ];
    const row = Object.fromEntries(hdrs.map((h) => [h, '']));
    applyFormXXVIITamilNaduPayrollToRow(row, payrollRow, hdrs, { overwrite: true });
    expect(row['DAILY RATED WAGES/PIECE RATES']).toBe('104065');
    expect(row['WAGE PERIOD- WEEKLY/FN/MONTHLY']).toBe(FORM_XXVII_TN_WAGE_PERIOD_DEFAULT);
    expect(row['OVERTIME RATE']).toBe(FORM_XXVII_TN_OVERTIME_RATE_DEFAULT);
    expect(row.HRA).toBe('12000');
    expect(row['OTHER ALLOWANCES, ECCA']).toBe(65284);
    expect(row.PT).toBe('200');
    expect(row['OTHER DEDUCTIONS']).toBe(1300);
    expect(row['TOTAL DEDUCTIONS']).toBe(5000);
    // Manual-only column stays blank.
    expect(row['WAGES INCLUDING CASH IN LIEU OF KINDS']).toBe('');
  });
});
