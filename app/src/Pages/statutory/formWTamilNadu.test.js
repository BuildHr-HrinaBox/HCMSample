import {
  FORM_W_TN_DEFAULT_PAYROLL,
  applyFormWTamilNaduDefaultPayrollToMap,
  applyFormWTamilNaduGenderCountsToHeaderFormData,
  applyFormWTamilNaduMonthYearToHeaderFormData,
  buildFormWTamilNaduWagePeriodLine,
  computeFormWTamilNaduOtherAllowances,
  computeFormWTamilNaduTotalDeductions,
  countFormWTamilNaduGenderBoxes,
  ensureFormWTamilNaduGenderHeaderFields,
  formatFormWTamilNaduMonthEndPaymentDate,
  isFormWTamilNaduRegisterContext,
  looksLikeFormWFilename,
  pickFormWTamilNaduGenderCount,
  resolveFormWTamilNaduDefaultPayroll,
  resolveFormWTamilNaduPeriodParts,
} from './formWTamilNadu';

describe('Form W Tamil Nadu payroll helpers', () => {
  test('detects Form_W_-_TamilNadu.xlsx filename (underscore after W)', () => {
    expect(looksLikeFormWFilename('Form_W_-_TamilNadu.xlsx')).toBe(true);
    expect(looksLikeFormWFilename('Form W')).toBe(true);
    expect(looksLikeFormWFilename('Form_U_-_TamilNadu.xlsx')).toBe(false);
    expect(
      isFormWTamilNaduRegisterContext({
        fileName: 'Form_W_-_TamilNadu.xlsx',
        item: { state: 'Tamil Nadu', formName: 'Form W' },
      })
    ).toBe(true);
  });

  test('resolves default Basic / HRA by employee id', () => {
    expect(resolveFormWTamilNaduDefaultPayroll({ EmployeeID: 'VE0147' })).toEqual({
      basic: '67729',
      hra: '32665',
    });
    expect(resolveFormWTamilNaduDefaultPayroll({ EmployeeID: 'VE0079' })).toEqual({
      basic: '52628',
      hra: '22476',
    });
    expect(resolveFormWTamilNaduDefaultPayroll({ EmployeeID: 'VE0054' })).toEqual({
      basic: '67064',
      hra: '32345',
    });
    expect(resolveFormWTamilNaduDefaultPayroll({ EmployeeID: 'VE0051' })).toEqual({
      basic: '48145',
      hra: '22647',
    });
    expect(resolveFormWTamilNaduDefaultPayroll({ EmployeeID: 'VE0042' })).toEqual({
      basic: '43688',
      hra: '20704',
    });
  });

  test('resolves truncated display names', () => {
    expect(
      resolveFormWTamilNaduDefaultPayroll({ FirstName: 'Avudaiappa', LastName: '' })
    ).toEqual({ basic: '67729', hra: '32665' });
    expect(
      resolveFormWTamilNaduDefaultPayroll({ FirstName: 'Vinu', LastName: 'Monik' })
    ).toEqual({ basic: '43688', hra: '20704' });
  });

  test('Other Allowances = gross_pay − basic − hra', () => {
    expect(computeFormWTamilNaduOtherAllowances(206820, 67729, 32665)).toBe(106426);
    expect(computeFormWTamilNaduOtherAllowances('194566', '52628', '22476')).toBe(119462);
  });

  test('Total Deductions = gross_pay − net_pay', () => {
    expect(computeFormWTamilNaduTotalDeductions(206820, 180000)).toBe(26820);
    expect(computeFormWTamilNaduTotalDeductions(100, 200)).toBe('');
  });

  test('apply map forces Basic/HRA defaults and computes other + deductions', () => {
    const map = applyFormWTamilNaduDefaultPayrollToMap(
      {
        basicWage: 1,
        houseRentAllowance: 2,
        grossPay: 206820,
        grossWages: 206820,
        netPay: 180000,
        netWages: 180000,
        otherAllowances: 999,
      },
      { EmployeeID: 'VE0147', FirstName: 'Avudaiappan' }
    );
    expect(map.basicWage).toBe('67729');
    expect(map.houseRentAllowance).toBe('32665');
    expect(map.otherAllowances).toBe(106426);
    expect(map.deductionsFromGrossNet).toBe(26820);
    expect(map.totalDeductions).toBe(26820);
  });

  test('Date of payment is month last date', () => {
    expect(formatFormWTamilNaduMonthEndPaymentDate('2026-04')).toBe('30-04-2026');
    expect(formatFormWTamilNaduMonthEndPaymentDate('2026-02')).toBe('28-02-2026');
  });

  test('counts Men / Women / young person boxes by gender', () => {
    const counts = countFormWTamilNaduGenderBoxes(
      [
        { EmployeeID: '1', Sex: 'Male', Date_of_birth: '1990-01-01' },
        { EmployeeID: '2', Gender: 'Female', Date_of_birth: '1992-05-05' },
        { EmployeeID: '3', Sex: 'Male', Date_of_birth: '2015-01-01' },
        { EmployeeID: '4', Sex: 'Female', Date_of_birth: '2016-06-01' },
      ],
      '2026-04'
    );
    expect(counts).toEqual({
      men: 1,
      women: 1,
      maleYoung: 1,
      femaleYoung: 1,
      total: 4,
    });
  });

  test('ensures Men/Women header fields and applies counts', () => {
    const fields = ensureFormWTamilNaduGenderHeaderFields([]);
    expect(fields.some((f) => /men/i.test(f.label) && !/young/i.test(f.label))).toBe(true);
    expect(fields.some((f) => /women/i.test(f.label) && !/young/i.test(f.label))).toBe(true);
    expect(fields.some((f) => /male young person/i.test(f.label))).toBe(true);
    expect(fields.some((f) => /female young person/i.test(f.label))).toBe(true);
    const data = applyFormWTamilNaduGenderCountsToHeaderFormData(
      {},
      {
        men: 5,
        women: 0,
        maleYoung: 0,
        femaleYoung: 0,
        total: 5,
      }
    );
    expect(data.form_w_tn_men).toBe('5');
    expect(data.form_w_tn_women).toBe('0');
    expect(data.form_w_tn_total_persons_employed).toBe('5');
  });

  test('builds full Wage Period year and Month/Year parts', () => {
    expect(buildFormWTamilNaduWagePeriodLine('2026-04')).toBe(
      'Wage Period from 1st April 2026 to 30th April 2026'
    );
    expect(resolveFormWTamilNaduPeriodParts('2026-04')).toEqual({
      iso: '2026-04',
      monthName: 'April',
      year: '2026',
      monthIdx: 3,
    });
    const hdr = applyFormWTamilNaduMonthYearToHeaderFormData({}, '2026-04');
    expect(hdr.form_x_month).toBe('April');
    expect(hdr.form_x_year).toBe('2026');
  });

  test('pickFormWTamilNaduGenderCount keeps explicit zero (does not fall back to company count)', () => {
    expect(pickFormWTamilNaduGenderCount('0', 44)).toBe(0);
    expect(pickFormWTamilNaduGenderCount(0, 44)).toBe(0);
    expect(pickFormWTamilNaduGenderCount('', 44)).toBe(44);
    expect(pickFormWTamilNaduGenderCount(undefined, 5)).toBe(5);
    expect(pickFormWTamilNaduGenderCount('5', 99)).toBe(5);
  });

  test('default payroll table covers the five listed employees', () => {
    expect(FORM_W_TN_DEFAULT_PAYROLL).toHaveLength(5);
  });
});
