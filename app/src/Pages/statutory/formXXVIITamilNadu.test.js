import {
  FORM_XXVII_TN_GROUP_OTHER_ALLOWANCES,
  FORM_XXVII_TN_GROUP_OTHER_DEDUCTIONS,
  FORM_XXVII_TN_OTHER_ALLOWANCE_LEAVES,
  FORM_XXVII_TN_OTHER_DEDUCTION_LEAVES,
  buildFormXXVIITamilNaduColumnGroupLabels,
  buildFormXXVIITamilNaduWagePeriodLine,
  formXXVIITamilNaduNeedsOtherAllowancesGroupThead,
  isFormXXVIITamilNaduContext,
  isFormXXVIITamilNaduSkipAutofillHeader,
  looksLikeFormXXVIITamilNaduWageHeaders,
  resolveFormXXVIITamilNaduPeriodParts,
  resolveFormXXVIITamilNaduTableHeaders,
  sanitizeFormXXVIITamilNaduColumnGroupLabels,
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
