import {
  applyFormCGJGujaratEmployeeToRow,
  applyFormCGJGujaratNilToMappedRows,
  FORM_C_GJ_DEDUCTION_NIL_TEXT,
  isFormCGJDeductionNilHeader,
  isFormCGJSrNumberRegisterHeader,
  readFormCGJEmployeeId,
} from './formCGJGujarat';

const FORM_C_GJ_HEADERS = [
  'Sr. Number in Employee / Workman / Worker Register',
  'Recovery type (Damage/Loss/fine/Advance/Loans/Absence)',
  'Particulars',
  'Date Of Damage/ Loss / Absence',
  'Amount',
  'Whether Show cause issued',
  'Explanation heard in presence of',
  'Number of instalments',
  'First Month/ Year',
  'Last Month / Year',
  'Date of Complete Recovery',
];

describe('Form C Gujarat Sr. Number ← EmployeeID and NIL columns', () => {
  test('detects Sr. Number register header', () => {
    expect(
      isFormCGJSrNumberRegisterHeader('Sr. Number in Employee / Workman / Worker Register')
    ).toBe(true);
    expect(isFormCGJSrNumberRegisterHeader('Particulars')).toBe(false);
    expect(isFormCGJSrNumberRegisterHeader('Name')).toBe(false);
  });

  test('detects deduction NIL headers from Form C GJ template', () => {
    expect(isFormCGJDeductionNilHeader('Recovery type (Damage/Loss/fine/Advance/Loans/Absence)')).toBe(
      true
    );
    expect(isFormCGJDeductionNilHeader('Particulars')).toBe(true);
    expect(isFormCGJDeductionNilHeader('Date Of Damage/ Loss / Absence')).toBe(true);
    expect(isFormCGJDeductionNilHeader('Amount')).toBe(true);
    expect(isFormCGJDeductionNilHeader('Whether Show cause issued')).toBe(true);
    expect(isFormCGJDeductionNilHeader('Explanation heard in presence of')).toBe(true);
    expect(isFormCGJDeductionNilHeader('Number of instalments')).toBe(true);
    expect(isFormCGJDeductionNilHeader('First Month/ Year')).toBe(true);
    expect(isFormCGJDeductionNilHeader('Last Month / Year')).toBe(true);
    expect(isFormCGJDeductionNilHeader('Date of Complete Recovery')).toBe(true);
    expect(
      isFormCGJDeductionNilHeader('Sr. Number in Employee / Workman / Worker Register')
    ).toBe(false);
  });

  test('reads EmployeeID from People employee', () => {
    expect(readFormCGJEmployeeId({ EmployeeID: 'VE0147' })).toBe('VE0147');
    expect(readFormCGJEmployeeId({ Employee: { 'Employee ID': 'VE0099' } })).toBe('VE0099');
  });

  test('applyFormCGJGujaratEmployeeToRow fills EmployeeID and NIL defaults', () => {
    const row = applyFormCGJGujaratEmployeeToRow(
      {},
      { EmployeeID: 'VE0147', FirstName: 'Vasan', LastName: 'Jagad' },
      FORM_C_GJ_HEADERS
    );
    expect(row['Sr. Number in Employee / Workman / Worker Register']).toBe('VE0147');
    expect(row['Recovery type (Damage/Loss/fine/Advance/Loans/Absence)']).toBe(
      FORM_C_GJ_DEDUCTION_NIL_TEXT
    );
    expect(row.Particulars).toBe('NIL');
    expect(row['Date Of Damage/ Loss / Absence']).toBe('NIL');
    expect(row.Amount).toBe('NIL');
    expect(row['Whether Show cause issued']).toBe('NIL');
    expect(row['Explanation heard in presence of']).toBe('NIL');
    expect(row['Number of instalments']).toBe('NIL');
    expect(row['First Month/ Year']).toBe('NIL');
    expect(row['Last Month / Year']).toBe('NIL');
    expect(row['Date of Complete Recovery']).toBe('NIL');
  });

  test('applyFormCGJGujaratNilToMappedRows fills blank deduction columns with NIL', () => {
    const rows = applyFormCGJGujaratNilToMappedRows(
      [
        {
          'Sr. Number in Employee / Workman / Worker Register': 'VE0147',
          Particulars: '',
          Amount: '',
        },
      ],
      FORM_C_GJ_HEADERS
    );
    expect(rows[0]['Sr. Number in Employee / Workman / Worker Register']).toBe('VE0147');
    expect(rows[0].Particulars).toBe('NIL');
    expect(rows[0].Amount).toBe('NIL');
    expect(rows[0]['Date of Complete Recovery']).toBe('NIL');
  });
});
