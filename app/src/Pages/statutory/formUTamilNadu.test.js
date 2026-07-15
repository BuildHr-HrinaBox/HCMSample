import {
  FORM_U_TAMILNADU_HEADERS,
  buildFormUTamilNaduCombinedBankDetails,
  ensureFormUTamilNaduMonthYearHeaderFields,
  formUTamilNaduHeaderAliasBucket,
  headersIndicateFormUTamilNaduTemplateLayout,
  isFormUTamilNaduCombinedBankHeader,
  isFormUTamilNaduEmployeeRegisterContext,
  looksLikeFormUFilename,
  normalizeFormUTamilNaduEmployeeRegisterHeaders,
  resolveFormUTamilNaduExportCellValue,
  resolveFormUTamilNaduPresentAddress,
  resolveFormUTamilNaduPermanentAddress,
  resolveFormUTamilNaduBankAddress
} from './formUTamilNadu';

describe('Form U Tamil Nadu helpers', () => {
  test('detects Form_U_-_TamilNadu.xlsx filename (underscore after U)', () => {
    expect(looksLikeFormUFilename('Form_U_-_TamilNadu.xlsx')).toBe(true);
    expect(
      isFormUTamilNaduEmployeeRegisterContext({
        fileName: 'Form_U_-_TamilNadu.xlsx',
        item: { state: 'Tamil Nadu', formName: 'Form U' }
      })
    ).toBe(true);
  });

  test('canonical headers include Date of Joining and Designation before Present Address', () => {
    const dob = FORM_U_TAMILNADU_HEADERS.indexOf('Date of Birth');
    const doj = FORM_U_TAMILNADU_HEADERS.indexOf('Date of Joining');
    const designation = FORM_U_TAMILNADU_HEADERS.indexOf('Designation');
    const present = FORM_U_TAMILNADU_HEADERS.indexOf('Present Address');
    expect(doj).toBe(dob + 1);
    expect(designation).toBe(doj + 1);
    expect(present).toBe(designation + 1);
    expect(formUTamilNaduHeaderAliasBucket('Date of Joining')).toBe('doj');
  });

  test('export resolves Date of Joining from joined/joining aliases', () => {
    expect(
      resolveFormUTamilNaduExportCellValue(
        { 'Date on which joined': '01-Dec-2025' },
        'Date of Joining'
      )
    ).toBe('01-Dec-2025');
    expect(
      resolveFormUTamilNaduExportCellValue(
        { 'Date of Joining': '15-Mar-2024' },
        'Date of Joining'
      )
    ).toBe('15-Mar-2024');
  });

  test('distinguishes Permanent Address vs Date made permanent aliases', () => {
    expect(formUTamilNaduHeaderAliasBucket('Permanent address')).toBe('permanent_address');
    expect(formUTamilNaduHeaderAliasBucket('Date on which made permanent')).toBe('made_permanent');
    expect(formUTamilNaduHeaderAliasBucket("Employee's Provident Fund No.")).toBe('epf');
    expect(formUTamilNaduHeaderAliasBucket("Employee's State Insurance Corporation No.")).toBe(
      'esic'
    );
  });

  test('export cell resolver maps by alias, not shifted index', () => {
    const row = {
      'Present Address': '12 Main St',
      'Permanent Address': '12 Main St',
      'Worker Identity No.': 'VE0147',
      Aadhaar: '395259631300',
      'Date on which joined': '01-Dec-2025',
      'Bank A/c Number, Name of Bank, Branch (Indian Financial System Code) (IFSC Code)':
        '123456, HDFC (HDFC0001)',
      'Email ID': 'avudaiappan.muthukrishnan@vayonaenergy.com',
      'Mobile Number': '9876543210'
    };
    expect(
      resolveFormUTamilNaduExportCellValue(row, 'Permanent address')
    ).toBe('12 Main St');
    expect(
      resolveFormUTamilNaduExportCellValue(row, 'Worker Identity No.')
    ).toBe('VE0147');
    expect(resolveFormUTamilNaduExportCellValue(row, 'Aadhaar No.')).toBe('395259631300');
    expect(
      resolveFormUTamilNaduExportCellValue(
        row,
        'Bank A/c Number, Branch (Indian Financial System Code) (IFSC Code)'
      )
    ).toContain('123456');
    expect(resolveFormUTamilNaduExportCellValue(row, 'Email ID')).toBe(
      'avudaiappan.muthukrishnan@vayonaenergy.com'
    );
    expect(resolveFormUTamilNaduExportCellValue(row, 'E Mail Id')).toBe(
      'avudaiappan.muthukrishnan@vayonaenergy.com'
    );
    // Must NOT put worker id into permanent address via index fallback
    expect(resolveFormUTamilNaduExportCellValue(row, 'Date on which made permanent')).toBe('');
    expect(
      resolveFormUTamilNaduExportCellValue(row, "Employee's State Insurance Corporation No.")
    ).toBe('');
    // Index misalignment (extra Bank Address in modal) must not blank Email
    const modalHeaders = [
      ...FORM_U_TAMILNADU_HEADERS.slice(0, 15),
      'Bank Address',
      ...FORM_U_TAMILNADU_HEADERS.slice(15)
    ];
    expect(
      resolveFormUTamilNaduExportCellValue(row, 'Email ID', modalHeaders, 18)
    ).toBe('avudaiappan.muthukrishnan@vayonaenergy.com');
  });

  test('Email ID alias is not confused with employee id', () => {
    expect(formUTamilNaduHeaderAliasBucket('Email ID')).toBe('email');
    expect(formUTamilNaduHeaderAliasBucket('E Mail Id')).toBe('email');
    expect(formUTamilNaduHeaderAliasBucket('Worker Identity No.')).toBe('workerid');
  });

  test('keeps template header order for TN layout', () => {
    const templateHeaders = [
      'S.No',
      'Name of the employee',
      'Worker Identity No.',
      'Present Address',
      'Permanent address',
      "Employee's Provident Fund No.",
      "Employee's State Insurance Corporation No.",
      'Aadhaar No.',
      'Date on which made permanent',
      'Bank A/c Number, Name of Bank, Branch (Indian Financial System Code) (IFSC Code)'
    ];
    expect(headersIndicateFormUTamilNaduTemplateLayout(templateHeaders)).toBe(true);
    const normalized = normalizeFormUTamilNaduEmployeeRegisterHeaders(templateHeaders, {
      fileName: 'Form_U_-_TamilNadu.xlsx'
    });
    expect(normalized.slice(0, 6)).toEqual(templateHeaders.slice(0, 6));
  });

  test('maps present and permanent address from Permanent_Address', () => {
    const emp = {
      Present_Address: 'Should not use',
      Permanent_Address: '12 Main Street, Chennai'
    };
    expect(resolveFormUTamilNaduPresentAddress(emp)).toBe('12 Main Street, Chennai');
    expect(resolveFormUTamilNaduPermanentAddress(emp)).toBe('12 Main Street, Chennai');
  });

  test('builds combined bank details and resolves Bank_Address', () => {
    expect(
      buildFormUTamilNaduCombinedBankDetails({
        Bank_Account_Number: '1234567890',
        Bank_Name: 'HDFC Bank',
        IFSC_Code: 'HDFC0001234'
      })
    ).toBe('1234567890, HDFC Bank (HDFC0001234)');
    expect(isFormUTamilNaduCombinedBankHeader('Bank A/c Number, Name of Bank, Branch (IFSC Code)')).toBe(
      true
    );
    expect(resolveFormUTamilNaduBankAddress({ Bank_Address: 'Anna Nagar Branch' })).toBe(
      'Anna Nagar Branch'
    );
  });

  test('injects Month and Year header fields', () => {
    const header = ensureFormUTamilNaduMonthYearHeaderFields({ fields: [] });
    expect(header.fields).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ key: 'form_x_month', label: 'Month:' }),
        expect.objectContaining({ key: 'form_x_year', label: 'Year:' })
      ])
    );
  });
});
