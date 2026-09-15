import ExcelJS from 'exceljs';
import {
  FORM_U_TAMILNADU_HEADERS,
  buildFormUTamilNaduCombinedBankDetails,
  ensureFormUTamilNaduMonthYearHeaderFields,
  formUTamilNaduHeaderAliasBucket,
  headersIndicateFormUTamilNaduTemplateLayout,
  isFormUTamilNaduCombinedBankHeader,
  isFormUTamilNaduEmployeeRegisterContext,
  fillFormUTamilNaduEmployeeTable,
  isFormUStatutoryDownloadHint,
  looksLikeFormUFilename,
  normalizeFormUTamilNaduEmployeeRegisterHeaders,
  normalizeFormUTamilNaduMappedRowsForExcel,
  resolveFormUTamilNaduExportCellValue,
  resolveFormUTamilNaduPresentAddress,
  resolveFormUTamilNaduPermanentAddress,
  resolveFormUTamilNaduBankAddress,
  scoreFormUTamilNaduSheetForExport,
  writeFormUTamilNaduExcelVisibleCell
} from './formUTamilNadu';

describe('Form U Tamil Nadu helpers', () => {
  test('detects Form_U_-_TamilNadu.xlsx filename (underscore after U)', () => {
    expect(looksLikeFormUFilename('Form_U_-_TamilNadu.xlsx')).toBe(true);
    expect(isFormUStatutoryDownloadHint('Form U - TamilNadu (3).xlsx')).toBe(true);
    expect(isFormUStatutoryDownloadHint('Employee Register', 'Form U - TamilNadu.xlsx')).toBe(true);
    expect(isFormUStatutoryDownloadHint('Form XIV Tamil Nadu')).toBe(false);
    expect(isFormUStatutoryDownloadHint('Form XVI')).toBe(false);
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

  test('treats Employee Identification + present/permanent as TN template layout', () => {
    expect(
      headersIndicateFormUTamilNaduTemplateLayout([
        'S.No',
        'Name of the employee',
        'Employee Identification No.',
        'Gender',
        'Date of Birth',
        'Date of Joining',
        'Designation',
        'Present Address',
        'Permanent address'
      ])
    ).toBe(true);
  });

  test('maps array / numeric-key autofill rows onto template labels for Excel', () => {
    const headers = [
      'S.No',
      'Name of the employee',
      'Employee Identification No.',
      'Gender'
    ];
    const fromArray = normalizeFormUTamilNaduMappedRowsForExcel(
      [[1, 'Babu', 'VE1111', 'Male']],
      headers
    );
    expect(fromArray).toHaveLength(1);
    expect(fromArray[0]['Name of the employee']).toBe('Babu');
    expect(
      resolveFormUTamilNaduExportCellValue(fromArray[0], 'Name of the employee', headers, 1)
    ).toBe('Babu');
    expect(
      resolveFormUTamilNaduExportCellValue([1, 'Babu', 'VE1111', 'Male'], 'Worker Identity No.', headers, 2)
    ).toBe('VE1111');

    const fromNumericKeys = normalizeFormUTamilNaduMappedRowsForExcel(
      [{ 0: 1, 1: 'P M', 2: 'VE1114', 3: 'Male' }],
      headers
    );
    expect(fromNumericKeys[0]['Name of the employee']).toBe('P M');
    expect(
      normalizeFormUTamilNaduMappedRowsForExcel(
        [
          {
            'S.No': '1',
            'Name of the employee': '2',
            'Employee Identification No.': '3',
            Gender: '4'
          }
        ],
        headers
      )
    ).toEqual([]);
  });

  test('prefers visible FORM U sheet over a hidden high-scoring helper sheet', () => {
    const hiddenHelper = scoreFormUTamilNaduSheetForExport({
      sheetName: 'Sheet3',
      sheetState: 'hidden',
      tableAnchorScore: 80,
      hasEmployeeTable: true
    });
    const visibleFormU = scoreFormUTamilNaduSheetForExport({
      sheetName: 'FORM U',
      sheetState: 'visible',
      tableAnchorScore: 40,
      hasFormUTitle: true,
      hasEmployeeRegisterTitle: true,
      hasEmployeeTable: true
    });
    expect(visibleFormU).toBeGreaterThan(hiddenHelper);
  });

  test('writes Excel values onto the merge master so they are visible', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('FORM U');
    ws.mergeCells(10, 2, 10, 3);
    writeFormUTamilNaduExcelVisibleCell(ws, 10, 3, 'Babu');
    expect(ws.getCell(10, 2).value).toBe('Babu');
  });

  test('fills employee rows under the 1-2-3 index strip on the FORM U sheet', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('FORM U');
    ws.getCell(1, 8).value = 'FORM - U';
    ws.getCell(2, 8).value = 'EMPLOYEE REGISTER';
    ws.getCell(4, 1).value = 'Name and Address of the Establishment:';
    ws.getCell(5, 1).value = 'Name and Address of the Employer:';
    const headers = [
      'S.No',
      'Name of the employee',
      'Employee Identification No.',
      'Gender',
      'Father / Spouse Name',
      'Date of Birth',
      'Date of Joining',
      'Designation',
      'Present Address',
      'Permanent address'
    ];
    headers.forEach((label, i) => {
      ws.getCell(8, i + 1).value = label;
      ws.getCell(9, i + 1).value = i + 1;
    });
    const filled = fillFormUTamilNaduEmployeeTable({
      worksheet: ws,
      headerRow: 8,
      startCol: 1,
      numberingRowBelowHeader: true,
      headersToUse: FORM_U_TAMILNADU_HEADERS,
      mappedData: [
        {
          'S.No': 1,
          'Name of the employee': 'Babu',
          'Worker Identity No.': 'VE1111',
          Gender: 'Male',
          'Father / Spouse Name': 'PM MUTHU',
          'Date of Birth': '02-Jul-80',
          'Date of Joining': '01-Dec-2025',
          Designation: 'Engineer',
          'Present Address': '12 Main St',
          'Permanent Address': '12 Main St'
        }
      ]
    });
    expect(filled.totalRows).toBe(1);
    expect(filled.startRow).toBe(10);
    expect(String(ws.getCell(10, 2).value)).toBe('Babu');
    expect(String(ws.getCell(10, 3).value)).toBe('VE1111');
    expect(String(ws.getCell(10, 4).value)).toBe('Male');
    expect(String(ws.getCell(10, 8).value)).toBe('Engineer');
  });
});
