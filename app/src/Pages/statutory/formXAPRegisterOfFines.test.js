import {
  cloneFormXIIIAPWorksheetClean,
  isFormXIIIRegisterOfWorkmenContext,
  isFormXAPRegisterOfFinesContext,
  blobIndicatesFormIRegisterOfFinesNotFormX,
  sheetBlobIndicatesFormXAPRegisterOfFines,
  sheetBlobIndicatesFormXLeaveRegister,
  matchesFormXHint,
  FORM_XX_AP_DEDUCTION_COLUMN_NIL_TEXT,
  isFormXXAPDeductionNilHeader,
  applyFormXXAPDeductionsNilToMappedRows,
} from './formXAPRegisterOfFines';
import { isFormXRajasthanEmploymentCardContext } from './formXIVMPEmploymentCard';

describe('Form X hint vs Form X_RJ Employment Card', () => {
  it('does not treat Form X_RJ as AP/TN Form X', () => {
    expect(matchesFormXHint('Form X_RJ')).toBe(false);
    expect(matchesFormXHint('Form_X_RJ.xlsx')).toBe(false);
    expect(matchesFormXHint('FORM X RJ Employment Card')).toBe(false);
    expect(
      matchesFormXHint('Rajasthan CLRA Form X Employment Card Rule 75')
    ).toBe(false);
  });

  it('still matches plain Form X leave / fines labels', () => {
    expect(matchesFormXHint('Form X')).toBe(true);
    expect(matchesFormXHint('Form_X_-_TamilNadu.xlsx')).toBe(true);
    expect(matchesFormXHint('FORM X Register of Leave')).toBe(true);
  });

  it('detects Rajasthan Form X Employment Card context', () => {
    expect(
      isFormXRajasthanEmploymentCardContext(
        { title: 'FORM X', subtitle: 'Employment Card', reference: '(See rule 75)' },
        { state: 'Rajasthan', formName: 'Form X_RJ' },
        'Form_X_RJ.xlsx',
        'Employment Card'
      )
    ).toBe(true);
  });

  it('detects leave register sheets and rejects them for Form X_RJ sheet pick', () => {
    expect(
      sheetBlobIndicatesFormXLeaveRegister(
        'FORM X REGISTER OF LEAVE Name of the employee Earned Leave Medical Leave'
      )
    ).toBe(true);
    expect(
      sheetBlobIndicatesFormXLeaveRegister(
        'FORM X Employment Card Name of the workman Wage period Period of employment'
      )
    ).toBe(false);
  });
});

describe('Form XIII Register of Workmen detection', () => {
  it('does not classify Tamil Nadu Form 1 conferment register as Form XIII', () => {
    expect(
      isFormXIIIRegisterOfWorkmenContext(
        {
          title: 'FORM 1',
          subtitle: 'REGISTER OF WORKMEN',
          reference: 'Tamil Nadu Industrial Establishments (Conferment of Permanent Status) Act'
        },
        { state: 'TamilNadu', formName: 'Register of Conferment Status' },
        'Form_I_-_TamilNadu.xlsx',
        'REGISTER OF WORKMEN'
      )
    ).toBe(false);
  });

  it('still detects CLRA Form XIII Register of Workmen', () => {
    expect(
      isFormXIIIRegisterOfWorkmenContext(
        { title: 'FORM XIII', subtitle: 'Register of Workmen' },
        { state: 'Andhra Pradesh', act: 'Contract Labour' },
        'Form_XIII.xlsx',
        'Register of Workmen (Contract Labour)'
      )
    ).toBe(true);
  });

  it('detects AP Form XIII from Rule 75 / workmen surname headers', () => {
    expect(
      isFormXIIIRegisterOfWorkmenContext(
        {
          title: 'FORM - XIII REGISTER OF WORKMEN EMPLOYED BY CONTRACTOR',
          subtitle: '[Vide Rule 75 Contract Labour (Regulation and Abolition) Central/A.P Rules)'
        },
        { state: 'Andhra Pradesh', formName: 'Form XIII' },
        'Form XIII - Andhra Pradesh.xlsx',
        'Name and Surname of Workmen Age and Sex Local address Permanent House address'
      )
    ).toBe(true);
  });
});

describe('Form XIII AP clean workbook clone', () => {
  it('dedupes repeated title text and applies a single A3:K3 merge', async () => {
    const ExcelJS = require('exceljs');
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('FORM 1');
    for (let c = 1; c <= 11; c += 1) {
      ws.getCell(3, c).value =
        'FORM - XIII REGISTER OF WORKMEN EMPLOYED BY CONTRACTOR [Vide Rule 75 Contract Labour (Regulation and Abolition) Central/A.P Rules)';
      ws.getCell(4, c).value =
        '[Vide Rule 75 Contract Labour (Regulation and Abolition) Central/A.P Rules)';
      ws.getCell(13, c).value = c === 1 ? 'S.No' : c === 2 ? 'Name and Surname of Workmen' : `Col ${c}`;
      ws.getCell(14, c).value = ws.getCell(13, c).value;
    }
    ws.getCell(5, 1).value = 'Name and address of Contractor : VAYONA ENERGY';
    ws.getCell(5, 2).value = 'Name and Address of Contractor. : VAYONA ENERGY PRIVATE LIMITED';
    ws.getCell(9, 1).value = 'Nature and location of work : AP-Gurvepalli';
    ws.getCell(9, 2).value = 'Nature and location of work. : AP-Gurvepalli';
    ws.getCell(5, 9).value = 'Name and address of Establishemnt in/ under which contract is carried on:';
    ws.getCell(5, 10).value = 'Name and addr Establishemnt i contract is carri';
    ws.getCell(15, 1).value = '1';
    ws.getCell(15, 2).value = '2';

    const { workbook, worksheet } = cloneFormXIIIAPWorksheetClean(ws, 'FORM 1');
    expect(workbook.worksheets).toHaveLength(1);
    expect(String(worksheet.name)).toMatch(/XIII/i);
    expect(String(worksheet.getCell(3, 1).value)).toMatch(/FORM\s*[-–]?\s*XIII/i);
    expect(String(worksheet.getCell(3, 1).value)).not.toMatch(/Vide Rule 75/i);
    expect(String(worksheet.getCell(4, 1).value)).toMatch(/Vide Rule 75/i);
    const merges = Array.isArray(worksheet.model?.merges) ? worksheet.model.merges : [];
    expect(merges).toEqual(expect.arrayContaining(['A3:K3', 'A4:K4', 'A13:A14']));
    expect(String(worksheet.getCell(13, 1).value)).toMatch(/S\.?\s*No/i);
    // Columns A and I must be cleared; combined text stays in B / J only.
    expect(String(worksheet.getCell(5, 1).value || '')).toBe('');
    expect(String(worksheet.getCell(9, 1).value || '')).toBe('');
    expect(String(worksheet.getCell(5, 9).value || '')).toBe('');
    expect(String(worksheet.getCell(5, 2).value)).toMatch(/VAYONA ENERGY PRIVATE LIMITED/i);
    expect(String(worksheet.getCell(9, 2).value)).toMatch(/AP-Gurvepalli/i);
  });
});

describe('Form I Tamil Nadu vs Form XIII AP', () => {
  it('does not classify CLRA Form XIII as Tamil Nadu Form I workmen', () => {
    // isFormIRegisterOfWorkmenContext lives in Statutory.js; verify Form XIII detector wins.
    expect(
      isFormXIIIRegisterOfWorkmenContext(
        {
          title: 'FORM - XIII REGISTER OF WORKMEN EMPLOYED BY CONTRACTOR',
          subtitle: '[Vide Rule 75 Contract Labour (Regulation and Abolition) Central/A.P Rules)'
        },
        { state: 'Andhra Pradesh', formName: 'Form XIII' },
        'Form XIII - Andhra Pradesh.xlsx',
        'REGISTER OF WORKMEN EMPLOYED BY CONTRACTOR'
      )
    ).toBe(true);
  });
});

describe('Form I Tamil Nadu vs Form X AP Register of Fines', () => {
  it('does not classify Form_I_-_TamilNadu Register of Fines as Form X AP', () => {
    const sheetText =
      'FORM I REGISTER OF FINES [See Rule 3] Payment of Wages Act Minimum Wages Act Name of the Establishment';
    expect(
      blobIndicatesFormIRegisterOfFinesNotFormX(
        `Form I Register of Fines Form_I_-_TamilNadu.xlsx ${sheetText}`
      )
    ).toBe(true);
    expect(
      sheetBlobIndicatesFormXAPRegisterOfFines(
        `Form I Register of Fines Form_I_-_TamilNadu.xlsx ${sheetText}`
      )
    ).toBe(false);
    expect(
      isFormXAPRegisterOfFinesContext(
        { title: 'FORM I', subtitle: 'REGISTER OF FINES', reference: '[See Rule 3]' },
        { formName: 'Form I', state: 'Tamil Nadu' },
        'Form_I_-_TamilNadu.xlsx',
        sheetText
      )
    ).toBe(false);
  });

  it('still detects AP Shops Form X Register of Fines', () => {
    expect(
      isFormXAPRegisterOfFinesContext(
        { title: 'Form X', subtitle: 'Register of Fines' },
        { formName: 'Form X', state: 'Andhra Pradesh', act: 'Shops and Establishment' },
        'Form_X_AP.xlsx',
        'Form X Register of Fines Name of the worker Nature & date of offence'
      )
    ).toBe(true);
  });
});

describe('Form XX AP Register of Deductions NIL columns', () => {
  const headers = [
    'S.No',
    'Name of workmen',
    'Father/Husband',
    'Particulars of Damage or Loss',
    'Date of Damage or Loss',
    'Whether workman showed cause against deduction',
    "Name of Person in whose presence Employee's explanation was heard",
    'Amount of deduction imposed',
    'No. of instalments',
    'First instalment',
    'Last instalment',
    'Remarks',
  ];

  it('recognizes damage / recovery columns for NIL', () => {
    expect(isFormXXAPDeductionNilHeader('Particulars of Damage or Loss')).toBe(true);
    expect(isFormXXAPDeductionNilHeader('Date of Damage or Loss')).toBe(true);
    expect(isFormXXAPDeductionNilHeader('Whether workman showed cause against deduction')).toBe(true);
    expect(
      isFormXXAPDeductionNilHeader(
        "Name of Person in whose presence Employee's explanation was heard"
      )
    ).toBe(true);
    expect(isFormXXAPDeductionNilHeader('Amount of deduction imposed')).toBe(true);
    expect(isFormXXAPDeductionNilHeader('No. of instalments')).toBe(true);
    expect(isFormXXAPDeductionNilHeader('First instalment')).toBe(true);
    expect(isFormXXAPDeductionNilHeader('Last instalment')).toBe(true);
    expect(isFormXXAPDeductionNilHeader('Remarks')).toBe(true);
    expect(isFormXXAPDeductionNilHeader('No. of Installment')).toBe(true);
    expect(isFormXXAPDeductionNilHeader('First installment')).toBe(true);
    expect(isFormXXAPDeductionNilHeader('Last installment')).toBe(true);
    expect(
      isFormXXAPDeductionNilHeader(
        'Whether work man showed cause against deduction Amount of deduction imposed'
      )
    ).toBe(true);
    expect(isFormXXAPDeductionNilHeader('Name of workmen')).toBe(false);
    expect(isFormXXAPDeductionNilHeader('S.No')).toBe(false);
  });

  it('fills blank damage / recovery columns with NIL', () => {
    const rows = applyFormXXAPDeductionsNilToMappedRows(
      [{ 'Name of workmen': 'Ravi', 'Particulars of Damage or Loss': '' }],
      headers
    );
    expect(rows[0]['Name of workmen']).toBe('Ravi');
    expect(rows[0]['Particulars of Damage or Loss']).toBe(FORM_XX_AP_DEDUCTION_COLUMN_NIL_TEXT);
    expect(rows[0]['Date of Damage or Loss']).toBe(FORM_XX_AP_DEDUCTION_COLUMN_NIL_TEXT);
    expect(rows[0]['Amount of deduction imposed']).toBe(FORM_XX_AP_DEDUCTION_COLUMN_NIL_TEXT);
    expect(rows[0]['No. of instalments']).toBe(FORM_XX_AP_DEDUCTION_COLUMN_NIL_TEXT);
    expect(rows[0]['First instalment']).toBe(FORM_XX_AP_DEDUCTION_COLUMN_NIL_TEXT);
    expect(rows[0]['Last instalment']).toBe(FORM_XX_AP_DEDUCTION_COLUMN_NIL_TEXT);
    expect(rows[0]['Remarks']).toBe(FORM_XX_AP_DEDUCTION_COLUMN_NIL_TEXT);
  });

  it('overwrites People-fetched names with NIL on download', () => {
    const rows = applyFormXXAPDeductionsNilToMappedRows(
      [
        {
          'Name of workmen': 'Ravi',
          "Name of Person in whose presence Employee's explanation was heard": 'Karthick',
          'No. of Installment': '3',
          'First installment': '01-01-2026',
          'Last installment': '01-03-2026',
          'Whether workman showed cause against deduction': 'Yes',
        },
      ],
      [
        'Name of workmen',
        "Name of Person in whose presence Employee's explanation was heard",
        'Whether workman showed cause against deduction',
        'No. of Installment',
        'First installment',
        'Last installment',
      ],
      undefined,
      { overwrite: true }
    );
    expect(rows[0]['Name of workmen']).toBe('Ravi');
    expect(rows[0]["Name of Person in whose presence Employee's explanation was heard"]).toBe(
      FORM_XX_AP_DEDUCTION_COLUMN_NIL_TEXT
    );
    expect(rows[0]['Whether workman showed cause against deduction']).toBe(
      FORM_XX_AP_DEDUCTION_COLUMN_NIL_TEXT
    );
    expect(rows[0]['No. of Installment']).toBe(FORM_XX_AP_DEDUCTION_COLUMN_NIL_TEXT);
    expect(rows[0]['First installment']).toBe(FORM_XX_AP_DEDUCTION_COLUMN_NIL_TEXT);
    expect(rows[0]['Last installment']).toBe(FORM_XX_AP_DEDUCTION_COLUMN_NIL_TEXT);
  });
});
