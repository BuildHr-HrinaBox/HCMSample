import {
  cloneFormXIIIAPWorksheetClean,
  buildFormXIIIAPWorkbookWithTemplateStyles,
  isFormXIIIRegisterOfWorkmenContext,
  isFormXAPRegisterOfFinesContext,
  blobIndicatesFormIRegisterOfFinesNotFormX,
  sheetBlobIndicatesFormXAPRegisterOfFines,
  sheetBlobIndicatesFormXLeaveRegister,
  matchesFormXHint,
  FORM_XX_AP_DEDUCTION_COLUMN_NIL_TEXT,
  FORM_XXI_AP_FINE_COLUMN_NIL_TEXT,
  isFormXXAPDeductionNilHeader,
  applyFormXXAPDeductionsNilToMappedRows,
  resolveFormXXAPLeafHeaders,
  dedupeFormXXAPConcatenatedHeader,
  normalizeFormXXAPExportHeaders,
  pickFormXXAPExportHeaders,
  exportFormXXAPRowValuesByHeaders,
  reapplyFormXXAPDownloadTableBordersFromWorksheet,
  isFormXXIAPFineNilHeader,
  applyFormXXIAPFinesNilToMappedRows,
} from './formXAPRegisterOfFines';
import { excelJSCellHasFullBoxBorder } from '../../utils/excelTableBorders';
import ExcelJS from 'exceljs';
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

  it('does not classify Rajasthan Form_14_RJ as Form XIII', () => {
    expect(
      isFormXIIIRegisterOfWorkmenContext(
        { title: 'FORM XIII', subtitle: 'Register of Workmen' },
        { state: 'Rajasthan', formName: 'Form 14_RJ' },
        'Form_14_RJ_-_Rajasthan.xlsx',
        'FORM XIII Register of Workmen Rule 22 - sub rule 1'
      )
    ).toBe(false);
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
    expect(merges).toEqual(expect.arrayContaining(['J5:K5', 'J6:K8', 'J9:K12']));
    expect(merges).not.toEqual(expect.arrayContaining(['J5:K8']));
    expect(String(worksheet.getCell(5, 10).value || '')).toMatch(
      /establishemnt in\/ under which contract is carried on/i
    );
    expect(String(worksheet.getCell(5, 10).value || '')).not.toMatch(/VAYONA|Gurvepalli/i);
  });

  it('copies Principal Employer value below the establishment heading', async () => {
    const ExcelJS = require('exceljs');
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Form XIII-Register of Workmen');
    ws.getCell(3, 1).value = 'FORM - XIII REGISTER OF WORKMEN EMPLOYED BY CONTRACTOR';
    ws.getCell(4, 1).value =
      '[Vide Rule 75 Contract Labour (Regulation and Abolition) Central/A.P Rules)';
    ws.getCell(5, 2).value = 'Name and Address of Contractor. : VAYONA ENERGY PRIVATE LIMITED';
    ws.getCell(9, 2).value = 'Nature and location of work. : AP-Tadipatri';
    ws.getCell(5, 9).value = 'Name and address of Establishemnt in/ under which contract is carried on:';
    ws.getCell(9, 10).value =
      'Name and address of Principal Employer : Vibrant greentech India PVT LTD (HCL) & Traditional Customers, M/S VIBRANT';
    ws.getCell(13, 1).value = 'S.No';
    ws.getCell(13, 2).value = 'Name and Surname of Workmen';

    const { worksheet } = cloneFormXIIIAPWorksheetClean(ws, 'Form XIII-Register of Workmen');
    expect(String(worksheet.getCell(5, 10).value || '')).toMatch(
      /establishemnt in\/ under which contract is carried on/i
    );
    expect(String(worksheet.getCell(5, 10).value || '')).not.toMatch(/Vibrant/i);
    expect(String(worksheet.getCell(6, 10).value || '')).toMatch(/Vibrant greentech India PVT LTD \(HCL\)/i);
    expect(String(worksheet.getCell(6, 10).value || '')).not.toMatch(/Principal Employer/i);
    expect(String(worksheet.getCell(9, 10).value || '')).toMatch(/Name and address of Principal Employer/i);
    expect(String(worksheet.getCell(9, 10).value || '')).toMatch(/Vibrant greentech/i);
    const merges = Array.isArray(worksheet.model?.merges) ? worksheet.model.merges : [];
    expect(merges).toEqual(expect.arrayContaining(['J6:K8', 'J9:K12']));
  });

  it('writes headerFormData Principal Employer below the establishment heading on download', async () => {
    const ExcelJS = require('exceljs');
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Form XIII-Register of Workmen');
    ws.getCell(3, 1).value = 'FORM - XIII REGISTER OF WORKMEN EMPLOYED BY CONTRACTOR';
    ws.getCell(4, 1).value =
      '[Vide Rule 75 Contract Labour (Regulation and Abolition) Central/A.P Rules)';
    ws.getCell(5, 2).value = 'Name and Address of Contractor. : VAYONA ENERGY PRIVATE LIMITED';
    ws.getCell(9, 2).value = 'Nature and location of work. : AP-Tadipatri';
    ws.getCell(5, 10).value = 'Name and address of Establishemnt in/ under which contract is carried on:';
    ws.getCell(9, 10).value = 'Name and address of Principal Employer :';
    ws.getCell(13, 1).value = 'S.No';
    ws.getCell(13, 2).value = 'Name and Surname of Workmen';
    const templateArrayBuffer = await wb.xlsx.writeBuffer();
    const pe =
      'Vibrant greentech India PVT LTD (HCL) & Traditional Customers, M/S VIBRANT';
    const { blob } = await buildFormXIIIAPWorkbookWithTemplateStyles({
      templateArrayBuffer,
      mappedData: [],
      headersToUse: ['S.No', 'Name and Surname of Workmen'],
      parsedHeaderRowIndex: 12,
      parsedDataStartIndex: 14,
      parsedFormHeader: {
        fields: [
          {
            key: 'form_xiii_establishment',
            label: 'Name and address of Establishemnt in/ under which contract is carried on',
            value: ''
          },
          {
            key: 'statutory_principal_employer',
            label: 'Name and address of Principal Employer',
            value: pe
          }
        ]
      },
      headerFormData: { statutory_principal_employer: pe }
    });
    const outWb = new ExcelJS.Workbook();
    await outWb.xlsx.load(await new Response(blob).arrayBuffer());
    const outWs = outWb.worksheets[0];
    expect(String(outWs.getCell(5, 10).value || '')).toMatch(
      /establishemnt in\/ under which contract is carried on/i
    );
    expect(String(outWs.getCell(5, 10).value || '')).not.toMatch(/Vibrant/i);
    expect(String(outWs.getCell(6, 10).value || '')).toBe(pe);
    expect(String(outWs.getCell(9, 10).value || '')).toMatch(/Name and address of Principal Employer/i);
    expect(String(outWs.getCell(9, 10).value || '')).toMatch(/Vibrant/);
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
    // Short child labels under Date of recovery
    expect(isFormXXAPDeductionNilHeader('First')).toBe(true);
    expect(isFormXXAPDeductionNilHeader('Last')).toBe(true);
    expect(
      isFormXXAPDeductionNilHeader(
        'Whether work man showed cause against deduction Amount of deduction imposed'
      )
    ).toBe(true);
    expect(isFormXXAPDeductionNilHeader('Name of workmen')).toBe(false);
    expect(isFormXXAPDeductionNilHeader('S.No')).toBe(false);
    expect(isFormXXAPDeductionNilHeader('First Name')).toBe(false);
    expect(isFormXXAPDeductionNilHeader('Last Name')).toBe(false);
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

  it('puts NIL on short First / Last recovery child headers', () => {
    const rows = applyFormXXAPDeductionsNilToMappedRows(
      [{ 'Name of workmen': 'Ravi', First: '01-01-2026', Last: '01-03-2026' }],
      ['Name of workmen', 'First', 'Last'],
      undefined,
      { overwrite: true }
    );
    expect(rows[0]['Name of workmen']).toBe('Ravi');
    expect(rows[0].First).toBe(FORM_XX_AP_DEDUCTION_COLUMN_NIL_TEXT);
    expect(rows[0].Last).toBe(FORM_XX_AP_DEDUCTION_COLUMN_NIL_TEXT);
  });

  it('expands Date of recovery child First/Last into instalment leaf headers', () => {
    const rows = [
      [
        'S.No',
        'Name of workmen',
        'Date of recovery',
        'Date of recovery',
        'Date of recovery',
        'Remarks',
      ],
      ['', '', 'No. of instalments', 'First', 'Last', ''],
    ];
    const labels = resolveFormXXAPLeafHeaders(rows, [], 0, 0, 6);
    expect(labels).toEqual([
      'S.No',
      'Name of workmen',
      'No. of instalments',
      'First instalment',
      'Last instalment',
      'Remarks',
    ]);
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

describe('Form XX AP Father / Designation export headers', () => {
  const duplicatedUiHeaders = [
    'S.No',
    'Name of Workmen',
    "Father's/Husband's Name_Father's/Husband's Name",
    'Nature of employement /Designation_Nature of employement /Designation',
    'Particulars of Damage or Loss',
    'Remarks',
  ];

  const layoutHeaders = [
    'S.No',
    'Name of Workmen',
    "Father's/Husband's Name",
    'Nature of employment /Designation',
    'Particulars of Damage or Loss',
    'Remarks',
  ];

  it('dedupes concatenated modal headers', () => {
    expect(
      dedupeFormXXAPConcatenatedHeader(
        "Father's/Husband's Name_Father's/Husband's Name"
      )
    ).toBe("Father's/Husband's Name");
    expect(
      dedupeFormXXAPConcatenatedHeader(
        'Nature of employement /Designation_Nature of employement /Designation'
      )
    ).toBe('Nature of employement /Designation');
  });

  it('prefers UI headers when Father and Designation are present', () => {
    const picked = pickFormXXAPExportHeaders(duplicatedUiHeaders, layoutHeaders);
    expect(picked).toEqual(normalizeFormXXAPExportHeaders(duplicatedUiHeaders));
    expect(picked.some((h) => /father|husband/i.test(h))).toBe(true);
    expect(picked.some((h) => /nature\s+of\s+employ|designat/i.test(h))).toBe(true);
  });

  it('exports Father and Designation from duplicated UI row keys', () => {
    const row = {
      'Name of Workmen': 'Pulu Venkata Raman',
      "Father's/Husband's Name_Father's/Husband's Name": 'pulu nagaiah',
      'Nature of employement /Designation_Nature of employement /Designation': 'Senior Engineer',
    };
    const values = exportFormXXAPRowValuesByHeaders(row, duplicatedUiHeaders);
    expect(values[1]).toBe('Pulu Venkata Raman');
    expect(values[2]).toBe('pulu nagaiah');
    expect(values[3]).toBe('Senior Engineer');
  });
});

describe('Form XXI AP Register of Fines NIL columns', () => {
  const headers = [
    'S.No',
    'Name of workmen',
    'Father/Husband',
    'Act/Omission for which fine imposed',
    'Date of Offence',
    'Whether workman showed cause against fine',
    "Name of Person in whose presence Employee's explanation was heard",
    'Wage - period and wages payable',
    'Amount of fine Imposed',
    'Date on which fine realised',
  ];

  it('recognizes fine / offence / wage-period columns for NIL', () => {
    expect(isFormXXIAPFineNilHeader('Act/Omission for which fine imposed')).toBe(true);
    expect(isFormXXIAPFineNilHeader('Date of Offence')).toBe(true);
    expect(isFormXXIAPFineNilHeader('Whether workman showed cause against fine')).toBe(true);
    expect(
      isFormXXIAPFineNilHeader(
        "Name of Person in whose presence Employee's explanation was heard"
      )
    ).toBe(true);
    expect(isFormXXIAPFineNilHeader('Wage - period and wages payable')).toBe(true);
    expect(isFormXXIAPFineNilHeader('Amount of fine Imposed')).toBe(true);
    expect(isFormXXIAPFineNilHeader('Date on which fine realised')).toBe(true);
    expect(isFormXXIAPFineNilHeader('Name of workmen')).toBe(false);
    expect(isFormXXIAPFineNilHeader('S.No')).toBe(false);
  });

  it('fills listed Form XXI fine columns with NIL', () => {
    const rows = applyFormXXIAPFinesNilToMappedRows(
      [
        {
          'Name of workmen': 'Ravi',
          'Act/Omission for which fine imposed': '',
          'Wage - period and wages payable': '12500',
          "Name of Person in whose presence Employee's explanation was heard": 'Karthick',
        },
      ],
      headers,
      undefined,
      { overwrite: true }
    );
    expect(rows[0]['Name of workmen']).toBe('Ravi');
    expect(rows[0]['Act/Omission for which fine imposed']).toBe(FORM_XXI_AP_FINE_COLUMN_NIL_TEXT);
    expect(rows[0]['Date of Offence']).toBe(FORM_XXI_AP_FINE_COLUMN_NIL_TEXT);
    expect(rows[0]['Whether workman showed cause against fine']).toBe(
      FORM_XXI_AP_FINE_COLUMN_NIL_TEXT
    );
    expect(rows[0]["Name of Person in whose presence Employee's explanation was heard"]).toBe(
      FORM_XXI_AP_FINE_COLUMN_NIL_TEXT
    );
    expect(rows[0]['Wage - period and wages payable']).toBe(FORM_XXI_AP_FINE_COLUMN_NIL_TEXT);
    expect(rows[0]['Amount of fine Imposed']).toBe(FORM_XXI_AP_FINE_COLUMN_NIL_TEXT);
    expect(rows[0]['Date on which fine realised']).toBe(FORM_XXI_AP_FINE_COLUMN_NIL_TEXT);
  });
});

describe('Form XX AP download table borders', () => {
  it('applies full box borders on data rows 15–16 like the template model', () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Form XX');
    // Rich-text title (previously broke String(cell.value) detection).
    ws.getCell(2, 1).value = {
      richText: [
        { text: 'FORM - XX REGISTER OF DEDUCTIONS FOR DAMAGE OR LOSS' },
      ],
    };
    ws.getCell(13, 1).value = 'S.No';
    ws.getCell(13, 2).value = 'Name of Workmen';
    ws.getCell(13, 3).value = "Father's/Husband's Name";
    ws.getCell(13, 4).value = 'Nature of employment /Designation';
    ws.getCell(13, 5).value = 'Particulars of Damage or Loss';
    ws.getCell(13, 6).value = 'Date of Damage or Loss';
    ws.getCell(13, 7).value = 'Whether workman showed cause against';
    ws.getCell(13, 8).value = "Name of Person in whose presence Employee's explanation was heard";
    ws.getCell(13, 9).value = 'Amount of deduction imposed';
    ws.getCell(13, 10).value = 'Date of recovery';
    ws.getCell(13, 11).value = 'Date of recovery';
    ws.getCell(13, 12).value = 'Date of recovery';
    ws.getCell(13, 13).value = 'Remarks';
    ws.getCell(14, 10).value = 'No. of instalments';
    ws.getCell(14, 11).value = 'First instalment';
    ws.getCell(14, 12).value = 'Last instalment';

    ws.getCell(15, 1).value = 1;
    ws.getCell(15, 2).value = 'Polu Venkata Ramana';
    for (let c = 5; c <= 13; c += 1) ws.getCell(15, c).value = 'NIL';
    ws.getCell(16, 1).value = 2;
    ws.getCell(16, 2).value = 'Arun Kumar Krishnan';
    for (let c = 5; c <= 13; c += 1) ws.getCell(16, c).value = 'NIL';

    reapplyFormXXAPDownloadTableBordersFromWorksheet(ws);

    for (const row of [15, 16]) {
      for (let c = 1; c <= 13; c += 1) {
        expect(excelJSCellHasFullBoxBorder(ws.getCell(row, c))).toBe(true);
      }
      expect(ws.getCell(row, 1).alignment?.horizontal).toBe('center');
      expect(ws.getCell(row, 2).alignment?.horizontal).toBe('left');
      expect(ws.getCell(row, 5).alignment?.horizontal).toBe('center');
      expect(ws.getCell(row, 5).alignment?.vertical).toBe('middle');
    }
  });
});
