import ExcelJS from 'exceljs';
import {
  FORM_A_RJ_CANONICAL_TABLE_HEADERS,
  applyFormARajasthanContractorFromSite,
  buildFormARajasthanWorkbookWithTemplateStyles,
  enrichFormARajasthanDisplayHeader,
  headersIndicateFormARajasthanTable,
  headersIndicateFormARajasthanHybridTemplate,
  isFormARajasthanContext,
  isFormARajasthanLikeExport,
  resolveFormARajasthanExportHeaders,
  sheetBodyHasEmploymentCardLabels,
  sheetLooksLikeFormXIVEmploymentCard,
  writeFormARajasthanTitleBandsInColumnsEG,
} from './formARajasthan';
import { formAGJHeaderAliasBucket, formAGJGujaratHeaderNorm } from './formAGJGujarat';

async function buildBlankFormARjTemplate() {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('FORM A');
  ws.getCell(1, 1).value = '[See rule 2(1)]';
  ws.getCell(2, 1).value = 'FORM A';
  ws.getCell(3, 1).value = 'FORMAT OF EMPLOYEE REGISTER';
  ws.getCell(4, 1).value =
    'Name and address of establishment in / under which contract is carried on';
  ws.getCell(5, 1).value = 'Name and address of contractor';
  FORM_A_RJ_CANONICAL_TABLE_HEADERS.forEach((h, i) => {
    ws.getCell(7, i + 1).value = h;
  });
  for (let r = 8; r <= 18; r += 1) {
    for (let c = 1; c <= 15; c += 1) {
      ws.getCell(r, c).border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' },
      };
    }
  }
  ws.getCell(19, 1).value = '*(Highly Skilled/Skilled/Semi Skilled/Un Skilled)';
  ws.getCell(20, 1).value =
    '#Note: In case the age is between 14 to 18 years, mention the nature of work, daily hours of work and Intervals of rest in the remarks Column.';
  return wb.xlsx.writeBuffer();
}

/** Broken hybrid template — table header row but employment-card labels in column B body. */
async function buildHybridBrokenFormARjTemplate() {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('FORM A');
  ws.getCell(1, 1).value = '[See rule 2(1)]';
  ws.getCell(2, 1).value = 'FORM A';
  ws.getCell(3, 1).value = 'FORMAT OF EMPLOYEE REGISTER';
  ws.getCell(4, 1).value =
    'Name and address of establishment in / under which contract is carried on';
  ws.getCell(5, 1).value = 'Name and address of contractor';
  const wrongHeaders = [
    'Sl. N',
    'Employee Code',
    'Nature of work and location of work',
    'Surname',
    "Father's/Spouse N",
    'Date of Birth',
    'Nationality',
    'Education',
  ];
  wrongHeaders.forEach((h, i) => {
    ws.getCell(7, i + 1).value = h;
  });
  const cardLabels = [
    'Name of the workman',
    'Serial number in the register of workmen employed',
    'Nature of employment / designation',
    "Wage rate with particulars or unit, in case of piece of work",
    'Wage period',
    'Tenure of employment',
    'Remarks',
  ];
  cardLabels.forEach((label, i) => {
    ws.getCell(8 + i, 2).value = label;
  });
  ws.getColumn(2).width = 42;
  ws.getCell(19, 1).value = '*(Highly Skilled/Skilled/Semi Skilled/Un Skilled)';
  return wb.xlsx.writeBuffer();
}

describe('Form A RJ original template export', () => {
  test('detects Form A_RJ context from filename', () => {
    expect(
      isFormARajasthanContext(
        { title: 'FORM A', subtitle: 'FORMAT OF EMPLOYEE REGISTER' },
        { formFileName: 'Form_A_RJ_-_Rajasthan.xlsx', formName: 'Format of Employee Register' },
        'Form_A_RJ_-_Rajasthan.xlsx',
        '',
        FORM_A_RJ_CANONICAL_TABLE_HEADERS
      )
    ).toBe(true);
  });

  test('detects Form A_MH context from filename', () => {
    expect(
      isFormARajasthanContext(
        { title: 'FORM A', subtitle: 'FORMAT OF EMPLOYEE REGISTER' },
        { formFileName: 'Form_A_MH_-_Maharashtra.xlsx', formName: 'Format of Employee Register' },
        'Form_A_MH_-_Maharashtra.xlsx',
        '',
        FORM_A_RJ_CANONICAL_TABLE_HEADERS
      )
    ).toBe(true);
  });

  test('does not treat Form XIV employment card as Form A RJ', () => {
    expect(
      isFormARajasthanContext(
        { title: 'FORM XIV', subtitle: 'Employment Card' },
        { formFileName: 'Form_XIV_MP.xlsx', formName: 'Employment Card' },
        'Form_XIV_MP.xlsx',
        'Name of the workman Tenure of employment',
        ['Name of the workman', 'Tenure of employment']
      )
    ).toBe(false);
  });

  test('still detects Form A_RJ when hybrid template sheet text has employment-card labels', () => {
    expect(
      isFormARajasthanContext(
        { title: 'FORM A', subtitle: 'FORMAT OF EMPLOYEE REGISTER' },
        { formFileName: 'Form_A_RJ_-_Rajasthan.xlsx', formName: 'Format of Employee Register' },
        'Form A RJ - Rajasthan.xlsx',
        'Name of the workman Serial number in the register Tenure of employment Wage period',
        ['Sl. N', 'Employee Code', 'Nature of work and location of work', 'Surname']
      )
    ).toBe(true);
  });

  test('canonical headers match original Part-A columns', () => {
    expect(headersIndicateFormARajasthanTable(FORM_A_RJ_CANONICAL_TABLE_HEADERS)).toBe(true);
    expect(resolveFormARajasthanExportHeaders([])).toEqual(FORM_A_RJ_CANONICAL_TABLE_HEADERS);
    expect(formAGJHeaderAliasBucket(formAGJGujaratHeaderNorm('Employee Code'))).toBe('employeeId');
    expect(
      formAGJHeaderAliasBucket(formAGJGujaratHeaderNorm('Category (HS/S/SS/US)*'))
    ).toBe('skillCategory');
  });

  test('moves FORM A heading to columns E–G', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('FORM A');
    ws.getCell(2, 1).value = 'FORM A';
    ws.getCell(3, 1).value = 'FORMAT OF EMPLOYEE REGISTER';
    writeFormARajasthanTitleBandsInColumnsEG(ws, {});
    expect(String(ws.getCell(2, 1).value || '')).toBe('');
    expect(String(ws.getCell(2, 5).value || '')).toBe('FORM A');
    expect(String(ws.getCell(3, 5).value || '')).toBe('FORMAT OF EMPLOYEE REGISTER');
  });

  test('moves [See rule 2(1)] above FORM A out of column O', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('FORM A');
    ws.getCell(1, 15).value = '[See rule 2(1)]';
    ws.getCell(2, 5).value = 'FORM A';
    ws.getCell(3, 5).value = 'FORMAT OF EMPLOYEE REGISTER';
    writeFormARajasthanTitleBandsInColumnsEG(ws, {});
    expect(String(ws.getCell(1, 15).value || '')).toBe('');
    expect(String(ws.getCell(1, 5).value || '')).toMatch(/see rule 2\(1\)/i);
    expect(String(ws.getCell(2, 5).value || '')).toBe('FORM A');
    expect(String(ws.getCell(3, 5).value || '')).toBe('FORMAT OF EMPLOYEE REGISTER');
  });

  test('places [Part-A: For all Establishments] left below FORMAT OF EMPLOYEE REGISTER', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('FORM A');
    ws.getCell(1, 15).value = '[See rule 2(1)]';
    ws.getCell(2, 5).value = 'FORM A';
    ws.getCell(3, 5).value = 'FORMAT OF EMPLOYEE REGISTER';
    ws.getCell(4, 15).value = '[Part-A: For all Establishments]';
    ws.getCell(5, 1).value = 'Name and address of establishment in / under which contract is carried on';
    writeFormARajasthanTitleBandsInColumnsEG(ws, {});
    expect(String(ws.getCell(4, 15).value || '')).toBe('');
    expect(String(ws.getCell(3, 5).value || '')).toBe('FORMAT OF EMPLOYEE REGISTER');
    expect(String(ws.getCell(4, 1).value || '')).toMatch(/part-?a.*establish/i);
    expect(String(ws.getCell(4, 1).alignment?.horizontal || '')).toBe('left');
    expect(String(ws.getCell(5, 1).value || '')).toMatch(/establishment/i);
  });

  test('inserts Part-A row when next row already has establishment field', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('FORM A');
    ws.getCell(1, 1).value = '[See rule 2(1)]';
    ws.getCell(2, 1).value = 'FORM A';
    ws.getCell(3, 1).value = 'FORMAT OF EMPLOYEE REGISTER';
    ws.getCell(4, 1).value = 'Name and address of establishment in / under which contract is carried on';
    writeFormARajasthanTitleBandsInColumnsEG(ws, {});
    expect(String(ws.getCell(3, 5).value || '')).toBe('FORMAT OF EMPLOYEE REGISTER');
    expect(String(ws.getCell(4, 1).value || '')).toMatch(/part-?a.*establish/i);
    expect(String(ws.getCell(4, 1).alignment?.horizontal || '')).toBe('left');
    expect(String(ws.getCell(5, 1).value || '')).toMatch(/establishment/i);
  });

  test('fills original template columns without rewriting footer', async () => {
    const templateArrayBuffer = await buildBlankFormARjTemplate();
    const { blob, buffer } = await buildFormARajasthanWorkbookWithTemplateStyles({
      templateArrayBuffer,
      mappedData: [
        {
          'Sl. No.': 1,
          'Employee Code': 'E001',
          Name: 'Suryakanta',
          Surname: 'Jana',
          "Father's/Spouse Name": 'NIRMAL JANA',
          'Date of Birth': '10 Aug 1995',
          Nationality: 'Indian',
          'Education Level': 'DIPLOMA',
          'Date of Joining': '01 Jan 2020',
          Designation: 'Electrician',
          'Category (HS/S/SS/US)*': 'S',
          'Type of Employment': 'Permanent',
          Mobile: '9876543210',
          UAN: '100200300400',
          PAN: 'ABCDE1234F',
        },
      ],
      headersToUse: FORM_A_RJ_CANONICAL_TABLE_HEADERS,
      formFileName: 'Form_A_RJ_-_Rajasthan.xlsx',
      headerFormData: { 'Name of Establishment': 'Vayona Site' },
    });

    expect(blob).toBeTruthy();
    const outWb = new ExcelJS.Workbook();
    await outWb.xlsx.load(buffer);
    const ws = outWb.worksheets[0];

    expect(String(ws.getCell(1, 5).value || '')).toMatch(/see rule 2\(1\)/i);
    expect(String(ws.getCell(2, 5).value || '')).toBe('FORM A');
    expect(String(ws.getCell(3, 5).value || '')).toBe('FORMAT OF EMPLOYEE REGISTER');
    expect(String(ws.getCell(4, 1).value || '')).toMatch(/part-?a.*establish/i);
    expect(String(ws.getCell(4, 1).alignment?.horizontal || '')).toBe('left');
    expect(String(ws.getCell(8, 3).value || '')).toBe('Name');
    expect(String(ws.getCell(8, 2).value || '')).toBe('Employee Code');
    expect(Number(ws.getColumn(2).width || 0)).toBeLessThanOrEqual(12);
    expect(String(ws.getCell(9, 1).value)).toBe('1');
    expect(String(ws.getCell(9, 2).value)).toBe('E001');
    expect(String(ws.getCell(9, 3).value)).toBe('Suryakanta');
    expect(String(ws.getCell(9, 4).value)).toBe('Jana');
    expect(Number(ws.getRow(9).height || 0)).toBeGreaterThanOrEqual(30);
    expect(String(ws.getCell(20, 1).value || '')).toMatch(/Highly Skilled/i);
    expect(String(ws.getCell(9, 2).value || '')).not.toMatch(/Name of the workman/i);
  });

  test('Form A MH applies full box borders after PAN (ESIC IP onward)', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('FORM A');
    ws.getCell(1, 1).value = '[See rule 2(1)]';
    ws.getCell(2, 1).value = 'FORM A';
    ws.getCell(3, 1).value = 'FORMAT OF EMPLOYEE REGISTER';
    ws.getCell(4, 1).value = '[Part-A: For all Establishments]';
    ws.getCell(5, 1).value = 'Name of Establishment';
    FORM_A_RJ_CANONICAL_TABLE_HEADERS.forEach((h, i) => {
      ws.getCell(7, i + 1).value = h;
    });
    const mhExtra = ['ESIC IP', 'LWF', 'AADHAAR', 'Present Address', 'Permanent Address'];
    mhExtra.forEach((h, i) => {
      ws.getCell(7, 16 + i).value = h;
    });
    // Simulate MH template tall merge after PAN (no per-row boxes).
    ws.mergeCells(8, 16, 18, 20);
    for (let r = 8; r <= 18; r += 1) {
      for (let c = 1; c <= 15; c += 1) {
        ws.getCell(r, c).border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' },
        };
      }
    }
    ws.getCell(19, 1).value = '*(Highly Skilled/Skilled/Semi Skilled/Un Skilled)';
    const templateArrayBuffer = await wb.xlsx.writeBuffer();

    const { buffer } = await buildFormARajasthanWorkbookWithTemplateStyles({
      templateArrayBuffer,
      mappedData: [
        {
          'Sl. No.': 1,
          'Employee Code': 'E001',
          Name: 'A',
          Surname: 'B',
          PAN: 'ABCDE1234F',
        },
        {
          'Sl. No.': 2,
          'Employee Code': 'E002',
          Name: 'C',
          Surname: 'D',
          PAN: 'FGHIJ5678K',
        },
      ],
      headersToUse: FORM_A_RJ_CANONICAL_TABLE_HEADERS,
      formFileName: 'Form_A_MH_-_Maharashtra.xlsx',
    });

    const outWb = new ExcelJS.Workbook();
    await outWb.xlsx.load(buffer);
    const out = outWb.worksheets[0];
    const hasFullBox = (cell) => {
      const b = cell?.border || {};
      return !!(b.top && b.left && b.bottom && b.right);
    };
    // Template already had Part-A on its own row — header stays at 7, data at 8+.
    expect(String(out.getCell(8, 1).value)).toBe('1');
    expect(String(out.getCell(9, 1).value)).toBe('2');
    expect(hasFullBox(out.getCell(8, 15))).toBe(true); // PAN row 1
    expect(hasFullBox(out.getCell(8, 16))).toBe(true); // ESIC IP row 1
    expect(hasFullBox(out.getCell(8, 17))).toBe(true); // LWF row 1
    expect(hasFullBox(out.getCell(9, 16))).toBe(true); // ESIC IP row 2
    expect(hasFullBox(out.getCell(9, 18))).toBe(true); // AADHAAR row 2
  });

  test('writes Site Management contractor name and address like Form C_RJ', async () => {
    const headerFormData = applyFormARajasthanContractorFromSite(
      {},
      {
        contractorName: 'Vayona Contractor',
        contractorAddress: 'Jaipur',
        contractorState: 'Rajasthan',
      }
    );
    const parsedFormHeader = enrichFormARajasthanDisplayHeader({
      title: 'FORM A',
      subtitle: 'FORMAT OF EMPLOYEE REGISTER',
      fields: [{ label: 'Name and address of contractor', key: 'form_a_rj_contractor' }],
    });
    expect(headerFormData.form_a_rj_contractor).toBe('Vayona Contractor, Jaipur, Rajasthan');

    const { buffer } = await buildFormARajasthanWorkbookWithTemplateStyles({
      templateArrayBuffer: await buildBlankFormARjTemplate(),
      mappedData: [{ 'Employee Code': 'E001', Name: 'Suryakanta', Surname: 'Jana' }],
      headersToUse: FORM_A_RJ_CANONICAL_TABLE_HEADERS,
      headerFormData,
      parsedFormHeader,
    });
    const outWb = new ExcelJS.Workbook();
    await outWb.xlsx.load(buffer);
    const ws = outWb.worksheets[0];
    const contractorCell = String(ws.getCell(6, 1).value || '');
    expect(contractorCell.toLowerCase()).toContain('contractor');
    expect(contractorCell).toContain('Vayona Contractor');
    expect(contractorCell).toContain('Jaipur');
  });

  test('repairs hybrid template: Employee Code narrow, Name in column C, codes in B', async () => {
    const templateArrayBuffer = await buildHybridBrokenFormARjTemplate();
    const inWb = new ExcelJS.Workbook();
    await inWb.xlsx.load(templateArrayBuffer);
    expect(sheetBodyHasEmploymentCardLabels(inWb.worksheets[0], 8, 19)).toBe(true);

    const { buffer } = await buildFormARajasthanWorkbookWithTemplateStyles({
      templateArrayBuffer,
      mappedData: [
        {
          'Sl. No.': 1,
          'Employee Code': 'VE0705',
          Name: 'Suryakanta',
          Surname: 'Jana',
          "Father's/Spouse Name": 'NIRMAL JANA',
          'Date of Birth': '10 Aug 1995',
        },
        {
          'Sl. No.': 2,
          'Employee Code': 'VE0706',
          Name: 'Prem Shankar',
          Surname: 'Menaria',
          "Father's/Spouse Name": 'Madan lal menariy',
          'Date of Birth': '09 Sept 1999',
        },
      ],
      headersToUse: FORM_A_RJ_CANONICAL_TABLE_HEADERS,
      employeesOverride: [
        { EmployeeID: 'VE0705', FirstName: 'Suryakanta', LastName: 'Jana' },
        { EmployeeID: 'VE0706', FirstName: 'Prem Shankar', LastName: 'Menaria' },
      ],
    });

    const outWb = new ExcelJS.Workbook();
    await outWb.xlsx.load(buffer);
    const ws = outWb.worksheets[0];

    expect(String(ws.getCell(2, 5).value || '')).toBe('FORM A');
    expect(String(ws.getCell(4, 1).value || '')).toMatch(/part-?a.*establish/i);
    expect(String(ws.getCell(8, 2).value || '')).toBe('Employee Code');
    expect(String(ws.getCell(8, 3).value || '')).toBe('Name');
    expect(Number(ws.getColumn(2).width || 0)).toBeLessThanOrEqual(12);
    expect(String(ws.getCell(9, 2).value || '')).toBe('VE0705');
    expect(String(ws.getCell(9, 3).value || '')).toBe('Suryakanta');
    expect(String(ws.getCell(10, 2).value || '')).toBe('VE0706');
    expect(String(ws.getCell(10, 3).value || '')).toBe('Prem Shankar');
    expect(String(ws.getCell(9, 2).value || '')).not.toMatch(/Name of the workman/i);
    expect(String(ws.getCell(15, 2).value || '')).not.toMatch(/Remarks/i);
  });

  test('detects Form XIV employment card sheet pattern', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Card');
    ws.getCell(10, 2).value = 'Nature of work and location of work';
    ws.getCell(14, 2).value = 'Name of the workman';
    ws.getCell(23, 2).value = 'Tenure of employment';
    expect(sheetLooksLikeFormXIVEmploymentCard(wb.worksheets[0])).toBe(true);
  });

  test('repairs hybrid template with header on row 6', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('FORM A');
    ws.getCell(2, 1).value = 'FORM A';
    ws.getCell(3, 1).value = 'FORMAT OF EMPLOYEE REGISTER';
    ws.getCell(4, 1).value =
      'Name and address of establishment in / under which contract is carried on';
    ws.getCell(5, 1).value = 'Name and address of contractor';
    ['Sl. N', 'Employee Code', 'Nature of work and location of work', 'Surname'].forEach((h, i) => {
      ws.getCell(6, i + 1).value = h;
    });
    [
      'Name of the workman',
      'Serial number in the register of workmen employed',
      'Nature of employment / designation',
    ].forEach((label, i) => {
      ws.getCell(7 + i, 2).value = label;
    });
    ws.getColumn(2).width = 42;
    ws.getCell(15, 1).value = '*(Highly Skilled/Skilled/Semi Skilled/Un Skilled)';

    const { buffer } = await buildFormARajasthanWorkbookWithTemplateStyles({
      templateArrayBuffer: await wb.xlsx.writeBuffer(),
      mappedData: [{ 'Employee Code': 'VE0705', Name: 'Suryakanta', Surname: 'Jana' }],
      headersToUse: FORM_A_RJ_CANONICAL_TABLE_HEADERS,
      employeesOverride: [{ EmployeeID: 'VE0705', FirstName: 'Suryakanta', LastName: 'Jana' }],
    });

    const outWb = new ExcelJS.Workbook();
    await outWb.xlsx.load(buffer);
    const out = outWb.worksheets[0];
    expect(String(out.getCell(4, 1).value || '')).toMatch(/part-?a.*establish/i);
    expect(String(out.getCell(7, 3).value || '')).toBe('Name');
    expect(String(out.getCell(8, 2).value || '')).toBe('VE0705');
    expect(String(out.getCell(8, 3).value || '')).toBe('Suryakanta');
    expect(String(out.getCell(8, 2).value || '')).not.toMatch(/Name of the workman/i);
  });

  test('detects hybrid template headers without Name column', () => {
    const hybridHdrs = [
      'Sl. N',
      'Employee Code',
      'Nature of work and location of work',
      'Surname',
      "Father's/Spouse N",
    ];
    expect(headersIndicateFormARajasthanHybridTemplate(hybridHdrs)).toBe(true);
    expect(
      isFormARajasthanLikeExport(
        { title: 'FORM A', subtitle: 'FORMAT OF EMPLOYEE REGISTER' },
        { formName: 'Format of Employee Register', state: 'Rajasthan' },
        'FORM_A.xlsx',
        'Name of the workman Tenure of employment',
        hybridHdrs
      )
    ).toBe(true);
  });

  test('requires original template buffer', async () => {
    await expect(
      buildFormARajasthanWorkbookWithTemplateStyles({
        templateArrayBuffer: null,
        mappedData: [],
        headersToUse: FORM_A_RJ_CANONICAL_TABLE_HEADERS,
      })
    ).rejects.toThrow(/Original form template buffer is required/i);
  });

  test('modal grid values win over conflicting employeesOverride', async () => {
    const templateArrayBuffer = await buildHybridBrokenFormARjTemplate();
    const { buffer } = await buildFormARajasthanWorkbookWithTemplateStyles({
      templateArrayBuffer,
      mappedData: [
        {
          'Sl. No.': 1,
          'Employee Code': 'VE0705',
          Name: 'Suryakanta',
          Surname: 'Jana',
        },
      ],
      headersToUse: FORM_A_RJ_CANONICAL_TABLE_HEADERS,
      employeesOverride: [{ EmployeeID: 'VE107', FirstName: 'Wrong', LastName: 'Person' }],
    });

    const outWb = new ExcelJS.Workbook();
    await outWb.xlsx.load(buffer);
    const ws = outWb.worksheets[0];
    expect(String(ws.getCell(9, 2).value || '')).toBe('VE0705');
    expect(String(ws.getCell(9, 3).value || '')).toBe('Suryakanta');
  });

  test('parsedTableStartCol hint does not shift data away from column A', async () => {
    const templateArrayBuffer = await buildHybridBrokenFormARjTemplate();
    const { buffer } = await buildFormARajasthanWorkbookWithTemplateStyles({
      templateArrayBuffer,
      mappedData: [
        {
          'Sl. No.': 1,
          'Employee Code': 'VE0705',
          Name: 'Suryakanta',
          Surname: 'Jana',
        },
      ],
      headersToUse: FORM_A_RJ_CANONICAL_TABLE_HEADERS,
      parsedTableStartCol: 1,
      parsedHeaderRowIndex: 6,
      parsedDataStartIndex: 7,
    });

    const outWb = new ExcelJS.Workbook();
    await outWb.xlsx.load(buffer);
    const ws = outWb.worksheets[0];
    expect(String(ws.getCell(9, 1).value)).toBe('1');
    expect(String(ws.getCell(9, 2).value || '')).toBe('VE0705');
    expect(String(ws.getCell(9, 3).value || '')).toBe('Suryakanta');
    expect(String(ws.getCell(9, 4).value || '')).toBe('Jana');
  });

  test('writes Site Management contractor name and address like Form C_RJ', async () => {
    const headerFormData = applyFormARajasthanContractorFromSite(
      {},
      {
        contractorName: 'Vayona Contractor',
        contractorAddress: 'Jaipur',
        contractorState: 'Rajasthan',
      }
    );
    const parsedFormHeader = enrichFormARajasthanDisplayHeader({
      title: 'FORM A',
      subtitle: 'FORMAT OF EMPLOYEE REGISTER',
      fields: [{ label: 'Name and address of contractor', key: 'form_a_rj_contractor' }],
    });
    expect(headerFormData.form_a_rj_contractor).toBe('Vayona Contractor, Jaipur, Rajasthan');

    const { buffer } = await buildFormARajasthanWorkbookWithTemplateStyles({
      templateArrayBuffer: await buildBlankFormARjTemplate(),
      mappedData: [{ 'Employee Code': 'E001', Name: 'Suryakanta', Surname: 'Jana' }],
      headersToUse: FORM_A_RJ_CANONICAL_TABLE_HEADERS,
      headerFormData,
      parsedFormHeader,
    });
    const outWb = new ExcelJS.Workbook();
    await outWb.xlsx.load(buffer);
    const ws = outWb.worksheets[0];
    const contractorCell = String(ws.getCell(6, 1).value || '');
    expect(contractorCell.toLowerCase()).toContain('contractor');
    expect(contractorCell).toContain('Vayona Contractor');
    expect(contractorCell).toContain('Jaipur');
  });

  test('maps Name from hybrid Nature of work column in modal rows', async () => {
    const hybridHdrs = [
      'Sl. N',
      'Employee Code',
      'Nature of work and location of work',
      'Surname',
    ];
    expect(resolveFormARajasthanExportHeaders(hybridHdrs)).toEqual(FORM_A_RJ_CANONICAL_TABLE_HEADERS);

    const templateArrayBuffer = await buildHybridBrokenFormARjTemplate();
    const { buffer } = await buildFormARajasthanWorkbookWithTemplateStyles({
      templateArrayBuffer,
      mappedData: [
        {
          'Sl. N': 1,
          'Employee Code': 'VE0705',
          'Nature of work and location of work': 'Suryakanta',
          Surname: 'Jana',
        },
      ],
      headersToUse: hybridHdrs,
    });

    const outWb = new ExcelJS.Workbook();
    await outWb.xlsx.load(buffer);
    const ws = outWb.worksheets[0];
    expect(String(ws.getCell(9, 2).value || '')).toBe('VE0705');
    expect(String(ws.getCell(9, 3).value || '')).toBe('Suryakanta');
  });
});
