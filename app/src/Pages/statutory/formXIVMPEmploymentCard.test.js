import ExcelJS from 'exceljs';
import {
  FORM_XIV_MP_CANONICAL_TABLE_HEADERS,
  buildFormXIVMPWorkbookWithTemplateStyles,
  detectFormXIVKarnatakaWorksheetLayout,
  detectFormXIVStackedWorkmanLayout,
  resolveFormXIVMPWorkmanFieldPositions,
  writeFormXIVMPHeaderFieldsToWorksheet,
  writeFormXIVMPWorkmanFieldsToWorksheet,
} from './formXIVMPEmploymentCard';

/** Mimic MP Employment Card: ordinals in A, labels in B, values intended for E. */
async function buildSplitOrdinalStackedTemplate() {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Sheet1');
  for (let c = 1; c <= 12; c += 1) {
    ws.getColumn(c).width = 3;
  }
  ws.getCell(3, 6).value = '(See rule 76)';
  ws.getCell(6, 2).value = 'Name and Address of the Establishment in/under which Contract is carried on:';
  ws.getCell(8, 2).value = 'Name and Address of the Contractor';
  ws.getCell(10, 2).value = 'Nature of work and location of work';
  ws.getCell(12, 2).value = 'Name and address of Principal Employer';
  ws.getCell(14, 1).value = 1;
  ws.getCell(14, 2).value = 'Name of the workman';
  ws.getCell(15, 1).value = 2;
  ws.getCell(15, 2).value = 'Serial number in the register of workmen employed';
  ws.getCell(17, 1).value = 3;
  ws.getCell(17, 2).value = 'Nature of employment / designation';
  ws.getCell(19, 1).value = 4;
  ws.getCell(19, 2).value = "Wage rate with particulars or unit, in case of piece of work";
  ws.getCell(21, 1).value = 5;
  ws.getCell(21, 2).value = 'Wage period';
  ws.getCell(23, 1).value = 6;
  ws.getCell(23, 2).value = 'Tenure of employment';
  ws.getCell(25, 1).value = 7;
  ws.getCell(25, 2).value = 'Remarks';
  return { wb, ws };
}

describe('Form XIV MP stacked download alignment', () => {
  it('detects stacked layout when ordinals are split from labels', async () => {
    const { ws } = await buildSplitOrdinalStackedTemplate();
    expect(detectFormXIVStackedWorkmanLayout(ws)).toBe(true);
  });

  it('writes header and workman values into column E, not beside the label', async () => {
    const { ws } = await buildSplitOrdinalStackedTemplate();
    writeFormXIVMPHeaderFieldsToWorksheet(
      ws,
      {
        form_xiv_mp_establishment: 'VAYONA',
        form_xiv_mp_contractor: 'M/s Contractor',
        form_xiv_mp_nature_location: 'MP-Dhar',
        form_xiv_mp_principal_employer: 'Principal Employer',
      },
      { formXIVVariant: 'mp' }
    );

    expect(String(ws.getCell(6, 5).value || '')).toBe('VAYONA');
    expect(String(ws.getCell(6, 2).value || '')).toMatch(/establishment/i);
    expect(String(ws.getCell(8, 5).value || '')).toBe('M/s Contractor');
    expect(String(ws.getCell(10, 5).value || '')).toBe('MP-Dhar');
    expect(String(ws.getCell(12, 5).value || '')).toBe('Principal Employer');

    const headers = [...FORM_XIV_MP_CANONICAL_TABLE_HEADERS];
    const layout = resolveFormXIVMPWorkmanFieldPositions(ws, headers);
    expect(layout.stackedLayout).toBe(true);
    expect(layout.valueCol).toBe(5);

    writeFormXIVMPWorkmanFieldsToWorksheet(
      ws,
      {
        [headers[0]]: 'Jeevan',
        [headers[1]]: 'VE1256',
        [headers[2]]: 'Engineer',
      },
      headers,
      layout
    );

    expect(String(ws.getCell(14, 5).value || '')).toBe('Jeevan');
    expect(String(ws.getCell(15, 5).value || '')).toBe('VE1256');
    expect(String(ws.getCell(17, 5).value || '')).toBe('Engineer');
    // Labels stay in column B on a single unwrapped line (original MP template).
    expect(String(ws.getCell(14, 2).value || '')).toMatch(/name of the workman/i);
    expect(String(ws.getCell(14, 2).value || '')).not.toMatch(/\n/);
    expect(ws.getCell(14, 2).alignment?.wrapText).toBe(false);
    expect(String(ws.getCell(15, 2).value || '')).toMatch(/serial number in the register of workmen employed/i);
    expect(ws.getCell(15, 2).alignment?.wrapText).toBe(false);
    expect(String(ws.getCell(6, 2).value || '')).toMatch(
      /name and address of establishment in \/ under which contract is carried on/i
    );
    expect(ws.getCell(6, 2).alignment?.wrapText).toBe(false);
    expect(Number(ws.getColumn(2).width)).toBeGreaterThanOrEqual(70);
    expect(Number(ws.getColumn(5).width)).toBeGreaterThanOrEqual(18);
  });

  it('does not treat the MP split-ordinal template as Karnataka', async () => {
    const { ws } = await buildSplitOrdinalStackedTemplate();
    expect(detectFormXIVKarnatakaWorksheetLayout(ws)).toBe(false);
  });

  it('shows full original-template labels on one line even when wrap was clipping them', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Sheet1');
    for (let c = 1; c <= 12; c += 1) ws.getColumn(c).width = 8;
    ws.getCell(2, 6).value = 'FORM XIV';
    ws.getCell(3, 6).value = '(See rule 76)';
    ws.getCell(4, 6).value = 'Employment Card';
    const labels = [
      [6, 'Name and address of contractor'],
      [8, 'Name and address of establishment in / under which contract is carried on'],
      [10, 'Nature of work and location of work'],
      [12, 'Name and address of principal employer'],
    ];
    labels.forEach(([row, text]) => {
      const cell = ws.getCell(row, 2);
      cell.value = text;
      cell.alignment = { wrapText: true, vertical: 'top', horizontal: 'left' };
      ws.getRow(row).height = 15;
    });
    ws.getCell(14, 1).value = 1;
    ws.getCell(14, 2).value = 'Name of the workman';
    ws.getCell(14, 2).alignment = { wrapText: true };
    ws.getCell(16, 1).value = 2;
    ws.getCell(16, 2).value = 'Serial number in the register of workmen employed';
    ws.getCell(16, 2).alignment = { wrapText: true };
    ws.getCell(18, 1).value = 3;
    ws.getCell(18, 2).value = 'Nature of employment / designation';
    ws.getCell(20, 1).value = 4;
    ws.getCell(20, 2).value = "Wage' rate with particulars or unit, in case of piece of work";

    writeFormXIVMPHeaderFieldsToWorksheet(
      ws,
      {
        form_xiv_mp_contractor: 'VAYON',
        form_xiv_mp_establishment: 'Dhar',
        form_xiv_mp_nature_location: 'MP-Dhar',
        form_xiv_mp_principal_employer: 'M/s',
      },
      { formXIVVariant: 'mp' }
    );

    expect(String(ws.getCell(6, 2).value || '')).toBe('Name and address of contractor');
    expect(String(ws.getCell(8, 2).value || '')).toMatch(
      /name and address of establishment in \/ under which contract is carried on/i
    );
    expect(String(ws.getCell(8, 2).value || '')).not.toMatch(/\n/);
    expect(ws.getCell(6, 2).alignment?.wrapText).toBe(false);
    expect(ws.getCell(8, 2).alignment?.wrapText).toBe(false);
    expect(Number(ws.getColumn(2).width)).toBeGreaterThanOrEqual(70);
    expect(String(ws.getCell(6, 5).value || '')).toBe('VAYON');
    expect(String(ws.getCell(8, 5).value || '')).toBe('Dhar');
    expect(detectFormXIVKarnatakaWorksheetLayout(ws)).toBe(false);
  });

  it('rebuilds truncated draft labels like the downloaded MP employment card screenshot', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Sheet1');
    for (let c = 1; c <= 12; c += 1) ws.getColumn(c).width = 8;
    ws.getCell(2, 6).value = 'FORM XIV';
    ws.getCell(3, 6).value = '(See rule 76)';
    ws.getCell(4, 6).value = 'Employment Card';
    const truncated = [
      [6, 'Name'],
      [8, 'Name'],
      [10, 'Nature'],
      [12, 'Name'],
      [14, '1 Name of'],
      [16, '2 Serial'],
      [18, '3 Nature'],
      [20, "4 Wage'"],
    ];
    truncated.forEach(([row, text]) => {
      const cell = ws.getCell(row, 2);
      cell.value = text;
      cell.alignment = { wrapText: true, vertical: 'top', horizontal: 'left' };
      ws.getRow(row).height = 15;
    });
    ws.getCell(6, 4).value = 'VAYON';
    ws.getCell(8, 4).value = 'Dhar';
    ws.getCell(10, 4).value = 'MP-Dhar';
    ws.getCell(12, 4).value = 'M/s';
    ws.getCell(14, 4).value = 'Chinmaya';
    ws.getCell(16, 4).value = 'VE0782';
    ws.getCell(18, 4).value = 'Junior';
    ws.getCell(20, 4).value = '113708';

    expect(detectFormXIVStackedWorkmanLayout(ws)).toBe(false);
    expect(detectFormXIVKarnatakaWorksheetLayout(ws)).toBe(false);

    const headers = [...FORM_XIV_MP_CANONICAL_TABLE_HEADERS];
    writeFormXIVMPHeaderFieldsToWorksheet(
      ws,
      {
        form_xiv_mp_contractor: 'VAYON',
        form_xiv_mp_establishment: 'Dhar',
        form_xiv_mp_nature_location: 'MP-Dhar',
        form_xiv_mp_principal_employer: 'M/s',
      },
      { formXIVVariant: 'mp' }
    );
    writeFormXIVMPWorkmanFieldsToWorksheet(
      ws,
      {
        [headers[0]]: 'Chinmaya Kumar Swain',
        [headers[1]]: 'VE0782',
        [headers[2]]: 'Junior Engineer',
        [headers[3]]: '113708',
      },
      headers
    );

    expect(String(ws.getCell(6, 2).value || '')).toBe('Name and address of contractor');
    expect(String(ws.getCell(8, 2).value || '')).toMatch(
      /name and address of establishment in \/ under which contract is carried on/i
    );
    expect(String(ws.getCell(10, 2).value || '')).toBe('Nature of work and location of work');
    expect(String(ws.getCell(12, 2).value || '')).toBe('Name and address of principal employer');
    expect(String(ws.getCell(14, 2).value || '')).toMatch(
      /1 Name of the workman/i
    );
    expect(String(ws.getCell(16, 2).value || '')).toMatch(
      /2 Serial number in the register of workmen employed/i
    );
    expect(String(ws.getCell(18, 2).value || '')).toMatch(
      /3 Nature of employment \/ designation/i
    );
    expect(String(ws.getCell(20, 2).value || '')).toMatch(
      /4 Wage rate with particulars or unit, in case of piece of work/i
    );
    expect(ws.getCell(6, 2).alignment?.wrapText).toBe(false);
    expect(Number(ws.getColumn(2).width)).toBeGreaterThanOrEqual(90);
  });

  it('exports full labels through buildFormXIVMPWorkbookWithTemplateStyles', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Sheet1');
    for (let c = 1; c <= 12; c += 1) ws.getColumn(c).width = 8;
    ws.getCell(2, 6).value = 'FORM XIV';
    ws.getCell(3, 6).value = '(See rule 76)';
    ws.getCell(4, 6).value = 'Employment Card';
    const truncated = [
      [6, 'Name'],
      [8, 'Name'],
      [10, 'Nature'],
      [12, 'Name'],
      [14, '1 Name of'],
      [16, '2 Serial'],
      [18, '3 Nature'],
      [20, "4 Wage'"],
    ];
    truncated.forEach(([row, text]) => {
      ws.getCell(row, 2).value = text;
      ws.getCell(row, 2).alignment = { wrapText: true, vertical: 'top', horizontal: 'left' };
    });
    [[6, 'VAYON'], [8, 'Dhar'], [10, 'MP-Dhar'], [12, 'M/s'], [14, 'Chinmaya'], [16, 'VE0782'], [18, 'Junior'], [20, '113708']].forEach(
      ([row, value]) => {
        ws.getCell(row, 4).value = value;
      }
    );
    const templateBuffer = await wb.xlsx.writeBuffer();
    const headers = [...FORM_XIV_MP_CANONICAL_TABLE_HEADERS];
    const { blob } = await buildFormXIVMPWorkbookWithTemplateStyles({
      templateArrayBuffer: templateBuffer,
      mappedData: [
        {
          [headers[0]]: 'Chinmaya Kumar Swain',
          [headers[1]]: 'VE0782',
          [headers[2]]: 'Junior Engineer',
          [headers[3]]: '113708',
        },
      ],
      headersToUse: headers,
      parsedFormHeader: { title: 'FORM XIV', fields: [] },
      formFileName: 'Form_XIV_MP_Chinmaya_Kumar_Swain.xlsx',
      headerFormData: {
        form_xiv_mp_contractor: 'VAYON',
        form_xiv_mp_establishment: 'Dhar',
        form_xiv_mp_nature_location: 'MP-Dhar',
        form_xiv_mp_principal_employer: 'M/s',
      },
    });
    const outWb = new ExcelJS.Workbook();
    const outBuffer =
      typeof blob.arrayBuffer === 'function'
        ? await blob.arrayBuffer()
        : blob;
    await outWb.xlsx.load(outBuffer);
    const out = outWb.worksheets[0];
    expect(String(out.getCell(6, 2).value || '')).toBe('Name and address of contractor');
    expect(String(out.getCell(8, 2).value || '')).toMatch(
      /name and address of establishment in \/ under which contract is carried on/i
    );
    expect(String(out.getCell(14, 2).value || '')).toMatch(/1 Name of the workman/i);
    expect(String(out.getCell(16, 2).value || '')).toMatch(
      /2 Serial number in the register of workmen employed/i
    );
    expect(out.getCell(6, 2).alignment?.wrapText).not.toBe(true);
    expect(Number(out.getColumn(2).width)).toBeGreaterThanOrEqual(90);
    expect(String(out.getCell(14, 5).value || '')).toBe('Chinmaya Kumar Swain');
  });
});

/** Official Karnataka Form XIV — labels in B/C with dotted leaders, values belong in F. */
async function buildKarnatakaFormXIVTemplate() {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Sheet1');
  for (let c = 1; c <= 12; c += 1) {
    ws.getColumn(c).width = 8;
  }
  ws.getCell(3, 3).value = 'FORM XIV';
  ws.getCell(4, 3).value = '[See rule 76]';
  ws.getCell(5, 3).value = 'Employment Card';
  ws.getCell(7, 2).value = 'Name and address if contractor';
  ws.getCell(8, 2).value = 'Name and address of Establishment in/under which contract is carried on :';
  ws.getCell(9, 2).value = 'Nature and location of work:';
  ws.getCell(10, 2).value = 'Name and address of Principal Employer:';
  const workman = [
    '1 Name of the Workman ........................................................................',
    '2 Serial No. in the Register of workmen employed ........................................................',
    '3 Nature of employment/Designation ........................................................',
    '4 Wage rate (with particulars of unit in case of piece-work) ........................................................',
    '5 Wage period........................................................',
    '6 Tenure of employment........................................................',
    '7 Remarks........................................................',
  ];
  workman.forEach((label, i) => {
    ws.getCell(12 + i, 3).value = label;
  });
  ws.getCell(21, 6).value = '............................';
  ws.getCell(22, 6).value = 'Signature of Contractor';
  return { wb, ws };
}

describe('Form XIV Karnataka download alignment', () => {
  it('detects the official Karnataka stacked template', async () => {
    const { ws } = await buildKarnatakaFormXIVTemplate();
    expect(detectFormXIVKarnatakaWorksheetLayout(ws)).toBe(true);
    expect(detectFormXIVStackedWorkmanLayout(ws)).toBe(true);
  });

  it('keeps full official labels on one line and places values on the dotted fill-in', async () => {
    const { ws } = await buildKarnatakaFormXIVTemplate();
    writeFormXIVMPHeaderFieldsToWorksheet(
      ws,
      {
        form_xiv_mp_contractor: 'wrtyy.',
        form_xiv_mp_establishment: 'Sample',
        form_xiv_mp_nature_location: 'Karnataka',
        form_xiv_mp_principal_employer: 'Delphi,',
      },
      { formXIVVariant: 'mp' }
    );

    const contractorLine = String(ws.getCell(7, 2).value || '');
    expect(contractorLine).toMatch(/^Name and address of contractor/i);
    expect(contractorLine).toMatch(/wrtyy\./);
    expect(String(ws.getCell(8, 2).value || '')).toMatch(
      /^Name and address of Establishment in\/under which contract is carried on[\s\S]*Sample/i
    );
    expect(String(ws.getCell(9, 2).value || '')).toMatch(/^Nature and location of work:[\s\S]*Karnataka/i);
    expect(String(ws.getCell(10, 2).value || '')).toMatch(
      /^Name and address of Principal Employer:[\s\S]*Delphi/i
    );
    expect(ws.getCell(7, 2).alignment?.horizontal).toBe('left');
    expect(ws.getCell(8, 2).alignment?.horizontal).toBe('left');
    expect(ws.getCell(9, 2).alignment?.horizontal).toBe('left');
    expect(ws.getCell(10, 2).alignment?.horizontal).toBe('left');
    expect(ws.getCell(7, 2).font?.bold).toBe(false);
    expect(ws.getCell(8, 2).font?.bold).toBe(false);
    expect(ws.getCell(9, 2).font?.bold).toBe(false);
    expect(ws.getCell(10, 2).font?.bold).toBe(false);
    expect(String(ws.getCell(3, 3).value || '')).toMatch(/^FORM XIV$/i);
    expect(String(ws.getCell(4, 3).value || '')).toMatch(/see rule 76/i);
    expect(String(ws.getCell(5, 3).value || '')).toMatch(/^Employment Card$/i);
    expect(String(ws.getCell(3, 2).value || '')).not.toMatch(/FORM XIV/i);
    expect(ws.getCell(3, 3).alignment?.horizontal).toBe('center');
    expect(ws.getCell(4, 3).alignment?.horizontal).toBe('center');
    expect(ws.getCell(5, 3).alignment?.horizontal).toBe('center');
    expect(String(ws.getCell(7, 6).value || '')).not.toMatch(/wrtyy/);
    expect(String(ws.getCell(8, 6).value || '')).not.toMatch(/Sample/);

    const headers = [...FORM_XIV_MP_CANONICAL_TABLE_HEADERS];
    const layout = resolveFormXIVMPWorkmanFieldPositions(ws, headers);
    expect(layout.stackedLayout).toBe(true);
    expect(layout.kaInline).toBe(true);

    writeFormXIVMPWorkmanFieldsToWorksheet(
      ws,
      {
        [headers[0]]: 'Stalin T',
        [headers[1]]: 'VE0814',
        [headers[2]]: 'Engineer',
        [headers[3]]: '69220',
        [headers[4]]: 'May 2026',
        [headers[5]]: 'From 06',
      },
      headers,
      layout
    );

    expect(Number(ws.getCell(12, 1).value)).toBe(1);
    const workmanLine = String(ws.getCell(12, 2).value || '');
    expect(workmanLine).toMatch(/^Name of the Workman/i);
    expect(workmanLine).not.toMatch(/^1\s/);
    expect(workmanLine).toMatch(/\.{4,}/);
    expect(workmanLine).toMatch(/Stalin T/);
    expect(Number(ws.getCell(13, 1).value)).toBe(2);
    expect(String(ws.getCell(13, 2).value || '')).toMatch(
      /Serial No\. in the Register of workmen employed[\s\S]*\.{4,}[\s\S]*VE0814/i
    );
    expect(String(ws.getCell(14, 2).value || '')).toMatch(
      /Nature of employment\/Designation[\s\S]*\.{4,}[\s\S]*Engineer/i
    );
    expect(String(ws.getCell(15, 2).value || '')).toMatch(
      /particulars of unit in case of piece-work[\s\S]*\.{4,}[\s\S]*69220/i
    );
    expect(ws.getCell(12, 2).alignment?.horizontal).toBe('left');
    expect(ws.getCell(12, 2).font?.bold).toBe(false);
    expect(ws.getCell(7, 2).font?.bold).toBe(false);
    expect(ws.getCell(7, 2).alignment?.horizontal).toBe('left');
    expect(String(ws.getCell(22, 6).value || '')).toMatch(/signature of contractor/i);
  });

  it('rebuilds a saved draft that has truncated labels and values in column D', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Sheet1');
    for (let c = 1; c <= 12; c += 1) ws.getColumn(c).width = 8;
    ws.getCell(2, 4).value = 'FORM XIV';
    ws.getCell(3, 4).value = '[See rule 76]';
    ws.getCell(4, 4).value = 'Employment Card';
    ws.getCell(6, 2).value = 'Name';
    ws.getCell(6, 4).value = 'VAYON';
    ws.getCell(7, 2).value = 'Name';
    ws.getCell(7, 4).value = 'Babalesh';
    ws.getCell(8, 2).value = 'Nature';
    ws.getCell(8, 4).value = 'KA-';
    ws.getCell(9, 2).value = 'Name';
    ws.getCell(9, 4).value = 'M/s';
    ws.getCell(11, 2).value = '1 Name of';
    ws.getCell(11, 4).value = 'Ashok';
    ws.getCell(12, 2).value = '2 Serial No. in the Regis';
    ws.getCell(12, 4).value = 'VE0705';
    ws.getCell(13, 2).value = '3 Nature of employmen';
    ws.getCell(13, 4).value = 'Junior';
    ws.getCell(14, 2).value = '4 Wage rate (with partic';
    ws.getCell(14, 4).value = '104761';
    ws.getCell(15, 2).value = '5 Wage';
    ws.getCell(15, 4).value = 'April';
    ws.getCell(16, 2).value = '6 Tenure';
    ws.getCell(16, 4).value = 'From 01';
    ws.getCell(17, 2).value = '7 Remarks';
    ws.getCell(19, 6).value = '........................';
    ws.getCell(20, 6).value = 'Signature of Contractor';

    expect(detectFormXIVKarnatakaWorksheetLayout(ws)).toBe(true);

    writeFormXIVMPHeaderFieldsToWorksheet(
      ws,
      {
        form_xiv_mp_contractor: 'VAYON',
        form_xiv_mp_establishment: 'Babalesh',
        form_xiv_mp_nature_location: 'KA-',
        form_xiv_mp_principal_employer: 'M/s',
      },
      { formXIVVariant: 'ka' }
    );
    const headers = [...FORM_XIV_MP_CANONICAL_TABLE_HEADERS];
    writeFormXIVMPWorkmanFieldsToWorksheet(
      ws,
      {
        [headers[0]]: 'Ashok',
        [headers[1]]: 'VE0705',
        [headers[2]]: 'Junior',
        [headers[3]]: '104761',
        [headers[4]]: 'April',
        [headers[5]]: 'From 01',
      },
      headers
    );

    const contractorLine = String(ws.getCell(6, 2).value || '');
    expect(contractorLine).toMatch(/^Name and address of contractor/i);
    expect(contractorLine).toMatch(/VAYON/);
    expect(ws.getCell(6, 2).alignment?.horizontal).toBe('left');
    expect(ws.getCell(6, 2).font?.bold).toBe(false);
    expect(String(ws.getCell(6, 6).value || '')).not.toMatch(/VAYON/);
    expect(Number(ws.getCell(11, 1).value)).toBe(1);
    const workmanLine = String(ws.getCell(11, 2).value || '');
    expect(workmanLine).toMatch(/^Name of the Workman/i);
    expect(workmanLine).toMatch(/\.{4,}/);
    expect(workmanLine).toMatch(/Ashok/);
    expect(Number(ws.getCell(12, 1).value)).toBe(2);
    expect(String(ws.getCell(12, 2).value || '')).toMatch(/Serial No\.[\s\S]*\.{4,}[\s\S]*VE0705/i);
    expect(ws.getCell(11, 2).alignment?.horizontal).toBe('left');
    expect(ws.getCell(11, 2).font?.bold).toBe(false);
    expect(String(ws.getCell(6, 2).value || '')).toMatch(/^Name and address of contractor/i);
    expect(String(ws.getCell(7, 2).value || '')).toMatch(
      /^Name and address of Establishment in\/under which contract is carried on/i
    );
  });

  it('clears draft bold and center merge on the four header lines', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Sheet1');
    for (let c = 1; c <= 12; c += 1) ws.getColumn(c).width = 8;
    ws.getCell(2, 3).value = 'FORM XIV';
    ws.getCell(3, 3).value = '[See rule 76]';
    ws.getCell(4, 3).value = 'Employment Card';
    const headerLines = [
      [5, 'Name and address of contractor .......................... VAYONA ENERGY'],
      [7, 'Name and address of Establishment in/under which contract is carried on : .......................... Site'],
      [8, 'Nature and location of work: .......................... KA-Bableshwar'],
      [9, 'Name and address of Principal Employer: .......................... Clean Wind'],
    ];
    headerLines.forEach(([row, text]) => {
      const cell = ws.getCell(row, 2);
      cell.value = text;
      cell.font = { name: 'Times New Roman', size: 12, bold: true };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      ws.mergeCells(row, 2, row, 6);
    });
    ws.getCell(11, 1).value = 1;
    ws.getCell(11, 2).value = '1 Name of the Workman ........................ Issac';
    ws.getCell(12, 1).value = 2;
    ws.getCell(12, 2).value = '2 Serial No. in the Register of workmen employed ........................ VE0209';
    ws.getCell(13, 1).value = 3;
    ws.getCell(13, 2).value = '3 Nature of employment/Designation ........................ Assistant';
    ws.getCell(20, 6).value = 'Signature of Contractor';

    writeFormXIVMPHeaderFieldsToWorksheet(
      ws,
      {
        form_xiv_mp_contractor: 'VAYONA ENERGY',
        form_xiv_mp_establishment: 'Site',
        form_xiv_mp_nature_location: 'KA-Bableshwar',
        form_xiv_mp_principal_employer: 'Clean Wind',
      },
      { formXIVVariant: 'ka' }
    );

    [5, 7, 8, 9].forEach((row) => {
      expect(ws.getCell(row, 2).font?.bold).toBe(false);
      expect(ws.getCell(row, 2).alignment?.horizontal).toBe('left');
      expect(ws.getCell(row, 2).isMerged).toBe(false);
    });
    expect(String(ws.getCell(5, 2).value || '')).toMatch(/^Name and address of contractor/i);
    expect(String(ws.getCell(5, 2).value || '')).toMatch(/VAYONA ENERGY/);
  });
});
