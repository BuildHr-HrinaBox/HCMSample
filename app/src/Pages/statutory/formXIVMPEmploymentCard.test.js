import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import {
  FORM_XIV_MP_CANONICAL_TABLE_HEADERS,
  buildFormXIVMPPerEmployeeDownload,
  buildFormXIVMPWorkbookWithTemplateStyles,
  detectFormXIVKarnatakaWorksheetLayout,
  detectFormXIVStackedWorkmanLayout,
  finalizeFormXIVMadhyaPradeshWorksheet,
  isFormXIVKarnatakaContext,
  resolveFormXIVMPWorkmanFieldPositions,
  resolveFormXIVVariant,
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
    expect(String(ws.getCell(14, 5).value || '')).toBe('');
    expect(ws.getCell(6, 2).alignment?.wrapText).toBe(false);
    expect(Number(ws.getColumn(2).width)).toBeGreaterThanOrEqual(90);
    expect(Number(ws.getColumn(3).width)).toBeGreaterThanOrEqual(28);
  });

  it('matches the official MP template: titles in F, labels in B, values in C', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Sheet1');
    for (let c = 1; c <= 12; c += 1) ws.getColumn(c).width = 8;
    ws.getCell(2, 6).value = 'FORM XIV';
    ws.getCell(3, 6).value = '(See rule 76)';
    ws.getCell(4, 6).value = 'Employment Card';
    ws.getCell(6, 2).value = 'Name and address of contractor';
    ws.getCell(6, 2).font = { bold: true };
    ws.getCell(8, 2).value = 'Name and address of establishment in / under which contract is carried on';
    ws.getCell(8, 2).font = { bold: true };
    ws.getCell(10, 2).value = 'Nature of work and location of work';
    ws.getCell(10, 2).font = { bold: true };
    ws.getCell(12, 2).value = 'Name and address of principal employer';
    ws.getCell(12, 2).font = { bold: true };
    const workman = [
      '1 Name of the workman',
      '2 Serial number in the register of workmen employed',
      '3 Nature of employment / designation',
      '4 Wage rate with particulars or unit, in case of piece of work',
      '5 Wage period',
      '6 Tenure of employment',
      '7 Remarks',
    ];
    workman.forEach((label, i) => {
      ws.getCell(14 + i * 2, 3).value = label;
    });

    writeFormXIVMPHeaderFieldsToWorksheet(
      ws,
      {
        form_xiv_mp_contractor: 'VAYONA ENERGY PRIVATE LIMITED',
        form_xiv_mp_establishment: 'Dhar Site',
        form_xiv_mp_nature_location: 'MP-Dhar',
        form_xiv_mp_principal_employer: 'Principal Employer',
      },
      { formXIVVariant: 'mp' }
    );
    const headers = [...FORM_XIV_MP_CANONICAL_TABLE_HEADERS];
    writeFormXIVMPWorkmanFieldsToWorksheet(
      ws,
      {
        [headers[0]]: 'Ajaya Patra',
        [headers[1]]: 'VE0705',
        [headers[2]]: 'Engineer',
        [headers[3]]: '41991',
        [headers[4]]: 'July 2026',
        [headers[5]]: 'From 11',
      },
      headers,
      null,
      { formXIVVariant: 'mp' }
    );

    expect(String(ws.getCell(2, 3).value || '')).toMatch(/^FORM XIV$/i);
    expect(String(ws.getCell(3, 3).value || '')).toMatch(/see rule 76/i);
    expect(String(ws.getCell(4, 3).value || '')).toMatch(/^Employment Card$/i);
    expect(String(ws.getCell(2, 6).value || '')).toBe('');
    expect(String(ws.getCell(6, 2).value || '')).toBe('Name and address of contractor');
    expect(String(ws.getCell(6, 3).value || '')).toMatch(/VAYONA ENERGY PRIVATE LIMITED/);
    expect(String(ws.getCell(6, 5).value || '')).toBe('');
    expect(String(ws.getCell(8, 2).value || '')).toMatch(
      /name and address of establishment in \/ under which contract is carried on/i
    );
    expect(String(ws.getCell(10, 2).value || '')).toBe('Nature of work and location of work');
    expect(String(ws.getCell(12, 2).value || '')).toBe('Name and address of principal employer');
    expect(String(ws.getCell(14, 2).value || '')).toMatch(/^1 Name of the workman/i);
    expect(String(ws.getCell(14, 3).value || '')).toBe('Ajaya Patra');
    expect(String(ws.getCell(14, 1).value || '')).toBe('');
    expect(String(ws.getCell(14, 5).value || '')).toBe('');
    expect(String(ws.getCell(16, 2).value || '')).toMatch(/^2 Serial number in the register of workmen employed/i);
    expect(String(ws.getCell(16, 3).value || '')).toBe('VE0705');
    expect(String(ws.getCell(20, 2).value || '')).toMatch(
      /^4 Wage rate with particulars or unit, in case of piece of work/i
    );
    expect(String(ws.getCell(20, 3).value || '')).toBe('41991');
    expect(ws.getCell(6, 2).alignment?.wrapText).toBe(false);
    expect(ws.getCell(14, 2).alignment?.wrapText).toBe(false);
    expect(ws.getCell(2, 3).font?.bold).toBe(false);
    expect(ws.getCell(3, 3).font?.bold).toBe(false);
    expect(ws.getCell(4, 3).font?.bold).toBe(false);
    expect(ws.getCell(6, 2).font?.bold).toBe(false);
    expect(ws.getCell(8, 2).font?.bold).toBe(false);
    expect(ws.getCell(10, 2).font?.bold).toBe(false);
    expect(ws.getCell(12, 2).font?.bold).toBe(false);
    expect(ws.getCell(14, 2).font?.bold).toBe(false);
    expect(ws.getCell(6, 3).font?.bold).toBe(false);
    expect(ws.getCell(14, 3).font?.bold).toBe(false);
    expect(Number(ws.getColumn(2).width)).toBeGreaterThanOrEqual(70);
    expect(Number(ws.getColumn(3).width)).toBeGreaterThanOrEqual(28);
  });

  it('rebuilds the clipped MP download: name in A and truncated C workman lines', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Sheet1');
    for (let c = 1; c <= 12; c += 1) ws.getColumn(c).width = 3;
    ws.getCell(2, 6).value = 'FORM XIV';
    ws.getCell(3, 6).value = '(See rule 76)';
    ws.getCell(4, 6).value = 'Employment Card';
    ws.getCell(6, 2).value = 'Name';
    ws.getCell(8, 2).value = 'Name';
    ws.getCell(10, 2).value = 'Nature of';
    ws.getCell(12, 2).value = 'Name';
    ws.getCell(14, 1).value = 'Aditya';
    ws.getCell(20, 3).value = '4 Wage';
    ws.getCell(20, 4).value = '41991';
    ws.getCell(22, 3).value = '5 Wage';
    ws.getCell(22, 4).value = 'July 2026';
    ws.getCell(24, 3).value = '6 Tenure';
    ws.getCell(24, 4).value = 'From 11';
    ws.getCell(26, 3).value = '7 Remarks';

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
    const headers = [...FORM_XIV_MP_CANONICAL_TABLE_HEADERS];
    writeFormXIVMPWorkmanFieldsToWorksheet(
      ws,
      {
        [headers[0]]: 'Aditya',
        [headers[1]]: 'VE0705',
        [headers[2]]: 'Engineer',
        [headers[3]]: '41991',
        [headers[4]]: 'July 2026',
        [headers[5]]: 'From 11',
      },
      headers,
      null,
      { formXIVVariant: 'mp' }
    );

    expect(String(ws.getCell(6, 2).value || '')).toBe('Name and address of contractor');
    expect(String(ws.getCell(10, 2).value || '')).toBe('Nature of work and location of work');
    expect(String(ws.getCell(14, 2).value || '')).toMatch(/^1 Name of the workman/i);
    expect(String(ws.getCell(14, 3).value || '')).toBe('Aditya');
    expect(String(ws.getCell(14, 1).value || '')).toBe('');
    expect(String(ws.getCell(14, 5).value || '')).toBe('');
    expect(String(ws.getCell(20, 2).value || '')).toMatch(
      /^4 Wage rate with particulars or unit, in case of piece of work/i
    );
    expect(String(ws.getCell(20, 3).value || '')).toBe('41991');
    expect(String(ws.getCell(22, 3).value || '')).toBe('July 2026');
    expect(detectFormXIVKarnatakaWorksheetLayout(ws)).toBe(false);
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
    expect(String(out.getCell(14, 5).value || '')).toBe('');
    expect(out.getCell(6, 2).alignment?.wrapText).not.toBe(true);
    expect(Number(out.getColumn(2).width)).toBeGreaterThanOrEqual(90);
    expect(Number(out.getColumn(3).width)).toBeGreaterThanOrEqual(28);
    expect(String(out.getCell(14, 3).value || '')).toBe('Chinmaya Kumar Swain');
  });

  it('rebuilds the downloaded MP screenshot: name in A, truncated B labels, values in C', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Sheet1');
    for (let c = 1; c <= 12; c += 1) ws.getColumn(c).width = 8;
    ws.getCell(2, 6).value = 'FORM XIV';
    ws.getCell(2, 6).font = { bold: true, underline: true };
    ws.getCell(3, 6).value = '(See rule 76)';
    ws.getCell(3, 6).font = { italic: true };
    ws.getCell(4, 6).value = 'Employment Card';
    ws.getCell(4, 6).font = { bold: true };
    ws.getCell(6, 2).value = 'Name';
    ws.getCell(8, 2).value = 'Name';
    ws.getCell(10, 2).value = 'Nature of';
    ws.getCell(12, 2).value = 'Name';
    ws.getCell(14, 1).value = 'Aditya';
    ws.getCell(20, 2).value = "4 Wage'";
    ws.getCell(20, 3).value = '72298';
    ws.getCell(22, 2).value = '5 Wage';
    ws.getCell(22, 3).value = 'July 2026';
    ws.getCell(24, 2).value = '6 Tenure';
    ws.getCell(24, 3).value = 'From 01';
    ws.getCell(26, 2).value = '7 Remarks';
    const templateBuffer = await wb.xlsx.writeBuffer();
    const headers = [...FORM_XIV_MP_CANONICAL_TABLE_HEADERS];
    const { blob, buffer } = await buildFormXIVMPWorkbookWithTemplateStyles({
      templateArrayBuffer: templateBuffer,
      mappedData: [
        {
          [headers[0]]: 'Aditya',
          [headers[1]]: 'VE1256',
          [headers[2]]: 'Engineer',
          [headers[3]]: '72298',
          [headers[4]]: 'July 2026',
          [headers[5]]: 'From 01',
        },
      ],
      headersToUse: headers,
      parsedFormHeader: { title: 'FORM XIV', formXIVVariant: 'mp', fields: [] },
      formFileName: 'Form_XIV_MP_Abaji_Bangar.xlsx',
      headerFormData: {
        form_xiv_mp_contractor: 'VAYON',
        form_xiv_mp_establishment: 'Abaji Bangar',
        form_xiv_mp_nature_location: 'MP-Dhar',
        form_xiv_mp_principal_employer: 'M/s',
      },
    });
    const outWb = new ExcelJS.Workbook();
    const outBuffer =
      buffer != null
        ? buffer
        : typeof blob.arrayBuffer === 'function'
          ? await blob.arrayBuffer()
          : blob;
    await outWb.xlsx.load(outBuffer);
    const out = outWb.worksheets[0];
    expect(detectFormXIVKarnatakaWorksheetLayout(out)).toBe(false);
    expect(String(out.getCell(2, 3).value || '')).toMatch(/^FORM XIV$/i);
    expect(String(out.getCell(3, 3).value || '')).toMatch(/see rule 76/i);
    expect(String(out.getCell(4, 3).value || '')).toMatch(/^Employment Card$/i);
    expect(String(out.getCell(2, 6).value || '')).toBe('');
    expect(String(out.getCell(6, 2).value || '')).toBe('Name and address of contractor');
    expect(String(out.getCell(8, 2).value || '')).toMatch(
      /name and address of establishment in \/ under which contract is carried on/i
    );
    expect(String(out.getCell(10, 2).value || '')).toBe('Nature of work and location of work');
    expect(String(out.getCell(12, 2).value || '')).toBe('Name and address of principal employer');
    expect(String(out.getCell(6, 3).value || '')).toBe('VAYON');
    expect(String(out.getCell(8, 3).value || '')).toBe('Abaji Bangar');
    expect(String(out.getCell(6, 5).value || '')).toBe('');
    expect(String(out.getCell(14, 1).value || '')).toBe('');
    expect(String(out.getCell(14, 2).value || '')).toMatch(/^1 Name of the workman/i);
    expect(String(out.getCell(14, 3).value || '')).toBe('Aditya');
    expect(String(out.getCell(14, 5).value || '')).toBe('');
    expect(String(out.getCell(16, 2).value || '')).toMatch(
      /^2 Serial number in the register of workmen employed/i
    );
    expect(String(out.getCell(16, 3).value || '')).toBe('VE1256');
    expect(String(out.getCell(20, 2).value || '')).toMatch(
      /^4 Wage rate with particulars or unit, in case of piece of work/i
    );
    expect(String(out.getCell(20, 3).value || '')).toBe('72298');
    expect(String(out.getCell(20, 5).value || '')).toBe('');
    expect(String(out.getCell(22, 3).value || '')).toBe('July 2026');
    expect(String(out.getCell(24, 3).value || '')).toBe('From 01');
    expect(out.getCell(2, 3).font?.bold).not.toBe(true);
    expect(out.getCell(4, 3).font?.bold).not.toBe(true);
    expect(out.getCell(6, 2).font?.bold).not.toBe(true);
    expect(out.getCell(8, 2).font?.bold).not.toBe(true);
    expect(out.getCell(10, 2).font?.bold).not.toBe(true);
    expect(out.getCell(12, 2).font?.bold).not.toBe(true);
    expect(out.getCell(6, 3).font?.bold).not.toBe(true);
    expect(out.getCell(14, 3).font?.bold).not.toBe(true);
    expect(Number(out.getColumn(2).width)).toBeGreaterThanOrEqual(90);
    expect(Number(out.getColumn(3).width)).toBeGreaterThanOrEqual(28);
  });
});

/** Official Karnataka Form XIV — titles in C, labels in B with dotted workman lines, signature in F. */
async function buildKarnatakaFormXIVTemplate() {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Sheet1');
  for (let c = 1; c <= 12; c += 1) {
    ws.getColumn(c).width = 8;
  }
  ws.getCell(2, 3).value = 'FORM XIV';
  ws.getCell(3, 3).value = '[See rule 76]';
  ws.getCell(4, 3).value = 'Employment Card';
  ws.getCell(6, 2).value = 'Name and address of contractor';
  ws.getCell(7, 2).value = 'Name and address of Establishment in/under which contract is carried on :';
  ws.getCell(8, 2).value = 'Nature and location of work:';
  ws.getCell(9, 2).value = 'Name and address of Principal Employer:';
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
    ws.getCell(11 + i, 2).value = label;
  });
  ws.getCell(19, 6).value = '............................';
  ws.getCell(20, 6).value = 'Signature of Contractor';
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
      { formXIVVariant: 'ka' }
    );

    const contractorLine = String(ws.getCell(6, 2).value || '');
    expect(contractorLine).toMatch(/^Name and address of contractor/i);
    expect(contractorLine).toMatch(/wrtyy\./);
    expect(String(ws.getCell(7, 2).value || '')).toMatch(
      /^Name and address of Establishment in\/under which contract is carried on[\s\S]*Sample/i
    );
    expect(String(ws.getCell(8, 2).value || '')).toMatch(/^Nature and location of work:[\s\S]*Karnataka/i);
    expect(String(ws.getCell(9, 2).value || '')).toMatch(
      /^Name and address of Principal Employer:[\s\S]*Delphi/i
    );
    expect(ws.getCell(6, 2).alignment?.horizontal).toBe('left');
    expect(ws.getCell(7, 2).alignment?.horizontal).toBe('left');
    expect(ws.getCell(8, 2).alignment?.horizontal).toBe('left');
    expect(ws.getCell(9, 2).alignment?.horizontal).toBe('left');
    expect(ws.getCell(6, 2).font?.bold).toBe(false);
    expect(ws.getCell(7, 2).font?.bold).toBe(false);
    expect(ws.getCell(8, 2).font?.bold).toBe(false);
    expect(ws.getCell(9, 2).font?.bold).toBe(false);
    expect(String(ws.getCell(2, 3).value || '')).toMatch(/^FORM XIV$/i);
    expect(String(ws.getCell(3, 3).value || '')).toMatch(/see rule 76/i);
    expect(String(ws.getCell(4, 3).value || '')).toMatch(/^Employment Card$/i);
    expect(String(ws.getCell(2, 2).value || '')).not.toMatch(/FORM XIV/i);
    expect(ws.getCell(2, 3).alignment?.horizontal).toBe('center');
    expect(ws.getCell(3, 3).alignment?.horizontal).toBe('center');
    expect(ws.getCell(4, 3).alignment?.horizontal).toBe('center');
    expect(String(ws.getCell(6, 6).value || '')).not.toMatch(/wrtyy/);
    expect(String(ws.getCell(7, 6).value || '')).not.toMatch(/Sample/);

    const headers = [...FORM_XIV_MP_CANONICAL_TABLE_HEADERS];
    const layout = resolveFormXIVMPWorkmanFieldPositions(ws, headers, { formXIVVariant: 'ka' });
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
      layout,
      { formXIVVariant: 'ka' }
    );

    expect(String(ws.getCell(11, 1).value || '')).toBe('');
    const workmanLine = String(ws.getCell(11, 2).value || '');
    expect(workmanLine).toMatch(/^1\s+Name of the Workman/i);
    expect(workmanLine).toMatch(/\.{4,}/);
    expect(workmanLine).toMatch(/Stalin T/);
    expect(String(ws.getCell(12, 1).value || '')).toBe('');
    expect(String(ws.getCell(12, 2).value || '')).toMatch(
      /^2\s+Serial No\. in the Register of workmen employed[\s\S]*\.{4,}[\s\S]*VE0814/i
    );
    expect(String(ws.getCell(13, 2).value || '')).toMatch(
      /^3\s+Nature of employment\/Designation[\s\S]*\.{4,}[\s\S]*Engineer/i
    );
    expect(String(ws.getCell(14, 2).value || '')).toMatch(
      /particulars of unit in case of piece-work[\s\S]*\.{4,}[\s\S]*69220/i
    );
    expect(ws.getCell(11, 2).alignment?.horizontal).toBe('left');
    expect(ws.getCell(11, 2).font?.bold).toBe(false);
    expect(ws.getCell(6, 2).font?.bold).toBe(false);
    expect(ws.getCell(6, 2).alignment?.horizontal).toBe('left');
    expect(Number(ws.getColumn(2).width)).toBeGreaterThanOrEqual(90);
    expect(String(ws.getCell(20, 6).value || '')).toMatch(/signature of contractor/i);
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
    ws.getCell(11, 1).value =
      'M/s Clean Wind Power Bableshwar Pvt Ltd, 33/220KV Substation, Survey No: 132, Karjol Village, Bijapur Dist - 586113 Karnataka, Bableshwar, Karnataka';
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
      headers,
      null,
      { formXIVVariant: 'ka' }
    );

    const contractorLine = String(ws.getCell(6, 2).value || '');
    expect(contractorLine).toMatch(/^Name and address of contractor/i);
    expect(contractorLine).toMatch(/VAYON/);
    expect(ws.getCell(6, 2).alignment?.horizontal).toBe('left');
    expect(ws.getCell(6, 2).font?.bold).toBe(false);
    expect(String(ws.getCell(6, 6).value || '')).not.toMatch(/VAYON/);
    expect(String(ws.getCell(11, 1).value || '')).toBe('');
    const workmanLine = String(ws.getCell(11, 2).value || '');
    expect(workmanLine).toMatch(/^1\s+Name of the Workman/i);
    expect(workmanLine).toMatch(/\.{4,}/);
    expect(workmanLine).toMatch(/Ashok/);
    expect(workmanLine).not.toMatch(/Clean Wind Power/i);
    expect(String(ws.getCell(12, 1).value || '')).toBe('');
    expect(String(ws.getCell(12, 2).value || '')).toMatch(/^2\s+Serial No\.[\s\S]*\.{4,}[\s\S]*VE0705/i);
    expect(ws.getCell(11, 2).alignment?.horizontal).toBe('left');
    expect(ws.getCell(11, 2).font?.bold).toBe(false);
    expect(String(ws.getCell(6, 2).value || '')).toMatch(/^Name and address of contractor/i);
    expect(String(ws.getCell(7, 2).value || '')).toMatch(
      /^Name and address of Establishment in\/under which contract is carried on/i
    );
    expect(String(ws.getCell(2, 3).value || '')).toMatch(/^FORM XIV$/i);
    expect(Number(ws.getColumn(2).width)).toBeGreaterThanOrEqual(90);
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

  it('restores full left-aligned labels when the draft clips them in narrow column A', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Sheet1');
    ws.getColumn(1).width = 3;
    for (let c = 2; c <= 12; c += 1) ws.getColumn(c).width = 8;
    ws.getCell(2, 4).value = 'FORM XIV';
    ws.getCell(3, 4).value = '[See rule 76]';
    ws.getCell(4, 4).value = 'Employment Card';
    const clipped = [
      [6, 'Name and address if contractor'],
      [7, 'f Establishment in/under which contract is carried on :'],
      [8, 'nd location of work:'],
      [9, 'ress of Principal Employer:'],
    ];
    clipped.forEach(([row, text]) => {
      const cell = ws.getCell(row, 1);
      cell.value = text;
      cell.alignment = { horizontal: 'right', vertical: 'middle', wrapText: false };
    });
    ws.getCell(11, 1).value = 'ey No: 132,';
    ws.getCell(16, 2).value = '............';
    ws.getCell(17, 2).value = 7;
    ws.getCell(17, 3).value = '...............................';
    ws.getCell(19, 6).value = '.............................';
    ws.getCell(20, 6).value = 'Signature of Contractor';

    writeFormXIVMPHeaderFieldsToWorksheet(
      ws,
      {
        form_xiv_mp_contractor: 'VAYONA ENERGY',
        form_xiv_mp_establishment: 'Site',
        form_xiv_mp_nature_location: 'KA-Bableshwar',
        form_xiv_mp_principal_employer: 'Clean Wind Power Bableshwar Pvt Ltd',
      },
      { formXIVVariant: 'ka' }
    );

    const contractor = String(ws.getCell(6, 1).value || '');
    expect(contractor).toMatch(/^Name and address of contractor/i);
    expect(contractor).toMatch(/VAYONA ENERGY/);
    expect(contractor).not.toMatch(/^f Establishment/i);
    expect(ws.getCell(6, 1).alignment?.horizontal).toBe('left');
    expect(Number(ws.getColumn(1).width)).toBeGreaterThanOrEqual(90);
    expect(String(ws.getCell(7, 1).value || '')).toMatch(
      /^Name and address of Establishment in\/under which contract is carried on/i
    );
    expect(String(ws.getCell(8, 1).value || '')).toMatch(/^Nature and location of work:/i);
    expect(String(ws.getCell(9, 1).value || '')).toMatch(/^Name and address of Principal Employer:/i);
    expect(String(ws.getCell(6, 1).value || '')).not.toMatch(/^ress of/i);
    expect(String(ws.getCell(20, 6).value || '')).toMatch(/signature of contractor/i);
  });

  it('puts the workman name on line 1 even when headers and leftover cells are scrambled', async () => {
    const { ws } = await buildKarnatakaFormXIVTemplate();
    ws.getCell(13, 4).value = 'Ashok Jangamashetti';
    const scrambledHeaders = [
      'Name and address of contractor',
      'Name and address of Establishment in/under which contract is carried on :',
      'Nature and location of work:',
      'Name and address of Principal Employer:',
      'Nature of employment / designation',
      'Serial number in the register of workmen employed',
      'Name of the workman',
      'Wage rate with particulars or unit, in case of piece of work',
      'Wage period',
      'Tenure of employment',
      'Remarks',
    ];
    writeFormXIVMPWorkmanFieldsToWorksheet(
      ws,
      {
        'Name of the workman': 'Ashok Jangamashetti',
        'Serial number in the register of workmen employed': 'VE0705',
        'Nature of employment / designation': 'Junior Engineer',
        'Wage rate with particulars or unit, in case of piece of work': '104761',
        'Wage period': 'April 2026',
        'Tenure of employment': 'From 01 Dec 2025',
      },
      scrambledHeaders,
      null,
      { formXIVVariant: 'ka' }
    );

    const nameLine = String(ws.getCell(11, 2).value || '');
    expect(nameLine).toMatch(/^1\s+Name of the Workman/i);
    expect(nameLine).toMatch(/Ashok Jangamashetti/);
    const serialLine = String(ws.getCell(12, 2).value || '');
    expect(serialLine).toMatch(/^2\s+Serial No\./i);
    expect(serialLine).toMatch(/VE0705/);
    const natureLine = String(ws.getCell(13, 2).value || '');
    expect(natureLine).toMatch(/^3\s+Nature of employment\/Designation/i);
    expect(natureLine).toMatch(/Junior Engineer/);
    expect(natureLine).not.toMatch(/Ashok Jangamashetti/);
    expect(String(ws.getCell(14, 2).value || '')).toMatch(/104761/);
    expect(String(ws.getCell(15, 2).value || '')).toMatch(/April 2026/);
    expect(String(ws.getCell(16, 2).value || '')).toMatch(/From 01 Dec 2025/);
  });
});

describe('Form XIV Karnataka ZIP variant and corresponding employees', () => {
  it('treats Form XIV - Karnataka as ka even when sheet text mentions Gujarat', () => {
    expect(
      isFormXIVKarnatakaContext(
        { title: 'FORM XIV' },
        { formName: 'Form XIV - Karnataka', state: 'Karnataka' },
        'Form XIV - Karnataka.xlsx',
        'Employment Card Central & Gujarat Rules'
      )
    ).toBe(true);
    expect(
      resolveFormXIVVariant(
        { title: 'FORM XIV' },
        { formName: 'Form XIV - Karnataka', state: 'Karnataka' },
        'Form XIV - Karnataka.xlsx',
        'Employment Card Central & Gujarat Rules'
      )
    ).toBe('ka');
  });

  it('still treats Madhya Pradesh Form XIV as mp', () => {
    expect(
      resolveFormXIVVariant(
        { title: 'FORM XIV' },
        { formName: 'Form XIV - Madhya Pradesh', state: 'Madhya Pradesh' },
        'Form XIV - Madhya Pradesh.xlsx',
        'Employment Card'
      )
    ).toBe('mp');
  });

  it('zips one Karnataka card per employee with official left-aligned labels', async () => {
    const { wb } = await buildKarnatakaFormXIVTemplate();
    const templateArrayBuffer = await wb.xlsx.writeBuffer();
    const headers = [...FORM_XIV_MP_CANONICAL_TABLE_HEADERS];
    const result = await buildFormXIVMPPerEmployeeDownload({
      templateArrayBuffer,
      mappedData: [
        { [headers[0]]: 'Aakash Tiwari', [headers[1]]: 'VE0701', [headers[5]]: 'From 01' },
        { [headers[0]]: 'Stalin T', [headers[1]]: 'VE0814', [headers[5]]: 'From 06' },
      ],
      headersToUse: headers,
      parsedFormHeader: { title: 'FORM XIV' },
      formFileName: 'Form XIV - Karnataka.xlsx',
      headerFormData: {
        form_xiv_mp_contractor: 'VAYON',
        form_xiv_mp_establishment: 'Sample',
      },
      rowItem: { formName: 'Form XIV - Karnataka', state: 'Karnataka' },
      sheetText: 'FORM XIV Employment Card Gujarat rules',
    });

    expect(result.fileName).toMatch(/Form_XIV_-_Karnataka_Employees\.zip$/i);
    const zip = await JSZip.loadAsync(result.blob);
    const names = Object.keys(zip.files).filter((name) => /\.xlsx$/i.test(name));
    expect(names).toHaveLength(2);
    expect(names.every((name) => name.startsWith('Form_XIV_KA_'))).toBe(true);
    expect(names.some((name) => /Aakash_Tiwari/i.test(name))).toBe(true);
    expect(names.some((name) => /Stalin_T/i.test(name))).toBe(true);

    const aakashEntry = names.find((name) => /Aakash_Tiwari/i.test(name));
    const bytes = await zip.file(aakashEntry).async('arraybuffer');
    const outWb = new ExcelJS.Workbook();
    await outWb.xlsx.load(bytes);
    const ws = outWb.worksheets[0];
    expect(String(ws.getCell(2, 3).value || '')).toMatch(/^FORM XIV$/i);
    expect(String(ws.getCell(6, 2).value || '')).toMatch(/^Name and address of contractor/i);
    expect(String(ws.getCell(6, 2).value || '')).toMatch(/VAYON/);
    expect(ws.getCell(6, 2).alignment?.horizontal).toBe('left');
    const workmanLine = String(ws.getCell(11, 2).value || '');
    expect(workmanLine).toMatch(/^1\s+Name of the Workman/i);
    expect(workmanLine).toMatch(/\.{4,}/);
    expect(workmanLine).toMatch(/Aakash Tiwari/);
    expect(String(ws.getCell(11, 1).value || '')).toBe('');
    expect(String(ws.getCell(16, 2).value || '')).toMatch(/^6\s+Tenure of employment/i);
    expect(String(ws.getCell(16, 2).value || '')).toMatch(/From 01/);
    expect(String(ws.getCell(20, 6).value || '')).toMatch(/signature of contractor/i);
  });
});

describe('Form XIV MP finalize must not rewrite Rajasthan Form A', () => {
  async function buildFormARajasthanRegisterSheet() {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('FORM A');
    ws.getCell(1, 1).value = '[See rule 2(1)]';
    ws.getCell(2, 1).value = 'FORM A';
    ws.getCell(3, 1).value = 'FORMAT OF EMPLOYEE REGISTER';
    ws.getCell(4, 1).value = 'Name and address of establishment in / under which contract is carried on';
    ws.getCell(5, 1).value = 'Name and address of contractor';
    ws.getCell(7, 1).value = 'Sl. No.';
    ws.getCell(7, 2).value = 'Employee Code';
    ws.getCell(7, 3).value = 'Nature of work and location of work';
    ws.getCell(7, 4).value = 'Surname';
    ws.getCell(7, 5).value = "Father's/Spouse Name";
    ws.getCell(7, 6).value = 'Date of Birth';
    const employees = [
      ['Suryakanta', 'Jana', 'NIRMAL JANA', '10 Aug 1995'],
      ['Prem Shankar', 'Menaria', 'Madan lal menariy', '09 Sept 1999'],
    ];
    employees.forEach((emp, i) => {
      const r = 8 + i;
      ws.getCell(r, 1).value = i + 1;
      ws.getCell(r, 2).value = `E${1001 + i}`;
      ws.getCell(r, 3).value = emp[0];
      ws.getCell(r, 4).value = emp[1];
      ws.getCell(r, 5).value = emp[2];
      ws.getCell(r, 6).value = emp[3];
    });
    return ws;
  }

  it('does not stamp Employment Card labels into Form A employee-code column', async () => {
    const ws = await buildFormARajasthanRegisterSheet();
    const applied = finalizeFormXIVMadhyaPradeshWorksheet(ws);
    expect(applied).toBe(false);
    expect(String(ws.getCell(8, 2).value || '')).toBe('E1001');
    expect(String(ws.getCell(9, 2).value || '')).toBe('E1002');
    expect(String(ws.getCell(8, 3).value || '')).toBe('Suryakanta');
    expect(String(ws.getCell(8, 4).value || '')).toBe('Jana');
    expect(String(ws.getCell(2, 1).value || '')).toMatch(/^FORM A$/i);
  });
});
