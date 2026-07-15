import ExcelJS from 'exceljs';
import fs from 'fs';
import path from 'path';
import {
  FORM_I_TAMIL_NADU_NIL_DEFAULT,
  FORM_I_TN_DEFAULT_EMPLOYEES,
  applyFormITamilNaduNilDefaultsToRows,
  buildFormITamilNaduDefaultRows,
  cloneFormITamilNaduWorkmenWorksheetClean,
  ensureFormITamilNaduDefaultEmployeeRows,
  ensureFormITamilNaduWorkmenTitleLayout,
  isFormITamilNaduFinesOrWorkmenDefaultContext,
  isFormITamilNaduNilDefaultHeader,
  isFormITamilNaduSkipAutofillHeader,
  isFormITamilNaduSuspensionWorkbookContext,
  normalizeFormITamilNaduHeaderText,
  safeFormIExcelJsCellText,
  sheetLooksLikeFormITamilNaduWorkmenRegister
} from './formITamilNadu';

describe('Form I Tamil Nadu NIL defaults', () => {
  const nilHeaders = [
    'Nature of offence committed and date of offence',
    'Date of suspension',
    'Date of revocation of suspension',
    'Rate at which subsistence allowance calculated and period for which calculation made',
    'Amount of subsistence allowance paid and the date of payment',
    'Whether the employee had been exonerated or awarded any punishment',
    'Remarks',
    'Signature of employee with date for receiving money or postal acknowledgement of money order'
  ];

  const emolumentsHeader = 'Monthly emoluments (Wages) paid to the employee';

  it('normalizes header text and recognizes the NIL columns', () => {
    expect(normalizeFormITamilNaduHeaderText('  Date of Suspension  ')).toBe('date of suspension');
    nilHeaders.forEach((header) => {
      expect(isFormITamilNaduNilDefaultHeader(header)).toBe(true);
    });
    expect(
      isFormITamilNaduNilDefaultHeader(
        'Whether the employee had been exoncrated or awaerded any punishment'
      )
    ).toBe(true);
    expect(
      isFormITamilNaduNilDefaultHeader(
        'Department in which the employee was working last and his designation'
      )
    ).toBe(false);
    expect(
      isFormITamilNaduNilDefaultHeader('Name and Address of the Employee kept under suspension')
    ).toBe(false);
    expect(isFormITamilNaduNilDefaultHeader(emolumentsHeader)).toBe(false);
    expect(isFormITamilNaduSkipAutofillHeader(emolumentsHeader)).toBe(true);
  });

  it('detects the suspension workbook from filename and headers', () => {
    expect(
      isFormITamilNaduSuspensionWorkbookContext({
        fileName: 'Form_I_-_TamilNadu.xlsx',
        headers: [emolumentsHeader, ...nilHeaders]
      })
    ).toBe(true);
  });

  it('clears Monthly emoluments and fills blank NIL columns without overwriting real values', () => {
    const headers = [emolumentsHeader, ...nilHeaders];
    const rows = [
      {
        [emolumentsHeader]: '51',
        [nilHeaders[0]]: '',
        [nilHeaders[1]]: 'Enter Date of suspension',
        [nilHeaders[2]]: '2026-07-15',
        [nilHeaders[3]]: '   ',
        [nilHeaders[4]]: '',
        [nilHeaders[5]]: 'Awarded warning',
        [nilHeaders[6]]: '',
        [nilHeaders[7]]: 'Enter Signature'
      }
    ];

    expect(applyFormITamilNaduNilDefaultsToRows(rows, headers)).toEqual([
      {
        [emolumentsHeader]: '',
        [nilHeaders[0]]: FORM_I_TAMIL_NADU_NIL_DEFAULT,
        [nilHeaders[1]]: FORM_I_TAMIL_NADU_NIL_DEFAULT,
        [nilHeaders[2]]: '2026-07-15',
        [nilHeaders[3]]: FORM_I_TAMIL_NADU_NIL_DEFAULT,
        [nilHeaders[4]]: FORM_I_TAMIL_NADU_NIL_DEFAULT,
        [nilHeaders[5]]: 'Awarded warning',
        [nilHeaders[6]]: FORM_I_TAMIL_NADU_NIL_DEFAULT,
        [nilHeaders[7]]: FORM_I_TAMIL_NADU_NIL_DEFAULT
      }
    ]);
  });

  it('can force NIL overwrite on suspension columns', () => {
    const headers = [nilHeaders[0]];
    const rows = [{ [nilHeaders[0]]: 'Some offence' }];
    expect(applyFormITamilNaduNilDefaultsToRows(rows, headers, { overwriteNil: true })).toEqual([
      { [nilHeaders[0]]: FORM_I_TAMIL_NADU_NIL_DEFAULT }
    ]);
  });
});

describe('Form I Tamil Nadu Form 1.xlsx default employees', () => {
  const finesHeaders = [
    'Sl.No',
    'Name',
    "Father's/ Husband's Name or Workshop Departmental or Gang Number",
    'Department of Gang',
    'Act or omission for which fine imposed',
    'Whether workman showed cause against fine or not and if so, date on which cause was shown',
    'Total wages for the wage-period in which fine imposed',
    'Amount of and Date on which fine imposed',
    'Date on which fine realised'
  ];

  it('detects Register of Fines Tamil Nadu context for defaults', () => {
    expect(
      isFormITamilNaduFinesOrWorkmenDefaultContext({
        fileName: 'Form_I_-_TamilNadu.xlsx',
        formHeader: { title: 'FORM I', subtitle: 'REGISTER OF FINES' },
        headers: finesHeaders
      })
    ).toBe(true);
  });

  it('builds default Register of Fines rows from Form 1 Tamilnadu.xlsx employees', () => {
    const rows = buildFormITamilNaduDefaultRows(finesHeaders);
    expect(rows).toHaveLength(FORM_I_TN_DEFAULT_EMPLOYEES.length);
    expect(rows[0].Name).toBe('Avudaiappan');
    expect(rows[0]['Department of Gang']).toBe('Service');
    expect(rows[1].Name).toBe('Vijayakumar');
    expect(rows[4].Name).toBe('Vinu Monikandan');
    expect(rows[0]['Act or omission for which fine imposed']).toBe('');
  });

  it('seeds empty Sl.No-only grids with Form 1 defaults on download', () => {
    const templateRows = [
      { 'Sl.No': '1', Name: '' },
      { 'Sl.No': '2', Name: '' }
    ];
    const seeded = ensureFormITamilNaduDefaultEmployeeRows(templateRows, finesHeaders);
    expect(seeded).toHaveLength(5);
    expect(seeded.map((r) => r.Name)).toEqual([
      'Avudaiappan',
      'Vijayakumar',
      'Ajikumar',
      'Senthilkannan',
      'Vinu Monikandan'
    ]);
  });

  it('keeps existing named rows instead of overwriting', () => {
    const existing = [{ 'Sl.No': '1', Name: 'Someone Else', 'Department of Gang': 'Ops' }];
    expect(ensureFormITamilNaduDefaultEmployeeRows(existing, finesHeaders)).toEqual(existing);
  });
});

describe('Form I Tamil Nadu Register of Workmen clean workbook', () => {
  it('does not treat suspension Form_I_-_TamilNadu as workmen defaults', () => {
    expect(
      isFormITamilNaduFinesOrWorkmenDefaultContext({
        fileName: 'Form_I_-_TamilNadu.xlsx',
        headers: [
          'Date of suspension',
          'Nature of offence committed and date of offence',
          'Amount of subsistence allowance paid and the date of payment'
        ]
      })
    ).toBe(false);
    expect(
      isFormITamilNaduSuspensionWorkbookContext({
        fileName: 'Form_I_-_TamilNadu.xlsx',
        headers: ['Date of suspension', 'Nature of offence committed and date of offence']
      })
    ).toBe(true);
  });

  it('reads merge-slave cell text safely when the master value is null', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('FORM 1');
    ws.getCell(1, 1).value = 'FORM - I\nREGISTER OF WORKMEN';
    ws.mergeCells('A1:J1');
    ws.getCell(1, 1).value = null;
    expect(() => ws.getCell(1, 2).text).toThrow(/toString/);
    expect(safeFormIExcelJsCellText(ws.getCell(1, 2))).toBe('');
    expect(safeFormIExcelJsCellText(ws.getCell(1, 1))).toBe('');
  });

  it('detects the FORM 1 workmen sheet and rebuilds a clean A-J layout', async () => {
    const templatePath = path.join(__dirname, '../../Form 1 Tamilnadu.xlsx');
    if (!fs.existsSync(templatePath)) return;

    const srcWb = new ExcelJS.Workbook();
    await srcWb.xlsx.load(fs.readFileSync(templatePath));
    const form1 = srcWb.getWorksheet('FORM 1');
    expect(form1).toBeTruthy();
    expect(sheetLooksLikeFormITamilNaduWorkmenRegister(form1)).toBe(true);
    expect(sheetLooksLikeFormITamilNaduWorkmenRegister(srcWb.getWorksheet('FORM U'))).toBe(false);

    const { workbook, worksheet } = cloneFormITamilNaduWorkmenWorksheetClean(form1);
    expect(workbook.worksheets.length).toBe(1);
    expect(worksheet.name).toBe('FORM 1');
    expect(worksheet.model.merges || []).toEqual(
      expect.arrayContaining(['A1:J1', 'A2:D2', 'E2:J2', 'A3:A4', 'J3:J4'])
    );
    expect(worksheet.getColumn(3).width).toBeGreaterThan(20);
    expect(worksheet.getCell(1, 1).alignment?.horizontal).toBe('center');
    expect(worksheet.getCell(1, 1).alignment?.wrapText).toBe(true);
    expect(String(worksheet.getCell(1, 1).value)).toMatch(/FORM\s*-\s*I/i);
    expect(String(worksheet.getCell(1, 1).value)).toMatch(/REGISTER OF WORKMEN/i);

    ensureFormITamilNaduWorkmenTitleLayout(worksheet);
    const out = await workbook.xlsx.writeBuffer();
    const reload = new ExcelJS.Workbook();
    await reload.xlsx.load(out);
    expect(reload.worksheets[0].model.merges || []).toEqual(
      expect.arrayContaining(['A1:J1', 'A3:A4'])
    );
    expect(String(reload.worksheets[0].getCell(1, 1).value)).toMatch(/REGISTER OF WORKMEN/i);
  });
});
