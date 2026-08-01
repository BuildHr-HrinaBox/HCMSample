import ExcelJS from 'exceljs';
import fs from 'fs';
import path from 'path';
import {
  FORM_I_TAMIL_NADU_NIL_DEFAULT,
  FORM_I_TN_DEFAULT_EMPLOYEES,
  FORM_I_TN_SUSPENSION_NIL_OF_MONTH_TEXT,
  applyFormITamilNaduNilDefaultsToRows,
  applyFormITamilNaduSuspensionNilTableRows,
  buildFormITamilNaduDefaultRows,
  buildFormITamilNaduRegisterOfFinesWorkbookClean,
  buildFormITamilNaduSuspensionNilTableRows,
  cloneFormITamilNaduWorkmenWorksheetClean,
  FORM_I_TN_FINES_CANONICAL_HEADERS,
  ensureFormITamilNaduDefaultEmployeeRows,
  ensureFormITamilNaduWorkmenTitleLayout,
  findFormITamilNaduSuspensionNilMonthYearHeader,
  findFormITamilNaduSuspensionNilPrimaryHeader,
  formITamilNaduWorkmenDownloadHasSubstantiveRows,
  formITamilNaduWorkmenRowHasSubstantiveEntry,
  isFormITamilNaduAmountAllowancePaidHeader,
  isFormITamilNaduFinesNilDefaultHeader,
  isFormITamilNaduFinesOrWorkmenDefaultContext,
  isFormITamilNaduNilDefaultHeader,
  isFormITamilNaduSkipAutofillHeader,
  isFormITamilNaduSuspensionBlankDefaultHeader,
  isFormITamilNaduSuspensionWorkbookContext,
  normalizeFormITamilNaduHeaderText,
  resolveFormITamilNaduWorkmenExportCellValue,
  safeFormIExcelJsCellText,
  sheetLooksLikeFormITamilNaduWorkmenRegister
} from './formITamilNadu';

describe('Form I Tamil Nadu NIL defaults', () => {
  const blankHeaders = [
    'Nature of offence committed and date of offence',
    'Date of suspension',
    'Date of revocation of suspension',
    'Rate at which subsistence allowance calculated and period for which calculation made',
    'Amount of subsistence allowance paid and the date of payment',
    'Whether the employee had been exonerated or awarded any punishment',
    'Signature of employee with date for receiving money or postal acknowledgement of money order'
  ];
  const nilHeaders = ['Remarks'];
  const rateHeader =
    'Rate at which subsistence allowance calculated and period for which calculation made';
  const rateHeaderTypo =
    'Rate ar which subsustence allowance calculated and period for which calculation made';

  const emolumentsHeader = 'Monthly emoluments (Wages) paid to the employee';
  const nameHeader = 'Name and Address of the Employee kept under suspension';
  const amountPaidHeader = 'Amount of subsistence allowance paid and the date of payment';
  const suspensionHeaders = [
    'Sl.No',
    nameHeader,
    emolumentsHeader,
    amountPaidHeader,
    ...blankHeaders.filter((h) => h !== amountPaidHeader),
    ...nilHeaders
  ];

  it('normalizes header text and recognizes blank vs NIL columns', () => {
    expect(normalizeFormITamilNaduHeaderText('  Date of Suspension  ')).toBe('date of suspension');
    blankHeaders.forEach((header) => {
      expect(isFormITamilNaduSuspensionBlankDefaultHeader(header)).toBe(true);
      expect(isFormITamilNaduNilDefaultHeader(header)).toBe(false);
    });
    nilHeaders.forEach((header) => {
      expect(isFormITamilNaduNilDefaultHeader(header)).toBe(true);
      expect(isFormITamilNaduSuspensionBlankDefaultHeader(header)).toBe(false);
    });
    expect(
      isFormITamilNaduSuspensionBlankDefaultHeader(
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
    expect(isFormITamilNaduAmountAllowancePaidHeader(amountPaidHeader)).toBe(true);
    expect(isFormITamilNaduSuspensionBlankDefaultHeader(amountPaidHeader)).toBe(true);
    expect(isFormITamilNaduSuspensionBlankDefaultHeader(rateHeader)).toBe(true);
    expect(isFormITamilNaduSuspensionBlankDefaultHeader(rateHeaderTypo)).toBe(true);
  });

  it('detects the suspension workbook from filename and headers', () => {
    expect(
      isFormITamilNaduSuspensionWorkbookContext({
        fileName: 'Form_I_-_TamilNadu.xlsx',
        headers: [emolumentsHeader, ...blankHeaders, ...nilHeaders]
      })
    ).toBe(true);
  });

  it('detects the suspension workbook from SA Form 1 sheet naming', () => {
    expect(
      isFormITamilNaduSuspensionWorkbookContext({
        fileName: 'Form I_TN - TamilNadu.xlsx',
        sheetText: 'SA Form 1 Register of Employees Placed under suspension'
      })
    ).toBe(true);
  });

  it('builds a single Nill of the month row with blank offence/date/amount columns', () => {
    expect(findFormITamilNaduSuspensionNilPrimaryHeader(suspensionHeaders)).toBe(nameHeader);
    expect(findFormITamilNaduSuspensionNilMonthYearHeader(suspensionHeaders, nameHeader)).toBe(
      emolumentsHeader
    );
    const row = buildFormITamilNaduSuspensionNilTableRows(
      suspensionHeaders,
      FORM_I_TN_SUSPENSION_NIL_OF_MONTH_TEXT,
      'Jul 2026'
    )[0];
    expect(row[nameHeader]).toBe(FORM_I_TN_SUSPENSION_NIL_OF_MONTH_TEXT);
    expect(row[emolumentsHeader]).toBe('Jul 2026');
    expect(row[amountPaidHeader]).toBe('');
    expect(row['Date of suspension']).toBe('');
    expect(row['Nature of offence committed and date of offence']).toBe('');
    expect(row['Date of revocation of suspension']).toBe('');
    expect(row['Whether the employee had been exonerated or awarded any punishment']).toBe('');
    expect(
      row[
        'Signature of employee with date for receiving money or postal acknowledgement of money order'
      ]
    ).toBe('');
    expect(
      row['Rate at which subsistence allowance calculated and period for which calculation made']
    ).toBe('');
    expect(row.Remarks).toBe(FORM_I_TAMIL_NADU_NIL_DEFAULT);
  });

  it('forces Amount of subsistence allowance paid blank and drops footer-like rows', () => {
    const peopleRows = [
      {
        [nameHeader]: 'rajeshkumar ramasamy',
        [emolumentsHeader]: '',
        [amountPaidHeader]: '31-05-2026',
        [blankHeaders[0]]: 'NIL'
      },
      {
        [nameHeader]: '',
        [amountPaidHeader]: 'For (Company Name)'
      },
      {
        [amountPaidHeader]: 'Authorised Signatory'
      }
    ];
    const applied = applyFormITamilNaduSuspensionNilTableRows(suspensionHeaders, peopleRows, {
      nilPrimaryText: FORM_I_TN_SUSPENSION_NIL_OF_MONTH_TEXT,
      monthYearLabel: 'Jul 2026',
      force: true
    });
    expect(applied).toHaveLength(1);
    expect(applied[0][nameHeader]).toBe(FORM_I_TN_SUSPENSION_NIL_OF_MONTH_TEXT);
    expect(applied[0][amountPaidHeader]).toBe('');
    expect(applied[0][emolumentsHeader]).toBe('Jul 2026');
  });

  it('overwrites fetched payment dates in Amount of subsistence allowance paid', () => {
    const headers = [amountPaidHeader, emolumentsHeader];
    const rows = [{ [amountPaidHeader]: '31-05-2026', [emolumentsHeader]: '5000' }];
    expect(applyFormITamilNaduNilDefaultsToRows(rows, headers)).toEqual([
      { [amountPaidHeader]: '', [emolumentsHeader]: '' }
    ]);
  });

  it('clears Monthly emoluments, blanks offence/date/rate/amount columns, keeps NIL for remarks', () => {
    const remarksHeader = nilHeaders[0];
    const headers = [emolumentsHeader, ...blankHeaders, remarksHeader];
    const rows = [
      {
        [emolumentsHeader]: '51',
        [blankHeaders[0]]: '',
        [blankHeaders[1]]: 'Enter Date of suspension',
        [blankHeaders[2]]: '2026-07-15',
        [rateHeader]: '   ',
        [blankHeaders[4]]: '31-05-2026',
        [blankHeaders[5]]: 'Awarded warning',
        [blankHeaders[6]]: 'Enter Signature',
        [remarksHeader]: ''
      }
    ];

    expect(applyFormITamilNaduNilDefaultsToRows(rows, headers)).toEqual([
      {
        [emolumentsHeader]: '',
        [blankHeaders[0]]: '',
        [blankHeaders[1]]: '',
        [blankHeaders[2]]: '2026-07-15',
        [rateHeader]: '',
        [blankHeaders[4]]: '',
        [blankHeaders[5]]: 'Awarded warning',
        [blankHeaders[6]]: '',
        [remarksHeader]: FORM_I_TAMIL_NADU_NIL_DEFAULT
      }
    ]);
  });

  it('can force blank overwrite on suspension offence columns', () => {
    const headers = [blankHeaders[0]];
    const rows = [{ [blankHeaders[0]]: 'Some offence' }];
    expect(applyFormITamilNaduNilDefaultsToRows(rows, headers, { overwriteNil: true })).toEqual([
      { [blankHeaders[0]]: '' }
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
    'Date on which fine realised',
    'Remarks'
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

  it('builds a canonical Register of Fines workbook when workmen template was linked', () => {
    const { worksheet } = buildFormITamilNaduRegisterOfFinesWorkbookClean({
      establishmentName: 'Acme',
      establishmentAddress: 'Chennai'
    });
    expect(String(worksheet.name)).toMatch(/pw\s*form\s*i/i);
    const title = String(worksheet.getCell(1, 1).value || '');
    expect(title).toMatch(/register\s+of\s+fines/i);
    expect(title).not.toMatch(/register\s+of\s+workmen/i);
    expect(String(worksheet.getCell(2, 1).value || '')).toMatch(/Acme/);
    FORM_I_TN_FINES_CANONICAL_HEADERS.forEach((h, i) => {
      expect(String(worksheet.getCell(3, i + 1).value || '')).toBe(h);
    });
  });

  it('detects Register of Workmen defaults even when filename is form-draft.xlsx', () => {
    expect(
      isFormITamilNaduFinesOrWorkmenDefaultContext({
        fileName: 'form-draft.xlsx',
        formHeader: { title: 'Form 1 Register of Workmen', subtitle: 'REGISTER OF WORKMEN' },
        headers: [
          'S No',
          'Emp ID',
          'Name and Address of the workman',
          'Date on which he completed 480 days of service'
        ]
      })
    ).toBe(true);
  });

  it('recognizes Register of Fines NIL columns', () => {
    finesHeaders.slice(4).forEach((header) => {
      expect(isFormITamilNaduFinesNilDefaultHeader(header)).toBe(true);
    });
    expect(isFormITamilNaduFinesNilDefaultHeader('Name')).toBe(false);
    expect(isFormITamilNaduFinesNilDefaultHeader('Department of Gang')).toBe(false);
  });

  it('builds default Register of Fines rows from Form 1 Tamilnadu.xlsx employees', () => {
    const rows = buildFormITamilNaduDefaultRows(finesHeaders);
    expect(rows).toHaveLength(FORM_I_TN_DEFAULT_EMPLOYEES.length);
    expect(rows[0].Name).toBe('Avudaiappan');
    expect(rows[0]['Department of Gang']).toBe('Service');
    expect(rows[1].Name).toBe('Vijayakumar');
    expect(rows[4].Name).toBe('Vinu Monikandan');
    expect(rows[0]['Act or omission for which fine imposed']).toBe(FORM_I_TAMIL_NADU_NIL_DEFAULT);
    expect(rows[0]['Whether workman showed cause against fine or not and if so, date on which cause was shown']).toBe(
      FORM_I_TAMIL_NADU_NIL_DEFAULT
    );
    expect(rows[0]['Total wages for the wage-period in which fine imposed']).toBe(
      FORM_I_TAMIL_NADU_NIL_DEFAULT
    );
    expect(rows[0]['Amount of and Date on which fine imposed']).toBe(FORM_I_TAMIL_NADU_NIL_DEFAULT);
    expect(rows[0]['Date on which fine realised']).toBe(FORM_I_TAMIL_NADU_NIL_DEFAULT);
    expect(rows[0].Remarks).toBe(FORM_I_TAMIL_NADU_NIL_DEFAULT);
  });

  it('applies NIL to blank Register of Fines columns without overwriting names', () => {
    const rows = [
      {
        Name: 'Avudaiappan',
        'Act or omission for which fine imposed': '',
        'Whether workman showed cause against fine or not and if so, date on which cause was shown':
          'Enter Whether',
        'Total wages for the wage-period in which fine imposed': '',
        'Amount of and Date on which fine imposed': '',
        'Date on which fine realised': '',
        Remarks: ''
      }
    ];
    expect(applyFormITamilNaduNilDefaultsToRows(rows, finesHeaders)).toEqual([
      {
        Name: 'Avudaiappan',
        'Act or omission for which fine imposed': FORM_I_TAMIL_NADU_NIL_DEFAULT,
        'Whether workman showed cause against fine or not and if so, date on which cause was shown':
          FORM_I_TAMIL_NADU_NIL_DEFAULT,
        'Total wages for the wage-period in which fine imposed': FORM_I_TAMIL_NADU_NIL_DEFAULT,
        'Amount of and Date on which fine imposed': FORM_I_TAMIL_NADU_NIL_DEFAULT,
        'Date on which fine realised': FORM_I_TAMIL_NADU_NIL_DEFAULT,
        Remarks: FORM_I_TAMIL_NADU_NIL_DEFAULT
      }
    ]);
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

describe('Form I Tamil Nadu Register of Workmen download helpers', () => {
  const workmenHeaders = [
    'S No',
    'Emp ID',
    'Name and Address of the workman',
    'Designation of the workmen',
    'Whether Temporary, casual, Badli, or Apprentice (other than those covered under the Apprentices Act, 1961)',
    'Date of first entry into service',
    'Date on which he completed 480 days of service',
    'Date on which made permanent',
    'Remarks',
    'Signature of the workman with date (to attest the entries)'
  ];

  it('detects Sl.No-only rows as non-substantive', () => {
    const snoOnly = [
      { 'S No': '1', 'Emp ID': '', 'Name and Address of the workman': '' },
      { 'S No': '2', 'Emp ID': '', 'Name and Address of the workman': '' }
    ];
    expect(formITamilNaduWorkmenRowHasSubstantiveEntry(snoOnly[0], workmenHeaders)).toBe(false);
    expect(formITamilNaduWorkmenDownloadHasSubstantiveRows(snoOnly, workmenHeaders)).toBe(false);
  });

  it('resolves export fallbacks for Sl.No-only rows', () => {
    const row = { 'S No': '1', 'Emp ID': '', 'Name and Address of the workman': '' };
    const emp = FORM_I_TN_DEFAULT_EMPLOYEES[0];
    expect(
      resolveFormITamilNaduWorkmenExportCellValue(row, 'Emp ID', 1, emp, 0, workmenHeaders)
    ).toBe('VE0147');
    expect(
      resolveFormITamilNaduWorkmenExportCellValue(
        row,
        'Name and Address of the workman',
        2,
        emp,
        0,
        workmenHeaders
      )
    ).toBe('Avudaiappan');
    expect(
      resolveFormITamilNaduWorkmenExportCellValue(
        row,
        'Designation of the workmen',
        3,
        emp,
        0,
        workmenHeaders
      )
    ).toBe('Assistant Manager');
  });

  it('forces default employees onto Sl.No-only grids for download', () => {
    const templateRows = [{ 'S No': '1' }, { 'S No': '2' }];
    const seeded = ensureFormITamilNaduDefaultEmployeeRows(templateRows, workmenHeaders, { force: true });
    expect(seeded).toHaveLength(5);
    expect(seeded[0]['Name and Address of the workman']).toBe('Avudaiappan');
    expect(seeded[0]['Emp ID']).toBe('VE0147');
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
