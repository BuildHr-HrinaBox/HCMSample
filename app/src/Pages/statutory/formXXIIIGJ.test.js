import {
  FORM_XXIII_GJ_OT_NIL,
  applyFormXXIIIGJOtNilToMappedRows,
  applyFormXXIIIGJOtTemplateLayoutToExcelJs,
  isFormXXIIIGJContext,
  isFormXXIIIGJOtNilHeader,
  isFormXXIIIGJOtPreambleFieldLabel,
  isFormXXIIIGJSexHeader,
  prepareFormXXIIIGJOvertimeExportRows,
  resolveFormXXIIIGJOtTableHeaders,
  scoreFormXXIIIGJOtTableHeaderRow,
  isFormXXIIIGJOtWorkedDatesHeader,
  isFormXXIIIGJTotalOvertimeWorkedHeader,
  isFormXXIIIGJOvertimeRateHeader,
  isFormXXIIIGJOvertimeEarningsHeader,
  isFormXXIIIGJOtWagesPaidDateHeader,
} from './formXXIIIGJ';

describe('Form XXIII GJ overtime NIL columns', () => {
  const headers = [
    'Serial No.',
    'Name and surname of workmen',
    "Father's/Husband's name",
    'Sex',
    'Designation/ Nature of employment',
    'Dates on which overtime worked',
    'Total overtime worked or production in case of piece rates',
    'Normal rate of wages',
    'Overtime rate of wages',
    'Overtime earnings',
    'Date on which overtime wages paid',
    'Remarks',
  ];

  it('detects Gujarat Form XXIII from file name and state GJ', () => {
    expect(
      isFormXXIIIGJContext({}, { formName: 'Register of Overtime' }, 'Form_XXIII_GJ_-_Gujarat.xlsx', '')
    ).toBe(true);
    expect(
      isFormXXIIIGJContext({}, { formName: 'Register of Overtime', state: 'GJ' }, 'Form XXIII.xlsx', '')
    ).toBe(true);
    expect(
      isFormXXIIIGJContext({}, { formName: 'Register of Overtime' }, 'Form_XXIII_MP.xlsx', '')
    ).toBe(false);
  });

  it('rejects Form XIV-style preamble labels that are not OT table S.No.', () => {
    expect(
      isFormXXIIIGJOtPreambleFieldLabel(
        '2 Serial No. in the Register of workmen employed ..............................................'
      )
    ).toBe(true);
    expect(isFormXXIIIGJOtPreambleFieldLabel('1 Name of the Workman ..............................................')).toBe(
      true
    );
    expect(isFormXXIIIGJOtPreambleFieldLabel('5 Wage period ..............................................')).toBe(true);
    expect(isFormXXIIIGJOtPreambleFieldLabel('Serial No.')).toBe(false);
    expect(isFormXXIIIGJOtPreambleFieldLabel('S.No.')).toBe(false);
    expect(isFormXXIIIGJOtPreambleFieldLabel('Name and surname of workmen')).toBe(false);
    expect(isFormXXIIIGJOtPreambleFieldLabel('Designation/ Nature of employment')).toBe(false);
  });

  it('scores August-style OT thead higher than Form XIV preamble', () => {
    expect(
      scoreFormXXIIIGJOtTableHeaderRow(
        'Sl. No. Name of workman Father Sex Designation Overtime rate Normal rate'
      )
    ).toBeGreaterThan(14);
    expect(
      scoreFormXXIIIGJOtTableHeaderRow(
        '1 Name of the Workman 2 Serial No. in the Register of workmen employed'
      )
    ).toBeLessThan(0);
  });

  it('stamps August-style OT table and clears Form XIV preamble on ExcelJS sheet', () => {
    const cells = new Map();
    const columns = new Map();
    const worksheet = {
      model: { merges: [] },
      getCell(r, c) {
        const key = `${r}:${c}`;
        if (!cells.has(key)) {
          cells.set(key, {
            value: null,
            alignment: null,
            border: null,
            font: null,
          });
        }
        return cells.get(key);
      },
      getColumn(c) {
        if (!columns.has(c)) columns.set(c, { hidden: false, width: 20, model: { hidden: false } });
        return columns.get(c);
      },
      unMergeCells() {},
      mergeCells() {},
    };
    worksheet.getCell(15, 1).value = 'Name and address of Principal Employer: Acme';
    worksheet.getCell(16, 6).value = 'Rate of Wages_Normal Wages Earned_Gross Wage';
    worksheet.getCell(17, 1).value =
      '1 Name of the Workman ..............................................';
    worksheet.getCell(18, 1).value =
      '2 Serial No. in the Register of workmen employed ..............................................';
    worksheet.getCell(18, 9).value = 'Overtime rate';

    const layout = applyFormXXIIIGJOtTemplateLayoutToExcelJs(worksheet, headers);
    expect(layout).toBeTruthy();
    expect(layout.headerRow1).toBe(18);
    expect(layout.headers[0]).toMatch(/serial|sl\.?\s*no/i);
    expect(layout.headers).toContain('Sex');
    expect(String(worksheet.getCell(16, 6).value || '')).toBe('');
    expect(String(worksheet.getCell(17, 1).value || '')).toBe('');
    expect(String(worksheet.getCell(layout.headerRow1, 1).value || '')).toMatch(/serial|sl\.?\s*no/i);
    expect(worksheet.getCell(layout.indexRow1, 1).value).toBe(1);
    expect(worksheet.getCell(layout.indexRow1, 12).value).toBe(12);
    expect(layout.dataStartRow1).toBe(layout.indexRow1 + 1);
  });

  it('always uses August 12-col OT headers (never wages-register headers)', () => {
    const wagesHeaders = [
      'S. No.',
      'Name of the Employee',
      'Date of Appointment',
      'Sex',
      'Rate of Wages',
      'Normal Wages Earned',
      'Gross Wages',
      'Actual Wages',
    ];
    const resolved = resolveFormXXIIIGJOtTableHeaders(wagesHeaders);
    expect(resolved).toHaveLength(12);
    expect(resolved[0]).toBe('Serial No.');
    expect(resolved[3]).toBe('Sex');
    expect(resolved[4]).toMatch(/designation/i);
    expect(resolved.some((h) => /appointment|gross wages|actual wages/i.test(h))).toBe(false);
    expect(isFormXXIIIGJSexHeader('Sex')).toBe(true);
    expect(isFormXXIIIGJSexHeader('Designation/ Nature of employment')).toBe(false);
  });

  it('fills workman name and father from __employeeLookupName and row keys on export', () => {
    const rows = [
      {
        __employeeLookupId: 'VE1001',
        __employeeLookupName: 'Ravi Kumar',
        Sex: 'Male',
        'Designation/ Nature of employment': 'Junior Engineer',
        'Rate of Wages': '42596',
        "Father's/Husband's name": 'Suresh Kumar',
      },
    ];
    const prepared = prepareFormXXIIIGJOvertimeExportRows(rows, [
      'Sex',
      'Designation/ Nature of employment',
      'Rate of Wages',
    ]);
    expect(prepared[0]['Name and surname of workmen']).toBe('Ravi Kumar');
    expect(prepared[0]["Father's/Husband's name"]).toBe('Suresh Kumar');
  });

  it('fills father from __employeeLookupFatherOrSpouse when column cell is empty', () => {
    const rows = [
      {
        __employeeLookupId: 'E1002',
        __employeeLookupName: 'Santoo Kumar',
        __employeeLookupFatherOrSpouse: 'Ramesh Kumar',
        Sex: 'Male',
        'Designation/ Nature of employment': 'Engineer',
      },
    ];
    const prepared = prepareFormXXIIIGJOvertimeExportRows(rows, headers);
    expect(prepared[0]["Father's/Husband's name"]).toBe('Ramesh Kumar');
  });

  it('fills father from People record on export when grid only has lookup id', () => {
    const rows = [
      {
        __employeeLookupId: 'E1001',
        __employeeLookupName: 'sushil parmar',
        Sex: 'Male',
        'Designation/ Nature of employment': 'Engineer',
      },
    ];
    const employees = [
      {
        EmployeeID: 'E1001',
        FirstName: 'Sushil',
        LastName: 'Parmar',
        Father_s_Name: 'Ram Parmar',
      },
    ];
    const prepared = prepareFormXXIIIGJOvertimeExportRows(rows, headers, {
      employees,
    });
    expect(prepared[0]['Name and surname of workmen']).toBe('Sushil Parmar');
    expect(prepared[0]["Father's/Husband's name"]).toBe('Ram Parmar');
  });

  it('remaps autofill/wages row keys onto August OT headers for download', () => {
    const wagesHeaders = [
      'S. No.',
      'Name of the Employee',
      "Father's/Husband's name",
      'Sex',
      'Rate of Wages',
      'Remarks',
    ];
    const rows = [
      {
        'S. No.': '1',
        'Name of the Employee': 'Ravi Kumar',
        "Father's/Husband's name": 'Suresh',
        Sex: 'Male',
        'Rate of Wages': '30211',
        Remarks: 'OK',
      },
    ];
    const prepared = prepareFormXXIIIGJOvertimeExportRows(rows, wagesHeaders);
    expect(prepared).toHaveLength(1);
    expect(prepared[0]['Serial No.']).toBe('1');
    expect(prepared[0]['Name and surname of workmen']).toBe('Ravi Kumar');
    expect(prepared[0]["Father's/Husband's name"]).toBe('Suresh');
    expect(prepared[0]['Sex']).toBe('Male');
    expect(prepared[0]['Normal rate of wages']).toBe('30211');
    expect(prepared[0]['Remarks']).toBe('OK');
    expect(prepared[0]['Dates on which overtime worked']).toBe(FORM_XXIII_GJ_OT_NIL);
    expect(prepared[0]['Overtime earnings']).toBe(FORM_XXIII_GJ_OT_NIL);
  });

  it('keeps August OT autofill values when headers already match', () => {
    const rows = [
      {
        'Serial No.': '2',
        'Name and surname of workmen': 'Anita',
        "Father's/Husband's name": 'Ramesh',
        Sex: 'Female',
        'Designation/ Nature of employment': 'Operator',
        'Normal rate of wages': '25000',
        Remarks: '',
      },
    ];
    const prepared = prepareFormXXIIIGJOvertimeExportRows(rows, headers);
    expect(prepared[0]['Name and surname of workmen']).toBe('Anita');
    expect(prepared[0]['Designation/ Nature of employment']).toBe('Operator');
    expect(prepared[0]['Normal rate of wages']).toBe('25000');
    expect(prepared[0]['Serial No.']).toBe('1');
  });

  it('clears Remarks column thead bleed (header label and column index 12)', () => {
    const rows = [
      {
        'Serial No.': '16',
        'Name and surname of workmen': 'Harish Chandra Pandey',
        Sex: 'Male',
        'Designation/ Nature of employment': 'Engineer',
        Remarks: 'Remarks',
      },
      {
        'Serial No.': '17',
        'Name and surname of workmen': 'Vikram Sharma',
        Sex: 'Male',
        'Designation/ Nature of employment': 'Engineer',
        Remarks: '12',
      },
    ];
    const prepared = prepareFormXXIIIGJOvertimeExportRows(rows, headers);
    expect(prepared).toHaveLength(2);
    expect(prepared[0].Remarks).toBe('');
    expect(prepared[1].Remarks).toBe('');
  });

  it('drops template sample people that Autofill never bound', () => {
    const rows = [
      {
        'Name of the Employee': 'Jai Venkata Vignesh',
        Sex: 'Male',
        'Rate of Wages': '0',
      },
      {
        'Name of the Employee': 'Real Worker',
        Sex: 'Female',
        'Rate of Wages': '21000',
        __employeeLookupId: 'VE1001',
        __employeeLookupName: 'Real Worker',
      },
    ];
    const prepared = prepareFormXXIIIGJOvertimeExportRows(rows, [
      'Name of the Employee',
      'Sex',
      'Rate of Wages',
    ]);
    expect(prepared).toHaveLength(1);
    expect(prepared[0]['Name and surname of workmen']).toBe('Real Worker');
    expect(prepared[0]['Normal rate of wages']).toBe('21000');
  });

  it('hides leftover columns past L and rebuilds full-width header bands', () => {
    const cells = new Map();
    const columns = new Map();
    const worksheet = {
      model: { merges: [] },
      getCell(r, c) {
        const key = `${r}:${c}`;
        if (!cells.has(key)) cells.set(key, { value: null, alignment: null, border: null, font: null });
        return cells.get(key);
      },
      getColumn(c) {
        if (!columns.has(c)) columns.set(c, { hidden: c >= 13 && c <= 15, width: 40, model: { hidden: false } });
        return columns.get(c);
      },
      unMergeCells() {},
      mergeCells() {},
    };
    worksheet.getCell(15, 1).value = 'Name and address of Principal Employer: Acme';
    worksheet.getCell(9, 1).value = 'Name and address of Contractor: VAYONA';
    worksheet.getCell(9, 10).value = 'spill text';
    applyFormXXIIIGJOtTemplateLayoutToExcelJs(worksheet, headers);
    expect(columns.get(13).hidden).toBe(true);
    expect(columns.get(14).hidden).toBe(true);
    expect(columns.get(15).hidden).toBe(true);
    expect(columns.get(1).width).toBe(7);
    expect(columns.get(2).width).toBe(22);
    expect(String(worksheet.getCell(9, 1).value || '')).toMatch(/contractor.*vayona/i);
    expect(String(worksheet.getCell(9, 10).value || '')).toBe('');
    // August model: header meta is plain text (no box borders).
    expect(worksheet.getCell(9, 1).border).toEqual({});
    expect(worksheet.getCell(15, 1).border).toEqual({});
    // OT thead keeps thin borders.
    expect(worksheet.getCell(18, 1).border?.top?.style).toBe('thin');
    expect(worksheet.getCell(18, 1).border?.left?.style).toBe('thin');
  });

  it('recognizes the five OT columns for NIL', () => {
    expect(isFormXXIIIGJOtWorkedDatesHeader('Dates on which overtime worked')).toBe(true);
    expect(
      isFormXXIIIGJTotalOvertimeWorkedHeader(
        'Total overtime worked or production in case of piece rates'
      )
    ).toBe(true);
    expect(isFormXXIIIGJOvertimeRateHeader('Overtime rate of wages')).toBe(true);
    expect(isFormXXIIIGJOvertimeEarningsHeader('Overtime earnings')).toBe(true);
    expect(isFormXXIIIGJOtWagesPaidDateHeader('Date on which overtime wages paid')).toBe(true);
    expect(isFormXXIIIGJOtNilHeader('Normal rate of wages')).toBe(false);
    expect(isFormXXIIIGJOtNilHeader('Name and surname of workmen')).toBe(false);
  });

  it('fills blank OT columns with NIL without touching normal rate', () => {
    const rows = applyFormXXIIIGJOtNilToMappedRows(
      [
        {
          'Name and surname of workmen': 'Ravi',
          'Normal rate of wages': '30211',
          'Dates on which overtime worked': '',
          'Overtime rate of wages': '',
        },
      ],
      headers
    );
    expect(rows[0]['Name and surname of workmen']).toBe('Ravi');
    expect(rows[0]['Normal rate of wages']).toBe('30211');
    expect(rows[0]['Dates on which overtime worked']).toBe(FORM_XXIII_GJ_OT_NIL);
    expect(rows[0]['Total overtime worked or production in case of piece rates']).toBe(
      FORM_XXIII_GJ_OT_NIL
    );
    expect(rows[0]['Overtime rate of wages']).toBe(FORM_XXIII_GJ_OT_NIL);
    expect(rows[0]['Overtime earnings']).toBe(FORM_XXIII_GJ_OT_NIL);
    expect(rows[0]['Date on which overtime wages paid']).toBe(FORM_XXIII_GJ_OT_NIL);
  });

  it('overwrites attendance OT values with NIL when overwrite is true', () => {
    const rows = applyFormXXIIIGJOtNilToMappedRows(
      [
        {
          'Dates on which overtime worked': '2026-05-09',
          'Total overtime worked or production in case of piece rates': '71:29',
          'Overtime rate of wages': '100',
        },
      ],
      headers,
      FORM_XXIII_GJ_OT_NIL,
      { overwrite: true }
    );
    expect(rows[0]['Dates on which overtime worked']).toBe(FORM_XXIII_GJ_OT_NIL);
    expect(rows[0]['Total overtime worked or production in case of piece rates']).toBe(
      FORM_XXIII_GJ_OT_NIL
    );
    expect(rows[0]['Overtime rate of wages']).toBe(FORM_XXIII_GJ_OT_NIL);
  });
});
