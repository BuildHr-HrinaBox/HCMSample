import {
  FORM_Q_KARNATAKA_CANONICAL_TABLE_HEADERS,
  applyFormQKarnatakaEmployeeToRow,
  computeFormQKarnatakaOtherAllowances,
  enrichFormQKarnatakaPayrollRows,
  formQKarnatakaUpsertInlineStrCell,
  isFormQKarnatakaContext,
  patchFormQKarnatakaFastSheetXml,
  patchFormQKarnatakaWageRatesCellText,
  preferFormQKarnatakaPayrollRowsWithBasic,
  resolveFormQKarnatakaPayrollWages,
  sanitizeFormQKarnatakaWageAmount,
} from './formQKarnataka';

describe('Form Q Karnataka context detection', () => {
  test('detects Form Q when row state is Karnataka', () => {
    expect(
      isFormQKarnatakaContext(null, { formName: 'Form Q', state: 'Karnataka' }, '', '')
    ).toBe(true);
  });

  test('detects Form Q Appointment Order with Karnataka act', () => {
    expect(
      isFormQKarnatakaContext(
        null,
        {
          formName: 'Form Q',
          act: 'Karnataka Shops and Commercial Establishments Act - Appointment Order',
        },
        '',
        ''
      )
    ).toBe(true);
  });

  test('rejects Maharashtra Form Q attendance muster grid', () => {
    expect(
      isFormQKarnatakaContext(
        null,
        { formName: 'Form Q', state: 'Maharashtra' },
        '',
        'Full Name of the Worker Date of the Month Working Hours Interval for Rest'
      )
    ).toBe(false);
  });
});

describe('Form Q Karnataka wage rates cell patch', () => {
  test('places Basic and Total on separate vertical lines like the template', () => {
    const raw = [
      '1. Basic',
      '2. VDA',
      '3. Other allowances if any/',
      '4. Total                    / month(s)',
    ].join('\n');
    const patched = patchFormQKarnatakaWageRatesCellText(raw, {
      basic: '31948',
      total: '34602',
    });
    expect(patched.split('\n')).toEqual([
      '1. Basic  31948',
      '2. VDA',
      '3. Other allowances if any/',
      '4. Total  34602',
    ]);
    expect(patched).not.toMatch(/month/i);
  });

  test('rebuilds smashed single-line rates into vertical lines', () => {
    const raw = '1. Basic2. VDA3. Other allowances if any/4. Total / month(s)';
    const patched = patchFormQKarnatakaWageRatesCellText(raw, {
      basic: '29378',
      other: '34602',
      total: '63980',
    });
    expect(patched.split('\n')).toEqual([
      '1. Basic  29378',
      '2. VDA',
      '3. Other allowances if any/  34602',
      '4. Total  63980',
    ]);
  });

  test('rejects Total values like "7 month(s)"', () => {
    const patched = patchFormQKarnatakaWageRatesCellText(
      '1. Basic\n2. VDA\n3. Other allowances if any/\n4. Total',
      {
        basic: '29378',
        other: '34602',
        total: '7 month(s)',
      }
    );
    expect(patched.split('\n')[3]).toBe('4. Total');
    expect(patched).not.toMatch(/7 month/i);
  });

  test('sanitizeFormQKarnatakaWageAmount drops month labels', () => {
    expect(sanitizeFormQKarnatakaWageAmount('7 month(s)')).toBe('');
    expect(sanitizeFormQKarnatakaWageAmount('29378')).toBe('29378');
  });
});

describe('Form Q Karnataka fast ZIP XML patch', () => {
  test('upserts inline string cells by reference', () => {
    const xml =
      '<worksheet><sheetData><row r="5"><c r="C5" t="inlineStr"><is><t>old</t></is></c></row></sheetData></worksheet>';
    const out = formQKarnatakaUpsertInlineStrCell(xml, 'C5', 'Ramesh');
    expect(out).toContain('r="C5"');
    expect(out).toContain('Ramesh');
    expect(out).not.toContain('>old<');
  });

  test('patches employee name and wage rates cell from export row', () => {
    const baseXml = [
      '<worksheet><sheetData>',
      '<row r="8"><c r="C8" s="5"/></row>',
      '<row r="20"><c r="C20" s="12" t="inlineStr"><is><t>1. Basic&#10;2. VDA&#10;3. Other allowances if any/&#10;4. Total</t></is></c></row>',
      '</sheetData></worksheet>',
    ].join('');
    const positions = {
      fieldPositions: [
        {
          header: 'Name of the Employee',
          row: 8,
          col: 3,
          cellRef: 'C8',
          isWage: false,
        },
      ],
      wageRatesPosition: {
        row: 20,
        col: 3,
        cellRef: 'C20',
        baseRaw: '1. Basic\n2. VDA\n3. Other allowances if any/\n4. Total',
      },
    };
    const patched = patchFormQKarnatakaFastSheetXml(
      baseXml,
      {
        'Name of the Employee': 'Anita Rao',
        Basic: '38000',
        Total: '111786',
      },
      positions,
      FORM_Q_KARNATAKA_CANONICAL_TABLE_HEADERS
    );
    expect(patched).toContain('Anita Rao');
    expect(patched).toContain('s="5"');
    expect(patched).toContain('s="12"');
    expect(patched).toContain(
      '1. Basic  38000&#10;2. VDA&#10;3. Other allowances if any/&#10;4. Total  111786'
    );
    expect(patched).not.toMatch(/month/i);
  });
});

describe('Form Q Karnataka Basic from SamplePayroll', () => {
  test('resolveFormQKarnatakaPayrollWages reads SamplePayroll basic/gross/net aliases', () => {
    const wages = resolveFormQKarnatakaPayrollWages({
      employee_name: 'Test Emp',
      basic: 38000,
      hra: 12000,
      gross: 165620,
      netpay: 111786,
    });
    expect(wages.basic).toBe('38000');
    expect(wages.other).toBe(computeFormQKarnatakaOtherAllowances('165620', '38000', '12000'));
    expect(wages.total).toBe('111786');
  });

  test('resolveFormQKarnatakaPayrollWages reads earned_basic and earnings lines', () => {
    const wages = resolveFormQKarnatakaPayrollWages({
      employee_name: 'Test Emp',
      gross_pay: 165620,
      net_pay: 111786,
      earnings: [
        { name: 'Basic', type: 'basic', amount: 38000 },
        { name: 'HRA', type: 'hra', amount: 12000 },
      ],
    });
    expect(wages.basic).toBe('38000');
    expect(wages.other).toBe('115620');
    expect(wages.total).toBe('111786');
  });

  test('applyFormQKarnatakaEmployeeToRow fills Basic column', () => {
    const row = applyFormQKarnatakaEmployeeToRow(
      {},
      { FirstName: 'Test', LastName: 'Emp', Employee_ID: 'VE0452' },
      FORM_Q_KARNATAKA_CANONICAL_TABLE_HEADERS,
      {
        payrollRow: {
          basic: 38000,
          hra: 12000,
          gross_pay: 165620,
          net_pay: 111786,
        },
      }
    );
    expect(row.Basic).toBe('38000');
    expect(row['Other allowances if any']).toBe('115620');
    expect(row.Total).toBe('111786');
  });

  test('applyFormQKarnatakaEmployeeToRow fills Date of Birth from Date_of_birth', () => {
    const row = applyFormQKarnatakaEmployeeToRow(
      {},
      {
        FirstName: 'Test',
        LastName: 'Emp',
        Date_of_birth: '1990-05-15',
        Dateofjoining: '2020-01-01',
      },
      FORM_Q_KARNATAKA_CANONICAL_TABLE_HEADERS
    );
    expect(row['Date of Birth']).toBe('1990-05-15');
    expect(row['Date of his/her entry into employment']).toBe('2020-01-01');
  });

  test('enrichFormQKarnatakaPayrollRows fills empty Basic when Other/Total already set', () => {
    const mapped = [
      {
        Basic: '',
        'Other allowances if any': '115620',
        Total: '111786',
      },
    ];
    const hits = enrichFormQKarnatakaPayrollRows(
      mapped,
      [{ FirstName: 'Test', LastName: 'Emp' }],
      FORM_Q_KARNATAKA_CANONICAL_TABLE_HEADERS,
      {
        overwrite: false,
        resolvePayrollRow: () => ({
          basic: 38000,
          hra: 12000,
          gross_pay: 165620,
          net_pay: 111786,
        }),
      }
    );
    expect(hits).toBe(1);
    expect(mapped[0].Basic).toBe('38000');
  });

  test('preferFormQKarnatakaPayrollRowsWithBasic keeps rows that have Basic', () => {
    const preferred = preferFormQKarnatakaPayrollRowsWithBasic(
      [
        { employee_name: 'A', gross_pay: 100, net_pay: 80 },
        { employee_name: 'B', basic: 40, gross_pay: 100, net_pay: 80 },
      ],
      []
    );
    expect(preferred).toHaveLength(1);
    expect(preferred[0].employee_name).toBe('B');
  });
});
