import { statutoryDraftPdfTestUtils } from './statutoryDraftPdf';

const {
  buildStatutoryPdfHeaderModel,
  looksLikeFormBGJGujaratPdfContext,
  looksLikeFormPGJGujaratPdfContext,
  rewriteFormPGJGujaratPdfHeader,
  trimFormPGJGujaratLeadingBlankPdfColumns,
  detectFormPGJGujaratPdfGroupBands,
  dropFormPGJGujaratExtraDayPdfColumns,
  blankFormPGJGujaratMinimumRateOfWagesPdfCells,
  resolveFormPGJGujaratPdfDaysInMonth,
} = statutoryDraftPdfTestUtils;

describe('Form P GJ PDF Excel model', () => {
  const formPMeta = [
    'FORM - P',
    '(See rules 26)',
    'MUSTER - ROLL CUM WAGE REGISTER',
    'Name of the Establishment : Maliya Site, No.1',
    'Name of the employer: VAYONA ENERGY PRIVATE LIMITED',
    'Month: December',
  ];
  const formPRows = [
    [
      '',
      'Sr. No',
      'Full Name of the Worker',
      'Designation',
      'Age',
      'Sex',
      'Date of entry into service',
      'Working hours',
      'Interval for Rest',
      'Date of Month',
      ...Array.from({ length: 30 }, () => ''),
      'Total Days Worked',
      'Minimum Rate of Wages Rs.',
      'Actual Wages Paid Rs.',
      'DEDUCTION',
      'Date of Payment',
      'Signature/Thumb Impression',
    ],
    [
      '',
      '(1)',
      '(2)',
      '(3)',
      '(4)',
      '(5)',
      '(6)',
      '(7)',
      '(8)',
      ...Array.from({ length: 31 }, (_, i) => String(i + 1)),
      '(10)',
      '(11)',
      '(13)',
      'Provident Fund Rs.',
      '(29)',
      '(30)',
    ],
    ['', '1', 'Harshad Dhokiya', 'Engineer', '', '', '01 Dec 2025', '9 AM - 5 PM'],
  ];

  it('detects Form P Gujarat muster-roll PDF context', () => {
    expect(looksLikeFormPGJGujaratPdfContext(formPMeta, formPRows, 'Sheet1', 'Form_P_GJ_Gujarat.xlsx')).toBe(
      true
    );
  });

  it('does not treat Form P GJ as Form B GJ (no Rate of Minimum Wages box)', () => {
    expect(looksLikeFormBGJGujaratPdfContext(formPMeta, formPRows, 'Form_P_GJ_Gujarat', 'Form_P_GJ.xlsx')).toBe(
      false
    );
    const model = buildStatutoryPdfHeaderModel(formPMeta, formPRows, 0, 'Sheet1', {
      fileName: 'Form_P_GJ_Gujarat.xlsx',
    });
    expect(model.formPGJGujarat).toBe(true);
    expect(model.formBGJGujarat).toBe(false);
    expect(model.formBGJMinWagesBox).toBeFalsy();
  });

  it('keeps Establishment, Employer and Month as left Excel-style fields', () => {
    const model = buildStatutoryPdfHeaderModel(formPMeta, formPRows, 0, 'Sheet1', {
      fileName: 'Form_P_GJ_Gujarat.xlsx',
    });
    expect(model.titles.some((t) => /FORM\s*-?\s*P/i.test(t))).toBe(true);
    expect(model.titles.some((t) => /muster/i.test(t))).toBe(true);
    expect(model.fields.some((f) => /Name of the Establishment/i.test(f))).toBe(true);
    expect(model.fields.some((f) => /Name of the employer/i.test(f) && /VAYONA/i.test(f))).toBe(true);
    expect(model.fields.some((f) => /^Month\s*:/i.test(String(f)))).toBe(true);
    expect(model.rightFields.length).toBe(0);
    expect(model.fields.some((f) => /rate of minimum wages/i.test(f))).toBe(false);
  });

  it('rewriteFormPGJGujaratPdfHeader orders Excel header lines', () => {
    const rewritten = rewriteFormPGJGujaratPdfHeader(
      ['MUSTER - ROLL CUM WAGE REGISTER', 'FORM - P', '(See rules 26)'],
      ['Month: December', 'Name of the employer: VAYONA ENERGY PRIVATE LIMITED', 'Name of the Establishment : Maliya Site']
    );
    expect(rewritten.titles[0]).toMatch(/FORM/i);
    expect(rewritten.fields[0]).toMatch(/Establishment/i);
    expect(rewritten.fields[1]).toMatch(/employer/i);
    expect(rewritten.fields[2]).toMatch(/Month/i);
  });

  it('trims the empty column before Sr. No.', () => {
    const trimmed = trimFormPGJGujaratLeadingBlankPdfColumns(formPRows, formPRows[0].length, 0);
    expect(trimmed.colCount).toBe(formPRows[0].length - 1);
    expect(String(trimmed.rows[0][0])).toMatch(/Sr\.?\s*No/i);
  });

  it('spans Date of Month (9) across day columns 1–31 like Excel', () => {
    const bands = detectFormPGJGujaratPdfGroupBands(formPRows, 0, 1, formPRows[0].length);
    expect(bands.length).toBeGreaterThan(0);
    expect(bands[0].label).toBe('Date of Month (9)');
    expect(bands[0].end - bands[0].start).toBeGreaterThanOrEqual(20);
  });

  it('groups Date of Month (9) when the Excel cell includes the column index', () => {
    const rows = [
      [
        'Sr. No',
        'Full Name of the Worker',
        'Interval for Rest',
        'Date of Month (9)',
        ...Array.from({ length: 30 }, () => ''),
        'Total Days Worked',
      ],
      ['(1)', '(2)', '(8)', ...Array.from({ length: 31 }, (_, i) => String(i + 1)), '(10)'],
    ];
    const bands = detectFormPGJGujaratPdfGroupBands(rows, 0, 1, rows[0].length);
    expect(bands[0].label).toBe('Date of Month (9)');
    expect(bands[0].start).toBe(3);
    expect(bands[0].end - bands[0].start + 1).toBe(31);
  });

  it('June PDF keeps 30 day columns and July keeps 31', () => {
    expect(resolveFormPGJGujaratPdfDaysInMonth(['Month: June 2026'], [])).toBe(30);
    expect(resolveFormPGJGujaratPdfDaysInMonth(['Month: July 2026'], [])).toBe(31);
    const rows = [
      [
        'Sr. No',
        'Date of Month (9)',
        ...Array.from({ length: 30 }, () => ''),
        'Total Days Worked',
      ],
      ['(1)', ...Array.from({ length: 31 }, (_, i) => String(i + 1)), '(10)'],
      ['1', ...Array.from({ length: 31 }, () => 'P'), '30'],
    ];
    const june = dropFormPGJGujaratExtraDayPdfColumns(rows, rows[0].length, 0, 30);
    expect(june.rows[1].filter((c) => /^\d{1,2}$/.test(String(c))).join(',')).toBe(
      Array.from({ length: 30 }, (_, i) => String(i + 1)).join(',')
    );
    expect(june.rows[1]).toContain('30');
    expect(june.rows[1]).not.toContain('31');
    const july = dropFormPGJGujaratExtraDayPdfColumns(rows, rows[0].length, 0, 31);
    expect(july.rows[1]).toContain('31');
  });

  it('clears Minimum Rate of Wages Rs. (11) data cells on PDF and keeps Total Days Worked', () => {
    const rows = [
      [
        'Sr. No',
        'Date of Month (9)',
        'Total Days Worked',
        'Minimum Rate of Wages Rs.',
        'Actual Wages Paid Rs.',
      ],
      ['(1)', '1', '(10)', '(11)', '(13)'],
      ['1', 'P', '30', '21863', '21863'],
    ];
    blankFormPGJGujaratMinimumRateOfWagesPdfCells(rows, 0, 5);
    expect(rows[0][2]).toMatch(/Total Days Worked/i);
    expect(rows[1][2]).toBe('(10)');
    expect(rows[2][2]).toBe('30');
    expect(rows[0][3]).toMatch(/Minimum Rate of Wages/i);
    expect(rows[1][3]).toBe('(11)');
    expect(rows[2][3]).toBe('');
    expect(rows[2][4]).toBe('21863');
  });
});
