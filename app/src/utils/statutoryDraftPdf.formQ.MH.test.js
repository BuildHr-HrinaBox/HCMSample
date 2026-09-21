/**
 * Maharashtra Form Q PDF — detection, group bands, font size 8.
 */
import {
  FORM_Q_MH_PDF_TABLE_FONT_SIZE,
  detectFormQMaharashtraPdfGroupBands,
  looksLikeFormQMaharashtraPdfContext,
  normalizeFormQMaharashtraPdfMatrix,
  rewriteFormQMaharashtraPdfHeader,
  formQMaharashtraColumnWeight,
} from './statutoryDraftPdf.formQ.MH';

describe('Form Q Maharashtra PDF context', () => {
  test('detects muster-roll Form Q MH', () => {
    expect(
      looksLikeFormQMaharashtraPdfContext(
        ['FORM Q', '(See Rule 26(1))', 'MUSTER-ROLL CUM WAGE REGISTER'],
        [['Sr. No.', 'Full Name of the worker', 'Date of the Month', '1', '2']],
        'FORM Q',
        'Form_Q_-_Maharashtra.xlsx'
      )
    ).toBe(true);
  });

  test('rejects Karnataka appointment order', () => {
    expect(
      looksLikeFormQMaharashtraPdfContext(
        ['FORM Q', '(See Rule 24(9A))', 'APPOINTMENT ORDER'],
        [['1. Name of the Employee']],
        'Form_Q_Karnataka',
        'Form_Q_Karnataka.xlsx'
      )
    ).toBe(false);
  });
});

describe('Form Q Maharashtra PDF grouping + font', () => {
  test('font size is 8', () => {
    expect(FORM_Q_MH_PDF_TABLE_FONT_SIZE).toBe(8);
  });

  test('detects Excel-model 3-row header groups (parent / leaf / index)', () => {
    // Dense matrix after merge collapse: parent label only in top-left, slaves blank.
    const parent = [
      'Sr. No.',
      'Full Name of the worker',
      'Designation',
      'Age',
      'Sex',
      'Date of entry into service',
      'Working hours',
      '',
      'Interval for Rest',
      '',
      'Date of the Month',
      ...Array.from({ length: 30 }, () => ''),
      'Total Days worked',
      'Minimum rate',
      'Actual Wages',
      'HRA',
      'DA',
      'Gross',
      'OT hours',
      'OT earn',
      'Deductions',
      ...Array.from({ length: 7 }, () => ''),
      'Total Deduction Rs.',
      'Net Payable',
    ];
    const leaf = [
      '',
      '',
      '',
      '',
      '',
      '',
      'From',
      'To',
      'From',
      'To',
      ...Array.from({ length: 31 }, (_, i) => String(i + 1)),
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      'Provident Fund Contribution',
      'Family Pension Rs.',
      'ESI Contribution Rs.',
      'Professional Tax Rs.',
      'Income Tax Rs. (if any)',
      'Loan and Interest',
      'Advances Rs.',
      'Other Deductions Rs. (if any)',
      '',
      '',
    ];
    const index = parent.map((_, i) => String(i + 1));
    const rows = [parent, leaf, index];
    const bands = detectFormQMaharashtraPdfGroupBands(rows, 0, 2, parent.length);

    const working = bands.find((b) => /working\s+hours/i.test(b.label));
    expect(working).toBeTruthy();
    expect(working.end - working.start).toBe(1);

    const interval = bands.find((b) => /interval\s+for\s+rest/i.test(b.label));
    expect(interval).toBeTruthy();
    expect(interval.end - interval.start).toBe(1);

    const dateBand = bands.find((b) => /date\s+of\s+the\s+month/i.test(b.label));
    expect(dateBand).toBeTruthy();
    expect(dateBand.end - dateBand.start).toBeGreaterThanOrEqual(20);

    const deductions = bands.find((b) => /^deductions$/i.test(b.label));
    expect(deductions).toBeTruthy();
    expect(deductions.end - deductions.start).toBeGreaterThanOrEqual(5);
  });

  test('unifies split Deductions banners and drops duplicate Signature column', () => {
    const parent = [
      'Sr. No.',
      'Name',
      'Working hours',
      '',
      'Interval for Rest',
      '',
      'Date of the Month',
      ...Array.from({ length: 3 }, () => ''),
      'Gross',
      'Deductions',
      '',
      '',
      'Deductions',
      '',
      'Total Deduction Rs.',
      'Net Payable',
      'Amount Deposited Rs.',
      'Signature / Thumb Impression of the worker (if required)',
      '',
      'Amount Deposited Rs.',
      'Signature / Thumb Impression of the worker (if required)',
    ];
    const leaf = [
      '',
      '',
      'From',
      'To',
      'From',
      'To',
      '1',
      '2',
      '3',
      '4',
      '',
      'Provident Fund Contribution',
      'Family Pension Rs.',
      'ESI Contribution Rs.',
      'Professional Tax Rs.',
      'Other Deductions Rs. (if any)',
      '',
      '',
      '',
      '',
      '',
      '',
    ];
    const normalized = normalizeFormQMaharashtraPdfMatrix([parent, leaf], parent.length, 0, []);
    expect(normalized.colCount).toBeLessThan(parent.length);
    const lastStack = `${normalized.rows[0][normalized.colCount - 1] || ''} ${
      normalized.rows[1][normalized.colCount - 1] || ''
    }`;
    expect(lastStack).toMatch(/signature|thumb/i);
    expect(normalized.rows[0].filter((v) => /^deductions?$/i.test(String(v || '').trim())).length).toBe(
      1
    );
    const bands = detectFormQMaharashtraPdfGroupBands(
      normalized.rows,
      0,
      1,
      normalized.colCount
    );
    const interval = bands.find((b) => /interval\s+for\s+rest/i.test(b.label));
    expect(interval.end - interval.start).toBe(1);
    const deductions = bands.find((b) => /^deductions$/i.test(b.label));
    expect(deductions.end - deductions.start).toBeGreaterThanOrEqual(3);
    expect(bands.filter((b) => /^deductions$/i.test(b.label)).length).toBe(1);
  });

  test('finds parent groups one row above tableStart (Sr. No. row)', () => {
    const rows = [
      [
        '',
        '',
        'Working hours',
        '',
        'Interval for Rest',
        '',
        'Date of the Month',
        ...Array(30).fill(''),
      ],
      [
        'Sr. No.',
        'Full Name of the worker',
        'From',
        'To',
        'From',
        'To',
        ...Array.from({ length: 31 }, (_, i) => String(i + 1)),
      ],
    ];
    const bands = detectFormQMaharashtraPdfGroupBands(rows, 1, 1, rows[1].length);
    expect(bands.some((b) => /working\s+hours/i.test(b.label) && b.end > b.start)).toBe(true);
    expect(bands.some((b) => /date\s+of\s+the\s+month/i.test(b.label) && b.end > b.start)).toBe(
      true
    );
  });

  test('rewrites header titles in Excel order', () => {
    const { titles, fields } = rewriteFormQMaharashtraPdfHeader(
      ['MUSTER-ROLL CUM WAGE REGISTER', 'FORM Q'],
      [
        'Month : May',
        'Name of the Employer: Acme',
        'Name of the Establishment: Site A',
        '(See Rule 26(1))',
      ]
    );
    expect(titles[0]).toMatch(/FORM Q/i);
    expect(titles.some((t) => /see\s+rule\s*26/i.test(t))).toBe(true);
    expect(fields[0]).toMatch(/establishment/i);
    expect(fields[1]).toMatch(/employer/i);
  });

  test('column weights keep day cells narrow and name wide', () => {
    expect(formQMaharashtraColumnWeight('Full Name of the worker')).toBeGreaterThan(10);
    expect(formQMaharashtraColumnWeight('15')).toBeLessThan(3);
    expect(formQMaharashtraColumnWeight('Sr. No.')).toBeLessThan(5);
  });

  test('scrubs stray Advances/Other under Total Deduction/Net and caps Deductions band', () => {
    // Parent above tableStart (Sr. No. leaf row) — mirrors real Excel → PDF matrix.
    const parent = [
      '',
      '',
      'Working hours',
      '',
      'Interval for Rest',
      '',
      'Date of the Month',
      '',
      '',
      'Gross',
      'OT',
      'Deductions',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      'Total Deduction Rs.',
      'Net Payable Rs.',
      'Date of Payment',
      'Amount Deposited Rs.',
      'Signature / Thumb Impression of the worker (if required)',
      '',
      'Signature / Thumb Impression of the worker (if required)',
    ];
    const leaf = [
      'Sr. No.',
      'Name',
      'From',
      'To',
      'From',
      'To',
      '1',
      '2',
      '3',
      '',
      '',
      'Provident Fund Contribution',
      'Family Pension Rs.',
      'ESI Contribution Rs.',
      'Professional Tax Rs.',
      'Income Tax Rs.',
      'Loan and Interest',
      'Advances Rs.',
      'Other Deductions Rs. (if any)',
      // Stray duplicates under Total Deduction / Net (the bug in the user PDF)
      'Advances Rs.',
      'Other Deductions Rs. (if any)',
      '',
      '',
      '',
      '',
      '',
    ];
    const normalized = normalizeFormQMaharashtraPdfMatrix([parent, leaf], parent.length, 1, []);
    const totalCol = normalized.rows[0].findIndex((v) => /total\s+deduction/i.test(String(v || '')));
    const netCol = normalized.rows[0].findIndex((v) => /net\s+payable/i.test(String(v || '')));
    expect(totalCol).toBeGreaterThan(0);
    expect(netCol).toBeGreaterThan(totalCol);
    expect(String(normalized.rows[1][totalCol] || '')).not.toMatch(/advances?/i);
    expect(String(normalized.rows[1][netCol] || '')).not.toMatch(/other\s+deductions?/i);

    const bands = detectFormQMaharashtraPdfGroupBands(
      normalized.rows,
      1,
      1,
      normalized.colCount
    );
    const deductions = bands.find((b) => /^deductions$/i.test(b.label));
    expect(deductions).toBeTruthy();
    expect(deductions.end).toBeLessThan(totalCol);
    expect(deductions.end - deductions.start).toBeGreaterThanOrEqual(5);

    const interval = bands.find((b) => /interval\s+for\s+rest/i.test(b.label));
    expect(interval.end - interval.start).toBe(1);

    const lastStack = `${normalized.rows[0][normalized.colCount - 1] || ''} ${
      normalized.rows[1][normalized.colCount - 1] || ''
    }`;
    expect(lastStack).toMatch(/signature|thumb/i);
  });
});
