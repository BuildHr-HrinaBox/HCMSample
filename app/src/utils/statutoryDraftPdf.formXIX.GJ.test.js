import {
  looksLikeFormXIXGJWageSlipPdfContext,
  normalizeFormXIXGJWageSlipPdfMatrix
} from './statutoryDraftPdf.formXIX.GJ';
import { statutoryDraftPdfTestUtils } from './statutoryDraftPdf';

const { isFormXIXAPWageSlipPdfContext } = statutoryDraftPdfTestUtils;

describe('Form XIX GJ Wage Slip PDF layout', () => {
  const gjRows = [
    ['FORM XIX'],
    ['[See rule 78 (2)(b)]'],
    ['Wage Slip'],
    ['Name and address of contractor', 'VAYONA ENERGY PRIVATE LIMITED'],
    ["Name and father's/husband's name of the workman", 'Aditya Sharma Nandlal Sharma'],
    ['Nature and location of work', 'GJ-Amreli'],
    ['Name of the workman', 'Aditya Sharma'],
    ['For the week/fortnight/month ending', ''],
    ['1 Number of days worked', '31'],
    ['2 Number of units worked in case of piece rate workers', ''],
    ['3 Rate of daily wages/piece rate', 'Monthly Wages'],
    ['4 Amount of overtime wages', 'NIL'],
    ['5 Gross wages payable', '61982'],
    ['6 Deductions, if any', '2899'],
    ['7 Net amount of wages paid', '59083'],
    ['.............................'],
    ['Initial of the Contractor or his representative']
  ];

  const gjExcelTypoRows = [
    ['', '', '', '', '', 'FORM XIX'],
    ['', '', '', '', '', 'Wage Slip'],
    ['[See rule 78 (2)(b)]'],
    ['Name and address if contractor........................'],
    ['VAYONA ENERGY PRIVATE LIMITED'],
    ["Name and father's/husband's name of the workman................", 'Anantharaman Sivaraman'],
    ['Sivaraman Anantharaman'],
    ['Nature and location of work............................', 'GJ-Amreli'],
    ['Name of the workman............................'],
    ['Anantharaman Sivaraman'],
    ['For the week/fought/month/ending............................'],
    ['Rate of daily wages/piece rate........................', '-433.48'],
    ['Amount of overtime wages........................', 'NIL'],
    ['Gross wages payable........................', '387491'],
    ['Deductions, if any........................', '700'],
    ['Net amount of wages paid........................', '255708'],
    ['1', 'Number of days worked........................ 31'],
    ['2', 'Number of units worked in case of piece rate workers........................'],
    ['.............................'],
    ['Initial of the Contractor or his representative']
  ];

  test('detects Gujarat Form XIX wage-slip PDF context', () => {
    expect(
      looksLikeFormXIXGJWageSlipPdfContext([], gjRows, 'Form_XIX_GJ_-_Gujarat')
    ).toBe(true);
  });

  test('detects GJ even with Excel "address if contractor" typo', () => {
    expect(
      looksLikeFormXIXGJWageSlipPdfContext([], gjExcelTypoRows, 'Sheet1', 'Form_XIX_GJ_-_Gujarat.xlsx')
    ).toBe(true);
  });

  test('does not treat Madhya Pradesh Form XIX as Gujarat', () => {
    expect(
      looksLikeFormXIXGJWageSlipPdfContext(
        [],
        [
          ['FORM XIX'],
          ['Wage Slip'],
          ['Name and address of contractor', 'ACME'],
          ['Nature and location of work', 'Bhopal'],
          ['1 Number of days worked', '26'],
          ['7 Net amount of wages paid', '1000']
        ],
        'Form_XIX_MP_-_Madhya_Pradesh'
      )
    ).toBe(false);
  });

  test('does not route Gujarat XIX through AP wage-slip context', () => {
    expect(isFormXIXAPWageSlipPdfContext([], gjRows, 'Form_XIX_GJ_-_Gujarat')).toBe(false);
    expect(isFormXIXAPWageSlipPdfContext([], gjExcelTypoRows, 'Form_XIX_GJ_-_Gujarat')).toBe(false);
  });

  test('uses MP-style single box: headers + wages inside, initials below', () => {
    const out = normalizeFormXIXGJWageSlipPdfMatrix(gjRows, 2, 0, []);

    expect(out.formXIXGJBoxedLayout).toBe(true);
    expect(out.formXIXAPLayout).toBe(true);
    expect(out.formXIXGJLayout).toBe(false);
    expect(out.metaLines).toEqual(['FORM XIX', '[See rule 78 (2)(b)]', 'Wage Slip']);
    expect(out.formXIXGJHeaderLines).toEqual([]);

    // Header fields are inside the table rows (not above-box meta lines).
    expect(out.rows[0][0]).toMatch(/Name and address of contractor/i);
    expect(out.rows[0][1]).toMatch(/VAYONA ENERGY PRIVATE LIMITED/i);
    expect(out.rows[1][0]).toMatch(/Name and father's\/husband's name of the workman/i);
    expect(out.rows[2][0]).toMatch(/Nature and location/i);
    expect(out.rows[2][1]).toMatch(/GJ-Amreli/i);
    expect(out.rows[3][0]).toMatch(/week\/fortnight\/month/i);

    // Wage particulars 1–7 follow headers inside the same box.
    const wageStart = out.rows.findIndex((r) => /^1\s+Number of days worked/i.test(String(r[0] || '')));
    expect(wageStart).toBeGreaterThan(3);
    expect(out.rows[wageStart][1]).toBe('31');
    expect(out.rows[wageStart + 2][1]).toBe('Monthly Wages');
    expect(out.rows[wageStart + 6][1]).toBe('59083');

    expect(out.formXIXAPFooterLines.join(' ')).toMatch(
      /Initial of the Contractor or his representative/i
    );
  });

  test('normalizes production Excel typo layout into boxed MP-style order', () => {
    const out = normalizeFormXIXGJWageSlipPdfMatrix(gjExcelTypoRows, 6, 0, []);

    expect(out.formXIXGJBoxedLayout).toBe(true);
    expect(out.metaLines[0]).toMatch(/FORM XIX/i);
    expect(out.rows[0][1]).toMatch(/VAYONA ENERGY PRIVATE LIMITED/i);
    expect(out.rows[2][1]).toMatch(/GJ-Amreli/i);

    const wageStart = out.rows.findIndex((r) => /^1\s+Number of days worked/i.test(String(r[0] || '')));
    expect(out.rows[wageStart][1]).toBe('31');
    expect(out.rows[wageStart + 2][1]).toBe('-433.48');
    expect(out.rows[wageStart + 6][1]).toBe('255708');

    const labels = out.rows.map((r) => r[0]).join(' | ');
    expect(labels.indexOf('Number of days worked')).toBeLessThan(labels.indexOf('Rate of daily'));
    expect(out.formXIXAPFooterLines.join(' ')).toMatch(/Initial of the Contractor/i);
  });
});
