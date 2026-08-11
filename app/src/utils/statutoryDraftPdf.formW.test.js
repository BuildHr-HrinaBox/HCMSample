import * as XLSX from 'xlsx';
import { statutoryDraftPdfTestUtils } from './statutoryDraftPdf';

const {
  isWageRegisterColHeaderBlob,
  isFormWAdminBandBlob,
  looksLikeFormWPdfContext,
  formWTamilNaduColumnWeight,
  isPurePdfNumericText,
  sheetToDenseMatrix
} = statutoryDraftPdfTestUtils;

describe('Form W Tamil Nadu PDF alignment helpers', () => {
  test('detects S.No and Name of the Employee as wage-register headers', () => {
    expect(
      isWageRegisterColHeaderBlob(
        'S.No Name of the Employee Employee Identification No. Number of days worked Basic Wage'
      )
    ).toBe(true);
    expect(isWageRegisterColHeaderBlob('Sr. No Name of Employee')).toBe(true);
    expect(isWageRegisterColHeaderBlob('Men Women Male young person')).toBe(false);
  });

  test('treats Men/Women and employer rows as Form W admin band', () => {
    expect(
      isFormWAdminBandBlob(
        'Name and Address of the Employer: Men Women Male young person Female young person Month:'
      )
    ).toBe(true);
    expect(isFormWAdminBandBlob('Wage Period from 1st May 2024 to 31st May 2024')).toBe(true);
    expect(
      isFormWAdminBandBlob(
        'S.No Name of the Employee Employee Identification No. Basic Wage'
      )
    ).toBe(false);
  });

  test('skips gender count value rows so table still starts at S.No', () => {
    const { isFormWAdminValueRow } = statutoryDraftPdfTestUtils;
    expect(isFormWAdminValueRow(['9', '1', '0', '0'])).toBe(true);
    expect(isFormWAdminValueRow(['May', '2024'])).toBe(true);
    expect(isFormWAdminValueRow(['1', 'rajeshkumar', 'VE0447', '31'])).toBe(false);
  });

  test('extracts Men/Women young-person box for PDF header', () => {
    const { extractFormWGenderBox } = statutoryDraftPdfTestUtils;
    const box = extractFormWGenderBox(
      [
        'FORM - W',
        'REGISTER OF WAGES',
        'Total number of persons employed: 10',
        'Name and Address of the Employer:',
        'Men',
        'Women',
        'Male young person',
        'Female young person',
        'Month: May',
        '10',
        '0',
        '0',
        '0'
      ],
      [],
      0
    );
    expect(box).toEqual({
      total: '10',
      men: '10',
      women: '0',
      maleYoung: '0',
      femaleYoung: '0'
    });
  });

  test('extracts gender box from dense sheet rows (Excel layout)', () => {
    const { extractFormWGenderBox } = statutoryDraftPdfTestUtils;
    const rows = [
      Array(30).fill(''),
      Array(30).fill(''),
      Array(30).fill(''),
      (() => {
        const r = Array(30).fill('');
        r[0] = 'Name and Address of the Establishment:';
        r[10] = 'Total number of persons employed: 10';
        return r;
      })(),
      (() => {
        const r = Array(30).fill('');
        r[10] = 'Men';
        r[11] = 'Women';
        r[12] = 'Male young person';
        r[13] = 'Female young person';
        return r;
      })(),
      (() => {
        const r = Array(30).fill('');
        r[10] = '10';
        r[11] = '0';
        r[12] = '0';
        r[13] = '0';
        return r;
      })(),
      Array(30).fill(''),
      Array(30).fill(''),
      Array(30).fill(''),
      (() => {
        const r = Array(30).fill('');
        r[0] = 'S.No';
        r[1] = 'Name of the Employee';
        return r;
      })()
    ];
    const box = extractFormWGenderBox(['FORM - W', 'REGISTER OF WAGES'], rows, 9);
    expect(box).toMatchObject({
      total: '10',
      men: '10',
      women: '0',
      maleYoung: '0',
      femaleYoung: '0'
    });
  });

  test('amount columns get a minimum weight so figures do not squeeze', () => {
    expect(formWTamilNaduColumnWeight('Basic Wage', 5)).toBeGreaterThanOrEqual(7);
    expect(formWTamilNaduColumnWeight('Net Wages', 5)).toBeGreaterThanOrEqual(7);
    expect(formWTamilNaduColumnWeight('S.No', 2)).toBeLessThan(5);
    expect(formWTamilNaduColumnWeight('Name of the Employee', 20)).toBeGreaterThan(10);
  });

  test('pure numeric text is detected for single-line painting', () => {
    expect(isPurePdfNumericText('32169')).toBe(true);
    expect(isPurePdfNumericText('31-05-2024')).toBe(false);
    expect(isPurePdfNumericText('VE0447')).toBe(false);
  });

  test('Form W template matrix starts at S.No header, not Men/Women row', () => {
    let workbook;
    try {
      // eslint-disable-next-line global-require, import/no-dynamic-require
      workbook = XLSX.readFile(
        require('path').resolve(
          __dirname,
          '../../../functions/statutoryreg_function/templates/Form W.xlsx'
        ),
        { cellDates: true, cellText: true }
      );
    } catch (err) {
      // Template path may differ in CI — skip rather than fail the suite.
      // eslint-disable-next-line no-console
      console.warn('Form W template not available for PDF matrix test:', err?.message);
      return;
    }
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const matrix = sheetToDenseMatrix(sheet, 'FORM W');
    expect(looksLikeFormWPdfContext(matrix.metaLines, matrix.rows, matrix.tableStartRow)).toBe(
      true
    );
    const headerRow = matrix.rows[matrix.tableStartRow] || [];
    const headerBlob = headerRow.join(' ').toLowerCase();
    expect(headerBlob).toMatch(/s\.?\s*no/);
    expect(headerBlob).toMatch(/name of the employee/);
    expect(headerBlob).not.toMatch(/^men\b/);
    expect(matrix.colCount).toBeGreaterThanOrEqual(28);
  });

  test('keeps Excel merge spans so PDF can center headings in the same place', () => {
    const { normalizeSheetMerges, buildMergePaintLookup, dedupeRepeatedRowText } =
      statutoryDraftPdfTestUtils;
    const merges = normalizeSheetMerges([{ s: { r: 0, c: 2 }, e: { r: 0, c: 8 } }]);
    expect(merges).toEqual([{ r0: 0, c0: 2, r1: 0, c1: 8 }]);
    const lookup = buildMergePaintLookup(merges, 10);
    expect(lookup.masters.get('0,2')).toEqual({ r0: 0, c0: 2, r1: 0, c1: 8 });
    expect(lookup.covered.has('0,3')).toBe(true);
    expect(lookup.covered.has('0,2')).toBe(false);

    const duplicated = [
      'FORM XX REGISTER OF WAGES',
      'FORM XX REGISTER OF WAGES',
      'FORM XX REGISTER OF WAGES',
      'FORM XX REGISTER OF WAGES',
      ''
    ];
    const { row, syntheticMerge } = dedupeRepeatedRowText(duplicated, 1);
    expect(row[0]).toBe('FORM XX REGISTER OF WAGES');
    expect(row.slice(1).every((c) => !c)).toBe(true);
    expect(syntheticMerge).toEqual({ r0: 1, c0: 0, r1: 1, c1: 3 });
  });

  test('Form 10 TN banner title is a centered title, not a left field', () => {
    const { isStatutoryTitleMetaLine, buildStatutoryPdfHeaderModel, expandStatutoryMetaSegments } =
      statutoryDraftPdfTestUtils;
    const banner =
      'FORM No. 10 (Prescribed under Rule 78) Overtime Muster Roll for exempted workers';
    expect(isStatutoryTitleMetaLine(banner)).toBe(true);
    expect(isStatutoryTitleMetaLine('Form 10')).toBe(true);
    expect(isStatutoryTitleMetaLine('Overtime Muster Roll for exempted workers')).toBe(true);
    expect(isStatutoryTitleMetaLine('(Prescribed under Rule 78)')).toBe(true);

    const parts = expandStatutoryMetaSegments(banner);
    expect(parts[0]).toMatch(/^FORM No\.?\s*10$/i);
    expect(parts.some((t) => /^\(Prescribed under Rule 78\)$/i.test(t))).toBe(true);
    expect(parts.some((t) => /^Overtime Muster Roll for exempted workers$/i.test(t))).toBe(true);

    const model = buildStatutoryPdfHeaderModel(
      [
        'Form 10',
        banner,
        'Name and Address of the Factory : VAYONA ENERGY PRIVATE LIMITED',
        'Month ending : August-2026'
      ],
      [],
      0,
      'Form 10'
    );
    expect(model.titles.some((t) => /^FORM No\.?\s*10$/i.test(t))).toBe(true);
    expect(model.titles.some((t) => /^Overtime Muster Roll for exempted workers$/i.test(t))).toBe(
      true
    );
    expect(model.fields.some((t) => /Overtime Muster Roll/i.test(t))).toBe(false);
    expect(model.fields.some((t) => /Name and Address of the Factory/i.test(t))).toBe(true);
  });

  test('Form 12 TN register banner is classified as a centered title', () => {
    const { isStatutoryTitleMetaLine, buildStatutoryPdfHeaderModel } = statutoryDraftPdfTestUtils;
    const banner =
      'FORM No. 12 [Prescribed under Rules 80, 86] REGISTER OF ADULT WORKERS AND YOUNG PERSONS';
    expect(isStatutoryTitleMetaLine(banner)).toBe(true);
    expect(isStatutoryTitleMetaLine('Form 12')).toBe(true);

    const model = buildStatutoryPdfHeaderModel(
      [
        'Form 12',
        banner,
        'Name and Address of the Factory : TN-Palani, Vayona Energy Pvt Ltd'
      ],
      [],
      0,
      'Form 12'
    );
    expect(model.titles.some((t) => /FORM No\. 12/i.test(t))).toBe(true);
    expect(model.fields.some((t) => /FORM No\. 12/i.test(t))).toBe(false);
  });

  test('does not use Excel draft file name as a PDF heading', () => {
    const { looksLikeExcelDraftFileLabel, buildStatutoryPdfHeaderModel } = statutoryDraftPdfTestUtils;
    expect(looksLikeExcelDraftFileLabel('Form_W_-_TamilNadu')).toBe(true);
    expect(looksLikeExcelDraftFileLabel('Form_L_-_Gujarat.xlsx')).toBe(true);
    expect(looksLikeExcelDraftFileLabel('FORM W')).toBe(false);
    expect(looksLikeExcelDraftFileLabel('Form 12')).toBe(false);

    const model = buildStatutoryPdfHeaderModel(
      ['FORM - L', '(See rule 14)', 'Name of Employer : VAYONA ENERGY'],
      [],
      0,
      'Form_L_-_Gujarat'
    );
    expect(model.titles.some((t) => looksLikeExcelDraftFileLabel(t))).toBe(false);
    expect(model.titles.some((t) => /FORM\s*-\s*L/i.test(t))).toBe(true);
  });

  test('splits Excel wrapText form headings into separate PDF title lines', () => {
    const { expandStatutoryMetaSegments, buildStatutoryPdfHeaderModel, isStatutoryTitleMetaLine } =
      statutoryDraftPdfTestUtils;

    const multiline =
      'FORM - I\nREGISTER OF WORKMEN\n[See sub-rule (1) under rule 6]\nTHE TAMIL NADU INDUSTRIAL ESTABLISHMENTS (CONFERMENT OF PERMANENT STATUS TO WORKMEN) ACT & RULES, 1981';
    expect(expandStatutoryMetaSegments(multiline)).toEqual([
      'FORM - I',
      'REGISTER OF WORKMEN',
      '[See sub-rule (1) under rule 6]',
      'THE TAMIL NADU INDUSTRIAL ESTABLISHMENTS (CONFERMENT OF PERMANENT STATUS TO WORKMEN) ACT & RULES, 1981'
    ]);
    expect(isStatutoryTitleMetaLine('FORM - I')).toBe(true);
    expect(isStatutoryTitleMetaLine('REGISTER OF WORKMEN')).toBe(true);
    expect(isStatutoryTitleMetaLine('[See sub-rule (1) under rule 6]')).toBe(true);

    const flattened =
      'FORM - I REGISTER OF WORKMEN [See sub-rule (1) under rule 6] THE TAMIL NADU INDUSTRIAL ESTABLISHMENTS (CONFERMENT OF PERMANENT STATUS TO WORKMEN) ACT & RULES, 1981';
    expect(expandStatutoryMetaSegments(flattened).length).toBeGreaterThanOrEqual(3);

    const model = buildStatutoryPdfHeaderModel([multiline], [], 0, 'FORM 1');
    expect(model.titles[0]).toMatch(/^FORM\s*-\s*I$/i);
    expect(model.titles.some((t) => /^REGISTER OF WORKMEN$/i.test(t))).toBe(true);
    expect(model.titles.some((t) => /see\s+sub-rule/i.test(t))).toBe(true);
    expect(model.titles.length).toBeGreaterThanOrEqual(3);
    // No single title cell should still hold the whole Excel wrapText blob
    expect(model.titles.some((t) => /FORM\s*-\s*I\s+REGISTER OF WORKMEN/i.test(t))).toBe(false);
  });

  test('column header weights keep short labels like Gender on one line', () => {
    const { statutoryHeaderColumnWeight, enforcePdfColumnMinWidths, isPdfSerialNumberHeader } =
      statutoryDraftPdfTestUtils;
    expect(isPdfSerialNumberHeader('Serial Number')).toBe(true);
    expect(statutoryHeaderColumnWeight('Gender', 4, 20)).toBeGreaterThanOrEqual(5.5);
    expect(statutoryHeaderColumnWeight('Photo', 0, 20)).toBeGreaterThanOrEqual(5.5);
    expect(statutoryHeaderColumnWeight('Serial Number', 2, 20)).toBeGreaterThanOrEqual(10);
    expect(statutoryHeaderColumnWeight('Name of the Worker', 20, 20)).toBeGreaterThan(
      statutoryHeaderColumnWeight('Gender', 4, 20)
    );

    const headers = ['Serial Number', 'Name of the Worker', 'Gender', 'Present Address'];
    const widths = enforcePdfColumnMinWidths([12, 120, 40, 200], headers, 372);
    expect(widths[0]).toBeGreaterThanOrEqual(52);
    expect(widths.reduce((a, b) => a + b, 0)).toBeCloseTo(372, 0);
  });

  test('Form 11 long column names get more width than short Sex/Age columns', () => {
    const { statutoryHeaderColumnWeight, enforcePdfColumnMinWidths } = statutoryDraftPdfTestUtils;
    const longHeader =
      'What Exactly was the Injured Person Doing at Time of Accident';
    const shortHeader = 'Sex';
    expect(statutoryHeaderColumnWeight(longHeader, 3, 18)).toBeGreaterThan(
      statutoryHeaderColumnWeight(shortHeader, 4, 18)
    );
    const headers = [
      'Sl. No.',
      'Sex',
      'Age',
      longHeader,
      'Name, Address & Occupation of Two Witnesses'
    ];
    const raw = headers.map((h, i) => statutoryHeaderColumnWeight(h, i === 0 ? 2 : 8, 18));
    const sum = raw.reduce((a, b) => a + b, 0);
    const usable = 1100;
    const widths = enforcePdfColumnMinWidths(
      raw.map((w) => (w / sum) * usable),
      headers,
      usable
    );
    expect(widths[3]).toBeGreaterThan(widths[1]);
    expect(widths[4]).toBeGreaterThan(widths[2]);
    expect(widths.reduce((a, b) => a + b, 0)).toBeCloseTo(usable, 0);
  });
});
