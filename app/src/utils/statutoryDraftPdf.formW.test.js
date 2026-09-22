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

describe('Form 26 nil-of-the-month PDF merge', () => {
  const { isNilOfTheMonthPdfText, resolveNilOfTheMonthPdfSpan } = statutoryDraftPdfTestUtils;

  test('recognizes Nil / Nill of the month and Nil for the month lines', () => {
    expect(isNilOfTheMonthPdfText('Nill of the month')).toBe(true);
    expect(isNilOfTheMonthPdfText('Nil of the month')).toBe(true);
    expect(isNilOfTheMonthPdfText('Nil for the month of Jun 2026')).toBe(true);
    expect(isNilOfTheMonthPdfText('Accident description')).toBe(false);
  });

  test('spans the full table band for a single nil cell after merge collapse', () => {
    const row = ['Nill of the month', '', '', '', '', '', '', '', '', '', '', '', '', ''];
    expect(resolveNilOfTheMonthPdfSpan(row, 14, 5, 4)).toEqual({
      start: 0,
      end: 13,
      text: 'Nil of the Month'
    });
  });

  test('Form 1 TN / Form XXIX nil row includes month and year in centered text', () => {
    const row = [
      '1',
      'Nill of the month',
      'Aug 2026',
      '',
      '',
      '',
      '',
      '',
      '',
      ''
    ];
    expect(resolveNilOfTheMonthPdfSpan(row, 10, 5, 4)).toEqual({
      start: 0,
      end: 9,
      text: 'Nil of the Month Aug 2026'
    });
  });

  test('Form XXIX combined nil line already containing period stays formatted', () => {
    const row = ['NIL of the Month Aug 2026', '', '', '', '', '', '', '', '', ''];
    expect(resolveNilOfTheMonthPdfSpan(row, 10, 5, 4)).toEqual({
      start: 0,
      end: 9,
      text: 'Nil of the Month Aug 2026'
    });
  });

  test('Form 1 TN nil row ignores leftover payment date and still merges full width', () => {
    const row = [
      '1',
      'Nill of the month',
      'Jul 2026',
      '',
      '',
      '',
      '',
      '31-07-2026',
      '',
      '',
      '',
      ''
    ];
    expect(resolveNilOfTheMonthPdfSpan(row, 12, 5, 4)).toEqual({
      start: 0,
      end: 11,
      text: 'Nil of the Month Jul 2026'
    });
  });

  test('Form 1 TN nil row ignores signature date spill and still merges full width', () => {
    const row = [
      '1',
      'Nil of the month',
      'Aug 2026',
      '',
      '',
      '',
      '',
      '',
      '31-08-2026\nFor (Company Name)\nAuthorised Signatory\nSignature of Employer / Manager / Authorised Person',
      '',
      '',
      ''
    ];
    expect(resolveNilOfTheMonthPdfSpan(row, 12, 5, 4)).toEqual({
      start: 0,
      end: 11,
      text: 'Nil of the Month Aug 2026'
    });
  });

  test('does not treat header rows or accident data rows as nil spans', () => {
    const header = ['(1) Sl. No.', '(2) Date', '(3) Name'];
    expect(resolveNilOfTheMonthPdfSpan(header, 3, 3, 4)).toBeNull();
    const data = ['1', '01-01-2026', 'Ravi'];
    expect(resolveNilOfTheMonthPdfSpan(data, 3, 5, 4)).toBeNull();
  });
});

describe('Form C LWF PDF layout', () => {
  const {
    isFormCLwfColHeaderBlob,
    looksLikeExcelSheetTabName,
    buildStatutoryPdfHeaderModel,
    sheetToDenseMatrix,
    isPureNilPdfText
  } = statutoryDraftPdfTestUtils;

  test('detects Details + Quarter ending column header row', () => {
    expect(
      isFormCLwfColHeaderBlob(
        'Details of Fines and Unpaid Accumulations (1) Quarter ending 31-March-2024 (2) Quarter ending 30-June-2024 (3) Quarter ending 30-September-2024 (4) Quarter ending 31-December-2024 (5)'
      )
    ).toBe(true);
    expect(isFormCLwfColHeaderBlob('S.No Name of the Employee Basic Wage')).toBe(false);
  });

  test('treats LWF Act - Form C sheet tab as non-printable title', () => {
    expect(looksLikeExcelSheetTabName('LWF Act - Form C')).toBe(true);
    expect(looksLikeExcelSheetTabName('Form C')).toBe(true);
    expect(looksLikeExcelSheetTabName('Form-C')).toBe(true);
    expect(looksLikeExcelSheetTabName('Register of Fines and Unpaid Accumulations for the year - 2024')).toBe(
      false
    );
  });

  test('does not put sheet tab name above Form-C titles', () => {
    const model = buildStatutoryPdfHeaderModel(
      [
        'Form-C',
        '[See rule 29 of the Tamil Nadu Labour Welfare Fund Rules, 1973]',
        'Register of Fines and Unpaid Accumulations for the year - 2024'
      ],
      [],
      0,
      'LWF Act - Form C'
    );
    expect(model.titles.some((t) => /lwf\s+act/i.test(t))).toBe(false);
    expect(model.titles[0]).toMatch(/form\s*-?\s*c/i);
  });

  test('keeps quarter header row in the table matrix instead of meta bands', () => {
    const XLSX = require('xlsx');
    const aoa = [
      ['Form-C', '', '', '', ''],
      ['[See rule 29 of the Tamil Nadu Labour Welfare Fund Rules, 1973]', '', '', '', ''],
      ['Register of Fines and Unpaid Accumulations for the year - 2024', '', '', '', ''],
      ['Name of the Establishment : TN-Palani, Vayona Energy Pvt Ltd', '', '', '', ''],
      [
        'Details of Fines and Unpaid Accumulations (1)',
        'Quarter ending 31-March-2024 (2)',
        'Quarter ending 30-June-2024 (3)',
        'Quarter ending 30-September-2024 (4)',
        'Quarter ending 31-December-2024 (5)'
      ],
      ['1. Total Realisations under fines', 'Nil', 'Nil', 'Nil', 'Nil'],
      ['(i) Basic Wages', 'Nil', 'Nil', 'Nil', 'Nil']
    ];
    const sheet = XLSX.utils.aoa_to_sheet(aoa);
    const matrix = sheetToDenseMatrix(sheet, 'LWF Act - Form C');
    expect(matrix.tableStartRow).toBe(4);
    expect(matrix.metaLines.join(' ')).not.toMatch(/quarter\s+ending/i);
    expect(matrix.rows[matrix.tableStartRow][1]).toMatch(/quarter\s+ending/i);
    expect(isPureNilPdfText('Nil')).toBe(true);
    expect(isPureNilPdfText('NIL')).toBe(true);
  });

  test('PDF header keeps Name of the Establishment and drops signatory footer', () => {
    const {
      rewriteFormCTamilNaduLwfPdfHeader,
      isFormCTamilNaduLwfPdfSignatoryRow,
      buildStatutoryPdfHeaderModel
    } = statutoryDraftPdfTestUtils;
    const model = buildStatutoryPdfHeaderModel(
      [
        'Form-C',
        '[See rule 29 of the Tamil Nadu Labour Welfare Fund Rules, 1973]',
        'Register of Fines and Unpaid Accumulations for the year - 2026',
        'Theni Site, Vayona Energy Pvt Ltd, 274 A, Rani mangammal main road',
        'Authorised Signatory',
        'Signature of Employer / Manager / Authorised Person'
      ],
      [
        [
          'Details of Fines and Unpaid Accumulations (1)',
          'Quarter ending 31-March-2026 (2)',
          'Quarter ending 30-June-2026 (3)',
          'Quarter ending 30-September-2026 (4)',
          'Quarter ending 31-December-2026 (5)'
        ],
        ['1. Total Realisations under fines', 'Nil', 'Nil', 'Nil', 'Nil'],
        ['Authorised Signatory', '', '', '', ''],
        ['Signature of Employer / Manager / Authorised Person', '', '', '', '']
      ],
      0,
      'Form C'
    );
    expect(model.formCTamilNaduLwf).toBe(true);
    expect(model.fields.some((f) => /Name of the Establishment\s*:/i.test(f))).toBe(true);
    expect(model.fields.join('\n')).not.toMatch(/Authorised Signatory/i);
    expect(model.fields.join('\n')).not.toMatch(/Signature of Employer/i);
    expect(model.titles.join('\n')).not.toMatch(/Authorised Signatory/i);
    expect(isFormCTamilNaduLwfPdfSignatoryRow(['Authorised Signatory'])).toBe(true);
    const rewritten = rewriteFormCTamilNaduLwfPdfHeader(model.titles, model.fields);
    expect(rewritten.fields[0]).toMatch(/^Name of the Establishment\s*:/);
  });

  test('detects unpaid accumulations footnote for placement above system note', () => {
    const { isUnpaidAccumulationsFootnoteText, isUnpaidAccumulationsFootnoteRow, FORM_C_UNPAID_ACCUMULATIONS_FOOTNOTE } =
      statutoryDraftPdfTestUtils;
    expect(
      isUnpaidAccumulationsFootnoteText(
        '*See definition of "Unpaid Accumulations" under Section 2(I) of the Tamil Nadu Labour Welfare Fund Act, 1972'
      )
    ).toBe(true);
    expect(
      isUnpaidAccumulationsFootnoteRow([
        '*See definition of" Unpaid Accumulations" under Section 2(I) of the Tamil Nadu Labour WelfareFund Act.1972',
        '',
        '',
        ''
      ])
    ).toBe(true);
    expect(FORM_C_UNPAID_ACCUMULATIONS_FOOTNOTE).toMatch(/unpaid\s+accumulations/i);
    expect(isUnpaidAccumulationsFootnoteText('Nil')).toBe(false);
  });
});

describe('Form 25 Tamil Nadu PDF trailing columns', () => {
  const {
    looksLikeForm25TamilNaduPdfContext,
    findRemarksColumnIndex,
    trimTrailingBlankPdfColumns,
    sheetToDenseMatrix,
    buildStatutoryPdfHeaderModel
  } = statutoryDraftPdfTestUtils;

  test('detects Form 25 compensatory holidays context', () => {
    expect(
      looksLikeForm25TamilNaduPdfContext(
        ['FORM No - 25', 'MUSTER ROLL AND REGISTER OF COMPENSATORY HOLIDAYS'],
        [],
        'Form 25'
      )
    ).toBe(true);
  });

  test('trims blank columns after Remarks', () => {
    const rows = [
      ['S.No', 'Name', '1', '2', 'Total Days Worked', 'Remarks', '', '', ''],
      ['1', 'Ravi', '8', '8', '2', '', '', '', '']
    ];
    expect(findRemarksColumnIndex(rows)).toBe(5);
    const trimmed = trimTrailingBlankPdfColumns(rows, 9, { clampToRemarks: true });
    expect(trimmed.colCount).toBe(6);
    expect(trimmed.rows[0]).toEqual(['S.No', 'Name', '1', '2', 'Total Days Worked', 'Remarks']);
    expect(trimmed.rows[0].length).toBe(6);
  });

  test('sheet matrix does not pad empty columns after Remarks for Form 25', () => {
    const XLSX = require('xlsx');
    const header = Array(50).fill('');
    header[0] = 'S.No';
    header[1] = 'Name of the Worker';
    header[2] = 'Worker Identity Number';
    header[3] = 'Time at which work commences';
    header[4] = 'Rest interval';
    header[5] = 'Time at which work ends';
    header[6] = 'Scheme of Shifts';
    header[7] = 'Daily Hours of work including overtime (if any)*';
    for (let d = 1; d <= 31; d += 1) header[6 + d] = String(d);
    header[38] = 'Total Days Worked';
    header[39] = 'Total Hours Worked';
    header[40] = 'Number of days on Loss of Pay';
    header[41] = 'Remarks';
    // Intentionally leave cols 42–49 blank (template padding).
    const aoa = [
      ['FORM No - 25', ...Array(49).fill('')],
      ['MUSTER ROLL AND REGISTER OF COMPENSATORY HOLIDAYS', ...Array(49).fill('')],
      ['[Prescribed under rules 77(4), 103]', ...Array(49).fill('')],
      header,
      ['1', 'Ravi', 'E1', '09:00', '1hr', '18:00', 'General', ...Array(31).fill('8'), '20', '160', '0', '', ...Array(8).fill('')]
    ];
    const sheet = XLSX.utils.aoa_to_sheet(aoa);
    // Inflate !ref past Remarks like a padded Excel template.
    sheet['!ref'] = 'A1:AX10';
    const matrix = sheetToDenseMatrix(sheet, 'Form 25');
    expect(matrix.colCount).toBe(42); // 0..41 inclusive → Remarks
    expect(matrix.rows[matrix.tableStartRow][41]).toMatch(/remarks/i);
    expect(matrix.rows[matrix.tableStartRow][42]).toBeUndefined();
  });

  test('does not promote festival holiday box numbers into PDF titles', () => {
    const model = buildStatutoryPdfHeaderModel(
      ['FORM No - 25', 'MUSTER ROLL AND REGISTER OF COMPENSATORY HOLIDAYS', '1', '2', '3', '4', '5'],
      [],
      0,
      'Form 25'
    );
    expect(model.titles.some((t) => /^\d+$/.test(t))).toBe(false);
    expect(model.titles.some((t) => /form\s*no/i.test(t))).toBe(true);
  });

  test('drops leftover template employee rows after the last real worker', () => {
    const { trimForm25TamilNaduPdfTrailingEmployeeRows, isForm25TamilNaduEmployeePdfRow } =
      statutoryDraftPdfTestUtils;
    const header = [
      'S.No',
      'Name of the Worker',
      'Scheme of Shifts',
      '1',
      '2',
      '3',
      'Remarks'
    ];
    const rows = [
      header,
      ['1', 'S Muthu Kumaran', 'General Shift', 'A', 'WO', 'A', ''],
      ['2', 'Mugundhan K', 'General Shift', 'WO', 'A', 'A', ''],
      ['3', '', 'General Shift', '09:00', '', '', ''], // template shell — no name / attendance
      ['4', '', 'General Shift', '', '', '', ''],
      ['This is a System Generated Document', '', '', '', '', '', '']
    ];
    expect(isForm25TamilNaduEmployeePdfRow(rows[1])).toBe(true);
    expect(isForm25TamilNaduEmployeePdfRow(rows[3])).toBe(false);
    const trimmed = trimForm25TamilNaduPdfTrailingEmployeeRows(rows, 0);
    expect(trimmed).toHaveLength(4); // header + 2 employees + system note
    expect(trimmed[1][1]).toBe('S Muthu Kumaran');
    expect(trimmed[2][1]).toBe('Mugundhan K');
    expect(trimmed[trimmed.length - 1][0]).toMatch(/system generated/i);
  });

  test('drops Sheet3 pivot pages and keeps only Form 25 sheet', () => {
    const { filterStatutoryPdfMatrices, looksLikeAuxiliaryOrPivotPdfSheet } = statutoryDraftPdfTestUtils;
    const form25 = {
      name: 'Form 25',
      metaLines: ['FORM No - 25', 'MUSTER ROLL AND REGISTER OF COMPENSATORY HOLIDAYS'],
      rows: [
        ['S.No', 'Name of the Worker', '1', '2', 'Remarks'],
        ['1', 'Ravi', 'A', 'WO', '']
      ],
      tableStartRow: 0,
      colCount: 5
    };
    const sheet3 = {
      name: 'Sheet3',
      metaLines: [],
      rows: [
        ['Row Labels', '', '', 'Count of MALE', 'Count of FEMALE', 'Sum of Gross Wages'],
        ['Admin Assistant', '', '', '1', '0', '25000']
      ],
      tableStartRow: 0,
      colCount: 6
    };
    expect(looksLikeAuxiliaryOrPivotPdfSheet(sheet3)).toBe(true);
    expect(looksLikeAuxiliaryOrPivotPdfSheet(form25)).toBe(false);
    const filtered = filterStatutoryPdfMatrices([form25, sheet3]);
    expect(filtered).toHaveLength(1);
    expect(filtered[0].name).toBe('Form 25');
  });

  test('keeps Form 15 Part 1 and drops Sheet3 pivot pages', () => {
    const {
      filterStatutoryPdfMatrices,
      looksLikeForm15TamilNaduPdfContext,
      stripLeakedPivotRowsFromPdfMatrix
    } = statutoryDraftPdfTestUtils;
    const form15 = {
      name: 'Form 15 Part 1',
      metaLines: ['FORM-15', 'REGISTER OF LEAVE WITH WAGES', '[See sub-rule (1)]'],
      rows: [
        ['S.No', 'Name of the employee', 'Employee Identification No.', 'Gender', 'Earned Leave'],
        ['1', 'Prabakaran', 'VE1189', 'Male', '0']
      ],
      tableStartRow: 0,
      colCount: 5
    };
    const sheet3 = {
      name: 'Sheet3',
      metaLines: [],
      rows: [
        ['Row Labels', '', '', 'Count of MALE', 'Count of FEMALE'],
        ['Admin Assistant', '', '', '2', '1']
      ],
      tableStartRow: 0,
      colCount: 5
    };
    expect(looksLikeForm15TamilNaduPdfContext(form15.metaLines, form15.rows, form15.name)).toBe(
      true
    );
    const filtered = filterStatutoryPdfMatrices([form15, sheet3]);
    expect(filtered).toHaveLength(1);
    expect(filtered[0].name).toBe('Form 15 Part 1');

    const leaked = [
      ['S.No', 'Name of the employee', 'Gender'],
      ['1', 'Prabakaran', 'Male'],
      ['Row Labels', '', 'Count of MALE'],
      ['Admin Assistant', '', '1'],
      ['This is a System Generated Document', '', '']
    ];
    const cleaned = stripLeakedPivotRowsFromPdfMatrix(leaked, 0);
    expect(cleaned.map((r) => r[0])).toEqual([
      'S.No',
      '1',
      'This is a System Generated Document'
    ]);
  });

  test('rewrites Form X titles when preferred context is Form 15 Part 1', () => {
    const {
      looksLikeForm15Part1PreferredContext,
      rewriteForm15Part1PdfTitles,
      buildStatutoryPdfHeaderModel
    } = statutoryDraftPdfTestUtils;
    expect(
      looksLikeForm15Part1PreferredContext('Form 15 Part 1', 'Form_15_Part_1_-_TamilNadu.xlsx')
    ).toBe(true);
    expect(looksLikeForm15Part1PreferredContext('Form X', 'Form_X_-_TamilNadu.xlsx')).toBe(false);

    expect(
      rewriteForm15Part1PdfTitles([
        'FORM-X',
        'REGISTER OF LEAVE AND SOCIAL SECURITY BENEFITS',
        '[See sub-rule (1) of rule (16)]'
      ])
    ).toEqual([
      'FORM-15',
      'REGISTER OF LEAVE WITH WAGES',
      '[See sub-rule (1) of rule (16)]'
    ]);

    // FORM-15 already present + leftover FORM-X must drop Form X.
    expect(rewriteForm15Part1PdfTitles(['FORM-15', 'FORM-X', 'REGISTER OF LEAVE WITH WAGES'])).toEqual(
      ['FORM-15', 'REGISTER OF LEAVE WITH WAGES']
    );

    const model = buildStatutoryPdfHeaderModel(
      ['FORM-15', 'FORM-X', 'REGISTER OF LEAVE WITH WAGES', '[See sub-rule (1) of rule (16)]'],
      [],
      0,
      'Sheet1',
      {
        preferredTitle: 'Form 15 Part 1',
        fileName: 'Form_15_Part_1_-_TamilNadu.xlsx'
      }
    );
    expect(model.titles[0]).toMatch(/form[\s._-]*15/i);
    expect(model.titles.some((t) => /leave with wages/i.test(t))).toBe(true);
    expect(model.titles.some((t) => /^form[\s._-]*x$/i.test(t))).toBe(false);
    expect(model.titles.join(' ')).not.toMatch(/\bform[\s._-]*x\b/i);
  });

  test('detects Earned / Medical / Other / Maternity leave group header bands', () => {
    const { detectLeaveCategoryBands } = statutoryDraftPdfTestUtils;
    const rows = [
      [
        'Name of the employee',
        'Employee Identification No.',
        '',
        'Earned Leave',
        '',
        '',
        '',
        'Medical Leave',
        '',
        '',
        'Other Leave',
        '',
        '',
        'Maternity Benefits',
        '',
        ''
      ],
      [
        '',
        '',
        'Gender',
        'Leave at the beginning of the Month',
        'Leave earned during the Period',
        'Leave availed during the Month',
        'Leave balance at the end of the Month',
        'Leave at beginning of the Month',
        'Leave availed during the Month',
        'Leave balance at end of the Month',
        'Leave at beginning of the Month',
        'Leave availed during the Month',
        'Leave Balance at end of the Month',
        'Date of giving notice',
        'Amount of Maternity Benefit',
        'Subsequent Maternity'
      ]
    ];
    const bands = detectLeaveCategoryBands(rows, 0, 1, 16);
    expect(bands).toHaveLength(4);
    expect(bands[0]).toMatchObject({ label: 'Earned Leave', start: 3, end: 6 });
    expect(bands[1]).toMatchObject({ label: 'Medical Leave', start: 7, end: 9 });
    expect(bands[2]).toMatchObject({ label: 'Other Leave', start: 10, end: 12 });
    expect(bands[3]).toMatchObject({ label: 'Maternity Benefits', start: 13, end: 15 });
  });

  test('detects Form 15 Part 2 Deductions / Advances / Damages group bands', () => {
    const { detectStatutoryGroupHeaderBands } = statutoryDraftPdfTestUtils;
    // Cols: OT, LeaveWages, Gross, PF, ESI, LWF, AdvPaid, AdvPend, AdvRec, AdvPend2,
    //       DamImp, DamPend, DamMade, DamPend2, AnyOther, TotalDed, Net, DatePay
    const rows = [
      [
        '',
        '',
        '',
        'Deductions',
        '',
        '',
        '',
        '',
        '',
        '',
        '',
        '',
        '',
        '',
        '',
        '',
        '',
        ''
      ],
      [
        '',
        '',
        '',
        '',
        '',
        '',
        'Advances',
        '',
        '',
        '',
        'Damages / Fine',
        '',
        '',
        '',
        '',
        '',
        '',
        ''
      ],
      [
        'Overtime Wages',
        'Leave Wages (Earned Leave / National, Festival & Special Holidays / Other)',
        'Gross Wages',
        'Provident Fund No.',
        "Employees' State Insurance Corporation No.",
        'Labour Welfare Fund',
        'Advance Paid',
        'Advance recovery pending at the beginning of the month',
        'Advance Recovered',
        'Pending Recovery',
        'Deduction imposed on Damages, Loss or Fines',
        'Deduction recovery pending at beginning of the month',
        'Deduction made on Damages, Loss or Fines',
        'Pending Recovery',
        'Any other Deductions',
        'Total Deductions',
        'Net Wages',
        'Date of payment'
      ],
      ['10', '11', '12', '13', '14', '15', '16', '17', '18', '19', '20', '21', '22', '23', '24', '25', '26', '27']
    ];
    const bands = detectStatutoryGroupHeaderBands(rows, 0, 3, 18);
    const byLabel = Object.fromEntries(bands.map((b) => [b.label, b]));
    expect(byLabel.Deductions).toMatchObject({ start: 3, end: 15, labelRow: 0 });
    expect(byLabel.Advances).toMatchObject({ start: 6, end: 9, labelRow: 1 });
    expect(byLabel['Damages / Fine']).toMatchObject({ start: 10, end: 13, labelRow: 1 });
  });

  test('detects Form VI festival holiday group band across day/date columns', () => {
    const {
      detectStatutoryGroupHeaderBands,
      isFormVIFestivalGroupLabel
    } = statutoryDraftPdfTestUtils;
    const festivalTitle =
      'Days, dates and months of the year on which National and Festival Holidays are allowed under the section 3 of the Tamil Nadu Industrial Establishments (National and Festival Holidays) Act, 1958 (Tamil Nadu Act XXXIII of 1958)';
    expect(isFormVIFestivalGroupLabel(festivalTitle)).toBe(true);
    const rows = [
      [
        'S.No',
        'Employee Code',
        'Name of the Employee',
        'D.O.J',
        festivalTitle,
        '',
        '',
        '',
        '',
        '',
        '',
        '',
        '',
        '',
        '',
        '',
        '',
        'Remarks'
      ],
      [
        '',
        '',
        '',
        '',
        '01/01/2022 (PONGAL)',
        '26/01/2022 (REPUBLIC DAY)',
        '14/04/2022 (TAMIL NEW YEAR)',
        '15/04/2022 (GOOD FRIDAY)',
        '01/05/2022 (MAY DAY)',
        '03/05/2022 (RAMZAN)',
        '15/08/2022 (INDEPENDENCE DAY)',
        '19/08/2022 (KRISHNA JAYANTHI)',
        '31/08/2022 (VINAYAKAR CHATHURTHI)',
        '04/10/2022 (AYUDHA POOJA)',
        '05/10/2022 (VIJAYA DASHAMI)',
        '24/10/2022 (DIWALI)',
        '25/12/2022 (CHRISTMAS)',
        ''
      ]
    ];
    const bands = detectStatutoryGroupHeaderBands(rows, 0, 1, 18);
    expect(bands).toHaveLength(1);
    expect(bands[0].start).toBe(4);
    expect(bands[0].end).toBe(16);
    expect(bands[0].label).toContain('Days, dates and months');
  });
});

describe('Form XXVII Tamil Nadu Register of Wages PDF', () => {
  test('detects Form XXVII register context and maps Wage Period to Month', () => {
    const {
      looksLikeFormXXVIITamilNaduRegisterPdfContext,
      extractMonthFromWagePeriodLine,
      buildStatutoryPdfHeaderModel
    } = statutoryDraftPdfTestUtils;

    expect(extractMonthFromWagePeriodLine('Wage Period : May')).toBe('May');
    expect(extractMonthFromWagePeriodLine('Wage Period from 1st May 2024 to 31st May 2024')).toBe(
      ''
    );

    const meta = [
      'FORM XXVII',
      'See Rule 78 (1) (a) of the Tamil Nadu Contract Labour (Regulation and Abolition) Rules, 1975',
      'REGISTER OF WAGES',
      'Address of the Establishment : Vayona Energy Pvt Ltd, Theni',
      'Wage Period : May',
      'Year: 2026'
    ];
    expect(
      looksLikeFormXXVIITamilNaduRegisterPdfContext(meta, [], 'Sheet1', 'Form_XXVII_-_TamilNadu.xlsx')
    ).toBe(true);

    const model = buildStatutoryPdfHeaderModel(meta, [], 0, 'Sheet1', {
      fileName: 'Form_XXVII_-_TamilNadu.xlsx'
    });
    expect(model.isFormXXVIIRegister).toBe(true);
    expect(model.titleBoxFullBorder).toBe(true);
    expect(model.hideRightBandSplit).toBe(true);
    expect(model.titles[0]).toMatch(/FORM XXVII/i);
    expect(model.titles.some((t) => /REGISTER OF WAGES/i.test(t))).toBe(true);
    expect(model.rightFields[0]).toMatch(/Month\s*:\s*May/i);
    expect(model.rightFields[1]).toMatch(/Year\s*:\s*2026/i);
    expect(model.fields.some((t) => /Wage Period/i.test(t))).toBe(false);
  });

  test('merges WAGES EARNED and DEDUCTIONS group bands', () => {
    const { detectStatutoryGroupHeaderBands, isWageDeductionGroupLabel } =
      statutoryDraftPdfTestUtils;
    expect(isWageDeductionGroupLabel('WAGES EARNED')).toBe(true);
    expect(isWageDeductionGroupLabel('DEDUCTIONS')).toBe(true);

    const rows = [
      [
        'S.No',
        'Name',
        'WAGES EARNED',
        '',
        '',
        '',
        'GROSS WAGES',
        'DEDUCTIONS',
        '',
        '',
        'NET WAGES'
      ],
      [
        '',
        '',
        'BASIC WAGE',
        'DA',
        'HRA',
        'OTHER ALLOWANCES, ECCA',
        'GROSS WAGES',
        'PF',
        'ESI',
        'TOTAL DEDUCTIONS',
        'NET WAGES'
      ],
      ['1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11']
    ];
    const bands = detectStatutoryGroupHeaderBands(rows, 0, 2, 11);
    const wages = bands.find((b) => /wages?\s+earned/i.test(b.label));
    const deductions = bands.find((b) => /^deductions?$/i.test(b.label));
    expect(wages).toBeTruthy();
    expect(wages.start).toBe(2);
    expect(wages.end).toBeGreaterThanOrEqual(4);
    expect(deductions).toBeTruthy();
    expect(deductions.start).toBe(7);
  });

  test('Form XXVII column weights prefer name over short amount labels', () => {
    const { formXXVIITamilNaduColumnWeight } = statutoryDraftPdfTestUtils;
    expect(formXXVIITamilNaduColumnWeight('Name of the Workman', 20)).toBeGreaterThan(
      formXXVIITamilNaduColumnWeight('HRA', 6)
    );
    expect(formXXVIITamilNaduColumnWeight('S.No', 2)).toBeLessThan(
      formXXVIITamilNaduColumnWeight('BASIC WAGE', 8)
    );
  });
});

describe('Form XXVI Tamil Nadu PDF header widths', () => {
  test('Permanent Home Address weight exceeds day leaf and Rate of Wages', () => {
    const { formXXVITamilNaduColumnWeight, formXXVITamilNaduLongestHeaderWord } =
      statutoryDraftPdfTestUtils;
    expect(formXXVITamilNaduLongestHeaderWord('Permanent Home Address')).toBe('Permanent');
    expect(formXXVITamilNaduColumnWeight('Permanent Home Address', 40)).toBeGreaterThan(
      formXXVITamilNaduColumnWeight('1', 1)
    );
    expect(formXXVITamilNaduColumnWeight('Permanent Home Address', 40)).toBeGreaterThan(
      formXXVITamilNaduColumnWeight('Rate of Wages', 8)
    );
    expect(formXXVITamilNaduColumnWeight('Number of days of work', 4)).toBeGreaterThan(
      formXXVITamilNaduColumnWeight('Age & Sex', 4)
    );
  });

  test('trailing Date of Termination and contractor signature get extra width', () => {
    const { formXXVITamilNaduColumnWeight, isFormXXVITamilNaduTrailingWideHeader } =
      statutoryDraftPdfTestUtils;
    expect(isFormXXVITamilNaduTrailingWideHeader('Date of Termination of Employment')).toBe(true);
    expect(isFormXXVITamilNaduTrailingWideHeader('Signature of contractor/representative')).toBe(
      true
    );
    expect(
      isFormXXVITamilNaduTrailingWideHeader('Signature/thumb impression of workman')
    ).toBe(true);
    expect(formXXVITamilNaduColumnWeight('Date of Termination of Employment', 0)).toBeGreaterThan(
      formXXVITamilNaduColumnWeight('Rate of Wages', 8)
    );
    expect(
      formXXVITamilNaduColumnWeight('Signature of contractor/representative', 0)
    ).toBeGreaterThan(formXXVITamilNaduColumnWeight('Number of days of work', 4));
  });

  test('ensureFormXXVIPdfHeaderBand adds index row and day band when missing', () => {
    const {
      ensureFormXXVIPdfHeaderBand,
      isFormXXVIStatutoryColumnIndexRow,
      isFormXXVIDayLeafNumberRow,
      inferFormXXVIPdfDayBand
    } = statutoryDraftPdfTestUtils;
    const colCount = 44;
    const textRow = [
      'Sr. No.',
      'Name of the Workman',
      'Age & Sex',
      'Permanent Home Address',
      'Local Address',
      'Designation (Nature of work)',
      "Father's/Husband's Name",
      'Date of entry into the service',
      'Rate of Wages',
      'Daily hours of work',
      ...Array.from({ length: 30 }, () => ''),
      'Number of days of work',
      'Signature/thumb impression of workman',
      'Date of Termination of Employment',
      'Signature of contractor/representative'
    ];
    while (textRow.length < colCount) textRow.push('');
    const matrix = {
      name: 'Form_XXVI_TamilNadu.xlsx',
      colCount,
      tableStartRow: 0,
      metaLines: ['FORM XXVI', 'Register of Employment of Contractual Labour'],
      rows: [textRow, ['Employee', 'Test']]
    };
    ensureFormXXVIPdfHeaderBand(matrix);
    expect(matrix.rows.length).toBeGreaterThanOrEqual(3);
    const dayBand = inferFormXXVIPdfDayBand(matrix.rows, 0, colCount);
    expect(isFormXXVIStatutoryColumnIndexRow(matrix.rows[1], dayBand, false)).toBe(true);
    expect(isFormXXVIDayLeafNumberRow(matrix.rows[2], dayBand, false)).toBe(true);
    expect(matrix.rows[2][dayBand.start]).toBe('1');
    expect(matrix.rows[2][dayBand.start + 30]).toBe('31');
    expect(matrix.rows[2][0]).toBe('');
    expect(matrix.rows[2][dayBand.end + 1]).toBe('');
  });

  test('redistributeFormXXVIPdfColumnWidths widens day band and trims leading cols', () => {
    const { redistributeFormXXVIPdfColumnWidths } = statutoryDraftPdfTestUtils;
    const headers = [
      'Sr. No.',
      'Name of the Workman',
      'Age & Sex',
      'Permanent Home Address',
      'Local Address',
      ...Array.from({ length: 31 }, (_, i) => String(i + 1)),
      'Number of days of work',
      'Signature/thumb impression of workman',
      'Date of Termination of Employment',
      'Signature of contractor/representative'
    ];
    const widths = headers.map((h, i) => {
      if (/^\d+$/.test(h)) return 10;
      if (i <= 4) return 55;
      return 30;
    });
    const usable = widths.reduce((a, b) => a + b, 0);
    const out = redistributeFormXXVIPdfColumnWidths(widths, headers, { start: 5, end: 35 }, usable);
    const dayBandWidth = out.slice(5, 36).reduce((a, b) => a + b, 0);
    const dayBandBefore = widths.slice(5, 36).reduce((a, b) => a + b, 0);
    expect(dayBandWidth).toBeGreaterThan(dayBandBefore);
    expect(out[0]).toBeLessThan(widths[0]);
    expect(out[1]).toBeLessThan(widths[1]);
    expect(out[3]).toBeLessThan(widths[3]);
    expect(out[4]).toBeLessThan(widths[4]);
  });

  test('wrapPdfTextPreferWholeWords keeps Permanent / Home / Address intact', () => {
    const { wrapPdfTextPreferWholeWords } = statutoryDraftPdfTestUtils;
    const fakeDoc = {
      getTextWidth: (s) => String(s || '').length * 5
    };
    // Width fits one long word (~9*5=45) but not two (~14*5=70).
    const lines = wrapPdfTextPreferWholeWords(fakeDoc, 'Permanent Home Address', 48, 6);
    expect(lines).toEqual(['Permanent', 'Home', 'Address']);
    expect(lines.every((line) => !/Permanen$|^t |W$|^ages|Num$|^ber/.test(line))).toBe(true);
  });
});
