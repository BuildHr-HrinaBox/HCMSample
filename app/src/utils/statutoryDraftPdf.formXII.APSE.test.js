import { statutoryDraftPdfTestUtils } from './statutoryDraftPdf';
import {
  collapseFormXIIAPSERepeatedDisplayText,
  extractFormXIIAPSEAdminLayout,
  findFormXIIAPSEDuplicateColumnSpans,
  formXIIAPSEColumnWeight,
  formXIIAPSEDisplayCellText,
  formXIIAPSEHeaderWrapLines,
  formXIIAPSEPairedCellValue,
  getFormXIIAPSEHeaderTitles,
  isFormXIIAPSECenterValueHeader,
  looksLikeFormXIIAPSEPdfContext
} from './statutoryDraftPdf.formXII.AP.S&D.js';

describe('Form XII AP Shops & Establishment PDF layout helpers', () => {
  const meta = [
    'Form XII –',
    'Register of Advances of Wages',
    '[Vide Rule 18(4) of A.P. Shops & Establishments Rules, 1990]',
    'Name of Establishment : Nimbagallu Site, Vayona Energy Pvt Ltd, 274 A, Rani mangammal main road, Govindanagaram - 625517 Theni Taluk, Theni District., Nimbagallu, Andhra Pradesh'
  ];

  test('detects AP S&E Form XII Register of Advances and not XIII / XXII / XXIII', () => {
    expect(
      looksLikeFormXIIAPSEPdfContext(meta, [], 'Form XII – Register of Advances of Wages')
    ).toBe(true);
    expect(
      looksLikeFormXIIAPSEPdfContext(
        ['FORM XIII', 'Register of Workmen Employed by Contractor', '[Vide Rule 75]'],
        [],
        'Form XIII Andhra Pradesh'
      )
    ).toBe(false);
    expect(
      looksLikeFormXIIAPSEPdfContext(
        ['FORM XXII', 'Register of Advances'],
        [],
        'Form XXII Tamil Nadu'
      )
    ).toBe(false);
    expect(
      looksLikeFormXIIAPSEPdfContext(
        ['Form XXIII – Register of Wages', '(Vide Rule 29(2) of A.P. Shops & Establishment Rules, 1990)'],
        [],
        'Form XXIII – Register of Wages'
      )
    ).toBe(false);
  });

  test('keeps Form XII, register title and vide-rule as separate header lines', () => {
    const titles = getFormXIIAPSEHeaderTitles(meta, [], 4);
    expect(titles.filter((line) => /form\s*xii/i.test(line))).toHaveLength(1);
    expect(titles.find((line) => /form\s*xii/i.test(line))).toBe('Form XII');
    expect(titles.find((line) => /form\s*xii/i.test(line))).not.toMatch(/[-–—]/);
    expect(titles.some((line) => /^register of advances of wages$/i.test(line))).toBe(true);
    expect(titles.some((line) => /vide\s+rule\s*18/i.test(line))).toBe(true);
    expect(titles.join(' ')).not.toMatch(/form xii.+\s+form xii/i);
  });

  test('extracts establishment text without changing the value', () => {
    const layout = extractFormXIIAPSEAdminLayout(meta, [], 4);
    expect(layout.establishmentValue).toContain('Nimbagallu Site, Vayona Energy Pvt Ltd');
    expect(layout.establishmentLine).toContain('Name of Establishment :');
    expect(layout.establishmentLine).toContain(layout.establishmentValue);
  });

  test('wraps Name and Father headers onto two centered lines', () => {
    expect(formXIIAPSEHeaderWrapLines('Name of the Employee')).toEqual(['Name of', 'the Employee']);
    expect(formXIIAPSEHeaderWrapLines("Father's / Husband's Name")).toEqual([
      "Father's /",
      "Husband's Name"
    ]);
  });

  test('paints the same name in both paired sub-columns', () => {
    const value = formXIIAPSEPairedCellValue(
      ['1', 'Jai Venkata Vignesh', 'Jai Venkata Vignesh', 'Vignesh p', 'Vignesh p'],
      { start: 1, end: 2, header: 'Name of the Employee' }
    );
    expect(value).toBe('Jai Venkata Vignesh');
    expect(collapseFormXIIAPSERepeatedDisplayText('Jai Venkata Vignesh Jai Venkata Vignesh')).toBe(
      'Jai Venkata Vignesh'
    );
    expect(collapseFormXIIAPSERepeatedDisplayText('Lakshmanan Venkatachalam')).toBe(
      'Lakshmanan Venkatachalam'
    );
    expect(formXIIAPSEDisplayCellText('Jai Venkata Vignesh', 'Jai Venkata Vignesh')).toBe(
      'Jai Venkata Vignesh'
    );
  });

  test('spans empty-header duplicate name columns without concatenating values', () => {
    const rows = [
      [
        'S. No.',
        'Name of the Employee',
        '',
        "Father's / Husband's Name",
        '',
        'Amount of Advance Given',
        'Date on which Advance was Given',
        'Purpose(s) for which Advance was Given',
        'No. of Instalments by which Advance to be Recovered',
        'Postponement Granted',
        'Date on which Total Amount is Recovered',
        'Remarks'
      ],
      [
        '1',
        'Jai Venkata Vignesh',
        'Jai Venkata Vignesh',
        'Father One',
        'Father One',
        'NIL',
        'NIL',
        'NIL',
        'NIL',
        'NIL',
        'NIL',
        ''
      ]
    ];
    const spans = findFormXIIAPSEDuplicateColumnSpans(rows, 0, 0, 12);
    expect(spans).toEqual([
      { start: 1, end: 2, header: 'Name of the Employee' },
      { start: 3, end: 4, header: "Father's / Husband's Name" }
    ]);
  });

  test('aligns S.No and NIL columns center, and keeps names wider than serial', () => {
    expect(isFormXIIAPSECenterValueHeader('S. No.')).toBe(true);
    expect(isFormXIIAPSECenterValueHeader('Amount of Advance Given')).toBe(true);
    expect(isFormXIIAPSECenterValueHeader('Name of the Employee')).toBe(false);
    expect(formXIIAPSEColumnWeight('Name of the Employee')).toBeGreaterThan(
      formXIIAPSEColumnWeight('S. No.')
    );
  });

  test('PDF header model uses Form XII AP S&E layout flags', () => {
    const model = statutoryDraftPdfTestUtils.buildStatutoryPdfHeaderModel(
      meta,
      [
        [
          'S. No.',
          'Name of the Employee',
          "Father's / Husband's Name",
          'Amount of Advance Given',
          'Remarks'
        ]
      ],
      0,
      'Form XII – Register of Advances of Wages'
    );
    expect(model.formXIIAPSE).toBe(true);
    expect(model.formXIIAPSEAdminLayout.establishmentValue).toContain('Nimbagallu Site');
    expect(model.titles.some((line) => /vide\s+rule\s*18/i.test(line))).toBe(true);
  });
});
