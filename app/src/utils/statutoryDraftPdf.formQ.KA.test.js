import * as XLSX from 'xlsx';
import { statutoryDraftPdfTestUtils } from './statutoryDraftPdf';
import {
  FORM_Q_KA_PDF_ACK,
  FORM_Q_KA_PDF_FIELDS,
  FORM_Q_KA_PDF_RULE,
  FORM_Q_KA_PDF_SEAL,
  FORM_Q_KA_PDF_SIGNATURE,
  FORM_Q_KA_PDF_SUBTITLE,
  FORM_Q_KA_PDF_TITLE,
  looksLikeFormQKarnatakaPdfContext,
  normalizeFormQKarnatakaPdfMatrix,
} from './statutoryDraftPdf.formQ.KA';

const { sheetToDenseMatrix, buildStatutoryPdfHeaderModel } = statutoryDraftPdfTestUtils;

const excelModelAoa = () => [
  ['', 'FORM Q', ''],
  ['', '(See Rule 24(9A))', ''],
  ['', 'APPOINTMENT ORDER', ''],
  ['', '', ''],
  [
    '',
    '1. Name & Address of the Establishment',
    'Gurmatkal Site, 110/33KV Substation, 8T-6902, Chanderki Village',
  ],
  ['', '2. Name & Address of the Employer', "M/s Orange Mamatkheda"],
  ['', '3. Name of the Employee', 'Sathisha Ningappa'],
  ['', '4. His/Her Postal Address', 'At post gurmitkal Tq.'],
  ['', '5. His/Her Permanent Address', 'At post gurmitkal Tq.'],
  ['', '6. Father/Husband Name', 'Ningappa Halakavadi'],
  ['', '7. Date of Birth', '05 Jun 1989'],
  ['', '8. Date of his/her entry into employment', '10 Aug 2026'],
  ['', '9. Designation', 'Engineer'],
  ['', '10. Nature of work entrusted to him/her', 'Engineer'],
  ['', '11. His/Her serial number in the Register of employment', 'VE1412'],
  [
    '',
    '12. Rates of wages payable to him/her',
    ['1. Basic  28376', '2. VDA', '3. Other allowances if any/  32638', '4. Total  66720'].join('\n'),
  ],
  ['', 'Place:\nDate:', 'Signature of employer'],
  ['', 'Acknowledgement by the employer with date & signature', 'Seal of the Establishment'],
];

describe('Form Q Karnataka PDF context', () => {
  test('detects appointment-order Excel and filename', () => {
    expect(
      looksLikeFormQKarnatakaPdfContext(
        ['FORM Q', '(See Rule 24(9A))', 'APPOINTMENT ORDER'],
        [['1. Name & Address of the Establishment', 'Site'], ['3. Name of the Employee', 'Sathisha']],
        'Sheet1',
        'Form_Q_Karnataka.xlsx'
      )
    ).toBe(true);
    expect(
      looksLikeFormQKarnatakaPdfContext(
        ['FORM Q', 'APPOINTMENT ORDER'],
        [['3. Name of the Employee', 'A'], ['12. Rates of wages payable to him/her', '1. Basic']],
        'Sheet1'
      )
    ).toBe(true);
  });

  test('rejects Maharashtra Form Q muster and Gujarat annual return', () => {
    expect(
      looksLikeFormQKarnatakaPdfContext(
        ['FORM Q'],
        [['Full Name of the Worker', 'Date of the Month', 'Working Hours', 'Interval for Rest']],
        'Form_Q_Maharashtra'
      )
    ).toBe(false);
    expect(
      looksLikeFormQKarnatakaPdfContext(
        ['FORM Q', 'Annual Return', 'Gujarat Shops And Establishments Act'],
        [['Particulars', 'Value']],
        'Form_Q_GJ'
      )
    ).toBe(false);
  });
});

describe('Form Q Karnataka PDF alignment (Excel model)', () => {
  test('keeps 2-column label|value rows in Excel order 1–12', () => {
    const normalized = normalizeFormQKarnatakaPdfMatrix(excelModelAoa(), 3, 0, []);
    expect(normalized.colCount).toBe(2);
    expect(normalized.formQKALayout).toBe(true);
    expect(normalized.metaLines).toEqual([
      FORM_Q_KA_PDF_TITLE,
      FORM_Q_KA_PDF_RULE,
      FORM_Q_KA_PDF_SUBTITLE,
    ]);

    FORM_Q_KA_PDF_FIELDS.forEach((spec, idx) => {
      expect(String(normalized.rows[idx][0])).toBe(spec.label);
    });
    expect(normalized.rows[0][1]).toMatch(/Gurmatkal Site/i);
    expect(normalized.rows[1][1]).toMatch(/Orange Mamatkheda/i);
    expect(normalized.rows[0][0]).toMatch(/^1\.\s*Name & Address of the Establishment/i);
    expect(normalized.rows[1][0]).toMatch(/^2\.\s*Name & Address of the Employer/i);
    expect(normalized.rows[2][0]).toMatch(/^3\.\s*Name of the Employee/i);
    expect(normalized.rows[2][1]).toBe('Sathisha Ningappa');
    expect(normalized.rows[8][1]).toBe('Engineer');
    expect(normalized.rows[9][1]).toBe('Engineer');
    expect(normalized.rows[10][1]).toBe('VE1412');
  });

  test('keeps wage rates in one right-hand cell as vertical lines', () => {
    const normalized = normalizeFormQKarnatakaPdfMatrix(excelModelAoa(), 3, 0, []);
    const wageRow = normalized.rows.find((r) => /rates of wages payable/i.test(String(r[0])));
    expect(wageRow).toBeTruthy();
    const lines = String(wageRow[1]).split(/\n/);
    expect(lines).toHaveLength(4);
    expect(lines[0]).toMatch(/1\.\s*Basic/);
    expect(lines[0]).toMatch(/28376/);
    expect(lines[1]).toMatch(/2\.\s*VDA/);
    expect(lines[2]).toMatch(/3\.\s*Other allowances/i);
    expect(lines[3]).toMatch(/4\.\s*Total/);
    expect(lines[3]).toMatch(/66720/);
  });

  test('footer matches Excel Place/Date | Signature and Acknowledgement | Seal', () => {
    const normalized = normalizeFormQKarnatakaPdfMatrix(excelModelAoa(), 3, 0, []);
    const footer = normalized.rows.slice(-2);
    expect(footer[0][0]).toMatch(/Place:/i);
    expect(footer[0][0]).toMatch(/Date:/i);
    expect(footer[0][1]).toBe(FORM_Q_KA_PDF_SIGNATURE);
    expect(footer[1][0]).toBe(FORM_Q_KA_PDF_ACK);
    expect(footer[1][1]).toBe(FORM_Q_KA_PDF_SEAL);
  });

  test('reorders stacked PDF meta (fields 3–12 then 1–2) back to Excel order', () => {
    const stacked = [
      ['FORM Q'],
      ['(See Rule 24(9A)'],
      ['APPOINTMENT ORDER'],
      ['3. Name of the Employee'],
      ['Saithivel Muthunaicker'],
      ['7. Date of Birth'],
      ['08 Jul 1989'],
      ['12. Rates of wages payable to him/her'],
      ['1. Basic  28376'],
      ['2. VDA'],
      ['3. Other allowances if any/  32638'],
      ['4. Total  66720'],
      ['1. Name & Address of the Establishment'],
      ['Gurmatkal Site, 110/33KV'],
      ['2. Name & Address of the Employer'],
      ["M/s Orange Mamatkheda Wind Pvt Ltd"],
      ['Place:'],
      ['Date:'],
      ['Signature of employer'],
      ['Acknowledgement by the employer with date & signature'],
      ['Seal of the Establishment'],
    ];
    const normalized = normalizeFormQKarnatakaPdfMatrix(stacked, 1, 0, []);
    expect(normalized.rows[0][0]).toMatch(/^1\.\s*Name & Address of the Establishment/i);
    expect(normalized.rows[0][1]).toMatch(/Gurmatkal Site/i);
    expect(normalized.rows[1][1]).toMatch(/Orange Mamatkheda/i);
    expect(normalized.rows[2][1]).toBe('Saithivel Muthunaicker');
    expect(String(normalized.rows[11][1]).split('\n')).toHaveLength(4);
  });

  test('sheetToDenseMatrix uses 2-column Excel layout and title-only header bands', () => {
    const ws = XLSX.utils.aoa_to_sheet(excelModelAoa());
    const matrix = sheetToDenseMatrix(ws, 'Form_Q_Karnataka');
    expect(matrix.formQKALayout).toBe(true);
    expect(matrix.colCount).toBe(2);
    expect(matrix.metaLines.join('\n')).toMatch(/FORM Q/i);
    expect(matrix.metaLines.join('\n')).toMatch(/See Rule 24\(9A\)/i);
    expect(matrix.metaLines.join('\n')).toMatch(/APPOINTMENT ORDER/i);

    const header = buildStatutoryPdfHeaderModel(
      matrix.metaLines,
      matrix.rows,
      matrix.tableStartRow,
      matrix.name
    );
    expect(header.titles.some((t) => /FORM Q/i.test(t))).toBe(true);
    expect(header.fields || []).toEqual([]);

    const employee = matrix.rows.find((r) => /name of the employee/i.test(String(r[0])));
    expect(employee[1]).toBe('Sathisha Ningappa');
    const wages = matrix.rows.find((r) => /rates of wages/i.test(String(r[0])));
    expect(String(wages[1])).toMatch(/1\.\s*Basic/);
    expect(String(wages[1])).toMatch(/\n2\.\s*VDA/);
  });
});
