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
});
