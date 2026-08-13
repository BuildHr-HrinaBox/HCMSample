import * as XLSX from 'xlsx';
import { statutoryDraftPdfTestUtils } from './statutoryDraftPdf';

const {
  looksLikeFormXIXWageSlipPdfContext,
  isFormXIXWageSlipColHeaderBlob,
  normalizeFormXIXWageSlipPdfMatrix,
  sheetToDenseMatrix,
} = statutoryDraftPdfTestUtils;

describe('Form XIX Wage Slip PDF layout', () => {
  test('detects Form XIX wage-slip column headers', () => {
    expect(
      isFormXIXWageSlipColHeaderBlob(
        'No. of days worked Rate of daily wages/piece - rate No. of units worked in case of piece rate Dates on which overtime worked Overtime hours and amount of overtime wages'
      )
    ).toBe(true);
  });

  test('collapses sparse Excel merges into Excel-matching 5 columns', () => {
    // Mimic template merges: headers in A,C,D,F,H with empty spacer columns.
    const aoa = [
      ['FORM XIX'],
      ['[See rule 78 (1)(b)]'],
      ['Wages Slip'],
      [
        'Name and address of contractor....... VAYONA ENERGY PRIVATE LIMITED',
        '',
        '',
        '',
        '',
        '',
        '',
        '',
      ],
      [
        'No. of days worked',
        '',
        'Rate of daily wages/piece - rate',
        'No. of units worked in case of piece rate',
        '',
        'Dates on which overtime worked',
        '',
        'Overtime hours and amount of overtime wages',
        '',
        '',
      ],
      ['1', '', '2', '3', '', '4', '', '5', '', ''],
      ['31', '', 'Monthly Wages', '', '', 'NIL', '', 'NIL', '', ''],
      [],
      [
        'Gross wages payable',
        '',
        '',
        'Deductions, any',
        '',
        '',
        '',
        'If Actual wages paid',
        '',
        '',
      ],
      ['71392', '', '', '3464', '', '', '', '67928', '', ''],
      ['Signature of the contractor or his Representative'],
    ];
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    const matrix = sheetToDenseMatrix(ws, 'Form_XIX_Karnataka');

    expect(
      looksLikeFormXIXWageSlipPdfContext(matrix.metaLines, matrix.rows, matrix.name)
    ).toBe(true);
    expect(matrix.colCount).toBe(5);

    const header = matrix.rows[matrix.tableStartRow] || [];
    expect(header[0]).toMatch(/days\s+worked/i);
    expect(header[1]).toMatch(/rate\s+of\s+daily\s+wages/i);
    expect(header[2]).toMatch(/units\s+worked/i);
    expect(header[3]).toMatch(/dates\s+on\s+which\s+overtime/i);
    expect(header[4]).toMatch(/overtime\s+hours/i);

    const data = matrix.rows[matrix.tableStartRow + 2] || [];
    expect(data[0]).toBe('31');
    expect(data[1]).toBe('Monthly Wages');
    expect(data[2]).toBe('');
    expect(data[3]).toBe('NIL');
    expect(data[4]).toBe('NIL');

    const footerValues = matrix.rows[matrix.tableStartRow + 5] || [];
    expect(footerValues[0]).toBe('71392');
    expect(footerValues[2]).toBe('3464');
    expect(footerValues[4]).toBe('67928');
  });

  test('normalize keeps blank units column without shifting NIL left', () => {
    const rows = [
      [
        'No. of days worked',
        '',
        'Rate of daily wages/piece - rate',
        'No. of units worked in case of piece rate',
        '',
        'Dates on which overtime worked',
        '',
        'Overtime hours and amount of overtime wages',
      ],
      ['1', '', '2', '3', '', '4', '', '5'],
      ['31', '', 'Monthly Wages', '', '', 'NIL', '', 'NIL'],
      ['Gross wages payable', '', '', 'Deductions, any', '', '', '', 'If Actual wages paid'],
      ['71392', '', '', '3464', '', '', '', '67928'],
    ];
    const out = normalizeFormXIXWageSlipPdfMatrix(rows, 8, 0);
    expect(out.colCount).toBe(5);
    expect(out.rows[out.tableStartRow + 2]).toEqual(['31', 'Monthly Wages', '', 'NIL', 'NIL']);
  });
});
