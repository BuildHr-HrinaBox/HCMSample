import * as XLSX from 'xlsx';
import { statutoryDraftPdfTestUtils } from './statutoryDraftPdf';
import { FORM_11_RJ_TABLE_HEADERS } from '../Pages/statutory/form11Rajasthan';

const {
  looksLikeForm11RajasthanPdfContext,
  scrubForm11RajasthanTotalHoursPdfColumn,
  sheetToDenseMatrix,
} = statutoryDraftPdfTestUtils;

describe('Form 11 RJ PDF total-hours blanking', () => {
  test('detects Form 11 RJ register context', () => {
    expect(
      looksLikeForm11RajasthanPdfContext(
        ['The Rajasthan Shops and Establishments Rules, 1959', 'FORM 11', 'Register of Employment'],
        [FORM_11_RJ_TABLE_HEADERS],
        'Form_11_RJ'
      )
    ).toBe(true);
  });

  test('scrubs leaked 5PM from Total hours column but keeps numeric hours', () => {
    const rows = [
      [...FORM_11_RJ_TABLE_HEADERS],
      ['virendra', 'Young', '9AM', '5PM', '9AM', '', '5PM', '5PM', 'NIL', 'NIL', 'NIL', 'NIL'],
      ['atul', 'Young', '9AM', '5PM', '9AM', '', '5PM', '184', 'NIL', 'NIL', 'NIL', 'NIL'],
      ['pradeep', 'Young', '9AM', '5PM', '9AM', '', '5PM', '240', 'NIL', 'NIL', 'NIL', 'NIL'],
    ];
    const scrubbed = scrubForm11RajasthanTotalHoursPdfColumn(rows, 0, 12);
    expect(scrubbed[1][7]).toBe('');
    expect(scrubbed[2][7]).toBe('184');
    expect(scrubbed[3][7]).toBe('240');
    // Rest-after column stays 5PM.
    expect(scrubbed[1][6]).toBe('5PM');
  });

  test('sheet matrix blanks non-numeric Total hours for Form 11 RJ', () => {
    const aoa = [
      ['The Rajasthan Shops and Establishments Rules, 1959'],
      ['FORM 11'],
      ['Register of Employment'],
      ['Rule 22 - sub rule 1'],
      [...FORM_11_RJ_TABLE_HEADERS],
      ['virendra', 'Young', '9AM', '5PM', '9AM', '', '5PM', '5PM', 'NIL', 'NIL', 'NIL', 'NIL'],
      ['atul', 'Young', '9AM', '5PM', '9AM', '', '5PM', '184', 'NIL', 'NIL', 'NIL', 'NIL'],
    ];
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    const matrix = sheetToDenseMatrix(ws, 'Form_11_RJ');
    expect(looksLikeForm11RajasthanPdfContext(matrix.metaLines, matrix.rows, matrix.name)).toBe(
      true
    );
    const dataRow = matrix.rows[matrix.tableStartRow + 1] || [];
    const hoursIdx = (matrix.rows[matrix.tableStartRow] || []).findIndex((h) =>
      /total\s+hours\s+worked/i.test(String(h || ''))
    );
    expect(hoursIdx).toBeGreaterThanOrEqual(0);
    expect(String(dataRow[hoursIdx] || '')).toBe('');
    const dataRow2 = matrix.rows[matrix.tableStartRow + 2] || [];
    expect(String(dataRow2[hoursIdx] || '')).toBe('184');
  });
});
