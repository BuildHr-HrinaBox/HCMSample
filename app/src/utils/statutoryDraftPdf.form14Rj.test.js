import * as XLSX from 'xlsx';
import { statutoryDraftPdfTestUtils } from './statutoryDraftPdf';
import {
  FORM_14_RJ_DAY_ENTRIES_NOTE,
  FORM_14_RJ_FOOTER_NOTE,
  FORM_14_RJ_TABLE_HEADERS,
} from '../Pages/statutory/form14Rajasthan';

const {
  isForm14RajasthanColHeaderBlob,
  looksLikeForm14RajasthanPdfContext,
  isForm14RajasthanOvertimeFootnoteText,
  isForm14RajasthanDayEntriesNoteText,
  isForm14RajasthanFootnoteRow,
  sheetToDenseMatrix,
  buildStatutoryPdfHeaderModel,
} = statutoryDraftPdfTestUtils;

describe('Form 14 RJ PDF column alignment', () => {
  test('detects Form 14 RJ long column headers', () => {
    expect(isForm14RajasthanColHeaderBlob(FORM_14_RJ_TABLE_HEADERS.join(' '))).toBe(true);
    expect(
      isForm14RajasthanColHeaderBlob(
        'Name of Persons Employed Whether young person or not Total hours worked during the month'
      )
    ).toBe(true);
    expect(isForm14RajasthanColHeaderBlob('Description of Department (if applicable)')).toBe(
      false
    );
  });

  test('keeps Form 14 RJ headers in the table (not full-width meta bands)', () => {
    const aoa = [
      ['The Rajasthan Shops and Establishments Rules, 1959'],
      ['FORM 14'],
      ['Rule 22 - sub rule 3'],
      ['Record of the Hours of Work of Persons Employed'],
      ['(To be used only when Notice in Form 13 is exhibited)'],
      ['Description of Department (if applicable)'],
      ['Month : August', 'Year : 2026'],
      [...FORM_14_RJ_TABLE_HEADERS],
      ['virendra .', 'Young', 'Nil', 'Nil', 'Nil', 'Nil'],
      ['atul dhakad', 'Young', 'Nil', 'Nil', 'Nil', 'Nil'],
    ];
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    const matrix = sheetToDenseMatrix(ws, 'Form_14_RJ');

    expect(looksLikeForm14RajasthanPdfContext(matrix.metaLines, matrix.rows, matrix.name)).toBe(
      true
    );

    // Title / description / month-year stay as meta — not the 6 column titles.
    const metaBlob = (matrix.metaLines || []).join(' ').toLowerCase();
    expect(metaBlob).toMatch(/form\s*14/);
    expect(metaBlob).toMatch(/description of department/);
    expect(metaBlob).not.toMatch(/name of persons employed/);
    expect(metaBlob).not.toMatch(/whether young person/);

    const headerRow = matrix.rows[matrix.tableStartRow] || [];
    expect(headerRow.join(' ')).toMatch(/name of persons employed/i);
    expect(headerRow.join(' ')).toMatch(/whether young person/i);
    expect(headerRow.filter((c) => String(c || '').trim()).length).toBeGreaterThanOrEqual(6);

    const model = buildStatutoryPdfHeaderModel(
      matrix.metaLines,
      matrix.rows,
      matrix.tableStartRow,
      matrix.name,
      { preferredTitle: 'FORM 14', fileName: 'Form_14_RJ_-_Rajasthan.xlsx' }
    );
    const titleBlob = (model.titles || []).join(' ').toLowerCase();
    expect(titleBlob).not.toMatch(/name of persons employed/);
    expect(titleBlob).not.toMatch(/extent of overtime/);
  });

  test('detects Form 14 RJ footnotes for placement above system-generated note', () => {
    expect(isForm14RajasthanOvertimeFootnoteText(FORM_14_RJ_FOOTER_NOTE)).toBe(true);
    expect(
      isForm14RajasthanOvertimeFootnoteText(
        '*This column need not be filled by Commercial Establishments. In case of shops, Residential Hotels, Restaurants and Eating Houses and Theatres and other places of public amusement or entertainment, the extent of such overtime on each day shall be recorded in the days column against the employed person distinctively in red ink, indicating the time up to which such overtime work was taken from the employee'
      )
    ).toBe(true);
    expect(isForm14RajasthanDayEntriesNoteText(FORM_14_RJ_DAY_ENTRIES_NOTE)).toBe(true);
    expect(
      isForm14RajasthanDayEntriesNoteText(
        'Note: Entries relating to any day must be made on that day'
      )
    ).toBe(true);
    expect(isForm14RajasthanFootnoteRow([FORM_14_RJ_FOOTER_NOTE, '', '', '', '', ''])).toBe(true);
    expect(isForm14RajasthanFootnoteRow([FORM_14_RJ_DAY_ENTRIES_NOTE])).toBe(true);
    expect(isForm14RajasthanFootnoteRow(['virendra .', 'Young', 'Nil'])).toBe(false);
  });

  test('keeps Form 14 RJ footnotes in the sheet matrix below employee rows', () => {
    const aoa = [
      ['The Rajasthan Shops and Establishments Rules, 1959'],
      ['FORM 14'],
      ['Record of the Hours of Work of Persons Employed'],
      [...FORM_14_RJ_TABLE_HEADERS],
      ['virendra .', 'Young', 'Nil', 'Nil', 'Nil', 'Nil'],
      [FORM_14_RJ_FOOTER_NOTE],
      [FORM_14_RJ_DAY_ENTRIES_NOTE],
      ['This is a System Generated Document'],
    ];
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    const matrix = sheetToDenseMatrix(ws, 'Form_14_RJ');
    const blob = (matrix.rows || []).flat().join(' ').toLowerCase();
    expect(blob).toMatch(/need not be filled/);
    expect(blob).toMatch(/entries relating to any day/);
    expect(blob).toMatch(/system generated/);
    // Footnotes must not become title/meta bands.
    const metaBlob = (matrix.metaLines || []).join(' ').toLowerCase();
    expect(metaBlob).not.toMatch(/need not be filled/);
    expect(metaBlob).not.toMatch(/entries relating to any day/);
  });

  test('drops empty leading column A so Name is the first PDF column', () => {
    const { trimForm14RajasthanLeadingBlankPdfColumns } = statutoryDraftPdfTestUtils;
    const aoa = [
      ['The Rajasthan Shops and Establishments Rules, 1959'],
      ['FORM 14'],
      ['Record of the Hours of Work of Persons Employed'],
      // Excel template: blank col A, then 6 Form 14 columns.
      ['', ...FORM_14_RJ_TABLE_HEADERS],
      ['', 'virendra .', 'Young', '176', 'Nil', 'Nil', 'Nil'],
      ['', 'atul dhakad', 'Young', '184', 'Nil', 'Nil', 'Nil'],
    ];
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    const matrix = sheetToDenseMatrix(ws, 'Form_14_RJ');
    const headerRow = matrix.rows[matrix.tableStartRow] || [];
    expect(String(headerRow[0] || '')).toMatch(/name of persons employed/i);
    expect(String(headerRow[1] || '')).toMatch(/young person/i);
    expect(matrix.colCount).toBe(6);

    const dataRow = matrix.rows[matrix.tableStartRow + 1] || [];
    expect(String(dataRow[0] || '')).toMatch(/virendra/i);
    expect(String(dataRow[1] || '')).toMatch(/young/i);

    const manual = trimForm14RajasthanLeadingBlankPdfColumns(
      [
        ['', ...FORM_14_RJ_TABLE_HEADERS],
        ['', 'x', 'Young', 'Nil', 'Nil', 'Nil', 'Nil'],
      ],
      7,
      0
    );
    expect(manual.colCount).toBe(6);
    expect(manual.rows[0][0]).toMatch(/Name of Persons Employed/i);
  });
});
