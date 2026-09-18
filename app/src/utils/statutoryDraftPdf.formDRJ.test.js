import { statutoryDraftPdfTestUtils } from './statutoryDraftPdf';
import {
  FORM_D_RJ_ABSENCE_CODES_NOTE,
  FORM_D_RJ_E_FORM_NOTE,
  FORM_D_RJ_OUTER_FOOTNOTES,
  FORM_D_RJ_RELAY_MINES_NOTE,
} from '../Pages/statutory/formDRajasthan';

const {
  looksLikeFormDRajasthanPdfContext,
  isFormDRJOuterFootnoteText,
  isFormDRJOuterFootnoteRow,
  extractFormDRJOuterFootnoteText,
} = statutoryDraftPdfTestUtils;

describe('Form D Rajasthan PDF outer footnotes', () => {
  test('detects Form D RJ attendance register context', () => {
    expect(
      looksLikeFormDRajasthanPdfContext(
        ['FORM D', 'Rajasthan', 'Format of Attendance Register'],
        [['Place of work*', '1', 'Summary No. of Days']],
        'Form_D_RJ'
      )
    ).toBe(true);
    expect(
      looksLikeFormDRajasthanPdfContext(
        ['FORM D', 'Gujarat'],
        [['Relay or Set Work', 'Summary No. of Days']],
        'Form_D_GJ'
      )
    ).toBe(false);
  });

  test('recognizes Mines / absence / E Form notes', () => {
    expect(isFormDRJOuterFootnoteText(FORM_D_RJ_RELAY_MINES_NOTE)).toBe(true);
    expect(isFormDRJOuterFootnoteText(FORM_D_RJ_ABSENCE_CODES_NOTE)).toBe(true);
    expect(isFormDRJOuterFootnoteText(FORM_D_RJ_E_FORM_NOTE)).toBe(true);
    expect(FORM_D_RJ_OUTER_FOOTNOTES).toHaveLength(3);
  });

  test('extracts wrapped footnote row text', () => {
    const row = [
      '#Relay and *Place of Work in case of Mines only (Underground/Opencast/Surface)',
      '',
      '',
    ];
    expect(isFormDRJOuterFootnoteRow(row)).toBe(true);
    expect(extractFormDRJOuterFootnoteText(row)).toBe(FORM_D_RJ_RELAY_MINES_NOTE);
  });
});
