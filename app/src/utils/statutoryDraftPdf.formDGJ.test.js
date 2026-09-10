import { statutoryDraftPdfTestUtils } from './statutoryDraftPdf';

const {
  buildStatutoryPdfHeaderModel,
  looksLikeFormDGJMusterRollPdfContext,
  isFormDGJGujaratOuterFootnoteText,
  FORM_DGJ_GJ_ELECTRONIC_FORMAT_NOTE_1,
  FORM_DGJ_GJ_ELECTRONIC_FORMAT_NOTE_2,
  FORM_DGJ_GJ_GOVERNOR_ORDER_NOTE,
} = statutoryDraftPdfTestUtils;

describe('Form D GJ PDF outer footnotes', () => {
  it('detects Form D Gujarat muster-roll PDF context', () => {
    expect(
      looksLikeFormDGJMusterRollPdfContext(
        ['FORM D', 'FORMAT OF ATTENDANCE/MUSTER-ROLL REGISTER'],
        [['Relay or Set Work', 'Summary No. of Days']],
        'Form_D_GJ'
      )
    ).toBe(true);
  });

  it('does not treat Rajasthan Form D as Gujarat Form D', () => {
    expect(
      looksLikeFormDGJMusterRollPdfContext(
        ['FORM D', 'Format of Attendance Register'],
        [['Place of work', 'Date', 'IN', 'OUT']],
        'Form_D_RJ'
      )
    ).toBe(false);
  });

  it('recognizes electronic-format and Governor footnote texts', () => {
    expect(isFormDGJGujaratOuterFootnoteText(FORM_DGJ_GJ_ELECTRONIC_FORMAT_NOTE_1)).toBe(true);
    expect(isFormDGJGujaratOuterFootnoteText(FORM_DGJ_GJ_ELECTRONIC_FORMAT_NOTE_2)).toBe(true);
    expect(isFormDGJGujaratOuterFootnoteText(FORM_DGJ_GJ_GOVERNOR_ORDER_NOTE)).toBe(true);
    expect(
      isFormDGJGujaratOuterFootnoteText('*Not necessary in case of elecrtonic format')
    ).toBe(true);
  });

  it('keeps Form D GJ footnotes out of the title band', () => {
    const model = buildStatutoryPdfHeaderModel(
      [
        'FORM D',
        'FORMAT OF ATTENDANCE/MUSTER-ROLL REGISTER',
        'Name of Establishment: Amreli Site',
        '*Not necessary in case of electronic format',
        '**Not necessary in case of electronic format',
        'By order and in the name of the Governor of Gujarat',
      ],
      [],
      0,
      'Form_D_GJ'
    );
    expect(model.titles.some((t) => /not\s+necessary/i.test(t))).toBe(false);
    expect(model.titles.some((t) => /governor\s+of\s+gujarat/i.test(t))).toBe(false);
    expect(model.fields.some((t) => /not\s+necessary/i.test(t))).toBe(false);
    expect(model.fields.some((t) => /governor\s+of\s+gujarat/i.test(t))).toBe(false);
  });
});
