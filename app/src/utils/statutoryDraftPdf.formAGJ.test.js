import { statutoryDraftPdfTestUtils } from './statutoryDraftPdf';

const {
  buildStatutoryPdfHeaderModel,
  looksLikeFormAGJGujaratPdfContext,
  looksLikeFormAPdfContext,
  isFormAGJGujaratOuterFootnoteText,
  reorderFormAGJPdfHeaderFields,
  FORM_AGJ_GJ_AGE_NOTE,
  FORM_AGJ_GJ_SKILL_NOTE,
  FORM_AGJ_GJ_ELECTRONIC_NOTE,
  FORM_AGJ_GJ_WHEREVER_NOTE,
} = statutoryDraftPdfTestUtils;

describe('Form A GJ PDF layout', () => {
  it('detects Form A Gujarat employee-format PDF context', () => {
    expect(
      looksLikeFormAGJGujaratPdfContext(
        ['FORM A', 'FORMAT OF EMPLOYEE/ WORKMAN/ WORKER'],
        [["Father's/Spouse Name", 'Date of Birth', 'Education Level']],
        'Form_A_GJ'
      )
    ).toBe(true);
  });

  it('does not treat Rajasthan Form A as Gujarat Form A', () => {
    expect(
      looksLikeFormAGJGujaratPdfContext(
        ['FORM A', 'FORMAT OF EMPLOYEE REGISTER'],
        [['Employee Code', 'Category (HS/S/SS/US)']],
        'Form_A_RJ'
      )
    ).toBe(false);
  });

  it('detects Rajasthan Form A employee register PDF context', () => {
    expect(
      looksLikeFormAPdfContext(
        ['FORM A', 'FORMAT OF EMPLOYEE REGISTER', '[See rule 2(1)]'],
        [['Employee Code', 'Category (HS/S/SS/US)']],
        'Form_A_RJ_-_Rajasthan.xlsx'
      )
    ).toBe(true);
  });

  it('recognizes Form A GJ outer footnote texts', () => {
    expect(isFormAGJGujaratOuterFootnoteText(FORM_AGJ_GJ_AGE_NOTE)).toBe(true);
    expect(isFormAGJGujaratOuterFootnoteText(FORM_AGJ_GJ_SKILL_NOTE)).toBe(true);
    expect(isFormAGJGujaratOuterFootnoteText(FORM_AGJ_GJ_ELECTRONIC_NOTE)).toBe(true);
    expect(isFormAGJGujaratOuterFootnoteText(FORM_AGJ_GJ_WHEREVER_NOTE)).toBe(true);
    expect(isFormAGJGujaratOuterFootnoteText('***Wherever applicable')).toBe(true);
  });

  it('puts NAME OF OWNER on the row below Name of Establishment', () => {
    const ordered = reorderFormAGJPdfHeaderFields([
      'LABOUR IDENTIFICATION NO',
      'Name of Establishment: Amreli EDF Nanikundal Site',
      'NAME OF OWNER',
      'Name and address of Principal Employer: Amreli R',
    ]);
    expect(ordered[0]).toMatch(/Name of Establishment/i);
    expect(ordered[1]).toMatch(/NAME OF OWNER/i);
    expect(ordered.some((f) => /principal\s+employer/i.test(f))).toBe(true);
  });

  it('keeps Form A GJ footnotes out of the title/field bands', () => {
    const model = buildStatutoryPdfHeaderModel(
      [
        'FORM A',
        'FORMAT OF EMPLOYEE/ WORKMAN/ WORKER',
        'Name of Establishment: Amreli Site',
        'NAME OF OWNER',
        FORM_AGJ_GJ_AGE_NOTE,
        FORM_AGJ_GJ_SKILL_NOTE,
        FORM_AGJ_GJ_ELECTRONIC_NOTE,
        FORM_AGJ_GJ_WHEREVER_NOTE,
      ],
      [],
      0,
      'Form_A_GJ'
    );
    expect(model.titles.some((t) => /14\s*to\s*18/i.test(t))).toBe(false);
    expect(model.titles.some((t) => /highly\s*skilled/i.test(t))).toBe(false);
    expect(model.fields.some((t) => /14\s*to\s*18/i.test(t))).toBe(false);
    expect(model.fields.some((t) => /wherever\s+applicable/i.test(t))).toBe(false);
    expect(model.formAGJSplitRows || []).toHaveLength(0);
    expect(model.fields[0]).toMatch(/Name of Establishment/i);
    expect(model.fields[1]).toMatch(/NAME OF OWNER/i);
  });
});
