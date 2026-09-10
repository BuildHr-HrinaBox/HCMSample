import { statutoryDraftPdfTestUtils } from './statutoryDraftPdf';

const {
  buildStatutoryPdfHeaderModel,
  isStatutoryFieldMetaLine,
  isStatutoryTitleMetaLine,
  looksLikeFormCGJLoanRecoveriesPdfContext,
  looksLikeFormCRajasthanPdfContext,
  reorderFormCGJPdfHeaderFields,
  isFormCWhereverApplicableFootnoteText,
  isFormCDamageLossFineFootnoteText,
} = statutoryDraftPdfTestUtils;

describe('Form C GJ PDF header field order', () => {
  it('treats Labour Identification No. as a field, not a title', () => {
    expect(isStatutoryFieldMetaLine('Labour Identification No.')).toBe(true);
    expect(isStatutoryTitleMetaLine('Labour Identification No.')).toBe(false);
    expect(
      isStatutoryFieldMetaLine('*Labour Identification No. of Principal Employer')
    ).toBe(true);
  });

  it('detects Form C Gujarat Loan/Recoveries PDF context', () => {
    expect(
      looksLikeFormCGJLoanRecoveriesPdfContext(
        [
          'FORM C',
          'Register of Loan/ Recoveries/ Damage/ Loss/ Fine/ Advance/ Absences',
        ],
        [],
        'Form_C_GJ'
      )
    ).toBe(true);
  });

  it('does not treat Gujarat Loan/Recoveries as Rajasthan Form C', () => {
    expect(
      looksLikeFormCRajasthanPdfContext(
        [
          'FORM C',
          'Register of Loan/ Recoveries/ Damage/ Loss/ Fine/ Advance/ Absences',
        ],
        [],
        'Form_C_GJ_-_Gujarat'
      )
    ).toBe(false);
  });

  it('orders Name of Establishment above Labour Identification No.', () => {
    const ordered = reorderFormCGJPdfHeaderFields([
      'Labour Identification No.',
      'Name of Establishment: Amreli Site',
      'Name and address of Principal Employer : Amreli Renewable',
      '*Labour Identification No. of Principal Employer',
    ]);
    expect(ordered[0]).toMatch(/Name of Establishment/i);
    expect(ordered[1]).toMatch(/^Labour Identification No\.?$/i);
    expect(ordered[2]).toMatch(/Principal Employer/i);
    expect(ordered[3]).toMatch(/Labour Identification No\. of Principal Employer/i);
  });

  it('builds header model with establishment before LIN (not under title)', () => {
    const model = buildStatutoryPdfHeaderModel(
      [
        'FORM C',
        'Register of Loan/ Recoveries/ Damage/ Loss/ Fine/ Advance/ Absences',
        'Labour Identification No.',
        'Name of Establishment: Amreli EDF Nanikundal Site',
        'Name and address of Principal Employer : Amreli Renewable Energy Pvt Ltd',
        '*Labour Identification No. of Principal Employer',
      ],
      [],
      0,
      'Form_C_GJ_-_Gujarat'
    );

    expect(model.titles.some((t) => /labour\s+identification/i.test(t))).toBe(false);
    expect(model.titles[0]).toMatch(/FORM\s*C/i);
    expect(model.titles[1]).toMatch(/Register of Loan/i);

    expect(model.fields[0]).toMatch(/Name of Establishment/i);
    expect(model.fields[1]).toMatch(/^Labour Identification No/i);
    expect(model.fields[2]).toMatch(/Principal Employer/i);
    expect(model.fields[3]).toMatch(/Labour Identification No\. of Principal Employer/i);
  });

  it('recognizes Form C footnote texts for outer-box / remove handling', () => {
    expect(isFormCWhereverApplicableFootnoteText('*Wherever applicable')).toBe(true);
    expect(isFormCWhereverApplicableFootnoteText('Wherever applicable')).toBe(true);
    expect(
      isFormCDamageLossFineFootnoteText('*Applicable only in case of damage/loss/fine')
    ).toBe(true);
  });

  it('keeps Wherever applicable / damage footnotes out of the title band', () => {
    const model = buildStatutoryPdfHeaderModel(
      [
        'FORM C',
        'Register of Loan/ Recoveries/ Damage/ Loss/ Fine/ Advance/ Absences',
        'Name of Establishment: Amreli Site',
        '*Wherever applicable',
        '*Applicable only in case of damage/loss/fine',
      ],
      [],
      0,
      'Form_C_GJ'
    );
    expect(model.titles.some((t) => /wherever\s+applicable/i.test(t))).toBe(false);
    expect(model.titles.some((t) => /applicable\s+only\s+in\s+case/i.test(t))).toBe(false);
    expect(model.fields.some((t) => /wherever\s+applicable/i.test(t))).toBe(false);
    expect(model.fields.some((t) => /applicable\s+only\s+in\s+case/i.test(t))).toBe(false);
  });
});
