import {
  FORM_XXI_TN_NIL,
  applyFormXXITamilNaduNilToMappedRows,
  isFormXXITamilNaduContext,
  isFormXXITamilNaduNilDefaultHeader,
} from './formXXITamilNadu';

describe('Form XXI Tamil Nadu Register of Fines NIL columns', () => {
  it('detects Form_XXI_-_TamilNadu.xlsx context', () => {
    expect(
      isFormXXITamilNaduContext(
        { title: 'FORM XXI', subtitle: 'Register of Fines' },
        { state: 'Tamil Nadu', formName: 'Form XXI' },
        'Form_XXI_-_TamilNadu.xlsx'
      )
    ).toBe(true);
    expect(isFormXXITamilNaduContext({}, null, 'Form_XXI_-_TamilNadu.xlsx')).toBe(true);
    expect(
      isFormXXITamilNaduContext(
        { title: 'FORM XXII' },
        { state: 'Tamil Nadu' },
        'Form_XXII_-_TamilNadu.xlsx'
      )
    ).toBe(false);
    expect(
      isFormXXITamilNaduContext(
        { title: 'FORM XXI' },
        { state: 'Andhra Pradesh' },
        'Form_XXI_-_Andhra_Pradesh.xlsx'
      )
    ).toBe(false);
  });

  it('matches the fine NIL default headers', () => {
    expect(isFormXXITamilNaduNilDefaultHeader('Act/Omission for which fine imposed')).toBe(true);
    expect(isFormXXITamilNaduNilDefaultHeader('Date of offence')).toBe(true);
    expect(isFormXXITamilNaduNilDefaultHeader('Whether workman showed cause against fine')).toBe(
      true
    );
    expect(
      isFormXXITamilNaduNilDefaultHeader(
        "Name of person in whose presence employee's explanation was heard"
      )
    ).toBe(true);
    expect(isFormXXITamilNaduNilDefaultHeader('Name of person in Whose')).toBe(true);
    expect(isFormXXITamilNaduNilDefaultHeader('Wage periods and wages payable')).toBe(true);
    expect(isFormXXITamilNaduNilDefaultHeader('Amount of fine Imposed')).toBe(true);
    expect(isFormXXITamilNaduNilDefaultHeader('Date on which fine realized')).toBe(true);
    expect(isFormXXITamilNaduNilDefaultHeader('Name and surname of workmen')).toBe(false);
    expect(isFormXXITamilNaduNilDefaultHeader("Father's/Husband's name")).toBe(false);
    expect(isFormXXITamilNaduNilDefaultHeader('Remarks')).toBe(false);
  });

  it('applies NIL to fine columns', () => {
    const headers = [
      'Name and surname of workmen',
      'Act/Omission for which fine imposed',
      'Date of offence',
      'Whether workman showed cause against fine',
      'Name of person in Whose',
      'Wage periods and wages payable',
      'Amount of fine Imposed',
      'Date on which fine realized',
      'Remarks',
    ];
    const rows = applyFormXXITamilNaduNilToMappedRows(
      [{ 'Name and surname of workmen': 'A', Remarks: '' }],
      headers,
      FORM_XXI_TN_NIL
    );
    expect(rows[0]['Name and surname of workmen']).toBe('A');
    expect(rows[0]['Act/Omission for which fine imposed']).toBe(FORM_XXI_TN_NIL);
    expect(rows[0]['Date of offence']).toBe(FORM_XXI_TN_NIL);
    expect(rows[0]['Whether workman showed cause against fine']).toBe(FORM_XXI_TN_NIL);
    expect(rows[0]['Name of person in Whose']).toBe(FORM_XXI_TN_NIL);
    expect(rows[0]['Wage periods and wages payable']).toBe(FORM_XXI_TN_NIL);
    expect(rows[0]['Amount of fine Imposed']).toBe(FORM_XXI_TN_NIL);
    expect(rows[0]['Date on which fine realized']).toBe(FORM_XXI_TN_NIL);
    expect(rows[0].Remarks).toBe('');
  });
});
