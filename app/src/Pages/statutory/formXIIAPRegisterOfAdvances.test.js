import {
  FORM_XII_AP_NIL,
  applyFormXIIAPNilToMappedRows,
  isFormXIIAPNilDefaultHeader,
  isFormXIIAPRegisterOfAdvancesContext,
  matchesFormXIIAPFileHint,
} from './formXIIAPRegisterOfAdvances';

describe('Form XII Andhra Pradesh Register of Advances NIL columns', () => {
  it('detects Form_XII_-_Andhra_Pradesh.xlsx context', () => {
    expect(matchesFormXIIAPFileHint('Form_XII_-_Andhra_Pradesh.xlsx')).toBe(true);
    expect(matchesFormXIIAPFileHint('Form___XII_-_Andhra_Pradesh.xlsx')).toBe(true);
    expect(
      isFormXIIAPRegisterOfAdvancesContext(
        {
          title: 'Form XII – Register of Advances of Wages',
          subtitle: 'Register of Advances (Contract Labour)',
          reference: '[Vide Rule 18(4) of A.P. Shops & Establishments Rules, 1990]',
        },
        { state: 'Andhra Pradesh', formName: 'Form XII' },
        'Form_XII_-_Andhra_Pradesh.xlsx'
      )
    ).toBe(true);
    expect(isFormXIIAPRegisterOfAdvancesContext({}, null, 'Form_XII_-_Andhra_Pradesh.xlsx')).toBe(
      true
    );
  });

  it('does not match Form XXII / XXIX / XIII / TN advances', () => {
    expect(matchesFormXIIAPFileHint('Form_XXII_-_TamilNadu.xlsx')).toBe(false);
    expect(matchesFormXIIAPFileHint('Form_XXII_-_Andhra_Pradesh.xlsx')).toBe(false);
    expect(matchesFormXIIAPFileHint('Form_XXIX_-_TamilNadu.xlsx')).toBe(false);
    expect(matchesFormXIIAPFileHint('Form_XIII_-_Andhra_Pradesh.xlsx')).toBe(false);
    expect(
      isFormXIIAPRegisterOfAdvancesContext(
        { title: 'FORM XXII', subtitle: 'Register of Advances' },
        { state: 'Tamil Nadu' },
        'Form_XXII_-_TamilNadu.xlsx'
      )
    ).toBe(false);
  });

  it('matches the six advance NIL default headers', () => {
    expect(isFormXIIAPNilDefaultHeader('Amount of Advance Given')).toBe(true);
    expect(isFormXIIAPNilDefaultHeader('Date on which Advance was Given')).toBe(true);
    expect(isFormXIIAPNilDefaultHeader('Purpose(s) for which Advance was Given')).toBe(true);
    expect(
      isFormXIIAPNilDefaultHeader('No. of Instalments by which Advance to be Recovered')
    ).toBe(true);
    expect(isFormXIIAPNilDefaultHeader('Postponement Granted')).toBe(true);
    expect(isFormXIIAPNilDefaultHeader('Date on which Total Amount is Recovered')).toBe(true);
    expect(isFormXIIAPNilDefaultHeader('Name of the Workman')).toBe(false);
    expect(isFormXIIAPNilDefaultHeader('Remarks')).toBe(false);
  });

  it('applies NIL to advance columns', () => {
    const headers = [
      'Name of the Workman',
      'Amount of Advance Given',
      'Date on which Advance was Given',
      'Purpose(s) for which Advance was Given',
      'No. of Instalments by which Advance to be Recovered',
      'Postponement Granted',
      'Date on which Total Amount is Recovered',
      'Remarks',
    ];
    const rows = applyFormXIIAPNilToMappedRows(
      [{ 'Name of the Workman': 'A', Remarks: '' }],
      headers,
      FORM_XII_AP_NIL
    );
    expect(rows[0]['Name of the Workman']).toBe('A');
    expect(rows[0]['Amount of Advance Given']).toBe(FORM_XII_AP_NIL);
    expect(rows[0]['Date on which Advance was Given']).toBe(FORM_XII_AP_NIL);
    expect(rows[0]['Purpose(s) for which Advance was Given']).toBe(FORM_XII_AP_NIL);
    expect(rows[0]['No. of Instalments by which Advance to be Recovered']).toBe(FORM_XII_AP_NIL);
    expect(rows[0]['Postponement Granted']).toBe(FORM_XII_AP_NIL);
    expect(rows[0]['Date on which Total Amount is Recovered']).toBe(FORM_XII_AP_NIL);
    expect(rows[0].Remarks).toBe('');
  });
});
