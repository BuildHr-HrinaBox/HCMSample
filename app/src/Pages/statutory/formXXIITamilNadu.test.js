import {
  FORM_XXII_TN_NIL,
  applyFormXXIITamilNaduNilToMappedRows,
  isFormXXIITamilNaduContext,
  isFormXXIITamilNaduNilDefaultHeader,
} from './formXXIITamilNadu';

describe('Form XXII Tamil Nadu Register of Advances NILL columns', () => {
  it('detects Form_XXII_-_TamilNadu.xlsx context', () => {
    expect(
      isFormXXIITamilNaduContext(
        { title: 'FORM XXII', subtitle: 'Register of Advances' },
        { state: 'Tamil Nadu', formName: 'Form XXII' },
        'Form_XXII_-_TamilNadu.xlsx'
      )
    ).toBe(true);
    expect(
      isFormXXIITamilNaduContext({}, null, 'Form_XXII_-_TamilNadu.xlsx')
    ).toBe(true);
    expect(
      isFormXXIITamilNaduContext(
        { title: 'FORM XXIX' },
        { state: 'Tamil Nadu' },
        'Form_XXIX_-_TamilNadu.xlsx'
      )
    ).toBe(false);
    expect(
      isFormXXIITamilNaduContext(
        { title: 'FORM XXII' },
        { state: 'Andhra Pradesh' },
        'Form_XXII_-_Andhra_Pradesh.xlsx'
      )
    ).toBe(false);
  });

  it('matches the six advance NILL default headers', () => {
    expect(isFormXXIITamilNaduNilDefaultHeader('Wage period and wages payable')).toBe(true);
    expect(isFormXXIITamilNaduNilDefaultHeader('Date and amount of advance given')).toBe(true);
    expect(isFormXXIITamilNaduNilDefaultHeader('Purpose (s) for which advance made')).toBe(true);
    expect(
      isFormXXIITamilNaduNilDefaultHeader('No. of installments by which advance to be repaid')
    ).toBe(true);
    expect(
      isFormXXIITamilNaduNilDefaultHeader('Date and amount of each installment repaid')
    ).toBe(true);
    expect(
      isFormXXIITamilNaduNilDefaultHeader('Date on which last installment was repaid')
    ).toBe(true);
    expect(isFormXXIITamilNaduNilDefaultHeader('Name of Workman')).toBe(false);
    expect(isFormXXIITamilNaduNilDefaultHeader('Remarks')).toBe(false);
  });

  it('applies NILL to advance columns', () => {
    const headers = [
      'Name of Workman',
      'Wage period and wages payable',
      'Date and amount of advance given',
      'Purpose (s) for which advance made',
      'No. of installments by which advance to be repaid',
      'Date and amount of each installment repaid',
      'Date on which last installment was repaid',
      'Remarks',
    ];
    const rows = applyFormXXIITamilNaduNilToMappedRows(
      [{ 'Name of Workman': 'A', Remarks: '' }],
      headers,
      FORM_XXII_TN_NIL
    );
    expect(rows[0]['Name of Workman']).toBe('A');
    expect(rows[0]['Wage period and wages payable']).toBe(FORM_XXII_TN_NIL);
    expect(rows[0]['Date and amount of advance given']).toBe(FORM_XXII_TN_NIL);
    expect(rows[0]['Purpose (s) for which advance made']).toBe(FORM_XXII_TN_NIL);
    expect(rows[0]['No. of installments by which advance to be repaid']).toBe(FORM_XXII_TN_NIL);
    expect(rows[0]['Date and amount of each installment repaid']).toBe(FORM_XXII_TN_NIL);
    expect(rows[0]['Date on which last installment was repaid']).toBe(FORM_XXII_TN_NIL);
    expect(rows[0].Remarks).toBe('');
  });
});
