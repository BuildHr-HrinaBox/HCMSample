import {
  isFormXIIIRegisterOfWorkmenContext,
  isFormXAPRegisterOfFinesContext,
  blobIndicatesFormIRegisterOfFinesNotFormX,
  sheetBlobIndicatesFormXAPRegisterOfFines,
} from './formXAPRegisterOfFines';

describe('Form XIII Register of Workmen detection', () => {
  it('does not classify Tamil Nadu Form 1 conferment register as Form XIII', () => {
    expect(
      isFormXIIIRegisterOfWorkmenContext(
        {
          title: 'FORM 1',
          subtitle: 'REGISTER OF WORKMEN',
          reference: 'Tamil Nadu Industrial Establishments (Conferment of Permanent Status) Act'
        },
        { state: 'TamilNadu', formName: 'Register of Conferment Status' },
        'Form_I_-_TamilNadu.xlsx',
        'REGISTER OF WORKMEN'
      )
    ).toBe(false);
  });

  it('still detects CLRA Form XIII Register of Workmen', () => {
    expect(
      isFormXIIIRegisterOfWorkmenContext(
        { title: 'FORM XIII', subtitle: 'Register of Workmen' },
        { state: 'Andhra Pradesh', act: 'Contract Labour' },
        'Form_XIII.xlsx',
        'Register of Workmen (Contract Labour)'
      )
    ).toBe(true);
  });
});

describe('Form I Tamil Nadu vs Form X AP Register of Fines', () => {
  it('does not classify Form_I_-_TamilNadu Register of Fines as Form X AP', () => {
    const sheetText =
      'FORM I REGISTER OF FINES [See Rule 3] Payment of Wages Act Minimum Wages Act Name of the Establishment';
    expect(
      blobIndicatesFormIRegisterOfFinesNotFormX(
        `Form I Register of Fines Form_I_-_TamilNadu.xlsx ${sheetText}`
      )
    ).toBe(true);
    expect(
      sheetBlobIndicatesFormXAPRegisterOfFines(
        `Form I Register of Fines Form_I_-_TamilNadu.xlsx ${sheetText}`
      )
    ).toBe(false);
    expect(
      isFormXAPRegisterOfFinesContext(
        { title: 'FORM I', subtitle: 'REGISTER OF FINES', reference: '[See Rule 3]' },
        { formName: 'Form I', state: 'Tamil Nadu' },
        'Form_I_-_TamilNadu.xlsx',
        sheetText
      )
    ).toBe(false);
  });

  it('still detects AP Shops Form X Register of Fines', () => {
    expect(
      isFormXAPRegisterOfFinesContext(
        { title: 'Form X', subtitle: 'Register of Fines' },
        { formName: 'Form X', state: 'Andhra Pradesh', act: 'Shops and Establishment' },
        'Form_X_AP.xlsx',
        'Form X Register of Fines Name of the worker Nature & date of offence'
      )
    ).toBe(true);
  });
});
