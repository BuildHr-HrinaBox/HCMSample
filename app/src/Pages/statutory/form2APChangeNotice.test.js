import {
  enrichForm2APDisplayHeader,
  FORM_2AP_DISPLAY_SUBTITLE,
  FORM_2AP_DISPLAY_TITLE,
  isForm2APChangeNoticeContext,
  resolveForm2APHeaderFieldLayout
} from './form2APChangeNotice';

describe('Form 2-A Andhra Pradesh change notice', () => {
  it('detects Form_2A_-_Andhra_Pradesh.xlsx by filename', () => {
    expect(
      isForm2APChangeNoticeContext(
        null,
        { state: 'Andhra Pradesh', formName: 'Form 2-A' },
        'Form_2A_-_Andhra_Pradesh.xlsx'
      )
    ).toBe(true);
  });

  it('forces Form 2-A title even when prior header was Form 11', () => {
    const layout = resolveForm2APHeaderFieldLayout(
      {
        formHeader: {
          title: 'FORM NO. 11',
          subtitle: '(Prescribed under Rule 79) Notice of period of work for adult and children'
        }
      },
      null,
      {
        item: { state: 'Andhra Pradesh' },
        fileName: 'Form_2A_-_Andhra_Pradesh.xlsx'
      }
    );
    expect(layout).not.toBeNull();
    expect(layout.formHeader.title).toBe(FORM_2AP_DISPLAY_TITLE);
    expect(layout.formHeader.subtitle).toBe(FORM_2AP_DISPLAY_SUBTITLE);
    expect(layout.formHeader.form2APColumnBoxLayout).toBe(true);
  });

  it('enrichForm2APDisplayHeader replaces Form 11 display title', () => {
    const enriched = enrichForm2APDisplayHeader(
      {
        title: 'FORM NO. 11',
        subtitle: '(Prescribed under Rule 79) Notice of period of work for adult and children',
        form2APColumnBoxLayout: true,
        fields: [{ label: 'Licence Number', key: 'form2_ap_licence_number', value: '' }]
      },
      'Form_2A_-_Andhra_Pradesh.xlsx',
      { state: 'Andhra Pradesh' }
    );
    expect(enriched.title).toBe(FORM_2AP_DISPLAY_TITLE);
    expect(enriched.subtitle).toBe(FORM_2AP_DISPLAY_SUBTITLE);
  });
});
