import { statutoryDraftPdfTestUtils } from './statutoryDraftPdf';
import { getFormXVIAPHeaderTitles } from './statutoryDraftPdf.formXVI.AP';

const { buildStatutoryPdfHeaderModel } = statutoryDraftPdfTestUtils;

describe('Form XVI AP PDF header alignment', () => {
  it('keeps For the Month of as month/year only when Excel concatenates leftover header text', () => {
    const rows = [
      ['Form XVI - Muster Roll'],
      ['Address of the Establishment : M/S VIBRANT GREENTECH INDIA PRIVATE LIMITED'],
      ['Name and Address of Contractor. : VAYONA ENERGY PRIVATE LIMITED'],
      ['Name and address of Principal Employer : Vibrant greentech India PVT LTD'],
      [
        'Nature and Location of work : AP-Tadipatri',
        'For the Month of : April 2026 Form XVI - Muster Roll Address of the Establishment : M/S VIBRANT GREENTECH Name and Address of Contractor. : VAYONA ENERGY Name and address of Principal Employer : Vibrant'
      ],
      ['S.No', 'Name of the Employee', 'Father\'s/ Husband\'s Name', 'Sex', 'Dates', '1', '2']
    ];
    const model = buildStatutoryPdfHeaderModel([], rows, 5, 'XVI-Muster Roll');
    expect(model.formXVIAP).toBe(true);
    expect(model.rightFields[1]).toBe('For the Month of : April 2026');
    expect(model.rightFields[1]).not.toMatch(/Form XVI|VAYONA|Establishment|Principal Employer/i);
    expect(model.rightFields[0]).toMatch(/AP-Tadipatri/);
    expect(model.rightFields[0]).not.toMatch(/April 2026/);
  });

  it('keeps only Nature/Location and month when the month cell dumps the full header', () => {
    const dump =
      'For the Month of : May 2026 Form XVI - Muster Roll Address of the Establishment : M/S VIBRANT GREENTECH INDIA PRIVATE LIMITED, SY NO.265,Chinthakunta village,Putlur (Mandal),Anantapur(Dist),Andhra Pradesh-515414, Tadipatri, Andhra Pradesh (Vide rule 78 (1) (a) (I) of Contract Labour (Regulation and Abolition ) Central/ A.P Rules. Name and Address of Contractor. : VAYONA ENERGY PRIVATE LIMITED Name and address of Principal Employer : Vibrant greentech India PVT LTD (HCL) & Traditional Customers, M/S VIBRANT GREENTECH INDIA PRIVATE LIMITED, SY NO.265,Chinthakunta village,Putlur (Mandal),Anantapur(Dist),Andhra Pradesh-515414 Nature and Location of work : AP-Tadipatri For the Month of : May 2026';
    const rows = [
      ['Form XVI - Muster Roll'],
      ['Address of the Establishment : M/S VIBRANT GREENTECH INDIA PRIVATE LIMITED, SY NO.265,Chinthakunta village,Putlur (Mandal),Anantapur(Dist),Andhra Pradesh-515414, Tadipatri, Andhra Pradesh'],
      ['Name and Address of Contractor. : VAYONA ENERGY PRIVATE LIMITED'],
      ['Name and address of Principal Employer : Vibrant greentech India PVT LTD (HCL) & Traditional Customers'],
      ['Nature and Location of work : AP-Tadipatri', dump],
      ['S.No', 'Name of the Employee', 'Father\'s/ Husband\'s Name', 'Sex', 'Dates', '1', '2']
    ];
    const model = buildStatutoryPdfHeaderModel([], rows, 5, 'XVI-Muster Roll');
    expect(model.rightFields[0]).toBe('Nature and Location of work : AP-Tadipatri');
    expect(model.rightFields[1]).toBe('For the Month of : May 2026');
    expect(model.rightFields.join(' ')).not.toMatch(/VAYONA|Vide rule|Principal Employer|Establishment/i);
  });

  it('does not duplicate Form XVI - Muster Roll as two title bands', () => {
    const titles = getFormXVIAPHeaderTitles(
      [],
      [['Form XVI - Muster Roll'], ['Vide Rule 78 (1) (a) (i) of Contract Labour (Regulation and Abolition ) Central / A.P. Rules']]
    );
    const formTitleHits = titles.filter((t) => /form[\s._-]*xvi/i.test(t) && /muster\s+roll/i.test(t));
    expect(formTitleHits.length).toBeLessThanOrEqual(1);
  });
});
