import {
  FORM_XXVI_TN_PDF_REFERENCE,
  FORM_XXVI_TN_PDF_SUBTITLE,
  FORM_XXVI_TN_PDF_TITLE,
  buildFormXXVITamilNaduPdfHtml,
  extractFormXXVITamilNaduPdfHeader,
  findFormXXVITamilNaduPdfTableStart,
  looksLikeFormXXVITamilNaduPdfContext
} from './statutoryDraftPdf.formXXVI.TN';

describe('Form XXVI Tamil Nadu SmartBrowz PDF', () => {
  test('detects CLRA Form XXVI from filename and official heading', () => {
    expect(
      looksLikeFormXXVITamilNaduPdfContext({
        fileName: 'Form_26_-_TamilNadu.xlsx',
        title: 'FORM XXVI',
        rows: [
          ['FORM XXVI'],
          ['See Rule 75 of the Tamil Nadu Contract Labour (Regulation and Abolition) Rules, 1975'],
          ['Register of Employment of Contractual Labour']
        ]
      })
    ).toBe(true);
  });

  test('does not treat Form XXVII or AP appointment letter as Form XXVI', () => {
    expect(
      looksLikeFormXXVITamilNaduPdfContext({
        fileName: 'Form_XXVII_-_TamilNadu.xlsx',
        title: 'FORM XXVII'
      })
    ).toBe(false);
    expect(
      looksLikeFormXXVITamilNaduPdfContext({
        fileName: 'Form_XXVI_AP.xlsx',
        title: 'Letter of Appointment',
        rows: [['Letter of Appointment']]
      })
    ).toBe(false);
  });

  test('HTML uses official Rule 75 banner and extracted header fields', () => {
    const rows = [
      ['FORM XXVI'],
      ['See Rule 75 of the Tamil Nadu Contract Labour (Regulation and Abolition) Rules, 1975'],
      ['Register of Employment of Contractual Labour'],
      ['Name and address of Principal Employer', 'Acme Industries', 'Month', 'April'],
      ['Name and Address of Contractor.', 'Vayona Contracting', 'Date', '2026'],
      ['Nature and location of work.', 'Chennai Site'],
      [
        'Serial Number',
        'Name of the Workman',
        'Age and Sex',
        'Rate of Wages',
        'Daily hours of work',
        '',
        'Number of Days Worked'
      ],
      ['', '', '', '', '1', '2', ''],
      ['1', 'Ravi', '32 / M', '500', '8', '8', '2']
    ];

    expect(findFormXXVITamilNaduPdfTableStart(rows)).toBe(6);
    const header = extractFormXXVITamilNaduPdfHeader(rows, 6);
    expect(header.title).toBe(FORM_XXVI_TN_PDF_TITLE);
    expect(header.reference).toBe(FORM_XXVI_TN_PDF_REFERENCE);
    expect(header.subtitle).toBe(FORM_XXVI_TN_PDF_SUBTITLE);
    expect(header.principal).toBe('Acme Industries');
    expect(header.contractor).toBe('Vayona Contracting');
    expect(header.worksite).toBe('Chennai Site');
    expect(header.month).toBe('April');
    expect(header.date).toBe('2026');

    const html = buildFormXXVITamilNaduPdfHtml({
      rows,
      fileName: 'Form_26_-_TamilNadu.xlsx'
    });
    expect(html).toContain(FORM_XXVI_TN_PDF_TITLE);
    expect(html).toContain(FORM_XXVI_TN_PDF_REFERENCE);
    expect(html).toContain(FORM_XXVI_TN_PDF_SUBTITLE);
    expect(html).toContain('Acme Industries');
    expect(html).toContain('Ravi');
    expect(html).toContain('class="xxvi-grid"');
  });
});
