import { statutoryDraftPdfTestUtils } from './statutoryDraftPdf';

const {
  buildStatutoryPdfHeaderModel,
  looksLikeFormBTamilNaduPdfContext,
  isFormBTamilNaduPdfSignatoryRow,
  detectFormBTamilNaduPdfGroupBands,
  normalizeFormBTamilNaduPdfMatrix,
} = statutoryDraftPdfTestUtils;

describe('Form B Tamil Nadu LWF PDF header', () => {
  const meta = [
    'Form B',
    '(See rule 29)',
    'Tamilnadu Labour Welfare Fund Rules, 1973',
    'Register of Wages',
    'Vayona Energy Pvt Ltd, 274 A, Rani mangammal main road, Govindanagaram - 625517 Theni Taluk, Theni District, Tamilnadu',
    'July2026',
    'For (Company Name)',
    'Authorised Signatory',
    'Signature of Employer / Manager / Authorised Person',
  ];
  const rows = [
    [
      'Total number of employees',
      'Total emoluments payable during the month including basic wages, D.A, O.T., and bonus',
      'Fine',
      'Other deductions',
      'Amount actually paid during the month',
      'Balance due to the employees',
    ],
    ['15', '1163466', '', '458439', '1035636', '0'],
    ['For (Company Name)'],
    ['Authorised Signatory'],
    ['Signature of Employer / Manager / Authorised Person'],
  ];

  it('detects Form B Tamil Nadu LWF register', () => {
    expect(
      looksLikeFormBTamilNaduPdfContext(meta, rows, 'Form B', 'Form_B_-_TamilNadu.xlsx')
    ).toBe(true);
  });

  it('prints establishment and month headings and drops signatory lines', () => {
    const model = buildStatutoryPdfHeaderModel(meta, rows, 0, 'Form B', {
      fileName: 'Form_B_-_TamilNadu.xlsx',
    });
    expect(model.formBTamilNadu).toBe(true);
    expect(model.titles.some((t) => /^Form B$/i.test(t))).toBe(true);
    expect(model.titles.some((t) => /See rule 29/i.test(t))).toBe(true);
    expect(model.titles.some((t) => /Labour Welfare Fund Rules/i.test(t))).toBe(true);
    expect(model.titles.some((t) => /Register of Wages/i.test(t))).toBe(true);
    expect(model.fields.some((f) => /Name of the Establishment\s*:/i.test(f))).toBe(true);
    expect(model.fields.some((f) => /For the Month of\s*:\s*July 2026/i.test(f))).toBe(true);
    expect(model.fields.join('\n')).not.toMatch(/Authorised Signatory/i);
    expect(model.fields.join('\n')).not.toMatch(/For\s*\(/);
    expect(model.fields.join('\n')).not.toMatch(/Signature of Employer/i);
    expect(model.titles.join('\n')).not.toMatch(/Authorised Signatory/i);
  });

  it('keeps Amounts deducted during the month as a two-column group banner', () => {
    const headerRows = [
      [
        'Total number of employees',
        'Total emoluments payable during the month including basic wages, D.A, O.T., and bonus',
        'Amounts deducted during the month',
        '',
        'Amount actually paid during the month',
        'Balance due to the employees',
      ],
      ['', '', 'Fine', 'Other deductions', '', ''],
    ];
    const bands = detectFormBTamilNaduPdfGroupBands(headerRows, 0, 1, 6);
    expect(bands).toHaveLength(1);
    expect(bands[0].label).toMatch(/Amounts deducted during the month/i);
    expect(bands[0].start).toBe(2);
    expect(bands[0].end).toBe(3);
  });

  it('treats footer signatory rows as skippable', () => {
    expect(isFormBTamilNaduPdfSignatoryRow(['For (Vayona Energy Pvt Ltd)'])).toBe(true);
    expect(isFormBTamilNaduPdfSignatoryRow(['Authorised Signatory'])).toBe(true);
    expect(
      isFormBTamilNaduPdfSignatoryRow(['Signature of Employer / Manager / Authorised Person'])
    ).toBe(true);
  });

  it('keeps Fine blank and Other deductions in the Other column', () => {
    const normalized = normalizeFormBTamilNaduPdfMatrix(
      [
        ...rows.slice(0, 1),
        ['15', '1163466', '458439', '', '1035636', '0'],
      ],
      6,
      0,
      meta
    );
    expect(normalized.rows[1][2]).toBe('');
    expect(String(normalized.rows[1][3])).toBe('458439');
  });
});
