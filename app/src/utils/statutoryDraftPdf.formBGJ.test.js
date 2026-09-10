import { statutoryDraftPdfTestUtils } from './statutoryDraftPdf';

const {
  buildStatutoryPdfHeaderModel,
  looksLikeFormAGJGujaratPdfContext,
  looksLikeFormBGJGujaratPdfContext,
  looksLikeFormBRajasthanPdfContext,
  extractFormBGJMinimumWagesBox,
  normalizeFormBGJGujaratPdfMatrix,
  reorderFormBGJPdfHeaderFields,
  FORM_AGJ_GJ_AGE_NOTE,
} = statutoryDraftPdfTestUtils;

describe('Form B GJ PDF Excel model', () => {
  const formBMeta = [
    'FORM B',
    'FORMAT OF WAGE REGISTER',
    'Rate of Minimum Wages and Since the date 01-Oct-2018 to 31-Mar-2019',
    'Name of Establishment: Amreli EDF Nanikundal Site, 66/33Kv Substation Amreli Renewable Energy Private Limited,Village:-Nanikundal,Tall.Babra,Pin code:-365421, Amreli, Gujarat',
    'Name of Owner',
    'Labour Identification No.',
    'Name and address of Principal Employer: Amreli Renewable Energy Pvt Ltd, 66/33Kv Substation Amreli Renewable Energy Private Limited,Village:-Nanikundal,Tall.Babra,Pin code:-365421, Chennai, Gujarat',
    '*Labour Identification No. Of Principal Employer',
    'Wage period From 01-07-2019 To 31-07-2019',
  ];
  const formBRows = [
    ['Highly Skilled', 'Skilled', 'Semi Skilled', 'Un Skilled'],
    ['Minimum Basic', '', '', '', ''],
    ['DA', '', '', '', ''],
    ['Overtime', '', '', '', ''],
    [
      'Sr.No. in Employer/ Workman/ Worker Register',
      'Name',
      'Rate of Wage',
      'No. Of Days worked',
      'Basic',
      'HRA',
      'Total',
      'Net Payment',
      'Receipt by Employee/ Workman/ Worker Bank Transaction ID',
    ],
  ];

  it('detects Form B Gujarat wage-register PDF context', () => {
    expect(looksLikeFormBGJGujaratPdfContext(formBMeta, formBRows, 'Form_B_GJ')).toBe(true);
  });

  it('does not treat Form B GJ as Form A GJ (no Form A footnotes)', () => {
    expect(looksLikeFormAGJGujaratPdfContext(formBMeta, formBRows, 'Form_B_GJ_-_Gujarat')).toBe(
      false
    );
  });

  it('does not treat Form B GJ as Rajasthan Form B', () => {
    expect(looksLikeFormBRajasthanPdfContext(formBMeta, formBRows, 'Form_B_GJ')).toBe(false);
  });

  it('extracts Rate of Minimum Wages box like Excel', () => {
    const box = extractFormBGJMinimumWagesBox(formBMeta, formBRows, 4);
    expect(box.title).toMatch(/Rate of Minimum Wages/i);
    expect(box.skills.join(' ')).toMatch(/Highly Skilled/i);
    expect(box.rows.some((r) => /Minimum Basic/i.test(r.label))).toBe(true);
    expect(box.rows.some((r) => /^DA$/i.test(r.label))).toBe(true);
    expect(box.rows.some((r) => /Overtime/i.test(r.label))).toBe(true);
  });

  it('keeps Form A footnotes out of Form B header model', () => {
    const model = buildStatutoryPdfHeaderModel(
      [...formBMeta, FORM_AGJ_GJ_AGE_NOTE, '***Wherever applicable'],
      formBRows,
      4,
      'Form_B_GJ'
    );
    expect(model.formBGJGujarat).toBe(true);
    expect(model.formBGJMinWagesBox?.title).toMatch(/Rate of Minimum Wages/i);
    expect(model.titles.some((t) => /FORM B/i.test(t))).toBe(true);
    expect(model.fields.some((f) => /Name of Establishment/i.test(f))).toBe(true);
    expect(model.fields.some((f) => /14\s*to\s*18/i.test(f))).toBe(false);
    expect(model.fields.some((f) => /rate\s+of\s+minimum\s+wages/i.test(f))).toBe(false);
  });

  it('keeps establishment header fields outside min-wages box (below in paint)', () => {
    const model = buildStatutoryPdfHeaderModel(formBMeta, formBRows, 4, 'Form_B_GJ');
    expect(model.fields.some((f) => /Name of Owner/i.test(f))).toBe(true);
    expect(model.fields.some((f) => /Labour Identification No/i.test(f))).toBe(true);
    expect(model.fields.some((f) => /Principal Employer/i.test(f))).toBe(true);
    expect(model.fields.some((f) => /Wage period/i.test(f))).toBe(true);
    expect(model.formBGJMinWagesBox).toBeTruthy();
    expect(model.titles.some((t) => /Name of Owner/i.test(t))).toBe(false);
  });

  it('promotes min-wages + establishment rows out of the wage table start', () => {
    const excelLikeRows = [
      ['Highly Skilled', 'Skilled', 'Semi-Skilled', 'Un-Skilled'],
      ['Minimum Basic', '', '', '', ''],
      ['DA', '', '', '', ''],
      ['Overtime', '', '', '', ''],
      [
        'Name of Establishment: Amreli EDF Nanikundal Site, Village:-Nanikundal, Pin code:-365421',
      ],
      ['Name of Owner'],
      ['Labour Identification No.'],
      [
        'Name and address of Principal Employer: Amreli Renewable Energy Pvt Ltd, Chennai, Gujarat',
      ],
      ['*Labour Identification No. Of Principal Employer'],
      ['Wage period From 01-07-2019 To 31-07-2019'],
      [
        'Sr.No. in Employer/ Workman/ Worker Register',
        'Name',
        'Rate of Wage',
        'Net Payment',
      ],
    ];
    const normalized = normalizeFormBGJGujaratPdfMatrix(
      excelLikeRows,
      9,
      0,
      [
        'FORM B',
        'FORMAT OF WAGE REGISTER',
        'Rate of Minimum Wages and Since the date 01-Oct-2018 to 31-Mar-2019',
      ]
    );
    expect(normalized.tableStartRow).toBe(10);
    expect(normalized.metaLines.some((l) => /Name of Establishment/i.test(l))).toBe(true);
    expect(normalized.metaLines.some((l) => /Wage period From/i.test(l))).toBe(true);
    const model = buildStatutoryPdfHeaderModel(
      normalized.metaLines,
      normalized.rows,
      normalized.tableStartRow,
      'Form_B_GJ_-_Gujarat'
    );
    expect(model.formBGJGujarat).toBe(true);
    expect(model.fields.some((f) => /Amreli EDF/i.test(f))).toBe(true);
    expect(model.fields.some((f) => /Principal Employer/i.test(f))).toBe(true);
    expect(model.fields.some((f) => /Wage period From/i.test(f))).toBe(true);
    const ordered = reorderFormBGJPdfHeaderFields(model.fields);
    expect(ordered[0]).toMatch(/Name of Establishment/i);
  });
});
