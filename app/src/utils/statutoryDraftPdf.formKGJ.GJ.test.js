import { statutoryDraftPdfTestUtils } from './statutoryDraftPdf';
import {
  FORM_KGJ_GJ_WEEKLY_HOLIDAY_INTRO,
  isFormKGJHolidayCaptionText,
  isFormKGJTruncatedWeeklyIntroText,
  looksLikeFormKGJGujaratPdfContext,
  rewriteFormKGJGujaratPdfHeader,
  sanitizeFormKGJGujaratPdfDataRows,
  trimFormKGJGujaratLeadingBlankPdfColumns,
} from './statutoryDraftPdf.formKGJ.GJ';

const { buildStatutoryPdfHeaderModel } = statutoryDraftPdfTestUtils;

describe('Form K GJ PDF weekly-holiday caption', () => {
  const truncated =
    'All the workers in the establishment are hereby informed that the weekly';
  const shortCaption = 'Holiday for each worker is given below:';

  it('detects the truncated title and short caption', () => {
    expect(isFormKGJTruncatedWeeklyIntroText(truncated)).toBe(true);
    expect(isFormKGJHolidayCaptionText(shortCaption)).toBe(true);
    expect(isFormKGJTruncatedWeeklyIntroText(FORM_KGJ_GJ_WEEKLY_HOLIDAY_INTRO)).toBe(
      false
    );
    expect(isFormKGJHolidayCaptionText(FORM_KGJ_GJ_WEEKLY_HOLIDAY_INTRO)).toBe(false);
  });

  it('detects Form K Gujarat PDF context', () => {
    expect(
      looksLikeFormKGJGujaratPdfContext(
        ['FORM - K', '(See rule 12)', 'NOTICE OF WEEKLY HOLIDAY', truncated],
        [[shortCaption], ['Sr. No. (1)', 'Name of Worker (2)']],
        'Form_K_GJ_-_Gujarat'
      )
    ).toBe(true);
  });

  it('moves the full sentence above the table and drops the truncated title', () => {
    const rewritten = rewriteFormKGJGujaratPdfHeader(
      ['FORM - K', '(See rule 12)', 'NOTICE OF WEEKLY HOLIDAY', truncated],
      [
        'Address of the Establishment : Site 1',
        'Name of the Manager/Incharge : Manager',
        shortCaption,
      ]
    );

    expect(rewritten.titles.some((t) => /all the workers/i.test(t))).toBe(false);
    expect(rewritten.titles).toEqual([
      'FORM - K',
      '(See rule 12)',
      'NOTICE OF WEEKLY HOLIDAY',
    ]);
    expect(rewritten.fields[rewritten.fields.length - 1]).toBe(
      FORM_KGJ_GJ_WEEKLY_HOLIDAY_INTRO
    );
    expect(rewritten.fields.filter((f) => /all the workers/i.test(f))).toHaveLength(1);
    expect(rewritten.fields.some((f) => /^holiday for each worker/i.test(f))).toBe(
      false
    );
  });

  it('builds the PDF header model without the truncated NOTICE line', () => {
    const model = buildStatutoryPdfHeaderModel(
      [
        'FORM - K',
        '(See rule 12)',
        'NOTICE OF WEEKLY HOLIDAY',
        truncated,
        'Address of the Establishment : No.114/1 Khirai Patiya',
        'Name of the Manager/Incharge : Nilakantan Govinden',
        shortCaption,
      ],
      [
        ['Sr. No. (1)', 'Name of Worker (2)', 'Designation (3)'],
        ['1', 'Harshad Dk', 'Engineer'],
      ],
      0,
      'Form_K_GJ_-_Gujarat',
      { fileName: 'Form K GJ - Gujarat.pdf' }
    );

    expect(model.titles.some((t) => /all the workers/i.test(t))).toBe(false);
    expect(model.titles.some((t) => /NOTICE OF WEEKLY HOLIDAY/i.test(t))).toBe(true);
    expect(model.fields.some((f) => /all the workers in the establishment/i.test(f))).toBe(
      true
    );
    expect(
      model.fields.some((f) => /^holiday for each worker is given below/i.test(f))
    ).toBe(false);
    expect(
      model.fields[model.fields.length - 1]
    ).toBe(FORM_KGJ_GJ_WEEKLY_HOLIDAY_INTRO);
  });

  it('removes the leading blank column before Sr. No.', () => {
    const rows = [
      ['FORM - K', '', '', '', '', ''],
      ['NOTICE OF WEEKLY HOLIDAY', '', '', '', '', ''],
      ['', 'Sr. No. (1)', 'Name of Worker (2)', 'Designation (3)', 'Day of Weekly Holiday (4)', 'Hours of Work (5)'],
      ['', '1', 'Harshad Dk', 'Engineer', 'Saturday & Sunday', ''],
    ];
    const out = trimFormKGJGujaratLeadingBlankPdfColumns(rows, 6, 2);
    expect(out.colCount).toBe(5);
    expect(out.rows[2]).toEqual([
      'Sr. No. (1)',
      'Name of Worker (2)',
      'Designation (3)',
      'Day of Weekly Holiday (4)',
      'Hours of Work (5)',
    ]);
    expect(out.rows[3]).toEqual([
      '1',
      'Harshad Dk',
      'Engineer',
      'Saturday & Sunday',
      '',
    ]);
  });

  it('clears weekly-holiday placeholder text from the Hours of Work PDF column', () => {
    const rows = [
      ['Sr. No. (1)', 'Name of Worker (2)', 'Designation (3)', 'Day of Weekly Holiday (4)', 'Hours of Work (5)'],
      ['1', 'Vikash Gupta', 'Engineer', 'Saturday & Sunday', 'Saturday & Sunday'],
      ['1', 'Vikash Gupta', 'Engineer', 'Saturday & Sunday', '09:00 to 18:00'],
    ];
    const out = sanitizeFormKGJGujaratPdfDataRows(rows, 5, 0);
    expect(out[1][4]).toBe('');
    expect(out[2][4]).toBe('09:00 to 18:00');
  });
});
