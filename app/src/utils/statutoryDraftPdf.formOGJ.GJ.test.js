import { statutoryDraftPdfTestUtils } from './statutoryDraftPdf';
import {
  looksLikeFormOGJGujaratPdfContext,
  rewriteFormOGJGujaratPdfHeader,
  trimFormOGJGujaratLeadingBlankPdfColumns,
} from './statutoryDraftPdf.formOGJ.GJ';

const { buildStatutoryPdfHeaderModel } = statutoryDraftPdfTestUtils;

const LEGAL =
  'As per section 18 (5) of the Gujarat Shops and Establishments (Regulation of Employment and Conditions of Service) Act, 2019 (Guj. 4 of 2019) the maximum leave that can be accumulated is for 45 days.';

describe('Form O GJ PDF Excel order and leading blank column', () => {
  it('detects Form O Gujarat PDF context', () => {
    expect(
      looksLikeFormOGJGujaratPdfContext(
        ['FORM - O', '(See rule 18)', 'NOTICE OF MAXIMUM LEAVE ACCUMULATED'],
        [['Sr. No', 'Name of Workers', 'Name of accumulated leave']],
        'Form_O_GJ_-_Gujarat'
      )
    ).toBe(true);
  });

  it('does not treat Form N Gujarat as Form O', () => {
    expect(
      looksLikeFormOGJGujaratPdfContext(
        ['FORM - N', '(See rule 17)', 'LEAVE BOOK'],
        [['Leave due on', 'No. of days']],
        'Form_N_GJ_-_Gujarat'
      )
    ).toBe(false);
  });

  it('reorders header fields to match Excel', () => {
    const rewritten = rewriteFormOGJGujaratPdfHeader(
      [
        'FORM - O',
        '(See rule 18)',
        'NOTICE OF MAXIMUM LEAVE ACCUMULATED',
        'Name of the Authorized person / Manager',
        'Notice',
        'Details of Workers',
      ],
      [
        'Address of the Establishment : No.114/1, Khirai Patiya',
        LEGAL,
      ]
    );

    expect(rewritten.titles).toEqual([
      'FORM - O',
      '(See rule 18)',
      'NOTICE OF MAXIMUM LEAVE ACCUMULATED',
    ]);
    expect(rewritten.fields[0]).toMatch(/Address of the Establishment/i);
    expect(rewritten.fields[1]).toMatch(/Authorized person/i);
    expect(rewritten.fields[2]).toBe('Notice');
    expect(rewritten.fields[3]).toMatch(/maximum leave that can be accumulated/i);
    expect(rewritten.fields[rewritten.fields.length - 1]).toBe('Details of Workers');
  });

  it('builds the PDF header model in Excel field order', () => {
    const model = buildStatutoryPdfHeaderModel(
      [
        'FORM - O',
        '(See rule 18)',
        'NOTICE OF MAXIMUM LEAVE ACCUMULATED',
        'Name of the Authorized person / Manager',
        'Notice',
        'Details of Workers',
        'Address of the Establishment : No.114/1, Khirai Patiya',
        LEGAL,
      ],
      [
        ['Sr. No', 'Name of Workers', 'Name of accumulated leave', 'From', 'Till'],
        ['1', 'Harshad Dk', '', '', ''],
      ],
      0,
      'Form_O_GJ_-_Gujarat',
      { fileName: 'Form_O_GJ_-_Gujarat.pdf' }
    );

    expect(model.titles).toEqual([
      'FORM - O',
      '(See rule 18)',
      'NOTICE OF MAXIMUM LEAVE ACCUMULATED',
    ]);
    expect(model.fields[0]).toMatch(/Address of the Establishment/i);
    expect(model.fields[1]).toMatch(/Authorized person/i);
    expect(model.fields[2]).toBe('Notice');
    expect(model.fields.some((f) => /maximum leave that can be accumulated/i.test(f))).toBe(
      true
    );
    expect(model.fields[model.fields.length - 1]).toBe('Details of Workers');
  });

  it('removes the leading blank column before Sr. No.', () => {
    const rows = [
      ['FORM - O', '', '', '', '', ''],
      ['NOTICE OF MAXIMUM LEAVE ACCUMULATED', '', '', '', '', ''],
      ['', 'Sr. No', 'Name of Workers', 'Name of accumulated leave', 'From', 'Till'],
      ['', '1', 'Harshad Dk', '', '', ''],
      ['', '2', 'Shailesh Babaria', '', '', ''],
    ];
    const out = trimFormOGJGujaratLeadingBlankPdfColumns(rows, 6, 2);
    expect(out.colCount).toBe(5);
    expect(out.rows[2]).toEqual([
      'Sr. No',
      'Name of Workers',
      'Name of accumulated leave',
      'From',
      'Till',
    ]);
    expect(out.rows[3]).toEqual(['1', 'Harshad Dk', '', '', '']);
    expect(out.rows[4]).toEqual(['2', 'Shailesh Babaria', '', '', '']);
  });
});
