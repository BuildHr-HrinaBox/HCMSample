import { statutoryDraftPdfTestUtils } from './statutoryDraftPdf';
import {
  detectFormLGJGujaratPdfGroupBands,
  formLGJGujaratColumnWeight,
  looksLikeFormLGJGujaratPdfContext,
  trimFormLGJGujaratLeadingBlankPdfColumns,
} from './statutoryDraftPdf.formLGJ.GJ';

describe('Form L GJ PDF leading blank column', () => {
  it('detects Form L Gujarat PDF context', () => {
    expect(
      looksLikeFormLGJGujaratPdfContext(
        ['FORM - L', '(See rule 14)', 'LIST OF WORKERS ENGAGED IN SHIFT'],
        [['Sr. No.', 'Name of the Worker', 'Designation']],
        'Form_L_GJ_-_Gujarat'
      )
    ).toBe(true);
  });

  it('does not treat Form K Gujarat as Form L', () => {
    expect(
      looksLikeFormLGJGujaratPdfContext(
        ['FORM - K', '(See rule 12)', 'NOTICE OF WEEKLY HOLIDAY'],
        [['Sr. No. (1)', 'Name of Worker (2)']],
        'Form_K_GJ_-_Gujarat'
      )
    ).toBe(false);
  });

  it('removes the leading blank column before Sr. No.', () => {
    const rows = [
      ['FORM - L', '', '', '', '', '', '', ''],
      ['LIST OF WORKERS ENGAGED IN SHIFT', '', '', '', '', '', '', ''],
      [
        '',
        'Sr. No.',
        'Name of the Worker',
        'Designation',
        'Date of the Month',
        'Date of the Month',
        'Date of the Month',
        'Weekly holiday day',
      ],
      ['', '', '', '', '1st Shift', '2nd Shift', '3rd Shift', ''],
      ['', '1', 'harshad dk', 'Engineer', '9 AM\n6 PM', '', '', 'Saturday & Sunday'],
    ];
    const out = trimFormLGJGujaratLeadingBlankPdfColumns(rows, 8, 2);
    expect(out.colCount).toBe(7);
    expect(out.rows[2]).toEqual([
      'Sr. No.',
      'Name of the Worker',
      'Designation',
      'Date of the Month',
      'Date of the Month',
      'Date of the Month',
      'Weekly holiday day',
    ]);
    expect(out.rows[4]).toEqual([
      '1',
      'harshad dk',
      'Engineer',
      '9 AM\n6 PM',
      '',
      '',
      'Saturday & Sunday',
    ]);
  });
});

describe('Form L GJ PDF heading width', () => {
  it('merges Date of the Month into one wide heading box', () => {
    const rows = [
      [
        'Sr. No.',
        'Name of the Worker',
        'Designation',
        'Date of the Month',
        'Date of the Month',
        'Date of the Month',
        'Weekly holiday day',
      ],
      ['', '', '', '1st Shift', '2nd Shift', '3rd Shift', ''],
      ['', '', '', 'From - To -', 'From - To -', 'From - To -', ''],
    ];
    const bands = detectFormLGJGujaratPdfGroupBands(rows, 0, 2, 7);
    expect(bands).toEqual([
      {
        labelRow: 0,
        start: 3,
        end: 5,
        label: 'Date of the Month',
      },
    ]);
  });

  it('gives shift and weekly-holiday heading columns more weight than Sr. No.', () => {
    expect(formLGJGujaratColumnWeight('Sr. No.')).toBeLessThan(
      formLGJGujaratColumnWeight('1st Shift')
    );
    expect(formLGJGujaratColumnWeight('Weekly holiday day')).toBeGreaterThan(
      formLGJGujaratColumnWeight('1st Shift')
    );
    expect(formLGJGujaratColumnWeight('Name of the Worker')).toBeGreaterThan(
      formLGJGujaratColumnWeight('Designation')
    );
  });

  it('flags Form L Gujarat on the PDF header model', () => {
    const { buildStatutoryPdfHeaderModel } = statutoryDraftPdfTestUtils;
    const model = buildStatutoryPdfHeaderModel(
      ['FORM - L', '(See rule 14)', 'LIST OF WORKERS ENGAGED IN SHIFT'],
      [
        [
          'Sr. No.',
          'Name of the Worker',
          'Designation',
          'Date of the Month',
          '1st Shift',
          'Weekly holiday day',
        ],
      ],
      0,
      'Form_L_GJ_-_Gujarat',
      { fileName: 'Form L GJ - Gujarat.pdf' }
    );
    expect(model.formLGJGujarat).toBe(true);
  });
});

