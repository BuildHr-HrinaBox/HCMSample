import { getFormXXAPHeaderTitles } from './statutoryDraftPdf.formXX.AP';
import { getFormXXIAPHeaderTitles } from './statutoryDraftPdf.formXXI.AP';
import { formXVIIAPHeaderTextMetrics, getFormXVIIAPHeaderTitles } from './statutoryDraftPdf.formXVII.AP';
import { getFormXVIAPHeaderTitles } from './statutoryDraftPdf.formXVI.AP';
import {
  extractFormXIIIAPAdministrativeRows,
  getFormXIIIAPHeaderTitles
} from './statutoryDraftPdf.formXIII.AP';
import { isApPdfHeadingRow } from './statutoryDraftPdf.apHeaderTitles';
import { statutoryDraftPdfTestUtils } from './statutoryDraftPdf';

const { uniqueExpandedTitleBands, buildStatutoryPdfHeaderModel } = statutoryDraftPdfTestUtils;

const formXXConcatenated =
  'FORM - XX REGISTER OF DEDUCTIONS FOR DAMAGE OR LOSS [Vide Rule 78 (1) (a) (ii) of Contract Labour (Reg. & Abolition) Central & A.P.Rules]';

describe('AP statutory PDF dynamic headings', () => {
  test('Form XX concatenated Excel title is split into unique form / register / rule bands', () => {
    const titles = getFormXXAPHeaderTitles([], [[formXXConcatenated, formXXConcatenated]]);
    expect(titles).toHaveLength(3);
    expect(titles[0]).toMatch(/^form\s*[-._ ]*xx\b/i);
    expect(titles[0]).not.toMatch(/register of deductions/i);
    expect(titles[1]).toMatch(/register of deductions for damage or loss/i);
    expect(titles[1]).not.toMatch(/^form\s/i);
    expect(titles[2]).toMatch(/vide\s+rule\s+78/i);
    expect(uniqueExpandedTitleBands(titles)).toEqual(titles);
  });

  test('Form XX duplicated heading rows do not reprint FORM - XX or the register title', () => {
    const rows = [
      Array(11).fill('FORM - XX'),
      Array(11).fill('REGISTER OF DEDUCTIONS FOR DAMAGE OR LOSS'),
      Array(11).fill('FORM - XX'),
      Array(11).fill('REGISTER OF DEDUCTIONS FOR DAMAGE OR LOSS'),
      Array(11).fill(
        '[Vide Rule 78 (1) (a) (ii) of Contract Labour (Reg. & Abolition) Central & A.P.Rules]'
      )
    ];
    const titles = getFormXXAPHeaderTitles([], rows);
    const painted = uniqueExpandedTitleBands([...titles, ...titles]);
    expect(painted.filter((line) => /^form\s*[-._ ]*xx\b/i.test(line))).toHaveLength(1);
    expect(painted.filter((line) => /register of deductions/i.test(line))).toHaveLength(1);
    expect(painted.filter((line) => /vide\s+rule\s+78/i.test(line))).toHaveLength(1);
  });

  test('Form XXI / XVII / XVI / XIII getters keep a single dynamic heading set', () => {
    expect(getFormXXIAPHeaderTitles([], [['FORM - XXI REGISTER OF FINES [Vide Rule 78 (1) (a) (ii)]']])).toEqual([
      'FORM - XXI',
      'REGISTER OF FINES',
      '[Vide Rule 78 (1) (a) (ii)]'
    ]);
    expect(
      getFormXVIIAPHeaderTitles([], [['FORM XVII REGISTER OF WAGES [Vide Rule 78 (1) (a) (i)]']])
    ).toHaveLength(3);
    expect(getFormXVIAPHeaderTitles([], [['FORM XVI MUSTER ROLL Vide Rule 78 (1) (a) (i)']])).toHaveLength(3);
    expect(
      getFormXIIIAPHeaderTitles(
        [],
        [['FORM XIII REGISTER OF WORKMEN EMPLOYED BY CONTRACTOR [Vide Rule 75]']],
        1
      )
    ).toHaveLength(3);
  });

  test('AP heading rows are skipped so titles are not painted again in the table', () => {
    expect(isApPdfHeadingRow(Array(11).fill('FORM - XX'))).toBe(true);
    expect(isApPdfHeadingRow(Array(11).fill('REGISTER OF DEDUCTIONS FOR DAMAGE OR LOSS'))).toBe(true);
    expect(
      isApPdfHeadingRow([
        'S.No',
        'Name of Workmen',
        "Father's/Husband's Name",
        'Nature of employment/Designation',
        'Date of recovery'
      ])
    ).toBe(false);
  });

  test('Form XX header model uses unique dynamic titles', () => {
    const model = buildStatutoryPdfHeaderModel(
      [],
      [
        Array(11).fill(formXXConcatenated),
        Array(11).fill(formXXConcatenated),
        [
          'S.No',
          'Name of Workmen',
          "Father's/Husband's Name",
          'Nature of employment/Designation',
          'Name of person in whose presence employee\'s explanation was heard',
          'Amount of deduction imposed',
          'Date of damage or loss',
          'Whether worker showed cause against deduction',
          'Name of person in whose presence employee\'s explanation was heard',
          'Date of recovery',
          'Remarks'
        ]
      ],
      2,
      'Form XX'
    );
    expect(model.formXXAP).toBe(true);
    expect(model.titles.filter((line) => /^form\s*[-._ ]*xx\b/i.test(line))).toHaveLength(1);
    expect(model.titles.filter((line) => /register of deductions/i.test(line))).toHaveLength(1);
  });
});

describe('Form XIII AP administrative header', () => {
  test('uses Contractor (company) value for empty establishment-in-which-contract field', () => {
    const rows = [
      ['Name and Address of Contractor. : VAYONA ENERGY PRIVATE LIMITED'],
      ['Name and address of Establishemnt in/ under which contract is carried on:'],
      ['Nature and location of work. : AP-Tadipatri'],
      [
        'Name and address of Principal Employer : Vibrant greentech India PVT LTD (HCL) & Traditional Customers, M/S VIBRANT'
      ]
    ];
    const admin = extractFormXIIIAPAdministrativeRows([], rows, 4);
    expect(admin[0][1]).toMatch(/VAYONA ENERGY PRIVATE LIMITED/i);
    expect(admin[0][1]).not.toMatch(/Vibrant greentech/i);
    expect(admin[1][1]).toMatch(/Vibrant greentech India PVT LTD \(HCL\)/i);
  });
});

describe('Form XVII AP column header vertical fit', () => {
  test('moves wrapped header text up so the last line stays inside the cell', () => {
    const cellHeight = 82;
    const previousFirstBaseline = cellHeight / 2 + 9;
    const { firstBaseline, lastBaseline } = formXVIIAPHeaderTextMetrics({
      lineCount: 5,
      fontSize: 8,
      cellHeight,
      vAlign: 'top'
    });
    expect(firstBaseline).toBeLessThan(previousFirstBaseline);
    expect(lastBaseline).toBeLessThan(cellHeight - 6);
  });
});
