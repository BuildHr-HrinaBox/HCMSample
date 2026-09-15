import * as XLSX from 'xlsx';
import { statutoryDraftPdfTestUtils } from './statutoryDraftPdf';
import {
  FORM_P_KA_PDF_DETAILS,
  FORM_P_KA_PDF_RULE,
  FORM_P_KA_PDF_SUBTITLE,
  FORM_P_KA_PDF_TITLE,
  looksLikeFormPKarnatakaPdfContext,
  normalizeFormPKarnatakaPdfMatrix,
} from './statutoryDraftPdf.formP.KA';
import { looksLikeFormPGJGujaratPdfContext } from './statutoryDraftPdf.formPGJ.GJ';
import { looksLikeFormOGJGujaratPdfContext } from './statutoryDraftPdf.formOGJ.GJ';

const { sheetToDenseMatrix } = statutoryDraftPdfTestUtils;

const excelModelAoa = () => [
  [],
  [],
  ["Form – 'P'", '', '', ''],
  ['(See rule 20)', '', '', ''],
  ['NOTICE OF MAXIMUM LEAVE ACCUMULATED', '', '', ''],
  [
    'Name and address of the establishment : Gurmatkal Site, 110/33KV Substation, SF-490/2, Chanderki Village',
    '',
    '',
    '',
  ],
  ['Name of the Authorised person / Manager', '', '', ''],
  ['To,', '', '', ''],
  ['Shri/Smt. Vijeesh Vijayan', '', '', ''],
  [
    'Address: Gurmatkal Site, 110/33KV Substation, SF-490/2, Chanderki Village',
    '',
    '',
    '',
  ],
  [],
  [
    'It is hereby informed that as per section 18 (5) of the Maharashtra Shops and Establishments (Regulation of Employment and Conditions of Service) Act, 2017 (Mah. LXI of 2017) the maximum leave that can be accumulated is for 45 days.',
    '',
    '',
    '',
  ],
  ['Details of the leave accumulated', '', '', ''],
  ['Sr. No.', 'Number of accumulated leave', 'Period for which leave is accumulated', ''],
  ['', '', 'From', 'Till'],
  ['1', '11', '29 May 2026', '10 Jun 2026'],
];

describe('Form P Karnataka PDF Excel model', () => {
  test('detects Form P KA notice and not Gujarat Form P muster or Form O', () => {
    const rows = excelModelAoa();
    expect(
      looksLikeFormPKarnatakaPdfContext([], rows, 'FORM P', 'Form_P_-_Karnataka.xlsx')
    ).toBe(true);
    expect(
      looksLikeFormPGJGujaratPdfContext([], rows, 'FORM P', 'Form_P_-_Karnataka.xlsx')
    ).toBe(false);
    expect(
      looksLikeFormOGJGujaratPdfContext([], rows, 'FORM P', 'Form_P_-_Karnataka.xlsx')
    ).toBe(false);
    expect(
      looksLikeFormPKarnatakaPdfContext([], [['NOTICE']], 'Sheet1', 'Form_P_-_Karnataka.xlsx')
    ).toBe(true);
  });

  test('normalizeFormPKarnatakaPdfMatrix keeps To block above the leave table', () => {
    const normalized = normalizeFormPKarnatakaPdfMatrix(excelModelAoa(), 4, 0, []);
    expect(normalized.formPKALayout).toBe(true);
    expect(normalized.formPKAModel.titles).toEqual([
      FORM_P_KA_PDF_TITLE,
      FORM_P_KA_PDF_RULE,
      FORM_P_KA_PDF_SUBTITLE,
    ]);
    expect(normalized.formPKAModel.establishment).toMatch(/Gurmatkal Site/i);
    expect(normalized.formPKAModel.shriSmt).toMatch(/Shri\/Smt\.\s+Vijeesh Vijayan/i);
    expect(normalized.formPKAModel.address).toMatch(/Gurmatkal Site/i);
    expect(normalized.formPKAModel.details).toBe(FORM_P_KA_PDF_DETAILS);
    const detailsIdx = normalized.rows.findIndex((row) =>
      /details of the leave accumulated/i.test(String(row?.[0] || ''))
    );
    const estIdx = normalized.rows.findIndex((row) =>
      /name and address of the establishment/i.test(String(row?.[0] || ''))
    );
    expect(estIdx).toBeGreaterThanOrEqual(0);
    expect(detailsIdx).toBeGreaterThan(estIdx);
    expect(normalized.formPKAModel.leaveRows[0][1]).toBe('11');
    expect(normalized.formPKAModel.leaveRows[0][2]).toMatch(/29 May 2026/);
    expect(normalized.formPKAModel.leaveRows[0][3]).toMatch(/10 Jun 2026/);
  });

  test('sheetToDenseMatrix uses Form P KA letter layout, not a generic grid', () => {
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(excelModelAoa());
    XLSX.utils.book_append_sheet(wb, ws, 'FORM P');
    const matrix = sheetToDenseMatrix(wb.Sheets['FORM P'], 'FORM P');
    expect(matrix.formPKALayout).toBe(true);
    expect(matrix.formPKAModel.shriSmt).toMatch(/Vijeesh Vijayan/i);
    expect(matrix.colCount).toBe(4);
    const joined = (matrix.rows || []).map((row) => (row || []).join(' ')).join('\n');
    const detailsAt = joined.toLowerCase().indexOf('details of the leave accumulated');
    const estAt = joined.toLowerCase().indexOf('name and address of the establishment');
    expect(estAt).toBeGreaterThanOrEqual(0);
    expect(detailsAt).toBeGreaterThan(estAt);
  });
});
