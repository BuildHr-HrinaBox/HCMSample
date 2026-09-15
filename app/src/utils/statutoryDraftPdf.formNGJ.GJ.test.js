import * as XLSX from 'xlsx';
import {
  looksLikeFormNGJGujaratPdfContext,
  normalizeFormNGJGujaratPdfMatrix,
} from './statutoryDraftPdf.formNGJ.GJ';
import { looksLikeFormLGJGujaratPdfContext } from './statutoryDraftPdf.formLGJ.GJ';
import { statutoryDraftPdfTestUtils } from './statutoryDraftPdf';

const { sheetToDenseMatrix } = statutoryDraftPdfTestUtils;

const excelLikeRows = [
  ['', '', '', '', 'FORM - N'],
  ['', '', '', '', '(See rule 17)'],
  ['', '', '', '', 'LEAVE BOOK'],
  [],
  [
    'Name of the establishment: Maliya Site\nNo.114/1,Khirai Patiya behind Kerosene Depot,Taluka-Maliya,Dist-Morbi GJ - 3636370, Gujarat\nName of the worker: Patel Chandrakant Virabhai\nDescription of the Department\n(if applicable): Site Operations',
    '',
    '',
    '',
    'Name of the employer: VAYONA ENERGY PRIVATE LIMITED\nDate of entry into service: 01 Dec 2025\n(Signature or thumb impression of worker)',
  ],
  [
    'Accumulation of leave',
    '',
    'Leave allowed',
    'Payment for leave made on',
    '',
    'Refusal of leave',
    '',
    'Payment for Leave on discharge of an worker quitting employment if admissible',
    '',
    '',
    'Remarks',
  ],
  ['1', '2', '3', '4', '', '5', '', '6', '', '', '7'],
  [
    'Leave due on',
    'No. of days',
    'From - To -',
    '1st Moiety',
    '2nd Moiety',
    'Application Date',
    'Date of Refusal',
    'Date of discharge',
    'Date and amount paid',
    'Signature or thumb impression of worker',
    'Remarks',
  ],
  ['', '2', '04-Sep-2026 - 05-Sep-2026', '', '', '', '', '', '', '', ''],
  ['DETAILS OF FESTIVAL LEAVE'],
  [
    'Period',
    '',
    'Total Leave',
    'Availed Leave',
    'Balance Leave',
    'Payment made in lieu of Festival Leave, when called for work',
    'Remarks',
  ],
  ['From', 'To', '', '', '', '', ''],
  ['', '', '', '', '', '', ''],
  ['DETAILS OF CASUAL LEAVE'],
  ['Period', '', 'Total Leave', 'Availed Leave', 'Balance Leave', 'Remarks'],
  ['From', 'To', '', '', '', ''],
  ['04-Jul-2026', '05-Jul-2026', '10', '2', '8', ''],
  ['', '', '', '', '', '', '', '', '', 'Name and Signature of Authority'],
];

describe('Form N GJ Leave Book PDF layout', () => {
  test('detects Gujarat Form N leave book and not Form L shift list', () => {
    expect(
      looksLikeFormNGJGujaratPdfContext(
        ['FORM - N', '(See rule 17)', 'LEAVE BOOK'],
        excelLikeRows,
        'Form_N_GJ.xlsx'
      )
    ).toBe(true);
    expect(
      looksLikeFormLGJGujaratPdfContext(
        ['FORM - N', '(See rule 17)', 'LEAVE BOOK'],
        excelLikeRows,
        'Form_N_GJ.xlsx'
      )
    ).toBe(false);
    expect(
      looksLikeFormNGJGujaratPdfContext(
        ['FORM - L', '(See rule 14)', 'LIST OF WORKERS ENGAGED IN SHIFT'],
        [['Sr. No.', 'Name of the Worker', '1st Shift']],
        'Form_L_GJ.xlsx'
      )
    ).toBe(false);
  });

  test('rebuilds identity boxes and three separate leave sections from Excel rows', () => {
    const normalized = normalizeFormNGJGujaratPdfMatrix(excelLikeRows, 11, 0, []);
    expect(normalized.formNGJLayout).toBe(true);
    expect(normalized.formNGJModel.titles[0]).toMatch(/FORM\s*-?\s*N/i);
    expect(normalized.formNGJModel.identity.establishment).toMatch(/Maliya Site/i);
    expect(normalized.formNGJModel.identity.establishment).toMatch(/Khirai Patiya/i);
    expect(normalized.formNGJModel.identity.workerName).toMatch(/Patel Chandrakant Virabhai/i);
    expect(normalized.formNGJModel.identity.department).toMatch(/Site Operations/i);
    expect(normalized.formNGJModel.identity.employer).toMatch(/VAYONA ENERGY PRIVATE LIMITED/i);
    expect(normalized.formNGJModel.identity.dateOfEntry).toMatch(/01 Dec 2025/i);
    expect(normalized.formNGJModel.identity.employer).not.toMatch(/01 Dec 2025/);

    expect(normalized.formNGJModel.leaveValues.leaveDueOn).toBe('');
    expect(normalized.formNGJModel.leaveValues.noOfDays).toBe('2');
    expect(normalized.formNGJModel.leaveValues.leaveFromTo).toBe('04-Sep-2026 - 05-Sep-2026');

    expect(normalized.formNGJModel.casualValues.casualPeriodFrom).toBe('04-Jul-2026');
    expect(normalized.formNGJModel.casualValues.casualPeriodTo).toBe('05-Jul-2026');
    expect(normalized.formNGJModel.casualValues.casualTotal).toBe('10');
    expect(normalized.formNGJModel.casualValues.casualAvailed).toBe('2');
    expect(normalized.formNGJModel.casualValues.casualBalance).toBe('8');
    expect(normalized.formNGJModel.footer).toMatch(/Name and Signature of Authority/i);
  });

  test('does not treat festival or casual titles as accumulation data', () => {
    const mashedRows = [
      ['FORM - N'],
      ['LEAVE BOOK'],
      ['Name of the establishment: Maliya Site'],
      ['Name of the worker: Patel Chandrakant Virabhai'],
      ['Accumulation of leave', 'Leave allowed', 'Payment for leave made on'],
      ['Leave due on', 'No. of days', 'From - To -'],
      ['DETAILS OF FESTIVAL LEAVE'],
      ['From', 'To', 'Total Leave', 'Availed Leave', 'Balance Leave', 'Remarks'],
      ['DETAILS OF CASUAL LEAVE'],
      ['From', 'To', 'Total Leave', 'Availed Leave', 'Balance Leave', 'Remarks'],
      ['01-Aug-2026', '02-Aug-2026', '12', '1', '11', ''],
    ];
    const normalized = normalizeFormNGJGujaratPdfMatrix(mashedRows, 6, 0, []);
    expect(normalized.formNGJModel.leaveValues.leaveDueOn).not.toMatch(/FESTIVAL/i);
    expect(normalized.formNGJModel.leaveValues.noOfDays).not.toMatch(/CASUAL/i);
    expect(normalized.formNGJModel.casualValues.casualPeriodFrom).toBe('01-Aug-2026');
    expect(normalized.formNGJModel.casualValues.casualAvailed).toBe('1');
  });

  test('sheetToDenseMatrix marks Form N leave-book layout before the generic grid', () => {
    const ws = XLSX.utils.aoa_to_sheet(excelLikeRows);
    const matrix = sheetToDenseMatrix(ws, 'Form_N_GJ.xlsx');
    expect(matrix.formNGJLayout).toBe(true);
    expect(matrix.formNGJModel.identity.workerName).toMatch(/Patel Chandrakant Virabhai/i);
    expect(matrix.formNGJModel.leaveValues.noOfDays).toBe('2');
  });
});
