import * as XLSX from 'xlsx';
import { statutoryDraftPdfTestUtils } from './statutoryDraftPdf';
import {
  FORM_F_KA_PDF_PART_I_TITLE,
  FORM_F_KA_PDF_PART_II_TITLE,
  FORM_F_KA_PDF_RULE,
  FORM_F_KA_PDF_SUBTITLE,
  FORM_F_KA_PDF_TITLE,
  looksLikeFormFKarnatakaPdfContext,
  normalizeFormFKarnatakaPdfMatrix,
} from './statutoryDraftPdf.formF.KA';

const { looksLikeForm15TamilNaduPdfContext, sheetToDenseMatrix } = statutoryDraftPdfTestUtils;

const excelModelAoa = () => [
  ['FORM F', '', '', '', '', '', '', '', '', '', ''],
  ['(SEE RULE 8)', '', '', '', '', '', '', '', '', '', ''],
  [],
  ['REGISTER OF LEAVE WITH WAGES', '', '', '', '', '', '', '', '', '', ''],
  [],
  [
    '1',
    'SI No in the Register of Adult/young person',
    'VE0761',
    '',
    '',
    '2',
    'Date of entry into service',
    '01 Dec 2025',
    '',
    '',
    '',
  ],
  [],
  [
    '3',
    'Name of the person',
    'Vijeesh Vijayan',
    '',
    '',
    '4',
    "Father's Name",
    'abc',
    '',
    '',
    '',
  ],
  [],
  ['PART I EARNED LEAVE', '', '', '', '', '', '', '', '', '', ''],
  [],
  [
    'No of Days worked',
    '',
    '',
    'Leave earned',
    'Leave at credit (incl balance if any, on return from leave on last occasion)',
    'leave availed',
    '',
    '',
    'Balance on return from leave',
    'Date on which wages for leave paid and amount paid',
    'Remarks',
  ],
  ['From', 'To', 'Total days worked', '', '', 'From', 'To', 'No. of days', '', '', ''],
  ['1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11'],
  [
    '01-May-2026',
    '31-May-2026',
    '31',
    '11',
    '11',
    '29-May-2026',
    '10-Jun-2026',
    '11',
    '0',
    '29-May-2026 to 10-Jun-2026',
    '',
  ],
  [],
  [],
  [],
  [],
  ['PART II - Sick/Accident Leave (with pay)', '', '', '', '', '', '', '', '', '', ''],
  [],
  ['2026', 'of credit', 'Sick/Accident leave Availed', '0', '', '', '', '', '', '', ''],
  ['1', '2', '3', '4', '', '', '', '', '', '', ''],
];

describe('Form F Karnataka PDF context', () => {
  test('detects Form F leave register Excel and filename', () => {
    expect(
      looksLikeFormFKarnatakaPdfContext(
        ['FORM F', '(SEE RULE 8)', 'REGISTER OF LEAVE WITH WAGES'],
        [['PART I EARNED LEAVE'], ['From', 'To', 'Total days worked']],
        'Sheet1',
        'Form_F_Karnataka.xlsx'
      )
    ).toBe(true);
    expect(
      looksLikeFormFKarnatakaPdfContext([], excelModelAoa(), 'Sheet1', 'Form_F_-_Karnataka.xlsx')
    ).toBe(true);
  });

  test('rejects Form 15 leave register and Form Q appointment order', () => {
    expect(
      looksLikeFormFKarnatakaPdfContext(
        ['FORM 15', 'REGISTER OF LEAVE WITH WAGES'],
        [['PART I'], ['Name of the employee']],
        'Form_15_KA'
      )
    ).toBe(false);
    expect(
      looksLikeFormFKarnatakaPdfContext(
        ['FORM Q', '(See Rule 24(9A))', 'APPOINTMENT ORDER'],
        [['3. Name of the Employee', 'A']],
        'Form_Q_KA'
      )
    ).toBe(false);
  });

  test('does not treat Form F as Tamil Nadu Form 15 leave register', () => {
    expect(looksLikeForm15TamilNaduPdfContext([], excelModelAoa(), 'Form_F_Karnataka')).toBe(false);
  });
});

describe('Form F Karnataka PDF alignment (Excel model)', () => {
  test('rebuilds identity 1–4 beside labels like the Excel register', () => {
    const normalized = normalizeFormFKarnatakaPdfMatrix(excelModelAoa(), 11, 0, []);
    expect(normalized.formFKALayout).toBe(true);
    expect(normalized.metaLines).toEqual([
      FORM_F_KA_PDF_TITLE,
      FORM_F_KA_PDF_RULE,
      FORM_F_KA_PDF_SUBTITLE,
    ]);
    expect(normalized.formFKAModel.identity).toEqual({
      slNo: 'VE0761',
      dateOfEntry: '01 Dec 2025',
      personName: 'Vijeesh Vijayan',
      fatherName: 'abc',
    });
    expect(normalized.rows[0][0]).toBe('1');
    expect(normalized.rows[0][1]).toMatch(/SI No in the Register of Adult/i);
    expect(normalized.rows[0][2]).toBe('VE0761');
    expect(normalized.rows[0][3]).toBe('2');
    expect(normalized.rows[0][5]).toBe('01 Dec 2025');
    expect(normalized.rows[1][0]).toBe('3');
    expect(normalized.rows[1][2]).toBe('Vijeesh Vijayan');
    expect(normalized.rows[1][3]).toBe('4');
    expect(normalized.rows[1][5]).toBe('abc');
  });

  test('keeps PART I 11-column earned-leave row and grouped headers', () => {
    const normalized = normalizeFormFKarnatakaPdfMatrix(excelModelAoa(), 11, 0, []);
    expect(normalized.rows[2][0]).toBe(FORM_F_KA_PDF_PART_I_TITLE);
    const groupRow = normalized.rows[3];
    expect(groupRow[0]).toMatch(/No of Days worked/i);
    expect(groupRow[3]).toMatch(/Leave earned/i);
    expect(groupRow[5]).toMatch(/leave availed/i);
    expect(groupRow[10]).toMatch(/Remarks/i);
    const numberRow = normalized.rows[5];
    expect(numberRow.slice(0, 11)).toEqual(['1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11']);
    const data = normalized.formFKAModel.partIRows[0];
    expect(data[0]).toBe('01-May-2026');
    expect(data[1]).toBe('31-May-2026');
    expect(data[2]).toBe('31');
    expect(data[3]).toBe('11');
    expect(data[4]).toBe('11');
    expect(data[5]).toBe('29-May-2026');
    expect(data[6]).toBe('10-Jun-2026');
    expect(data[7]).toBe('11');
    expect(data[8]).toBe('0');
    expect(data[9]).toMatch(/29-May-2026 to 10-Jun-2026/);
  });

  test('keeps PART II year / of credit / availed / balance without [object Object]', () => {
    const messy = excelModelAoa();
    messy.push(['', { foo: 'bar' }, '', '', '']);
    const normalized = normalizeFormFKarnatakaPdfMatrix(messy, 11, 0, []);
    expect(normalized.rows.some((row) => /PART II/i.test(String(row[0])))).toBe(true);
    expect(normalized.formFKAModel.partIiTitle).toBe(FORM_F_KA_PDF_PART_II_TITLE);
    expect(normalized.formFKAModel.partIi.year).toBe('2026');
    expect(normalized.formFKAModel.partIi.balance).toBe('0');
    const blob = JSON.stringify(normalized);
    expect(blob).not.toMatch(/\[object Object\]/);
  });

  test('sheetToDenseMatrix uses Form F layout instead of a generic identity table', () => {
    const ws = XLSX.utils.aoa_to_sheet(excelModelAoa());
    const matrix = sheetToDenseMatrix(ws, 'Form_F_Karnataka');
    expect(matrix.formFKALayout).toBe(true);
    expect(matrix.formFKAModel.identity.slNo).toBe('VE0761');
    expect(matrix.formFKAModel.identity.personName).toBe('Vijeesh Vijayan');
    expect(matrix.metaLines[0]).toBe(FORM_F_KA_PDF_TITLE);
    expect(matrix.rows[0][1]).toMatch(/SI No in the Register of Adult/i);
    expect(matrix.rows.some((row) => /PART I EARNED LEAVE/i.test(String(row[0])))).toBe(true);
  });
});
