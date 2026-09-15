import * as XLSX from 'xlsx';
import {
  looksLikeFormMGJGujaratPdfContext,
  normalizeFormMGJGujaratPdfMatrix,
} from './statutoryDraftPdf.formMGJ.GJ';
import { statutoryDraftPdfTestUtils } from './statutoryDraftPdf';

const { sheetToDenseMatrix, looksLikeFormXIVEmploymentCardPdfContext } = statutoryDraftPdfTestUtils;

describe('Form M GJ Identity Card PDF layout', () => {
  const alignedRows = [
    ['', '', '', '', '', 'FORM - M', '', '', '', 'Photograph'],
    ['', '', '', '', '', '(See Rule 16)'],
    ['', '', '', '', '', 'IDENTITY CARD'],
    [],
    ['', '(a) Name and address of the establishment:', 'Maliya Site'],
    ['', '(b) The full name and address of the worker:', 'Patel Chandrakant Viralbhai'],
    ['', '(c) Date of birth of the worker;', '24 Jun 2004'],
    ['', '(d) Date of joining the service in the establishment:', '01 Dec 2025'],
    ['', '(e) Recent passport size photograph of the worker.'],
    ['', '(f) Contact No.', '9876543210'],
    [],
    ['', '', '', '', '', '', '', 'Signature or left thumb impression of the worker.'],
    ['', '', '', '', '', '', '', 'Signature of Manager or Authorized Agent.'],
    ['', '', '', '', '', '', '', 'Date of Issue.'],
  ];

  const scatteredRows = [
    ['', '', '', '', '', 'FORM - M', '', '', '', 'Photograph'],
    ['', '', '', '', '', '(See Rule 16)'],
    ['', '', '', '', '', 'IDENTITY CARD'],
    ['', 'Maliya Site'],
    ['', 'Harshad Dk'],
    ['', '(c) Date of birth of the worker;'],
    ['', '22 Jun 2000'],
    ['', '01 Dec 2025'],
    ['', '(e) Recent passport size photograph of the worker.'],
    ['', '(f) Contact No.'],
    ['', '', '', '', '', '', '', 'Signature or left thumb impression of the worker:'],
    [
      '',
      '(a) Name and address of the establishment:',
      'No.114/1,Khirai Patiya behind Kerosene Depot,Takuka-Maliya,Dist-Morbi GJ - 3636370, Maliya, Gujarat',
    ],
    ['', '(b) The full name and address of the worker: -'],
    ['', '(d) Date of joining the service in the establishment:'],
  ];

  test('detects Gujarat Form M identity card and not Form XIV employment card', () => {
    expect(
      looksLikeFormMGJGujaratPdfContext(
        ['FORM - M', '(See Rule 16)', 'IDENTITY CARD'],
        alignedRows,
        'Form_M_GJ_Gujarat_Employees.xlsx'
      )
    ).toBe(true);
    expect(
      looksLikeFormXIVEmploymentCardPdfContext(
        ['FORM - M', '(See Rule 16)', 'IDENTITY CARD'],
        alignedRows,
        'Form_M_GJ_Gujarat.xlsx'
      )
    ).toBe(false);
    expect(
      looksLikeFormMGJGujaratPdfContext(
        ['FORM XIV', '[See rule 76]', 'Employment Card'],
        [['Name of the Workman', 'A']],
        'Form_XIV_KA.xlsx'
      )
    ).toBe(false);
  });

  test('pairs labels with same-row values and keeps photograph out of the field list', () => {
    const normalized = normalizeFormMGJGujaratPdfMatrix(alignedRows, 10, 0, []);
    expect(normalized.formMGJLayout).toBe(true);
    expect(normalized.colCount).toBe(2);
    expect(normalized.formMGJTitles[0]).toMatch(/FORM\s*-?\s*M/i);
    expect(normalized.formMGJTitles).toEqual(
      expect.arrayContaining(['(See Rule 16)', 'IDENTITY CARD'])
    );
    const byKey = Object.fromEntries(normalized.formMGJFields.map((f) => [f.key, f]));
    expect(byKey.establishment.value).toBe('Maliya Site');
    expect(byKey.workerNameAddress.value).toBe('Patel Chandrakant Viralbhai');
    expect(byKey.dateOfBirth.value).toBe('24 Jun 2004');
    expect(byKey.dateOfJoining.value).toBe('01 Dec 2025');
    expect(byKey.contactNo.value).toBe('9876543210');
    expect(byKey.photograph.value).toBe('');
    expect(normalized.formMGJFields.map((f) => f.label).join(' ')).not.toMatch(/^Photograph$/m);
    expect(normalized.formMGJSignatures[0]).toMatch(/thumb impression/i);
    expect(normalized.formMGJSignatures[1]).toMatch(/Manager or Authorized Agent/i);
    expect(normalized.formMGJSignatures[2]).toMatch(/Date of Issue/i);
  });

  test('rebuilds scattered Excel/PDF rows into official a–f order', () => {
    const normalized = normalizeFormMGJGujaratPdfMatrix(scatteredRows, 10, 0, []);
    const byKey = Object.fromEntries(normalized.formMGJFields.map((f) => [f.key, f]));
    expect(normalized.formMGJFields.map((f) => f.key)).toEqual([
      'establishment',
      'workerNameAddress',
      'dateOfBirth',
      'dateOfJoining',
      'photograph',
      'contactNo',
    ]);
    expect(byKey.establishment.value).toMatch(/Maliya Site/i);
    expect(byKey.establishment.value).toMatch(/Khirai Patiya/i);
    expect(byKey.workerNameAddress.value).toBe('Harshad Dk');
    expect(byKey.dateOfBirth.value).toBe('22 Jun 2000');
    expect(byKey.dateOfJoining.value).toBe('01 Dec 2025');
    expect(byKey.photograph.value).toBe('');
  });

  test('splits concatenated label:value cells', () => {
    const rows = [
      ['FORM - M'],
      ['(See Rule 16)'],
      ['IDENTITY CARD'],
      ['(a) Name and address of the establishment: Maliya Site'],
      ['(b) The full name and address of the worker: Patel Chandrakant Viralbhai'],
      ['(c) Date of birth of the worker; 24 Jun 2004'],
      ['(d) Date of joining the service in the establishment: 01 Dec 2025'],
      ['(e) Recent passport size photograph of the worker.'],
      ['(f) Contact No. 9988776655'],
    ];
    const normalized = normalizeFormMGJGujaratPdfMatrix(rows, 1, 0, []);
    const byKey = Object.fromEntries(normalized.formMGJFields.map((f) => [f.key, f]));
    expect(byKey.establishment.value).toBe('Maliya Site');
    expect(byKey.workerNameAddress.value).toBe('Patel Chandrakant Viralbhai');
    expect(byKey.dateOfBirth.value).toBe('24 Jun 2004');
    expect(byKey.dateOfJoining.value).toBe('01 Dec 2025');
    expect(byKey.contactNo.value).toBe('9988776655');
  });

  test('sheetToDenseMatrix marks Form M GJ as identity-card layout', () => {
    const aoa = [
      ['FORM - M', '', '', '', '', '', '', '', '', 'Photograph'],
      ['(See Rule 16)'],
      ['IDENTITY CARD'],
      ['(a) Name and address of the establishment:', 'Maliya Site'],
      ['(b) The full name and address of the worker:', 'Patel Chandrakant Viralbhai'],
      ['(c) Date of birth of the worker;', '24 Jun 2004'],
      ['(d) Date of joining the service in the establishment:', '01 Dec 2025'],
      ['(e) Recent passport size photograph of the worker.'],
      ['(f) Contact No.', '9000000000'],
    ];
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    const matrix = sheetToDenseMatrix(ws, 'Form_M_GJ_Gujarat.xlsx');
    expect(matrix.formMGJLayout).toBe(true);
    expect(matrix.colCount).toBe(2);
    expect(matrix.formMGJFields[0].value).toBe('Maliya Site');
    expect(matrix.formMGJFields[1].value).toBe('Patel Chandrakant Viralbhai');
    expect(matrix.rows.every((row) => row.length === 2)).toBe(true);
    expect(matrix.rows.some((row) => /^Photograph$/i.test(String(row[0] || '')))).toBe(false);
  });
});
