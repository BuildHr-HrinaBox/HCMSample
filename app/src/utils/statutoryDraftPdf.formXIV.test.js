import * as XLSX from 'xlsx';
import { statutoryDraftPdfTestUtils } from './statutoryDraftPdf';

const {
  looksLikeFormXIVEmploymentCardPdfContext,
  looksLikeFormXRajasthanEmploymentCardPdfContext,
  buildFormXRajasthanEmploymentHeaderModel,
  trimFormXRajasthanEmploymentLeadingBlankPdfColumns,
  looksLikeFormXIRajasthanServiceCertificatePdfContext,
  buildFormXIRajasthanServiceHeaderModel,
  looksLikeFormXVRajasthanWageSlipPdfContext,
  buildFormXVRajasthanWageSlipHeaderModel,
  normalizeFormXIVEmploymentCardPdfMatrix,
  sheetToDenseMatrix,
} = statutoryDraftPdfTestUtils;

describe('Form XIV Employment Card PDF layout', () => {
  test('pairs Rajasthan Form X Employment Card headings with following values', () => {
    const metaLines = [
      '[See Rule 75]',
      'Employment Card',
      'Sample',
      'RJ-Fatehgarh-2',
      'Name and address of contractor........................................',
      'Sample',
      'Nature and location of work............................................',
      'RJ-Fatehgarh-2',
      'Name and address of establishment under which contract is carried on........................',
      'Rajasthan, RSEPL HYBRID POWER ONER ONE LIMITED, AEML-1, 250 MW, Sandhua, Fatehgarh, Jaisalmer,245027, Chennai, Rajasthan',
      'Name and address of principal employer................................',
      'VAYONA ENERGY PRIVATE LIMITED, Vayona Energy Pvt Ltd, 274 A, Rani mangammal main road, Govindanagaram - 625517 Theni Taluk, Theni District., Theni, TamilNadu',
    ];
    expect(
      looksLikeFormXRajasthanEmploymentCardPdfContext(metaLines, [], 'Form_X_RJ_Rajasthan.xlsx')
    ).toBe(true);

    const model = buildFormXRajasthanEmploymentHeaderModel(metaLines);
    expect(model.fields[0]).toMatch(/contractor: Sample$/i);
    expect(model.fields[1]).toMatch(/work: RJ-Fatehgarh-2$/i);
    expect(model.fields[2]).toMatch(/establishment under which contract is carried on: Rajasthan/i);
    expect(model.fields[3]).toMatch(/principal employer: VAYONA ENERGY/i);
    expect(model.titles).not.toContain('Sample');
    expect(model.titles).not.toContain('RJ-Fatehgarh-2');
  });

  test('removes the empty leading table column from Rajasthan Form X', () => {
    const rows = [
      ['', 'Name of the workman', 'Sl. No. of the register of workman employed'],
      ['', '', '1'],
      ['', 'Darshan Darji', 'VE0576'],
    ];
    const out = trimFormXRajasthanEmploymentLeadingBlankPdfColumns(rows, 3, 0);
    expect(out.colCount).toBe(2);
    expect(out.rows[0]).toEqual([
      'Name of the workman',
      'Sl. No. of the register of workman employed'
    ]);
    expect(out.rows[2]).toEqual(['Darshan Darji', 'VE0576']);
  });

  test('pairs Rajasthan Form XI Service Certificate headings with values', () => {
    const metaLines = [
      'FORM XI', '[See Rule 76]', 'Service Certificate',
      'Name and address of contractor', 'Sample',
      'Nature and location of work', 'RJ-Fatehgarh-2',
      'Name and address of establishment under which contract is carried on', 'Rajasthan establishment address',
      'Name and address of principal employer', 'Vayona Energy address',
      'Name and address of the workman', 'Hariom Dholi',
      'Age or date of birth', '05-Jan-1994',
      'Identification marks', 'NIL',
      "Father's / Husband's Name", 'Bh enru Lal'
    ];
    expect(looksLikeFormXIRajasthanServiceCertificatePdfContext(metaLines, [], 'Form_XI_RJ.xlsx')).toBe(true);
    const model = buildFormXIRajasthanServiceHeaderModel(metaLines);
    expect(model.fields[0]).toBe('Name and address of contractor: Sample');
    expect(model.fields[1]).toBe('Nature and location of work: RJ-Fatehgarh-2');
    expect(model.fields[3]).toBe('Name and address of principal employer: Vayona Energy address');
    expect(model.fields[7]).toBe("Father's / Husband's Name: Bh enru Lal");
    expect(model.titles).not.toContain('Sample');
  });

  test('pairs Rajasthan Form XV Wage Slip headings with values', () => {
    const metaLines = [
      'FORM XV', '[See Rule 77(2)(b)]', 'Wages Slip', 'Sample', 'RJ-Fatehgarh-2',
      'Name and address of contractor', 'Sample',
      'Name and location of work', 'RJ-Fatehgarh-2',
      'Name and address of establishment in/under which contract is carried on', 'Rajasthan establishment address',
      'Name and address of principal employer', 'Vayona Energy address',
      "Name and Father's name of the workman", 'Janga Babu Kotaiah',
      'Sex and identification token/ticket No.', 'Male VE1476',
      'For the week/fortnight/month', 'July 2026',
    ];
    expect(looksLikeFormXVRajasthanWageSlipPdfContext(metaLines, [], 'Form_XV_RJ.xlsx')).toBe(true);
    const model = buildFormXVRajasthanWageSlipHeaderModel(metaLines);
    expect(model.fields[0]).toBe('Name and address of contractor: Sample');
    expect(model.fields[1]).toBe('Name and location of work: RJ-Fatehgarh-2');
    expect(model.fields[4]).toMatch(/workman: Janga Babu Kotaiah$/i);
    expect(model.fields[6]).toBe('For the week/fortnight/month: July 2026');
    expect(model.titles).not.toContain('Sample');
  });

  test('detects Form XIV Employment Card context', () => {
    expect(
      looksLikeFormXIVEmploymentCardPdfContext(
        ['FORM XIV', '[See rule 76]', 'Employment Card'],
        [['1', 'Name of the Workman', '', '', 'Ameerkhan']],
        'Form_XIV_KA'
      )
    ).toBe(true);
  });

  test('always shows the four header fields as label|value rows', () => {
    // Mimic KA/MP stacked template: ordinals in A, labels in B, values in E, empty F+.
    const aoa = [
      ['FORM XIV', '', '', '', '', '', '', ''],
      ['[See rule 76]', '', '', '', '', '', '', ''],
      ['Employment Card', '', '', '', '', '', '', ''],
      ['VAYONA ENERGY PRIVATE LIMITED, vvd/14', '', '', '', '', '', '', ''],
      ['KA-Bableshwar', '', '', '', '', '', '', ''],
      ['Name and address if contractor', '', '', '', '', '', '', ''],
      [
        'Name and address of Establishment in/under which contract is carried on',
        '',
        '',
        '',
        'Babaleshwar Hero Site, Karnataka',
        '',
        '',
        '',
      ],
      ['Nature and location of work', '', '', '', '', '', '', ''],
      [
        'Name and address of Principal Employer',
        '',
        '',
        '',
        'M/s Clean Wind Power Bableshwar Pvt Ltd',
        '',
        '',
        '',
      ],
      ['1', 'Name of the Workman', '', '', 'Ameerkhan Naregal', '', '', ''],
      [
        '2',
        'Serial No. in the Register of workmen employed',
        '',
        '',
        'VE0712',
        '',
        '',
        '',
      ],
      ['3', 'Nature of employment/Designation', '', '', 'Engineer', '', '', ''],
      [
        '4',
        'Wage rate (with particulars of unit in case of piece-work)',
        '',
        '',
        '71392',
        '',
        '',
        '',
      ],
      ['5', 'Wage period', '', '', 'May 2026', '', '', ''],
      ['6', 'Tenure of employment', '', '', 'From 01 Dec 2025', '', '', ''],
      ['7', 'Remarks', '', '', '', '', '', ''],
      ['Signature of Contractor', '', '', '', '', '', '', ''],
    ];
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    const matrix = sheetToDenseMatrix(ws, 'Form_XIV_Karnataka');

    expect(
      looksLikeFormXIVEmploymentCardPdfContext(matrix.metaLines, matrix.rows, matrix.name)
    ).toBe(true);
    expect(matrix.colCount).toBe(2);

    const contractorRow = matrix.rows.find((r) => /contractor/i.test(String(r?.[0] || '')));
    expect(contractorRow).toBeTruthy();
    expect(String(contractorRow[0])).toMatch(/name and address if contractor/i);

    const establishmentRow = matrix.rows.find((r) => /establishment/i.test(String(r?.[0] || '')));
    expect(establishmentRow).toBeTruthy();
    expect(String(establishmentRow[1])).toMatch(/Babaleshwar Hero Site/i);

    const natureRow = matrix.rows.find((r) => /nature and location of work/i.test(String(r?.[0] || '')));
    expect(natureRow).toBeTruthy();

    const principalRow = matrix.rows.find((r) => /principal employer/i.test(String(r?.[0] || '')));
    expect(principalRow).toBeTruthy();
    expect(String(principalRow[1])).toMatch(/Clean Wind Power/i);

    // Workman rows stay in correct columns.
    const nameRow = matrix.rows.find((r) => /name of the workman/i.test(String(r?.[0] || '')));
    expect(String(nameRow[0])).toMatch(/^1\.\s*Name of the Workman/i);
    expect(String(nameRow[1])).toBe('Ameerkhan Naregal');
    expect(String(matrix.rows.find((r) => /serial/i.test(String(r?.[0] || '')))[1])).toBe('VE0712');
    expect(String(matrix.rows.find((r) => /wage rate/i.test(String(r?.[0] || '')))[1])).toBe(
      '71392'
    );
  });

  test('matches header labels that include "the" and sit in column B', () => {
    const rows = [
      ['FORM XIV'],
      ['Employment Card'],
      ['', 'Name and Address of the Contractor', '', '', 'VAYONA ENERGY'],
      [
        '',
        'Name and Address of the Establishment in/under which Contract is carried on:',
        '',
        '',
        'Babaleshwar Hero Site',
      ],
      ['', 'Nature and location of work:', '', '', 'Wind Site'],
      ['', 'Name and address of the Principal Employer', '', '', 'Clean Wind Power'],
      ['1', 'Name of the Workman', '', '', 'Ameerkhan Naregal'],
    ];
    const normalized = normalizeFormXIVEmploymentCardPdfMatrix(rows, 5, 0);
    expect(normalized.colCount).toBe(2);
    expect(normalized.rows[0][0]).toMatch(/contractor/i);
    expect(normalized.rows[0][1]).toBe('VAYONA ENERGY');
    expect(normalized.rows[1][1]).toBe('Babaleshwar Hero Site');
    expect(normalized.rows[2][1]).toBe('Wind Site');
    expect(normalized.rows[3][1]).toBe('Clean Wind Power');
    expect(normalized.rows[4][1]).toBe('Ameerkhan Naregal');
  });

  test('normalize prefers column E value over empty spacer cells', () => {
    const rows = [
      ['1', 'Name of the Workman', '', '', 'Jeevan', '', ''],
      ['2', 'Serial number in the register of workmen employed', '', '', 'VE1256', '', ''],
    ];
    const normalized = normalizeFormXIVEmploymentCardPdfMatrix(rows, 7, 0);
    expect(normalized.colCount).toBe(2);
    // Four header rows are always prepended (blank values when absent).
    expect(normalized.rows[0][0]).toMatch(/contractor/i);
    expect(normalized.rows[4][0]).toMatch(/name of the workman/i);
    expect(normalized.rows[4][1]).toBe('Jeevan');
    expect(normalized.rows[5][1]).toBe('VE1256');
  });

  test('second normalize does not wipe already-compacted header rows', () => {
    const compact = [
      ['Name and address if contractor:', ''],
      ['Name and address of Establishment in/under which contract is carried on:', 'Site A'],
      ['Nature and location of work:', ''],
      ['Name and address of Principal Employer:', 'Employer A'],
      ['1. Name of the Workman', 'Jeevan'],
    ];
    const again = normalizeFormXIVEmploymentCardPdfMatrix(compact, 2, 0);
    expect(again.metaLines).toBeNull();
    expect(again.rows[1][1]).toBe('Site A');
    expect(again.rows[3][1]).toBe('Employer A');
    expect(again.rows[4][1]).toBe('Jeevan');
  });
});
