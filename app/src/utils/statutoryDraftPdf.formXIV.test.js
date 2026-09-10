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
  buildStatutoryPdfHeaderModel,
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
    expect(matrix.formXIVKALayout).toBe(true);

    // Karnataka: header fields sit in the 2-column table (label | value), not as meta bands.
    const contractorRow = matrix.rows.find((r) => /name and address if contractor/i.test(String(r?.[0] || '')));
    const establishmentRow = matrix.rows.find((r) => /establishment/i.test(String(r?.[0] || '')));
    const natureRow = matrix.rows.find((r) => /nature and location of work/i.test(String(r?.[0] || '')));
    const principalRow = matrix.rows.find((r) => /principal employer/i.test(String(r?.[0] || '')));
    expect(contractorRow?.[1]).toMatch(/VAYONA ENERGY PRIVATE LIMITED/i);
    expect(establishmentRow?.[1]).toMatch(/Babaleshwar Hero Site/i);
    expect(natureRow?.[1]).toMatch(/KA-Bableshwar/i);
    expect(principalRow?.[1]).toMatch(/Clean Wind Power/i);

    const metaBlob = (matrix.metaLines || []).join('\n');
    expect(metaBlob).toMatch(/FORM XIV/i);
    expect(metaBlob).toMatch(/see rule 76/i);
    expect(metaBlob).toMatch(/Employment Card/i);

    const nameRow = matrix.rows.find((r) => /name of the workman/i.test(String(r?.[0] || '')));
    expect(String(nameRow[0])).toMatch(/^1\.\s*Name of the Workman/i);
    expect(String(nameRow[1])).toBe('Ameerkhan Naregal');
    expect(String(matrix.rows.find((r) => /serial/i.test(String(r?.[0] || '')))[1])).toBe('VE0712');
    expect(String(matrix.rows.find((r) => /wage rate/i.test(String(r?.[0] || '')))[1])).toBe(
      '71392'
    );
    expect(matrix.formXIVFooterLines?.[0]).toMatch(/signature of (?:the )?contractor/i);
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
    expect(normalized.metaLines[0]).toMatch(/FORM XIV/i);
    expect(normalized.metaLines.find((l) => /contractor/i.test(l))).toMatch(/VAYONA ENERGY/i);
    expect(normalized.metaLines.find((l) => /establishment/i.test(l))).toMatch(/Babaleshwar Hero Site/i);
    expect(normalized.metaLines.find((l) => /nature and location/i.test(l))).toMatch(/Wind Site/i);
    expect(normalized.metaLines.find((l) => /principal employer/i.test(l))).toMatch(/Clean Wind Power/i);
    expect(normalized.rows[0][0]).toMatch(/name of the workman/i);
    expect(normalized.rows[0][1]).toBe('Ameerkhan Naregal');
    expect(normalized.formXIVFooterLines?.[0]).toMatch(/signature of the contractor/i);
  });

  test('normalize prefers column E value over empty spacer cells', () => {
    const rows = [
      ['1', 'Name of the Workman', '', '', 'Jeevan', '', ''],
      ['2', 'Serial number in the register of workmen employed', '', '', 'VE1256', '', ''],
    ];
    const normalized = normalizeFormXIVEmploymentCardPdfMatrix(rows, 7, 0);
    expect(normalized.colCount).toBe(2);
    // Four header fields are meta (outside the box); workman rows start at index 0.
    expect(normalized.metaLines.some((l) => /contractor/i.test(l))).toBe(true);
    expect(normalized.rows[0][0]).toMatch(/name of the workman/i);
    expect(normalized.rows[0][1]).toBe('Jeevan');
    expect(normalized.rows[1][1]).toBe('VE1256');
  });

  test('second normalize keeps header values outside the workman box', () => {
    const compact = [
      ['Name and address if contractor:', ''],
      ['Name and address of Establishment in/under which contract is carried on:', 'Site A'],
      ['Nature and location of work:', ''],
      ['Name and address of Principal Employer:', 'Employer A'],
      ['1. Name of the Workman', 'Jeevan'],
    ];
    const again = normalizeFormXIVEmploymentCardPdfMatrix(compact, 2, 0);
    expect(again.metaLines.find((l) => /establishment/i.test(l))).toMatch(/Site A/i);
    expect(again.metaLines.find((l) => /principal employer/i.test(l))).toMatch(/Employer A/i);
    expect(again.rows[0][0]).toMatch(/name of the workman/i);
    expect(again.rows[0][1]).toBe('Jeevan');
    expect(again.rows.every((r) => !/contractor|establishment/i.test(String(r?.[0] || '')))).toBe(true);
  });

  test('GJ boxed layout pairs values below labels and keeps them outside the table', () => {
    const rows = [
      ['FORM XIV'],
      ['EMPLOYMENT CARD'],
      ['(Vide Rule 76 of Contract Labour (R&A) Central & Gujarat Rules)'],
      ['Name and Address of the Contractor'],
      ['VAYONA ENERGY PRIVATE LIMITED Amreli Renewable Energy Pvt Ltd'],
      ['Name and Address of the Establishment in/under which Contract is carried on:'],
      ['Nanikundal Site, Gujarat'],
      ['Nature of work and location of work:'],
      ['Wind Farm'],
      ['Name and address of Principal Employer:'],
      ['Principal Co'],
      ['1. Name of the Workman', 'Aditya Sharma'],
      ['5. Wage rate (With particulars of unit in case of piece - work)', '59083'],
      ['Signature of the Contractor'],
    ];
    const normalized = normalizeFormXIVEmploymentCardPdfMatrix(rows, 2, 0);
    const meta = (normalized.metaLines || []).join('\n');
    expect(meta).toMatch(/contractor:\s*VAYONA ENERGY/i);
    expect(meta).toMatch(/establishment[\s\S]*Nanikundal Site/i);
    expect(meta).toMatch(/nature and location of work:\s*Wind Farm/i);
    expect(meta).toMatch(/principal employer:\s*Principal Co/i);
    expect(normalized.rows.some((r) => /contractor/i.test(String(r?.[0] || '')))).toBe(false);
    expect(normalized.rows[0][1]).toBe('Aditya Sharma');
    expect(normalized.rows.find((r) => /wage rate/i.test(String(r?.[0] || '')))[1]).toBe('59083');
    expect(normalized.formXIVFooterLines?.[0]).toMatch(/signature of the contractor/i);
  });

  test('GJ 2x2 header boxes map left/right values to Establishment and Principal Employer', () => {
    // Official Gujarat Employment Card: labels on one row, values in side-by-side boxes below.
    const rows = [
      ['FORM XIV'],
      ['EMPLOYMENT CARD'],
      ['(Vide Rule 76 of Contract Labour (R&A) Central & Gujarat Rules)'],
      [
        'Name and Address of the Contractor',
        '',
        '',
        '',
        '',
        '',
        'Name and Address of the Establishment in/under which Contract is carried on:',
      ],
      [
        'VAYONA ENERGY PRIVATE LIMITED',
        '',
        '',
        '',
        '',
        '',
        'Amreli EDF Nanikundal Site 66/33Kv Substation',
      ],
      [
        'Nature of work and location of work:',
        '',
        '',
        '',
        '',
        '',
        'Name and address of Principal Employer:',
      ],
      [
        'GJ-Amreli Junior Engineer',
        '',
        '',
        '',
        '',
        '',
        'Amreli Renewable Energy Pvt Ltd Village:-Nanikundal',
      ],
      ['1. Name of the Workman', '', '', '', '', '', 'Aditya Sharma'],
      ['Signature of the Contractor'],
    ];
    const normalized = normalizeFormXIVEmploymentCardPdfMatrix(rows, 13, 0);
    const contractor = normalized.metaLines.find((l) => /if contractor/i.test(l));
    const establishment = normalized.metaLines.find((l) => /establishment/i.test(l));
    const nature = normalized.metaLines.find((l) => /nature and location/i.test(l));
    const principal = normalized.metaLines.find((l) => /principal employer/i.test(l));
    expect(contractor).toMatch(/VAYONA ENERGY PRIVATE LIMITED/i);
    expect(contractor).not.toMatch(/Amreli EDF|Junior Engineer|Village:-Nanikundal/i);
    expect(establishment).toMatch(/Amreli EDF Nanikundal Site/i);
    expect(nature).toMatch(/GJ-Amreli Junior Engineer/i);
    expect(principal).toMatch(/Amreli Renewable Energy Pvt Ltd/i);
    expect(establishment).toMatch(/:/);
    expect(String(establishment).split(':')[1] || '').toMatch(/\S/);
    expect(String(principal).split(':')[1] || '').toMatch(/\S/);
  });

  test('does not paint designation as a title band under the Form XIV GJ rule line', () => {
    const metaLines = [
      'FORM XIV',
      'EMPLOYMENT CARD',
      '(Vide Rule 76 of Contract Labour (R&A) Central & Gujarat Rules)',
      'Name and address if contractor: VAYONA ENERGY PRIVATE LIMITED',
      'Name and address of Establishment in/under which contract is carried on: Amreli Site',
      'Nature and location of work: GJ-Amreli\nJunior Engineer',
      'Name and address of Principal Employer: Principal Co',
    ];
    const rows = [
      ['1. Name of the Workman', 'Aditya Sharma'],
      ['3. Nature of Employment /Designation', 'Junior Engineer'],
    ];
    const normalized = normalizeFormXIVEmploymentCardPdfMatrix(
      [
        ['FORM XIV'],
        ['EMPLOYMENT CARD'],
        ['(Vide Rule 76 of Contract Labour (R&A) Central & Gujarat Rules)'],
        ['Name and address if contractor:', 'VAYONA ENERGY'],
        ['Name and address of Establishment in/under which contract is carried on:', 'Amreli Site'],
        ['Nature and location of work:', 'GJ-Amreli\nJunior Engineer'],
        ['Name and address of Principal Employer:', 'Principal Co'],
        ['1. Name of the Workman', 'Aditya Sharma'],
        ['3. Nature of Employment /Designation', 'Junior Engineer'],
      ],
      2,
      0
    );
    expect(normalized.metaLines.every((l) => !/^Junior Engineer$/i.test(String(l).trim()))).toBe(
      true
    );
    expect(normalized.metaLines.find((l) => /nature and location/i.test(l))).toMatch(
      /GJ-Amreli\s+Junior Engineer/i
    );
    expect(normalized.metaLines.find((l) => /nature and location/i.test(l))).not.toMatch(/\n/);

    const model = buildStatutoryPdfHeaderModel(metaLines, rows, 0, 'Form_XIV_GJ');
    expect(model.titles.some((t) => /^Junior Engineer$/i.test(String(t).trim()))).toBe(false);
    expect(model.fields.find((f) => /nature and location/i.test(f))).toMatch(/Junior Engineer/i);
  });
});
