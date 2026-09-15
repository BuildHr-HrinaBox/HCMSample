import * as XLSX from 'xlsx';
import { statutoryDraftPdfTestUtils } from './statutoryDraftPdf';
import {
  detectFormTSEKarnatakaAttendanceBand,
  extractFormTSEKarnatakaHeaderFields,
  FORM_T_KA_PDF_ATTENDANCE_LABEL,
  FORM_T_KA_PDF_DATE_LABEL,
  FORM_T_KA_PDF_DEDUCTIONS_GROUP,
  FORM_T_KA_PDF_EARNED_WAGES_GROUP,
  FORM_T_KA_PDF_IDENTITY_HEADERS,
  FORM_T_KA_PDF_SIGNATORY_LABEL,
  FORM_T_KA_PDF_SUBTITLE,
  FORM_T_KA_PDF_TITLE,
  isFormTSEKarnatakaPdfFooterOnlyRow,
  looksLikeFormTSEKarnatakaPdfContext,
  normalizeFormTSEKarnatakaPdfMatrix,
  shiftFormTSEKarnatakaPdfIdentityFromJToA,
} from './statutoryDraftPdf.formT.KA';
import { looksLikeFormXIIIAPPdfContext } from './statutoryDraftPdf.formXIII.AP';

const { buildStatutoryPdfHeaderModel, sheetToDenseMatrix } = statutoryDraftPdfTestUtils;

const formTAoa = () => {
  const identity = [
    'S.NO',
    'Name of Employee',
    "Father / Husband's Name",
    'Gender',
    'Designation / Department',
    'Date of Joining',
    'ESI No.',
    'UAN No.',
    'Wages fixed including VDA',
  ];
  const attendanceBanner = Array.from({ length: 31 }, () => 'ATTENDANCE');
  const dayNumbers = Array.from({ length: 31 }, (_, i) => String(i + 1));
  const wageLeaves = [
    'No. of payable days',
    'Total OT hours',
    'BASIC',
    'DA/VDA',
    'HRA',
    'Conveyance',
    'Medical Allowance',
    'Attendance Bonus',
    'Special Allowance',
    'OT',
    'NFH',
    'Maternity Benefit',
    'Others',
    'Subsistence Allowance if any',
    'Total',
    'ESI',
    'PF',
    'PT',
    'Society',
    'Insurance',
    'Salary Advance',
    'Fines',
    'Damages',
    'Others',
    'Total',
    'Net Amount Payable',
    'Mode of Payment Cash/ Cheque No.',
    'Employee signature or thumb impression',
  ];
  const header = [...identity, ...attendanceBanner, ...wageLeaves];
  const days = [...Array(9).fill(''), ...dayNumbers, ...Array(wageLeaves.length).fill('')];
  const employee = [
    '1',
    'Vinay Kumar',
    'Ramesh',
    'Male',
    'Engineer',
    '12 Aug 2025',
    '101455041067',
    '',
    '25000',
    ...Array.from({ length: 31 }, () => 'P'),
    '31',
    'NIL',
    '18000',
    '2000',
    '1500',
    '800',
    '500',
    '0',
    '0',
    '0',
    '0',
    '0',
    '0',
    '0',
    '22800',
    '111',
    '1800',
    '200',
    '0',
    '0',
    '0',
    '0',
    '0',
    '0',
    '2111',
    '20689',
    'Bank Transfer',
    '',
  ];
  return [
    ['FORM T'],
    ['COMBINED MUSTER ROLL CUM REGISTER OF WAGES'],
    ['[See Rule 24(9-B) of Karnataka Shops & Commercial Establishment Rules, 1963]'],
    ['in lieu of'],
    [
      '1. Form I, II of Rule 22(4); Form IV of Rule 29(2); Forms V & VII of Rule 29(1) & (5) of Karnataka Minimum Wages Rules, 1958',
    ],
    [
      '2. Form I of Rules 3(1) of Karnataka Payment of Wages Rules, 1963',
    ],
    [
      '3. Form XIII of Rules 75; Form XV, XVII, XX, XXI, XXII, XXIII of 78(1)(a)(i), (ii) & (iii) of Karnataka Contract Labour (Regulation & Abolition) Rules, 1974',
    ],
    [
      '4. Form XIII of Rule 43; Forms XVII, XVIII, XIX, XX, XXI, XXII of Rule 46(2)(a),(c) & (d) of Inter-state Migrant Workmen (Regulation of Employment and conditions of service) Karnataka Rules, 1981',
    ],
    ['Month / Year : September 2025'],
    [
      'Name and address of the Establishment : Bableshwar Hero Site, 33/220KV Substation, Survey No: 132, Karjol Village, Bijapur Dist - 586113 Karnataka, Bableshwar, Karnataka',
    ],
    [
      'Name and Address of employer : M/s. Clean Wind Power Bableshwar Pvt Ltd, 33/220KV Substation, Survey No: 132, Karjol Village, Bijapur Dist - 586113 Karnataka, Bableshwar, Karnataka',
    ],
    header,
    days,
    employee,
  ];
};

describe('Form T Karnataka PDF layout', () => {
  test('detects Form T KA even when citation lines mention Form XIII', () => {
    const rows = formTAoa();
    expect(
      looksLikeFormTSEKarnatakaPdfContext(
        ['FORM T', 'COMBINED MUSTER ROLL CUM REGISTER OF WAGES'],
        rows,
        'Form T',
        'Form_T_KA_Karnataka.xlsx'
      )
    ).toBe(true);
    expect(looksLikeFormXIIIAPPdfContext(rows.slice(0, 8).map((r) => r[0]), rows, 'Form T')).toBe(
      false
    );
  });

  test('does not treat a real AP Form XIII sheet as Form T', () => {
    const meta = ['FORM XIII', 'REGISTER OF WORKMEN EMPLOYED BY CONTRACTOR', '[Vide Rule 75]'];
    const rows = [['S.No', 'Name and Surname of Workmen']];
    expect(
      looksLikeFormTSEKarnatakaPdfContext(
        meta,
        rows,
        'Form XIII-Register of Workmen',
        'Form_XIII_AP.xlsx'
      )
    ).toBe(false);
    expect(
      looksLikeFormXIIIAPPdfContext(meta, rows, 'Form XIII-Register of Workmen', 'Form_XIII_AP.xlsx')
    ).toBe(true);
  });

  test('builds Form T header like the official combined-muster model, not Form XIII', () => {
    const rows = formTAoa();
    const model = buildStatutoryPdfHeaderModel(
      [],
      rows,
      11,
      'Form T',
      { fileName: 'Form_T_KA_Karnataka.xlsx', preferredTitle: 'Form T Karnataka' }
    );
    expect(model.formTKA).toBe(true);
    expect(model.formXIIIAP).toBeFalsy();
    expect(model.titles[0]).toBe(FORM_T_KA_PDF_TITLE);
    expect(model.titles[1]).toBe(FORM_T_KA_PDF_SUBTITLE);
    expect(model.titles.join('\n')).toMatch(/See Rule 24\(9-B\)/i);
    expect(model.titles.join('\n')).toMatch(/in lieu of/i);
    expect(model.titles.join('\n')).not.toMatch(/REGISTER OF WORKMEN EMPLOYED BY CONTRACTOR/i);
    expect(model.fields.join('\n')).toMatch(/Month \/ Year : September 2025/i);
    expect(model.fields.join('\n')).toMatch(/Name and address of the Establishment :/i);
    expect(model.fields.join('\n')).toMatch(/Bableshwar Hero Site/i);
    expect(model.fields.join('\n')).toMatch(/Clean Wind Power Bableshwar/i);
    expect(model.fields.join('\n')).not.toMatch(/Nature and location of work/i);
  });

  test('keeps employee name and attendance marks in the PDF table', () => {
    const ws = XLSX.utils.aoa_to_sheet(formTAoa());
    const matrix = sheetToDenseMatrix(ws, 'Form_T_KA');
    expect(looksLikeFormTSEKarnatakaPdfContext(matrix.metaLines, matrix.rows, matrix.name)).toBe(
      true
    );
    const normalized = normalizeFormTSEKarnatakaPdfMatrix(
      matrix.rows,
      matrix.colCount,
      matrix.tableStartRow,
      matrix.metaLines
    );
    const flat = normalized.rows.flat().join(' ');
    expect(flat).toMatch(/Vinay Kumar/);
    expect(flat).toMatch(/\bP\b/);
    expect(flat).toMatch(/Bank Transfer/);
    expect(normalized.formTKALayout).toBe(true);
    expect(normalized.metaLines.join('\n')).toMatch(/FORM T/);
    expect(normalized.metaLines.join('\n')).not.toMatch(
      /REGISTER OF WORKMEN EMPLOYED BY CONTRACTOR/i
    );
  });

  test('extracts establishment and employer values from Label : value cells', () => {
    const fields = extractFormTSEKarnatakaHeaderFields(
      [
        'Month / Year : September 2025',
        'Name and address of the Establishment : Bableshwar Hero Site',
        'Name and Address of employer : Clean Wind Power',
      ],
      [],
      0
    );
    expect(fields[0]).toBe('Month / Year : September 2025');
    expect(fields[1]).toMatch(/Bableshwar Hero Site/);
    expect(fields[2]).toMatch(/Clean Wind Power/);
  });

  test('shifts J-dumped employee identity onto column A so S.NO starts first', () => {
    const header = [
      'S.NO',
      'Name and address of principal employer',
      "Father / Husband's Name",
      'Gender',
      'Designation / Department',
      'Date of Joining',
      'ESI No.',
      'UAN No.',
      'Wages fixed including VDA',
      'ATTENDANCE',
      '1',
      '2',
      '3',
    ];
    const index = ['1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '', '', ''];
    const shifted = [
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '1',
      'Prakash',
      'Lakshman',
      'Male',
      'Engineer',
      '15 Jul 2026',
      '123456',
      '',
      '25000',
      'P',
      'P',
      'WO',
    ];
    const dateRow = ['Date:', '', '', '', '', '', '', '', '', '6', 'Elbaz', 'abc', 'Male'];
    const normalized = normalizeFormTSEKarnatakaPdfMatrix(
      [header, index, shifted, dateRow],
      21,
      0,
      ['FORM T', 'COMBINED MUSTER ROLL CUM REGISTER OF WAGES']
    );
    const dataRow = normalized.rows.find((row) => String(row?.[1] || '').includes('Prakash'));
    expect(dataRow).toBeTruthy();
    expect(String(dataRow[0])).toBe('1');
    expect(String(dataRow[1])).toBe('Prakash');
    expect(String(dataRow[3])).toBe('Male');
    expect(String(dataRow[4])).toBe('Engineer');
    expect(String(dataRow[9])).toBe('P');
    expect(normalized.rows.some((row) => String(row?.[1] || '').includes('Elbaz'))).toBe(true);
    const elbaz = normalized.rows.find((row) => String(row?.[1] || '').includes('Elbaz'));
    expect(String(elbaz[0])).toBe('6');
    expect(shiftFormTSEKarnatakaPdfIdentityFromJToA([header])[0][0]).toBe('S.NO');
  });

  test('rebuilds official Excel-model headers and merges ATTENDANCE across days', () => {
    const rows = formTAoa();
    const normalized = normalizeFormTSEKarnatakaPdfMatrix(rows, rows[11].length, 11, []);
    expect(normalized.rows[0][1]).toBe(FORM_T_KA_PDF_IDENTITY_HEADERS[1]);
    expect(normalized.rows[0][1]).not.toMatch(/Name of Employee/i);
    expect(normalized.rows[0][9]).toBe(FORM_T_KA_PDF_ATTENDANCE_LABEL);
    expect(String(normalized.rows[0][10] || '')).toBe('');
    expect(String(normalized.rows[0][39] || '')).toBe('');
    expect(String(normalized.rows[1][9])).toBe('1');
    expect(String(normalized.rows[1][39])).toBe('31');
    expect(String(normalized.rows[2][0])).toBe('1');
    expect(String(normalized.rows[2][8])).toBe('9');
    expect(String(normalized.rows[2][9])).toBe('10');
    expect(normalized.attendanceBand.start).toBe(9);
    expect(normalized.attendanceBand.end).toBe(39);
    expect(normalized.attendanceBand.labelRow).toBe(0);
    expect(normalized.attendanceBand.label).toBe(FORM_T_KA_PDF_ATTENDANCE_LABEL);
    const band = detectFormTSEKarnatakaAttendanceBand(
      normalized.rows,
      0,
      3,
      normalized.colCount
    );
    expect(band.end - band.start + 1).toBe(31);
  });

  test('puts earned-wage leaves under the group and shows the Deductions band', () => {
    const rows = formTAoa();
    const normalized = normalizeFormTSEKarnatakaPdfMatrix(rows, rows[11].length, 11, []);
    const wageStart = 40;
    expect(normalized.rows[0][wageStart]).toMatch(/No\. of payable days/i);
    expect(normalized.rows[0][wageStart + 1]).toMatch(/Total OT hours/i);
    expect(normalized.rows[0][wageStart + 2]).toBe(FORM_T_KA_PDF_EARNED_WAGES_GROUP);
    expect(normalized.rows[0][wageStart + 15]).toBe(FORM_T_KA_PDF_DEDUCTIONS_GROUP);
    expect(String(normalized.rows[1][wageStart + 2])).toMatch(/BASIC/i);
    expect(String(normalized.rows[1][wageStart + 4])).toMatch(/HRA/i);
    expect(String(normalized.rows[1][wageStart + 15])).toMatch(/ESI/i);
    expect(String(normalized.rows[1][wageStart + 16])).toMatch(/PF/i);
    expect(String(normalized.rows[1][wageStart + 17])).toMatch(/PT/i);
    expect(String(normalized.rows[2][wageStart])).toBe('11');
    expect(String(normalized.rows[2][wageStart + 15])).toBe('26');
    const employee = normalized.rows.find((row) => String(row?.[1] || '').includes('Vinay Kumar'));
    expect(String(employee[wageStart + 15])).toBe('111');
  });

  test('reads Establishment and Employer from the following line when the label has no value', () => {
    const fields = extractFormTSEKarnatakaHeaderFields(
      [
        'Name and address of the Establishment :',
        'Bableshwar Hero Site, Karjol Village',
        'Name and Address of employer :',
        'M/s Clean Wind Power Bableshwar Pvt Ltd',
      ],
      [],
      0
    );
    expect(fields[1]).toMatch(/Bableshwar Hero Site/);
    expect(fields[2]).toMatch(/Clean Wind Power Bableshwar/);
  });

  test('puts Date: and Authorised Signatory below the table, not inside the grid', () => {
    const rows = formTAoa();
    rows.push(['Date:', '', '', '', '', 'Authorised Signatory']);
    const normalized = normalizeFormTSEKarnatakaPdfMatrix(rows, rows[11].length, 11, []);
    expect(normalized.formTKAFooter.date).toBe(FORM_T_KA_PDF_DATE_LABEL);
    expect(normalized.formTKAFooter.signatory).toBe(FORM_T_KA_PDF_SIGNATORY_LABEL);
    expect(
      normalized.rows.some((row) => isFormTSEKarnatakaPdfFooterOnlyRow(row))
    ).toBe(false);
    expect(normalized.rows.some((row) => String(row?.[0] || '') === 'Date:')).toBe(false);
    expect(
      normalized.rows.some((row) =>
        (row || []).some((cell) => /authorised\s+signatory/i.test(String(cell || '')))
      )
    ).toBe(false);

    const withValue = normalizeFormTSEKarnatakaPdfMatrix(
      [...formTAoa(), ['Date: 11-09-2026']],
      20,
      11,
      []
    );
    expect(withValue.formTKAFooter.date).toBe('Date: 11-09-2026');
    expect(withValue.formTKAFooter.signatory).toBe(FORM_T_KA_PDF_SIGNATORY_LABEL);
  });
});
