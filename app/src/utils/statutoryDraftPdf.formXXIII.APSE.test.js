import { statutoryDraftPdfTestUtils } from './statutoryDraftPdf';
import {
  extractFormXXIIIAPSEAdminLayout,
  findFormXXIIIAPSEDeductionsBand,
  formXXIIIAPSEColumnWeight,
  formatFormXXIIIAPSEMoneyPdfValue,
  getFormXXIIIAPSEHeaderTitles,
  isFormXXIIIAPSECenterValueHeader,
  isFormXXIIIAPSEMoneyHeader,
  looksLikeFormXXIIIAPSEPdfContext,
  trimFormXXIIIAPSEEmptyPdfColumns
} from './statutoryDraftPdf.formXXIII.AP.S&D.js';

describe('Form XXIII AP Shops & Establishment PDF layout helpers', () => {
  const meta = [
    'Form XXIII – Register of Wages',
    '(Vide Rule 29(2) of A.P. Shops & Establishment Rules, 1990)',
    'Name of Establishment / Shop:',
    'Nimbagallu Site, Vayona Energy Pvt Ltd, 274 A, Rani mangammal main road, Govindanagaram - 625517 Theni Taluk, Theni District., Nimbagallu, Andhra Pradesh',
    'Registration No.:',
    'Address of the Establishment:',
    'Vayona Energy Pvt Ltd, 274 A, Rani mangammal main road, Govindanagaram - 625517 Theni Taluk, Theni District., Nimbagallu, Andhra Pradesh',
    'Wage Period',
    'From:',
    'August',
    'To:',
    'August'
  ];

  test('detects AP S&E Form XXIII Register of Wages and not overtime Form XXIII', () => {
    expect(
      looksLikeFormXXIIIAPSEPdfContext(meta, [], 'Form XXIII – Register of Wages')
    ).toBe(true);
    expect(
      looksLikeFormXXIIIAPSEPdfContext(
        ['FORM XXIII', 'Register of Overtime', '[See Rule 78(1)(a)(iii)]'],
        [],
        'Form XXIII Tamil Nadu'
      )
    ).toBe(false);
  });

  test('keeps form title plus centered vide-rule line without duplicating Register of Wages', () => {
    const titles = getFormXXIIIAPSEHeaderTitles(meta, [], 8);
    expect(titles.filter((line) => /form\s*xxiii/i.test(line))).toHaveLength(1);
    expect(titles.some((line) => /vide\s+rule\s*29/i.test(line))).toBe(true);
    expect(titles.filter((line) => /^register of wages$/i.test(line))).toHaveLength(0);
  });

  test('extracts establishment, address, registration and wage-period cells without changing values', () => {
    const layout = extractFormXXIIIAPSEAdminLayout(meta, [], 8);
    expect(layout.establishmentValue).toContain('Nimbagallu Site, Vayona Energy Pvt Ltd');
    expect(layout.addressValue).toContain('Vayona Energy Pvt Ltd, 274 A, Rani mangammal main road');
    expect(layout.registrationLabel).toMatch(/registration no/i);
    expect(layout.fromValue).toBe('August');
    expect(layout.toValue).toBe('August');
  });

  test('groups Deductions parent over Amount and Reason on the next header row', () => {
    const rows = [
      [
        'S. No.',
        'Name of the Employee',
        'Date of Appointment',
        'Rate of Wages',
        'Normal Wages Earned',
        'Wages Earned',
        'Wages Earned for Overtime',
        'Gross Wages Payable',
        'Deductions if Any, and Reasons Thereof',
        '',
        'Actual Wages Paid',
        'Date of Payment',
        'Signature / Thumb Impression of Employee',
        'Remarks'
      ],
      ['', '', '', '', '', '', '', '', 'Amount', 'Reason', '', '', '', ''],
      ['1', 'Test Worker', '01-08-2026', '500', '15000', '15000', '0', '15000', '0', 'NIL', '15000', '31-08-2026', '', '']
    ];
    const band = findFormXXIIIAPSEDeductionsBand(rows, 0, 14);
    expect(band).toEqual({
      parentRow: 0,
      childRow: 1,
      start: 8,
      end: 9
    });
  });

  test('gives employee name more width than serial number', () => {
    expect(formXXIIIAPSEColumnWeight('Name of the Employee')).toBeGreaterThan(
      formXXIIIAPSEColumnWeight('S. No.')
    );
  });

  test('centers S.No, appointment date and payment date values only', () => {
    expect(isFormXXIIIAPSECenterValueHeader('S. No.')).toBe(true);
    expect(isFormXXIIIAPSECenterValueHeader('Date of Appointment')).toBe(true);
    expect(isFormXXIIIAPSECenterValueHeader('Date of Payment')).toBe(true);
    expect(isFormXXIIIAPSECenterValueHeader('Name of the Employee')).toBe(false);
    expect(isFormXXIIIAPSECenterValueHeader('Rate of Wages')).toBe(false);
  });

  test('formats money display with Indian grouping and keeps decimals', () => {
    expect(isFormXXIIIAPSEMoneyHeader('Rate of Wages')).toBe(true);
    expect(isFormXXIIIAPSEMoneyHeader('Normal Wages Earned')).toBe(true);
    expect(isFormXXIIIAPSEMoneyHeader('Wages Earned')).toBe(true);
    expect(isFormXXIIIAPSEMoneyHeader('Wages Earned for Overtime')).toBe(true);
    expect(isFormXXIIIAPSEMoneyHeader('Gross Wages Payable')).toBe(true);
    expect(isFormXXIIIAPSEMoneyHeader('Amount')).toBe(true);
    expect(isFormXXIIIAPSEMoneyHeader('Actual Wages Paid')).toBe(true);
    expect(isFormXXIIIAPSEMoneyHeader('Reason')).toBe(false);
    expect(formatFormXXIIIAPSEMoneyPdfValue('45997')).toBe('45,997');
    expect(formatFormXXIIIAPSEMoneyPdfValue('5000')).toBe('5,000');
    expect(formatFormXXIIIAPSEMoneyPdfValue('125000')).toBe('1,25,000');
    expect(formatFormXXIIIAPSEMoneyPdfValue('45997.50')).toBe('45,997.50');
    expect(formatFormXXIIIAPSEMoneyPdfValue('810.65')).toBe('810.65');
    expect(formatFormXXIIIAPSEMoneyPdfValue('NIL')).toBe('NIL');
  });

  test('drops the unlabeled column between Signature and Remarks without removing those columns', () => {
    const rows = [
      [
        'S. No.',
        'Name of the Employee',
        'Date of Appointment',
        'Rate of Wages',
        'Normal Wages Earned',
        'Wages Earned',
        'Wages Earned for Overtime',
        'Gross Wages Payable',
        'Deductions if Any, and Reasons Thereof',
        '',
        'Actual Wages Paid',
        'Date of Payment',
        'Signature / Thumb Impression of Employee',
        '',
        'Remarks'
      ],
      ['', '', '', '', '', '', '', '', 'Amount', 'Reason', '', '', '', '', ''],
      [
        '1',
        'Test Worker',
        '10 Aug 2026',
        '810.65',
        '21077',
        '21077',
        '',
        '21077',
        '1465',
        '',
        '19612',
        '31-08-2026',
        '',
        '',
        ''
      ]
    ];
    const trimmed = trimFormXXIIIAPSEEmptyPdfColumns(rows, 0, 1, 15);
    expect(trimmed.colCount).toBe(14);
    expect(trimmed.rows[0][12]).toMatch(/signature/i);
    expect(trimmed.rows[0][13]).toMatch(/^remarks$/i);
    expect(trimmed.rows[2][0]).toBe('1');
    expect(trimmed.rows[2][3]).toBe('810.65');
    expect(trimmed.rows[2][11]).toBe('31-08-2026');
  });

  test('PDF header model uses Form XXIII AP S&E layout flags', () => {
    const model = statutoryDraftPdfTestUtils.buildStatutoryPdfHeaderModel(
      meta,
      [
        [
          'S. No.',
          'Name of the Employee',
          'Deductions if Any, and Reasons Thereof',
          '',
          'Remarks'
        ],
        ['', '', 'Amount', 'Reason', '']
      ],
      0,
      'Form XXIII – Register of Wages'
    );
    expect(model.formXXIIIAPSE).toBe(true);
    expect(model.formXXIIIAPSEAdminLayout.establishmentValue).toContain('Nimbagallu Site');
    expect(model.titles.some((line) => /vide\s+rule\s*29/i.test(line))).toBe(true);
  });
});
