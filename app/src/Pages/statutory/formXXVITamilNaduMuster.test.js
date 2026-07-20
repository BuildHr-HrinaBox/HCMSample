import {
  FORM_XXVI_TN_REFERENCE,
  FORM_XXVI_TN_SUBTITLE,
  FORM_XXVI_TN_TITLE,
  applyFormXXVITamilNaduAutofillFromSite,
  applyFormXXVITamilNaduPaidDaysToMappedRows,
  buildFormXXVITamilNaduWorksiteText,
  enrichFormXXVITamilNaduDisplayHeader,
  getFormXXVITamilNaduEmployeeName,
  isFormXXVITamilNaduClraContext,
  isFormXXVITamilNaduDayHeaderKey,
  isFormXXVITamilNaduHeaderFieldLayoutFormHeader,
  isFormXXVITamilNaduRateOfWagesHeaderKey,
  detectFormXXVITamilNaduDayColumnMap,
  readFormXXVITamilNaduCellValue,
  stripFormXXVITamilNaduNonTableHeaders,
} from './formXXVITamilNaduMuster';
import { isFormXVIIITamilNaduClraContext } from './formXVIIITamilNaduWagesMuster';

describe('formXXVITamilNaduMuster heading', () => {
  test('Form_26_-_TamilNadu.xlsx is Form XXVI CLRA, not Form XVIII', () => {
    const fileName = 'Form_26_-_TamilNadu.xlsx';
    const sheetText =
      'Form XVIII – Register of Wages-cum-Muster Roll Form of Register of Wages-cum-Muster Roll [See rule 78(1)(a)(i)]';
    const formHeader = {
      title: 'Form XVIII – Register of Wages-cum-Muster Roll',
      subtitle: 'Form of Register of Wages-cum-Muster Roll',
      reference: '[See rule 78(1)(a)(i)]'
    };

    expect(
      isFormXVIIITamilNaduClraContext(formHeader, null, fileName, [], sheetText)
    ).toBe(false);
    expect(
      isFormXXVITamilNaduClraContext(formHeader, null, fileName, [], sheetText)
    ).toBe(true);
  });

  test('enrichFormXXVITamilNaduDisplayHeader replaces wrong Form XVIII title', () => {
    const enriched = enrichFormXXVITamilNaduDisplayHeader(
      {
        title: 'Form XVIII – Register of Wages-cum-Muster Roll',
        subtitle: 'Form of Register of Wages-cum-Muster Roll',
        reference: '[See rule 78(1)(a)(i)]'
      },
      { formName: 'Form 26', state: 'Tamil Nadu' },
      'Form_26_-_TamilNadu.xlsx',
      [],
      'Form XVIII Register of Wages-cum-Muster Roll'
    );

    expect(enriched.title).toBe(FORM_XXVI_TN_TITLE);
    expect(enriched.subtitle).toBe(FORM_XXVI_TN_SUBTITLE);
    expect(enriched.reference).toBe(FORM_XXVI_TN_REFERENCE);
    expect(isFormXXVITamilNaduHeaderFieldLayoutFormHeader(enriched)).toBe(true);
    expect(enriched.fields.some((f) => f.key === 'form_xxvi_worksite')).toBe(true);
    expect(
      enriched.fields.find((f) => f.key === 'form_xxvi_worksite')?.label
    ).toBe('Name and Location of Worksite');
  });

  test('Form XVIII Tamil Nadu filename still matches Form XVIII', () => {
    expect(
      isFormXVIIITamilNaduClraContext(
        { title: 'Form XVIII – Register of Wages-cum-Muster Roll' },
        null,
        'Form_XVIII_-_TamilNadu.xlsx',
        [],
        'Register of Wages-cum-Muster Roll'
      )
    ).toBe(true);
  });
});

describe('formXXVITamilNaduMuster worksite + employee name', () => {
  test('buildFormXXVITamilNaduWorksiteText joins site Name and Location name', () => {
    expect(
      buildFormXXVITamilNaduWorksiteText(
        { siteName: 'Chennai Site', location: 'Teynampet' },
        {}
      )
    ).toBe('Chennai Site, Teynampet');
  });

  test('worksite accepts EmployeeName + LocationName extras', () => {
    expect(
      buildFormXXVITamilNaduWorksiteText(
        {},
        { EmployeeName: 'Buildgr Works', LocationName: 'Madurai' }
      )
    ).toBe('Buildgr Works, Madurai');
  });

  test('getFormXXVITamilNaduEmployeeName prefers EmployeeName', () => {
    expect(
      getFormXXVITamilNaduEmployeeName({ EmployeeName: 'Rajesh Kumar', Name: 'Other' })
    ).toBe('Rajesh Kumar');
  });

  test('applyFormXXVITamilNaduAutofillFromSite fills worksite', () => {
    const out = applyFormXXVITamilNaduAutofillFromSite(
      { form_xxvi_worksite: '' },
      {
        siteName: 'ELCOT SEZ',
        locationName: 'Madurai',
        contractorText: 'Contractor LLC',
        principalEmployerText: 'Principal Co',
        monthName: 'July',
        year: '2026'
      }
    );
    expect(out.form_xxvi_worksite).toBe('ELCOT SEZ, Madurai');
    expect(out.form_xxvi_contractor).toBe('Contractor LLC');
    expect(out.form_xxvi_principal_employer).toBe('Principal Co');
    expect(out.form_x_month).toBe('July');
    expect(out.form_x_year).toBe('2026');
  });

  test('stripFormXXVITamilNaduNonTableHeaders removes worksite column', () => {
    const headers = [
      'Name and Location of Worksite :',
      'Name of the Workman',
      'Age & Sex',
      'Number of Days Worked'
    ];
    expect(stripFormXXVITamilNaduNonTableHeaders(headers)).toEqual([
      'Name of the Workman',
      'Age & Sex',
      'Number of Days Worked'
    ]);
  });

  test('applyFormXXVITamilNaduPaidDaysToMappedRows sets Paid_days from SamplePayroll shape', () => {
    const headers = ['Name of the Workman', 'Number of Days Worked', 'Rate of Wages'];
    const rows = [
      { 'Name of the Workman': 'Rajeshkumar', 'Number of Days Worked': '', 'Rate of Wages': '' },
      { 'Name of the Workman': 'Raja', 'Number of Days Worked': '', 'Rate of Wages': '' }
    ];
    const payroll = [
      { employee_name: 'Rajeshkumar', paid_days: 26, gross_pay: 25000 },
      { employee_name: 'Raja', Paid_days: 24, gross: 22000 }
    ];
    const hits = applyFormXXVITamilNaduPaidDaysToMappedRows(rows, headers, payroll, {
      overwrite: true
    });
    expect(hits.paidDaysHits).toBe(2);
    expect(hits.rateHits).toBe(2);
    expect(rows[0]['Number of Days Worked']).toBe('26');
    expect(rows[0]['Rate of Wages']).toBe('25000');
    expect(rows[1]['Number of Days Worked']).toBe('24');
    expect(rows[1]['Rate of Wages']).toBe('22000');
  });

  test('applyFormXXVITamilNaduPaidDaysToMappedRows uses resolvePayrollRow when provided', () => {
    const headers = ['Name of the Workman', 'Number of Days Worked', 'Rate of Wages'];
    const rows = [
      { 'Name of the Workman': 'Venkatesan', 'Number of Days Worked': '', 'Rate of Wages': '' }
    ];
    const employees = [{ EmployeeName: 'Venkatesan', EmployeeID: 'E1' }];
    const hits = applyFormXXVITamilNaduPaidDaysToMappedRows(rows, headers, [], {
      overwrite: true,
      employeesForMapping: employees,
      resolvePayrollRow: () => ({ employee_name: 'Venkatesan', paidDays: 28, gross_pay: 28000 })
    });
    expect(hits.paidDaysHits).toBe(1);
    expect(hits.rateHits).toBe(1);
    expect(rows[0]['Number of Days Worked']).toBe('28');
    expect(rows[0]['Rate of Wages']).toBe('28000');
  });

  test('Rate of Wages is not treated as a day header', () => {
    expect(isFormXXVITamilNaduRateOfWagesHeaderKey('Rate of Wages')).toBe(true);
    expect(isFormXXVITamilNaduDayHeaderKey('Rate of Wages')).toBe(false);
    expect(isFormXXVITamilNaduDayHeaderKey('10_1')).toBe(true);
    const row = {
      'Rate of Wages': '',
      '10_1': 'A',
      '10_2': 'WO',
      'Number of Days Worked': '26'
    };
    expect(readFormXXVITamilNaduCellValue(row, 'Rate of Wages')).toBe('');
    expect(readFormXXVITamilNaduCellValue(row, '10_1')).toBe('A');
  });

  test('detectFormXXVITamilNaduDayColumnMap ignores column-index (1)-(14) row', () => {
    // Row 1: identity column numbers 1–14 at cols 1–14
    // Row 2: real days 1–31 starting at col 11
    const getCell = (r, c) => {
      if (r === 1 && c >= 1 && c <= 14) return String(c);
      if (r === 2 && c >= 11 && c <= 41) return String(c - 10);
      return '';
    };
    const map = detectFormXXVITamilNaduDayColumnMap(getCell, 1, 1, 50);
    expect(map.get(1)).toBe(11);
    expect(map.get(31)).toBe(41);
    expect(map.size).toBe(31);
  });
});
