import {
  FORM_XXVI_TN_REFERENCE,
  FORM_XXVI_TN_SUBTITLE,
  FORM_XXVI_TN_TITLE,
  applyFormXXVITamilNaduAutofillFromSite,
  applyFormXXVITamilNaduPaidDaysToMappedRows,
  buildFormXXVITamilNaduWorksiteText,
  enrichFormXXVITamilNaduDisplayHeader,
  getFormXXVITamilNaduEmployeeName,
  hasFormXXVITamilNaduPersonNameParts,
  isFormXXVITamilNaduClraContext,
  isFormXXVITamilNaduDayHeaderKey,
  isFormXXVITamilNaduHeaderFieldLayoutFormHeader,
  isFormXXVITamilNaduRateOfWagesHeaderKey,
  isFormXXVITamilNaduWorkmanNameHeader,
  detectFormXXVITamilNaduDayColumnMap,
  resolveFormXXVITamilNaduTableTextHeaderRow,
  ensureFormXXVITamilNaduDayColumnHeaders,
  readFormXXVITamilNaduCellValue,
  resolveFormXXVITamilNaduRateOfWages,
  stripFormXXVITamilNaduNonTableHeaders,
  scrubFormXXVITamilNaduPreJoinDayMarks,
  resolveFormXXVITamilNaduEmployeeJoinDate,
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

  test('getFormXXVITamilNaduEmployeeName joins FirstName and LastName', () => {
    expect(
      getFormXXVITamilNaduEmployeeName({ FirstName: 'Rajesh', LastName: 'Kumar' })
    ).toBe('Rajesh Kumar');
    expect(
      getFormXXVITamilNaduEmployeeName({ first_name: 'Raja', last_name: 'K' })
    ).toBe('Raja K');
    expect(
      getFormXXVITamilNaduEmployeeName({
        EmployeeName: 'Rajeshkumar',
        FirstName: 'Rajesh',
        LastName: 'Kumar'
      })
    ).toBe('Rajesh Kumar');
    expect(getFormXXVITamilNaduEmployeeName({ EmployeeName: 'OnlyFirst' })).toBe('OnlyFirst');
  });

  test('isFormXXVITamilNaduWorkmanNameHeader matches Name of the Workman', () => {
    expect(isFormXXVITamilNaduWorkmanNameHeader('Name of the Workman')).toBe(true);
    expect(isFormXXVITamilNaduWorkmanNameHeader("Father's / Husband's Name")).toBe(false);
    expect(isFormXXVITamilNaduWorkmanNameHeader('Name and Address of the Contractor')).toBe(false);
  });

  test('applyFormXXVITamilNaduPaidDaysToMappedRows writes FirstName + LastName into Name of the Workman', () => {
    const headers = ['Name of the Workman', 'Number of Days Worked', 'Rate of Wages'];
    const rows = [
      { 'Name of the Workman': 'Rajeshkumar', 'Number of Days Worked': '', 'Rate of Wages': '' }
    ];
    applyFormXXVITamilNaduPaidDaysToMappedRows(
      rows,
      headers,
      [
        {
          employee_name: 'Rajeshkumar',
          first_name: 'Rajesh',
          last_name: 'Kumar',
          paid_days: 26,
          gross_pay: 25000
        }
      ],
      {
        overwrite: true,
        employeesForMapping: [{ FirstName: 'Rajesh', LastName: 'Kumar' }]
      }
    );
    expect(rows[0]['Name of the Workman']).toBe('Rajesh Kumar');
    expect(rows[0]['Number of Days Worked']).toBe('26');
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
      {
        employee_name: 'Rajeshkumar',
        first_name: 'Rajesh',
        last_name: 'Kumar',
        paid_days: 26,
        gross_pay: 25000
      },
      {
        employee_name: 'Raja',
        FirstName: 'Raja',
        LastName: 'K',
        Paid_days: 24,
        gross: 22000
      }
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
    const employees = [{ FirstName: 'Venkat', LastName: 'Esan', EmployeeID: 'E1' }];
    const hits = applyFormXXVITamilNaduPaidDaysToMappedRows(rows, headers, [], {
      overwrite: true,
      employeesForMapping: employees,
      resolvePayrollRow: () => ({
        employee_name: 'Venkatesan',
        first_name: 'Venkat',
        last_name: 'Esan',
        paidDays: 28,
        gross_pay: 28000
      })
    });
    expect(hits.paidDaysHits).toBe(1);
    expect(hits.rateHits).toBe(1);
    expect(rows[0]['Number of Days Worked']).toBe('28');
    expect(rows[0]['Rate of Wages']).toBe('28000');
  });

  test('clears Rate of Wages and Days Worked when person has no SamplePayroll row', () => {
    const headers = ['Name of the Workman', 'Number of Days Worked', 'Rate of Wages'];
    const rows = [
      {
        'Name of the Workman': 'Prabakaran',
        'Number of Days Worked': '31',
        'Rate of Wages': '80990'
      },
      {
        'Name of the Workman': 'NoPayrollPerson',
        'Number of Days Worked': '31',
        'Rate of Wages': '80990'
      }
    ];
    const payroll = [
      {
        employee_name: 'Prabakaran',
        first_name: 'Prabakaran',
        last_name: 'D',
        paid_days: 26,
        gross_pay: 25000
      }
    ];
    const hits = applyFormXXVITamilNaduPaidDaysToMappedRows(rows, headers, payroll, {
      overwrite: true,
      employeesForMapping: [
        { FirstName: 'Prabakaran', LastName: 'D' },
        { FirstName: 'NoPayroll', LastName: 'Person' }
      ],
      resolvePayrollRow: (emp) => {
        const name = `${emp?.FirstName || ''} ${emp?.LastName || ''}`.trim().toLowerCase();
        if (name.includes('prabakaran')) return payroll[0];
        return null;
      }
    });
    expect(hits.paidDaysHits).toBe(1);
    expect(hits.rateHits).toBe(1);
    expect(rows[0]['Number of Days Worked']).toBe('26');
    expect(rows[0]['Rate of Wages']).toBe('25000');
    expect(rows[1]['Number of Days Worked']).toBe('');
    expect(rows[1]['Rate of Wages']).toBe('');
  });

  test('does not assign another employee payroll via substring name match', () => {
    const headers = ['Name of the Workman', 'Number of Days Worked', 'Rate of Wages'];
    const rows = [
      { 'Name of the Workman': 'Raja', 'Number of Days Worked': '99', 'Rate of Wages': '1' }
    ];
    const payroll = [
      {
        employee_name: 'Rajeshkumar',
        first_name: 'Rajesh',
        last_name: 'Kumar',
        paid_days: 26,
        gross_pay: 25000
      }
    ];
    applyFormXXVITamilNaduPaidDaysToMappedRows(rows, headers, payroll, { overwrite: true });
    expect(rows[0]['Number of Days Worked']).toBe('');
    expect(rows[0]['Rate of Wages']).toBe('');
  });

  test('Rate of Wages requires firstname and lastname', () => {
    expect(
      hasFormXXVITamilNaduPersonNameParts({ FirstName: 'Prabakaran', LastName: 'D' })
    ).toBe(true);
    expect(hasFormXXVITamilNaduPersonNameParts({ FirstName: 'Prabakaran' })).toBe(false);
    expect(
      resolveFormXXVITamilNaduRateOfWages(
        { gross_pay: 80990, first_name: 'Prabakaran' },
        { FirstName: 'Prabakaran' }
      )
    ).toBe('');
    expect(
      resolveFormXXVITamilNaduRateOfWages(
        { gross_pay: 80990, first_name: 'Prabakaran', last_name: 'D' },
        null
      )
    ).toBe('80990');

    const headers = ['Name of the Workman', 'Rate of Wages'];
    const rows = [{ 'Name of the Workman': 'Prabakaran', 'Rate of Wages': '80990' }];
    const hits = applyFormXXVITamilNaduPaidDaysToMappedRows(
      rows,
      headers,
      [{ employee_name: 'Prabakaran', first_name: 'Prabakaran', gross_pay: 80990 }],
      {
        overwrite: true,
        employeesForMapping: [{ FirstName: 'Prabakaran' }]
      }
    );
    expect(hits.rateHits).toBe(0);
    expect(rows[0]['Rate of Wages']).toBe('');
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

  test('resolveFormXXVITamilNaduTableTextHeaderRow prefers descriptive row over numeric band', () => {
    const getCell = (r, c) => {
      if (r === 8 && c === 1) return 'Sr. No.';
      if (r === 8 && c === 2) return 'Name of the Workman';
      if (r === 8 && c === 10) return 'Daily hours of work';
      if (r === 9 && c >= 1 && c <= 9) return String(c);
      if (r === 9 && c === 10) return '10';
      if (r === 10 && c >= 10 && c <= 40) return String(c - 9);
      return '';
    };
    expect(resolveFormXXVITamilNaduTableTextHeaderRow(getCell, 9, 20, 50)).toBe(8);
    expect(resolveFormXXVITamilNaduTableTextHeaderRow(getCell, 8, 20, 50)).toBe(8);
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

  test('ensureFormXXVITamilNaduDayColumnHeaders pads partial 1–14 day band to 1–31', () => {
    const headers = [
      'Serial Number',
      'Name of the Workman',
      'Rate of Wages',
      ...Array.from({ length: 14 }, (_, i) => `10_${i + 1}`),
      'Number of Days Worked',
      'Signature or Thumb impression of the Workman'
    ];
    const ensured = ensureFormXXVITamilNaduDayColumnHeaders(headers);
    const dayKeys = ensured.filter((h) => /^10_(\d{1,2})$/.test(h));
    expect(dayKeys).toHaveLength(31);
    expect(dayKeys[0]).toBe('10_1');
    expect(dayKeys[30]).toBe('10_31');
    expect(ensured.indexOf('Rate of Wages')).toBeLessThan(ensured.indexOf('10_1'));
    expect(ensured.indexOf('10_31')).toBeLessThan(ensured.indexOf('Number of Days Worked'));
  });

  test('ensureFormXXVITamilNaduDayColumnHeaders keeps a full 31-day band and injects Rate of Wages', () => {
    const headers = [
      'Name of the Workman',
      ...Array.from({ length: 31 }, (_, i) => `10_${i + 1}`),
      'Number of Days Worked'
    ];
    const ensured = ensureFormXXVITamilNaduDayColumnHeaders(headers);
    expect(ensured.filter((h) => /^10_/.test(h))).toHaveLength(31);
    expect(ensured).toContain('Rate of Wages');
    expect(ensured.indexOf('Rate of Wages')).toBeLessThan(ensured.indexOf('10_1'));
    expect(ensured.indexOf('10_31')).toBeLessThan(ensured.indexOf('Number of Days Worked'));
  });

  test('resolveFormXXVITamilNaduEmployeeJoinDate prefers Zoho DOJ over row entry date', () => {
    const row = { 'Date of Entry into Service': '11 Aug 2026' };
    const emp = { Dateofjoining: '17-Aug-2026' };
    const join = resolveFormXXVITamilNaduEmployeeJoinDate(emp, row, []);
    expect(join?.getDate()).toBe(17);
    expect(join?.getMonth()).toBe(7);
  });

  test('scrubFormXXVITamilNaduPreJoinDayMarks clears days before Date of Entry', () => {
    const headers = ensureFormXXVITamilNaduDayColumnHeaders([
      'Name of the Workman',
      'Date of Entry into Service',
      ...Array.from({ length: 31 }, (_, i) => `10_${i + 1}`),
      'Number of Days Worked'
    ]);
    const row = {
      'Name of the Workman': 'Selva P',
      'Date of Entry into Service': '17 Aug 2026',
      ...Object.fromEntries(Array.from({ length: 31 }, (_, i) => [`10_${i + 1}`, 'P'])),
    };
    scrubFormXXVITamilNaduPreJoinDayMarks([row], headers, '2026-08', {
      form_x_month: 'August',
      form_x_year: '2026'
    });
    expect(row['10_16']).toBe('');
    expect(row['10_17']).toBe('P');
    expect(row['10_1']).toBe('');
  });

  test('scrubFormXXVITamilNaduPreJoinDayMarks uses employee DOJ when row date differs', () => {
    const headers = ensureFormXXVITamilNaduDayColumnHeaders([
      'Name of the Workman',
      'Date of Entry into Service',
      ...Array.from({ length: 31 }, (_, i) => `10_${i + 1}`),
    ]);
    const row = {
      'Name of the Workman': 'Selva P',
      'Date of Entry into Service': '11 Aug 2026',
      ...Object.fromEntries(Array.from({ length: 31 }, (_, i) => [`10_${i + 1}`, 'P'])),
    };
    scrubFormXXVITamilNaduPreJoinDayMarks([row], headers, '2026-08', { form_x_month: 'August', form_x_year: '2026' }, {
      employees: [{ Dateofjoining: '17-Aug-2026' }],
    });
    expect(row['10_13']).toBe('');
    expect(row['10_17']).toBe('P');
  });

  test('short identity prefix still maps Age / wages / attendance by alias (download model)', () => {
    // Regression: only Serial/Name/Age before day band used to push Age into Rate of Wages
    // and leave Daily hours empty on PDF/Excel download.
    const headers = ensureFormXXVITamilNaduDayColumnHeaders([
      'Serial Number',
      'Name of the Workman',
      'Age and Sex',
      ...Array.from({ length: 31 }, (_, i) => `10_${i + 1}`),
      'Number of Days Worked'
    ]);
    expect(headers).toContain('Rate of Wages');
    expect(headers.indexOf('Rate of Wages')).toBeLessThan(headers.indexOf('10_1'));
    expect(headers.filter((h) => /^10_/.test(h))).toHaveLength(31);

    const row = {
      'Serial Number': '1',
      'Name of the Workman': 'Selva P',
      'Age and Sex': '26 Male',
      'Age & Sex': '26 Male',
      'Rate of Wages': '8965',
      '10_1': 'P',
      '10_2': 'WO',
      '10_3': 'CL',
      'Number of Days Worked': '17'
    };
    expect(readFormXXVITamilNaduCellValue(row, 'Age and Sex')).toMatch(/26/);
    expect(readFormXXVITamilNaduCellValue(row, 'Rate of Wages')).toBe('8965');
    expect(readFormXXVITamilNaduCellValue(row, '10_1')).toBe('P');
    expect(readFormXXVITamilNaduCellValue(row, '10_2')).toBe('WO');
    expect(readFormXXVITamilNaduCellValue(row, 'Number of Days Worked')).toBe('17');
  });
});
