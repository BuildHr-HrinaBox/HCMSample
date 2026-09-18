import {
  applyFormBRajasthanEmployeeToRow,
  enrichFormBRajasthanPayrollRows,
  filterFormBRajasthanExportRows,
  findFormBRajasthanPayrollRowByName,
  formBRJNamesMatch,
  formBRJPayrollRowHasSampleDeductionFields,
  formBRJRowHasNoDaysWorked,
  isFormBRajasthanContext,
  isFormBRJPfHeader,
  isFormBRJRateOfWageHeader,
  isFormBRJVoluntaryPfHeader,
  readEmployeeFullName,
  readFormBRJPersonNameParts,
  resolveFormBRajasthanPayrollFields,
  resolveFormBRajasthanPayrollRowForEmployee,
  resolveFormBRajasthanPayrollRowsForAutofill,
  sanitizeFormBRajasthanAbsentWageRow,
  sanitizeFormBRajasthanOrphanAllowanceRow,
} from './formBRajasthan';
import { collectEmployeeNameCandidates } from './formXIXMPWageSlip';

describe('Form B Rajasthan FORMAT OF WAGE REGISTER payroll', () => {
  test('builds People name from FirstName + MiddleName + LastName', () => {
    expect(
      readEmployeeFullName({
        FirstName: 'Ram',
        MiddleName: 'Kumar',
        LastName: 'Sharma',
      })
    ).toBe('Ram Kumar Sharma');
    expect(
      collectEmployeeNameCandidates({
        FirstName: 'Ram',
        MiddleName: 'Kumar',
        LastName: 'Sharma',
      })
    ).toContain('ram kumar sharma');
  });

  test('matches Sample Payroll only when Employee ID agrees (not name alone)', () => {
    const emp = {
      FirstName: 'Ram',
      MiddleName: 'Kumar',
      LastName: 'Sharma',
      EmployeeID: 'VE1001',
    };
    const payroll = {
      employee_name: 'Ram Kumar Sharma',
      employee_id: 'VE1001',
      paid_days: 26,
      basic: 12000,
      hra: 3000,
      gross_pay: 18000,
      pf: 1800,
      voluntary_provident_fund: 500,
    };
    expect(formBRJNamesMatch(emp, payroll)).toBe(true);
    expect(formBRJNamesMatch(emp, { employee_name: 'Ram Sharma' })).toBe(false);
    expect(formBRJNamesMatch({ FirstName: 'Ram', LastName: 'Sharma' }, { employee_name: 'Ram Sharma' })).toBe(
      true
    );
    expect(findFormBRajasthanPayrollRowByName(emp, [payroll, { employee_name: 'Other Person' }])).toEqual(
      payroll
    );
    // Same name, no Employee ID on People → do not bind wages.
    expect(
      findFormBRajasthanPayrollRowByName(
        { FirstName: 'Ram', MiddleName: 'Kumar', LastName: 'Sharma' },
        [payroll]
      )
    ).toBeNull();
  });

  test('when two same-name people exist, matches only by Employee ID', () => {
    const siteA = {
      FirstName: 'Ram',
      LastName: 'Sharma',
      EmployeeID: 'VE1001',
      Work_location: 'RJ-Fatehgarh-2',
    };
    const siteBPayroll = {
      employee_name: 'Ram Sharma',
      employee_id: 'VE2002',
      work_location: 'MH-Pune',
      paid_days: 20,
      gross_pay: 10000,
    };
    const siteAPayroll = {
      employee_name: 'Ram Sharma',
      employee_id: 'VE1001',
      work_location: 'RJ-Fatehgarh-2',
      paid_days: 26,
      gross_pay: 18000,
    };
    expect(findFormBRajasthanPayrollRowByName(siteA, [siteBPayroll, siteAPayroll])).toEqual(siteAPayroll);

    // Location without Employee ID must not pick payroll by name.
    const byLocationOnly = {
      FirstName: 'Ram',
      LastName: 'Sharma',
      Work_location: 'MH-Pune',
    };
    expect(findFormBRajasthanPayrollRowByName(byLocationOnly, [siteBPayroll, siteAPayroll])).toBeNull();
  });

  test('same name at same location requires Employee ID (does not guess)', () => {
    const emp = {
      FirstName: 'Ram',
      LastName: 'Sharma',
      EmployeeID: 'VE1001',
      Work_location: 'RJ-Fatehgarh-2',
    };
    const twinA = {
      employee_name: 'Ram Sharma',
      employee_id: 'VE1001',
      work_location: 'RJ-Fatehgarh-2',
      paid_days: 26,
    };
    const twinB = {
      employee_name: 'Ram Sharma',
      employee_id: 'VE1002',
      work_location: 'RJ-Fatehgarh-2',
      paid_days: 20,
    };
    expect(findFormBRajasthanPayrollRowByName(emp, [twinB, twinA])).toEqual(twinA);

    const noId = { FirstName: 'Ram', LastName: 'Sharma', Work_location: 'RJ-Fatehgarh-2' };
    expect(findFormBRajasthanPayrollRowByName(noId, [twinA, twinB])).toBeNull();
  });

  test('Karthik P twins: never bind the other Employee ID by name alone', () => {
    const karthikA = { FirstName: 'Karthik', LastName: 'P', EmployeeID: 'VE0101' };
    const karthikBPayroll = {
      employee_name: 'Karthik P',
      employee_id: 'VE0202',
      paid_days: 22,
      gross_pay: 15000,
    };
    const karthikAPayroll = {
      employee_name: 'Karthik P',
      employee_id: 'VE0101',
      paid_days: 26,
      gross_pay: 18000,
    };
    expect(findFormBRajasthanPayrollRowByName(karthikA, [karthikBPayroll, karthikAPayroll])).toEqual(
      karthikAPayroll
    );
    // Only the other Karthik is in payroll → do not attach wrong ID wages.
    expect(findFormBRajasthanPayrollRowByName(karthikA, [karthikBPayroll])).toBeNull();
  });

  test('Karthik P VE1428 must not take VE948 May payroll by name', () => {
    const karthikAug = {
      FirstName: 'Karthik',
      LastName: 'P',
      EmployeeID: 'VE1428',
      Zoho_ID: '999001',
    };
    const mayOtherTwin = {
      employee_name: 'Karthik P',
      employee_id: 'VE948',
      paid_days: 30,
      basic: 88238,
      rate_of_wage: 6787.54,
      gross_pay: 88238,
      payrollMonth: '2026-05',
    };
    expect(findFormBRajasthanPayrollRowByName(karthikAug, [mayOtherTwin])).toBeNull();
    expect(resolveFormBRajasthanPayrollRowForEmployee(karthikAug, [mayOtherTwin])).toBeNull();

    const headers = ['Name', 'Rate of Wage', 'No. Of Days worked', 'Basic'];
    const mapped = [
      {
        Name: 'Karthik P',
        'Rate of Wage': '6787.54',
        'No. Of Days worked': '30',
        Basic: '88238',
        __employeeLookupId: 'VE1428',
      },
    ];
    const hits = enrichFormBRajasthanPayrollRows(mapped, [karthikAug], headers, {
      payrollRows: [mayOtherTwin],
      overwrite: true,
    });
    expect(hits).toBe(0);
    expect(String(mapped[0]['Rate of Wage'] || '').trim()).toBe('');
    expect(String(mapped[0]['No. Of Days worked'] || '').trim()).toBe('');
    expect(String(mapped[0].Basic || '').trim()).toBe('');
  });

  test('Karthik P with People ID never binds name-only payroll (no employee_id on payroll)', () => {
    const karthikA = { FirstName: 'Karthik', LastName: 'P', EmployeeID: 'VE0101' };
    const nameOnlyPayroll = {
      employee_name: 'Karthik P',
      paid_days: 30,
      basic: 88238,
      rate_of_wage: 6787.54,
    };
    expect(findFormBRajasthanPayrollRowByName(karthikA, [nameOnlyPayroll])).toBeNull();
    expect(resolveFormBRajasthanPayrollRowForEmployee(karthikA, [nameOnlyPayroll])).toBeNull();
  });

  test('clears orphan HRA / VPF when Rate and Days were not filled from payroll', () => {
    const headers = [
      'Name',
      'Rate of Wage',
      'No. Of Days worked',
      'Basic',
      'HRA',
      'Voluntary Provident Fund',
    ];
    const row = sanitizeFormBRajasthanOrphanAllowanceRow(
      {
        Name: 'Prakash Jadhav',
        'Rate of Wage': '',
        'No. Of Days worked': '',
        Basic: '',
        HRA: '8157',
        'Voluntary Provident Fund': '1165',
      },
      headers
    );
    expect(row.HRA).toBe('');
    expect(row['Voluntary Provident Fund']).toBe('');
  });

  test('enrichFormBRajasthanPayrollRows does not name-only backfill after ID miss', () => {
    const headers = [
      'Sr.No. In Employer/ Workman/ Worker Register',
      'Name',
      'Rate of Wage',
      'No. Of Days worked',
      'Basic',
    ];
    const emp = { FirstName: 'Karthik', LastName: 'P', EmployeeID: 'VE0101' };
    const mapped = [
      {
        Name: 'Karthik P',
        __employeeLookupName: 'Karthik P',
        __employeeLookupId: 'VE0101',
      },
    ];
    const payrollRows = [
      {
        employee_name: 'Karthik P',
        employee_id: 'VE0202',
        paid_days: 30,
        basic: 88238,
        rate_of_wage: 6787.54,
      },
    ];
    const hits = enrichFormBRajasthanPayrollRows(mapped, [emp], headers, {
      payrollRows,
      overwrite: true,
    });
    expect(hits).toBe(0);
    expect(String(mapped[0]['Rate of Wage'] || '').trim()).toBe('');
    expect(String(mapped[0].Basic || mapped[0]['Basic'] || '').trim()).toBe('');
  });

  test('refuses to guess when same name appears twice with no id or location', () => {
    const emp = { FirstName: 'Ram', LastName: 'Sharma' };
    expect(
      findFormBRajasthanPayrollRowByName(emp, [
        { employee_name: 'Ram Sharma', paid_days: 10 },
        { employee_name: 'Ram Sharma', paid_days: 26 },
      ])
    ).toBeNull();
  });

  test('detects Form B_MH context from filename', () => {
    expect(
      isFormBRajasthanContext(
        { title: 'FORM B', subtitle: 'FORMAT FOR WAGE REGISTER' },
        { formFileName: 'Form_B_MH_-_Maharashtra.xlsx', formName: 'FORMAT FOR WAGE REGISTER' },
        'Form_B_MH_-_Maharashtra.xlsx',
        ''
      )
    ).toBe(true);
  });

  test('detects Form_B_MH from filename alone (underscore after B)', () => {
    expect(
      isFormBRajasthanContext(
        {},
        { formFileName: 'Form_B_MH_-_Maharashtra.xlsx' },
        'Form_B_MH_-_Maharashtra.xlsx',
        '',
        []
      )
    ).toBe(true);
  });

  test('does not match first-name-only', () => {
    expect(
      formBRJNamesMatch(
        { FirstName: 'Suryakanta', LastName: 'Jana' },
        { employee_name: 'Suryakanta' }
      )
    ).toBe(false);
  });

  test('distinguishes PF and Voluntary Provident Fund headers', () => {
    expect(isFormBRJPfHeader('PF')).toBe(true);
    expect(isFormBRJPfHeader('Provident Fund')).toBe(true);
    expect(isFormBRJPfHeader('Voluntary Provident Fund')).toBe(false);
    expect(isFormBRJVoluntaryPfHeader('Voluntary Provident Fund')).toBe(true);
    expect(isFormBRJVoluntaryPfHeader('VPF')).toBe(true);
    expect(isFormBRJVoluntaryPfHeader('PF')).toBe(false);
    expect(isFormBRJRateOfWageHeader('Rate of Wage')).toBe(true);
  });

  test('resolveFormBRajasthanPayrollFields reads days, rate, basic, HRA, total, PF, VPF', () => {
    const fields = resolveFormBRajasthanPayrollFields({
      employee_name: 'Ram Kumar Sharma',
      paid_days: 26,
      basic: 12000,
      hra: 3000,
      da: 500,
      gross_pay: 18000,
      pf: 1800,
      voluntary_provident_fund: 500,
      net_pay: 15000,
    });
    expect(fields.paidDays).toBe(26);
    expect(fields.rateOfWage).toBe(Math.round((18000 / 26) * 100) / 100);
    expect(fields.basic).toBe(12000);
    expect(fields.da).toBe(500);
    expect(fields.hra).toBe(3000);
    expect(fields.earningsTotal).toBe(18000);
    expect(fields.pf).toBe(1800);
    expect(fields.voluntaryProvidentFund).toBe(500);
  });

  test('clears HRA / VPF / wages when No. Of Days worked is missing (absent month)', () => {
    const fields = resolveFormBRajasthanPayrollFields({
      employee_name: 'Suryakanta Jana',
      // no paid_days — person did not attend
      basic: 10000,
      hra: 4854,
      gross_pay: 0,
      voluntary_provident_fund: 1165,
      net_pay: 0,
    });
    expect(fields.paidDays).toBe('');
    expect(fields.rateOfWage).toBe('');
    expect(fields.basic).toBe('');
    expect(fields.hra).toBe('');
    expect(fields.earningsTotal).toBe('');
    expect(fields.voluntaryProvidentFund).toBe('');
    expect(fields.netPay).toBe('');
    expect(fields.paymentDate).toBe('');
  });

  test('clears wage columns when paid_days is 0', () => {
    const fields = resolveFormBRajasthanPayrollFields({
      employee_name: 'Suryakanta Jana',
      paid_days: 0,
      hra: 4854,
      voluntary_provident_fund: 1165,
      gross_pay: 19418,
    });
    expect(fields.paidDays).toBe(0);
    expect(fields.hra).toBe('');
    expect(fields.voluntaryProvidentFund).toBe('');
    expect(fields.earningsTotal).toBe('');
  });

  test('applyFormBRajasthanEmployeeToRow blanks HRA when days worked empty', () => {
    const headers = [
      'Name',
      'No. Of Days worked',
      'Basic',
      'HRA',
      'Total',
      'Voluntary Provident Fund',
    ];
    const row = applyFormBRajasthanEmployeeToRow(
      {
        Name: 'Old',
        HRA: '4854',
        'Voluntary Provident Fund': '1165',
        Total: '1',
      },
      { FirstName: 'Suryakanta', LastName: 'Jana' },
      headers,
      {
        overwrite: true,
        payrollRow: {
          employee_name: 'Suryakanta Jana',
          hra: 4854,
          voluntary_provident_fund: 1165,
          gross_pay: 0,
        },
      }
    );
    expect(row.Name).toBe('Suryakanta Jana');
    expect(row['No. Of Days worked']).toBe('');
    expect(row.Basic).toBe('');
    expect(row.HRA).toBe('');
    expect(row.Total).toBe('');
    expect(row['Voluntary Provident Fund']).toBe('');
  });

  test('export sanitize strips HRA/VPF when days worked missing (Excel download)', () => {
    const headers = [
      'Name',
      'No. Of Days worked',
      'Basic',
      'HRA',
      'Total',
      'Voluntary Provident Fund',
      'Net Payment',
    ];
    const stale = {
      Name: 'Suryakanta Jana',
      'No. Of Days worked': '',
      Basic: '',
      HRA: '4854',
      Total: '',
      'Voluntary Provident Fund': '1165',
      'Net Payment': '',
    };
    expect(formBRJRowHasNoDaysWorked(stale, headers)).toBe(true);
    const cleaned = sanitizeFormBRajasthanAbsentWageRow(stale, headers);
    expect(cleaned.HRA).toBe('');
    expect(cleaned['Voluntary Provident Fund']).toBe('');
    const exported = filterFormBRajasthanExportRows([stale], headers);
    expect(exported).toHaveLength(1);
    expect(exported[0].HRA).toBe('');
    expect(exported[0]['Voluntary Provident Fund']).toBe('');
    expect(exported[0].Name).toBe('Suryakanta Jana');
  });

  test('applyFormBRajasthanEmployeeToRow fills wage-register columns from Sample Payroll', () => {
    const headers = [
      'Name',
      'Rate of Wage',
      'No. Of Days worked',
      'Basic',
      'DA',
      'HRA',
      'Total',
      'PF',
      'Voluntary Provident Fund',
    ];
    const row = applyFormBRajasthanEmployeeToRow(
      {},
      { FirstName: 'Ram', MiddleName: 'Kumar', LastName: 'Sharma' },
      headers,
      {
        payrollRow: {
          employee_name: 'Ram Kumar Sharma',
          paid_days: 26,
          basic: 12000,
          da: 400,
          hra: 3000,
          gross_pay: 18000,
          pf: 1800,
          voluntaryProvidentFund: 500,
        },
      }
    );
    expect(row.Name).toBe('Ram Kumar Sharma');
    expect(row['Rate of Wage']).toBe(String(Math.round((18000 / 26) * 100) / 100));
    expect(row['No. Of Days worked']).toBe('26');
    expect(row.Basic).toBe('12000');
    expect(row.DA).toBe('400');
    expect(row.HRA).toBe('3000');
    expect(row.Total).toBe('18000');
    expect(row.PF).toBe('1800');
    expect(row['Voluntary Provident Fund']).toBe('500');
    expect(readFormBRJPersonNameParts(row).fullName || row.Name).toBeTruthy();
  });

  test('clears stale template Basic when Sample Payroll has no basic for the matched name', () => {
    const headers = ['Name', 'Rate of Wage', 'No. Of Days worked', 'Basic', 'HRA', 'Total', 'PF'];
    const row = applyFormBRajasthanEmployeeToRow(
      {
        Name: 'Old',
        Basic: '67729',
        HRA: '999',
        Total: '1',
        PF: '1',
      },
      { FirstName: 'Suryakanta', LastName: 'Jana' },
      headers,
      {
        overwrite: true,
        payrollRow: {
          employee_name: 'Suryakanta Jana',
          paid_days: 17,
          gross_pay: 19418.1,
          // no basic / hra / pf — must not keep template leftovers
        },
      }
    );
    expect(row.Name).toBe('Suryakanta Jana');
    expect(row['No. Of Days worked']).toBe('17');
    expect(row['Rate of Wage']).toBe(String(Math.round((19418.1 / 26) * 100) / 100));
    expect(row.Basic).toBe('');
    expect(row.HRA).toBe('');
    expect(row.PF).toBe('');
    expect(row.Total).toBe('19418.1');
  });

  test('matches Prem Shankar Menaria on First + Middle + Last', () => {
    const emp = { FirstName: 'Prem', MiddleName: 'Shankar', LastName: 'Menaria' };
    const payroll = {
      employee_name: 'Prem Shankar Menaria',
      paid_days: 31,
      basic: 50000,
      hra: 8000,
      gross_pay: 71678.1,
      pf: 1800,
      voluntary_provident_fund: 0,
    };
    expect(formBRJNamesMatch(emp, payroll)).toBe(true);
    expect(readEmployeeFullName(emp)).toBe('Prem Shankar Menaria');
    const row = applyFormBRajasthanEmployeeToRow({}, emp, ['Name', 'Basic', 'No. Of Days worked'], {
      payrollRow: payroll,
    });
    expect(row.Name).toBe('Prem Shankar Menaria');
    expect(row.Basic).toBe('50000');
    expect(row['No. Of Days worked']).toBe('31');
  });

  test('prefers Sample Payroll rows that carry PF / VPF', () => {
    const withDeductions = {
      employee_name: 'Ram Kumar Sharma',
      paid_days: 26,
      basic: 12000,
      hra: 3000,
      gross_pay: 18000,
      pf: 1800,
      voluntary_provident_fund: 500,
      payrollMonth: '2026-05',
    };
    const withoutDeductions = {
      employee_name: 'Ram Kumar Sharma',
      paid_days: 26,
      basic: 12000,
      hra: 3000,
      gross_pay: 18000,
      payrollMonth: '2026-05',
    };
    expect(formBRJPayrollRowHasSampleDeductionFields(withDeductions)).toBe(true);
    expect(formBRJPayrollRowHasSampleDeductionFields(withoutDeductions)).toBe(false);
    const preferred = resolveFormBRajasthanPayrollRowsForAutofill([withoutDeductions], ['2026-05'], {
      cachedSampleRows: [withDeductions],
      scopedPayrollMonth: '2026-05',
    });
    expect(preferred).toHaveLength(1);
    expect(preferred[0].pf).toBe(1800);
  });

  test('does not fill May form from another month bulk row by name', () => {
    const aprilOnly = {
      employee_name: 'Suryakanta Jana',
      paid_days: 17,
      basic: 9709,
      gross_pay: 19418,
      pf: 100,
      payrollMonth: '2026-04',
    };
    const mayOther = {
      employee_name: 'Hariom Dholi',
      paid_days: 31,
      basic: 44281,
      gross_pay: 114696,
      pf: 200,
      payrollMonth: '2026-05',
    };
    const rows = resolveFormBRajasthanPayrollRowsForAutofill([], ['2026-05'], {
      bulkSampleRows: [aprilOnly, mayOther],
      scopedPayrollMonth: '2026-05',
    });
    expect(rows).toHaveLength(1);
    expect(rows[0].employee_name).toBe('Hariom Dholi');
    expect(findFormBRajasthanPayrollRowByName({ FirstName: 'Suryakanta', LastName: 'Jana' }, rows)).toBe(
      null
    );
  });

  test('rejects latest-month fallback stamped as selected month', () => {
    const aprilLatest = {
      employee_name: 'Suryakanta Jana',
      paid_days: 17,
      basic: 9709,
      gross_pay: 19418,
      pf: 100,
      payrollMonth: '2026-04',
    };
    // Polluted "May cache" that actually holds April rows — must not autofill.
    const rows = resolveFormBRajasthanPayrollRowsForAutofill([], ['2026-05'], {
      cachedSampleRows: [aprilLatest],
      scopedPayrollMonth: '2026-05',
      trustedPrimaryMonthFetch: true,
    });
    expect(rows).toHaveLength(0);
  });

  test('trusted May fetch keeps untagged May Sample Payroll rows only', () => {
    const mayUntagged = {
      employee_name: 'Hariom Dholi',
      paid_days: 31,
      basic: 44281,
      gross_pay: 114696,
      pf: 200,
    };
    const rows = resolveFormBRajasthanPayrollRowsForAutofill([], ['2026-05'], {
      cachedSampleRows: [mayUntagged],
      scopedPayrollMonth: '2026-05',
      trustedPrimaryMonthFetch: true,
    });
    expect(rows).toHaveLength(1);
    expect(rows[0].payrollMonth).toBe('2026-05');
  });
});
