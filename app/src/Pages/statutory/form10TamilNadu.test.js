import {
  applyForm10UnmatchedPayrollAmountsNil,
  buildForm10EmployeeDisplayName,
  buildForm10PayrollMatchIdentity,
  buildFormTamilNaduPayrollRowResolver,
  buildKarnatakaPayrollRowResolver,
  collectForm10RowNameParts,
  findForm10EmployeeByFirstAndLastName,
  findForm10PayrollRowByFirstAndLastName,
  findKarnatakaPayrollRowByFirstAndLastName,
  form10FirstAndLastNamesMatch,
  form10HasFirstAndLastName,
  form10PayrollRowAgreesWithEmployeeNames,
  isForm10CombinedNameHeader,
  isForm10FirstNameHeader,
  isForm10LastNameHeader,
  karnatakaFirstAndLastNamesMatch,
  readForm10PersonNameParts,
} from './form10TamilNadu';

describe('Form 10 Tamil Nadu firstname + lastname payroll mapping', () => {
  const selvaP = { FirstName: 'Selva', LastName: 'P', EmailID: 'selva.p@example.com' };
  const julyPayroll = [
    {
      first_name: 'Selva',
      last_name: 'Kumar',
      employee_name: 'Selva Kumar',
      gross_pay: 76140,
      net_pay: 72808,
    },
    {
      first_name: 'Ravi',
      last_name: 'S',
      employee_name: 'Ravi S',
      gross_pay: 49124,
      net_pay: 47053,
    },
  ];

  test('reads firstname and lastname from People and payroll rows', () => {
    expect(readForm10PersonNameParts(selvaP)).toEqual({
      firstName: 'Selva',
      lastName: 'P',
      fullName: 'Selva P',
    });
    expect(
      readForm10PersonNameParts({
        first_name: 'Selva',
        last_name: 'P',
        employee_name: 'Selva P',
      })
    ).toEqual({
      firstName: 'Selva',
      lastName: 'P',
      fullName: 'Selva P',
    });
    expect(form10HasFirstAndLastName(selvaP)).toBe(true);
    expect(form10HasFirstAndLastName({ FirstName: 'Selva' })).toBe(false);
  });

  test('Selva P does not match another Selva in July payroll', () => {
    expect(form10FirstAndLastNamesMatch(selvaP, julyPayroll[0])).toBe(false);
    expect(form10FirstAndLastNamesMatch(selvaP, { first_name: 'Selva', employee_name: 'Selva' })).toBe(
      false
    );
    expect(findForm10PayrollRowByFirstAndLastName(selvaP, julyPayroll)).toBeNull();
  });

  test('Selva P matches only the payroll row with the same firstname and lastname', () => {
    const augustRow = {
      first_name: 'Selva',
      last_name: 'P',
      employee_name: 'Selva P',
      gross_pay: 50000,
      net_pay: 45000,
    };
    expect(form10FirstAndLastNamesMatch(selvaP, augustRow)).toBe(true);
    expect(findForm10PayrollRowByFirstAndLastName(selvaP, [...julyPayroll, augustRow])).toEqual(
      augustRow
    );
  });

  test('matches payroll employee_name when first/last columns are missing', () => {
    const pay = { employee_name: 'Selva P', gross_pay: 12000 };
    expect(form10FirstAndLastNamesMatch(selvaP, pay)).toBe(true);
    expect(form10FirstAndLastNamesMatch(selvaP, { employee_name: 'Selva Kumar' })).toBe(false);
  });

  test('email/id hit is rejected when payroll firstname+lastname disagree', () => {
    expect(form10PayrollRowAgreesWithEmployeeNames(selvaP, julyPayroll[0])).toBe(false);
    expect(
      form10PayrollRowAgreesWithEmployeeNames(selvaP, {
        first_name: 'Selva',
        last_name: 'P',
      })
    ).toBe(true);
    expect(form10PayrollRowAgreesWithEmployeeNames(selvaP, { employee_id: 'VE0099' })).toBe(true);
  });

  test('display name uses firstname and lastname together', () => {
    expect(buildForm10EmployeeDisplayName(selvaP)).toBe('Selva P');
    expect(buildForm10EmployeeDisplayName({ FirstName: 'Ravi' })).toBe('Ravi');
  });

  test('collects first/last from Form 10 name columns', () => {
    const headers = ['First Name', 'Last Name', 'Normal rate of pay'];
    const parts = collectForm10RowNameParts(
      { 'First Name': 'Selva', 'Last Name': 'P', 'Normal rate of pay': '76140' },
      headers
    );
    expect(parts).toEqual({ firstName: 'Selva', lastName: 'P', fullName: 'Selva P' });
    expect(isForm10FirstNameHeader('First Name')).toBe(true);
    expect(isForm10LastNameHeader('Last Name')).toBe(true);
    expect(isForm10CombinedNameHeader('Name of the person employed')).toBe(true);
    expect(isForm10CombinedNameHeader('First Name')).toBe(false);
  });

  test('unmatched payroll amount columns become Nil', () => {
    const headers = [
      'Normal rate of pay',
      'Normal earnings',
      'Total earnings',
      'Overtime earnings',
    ];
    const row = {
      'Normal rate of pay': '76140',
      'Normal earnings': '76140',
      'Total earnings': '72808',
      'Overtime earnings': 'Nil',
    };
    applyForm10UnmatchedPayrollAmountsNil(row, headers, 'Nil');
    expect(row['Normal rate of pay']).toBe('Nil');
    expect(row['Normal earnings']).toBe('Nil');
    expect(row['Total earnings']).toBe('Nil');
    expect(row['Overtime earnings']).toBe('Nil');
  });

  test('buildForm10PayrollMatchIdentity prefers row First/Last over People record', () => {
    const headers = ['First Name', 'Last Name'];
    const identity = buildForm10PayrollMatchIdentity(
      { FirstName: 'Wrong', LastName: 'Person' },
      { 'First Name': 'Selva', 'Last Name': 'P' },
      headers
    );
    expect(identity.FirstName).toBe('Selva');
    expect(identity.LastName).toBe('P');
    expect(identity.employee_name).toBe('Selva P');
  });

  test('findForm10EmployeeByFirstAndLastName locates People by row name parts', () => {
    const employees = [{ FirstName: 'Selva', LastName: 'P' }, { FirstName: 'Ravi', LastName: 'S' }];
    const hit = findForm10EmployeeByFirstAndLastName(employees, {
      firstName: 'Selva',
      lastName: 'P',
      fullName: 'Selva P',
    });
    expect(hit).toEqual({ FirstName: 'Selva', LastName: 'P' });
  });

  test('buildFormTamilNaduPayrollRowResolver matches strict first+last only', () => {
    const payroll = [
      { first_name: 'Selva', last_name: 'Kumar', employee_name: 'Selva Kumar' },
      { first_name: 'Selva', last_name: 'P', employee_name: 'Selva P' },
    ];
    const resolve = buildFormTamilNaduPayrollRowResolver(payroll);
    const hit = resolve({ FirstName: 'Selva', LastName: 'P' });
    expect(hit).toEqual(payroll[1]);
    expect(resolve({ FirstName: 'Selva', LastName: 'P' })).toBeNull();
  });

  test('Karnataka matches People FirstName that already contains payroll first+last', () => {
    const emp = { FirstName: 'Umesh S O', Father_s_Name: 'Shivaputrappa oli' };
    const pay = { first_name: 'Umesh', last_name: 'S O', paid_days: 26, gross_pay: 50000 };
    expect(karnatakaFirstAndLastNamesMatch(emp, pay)).toBe(true);
    expect(findKarnatakaPayrollRowByFirstAndLastName(emp, [pay])).toEqual(pay);
    const resolve = buildKarnatakaPayrollRowResolver([pay, { first_name: 'Selva', last_name: 'Kumar' }]);
    expect(resolve(emp)).toEqual(pay);
    expect(resolve({ FirstName: 'Selva' })).toBeNull();
    expect(form10FirstAndLastNamesMatch(emp, pay)).toBe(false);
  });

  test('Karnataka matches payroll employee_name when first/last are not split', () => {
    const emp = { FirstName: 'Prakash', LastName: 'Talawar' };
    const pay = { employee_name: 'Prakash Talawar', gross_pay: 74992 };
    expect(karnatakaFirstAndLastNamesMatch(emp, pay)).toBe(true);
    expect(findKarnatakaPayrollRowByFirstAndLastName(emp, [pay])).toEqual(pay);
    expect(karnatakaFirstAndLastNamesMatch({ FirstName: 'Prakash' }, pay)).toBe(false);
  });
});
