import {
  FORM_XXVI_AP_TABLE_HEADERS,
  applyFormXXVIAPEmployeeToRow,
  computeFormXXVIAPOtherAllowance,
  isFormXXVIAPAppointmentLetterContext,
  isFormXXVIAPTableLayoutFormHeader,
  mergeFormXXVIAPTableRowIntoHeaderData,
  readFormXXVIAPPayrollOtherAllowance,
  resolveFormXXVIAPDateOfBirth,
  resolveFormXXVIAPEmployeeName,
  resolveFormXXVIAPHeaderFieldLayout,
  resolveFormXXVIAPTableHeaders,
} from './formXXVIAPAppointmentLetter';

describe('Form XXVI Andhra Pradesh appointment letter', () => {
  it('detects Form___XXVI_-_Andhra_Pradesh.xlsx by filename', () => {
    expect(
      isFormXXVIAPAppointmentLetterContext(
        { title: 'Form XXVI – Letter of Appointment' },
        { state: 'Andhra Pradesh', formName: 'Form XXVI – Letter of Appointment' },
        'Form___XXVI_-_Andhra_Pradesh.xlsx'
      )
    ).toBe(true);
  });

  it('resolves table layout with establishment header fields and employee columns', () => {
    const layout = resolveFormXXVIAPHeaderFieldLayout(
      { headers: [] },
      null,
      {
        formHeader: { title: 'Form XXVI – Letter of Appointment' },
        item: { state: 'Andhra Pradesh', formName: 'Letter of Appointment' },
        fileName: 'Form___XXVI_-_Andhra_Pradesh.xlsx',
        sheetText: 'Letter of Appointment Rule 30',
      }
    );
    expect(layout).toBeTruthy();
    expect(layout.formHeader.formXXVIAPHeaderFieldLayout).toBe(true);
    expect(layout.formHeader.formXXVIAPTableLayout).toBe(true);
    expect(isFormXXVIAPTableLayoutFormHeader(layout.formHeader)).toBe(true);
    expect(layout.formHeader.fields.every((f) => f.group === 'header')).toBe(true);
    expect(resolveFormXXVIAPTableHeaders(layout.headers)).toEqual(FORM_XXVI_AP_TABLE_HEADERS);
  });

  it('maps employee and payroll data onto table columns', () => {
    const row = {};
    FORM_XXVI_AP_TABLE_HEADERS.forEach((h) => {
      row[h] = '';
    });
    Object.assign(
      row,
      applyFormXXVIAPEmployeeToRow(
        row,
        {
          EmployeeName: 'Ravi Kumar',
          FatherName: 'Suresh Kumar',
          Date_of_birth: '1990-05-15',
          Dateofjoining: '2020-01-10',
          Designation: 'Technician',
          PresentAddress: 'Hyderabad',
          Age: '35',
        },
        FORM_XXVI_AP_TABLE_HEADERS,
        {
          formatStatutoryDateDisplay: (v) => String(v || '').trim(),
          payrollRow: {
            basic_pay: 12000,
            basic: 12000,
            hra: 3000,
            dearness_allowance: 2000,
            gross_pay: 15500,
          },
        }
      )
    );
    expect(row['Name of Employee (Sri / Srimathi / Kumari)']).toBe('Ravi Kumar');
    expect(row['Son / Wife / Daughter of']).toBe('Suresh Kumar');
    expect(row.Aged).toBe('35');
    expect(row['Date of Birth']).toBe('1990-05-15');
    expect(row['Appointed as']).toBe('Technician');
    expect(row['Basic Pay']).toBe('12000');
    expect(row['Dearness Allowance']).toBe('2000');
    // Other Allowance = gross_pay − basic − hra = 15500 − 12000 − 3000
    expect(row['Other Allowance']).toBe('500');
    expect(row['Total wages per day / month / week']).toBe('15500');
    expect(row['To (Name and Address of Employee)']).toContain('Ravi Kumar');
    expect(row['To (Name and Address of Employee)']).toContain('Hyderabad');
  });

  it('resolves EmployeeName and Date_of_birth field keys', () => {
    expect(resolveFormXXVIAPEmployeeName({ EmployeeName: 'Anita Devi' })).toBe('Anita Devi');
    expect(resolveFormXXVIAPDateOfBirth({ Date_of_birth: '1992-03-01' })).toBe('1992-03-01');
  });

  it('computes Other Allowance as gross_pay − basic − hra', () => {
    expect(computeFormXXVIAPOtherAllowance(15500, 12000, 3000)).toBe(500);
    expect(computeFormXXVIAPOtherAllowance('179137', '57799', '20000')).toBe(101338);
    expect(
      readFormXXVIAPPayrollOtherAllowance({
        gross_pay: 179137,
        basic: 57799,
        hra: 20000,
      })
    ).toBe('101338');
  });

  it('merges table row values back into form_xxvi_ap_* keys for export', () => {
    const headers = FORM_XXVI_AP_TABLE_HEADERS;
    const row = Object.fromEntries(headers.map((h) => [h, '']));
    row['Name of Employee (Sri / Srimathi / Kumari)'] = 'Anita';
    row['Basic Pay'] = '9000';
    row['Dearness Allowance'] = '1000';
    row['Other Allowance'] = '500';
    row['Total wages per day / month / week'] = '10500';
    const merged = mergeFormXXVIAPTableRowIntoHeaderData(
      { form_xxvi_ap_establishment: 'ACME Est.' },
      row,
      headers
    );
    expect(merged.form_xxvi_ap_establishment).toBe('ACME Est.');
    expect(merged.form_xxvi_ap_employee_name).toBe('Anita');
    expect(merged.form_xxvi_ap_basic_pay).toBe('9000');
    expect(merged.form_xxvi_ap_dearness_allowance).toBe('1000');
    expect(merged.form_xxvi_ap_other_allowance).toBe('500');
    expect(merged.form_xxvi_ap_total_wages).toBe('10500');
  });

  it('resolves empty headers to canonical Form XXVI AP columns', () => {
    expect(resolveFormXXVIAPTableHeaders([])).toEqual(FORM_XXVI_AP_TABLE_HEADERS);
    expect(resolveFormXXVIAPTableHeaders(null)).toEqual(FORM_XXVI_AP_TABLE_HEADERS);
  });

  it('does not look like Form D structural headers (avoids Form D download mis-route)', () => {
    const joined = FORM_XXVI_AP_TABLE_HEADERS.join(' ').toLowerCase();
    // Dearness/Other allowance alone used to trip isFormDAutofillHeaders — Form D needs these markers.
    expect(joined).not.toMatch(/category of workers/);
    expect(joined).not.toMatch(/men employed|women employed/);
    expect(joined).not.toMatch(/rate of remuneration|components of remuneration/);
    expect(joined).not.toMatch(/brief description of work/);
  });
});
