import {
  FORM_XXIX_TN_GROUP_ADVANCE,
  FORM_XXIX_TN_GROUP_DAMAGE,
  FORM_XXIX_TN_GROUP_FINES,
  FORM_XXIX_TN_NIL,
  FORM_XXIX_TN_TABLE_HEADERS,
  applyFormXXIXTamilNaduAutofillFromSite,
  applyFormXXIXTamilNaduNilToMappedRows,
  enrichFormXXIXTamilNaduDisplayHeader,
  ensureFormXXIXTamilNaduMonthYearHeaderFields,
  finalizeFormXXIXTamilNaduHeaderFields,
  isFormXXIXTamilNaduContext,
  isFormXXIXTamilNaduEmployeeNumberHeader,
  isFormXXIXTamilNaduHeaderFieldLayoutFormHeader,
  isFormXXIXTamilNaduNilDefaultHeader,
  resolveFormXXIXTamilNaduEmployeeId,
  resolveFormXXIXTamilNaduTableHeaders,
  looksLikeFormXXIXTamilNaduTableHeaders,
  sanitizeFormXXIXTamilNaduColumnGroupLabels,
  stripFormXXIXTamilNaduMonthYearFromTableHeaders,
} from './formXXIXTamilNadu';

describe('Form XXIX Tamil Nadu Register of Advances', () => {
  it('detects Form_XXIX_-_TamilNadu.xlsx context', () => {
    expect(
      isFormXXIXTamilNaduContext(
        { title: 'FORM XXIX', subtitle: 'Register of Advances (Contract Labour)' },
        { state: 'Tamil Nadu', formName: 'Form XXIX' },
        'Form_XXIX_-_TamilNadu.xlsx'
      )
    ).toBe(true);
    // Filename alone is enough (item.state may be blank during Autofill).
    expect(
      isFormXXIXTamilNaduContext({}, null, 'Form_XXIX_-_TamilNadu.xlsx', [
        'Name of the Workman',
        'Employee Number',
        'Amount Paid',
        'Signature or Thumb Impression of the Workman',
      ])
    ).toBe(true);
    expect(
      isFormXXIXTamilNaduContext(
        { title: 'FORM XII' },
        { state: 'Andhra Pradesh' },
        'Form_XII_-_Andhra_Pradesh.xlsx'
      )
    ).toBe(false);
    // Form XXII TN advances register must not be rewritten as Form XXIX.
    expect(
      isFormXXIXTamilNaduContext(
        {
          title: 'FORM XXII',
          subtitle: 'Register of Advances',
          reference:
            '[See Rule 78 (1) (d) of the Tamil Nadu Contract Labour (Regulation and Abolition) Rules, 1975]',
        },
        { state: 'Tamil Nadu', formName: 'Form XXII' },
        'Form_XXII_-_TamilNadu.xlsx',
        [
          'Name of the Workman',
          'Employee Number',
          'Amount Paid',
          'Total amount of deduction imposed',
          'Number of installments to be recovered',
          'Signature or Thumb Impression of the Workman',
        ],
        'Register of Advances Deduction for Damage/Loss'
      )
    ).toBe(false);
  });

  it('does not treat Form_XVIII_-_TamilNadu.xlsx wages-cum-muster as Form XXIX', () => {
    const formXviiiHeaders = [
      'Name of employee',
      'Designation/nature of work',
      'Daily attendance /units worked',
      'Total attendance/ units of work done',
      'Daily rate of wages/piece-rate',
      'Basic wages',
      'Dearness allowance',
      'Overtime',
      'Other cash payments (nature of payment to be indicated)',
      'Total',
      'Deductions, if any (indicate nature)',
      'Net Amount Paid',
      'Signature / Thumb impression of workman',
      'Initial of Contractor or representative',
    ];
    expect(looksLikeFormXXIXTamilNaduTableHeaders(formXviiiHeaders)).toBe(false);
    expect(
      isFormXXIXTamilNaduContext(
        {
          title: 'Form XVIII – Register of Wages-cum-Muster Roll',
          subtitle: 'Form of Register of Wages-cum-Muster Roll',
          reference: '[See rule 78(1)(a)(i)]',
        },
        { state: 'Tamil Nadu', formName: 'Form XVIII' },
        'Form_XVIII_-_TamilNadu.xlsx',
        formXviiiHeaders,
        'Form XVIII Register of Wages-cum-Muster Roll'
      )
    ).toBe(false);
  });

  it('resolves People EmployeeID aliases for Employee Number', () => {
    expect(resolveFormXXIXTamilNaduEmployeeId({ EmployeeID: 'VE0147' })).toBe('VE0147');
    expect(resolveFormXXIXTamilNaduEmployeeId({ Employee_ID: 'VE0099' })).toBe('VE0099');
    expect(resolveFormXXIXTamilNaduEmployeeId({ EmpID: 'E1' })).toBe('E1');
  });

  it('enriches header with aligned Month/Year fields', () => {
    const header = enrichFormXXIXTamilNaduDisplayHeader(
      { title: 'Wrong', fields: [{ label: 'YEAR :', value: '', key: 'year_raw' }] },
      { state: 'Tamil Nadu' },
      'Form_XXIX_-_TamilNadu.xlsx',
      ['Name', 'Father', 'YEAR :'],
      'Register of Advances Rule 78'
    );
    expect(header.title).toBe('FORM XXIX');
    expect(header.subtitle).toMatch(/Register of Advances/i);
    expect(header.reference).toMatch(/Rule 78/);
    expect(isFormXXIXTamilNaduHeaderFieldLayoutFormHeader(header)).toBe(true);
    expect(header.fields.map((f) => f.key)).toEqual(
      expect.arrayContaining(['form_x_month', 'form_x_year', 'form_xxix_contractor'])
    );
  });

  it('strips Month/Year labels from table headers', () => {
    expect(
      stripFormXXIXTamilNaduMonthYearFromTableHeaders([
        'Name',
        'YEAR :',
        'Month:',
        'Designation',
        'Column 15',
      ])
    ).toEqual(['Name', 'Designation', 'Column 15']);
  });

  it('resolves leaf headers to official Form XXIX (1)–(20) model', () => {
    const misaligned = [
      'Serial Number',
      'Name of the Workman',
      "Father/Husband's Name",
      'Employee Number',
      'Designation',
      'Date of Payment',
      'Amount Paid',
      'Number of Installments to be recovered',
      'Date on which recovery completed',
      'Date on which caused Damage or Loss',
      'Date of show cause Notice',
      'Total amount of deduction imposed',
      'Number of installments to be recovered',
      'DEDUCTION FOR DAMAGE AND LOSS (5)',
      'Column 15',
      'Column 16',
      'Column 17',
      'Column 18',
      'Signature or Thumb Impression of the Workman',
      'Remarks',
    ];
    const resolved = resolveFormXXIXTamilNaduTableHeaders(misaligned);
    expect(resolved).toHaveLength(20);
    expect(resolved[13]).toBe('Date on which deduction completed');
    expect(resolved[14]).toBe('Act or Omission');
    expect(resolved[15]).toBe('Date of Show Cause Notice');
    expect(resolved[16]).toBe('Amount of fine imposed');
    expect(resolved[17]).toBe('Date on which fine recovery completed');
    expect(resolved[18]).toMatch(/Signature/);
    expect(resolved[19]).toBe('Remarks');
  });

  it('sanitizes YEAR group labels to Advance/Damage/Fines bands', () => {
    const badGroups = Array.from({ length: 20 }, (_, i) =>
      i >= 14 && i <= 17 ? 'YEAR :' : i >= 5 && i <= 8 ? 'Advance Paid' : ''
    );
    const groups = sanitizeFormXXIXTamilNaduColumnGroupLabels(badGroups, 20);
    expect(groups[5]).toBe(FORM_XXIX_TN_GROUP_ADVANCE);
    expect(groups[9]).toBe(FORM_XXIX_TN_GROUP_DAMAGE);
    expect(groups[14]).toBe(FORM_XXIX_TN_GROUP_FINES);
    expect(groups[15]).toBe(FORM_XXIX_TN_GROUP_FINES);
    expect(groups.every((g) => !/year/i.test(g))).toBe(true);
    expect(FORM_XXIX_TN_TABLE_HEADERS).toHaveLength(20);
  });

  it('ensures Month and Year header fields', () => {
    const header = ensureFormXXIXTamilNaduMonthYearHeaderFields({ fields: [] });
    expect(header.fields).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ key: 'form_x_month', label: 'Month:' }),
        expect.objectContaining({ key: 'form_x_year', label: 'Year:' }),
      ])
    );
  });

  it('autofills contractor fields and month/year from site context', () => {
    const out = applyFormXXIXTamilNaduAutofillFromSite(
      {},
      {
        contractorText: 'Contractor A',
        establishmentText: 'Establishment B',
        natureLocationText: 'Chennai',
        principalEmployerText: 'Employer C',
        monthName: 'April',
        year: 2026,
      }
    );
    expect(out.form_xxix_contractor).toBe('Contractor A');
    expect(out.form_xxix_establishment_contract_carried).toBe('Establishment B');
    expect(out.form_xxix_nature_location_work).toBe('Chennai');
    expect(out.form_xxix_principal_employer).toBe('Employer C');
    expect(out.form_x_month).toBe('April');
    expect(out.form_x_year).toBe('2026');
  });

  it('finalizes header fields in 2-column layout order', () => {
    const fields = finalizeFormXXIXTamilNaduHeaderFields([
      { label: 'Year:', value: '2025', key: 'old_year' },
      { label: '1. Name and Address of Contractor.', value: 'ACME', key: 'c' },
    ]);
    expect(fields[0].key).toBe('form_xxix_contractor');
    expect(fields[0].value).toBe('ACME');
    expect(fields[4].key).toBe('form_x_month');
    expect(fields[5].key).toBe('form_x_year');
    expect(fields[5].value).toBe('2025');
  });

  it('detects Employee Number header and NILL default columns', () => {
    expect(isFormXXIXTamilNaduEmployeeNumberHeader('Employee Number')).toBe(true);
    expect(isFormXXIXTamilNaduNilDefaultHeader('Amount Paid')).toBe(true);
    expect(isFormXXIXTamilNaduNilDefaultHeader('Number of Installments to be recovered')).toBe(true);
    expect(isFormXXIXTamilNaduNilDefaultHeader('Date on which recovery completed')).toBe(true);
    expect(isFormXXIXTamilNaduNilDefaultHeader('Date on which caused Damage or Loss')).toBe(true);
    expect(isFormXXIXTamilNaduNilDefaultHeader('Date of show cause Notice')).toBe(true);
    expect(isFormXXIXTamilNaduNilDefaultHeader('Total amount of deduction imposed')).toBe(true);
    expect(isFormXXIXTamilNaduNilDefaultHeader('Date on which deduction completed')).toBe(true);
    expect(isFormXXIXTamilNaduNilDefaultHeader('Act or Omission')).toBe(true);
    expect(isFormXXIXTamilNaduNilDefaultHeader('Amount of fine imposed')).toBe(true);
    expect(isFormXXIXTamilNaduNilDefaultHeader('Date on which fine recovery completed')).toBe(true);
    expect(
      isFormXXIXTamilNaduNilDefaultHeader('Signature or Thumb Impression of the Workman')
    ).toBe(true);
    expect(isFormXXIXTamilNaduNilDefaultHeader('Name of the Workman')).toBe(false);
    expect(isFormXXIXTamilNaduNilDefaultHeader('Employee Number')).toBe(false);
    expect(isFormXXIXTamilNaduNilDefaultHeader('Date of Payment')).toBe(false);
  });

  it('applies NILL defaults to advance/damage/fines columns', () => {
    const rows = applyFormXXIXTamilNaduNilToMappedRows(
      [
        {
          'Employee Number': 'E1',
          'Amount Paid': '',
          'Act or Omission': 'Enter Act',
          Remarks: '',
        },
      ],
      FORM_XXIX_TN_TABLE_HEADERS,
      FORM_XXIX_TN_NIL
    );
    expect(rows[0]['Amount Paid']).toBe(FORM_XXIX_TN_NIL);
    expect(rows[0]['Act or Omission']).toBe(FORM_XXIX_TN_NIL);
    expect(rows[0]['Signature or Thumb Impression of the Workman']).toBe(FORM_XXIX_TN_NIL);
    expect(rows[0]['Employee Number']).toBe('E1');
    expect(rows[0].Remarks).toBe('');
  });
});
