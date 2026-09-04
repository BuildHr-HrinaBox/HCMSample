import {
  applyFormXVIIITamilNaduAutofillFromSite,
  applyFormXVIIITamilNaduOvertimeNilToMappedRows,
  applyFormXVIIITamilNaduPayrollToExportRow,
  computeFormXVIIITamilNaduDeductions,
  computeFormXVIIITamilNaduOtherCashPayments,
  enrichFormXVIIITamilNaduExportRows,
  FORM_XVIII_TN_OVERTIME_NIL,
  formXVIIITamilNaduRowsRicherThan,
  isFormXVIIITamilNaduDailyAttendanceHeader,
  isFormXVIIITamilNaduDailyRateHeader,
  isFormXVIIITamilNaduDeductionsHeader,
  isFormXVIIITamilNaduOtherCashPaymentsHeader,
  isFormXVIIITamilNaduOvertimeHeader,
  isFormXVIIITamilNaduPlaceholderHeaderValue,
  isFormXVIIITamilNaduSerialRegisterHeader,
  isFormXVIIITamilNaduTotalAttendanceHeader,
  looksLikeFormXVIIITamilNaduStaleDeductionCell,
  resolveFormXVIIITamilNaduDeductionsCellValue,
  resolveFormXVIIITamilNaduPayrollColumnHeaders,
  toFormXVIIITamilNaduPersonNameDisplay,
  isFormXVIIITamilNaduResolvedDeductionsHeader,
} from './formXVIIITamilNaduWagesMuster';

describe('Form XVIII Tamil Nadu payroll column headers', () => {
  it('detects SL.No.in register of workmen (sl.no, not only serial)', () => {
    expect(isFormXVIIITamilNaduSerialRegisterHeader('SL.No.in register of workmen')).toBe(true);
    expect(
      isFormXVIIITamilNaduSerialRegisterHeader('Serial Number in the Register of Workmen')
    ).toBe(true);
    expect(isFormXVIIITamilNaduSerialRegisterHeader('Sl No.')).toBe(false);
    expect(isFormXVIIITamilNaduSerialRegisterHeader('Name of employee')).toBe(false);
  });

  it('defaults Overtime to Nil when empty and keeps real OT amounts', () => {
    expect(isFormXVIIITamilNaduOvertimeHeader('Overtime')).toBe(true);
    expect(isFormXVIIITamilNaduOvertimeHeader('Other cash payments')).toBe(false);
    const headers = ['Name of employee', 'Overtime', 'Basic wages'];
    const rows = applyFormXVIIITamilNaduOvertimeNilToMappedRows(
      [
        { 'Name of employee': 'rajesh', Overtime: '', 'Basic wages': 100 },
        { 'Name of employee': 'raja', Overtime: 2500, 'Basic wages': 200 },
      ],
      headers,
      FORM_XVIII_TN_OVERTIME_NIL,
      { overwrite: false }
    );
    expect(rows[0].Overtime).toBe('Nil');
    expect(rows[1].Overtime).toBe(2500);
  });

  it('distinguishes Daily attendance from Total attendance', () => {
    expect(
      isFormXVIIITamilNaduDailyAttendanceHeader('Daily attendance /units worked')
    ).toBe(true);
    expect(
      isFormXVIIITamilNaduTotalAttendanceHeader('Daily attendance /units worked')
    ).toBe(false);
    expect(
      isFormXVIIITamilNaduTotalAttendanceHeader('Total attendance/ units of work done')
    ).toBe(true);
    expect(
      isFormXVIIITamilNaduDailyAttendanceHeader('Total attendance/ units of work done')
    ).toBe(false);
  });

  it('detects daily rate and other cash payment headers', () => {
    expect(
      isFormXVIIITamilNaduDailyRateHeader('Daily rate of wages/piece-rate')
    ).toBe(true);
    expect(
      isFormXVIIITamilNaduOtherCashPaymentsHeader(
        'Other cash payments (nature of payment to be indicated)'
      )
    ).toBe(true);
  });

  it('computes Other cash payments as gross_pay − basic − hra', () => {
    expect(computeFormXVIIITamilNaduOtherCashPayments(100000, 40000, 20000)).toBe(40000);
    expect(computeFormXVIIITamilNaduOtherCashPayments('113708', '50000', '20000')).toBe(43708);
    expect(computeFormXVIIITamilNaduOtherCashPayments(70326, '', '')).toBe('');
    expect(computeFormXVIIITamilNaduOtherCashPayments(74132, 74132, 0)).toBe(0);
  });

  it('detects Deductions, if any (indicate nature) header', () => {
    expect(
      isFormXVIIITamilNaduDeductionsHeader('Deductions, if any (indicate nature)')
    ).toBe(true);
    expect(isFormXVIIITamilNaduDeductionsHeader('Net Amount Paid')).toBe(false);
  });

  it('computes Deductions as gross_pay − net_pay (Total − Net Amount Paid)', () => {
    expect(computeFormXVIIITamilNaduDeductions(129102, 113888)).toBe(15214);
    expect(computeFormXVIIITamilNaduDeductions('74162', '70486')).toBe(3676);
  });

  it('title-cases employee names for export', () => {
    expect(toFormXVIIITamilNaduPersonNameDisplay('selva p')).toBe('Selva P');
    expect(toFormXVIIITamilNaduPersonNameDisplay('pal pandian v')).toBe('Pal Pandian V');
  });

  it('overwrites template contractor placeholder from site autofill', () => {
    expect(isFormXVIIITamilNaduPlaceholderHeaderValue('Name and address of contractor')).toBe(true);
    const out = applyFormXVIIITamilNaduAutofillFromSite(
      { form_xviii_contractor: 'Name and address of contractor' },
      { contractorText: 'Acme Contractors, Chennai' },
      { overwrite: true, formHeaderFields: [{ key: 'custom_contractor', label: 'Name and Address of Contractor.' }] }
    );
    expect(out.form_xviii_contractor).toBe('Acme Contractors, Chennai');
    expect(out.custom_contractor).toBe('Acme Contractors, Chennai');
  });

  it('applies payroll gross_pay and net_pay to export row and deductions', () => {
    const headers = [
      'Name of employee',
      'Total',
      'Deductions, if any (indicate nature)',
      'Net Amount Paid',
    ];
    const row = applyFormXVIIITamilNaduPayrollToExportRow(
      { 'Name of employee': 'selva p' },
      { gross_pay: 129102, net_pay: 113888, paid_days: 31 },
      headers
    );
    expect(row.Total).toBe(129102);
    expect(row['Net Amount Paid']).toBe(113888);
    expect(row['Deductions, if any (indicate nature)']).toBe(15214);
  });

  it('enriches export rows with title-case names and computed deductions', () => {
    const headers = [
      'Name of employee',
      'Total',
      'Deductions, if any (indicate nature)',
      'Net Amount Paid',
    ];
    const rows = enrichFormXVIIITamilNaduExportRows(
      [
        {
          'Name of employee': 'selva p',
          Total: 129102,
          'Deductions, if any (indicate nature)': '07:45',
          'Net Amount Paid': 113888,
        },
      ],
      headers
    );
    expect(rows[0]['Name of employee']).toBe('Selva P');
    expect(rows[0]['Deductions, if any (indicate nature)']).toBe(15214);
  });

  it('treats template time text 07:45 as stale deductions placeholder', () => {
    expect(looksLikeFormXVIIITamilNaduStaleDeductionCell('07:45')).toBe(true);
    expect(looksLikeFormXVIIITamilNaduStaleDeductionCell(3214)).toBe(false);
  });

  it('resolves modal deductions display from Total − Net when stale', () => {
    const headers = ['Total', '13', 'Net Amount Paid'];
    expect(
      resolveFormXVIIITamilNaduDeductionsCellValue(
        {
          Total: 70562,
          '13': '07:45',
          'Net Amount Paid': 67348,
        },
        headers
      )
    ).toBe(3214);
  });

  it('resolves numeric column key "13" as Deductions when 12/14 band present', () => {
    const headers = ['Total', '13', 'Net Amount Paid'];
    const cols = resolveFormXVIIITamilNaduPayrollColumnHeaders(headers);
    expect(cols.totalAmount).toBe('Total');
    expect(cols.deductions).toBe('13');
    expect(cols.netAmountPaid).toBe('Net Amount Paid');
    expect(isFormXVIIITamilNaduResolvedDeductionsHeader('13', headers)).toBe(true);
  });

  it('enriches rows when deductions column key is "13"', () => {
    const headers = ['Total', '13', 'Net Amount Paid'];
    const rows = enrichFormXVIIITamilNaduExportRows(
      [{ Total: 70562, '13': '07:45', 'Net Amount Paid': 67348 }],
      headers
    );
    expect(rows[0]['13']).toBe(3214);
  });

  it('scores autofill rows richer than sparse name/basic SampleData', () => {
    const headers = [
      'Name of employee',
      'Designation/nature of work',
      'Daily attendance /units worked',
      'Total attendance/ units of work done',
      'Daily rate of wages/piece-rate',
      'Basic wages',
    ];
    const rich = [
      {
        'Name of employee': 'rajesh',
        'Designation/nature of work': 'Engineer',
        'Daily attendance /units worked': 31,
        'Total attendance/ units of work done': 31,
        'Daily rate of wages/piece-rate': 70562,
        'Basic wages': 26781,
      },
    ];
    const sparse = [{ 'Name of employee': 'rajesh', 'Basic wages': 26781 }];
    expect(formXVIIITamilNaduRowsRicherThan(rich, headers, sparse, headers)).toBe(true);
    expect(formXVIIITamilNaduRowsRicherThan(sparse, headers, rich, headers)).toBe(false);
  });
});
