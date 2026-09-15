import {
  applyFormXVIIITamilNaduAutofillFromSite,
  applyFormXVIIITamilNaduOvertimeNilToMappedRows,
  applyFormXVIIITamilNaduPayrollToExportRow,
  computeFormXVIIITamilNaduDeductions,
  computeFormXVIIITamilNaduOtherCashPayments,
  enrichFormXVIIITamilNaduExportRows,
  enrichFormXVIIITamilNaduPayrollExportRows,
  findFormXVIIITamilNaduRecordByFirstAndLastName,
  FORM_XVIII_TN_OVERTIME_NIL,
  FORM_XVIII_TN_PDF_TABLE_FONT_SIZE,
  FORM_XVIII_TN_AMOUNT_OF_WAGES_EARNED,
  coalesceFormXVIIITamilNaduTitleLines,
  detectFormXVIIITamilNaduPdfGroupBands,
  isFormXVIIITamilNaduWagesCumMusterTitle,
  FORM_XVIII_TN_TRAILING_COL_R0,
  FORM_XVIII_TN_TRAILING_COL_W0,
  stripFormXVIIITamilNaduTrailingExcelJsColumns,
  stripFormXVIIITamilNaduTrailingSheetColumns,
  formXVIIITamilNaduAttendanceNameKeys,
  formXVIIITamilNaduFirstAndLastNamesMatch,
  formXVIIITamilNaduFirstLastKey,
  resolveFormXVIIITamilNaduPaidDays,
  countFormXVIIITamilNaduWorkedDaysFromAttendance,
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
  looksLikeFormXVIIITamilNaduPdfContext,
  resolveFormXVIIITamilNaduDeductionsCellValue,
  resolveFormXVIIITamilNaduPayrollColumnHeaders,
  toFormXVIIITamilNaduPersonNameDisplay,
  isFormXVIIITamilNaduResolvedDeductionsHeader,
} from './formXVIIITamilNaduWagesMuster';
import { statutoryDraftPdfTestUtils } from '../../utils/statutoryDraftPdf';
import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';

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

  it('fills Daily attendance and Total attendance from payroll Paid_days', () => {
    const headers = [
      'Name of employee',
      'Daily attendance /units worked',
      'Total attendance/ units of work done',
    ];
    const row = applyFormXVIIITamilNaduPayrollToExportRow(
      { 'Name of employee': 'Selva P' },
      { first_name: 'Selva', last_name: 'P', paid_days: 26 },
      headers
    );
    expect(row['Daily attendance /units worked']).toBe(26);
    expect(row['Total attendance/ units of work done']).toBe(26);
  });

  it('reads Paid_days from Sample Payroll payload aliases', () => {
    expect(resolveFormXVIIITamilNaduPaidDays({ Paid_days: 22 })).toBe(22);
    expect(resolveFormXVIIITamilNaduPaidDays({ payroll_payload: { paid_days: 18 } })).toBe(18);
    expect(resolveFormXVIIITamilNaduPaidDays({ 'Paid Days': '31' })).toBe(31);
  });

  it('matches Sample Payroll by fullname or firstname + lastname', () => {
    const selvaP = { FirstName: 'Selva', LastName: 'P' };
    expect(formXVIIITamilNaduFirstAndLastNamesMatch(selvaP, { employee_name: 'Selva P' })).toBe(
      true
    );
    expect(
      formXVIIITamilNaduFirstAndLastNamesMatch(
        { Name: 'Pal Pandian V' },
        { first_name: 'Pal', last_name: 'V', employee_name: 'Pal Pandian V' }
      )
    ).toBe(true);
    expect(
      formXVIIITamilNaduFirstAndLastNamesMatch(
        { Name: 'Nagaraju' },
        { employee_name: 'Nagaraju', paid_days: 20 }
      )
    ).toBe(true);
    expect(formXVIIITamilNaduFirstLastKey({ Name: 'Nagaraju' })).toBe('nagaraju');
    expect(formXVIIITamilNaduAttendanceNameKeys(selvaP)).toEqual(
      expect.arrayContaining(['selva p'])
    );
    expect(
      findFormXVIIITamilNaduRecordByFirstAndLastName(
        { Name: 'Dumulai' },
        [
          { employee_name: 'Dumulai', Paid_days: 24 },
          { first_name: 'Selva', last_name: 'P', paid_days: 26 },
        ]
      )?.Paid_days
    ).toBe(24);
  });

  it('matches Form 25 / payroll attendance on FirstName AND LastName only', () => {
    const selvaP = { FirstName: 'Selva', LastName: 'P' };
    const selvaKumar = { FirstName: 'Selva', LastName: 'Kumar' };
    const records = [
      { nameOfTheWorker: 'Selva Kumar', totalDaysWorked: 18, paid_days: 18 },
      { first_name: 'Selva', last_name: 'P', paid_days: 26, nameOfTheWorker: 'Selva P' },
    ];
    expect(formXVIIITamilNaduFirstAndLastNamesMatch(selvaP, records[1])).toBe(true);
    expect(formXVIIITamilNaduFirstAndLastNamesMatch(selvaP, records[0])).toBe(false);
    expect(formXVIIITamilNaduFirstAndLastNamesMatch(selvaP, { first_name: 'Selva' })).toBe(false);
    expect(
      formXVIIITamilNaduFirstAndLastNamesMatch(
        { FirstName: 'Selva P' },
        { first_name: 'Selva', last_name: 'P', paid_days: 26 }
      )
    ).toBe(true);
    expect(
      formXVIIITamilNaduFirstAndLastNamesMatch(
        { FirstName: 'Selva', LastName: 'P', EmployeeID: 'VE0147' },
        { first_name: 'Selva', last_name: 'Kumar', employee_id: 'VE0147', paid_days: 18 }
      )
    ).toBe(false);
    expect(findFormXVIIITamilNaduRecordByFirstAndLastName(selvaP, records)?.paid_days).toBe(26);
    expect(findFormXVIIITamilNaduRecordByFirstAndLastName(selvaKumar, records)?.totalDaysWorked).toBe(
      18
    );
  });

  it('does not treat last initial P as Pal / Kumar P', () => {
    const selvaP = { FirstName: 'Selva', LastName: 'P' };
    expect(formXVIIITamilNaduFirstLastKey(selvaP)).toBe('selva p');
    expect(formXVIIITamilNaduFirstLastKey({ FirstName: 'Selva', LastName: 'Kumar' })).toBe(
      'selva kumar'
    );
    expect(formXVIIITamilNaduFirstLastKey({ nameOfTheWorker: 'Selva Kumar P' })).toBe(
      'selva kumar p'
    );
    expect(formXVIIITamilNaduFirstLastKey({ nameOfTheWorker: 'Selva Pal' })).toBe('selva pal');
    expect(
      formXVIIITamilNaduFirstAndLastNamesMatch(selvaP, { first_name: 'Selva', last_name: 'Pal' })
    ).toBe(false);
    expect(
      formXVIIITamilNaduFirstAndLastNamesMatch(selvaP, { employee_name: 'Selva Kumar P' })
    ).toBe(false);
    expect(
      formXVIIITamilNaduFirstAndLastNamesMatch(selvaP, { employee_name: 'Selva Kumar' })
    ).toBe(false);
    expect(
      findFormXVIIITamilNaduRecordByFirstAndLastName(selvaP, [
        { employee_name: 'Selva Pal', paid_days: 4 },
        { employee_name: 'Selva Kumar P', paid_days: 18 },
      ])
    ).toBeNull();
  });

  it('does not fill June attendance from a May payroll row', () => {
    const selvaP = { FirstName: 'Selva', LastName: 'P' };
    const records = [
      { first_name: 'Selva', last_name: 'P', paid_days: 4, payroll_month: '2026-05' },
    ];
    expect(
      findFormXVIIITamilNaduRecordByFirstAndLastName(selvaP, records, null, {
        monthIso: '2026-06',
      })
    ).toBeNull();
    expect(
      findFormXVIIITamilNaduRecordByFirstAndLastName(selvaP, records, null, {
        monthIso: '2026-05',
      })?.paid_days
    ).toBe(4);
  });

  it('leaves Daily/Total attendance blank when skipAttendance and no June days', () => {
    const headers = [
      'Name of employee',
      'Daily attendance /units worked',
      'Total attendance/ units of work done',
      'Total',
    ];
    const row = applyFormXVIIITamilNaduPayrollToExportRow(
      { 'Name of employee': 'Selva P' },
      { first_name: 'Selva', last_name: 'P', paid_days: 4, gross_pay: 10000 },
      headers,
      { skipAttendance: true }
    );
    expect(row['Daily attendance /units worked']).toBeUndefined();
    expect(row['Total attendance/ units of work done']).toBeUndefined();
    expect(row.Total).toBe(10000);
  });

  it('counts Daily/Total attendance from June punches by first+last only', () => {
    const selvaP = { FirstName: 'Selva', LastName: 'P' };
    const records = [
      { first_name: 'Selva', last_name: 'P', date: '2026-06-02' },
      { first_name: 'Selva', last_name: 'P', date: '2026-06-03' },
      { first_name: 'Selva', last_name: 'Kumar', date: '2026-06-02' },
      { first_name: 'Selva', last_name: 'P', date: '2026-05-20' },
      { employee_name: 'Selva Pal', date: '2026-06-04' },
    ];
    expect(
      countFormXVIIITamilNaduWorkedDaysFromAttendance(selvaP, records, null, {
        monthIso: '2026-06',
      })
    ).toBe(2);
    expect(
      countFormXVIIITamilNaduWorkedDaysFromAttendance(selvaP, records, null, {
        monthIso: '2026-06',
      })
    ).not.toBe(4);
    expect(
      countFormXVIIITamilNaduWorkedDaysFromAttendance(
        { FirstName: 'Selva', LastName: 'Kumar' },
        records,
        null,
        { monthIso: '2026-06' }
      )
    ).toBe(1);
    expect(
      countFormXVIIITamilNaduWorkedDaysFromAttendance(selvaP, [], null, { monthIso: '2026-06' })
    ).toBe(0);
  });

  it('enriches rows from payroll using first+last resolver, not row index', () => {
    const headers = [
      'Name of employee',
      'Daily attendance /units worked',
      'Total attendance/ units of work done',
    ];
    const payroll = [
      { first_name: 'Selva', last_name: 'Kumar', paid_days: 18 },
      { first_name: 'Selva', last_name: 'P', paid_days: 26 },
    ];
    const employees = [
      { FirstName: 'Selva', LastName: 'P' },
      { FirstName: 'Selva', LastName: 'Kumar' },
    ];
    const rows = enrichFormXVIIITamilNaduPayrollExportRows(
      [{ 'Name of employee': 'Selva P' }, { 'Name of employee': 'Selva Kumar' }],
      headers,
      {
        employees,
        resolvePayrollRow: (emp) =>
          findFormXVIIITamilNaduRecordByFirstAndLastName(emp, payroll),
      }
    );
    expect(rows[0]['Daily attendance /units worked']).toBe(26);
    expect(rows[0]['Total attendance/ units of work done']).toBe(26);
    expect(rows[1]['Daily attendance /units worked']).toBe(18);
    expect(rows[1]['Total attendance/ units of work done']).toBe(18);
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

  it('detects Form XVIII wages-cum-muster PDF context and uses table font 10', () => {
    const meta = [
      'Form XVIII – Register of Wages-cum-Muster Roll',
      'Form of Register of Wages-cum-Muster Roll',
      '[See rule 78(1)(a)(i)]',
    ];
    const rows = [
      [
        'Name of employee',
        'Daily attendance /units worked',
        'Total attendance/ units of work done',
        'Daily rate of wages/piece-rate',
      ],
    ];
    expect(
      looksLikeFormXVIIITamilNaduPdfContext(meta, rows, 'Sheet1', 'Form_XVIII_-_TamilNadu.xlsx')
    ).toBe(true);
    expect(FORM_XVIII_TN_PDF_TABLE_FONT_SIZE).toBe(10);
    expect(
      isFormXVIIITamilNaduWagesCumMusterTitle('Form of Register of Wages-cum-Muster Roll')
    ).toBe(true);
    expect(
      coalesceFormXVIIITamilNaduTitleLines([
        'Form of',
        'Register of Wages-cum-Muster Roll',
      ])
    ).toEqual(['Form of Register of Wages-cum-Muster Roll']);
    const { expandStatutoryMetaSegments } = statutoryDraftPdfTestUtils;
    expect(
      expandStatutoryMetaSegments('Form of Register of Wages-cum-Muster Roll')
    ).toEqual(['Form of Register of Wages-cum-Muster Roll']);
    const wageRows = [
      ['', '', '', '', '', '', '', 'Amount of wages earned', '', '', '', '', '', '', '', ''],
      [
        'Sl No.',
        'SL.No.in register of workmen',
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
        'Initial of Contractor or his representative',
      ],
    ];
    const bands = detectFormXVIIITamilNaduPdfGroupBands(wageRows, 0, 1, 16);
    expect(bands[0]?.label).toBe('Amount of wages earned');
    expect(bands[0]?.start).toBe(7);
    expect(bands[0]?.end).toBe(11);
    expect(FORM_XVIII_TN_AMOUNT_OF_WAGES_EARNED).toBe('Amount of wages earned');
    const { buildStatutoryPdfHeaderModel, looksLikeFormWPdfContext } = statutoryDraftPdfTestUtils;
    expect(
      looksLikeFormWPdfContext(meta, rows, 0)
    ).toBe(false);
    const model = buildStatutoryPdfHeaderModel(meta, rows, 0, 'Sheet1', {
      fileName: 'Form_XVIII_-_TamilNadu.xlsx',
    });
    expect(model.titles.some((t) => isFormXVIIITamilNaduWagesCumMusterTitle(t))).toBe(true);
    expect(model.formXVIIITamilNadu).toBe(true);
    expect(model.isFormW).toBe(false);
    expect(
      looksLikeFormXVIIITamilNaduPdfContext(
        ['FORM XXVII', 'REGISTER OF WAGES'],
        [],
        'Sheet1',
        'Form_XXVII_-_TamilNadu.xlsx'
      )
    ).toBe(false);
  });
});

describe('Form XVIII Tamil Nadu Excel layout', () => {
  it('defaults leftover boxes to Excel columns R–W', () => {
    expect(FORM_XVIII_TN_TRAILING_COL_R0).toBe(17);
    expect(FORM_XVIII_TN_TRAILING_COL_W0).toBe(22);
  });

  it('clears SheetJS cells in R–W', () => {
    const ws = {};
    const rRef = XLSX.utils.encode_cell({ r: 12, c: 17 });
    const wRef = XLSX.utils.encode_cell({ r: 12, c: 22 });
    ws[rRef] = { t: 's', v: 'box', s: { border: { top: { style: 'thin' } } } };
    ws[wRef] = { t: 's', v: 'box', s: { border: { top: { style: 'thin' } } } };
    ws['!ref'] = 'A1:W40';
    stripFormXVIIITamilNaduTrailingSheetColumns(ws, {
      writeCell: (ref, value) => {
        ws[ref] = { t: 's', v: value };
      },
      maxRow: 20,
    });
    expect(ws[rRef]).toBeUndefined();
    expect(ws[wRef]).toBeUndefined();
  });

  it('hides ExcelJS R–W boxes on Form XVIII', () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('FORMXVIII');
    ws.getCell(1, 1).value = 'FORM XVIII';
    ws.getCell(2, 1).value = 'Form of Register of Wages-cum-Muster Roll';
    ws.getCell(11, 1).value = 'Name of employee';
    ws.getCell(11, 4).value = 'Daily attendance /units worked';
    ws.getCell(11, 5).value = 'Total attendance/ units of work done';
    const box = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'thin' },
      right: { style: 'thin' },
    };
    ws.getCell(13, 18).value = 'x';
    ws.getCell(13, 18).border = box;
    ws.getCell(13, 23).value = 'x';
    ws.getCell(13, 23).border = box;
    stripFormXVIIITamilNaduTrailingExcelJsColumns(ws);
    expect(ws.getCell(13, 18).value).toBeNull();
    expect(ws.getCell(13, 23).value).toBeNull();
    expect(ws.getColumn(18).hidden).toBe(true);
    expect(ws.getColumn(23).hidden).toBe(true);
  });
});
