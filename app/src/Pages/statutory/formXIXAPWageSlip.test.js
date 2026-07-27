import * as XLSX from 'xlsx';
import {
  FORM_XIX_AP_FIELD_GROUPS,
  FORM_XIX_AP_WAGE_TABLE_HEADERS,
  buildFormXIXAPTemplateFields,
  isFormXIXAPTableLayoutFormHeader,
  isFormXIXAPTableWageSlipContext,
  isFormXIXAPWageSlipContext,
  mergeFormXIXAPWageTableRowIntoHeaderData,
  resolveFormXIXAPHeaderFieldLayout,
  resolveFormXIXAPWageTableHeaders,
  resolveFormXIXWorkbookSheetName,
} from './formXIXAPWageSlip';
import {
  applyFormXIXMPEmployeeToRow,
  resolveFormXIXMPPayrollFields,
  resolveFormXIXMPPayrollHelpers,
} from './formXIXMPWageSlip';

describe('Form XIX AP wage slip detection', () => {
  it('detects Andhra Pradesh Form XIX wage slip from file and form name', () => {
    expect(
      isFormXIXAPWageSlipContext(
        { title: 'Form XIX – Wage Slip', subtitle: 'Wage Slip (Contract Labour)' },
        { state: 'Andhra Pradesh', formName: 'Form XIX – Wage Slip' },
        'Form_XIX_-_Andhra_Pradesh.xlsx',
        'FORM XIX WAGE SLIP Rule 78(1)(b)'
      )
    ).toBe(true);
  });

  it('treats AP Form XIX as tabular wage particulars layout', () => {
    expect(
      isFormXIXAPTableWageSlipContext(
        { title: 'Form XIX – Wage Slip', formXIXAPHeaderFieldLayout: true },
        { state: 'Andhra Pradesh', formName: 'Form XIX – Wage Slip' },
        'Form_XIX_-_Andhra_Pradesh.xlsx',
        'FORM XIX WAGE SLIP'
      )
    ).toBe(true);
  });

  it('does not treat Madhya Pradesh Form XIX as AP table layout', () => {
    expect(
      isFormXIXAPTableWageSlipContext(
        { title: 'Form XIX – Wage Slip', formXIXMPTableLayout: true },
        { state: 'Madhya Pradesh', formName: 'Form XIX – Wage Slip' },
        'Form_XIX_MP.xlsx',
        'FORM XIX WAGE SLIP'
      )
    ).toBe(false);
  });
});

describe('Form XIX AP wage table layout', () => {
  it('exposes official wage particulars 1–7 as table headers', () => {
    const headers = resolveFormXIXAPWageTableHeaders([]);
    expect(headers).toEqual(FORM_XIX_AP_WAGE_TABLE_HEADERS);
    expect(headers).toContain('1. No. of days worked');
    expect(headers).toContain('7. Net amount of wages paid');
    expect(FORM_XIX_AP_FIELD_GROUPS.map((g) => g.id)).toEqual(['header', 'footer']);
  });

  it('builds header fields without wage particulars group', () => {
    const fields = buildFormXIXAPTemplateFields().filter((f) => f.group !== 'wages');
    const labels = fields.map((f) => f.label);
    expect(labels).toContain('Name and address of contractor:');
    expect(labels).toContain('Nature and location of work:');
    expect(labels.some((l) => /days worked/i.test(l))).toBe(false);
  });

  it('resolves AP layout with wage table headers and formXIXAPTableLayout', () => {
    const wb = XLSX.utils.book_new();
    const rows = [
      ['FORM XIX', '', '', 'WAGE SLIP', '', '[Rule 78(1)(b)]'],
      ['Name and address of contractor:', ''],
      ['Nature and location of work:', ''],
      ['', '', "Name and Father's/Husband's Name of the workman:", ''],
      ['', '', 'For the week/Fortnight/Month ending:', ''],
      ['1. No. of days worked', ''],
      ['2. No. of units worked in case of piece-rate Workers', ''],
      ['3. Rate of daily wages/piece-rate', ''],
      ['4. Amount of overtime wages', ''],
      ['5. Gross wages payable', ''],
      ['6. Deductions, if any', ''],
      ['7. Net amount of wages paid', ''],
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), 'XIX-Wage Slip');

    const sheetName = resolveFormXIXWorkbookSheetName(wb, {
      fileName: 'Form_XIX_-_Andhra_Pradesh.xlsx',
      formFileName: 'Form_XIX_-_Andhra_Pradesh.xlsx',
    });
    expect(sheetName).toMatch(/wage\s+slip/i);

    const layout = resolveFormXIXAPHeaderFieldLayout(
      { formHeader: { title: 'Form XIX – Wage Slip' }, headers: [], tableData: [] },
      wb,
      {
        item: { state: 'Andhra Pradesh', formName: 'Form XIX – Wage Slip' },
        fileName: 'Form_XIX_-_Andhra_Pradesh.xlsx',
        formFileName: 'Form_XIX_-_Andhra_Pradesh.xlsx',
        sheetText: 'FORM XIX WAGE SLIP Rule 78(1)(b)',
      }
    );

    expect(layout?.formHeader?.formXIXAPHeaderFieldLayout).toBe(true);
    expect(isFormXIXAPTableLayoutFormHeader(layout?.formHeader)).toBe(true);
    expect(layout?.headers).toEqual(FORM_XIX_AP_WAGE_TABLE_HEADERS);
    expect(layout?.formHeader?.fields?.every((f) => f.group !== 'wages')).toBe(true);
  });

  it('merges wage table row into header form keys for export', () => {
    const row = {
      "Name and Father's/Husband's Name of the workman": 'Ravi\nFather',
      '1. No. of days worked': '26',
      '2. No. of units worked in case of piece-rate Workers': '',
      '3. Rate of daily wages/piece-rate': '4646.31',
      '5. Gross wages payable': '15000',
      '7. Net amount of wages paid': '14200',
    };
    const merged = mergeFormXIXAPWageTableRowIntoHeaderData(
      { form_xix_ap_contractor: 'ABC Ltd' },
      row
    );
    expect(merged.form_xix_ap_contractor).toBe('ABC Ltd');
    expect(merged.form_xix_ap_workman).toBe('Ravi\nFather');
    expect(merged.form_xix_ap_days_worked).toBe('26');
    expect(merged.form_xix_ap_rate).toBe('4646.31');
    expect(merged.form_xix_ap_gross).toBe('15000');
    expect(merged.form_xix_ap_net).toBe('14200');
  });
});

describe('Form XIX AP payroll autofill rules', () => {
  const apHelpers = resolveFormXIXMPPayrollHelpers({
    formHeader: { title: 'Form XIX – Wage Slip', formXIXAPTableLayout: true },
    item: { state: 'Andhra Pradesh', formName: 'Form XIX – Wage Slip' },
    fileName: 'Form_XIX_-_Andhra_Pradesh.xlsx',
    sheetText: 'FORM XIX WAGE SLIP Rule 78(1)(b)',
  });

  it('enables Andhra Pradesh payroll rules for AP Form XIX', () => {
    expect(apHelpers.andhraPradeshPayrollRules).toBe(true);
  });

  it('sets overtime to NIL and deductions to gross_pay − net_pay', () => {
    const fields = resolveFormXIXMPPayrollFields(
      { gross_pay: '15000', net_pay: '14200', overtime: '500', total_deductions: '50' },
      apHelpers
    );
    expect(fields.overtimeWages).toBe('NIL');
    expect(Number(fields.deductions)).toBe(800);
    expect(String(fields.grossWages)).toBe('15000');
    expect(String(fields.netWages)).toBe('14200');
  });

  it('fills overtime NIL on employee row even without payroll', () => {
    const row = applyFormXIXMPEmployeeToRow(
      {},
      { FirstName: 'Ravi', LastName: 'Kumar' },
      FORM_XIX_AP_WAGE_TABLE_HEADERS,
      { ...apHelpers, payrollRow: null }
    );
    expect(row['4. Amount of overtime wages']).toBe('NIL');
  });

  it('keeps AP overtime NIL and gross−net deductions after payroll enrich overwrite', () => {
    const { enrichFormXIXMPPayrollRows, applyFormXIXAPWageParticularsToRow } = require('./formXIXMPWageSlip');
    const mapped = [
      {
        "Name and Father's/Husband's Name of the workman": 'Ravi',
        '4. Amount of overtime wages': '',
        '5. Gross wages payable': '120804',
        '6. Deductions, if any': '0',
        '7. Net amount of wages paid': '109883',
      },
    ];
    enrichFormXIXMPPayrollRows(
      mapped,
      [{ FirstName: 'Ravi' }],
      FORM_XIX_AP_WAGE_TABLE_HEADERS,
      {
        andhraPradeshPayrollRules: true,
        overwrite: true,
        resolvePayrollRow: () => ({
          gross_pay: 120804,
          net_pay: 109883,
          overtime: '',
          total_deductions: 0,
        }),
      }
    );
    expect(mapped[0]['4. Amount of overtime wages']).toBe('NIL');
    expect(String(mapped[0]['6. Deductions, if any'])).toBe('10921');
    applyFormXIXAPWageParticularsToRow(mapped[0], FORM_XIX_AP_WAGE_TABLE_HEADERS);
    expect(mapped[0]['4. Amount of overtime wages']).toBe('NIL');
    expect(String(mapped[0]['6. Deductions, if any'])).toBe('10921');
  });

  it('builds a ZIP with one wage-slip workbook per employee', async () => {
    const XLSX = require('xlsx');
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(
      wb,
      XLSX.utils.aoa_to_sheet([
        ['FORM XIX', 'WAGE SLIP'],
        ['Name and address of contractor:', ''],
        ['1. No. of days worked', ''],
        ['4. Amount of overtime wages', ''],
        ['5. Gross wages payable', ''],
        ['6. Deductions, if any', ''],
        ['7. Net amount of wages paid', ''],
      ]),
      'XIX-Wage Slip'
    );
    const templateArrayBuffer = XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
    const { buildFormXIXAPPerEmployeeDownload } = require('./formXIXAPWageSlip');
    const result = await buildFormXIXAPPerEmployeeDownload({
      templateArrayBuffer,
      mappedData: [
        {
          "Name and Father's/Husband's Name of the workman": 'Ravi Kumar',
          '5. Gross wages payable': '120804',
          '7. Net amount of wages paid': '109883',
        },
        {
          "Name and Father's/Husband's Name of the workman": 'Anita Devi',
          '5. Gross wages payable': '98189',
          '7. Net amount of wages paid': '93953',
        },
      ],
      headersToUse: FORM_XIX_AP_WAGE_TABLE_HEADERS,
      parsedFormHeader: { title: 'Form XIX – Wage Slip', formXIXAPTableLayout: true },
      formFileName: 'Form_XIX_-_Andhra_Pradesh.xlsx',
      headerFormData: { form_xix_ap_contractor: 'ABC Ltd' },
    });
    expect(result.fileName).toMatch(/\.zip$/i);
    expect(result.blob).toBeTruthy();
    expect(result.blob.size).toBeGreaterThan(100);
  });

  it('writes wage particulars onto dotted H-column cells even when modal fields omit wages', async () => {
    const ExcelJS = require('exceljs');
    const {
      writeFormXIXAPFieldsToExcelJsWorksheet,
      applyFormXIXAPExportWageRulesToHeaderData,
      mergeFormXIXAPWageTableRowIntoHeaderData,
    } = require('./formXIXAPWageSlip');
    const workbook = new ExcelJS.Workbook();
    const ws = workbook.addWorksheet('XIX-Wage Slip');
    ws.getCell(1, 1).value = 'FORM XIX';
    ws.getCell(5, 1).value = 'Name and address of contractor:';
    ws.getCell(5, 8).value = '……';
    ws.getCell(12, 1).value = '1. No. of days worked';
    ws.getCell(12, 8).value = '……';
    ws.getCell(15, 1).value = '3. Rate of daily wages/piece-rate';
    ws.getCell(15, 8).value = '……';
    ws.getCell(16, 1).value = '4. Amount of overtime wages';
    ws.getCell(16, 8).value = '……';
    ws.getCell(17, 1).value = '5. Gross wages payable';
    ws.getCell(17, 8).value = '……';
    ws.getCell(18, 1).value = '6. Deductions, if any';
    ws.getCell(18, 8).value = '……';
    ws.getCell(19, 1).value = '7. Net amount of wages paid';
    ws.getCell(19, 8).value = '……';

    const merged = applyFormXIXAPExportWageRulesToHeaderData(
      mergeFormXIXAPWageTableRowIntoHeaderData(
        { form_xix_ap_contractor: 'VAYONA ENERGY PRIVATE LIMITED' },
        {
          '1. No. of days worked': '26',
          '3. Rate of daily wages/piece-rate': '4646.31',
          '5. Gross wages payable': '120804',
          '7. Net amount of wages paid': '109883',
        },
        FORM_XIX_AP_WAGE_TABLE_HEADERS
      )
    );
    // Simulate modal: only header fields present (wages live in table only).
    writeFormXIXAPFieldsToExcelJsWorksheet(ws, merged, {
      formXIXAPTableLayout: true,
      fields: [
        {
          key: 'form_xix_ap_contractor',
          label: 'Name and address of contractor:',
          group: 'header',
          labelRow: 4,
          valueCol: 7,
        },
      ],
    });

    expect(String(ws.getCell(5, 8).value)).toBe('VAYONA ENERGY PRIVATE LIMITED');
    expect(String(ws.getCell(12, 8).value)).toBe('26');
    expect(String(ws.getCell(15, 8).value)).toBe('4646.31');
    expect(String(ws.getCell(16, 8).value)).toBe('NIL');
    expect(String(ws.getCell(17, 8).value)).toBe('120804');
    expect(String(ws.getCell(18, 8).value)).toBe('10921');
    expect(String(ws.getCell(19, 8).value)).toBe('109883');
    // Must not dump wage values into empty column B next to the label.
    expect(String(ws.getCell(12, 2).value || '')).toBe('');
  });

  it('writes Gujarat Form XIX contractor name/address to stacked column E (not H)', async () => {
    const ExcelJS = require('exceljs');
    const { writeFormXIXAPFieldsToExcelJsWorksheet } = require('./formXIXAPWageSlip');
    const { FORM_XIX_MP_STACKED_VALUE_COL } = require('./formXIXMPWageSlip');

    const workbook = new ExcelJS.Workbook();
    const ws = workbook.addWorksheet('Sheet1');
    ws.getCell(1, 1).value = 'FORM XIX';
    ws.getCell(2, 1).value = '[See rule 78 (2)(b)]';
    ws.getCell(3, 1).value = 'Wage Slip';
    ws.getCell(7, 1).value = 'Name and address if contractor………………..';

    // No parsed valueCol — scan path previously preferred AP H-band; stacked GJ must use E.
    writeFormXIXAPFieldsToExcelJsWorksheet(
      ws,
      { form_xix_ap_contractor: 'Sample, vvd/14' },
      {
        title: 'FORM XIX',
        formXIXMPTableLayout: true,
        fields: [
          {
            key: 'form_xix_ap_contractor',
            label: 'Name and address if contractor………………..',
            group: 'header',
          },
        ],
      }
    );

    expect(String(ws.getCell(7, FORM_XIX_MP_STACKED_VALUE_COL).value)).toBe('Sample, vvd/14');
    expect(String(ws.getCell(7, 8).value || '')).toBe('');
  });
});
