import {
  FORM_XIX_TN_TABLE_HEADERS,
  applyFormXIXTamilNaduAllBorders,
  applyFormXIXTamilNaduEmployeeToRow,
  clearFormXIXTamilNaduExtraColumns,
  isFormXIXTamilNaduWageSlipContext,
  looksLikeFormXIXTamilNaduTableHeaders,
  resolveFormXIXTamilNaduTableHeaders,
} from './formXIXTamilNadu';
import ExcelJS from 'exceljs';

describe('formXIXTamilNadu', () => {
  it('detects Form_XIX_-_TamilNadu.xlsx wage slip context', () => {
    expect(
      isFormXIXTamilNaduWageSlipContext(
        { title: 'Form XIX', subtitle: 'Wage Slip' },
        { formFileName: 'Form_XIX_-_TamilNadu.xlsx' },
        'Form_XIX_-_TamilNadu.xlsx',
        'Wage Slip'
      )
    ).toBe(true);
  });

  it('resolves canonical Excel workman columns for autofill', () => {
    expect(resolveFormXIXTamilNaduTableHeaders([])).toEqual(FORM_XIX_TN_TABLE_HEADERS);
    expect(looksLikeFormXIXTamilNaduTableHeaders(FORM_XIX_TN_TABLE_HEADERS)).toBe(true);
  });

  it('maps employee people fields into Excel columns', () => {
    const emp = {
      EmployeeID: 'E101',
      FirstName: 'Ravi',
      LastName: 'Kumar',
      Father_s_Name: 'Suresh',
      UAN: '100200300400',
      ESIC_Number: 'TN1234567',
      DateOfJoining: '2024-01-15',
      Designation: 'Technician',
    };
    const row = applyFormXIXTamilNaduEmployeeToRow({}, emp, FORM_XIX_TN_TABLE_HEADERS, {
      sanitizeValue: (v) => String(v ?? '').trim(),
      formatStatutoryDateDisplay: (v) => String(v || '').trim(),
      payrollRow: {
        paid_days: 26,
        overtime: '',
        basic: 50000,
        dearness_allowance: 2000,
        other_allowance: 3000,
        gross_pay: 55000,
        epf_contribution: 1800,
        professional_tax: 200,
        total_deductions: 2000,
        net_pay: 53000,
        deductions: [
          { type: 'esi', name: 'ESIC', amount: 400 },
          { type: 'lwf', name: 'Labour Welfare Fund', amount: 20 },
        ],
      },
    });
    expect(row['Workman Code']).toBe('E101');
    expect(row['Workman Name']).toBe('Ravi Kumar');
    expect(row["Father's Name"]).toBe('Suresh');
    expect(row.UAN).toBe('100200300400');
    expect(row['ESIC IP Number']).toBe('TN1234567');
    expect(row['Date of Joining']).toBe('2024-01-15');
    expect(row['No. of. Working Days']).toBe('26');
    expect(row['No. of. Overtime']).toBe('NIL');
    expect(row['Rate of Daily Wages/Piece Rate']).toBe('Monthly Wages');
    expect(row['Nature of Work']).toBe('Technician');
    expect(row.Basic).toBe('50000');
    expect(row['Dearness Allowance']).toBe('2000');
    expect(row['Leave with Wages Including Cash in Lieu of Kinds']).toBe('NIL');
    expect(row['Other Allowances']).toBe('3000');
    expect(row['Gross Wages']).toBe('55000');
    expect(row['Employee Provident Fund']).toBe('1800');
    expect(row.ESIC).toBe('400');
    expect(row['Advance/Loan']).toBe('NIL');
    expect(row['Labour Welfare Fund']).toBe('20');
    expect(row['Professional Tax']).toBe('200');
    expect(row['Total Wage Deductions']).toBe('2000');
    expect(row['Net Amount of Wages Paid']).toBe('53000');
  });

  it('computes Total Wage Deductions as gross_pay − net_pay', () => {
    const row = applyFormXIXTamilNaduEmployeeToRow(
      {},
      { FirstName: 'A', LastName: 'B' },
      FORM_XIX_TN_TABLE_HEADERS,
      {
        sanitizeValue: (v) => String(v ?? '').trim(),
        payrollRow: {
          gross_pay: 104065,
          net_pay: 100851,
          total_deductions: 0,
        },
      }
    );
    expect(row['Gross Wages']).toBe('104065');
    expect(row['Net Amount of Wages Paid']).toBe('100851');
    expect(row['Total Wage Deductions']).toBe('3214');
  });

  it('upgrades workman-only headers to full wage-computation columns', () => {
    const headers = resolveFormXIXTamilNaduTableHeaders([
      'Workman Code',
      'Workman Name',
      "Father's Name",
      'UAN',
      'ESIC IP Number',
      'Date of Joining',
      'No. of. Working Days',
      'No. of. Overtime',
      'Rate of Daily Wages/Piece Rate',
      'Nature of Work',
    ]);
    expect(headers).toContain('Basic');
    expect(headers).toContain('Gross Wages');
    expect(headers).toContain('Net Amount of Wages Paid');
    expect(headers.length).toBe(FORM_XIX_TN_TABLE_HEADERS.length);
  });

  it('applies thin all-borders on A1:D17', () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Form XIX');
    applyFormXIXTamilNaduAllBorders(ws);
    const cell = ws.getCell(5, 2);
    expect(cell.border?.top?.style).toBe('thin');
    expect(cell.border?.left?.style).toBe('thin');
    expect(cell.border?.bottom?.style).toBe('thin');
    expect(cell.border?.right?.style).toBe('thin');
    expect(ws.getCell(17, 4).border?.top?.style).toBe('thin');
    expect(ws.getCell(1, 1).border?.right?.style).toBe('thin');
  });

  it('clears stray E–J column data beside the A–D form', () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Form XIX');
    ws.getCell(5, 7).value = '30';
    ws.getCell(7, 7).value = 'Monthly Wages';
    ws.getCell(16, 8).value = '173350';
    ws.getCell(16, 9).value = '22353';
    ws.getCell(5, 2).value = 'keep-me';
    clearFormXIXTamilNaduExtraColumns(ws);
    expect(ws.getCell(5, 7).value).toBeNull();
    expect(ws.getCell(7, 7).value).toBeNull();
    expect(ws.getCell(16, 8).value).toBeNull();
    expect(ws.getCell(16, 9).value).toBeNull();
    expect(ws.getCell(5, 2).value).toBe('keep-me');
  });

  it('writes Net Amount of Wages Paid label then amount on row 17', () => {
    const { writeFormXIXTamilNaduNetAmountRow } = require('./formXIXTamilNadu');
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Form XIX');
    ws.mergeCells('A17:D17');
    ws.getCell(17, 1).value = '158123';
    writeFormXIXTamilNaduNetAmountRow(ws, '158123');
    expect(String(ws.getCell(17, 1).value || '')).toBe('Net Amount of Wages Paid');
    expect(String(ws.getCell(17, 4).value || '')).toBe('158123');
  });
});
