import {
  applyFormPGJGujaratApprovedLeaveToDayColumns,
  applyFormPGJGujaratPayrollToRow,
  applyFormPGJGujaratStaticDefaults,
  applyFormPGJGujaratHeaderAutofill,
  buildFormPGJGujaratWorkbookWithTemplateStyles,
  computeFormPGJGujaratOtherDeductions,
  computeFormPGJGujaratTotalDeduction,
  countFormPGJGujaratDaysWorkedFromAttendance,
  FORM_PGJ_DATE_OF_MONTH_GROUP_LABEL,
  FORM_PGJ_TABLE_FONT_SIZE,
  FORM_PGJ_WORKING_HOURS_FROM,
  FORM_PGJ_WORKING_HOURS_TO,
  formPGJEmployerDisplayName,
  formPGJLeaveCodeFromType,
  formatFormPGJPayrollPayDate,
  resolveFormPGJGujaratDaysInMonth,
  resolveFormPGJGujaratPayrollFields,
  resolveFormPGJPayDateFromCache,
  writeFormPGJGujaratHeaderMetaToExcelJsWorksheet,
} from './formPGJGujarat';
import { excelJSCellHasFullBoxBorder } from '../../utils/excelTableBorders';

describe('Form P Gujarat Sample Payroll deductions', () => {
  test('Total Deduction Rs. = gross_pay − net_pay', () => {
    expect(computeFormPGJGujaratTotalDeduction(236887, 177014)).toBe(59873);
    expect(computeFormPGJGujaratTotalDeduction('95316', '89026')).toBe(6290);
    expect(computeFormPGJGujaratTotalDeduction('', 100)).toBe('');
  });

  test('Other Deductions Rs. = Sample Payroll Total Deduction − Professional Tax', () => {
    expect(computeFormPGJGujaratOtherDeductions(1600, 200)).toBe(1400);
    expect(computeFormPGJGujaratOtherDeductions(700, '')).toBe(700);
    expect(computeFormPGJGujaratOtherDeductions('', 200)).toBe('');
  });

  test('resolveFormPGJGujaratPayrollFields reads PF / PT / Income Tax from Sample Payroll', () => {
    const fields = resolveFormPGJGujaratPayrollFields({
      employee_name: 'Test Emp',
      pf: 9211,
      professional_tax: 200,
      income_tax: 30640,
      gross_pay: 236887,
      net_pay: 177014,
      total_deductions: 1600,
    });
    expect(fields.pf).toBe(9211);
    expect(fields.professionalTax).toBe(200);
    expect(fields.incomeTax).toBe(30640);
    expect(fields.totalDeduction).toBe(59873);
    expect(fields.otherDeductions).toBe(1400);
  });

  test('resolveFormPGJGujaratPayrollFields prefers flatten epf_contribution / income_tax', () => {
    const fields = resolveFormPGJGujaratPayrollFields({
      employee_name: 'Test Emp',
      epf_contribution: 9211,
      income_tax: 30640,
      professional_tax: 200,
      Gross: 22832,
      'Net Pay': 20000,
      'Total Deduction': 700,
    });
    expect(fields.pf).toBe(9211);
    expect(fields.incomeTax).toBe(30640);
    expect(fields.grossPay).toBe(22832);
    expect(fields.totalDeduction).toBe(2832);
    expect(fields.otherDeductions).toBe(500);
  });

  test('applyFormPGJGujaratPayrollToRow fills PF, PT, Income Tax, Other / Total Deduction', () => {
    const headers = [
      'Provident Fund Rs.',
      'Professional Tax Rs.',
      'Income Tax Rs.',
      'Other Deductions Rs.',
      'Total Deduction Rs.',
    ];
    const row = {};
    applyFormPGJGujaratPayrollToRow(
      row,
      {
        pf: 1800,
        professional_tax: 200,
        income_tax: 1200,
        gross_pay: 50000,
        net_pay: 45000,
        total_deductions: 1600,
      },
      headers
    );
    expect(row['Provident Fund Rs.']).toBe('1800');
    expect(row['Professional Tax Rs.']).toBe('200');
    expect(row['Income Tax Rs.']).toBe('1200');
    expect(row['Other Deductions Rs.']).toBe('1400');
    expect(row['Total Deduction Rs.']).toBe('5000');
  });

  test('PF column does not receive Professional Tax', () => {
    const row = {};
    applyFormPGJGujaratPayrollToRow(
      row,
      { pf: 1800, professional_tax: 200 },
      ['Provident Fund Rs.']
    );
    expect(row['Provident Fund Rs.']).toBe('1800');
    expect(row['Provident Fund Rs.']).not.toBe('200');
  });

  test('Working hours From / To default to 9 AM / 5 PM', () => {
    expect(FORM_PGJ_WORKING_HOURS_FROM).toBe('9 AM');
    expect(FORM_PGJ_WORKING_HOURS_TO).toBe('5 PM');

    const row = {};
    applyFormPGJGujaratPayrollToRow(row, { pf: 100 }, [
      'Working hours_From',
      'Working hours_To',
    ]);
    expect(row['Working hours_From']).toBe('9 AM');
    expect(row['Working hours_To']).toBe('5 PM');

    const staticRow = {};
    applyFormPGJGujaratStaticDefaults([staticRow], [
      'Working hours_From',
      'Working hours_To',
    ]);
    expect(staticRow['Working hours_From']).toBe('9 AM');
    expect(staticRow['Working hours_To']).toBe('5 PM');
  });

  test('Date of Payment uses month end when provided', () => {
    const fields = resolveFormPGJGujaratPayrollFields(
      { employee_name: 'A', pf: 100, pay_date: '2026-04-15' },
      { monthEndDate: '30-04-2026', payDate: '2026-04-15' }
    );
    expect(fields.paymentDate).toBe('30-04-2026');

    const row = {};
    applyFormPGJGujaratPayrollToRow(
      row,
      { pf: 100, pay_date: '2026-04-15' },
      ['Date of Payment'],
      { monthEndDate: '30-04-2026' }
    );
    expect(row['Date of Payment']).toBe('30-04-2026');
  });

  test('Date of Payment fills from Sample Payroll Pay date', () => {
    expect(formatFormPGJPayrollPayDate('2026-04-30')).toBe('30-04-2026');

    const fields = resolveFormPGJGujaratPayrollFields(
      { employee_name: 'A', pf: 100 },
      { payDate: '2026-04-30' }
    );
    expect(fields.paymentDate).toBe('30-04-2026');

    const row = {};
    applyFormPGJGujaratPayrollToRow(
      row,
      { pf: 100, pay_date: '2026-04-30' },
      ['Date of Payment']
    );
    expect(row['Date of Payment']).toBe('30-04-2026');

    const metaOnly = {};
    applyFormPGJGujaratPayrollToRow(metaOnly, null, ['Date of Payment'], {
      payDate: '2026-04-30',
    });
    expect(metaOnly['Date of Payment']).toBe('30-04-2026');
  });

  test('empty row pay_date still uses helpers Pay date', () => {
    const fields = resolveFormPGJGujaratPayrollFields(
      { employee_name: 'A', pf: 100, pay_date: '', payDate: '' },
      { payDate: '2026-04-30' }
    );
    expect(fields.paymentDate).toBe('30-04-2026');

    const row = { 'Date of Payment': '' };
    applyFormPGJGujaratPayrollToRow(
      row,
      { pf: 100, net_pay: 90000, gross_pay: 95000, pay_date: '' },
      ['Date of Payment', 'Net Payable Rs.'],
      { payDate: '2026-04-15' }
    );
    expect(row['Date of Payment']).toBe('15-04-2026');
  });

  test('formPGJTemplateBucket maps DEDUCTION child labels', () => {
    // Exercised via apply + export value helpers — headers under DEDUCTION parent.
    const row = {};
    applyFormPGJGujaratPayrollToRow(
      row,
      {
        pf: 1800,
        professional_tax: 200,
        income_tax: 1200,
        gross_pay: 50000,
        net_pay: 45000,
        total_deductions: 1600,
      },
      [
        'Provident Fund Rs.',
        'Professional Tax Rs.',
        'Income Tax Rs.',
        'Other Deductions Rs.',
        'Total Deduction Rs.',
        'Date of Payment',
      ],
      { monthEndDate: '30-04-2026' }
    );
    expect(row['Provident Fund Rs.']).toBe('1800');
    expect(row['Professional Tax Rs.']).toBe('200');
    expect(row['Income Tax Rs.']).toBe('1200');
    expect(row['Other Deductions Rs.']).toBe('1400');
    expect(row['Total Deduction Rs.']).toBe('5000');
    expect(row['Date of Payment']).toBe('30-04-2026');
  });

  test('resolveFormPGJPayDateFromCache reads top-level payDate', () => {
    expect(
      resolveFormPGJPayDateFromCache({
        payDate: '2026-04-30',
        meta: {},
      })
    ).toBe('2026-04-30');
    expect(
      resolveFormPGJPayDateFromCache({
        payDate: '',
        meta: { pay_date: '2026-05-01' },
      })
    ).toBe('2026-05-01');
  });
});

describe('Form P GJ header employer / month + Excel borders', () => {
  test('formPGJEmployerDisplayName keeps company name before address', () => {
    expect(
      formPGJEmployerDisplayName('VAYONA ENERGY PRIVATE LIMITED, 274 A, Theni')
    ).toBe('VAYONA ENERGY PRIVATE LIMITED');
    expect(formPGJEmployerDisplayName('Vayona Energy')).toBe('Vayona Energy');
  });

  test('applyFormPGJGujaratHeaderAutofill sets employer and month keys', () => {
    const out = applyFormPGJGujaratHeaderAutofill(
      {},
      {
        monthText: 'December',
        employerText: 'VAYONA ENERGY PRIVATE LIMITED, Theni',
      }
    );
    expect(out.form_p_gj_month).toBe('December');
    expect(out.form_x_month).toBe('December');
    expect(out.form_p_gj_employer).toBe('VAYONA ENERGY PRIVATE LIMITED');
  });

  test('writeFormPGJGujaratHeaderMetaToExcelJsWorksheet fills Name of the employer and Month', async () => {
    const ExcelJS = require('exceljs');
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Sheet1');
    ws.getCell(6, 2).value = 'Name of the Establishment :';
    ws.getCell(8, 2).value = 'Name of the employer:';
    ws.getCell(9, 2).value = 'Month:';
    const wrote = writeFormPGJGujaratHeaderMetaToExcelJsWorksheet(
      ws,
      {
        form_p_gj_employer: 'VAYONA ENERGY PRIVATE LIMITED',
        form_p_gj_month: 'December',
        statutory_establishment_name_address:
          'Maliya Site, No.1,(1) Khiras,Parava behind,Raveswar, Dopor,Taluka Maliya Dist.Morbi GJ',
      },
      { headerRowEnd: 10 }
    );
    expect(wrote).toBe(true);
    expect(String(ws.getCell(6, 2).value || '')).toMatch(
      /Name of the Establishment\s*:\s*Maliya Site, No\.1/i
    );
    expect(String(ws.getCell(8, 2).value || '')).toMatch(/Name of the employer:\s*VAYONA ENERGY PRIVATE LIMITED/i);
    expect(String(ws.getCell(9, 2).value || '')).toMatch(/Month:\s*December/i);
  });

  test('Excel download fills header meta and boxes B + signature columns for 3 data rows', async () => {
    const ExcelJS = require('exceljs');
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Sheet1');
    ws.getCell(1, 2).value = 'FORM - P';
    ws.getCell(2, 2).value = '(See rule 26)';
    ws.getCell(3, 2).value = 'MUSTER-ROLL CUM WAGE REGISTER';
    ws.getCell(6, 2).value = 'Name of the Establishment :';
    ws.getCell(8, 2).value = 'Name of the employer:';
    ws.getCell(9, 2).value = 'Month:';
    // Table header row 11 (Excel 1-based) — Sr.No in column B.
    ws.getCell(11, 2).value = 'Sr. No';
    ws.getCell(11, 3).value = 'Full Name of the Worker';
    ws.getCell(11, 4).value = 'Designation';
    ws.getCell(11, 5).value = 'Date of entry into service';
    ws.getCell(11, 6).value = 'Working hours';
    ws.getCell(11, 7).value = 'Interval for Rest';
    ws.getCell(11, 8).value = 'Date of Month';
    ws.getCell(11, 9).value = 'Date of Payment';
    ws.getCell(11, 10).value = 'Signature/Thumb Impression';
    ws.getCell(12, 6).value = 'From';
    ws.getCell(12, 8).value = '1';
    const templateArrayBuffer = await wb.xlsx.writeBuffer();

    const rows = [
      { 'Sr. No.': '1', 'Name of the Worker': 'Harshad Dhokiya', Designation: 'Engineer' },
      { 'Sr. No.': '2', 'Name of the Worker': 'Shailesh R. Patel', Designation: 'Engineer' },
      { 'Sr. No.': '3', 'Name of the Worker': 'Chandreshbhai', Designation: 'Engineer' },
    ];

    const { blob } = await buildFormPGJGujaratWorkbookWithTemplateStyles({
      templateArrayBuffer,
      mappedData: rows,
      headersToUse: ['Sr. No.', 'Name of the Worker', 'Designation'],
      parsedFormHeader: { title: 'FORM P' },
      headerFormData: {
        form_p_gj_employer: 'VAYONA ENERGY PRIVATE LIMITED',
        form_p_gj_month: 'December',
        statutory_principal_employer: 'VAYONA ENERGY PRIVATE LIMITED',
        statutory_establishment_name_address:
          'Maliya Site, No.1,(1) Khiras,Parava behind,Raveswar, Dopor,Taluka Maliya Dist.Morbi GJ',
      },
      formFileName: 'Form_P_GJ.xlsx',
    });

    const outWb = new ExcelJS.Workbook();
    const outBuffer = Buffer.from(await new Response(blob).arrayBuffer());
    await outWb.xlsx.load(outBuffer);
    const outWs = outWb.worksheets[0];

    expect(String(outWs.getCell(8, 2).value || '')).toMatch(
      /Name of the employer:\s*VAYONA ENERGY PRIVATE LIMITED/i
    );
    expect(String(outWs.getCell(9, 2).value || '')).toMatch(/Month:\s*December/i);

    // 3 employee rows start at headerRow+2 = 13.
    expect(String(outWs.getCell(13, 3).value || '')).toMatch(/Harshad/i);
    expect(String(outWs.getCell(14, 3).value || '')).toMatch(/Shailesh/i);
    expect(String(outWs.getCell(15, 3).value || '')).toMatch(/Chandresh/i);

    for (let r = 13; r <= 15; r += 1) {
      expect(excelJSCellHasFullBoxBorder(outWs.getCell(r, 2))).toBe(true);
      expect(excelJSCellHasFullBoxBorder(outWs.getCell(r, 10))).toBe(true);
      expect(Number(outWs.getCell(r, 3).font?.size)).toBe(FORM_PGJ_TABLE_FONT_SIZE);
    }
  });

  test('Excel download groups Date of Month (9) across days 1–31 at font 8', async () => {
    const ExcelJS = require('exceljs');
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Sheet1');
    ws.getCell(1, 2).value = 'FORM - P';
    ws.getCell(11, 2).value = 'Sr. No';
    ws.getCell(11, 3).value = 'Full Name of the Worker';
    ws.getCell(11, 4).value = 'Designation';
    ws.getCell(11, 5).value = 'Working hours';
    ws.getCell(11, 6).value = 'Interval for Rest';
    ws.getCell(11, 7).value = 'Date of Month';
    for (let d = 1; d <= 31; d += 1) {
      ws.getCell(12, 6 + d).value = String(d);
    }
    ws.getCell(11, 38).value = 'Total Days Worked';
    ws.getCell(11, 39).value = 'Date of Payment';
    ws.getCell(11, 40).value = 'Signature/Thumb Impression';
    const templateArrayBuffer = await wb.xlsx.writeBuffer();

    const { blob } = await buildFormPGJGujaratWorkbookWithTemplateStyles({
      templateArrayBuffer,
      mappedData: [{ 'Sr. No.': '1', 'Name of the Worker': 'Harshad Dhokiya', Designation: 'Engineer' }],
      headersToUse: ['Sr. No.', 'Name of the Worker', 'Designation'],
      parsedFormHeader: { title: 'FORM P' },
      headerFormData: {},
      formFileName: 'Form_P_GJ.xlsx',
    });

    const outWb = new ExcelJS.Workbook();
    await outWb.xlsx.load(Buffer.from(await new Response(blob).arrayBuffer()));
    const outWs = outWb.worksheets[0];
    expect(String(outWs.getCell(11, 7).value || '')).toBe(FORM_PGJ_DATE_OF_MONTH_GROUP_LABEL);
    expect(Number(outWs.getCell(11, 7).font?.size)).toBe(FORM_PGJ_TABLE_FONT_SIZE);
    expect(Number(outWs.getCell(13, 3).font?.size)).toBe(FORM_PGJ_TABLE_FONT_SIZE);
    const grouped = (outWs.model?.merges || []).some((label) => {
      const parts = String(label || '').split(':');
      if (parts.length !== 2) return false;
      const tl = outWs.getCell(parts[0]);
      const br = outWs.getCell(parts[1]);
      return tl.row === 11 && tl.col === 7 && br.row === 11 && br.col >= 37;
    });
    expect(grouped).toBe(true);
  });

  test('resolveFormPGJGujaratDaysInMonth uses calendar length', () => {
    expect(resolveFormPGJGujaratDaysInMonth('June', '2026')).toBe(30);
    expect(resolveFormPGJGujaratDaysInMonth('July', '2026')).toBe(31);
    expect(resolveFormPGJGujaratDaysInMonth('February', '2026')).toBe(28);
    expect(resolveFormPGJGujaratDaysInMonth('February', '2024')).toBe(29);
  });

  test('countFormPGJGujaratDaysWorkedFromAttendance skips A / WO', () => {
    expect(
      countFormPGJGujaratDaysWorkedFromAttendance({
        Date_of_Month_1: 'P',
        'Date of Month_1': 'P',
        'Date of Month_2': 'P',
        'Date of Month_3': 'WO',
        'Date of Month_4': 'A',
        'Date of Month_5': 'CL',
      })
    ).toBe(3);
  });

  test('Excel download for June shows days 1–30 and hides 31', async () => {
    const ExcelJS = require('exceljs');
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Sheet1');
    ws.getCell(1, 2).value = 'FORM - P';
    ws.getCell(6, 2).value = 'Name of the Establishment :';
    ws.getCell(8, 2).value = 'Name of the employer:';
    ws.getCell(9, 2).value = 'Month:';
    ws.getCell(11, 2).value = 'Sr. No';
    ws.getCell(11, 3).value = 'Full Name of the Worker';
    ws.getCell(11, 4).value = 'Designation';
    ws.getCell(11, 5).value = 'Working hours';
    ws.getCell(11, 6).value = 'Interval for Rest';
    ws.getCell(11, 7).value = 'Date of Month';
    for (let d = 1; d <= 31; d += 1) {
      ws.getCell(12, 6 + d).value = String(d);
    }
    ws.getCell(11, 38).value = 'Total Days Worked';
    ws.getCell(11, 39).value = 'Date of Payment';
    ws.getCell(11, 40).value = 'Signature/Thumb Impression';
    const templateArrayBuffer = await wb.xlsx.writeBuffer();

    const { blob } = await buildFormPGJGujaratWorkbookWithTemplateStyles({
      templateArrayBuffer,
      mappedData: [
        {
          'Sr. No.': '1',
          'Name of the Worker': 'Harshad Dhokiya',
          Designation: 'Engineer',
          'Total Days Worked': 30,
          Date_of_Month_31: 'P',
        },
      ],
      headersToUse: ['Sr. No.', 'Name of the Worker', 'Designation'],
      parsedFormHeader: { title: 'FORM P' },
      headerFormData: {
        form_p_gj_month: 'June',
        form_x_year: '2026',
        form_p_gj_employer: 'VAYONA ENERGY PRIVATE LIMITED',
        statutory_establishment_name_address:
          'Maliya Site, No.1,(1) Khiras,Parava behind,Raveswar, Dopor,Taluka Maliya Dist.Morbi GJ',
      },
      formFileName: 'Form_P_GJ.xlsx',
    });

    const outWb = new ExcelJS.Workbook();
    await outWb.xlsx.load(Buffer.from(await new Response(blob).arrayBuffer()));
    const outWs = outWb.worksheets[0];
    expect(String(outWs.getCell(12, 36).value || '')).toBe('30');
    expect(String(outWs.getCell(12, 37).value || '')).toBe('');
    expect(outWs.getColumn(37).hidden).toBe(true);
    expect(String(outWs.getCell(13, 38).value ?? '')).toBe('30');
    expect(String(outWs.getCell(6, 2).value || '')).toMatch(
      /Name of the Establishment\s*:\s*Maliya Site, No\.1,\(1\) Khiras/i
    );
    expect(outWs.getCell(6, 2).alignment?.wrapText).toBe(true);
    expect(Number(outWs.getRow(6).height)).toBeGreaterThanOrEqual(32);
    expect(String(outWs.getCell(8, 2).value || '')).toMatch(
      /Name of the employer:\s*VAYONA ENERGY PRIVATE LIMITED/i
    );
    expect(outWs.getCell(8, 2).alignment?.wrapText).toBe(true);
    const grouped = (outWs.model?.merges || []).some((label) => {
      const parts = String(label || '').split(':');
      if (parts.length !== 2) return false;
      const tl = outWs.getCell(parts[0]);
      const br = outWs.getCell(parts[1]);
      return tl.row === 11 && tl.col === 7 && br.row === 11 && br.col === 36;
    });
    expect(grouped).toBe(true);
  });

  test('formPGJLeaveCodeFromType maps Contingency / Earned leave', () => {
    expect(formPGJLeaveCodeFromType('Contingency Leave')).toBe('CL');
    expect(formPGJLeaveCodeFromType('Earned Leave')).toBe('EL');
    expect(formPGJLeaveCodeFromType('Sick Leave')).toBe('SL');
  });

  test('applyFormPGJGujaratApprovedLeaveToDayColumns stamps June leave for Patel Chandrakant Virabhai', () => {
    const headers = ['Name of the Worker', 'Date of Month_11', 'Date of Month_12', 'Date of Month_13'];
    const mappedData = [
      {
        'Name of the Worker': 'Patel Chandrakant Virabhai',
        'Date of Month_11': 'A',
        'Date of Month_12': 'A',
        'Date of Month_13': 'A',
      },
    ];
    const employees = [
      {
        FirstName: 'Patel Chandrakant',
        LastName: 'Virabhai',
        EmployeeID: 'VE0104',
      },
    ];
    const approved = [
      {
        'Leave Type': 'Contingency Leave',
        From: '11-Jun-2026',
        To: '13-Jun-2026',
        Days: {
          '11-Jun-2026': { LeaveCount: 0 },
          '12-Jun-2026': { LeaveCount: 1 },
          '13-Jun-2026': { LeaveCount: 0 },
        },
        ApprovalStatus: 'APPROVED',
        'Employee Name': 'Patel Chandrakant Virabhai',
        EmployeeID: 'VE0104',
      },
    ];
    const hits = applyFormPGJGujaratApprovedLeaveToDayColumns(
      mappedData,
      employees,
      headers,
      approved,
      { monthFrom: '01-Jun-2026', monthTo: '30-Jun-2026' }
    );
    expect(hits).toBe(1);
    expect(mappedData[0]['Date of Month_11']).toBe('A');
    expect(mappedData[0]['Date of Month_12']).toBe('CL');
    expect(mappedData[0]['Date of Month_13']).toBe('A');
  });
});
