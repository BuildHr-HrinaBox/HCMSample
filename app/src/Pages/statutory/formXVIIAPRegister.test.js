import {
  applyFormXVIIAPHeaderLayoutFixes,
  applyFormXVIIAPPfPtToRow,
  collapseFormXVIIAPRepeatedText,
  FORM_XVII_AP_DEFAULT_CONTRACTOR_NAME,
  formXVIIAPPayrollRowHasPfOrPt,
  inferFormXVIIAPDeductionExportHeaders,
  isFormXVIIAPPfHeader,
  isFormXVIIAPPtHeader,
  isFormXVIIAPTotalDeductionsHeader,
  locateFormXVIIAPDeductionColumnsFromHeaderCells,
  pickFormXVIIAPExportPtAndTotalDeductions,
  preferFormXVIIAPPayrollRowsWithPfPt,
  readFormXVIIAPExportAmount,
  resolveFormXVIIAPSamplePayrollPf,
  resolveFormXVIIAPSamplePayrollPt,
} from './formXVIIAPRegister';

const FORM_XVII_AP_HEADERS = [
  'S.No',
  'Name of workman',
  'Basic wages',
  'Dearness Allowances',
  'Over time',
  'Other cash payments',
  'Total',
  'ESI',
  'PF',
  'PT',
  'SEVA',
  'Total Deductions',
  'Net Amount paid',
];

describe('Form XVII AP Sample Payroll PF / PT', () => {
  test('matches PF / PT deduction leaves and not registration columns', () => {
    expect(isFormXVIIAPPfHeader('PF')).toBe(true);
    expect(isFormXVIIAPPfHeader('Provident Fund')).toBe(true);
    expect(isFormXVIIAPPfHeader('Deductions, If any, (Indicate nature)\nPF')).toBe(true);
    expect(isFormXVIIAPPfHeader('EPF/UAN No.')).toBe(false);
    expect(isFormXVIIAPPfHeader('Total Deductions')).toBe(false);
    expect(isFormXVIIAPPtHeader('PT')).toBe(true);
    expect(isFormXVIIAPPtHeader('Professional Tax')).toBe(true);
    expect(isFormXVIIAPPtHeader('Deductions, If any, (Indicate nature)\nPT')).toBe(true);
    expect(isFormXVIIAPPtHeader('Total Deductions')).toBe(false);
  });

  test('reads PF and Professional Tax from Sample Payroll table keys', () => {
    expect(resolveFormXVIIAPSamplePayrollPf({ PF: 1800, professional_tax: 200 })).toBe(1800);
    expect(resolveFormXVIIAPSamplePayrollPt({ PF: 1800, 'Professional Tax': 200 })).toBe(200);
    expect(resolveFormXVIIAPSamplePayrollPf({ pf: 3562, professionalTax: 208 })).toBe(3562);
    expect(resolveFormXVIIAPSamplePayrollPt({ pf: 3562, professionalTax: 208 })).toBe(208);
    expect(resolveFormXVIIAPSamplePayrollPf({ epf_contribution: 3562 })).toBe(3562);
    expect(resolveFormXVIIAPSamplePayrollPf({ employer_pf: 3562 })).toBe(3562);
  });

  test('applies Sample Payroll PF / PT onto Form XVII deduction columns', () => {
    const row = {
      PF: 'Enter PF',
      PT: 'Enter PT',
      ESI: '',
      SEVA: '',
      'Total Deductions': 4039,
    };
    const hit = applyFormXVIIAPPfPtToRow(
      row,
      { PF: 3562, professional_tax: 208, gross_pay: 135142, net_pay: 131103 },
      FORM_XVII_AP_HEADERS
    );
    expect(hit).toBe(true);
    expect(row.PF).toBe(3562);
    expect(row.PT).toBe(208);
    expect(row.ESI).toBe('');
    expect(row.SEVA).toBe('');
    expect(row['Total Deductions']).toBe(4039);
  });

  test('prefers Sample Payroll rows that carry PF / PT over gross-only pay-run rows', () => {
    const rows = preferFormXVIIAPPayrollRowsWithPfPt(
      [{ gross_pay: 135142, net_pay: 131103 }],
      [{ PF: 3562, professional_tax: 208, gross_pay: 135142, net_pay: 131103 }]
    );
    expect(rows).toHaveLength(1);
    expect(formXVIIAPPayrollRowHasPfOrPt(rows[0])).toBe(true);
    expect(resolveFormXVIIAPSamplePayrollPf(rows[0])).toBe(3562);
    expect(resolveFormXVIIAPSamplePayrollPt(rows[0])).toBe(208);
  });

  test('infers PT after PF and Total Deductions before Net when Excel leaves are blank', () => {
    expect(
      inferFormXVIIAPDeductionExportHeaders([
        'Basic wages',
        'Total',
        'ESI',
        'PF',
        'Column 15',
        'SEVA',
        'Column 17',
        'Net Amount paid',
      ])
    ).toEqual([
      'Basic wages',
      'Total',
      'ESI',
      'PF',
      'PT',
      'SEVA',
      'Total Deductions',
      'Net Amount paid',
    ]);
  });

  test('relabels duplicate PF leaf after PF as PT for merged Form XVII headers', () => {
    expect(
      inferFormXVIIAPDeductionExportHeaders([
        'Total',
        'ESI',
        'PF',
        'PF',
        'SEVA',
        'Column 17',
        'Net Amount paid',
      ])
    ).toEqual([
      'Total',
      'ESI',
      'PF',
      'PT',
      'SEVA',
      'Total Deductions',
      'Net Amount paid',
    ]);
  });

  test('reads PT / Total Deductions from Autofill row keys during export', () => {
    const row = {
      PF: 3839,
      PT: 200,
      'Total Deductions': 4039,
      'Net Amount paid': 76951,
    };
    expect(readFormXVIIAPExportAmount(row, isFormXVIIAPPtHeader)).toBe(200);
    expect(readFormXVIIAPExportAmount(row, isFormXVIIAPTotalDeductionsHeader)).toBe(4039);
    expect(pickFormXVIIAPExportPtAndTotalDeductions({ PF: 3839 }, row)).toEqual({
      pf: 3839,
      pt: 200,
      totalDeductions: 4039,
    });
  });

  test('reads PT from duplicate PF (2) column when merged headers dedupe', () => {
    const headers = ['Total', 'ESI', 'PF', 'PF (2)', 'SEVA', 'Column 17', 'Net Amount paid'];
    const row = {
      PF: 3839,
      'PF (2)': 200,
      'Column 17': 4039,
      'Net Amount paid': 76951,
    };
    expect(pickFormXVIIAPExportPtAndTotalDeductions(row, null, headers)).toEqual({
      pf: 3839,
      pt: 200,
      totalDeductions: 4039,
    });
  });

  test('derives PT from payroll and Total Deductions from gross − net when grid keys are blank', () => {
    expect(
      pickFormXVIIAPExportPtAndTotalDeductions({
        PF: 3562,
        Total: 80990,
        'Net Amount paid': 76951,
        gross_pay: 80990,
        net_pay: 76951,
        professional_tax: 208,
      })
    ).toEqual({
      pf: 3562,
      pt: 208,
      totalDeductions: 4039,
    });
  });

  test('derives Total Deductions as PF + PT when gross / net are missing', () => {
    expect(
      pickFormXVIIAPExportPtAndTotalDeductions({
        PF: 3839,
        professional_tax: 200,
      })
    ).toEqual({
      pf: 3839,
      pt: 200,
      totalDeductions: 4039,
    });
  });

  test('locates PT after PF and Total Deductions before Net from Excel header cells', () => {
    expect(
      locateFormXVIIAPDeductionColumnsFromHeaderCells([
        { col: 12, text: 'Total' },
        { col: 13, text: 'ESI' },
        { col: 14, text: 'PF' },
        { col: 15, text: 'PT' },
        { col: 16, text: 'SEVA' },
        { col: 17, text: 'Total Deductions' },
        { col: 18, text: 'Net Amount paid' },
      ])
    ).toEqual({ pfCol: 14, ptCol: 15, totalDedCol: 17, netCol: 18 });
    expect(
      locateFormXVIIAPDeductionColumnsFromHeaderCells([
        { col: 13, text: 'ESI' },
        { col: 14, text: 'PF' },
        { col: 16, text: 'SEVA' },
        { col: 18, text: 'Net Amount paid' },
      ])
    ).toEqual({ pfCol: 14, ptCol: 15, totalDedCol: 17, netCol: 18 });
  });
});

describe('Form XVII AP header layout fixes', () => {
  test('collapses concatenated nature / location text', () => {
    expect(collapseFormXVIIAPRepeatedText('AP-NimbagalluAP-Nimbagallu')).toBe('AP-Nimbagallu');
    expect(collapseFormXVIIAPRepeatedText('AP-Nimbagallu / AP-Nimbagallu')).toBe('AP-Nimbagallu');
    expect(collapseFormXVIIAPRepeatedText('AP-Tadipatri AP-Tadipatri')).toBe('AP-Tadipatri');
  });

  test('places company establishment below the split label and keeps nature once', async () => {
    const ExcelJS = (await import('exceljs')).default;
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('XVII-Register of Wages');
    ws.getCell(5, 1).value = 'Name and address of Contractor:';
    ws.getCell(5, 3).value = 'VAYONA ENERGY';
    ws.getCell(5, 4).value = 'VAYONA ENERGY PRIVATE LIMITED';
    ws.getCell(5, 9).value = 'Name and address of Establishemnt in/';
    ws.getCell(6, 9).value = 'under which contract is carried on:';
    ws.getCell(5, 13).value = 'VAYONA ENERGY PRIVATE LIMITED';
    ws.getCell(7, 1).value = 'Nature and location of work:';
    ws.getCell(7, 3).value = 'AP-Tadipatri';
    ws.getCell(7, 4).value = 'AP-Tadipatri';

    applyFormXVIIAPHeaderLayoutFixes(ws, {
      headerFormData: {
        form_xvii_establishment_contract_carried: 'VAYONA ENERGY PRIVATE LIMITED, Hyderabad',
        form_xvii_nature_location_work: 'AP-Tadipatri',
        form_xvii_contractor: 'Site Contractor Pvt Ltd',
      },
      parsedFormHeader: {
        fields: [
          { key: 'form_xvii_contractor', label: 'Name and Address of Contractor.' },
          { key: 'form_xvii_nature_location_work', label: 'Nature and location of work.' },
          {
            key: 'form_xvii_establishment_contract_carried',
            label: 'Name and address of establishment in/under which contract is carried on',
          },
        ],
      },
      headerRowEnd: 12,
      maxCol: 20,
    });

    expect(String(ws.getCell(5, 1).value || '')).toMatch(/Name and address of Contractor/i);
    expect(String(ws.getCell(5, 1).value || '')).toContain('Site Contractor Pvt Ltd');
    expect(ws.getCell(5, 1).alignment?.wrapText).toBe(false);
    expect(String(ws.getCell(7, 1).value || '')).toMatch(/Nature and location of work/i);
    expect(String(ws.getCell(7, 1).value || '')).toContain('AP-Tadipatri');
    expect(String(ws.getCell(7, 1).value || '')).not.toMatch(/AP-Tadipatri.*AP-Tadipatri/i);
    expect(String(ws.getCell(5, 9).value || '')).toMatch(/Establishment/i);
    expect(String(ws.getCell(5, 9).value || '')).toContain('VAYONA ENERGY PRIVATE LIMITED');
    expect(ws.getCell(5, 9).alignment?.wrapText).toBe(false);
    expect(String(ws.getCell(6, 9).value || '')).toBe('');
  });

  test('defaults contractor to VAYONA ENERGY PRIVATE LIMITED and writes establishment in a taller box', async () => {
    const ExcelJS = (await import('exceljs')).default;
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('XVII-Register of Wages');
    ws.getCell(5, 1).value = 'Name and address of Contractor:';
    ws.getCell(5, 9).value = 'Name and address of Establishemnt in/';
    ws.getCell(6, 9).value = 'under which contract is carried on:';
    ws.getCell(7, 1).value = 'Nature and location of work:AP-NimbagalluAP-Nimbagallu';
    ws.getCell(7, 3).value = 'AP-Nimbagallu';
    ws.getCell(9, 1).value = 'Wage period: monthly:';
    ws.mergeCells(9, 10, 10, 18);
    ws.getCell(9, 10).value = 'Name and address of Establishment : leftover';
    ws.getRow(9).height = 15;
    ws.getRow(10).height = 15;

    applyFormXVIIAPHeaderLayoutFixes(ws, {
      headerFormData: {
        form_xvii_establishment_contract_carried:
          'VAYONA ENERGY PRIVATE LIMITED, Vayona Energy Pvt Ltd, 277A, Ranerimangalam road, Govindanagaram',
        form_xvii_nature_location_work: 'AP-NimbagalluAP-Nimbagallu',
      },
      parsedFormHeader: {
        fields: [
          { key: 'form_xvii_contractor', label: 'Name and Address of Contractor.' },
          { key: 'form_xvii_nature_location_work', label: 'Nature and location of work.' },
          {
            key: 'form_xvii_establishment_contract_carried',
            label: 'Name and address of establishment in/under which contract is carried on',
          },
        ],
      },
      headerRowEnd: 12,
      maxCol: 20,
    });

    expect(FORM_XVII_AP_DEFAULT_CONTRACTOR_NAME).toBe('VAYONA ENERGY PRIVATE LIMITED');
    expect(String(ws.getCell(5, 1).value || '')).toMatch(/Name and address of Contractor/i);
    expect(String(ws.getCell(5, 1).value || '')).toContain(FORM_XVII_AP_DEFAULT_CONTRACTOR_NAME);
    expect(ws.getCell(5, 1).alignment?.wrapText).toBe(false);
    expect(String(ws.getCell(7, 1).value || '')).toMatch(/Nature and location of work/i);
    expect(String(ws.getCell(7, 1).value || '')).toContain('AP-Nimbagallu');
    expect(String(ws.getCell(7, 1).value || '')).not.toMatch(/AP-Nimbagallu.*AP-Nimbagallu/i);
    expect(String(ws.getCell(5, 9).value || '')).toContain('VAYONA ENERGY PRIVATE LIMITED');
    expect(String(ws.getCell(5, 9).value || '')).toMatch(/Establishment/i);
    expect(ws.getCell(5, 9).alignment?.wrapText).toBe(false);
    expect(ws.getRow(5).height).toBe(20);
    expect(String(ws.getCell(9, 10).value || '')).toBe('');
  });
});
