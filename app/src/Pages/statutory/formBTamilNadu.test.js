import ExcelJS from 'exceljs';
import { excelJSCellHasFullBoxBorder } from '../../utils/excelTableBorders';
import {
  detectFormBTamilNaduPdfGroupBands,
  FORM_B_TN_AMOUNTS_DEDUCTED_GROUP_LABEL,
  isFormBTamilNaduAmountsDeductedGroupLabel,
  applyFormBTamilNaduEstablishmentFromCompany,
  applyFormBTamilNaduSummaryTotals,
  computeFormBTamilNaduBalanceDue,
  computeFormBTamilNaduBasicPlusHra,
  computeFormBTamilNaduOtherDeductions,
  extractFormBTamilNaduSummaryExportValues,
  formatFormBTamilNaduEstablishmentLine,
  formatFormBTamilNaduMonthDisplay,
  formatFormBTamilNaduMonthLine,
  formBTamilNaduSummaryHasExportAmounts,
  isFormBTamilNaduAmountActuallyPaidHeader,
  isFormBTamilNaduBalanceDueHeader,
  isFormBTamilNaduEstablishmentFromCompanyContext,
  isFormBTamilNaduOtherDeductionsHeader,
  isFormBTamilNaduPayrollSummaryHeader,
  isFormBTamilNaduPdfSignatoryRow,
  isFormBTamilNaduTotalEmolumentsHeader,
  looksLikeFormBTamilNaduPdfContext,
  normalizeFormBTamilNaduPdfMatrix,
  resolveFormBTamilNaduBasicAndHra,
  resolveFormBTamilNaduHeaderExportValues,
  rewriteFormBTamilNaduPdfHeader,
  summarizeFormBTamilNaduPayrollAmounts,
  writeFormBTamilNaduHeaderFieldsToWorksheet,
} from './formBTamilNadu';

describe('Form B Tamil Nadu payroll summary', () => {
  test('detects Form B TN for company-as-establishment from filename', () => {
    expect(
      isFormBTamilNaduEstablishmentFromCompanyContext({
        fileName: 'Form_B_-_TamilNadu.xlsx',
        rowItem: { state: 'Tamil Nadu' },
      })
    ).toBe(true);
  });

  test('overwrites establishment with company text (not Theni Site)', () => {
    const companyText =
      'Vayona Energy Pvt Ltd, 274 A, Rani mangammal main road, Govindanagaram - 625517 Theni Taluk, Theni District., Chennai, TamilNadu';
    const next = applyFormBTamilNaduEstablishmentFromCompany(
      { statutory_establishment_name_address: 'Theni Site' },
      companyText
    );
    expect(next.statutory_establishment_name_address).toBe(companyText);
  });

  test('detects Register of Wages column headers', () => {
    expect(
      isFormBTamilNaduTotalEmolumentsHeader(
        'Total emoluments payable during the month including basic wages, D.A, O.T., and bonus'
      )
    ).toBe(true);
    expect(isFormBTamilNaduOtherDeductionsHeader('Amounts deducted during the month_Other deductions')).toBe(
      true
    );
    expect(isFormBTamilNaduOtherDeductionsHeader('Fine')).toBe(false);
    expect(isFormBTamilNaduAmountActuallyPaidHeader('Amount actually paid during the month')).toBe(true);
    expect(isFormBTamilNaduBalanceDueHeader('Balance due to the employees')).toBe(true);
  });

  test('Other deductions = gross_pay − basic − hra', () => {
    expect(computeFormBTamilNaduOtherDeductions(100000, 40000, 20000)).toBe(40000);
    expect(computeFormBTamilNaduOtherDeductions('194566', '52628', '22476')).toBe(119462);
    expect(computeFormBTamilNaduOtherDeductions(50000, '', '')).toBe(50000);
  });

  test('Total emoluments helper still computes basic + hra for other-deduction math', () => {
    expect(computeFormBTamilNaduBasicPlusHra(40000, 20000)).toBe(60000);
    expect(computeFormBTamilNaduBasicPlusHra('67729', '32665')).toBe(100394);
    expect(computeFormBTamilNaduBasicPlusHra(50000, '')).toBe(50000);
  });

  test('forces Form W TN defaults so basic alone (67729) is never used without HRA', () => {
    const resolved = resolveFormBTamilNaduBasicAndHra({
      basic: '67729',
      hra: '',
      emp: { EmployeeID: 'VE0147', FirstName: 'Avudaiappan' },
    });
    expect(resolved).toEqual({ basic: '67729', hra: '32665' });
    expect(computeFormBTamilNaduBasicPlusHra(resolved.basic, resolved.hra)).toBe(100394);
  });

  test('Balance due is 0 when (gross − net) equals net_pay, or when net is paid', () => {
    // gross − net = net → 0
    expect(computeFormBTamilNaduBalanceDue(200000, 100000)).toBe(0);
    // normal paid payroll → 0
    expect(computeFormBTamilNaduBalanceDue(150000, 135000)).toBe(0);
    expect(computeFormBTamilNaduBalanceDue(1118499, 1007977)).toBe(0);
  });

  test('summarizes gross emoluments, net paid, and zero balance for location employees', () => {
    const totals = summarizeFormBTamilNaduPayrollAmounts([
      { gross: 100000, basic: 40000, hra: 20000, net: 90000 },
      { gross: 50000, basic: 30000, hra: 10000, net: 45000 },
      { gross: '', basic: 1, hra: 1, net: 1 },
    ]);
    expect(totals.payrollMatchedCount).toBe(2);
    expect(totals.emolumentsEmployeeCount).toBe(2);
    expect(totals.totalGrossPay).toBe(150000);
    expect(totals.totalEmoluments).toBe(150000);
    expect(totals.totalBasicPlusHra).toBe(100002);
    expect(totals.totalOtherDeductions).toBe(50000);
    expect(totals.totalAmountActuallyPaid).toBe(135000);
    expect(totals.totalNetPay).toBe(135000);
    expect(totals.totalBalanceDue).toBe(0);
  });

  test('does not count basic+hra-only rows toward total emoluments (needs gross_pay)', () => {
    const totals = summarizeFormBTamilNaduPayrollAmounts([
      { gross: '', basic: 67729, hra: 32665 },
      { gross: '', basic: 52628, hra: 22476 },
      { gross: '', basic: 67064, hra: 32345 },
      { gross: '', basic: 48145, hra: 22647 },
      { gross: '', basic: 43688, hra: 20704 },
    ]);
    expect(totals.emolumentsEmployeeCount).toBe(0);
    expect(totals.totalGrossPay).toBe(0);
    expect(totals.totalEmoluments).toBe(0);
    expect(totals.totalBasicPlusHra).toBe(410091);
  });

  test('Other deductions sums gross − basic − hra across employees', () => {
    const totals = summarizeFormBTamilNaduPayrollAmounts([
      { gross: 206820, basic: 67729, hra: 32665, net: 180000 },
      { gross: 194566, basic: 52628, hra: 22476, net: 170000 },
    ]);
    expect(totals.totalGrossPay).toBe(401386);
    expect(totals.totalEmoluments).toBe(401386);
    expect(totals.totalBasicPlusHra).toBe(67729 + 32665 + 52628 + 22476);
    expect(totals.totalOtherDeductions).toBe(106426 + 119462);
    expect(totals.totalAmountActuallyPaid).toBe(350000);
    expect(totals.totalBalanceDue).toBe(0);
  });

  test('applies totals onto the summary row headers', () => {
    const headers = [
      'Total emoluments payable during the month including basic wages, D.A, O.T., and bonus',
      'Amounts deducted during the month_Fine',
      'Amounts deducted during the month_Other deductions',
      'Amount actually paid during the month',
      'Balance due to the employees',
    ];
    const row = {};
    headers.forEach((h) => {
      row[h] = '';
    });
    applyFormBTamilNaduSummaryTotals(
      row,
      headers,
      {
        matchedCount: 2,
        totalGrossPay: 150000,
        totalEmoluments: 150000,
        totalOtherDeductions: 50000,
        totalAmountActuallyPaid: 135000,
        totalNetPay: 135000,
        totalBalanceDue: 0,
      },
      { sanitizeValue: (v) => String(v) }
    );
    expect(row[headers[0]]).toBe('150000');
    expect(row[headers[1]]).toBe('');
    expect(row[headers[2]]).toBe('50000');
    expect(row[headers[3]]).toBe('135000');
    expect(row[headers[4]]).toBe('0');
  });

  test('extracts Form B summary export values from Parent_Sub headers', () => {
    const headers = [
      'Total number of employees',
      'Total emoluments payable during the month including basic wages, D.A, O.T., and bonus',
      'Amounts deducted during the month_Fine',
      'Amounts deducted during the month_Other deductions',
      'Amount actually paid during the month',
      'Balance due to the employees',
    ];
    const row = {
      [headers[0]]: '5',
      [headers[1]]: '150000',
      [headers[2]]: '',
      [headers[3]]: '50000',
      [headers[4]]: '135000',
      [headers[5]]: '0',
    };
    const values = extractFormBTamilNaduSummaryExportValues(row, headers);
    expect(values.employeeCount).toBe('5');
    expect(values.totalEmoluments).toBe('150000');
    expect(values.otherDeductions).toBe('50000');
    expect(values.amountActuallyPaid).toBe('135000');
    expect(values.balanceDue).toBe('0');
    expect(formBTamilNaduSummaryHasExportAmounts(values)).toBe(true);
  });

  test('keeps Name of the Establishment and For the Month of on Excel label cells', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Form B');
    ws.mergeCells(5, 1, 5, 4);
    ws.mergeCells(6, 1, 6, 4);
    ws.getCell(5, 1).value = 'Name of the Establishment :';
    ws.getCell(6, 1).value = 'For the Month of';
    const establishment =
      'Vayona Energy Pvt Ltd, 274 A, Rani mangammal main road, Govindanagaram - 625517 Theni Taluk, Theni District, Tamilnadu';
    writeFormBTamilNaduHeaderFieldsToWorksheet(ws, {
      statutory_establishment_name_address: establishment,
      form_b_header_for_the_month_of: 'July2026',
    });
    expect(String(ws.getCell(5, 1).value)).toBe(
      `Name of the Establishment : ${establishment}`
    );
    expect(String(ws.getCell(6, 1).value)).toBe('For the Month of : July 2026');
    expect(Number(ws.getRow(5).height)).toBeGreaterThanOrEqual(48);
    expect(excelJSCellHasFullBoxBorder(ws.getCell(5, 1))).toBe(true);
    expect(excelJSCellHasFullBoxBorder(ws.getCell(6, 1))).toBe(true);
  });

  test('moves collapsed Fine amount back to Other deductions for PDF', () => {
    const normalized = normalizeFormBTamilNaduPdfMatrix(
      [
        [
          'Total number of employees',
          'Total emoluments payable during the month including basic wages, D.A, O.T., and bonus',
          'Fine',
          'Other deductions',
          'Amount actually paid during the month',
          'Balance due to the employees',
        ],
        ['', '', '(a)', '(b)', '', ''],
        ['15', '1163466', '458439', '', '1035636', '0'],
      ],
      6,
      0,
      []
    );
    expect(normalized.rows[2][2]).toBe('');
    expect(String(normalized.rows[2][3])).toBe('458439');
    expect(String(normalized.rows[2][4])).toBe('1035636');
  });

  test('formats glued month tokens and header export values', () => {
    expect(formatFormBTamilNaduMonthDisplay('July2026')).toBe('July 2026');
    expect(formatFormBTamilNaduMonthLine('july 2026')).toBe('For the Month of : July 2026');
    expect(formatFormBTamilNaduEstablishmentLine('Vayona Energy Pvt Ltd')).toBe(
      'Name of the Establishment : Vayona Energy Pvt Ltd'
    );
    const resolved = resolveFormBTamilNaduHeaderExportValues({
      statutory_establishment_name: 'Vayona Energy Pvt Ltd',
      form_b_header_for_the_month_of: 'July2026',
    });
    expect(resolved.establishment).toBe('Vayona Energy Pvt Ltd');
    expect(resolved.month).toBe('July 2026');
  });

  test('rewrites PDF header with labels and drops signatory lines', () => {
    const rewritten = rewriteFormBTamilNaduPdfHeader(
      [
        'Form B',
        '(See rule 29)',
        'Tamilnadu Labour Welfare Fund Rules, 1973',
        'Register of Wages',
        'July2026',
        'For (Company Name)',
        'Authorised Signatory',
      ],
      [
        'Vayona Energy Pvt Ltd, 274 A, Rani mangammal main road, Govindanagaram - 625517 Theni Taluk, Theni District, Tamilnadu',
        'Signature of Employer / Manager / Authorised Person',
      ]
    );
    expect(rewritten.titles).toEqual([
      'Form B',
      '(See rule 29)',
      'Tamilnadu Labour Welfare Fund Rules, 1973',
      'Register of Wages',
    ]);
    expect(rewritten.fields[0]).toMatch(/^Name of the Establishment\s*:/);
    expect(rewritten.fields.some((f) => /^For the Month of\s*:\s*July 2026$/i.test(f))).toBe(true);
    expect(rewritten.fields.join('\n')).not.toMatch(/Authorised Signatory/i);
    expect(rewritten.fields.join('\n')).not.toMatch(/For\s*\(/);
    expect(rewritten.fields.join('\n')).not.toMatch(/Signature of Employer/i);
  });

  test('spans Amounts deducted during the month across Fine and Other deductions', () => {
    expect(isFormBTamilNaduAmountsDeductedGroupLabel('Amounts deducted during the month')).toBe(
      true
    );
    expect(isFormBTamilNaduAmountsDeductedGroupLabel('Other deductions')).toBe(false);
    const bands = detectFormBTamilNaduPdfGroupBands(
      [
        [
          'Total number of employees',
          'Total emoluments payable during the month including basic wages, D.A, O.T., and bonus',
          'Amounts deducted during the month',
          '',
          'Amount actually paid during the month',
          'Balance due to the employees',
        ],
        ['', '', 'Fine', 'Other deductions', '', ''],
        ['', '', '(a)', '(b)', '', ''],
      ],
      0,
      2,
      6
    );
    expect(bands).toHaveLength(1);
    expect(bands[0].start).toBe(2);
    expect(bands[0].end).toBe(3);
    expect(bands[0].label).toBe(FORM_B_TN_AMOUNTS_DEDUCTED_GROUP_LABEL);
  });

  test('detects Form B TN PDF context and signatory rows', () => {
    expect(
      looksLikeFormBTamilNaduPdfContext(
        ['Form B', '(See rule 29)', 'Tamilnadu Labour Welfare Fund Rules, 1973', 'Register of Wages'],
        [['Total number of employees', 'Total emoluments payable during the month']],
        'Form B',
        'Form_B_-_TamilNadu.xlsx'
      )
    ).toBe(true);
    expect(isFormBTamilNaduPdfSignatoryRow(['For (Company Name)'])).toBe(true);
    expect(isFormBTamilNaduPdfSignatoryRow(['Authorised Signatory'])).toBe(true);
    expect(
      isFormBTamilNaduPdfSignatoryRow(['Signature of Employer / Manager / Authorised Person'])
    ).toBe(true);
    expect(isFormBTamilNaduPdfSignatoryRow(['15'])).toBe(false);
  });

  test('detects payroll summary headers blocked from statutory overlay', () => {
    expect(
      isFormBTamilNaduPayrollSummaryHeader(
        'Total emoluments payable during the month including basic wages, D.A, O.T., and bonus'
      )
    ).toBe(true);
    expect(isFormBTamilNaduPayrollSummaryHeader('Amount actually paid during the month')).toBe(true);
    expect(isFormBTamilNaduPayrollSummaryHeader('Balance due to the employees')).toBe(true);
    expect(isFormBTamilNaduPayrollSummaryHeader('Total number of employees')).toBe(true);
  });
});
