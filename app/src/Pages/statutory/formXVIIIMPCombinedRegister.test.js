import {
  FORM_XVIII_MP_LEAVE_CATEGORY,
  FORM_XVIII_MP_NIL,
  FORM_XVIII_MP_FOR_THE_MONTH_KEY,
  FORM_XVIII_MP_FOR_THE_MONTH_LABEL,
  applyFormXVIIIMPForTheMonthToHeaderData,
  applyFormXVIIIMPLeaveToRow,
  applyFormXVIIIMPNilDefaultsToRow,
  applyFormXVIIIMPPayrollToRow,
  computeFormXVIIIMPOtherAllowances,
  ensureFormXVIIIMPForTheMonthHeaderField,
  extractFormXVIIIMPMonthYearFromForTheMonthText,
  finalizeFormXVIIIMPOtherAllowancesForDownload,
  formatFormXVIIIMPForTheMonthDisplay,
  isFormXVIIIMPAnyOtherAmountHeader,
  isFormXVIIIMPCombinedRegisterContext,
  isFormXVIIIMPForTheMonthCellText,
  dedupeFormXVIIIMPHeaderFields,
  normalizeFormXVIIIMPHeaderFieldLabel,
  isFormXVIIIMPMaternityBenefitHeader,
  isFormXVIIIMPNilDefaultHeader,
  isFormXVIIIMPOtherDeductionsGroupParentHeader,
  isFormXVIIIMPOvertimeHoursHeader,
  isFormXVIIIMPOvertimeWagesHeader,
  formXVIIIMPRowsNeedPayrollEnrich,
  isFormXVIIIMPPayrollDeductionHeader,
  isFormXVIIIMPPfHeader,
  isFormXVIIIMPPtHeader,
  isFormXVIIIMPWageRateHeader,
  resolveFormXVIIIMPPayrollRowsForAutofill,
  resolveFormXVIIIMPForTheMonthFromHeaderData,
  mapFormXVIIIMPRowsFromEmployees,
  looksLikeKarnatakaFormTSheet,
  looksLikeMPCombinedRegisterSheet,
  looksLikeTamilNaduFormXVIIISheet,
  readFormXVIIIMPOtherAllowanceFromRow,
  remapMPCombinedRegisterRows,
  repairFormXVIIIMPDuplicateWageHeaders,
  repairFormXVIIIMPOtherDeductionsPfHeader,
  resolveFormXVIIIMPOtherAllowanceForWrite,
  resolveFormXVIIIMPSamplePayrollPf,
  resolveFormXVIIIMPSamplePayrollPt,
  resolveFormXVIIIMPSectionHeaderLabel,
  resolveFormXVIIIMPTableHeaders,
  restoreFormXVIIIMPDistinctOtherAllowances,
  snapshotFormXVIIIMPDistinctOtherAllowances,
  sumFormXVIIIMPLeaveBookedAndBalance,
  writeFormXVIIIMPForTheMonthToWorksheet,
} from './formXVIIIMPCombinedRegister';
import { flattenPayrollEarningColumns } from '../../utils/payrollEarnings';

describe('Form XVIII MP Combined Register mappings', () => {
  it('does not treat Karnataka Form T as MP Form XVIII combined register', () => {
    const formTBlob =
      'Form___T_-_Karnataka.xlsx FORM T COMBINED MUSTER ROLL CUM REGISTER OF WAGES Karnataka Rule 24(9-B)';
    expect(looksLikeKarnatakaFormTSheet(formTBlob)).toBe(true);
    expect(looksLikeMPCombinedRegisterSheet(formTBlob)).toBe(false);
    expect(
      isFormXVIIIMPCombinedRegisterContext(
        {
          title: 'FORM T',
          subtitle: 'COMBINED MUSTER ROLL CUM REGISTER OF WAGES',
          reference: '[See Rule 24(9-B) of Karnataka Shops & Commercial Establishment Rules, 1963]',
        },
        { formName: 'Form T', state: 'Karnataka' },
        'Form___T_-_Karnataka.xlsx',
        'COMBINED MUSTER ROLL CUM REGISTER OF WAGES'
      )
    ).toBe(false);
    expect(looksLikeKarnatakaFormTSheet('Form T_KA COMBINED MUSTER ROLL CUM REGISTER OF WAGES Karnataka')).toBe(
      true
    );
  });

  it('still detects Madhya Pradesh Form XVIII combined register', () => {
    const mpBlob =
      'Form XVIII Madhya Pradesh Muster Roll-cum Register of Wages Name of the establishment and address Location of work';
    expect(looksLikeKarnatakaFormTSheet(mpBlob)).toBe(false);
    expect(looksLikeTamilNaduFormXVIIISheet(mpBlob)).toBe(false);
    expect(looksLikeMPCombinedRegisterSheet(mpBlob)).toBe(true);
    expect(
      isFormXVIIIMPCombinedRegisterContext(
        { title: 'FORM XVIII', subtitle: 'Muster Roll-cum Register of Wages' },
        { formName: 'Form XVIII', state: 'Madhya Pradesh' },
        'Form_XVIII_MP.xlsx',
        'Name of the establishment and address Location of work'
      )
    ).toBe(true);
  });

  it('peopleOnly mapping fills worker name without waiting on payroll', () => {
    const headers = ['Sr No', 'EMP Id', 'Full name of the Worker (ID/ Tocken No. If any)', 'PF (22)'];
    const rows = mapFormXVIIIMPRowsFromEmployees(
      [{ EmployeeID: 'VE1257', FirstName: 'Tejpal', LastName: 'Singh' }],
      headers,
      { peopleOnly: true }
    );
    expect(rows[0]['EMP Id']).toBe('VE1257');
    expect(String(rows[0]['Full name of the Worker (ID/ Tocken No. If any)'])).toMatch(/Tejpal/i);
    expect(rows[0]['PF (22)']).toBe('');
  });

  it('dedupes repeated Form XVIII MP header labels so the modal stays short', () => {
    expect(normalizeFormXVIIIMPHeaderFieldLabel('2. Nature and location of work.')).toBe(
      'nature and location of work'
    );
    const fields = dedupeFormXVIIIMPHeaderFields([
      { key: 'form_xviii_est', label: 'Name and Address of the Establishment', value: 'Site A' },
      { key: 'form_xviii_est_addr', label: 'Address of the Establishment', value: '' },
      { key: 'form_xviii_contractor', label: 'Name and Address of Contractor', value: 'Vayona' },
      { key: 'form_xviii_nature', label: 'Nature and location of work', value: 'MP-Dhar' },
      { key: 'form_xviii_nature_dup', label: '2. Nature and location of work.', value: '' },
      { key: 'form_xviii_pe', label: 'Name and address of Principal Employer', value: 'PE' },
      { key: 'form_xviii_pe_mgr', label: 'Name and address of Principal Employer/ Manager', value: '' },
      { key: 'form_xviii_month_year', label: 'Month / Year', value: '' },
      { key: 'form_xviii_month', label: 'Month', value: '' },
    ]);
    expect(fields.map((f) => f.key)).toEqual([
      'form_xviii_est',
      'form_xviii_contractor',
      'form_xviii_nature',
      'form_xviii_pe',
      'form_xviii_month_year',
    ]);
  });

  it('does not treat Form_XVIII_-_TamilNadu.xlsx as MP Form XVIII combined register', () => {
    const tnBlob =
      'Form_XVIII_-_TamilNadu.xlsx Form XVIII – Register of Wages-cum-Muster Roll Form of Register of Wages-cum-Muster Roll [See rule 78(1)(a)(i)] Tamil Nadu';
    expect(looksLikeTamilNaduFormXVIIISheet(tnBlob)).toBe(true);
    expect(looksLikeMPCombinedRegisterSheet(tnBlob)).toBe(false);
    expect(
      isFormXVIIIMPCombinedRegisterContext(
        {
          title: 'Form XVIII – Register of Wages-cum-Muster Roll',
          subtitle: 'Form of Register of Wages-cum-Muster Roll',
          reference: '[See rule 78(1)(a)(i)]',
        },
        { formName: 'Form XVIII', state: 'Tamil Nadu' },
        'Form_XVIII_-_TamilNadu.xlsx',
        'Form XVIII – Register of Wages-cum-Muster Roll Daily attendance /units worked Amount of wages earned'
      )
    ).toBe(false);
  });

  it('computes Other allowances as gross_pay − basic − hra', () => {
    expect(computeFormXVIIIMPOtherAllowances(100000, 40000, 20000)).toBe(40000);
    expect(computeFormXVIIIMPOtherAllowances('113708', '50000', '20000')).toBe(43708);
    expect(computeFormXVIIIMPOtherAllowances('', 1, 1)).toBe('');
    // Missing basic+hra must not dump full gross into Other allowances
    expect(computeFormXVIIIMPOtherAllowances(70326, '', '')).toBe('');
    // Zero basic/hra must not mirror wage-rate/gross
    expect(computeFormXVIIIMPOtherAllowances(74132, 0, 0)).toBe('');
    // basic == gross ⇒ other allowances is 0 (not the wage-rate amount)
    expect(computeFormXVIIIMPOtherAllowances(74132, 74132, 0)).toBe(0);
  });

  it('repairs merged Wage-rate label on section column 14 to Other allowances', () => {
    expect(
      resolveFormXVIIIMPSectionHeaderLabel(
        'Wage rate/ pay or (piece rate/ wages per unit)',
        '14',
        'Wage rate/ pay or (piece rate/ wages per unit)'
      )
    ).toBe('Other allowances');
    expect(isFormXVIIIMPWageRateHeader('Other allowances')).toBe(false);
  });

  it('repairs consecutive Wage-rate labels even when section index wrongly says 13', () => {
    const wage = 'Wage rate/ pay or (piece rate/ wages per unit)';
    expect(resolveFormXVIIIMPSectionHeaderLabel(wage, '13', wage)).toBe('Other allowances');
    expect(resolveFormXVIIIMPSectionHeaderLabel(wage, '', wage)).toBe('Other allowances');
    expect(resolveFormXVIIIMPSectionHeaderLabel('', '14', wage)).toBe('Other allowances');
  });

  it('repairs duplicate Wage-rate header list to Other allowances', () => {
    const wage = 'Wage rate/ pay or (piece rate/ wages per unit)';
    expect(repairFormXVIIIMPDuplicateWageHeaders([wage, wage, 'Total/ gross Wages/ Earnings'])).toEqual([
      wage,
      'Other allowances',
      'Total/ gross Wages/ Earnings',
    ]);
  });

  it('keeps Other allowances when remapping headers after download rebuild', () => {
    const oldHeaders = [
      'Wage rate/ pay or (piece rate/ wages per unit)',
      'Other allowances',
      'Total/ gross Wages/ Earnings',
    ];
    const newHeaders = [
      'Wage rate/ pay or (piece rate/ wages per unit)',
      'Other allowances',
      'Total/ gross Wages/ Earnings',
    ];
    const rows = remapMPCombinedRegisterRows(
      [
        {
          'Wage rate/ pay or (piece rate/ wages per unit)': '74132',
          'Other allowances': '24132',
          'Total/ gross Wages/ Earnings': '74132',
        },
      ],
      oldHeaders,
      newHeaders
    );
    expect(rows[0]['Other allowances']).toBe('24132');
    expect(rows[0]['Wage rate/ pay or (piece rate/ wages per unit)']).toBe('74132');
  });

  it('preserves Other allowances when download rebuild duplicates Wage-rate header', () => {
    const wage = 'Wage rate/ pay or (piece rate/ wages per unit)';
    const newHeaders = [wage, wage, 'Total/ gross Wages/ Earnings'];
    const rows = remapMPCombinedRegisterRows(
      [
        {
          [wage]: '74132',
          'Other allowances': '32165',
          'Total/ gross Wages/ Earnings': '74132',
        },
      ],
      [wage, 'Other allowances', 'Total/ gross Wages/ Earnings'],
      newHeaders
    );
    expect(newHeaders[1]).toBe('Other allowances');
    expect(rows[0]['Other allowances']).toBe('32165');
    expect(rows[0][wage]).toBe('74132');
    expect(readFormXVIIIMPOtherAllowanceFromRow(rows[0], newHeaders)).toBe('32165');
  });

  it('snapshots and restores UI Other allowances when download mirrored wage rate', () => {
    const wage = 'Wage rate/ pay or (piece rate/ wages per unit)';
    const headers = [wage, 'Other allowances', 'Total/ gross Wages/ Earnings'];
    const uiRows = [
      {
        [wage]: '74132',
        'Other allowances': '32165',
        'Total/ gross Wages/ Earnings': '74132',
      },
    ];
    const snap = snapshotFormXVIIIMPDistinctOtherAllowances(uiRows, headers);
    expect(snap[0]).toBe('32165');
    const broken = [
      {
        [wage]: '74132',
        'Other allowances': '74132',
        'Total/ gross Wages/ Earnings': '74132',
      },
    ];
    restoreFormXVIIIMPDistinctOtherAllowances(broken, headers, snap);
    expect(broken[0]['Other allowances']).toBe('32165');
    expect(resolveFormXVIIIMPOtherAllowanceForWrite(broken[0], headers)).toBe('32165');
    expect(
      resolveFormXVIIIMPOtherAllowanceForWrite(
        { [wage]: '74132', 'Other allowances': '74132', 'Total/ gross Wages/ Earnings': '74132' },
        headers
      )
    ).toBe('');
  });

  it('does not write gross into Other allowances when basic/hra missing on download enrich', () => {
    const headers = [
      'Other allowances',
      'Total/ gross Wages/ Earnings',
      'Wage rate/ pay or (piece rate/ wages per unit)',
    ];
    const mpHeaders = resolveFormXVIIIMPTableHeaders(headers);
    const row = {
      'Other allowances': '24132',
      'Total/ gross Wages/ Earnings': '',
      'Wage rate/ pay or (piece rate/ wages per unit)': '',
    };
    applyFormXVIIIMPPayrollToRow(
      row,
      { gross_pay: 74132, net_pay: 70486, other_allowance: 74132 },
      mpHeaders,
      {
        headers,
        sanitizeValue: (v) => String(v ?? '').trim(),
        findDeductionAmount: () => '',
        getDeductionsArray: () => [],
      }
    );
    expect(row['Other allowances']).toBe('24132');
    expect(row['Wage rate/ pay or (piece rate/ wages per unit)']).toBe('74132');
    expect(row['Total/ gross Wages/ Earnings']).toBe('74132');
  });

  it('maps PF (22) and PT (22) from Sample Payroll PF / Professional Tax', () => {
    expect(isFormXVIIIMPPayrollDeductionHeader('PF (22)')).toBe(true);
    expect(isFormXVIIIMPPayrollDeductionHeader('PT (22)')).toBe(true);
    expect(resolveFormXVIIIMPSamplePayrollPf({ PF: 1800, professional_tax: 200 })).toBe(1800);
    expect(resolveFormXVIIIMPSamplePayrollPt({ PF: 1800, 'Professional Tax': 200 })).toBe(200);

    const headers = ['PF (22)', 'PT (22)', 'Total/ gross Wages/ Earnings'];
    const mpHeaders = resolveFormXVIIIMPTableHeaders(headers);
    expect(mpHeaders.pf).toBe('PF (22)');
    expect(mpHeaders.pt).toBe('PT (22)');

    const row = {};
    headers.forEach((h) => {
      row[h] = '';
    });
    applyFormXVIIIMPPayrollToRow(
      row,
      {
        gross_pay: 50000,
        net_pay: 45000,
        PF: 1800,
        professional_tax: 200,
      },
      mpHeaders,
      {
        headers,
        sanitizeValue: (v) => String(v ?? '').trim(),
        flattenPayrollEarningColumns: (r) => r,
      }
    );
    expect(row['PF (22)']).toBe('1800');
    expect(row['PT (22)']).toBe('200');
  });

  it('maps PF (22) / PT (22) from Sample Payroll UI keys through flatten', () => {
    expect(resolveFormXVIIIMPSamplePayrollPf({ pf: 3562, professionalTax: 208 })).toBe(3562);
    expect(resolveFormXVIIIMPSamplePayrollPt({ pf: 3562, professionalTax: 208 })).toBe(208);
    expect(isFormXVIIIMPPtHeader('PT (22)')).toBe(true);
    expect(
      isFormXVIIIMPPtHeader('Other Deductions Like EPF/ ESI/ Welfare Fund etc. (if any)_PT (22)')
    ).toBe(true);

    const headers = ['PF (22)', 'PT (22)'];
    const mpHeaders = resolveFormXVIIIMPTableHeaders(headers);
    const row = {};
    headers.forEach((h) => {
      row[h] = '';
    });
    applyFormXVIIIMPPayrollToRow(
      row,
      {
        employee_name: 'Test Worker',
        gross_pay: 50000,
        net_pay: 45000,
        pf: 3562,
        professionalTax: 208,
      },
      mpHeaders,
      {
        headers,
        sanitizeValue: (v) => String(v ?? '').trim(),
        flattenPayrollEarningColumns,
      }
    );
    expect(row['PF (22)']).toBe('3562');
    expect(row['PT (22)']).toBe('208');
  });

  it('prefers Sample Payroll rows that carry PF over gross-only pay-run rows', () => {
    const rows = resolveFormXVIIIMPPayrollRowsForAutofill(
      [
        { employee_name: 'A', gross_pay: 50000, net_pay: 45000 },
        { employee_name: 'B', gross_pay: 50000, net_pay: 45000, pf: 1800, professionalTax: 200 },
      ],
      ['2026-07'],
      { cachedSampleRows: [{ employee_name: 'A', gross_pay: 50000, net_pay: 45000 }] }
    );
    expect(
      rows.some((row) => Number(row.epf_contribution ?? row.pf ?? row.PF) === 1800)
    ).toBe(true);
    expect(
      formXVIIIMPRowsNeedPayrollEnrich(
        [{ 'PF (22)': '', 'PT (22)': '', 'Total/ gross Wages/ Earnings': '50000' }],
        ['PF (22)', 'PT (22)', 'Total/ gross Wages/ Earnings']
      )
    ).toBe(true);
  });

  it('fills PF when merged Other Deductions parent is the column key', () => {
    const otherDed =
      'Other Deductions Like EPF/ ESI/ Welfare Fund etc. (if any)';
    expect(isFormXVIIIMPOtherDeductionsGroupParentHeader(otherDed)).toBe(true);
    expect(isFormXVIIIMPPfHeader(otherDed)).toBe(true);
    expect(isFormXVIIIMPOtherDeductionsGroupParentHeader('Other Deductions')).toBe(true);
    expect(isFormXVIIIMPPfHeader('Other Deductions')).toBe(true);
    expect(
      repairFormXVIIIMPOtherDeductionsPfHeader([otherDed, 'ESIC (22)', 'PT (22)', 'LWF (22)'])[0]
    ).toBe('PF (22)');
    expect(
      resolveFormXVIIIMPSectionHeaderLabel(otherDed, '22', 'Deductions of Fines imposed, if any')
    ).toBe('PF (22)');

    const headers = [otherDed, 'ESIC (22)', 'PT (22)', 'LWF (22)'];
    const mpHeaders = resolveFormXVIIIMPTableHeaders(headers);
    expect(mpHeaders.pf).toBe(otherDed);
    expect(mpHeaders.pt).toBe('PT (22)');

    const row = {};
    headers.forEach((h) => {
      row[h] = '';
    });
    applyFormXVIIIMPPayrollToRow(
      row,
      { PF: 3562, professional_tax: 208, gross_pay: 113708, net_pay: 109938 },
      mpHeaders,
      {
        headers,
        sanitizeValue: (v) => String(v ?? '').trim(),
        flattenPayrollEarningColumns: (r) => r,
      }
    );
    expect(row[otherDed]).toBe('3562');
    expect(row['PT (22)']).toBe('208');
  });

  it('remaps PF from Other Deductions parent key to PF (22) on download', () => {
    const otherDed = 'Other Deductions Like EPF/ ESI/ Welfare Fund etc. (if any)';
    const oldHeaders = [otherDed, 'ESIC (22)', 'PT (22)', 'LWF (22)'];
    const newHeaders = ['PF (22)', 'ESIC (22)', 'PT (22)', 'LWF (22)'];
    const remapped = remapMPCombinedRegisterRows(
      [{ [otherDed]: '3562', 'PT (22)': '208', 'ESIC (22)': '', 'LWF (22)': '' }],
      oldHeaders,
      newHeaders
    );
    expect(remapped[0]['PF (22)']).toBe('3562');
    expect(remapped[0]['PT (22)']).toBe('208');
  });

  it('reads Sample Payroll PF including employer_pf fallback', () => {
    expect(resolveFormXVIIIMPSamplePayrollPf({ PF: 3562 })).toBe(3562);
    expect(resolveFormXVIIIMPSamplePayrollPf({ epf_contribution: 3562 })).toBe(3562);
    expect(resolveFormXVIIIMPSamplePayrollPf({ employer_pf: 3562 })).toBe(3562);
  });

  it('prefers Sample Payroll PF / Professional Tax over Zoho deduction line items', () => {
    const headers = ['PF (22)', 'PT (22)'];
    const mpHeaders = resolveFormXVIIIMPTableHeaders(headers);
    const row = {};
    headers.forEach((h) => {
      row[h] = '';
    });
    applyFormXVIIIMPPayrollToRow(
      row,
      {
        PF: 9211,
        'Professional Tax': 200,
        deductions: [
          { type: 'pf', name: 'PF', amount: 999 },
          { type: 'pt', name: 'Professional Tax', amount: 50 },
        ],
      },
      mpHeaders,
      {
        headers,
        sanitizeValue: (v) => String(v ?? '').trim(),
        flattenPayrollEarningColumns: (r) => r,
        getDeductionsArray: (p) => p?.deductions || [],
        findDeductionAmount: (list, pred) => {
          for (const item of list || []) {
            const n = String(item?.name || '').toLowerCase();
            const t = String(item?.type || '').toLowerCase();
            if (pred(t, n) && item?.amount != null) return item.amount;
          }
          return '';
        },
      }
    );
    expect(row['PF (22)']).toBe('9211');
    expect(row['PT (22)']).toBe('200');
  });

  it('PF column does not receive Professional Tax', () => {
    const headers = ['PF (22)', 'PT (22)'];
    const mpHeaders = resolveFormXVIIIMPTableHeaders(headers);
    const row = {};
    headers.forEach((h) => {
      row[h] = '';
    });
    applyFormXVIIIMPPayrollToRow(
      row,
      { pf: 1800, professional_tax: 200 },
      mpHeaders,
      {
        headers,
        sanitizeValue: (v) => String(v ?? '').trim(),
        flattenPayrollEarningColumns: (r) => r,
      }
    );
    expect(row['PF (22)']).toBe('1800');
    expect(row['PF (22)']).not.toBe('200');
    expect(row['PT (22)']).toBe('200');
  });

  it('reads basic/hra from flattened payroll for other allowances', () => {
    const headers = ['Other allowances', 'Total/ gross Wages/ Earnings', 'Wage rate/ pay'];
    const mpHeaders = resolveFormXVIIIMPTableHeaders(headers);
    const row = {};
    headers.forEach((h) => {
      row[h] = '';
    });
    applyFormXVIIIMPPayrollToRow(
      row,
      {
        gross_pay: 100000,
        net_pay: 90000,
        earnings: [
          { name: 'Basic', amount: 40000 },
          { name: 'HRA', amount: 20000 },
        ],
      },
      mpHeaders,
      {
        headers,
        sanitizeValue: (v) => String(v ?? '').trim(),
        findDeductionAmount: () => '',
        findEarningAmount: (earnings, pred) => {
          for (const item of earnings || []) {
            const n = String(item?.name || '').toLowerCase();
            const t = String(item?.type || '').toLowerCase();
            const amt = item?.amount;
            if (pred(t, n) && amt != null) return amt;
          }
          return '';
        },
        getEarningsArray: (p) => p?.earnings || [],
        getDeductionsArray: () => [],
        flattenPayrollEarningColumns: (r) => r,
      }
    );
    expect(row['Other allowances']).toBe('40000');
    expect(row['Total/ gross Wages/ Earnings']).toBe('100000');
  });

  it('finalizes Other allowances on download when mirrored from wage rate', () => {
    const headers = [
      'Other allowances',
      'Total/ gross Wages/ Earnings',
      'Wage rate/ pay or (piece rate/ wages per unit)',
    ];
    const rows = [
      {
        'Other allowances': '74132',
        'Total/ gross Wages/ Earnings': '74132',
        'Wage rate/ pay or (piece rate/ wages per unit)': '74132',
      },
    ];
    finalizeFormXVIIIMPOtherAllowancesForDownload(rows, headers, [{}], {
      sanitizeValue: (v) => String(v ?? '').trim(),
      payrollRows: [
        {
          gross_pay: 74132,
          net_pay: 70486,
          earnings: [
            { name: 'Basic', amount: 40000 },
            { name: 'HRA', amount: 10000 },
          ],
        },
      ],
    });
    expect(rows[0]['Other allowances']).toBe('24132');
    expect(rows[0]['Wage rate/ pay or (piece rate/ wages per unit)']).toBe('74132');
  });

  it('computes Other allowances on download path without autofill helpers', () => {
    const headers = [
      'Other allowances',
      'Total/ gross Wages/ Earnings',
      'Wage rate/ pay or (piece rate/ wages per unit)',
    ];
    const mpHeaders = resolveFormXVIIIMPTableHeaders(headers);
    const row = {};
    headers.forEach((h) => {
      row[h] = '';
    });
    applyFormXVIIIMPPayrollToRow(
      row,
      {
        gross_pay: 74132,
        net_pay: 70486,
        earnings: [
          { name: 'Basic', amount: 40000 },
          { name: 'HRA', amount: 10000 },
        ],
      },
      mpHeaders,
      {
        headers,
        sanitizeValue: (v) => String(v ?? '').trim(),
        findDeductionAmount: () => '',
        getDeductionsArray: () => [],
      }
    );
    expect(row['Other allowances']).toBe('24132');
    expect(row['Wage rate/ pay or (piece rate/ wages per unit)']).toBe('74132');
    expect(row['Total/ gross Wages/ Earnings']).toBe('74132');
  });

  it('detects Over time hours / overtime wages / NIL columns', () => {
    expect(
      isFormXVIIIMPOvertimeHoursHeader('Over time worked (Number of hours in the month)')
    ).toBe(true);
    expect(isFormXVIIIMPOvertimeWagesHeader('Amount of overtime wages')).toBe(true);
    expect(isFormXVIIIMPMaternityBenefitHeader('Amount of Maternity benefit (If any)')).toBe(true);
    expect(isFormXVIIIMPAnyOtherAmountHeader('Any other Amount (Please mention)')).toBe(true);
    expect(isFormXVIIIMPNilDefaultHeader('Amount of advances/ loans, if any and purpose of advance')).toBe(
      true
    );
    expect(isFormXVIIIMPNilDefaultHeader('Deductions of Fines imposed, if any')).toBe(true);
    expect(isFormXVIIIMPNilDefaultHeader('Other allowances')).toBe(false);
  });

  it('writes NIL defaults onto OT / maternity / advances / fines columns', () => {
    const headers = [
      'Over time worked (Number of hours in the month)',
      'Amount of overtime wages',
      'Amount of Maternity benefit (If any)',
      'Any other Amount (Please mention)',
      'Amount of advances/ loans, if any and purpose of advance',
      'Deductions of Fines imposed, if any',
      'Other allowances',
    ];
    const row = {
      'Over time worked (Number of hours in the month)': 'Enter Over t',
      'Amount of overtime wages': '',
      'Amount of Maternity benefit (If any)': '',
      'Any other Amount (Please mention)': '',
      'Amount of advances/ loans, if any and purpose of advance': '',
      'Deductions of Fines imposed, if any': '',
      'Other allowances': '123',
    };
    applyFormXVIIIMPNilDefaultsToRow(row, headers);
    expect(row['Over time worked (Number of hours in the month)']).toBe(FORM_XVIII_MP_NIL);
    expect(row['Amount of overtime wages']).toBe(FORM_XVIII_MP_NIL);
    expect(row['Amount of Maternity benefit (If any)']).toBe(FORM_XVIII_MP_NIL);
    expect(row['Any other Amount (Please mention)']).toBe(FORM_XVIII_MP_NIL);
    expect(row['Amount of advances/ loans, if any and purpose of advance']).toBe(FORM_XVIII_MP_NIL);
    expect(row['Deductions of Fines imposed, if any']).toBe(FORM_XVIII_MP_NIL);
    expect(row['Other allowances']).toBe('123');
  });

  it('applies payroll other allowances + leave category + NIL, not days as leave category', () => {
    const headers = [
      'Other allowances',
      'Category of Leave',
      'Total no. Of days worked',
      'Wage rate/ pay or (piece rate/ wages per unit)',
      'Over time worked (Number of hours in the month)',
      'Amount of overtime wages',
    ];
    const mpHeaders = resolveFormXVIIIMPTableHeaders(headers);
    const row = {};
    headers.forEach((h) => {
      row[h] = '';
    });
    applyFormXVIIIMPPayrollToRow(
      row,
      { gross_pay: 100000, basic: 40000, hra: 20000, paid_days: 30, net_pay: 90000 },
      mpHeaders,
      {
        headers,
        sanitizeValue: (v) => String(v ?? '').trim(),
        findDeductionAmount: () => '',
        getDeductionsArray: () => [],
      }
    );
    expect(row['Other allowances']).toBe('40000');
    expect(row['Category of Leave']).toBe(FORM_XVIII_MP_LEAVE_CATEGORY);
    expect(row['Total no. Of days worked']).toBe('30');
    expect(row['Over time worked (Number of hours in the month)']).toBe(FORM_XVIII_MP_NIL);
    expect(row['Amount of overtime wages']).toBe(FORM_XVIII_MP_NIL);
  });

  it('sums booked/balance for Earned + Legacy Earned + Contingency only', () => {
    const leaveRecord = {
      employee: { name: 'Test' },
      'Earned Leave': { paidBalance: 20, paidBooked: 2 },
      'Legacy Earned Leave': { balance: 10, booked: 1 },
      'Contingency Leave': { balance: 5, booked: 3 },
      'Paternity Leave': { balance: 15, booked: 4 },
    };
    const { booked, balance } = sumFormXVIIIMPLeaveBookedAndBalance(leaveRecord);
    expect(booked).toBe('6');
    expect(balance).toBe('35');
  });

  it('applies leave API metrics and fixed category text', () => {
    const headers = [
      'Category of Leave',
      'Leaves availed (No. Of days)',
      'Total Balance Leaves',
    ];
    const mpHeaders = resolveFormXVIIIMPTableHeaders(headers);
    const row = {
      'Category of Leave': '30',
      'Leaves availed (No. Of days)': 'Enter Leave',
      'Total Balance Leaves': '',
    };
    applyFormXVIIIMPLeaveToRow(
      row,
      {
        'Earned Leave': { paidBalance: 64, paidBooked: 12 },
        'Contingency Leave': { balance: 0, booked: 0 },
      },
      mpHeaders,
      { headers, sanitizeValue: (v) => String(v ?? '').trim() }
    );
    expect(row['Category of Leave']).toBe(FORM_XVIII_MP_LEAVE_CATEGORY);
    expect(row['Leaves availed (No. Of days)']).toBe('12');
    expect(row['Total Balance Leaves']).toBe('64');
  });

  it('formats FOR THE MONTH from selected month/year (replaces template dashes)', () => {
    expect(isFormXVIIIMPForTheMonthCellText('FOR THE MONTH ------------')).toBe(true);
    expect(extractFormXVIIIMPMonthYearFromForTheMonthText('FOR THE MONTH ------------')).toBe('');
    expect(extractFormXVIIIMPMonthYearFromForTheMonthText('FOR THE MONTH August 2026')).toBe(
      'August 2026'
    );
    expect(formatFormXVIIIMPForTheMonthDisplay('August 2026')).toBe('FOR THE MONTH August 2026');
    expect(formatFormXVIIIMPForTheMonthDisplay('FOR THE MONTH ------------')).toBe(
      FORM_XVIII_MP_FOR_THE_MONTH_LABEL
    );

    const fields = ensureFormXVIIIMPForTheMonthHeaderField([]);
    expect(fields).toEqual([
      {
        label: FORM_XVIII_MP_FOR_THE_MONTH_LABEL,
        value: '',
        key: FORM_XVIII_MP_FOR_THE_MONTH_KEY,
      },
    ]);

    const headerData = applyFormXVIIIMPForTheMonthToHeaderData(
      { [FORM_XVIII_MP_FOR_THE_MONTH_KEY]: 'FOR THE MONTH ------------' },
      'August 2026',
      fields
    );
    expect(headerData[FORM_XVIII_MP_FOR_THE_MONTH_KEY]).toBe('August 2026');
    expect(resolveFormXVIIIMPForTheMonthFromHeaderData(headerData, fields)).toBe('August 2026');
  });

  it('writes FOR THE MONTH August 2026 onto the worksheet merge cell', () => {
    const cellStore = new Map();
    const worksheet = {
      model: { merges: ['S2:Z2'] },
      getCell(rOrAddr, c) {
        let r;
        let col;
        if (typeof rOrAddr === 'string') {
          const m = String(rOrAddr).match(/^([A-Z]+)(\d+)$/i);
          if (!m) return { value: null, alignment: {} };
          col = m[1].toUpperCase().charCodeAt(0) - 64;
          if (m[1].length > 1) {
            col = 0;
            for (let i = 0; i < m[1].length; i += 1) {
              col = col * 26 + (m[1].toUpperCase().charCodeAt(i) - 64);
            }
          }
          r = Number(m[2]);
        } else {
          r = rOrAddr;
          col = c;
        }
        const key = `${r}:${col}`;
        if (!cellStore.has(key)) {
          cellStore.set(key, {
            row: r,
            col,
            value: r === 2 && col === 19 ? 'FOR THE MONTH ------------' : null,
            alignment: {},
          });
        }
        return cellStore.get(key);
      },
    };
    expect(writeFormXVIIIMPForTheMonthToWorksheet(worksheet, 'August 2026', { headerRowEnd: 5 })).toBe(
      true
    );
    expect(cellStore.get('2:19').value).toBe('FOR THE MONTH August 2026');
  });
});
