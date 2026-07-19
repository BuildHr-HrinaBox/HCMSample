import {
  FORM_XVIII_MP_LEAVE_CATEGORY,
  FORM_XVIII_MP_NIL,
  applyFormXVIIIMPLeaveToRow,
  applyFormXVIIIMPNilDefaultsToRow,
  applyFormXVIIIMPPayrollToRow,
  computeFormXVIIIMPOtherAllowances,
  finalizeFormXVIIIMPOtherAllowancesForDownload,
  isFormXVIIIMPAnyOtherAmountHeader,
  isFormXVIIIMPCombinedRegisterContext,
  isFormXVIIIMPMaternityBenefitHeader,
  isFormXVIIIMPNilDefaultHeader,
  isFormXVIIIMPOvertimeHoursHeader,
  isFormXVIIIMPOvertimeWagesHeader,
  isFormXVIIIMPWageRateHeader,
  looksLikeKarnatakaFormTSheet,
  looksLikeMPCombinedRegisterSheet,
  readFormXVIIIMPOtherAllowanceFromRow,
  remapMPCombinedRegisterRows,
  repairFormXVIIIMPDuplicateWageHeaders,
  resolveFormXVIIIMPOtherAllowanceForWrite,
  resolveFormXVIIIMPSectionHeaderLabel,
  resolveFormXVIIIMPTableHeaders,
  restoreFormXVIIIMPDistinctOtherAllowances,
  snapshotFormXVIIIMPDistinctOtherAllowances,
  sumFormXVIIIMPLeaveBookedAndBalance,
} from './formXVIIIMPCombinedRegister';

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
  });

  it('still detects Madhya Pradesh Form XVIII combined register', () => {
    const mpBlob =
      'Form XVIII Madhya Pradesh Muster Roll-cum Register of Wages Name of the establishment and address Location of work';
    expect(looksLikeKarnatakaFormTSheet(mpBlob)).toBe(false);
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
});
