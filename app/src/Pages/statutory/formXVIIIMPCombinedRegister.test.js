import {
  FORM_XVIII_MP_LEAVE_CATEGORY,
  FORM_XVIII_MP_NIL,
  applyFormXVIIIMPLeaveToRow,
  applyFormXVIIIMPNilDefaultsToRow,
  applyFormXVIIIMPPayrollToRow,
  computeFormXVIIIMPOtherAllowances,
  isFormXVIIIMPAnyOtherAmountHeader,
  isFormXVIIIMPCombinedRegisterContext,
  isFormXVIIIMPMaternityBenefitHeader,
  isFormXVIIIMPNilDefaultHeader,
  isFormXVIIIMPOvertimeHoursHeader,
  isFormXVIIIMPOvertimeWagesHeader,
  looksLikeKarnatakaFormTSheet,
  looksLikeMPCombinedRegisterSheet,
  resolveFormXVIIIMPTableHeaders,
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
