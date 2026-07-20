import {
  computeFormXVIIITamilNaduOtherCashPayments,
  isFormXVIIITamilNaduDailyAttendanceHeader,
  isFormXVIIITamilNaduDailyRateHeader,
  isFormXVIIITamilNaduOtherCashPaymentsHeader,
  isFormXVIIITamilNaduTotalAttendanceHeader,
} from './formXVIIITamilNaduWagesMuster';

describe('Form XVIII Tamil Nadu payroll column headers', () => {
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
});
