import {
  STATUTORY_MONTH_ABBR,
  STATUTORY_MONTH_NAMES,
  daysInStatutoryMonth,
  resolveToFullMonthName,
  statutoryMonthDayCountsForYear,
  statutoryMonthIndex,
} from './statutoryMonthResolve';

describe('statutory month resolve — all 12 months (not May/July-only)', () => {
  test('full names, abbreviations, and Month YYYY all resolve uniquely', () => {
    STATUTORY_MONTH_NAMES.forEach((name, i) => {
      expect(resolveToFullMonthName(name)).toBe(name);
      expect(resolveToFullMonthName(name.toUpperCase())).toBe(name);
      expect(resolveToFullMonthName(STATUTORY_MONTH_ABBR[i])).toBe(name);
      expect(resolveToFullMonthName(`${name} 2026`)).toBe(name);
      expect(resolveToFullMonthName(`${STATUTORY_MONTH_ABBR[i]}-2026`)).toBe(name);
      expect(resolveToFullMonthName(`2026-${String(i + 1).padStart(2, '0')}`)).toBe(name);
      expect(statutoryMonthIndex(name)).toBe(i);
    });
  });

  test('July is never June; Jun stays June', () => {
    expect(resolveToFullMonthName('July')).toBe('July');
    expect(resolveToFullMonthName('july')).toBe('July');
    expect(resolveToFullMonthName('Jul')).toBe('July');
    expect(resolveToFullMonthName('JUL')).toBe('July');
    expect(resolveToFullMonthName('July 2026')).toBe('July');
    expect(resolveToFullMonthName('2026-07')).toBe('July');
    expect(resolveToFullMonthName('June')).toBe('June');
    expect(resolveToFullMonthName('Jun')).toBe('June');
    expect(resolveToFullMonthName('2026-06')).toBe('June');
    expect(statutoryMonthIndex('July')).not.toBe(statutoryMonthIndex('June'));
  });

  test('May is never March; Mar stays March', () => {
    expect(resolveToFullMonthName('May')).toBe('May');
    expect(resolveToFullMonthName('May 2026')).toBe('May');
    expect(resolveToFullMonthName('2026-05')).toBe('May');
    expect(resolveToFullMonthName('March')).toBe('March');
    expect(resolveToFullMonthName('Mar')).toBe('March');
    expect(resolveToFullMonthName('2026-03')).toBe('March');
  });

  test('unique prefixes do not steal a later month', () => {
    expect(resolveToFullMonthName('sept')).toBe('September');
    expect(resolveToFullMonthName('aug')).toBe('August');
    expect(resolveToFullMonthName('oct')).toBe('October');
    expect(resolveToFullMonthName('dec')).toBe('December');
    expect(resolveToFullMonthName('jan')).toBe('January');
  });

  test('2026 day counts match the calendar for every month', () => {
    const expected = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
    const got = statutoryMonthDayCountsForYear(2026);
    got.forEach((row, i) => {
      expect(row.month).toBe(STATUTORY_MONTH_NAMES[i]);
      expect(row.days).toBe(expected[i]);
      expect(daysInStatutoryMonth(STATUTORY_MONTH_ABBR[i], 2026)).toBe(expected[i]);
    });
    expect(daysInStatutoryMonth('July', 2026)).toBe(31);
    expect(daysInStatutoryMonth('June', 2026)).toBe(30);
    expect(daysInStatutoryMonth('May', 2026)).toBe(31);
    expect(daysInStatutoryMonth('February', 2024)).toBe(29);
  });

  test('wage-period style text uses the month token, not a 3-letter prefix collision', () => {
    expect(resolveToFullMonthName('Wage Period : July 2026')).toBe('July');
    expect(resolveToFullMonthName('For the month of May 2026')).toBe('May');
    expect(resolveToFullMonthName('1st March 2026 to 31st March 2026')).toBe('March');
    expect(resolveToFullMonthName('1st August 2026 to 31st August 2026')).toBe('August');
  });
});
