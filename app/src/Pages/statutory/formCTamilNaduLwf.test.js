import {
  applyFormCLabourWelfareYearToFormHeader,
  applyFormCLabourWelfareYearToHeaderLabel,
  applyFormCLabourWelfareYearToHeaders,
  applyFormCLabourWelfareYearToSubtitle,
  remapFormCLabourWelfareRowsToHeaders,
  resolveFormCLabourWelfareYear,
} from './formCTamilNaduLwf';

describe('Form C Tamil Nadu LWF year helpers', () => {
  it('resolves year from due date, then selected month, then current year', () => {
    expect(resolveFormCLabourWelfareYear('April', { dueDate: '2024-06-15' })).toBe(2024);
    expect(resolveFormCLabourWelfareYear('April 2025', {})).toBe(2025);
    expect(resolveFormCLabourWelfareYear('', {}, 'for the year - 2023')).toBe(2023);
  });

  it('appends corresponding year to subtitle', () => {
    expect(applyFormCLabourWelfareYearToSubtitle('', 2026)).toBe(
      'Register of Fines and Unpaid Accumulations for the year - 2026'
    );
    expect(
      applyFormCLabourWelfareYearToSubtitle(
        'Register of Fines and Unpaid Accumulations for the year',
        2024
      )
    ).toBe('Register of Fines and Unpaid Accumulations for the year - 2024');
    expect(
      applyFormCLabourWelfareYearToSubtitle(
        'Register of Fines and Unpaid Accumulations for the year - 2014',
        2026
      )
    ).toBe('Register of Fines and Unpaid Accumulations for the year - 2026');
  });

  it('rewrites truncated quarter header years', () => {
    expect(applyFormCLabourWelfareYearToHeaderLabel('Quarter ending 31-March-14 (2)', 2026)).toBe(
      'Quarter ending 31-March-2026 (2)'
    );
    expect(applyFormCLabourWelfareYearToHeaderLabel('Quarter ending 30-June-14 (3)', 2024)).toBe(
      'Quarter ending 30-June-2024 (3)'
    );
    expect(
      applyFormCLabourWelfareYearToHeaderLabel('Quarter ending 30-September-2014 (4)', 2026)
    ).toBe('Quarter ending 30-September-2026 (4)');
    expect(applyFormCLabourWelfareYearToHeaderLabel('Details of fines (1)', 2026)).toBe(
      'Details of fines (1)'
    );
  });

  it('rewrites all quarter headers in a list', () => {
    const headers = [
      'Details of fines and unpaid accumulations (1)',
      'Quarter ending 31-March-14 (2)',
      'Quarter ending 30-June-14 (3)',
      'Quarter ending 30-September-14 (4)',
      'Quarter ending 31-December-14 (5)',
    ];
    expect(applyFormCLabourWelfareYearToHeaders(headers, 2026)).toEqual([
      'Details of fines and unpaid accumulations (1)',
      'Quarter ending 31-March-2026 (2)',
      'Quarter ending 30-June-2026 (3)',
      'Quarter ending 30-September-2026 (4)',
      'Quarter ending 31-December-2026 (5)',
    ]);
  });

  it('updates form header subtitle with year', () => {
    const next = applyFormCLabourWelfareYearToFormHeader(
      { title: 'FORM C', subtitle: 'Register of Fines and Unpaid Accumulations for the year' },
      2026
    );
    expect(next.subtitle).toBe('Register of Fines and Unpaid Accumulations for the year - 2026');
  });

  it('remaps row keys when quarter headers change', () => {
    const oldHeaders = [
      'Details of fines (1)',
      'Quarter ending 31-March-14 (2)',
      'Quarter ending 30-June-14 (3)',
    ];
    const newHeaders = applyFormCLabourWelfareYearToHeaders(oldHeaders, 2026);
    const rows = [
      {
        'Details of fines (1)': '1. Total Realisations',
        'Quarter ending 31-March-14 (2)': 'Nil',
        'Quarter ending 30-June-14 (3)': '',
      },
    ];
    const remapped = remapFormCLabourWelfareRowsToHeaders(rows, oldHeaders, newHeaders);
    expect(remapped[0]['Quarter ending 31-March-2026 (2)']).toBe('Nil');
    expect(remapped[0]['Details of fines (1)']).toBe('1. Total Realisations');
    expect(remapped[0]['Quarter ending 31-March-14 (2)']).toBeUndefined();
  });
});
