import {
  FORM_15_RJ_TABLE_HEADERS,
  collapseForm15RajasthanDuplicateHeaders,
  headersIndicateForm15RajasthanWeeklyHolidays,
  isForm15RajasthanContext,
  remapForm15RajasthanRowsToHeaders,
  resolveForm15RajasthanTableHeaders,
} from './form15Rajasthan';

describe('form15Rajasthan', () => {
  const duplicatedHeaders = [
    'Name of persons employed',
    'Name of persons employed',
    'Name of persons employed',
    'Designation / Department',
    'Designation / Department',
    'Designation / Department',
    'Day of the week',
    'Day of the week',
    'Day of the week',
  ];

  it('detects Form_15_RJ filename context', () => {
    expect(
      isForm15RajasthanContext(
        { title: 'FORM 15' },
        { formFileName: 'Form_15_RJ_-_Rajasthan.xlsx' },
        'Form_15_RJ_-_Rajasthan.xlsx'
      )
    ).toBe(true);
  });

  it('does not treat TN Form 15 leave register as RJ weekly holidays', () => {
    expect(
      isForm15RajasthanContext(
        { title: 'FORM 15', subtitle: 'Leave with Wages - Part 1' },
        { formFileName: 'Form_15_Part_1_-_TamilNadu.xlsx', siteState: 'Tamil Nadu' },
        'Form_15_Part_1_-_TamilNadu.xlsx'
      )
    ).toBe(false);
  });

  it('collapses merge-expanded duplicate column titles to one each', () => {
    expect(collapseForm15RajasthanDuplicateHeaders(duplicatedHeaders)).toEqual([
      'Name of persons employed',
      'Designation / Department',
      'Day of the week',
    ]);
  });

  it('resolves duplicated modal headers to three unique columns', () => {
    expect(resolveForm15RajasthanTableHeaders(duplicatedHeaders)).toEqual([
      'Name of persons employed',
      'Designation / Department',
      'Day of the week',
    ]);
    expect(headersIndicateForm15RajasthanWeeklyHolidays(duplicatedHeaders)).toBe(true);
  });

  it('falls back to canonical headers when empty', () => {
    expect(resolveForm15RajasthanTableHeaders([])).toEqual([...FORM_15_RJ_TABLE_HEADERS]);
  });

  it('remaps row values onto unique headers', () => {
    const rows = [
      {
        'Name of persons employed': 'Virendra',
        'Designation / Department': 'Junior Engineer',
        'Day of the week': 'Sunday',
      },
    ];
    const remapped = remapForm15RajasthanRowsToHeaders(rows, duplicatedHeaders, duplicatedHeaders);
    expect(remapped).toHaveLength(1);
    expect(remapped[0]['Name of persons employed']).toBe('Virendra');
    expect(remapped[0]['Designation / Department']).toBe('Junior Engineer');
    expect(remapped[0]['Day of the week']).toBe('Sunday');
    expect(Object.keys(remapped[0]).filter((k) => !k.startsWith('__'))).toEqual([
      'Name of persons employed',
      'Designation / Department',
      'Day of the week',
    ]);
  });
});
