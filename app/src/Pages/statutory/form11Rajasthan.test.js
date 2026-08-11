import {
  FORM_11_RJ_TABLE_HEADERS,
  classifyForm11RJHeader,
  headersIndicateForm11RajasthanEmployment,
  isForm11RajasthanContext,
  remapForm11RajasthanRowsToHeaders,
  repairForm11RajasthanTableHeadersFromWorkbook,
  resolveForm11RajasthanTableHeaders,
  resolveForm11RajasthanYoungPersonValue,
  sanitizeForm11RajasthanCellValue,
} from './form11Rajasthan';

describe('form11Rajasthan', () => {
  test('detects Form_11_RJ filename context', () => {
    expect(
      isForm11RajasthanContext({}, { formFileName: 'Form_11_RJ_-_Rajasthan.xlsx' }, '', '', [])
    ).toBe(true);
  });

  test('detects Rajasthan register of employment headers', () => {
    expect(headersIndicateForm11RajasthanEmployment(FORM_11_RJ_TABLE_HEADERS)).toBe(true);
  });

  test('does not treat ESIC accident book as RJ Form 11', () => {
    expect(
      isForm11RajasthanContext(
        { title: 'FORM 11', subtitle: 'Accident Book' },
        { formFileName: 'Form_11_TamilNadu.xlsx' },
        '',
        'employees state insurance accident book',
        ['Date of Notice', 'Insurance No.', 'Cause of Injury']
      )
    ).toBe(false);
  });

  test('repairs stacked Days of Month headers to 12 leaf columns', () => {
    const broken = [
      'Name of Persons Employeed',
      'Whether young person or not',
      'Days of Month Time at which employment commences',
      'Days of Month Time at which employment ceases',
      'Rest interval',
      'Total hours worked during the month',
      '*Days on which overtime work is done and extent of such overtime on each day',
      'Extent of overtime worked during the month',
      'Extent of overtime worked during the quarter',
      'Extent of overtime worked during the year',
    ];
    const resolved = resolveForm11RajasthanTableHeaders(broken);
    expect(resolved).toHaveLength(12);
    expect(classifyForm11RJHeader(resolved[2])).toBe('commences');
    expect(classifyForm11RJHeader(resolved[3])).toBe('ceases');
    expect(classifyForm11RJHeader(resolved[4])).toBe('day1');
    expect(classifyForm11RJHeader(resolved[5])).toBe('rest');
    expect(classifyForm11RJHeader(resolved[6])).toBe('day3');
  });

  test('sanitizes HTML help text from young-person cells', () => {
    expect(sanitizeForm11RajasthanCellValue('<div>Please follow the steps below</div>')).toBe('');
    expect(sanitizeForm11RajasthanCellValue('No')).toBe('No');
  });

  test('remaps rows and clears junk young-person values', () => {
    const srcHdrs = [
      'Name of Persons Employeed',
      'Whether young person or not',
      'Days of Month Time at which employment commences',
    ];
    const rows = [
      {
        'Name of Persons Employeed': 'Virendra',
        'Whether young person or not': '<div>Please follow the steps belc',
        'Days of Month Time at which employment commences': '09:00',
      },
    ];
    const out = remapForm11RajasthanRowsToHeaders(rows, srcHdrs, FORM_11_RJ_TABLE_HEADERS);
    expect(out[0][FORM_11_RJ_TABLE_HEADERS[0]]).toBe('Virendra');
    expect(out[0][FORM_11_RJ_TABLE_HEADERS[1]]).toBe('');
    expect(out[0][FORM_11_RJ_TABLE_HEADERS[2]]).toBe('09:00');
  });

  test('young person Yes/No from age', () => {
    expect(resolveForm11RajasthanYoungPersonValue(17)).toBe('Yes');
    expect(resolveForm11RajasthanYoungPersonValue(25)).toBe('No');
    expect(resolveForm11RajasthanYoungPersonValue('')).toBe('No');
  });

  test('repairs headers from Form 11_RJ.xlsx when available', () => {
    let XLSX;
    let fs;
    try {
      XLSX = require('xlsx');
      fs = require('fs');
    } catch {
      return;
    }
    const candidate = 'C:/Users/HP/Downloads/Form 11_RJ.xlsx';
    if (!fs.existsSync(candidate)) return;
    const wb = XLSX.readFile(candidate);
    const repaired = repairForm11RajasthanTableHeadersFromWorkbook(wb);
    expect(repaired).not.toBeNull();
    expect(repaired.headers).toHaveLength(12);
    expect(repaired.dataStartIndex).toBeGreaterThan(repaired.headerRowIndex);
    expect(classifyForm11RJHeader(repaired.headers[0])).toBe('name');
    expect(classifyForm11RJHeader(repaired.headers[5])).toBe('rest');
  });
});
