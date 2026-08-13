import {
  FORM_12_RJ_CEASES_DEFAULT,
  FORM_12_RJ_COMMENCES_DEFAULT,
  FORM_12_RJ_OT_DEFAULT,
  FORM_12_RJ_TABLE_HEADERS,
  FORM_12_RJ_YOUNG_DEFAULT,
  applyForm12RajasthanAutofillDefaultsToRows,
  classifyForm12RJHeader,
  clearForm12RajasthanAttendanceDayColumns,
  headersIndicateForm12RajasthanEmployment,
  isForm12RajasthanContext,
  looksLikeForm12RajasthanFileName,
  remapForm12RajasthanRowsToHeaders,
  repairForm12RajasthanTableHeadersFromWorkbook,
  resolveForm12RajasthanTableHeaders,
} from './form12Rajasthan';
import { isForm11RajasthanContext } from './form11Rajasthan';
import * as XLSX from 'xlsx';
import fs from 'fs';

describe('form12Rajasthan', () => {
  test('detects Form_12_RJ filename context', () => {
    expect(looksLikeForm12RajasthanFileName('Form_12_RJ_-_Rajasthan.xlsx')).toBe(true);
    expect(
      isForm12RajasthanContext({}, { formFileName: 'Form_12_RJ_-_Rajasthan.xlsx' }, '', '', [])
    ).toBe(true);
  });

  test('does not treat Form 12 as Form 11', () => {
    expect(
      isForm11RajasthanContext({}, { formFileName: 'Form_12_RJ_-_Rajasthan.xlsx' }, '', '', [])
    ).toBe(false);
    expect(
      isForm11RajasthanContext(
        { title: 'FORM 12', subtitle: 'Where opening and closing hours are ordinarily uniform' },
        { formFileName: 'Form_12_RJ.xlsx', siteState: 'Rajasthan' },
        '',
        'opening and closing hours are ordinarily uniform',
        FORM_12_RJ_TABLE_HEADERS
      )
    ).toBe(false);
  });

  test('detects Form 12 employment headers', () => {
    expect(headersIndicateForm12RajasthanEmployment(FORM_12_RJ_TABLE_HEADERS)).toBe(true);
  });

  test('resolves canonical 19 leaf headers', () => {
    const resolved = resolveForm12RajasthanTableHeaders(['broken']);
    expect(resolved).toHaveLength(19);
    expect(classifyForm12RJHeader(resolved[0])).toBe('name');
    expect(classifyForm12RJHeader(resolved[4])).toBe('rest');
    expect(classifyForm12RJHeader(resolved[5])).toBe('hours1');
    expect(classifyForm12RJHeader(resolved[14])).toBe('hours10');
    expect(classifyForm12RJHeader(resolved[18])).toBe('otPrevYear');
  });

  test('autofill defaults write Young / times / OT NIL and leave hours 1–10 blank', () => {
    const headers = [...FORM_12_RJ_TABLE_HEADERS];
    const rows = [
      {
        [headers[0]]: 'Virendra',
        [headers[5]]: '9:00',
        [headers[6]]: 'A',
        [headers[7]]: 'WO',
      },
    ];
    applyForm12RajasthanAutofillDefaultsToRows(rows, headers);
    expect(rows[0][headers[1]]).toBe(FORM_12_RJ_YOUNG_DEFAULT);
    expect(rows[0][headers[2]]).toBe(FORM_12_RJ_COMMENCES_DEFAULT);
    expect(rows[0][headers[3]]).toBe(FORM_12_RJ_CEASES_DEFAULT);
    expect(rows[0][headers[4]]).toBe('');
    expect(rows[0][headers[5]]).toBe('');
    expect(rows[0][headers[6]]).toBe('');
    expect(rows[0][headers[7]]).toBe('');
    expect(rows[0][headers[14]]).toBe('');
    expect(rows[0][headers[16]]).toBe(FORM_12_RJ_OT_DEFAULT);
    expect(rows[0][headers[17]]).toBe(FORM_12_RJ_OT_DEFAULT);
    expect(rows[0][headers[18]]).toBe(FORM_12_RJ_OT_DEFAULT);
  });

  test('clearForm12RajasthanAttendanceDayColumns blanks hours 1–10', () => {
    const headers = [...FORM_12_RJ_TABLE_HEADERS];
    const rows = [{ [headers[5]]: 'P', [headers[8]]: '10:33', [headers[0]]: 'Atul' }];
    clearForm12RajasthanAttendanceDayColumns(rows, headers);
    expect(rows[0][headers[0]]).toBe('Atul');
    expect(rows[0][headers[5]]).toBe('');
    expect(rows[0][headers[8]]).toBe('');
  });

  test('remap keeps person name on name column', () => {
    const srcHdrs = FORM_12_RJ_TABLE_HEADERS;
    const rows = [{ [srcHdrs[0]]: 'Atul Dhakad', [srcHdrs[1]]: 'Young' }];
    const out = remapForm12RajasthanRowsToHeaders(rows, srcHdrs, FORM_12_RJ_TABLE_HEADERS);
    expect(out[0][FORM_12_RJ_TABLE_HEADERS[0]]).toBe('Atul Dhakad');
  });

  test('repairs headers from Form 12_RJ.xlsx when available', () => {
    const candidate = 'C:/Users/HP/Downloads/Form 12_RJ.xlsx';
    if (!fs.existsSync(candidate)) return;
    const wb = XLSX.readFile(candidate);
    const repaired = repairForm12RajasthanTableHeadersFromWorkbook(wb, {
      fileName: 'Form_12_RJ.xlsx',
    });
    expect(repaired).toBeTruthy();
    expect(repaired.headers).toHaveLength(19);
    expect(classifyForm12RJHeader(repaired.headers[0])).toBe('name');
    expect(classifyForm12RJHeader(repaired.headers[5])).toBe('hours1');
    expect(repaired.tableStartCol).toBe(1); // column B
  });
});
