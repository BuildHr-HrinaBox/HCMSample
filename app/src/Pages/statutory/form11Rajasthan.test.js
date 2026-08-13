import {
  FORM_11_RJ_CEASES_DEFAULT,
  FORM_11_RJ_COMMENCES_DEFAULT,
  FORM_11_RJ_OT_DEFAULT,
  FORM_11_RJ_REST_AFTER_DEFAULT,
  FORM_11_RJ_REST_BEFORE_DEFAULT,
  FORM_11_RJ_TABLE_HEADERS,
  FORM_11_RJ_YOUNG_DEFAULT,
  applyForm11RajasthanAutofillDefaultsToRows,
  applyForm11RajasthanPaidDaysToRows,
  buildForm11RajasthanExportColMap,
  classifyForm11RJHeader,
  computeForm11RajasthanTotalHours,
  exportForm11RajasthanRowValues,
  getForm11RajasthanRowValueForHeader,
  headersIndicateForm11RajasthanEmployment,
  isForm11RajasthanContext,
  isForm11RajasthanJunkTotalHoursValue,
  looksLikeForm11RajasthanPersonName,
  prepareForm11RajasthanExport,
  remapForm11RajasthanRowsToHeaders,
  repairForm11RajasthanExportRow,
  repairForm11RajasthanTableHeadersFromWorkbook,
  resolveForm11RajasthanTableHeaders,
  resolveForm11RajasthanYoungPersonValue,
  sanitizeForm11RajasthanCellValue,
  sanitizeForm11RajasthanTotalHoursValue,
  sheetJsWorkbookLooksLikeForm11Rajasthan,
  repairForm11RajasthanExcelJsWorksheet,
  writeForm11RajasthanMonthYearToSheetJs,
} from './form11Rajasthan';

describe('form11Rajasthan', () => {
  test('detects Form_11_RJ filename context', () => {
    expect(
      isForm11RajasthanContext({}, { formFileName: 'Form_11_RJ_-_Rajasthan.xlsx' }, '', '', [])
    ).toBe(true);
  });

  test('does not treat Form_12_RJ as Form 11', () => {
    expect(
      isForm11RajasthanContext({}, { formFileName: 'Form_12_RJ_-_Rajasthan.xlsx' }, '', '', [])
    ).toBe(false);
  });

  test('does not treat Form_14_RJ as Form 11', () => {
    expect(
      isForm11RajasthanContext({}, { formFileName: 'Form_14_RJ_-_Rajasthan.xlsx' }, '', '', [])
    ).toBe(false);
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
    // HTML junk is dropped; export repair fills the Young default.
    expect(out[0][FORM_11_RJ_TABLE_HEADERS[1]]).toBe(FORM_11_RJ_YOUNG_DEFAULT);
    expect(out[0][FORM_11_RJ_TABLE_HEADERS[2]]).toBe('09:00');
  });

  test('autofill defaults: Young / 9 AM / 5PM / rest before-after / OT NIL', () => {
    expect(resolveForm11RajasthanYoungPersonValue(25)).toBe(FORM_11_RJ_YOUNG_DEFAULT);
    expect(FORM_11_RJ_YOUNG_DEFAULT).toBe('Young');
    expect(FORM_11_RJ_COMMENCES_DEFAULT).toBe('9AM');
    expect(FORM_11_RJ_CEASES_DEFAULT).toBe('5PM');
    expect(FORM_11_RJ_REST_BEFORE_DEFAULT).toBe('9AM');
    expect(FORM_11_RJ_REST_AFTER_DEFAULT).toBe('5PM');
    expect(FORM_11_RJ_OT_DEFAULT).toBe('NIL');
  });

  test('total hours = Paid_days × 8; blank when missing or junk', () => {
    expect(computeForm11RajasthanTotalHours(26)).toBe('208');
    expect(computeForm11RajasthanTotalHours('22.5')).toBe('180');
    expect(computeForm11RajasthanTotalHours('')).toBe('');
    expect(computeForm11RajasthanTotalHours('1 month(s)')).toBe('');
    expect(isForm11RajasthanJunkTotalHoursValue('1 month(s)')).toBe(true);
    expect(isForm11RajasthanJunkTotalHoursValue('5PM')).toBe(true);
    expect(isForm11RajasthanJunkTotalHoursValue('9AM')).toBe(true);
    expect(isForm11RajasthanJunkTotalHoursValue('NIL')).toBe(true);
    expect(sanitizeForm11RajasthanTotalHoursValue('1 month(s)')).toBe('');
    expect(sanitizeForm11RajasthanTotalHoursValue('5PM')).toBe('');
    expect(sanitizeForm11RajasthanTotalHoursValue('184')).toBe('184');
  });

  test('autofill blanks leaked 5PM from Total hours when no Paid_days', () => {
    const headers = [...FORM_11_RJ_TABLE_HEADERS];
    const rows = [
      {
        [headers[0]]: 'Virendra',
        [headers[7]]: '5PM',
      },
    ];
    applyForm11RajasthanAutofillDefaultsToRows(rows, headers);
    expect(rows[0][headers[7]]).toBe('');
    expect(rows[0][headers[3]]).toBe('5PM');
  });

  test('applyForm11RajasthanPaidDaysToRows fills total hours from payroll', () => {
    const headers = [...FORM_11_RJ_TABLE_HEADERS];
    const rows = [{ [headers[0]]: 'Virendra', [headers[7]]: '' }];
    const hits = applyForm11RajasthanPaidDaysToRows(rows, [{ EmployeeID: '1' }], headers, {
      overwrite: true,
      resolvePayrollRow: () => ({ paid_days: 20 }),
    });
    expect(hits).toBe(1);
    expect(rows[0][headers[7]]).toBe('160');
  });

  test('applyForm11RajasthanPaidDaysToRows clears 1 month(s) when Paid_days missing', () => {
    const headers = [...FORM_11_RJ_TABLE_HEADERS];
    const rows = [{ [headers[0]]: 'Atul', [headers[7]]: '1 month(s)' }];
    const hits = applyForm11RajasthanPaidDaysToRows(rows, [{ EmployeeID: '2' }], headers, {
      overwrite: true,
      resolvePayrollRow: () => null,
    });
    expect(hits).toBe(1);
    expect(rows[0][headers[7]]).toBe('');
  });

  test('applyForm11RajasthanAutofillDefaultsToRows writes OT NIL on all overtime columns', () => {
    const headers = [...FORM_11_RJ_TABLE_HEADERS];
    const rows = [{ [headers[0]]: 'Virendra' }];
    applyForm11RajasthanAutofillDefaultsToRows(rows, headers);
    expect(rows[0][headers[1]]).toBe('Young');
    expect(rows[0][headers[2]]).toBe('9AM');
    expect(rows[0][headers[3]]).toBe('5PM');
    expect(rows[0][headers[4]]).toBe('9AM');
    expect(rows[0][headers[6]]).toBe('5PM');
    expect(rows[0][headers[8]]).toBe('NIL');
    expect(rows[0][headers[9]]).toBe('NIL');
    expect(rows[0][headers[10]]).toBe('NIL');
    expect(rows[0][headers[11]]).toBe('NIL');
  });

  test('getForm11RajasthanRowValueForHeader finds OT NIL by role when key differs', () => {
    const row = {
      '*Days on which overtime work is done and extent of such overtime on each day': 'NIL',
    };
    expect(
      getForm11RajasthanRowValueForHeader(
        row,
        FORM_11_RJ_TABLE_HEADERS[8]
      )
    ).toBe('NIL');
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

  test('export col map stays sequential from Name column (B=1)', () => {
    const { headerToCol, headers } = buildForm11RajasthanExportColMap(1);
    expect(headers).toEqual(FORM_11_RJ_TABLE_HEADERS);
    expect(headerToCol.get(FORM_11_RJ_TABLE_HEADERS[0])).toBe(1); // B
    expect(headerToCol.get(FORM_11_RJ_TABLE_HEADERS[3])).toBe(4); // E ceases
    expect(headerToCol.get(FORM_11_RJ_TABLE_HEADERS[11])).toBe(12); // M
  });

  test('export values keep name under Name even when Object.keys reorders 1/3', () => {
    const row = {
      [FORM_11_RJ_TABLE_HEADERS[4]]: '9AM',
      [FORM_11_RJ_TABLE_HEADERS[6]]: '5PM',
      [FORM_11_RJ_TABLE_HEADERS[0]]: 'virendra',
      [FORM_11_RJ_TABLE_HEADERS[1]]: 'Young',
      [FORM_11_RJ_TABLE_HEADERS[2]]: '9 AM',
      [FORM_11_RJ_TABLE_HEADERS[3]]: '5PM',
      [FORM_11_RJ_TABLE_HEADERS[8]]: 'NIL',
      EmployeeID: '4',
    };
    // Integer keys "1"/"3" sort first — must not drive export order.
    expect(Object.keys(row)[0]).toBe('1');
    const values = exportForm11RajasthanRowValues(row);
    expect(values[0]).toBe('virendra');
    expect(values[1]).toBe('Young');
    expect(values[3]).toBe('5PM');
    expect(values[4]).toBe('9AM');
  });

  test('repairs shifted row: EmployeeID in Name, real name parked under Ceases', () => {
    const shifted = {
      [FORM_11_RJ_TABLE_HEADERS[0]]: '4',
      [FORM_11_RJ_TABLE_HEADERS[1]]: 'NIL',
      [FORM_11_RJ_TABLE_HEADERS[2]]: 'NIL',
      [FORM_11_RJ_TABLE_HEADERS[3]]: 'virendra',
      [FORM_11_RJ_TABLE_HEADERS[8]]: 'NIL',
      [FORM_11_RJ_TABLE_HEADERS[9]]: 'NIL',
      EmployeeID: '4',
    };
    expect(looksLikeForm11RajasthanPersonName('virendra')).toBe(true);
    expect(looksLikeForm11RajasthanPersonName('4')).toBe(false);
    const fixed = repairForm11RajasthanExportRow(shifted);
    expect(fixed[FORM_11_RJ_TABLE_HEADERS[0]]).toBe('virendra');
    expect(fixed[FORM_11_RJ_TABLE_HEADERS[1]]).toBe(FORM_11_RJ_YOUNG_DEFAULT);
    expect(fixed[FORM_11_RJ_TABLE_HEADERS[2]]).toBe(FORM_11_RJ_COMMENCES_DEFAULT);
    expect(fixed[FORM_11_RJ_TABLE_HEADERS[3]]).toBe(FORM_11_RJ_CEASES_DEFAULT);
    expect(
      getForm11RajasthanRowValueForHeader(shifted, FORM_11_RJ_TABLE_HEADERS[0])
    ).toBe('virendra');
    const exported = exportForm11RajasthanRowValues(shifted);
    expect(exported[0]).toBe('virendra');
  });

  test('defaults recover name from Ceases before writing 5PM (SampleData shift)', () => {
    const shifted = {
      [FORM_11_RJ_TABLE_HEADERS[0]]: '4',
      [FORM_11_RJ_TABLE_HEADERS[1]]: 'NIL',
      [FORM_11_RJ_TABLE_HEADERS[2]]: 'NIL',
      [FORM_11_RJ_TABLE_HEADERS[3]]: 'virendra',
      [FORM_11_RJ_TABLE_HEADERS[8]]: 'NIL',
      EmployeeID: '4',
    };
    const rows = [shifted];
    applyForm11RajasthanAutofillDefaultsToRows(rows, FORM_11_RJ_TABLE_HEADERS);
    expect(rows[0][FORM_11_RJ_TABLE_HEADERS[0]]).toBe('virendra');
    expect(rows[0][FORM_11_RJ_TABLE_HEADERS[1]]).toBe(FORM_11_RJ_YOUNG_DEFAULT);
    expect(rows[0][FORM_11_RJ_TABLE_HEADERS[2]]).toBe(FORM_11_RJ_COMMENCES_DEFAULT);
    expect(rows[0][FORM_11_RJ_TABLE_HEADERS[3]]).toBe(FORM_11_RJ_CEASES_DEFAULT);
    expect(rows[0][FORM_11_RJ_TABLE_HEADERS[4]]).toBe(FORM_11_RJ_REST_BEFORE_DEFAULT);
    expect(rows[0][FORM_11_RJ_TABLE_HEADERS[6]]).toBe(FORM_11_RJ_REST_AFTER_DEFAULT);
    expect(rows[0].EmployeeID).toBe('4');
    expect(exportForm11RajasthanRowValues(rows[0])[0]).toBe('virendra');
  });

  test('defaults recover name from Employeed typo key', () => {
    const rows = [
      {
        'Name of Persons Employeed': '4',
        'Whether young person or not': 'NIL',
        'Time at which employment commences': 'NIL',
        'Time at which employment ceases': 'sharwan kumar',
        EmployeeID: '7',
      },
    ];
    applyForm11RajasthanAutofillDefaultsToRows(rows, FORM_11_RJ_TABLE_HEADERS);
    expect(rows[0][FORM_11_RJ_TABLE_HEADERS[0]]).toBe('sharwan kumar');
    expect(rows[0][FORM_11_RJ_TABLE_HEADERS[3]]).toBe(FORM_11_RJ_CEASES_DEFAULT);
    expect(Object.prototype.hasOwnProperty.call(rows[0], 'Name of Persons Employeed')).toBe(false);
  });

  test('prepareForm11RajasthanExport never uses tableStartCol 0 (Name is column B)', () => {
    const rows = [
      {
        [FORM_11_RJ_TABLE_HEADERS[0]]: '4',
        [FORM_11_RJ_TABLE_HEADERS[1]]: 'NIL',
        [FORM_11_RJ_TABLE_HEADERS[2]]: 'NIL',
        [FORM_11_RJ_TABLE_HEADERS[3]]: 'virendra',
        EmployeeID: '4',
      },
    ];
    const prepared = prepareForm11RajasthanExport(null, rows, FORM_11_RJ_TABLE_HEADERS, {
      tableStartCol: 0,
    });
    expect(prepared.tableStartCol).toBe(1);
    expect(prepared.headerToCol.get(FORM_11_RJ_TABLE_HEADERS[0])).toBe(1);
    expect(prepared.headerToCol.get(FORM_11_RJ_TABLE_HEADERS[3])).toBe(4);
    expect(prepared.rows[0][FORM_11_RJ_TABLE_HEADERS[0]]).toBe('virendra');
    expect(exportForm11RajasthanRowValues(prepared.rows[0])[0]).toBe('virendra');
    expect(exportForm11RajasthanRowValues(prepared.rows[0])[3]).toBe(FORM_11_RJ_CEASES_DEFAULT);
  });

  test('writes Month/Year into Form 11 RJ underline placeholders', () => {
    let XLSX;
    try {
      XLSX = require('xlsx');
    } catch {
      return;
    }
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet([
      [],
      [],
      [],
      [],
      [],
      [],
      ['', 'Month––––––––––––––––', '', '', '', '', '', 'Year––––––––––'],
    ]);
    wb.Sheets.Sheet1 = ws;
    wb.SheetNames = ['Sheet1'];
    writeForm11RajasthanMonthYearToSheetJs(ws, {
      form_11_rj_month: 'August',
      form_11_rj_year: '2026',
    });
    expect(String(ws.B7.v)).toBe('Month : August');
    expect(String(ws.H7.v)).toBe('Year : 2026');
  });

  test('detects Form 11 RJ workbook from title band', () => {
    let XLSX;
    try {
      XLSX = require('xlsx');
    } catch {
      return;
    }
    const ws = XLSX.utils.aoa_to_sheet([
      [],
      ['The Rajasthan Shops and Establishments Rules, 1959'],
      ['FORM 11'],
      ['Rule 22 - sub rule 1'],
      ['Register of Employment'],
      [],
      ['', 'Month––––', '', '', '', '', '', 'Year––––'],
      [],
      ['', 'Name of Persons Employed', 'Whether young person or not'],
    ]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    expect(sheetJsWorkbookLooksLikeForm11Rajasthan(wb)).toBe(true);
  });

  test('ExcelJS repair restores %PM to 5PM as text', async () => {
    let ExcelJS;
    try {
      ExcelJS = require('exceljs');
    } catch {
      return;
    }
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Sheet1');
    ws.getCell('A3').value = 'FORM 11';
    ws.getCell('A5').value = 'Register of Employment';
    ws.getCell('E15').value = '%PM';
    ws.getCell('E15').numFmt = '0%';
    ws.getCell('D15').value = '9AM';
    expect(repairForm11RajasthanExcelJsWorksheet(ws)).toBe(true);
    expect(String(ws.getCell('E15').value)).toBe('5PM');
    expect(ws.getCell('E15').numFmt).toBe('@');
    expect(String(ws.getCell('D15').value)).toBe('9AM');
  });

  test('buildForm11RajasthanWorkbookWithTemplateStyles writes names into column B boxes', async () => {
    let ExcelJS;
    try {
      ExcelJS = require('exceljs');
    } catch {
      return;
    }
    const {
      buildForm11RajasthanWorkbookWithTemplateStyles,
    } = require('./form11Rajasthan');
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Sheet1');
    ws.getCell('A2').value = 'The Rajasthan Shops and Establishments Rules, 1959';
    ws.getCell('A3').value = 'FORM 11';
    ws.getCell('A4').value = 'Rule 22 - sub rule 1';
    ws.getCell('A5').value = 'Register of Employment';
    ws.getCell('B7').value = 'Month----------';
    ws.getCell('H7').value = 'Year----------';
    ws.getCell('B10').value = 'Name of Persons Employed';
    ws.getCell('C10').value = 'Whether young person or not';
    ws.getCell('D10').value = 'Days of Month';
    ws.mergeCells('D10:E10');
    ws.getCell('F10').value = '1';
    ws.getCell('G10').value = '2';
    ws.getCell('H10').value = '3';
    ws.getCell('D11').value = 'Time at which employment commences';
    ws.getCell('E11').value = 'Time at which employment ceases';
    ws.getCell('G11').value = 'Rest interval';
    ws.getCell('I10').value = 'Total hours worked during the month';
    ws.mergeCells('B10:B11');
    ws.mergeCells('C10:C11');
    ws.mergeCells('I10:I11');
    for (let i = 0; i < 12; i += 1) {
      ws.getCell(12, 2 + i).value = String(i + 1);
    }
    // Fake vertical merge box over first data rows (template quirk).
    ws.mergeCells('B13:B14');
    const buf = await wb.xlsx.writeBuffer();
    const { blob } = await buildForm11RajasthanWorkbookWithTemplateStyles({
      templateArrayBuffer: buf,
      mappedData: [
        {
          'Name of Persons Employed': '4',
          'Whether young person or not': 'NIL',
          'Time at which employment ceases': 'virendra',
        },
        {
          EmployeeID: '5',
          'Name of Persons Employed': '5',
          'Time at which employment ceases': 'sharwan kumar',
        },
      ],
      headersToUse: FORM_11_RJ_TABLE_HEADERS,
      headerFormData: { form_11_rj_month: 'August', form_11_rj_year: '2026' },
      parsedTableStartCol: 1,
      formFileName: 'Form_11_RJ_-_Rajasthan.xlsx',
    });
    const outWb = new ExcelJS.Workbook();
    await outWb.xlsx.load(await new Response(blob).arrayBuffer());
    const outWs = outWb.worksheets[0];
    expect(String(outWs.getCell('B7').value)).toMatch(/Month\s*:\s*August/i);
    expect(String(outWs.getCell('H7').value)).toMatch(/Year\s*:\s*2026/i);
    // 1st-image model: Days of Month over D–E; day numbers 1/2/3 on top; Rest under day 2.
    expect(String(outWs.getCell('D10').value || '')).toMatch(/days\s+of\s+month/i);
    expect(String(outWs.getCell('F10').value || '').trim()).toBe('1');
    expect(String(outWs.getCell('G10').value || '').trim()).toBe('2');
    expect(String(outWs.getCell('H10').value || '').trim()).toBe('3');
    expect(String(outWs.getCell('D11').value || '')).toMatch(/commences/i);
    expect(String(outWs.getCell('E11').value || '')).toMatch(/ceases/i);
    expect(String(outWs.getCell('F11').value || '').trim()).toBe('');
    expect(String(outWs.getCell('G11').value || '')).toMatch(/rest\s+interval/i);
    expect(String(outWs.getCell('H11').value || '').trim()).toBe('');
    expect(String(outWs.getCell('B13').value)).toBe('virendra');
    expect(String(outWs.getCell('C13').value)).toBe('Young');
    expect(String(outWs.getCell('D13').value)).toBe('9AM');
    expect(String(outWs.getCell('E13').value)).toBe('5PM');
    expect(String(outWs.getCell('B14').value)).toBe('sharwan kumar');
    // Column A must stay empty (no EmployeeID dump).
    expect(String(outWs.getCell('A13').value || '').trim()).toBe('');
  });

  test('prepareForm11RajasthanExport remaps Employeed rows onto canonical B…M map', () => {
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
    const srcHdrs = [
      'Name of Persons Employeed',
      'Whether young person or not',
      'Time at which employment commences',
      'Time at which employment ceases',
      '1',
      '2 Rest interval',
      '3',
      'Total hours worked during the month',
      '*Days on which overtime work is done and extent of such overtime on each day',
      'Extent of overtime worked during the month',
      'Extent of overtime worked during the quarter',
      'Extent of overtime worked during the year',
    ];
    const rows = [
      {
        'Name of Persons Employeed': 'virendra',
        'Whether young person or not': 'Young',
        'Time at which employment commences': '9AM',
        'Time at which employment ceases': '5PM',
        1: '9AM',
        '2 Rest interval': '',
        3: '5PM',
        'Total hours worked during the month': '160',
        '*Days on which overtime work is done and extent of such overtime on each day': 'NIL',
        'Extent of overtime worked during the month': 'NIL',
        'Extent of overtime worked during the quarter': 'NIL',
        'Extent of overtime worked during the year': 'NIL',
        EmployeeID: '4',
      },
    ];
    const prepared = prepareForm11RajasthanExport(wb, rows, srcHdrs);
    expect(prepared.headers).toEqual(FORM_11_RJ_TABLE_HEADERS);
    expect(prepared.tableStartCol).toBe(1);
    expect(prepared.headerToCol.get(FORM_11_RJ_TABLE_HEADERS[0])).toBe(1);
    expect(prepared.rows[0][FORM_11_RJ_TABLE_HEADERS[0]]).toBe('virendra');
    expect(exportForm11RajasthanRowValues(prepared.rows[0])[0]).toBe('virendra');
  });
});
