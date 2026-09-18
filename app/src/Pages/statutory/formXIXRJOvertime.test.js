import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import {
  FORM_XIX_RJ_DEFAULT_TABLE_START_COL0,
  FORM_XIX_RJ_TABLE_HEADERS,
  applyFormXIXRJOriginalTemplateLayoutToExcelJs,
  applyFormXIXRJOvertimeAutofillToRow,
  buildFormXIXRJExportColMap,
  isFormXIXRJOvertimeAutofillContext,
  mapFormXIXRJHeadersToSheetJsCols,
  mapFormXIXRJHeadersToTemplateCols,
  prepareFormXIXRJOvertimeExportRows,
  resolveFormXIXRJNormalHours,
  resolveFormXIXRJTableHeaders,
  writeFormXIXRJHeaderValuesOntoOriginalLabels,
} from './formXIXRJOvertime';
import { formatWorkmanNameAndGuardian, isFormXIXRajasthanOvertimeRegisterContext } from './formXIXAPWageSlip';

const gridRead = (grid) => (r, c) =>
  grid[r] && grid[r][c] != null ? String(grid[r][c]) : '';

describe('Form XIX RJ Normal Hours', () => {
  test('uses paid_days × 8 when present', () => {
    expect(resolveFormXIXRJNormalHours({ paid_days: 31 })).toBe('248');
    expect(resolveFormXIXRJNormalHours({ paid_days: 1 })).toBe('8');
  });

  test('writes FirstName + MiddleName + LastName into Name of workman', () => {
    const row = {};
    applyFormXIXRJOvertimeAutofillToRow(
      row,
      { FirstName: 'Prem', MiddleName: 'Shankar', LastName: 'Menaria' },
      FORM_XIX_RJ_TABLE_HEADERS,
      { paid_days: 26, gross_pay: 18000, net_pay: 15000 }
    );
    expect(row['Name of workman']).toBe('Prem Shankar Menaria');
    expect(formatWorkmanNameAndGuardian({ FirstName: 'Ram', MiddleName: 'Kumar', LastName: 'Sharma' })).toBe(
      'Ram Kumar Sharma'
    );
  });

  test('detects Form_XIX_MH filename', () => {
    expect(
      isFormXIXRajasthanOvertimeRegisterContext(
        { title: 'FORM XIX', subtitle: 'Register of Overtime' },
        { formFileName: 'Form_XIX_MH_-_Maharashtra.xlsx', formName: 'Register of Overtime' },
        'Form_XIX_MH_-_Maharashtra.xlsx',
        ''
      )
    ).toBe(true);
  });

  test('stays blank when absent or unpaid (never 08:00)', () => {
    expect(resolveFormXIXRJNormalHours({ paid_days: 0 })).toBe('');
    expect(resolveFormXIXRJNormalHours({})).toBe('');
    expect(resolveFormXIXRJNormalHours(null)).toBe('');
  });

  test('export strips clock-style 08:00 leftovers from Normal Hours', () => {
    const rows = [
      {
        'Serial No.': '1',
        'Name of workman': 'Ram Lal',
        'Normal Hours': '08:00',
        'Normal rate': '12000',
      },
      {
        'Serial No.': '2',
        'Name of workman': 'Sita Devi',
        'Normal Hours': '240',
        'Normal rate': '15000',
      },
    ];
    const out = prepareFormXIXRJOvertimeExportRows(rows, FORM_XIX_RJ_TABLE_HEADERS);
    expect(out[0]['Normal Hours']).toBe('');
    expect(out[1]['Normal Hours']).toBe('240');
  });
});

describe('Form XIX RJ original template column map', () => {
  test('does not treat Form XXIII overtime files as Form XIX RJ', () => {
    expect(
      isFormXIXRajasthanOvertimeRegisterContext(
        { title: 'FORM XXIII', subtitle: 'Register of Overtime' },
        { formFileName: 'Form_XXIII_-_TamilNadu.xlsx', formName: 'Register of Overtime' },
        'Form_XXIII_-_TamilNadu.xlsx',
        ''
      )
    ).toBe(false);
    expect(
      isFormXIXRJOvertimeAutofillContext(
        { title: 'FORM XXIII', subtitle: 'Register of Overtime' },
        { formFileName: 'Form_XXIII_GJ_-_Gujarat.xlsx', formName: 'Register of Overtime' },
        'Form_XXIII_GJ_-_Gujarat.xlsx',
        '',
        FORM_XIX_RJ_TABLE_HEADERS
      )
    ).toBe(false);
  });

  test('detects Form_XIX_RJ filename', () => {
    expect(
      isFormXIXRajasthanOvertimeRegisterContext(
        { title: 'FORM XIX', subtitle: 'Register of Overtime' },
        { formFileName: 'Form_XIX_RJ_-_Rajasthan.xlsx', formName: 'Register of Overtime' },
        'Form_XIX_RJ_-_Rajasthan.xlsx',
        ''
      )
    ).toBe(true);
  });

  test('fallback export map starts at column B, never A', () => {
    const mapped = buildFormXIXRJExportColMap(0);
    expect(mapped.tableStartCol).toBe(FORM_XIX_RJ_DEFAULT_TABLE_START_COL0);
    expect(mapped.headerToCol.get('Serial No.')).toBe(1);
    expect(mapped.headerToCol.get('Name of workman')).toBe(2);
    expect(mapped.headerToCol.get('Date on which overtime payment made')).toBe(15);
  });

  test('blank column A spacer maps Serial No. to column B', () => {
    const grid = [
      ['FORM XIX', '', ''],
      ['Register of Overtime', '', ''],
      [
        '',
        'Serial No.',
        'Name of workman',
        "Father's/Husband's name",
        'Sex',
        'Designation and department',
        'Date on which overtime work was put in',
        'Wages of overtime on each occasion',
        'Total over time worked or production in case of piece rates',
        'Normal Hours',
        'Normal rate',
        'Overtime rate',
        'Normal earning',
        'Overtime earnings',
        'Total earnings',
        'Date on which overtime payment made',
      ],
      ['', '1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12', '13', '14', '15'],
      ['', '1', 'Ram Lal', 'Sohan Lal', 'M', 'Civil / Helper', 'Nil', 'Nil', '', '208', '12000', 'Nil', '11000', 'Nil', '11000', 'Nil'],
    ];
    const layout = mapFormXIXRJHeadersToTemplateCols(gridRead(grid));
    expect(layout.tableStartCol).toBe(1);
    expect(layout.headerToCol.get('Serial No.')).toBe(1);
    expect(layout.headerToCol.get('Name of workman')).toBe(2);
    expect(layout.headerToCol.get('Sex')).toBe(4);
    expect(layout.headerToCol.get('Normal Hours')).toBe(9);
    expect(layout.dataStartRow).toBe(4);
  });

  test('Serial stays beside Name — no empty gap column B', () => {
    const grid = [
      [
        'Serial No.',
        'Serial No.',
        'Name of workman',
        "Father's/Husband's name",
        'Sex',
      ],
    ];
    const layout = mapFormXIXRJHeadersToTemplateCols(gridRead(grid));
    expect(layout.headerToCol.get('Serial No.')).toBe(1);
    expect(layout.headerToCol.get('Name of workman')).toBe(2);
    expect(layout.headerToCol.get("Father's/Husband's name")).toBe(3);
  });

  test('number row 1..15 starting at B wins over merge-aware Serial in A', () => {
    const headers = [
      'Serial No.',
      'Name of workman',
      "Father's/Husband's name",
      'Sex',
      'Designation and department',
      'Date on which overtime work was put in',
      'Wages of overtime on each occasion',
      'Total over time worked or production in case of piece rates',
      'Normal Hours',
      'Normal rate',
      'Overtime rate',
      'Normal earning',
      'Overtime earnings',
      'Total earnings',
      'Date on which overtime payment made',
    ];
    const numbers = ['', '1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12', '13', '14', '15'];
    const mergeAwareHeaders = ['Serial No.', ...headers];
    const grid = [mergeAwareHeaders, numbers];
    const layout = mapFormXIXRJHeadersToTemplateCols(gridRead(grid));
    expect(layout.headerToCol.get('Serial No.')).toBe(1);
    expect(layout.headerToCol.get('Name of workman')).toBe(2);
    expect(layout.tableStartCol).toBe(1);
    expect(layout.headerToCol.get('Serial No.') + 1).toBe(layout.headerToCol.get('Name of workman'));
  });

  test('SheetJS empty-A template maps Serial to column B', () => {
    const aoa = [
      ['FORM XIX', '', 'Register of Overtime'],
      [
        '',
        'Serial No.',
        'Name of workman',
        "Father's/Husband's name",
        'Sex',
        'Designation and department',
      ],
      ['', '1', 'Sita Devi', 'Ram', 'F', 'Housekeeping / Attendant'],
    ];
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    const layout = mapFormXIXRJHeadersToSheetJsCols(ws, FORM_XIX_RJ_TABLE_HEADERS);
    expect(layout.headerToCol.get('Serial No.')).toBe(1);
    expect(layout.headerToCol.get('Name of workman')).toBe(2);
    expect(layout.tableStartCol).toBe(1);
  });

  test('export rows stay on canonical headers for template write', () => {
    const rows = [
      {
        '': '1',
        'Serial No.': 'Ram Lal',
        'Name of workman': 'Sohan Lal',
      },
    ];
    const headers = ['', 'Serial No.', 'Name of workman', "Father's/Husband's name"];
    const out = prepareFormXIXRJOvertimeExportRows(rows, headers);
    const canon = resolveFormXIXRJTableHeaders(headers);
    expect(canon[0]).toBe('Serial No.');
    expect(out[0]['Serial No.']).toBe('1');
    expect(out[0]['Name of workman']).toBe('Ram Lal');
  });

  test('Form_XIX_RJ filename wins even when company text mentions another state', () => {
    expect(
      isFormXIXRajasthanOvertimeRegisterContext(
        { title: 'FORM XIX', subtitle: 'Register of Overtime' },
        {
          formFileName: 'Form_XIX_RJ_-_Rajasthan.xlsx',
          formName: 'Register of Overtime',
          state: 'Rajasthan',
        },
        'Form_XIX_RJ_-_Rajasthan.xlsx',
        'Registered office Gujarat'
      )
    ).toBe(true);
    expect(
      isFormXIXRJOvertimeAutofillContext(
        { title: 'FORM XIX' },
        { formFileName: 'Form_XIX_RJ_-_Rajasthan.xlsx' },
        'Form_XIX_RJ_-_Rajasthan.xlsx',
        'Andhra Pradesh head office',
        FORM_XIX_RJ_TABLE_HEADERS
      )
    ).toBe(true);
  });

  test('ExcelJS restore moves Serial from A into B and sets vertical headers', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Sheet1');
    ws.getCell('F2').value = 'FORM XIX';
    ws.getCell('F3').value = '[See Rule 77 (2)(e)]';
    ws.getCell('F4').value = 'Register of Overtime';
    ws.getCell('A6').value = 'Name and address of the Contractor........................................................................';
    ws.getCell('A13').value = 'Serial No.';
    ws.getCell('C13').value = 'Name of workman';
    ws.getCell('D13').value = "Father's/Husband's name";
    ws.getCell('A14').value = '1';
    ws.getCell('C14').value = '2';
    ws.getCell('D14').value = '3';
    ws.getCell('A15').value = '1';
    ws.getCell('C15').value = 'Ram Lal';
    const layout = applyFormXIXRJOriginalTemplateLayoutToExcelJs(ws);
    expect(layout).toBeTruthy();
    expect(String(ws.getCell('A13').value || '')).toBe('');
    expect(String(ws.getCell('B13').value)).toBe('Serial No.');
    expect(String(ws.getCell('C13').value)).toBe('Name of workman');
    expect(ws.getCell('B13').alignment?.textRotation).toBe(90);
    expect(ws.getCell('C13').alignment?.textRotation).toBe(90);
    expect(Number(ws.getRow(13).height)).toBeGreaterThanOrEqual(78);
    expect(layout.tableStartCol1).toBe(2);
  });

  test('keeps original dotted contractor label and writes the value after the dots', () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Sheet1');
    ws.getCell('B6').value =
      'Name and address of the Contractor........................................................................';
    writeFormXIXRJHeaderValuesOntoOriginalLabels(
      ws,
      { form_xxiii_contractor: 'Vayona Energy' },
      { fields: [{ key: 'form_xxiii_contractor', label: 'Name and address of the Contractor' }] },
      12
    );
    const text = String(ws.getCell('B6').value);
    expect(text).toMatch(/Name and address of the Contractor/);
    expect(text).toMatch(/\.{5,}/);
    expect(text).toContain('Vayona Energy');
    expect(text).not.toMatch(/^Name and address of contractor\s*:/i);
  });

  test('does not rotate Form XXIII overtime headers', () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Sheet1');
    ws.getCell('A1').value = 'FORM XXIII';
    ws.getCell('A2').value = 'Register of Overtime';
    ws.getCell('A10').value = 'Serial No.';
    ws.getCell('B10').value = 'Name of workman';
    expect(applyFormXIXRJOriginalTemplateLayoutToExcelJs(ws)).toBeNull();
    expect(ws.getCell('A10').alignment?.textRotation || 0).not.toBe(90);
  });
});
