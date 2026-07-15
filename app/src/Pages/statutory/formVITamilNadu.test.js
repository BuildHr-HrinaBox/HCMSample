import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import {
  ensureFormVITamilNaduTitleLayout,
  isFormVITamilNaduHolidayContext,
  normalizeFormVITitleCellText,
  pickFormVITamilNaduExcelJsWorksheet,
  repickFormVITamilNaduWorkbookSheetIfNeeded,
  resolveFormVITamilNaduWorkbookSheetName,
  splitFormVIHeaderBlock
} from './formVITamilNadu';

function buildWorkbook() {
  const wb = XLSX.utils.book_new();

  const employeeSheet = XLSX.utils.aoa_to_sheet([
    ['S.No', 'Name of the employee', 'Employee Identification No.', 'Gender', 'Father / Spouse Name', 'Date of Birth', 'Date of Joining'],
    [1, 'Avudaiappan', 'VE0147', 'Male', 'Muthukrishnan M', '30-May-197', '01 Dec 2025']
  ]);
  XLSX.utils.book_append_sheet(wb, employeeSheet, 'Employee Details');

  const holidaySheet = XLSX.utils.aoa_to_sheet([
    ['FORM - VI'],
    ['REGISTER OF NATIONAL AND FESTIVAL HOLIDAYS'],
    ['S.No.', 'Employee Code', 'Name of the Employee', "Father's/Husband's Name", 'DOJ', '14-01-2026 (PONGAL)', '26/01/2026 (REPUBLIC DAY)', 'Remarks'],
    [1, 'VE0147', 'Avudaiappa', 'Muthukrishnan M', '01/01/2020', 'H', 'H', '']
  ]);
  XLSX.utils.book_append_sheet(wb, holidaySheet, 'Form VI');

  return wb;
}

function buildWorkbookWithRealisticEmployeeSheet() {
  const wb = XLSX.utils.book_new();
  const employeeSheet = XLSX.utils.aoa_to_sheet([
    ['S.No', 'Name of the employee', 'Employee Identification No.', 'Gender', 'Father / Spouse Name', 'Date of Birth', 'Date of Joining'],
    [1, 'Avudaiappan', 'VE0147', 'Male', 'Muthukrishnan M', '30-May-197', '01 Dec 2025']
  ]);
  XLSX.utils.book_append_sheet(wb, employeeSheet, 'Employee Details');

  const holidaySheet = XLSX.utils.aoa_to_sheet([
    ['FORM - VI'],
    ['REGISTER OF NATIONAL AND FESTIVAL HOLIDAYS'],
    ['S.No.', 'Employee Code', 'Name of the Employee', "Father's/Husband's Name", 'DOJ', '14-01-2026 (PONGAL)', '26/01/2026 (REPUBLIC DAY)', 'Remarks'],
    [1, 'VE0147', 'Avudaiappa', 'Muthukrishnan M', '01/01/2020', 'H', 'H', '']
  ]);
  XLSX.utils.book_append_sheet(wb, holidaySheet, 'Holiday Register');

  return wb;
}

describe('Form VI Tamil Nadu holiday workbook helpers', () => {
  it('detects the holiday workbook context from the file name', () => {
    expect(
      isFormVITamilNaduHolidayContext(
        null,
        { formName: 'Form VI' },
        'Form_VI_-_TamilNadu.xlsx'
      )
    ).toBe(true);

    expect(
      isFormVITamilNaduHolidayContext(
        null,
        { formName: 'Form 6' },
        'Form_6_-_AP.xlsx',
        'humidity register'
      )
    ).toBe(false);
  });

  it('prefers the holiday sheet over an employee-details sheet', () => {
    const wb = buildWorkbook();
    expect(
      resolveFormVITamilNaduWorkbookSheetName(wb, {
        fileName: 'Form_VI_-_TamilNadu.xlsx'
      })
    ).toBe('Form VI');
  });

  it('repicks the open sheet when the current sheet is wrong', () => {
    const wb = buildWorkbook();
    expect(
      repickFormVITamilNaduWorkbookSheetIfNeeded(
        wb,
        {
          fileName: 'Form_VI_-_TamilNadu.xlsx'
        },
        'Employee Details'
      )
    ).toBe('Form VI');
  });

  it('still chooses the holiday register when the other sheet looks like employee details', () => {
    const wb = buildWorkbookWithRealisticEmployeeSheet();
    expect(
      resolveFormVITamilNaduWorkbookSheetName(wb, {
        fileName: 'Form_VI_-_TamilNadu.xlsx'
      })
    ).toBe('Holiday Register');
  });

  it('splits the merged Form VI heading block', () => {
    const parsed = splitFormVIHeaderBlock(
      'FORM - VI\nREGISTER OF NATIONAL AND FESTIVAL HOLIDAYS\n[See sub-rule (1) under rule (7)]\nTHE TAMIL NADU INDUSTRIAL ESTABLISHMENTS (NATIONAL AND FESTIVAL HOLIDAYS) RULES, 1959'
    );
    expect(parsed).toEqual({
      title: 'FORM - VI',
      subtitle: 'REGISTER OF NATIONAL AND FESTIVAL HOLIDAYS',
      reference:
        '[See sub-rule (1) under rule (7)]\nTHE TAMIL NADU INDUSTRIAL ESTABLISHMENTS (NATIONAL AND FESTIVAL HOLIDAYS) RULES, 1959'
    });
  });

  it('normalizes SheetJS _x000d_ title artifacts', () => {
    expect(normalizeFormVITitleCellText('FORM - VI_x000d_REGISTER')).toBe('FORM - VI\nREGISTER');
  });

  it('restores a wide multi-row title merge when the header collapsed into column A', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Form VI');
    ws.getCell(1, 1).value =
      'FORM - VI\nREGISTER OF NATIONAL AND FESTIVAL HOLIDAYS\n[See sub-rule (1) under rule (7)]\nTHE TAMIL NADU INDUSTRIAL ESTABLISHMENTS (NATIONAL AND FESTIVAL HOLIDAYS) RULES, 1959';
    ws.getCell(6, 1).value = 'S.No.';
    ws.getCell(6, 19).value = 'Remarks';

    const changed = ensureFormVITamilNaduTitleLayout(ws, {
      colFrom: 1,
      colTo: 19,
      headerRow: 6
    });
    expect(changed).toBe(true);

    const merges = ws.model.merges || [];
    expect(merges.some((m) => /^A1:[A-Z]+4$/i.test(m) || /^A1:S4$/i.test(m))).toBe(true);
    expect(ws.getCell(1, 1).alignment?.horizontal).toBe('center');
    expect(ws.getCell(1, 1).alignment?.wrapText).toBe(true);
    expect(String(ws.getCell(1, 1).value)).toMatch(/FORM - VI/);
  });

  it('picks the Form VI worksheet from an ExcelJS workbook', async () => {
    const wb = new ExcelJS.Workbook();
    const other = wb.addWorksheet('Sheet2');
    other.getCell(1, 1).value = 'FEMALE';
    const holiday = wb.addWorksheet('Form VI');
    holiday.getCell(1, 1).value = 'FORM - VI\nREGISTER OF NATIONAL AND FESTIVAL HOLIDAYS';
    holiday.getCell(7, 6).value = '14-01-2026 (PONGAL)';

    const picked = pickFormVITamilNaduExcelJsWorksheet(wb, {
      fileName: 'Form_VI_-_TamilNadu.xlsx'
    });
    expect(picked?.name).toBe('Form VI');
  });

  it('builds a clean A-S workbook that keeps title merges and column widths', async () => {
    const fs = require('fs');
    const path = require('path');
    const templatePath = path.join(__dirname, '../../../Form_VI_-_TamilNadu.xlsx');
    if (!fs.existsSync(templatePath)) return;

    const { buildFormVITamilNaduWorkbookWithTemplateStyles } = require('./formVITamilNadu');
    const buf = fs.readFileSync(templatePath);
    const { arrayBuffer } = await buildFormVITamilNaduWorkbookWithTemplateStyles({
      templateArrayBuffer: buf,
      mappedData: [
        {
          'S.No.': 1,
          'Employee Code': 'VE0147',
          'Name of the Employee': 'Avudaiappan',
          "Father's /Husband's Name": 'Muthukrishnan M',
          DOJ: '01 Dec 2025',
          '14-01-2026 ( PONGAL)': 'H',
          '26/01/2026 (REPUBLIC DAY)': 'H'
        }
      ],
      headersToUse: [
        'S.No.',
        'Employee Code',
        'Name of the Employee',
        "Father's /Husband's Name",
        'DOJ',
        '14-01-2026 ( PONGAL)',
        '26/01/2026 (REPUBLIC DAY)',
        'Remarks'
      ],
      formFileName: 'Form_VI_-_TamilNadu.xlsx',
      ExcelJS
    });

    const outWb = new ExcelJS.Workbook();
    await outWb.xlsx.load(arrayBuffer);
    expect(outWb.worksheets.length).toBe(1);
    expect(outWb.worksheets[0].name).toBe('Form VI');
    const merges = outWb.worksheets[0].model.merges || [];
    expect(merges).toEqual(expect.arrayContaining(['A1:S4', 'F6:S6', 'A6:A8']));
    expect(outWb.worksheets[0].getColumn(3).width).toBeGreaterThan(20);
    expect(outWb.worksheets[0].getCell(1, 1).alignment?.horizontal).toBe('center');
    expect(outWb.worksheets[0].getCell(1, 1).alignment?.wrapText).toBe(true);
    expect(String(outWb.worksheets[0].getCell(9, 2).value)).toBe('VE0147');
  });
});
