import ExcelJS from 'exceljs';
import * as XLSX from 'xlsx';
import {
  bindOriginalTemplateLayoutPreserve,
  snapshotExcelJsWorkbookLayout,
  restoreExcelJsWorkbookLayout,
  validateExcelJsWorkbookLayout,
  snapshotSheetJsWorkbookLayout,
  preserveSheetJsOriginalTemplateLayout,
} from './excelTemplateLayoutPreserve';

describe('excel original template layout preserve', () => {
  test('ExcelJS write restores original column widths, row heights, wrap and merges', async () => {
    const template = new ExcelJS.Workbook();
    const ws = template.addWorksheet('FormSample');
    ws.getColumn(1).width = 8;
    ws.getColumn(2).width = 22;
    ws.getColumn(3).width = 11;
    ws.getRow(1).height = 21;
    ws.getRow(2).height = 36;
    ws.mergeCells('A1:C1');
    ws.getCell(1, 1).value = 'FORM XVIII';
    ws.getCell(1, 1).alignment = { horizontal: 'center', wrapText: false, vertical: 'middle' };
    ws.getCell(2, 1).value = 'Daily attendance / units of work done';
    ws.getCell(2, 1).alignment = { wrapText: true, horizontal: 'center', vertical: 'middle' };
    ws.getCell(3, 1).value = 'sample';
    const buf = await template.xlsx.writeBuffer();

    const loaded = new ExcelJS.Workbook();
    await loaded.xlsx.load(buf);
    const sheet = loaded.worksheets[0];
    sheet.getColumn(1).width = 40;
    sheet.getColumn(2).width = 40;
    sheet.getRow(1).height = 80;
    sheet.getCell(1, 1).alignment = { horizontal: 'left', wrapText: true };
    sheet.getCell(3, 1).value = 'Selva P';
    const out = await loaded.xlsx.writeBuffer();

    const reloaded = new ExcelJS.Workbook();
    reloaded.__skipStatutoryLayoutPreserve = true;
    await reloaded.xlsx.load(out, { skipLayoutPreserve: true });
    const outWs = reloaded.worksheets[0];
    expect(outWs.getColumn(1).width).toBeCloseTo(8, 0);
    expect(outWs.getColumn(2).width).toBeCloseTo(22, 0);
    expect(outWs.getRow(1).height).toBeCloseTo(21, 0);
    expect(outWs.getRow(2).height).toBeCloseTo(36, 0);
    expect(outWs.getCell(1, 1).alignment?.horizontal).toBe('center');
    expect(outWs.getCell(1, 1).alignment?.wrapText).not.toBe(true);
    expect(String(outWs.getCell(3, 1).value)).toBe('Selva P');
    const merges = outWs.model?.merges || [];
    expect(merges.some((m) => /^A1:C1$/i.test(m))).toBe(true);
  });

  test('validation reports width drift then restore clears it', () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('X');
    ws.getColumn(1).width = 10;
    ws.getRow(1).height = 18;
    ws.getCell(1, 1).value = 'FORM W';
    const snap = snapshotExcelJsWorkbookLayout(wb);
    ws.getColumn(1).width = 50;
    const before = validateExcelJsWorkbookLayout(wb, snap);
    expect(before.ok).toBe(false);
    restoreExcelJsWorkbookLayout(wb, snap);
    expect(validateExcelJsWorkbookLayout(wb, snap).ok).toBe(true);
  });

  test('SheetJS restore keeps original !cols and cell styles while data values change', () => {
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet([
      ['FORM XVIII', '', ''],
      ['Name', 'Overtime', 'Net Amount'],
      ['sample', 'Nil', '']
    ]);
    ws['!cols'] = [{ wch: 9 }, { wch: 14 }, { wch: 16 }];
    ws['!rows'] = [{ hpt: 20 }, { hpt: 32 }, { hpt: 18 }];
    ws['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 2 } }];
    ws.A1.s = { alignment: { horizontal: 'center', wrapText: false } };
    XLSX.utils.book_append_sheet(wb, ws, 'FormXVIII');
    const snap = snapshotSheetJsWorkbookLayout(wb);
    ws['!cols'] = [{ wch: 40 }, { wch: 40 }, { wch: 40 }];
    ws['!rows'][1].hpt = 90;
    ws.C3 = { t: 'n', v: 67348 };
    const report = preserveSheetJsOriginalTemplateLayout(wb, snap);
    expect(report.ok).toBe(true);
    expect(ws['!cols'][0].wch).toBe(9);
    expect(ws['!cols'][2].wch).toBe(16);
    expect(ws['!rows'][1].hpt).toBe(32);
    expect(ws.C3.v).toBe(67348);
  });

  test('bind is idempotent', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('S');
    ws.getColumn(1).width = 12;
    const buf = await wb.xlsx.writeBuffer();
    const loaded = new ExcelJS.Workbook();
    await loaded.xlsx.load(buf);
    bindOriginalTemplateLayoutPreserve(loaded);
    bindOriginalTemplateLayoutPreserve(loaded);
    loaded.worksheets[0].getColumn(1).width = 99;
    const out = await loaded.xlsx.writeBuffer();
    const check = new ExcelJS.Workbook();
    check.__skipStatutoryLayoutPreserve = true;
    await check.xlsx.load(out, { skipLayoutPreserve: true });
    expect(check.worksheets[0].getColumn(1).width).toBeCloseTo(12, 0);
  });
});
