import ExcelJS from 'exceljs';
import {
  ensureExcelJSDataRowsWithBorders,
  excelJSCellHasBorder,
  trimExcelJsTableToRecordCount,
} from './excelTableBorders';

describe('excel data rows fit employee count', () => {
  test('5 employees keep 5 boxed rows and extra template boxes are cleared', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('FormXXVI');
    for (let r = 11; r <= 25; r += 1) {
      for (let c = 2; c <= 10; c += 1) {
        const cell = ws.getCell(r, c);
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        };
        if (r <= 15) cell.value = `sample-${r}-${c}`;
      }
    }
    for (let i = 0; i < 5; i += 1) {
      ws.getCell(11 + i, 2).value = i + 1;
      ws.getCell(11 + i, 3).value = `Emp ${i + 1}`;
    }
    ensureExcelJSDataRowsWithBorders(ws, {
      dataStartRow: 11,
      dataRowCount: 5,
      colFrom: 2,
      colTo: 10,
      templateRow: 11
    });

    for (let r = 11; r <= 15; r += 1) {
      expect(excelJSCellHasBorder(ws.getCell(r, 3))).toBe(true);
    }
    expect(String(ws.getCell(11, 3).value)).toBe('Emp 1');
    expect(excelJSCellHasBorder(ws.getCell(16, 3))).toBe(false);
    expect(ws.getCell(16, 3).value == null || String(ws.getCell(16, 3).value).trim() === '').toBe(
      true
    );
    expect(ws.getCell(20, 3).value == null || String(ws.getCell(20, 3).value).trim() === '').toBe(
      true
    );
  });

  test('trim keeps the system-generated note row', () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('FormXXVI');
    ws.getCell(11, 2).value = 'Selva P';
    ws.getCell(12, 2).value = 'sample leftover';
    ws.getCell(12, 2).border = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'thin' },
      right: { style: 'thin' }
    };
    ws.getCell(18, 2).value = 'This is a System Generated Document';
    trimExcelJsTableToRecordCount(ws, {
      dataStartRow: 11,
      dataRowCount: 1,
      colFrom: 2,
      colTo: 8
    });
    expect(String(ws.getCell(18, 2).value)).toMatch(/system generated/i);
    expect(ws.getCell(12, 2).value == null || String(ws.getCell(12, 2).value).trim() === '').toBe(
      true
    );
  });
});
