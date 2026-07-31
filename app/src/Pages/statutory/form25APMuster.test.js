import ExcelJS from 'exceljs';
import {
  clearForm25APMusterRemarksColumn,
  filterForm25APMusterHeadersForMonthDays,
  hideForm25APMusterExtraDayColumns,
  shrinkForm25APMusterMergesPastColumn,
} from './form25APMuster';

describe('Form 25 AP Muster download column helpers', () => {
  test('filterForm25APMusterHeadersForMonthDays drops day 31 for 30-day months', () => {
    const headers = [
      'Serial',
      'Name',
      ...Array.from({ length: 31 }, (_, i) => `For the period ending_${i + 1}`),
      'Remarks',
    ];
    const filtered = filterForm25APMusterHeadersForMonthDays(headers, 30);
    expect(filtered).not.toContain('For the period ending_31');
    expect(filtered).toContain('For the period ending_30');
    expect(filtered).toContain('Remarks');
  });

  test('hideForm25APMusterExtraDayColumns hides day 31 and blanks Remarks body', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('25');
    // Days 30 / 31 at AL / AM (1-based 38 / 39); Remarks at AN (40).
    ws.getCell(6, 38).value = 30;
    ws.getCell(6, 39).value = 31;
    ws.getCell(5, 40).value = 'Remarks';
    ws.getCell(8, 39).value = 'P';
    ws.getCell(8, 40).value = 'ramnath k';
    ws.getCell(9, 40).value = 'aravind';
    ws.mergeCells('W5:AM5');

    const dayColumnMap = new Map([
      [30, 37],
      [31, 38],
    ]);
    const hidden = hideForm25APMusterExtraDayColumns(ws, dayColumnMap, 30, {
      headerRow: 6,
      clearRowTo: 20,
    });
    expect(hidden).toEqual([39]);
    expect(ws.getColumn(39).hidden).toBe(true);
    expect(ws.getCell(6, 39).value).toBeNull();
    expect(ws.getCell(8, 39).value).toBeNull();

    clearForm25APMusterRemarksColumn(ws, 40, { headerRow: 5, clearRowTo: 20 });
    expect(ws.getCell(5, 40).value).toBe('Remarks');
    expect(ws.getCell(8, 40).value).toBeNull();
    expect(ws.getCell(9, 40).value).toBeNull();

    const out = await wb.xlsx.writeBuffer();
    expect(out.byteLength).toBeGreaterThan(0);
  });

  test('shrinkForm25APMusterMergesPastColumn trims band into last visible day', () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('25');
    ws.mergeCells('W5:AM5');
    // Last visible day col = AL (38) for a 30-day month.
    shrinkForm25APMusterMergesPastColumn(ws, 38);
    const merges = ws.model.merges || [];
    expect(merges.some((m) => String(m).toUpperCase() === 'W5:AL5')).toBe(true);
    expect(merges.some((m) => String(m).toUpperCase().includes('AM5'))).toBe(false);
  });
});
