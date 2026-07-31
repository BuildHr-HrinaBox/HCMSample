import {
  applyStatutoryDownloadContentAlignment,
  detachOrphanedExcelJSSharedFormulas,
  isStatutoryFormTitleBandText,
  isStatutoryNumericCellValue,
  resetExcelJsWorkbookActiveSheet,
  statutoryCellValueToPlainText,
} from './excelTableBorders';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';

describe('statutory download content alignment helpers', () => {
  test('treats plain numbers and money-like values as numeric', () => {
    expect(isStatutoryNumericCellValue(77765)).toBe(true);
    expect(isStatutoryNumericCellValue('77765')).toBe(true);
    expect(isStatutoryNumericCellValue('7,776.50')).toBe(true);
    expect(isStatutoryNumericCellValue('₹1200')).toBe(true);
    expect(isStatutoryNumericCellValue('(1500)')).toBe(true);
    expect(isStatutoryNumericCellValue('12%')).toBe(true);
  });

  test('treats text codes and dates as non-numeric', () => {
    expect(isStatutoryNumericCellValue('hhhbhj')).toBe(false);
    expect(isStatutoryNumericCellValue('VE123')).toBe(false);
    expect(isStatutoryNumericCellValue('P')).toBe(false);
    expect(isStatutoryNumericCellValue('23-07-2026')).toBe(false);
    expect(isStatutoryNumericCellValue('')).toBe(false);
    expect(isStatutoryNumericCellValue(null)).toBe(false);
  });

  test('plain-text extraction handles rich text', () => {
    expect(
      statutoryCellValueToPlainText({
        richText: [{ text: 'hh' }, { text: 'hbhj' }],
      })
    ).toBe('hhhbhj');
  });

  test('detects Form Number / Rule / Form Name title band (Form XXIII model)', () => {
    expect(isStatutoryFormTitleBandText('FORM XXIII')).toBe(true);
    expect(isStatutoryFormTitleBandText('Form XXIII – Register of Wages')).toBe(true);
    expect(isStatutoryFormTitleBandText('*[See Rule 78(1)(a)(iii)]')).toBe(true);
    expect(
      isStatutoryFormTitleBandText(
        '(Vide Rule 29(2) of A.P. Shops & Establishment Rules, 1990)'
      )
    ).toBe(true);
    expect(isStatutoryFormTitleBandText('Register of Overtime')).toBe(true);
    expect(isStatutoryFormTitleBandText('FORM - W')).toBe(true);
    expect(isStatutoryFormTitleBandText('[See sub-rule(1) of rule (16)]')).toBe(true);
    expect(isStatutoryFormTitleBandText('REGISTER OF WAGES')).toBe(true);
    expect(
      isStatutoryFormTitleBandText('Name and Address of Contractor. : VAYONA ENERGY')
    ).toBe(false);
    expect(isStatutoryFormTitleBandText('Name of workman')).toBe(false);
    expect(isStatutoryFormTitleBandText('Overtime earnings')).toBe(false);
  });

  test('centers title band even when template had left, and never forces text left', () => {
    const cells = new Map();
    const makeCell = (addr, value, alignment = {}) => {
      const cell = { value, alignment: { ...alignment } };
      cells.set(addr, cell);
      return cell;
    };
    const worksheet = {
      eachRow: (_opts, cb) => {
        cb(
          {
            eachCell: (_o, cellCb) => {
              // Was wrongly left — must become center
              cellCb(makeCell('A1', 'FORM XXIII', { horizontal: 'left' }), 1);
              cellCb(makeCell('A2', '*[See Rule 78(1)(a)(iii)]', { horizontal: 'left' }), 1);
              cellCb(makeCell('A3', 'Register of Overtime', { horizontal: 'left' }), 1);
            },
          },
          1
        );
        cb(
          {
            eachCell: (_o, cellCb) => {
              cellCb(
                makeCell('A4', 'Name and Address of Contractor. : VAYONA', {
                  horizontal: 'left',
                }),
                1
              );
              cellCb(makeCell('H4', '67348'), 8);
            },
          },
          4
        );
        cb(
          {
            eachCell: (_o, cellCb) => {
              cellCb(makeCell('A10', 'Rajeshkumar', { horizontal: 'left' }), 1);
              cellCb(makeCell('H10', 67348), 8);
            },
          },
          10
        );
      },
    };

    applyStatutoryDownloadContentAlignment(worksheet);

    expect(cells.get('A1').alignment.horizontal).toBe('center');
    expect(cells.get('A1').alignment.wrapText).toBe(false);
    expect(cells.get('A2').alignment.horizontal).toBe('center');
    expect(cells.get('A2').alignment.wrapText).toBe(false);
    expect(cells.get('A3').alignment.horizontal).toBe('center');
    // Field / body text: left-align override removed — existing alignment kept
    expect(cells.get('A4').alignment.horizontal).toBe('left');
    expect(cells.get('A10').alignment.horizontal).toBe('left');
    expect(cells.get('H4').alignment.horizontal).toBe('right');
    expect(cells.get('H10').alignment.horizontal).toBe('right');
  });
});

describe('resetExcelJsWorkbookActiveSheet', () => {
  test('resets activeTab/firstSheet after pruning extra worksheets', async () => {
    const wb = new ExcelJS.Workbook();
    wb.addWorksheet('Form XII');
    wb.addWorksheet('Form XXIII – Register of Wages');
    wb.views = [
      {
        x: 0,
        y: 0,
        width: 12000,
        height: 16000,
        visibility: 'visible',
        activeTab: 1,
        firstSheet: 1,
      },
    ];
    const keep = wb.worksheets[1];
    [...wb.worksheets].forEach((ws) => {
      if (ws.id !== keep.id) wb.removeWorksheet(ws.id);
    });
    expect(wb.views[0].activeTab).toBe(1);
    resetExcelJsWorkbookActiveSheet(wb);
    expect(wb.views[0].activeTab).toBe(0);
    expect(wb.views[0].firstSheet).toBe(0);

    const out = await wb.xlsx.writeBuffer();
    const zip = await JSZip.loadAsync(out);
    const workbookXml = await zip.file('xl/workbook.xml').async('string');
    expect(workbookXml).toMatch(/activeTab="0"/);
    expect(workbookXml).toMatch(/firstSheet="0"/);
  });
});

describe('detachOrphanedExcelJSSharedFormulas', () => {
  test('converts orphaned shared-formula clones so writeBuffer succeeds', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Form25');
    // Simulate template: master at AM12, clone at AM30 (Form 25 day-31 column).
    ws.getCell('AM12').value = { formula: '1+1', result: 2, shareType: 'shared' };
    ws.getCell('AM30').value = { sharedFormula: 'AM12', result: 2 };
    // Overwrite master the way export does — leaves AM30 pointing at a non-formula cell.
    ws.getCell('AM12').value = 'P';

    expect(detachOrphanedExcelJSSharedFormulas(ws)).toBeGreaterThanOrEqual(1);
    const out = await wb.xlsx.writeBuffer();
    expect(out.byteLength).toBeGreaterThan(0);

    const reload = new ExcelJS.Workbook();
    await reload.xlsx.load(out);
    expect(reload.worksheets[0].getCell('AM12').value).toBe('P');
  });
});
