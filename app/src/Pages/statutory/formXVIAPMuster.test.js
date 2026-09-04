import ExcelJS from 'exceljs';
import {
  applyFormXVIAPMusterLayoutFixes,
  clampFormXVIAPHeadersToRemarks,
  FORM_XVI_AP_COL_AK,
  FORM_XVI_AP_COL_AL,
  resolveFormXVIAPLastKeepCol
} from './formXXIIAPRegisterOfEmployment';

describe('Form XVI AP Muster Roll Excel layout helpers', () => {
  it('stops headers at Remarks so overflow columns are not exported', () => {
    expect(
      clampFormXVIAPHeadersToRemarks([
        'S. No',
        'Name of the Employee',
        '1',
        '31',
        'Remarks',
        'Aadhaar',
        'Father leftover'
      ])
    ).toEqual(['S. No', 'Name of the Employee', '1', '31', 'Remarks']);
  });

  it('keeps the table through Remarks / day 31 and excludes AK/AL when days end before AK', () => {
    const dayColumnMap = new Map();
    for (let day = 1; day <= 31; day += 1) dayColumnMap.set(day, 4 + day); // E=5 … AI=35
    const lastKeep = resolveFormXVIAPLastKeepCol(dayColumnMap, 36, [1, 2, 3, 4, 35, 36]);
    expect(lastKeep).toBeLessThan(FORM_XVI_AP_COL_AK);
    expect(lastKeep).toBeLessThan(FORM_XVI_AP_COL_AL);
    expect(lastKeep).toBe(36);
  });

  it('writes Nature and For the Month of as full single-line bands above Dates', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('XVI-Muster Roll');
    // Template only has the "For the" fragment (real AP template behaviour).
    ws.getCell(8, 25).value = 'For the';
    ws.getCell(9, 5).value = 'Dates';
    ws.mergeCells(9, 5, 9, 35);
    for (let day = 1; day <= 31; day += 1) {
      ws.getCell(10, 4 + day).value = String(day);
    }
    ws.mergeCells(11, 5, 11, 35);
    ws.getCell(11, 5).value = 'P';
    ws.getCell(11, FORM_XVI_AP_COL_AK).value = 'ata ramana';
    ws.getCell(11, FORM_XVI_AP_COL_AK).border = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'thin' },
      right: { style: 'thin' }
    };
    ws.getCell(11, FORM_XVI_AP_COL_AL).value = 3.14e17;
    ws.getCell(11, FORM_XVI_AP_COL_AL).border = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'thin' },
      right: { style: 'thin' }
    };

    const dayColumnMap = new Map();
    for (let day = 1; day <= 31; day += 1) dayColumnMap.set(day, 4 + day);

    applyFormXVIAPMusterLayoutFixes(ws, {
      headerRow: 9,
      markerRow: 10,
      dataStartRow: 11,
      dayColumnMap,
      remarksCol: 36,
      lastKeepCol: 36,
      sourceRowCount: 2,
      monthYearText: 'April 2026',
      headerFormData: { form_xvi_nature_location_work: 'AP-Tadipatri' },
      maxScanRows: 20,
      attendanceYear: 2026,
      attendanceMonthIndex: 3
    });

    const merges = ws.model?.merges || [];
    expect(merges.some((m) => /^E9:AI9$/i.test(String(m)))).toBe(true);
    expect(String(ws.getCell(9, 5).value || '')).toMatch(/^dates?$/i);
    expect(merges.some((m) => /^E11:AI11$/i.test(String(m)))).toBe(false);

    expect(String(ws.getCell(8, 1).value || '')).toBe('Nature and Location of work : AP-Tadipatri');
    expect(ws.getCell(8, 1).alignment?.wrapText).toBe(false);
    expect(ws.getCell(8, 1).font?.bold).toBe(true);
    expect(String(ws.getCell(8, 5).value || '')).toBe('For the Month of : April 2026');
    expect(ws.getCell(8, 5).alignment?.wrapText).toBe(false);
    expect(ws.getCell(8, 5).font?.bold).toBe(true);
    // Old "For the" fragment must be cleared.
    expect(String(ws.getCell(8, 25).value || '')).not.toMatch(/^For the$/i);

    expect(ws.getCell(11, FORM_XVI_AP_COL_AK).value).toBeNull();
    expect(ws.getCell(11, FORM_XVI_AP_COL_AL).value).toBeNull();
    expect(ws.getCell(11, FORM_XVI_AP_COL_AK).border).toBeFalsy();
    expect(ws.getCell(11, FORM_XVI_AP_COL_AL).border).toBeFalsy();
  });

  it('strips leftover Form XVI header dump from the For the Month of cell', () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('XVI-Muster Roll');
    ws.getCell(5, 6).value =
      'For the Month of : May 2026 Form XVI - Muster Roll Address of the Establishment : M/S VIBRANT GREENTECH Name and Address of Contractor. : VAYONA ENERGY Name and address of Principal Employer : Vibrant Nature and Location of work : AP-Tadipatri For the Month of : May 2026';
    ws.getCell(6, 1).value = 'Nature and Location of work : AP-Tadipatri';
    ws.getCell(8, 5).value = 'Dates';

    applyFormXVIAPMusterLayoutFixes(ws, {
      headerRow: 8,
      markerRow: 9,
      dataStartRow: 10,
      dayColumnMap: new Map([[1, 5], [31, 35]]),
      remarksCol: -1,
      lastKeepCol: 36,
      sourceRowCount: 0,
      monthYearText: 'May 2026',
      headerFormData: { form_xvi_nature_location_work: 'AP-Tadipatri' },
      maxScanRows: 12
    });

    expect(String(ws.getCell(7, 1).value || '')).toBe('Nature and Location of work : AP-Tadipatri');
    expect(String(ws.getCell(7, 5).value || '')).toBe('For the Month of : May 2026');
    expect(String(ws.getCell(7, 5).value || '')).not.toMatch(/VAYONA|Establishment|Form XVI/i);
  });
});
