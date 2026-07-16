import {
  applyFormVTamilNaduPaidDaysToRows,
  getFormVTamilNaduMonthDayCount,
  headersIndicateFormVTamilNaduDayGrid,
  isFormVTamilNaduTotalDaysWorkedHeader,
  readFormVTamilNaduPaidDays,
  resolveFormVTamilNaduDayNumberFromHeader,
  resolveFormVTamilNaduTableHeaders,
} from './formVTamilNadu';

describe('Form V Tamil Nadu month-wise day columns', () => {
  test('resolves day numbers from plain and Parent_Sub headers', () => {
    expect(resolveFormVTamilNaduDayNumberFromHeader('31')).toBe(31);
    expect(resolveFormVTamilNaduDayNumberFromHeader('DATES_30')).toBe(30);
    expect(resolveFormVTamilNaduDayNumberFromHeader('Name of Worker')).toBe(0);
  });

  test('May has 31 days and April has 30', () => {
    expect(getFormVTamilNaduMonthDayCount('May', 2026)).toBe(31);
    expect(getFormVTamilNaduMonthDayCount('April', 2026)).toBe(30);
    expect(getFormVTamilNaduMonthDayCount('February', 2024)).toBe(29);
    expect(getFormVTamilNaduMonthDayCount('February', 2025)).toBe(28);
  });

  test('expands default 30-day band to 31 for May', () => {
    const headers = [
      'S.No',
      'Name of the person employed',
      ...Array.from({ length: 30 }, (_, i) => String(i + 1)),
      'Benefit of National Holiday',
      'Remarks',
    ];
    expect(headersIndicateFormVTamilNaduDayGrid(headers)).toBe(true);
    const next = resolveFormVTamilNaduTableHeaders(headers, 31);
    const days = next
      .map((h) => resolveFormVTamilNaduDayNumberFromHeader(h))
      .filter((d) => d >= 1);
    expect(days).toEqual(Array.from({ length: 31 }, (_, i) => i + 1));
    expect(next[next.length - 2]).toBe('Benefit of National Holiday');
    expect(next[next.length - 1]).toBe('Remarks');
  });

  test('trims day 31 for April (30-day month)', () => {
    const headers = [
      'S.No',
      ...Array.from({ length: 31 }, (_, i) => String(i + 1)),
      'Remarks',
    ];
    const next = resolveFormVTamilNaduTableHeaders(headers, 30);
    const days = next
      .map((h) => resolveFormVTamilNaduDayNumberFromHeader(h))
      .filter((d) => d >= 1);
    expect(days).toEqual(Array.from({ length: 30 }, (_, i) => i + 1));
    expect(next[next.length - 1]).toBe('Remarks');
  });

  test('keeps Parent_Sub day key style when expanding', () => {
    const headers = [
      'Name',
      ...Array.from({ length: 30 }, (_, i) => `DATES_${i + 1}`),
      'Remarks',
    ];
    const next = resolveFormVTamilNaduTableHeaders(headers, 31);
    expect(next).toContain('DATES_31');
    expect(resolveFormVTamilNaduDayNumberFromHeader('DATES_31')).toBe(31);
  });
});

describe('Form V Tamil Nadu Total Days Worked ← Paid_days', () => {
  test('detects Total Days Worked header but not day columns', () => {
    expect(isFormVTamilNaduTotalDaysWorkedHeader('Total Days Worked')).toBe(true);
    expect(isFormVTamilNaduTotalDaysWorkedHeader('31')).toBe(false);
    expect(isFormVTamilNaduTotalDaysWorkedHeader('Remarks')).toBe(false);
  });

  test('reads Paid_days from payroll row', () => {
    expect(readFormVTamilNaduPaidDays({ paid_days: 26 })).toBe('26');
    expect(readFormVTamilNaduPaidDays({ Paid_days: '28' })).toBe('28');
    expect(readFormVTamilNaduPaidDays({ fetch_error: true, paid_days: 26 })).toBe('');
  });

  test('Total Hours = Paid_days × 8 and LOP = month − Paid_days', () => {
    const {
      computeFormVTamilNaduTotalHoursWorked,
      computeFormVTamilNaduLossOfPayDays,
    } = require('./formVTamilNadu');
    expect(computeFormVTamilNaduTotalHoursWorked(31)).toBe('248');
    expect(computeFormVTamilNaduTotalHoursWorked(30)).toBe('240');
    expect(computeFormVTamilNaduLossOfPayDays(31, 31)).toBe('0');
    expect(computeFormVTamilNaduLossOfPayDays(30, 31)).toBe('1');
  });

  test('applies Paid_days onto days, hours, and LOP for each employee', () => {
    const headers = [
      'Name',
      'Total Days Worked',
      'Total Hours Worked',
      'Number of days on Loss of Pay',
      '1',
      '2',
    ];
    const rows = [
      {
        Name: 'A',
        'Total Days Worked': '',
        'Total Hours Worked': '',
        'Number of days on Loss of Pay': '',
      },
    ];
    const hits = applyFormVTamilNaduPaidDaysToRows(rows, [{ EmployeeID: '1' }], headers, {
      daysInMonth: 31,
      resolvePayrollRow: () => ({ paid_days: 30 }),
    });
    expect(hits).toBe(1);
    expect(rows[0]['Total Days Worked']).toBe('30');
    expect(rows[0]['Total Hours Worked']).toBe('240');
    expect(rows[0]['Number of days on Loss of Pay']).toBe('1');
  });
});

describe('Form V Tamil Nadu Excel day-column insert', () => {
  test('inserts missing day 31 after day 30 and shifts Total Days Worked right', async () => {
    const ExcelJS = require('exceljs');
    const { ensureFormVTamilNaduExcelDayColumns, locateFormVTamilNaduSummaryExcelColumns } = require('./formVTamilNadu');
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('FORM V');
    ws.getCell(11, 1).value = 28;
    ws.getCell(11, 2).value = 29;
    ws.getCell(11, 3).value = 30;
    ws.getCell(10, 4).value = 'Total Days Worked';
    ws.getCell(10, 5).value = 'Total Hours Worked';

    const dayColsByNumber = new Map([
      [28, 1],
      [29, 2],
      [30, 3],
    ]);
    const ensured = ensureFormVTamilNaduExcelDayColumns(ws, {
      dayRow: 11,
      dayColsByNumber,
      markerCols: [1, 2, 3],
      daysInMonth: 31,
    });
    expect(ensured.inserted).toBe(1);
    expect(ensured.dayColsByNumber.get(31)).toBe(4);
    expect(String(ws.getCell(11, 4).value)).toBe('31');
    expect(String(ws.getCell(10, 5).value)).toMatch(/total days worked/i);

    const summary = locateFormVTamilNaduSummaryExcelColumns(ws, {
      headerRow: 10,
      dayRow: 11,
      afterCol: 4,
      cellText: (cell) => String(cell?.value ?? ''),
    });
    expect(summary.totalDaysWorked).toBe(5);
    expect(summary.totalHoursWorked).toBe(6);
  });

  test('manual-shifts when summary label stays at insert column (splice no-op)', () => {
    const ExcelJS = require('exceljs');
    const { ensureFormVTamilNaduExcelDayColumns } = require('./formVTamilNadu');
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('FORM V');
    ws.getCell(11, 1).value = 30;
    ws.getCell(10, 2).value = 'Total Days Worked';
    // Simulate spliceColumns that does nothing useful by replacing it.
    ws.spliceColumns = () => {};
    const ensured = ensureFormVTamilNaduExcelDayColumns(ws, {
      dayRow: 11,
      dayColsByNumber: new Map([[30, 1]]),
      markerCols: [1],
      daysInMonth: 31,
      maxScanRows: 20,
      maxScanCols: 20,
    });
    expect(ensured.inserted).toBe(1);
    expect(ensured.dayColsByNumber.get(31)).toBe(2);
    expect(String(ws.getCell(11, 2).value)).toBe('31');
    expect(String(ws.getCell(10, 3).value)).toMatch(/total days worked/i);
  });

  test('on real Form V.xlsx template, inserts day 31 at AK (col 37) after day 30', async () => {
    const ExcelJS = require('exceljs');
    const fs = require('fs');
    const path = require('path');
    const {
      ensureFormVTamilNaduExcelDayColumns,
      locateFormVTamilNaduSummaryExcelColumns,
    } = require('./formVTamilNadu');
    const templatePath = path.resolve(
      __dirname,
      '../../../../functions/statutoryreg_function/templates/Form V.xlsx'
    );
    if (!fs.existsSync(templatePath)) {
      console.warn('Form V.xlsx not found — skip template layout test');
      return;
    }
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(fs.readFileSync(templatePath));
    const ws = wb.worksheets[0];
    const dayRow = 11;
    const dayColsByNumber = new Map();
    const markerCols = [];
    for (let c = 1; c <= 50; c += 1) {
      const t = String(ws.getCell(dayRow, c).value ?? '').trim();
      if (/^\d{1,2}$/.test(t)) {
        const n = Number(t);
        if (n >= 1 && n <= 31) {
          dayColsByNumber.set(n, c);
          markerCols.push(c);
        }
      }
    }
    expect(dayColsByNumber.get(30)).toBe(36); // AJ
    expect(dayColsByNumber.has(31)).toBe(false);
    expect(String(ws.getCell(10, 37).value)).toMatch(/total days worked/i); // AK before insert

    const ensured = ensureFormVTamilNaduExcelDayColumns(ws, {
      dayRow,
      dayColsByNumber,
      markerCols,
      daysInMonth: 31,
      maxScanRows: 80,
      maxScanCols: 80,
    });
    expect(ensured.inserted).toBe(1);
    expect(ensured.dayColsByNumber.get(31)).toBe(37); // AK
    expect(String(ws.getCell(11, 37).value)).toBe('31');
    // Totals must shift right: AL / AM
    expect(String(ws.getCell(10, 38).value)).toMatch(/total days worked/i);
    expect(String(ws.getCell(10, 39).value)).toMatch(/total hours worked/i);

    const summary = locateFormVTamilNaduSummaryExcelColumns(ws, {
      headerRow: 10,
      dayRow: 11,
      afterCol: 37,
      cellText: (cell) => String(cell?.value ?? ''),
    });
    expect(summary.totalDaysWorked).toBe(38);
    expect(summary.totalHoursWorked).toBe(39);
    expect(summary.lossOfPay).toBe(40);

    // Daily Hours merge should include new day-31 column (G10:AK10).
    const merges = Array.isArray(ws.model?.merges) ? ws.model.merges : [];
    expect(merges.some((m) => /^G10:AK10$/i.test(String(m)))).toBe(true);
  });

  test('inserts day 31 when Total Days Worked header is split vertically', () => {
    const ExcelJS = require('exceljs');
    const {
      ensureFormVTamilNaduExcelDayColumns,
      locateFormVTamilNaduSummaryExcelColumns,
    } = require('./formVTamilNadu');
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('FORM V');
    // Days 28-30 then vertically split summary (real Form V template layout).
    ws.getCell(11, 1).value = 28;
    ws.getCell(11, 2).value = 29;
    ws.getCell(11, 3).value = 30;
    ws.getCell(9, 4).value = 'Total';
    ws.getCell(10, 4).value = 'Days';
    ws.getCell(11, 4).value = 'Worked';
    ws.getCell(9, 5).value = 'Total';
    ws.getCell(10, 5).value = 'Hours';
    ws.getCell(11, 5).value = 'Worked';
    ws.getCell(9, 6).value = 'Number of days';
    ws.getCell(10, 6).value = 'Loss of Pay';
    ws.spliceColumns = () => {};

    const ensured = ensureFormVTamilNaduExcelDayColumns(ws, {
      dayRow: 11,
      dayColsByNumber: new Map([
        [28, 1],
        [29, 2],
        [30, 3],
      ]),
      markerCols: [1, 2, 3],
      daysInMonth: 31,
      maxScanRows: 30,
      maxScanCols: 30,
    });
    expect(ensured.inserted).toBe(1);
    expect(ensured.dayColsByNumber.get(31)).toBe(4);
    expect(String(ws.getCell(11, 4).value)).toBe('31');
    // Split "Total / Days / Worked" must move to column 5.
    expect(String(ws.getCell(9, 5).value)).toMatch(/total/i);
    expect(String(ws.getCell(10, 5).value)).toMatch(/days/i);

    const summary = locateFormVTamilNaduSummaryExcelColumns(ws, {
      headerRow: 10,
      dayRow: 11,
      afterCol: 4,
      cellText: (cell) => String(cell?.value ?? ''),
    });
    expect(summary.totalDaysWorked).toBe(5);
    expect(summary.totalHoursWorked).toBe(6);
    expect(summary.lossOfPay).toBe(7);
  });

  test('locates National/Festival/Remarks leave-blank columns after LOP', () => {
    const ExcelJS = require('exceljs');
    const {
      locateFormVTamilNaduLeaveBlankExcelColumns,
      locateFormVTamilNaduSummaryExcelColumns,
    } = require('./formVTamilNadu');
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('FORM V');
    ws.getCell(10, 37).value = 'Total Days Worked';
    ws.getCell(10, 38).value = 'Total Hours Worked';
    ws.getCell(10, 39).value = 'Number of days on Loss of Pay';
    ws.getCell(10, 40).value = 'Benefit availed for working on National Holiday (**)';
    ws.getCell(10, 41).value = 'Benefit availed for working on Festival Holiday (**)';
    ws.getCell(10, 42).value = 'Remarks';
    // Must not treat Approved Festival Holidays as a leave-blank data column.
    ws.getCell(7, 12).value = 'Approved Festival Holidays:';

    const summary = locateFormVTamilNaduSummaryExcelColumns(ws, {
      headerRow: 10,
      afterCol: 37,
      cellText: (cell) => String(cell?.value ?? ''),
    });
    expect(summary.lossOfPay).toBe(39);
    expect(summary.nationalHolidayBenefit).toBe(40);
    expect(summary.festivalHolidayBenefit).toBe(41);
    expect(summary.remarks).toBe(42);

    const blankCols = locateFormVTamilNaduLeaveBlankExcelColumns(ws, {
      headerRow: 10,
      afterCol: 37,
      lastDayCol: 36,
      cellText: (cell) => String(cell?.value ?? ''),
    });
    expect(blankCols).toEqual([40, 41, 42]);
    expect(blankCols).not.toContain(12);
  });
});
