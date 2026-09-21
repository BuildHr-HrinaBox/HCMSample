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

  test('every month uses the calendar length (July is 31, not June’s 30)', () => {
    expect(getFormVTamilNaduMonthDayCount('January', 2026)).toBe(31);
    expect(getFormVTamilNaduMonthDayCount('February', 2026)).toBe(28);
    expect(getFormVTamilNaduMonthDayCount('March', 2026)).toBe(31);
    expect(getFormVTamilNaduMonthDayCount('April', 2026)).toBe(30);
    expect(getFormVTamilNaduMonthDayCount('May', 2026)).toBe(31);
    expect(getFormVTamilNaduMonthDayCount('June', 2026)).toBe(30);
    expect(getFormVTamilNaduMonthDayCount('July', 2026)).toBe(31);
    expect(getFormVTamilNaduMonthDayCount('Jul', 2026)).toBe(31);
    expect(getFormVTamilNaduMonthDayCount('August', 2026)).toBe(31);
    expect(getFormVTamilNaduMonthDayCount('September', 2026)).toBe(30);
    expect(getFormVTamilNaduMonthDayCount('October', 2026)).toBe(31);
    expect(getFormVTamilNaduMonthDayCount('November', 2026)).toBe(30);
    expect(getFormVTamilNaduMonthDayCount('December', 2026)).toBe(31);
    expect(getFormVTamilNaduMonthDayCount('February', 2024)).toBe(29);
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

  test('clears 1 month(s) junk when Paid_days is missing', () => {
    const {
      applyFormVTamilNaduPaidDaysToRows: applyPaidDays,
      isFormVTamilNaduJunkSummaryValue,
      sanitizeFormVTamilNaduSummaryNumericValue,
      computeFormVTamilNaduTotalHoursWorked,
    } = require('./formVTamilNadu');
    expect(isFormVTamilNaduJunkSummaryValue('1 month(s)')).toBe(true);
    expect(sanitizeFormVTamilNaduSummaryNumericValue('1 month(s)')).toBe('');
    expect(computeFormVTamilNaduTotalHoursWorked('1 month(s)')).toBe('');

    const headers = ['Name', 'Total Days Worked', 'Total Hours Worked'];
    const rows = [
      {
        Name: 'A',
        'Total Days Worked': '1 month(s)',
        'Total Hours Worked': '1 month(s)',
      },
    ];
    const hits = applyPaidDays(rows, [{ EmployeeID: '1' }], headers, {
      daysInMonth: 31,
      resolvePayrollRow: () => null,
    });
    expect(hits).toBe(1);
    expect(rows[0]['Total Days Worked']).toBe('');
    expect(rows[0]['Total Hours Worked']).toBe('');
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

describe('Form V Tamil Nadu period banner', () => {
  test('builds full year period from selected month (not truncated 202)', () => {
    const { buildFormVTamilNaduPeriodLine, writeFormVTamilNaduPeriodToWorksheet } = require('./formVTamilNadu');
    expect(buildFormVTamilNaduPeriodLine('April', 2026)).toBe(
      'For the period from 1st April 2026 to 30th April 2026'
    );
    expect(buildFormVTamilNaduPeriodLine('May', 2026)).toBe(
      'For the period from 1st May 2026 to 31st May 2026'
    );
    expect(buildFormVTamilNaduPeriodLine('April', 202)).toBe('');
  });

  test('overwrites truncated template period cell on worksheet', async () => {
    const ExcelJS = require('exceljs');
    const { buildFormVTamilNaduPeriodLine, writeFormVTamilNaduPeriodToWorksheet } = require('./formVTamilNadu');
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('FORM V');
    ws.getCell(3, 1).value = 'For the period from 1st April 202  to 30th April 202';
    const line = buildFormVTamilNaduPeriodLine('April', 2026);
    expect(writeFormVTamilNaduPeriodToWorksheet(ws, line, 12)).toBe(true);
    expect(String(ws.getCell(3, 1).value)).toBe(
      'For the period from 1st April 2026 to 30th April 2026'
    );
  });

  test('centers FORM-V / REGISTER / period across full table width (not right-aligned)', async () => {
    const ExcelJS = require('exceljs');
    const {
      buildFormVTamilNaduPeriodLine,
      centerFormVTamilNaduTitleBannerRows,
    } = require('./formVTamilNadu');
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('FORM V');
    // Template-style merges only to AP (42); table grows to AQ (43) after day 31.
    ws.mergeCells(1, 1, 1, 42);
    ws.getCell(1, 1).value = 'FORM-V';
    ws.getCell(1, 1).alignment = { horizontal: 'right', vertical: 'middle' };
    ws.mergeCells(2, 1, 2, 42);
    ws.getCell(2, 1).value = 'REGISTER OF EMPLOYMENT.';
    ws.getCell(2, 1).alignment = { horizontal: 'right', vertical: 'middle' };
    ws.mergeCells(3, 1, 3, 42);
    ws.getCell(3, 1).value = '[See sub-rule (1) of rule (16)]';
    ws.getCell(3, 1).alignment = { horizontal: 'right', vertical: 'middle' };
    ws.mergeCells(4, 1, 4, 42);
    ws.getCell(4, 1).value = 'For the period from 1st April 202  to 30th April 202';
    ws.getCell(4, 1).alignment = { horizontal: 'right', vertical: 'middle' };

    const line = buildFormVTamilNaduPeriodLine('May', 2026);
    const result = centerFormVTamilNaduTitleBannerRows(ws, {
      headerRowEnd: 8,
      tableEndCol: 43,
      periodText: line,
    });
    expect(result.centered).toBeGreaterThanOrEqual(4);
    expect(result.endCol).toBe(43);
    expect(String(ws.getCell(1, 1).value)).toMatch(/FORM/i);
    expect(String(ws.getCell(2, 1).value)).toMatch(/REGISTER OF EMPLOYMENT/i);
    expect(String(ws.getCell(4, 1).value)).toBe(line);
    expect(ws.getCell(1, 1).alignment?.horizontal).toBe('center');
    expect(ws.getCell(2, 1).alignment?.horizontal).toBe('center');
    expect(ws.getCell(3, 1).alignment?.horizontal).toBe('center');
    expect(ws.getCell(4, 1).alignment?.horizontal).toBe('center');
  });
});

describe('Form V Tamil Nadu download hints', () => {
  const {
    isFormVStatutoryDownloadHint,
    looksLikeFormVTamilNaduFilename,
  } = require('./formVTamilNadu');

  test('detects Form_V_-_TamilNadu.xlsx and Form V - TamilNadu', () => {
    expect(looksLikeFormVTamilNaduFilename('Form_V_-_TamilNadu.xlsx')).toBe(true);
    expect(looksLikeFormVTamilNaduFilename('Form V - TamilNadu.xlsx')).toBe(true);
    expect(looksLikeFormVTamilNaduFilename('Form V.xlsx')).toBe(true);
    expect(isFormVStatutoryDownloadHint('Form V - TamilNadu')).toBe(true);
    expect(isFormVStatutoryDownloadHint('REGISTER OF EMPLOYMENT', 'Tamil Nadu')).toBe(true);
    expect(isFormVStatutoryDownloadHint('Form_VI_-_TamilNadu.xlsx')).toBe(false);
    expect(isFormVStatutoryDownloadHint('Form XXII Register of Employment', 'Andhra Pradesh')).toBe(
      false
    );
    expect(isFormVStatutoryDownloadHint('Form 11 RJ Register of Employment')).toBe(false);
  });
});

describe('Form V Tamil Nadu Excel visible writes', () => {
  test('maps Name of the Employee from Autofill aliases and Present→P', () => {
    const {
      coerceFormVTamilNaduRowToObject,
      formVTamilNaduIdentityColumnKind,
      normalizeFormVTamilNaduAttendanceCode,
      resolveFormVTamilNaduExportCellValue,
      resolveFormVTamilNaduDayValue,
    } = require('./formVTamilNadu');
    expect(formVTamilNaduIdentityColumnKind('Name of the Employee')).toBe('name');
    expect(formVTamilNaduIdentityColumnKind('Employee Identification No.')).toBe('empid');
    expect(formVTamilNaduIdentityColumnKind('Time at which work commences')).toBe('commence');
    expect(normalizeFormVTamilNaduAttendanceCode('Present')).toBe('P');
    expect(normalizeFormVTamilNaduAttendanceCode('WO')).toBe('WO');

    const headers = ['S.No', 'Name of the Employee', '1', '2'];
    const arrayRow = [1, 'Priya Nair', 'P', 'A'];
    const coerced = coerceFormVTamilNaduRowToObject(arrayRow, headers);
    expect(resolveFormVTamilNaduExportCellValue(coerced, 'Name of the Employee', headers)).toBe('Priya Nair');
    expect(resolveFormVTamilNaduDayValue(coerced, 1, headers)).toBe('P');
    expect(
      resolveFormVTamilNaduExportCellValue(
        { 'Name of the person employed': 'Arun K' },
        'Name of the Employee',
        headers
      )
    ).toBe('Arun K');
  });

  test('writes employee name onto merged non-master cells so Excel shows it', async () => {
    const ExcelJS = require('exceljs');
    const {
      locateFormVTamilNaduIdentityExcelColumns,
      writeFormVTamilNaduExcelVisibleCell,
    } = require('./formVTamilNadu');
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('FORM V');
    ws.mergeCells(10, 1, 11, 1);
    ws.getCell(10, 1).value = 'S.No';
    ws.mergeCells(10, 2, 11, 2);
    ws.getCell(10, 2).value = 'Name of the Employee';
    ws.getCell(10, 3).value = 'Employee Identification No.';
    ws.getCell(11, 7).value = 1;
    ws.getCell(11, 8).value = 2;
    ws.mergeCells(12, 2, 12, 3);

    const identity = locateFormVTamilNaduIdentityExcelColumns(ws, {
      headerRow: 10,
      dayRow: 11,
      firstDayCol: 7,
      startCol: 1,
      cellText: (cell) => String(cell?.value ?? ''),
    });
    expect(identity.sno).toBe(1);
    expect(identity.name).toBe(2);
    expect(identity.empid).toBe(3);

    writeFormVTamilNaduExcelVisibleCell(ws, 12, 2, 'Priya Nair');
    writeFormVTamilNaduExcelVisibleCell(ws, 12, 1, 1);
    writeFormVTamilNaduExcelVisibleCell(ws, 12, 7, 'P');
    expect(String(ws.getCell(12, 2).value)).toBe('Priya Nair');
    expect(Number(ws.getCell(12, 1).value)).toBe(1);
    expect(String(ws.getCell(12, 7).value)).toBe('P');
  });
});
