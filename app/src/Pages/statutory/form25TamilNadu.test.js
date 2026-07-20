import {
  applyForm25TamilNaduExportColumnWidths,
  applyForm25TamilNaduPrefixColumnsToRows,
  buildForm25TamilNaduOrderedExcelCols,
  ensureForm25TamilNaduDayColumnHeaders,
  ensureForm25TamilNaduPrefixHeaders,
  form25TamilNaduSummaryColForHeader,
  getForm25TamilNaduEmployeeName,
  isForm25TamilNaduSerialInRegisterHeader,
  isForm25TamilNaduWorkerNameHeader,
  locateForm25TamilNaduSummaryExcelColumns,
  resolveForm25TamilNaduDayNumberFromHeader,
  resolveForm25TamilNaduExportCellValue,
  resolveForm25TamilNaduExportHeaders,
  trimForm25TamilNaduHeadersAfterRemarks,
} from './form25TamilNadu';

describe('Form 25 Tamil Nadu Excel alignment', () => {
  test('resolves day numbers from bare and Dates_N headers', () => {
    expect(resolveForm25TamilNaduDayNumberFromHeader('1')).toBe(1);
    expect(resolveForm25TamilNaduDayNumberFromHeader('31')).toBe(31);
    expect(resolveForm25TamilNaduDayNumberFromHeader('Dates_12')).toBe(12);
    expect(resolveForm25TamilNaduDayNumberFromHeader('Total Days Worked')).toBe(0);
  });

  test('detects Serial-in-Register vs Name of the Worker', () => {
    expect(
      isForm25TamilNaduSerialInRegisterHeader(
        'Serial Number in Register of adult workers and young persons'
      )
    ).toBe(true);
    expect(isForm25TamilNaduWorkerNameHeader('Name of the Worker')).toBe(true);
    expect(
      isForm25TamilNaduWorkerNameHeader(
        'Serial Number in Register of adult workers and young persons'
      )
    ).toBe(false);
  });

  test('inserts missing Serial header before Name so Name is not written under Serial column', () => {
    const headers = ['S.No', 'Name of the Worker', 'Worker Identify Number'];
    const ensured = ensureForm25TamilNaduPrefixHeaders(headers);
    expect(ensured[0]).toBe('S.No');
    expect(isForm25TamilNaduSerialInRegisterHeader(ensured[1])).toBe(true);
    expect(ensured[2]).toBe('Name of the Worker');
  });

  test('array rows splice Serial slot when header was missing', () => {
    const headers = ['S.No', 'Name of the Worker', 'Worker Identify Number'];
    const rows = [['1', 'Rajeshkumar Ramasamy', 'VE0447']];
    const fixed = applyForm25TamilNaduPrefixColumnsToRows(rows, headers, [
      { EmployeeName: 'Rajeshkumar Ramasamy' },
    ]);
    const ensured = ensureForm25TamilNaduPrefixHeaders(headers);
    const serialIdx = ensured.findIndex((h) => isForm25TamilNaduSerialInRegisterHeader(h));
    const nameIdx = ensured.findIndex((h) => isForm25TamilNaduWorkerNameHeader(h));
    expect(fixed[0][serialIdx]).toBe('1');
    expect(fixed[0][nameIdx]).toBe('Rajeshkumar Ramasamy');
    expect(fixed[0][nameIdx + 1]).toBe('VE0447');
  });

  test('Serial Number in Register → 1..n and Name ← EmployeeName; clears after Remarks', () => {
    const headers = [
      'S.No',
      'Serial Number in Register of adult workers and young persons',
      'Name of the Worker',
      'Worker Identify Number',
      'Remarks',
      'Orphan Name',
      'Orphan Id',
    ];
    const rows = [
      {
        'S.No': '1',
        'Serial Number in Register of adult workers and young persons': 'Rajeshkumar',
        'Name of the Worker': '3',
        'Worker Identify Number': 'VE0447',
        Remarks: '',
        'Orphan Name': 'spill-name',
        'Orphan Id': '313989000000559040',
      },
    ];
    const fixed = applyForm25TamilNaduPrefixColumnsToRows(rows, headers, [
      { EmployeeName: 'Rajeshkumar R' },
    ]);
    expect(fixed[0]['Serial Number in Register of adult workers and young persons']).toBe('1');
    expect(fixed[0]['Name of the Worker']).toBe('Rajeshkumar R');
    expect(fixed[0]['Orphan Name']).toBe('');
    expect(fixed[0]['Orphan Id']).toBe('');
  });

  test('recovers Name from Serial when EmployeeName missing', () => {
    const headers = [
      'Serial Number in Register of adult workers and young persons',
      'Name of the Worker',
    ];
    const rows = [
      {
        'Serial Number in Register of adult workers and young persons': 'Raja',
        'Name of the Worker': '',
      },
    ];
    const fixed = applyForm25TamilNaduPrefixColumnsToRows(rows, headers, [{}]);
    expect(fixed[0]['Name of the Worker']).toBe('Raja');
    expect(fixed[0]['Serial Number in Register of adult workers and young persons']).toBe('1');
  });

  test('trims headers after Remarks', () => {
    const headers = ['S.No', 'Name of the Worker', 'Remarks', 'Extra1', 'Extra2'];
    expect(trimForm25TamilNaduHeadersAfterRemarks(headers)).toEqual([
      'S.No',
      'Name of the Worker',
      'Remarks',
    ]);
  });

  test('keeps blank Rest Interval so Scheme / day-1 do not shift left', () => {
    const headers = [
      'S.No',
      'Serial Number in Register of adult workers and young persons',
      'Name of the Worker',
      'Worker Identify Number',
      'Time at which work commences',
      '',
      'Time at which work ends',
      'Scheme of Shifts',
      '1',
      '2',
      'Total Days Worked',
    ];
    const ensured = ensureForm25TamilNaduDayColumnHeaders(headers);
    expect(ensured[5]).toBe('');
    expect(ensured[6]).toBe('Time at which work ends');
    expect(ensured[7]).toBe('Scheme of Shifts');
  });

  test('inserts missing day columns before summary headers', () => {
    const headers = [
      'S.No',
      'Name of the Worker',
      '1',
      '2',
      'Total Days Worked',
      'Total Hours Worked',
    ];
    const ensured = ensureForm25TamilNaduDayColumnHeaders(headers);
    expect(ensured.filter((h) => resolveForm25TamilNaduDayNumberFromHeader(h) >= 1)).toHaveLength(31);
    expect(ensured[ensured.length - 2]).toBe('Total Days Worked');
    expect(ensured[ensured.length - 1]).toBe('Total Hours Worked');
  });

  test('maps Total Days / Hours / LOP after day band — not into day columns', () => {
    const prefix = [
      'S.No',
      'Serial Number in Register of adult workers and young persons',
      'Name of the Worker',
      'Worker Identify Number',
      'Time at which work commences',
      'Rest Interval',
      'Time at which work ends',
      'Scheme of Shifts',
    ];
    const days = Array.from({ length: 21 }, (_, i) => String(i + 1));
    const headers = resolveForm25TamilNaduExportHeaders(
      [...prefix, ...days, 'Total Days Worked', 'Total Hours Worked', 'Number of days on Loss of Pay'],
      []
    );
    const dayColsByNumber = new Map();
    const markerCols = [];
    for (let d = 1; d <= 31; d += 1) {
      const col = 8 + d;
      dayColsByNumber.set(d, col);
      markerCols.push(col);
    }
    const summaryCols = {
      totalDaysWorked: 40,
      totalHoursWorked: 41,
      lossOfPay: 42,
    };
    const dayStartIdx = headers.findIndex((h) => resolveForm25TamilNaduDayNumberFromHeader(h) >= 1);
    let lastDayHeaderIdx = -1;
    headers.forEach((h, i) => {
      if (resolveForm25TamilNaduDayNumberFromHeader(h) >= 1) lastDayHeaderIdx = i;
    });
    const ordered = buildForm25TamilNaduOrderedExcelCols({
      sourceHeaders: headers,
      startCol: 1,
      dayStartIdx,
      lastDayHeaderIdx,
      firstDayCol: 9,
      lastDayCol: 39,
      dayColsByNumber,
      markerCols,
      summaryCols,
      daysInMonth: 31,
    });

    const totalDaysIdx = headers.indexOf('Total Days Worked');
    const totalHoursIdx = headers.indexOf('Total Hours Worked');
    const lopIdx = headers.indexOf('Number of days on Loss of Pay');
    expect(ordered[totalDaysIdx]).toBe(40);
    expect(ordered[totalHoursIdx]).toBe(41);
    expect(ordered[lopIdx]).toBe(42);
    expect(ordered[headers.indexOf('22')]).toBe(30);
    expect(ordered[headers.indexOf('22')]).not.toBe(40);
  });

  test('does not put worker name into Serial Number via index fallback', () => {
    const row = {
      'S.No': '1',
      'Name of the Worker': 'Rajeshkumar',
      'Worker Identify Number': 'VE0447',
      'Time at which work ends': '05:00',
      'Scheme of Shifts': 'General Shift Sites',
    };
    expect(
      resolveForm25TamilNaduExportCellValue(
        row,
        'Serial Number in Register of adult workers and young persons'
      )
    ).toBe('');
    expect(resolveForm25TamilNaduExportCellValue(row, 'Name of the Worker')).toBe('Rajeshkumar');
    expect(resolveForm25TamilNaduExportCellValue(row, 'Time at which work ends')).toBe('05:00');
    expect(resolveForm25TamilNaduExportCellValue(row, 'Scheme of Shifts')).toBe(
      'General Shift Sites'
    );
  });

  test('summaryColForHeader classifies Form 25 trailing labels', () => {
    const cols = {
      totalDaysWorked: 40,
      totalHoursWorked: 41,
      lossOfPay: 42,
      signatureWorker: 47,
    };
    expect(form25TamilNaduSummaryColForHeader('Total Days Worked', cols)).toBe(40);
    expect(form25TamilNaduSummaryColForHeader('Signature of Worker', cols)).toBe(47);
    expect(form25TamilNaduSummaryColForHeader('22', cols)).toBeNull();
  });

  test('locateForm25TamilNaduSummaryExcelColumns finds labels after day band', () => {
    const cells = new Map();
    const set = (r, c, v) => cells.set(`${r}:${c}`, v);
    set(9, 40, 'Total Days Worked');
    set(9, 41, 'Total Hours Worked');
    set(9, 42, 'Number of days on Loss of Pay');
    set(9, 43, 'Number of days entitled for National Holidays Benefit');
    const worksheet = {
      getCell: (r, c) => ({ value: cells.get(`${r}:${c}`) ?? '' }),
    };
    const found = locateForm25TamilNaduSummaryExcelColumns(worksheet, {
      headerRow: 9,
      dayRow: 10,
      afterCol: 40,
      maxScanCols: 50,
      cellText: (cell) => String(cell?.value ?? ''),
    });
    expect(found.totalDaysWorked).toBe(40);
    expect(found.totalHoursWorked).toBe(41);
    expect(found.lossOfPay).toBe(42);
    expect(found.nationalHolidayBenefit).toBe(43);
  });

  test('applyForm25TamilNaduExportColumnWidths widens narrow day columns', () => {
    const widths = new Map();
    const worksheet = {
      getColumn: (col) => {
        if (!widths.has(col)) widths.set(col, { width: 2 });
        return widths.get(col);
      },
    };
    applyForm25TamilNaduExportColumnWidths(worksheet, {
      dayCols: [9, 10, 11],
      summaryCols: { totalDaysWorked: 40, signatureWorker: 47 },
      startCol: 1,
      prefixCount: 8,
    });
    expect(widths.get(9).width).toBeGreaterThanOrEqual(4);
    expect(widths.get(40).width).toBeGreaterThanOrEqual(12);
    expect(widths.get(47).width).toBeGreaterThanOrEqual(16);
  });

  test('does not treat Tamil Nadu Form 25 day grid as AP Muster', () => {
    const {
      looksLikeForm25APMusterPartialHeaders,
      headersIndicateForm25APMusterTable,
    } = require('./form25APMuster');
    const tnHeaders = [
      'S.No',
      'Serial Number in Register of adult workers and young persons',
      'Name of the Worker',
      'Worker Identify Number',
      'Scheme of Shifts',
      ...Array.from({ length: 31 }, (_, i) => String(i + 1)),
      'Total Days Worked',
      'Remarks',
    ];
    expect(looksLikeForm25APMusterPartialHeaders(tnHeaders)).toBe(false);
    expect(headersIndicateForm25APMusterTable(tnHeaders)).toBe(false);
  });
});
