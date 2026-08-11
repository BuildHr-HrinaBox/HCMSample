import {
  FORM26_ACCIDENTS_NIL_OF_MONTH_TEXT,
  applyForm26AccidentsColumnLayout,
  applyForm26NilTableRows,
  buildForm26NilTableRows,
  findForm26NilPrimaryHeader,
  isForm26AccidentsNilMonthValue,
  resolveForm26NilSpanInfo,
  tableHeadersLookLikeForm26Accidents,
} from './form26RegisterOfAccidents';
import { isForm26NilMonthLineValue } from './form26ADangerousOccurrences';

const FORM26_ACCIDENT_HEADERS = [
  '(1)Sl. No. (Running – Calendar Year)',
  '(2)Date & Hour of Accident',
  '(3)Name & Designation of Person Injured',
  '(4)Exact Place in the Factory(Branch / Dept)',
  '(5)Full Description of Accident',
  '(6)Nature & Extent of Injury',
  '(7)Date of Despatch of Form 18',
];

const FORM26_ACCIDENT_HEADERS_14 = [
  '(1)Sl. No.',
  '(2)Date & Hour of Accident',
  '(3)Name & Designation of Injured Person',
  '(4)Exact Place in the Factory',
  '(5)Full Description of Accident',
  '(6)Nature, Extent & Location of Injury',
  '(7)Date of Return to Work',
  '(8)Date of Despatch of Form 18',
  '(9)Date of Report to Inspector',
  '(10)Date(s) of Absence',
  '(11)Number of Days Lost',
  '(12)Number of Hours Lost',
  '(13)Details of Compensation',
  '(14)Remarks & Initials of Manager',
];

describe('form26RegisterOfAccidents nil month', () => {
  test('table headers look like Form 26 Accidents, not 26-A', () => {
    expect(tableHeadersLookLikeForm26Accidents(FORM26_ACCIDENT_HEADERS)).toBe(true);
  });

  test('finds Sl. No. (Running – Calendar Year) as nil primary column', () => {
    expect(findForm26NilPrimaryHeader(FORM26_ACCIDENT_HEADERS)).toBe(
      '(1)Sl. No. (Running – Calendar Year)'
    );
  });

  test('nil row puts Nill of the month in column (1), not the calendar year', () => {
    const rows = buildForm26NilTableRows(FORM26_ACCIDENT_HEADERS);
    expect(rows).toHaveLength(1);
    expect(rows[0]['(1)Sl. No. (Running – Calendar Year)']).toBe(FORM26_ACCIDENTS_NIL_OF_MONTH_TEXT);
    expect(rows[0]['(1)Sl. No. (Running – Calendar Year)']).not.toBe('2026');
    expect(rows[0]['(2)Date & Hour of Accident']).toBe('');
  });

  test('applyForm26NilTableRows replaces a bare year placeholder row', () => {
    const existing = [
      {
        '(1)Sl. No. (Running – Calendar Year)': '2026',
        '(2)Date & Hour of Accident': '',
        '(3)Name & Designation of Person Injured': '',
      },
    ];
    const rows = applyForm26NilTableRows(FORM26_ACCIDENT_HEADERS, existing);
    expect(rows[0]['(1)Sl. No. (Running – Calendar Year)']).toBe('Nill of the month');
  });

  test('applyForm26NilTableRows keeps real accident entries', () => {
    const existing = [
      {
        '(1)Sl. No. (Running – Calendar Year)': '1',
        '(2)Date & Hour of Accident': '01-Jul-2026 10:00',
        '(3)Name & Designation of Person Injured': 'Ravi / Helper',
      },
    ];
    const rows = applyForm26NilTableRows(FORM26_ACCIDENT_HEADERS, existing);
    expect(rows).toEqual(existing);
  });

  test('recognizes Nill of the month values', () => {
    expect(isForm26AccidentsNilMonthValue('Nill of the month')).toBe(true);
    expect(isForm26NilMonthLineValue('Nill of the month')).toBe(true);
    expect(isForm26NilMonthLineValue('Nil for the month of Jul 2026')).toBe(true);
  });

  test('nil span covers the full table band including Remarks on 14-col Form 26', () => {
    const info = resolveForm26NilSpanInfo(FORM26_ACCIDENT_HEADERS_14);
    expect(info).toEqual({
      startIdx: 0,
      span: 14,
      primaryHeader: '(1)Sl. No.',
    });
  });

  test('applyForm26AccidentsColumnLayout widens narrow columns and wraps headers', () => {
    const widths = {};
    const cells = {};
    const headerRowObj = { height: 15 };
    const worksheet = {
      getColumn: (c) => {
        if (!widths[c]) widths[c] = { width: 8 };
        return widths[c];
      },
      getCell: (r, c) => {
        const key = `${r}:${c}`;
        if (!cells[key]) cells[key] = { alignment: {} };
        return cells[key];
      },
      getRow: () => headerRowObj,
    };
    applyForm26AccidentsColumnLayout(worksheet, [1, 2, 3, 4], 4);
    expect(widths[1].width).toBeGreaterThanOrEqual(10);
    expect(widths[2].width).toBeGreaterThanOrEqual(14);
    expect(cells['4:1'].alignment.wrapText).toBe(true);
    expect(cells['4:2'].alignment.wrapText).toBe(true);
    expect(headerRowObj.height).toBeGreaterThanOrEqual(45);
  });
});
