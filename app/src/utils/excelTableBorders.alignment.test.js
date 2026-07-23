import {
  applyStatutoryDownloadContentAlignment,
  isStatutoryFormTitleBandText,
  isStatutoryNumericCellValue,
  statutoryCellValueToPlainText,
} from './excelTableBorders';

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
    expect(isStatutoryFormTitleBandText('*[See Rule 78(1)(a)(iii)]')).toBe(true);
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
    expect(cells.get('A2').alignment.horizontal).toBe('center');
    expect(cells.get('A3').alignment.horizontal).toBe('center');
    // Field / body text: left-align override removed — existing alignment kept
    expect(cells.get('A4').alignment.horizontal).toBe('left');
    expect(cells.get('A10').alignment.horizontal).toBe('left');
    expect(cells.get('H4').alignment.horizontal).toBe('right');
    expect(cells.get('H10').alignment.horizontal).toBe('right');
  });
});
