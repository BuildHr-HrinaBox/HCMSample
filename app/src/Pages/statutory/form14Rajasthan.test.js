import {
  FORM_14_RJ_DISPLAY_REFERENCE,
  FORM_14_RJ_DISPLAY_SUBTITLE,
  FORM_14_RJ_DISPLAY_TITLE,
  FORM_14_RJ_YOUNG_DEFAULT,
  applyForm14RajasthanAutofillDefaultsToRows,
  applyForm14RajasthanPaidDaysToRows,
  computeForm14RajasthanTotalHours,
  enrichForm14RajasthanDisplayHeader,
  isForm14RajasthanContext,
  isForm14RajasthanTotalHoursHeader,
  isForm14RajasthanYoungPersonHeader,
  looksLikeForm14RajasthanFileName,
  readForm14RajasthanPaidDays,
} from './form14Rajasthan';
import { isForm11RajasthanContext } from './form11Rajasthan';
import { isFormXIIIRegisterOfWorkmenContext } from './formXAPRegisterOfFines';

describe('form14Rajasthan', () => {
  test('detects Form_14_RJ filename context', () => {
    expect(looksLikeForm14RajasthanFileName('Form_14_RJ_-_Rajasthan.xlsx')).toBe(true);
    expect(
      isForm14RajasthanContext({}, { formFileName: 'Form_14_RJ_-_Rajasthan.xlsx' }, '', '', [])
    ).toBe(true);
  });

  test('does not treat Form XIV Employment Card as Form 14 RJ', () => {
    expect(looksLikeForm14RajasthanFileName('Form_XIV_MP.xlsx')).toBe(false);
    expect(
      isForm14RajasthanContext(
        { title: 'FORM XIV', subtitle: 'Employment Card' },
        { state: 'Madhya Pradesh', formName: 'Form XIV' },
        'Form_XIV.xlsx',
        'Employment Card'
      )
    ).toBe(false);
  });

  test('does not treat Form_11_RJ as Form 14', () => {
    expect(
      isForm14RajasthanContext({}, { formFileName: 'Form_11_RJ_-_Rajasthan.xlsx' }, '', '', [])
    ).toBe(false);
  });

  test('Form_14_RJ is not Form 11 RJ', () => {
    expect(
      isForm11RajasthanContext({}, { formFileName: 'Form_14_RJ_-_Rajasthan.xlsx' }, '', '', [
        'Name of Persons Employed',
        'Whether young person or not',
        'Time at which employment commences',
        '2 Rest interval',
        'Total hours worked during the month',
      ])
    ).toBe(false);
  });

  test('Form_14_RJ is not Form XIII Register of Workmen', () => {
    expect(
      isFormXIIIRegisterOfWorkmenContext(
        { title: 'FORM XIII', subtitle: 'Register of Workmen' },
        { state: 'Rajasthan', formName: 'Form 14_RJ' },
        'Form_14_RJ_-_Rajasthan.xlsx',
        'FORM XIII Register of Workmen (Contract Labour) Rule 22'
      )
    ).toBe(false);
  });

  test('enrich sets Form 14 display title', () => {
    const {
      FORM_14_RJ_ACT_NAME,
      FORM_14_RJ_NOTICE_LINE,
    } = require('./form14Rajasthan');
    const next = enrichForm14RajasthanDisplayHeader(
      { title: 'Form XIII – Register of Workmen', subtitle: 'Register of Workmen (Contract Labour)' },
      'Form_14_RJ_-_Rajasthan.xlsx',
      { formName: 'Form 14_RJ', state: 'Rajasthan' },
      [],
      'Rule 22'
    );
    expect(next.title).toBe(FORM_14_RJ_DISPLAY_TITLE);
    expect(next.subtitle).toBe(FORM_14_RJ_DISPLAY_SUBTITLE);
    expect(next.reference).toBe(FORM_14_RJ_DISPLAY_REFERENCE);
    expect(next.actName).toBe(FORM_14_RJ_ACT_NAME);
    expect(next.noticeLine).toBe(FORM_14_RJ_NOTICE_LINE);
    expect(next.form14RJHoursOfWorkLayout).toBe(true);
    expect(next.form14RJCenterHeader).toBe(true);
    expect(next.headerTextAlign).toBe('center');
  });

  test('defaults Whether young person or not to Young', () => {
    expect(isForm14RajasthanYoungPersonHeader('Whether young person or not')).toBe(true);
    const rows = [{ 'Whether young person or not': '', 'Total hours worked during the month': '' }];
    applyForm14RajasthanAutofillDefaultsToRows(rows, [
      'Whether young person or not',
      'Total hours worked during the month',
    ]);
    expect(rows[0]['Whether young person or not']).toBe(FORM_14_RJ_YOUNG_DEFAULT);
  });

  test('Total hours worked during the month = Paid_days × 8', () => {
    expect(isForm14RajasthanTotalHoursHeader('Total hours worked during the month')).toBe(true);
    expect(computeForm14RajasthanTotalHours(22)).toBe('176');
    expect(readForm14RajasthanPaidDays({ Paid_days: 22 })).toBe('22');
    const rows = [{ 'Whether young person or not': '', 'Total hours worked during the month': '' }];
    applyForm14RajasthanPaidDaysToRows(
      rows,
      [{ Employee: { EmpID: 'E1' } }],
      ['Whether young person or not', 'Total hours worked during the month'],
      {
        resolvePayrollRow: () => ({ Paid_days: 22 }),
      }
    );
    expect(rows[0]['Total hours worked during the month']).toBe('176');
  });

  test('ExcelJS Form 14 RJ export keeps model table headers (not Young/Nil)', async () => {
    const ExcelJS = require('exceljs');
    const template = new ExcelJS.Workbook();
    const ws = template.addWorksheet('Sheet1');
    ws.getCell(2, 2).value = 'The Rajasthan Shops and Establishments Rules, 1959';
    ws.getCell(3, 2).value = 'FORM 14';
    ws.getCell(4, 2).value = 'Rule 22 - sub rule 3';
    ws.getCell(5, 2).value = 'Record of the Hours of Work of Persons Employed';
    ws.getCell(6, 2).value = '(To be used only when Notice in Form 13 is exhibited)';
    ws.getCell(8, 2).value = 'Description of Department (if applicable)';
    ws.getCell(9, 2).value = 'Month---------';
    ws.getCell(9, 6).value = 'Year----------';
    // Corrupt header row like a bad prior download (Young / Nil as headers).
    ws.getCell(11, 2).value = 'Virendra';
    ws.getCell(11, 3).value = 'Young';
    ws.getCell(11, 4).value = 'Total hours worked during the month';
    ws.getCell(11, 5).value = 'Nil';
    ws.getCell(11, 6).value = 'Nil';
    ws.getCell(11, 7).value = 'Nil';
    ws.getCell(14, 2).value =
      '*This column need not be filled by Commercial Establishments. In case of shops, Residential Hotels, Restaurants and Eating Houses and Theatres and places of public amusement or entertainment, the extent of such overtime on each day shall be recorded in the days column against the employed distinctively in red ink, indicating the time up to which such overtime work was taken from the employee';
    const templateBuf = await template.xlsx.writeBuffer();

    const {
      buildForm14RajasthanWorkbookWithTemplateStyles,
      FORM_14_RJ_TABLE_HEADERS,
      applyForm14RajasthanMonthYearToHeaderData,
    } = require('./form14Rajasthan');
    const headerFormData = applyForm14RajasthanMonthYearToHeaderData({}, 'August', '2026');
    const { blob } = await buildForm14RajasthanWorkbookWithTemplateStyles({
      templateArrayBuffer: templateBuf,
      mappedData: [
        {
          'Name of Persons Employed': 'Virendra',
          'Whether young person or not': '',
          'Total hours worked during the month': '176',
        },
      ],
      headersToUse: [...FORM_14_RJ_TABLE_HEADERS],
      headerFormData,
      formFileName: 'Form_14_RJ_-_Rajasthan.xlsx',
      parsedHeaderRowIndex: 10,
      parsedDataStartIndex: 10,
    });
    expect(blob).toBeTruthy();
    const outBuf = await new Response(blob).arrayBuffer();
    const reloaded = new ExcelJS.Workbook();
    await reloaded.xlsx.load(outBuf);
    const outWs = reloaded.worksheets[0];
    expect(String(outWs.getCell(11, 2).value || '')).toBe(FORM_14_RJ_TABLE_HEADERS[0]);
    expect(String(outWs.getCell(11, 3).value || '')).toBe(FORM_14_RJ_TABLE_HEADERS[1]);
    expect(String(outWs.getCell(11, 4).value || '')).toBe(FORM_14_RJ_TABLE_HEADERS[2]);
    expect(String(outWs.getCell(11, 5).value || '')).toBe(FORM_14_RJ_TABLE_HEADERS[3]);
    expect(String(outWs.getCell(11, 6).value || '')).toBe(FORM_14_RJ_TABLE_HEADERS[4]);
    expect(String(outWs.getCell(11, 7).value || '')).toBe(FORM_14_RJ_TABLE_HEADERS[5]);
    expect(String(outWs.getCell(12, 2).value || '')).toBe('Virendra');
    expect(String(outWs.getCell(12, 3).value || '')).toBe('Young');
    expect(String(outWs.getCell(12, 4).value || '')).toBe('176');
    expect(String(outWs.getCell(12, 5).value || '')).toMatch(/^Nil$/i);
    expect(outWs.getCell(12, 2).alignment?.horizontal).toBe('left');
    expect(outWs.getCell(12, 3).alignment?.horizontal).toBe('left');
    expect(outWs.getCell(12, 4).alignment?.horizontal).toBe('left');
    expect(outWs.getCell(12, 5).alignment?.horizontal).toBe('left');

    // Finalize normally right-aligns numbers — Form 14 must stay left.
    const { applyStatutoryDownloadContentAlignment } = require('../../utils/excelTableBorders');
    applyStatutoryDownloadContentAlignment(outWs);
    expect(outWs.getCell(12, 2).alignment?.horizontal).toBe('left');
    expect(outWs.getCell(12, 4).alignment?.horizontal).toBe('left');
    expect(outWs.getCell(12, 5).alignment?.horizontal).toBe('left');
  });

  test('ExcelJS Form 14 RJ export reloads without SheetJS round-trip', async () => {
    const ExcelJS = require('exceljs');
    const template = new ExcelJS.Workbook();
    const ws = template.addWorksheet('Sheet1');
    ws.getCell(2, 1).value = 'The Rajasthan Shops and Establishments Rules, 1959';
    ws.getCell(3, 1).value = 'FORM 14';
    ws.getCell(4, 1).value = 'Rule 22 - sub rule 3';
    ws.getCell(5, 1).value = 'Record of the Hours of Work of Persons Employed';
    ws.getCell(8, 1).value = 'Description of Department (if applicable)';
    ws.getCell(9, 1).value = 'Month---------';
    ws.getCell(9, 4).value = 'Year----------';
    ws.getCell(10, 1).value = 'Name of Persons Employed';
    ws.getCell(10, 2).value = 'Whether young person or not';
    ws.getCell(10, 3).value = 'Total hours worked during the month';
    ws.getCell(14, 1).value =
      '*This column need not be filled by Commercial Establishments. In case of shops, Residential Hotels, Restaurants and Eating Houses and Theatres and places of public amusement or entertainment, the extent of such overtime on each day shall be recorded in the days column against the employed distinctively in red ink, indicating the time up to which such overtime work was taken from the employee';
    ws.getCell(16, 1).value = 'Note: Entries relating\nto any day must be\nmade on that day';
    const templateBuf = await template.xlsx.writeBuffer();

    const {
      buildForm14RajasthanWorkbookWithTemplateStyles,
      FORM_14_RJ_FOOTER_NOTE,
      FORM_14_RJ_DAY_ENTRIES_NOTE,
      applyForm14RajasthanMonthYearToHeaderData,
    } = require('./form14Rajasthan');
    const headerFormData = applyForm14RajasthanMonthYearToHeaderData({}, 'August', '2026');
    const { blob } = await buildForm14RajasthanWorkbookWithTemplateStyles({
      templateArrayBuffer: templateBuf,
      mappedData: [
        {
          'Name of Persons Employed': 'Virendra',
          'Whether young person or not': '',
          'Total hours worked during the month': '176',
        },
      ],
      headersToUse: [
        'Name of Persons Employed',
        'Whether young person or not',
        'Total hours worked during the month',
      ],
      headerFormData,
      formFileName: 'Form_14_RJ_-_Rajasthan.xlsx',
      // Bad modal hint that previously wiped the header row — must be ignored.
      parsedHeaderRowIndex: 9,
      parsedDataStartIndex: 9,
    });
    expect(blob).toBeTruthy();
    const outBuf = await new Response(blob).arrayBuffer();
    const reloaded = new ExcelJS.Workbook();
    await reloaded.xlsx.load(outBuf);
    const outWs = reloaded.worksheets[0];
    expect(String(outWs.getCell(3, 1).value || '')).toMatch(/FORM\s*14/i);
    expect(outWs.getCell(3, 1).alignment?.horizontal).toBe('center');
    expect(String(outWs.getCell(2, 1).value || '')).toMatch(/Rajasthan Shops/i);
    expect(outWs.getCell(2, 1).alignment?.horizontal).toBe('center');
    expect(String(outWs.getCell(9, 1).value || '')).toMatch(/Month\s*:\s*August/i);
    expect(String(outWs.getCell(9, 4).value || '')).toMatch(/Year\s*:\s*2026/i);
    // Table column headers must remain (not overwritten by employee rows).
    expect(String(outWs.getCell(10, 1).value || '')).toMatch(/Name of Persons Employed/i);
    expect(String(outWs.getCell(10, 2).value || '')).toMatch(/young person/i);
    expect(String(outWs.getCell(11, 1).value || '')).toBe('Virendra');
    expect(String(outWs.getCell(11, 2).value || '')).toBe('Young');
    expect(String(outWs.getCell(11, 3).value || '')).toBe('176');
    expect(outWs.getCell(11, 1).alignment?.horizontal).toBe('left');
    expect(outWs.getCell(11, 2).alignment?.horizontal).toBe('left');
    expect(outWs.getCell(11, 3).alignment?.horizontal).toBe('left');
    expect(outWs.getCell(9, 1).alignment?.horizontal).toBe('left');
    expect(Number(outWs.getColumn(1).width || 0)).toBeGreaterThanOrEqual(28);
    expect(Number(outWs.getColumn(2).width || 0)).toBeGreaterThanOrEqual(20);
    expect(Number(outWs.getColumn(3).width || 0)).toBeGreaterThanOrEqual(20);
    const footerText = String(outWs.getCell(12, 1).value || '');
    expect(footerText).toMatch(/need not be filled by Commercial Establishments/i);
    expect(footerText).toMatch(/red ink/i);
    expect(footerText).not.toMatch(/\n/);
    expect(footerText.length).toBeGreaterThan(80);
    expect(FORM_14_RJ_FOOTER_NOTE).toMatch(/red ink/i);
    const dayNote = String(outWs.getCell(13, 1).value || '');
    expect(dayNote).toBe(FORM_14_RJ_DAY_ENTRIES_NOTE);
    expect(dayNote).not.toMatch(/\n/);
    // Prefer a wide merge so the note reads as one paragraph like the template model.
    const merges = outWs.model?.merges || [];
    expect(merges.some((m) => String(m).startsWith('A12:'))).toBe(true);
  });
});
