import ExcelJS from 'exceljs';
import {
  applyFormCLabourWelfareYearToFormHeader,
  applyFormCLabourWelfareYearToHeaderLabel,
  applyFormCLabourWelfareYearToHeaders,
  applyFormCLabourWelfareYearToSubtitle,
  ensureFormCTamilNaduLwfEstablishmentLayout,
  ensureFormCTamilNaduLwfTableAlignment,
  ensureFormCTamilNaduLwfTitleLayout,
  remapFormCLabourWelfareRowsToHeaders,
  resolveFormCLabourWelfareYear,
} from './formCTamilNaduLwf';

describe('Form C Tamil Nadu LWF year helpers', () => {
  it('resolves year from due date, then selected month, then current year', () => {
    expect(resolveFormCLabourWelfareYear('April', { dueDate: '2024-06-15' })).toBe(2024);
    expect(resolveFormCLabourWelfareYear('April 2025', {})).toBe(2025);
    expect(resolveFormCLabourWelfareYear('', {}, 'for the year - 2023')).toBe(2023);
  });

  it('appends corresponding year to subtitle', () => {
    expect(applyFormCLabourWelfareYearToSubtitle('', 2026)).toBe(
      'Register of Fines and Unpaid Accumulations for the year - 2026'
    );
    expect(
      applyFormCLabourWelfareYearToSubtitle(
        'Register of Fines and Unpaid Accumulations for the year',
        2024
      )
    ).toBe('Register of Fines and Unpaid Accumulations for the year - 2024');
    expect(
      applyFormCLabourWelfareYearToSubtitle(
        'Register of Fines and Unpaid Accumulations for the year - 2014',
        2026
      )
    ).toBe('Register of Fines and Unpaid Accumulations for the year - 2026');
  });

  it('rewrites truncated quarter header years', () => {
    expect(applyFormCLabourWelfareYearToHeaderLabel('Quarter ending 31-March-14 (2)', 2026)).toBe(
      'Quarter ending 31-March-2026 (2)'
    );
    expect(applyFormCLabourWelfareYearToHeaderLabel('Quarter ending 30-June-14 (3)', 2024)).toBe(
      'Quarter ending 30-June-2024 (3)'
    );
    expect(
      applyFormCLabourWelfareYearToHeaderLabel('Quarter ending 30-September-2014 (4)', 2026)
    ).toBe('Quarter ending 30-September-2026 (4)');
    expect(applyFormCLabourWelfareYearToHeaderLabel('Details of fines (1)', 2026)).toBe(
      'Details of fines (1)'
    );
  });

  it('rewrites all quarter headers in a list', () => {
    const headers = [
      'Details of fines and unpaid accumulations (1)',
      'Quarter ending 31-March-14 (2)',
      'Quarter ending 30-June-14 (3)',
      'Quarter ending 30-September-14 (4)',
      'Quarter ending 31-December-14 (5)',
    ];
    expect(applyFormCLabourWelfareYearToHeaders(headers, 2026)).toEqual([
      'Details of fines and unpaid accumulations (1)',
      'Quarter ending 31-March-2026 (2)',
      'Quarter ending 30-June-2026 (3)',
      'Quarter ending 30-September-2026 (4)',
      'Quarter ending 31-December-2026 (5)',
    ]);
  });

  it('updates form header subtitle with year', () => {
    const next = applyFormCLabourWelfareYearToFormHeader(
      { title: 'FORM C', subtitle: 'Register of Fines and Unpaid Accumulations for the year' },
      2026
    );
    expect(next.subtitle).toBe('Register of Fines and Unpaid Accumulations for the year - 2026');
  });

  it('remaps row keys when quarter headers change', () => {
    const oldHeaders = [
      'Details of fines (1)',
      'Quarter ending 31-March-14 (2)',
      'Quarter ending 30-June-14 (3)',
    ];
    const newHeaders = applyFormCLabourWelfareYearToHeaders(oldHeaders, 2026);
    const rows = [
      {
        'Details of fines (1)': '1. Total Realisations',
        'Quarter ending 31-March-14 (2)': 'Nil',
        'Quarter ending 30-June-14 (3)': '',
      },
    ];
    const remapped = remapFormCLabourWelfareRowsToHeaders(rows, oldHeaders, newHeaders);
    expect(remapped[0]['Quarter ending 31-March-2026 (2)']).toBe('Nil');
    expect(remapped[0]['Details of fines (1)']).toBe('1. Total Realisations');
    expect(remapped[0]['Quarter ending 31-March-14 (2)']).toBeUndefined();
  });
});

describe('Form C Tamil Nadu LWF download alignment', () => {
  async function buildFormCSheet() {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Form C');
    ws.getCell(2, 1).value = 'Form-C';
    ws.getCell(2, 1).alignment = { horizontal: 'left' };
    ws.getCell(3, 1).value = '[See rule 29 of the Tamil Nadu Labour Welfare fund Rules.1973]';
    ws.getCell(4, 1).value = 'Register of Fines and Unpaid Accumulations for the year - 2024';
    ws.getCell(5, 1).value = 'Name of the Establishment : TN-Palani, Vayona Energy Pvt Ltd';
    ws.getCell(11, 1).value = 'Details of Fines and Unpaid Accumulations (1)';
    ws.getCell(11, 2).value = 'Quarter ending 31-March-2024 (2)';
    ws.getCell(11, 3).value = 'Quarter ending 30-June-2024 (3)';
    ws.getCell(11, 4).value = 'Quarter ending 30-September-2024 (4)';
    ws.getCell(11, 5).value = 'Quarter ending 31-December-2024 (5)';
    ws.mergeCells(12, 1, 14, 1);
    ws.getCell(12, 1).value = '1. Total Realisations under fines';
    ws.getCell(12, 2).value = 'Nil';
    ws.mergeCells(15, 1, 17, 1);
    ws.getCell(15, 1).value = '2. Total amount being unpaid Accumulations*of--- (I) Basic Wages';
    ws.getCell(15, 2).value = 'Nil';
    return ws;
  }

  it('centers Form-C / See rule / Register title across the table width', async () => {
    const ws = await buildFormCSheet();
    const ok = ensureFormCTamilNaduLwfTitleLayout(ws, { colFrom: 1, colTo: 5, force: true });
    expect(ok).toBe(true);
    expect(ws.getCell(2, 1).value).toBe('Form-C');
    expect(ws.getCell(2, 1).alignment.horizontal).toBe('center');
    expect(ws.getCell(3, 1).alignment.horizontal).toBe('center');
    expect(ws.getCell(4, 1).alignment.horizontal).toBe('center');
    expect(ws.getCell(4, 1).font?.bold).toBe(true);
    const merges = ws.model.merges || [];
    expect(merges.some((m) => /^A2:E2$/i.test(m))).toBe(true);
    expect(merges.some((m) => /^A4:E4$/i.test(m))).toBe(true);
  });

  it('left-aligns establishment across the form and strips the label prefix', async () => {
    const ws = await buildFormCSheet();
    const ok = ensureFormCTamilNaduLwfEstablishmentLayout(ws, {
      headerRow: 11,
      colFrom: 1,
      colTo: 5,
      establishmentText: 'TN-Palani, Vayona Energy Pvt Ltd, 274 A, Rani mangammal main road',
    });
    expect(ok).toBe(true);
    expect(String(ws.getCell(5, 1).value)).toContain('TN-Palani, Vayona Energy Pvt Ltd');
    expect(String(ws.getCell(5, 1).value)).not.toMatch(/^Name of the Establishment/i);
    expect(ws.getCell(5, 1).alignment.horizontal).toBe('left');
    expect(ws.getCell(5, 1).alignment.wrapText).toBe(true);
  });

  it('centers Nil inside vertical merges matching each category block', async () => {
    const ws = await buildFormCSheet();
    const ok = ensureFormCTamilNaduLwfTableAlignment(ws, {
      headerRow: 11,
      dataStartRow: 12,
      colFrom: 1,
      colTo: 5,
      quarterCols: [2, 3, 4, 5],
      categorySlots: [
        { row: 12, key: 'realisations' },
        { row: 15, key: 'basic' },
      ],
    });
    expect(ok).toBe(true);
    expect(ws.getCell(11, 1).alignment.horizontal).toBe('center');
    expect(ws.getCell(11, 2).alignment.wrapText).toBe(true);
    expect(ws.getCell(12, 1).alignment.horizontal).toBe('left');
    expect(ws.getCell(12, 2).value).toBe('Nil');
    expect(ws.getCell(12, 2).alignment.horizontal).toBe('center');
    expect(ws.getCell(12, 2).alignment.vertical).toBe('middle');
    expect(ws.getCell(15, 2).value).toBe('Nil');
    const merges = ws.model.merges || [];
    expect(merges.some((m) => /^B12:B14$/i.test(m))).toBe(true);
    expect(merges.some((m) => /^B15:B17$/i.test(m))).toBe(true);
  });
});
