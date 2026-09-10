import * as XLSX from 'xlsx';
import { statutoryDraftPdfTestUtils } from './statutoryDraftPdf';

const {
  looksLikeFormXIXKarnatakaPdfContext,
  normalizeFormXIXKarnatakaWageSlipPdfMatrix,
  sheetToDenseMatrix,
} = statutoryDraftPdfTestUtils;

describe('Form XIX Karnataka PDF normalization', () => {
  test('keeps Sex and Token lines after principal-employer details', () => {
    const aoa = [
      ['FORM XIX'],
      ['[See rule 78 (1)(b)]'],
      ['Wages Slip'],
      ['Name and address of contractor....... VAYONA ENERGY PRIVATE LIMITED'],
      ['Name and address of establishment in/under which contract is carried on'],
      ['M/s Clean Wind Power Bableshwar Pvt Ltd, 33/220KV Substation, Survey No: 132, Karjol Village, Bijapur Dist - 586113 Karnataka, Bableshwar, Karnataka'],
      ['Nature of work and location of work..............................'],
      ['Name and address of Principal Employer, for the week/fortnight/month.................'],
      ['Token/Ticket No............................. VE1491'],
      ['Sex and identification marks............... Male'],
      [
        'No. of days worked',
        '',
        'Rate of daily wages/piece - rate',
        'No. of units worked in case of piece rate',
        '',
        'Dates on which overtime worked',
        '',
        'Overtime hours and amount of overtime wages',
      ],
      ['1', '', '2', '3', '', '4', '', '5'],
      ['31', '', 'Monthly Wages', '', '', 'NIL', '', 'NIL'],
      ['Gross wages payable', '', '', 'Deductions, any', '', '', '', 'If Actual wages paid'],
      ['71392', '', '', '3464', '', '', '', '67928'],
      ['Signature of the contractor or his Representative'],
    ];
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    const matrix = sheetToDenseMatrix(ws, 'Form_XIX_Karnataka');
    expect(looksLikeFormXIXKarnatakaPdfContext(matrix.metaLines, matrix.rows, matrix.name)).toBe(true);

    const normalized = normalizeFormXIXKarnatakaWageSlipPdfMatrix(
      matrix.rows,
      matrix.colCount,
      matrix.tableStartRow,
      matrix.metaLines
    );
    const text = (normalized.metaLines || []).join('\n').toLowerCase();
    expect(text.indexOf('principal employer')).toBeGreaterThanOrEqual(0);
    expect(text.indexOf('sex and identification marks')).toBeGreaterThan(text.indexOf('principal employer'));
    expect(text.indexOf('token/ticket no')).toBeGreaterThan(text.indexOf('sex and identification marks'));
    expect(text).toMatch(/name and address of contractor\.\s*vayona energy private limited/i);
    expect(normalized.formXIXKarnatakaLayout).toBe(true);
  });

  test('does not treat Gujarat or AP Form XIX as Karnataka', () => {
    expect(
      looksLikeFormXIXKarnatakaPdfContext(
        ['FORM XIX', '[See rule 78 (2)(b)]', 'Wage Slip'],
        [['Name and address of contractor', 'VAYONA'], ['Number of days worked', '31']],
        'Form_XIX_GJ'
      )
    ).toBe(false);
    expect(
      looksLikeFormXIXKarnatakaPdfContext(
        ['FORM XIX', '[See rule 78 (1)(b)]', 'Wage Slip'],
        [['Nature and location of work', 'Site'], ['Net amount of wages paid', '100']],
        'Form_XIX_AP'
      )
    ).toBe(false);
  });

  test('rebuilds official dotted Karnataka template into boxed model header order', () => {
    const rows = [
      ['FORM XIX'],
      ['[See rule 78 (1)(b)]'],
      ['Wages Slip'],
      ['VAYONA ENERGY PRIVATE LIMITED'],
      ['KA-Bableshwar'],
      ['Sex and identification marks............................'],
      ['Male'],
      ['Token/Ticket No............................'],
      ['VE0712'],
      ['Name and address of contractor...............'],
      ['Name and address of establishment in/under which contract is carried on'],
      ['M/s Clean Wind Power Bableshwar Pvt Ltd, 33/220KV Substation, Survey No: 132, Karjol Village, Bijapur Dist - 586113 Karnataka, Bableshwar, Karnataka'],
      ['Nature of work and location of work.............................'],
      ['Name and address of Principal Employer, for the week/fortnight/month'],
      ['M/s Clean Wind Power Bableshwar Pvt Ltd, 33/220KV Substation, Survey No: 132, Karjol Village, Bijapur Dist - 586113 Karnataka, Bableshwar, Karnataka'],
      [
        'No.of days worked',
        'Rate of daily wages/piece - rate',
        'No. of units worked in case of piece rate',
        'Dates on which overtime worked',
        'Overtime hours and amount of overtime wages',
      ],
      ['1', '2', '3', '4', '5'],
      ['31', 'Monthly Wages', '2', 'NIL', 'NIL'],
      ['Gross wages payable', '', 'Deductions, any', '', 'If Actual wages paid'],
      ['74992', '', '3464', '', '71528'],
      ['Signature of the contractor or his Representative'],
    ];
    expect(looksLikeFormXIXKarnatakaPdfContext([], rows, 'Form_XIX_KA')).toBe(true);
    const out = normalizeFormXIXKarnatakaWageSlipPdfMatrix(rows, 5, 15, []);
    const meta = (out.metaLines || []).join('\n');
    expect(out.formXIXKarnatakaLayout).toBe(true);
    expect(meta).toMatch(/^FORM XIX/m);
    expect(meta).toMatch(/\[See rule 78 \(1\)\(b\)\]/);
    expect(meta).toMatch(/Wages Slip/);
    expect(meta).toMatch(/Name and address of contractor\.\s*VAYONA ENERGY PRIVATE LIMITED/);
    expect(meta).toMatch(/Nature of work and location of work\.\s*KA Bableshwar/);
    expect(meta).toMatch(/Sex and identification marks\.\s*Male/);
    expect(meta).toMatch(/Token\/Ticket No\.\s*VE0712/);
    expect(meta.indexOf('Sex and identification marks')).toBeGreaterThan(meta.indexOf('Principal Employer'));
    expect((out.rows || [])[0].join(' ')).toMatch(/No\.?\s*of days worked/i);
    expect((out.rows || [])[1]).toEqual(['1', '2', '3', '4', '5']);
    expect((out.rows || [])[2][0]).toBe('31');
    expect((out.rows || []).some((r) => /gross wages payable/i.test((r || []).join(' ')))).toBe(false);
    expect(out.formXIXKABottomSection?.values).toEqual(['74992', '3464', '71528']);
    expect(out.formXIXKABottomSection?.signatureText).toMatch(/signature of the contractor/i);
  });

  test('does not duplicate bottom-section labels and keeps one signature line', () => {
    const rows = [
      ['FORM XIX'],
      ['Wages Slip'],
      ['No. of days worked', '', 'Rate of daily wages/piece - rate', 'No. of units worked in case of piece rate', '', 'Dates on which overtime worked', '', 'Overtime hours and amount of overtime wages'],
      ['1', '', '2', '3', '', '4', '', '5'],
      ['31', '', 'Monthly Wages', '', '', 'NIL', '', 'NIL'],
      ['Gross wages payable', '', '', 'Deductions, any', '', '', '', 'If Actual wages paid'],
      ['71392', '', '', '3464', '', '', '', '67928'],
      ['Signature of the contractor or his Representative'],
      ['Gross wages payable', '', '', 'Deductions, any', '', '', '', 'If Actual wages paid'],
      ['Signature of the contractor or his Representative'],
    ];
    const out = normalizeFormXIXKarnatakaWageSlipPdfMatrix(rows, 8, 2, []);
    const section = out.formXIXKABottomSection || {};
    const bottomBlob = [
      ...(section.labels || []),
      ...(section.values || []),
      section.signatureText || '',
    ]
      .join('\n')
      .toLowerCase();
    expect((bottomBlob.match(/gross wages payable/g) || []).length).toBe(1);
    expect((bottomBlob.match(/deductions, any/g) || []).length).toBe(1);
    expect((bottomBlob.match(/if actual wages paid/g) || []).length).toBe(1);
    expect((bottomBlob.match(/signature of the contractor or his representative/g) || []).length).toBe(1);
    expect((out.rows || []).some((r) => /gross wages payable/i.test((r || []).join(' ')))).toBe(false);
  });

  test('maps packed single footer value to deductions column when anchors are empty', () => {
    const rows = [
      ['FORM XIX'],
      ['Wages Slip'],
      ['No. of days worked', '', 'Rate of daily wages/piece - rate', 'No. of units worked in case of piece rate', '', 'Dates on which overtime worked', '', 'Overtime hours and amount of overtime wages'],
      ['1', '', '2', '3', '', '4', '', '5'],
      ['31', '', 'Monthly Wages', '', '', 'NIL', '', 'NIL'],
      ['Gross wages payable', '', '', 'Deductions, any', '', '', '', 'If Actual wages paid'],
      ['', '', '', '3464', '', '', '', ''],
      ['Signature of the contractor or his Representative'],
    ];
    const out = normalizeFormXIXKarnatakaWageSlipPdfMatrix(rows, 8, 2, []);
    expect(out.formXIXKABottomSection?.values?.[1]).toBe('3464');
  });

  test('preserves numeric zero in deductions footer value', () => {
    const rows = [
      ['FORM XIX'],
      ['Wages Slip'],
      ['No. of days worked', '', 'Rate of daily wages/piece - rate', 'No. of units worked in case of piece rate', '', 'Dates on which overtime worked', '', 'Overtime hours and amount of overtime wages'],
      ['1', '', '2', '3', '', '4', '', '5'],
      ['31', '', 'Monthly Wages', '', '', 'NIL', '', 'NIL'],
      ['Gross wages payable', '', '', 'Deductions, any', '', '', '', 'If Actual wages paid'],
      ['', '', '', 0, '', '', '', ''],
      ['Signature of the contractor or his Representative'],
    ];
    const out = normalizeFormXIXKarnatakaWageSlipPdfMatrix(rows, 8, 2, []);
    expect(out.formXIXKABottomSection?.values?.[1]).toBe('0');
  });

  test('reads deductions from actual deductions label column', () => {
    const rows = [
      ['FORM XIX'],
      ['Wages Slip'],
      ['No. of days worked', '', 'Rate of daily wages/piece - rate', 'No. of units worked in case of piece rate', '', 'Dates on which overtime worked', '', 'Overtime hours and amount of overtime wages'],
      ['1', '', '2', '3', '', '4', '', '5'],
      ['31', '', 'Monthly Wages', '', '', 'NIL', '', 'NIL'],
      ['', 'Gross wages payable', '', '', 'Deductions, any', '', '', 'If Actual wages paid', ''],
      ['', '', '', '', 0, '', '', '', ''],
      ['Signature of the contractor or his Representative'],
    ];
    const out = normalizeFormXIXKarnatakaWageSlipPdfMatrix(rows, 9, 2, []);
    expect(out.formXIXKABottomSection?.values?.[1]).toBe('0');
  });
});

