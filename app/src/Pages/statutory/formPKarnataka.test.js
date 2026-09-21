import ExcelJS from 'exceljs';
import {
  buildFormPKarnatakaPerEmployeeDownload,
  writeFormPKarnatakaNoticeIdentity,
} from './formPKarnataka';
import { FORM_OGJ_PERIOD_PARENT } from './formOGJGujarat';

const ESTABLISHMENT_TEXT =
  '110/33KV Substation, SF-490/2, Chanderki Village, Gurumatkal Taluk, Yadgir District, Karnataka-585214';

const formPKaHeaderFormData = {
  statutory_establishment_name_address: ESTABLISHMENT_TEXT,
};

const FROM_HEADER = `${FORM_OGJ_PERIOD_PARENT}_From`;
const TILL_HEADER = `${FORM_OGJ_PERIOD_PARENT}_Till`;

async function makeFormPKarnatakaNoticeTemplateBuffer() {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('FORM P');
  ws.getCell(3, 3).value = "Form – 'P'";
  ws.getCell(4, 3).value = '(See rule 20)';
  ws.getCell(5, 3).value = 'NOTICE OF MAXIMUM LEAVE ACCUMULATED';
  ws.getCell(6, 3).value = 'Name and address of the establishment';
  ws.getCell(7, 3).value = 'Name of the Authorised person / Manager';
  ws.getCell(8, 3).value = 'To,';
  ws.getCell(9, 3).value = 'Shri/Smt. ...................... (Name of worker)';
  ws.getCell(10, 3).value = 'Address......................';
  ws.getCell(14, 3).value = 'Sr. No.';
  ws.getCell(14, 4).value = 'Number of accumulated leave';
  ws.mergeCells(14, 5, 14, 6);
  ws.getCell(14, 5).value = 'Period for which leave is accumulated';
  ws.getCell(15, 5).value = 'From';
  ws.getCell(15, 6).value = 'Till';
  return wb.xlsx.writeBuffer();
}

function worksheetPlainText(ws) {
  const parts = [];
  ws.eachRow((row) => {
    row.eachCell((cell) => {
      const value = cell.value;
      if (value == null || value === '') return;
      parts.push(String(value));
    });
  });
  return parts.join(' | ');
}

describe('Form P Karnataka per-employee notice ZIP', () => {
  test('buildFormPKarnatakaPerEmployeeDownload emits one xlsx per employee in a ZIP', async () => {
    const buf = await makeFormPKarnatakaNoticeTemplateBuffer();
    const mapped = [
      {
        __employeeLookupName: 'A B',
        __employeeLookupId: 'E1',
        'Sr. No.': '1',
        'Number of accumulated leave': '11',
        [FROM_HEADER]: '29 May 2026',
        [TILL_HEADER]: '10 Jun 2026',
      },
      {
        __employeeLookupName: 'C D',
        __employeeLookupId: 'E2',
        'Sr. No.': '2',
        'Number of accumulated leave': '4',
        [FROM_HEADER]: '01 Jun 2026',
        [TILL_HEADER]: '04 Jun 2026',
      },
      {
        __employeeLookupName: 'E F',
        __employeeLookupId: 'E3',
        'Sr. No.': '3',
        'Number of accumulated leave': '2',
        [FROM_HEADER]: '08 Jul 2026',
        [TILL_HEADER]: '09 Jul 2026',
      },
    ];
    const employees = [
      {
        FirstName: 'A',
        LastName: 'B',
        Employee_ID: 'E1',
        Address: 'House 1, Bengaluru',
      },
      {
        FirstName: 'C',
        LastName: 'D',
        Employee_ID: 'E2',
        Address: 'House 2, Mysuru',
      },
      {
        FirstName: 'E',
        LastName: 'F',
        Employee_ID: 'E3',
        Address: 'House 3, Hubballi',
      },
    ];

    const { blob, fileName } = await buildFormPKarnatakaPerEmployeeDownload({
      templateArrayBuffer: buf,
      mappedData: mapped,
      headersToUse: [
        'Sr. No.',
        'Number of accumulated leave',
        FROM_HEADER,
        TILL_HEADER,
      ],
      formFileName: 'Form_P_-_Karnataka.xlsx',
      employeesOverride: employees,
      headerFormData: formPKaHeaderFormData,
    });

    expect(fileName).toMatch(/_Employees\.zip$/i);
    expect(blob.type).toMatch(/zip/i);

    const JSZip = (await import('jszip')).default;
    const zipBytes =
      typeof blob?.arrayBuffer === 'function'
        ? await blob.arrayBuffer()
        : await new Response(blob).arrayBuffer();
    const zip = await JSZip.loadAsync(zipBytes);
    const names = Object.keys(zip.files).filter((n) => !zip.files[n].dir);
    expect(names).toHaveLength(3);
    expect(names.some((n) => /Form_P_Karnataka_A_B\.xlsx$/i.test(n))).toBe(true);
    expect(names.some((n) => /Form_P_Karnataka_C_D\.xlsx$/i.test(n))).toBe(true);
    expect(names.some((n) => /Form_P_Karnataka_E_F\.xlsx$/i.test(n))).toBe(true);

    const firstBytes = await zip.files[names.find((n) => /A_B/i.test(n))].async('arraybuffer');
    const outWb = new ExcelJS.Workbook();
    await outWb.xlsx.load(firstBytes);
    const outWs = outWb.worksheets[0];
    const text = worksheetPlainText(outWs);
    expect(text).toMatch(/Shri\/Smt\.\s+A B/i);
    expect(text).not.toMatch(/Name of worker/i);
    expect(outWs.getCell(9, 3).alignment?.wrapText).not.toBe(true);
    expect(String(outWs.getCell(6, 3).value)).toMatch(/Name and address of the establishment\s*:/i);
    expect(String(outWs.getCell(6, 3).value)).toMatch(/110\/33KV Substation/i);
    expect(String(outWs.getCell(10, 3).value)).toBe(`Address: ${ESTABLISHMENT_TEXT}`);
    expect(text).not.toMatch(/House 1, Bengaluru/i);
    expect(text).not.toMatch(/House 2, Mysuru/i);
    expect(String(outWs.getCell(16, 4).value)).toBe('11');
    expect(String(outWs.getCell(16, 5).value)).toMatch(/29 May 2026/);
    expect(String(outWs.getCell(16, 6).value)).toMatch(/10 Jun 2026/);
    expect(outWs.getCell(17, 4).value == null || String(outWs.getCell(17, 4).value).trim() === '').toBe(
      true
    );
    expect(text).not.toMatch(/House 2, Mysuru/i);
  });

  test('writeFormPKarnatakaLeaveTableRow clears shift-time residue and stamps leave dates', async () => {
    const { writeFormPKarnatakaLeaveTableRow } = await import('./formPKarnataka');
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('FORM P');
    ws.getCell(14, 3).value = 'Sr. No.';
    ws.getCell(14, 4).value = 'Number of accumulated leave';
    ws.mergeCells(14, 5, 14, 6);
    ws.getCell(14, 5).value = 'Period for which leave is accumulated';
    ws.getCell(15, 5).value = 'From';
    ws.getCell(15, 6).value = 'Till';
    ws.getCell(16, 3).value = '1';
    ws.getCell(16, 4).value = '';
    ws.getCell(16, 5).value = '09:00 AM';
    ws.getCell(16, 6).value = '05:00 PM';
    ws.getCell(17, 5).value = '09:00 AM';
    ws.getCell(17, 6).value = '05:00 PM';

    const written = writeFormPKarnatakaLeaveTableRow(
      ws,
      {
        'Number of accumulated leave': '2',
        'Period for which leave is accumulated_From': '31 Aug 2026',
        'Period for which leave is accumulated_Till': '01 Sept 2026',
      },
      {}
    );
    expect(written).toBe(1);
    expect(String(ws.getCell(16, 4).value)).toBe('2');
    expect(String(ws.getCell(16, 5).value)).toMatch(/31 Aug 2026/);
    expect(String(ws.getCell(16, 6).value)).toMatch(/01 Sept 2026/);
    expect(String(ws.getCell(17, 5).value || '')).toBe('');
    expect(String(ws.getCell(17, 6).value || '')).toBe('');
  });

  test('single employee download stays a single xlsx (not a ZIP)', async () => {
    const buf = await makeFormPKarnatakaNoticeTemplateBuffer();
    const { blob, fileName } = await buildFormPKarnatakaPerEmployeeDownload({
      templateArrayBuffer: buf,
      mappedData: [
        {
          __employeeLookupName: 'Vijeesh Vijayan',
          'Number of accumulated leave': '11',
          [FROM_HEADER]: '29 May 2026',
          [TILL_HEADER]: '10 Jun 2026',
        },
      ],
      headersToUse: ['Sr. No.', 'Number of accumulated leave', FROM_HEADER, TILL_HEADER],
      formFileName: 'Form_P_-_Karnataka.xlsx',
      headerFormData: formPKaHeaderFormData,
      employeesOverride: [
        {
          FirstName: 'Vijeesh',
          LastName: 'Vijayan',
          Employee_ID: 'VE0761',
          'Permanent Address': 'Gurumatkal, Karnataka',
        },
      ],
    });
    expect(fileName).toMatch(/Form_P_Karnataka_Vijeesh_Vijayan\.xlsx$/i);
    expect(blob.type).toMatch(/spreadsheetml/i);

    const bytes =
      typeof blob?.arrayBuffer === 'function'
        ? await blob.arrayBuffer()
        : await new Response(blob).arrayBuffer();
    const outWb = new ExcelJS.Workbook();
    await outWb.xlsx.load(bytes);
    const text = worksheetPlainText(outWb.worksheets[0]);
    expect(text).toMatch(/Shri\/Smt\.\s+Vijeesh Vijayan/i);
    expect(text).not.toMatch(/Name of worker/i);
    expect(text).toMatch(/110\/33KV Substation/i);
    expect(text).toMatch(/Address: 110\/33KV Substation/i);
    expect(text).not.toMatch(/Gurumatkal, Karnataka/i);
  });

  test('writeFormPKarnatakaNoticeIdentity fills establishment and copies it to Address row', () => {
    const ws = {
      rowCount: 12,
      getCell(r, c) {
        const key = `${r}:${c}`;
        if (!this._cells) this._cells = {};
        if (!this._cells[key]) this._cells[key] = { value: '', alignment: {} };
        return this._cells[key];
      },
      getRow() {
        return { height: 18 };
      },
    };
    ws.getCell(6, 3).value = 'Name and address of the establishment';
    ws.getCell(7, 3).value = 'Name of the Authorised person / Manager';
    ws.getCell(8, 3).value = 'To,';
    ws.getCell(9, 3).value = 'Shri/Smt. ...................... (Name of worker)';
    ws.getCell(10, 3).value = 'Address......................';
    const written = writeFormPKarnatakaNoticeIdentity(
      ws,
      { FirstName: 'Sornaraja', LastName: 'Gurusamy', Address: 'Worker Lane 9' },
      {},
      { establishmentText: ESTABLISHMENT_TEXT }
    );
    expect(written).toBeGreaterThan(0);
    expect(String(ws.getCell(6, 3).value)).toBe(
      `Name and address of the establishment : ${ESTABLISHMENT_TEXT}`
    );
    expect(ws.getCell(6, 3).alignment.wrapText).toBe(false);
    expect(String(ws.getCell(9, 3).value)).toBe('Shri/Smt. Sornaraja Gurusamy');
    expect(ws.getCell(9, 3).alignment.wrapText).toBe(false);
    expect(String(ws.getCell(10, 3).value)).toBe(`Address: ${ESTABLISHMENT_TEXT}`);
    expect(ws.getCell(10, 3).alignment.wrapText).toBe(false);
    expect(String(ws.getCell(10, 3).value)).not.toMatch(/Worker Lane 9/);
    expect(String(ws.getCell(7, 3).value)).toBe('Name of the Authorised person / Manager');
    expect(String(ws.getCell(8, 3).value)).toBe('To,');
  });
});
