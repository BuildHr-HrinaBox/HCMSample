import ExcelJS from 'exceljs';
import { buildFormPMaharashtraPerEmployeeDownload } from './formPMaharashtra';
import { FORM_OGJ_PERIOD_PARENT } from './formOGJGujarat';

const ESTABLISHMENT_TEXT =
  'Torrent Solargen Ltd (Lohara) Gut No ;206,207 Salgardivti Lohara Road Village,Hippargarava';

const formPMhHeaderFormData = {
  statutory_establishment_name_address: ESTABLISHMENT_TEXT,
};

const FROM_HEADER = `${FORM_OGJ_PERIOD_PARENT}_From`;
const TILL_HEADER = `${FORM_OGJ_PERIOD_PARENT}_Till`;

async function makeFormPMaharashtraNoticeTemplateBuffer() {
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

describe('Form P Maharashtra per-employee notice ZIP', () => {
  test('buildFormPMaharashtraPerEmployeeDownload emits one xlsx per employee in a ZIP', async () => {
    const buf = await makeFormPMaharashtraNoticeTemplateBuffer();
    const mapped = [
      {
        __employeeLookupName: 'A B',
        __employeeLookupId: 'E1',
        'Sr. No.': '1',
        'Number of accumulated leave': '2',
        [FROM_HEADER]: '31 Aug 2026',
        [TILL_HEADER]: '01 Sept 2026',
      },
      {
        __employeeLookupName: 'C D',
        __employeeLookupId: 'E2',
        'Sr. No.': '2',
        'Number of accumulated leave': '1',
        [FROM_HEADER]: '07 Sept 2026',
        [TILL_HEADER]: '07 Sept 2026',
      },
    ];
    const employees = [
      { FirstName: 'A', LastName: 'B', Employee_ID: 'E1' },
      { FirstName: 'C', LastName: 'D', Employee_ID: 'E2' },
    ];

    const { blob, fileName } = await buildFormPMaharashtraPerEmployeeDownload({
      templateArrayBuffer: buf,
      mappedData: mapped,
      headersToUse: ['Sr. No.', 'Number of accumulated leave', FROM_HEADER, TILL_HEADER],
      formFileName: 'Form_P_-_Maharashtra.xlsx',
      employeesOverride: employees,
      headerFormData: formPMhHeaderFormData,
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
    expect(names).toHaveLength(2);
    expect(names.some((n) => /Form_P_Maharashtra_A_B\.xlsx$/i.test(n))).toBe(true);
    expect(names.some((n) => /Form_P_Maharashtra_C_D\.xlsx$/i.test(n))).toBe(true);
  });
});
