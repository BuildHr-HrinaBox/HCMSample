import ExcelJS from 'exceljs';
import {
  buildFormAGJGujaratWorkbookWithTemplateStyles,
  writeFormAGJGujaratTitleBandsInColumnS,
} from './formAGJGujarat';

function assertFormAGJTitleBandsCleared(ws) {
  for (let r = 1; r <= 15; r += 1) {
    for (let c = 1; c <= 26; c += 1) {
      const v = String(ws.getCell(r, c).value || '').trim();
      if (!v) continue;
      expect(v).not.toMatch(/^SCHEDULE$/i);
      expect(v).not.toMatch(/^FORM A$/i);
      expect(v).not.toMatch(/FORMAT OF EMPLOYEE/i);
    }
  }
  expect(String(ws.getCell(1, 6).value || '')).not.toMatch(/See rule 2/i);
}

async function buildTemplateWithTitlesInEF() {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('FORM A');
  ws.getCell(1, 1).value = 'SCHEDULE';
  ws.getCell(1, 6).value = '[See rule 2(1)]';
  ws.mergeCells(2, 5, 2, 26);
  ws.getCell(2, 5).value = 'FORM A';
  ws.getCell(2, 5).alignment = { horizontal: 'center' };
  ws.mergeCells(4, 5, 4, 26);
  ws.getCell(4, 5).value = 'FORMAT OF EMPLOYEE/ WORKMAN/ WORKER';
  ws.getCell(4, 5).alignment = { horizontal: 'center' };

  const headers = [
    
    'Sl. No.',
    'Employee Code',
    'Name',
    'Surname',
    "Father's/Spouse Name",
    'Date of Birth',
    'Nationality',
    'Education Level',
    'Date of Joining',
    'Designation',
    'Category (HS/S/SS/US)',
    'Type of Employment',
    'Mobile',
    'UAN',
    'PAN',
    'ESIC IP',
    'LWF',
    'AADHAAR',
    'Bank A/C No.',
    'Bank',
    'Branch (IFSC)',
    'Present Address',
    'Permanent Address',
  ];
  headers.forEach((h, i) => {
    ws.getCell(15, i + 1).value = h;
    ws.getCell(16, i + 1).value = i + 1;
  });
  return wb.xlsx.writeBuffer();
}

function assertFormAGJTitleBandOrder(ws) {
  expect(String(ws.getCell(1, 5).value || '').trim()).toMatch(/^SCHEDULE$/i);
  expect(String(ws.getCell(2, 5).value || '').trim()).toMatch(/^FORM A$/i);
  expect(String(ws.getCell(3, 5).value || '').trim()).toMatch(/FORMAT OF EMPLOYEE/i);
  expect(String(ws.getCell(4, 5).value || '').trim()).not.toMatch(/FORMAT OF EMPLOYEE/i);
  expect(String(ws.getCell(1, 1).value || '').trim()).not.toMatch(/^SCHEDULE$/i);
  expect(String(ws.getCell(1, 6).value || '')).toMatch(/See rule 2/i);
}

describe('Form A GJ title band rows 1–3', () => {
  test('writeFormAGJGujaratTitleBandsInColumnS orders SCHEDULE / FORM A / FORMAT', async () => {
    const buf = await buildTemplateWithTitlesInEF();
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buf);
    const ws = wb.worksheets[0];

    writeFormAGJGujaratTitleBandsInColumnS(ws, {
      title: 'FORM A',
      subtitle: 'FORMAT OF EMPLOYEE/ WORKMAN/ WORKER',
    });

    assertFormAGJTitleBandOrder(ws);
  });

  test('download clears SCHEDULE / FORM A / FORMAT title band', async () => {
    const templateArrayBuffer = await buildTemplateWithTitlesInEF();
    const { buffer } = await buildFormAGJGujaratWorkbookWithTemplateStyles({
      templateArrayBuffer,
      mappedData: [
        {
          'Sl. No.': 1,
          'Employee Code': 'VE0731',
          Name: 'Rajanish',
          Surname: 'Kumar',
        },
      ],
      headersToUse: [
        'Sl. No.',
        'Employee Code',
        'Name',
        'Surname',
        "Father's/Spouse Name",
        'Date of Birth',
        'Nationality',
        'Education Level',
        'Date of Joining',
        'Designation',
        'Category (HS/S/SS/US)',
        'Type of Employment',
        'Mobile',
        'UAN',
        'PAN',
        'ESIC IP',
        'LWF',
        'AADHAAR',
        'Bank A/C No.',
        'Bank',
        'Branch (IFSC)',
        'Present Address',
        'Permanent Address',
      ],
      parsedFormHeader: {
        title: 'FORM A',
        subtitle: 'FORMAT OF EMPLOYEE/ WORKMAN/ WORKER',
      },
      formFileName: 'Form_A_GJ_-_Gujarat.xlsx',
    });

    const outWb = new ExcelJS.Workbook();
    await outWb.xlsx.load(buffer);
    assertFormAGJTitleBandsCleared(outWb.worksheets[0]);
  });
});
