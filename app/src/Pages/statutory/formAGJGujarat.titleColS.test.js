import ExcelJS from 'exceljs';
import {
  buildFormAGJGujaratWorkbookWithTemplateStyles,
  writeFormAGJGujaratTitleBandsInColumnS,
} from './formAGJGujarat';

async function buildTemplateWithTitlesInEF() {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('FORM A');
  // Wide merge past S (E2:Z2) — write to S must not land on master E.
  ws.mergeCells(2, 5, 2, 26); // E2:Z2
  ws.getCell(2, 5).value = 'FORM A';
  ws.getCell(2, 5).alignment = { horizontal: 'center' };
  ws.mergeCells(4, 5, 4, 26); // E4:Z4
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

function assertTitlesOnlyInColumnS(ws) {
  for (let r = 1; r <= 8; r += 1) {
    for (let c = 1; c <= 18; c += 1) {
      const v = String(ws.getCell(r, c).value || '');
      expect(v).not.toMatch(/FORM A|FORMAT OF EMPLOYEE/i);
    }
  }
  const sTexts = [];
  for (let r = 1; r <= 8; r += 1) {
    const v = String(ws.getCell(r, 19).value || '').trim();
    if (v) sTexts.push(v);
  }
  expect(sTexts.some((t) => /^FORM A$/i.test(t))).toBe(true);
  expect(sTexts.some((t) => /FORMAT OF EMPLOYEE/i.test(t))).toBe(true);

  const merges = Array.isArray(ws.model?.merges) ? ws.model.merges : [];
  merges.forEach((range) => {
    const parts = String(range || '').split(':');
    if (parts.length !== 2) return;
    const start = parts[0].match(/^([A-Z]+)(\d+)$/i);
    const end = parts[1].match(/^([A-Z]+)(\d+)$/i);
    if (!start || !end) return;
    const r1 = parseInt(start[2], 10);
    const r2 = parseInt(end[2], 10);
    if (r2 < 1 || r1 > 8) return;
    // No leftover title merges that span into E–P.
    expect(String(range)).not.toMatch(/^[E-R]\d+:[E-R]\d+$/i);
  });
}

describe('Form A GJ titles in column S', () => {
  test('writeFormAGJGujaratTitleBandsInColumnS clears E–F and writes S', async () => {
    const buf = await buildTemplateWithTitlesInEF();
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buf);
    const ws = wb.worksheets[0];

    writeFormAGJGujaratTitleBandsInColumnS(ws, {
      title: 'FORM A',
      subtitle: 'FORMAT OF EMPLOYEE/ WORKMAN/ WORKER',
    });

    assertTitlesOnlyInColumnS(ws);
    expect(String(ws.getCell(2, 19).value || '')).toMatch(/FORM A/i);
    expect(String(ws.getCell(4, 19).value || '')).toMatch(/FORMAT OF EMPLOYEE/i);
    expect(String(ws.getCell(2, 5).value || '')).toBe('');
    expect(String(ws.getCell(4, 5).value || '')).toBe('');
  });

  test('download places titles in column S not E–F', async () => {
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
    assertTitlesOnlyInColumnS(outWb.worksheets[0]);
  });
});
