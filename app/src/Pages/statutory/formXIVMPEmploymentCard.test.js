import ExcelJS from 'exceljs';
import {
  FORM_XIV_MP_CANONICAL_TABLE_HEADERS,
  detectFormXIVStackedWorkmanLayout,
  resolveFormXIVMPWorkmanFieldPositions,
  writeFormXIVMPHeaderFieldsToWorksheet,
  writeFormXIVMPWorkmanFieldsToWorksheet,
} from './formXIVMPEmploymentCard';

/** Mimic MP Employment Card: ordinals in A, labels in B, values intended for E. */
async function buildSplitOrdinalStackedTemplate() {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Sheet1');
  for (let c = 1; c <= 12; c += 1) {
    ws.getColumn(c).width = 3;
  }
  ws.getCell(3, 6).value = '(See rule 76)';
  ws.getCell(6, 2).value = 'Name and Address of the Establishment in/under which Contract is carried on:';
  ws.getCell(8, 2).value = 'Name and Address of the Contractor';
  ws.getCell(10, 2).value = 'Nature of work and location of work';
  ws.getCell(12, 2).value = 'Name and address of Principal Employer';
  ws.getCell(14, 1).value = 1;
  ws.getCell(14, 2).value = 'Name of the workman';
  ws.getCell(15, 1).value = 2;
  ws.getCell(15, 2).value = 'Serial number in the register of workmen employed';
  ws.getCell(17, 1).value = 3;
  ws.getCell(17, 2).value = 'Nature of employment / designation';
  ws.getCell(19, 1).value = 4;
  ws.getCell(19, 2).value = "Wage rate with particulars or unit, in case of piece of work";
  ws.getCell(21, 1).value = 5;
  ws.getCell(21, 2).value = 'Wage period';
  ws.getCell(23, 1).value = 6;
  ws.getCell(23, 2).value = 'Tenure of employment';
  ws.getCell(25, 1).value = 7;
  ws.getCell(25, 2).value = 'Remarks';
  return { wb, ws };
}

describe('Form XIV MP stacked download alignment', () => {
  it('detects stacked layout when ordinals are split from labels', async () => {
    const { ws } = await buildSplitOrdinalStackedTemplate();
    expect(detectFormXIVStackedWorkmanLayout(ws)).toBe(true);
  });

  it('writes header and workman values into column E, not beside the label', async () => {
    const { ws } = await buildSplitOrdinalStackedTemplate();
    writeFormXIVMPHeaderFieldsToWorksheet(
      ws,
      {
        form_xiv_mp_establishment: 'VAYONA',
        form_xiv_mp_contractor: 'M/s Contractor',
        form_xiv_mp_nature_location: 'MP-Dhar',
        form_xiv_mp_principal_employer: 'Principal Employer',
      },
      { formXIVVariant: 'mp' }
    );

    expect(String(ws.getCell(6, 5).value || '')).toBe('VAYONA');
    expect(String(ws.getCell(6, 3).value || '')).toBe('');
    expect(String(ws.getCell(8, 5).value || '')).toBe('M/s Contractor');
    expect(String(ws.getCell(10, 5).value || '')).toBe('MP-Dhar');
    expect(String(ws.getCell(12, 5).value || '')).toBe('Principal Employer');

    const headers = [...FORM_XIV_MP_CANONICAL_TABLE_HEADERS];
    const layout = resolveFormXIVMPWorkmanFieldPositions(ws, headers);
    expect(layout.stackedLayout).toBe(true);
    expect(layout.valueCol).toBe(5);

    writeFormXIVMPWorkmanFieldsToWorksheet(
      ws,
      {
        [headers[0]]: 'Jeevan',
        [headers[1]]: 'VE1256',
        [headers[2]]: 'Engineer',
      },
      headers,
      layout
    );

    expect(String(ws.getCell(14, 5).value || '')).toBe('Jeevan');
    expect(String(ws.getCell(14, 3).value || '')).toBe('');
    expect(String(ws.getCell(15, 5).value || '')).toBe('VE1256');
    expect(String(ws.getCell(17, 5).value || '')).toBe('Engineer');
    // Labels stay in column B
    expect(String(ws.getCell(14, 2).value || '')).toMatch(/name of the workman/i);
    expect(String(ws.getCell(15, 2).value || '')).toMatch(/serial number/i);
    // Column widths widened for readable alignment after download
    expect(Number(ws.getColumn(2).width)).toBeGreaterThanOrEqual(18);
    expect(Number(ws.getColumn(5).width)).toBeGreaterThanOrEqual(18);
  });
});
