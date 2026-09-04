import ExcelJS from 'exceljs';
import {
  buildFormTSEWorkbookWithTemplateStyles,
  prepareFormTSEExportRows,
  formTSEDownloadHasSubstantiveRows,
  expandFormTSEHeaderValueBoxes,
  ensureFormTSEKarnatakaTitleLayout,
  ensureFormTSEKarnatakaSystemGeneratedNoteCentered,
  finalizeFormTSEKarnatakaWorksheetExportBorders,
  renameFormTSEKarnatakaHeaderLabels,
  writeFormTSEHeaderFieldsToWorksheet,
  findFormTSEIdentityTableStartCol0,
  resolveFormTSEWriteExcelCols0,
  repairFormTSEWorkbookColumnAlignment,
  formTSEWorkbookHasIdentityInColumnA,
  looksLikeFormTSEEmployeeNameCell,
  applyFormTSEWorkbookOpenAtColumnA,
  prepareFormTSEWorkbookForDownload,
  computeFormTSEKarnatakaTotalDeductions,
  FORM_T_KA_DEFAULT_PAYMENT_MODE,
  FORM_T_KA_OT_HOURS_NIL,
  FORM_T_KA_RULE_CITATION,
  isFormTSEKarnatakaTotalOtHoursHeader,
  applyFormTSEKarnatakaOtHoursNilToMappedRows,
  resolveFormTSEKarnatakaDeductionTotalHeader,
  resolveFormTSEKarnatakaEarningsTotalHeader,
  resolveFormTSEKarnatakaNetAmountPayableHeader,
  isFormTSEKarnatakaEarningsTotalHeader,
  formatFormTSERegistrationIdForExcel,
  resolveFormTSEDownloadWriteHeaders,
  FORM_T_KA_ATTENDANCE_START_COL0,
  FORM_T_KARNATAKA_HEADER_BOX_END_COL,
} from './formTSEKarnataka';
import { excelJSCellHasFullBoxBorder, excelJSCellHasBorder } from '../../utils/excelTableBorders';

/** Minimal Form T Karnataka-like sheet: header R12, days R13, index R14, data from R15. */
async function buildMinimalFormTTemplateBuffer() {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Form T');

  ws.getCell(1, 1).value = 'FORM T';
  ws.getCell(2, 1).value = 'COMBINED MUSTER ROLL CUM REGISTER OF WAGES';
  ws.getCell(3, 1).value =
    '[See Rule 24(9-B) of Karnataka Shops & Commercial Establishment Rules, 1963]';
  ws.getCell(2, 12).value = 'COMBINED MUSTER ROLL CUM REGISTER OF WAGES';
  ws.getCell(3, 12).value =
    '[See Rule 24(9-B) of Karnataka Shops & Commercial Establishment Rules, 1963]';
  ws.mergeCells(2, 5, 2, 14);
  ws.mergeCells(3, 5, 3, 14);
  ws.getCell(9, 1).value = 'Month / Year';
  ws.getCell(10, 1).value = 'Address of the Establishment';
  ws.getCell(11, 1).value = 'Name and Address of the Employer';
  // Narrow template boxes (A:D) — export must widen these to A:I.
  ws.mergeCells(9, 1, 9, 4);
  ws.mergeCells(10, 1, 10, 4);
  ws.mergeCells(11, 1, 11, 4);

  const identity = [
    'S.NO',
    'Name of Employee',
    "Father / Husband's Name",
    'Gender',
    'Designation / Department',
    'Date of Joining',
    'ESI No.',
    'UAN No.',
    'Wages fixed including VDA',
  ];
  identity.forEach((h, i) => {
    ws.getCell(12, i + 1).value = h;
  });
  // Attendance banner + day numbers (cols 10..40)
  for (let d = 1; d <= 31; d += 1) {
    const c = 9 + d;
    ws.getCell(12, c).value =
      'ATTENDANCE (Please mention the date of suspension of employees, if any)';
    ws.getCell(13, c).value = d;
  }
  // Statutory column index strip
  for (let c = 1; c <= 40; c += 1) {
    ws.getCell(14, c).value = c;
  }
  // One empty template body row with borders area
  for (let c = 1; c <= 9; c += 1) {
    ws.getCell(15, c).value = '';
  }
  ws.getCell(20, 1).value = 'Date:';

  return wb.xlsx.writeBuffer();
}

describe('Form T Karnataka workbook write', () => {
  const modalHeaders = [
    'S.NO',
    'Name of Employee',
    "Father / Husband's Name",
    'Gender',
    'Designation / Department',
    'Date of Joining',
    'ESI No.',
    'UAN No.',
    'Wages fixed including VDA',
  ];

  const liveRows = [
    {
      'S.NO': 1,
      'Name of Employee': 'Issac Kanagaraj',
      "Father / Husband's Name": 'abc',
      Gender: 'Male',
      'Designation / Department': 'Assistant Manager',
      'Date of Joining': '01 Dec 2025',
      'ESI No.': '',
      'UAN No.': '',
      'Wages fixed including VDA': '',
    },
    {
      'S.NO': 2,
      'Name of Employee': 'Saravanan',
      "Father / Husband's Name": 'abc',
      Gender: 'Male',
      'Designation / Department': 'Engineer',
      'Date of Joining': '01 Dec 2025',
      'ESI No.': '',
      'UAN No.': '101217123456',
      'Wages fixed including VDA': '',
    },
  ];

  it('writes autofill employee names into Excel data rows (not blank)', async () => {
    const templateArrayBuffer = await buildMinimalFormTTemplateBuffer();
    expect(formTSEDownloadHasSubstantiveRows(liveRows, modalHeaders)).toBe(true);

    const { blob } = await buildFormTSEWorkbookWithTemplateStyles({
      templateArrayBuffer,
      mappedData: liveRows,
      headersToUse: modalHeaders,
      parsedHeaderRowIndex: 11, // 0-based → Excel row 12
      parsedDataStartIndex: 14, // 0-based → Excel row 15
      parsedTableStartCol: 0,
      parsedFormHeader: { title: 'Form T' },
      headerFormData: {
        form_t_month_year: 'May 2026',
        form_t_establishment_name_address:
          '33/220KV Substation, Survey No: 132, Karjol Village, Bijapur Dist – 586113 Karnataka',
        form_t_employer:
          'M/s Clean Wind Power Bableshwar Pvt Ltd, 33/220KV Substation, Survey No: 132, Karjol Village',
      },
      formFileName: 'Form_T_Karnataka.xlsx',
      sheetNameHint: 'Form T',
    });

    const outWb = new ExcelJS.Workbook();
    await outWb.xlsx.load(await new Response(blob).arrayBuffer());
    const ws = outWb.getWorksheet('Form T');
    expect(ws).toBeTruthy();

    expect(String(ws.getCell(15, 1).value ?? '').trim()).toBe('1');
    expect(String(ws.getCell(15, 2).value ?? '')).toMatch(/Issac/i);
    expect(String(ws.getCell(15, 4).value ?? '')).toMatch(/Male/i);
    expect(String(ws.getCell(16, 2).value ?? '')).toMatch(/Saravanan/i);
    expect(String(ws.getCell(16, 8).value ?? '').trim()).toBe('101217123456');

    // Header values live in widened A–I boxes (not clipped A:D).
    const merges = Array.isArray(ws.model?.merges) ? ws.model.merges : [];
    expect(merges.some((m) => /^A9:I9$/i.test(String(m)))).toBe(true);
    expect(merges.some((m) => /^A10:I10$/i.test(String(m)))).toBe(true);
    expect(merges.some((m) => /^A11:I11$/i.test(String(m)))).toBe(true);
    expect(String(ws.getCell(9, 1).value ?? '')).toMatch(/May 2026/i);
    expect(String(ws.getCell(10, 1).value ?? '')).toMatch(/Karjol/i);
    expect(String(ws.getCell(11, 1).value ?? '')).toMatch(/Clean Wind/i);
    expect(ws.getCell(10, 1).alignment?.wrapText).toBe(true);

    // Table grid borders start at row 12 only — not on Month/Year header boxes (9–11).
    [9, 10, 11].forEach((row) => {
      expect(excelJSCellHasFullBoxBorder(ws.getCell(row, 1))).toBe(false);
    });
    // Row 12+ table grid uses full thin borders across identity and attendance columns.
    expect(excelJSCellHasFullBoxBorder(ws.getCell(12, 1))).toBe(true);
    expect(excelJSCellHasFullBoxBorder(ws.getCell(12, 9))).toBe(true);
    expect(excelJSCellHasFullBoxBorder(ws.getCell(13, 2))).toBe(true);
    expect(excelJSCellHasFullBoxBorder(ws.getCell(13, 10))).toBe(true);
    expect(excelJSCellHasFullBoxBorder(ws.getCell(14, 5))).toBe(true);
    expect(excelJSCellHasFullBoxBorder(ws.getCell(15, 1))).toBe(true);
    expect(excelJSCellHasFullBoxBorder(ws.getCell(15, 8))).toBe(true);
    expect(excelJSCellHasFullBoxBorder(ws.getCell(16, 2))).toBe(true);
  });

  it('left-aligns title rows 2–3 starting in column A with full text', async () => {
    const templateArrayBuffer = await buildMinimalFormTTemplateBuffer();
    const { blob } = await buildFormTSEWorkbookWithTemplateStyles({
      templateArrayBuffer,
      mappedData: liveRows.slice(0, 1),
      headersToUse: modalHeaders,
      parsedHeaderRowIndex: 11,
      parsedDataStartIndex: 14,
      parsedFormHeader: { title: 'Form T' },
      sheetNameHint: 'Form T',
    });

    const outWb = new ExcelJS.Workbook();
    await outWb.xlsx.load(await new Response(blob).arrayBuffer());
    const ws = outWb.getWorksheet('Form T');
    expect(String(ws.getCell(2, 1).value ?? '')).toMatch(/^COMBINED MUSTER ROLL/i);
    expect(String(ws.getCell(3, 1).value ?? '')).toMatch(/^\[See Rule 24\(9-B\)/i);
    expect(ws.getCell(2, 1).alignment?.horizontal).toBe('left');
    expect(ws.getCell(3, 1).alignment?.horizontal).toBe('left');
    expect(Number(ws.getCell(2, 1).alignment?.indent || 0)).toBe(0);
    const merges = Array.isArray(ws.model?.merges) ? ws.model.merges : [];
    expect(merges.some((m) => /^A2:K2$/i.test(String(m)))).toBe(true);
    expect(merges.some((m) => /^A3:K3$/i.test(String(m)))).toBe(true);
  });

  it('preserves title text when template rows use wide horizontal merges', () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Form T');
    ws.getCell(1, 1).value = 'FORM T';
    ws.getCell(2, 1).value = 'COMBINED MUSTER ROLL CUM REGISTER OF WAGES';
    ws.getCell(3, 1).value =
      '[See Rule 24(9-B) of Karnataka Shops & Commercial Establishment Rules, 1963]';
    ws.getCell(9, 1).value = 'Month / Year : July 2026';
    ws.getCell(12, 2).value = 'Name of Employee';
    ws.mergeCells(2, 5, 2, 14);
    ws.mergeCells(3, 5, 3, 14);

    ensureFormTSEKarnatakaTitleLayout(ws);

    expect(String(ws.getCell(1, 1).value ?? '')).toMatch(/^FORM T/i);
    expect(String(ws.getCell(2, 1).value ?? '')).toMatch(/^COMBINED MUSTER ROLL/i);
    expect(String(ws.getCell(3, 1).value ?? '')).toMatch(/^\[See Rule 24\(9-B\)/i);
    expect(ws.getCell(2, 1).alignment?.horizontal).toBe('left');
  });

  it('centers System Generated Document footer across the table width', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Form T');
    ws.getCell(1, 1).value = 'FORM T';
    ws.getCell(2, 1).value = 'COMBINED MUSTER ROLL CUM REGISTER OF WAGES';
    ws.getCell(3, 1).value =
      '[See Rule 24(9-B) of Karnataka Shops & Commercial Establishment Rules, 1963]';
    ws.getCell(9, 1).value = 'Month / Year : May 2026';
    ws.getCell(12, 1).value = 'S.NO';
    ws.getCell(12, 2).value = 'Name of Employee';
    for (let d = 1; d <= 8; d += 1) {
      ws.getCell(13, 9 + d).value = d;
    }
    ws.getCell(24, 1).value = 'stem Generated Document';

    finalizeFormTSEKarnatakaWorksheetExportBorders(ws, { colTo: 17, tableLastRow: 17 });

    expect(String(ws.getCell(24, 1).value ?? '')).toBe('This is a System Generated Document');
    expect(ws.getCell(24, 1).alignment?.horizontal).toBe('center');
    const merges = Array.isArray(ws.model?.merges) ? ws.model.merges : [];
    expect(merges.some((m) => /^A24:Q24$/i.test(String(m)))).toBe(true);
  });

  it('fills employee names from People when mappedData is an empty template grid', async () => {
    const templateArrayBuffer = await buildMinimalFormTTemplateBuffer();
    const { blob } = await buildFormTSEWorkbookWithTemplateStyles({
      templateArrayBuffer,
      mappedData: [{ 'S.NO': '', 'Name of Employee': '', Gender: '' }],
      headersToUse: modalHeaders,
      employees: [
        { FirstName: 'Issac', LastName: 'Kanagaraj', Sex: 'Male', Designation: 'Assistant Manager' },
        { FirstName: 'Saravanan', Sex: 'Male', Designation: 'Engineer' },
      ],
      parsedHeaderRowIndex: 11,
      parsedDataStartIndex: 19, // stale hint at/after Date: footer
      parsedTableStartCol: 0,
      parsedFormHeader: { title: 'Form T' },
      headerFormData: {
        form_t_month_year: 'March 2026',
        form_t_establishment_name_address:
          '33/220KV Substation, Survey No: 132, Karjol Village, Bijapur Dist – 586113 Karnataka',
        form_t_employer: 'M/s Clean Wind Power Bableshwar Pvt Ltd',
      },
      formFileName: 'Form_T_KA.xlsx',
      sheetNameHint: 'Form T',
    });

    const outWb = new ExcelJS.Workbook();
    await outWb.xlsx.load(await new Response(blob).arrayBuffer());
    const ws = outWb.getWorksheet('Form T');
    expect(String(ws.getCell(15, 1).value ?? '').trim()).toBe('1');
    expect(String(ws.getCell(15, 2).value ?? '')).toMatch(/Issac/i);
    expect(String(ws.getCell(16, 2).value ?? '')).toMatch(/Saravanan/i);
    expect(String(ws.getCell(15, 2).value ?? '')).not.toMatch(/Date/i);
  });

  it('expands narrow A:D header boxes to A:I for long address/employer text', () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Form T');
    ws.getCell(9, 1).value = 'Month / Year';
    ws.getCell(10, 1).value = 'Address of the Establishment';
    ws.getCell(11, 1).value = 'Name and Address of the Employer';
    ws.mergeCells(9, 1, 9, 4);
    ws.mergeCells(10, 1, 10, 4);
    ws.mergeCells(11, 1, 11, 4);
    const boxBorder = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'thin' },
      right: { style: 'thin' },
    };
    [9, 10, 11].forEach((row) => {
      for (let col = 1; col <= 4; col += 1) {
        ws.getCell(row, col).border = { ...boxBorder };
      }
    });

    writeFormTSEHeaderFieldsToWorksheet(ws, {
      form_t_month_year: 'May 2026',
      form_t_establishment_name_address:
        '33/220KV Substation, Survey No: 132, Karjol Village, Bijapur Dist – 586113 Karnataka, Karnataka',
      form_t_employer:
        'M/s Clean Wind Power Bableshwar Pvt Ltd, 33/220KV Substation, Survey No: 132, Karjol Village',
    });
    expandFormTSEHeaderValueBoxes(ws, FORM_T_KARNATAKA_HEADER_BOX_END_COL);

    const merges = Array.isArray(ws.model?.merges) ? ws.model.merges : [];
    expect(merges).toEqual(expect.arrayContaining(['A9:I9', 'A10:I10', 'A11:I11']));
    expect(String(ws.getCell(10, 1).value ?? '')).toMatch(
      /Name and address of the Establishment\s*:/
    );
    expect(String(ws.getCell(10, 1).value ?? '')).toMatch(/Karjol Village/);
    expect(String(ws.getCell(11, 1).value ?? '')).toMatch(/Name and Address of employer\s*:/);
    expect(String(ws.getCell(11, 1).value ?? '')).toMatch(/Clean Wind Power/);
    [9, 10, 11].forEach((row) => {
      expect(excelJSCellHasFullBoxBorder(ws.getCell(row, 1))).toBe(false);
      expect(excelJSCellHasFullBoxBorder(ws.getCell(row, FORM_T_KARNATAKA_HEADER_BOX_END_COL))).toBe(
        false
      );
    });
  });

  it('clears rows 3–8 / L–AN boxes and renames contractor + nature-of-work labels', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Form T');
    const boxBorder = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'thin' },
      right: { style: 'thin' },
    };
    ws.getCell(1, 1).value = 'FORM T';
    ws.getCell(2, 1).value = 'COMBINED MUSTER ROLL CUM REGISTER OF WAGES';
    ws.getCell(3, 1).value =
      'Name and address of establishment in / under which contract is carried on';
    ws.getCell(4, 1).value = 'in lieu of';
    ws.getCell(5, 1).value =
      '1. Form I, II of Rule 22(4); Form IV of Rule 29(2); Forms V & VII of Rule 29(1) & (5) of Karnataka Minimum Wages Rules, 1958';
    ws.getCell(6, 1).value = '2. Form I of Rules 3(1) of Karnataka Payment of Wages Rules, 1963';
    ws.getCell(7, 1).value =
      '3. Form XIII of Rules 75; Form XV, XVII, XX, XXI, XXII, XXIII of 78(1)(a)(i), (ii) & (iii) of Karnataka Contract Labour (Regulation & Abolition) Rules, 1974';
    ws.getCell(8, 1).value =
      '4. Form XIII of Rule 43; Forms XVII, XVIII, XIX, XX, XXI, XXII of Rule 46(2)(a),(c) & (d) of Inter-state Migrant Workmen (Regulation of Employment and conditions of service) Karnataka Rules, 1981';
    ws.getCell(9, 1).value = 'Month / Year';
    ws.getCell(10, 1).value = 'Name and address of contractor';
    ws.getCell(11, 1).value = 'Nature of work and location of work';
    ws.getCell(12, 2).value = 'Name of Employee';
    // Obsolete bordered boxes on rows 3–8 and L–AN (cols 12–40).
    for (let r = 3; r <= 8; r += 1) {
      for (let c = 1; c <= 11; c += 1) {
        ws.getCell(r, c).border = { ...boxBorder };
      }
    }
    for (let r = 1; r <= 8; r += 1) {
      for (let c = 12; c <= 40; c += 1) {
        ws.getCell(r, c).border = { ...boxBorder };
      }
    }

    finalizeFormTSEKarnatakaWorksheetExportBorders(ws, { colTo: 40, tableLastRow: 16 });
    renameFormTSEKarnatakaHeaderLabels(ws);
    writeFormTSEHeaderFieldsToWorksheet(ws, {
      form_t_month_year: 'June 2026',
      form_t_establishment_name_address: 'Karjol Village Establishment, Bijapur Dist',
      form_t_employer: 'M/s Clean Wind Power Bableshwar Pvt Ltd',
    });

    expect(String(ws.getCell(3, 1).value ?? '')).toBe(FORM_T_KA_RULE_CITATION);
    expect(String(ws.getCell(4, 1).value ?? '')).toMatch(/^in lieu of$/i);
    expect(String(ws.getCell(5, 1).value ?? '')).toMatch(/Minimum Wages Rules, 1958/i);
    expect(String(ws.getCell(6, 1).value ?? '')).toMatch(/Payment of Wages Rules, 1963/i);
    expect(String(ws.getCell(7, 1).value ?? '')).toMatch(/Contract Labour/i);
    expect(String(ws.getCell(8, 1).value ?? '')).toMatch(/Migrant Workmen/i);
    // CLRA rows 4–8 stay visible (not collapsed); boxes stripped.
    expect(ws.getRow(4).hidden).toBe(false);
    expect(ws.getRow(8).hidden).toBe(false);
    expect(ws.getRow(3).hidden).toBe(false);
    expect(ws.getRow(9).hidden).toBe(false);
    expect(excelJSCellHasBorder(ws.getCell(5, 3))).toBe(false);
    expect(excelJSCellHasBorder(ws.getCell(3, 20))).toBe(false);
    expect(excelJSCellHasBorder(ws.getCell(1, 12))).toBe(false);
    expect(excelJSCellHasBorder(ws.getCell(8, 40))).toBe(false);
    expect(String(ws.getCell(10, 1).value ?? '')).toMatch(
      /Name and address of the Establishment\s*:/
    );
    expect(String(ws.getCell(10, 1).value ?? '')).toMatch(/Karjol Village/);
    expect(String(ws.getCell(11, 1).value ?? '')).toMatch(/Name and Address of employer\s*:/);
    expect(String(ws.getCell(11, 1).value ?? '')).toMatch(/Clean Wind Power/);
    expect(String(ws.getCell(10, 1).value ?? '')).not.toMatch(/contractor/i);
    expect(String(ws.getCell(11, 1).value ?? '')).not.toMatch(/nature of work/i);
    // Establishment/employer must not be left only on wiped row 3 (Rule citation).
    expect(String(ws.getCell(3, 1).value ?? '')).not.toMatch(/Karjol Village|Clean Wind/i);
  });

  it('force-stamps Establishment and Employer Label : value on rows 10–11', () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Form T');
    ws.getCell(1, 1).value = 'FORM T';
    ws.getCell(2, 1).value = 'COMBINED MUSTER ROLL CUM REGISTER OF WAGES';
    ws.getCell(3, 1).value = FORM_T_KA_RULE_CITATION;
    ws.getCell(9, 1).value = 'Month / Year';
    ws.getCell(10, 1).value = 'Name and address of the Establishment';
    ws.getCell(11, 1).value = 'Name and Address of employer';
    ws.getCell(12, 2).value = 'Name of Employee';
    // No A:I merge — previously left values only in column B (looked blank in A).
    writeFormTSEHeaderFieldsToWorksheet(ws, {
      form_t_month_year: 'May 2026',
      statutory_establishment_name_address: 'Karjol Village Establishment, Bijapur Dist',
      statutory_employer_name_address: 'M/s Clean Wind Power Bableshwar Pvt Ltd',
    });
    expect(String(ws.getCell(10, 1).value ?? '')).toMatch(
      /Name and address of the Establishment\s*:\s*Karjol Village/
    );
    expect(String(ws.getCell(11, 1).value ?? '')).toMatch(
      /Name and Address of employer\s*:\s*M\/s Clean Wind/
    );
  });

  it('prepareFormTSEWorkbookForDownload re-stamps Establishment/Employer on an existing buffer', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Form T');
    ws.getCell(1, 1).value = 'FORM T';
    ws.getCell(2, 1).value = 'COMBINED MUSTER ROLL CUM REGISTER OF WAGES';
    ws.getCell(3, 1).value = FORM_T_KA_RULE_CITATION;
    ws.getCell(9, 1).value = 'Month / Year : June 2026';
    ws.getCell(10, 1).value = 'Name and address of the Establishment';
    ws.getCell(11, 1).value = 'Name and Address of employer';
    ws.getCell(12, 1).value = 'S.NO';
    ws.getCell(12, 2).value = 'Name of Employee';
    ws.getCell(15, 1).value = '1';
    ws.getCell(15, 2).value = 'Vinay Kumar';
    const buf = await wb.xlsx.writeBuffer();

    const stamped = await prepareFormTSEWorkbookForDownload(buf, {
      force: true,
      headerFormData: {
        form_t_month_year: 'June 2026',
        form_t_establishment_name_address: 'Karjol Village Establishment, Bijapur Dist',
        form_t_employer: 'M/s Clean Wind Power Bableshwar Pvt Ltd',
      },
    });

    const outWb = new ExcelJS.Workbook();
    await outWb.xlsx.load(stamped);
    const out = outWb.getWorksheet('Form T');
    expect(String(out.getCell(10, 1).value ?? '')).toMatch(/Establishment\s*:\s*Karjol Village/);
    expect(String(out.getCell(11, 1).value ?? '')).toMatch(/employer\s*:\s*M\/s Clean Wind/i);
  });

  it('writes Establishment/Employer to rows 10–11 even when row 3 has contract-carried label', () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Form T');
    ws.getCell(1, 1).value = 'FORM T';
    ws.getCell(2, 1).value = 'COMBINED MUSTER ROLL CUM REGISTER OF WAGES';
    ws.getCell(3, 1).value =
      'Name and address of establishment in / under which contract is carried on';
    ws.getCell(9, 1).value = 'Month / Year';
    ws.getCell(10, 1).value = 'Name and address of contractor';
    ws.getCell(11, 1).value = 'Nature of work and location of work';
    ws.getCell(12, 2).value = 'Name of Employee';

    writeFormTSEHeaderFieldsToWorksheet(ws, {
      form_t_month_year: 'May 2026',
      form_t_establishment_name_address: 'Karjol Village Establishment, Bijapur Dist',
      form_t_employer: 'M/s Clean Wind Power Bableshwar Pvt Ltd',
    });
    finalizeFormTSEKarnatakaWorksheetExportBorders(ws, { colTo: 20, tableLastRow: 16 });
    writeFormTSEHeaderFieldsToWorksheet(ws, {
      form_t_month_year: 'May 2026',
      form_t_establishment_name_address: 'Karjol Village Establishment, Bijapur Dist',
      form_t_employer: 'M/s Clean Wind Power Bableshwar Pvt Ltd',
    });

    expect(String(ws.getCell(3, 1).value ?? '')).toBe(FORM_T_KA_RULE_CITATION);
    expect(String(ws.getCell(3, 1).value ?? '')).not.toMatch(/Karjol|Clean Wind/i);
    expect(String(ws.getCell(10, 1).value ?? '')).toMatch(/Establishment\s*:/);
    expect(String(ws.getCell(10, 1).value ?? '')).toMatch(/Karjol Village/);
    expect(String(ws.getCell(11, 1).value ?? '')).toMatch(/employer\s*:/i);
    expect(String(ws.getCell(11, 1).value ?? '')).toMatch(/Clean Wind Power/);
  });

  it('still writes identity columns when modal hints tableStartCol under ATTENDANCE', async () => {
    const templateArrayBuffer = await buildMinimalFormTTemplateBuffer();
    const { blob } = await buildFormTSEWorkbookWithTemplateStyles({
      templateArrayBuffer,
      mappedData: liveRows,
      headersToUse: modalHeaders,
      parsedHeaderRowIndex: 11,
      parsedDataStartIndex: 13, // wrongly points at index strip
      parsedTableStartCol: 9, // wrongly under attendance
      parsedFormHeader: { title: 'Form T' },
      headerFormData: {},
      formFileName: 'Form_T_Karnataka.xlsx',
      sheetNameHint: 'Form T',
    });

    const outWb = new ExcelJS.Workbook();
    await outWb.xlsx.load(await new Response(blob).arrayBuffer());
    const ws = outWb.getWorksheet('Form T');
    // Must land in A–I identity band — not under ATTENDANCE (col J+).
    expect(String(ws.getCell(15, 1).value ?? '').trim()).toBe('1');
    expect(String(ws.getCell(15, 2).value ?? '')).toMatch(/Issac/i);
    expect(String(ws.getCell(15, 4).value ?? '')).toMatch(/Male/i);
    expect(String(ws.getCell(15, 5).value ?? '')).toMatch(/Assistant/i);
    expect(String(ws.getCell(16, 2).value ?? '')).toMatch(/Saravanan/i);
    // Column J must not hold the employee name / serial dump.
    expect(String(ws.getCell(15, 10).value ?? '')).not.toMatch(/Issac/i);
    expect(String(ws.getCell(15, 11).value ?? '')).not.toMatch(/Issac/i);
  });

  it('writes identity to A–I even when headersToUse are attendance-first', async () => {
    const templateArrayBuffer = await buildMinimalFormTTemplateBuffer();
    // Broken parse order: attendance keys first, then identity — values still carry names.
    const brokenHeaders = [
      'ATTENDANCE_1',
      'ATTENDANCE_2',
      'S.NO',
      'Name of Employee',
      "Father / Husband's Name",
      'Gender',
      'Designation / Department',
    ];
    const brokenRows = liveRows.map((row, i) => ({
      ATTENDANCE_1: 'P',
      ATTENDANCE_2: 'WO',
      'S.NO': i + 1,
      'Name of Employee': row['Name of Employee'],
      "Father / Husband's Name": row["Father / Husband's Name"],
      Gender: row.Gender,
      'Designation / Department': row['Designation / Department'],
    }));
    const { blob } = await buildFormTSEWorkbookWithTemplateStyles({
      templateArrayBuffer,
      mappedData: brokenRows,
      headersToUse: brokenHeaders,
      parsedHeaderRowIndex: 11,
      parsedDataStartIndex: 14,
      parsedTableStartCol: 9,
      parsedFormHeader: { title: 'Form T' },
      headerFormData: {},
      formFileName: 'Form_T_Karnataka.xlsx',
      sheetNameHint: 'Form T',
    });
    const outWb = new ExcelJS.Workbook();
    await outWb.xlsx.load(await new Response(blob).arrayBuffer());
    const ws = outWb.getWorksheet('Form T');
    expect(String(ws.getCell(15, 1).value ?? '').trim()).toBe('1');
    expect(String(ws.getCell(15, 2).value ?? '')).toMatch(/Issac/i);
    expect(String(ws.getCell(15, 10).value ?? '')).not.toMatch(/Issac/i);
  });

  it('findFormTSEIdentityTableStartCol0 ignores ATTENDANCE-band preferCol', () => {
    const grid = Array.from({ length: 20 }, () => Array(20).fill(''));
    grid[11][0] = 'S.NO';
    grid[11][1] = 'Name of Employee';
    grid[11][9] = 'ATTENDANCE';
    const start = findFormTSEIdentityTableStartCol0({
      getMergedAwareCellText: (r, c) => grid[r]?.[c] || '',
      headerRowIndex: 11,
      preferCol: 9,
    });
    expect(start).toBe(0);
  });

  it('resolveFormTSEWriteExcelCols0 maps identity to A–I and days to J+', () => {
    const headers = [
      'S.NO',
      'Name of Employee',
      "Father / Husband's Name",
      'Gender',
      'Designation / Department',
      'Date of Joining',
      'ESI No.',
      'UAN No.',
      'Wages fixed including VDA',
      'ATTENDANCE_1',
      'ATTENDANCE_2',
      'ATTENDANCE_10',
    ];
    const cols = resolveFormTSEWriteExcelCols0(headers);
    expect(cols[0]).toBe(0); // A
    expect(cols[1]).toBe(1); // B Name
    expect(cols[5]).toBe(5); // F Date of Joining
    expect(cols[8]).toBe(8); // I Wages
    expect(cols[9]).toBe(FORM_T_KA_ATTENDANCE_START_COL0); // J day 1
    expect(cols[10]).toBe(FORM_T_KA_ATTENDANCE_START_COL0 + 1); // K day 2
    expect(cols[11]).toBe(FORM_T_KA_ATTENDANCE_START_COL0 + 9); // day 10
  });

  it('maps row values when live keys lack statutory (N) suffixes but export headers have them', () => {
    const exportHeaders = modalHeaders.map((h, i) => `${h} (${i + 1})`);
    const prepared = prepareFormTSEExportRows(liveRows, exportHeaders, modalHeaders);
    expect(formTSEDownloadHasSubstantiveRows(prepared, exportHeaders)).toBe(true);
    expect(String(prepared[0][exportHeaders[1]])).toMatch(/Issac/i);
  });

  it('maps row values when live keys use statutory (N) suffixes but export headers do not', () => {
    const liveWithSuffix = [
      {
        'S.NO (1)': 1,
        'Name of Employee (2)': 'Issac Kanagaraj',
        "Father / Husband's Name (3)": 'abc',
        'Gender (4)': 'Male',
      },
    ];
    const plainHeaders = ['S.NO', 'Name of Employee', "Father / Husband's Name", 'Gender'];
    const prepared = prepareFormTSEExportRows(liveWithSuffix, plainHeaders, plainHeaders);
    expect(String(prepared[0]['Name of Employee'])).toMatch(/Issac/i);
    expect(String(prepared[0].Gender)).toMatch(/Male/i);
  });

  it('repairFormTSEWorkbookColumnAlignment shifts identity from J–K back to A–B', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Form T');
    // Header labels in A–I (correct) — data wrongly dumped under attendance (J+).
    [
      'S.NO',
      'Name of Employee',
      "Father / Husband's Name",
      'Gender',
      'Designation / Department',
      'Date of Joining',
      'ESI No.',
      'UAN No.',
      'Wages fixed including VDA',
    ].forEach((h, i) => {
      ws.getCell(12, i + 1).value = h;
    });
    // Bleed header text into B15 (old bug made repair think row was aligned).
    ws.getCell(15, 2).value = 'Name of Employee';
    ws.getCell(15, 10).value = 1;
    ws.getCell(15, 11).value = 'Issac Kanagaraj';
    ws.getCell(15, 12).value = 'abc';
    ws.getCell(15, 13).value = 'Male';
    ws.getCell(15, 14).value = 'Assistant Manager';
    ws.getCell(15, 19).value = 'P';
    ws.getCell(16, 10).value = 2;
    ws.getCell(16, 11).value = 'Saravanan';
    ws.getCell(16, 13).value = 'Male';

    expect(looksLikeFormTSEEmployeeNameCell('Name of Employee')).toBe(false);
    expect(looksLikeFormTSEEmployeeNameCell('Issac Kanagaraj')).toBe(true);

    const shiftedBuf = await wb.xlsx.writeBuffer();
    expect(await formTSEWorkbookHasIdentityInColumnA(shiftedBuf)).toBe(false);

    const repairedBuf = await repairFormTSEWorkbookColumnAlignment(shiftedBuf);
    expect(await formTSEWorkbookHasIdentityInColumnA(repairedBuf)).toBe(true);

    const outWb = new ExcelJS.Workbook();
    await outWb.xlsx.load(repairedBuf);
    const out = outWb.getWorksheet('Form T');
    expect(String(out.getCell(15, 1).value ?? '').trim()).toBe('1');
    expect(String(out.getCell(15, 2).value ?? '')).toMatch(/Issac/i);
    expect(String(out.getCell(15, 4).value ?? '')).toMatch(/Male/i);
    expect(String(out.getCell(15, 5).value ?? '')).toMatch(/Assistant/i);
    expect(String(out.getCell(15, 11).value ?? '')).not.toMatch(/Issac/i);
    expect(String(out.getCell(16, 2).value ?? '')).toMatch(/Saravanan/i);
  });

  it('opens the sheet at column A instead of frozen attendance column J', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Form T');
    ws.getCell(12, 1).value = 'S.NO';
    ws.getCell(12, 2).value = 'Name of Employee';
    ws.getCell(15, 1).value = 1;
    ws.getCell(15, 2).value = 'Kanagaraj';
    ws.views = [
      {
        state: 'frozen',
        xSplit: 9,
        ySplit: 14,
        topLeftCell: 'J15',
        activeCell: 'J15',
      },
    ];
    const frozenBuf = await wb.xlsx.writeBuffer();
    const opened = await applyFormTSEWorkbookOpenAtColumnA(frozenBuf);
    const outWb = new ExcelJS.Workbook();
    await outWb.xlsx.load(opened);
    const out = outWb.getWorksheet('Form T');
    const view = (out.views && out.views[0]) || {};
    expect(String(view.topLeftCell || 'A1').toUpperCase()).toBe('A1');
    expect(String(view.activeCell || 'A1').toUpperCase()).toBe('A1');
    expect(Number(view.xSplit || 0)).toBe(0);
    expect(view.state === 'frozen' ? Number(view.xSplit || 0) : 0).toBe(0);
  });

  it('prepareFormTSEWorkbookForDownload shifts J identity and opens at A', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Form T');
    ws.getCell(12, 1).value = 'S.NO';
    ws.getCell(12, 2).value = 'Name of Employee';
    ws.getCell(15, 10).value = 1;
    ws.getCell(15, 11).value = 'Kanagaraj';
    ws.views = [{ state: 'frozen', xSplit: 9, ySplit: 14, topLeftCell: 'J15', activeCell: 'J15' }];
    const buf = await wb.xlsx.writeBuffer();
    const ready = await prepareFormTSEWorkbookForDownload(buf);
    expect(await formTSEWorkbookHasIdentityInColumnA(ready)).toBe(true);
    const outWb = new ExcelJS.Workbook();
    await outWb.xlsx.load(ready);
    const out = outWb.getWorksheet('Form T');
    expect(String(out.getCell(15, 2).value ?? '')).toMatch(/Kanagaraj/i);
    const view = (out.views && out.views[0]) || {};
    expect(String(view.topLeftCell || 'A1').toUpperCase()).toBe('A1');
  });

  it('moves screenshot-style J-grid (serial in J, name in K, Male in M) onto A–D', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Form T');
    ws.getCell(12, 1).value = 'S.NO';
    ws.getCell(12, 2).value = 'Name of Employee';
    ws.getCell(12, 9).value = 'Wages fixed including VDA';
    ws.getCell(12, 20).value = 'ATTENDANCE';
    for (let d = 1; d <= 21; d += 1) {
      ws.getCell(13, 9 + d).value = d;
    }
    const people = [
      { name: 'nagaraj', desig: 'Manager' },
      { name: 'avanan', desig: 'engineer' },
      { name: 'erkhan', desig: 'engineer' },
    ];
    people.forEach((p, i) => {
      const r = 15 + i;
      ws.getCell(r, 10).value = i + 1;
      ws.getCell(r, 11).value = p.name;
      ws.getCell(r, 12).value = 'abc';
      ws.getCell(r, 13).value = 'Male';
      ws.getCell(r, 14).value = p.desig;
      ws.getCell(r, 20).value = 'P';
    });
    const buf = await wb.xlsx.writeBuffer();
    const ready = await prepareFormTSEWorkbookForDownload(buf);
    const outWb = new ExcelJS.Workbook();
    await outWb.xlsx.load(ready);
    const out = outWb.getWorksheet('Form T');
    expect(String(out.getCell(15, 1).value ?? '').trim()).toBe('1');
    expect(String(out.getCell(15, 2).value ?? '')).toMatch(/nagaraj/i);
    expect(String(out.getCell(15, 4).value ?? '')).toMatch(/Male/i);
    expect(String(out.getCell(16, 2).value ?? '')).toMatch(/avanan/i);
    expect(String(out.getCell(15, 11).value ?? '')).not.toMatch(/nagaraj/i);
  });

  it('treats Save auto-download names in row 6 as aligned identity (not only row 15)', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Form T');
    ws.getCell(1, 1).value = 'Month / Year : April 2026';
    ws.getCell(4, 1).value = 'S.NO';
    ws.getCell(4, 2).value = 'Name of Employee';
    ws.getCell(6, 1).value = 1;
    ws.getCell(6, 2).value = 'Kanagaraj';
    ws.getCell(7, 1).value = 2;
    ws.getCell(7, 2).value = 'Saravanan';
    const buf = await wb.xlsx.writeBuffer();
    expect(await formTSEWorkbookHasIdentityInColumnA(buf)).toBe(true);
  });

  it('does not treat a header-only Form T_KA template as having employee rows', async () => {
    expect(looksLikeFormTSEEmployeeNameCell('Address of the Establishment : Karjol Village')).toBe(
      false
    );
    expect(looksLikeFormTSEEmployeeNameCell('Name and Address of the Employer')).toBe(false);
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Form T');
    ws.getCell(6, 1).value = 'Karnataka Payment of Wages Rules, 1963';
    ws.getCell(9, 1).value = 'Month / Year : March 2026';
    ws.getCell(10, 1).value = 'Address of the Establishment : 33/220KV Substation Karjol Village';
    ws.mergeCells(10, 1, 10, 9);
    ws.getCell(11, 1).value = 'Name and Address of the Employer : M/s Clean Wind Power';
    ws.mergeCells(11, 1, 11, 9);
    [
      'S.NO',
      'Name of Employee',
      "Father / Husband's Name",
      'Gender',
      'Designation / Department',
      'Date of Joining',
      'ESI No.',
      'UAN No.',
    ].forEach((h, i) => {
      ws.getCell(12, i + 1).value = h;
    });
    for (let c = 1; c <= 8; c += 1) ws.getCell(13, c).value = c;
    ws.getCell(20, 1).value = 'Date:';
    const buf = await wb.xlsx.writeBuffer();
    expect(await formTSEWorkbookHasIdentityInColumnA(buf)).toBe(false);
  });

  it('clears prior J-shifted Draft leftovers when rewriting onto a dirty template', async () => {
    // Simulate saved Draft that already has identity under ATTENDANCE (user screenshot).
    const dirtyWb = new ExcelJS.Workbook();
    const dirtyWs = dirtyWb.addWorksheet('Form T');
    dirtyWs.getCell(1, 1).value = 'FORM T';
    dirtyWs.getCell(9, 1).value = 'Month / Year';
    dirtyWs.getCell(10, 1).value = 'Address of the Establishment';
    dirtyWs.getCell(11, 1).value = 'Name and Address of the Employer';
    [
      'S.NO',
      'Name of Employee',
      "Father / Husband's Name",
      'Gender',
      'Designation / Department',
      'Date of Joining',
      'ESI No.',
      'UAN No.',
      'Wages fixed including VDA',
    ].forEach((h, i) => {
      dirtyWs.getCell(12, i + 1).value = h;
    });
    for (let d = 1; d <= 31; d += 1) {
      dirtyWs.getCell(12, 9 + d).value = 'ATTENDANCE';
      dirtyWs.getCell(13, 9 + d).value = d;
    }
    for (let c = 1; c <= 40; c += 1) dirtyWs.getCell(14, c).value = c;
    // Wrong prior write: S.NO/Name under J/K
    dirtyWs.getCell(15, 10).value = 1;
    dirtyWs.getCell(15, 11).value = 'OldShifted Name';
    dirtyWs.getCell(15, 12).value = 'abc';
    dirtyWs.getCell(15, 13).value = 'Male';
    dirtyWs.getCell(15, 19).value = 'P';

    const dirtyBuf = await dirtyWb.xlsx.writeBuffer();
    const { blob } = await buildFormTSEWorkbookWithTemplateStyles({
      templateArrayBuffer: dirtyBuf,
      mappedData: liveRows,
      headersToUse: modalHeaders,
      parsedHeaderRowIndex: 11,
      parsedDataStartIndex: 14,
      parsedTableStartCol: 9,
      parsedFormHeader: { title: 'Form T' },
      headerFormData: {},
      formFileName: 'Form_T_Karnataka.xlsx',
      sheetNameHint: 'Form T',
    });

    const outWb = new ExcelJS.Workbook();
    await outWb.xlsx.load(await new Response(blob).arrayBuffer());
    const ws = outWb.getWorksheet('Form T');
    expect(String(ws.getCell(15, 1).value ?? '').trim()).toBe('1');
    expect(String(ws.getCell(15, 2).value ?? '')).toMatch(/Issac/i);
    expect(String(ws.getCell(15, 10).value ?? '')).not.toMatch(/OldShifted/i);
    expect(String(ws.getCell(15, 11).value ?? '')).not.toMatch(/OldShifted/i);
    expect(String(ws.getCell(15, 11).value ?? '')).not.toMatch(/Issac/i);
  });
});

describe('Form T Karnataka payroll column mapping', () => {
  const wageHeaders = [
    'Subsistence Allowance (24)',
    'Total (25)',
    'ESI (26)',
    'Salary Advance',
    'Fines',
    'Damages/Loss',
    'Others',
    'Total',
    'Net Amount Payable',
    'Mode of Payment Cash/ Cheque No.',
    'Total OT hours',
  ];

  test('Deductions Total = gross_pay − net_pay', () => {
    expect(computeFormTSEKarnatakaTotalDeductions(107741, 103828)).toBe(3913);
    expect(computeFormTSEKarnatakaTotalDeductions('10,000', '9,250')).toBe(750);
    expect(computeFormTSEKarnatakaTotalDeductions(9000, 10000)).toBe('');
    expect(computeFormTSEKarnatakaTotalDeductions('', 1000)).toBe('');
  });

  test('Mode of Payment defaults to Bank Transfer', () => {
    expect(FORM_T_KA_DEFAULT_PAYMENT_MODE).toBe('Bank Transfer');
  });

  test('Total OT hours defaults to NIL', () => {
    expect(FORM_T_KA_OT_HOURS_NIL).toBe('NIL');
    expect(isFormTSEKarnatakaTotalOtHoursHeader('Total OT hours')).toBe(true);
    expect(isFormTSEKarnatakaTotalOtHoursHeader('Total OT hours (12)')).toBe(true);
    expect(isFormTSEKarnatakaTotalOtHoursHeader('No. of payable days')).toBe(false);
    const rows = applyFormTSEKarnatakaOtHoursNilToMappedRows(
      [{ 'Total OT hours': '10:58', Name: 'A' }, { 'Total OT hours': '08:40', Name: 'B' }],
      ['Name', 'Total OT hours']
    );
    expect(rows[0]['Total OT hours']).toBe('NIL');
    expect(rows[1]['Total OT hours']).toBe('NIL');
    const exported = prepareFormTSEExportRows(
      [{ 'Total OT hours': '10:58', 'No. of payable days': '30' }],
      ['No. of payable days', 'Total OT hours'],
      ['No. of payable days', 'Total OT hours']
    );
    expect(exported[0]['Total OT hours']).toBe('NIL');
    expect(exported[0]['No. of payable days']).toBe('30');
  });

  test('does not treat Deduction-band Total as earned-wages Total', () => {
    expect(resolveFormTSEKarnatakaEarningsTotalHeader(wageHeaders)).toBe('Total (25)');
    expect(resolveFormTSEKarnatakaDeductionTotalHeader(wageHeaders)).toBe('Total');
    expect(resolveFormTSEKarnatakaNetAmountPayableHeader(wageHeaders)).toBe('Net Amount Payable');
  });

  test('Earned wages and other allowances / Earnings Total ← gross_pay column', () => {
    expect(
      resolveFormTSEKarnatakaEarningsTotalHeader([
        'Medical Allowance',
        'Special allowance',
        'Earned wages and other allowances',
        'ESI',
      ])
    ).toBe('Earned wages and other allowances');
    expect(
      resolveFormTSEKarnatakaEarningsTotalHeader([
        'Basic',
        'HRA',
        'Earnings Total',
        'Net Amount Payable',
      ])
    ).toBe('Earnings Total');
    expect(isFormTSEKarnatakaEarningsTotalHeader('Medical Allowance')).toBe(false);
    expect(isFormTSEKarnatakaEarningsTotalHeader('Special allowance')).toBe(false);
    expect(isFormTSEKarnatakaEarningsTotalHeader('Earned wages and other allowances')).toBe(true);
  });

  test('bare Total before Net Amount Payable is Deductions Total', () => {
    const headers = ['Others', 'Total', 'Net Amount Payable', 'Mode of Payment Cash/ Cheque No.'];
    expect(resolveFormTSEKarnatakaDeductionTotalHeader(headers)).toBe('Total');
    expect(resolveFormTSEKarnatakaEarningsTotalHeader(headers)).toBeNull();
    expect(resolveFormTSEKarnatakaNetAmountPayableHeader(headers)).toBe('Net Amount Payable');
  });

  test('UAN / ESI stay as plain text (no 1.01E+11)', () => {
    expect(formatFormTSERegistrationIdForExcel(1.01e11)).toBe('101000000000');
    expect(formatFormTSERegistrationIdForExcel('1.01E+11')).toBe('101000000000');
    expect(formatFormTSERegistrationIdForExcel('100123456789')).toBe('100123456789');
  });

  test('download prefers fuller original template headers over truncated modal headers', () => {
    const modalTruncated = [
      'S.NO',
      'Name of Employee',
      'ATTENDANCE_1',
      'ATTENDANCE_2',
      'ATTENDANCE_3',
      'Basic',
      'Medical Allowance',
    ];
    const templateFull = [
      ...modalTruncated,
      'Special allowance',
      'OT',
      'NFH',
      'Total (25)',
      'Net Amount Payable',
      'Mode of Payment Cash/ Cheque No.',
    ];
    expect(resolveFormTSEDownloadWriteHeaders(modalTruncated, templateFull)).toEqual(templateFull);
  });
});
