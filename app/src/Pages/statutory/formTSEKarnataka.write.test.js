import ExcelJS from 'exceljs';
import JSZip from 'jszip';
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
  prepareFormTSEDownloadHeaderData,
  applyFormTSEKarnatakaAutofillFromSite,
  computeFormTSEKarnatakaTotalDeductions,
  FORM_T_KA_DEFAULT_PAYMENT_MODE,
  FORM_T_KA_OT_HOURS_NIL,
  FORM_T_KA_RULE_CITATION,
  FORM_T_KA_HEADER_IDENTITY_FONT_SIZE,
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
  getFormTSEKarnatakaEmployeeNameFromRow,
  applyFormTSEKarnatakaEmployeeToRow,
  forceFormTSEIdentityCellsInSheetXml,
  sanitizeFormTSEAttendanceMark,
  overlayFormTSEPayrollOntoRowsByName,
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
    expect(Number(ws.getCell(10, 1).font?.size)).toBe(FORM_T_KA_HEADER_IDENTITY_FONT_SIZE);
    expect(Number(ws.getCell(11, 1).font?.size)).toBe(FORM_T_KA_HEADER_IDENTITY_FONT_SIZE);

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

  it('writes the last employee name (not Father) when the Date row has a Name+Father merge', async () => {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(await buildMinimalFormTTemplateBuffer());
    const templateWs = wb.getWorksheet('Form T');
    // Template Date: row sits at Excel 20. A B:C merge makes ExcelJS write Father into Name.
    templateWs.mergeCells(20, 2, 20, 3);
    const templateArrayBuffer = await wb.xlsx.writeBuffer();

    const sixRows = [
      {
        'S.NO': 1,
        'Name of Employee': 'Vinay Kumar',
        "Father / Husband's Name": 'Kadesh j kamble',
        Gender: 'Male',
        'Designation / Department': 'Engineer',
        'Date of Joining': '17 Aug 2026',
      },
      {
        'S.NO': 2,
        'Name of Employee': 'Issac Kanagaraj',
        "Father / Husband's Name": 'abc',
        Gender: 'Male',
        'Designation / Department': 'Assistant Manager',
        'Date of Joining': '01 Dec 2025',
      },
      {
        'S.NO': 3,
        'Name of Employee': 'Saravanan',
        "Father / Husband's Name": 'abc',
        Gender: 'Male',
        'Designation / Department': 'Engineer',
        'Date of Joining': '01 Dec 2025',
      },
      {
        'S.NO': 4,
        'Name of Employee': 'Ameerkhan',
        "Father / Husband's Name": 'abc',
        Gender: 'Male',
        'Designation / Department': 'Engineer',
        'Date of Joining': '01 Dec 2025',
      },
      {
        'S.NO': 5,
        'Name of Employee': 'Rajesh',
        "Father / Husband's Name": 'abc',
        Gender: 'Male',
        'Designation / Department': 'Engineer',
        'Date of Joining': '01 Dec 2025',
      },
      {
        'S.NO': 6,
        'Name of Employee': 'Ashok',
        "Father / Husband's Name": 'abc',
        Gender: 'Male',
        'Designation / Department': 'Junior Engineer',
        'Date of Joining': '01 Dec 2025',
      },
    ];

    expect(getFormTSEKarnatakaEmployeeNameFromRow(sixRows[5], modalHeaders)).toBe('Ashok');

    const { blob } = await buildFormTSEWorkbookWithTemplateStyles({
      templateArrayBuffer,
      mappedData: sixRows,
      headersToUse: modalHeaders,
      parsedHeaderRowIndex: 11,
      parsedDataStartIndex: 14,
      parsedTableStartCol: 0,
      parsedFormHeader: { title: 'Form T' },
      headerFormData: {
        form_t_month_year: 'April 2026',
        form_t_establishment_name_address: 'Babaleshwar Hero Site Karnataka',
        form_t_employer: 'M/s Clean Wind Power Bableshwar Pvt Ltd',
      },
      formFileName: 'Form_T_KA.xlsx',
      sheetNameHint: 'Form T',
    });

    const outWb = new ExcelJS.Workbook();
    await outWb.xlsx.load(await new Response(blob).arrayBuffer());
    const ws = outWb.getWorksheet('Form T');
    expect(String(ws.getCell(15, 2).value ?? '')).toMatch(/Vinay Kumar/i);
    expect(String(ws.getCell(20, 1).value ?? '').trim()).toBe('6');
    expect(String(ws.getCell(20, 2).value ?? '')).toMatch(/Ashok/i);
    expect(String(ws.getCell(20, 2).value ?? '')).not.toMatch(/^abc$/i);
    expect(String(ws.getCell(20, 3).value ?? '')).toMatch(/^abc$/i);
    expect(String(ws.getCell(20, 5).value ?? '')).toMatch(/Junior Engineer/i);
    let dateRow = 0;
    for (let r = 20; r <= 28; r += 1) {
      if (/^date\s*:/i.test(String(ws.getCell(r, 1).value ?? '').trim())) {
        dateRow = r;
        break;
      }
    }
    expect(dateRow).toBeGreaterThan(20);
  });

  it('writes the 6th employee name when column B is labelled principal employer', async () => {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(await buildMinimalFormTTemplateBuffer());
    const templateWs = wb.getWorksheet('Form T');
    templateWs.getCell(12, 2).value = 'Name and address of principal employer';
    templateWs.mergeCells(20, 2, 20, 3);
    const templateArrayBuffer = await wb.xlsx.writeBuffer();

    const clraHeaders = [
      'S.NO',
      'Name and address of principal employer',
      "Father / Husband's Name",
      'Gender',
      'Designation / Department',
      'Date of Joining',
    ];
    const people = ['Vinay Kumar', 'Issac Kanagaraj', 'Saravanan', 'Ameerkhan', 'Rajesh', 'Ashok'];
    const sixRows = people.map((name, i) => ({
      'S.NO': i + 1,
      'Name and address of principal employer': name,
      "Father / Husband's Name": i === 0 ? 'Kadesh j kamble' : 'abc',
      Gender: 'Male',
      'Designation / Department': i === 5 ? 'Junior Engineer' : 'Engineer',
      'Date of Joining': '01 Dec 2025',
    }));

    expect(getFormTSEKarnatakaEmployeeNameFromRow(sixRows[5], clraHeaders)).toBe('Ashok');

    const { blob } = await buildFormTSEWorkbookWithTemplateStyles({
      templateArrayBuffer,
      mappedData: sixRows,
      headersToUse: clraHeaders,
      parsedHeaderRowIndex: 11,
      parsedDataStartIndex: 14,
      parsedTableStartCol: 0,
      parsedFormHeader: { title: 'Form T' },
      headerFormData: {
        form_t_month_year: 'April 2026',
        form_t_establishment_name_address: 'Babaleshwar Hero Site Karnataka',
        form_t_employer: 'M/s Clean Wind Power Bableshwar Pvt Ltd',
      },
      formFileName: 'Form_T_KA.xlsx',
      sheetNameHint: 'Form T',
    });

    const outWb = new ExcelJS.Workbook();
    await outWb.xlsx.load(await new Response(blob).arrayBuffer());
    const ws = outWb.getWorksheet('Form T');
    expect(String(ws.getCell(12, 2).value ?? '')).toMatch(/Name of Employee/i);
    expect(String(ws.getCell(12, 2).value ?? '')).not.toMatch(/principal employer/i);
    expect(String(ws.getCell(11, 1).value ?? '')).toMatch(/Clean Wind/i);
    expect(String(ws.getCell(20, 1).value ?? '').trim()).toBe('6');
    expect(String(ws.getCell(20, 2).value ?? '')).toMatch(/Ashok/i);
    expect(String(ws.getCell(20, 2).value ?? '')).not.toMatch(/^abc$/i);
    expect(String(ws.getCell(20, 3).value ?? '')).toMatch(/^abc$/i);
  });

  it('writes Ashok (not Father abc) on serial 6 when official template has empty row 14 and Date B:C merge', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Form T');
    ws.getCell(1, 1).value = 'FORM T';
    ws.getCell(2, 1).value = 'COMBINED MUSTER ROLL CUM REGISTER OF WAGES';
    ws.getCell(3, 1).value =
      '[See Rule 24(9-B) of Karnataka Shops & Commercial Establishment Rules, 1963]';
    ws.getCell(9, 1).value = 'Month / Year : April 2026';
    ws.getCell(10, 1).value = 'Address of the Establishment : Karjol Village';
    ws.getCell(11, 1).value = 'Name and Address of the Employer';
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
    for (let c = 1; c <= 7; c += 1) ws.getCell(13, c).value = c;
    ws.getCell(20, 1).value = 'Date:';
    ws.mergeCells(20, 2, 20, 3);
    const templateArrayBuffer = await wb.xlsx.writeBuffer();

    const sixRows = [
      'Vinay Kumar',
      'Issac Kanagaraj',
      'Saravanan',
      'Ameerkhan',
      'Rajesh',
      'Ashok',
    ].map((name, i) => ({
      'S.NO': i + 1,
      'Name of Employee': name,
      "Father / Husband's Name": i === 0 ? 'Kadesh j kamble' : 'abc',
      Gender: 'Male',
      'Designation / Department': i === 5 ? 'Junior Engineer' : 'Engineer',
      'Date of Joining': '01 Dec 2025',
    }));

    const { blob } = await buildFormTSEWorkbookWithTemplateStyles({
      templateArrayBuffer,
      mappedData: sixRows,
      headersToUse: modalHeaders,
      parsedHeaderRowIndex: 11,
      parsedDataStartIndex: 14,
      parsedTableStartCol: 0,
      parsedFormHeader: { title: 'Form T' },
      headerFormData: {
        form_t_month_year: 'April 2026',
        form_t_establishment_name_address: 'Babaleshwar Hero Site Karnataka',
        form_t_employer: 'M/s Clean Wind Power Bableshwar Pvt Ltd',
      },
      formFileName: 'Form_T_KA.xlsx',
      sheetNameHint: 'Form T',
    });

    const outWb = new ExcelJS.Workbook();
    await outWb.xlsx.load(await new Response(blob).arrayBuffer());
    const out = outWb.getWorksheet('Form T');
    let serialSix = 0;
    for (let r = 14; r <= 28; r += 1) {
      if (String(out.getCell(r, 1).value ?? '').trim() === '6') {
        serialSix = r;
        break;
      }
    }
    expect(serialSix).toBeGreaterThan(0);
    expect(String(out.getCell(serialSix, 2).value ?? '')).toMatch(/Ashok/i);
    expect(String(out.getCell(serialSix, 2).value ?? '')).not.toMatch(/^abc$/i);
    expect(String(out.getCell(serialSix, 3).value ?? '')).toMatch(/^abc$/i);
    expect(String(out.getCell(serialSix, 5).value ?? '')).toMatch(/Junior Engineer/i);
  });

  it('writes Ponnusamy on serial 6 when more employees follow (row 20 merge is not the last row)', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Form T');
    ws.getCell(1, 1).value = 'FORM T';
    ws.getCell(2, 1).value = 'COMBINED MUSTER ROLL CUM REGISTER OF WAGES';
    ws.getCell(3, 1).value =
      '[See Rule 24(9-B) of Karnataka Shops & Commercial Establishment Rules, 1963]';
    ws.getCell(9, 1).value = 'Month / Year : July 2026';
    ws.getCell(11, 1).value = 'Name and Address of the Employer';
    [
      'S.NO',
      'Name of Employee',
      "Father / Husband's Name",
      'Gender',
      'Designation / Department',
      'Date of Joining',
    ].forEach((h, i) => {
      ws.getCell(12, i + 1).value = h;
    });
    for (let c = 1; c <= 7; c += 1) ws.getCell(14, c).value = c;
    ws.getCell(20, 1).value = 'Date:';
    ws.mergeCells(20, 2, 20, 3);
    const templateArrayBuffer = await wb.xlsx.writeBuffer();

    const people = [
      ['Prakash', 'Lokappa', 'Engineer'],
      ['Umesh', 'Shivaputrappa oli', 'Senior Technician'],
      ['Naveen', 'Manjappa K', 'Engineer'],
      ['Vinayak P', 'Parashuram V hadapad', 'Junior Engineer'],
      ['Bharat', 'Vasudev kolkar', 'Junior Engineer'],
      ['Ponnusamy', 'abc', 'Senior Engineer'],
      ['Eliza', 'abc', 'Engineer'],
      ['Basavaraj', 'abc', 'Senior Engineer'],
      ['Manoj', 'Duraiswamy H', 'Deputy Manager'],
      ['PrabakaranKumar', 'Jeyakumar P', 'Associate Master Technician'],
      ['Sureshkumar', 'Ranganathan', 'Assistant Manager'],
    ];
    const rows = people.map((p, i) => ({
      'S.NO': i + 1,
      'Name of Employee': p[0],
      "Father / Husband's Name": p[1],
      Gender: 'Male',
      'Designation / Department': p[2],
      'Date of Joining': '01 Dec 2025',
    }));

    const { blob } = await buildFormTSEWorkbookWithTemplateStyles({
      templateArrayBuffer,
      mappedData: rows,
      headersToUse: modalHeaders,
      parsedHeaderRowIndex: 11,
      parsedDataStartIndex: 14,
      parsedTableStartCol: 0,
      parsedFormHeader: { title: 'Form T' },
      formFileName: 'Form_T_KA.xlsx',
      sheetNameHint: 'Form T',
    });

    const outWb = new ExcelJS.Workbook();
    await outWb.xlsx.load(await new Response(blob).arrayBuffer());
    const out = outWb.getWorksheet('Form T') || outWb.worksheets[0];
    let serialSix = 0;
    for (let r = 14; r <= 40; r += 1) {
      if (String(out.getCell(r, 1).value ?? '').trim() === '6') {
        serialSix = r;
        break;
      }
    }
    expect(serialSix).toBeGreaterThan(0);
    expect(String(out.getCell(serialSix, 2).value ?? '')).toMatch(/Ponnusamy/i);
    expect(String(out.getCell(serialSix, 2).value ?? '')).not.toMatch(/^abc$/i);
    expect(String(out.getCell(serialSix, 3).value ?? '')).toMatch(/^abc$/i);
    expect(String(out.getCell(serialSix, 5).value ?? '')).toMatch(/Senior Engineer/i);
    expect(out.getCell(serialSix, 3).alignment?.horizontal || 'center').toBe('center');
    const sampleFatherAlign = out.getCell(serialSix > 15 ? serialSix - 1 : serialSix + 1, 3)
      .alignment?.horizontal;
    if (sampleFatherAlign) {
      expect(out.getCell(serialSix, 3).alignment?.horizontal).toBe(sampleFatherAlign);
    }

    const zip = await JSZip.loadAsync(await new Response(blob).arrayBuffer());
    const sheetPath = Object.keys(zip.files).find((p) => /xl\/worksheets\/sheet\d+\.xml$/i.test(p));
    const xml = await zip.file(sheetPath).async('string');
    expect(xml).not.toMatch(/ref="B20:C20"/i);
    expect(xml).toMatch(/Ponnusamy/i);
    const b20 = xml.match(/<c\b(?=[^>]*\br=["']B20["'])[^>]*>[\s\S]*?<\/c>/i);
    expect(b20 && b20[0]).toMatch(/Ponnusamy/i);
    expect(b20 && b20[0]).not.toMatch(/>abc</i);
    const c19open = xml.match(/<c\b(?=[^>]*\br=["']C19["'])[^>]*>/i);
    const c20open = xml.match(/<c\b(?=[^>]*\br=["']C20["'])[^>]*>/i);
    const s19 = c19open && c19open[0].match(/\bs="([^"]+)"/i);
    const s20 = c20open && c20open[0].match(/\bs="([^"]+)"/i);
    if (s19) expect(s20 && s20[1]).toBe(s19[1]);
  });

  it('XML last-pass writes Ponnusamy into B20 even when ExcelJS stored abc as t=s with s before r', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Form T');
    ws.getCell(1, 1).value = 'FORM T';
    ws.getCell(12, 1).value = 'S.NO';
    ws.getCell(12, 2).value = 'Name of Employee';
    ws.getCell(12, 3).value = "Father / Husband's Name";
    ws.getCell(20, 1).value = 6;
    ws.getCell(20, 2).value = 'abc';
    ws.mergeCells(20, 2, 20, 3);
    const buf = await wb.xlsx.writeBuffer();
    const zip0 = await JSZip.loadAsync(buf);
    const sheetPath = Object.keys(zip0.files).find((p) => /xl\/worksheets\/sheet[^/]*\.xml$/i.test(p));
    let xml = await zip0.file(sheetPath).async('string');
    xml = xml.replace(
      /<c\b(?=[^>]*\br=["']B20["'])[^>]*>[\s\S]*?<\/c>|<c\b(?=[^>]*\br=["']B20["'])[^>]*\/>/i,
      '<c s="5" r="B20" t="s"><v>0</v></c>'
    );
    zip0.file(sheetPath, xml);
    const dirty = await zip0.generateAsync({ type: 'arraybuffer' });

    const people = [
      ['Prakash', 'Lokappa'],
      ['Umesh', 'Shivaputrappa oli'],
      ['Naveen', 'Manjappa K'],
      ['Vinayak P', 'Parashuram V hadapad'],
      ['Bharat', 'Vasudev kolkar'],
      ['Ponnusamy', 'abc'],
      ['Eliza', 'abc'],
    ].map((p, i) => ({
      'S.NO': i + 1,
      'Name of Employee': p[0],
      "Father / Husband's Name": p[1],
    }));

    const patched = await forceFormTSEIdentityCellsInSheetXml(dirty, people, modalHeaders, []);
    const zip1 = await JSZip.loadAsync(patched);
    const xml1 = await zip1.file(sheetPath).async('string');
    expect(xml1).not.toMatch(/ref="B20:C20"/i);
    const b20 = xml1.match(/<c\b(?=[^>]*\br=["']B20["'])[^>]*>[\s\S]*?<\/c>/i);
    expect(b20 && b20[0]).toMatch(/Ponnusamy/i);
    expect(b20 && b20[0]).not.toMatch(/>abc</i);

    const second = await prepareFormTSEWorkbookForDownload(patched, {
      force: true,
      mappedData: people,
      headers: modalHeaders,
    });
    const zip2 = await JSZip.loadAsync(second);
    const sheetPath2 = Object.keys(zip2.files).find((p) => /xl\/worksheets\/sheet[^/]*\.xml$/i.test(p));
    const xml2 = await zip2.file(sheetPath2).async('string');
    expect(xml2).not.toMatch(/ref="B20:C20"/i);
    const b20b = xml2.match(/<c\b(?=[^>]*\br=["']B20["'])[^>]*>[\s\S]*?<\/c>/i);
    expect(b20b && b20b[0]).toMatch(/Ponnusamy/i);
    expect(b20b && b20b[0]).not.toMatch(/>abc</i);
  });

  it('overwrites leaked Father abc in Name when Autofill applies People Ashok', () => {
    const row = {
      'S.NO': 6,
      'Name of Employee': 'abc',
      "Father / Husband's Name": 'abc',
      Gender: 'Male',
      'Designation / Department': 'Junior Engineer',
    };
    applyFormTSEKarnatakaEmployeeToRow(
      row,
      { FirstName: 'Ashok', Father_s_Name: 'abc', Sex: 'Male', Designation: 'Junior Engineer' },
      modalHeaders,
      { onlyEmpty: true }
    );
    expect(row['Name of Employee']).toMatch(/Ashok/i);
    expect(row['Name of Employee']).not.toMatch(/^abc$/i);
  });

  it('restamps the 6th name from Autofill lookup when the saved sheet still has abc', async () => {
    const dirty = new ExcelJS.Workbook();
    await dirty.xlsx.load(await buildMinimalFormTTemplateBuffer());
    const dirtyWs = dirty.getWorksheet('Form T');
    dirtyWs.getCell(12, 2).value = 'Name of Employee';
    dirtyWs.mergeCells(20, 2, 20, 3);
    dirtyWs.getCell(15, 1).value = 1;
    dirtyWs.getCell(15, 2).value = 'Vinay Kumar';
    dirtyWs.getCell(15, 3).value = 'Kadesh j kamble';
    dirtyWs.getCell(16, 1).value = 2;
    dirtyWs.getCell(16, 2).value = 'Issac Kanagaraj';
    dirtyWs.getCell(16, 3).value = 'abc';
    dirtyWs.getCell(17, 1).value = 3;
    dirtyWs.getCell(17, 2).value = 'Saravanan';
    dirtyWs.getCell(17, 3).value = 'abc';
    dirtyWs.getCell(18, 1).value = 4;
    dirtyWs.getCell(18, 2).value = 'Ameerkhan';
    dirtyWs.getCell(18, 3).value = 'abc';
    dirtyWs.getCell(19, 1).value = 5;
    dirtyWs.getCell(19, 2).value = 'Rajesh';
    dirtyWs.getCell(19, 3).value = 'abc';
    dirtyWs.getCell(20, 1).value = 6;
    dirtyWs.getCell(20, 2).value = 'abc';
    dirtyWs.getCell(20, 4).value = 'Male';
    dirtyWs.getCell(20, 5).value = 'Junior Engineer';
    const dirtyBuf = await dirty.xlsx.writeBuffer();

    const liveRows = [
      { 'S.NO': 1, 'Name of Employee': 'Vinay Kumar', "Father / Husband's Name": 'Kadesh j kamble' },
      { 'S.NO': 2, 'Name of Employee': 'Issac Kanagaraj', "Father / Husband's Name": 'abc' },
      { 'S.NO': 3, 'Name of Employee': 'Saravanan', "Father / Husband's Name": 'abc' },
      { 'S.NO': 4, 'Name of Employee': 'Ameerkhan', "Father / Husband's Name": 'abc' },
      { 'S.NO': 5, 'Name of Employee': 'Rajesh', "Father / Husband's Name": 'abc' },
      {
        'S.NO': 6,
        'Name of Employee': 'abc',
        "Father / Husband's Name": '',
        Gender: 'Male',
        'Designation / Department': 'Junior Engineer',
        __employeeLookupName: 'Ashok',
      },
    ];
    expect(getFormTSEKarnatakaEmployeeNameFromRow(liveRows[5], modalHeaders)).toBe('Ashok');

    const ready = await prepareFormTSEWorkbookForDownload(dirtyBuf, {
      force: true,
      mappedData: liveRows,
      headers: modalHeaders,
      employees: [{ FirstName: 'Ashok' }],
    });
    const outWb = new ExcelJS.Workbook();
    await outWb.xlsx.load(ready);
    const ws = outWb.getWorksheet('Form T') || outWb.worksheets[0];
    expect(String(ws.getCell(20, 2).value ?? '')).toMatch(/Ashok/i);
    expect(String(ws.getCell(20, 2).value ?? '')).not.toMatch(/^abc$/i);
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
    expect(Number(ws.getCell(10, 1).font?.size)).toBe(FORM_T_KA_HEADER_IDENTITY_FONT_SIZE);
    expect(Number(ws.getCell(11, 1).font?.size)).toBe(FORM_T_KA_HEADER_IDENTITY_FONT_SIZE);
  });

  it('fills Establishment/Employer from sibling statutory keys when form_t keys are blank', () => {
    const out = applyFormTSEKarnatakaAutofillFromSite(
      {
        statutory_establishment_name_address: 'Bableshwar Hero Site, Karjol Village',
        statutory_employer_name_address: 'M/s Clean Wind Power Bableshwar Pvt Ltd',
      },
      { monthYearText: '', establishmentText: '', employerText: '' }
    );
    expect(out.form_t_establishment_name_address).toMatch(/Bableshwar Hero Site/);
    expect(out.form_t_employer).toMatch(/Clean Wind Power/);
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

  it('prepareFormTSEDownloadHeaderData keeps modal Establishment/Employer over empty site overlay', () => {
    const out = prepareFormTSEDownloadHeaderData(
      {
        form_t_month_year: 'June 2026',
        form_t_establishment_name_address:
          'Babaleshwar Hero Site, 33/220KV Substation, Survey No: 132',
        form_t_employer:
          'M/s Clean Wind Power Bableshwar Pvt Ltd, 33/220KV Substation, Survey No: 132, Karjol Village, Bijapur Dist - 586113',
      },
      null,
      { monthYearText: '', establishmentText: '', employerText: '' }
    );
    expect(out.form_t_establishment_name_address).toMatch(/Babaleshwar Hero Site/);
    expect(out.form_t_employer).toMatch(/Clean Wind Power/);
  });

  it('prepareFormTSEDownloadHeaderData fills blank Establishment/Employer from site/company', () => {
    const out = prepareFormTSEDownloadHeaderData(
      { form_t_month_year: 'July 2026' },
      {
        fields: [
          { label: 'Name and address of the Establishment', key: 'form_t_establishment_name_address' },
          { label: 'Name and Address of employer', key: 'form_t_employer' },
        ],
      },
      {
        monthYearText: 'July 2026',
        establishmentText: 'Babaleshwar Hero Site, Karjol Village, Bijapur Dist',
        employerText: 'M/s Clean Wind Power Bableshwar Pvt Ltd, Karjol Village',
      }
    );
    expect(out.form_t_establishment_name_address).toMatch(/Babaleshwar Hero Site/);
    expect(out.form_t_employer).toMatch(/Clean Wind Power/);
    expect(out.statutory_establishment_name_address).toMatch(/Babaleshwar Hero Site/);
    expect(out.statutory_employer_name_address).toMatch(/Clean Wind Power/);
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

  it('does not put payroll into empty attendance days when the employee joined mid-month', async () => {
    expect(sanitizeFormTSEAttendanceMark('31')).toBe('');
    expect(sanitizeFormTSEAttendanceMark('NIL')).toBe('');
    expect(sanitizeFormTSEAttendanceMark('18563')).toBe('');
    expect(sanitizeFormTSEAttendanceMark('P')).toBe('P');
    expect(sanitizeFormTSEAttendanceMark('WO')).toBe('WO');

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
    const attendance = Array.from({ length: 31 }, (_, i) => `ATTENDANCE_${i + 1}`);
    const exportHeaders = [
      ...identity,
      ...attendance,
      'No. of payable days',
      'Total OT hours',
      'Basic',
    ];
    const row = {
      'S.NO': 1,
      'Name of Employee': 'Prakash',
      "Father / Husband's Name": 'Lokappa',
      Gender: 'Male',
      'Designation / Department': 'Engineer',
      'Date of Joining': '15 Jul 2026',
      ATTENDANCE_15: 'P',
      ATTENDANCE_16: 'WO',
      'No. of payable days': '31',
      'Total OT hours': 'NIL',
      Basic: '18563',
    };
    const prepared = prepareFormTSEExportRows([row], exportHeaders, exportHeaders);
    expect(prepared[0]['ATTENDANCE_1']).toBe('');
    expect(prepared[0]['ATTENDANCE_11']).toBe('');
    expect(prepared[0]['ATTENDANCE_14']).toBe('');
    expect(prepared[0]['ATTENDANCE_15']).toBe('P');
    expect(prepared[0]['ATTENDANCE_16']).toBe('WO');
    expect(String(prepared[0]['No. of payable days'])).toBe('31');
    expect(prepared[0]['Total OT hours']).toBe('NIL');
    expect(String(prepared[0].Basic)).toBe('18563');

    const { blob } = await buildFormTSEWorkbookWithTemplateStyles({
      templateArrayBuffer: await buildMinimalFormTTemplateBuffer(),
      mappedData: [row],
      headersToUse: exportHeaders,
      parsedHeaderRowIndex: 11,
      parsedDataStartIndex: 14,
      parsedTableStartCol: 0,
      parsedFormHeader: { title: 'Form T' },
      formFileName: 'Form_T_KA.xlsx',
      sheetNameHint: 'Form T',
    });
    const outWb = new ExcelJS.Workbook();
    await outWb.xlsx.load(await new Response(blob).arrayBuffer());
    const ws = outWb.getWorksheet('Form T') || outWb.worksheets[0];
    const day1Col = FORM_T_KA_ATTENDANCE_START_COL0 + 1;
    expect(String(ws.getCell(15, day1Col).value ?? '').trim()).toBe('');
    expect(String(ws.getCell(15, day1Col + 10).value ?? '').trim()).toBe('');
    expect(String(ws.getCell(15, day1Col + 14).value ?? '')).toMatch(/^P$/i);
    expect(String(ws.getCell(15, day1Col + 15).value ?? '')).toMatch(/^WO$/i);
  });

  it('puts payable days / Basic / HRA / Total / PF / PT on the row matching first and last name', () => {
    const headers = [
      'S.NO',
      'Name of Employee',
      "Father / Husband's Name",
      'No. of payable days',
      'Basic',
      'HRA',
      'Total (25)',
      'PF',
      'PT',
    ];
    const autofill = [
      {
        'Name of Employee': 'Umesh Shivaputrappa',
        'No. of payable days': '31',
        Basic: '100',
        HRA: '50',
        'Total (25)': '200',
        PF: '12',
        PT: '0',
      },
      {
        'Name of Employee': 'Prakash Lokappa',
        'No. of payable days': '16',
        Basic: '18563',
        HRA: '2228',
        'Total (25)': '37126',
        PF: '1800',
        PT: '200',
      },
    ];
    const target = [
      { 'S.NO': 1, 'Name of Employee': 'Prakash Lokappa' },
      { 'S.NO': 2, 'Name of Employee': 'Umesh Shivaputrappa' },
    ];
    const out = overlayFormTSEPayrollOntoRowsByName(target, headers, autofill);
    expect(String(out[0]['No. of payable days'])).toBe('16');
    expect(String(out[0].Basic)).toBe('18563');
    expect(String(out[0].HRA)).toBe('2228');
    expect(String(out[0]['Total (25)'])).toBe('37126');
    expect(String(out[0].PF)).toBe('1800');
    expect(String(out[0].PT)).toBe('200');
    expect(String(out[1].Basic)).toBe('100');

    const prepared = prepareFormTSEExportRows(target, headers, headers);
    const named = overlayFormTSEPayrollOntoRowsByName(prepared, headers, autofill);
    expect(String(named[0].Basic)).toBe('18563');
    expect(String(named[0]['ATTENDANCE_1'] || '')).toBe('');
  });

  it('clears payable days / Basic / HRA / Total / PF / PT when the employee has no attendance that month', () => {
    const attendance = Array.from({ length: 31 }, (_, i) => `ATTENDANCE_${i + 1}`);
    const headers = [
      'Name of Employee',
      ...attendance,
      'No. of payable days',
      'Total OT hours',
      'Basic',
      'HRA',
      'Total (25)',
      'PF',
      'PT',
    ];
    const row = {
      'Name of Employee': 'Prakash Lokappa',
      'No. of payable days': '31',
      'Total OT hours': 'NIL',
      Basic: '18563',
      HRA: '2228',
      'Total (25)': '37126',
      PF: '1800',
      PT: '200',
    };
    const prepared = prepareFormTSEExportRows([row], headers, headers);
    expect(prepared[0]['No. of payable days']).toBe('');
    expect(prepared[0].Basic).toBe('');
    expect(prepared[0].HRA).toBe('');
    expect(prepared[0]['Total (25)']).toBe('');
    expect(prepared[0].PF).toBe('');
    expect(prepared[0].PT).toBe('');
    expect(prepared[0]['Total OT hours']).toBe('NIL');

    const overlayTarget = {
      'Name of Employee': 'Prakash Lokappa',
      ATTENDANCE_1: '',
      ATTENDANCE_26: '',
    };
    const overlayHeaders = [
      'Name of Employee',
      'ATTENDANCE_1',
      'ATTENDANCE_26',
      'No. of payable days',
      'Basic',
      'HRA',
      'PF',
      'PT',
    ];
    const out = overlayFormTSEPayrollOntoRowsByName(
      [overlayTarget],
      overlayHeaders,
      [
        {
          'Name of Employee': 'Prakash Lokappa',
          ATTENDANCE_1: 'P',
          'No. of payable days': '31',
          Basic: '18563',
          HRA: '2228',
          PF: '1800',
          PT: '200',
        },
      ]
    );
    expect(String(out[0]['No. of payable days'] || '')).toBe('');
    expect(String(out[0].Basic || '')).toBe('');
    expect(String(out[0].HRA || '')).toBe('');
    expect(String(out[0].PF || '')).toBe('');
    expect(String(out[0].PT || '')).toBe('');
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

describe('Form T Karnataka export grid (Remarks cap + wage groups)', () => {
  const boxBorder = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thin' },
  };

  function mergeCoversRow(ws, row, colFrom, colTo) {
    const merges = Array.isArray(ws.model?.merges) ? ws.model.merges : [];
    return merges.some((label) => {
      const parts = String(label || '').split(':');
      if (parts.length !== 2) return false;
      const tl = ws.getCell(parts[0]);
      const br = ws.getCell(parts[1]);
      return tl.row === row && br.row === row && tl.col === colFrom && br.col === colTo;
    });
  }

  it('strips leftover BS–CB boxes after Remarks (col 70)', () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Form T');
    ws.getCell(1, 1).value = 'FORM T';
    ws.getCell(2, 1).value = 'COMBINED MUSTER ROLL CUM REGISTER OF WAGES';
    ws.getCell(9, 1).value = 'Month / Year';
    ws.getCell(12, 1).value = 'S.NO';
    ws.getCell(12, 2).value = 'Name of Employee';
    ws.getCell(12, 70).value = 'Remarks';
    for (let r = 12; r <= 16; r += 1) {
      for (let c = 70; c <= 80; c += 1) {
        ws.getCell(r, c).border = { ...boxBorder };
      }
    }
    ws.getCell(14, 71).value = 41;
    ws.getCell(14, 80).value = 50;

    finalizeFormTSEKarnatakaWorksheetExportBorders(ws, { colTo: 80, tableLastRow: 16 });

    expect(String(ws.getCell(12, 70).value ?? '')).toMatch(/^Remarks$/i);
    expect(excelJSCellHasBorder(ws.getCell(12, 70))).toBe(true);
    expect(excelJSCellHasBorder(ws.getCell(12, 71))).toBe(false);
    expect(excelJSCellHasBorder(ws.getCell(12, 80))).toBe(false);
    expect(excelJSCellHasBorder(ws.getCell(15, 80))).toBe(false);
    expect(String(ws.getCell(14, 71).value ?? '').trim()).toBe('');
    expect(String(ws.getCell(12, 80).value ?? '').trim()).toBe('');
  });

  it('merges Earned wages and Deductions groups including TDS', () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Form T');
    ws.getCell(1, 1).value = 'FORM T';
    ws.getCell(2, 1).value = 'COMBINED MUSTER ROLL CUM REGISTER OF WAGES';
    ws.getCell(9, 1).value = 'Month / Year';
    ws.getCell(12, 1).value = 'S.NO';
    ws.getCell(12, 2).value = 'Name of Employee';
    ws.getCell(12, 41).value = 'No. of payable days';
    ws.getCell(12, 42).value = 'Total OT hours';
    const earned = ['BASIC', 'HRA', 'Conveyance', 'Medical Allowance', 'Total'];
    const deductions = [
      'ESI',
      'PF',
      'PT',
      'TDS',
      'Society',
      'Insurance',
      'Salary Advance',
      'Fines',
      'Damages/Loss',
      'Others',
      'Total',
    ];
    earned.forEach((h, i) => {
      ws.getCell(13, 43 + i).value = h;
    });
    deductions.forEach((h, i) => {
      ws.getCell(13, 48 + i).value = h;
    });
    ws.getCell(13, 59).value = 'Net Amount Payable';
    ws.getCell(13, 60).value = 'Mode of Payment Cash/ Cheque No.';
    ws.getCell(13, 61).value = "Employee's signature or Thumb impression";
    ws.getCell(13, 62).value = 'Remarks';

    finalizeFormTSEKarnatakaWorksheetExportBorders(ws, { tableLastRow: 16 });

    expect(String(ws.getCell(12, 43).value ?? '')).toMatch(/Earned wages and other allowances/i);
    expect(String(ws.getCell(12, 48).value ?? '')).toMatch(/^Deductions$/i);
    expect(mergeCoversRow(ws, 12, 43, 47)).toBe(true);
    expect(mergeCoversRow(ws, 12, 48, 58)).toBe(true);
    expect(String(ws.getCell(13, 43).value ?? '')).toMatch(/BASIC/i);
    expect(String(ws.getCell(13, 51).value ?? '')).toMatch(/TDS/i);
    expect(String(ws.getCell(13, 48).value ?? '')).toMatch(/^ESI$/i);
    expect(String(ws.getCell(12, 41).value ?? '')).toMatch(/payable days/i);
    expect(String(ws.getCell(13, 59).value ?? '')).toMatch(/Net Amount Payable/i);
  });

  it('copies row-12 wage leaves under the group banners', () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Form T');
    ws.getCell(1, 1).value = 'FORM T';
    ws.getCell(2, 1).value = 'COMBINED MUSTER ROLL CUM REGISTER OF WAGES';
    ws.getCell(9, 1).value = 'Month / Year';
    ws.getCell(12, 2).value = 'Name of Employee';
    ['BASIC', 'Conveyance', 'Total', 'ESI', 'PF', 'TDS', 'Total', 'Net Amount Payable', 'Remarks'].forEach(
      (h, i) => {
        ws.getCell(12, 43 + i).value = h;
      }
    );

    finalizeFormTSEKarnatakaWorksheetExportBorders(ws, { tableLastRow: 16 });

    expect(String(ws.getCell(12, 43).value ?? '')).toMatch(/Earned wages and other allowances/i);
    expect(String(ws.getCell(12, 46).value ?? '')).toMatch(/^Deductions$/i);
    expect(String(ws.getCell(13, 43).value ?? '')).toMatch(/BASIC/i);
    expect(String(ws.getCell(13, 48).value ?? '')).toMatch(/TDS/i);
    expect(mergeCoversRow(ws, 12, 43, 45)).toBe(true);
    expect(mergeCoversRow(ws, 12, 46, 49)).toBe(true);
  });

  it('merges ATTENDANCE across calendar-day columns', () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Form T');
    ws.getCell(1, 1).value = 'FORM T';
    ws.getCell(2, 1).value = 'COMBINED MUSTER ROLL CUM REGISTER OF WAGES';
    ws.getCell(9, 1).value = 'Month / Year';
    ws.getCell(12, 1).value = 'S.NO';
    ws.getCell(12, 2).value = 'Name of Employee';
    for (let d = 1; d <= 31; d += 1) {
      const c = 9 + d;
      ws.getCell(12, c).value =
        'ATTENDANCE (Please mention the date of suspension of employees, if any)';
      ws.getCell(13, c).value = d;
    }
    ws.getCell(12, 41).value = 'No. of payable days';

    finalizeFormTSEKarnatakaWorksheetExportBorders(ws, { tableLastRow: 16 });

    expect(String(ws.getCell(12, 10).value ?? '')).toMatch(
      /ATTENDANCE \(Please mention the date of suspension of employees, if any\)/i
    );
    expect(mergeCoversRow(ws, 12, 10, 40)).toBe(true);
    expect(Number(ws.getCell(13, 10).value)).toBe(1);
    expect(Number(ws.getCell(13, 40).value)).toBe(31);
    expect(String(ws.getCell(12, 41).value ?? '')).toMatch(/payable days/i);
  });

  it('makes title, headers, and table data bold', () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Form T');
    ws.getCell(1, 1).value = 'FORM T';
    ws.getCell(2, 1).value = 'COMBINED MUSTER ROLL CUM REGISTER OF WAGES';
    ws.getCell(9, 1).value = 'Month / Year';
    ws.getCell(12, 1).value = 'S.NO';
    ws.getCell(12, 2).value = 'Name of Employee';
    ws.getCell(15, 1).value = 1;
    ws.getCell(15, 2).value = 'Issac Kanagaraj';
    ws.getCell(15, 2).font = { bold: false };

    finalizeFormTSEKarnatakaWorksheetExportBorders(ws, { tableLastRow: 16 });

    expect(ws.getCell(1, 1).font?.bold).toBe(true);
    expect(ws.getCell(2, 1).font?.bold).toBe(true);
    expect(ws.getCell(3, 1).font?.bold).toBe(true);
    expect(ws.getCell(12, 2).font?.bold).toBe(true);
    expect(ws.getCell(15, 2).font?.bold).toBe(true);
    expect(ws.getCell(15, 1).font?.bold).toBe(true);
  });
});
