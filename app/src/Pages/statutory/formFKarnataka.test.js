import ExcelJS from 'exceljs';
import {
  FORM_F_KARNATAKA_PART_I_HEADERS,
  FORM_F_KARNATAKA_PART_II_HEADERS,
  FORM_F_KARNATAKA_TEMPLATE_BODY_ROWS,
  alignFormFKarnatakaMappedRowsToEmployees,
  applyFormFKarnatakaApprovedLeaveAutofill,
  applyFormFKarnatakaApprovedLeaveToRow,
  applyFormFKarnatakaEarnedLeaveToRow,
  buildFormFKarnatakaPerEmployeeDownload,
  buildFormFKarnatakaWorkbookWithTemplateStyles,
  chunkFormFKarnatakaRowsForTemplate,
  computeFormFKarnatakaBalanceOnReturn,
  computeFormFKarnatakaLeaveAtCredit,
  ensureFormFKarnatakaTitleLayout,
  filterApprovedLeaveRecordsForFormFKarnataka,
  findApprovedLeaveForEmployee,
  formFKarnatakaPersonNamesMatchStrict,
  groupFormFKarnatakaRowsByLocation,
  isFormFKarnatakaEarnedApprovedLeave,
  isFormFKarnatakaSickAccidentApprovedLeave,
  resolveFormFKarnatakaEmployeeIdentityValues,
  resolveFormFKarnatakaPartIiValues,
  writeFormFKarnatakaIdentityFieldsToWorksheet,
} from './formFKarnataka';

describe('Form F Karnataka PART II Sick/Accident', () => {
  test('isFormFKarnatakaSickAccidentApprovedLeave keeps sick only', () => {
    expect(
      isFormFKarnatakaSickAccidentApprovedLeave({ LeaveType: 'Sick Leave', From: '01-May-2026' })
    ).toBe(true);
    expect(
      isFormFKarnatakaSickAccidentApprovedLeave({
        LeaveType: 'Contingency Leave',
        From: '01-May-2026',
      })
    ).toBe(false);
    expect(
      isFormFKarnatakaSickAccidentApprovedLeave({ LeaveType: 'Earned Leave', From: '01-May-2026' })
    ).toBe(false);
  });

  test('resolveFormFKarnatakaPartIiValues maps sick leave credit / availed / balance', () => {
    const values = resolveFormFKarnatakaPartIiValues({
      leaveRecord: {
        employee: { name: 'A B' },
        'Sick Leave': { paidBalance: '12', paidBooked: '2' },
      },
      monthFrom: '01-May-2026',
    });
    expect(values[FORM_F_KARNATAKA_PART_II_HEADERS[0]]).toBe('2026');
    expect(values[FORM_F_KARNATAKA_PART_II_HEADERS[1]]).toBe('12');
    expect(values[FORM_F_KARNATAKA_PART_II_HEADERS[2]]).toBe('2');
    expect(values[FORM_F_KARNATAKA_PART_II_HEADERS[3]]).toBe('10');
  });
});

describe('Form F Karnataka identity header fields', () => {
  test('resolveFormFKarnatakaEmployeeIdentityValues maps employee master fields', () => {
    const values = resolveFormFKarnatakaEmployeeIdentityValues(
      {
        FirstName: 'Sakthivel',
        LastName: 'Muthunaicker',
        Employee_ID: 'VE0402',
        Father_s_Name: 'Muthunaicker',
        Dateofjoining: '2019-03-15',
      },
      {
        formatStatutoryDateDisplay: (v) =>
          String(v).startsWith('2019') ? '15-Mar-2019' : String(v),
      }
    );
    expect(values.slNo).toBe('VE0402');
    expect(values.personName).toBe('Sakthivel Muthunaicker');
    expect(values.fatherName).toBe('Muthunaicker');
    expect(values.dateOfEntry).toBe('15-Mar-2019');
  });

  test('download fills Sl No / Date of entry / Name / Father Name beside labels', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Form F');
    ws.getCell(1, 1).value = 'FORM F';
    ws.getCell(2, 1).value = '(SEE RULE 8)';
    ws.getCell(4, 1).value = 'REGISTER OF LEAVE WITH WAGES';
    ws.getCell(5, 2).value = 'Sl No in the Register of Adult/young person:';
    ws.getCell(5, 8).value = 'Date of entry into service:';
    ws.getCell(8, 2).value = 'Name of the person :';
    ws.getCell(8, 8).value = "Father's Name :";
    ws.getCell(10, 1).value = 'PART I EARNED LEAVE';
    const headers = [
      'From',
      'To',
      'Total days worked',
      'Leave earned',
      'Leave at credit',
      'From',
      'To',
      'No. of days',
      'Balance on return from leave',
      'Date on which wages for leave paid',
      'Remarks',
    ];
    headers.forEach((h, i) => {
      ws.getCell(11, 1 + i).value = h;
    });
    for (let i = 0; i < 11; i += 1) {
      ws.getCell(12, 1 + i).value = String(i + 1);
    }
    for (let r = 0; r < 3; r += 1) {
      for (let c = 0; c < 11; c += 1) {
        const cell = ws.getCell(13 + r, 1 + c);
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' },
        };
      }
    }
    ws.getCell(17, 1).value = 'PART II - Sick/Accident Leave (with pay)';
    const buf = await wb.xlsx.writeBuffer();

    const { blob } = await buildFormFKarnatakaWorkbookWithTemplateStyles({
      templateArrayBuffer: buf,
      mappedData: [
        {
          [FORM_F_KARNATAKA_PART_I_HEADERS[3]]: '11',
          __employeeLookupName: 'Sakthivel Muthunaicker',
        },
      ],
      headersToUse: FORM_F_KARNATAKA_PART_I_HEADERS,
      sheetNameHint: 'Form F',
      employee: {
        FirstName: 'Sakthivel',
        LastName: 'Muthunaicker',
        Employee_ID: 'VE0402',
        Father_s_Name: 'Ramasamy',
        Dateofjoining: '2019-03-15',
      },
      formatStatutoryDateDisplay: () => '15-Mar-2019',
    });

    const outWb = new ExcelJS.Workbook();
    const outBytes =
      typeof blob?.arrayBuffer === 'function'
        ? await blob.arrayBuffer()
        : await new Response(blob).arrayBuffer();
    await outWb.xlsx.load(outBytes);
    const outWs = outWb.worksheets[0];
    expect(String(outWs.getCell(5, 3).value || '')).toBe('VE0402');
    expect(String(outWs.getCell(5, 9).value || '')).toBe('15-Mar-2019');
    expect(String(outWs.getCell(8, 3).value || '')).toBe('Sakthivel Muthunaicker');
    expect(String(outWs.getCell(8, 9).value || '')).toBe('Ramasamy');
  });

  test('writeFormFKarnatakaIdentityFieldsToWorksheet does not confuse Name with Father Name', () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Form F');
    ws.getCell(3, 2).value = 'Name of the person :';
    ws.getCell(3, 8).value = "Father's Name :";
    const written = writeFormFKarnatakaIdentityFieldsToWorksheet(ws, {
      personName: 'A B',
      fatherName: 'C D',
      slNo: '',
      dateOfEntry: '',
    });
    expect(written).toBe(2);
    expect(String(ws.getCell(3, 3).value || '')).toBe('A B');
    expect(String(ws.getCell(3, 9).value || '')).toBe('C D');
  });
});

describe('Form F Karnataka leave mapping', () => {
  const leaveRecord = {
    employee: { name: 'Test Emp', id: 'E1' },
    'Earned Leave': { paidBalance: '23', paidBooked: '6' },
  };

  test('Leave earned ← Leave earned during the Period (paidBalance)', () => {
    const row = {};
    applyFormFKarnatakaEarnedLeaveToRow(row, leaveRecord, {}, null);
    expect(row[FORM_F_KARNATAKA_PART_I_HEADERS[3]]).toBe('23');
  });

  test('Leave at credit ← Leave earned during the Period', () => {
    const row = {};
    applyFormFKarnatakaEarnedLeaveToRow(row, leaveRecord, {}, null);
    expect(row[FORM_F_KARNATAKA_PART_I_HEADERS[4]]).toBe('23');
  });

  test('Balance on return ← Balance count when no leave availed', () => {
    const row = {};
    applyFormFKarnatakaEarnedLeaveToRow(row, leaveRecord, {}, null);
    expect(row[FORM_F_KARNATAKA_PART_I_HEADERS[8]]).toBe('23');
    expect(computeFormFKarnatakaBalanceOnReturn('23', '0')).toBe('23');
  });

  test('Balance on return ← Balance count after approved LeaveCount', () => {
    const row = {};
    const approved = {
      Days: {
        '04-May-2026': { LeaveCount: 3 },
        '05-May-2026': { LeaveCount: 3 },
      },
      From: '04-May-2026',
      To: '09-May-2026',
    };
    applyFormFKarnatakaEarnedLeaveToRow(row, leaveRecord, {}, approved);
    expect(row[FORM_F_KARNATAKA_PART_I_HEADERS[3]]).toBe('23');
    expect(row[FORM_F_KARNATAKA_PART_I_HEADERS[4]]).toBe('23');
    expect(row[FORM_F_KARNATAKA_PART_I_HEADERS[8]]).toBe('17');
    expect(computeFormFKarnatakaLeaveAtCredit('23', '6')).toBe('17');
  });

  test('approved leave updates Balance on return, not Leave at credit', () => {
    const row = {
      [FORM_F_KARNATAKA_PART_I_HEADERS[3]]: '23',
      [FORM_F_KARNATAKA_PART_I_HEADERS[4]]: '23',
    };
    const approved = {
      Days: {
        '04-May-2026': { LeaveCount: 6 },
      },
      From: '04-May-2026',
      To: '09-May-2026',
      ApprovalStatus: 'APPROVED',
      LeaveType: 'Earned Leave',
    };
    applyFormFKarnatakaApprovedLeaveToRow(
      row,
      approved,
      FORM_F_KARNATAKA_PART_I_HEADERS,
      FORM_F_KARNATAKA_PART_I_HEADERS,
      { overwrite: true, monthFrom: '01-May-2026', monthTo: '31-May-2026' }
    );
    expect(row[FORM_F_KARNATAKA_PART_I_HEADERS[4]]).toBe('23');
    expect(row[FORM_F_KARNATAKA_PART_I_HEADERS[7]]).toBe('6');
    expect(row[FORM_F_KARNATAKA_PART_I_HEADERS[8]]).toBe('17');
  });

  test('ensureFormFKarnatakaTitleLayout centers FORM F / SEE RULE 8 / REGISTER across A–K', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Form F');
    ws.getCell(1, 12).value = 'FORM F';
    ws.getCell(2, 12).value = '(SEE RULE 8)';
    ws.getCell(4, 12).value = 'REGISTER OF LEAVE WITH WAGES';
    ws.getCell(11, 12).value = 'PART I EARNED LEAVE';
    ws.getCell(13, 1).value = 'From';
    ws.getCell(13, 4).value = 'Leave earned';

    const ok = ensureFormFKarnatakaTitleLayout(ws);
    expect(ok).toBe(true);
    expect(String(ws.getCell(1, 1).value || '')).toBe('FORM F');
    expect(String(ws.getCell(2, 1).value || '')).toBe('(SEE RULE 8)');
    expect(String(ws.getCell(4, 1).value || '')).toBe('REGISTER OF LEAVE WITH WAGES');
    expect(ws.getCell(1, 1).alignment?.horizontal).toBe('center');
    expect(String(ws.getCell(1, 12).value || '')).toBe('');
  });

  test('ensureFormFKarnatakaTitleLayout does not rewrite Form 15 Part II titles', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Form 15 Part II');
    ws.getCell(1, 1).value = 'FORM 15';
    ws.getCell(2, 1).value = 'PART II';
    ws.getCell(3, 1).value = 'REGISTER OF LEAVE WITH WAGES';
    ws.getCell(4, 1).value = '(SEE RULE 95)';
    ws.getCell(11, 3).value = 'Name of the Worker';
    ws.getCell(11, 7).value = 'Basic Wages';
    ws.getCell(11, 13).value = 'Gross Wages';

    const ok = ensureFormFKarnatakaTitleLayout(ws);
    expect(ok).toBe(false);
    expect(String(ws.getCell(1, 1).value || '')).toBe('FORM 15');
    expect(String(ws.getCell(2, 1).value || '')).toBe('PART II');
    expect(String(ws.getCell(3, 1).value || '')).toBe('REGISTER OF LEAVE WITH WAGES');
    expect(String(ws.getCell(4, 1).value || '')).toBe('(SEE RULE 95)');
  });
});

describe('Form F Karnataka leave availed identity', () => {
  test('strict name match rejects first-name-only false positives', () => {
    expect(formFKarnatakaPersonNamesMatchStrict('Sakthivel Muthunaicker', 'Sakthivel Kumar')).toBe(
      false
    );
    expect(formFKarnatakaPersonNamesMatchStrict('Sakthivel', 'Sakthivel Muthunaicker')).toBe(false);
    expect(
      formFKarnatakaPersonNamesMatchStrict('Sakthivel Muthunaicker', 'Sakthivel Muthunaicker')
    ).toBe(true);
  });

  test('Contingency leave is excluded from Form F PART I', () => {
    expect(
      isFormFKarnatakaEarnedApprovedLeave({ LeaveType: 'Contingency Leave', From: '04-May-2026' })
    ).toBe(false);
    expect(
      isFormFKarnatakaEarnedApprovedLeave({ LeaveType: 'Earned Leave', From: '04-May-2026' })
    ).toBe(true);
  });

  test('Booked: 0 employee does not get Leave availed dates from another person', () => {
    const mapped = [
      {
        [FORM_F_KARNATAKA_PART_I_HEADERS[3]]: '33',
        [FORM_F_KARNATAKA_PART_I_HEADERS[4]]: '33',
        'Name of the person': 'Sakthivel Muthunaicker',
      },
    ];
    const employees = [
      {
        FirstName: 'Sakthivel',
        LastName: 'Muthunaicker',
        Employee_ID: 'VE0402',
      },
    ];
    const leaveRecords = [
      {
        employee: { name: 'Sakthivel Muthunaicker (VE0402)', id: 'VE0402' },
        EmployeeName: 'Sakthivel Muthunaicker (VE0402)',
        'Earned Leave': { paidBalance: '33', paidBooked: '0' },
      },
    ];
    const approved = [
      {
        Employee: 'Sakthivel SomeoneElse',
        LeaveType: 'Earned Leave',
        ApprovalStatus: 'APPROVED',
        From: '04-May-2026',
        To: '09-May-2026',
        Days: { '04-May-2026': { LeaveCount: 6 } },
      },
    ];
    const hits = applyFormFKarnatakaApprovedLeaveAutofill(
      mapped,
      employees,
      FORM_F_KARNATAKA_PART_I_HEADERS,
      approved,
      {
        overwrite: true,
        monthFrom: '01-May-2026',
        monthTo: '31-May-2026',
        leaveRecords,
        leaveTypeLabels: {},
      }
    );
    expect(hits).toBe(0);
    expect(mapped[0][FORM_F_KARNATAKA_PART_I_HEADERS[5]] || '').toBe('');
    expect(mapped[0][FORM_F_KARNATAKA_PART_I_HEADERS[6]] || '').toBe('');
    expect(mapped[0][FORM_F_KARNATAKA_PART_I_HEADERS[7]] || '').toBe('');
  });

  test('findApprovedLeaveForEmployee does not match on first name alone when strict', () => {
    const approved = [
      {
        Employee: 'Sakthivel Kumar',
        LeaveType: 'Earned Leave',
        ApprovalStatus: 'APPROVED',
        From: '04-May-2026',
        To: '09-May-2026',
        Days: { '04-May-2026': { LeaveCount: 6 } },
      },
    ];
    const hit = findApprovedLeaveForEmployee(
      null,
      { FirstName: 'Sakthivel', LastName: 'Muthunaicker', Employee_ID: 'VE0402' },
      { 'Name of the person': 'Sakthivel Muthunaicker' },
      '',
      'Name of the person',
      {
        approvedRecords: approved,
        monthFrom: '01-May-2026',
        monthTo: '31-May-2026',
        strictIdentity: true,
      }
    );
    expect(hit).toBeNull();
  });

  test('filterApprovedLeaveRecordsForFormFKarnataka keeps earned only', () => {
    const filtered = filterApprovedLeaveRecordsForFormFKarnataka(
      [
        {
          LeaveType: 'Earned Leave',
          ApprovalStatus: 'APPROVED',
          From: '04-May-2026',
          To: '09-May-2026',
          Days: { '04-May-2026': { LeaveCount: 1 } },
        },
        {
          LeaveType: 'Contingency Leave',
          ApprovalStatus: 'APPROVED',
          From: '02-May-2026',
          To: '02-May-2026',
          Days: { '02-May-2026': { LeaveCount: 1 } },
        },
      ],
      '01-May-2026',
      '31-May-2026'
    );
    expect(filtered).toHaveLength(1);
    expect(filtered[0].LeaveType).toBe('Earned Leave');
  });
});

describe('Form F Karnataka location ZIP helpers', () => {
  test('chunkFormFKarnatakaRowsForTemplate splits into 15-row batches', () => {
    const rows = Array.from({ length: 37 }, (_, i) => ({ i }));
    const chunks = chunkFormFKarnatakaRowsForTemplate(rows);
    expect(FORM_F_KARNATAKA_TEMPLATE_BODY_ROWS).toBe(15);
    expect(chunks).toHaveLength(3);
    expect(chunks[0]).toHaveLength(15);
    expect(chunks[1]).toHaveLength(15);
    expect(chunks[2]).toHaveLength(7);
  });

  test('groupFormFKarnatakaRowsByLocation folders by Work_location', () => {
    const mapped = [
      { 'Name of the person': 'A B', [FORM_F_KARNATAKA_PART_I_HEADERS[3]]: '1' },
      { 'Name of the person': 'C D', [FORM_F_KARNATAKA_PART_I_HEADERS[3]]: '2' },
      { 'Name of the person': 'E F', [FORM_F_KARNATAKA_PART_I_HEADERS[3]]: '3' },
    ];
    const employees = [
      { FirstName: 'A', LastName: 'B', Work_location: 'Theni' },
      { FirstName: 'C', LastName: 'D', Work_location: 'Madurai' },
      { FirstName: 'E', LastName: 'F', Work_location: 'Theni' },
    ];
    const groups = groupFormFKarnatakaRowsByLocation(mapped, employees);
    expect(groups).toHaveLength(2);
    const theni = groups.find((g) => /theni/i.test(g.location));
    const madurai = groups.find((g) => /madurai/i.test(g.location));
    expect(theni.rows).toHaveLength(2);
    expect(madurai.rows).toHaveLength(1);
  });

  test('alignFormFKarnatakaMappedRowsToEmployees filters when site list is smaller', () => {
    const mapped = [
      { __employeeLookupName: 'A B', [FORM_F_KARNATAKA_PART_I_HEADERS[3]]: '1' },
      { __employeeLookupName: 'C D', [FORM_F_KARNATAKA_PART_I_HEADERS[3]]: '2' },
      { __employeeLookupName: 'E F', [FORM_F_KARNATAKA_PART_I_HEADERS[3]]: '3' },
    ];
    const siteEmployees = [
      { FirstName: 'C', LastName: 'D', Work_location: 'Madurai' },
      { FirstName: 'E', LastName: 'F', Work_location: 'Theni' },
    ];
    const aligned = alignFormFKarnatakaMappedRowsToEmployees(mapped, siteEmployees);
    expect(aligned).toHaveLength(2);
    expect(aligned[0].__employeeLookupName).toMatch(/c d/i);
    expect(aligned[1].__employeeLocation).toMatch(/theni/i);
  });
});

describe('Form F Karnataka per-employee download', () => {
  async function makeFormFTemplateBuffer() {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Form F');
    ws.getCell(1, 1).value = 'FORM F';
    ws.getCell(2, 1).value = '(SEE RULE 8)';
    ws.getCell(4, 1).value = 'REGISTER OF LEAVE WITH WAGES';
    ws.getCell(7, 1).value = 'PART I EARNED LEAVE';
    const headers = [
      'From',
      'To',
      'Total days worked',
      'Leave earned',
      'Leave at credit',
      'From',
      'To',
      'No. of days',
      'Balance on return from leave',
      'Date on which wages for leave paid',
      'Remarks',
    ];
    headers.forEach((h, i) => {
      ws.getCell(9, 1 + i).value = h;
    });
    for (let i = 0; i < 11; i += 1) {
      ws.getCell(10, 1 + i).value = String(i + 1);
    }
    for (let r = 0; r < 15; r += 1) {
      for (let c = 0; c < 11; c += 1) {
        const cell = ws.getCell(11 + r, 1 + c);
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' },
        };
      }
    }
    ws.getCell(27, 1).value = 'PART II - Sick/Accident Leave';
    return wb.xlsx.writeBuffer();
  }

  test('buildFormFKarnatakaPerEmployeeDownload emits one xlsx per employee in a ZIP', async () => {
    const buf = await makeFormFTemplateBuffer();
    const mapped = [
      {
        __employeeLookupName: 'A B',
        [FORM_F_KARNATAKA_PART_I_HEADERS[3]]: '11',
        [FORM_F_KARNATAKA_PART_I_HEADERS[4]]: '11',
      },
      {
        __employeeLookupName: 'C D',
        [FORM_F_KARNATAKA_PART_I_HEADERS[3]]: '33',
        [FORM_F_KARNATAKA_PART_I_HEADERS[4]]: '33',
      },
      {
        __employeeLookupName: 'E F',
        [FORM_F_KARNATAKA_PART_I_HEADERS[3]]: '38',
        [FORM_F_KARNATAKA_PART_I_HEADERS[4]]: '38',
      },
    ];
    const employees = [
      { FirstName: 'A', LastName: 'B', Employee_ID: 'E1' },
      { FirstName: 'C', LastName: 'D', Employee_ID: 'E2' },
      { FirstName: 'E', LastName: 'F', Employee_ID: 'E3' },
    ];
    const { blob, fileName } = await buildFormFKarnatakaPerEmployeeDownload({
      templateArrayBuffer: buf,
      mappedData: mapped,
      headersToUse: FORM_F_KARNATAKA_PART_I_HEADERS,
      formFileName: 'Form_F_-_Karnataka.xlsx',
      employeesOverride: employees,
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
    expect(names.some((n) => /Form_F_Karnataka_A_B\.xlsx$/i.test(n))).toBe(true);
    expect(names.some((n) => /Form_F_Karnataka_C_D\.xlsx$/i.test(n))).toBe(true);
    expect(names.some((n) => /Form_F_Karnataka_E_F\.xlsx$/i.test(n))).toBe(true);

    const firstBytes = await zip.files[names.find((n) => /A_B/i.test(n))].async('arraybuffer');
    const outWb = new ExcelJS.Workbook();
    await outWb.xlsx.load(firstBytes);
    const outWs = outWb.worksheets[0];
    expect(String(outWs.getCell(11, 4).value)).toBe('11');
    // Only that employee's row — second PART I body row must stay empty.
    expect(outWs.getCell(12, 4).value == null || String(outWs.getCell(12, 4).value).trim() === '').toBe(
      true
    );
  });

  test('buildFormFKarnatakaWorkbookWithTemplateStyles writes into 15 PART I body rows', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Form F');
    ws.getCell(1, 2).value = 'REGISTER OF LEAVE WITH WAGES';
    ws.getCell(7, 2).value = 'PART I EARNED LEAVE';
    const headers = [
      'From',
      'To',
      'Total days worked',
      'Leave earned',
      'Leave at credit',
      'From',
      'To',
      'No. of days',
      'Balance on return from leave',
      'Date on which wages for leave paid',
      'Remarks',
    ];
    headers.forEach((h, i) => {
      ws.getCell(9, 2 + i).value = h;
    });
    for (let i = 0; i < 11; i += 1) {
      ws.getCell(10, 2 + i).value = String(i + 1);
    }
    for (let r = 0; r < 15; r += 1) {
      for (let c = 0; c < 11; c += 1) {
        const cell = ws.getCell(11 + r, 2 + c);
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' },
        };
      }
    }
    ws.getCell(27, 2).value = 'PART II - Sick/Accident Leave';
    const buf = await wb.xlsx.writeBuffer();

    const mapped = [
      {
        [FORM_F_KARNATAKA_PART_I_HEADERS[0]]: '01-May-2026',
        [FORM_F_KARNATAKA_PART_I_HEADERS[1]]: '31-May-2026',
        [FORM_F_KARNATAKA_PART_I_HEADERS[2]]: '31',
        [FORM_F_KARNATAKA_PART_I_HEADERS[3]]: '33',
        [FORM_F_KARNATAKA_PART_I_HEADERS[4]]: '33',
        [FORM_F_KARNATAKA_PART_I_HEADERS[8]]: '33',
      },
    ];
    const { blob, layout } = await buildFormFKarnatakaWorkbookWithTemplateStyles({
      templateArrayBuffer: buf,
      mappedData: mapped,
      headersToUse: FORM_F_KARNATAKA_PART_I_HEADERS,
      sheetNameHint: 'Form F',
    });
    expect(layout.dataStartRow).toBe(11);
    expect(layout.bodyRows).toBe(15);
    expect(layout.colFrom).toBe(2);

    const outWb = new ExcelJS.Workbook();
    const outBytes =
      typeof blob?.arrayBuffer === 'function'
        ? await blob.arrayBuffer()
        : await new Response(blob).arrayBuffer();
    await outWb.xlsx.load(outBytes);
    const outWs = outWb.worksheets[0];
    expect(String(outWs.getCell(11, 2).value)).toBe('01-May-2026');
    expect(String(outWs.getCell(11, 3).value)).toBe('31-May-2026');
    expect(String(outWs.getCell(11, 5).value)).toBe('33');
    // Must not spill below the 15-row body (row 26 is last body; 27 is PART II).
    expect(outWs.getCell(26, 2).value == null || String(outWs.getCell(26, 2).value).trim() === '').toBe(
      true
    );
    expect(String(outWs.getCell(27, 2).value || '')).toMatch(/PART II/i);
  });

  test('download preserves PART II Sick/Accident Leave when template PART I is short', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Form F');
    ws.getCell(1, 1).value = 'FORM F';
    ws.getCell(2, 1).value = '(SEE RULE 8)';
    ws.getCell(4, 1).value = 'REGISTER OF LEAVE WITH WAGES';
    ws.getCell(6, 1).value = 'PART I EARNED LEAVE';
    const headers = [
      'From',
      'To',
      'Total days worked',
      'Leave earned',
      'Leave at credit',
      'From',
      'To',
      'No. of days',
      'Balance on return from leave',
      'Date on which wages for leave paid',
      'Remarks',
    ];
    headers.forEach((h, i) => {
      ws.getCell(7, 1 + i).value = h;
    });
    for (let i = 0; i < 11; i += 1) {
      ws.getCell(8, 1 + i).value = String(i + 1);
    }
    // Only 7 PART I body rows before PART II (matches compact Form F template).
    for (let r = 0; r < 7; r += 1) {
      for (let c = 0; c < 11; c += 1) {
        const cell = ws.getCell(9 + r, 1 + c);
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' },
        };
      }
    }
    ws.getCell(16, 1).value = 'PART II - Sick/Accident Leave (with pay)';
    ws.getCell(17, 1).value = 'Year';
    ws.getCell(17, 2).value = 'of credit';
    ws.getCell(17, 3).value = 'Availed';
    ws.getCell(17, 4).value = 'Balance at the end of the year';
    ws.getCell(18, 4).value = '-';
    const buf = await wb.xlsx.writeBuffer();

    const { blob, layout } = await buildFormFKarnatakaWorkbookWithTemplateStyles({
      templateArrayBuffer: buf,
      mappedData: [
        {
          [FORM_F_KARNATAKA_PART_I_HEADERS[3]]: '11',
          [FORM_F_KARNATAKA_PART_I_HEADERS[4]]: '11',
        },
      ],
      headersToUse: FORM_F_KARNATAKA_PART_I_HEADERS,
      sheetNameHint: 'Form F',
    });
    expect(layout.partIiRow).toBe(16);
    expect(layout.bodyRows).toBeLessThanOrEqual(7);
    expect(layout.dataStartRow + layout.bodyRows).toBeLessThanOrEqual(16);

    const outWb = new ExcelJS.Workbook();
    const outBytes =
      typeof blob?.arrayBuffer === 'function'
        ? await blob.arrayBuffer()
        : await new Response(blob).arrayBuffer();
    await outWb.xlsx.load(outBytes);
    const outWs = outWb.worksheets[0];
    expect(String(outWs.getCell(16, 1).value || '')).toMatch(/PART II.*Sick/i);
    expect(String(outWs.getCell(17, 1).value || '')).toMatch(/Year/i);
    expect(String(outWs.getCell(17, 2).value || '')).toMatch(/credit/i);
    expect(String(outWs.getCell(17, 3).value || '')).toMatch(/Availed/i);
    expect(String(outWs.getCell(17, 4).value || '')).toMatch(/Balance/i);
    expect(String(outWs.getCell(18, 4).value || '')).toBe('-');
  });
});
