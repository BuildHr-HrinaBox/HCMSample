import {
  FORM_NGJ_CASUAL_HEADERS,
  FORM_NGJ_CASUAL_PARENT,
  FORM_NGJ_FESTIVAL_PARENT,
  applyFormNGJGujaratCasualLeaveAutofill,
  applyFormNGJGujaratCasualLeaveToRow,
  applyFormNGJGujaratEmployeeToRow,
  buildFormNGJCasualLeaveValues,
  buildFormNGJGujaratWorkbookWithTemplateStyles,
  filterApprovedLeaveRecordsForFormNGJCasual,
  filterFormNGJGujaratExportRows,
  getFormNGJRowValueForHeader,
  isApprovedLeaveContingencyForFormNGJ,
  isFormNGJCasualAvailedLeaveHeader,
  isFormNGJCasualBalanceLeaveHeader,
  isFormNGJCasualPeriodFromHeader,
  isFormNGJCasualPeriodToHeader,
  isFormNGJCasualTotalLeaveHeader,
  remapFormNGJGujaratRowsToHeaders,
  resolveFormNGJGujaratHeaderFieldLayout,
  resolveFormNGJGujaratTableHeaders,
  rowHasMeaningfulFormNGJGujaratExportData,
} from './formNGJGujarat';

describe('Form N GJ Casual Leave (Contingency)', () => {
  const headers = resolveFormNGJGujaratTableHeaders();

  it('recognizes Casual Leave From / To / Total / Availed / Balance headers', () => {
    expect(isFormNGJCasualPeriodFromHeader(`${FORM_NGJ_CASUAL_PARENT}_Period_From`)).toBe(true);
    expect(isFormNGJCasualPeriodToHeader(`${FORM_NGJ_CASUAL_PARENT}_Period_To`)).toBe(true);
    expect(isFormNGJCasualTotalLeaveHeader(`${FORM_NGJ_CASUAL_PARENT}_Total Leave`)).toBe(true);
    expect(isFormNGJCasualAvailedLeaveHeader(`${FORM_NGJ_CASUAL_PARENT}_Availed Leave`)).toBe(true);
    expect(isFormNGJCasualBalanceLeaveHeader(`${FORM_NGJ_CASUAL_PARENT}_Balance Leave`)).toBe(true);
    expect(isFormNGJCasualPeriodFromHeader('Details of Festival Leave_Period_From')).toBe(false);
  });

  it('keeps Period_From / Period_To intact when resolving headers', () => {
    const resolved = resolveFormNGJGujaratTableHeaders([...headers]);
    expect(resolved).toContain(`${FORM_NGJ_CASUAL_PARENT}_Period_From`);
    expect(resolved).toContain(`${FORM_NGJ_CASUAL_PARENT}_Period_To`);
    expect(resolved).not.toContain(`${FORM_NGJ_CASUAL_PARENT}_From`);
  });

  it('filters Contingency Leave approved records only', () => {
    const records = [
      {
        'Leave Type': 'Contingency Leave',
        From: '02-Apr-2026',
        To: '03-Apr-2026',
        Days: {
          '02-Apr-2026': { LeaveCount: 1 },
          '03-Apr-2026': { LeaveCount: 1 },
        },
        ApprovalStatus: 'APPROVED',
        'Employee Name': 'Ravi Kumar',
      },
      {
        'Leave Type': 'Earned Leave',
        From: '05-Apr-2026',
        To: '06-Apr-2026',
        Days: {
          '05-Apr-2026': { LeaveCount: 1 },
          '06-Apr-2026': { LeaveCount: 1 },
        },
        ApprovalStatus: 'APPROVED',
        'Employee Name': 'Ravi Kumar',
      },
    ];
    expect(isApprovedLeaveContingencyForFormNGJ(records[0])).toBe(true);
    expect(isApprovedLeaveContingencyForFormNGJ(records[1])).toBe(false);
    const filtered = filterApprovedLeaveRecordsForFormNGJCasual(
      records,
      '01-Apr-2026',
      '30-Apr-2026'
    );
    expect(filtered).toHaveLength(1);
    expect(filtered[0]['Leave Type']).toBe('Contingency Leave');
  });

  it('builds From/To/Availed from approved leave; Balance from Contingency LeaveData', () => {
    const approved = {
      'Leave Type': 'Contingency Leave',
      From: '10-Apr-2026',
      To: '11-Apr-2026',
      Days: {
        '10-Apr-2026': { LeaveCount: 1 },
        '11-Apr-2026': { LeaveCount: 1 },
      },
    };
    const leaveRecord = {
      'Contingency Leave': { balance: 5, booked: 2 },
    };
    const values = buildFormNGJCasualLeaveValues(approved, leaveRecord);
    expect(values.from).toBe('10-Apr-2026');
    expect(values.to).toBe('11-Apr-2026');
    expect(values.availedLeave).toBe('2');
    expect(values.balanceLeave).toBe('5');
    expect(values.totalLeave).toBe('7');
  });

  it('uses approved LeaveCount for Availed, not LeaveData YTD booked', () => {
    const approved = {
      'Leave Type': 'Contingency Leave',
      From: '30-Apr-2026',
      To: '02-May-2026',
      Days: {
        '30-Apr-2026': { LeaveCount: 1 },
        '01-May-2026': { LeaveCount: 1 },
        '02-May-2026': { LeaveCount: 1 },
      },
    };
    const leaveRecord = {
      'Contingency Leave': { balance: 5, booked: 8 },
    };
    const values = buildFormNGJCasualLeaveValues(approved, leaveRecord);
    expect(values.from).toBe('30-Apr-2026');
    expect(values.to).toBe('02-May-2026');
    expect(values.availedLeave).toBe('3');
    expect(values.balanceLeave).toBe('5');
    expect(values.totalLeave).toBe('8');
  });

  it('returns empty when there is no approved Contingency leave (no LeaveData defaults)', () => {
    const leaveRecord = {
      'Contingency Leave': { balance: 8, booked: 8 },
    };
    const values = buildFormNGJCasualLeaveValues(null, leaveRecord);
    expect(values.from).toBe('');
    expect(values.to).toBe('');
    expect(values.availedLeave).toBe('');
    expect(values.balanceLeave).toBe('');
    expect(values.totalLeave).toBe('');
    expect(values.hasApprovedPeriod).toBe(false);
    expect(values.hasLeaveMetrics).toBe(false);
  });

  it('uses approved LeaveCount for Availed/Total when LeaveData missing', () => {
    const approved = {
      'Leave Type': 'CL',
      From: '01-May-2026',
      To: '01-May-2026',
      Days: { '01-May-2026': { LeaveCount: 1 } },
    };
    const values = buildFormNGJCasualLeaveValues(approved, null);
    expect(values.from).toBe('01-May-2026');
    expect(values.to).toBe('01-May-2026');
    expect(values.availedLeave).toBe('1');
    expect(values.totalLeave).toBe('1');
    expect(values.balanceLeave).toBe('');
  });

  it('applies Casual Leave cells onto a Form N row', () => {
    const row = {
      'Name of the worker': 'Anita Shah',
    };
    FORM_NGJ_CASUAL_HEADERS.forEach((h) => {
      row[h] = '';
    });
    const applied = applyFormNGJGujaratCasualLeaveToRow(
      row,
      {
        'Leave Type': 'Contingency Leave',
        From: '12-Jun-2026',
        To: '13-Jun-2026',
        Days: {
          '12-Jun-2026': { LeaveCount: 1 },
          '13-Jun-2026': { LeaveCount: 1 },
        },
      },
      { 'Contingency Leave': { unpaidBalance: 4, unpaidBooked: 2 } },
      headers
    );
    expect(applied).toBeGreaterThan(0);
    expect(row[`${FORM_NGJ_CASUAL_PARENT}_Period_From`]).toBe('12-Jun-2026');
    expect(row[`${FORM_NGJ_CASUAL_PARENT}_Period_To`]).toBe('13-Jun-2026');
    expect(row[`${FORM_NGJ_CASUAL_PARENT}_Availed Leave`]).toBe('2');
    expect(row[`${FORM_NGJ_CASUAL_PARENT}_Balance Leave`]).toBe('4');
    expect(row[`${FORM_NGJ_CASUAL_PARENT}_Total Leave`]).toBe('6');
  });

  it('autofills Casual Leave via Form O-style Contingency matching', () => {
    const mappedData = [
      {
        'Name of the worker': 'Chetan Kumar',
        [`${FORM_NGJ_CASUAL_PARENT}_Period_From`]: '',
        [`${FORM_NGJ_CASUAL_PARENT}_Period_To`]: '',
        [`${FORM_NGJ_CASUAL_PARENT}_Total Leave`]: '',
        [`${FORM_NGJ_CASUAL_PARENT}_Availed Leave`]: '',
        [`${FORM_NGJ_CASUAL_PARENT}_Balance Leave`]: '',
      },
    ];
    const employees = [{ FirstName: 'Chetan', LastName: 'Kumar', Zoho_ID: '5001001' }];
    const approved = [
      {
        'Leave Type': 'Contingency Leave',
        From: '04-Jul-2026',
        To: '05-Jul-2026',
        Days: {
          '04-Jul-2026': { LeaveCount: 1 },
          '05-Jul-2026': { LeaveCount: 1 },
        },
        ApprovalStatus: 'APPROVED',
        'Employee Name': 'Chetan Kumar',
        ZohoID: '5001001',
      },
    ];
    const leaveRecords = [
      {
        employee: { name: 'Chetan Kumar', id: '5001001' },
        'Contingency Leave': { balance: 8, booked: 2 },
      },
    ];
    const hits = applyFormNGJGujaratCasualLeaveAutofill(
      mappedData,
      employees,
      headers,
      approved,
      leaveRecords,
      { monthFrom: '01-Jul-2026', monthTo: '31-Jul-2026' }
    );
    expect(hits).toBe(1);
    expect(mappedData[0][`${FORM_NGJ_CASUAL_PARENT}_Period_From`]).toBe('04-Jul-2026');
    expect(mappedData[0][`${FORM_NGJ_CASUAL_PARENT}_Period_To`]).toBe('05-Jul-2026');
    expect(mappedData[0][`${FORM_NGJ_CASUAL_PARENT}_Total Leave`]).toBe('10');
    expect(mappedData[0][`${FORM_NGJ_CASUAL_PARENT}_Availed Leave`]).toBe('2');
    expect(mappedData[0][`${FORM_NGJ_CASUAL_PARENT}_Balance Leave`]).toBe('8');
  });

  it('clears Casual Leave columns when worker has no Contingency leave', () => {
    const mappedData = [
      {
        'Name of the worker': 'No Leave Worker',
        [`${FORM_NGJ_CASUAL_PARENT}_Period_From`]: '30-Apr-2026',
        [`${FORM_NGJ_CASUAL_PARENT}_Period_To`]: '02-May-2026',
        [`${FORM_NGJ_CASUAL_PARENT}_Total Leave`]: '8',
        [`${FORM_NGJ_CASUAL_PARENT}_Availed Leave`]: '8',
        [`${FORM_NGJ_CASUAL_PARENT}_Balance Leave`]: '5',
      },
    ];
    const leaveRecords = [
      {
        employee: { name: 'No Leave Worker', id: '9001' },
        'Contingency Leave': { balance: 8, booked: 8 },
      },
    ];
    const hits = applyFormNGJGujaratCasualLeaveAutofill(
      mappedData,
      [{ FirstName: 'No', LastName: 'Leave Worker', Zoho_ID: '9001' }],
      headers,
      [],
      leaveRecords,
      { monthFrom: '01-May-2026', monthTo: '31-May-2026' }
    );
    expect(hits).toBe(0);
    expect(mappedData[0][`${FORM_NGJ_CASUAL_PARENT}_Period_From`]).toBe('');
    expect(mappedData[0][`${FORM_NGJ_CASUAL_PARENT}_Period_To`]).toBe('');
    expect(mappedData[0][`${FORM_NGJ_CASUAL_PARENT}_Total Leave`]).toBe('');
    expect(mappedData[0][`${FORM_NGJ_CASUAL_PARENT}_Availed Leave`]).toBe('');
    expect(mappedData[0][`${FORM_NGJ_CASUAL_PARENT}_Balance Leave`]).toBe('');
  });

  it('matches Contingency LeaveData by EmployeeID (VE code), not a similarly named worker', () => {
    const mappedData = [
      {
        'Name of the worker': 'Ajaykumar Mansingbhai',
        [`${FORM_NGJ_CASUAL_PARENT}_Period_From`]: '',
        [`${FORM_NGJ_CASUAL_PARENT}_Period_To`]: '',
        [`${FORM_NGJ_CASUAL_PARENT}_Total Leave`]: '',
        [`${FORM_NGJ_CASUAL_PARENT}_Availed Leave`]: '',
        [`${FORM_NGJ_CASUAL_PARENT}_Balance Leave`]: '',
      },
    ];
    const employees = [
      {
        FirstName: 'Ajaykumar',
        LastName: 'Mansingbhai',
        EmployeeID: 'VE0471',
        Zoho_ID: '9990001',
      },
    ];
    const approved = [
      {
        'Leave Type': 'Contingency Leave',
        From: '30-Apr-2026',
        To: '02-May-2026',
        Days: {
          '30-Apr-2026': { LeaveCount: 1 },
          '01-May-2026': { LeaveCount: 1 },
          '02-May-2026': { LeaveCount: 1 },
        },
        ApprovalStatus: 'APPROVED',
        'Employee Name': 'Ajaykumar Mansingbhai',
        ZohoID: '9990001',
      },
    ];
    const leaveRecords = [
      {
        employee: { name: 'Vijaya Sankar Chandrasekar', id: 'VE0465' },
        'Contingency Leave': { balance: 21, booked: 3 },
      },
      {
        employee: { name: 'Ajaykumar Mansingbhai', id: 'VE0471' },
        'Contingency Leave': { balance: 25, booked: 3 },
      },
    ];
    applyFormNGJGujaratCasualLeaveAutofill(
      mappedData,
      employees,
      headers,
      approved,
      leaveRecords,
      { monthFrom: '01-Apr-2026', monthTo: '30-Apr-2026' }
    );
    expect(mappedData[0][`${FORM_NGJ_CASUAL_PARENT}_Availed Leave`]).toBe('3');
    expect(mappedData[0][`${FORM_NGJ_CASUAL_PARENT}_Balance Leave`]).toBe('25');
    expect(mappedData[0][`${FORM_NGJ_CASUAL_PARENT}_Total Leave`]).toBe('28');
  });

  it("does not take another Ajaykumar's 5-day leave when FirstName+LastName is Mansingbhai (VE0471)", () => {
    const mappedData = [
      {
        'Name of the worker': 'Ajaykumar Mansingbhai',
        [`${FORM_NGJ_CASUAL_PARENT}_Period_From`]: '',
        [`${FORM_NGJ_CASUAL_PARENT}_Period_To`]: '',
        [`${FORM_NGJ_CASUAL_PARENT}_Total Leave`]: '',
        [`${FORM_NGJ_CASUAL_PARENT}_Availed Leave`]: '',
        [`${FORM_NGJ_CASUAL_PARENT}_Balance Leave`]: '',
      },
    ];
    const employees = [
      {
        FirstName: 'Ajaykumar',
        LastName: 'Mansingbhai',
        EmployeeID: 'VE0471',
        Zoho_ID: '9990001',
      },
    ];
    const approved = [
      {
        'Leave Type': 'Contingency Leave',
        From: '23-May-2026',
        To: '28-May-2026',
        Days: {
          '23-May-2026': { LeaveCount: 1 },
          '24-May-2026': { LeaveCount: 1 },
          '25-May-2026': { LeaveCount: 1 },
          '26-May-2026': { LeaveCount: 1 },
          '27-May-2026': { LeaveCount: 1 },
        },
        ApprovalStatus: 'APPROVED',
        'Employee Name': 'Ajaykumar Mehta',
        ZohoID: '8880001',
      },
      {
        'Leave Type': 'Contingency Leave',
        From: '23-May-2026',
        To: '28-May-2026',
        Days: {
          '23-May-2026': { LeaveCount: 1 },
          '24-May-2026': { LeaveCount: 1 },
          '25-May-2026': { LeaveCount: 1 },
          '26-May-2026': { LeaveCount: 1 },
          '27-May-2026': { LeaveCount: 1 },
        },
        ApprovalStatus: 'APPROVED',
        'Employee Name': 'Ajaykumar M',
        ZohoID: '7770001',
      },
      {
        'Leave Type': 'Contingency Leave',
        From: '26-May-2026',
        To: '28-May-2026',
        Days: {
          '26-May-2026': { LeaveCount: 1 },
          '27-May-2026': { LeaveCount: 1 },
          '28-May-2026': { LeaveCount: 1 },
        },
        ApprovalStatus: 'APPROVED',
        'Employee Name': 'Ajaykumar Mansingbhai',
        FirstName: 'Ajaykumar',
        LastName: 'Mansingbhai',
        ZohoID: '9990001',
      },
    ];
    const leaveRecords = [
      {
        employee: { name: 'Ajaykumar Mansingbhai', id: 'VE0471' },
        'Contingency Leave': { balance: 25, booked: 3 },
      },
    ];
    applyFormNGJGujaratCasualLeaveAutofill(
      mappedData,
      employees,
      headers,
      approved,
      leaveRecords,
      { monthFrom: '01-May-2026', monthTo: '31-May-2026' }
    );
    expect(mappedData[0][`${FORM_NGJ_CASUAL_PARENT}_Period_From`]).toBe('26-May-2026');
    expect(mappedData[0][`${FORM_NGJ_CASUAL_PARENT}_Period_To`]).toBe('28-May-2026');
    expect(mappedData[0][`${FORM_NGJ_CASUAL_PARENT}_Availed Leave`]).toBe('3');
    expect(mappedData[0][`${FORM_NGJ_CASUAL_PARENT}_Balance Leave`]).toBe('25');
    expect(mappedData[0][`${FORM_NGJ_CASUAL_PARENT}_Total Leave`]).toBe('28');
  });

  it('caps Availed at Contingency booked when LeaveCount exceeds booked (wrong leave guard)', () => {
    const approved = {
      'Leave Type': 'Contingency Leave',
      From: '23-May-2026',
      To: '28-May-2026',
      Days: {
        // Sat 23 + Sun 24 + Mon–Wed — Zoho marks weekends LeaveCount 1; Contingency booked is 3.
        '23-May-2026': { LeaveCount: 1 },
        '24-May-2026': { LeaveCount: 0 },
        '25-May-2026': { LeaveCount: 1 },
        '26-May-2026': { LeaveCount: 1 },
        '27-May-2026': { LeaveCount: 1 },
        '28-May-2026': { LeaveCount: 0 },
      },
    };
    const leaveRecord = {
      'Contingency Leave': { balance: 25, booked: 3 },
    };
    const values = buildFormNGJCasualLeaveValues(approved, leaveRecord);
    expect(values.from).toBe('25-May-2026');
    expect(values.to).toBe('27-May-2026');
    expect(values.availedLeave).toBe('3');
    expect(values.balanceLeave).toBe('25');
    expect(values.totalLeave).toBe('28');
  });

  it('rejects Contingency leave that only matches a duplicated Zoho ID with a different name', () => {
    const mappedData = [
      {
        'Name of the worker': 'Ajaykumar Mansingbhai',
        [`${FORM_NGJ_CASUAL_PARENT}_Period_From`]: '',
        [`${FORM_NGJ_CASUAL_PARENT}_Period_To`]: '',
        [`${FORM_NGJ_CASUAL_PARENT}_Total Leave`]: '',
        [`${FORM_NGJ_CASUAL_PARENT}_Availed Leave`]: '',
        [`${FORM_NGJ_CASUAL_PARENT}_Balance Leave`]: '',
      },
    ];
    const employees = [
      {
        FirstName: 'Ajaykumar',
        LastName: 'Mansingbhai',
        EmployeeID: 'VE0471',
        Zoho_ID: '313989000000559170',
      },
    ];
    // Same Zoho erecno as another worker (real Approved Leaves bug) + 5-day leave.
    const approved = [
      {
        'Leave Type': 'Contingency Leave',
        From: '23-May-2026',
        To: '28-May-2026',
        Days: {
          '23-May-2026': { LeaveCount: 1 },
          '24-May-2026': { LeaveCount: 1 },
          '25-May-2026': { LeaveCount: 1 },
          '26-May-2026': { LeaveCount: 1 },
          '27-May-2026': { LeaveCount: 1 },
        },
        ApprovalStatus: 'APPROVED',
        'Employee Name': 'Santhoshkumar Raman',
        ZohoID: '313989000000559170',
      },
      {
        'Leave Type': 'Contingency Leave',
        From: '26-May-2026',
        To: '28-May-2026',
        Days: {
          '26-May-2026': { LeaveCount: 1 },
          '27-May-2026': { LeaveCount: 1 },
          '28-May-2026': { LeaveCount: 1 },
        },
        ApprovalStatus: 'APPROVED',
        'Employee Name': 'Ajaykumar Mansingbhai',
        ZohoID: '9990001',
      },
    ];
    const leaveRecords = [
      {
        employee: { name: 'Ajaykumar Mansingbhai', id: 'VE0471' },
        'Contingency Leave': { balance: 25, booked: 3 },
      },
    ];
    applyFormNGJGujaratCasualLeaveAutofill(
      mappedData,
      employees,
      headers,
      approved,
      leaveRecords,
      { monthFrom: '01-May-2026', monthTo: '31-May-2026' }
    );
    expect(mappedData[0][`${FORM_NGJ_CASUAL_PARENT}_Period_From`]).toBe('26-May-2026');
    expect(mappedData[0][`${FORM_NGJ_CASUAL_PARENT}_Period_To`]).toBe('28-May-2026');
    expect(mappedData[0][`${FORM_NGJ_CASUAL_PARENT}_Availed Leave`]).toBe('3');
    expect(mappedData[0][`${FORM_NGJ_CASUAL_PARENT}_Balance Leave`]).toBe('25');
  });

  it('does not remap Festival leave values into Casual Leave columns', () => {
    const sourceHeaders = [
      'Name of the worker',
      `${FORM_NGJ_FESTIVAL_PARENT}_Total Leave`,
      `${FORM_NGJ_FESTIVAL_PARENT}_Availed Leave`,
      `${FORM_NGJ_FESTIVAL_PARENT}_Balance Leave`,
      `${FORM_NGJ_CASUAL_PARENT}_Total Leave`,
      `${FORM_NGJ_CASUAL_PARENT}_Availed Leave`,
      `${FORM_NGJ_CASUAL_PARENT}_Balance Leave`,
    ];
    const rows = [
      {
        'Name of the worker': 'Ravi Patel',
        [`${FORM_NGJ_FESTIVAL_PARENT}_Total Leave`]: '99',
        [`${FORM_NGJ_FESTIVAL_PARENT}_Availed Leave`]: '9',
        [`${FORM_NGJ_FESTIVAL_PARENT}_Balance Leave`]: '90',
        [`${FORM_NGJ_CASUAL_PARENT}_Total Leave`]: '28',
        [`${FORM_NGJ_CASUAL_PARENT}_Availed Leave`]: '3',
        [`${FORM_NGJ_CASUAL_PARENT}_Balance Leave`]: '25',
      },
    ];
    const remapped = remapFormNGJGujaratRowsToHeaders(rows, sourceHeaders, headers);
    expect(remapped[0][`${FORM_NGJ_CASUAL_PARENT}_Total Leave`]).toBe('28');
    expect(remapped[0][`${FORM_NGJ_CASUAL_PARENT}_Availed Leave`]).toBe('3');
    expect(remapped[0][`${FORM_NGJ_CASUAL_PARENT}_Balance Leave`]).toBe('25');
    expect(remapped[0][`${FORM_NGJ_FESTIVAL_PARENT}_Total Leave`]).toBe('99');
  });

  it('stores EmployeeID on Form N rows for LeaveData matching', () => {
    const row = applyFormNGJGujaratEmployeeToRow(
      {},
      { FirstName: 'Ajaykumar', LastName: 'Mansingbhai', EmployeeID: 'VE0471' },
      headers
    );
    expect(row.__employeeLookupName).toBe('Ajaykumar Mansingbhai');
    expect(row.__employeeLookupId).toBe('VE0471');
  });

  it('preserves table anchors for draft export (does not force headerRowIndex=-1)', () => {
    const layout = resolveFormNGJGujaratHeaderFieldLayout(
      {
        formHeader: { title: 'FORM - N', subtitle: 'LEAVE BOOK' },
        headers,
        headerRowIndex: 6,
        dataStartIndex: 9,
        tableStartCol: 0,
      },
      null,
      {
        formHeader: { title: 'FORM - N', subtitle: 'LEAVE BOOK' },
        fileName: 'Form_N_GJ.xlsx',
        tableHeaders: headers,
      }
    );
    expect(layout).not.toBeNull();
    expect(layout.headerRowIndex).toBe(6);
    expect(layout.dataStartIndex).toBe(9);
    expect(layout.headers.length).toBeGreaterThan(0);
  });
});

describe('Form N GJ ExcelJS workbook export', () => {
  const headers = resolveFormNGJGujaratTableHeaders();

  it('reads worker/leave values from a row without treating worker fields as leave columns', () => {
    const row = {
      'Name of the worker': 'Ravi Patel',
      'Description of the Department (if applicable)': 'Corrective Maintenance',
      'Name of the employer': 'Vayona',
      [`${FORM_NGJ_CASUAL_PARENT}_Period_From`]: '04-Jul-2026',
      [`${FORM_NGJ_CASUAL_PARENT}_Period_To`]: '05-Jul-2026',
      [`${FORM_NGJ_CASUAL_PARENT}_Availed Leave`]: '2',
    };
    expect(getFormNGJRowValueForHeader(row, 'Name of the worker')).toBe('Ravi Patel');
    expect(
      getFormNGJRowValueForHeader(row, 'Description of the Department (if applicable)')
    ).toBe('Corrective Maintenance');
    expect(rowHasMeaningfulFormNGJGujaratExportData(row, headers)).toBe(true);
    expect(filterFormNGJGujaratExportRows([row, {}], headers)).toHaveLength(1);
  });

  it('writes Leave Book into template with ExcelJS and keeps a valid xlsx (no SheetJS repair)', async () => {
    const ExcelJS = require('exceljs');
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Sheet1');
    ws.getCell(1, 1).value = 'FORM - N';
    ws.getCell(2, 1).value = '(See rule 17)';
    ws.getCell(3, 1).value = 'LEAVE BOOK';
    ws.getCell(5, 1).value = 'Name of the establishment:';
    ws.getCell(5, 3).value = 'Name of the worker:';
    ws.getCell(5, 5).value = 'Description of the Department (if applicable):';
    ws.getCell(5, 7).value = 'Name of the employer:';
    ws.getCell(6, 1).value = 'Receipt of level book:';
    ws.getCell(6, 5).value = 'Date of entry into service:';
    ws.getCell(8, 1).value = 'Accumulation of leave';
    ws.getCell(8, 3).value = 'Leave allowed';
    ws.getCell(8, 4).value = 'Payment for leave made on';
    ws.getCell(8, 6).value = 'Refusal of leave';
    ws.getCell(8, 8).value =
      'Payment for Leave on discharge of an worker quitting employment if admissible';
    ws.getCell(8, 11).value = 'Remarks';
    ws.mergeCells(8, 1, 8, 2);
    ws.mergeCells(8, 4, 8, 5);
    ws.mergeCells(8, 6, 8, 7);
    ws.mergeCells(8, 8, 8, 10);
    ws.getCell(9, 1).value = 'Leave due on';
    ws.getCell(9, 2).value = 'No. of days';
    ws.getCell(9, 3).value = 'From - To -';
    ws.getCell(9, 4).value = '1st Moiety';
    ws.getCell(9, 5).value = '2nd Moiety';
    ws.getCell(9, 6).value = 'Application Date';
    ws.getCell(9, 7).value = 'Date of Refusal';
    ws.getCell(9, 8).value = 'Date of discharge';
    ws.getCell(9, 9).value = 'Date and amount paid';
    ws.getCell(9, 10).value = 'Signature or thumb impression of worker';
    ws.getCell(9, 11).value = 'Remarks';
    ws.getCell(11, 1).value = 'DETAILS OF FESTIVAL LEAVE';
    ws.getCell(12, 1).value = 'Period';
    ws.getCell(12, 3).value = 'Total Leave';
    ws.getCell(12, 4).value = 'Availed Leave';
    ws.getCell(12, 5).value = 'Balance Leave';
    ws.getCell(12, 6).value = 'Payment made in lieu of Festival Leave, when called';
    ws.getCell(12, 7).value = 'Remarks';
    ws.mergeCells(12, 1, 12, 2);
    ws.getCell(13, 1).value = 'From';
    ws.getCell(13, 2).value = 'To';
    ws.getCell(15, 1).value = 'DETAILS OF CASUAL LEAVE';
    ws.getCell(16, 1).value = 'Period';
    ws.getCell(16, 3).value = 'Total Leave';
    ws.getCell(16, 4).value = 'Availed Leave';
    ws.getCell(16, 5).value = 'Balance Leave';
    ws.getCell(16, 6).value = 'Remarks';
    ws.mergeCells(16, 1, 16, 2);
    ws.getCell(17, 1).value = 'From';
    ws.getCell(17, 2).value = 'To';

    const templateArrayBuffer = await wb.xlsx.writeBuffer();
    const row = {
      'Name of the worker': 'Ravi Patel',
      'Description of the Department (if applicable)': 'Corrective Maintenance',
      'Name of the employer': 'Vayona Energy',
      'Date of entry into service': '01-Jan-2024',
      [`${FORM_NGJ_CASUAL_PARENT}_Period_From`]: '04-Jul-2026',
      [`${FORM_NGJ_CASUAL_PARENT}_Period_To`]: '05-Jul-2026',
      [`${FORM_NGJ_CASUAL_PARENT}_Total Leave`]: '10',
      [`${FORM_NGJ_CASUAL_PARENT}_Availed Leave`]: '2',
      [`${FORM_NGJ_CASUAL_PARENT}_Balance Leave`]: '8',
    };

    const { blob, fileName } = await buildFormNGJGujaratWorkbookWithTemplateStyles({
      templateArrayBuffer,
      mappedData: [row],
      headersToUse: headers,
      parsedFormHeader: { title: 'FORM - N', subtitle: 'LEAVE BOOK' },
      headerFormData: { form_n_gj_establishment: 'Site Alpha' },
      formFileName: 'Form_N_GJ.xlsx',
    });

    expect(fileName).toMatch(/Form_N_GJ/i);
    expect(blob).toBeInstanceOf(Blob);
    expect(blob.size).toBeGreaterThan(1000);

    const outWb = new ExcelJS.Workbook();
    const outBuffer = Buffer.from(await new Response(blob).arrayBuffer());
    await outWb.xlsx.load(outBuffer);
    const outWs = outWb.worksheets[0];
    expect(String(outWs.getCell(1, 1).value || '')).toContain('FORM - N');
    // Department must not land in leave "No. of days" (col B data row).
    expect(String(outWs.getCell(10, 2).value || '')).not.toBe('Corrective Maintenance');
    // Casual From/To written into casual section.
    expect(String(outWs.getCell(18, 1).value || '')).toBe('04-Jul-2026');
    expect(String(outWs.getCell(18, 2).value || '')).toBe('05-Jul-2026');
    expect(String(outWs.getCell(18, 4).value || '')).toBe('2');
  });
});
