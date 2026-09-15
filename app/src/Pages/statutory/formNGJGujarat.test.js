import {
  FORM_NGJ_CASUAL_HEADERS,
  FORM_NGJ_CASUAL_PARENT,
  FORM_NGJ_FESTIVAL_PARENT,
  FORM_NGJ_LEAVE_PARENT,
  applyFormNGJGujaratAccumulationLeaveAutofill,
  applyFormNGJGujaratAccumulationLeaveToRow,
  applyFormNGJGujaratCasualLeaveAutofill,
  applyFormNGJGujaratCasualLeaveToRow,
  applyFormNGJGujaratEmployeeToRow,
  buildFormNGJAccumulationLeaveValues,
  buildFormNGJCasualLeaveValues,
  buildFormNGJGujaratWorkbookWithTemplateStyles,
  filterApprovedLeaveRecordsForFormNGJCasual,
  filterFormNGJGujaratExportRows,
  getFormNGJRowValueForHeader,
  isApprovedLeaveContingencyForFormNGJ,
  isApprovedLeaveEarnedForFormNGJ,
  isFormNGJCasualAvailedLeaveHeader,
  isFormNGJCasualBalanceLeaveHeader,
  isFormNGJCasualPeriodFromHeader,
  isFormNGJCasualPeriodToHeader,
  isFormNGJCasualTotalLeaveHeader,
  isFormNGJLeaveAllowedFromToHeader,
  isFormNGJLeaveDueOnHeader,
  isFormNGJNoOfDaysHeader,
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

  it('sums multiple Contingency leaves in the same month (1 day + 5 days = 6)', () => {
    const mappedData = [
      {
        'Name of the worker': 'Ravi Patel',
        [`${FORM_NGJ_CASUAL_PARENT}_Period_From`]: '',
        [`${FORM_NGJ_CASUAL_PARENT}_Period_To`]: '',
        [`${FORM_NGJ_CASUAL_PARENT}_Total Leave`]: '',
        [`${FORM_NGJ_CASUAL_PARENT}_Availed Leave`]: '',
        [`${FORM_NGJ_CASUAL_PARENT}_Balance Leave`]: '',
      },
    ];
    const employees = [{ FirstName: 'Ravi', LastName: 'Patel', EmployeeID: 'VE2001' }];
    const approved = [
      {
        'Leave Type': 'Contingency Leave',
        From: '08-Jul-2026',
        To: '08-Jul-2026',
        Days: { '08-Jul-2026': { LeaveCount: 1 } },
        ApprovalStatus: 'APPROVED',
        'Employee Name': 'Ravi Patel',
        EmployeeID: 'VE2001',
        ZohoID: 'leave-1',
      },
      {
        'Leave Type': 'Contingency Leave',
        From: '21-Jul-2026',
        To: '25-Jul-2026',
        Days: {
          '21-Jul-2026': { LeaveCount: 1 },
          '22-Jul-2026': { LeaveCount: 1 },
          '23-Jul-2026': { LeaveCount: 1 },
          '24-Jul-2026': { LeaveCount: 1 },
          '25-Jul-2026': { LeaveCount: 1 },
        },
        ApprovalStatus: 'APPROVED',
        'Employee Name': 'Ravi Patel',
        EmployeeID: 'VE2001',
        ZohoID: 'leave-2',
      },
    ];
    const leaveRecords = [
      {
        employee: { name: 'Ravi Patel', id: 'VE2001' },
        'Contingency Leave': { balance: 10, booked: 6 },
      },
    ];
    applyFormNGJGujaratCasualLeaveAutofill(
      mappedData,
      employees,
      headers,
      approved,
      leaveRecords,
      { monthFrom: '01-Jul-2026', monthTo: '31-Jul-2026' }
    );
    expect(mappedData[0][`${FORM_NGJ_CASUAL_PARENT}_Period_From`]).toBe(
      '08-Jul-2026, 21-Jul-2026 to 25-Jul-2026'
    );
    expect(mappedData[0][`${FORM_NGJ_CASUAL_PARENT}_Period_To`]).toBe(
      '08-Jul-2026, 21-Jul-2026 to 25-Jul-2026'
    );
    expect(mappedData[0][`${FORM_NGJ_CASUAL_PARENT}_Availed Leave`]).toBe('6');
    expect(mappedData[0][`${FORM_NGJ_CASUAL_PARENT}_Balance Leave`]).toBe('10');
    expect(mappedData[0][`${FORM_NGJ_CASUAL_PARENT}_Total Leave`]).toBe('16');
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

describe('Form N GJ Accumulation of leave (Earned Leave)', () => {
  const headers = resolveFormNGJGujaratTableHeaders();
  const dueOnHeader = `${FORM_NGJ_LEAVE_PARENT.accumulation}_Leave due on`;
  const daysHeader = `${FORM_NGJ_LEAVE_PARENT.accumulation}_No. of days`;
  const fromToHeader = `${FORM_NGJ_LEAVE_PARENT.leaveAllowed}_From To`;

  it('recognizes Leave due on / No. of days / Leave allowed From-To headers', () => {
    expect(isFormNGJLeaveDueOnHeader(dueOnHeader)).toBe(true);
    expect(isFormNGJNoOfDaysHeader(daysHeader)).toBe(true);
    expect(isFormNGJLeaveAllowedFromToHeader(fromToHeader)).toBe(true);
    expect(isFormNGJLeaveDueOnHeader(`${FORM_NGJ_CASUAL_PARENT}_Period_From`)).toBe(false);
    expect(isFormNGJNoOfDaysHeader(`${FORM_NGJ_FESTIVAL_PARENT}_Total Leave`)).toBe(false);
  });

  it('treats Earned Leave approvals as accumulation, not Contingency', () => {
    expect(isApprovedLeaveEarnedForFormNGJ({ 'Leave Type': 'Earned Leave' })).toBe(true);
    expect(isApprovedLeaveEarnedForFormNGJ({ 'Leave Type': 'Contingency Leave' })).toBe(false);
    expect(isApprovedLeaveContingencyForFormNGJ({ 'Leave Type': 'Earned Leave' })).toBe(false);
  });

  it('keeps Leave due on blank and uses Earned Leave booked count for No. of days', () => {
    const values = buildFormNGJAccumulationLeaveValues(
      null,
      { 'Earned Leave': { paidBalance: 10, paidBooked: 2 } },
      { monthFrom: '01-Sep-2026' }
    );
    expect(values.leaveDueOn).toBe('');
    expect(values.noOfDays).toBe('2');
    expect(values.leaveFromTo).toBe('');
    expect(values.hasLeaveMetrics).toBe(true);
  });

  it('fills Leave allowed From-To from approved Earned Leave, not Contingency', () => {
    const values = buildFormNGJAccumulationLeaveValues(
      {
        'Leave Type': 'Earned Leave',
        From: '04-Sep-2026',
        To: '05-Sep-2026',
        Days: {
          '04-Sep-2026': { LeaveCount: 1 },
          '05-Sep-2026': { LeaveCount: 1 },
        },
      },
      { 'Earned Leave': { paidBalance: 8, paidBooked: 2 } },
      { monthFrom: '01-Sep-2026' }
    );
    expect(values.leaveDueOn).toBe('');
    expect(values.noOfDays).toBe('2');
    expect(values.leaveFromTo).toBe('04-Sep-2026 - 05-Sep-2026');
  });

  it('applies accumulation cells onto a Form N row', () => {
    const row = { 'Name of the worker': 'Ravi Patel', [dueOnHeader]: '01-Sep-2026' };
    const applied = applyFormNGJGujaratAccumulationLeaveToRow(
      row,
      {
        'Leave Type': 'Earned Leave',
        From: '10-Sep-2026',
        To: '11-Sep-2026',
        Days: {
          '10-Sep-2026': { LeaveCount: 1 },
          '11-Sep-2026': { LeaveCount: 1 },
        },
      },
      { 'Earned Leave': { paidBalance: 5, paidBooked: 2 } },
      headers,
      { monthFrom: '01-Sep-2026' }
    );
    expect(applied).toBeGreaterThan(0);
    expect(row[dueOnHeader]).toBe('');
    expect(row[daysHeader]).toBe('2');
    expect(row[fromToHeader]).toBe('10-Sep-2026 - 11-Sep-2026');
  });

  it('sums multiple Earned Leave periods for No. of days and From-To', () => {
    const values = buildFormNGJAccumulationLeaveValues(
      [
        {
          'Leave Type': 'Earned Leave',
          From: '08-Jul-2026',
          To: '08-Jul-2026',
          Days: { '08-Jul-2026': { LeaveCount: 1 } },
        },
        {
          'Leave Type': 'Earned Leave',
          From: '21-Jul-2026',
          To: '25-Jul-2026',
          Days: {
            '21-Jul-2026': { LeaveCount: 1 },
            '22-Jul-2026': { LeaveCount: 1 },
            '23-Jul-2026': { LeaveCount: 1 },
            '24-Jul-2026': { LeaveCount: 1 },
            '25-Jul-2026': { LeaveCount: 1 },
          },
        },
      ],
      { 'Earned Leave': { paidBalance: 4, paidBooked: 6 } },
      { monthFrom: '01-Jul-2026', monthTo: '31-Jul-2026' }
    );
    expect(values.noOfDays).toBe('6');
    expect(values.leaveFromTo).toBe('08-Jul-2026, 21-Jul-2026 to 25-Jul-2026');
  });

  it('autofills Accumulation of leave from Earned Leave matching', () => {
    const mappedData = [
      {
        'Name of the worker': 'Chetan Kumar',
        [dueOnHeader]: '',
        [daysHeader]: '',
        [fromToHeader]: '',
      },
    ];
    const employees = [{ FirstName: 'Chetan', LastName: 'Kumar', EmployeeID: 'VE1001' }];
    const approved = [
      {
        'Leave Type': 'Earned Leave',
        From: '04-Sep-2026',
        To: '05-Sep-2026',
        Days: {
          '04-Sep-2026': { LeaveCount: 1 },
          '05-Sep-2026': { LeaveCount: 1 },
        },
        EmployeeID: 'VE1001',
        'Employee Name': 'Chetan Kumar',
      },
      {
        'Leave Type': 'Contingency Leave',
        From: '08-Sep-2026',
        To: '08-Sep-2026',
        Days: { '08-Sep-2026': { LeaveCount: 1 } },
        EmployeeID: 'VE1001',
        'Employee Name': 'Chetan Kumar',
      },
    ];
    const leaveRecords = [
      {
        EmployeeID: 'VE1001',
        EmployeeName: 'Chetan Kumar',
        'Earned Leave': { paidBalance: 14, paidBooked: 2 },
      },
    ];
    const hits = applyFormNGJGujaratAccumulationLeaveAutofill(
      mappedData,
      employees,
      headers,
      approved,
      leaveRecords,
      { monthFrom: '01-Sep-2026', monthTo: '30-Sep-2026' }
    );
    expect(hits).toBe(1);
    expect(mappedData[0][dueOnHeader]).toBe('');
    expect(mappedData[0][daysHeader]).toBe('2');
    expect(mappedData[0][fromToHeader]).toBe('04-Sep-2026 - 05-Sep-2026');
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
      [`${FORM_NGJ_LEAVE_PARENT.accumulation}_Leave due on`]: '01-Jul-2026',
      [`${FORM_NGJ_LEAVE_PARENT.accumulation}_No. of days`]: '12',
      [`${FORM_NGJ_LEAVE_PARENT.leaveAllowed}_From To`]: '04-Jul-2026 - 05-Jul-2026',
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
    expect(String(outWs.getCell(10, 1).value || '')).toBe('01-Jul-2026');
    expect(String(outWs.getCell(10, 2).value || '')).toBe('12');
    expect(String(outWs.getCell(10, 3).value || '')).toBe('04-Jul-2026 - 05-Jul-2026');
    expect(formNGJSheetContains(outWs, 'Site Alpha')).toBe(true);
    expect(formNGJSheetContains(outWs, 'Ravi Patel')).toBe(true);
    expect(formNGJSheetContains(outWs, 'Corrective Maintenance')).toBe(true);
    // Casual From/To written into casual section.
    expect(String(outWs.getCell(18, 1).value || '')).toBe('04-Jul-2026');
    expect(String(outWs.getCell(18, 2).value || '')).toBe('05-Jul-2026');
    expect(String(outWs.getCell(18, 4).value || '')).toBe('2');
  });

  it('writes establishment, worker and department into stacked Leave Book header labels', async () => {
    const ExcelJS = require('exceljs');
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Sheet1');
    ws.getCell(2, 6).value = 'FORM - N';
    ws.getCell(3, 6).value = '(See rule 17)';
    ws.getCell(4, 6).value = 'LEAVE BOOK';
    ws.mergeCells(7, 2, 7, 6);
    ws.getCell(7, 2).value = 'Name of the establishment:';
    ws.mergeCells(8, 2, 8, 6);
    ws.getCell(8, 2).value = 'Name of the worker:';
    ws.mergeCells(9, 2, 10, 6);
    ws.getCell(9, 2).value = 'Description of the Department\n(if applicable)';
    ws.getCell(7, 7).value = 'Name of the employer:';
    ws.getCell(7, 11).value = 'Receipt of level book';
    ws.getCell(8, 7).value = 'Date of entry into service';
    ws.getCell(11, 2).value = 'Accumulation of leave';
    ws.getCell(11, 4).value = 'Leave allowed';
    ws.getCell(11, 5).value = 'Payment for leave made on';
    ws.getCell(11, 7).value = 'Refusal of leave';
    ws.getCell(11, 9).value =
      'Payment for Leave on discharge of an worker quitting employment if admissible';
    ws.getCell(11, 12).value = 'Remarks';
    ws.getCell(12, 2).value = 'Leave due on';
    ws.getCell(12, 3).value = 'No. of days';
    ws.getCell(12, 4).value = 'From - To -';
    ws.getCell(12, 5).value = '1st Moiety';
    ws.getCell(12, 6).value = '2nd Moiety';
    ws.getCell(12, 7).value = 'Application Date';
    ws.getCell(12, 8).value = 'Date of Refusal';
    ws.getCell(12, 9).value = 'Date of discharge';
    ws.getCell(12, 10).value = 'Date and amount paid';
    ws.getCell(12, 11).value = 'Signature or thumb impression of worker';
    ws.getCell(12, 12).value = 'Remarks';
    ws.getCell(16, 2).value = 'DETAILS OF FESTIVAL LEAVE';
    ws.getCell(17, 2).value = 'Period';
    ws.getCell(17, 4).value = 'Total Leave';
    ws.getCell(17, 5).value = 'Availed Leave';
    ws.getCell(17, 6).value = 'Balance Leave';
    ws.getCell(17, 7).value = 'Payment made in lieu of Festival Leave, when called';
    ws.getCell(17, 8).value = 'Remarks';
    ws.mergeCells(17, 2, 17, 3);
    ws.getCell(18, 2).value = 'From';
    ws.getCell(18, 3).value = 'To';
    ws.getCell(21, 2).value = 'DETAILS OF CASUAL LEAVE';
    ws.getCell(22, 2).value = 'Period';
    ws.getCell(22, 4).value = 'Total Leave';
    ws.getCell(22, 5).value = 'Availed Leave';
    ws.getCell(22, 6).value = 'Balance Leave';
    ws.getCell(22, 7).value = 'Remarks';
    ws.mergeCells(22, 2, 22, 3);
    ws.getCell(23, 2).value = 'From';
    ws.getCell(23, 3).value = 'To';

    const templateArrayBuffer = await wb.xlsx.writeBuffer();
    const { blob } = await buildFormNGJGujaratWorkbookWithTemplateStyles({
      templateArrayBuffer,
      mappedData: [
        {
          'Name of the worker': 'Ajaykumar Mansingbhai',
          'Description of the Department (if applicable)': 'Corrective Maintenance',
          __employeeLookupName: 'Ajaykumar Mansingbhai',
        },
      ],
      headersToUse: headers,
      parsedFormHeader: { title: 'FORM - N', subtitle: 'LEAVE BOOK' },
      headerFormData: { form_n_gj_establishment: 'Vayona Energy Pvt Ltd' },
      formFileName: 'Form_N_GJ.xlsx',
    });

    const outWb = new ExcelJS.Workbook();
    await outWb.xlsx.load(Buffer.from(await new Response(blob).arrayBuffer()));
    const outWs = outWb.worksheets[0];
    expect(formNGJSheetContains(outWs, 'Vayona Energy Pvt Ltd')).toBe(true);
    expect(formNGJSheetContains(outWs, 'Ajaykumar Mansingbhai')).toBe(true);
    expect(formNGJSheetContains(outWs, 'Corrective Maintenance')).toBe(true);
    expect(String(outWs.getCell(8, 2).value || '')).toMatch(/Ajaykumar Mansingbhai/);
    expect(String(outWs.getCell(9, 2).value || '')).toMatch(/Corrective Maintenance/);
    expect(String(outWs.getCell(7, 2).value || '')).toMatch(/Vayona Energy Pvt Ltd/);
  });

  it('keeps worker and department lines inside the official merged Leave Book header box', async () => {
    const ExcelJS = require('exceljs');
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Sheet1');
    ws.getCell(2, 6).value = 'FORM - N';
    ws.getCell(3, 6).value = '(See rule 17)';
    ws.getCell(4, 6).value = 'LEAVE BOOK';
    ws.mergeCells(7, 2, 10, 6);
    ws.getCell(7, 2).value = [
      'Name of the establishment:',
      'Name of the worker:',
      'Description of the Department',
      '(if applicable)',
    ].join('\n');
    ws.mergeCells(7, 7, 10, 12);
    ws.getCell(7, 7).value = [
      'Name of the employer:',
      'Receipt of level book',
      'Date of entry into service',
    ].join('\n');
    ws.getCell(11, 2).value = 'Accumulation of leave';
    ws.getCell(11, 4).value = 'Leave allowed';
    ws.getCell(11, 5).value = 'Payment for leave made on';
    ws.getCell(11, 7).value = 'Refusal of leave';
    ws.getCell(11, 9).value =
      'Payment for Leave on discharge of an worker quitting employment if admissible';
    ws.getCell(11, 12).value = 'Remarks';
    ws.getCell(12, 2).value = 'Leave due on';
    ws.getCell(12, 3).value = 'No. of days';
    ws.getCell(12, 4).value = 'From - To -';
    ws.getCell(12, 5).value = '1st Moiety';
    ws.getCell(12, 6).value = '2nd Moiety';
    ws.getCell(12, 7).value = 'Application Date';
    ws.getCell(12, 8).value = 'Date of Refusal';
    ws.getCell(12, 9).value = 'Date of discharge';
    ws.getCell(12, 10).value = 'Date and amount paid';
    ws.getCell(12, 11).value = 'Signature or thumb impression of worker';
    ws.getCell(12, 12).value = 'Remarks';
    ws.getCell(16, 2).value = 'DETAILS OF FESTIVAL LEAVE';
    ws.getCell(17, 2).value = 'Period';
    ws.getCell(17, 4).value = 'Total Leave';
    ws.getCell(17, 5).value = 'Availed Leave';
    ws.getCell(17, 6).value = 'Balance Leave';
    ws.mergeCells(17, 2, 17, 3);
    ws.getCell(18, 2).value = 'From';
    ws.getCell(18, 3).value = 'To';
    ws.getCell(21, 2).value = 'DETAILS OF CASUAL LEAVE';
    ws.getCell(22, 2).value = 'Period';
    ws.getCell(22, 4).value = 'Total Leave';
    ws.getCell(22, 5).value = 'Availed Leave';
    ws.getCell(22, 6).value = 'Balance Leave';
    ws.mergeCells(22, 2, 22, 3);
    ws.getCell(23, 2).value = 'From';
    ws.getCell(23, 3).value = 'To';

    const templateArrayBuffer = await wb.xlsx.writeBuffer();
    const { blob } = await buildFormNGJGujaratWorkbookWithTemplateStyles({
      templateArrayBuffer,
      mappedData: [
        {
          'Name of the worker': 'Ajaykumar Mansingbhai',
          'Description of the Department (if applicable)': 'Corrective Maintenance',
          'Name of the employer': 'Vayona Energy',
          'Date of entry into service': '01 Dec 2025',
          __employeeLookupName: 'Ajaykumar Mansingbhai',
        },
      ],
      headersToUse: headers,
      parsedFormHeader: { title: 'FORM - N', subtitle: 'LEAVE BOOK' },
      headerFormData: {
        form_n_gj_establishment:
          'Maliya Site\nNo.114/1,Khirai Patiya behind Kerosene Depot,Taluka-Maliya,Dist-Morbi GJ - 363670,\nMaliya, Gujarat',
      },
      formFileName: 'Form_N_GJ.xlsx',
    });

    const outWb = new ExcelJS.Workbook();
    await outWb.xlsx.load(Buffer.from(await new Response(blob).arrayBuffer()));
    const left = String(outWb.worksheets[0].getCell(7, 2).value || '');
    const right = String(outWb.worksheets[0].getCell(7, 7).value || '');
    expect(left).toMatch(/Name of the establishment:\s*Maliya Site/i);
    expect(left).toMatch(/Name of the worker:\s*Ajaykumar Mansingbhai/i);
    expect(left).toMatch(/Corrective Maintenance/);
    expect(left).toMatch(/Description of the Department/i);
    expect((left.match(/Khirai Patiya/g) || []).length).toBe(1);
    expect((left.match(/Corrective Maintenance/g) || []).length).toBe(1);
    expect(right).toMatch(/Name of the employer:\s*Vayona Energy/i);
    expect(right).not.toMatch(/Name of the employer:\s*01 Dec 2025/i);
    expect(right).toMatch(/Date of entry into service:\s*01 Dec 2025/i);
  });
});

function formNGJSheetContains(ws, text) {
  const needle = String(text || '').trim();
  if (!needle) return false;
  for (let r = 1; r <= 30; r += 1) {
    for (let c = 1; c <= 14; c += 1) {
      if (String(ws.getCell(r, c).value || '').includes(needle)) return true;
    }
  }
  return false;
}
