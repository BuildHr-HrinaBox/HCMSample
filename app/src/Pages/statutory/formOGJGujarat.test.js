import {
  FORM_OGJ_PERIOD_PARENT,
  applyFormOGJGujaratApprovedLeaveAutofill,
  formOGJPersonNamesMatchStrict,
  formOGJRecordMatchesWorker,
  resolveFormOGJGujaratTableHeaders,
  toFormOGJPersonNameDisplay,
} from './formOGJGujarat';

const headers = resolveFormOGJGujaratTableHeaders();
const leaveCountHeader = 'Number of accumulated leave';
const fromHeader = `${FORM_OGJ_PERIOD_PARENT}_From`;
const tillHeader = `${FORM_OGJ_PERIOD_PARENT}_Till`;

describe('Form O Gujarat approved leave matching', () => {
  it('does not match on last-name-initial alone', () => {
    expect(formOGJPersonNamesMatchStrict('Harshad Dk', 'Harshad Dineshkumar')).toBe(false);
    expect(
      formOGJPersonNamesMatchStrict('Harshad Dk', 'Harshad Dineshkumar', 'Harshad', 'Dk')
    ).toBe(false);
    expect(
      formOGJPersonNamesMatchStrict(
        'Shailesh Babaria',
        'Shailesh Babaria',
        'Shailesh',
        'Babaria'
      )
    ).toBe(true);
  });

  it('rejects leave that only matches a duplicated Zoho ID with a different name', () => {
    const mappedData = [
      {
        'Name of Workers': 'Shailesh Babaria',
        [leaveCountHeader]: '',
        [fromHeader]: '',
        [tillHeader]: '',
      },
    ];
    const employees = [
      {
        FirstName: 'Shailesh',
        LastName: 'Babaria',
        EmployeeID: 'VE0102',
        Zoho_ID: '313989000000559170',
      },
    ];
    const approved = [
      {
        'Leave Type': 'Contingency Leave',
        From: '01-Jun-2026',
        To: '05-Jun-2026',
        Days: {
          '01-Jun-2026': { LeaveCount: 1 },
          '02-Jun-2026': { LeaveCount: 1 },
          '03-Jun-2026': { LeaveCount: 1 },
          '04-Jun-2026': { LeaveCount: 1 },
          '05-Jun-2026': { LeaveCount: 1 },
        },
        ApprovalStatus: 'APPROVED',
        'Employee Name': 'Santhoshkumar Raman',
        ZohoID: '313989000000559170',
      },
      {
        'Leave Type': 'Contingency Leave',
        From: '03-Jun-2026',
        To: '03-Jun-2026',
        Days: {
          '03-Jun-2026': { LeaveCount: 0.5 },
        },
        ApprovalStatus: 'APPROVED',
        'Employee Name': 'Shailesh Babaria',
        ZohoID: '9990002',
      },
    ];

    applyFormOGJGujaratApprovedLeaveAutofill(mappedData, employees, headers, approved, {
      monthFrom: '01-Jun-2026',
      monthTo: '30-Jun-2026',
    });

    expect(mappedData[0][leaveCountHeader]).toBe('0.5');
    expect(mappedData[0][fromHeader]).toBe('03-Jun-2026');
    expect(mappedData[0][tillHeader]).toBe('03-Jun-2026');
  });

  it('clears leave cells when the only ID hit belongs to another person', () => {
    const mappedData = [
      {
        'Name of Workers': 'Harshad Dk',
        [leaveCountHeader]: '9',
        [fromHeader]: '01-Jun-2026',
        [tillHeader]: '09-Jun-2026',
      },
    ];
    const employees = [
      {
        FirstName: 'Harshad',
        LastName: 'Dk',
        EmployeeID: 'VE0103',
        Zoho_ID: '313989000000559170',
      },
    ];
    const approved = [
      {
        'Leave Type': 'Contingency Leave',
        From: '01-Jun-2026',
        To: '05-Jun-2026',
        Days: {
          '01-Jun-2026': { LeaveCount: 1 },
          '02-Jun-2026': { LeaveCount: 1 },
          '03-Jun-2026': { LeaveCount: 1 },
          '04-Jun-2026': { LeaveCount: 1 },
          '05-Jun-2026': { LeaveCount: 1 },
        },
        ApprovalStatus: 'APPROVED',
        'Employee Name': 'Santhoshkumar Raman',
        ZohoID: '313989000000559170',
      },
    ];

    applyFormOGJGujaratApprovedLeaveAutofill(mappedData, employees, headers, approved, {
      monthFrom: '01-Jun-2026',
      monthTo: '30-Jun-2026',
    });

    expect(mappedData[0][leaveCountHeader]).toBe('');
    expect(mappedData[0][fromHeader]).toBe('');
    expect(mappedData[0][tillHeader]).toBe('');
  });

  it('uses positive LeaveCount days for From/Till (skips LeaveCount 0.0 weekly offs)', () => {
    const mappedData = [
      {
        'Name of Workers': 'Patel Chandrakant Virabhai',
        [leaveCountHeader]: '',
        [fromHeader]: '',
        [tillHeader]: '',
      },
    ];
    const employees = [
      {
        FirstName: 'Patel',
        LastName: 'Virabhai',
        EmployeeID: 'VE0104',
      },
    ];
    const approved = [
      {
        'Leave Type': 'Contingency Leave',
        From: '11-Jun-2026',
        To: '13-Jun-2026',
        Days: {
          '11-Jun-2026': { LeaveCount: 0 },
          '12-Jun-2026': { LeaveCount: 1 },
          '13-Jun-2026': { LeaveCount: 0 },
        },
        ApprovalStatus: 'APPROVED',
        'Employee Name': 'Patel Chandrakant Virabhai',
        EmployeeID: 'VE0104',
      },
    ];

    applyFormOGJGujaratApprovedLeaveAutofill(mappedData, employees, headers, approved, {
      monthFrom: '01-Jun-2026',
      monthTo: '30-Jun-2026',
    });

    expect(mappedData[0][leaveCountHeader]).toBe('1');
    expect(mappedData[0][fromHeader]).toBe('12-Jun-2026');
    expect(mappedData[0][tillHeader]).toBe('12-Jun-2026');
  });

  it('matches by EmployeeID even when Zoho_ID is absent', () => {
    const emp = { FirstName: 'Shailesh', LastName: 'Babaria', EmployeeID: 'VE0102' };
    const record = {
      'Employee Name': 'Shailesh Babaria',
      EmployeeID: 'VE0102',
      ApprovalStatus: 'APPROVED',
      From: '03-Jun-2026',
      To: '03-Jun-2026',
      Days: { '03-Jun-2026': { LeaveCount: 0.5 } },
    };
    expect(formOGJRecordMatchesWorker(record, 'Shailesh Babaria', emp)).toBe(true);
  });

  it('title-cases lowercase lookup names for Excel download', () => {
    expect(toFormOGJPersonNameDisplay('harshad dk')).toBe('Harshad Dk');
    expect(toFormOGJPersonNameDisplay('shailesh babaria')).toBe('Shailesh Babaria');
    expect(toFormOGJPersonNameDisplay('patel chandrakant virabhai')).toBe(
      'Patel Chandrakant Virabhai'
    );
  });
});
