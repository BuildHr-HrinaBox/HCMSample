import {
  FORM_OGJ_PERIOD_PARENT,
  applyFormOGJGujaratApprovedLeaveAutofill,
  formOGJPersonNamesMatchStrict,
  formOGJRecordMatchesWorker,
  isFormOGJGujaratContext,
  isFormPKarnatakaAccumulatedLeaveContext,
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

  it('matches Patel Chandrakant Virabhai when FirstName includes the middle name', () => {
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
        FirstName: 'Patel Chandrakant',
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

  it('matches Patel Chandrakant Virabhai when FirstName is the given name only', () => {
    expect(
      formOGJPersonNamesMatchStrict(
        'Patel Chandrakant Virabhai',
        'Patel Chandrakant Virabhai',
        'Chandrakant',
        'Virabhai'
      )
    ).toBe(true);
    expect(
      formOGJPersonNamesMatchStrict(
        'Patel Chandrakant Virabhai',
        'Patel Chandrakant Virabhai',
        'Patel Chandrakant',
        'Virabhai'
      )
    ).toBe(true);
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

  it('sums multiple approved leaves in the same month (1 day + 5 days = 6)', () => {
    const mappedData = [
      {
        'Name of Workers': 'Ravi Patel',
        [leaveCountHeader]: '',
        [fromHeader]: '',
        [tillHeader]: '',
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

    applyFormOGJGujaratApprovedLeaveAutofill(mappedData, employees, headers, approved, {
      monthFrom: '01-Jul-2026',
      monthTo: '31-Jul-2026',
    });

    expect(mappedData[0][leaveCountHeader]).toBe('6');
    expect(mappedData[0][fromHeader]).toBe('08-Jul-2026, 21-Jul-2026 to 25-Jul-2026');
    expect(mappedData[0][tillHeader]).toBe('08-Jul-2026, 21-Jul-2026 to 25-Jul-2026');
  });
});

describe('Form P Karnataka accumulated leave', () => {
  const kaHeaders = [
    'Sr. No.',
    'Number of accumulated leave',
    'Period for which leave is accumulated From_From',
    'Period for which leave is accumulated Till_Till',
  ];

  it('detects Form_P_Karnataka accumulated-leave notice, not Gujarat Form P muster', () => {
    expect(
      isFormPKarnatakaAccumulatedLeaveContext(
        { title: 'FORM P' },
        { formFileName: 'Form_P_-_Karnataka.xlsx' },
        'Form_P_-_Karnataka.xlsx',
        '',
        kaHeaders
      )
    ).toBe(true);
    expect(
      isFormOGJGujaratContext(
        { title: 'FORM P' },
        { formFileName: 'Form_P_-_Karnataka.xlsx' },
        'Form_P_-_Karnataka.xlsx',
        '',
        kaHeaders
      )
    ).toBe(true);
    expect(
      isFormOGJGujaratContext(
        { title: 'FORM P' },
        { formFileName: 'Form_P_GJ.xlsx' },
        'Form_P_GJ.xlsx',
        'MUSTER-ROLL Date of the Month',
        ['Sr. No.', 'Full Name of the Worker', 'Date of Month_1']
      )
    ).toBe(false);
  });

  it('autofills Number of accumulated leave / From / Till from approved leave', () => {
    const mappedData = [
      {
        'Sr. No.': '1',
        'Number of accumulated leave': '',
        'Period for which leave is accumulated From_From': '',
        'Period for which leave is accumulated Till_Till': '',
      },
    ];
    const employees = [{ FirstName: 'Vijeesh', LastName: 'Vijayan', EmployeeID: 'VE0761' }];
    const approved = [
      {
        'Leave Type': 'Earned Leave',
        From: '02-May-2026',
        To: '03-May-2026',
        Days: {
          '02-May-2026': { LeaveCount: 1 },
          '03-May-2026': { LeaveCount: 1 },
        },
        ApprovalStatus: 'APPROVED',
        'Employee Name': 'Vijeesh Vijayan',
        EmployeeID: 'VE0761',
        ZohoID: 'leave-pka-1',
      },
    ];

    const hits = applyFormOGJGujaratApprovedLeaveAutofill(
      mappedData,
      employees,
      kaHeaders,
      approved,
      { monthFrom: '01-May-2026', monthTo: '31-May-2026' }
    );
    expect(hits).toBe(1);
    expect(mappedData[0][leaveCountHeader]).toBe('2');
    expect(mappedData[0][fromHeader]).toBe('02-May-2026');
    expect(mappedData[0][tillHeader]).toBe('03-May-2026');
  });
});
