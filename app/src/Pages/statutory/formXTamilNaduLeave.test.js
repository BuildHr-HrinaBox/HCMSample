import {
  buildFormXLeaveSectionValues,
  computeFormXLeaveBeginning,
  computeFormXLeaveEndBalance,
  getFormXEarnedLeaveApiMetrics,
  getFormXMedicalLeaveApiMetrics,
  getFormXOtherLeaveApiMetrics,
  leaveTypeLabelMatchesAliases,
  sumApprovedLeaveCountForEmployee,
  FORM_X_EARNED_LEAVE_TYPE_ALIASES,
  FORM_X_MEDICAL_LEAVE_TYPE_ALIASES,
  FORM_X_OTHER_LEAVE_TYPE_ALIASES,
} from './formXTamilNaduLeave';

describe('formXTamilNaduLeave', () => {
  it('computes beginning as Leave API balance + approved LeaveCount (Apr 24 → May 23)', () => {
    expect(computeFormXLeaveBeginning(23, 1)).toBe('24');
    expect(computeFormXLeaveBeginning(23, 0)).toBe('23');
    expect(computeFormXLeaveBeginning(23, '')).toBe('23');
    expect(computeFormXLeaveEndBalance(24, 1)).toBe('23');
  });

  it('maps Earned / Contingency / Legacy leave cells from Leave API rows', () => {
    const row = {
      employee: 'Rajeshkumar Inbaraj (VE0507)',
      'Earned Leave (Test)': JSON.stringify({ paidBalance: 72, paidBooked: 8 }),
      'Contingency Leave': JSON.stringify({ balance: 26, booked: 2 }),
      'Legacy Earned Leave': 0,
    };
    expect(getFormXEarnedLeaveApiMetrics(row).balance).toBe('72');
    expect(getFormXMedicalLeaveApiMetrics(row).balance).toBe('26');
    expect(getFormXOtherLeaveApiMetrics(row).balance).toBe('0');
  });

  it('builds section values with approved LeaveCount as availed', () => {
    const leaveRecord = {
      'Earned Leave': { paidBalance: 23, paidBooked: 1 },
    };
    const values = buildFormXLeaveSectionValues(leaveRecord, '1', {}, 'earned');
    expect(values.beginning).toBe('24');
    expect(values.availed).toBe('1');
    expect(values.balance).toBe('23');
  });

  it('sums LeaveCount from matching approved leave Days JSON', () => {
    const approved = [
      {
        Employee: 'Arun Pradhaban',
        'Employee.ID': '313989000000537200',
        'Leave Type': 'Earned Leave',
        From: '10-Apr-2026',
        To: '10-Apr-2026',
        Days: '{"10-Apr-2026":{"LeaveCount":"1.0","EndTime":"17:00","StartTime":"08:00"}}',
        ApprovalStatus: 'APPROVED',
      },
      {
        Employee: 'Arun Pradhaban',
        'Employee.ID': '313989000000537200',
        'Leave Type': 'Contingency Leave',
        From: '12-Apr-2026',
        To: '12-Apr-2026',
        Days: '{"12-Apr-2026":{"LeaveCount":"1.0"}}',
        ApprovalStatus: 'APPROVED',
      },
    ];
    const emp = { FirstName: 'Arun', LastName: 'Pradhaban', Zoho_ID: '313989000000537200' };
    const row = { Name: 'Arun Pradhaban' };
    const earned = sumApprovedLeaveCountForEmployee(
      approved,
      emp,
      row,
      '',
      'Name',
      FORM_X_EARNED_LEAVE_TYPE_ALIASES,
      { monthFrom: '01-Apr-2026', monthTo: '30-Apr-2026' }
    );
    const medical = sumApprovedLeaveCountForEmployee(
      approved,
      emp,
      row,
      '',
      'Name',
      FORM_X_MEDICAL_LEAVE_TYPE_ALIASES,
      { monthFrom: '01-Apr-2026', monthTo: '30-Apr-2026' }
    );
    const other = sumApprovedLeaveCountForEmployee(
      approved,
      emp,
      row,
      '',
      'Name',
      FORM_X_OTHER_LEAVE_TYPE_ALIASES,
      { monthFrom: '01-Apr-2026', monthTo: '30-Apr-2026' }
    );
    expect(earned).toBe('1');
    expect(medical).toBe('1');
    expect(other).toBe('');
  });

  it('matches leave type aliases for Contingency and Legacy', () => {
    expect(leaveTypeLabelMatchesAliases('Contingency Leave', FORM_X_MEDICAL_LEAVE_TYPE_ALIASES)).toBe(
      true
    );
    expect(leaveTypeLabelMatchesAliases('Medical Leave', FORM_X_MEDICAL_LEAVE_TYPE_ALIASES)).toBe(
      false
    );
    expect(leaveTypeLabelMatchesAliases('Legacy Earned Leave', FORM_X_OTHER_LEAVE_TYPE_ALIASES)).toBe(
      true
    );
    expect(leaveTypeLabelMatchesAliases('Legacy Earned Leave', FORM_X_EARNED_LEAVE_TYPE_ALIASES)).toBe(
      false
    );
    expect(leaveTypeLabelMatchesAliases('Earned Leave (Test)', FORM_X_EARNED_LEAVE_TYPE_ALIASES)).toBe(
      true
    );
  });
});
