import {
  applyForm15Part1TamilNaduLeaveToRow,
  buildForm15Part1LeaveSectionValues,
  resolveForm15Part1LeaveColumnHeaders,
  toFormXSectionHeadersFromForm15,
} from './form15TamilNaduLeave';

describe('form15TamilNaduLeave (Form X mapping)', () => {
  const headers = [
    'S.No',
    'Name of the Worker',
    'Worker Identity No.',
    'Leave at the beginning of the Month',
    'Leave earned during the Period',
    'Leave availed during the Month',
    'Leave balance at the end of the Month',
    'Leave at beginning of the Month',
    'Leave availed during the Month',
    'Leave balance at end of the Month',
  ];
  const groupLabels = [
    '',
    '',
    '',
    'Earned Leave',
    'Earned Leave',
    'Earned Leave',
    'Earned Leave',
    'Medical Leave',
    'Medical Leave',
    'Medical Leave',
  ];

  it('resolves Earned + Medical leaf columns under group headers', () => {
    const resolved = resolveForm15Part1LeaveColumnHeaders(headers, groupLabels);
    expect(resolved.earnedBeginning).toBe('Leave at the beginning of the Month');
    expect(resolved.earnedDuring).toBe('Leave earned during the Period');
    expect(resolved.earnedAvailed).toBe('Leave availed during the Month');
    expect(resolved.earnedBalance).toBe('Leave balance at the end of the Month');
    expect(resolved.medicalBeginning).toBe('Leave at beginning of the Month');
    expect(resolved.medicalAvailed).toBe('Leave availed during the Month');
    expect(resolved.medicalBalance).toBe('Leave balance at end of the Month');
  });

  it('falls back to leaf order when group labels exist but are blank', () => {
    const blankGroups = headers.map(() => '');
    const resolved = resolveForm15Part1LeaveColumnHeaders(headers, blankGroups);
    expect(resolved.earnedBeginning).toBe('Leave at the beginning of the Month');
    expect(resolved.earnedDuring).toBe('Leave earned during the Period');
    expect(resolved.earnedAvailed).toBe('Leave availed during the Month');
    expect(resolved.earnedBalance).toBe('Leave balance at the end of the Month');
    expect(resolved.medicalBeginning).toBe('Leave at beginning of the Month');
    expect(resolved.medicalAvailed).toBe('Leave availed during the Month');
    expect(resolved.medicalBalance).toBe('Leave balance at end of the Month');
  });

  it('adapts Form 15 headers to Form X section shape (earnedDuring → earnedEarned)', () => {
    const form15 = resolveForm15Part1LeaveColumnHeaders(headers, groupLabels);
    const section = toFormXSectionHeadersFromForm15(form15);
    expect(section.earnedEarned).toBe(form15.earnedDuring);
    expect(section.medicalBeginning).toBe(form15.medicalBeginning);
  });

  it('maps Earned Leave with Form X formulas', () => {
    const leaveRecord = {
      'Earned Leave': { paidBalance: 23, paidBooked: 1 },
    };
    const values = buildForm15Part1LeaveSectionValues(leaveRecord, '9', {}, 'earned');
    expect(values.beginning).toBe('24');
    expect(values.earnedDuring).toBe('0');
    expect(values.availed).toBe('1');
    expect(values.balance).toBe('23');
  });

  it('maps Medical Leave from Contingency Leave Balance/Booked', () => {
    const leaveRecord = {
      'Contingency Leave': { balance: 26, booked: 2 },
    };
    const values = buildForm15Part1LeaveSectionValues(leaveRecord, '9', {}, 'medical');
    expect(values.beginning).toBe('28');
    expect(values.availed).toBe('2');
    expect(values.balance).toBe('26');
  });

  it('writes Earned + Medical columns onto a Form 15 Part 1 row', () => {
    // Deduped leaf keys (same as Statutory dedupeStatutoryTableHeaders).
    const dedupedHeaders = [
      'S.No',
      'Name of the Worker',
      'Worker Identity No.',
      'Leave at the beginning of the Month',
      'Leave earned during the Period',
      'Leave availed during the Month',
      'Leave balance at the end of the Month',
      'Leave at beginning of the Month',
      'Leave availed during the Month (2)',
      'Leave balance at end of the Month',
    ];
    const form15Headers = resolveForm15Part1LeaveColumnHeaders(dedupedHeaders, groupLabels);
    const row = {
      'Name of the Worker': 'Rajeshkumar',
      'Worker Identity No.': 'VE0447',
    };

    const leaveRecord = {
      employee: 'Rajeshkumar (VE0447)',
      'Earned Leave': { paidBalance: 39, paidBooked: 0 },
      'Contingency Leave': { unpaidBalance: 26, unpaidBooked: 2 },
    };

    const applied = applyForm15Part1TamilNaduLeaveToRow(
      row,
      leaveRecord,
      form15Headers,
      {},
      [],
      {
        employeeIdHeader: 'Worker Identity No.',
        employeeNameHeader: 'Name of the Worker',
        overwrite: true,
      }
    );

    expect(applied).toBeGreaterThan(0);
    expect(row[form15Headers.earnedBeginning]).toBe('39');
    expect(row[form15Headers.earnedDuring]).toBe('0');
    expect(row[form15Headers.earnedAvailed]).toBe('0');
    expect(row[form15Headers.earnedBalance]).toBe('39');
    expect(row[form15Headers.medicalBeginning]).toBe('28');
    expect(row[form15Headers.medicalAvailed]).toBe('2');
    expect(row[form15Headers.medicalBalance]).toBe('26');
  });
});
