import {
  applyForm15Part1TamilNaduLeaveAutofillWithSummary,
  applyForm15Part1TamilNaduLeaveToRow,
  buildForm15Part1LeaveExportColumnValues,
  buildForm15Part1LeaveSectionValues,
  ensureForm15Part1LeaveColumnsLikeFormX,
  formatForm15Part1LeaveTransferSummary,
  mapForm15Part1LeaveValuesToExcelLeaveColumns,
  matchForm15Part1LeaveRecord,
  resolveForm15Part1LeaveColumnHeaders,
  toFormXSectionHeadersFromForm15,
} from './form15TamilNaduLeave';
import { buildLeaveRecordLookupMap } from '../../utils/leaveMetrics';

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

  it('inserts Medical + Other Leave bands when Form 15 only has Earned Leave', () => {
    const earnedOnlyHeaders = [
      'Name of the Worker',
      'Worker Identity No.',
      'Leave at the beginning of the Month',
      'Leave earned during the Period',
      'Leave availed during the Month',
      'Leave balance at the end of the Month',
    ];
    const earnedOnlyGroups = [
      '',
      '',
      'Earned Leave',
      'Earned Leave',
      'Earned Leave',
      'Earned Leave',
    ];
    const rows = [
      {
        'Name of the Worker': 'Rajeshkumar Ramasamy',
        'Worker Identity No.': 'VE0447',
        'Leave at the beginning of the Month': '0',
      },
    ];

    const ensured = ensureForm15Part1LeaveColumnsLikeFormX(
      earnedOnlyHeaders,
      earnedOnlyGroups,
      rows
    );

    expect(ensured.changed).toBe(true);
    expect(ensured.insertedMedical).toBe(true);
    expect(ensured.insertedOther).toBe(true);
    expect(ensured.groupLabels).toContain('Medical Leave');
    expect(ensured.groupLabels).toContain('Other Leave');

    const resolved = resolveForm15Part1LeaveColumnHeaders(
      ensured.headers,
      ensured.groupLabels
    );
    expect(resolved.medicalBeginning).toBeTruthy();
    expect(resolved.medicalAvailed).toBeTruthy();
    expect(resolved.medicalBalance).toBeTruthy();
    expect(resolved.otherBeginning).toBeTruthy();
    expect(resolved.otherAvailed).toBeTruthy();
    expect(resolved.otherBalance).toBeTruthy();
    // Existing earned data preserved
    expect(ensured.rows[0]['Leave at the beginning of the Month']).toBe('0');
    expect(ensured.rows[0]['Name of the Worker']).toBe('Rajeshkumar Ramasamy');
  });

  it('does not re-insert when Medical + Other already exist with unique keys', () => {
    const fullHeaders = [
      'Name of the Worker',
      'Worker Identity No.',
      'Leave at the beginning of the Month',
      'Leave earned during the Period',
      'Leave availed during the Month',
      'Leave balance at the end of the Month',
      'Leave at beginning of the Month',
      'Leave availed during the Month (1)',
      'Leave balance at end of the Month',
      'Leave at beginning of the Month (1)',
      'Leave availed during the Month (2)',
      'Leave Balance at end of the Month (1)',
    ];
    const fullGroups = [
      '',
      '',
      'Earned Leave',
      'Earned Leave',
      'Earned Leave',
      'Earned Leave',
      'Medical Leave',
      'Medical Leave',
      'Medical Leave',
      'Other Leave',
      'Other Leave',
      'Other Leave',
    ];
    const ensured = ensureForm15Part1LeaveColumnsLikeFormX(fullHeaders, fullGroups, []);
    expect(ensured.changed).toBe(false);
    expect(ensured.insertedMedical).toBe(false);
    expect(ensured.insertedOther).toBe(false);
  });

  it('dedupes repeated leave leaf headers when Medical already exists', () => {
    const fullHeaders = [
      'Name of the Worker',
      'Worker Identity No.',
      'Leave at the beginning of the Month',
      'Leave earned during the Period',
      'Leave availed during the Month',
      'Leave balance at the end of the Month',
      'Leave at beginning of the Month',
      'Leave availed during the Month',
      'Leave balance at end of the Month',
      'Leave at beginning of the Month (1)',
      'Leave availed during the Month (2)',
      'Leave Balance at end of the Month (1)',
    ];
    const fullGroups = [
      '',
      '',
      'Earned Leave',
      'Earned Leave',
      'Earned Leave',
      'Earned Leave',
      'Medical Leave',
      'Medical Leave',
      'Medical Leave',
      'Other Leave',
      'Other Leave',
      'Other Leave',
    ];
    const ensured = ensureForm15Part1LeaveColumnsLikeFormX(fullHeaders, fullGroups, [
      {
        'Name of the Worker': 'A',
        'Leave availed during the Month': '5',
      },
    ]);
    expect(ensured.changed).toBe(true);
    expect(ensured.insertedMedical).toBe(false);
    expect(ensured.insertedOther).toBe(false);
    expect(ensured.headers.filter((h) => /^Leave availed during the Month/i.test(h))).toEqual([
      'Leave availed during the Month',
      'Leave availed during the Month (2)',
      'Leave availed during the Month (2)',
    ]);
    expect(ensured.rows[0]['Leave availed during the Month']).toBe('5');
  });

  it('matches by Worker Identity No. and reports not-found / ambiguous', () => {
    const leaveRecords = [
      {
        employee: 'Rajeshkumar Ramasamy (VE0447)',
        EmployeeID: 'VE0447',
        'Earned Leave': { paidBalance: 39, paidBooked: 0 },
        'Contingency Leave': { unpaidBalance: 26, unpaidBooked: 2 },
        'Legacy Earned Leave': 0,
      },
      {
        employee: 'Duplicate Name (VE0999)',
        EmployeeID: 'VE0999',
        'Earned Leave': { paidBalance: 1, paidBooked: 0 },
      },
      {
        employee: 'Duplicate Name (VE0888)',
        EmployeeID: 'VE0888',
        'Earned Leave': { paidBalance: 2, paidBooked: 0 },
      },
    ];

    const earnedOnlyHeaders = [
      'Name of the Worker',
      'Worker Identity No.',
      'Leave at the beginning of the Month',
      'Leave earned during the Period',
      'Leave availed during the Month',
      'Leave balance at the end of the Month',
    ];
    const earnedOnlyGroups = [
      '',
      '',
      'Earned Leave',
      'Earned Leave',
      'Earned Leave',
      'Earned Leave',
    ];
    const rows = [
      {
        'Name of the Worker': 'Rajeshkumar Ramasamy',
        'Worker Identity No.': 'VE0447',
      },
      {
        'Name of the Worker': 'Missing Person',
        'Worker Identity No.': 'VE0000',
      },
      {
        'Name of the Worker': 'Duplicate Name',
        'Worker Identity No.': '',
      },
    ];

    const result = applyForm15Part1TamilNaduLeaveAutofillWithSummary(
      rows,
      [],
      leaveRecords,
      {},
      earnedOnlyHeaders,
      null,
      [],
      {
        overwrite: true,
        groupLabels: earnedOnlyGroups,
        resolveEmployeeHeaders: () => ({
          employeeNameHeader: 'Name of the Worker',
          employeeIdHeader: 'Worker Identity No.',
        }),
      }
    );

    expect(result.summary.matched).toBe(1);
    expect(result.summary.updated).toBe(1);
    expect(result.summary.notFound).toHaveLength(1);
    expect(result.summary.notFound[0].employeeId).toBe('VE0000');
    expect(result.summary.ambiguous).toHaveLength(1);
    expect(result.summary.ambiguous[0].employeeName).toBe('Duplicate Name');
    expect(result.summary.insertedMedical).toBe(true);
    expect(result.summary.insertedOther).toBe(true);

    const resolved = resolveForm15Part1LeaveColumnHeaders(
      result.headers,
      result.groupLabels
    );
    expect(result.rows[0][resolved.earnedBeginning]).toBe('39');
    expect(result.rows[0][resolved.earnedDuring]).toBe('0');
    expect(result.rows[0][resolved.medicalBeginning]).toBe('28');
    // Not found / ambiguous rows stay blank on leave columns
    expect(String(result.rows[1][resolved.earnedBeginning] ?? '').trim()).toBe('');
    expect(String(result.rows[2][resolved.earnedBeginning] ?? '').trim()).toBe('');

    const line = formatForm15Part1LeaveTransferSummary(result.summary);
    expect(line).toContain('Matched 1');
    expect(line).toContain('updated 1');
    expect(line).toContain('not found 1');
    expect(line).toContain('ambiguous 1');
  });

  it('prefers Employee ID over duplicate names', () => {
    const leaveRecords = [
      {
        employee: 'Duplicate Name (VE0999)',
        EmployeeID: 'VE0999',
        'Earned Leave': { paidBalance: 10, paidBooked: 0 },
      },
      {
        employee: 'Duplicate Name (VE0888)',
        EmployeeID: 'VE0888',
        'Earned Leave': { paidBalance: 20, paidBooked: 0 },
      },
    ];
    const lookup = buildLeaveRecordLookupMap(leaveRecords);
    const nameDupCounts = new Map([
      ['duplicate name', 2],
      ['duplicate name (ve0999)', 1],
      ['duplicate name (ve0888)', 1],
    ]);
    const match = matchForm15Part1LeaveRecord(
      lookup,
      { 'Name of the Worker': 'Duplicate Name', 'Worker Identity No.': 'VE0888' },
      'Worker Identity No.',
      'Name of the Worker',
      nameDupCounts
    );
    expect(match.ambiguous).toBe(false);
    expect(match.record.EmployeeID).toBe('VE0888');
    expect(match.reason).toBe('matched_id');
  });

  it('builds leave export values in header order including earned-during 0', () => {
    const exportHeaders = [
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
    const row = {
      'Name of the Worker': 'Rajeshkumar Ramasamy',
      'Worker Identity No.': 'VE0447',
      'Leave at the beginning of the Month': '39',
      'Leave earned during the Period': '0',
      'Leave availed during the Month': '11',
      'Leave balance at the end of the Month': '28',
      'Leave at beginning of the Month': '28',
      'Leave availed during the Month (2)': '2',
      'Leave balance at end of the Month': '26',
    };
    const values = buildForm15Part1LeaveExportColumnValues(row, exportHeaders);
    expect(values).toEqual(['39', '0', '11', '28', '28', '2', '26']);
    expect(mapForm15Part1LeaveValuesToExcelLeaveColumns(row, exportHeaders, 7)).toEqual([
      '39',
      '0',
      '11',
      '28',
      '28',
      '2',
      '26',
    ]);
  });
});
