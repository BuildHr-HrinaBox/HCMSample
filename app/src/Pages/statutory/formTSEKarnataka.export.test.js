import {
  buildFormTSERowsFromEmployees,
  formTSEDownloadHasSubstantiveRows,
  resolveFormTSERowsForExport,
  rebuildFormTSETableHeadersFromSheet,
} from './formTSEKarnataka';

describe('Form T Karnataka export row resolution', () => {
  const headers = [
    'S.NO',
    'Name of Employee',
    'Father / Husband\'s Name',
    'Gender',
    'Designation / Department',
  ];

  it('detects substantive employee name rows', () => {
    expect(
      formTSEDownloadHasSubstantiveRows(
        [{ 'S.NO': 1, 'Name of Employee': 'Issac Kanagaraj', Gender: 'Male' }],
        headers
      )
    ).toBe(true);
    expect(
      formTSEDownloadHasSubstantiveRows(
        [{ 'S.NO': 1, 'Name of Employee': '2', Gender: '3' }],
        headers
      )
    ).toBe(false);
    expect(
      formTSEDownloadHasSubstantiveRows(
        [{ 'S.NO': 'S.NO', 'Name of Employee': 'Name of Employee', Gender: 'Gender' }],
        headers
      )
    ).toBe(false);
  });

  it('rebuilds export rows from People employees when live grid is empty', () => {
    const rows = resolveFormTSERowsForExport({
      liveRows: [],
      headers,
      employees: [
        { FirstName: 'Issac', LastName: 'Kanagaraj', Sex: 'Male', Designation: 'Assistant Manager' },
        { FirstName: 'Saravanan', Sex: 'Male', Designation: 'Engineer' },
      ],
    });
    expect(formTSEDownloadHasSubstantiveRows(rows, headers)).toBe(true);
    expect(rows[0]['Name of Employee']).toMatch(/Issac/i);
    expect(rows[1]['Name of Employee']).toMatch(/Saravanan/i);
  });

  it('keeps live modal rows when they already have names', () => {
    const live = [
      { 'S.NO': 1, 'Name of Employee': 'Ameerkhan', Gender: 'Male', 'Designation / Department': 'Engineer' },
    ];
    const rows = resolveFormTSERowsForExport({
      liveRows: live,
      headers,
      employees: [{ FirstName: 'Someone', LastName: 'Else' }],
    });
    expect(rows).toHaveLength(1);
    expect(rows[0]['Name of Employee']).toBe('Ameerkhan');
  });

  it('fills the 6th name from People when Autofill shows lookup name but Name cell is abc', () => {
    const live = [
      { 'S.NO': 1, 'Name of Employee': 'Vinay Kumar', "Father / Husband's Name": 'Kadesh j kamble', Gender: 'Male' },
      { 'S.NO': 2, 'Name of Employee': 'Issac Kanagaraj', "Father / Husband's Name": 'abc', Gender: 'Male' },
      { 'S.NO': 3, 'Name of Employee': 'Saravanan', "Father / Husband's Name": 'abc', Gender: 'Male' },
      { 'S.NO': 4, 'Name of Employee': 'Ameerkhan', "Father / Husband's Name": 'abc', Gender: 'Male' },
      { 'S.NO': 5, 'Name of Employee': 'Rajesh', "Father / Husband's Name": 'abc', Gender: 'Male' },
      {
        'S.NO': 6,
        'Name of Employee': 'abc',
        "Father / Husband's Name": '',
        Gender: 'Male',
        'Designation / Department': 'Junior Engineer',
        __employeeLookupName: 'Ashok',
      },
    ];
    const rows = resolveFormTSERowsForExport({
      liveRows: live,
      headers,
      employees: [
        { FirstName: 'Vinay', LastName: 'Kumar' },
        { FirstName: 'Issac', LastName: 'Kanagaraj' },
        { FirstName: 'Saravanan' },
        { FirstName: 'Ameerkhan' },
        { FirstName: 'Rajesh' },
        { FirstName: 'Ashok', Designation: 'Junior Engineer' },
      ],
    });
    expect(rows[5]['Name of Employee']).toMatch(/Ashok/i);
    expect(rows[5]['Name of Employee']).not.toMatch(/^abc$/i);
  });

  it('buildFormTSERowsFromEmployees fills serial and name', () => {
    const rows = buildFormTSERowsFromEmployees(
      [{ FirstName: 'Rajesh', Sex: 'Male' }],
      headers
    );
    expect(String(rows[0]['S.NO'])).toBe('1');
    expect(rows[0]['Name of Employee']).toMatch(/Rajesh/i);
  });
});

describe('Form T Karnataka original template column layout', () => {
  it('does not treat column-index row as employee data when parser points at it', () => {
    // Mimic Form T: R12 headers, R13 day numbers, R14 statutory 1..n, R15+ data.
    // 0-based: header=0, days=1, index=2, data=3
    const jsonData = [];
    for (let r = 0; r < 8; r += 1) jsonData.push(Array(45).fill(''));

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
    ];
    headers.forEach((h, c) => {
      jsonData[0][c] = h;
      jsonData[1][c] = h; // vertically repeated identity labels
    });
    for (let d = 1; d <= 31; d += 1) {
      jsonData[0][8 + d] =
        'ATTENDANCE (Please mention the date of suspension of employees, if any)';
      jsonData[1][8 + d] = String(d);
    }
    for (let c = 0; c < 40; c += 1) {
      // Template quirk: attendance band may repeat statutory col "10"
      jsonData[2][c] = String(c < 10 ? c + 1 : c === 10 ? 10 : c);
    }

    const get = (r, c) => (jsonData[r] && jsonData[r][c] != null ? String(jsonData[r][c]) : '');

    const rebuilt = rebuildFormTSETableHeadersFromSheet({
      headerRowIndex: 0,
      getMergedAwareCellText: get,
      getRawCellText: get,
      jsonData,
      effectiveSheetCols: 45,
      tableStartCol: 9, // wrong modal hint (would shift under ATTENDANCE)
      dataStartIndex: 2, // wrong modal hint (column-index row)
    });

    expect(rebuilt).toBeTruthy();
    expect(rebuilt.tableStartCol).toBe(0);
    expect(rebuilt.columnIndexRow).toBe(2);
    expect(rebuilt.dataStartIndex).toBe(3);
    expect(rebuilt.headers.some((h) => /^ATTENDANCE_1$/i.test(h))).toBe(true);
  });

  it('keeps wage columns after Medical Allowance even when body cells are blank', () => {
    // Identity A–I + attendance J… + wage band. Leaf labels after Medical exist, but
    // empty body + empty gap must not truncate the header scan (early-stop bug).
    const cols = 70;
    const jsonData = [];
    for (let r = 0; r < 6; r += 1) jsonData.push(Array(cols).fill(''));

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
    identity.forEach((h, c) => {
      jsonData[0][c] = h;
    });
    for (let d = 1; d <= 31; d += 1) {
      jsonData[0][8 + d] = 'ATTENDANCE';
      jsonData[1][8 + d] = String(d);
      jsonData[2][8 + d] = String(d); // calendar / index mix under attendance
    }
    // Wage leaf row (R1) + statutory index (R2) after attendance (col 40+)
    const wageStart = 40;
    const wageLeaves = [
      'No. of payable days',
      'Total OT hours',
      'Basic',
      'DA/VDA',
      'HRA',
      'Conveyance',
      'Medical Allowance',
      '', // gap under parent merge (no leaf text)
      'Special allowance',
      'OT',
      'NFH',
      'Maternity Benefit',
      'Others',
      'Subsistence Allowance',
      'Total',
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
      'Net Amount Payable',
      'Mode of Payment Cash/ Cheque No.',
    ];
    wageLeaves.forEach((label, i) => {
      const c = wageStart + i;
      jsonData[0][c] = 'Earned wages and other allowances'; // parent banner
      if (i >= 15) jsonData[0][c] = 'Deductions';
      if (i >= 26) jsonData[0][c] = label;
      jsonData[1][c] = label;
      jsonData[2][c] = String(11 + i);
    });
    // Body blank — no sample data after Medical.
    jsonData[3][0] = '1';
    jsonData[3][1] = 'Test Employee';

    const get = (r, c) => (jsonData[r] && jsonData[r][c] != null ? String(jsonData[r][c]) : '');
    // Merged-aware: parent banner fills empty leaf cells (simulates Excel merge read).
    const getMerged = (r, c) => {
      const direct = get(r, c);
      if (direct) return direct;
      if (r === 1 && c >= wageStart && c < wageStart + 15) {
        return 'Earned wages and other allowances';
      }
      return '';
    };

    const rebuilt = rebuildFormTSETableHeadersFromSheet({
      headerRowIndex: 0,
      getMergedAwareCellText: getMerged,
      getRawCellText: get,
      jsonData,
      effectiveSheetCols: 45, // intentionally short — must still scan past Medical
      tableStartCol: 0,
      dataStartIndex: 3,
    });

    expect(rebuilt).toBeTruthy();
    const labels = (rebuilt.headers || []).map((h) =>
      String(h || '')
        .replace(/\s*\(\s*\d{1,2}\s*\)\s*$/, '')
        .trim()
        .toLowerCase()
    );
    expect(labels.some((h) => h.includes('medical'))).toBe(true);
    expect(labels.some((h) => h.includes('special'))).toBe(true);
    expect(labels.some((h) => h === 'ot' || h.includes('nfh'))).toBe(true);
    expect(labels.some((h) => h.includes('net') && h.includes('payable'))).toBe(true);
    expect(labels.some((h) => h.includes('mode') && h.includes('payment'))).toBe(true);
  });
});
