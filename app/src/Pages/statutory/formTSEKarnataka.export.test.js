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
});
