import ExcelJS from 'exceljs';
import {
  applyFormXVIAPMusterLayoutFixes,
  clampFormXVIAPHeadersToRemarks,
  FORM_XVI_AP_COL_AK,
  FORM_XVI_AP_COL_AL,
  readFormXXIIAPCellValue,
  resolveFormXVIAPContractorText,
  resolveFormXVIAPLastKeepCol,
  writeFormXVIAPContractorHeader,
  writeFormXVIAPEstablishmentContractHeader,
  detectFormXXIIAPDayColumnMapWithMarker,
} from './formXXIIAPRegisterOfEmployment';
import {
  findFormXVIAPRecordByFullOrFirstLastName,
  findFormXVIAPSiteRecord,
  formXVIAPAttendanceNameKeys,
  formXVIAPFirstLastKey,
  formXVIAPNamesMatch,
  FORM_XVI_AP_DEFAULT_CONTRACTOR_NAME,
  resolveFormXVIAPContractorFromSources,
  resolveFormXVIAPEmployeeName,
} from './formXVIAPMuster';
import { headersIndicateForm25APMusterTable } from './form25APMuster';

describe('Form XVI AP Muster Roll Excel layout helpers', () => {
  it('stops headers at Remarks so overflow columns are not exported', () => {
    expect(
      clampFormXVIAPHeadersToRemarks([
        'S. No',
        'Name of the Employee',
        '1',
        '31',
        'Remarks',
        'Aadhaar',
        'Father leftover'
      ])
    ).toEqual(['S. No', 'Name of the Employee', '1', '31', 'Remarks']);
  });

  it('keeps the table through Remarks / day 31 and excludes AK/AL when days end before AK', () => {
    const dayColumnMap = new Map();
    for (let day = 1; day <= 31; day += 1) dayColumnMap.set(day, 4 + day); // E=5 … AI=35
    const lastKeep = resolveFormXVIAPLastKeepCol(dayColumnMap, 36, [1, 2, 3, 4, 35, 36]);
    expect(lastKeep).toBeLessThan(FORM_XVI_AP_COL_AK);
    expect(lastKeep).toBeLessThan(FORM_XVI_AP_COL_AL);
    expect(lastKeep).toBe(36);
  });

  it('writes Nature and For the Month of as full single-line bands above Dates', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('XVI-Muster Roll');
    // Template only has the "For the" fragment (real AP template behaviour).
    ws.getCell(8, 25).value = 'For the';
    ws.getCell(9, 5).value = 'Dates';
    ws.mergeCells(9, 5, 9, 35);
    for (let day = 1; day <= 31; day += 1) {
      ws.getCell(10, 4 + day).value = String(day);
    }
    ws.mergeCells(11, 5, 11, 35);
    ws.getCell(11, 5).value = 'P';
    ws.getCell(11, FORM_XVI_AP_COL_AK).value = 'ata ramana';
    ws.getCell(11, FORM_XVI_AP_COL_AK).border = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'thin' },
      right: { style: 'thin' }
    };
    ws.getCell(11, FORM_XVI_AP_COL_AL).value = 3.14e17;
    ws.getCell(11, FORM_XVI_AP_COL_AL).border = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'thin' },
      right: { style: 'thin' }
    };

    const dayColumnMap = new Map();
    for (let day = 1; day <= 31; day += 1) dayColumnMap.set(day, 4 + day);

    applyFormXVIAPMusterLayoutFixes(ws, {
      headerRow: 9,
      markerRow: 10,
      dataStartRow: 11,
      dayColumnMap,
      remarksCol: 36,
      lastKeepCol: 36,
      sourceRowCount: 2,
      monthYearText: 'April 2026',
      headerFormData: { form_xvi_nature_location_work: 'AP-Tadipatri' },
      maxScanRows: 20,
      attendanceYear: 2026,
      attendanceMonthIndex: 3
    });

    const merges = ws.model?.merges || [];
    expect(merges.some((m) => /^E9:AI9$/i.test(String(m)))).toBe(true);
    expect(String(ws.getCell(9, 5).value || '')).toMatch(/^dates?$/i);
    expect(merges.some((m) => /^E11:AI11$/i.test(String(m)))).toBe(false);

    expect(String(ws.getCell(8, 1).value || '')).toBe('Nature and Location of work : AP-Tadipatri');
    expect(ws.getCell(8, 1).alignment?.wrapText).toBe(false);
    expect(ws.getCell(8, 1).font?.bold).toBe(true);
    expect(String(ws.getCell(8, 5).value || '')).toBe('For the Month of : April 2026');
    expect(ws.getCell(8, 5).alignment?.wrapText).toBe(false);
    expect(ws.getCell(8, 5).font?.bold).toBe(true);
    // Old "For the" fragment must be cleared.
    expect(String(ws.getCell(8, 25).value || '')).not.toMatch(/^For the$/i);

    expect(ws.getCell(11, FORM_XVI_AP_COL_AK).value).toBeNull();
    expect(ws.getCell(11, FORM_XVI_AP_COL_AL).value).toBeNull();
    expect(ws.getCell(11, FORM_XVI_AP_COL_AK).border).toBeFalsy();
    expect(ws.getCell(11, FORM_XVI_AP_COL_AL).border).toBeFalsy();
  });

  it('strips leftover Form XVI header dump from the For the Month of cell', () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('XVI-Muster Roll');
    ws.getCell(5, 6).value =
      'For the Month of : May 2026 Form XVI - Muster Roll Address of the Establishment : M/S VIBRANT GREENTECH Name and Address of Contractor. : VAYONA ENERGY Name and address of Principal Employer : Vibrant Nature and Location of work : AP-Tadipatri For the Month of : May 2026';
    ws.getCell(6, 1).value = 'Nature and Location of work : AP-Tadipatri';
    ws.getCell(8, 5).value = 'Dates';

    applyFormXVIAPMusterLayoutFixes(ws, {
      headerRow: 8,
      markerRow: 9,
      dataStartRow: 10,
      dayColumnMap: new Map([[1, 5], [31, 35]]),
      remarksCol: -1,
      lastKeepCol: 36,
      sourceRowCount: 0,
      monthYearText: 'May 2026',
      headerFormData: { form_xvi_nature_location_work: 'AP-Tadipatri' },
      maxScanRows: 12
    });

    expect(String(ws.getCell(7, 1).value || '')).toBe('Nature and Location of work : AP-Tadipatri');
    expect(String(ws.getCell(7, 5).value || '')).toBe('For the Month of : May 2026');
    expect(String(ws.getCell(7, 5).value || '')).not.toMatch(/VAYONA|Establishment|Form XVI/i);
  });

  it('writes Name and address of Contractor from site header data', () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('XVI-Muster Roll');
    ws.getCell(5, 1).value = 'Name and address of Contractor :';
    ws.getCell(5, 13).value = 'Name and address of Principal Employer : VAYON';
    ws.getCell(8, 5).value = 'Dates';

    applyFormXVIAPMusterLayoutFixes(ws, {
      headerRow: 9,
      markerRow: 10,
      dataStartRow: 11,
      dayColumnMap: new Map([[1, 5], [31, 35]]),
      remarksCol: 36,
      lastKeepCol: 36,
      sourceRowCount: 0,
      monthYearText: 'July 2026',
      headerFormData: {
        form_xvi_nature_location_work: 'AP-Nimbagallu',
        form_xvi_contractor: 'Vayona Energy Pvt Ltd, Theni'
      },
      maxScanRows: 12
    });

    expect(resolveFormXVIAPContractorText({ form_xvi_contractor: 'Acme Contractors' })).toBe(
      'Acme Contractors'
    );
    expect(resolveFormXVIAPContractorText({})).toBe(FORM_XVI_AP_DEFAULT_CONTRACTOR_NAME);
    expect(String(ws.getCell(5, 1).value || '')).toBe(
      'Name and address of Contractor : Vayona Energy Pvt Ltd, Theni'
    );
    expect(String(ws.getCell(5, 13).value || '')).toMatch(/Principal Employer/i);
  });

  it('prefers Dates 1–31 row over official column-number row', () => {
    const grid = {};
    // Official CLRA column index row (1=S.No …) starts at col A.
    for (let n = 1; n <= 31; n += 1) grid[`8:${n}`] = String(n);
    // Dates 1–31 start at col E (after S.No / Name / Father / Sex).
    for (let n = 1; n <= 31; n += 1) grid[`10:${4 + n}`] = String(n);
    const getCell = (r, c) => grid[`${r}:${c}`] || '';
    const { dayColumnMap, markerRow } = detectFormXXIIAPDayColumnMapWithMarker(getCell, 9, 20, 40);
    expect(markerRow).toBe(10);
    expect(dayColumnMap.get(1)).toBe(5);
    expect(dayColumnMap.get(15)).toBe(19);
  });

  it('writeFormXVIAPContractorHeader fills blank contractor label cell', () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('XVI-Muster Roll');
    ws.getCell(5, 1).value = 'Name and address of Contractor :';
    writeFormXVIAPContractorHeader(ws, {
      headerFormData: { form_xxiii_contractor: 'Site Contractor LLC, Hyderabad' },
      headerScanEnd: 8
    });
    expect(String(ws.getCell(5, 1).value || '')).toBe(
      'Name and address of Contractor : Site Contractor LLC, Hyderabad'
    );
  });

  it('renames Address of the Establishment to establishment in/under which contract is carried on', () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('XVI-Muster Roll');
    ws.getCell(2, 1).value = 'Address of the Establishment :';
    writeFormXVIAPEstablishmentContractHeader(ws, {
      headerFormData: {
        form_xvi_establishment_contract_carried:
          'VAYONA ENERGY PRIVATE LIMITED, Vayona Energy Pvt Ltd, Theni, TamilNadu'
      },
      headerScanEnd: 8
    });
    expect(String(ws.getCell(2, 1).value || '')).toMatch(
      /Name and Address of the Establishment in\/ under which contract is carried on/i
    );
    expect(String(ws.getCell(2, 1).value || '')).toMatch(/VAYONA ENERGY PRIVATE LIMITED/);
    expect(String(ws.getCell(2, 1).value || '')).toMatch(/Theni/);
  });

  it('writeFormXVIAPContractorHeader fills compacted contractor label without spaces', () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('XVI-Muster Roll');
    ws.getCell(5, 1).value = 'NameandaddressofContractor:';
    writeFormXVIAPContractorHeader(ws, {
      headerFormData: { form_xvi_contractor: 'Vayona Energy Pvt Ltd, Theni' },
      headerScanEnd: 8
    });
    expect(String(ws.getCell(5, 1).value || '')).toMatch(/Vayona Energy Pvt Ltd, Theni/);
  });

  it('does not copy later-day attendance into empty early days', () => {
    const row = {
      'Name of the Employee': 'Rabindra Kumar Behera',
      15: 'P',
      16: 'P',
      Dates_21: 'WO'
    };
    expect(readFormXXIIAPCellValue(row, '1')).toBe('');
    expect(readFormXXIIAPCellValue(row, '2')).toBe('');
    expect(readFormXXIIAPCellValue(row, '11')).toBe('');
    expect(readFormXXIIAPCellValue(row, '15')).toBe('P');
    expect(readFormXXIIAPCellValue(row, '21')).toBe('WO');
  });
});

describe('Form XVI AP attendance name matching', () => {
  const jai = { FirstName: 'Jai', LastName: 'Venkata', Name: 'Jai Venkata' };
  const nagaraju = { FirstName: 'Nagaraju', LastName: '', Name: 'Nagaraju' };
  const records = [
    { nameOfTheWorker: 'Jai Venkata' },
    { first_name: 'Rabindra', last_name: 'Kumar' },
    { FirstName: 'Nagaraju', LastName: 'Nageswararao' },
    { employee_name: 'Balasuthan Subramanian' }
  ];

  it('matches by fullname or firstname+lastname', () => {
    expect(formXVIAPNamesMatch(jai, { nameOfTheWorker: 'Jai Venkata' })).toBe(true);
    expect(formXVIAPNamesMatch(jai, { FirstName: 'Jai', LastName: 'Venkata' })).toBe(true);
    expect(formXVIAPNamesMatch(jai, { nameOfTheWorker: 'Jai Kumar' })).toBe(false);
    expect(formXVIAPFirstLastKey(jai)).toBe('jai venkata');
    expect(formXVIAPAttendanceNameKeys(jai)).toEqual(expect.arrayContaining(['jai venkata']));
  });

  it('does not match on first name alone', () => {
    expect(formXVIAPNamesMatch(jai, { FirstName: 'Jai' })).toBe(false);
    expect(
      formXVIAPNamesMatch(
        { FirstName: 'Selva', LastName: 'P' },
        { nameOfTheWorker: 'Selva Kumar' }
      )
    ).toBe(false);
  });

  it('finds attendance/Form 25 rows by fullname or first+last', () => {
    expect(findFormXVIAPRecordByFullOrFirstLastName(jai, records)?.nameOfTheWorker).toBe(
      'Jai Venkata'
    );
    expect(
      findFormXVIAPRecordByFullOrFirstLastName(
        { FirstName: 'Rabindra', LastName: 'Kumar' },
        records
      )?.first_name
    ).toBe('Rabindra');
    expect(
      findFormXVIAPRecordByFullOrFirstLastName(
        { Name: 'Nagaraju' },
        [{ nameOfTheWorker: 'Nagaraju' }, ...records]
      )?.nameOfTheWorker
    ).toBe('Nagaraju');
    expect(findFormXVIAPRecordByFullOrFirstLastName(nagaraju, records)).toBeNull();
  });

  it('Name of the Employee uses fullname or firstname + lastname', () => {
    expect(resolveFormXVIAPEmployeeName(jai)).toBe('Jai Venkata');
    expect(
      resolveFormXVIAPEmployeeName({ FirstName: 'Rabindra', LastName: 'Kumar' })
    ).toBe('Rabindra Kumar');
    expect(resolveFormXVIAPEmployeeName({ Name: 'Kalaiselvan' })).toBe('Kalaiselvan');
  });

  it('does not treat Form XVI identity+day headers as Form 25 AP muster', () => {
    const xviHeaders = [
      'S. No',
      'Name of the Employee',
      "Father's / Husband's Name",
      'Sex',
      ...Array.from({ length: 31 }, (_, i) => String(i + 1)),
      'Remarks'
    ];
    expect(headersIndicateForm25APMusterTable(xviHeaders)).toBe(false);
  });
});

describe('Form XVI AP contractor from Site Management', () => {
  const sites = [
    {
      siteName: 'Nimbagallu Plant',
      location: 'AP-Nimbagallu',
      siteState: 'Andhra Pradesh',
      contractorName: 'Site Contractor Pvt Ltd',
      contractorAddress: 'Hyderabad',
      contractorCity: 'Hyderabad',
      contractorState: 'Telangana'
    }
  ];

  it('finds the site by Location when the nature line is AP-Nimbagallu', () => {
    expect(
      findFormXVIAPSiteRecord(sites, { natureText: 'AP-Nimbagallu' })?.contractorName
    ).toBe('Site Contractor Pvt Ltd');
  });

  it('builds contractor name and address from the matched site', () => {
    expect(
      resolveFormXVIAPContractorFromSources({
        sites,
        natureText: 'AP-Nimbagallu'
      })
    ).toBe('Site Contractor Pvt Ltd, Hyderabad, Hyderabad, Telangana');
  });

  it('defaults contractor to VAYONA ENERGY PRIVATE LIMITED', () => {
    expect(resolveFormXVIAPContractorFromSources({})).toBe(FORM_XVI_AP_DEFAULT_CONTRACTOR_NAME);
    expect(
      resolveFormXVIAPContractorFromSources({
        sites: [{ siteName: 'AP-Nimbagallu', location: 'AP-Nimbagallu' }],
        natureText: 'AP-Nimbagallu'
      })
    ).toBe(FORM_XVI_AP_DEFAULT_CONTRACTOR_NAME);
  });

  it('falls back to employee Contractor_Name when site contractor is blank', () => {
    expect(
      resolveFormXVIAPContractorFromSources({
        sites: [{ siteName: 'AP-Nimbagallu', location: 'AP-Nimbagallu' }],
        siteName: 'AP-Nimbagallu',
        employees: [
          { Contractor_Name: 'CLRA Contractor LLC' },
          { Contractor_Name: 'CLRA Contractor LLC' }
        ]
      })
    ).toBe('CLRA Contractor LLC');
  });
});
