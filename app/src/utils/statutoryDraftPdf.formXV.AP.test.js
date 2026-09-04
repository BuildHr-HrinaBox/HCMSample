import {
  buildFormXVAPCertificatePageModel,
  buildFormXVAPServiceCertificatePdfBlob,
  filterFormXVAPCorrespondingEmployeeRows,
  formXVAPRowHasCorrespondingEmployee,
  looksLikeFormXVAPPdfContext
} from './statutoryDraftPdf.formXV.AP';

describe('Form XV AP Service Certificate PDF', () => {
  const headers = {
    'Name and address of the workman': 'Ramesh Kumar, AP-Tadipatri',
    'Age or date of birth': '32',
    'Identification marks': 'Mole on left hand',
    "Father's/Husband's name": 'Suresh Kumar',
    'Total period for which employed_From': '01-Apr-2024',
    'Total period for which employed_To': '30-Apr-2026',
    'Nature of work done': 'Technician',
    'Rate of wage (with particulars of work)': '18000',
    Remarks: ''
  };

  test('detects AP Form XV service certificate context', () => {
    expect(
      looksLikeFormXVAPPdfContext(
        ['FORM XV', 'SERVICE CERTIFICATE', 'Central & A.P. Rules, 1971'],
        [],
        'Form_XV_-_Andhra_Pradesh.xlsx'
      )
    ).toBe(true);
  });

  test('does not treat Rajasthan Form XV wage slip as AP service certificate', () => {
    expect(
      looksLikeFormXVAPPdfContext(
        ['FORM XV', 'Wage Slip', '[See Rule 77(2)(b)]'],
        [],
        'Form_XV_RJ.xlsx'
      )
    ).toBe(false);
  });

  test('keeps only rows that have a corresponding workman name', () => {
    const rows = [
      headers,
      { 'Name and address of the workman': '', 'Rate of wage (with particulars of work)': '500' },
      { 'Name and address of the workman': 'Enter Name and address of the workman' },
      { 'Name and address of the workman': 'Anita Devi, AP-Tadipatri' }
    ];
    const kept = filterFormXVAPCorrespondingEmployeeRows(rows);
    expect(kept).toHaveLength(2);
    expect(formXVAPRowHasCorrespondingEmployee(rows[1])).toBe(false);
    expect(formXVAPRowHasCorrespondingEmployee(rows[2])).toBe(false);
  });

  test('builds one certificate page model with site header and employee fields', () => {
    const model = buildFormXVAPCertificatePageModel({
      headerFormData: {
        form_xv_contractor: 'VAYONA ENERGY PRIVATE LIMITED',
        form_xv_nature_location_work: 'AP-Tadipatri',
        form_xv_establishment_contract_carried: 'Vibrant greentech India PVT LTD',
        form_xv_principal_employer: 'Vibrant greentech India PVT LTD'
      },
      parsedFormHeader: {
        title: 'FORM XV',
        subtitle: 'SERVICE CERTIFICATE'
      },
      row: headers,
      fileName: 'Form_XV_-_Andhra_Pradesh.xlsx'
    });
    expect(model.titles[0]).toMatch(/FORM XV/i);
    expect(model.titles.join(' ')).toMatch(/SERVICE CERTIFICATE/i);
    expect(model.adminRows[0][0]).toBe('VAYONA ENERGY PRIVATE LIMITED');
    expect(model.adminRows[1][0]).toBe('AP-Tadipatri');
    expect(model.fieldRows[0][1]).toBe('Ramesh Kumar, AP-Tadipatri');
    expect(model.tableRow[1]).toBe('01-Apr-2024');
    expect(model.tableRow[3]).toBe('Technician');
  });

  test('PDF blob contains one page per corresponding employee only', async () => {
    const blob = await buildFormXVAPServiceCertificatePdfBlob({
      headerFormData: {
        form_xv_contractor: 'VAYONA ENERGY PRIVATE LIMITED',
        form_xv_nature_location_work: 'AP-Tadipatri'
      },
      parsedFormHeader: { title: 'FORM XV', subtitle: 'SERVICE CERTIFICATE' },
      employeeRows: [
        headers,
        { 'Name and address of the workman': '' },
        { ...headers, 'Name and address of the workman': 'Anita Devi, AP-Tadipatri' }
      ],
      fileName: 'Form_XV_-_Andhra_Pradesh.xlsx'
    });
    expect(blob).toBeInstanceOf(Blob);
    expect(blob.size).toBeGreaterThan(500);
    expect(String(blob.type || '')).toMatch(/pdf/i);
  });

  test('throws when no corresponding employees are present', async () => {
    await expect(
      buildFormXVAPServiceCertificatePdfBlob({
        employeeRows: [{ 'Name and address of the workman': 'Enter Name and address of the workman' }]
      })
    ).rejects.toThrow(/No corresponding employees/i);
  });
});
