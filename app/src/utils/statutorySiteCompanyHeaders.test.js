import {
  isFormXLeaveGratuityHeader,
  isFormXMonthOnlyHeaderLabel,
  isFormXYearOnlyHeaderLabel,
  isFormVTamilNaduSkipAutofillHeader,
  resolveHeaderFieldExportValue,
  resolveCompanyRecordForStatutory,
  looksLikeDemoCompanyHeaderValue,
  applySiteCompanyHeaderAutofill,
  buildCompanyNameAndAddress,
  buildCompanyNameWithSiteAddress,
  buildStatutoryEmployerTextFromCompanies,
  writeFormUEstablishmentNameAddressToWorksheet,
} from './statutorySiteCompanyHeaders';

describe('statutory header export helpers', () => {
  it('resolves month/year header values across the Tamil Nadu Form X aliases', () => {
    const headerFormData = { form_t_month_year: 'June 2024' };
    const field = { key: 'form_xviii_month_year', label: 'Month/Year' };

    expect(resolveHeaderFieldExportValue(headerFormData, field)).toBe('June 2024');
  });

  it('resolves separate Form X Month and Year header values', () => {
    const headerFormData = { form_x_month: 'June', form_x_year: '2024' };

    expect(
      resolveHeaderFieldExportValue(headerFormData, { key: 'form_x_month', label: 'Month:' })
    ).toBe('June');
    expect(
      resolveHeaderFieldExportValue(headerFormData, { key: 'form_x_year', label: 'Year:' })
    ).toBe('2024');
  });

  it('recognizes Form X month-only and year-only labels', () => {
    expect(isFormXMonthOnlyHeaderLabel('Month:')).toBe(true);
    expect(isFormXMonthOnlyHeaderLabel('Month / Year')).toBe(false);
    expect(isFormXYearOnlyHeaderLabel('Year:')).toBe(true);
    expect(isFormXYearOnlyHeaderLabel('Calendar Year:')).toBe(false);
  });

  it('recognizes Form X gratuity header labels', () => {
    expect(isFormXLeaveGratuityHeader('Amount paid as Gratuity in case of exit of the employee')).toBe(true);
    expect(isFormXLeaveGratuityHeader('Gratuity benefits')).toBe(true);
    expect(isFormXLeaveGratuityHeader('Name of the employee')).toBe(false);
    expect(isFormXLeaveGratuityHeader('Employee Identification No.')).toBe(false);
  });

  it('recognizes Form V TN skip-autofill headers', () => {
    expect(
      isFormVTamilNaduSkipAutofillHeader('Benefit availed for working on National Holiday (**)')
    ).toBe(true);
    expect(
      isFormVTamilNaduSkipAutofillHeader('Benefit availed for working on Festival Holiday (**)')
    ).toBe(true);
    expect(isFormVTamilNaduSkipAutofillHeader('Remarks')).toBe(true);
    expect(isFormVTamilNaduSkipAutofillHeader('Name of the Employee')).toBe(false);
    expect(isFormVTamilNaduSkipAutofillHeader('Total Days Worked')).toBe(false);
    expect(isFormVTamilNaduSkipAutofillHeader('Approved Festival Holidays:')).toBe(false);
  });
});

describe('resolveCompanyRecordForStatutory', () => {
  const companies = [
    {
      companyName: 'Company Name 2',
      companyAddress: '123, chamiers Road',
      city: 'Chennai',
      state: 'Tamil Nadu',
      postalcode: '600028',
      ROWID: 'co2',
    },
    {
      companyName: 'VAYONA ENERGY PRIVATE LIMITED',
      companyAddress: 'Arumbakkam',
      city: 'Chennai',
      state: 'Tamil Nadu',
      postalcode: '600106',
      ROWID: 'vayona',
    },
  ];

  it('does not fall back to first company when multiple real companies exist', () => {
    const realCompanies = [
      companies[1],
      {
        companyName: 'OTHER ENERGY LTD',
        companyAddress: 'T Nagar',
        city: 'Chennai',
        state: 'Tamil Nadu',
        postalcode: '600017',
        ROWID: 'other',
      },
    ];
    expect(resolveCompanyRecordForStatutory({}, realCompanies)).toBeNull();
  });

  it('skips demo Company Name 2 and uses the sole real company_function record', () => {
    expect(resolveCompanyRecordForStatutory({}, companies)?.companyName).toBe(
      'VAYONA ENERGY PRIVATE LIMITED'
    );
  });

  it('prefers Site Management company link over list order', () => {
    const site = { companyId: 'vayona', companyName: 'VAYONA ENERGY PRIVATE LIMITED' };
    const resolved = resolveCompanyRecordForStatutory({}, companies, site);
    expect(resolved?.companyName).toBe('VAYONA ENERGY PRIVATE LIMITED');
  });

  it('matches partial site company name to full company record', () => {
    const site = { companyName: 'Vayona' };
    const resolved = resolveCompanyRecordForStatutory({}, companies, site);
    expect(resolved?.companyName).toBe('VAYONA ENERGY PRIVATE LIMITED');
  });

  it('uses sole company when only one exists', () => {
    expect(resolveCompanyRecordForStatutory({}, [companies[1]])?.companyName).toBe(
      'VAYONA ENERGY PRIVATE LIMITED'
    );
  });

  it('detects demo company header leftovers', () => {
    expect(
      looksLikeDemoCompanyHeaderValue('Company Name 2, 123, chamiers Road, Chennai, Tamil Nadu, 600028')
    ).toBe(true);
    expect(
      looksLikeDemoCompanyHeaderValue('Delphi, Arumbakkam, Chennai, Tamil Nadu, 600106')
    ).toBe(true);
    expect(
      looksLikeDemoCompanyHeaderValue(
        'VAYONA ENERGY PRIVATE LIMITED, Arumbakkam, Chennai, Tamil Nadu, 600106'
      )
    ).toBe(false);
  });

  it('ignores Delphi site company link and uses sole company_function record', () => {
    const sole = [companies[1]];
    const site = { companyName: 'Delphi', siteName: 'Vayona' };
    const resolved = resolveCompanyRecordForStatutory({}, sole, site);
    expect(resolved?.companyName).toBe('VAYONA ENERGY PRIVATE LIMITED');
  });

  it('replaces Delphi template employer with company_function on autofill', () => {
    const company = companies[1];
    const next = applySiteCompanyHeaderAutofill(
      {
        statutory_employer_name_address: 'Delphi, Arumbakkam, Chennai, Tamil Nadu, 600106',
      },
      {
        site: { siteName: 'Vayona' },
        company,
        formHeaderFields: [
          { label: 'Name and Address of the Employer:', key: 'statutory_employer_name_address' },
        ],
      }
    );
    expect(next.statutory_employer_name_address).toBe(buildCompanyNameAndAddress(company));
  });

  it('replaces demo employer text with site-linked company on autofill', () => {
    const site = {
      siteName: 'Vayona',
      companyName: 'VAYONA ENERGY PRIVATE LIMITED',
      companyId: 'vayona',
    };
    const company = resolveCompanyRecordForStatutory({}, companies, site);
    const next = applySiteCompanyHeaderAutofill(
      {
        statutory_employer_name_address:
          'Company Name 2, 123, chamiers Road, Chennai, Tamil Nadu, 600028',
      },
      {
        site,
        company,
        formHeaderFields: [
          { label: 'Name and Address of the Employer:', key: 'statutory_employer_name_address' },
        ],
      }
    );
    expect(next.statutory_employer_name_address).toBe(buildCompanyNameAndAddress(companies[1]));
  });

  it('always overwrites employer from company_function (never leaves site/stale text)', () => {
    const company = companies[1];
    const companyText = buildCompanyNameAndAddress(company);
    const site = {
      siteName: 'Theni Plant',
      siteAddress: 'Site Road',
      siteCity: 'Theni',
      siteState: 'Tamil Nadu',
      companyId: 'vayona',
    };
    const next = applySiteCompanyHeaderAutofill(
      {
        statutory_employer_name_address: 'Theni Plant, Site Road, Theni, Tamil Nadu',
        form_t_employer: 'Some other site address',
      },
      {
        site,
        company,
        formHeaderFields: [
          { label: 'Name and Address of the Employer:', key: 'statutory_employer_name_address' },
          { label: 'Name and Address of employer', key: 'form_t_employer' },
        ],
      }
    );
    expect(next.statutory_employer_name_address).toBe(companyText);
    expect(next.form_t_employer).toBe(companyText);
  });

  it('fills employer from sole company even without a site record', () => {
    const company = companies[1];
    const next = applySiteCompanyHeaderAutofill(
      { statutory_employer_name_address: '' },
      {
        site: null,
        company,
        formHeaderFields: [
          { label: 'Name and Address of the Employer:', key: 'statutory_employer_name_address' },
        ],
      }
    );
    expect(next.statutory_employer_name_address).toBe(buildCompanyNameAndAddress(company));
  });

  it('builds employer text from sole company_function record', () => {
    const text = buildStatutoryEmployerTextFromCompanies([companies[1]]);
    expect(text).toBe(buildCompanyNameAndAddress(companies[1]));
  });

  it('overwrites Manager/Incharge from Site Management incharge Name', () => {
    const site = {
      siteName: 'Vayona',
      inchargeName: 'Sample',
    };
    const next = applySiteCompanyHeaderAutofill(
      { form_header_manager_incharge: 'hjghh' },
      {
        site,
        company: companies[1],
        formHeaderFields: [
          { label: 'Name of the Manager/Incharge:', key: 'form_header_manager_incharge' },
        ],
      }
    );
    expect(next.form_header_manager_incharge).toBe('Sample');
  });

  it('fills Form B establishment from company name (not site) when requested', () => {
    const company = companies[1];
    const site = {
      siteName: 'Theni Site',
      siteAddress:
        'Vayona Energy Pvt Ltd, 274 A, Rani mangammal main road, Govindanagaram - 625517 Theni Taluk, Theni District.',
      siteCity: 'Theni',
      siteState: 'TamilNadu',
      companyId: 'vayona',
    };
    const expected =
      'VAYONA ENERGY PRIVATE LIMITED, Vayona Energy Pvt Ltd, 274 A, Rani mangammal main road, Govindanagaram - 625517 Theni Taluk, Theni District., Theni, TamilNadu';
    const next = applySiteCompanyHeaderAutofill(
      {
        statutory_establishment_name_address: 'Theni Site',
      },
      {
        site,
        company,
        establishmentFromCompany: true,
        formHeaderFields: [
          {
            label: 'Name and Address of the Establishment:',
            key: 'statutory_establishment_name_address',
          },
        ],
      }
    );
    expect(next.statutory_establishment_name_address).toBe(expected);
    expect(next.statutory_establishment_name_address).not.toMatch(/^Theni Site/);
  });

  it('builds company name with site address without site name', () => {
    const expected =
      'VAYONA ENERGY PRIVATE LIMITED, Vayona Energy Pvt Ltd, 274 A, Rani mangammal main road, Govindanagaram - 625517 Theni Taluk, Theni District., Theni, TamilNadu';
    expect(
      buildCompanyNameWithSiteAddress(companies[1], {
        siteName: 'Theni Site',
        siteAddress:
          'Vayona Energy Pvt Ltd, 274 A, Rani mangammal main road, Govindanagaram - 625517 Theni Taluk, Theni District.',
        siteCity: 'Theni',
        siteState: 'TamilNadu',
      })
    ).toBe(expected);
  });

  it('writes Form U establishment over Theni Site template value', () => {
    const cells = {
      '4:1': {
        value: 'Name and Address of the Establishment : Theni Site, old address',
        alignment: {},
      },
      '4:2': { value: 'Theni Site, old address', alignment: {} },
    };
    const worksheet = {
      getCell: (r, c) => {
        const key = `${r}:${c}`;
        if (!cells[key]) cells[key] = { value: '', alignment: {} };
        return cells[key];
      },
    };
    const text =
      'VAYONA ENERGY PRIVATE LIMITED, Vayona Energy Pvt Ltd, 274 A, Rani mangammal main road, Govindanagaram - 625517 Theni Taluk, Theni District., Theni, TamilNadu';
    const wrote = writeFormUEstablishmentNameAddressToWorksheet(worksheet, text, { maxRow: 8 });
    expect(wrote).toBe(true);
    expect(String(cells['4:1'].value)).toContain('VAYONA ENERGY PRIVATE LIMITED');
    expect(String(cells['4:1'].value)).not.toMatch(/Theni Site/);
    // Default: combined on label only (no adjacent spill that caused "VAYONA ENERVAYONA" in Excel).
    expect(String(cells['4:2'].value)).toBe('Theni Site, old address');
  });

  it('optionally writes Form U establishment into adjacent value cell', () => {
    const cells = {
      '4:1': {
        value: 'Name and Address of the Establishment :',
        alignment: {},
      },
      '4:2': { value: '', alignment: {} },
    };
    const worksheet = {
      getCell: (r, c) => {
        const key = `${r}:${c}`;
        if (!cells[key]) cells[key] = { value: '', alignment: {} };
        return cells[key];
      },
    };
    const text = 'VAYONA ENERGY PRIVATE LIMITED, Site Address';
    writeFormUEstablishmentNameAddressToWorksheet(worksheet, text, {
      maxRow: 8,
      writeAdjacent: true,
    });
    expect(String(cells['4:2'].value)).toBe(text);
  });

  it('clears duplicated Form U employer/manager spill across header columns', () => {
    const { clearDuplicatedStatutoryHeaderValueSpill } = require('./statutorySiteCompanyHeaders');
    const cells = {
      '5:1': {
        value:
          'Name and Address of the Employer: VAYONA ENERGY PRIVATE LIMITED, Vayona Energy Pvt Ltd',
      },
      '5:13': { value: 'VAYONA ENERGY PRIVATE LIMITED, Vayona Energy Pvt Ltd' },
      '5:14': { value: 'VAYONA ENERVAYONA ENERGY PRIVATE LIMITED' },
      '6:1': { value: 'Name of the Manager/Incharge: Nilakantan Govindan' },
      '6:13': { value: 'Nilakantan Govindan' },
      '6:14': { value: 'Nilakantan GovNilakantan Govindan' },
    };
    const worksheet = {
      getCell: (r, c) => {
        const key = `${r}:${c}`;
        if (!cells[key]) cells[key] = { value: null };
        return cells[key];
      },
    };
    clearDuplicatedStatutoryHeaderValueSpill(worksheet, {
      rowFrom: 5,
      rowTo: 6,
      colFrom: 2,
      colTo: 20,
      valueHints: [
        'VAYONA ENERGY PRIVATE LIMITED, Vayona Energy Pvt Ltd',
        'Nilakantan Govindan',
      ],
    });
    expect(String(cells['5:1'].value)).toContain('Name and Address of the Employer');
    expect(cells['5:13'].value).toBeNull();
    expect(cells['5:14'].value).toBeNull();
    expect(String(cells['6:1'].value)).toContain('Nilakantan Govindan');
    expect(cells['6:13'].value).toBeNull();
    expect(cells['6:14'].value).toBeNull();
  });
});
