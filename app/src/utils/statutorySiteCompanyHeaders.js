/** Site + company header autofill for statutory forms (display + Excel export). */

export function normalizeStatutoryHeaderLabel(label) {
  return String(label || '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

/** Canonical key for matching template Excel labels to header field labels (ignores "the", numbering). */
export function statutoryHeaderLabelMatchKey(label) {
  const compact = normalizeStatutoryHeaderLabel(label)
    .replace(/^\d+\s+/, '')
    .replace(/\bthe\b/g, ' ')
    .replace(/\s*&\s*/g, ' and ')
    .replace(/\s+/g, ' ')
    .trim();
  if (!compact) return '';
  if (/name\s+and\s+address\s+of\s+contractor/.test(compact) && !/principal/.test(compact)) {
    return 'statutory_contractor';
  }
  if (/(?:name|nature)\s+and\s+location\s+of\s+work/.test(compact)) {
    return 'statutory_nature_location';
  }
  if (/establishment\s+in.*under\s+which\s+contract/.test(compact)) {
    return 'statutory_establishment_contract';
  }
  if (/name\s+and\s+address\s+of\s+principal\s+employer/.test(compact)) {
    return 'statutory_principal_employer';
  }
  if (/for\s+the\s+period\s+from/.test(compact)) {
    return 'statutory_period_from';
  }
  return compact;
}

export function buildSiteEstablishmentNameAndAddress(site) {
  if (!site || typeof site !== 'object') return '';
  const name = String(site.siteName ?? site.SiteName ?? '').trim();
  const addr = String(site.siteAddress ?? site.SiteAddress ?? '').trim();
  const city = String(site.siteCity ?? site.SiteCity ?? '').trim();
  const state = String(site.siteState ?? site.SiteState ?? '').trim();
  const parts = [];
  if (name) parts.push(name);
  const addrLine = [addr, city, state].filter(Boolean).join(', ');
  if (addrLine) parts.push(addrLine);
  return parts.join(', ').trim();
}

export function buildSiteEstablishmentAddressOnly(site) {
  if (!site || typeof site !== 'object') return '';
  const addr = String(site.siteAddress ?? site.SiteAddress ?? '').trim();
  const city = String(site.siteCity ?? site.SiteCity ?? '').trim();
  const state = String(site.siteState ?? site.SiteState ?? '').trim();
  return [addr, city, state].filter(Boolean).join(', ').trim();
}

export function buildSiteLocationText(site) {
  if (!site || typeof site !== 'object') return '';
  return String(site.location ?? site.Location ?? '').trim();
}

/** Site Management contractor fields → combined name and address line. */
export function buildSiteContractorNameAndAddress(site) {
  if (!site || typeof site !== 'object') return '';
  const name = String(site.contractorName ?? site.ContractorName ?? '').trim();
  const addr = String(site.contractorAddress ?? site.ContractorAddress ?? '').trim();
  const city = String(site.contractorCity ?? site.ContractorCity ?? '').trim();
  const state = String(site.contractorState ?? site.ContractorState ?? '').trim();
  const parts = [];
  if (name) parts.push(name);
  const addrLine = [addr, city, state].filter(Boolean).join(', ');
  if (addrLine) parts.push(addrLine);
  return parts.join(', ').trim();
}

/**
 * Form XXIII MP overtime register — header mapping from Site Management:
 * - Contractor field ← contractor name and address
 * - Nature/location of work ← site Location
 */
export function applyFormXXIIIMPHeaderFieldsFromSite(
  headerData,
  site,
  { formHeaderFields = [], isPlaceholder = () => false } = {}
) {
  if (!headerData || typeof headerData !== 'object' || !site) return headerData;
  const locationText = buildSiteLocationText(site);
  const contractorText = buildSiteContractorNameAndAddress(site);
  const out = { ...headerData };
  let changed = false;

  const setValue = (key, value, force = false) => {
    if (!key || !value) return;
    const cur = Object.prototype.hasOwnProperty.call(out, key) ? String(out[key] ?? '').trim() : '';
    if (!force && cur && !isPlaceholder(cur)) return;
    if (force && cur === value) return;
    out[key] = value;
    changed = true;
  };

  setValue('form_xxiii_contractor', contractorText, true);
  setValue('form_xxiii_nature_location_work', locationText, true);

  const fields = Array.isArray(formHeaderFields) ? formHeaderFields : [];
  for (const field of fields) {
    const key = field?.key;
    if (!key) continue;
    const labelNorm = normalizeStatutoryHeaderLabel(field.label);
    if (
      /name\s+and\s+address\s+of\s+contractor/.test(labelNorm) &&
      !/principal/.test(labelNorm)
    ) {
      setValue(key, contractorText, true);
    } else if (/nature\s+and\s+location\s+of\s+work/.test(labelNorm) || /name\s+and\s+location\s+of\s+work/.test(labelNorm)) {
      setValue(key, locationText, true);
    }
  }

  return changed ? out : headerData;
}

function isClraIndustryLabel(industry) {
  const s = String(industry || '').trim().toLowerCase();
  return s === 'clra' || s.includes('clra') || s.includes('contract labour') || s.includes('contract labor');
}

function isShopsAndEstablishmentIndustryLabel(industry) {
  const s = String(industry || '').trim().toLowerCase();
  return (s.includes('shop') || s.includes('shops')) && s.includes('establishment');
}

function isFactoryIndustryLabel(industry) {
  const s = String(industry || '').trim().toLowerCase();
  if (!s || isClraIndustryLabel(industry) || isShopsAndEstablishmentIndustryLabel(industry)) {
    return false;
  }
  return (
    s.includes('factories act') ||
    s.includes('factory act') ||
    s.includes('factories') ||
    s.includes('factory')
  );
}

/** Site Management RC number — industry-specific field first, then first non-empty RC. */
export function buildSiteRegistrationNumber(site) {
  if (!site || typeof site !== 'object') return '';
  const industry = String(site.industry ?? site.Industry ?? '').trim();
  const sand = String(site.sandERCNumber ?? site.SandERCNumber ?? '').trim();
  const factory = String(site.factoryRCNumber ?? site.FactoryRCNumber ?? '').trim();
  const clra = String(site.clraRCNumber ?? site.CLRARCNumber ?? '').trim();
  if (isClraIndustryLabel(industry) && clra) return clra;
  if (isShopsAndEstablishmentIndustryLabel(industry) && sand) return sand;
  if (isFactoryIndustryLabel(industry) && factory) return factory;
  return clra || sand || factory || '';
}

export function buildCompanyNameAndAddress(company) {
  if (!company || typeof company !== 'object') return '';
  const name = String(company.companyName ?? company.CompanyName ?? '').trim();
  const addr = String(company.companyAddress ?? company.CompanyAddress ?? '').trim();
  const city = String(company.city ?? company.City ?? '').trim();
  const state = String(company.state ?? company.State ?? '').trim();
  const postal = String(company.postalcode ?? company.postalCode ?? company.PostalCode ?? '').trim();
  const parts = [];
  if (name) parts.push(name);
  const addrLine = [addr, city, state, postal].filter(Boolean).join(', ');
  if (addrLine) parts.push(addrLine);
  return parts.join(', ').trim();
}

export function resolveCompanyRecordForStatutory(item, companyDetailsList) {
  const list = Array.isArray(companyDetailsList) ? companyDetailsList : [];
  if (list.length === 0) return null;
  const target = String(item?.companyName ?? item?.CompanyName ?? '').trim().toLowerCase();
  if (target) {
    const match = list.find(
      (co) => String(co?.companyName ?? co?.CompanyName ?? '').trim().toLowerCase() === target
    );
    if (match) return match;
  }
  return list[0] || null;
}

export function isEstablishmentNameHeaderLabel(label) {
  const compact = normalizeStatutoryHeaderLabel(label);
  if (!compact) return false;
  if (/address\s+of\s+the\s+establishment/.test(compact)) return false;
  if (/already\s+registered/.test(compact)) return false;
  if (/principal\s+employer/.test(compact)) return false;
  if (/contractor/.test(compact)) return false;
  if (/factory/.test(compact)) return false;
  if (
    /^name\s+of\s+the\s+establishment$/.test(compact) ||
    /^name\s+of\s+establishment$/.test(compact) ||
    /name\s+of\s+establishment\s+shop/.test(compact) ||
    /name\s+of\s+the\s+establishment\s+shop/.test(compact) ||
    /name\s+and\s+address\s+of\s+the\s+establishment$/.test(compact) ||
    /name\s+address\s+of\s+the\s+establishment$/.test(compact)
  ) {
    return true;
  }
  return false;
}

export function isEstablishmentAddressHeaderLabel(label) {
  const compact = normalizeStatutoryHeaderLabel(label);
  if (!compact) return false;
  return /^address\s+of\s+the\s+establishment$/.test(compact);
}

export function isPrincipalEmployerHeaderLabel(label) {
  const compact = normalizeStatutoryHeaderLabel(label);
  if (!compact) return false;
  return (
    /name\s+and\s+address\s+of\s+principal\s+employer/.test(compact) ||
    /name\s+address\s+of\s+the\s+employer/.test(compact) ||
    /name\s+and\s+address\s+of\s+the\s+employer/.test(compact) ||
    /name\s+and\s+address\s+of\s+employer/.test(compact) ||
    /^employer$/.test(compact)
  );
}

export function isMonthYearHeaderLabel(label) {
  const compact = normalizeStatutoryHeaderLabel(label);
  if (!compact) return false;
  return /^month\s+year$/.test(compact) || /^month\s*\/\s*year$/.test(compact);
}

export function isNatureLocationHeaderLabel(label) {
  const compact = normalizeStatutoryHeaderLabel(label);
  if (!compact) return false;
  return /(?:name|nature)\s+and\s+location\s+of\s+work/.test(compact);
}

export function isRegistrationNoHeaderLabel(label) {
  const compact = normalizeStatutoryHeaderLabel(label);
  if (!compact) return false;
  return /^registration\s+no/.test(compact);
}

export const STATUTORY_ESTABLISHMENT_NAME_HEADER_KEYS = new Set([
  'form_a_establishment_name',
  'form_q_establishment',
  'form_h_establishment_name_address',
  'form_f_establishment_name_address',
  'form_t_establishment_name_address',
  'form_xxvi_ap_establishment',
  'statutory_establishment_name',
  'statutory_establishment_name_shop'
]);

export const STATUTORY_ESTABLISHMENT_ADDRESS_HEADER_KEYS = new Set([
  'form_a_establishment_address',
  'statutory_establishment_address'
]);

export const STATUTORY_PRINCIPAL_EMPLOYER_HEADER_KEYS = new Set([
  'form25_principal_employer',
  'form_xv_principal_employer',
  'form_xvi_principal_employer',
  'form_xvii_principal_employer',
  'form_xviii_principal_employer',
  'form_xxiii_principal_employer',
  'form_q_ka_employer',
  'form_t_employer',
  'statutory_principal_employer'
]);

export const STATUTORY_MONTH_YEAR_HEADER_KEYS = new Set(['form_t_month_year', 'form_xviii_month_year']);

export const STATUTORY_NATURE_LOCATION_HEADER_KEYS = new Set([
  'form_xv_nature_location_work',
  'form_xvi_nature_location_work',
  'form_xvii_nature_location_work',
  'form_xviii_nature_location_work',
  'form_xxiii_nature_location_work',
  'form_xix_ap_nature_location'
]);

export const STATUTORY_REGISTRATION_HEADER_KEYS = new Set([
  'statutory_registration_no',
  'form12_header_registration'
]);

export const STATUTORY_SITE_COMPANY_SHEET_HEADER_SPECS = [
  {
    match: /name\s+of\s+the\s+establishment(?!\s+already)/i,
    label: 'Name of the Establishment:',
    key: 'statutory_establishment_name',
    kind: 'establishment_name'
  },
  {
    match: /^name\s+of\s+(?:the\s+)?establishment(?!.*principal)(?!.*employer)/i,
    label: 'Name of Establishment',
    key: 'statutory_establishment_name',
    kind: 'establishment_name'
  },
  {
    match: /name\s+of\s+establishment\s*\/\s*shop/i,
    label: 'Name of Establishment / Shop:',
    key: 'statutory_establishment_name_shop',
    kind: 'establishment_name'
  },
  {
    match: /address\s+of\s+the\s+establishment/i,
    label: 'Address of the Establishment:',
    key: 'statutory_establishment_address',
    kind: 'establishment_address'
  },
  {
    match: /name\s+and\s+address\s+of\s+principal\s+employer/i,
    label: 'Name and address of Principal Employer:',
    key: 'statutory_principal_employer',
    kind: 'principal_employer'
  },
  {
    match: /^employer\s*:?$/i,
    label: 'Employer:',
    key: 'statutory_principal_employer',
    kind: 'principal_employer'
  },
  {
    match: /name\s+and\s+address\s+of\s+(?:the\s+)?contractor/i,
    label: '1. Name and Address of Contractor.',
    key: 'form_xxiii_contractor'
  },
  {
    match: /(?:name|nature)\s+and\s+location\s+of\s+work/i,
    label: '2. Nature and location of work.',
    key: 'form_xxiii_nature_location_work'
  },
  {
    match: /establishment\s+in\s*\/?\s*under\s+which\s+contract\s+is\s+carried\s+on/i,
    label: '3. Name and address of establishment in/under which contract is carried on',
    key: 'form_xxiii_establishment_contract_carried'
  },
  {
    match: /^registration\s+no\.?/i,
    label: 'Registration No.:',
    key: 'statutory_registration_no'
  },
  {
    match: /^month\s*\/\s*year$/i,
    label: 'Month / Year',
    key: 'form_t_month_year'
  },
  {
    match: /name\s+and\s+address\s+of\s+the\s+establishment/i,
    label: 'Name and address of the Establishment',
    key: 'form_t_establishment_name_address',
    kind: 'establishment_name'
  },
  {
    match: /name\s+and\s+address\s+of\s+employer/i,
    label: 'Name and Address of employer',
    key: 'form_t_employer',
    kind: 'principal_employer'
  },
  {
    match: /for\s+the\s+period\s+from/i,
    label: 'For the period From',
    key: 'form_d_gj_period'
  }
];

export function enrichEstablishmentPrincipalEmployerHeaderFields(headerFields) {
  const fields = Array.isArray(headerFields) ? [...headerFields] : [];
  const norm = normalizeStatutoryHeaderLabel;
  const hasLabel = (label) => fields.some((f) => norm(f?.label) === norm(label));
  const hasKind = (kind) =>
    fields.some((f) => {
      if (kind === 'establishment_name') {
        return (
          STATUTORY_ESTABLISHMENT_NAME_HEADER_KEYS.has(f?.key) ||
          isEstablishmentNameHeaderLabel(f?.label)
        );
      }
      if (kind === 'establishment_address') {
        return (
          STATUTORY_ESTABLISHMENT_ADDRESS_HEADER_KEYS.has(f?.key) ||
          isEstablishmentAddressHeaderLabel(f?.label)
        );
      }
      if (kind === 'principal_employer') {
        return (
          STATUTORY_PRINCIPAL_EMPLOYER_HEADER_KEYS.has(f?.key) ||
          isPrincipalEmployerHeaderLabel(f?.label)
        );
      }
      return false;
    });

  STATUTORY_SITE_COMPANY_SHEET_HEADER_SPECS.forEach((spec) => {
    if (hasKind(spec.kind) || hasLabel(spec.label)) return;
    fields.push({ label: spec.label, value: '', key: spec.key });
  });
  return fields;
}

export function enrichEstablishmentPrincipalEmployerHeadersFromSheet(
  headerFields,
  headerRowIndex,
  effectiveSheetCols,
  getMergedAwareCellText
) {
  let fields = enrichEstablishmentPrincipalEmployerHeaderFields(headerFields);
  const norm = normalizeStatutoryHeaderLabel;
  const hasLabel = (label) => fields.some((f) => norm(f?.label) === norm(label));
  const maxR = Math.max(0, Number(headerRowIndex) || 0);
  const maxC = Math.max(40, Number(effectiveSheetCols) || 40);

  for (let r = 0; r < maxR; r += 1) {
    for (let c = 0; c < maxC; c += 1) {
      const cellText = String(getMergedAwareCellText(r, c) || '').trim();
      if (!cellText) continue;
      for (const spec of STATUTORY_SITE_COMPANY_SHEET_HEADER_SPECS) {
        if (!spec.match.test(cellText) || hasLabel(spec.label)) continue;
        fields.push({ label: spec.label, value: '', key: spec.key });
      }
    }
  }
  return fields;
}

export function applySiteCompanyHeaderAutofill(
  headerData,
  { site, company, formHeaderFields = [], onlyIfEmpty = true, isPlaceholder = () => false } = {}
) {
  if (!headerData || typeof headerData !== 'object') return headerData;
  const establishmentNameText = buildSiteEstablishmentNameAndAddress(site);
  const establishmentAddressText = buildSiteEstablishmentAddressOnly(site);
  const principalEmployerText = buildCompanyNameAndAddress(company);

  const out = { ...headerData };
  let changed = false;

  const shouldFill = (current) => {
    const cur = String(current ?? '').trim();
    if (!onlyIfEmpty) return true;
    return !cur || isPlaceholder(cur);
  };

  const fillKey = (key, value) => {
    if (!key || !value) return;
    if (!shouldFill(out[key])) return;
    out[key] = value;
    changed = true;
  };

  if (establishmentNameText) {
    STATUTORY_ESTABLISHMENT_NAME_HEADER_KEYS.forEach((key) => fillKey(key, establishmentNameText));
    fillKey('form_xv_establishment_contract_carried', establishmentNameText);
    fillKey('form25_establishment', establishmentNameText);
    fillKey('form_xxiii_establishment_contract_carried', establishmentNameText);
  }
  if (establishmentAddressText) {
    STATUTORY_ESTABLISHMENT_ADDRESS_HEADER_KEYS.forEach((key) => fillKey(key, establishmentAddressText));
  }
  if (principalEmployerText) {
    STATUTORY_PRINCIPAL_EMPLOYER_HEADER_KEYS.forEach((key) => fillKey(key, principalEmployerText));
  }

  const locationText = buildSiteLocationText(site);
  const registrationText = buildSiteRegistrationNumber(site);
  if (locationText) {
    STATUTORY_NATURE_LOCATION_HEADER_KEYS.forEach((key) => fillKey(key, locationText));
  }
  if (registrationText) {
    STATUTORY_REGISTRATION_HEADER_KEYS.forEach((key) => fillKey(key, registrationText));
  }

  const fields = Array.isArray(formHeaderFields) ? formHeaderFields : [];
  for (const field of fields) {
    const key = field?.key;
    if (!key) continue;
    let value = '';
    if (isEstablishmentNameHeaderLabel(field.label)) value = establishmentNameText;
    else if (isEstablishmentAddressHeaderLabel(field.label)) value = establishmentAddressText;
    else if (isPrincipalEmployerHeaderLabel(field.label)) value = principalEmployerText;
    else if (isNatureLocationHeaderLabel(field.label)) value = locationText;
    else if (isRegistrationNoHeaderLabel(field.label)) value = registrationText;
    fillKey(key, value);
  }

  return changed ? out : headerData;
}

export function excelCellValueToString(val) {
  if (val == null) return '';
  if (typeof val === 'string' || typeof val === 'number' || typeof val === 'boolean') return String(val);
  if (val instanceof Date) return val.toISOString();
  if (typeof val === 'object') {
    if (Array.isArray(val.richText)) return val.richText.map((rt) => rt?.text || '').join('');
    if (val.text != null) return String(val.text);
    if (val.result != null) return String(val.result);
  }
  return '';
}

/** Resolve export value for a header field (direct key + establishment/principal aliases). */
export function resolveHeaderFieldExportValue(headerFormData, field) {
  if (!headerFormData || typeof headerFormData !== 'object') return '';
  const tryKeys = (keys) => {
    for (const k of keys) {
      if (!k) continue;
      const v = headerFormData[k];
      if (v != null && String(v).trim() !== '') return String(v).trim();
    }
    return '';
  };
  const key = field?.key;
  const label = field?.label || '';
  const direct = tryKeys([key]);
  if (direct) return direct;
  if (isEstablishmentNameHeaderLabel(label)) {
    return tryKeys([
      'statutory_establishment_name_shop',
      'statutory_establishment_name',
      'form_a_establishment_name',
      'form_q_establishment',
      'form25_establishment',
      'form_xv_establishment_contract_carried',
      'form_xxiii_establishment_contract_carried',
      'form_xxvi_ap_establishment',
      'form_h_establishment_name_address',
      'form_f_establishment_name_address',
      'form_t_establishment_name_address'
    ]);
  }
  if (isEstablishmentAddressHeaderLabel(label)) {
    return tryKeys(['statutory_establishment_address', 'form_a_establishment_address']);
  }
  if (isPrincipalEmployerHeaderLabel(label)) {
    return tryKeys([
      'form_t_employer',
      'form_q_ka_employer',
      'statutory_principal_employer',
      'form25_principal_employer',
      'form_xv_principal_employer',
      'form_xvi_principal_employer',
      'form_xvii_principal_employer',
      'form_xviii_principal_employer',
      'form_xxiii_principal_employer'
    ]);
  }
  if (isMonthYearHeaderLabel(label)) {
    return tryKeys(['form_t_month_year', 'form_xviii_month_year']);
  }
  if (/for\s+the\s+period\s+from/i.test(normalizeStatutoryHeaderLabel(label))) {
    return tryKeys(['form_d_gj_period', 'statutory_period_from']);
  }
  if (isRegistrationNoHeaderLabel(label)) {
    return tryKeys(['statutory_registration_no', 'form12_header_registration']);
  }
  if (/contractor/i.test(normalizeStatutoryHeaderLabel(label)) && !/principal/.test(normalizeStatutoryHeaderLabel(label))) {
    const contractorVal = tryKeys(['form_xxiii_contractor', 'form_xviii_contractor', 'form_xvii_contractor', 'form_xvi_contractor']);
    if (contractorVal) return contractorVal;
    if (/name\s+and\s+address\s+of\s+contractor/.test(normalizeStatutoryHeaderLabel(label))) {
      return tryKeys(['form_xxiii_contractor']);
    }
  }
  if (isNatureLocationHeaderLabel(label)) {
    return tryKeys([
      'form_xxiii_nature_location_work',
      'form_xvii_nature_location_work',
      'form_xviii_nature_location_work',
      'form_xvi_nature_location_work',
      'form_xv_nature_location_work',
      'form_xix_ap_nature_location'
    ]);
  }
  if (/establishment\s+in.*under\s+which\s+contract/i.test(normalizeStatutoryHeaderLabel(label))) {
    return tryKeys([
      'form_xxiii_establishment_contract_carried',
      'form_xv_establishment_contract_carried',
      'form_xviii_establishment_contract_carried',
      'form_xvii_establishment_contract_carried',
      'form_xvi_establishment_contract_carried'
    ]);
  }
  return String(field?.value ?? '').trim();
}

/** Format exported header as "Label : value" (matches statutory form layout). */
export function formatStatutoryHeaderLabelValueExport(label, rawLabel, value) {
  let labelText = String(label || rawLabel || '')
    .trim()
    .replace(/^\d+\.\s*/, '')
    .replace(/:+\s*$/, '')
    .replace(/\s+/g, ' ')
    .trim();
  if (/name\s+of\s+establishment\s*\/\s*shop/i.test(labelText)) {
    labelText = 'Name of Establishment';
  }
  const val = String(value ?? '').trim();
  if (!labelText) return val;
  if (!val) return `${labelText} :`;
  return `${labelText} : ${val}`;
}

/**
 * Write header field values into Excel beside their labels (merge-aware empty-cell scan).
 * Exported cell text: "Label : value"
 */
export function writeStatutoryHeaderFieldsToExcelJsWorksheet(
  worksheet,
  {
    headerFormData,
    parsedFormHeader,
    headerRowEnd,
    maxScanRows,
    maxScanCols,
    colRightBound,
    writeMode = 'combined'
  } = {}
) {
  if (!worksheet || !headerFormData || typeof headerFormData !== 'object') return;
  const fields = Array.isArray(parsedFormHeader?.fields) ? parsedFormHeader.fields : [];
  const specs = STATUTORY_SITE_COMPANY_SHEET_HEADER_SPECS;
  if (fields.length === 0 && specs.length === 0) return;

  const normalize = normalizeStatutoryHeaderLabel;
  const rowEnd = Math.max(1, Number(headerRowEnd) || 35);
  const colLimit = Math.max(20, Number(maxScanCols) || 80);
  const scanColMax =
    Number(colRightBound) > 0 ? Math.min(colLimit, Number(colRightBound)) : colLimit;
  const useCombined = writeMode === 'combined' || writeMode === 'both';
  const useAdjacent = writeMode === 'adjacent' || writeMode === 'both';

  const writeCombinedOnLabelRow = (row, startCol, label, rawLabel, value) => {
    const text = formatStatutoryHeaderLabelValueExport(label, rawLabel || label, value);
    if (!String(text).trim()) return;
    const mergeEndCol =
      Number(colRightBound) > startCol
        ? Number(colRightBound)
        : Math.min(startCol + 11, scanColMax);
    if (mergeEndCol > startCol) {
      try {
        worksheet.mergeCells(row, startCol, row, mergeEndCol);
      } catch (_) {
        // Template may already define merges — still write into the anchor cell.
      }
    }
    const cell = worksheet.getCell(row, startCol);
    cell.value = text;
    cell.alignment = {
      ...(cell.alignment || {}),
      wrapText: false,
      vertical: 'middle',
      horizontal: 'left'
    };
    const wsRow = worksheet.getRow(row);
    if (wsRow) {
      wsRow.height = Math.min(Math.max(wsRow.height || 18, 15), 22);
    }
  };

  const writeAdjacentValue = (row, startCol, value) => {
    const val = String(value ?? '').trim();
    if (!val) return;
    const adjMax =
      Number(colRightBound) > 0
        ? Math.min(startCol + 14, Number(colRightBound))
        : startCol + 14;
    for (let ac = startCol + 1; ac <= adjMax; ac += 1) {
      const adj = excelCellValueToString(worksheet.getCell(row, ac)?.value).trim();
      if (!adj || /^enter\b/i.test(adj)) {
        worksheet.getCell(row, ac).value = val;
        return;
      }
    }
  };

  const writeHeaderValue = (row, startCol, label, rawLabel, value) => {
    const val = String(value ?? '').trim();
    if (!val) return;
    if (useCombined) writeCombinedOnLabelRow(row, startCol, label, rawLabel, val);
    if (useAdjacent) writeAdjacentValue(row, startCol, val);
  };

  const cellLooksLikeCompletedExport = (raw, newVal) => {
    const text = String(raw || '').trim();
    if (!text || !/:/.test(text)) return false;
    const parts = text.split(':');
    const existingVal = String(parts.slice(1).join(':') || '').trim();
    const nextVal = String(newVal ?? '').trim();
    if (nextVal && existingVal !== nextVal) return false;
    return parts.length >= 2 && existingVal.length >= 3;
  };

  const labelMatchesField = (raw, label) => {
    const rawLabelOnly = String(raw).split(':')[0].trim();
    const rawKey = statutoryHeaderLabelMatchKey(rawLabelOnly);
    const labelKey = statutoryHeaderLabelMatchKey(label);
    if (rawKey && labelKey && rawKey === labelKey) return true;
    const rawNorm = normalize(String(rawLabelOnly).replace(/:+$/, ''));
    const labelNorm = normalize(String(label).replace(/:+$/, ''));
    if (!rawNorm || !labelNorm) return false;
    return rawNorm === labelNorm || rawNorm.includes(labelNorm) || labelNorm.includes(rawNorm);
  };

  for (let r = 1; r <= rowEnd; r += 1) {
    for (let c = 1; c <= scanColMax; c += 1) {
      const raw = excelCellValueToString(worksheet.getCell(r, c)?.value).trim();
      if (!raw) continue;
      let wrote = false;

      for (const field of fields) {
        const label = String(field?.label || '').trim();
        if (!label || !labelMatchesField(raw, label)) continue;
        const val = resolveHeaderFieldExportValue(headerFormData, field);
        if (cellLooksLikeCompletedExport(raw, val) && !val) continue;
        if (val) {
          writeHeaderValue(r, c, label, raw, val);
          wrote = true;
        }
        break;
      }
      if (wrote) continue;

      for (const spec of specs) {
        if (!spec.match.test(raw.split(':')[0])) continue;
        const val = resolveHeaderFieldExportValue(headerFormData, { key: spec.key, label: spec.label });
        if (cellLooksLikeCompletedExport(raw, val) && !val) continue;
        if (val) writeHeaderValue(r, c, spec.label, raw, val);
        break;
      }
    }
  }

  // Repair pass: label-only template cells (no "Label : value" yet).
  for (let r = 1; r <= rowEnd; r += 1) {
    for (let c = 1; c <= scanColMax; c += 1) {
      const raw = excelCellValueToString(worksheet.getCell(r, c)?.value).trim();
      if (!raw || (/:/.test(raw) && cellLooksLikeCompletedExport(raw))) continue;
      for (const field of fields) {
        const label = String(field?.label || '').trim();
        if (!label || !labelMatchesField(raw, label)) continue;
        const val = resolveHeaderFieldExportValue(headerFormData, field);
        if (val) writeCombinedOnLabelRow(r, c, label, raw, val);
        break;
      }
      if (/:/.test(raw)) continue;
      for (const spec of specs) {
        if (!spec.match.test(raw)) continue;
        const val = resolveHeaderFieldExportValue(headerFormData, { key: spec.key, label: spec.label });
        if (val) writeCombinedOnLabelRow(r, c, spec.label, raw, val);
        break;
      }
    }
  }
}

/** Remove accidental header values written outside the wage-register table box. */
export function clearStatutoryHeaderCellsBeyondColumn(
  worksheet,
  { colFrom, colTo, rowFrom = 1, rowTo = 25 } = {}
) {
  if (!worksheet) return;
  const c0 = Math.max(1, Number(colFrom) || 1);
  const c1 = Math.max(c0, Number(colTo) || c0);
  const r0 = Math.max(1, Number(rowFrom) || 1);
  const r1 = Math.max(r0, Number(rowTo) || r0);
  const spillRe =
    /address\s+of\s+the\s+establishment|name\s+and\s+address\s+of\s+principal\s+employer|establishment\s+in.*under\s+which\s+contract|name\s+of\s+establishment\s*\/\s*shop|registration\s+no|name\s+and\s+address\s+of\s+contractor|nature\s+and\s+location\s+of\s+work/i;
  for (let r = r0; r <= r1; r += 1) {
    for (let c = c0; c <= c1; c += 1) {
      const raw = excelCellValueToString(worksheet.getCell(r, c)?.value).trim();
      if (!raw) continue;
      if (spillRe.test(raw)) {
        worksheet.getCell(r, c).value = '';
      }
    }
  }
}
