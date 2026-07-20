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
  if (/name\s+and\s+location\s+of\s+(?:the\s+)?work\s*site|name\s+and\s+location\s+of\s+worksite/.test(compact)) {
    return 'statutory_worksite';
  }
  if (/(?:name|nature)\s+and\s+location\s+of\s+work/.test(compact)) {
    return 'statutory_nature_location';
  }
  if (/establishment\s+in.*under\s+which\s+contract/.test(compact)) {
    return 'statutory_establishment_contract';
  }
  if (/name\s+and\s+address\s+of\s+principal\s+employer/.test(compact) ||
      /name\s+and\s+addressof\s+pricipal\s+employer/.test(compact) ||
      /name\s+and\s+address\s+of\s+pricipal\s+employer/.test(compact) ||
      (/principal\s+employer/.test(compact) && /name/.test(compact) && /address/.test(compact))) {
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

/**
 * Form B TN establishment: company legal name + site physical address (never siteName).
 * e.g. "VAYONA ENERGY PRIVATE LIMITED, Vayona Energy Pvt Ltd, 274 A, …, Theni, TamilNadu"
 */
export function buildCompanyNameWithSiteAddress(company, site) {
  const name = String(company?.companyName ?? company?.CompanyName ?? '').trim();
  const siteAddr = buildSiteEstablishmentAddressOnly(site);
  if (name && siteAddr) return `${name}, ${siteAddr}`;
  if (name) return buildCompanyNameAndAddress(company);
  if (siteAddr) return siteAddr;
  return buildCompanyNameAndAddress(company);
}

/** True for sample/demo company values that should not stick as employer autofill. */
export function looksLikeDemoCompanyHeaderValue(value) {
  const s = String(value || '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ');
  if (!s) return false;
  // e.g. "Company Name 2, 123, chamiers Road, Chennai..."
  if (/^company\s*name(\s*\d+)?(\b|,|:|$)/.test(s)) return true;
  // Excel template / Audit demo leftovers (not Company Details): "Delphi, Arumbakkam..."
  if (/^delphi(\b|,|:|$)/.test(s)) return true;
  return false;
}

/** Site/item companyName that is a sample brand — never use it to pick company_function. */
export function looksLikePlaceholderCompanyLinkName(name) {
  const s = String(name || '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ');
  if (!s) return false;
  if (/^company\s*name(\s*\d+)?$/.test(s)) return true;
  if (/^delphi(\s|$)/.test(s) || s === 'delhi') return true;
  return false;
}

function statutoryCompanyId(rec) {
  return String(
    rec?.ROWID ?? rec?.rowId ?? rec?.id ?? rec?.companyId ?? rec?.CompanyId ?? rec?.companyROWID ?? ''
  ).trim();
}

function statutoryCompanyName(rec) {
  return String(rec?.companyName ?? rec?.CompanyName ?? '').trim();
}

function findCompanyInListById(list, rawId) {
  const id = String(rawId ?? '').trim();
  if (!id || !Array.isArray(list)) return null;
  return list.find((co) => statutoryCompanyId(co) === id) || null;
}

function findCompanyInListByName(list, rawName) {
  const target = String(rawName ?? '').trim().toLowerCase();
  if (!target || !Array.isArray(list)) return null;
  const exact = list.find((co) => statutoryCompanyName(co).toLowerCase() === target);
  if (exact) return exact;
  // Partial: site "Vayona" ↔ "VAYONA ENERGY PRIVATE LIMITED"
  return (
    list.find((co) => {
      const n = statutoryCompanyName(co).toLowerCase();
      if (!n) return false;
      return n.includes(target) || target.includes(n);
    }) || null
  );
}

/**
 * Resolve the employer company for statutory headers.
 * Prefer Site Management company link, then the statutory row, then the sole company.
 * Never pick an arbitrary list[0] when multiple companies exist (avoids "Company Name 2").
 * Ignore Delphi / demo site links so Form U never exports template sample employer text.
 *
 * @param {object|null|undefined} item Statutory row
 * @param {array} companyDetailsList Company Details records
 * @param {object|null|undefined} site Optional Site Management record for the form's establishment
 */
export function resolveCompanyRecordForStatutory(item, companyDetailsList, site = null) {
  const list = Array.isArray(companyDetailsList) ? companyDetailsList : [];
  if (list.length === 0) return null;

  const nonDemoList = list.filter((co) => !looksLikeDemoCompanyHeaderValue(statutoryCompanyName(co)));
  const pickList = nonDemoList.length > 0 ? nonDemoList : list;

  const acceptResolved = (rec) => {
    if (!rec) return null;
    if (looksLikeDemoCompanyHeaderValue(statutoryCompanyName(rec)) && pickList !== list) {
      return null;
    }
    return rec;
  };

  if (site && typeof site === 'object') {
    const siteLinkName = site.companyName ?? site.CompanyName ?? site.company ?? site.Company;
    const fromSiteId = findCompanyInListById(
      pickList,
      site.companyId ?? site.CompanyId ?? site.companyROWID
    );
    const fromSiteName = looksLikePlaceholderCompanyLinkName(siteLinkName)
      ? null
      : findCompanyInListByName(pickList, siteLinkName);
    const fromSite = acceptResolved(fromSiteId || fromSiteName);
    if (fromSite) return fromSite;
  }

  const itemLinkName = item?.companyName ?? item?.CompanyName;
  const fromItemId = findCompanyInListById(pickList, item?.companyId ?? item?.CompanyId);
  const fromItemName = looksLikePlaceholderCompanyLinkName(itemLinkName)
    ? null
    : findCompanyInListByName(pickList, itemLinkName);
  const fromItem = acceptResolved(fromItemId || fromItemName);
  if (fromItem) return fromItem;

  // Prefer a single real company from company_function
  if (pickList.length === 1) return pickList[0];
  if (list.length === 1) return list[0];

  return null;
}

/**
 * Build employer "name, address" from company_function rows.
 * Never returns site/Delphi text. Falls back to first real company if link is incomplete.
 */
export function buildStatutoryEmployerTextFromCompanies(companyDetailsList, item = null, site = null) {
  const list = Array.isArray(companyDetailsList) ? companyDetailsList : [];
  let company = resolveCompanyRecordForStatutory(item, list, site);
  if (!company) {
    const real = list.filter((co) => !looksLikeDemoCompanyHeaderValue(statutoryCompanyName(co)));
    if (real.length === 1) company = real[0];
    else if (real.length > 1) {
      // Site GET often omits companyId; pick best name match if possible.
      const siteHint = String(
        site?.companyName ?? site?.CompanyName ?? site?.company ?? site?.Company ?? ''
      ).trim();
      if (siteHint && !looksLikePlaceholderCompanyLinkName(siteHint)) {
        company = findCompanyInListByName(real, siteHint);
      }
      if (!company) {
        company = real.find((co) => /vayona/i.test(statutoryCompanyName(co))) || real[0];
      }
    } else if (list.length === 1) {
      company = list[0];
    }
  }
  return buildCompanyNameAndAddress(company);
}

/** Write employer name/address onto Form U style header rows (label + adjacent value cell). */
export function writeFormUEmployerNameAddressToWorksheet(
  worksheet,
  employerText,
  { maxRow = 12, writeAdjacent = false } = {}
) {
  const text = String(employerText || '').trim();
  if (!worksheet || !text) return false;
  const rowLimit = Math.max(4, Math.min(30, Number(maxRow) || 12));
  let wrote = false;
  for (let r = 1; r <= rowLimit; r += 1) {
    for (let c = 1; c <= 12; c += 1) {
      const raw = excelCellValueToString(worksheet.getCell(r, c)?.value).trim();
      if (!raw) continue;
      const labelOnly = raw.split(':')[0].trim();
      if (!isPrincipalEmployerHeaderLabel(labelOnly) && !isPrincipalEmployerHeaderLabel(raw)) continue;
      // Prefer template layout: label in A, value in B (adjacent). Also write combined on label.
      const combined = formatStatutoryHeaderLabelValueExport(
        'Name and Address of the Employer:',
        raw,
        text
      );
      worksheet.getCell(r, c).value = combined;
      worksheet.getCell(r, c).alignment = {
        ...(worksheet.getCell(r, c).alignment || {}),
        wrapText: false,
        vertical: 'middle',
        horizontal: 'left'
      };
      // Optional adjacent value — disabled by default (repeating adjacent fills caused
      // "VAYONA ENERVAYONA ENI" spill across Form U header columns).
      if (writeAdjacent) {
        for (let ac = c + 1; ac <= Math.min(c + 2, 20); ac += 1) {
          const adj = excelCellValueToString(worksheet.getCell(r, ac)?.value).trim();
          if (!adj || /^enter\b/i.test(adj) || looksLikeDemoCompanyHeaderValue(adj)) {
            worksheet.getCell(r, ac).value = text;
            break;
          }
          // Don't overwrite another labeled header field.
          if (/:/.test(adj) && isPrincipalEmployerHeaderLabel(adj.split(':')[0])) break;
          if (/name\s+and\s+address|manager|registration|establishment/i.test(adj.split(':')[0] || '')) {
            break;
          }
        }
      }
      wrote = true;
      break;
    }
    if (wrote) break;
  }
  return wrote;
}

/**
 * Write Form U establishment name/address (company legal name + site address; never siteName alone).
 * Always overwrites stale site-name values such as "Theni Site, …".
 */
export function writeFormUEstablishmentNameAddressToWorksheet(
  worksheet,
  establishmentText,
  { maxRow = 12, writeAdjacent = false } = {}
) {
  const text = String(establishmentText || '').trim();
  if (!worksheet || !text) return false;
  const rowLimit = Math.max(4, Math.min(30, Number(maxRow) || 12));
  let wrote = false;
  for (let r = 1; r <= rowLimit; r += 1) {
    for (let c = 1; c <= 12; c += 1) {
      const raw = excelCellValueToString(worksheet.getCell(r, c)?.value).trim();
      if (!raw) continue;
      const labelOnly = raw.split(':')[0].trim();
      if (!isEstablishmentNameHeaderLabel(labelOnly) && !isEstablishmentNameHeaderLabel(raw)) continue;
      const combined = formatStatutoryHeaderLabelValueExport(
        'Name and Address of the Establishment:',
        raw,
        text
      );
      worksheet.getCell(r, c).value = combined;
      worksheet.getCell(r, c).alignment = {
        ...(worksheet.getCell(r, c).alignment || {}),
        wrapText: false,
        vertical: 'middle',
        horizontal: 'left'
      };
      if (writeAdjacent) {
        for (let ac = c + 1; ac <= Math.min(c + 2, 20); ac += 1) {
          const adj = excelCellValueToString(worksheet.getCell(r, ac)?.value).trim();
          const adjLabel = adj.split(':')[0] || '';
          if (/name\s+and\s+address|manager|registration|employer/i.test(adjLabel) && /:/.test(adj)) {
            break;
          }
          // Force overwrite prior site-name autofill (e.g. "Theni Site, …").
          if (
            !adj ||
            /^enter\b/i.test(adj) ||
            looksLikeDemoCompanyHeaderValue(adj) ||
            !/name\s+and\s+address\s+of\s+(?:the\s+)?employer|manager|registration/i.test(adj)
          ) {
            worksheet.getCell(r, ac).value = text;
            break;
          }
        }
      }
      wrote = true;
      break;
    }
    if (wrote) break;
  }
  return wrote;
}

/** Write Site Management Incharge Name onto Form U Manager/Incharge header rows. */
export function writeFormUManagerInchargeToWorksheet(
  worksheet,
  inchargeName,
  { maxRow = 12, writeAdjacent = false } = {}
) {
  const text = String(inchargeName || '').trim();
  if (!worksheet || !text) return false;
  const rowLimit = Math.max(4, Math.min(30, Number(maxRow) || 12));
  let wrote = false;
  for (let r = 1; r <= rowLimit; r += 1) {
    for (let c = 1; c <= 12; c += 1) {
      const raw = excelCellValueToString(worksheet.getCell(r, c)?.value).trim();
      if (!raw) continue;
      const labelOnly = raw.split(':')[0].trim();
      if (!isManagerInchargeHeaderLabel(labelOnly) && !isManagerInchargeHeaderLabel(raw)) continue;
      const combined = formatStatutoryHeaderLabelValueExport(
        'Name of the Manager/Incharge:',
        raw,
        text
      );
      worksheet.getCell(r, c).value = combined;
      worksheet.getCell(r, c).alignment = {
        ...(worksheet.getCell(r, c).alignment || {}),
        wrapText: false,
        vertical: 'middle',
        horizontal: 'left'
      };
      if (writeAdjacent) {
        for (let ac = c + 1; ac <= Math.min(c + 2, 20); ac += 1) {
          const adj = excelCellValueToString(worksheet.getCell(r, ac)?.value).trim();
          const adjLabel = adj.split(':')[0] || '';
          if (/name\s+and\s+address|registration|establishment|employer/i.test(adjLabel)) break;
          if (!adj || /^enter\b/i.test(adj) || !/name\s+of\s+the\s+manager|registration/i.test(adj)) {
            worksheet.getCell(r, ac).value = text;
            break;
          }
        }
      }
      wrote = true;
      break;
    }
    if (wrote) break;
  }
  return wrote;
}

export function isEstablishmentNameHeaderLabel(label) {
  const compact = normalizeStatutoryHeaderLabel(label);
  if (!compact) return false;
  // Address-only labels (not "Name and Address …").
  if (/^address\s+of\s+the\s+establishment$/.test(compact)) return false;
  if (/already\s+registered/.test(compact)) return false;
  if (/principal\s+employer/.test(compact)) return false;
  if (/contractor/.test(compact)) return false;
  if (/factory/.test(compact)) return false;
  if (
    /^name\s+of\s+the\s+establishment$/.test(compact) ||
    /^name\s+of\s+establishment$/.test(compact) ||
    /name\s+of\s+establishment\s+shop/.test(compact) ||
    /name\s+of\s+the\s+establishment\s+shop/.test(compact) ||
    /name\s+and\s+address\s+of\s+(?:the\s+)?establishment/.test(compact) ||
    /name\s+address\s+of\s+(?:the\s+)?establishment/.test(compact)
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
  // Tolerate "addressof", "Pricipal", and "Employer/Manager" template variants.
  const normalized = compact
    .replace(/\baddressof\b/g, 'address of')
    .replace(/\bpricipal\b/g, 'principal');
  return (
    /name\s+and\s+address\s+of\s+principal\s+employer/.test(normalized) ||
    /name\s+address\s+of\s+the\s+employer/.test(normalized) ||
    /name\s+and\s+address\s+of\s+the\s+employer/.test(normalized) ||
    /name\s+and\s+address\s+of\s+employer/.test(normalized) ||
    /principal\s+employer\s*(?:manager)?/.test(normalized) ||
    /^employer$/.test(normalized)
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
  return (
    /^registration\s+no/.test(compact) ||
    /registration\s+certificate\s+no/.test(compact) ||
    (/registration/.test(compact) && /\bno\b/.test(compact) && !/already\s+registered/.test(compact))
  );
}

export function isManagerInchargeHeaderLabel(label) {
  const compact = normalizeStatutoryHeaderLabel(label);
  if (!compact) return false;
  if (/name\s+of\s+the\s+manager/.test(compact) && /incharge|in\s*charge/.test(compact)) return true;
  if (/name\s+of\s+the\s+manager/.test(compact)) return true;
  if (/manager/.test(compact) && /incharge|in\s*charge/.test(compact)) return true;
  return false;
}

/** Tamil Nadu Form X gratuity column - keep manual / blank during autofill & export. */
export function isFormXLeaveGratuityHeader(label) {
  const compact = normalizeStatutoryHeaderLabel(label);
  if (!compact) return false;
  return /\bgratuit(?:y|ies)\b/.test(compact);
}

/** Form X leave register — separate "Month:" / "Year:" header labels (not "Month / Year"). */
export function isFormXMonthOnlyHeaderLabel(label) {
  const compact = normalizeStatutoryHeaderLabel(label);
  if (!compact) return false;
  return /^month$/.test(compact);
}

export function isFormXYearOnlyHeaderLabel(label) {
  const compact = normalizeStatutoryHeaderLabel(label);
  if (!compact) return false;
  return /^year$/.test(compact);
}

/** Form V TN Register of Employment — National/Festival benefit + Remarks stay blank. */
export function isFormVTamilNaduSkipAutofillHeader(label) {
  const compact = normalizeStatutoryHeaderLabel(label);
  if (!compact) return false;
  if (/^remarks?$/.test(compact)) return true;
  if (/\bbenefit\b/.test(compact) && /\bnational\b/.test(compact) && /\bholiday\b/.test(compact)) {
    return true;
  }
  if (/\bbenefit\b/.test(compact) && /\bfestival\b/.test(compact) && /\bholiday\b/.test(compact)) {
    return true;
  }
  if (/\bavailed\b/.test(compact) && /\bnational\b/.test(compact) && /\bholiday\b/.test(compact)) {
    return true;
  }
  if (/\bavailed\b/.test(compact) && /\bfestival\b/.test(compact) && /\bholiday\b/.test(compact)) {
    return true;
  }
  return false;
}

export const STATUTORY_ESTABLISHMENT_NAME_HEADER_KEYS = new Set([
  'form_a_establishment_name',
  'form_q_establishment',
  'form_h_establishment_name_address',
  'form_f_establishment_name_address',
  'form_t_establishment_name_address',
  'form_xxvi_ap_establishment',
  'statutory_establishment_name',
  'statutory_establishment_name_shop',
  'statutory_establishment_name_address'
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
  'statutory_principal_employer',
  'statutory_employer_name_address'
]);

export const STATUTORY_MONTH_YEAR_HEADER_KEYS = new Set(['form_t_month_year', 'form_xviii_month_year']);

export const STATUTORY_FORM_X_MONTH_HEADER_KEYS = new Set(['form_x_month']);
export const STATUTORY_FORM_X_YEAR_HEADER_KEYS = new Set(['form_x_year']);

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
    match: /name\s+and\s+address\s+of\s+(?:the\s+)?establishment/i,
    label: 'Name and Address of the Establishment:',
    key: 'statutory_establishment_name_address',
    kind: 'establishment_name'
  },
  {
    match: /name\s+and\s+address\s+of\s+(?:the\s+)?employer(?!\s+already)/i,
    label: 'Name and Address of the Employer:',
    key: 'statutory_employer_name_address',
    kind: 'principal_employer'
  },
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
    match: /name\s+and\s+address\s*of\s+pr(?:i|in)cipal\s+employer/i,
    label: 'Name and address of Principal Employer:',
    key: 'statutory_principal_employer',
    kind: 'principal_employer'
  },
  {
    match: /name\s+and\s+addressof\s+pr(?:i|in)cipal\s+employer/i,
    label: 'Name and address of Principal Employer:',
    key: 'statutory_principal_employer',
    kind: 'principal_employer'
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
    match: /registration\s+(certificate\s+)?no\.?/i,
    label: 'Registration Certificate No:',
    key: 'statutory_registration_no'
  },
  {
    match: /name\s+of\s+the\s+manager/i,
    label: 'Name of the Manager/Incharge:',
    key: 'form_header_manager_incharge'
  },
  {
    match: /manager\s*\/\s*incharge|manager.*in\s*charge/i,
    label: 'Name of the Manager/Incharge:',
    key: 'form_header_manager_incharge'
  },
  {
    match: /^month\s*\/\s*year$/i,
    label: 'Month / Year',
    key: 'form_t_month_year'
  },
  {
    match: /^month\s*:?\s*$/i,
    label: 'Month:',
    key: 'form_x_month'
  },
  {
    match: /^year\s*:?\s*$/i,
    label: 'Year:',
    key: 'form_x_year'
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
  {
    site,
    company,
    formHeaderFields = [],
    onlyIfEmpty = true,
    isPlaceholder = () => false,
    /** Form B TN: "Name and Address of the Establishment" uses company, not site. */
    establishmentFromCompany = false,
  } = {}
) {
  if (!headerData || typeof headerData !== 'object') return headerData;
  const companyNameAndAddress = buildCompanyNameAndAddress(company);
  const companyAddressOnly = (() => {
    if (!company || typeof company !== 'object') return '';
    const addr = String(company.companyAddress ?? company.CompanyAddress ?? '').trim();
    const city = String(company.city ?? company.City ?? '').trim();
    const state = String(company.state ?? company.State ?? '').trim();
    const postal = String(company.postalcode ?? company.PostalCode ?? '').trim();
    return [addr, city, state, postal].filter(Boolean).join(', ');
  })();
  const establishmentNameText = establishmentFromCompany
    ? buildCompanyNameWithSiteAddress(company, site) || companyNameAndAddress
    : buildSiteEstablishmentNameAndAddress(site);
  const establishmentAddressText = establishmentFromCompany
    ? buildSiteEstablishmentAddressOnly(site) || companyAddressOnly || companyNameAndAddress
    : buildSiteEstablishmentAddressOnly(site);
  // Employer always from company_function (Company Details) — never site address.
  const principalEmployerText = companyNameAndAddress;
  const establishmentFillOpts = establishmentFromCompany ? { force: true } : {};

  const out = { ...headerData };
  let changed = false;

  const fields = Array.isArray(formHeaderFields) ? formHeaderFields : [];

  // Strip Excel-template sample employer text (e.g. "Delphi, Arumbakkam...") before fill.
  const clearDemoEmployerKey = (key) => {
    if (!key) return;
    if (!looksLikeDemoCompanyHeaderValue(out[key])) return;
    out[key] = '';
    changed = true;
  };
  STATUTORY_PRINCIPAL_EMPLOYER_HEADER_KEYS.forEach((key) => clearDemoEmployerKey(key));
  clearDemoEmployerKey('form_xxvi_ap_employer');
  for (const field of fields) {
    if (isPrincipalEmployerHeaderLabel(field?.label)) clearDemoEmployerKey(field.key);
  }

  const shouldFill = (current, { allowReplaceDemoCompany = false, force = false } = {}) => {
    if (force) return true;
    const cur = String(current ?? '').trim();
    if (!onlyIfEmpty) return true;
    if (!cur || isPlaceholder(cur)) return true;
    // Replace leftover demo employer text (e.g. "Company Name 2, ...") with the real company
    if (allowReplaceDemoCompany && looksLikeDemoCompanyHeaderValue(cur)) return true;
    return false;
  };

  const fillKey = (key, value, opts = {}) => {
    if (!key || !value) return;
    if (!shouldFill(out[key], opts)) return;
    if (String(out[key] ?? '').trim() === String(value).trim()) return;
    out[key] = value;
    changed = true;
  };

  if (establishmentNameText) {
    STATUTORY_ESTABLISHMENT_NAME_HEADER_KEYS.forEach((key) =>
      fillKey(key, establishmentNameText, establishmentFillOpts)
    );
    fillKey('form_xv_establishment_contract_carried', establishmentNameText, establishmentFillOpts);
    fillKey('form25_establishment', establishmentNameText, establishmentFillOpts);
    fillKey('form_xxiii_establishment_contract_carried', establishmentNameText, establishmentFillOpts);
  }
  if (establishmentAddressText) {
    STATUTORY_ESTABLISHMENT_ADDRESS_HEADER_KEYS.forEach((key) =>
      fillKey(key, establishmentAddressText, establishmentFillOpts)
    );
  }
  // Always overwrite employer with company_function so forms stay consistent.
  if (principalEmployerText) {
    STATUTORY_PRINCIPAL_EMPLOYER_HEADER_KEYS.forEach((key) =>
      fillKey(key, principalEmployerText, { force: true })
    );
    fillKey('form_xxvi_ap_employer', principalEmployerText, { force: true });
  }

  const locationText = buildSiteLocationText(site);
  const registrationText = buildSiteRegistrationNumber(site);
  if (locationText) {
    STATUTORY_NATURE_LOCATION_HEADER_KEYS.forEach((key) => fillKey(key, locationText));
  }
  if (registrationText) {
    STATUTORY_REGISTRATION_HEADER_KEYS.forEach((key) => fillKey(key, registrationText));
  }
  const inchargeName = String(site?.inchargeName ?? site?.InchargeName ?? '').trim();
  // Always overwrite Manager/Incharge from Site Management Incharge Name.
  if (inchargeName) {
    fillKey('form_header_manager_incharge', inchargeName, { force: true });
  }

  for (const field of fields) {
    const key = field?.key;
    if (!key) continue;
    let value = '';
    let fillOpts = {};
    if (isEstablishmentNameHeaderLabel(field.label)) {
      value = establishmentNameText;
      fillOpts = establishmentFillOpts;
    } else if (isEstablishmentAddressHeaderLabel(field.label)) {
      value = establishmentAddressText;
      fillOpts = establishmentFillOpts;
    } else if (isPrincipalEmployerHeaderLabel(field.label)) {
      value = principalEmployerText;
      fillOpts = { force: true };
    } else if (isNatureLocationHeaderLabel(field.label)) value = locationText;
    else if (isRegistrationNoHeaderLabel(field.label)) value = registrationText;
    else if (isManagerInchargeHeaderLabel(field.label)) {
      value = inchargeName;
      fillOpts = { force: true };
    }
    fillKey(key, value, fillOpts);
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
      'statutory_establishment_name_address',
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
  if (
    /name\s+and\s+address\s+of\s+the\s+factory/.test(normalizeStatutoryHeaderLabel(label)) ||
    /^name\s+of\s+the\s+factory$/.test(normalizeStatutoryHeaderLabel(label))
  ) {
    return tryKeys([
      key,
      'statutory_factory_name_address',
      'form_vi_header_factory',
      'form12_header_factory',
      'form10_header_factory',
      'form14_header_factory',
      'form25_ap_header_factory'
    ]);
  }
  if (isPrincipalEmployerHeaderLabel(label)) {
    return tryKeys([
      'statutory_employer_name_address',
      'form_xxvi_ap_employer',
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
  if (isFormXMonthOnlyHeaderLabel(label) || STATUTORY_FORM_X_MONTH_HEADER_KEYS.has(key)) {
    return tryKeys(['form_x_month']);
  }
  if (isFormXYearOnlyHeaderLabel(label) || STATUTORY_FORM_X_YEAR_HEADER_KEYS.has(key)) {
    return tryKeys(['form_x_year']);
  }
  if (/for\s+the\s+period\s+from/i.test(normalizeStatutoryHeaderLabel(label))) {
    return tryKeys(['form_d_gj_period', 'statutory_period_from']);
  }
  if (isRegistrationNoHeaderLabel(label)) {
    return tryKeys(['statutory_registration_no', 'form12_header_registration']);
  }
  if (isManagerInchargeHeaderLabel(label)) {
    return tryKeys(['form_header_manager_incharge']);
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
  // Write each logical header field only once. Merged cells repeat the same label across
  // many columns — rewriting each one spills "VAYONA…" / manager name across the row.
  const writtenFieldKeys = new Set();
  const writtenSpecKeys = new Set();

  const writeCombinedOnLabelRow = (row, startCol, label, rawLabel, value) => {
    const text = formatStatutoryHeaderLabelValueExport(label, rawLabel || label, value);
    if (!String(text).trim()) return;
    // Prefer a short merge (label + a few value columns). Wide merges cause Excel to show
    // truncated fragments ("VA" / "VAYONA ENER") in every column of the band.
    const mergeEndCol =
      Number(colRightBound) > startCol
        ? Number(colRightBound)
        : Math.min(startCol + 3, scanColMax);
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
    // Only fill the immediate next empty cell — never walk far across the sheet.
    const adjMax = Math.min(startCol + 2, scanColMax);
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
    // Adjacent-only mode fills the value cell; combined already embeds Label : value.
    if (useAdjacent && !useCombined) writeAdjacentValue(row, startCol, val);
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
    if (rawNorm === labelNorm) return true;
    // Avoid matching short table headers like "Name" to "Name of the Establishment".
    if (rawNorm.length <= 12 || labelNorm.length <= 12) {
      const shorter = rawNorm.length <= labelNorm.length ? rawNorm : labelNorm;
      const longer = rawNorm.length <= labelNorm.length ? labelNorm : rawNorm;
      if (shorter.length < 8) return false;
      return longer.startsWith(shorter + ' ') || longer.startsWith(shorter + ':');
    }
    return rawNorm.includes(labelNorm) || labelNorm.includes(rawNorm);
  };

  for (let r = 1; r <= rowEnd; r += 1) {
    for (let c = 1; c <= scanColMax; c += 1) {
      const raw = excelCellValueToString(worksheet.getCell(r, c)?.value).trim();
      if (!raw) continue;
      let wrote = false;

      for (const field of fields) {
        const label = String(field?.label || '').trim();
        if (!label || !labelMatchesField(raw, label)) continue;
        const fieldKey = String(field?.key || label).trim();
        if (fieldKey && writtenFieldKeys.has(fieldKey)) continue;
        const val = resolveHeaderFieldExportValue(headerFormData, field);
        const existingAfterColon = String(raw.split(':').slice(1).join(':') || '').trim();
        const existingIsDemoEmployer =
          isPrincipalEmployerHeaderLabel(label) && looksLikeDemoCompanyHeaderValue(existingAfterColon);
        if (cellLooksLikeCompletedExport(raw, val) && !val && !existingIsDemoEmployer) continue;
        if (val) {
          writeHeaderValue(r, c, label, raw, val);
          if (fieldKey) writtenFieldKeys.add(fieldKey);
          wrote = true;
        } else if (existingIsDemoEmployer) {
          // Leave label-only cell until company_function value is available (do not export empty employer).
        }
        break;
      }
      if (wrote) continue;

      for (const spec of specs) {
        if (!spec.match.test(raw.split(':')[0])) continue;
        const specKey = String(spec.key || spec.label || '').trim();
        if (specKey && writtenSpecKeys.has(specKey)) continue;
        const val = resolveHeaderFieldExportValue(headerFormData, { key: spec.key, label: spec.label });
        const existingAfterColon = String(raw.split(':').slice(1).join(':') || '').trim();
        const existingIsDemoEmployer =
          spec.kind === 'principal_employer' && looksLikeDemoCompanyHeaderValue(existingAfterColon);
        if (cellLooksLikeCompletedExport(raw, val) && !val && !existingIsDemoEmployer) continue;
        if (val) {
          writeHeaderValue(r, c, spec.label, raw, val);
          if (specKey) writtenSpecKeys.add(specKey);
        } else if (existingIsDemoEmployer) {
          // Leave label-only cell until company_function value is available.
        }
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
        const fieldKey = String(field?.key || label).trim();
        if (fieldKey && writtenFieldKeys.has(fieldKey)) continue;
        const val = resolveHeaderFieldExportValue(headerFormData, field);
        if (val) {
          writeCombinedOnLabelRow(r, c, label, raw, val);
          if (fieldKey) writtenFieldKeys.add(fieldKey);
        }
        break;
      }
      if (/:/.test(raw)) continue;
      for (const spec of specs) {
        if (!spec.match.test(raw)) continue;
        const specKey = String(spec.key || spec.label || '').trim();
        if (specKey && writtenSpecKeys.has(specKey)) continue;
        const val = resolveHeaderFieldExportValue(headerFormData, { key: spec.key, label: spec.label });
        if (val) {
          writeCombinedOnLabelRow(r, c, spec.label, raw, val);
          if (specKey) writtenSpecKeys.add(specKey);
        }
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

/**
 * Form U / Form V: clear duplicated employer / manager / registration value fragments
 * that were spilled across header columns (shows as "VAYONA ENERVAYONA ENI" in Excel).
 * Keeps the first (leftmost) completed "Label : value" cell on each row.
 */
export function clearDuplicatedStatutoryHeaderValueSpill(
  worksheet,
  {
    rowFrom = 1,
    rowTo = 12,
    colFrom = 2,
    colTo = 80,
    valueHints = [],
  } = {}
) {
  if (!worksheet) return;
  const r0 = Math.max(1, Number(rowFrom) || 1);
  const r1 = Math.max(r0, Number(rowTo) || r0);
  const c0 = Math.max(1, Number(colFrom) || 2);
  const c1 = Math.max(c0, Number(colTo) || c0);
  const hints = (Array.isArray(valueHints) ? valueHints : [])
    .map((h) => String(h || '').trim())
    .filter((h) => h.length >= 4);
  const hintPrefixes = hints.map((h) => h.slice(0, Math.min(16, h.length)).toLowerCase());

  const looksLikeSpill = (raw) => {
    const s = String(raw || '').trim();
    if (!s) return false;
    // Keep real Month:/Year: and short markers.
    if (/^(month|year)\s*:/i.test(s) && s.length < 40) return false;
    if (/^\d+$/.test(s)) return false;
    const lower = s.toLowerCase();
    // Repeated company/manager fragments without a proper leading label.
    if (
      !/name\s+and\s+address|name\s+of\s+the\s+manager|registration\s+(certificate\s+)?no/i.test(s) &&
      hintPrefixes.some((p) => p && (lower.startsWith(p) || lower.includes(p)))
    ) {
      return true;
    }
    // Concatenated duplicates: "VAYONA ENERVAYONA" / "Nilakantan GovNilakantan"
    if (/([A-Za-z]{4,})\1/i.test(s.replace(/\s+/g, ''))) return true;
    return false;
  };

  for (let r = r0; r <= r1; r += 1) {
    // Find leftmost keep cell (completed label:value for identity headers).
    let keepCol = 0;
    for (let c = 1; c <= Math.min(12, c1); c += 1) {
      const raw = excelCellValueToString(worksheet.getCell(r, c)?.value).trim();
      if (
        /name\s+and\s+address\s+of\s+(?:the\s+)?(?:employer|establishment)|name\s+of\s+the\s+manager|registration\s+(certificate\s+)?no/i.test(
          raw
        ) &&
        /:/.test(raw)
      ) {
        keepCol = c;
        break;
      }
    }
    for (let c = c0; c <= c1; c += 1) {
      if (keepCol && c === keepCol) continue;
      const raw = excelCellValueToString(worksheet.getCell(r, c)?.value).trim();
      if (!raw) continue;
      // Never clear another labeled header on the same row.
      if (
        keepCol &&
        c !== keepCol &&
        /name\s+and\s+address|name\s+of\s+the\s+manager|registration\s+(certificate\s+)?no|month\s*:|year\s*:/i.test(
          raw.split(':')[0] || ''
        ) &&
        /:/.test(raw)
      ) {
        continue;
      }
      if (looksLikeSpill(raw) || (keepCol && c > keepCol && hintPrefixes.some((p) => p && raw.toLowerCase().includes(p)))) {
        // Only clear unlabeled / duplicate value cells to the right of the keep label.
        if (!keepCol || c > keepCol) {
          if (
            /name\s+and\s+address|name\s+of\s+the\s+manager|registration\s+(certificate\s+)?no/i.test(
              raw.split(':')[0] || ''
            ) &&
            /:/.test(raw) &&
            String(raw.split(':').slice(1).join(':') || '').trim().length >= 3
          ) {
            // Another complete labeled field — keep.
            continue;
          }
          worksheet.getCell(r, c).value = null;
        }
      }
    }
  }
}
