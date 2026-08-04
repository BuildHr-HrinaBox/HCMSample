import React, { useState, useEffect, useCallback, useMemo, useRef } from 'react';
import { Link } from 'react-router-dom';
import { CheckCircle2, Trash2 } from 'lucide-react';
import * as XLSX from 'xlsx';
import './SiteManagement.css';
import './CompanyDetails.css';
import { resolveLoginEmailString, stringifyUserEmail } from '../utils/resolveLoginEmail';
import {
  filterSitesForLoginUser,
  hasInchargeSiteScope,
  siteIndustry,
  siteInchargeEmail,
  siteStateFromRecord,
} from '../utils/siteInchargeScope';
import { INDIAN_CITIES } from '../utils/indianCities';
import { INDIAN_STATES } from '../utils/indianStates';
import CityCombobox, { StateCombobox } from '../components/CityCombobox';

const API_BASE = '/server/sitemanagement_function';
const COMPANY_API = '/server/company_function';

const SITE_TABLE_PAGE_SIZE = 10;

/** Site CSV export/import columns (order matters for round-trip). */
/** Site CSV export/import columns (ID omitted from export; import still accepts ID if present). */
const SITE_CSV_COLS = [
  { key: 'companyName', label: 'Company Name' },
  { key: 'siteName', label: 'Site Name' },
  { key: 'siteAddress', label: 'Address' },
  { key: 'siteCity', label: 'City' },
  { key: 'siteState', label: 'State' },
  { key: 'sitePostalCode', label: 'Postal code' },
  { key: 'unitNo', label: 'Unit no' },
  { key: 'contractorName', label: 'Contractor Name' },
  { key: 'contractorAddress', label: 'Contractor Address' },
  { key: 'contractorEmail', label: 'Contractor Email' },
  { key: 'contractorPhone', label: 'Contractor Phone' },
  { key: 'contractorCity', label: 'Contractor City' },
  { key: 'contractorState', label: 'Contractor State' },
  { key: 'inchargeName', label: 'Incharge Name' },
  { key: 'inchargePhone', label: 'Incharge Phone' },
  { key: 'inchargeEmail', label: 'Incharge Mail Id' },
  { key: 'inchargeDesignation', label: 'Incharge Designation' },
  { key: 'industry', label: 'Industry' },
  { key: 'location', label: 'Location' }
];

/** Older exports used ID + short incharge labels (Name / Phone / Mail Id / Designation). */
const SITE_CSV_LEGACY_LABELS = [
  'ID',
  'Company Name',
  'Site Name',
  'Address',
  'City',
  'State',
  'Postal code',
  'Unit no',
  'Contractor Name',
  'Contractor Address',
  'Contractor Email',
  'Contractor Phone',
  'Contractor City',
  'Contractor State',
  'Name',
  'Phone',
  'Mail Id',
  'Designation',
  'Industry',
  'Location'
];

/** Older export with ID + unique incharge labels (before ID was removed from export). */
const SITE_CSV_LEGACY_WITH_ID_LABELS = ['ID', ...SITE_CSV_COLS.map((c) => c.label)];

function normalizeCsvHeader(h) {
  return String(h ?? '')
    .replace(/^\uFEFF/, '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ');
}

function stripCsvExcelTextPrefix(v) {
  const s = String(v ?? '').trim();
  return s.startsWith('\t') ? s.slice(1).trim() : s;
}

function headersMatchLabelList(headerRow, labels) {
  if (!Array.isArray(headerRow) || headerRow.length < labels.length) return false;
  return labels.every(
    (label, i) => normalizeCsvHeader(headerRow[i]) === normalizeCsvHeader(label)
  );
}

function buildSiteCsvHeaderIndexMap(headerRow) {
  const map = {};

  // Older export: ID + short incharge labels
  if (headersMatchLabelList(headerRow, SITE_CSV_LEGACY_LABELS)) {
    const legacyKeys = ['id', ...SITE_CSV_COLS.map((c) => c.key)];
    legacyKeys.forEach((key, i) => {
      map[key] = i;
    });
    return map;
  }

  // Older export: ID + unique incharge labels
  if (headersMatchLabelList(headerRow, SITE_CSV_LEGACY_WITH_ID_LABELS)) {
    map.id = 0;
    SITE_CSV_COLS.forEach((c, i) => {
      map[c.key] = i + 1;
    });
    return map;
  }

  // Current export (no ID column)
  if (
    SITE_CSV_COLS.every(
      (c, i) => normalizeCsvHeader(headerRow[i]) === normalizeCsvHeader(c.label)
    )
  ) {
    SITE_CSV_COLS.forEach((c, i) => {
      map[c.key] = i;
    });
    return map;
  }

  const aliases = {
    id: ['id', 'rowid'],
    companyName: ['company name', 'company', 'companyname'],
    siteName: ['site name', 'sitename', 'site'],
    siteAddress: ['address', 'site address'],
    siteCity: ['city', 'site city'],
    siteState: ['state', 'site state'],
    sitePostalCode: ['postal code', 'postalcode', 'pincode', 'pin code', 'site postal code'],
    unitNo: ['unit no', 'unit', 'unit no.', 'unitno'],
    contractorName: ['contractor name', 'contractor'],
    contractorAddress: ['contractor address'],
    contractorEmail: ['contractor email'],
    contractorPhone: ['contractor phone'],
    contractorCity: ['contractor city'],
    contractorState: ['contractor state'],
    inchargeName: ['incharge name', 'name', 'incharge'],
    inchargePhone: ['incharge phone', 'phone'],
    inchargeEmail: ['incharge mail id', 'incharge email', 'mail id', 'email'],
    inchargeDesignation: ['incharge designation', 'designation'],
    industry: ['industry'],
    location: ['location']
  };
  const normalizedHeaders = headerRow.map(normalizeCsvHeader);
  Object.keys(aliases).forEach((key) => {
    const idx = normalizedHeaders.findIndex((h) => aliases[key].includes(h));
    if (idx >= 0) map[key] = idx;
  });
  return map;
}

function sheetToAoaPreferDisplayText(ws) {
  if (!ws || !ws['!ref']) return [];
  const range = XLSX.utils.decode_range(ws['!ref']);
  const rows = [];
  for (let R = range.s.r; R <= range.e.r; R += 1) {
    const row = [];
    for (let C = range.s.c; C <= range.e.c; C += 1) {
      const cell = ws[XLSX.utils.encode_cell({ r: R, c: C })];
      if (!cell) {
        row.push('');
        continue;
      }
      if (cell.w != null && String(cell.w).trim() !== '') {
        row.push(String(cell.w));
        continue;
      }
      row.push(cell.v == null ? '' : cell.v);
    }
    rows.push(row);
  }
  return rows;
}

function parseSiteImportWorkbook(arrayBuffer) {
  const wb = XLSX.read(arrayBuffer, { type: 'array', cellDates: false });
  const sheetName = wb.SheetNames[0];
  if (!sheetName) throw new Error('No sheets found in the file.');
  const ws = wb.Sheets[sheetName];
  const rows = sheetToAoaPreferDisplayText(ws);
  if (!rows.length) throw new Error('No data found in the file.');
  const headerRow = rows[0].map((h) => stripCsvExcelTextPrefix(h));
  const indexMap = buildSiteCsvHeaderIndexMap(headerRow);
  if (indexMap.siteName == null) {
    throw new Error('Could not find a Site Name column. Use Export CSV as a template.');
  }
  const dataRows = rows.slice(1).filter((row) =>
    Array.isArray(row) && row.some((cell) => String(cell ?? '').trim() !== '')
  );
  return dataRows.map((row, rowIndex) => {
    const record = { __row: rowIndex + 2, id: '' };
    if (indexMap.id != null) {
      record.id = stripCsvExcelTextPrefix(row[indexMap.id]);
    }
    SITE_CSV_COLS.forEach(({ key }) => {
      const idx = indexMap[key];
      if (idx == null) {
        record[key] = '';
        return;
      }
      record[key] = stripCsvExcelTextPrefix(row[idx]);
    });
    return record;
  });
}

function resolveCompanyForImport(companyName, companiesList) {
  const name = String(companyName || '').trim().toLowerCase();
  if (!name) return { companyId: '', companyName: '' };
  const match = (companiesList || []).find(
    (c) => String(c.companyName || '').trim().toLowerCase() === name
  );
  if (match) {
    return {
      companyId: String(match.id ?? match.ROWID ?? '').trim(),
      companyName: String(match.companyName || companyName).trim()
    };
  }
  return { companyId: '', companyName: String(companyName || '').trim() };
}

function buildSitePayloadFromImportRow(row, companyResolved) {
  return {
    siteName: String(row.siteName || '').trim(),
    companyId: String(companyResolved.companyId || '').trim(),
    companyName: String(companyResolved.companyName || row.companyName || '').trim(),
    company: String(companyResolved.companyName || row.companyName || '').trim(),
    siteAddress: String(row.siteAddress || '').trim(),
    siteCity: String(row.siteCity || '').trim(),
    siteState: String(row.siteState || '').trim(),
    sitePostalCode: digitsOnly(row.sitePostalCode).slice(0, 6),
    unitNo: String(row.unitNo || '').trim(),
    contractorName: String(row.contractorName || '').trim(),
    contractorAddress: String(row.contractorAddress || '').trim(),
    contractorEmail: String(row.contractorEmail || '').trim(),
    contractorPhone: digitsOnly(row.contractorPhone).slice(0, 10),
    contractorCity: String(row.contractorCity || '').trim(),
    contractorState: String(row.contractorState || '').trim(),
    inchargeName: String(row.inchargeName || '').trim(),
    inchargePhone: digitsOnly(row.inchargePhone).slice(0, 10),
    inchargeEmail: String(row.inchargeEmail || '').trim(),
    inchargeDesignation: String(row.inchargeDesignation || '').trim(),
    industry: normalizeSiteIndustryLabel(row.industry),
    sandERCNumber: '',
    factoryRCNumber: '',
    clraRCNumber: '',
    location: String(row.location || '').trim(),
    audit: 'false'
  };
}

/** Company Details fields shown read-only after a company is selected (from company_function). */
const COMPANY_LINK_DISPLAY_FIELDS = [
  { key: 'companyName', label: 'Company name' },
  { key: 'companyMail', label: 'Mail Id' },
  { key: 'companyPhoneNumber', label: 'Mobile number' },
  { key: 'companyAddress', label: 'Address' },
  { key: 'city', label: 'City' },
  { key: 'state', label: 'State' },
  { key: 'postalcode', label: 'Postal code' },
  { key: 'incorprationDate', label: 'Incorporation date' },
  { key: 'incorporationNo', label: 'Incorporation no.' },
  { key: 'companyPANNumber', label: 'Company PAN' },
  { key: 'gstNo', label: 'GST no.' },
  { key: 'pfNo', label: 'PF no.' },
  { key: 'esiNo', label: 'ESI no.' },
  { key: 'directorName', label: 'Director name' },
  { key: 'directorPhoneNumber', label: 'Director phone' },
  { key: 'directorMail', label: 'Director mail' },
  { key: 'directorAddress', label: 'Director address' },
  { key: 'ownerName', label: 'Owner name' },
  { key: 'ownerPAN', label: 'Owner PAN' },
  { key: 'ownerAaadhar', label: 'Owner Aadhaar' },
  { key: 'ownerDesignation', label: 'Owner designation' },
  { key: 'safetyOfficerName', label: 'Safety officer name' },
  { key: 'safetyOfficerPhone', label: 'Safety officer phone' }
];

function siteCompanyName(s) {
  if (!s || typeof s !== 'object') return '';
  return String(s.companyName ?? s.CompanyName ?? s.company ?? s.Company ?? '').trim();
}

function siteCompanyId(s) {
  if (!s || typeof s !== 'object') return '';
  const id = s.companyId ?? s.CompanyId ?? s.companyROWID ?? '';
  return id == null || id === '' ? '' : String(id);
}

/** Natural key for import upsert: Company Name + Site Name + State (case-insensitive). */
function siteImportNaturalKey(companyName, siteName, siteState) {
  const norm = (v) =>
    String(v ?? '')
      .trim()
      .toLowerCase()
      .replace(/\s+/g, ' ');
  const company = norm(companyName);
  const site = norm(siteName);
  const state = norm(siteState);
  if (!company || !site || !state) return '';
  return `${company}|${site}|${state}`;
}

function siteRecordNaturalKey(s) {
  if (!s || typeof s !== 'object') return '';
  const siteName = String(s.siteName ?? s.SiteName ?? '').trim();
  return siteImportNaturalKey(siteCompanyName(s), siteName, siteStateFromRecord(s));
}

function siteRecordId(s) {
  if (!s || typeof s !== 'object') return '';
  return String(s.id ?? s.ROWID ?? s.rowid ?? '').trim();
}

/** @returns {(number | 'ellipsis')[]} */
function buildPaginationItems(currentPage, totalPages) {
  if (totalPages <= 0) return [];
  if (totalPages <= 9) {
    return Array.from({ length: totalPages }, (_, i) => i + 1);
  }
  const set = new Set([1, totalPages, currentPage]);
  for (let i = currentPage - 2; i <= currentPage + 2; i += 1) {
    if (i >= 1 && i <= totalPages) set.add(i);
  }
  const sorted = [...set].sort((a, b) => a - b);
  const out = [];
  for (let i = 0; i < sorted.length; i += 1) {
    if (i > 0 && sorted[i] - sorted[i - 1] > 1) {
      out.push('ellipsis');
    }
    out.push(sorted[i]);
  }
  return out;
}

function normalizeEmail(value) {
  return String(value || '')
    .trim()
    .toLowerCase();
}

function siteLocation(s) {
  if (!s || typeof s !== 'object') return '';
  return String(s.location ?? s.Location ?? '').trim();
}

function siteSandERCNumber(s) {
  if (!s || typeof s !== 'object') return '';
  return String(s.sandERCNumber ?? s.SandERCNumber ?? '').trim();
}

function siteFactoryRCNumber(s) {
  if (!s || typeof s !== 'object') return '';
  return String(s.factoryRCNumber ?? s.FactoryRCNumber ?? '').trim();
}

function siteCLRARCNumber(s) {
  if (!s || typeof s !== 'object') return '';
  return String(s.clraRCNumber ?? s.CLRARCNumber ?? '').trim();
}

function isClraIndustryLabel(industry) {
  const s = String(industry || '').trim().toLowerCase();
  return s === 'clra' || s.includes('clra') || s.includes('contract labour') || s.includes('contract labor');
}

function isShopsAndEstablishmentIndustryLabel(industry) {
  const s = String(industry || '').trim().toLowerCase();
  return (
    (s.includes('shop') || s.includes('shops')) &&
    s.includes('establishment')
  );
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

/** Only these three Industry values are allowed in Site Management. */
const SITE_INDUSTRY_OPTIONS = ['CLRA', 'Shops and Establishment', 'Factories Act'];

/**
 * Map misspellings / variants (e.g. "Shop and establishment") to the canonical three labels.
 * Unknown values return '' so they are not added as extra dropdown options.
 */
function normalizeSiteIndustryLabel(industry) {
  const raw = String(industry || '').trim();
  if (!raw) return '';
  if (SITE_INDUSTRY_OPTIONS.includes(raw)) return raw;
  if (isClraIndustryLabel(raw)) return 'CLRA';
  if (isShopsAndEstablishmentIndustryLabel(raw)) return 'Shops and Establishment';
  if (isFactoryIndustryLabel(raw)) return 'Factories Act';
  return '';
}

function getSiteRcFieldVisibility(industry) {
  const normalized = normalizeSiteIndustryLabel(industry) || industry;
  if (isClraIndustryLabel(normalized)) {
    return { showSandERC: false, showFactoryRC: false, showClraRC: true };
  }
  if (isShopsAndEstablishmentIndustryLabel(normalized)) {
    return { showSandERC: true, showFactoryRC: false, showClraRC: false };
  }
  if (isFactoryIndustryLabel(normalized)) {
    return { showSandERC: false, showFactoryRC: true, showClraRC: false };
  }
  return { showSandERC: false, showFactoryRC: false, showClraRC: false };
}

/** Avoid opaque `Unexpected token '<'` when the server returns an HTML error page. */
async function readJsonFromResponse(res) {
  const text = await res.text();
  const trimmed = text.trim();
  if (!trimmed) return {};
  if (trimmed[0] === '<') {
    throw new Error(
      `Server returned a web page instead of JSON (HTTP ${res.status}). Redeploy sitemanagement_function with the latest code, or verify /server/sitemanagement_function is reachable.`
    );
  }
  try {
    return JSON.parse(text);
  } catch (e) {
    throw new Error(`Invalid response from server (HTTP ${res.status}).`);
  }
}

const initialForm = {
  siteName: '',
  companyId: '',
  companyName: '',
  siteAddress: '',
  siteCity: '',
  siteState: '',
  sitePostalCode: '',
  unitNo: '',
  contractorName: '',
  contractorAddress: '',
  contractorEmail: '',
  contractorPhone: '',
  contractorCity: '',
  contractorState: '',
  inchargeName: '',
  inchargePhone: '',
  inchargeEmail: '',
  inchargeDesignation: '',
  industry: '',
  sandERCNumber: '',
  factoryRCNumber: '',
  clraRCNumber: '',
  location: '',
  audit: 'false'
};

function digitsOnly(s) {
  return String(s || '').replace(/\D/g, '');
}

const SITE_FORM_EMAIL_REGEX = /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/i;

/** Constrain input while typing (PIN / phone length, email length, reasonable text caps). */
function sanitizeSiteFormField(name, raw) {
  const v = raw == null ? '' : String(raw);
  switch (name) {
    case 'sitePostalCode':
      return digitsOnly(v).slice(0, 6);
    case 'inchargePhone':
    case 'contractorPhone':
      return digitsOnly(v).slice(0, 10);
    case 'inchargeEmail':
    case 'contractorEmail':
      return v.replace(/\s/g, '').slice(0, 254);
    case 'siteName':
    case 'contractorName':
      return v.slice(0, 200);
    case 'siteCity':
    case 'siteState':
    case 'contractorCity':
    case 'contractorState':
      return v.slice(0, 120);
    case 'siteAddress':
    case 'contractorAddress':
      return v.slice(0, 500);
    case 'unitNo':
      return v.slice(0, 80);
    case 'inchargeName':
      return v.slice(0, 200);
    case 'inchargeDesignation':
      return v.slice(0, 150);
    case 'sandERCNumber':
    case 'factoryRCNumber':
    case 'clraRCNumber':
    case 'location':
      return v.slice(0, 200);
    default:
      return v;
  }
}

/** Client-side validation (Add / Edit site modal). */
function validateSiteFormValues(form) {
  const errors = {};
  if (!form.siteName.trim()) errors.siteName = 'Site name is required';
  if (!String(form.companyId || '').trim() && !String(form.companyName || '').trim()) {
    errors.companyId = 'Company is required';
  }
  if (!form.siteCity.trim()) errors.siteCity = 'City is required';
  if (!form.siteState.trim()) errors.siteState = 'State is required';
  const pin = digitsOnly(form.sitePostalCode);
  if (!String(form.sitePostalCode || '').trim()) {
    errors.sitePostalCode = 'Postal code is required';
  } else if (pin.length !== 6) {
    errors.sitePostalCode = 'Enter a valid 6-digit PIN code';
  }
  if (!form.siteAddress.trim()) errors.siteAddress = 'Address is required';
  if (!form.unitNo.trim()) errors.unitNo = 'Unit no. is required';
  if (!form.inchargeName.trim()) errors.inchargeName = 'Name is required';
  if (!String(form.inchargePhone || '').trim()) {
    errors.inchargePhone = 'Phone is required';
  } else if (digitsOnly(form.inchargePhone).length !== 10) {
    errors.inchargePhone = 'Enter a valid 10-digit phone number';
  }
  const em = String(form.inchargeEmail || '').trim();
  if (em && !SITE_FORM_EMAIL_REGEX.test(em)) {
    errors.inchargeEmail = 'Enter a valid email address (e.g., name@company.com)';
  }
  const contractorEmail = String(form.contractorEmail || '').trim();
  if (contractorEmail && !SITE_FORM_EMAIL_REGEX.test(contractorEmail)) {
    errors.contractorEmail = 'Enter a valid email address (e.g., name@company.com)';
  }
  const contractorPhone = String(form.contractorPhone || '').trim();
  if (contractorPhone && digitsOnly(contractorPhone).length !== 10) {
    errors.contractorPhone = 'Enter a valid 10-digit phone number';
  }
  if (!form.inchargeDesignation.trim()) errors.inchargeDesignation = 'Designation is required';
  return errors;
}

const SITE_FORM_BLUR_VALIDATE_NAMES = new Set([
  'siteName',
  'companyId',
  'siteCity',
  'siteState',
  'sitePostalCode',
  'siteAddress',
  'unitNo',
  'contractorEmail',
  'contractorPhone',
  'inchargeName',
  'inchargePhone',
  'inchargeEmail',
  'inchargeDesignation'
]);

function extractSiteIdFromResponse(data) {
  const candidates = [
    data?.data?.id,
    data?.data?.ROWID,
    data?.data?.site?.id,
    data?.data?.site?.ROWID,
    data?.data?.siteDetails?.id,
    data?.data?.siteDetails?.ROWID,
    data?.id,
    data?.ROWID
  ];
  const id = candidates.find((v) => v != null && typeof v !== 'object');
  return id == null ? '' : String(id);
}

const SiteManagement = ({ userEmail, userRole }) => {
  const [sites, setSites] = useState([]);
  const [loading, setLoading] = useState(false);
  const [message, setMessage] = useState('');
  const [showForm, setShowForm] = useState(false);
  const [editingId, setEditingId] = useState(null);
  const [viewOnly, setViewOnly] = useState(false);
  const [form, setForm] = useState(initialForm);
  const formRef = useRef(form);
  formRef.current = form;
  const [formErrors, setFormErrors] = useState({});
  const [tableSearch, setTableSearch] = useState('');
  const [tablePage, setTablePage] = useState(1);
  const [selectedIds, setSelectedIds] = useState(() => new Set());
  const [prioritySiteId, setPrioritySiteId] = useState('');
  const [toast, setToast] = useState('');
  const [bulkDeleting, setBulkDeleting] = useState(false);
  const [importing, setImporting] = useState(false);
  const toastTimerRef = useRef(null);
  const importFileRef = useRef(null);
  /** Companies from company_function for Site → Company link */
  const [companies, setCompanies] = useState([]);
  /** Login email aligned with Catalyst + localStorage (prop alone is often stale vs real sign-in). */
  const [sessionLoginEmail, setSessionLoginEmail] = useState(() => {
    try {
      return (localStorage.getItem('userEmail') || '').trim();
    } catch {
      return '';
    }
  });

  useEffect(() => {
    let cancelled = false;
    (async () => {
      const resolved = await resolveLoginEmailString(userEmail);
      if (!cancelled) setSessionLoginEmail((prev) => resolved || prev);
    })();
    return () => {
      cancelled = true;
    };
  }, [userEmail]);

  const loginEmailNorm = normalizeEmail(sessionLoginEmail || stringifyUserEmail(userEmail));

  const hasInchargeIndustryScope = useMemo(
    () => hasInchargeSiteScope(sites, loginEmailNorm),
    [sites, loginEmailNorm]
  );

  const displaySites = useMemo(
    () => filterSitesForLoginUser(sites, loginEmailNorm, userRole),
    [sites, loginEmailNorm, userRole]
  );

  const scopeSubtitle = useMemo(() => {
    if (!hasInchargeIndustryScope) return '';
    const states = [...new Set(displaySites.map(siteStateFromRecord).filter(Boolean))];
    const industries = [...new Set(displaySites.map(siteIndustry).filter(Boolean))];
    const parts = [];
    if (states.length) parts.push(states.join(', '));
    if (industries.length) parts.push(industries.join(', '));
    return parts.length ? `Showing sites for ${parts.join(' · ')}` : '';
  }, [hasInchargeIndustryScope, displaySites]);

  const filteredSites = useMemo(() => {
    const q = tableSearch.trim().toLowerCase();
    if (!q) return displaySites;
    return displaySites.filter((s) => {
      const hay = [
        s.siteName,
        siteCompanyName(s),
        s.siteAddress,
        s.siteCity,
        s.siteState,
        s.sitePostalCode,
        s.unitNo,
        s.inchargeName,
        s.inchargePhone,
        siteInchargeEmail(s),
        s.inchargeDesignation,
        siteIndustry(s),
        siteLocation(s)
      ]
        .map((x) => String(x || '').toLowerCase())
        .join(' ');
      return hay.includes(q);
    });
  }, [displaySites, tableSearch]);

  const sortedDisplaySites = useMemo(() => {
    const list = [...filteredSites];
    if (!prioritySiteId) return list;
    const idx = list.findIndex((s) => String(s?.id) === String(prioritySiteId));
    if (idx <= 0) return list;
    const [prioritySite] = list.splice(idx, 1);
    list.unshift(prioritySite);
    return list;
  }, [filteredSites, prioritySiteId]);

  const totalTablePages = Math.max(1, Math.ceil(sortedDisplaySites.length / SITE_TABLE_PAGE_SIZE));

  useEffect(() => {
    setTablePage(1);
  }, [tableSearch]);

  useEffect(() => {
    setTablePage((p) => Math.min(p, totalTablePages));
  }, [totalTablePages]);

  const effectiveTablePage = Math.min(tablePage, totalTablePages);

  const pagedSites = useMemo(() => {
    const start = (effectiveTablePage - 1) * SITE_TABLE_PAGE_SIZE;
    return sortedDisplaySites.slice(start, start + SITE_TABLE_PAGE_SIZE);
  }, [sortedDisplaySites, effectiveTablePage]);

  const totalVisible = sortedDisplaySites.length;
  const pageStart = (effectiveTablePage - 1) * SITE_TABLE_PAGE_SIZE;
  const showingFrom = totalVisible === 0 ? 0 : pageStart + 1;
  const showingTo = Math.min(pageStart + pagedSites.length, totalVisible);

  const paginationItems = useMemo(
    () => buildPaginationItems(effectiveTablePage, totalTablePages),
    [effectiveTablePage, totalTablePages]
  );

  const allPageSelected =
    pagedSites.length > 0 && pagedSites.every((s) => selectedIds.has(String(s.id)));
  const somePageSelected = pagedSites.some((s) => selectedIds.has(String(s.id)));

  const toggleSelectAllDisplay = () => {
    if (allPageSelected) {
      setSelectedIds((prev) => {
        const next = new Set(prev);
        pagedSites.forEach((s) => next.delete(String(s.id)));
        return next;
      });
    } else {
      setSelectedIds((prev) => {
        const next = new Set(prev);
        pagedSites.forEach((s) => next.add(String(s.id)));
        return next;
      });
    }
  };

  const toggleRowSelected = (id) => {
    const key = String(id);
    setSelectedIds((prev) => {
      const next = new Set(prev);
      if (next.has(key)) next.delete(key);
      else next.add(key);
      return next;
    });
  };

  const handleBulkDeleteSelected = async () => {
    const ids = [...selectedIds];
    if (ids.length === 0) return;
    if (!window.confirm(`Delete ${ids.length} selected site${ids.length > 1 ? 's' : ''}?`)) return;
    setBulkDeleting(true);
    setMessage('');
    try {
      const results = await Promise.all(
        ids.map(async (id) => {
          const res = await fetch(`${API_BASE}/sitemanagement/${id}`, { method: 'DELETE' });
          const data = await readJsonFromResponse(res);
          return data.status === 'success';
        })
      );
      const failedCount = results.filter((ok) => !ok).length;
      if (failedCount > 0) {
        setMessage(`${failedCount} delete request(s) failed. Please try again.`);
      } else {
        setMessage('');
        showToast('Deletedd Successfully');
      }
      setSelectedIds(new Set());
      fetchSites({ clearMessage: failedCount > 0 });
    } catch (err) {
      setMessage(err.message || 'Delete failed');
    } finally {
      setBulkDeleting(false);
    }
  };

  const exportSitesCsv = () => {
    /** Site export: main form fields (audit omitted); all rows in scope (ignores table search). */
    const CSV_EXCEL_TEXT_KEYS = new Set([
      'sitePostalCode',
      'contractorPhone',
      'contractorEmail',
      'inchargePhone',
      'inchargeEmail',
      'unitNo'
    ]);
    const escCsvCell = (raw, colKey) => {
      const s = String(raw ?? '').replace(/"/g, '""');
      if (s === '') return '""';
      const body = CSV_EXCEL_TEXT_KEYS.has(colKey) ? `\t${s}` : s;
      return `"${body}"`;
    };
    const escLabel = (v) => {
      const s = String(v ?? '').replace(/"/g, '""');
      return `"${s}"`;
    };
    const cellValue = (site, key) => {
      if (key === 'companyName') return siteCompanyName(site) || site.companyName || '';
      if (key === 'inchargeEmail') return siteInchargeEmail(site) || site.inchargeEmail || '';
      if (key === 'industry') return siteIndustry(site) || site.industry || '';
      if (key === 'location') return siteLocation(site) || site.location || '';
      if (key === 'sitePostalCode') {
        return String(site.sitePostalCode ?? site.SitePostalCode ?? site.site_postal_code ?? '').trim();
      }
      return site[key];
    };
    const strCmp = (a, b, field) =>
      String(a[field] ?? '')
        .toLowerCase()
        .localeCompare(String(b[field] ?? '').toLowerCase(), undefined, { sensitivity: 'base' });
    const list = [...displaySites].sort((a, b) => strCmp(a, b, 'siteName'));
    const header = SITE_CSV_COLS.map((c) => escLabel(c.label)).join(',');
    const rows = list.map((site) => SITE_CSV_COLS.map((c) => escCsvCell(cellValue(site, c.key), c.key)).join(','));
    const csv = '\ufeff' + [header, ...rows].join('\r\n');
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `site-management-${new Date().toISOString().slice(0, 10)}.csv`;
    a.click();
    URL.revokeObjectURL(url);
  };

  const emptyListMessage = hasInchargeIndustryScope
    ? 'No sites are assigned to your login for this state and industry.'
    : userRole && userRole !== 'App Administrator' && userRole !== 'HR Admin'
      ? 'No sites are assigned to your login as Incharge.'
      : 'No sites. Click Add to create one.';

  const fetchSites = useCallback(async (options = {}) => {
    const { clearMessage = true, useCacheFirst = false, silentRefresh = false } = options;

    if (useCacheFirst) {
      try {
        const cached = localStorage.getItem('siteManagementData');
        if (cached) {
          const parsed = JSON.parse(cached);
          if (Array.isArray(parsed) && parsed.length > 0) {
            setSites(parsed);
          }
        }
      } catch (_) {}
    }

    if (!silentRefresh) setLoading(true);
    if (clearMessage) setMessage('');
    try {
      const res = await fetch(`${API_BASE}/sitemanagement`);
      const data = await readJsonFromResponse(res);
      if (data.status === 'success' && data.data && Array.isArray(data.data.siteDetails)) {
        setSites(data.data.siteDetails);
        localStorage.setItem('siteManagementData', JSON.stringify(data.data.siteDetails));
      } else {
        setSites([]);
      }
    } catch (err) {
      setMessage('Failed to load site details');
      setSites([]);
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    // Show cached rows immediately on navigation, refresh latest data in background.
    fetchSites({ useCacheFirst: true, silentRefresh: true });
  }, [fetchSites]);

  useEffect(() => {
    const valid = new Set(sites.map((s) => String(s.id)));
    setSelectedIds((prev) => {
      const next = new Set();
      prev.forEach((id) => {
        if (valid.has(id)) next.add(id);
      });
      return next;
    });
  }, [sites]);

  const fetchCompanies = useCallback(async () => {
    try {
      const res = await fetch(`${COMPANY_API}/company`, { cache: 'no-store' });
      const data = await res.json().catch(() => ({}));
      const list = Array.isArray(data?.data?.companyDetails) ? data.data.companyDetails : [];
      const sorted = [...list].sort((a, b) =>
        String(a.companyName || '')
          .toLowerCase()
          .localeCompare(String(b.companyName || '').toLowerCase(), undefined, { sensitivity: 'base' })
      );
      setCompanies(sorted);
    } catch (_) {
      setCompanies([]);
    }
  }, []);

  useEffect(() => {
    fetchCompanies();
  }, [fetchCompanies]);

  const selectedCompanyDetails = useMemo(() => {
    const id = String(form.companyId || '').trim();
    if (id) {
      const byId = companies.find((c) => String(c.id) === id);
      if (byId) return byId;
    }
    const name = String(form.companyName || '').trim().toLowerCase();
    if (!name) return null;
    return companies.find((c) => String(c.companyName || '').trim().toLowerCase() === name) || null;
  }, [companies, form.companyId, form.companyName]);

  /** When Site.Company has only the name, match dropdown id from company_function list. */
  useEffect(() => {
    if (!showForm) return;
    const name = String(form.companyName || '').trim();
    if (!name || !companies.length) return;
    const match = companies.find(
      (c) => String(c.companyName || '').trim().toLowerCase() === name.toLowerCase()
    );
    if (!match) return;
    const matchId = String(match.id);
    if (String(form.companyId || '') === matchId) return;
    setForm((prev) => ({
      ...prev,
      companyId: matchId,
      companyName: match.companyName || prev.companyName
    }));
  }, [showForm, companies, form.companyId, form.companyName]);

  /** Company dropdown options; keep legacy linked name if company list has not loaded that id yet */
  const companySelectOptions = useMemo(() => {
    const list = [...companies];
    const curId = String(form.companyId || '').trim();
    const curName = String(form.companyName || '').trim();
    if (curId && !list.some((c) => String(c.id) === curId)) {
      list.unshift({ id: curId, companyName: curName || `Company #${curId}` });
    }
    return list;
  }, [companies, form.companyId, form.companyName]);

  /** Industry dropdown: only the three canonical values (no Checklist Master spelling variants). */
  const industrySelectOptions = SITE_INDUSTRY_OPTIONS;

  const rcFieldVisibility = useMemo(
    () => getSiteRcFieldVisibility(form.industry),
    [form.industry]
  );

  /** City list — same model as Company Details: master list + saved site values + current form. */
  const citySelectOptions = useMemo(() => {
    const set = new Set(INDIAN_CITIES);
    sites.forEach((s) => {
      const siteCity = String(s.siteCity ?? s.SiteCity ?? '').trim();
      const contractorCity = String(s.contractorCity ?? s.ContractorCity ?? '').trim();
      if (siteCity) set.add(siteCity);
      if (contractorCity) set.add(contractorCity);
    });
    [form.siteCity, form.contractorCity].forEach((c) => {
      const v = String(c || '').trim();
      if (v) set.add(v);
    });
    return [...set].sort((a, b) => a.localeCompare(b, undefined, { sensitivity: 'base' }));
  }, [sites, form.siteCity, form.contractorCity]);

  /** State list — same model as Company Details: master list + saved site values + current form. */
  const stateSelectOptions = useMemo(() => {
    const set = new Set(INDIAN_STATES);
    sites.forEach((s) => {
      const siteState = String(s.siteState ?? s.SiteState ?? '').trim();
      const contractorState = String(s.contractorState ?? s.ContractorState ?? '').trim();
      if (siteState) set.add(siteState);
      if (contractorState) set.add(contractorState);
    });
    [form.siteState, form.contractorState].forEach((st) => {
      const v = String(st || '').trim();
      if (v) set.add(v);
    });
    return [...set].sort((a, b) => a.localeCompare(b, undefined, { sensitivity: 'base' }));
  }, [sites, form.siteState, form.contractorState]);

  const openAdd = () => {
    setViewOnly(false);
    setEditingId(null);
    setForm(
      hasInchargeIndustryScope && (sessionLoginEmail || stringifyUserEmail(userEmail))
        ? {
            ...initialForm,
            inchargeEmail: (sessionLoginEmail || stringifyUserEmail(userEmail)).trim(),
          }
        : initialForm
    );
    setFormErrors({});
    fetchCompanies();
    setShowForm(true);
  };

  const openView = (site) => {
    setViewOnly(true);
    setEditingId(site.id);
    setForm({
      siteName: site.siteName || '',
      companyId: siteCompanyId(site),
      companyName: siteCompanyName(site),
      siteAddress: site.siteAddress || '',
      siteCity: site.siteCity || '',
      siteState: site.siteState || '',
      sitePostalCode: site.sitePostalCode || '',
      unitNo: site.unitNo || '',
      contractorName: site.contractorName || '',
      contractorAddress: site.contractorAddress || '',
      contractorEmail: site.contractorEmail || '',
      contractorPhone: site.contractorPhone || '',
      contractorCity: site.contractorCity || '',
      contractorState: site.contractorState || '',
      inchargeName: site.inchargeName || '',
      inchargePhone: site.inchargePhone || '',
      inchargeEmail: siteInchargeEmail(site) || site.inchargeEmail || '',
      inchargeDesignation: site.inchargeDesignation || '',
      industry: normalizeSiteIndustryLabel(siteIndustry(site) || site.industry || ''),
      sandERCNumber: siteSandERCNumber(site) || site.sandERCNumber || '',
      factoryRCNumber: siteFactoryRCNumber(site) || site.factoryRCNumber || '',
      clraRCNumber: siteCLRARCNumber(site) || site.clraRCNumber || '',
      location: siteLocation(site) || site.location || '',
      audit: site.audit === true || site.audit === 'true' ? 'true' : 'false'
    });
    setFormErrors({});
    fetchCompanies();
    setShowForm(true);
  };

  const openEdit = (site) => {
    setViewOnly(false);
    setEditingId(site.id);
    setForm({
      siteName: site.siteName || '',
      companyId: siteCompanyId(site),
      companyName: siteCompanyName(site),
      siteAddress: site.siteAddress || '',
      siteCity: site.siteCity || '',
      siteState: site.siteState || '',
      sitePostalCode: site.sitePostalCode || '',
      unitNo: site.unitNo || '',
      contractorName: site.contractorName || '',
      contractorAddress: site.contractorAddress || '',
      contractorEmail: site.contractorEmail || '',
      contractorPhone: site.contractorPhone || '',
      contractorCity: site.contractorCity || '',
      contractorState: site.contractorState || '',
      inchargeName: site.inchargeName || '',
      inchargePhone: site.inchargePhone || '',
      inchargeEmail: siteInchargeEmail(site) || site.inchargeEmail || '',
      inchargeDesignation: site.inchargeDesignation || '',
      industry: normalizeSiteIndustryLabel(siteIndustry(site) || site.industry || ''),
      sandERCNumber: siteSandERCNumber(site) || site.sandERCNumber || '',
      factoryRCNumber: siteFactoryRCNumber(site) || site.factoryRCNumber || '',
      clraRCNumber: siteCLRARCNumber(site) || site.clraRCNumber || '',
      location: siteLocation(site) || site.location || '',
      audit: site.audit === true || site.audit === 'true' ? 'true' : 'false'
    });
    setFormErrors({});
    fetchCompanies();
    setShowForm(true);
  };

  const closeForm = () => {
    setShowForm(false);
    setViewOnly(false);
    setEditingId(null);
    setForm(initialForm);
    setFormErrors({});
  };

  const dismissToast = useCallback(() => {
    if (toastTimerRef.current) {
      clearTimeout(toastTimerRef.current);
      toastTimerRef.current = null;
    }
    setToast('');
  }, []);

  const showToast = useCallback((text) => {
    if (!text) return;
    if (toastTimerRef.current) {
      clearTimeout(toastTimerRef.current);
      toastTimerRef.current = null;
    }
    setToast(text);
    toastTimerRef.current = window.setTimeout(() => {
      setToast('');
      toastTimerRef.current = null;
    }, 4000);
  }, []);

  const handleImportSites = useCallback(
    async (event) => {
      const file = event.target.files?.[0];
      if (importFileRef.current) importFileRef.current.value = '';
      if (!file) return;

      const name = String(file.name || '').toLowerCase();
      const okExt = name.endsWith('.csv') || name.endsWith('.xlsx') || name.endsWith('.xls');
      if (!okExt) {
        setMessage('Invalid file type. Please upload a CSV or Excel file (.csv, .xlsx, .xls).');
        return;
      }
      if (file.size > 5 * 1024 * 1024) {
        setMessage('File size exceeds 5MB limit.');
        return;
      }

      setImporting(true);
      setMessage('');
      try {
        let companiesList = companies;
        if (!companiesList.length) {
          try {
            const res = await fetch(`${COMPANY_API}/company`, { cache: 'no-store' });
            const data = await res.json().catch(() => ({}));
            companiesList = Array.isArray(data?.data?.companyDetails) ? data.data.companyDetails : [];
            if (companiesList.length) setCompanies(companiesList);
          } catch (_) {
            companiesList = [];
          }
        }

        const arrayBuffer = await file.arrayBuffer();
        const importRows = parseSiteImportWorkbook(arrayBuffer);
        if (!importRows.length) {
          setMessage('No data rows found in the file.');
          return;
        }

        // Refresh current sites so upsert matches latest DB rows (not a stale client list)
        let sitesForMatch = sites;
        try {
          const res = await fetch(`${API_BASE}/sitemanagement`);
          const data = await readJsonFromResponse(res);
          if (data.status === 'success' && Array.isArray(data.data?.siteDetails)) {
            sitesForMatch = data.data.siteDetails;
            setSites(sitesForMatch);
          }
        } catch (_) {
          // keep in-memory sites
        }

        const existingById = new Map();
        const existingByNaturalKey = new Map();
        sitesForMatch.forEach((s) => {
          const id = siteRecordId(s);
          if (id) existingById.set(id, s);
          const key = siteRecordNaturalKey(s);
          // Keep the first match when duplicates already exist
          if (key && !existingByNaturalKey.has(key)) existingByNaturalKey.set(key, s);
        });

        let created = 0;
        let updated = 0;
        const errors = [];

        for (const row of importRows) {
          const companyResolved = resolveCompanyForImport(row.companyName, companiesList);
          const companyNameForMatch = companyResolved.companyName || row.companyName;
          const formValues = {
            ...initialForm,
            siteName: row.siteName,
            companyId: companyResolved.companyId,
            companyName: companyNameForMatch,
            siteAddress: row.siteAddress,
            siteCity: row.siteCity,
            siteState: row.siteState,
            sitePostalCode: row.sitePostalCode,
            unitNo: row.unitNo,
            contractorName: row.contractorName,
            contractorAddress: row.contractorAddress,
            contractorEmail: row.contractorEmail,
            contractorPhone: row.contractorPhone,
            contractorCity: row.contractorCity,
            contractorState: row.contractorState,
            inchargeName: row.inchargeName,
            inchargePhone: row.inchargePhone,
            inchargeEmail: row.inchargeEmail,
            inchargeDesignation: row.inchargeDesignation,
            industry: normalizeSiteIndustryLabel(row.industry),
            location: row.location
          };
          const errs = validateSiteFormValues(formValues);
          if (Object.keys(errs).length > 0) {
            errors.push(`Row ${row.__row}: ${Object.values(errs)[0]}`);
            continue;
          }
          if (!companyNameForMatch) {
            errors.push(`Row ${row.__row}: Company is required`);
            continue;
          }

          const payload = buildSitePayloadFromImportRow(row, companyResolved);
          const rowId = String(row.id || '').trim();
          const naturalKey = siteImportNaturalKey(companyNameForMatch, row.siteName, row.siteState);
          let matchId = '';
          if (rowId && existingById.has(rowId)) {
            matchId = rowId;
          } else if (naturalKey && existingByNaturalKey.has(naturalKey)) {
            matchId = siteRecordId(existingByNaturalKey.get(naturalKey));
          }

          try {
            if (matchId) {
              const res = await fetch(`${API_BASE}/sitemanagement/${encodeURIComponent(matchId)}`, {
                method: 'PUT',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(payload)
              });
              const data = await readJsonFromResponse(res);
              if (data.status === 'success') {
                updated += 1;
                const updatedSite = {
                  ...(existingById.get(matchId) || existingByNaturalKey.get(naturalKey) || {}),
                  ...payload,
                  id: matchId
                };
                existingById.set(matchId, updatedSite);
                if (naturalKey) existingByNaturalKey.set(naturalKey, updatedSite);
              } else {
                errors.push(`Row ${row.__row}: ${data.message || 'Update failed'}`);
              }
            } else {
              const res = await fetch(`${API_BASE}/sitemanagement`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(payload)
              });
              const data = await readJsonFromResponse(res);
              if (data.status === 'success') {
                created += 1;
                const createdId = extractSiteIdFromResponse(data);
                const createdSite = { ...payload, id: createdId };
                if (createdId) existingById.set(createdId, createdSite);
                // Prevent later rows in the same file from creating another duplicate
                if (naturalKey) existingByNaturalKey.set(naturalKey, createdSite);
              } else {
                errors.push(`Row ${row.__row}: ${data.message || 'Add failed'}`);
              }
            }
          } catch (err) {
            errors.push(`Row ${row.__row}: ${err.message || 'Request failed'}`);
          }
        }

        await fetchSites({ clearMessage: false });
        const parts = [];
        if (created) parts.push(`${created} added`);
        if (updated) parts.push(`${updated} updated`);
        if (parts.length) showToast(`Import complete: ${parts.join(', ')}`);
        if (errors.length) {
          const preview = errors.slice(0, 5).join(' · ');
          const more = errors.length > 5 ? ` (+${errors.length - 5} more)` : '';
          setMessage(
            parts.length
              ? `Imported with ${errors.length} error(s): ${preview}${more}`
              : `Import failed: ${preview}${more}`
          );
        } else if (!parts.length) {
          setMessage('No sites were imported.');
        }
      } catch (err) {
        setMessage(err.message || 'Import failed');
      } finally {
        setImporting(false);
      }
    },
    [companies, sites, fetchSites, showToast]
  );

  useEffect(
    () => () => {
      if (toastTimerRef.current) clearTimeout(toastTimerRef.current);
    },
    []
  );

  const handleChange = (e) => {
    const { name, value } = e.target;
    if (name === 'companyId') {
      const company = companies.find((c) => String(c.id) === String(value));
      setForm((prev) => ({
        ...prev,
        companyId: value,
        companyName: company?.companyName || ''
      }));
      setFormErrors((prev) => {
        if (!prev.companyId && !prev.companyName) return prev;
        const next = { ...prev };
        delete next.companyId;
        delete next.companyName;
        return next;
      });
      return;
    }
    const nextValue =
      name === 'industry' ? normalizeSiteIndustryLabel(value) : sanitizeSiteFormField(name, value);
    setForm((prev) => {
      const next = { ...prev, [name]: nextValue };
      if (name === 'industry') {
        const vis = getSiteRcFieldVisibility(nextValue);
        if (!vis.showSandERC) next.sandERCNumber = '';
        if (!vis.showFactoryRC) next.factoryRCNumber = '';
        if (!vis.showClraRC) next.clraRCNumber = '';
      }
      return next;
    });
    setFormErrors((prev) => {
      if (!prev[name]) return prev;
      const next = { ...prev };
      delete next[name];
      return next;
    });
  };

  /** Show field error after leaving a mandatory control empty or invalid (same rules as submit). */
  const handleSiteFieldBlur = (e) => {
    if (viewOnly) return;
    const name = e.target?.name;
    if (!name || !SITE_FORM_BLUR_VALIDATE_NAMES.has(name)) return;
    const fieldErr = validateSiteFormValues(formRef.current)[name];
    setFormErrors((fe) => {
      const next = { ...fe };
      if (fieldErr) next[name] = fieldErr;
      else delete next[name];
      return next;
    });
  };

  const handleSubmit = async (e) => {
    e.preventDefault();
    setMessage('');
    if (viewOnly) return;
    const errors = validateSiteFormValues(form);
    setFormErrors(errors);
    if (Object.keys(errors).length > 0) return;
    try {
      const payload = {
        siteName: form.siteName.trim(),
        companyId: String(form.companyId || '').trim(),
        companyName: form.companyName.trim(),
        company: form.companyName.trim(),
        siteAddress: form.siteAddress.trim(),
        siteCity: form.siteCity.trim(),
        siteState: form.siteState.trim(),
        sitePostalCode: form.sitePostalCode.trim(),
        unitNo: form.unitNo.trim(),
        contractorName: form.contractorName.trim(),
        contractorAddress: form.contractorAddress.trim(),
        contractorEmail: form.contractorEmail.trim(),
        contractorPhone: form.contractorPhone.trim(),
        contractorCity: form.contractorCity.trim(),
        contractorState: form.contractorState.trim(),
        inchargeName: form.inchargeName.trim(),
        inchargePhone: form.inchargePhone.trim(),
        inchargeEmail: form.inchargeEmail.trim(),
        inchargeDesignation: form.inchargeDesignation.trim(),
        industry: normalizeSiteIndustryLabel(form.industry),
        sandERCNumber: rcFieldVisibility.showSandERC ? form.sandERCNumber.trim() : '',
        factoryRCNumber: rcFieldVisibility.showFactoryRC ? form.factoryRCNumber.trim() : '',
        clraRCNumber: rcFieldVisibility.showClraRC ? form.clraRCNumber.trim() : '',
        location: form.location.trim(),
        audit: form.audit || 'false'
      };

      if (editingId) {
        const res = await fetch(`${API_BASE}/sitemanagement/${editingId}`, {
          method: 'PUT',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(payload)
        });
        const data = await readJsonFromResponse(res);
        if (data.status === 'success') {
          const successText = 'Update Successfully';
          setMessage('');
          showToast(successText);
          setPrioritySiteId(String(editingId));
          closeForm();
          fetchSites({ clearMessage: false });
        } else {
          setMessage(data.message || 'Update failed');
        }
      } else {
        const res = await fetch(`${API_BASE}/sitemanagement`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(payload)
        });
        const data = await readJsonFromResponse(res);
        if (data.status === 'success') {
          const successText = 'Added Successfully';
          setMessage('');
          showToast(successText);
          const createdId = extractSiteIdFromResponse(data);
          if (createdId) setPrioritySiteId(createdId);
          closeForm();
          fetchSites({ clearMessage: false });
        } else {
          setMessage(data.message || 'Add failed');
        }
      }
    } catch (err) {
      setMessage(err.message || 'Request failed');
    }
  };

  const handleDelete = async (id) => {
    if (!window.confirm('Delete this site?')) return;
    setMessage('');
    try {
      const res = await fetch(`${API_BASE}/sitemanagement/${id}`, { method: 'DELETE' });
      const data = await readJsonFromResponse(res);
      if (data.status === 'success') {
        const successText = 'Deletedd Successfully';
        setMessage('');
        showToast(successText);
        if (String(id) === String(prioritySiteId)) setPrioritySiteId('');
        fetchSites({ clearMessage: false });
      } else {
        setMessage(data.message || 'Delete failed');
      }
    } catch (err) {
      setMessage(err.message || 'Delete failed');
    }
  };

  const subtitle = scopeSubtitle;

  const TABLE_COL_COUNT = 14;

  const siteFormModalTitle = viewOnly ? 'View Site' : editingId ? 'Edit Site' : 'Add Site';

  return (
    <div className={`company-details-page site-management-hcm-route${showForm ? ' company-details-page-form-open' : ''}`}>
      {toast ? (
        <div className="company-details-toast-wrap" role="status" aria-live="polite" aria-atomic="true">
          <div className="company-details-toast">
            <CheckCircle2 className="company-details-toast-icon" size={22} strokeWidth={2} aria-hidden />
            <span className="company-details-toast-text">{toast}</span>
            <button type="button" className="company-details-toast-dismiss" onClick={dismissToast} aria-label="Dismiss notification">
              ×
            </button>
          </div>
        </div>
      ) : null}
      {showForm ? (
        <>
          <header className="company-details-page-heading">
            <h1 className="company-details-page-title">Site Management</h1>
            <nav className="company-details-breadcrumb" aria-label="Breadcrumb">
              <Link to="/hcm-dashboard" className="company-details-bc-link">
                Dashboard
              </Link>
              <span className="company-details-bc-sep" aria-hidden>
                ›
              </span>
              <span className="company-details-bc-muted">Organization Master</span>
              <span className="company-details-bc-sep" aria-hidden>
                ›
              </span>
              <button
                type="button"
                className="company-details-bc-link company-details-bc-as-link"
                onClick={closeForm}
                aria-label="Return to site list"
              >
                Site Management
              </button>
              <span className="company-details-bc-sep" aria-hidden>
                ›
              </span>
              <span className="company-details-bc-current">{siteFormModalTitle}</span>
            </nav>
          </header>
          <div className="company-details-form-center">
            {message ? <div className="company-details-message company-details-message--modal">{message}</div> : null}
            <div className="company-details-modal" role="dialog" aria-labelledby="site-management-modal-title" aria-modal="true">
              <div className="company-details-modal-header">
                <h2 id="site-management-modal-title" className="company-details-modal-title">
                  {siteFormModalTitle}
                </h2>
                <button type="button" className="company-details-modal-close" onClick={closeForm} aria-label="Close">
                  ×
                </button>
              </div>
              <form onSubmit={viewOnly ? (e) => e.preventDefault() : handleSubmit} className="company-details-modal-form" noValidate>
                <div className="company-details-modal-scroll">
                  <div className="company-details-modal-body">
                    <section className="company-details-section-card">
                      <header className="company-details-section-head">
                        <h3 className="company-details-section-title">
                          <svg
                            className="company-details-section-title-icon"
                            width="20"
                            height="20"
                            viewBox="0 0 24 24"
                            fill="none"
                            stroke="currentColor"
                            strokeWidth="2"
                            aria-hidden
                          >
                            <path d="M21 10c0 7-9 13-9 13s-9-6-9-13a9 9 0 0 1 18 0z" />
                            <circle cx="12" cy="10" r="3" />
                          </svg>
                          Site details
                        </h3>
                      </header>
                      <div className="company-details-fields-grid">
                        <div className="company-details-field">
                          <label htmlFor="sm-companyId">
                            Company <span className="required" aria-hidden="true">*</span>
                          </label>
                          <select
                            id="sm-companyId"
                            name="companyId"
                            value={form.companyId}
                            onChange={handleChange}
                            onBlur={handleSiteFieldBlur}
                            disabled={viewOnly}
                            required
                            aria-describedby={formErrors.companyId ? 'sm-companyId-error' : undefined}
                          >
                            <option value="">Select company</option>
                            {companySelectOptions.map((co) => (
                              <option key={String(co.id)} value={String(co.id)}>
                                {co.companyName || `Company #${co.id}`}
                              </option>
                            ))}
                          </select>
                          {formErrors.companyId && (
                            <div id="sm-companyId-error" style={{ color: 'red', fontSize: '0.95em', marginTop: 2 }}>
                              {formErrors.companyId}
                            </div>
                          )}
                        </div>
                        <div className="company-details-field">
                          <label htmlFor="sm-siteName">
                            Site name <span className="required" aria-hidden="true">*</span>
                          </label>
                          <input
                            id="sm-siteName"
                            name="siteName"
                            value={form.siteName}
                            onChange={handleChange}
                            onBlur={handleSiteFieldBlur}
                            placeholder="Enter site name"
                            disabled={viewOnly}
                            maxLength={200}
                            required
                          />
                          {formErrors.siteName && (
                            <div style={{ color: 'red', fontSize: '0.95em', marginTop: 2 }}>{formErrors.siteName}</div>
                          )}
                        </div>
                        <div className="company-details-field">
                          <label htmlFor="sm-siteAddress">
                            Address <span className="required" aria-hidden="true">*</span>
                          </label>
                          <input
                            id="sm-siteAddress"
                            name="siteAddress"
                            value={form.siteAddress}
                            onChange={handleChange}
                            onBlur={handleSiteFieldBlur}
                            placeholder="Enter address"
                            disabled={viewOnly}
                            maxLength={500}
                            required
                          />
                          {formErrors.siteAddress && (
                            <div style={{ color: 'red', fontSize: '0.95em', marginTop: 2 }}>{formErrors.siteAddress}</div>
                          )}
                        </div>
                        <div className="company-details-field">
                          <label htmlFor="sm-siteCity">
                            City <span className="required" aria-hidden="true">*</span>
                          </label>
                          <CityCombobox
                            id="sm-siteCity"
                            name="siteCity"
                            value={form.siteCity}
                            options={citySelectOptions}
                            onChange={handleChange}
                            onBlur={handleSiteFieldBlur}
                            disabled={viewOnly}
                            required
                            placeholder="Type letter to filter cities"
                            otherPlaceholder="Enter your city name"
                          />
                          {formErrors.siteCity && (
                            <div style={{ color: 'red', fontSize: '0.95em', marginTop: 2 }}>{formErrors.siteCity}</div>
                          )}
                        </div>
                        <div className="company-details-field">
                          <label htmlFor="sm-siteState">
                            State <span className="required" aria-hidden="true">*</span>
                          </label>
                          <StateCombobox
                            id="sm-siteState"
                            name="siteState"
                            value={form.siteState}
                            options={stateSelectOptions}
                            onChange={handleChange}
                            onBlur={handleSiteFieldBlur}
                            disabled={viewOnly}
                            required
                            placeholder="Type letter to filter states"
                            otherPlaceholder="Enter your state name"
                          />
                          {formErrors.siteState && (
                            <div style={{ color: 'red', fontSize: '0.95em', marginTop: 2 }}>{formErrors.siteState}</div>
                          )}
                        </div>
                        <div className="company-details-field">
                          <label htmlFor="sm-sitePostalCode">
                            Postal code <span className="required" aria-hidden="true">*</span>
                          </label>
                          <input
                            id="sm-sitePostalCode"
                            name="sitePostalCode"
                            type="tel"
                            inputMode="numeric"
                            value={form.sitePostalCode}
                            onChange={handleChange}
                            onBlur={handleSiteFieldBlur}
                            placeholder="Enter 6-digit PIN"
                            disabled={viewOnly}
                            maxLength={6}
                            title="Enter exactly 6 digits"
                            required
                          />
                          {formErrors.sitePostalCode && (
                            <div style={{ color: 'red', fontSize: '0.95em', marginTop: 2 }}>{formErrors.sitePostalCode}</div>
                          )}
                        </div>
                        <div className="company-details-field">
                          <label htmlFor="sm-unitNo">
                            Unit no. <span className="required" aria-hidden="true">*</span>
                          </label>
                          <input
                            id="sm-unitNo"
                            name="unitNo"
                            value={form.unitNo}
                            onChange={handleChange}
                            onBlur={handleSiteFieldBlur}
                            placeholder="Enter unit no."
                            disabled={viewOnly}
                            maxLength={80}
                            required
                          />
                          {formErrors.unitNo && (
                            <div style={{ color: 'red', fontSize: '0.95em', marginTop: 2 }}>{formErrors.unitNo}</div>
                          )}
                        </div>
                        <div className="company-details-field">
                          <label htmlFor="sm-industry">Industry</label>
                          <select
                            id="sm-industry"
                            name="industry"
                            value={form.industry}
                            onChange={handleChange}
                            disabled={viewOnly}
                          >
                            <option value="">Select</option>
                            {industrySelectOptions.map((sector) => (
                              <option key={sector} value={sector}>
                                {sector}
                              </option>
                            ))}
                          </select>
                        </div>
                        {rcFieldVisibility.showSandERC && (
                          <div className="company-details-field">
                            <label htmlFor="sm-sandERCNumber">Shops and Establishment RC number</label>
                            <input
                              id="sm-sandERCNumber"
                              name="sandERCNumber"
                              value={form.sandERCNumber}
                              onChange={handleChange}
                              placeholder="Enter Shops and Establishment RC number"
                              disabled={viewOnly}
                              maxLength={200}
                            />
                          </div>
                        )}
                        {rcFieldVisibility.showFactoryRC && (
                          <div className="company-details-field">
                            <label htmlFor="sm-factoryRCNumber">Factory RC number</label>
                            <input
                              id="sm-factoryRCNumber"
                              name="factoryRCNumber"
                              value={form.factoryRCNumber}
                              onChange={handleChange}
                              placeholder="Enter Factory RC number"
                              disabled={viewOnly}
                              maxLength={200}
                            />
                          </div>
                        )}
                        {rcFieldVisibility.showClraRC && (
                          <div className="company-details-field">
                            <label htmlFor="sm-clraRCNumber">CLRA RC number</label>
                            <input
                              id="sm-clraRCNumber"
                              name="clraRCNumber"
                              value={form.clraRCNumber}
                              onChange={handleChange}
                              placeholder="Enter CLRA RC number"
                              disabled={viewOnly}
                              maxLength={200}
                            />
                          </div>
                        )}
                        <div className="company-details-field">
                          <label htmlFor="sm-location">Location</label>
                          <input
                            id="sm-location"
                            name="location"
                            value={form.location}
                            onChange={handleChange}
                            placeholder="Enter location"
                            disabled={viewOnly}
                            maxLength={200}
                          />
                        </div>
                      </div>
                    </section>

                    {form.companyId ? (
                      <section className="company-details-section-card">
                        <header className="company-details-section-head">
                          <h3 className="company-details-section-title">
                            <svg
                              className="company-details-section-title-icon"
                              width="20"
                              height="20"
                              viewBox="0 0 24 24"
                              fill="none"
                              stroke="currentColor"
                              strokeWidth="2"
                              aria-hidden
                            >
                              <path d="M3 21h18" />
                              <path d="M5 21V7l7-4 7 4v14" />
                              <path d="M9 21v-6h6v6" />
                            </svg>
                            Company details
                          </h3>
                        </header>
                        <div className="company-details-fields-grid">
                          {COMPANY_LINK_DISPLAY_FIELDS.map(({ key, label }) => {
                            const raw =
                              selectedCompanyDetails?.[key] ??
                              (key === 'companyName' ? form.companyName : '');
                            const display = String(raw ?? '').trim() || '—';
                            return (
                              <div
                                key={key}
                                className={`company-details-field${
                                  key === 'companyAddress' || key === 'directorAddress'
                                    ? ' company-details-field--full-row'
                                    : ''
                                }`}
                              >
                                <label htmlFor={`sm-co-${key}`}>{label}</label>
                                <input
                                  id={`sm-co-${key}`}
                                  value={display}
                                  readOnly
                                  disabled
                                  title={display === '—' ? undefined : display}
                                />
                              </div>
                            );
                          })}
                        </div>
                        {!selectedCompanyDetails && form.companyName ? (
                          <p style={{ margin: '8px 0 0', fontSize: '0.9em', color: '#666' }}>
                            Linked company: {form.companyName}. Full details will appear once company list loads.
                          </p>
                        ) : null}
                      </section>
                    ) : null}

                    <section className="company-details-section-card">
                      <header className="company-details-section-head">
                        <h3 className="company-details-section-title">
                          <svg
                            className="company-details-section-title-icon"
                            width="20"
                            height="20"
                            viewBox="0 0 24 24"
                            fill="none"
                            stroke="currentColor"
                            strokeWidth="2"
                            aria-hidden
                          >
                            <path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2" />
                            <circle cx="12" cy="7" r="4" />
                          </svg>
                          Incharge details
                        </h3>
                      </header>
                      <div className="company-details-fields-grid company-details-fields-grid--four-in-row">
                        <div className="company-details-field">
                          <label htmlFor="sm-inchargeName">
                            Name <span className="required" aria-hidden="true">*</span>
                          </label>
                          <input
                            id="sm-inchargeName"
                            name="inchargeName"
                            value={form.inchargeName}
                            onChange={handleChange}
                            onBlur={handleSiteFieldBlur}
                            placeholder="Enter name"
                            disabled={viewOnly}
                            maxLength={200}
                            required
                          />
                          {formErrors.inchargeName && (
                            <div style={{ color: 'red', fontSize: '0.95em', marginTop: 2 }}>{formErrors.inchargeName}</div>
                          )}
                        </div>
                        <div className="company-details-field">
                          <label htmlFor="sm-inchargePhone">
                            Phone <span className="required" aria-hidden="true">*</span>
                          </label>
                          <input
                            id="sm-inchargePhone"
                            name="inchargePhone"
                            type="tel"
                            inputMode="numeric"
                            autoComplete="tel"
                            value={form.inchargePhone}
                            onChange={handleChange}
                            onBlur={handleSiteFieldBlur}
                            placeholder="Enter 10-digit phone"
                            disabled={viewOnly}
                            maxLength={10}
                            title="Enter exactly 10 digits"
                            required
                          />
                          {formErrors.inchargePhone && (
                            <div style={{ color: 'red', fontSize: '0.95em', marginTop: 2 }}>{formErrors.inchargePhone}</div>
                          )}
                        </div>
                        <div className="company-details-field">
                          <label htmlFor="sm-inchargeEmail">Mail Id</label>
                          <input
                            id="sm-inchargeEmail"
                            name="inchargeEmail"
                            type="email"
                            value={form.inchargeEmail}
                            onChange={handleChange}
                            onBlur={handleSiteFieldBlur}
                            placeholder="Enter mail id"
                            disabled={viewOnly}
                            maxLength={254}
                          />
                          {formErrors.inchargeEmail && (
                            <div style={{ color: 'red', fontSize: '0.95em', marginTop: 2 }}>{formErrors.inchargeEmail}</div>
                          )}
                        </div>
                        <div className="company-details-field">
                          <label htmlFor="sm-inchargeDesignation">
                            Designation <span className="required" aria-hidden="true">*</span>
                          </label>
                          <input
                            id="sm-inchargeDesignation"
                            name="inchargeDesignation"
                            value={form.inchargeDesignation}
                            onChange={handleChange}
                            onBlur={handleSiteFieldBlur}
                            placeholder="Enter designation"
                            disabled={viewOnly}
                            maxLength={150}
                            required
                          />
                          {formErrors.inchargeDesignation && (
                            <div style={{ color: 'red', fontSize: '0.95em', marginTop: 2 }}>{formErrors.inchargeDesignation}</div>
                          )}
                        </div>
                      </div>
                    </section>

                    <section className="company-details-section-card">
                      <header className="company-details-section-head">
                        <h3 className="company-details-section-title">
                          <svg
                            className="company-details-section-title-icon"
                            width="20"
                            height="20"
                            viewBox="0 0 24 24"
                            fill="none"
                            stroke="currentColor"
                            strokeWidth="2"
                            aria-hidden
                          >
                            <path d="M3 21h18" />
                            <path d="M5 21V7l8-4 6 3v15" />
                            <path d="M9 9h.01" />
                            <path d="M9 13h.01" />
                            <path d="M9 17h.01" />
                            <path d="M14 9h.01" />
                            <path d="M14 13h.01" />
                            <path d="M14 17h.01" />
                          </svg>
                          Contractor details
                        </h3>
                      </header>
                      <div className="company-details-fields-grid">
                        <div className="company-details-field">
                          <label htmlFor="sm-contractorName">Name</label>
                          <input
                            id="sm-contractorName"
                            name="contractorName"
                            value={form.contractorName}
                            onChange={handleChange}
                            placeholder="Enter contractor name"
                            disabled={viewOnly}
                            maxLength={200}
                          />
                        </div>
                        <div className="company-details-field">
                          <label htmlFor="sm-contractorEmail">Mail Id</label>
                          <input
                            id="sm-contractorEmail"
                            name="contractorEmail"
                            type="email"
                            value={form.contractorEmail}
                            onChange={handleChange}
                            onBlur={handleSiteFieldBlur}
                            placeholder="Enter contractor email"
                            disabled={viewOnly}
                            maxLength={254}
                          />
                          {formErrors.contractorEmail && (
                            <div style={{ color: 'red', fontSize: '0.95em', marginTop: 2 }}>{formErrors.contractorEmail}</div>
                          )}
                        </div>
                        <div className="company-details-field">
                          <label htmlFor="sm-contractorPhone">Mobile Number</label>
                          <input
                            id="sm-contractorPhone"
                            name="contractorPhone"
                            type="tel"
                            inputMode="numeric"
                            autoComplete="tel"
                            value={form.contractorPhone}
                            onChange={handleChange}
                            onBlur={handleSiteFieldBlur}
                            placeholder="Enter 10-digit phone"
                            disabled={viewOnly}
                            maxLength={10}
                            title="Enter exactly 10 digits"
                          />
                          {formErrors.contractorPhone && (
                            <div style={{ color: 'red', fontSize: '0.95em', marginTop: 2 }}>{formErrors.contractorPhone}</div>
                          )}
                        </div>
                        <div className="company-details-field">
                          <label htmlFor="sm-contractorAddress">Address</label>
                          <input
                            id="sm-contractorAddress"
                            name="contractorAddress"
                            value={form.contractorAddress}
                            onChange={handleChange}
                            placeholder="Enter contractor address"
                            disabled={viewOnly}
                            maxLength={500}
                          />
                        </div>
                        <div className="company-details-field">
                          <label htmlFor="sm-contractorCity">City</label>
                          <CityCombobox
                            id="sm-contractorCity"
                            name="contractorCity"
                            value={form.contractorCity}
                            options={citySelectOptions}
                            onChange={handleChange}
                            onBlur={handleSiteFieldBlur}
                            disabled={viewOnly}
                            placeholder="Type letter to filter cities"
                            otherPlaceholder="Enter your city name"
                          />
                        </div>
                        <div className="company-details-field">
                          <label htmlFor="sm-contractorState">State</label>
                          <StateCombobox
                            id="sm-contractorState"
                            name="contractorState"
                            value={form.contractorState}
                            options={stateSelectOptions}
                            onChange={handleChange}
                            onBlur={handleSiteFieldBlur}
                            disabled={viewOnly}
                            placeholder="Type letter to filter states"
                            otherPlaceholder="Enter your state name"
                          />
                        </div>
                      </div>
                    </section>
                  </div>
                </div>
                <div className="company-details-modal-footer">
                  <button type="button" className="company-details-btn company-details-btn-secondary" onClick={closeForm}>
                    {viewOnly ? 'Close' : 'Cancel'}
                  </button>
                  {!viewOnly && (
                    <button type="submit" className="company-details-btn company-details-btn-primary company-details-btn-submit-modal">
                      {editingId ? 'Update' : 'Submit'}
                    </button>
                  )}
                </div>
              </form>
            </div>
          </div>
        </>
      ) : (
        <>
          {!showForm && message ? (
            <div className="company-details-message company-details-message--flush">{message}</div>
          ) : null}
          <header className="company-details-page-heading">
            <h1 className="company-details-page-title">Site Management</h1>
            {subtitle ? <p className="site-management-page-subline">{subtitle}</p> : null}
            <nav className="company-details-breadcrumb" aria-label="Breadcrumb">
              <Link to="/hcm-dashboard" className="company-details-bc-link">
                Dashboard
              </Link>
              <span className="company-details-bc-sep" aria-hidden>
                ›
              </span>
              <span className="company-details-bc-muted">Organization Master</span>
              <span className="company-details-bc-sep" aria-hidden>
                ›
              </span>
              <span className="company-details-bc-current">Site Management</span>
            </nav>
          </header>
          <div className="company-details-shell-box">
            <input
              ref={importFileRef}
              type="file"
              accept=".csv,.xlsx,.xls,text/csv,application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
              style={{ display: 'none' }}
              onChange={handleImportSites}
              aria-hidden
            />
            {loading ? (
              <p className="site-management-loading">Loading...</p>
            ) : displaySites.length === 0 ? (
              <div className="company-details-empty-state">
                <p className="site-management-empty">{emptyListMessage}</p>
                <div className="company-details-empty-actions">
                  <button
                    type="button"
                    className="company-details-btn company-details-btn-export"
                    onClick={() => importFileRef.current?.click()}
                    disabled={importing}
                    title="Import CSV"
                    aria-label="Import CSV"
                  >
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden>
                      <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
                      <polyline points="7 10 12 15 17 10" />
                      <line x1="12" y1="15" x2="12" y2="3" />
                    </svg>
                    {importing ? 'Importing…' : 'Import CSV'}
                  </button>
                  <button type="button" className="company-details-btn company-details-btn-add" onClick={openAdd}>
                    <span className="company-details-btn-add-icon">+</span> Add Site
                  </button>
                </div>
              </div>
            ) : (
              <>
                <div className="company-details-inner-toolbar">
                  <div className="company-details-table-search company-details-table-search--flex">
                    <svg className="company-details-table-search-icon" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" aria-hidden>
                      <circle cx="11" cy="11" r="8" />
                      <path d="m21 21-4.35-4.35" />
                    </svg>
                    <input
                      type="search"
                      className="company-details-table-search-input"
                      placeholder="Search by Site Name, City, Incharge, Email…"
                      value={tableSearch}
                      onChange={(e) => setTableSearch(e.target.value)}
                      aria-label="Search sites"
                    />
                  </div>
                  <div className="company-details-inner-toolbar-right">
                    <button
                      type="button"
                      className="company-details-btn company-details-btn-export"
                      onClick={() => importFileRef.current?.click()}
                      disabled={importing}
                      title="Import CSV"
                      aria-label="Import CSV"
                    >
                      <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden>
                        <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
                        <polyline points="7 10 12 15 17 10" />
                        <line x1="12" y1="15" x2="12" y2="3" />
                      </svg>
                      {importing ? 'Importing…' : 'Import CSV'}
                    </button>
                    <button type="button" className="company-details-btn company-details-btn-export" onClick={exportSitesCsv} title="Export CSV" aria-label="Export CSV">
                      <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden>
                        <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
                        <polyline points="17 8 12 3 7 8" />
                        <line x1="12" y1="3" x2="12" y2="15" />
                      </svg>
                      Export CSV
                    </button>
                    <button type="button" className="company-details-btn company-details-btn-add" onClick={openAdd}>
                      <span className="company-details-btn-add-icon">+</span> Add Site
                    </button>
                    {allPageSelected && selectedIds.size > 0 ? (
                      <button
                        type="button"
                        className="company-details-btn company-details-btn-icon site-management-bulk-delete-btn"
                        onClick={handleBulkDeleteSelected}
                        disabled={bulkDeleting}
                        title={bulkDeleting ? 'Deleting selected sites...' : 'Delete selected sites'}
                        aria-label="Delete selected sites"
                      >
                        <Trash2 size={18} strokeWidth={2} aria-hidden />
                      </button>
                    ) : null}
                  </div>
                </div>
                <div className="company-details-table-wrap">
                  <table className="company-details-data-table site-management-data-table">
                    <thead>
                      <tr>
                        <th className="company-details-th-check" scope="col">
                          <input
                            type="checkbox"
                            className="company-details-table-checkbox"
                            checked={allPageSelected}
                            ref={(el) => {
                              if (el) el.indeterminate = somePageSelected && !allPageSelected;
                            }}
                            onChange={toggleSelectAllDisplay}
                            aria-label="Select all rows on this page"
                          />
                        </th>
                        <th scope="col">#</th>
                        <th scope="col">Company</th>
                        <th scope="col">Site name</th>
                        <th scope="col">City</th>
                        <th scope="col">State</th>
                        <th scope="col">Address</th>
                        <th scope="col">Unit</th>
                        <th scope="col">Incharge</th>
                        <th scope="col">Phone</th>
                        <th scope="col">Mail Id</th>
                        <th scope="col">Industry</th>
                        <th scope="col">Location</th>
                        <th className="company-details-th-actions" scope="col">
                          Actions
                        </th>
                      </tr>
                    </thead>
                    <tbody>
                      {sortedDisplaySites.length === 0 ? (
                        <tr>
                          <td colSpan={TABLE_COL_COUNT} className="company-details-table-empty-cell">
                            No sites match your search.
                          </td>
                        </tr>
                      ) : (
                        pagedSites.map((site, idx) => (
                          <tr key={site.id} className="company-details-data-row">
                            <td>
                              <input
                                type="checkbox"
                                className="company-details-table-checkbox"
                                checked={selectedIds.has(String(site.id))}
                                onChange={() => toggleRowSelected(site.id)}
                                aria-label={`Select ${site.siteName || 'site'}`}
                              />
                            </td>
                            <td className="company-details-td-num">{(effectiveTablePage - 1) * SITE_TABLE_PAGE_SIZE + idx + 1}</td>
                            <td className="site-management-td-clip" title={siteCompanyName(site)}>
                              {siteCompanyName(site) || '—'}
                            </td>
                            <td className="company-details-td-strong">{site.siteName || '—'}</td>
                            <td>{site.siteCity || '—'}</td>
                            <td>{site.siteState || '—'}</td>
                            <td className="site-management-td-clip" title={site.siteAddress || ''}>
                              {site.siteAddress || '—'}
                            </td>
                            <td>{site.unitNo || '—'}</td>
                            <td>{site.inchargeName || '—'}</td>
                            <td>{site.inchargePhone || '—'}</td>
                            <td className="site-management-td-email" title={siteInchargeEmail(site)}>
                              {siteInchargeEmail(site) || '—'}
                            </td>
                            <td className="site-management-td-industry" title={siteIndustry(site)}>
                              {siteIndustry(site) || '—'}
                            </td>
                            <td className="site-management-td-clip" title={siteLocation(site)}>
                              {siteLocation(site) || '—'}
                            </td>
                            <td>
                              <div className="company-details-table-actions">
                                <button
                                  type="button"
                                  className="company-details-table-action company-details-table-action--view"
                                  onClick={() => openView(site)}
                                  title="View"
                                  aria-label={`View ${site.siteName || 'site'}`}
                                >
                                  <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                                    <path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z" />
                                    <circle cx="12" cy="12" r="3" />
                                  </svg>
                                </button>
                                <button
                                  type="button"
                                  className="company-details-table-action company-details-table-action--edit"
                                  onClick={() => openEdit(site)}
                                  title="Edit"
                                  aria-label={`Edit ${site.siteName || 'site'}`}
                                >
                                  <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                                    <path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7" />
                                    <path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z" />
                                  </svg>
                                </button>
                                <button
                                  type="button"
                                  className="company-details-table-action company-details-table-action--delete"
                                  onClick={() => handleDelete(site.id)}
                                  title="Delete"
                                  aria-label={`Delete ${site.siteName || 'site'}`}
                                >
                                  <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                                    <polyline points="3 6 5 6 21 6" />
                                    <path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3-3V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2" />
                                    <line x1="10" y1="11" x2="10" y2="17" />
                                    <line x1="14" y1="11" x2="14" y2="17" />
                                  </svg>
                                </button>
                              </div>
                            </td>
                          </tr>
                        ))
                      )}
                    </tbody>
                  </table>
                </div>
                {sortedDisplaySites.length > 0 ? (
                  <div className="company-details-pagination-footer">
                    <span className="company-details-pagination-info">
                      Showing {showingFrom} to {showingTo} of {totalVisible} results
                    </span>
                    <nav className="company-details-pagination" aria-label="Table pagination">
                      <button
                        type="button"
                        className="company-details-pagination-nav"
                        disabled={effectiveTablePage <= 1}
                        onClick={() => setTablePage(1)}
                        title="First page"
                        aria-label="First page"
                      >
                        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden>
                          <polyline points="11 17 6 12 11 7" />
                          <polyline points="18 17 13 12 18 7" />
                        </svg>
                      </button>
                      <button
                        type="button"
                        className="company-details-pagination-nav"
                        disabled={effectiveTablePage <= 1}
                        onClick={() => setTablePage((p) => Math.max(1, p - 1))}
                        title="Previous page"
                        aria-label="Previous page"
                      >
                        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden>
                          <polyline points="15 18 9 12 15 6" />
                        </svg>
                      </button>
                      <div className="company-details-pagination-pages">
                        {paginationItems.map((item, i) =>
                          item === 'ellipsis' ? (
                            <span key={`e-${i}`} className="company-details-pagination-ellipsis" aria-hidden>
                              ...
                            </span>
                          ) : (
                            <button
                              key={item}
                              type="button"
                              className={`company-details-pagination-page${item === effectiveTablePage ? ' company-details-pagination-page--active' : ''}`}
                              onClick={() => setTablePage(item)}
                              aria-label={`Page ${item}`}
                              aria-current={item === effectiveTablePage ? 'page' : undefined}
                            >
                              {item}
                            </button>
                          )
                        )}
                      </div>
                      <button
                        type="button"
                        className="company-details-pagination-nav"
                        disabled={effectiveTablePage >= totalTablePages}
                        onClick={() => setTablePage((p) => Math.min(totalTablePages, p + 1))}
                        title="Next page"
                        aria-label="Next page"
                      >
                        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden>
                          <polyline points="9 18 15 12 9 6" />
                        </svg>
                      </button>
                      <button
                        type="button"
                        className="company-details-pagination-nav"
                        disabled={effectiveTablePage >= totalTablePages}
                        onClick={() => setTablePage(totalTablePages)}
                        title="Last page"
                        aria-label="Last page"
                      >
                        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden>
                          <polyline points="13 17 18 12 13 7" />
                          <polyline points="6 17 11 12 6 7" />
                        </svg>
                      </button>
                    </nav>
                  </div>
                ) : null}
              </>
            )}
          </div>
        </>
      )}
    </div>
  );
};

export default SiteManagement;
