import React, { useState, useEffect, useCallback, useMemo, useRef } from 'react';
import { Link } from 'react-router-dom';
import { Calendar, CheckCircle2, Trash2, Upload } from 'lucide-react';
import * as XLSX from 'xlsx';
import './CompanyDetails.css';
import { INDIAN_CITIES } from '../utils/indianCities';
import { INDIAN_STATES } from '../utils/indianStates';
import CityCombobox, { StateCombobox } from '../components/CityCombobox';

const API_BASE = '/server/company_function';

/** Full company form fields for CSV export/import (ID omitted from export; import still accepts ID if present). */
const COMPANY_CSV_COLS = [
  { key: 'companyName', label: 'Name' },
  { key: 'companyMail', label: 'Mail Id' },
  { key: 'companyPhoneNumber', label: 'Moblie Number' },
  { key: 'companyAddress', label: 'Address' },
  { key: 'city', label: 'City' },
  { key: 'state', label: 'State' },
  { key: 'postalcode', label: 'Postal code' },
  { key: 'incorprationDate', label: 'Incorporation date' },
  { key: 'incorporationNo', label: 'Incorporation number' },
  { key: 'companyPANNumber', label: 'PAN' },
  { key: 'gstNo', label: 'GST number' },
  { key: 'pfNo', label: 'PF number' },
  { key: 'esiNo', label: 'ESI number' },
  { key: 'directorName', label: 'Director Name' },
  { key: 'directorMail', label: 'Director Mail Id' },
  { key: 'directorPhoneNumber', label: 'Director Mobile Number' },
  { key: 'directorAddress', label: 'Director Address' },
  { key: 'ownerName', label: 'Owner / In-charge name' },
  { key: 'ownerPAN', label: 'Owner / In-charge PAN' },
  { key: 'ownerAaadhar', label: 'Owner / In-charge Aadhar' },
  { key: 'ownerDesignation', label: 'Owner / In-charge designation' },
  { key: 'safetyOfficerName', label: 'Safety officer name' },
  { key: 'safetyOfficerPhone', label: 'Safety officer phone' }
];

/** Older exports included ID and used duplicate labels for director fields — map by position. */
const COMPANY_CSV_LEGACY_LABELS = [
  'ID',
  'Name',
  'Mail Id',
  'Moblie Number',
  'Address',
  'City',
  'State',
  'Postal code',
  'Incorporation date',
  'Incorporation number',
  'PAN',
  'GST number',
  'PF number',
  'ESI number',
  'Name',
  'Mail Id',
  'Mobile Number',
  'Address',
  'Owner / In-charge name',
  'Owner / In-charge PAN',
  'Owner / In-charge Aadhar',
  'Owner / In-charge designation',
  'Safety officer name',
  'Safety officer phone'
];

/** Older export with ID + unique director labels (before ID was removed from export). */
const COMPANY_CSV_LEGACY_WITH_ID_LABELS = [
  'ID',
  ...COMPANY_CSV_COLS.map((c) => c.label)
];

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

function buildCompanyCsvHeaderIndexMap(headerRow) {
  const map = {};

  // Older export: ID + duplicate director labels (Name / Mail Id / …)
  if (headersMatchLabelList(headerRow, COMPANY_CSV_LEGACY_LABELS)) {
    const legacyKeys = ['id', ...COMPANY_CSV_COLS.map((c) => c.key)];
    legacyKeys.forEach((key, i) => {
      map[key] = i;
    });
    return map;
  }

  // Older export: ID + unique director labels
  if (headersMatchLabelList(headerRow, COMPANY_CSV_LEGACY_WITH_ID_LABELS)) {
    map.id = 0;
    COMPANY_CSV_COLS.forEach((c, i) => {
      map[c.key] = i + 1;
    });
    return map;
  }

  // Current export (no ID column)
  if (
    COMPANY_CSV_COLS.every(
      (c, i) => normalizeCsvHeader(headerRow[i]) === normalizeCsvHeader(c.label)
    )
  ) {
    COMPANY_CSV_COLS.forEach((c, i) => {
      map[c.key] = i;
    });
    return map;
  }

  const aliases = {
    id: ['id', 'rowid'],
    companyName: ['name', 'company name', 'companyname'],
    companyMail: ['mail id', 'email', 'company mail', 'company email', 'companymail'],
    companyPhoneNumber: ['moblie number', 'mobile number', 'phone', 'company phone', 'companyphone'],
    companyAddress: ['address', 'company address'],
    city: ['city'],
    state: ['state'],
    postalcode: ['postal code', 'postalcode', 'pincode', 'pin code'],
    incorprationDate: ['incorporation date', 'incorpration date', 'incorporationdate'],
    incorporationNo: ['incorporation number', 'incorporation no', 'incorporationno'],
    companyPANNumber: ['pan', 'company pan', 'companypannumber'],
    gstNo: ['gst number', 'gst', 'gstin', 'gstno'],
    pfNo: ['pf number', 'pf', 'pfno'],
    esiNo: ['esi number', 'esi', 'esino'],
    directorName: ['director name', 'director'],
    directorMail: ['director mail id', 'director email', 'director mail'],
    directorPhoneNumber: ['director mobile number', 'director phone', 'director phone number'],
    directorAddress: ['director address'],
    ownerName: ['owner / in-charge name', 'owner name', 'owner/in-charge name'],
    ownerPAN: ['owner / in-charge pan', 'owner pan'],
    ownerAaadhar: ['owner / in-charge aadhar', 'owner aadhar', 'owner aadhaar'],
    ownerDesignation: ['owner / in-charge designation', 'owner designation'],
    safetyOfficerName: ['safety officer name'],
    safetyOfficerPhone: ['safety officer phone']
  };
  const normalizedHeaders = headerRow.map(normalizeCsvHeader);
  Object.keys(aliases).forEach((key) => {
    const idx = normalizedHeaders.findIndex((h) => aliases[key].includes(h));
    if (idx >= 0) map[key] = idx;
  });
  return map;
}

/** Format a JS Date as DD/MM/YYYY using local calendar parts. */
function formatJsDateToDdMmYyyy(d) {
  if (!(d instanceof Date) || isNaN(d.getTime())) return '';
  const day = String(d.getDate()).padStart(2, '0');
  const mo = String(d.getMonth() + 1).padStart(2, '0');
  return `${day}/${mo}/${d.getFullYear()}`;
}

/** Convert Excel serial (or SheetJS date code) to DD/MM/YYYY via SSF (avoids timezone off-by-one). */
function excelSerialToDdMmYyyy(serial) {
  const num = Number(serial);
  if (!Number.isFinite(num) || num <= 0) return '';
  if (XLSX.SSF && typeof XLSX.SSF.parse_date_code === 'function') {
    const parsed = XLSX.SSF.parse_date_code(num);
    if (parsed && parsed.y && parsed.m && parsed.d) {
      const day = String(parsed.d).padStart(2, '0');
      const mo = String(parsed.m).padStart(2, '0');
      return `${day}/${mo}/${parsed.y}`;
    }
  }
  const utc = new Date(Date.UTC(1899, 11, 30) + Math.round(num) * 86400000);
  if (isNaN(utc.getTime())) return '';
  const day = String(utc.getUTCDate()).padStart(2, '0');
  const mo = String(utc.getUTCMonth() + 1).padStart(2, '0');
  return `${day}/${mo}/${utc.getUTCFullYear()}`;
}

/**
 * Read sheet as AOA, preferring Excel display text (`cell.w`) so DD/MM/YYYY dates
 * are not reformatted to US short dates like 6/11/26.
 */
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
      if (cell.t === 'd' && cell.v instanceof Date) {
        row.push(cell.v);
        continue;
      }
      if (cell.t === 'n' && cell.v != null) {
        row.push(cell.v);
        continue;
      }
      row.push(cell.v == null ? '' : cell.v);
    }
    rows.push(row);
  }
  return rows;
}

/**
 * Normalize Excel/CSV date cells to DD/MM/YYYY for validation.
 * Excel stores dates as serials/Date objects; display may look like DD/MM/YYYY
 * while the raw value is not a plain date string.
 */
function normalizeImportIncorporationDate(val) {
  if (val == null || val === '') return '';
  if (val instanceof Date) {
    // Prefer SSF path when we only have a Date — use UTC+local safe formatting via serial
    const serial = val.getTime() / 86400000 + 25569;
    const fromSerial = excelSerialToDdMmYyyy(serial);
    if (fromSerial) return fromSerial;
    return formatJsDateToDdMmYyyy(val);
  }
  if (typeof val === 'number' && Number.isFinite(val)) {
    if (val >= 1000 && val <= 9999 && Number.isInteger(val)) return `01/01/${val}`;
    if (val > 20000 && val < 1000000) return excelSerialToDdMmYyyy(val);
  }
  let s = stripCsvExcelTextPrefix(val);
  if (!s) return '';
  if (/^\d+(\.\d+)?$/.test(s)) {
    const num = Number(s);
    if (num >= 1000 && num <= 9999 && Number.isInteger(num)) return `01/01/${num}`;
    if (num > 20000 && num < 1000000) return excelSerialToDdMmYyyy(num);
  }
  const isoDt = /^(\d{4}-\d{2}-\d{2})(?:[T\s].*)?$/.exec(s);
  if (isoDt) {
    const [y, mo, d] = isoDt[1].split('-');
    return `${d}/${mo}/${y}`;
  }
  // Strip trailing time: "11/06/2026 00:00:00", "11/06/2026 12:00 AM"
  s = s.replace(/\s+\d{1,2}:\d{2}(:\d{2})?(\.\d+)?(\s*[AaPp][Mm])?\s*$/, '').trim();
  return s;
}

function parseCompanyImportWorkbook(arrayBuffer) {
  // cellDates:false keeps Excel dates as serials; we use cell.w (display) when present
  const wb = XLSX.read(arrayBuffer, { type: 'array', cellDates: false });
  const sheetName = wb.SheetNames[0];
  if (!sheetName) throw new Error('No sheets found in the file.');
  const ws = wb.Sheets[sheetName];
  const rows = sheetToAoaPreferDisplayText(ws);
  if (!rows.length) throw new Error('No data found in the file.');
  const headerRow = rows[0].map((h) => stripCsvExcelTextPrefix(h));
  const indexMap = buildCompanyCsvHeaderIndexMap(headerRow);
  if (indexMap.companyName == null) {
    throw new Error('Could not find a Name / Company Name column. Use Export CSV as a template.');
  }
  const dataRows = rows.slice(1).filter((row) =>
    Array.isArray(row) && row.some((cell) => String(cell ?? '').trim() !== '')
  );
  return dataRows.map((row, rowIndex) => {
    const record = { __row: rowIndex + 2, id: '' };
    if (indexMap.id != null) {
      record.id = stripCsvExcelTextPrefix(row[indexMap.id]);
    }
    COMPANY_CSV_COLS.forEach(({ key }) => {
      const idx = indexMap[key];
      if (idx == null) {
        record[key] = '';
        return;
      }
      const rawVal = row[idx];
      if (key === 'incorprationDate') {
        record[key] = normalizeImportIncorporationDate(rawVal);
        return;
      }
      if (rawVal instanceof Date && !isNaN(rawVal.getTime())) {
        record[key] = formatJsDateToDdMmYyyy(rawVal);
        return;
      }
      record[key] = stripCsvExcelTextPrefix(rawVal);
    });
    return record;
  });
}

function buildCompanyPayloadFromImportRow(row) {
  const incParsed = parseFlexibleIncorporationDate(String(row.incorprationDate || '').trim());
  return {
    companyName: String(row.companyName || '').trim(),
    companyPANNumber: String(row.companyPANNumber || '').trim().toUpperCase(),
    companyAddress: String(row.companyAddress || '').trim(),
    city: String(row.city || '').trim(),
    state: String(row.state || '').trim(),
    postalcode: String(row.postalcode || '').trim(),
    incorprationDate: incParsed || null,
    incorporationNo: String(row.incorporationNo || '').trim(),
    gstNo: String(row.gstNo || '').trim().toUpperCase(),
    pfNo: normalizePfNumberInput(row.pfNo),
    esiNo: normalizeEsiNumberInput(row.esiNo),
    companyMail: String(row.companyMail || '').trim(),
    companyPhoneNumber: digitsOnly(row.companyPhoneNumber),
    directorName: String(row.directorName || '').trim(),
    directorPhoneNumber: digitsOnly(row.directorPhoneNumber),
    directorMail: String(row.directorMail || '').trim(),
    directorAddress: String(row.directorAddress || '').trim(),
    ownerName: String(row.ownerName || '').trim(),
    ownerPAN: String(row.ownerPAN || '').trim().toUpperCase(),
    ownerAaadhar: digitsOnly(row.ownerAaadhar),
    ownerDesignation: String(row.ownerDesignation || '').trim(),
    safetyOfficerName: String(row.safetyOfficerName || '').trim(),
    safetyOfficerPhone: digitsOnly(row.safetyOfficerPhone),
    doctroName: '',
    doctroPhone: ''
  };
}

/** Keep in sync with HcmDashboardSidebar (same-route Company Details nav). */
const COMPANY_DETAILS_CLOSE_MODAL_EVENT = 'company-details-close-modal';

const COMPANY_TABLE_PAGE_SIZE = 10;

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

const initialForm = {
  companyName: '',
  companyAddress: '',
  city: '',
  state: '',
  postalcode: '',
  incorporationNo: '',
  pfNo: '',
  companyMail: '',
  directorName: '',
  directorMail: '',
  directorAddress: '',
  companyPANNumber: '',
  incorprationDate: '',
  gstNo: '',
  esiNo: '',
  companyPhoneNumber: '',
  directorPhoneNumber: '',
  ownerName: '',
  ownerPAN: '',
  ownerAaadhar: '',
  ownerDesignation: '',
  safetyOfficerName: '',
  safetyOfficerPhone: ''
};

const DOC_TYPES = [
  { docType: 'DoctroAppoitnmentorder', label: 'Director Appointment Order' },
  { docType: 'DoctroApporvalcopy', label: 'Director Approval Copy' },
  { docType: 'HeadHRSign', label: 'Head HR Sign' },
  { docType: 'HeadHRSeal', label: 'Head HR Seal' }
];

/** Same mapping as company_function DOC_TYPE_TO_COMPANY_KEY — list rows use camelCase file id fields. */
const DOC_TYPE_TO_COMPANY_KEY_UI = {
  DoctroAppoitnmentorder: 'doctroAppoitnmentorder',
  DoctroApporvalcopy: 'doctroApporvalcopy',
  HeadHRSign: 'headHRSign',
  HeadHRSeal: 'headHRSeal'
};

const PAN_REGEX = /^[A-Z]{5}[0-9]{4}[A-Z]$/i;
const EMAIL_REGEX = /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/i;
/** PF establishment id e.g. MH/BAN/1234567/000/1234567 */
const PF_NUMBER_REGEX = /^[A-Z]{2}\/[A-Z]{3}\/\d{7}\/\d{3}\/\d{7}$/i;
/** ESI code e.g. 31-00-123456-000-0001 (hyphen or en/em dash) */
const ESI_NUMBER_REGEX = /^\d{2}-\d{2}-\d{6}-\d{3}-\d{4}$/;
/** GSTIN 15 chars e.g. 29AAACH7409R1ZX */
const GSTIN_REGEX = /^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1}$/;

const COMPANY_FORM_FIELD_ORDER = [
  'companyName',
  'companyMail',
  'companyPhoneNumber',
  'companyAddress',
  'city',
  'state',
  'postalcode',
  'incorprationDate',
  'incorporationNo',
  'companyPANNumber',
  'gstNo',
  'pfNo',
  'esiNo',
  'directorName',
  'directorMail',
  'directorPhoneNumber',
  'directorAddress',
  'ownerName',
  'ownerPAN',
  'ownerAaadhar',
  'ownerDesignation',
  'safetyOfficerName',
  'safetyOfficerPhone'
];

const SECTION_FIELD_GROUPS = [
  { keys: ['companyName', 'companyMail', 'companyPhoneNumber', 'companyAddress', 'city', 'state', 'postalcode'] },
  { keys: ['incorprationDate', 'incorporationNo', 'companyPANNumber', 'gstNo', 'pfNo', 'esiNo'] },
  { keys: ['directorName', 'directorMail', 'directorPhoneNumber', 'directorAddress'] },
  { keys: ['ownerName', 'ownerPAN', 'ownerAaadhar', 'ownerDesignation', 'safetyOfficerName', 'safetyOfficerPhone'] }
];

function companyInputId(name) {
  return name === 'incorprationDate' ? 'cd-incorprationDate' : `cd-${name}`;
}

function digitsOnly(s) {
  return String(s || '').replace(/\D/g, '');
}

function normalizePfNumberInput(s) {
  return String(s || '').trim();
}

function normalizeEsiNumberInput(s) {
  return String(s || '')
    .trim()
    .replace(/[\u2010\u2011\u2012\u2013\u2014\u2015\u2212]/g, '-')
    .replace(/\s+/g, '');
}

function isValidCalendarDate(y, m, d) {
  if (m < 1 || m > 12 || d < 1 || d > 31) return false;
  const dt = new Date(y, m - 1, d);
  return dt.getFullYear() === y && dt.getMonth() === m - 1 && dt.getDate() === d;
}

function padYmd(y, m, d) {
  return `${y}-${String(m).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
}

/** Normalize unicode dashes to ASCII hyphen for parsing. */
function normalizeDateSeparators(t) {
  return String(t || '').replace(/[\u2010\u2011\u2012\u2013\u2014\u2015\u2212]/g, '-');
}

/**
 * Parse incorporation date from user input to canonical YYYY-MM-DD (local calendar).
 * Accepts YYYY-MM-DD, DD/MM/YYYY, DD-MM-YYYY, DD.MM.YYYY, optional time suffix,
 * 2-digit years, and 4-digit year → YYYY-01-01.
 */
function parseFlexibleIncorporationDate(s) {
  let t = normalizeDateSeparators(String(s || '').trim());
  if (!t) return null;
  // Drop time portion Excel often appends (e.g. "11/06/2026 00:00:00")
  t = t.replace(/\s+\d{1,2}:\d{2}(:\d{2})?(\.\d+)?(\s*[AaPp][Mm])?\s*$/, '').trim();
  const isoDt = /^(\d{4}-\d{2}-\d{2})(?:[T\s].*)?$/.exec(t);
  if (isoDt) t = isoDt[1];
  let m = /^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$/.exec(t);
  if (m) {
    const y = +m[1];
    const mo = +m[2];
    const d = +m[3];
    if (isValidCalendarDate(y, mo, d)) return padYmd(y, mo, d);
    return null;
  }
  m = /^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})$/.exec(t);
  if (m) {
    const d = +m[1];
    const mo = +m[2];
    const y = +m[3];
    if (isValidCalendarDate(y, mo, d)) return padYmd(y, mo, d);
    return null;
  }
  // DD/MM/YY (Excel short year)
  m = /^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2})$/.exec(t);
  if (m) {
    const d = +m[1];
    const mo = +m[2];
    const yy = +m[3];
    const y = yy >= 70 ? 1900 + yy : 2000 + yy;
    if (isValidCalendarDate(y, mo, d)) return padYmd(y, mo, d);
    return null;
  }
  m = /^(\d{4})$/.exec(t);
  if (m) {
    const y = +m[1];
    if (y >= 1000 && y <= 9999) return `${y}-01-01`;
  }
  return null;
}

function sanitizeIncorporationDateInput(raw) {
  const digits = String(raw || '').replace(/\D/g, '').slice(0, 8);
  if (digits.length === 0) return '';
  if (digits.length <= 2) return `${digits}${digits.length === 2 ? '/' : ''}`;
  if (digits.length <= 4) {
    const month = digits.slice(2);
    return `${digits.slice(0, 2)}/${month}${digits.length === 4 ? '/' : ''}`;
  }
  return `${digits.slice(0, 2)}/${digits.slice(2, 4)}/${digits.slice(4)}`;
}

function isoStringToDdMmYyyy(iso) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(iso || '').trim());
  if (!m) return '';
  return `${m[3]}/${m[2]}/${m[1]}`;
}

/** Form state stores YYYY-MM-DD when known; show DD/MM/YYYY in the text field. */
function incorporationInputDisplay(stored) {
  const s = String(stored || '').trim();
  if (!s) return '';
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return isoStringToDdMmYyyy(s);
  return s;
}

function normalizeIncorporationDateOnBlurString(raw) {
  const t = String(raw || '').trim();
  if (!t) return '';
  const p = parseFlexibleIncorporationDate(t);
  return p || t;
}

function normalizeStoredIncorporationDateForForm(raw) {
  const s = String(raw || '').trim();
  if (!s) return '';
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  const p = parseFlexibleIncorporationDate(s);
  if (p) return p;
  const d = new Date(s);
  if (!isNaN(d.getTime())) {
    return padYmd(d.getFullYear(), d.getMonth() + 1, d.getDate());
  }
  return s;
}

/** Value for a native `<input type="date">` (YYYY-MM-DD or empty). */
function isoForNativeDatePickerValue(stored) {
  const s = String(stored || '').trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  const p = parseFlexibleIncorporationDate(s);
  return p || '';
}

/** Constrain input while typing (prevents e.g. 11th phone digit, over-long PAN/GST). */
function sanitizeCompanyFormField(name, raw) {
  const v = raw == null ? '' : String(raw);
  switch (name) {
    case 'companyPhoneNumber':
    case 'directorPhoneNumber':
    case 'safetyOfficerPhone':
      return digitsOnly(v).slice(0, 10);
    case 'ownerAaadhar':
      return digitsOnly(v).slice(0, 12);
    case 'companyPANNumber':
    case 'ownerPAN':
      return v
        .toUpperCase()
        .replace(/[^A-Z0-9]/g, '')
        .slice(0, 10);
    case 'gstNo':
      return v.toUpperCase().replace(/[^A-Z0-9]/g, '').slice(0, 15);
    case 'pfNo':
      return v.toUpperCase().replace(/[^A-Z0-9/]/g, '').slice(0, 26);
    case 'esiNo':
      return v
        .replace(/[\u2010\u2011\u2012\u2013\u2014\u2015\u2212]/g, '-')
        .replace(/[^0-9-]/g, '')
        .slice(0, 21);
    case 'companyMail':
    case 'directorMail':
      return v.replace(/\s/g, '').slice(0, 254);
    case 'incorprationDate':
      return sanitizeIncorporationDateInput(v);
    case 'city':
    case 'state':
      return v.slice(0, 120);
    default:
      return v;
  }
}

/** Client-side validation (Add / Edit company modal) */
function validateCompanyFormValues(values) {
  const e = {};

  if (!String(values.companyName || '').trim()) {
    e.companyName = 'Name is required';
  }

  const mail = String(values.companyMail || '').trim();
  if (!mail) {
    e.companyMail = 'Mail Id is required';
  } else if (!EMAIL_REGEX.test(mail)) {
    e.companyMail = 'Enter a valid email address (e.g., info@company.com)';
  }

  if (!String(values.companyPhoneNumber || '').trim()) {
    e.companyPhoneNumber = 'Moblie Number is required';
  } else if (digitsOnly(values.companyPhoneNumber).length !== 10) {
    e.companyPhoneNumber = 'Enter a valid 10-digit phone number';
  }

  if (!String(values.companyAddress || '').trim()) {
    e.companyAddress = 'Address is required';
  }

  const incRaw = String(values.incorprationDate || '').trim();
  if (!incRaw) {
    e.incorprationDate = 'Incorporation date is required';
  } else if (!parseFlexibleIncorporationDate(incRaw)) {
    e.incorprationDate = 'Enter a valid date (DD/MM/YYYY)';
  }

  if (!String(values.incorporationNo || '').trim()) {
    e.incorporationNo = 'Incorporation number is required';
  }

  const pan = String(values.companyPANNumber || '').trim().toUpperCase();
  if (!pan) {
    e.companyPANNumber = 'PAN is required';
  } else if (!PAN_REGEX.test(pan)) {
    e.companyPANNumber = 'Enter a valid PAN (e.g., ABCDE1234F)';
  }

  const gst = String(values.gstNo || '').trim().toUpperCase();
  if (gst.length > 0 && !GSTIN_REGEX.test(gst)) {
    e.gstNo = 'Enter a valid GSTIN (e.g., 29AAACH7409R1ZX)';
  }

  const pfRaw = normalizePfNumberInput(values.pfNo);
  if (pfRaw.length > 0 && !PF_NUMBER_REGEX.test(pfRaw)) {
    e.pfNo = 'Enter a valid PF number (e.g., MH/BAN/1234567/000/1234567)';
  }

  const esiNorm = normalizeEsiNumberInput(values.esiNo);
  if (esiNorm.length > 0 && !ESI_NUMBER_REGEX.test(esiNorm)) {
    e.esiNo = 'Enter a valid ESI number (e.g., 31-00-123456-000-0001)';
  }

  const dMail = String(values.directorMail || '').trim();
  if (dMail && !EMAIL_REGEX.test(dMail)) {
    e.directorMail = 'Enter a valid email address (e.g., info@company.com)';
  }
  if (String(values.directorPhoneNumber || '').trim() && digitsOnly(values.directorPhoneNumber).length !== 10) {
    e.directorPhoneNumber = 'Enter a valid 10-digit phone number';
  }

  if (!String(values.ownerName || '').trim()) {
    e.ownerName = 'Owner / In-charge name is required';
  }
  const ownerPan = String(values.ownerPAN || '').trim().toUpperCase();
  if (!ownerPan) {
    e.ownerPAN = 'Owner / In-charge PAN is required';
  } else if (!PAN_REGEX.test(ownerPan)) {
    e.ownerPAN = 'Enter a valid PAN (e.g., ABCDE1234F)';
  }
  if (!String(values.ownerAaadhar || '').trim()) {
    e.ownerAaadhar = 'Owner / In-charge Aadhar is required';
  } else if (digitsOnly(values.ownerAaadhar).length !== 12) {
    e.ownerAaadhar = 'Enter a valid 12-digit Aadhar number';
  }
  if (!String(values.ownerDesignation || '').trim()) {
    e.ownerDesignation = 'Owner / In-charge designation is required';
  }

  if (String(values.safetyOfficerPhone || '').trim() && digitsOnly(values.safetyOfficerPhone).length !== 10) {
    e.safetyOfficerPhone = 'Enter a valid 10-digit phone number';
  }

  return e;
}

function countSectionErrors(errs, keys) {
  return keys.filter((k) => errs[k]).length;
}

function companyRowId(c) {
  if (!c || typeof c !== 'object') return '';
  return String(c.id ?? c.ROWID ?? c.rowid ?? c.Company?.ROWID ?? '').trim();
}

/** Natural key for import upsert: Company Name (case-insensitive). */
function companyImportNameKey(name) {
  return String(name ?? '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ');
}

function companyRecordNameKey(c) {
  if (!c || typeof c !== 'object') return '';
  return companyImportNameKey(c.companyName ?? c.CompanyName ?? c.Name ?? '');
}

/** Normalize location fields from API list row or Datastore payload (PascalCase or camelCase). */
function normalizeCompanyRecord(company) {
  if (!company || typeof company !== 'object') return company;
  return {
    ...company,
    city: String(company.city ?? company.City ?? '').trim(),
    state: String(company.state ?? company.State ?? '').trim(),
    postalcode: String(company.postalcode ?? company.Postalcode ?? company.PostalCode ?? '').trim()
  };
}

function buildFormFromCompany(company) {
  const c = normalizeCompanyRecord(company);
  return {
    companyName: c.companyName || '',
    companyAddress: c.companyAddress || '',
    city: c.city || '',
    state: c.state || '',
    postalcode: c.postalcode || '',
    incorporationNo: c.incorporationNo || '',
    pfNo: c.pfNo || '',
    companyMail: c.companyMail || '',
    directorName: c.directorName || '',
    directorMail: c.directorMail || '',
    directorAddress: c.directorAddress || '',
    companyPANNumber: c.companyPANNumber || '',
    incorprationDate: normalizeStoredIncorporationDateForForm(c.incorprationDate || c.incorporationDate || ''),
    gstNo: c.gstNo || '',
    esiNo: c.esiNo || '',
    companyPhoneNumber: c.companyPhoneNumber || '',
    directorPhoneNumber: c.directorPhoneNumber || '',
    ownerName: c.ownerName || '',
    ownerPAN: c.ownerPAN || '',
    ownerAaadhar: c.ownerAaadhar || '',
    ownerDesignation: c.ownerDesignation || '',
    safetyOfficerName: c.safetyOfficerName || '',
    safetyOfficerPhone: c.safetyOfficerPhone || ''
  };
}

async function enrichCompanyForForm(company) {
  const normalized = normalizeCompanyRecord(company);
  const id = companyRowId(normalized);
  if (!id || normalized.city || normalized.state || normalized.postalcode) {
    return normalized;
  }
  try {
    const res = await fetch(`${API_BASE}/company/${encodeURIComponent(id)}`);
    const data = await res.json();
    if (data.status === 'success' && data.data?.company) {
      return normalizeCompanyRecord({ ...normalized, ...data.data.company });
    }
  } catch (_) {
    // keep list row if single-record fetch fails
  }
  return normalized;
}

const CompanyDetails = ({ userRole, userEmail }) => {
  const [companies, setCompanies] = useState([]);
  const [loading, setLoading] = useState(false);
  const [message, setMessage] = useState('');
  const [showForm, setShowForm] = useState(false);
  const [editingId, setEditingId] = useState(null);
  const [editingCompany, setEditingCompany] = useState(null);
  const [form, setForm] = useState(initialForm);
  const [uploadingDoc, setUploadingDoc] = useState(null);
  const [pendingDocFiles, setPendingDocFiles] = useState({});
  const [uploadedDocNames, setUploadedDocNames] = useState({});
  const [tableSearch, setTableSearch] = useState('');
  const [selectedIds, setSelectedIds] = useState(() => new Set());
  const [tablePage, setTablePage] = useState(1);
  const [viewOnly, setViewOnly] = useState(false);
  const [formErrors, setFormErrors] = useState({});
  const [toast, setToast] = useState('');
  const [bulkDeleting, setBulkDeleting] = useState(false);
  const [importing, setImporting] = useState(false);
  const toastTimerRef = useRef(null);
  const importFileRef = useRef(null);

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

  useEffect(
    () => () => {
      if (toastTimerRef.current) clearTimeout(toastTimerRef.current);
    },
    []
  );

  const fetchCompanies = useCallback(async (prioritizeId = null, options = {}) => {
    const { useCacheFirst = false, silentRefresh = false } = options;

    if (useCacheFirst) {
      try {
        const cached = localStorage.getItem('companyDetailsData_v3');
        if (cached) {
          const parsed = JSON.parse(cached);
          if (Array.isArray(parsed) && parsed.length > 0) {
            setCompanies(parsed);
          }
        }
      } catch (_) {}
    }

    if (!silentRefresh) setLoading(true);
    setMessage('');
    try {
      const res = await fetch(`${API_BASE}/company`);
      const data = await res.json();
      if (data.status === 'success' && data.data && Array.isArray(data.data.companyDetails)) {
        const list = data.data.companyDetails.map(normalizeCompanyRecord);
        if (prioritizeId != null && prioritizeId !== '') {
          const key = String(prioritizeId).trim();
          const idx = list.findIndex((c) => companyRowId(c) === key);
          if (idx > 0) {
            const [moved] = list.splice(idx, 1);
            list.unshift(moved);
          }
        }
        setCompanies(list);
        localStorage.setItem('companyDetailsData_v3', JSON.stringify(list));
      } else {
        setCompanies([]);
      }
    } catch (err) {
      setMessage('Failed to load company details');
      setCompanies([]);
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    // Show cached data immediately on navigation, then refresh in background.
    fetchCompanies(null, { useCacheFirst: true, silentRefresh: true });
  }, [fetchCompanies]);

  useEffect(() => {
    const valid = new Set(companies.map((c) => String(c.id)));
    setSelectedIds((prev) => {
      const next = new Set();
      prev.forEach((id) => {
        if (valid.has(id)) next.add(id);
      });
      return next;
    });
  }, [companies]);

  const openAdd = () => {
    setViewOnly(false);
    setEditingId(null);
    setEditingCompany(null);
    setForm(initialForm);
    setPendingDocFiles({});
    setUploadedDocNames({});
    setFormErrors({});
    setShowForm(true);
  };

  const openView = async (company) => {
    const record = await enrichCompanyForForm(company);
    setViewOnly(true);
    setEditingId(companyRowId(record) || record.id);
    setEditingCompany(record);
    setForm(buildFormFromCompany(record));
    setPendingDocFiles({});
    setUploadedDocNames(record?.documentFileNames || {});
    setFormErrors({});
    setShowForm(true);
    loadDocNamesForCompany(record);
  };

  const openEdit = async (company) => {
    const record = await enrichCompanyForForm(company);
    setViewOnly(false);
    setEditingId(companyRowId(record) || record.id);
    setEditingCompany(record);
    setPendingDocFiles({});
    setForm(buildFormFromCompany(record));
    setFormErrors({});
    setUploadedDocNames(record?.documentFileNames || {});
    setShowForm(true);
    loadDocNamesForCompany(record);
  };

  const closeForm = useCallback(() => {
    setShowForm(false);
    setViewOnly(false);
    setEditingId(null);
    setEditingCompany(null);
    setForm(initialForm);
    setPendingDocFiles({});
    setUploadedDocNames({});
    setFormErrors({});
  }, []);

  useEffect(() => {
    const onCloseModal = () => closeForm();
    window.addEventListener(COMPANY_DETAILS_CLOSE_MODAL_EVENT, onCloseModal);
    return () => window.removeEventListener(COMPANY_DETAILS_CLOSE_MODAL_EVENT, onCloseModal);
  }, [closeForm]);

  const getFileIdForDocType = (docType) => {
    if (!editingCompany) return null;
    const key = DOC_TYPE_TO_COMPANY_KEY_UI[docType] || docType;
    return editingCompany[key] || editingCompany[docType] || null;
  };

  const loadDocNamesForCompany = useCallback(async (company) => {
    const id = companyRowId(company);
    if (!id) return;

    const namesFromList =
      company?.documentFileNames && typeof company.documentFileNames === 'object'
        ? { ...company.documentFileNames }
        : {};

    const metaByDoc = {};
    await Promise.all(
      DOC_TYPES.map(async ({ docType }) => {
        try {
          const res = await fetch(
            `${API_BASE}/company/${encodeURIComponent(id)}/file-meta/${encodeURIComponent(docType)}`,
            { cache: 'no-store' }
          );
          const data = await res.json().catch(() => null);
          if (data?.status === 'success' && data.data) {
            const fid = data.data.fileId;
            const hasId = fid != null && String(fid).trim() !== '';
            metaByDoc[docType] = {
              fileId: hasId ? fid : null,
              fileName: String(data.data.fileName || '').trim()
            };
          }
        } catch (_) {
          // Ignore per-doc failures.
        }
      })
    );

    const idsPatch = {};
    for (const { docType } of DOC_TYPES) {
      const m = metaByDoc[docType];
      if (m?.fileId) {
        const k = DOC_TYPE_TO_COMPANY_KEY_UI[docType];
        if (k) idsPatch[k] = m.fileId;
      }
    }

    setUploadedDocNames((prev) => {
      const next = { ...namesFromList, ...prev };
      for (const { docType } of DOC_TYPES) {
        const m = metaByDoc[docType];
        if (!m) continue;
        if (m.fileName) {
          next[docType] = m.fileName;
        } else if (m.fileId && !next[docType]) {
          next[docType] = 'File attached';
        }
      }
      return next;
    });

    if (Object.keys(idsPatch).length > 0) {
      setEditingCompany((prev) => {
        if (!prev || companyRowId(prev) !== id) return prev;
        return { ...prev, ...idsPatch };
      });
    }
  }, []);

  const handleFileUpload = async (docType, file) => {
    if (!editingId || !file) return;
    setMessage('');
    setUploadingDoc(docType);
    try {
      const formData = new FormData();
      formData.append('file', file);
      const res = await fetch(`${API_BASE}/company/${editingId}/upload/${docType}`, {
        method: 'POST',
        body: formData
      });
      const data = await res.json();
      if (data.status === 'success') {
        setMessage(`${DOC_TYPES.find(d => d.docType === docType)?.label || docType} uploaded`);
        const uploadedFileName = String(data.fileName || file.name || '').trim();
        if (uploadedFileName) {
          setUploadedDocNames((prev) => ({ ...prev, [docType]: uploadedFileName }));
        }
        const list = await fetch(`${API_BASE}/company`).then(r => r.json());
        if (list.status === 'success' && list.data?.companyDetails) {
          const updated = list.data.companyDetails.find(c => String(c.id) === String(editingId));
          if (updated) setEditingCompany(updated);
          setCompanies(list.data.companyDetails);
        }
      } else {
        setMessage(data.message || 'Upload failed');
      }
    } catch (err) {
      setMessage(err.message || 'Upload failed');
    } finally {
      setUploadingDoc(null);
    }
  };

  const handleFileDownload = (docType) => {
    if (!editingId) return;
    const fileId = getFileIdForDocType(docType);
    if (!fileId) return;
    window.open(`${API_BASE}/company/${editingId}/file/${docType}`, '_blank');
  };

  const handleFileRemove = async (docType) => {
    if (!editingId || !window.confirm('Remove this document?')) return;
    setMessage('');
    try {
      const res = await fetch(`${API_BASE}/company/${editingId}/file/${docType}`, { method: 'DELETE' });
      const data = await res.json();
      if (data.status === 'success') {
        setMessage('Document removed');
        const list = await fetch(`${API_BASE}/company`).then(r => r.json());
        if (list.status === 'success' && list.data?.companyDetails) {
          const updated = list.data.companyDetails.find(c => String(c.id) === String(editingId));
          if (updated) setEditingCompany(updated);
          setCompanies(list.data.companyDetails);
        }
      } else {
        setMessage(data.message || 'Remove failed');
      }
    } catch (err) {
      setMessage(err.message || 'Remove failed');
    }
  };

  const handleChange = (e) => {
    const { name, value } = e.target;
    const nextValue = sanitizeCompanyFormField(name, value);
    setForm((prev) => ({ ...prev, [name]: nextValue }));
    setFormErrors((prev) => {
      if (!prev[name]) return prev;
      const next = { ...prev };
      delete next[name];
      return next;
    });
  };

  const handleFieldBlur = useCallback((e) => {
    if (viewOnly) return;
    const target = e.target;
    if (!target || target.type === 'file') return;
    const { name } = target;
    if (!name || !(name in initialForm)) return;

    const nextForm =
      name === 'incorprationDate'
        ? { ...form, incorprationDate: normalizeIncorporationDateOnBlurString(form.incorprationDate) }
        : form;

    if (nextForm !== form) {
      setForm(nextForm);
    }

    const fieldError = validateCompanyFormValues(nextForm)[name];
    setFormErrors((prev) => {
      if (fieldError) {
        if (prev[name] === fieldError) return prev;
        return { ...prev, [name]: fieldError };
      }
      if (!prev[name]) return prev;
      const next = { ...prev };
      delete next[name];
      return next;
    });
  }, [form, viewOnly]);

  const incorporationNativePickerRef = useRef(null);

  const openNativeIncorporationDatePicker = () => {
    if (viewOnly) return;
    const el = incorporationNativePickerRef.current;
    if (!el) return;
    try {
      if (typeof el.showPicker === 'function') {
        el.showPicker();
      } else {
        el.click();
      }
    } catch {
      el.click();
    }
  };

  const handleNativeIncorporationDateChange = (e) => {
    const v = e.target.value;
    setForm((prev) => ({ ...prev, incorprationDate: v }));
    setFormErrors((prev) => {
      if (!prev.incorprationDate) return prev;
      const next = { ...prev };
      delete next.incorprationDate;
      return next;
    });
  };

  const scrollToFirstFormError = useCallback((errs) => {
    const firstKey = COMPANY_FORM_FIELD_ORDER.find((k) => errs[k]);
    if (!firstKey) return;
    const el = document.getElementById(companyInputId(firstKey));
    if (el) {
      el.scrollIntoView({ behavior: 'smooth', block: 'center' });
      if (typeof el.focus === 'function') {
        try {
          el.focus({ preventScroll: true });
        } catch {
          el.focus();
        }
      }
    }
  }, []);

  const handleSubmit = async (e) => {
    e.preventDefault();
    setMessage('');
    if (!viewOnly) {
      const errs = validateCompanyFormValues(form);
      if (Object.keys(errs).length > 0) {
        setFormErrors(errs);
        setTimeout(() => scrollToFirstFormError(errs), 50);
        return;
      }
      setFormErrors({});
    }
    try {
      const incParsed = parseFlexibleIncorporationDate(String(form.incorprationDate || '').trim());
      const payload = {
        companyName: form.companyName.trim(),
        companyPANNumber: form.companyPANNumber.trim(),
        companyAddress: form.companyAddress.trim(),
        city: form.city.trim(),
        state: form.state.trim(),
        postalcode: form.postalcode.trim(),
        incorprationDate: incParsed || null,
        incorporationNo: form.incorporationNo.trim(),
        gstNo: form.gstNo.trim(),
        pfNo: form.pfNo.trim(),
        esiNo: form.esiNo.trim(),
        companyMail: form.companyMail.trim(),
        companyPhoneNumber: form.companyPhoneNumber.trim(),
        directorName: form.directorName.trim(),
        directorPhoneNumber: form.directorPhoneNumber.trim(),
        directorMail: form.directorMail.trim(),
        directorAddress: form.directorAddress.trim(),
        ownerName: form.ownerName.trim(),
        ownerPAN: form.ownerPAN.trim(),
        ownerAaadhar: form.ownerAaadhar.trim(),
        ownerDesignation: form.ownerDesignation.trim(),
        safetyOfficerName: form.safetyOfficerName.trim(),
        safetyOfficerPhone: form.safetyOfficerPhone.trim(),
        doctroName: editingId ? String(editingCompany?.doctroName ?? '').trim() : '',
        doctroPhone: editingId ? String(editingCompany?.doctroPhone ?? '').trim() : ''
      };

      if (editingId) {
        const res = await fetch(`${API_BASE}/company/${editingId}`, {
          method: 'PUT',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(payload)
        });
        const data = await res.json();
        if (data.status === 'success') {
          const updatedId = editingId;
          showToast('Update Successfully');
          closeForm();
          setTablePage(1);
          fetchCompanies(updatedId);
        } else {
          setMessage(data.message || 'Update failed');
        }
      } else {
        const res = await fetch(`${API_BASE}/company`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(payload)
        });
        const data = await res.json();
        if (data.status === 'success') {
          const createdRow = data.data?.company;
          let createdId = null;
          const newId = createdRow?.ROWID ?? createdRow?.Company?.ROWID ?? createdRow?.id;
          if (newId != null && typeof newId !== 'object') {
            createdId = String(newId);
          }
          if (newId != null && typeof newId !== 'object') {
            for (const { docType } of DOC_TYPES) {
              const file = pendingDocFiles[docType];
              if (file) {
                try {
                  const fd = new FormData();
                  fd.append('file', file);
                  await fetch(`${API_BASE}/company/${newId}/upload/${docType}`, { method: 'POST', body: fd });
                } catch (err) {
                  setMessage((m) => (m ? `${m} Upload failed for ${docType}.` : `Company saved. Upload failed for ${docType}.`));
                }
              }
            }
          }
          showToast('Added Successfully');
          closeForm();
          setTablePage(1);
          fetchCompanies(createdId);
        } else {
          setMessage(data.message || 'Add failed');
        }
      }
    } catch (err) {
      setMessage(err.message || 'Request failed');
    }
  };

  const handleDelete = async (id) => {
    if (!window.confirm('Delete this company?')) return;
    setMessage('');
    try {
      const res = await fetch(`${API_BASE}/company/${id}`, { method: 'DELETE' });
      const data = await res.json();
      if (data.status === 'success') {
        showToast('Deleted Successfully');
        fetchCompanies();
      } else {
        setMessage(data.message || 'Delete failed');
      }
    } catch (err) {
      setMessage(err.message || 'Delete failed');
    }
  };

  const handleDeleteSelected = async () => {
    const ids = [...selectedIds];
    if (ids.length === 0 || bulkDeleting) return;
    if (!window.confirm(`Delete ${ids.length} selected compan${ids.length === 1 ? 'y' : 'ies'}?`)) return;
    setMessage('');
    setBulkDeleting(true);
    try {
      const results = await Promise.allSettled(
        ids.map(async (id) => {
          const res = await fetch(`${API_BASE}/company/${id}`, { method: 'DELETE' });
          const data = await res.json();
          return data.status === 'success';
        })
      );
      const deletedCount = results.filter((r) => r.status === 'fulfilled' && r.value === true).length;
      const failedCount = results.length - deletedCount;
      if (deletedCount > 0) {
        showToast('Deleted Successfully');
      }
      if (failedCount > 0) {
        setMessage(`${failedCount} delete request(s) failed. Please try again.`);
      }
      setSelectedIds(new Set());
      fetchCompanies();
    } finally {
      setBulkDeleting(false);
    }
  };

  /** City list for search — master list + cities already saved on companies. */
  const citySelectOptions = useMemo(() => {
    const set = new Set(INDIAN_CITIES);
    companies.forEach((c) => {
      const savedCity = String(c.city ?? '').trim();
      if (savedCity) set.add(savedCity);
    });
    const cur = String(form.city || '').trim();
    if (cur) set.add(cur);
    return [...set].sort((a, b) => a.localeCompare(b, undefined, { sensitivity: 'base' }));
  }, [companies, form.city]);

  /** State list — standard list + states already saved on companies. */
  const stateSelectOptions = useMemo(() => {
    const set = new Set(INDIAN_STATES);
    companies.forEach((c) => {
      const savedState = String(c.state ?? '').trim();
      if (savedState) set.add(savedState);
    });
    const cur = String(form.state || '').trim();
    if (cur) set.add(cur);
    return [...set].sort((a, b) => a.localeCompare(b, undefined, { sensitivity: 'base' }));
  }, [companies, form.state]);

  const filteredCompanies = useMemo(() => {
    const q = tableSearch.trim().toLowerCase();
    if (!q) return companies;
    return companies.filter((c) => {
      const hay = [
        c.companyName,
        c.companyMail,
        c.companyPhoneNumber,
        c.directorName,
        c.directorMail,
        c.companyPANNumber,
        c.gstNo,
        c.incorporationNo,
        c.pfNo,
        c.esiNo,
      ]
        .map((x) => String(x || '').toLowerCase())
        .join(' ');
      return hay.includes(q);
    });
  }, [companies, tableSearch]);

  const displayCompanies = useMemo(() => [...filteredCompanies], [filteredCompanies]);

  const totalTablePages = Math.max(1, Math.ceil(displayCompanies.length / COMPANY_TABLE_PAGE_SIZE));

  useEffect(() => {
    setTablePage(1);
  }, [tableSearch]);

  useEffect(() => {
    setTablePage((p) => Math.min(p, totalTablePages));
  }, [totalTablePages]);

  const effectiveTablePage = Math.min(tablePage, totalTablePages);

  const pagedCompanies = useMemo(() => {
    const start = (effectiveTablePage - 1) * COMPANY_TABLE_PAGE_SIZE;
    return displayCompanies.slice(start, start + COMPANY_TABLE_PAGE_SIZE);
  }, [displayCompanies, effectiveTablePage]);

  const paginationItems = useMemo(
    () => buildPaginationItems(effectiveTablePage, totalTablePages),
    [effectiveTablePage, totalTablePages]
  );

  const allPageSelected =
    pagedCompanies.length > 0 && pagedCompanies.every((c) => selectedIds.has(String(c.id)));
  const somePageSelected = pagedCompanies.some((c) => selectedIds.has(String(c.id)));

  const toggleSelectAllDisplay = () => {
    if (allPageSelected) {
      setSelectedIds((prev) => {
        const next = new Set(prev);
        pagedCompanies.forEach((c) => next.delete(String(c.id)));
        return next;
      });
    } else {
      setSelectedIds((prev) => {
        const next = new Set(prev);
        pagedCompanies.forEach((c) => next.add(String(c.id)));
        return next;
      });
    }
  };

  const formatDisplayDate = (val) => {
    if (!val) return '';
    const d = new Date(val);
    if (isNaN(d.getTime())) return '';
    const day = String(d.getDate()).padStart(2, '0');
    const mo = String(d.getMonth() + 1).padStart(2, '0');
    return `${day}/${mo}/${d.getFullYear()}`;
  };

  const exportCompaniesCsv = () => {
    /** Leading tab inside a quoted field forces Excel to treat the cell as text (avoids 5.64E+09, ###, etc.). */
    const CSV_EXCEL_TEXT_KEYS = new Set([
      'companyPhoneNumber',
      'directorPhoneNumber',
      'safetyOfficerPhone',
      'ownerAaadhar',
      'incorprationDate',
      'incorporationNo',
      'companyPANNumber',
      'ownerPAN',
      'gstNo',
      'pfNo',
      'esiNo',
      'postalcode',
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
    const cellValue = (co, key) => {
      if (key === 'incorprationDate') {
        const dateRaw = co.incorprationDate || co.incorporationDate;
        if (!dateRaw) return '';
        const d = new Date(dateRaw);
        if (isNaN(d.getTime())) return '';
        const day = String(d.getDate()).padStart(2, '0');
        const m = String(d.getMonth() + 1).padStart(2, '0');
        return `${day}/${m}/${d.getFullYear()}`;
      }
      return co[key];
    };
    const strCmp = (a, b, field) =>
      String(a[field] ?? '')
        .toLowerCase()
        .localeCompare(String(b[field] ?? '').toLowerCase(), undefined, { sensitivity: 'base' });
    const list = [...companies].sort((a, b) => strCmp(a, b, 'companyName'));
    const header = COMPANY_CSV_COLS.map((c) => escLabel(c.label)).join(',');
    const rows = list.map((co) => COMPANY_CSV_COLS.map((c) => escCsvCell(cellValue(co, c.key), c.key)).join(','));
    const csv = '\ufeff' + [header, ...rows].join('\r\n');
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `company-details-${new Date().toISOString().slice(0, 10)}.csv`;
    a.click();
    URL.revokeObjectURL(url);
  };

  const handleImportCompanies = useCallback(
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
        const arrayBuffer = await file.arrayBuffer();
        const importRows = parseCompanyImportWorkbook(arrayBuffer);
        if (!importRows.length) {
          setMessage('No data rows found in the file.');
          return;
        }

        // Refresh so upsert matches latest DB rows (not a stale client list)
        let companiesForMatch = companies;
        try {
          const res = await fetch(`${API_BASE}/company`);
          const data = await res.json();
          if (data.status === 'success' && Array.isArray(data.data?.companyDetails)) {
            companiesForMatch = data.data.companyDetails.map(normalizeCompanyRecord);
            setCompanies(companiesForMatch);
          }
        } catch (_) {
          // keep in-memory list
        }

        const existingById = new Map();
        const existingByName = new Map();
        companiesForMatch.forEach((c) => {
          const id = companyRowId(c);
          if (id) existingById.set(id, c);
          const nameKey = companyRecordNameKey(c);
          // Keep the first match when duplicates already exist
          if (nameKey && !existingByName.has(nameKey)) existingByName.set(nameKey, c);
        });

        let created = 0;
        let updated = 0;
        const errors = [];

        for (const row of importRows) {
          const formValues = {
            ...initialForm,
            companyName: row.companyName,
            companyMail: row.companyMail,
            companyPhoneNumber: row.companyPhoneNumber,
            companyAddress: row.companyAddress,
            city: row.city,
            state: row.state,
            postalcode: row.postalcode,
            incorprationDate: row.incorprationDate,
            incorporationNo: row.incorporationNo,
            companyPANNumber: row.companyPANNumber,
            gstNo: row.gstNo,
            pfNo: row.pfNo,
            esiNo: row.esiNo,
            directorName: row.directorName,
            directorMail: row.directorMail,
            directorPhoneNumber: row.directorPhoneNumber,
            directorAddress: row.directorAddress,
            ownerName: row.ownerName,
            ownerPAN: row.ownerPAN,
            ownerAaadhar: row.ownerAaadhar,
            ownerDesignation: row.ownerDesignation,
            safetyOfficerName: row.safetyOfficerName,
            safetyOfficerPhone: row.safetyOfficerPhone
          };
          const errs = validateCompanyFormValues(formValues);
          if (Object.keys(errs).length > 0) {
            const first = Object.values(errs)[0];
            errors.push(`Row ${row.__row}: ${first}`);
            continue;
          }

          const payload = buildCompanyPayloadFromImportRow(row);
          const rowId = String(row.id || '').trim();
          const nameKey = companyImportNameKey(row.companyName);
          let matchId = '';
          let existing = null;
          if (rowId && existingById.has(rowId)) {
            matchId = rowId;
            existing = existingById.get(rowId);
          } else if (nameKey && existingByName.has(nameKey)) {
            existing = existingByName.get(nameKey);
            matchId = companyRowId(existing);
          }

          try {
            if (matchId) {
              payload.doctroName = String(existing?.doctroName ?? '').trim();
              payload.doctroPhone = String(existing?.doctroPhone ?? '').trim();
              const res = await fetch(`${API_BASE}/company/${encodeURIComponent(matchId)}`, {
                method: 'PUT',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(payload)
              });
              const data = await res.json();
              if (data.status === 'success') {
                updated += 1;
                const updatedCo = { ...(existing || {}), ...payload, id: matchId };
                existingById.set(matchId, updatedCo);
                if (nameKey) existingByName.set(nameKey, updatedCo);
              } else {
                errors.push(`Row ${row.__row}: ${data.message || 'Update failed'}`);
              }
            } else {
              const res = await fetch(`${API_BASE}/company`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(payload)
              });
              const data = await res.json();
              if (data.status === 'success') {
                created += 1;
                const createdRow = data.data?.company;
                const createdId = companyRowId(createdRow);
                const createdCo = { ...payload, id: createdId };
                if (createdId) existingById.set(createdId, createdCo);
                // Prevent later rows in the same file from creating another duplicate
                if (nameKey) existingByName.set(nameKey, createdCo);
              } else {
                errors.push(`Row ${row.__row}: ${data.message || 'Add failed'}`);
              }
            }
          } catch (err) {
            errors.push(`Row ${row.__row}: ${err.message || 'Request failed'}`);
          }
        }

        await fetchCompanies();
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
          setMessage('No companies were imported.');
        }
      } catch (err) {
        setMessage(err.message || 'Import failed');
      } finally {
        setImporting(false);
      }
    },
    [companies, fetchCompanies, showToast]
  );

  const toggleRowSelected = (id) => {
    const key = String(id);
    setSelectedIds((prev) => {
      const next = new Set(prev);
      if (next.has(key)) next.delete(key);
      else next.add(key);
      return next;
    });
  };

  const modalTitle = viewOnly ? 'View Company' : editingId ? 'Edit Company' : 'Add New Company';
  const breadcrumbFormCrumb = viewOnly ? 'View Company' : editingId ? 'Edit Company' : 'Add Company';
  const sectionErrorCount = (idx) => countSectionErrors(formErrors, SECTION_FIELD_GROUPS[idx].keys);

  return (
    <div className={`company-details-page company-details-hcm-route${showForm ? ' company-details-page-form-open' : ''}`}>
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
      {!showForm && message ? <div className="company-details-message company-details-message--flush">{message}</div> : null}
      <header className="company-details-page-heading">
        <h1 className="company-details-page-title">Company Details</h1>
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
          {showForm ? (
            <button
              type="button"
              className="company-details-bc-link company-details-bc-as-link"
              onClick={closeForm}
              aria-label="Return to company list"
            >
              Company Details
            </button>
          ) : (
            <span className="company-details-bc-current company-details-bc-active">Company Details</span>
          )}
          {showForm ? (
            <>
              <span className="company-details-bc-sep" aria-hidden>
                ›
              </span>
              <span className="company-details-bc-current company-details-bc-active">{breadcrumbFormCrumb}</span>
            </>
          ) : null}
        </nav>
      </header>
      {showForm ? (
        <div className="company-details-form-center">
          {message ? <div className="company-details-message company-details-message--modal">{message}</div> : null}
          <div className="company-details-modal" role="dialog" aria-labelledby="company-details-modal-title" aria-modal="true">
            <div className="company-details-modal-header">
              <h2 id="company-details-modal-title" className="company-details-modal-title">
                {modalTitle}
              </h2>
              <button type="button" className="company-details-modal-close" onClick={closeForm} aria-label="Close">
                ×
              </button>
            </div>
            <form
              onSubmit={viewOnly ? (e) => e.preventDefault() : handleSubmit}
              onBlur={handleFieldBlur}
              className="company-details-modal-form"
              noValidate
            >
              <div className="company-details-modal-scroll">
              <div className="company-details-modal-body">
                <section className="company-details-section-card">
                  <header className="company-details-section-head">
                    <h3 className="company-details-section-title">
                      <svg className="company-details-section-title-icon" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" aria-hidden>
                        <path d="M9 5H7a2 2 0 0 0-2 2v12a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V7a2 2 0 0 0-2-2h-2M9 5a2 2 0 0 0 2 2h2a2 2 0 0 0 2-2M9 5a2 2 0 0 1 2-2h2a2 2 0 0 1 2 2" />
                      </svg>
                      Company info
                    </h3>
                    {/* Error badge removed as requested */}
                  </header>
                  <div className="company-details-fields-grid company-details-fields-grid--company-info">
                    <div className={`company-details-field${formErrors.companyName ? ' company-details-field--error' : ''}`}>
                      <label htmlFor="cd-companyName">
                        Name <span className="required" aria-hidden="true">*</span>
                      </label>
                      <input id="cd-companyName" name="companyName" value={form.companyName} onChange={handleChange} placeholder="Enter name" disabled={viewOnly} />
                      {formErrors.companyName ? (
                        <p className="company-details-field-error" role="alert">
                          <span aria-hidden="true">⚠️</span> {formErrors.companyName}
                        </p>
                      ) : null}
                    </div>
                    <div className={`company-details-field${formErrors.companyMail ? ' company-details-field--error' : ''}`}>
                      <label htmlFor="cd-companyMail">
                        Mail Id <span className="required" aria-hidden="true">*</span>
                      </label>
                      <input
                        id="cd-companyMail"
                        name="companyMail"
                        type="email"
                        value={form.companyMail}
                        onChange={handleChange}
                        placeholder="Enter mail id"
                        disabled={viewOnly}
                        maxLength={254}
                      />
                      {formErrors.companyMail ? (
                        <p className="company-details-field-error" role="alert">
                          <span aria-hidden="true">⚠️</span> {formErrors.companyMail}
                        </p>
                      ) : null}
                    </div>
                    <div className={`company-details-field${formErrors.companyPhoneNumber ? ' company-details-field--error' : ''}`}>
                      <label htmlFor="cd-companyPhoneNumber">
                        Moblie Number <span className="required" aria-hidden="true">*</span>
                      </label>
                      <input
                        id="cd-companyPhoneNumber"
                        name="companyPhoneNumber"
                        type="tel"
                        inputMode="numeric"
                        autoComplete="tel"
                        value={form.companyPhoneNumber}
                        onChange={handleChange}
                        placeholder="Enter mobile number"
                        disabled={viewOnly}
                        maxLength={10}
                        title="Enter exactly 10 digits"
                      />
                      {formErrors.companyPhoneNumber ? (
                        <p className="company-details-field-error" role="alert">
                          <span aria-hidden="true">⚠️</span> {formErrors.companyPhoneNumber}
                        </p>
                      ) : null}
                    </div>
                    <div className={`company-details-field company-details-field--full-row${formErrors.companyAddress ? ' company-details-field--error' : ''}`}>
                      <label htmlFor="cd-companyAddress">
                        Address <span className="required" aria-hidden="true">*</span>
                      </label>
                      <input id="cd-companyAddress" name="companyAddress" value={form.companyAddress} onChange={handleChange} placeholder="Enter address" disabled={viewOnly} />
                      {formErrors.companyAddress ? (
                        <p className="company-details-field-error" role="alert">
                          <span aria-hidden="true">⚠️</span> {formErrors.companyAddress}
                        </p>
                      ) : null}
                    </div>
                    <div className="company-details-field">
                      <label htmlFor="cd-city">City</label>
                      <CityCombobox
                        id="cd-city"
                        name="city"
                        value={form.city}
                        options={citySelectOptions}
                        onChange={handleChange}
                        onBlur={handleFieldBlur}
                        disabled={viewOnly}
                        placeholder="Type letter to filter cities"
                        otherPlaceholder="Enter your city name"
                      />
                    </div>
                    <div className="company-details-field">
                      <label htmlFor="cd-state">State</label>
                      <StateCombobox
                        id="cd-state"
                        name="state"
                        value={form.state}
                        options={stateSelectOptions}
                        onChange={handleChange}
                        onBlur={handleFieldBlur}
                        disabled={viewOnly}
                        placeholder="Type letter to filter states"
                        otherPlaceholder="Enter your state name"
                      />
                    </div>
                    <div className="company-details-field">
                      <label htmlFor="cd-postalcode">Postal code</label>
                      <input
                        id="cd-postalcode"
                        name="postalcode"
                        value={form.postalcode}
                        onChange={handleChange}
                        placeholder="Enter postal code"
                        disabled={viewOnly}
                        inputMode="numeric"
                        maxLength={10}
                      />
                    </div>
                  </div>
                </section>

                <section className="company-details-section-card">
                  <header className="company-details-section-head">
                    <h3 className="company-details-section-title">
                      <svg className="company-details-section-title-icon" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" aria-hidden>
                        <rect x="2" y="7" width="20" height="14" rx="2" ry="2" />
                        <path d="M16 21V5a2 2 0 0 0-2-2h-4a2 2 0 0 0-2 2v16" />
                      </svg>
                      Registration & tax
                    </h3>
                    <div className="company-details-section-head-right">
                      {/* Error badge and computer icon removed as requested */}
                    </div>
                  </header>
                  <div className="company-details-fields-grid">
                    <div className={`company-details-field${formErrors.incorprationDate ? ' company-details-field--error' : ''}`}>
                      <label htmlFor="cd-incorprationDate">
                        Incorporation date <span className="required" aria-hidden="true">*</span>
                      </label>
                      <div className="company-details-date-wrap">
                        <input
                          id="cd-incorprationDate"
                          name="incorprationDate"
                          type="text"
                          inputMode="numeric"
                          autoComplete="off"
                          placeholder="DD/MM/YYYY"
                          value={incorporationInputDisplay(form.incorprationDate)}
                          onChange={handleChange}
                          className="company-details-date-input company-details-field-input"
                          title="Enter incorporation date as DD/MM/YYYY, or use the calendar"
                          disabled={viewOnly}
                        />
                        <input
                          ref={incorporationNativePickerRef}
                          type="date"
                          className="company-details-date-native-sr"
                          tabIndex={-1}
                          aria-hidden="true"
                          value={isoForNativeDatePickerValue(form.incorprationDate)}
                          onChange={handleNativeIncorporationDateChange}
                          disabled={viewOnly}
                        />
                        <button
                          type="button"
                          className="company-details-date-picker-btn"
                          onClick={openNativeIncorporationDatePicker}
                          disabled={viewOnly}
                          title="Open calendar"
                          aria-label="Open calendar to choose incorporation date"
                        >
                          <Calendar size={18} strokeWidth={2} aria-hidden />
                        </button>
                      </div>
                      {formErrors.incorprationDate ? (
                        <p className="company-details-field-error" role="alert">
                          <span aria-hidden="true">⚠️</span> {formErrors.incorprationDate}
                        </p>
                      ) : null}
                    </div>
                    <div className={`company-details-field${formErrors.incorporationNo ? ' company-details-field--error' : ''}`}>
                      <label htmlFor="cd-incorporationNo">
                        Incorporation number <span className="required" aria-hidden="true">*</span>
                      </label>
                      <input id="cd-incorporationNo" name="incorporationNo" value={form.incorporationNo} onChange={handleChange} placeholder="Enter incorporation number" disabled={viewOnly} />
                      {formErrors.incorporationNo ? (
                        <p className="company-details-field-error" role="alert">
                          <span aria-hidden="true">⚠️</span> {formErrors.incorporationNo}
                        </p>
                      ) : null}
                    </div>
                    <div className={`company-details-field${formErrors.companyPANNumber ? ' company-details-field--error' : ''}`}>
                      <label htmlFor="cd-companyPANNumber">
                        PAN <span className="required" aria-hidden="true">*</span>
                      </label>
                      <input
                        id="cd-companyPANNumber"
                        name="companyPANNumber"
                        value={form.companyPANNumber}
                        onChange={handleChange}
                        placeholder="e.g. ABCDE1234F"
                        disabled={viewOnly}
                        maxLength={10}
                        autoCapitalize="characters"
                      />
                      {formErrors.companyPANNumber ? (
                        <p className="company-details-field-error" role="alert">
                          <span aria-hidden="true">⚠️</span> {formErrors.companyPANNumber}
                        </p>
                      ) : null}
                    </div>
                    <div className={`company-details-field${formErrors.gstNo ? ' company-details-field--error' : ''}`}>
                      <label htmlFor="cd-gstNo">GST number</label>
                      <input
                        id="cd-gstNo"
                        name="gstNo"
                        value={form.gstNo}
                        onChange={handleChange}
                        placeholder="e.g. 29AAACH7409R1ZX"
                        disabled={viewOnly}
                        maxLength={15}
                        autoCapitalize="characters"
                      />
                      {formErrors.gstNo ? (
                        <p className="company-details-field-error" role="alert">
                          <span aria-hidden="true">⚠️</span> {formErrors.gstNo}
                        </p>
                      ) : null}
                    </div>
                    <div className={`company-details-field${formErrors.pfNo ? ' company-details-field--error' : ''}`}>
                      <label htmlFor="cd-pfNo">PF number</label>
                      <input
                        id="cd-pfNo"
                        name="pfNo"
                        value={form.pfNo}
                        onChange={handleChange}
                        placeholder="e.g. MH/BAN/1234567/000/1234567"
                        disabled={viewOnly}
                        maxLength={26}
                        autoCapitalize="characters"
                      />
                      {formErrors.pfNo ? (
                        <p className="company-details-field-error" role="alert">
                          <span aria-hidden="true">⚠️</span> {formErrors.pfNo}
                        </p>
                      ) : null}
                    </div>
                    <div className={`company-details-field${formErrors.esiNo ? ' company-details-field--error' : ''}`}>
                      <label htmlFor="cd-esiNo">ESI number</label>
                      <input
                        id="cd-esiNo"
                        name="esiNo"
                        value={form.esiNo}
                        onChange={handleChange}
                        placeholder="e.g. 31-00-123456-000-0001"
                        disabled={viewOnly}
                        maxLength={21}
                        inputMode="numeric"
                      />
                      {formErrors.esiNo ? (
                        <p className="company-details-field-error" role="alert">
                          <span aria-hidden="true">⚠️</span> {formErrors.esiNo}
                        </p>
                      ) : null}
                    </div>
                  </div>
                </section>

                <section className="company-details-section-card">
                  <header className="company-details-section-head">
                    <h3 className="company-details-section-title">
                      <svg className="company-details-section-title-icon" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" aria-hidden>
                        <path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2" />
                        <circle cx="12" cy="7" r="4" />
                      </svg>
                      Director info
                    </h3>
                    <div className="company-details-section-head-right">
                      {!viewOnly && sectionErrorCount(2) > 0 && (
                        <span className="company-details-section-error-badge">
                          {sectionErrorCount(2)} error{sectionErrorCount(2) === 1 ? '' : 's'}
                        </span>
                      )}
                    </div>
                  </header>
                  <div className="company-details-fields-grid company-details-fields-grid--four-in-row">
                    <div className="company-details-field">
                      <label htmlFor="cd-directorName">Name</label>
                      <input id="cd-directorName" name="directorName" value={form.directorName} onChange={handleChange} placeholder="Enter name" disabled={viewOnly} />
                    </div>
                    <div className={`company-details-field${formErrors.directorMail ? ' company-details-field--error' : ''}`}>
                      <label htmlFor="cd-directorMail">Mail Id</label>
                      <input
                        id="cd-directorMail"
                        name="directorMail"
                        type="email"
                        value={form.directorMail}
                        onChange={handleChange}
                        placeholder="Enter mail id"
                        disabled={viewOnly}
                        maxLength={254}
                      />
                      {formErrors.directorMail ? (
                        <p className="company-details-field-error" role="alert">
                          <span aria-hidden="true">⚠️</span> {formErrors.directorMail}
                        </p>
                      ) : null}
                    </div>
                    <div className={`company-details-field${formErrors.directorPhoneNumber ? ' company-details-field--error' : ''}`}>
                      <label htmlFor="cd-directorPhoneNumber">Mobile Number</label>
                      <input
                        id="cd-directorPhoneNumber"
                        name="directorPhoneNumber"
                        type="tel"
                        inputMode="numeric"
                        autoComplete="tel"
                        value={form.directorPhoneNumber}
                        onChange={handleChange}
                        placeholder="Enter mobile number"
                        disabled={viewOnly}
                        maxLength={10}
                        title="Enter exactly 10 digits"
                      />
                      {formErrors.directorPhoneNumber ? (
                        <p className="company-details-field-error" role="alert">
                          <span aria-hidden="true">⚠️</span> {formErrors.directorPhoneNumber}
                        </p>
                      ) : null}
                    </div>
                    <div className="company-details-field">
                      <label htmlFor="cd-directorAddress">Address</label>
                      <input id="cd-directorAddress" name="directorAddress" value={form.directorAddress} onChange={handleChange} placeholder="Enter address" disabled={viewOnly} />
                    </div>
                  </div>
                </section>

                <section className="company-details-section-card">
                  <header className="company-details-section-head">
                    <h3 className="company-details-section-title">
                      <svg className="company-details-section-title-icon" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" aria-hidden>
                        <path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2" />
                        <circle cx="9" cy="7" r="4" />
                        <path d="M23 21v-2a4 4 0 0 0-3-3.87" />
                        <path d="M16 3.13a4 4 0 0 1 0 7.75" />
                      </svg>
                      Owner & compliance
                    </h3>
                    <div className="company-details-section-head-right">
                      {!viewOnly && sectionErrorCount(3) > 0 && (
                        <span className="company-details-section-error-badge">
                          {sectionErrorCount(3)} error{sectionErrorCount(3) === 1 ? '' : 's'}
                        </span>
                      )}
                    </div>
                  </header>
                  <div className="company-details-fields-grid company-details-fields-grid--four-in-row">
                    <div className={`company-details-field${formErrors.ownerName ? ' company-details-field--error' : ''}`}>
                      <label htmlFor="cd-ownerName">
                        Owner / In-charge Name <span className="required" aria-hidden="true">*</span>
                      </label>
                      <input id="cd-ownerName" name="ownerName" value={form.ownerName} onChange={handleChange} placeholder="Enter owner / in-charge name" disabled={viewOnly} />
                      {formErrors.ownerName ? (
                        <p className="company-details-field-error" role="alert">
                          <span aria-hidden="true">⚠️</span> {formErrors.ownerName}
                        </p>
                      ) : null}
                    </div>
                    <div className={`company-details-field${formErrors.ownerPAN ? ' company-details-field--error' : ''}`}>
                      <label htmlFor="cd-ownerPAN">
                        Owner / In-charge PAN <span className="required" aria-hidden="true">*</span>
                      </label>
                      <input
                        id="cd-ownerPAN"
                        name="ownerPAN"
                        value={form.ownerPAN}
                        onChange={handleChange}
                        placeholder="e.g. ABCDE1234F"
                        disabled={viewOnly}
                        maxLength={10}
                        autoCapitalize="characters"
                      />
                      {formErrors.ownerPAN ? (
                        <p className="company-details-field-error" role="alert">
                          <span aria-hidden="true">⚠️</span> {formErrors.ownerPAN}
                        </p>
                      ) : null}
                    </div>
                    <div className={`company-details-field${formErrors.ownerAaadhar ? ' company-details-field--error' : ''}`}>
                      <label htmlFor="cd-ownerAaadhar">
                        Owner / In-charge Aadhar <span className="required" aria-hidden="true">*</span>
                      </label>
                      <input
                        id="cd-ownerAaadhar"
                        name="ownerAaadhar"
                        type="tel"
                        inputMode="numeric"
                        value={form.ownerAaadhar}
                        onChange={handleChange}
                        placeholder="Enter 12-digit Aadhar"
                        disabled={viewOnly}
                        maxLength={12}
                        title="Enter exactly 12 digits"
                      />
                      {formErrors.ownerAaadhar ? (
                        <p className="company-details-field-error" role="alert">
                          <span aria-hidden="true">⚠️</span> {formErrors.ownerAaadhar}
                        </p>
                      ) : null}
                    </div>
                    <div
                      className={`company-details-field company-details-field--owner-designation${formErrors.ownerDesignation ? ' company-details-field--error' : ''}`}
                    >
                      <label htmlFor="cd-ownerDesignation">
                        Owner / In-charge designation<span className="required" aria-hidden="true">*</span>
                      </label>
                      <input
                        id="cd-ownerDesignation"
                        name="ownerDesignation"
                        value={form.ownerDesignation}
                        onChange={handleChange}
                        placeholder="Enter owner / in-charge designation"
                        disabled={viewOnly}
                      />
                      {formErrors.ownerDesignation ? (
                        <p className="company-details-field-error" role="alert">
                          <span aria-hidden="true">⚠️</span> {formErrors.ownerDesignation}
                        </p>
                      ) : null}
                    </div>
                    <div className={`company-details-field${formErrors.safetyOfficerName ? ' company-details-field--error' : ''}`}>
                      <label htmlFor="cd-safetyOfficerName">Safety officer name</label>
                      <input id="cd-safetyOfficerName" name="safetyOfficerName" value={form.safetyOfficerName} onChange={handleChange} placeholder="Enter name" disabled={viewOnly} />
                      {formErrors.safetyOfficerName ? (
                        <p className="company-details-field-error" role="alert">
                          <span aria-hidden="true">⚠️</span> {formErrors.safetyOfficerName}
                        </p>
                      ) : null}
                    </div>
                    <div className={`company-details-field${formErrors.safetyOfficerPhone ? ' company-details-field--error' : ''}`}>
                      <label htmlFor="cd-safetyOfficerPhone">Safety officer phone</label>
                      <input
                        id="cd-safetyOfficerPhone"
                        name="safetyOfficerPhone"
                        type="tel"
                        inputMode="numeric"
                        autoComplete="tel"
                        value={form.safetyOfficerPhone}
                        onChange={handleChange}
                        placeholder="Enter phone (10 digits)"
                        disabled={viewOnly}
                        maxLength={10}
                        title="Enter exactly 10 digits"
                      />
                      {formErrors.safetyOfficerPhone ? (
                        <p className="company-details-field-error" role="alert">
                          <span aria-hidden="true">⚠️</span> {formErrors.safetyOfficerPhone}
                        </p>
                      ) : null}
                    </div>
                  </div>
                </section>

                <section className="company-details-section-card">
                  <header className="company-details-section-head">
                    <h3 className="company-details-section-title">
                      <svg className="company-details-section-title-icon" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" aria-hidden>
                        <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" />
                        <polyline points="14 2 14 8 20 8" />
                        <line x1="16" y1="13" x2="8" y2="13" />
                        <line x1="16" y1="17" x2="8" y2="17" />
                        <polyline points="10 9 9 9 8 9" />
                      </svg>
                      Documents
                    </h3>
                  </header>
                  {!editingId && !viewOnly && (
                    <p className="company-details-doc-hint company-details-doc-hint--section">Files upload on submit; you can change them later from Edit.</p>
                  )}
                  <div className="company-details-doc-grid company-details-doc-grid--section">
                    {DOC_TYPES.map(({ docType, label }) => {
                      const fileId = getFileIdForDocType(docType);
                      const isUploading = uploadingDoc === docType;
                      const isEdit = !!editingId;
                      const pendingFile = pendingDocFiles[docType];
                      const uploadedName = uploadedDocNames[docType] || '';
                      const uploadedText = uploadedName;
                      return (
                        <div key={docType} className="company-details-doc-field">
                          <label className="company-details-doc-label">{label}</label>
                          <div className="company-details-doc-upload-input-wrap">
                            <div className="company-details-doc-upload-icon-inside" aria-hidden>
                              <Upload size={14} strokeWidth={2} />
                              <span>Import</span>
                            </div>
                            <input
                              type="file"
                              accept=".pdf,.doc,.docx,.jpg,.jpeg,.png"
                              onChange={(e) => {
                                const f = e.target.files?.[0];
                                if (isEdit && f) {
                                  handleFileUpload(docType, f);
                                  e.target.value = '';
                                } else if (!isEdit) {
                                  setPendingDocFiles((prev) => ({ ...prev, [docType]: f || null }));
                                }
                              }}
                              disabled={isUploading || viewOnly}
                              className="company-details-doc-file-input"
                            />
                          </div>
                          {!isEdit && pendingFile && (
                            <span className="company-details-doc-pending">Selected: {pendingFile.name}</span>
                          )}
                          {isEdit && uploadedText && (
                            <span className="company-details-doc-pending">Uploaded: {uploadedText}</span>
                          )}
                          {isEdit && !uploadedText && fileId && (
                            <span className="company-details-doc-pending">File attached{uploadingDoc === docType ? ' — loading name…' : ''}</span>
                          )}
                          <p className="company-details-doc-hint-line">Max 5MB · PDF, DOC, DOCX, JPG, PNG</p>
                          {isUploading && <span className="company-details-doc-uploading">Uploading…</span>}
                        </div>
                      );
                    })}
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
      ) : (
          <div className="company-details-shell-box">
            <input
              ref={importFileRef}
              type="file"
              accept=".csv,.xlsx,.xls,text/csv,application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
              style={{ display: 'none' }}
              onChange={handleImportCompanies}
              aria-hidden
            />
            {loading ? (
              <p className="company-details-loading">Loading...</p>
            ) : companies.length === 0 ? (
              <div className="company-details-empty-state">
                <p className="company-details-empty">No company details. Click Add to create one.</p>
                <div className="company-details-empty-actions">
                  <button
                    type="button"
                    className="company-details-btn company-details-btn-export"
                    onClick={() => importFileRef.current?.click()}
                    disabled={importing}
                    title="Import CSV"
                    aria-label="Import CSV"
                  >
                    <svg
                      width="16"
                      height="16"
                      viewBox="0 0 24 24"
                      fill="none"
                      stroke="currentColor"
                      strokeWidth="2"
                      strokeLinecap="round"
                      strokeLinejoin="round"
                      aria-hidden
                    >
                      <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
                      <polyline points="7 10 12 15 17 10" />
                      <line x1="12" y1="15" x2="12" y2="3" />
                    </svg>
                    {importing ? 'Importing…' : 'Import CSV'}
                  </button>
                  <button type="button" className="company-details-btn company-details-btn-add" onClick={openAdd}>
                    <span className="company-details-btn-add-icon">+</span> Add Company
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
                      placeholder="Search by Name, Email, Number…"
                      value={tableSearch}
                      onChange={(e) => setTableSearch(e.target.value)}
                      aria-label="Search companies"
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
                      <svg
                        width="16"
                        height="16"
                        viewBox="0 0 24 24"
                        fill="none"
                        stroke="currentColor"
                        strokeWidth="2"
                        strokeLinecap="round"
                        strokeLinejoin="round"
                        aria-hidden
                      >
                        <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
                        <polyline points="7 10 12 15 17 10" />
                        <line x1="12" y1="15" x2="12" y2="3" />
                      </svg>
                      {importing ? 'Importing…' : 'Import CSV'}
                    </button>
                    <button type="button" className="company-details-btn company-details-btn-export" onClick={exportCompaniesCsv} title="Export CSV" aria-label="Export CSV">
                      <svg
                        width="16"
                        height="16"
                        viewBox="0 0 24 24"
                        fill="none"
                        stroke="currentColor"
                        strokeWidth="2"
                        strokeLinecap="round"
                        strokeLinejoin="round"
                        aria-hidden
                      >
                        <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
                        <polyline points="17 8 12 3 7 8" />
                        <line x1="12" y1="3" x2="12" y2="15" />
                      </svg>
                      Export CSV
                    </button>
                    <button type="button" className="company-details-btn company-details-btn-add" onClick={openAdd}>
                      <span className="company-details-btn-add-icon">+</span> Add Company
                    </button>
                    {allPageSelected && (
                      <button
                        type="button"
                        className="company-details-btn company-details-btn-delete-selected"
                        onClick={handleDeleteSelected}
                        title="Delete selected"
                        aria-label="Delete selected companies"
                        disabled={bulkDeleting}
                      >
                        <Trash2 size={20} aria-hidden />
                      </button>
                    )}
                  </div>
                </div>
                <div className="company-details-table-wrap">
                  <table className="company-details-data-table">
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
                        <th scope="col">Name</th>
                        <th scope="col">Email</th>
                        <th scope="col">Phone</th>
                        <th scope="col">Director</th>
                        <th scope="col">PAN</th>
                        <th scope="col">GST</th>
                        <th scope="col">Incorp. date</th>
                        <th className="company-details-th-actions" scope="col">
                          Actions
                        </th>
                      </tr>
                    </thead>
                    <tbody>
                      {displayCompanies.length === 0 ? (
                        <tr>
                          <td colSpan={10} className="company-details-table-empty-cell">
                            No companies match your search.
                          </td>
                        </tr>
                      ) : (
                          pagedCompanies.map((company, idx) => (
                            <tr key={company.id} className="company-details-data-row">
                              <td>
                                <input
                                  type="checkbox"
                                  className="company-details-table-checkbox"
                                  checked={selectedIds.has(String(company.id))}
                                  onChange={() => toggleRowSelected(company.id)}
                                  aria-label={`Select ${company.companyName || 'company'}`}
                                />
                              </td>
                              <td className="company-details-td-num">
                                {(effectiveTablePage - 1) * COMPANY_TABLE_PAGE_SIZE + idx + 1}
                              </td>
                              <td className="company-details-td-strong">{company.companyName || ''}</td>
                              <td>{company.companyMail || ''}</td>
                              <td>{company.companyPhoneNumber || ''}</td>
                              <td>{company.directorName || ''}</td>
                              <td className="company-details-td-mono">{company.companyPANNumber || ''}</td>
                              <td className="company-details-td-mono">{company.gstNo || ''}</td>
                              <td>{formatDisplayDate(company.incorprationDate || company.incorporationDate)}</td>
                              <td>
                                <div className="company-details-table-actions">
                                  <button
                                    type="button"
                                    className="company-details-table-action company-details-table-action--view"
                                    onClick={() => openView(company)}
                                    title="View"
                                    aria-label={`View ${company.companyName || 'company'}`}
                                  >
                                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                                      <path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z" />
                                      <circle cx="12" cy="12" r="3" />
                                    </svg>
                                  </button>
                                  <button
                                    type="button"
                                    className="company-details-table-action company-details-table-action--edit"
                                    onClick={() => openEdit(company)}
                                    title="Edit"
                                    aria-label={`Edit ${company.companyName || 'company'}`}
                                  >
                                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                                      <path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7" />
                                      <path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z" />
                                    </svg>
                                  </button>
                                  <button
                                    type="button"
                                    className="company-details-table-action company-details-table-action--delete"
                                    onClick={() => handleDelete(company.id)}
                                    title="Delete"
                                    aria-label={`Delete ${company.companyName || 'company'}`}
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
                {displayCompanies.length > 0 ? (
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
                ) : null}
              </>
            )}
          </div>
      )}
    </div>
  );
};

export default CompanyDetails;
