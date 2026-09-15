/**
 * Form XVI Andhra Pradesh — Muster Roll.
 * Attendance matching: fullname OR firstname + lastname (never first-name-only / loose).
 */

import { buildSiteContractorNameAndAddress } from '../../utils/statutorySiteCompanyHeaders';
import {
  formXVIIITamilNaduFirstAndLastNamesMatch,
  formXVIIITamilNaduFirstLastKey,
  readFormXVIIITamilNaduFirstLastParts,
} from './formXVIIITamilNaduWagesMuster';

export const FORM_XVI_AP_CONTRACTOR_HEADER_KEY = 'form_xvi_contractor';
export const FORM_XVI_AP_DEFAULT_CONTRACTOR_NAME = 'VAYONA ENERGY PRIVATE LIMITED';
export const FORM_XVI_AP_ESTABLISHMENT_CONTRACT_KEY = 'form_xvi_establishment_contract_carried';
export const FORM_XVI_AP_ESTABLISHMENT_CONTRACT_LABEL =
  'Name and Address of the Establishment in/ under which contract is carried on';

export function formXVIAPNormName(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ');
}

/** Same first+last / fullname key used for Form XVIII attendance. */
export function formXVIAPFirstLastKey(src, extraParts = null) {
  return formXVIIITamilNaduFirstLastKey(src, extraParts);
}

/**
 * True when employee and attendance/Form 25 record share the same fullname
 * or the same firstname + lastname (never first-name-only / loose substring).
 */
export function formXVIAPNamesMatch(employeeOrRow, record, extraParts = null) {
  const left = readFormXVIIITamilNaduFirstLastParts(employeeOrRow, extraParts);
  const right = readFormXVIIITamilNaduFirstLastParts(record);
  const leftFull = formXVIAPNormName(left.fullName);
  const rightFull = formXVIAPNormName(right.fullName);
  // Exact fullname (covers multi-word and single-token names like "Nagaraju").
  if (leftFull && rightFull && leftFull === rightFull) return true;
  return formXVIIITamilNaduFirstAndLastNamesMatch(employeeOrRow, record, extraParts);
}

/** Lookup keys for indexing attendance by fullname and/or first+last. */
export function formXVIAPAttendanceNameKeys(src, extraParts = null) {
  const keys = new Set();
  const flKey = formXVIAPFirstLastKey(src, extraParts);
  if (flKey) keys.add(flKey);
  const { fullName, firstName, lastName } = readFormXVIIITamilNaduFirstLastParts(src, extraParts);
  const full = formXVIAPNormName(fullName);
  // Index exact fullname even for single-token names (e.g. Nagaraju).
  if (full) keys.add(full);
  if (firstName && lastName) keys.add(formXVIAPNormName(`${firstName} ${lastName}`));
  return Array.from(keys).filter(Boolean);
}

export function findFormXVIAPRecordByFullOrFirstLastName(employeeOrRow, records, extraParts = null) {
  const list = Array.isArray(records) ? records : [];
  return (
    list.find((row) => formXVIAPNamesMatch(employeeOrRow, row, extraParts)) || null
  );
}

function titleCaseFormXVIAPName(value) {
  const s = String(value || '')
    .trim()
    .replace(/\s+/g, ' ');
  if (!s) return '';
  return s
    .split(' ')
    .map((word) => (word ? word.charAt(0).toUpperCase() + word.slice(1).toLowerCase() : word))
    .join(' ');
}

/** Name of the Employee ← fullname, else firstname + lastname. */
export function resolveFormXVIAPEmployeeName(emp = {}) {
  const { firstName, lastName, fullName } = readFormXVIIITamilNaduFirstLastParts(emp);
  const combo = `${firstName} ${lastName}`.trim();
  const raw = String(fullName || combo || '').trim();
  return titleCaseFormXVIAPName(raw);
}

export function isFormXVIAPEmployeeNameHeader(header) {
  const n = String(header || '')
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .trim();
  if (!n) return false;
  if (/father|husband|contractor|employer|establishment/.test(n)) return false;
  return (
    /name\s+of\s+the\s+employee/.test(n) ||
    /name\s+of\s+the\s+workm[ae]n/.test(n) ||
    (/name/.test(n) && /employee|workman|workmen|worker/.test(n))
  );
}

export function formXVIAPCompactKey(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '');
}

/** Find Site Management row by site name or Location (e.g. AP-Nimbagallu). */
export function findFormXVIAPSiteRecord(sites, hints = {}) {
  const list = Array.isArray(sites) ? sites : [];
  if (!list.length) return null;
  const hintValues = [hints.siteName, hints.locationHint, hints.natureText]
    .concat(Array.isArray(hints.extraNames) ? hints.extraNames : [])
    .flatMap((v) => String(v || '').split(/[,;/|]+/))
    .map((v) => v.trim())
    .filter(Boolean);

  const siteKey = (s) => formXVIAPCompactKey(s?.siteName ?? s?.SiteName ?? s?.name ?? s?.Name);
  const locKey = (s) => formXVIAPCompactKey(s?.location ?? s?.Location);

  for (let i = 0; i < hintValues.length; i += 1) {
    const n = formXVIAPCompactKey(hintValues[i]);
    if (!n) continue;
    const exact = list.find((s) => siteKey(s) === n || locKey(s) === n);
    if (exact) return exact;
  }
  for (let i = 0; i < hintValues.length; i += 1) {
    const n = formXVIAPCompactKey(hintValues[i]);
    if (n.length < 3) continue;
    const fuzzy = list.find((s) => {
      const sn = siteKey(s);
      const loc = locKey(s);
      return (
        (sn && (sn.includes(n) || n.includes(sn))) ||
        (loc && (loc.includes(n) || n.includes(loc)))
      );
    });
    if (fuzzy) return fuzzy;
  }
  return null;
}

export function resolveContractorTextFromRecords(records) {
  const list = Array.isArray(records) ? records : [];
  const counts = new Map();
  for (let i = 0; i < list.length; i += 1) {
    const rec = list[i];
    if (!rec || typeof rec !== 'object') continue;
    const name = String(
      rec.contractorName ??
        rec.ContractorName ??
        rec.Contractor_Name ??
        rec.contractor_name ??
        (typeof rec.contractor === 'string' ? rec.contractor : '') ??
        (typeof rec.Contractor === 'string' ? rec.Contractor : '') ??
        ''
    ).trim();
    if (!name) continue;
    const squeezed = formXVIAPCompactKey(name);
    if (!squeezed || squeezed.startsWith('nameandaddressofcontractor')) continue;
    const key = name.toLowerCase();
    const prev = counts.get(key);
    counts.set(key, { name, n: (prev?.n || 0) + 1 });
  }
  let best = '';
  let bestN = 0;
  counts.forEach((v) => {
    if (v.n > bestN) {
      best = v.name;
      bestN = v.n;
    }
  });
  return best;
}

/** Site Management contractor first; then employee / CLRA; else company default. */
export function resolveFormXVIAPContractorFromSources({
  sites,
  siteName,
  locationHint,
  natureText,
  employees
} = {}) {
  const site = findFormXVIAPSiteRecord(sites, { siteName, locationHint, natureText });
  const fromSite = buildSiteContractorNameAndAddress(site);
  if (fromSite) return fromSite;
  const fromEmployees = resolveContractorTextFromRecords(employees);
  if (fromEmployees) return fromEmployees;
  return FORM_XVI_AP_DEFAULT_CONTRACTOR_NAME;
}
