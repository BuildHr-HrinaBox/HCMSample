import ExcelJS from 'exceljs';
import {
  flattenPayrollEarningColumns,
  readPayrollForm15WageAmounts,
  readPayrollScalar,
  readPayrollTextScalar,
} from '../../utils/payrollEarnings';
import { ensureExcelJSDataRowsWithBorders } from '../../utils/excelTableBorders';
import {
  applyFormCRJContractorFromSite,
  enrichEstablishmentPrincipalEmployerHeaderFields,
  writeStatutoryHeaderFieldsToExcelJsWorksheet,
} from '../../utils/statutorySiteCompanyHeaders';
import {
  employeeGidCandidates,
  payrollRowGidCandidates,
} from './formXIXMPWageSlip';

/** Rajasthan Form B — Register of Wages (FORMAT OF WAGE REGISTER). */

export const FORM_B_RJ_NIL = 'Nil';

/** True when a Sample Payroll / pay-run row carries PF or Voluntary Provident Fund. */
export function formBRJPayrollRowHasSampleDeductionFields(payrollRow) {
  if (!payrollRow || payrollRow.fetch_error) return false;
  const fields = resolveFormBRajasthanPayrollFields(payrollRow);
  return (
    (fields.pf !== '' && fields.pf != null) ||
    (fields.voluntaryProvidentFund !== '' && fields.voluntaryProvidentFund != null)
  );
}

/** Read YYYY-MM from a Sample Payroll / pay-run row. */
export function readFormBRJPayrollMonthIso(row) {
  if (!row || typeof row !== 'object') return '';
  const keys = [
    'payroll_month',
    'Payroll_Month',
    'payrollMonth',
    'PayrollMonth',
    'salary_month',
    'Salary_Month',
    'yearmonth',
    'year_month',
    'YearMonth',
    'monthFilter',
    'MonthFilter',
    'monthfilter',
  ];
  for (let i = 0; i < keys.length; i += 1) {
    const raw = String(row[keys[i]] ?? '').trim();
    if (!raw) continue;
    const iso = raw.match(/^(\d{4}-\d{2})/);
    if (iso) return iso[1];
  }
  const payDate = String(row.pay_date || row.payDate || row.payment_date || '').trim();
  const fromPay = payDate.match(/^(\d{4})-(\d{2})/);
  if (fromPay) return `${fromPay[1]}-${fromPay[2]}`;
  const fromPayDmy = payDate.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})$/);
  if (fromPayDmy) {
    return `${fromPayDmy[3]}-${String(fromPayDmy[2]).padStart(2, '0')}`;
  }
  return '';
}

/**
 * Keep only rows for the selected wage month.
 * - Rows tagged with another month are always dropped.
 * - Untagged rows are kept only when trustedPrimaryMonthFetch (fresh Sample Payroll for that month).
 */
export function filterFormBRJPayrollRowsForMonth(records, monthIso, helpers = {}) {
  const rows = (Array.isArray(records) ? records : []).filter(
    (row) => row && typeof row === 'object' && row.fetch_error !== true
  );
  const want = String(monthIso || '').trim();
  if (!/^\d{4}-\d{2}$/.test(want)) return rows;

  const trusted = helpers.trustedPrimaryMonthFetch === true;
  const matched = [];
  let sawOtherMonth = false;
  const untagged = [];

  rows.forEach((row) => {
    const tagged = readFormBRJPayrollMonthIso(row);
    if (!tagged) {
      untagged.push(row);
      return;
    }
    if (tagged === want) {
      matched.push(row);
      return;
    }
    sawOtherMonth = true;
  });

  if (matched.length > 0) return matched;
  // Never treat another month's batch as the selected month (even if untagged mixed in).
  if (sawOtherMonth) return [];
  if (trusted && untagged.length > 0) {
    return untagged.map((row) => ({
      ...row,
      payroll_month: want,
      payrollMonth: want,
    }));
  }
  return [];
}

/**
 * Form B_RJ Sample Payroll for the selected wage month only.
 * Never uses "latest month" fallback or multi-month bulk name matching.
 */
export function resolveFormBRajasthanPayrollRowsForAutofill(
  statutoryPayrollRows,
  monthCandidates = [],
  helpers = {}
) {
  const {
    cachedSampleRows = null,
    bulkSampleRows = null,
    payDate = '',
    scopedPayrollMonth = '',
    trustedPrimaryMonthFetch = false,
  } = helpers;

  const months = (Array.isArray(monthCandidates) ? monthCandidates : [])
    .map((m) => String(m || '').trim())
    .filter((m) => /^\d{4}-\d{2}$/.test(m));
  const primaryMonth = String(scopedPayrollMonth || months[0] || '').trim();

  const stampPayDate = (rows) => {
    const date = String(payDate || '').trim();
    if (!date) return rows;
    return rows.map((row) => {
      if (!row || typeof row !== 'object') return row;
      if (String(row.pay_date || row.payDate || '').trim()) return row;
      return {
        ...row,
        pay_date: date,
        payDate: date,
        payment_date: date,
        date_of_payment: date,
      };
    });
  };

  const normalizeRows = (rows) =>
    (Array.isArray(rows) ? rows : [])
      .filter((row) => row && typeof row === 'object' && row.fetch_error !== true)
      .map((row) => flattenPayrollEarningColumns(row));

  const preferDeductionRows = (rows) => {
    const list = normalizeRows(rows);
    if (list.length === 0) return [];
    const withDeductions = list.filter((row) => formBRJPayrollRowHasSampleDeductionFields(row));
    return stampPayDate(withDeductions.length > 0 ? withDeductions : list);
  };

  const takeMonthRows = (rows, trusted) => {
    const scoped = filterFormBRJPayrollRowsForMonth(rows, primaryMonth, {
      trustedPrimaryMonthFetch: trusted === true,
    });
    const preferred = preferDeductionRows(scoped);
    return preferred.length > 0 ? preferred : null;
  };

  // 1) Trusted fresh May (etc.) Sample Payroll batch from caller.
  if (trustedPrimaryMonthFetch) {
    const trustedHit =
      takeMonthRows(cachedSampleRows, true) ||
      takeMonthRows(statutoryPayrollRows, true);
    if (trustedHit) return trustedHit;
  }

  // 2) Month-keyed cache — keep only rows tagged for primary month (drop other-month pollution).
  const fromCached = takeMonthRows(cachedSampleRows, false);
  if (fromCached) return fromCached;

  // 3) Bulk / statutory — require explicit payroll_month === primary (no untagged latest-month reuse).
  const fromBulk = takeMonthRows(bulkSampleRows, false);
  if (fromBulk) return fromBulk;

  const fromStatutory = takeMonthRows(statutoryPayrollRows, false);
  if (fromStatutory) return fromStatutory;

  // Do not call resolveFormXIXMPPayrollRowsForAutofill — it falls back to "latest" month.
  return [];
}

export function formBRajasthanHeaderNorm(txt) {
  return String(txt || '')
    .replace(/\r?\n/g, ' ')
    .replace(/[''`´]/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

const stripLeadingNumber = (s) =>
  String(s || '')
    .replace(/^\d+[\.\)]\s*/, '')
    .trim();

function unwrapEmployeeRecord(emp) {
  return emp?.Employee || emp?.employee || emp;
}

function pickEmployeeValue(emp, keys) {
  const src = unwrapEmployeeRecord(emp);
  if (!src || typeof src !== 'object') return '';
  for (let i = 0; i < keys.length; i += 1) {
    const raw = src[keys[i]];
    if (raw == null || raw === '') continue;
    if (typeof raw === 'object') {
      const nested = String(raw.displayValue ?? raw.name ?? raw.Name ?? '').trim();
      if (nested) return nested;
      continue;
    }
    const text = String(raw).trim();
    if (text) return text;
  }
  return '';
}

/** People full name: FirstName + MiddleName + LastName (Form B_RJ payroll match key). */
export function readEmployeeFullName(emp) {
  const parts = readFormBRJPersonNameParts(emp);
  if (parts.fullName) return parts.fullName;
  return pickEmployeeValue(unwrapEmployeeRecord(emp), [
    'DisplayName',
    'Display Name',
    'Employee_Name',
    'Employee Name',
    'full_name',
    'Full Name',
  ]);
}

const normFormBRJPersonName = (value) =>
  String(value || '')
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();

/** Read FirstName / MiddleName / LastName from People or a payroll row. */
export function readFormBRJPersonNameParts(personOrRow) {
  const src = unwrapEmployeeRecord(personOrRow) || {};
  const firstName = pickEmployeeValue(src, [
    'FirstName',
    'First Name',
    'firstName',
    'first_name',
  ]);
  const middleName = pickEmployeeValue(src, [
    'MiddleName',
    'Middle Name',
    'middleName',
    'middle_name',
    'Middle_Name',
  ]);
  const lastName = pickEmployeeValue(src, [
    'LastName',
    'Last Name',
    'lastName',
    'last_name',
  ]);
  let fullName = [firstName, middleName, lastName].filter(Boolean).join(' ').trim();
  if (!fullName) {
    fullName = pickEmployeeValue(src, [
      'employee_name',
      'Employee Name',
      'EmployeeName',
      'full_name',
      'Full Name',
      'DisplayName',
      'Display Name',
      'name',
      'Name',
    ]);
  }
  // Sample Payroll often stores only employee_name — recover edge tokens.
  if ((!firstName || !lastName) && fullName) {
    const tokens = normFormBRJPersonName(fullName).split(' ').filter(Boolean);
    return {
      firstName: firstName || (tokens[0] ? String(fullName).trim().split(/\s+/).filter(Boolean)[0] : ''),
      middleName:
        middleName ||
        (tokens.length >= 3
          ? String(fullName)
              .trim()
              .split(/\s+/)
              .filter(Boolean)
              .slice(1, -1)
              .join(' ')
          : ''),
      lastName:
        lastName ||
        (tokens.length >= 2
          ? String(fullName).trim().split(/\s+/).filter(Boolean).slice(-1)[0]
          : ''),
      fullName: String(fullName || '').trim(),
    };
  }
  return {
    firstName,
    middleName,
    lastName,
    fullName: String(fullName || '').trim(),
  };
}

function formBRJPayrollDisplayNames(payrollRow) {
  if (!payrollRow || typeof payrollRow !== 'object') return [];
  const parts = readFormBRJPersonNameParts(payrollRow);
  return [
    parts.fullName,
    [parts.firstName, parts.middleName, parts.lastName].filter(Boolean).join(' '),
    [parts.firstName, parts.lastName].filter(Boolean).join(' '),
    payrollRow.employee_name,
    payrollRow.EmployeeName,
    payrollRow.full_name,
    payrollRow.name,
  ]
    .map((v) => String(v || '').replace(/\r?\n/g, ' ').replace(/\s+/g, ' ').trim())
    .filter(Boolean);
}

/**
 * Form B_RJ payroll match: FirstName + MiddleName + LastName.
 * Requires first AND last; when MiddleName is present it must appear in payroll name.
 * Never matches on first-name-only.
 */
export function formBRJNamesMatch(employeeOrRow, payrollRow) {
  if (!payrollRow || typeof payrollRow !== 'object' || payrollRow.fetch_error) return false;
  const left = readFormBRJPersonNameParts(employeeOrRow);
  const payNames = formBRJPayrollDisplayNames(payrollRow);
  if (payNames.length === 0) return false;

  const leftFull = normFormBRJPersonName(left.fullName);
  if (leftFull) {
    for (let i = 0; i < payNames.length; i += 1) {
      if (normFormBRJPersonName(payNames[i]) === leftFull) return true;
    }
  }

  const peopleTokens = [
    ...normFormBRJPersonName(left.firstName).split(' '),
    ...normFormBRJPersonName(left.middleName).split(' '),
    ...normFormBRJPersonName(left.lastName).split(' '),
  ].filter(Boolean);
  if (peopleTokens.length < 2) return false;
  const leftFirst = peopleTokens[0];
  const leftLast = peopleTokens[peopleTokens.length - 1];
  const middleTokens = peopleTokens.slice(1, -1);

  for (let i = 0; i < payNames.length; i += 1) {
    const tokens = normFormBRJPersonName(payNames[i]).split(' ').filter(Boolean);
    if (tokens.length < 2) continue;
    const payFirst = tokens[0];
    const payLast = tokens[tokens.length - 1];
    if (payFirst !== leftFirst || payLast !== leftLast) continue;
    // When People has middle token(s), require they appear in the payroll name.
    if (middleTokens.length > 0 && !middleTokens.every((tok) => tokens.includes(tok))) {
      continue;
    }
    return true;
  }
  return false;
}

const normFormBRJCompareToken = (value) =>
  String(value || '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

function normalizeFormBRJIdentityCode(value) {
  let s = String(value || '')
    .trim()
    .toLowerCase();
  if (!s) return '';
  // Align VE0705 / ve0705 / 0705 / 705 style codes when comparing People ↔ payroll.
  s = s.replace(/^0+/, '') || '0';
  const stripped = s.replace(/^(ve|emp|e|gid)[\s._-]*/i, '').replace(/^0+/, '');
  if (stripped) return stripped;
  return s;
}

function readFormBRJPersonIdentityCodes(personOrRow) {
  const src = unwrapEmployeeRecord(personOrRow) || {};
  const roleId =
    src.Role && typeof src.Role === 'object' ? src.Role.ID || src.Role.id || '' : '';
  const codes = [
    ...employeeGidCandidates(src),
    pickEmployeeValue(src, [
      'EmployeeID',
      'Employee ID',
      'Employee_ID',
      'EmployeeId',
      'employeeId',
      'employee_id',
      'employee_number',
      'Employee_Number',
      'Employee Number',
      'EmployeeCode',
      'Employee Code',
      'EmpCode',
      'Zoho_ID',
      'ZohoID',
      'zoho_id',
      'zohoId',
      'Role.ID',
      'erecno',
      'Erecno',
      '__employeeLookupId',
    ]),
    roleId,
    personOrRow?.__employeeLookupId,
    src.__employeeLookupId,
  ]
    .map((v) => normalizeFormBRJIdentityCode(v))
    .filter(Boolean);
  return [...new Set(codes)];
}

function readFormBRJPayrollIdentityCodes(payrollRow) {
  if (!payrollRow || typeof payrollRow !== 'object') return [];
  const codes = [
    ...payrollRowGidCandidates(payrollRow),
    payrollRow.employee_id,
    payrollRow.employee_number,
    payrollRow.employee_code,
    payrollRow.EmployeeID,
    payrollRow['Employee ID'],
    payrollRow.Zoho_ID,
    payrollRow.zoho_id,
  ]
    .map((v) => normalizeFormBRJIdentityCode(v))
    .filter(Boolean);
  return [...new Set(codes)];
}

/** People EmployeeID / Emp Code only — never Zoho_ID (avoids VE1428 row taking VE948 May wages). */
function readFormBRJPreferredPeopleEmployeeCodes(personOrRow) {
  const src = unwrapEmployeeRecord(personOrRow) || {};
  const codes = [
    pickEmployeeValue(src, [
      'EmployeeID',
      'Employee ID',
      'Employee_ID',
      'EmployeeId',
      'employeeId',
      'employee_id',
      'employee_number',
      'Employee_Number',
      'Employee Number',
      'EmployeeCode',
      'Employee Code',
      'EmpCode',
      '__employeeLookupId',
    ]),
    personOrRow?.__employeeLookupId,
    src.__employeeLookupId,
  ]
    .map((v) => normalizeFormBRJIdentityCode(v))
    .filter(Boolean);
  return [...new Set(codes)];
}

function formBRJPreferredCodesAgree(personOrRow, payrollRow) {
  const left = readFormBRJPreferredPeopleEmployeeCodes(personOrRow);
  const right = readFormBRJPayrollIdentityCodes(payrollRow);
  if (left.length === 0 || right.length === 0) return false;
  const rightSet = new Set(right);
  return left.some((code) => rightSet.has(code));
}

function readFormBRJPersonLocation(personOrRow) {
  const src = unwrapEmployeeRecord(personOrRow) || {};
  return pickEmployeeValue(src, [
    'Work_location',
    'Work Location',
    'work_location',
    'workLocation',
    'Location',
    'location',
    'LocationName',
    'Location Name',
    'Site',
    'site',
    'SiteName',
    'Site Name',
    'siteName',
  ]);
}

function readFormBRJPayrollLocation(payrollRow) {
  if (!payrollRow || typeof payrollRow !== 'object') return '';
  return String(
    payrollRow.work_location ||
      payrollRow.Work_location ||
      payrollRow.location ||
      payrollRow.Location ||
      payrollRow.site ||
      payrollRow.Site ||
      payrollRow.site_name ||
      payrollRow.SiteName ||
      payrollRow.department ||
      payrollRow.Department ||
      ''
  ).trim();
}

function formBRJLocationsAgree(personOrRow, payrollRow) {
  const left = normFormBRJCompareToken(readFormBRJPersonLocation(personOrRow));
  const right = normFormBRJCompareToken(readFormBRJPayrollLocation(payrollRow));
  if (!left || !right) return false;
  if (left === right) return true;
  const short = left.length <= right.length ? left : right;
  const long = left.length <= right.length ? right : left;
  if (short.length < 4) return false;
  return long.includes(short) || short.split(' ').some((tok) => tok.length >= 4 && long.includes(tok));
}

function formBRJIdentityCodesAgree(personOrRow, payrollRow) {
  const left = readFormBRJPersonIdentityCodes(personOrRow);
  const right = readFormBRJPayrollIdentityCodes(payrollRow);
  if (left.length === 0 || right.length === 0) return false;
  const rightSet = new Set(right);
  return left.some((code) => rightSet.has(code));
}

/** True when both sides have IDs and none overlap (e.g. two "Karthik P" with different EmployeeIDs). */
function formBRJIdentityCodesConflict(personOrRow, payrollRow) {
  const left = readFormBRJPersonIdentityCodes(personOrRow);
  const right = readFormBRJPayrollIdentityCodes(payrollRow);
  if (left.length === 0 || right.length === 0) return false;
  const rightSet = new Set(right);
  return !left.some((code) => rightSet.has(code));
}

/**
 * Same First(+Middle)+Last name on more than one payroll row:
 * - Prefer Employee ID / GID / Zoho ID (even for a single name hit)
 * - Else unique Work Location / Site
 * Never attach another person's wages when IDs disagree.
 */
export function pickFormBRajasthanPayrollRowAmongNameMatches(employeeOrRow, nameMatches) {
  const hits = (Array.isArray(nameMatches) ? nameMatches : []).filter(
    (row) => row && typeof row === 'object' && row.fetch_error !== true
  );
  if (hits.length === 0) return null;

  const byId = hits.filter((row) => formBRJIdentityCodesAgree(employeeOrRow, row));
  if (byId.length === 1) return byId[0];
  if (byId.length > 1) {
    const byIdAndLoc = byId.filter((row) => formBRJLocationsAgree(employeeOrRow, row));
    if (byIdAndLoc.length === 1) return byIdAndLoc[0];
    return byId[0];
  }

  // People has an ID and every name-hit has a different ID → wrong person(s); do not use.
  const conflicting = hits.filter((row) => formBRJIdentityCodesConflict(employeeOrRow, row));
  if (conflicting.length === hits.length) return null;

  const nonConflict = hits.filter((row) => !formBRJIdentityCodesConflict(employeeOrRow, row));
  if (nonConflict.length === 1) return nonConflict[0];

  const byLoc = nonConflict.filter((row) => formBRJLocationsAgree(employeeOrRow, row));
  if (byLoc.length === 1) return byLoc[0];
  // Same name + same location without unique ID → do not guess.
  return null;
}

/** Match Sample Payroll by People EmployeeID (VE…) only — never by name / wrong twin. */
export function findFormBRajasthanPayrollRowByName(employeeOrRow, payrollRows) {
  const rows = (Array.isArray(payrollRows) ? payrollRows : []).filter(
    (row) => row && typeof row === 'object' && row.fetch_error !== true
  );
  if (rows.length === 0) return null;

  // VE1428 on Form A/People must not take VE948 May wages just because both are "Karthik P".
  const preferred = readFormBRJPreferredPeopleEmployeeCodes(employeeOrRow);
  if (preferred.length > 0) {
    const preferredHits = rows.filter((row) => formBRJPreferredCodesAgree(employeeOrRow, row));
    if (preferredHits.length === 0) return null;
    if (preferredHits.length === 1) return preferredHits[0];
    const named = preferredHits.filter((row) => formBRJNamesMatch(employeeOrRow, row));
    if (named.length === 1) return named[0];
    const byLoc = preferredHits.filter((row) => formBRJLocationsAgree(employeeOrRow, row));
    if (byLoc.length === 1) return byLoc[0];
    return preferredHits[0];
  }

  // No EmployeeID on People — Zoho/GID only (still never name-only).
  const idHits = rows.filter((row) => formBRJIdentityCodesAgree(employeeOrRow, row));
  if (idHits.length === 0) return null;
  if (idHits.length === 1) return idHits[0];

  const namedId = idHits.filter((row) => formBRJNamesMatch(employeeOrRow, row));
  if (namedId.length === 1) return namedId[0];

  const byLoc = idHits.filter((row) => formBRJLocationsAgree(employeeOrRow, row));
  if (byLoc.length === 1) return byLoc[0];

  return idHits[0];
}

export function resolveFormBRajasthanPayrollRowForEmployee(emp, payrollRows) {
  // No XIX / name-only fallback — wrong twin wages must never attach.
  return findFormBRajasthanPayrollRowByName(emp, payrollRows);
}

/** One payroll row per employee — name-first, no GID gate. */
export function buildFormBRajasthanPayrollRowResolver(payrollRows) {
  const rows = Array.isArray(payrollRows) ? payrollRows : [];
  const used = new Set();
  const rowKey = (row) => {
    if (!row || typeof row !== 'object') return '';
    const name = normFormBRJPersonName(formBRJPayrollDisplayNames(row)[0] || '');
    const id = String(row.employee_id || row.employee_number || row.gidNumber || '').trim();
    return id ? `id:${id}|name:${name}` : name ? `name:${name}` : '';
  };
  return (emp) => {
    const available = rows.filter((row) => {
      const key = rowKey(row);
      return !key || !used.has(key);
    });
    const hit = resolveFormBRajasthanPayrollRowForEmployee(
      emp,
      available.length > 0 ? available : rows
    );
    const key = rowKey(hit);
    if (hit && key) used.add(key);
    return hit;
  };
}

function readBankAccountNumber(emp) {
  return pickEmployeeValue(emp, [
    'Bank_Account_Number',
    'Bank Account Number',
    'Account_Number',
    'Account Number',
    'AccountNumber',
    'accountNumber',
  ]);
}

function readEmployeeId(emp) {
  return pickEmployeeValue(emp, [
    'EmployeeID',
    'Employee ID',
    'Employee_ID',
    'EmployeeId',
    'employeeId',
    'Employee.ID',
    'erecno',
    'Erecno',
    'Zoho_ID',
  ]);
}

function parseFormBRJHeaderLabel(h) {
  const raw = stripLeadingNumber(h);
  const m = String(raw || '').match(/^(.+)_([\s\S]+)$/);
  if (m) return String(m[2] || '').trim();
  return raw;
}

const normHeader = (h) =>
  formBRajasthanHeaderNorm(parseFormBRJHeaderLabel(h))
    .replace(/[^a-z0-9 ]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

export function isFormBRJSerialHeader(h) {
  const s = normHeader(h);
  return /^sr\.?\s*no|^s\.?\s*no|serial|sl\.?\s*no|^no\.?$/.test(s);
}

export function isFormBRJNameHeader(h) {
  const s = normHeader(h);
  if (!/\bname\b/.test(s)) return false;
  if (/father|spouse|husband|employer|establishment|bank/.test(s)) return false;
  return true;
}

export function isFormBRJEmployeeCodeHeader(h) {
  const s = normHeader(h);
  return (
    (s.includes('employee') && s.includes('code')) ||
    s === 'employee code' ||
    s === 'emp code'
  );
}

export function isFormBRJDaysWorkedHeader(h) {
  const s = normHeader(h);
  return (s.includes('day') || s.includes('days')) && s.includes('work');
}

export function isFormBRJRateOfWageHeader(h) {
  const s = normHeader(h);
  return s.includes('rate') && s.includes('wage');
}

export function isFormBRJOvertimeHoursHeader(h) {
  const s = normHeader(h);
  return /\bover[\s-]*time\b/.test(s) && (s.includes('hour') || s.includes('hrs') || s.includes('worked'));
}

export function isFormBRJPaymentsOvertimeHeader(h) {
  const s = normHeader(h);
  if (!/\bover[\s-]*time\b/.test(s)) return false;
  if (s.includes('hour') || s.includes('hrs') || s.includes('worked')) return false;
  return s.includes('payment') || s.includes('pay');
}

export function isFormBRJBasicHeader(h) {
  const s = normHeader(h);
  return (s === 'basic' || /^basic\b/.test(s)) && !s.includes('special');
}

export function isFormBRJHraHeader(h) {
  const s = normHeader(h);
  return s === 'hra' || /\bhra\b/.test(s) || (s.includes('house') && s.includes('rent'));
}

export function isFormBRJVoluntaryPfHeader(h) {
  const s = normHeader(h);
  const compact = s.replace(/\s+/g, '');
  return (
    compact === 'vpf' ||
    (s.includes('voluntary') && (s.includes('provident') || s.includes('pf') || compact.includes('vpf'))) ||
    s.includes('voluntary provident fund')
  );
}

export function isFormBRJPfHeader(h) {
  const s = normHeader(h);
  if (isFormBRJVoluntaryPfHeader(h)) return false;
  const compact = s.replace(/\s+/g, '');
  return (
    compact === 'pf' ||
    compact === 'epf' ||
    s.includes('provident fund') ||
    s.includes('epf') ||
    /^p\.?\s*f\.?$/.test(s)
  );
}

export function isFormBRJSpecialBasicHeader(h) {
  const s = normHeader(h);
  return s.includes('special') && s.includes('basic');
}

export function isFormBRJDaHeader(h) {
  const s = normHeader(h);
  return s === 'da' || /\bda\b/.test(s) || s.includes('dearness');
}

function isFormBRJPlainTotalHeader(h) {
  const s = normHeader(h);
  if (!s) return false;
  if (s.includes('deduction') || s.includes('deducation')) return false;
  return s === 'total' || /^total\b/.test(s);
}

export function isFormBRJOthersHeader(h) {
  const s = normHeader(h);
  return s === 'others' || /^others\b/.test(s);
}

/** Deduction total / deducation — gross_pay − net_pay (not gross_pay). */
export function isFormBRJDeductionTotalHeader(header, headers = [], headerIndex = null) {
  const s = normHeader(header);
  if (!s) return false;
  if (
    (s.includes('deduction') || s.includes('deducation')) &&
    (s.includes('total') || s === 'deduction' || s === 'deducation')
  ) {
    return true;
  }
  const list = Array.isArray(headers) ? headers : [];
  if (!isFormBRJPlainTotalHeader(header)) return false;
  const idx =
    Number.isInteger(headerIndex) && headerIndex >= 0 ? headerIndex : list.indexOf(header);
  if (idx < 0) return false;
  const totalIndices = list
    .map((h, i) => (isFormBRJPlainTotalHeader(h) ? i : -1))
    .filter((i) => i >= 0);
  // Last plain "Total" is Deduction → Total (duplicate labels break indexOf).
  return totalIndices.length >= 2 && idx === totalIndices[totalIndices.length - 1];
}

/** Earnings total (column after HRA) — payroll gross_pay. */
export function isFormBRJEarningsTotalHeader(header, headers = [], headerIndex = null) {
  if (isFormBRJDeductionTotalHeader(header, headers, headerIndex)) return false;
  return isFormBRJPlainTotalHeader(header);
}

export function isFormBRJNetPaymentHeader(h) {
  const s = normHeader(h);
  return s.includes('net') && (s.includes('payment') || s.includes('payable') || s.includes('paid'));
}

export function isFormBRJBankReceiptHeader(h) {
  const s = normHeader(h);
  return (
    (s.includes('receipt') && (s.includes('employee') || s.includes('workman') || s.includes('worker'))) ||
    (s.includes('bank') && (s.includes('transaction') || s.includes('account')))
  );
}

export function isFormBRJPaymentDateHeader(h) {
  const s = normHeader(h);
  return s.includes('date') && s.includes('payment');
}

/** Payroll / wage columns — skip generic People field mapping. Name, code, and bank receipt come from People. */
export function isFormBRJSkipPeopleAutofillHeader(h) {
  return (
    isFormBRJEmployeeCodeHeader(h) ||
    isFormBRJRateOfWageHeader(h) ||
    isFormBRJDaysWorkedHeader(h) ||
    isFormBRJOvertimeHoursHeader(h) ||
    isFormBRJPaymentsOvertimeHeader(h) ||
    isFormBRJBasicHeader(h) ||
    isFormBRJSpecialBasicHeader(h) ||
    isFormBRJDaHeader(h) ||
    isFormBRJHraHeader(h) ||
    isFormBRJOthersHeader(h) ||
    isFormBRJPfHeader(h) ||
    isFormBRJVoluntaryPfHeader(h) ||
    isFormBRJEarningsTotalHeader(h) ||
    isFormBRJDeductionTotalHeader(h) ||
    isFormBRJNetPaymentHeader(h) ||
    isFormBRJPaymentDateHeader(h)
  );
}

export function formBRJHeaderAliasBucket(norm) {
  const n = String(norm || '').trim();
  if (!n) return '';
  if (/^sr\.?\s*no|^s\.?\s*no|serial|sl\.?\s*no|^no\.?$/.test(n)) return 'sno';
  if (n === 'name' || /^name$/.test(n)) return 'name';
  if (n.includes('rate') && n.includes('wage')) return 'rateOfWage';
  if ((n.includes('day') || n.includes('days')) && n.includes('work')) return 'daysWorked';
  if (/\bover[\s-]*time\b/.test(n) && (n.includes('hour') || n.includes('hrs') || n.includes('worked'))) {
    return 'overtimeHours';
  }
  if (/\bover[\s-]*time\b/.test(n) && (n.includes('payment') || n.includes('pay'))) return 'paymentsOvertime';
  if (n.includes('special') && n.includes('basic')) return 'specialBasic';
  if (n === 'basic' || /^basic\b/.test(n)) return 'basic';
  if (n === 'da' || /\bda\b/.test(n) || n.includes('dearness')) return 'da';
  if (n === 'hra' || /\bhra\b/.test(n) || (n.includes('house') && n.includes('rent'))) return 'hra';
  if (
    n === 'vpf' ||
    (n.includes('voluntary') && (n.includes('provident') || n.includes('pf'))) ||
    n.includes('voluntary provident fund')
  ) {
    return 'voluntaryPf';
  }
  if (
    n.replace(/\s+/g, '') === 'pf' ||
    n.replace(/\s+/g, '') === 'epf' ||
    n.includes('provident fund') ||
    n.includes('epf') ||
    /^p\.?\s*f\.?$/.test(n)
  ) {
    return 'pf';
  }
  if (n === 'others' || /^others\b/.test(n)) return 'others';
  if (
    ((n.includes('deduction') || n.includes('deducation')) && n.includes('total')) ||
    n === 'deduction' ||
    n === 'deducation'
  ) {
    return 'deductionTotal';
  }
  if (n.includes('net') && (n.includes('payment') || n.includes('payable') || n.includes('paid'))) {
    return 'netPay';
  }
  if (
    (n.includes('receipt') && (n.includes('employee') || n.includes('workman') || n.includes('worker'))) ||
    (n.includes('bank') && (n.includes('transaction') || n.includes('account')))
  ) {
    return 'bankReceipt';
  }
  if (n.includes('date') && n.includes('payment')) return 'paymentDate';
  if (n === 'total' || /^total\b/.test(n)) return 'earningsTotal';
  if (n.includes('employee') && n.includes('code')) return 'employeeCode';
  return n;
}

export function headersIndicateFormBRajasthanTable(tableHeaders) {
  if (!Array.isArray(tableHeaders) || tableHeaders.length < 5) return false;
  const joined = tableHeaders.map((h) => formBRajasthanHeaderNorm(h)).join('\n');
  const hasName = /\bname\b/.test(joined);
  const hasDaysWorked = /day/.test(joined) && /work/.test(joined);
  const hasBasic = /\bbasic\b/.test(joined);
  const hasWageRegister =
    /\bhra\b/.test(joined) ||
    /\bpf\b/.test(joined) ||
    /voluntary\s+provident/.test(joined) ||
    /net\s+payment/.test(joined) ||
    /payment\s+overtime/.test(joined) ||
    /deduction|deducation/.test(joined);
  const looksLikeFormAEmployeeRegister =
    /education/.test(joined) &&
    (/type/.test(joined) && /employ/.test(joined)) &&
    (/national/.test(joined) || /aadha?ar/.test(joined));
  return hasName && hasDaysWorked && hasBasic && hasWageRegister && !looksLikeFormAEmployeeRegister;
}

export function isFormBRajasthanContext(
  formHeader,
  rowItem,
  fileName,
  sheetText = '',
  tableHeaders = []
) {
  const parts = [
    rowItem?.formName,
    rowItem?.FormName,
    rowItem?.formFileName,
    rowItem?.FormFileName,
    rowItem?.siteState,
    rowItem?.SiteState,
    rowItem?.state,
    rowItem?.State,
    fileName,
    formHeader?.title,
    formHeader?.subtitle,
    formHeader?.reference,
    sheetText,
    Array.isArray(tableHeaders) ? tableHeaders.join(' ') : '',
  ]
    .filter((x) => x != null && String(x).trim() !== '')
    .join(' ')
    .toLowerCase();

  if (/gujarat|\b_gj\b|form[\s._-]*b[\s._-]*gj|form_b_gj/.test(parts)) return false;

  // Form_B_MH / Form_B_RJ filenames: underscore after B breaks \bform…b\b, so match codes first.
  if (/form[\s._-]*b[\s._-]*(rj|mh)\b|form_b_(rj|mh)\b/.test(parts)) return true;

  const hasRajasthan =
    /rajasthan|\b_rj\b|form[\s._-]*b[\s._-]*rj|form_b_rj/.test(parts);
  const hasMaharashtra =
    /maharashtra|\b_mh\b|form[\s._-]*b[\s._-]*mh|form_b_mh/.test(parts);
  // form_b_mh: "b" is followed by "_" (word char) so trailing \b fails — use lookahead.
  const hasFormB =
    /form[\s._-]*b(?=[\s._-]|$)/.test(parts) || /\bform[\s._-]*b\b/.test(parts);

  if ((hasRajasthan || hasMaharashtra) && hasFormB) return true;
  if (headersIndicateFormBRajasthanTable(tableHeaders) && (hasRajasthan || hasMaharashtra) && hasFormB) {
    return true;
  }
  if (headersIndicateFormBRajasthanTable(tableHeaders) && (hasRajasthan || hasMaharashtra)) return true;
  return false;
}

/** Same contractor header field Form C_RJ uses, so Site Management name+address can bind. */
export function enrichFormBRajasthanDisplayHeader(formHeader) {
  const base = formHeader && typeof formHeader === 'object' ? { ...formHeader } : {};
  let fields = enrichEstablishmentPrincipalEmployerHeaderFields(base.fields || []);
  const hasContractor = fields.some((field) =>
    /name\s+and\s+address\s+of\s+(?:the\s+)?contractor/i.test(String(field?.label || ''))
  );
  if (!hasContractor) {
    fields = [
      ...fields,
      { label: 'Name and address of contractor', value: '', key: 'form_b_rj_contractor' },
    ];
  }
  return { ...base, fields };
}

export function applyFormBRajasthanContractorFromSite(headerData, site, formHeaderFields = []) {
  return applyFormCRJContractorFromSite(headerData, site, formHeaderFields);
}

export function resolveFormBRajasthanTableHeaders(tableHeaders) {
  const parsed = Array.isArray(tableHeaders)
    ? tableHeaders.map((h) => String(h ?? '')).filter((h) => String(h).trim() !== '')
    : [];
  if (parsed.length === 0) return [];

  // Keep duplicate captions addressable as separate object keys.
  // Example: "Total" + "Total" → "Total" + "Total " so earnings Total (gross_pay)
  // is not overwritten by deduction Total (gross_pay − net_pay).
  const seen = new Map();
  return parsed.map((header) => {
    const count = seen.get(header) || 0;
    seen.set(header, count + 1);
    return count > 0 ? `${header}${' '.repeat(count)}` : header;
  });
}

function parsePayrollNumber(value) {
  const n = Number(String(value ?? '').replace(/[,₹]/g, '').trim());
  return Number.isFinite(n) ? n : NaN;
}

function sumPayrollNumbers(values) {
  let sum = 0;
  let any = false;
  values.forEach((value) => {
    const n = parsePayrollNumber(value);
    if (Number.isFinite(n)) {
      sum += n;
      any = true;
    }
  });
  return any ? Math.round(sum * 100) / 100 : '';
}

function formatBRJPayrollPayDate(payDateRaw) {
  const raw = String(payDateRaw || '').trim();
  if (!raw) return '';
  const iso = raw.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) return `${iso[3]}-${iso[2]}-${iso[1]}`;
  const dmy = raw.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})$/);
  if (dmy) {
    const dd = String(dmy[1]).padStart(2, '0');
    const mm = String(dmy[2]).padStart(2, '0');
    return `${dd}-${mm}-${dmy[3]}`;
  }
  return raw;
}

function resolveFormBRJRateOfWageFromGross(grossPay) {
  const gross = parsePayrollNumber(grossPay);
  if (!Number.isFinite(gross) || gross <= 0) return '';
  return Math.round((gross / 26) * 100) / 100;
}

export function resolveFormBRajasthanPayrollFields(payrollRow, helpers = {}) {
  const monthEndDate = String(helpers.monthEndDate || '').trim();
  const empty = {
    paidDays: '',
    rateOfWage: '',
    overtimeHours: FORM_B_RJ_NIL,
    paymentsOvertime: FORM_B_RJ_NIL,
    basic: '',
    specialBasic: '',
    da: '',
    hra: '',
    earningsTotal: '',
    pf: '',
    voluntaryProvidentFund: '',
    deductionsTotal: '',
    netPay: '',
    paymentDate: monthEndDate,
  };
  if (!payrollRow || payrollRow.fetch_error) return empty;

  const flat = flattenPayrollEarningColumns(payrollRow);
  const source = { ...flat, ...payrollRow };
  const wageAmounts = readPayrollForm15WageAmounts(payrollRow);

  const paidDays = readPayrollScalar(
    source,
    ['paid_days', 'Paid_days', 'Paid Days', 'days_worked', 'Days Worked', 'paidDays', 'no_of_days_worked'],
    [/^paid_days$/, /paiddays/, /daysworked/]
  );

  // Absent / no attendance for the wage month → do not carry CTC leftovers (HRA, VPF, etc.).
  const paidDaysNum = parsePayrollNumber(paidDays);
  const hasNoDaysWorked =
    paidDays === '' ||
    paidDays == null ||
    (Number.isFinite(paidDaysNum) && paidDaysNum <= 0);
  if (hasNoDaysWorked) {
    return {
      ...empty,
      paidDays:
        paidDays === '' || paidDays == null
          ? ''
          : Number.isFinite(paidDaysNum)
            ? paidDaysNum
            : paidDays,
      paymentDate: '',
    };
  }

  const basic =
    readPayrollScalar(
      source,
      ['basic', 'Basic', 'earned_basic', 'Earned Basic', 'basic_pay', 'Basic Earnings'],
      [/^basic$/, /^earned_basic$/]
    ) || wageAmounts.basic;

  const specialBasic = readPayrollScalar(
    source,
    ['special_basic', 'Special Basic', 'specialBasic', 'SpecialBasic'],
    [/^special_basic$/, /special.*basic/]
  );

  const da = readPayrollScalar(
    source,
    ['da', 'DA', 'dearness_allowance', 'Dearness Allowance', 'dearnessAllowance'],
    [/^da$/, /dearness/]
  );

  const hra = readPayrollScalar(
    source,
    ['hra', 'HRA', 'hra_fbp', 'HRA FBP', 'house_rent_allowance', 'House Rent Allowance'],
    [/^hra$/, /house.*rent/]
  );

  const grossPay = readPayrollScalar(
    source,
    ['gross_pay', 'Gross Pay', 'Gross_pay', 'grossPay', 'total_earnings'],
    [/^gross_pay$/, /^total_earnings$/]
  );

  const netPay = readPayrollScalar(
    source,
    ['net_pay', 'Net Pay', 'netPay', 'monthly_salary'],
    [/^net_pay$/]
  );

  const pickAmount = (...values) => {
    for (let i = 0; i < values.length; i += 1) {
      const value = values[i];
      if (value === '' || value == null) continue;
      const n = parsePayrollNumber(value);
      if (Number.isFinite(n)) return n;
    }
    return '';
  };

  const pf = pickAmount(
    flat.epf_contribution,
    flat.pf,
    flat.PF,
    flat.provident_fund,
    readPayrollScalar(
      source,
      [
        'pf',
        'PF',
        'epf_contribution',
        'EPF Contribution',
        'epf',
        'EPF',
        'employee_pf',
        'Employee PF',
        'provident_fund',
        'Provident Fund',
      ],
      [/^pf$/, /^epf(_contribution)?$/, /^provident_fund$/]
    )
  );

  const voluntaryProvidentFund = pickAmount(
    flat.voluntary_provident_fund,
    flat.vpf,
    flat.VoluntaryProvidentFund,
    flat.voluntaryProvidentFund,
    readPayrollScalar(
      source,
      [
        'voluntary_provident_fund',
        'VoluntaryProvidentFund',
        'Voluntary Provident Fund',
        'voluntaryProvidentFund',
        'vpf',
        'VPF',
      ],
      [/^vpf$/, /voluntary.*provident/i]
    )
  );

  const grossN = parsePayrollNumber(grossPay);
  const netN = parsePayrollNumber(netPay);
  let deductionsTotal = '';
  if (Number.isFinite(grossN) && Number.isFinite(netN)) {
    deductionsTotal = Math.round((grossN - netN) * 100) / 100;
    if (deductionsTotal < 0) deductionsTotal = '';
  }

  // Total after HRA ← gross_pay (not Basic + HRA).
  const earningsTotal = grossPay || sumPayrollNumbers([basic, hra]);
  const rateOfWage = resolveFormBRJRateOfWageFromGross(grossPay);

  // Date of Payment → month end (selected wage month), not payroll pay_date.
  const paymentDate =
    monthEndDate ||
    formatBRJPayrollPayDate(
      readPayrollTextScalar(
        source,
        ['pay_date', 'Pay Date', 'payment_date', 'Payment Date', 'paid_date', 'date_of_payment', 'Date of Payment'],
        [/^pay_date$/, /payment.*date/i, /^paid_date$/]
      ) || String(payrollRow?.pay_date ?? flat?.pay_date ?? helpers.payDate ?? '').trim()
    );

  return {
    paidDays,
    rateOfWage,
    overtimeHours: FORM_B_RJ_NIL,
    paymentsOvertime: FORM_B_RJ_NIL,
    basic,
    specialBasic,
    da,
    hra,
    earningsTotal,
    pf,
    voluntaryProvidentFund,
    deductionsTotal,
    netPay,
    paymentDate,
  };
}

function formatCellValue(value, { allowZero = false, allowNil = false } = {}) {
  if (value == null || value === '') return '';
  if (allowNil && String(value).trim().toLowerCase() === 'nil') return FORM_B_RJ_NIL;
  const num = parsePayrollNumber(value);
  if (Number.isFinite(num) && (allowZero || num !== 0)) return num;
  return String(value).trim();
}

export function enrichFormBRajasthanPayrollRows(mappedData, employees, headers, helpers = {}) {
  const hdrs = resolveFormBRajasthanTableHeaders(headers);
  const {
    resolvePayrollRow = null,
    payrollRows = null,
    sanitizeValue = (v) => String(v ?? '').trim(),
    formatStatutoryDateDisplay = (v) => String(v || '').trim(),
    payDate = '',
    monthEndDate = '',
    overwrite = true,
    rowIndexOffset = 0,
  } = helpers;
  if (!Array.isArray(mappedData) || mappedData.length === 0) {
    return 0;
  }
  const defaultResolver =
    typeof resolvePayrollRow === 'function'
      ? resolvePayrollRow
      : Array.isArray(payrollRows) && payrollRows.length > 0
        ? buildFormBRajasthanPayrollRowResolver(payrollRows)
        : null;
  if (typeof defaultResolver !== 'function') return 0;

  let hits = 0;
  mappedData.forEach((row, rowIndex) => {
    const empItem = employees[rowIndex];
    const emp = unwrapEmployeeRecord(empItem);
    const lookupId = String(
      row?.__employeeLookupId ||
        pickEmployeeValue(emp || {}, [
          'EmployeeID',
          'Employee ID',
          'Employee_ID',
          'EmployeeId',
          'Zoho_ID',
          'ZohoID',
        ]) ||
        ''
    ).trim();
    // Prefer People record; always carry lookup Employee ID so wage bind cannot fall back to name.
    const matchSrc =
      emp && typeof emp === 'object'
        ? {
            ...emp,
            ...(lookupId
              ? {
                  EmployeeID: emp.EmployeeID || emp['Employee ID'] || lookupId,
                  __employeeLookupId: emp.__employeeLookupId || lookupId,
                }
              : {}),
            ...(row?.__employeeLookupName
              ? { __employeeLookupName: row.__employeeLookupName }
              : {}),
          }
        : row && typeof row === 'object'
          ? row
          : null;
    let payrollRow = defaultResolver(matchSrc, row, rowIndex);
    if ((!payrollRow || payrollRow.fetch_error) && Array.isArray(payrollRows) && payrollRows.length > 0) {
      payrollRow = findFormBRajasthanPayrollRowByName(matchSrc, payrollRows);
    }
    const hadPayroll = Boolean(payrollRow && !payrollRow.fetch_error);
    const merged = applyFormBRajasthanEmployeeToRow(row, emp, hdrs, {
      sanitizeValue,
      formatStatutoryDateDisplay,
      payDate,
      monthEndDate,
      payrollRow: hadPayroll ? payrollRow : null,
      rowIndex: rowIndexOffset + rowIndex,
      overwrite,
    });
    Object.assign(row, merged);
    if (hadPayroll) hits += 1;
  });
  const cleaned = sanitizeFormBRajasthanMappedWageRows(mappedData, hdrs);
  cleaned.forEach((row, idx) => {
    if (mappedData[idx] && row) Object.assign(mappedData[idx], row);
  });
  return hits;
}

export function applyFormBRajasthanEmployeeToRow(row, emp, headers, helpers = {}) {
  const hdrs = resolveFormBRajasthanTableHeaders(headers);
  const out = row && typeof row === 'object' ? { ...row } : {};
  const {
    sanitizeValue = (v) => String(v ?? '').trim(),
    payrollRow = null,
    rowIndex = 0,
    formatStatutoryDateDisplay = (v) => String(v || '').trim(),
    payDate = '',
    monthEndDate = '',
    overwrite = true,
  } = helpers;

  const cellIsEmpty = (header) => {
    const v = String(out[header] || '').trim();
    if (!v) return true;
    const s = v.toLowerCase();
    return /^enter\b/.test(s) || s.includes('enter ');
  };
  const setCell = (header, value, opts = {}) => {
    if (!header) return;
    // Always clear stale template / People values when payroll field is empty.
    if (value == null || value === '') {
      if (opts.clearWhenEmpty && (overwrite || cellIsEmpty(header))) out[header] = '';
      return;
    }
    if (!overwrite && !cellIsEmpty(header)) return;
    out[header] = sanitizeValue(formatCellValue(value, opts));
  };

  // Final gate: if People has EmployeeID (VE1428), refuse payroll that does not share it
  // (blocks VE948 May wages on the Aug Karthik P row even if a caller matched by name).
  const matchEmp = emp && typeof emp === 'object' ? emp : out;
  const preferredIds = readFormBRJPreferredPeopleEmployeeCodes(matchEmp);
  let safePayrollRow = payrollRow && !payrollRow.fetch_error ? payrollRow : null;
  if (safePayrollRow && preferredIds.length > 0 && !formBRJPreferredCodesAgree(matchEmp, safePayrollRow)) {
    safePayrollRow = null;
  }

  const payroll = resolveFormBRajasthanPayrollFields(
    safePayrollRow,
    { formatStatutoryDateDisplay, payDate, monthEndDate }
  );
  // Name ← People FirstName + MiddleName + LastName (payroll match key).
  const fullName = readEmployeeFullName(emp);
  if (fullName) out.__employeeLookupName = fullName;
  const lookupId = readEmployeeId(emp);
  if (lookupId) out.__employeeLookupId = lookupId;

  hdrs.forEach((header, headerIndex) => {
    if (isFormBRJSerialHeader(header)) {
      setCell(header, String(rowIndex + 1), { allowZero: true });
      return;
    }
    if (isFormBRJNameHeader(header)) {
      setCell(header, fullName, { clearWhenEmpty: true });
      return;
    }
    if (isFormBRJEmployeeCodeHeader(header)) {
      setCell(header, readEmployeeId(emp), { clearWhenEmpty: true });
      return;
    }
    if (isFormBRJRateOfWageHeader(header)) {
      setCell(header, payroll.rateOfWage, { clearWhenEmpty: true });
      return;
    }
    if (isFormBRJDaysWorkedHeader(header)) {
      setCell(header, payroll.paidDays, { allowZero: true, clearWhenEmpty: true });
      return;
    }
    if (isFormBRJOvertimeHoursHeader(header)) {
      setCell(header, payroll.overtimeHours, { allowNil: true, clearWhenEmpty: true });
      return;
    }
    if (isFormBRJPaymentsOvertimeHeader(header)) {
      setCell(header, payroll.paymentsOvertime, { allowNil: true, clearWhenEmpty: true });
      return;
    }
    if (isFormBRJSpecialBasicHeader(header)) {
      setCell(header, payroll.specialBasic, { allowZero: true, clearWhenEmpty: true });
      return;
    }
    if (isFormBRJBasicHeader(header)) {
      setCell(header, payroll.basic, { clearWhenEmpty: true });
      return;
    }
    if (isFormBRJDaHeader(header)) {
      setCell(header, payroll.da, { allowZero: true, clearWhenEmpty: true });
      return;
    }
    if (isFormBRJHraHeader(header)) {
      setCell(header, payroll.hra, { clearWhenEmpty: true });
      return;
    }
    if (isFormBRJVoluntaryPfHeader(header)) {
      setCell(header, payroll.voluntaryProvidentFund, { allowZero: true, clearWhenEmpty: true });
      return;
    }
    if (isFormBRJPfHeader(header)) {
      setCell(header, payroll.pf, { allowZero: true, clearWhenEmpty: true });
      return;
    }
    if (isFormBRJOthersHeader(header)) {
      // Deduction → Others stays blank (do not fetch payroll residual).
      if (overwrite || cellIsEmpty(header)) out[header] = '';
      return;
    }
    if (isFormBRJDeductionTotalHeader(header, hdrs, headerIndex)) {
      setCell(header, payroll.deductionsTotal, { allowZero: true, clearWhenEmpty: true });
      return;
    }
    if (isFormBRJEarningsTotalHeader(header, hdrs, headerIndex)) {
      setCell(header, payroll.earningsTotal, { clearWhenEmpty: true });
      return;
    }
    if (isFormBRJNetPaymentHeader(header)) {
      setCell(header, payroll.netPay, { clearWhenEmpty: true });
      return;
    }
    if (isFormBRJBankReceiptHeader(header)) {
      setCell(header, readBankAccountNumber(emp), { clearWhenEmpty: true });
      return;
    }
    if (isFormBRJPaymentDateHeader(header)) {
      setCell(header, payroll.paymentDate, { clearWhenEmpty: true });
    }
  });

  return out;
}

export function getFormBRajasthanRowValueForHeader(row, header, rowIndex = 0) {
  if (!row || !header) return '';
  const direct = row[header];
  if (direct != null && String(direct).trim() !== '') return String(direct).trim();
  const bucket = formBRJHeaderAliasBucket(formBRajasthanHeaderNorm(header));
  for (const [k, v] of Object.entries(row)) {
    if (
      formBRJHeaderAliasBucket(formBRajasthanHeaderNorm(k)) === bucket &&
      v != null &&
      String(v).trim() !== ''
    ) {
      return String(v).trim();
    }
  }
  if (bucket === 'sno') return String(rowIndex + 1);
  return '';
}

export function rowHasMeaningfulFormBRajasthanExportData(row, headers) {
  const hdrs = resolveFormBRajasthanTableHeaders(headers);
  return hdrs.some((header) => {
    const bucket = formBRJHeaderAliasBucket(formBRajasthanHeaderNorm(header));
    if (bucket === 'sno') return false;
    return getFormBRajasthanRowValueForHeader(row, header) !== '';
  });
}

export function remapFormBRajasthanRowsToHeaders(rows, sourceHeaders, targetHeaders) {
  const src = resolveFormBRajasthanTableHeaders(sourceHeaders);
  const tgt = resolveFormBRajasthanTableHeaders(targetHeaders);
  if (!Array.isArray(rows)) return [];
  return rows.map((row, rowIndex) => {
    const out = {};
    tgt.forEach((targetHeader, colIdx) => {
      let val = '';
      if (row && typeof row === 'object') {
        if (Object.prototype.hasOwnProperty.call(row, targetHeader)) {
          val = row[targetHeader];
        } else if (src[colIdx] && Object.prototype.hasOwnProperty.call(row, src[colIdx])) {
          val = row[src[colIdx]];
        } else {
          val = getFormBRajasthanRowValueForHeader(row, targetHeader, rowIndex);
        }
      }
      if (
        (val == null || String(val).trim() === '') &&
        formBRJHeaderAliasBucket(formBRajasthanHeaderNorm(targetHeader)) === 'sno'
      ) {
        val = String(rowIndex + 1);
      }
      out[targetHeader] = val == null ? '' : val;
    });
    return out;
  });
}

/** True when No. Of Days worked is blank or ≤ 0 (absent for the wage month). */
export function formBRJRowHasNoDaysWorked(row, headers) {
  if (!row || typeof row !== 'object') return true;
  const hdrs = resolveFormBRajasthanTableHeaders(headers);
  let raw = '';
  for (let i = 0; i < hdrs.length; i += 1) {
    if (!isFormBRJDaysWorkedHeader(hdrs[i])) continue;
    raw = String(getFormBRajasthanRowValueForHeader(row, hdrs[i]) || '').trim();
    if (raw) break;
  }
  if (!raw) {
    for (const [k, v] of Object.entries(row)) {
      if (formBRJHeaderAliasBucket(formBRajasthanHeaderNorm(k)) !== 'daysWorked') continue;
      raw = String(v ?? '').trim();
      if (raw) break;
    }
  }
  if (!raw || /^enter\b/i.test(raw)) return true;
  const n = parsePayrollNumber(raw);
  return Number.isFinite(n) && n <= 0;
}

/**
 * Absent employees must not keep CTC leftovers (HRA, VPF, etc.) on download/export
 * when the modal/cache was filled before payroll absent rules ran.
 */
export function sanitizeFormBRajasthanAbsentWageRow(row, headers) {
  if (!row || typeof row !== 'object') return row;
  if (!formBRJRowHasNoDaysWorked(row, headers)) return row;
  const hdrs = resolveFormBRajasthanTableHeaders(headers);
  const out = { ...row };
  const clearBucket = (bucket) =>
    bucket === 'rateOfWage' ||
    bucket === 'basic' ||
    bucket === 'specialBasic' ||
    bucket === 'da' ||
    bucket === 'hra' ||
    bucket === 'earningsTotal' ||
    bucket === 'pf' ||
    bucket === 'voluntaryPf' ||
    bucket === 'others' ||
    bucket === 'deductionTotal' ||
    bucket === 'netPayment' ||
    bucket === 'paymentDate';

  hdrs.forEach((header) => {
    const bucket = formBRJHeaderAliasBucket(formBRajasthanHeaderNorm(header));
    if (bucket === 'overtimeHours' || bucket === 'paymentsOvertime') {
      out[header] = FORM_B_RJ_NIL;
      return;
    }
    if (clearBucket(bucket)) out[header] = '';
  });
  Object.keys(out).forEach((key) => {
    if (String(key).startsWith('__')) return;
    const bucket = formBRJHeaderAliasBucket(formBRajasthanHeaderNorm(key));
    if (bucket === 'overtimeHours' || bucket === 'paymentsOvertime') {
      out[key] = FORM_B_RJ_NIL;
      return;
    }
    if (clearBucket(bucket)) out[key] = '';
  });
  return out;
}

/**
 * Clear HRA / VPF / other wage leftovers when Rate/Days/Basic were not filled from
 * Sample Payroll (People fuzzy match or StatutoryData overlay leak).
 */
export function sanitizeFormBRajasthanOrphanAllowanceRow(row, headers) {
  if (!row || typeof row !== 'object') return row;
  const hdrs = resolveFormBRajasthanTableHeaders(headers);
  if (hdrs.length === 0) return row;
  const readBucket = (bucket) => {
    for (let i = 0; i < hdrs.length; i += 1) {
      const h = hdrs[i];
      if (formBRJHeaderAliasBucket(formBRajasthanHeaderNorm(h)) !== bucket) continue;
      const v = String(row[h] ?? '').trim();
      if (v && !/^enter\b/i.test(v)) return v;
    }
    return '';
  };
  const hasCorePayroll =
    Boolean(readBucket('daysWorked')) ||
    Boolean(readBucket('rateOfWage')) ||
    Boolean(readBucket('basic')) ||
    Boolean(readBucket('earningsTotal')) ||
    Boolean(readBucket('netPayment'));
  if (hasCorePayroll) return row;
  return sanitizeFormBRajasthanAbsentWageRow(row, hdrs);
}

export function sanitizeFormBRajasthanMappedWageRows(rows, headers) {
  const hdrs = resolveFormBRajasthanTableHeaders(headers);
  return (Array.isArray(rows) ? rows : []).map((row) =>
    sanitizeFormBRajasthanOrphanAllowanceRow(row, hdrs)
  );
}

export function filterFormBRajasthanExportRows(rows, headers) {
  const hdrs = resolveFormBRajasthanTableHeaders(headers);
  return (Array.isArray(rows) ? rows : [])
    .map((row) => sanitizeFormBRajasthanAbsentWageRow(row, hdrs))
    .filter((row) => rowHasMeaningfulFormBRajasthanExportData(row, hdrs));
}

const excelCellValueToString = (val) => {
  if (val == null) return '';
  if (typeof val === 'string' || typeof val === 'number' || typeof val === 'boolean') {
    return String(val);
  }
  if (typeof val === 'object') {
    if (val instanceof Date) return val.toISOString().slice(0, 10);
    if (Array.isArray(val.richText)) return val.richText.map((rt) => rt?.text || '').join('');
    if (val.text != null) return String(val.text);
    if (val.result != null) return String(val.result);
  }
  return '';
};

function excelCellLooksLikeSerialHeader(text) {
  const t = formBRajasthanHeaderNorm(text).replace(/[^a-z0-9 ]/g, ' ').replace(/\s+/g, ' ').trim();
  return /^sr\.?\s*no|^s\.?\s*no|serial|sl\.?\s*no/.test(t);
}

function buildMergeTopLeftResolver(worksheet) {
  const cache = new Map();
  return (r, c) => {
    const key = `${r}:${c}`;
    if (cache.has(key)) return cache.get(key);
    let topLeft = { r, c };
    const merges = worksheet.model?.merges;
    if (Array.isArray(merges)) {
      for (let mi = 0; mi < merges.length; mi += 1) {
        const parts = String(merges[mi] || '').split(':');
        if (parts.length !== 2) continue;
        const tl = worksheet.getCell(parts[0]);
        const br = worksheet.getCell(parts[1]);
        if (!tl || !br) continue;
        if (r >= tl.row && r <= br.row && c >= tl.col && c <= br.col) {
          topLeft = { r: tl.row, c: tl.col };
          break;
        }
      }
    }
    cache.set(key, topLeft);
    return topLeft;
  };
}

function detectFormBRajasthanTableLayout(worksheet, hints = {}) {
  const getMergeTopLeft = buildMergeTopLeftResolver(worksheet);
  const getMergedAwareCellText = (r, c) => {
    const tl = getMergeTopLeft(r, c);
    return excelCellValueToString(worksheet.getCell(tl.r, tl.c)?.value).trim();
  };

  let headerRow =
    hints.parsedHeaderRowIndex != null && hints.parsedHeaderRowIndex >= 0
      ? hints.parsedHeaderRowIndex + 1
      : -1;
  let startCol =
    hints.parsedTableStartCol != null && hints.parsedTableStartCol >= 0
      ? hints.parsedTableStartCol + 1
      : 1;

  if (headerRow < 1) {
    const maxScanRows = Math.max(35, worksheet.rowCount + 5);
    for (let r = 1; r <= maxScanRows; r += 1) {
      for (let c = 1; c <= 30; c += 1) {
        if (excelCellLooksLikeSerialHeader(getMergedAwareCellText(r, c))) {
          headerRow = r;
          startCol = c;
          break;
        }
      }
      if (headerRow > 0) break;
    }
  }
  if (headerRow < 1) return null;

  const templateCols = [];
  for (let c = startCol; c <= startCol + 40; c += 1) {
    const label = getMergedAwareCellText(headerRow, c);
    if (!label) {
      if (templateCols.length >= 8) break;
      continue;
    }
    const bucket = formBRJHeaderAliasBucket(formBRajasthanHeaderNorm(label));
    templateCols.push({ col: c, label, bucket });
    if (templateCols.length >= 30) break;
  }
  if (templateCols.length < 4) return null;

  // Disambiguate duplicate plain "Total" columns: last → Deduction Total.
  const plainTotalIdxs = templateCols
    .map((entry, i) => (entry.bucket === 'earningsTotal' ? i : -1))
    .filter((i) => i >= 0);
  if (plainTotalIdxs.length >= 2) {
    const lastIdx = plainTotalIdxs[plainTotalIdxs.length - 1];
    templateCols[lastIdx] = { ...templateCols[lastIdx], bucket: 'deductionTotal' };
  }

  const dataStartRow =
    hints.parsedDataStartIndex != null && hints.parsedDataStartIndex >= 0
      ? hints.parsedDataStartIndex + 1
      : headerRow + 1;

  return { headerRow, dataStartRow, templateCols, startCol };
}

export async function buildFormBRajasthanWorkbookWithTemplateStyles({
  templateArrayBuffer,
  mappedData,
  headersToUse,
  parsedFormHeader,
  formFileName,
  headerFormData,
  parsedHeaderRowIndex,
  parsedDataStartIndex,
  parsedTableStartCol,
  formatStatutoryDateDisplay = null,
}) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(templateArrayBuffer);
  const worksheet = workbook.worksheets[0];
  if (!worksheet) throw new Error('Template worksheet not found.');

  const layout = detectFormBRajasthanTableLayout(worksheet, {
    parsedHeaderRowIndex,
    parsedDataStartIndex,
    parsedTableStartCol,
  });
  if (!layout) throw new Error('Could not locate Rajasthan Form B table header row.');

  const { dataStartRow, templateCols } = layout;

  if (headerFormData && typeof headerFormData === 'object') {
    writeStatutoryHeaderFieldsToExcelJsWorksheet(worksheet, {
      headerFormData,
      parsedFormHeader,
      headerRowEnd: Math.max(1, layout.headerRow - 1),
      maxScanCols: 80,
    });
  }

  const templateHeaderLabels = templateCols.map((entry) => entry.label);
  const normalizedHeaders = resolveFormBRajasthanTableHeaders(
    Array.isArray(headersToUse) && headersToUse.length >= templateHeaderLabels.length
      ? headersToUse
      : templateHeaderLabels.length > 0
        ? templateHeaderLabels
        : headersToUse
  );
  const rows = filterFormBRajasthanExportRows(
    remapFormBRajasthanRowsToHeaders(
      Array.isArray(mappedData) ? mappedData : [],
      normalizedHeaders,
      normalizedHeaders
    ),
    normalizedHeaders
  );

  const tableColMin = Math.min(...templateCols.map(({ col }) => col));
  const tableColMax = Math.max(...templateCols.map(({ col }) => col));
  const clearToRow = Math.max(dataStartRow + rows.length + 10, dataStartRow + 15);
  for (let r = dataStartRow; r <= clearToRow; r += 1) {
    for (let c = tableColMin; c <= tableColMax; c += 1) {
      worksheet.getCell(r, c).value = '';
    }
  }

  rows.forEach((row, idx) => {
    const targetRow = worksheet.getRow(dataStartRow + idx);
    if (targetRow) targetRow.height = 18;
    templateCols.forEach(({ col, bucket, label }, colIdx) => {
      const uniqueLabel = normalizedHeaders[colIdx] || label;
      let val = '';
      if (Object.prototype.hasOwnProperty.call(row, uniqueLabel)) {
        val = row[uniqueLabel];
      } else if (Object.prototype.hasOwnProperty.call(row, label)) {
        val = row[label];
      } else {
        val = getFormBRajasthanRowValueForHeader(row, uniqueLabel, idx);
      }
      if ((val == null || String(val).trim() === '') && bucket === 'deductionTotal') {
        // Prefer the last plain-Total key (uniquified trailing spaces), not earnings Total.
        const totalKeys = Object.keys(row || {}).filter((k) =>
          isFormBRJPlainTotalHeader(k)
        );
        if (totalKeys.length >= 2) {
          const lastKey = totalKeys[totalKeys.length - 1];
          if (row[lastKey] != null && String(row[lastKey]).trim() !== '') {
            val = row[lastKey];
          }
        }
      } else if ((val == null || String(val).trim() === '') && bucket) {
        for (const [k, v] of Object.entries(row)) {
          if (
            formBRJHeaderAliasBucket(formBRajasthanHeaderNorm(k)) === bucket &&
            v != null &&
            String(v).trim() !== ''
          ) {
            val = String(v).trim();
            break;
          }
        }
      }
      if ((val == null || String(val).trim() === '') && bucket === 'sno') val = String(idx + 1);
      const cell = worksheet.getCell(dataStartRow + idx, col);
      if (val == null || String(val).trim() === '') {
        cell.value = '';
        return;
      }
      if (bucket === 'sno') {
        const n = Number(String(val).replace(/[,]/g, '').trim());
        cell.value = Number.isFinite(n) ? n : String(val);
      } else if (bucket === 'bankReceipt') {
        const digits = String(val).replace(/\D/g, '');
        cell.value = digits || String(val);
        if (digits) cell.numFmt = '0';
      } else if (
        bucket === 'overtimeHours' ||
        bucket === 'paymentsOvertime'
      ) {
        cell.value = String(val);
      } else {
        const n = Number(String(val).replace(/[,₹]/g, '').trim());
        cell.value = Number.isFinite(n) ? n : String(val);
      }
      cell.alignment = { ...(cell.alignment || {}), vertical: 'middle', wrapText: true };
    });
  });

  if (rows.length > 0) {
    ensureExcelJSDataRowsWithBorders(worksheet, {
      dataStartRow,
      dataRowCount: rows.length,
      colFrom: tableColMin,
      colTo: tableColMax,
      templateRow: dataStartRow,
      templateBodyRows: 1,
    });
  }

  const out = await workbook.xlsx.writeBuffer();
  const fileName =
    formFileName ||
    parsedFormHeader?.title?.replace(/[^a-zA-Z0-9]/g, '_') ||
    'Form_B_RJ_Rajasthan.xlsx';
  return {
    blob: new Blob([out], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    }),
    fileName,
  };
}
