/**
 * Form 10 Tamil Nadu — Overtime Muster payroll mapping.
 * Match Sample Payroll / pay-run rows by FirstName AND LastName only.
 * Never fall back to first-name-only (avoids Selva P picking up another Selva's July pay).
 */

const normForm10Name = (value) =>
  String(value || '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ');

const pickForm10NameField = (src, keys) => {
  if (!src || typeof src !== 'object') return '';
  for (let i = 0; i < keys.length; i += 1) {
    const raw = src[keys[i]];
    if (raw == null) continue;
    const text = String(raw).trim();
    if (text) return text;
  }
  return '';
};

const FORM10_FIRST_KEYS = [
  'FirstName',
  'firstName',
  'first_name',
  'firstname',
  'First Name',
  'First_Name',
];

const FORM10_LAST_KEYS = [
  'LastName',
  'lastName',
  'last_name',
  'lastname',
  'Last Name',
  'Last_Name',
  'Surname',
  'surname',
];

const FORM10_FULL_KEYS = [
  'employee_name',
  'EmployeeName',
  'Employee Name',
  'employeeName',
  'full_name',
  'Full_Name',
  'name',
  'Name',
];

export function readForm10PersonNameParts(src = null) {
  if (!src || typeof src !== 'object') {
    return { firstName: '', lastName: '', fullName: '' };
  }
  const payload =
    src.payroll_payload && typeof src.payroll_payload === 'object' ? src.payroll_payload : {};
  const employee =
    src.employee && typeof src.employee === 'object' && !Array.isArray(src.employee)
      ? src.employee
      : {};
  const pools = [src, payload, employee];
  let firstName = '';
  let lastName = '';
  let fullName = '';
  pools.forEach((pool) => {
    if (!firstName) firstName = pickForm10NameField(pool, FORM10_FIRST_KEYS);
    if (!lastName) lastName = pickForm10NameField(pool, FORM10_LAST_KEYS);
    if (!fullName) fullName = pickForm10NameField(pool, FORM10_FULL_KEYS);
  });
  const combo = `${firstName} ${lastName}`.trim();
  return {
    firstName,
    lastName,
    fullName: fullName || combo,
  };
}

export function form10HasFirstAndLastName(src = null) {
  const { firstName, lastName } = readForm10PersonNameParts(src);
  return Boolean(firstName && lastName);
}

export function isForm10FirstNameHeader(header) {
  const s = String(header || '')
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
  if (!s) return false;
  if (/last/.test(s) || /surname/.test(s) || /father/.test(s) || /husband/.test(s)) return false;
  return s === 'first name' || s === 'firstname' || s === 'first_name' || /^first\s*name$/.test(s);
}

export function isForm10LastNameHeader(header) {
  const s = String(header || '')
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
  if (!s) return false;
  return (
    s === 'last name' ||
    s === 'lastname' ||
    s === 'last_name' ||
    s === 'surname' ||
    /^last\s*name$/.test(s)
  );
}

/** Combined employee/workman name column on Form 10 (not First/Last alone). */
export function isForm10CombinedNameHeader(header) {
  if (isForm10FirstNameHeader(header) || isForm10LastNameHeader(header)) return false;
  const s = String(header || '')
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
  if (!s || /father|husband|guardian/.test(s)) return false;
  return (
    s === 'name' ||
    /employee\s+name/.test(s) ||
    /name\s+of\s+the\s+employee/.test(s) ||
    /name\s+of\s+(the\s+)?(person|workman|worker|workmen)/.test(s) ||
    (s.includes('name') && (s.includes('employee') || s.includes('workman') || s.includes('worker')))
  );
}

export function collectForm10RowNameParts(row, headers = []) {
  const out = { firstName: '', lastName: '', fullName: '' };
  if (!row || typeof row !== 'object') return out;
  const hdrs = Array.isArray(headers) ? headers : [];
  hdrs.forEach((header) => {
    const value = String(row[header] ?? '').trim();
    if (!value) return;
    if (isForm10FirstNameHeader(header) && !out.firstName) out.firstName = value;
    else if (isForm10LastNameHeader(header) && !out.lastName) out.lastName = value;
    else if (isForm10CombinedNameHeader(header) && !out.fullName) out.fullName = value;
  });
  if (row.__employeeLookupName && !out.fullName) {
    out.fullName = String(row.__employeeLookupName).trim();
  }
  if ((!out.firstName || !out.lastName) && out.fullName) {
    const parts = out.fullName.split(/\s+/).filter(Boolean);
    if (parts.length >= 2) {
      if (!out.firstName) out.firstName = parts[0];
      if (!out.lastName) out.lastName = parts[parts.length - 1];
    }
  }
  if (!out.fullName && out.firstName && out.lastName) {
    out.fullName = `${out.firstName} ${out.lastName}`.trim();
  }
  return out;
}

function form10NameTokensMatch(left, right) {
  const a = normForm10Name(left);
  const b = normForm10Name(right);
  return Boolean(a && b && a === b);
}

function form10FullNamesMatchExact(left, right) {
  const a = normForm10Name(left);
  const b = normForm10Name(right);
  if (!a || !b) return false;
  if (a === b) return true;
  const aParts = a.split(' ').filter(Boolean);
  const bParts = b.split(' ').filter(Boolean);
  if (aParts.length < 2 || bParts.length < 2) return false;
  return aParts[0] === bParts[0] && aParts[aParts.length - 1] === bParts[bParts.length - 1];
}

/**
 * Require FirstName AND LastName on both sides.
 * "Selva P" must not match "Selva Kumar" / first-name-only "Selva".
 */
export function form10FirstAndLastNamesMatch(employeeOrRow, payrollRow, extraParts = null) {
  const emp = readForm10PersonNameParts(employeeOrRow);
  const extra = extraParts && typeof extraParts === 'object' ? extraParts : {};
  const firstName = emp.firstName || extra.firstName || '';
  const lastName = emp.lastName || extra.lastName || '';
  const fullName = emp.fullName || extra.fullName || `${firstName} ${lastName}`.trim();
  if (!firstName || !lastName) return false;

  const pay = readForm10PersonNameParts(payrollRow);
  if (pay.firstName && pay.lastName) {
    return form10NameTokensMatch(firstName, pay.firstName) && form10NameTokensMatch(lastName, pay.lastName);
  }
  if (pay.fullName) {
    return (
      form10FullNamesMatchExact(`${firstName} ${lastName}`, pay.fullName) ||
      form10FullNamesMatchExact(fullName, pay.fullName)
    );
  }
  return false;
}

/** True when payroll identity does not contradict employee FirstName+LastName. */
export function form10PayrollRowAgreesWithEmployeeNames(employeeOrRow, payrollRow, extraParts = null) {
  const pay = readForm10PersonNameParts(payrollRow);
  if (!(pay.firstName && pay.lastName) && !pay.fullName) return true;
  const emp = readForm10PersonNameParts(employeeOrRow);
  const extra = extraParts && typeof extraParts === 'object' ? extraParts : {};
  const firstName = emp.firstName || extra.firstName || '';
  const lastName = emp.lastName || extra.lastName || '';
  if (!firstName || !lastName) return true;
  return form10FirstAndLastNamesMatch(employeeOrRow, payrollRow, extraParts);
}

export function findForm10PayrollRowByFirstAndLastName(employeeOrRow, payrollRows, extraParts = null) {
  const rows = (Array.isArray(payrollRows) ? payrollRows : []).filter(
    (row) => row && typeof row === 'object' && row.fetch_error !== true
  );
  if (rows.length === 0) return null;
  const hits = rows.filter((row) => form10FirstAndLastNamesMatch(employeeOrRow, row, extraParts));
  if (hits.length === 0) return null;
  return hits[0];
}

export function buildForm10EmployeeDisplayName(emp) {
  const { firstName, lastName, fullName } = readForm10PersonNameParts(emp);
  if (firstName && lastName) return `${firstName} ${lastName}`.trim();
  return fullName || firstName || lastName || '';
}

/**
 * Merge People identity with Form 10 row First/Last name columns for strict TN payroll lookup.
 */
export function buildForm10PayrollMatchIdentity(emp, row, headers = []) {
  const empParts = readForm10PersonNameParts(emp);
  const rowParts = collectForm10RowNameParts(row, headers);
  const firstName = rowParts.firstName || empParts.firstName || '';
  const lastName = rowParts.lastName || empParts.lastName || '';
  const fullName =
    rowParts.fullName || empParts.fullName || `${firstName} ${lastName}`.trim();
  const base = emp && typeof emp === 'object' ? { ...emp } : {};
  return {
    ...base,
    FirstName: firstName || base.FirstName || base.firstName || base.first_name || '',
    LastName: lastName || base.LastName || base.lastName || base.last_name || '',
    firstName: firstName || base.firstName || base.first_name || '',
    lastName: lastName || base.lastName || base.last_name || '',
    first_name: firstName || base.first_name || base.firstName || '',
    last_name: lastName || base.last_name || base.lastName || '',
    EmployeeName: fullName || base.EmployeeName || base.employee_name || '',
    employee_name: fullName || base.employee_name || base.EmployeeName || '',
    Name: fullName || base.Name || base.name || '',
    name: fullName || base.name || base.Name || '',
  };
}

/** Find a People record whose FirstName+LastName match the form row name parts. */
export function findForm10EmployeeByFirstAndLastName(employees, nameParts = null) {
  const parts = nameParts && typeof nameParts === 'object' ? nameParts : {};
  const { firstName, lastName } = parts;
  if (!firstName || !lastName) return null;
  const list = Array.isArray(employees) ? employees : [];
  for (let i = 0; i < list.length; i += 1) {
    const item = list[i];
    const emp = item?.Employee || item?.employee || item;
    if (
      form10FirstAndLastNamesMatch(emp, {
        firstName,
        lastName,
        fullName: parts.fullName || `${firstName} ${lastName}`.trim(),
      })
    ) {
      return emp;
    }
  }
  return null;
}

const form10PayrollRowIdentityKey = (row) => {
  const { firstName, lastName, fullName } = readForm10PersonNameParts(row);
  if (firstName && lastName) return `fn:${normForm10Name(firstName)}|ln:${normForm10Name(lastName)}`;
  if (fullName) return `name:${normForm10Name(fullName)}`;
  return '';
};

/**
 * Tamil Nadu strict payroll resolver — FirstName AND LastName only (Form 10 / Form XV / Form XXIII TN).
 */
export function buildFormTamilNaduPayrollRowResolver(payrollRows, options = {}) {
  const rows = (Array.isArray(payrollRows) ? payrollRows : []).filter(
    (row) => row && typeof row === 'object' && row.fetch_error !== true
  );
  const getExtraParts =
    typeof options.getExtraParts === 'function' ? options.getExtraParts : () => null;
  const usedKeys = new Set();
  return (emp, formRow) => {
    const extraParts = formRow != null ? getExtraParts(emp, formRow) : null;
    const available = rows.filter((row) => {
      const key = form10PayrollRowIdentityKey(row);
      return !key || !usedKeys.has(key);
    });
    const hit = findForm10PayrollRowByFirstAndLastName(
      emp,
      available.length > 0 ? available : rows,
      extraParts
    );
    const hitKey = form10PayrollRowIdentityKey(hit);
    if (hit && hitKey) usedKeys.add(hitKey);
    return hit;
  };
}

export function applyForm10UnmatchedPayrollAmountsNil(row, headers, nilText = 'Nil') {
  if (!row || typeof row !== 'object' || !Array.isArray(headers)) return row;
  const norm = (h) =>
    String(h || '')
      .replace(/\r?\n/g, ' ')
      .replace(/\s+/g, ' ')
      .trim()
      .toLowerCase();
  headers.forEach((header) => {
    const s = norm(header);
    if (!s) return;
    const isNormalRate = s.includes('normal') && s.includes('rate') && s.includes('pay');
    const isNormalEarnings = s.includes('normal') && s.includes('earning');
    const isTotalEarnings =
      !s.includes('overtime') && ((s.includes('total') && s.includes('earning')) || s === 'total earnings');
    if (isNormalRate || isNormalEarnings || isTotalEarnings) {
      row[header] = nilText;
    }
  });
  return row;
}
