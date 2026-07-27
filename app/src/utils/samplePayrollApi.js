import {
  fetchJsonWithTimeout,
  fetchPeopleDataForAutofillDisplay,
  flattenZohoPeopleEmployees,
  getCachedPeopleData,
} from './statutoryAutofillCache';
import { flattenPayrollEarningColumns } from './payrollEarnings';

export const SAMPLE_PAYROLL_API_BASE = '/server/samplepayroll_function';

const EMAIL_FIELD_KEYS = [
  'EmailID',
  'Email',
  'email',
  'emailId',
  'Email ID',
  'Work_Email',
  'work_email',
  'workEmail',
  'work_mail',
  'Personal_Email',
  'personal_email',
  'mail_id',
];

const PEOPLE_ID_KEYS = [
  'Zoho_ID',
  'zoho_id',
  'ZohoID',
  'zohoId',
  'EmployeeID',
  'Employee ID',
  'employee_id',
  'employee_number',
  'Employee_Number',
  'Employee Number',
  'Role.ID',
  'employeeId',
  'EmpID',
  'EmpId',
  'empId',
  'erecno',
  'id',
];

function pickFirstString(row, keys) {
  if (!row || typeof row !== 'object') return '';
  for (let i = 0; i < keys.length; i += 1) {
    const value = row[keys[i]];
    if (value != null && String(value).trim() !== '') return String(value).trim();
  }
  return '';
}

function normalizePersonName(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ');
}

export function pickEmailFromPeopleRecord(emp) {
  if (!emp || typeof emp !== 'object') return '';
  const direct = pickFirstString(emp, EMAIL_FIELD_KEYS);
  if (direct && direct.includes('@')) return direct;
  const role =
    emp.Role && typeof emp.Role === 'object' && !Array.isArray(emp.Role) ? emp.Role : null;
  if (role) {
    const roleEmail = pickFirstString(role, EMAIL_FIELD_KEYS);
    if (roleEmail && roleEmail.includes('@')) return roleEmail;
  }
  return '';
}

function pickPeopleIdCandidates(emp) {
  if (!emp || typeof emp !== 'object') return [];
  const ids = new Set();
  PEOPLE_ID_KEYS.forEach((key) => {
    const value = emp[key];
    if (value != null && String(value).trim() !== '') ids.add(String(value).trim());
  });
  if (emp.Role && typeof emp.Role === 'object' && !Array.isArray(emp.Role)) {
    const roleId = emp.Role.ID ?? emp.Role.Id ?? emp.Role.id;
    if (roleId != null && String(roleId).trim() !== '') ids.add(String(roleId).trim());
  }
  return Array.from(ids);
}

function pickPeopleNameCandidates(emp) {
  if (!emp || typeof emp !== 'object') return [];
  const names = new Set();
  const first = pickFirstString(emp, ['FirstName', 'First Name', 'first_name']);
  const last = pickFirstString(emp, ['LastName', 'Last Name', 'last_name']);
  const full = pickFirstString(emp, [
    'FullName',
    'Full Name',
    'Employee_Name',
    'Employee Name',
    'employee_name',
    'name',
  ]);
  if (full) names.add(normalizePersonName(full));
  const combined = `${first} ${last}`.trim();
  if (combined) names.add(normalizePersonName(combined));
  return Array.from(names).filter(Boolean);
}

export function buildPeopleEmailLookup(employees) {
  const byId = new Map();
  const byName = new Map();
  const list = Array.isArray(employees) ? employees : [];
  list.forEach((emp) => {
    const email = pickEmailFromPeopleRecord(emp);
    if (!email || !email.includes('@')) return;
    pickPeopleIdCandidates(emp).forEach((id) => {
      if (!byId.has(id)) byId.set(id, email);
    });
    pickPeopleNameCandidates(emp).forEach((name) => {
      if (!byName.has(name)) byName.set(name, email);
    });
  });
  return { byId, byName };
}

function payrollRowHasEmail(row) {
  const flat = flattenPayrollEarningColumns(row);
  const email = pickFirstString(flat, EMAIL_FIELD_KEYS);
  return Boolean(email && email.includes('@'));
}

export function enrichPayrollRowWithPeopleEmail(row, lookup) {
  if (!row || typeof row !== 'object' || !lookup) return row;
  if (payrollRowHasEmail(row)) return row;

  const flat = flattenPayrollEarningColumns(row);
  const empId = pickFirstString(flat, [
    'employee_id',
    'employee_number',
    'employeeId',
    'EmployeeID',
  ]);
  let email = empId ? lookup.byId.get(empId) || '' : '';
  if (!email) {
    const name = normalizePersonName(
      pickFirstString(flat, ['employee_name', 'full_name', 'name', 'Employee_Name'])
    );
    if (name) email = lookup.byName.get(name) || '';
  }
  if (!email || !email.includes('@')) return row;

  return flattenPayrollEarningColumns({
    ...row,
    ...flat,
    email,
    work_email: email,
    work_mail: email,
  });
}

export function enrichPayrollRowsWithPeopleEmail(rows, employees) {
  const lookup = buildPeopleEmailLookup(employees);
  return (Array.isArray(rows) ? rows : []).map((row) => enrichPayrollRowWithPeopleEmail(row, lookup));
}

export async function loadPeopleEmployeesForEmailEnrichment() {
  const cached = getCachedPeopleData();
  if (cached) {
    return flattenZohoPeopleEmployees(cached);
  }
  const result = await fetchPeopleDataForAutofillDisplay();
  return flattenZohoPeopleEmployees(result);
}

export async function enrichPayrollRowsWithPeopleEmailFromApi(rows) {
  try {
    const employees = await loadPeopleEmployeesForEmailEnrichment();
    if (!Array.isArray(employees) || employees.length === 0) return rows;
    return enrichPayrollRowsWithPeopleEmail(rows, employees);
  } catch (_) {
    return rows;
  }
}

export function isSamplePayrollApiSuccess(json) {
  return json?.status === 'success' || json?.success === true;
}

function pickSamplePayrollAmount(record, keys) {
  if (!record || typeof record !== 'object') return '';
  for (let i = 0; i < keys.length; i += 1) {
    const value = record[keys[i]];
    if (value == null || value === '') continue;
    const n = Number(String(value).replace(/,/g, '').trim());
    if (Number.isFinite(n)) return n;
    const s = String(value).trim();
    if (s) return s;
  }
  return '';
}

export function mapSamplePayrollRecordToPayrollRow(record) {
  if (!record || typeof record !== 'object') return null;
  const email = String(record.email || '').trim();
  const employeeName = String(record.employeeName || '').trim();
  const employeeId = String(record.employeeId || '').trim();
  const gidNumber = String(record.gidNumber || '').trim();
  const basic = pickSamplePayrollAmount(record, [
    'basic',
    'Basic',
    'earned_basic',
    'basic_pay',
    'Basic Earnings',
  ]);
  const hra = pickSamplePayrollAmount(record, [
    'hra',
    'HRA',
    'hra_fbp',
    'house_rent_allowance',
    'House Rent Allowance',
  ]);
  const gross = pickSamplePayrollAmount(record, ['gross', 'Gross', 'gross_pay', 'total_earnings']);
  const netpay = pickSamplePayrollAmount(record, ['netpay', 'Netpay', 'net_pay', 'netPay']);
  const totalDeduction = pickSamplePayrollAmount(record, [
    'totalDeduction',
    'TotalDeduction',
    'total_deductions',
  ]);
  const incomeTax = pickSamplePayrollAmount(record, [
    'incomeTax',
    'IncomeTax',
    'income_tax',
  ]);
  const pf = pickSamplePayrollAmount(record, [
    'pf',
    'PF',
    'epf_contribution',
    'EPF Contribution',
  ]);
  const voluntaryProvidentFund = pickSamplePayrollAmount(record, [
    'voluntaryProvidentFund',
    'VoluntaryProvidentFund',
    'voluntary_provident_fund',
    'Voluntary Provident Fund',
    'vpf',
    'VPF',
  ]);
  const professionalTax = pickSamplePayrollAmount(record, [
    'professionalTax',
    'ProfessionalTax',
    'professional_tax',
  ]);
  return flattenPayrollEarningColumns({
    employee_name: employeeName,
    full_name: employeeName,
    name: employeeName,
    employee_id: employeeId,
    employee_number: gidNumber || employeeId,
    GIDNumber: gidNumber,
    gidNumber,
    email,
    work_email: email,
    work_mail: email,
    date_of_birth: record.dateofBirth,
    Date_of_birth: record.dateofBirth,
    DateofBirth: record.dateofBirth,
    dateofBirth: record.dateofBirth,
    paid_days: record.paidDays,
    basic,
    earned_basic: basic,
    Basic: basic,
    hra,
    hra_fbp: hra,
    HRA: hra,
    house_rent_allowance: hra,
    gross_pay: gross,
    total_earnings: gross,
    gross,
    net_pay: netpay,
    netPay: netpay,
    total_deductions: totalDeduction,
    income_tax: incomeTax,
    IncomeTax: incomeTax,
    epf_contribution: pf,
    PF: pf,
    pf,
    voluntary_provident_fund: voluntaryProvidentFund,
    VoluntaryProvidentFund: voluntaryProvidentFund,
    vpf: voluntaryProvidentFund,
    professional_tax: professionalTax,
    ProfessionalTax: professionalTax,
    payrollMonth: record.payrollMonth,
  });
}

function pickPayrollEmployeeId(row) {
  const flat = flattenPayrollEarningColumns(row);
  for (const key of ['employee_id', 'employee_number', 'employeeId', 'EmployeeID']) {
    const value = flat[key];
    if (value != null && String(value).trim() !== '') return String(value).trim();
  }
  return '';
}

function pickPayrollGidNumber(row) {
  const flat = flattenPayrollEarningColumns(row);
  for (const key of ['gidNumber', 'GIDNumber', 'gid_number', 'GID Number', 'GID_Number', 'employee_number']) {
    const value = flat[key];
    if (value != null && String(value).trim() !== '') return String(value).trim();
  }
  return '';
}

function samplePayrollRowMatchesEmployeeId(row, employeeId) {
  const id = String(employeeId || '').trim();
  if (!id || !row) return false;
  if (pickPayrollEmployeeId(row) === id) return true;
  const gid = pickPayrollGidNumber(row);
  return gid && gid === id;
}

function normalizePayrollMonthCandidates(monthCandidates) {
  const list = Array.isArray(monthCandidates) ? monthCandidates : [];
  const out = [];
  const seen = new Set();
  list.forEach((value) => {
    const month = String(value || '').trim();
    if (!/^\d{4}-\d{2}$/.test(month) || seen.has(month)) return;
    seen.add(month);
    out.push(month);
  });
  return out;
}

export async function fetchSamplePayrollRecords(payrollMonth, { timeoutMs = 45000 } = {}) {
  const month = String(payrollMonth || '').trim();
  const qs = new URLSearchParams({ fetch_all: '1' });
  if (/^\d{4}-\d{2}$/.test(month)) qs.set('payroll_month', month);

  try {
    const { resp, json } = await fetchJsonWithTimeout(
      `${SAMPLE_PAYROLL_API_BASE}/samplepayroll?${qs.toString()}`,
      { cache: 'no-store' },
      timeoutMs
    );
    if (!resp.ok || !isSamplePayrollApiSuccess(json)) {
      return { records: [], payrollMonth: month };
    }
    const raw = Array.isArray(json.data?.records) ? json.data.records : [];
    return {
      records: raw.map(mapSamplePayrollRecordToPayrollRow).filter(Boolean),
      payrollMonth: json.data?.payrollMonth || month,
    };
  } catch (_) {
    return { records: [], payrollMonth: month };
  }
}

/** @deprecated Statutory reads SamplePayroll table only — not the Payroll JSON snapshot table. */
export async function fetchSamplePayrollTableSnapshot(payrollMonth) {
  return { records: [], meta: null, payrollMonth: String(payrollMonth || '').trim() };
}

export async function fetchSamplePayrollRowsForMonth(payrollMonth, { timeoutMs = 45000 } = {}) {
  const listLoad = await fetchSamplePayrollRecords(payrollMonth, { timeoutMs });
  return {
    records: listLoad.records,
    meta: null,
    payrollMonth: listLoad.payrollMonth || payrollMonth,
    source: listLoad.records.length > 0 ? 'sample_payroll' : 'none',
  };
}

export async function fetchLatestSamplePayrollRows({ timeoutMs = 45000 } = {}) {
  const listLoad = await fetchSamplePayrollRecords('', { timeoutMs });
  if (listLoad.records.length === 0) {
    return { records: [], meta: null, payrollMonth: '', source: 'none' };
  }

  const months = [
    ...new Set(
      listLoad.records
        .map((row) => String(row.payrollMonth || row.PayrollMonth || '').trim())
        .filter((month) => /^\d{4}-\d{2}$/.test(month))
    ),
  ].sort();

  const latestMonth = months[months.length - 1] || '';
  const monthRows = latestMonth
    ? listLoad.records.filter(
        (row) => String(row.payrollMonth || row.PayrollMonth || '').trim() === latestMonth
      )
    : listLoad.records;

  return {
    records: monthRows.length > 0 ? monthRows : listLoad.records,
    meta: null,
    payrollMonth: latestMonth,
    source: 'sample_payroll_latest',
  };
}

export async function fetchSamplePayrollEmployeeRow(
  employeeId,
  monthCandidates = [],
  { timeoutMs = 12000 } = {}
) {
  const id = String(employeeId || '').trim();
  if (!id) return null;

  const months = normalizePayrollMonthCandidates(monthCandidates);
  for (let i = 0; i < months.length; i += 1) {
    const listLoad = await fetchSamplePayrollRecords(months[i], { timeoutMs });
    const listHit = listLoad.records.find((row) => samplePayrollRowMatchesEmployeeId(row, id));
    if (listHit) return listHit;
  }

  if (months.length === 0) {
    const listLoad = await fetchSamplePayrollRecords('', { timeoutMs });
    const listHit = listLoad.records.find((row) => samplePayrollRowMatchesEmployeeId(row, id));
    if (listHit) return listHit;
  }

  return null;
}
