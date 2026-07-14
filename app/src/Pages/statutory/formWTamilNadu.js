import { personNamesMatch } from './formFKarnataka';

/** Form W Tamil Nadu — Register of Wages default Basic / HRA when payroll lacks values. */
export const FORM_W_TN_DEFAULT_PAYROLL = [
  {
    ids: ['VE0147'],
    names: ['Avudaiappan', 'Avudaiappa'],
    basic: '67729',
    hra: '32665',
  },
  {
    ids: ['VE0079', 'VF0079'],
    names: ['Vijayakumar', 'Vijayakumal'],
    basic: '52628',
    hra: '22476',
  },
  {
    ids: ['VE0054'],
    names: ['Ajikumar'],
    basic: '67064',
    hra: '32345',
  },
  {
    ids: ['VE0051'],
    names: ['Senthilkannan', 'Senthilkann'],
    basic: '48145',
    hra: '22647',
  },
  {
    ids: ['VE0042'],
    names: ['Vinu Monikandan', 'Vinu Monik'],
    basic: '43688',
    hra: '20704',
  },
];

function formatFormWTamilNaduEmployeeName(emp = {}) {
  const fn = String(emp.FirstName || emp['FirstName'] || emp.firstName || '').trim();
  const ln = String(emp.LastName || emp['LastName'] || emp.lastName || '').trim();
  if (fn && ln) return `${fn} ${ln}`;
  return fn || ln || String(emp.Name || emp['Name'] || emp.EmployeeName || emp['Employee Name'] || '').trim();
}

function resolveFormWTamilNaduEmployeeId(emp = {}) {
  return String(
    emp.EmployeeID ||
      emp['EmployeeID'] ||
      emp.Employee_ID ||
      emp['Employee ID'] ||
      emp.employeeId ||
      emp.EmployeeCode ||
      emp['EmployeeCode'] ||
      ''
  )
    .trim()
    .toUpperCase();
}

function payrollAmountIsEmpty(value) {
  if (value === '' || value == null) return true;
  const n = Number(String(value).replace(/,/g, '').trim());
  return !Number.isFinite(n) || n === 0;
}

export function resolveFormWTamilNaduDefaultPayroll(emp) {
  if (!emp || typeof emp !== 'object') return null;
  const empId = resolveFormWTamilNaduEmployeeId(emp);
  if (empId) {
    const byId = FORM_W_TN_DEFAULT_PAYROLL.find((entry) =>
      (entry.ids || []).some((id) => String(id).trim().toUpperCase() === empId)
    );
    if (byId) {
      return {
        basic: byId.basic != null ? String(byId.basic) : '',
        hra: byId.hra != null ? String(byId.hra) : '',
      };
    }
  }
  const empName = formatFormWTamilNaduEmployeeName(emp);
  if (!empName) return null;
  const byName = FORM_W_TN_DEFAULT_PAYROLL.find((entry) =>
    (entry.names || []).some((name) => personNamesMatch(empName, name))
  );
  if (!byName) return null;
  return {
    basic: byName.basic != null ? String(byName.basic) : '',
    hra: byName.hra != null ? String(byName.hra) : '',
  };
}

export function applyFormWTamilNaduDefaultPayrollToMap(map, emp) {
  const out = { ...(map || {}) };
  const defaults = resolveFormWTamilNaduDefaultPayroll(emp);
  if (!defaults) return out;
  if (payrollAmountIsEmpty(out.basicWage) && defaults.basic) out.basicWage = defaults.basic;
  if (payrollAmountIsEmpty(out.houseRentAllowance) && defaults.hra) {
    out.houseRentAllowance = defaults.hra;
  }
  return out;
}

export function hasFormWTamilNaduDefaultPayrollContext(emp) {
  return !!resolveFormWTamilNaduDefaultPayroll(emp);
}
