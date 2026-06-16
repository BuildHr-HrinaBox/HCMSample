'use strict';

function toNumber(value) {
  if (value == null || value === '') return NaN;
  if (typeof value === 'number') return Number.isFinite(value) ? value : NaN;
  const n = Number(String(value).replace(/,/g, '').trim());
  return Number.isFinite(n) ? n : NaN;
}

function coalesceAmount(...values) {
  for (let i = 0; i < values.length; i += 1) {
    const n = toNumber(values[i]);
    if (Number.isFinite(n)) return n;
  }
  return '';
}

function normalizeKey(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '');
}

function payrollEarningType(item) {
  return normalizeKey(
    item?.type ??
      item?.earning_type ??
      item?.salary_component_type ??
      item?.component_type ??
      ''
  );
}

function payrollEarningName(item) {
  return String(
    item?.name ??
      item?.earning_name ??
      item?.component_name ??
      item?.salary_component_name ??
      item?.label ??
      item?.display_name ??
      ''
  )
    .trim()
    .toLowerCase();
}

function getPayrollLineAmount(lineItem) {
  if (!lineItem || typeof lineItem !== 'object') return toNumber(lineItem);
  return toNumber(
    lineItem.earned_amount ??
      lineItem.actual_amount ??
      lineItem.actual_earning_amount ??
      lineItem.calculated_amount ??
      lineItem.payable_amount ??
      lineItem.component_amount ??
      lineItem.amount ??
      lineItem.value ??
      lineItem.earning_amount ??
      lineItem.monthly_amount ??
      lineItem.Amount
  );
}

function objectMapToEarnings(obj) {
  if (!obj || typeof obj !== 'object' || Array.isArray(obj)) return [];
  return Object.entries(obj).map(([key, value]) => {
    if (value && typeof value === 'object' && !Array.isArray(value)) {
      return { name: value.name || value.earning_name || key, type: value.type || key, ...value };
    }
    return { name: key, type: key, amount: value };
  });
}

function pushEarningsFromCandidate(candidate, collected) {
  if (Array.isArray(candidate)) {
    collected.push(...candidate);
    return;
  }
  if (typeof candidate === 'string') {
    try {
      pushEarningsFromCandidate(JSON.parse(candidate), collected);
    } catch (_) {
      /* ignore */
    }
    return;
  }
  if (candidate && typeof candidate === 'object') {
    collected.push(...objectMapToEarnings(candidate));
  }
}

/** Unwrap Zoho salary/payrun payloads: { employee: {...} } or nested salary.employee */
function normalizePayrollEmployee(row) {
  if (!row || typeof row !== 'object') return {};
  if (row.employee && typeof row.employee === 'object' && !Array.isArray(row.employee)) {
    return row.employee;
  }
  if (row.salary && typeof row.salary === 'object') {
    if (row.salary.employee && typeof row.salary.employee === 'object') return row.salary.employee;
    if (Array.isArray(row.salary.earnings) || Array.isArray(row.salary.fbp_components)) return row.salary;
  }
  return row;
}

function getEarningsArray(row) {
  if (!row || typeof row !== 'object') return [];
  const employee = normalizePayrollEmployee(row);
  const payrollEmployee =
    row.payroll_employee && typeof row.payroll_employee === 'object' ? row.payroll_employee : null;
  const collected = [];
  [
    row.earnings,
    employee.earnings,
    employee.fbp_components,
    employee.variable_earnings,
    employee.reimbursements,
    payrollEmployee?.earnings,
    payrollEmployee?.fbp_components,
    payrollEmployee?.variable_earnings,
    row.fbp_components,
    row.variable_earnings,
    row.employee_earnings,
    row.salary_components,
    row.earning_components,
    row.earning_details,
    row.regular_earnings,
    row.additional_earnings,
    row.one_time_earnings,
    row.payroll?.earnings,
    row.employee_payroll?.earnings,
  ].forEach((candidate) => pushEarningsFromCandidate(candidate, collected));
  return collected;
}

function pickScalarAmount(row, keys) {
  if (!row || typeof row !== 'object') return '';
  const sources = [row, normalizePayrollEmployee(row)];
  const wanted = keys.map((key) => normalizeKey(key));
  for (const source of sources) {
    for (const [rawKey, value] of Object.entries(source)) {
      const normKey = normalizeKey(rawKey);
      if (!wanted.includes(normKey)) continue;
      if (value != null && typeof value === 'object') {
        const nested = getPayrollLineAmount(value);
        if (Number.isFinite(nested)) return nested;
        continue;
      }
      const n = toNumber(value);
      if (Number.isFinite(n)) return n;
    }
  }
  return '';
}

function pickAmountByPatterns(row, patterns) {
  const sources = [row, normalizePayrollEmployee(row)];
  for (const source of sources) {
    if (!source || typeof source !== 'object') continue;
    for (const [rawKey, value] of Object.entries(source)) {
      if (value == null || value === '') continue;
      if (Array.isArray(value) || (typeof value === 'object' && value !== null)) continue;
      const normKey = normalizeKey(rawKey);
      const label = String(rawKey).trim().toLowerCase();
      for (let i = 0; i < patterns.length; i += 1) {
        if (patterns[i].test(normKey) || patterns[i].test(label)) {
          const n = toNumber(value);
          if (Number.isFinite(n)) return n;
        }
      }
    }
  }
  return '';
}

function findEarningAmount(earnings, matcher) {
  const hit = earnings.find((item) => matcher(payrollEarningType(item), payrollEarningName(item)));
  if (!hit) return '';
  const amount = getPayrollLineAmount(hit);
  return Number.isFinite(amount) ? amount : '';
}

function pickBasicAmount(earnings) {
  const exact = earnings.find((item) => {
    const name = payrollEarningName(item);
    const type = payrollEarningType(item);
    return (
      name === 'basic' ||
      name === 'basic earnings' ||
      name === 'basic pay' ||
      name === 'earned basic' ||
      type === 'basic' ||
      type === 'earned_basic'
    );
  });
  if (exact) {
    const amount = getPayrollLineAmount(exact);
    if (Number.isFinite(amount)) return amount;
  }
  return findEarningAmount(
    earnings,
    (type, name) =>
      type === 'basic' ||
      type === 'earned_basic' ||
      name === 'basic' ||
      name.includes('basic earnings') ||
      name.includes('basic pay') ||
      name.includes('earned basic') ||
      name.includes('basic wage')
  );
}

function pickHraAmount(earnings) {
  const exact = earnings.find((item) => {
    const name = payrollEarningName(item);
    const type = payrollEarningType(item);
    return (
      name === 'hra' ||
      name === 'house rent allowance' ||
      name.includes('hra (fbp)') ||
      name.includes('house rent') ||
      type === 'hra'
    );
  });
  if (exact) {
    const amount = getPayrollLineAmount(exact);
    if (Number.isFinite(amount)) return amount;
  }
  return findEarningAmount(
    earnings,
    (type, name) => type === 'hra' || name.includes('house rent') || /\bhra\b/.test(name)
  );
}

function pickOtherAllowanceAmount(earnings) {
  const exact = earnings.find((item) => {
    const name = payrollEarningName(item);
    return name === 'other allowance' || name.includes('supplementary allowance');
  });
  if (exact) {
    const amount = getPayrollLineAmount(exact);
    if (Number.isFinite(amount)) return amount;
  }
  return findEarningAmount(
    earnings,
    (type, name) =>
      name.includes('other allowance') ||
      name.includes('supplementary allowance') ||
      type === 'other_allowance' ||
      (type === 'fixed' && name.includes('allowance') && !name.includes('house rent'))
  );
}

function readPayrollForm15WageAmounts(payrollRow) {
  const flat = flattenPayrollEarningColumns(payrollRow || {});
  return {
    flat,
    basic: coalesceAmount(flat.basic, flat.earned_basic),
    hra: coalesceAmount(flat.hra_fbp, flat.hra),
    other_allowance: coalesceAmount(flat.other_allowance),
  };
}

function flattenPayrollEarningColumns(row) {
  if (!row || typeof row !== 'object') return row;
  const employee = normalizePayrollEmployee(row);
  const earnings = getEarningsArray(row);
  const componentColumns = {};
  earnings.forEach((item) => {
    const name = payrollEarningName(item);
    if (!name) return;
    const amount = getPayrollLineAmount(item);
    if (!Number.isFinite(amount)) return;
    const key = normalizeKey(name);
    if (key) componentColumns[key] = amount;
  });

  const earned_basic = coalesceAmount(
    pickBasicAmount(earnings),
    componentColumns.basic,
    componentColumns.basic_earnings,
    componentColumns.basic_pay,
    componentColumns.earned_basic,
    pickScalarAmount(row, [
      'earned_basic',
      'Earned Basic',
      'basic_pay',
      'basic',
      'Basic',
      'basic_earnings',
      'Basic Earnings',
    ]),
    pickAmountByPatterns(row, [/^earned_basic$/, /^basic_earnings?$/, /^basic_pay$/, /^basic$/])
  );

  const basic = coalesceAmount(
    pickScalarAmount(row, ['basic', 'Basic']),
    componentColumns.basic,
    earned_basic
  );

  const hra = coalesceAmount(
    pickHraAmount(earnings),
    componentColumns.hra,
    componentColumns.hra_fbp,
    componentColumns.house_rent_allowance,
    pickScalarAmount(row, ['hra', 'HRA', 'hra_fbp', 'house_rent_allowance', 'House Rent Allowance']),
    pickAmountByPatterns(row, [/^hra(_fbp)?$/, /^house_rent_allowance$/, /house.*rent/])
  );

  const hra_fbp = coalesceAmount(
    componentColumns.hra_fbp,
    pickScalarAmount(row, ['hra_fbp', 'HRA (FBP)', 'hra (fbp)']),
    hra
  );

  const other_allowance = coalesceAmount(
    pickOtherAllowanceAmount(earnings),
    componentColumns.other_allowance,
    componentColumns.supplementary_allowance,
    pickScalarAmount(row, ['other_allowance', 'Other Allowance', 'Supplementary Allowance']),
    pickAmountByPatterns(row, [/^other_allowance$/, /supplementary.*allowance/])
  );

  return {
    ...row,
    ...employee,
    ...componentColumns,
    earned_basic,
    basic,
    hra,
    hra_fbp,
    other_allowance,
    dearness_allowance: coalesceAmount(
      findEarningAmount(earnings, (type, name) => type === 'da' || name.includes('dearness')),
      componentColumns.dearness_allowance,
      pickScalarAmount(row, ['dearness_allowance', 'Dearness Allowance', 'da']),
      pickAmountByPatterns(row, [/dearness/])
    ),
    conveyance_allowance: coalesceAmount(
      findEarningAmount(
        earnings,
        (type, name) =>
          name.includes('conveyance') ||
          type === 'conveyance' ||
          name.includes('fuel allowance')
      ),
      componentColumns.conveyance_allowance,
      componentColumns.fuel_allowance,
      pickScalarAmount(row, ['conveyance_allowance', 'Conveyance Allowance', 'Fuel Allowance']),
      pickAmountByPatterns(row, [/conveyance/, /fuel.*allowance/])
    ),
    overtime: coalesceAmount(
      findEarningAmount(
        earnings,
        (type, name) => type === 'overtime' || type === 'ot' || name.includes('overtime')
      ),
      componentColumns.overtime,
      pickScalarAmount(row, ['overtime', 'Overtime', 'overtime_wages']),
      pickAmountByPatterns(row, [/overtime/])
    ),
    gross_pay: coalesceAmount(
      pickScalarAmount(row, ['gross_pay', 'Gross Pay', 'grossPay', 'total_earnings', 'monthly_gross_amount']),
      pickAmountByPatterns(row, [/^gross_pay$/, /^total_earnings$/])
    ),
    net_pay: coalesceAmount(
      pickScalarAmount(row, ['net_pay', 'Net Pay', 'netPay', 'monthly_salary', 'MonthlySalary']),
      pickAmountByPatterns(row, [/^net_pay$/, /^monthly_salary$/])
    ),
    total_deductions: coalesceAmount(
      pickScalarAmount(row, [
        'total_deductions',
        'Total Deductions',
        'totalDeductions',
        'total_employee_deductions',
        'total_deduction',
      ]),
      pickAmountByPatterns(row, [/^total_deductions?$/, /^total_employee_deductions$/])
    ),
  };
}

function mergePayrollRunEmployeePayload(data) {
  if (!data || typeof data !== 'object') return data;
  const employee =
    data.employee && typeof data.employee === 'object' ? { ...data.employee } : { ...data };
  const nestedKeys = [
    'earnings',
    'deductions',
    'benefits',
    'taxes',
    'donations',
    'reimbursements',
    'earning_details',
    'salary_components',
    'fbp_components',
    'variable_earnings',
  ];
  const nestedSources = [data];
  if (data.payroll_employee && typeof data.payroll_employee === 'object') {
    nestedSources.push(data.payroll_employee);
  }
  nestedSources.forEach((source) => {
    nestedKeys.forEach((key) => {
      if (source[key] != null && employee[key] == null) employee[key] = source[key];
    });
  });
  Object.entries(data).forEach(([key, value]) => {
    if (['employee', 'code', 'message', 'page_context', 'payroll_employee'].includes(key)) return;
    if (nestedKeys.includes(key)) return;
    if (Array.isArray(value)) return;
    if (value != null && typeof value !== 'object') {
      if (employee[key] == null || employee[key] === '') employee[key] = value;
    }
  });
  return employee;
}

function unwrapSalaryEmployeePayload(body) {
  if (!body || typeof body !== 'object') return body;
  return mergePayrollRunEmployeePayload(body);
}

module.exports = {
  flattenPayrollEarningColumns,
  getEarningsArray,
  mergePayrollRunEmployeePayload,
  normalizePayrollEmployee,
  readPayrollForm15WageAmounts,
  unwrapSalaryEmployeePayload,
};
