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
export function normalizePayrollEmployee(row) {
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

function payrollComponentType(item) {
  return normalizeKey(
    item?.type ??
      item?.deduction_type ??
      item?.tax_type ??
      item?.benefit_type ??
      item?.salary_component_type ??
      item?.component_type ??
      ''
  );
}

function payrollComponentName(item) {
  return String(
    item?.name ??
      item?.deduction_name ??
      item?.tax_name ??
      item?.benefit_name ??
      item?.component_name ??
      item?.label ??
      item?.display_name ??
      ''
  )
    .trim()
    .toLowerCase();
}

function pushPayrollLineItemsFromCandidate(candidate, collected) {
  if (Array.isArray(candidate)) {
    collected.push(...candidate);
    return;
  }
  if (typeof candidate === 'string') {
    try {
      const parsed = JSON.parse(candidate);
      if (Array.isArray(parsed)) collected.push(...parsed);
    } catch (_) {
      /* ignore */
    }
    return;
  }
  if (candidate && typeof candidate === 'object') {
    collected.push(...objectMapToEarnings(candidate));
  }
}

export function getEarningsArray(row) {
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

export function getDeductionsArray(row) {
  if (!row || typeof row !== 'object') return [];
  const employee = normalizePayrollEmployee(row);
  const payrollEmployee =
    row.payroll_employee && typeof row.payroll_employee === 'object' ? row.payroll_employee : null;
  const collected = [];
  [
    row.deductions,
    row.deduction,
    row.employee_deductions,
    row.deduction_components,
    row.deduction_details,
    row.statutory_deductions,
    employee.deductions,
    payrollEmployee?.deductions,
    row.employee_salary_details?.deductions,
    row.salary?.deductions,
    row.salary_details?.deductions,
    row.pay_structure?.deductions,
    row.employee_salary?.deductions,
  ].forEach((candidate) => pushPayrollLineItemsFromCandidate(candidate, collected));
  return collected;
}

export function getTaxesArray(row) {
  if (!row || typeof row !== 'object') return [];
  const employee = normalizePayrollEmployee(row);
  const payrollEmployee =
    row.payroll_employee && typeof row.payroll_employee === 'object' ? row.payroll_employee : null;
  const collected = [];
  [
    row.taxes,
    row.tax,
    row.employee_taxes,
    employee.taxes,
    payrollEmployee?.taxes,
    row.employee_salary_details?.taxes,
    row.salary?.taxes,
    row.salary_details?.taxes,
    row.pay_structure?.taxes,
    row.employee_salary?.taxes,
  ].forEach((candidate) => pushPayrollLineItemsFromCandidate(candidate, collected));
  return collected;
}

function findPayrollComponentAmount(items, matcher) {
  const list = Array.isArray(items) ? items : [];
  const hit = list.find((item) => matcher(payrollComponentType(item), payrollComponentName(item)));
  if (!hit) return '';
  const amount = getPayrollLineAmount(hit);
  return Number.isFinite(amount) ? amount : '';
}

function indexPayrollLineItems(items, componentColumns) {
  const list = Array.isArray(items) ? items : [];
  list.forEach((item) => {
    const type = payrollComponentType(item);
    const name = payrollComponentName(item);
    const amount = getPayrollLineAmount(item);
    if (!Number.isFinite(amount)) return;
    if (type) componentColumns[type] = amount;
    const nameKey = normalizeKey(name);
    if (nameKey) componentColumns[nameKey] = amount;
  });
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

function pickTextScalar(row, keys) {
  if (!row || typeof row !== 'object') return '';
  const sources = [row, normalizePayrollEmployee(row)];
  const wanted = keys.map((key) => normalizeKey(key));
  for (const source of sources) {
    for (const [rawKey, value] of Object.entries(source)) {
      const normKey = normalizeKey(rawKey);
      if (!wanted.includes(normKey)) continue;
      if (value == null || typeof value === 'object') continue;
      const text = String(value).trim();
      if (text !== '') return text;
    }
  }
  return '';
}

function pickTextByPatterns(row, patterns) {
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
          const text = String(value).trim();
          if (text !== '') return text;
        }
      }
    }
  }
  return '';
}

/** Read a numeric payroll field by explicit keys and/or key-name regex patterns. */
export function readPayrollScalar(row, keys = [], patterns = []) {
  if (!row || row.fetch_error) return '';
  const fromKeys = keys.length > 0 ? pickScalarAmount(row, keys) : '';
  if (fromKeys !== '') return fromKeys;
  if (patterns.length > 0) {
    const fromPatterns = pickAmountByPatterns(row, patterns);
    if (fromPatterns !== '') return fromPatterns;
  }
  return '';
}

/** Read a text payroll field by explicit keys and/or key-name regex patterns. */
export function readPayrollTextScalar(row, keys = [], patterns = []) {
  if (!row || row.fetch_error) return '';
  const fromKeys = keys.length > 0 ? pickTextScalar(row, keys) : '';
  if (fromKeys !== '') return fromKeys;
  if (patterns.length > 0) {
    const fromPatterns = pickTextByPatterns(row, patterns);
    if (fromPatterns !== '') return fromPatterns;
  }
  return '';
}

/** Net pay from explicit net_pay fields only (no gross/total fallbacks). */
export function readStrictPayrollNetPay(payrollRow) {
  if (!payrollRow || payrollRow.fetch_error) return '';
  const flat = flattenPayrollEarningColumns(payrollRow);
  const amount = readPayrollScalar(flat, ['net_pay', 'Net Pay', 'netPay', 'netWages'], [/^net_pay$/, /^netwages$/]);
  if (amount === '') return '';
  const n = toNumber(amount);
  return Number.isFinite(n) && n > 0 ? n : '';
}

/** Net pay for statutory autofill — net_pay with monthly_salary / net salary fallbacks, not gross. */
export function readPayrollNetPayForStatutory(payrollRow) {
  if (!payrollRow || payrollRow.fetch_error) return '';
  const strict = readStrictPayrollNetPay(payrollRow);
  if (strict !== '') return strict;
  const flat = flattenPayrollEarningColumns(payrollRow);
  const amount = readPayrollScalar(
    flat,
    ['monthly_salary', 'MonthlySalary', 'monthly_net_pay', 'net_salary', 'netSalary'],
    [/^monthly_salary$/, /^net_salary$/, /^monthly_net_pay$/]
  );
  if (amount === '') return '';
  const n = toNumber(amount);
  return Number.isFinite(n) && n > 0 ? n : '';
}

/** Unwrap nested Zoho salary / employee_salary / data wrappers into a plain payroll object. */
export function getPayrollPayloadObject(data) {
  const parseMaybeJson = (v) => {
    if (!v) return {};
    if (typeof v === 'object') return v;
    if (typeof v === 'string') {
      try {
        const parsed = JSON.parse(v);
        return parsed && typeof parsed === 'object' ? parsed : {};
      } catch (_) {
        return {};
      }
    }
    return {};
  };

  if (!data) return {};
  if (Array.isArray(data)) {
    const first = data[0];
    if (!first) return {};
    if (typeof first === 'string') return parseMaybeJson(first);
    return getPayrollPayloadObject(first);
  }
  if (typeof data !== 'object') return {};
  if (data.salary) return getPayrollPayloadObject(parseMaybeJson(data.salary));
  if (data.employee_salary) return getPayrollPayloadObject(parseMaybeJson(data.employee_salary));
  if (data.data) return getPayrollPayloadObject(parseMaybeJson(data.data));
  if (data.employee && typeof data.employee === 'object') {
    return { ...parseMaybeJson(data.employee), ...data };
  }
  return data;
}

export function findEarningAmount(earnings, matcher) {
  const list = Array.isArray(earnings) ? earnings : [];
  const hit = list.find((item) => matcher(payrollEarningType(item), payrollEarningName(item)));
  if (!hit) return '';
  const amount = getPayrollLineAmount(hit);
  return Number.isFinite(amount) ? amount : '';
}

export function findDeductionAmount(deductions, matcher) {
  const list = Array.isArray(deductions) ? deductions : [];
  const hit = list.find((item) =>
    matcher(payrollComponentType(item), payrollComponentName(item))
  );
  if (!hit) return '';
  const amount = getPayrollLineAmount(hit);
  return Number.isFinite(amount) ? amount : '';
}

export function sumPayrollLineItems(items) {
  const list = Array.isArray(items) ? items : [];
  let total = 0;
  let hasAny = false;
  list.forEach((it) => {
    const amount = getPayrollLineAmount(it);
    if (Number.isFinite(amount)) {
      total += amount;
      hasAny = true;
    }
  });
  return hasAny ? total : '';
}

function pickBasicAmount(earnings) {
  const exact = earnings.find((item) => {
    const name = payrollEarningName(item);
    const type = payrollEarningType(item);
    return (
      name === 'basic' ||
      name === 'basic earnings' ||
      name === 'basic pay' ||
      name === 'basic salary' ||
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
      (name.includes('basic') &&
        !name.includes('arrear') &&
        !name.includes('overtime') &&
        !name.includes('special'))
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

/** Form 15 Part 2 wage columns from flattened payroll row (basic, hra_fbp, other_allowance). */
export function readPayrollForm15WageAmounts(payrollRow) {
  const flat = flattenPayrollEarningColumns(payrollRow || {});
  return {
    flat,
    basic: coalesceAmount(flat.basic, flat.earned_basic),
    hra: coalesceAmount(flat.hra_fbp, flat.hra),
    hra_fbp: coalesceAmount(flat.hra_fbp, flat.hra),
    other_allowance: coalesceAmount(flat.other_allowance),
  };
}

function payrollAmountIsPresent(value) {
  return coalesceAmount(value) !== '';
}

/** True when a payroll row has explicit Basic or HRA amounts (not just gross/net summary). */
export function payrollRowHasWageBreakdown(payrollRow) {
  const wages = readPayrollForm15WageAmounts(payrollRow);
  return payrollAmountIsPresent(wages.basic) || payrollAmountIsPresent(wages.hra);
}

/** True when a Payroll table row has basic / hra_fbp / other_allowance scalars. */
export function payrollRowHasForm15Part2WageFields(payrollRow) {
  if (!payrollRow || payrollRow.fetch_error) return false;
  const flat = flattenPayrollEarningColumns(payrollRow);
  return (
    payrollAmountIsPresent(flat.basic) ||
    payrollAmountIsPresent(flat.earned_basic) ||
    payrollAmountIsPresent(flat.hra_fbp) ||
    payrollAmountIsPresent(flat.hra) ||
    payrollAmountIsPresent(flat.other_allowance)
  );
}

export function payrollRowsHaveForm15Part2WageFields(rows) {
  if (!Array.isArray(rows) || rows.length === 0) return false;
  return rows.some((row) => payrollRowHasForm15Part2WageFields(row));
}

export function payrollRowsHaveWageBreakdown(rows) {
  if (!Array.isArray(rows) || rows.length === 0) return false;
  return rows.some((row) => row && !row.fetch_error && payrollRowHasWageBreakdown(row));
}

/** Form 10 — net_pay from Payroll table / Zoho pay run (Total earnings). */
export function readForm10NetPayAmount(payrollRow) {
  if (!payrollRow || payrollRow.fetch_error) return '';
  const flat = flattenPayrollEarningColumns(payrollRow);
  const candidates = [flat.net_pay, payrollRow.net_pay];
  for (let i = 0; i < candidates.length; i += 1) {
    const n = Number(String(candidates[i] ?? '').replace(/,/g, '').trim());
    if (Number.isFinite(n) && n > 0) return n;
  }
  return '';
}

/** Form 10 — gross_pay from Payroll table / Zoho pay run (Normal earnings). */
export function readForm10GrossPayAmount(payrollRow) {
  if (!payrollRow || payrollRow.fetch_error) return '';
  const flat = flattenPayrollEarningColumns(payrollRow);
  const candidates = [flat.gross_pay, payrollRow.gross_pay];
  for (let i = 0; i < candidates.length; i += 1) {
    const n = Number(String(candidates[i] ?? '').replace(/,/g, '').trim());
    if (Number.isFinite(n) && n > 0) return n;
  }
  return '';
}

export function payrollRowsHaveForm10NetPay(rows) {
  if (!Array.isArray(rows) || rows.length === 0) return false;
  return rows.some((row) => readForm10NetPayAmount(row) !== '');
}

/** True when a payroll row has a usable net/gross pay amount for Form 10 autofill. */
export function payrollRowHasNetPay(payrollRow) {
  if (!payrollRow || payrollRow.fetch_error) return false;
  const flat = flattenPayrollEarningColumns(payrollRow);
  const candidates = [
    flat.net_pay,
    flat.total_earnings,
    flat.gross_pay,
    payrollRow.net_pay,
    payrollRow.total_earnings,
    payrollRow.gross_pay,
    payrollRow.monthly_salary,
    flat.monthly_salary,
  ];
  return candidates.some((value) => {
    const n = Number(String(value ?? '').replace(/,/g, '').trim());
    return Number.isFinite(n) && n > 0;
  });
}

/** True when a payroll row has gross_pay / total_earnings (Form XIV wage rate, Form Q, etc.). */
export function payrollRowHasGrossPayAmount(payrollRow) {
  if (!payrollRow || payrollRow.fetch_error) return false;
  const flat = flattenPayrollEarningColumns(payrollRow);
  const amount = readPayrollScalar(
    flat,
    ['gross_pay', 'Gross Pay', 'grossPay', 'total_earnings', 'monthly_gross_amount'],
    [/^gross_pay$/, /^total_earnings$/]
  );
  if (amount !== '') {
    const n = Number(String(amount).replace(/,/g, '').trim());
    return Number.isFinite(n) && n > 0;
  }
  if (flat !== payrollRow) {
    const top = readPayrollScalar(
      payrollRow,
      ['gross_pay', 'Gross Pay', 'grossPay', 'total_earnings', 'monthly_gross_amount'],
      [/^gross_pay$/, /^total_earnings$/]
    );
    if (top !== '') {
      const n = Number(String(top).replace(/,/g, '').trim());
      return Number.isFinite(n) && n > 0;
    }
  }
  return false;
}

/** Payroll table rows with usable gross_pay — for Form XIV Employment Card and similar. */
export function filterPayrollRowsWithGrossPay(rows) {
  if (!Array.isArray(rows) || rows.length === 0) return [];
  return rows
    .filter((row) => row && typeof row === 'object' && row.fetch_error !== true)
    .map((row) => flattenPayrollEarningColumns(row))
    .filter(payrollRowHasGrossPayAmount);
}

export function flattenPayrollEarningColumns(row) {
  if (!row || typeof row !== 'object') return row;
  const employee = normalizePayrollEmployee(row);
  const earnings = getEarningsArray(row);
  const deductions = getDeductionsArray(row);
  const taxes = getTaxesArray(row);
  const componentColumns = {};
  earnings.forEach((item) => {
    const name = payrollEarningName(item);
    const type = payrollEarningType(item);
    if (!name && !type) return;
    const amount = getPayrollLineAmount(item);
    if (!Number.isFinite(amount)) return;
    const key = normalizeKey(name);
    if (key) componentColumns[key] = amount;
    if (type) componentColumns[type] = amount;
  });
  indexPayrollLineItems(deductions, componentColumns);
  indexPayrollLineItems(taxes, componentColumns);

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
    net_pay: (() => {
      const grossVal = coalesceAmount(
        pickScalarAmount(row, ['gross_pay', 'Gross Pay', 'grossPay', 'total_earnings', 'monthly_gross_amount']),
        pickAmountByPatterns(row, [/^gross_pay$/, /^total_earnings$/])
      );
      const dedVal = coalesceAmount(
        pickScalarAmount(row, [
          'total_deductions',
          'Total Deductions',
          'totalDeductions',
          'total_employee_deductions',
          'total_deduction',
        ]),
        pickAmountByPatterns(row, [/^total_deductions?$/, /^total_employee_deductions$/])
      );
      const computedNet =
        Number.isFinite(grossVal) && Number.isFinite(dedVal)
          ? Math.max(0, grossVal - dedVal)
          : '';
      return coalesceAmount(
        pickScalarAmount(row, [
          'net_pay',
          'Net Pay',
          'netPay',
          'monthly_salary',
          'MonthlySalary',
          'net_wages',
          'Net Wages',
          'take_home_pay',
          'employee_net_pay',
        ]),
        pickAmountByPatterns(row, [/^net_pay$/, /^monthly_salary$/, /^net_wages$/, /^take_home/, /^employee_net/]),
        computedNet
      );
    })(),
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
    total_benefits: coalesceAmount(
      pickScalarAmount(row, ['total_benefits', 'Total Benefits', 'totalBenefits']),
      pickAmountByPatterns(row, [/^total_benefits$/])
    ),
    total_taxes: coalesceAmount(
      pickScalarAmount(row, ['total_taxes', 'Total Taxes', 'totalTaxes']),
      pickAmountByPatterns(row, [/^total_taxes$/])
    ),
    epf_contribution: coalesceAmount(
      componentColumns.epf_contribution,
      pickScalarAmount(row, ['epf_contribution', 'EPF Contribution', 'epf', 'EPF']),
      pickAmountByPatterns(row, [/^epf_contribution$/, /^epf$/]),
      findPayrollComponentAmount(
        deductions,
        (type, name) =>
          type === 'epf_contribution' ||
          type === 'epf' ||
          type === 'pf' ||
          name.includes('epf contribution') ||
          name.includes('provident fund') ||
          (name.includes('epf') && !name.includes('employer'))
      )
    ),
    professional_tax: coalesceAmount(
      componentColumns.professional_tax,
      pickScalarAmount(row, ['professional_tax', 'Professional Tax', 'pt', 'PT']),
      pickAmountByPatterns(row, [/^professional_tax$/, /^pt$/]),
      findPayrollComponentAmount(
        taxes,
        (type, name) => type === 'professional_tax' || type === 'pt' || name.includes('professional tax')
      ),
      findPayrollComponentAmount(
        deductions,
        (type, name) => type === 'professional_tax' || name.includes('professional tax')
      )
    ),
    payment_mode: pickTextScalar(row, ['payment_mode', 'Payment Mode', 'paymentMode'], [/^payment_mode$/]),
  };
}

export function mergePayrollRunEmployeePayload(data) {
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

export function unwrapSalaryEmployeePayload(body) {
  if (!body || typeof body !== 'object') return body;
  return mergePayrollRunEmployeePayload(body);
}

export const PAYROLL_PREFERRED_COLUMNS = [
  'employee_id',
  'employee_number',
  'employee_name',
  'full_name',
  'payment_status',
  'paid_days',
  'lop_days',
  'earned_basic',
  'basic',
  'hra',
  'hra_fbp',
  'other_allowance',
  'dearness_allowance',
  'conveyance_allowance',
  'overtime',
  'gross_pay',
  'net_pay',
  'total_earnings',
  'total_deductions',
  'total_benefits',
  'total_taxes',
  'total_donations',
  'payment_mode',
  'epf_contribution',
  'professional_tax',
];

const ALWAYS_SHOW_COLUMNS = new Set([
  'earned_basic',
  'basic',
  'hra',
  'hra_fbp',
  'other_allowance',
  'dearness_allowance',
  'conveyance_allowance',
  'overtime',
]);

export function collectPayrollColumnKeys(records) {
  const keySet = new Set();
  records.forEach((row) => {
    if (row && typeof row === 'object') {
      Object.keys(row).forEach((key) => {
        if (!/^_|^\./.test(key)) keySet.add(key);
      });
    }
  });
  const preferred = PAYROLL_PREFERRED_COLUMNS.filter(
    (key) => keySet.has(key) || ALWAYS_SHOW_COLUMNS.has(key)
  );
  const preferredSet = new Set(preferred);
  const rest = [...keySet].filter((key) => !preferredSet.has(key)).sort();
  return [...preferred, ...rest];
}