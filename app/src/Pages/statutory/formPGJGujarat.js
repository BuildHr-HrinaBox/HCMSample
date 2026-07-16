/** Gujarat Form P — wage / attendance register (Working hours, days 1–31, wage summary columns). */

import * as XLSX from 'xlsx';
import {
  flattenPayrollEarningColumns,
  readPayrollScalar,
  readPayrollTextScalar,
} from '../../utils/payrollEarnings';

export function formPGJGujaratHeaderNorm(txt) {
  return String(txt || '')
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

const stripLeadingNumber = (s) =>
  String(s || '')
    .replace(/^\(?\d+\)?\s*[\.\)]?\s*/i, '')
    .trim();

function normHeaderLabel(h) {
  return formPGJGujaratHeaderNorm(stripLeadingNumber(h))
    .replace(/[^a-z0-9 ]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

export function isFormPGJGujaratContext(
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

  if (/\bform\s*q\b|\bform_q\b|form[\s._-]*q[\s._-]*gj/.test(parts)) return false;
  const hasGujarat = /gujarat|\b_gj\b|form[\s._-]*p[\s._-]*gj|form_p_gj/.test(parts);
  const hasFormP = /\bform[\s._-]*p\b/.test(parts);
  if (hasGujarat && hasFormP) return true;
  if (/\bform_p_gj\b/.test(parts)) return true;
  const joined = parts;
  return (
    /sr\.?\s*no/.test(joined) &&
    /date\s+of\s+month/.test(joined) &&
    (/working\s+hours/.test(joined) || /total\s+days\s+worked/.test(joined) || /name\s+of\s+the\s+worker/.test(joined))
  );
}

export function isFormPGJGujaratTableLayoutFormHeader(formHeader) {
  return !!formHeader?.formPGJGujaratTableLayout;
}

const FORM_PGJ_WAGE_TAIL_HEADERS = [
  'Total Days Worked',
  'Minimum Rate of Wages Rs.',
  'Total Production in case of piece rate Rs.',
  'Actual Wages Paid Rs.',
  'House Rent Allowance Rs.',
  'Dearness Allowance Rs.',
  'Gross Amount Payable Rs.',
  'Total hours of Overtime during the month',
  'Overtime Earnings Rs.',
  'Provident Fund Rs.',
  'Family Pension Rs.',
  'ESI Contribution Rs.',
  'Professional Tax Rs.',
  'Income Tax Rs.',
  'Loan & Interest Rs.',
  'Advances Rs.',
  'Other Deductions Rs.',
  'Total Deduction Rs.',
  'Net Payable Rs.',
  'Date of Payment',
  'Signature/Thumb Impression',
];

/** Full modal / export column list — worker cols, hours, days 1–31, wage summary. */
export function buildFormPGJGujaratCanonicalHeaders(daysInMonth = 31) {
  const headers = [
    'Sr. No.',
    'Name of the Worker',
    "Father's / Husband's Name",
    'Designation',
    'Date of entry into service',
    'Wage Rate',
    'Working hours_From',
    'Working hours_To',
  ];
  const days = Math.min(Math.max(Number(daysInMonth) || 31, 28), 31);
  for (let d = 1; d <= days; d += 1) {
    headers.push(`Date of Month_${d}`);
  }
  FORM_PGJ_WAGE_TAIL_HEADERS.forEach((h) => headers.push(h));
  return headers;
}

export function buildFormPGJSubColumnsFromHeaders(headers) {
  const subs = {};
  (headers || []).forEach((h) => {
    const m = String(h || '').match(/^(.+)_([\s\S]+)$/);
    if (!m) return;
    const parent = m[1].trim();
    const sub = m[2].trim();
    if (!subs[parent]) subs[parent] = [];
    if (!subs[parent].includes(sub)) subs[parent].push(sub);
  });
  return subs;
}

export function resolveFormPGJGujaratTableHeaders(_tableHeaders, daysInMonth = 31) {
  return buildFormPGJGujaratCanonicalHeaders(daysInMonth);
}

const isEmployerMetaLabel = (txt) => {
  const t = formPGJGujaratHeaderNorm(txt);
  return (
    /name\s+of\s+the\s+employer/.test(t) ||
    /^month\s*:?$/.test(t) ||
    /name\s+of\s+the\s+establishment/.test(t)
  );
};

/** Excel section index (1–9) under worker parents — not a data sub-column. */
const isFormPGJSectionMarkerSub = (sub, main) => {
  const subText = String(sub || '').trim();
  if (!/^\d{1,2}$/.test(subText)) return false;
  const n = parseInt(subText, 10);
  if (n < 1 || n > 9) return false;
  const ml = formPGJGujaratHeaderNorm(main);
  if (ml.includes('date') && ml.includes('month')) return false;
  return (
    /sr\.?\s*no/.test(ml) ||
    /name\s+of\s+the\s+worker/.test(ml) ||
    /father|husband/.test(ml) ||
    /designation/.test(ml) ||
    /entry\s+into\s+service/.test(ml) ||
    /wage\s+rate/.test(ml) ||
    /working\s+hours/.test(ml)
  );
};

const isFormPGJBlankOrMetaTableHeader = (headerKey) => {
  const t = String(headerKey || '').trim();
  if (!t || t === ':') return true;
  if (/^column\s*\d+$/i.test(t)) return true;
  const m = t.match(/^(\d{1,2})$/);
  if (m) {
    const n = parseInt(m[1], 10);
    return n >= 1 && n <= 9;
  }
  const tl = t.toLowerCase().replace(/\s+/g, ' ');
  return (
    /name\s+of\s+the\s+employer/.test(tl) ||
    /name\s+of\s+the\s+establishment/.test(tl) ||
    /^month\s*:?$/.test(tl)
  );
};

const normalizeFormPGJHeaderKey = (headerKey) => {
  const h = String(headerKey || '').trim();
  if (!h) return h;
  const m = h.match(/^(.+)_([\s\S]+)$/);
  if (!m) return h;
  const parent = m[1].trim();
  const sub = String(m[2]).trim();
  if (/^(from|to)$/i.test(sub)) return h;
  if (isFormPGJSectionMarkerSub(sub, parent)) return parent;
  if (/^date\s+of\s+month$/i.test(parent) && /^\d{1,2}$/.test(sub)) {
    return `Date of Month_${sub}`;
  }
  return h;
};

export const stripAndNormalizeFormPGJTableHeaders = (headers) => {
  if (!Array.isArray(headers) || headers.length === 0) return [];
  const normalized = headers.map(normalizeFormPGJHeaderKey);
  let start = 0;
  while (start < normalized.length && isFormPGJBlankOrMetaTableHeader(normalized[start])) {
    start += 1;
  }
  const seen = new Set();
  const out = [];
  for (let i = start; i < normalized.length; i += 1) {
    const h = normalized[i];
    if (isFormPGJBlankOrMetaTableHeader(h)) continue;
    const key = h.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    out.push(h);
  }
  return out;
};

export function remapFormPGJRowsToHeaders(rows, oldHeaders, newHeaders) {
  if (!Array.isArray(rows) || !Array.isArray(newHeaders) || newHeaders.length === 0) {
    return rows || [];
  }
  const oldList = Array.isArray(oldHeaders) ? oldHeaders : [];
  const normOld = (h) => formPGJGujaratHeaderNorm(stripLeadingNumber(h));
  return rows.map((row, rowIndex) => {
    const next = {};
    newHeaders.forEach((newH) => {
      if (row && row[newH] != null && String(row[newH]).trim() !== '') {
        next[newH] = row[newH];
        return;
      }
      const newNorm = normOld(newH);
      for (let i = 0; i < oldList.length; i += 1) {
        const oldH = oldList[i];
        if (normOld(oldH) !== newNorm) continue;
        const v = row?.[oldH];
        if (v != null && String(v).trim() !== '') {
          next[newH] = v;
          return;
        }
      }
      if (/^sr\.?\s*no/i.test(newH)) {
        next[newH] = String(rowIndex + 1);
      } else {
        next[newH] = '';
      }
    });
    if (row && typeof row === 'object') {
      Object.entries(row).forEach(([k, v]) => {
        if (String(k).startsWith('__') && next[k] == null) next[k] = v;
      });
    }
    return next;
  });
}

/**
 * Parse Form P GJ Excel header rows (parent + From/To/day subs + wage tail columns).
 */
export function rebuildFormPGJGujaratTableHeadersFromSheet({
  worksheet,
  jsonData,
  effectiveSheetCols,
  getMergedAwareCellText,
  merges,
  daysInMonth = 31,
}) {
  const norm = (txt) =>
    String(txt || '')
      .replace(/\r?\n/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  const normLower = (txt) => norm(txt).toLowerCase();

  const getMergeSpanEndCol = (r, c) => {
    for (let i = 0; i < merges.length; i += 1) {
      const m = merges[i];
      if (!m?.s || !m?.e) continue;
      if (r >= m.s.r && r <= m.e.r && c >= m.s.c && c <= m.e.c) {
        return m.e.c + 1;
      }
    }
    return c + 1;
  };

  let headerRowIndex = -1;
  for (let r = 0; r < Math.min(30, jsonData.length); r += 1) {
    const parts = [];
    for (let c = 0; c < Math.max(effectiveSheetCols, 20); c += 1) {
      const t = normLower(getMergedAwareCellText(r, c));
      if (t) parts.push(t);
    }
    const joined = parts.join(' ');
    if (/sr\.?\s*no/.test(joined) && /name\s+of\s+the\s+worker/.test(joined)) {
      headerRowIndex = r;
      break;
    }
    if (/working\s+hours/.test(joined) && /date\s+of\s+month/.test(joined)) {
      headerRowIndex = r;
      break;
    }
  }
  if (headerRowIndex < 0) return null;

  let startCol = -1;
  for (let c = 0; c < Math.max(effectiveSheetCols, 12); c += 1) {
    const t = normLower(getMergedAwareCellText(headerRowIndex, c));
    if (/^sr\.?\s*no\.?$/.test(t) || t === 'sr no' || t.startsWith('sr. no')) {
      startCol = c;
      break;
    }
  }
  if (startCol < 0) startCol = 0;

  const subRowIndex = headerRowIndex + 1;
  const flatHeaders = [];
  const subColumnsData = {};
  const mainHeaders = [];

  let c = startCol;
  const maxCols = Math.max(effectiveSheetCols, (jsonData[headerRowIndex] || []).length, 70);
  while (c < maxCols) {
    const mergeEnd = getMergeSpanEndCol(headerRowIndex, c);
    let endC = mergeEnd;
    const main = norm(getMergedAwareCellText(headerRowIndex, c));
    if (!main || isEmployerMetaLabel(main) || isFormPGJBlankOrMetaTableHeader(main)) {
      c = Math.max(c + 1, endC);
      continue;
    }
    while (endC < maxCols) {
      const nextMain = norm(getMergedAwareCellText(headerRowIndex, endC));
      if (!nextMain) break;
      if (normLower(nextMain) !== normLower(main)) break;
      endC = Math.max(endC, getMergeSpanEndCol(headerRowIndex, endC));
    }

    const mainLower = normLower(main);
    const subs = [];
    for (let sc = c; sc < endC; sc += 1) {
      const sub = norm(getMergedAwareCellText(subRowIndex, sc));
      if (!sub || sub.toLowerCase() === main.toLowerCase()) continue;
      if (sub.endsWith(':') && sub.length < 40) continue;
      if (isFormPGJSectionMarkerSub(sub, main)) continue;
      subs.push(sub);
    }

    mainHeaders.push(main);
    if (!subColumnsData[main]) subColumnsData[main] = [];

    if (/date\s+of\s+month/i.test(mainLower)) {
      const dayCount = Math.min(Math.max(Number(daysInMonth) || 31, 28), 31);
      for (let d = 1; d <= dayCount; d += 1) {
        const key = `Date of Month_${d}`;
        flatHeaders.push(key);
        subColumnsData[main].push(String(d));
      }
    } else if (/working\s+hours/i.test(mainLower)) {
      const fromTo = subs.filter((s) => /^(from|to)$/i.test(s));
      if (fromTo.length >= 2) {
        fromTo.forEach((sub) => {
          const key = `Working hours_${sub}`;
          flatHeaders.push(key);
          subColumnsData[main].push(sub);
        });
      } else {
        flatHeaders.push('Working hours_From', 'Working hours_To');
        subColumnsData[main].push('From', 'To');
      }
    } else if (subs.length === 0) {
      flatHeaders.push(main);
    } else {
      const allNumericDays = subs.every((s) => /^\d{1,2}$/.test(s));
      if (allNumericDays && subs.length >= 20) {
        subs.forEach((sub) => {
          const key = `Date of Month_${sub}`;
          flatHeaders.push(key);
          subColumnsData[main].push(sub);
        });
      } else {
        subs.forEach((sub) => {
          const key = `${main}_${sub}`;
          flatHeaders.push(key);
          subColumnsData[main].push(sub);
        });
      }
    }
    c = endC;
  }

  const canonical = buildFormPGJGujaratCanonicalHeaders(daysInMonth);
  const expandedHeaders = stripAndNormalizeFormPGJTableHeaders(
    flatHeaders.length >= 10 ? flatHeaders : canonical
  );
  const finalHeaders =
    expandedHeaders.length >= canonical.length - 4
      ? expandedHeaders
      : stripAndNormalizeFormPGJTableHeaders(canonical);

  const headerFields = [];
  for (let r = 0; r < headerRowIndex; r += 1) {
    for (let col = 0; col < Math.max(20, effectiveSheetCols); col += 1) {
      const raw = norm(getMergedAwareCellText(r, col));
      if (!raw) continue;
      if (/^month\s*:?$/i.test(raw)) {
        let val = '';
        for (let nc = col + 1; nc < Math.min(col + 8, maxCols); nc += 1) {
          const v = norm(getMergedAwareCellText(r, nc));
          if (v && !/^month\s*:?$/i.test(v)) {
            val = v;
            break;
          }
        }
        headerFields.push({ label: 'Month', value: val, key: 'form_p_gj_month' });
      }
    }
  }

  return {
    headerRowIndex,
    startCol,
    headers: mainHeaders,
    expandedHeaders: finalHeaders,
    subColumnsData,
    headerFields: headerFields.length
      ? headerFields
      : [{ label: 'Month', value: '', key: 'form_p_gj_month' }],
  };
}

export function resolveFormPGJTableHeadersForWorkbook(workbook, daysInMonth = 31) {
  const canonical = stripAndNormalizeFormPGJTableHeaders(
    buildFormPGJGujaratCanonicalHeaders(daysInMonth)
  );
  if (!workbook?.SheetNames?.[0]) {
    return {
      headers: canonical,
      subColumns: buildFormPGJSubColumnsFromHeaders(canonical),
      headerRowIndex: 11,
      dataStartIndex: 14,
      headerFields: [{ label: 'Month', value: '', key: 'form_p_gj_month' }],
    };
  }
  const worksheet = workbook.Sheets[workbook.SheetNames[0]];
  const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
  const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1:A1');
  const maxSheetCols = range.e && typeof range.e.c === 'number' ? range.e.c + 1 : 0;
  const merges = worksheet['!merges'] || [];
  let mergeMaxCol = 0;
  merges.forEach((m) => {
    if (m?.e && typeof m.e.c === 'number') mergeMaxCol = Math.max(mergeMaxCol, m.e.c + 1);
  });
  const effectiveSheetCols = Math.max(maxSheetCols, mergeMaxCol);
  const getRawCellText = (r, c) => {
    if (r < 0 || c < 0) return '';
    const ref = XLSX.utils.encode_cell({ r, c });
    const cell = worksheet[ref];
    if (!cell || cell.v == null) return '';
    return String(cell.v).trim();
  };
  const getMergedAwareCellText = (r, c) => {
    const direct = getRawCellText(r, c);
    if (direct) return direct;
    for (let i = 0; i < merges.length; i += 1) {
      const m = merges[i];
      if (!m?.s || !m?.e) continue;
      if (r >= m.s.r && r <= m.e.r && c >= m.s.c && c <= m.e.c) {
        return getRawCellText(m.s.r, m.s.c);
      }
    }
    return '';
  };
  const rebuild = rebuildFormPGJGujaratTableHeadersFromSheet({
    worksheet,
    jsonData,
    effectiveSheetCols,
    getMergedAwareCellText,
    merges,
    daysInMonth,
  });
  const headers = rebuild?.expandedHeaders?.length >= 10 ? rebuild.expandedHeaders : canonical;
  let dataStartIndex = (rebuild?.headerRowIndex ?? 11) + 2;
  if (rebuild?.headerRowIndex != null && rebuild.headerRowIndex + 2 < jsonData.length) {
    let seqHits = 0;
    const sectionRow = rebuild.headerRowIndex + 2;
    const startCol = rebuild.startCol ?? 0;
    for (let hi = 0; hi < Math.min(9, headers.length); hi += 1) {
      const t = String(getMergedAwareCellText(sectionRow, startCol + hi) || '')
        .replace(/\s+/g, '')
        .trim();
      if (t === String(hi + 1)) seqHits += 1;
    }
    if (seqHits >= 4) dataStartIndex = sectionRow + 1;
  }
  return {
    headers: stripAndNormalizeFormPGJTableHeaders(headers),
    subColumns:
      rebuild?.subColumnsData &&
      Object.keys(rebuild.subColumnsData).some((k) => (rebuild.subColumnsData[k] || []).length > 0)
        ? rebuild.subColumnsData
        : buildFormPGJSubColumnsFromHeaders(headers),
    headerRowIndex: rebuild?.headerRowIndex ?? 11,
    dataStartIndex,
    headerFields: rebuild?.headerFields || [{ label: 'Month', value: '', key: 'form_p_gj_month' }],
  };
}

function unwrapEmployeeRecord(emp) {
  return emp?.Employee || emp?.employee || emp;
}

function pickScalarEmployeeValue(raw) {
  if (raw == null || raw === '') return '';
  if (typeof raw === 'object') {
    return String(
      raw.display_value ??
        raw.displayValue ??
        raw.name ??
        raw.Name ??
        raw.value ??
        raw.text ??
        ''
    ).trim();
  }
  return String(raw).trim();
}

function collectEmployeeSources(emp) {
  const sources = [];
  const seen = new Set();
  const push = (value) => {
    if (!value || typeof value !== 'object' || seen.has(value)) return;
    seen.add(value);
    sources.push(value);
  };
  push(emp);
  if (emp && typeof emp === 'object') {
    push(emp.Employee);
    push(emp.employee);
    push(unwrapEmployeeRecord(emp));
  }
  return sources;
}

function readScalarFromEmployee(emp, keys) {
  const sources = collectEmployeeSources(emp);
  for (let i = 0; i < sources.length; i += 1) {
    const src = sources[i];
    for (let k = 0; k < keys.length; k += 1) {
      const val = pickScalarEmployeeValue(src[keys[k]]);
      if (val && val !== '-') return val;
    }
    for (const [k, v] of Object.entries(src)) {
      const kn = String(k || '')
        .toLowerCase()
        .replace(/[^a-z0-9]/g, '');
      if (keys.some((key) => kn === String(key).toLowerCase().replace(/[^a-z0-9]/g, ''))) {
        const val = pickScalarEmployeeValue(v);
        if (val && val !== '-') return val;
      }
    }
  }
  return '';
}

export function readFormPGJShiftStartFromEmployee(emp) {
  return readScalarFromEmployee(emp, [
    'ShiftStartTime',
    'Shift_Start_Time',
    'shiftStartTime',
    'Shift Start Time',
  ]);
}

export function readFormPGJShiftEndFromEmployee(emp) {
  return readScalarFromEmployee(emp, [
    'ShiftEndTime',
    'Shift_End_Time',
    'shiftEndTime',
    'Shift End Time',
  ]);
}

export function readFormPGJShiftNameFromEmployee(emp) {
  return readScalarFromEmployee(emp, [
    'ShiftName',
    'Shift_Name',
    'shiftName',
    'Shift Name',
    'Shift',
  ]);
}

export function isFormPGJWorkingHoursFromHeader(header) {
  const n = normHeaderLabel(header);
  return n === 'working hours from' || (n.includes('working hours') && n.includes('from'));
}

export function isFormPGJWorkingHoursToHeader(header) {
  const n = normHeaderLabel(header);
  return n === 'working hours to' || (n.includes('working hours') && n.includes('to'));
}

export function isFormPGJDateOfMonthDayHeader(header) {
  return /^date\s+of\s+month_\d{1,2}$/i.test(String(header || '').trim());
}

export function isFormPGJWageTailHeader(header) {
  const n = normHeaderLabel(header);
  return FORM_PGJ_WAGE_TAIL_HEADERS.some((h) => normHeaderLabel(h) === n);
}

export function isFormPGJSignatureHeader(header) {
  const n = normHeaderLabel(header);
  return n.includes('signature') || n.includes('thumb impression');
}

function parsePayrollNumber(value) {
  const n = Number(String(value ?? '').replace(/[,₹]/g, '').trim());
  return Number.isFinite(n) ? n : NaN;
}

function readPayrollDeductionScalar(flat, payrollRow, keys, patterns = []) {
  const val = readPayrollScalar({ ...flat, ...payrollRow }, keys, patterns);
  return val !== '' && val != null ? val : '';
}

export function resolveFormPGJGujaratPayrollFields(payrollRow, helpers = {}) {
  const empty = {
    paidDays: '',
    basic: '',
    hra: '',
    grossPay: '',
    overtimeHours: '',
    esi: '',
    professionalTax: '',
    incomeTax: '',
    totalDeduction: '',
    netPay: '',
    paymentDate: '',
  };
  if (!payrollRow || payrollRow.fetch_error) return empty;

  const flat = flattenPayrollEarningColumns(payrollRow);
  const merged = { ...flat, ...payrollRow };

  const paidDays = readPayrollScalar(
    merged,
    ['paid_days', 'Paid Days', 'days_worked', 'Days Worked', 'paidDays', 'no_of_days_worked'],
    [/^paid_days$/, /paiddays/, /daysworked/]
  );

  const basic =
    readPayrollScalar(
      merged,
      ['basic', 'Basic', 'earned_basic', 'Earned Basic', 'basic_pay', 'Basic Earnings'],
      [/^basic$/, /^earned_basic$/, /^basic_pay$/]
    ) || (flat.basic != null && flat.basic !== '' ? flat.basic : '');

  const hra =
    readPayrollScalar(
      merged,
      ['hra_fbp', 'HRA FBP', 'hra (fbp)', 'hra_fp', 'HRA (FBP)'],
      [/^hra_fbp$/, /^hra_fp$/]
    ) ||
    readPayrollScalar(
      merged,
      ['hra', 'HRA', 'house_rent_allowance', 'House Rent Allowance'],
      [/^hra$/, /house.*rent/]
    ) ||
    (flat.hra_fbp != null && flat.hra_fbp !== '' ? flat.hra_fbp : flat.hra || '');

  const grossPay = readPayrollScalar(
    merged,
    ['gross_pay', 'Gross Pay', 'grossPay', 'total_earnings', 'Gross Amount Payable'],
    [/^gross_pay$/, /^total_earnings$/]
  );

  let overtimeHours = readPayrollTextScalar(
    merged,
    [
      'total_ot_hours',
      'Total OT hours',
      'total overtime hours',
      'overtime_hours',
      'Overtime Hours',
      'total_overtime_hours',
      'ot_hours',
    ],
    [/total.*ot.*hour/i, /^total_ot_hours$/, /^overtime_hours$/, /^ot_hours$/]
  );
  if (overtimeHours === '') {
    overtimeHours = readPayrollScalar(
      merged,
      ['total_ot_hours', 'overtime_hours', 'total_overtime_hours', 'ot_hours'],
      [/total.*ot.*hour/i, /^total_ot_hours$/, /^overtime_hours$/, /^ot_hours$/]
    );
  }

  const esi = readPayrollDeductionScalar(
    flat,
    payrollRow,
    ['esi', 'ESI', 'esic', 'ESIC', 'esi_contribution', 'ESI Contribution'],
    [/^esi/, /employee state insurance/]
  );

  const professionalTax = readPayrollDeductionScalar(
    flat,
    payrollRow,
    ['professional_tax', 'Professional Tax', 'pt', 'PT'],
    [/^professional_tax$/, /^pt$/]
  );

  const incomeTax =
    readPayrollScalar(merged, ['income_tax', 'Income Tax'], [/^income_tax$/]) ||
    readPayrollDeductionScalar(
      flat,
      payrollRow,
      ['income_tax', 'Income Tax', 'tds', 'TDS', 'tax_deducted_at_source'],
      [/income.*tax/i, /^tds$/]
    );

  const totalDeduction = readPayrollScalar(
    merged,
    ['total_deductions', 'Total Deductions', 'totalDeductions', 'total_employee_deductions', 'total_deduction'],
    [/^total_deductions?$/]
  );

  const netPay = readPayrollScalar(
    merged,
    ['net_pay', 'Net Pay', 'netPay', 'monthly_salary', 'Net Payable'],
    [/^net_pay$/]
  );

  const payDateRaw =
    readPayrollTextScalar(
      merged,
      ['pay_date', 'Pay Date', 'payment_date', 'Payment Date', 'paid_date', 'date_of_payment', 'Date of Payment'],
      [/^pay_date$/, /payment.*date/i, /^paid_date$/]
    ) ||
    String(payrollRow?.pay_date ?? flat?.pay_date ?? helpers.payDate ?? '').trim();

  const paymentDate =
    typeof helpers.formatStatutoryDateDisplay === 'function'
      ? helpers.formatStatutoryDateDisplay(payDateRaw) || payDateRaw
      : payDateRaw;

  return {
    paidDays,
    basic,
    hra,
    grossPay,
    overtimeHours,
    esi,
    professionalTax,
    incomeTax,
    totalDeduction,
    netPay,
    paymentDate,
  };
}

function formatCellValue(value, { allowZero = false } = {}) {
  if (value == null || value === '') return '';
  const num = parsePayrollNumber(value);
  if (Number.isFinite(num) && (allowZero || num !== 0)) return num;
  return String(value).trim();
}

function matchFormPGJWageTailBucket(header) {
  const n = normHeaderLabel(header);
  if (n.includes('total days worked')) return 'paidDays';
  if (n.includes('minimum rate of wages')) return 'skip';
  if (n.includes('total production') && n.includes('piece rate')) return 'skip';
  if (n.includes('actual wages paid')) return 'basic';
  if (n.includes('house rent allowance')) return 'hra';
  if (n.includes('dearness allowance')) return 'skip';
  if (n.includes('gross amount payable')) return 'grossPay';
  if (n.includes('total hours of overtime')) return 'overtimeHours';
  if (n.includes('overtime earnings')) return 'skip';
  if (n.includes('provident fund')) return 'skip';
  if (n.includes('family pension')) return 'skip';
  if (n.includes('esi contribution') || (n.includes('esi') && n.includes('contribution'))) return 'esi';
  if (n.includes('professional tax')) return 'professionalTax';
  if (n.includes('income tax')) return 'incomeTax';
  if (n.includes('loan') && n.includes('interest')) return 'skip';
  if (n.includes('advances')) return 'skip';
  if (n.includes('other deductions')) return 'skip';
  if (n.includes('total deduction')) return 'totalDeduction';
  if (n.includes('net payable')) return 'netPay';
  if (n.includes('date of payment')) return 'paymentDate';
  return '';
}

function unwrapPGJEmployee(empItem) {
  return empItem?.Employee || empItem?.employee || empItem || {};
}

/** Build People match candidates — unwrap nested Employee + row lookup metadata from grid. */
export function buildFormPGJEmployeePayrollCandidates(empItem, row = null) {
  const emp = unwrapPGJEmployee(empItem);
  const candidates = [];
  const seen = new Set();
  const push = (value) => {
    if (!value || typeof value !== 'object' || seen.has(value)) return;
    seen.add(value);
    candidates.push(value);
  };
  push(empItem);
  push(emp);
  const lookupId = String(row?.__employeeLookupId || '').trim();
  if (lookupId) {
    push({
      ...emp,
      Zoho_ID: lookupId,
      EmployeeID: lookupId,
      employee_number: lookupId,
      Employee_Number: lookupId,
    });
  }
  const lookupName = String(row?.__employeeLookupName || '').trim();
  if (lookupName) {
    const parts = lookupName.split(/\s+/).filter(Boolean);
    push({
      ...emp,
      DisplayName: lookupName,
      EmployeeName: lookupName,
      FirstName: parts[0] || '',
      LastName: parts.slice(1).join(' ') || '',
    });
  }
  return candidates;
}

export function applyFormPGJGujaratPayrollToRow(row, payrollRow, headers, helpers = {}) {
  const hdrs = Array.isArray(headers) && headers.length > 0 ? headers : buildFormPGJGujaratCanonicalHeaders();
  const out = row && typeof row === 'object' ? row : {};
  const {
    sanitizeValue = (v) => String(v ?? '').trim(),
    formatStatutoryDateDisplay = (v) => String(v || '').trim(),
    payDate = '',
    overwrite = true,
  } = helpers;
  const payroll = resolveFormPGJGujaratPayrollFields(
    payrollRow && !payrollRow.fetch_error ? payrollRow : null,
    { formatStatutoryDateDisplay, payDate }
  );
  const cellIsEmpty = (header) => {
    const v = String(out[header] || '').trim();
    if (!v) return true;
    const s = v.toLowerCase();
    return /^enter\b/.test(s) || s.includes('enter ');
  };
  const setCell = (header, value, opts = {}) => {
    if (!header || value == null || value === '') return false;
    if (!overwrite && !cellIsEmpty(header)) return false;
    out[header] = sanitizeValue(formatCellValue(value, opts));
    return true;
  };
  let wrote = false;
  hdrs.forEach((header) => {
    if (isFormPGJSignatureHeader(header)) return;
    const bucket = matchFormPGJWageTailBucket(header);
    if (!bucket || bucket === 'skip') return;
    if (Object.prototype.hasOwnProperty.call(payroll, bucket)) {
      if (setCell(header, payroll[bucket], { allowZero: true })) wrote = true;
    }
  });
  return wrote;
}

export function enrichFormPGJGujaratPayrollRows(mappedData, employees, headers, helpers = {}) {
  const hdrs = Array.isArray(headers) && headers.length > 0 ? headers : buildFormPGJGujaratCanonicalHeaders();
  const {
    resolvePayrollRow = null,
    sanitizeValue = (v) => String(v ?? '').trim(),
    formatStatutoryDateDisplay = (v) => String(v || '').trim(),
    payDate = '',
    overwrite = true,
    rowIndexOffset = 0,
  } = helpers;
  if (!Array.isArray(mappedData) || mappedData.length === 0 || typeof resolvePayrollRow !== 'function') {
    return 0;
  }
  let hits = 0;
  mappedData.forEach((row, rowIndex) => {
    const globalRowIndex = rowIndexOffset + rowIndex;
    const empItem = employees?.[rowIndex];
    let payrollRow = null;
    if (typeof resolvePayrollRow === 'function') {
      payrollRow =
        resolvePayrollRow(empItem, row, globalRowIndex) ||
        resolvePayrollRow(unwrapPGJEmployee(empItem), row, globalRowIndex);
      if (!payrollRow || payrollRow.fetch_error) {
        const candidates = buildFormPGJEmployeePayrollCandidates(empItem, row);
        for (let ci = 0; ci < candidates.length; ci += 1) {
          const candidate = candidates[ci];
          const hit = resolvePayrollRow(candidate, row, globalRowIndex);
          if (hit && !hit.fetch_error) {
            payrollRow = hit;
            break;
          }
        }
      }
    }
    const wrote = applyFormPGJGujaratPayrollToRow(row, payrollRow, hdrs, {
      sanitizeValue,
      formatStatutoryDateDisplay,
      payDate,
      overwrite,
    });
    if (wrote) hits += 1;
  });
  return hits;
}

export function enrichFormPGJGujaratDisplayHeader(formHeader, fileName, item, headers, sheetText = '') {
  if (!isFormPGJGujaratContext(formHeader, item, fileName, sheetText, headers)) return formHeader;
  return {
    ...(formHeader || {}),
    title: formHeader?.title || 'FORM P',
    formPGJGujaratTableLayout: true,
    fields: Array.isArray(formHeader?.fields) && formHeader.fields.length > 0
      ? formHeader.fields
      : [{ label: 'Month', value: '', key: 'form_p_gj_month' }],
  };
}
