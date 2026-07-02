import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import {
  flattenPayrollEarningColumns,
  readPayrollScalar,
  readPayrollTextScalar,
} from '../../utils/payrollEarnings';
import {
  formXIXAPHeaderNorm,
  formatWorkmanNameAndGuardian,
  isFormXIXAPWageSlipContext,
} from './formXIXAPWageSlip';

function getRowValueForHeader(row, header) {
  if (!row || !header) return '';
  const direct = row[header];
  if (direct != null && String(direct).trim() !== '') return String(direct).trim();
  const want = formXIXAPHeaderNorm(header);
  for (const [k, v] of Object.entries(row)) {
    if (formXIXAPHeaderNorm(k) === want) return String(v ?? '').trim();
  }
  return '';
}

function getKarnatakaRowValueForHeader(row, header) {
  if (!row || !header) return '';
  const direct = getRowValueForHeader(row, header);
  if (direct !== '') return direct;

  const headerMatchers = [
    isFormXIXKASexIdentificationHeader,
    isFormXIXKATokenHeader,
    isFormXIXKAWorkmanNameHeader,
    isFormXIXKADaysWorkedHeader,
    isFormXIXKARateHeader,
    isFormXIXKAUnitsHeader,
    isFormXIXKAOvertimeDatesHeader,
    isFormXIXKAOvertimeAmountHeader,
    isFormXIXKAGrossHeader,
    isFormXIXKADeductionsHeader,
    isFormXIXKANetHeader,
  ];
  const matchFn = headerMatchers.find((fn) => fn(header));
  if (!matchFn) return '';

  for (const [key, value] of Object.entries(row)) {
    if (String(key).startsWith('__')) continue;
    if (matchFn(key) && value != null && String(value).trim() !== '') {
      return String(value).trim();
    }
  }
  return '';
}

function resolveKarnatakaExportCellValue(row, header) {
  const value = getKarnatakaRowValueForHeader(row, header);
  if (value !== '') return value;
  if (isFormXIXKARateHeader(header)) return FORM_XIX_KA_RATE_DEFAULT;
  return '';
}

/** Karnataka CLRA Form XIX — tabular wage slip (Rule 78). */

export const FORM_XIX_KA_WAGE_TABLE_HEADERS = [
  'Sex and identification marks',
  'Token/Ticket No.',
  "Name and Father's/Husband's Name of the workman",
  'No. of days worked',
  'Rate of daily wages/piece - rate',
  'No. of units worked in case of piece rate',
  'Dates on which overtime worked',
  'Overtime hours and amount of overtime wages',
];

/** Excel template table band — identity fields sit in the header area, not these columns. */
export const FORM_XIX_KA_EXCEL_WAGE_HEADERS = [
  'No. of days worked',
  'Rate of daily wages/piece - rate',
  'No. of units worked in case of piece rate',
  'Dates on which overtime worked',
  'Overtime hours and amount of overtime wages',
];

export const FORM_XIX_KA_FOOTER_HEADERS = [
  'Gross wages payable',
  'Deductions, any',
  'If Actual wages paid',
];

export const FORM_XIX_KA_ALL_TABLE_HEADERS = [
  ...FORM_XIX_KA_WAGE_TABLE_HEADERS,
  ...FORM_XIX_KA_FOOTER_HEADERS,
];

/** Karnataka Form XIX — fixed text for rate column (not from payroll). */
export const FORM_XIX_KA_RATE_DEFAULT = 'Monthly Wages';

export const FORM_XIX_KA_FIELD_GROUPS = [
  { id: 'header', title: 'Wage slip — contractor details' },
];

/** Site-level header fields — workman identity columns render in the employee table below. */
export const FORM_XIX_KA_HEADER_SPECS = [
  {
    key: 'form_xix_ka_contractor',
    label: 'Name and address of contractor',
    group: 'header',
    fieldType: 'textarea',
    match: /name\s+and\s+address\s+(?:of|if)\s+contractor/i,
  },
  {
    key: 'form_xix_ka_establishment',
    label: 'Name and address of establishment in/under which contract is carried on',
    group: 'header',
    fieldType: 'textarea',
    match:
      /name\s+and\s+address\s+of\s+establishment[\s\S]{0,80}?contract\s+is\s+carried\s+on|establishment\s+in\s*\/\s*under\s+which\s+contract/i,
  },
  {
    key: 'form_xix_ka_nature_location',
    label: 'Nature and location of work',
    group: 'header',
    fieldType: 'textarea',
    match: /nature\s+(?:of\s+work\s+and\s+)?location\s+of\s+work|nature\s+and\s+location\s+of\s+work/i,
  },
  {
    key: 'form_xix_ka_principal_employer',
    label: 'Name and address of Principal Employer',
    group: 'header',
    fieldType: 'textarea',
    match: /name\s+and\s+address\s+of\s+principal\s+employer/i,
  },
  {
    key: 'form_xix_ka_period_ending',
    label: 'For the week/Fortnight/Month ending',
    group: 'header',
    fieldType: 'text',
    match: /week.*fortnight.*month\s+ending|fortnight.*month\s+ending/i,
  },
];

/** Per-employee fields written to the Excel template header area on export. */
const FORM_XIX_KA_EMPLOYEE_HEADER_SPECS = [
  {
    key: 'form_xix_ka_sex_identification',
    label: 'Sex and identification marks',
    match: /sex\s+and\s+identification\s+marks?/i,
  },
  {
    key: 'form_xix_ka_token',
    label: 'Token/Ticket No.',
    match: /token\/?\s*ticket\s+no\.?/i,
  },
  {
    key: 'form_xix_ka_workman',
    label: "Name and Father's/Husband's Name of the workman",
    match: /name\s+and\s+father.*husband.*workman|father.*husband.*name\s+of\s+the\s+workman/i,
  },
];

const FORM_XIX_KA_DEFAULT_VALUE_COL = 4;

const normHeader = (h) => formXIXAPHeaderNorm(String(h || '').replace(/^\d+[\.\)]\s*/, ''));

export function isFormXIXKarnatakaWageSlipContext(formHeader, rowItem, fileName, sheetText = '') {
  if (!isFormXIXAPWageSlipContext(formHeader, rowItem, fileName, sheetText)) return false;
  const parts = [
    rowItem?.formName,
    rowItem?.FormName,
    rowItem?.formFileName,
    rowItem?.FormFileName,
    rowItem?.state,
    rowItem?.State,
    fileName,
    formHeader?.title,
    sheetText,
  ]
    .filter((x) => x != null && String(x).trim() !== '')
    .join(' ')
    .toLowerCase();
  if (/karnataka|\(ka\)|form[\s._-]*xix[\s._-]*(?:ka|karnataka)/i.test(parts)) return true;
  const sheet = String(sheetText || '').toLowerCase();
  return (
    /sex\s+and\s+identification\s+marks?/i.test(sheet) &&
    /rate\s+of\s+daily\s+wages\/piece/i.test(sheet) &&
    /overtime\s+hours\s+and\s+amount\s+of\s+overtime/i.test(sheet) &&
    !/madhya\s+pradesh/.test(parts)
  );
}

export function isFormXIXKarnatakaTableLayoutFormHeader(formHeader) {
  return !!formHeader?.formXIXKarnatakaTableLayout;
}

export function isFormXIXKAGrossHeader(h) {
  const s = normHeader(h);
  return s.includes('gross') && s.includes('wage');
}

export function isFormXIXKADeductionsHeader(h) {
  return /deduction/.test(normHeader(h));
}

export function isFormXIXKANetHeader(h) {
  const s = normHeader(h);
  return (
    /actual\s+wages\s+paid/.test(s) ||
    (s.includes('net') && (s.includes('wage') || s.includes('paid') || s.includes('amount')))
  );
}

function headerMatchesKarnatakaFooter(a, b) {
  const na = normHeader(a);
  const nb = normHeader(b);
  if (na === nb) return true;
  if (isFormXIXKAGrossHeader(a) && isFormXIXKAGrossHeader(b)) return true;
  if (isFormXIXKADeductionsHeader(a) && isFormXIXKADeductionsHeader(b)) return true;
  if (isFormXIXKANetHeader(a) && isFormXIXKANetHeader(b)) return true;
  return false;
}

export function resolveFormXIXKarnatakaWageTableHeaders(tableHeaders) {
  const list = Array.isArray(tableHeaders) ? tableHeaders.filter(Boolean) : [];
  const hasIdentityCols =
    list.some(isFormXIXKASexIdentificationHeader) &&
    list.some(isFormXIXKATokenHeader) &&
    list.some(isFormXIXKAWorkmanNameHeader);
  if (list.length >= FORM_XIX_KA_ALL_TABLE_HEADERS.length) return list;
  if (!hasIdentityCols || list.length < FORM_XIX_KA_WAGE_TABLE_HEADERS.length) {
    return [...FORM_XIX_KA_ALL_TABLE_HEADERS];
  }
  const extras = FORM_XIX_KA_FOOTER_HEADERS.filter(
    (footerHeader) => !list.some((h) => headerMatchesKarnatakaFooter(h, footerHeader))
  );
  return extras.length > 0 ? [...list, ...extras] : list;
}

export function isFormXIXKADaysWorkedHeader(h) {
  return /days\s+worked/.test(normHeader(h));
}

export function isFormXIXKASexIdentificationHeader(h) {
  const s = normHeader(h);
  return s.includes('sex') && s.includes('identification');
}

export function isFormXIXKATokenHeader(h) {
  return /token\/?\s*ticket/.test(normHeader(h));
}

export function isFormXIXKAWorkmanNameHeader(h) {
  const s = normHeader(h);
  return (s.includes('workman') || s.includes('workmen')) && (s.includes('name') || s.includes('father') || s.includes('husband'));
}

export function isFormXIXKAIdentityTableHeader(h) {
  return (
    isFormXIXKASexIdentificationHeader(h) ||
    isFormXIXKATokenHeader(h) ||
    isFormXIXKAWorkmanNameHeader(h)
  );
}

export function isFormXIXKARateHeader(h) {
  const s = normHeader(h);
  if (isFormXIXKAGrossHeader(h) || isFormXIXKADeductionsHeader(h) || isFormXIXKANetHeader(h)) return false;
  return s.includes('rate') && (s.includes('wage') || s.includes('piece'));
}

export function isFormXIXKAUnitsHeader(h) {
  return /units\s+worked/.test(normHeader(h));
}

export function isFormXIXKAOvertimeDatesHeader(h) {
  return /dates\s+on\s+which\s+overtime/.test(normHeader(h));
}

export function isFormXIXKAOvertimeAmountHeader(h) {
  return /overtime\s+hours\s+and\s+amount/.test(normHeader(h));
}

export function isFormXIXKASkipPeopleAutofillHeader(h) {
  return (
    isFormXIXKASexIdentificationHeader(h) ||
    isFormXIXKATokenHeader(h) ||
    isFormXIXKAWorkmanNameHeader(h) ||
    isFormXIXKADaysWorkedHeader(h) ||
    isFormXIXKARateHeader(h) ||
    isFormXIXKAUnitsHeader(h) ||
    isFormXIXKAOvertimeDatesHeader(h) ||
    isFormXIXKAOvertimeAmountHeader(h) ||
    isFormXIXKAGrossHeader(h) ||
    isFormXIXKADeductionsHeader(h) ||
    isFormXIXKANetHeader(h)
  );
}

const isNarrativeBlob = (raw) => {
  const n = formXIXAPHeaderNorm(raw);
  if (!n) return false;
  return (
    /^form\s*xix\b/i.test(n) ||
    /wage\s+slip/i.test(n) ||
    /see\s+rule\s+78/i.test(n) ||
    /contract\s+labou?r/i.test(n) ||
    n.length > 160
  );
};

const isTableHeaderBlob = (raw) => {
  const n = normHeader(raw);
  return (
    isFormXIXKASexIdentificationHeader(n) ||
    isFormXIXKATokenHeader(n) ||
    isFormXIXKAWorkmanNameHeader(n) ||
    isFormXIXKADaysWorkedHeader(n) ||
    isFormXIXKARateHeader(n) ||
    isFormXIXKAUnitsHeader(n) ||
    isFormXIXKAOvertimeDatesHeader(n) ||
    isFormXIXKAOvertimeAmountHeader(n) ||
    isFormXIXKAGrossHeader(n) ||
    isFormXIXKADeductionsHeader(n) ||
    isFormXIXKANetHeader(n)
  );
};

export function isFormXIXKAPlaceholderCell(text) {
  const s = String(text ?? '').trim();
  if (!s) return true;
  if (/^[\s._\-…·]+$/u.test(s)) return true;
  if (/[.…_-]{4,}$/u.test(s) && s.replace(/[.…_\-\s]+/gu, '').length < 4) return true;
  return false;
}

const buildTemplateField = (spec, coords = {}) => ({
  label: spec.label,
  value: coords.value || '',
  key: spec.key,
  group: spec.group || 'header',
  fieldType: spec.fieldType || 'text',
  labelRow: coords.labelRow ?? null,
  labelCol: coords.labelCol ?? null,
  valueCol: coords.valueCol ?? null,
  valueRow: coords.valueRow ?? null,
});

const findHeaderLabelCell = (getMergedAwareCellText, matchRe, effectiveSheetCols, maxRows = 120) => {
  const maxC = Math.max(20, effectiveSheetCols || 0);
  for (let r = 0; r < maxRows; r += 1) {
    for (let c = 0; c < maxC; c += 1) {
      const raw = String(getMergedAwareCellText(r, c) || '').trim();
      if (!raw || isNarrativeBlob(raw) || isTableHeaderBlob(raw)) continue;
      const norm = formXIXAPHeaderNorm(raw);
      if (!matchRe.test(norm) && !matchRe.test(raw)) continue;
      if (norm.length > 120) continue;
      return { labelRow: r, labelCol: c, valueCol: FORM_XIX_KA_DEFAULT_VALUE_COL - 1, valueRow: r, value: '' };
    }
  }
  return null;
};

function buildWorkbookAccessor(workbook, hints = {}) {
  if (!workbook?.SheetNames?.length) return null;
  const preferred = String(hints.preferredSheetName || '').trim();
  const sheetName =
    preferred && workbook.Sheets?.[preferred]
      ? preferred
      : workbook.SheetNames.find((n) => {
          const ws = workbook.Sheets[n];
          const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
          const blob = rows
            .slice(0, 40)
            .map((row) => (Array.isArray(row) ? row : []).join(' '))
            .join(' ')
            .toLowerCase();
          return /wage\s+slip|form\s*xix/i.test(blob);
        }) || workbook.SheetNames[0];
  const worksheet = sheetName ? workbook.Sheets[sheetName] : null;
  if (!worksheet) return null;

  const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1:A1');
  const maxSheetCols = range.e && typeof range.e.c === 'number' ? range.e.c + 1 : 0;
  const merges = worksheet['!merges'] || [];
  let mergeMaxCol = 0;
  merges.forEach((m) => {
    if (m?.e && typeof m.e.c === 'number') mergeMaxCol = Math.max(mergeMaxCol, m.e.c + 1);
  });
  const effectiveSheetCols = Math.max(maxSheetCols, mergeMaxCol, 20);
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
        const topLeft = getRawCellText(m.s.r, m.s.c);
        if (topLeft) return topLeft;
      }
    }
    return '';
  };
  return { getMergedAwareCellText, effectiveSheetCols, sheetName };
}

function resolveKarnatakaTableLayout(getMergedAwareCellText, effectiveSheetCols) {
  const wageHdrs = FORM_XIX_KA_EXCEL_WAGE_HEADERS;
  const uiHdrs = FORM_XIX_KA_WAGE_TABLE_HEADERS;
  const maxC = Math.max(20, effectiveSheetCols || 0);
  for (let r = 0; r < 60; r += 1) {
    for (let c = 0; c < maxC; c += 1) {
      const raw = String(getMergedAwareCellText(r, c) || '').trim();
      if (!isFormXIXKADaysWorkedHeader(raw)) continue;
      const matched = uiHdrs.map((header) => {
        if (isFormXIXKAIdentityTableHeader(header)) {
          return { header, col: c - 1 };
        }
        const wageIdx = wageHdrs.findIndex((h) => normHeader(h) === normHeader(header));
        return { header, col: c + (wageIdx >= 0 ? wageIdx : 0) };
      });
      let dataStartIndex = r + 1;
      for (let dr = r + 1; dr <= r + 4; dr += 1) {
        const firstCell = String(getMergedAwareCellText(dr, c) || '').trim();
        if (/^\d+$/.test(firstCell)) {
          dataStartIndex = dr + 1;
          break;
        }
      }
      return {
        headerRowIndex: r,
        dataStartIndex,
        tableStartCol: c,
        columns: matched,
      };
    }
  }
  return {
    headerRowIndex: 12,
    dataStartIndex: 14,
    tableStartCol: 0,
    columns: uiHdrs.map((header, i) => ({ header, col: i })),
  };
}

function resolveKarnatakaFooterLayout(getMergedAwareCellText, effectiveSheetCols, dataStartIndex = 14) {
  const maxC = Math.max(20, effectiveSheetCols || 0);
  const footerHdrs = FORM_XIX_KA_FOOTER_HEADERS;
  for (let r = dataStartIndex + 1; r < dataStartIndex + 12; r += 1) {
    const matched = [];
    for (let c = 0; c < maxC; c += 1) {
      const cell = String(getMergedAwareCellText(r, c) || '').trim();
      if (!cell) continue;
      if (isFormXIXKAGrossHeader(cell)) matched.push({ header: footerHdrs[0], col: c, labelRow: r });
      else if (isFormXIXKADeductionsHeader(cell)) matched.push({ header: footerHdrs[1], col: c, labelRow: r });
      else if (isFormXIXKANetHeader(cell)) matched.push({ header: footerHdrs[2], col: c, labelRow: r });
    }
    if (matched.length >= 2) {
      return { footerRowIndex: r, footerValueRow: r + 1, footerColumns: matched };
    }
  }
  return {
    footerRowIndex: dataStartIndex + 4,
    footerValueRow: dataStartIndex + 5,
    footerColumns: [
      { header: footerHdrs[0], col: 1, labelRow: dataStartIndex + 4 },
      { header: footerHdrs[1], col: 3, labelRow: dataStartIndex + 4 },
      { header: footerHdrs[2], col: 5, labelRow: dataStartIndex + 4 },
    ],
  };
}

export function buildFormXIXKarnatakaHeaderFields(getMergedAwareCellText = null, effectiveSheetCols = 20) {
  return FORM_XIX_KA_HEADER_SPECS.map((spec) => {
    let coords = {};
    if (typeof getMergedAwareCellText === 'function' && spec.match) {
      coords = findHeaderLabelCell(getMergedAwareCellText, spec.match, effectiveSheetCols) || {};
    }
    return buildTemplateField(spec, coords);
  });
}

export function resolveFormXIXKarnatakaHeaderFieldLayout(parsed, workbook, hints = {}) {
  const formHeader = hints.formHeader || parsed?.formHeader || null;
  const item = hints.item || null;
  const fileName = hints.fileName || hints.formFileName || '';
  const sheetText = hints.sheetText || '';
  if (!isFormXIXKarnatakaWageSlipContext(formHeader, item, fileName, sheetText)) return null;

  const accessor = buildWorkbookAccessor(workbook, hints);
  const getMergedAwareCellText = accessor?.getMergedAwareCellText ?? null;
  const effectiveSheetCols = accessor?.effectiveSheetCols ?? 20;
  const headerFields = buildFormXIXKarnatakaHeaderFields(getMergedAwareCellText, effectiveSheetCols);
  const tableLayout =
    typeof getMergedAwareCellText === 'function'
      ? resolveKarnatakaTableLayout(getMergedAwareCellText, effectiveSheetCols)
      : { headerRowIndex: 12, dataStartIndex: 14, tableStartCol: 0, columns: [] };
  const footerLayout =
    typeof getMergedAwareCellText === 'function'
      ? resolveKarnatakaFooterLayout(
          getMergedAwareCellText,
          effectiveSheetCols,
          tableLayout.dataStartIndex ?? 14
        )
      : {
          footerRowIndex: 18,
          footerValueRow: 19,
          footerColumns: FORM_XIX_KA_FOOTER_HEADERS.map((header, i) => ({
            header,
            col: 1 + i * 2,
            labelRow: 18,
          })),
        };

  return {
    formHeader: {
      title: formHeader?.title || 'FORM XIX',
      subtitle: formHeader?.subtitle || 'Wage Slip (Contract Labour)',
      reference: formHeader?.reference || '(See rule 78 (1)(b))',
      formXIXAPHeaderFieldLayout: true,
      formXIXKarnatakaTableLayout: true,
      textRows: [],
      fields: headerFields,
      tableColumns: tableLayout.columns,
      footerColumns: footerLayout.footerColumns,
      footerRowIndex: footerLayout.footerRowIndex,
      footerValueRow: footerLayout.footerValueRow,
      headerRowIndex: tableLayout.headerRowIndex,
      dataStartIndex: tableLayout.dataStartIndex,
      tableStartCol: tableLayout.tableStartCol,
    },
    headers: resolveFormXIXKarnatakaWageTableHeaders(parsed?.headers || hints.tableHeaders || null),
    tableData: [],
    headerRowIndex: tableLayout.headerRowIndex,
    dataStartIndex: tableLayout.dataStartIndex,
    tableStartCol: tableLayout.tableStartCol,
    footerRowIndex: footerLayout.footerRowIndex,
    footerValueRow: footerLayout.footerValueRow,
    footerColumns: footerLayout.footerColumns,
    sheetName: accessor?.sheetName || hints.preferredSheetName || null,
  };
}

const setHeaderField = (headerData, key, value) => {
  if (!key || value == null || String(value).trim() === '') return headerData;
  const out = { ...(headerData || {}) };
  out[key] = String(value).trim();
  return out;
};

export function applyFormXIXKarnatakaAutofillFromSite(headerData, siteContext = {}, options = {}) {
  const { onlyEmpty = false } = options;
  const {
    contractorText = '',
    establishmentText = '',
    natureLocationText = '',
    principalEmployerText = '',
    periodEndingText = '',
  } = siteContext;
  let out = { ...(headerData || {}) };
  const assign = (key, value) => {
    if (onlyEmpty && String(out[key] ?? '').trim()) return;
    out = setHeaderField(out, key, value);
  };
  assign('form_xix_ka_contractor', contractorText);
  assign('form_xix_ka_establishment', establishmentText);
  assign('form_xix_ka_nature_location', natureLocationText);
  assign('form_xix_ka_principal_employer', principalEmployerText);
  assign('form_xix_ka_period_ending', periodEndingText);
  return out;
}

export function resolveFormXIXKAEmployeeToken(emp = {}) {
  return String(
    emp.EmployeeID ||
      emp['EmployeeID'] ||
      emp['Employee ID'] ||
      emp.Employee_Number ||
      emp['Employee Number'] ||
      emp.EmployeeCode ||
      emp['Employee Code'] ||
      ''
  ).trim();
}

export function formatFormXIXKASexAndIdentificationMarks(emp = {}) {
  const sex = String(emp.Sex || emp.Gender || emp['Gender'] || emp['Sex'] || '').trim();
  const marks = String(
    emp.Identification_Marks ||
      emp['Identification Marks'] ||
      emp.IdentificationMarks ||
      emp.identification_marks ||
      emp['Identification_marks'] ||
      ''
  ).trim();
  if (sex && marks) return `${sex}\n${marks}`;
  return sex || marks;
}

const pickKarnatakaPayrollValue = (flat, payrollRow, keys, patterns = []) => {
  const fromFlat = readPayrollScalar(flat, keys, patterns);
  if (fromFlat !== '') return fromFlat;
  if (payrollRow && payrollRow !== flat) {
    return readPayrollScalar(payrollRow, keys, patterns);
  }
  return '';
};

const pickKarnatakaPayrollText = (flat, payrollRow, keys, patterns = []) => {
  const fromFlat = readPayrollTextScalar(flat, keys, patterns);
  if (fromFlat !== '') return fromFlat;
  if (payrollRow && payrollRow !== flat) {
    return readPayrollTextScalar(payrollRow, keys, patterns);
  }
  return '';
};

const formatKarnatakaPayrollDate = (value) => {
  const s = String(value ?? '').trim();
  if (!s) return '';
  const iso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) return `${iso[3]}-${iso[2]}-${iso[1]}`;
  return s;
};

const pickGrossPayFromPayrollRow = (flat, payrollRow) => {
  if (flat.gross_pay !== '' && flat.gross_pay != null) return String(flat.gross_pay);
  return pickKarnatakaPayrollValue(
    flat,
    payrollRow,
    ['gross_pay', 'Gross Pay', 'grossPay', 'total_earnings', 'Total Earnings'],
    [/^gross_pay$/, /^total_earnings$/]
  );
};

const sumPayrollDeductionLines = (payrollRow) => {
  const items = payrollRow?.deductions ?? payrollRow?.Deductions;
  if (!Array.isArray(items) || items.length === 0) return '';
  let sum = 0;
  let any = false;
  items.forEach((item) => {
    const raw = item?.amount ?? item?.value ?? item?.employee_contribution ?? '';
    const n = Number(String(raw).replace(/,/g, '').trim());
    if (Number.isFinite(n)) {
      sum += n;
      any = true;
    }
  });
  return any ? String(Math.round(sum * 100) / 100) : '';
};

const pickDeductionsFromPayrollRow = (flat, payrollRow) => {
  let deductions = pickKarnatakaPayrollValue(
    flat,
    payrollRow,
    ['total_deductions', 'Total Deductions', 'totalDeductions', 'total_employee_deductions', 'total_deduction'],
    [/^total_deductions?$/, /^total_employee_deductions$/]
  );
  if (deductions === '') deductions = sumPayrollDeductionLines(payrollRow);
  if (deductions === '') {
    const gross = pickGrossPayFromPayrollRow(flat, payrollRow);
    const net = pickKarnatakaPayrollValue(
      flat,
      payrollRow,
      ['net_pay', 'Net Pay', 'netPay', 'net_wages'],
      [/^net_pay$/]
    );
    if (gross !== '' && net !== '') {
      const diff = Number(gross) - Number(net);
      if (Number.isFinite(diff) && diff >= 0) deductions = String(Math.round(diff * 100) / 100);
    }
  }
  return deductions !== '' ? String(deductions) : '';
};

const pickNetPayFromPayrollRow = (flat, payrollRow) => {
  let net = pickKarnatakaPayrollValue(
    flat,
    payrollRow,
    ['net_pay', 'Net Pay', 'netPay', 'net_wages', 'Net Wages', 'monthly_salary'],
    [/^net_pay$/, /^net_wages$/]
  );
  if (net !== '') return net;
  const gross = pickGrossPayFromPayrollRow(flat, payrollRow);
  const deductions = pickDeductionsFromPayrollRow(flat, payrollRow);
  if (gross !== '' && deductions !== '') {
    const diff = Number(String(gross).replace(/,/g, '')) - Number(String(deductions).replace(/,/g, ''));
    if (Number.isFinite(diff)) return String(Math.round(diff * 100) / 100);
  }
  return '';
};

const formatKarnatakaOvertimeHoursAndAmount = (hours, amount) => {
  const h = String(hours ?? '').trim();
  const a = String(amount ?? '').trim();
  if (h && a) return `${h} / ${a}`;
  return h || a || '';
};

/** Same OT dates list as Form XXIII, displayed DD-MM-YYYY for Karnataka wage slip. */
export function formatFormXIXKAOvertimeWorkedDatesList(records, dateKeyFn) {
  if (!Array.isArray(records) || records.length === 0) return '';
  const dates = [];
  const seen = new Set();
  records.forEach((rec) => {
    const dk =
      (typeof dateKeyFn === 'function' ? dateKeyFn(rec) : '') ||
      String(rec?.date || rec?.Date || '').trim();
    const iso = String(dk || '').trim();
    if (!iso || seen.has(iso)) return;
    seen.add(iso);
    dates.push(formatKarnatakaPayrollDate(iso) || iso);
  });
  dates.sort();
  return dates.join(', ');
}

const pickKarnatakaOvertimeAmount = (payrollRow) => {
  if (!payrollRow || payrollRow.fetch_error) return '';
  const flat = flattenPayrollEarningColumns(payrollRow);
  if (flat.overtime !== '' && flat.overtime != null) return String(flat.overtime);
  return pickKarnatakaPayrollValue(
    flat,
    payrollRow,
    ['overtime', 'Overtime', 'overtime_wages', 'overtimeWages', 'overtime amount', 'Overtime Amount'],
    [/^overtime$/, /overtime.*wage/i, /overtime.*amount/i]
  );
};

/**
 * Fill Karnataka Form XIX OT columns from attendance (same source as Form XXIII Register of Overtime).
 */
export function enrichFormXIXKarnatakaAttendanceOvertimeRows(mappedData, employees, headers, helpers = {}) {
  const {
    resolveOtRecords = null,
    getOtHours = null,
    dateKeyForOt = null,
    formatOtHours = (n) => String(n ?? ''),
    resolvePayrollRow = null,
    sanitizeValue = (v) => String(v ?? '').trim(),
    overwrite = true,
  } = helpers;
  if (
    !Array.isArray(mappedData) ||
    mappedData.length === 0 ||
    typeof resolveOtRecords !== 'function' ||
    typeof getOtHours !== 'function'
  ) {
    return 0;
  }

  const hdrs = resolveFormXIXKarnatakaWageTableHeaders(headers);
  const otDatesHdr = hdrs.find(isFormXIXKAOvertimeDatesHeader);
  const otHoursHdr = hdrs.find(isFormXIXKAOvertimeAmountHeader);
  if (!otDatesHdr && !otHoursHdr) return 0;

  const canWrite = (current) => {
    if (overwrite) return true;
    return !String(current ?? '').trim();
  };

  let hits = 0;
  mappedData.forEach((row, rowIndex) => {
    if (!row || typeof row !== 'object') return;
    const empItem = employees[rowIndex];
    const emp = empItem?.Employee || empItem?.employee || empItem;
    const otRecords = resolveOtRecords(emp, row, rowIndex);
    if (!Array.isArray(otRecords) || otRecords.length === 0) return;

    const totalOtHours = otRecords.reduce((sum, rec) => sum + (Number(getOtHours(rec)) || 0), 0);
    if (totalOtHours <= 0) return;

    const otDates = formatFormXIXKAOvertimeWorkedDatesList(otRecords, dateKeyForOt);
    const hoursDisplay = formatOtHours(totalOtHours);
    const payrollRow = typeof resolvePayrollRow === 'function' ? resolvePayrollRow(emp, row, rowIndex) : null;
    const otAmount =
      payrollRow && !payrollRow.fetch_error ? pickKarnatakaOvertimeAmount(payrollRow) : '';
    const hoursAndAmount = formatKarnatakaOvertimeHoursAndAmount(hoursDisplay, otAmount);

    if (otDatesHdr && otDates && canWrite(row[otDatesHdr])) {
      row[otDatesHdr] = sanitizeValue(otDates);
    }
    if (otHoursHdr && hoursAndAmount && canWrite(row[otHoursHdr])) {
      row[otHoursHdr] = sanitizeValue(hoursAndAmount);
    }
    hits += 1;
  });
  return hits;
}

/** Karnataka Form XIX — payroll column mapping (paid_days, gross_pay, monthly OT date/hours). */
export function resolveFormXIXKarnatakaPayrollFields(payrollRow) {
  if (!payrollRow || payrollRow.fetch_error) {
    return {
      daysWorked: '',
      overtimeDates: '',
      overtimeHoursAndAmount: '',
      grossWages: '',
      deductions: '',
      netWages: '',
    };
  }
  const flat = flattenPayrollEarningColumns(payrollRow);

  const daysWorked = pickKarnatakaPayrollValue(
    flat,
    payrollRow,
    ['paid_days', 'Paid Days', 'days_worked', 'Days Worked', 'paidDays', 'no_of_days_worked'],
    [/^paid_days$/, /paiddays/, /daysworked/, /days_present/, /noofdayspresent/, /no_of_days_present/]
  );

  let overtimeDates = formatKarnatakaPayrollDate(flat.overtime_date);
  if (!overtimeDates) {
    overtimeDates = formatKarnatakaPayrollDate(
      pickKarnatakaPayrollText(
        flat,
        payrollRow,
        [
          'overtime_date',
          'overtime dates',
          'Overtime Date',
          'Overtime Dates',
          'overtime_dates',
          'dates_on_which_overtime_worked',
          'overtime_worked_dates',
        ],
        [/overtime.*date/i, /dates.*overtime/i, /overtime.*worked.*date/i]
      )
    );
  }

  let overtimeHours = '';
  if (flat.overtime_hours !== '' && flat.overtime_hours != null) {
    overtimeHours = String(flat.overtime_hours);
  }
  if (overtimeHours === '') {
    overtimeHours =
      pickKarnatakaPayrollText(
        flat,
        payrollRow,
        [
          'overtime_hours',
          'overtime hours',
          'Overtime Hours',
          'total_overtime_hours',
          'total overtime hours',
          'overtime_hours_total',
          'overtime hours total',
          'ot_hours',
          'overtime_hrs',
          'monthly_overtime_hours',
          'total_ot_hours',
        ],
        [/total.*overtime.*hour/i, /^overtime_hours$/, /^total_overtime_hours$/, /^ot_hours$/, /overtime.*hrs/i]
      ) ||
      pickKarnatakaPayrollValue(
        flat,
        payrollRow,
        [
          'overtime_hours',
          'total_overtime_hours',
          'overtime_hours_total',
          'monthly_overtime_hours',
          'ot_hours',
          'total_ot_hours',
        ],
        [/total.*overtime.*hour/i, /^overtime_hours$/, /^total_overtime_hours$/, /^ot_hours$/]
      );
  }

  let overtimeAmount = '';
  if (flat.overtime !== '' && flat.overtime != null) {
    overtimeAmount = String(flat.overtime);
  }
  if (overtimeAmount === '') {
    overtimeAmount = pickKarnatakaPayrollValue(
      flat,
      payrollRow,
      ['overtime', 'Overtime', 'overtime_wages', 'overtimeWages', 'overtime amount', 'Overtime Amount'],
      [/^overtime$/, /overtime.*wage/i, /overtime.*amount/i]
    );
  }

  return {
    daysWorked,
    overtimeDates,
    overtimeHoursAndAmount: formatKarnatakaOvertimeHoursAndAmount(overtimeHours, overtimeAmount),
    grossWages: pickGrossPayFromPayrollRow(flat, payrollRow),
    deductions: pickDeductionsFromPayrollRow(flat, payrollRow),
    netWages: pickNetPayFromPayrollRow(flat, payrollRow),
  };
}

export function applyFormXIXKarnatakaEmployeeToRow(row, emp, headers, helpers = {}) {
  const headerList =
    Array.isArray(headers) && headers.length > 0
      ? headers.filter(Boolean)
      : resolveFormXIXKarnatakaWageTableHeaders(headers);
  const out = row && typeof row === 'object' ? { ...row } : {};
  const {
    sanitizeValue = (v) => String(v ?? '').trim(),
    payrollRow = null,
    resolvePayrollFields = null,
  } = helpers;
  const hasPayroll = payrollRow && !payrollRow.fetch_error;
  const payroll =
    hasPayroll && typeof resolvePayrollFields === 'function'
      ? resolvePayrollFields(payrollRow)
      : hasPayroll
        ? resolveFormXIXKarnatakaPayrollFields(payrollRow)
        : {};
  const wageValue = (value, current = '') => {
    if (hasPayroll && value !== '') return sanitizeValue(value);
    return String(current ?? '').trim();
  };

  headerList.forEach((header) => {
    if (isFormXIXKASexIdentificationHeader(header)) {
      out[header] = sanitizeValue(formatFormXIXKASexAndIdentificationMarks(emp));
      return;
    }
    if (isFormXIXKATokenHeader(header)) {
      out[header] = sanitizeValue(resolveFormXIXKAEmployeeToken(emp));
      return;
    }
    if (isFormXIXKAWorkmanNameHeader(header)) {
      out[header] = sanitizeValue(formatWorkmanNameAndGuardian(emp));
      return;
    }
    if (isFormXIXKADaysWorkedHeader(header)) {
      out[header] = wageValue(payroll.daysWorked, out[header]);
      return;
    }
    if (isFormXIXKARateHeader(header)) {
      out[header] = sanitizeValue(FORM_XIX_KA_RATE_DEFAULT);
      return;
    }
    if (isFormXIXKAUnitsHeader(header)) {
      out[header] = '';
      return;
    }
    if (isFormXIXKAOvertimeDatesHeader(header)) {
      out[header] = wageValue(payroll.overtimeDates, out[header]);
      return;
    }
    if (isFormXIXKAOvertimeAmountHeader(header)) {
      out[header] = wageValue(payroll.overtimeHoursAndAmount, out[header]);
      return;
    }
    if (isFormXIXKAGrossHeader(header)) {
      out[header] = wageValue(payroll.grossWages, out[header]);
      return;
    }
    if (isFormXIXKADeductionsHeader(header)) {
      out[header] = wageValue(payroll.deductions, out[header]);
      return;
    }
    if (isFormXIXKANetHeader(header)) {
      out[header] = wageValue(payroll.netWages, out[header]);
    }
  });
  Object.keys(out).forEach((key) => {
    if (String(key).startsWith('__')) return;
    if (isFormXIXKARateHeader(key)) out[key] = sanitizeValue(FORM_XIX_KA_RATE_DEFAULT);
    if (isFormXIXKAUnitsHeader(key)) out[key] = '';
  });
  out.__employeeLookupName = sanitizeValue(formatWorkmanNameAndGuardian(emp).split(/\r?\n/)[0]);
  return out;
}

export function mapFormXIXKarnatakaRowsFromEmployees(employees, headers, helpers = {}) {
  const hdrs = resolveFormXIXKarnatakaWageTableHeaders(headers);
  const list = Array.isArray(employees) ? employees : [];
  const { resolvePayrollRow = null, resolvePayrollFields = null, ...rest } = helpers;
  return list.map((empItem, rowIndex) => {
    const emp = empItem?.Employee || empItem?.employee || empItem;
    const payrollRow = typeof resolvePayrollRow === 'function' ? resolvePayrollRow(emp, rowIndex) : null;
    return applyFormXIXKarnatakaEmployeeToRow({}, emp, hdrs, {
      ...rest,
      payrollRow: payrollRow && !payrollRow.fetch_error ? payrollRow : null,
      resolvePayrollFields,
    });
  });
}

export function enrichFormXIXKarnatakaStaticFieldRows(mappedData, headers) {
  if (!Array.isArray(mappedData) || mappedData.length === 0) return 0;
  const hdrs = resolveFormXIXKarnatakaWageTableHeaders(headers);
  const rateHdrs = hdrs.filter(isFormXIXKARateHeader);
  if (rateHdrs.length === 0) return 0;
  let hits = 0;
  mappedData.forEach((row) => {
    if (!row || typeof row !== 'object') return;
    const rateKeys = new Set(rateHdrs);
    Object.keys(row).forEach((key) => {
      if (String(key).startsWith('__')) return;
      if (isFormXIXKARateHeader(key)) rateKeys.add(key);
      if (isFormXIXKAUnitsHeader(key)) row[key] = '';
    });
    rateKeys.forEach((header) => {
      row[header] = FORM_XIX_KA_RATE_DEFAULT;
    });
    if (rateKeys.size > 0) hits += 1;
  });
  return hits;
}

export function enrichFormXIXKarnatakaPayrollRows(mappedData, employees, headers, helpers = {}) {
  const hdrs = resolveFormXIXKarnatakaWageTableHeaders(headers);
  const {
    resolvePayrollRow = null,
    resolvePayrollFields = null,
    sanitizeValue = (v) => String(v ?? '').trim(),
    overwrite = true,
  } = helpers;
  if (!Array.isArray(mappedData) || mappedData.length === 0 || typeof resolvePayrollRow !== 'function') {
    return 0;
  }
  let hits = 0;
  mappedData.forEach((row, rowIndex) => {
    const empItem = employees[rowIndex];
    const emp = empItem?.Employee || empItem?.employee || empItem;
    const payrollRow = typeof resolvePayrollRow === 'function' ? resolvePayrollRow(emp, row, rowIndex) : null;
    const merged = applyFormXIXKarnatakaEmployeeToRow(row, emp, hdrs, {
      sanitizeValue,
      payrollRow: payrollRow && !payrollRow.fetch_error ? payrollRow : null,
      resolvePayrollFields,
    });
    hdrs.forEach((header) => {
      const prev = String(row[header] ?? '').trim();
      const next = String(merged[header] ?? '').trim();
      if (!overwrite && prev !== '' && next !== '') {
        merged[header] = row[header];
        return;
      }
      if (prev !== '' && next === '' && (isFormXIXKAOvertimeDatesHeader(header) || isFormXIXKAOvertimeAmountHeader(header))) {
        merged[header] = row[header];
      }
    });
    Object.assign(row, merged);
    hits += 1;
  });
  return hits;
}

export function rowHasMeaningfulFormXIXKarnatakaExportData(row, headers) {
  if (!row || typeof row !== 'object') return false;
  const hdrs = resolveFormXIXKarnatakaWageTableHeaders(headers);
  return hdrs.some((header) => getKarnatakaRowValueForHeader(row, header) !== '');
}

export function overlayFormXIXKarnatakaUserEditsOntoRows(mappedRows, tableRows, headers) {
  const hdrs = resolveFormXIXKarnatakaWageTableHeaders(headers);
  const savedRows = Array.isArray(tableRows) ? tableRows : [];
  return (Array.isArray(mappedRows) ? mappedRows : []).map((row, index) => {
    const saved = savedRows[index];
    if (!saved || typeof saved !== 'object' || !rowHasMeaningfulFormXIXKarnatakaExportData(saved, hdrs)) {
      return row;
    }
    const merged = { ...row };
    hdrs.forEach((header) => {
      const value = getKarnatakaRowValueForHeader(saved, header);
      if (value !== '') merged[header] = saved[header] ?? value;
    });
    return merged;
  });
}

const excelCellValueToString = (val) => {
  if (val == null) return '';
  if (typeof val === 'string' || typeof val === 'number' || typeof val === 'boolean') return String(val);
  if (val instanceof Date) return val.toISOString();
  if (typeof val === 'object') {
    if (Array.isArray(val.richText)) return val.richText.map((rt) => rt?.text || '').join('');
    if (val.text != null) return String(val.text);
    if (val.result != null) return String(val.result);
  }
  return '';
};

export function writeFormXIXKarnatakaHeaderFieldsToWorksheet(worksheet, headerFormData = {}, parsedFormHeader = null) {
  if (!worksheet || !headerFormData || typeof headerFormData !== 'object') return;
  const parsedFields = Array.isArray(parsedFormHeader?.fields) ? parsedFormHeader.fields : [];
  const writeAt = (row, col, value) => {
    const text = String(value ?? '').trim();
    if (!text || row < 1 || col < 1) return;
    worksheet.getCell(row, col).value = text;
  };
  FORM_XIX_KA_HEADER_SPECS.forEach((spec) => {
    const val = headerFormData[spec.key];
    if (val == null || String(val).trim() === '') return;
    const parsedField = parsedFields.find((f) => f.key === spec.key);
    if (parsedField?.labelRow != null) {
      writeAt(parsedField.labelRow + 1, FORM_XIX_KA_DEFAULT_VALUE_COL, val);
      return;
    }
    for (let r = 1; r <= 40; r += 1) {
      for (let c = 1; c <= 14; c += 1) {
        const raw = excelCellValueToString(worksheet.getCell(r, c)?.value).trim();
        if (!raw || (!spec.match.test(raw) && !spec.match.test(formXIXAPHeaderNorm(raw)))) continue;
        writeAt(r, FORM_XIX_KA_DEFAULT_VALUE_COL, val);
        return;
      }
    }
  });
  FORM_XIX_KA_EMPLOYEE_HEADER_SPECS.forEach((spec) => {
    const val = headerFormData[spec.key];
    if (val == null || String(val).trim() === '') return;
    const parsedField = parsedFields.find((f) => f.key === spec.key);
    if (parsedField?.labelRow != null) {
      writeAt(parsedField.labelRow + 1, FORM_XIX_KA_DEFAULT_VALUE_COL, val);
      return;
    }
    for (let r = 1; r <= 40; r += 1) {
      for (let c = 1; c <= 14; c += 1) {
        const raw = excelCellValueToString(worksheet.getCell(r, c)?.value).trim();
        if (!raw || (!spec.match.test(raw) && !spec.match.test(formXIXAPHeaderNorm(raw)))) continue;
        writeAt(r, FORM_XIX_KA_DEFAULT_VALUE_COL, val);
        return;
      }
    }
  });
}

function resolveTableColumns(parsedFormHeader) {
  const fromHeader = Array.isArray(parsedFormHeader?.tableColumns)
    ? parsedFormHeader.tableColumns.filter((h) => h && !isFormXIXKAIdentityTableHeader(h))
    : [];
  if (fromHeader.length >= FORM_XIX_KA_EXCEL_WAGE_HEADERS.length) return fromHeader;
  const startCol = Number(parsedFormHeader?.tableStartCol ?? 0);
  return FORM_XIX_KA_EXCEL_WAGE_HEADERS.map((header, i) => ({ header, col: startCol + i }));
}

function resolveKarnatakaTableColumnsFromWorksheet(worksheet, parsedFormHeader) {
  const parsedCols = resolveTableColumns(parsedFormHeader);
  if (!worksheet) return parsedCols;

  const maxR = 45;
  const maxC = 20;
  const wageHdrs = FORM_XIX_KA_EXCEL_WAGE_HEADERS;
  for (let r = 1; r <= maxR; r += 1) {
    for (let c = 1; c <= maxC; c += 1) {
      const raw = excelCellValueToString(worksheet.getCell(r, c)?.value).trim();
      if (!isFormXIXKADaysWorkedHeader(raw)) continue;
      return wageHdrs.map((header, i) => ({
        header,
        col: c - 1 + i,
        headerRow: r,
      }));
    }
  }
  return parsedCols;
}

function resolveKarnatakaDataRowFromWorksheet(worksheet, parsedFormHeader, columns) {
  const daysCol = columns.find((col) => isFormXIXKADaysWorkedHeader(col.header));
  const headerRow = Number(daysCol?.headerRow ?? 0);
  if (worksheet && headerRow > 0) {
    const col1 = Number(daysCol.col) + 1;
    for (let r = headerRow + 1; r <= headerRow + 5; r += 1) {
      const marker = excelCellValueToString(worksheet.getCell(r, col1)?.value).trim();
      if (/^[1-5]$/.test(marker) || /^\d+$/.test(marker)) {
        return r + 1;
      }
    }
    return headerRow + 2;
  }
  const parsedRow = Number(parsedFormHeader?.dataStartIndex);
  if (Number.isFinite(parsedRow) && parsedRow >= 0) {
    return parsedRow + 1;
  }
  return 17;
}

function resolveKarnatakaFooterColumnsFromWorksheet(worksheet, parsedFormHeader, columns, dataRow) {
  const parsedFooter = resolveFooterColumns(parsedFormHeader);
  if (!worksheet) return parsedFooter;

  const footerHdrs = FORM_XIX_KA_FOOTER_HEADERS;
  const headerRow = Number(columns.find((col) => isFormXIXKADaysWorkedHeader(col.header))?.headerRow ?? 0);
  const scanFrom = Math.max(headerRow, dataRow - 1);
  for (let r = scanFrom; r <= scanFrom + 10; r += 1) {
    const matched = [];
    for (let c = 1; c <= 20; c += 1) {
      const cell = excelCellValueToString(worksheet.getCell(r, c)?.value).trim();
      if (!cell) continue;
      if (isFormXIXKAGrossHeader(cell)) matched.push({ header: footerHdrs[0], col: c - 1, labelRow: r - 1, valueRow: r + 1 });
      else if (isFormXIXKADeductionsHeader(cell)) matched.push({ header: footerHdrs[1], col: c - 1, labelRow: r - 1, valueRow: r + 1 });
      else if (isFormXIXKANetHeader(cell)) matched.push({ header: footerHdrs[2], col: c - 1, labelRow: r - 1, valueRow: r + 1 });
    }
    if (matched.length >= 2) return matched;
  }
  return parsedFooter;
}

function resolveFooterColumns(parsedFormHeader) {
  const fromHeader = Array.isArray(parsedFormHeader?.footerColumns) ? parsedFormHeader.footerColumns : [];
  if (fromHeader.length > 0) return fromHeader;
  return FORM_XIX_KA_FOOTER_HEADERS.map((header, i) => ({
    header,
    col: 1 + i * 2,
    labelRow: Number(parsedFormHeader?.footerRowIndex ?? 18),
  }));
}

/** Resolve where footer amounts are written in the Excel template (beside or below labels). */
function buildKarnatakaFooterValuePositions(parsedFormHeader, worksheet, tableColumns, dataRow) {
  const footerColumns =
    Array.isArray(parsedFormHeader?.footerColumns) && parsedFormHeader.footerColumns.length >= 2
      ? parsedFormHeader.footerColumns
      : worksheet
        ? resolveKarnatakaFooterColumnsFromWorksheet(worksheet, parsedFormHeader, tableColumns, dataRow)
        : resolveFooterColumns(parsedFormHeader);

  const footerValueRow0 =
    parsedFormHeader?.footerValueRow != null
      ? Number(parsedFormHeader.footerValueRow)
      : footerColumns[0]?.labelRow != null
        ? Number(footerColumns[0].labelRow) + 1
        : Number(parsedFormHeader?.footerRowIndex ?? 18) + 1;

  const positions = [];
  const seen = new Set();

  const pushPos = (header, row, col) => {
    if (!header || row == null || col == null || row < 1 || col < 1) return;
    const cellRef = formXIXKAToCellRef(row, col);
    const key = `${header}|${cellRef}`;
    if (seen.has(key)) return;
    seen.add(key);
    positions.push({ header, row, col, cellRef });
  };

  footerColumns.forEach(({ header, col, labelRow }) => {
    const col0 = Number(col);
    const labelRow0 = labelRow != null ? Number(labelRow) : Number(parsedFormHeader?.footerRowIndex ?? 18);
    const labelExcelRow = labelRow0 + 1;
    const belowExcelRow = footerValueRow0 + 1;

    // Karnataka template: amount typically in the cell to the right of the label (same row).
    pushPos(header, labelExcelRow, col0 + 2);
    // Fallback: amount below the label in the same column.
    pushPos(header, belowExcelRow, col0 + 1);
  });

  return positions;
}

export function applyFormXIXKarnatakaFooterFieldsToRow(row, payrollRow, headers) {
  if (!row || typeof row !== 'object' || !payrollRow || payrollRow.fetch_error) return row;
  const payroll = resolveFormXIXKarnatakaPayrollFields(payrollRow);
  const out = { ...row };
  const hdrs = resolveFormXIXKarnatakaWageTableHeaders(headers);
  const footerValues = [payroll.grossWages, payroll.deductions, payroll.netWages];
  FORM_XIX_KA_FOOTER_HEADERS.forEach((footerHeader, index) => {
    const value = String(footerValues[index] ?? '').trim();
    if (!value) return;
    out[footerHeader] = value;
    hdrs.forEach((header) => {
      if (headerMatchesKarnatakaFooter(header, footerHeader)) {
        out[header] = value;
      }
    });
  });
  return out;
}

export function writeFormXIXKarnatakaFooterRowToWorksheet(worksheet, row, parsedFormHeader) {
  if (!worksheet || !row || typeof row !== 'object') return;
  const tableColumns = resolveKarnatakaTableColumnsFromWorksheet(worksheet, parsedFormHeader);
  const dataRow = resolveKarnatakaDataRowFromWorksheet(worksheet, parsedFormHeader, tableColumns);
  const footerPositions = buildKarnatakaFooterValuePositions(
    parsedFormHeader,
    worksheet,
    tableColumns,
    dataRow
  );
  footerPositions.forEach(({ header, row: writeRow, col: writeCol }) => {
    const value = resolveKarnatakaExportCellValue(row, header);
    if (value === '') return;
    worksheet.getCell(writeRow, writeCol).value = value;
  });
}

export function writeFormXIXKarnatakaTableRowToWorksheet(worksheet, row, parsedFormHeader, dataRowIndex) {
  if (!worksheet || !row || typeof row !== 'object') return;
  const columns = resolveKarnatakaTableColumnsFromWorksheet(worksheet, parsedFormHeader);
  const dataRow =
    dataRowIndex != null && Number.isFinite(Number(dataRowIndex))
      ? Number(dataRowIndex) + 1
      : resolveKarnatakaDataRowFromWorksheet(worksheet, parsedFormHeader, columns);
  columns.forEach(({ header, col }) => {
    const value = resolveKarnatakaExportCellValue(row, header);
    if (value === '') return;
    worksheet.getCell(dataRow, Number(col) + 1).value = value;
  });
  writeFormXIXKarnatakaFooterRowToWorksheet(worksheet, row, parsedFormHeader);
}

function buildEmployeeHeaderFormData(headerFormData, employeeRow, emp) {
  const base = headerFormData && typeof headerFormData === 'object' ? { ...headerFormData } : {};
  const hdrs = FORM_XIX_KA_WAGE_TABLE_HEADERS;
  const workmanHdr = hdrs.find(isFormXIXKAWorkmanNameHeader);
  const tokenHdr = hdrs.find(isFormXIXKATokenHeader);
  const sexHdr = hdrs.find(isFormXIXKASexIdentificationHeader);
  const workmanFromRow = workmanHdr ? getKarnatakaRowValueForHeader(employeeRow, workmanHdr) : '';
  const tokenFromRow = tokenHdr ? getKarnatakaRowValueForHeader(employeeRow, tokenHdr) : '';
  const sexFromRow = sexHdr ? getKarnatakaRowValueForHeader(employeeRow, sexHdr) : '';
  const workman = workmanFromRow || formatWorkmanNameAndGuardian(emp);
  if (workman) base.form_xix_ka_workman = workman;
  const token = tokenFromRow || resolveFormXIXKAEmployeeToken(emp);
  if (token) base.form_xix_ka_token = token;
  const sexIdentification = sexFromRow || formatFormXIXKASexAndIdentificationMarks(emp);
  if (sexIdentification) base.form_xix_ka_sex_identification = sexIdentification;
  return base;
}

export async function buildFormXIXKarnatakaWorkbookWithTemplateStyles({
  templateArrayBuffer,
  mappedData,
  headersToUse,
  parsedFormHeader,
  formFileName,
  headerFormData,
  employeesOverride,
  resolvePayrollRow,
}) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(templateArrayBuffer);
  const worksheet = workbook.worksheets[0];
  if (!worksheet) throw new Error('Template worksheet not found.');

  const hdrs = resolveFormXIXKarnatakaWageTableHeaders(headersToUse);
  let exportRows = (Array.isArray(mappedData) ? mappedData : []).filter((row) =>
    rowHasMeaningfulFormXIXKarnatakaExportData(row, hdrs)
  );
  enrichFormXIXKarnatakaStaticFieldRows(exportRows, hdrs);
  const employees = Array.isArray(employeesOverride) ? employeesOverride : [];
  const employeeRow = exportRows[0] || null;
  const emp = employees[0]?.Employee || employees[0]?.employee || employees[0] || null;
  const payrollRow =
    typeof resolvePayrollRow === 'function' && emp ? resolvePayrollRow(emp) : null;
  let rowToWrite = employeeRow;
  if (employeeRow && payrollRow && !payrollRow.fetch_error) {
    rowToWrite = applyFormXIXKarnatakaFooterFieldsToRow(employeeRow, payrollRow, hdrs);
  }
  const mergedHeader = buildEmployeeHeaderFormData(headerFormData, rowToWrite, emp);

  writeFormXIXKarnatakaHeaderFieldsToWorksheet(worksheet, mergedHeader, parsedFormHeader);
  if (rowToWrite) {
    writeFormXIXKarnatakaTableRowToWorksheet(
      worksheet,
      rowToWrite,
      { ...parsedFormHeader, dataStartIndex: parsedFormHeader?.dataStartIndex },
      parsedFormHeader?.dataStartIndex
    );
  }

  const out = await workbook.xlsx.writeBuffer();
  const fileName =
    formFileName ||
    parsedFormHeader?.title?.replace(/[^a-zA-Z0-9]/g, '_') ||
    `Form_XIX_Karnataka_${Date.now()}.xlsx`;
  return { blob: new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), fileName };
}

function resolveEmployeeDownloadBaseName(row, fallbackIndex = 0) {
  const raw = String(row?.__employeeLookupName ?? '').trim();
  const slug = raw
    .replace(/[^\w\s.-]/g, '')
    .replace(/\s+/g, '_')
    .replace(/_+/g, '_')
    .replace(/^_|_$/g, '')
    .slice(0, 48);
  return slug || `Employee_${fallbackIndex + 1}`;
}

const formXIXKASanitizeExportText = (value) =>
  String(value ?? '')
    .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '')
    .trim();

const formXIXKAEscapeXml = (value) =>
  String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');

const formXIXKAColToLetter = (col) => {
  let result = '';
  let n = col;
  while (n > 0) {
    const rem = (n - 1) % 26;
    result = String.fromCharCode(65 + rem) + result;
    n = Math.floor((n - 1) / 26);
  }
  return result;
};

const formXIXKAToCellRef = (row, col) => `${formXIXKAColToLetter(col)}${row}`;

const formXIXKAUpsertInlineStrCell = (sheetXml, cellRef, value) => {
  const text = formXIXKAEscapeXml(formXIXKASanitizeExportText(value));
  const cellXml = text
    ? `<c r="${cellRef}" t="inlineStr"><is><t xml:space="preserve">${text}</t></is></c>`
    : `<c r="${cellRef}"/>`;
  const cellRe = new RegExp(`<c\\s+r="${cellRef}"[^>]*(?:/>|>[\\s\\S]*?</c>)`, 'i');
  if (cellRe.test(sheetXml)) {
    return sheetXml.replace(cellRe, cellXml);
  }
  const rowNum = cellRef.replace(/^[A-Z]+/i, '');
  const rowRe = new RegExp(`(<row\\s+r="${rowNum}"[^>]*>)([\\s\\S]*?)(</row>)`, 'i');
  if (!rowRe.test(sheetXml)) return sheetXml;
  return sheetXml.replace(rowRe, `$1$2${cellXml}$3`);
};

const resolveFormXIXKAWorksheetEntry = (zipFiles) =>
  Object.keys(zipFiles)
    .filter((name) => /^xl\/worksheets\/sheet\d+\.xml$/i.test(name))
    .sort((a, b) => a.localeCompare(b, undefined, { numeric: true }))[0] || null;

function resolveFormXIXKarnatakaFastExportPositions(worksheet, parsedFormHeader) {
  const employeeHeaderPositions = [];
  const seenHeaderKeys = new Set();
  const parsedFields = Array.isArray(parsedFormHeader?.fields) ? parsedFormHeader.fields : [];

  const pushEmployeeHeader = (key, row, col) => {
    if (!key || row == null || col == null || seenHeaderKeys.has(key)) return;
    seenHeaderKeys.add(key);
    employeeHeaderPositions.push({ key, row, col, cellRef: formXIXKAToCellRef(row, col) });
  };

  FORM_XIX_KA_EMPLOYEE_HEADER_SPECS.forEach((spec) => {
    const parsedField = parsedFields.find((f) => f.key === spec.key);
    if (parsedField?.labelRow != null) {
      pushEmployeeHeader(
        spec.key,
        (parsedField.valueRow ?? parsedField.labelRow) + 1,
        parsedField.valueCol != null ? parsedField.valueCol + 1 : FORM_XIX_KA_DEFAULT_VALUE_COL
      );
      return;
    }
    if (!worksheet) return;
    for (let r = 1; r <= 40; r += 1) {
      for (let c = 1; c <= 14; c += 1) {
        const raw = excelCellValueToString(worksheet.getCell(r, c)?.value).trim();
        if (!raw || (!spec.match.test(raw) && !spec.match.test(formXIXAPHeaderNorm(raw)))) continue;
        pushEmployeeHeader(spec.key, r, FORM_XIX_KA_DEFAULT_VALUE_COL);
        return;
      }
    }
  });

  const tableColumns = worksheet
    ? resolveKarnatakaTableColumnsFromWorksheet(worksheet, parsedFormHeader)
    : resolveTableColumns(parsedFormHeader);
  const dataRow = worksheet
    ? resolveKarnatakaDataRowFromWorksheet(worksheet, parsedFormHeader, tableColumns)
    : Number(parsedFormHeader?.dataStartIndex ?? 14) + 1;
  const tablePositions = tableColumns.map(({ header, col }) => ({
    header,
    row: dataRow,
    col: Number(col) + 1,
    cellRef: formXIXKAToCellRef(dataRow, Number(col) + 1),
  }));

  const footerPositions = buildKarnatakaFooterValuePositions(
    parsedFormHeader,
    worksheet,
    tableColumns,
    dataRow
  );

  return { employeeHeaderPositions, tablePositions, footerPositions };
}

const clearFormXIXKarnatakaPerEmployeeValueCells = (worksheet, positions) => {
  if (!worksheet || !positions) return;
  const all = [
    ...(positions.employeeHeaderPositions || []),
    ...(positions.tablePositions || []),
    ...(positions.footerPositions || []),
  ];
  all.forEach((pos) => {
    if (pos?.row != null && pos?.col != null) {
      worksheet.getCell(pos.row, pos.col).value = '';
    }
  });
};

const patchFormXIXKarnatakaFastSheetXml = (baseSheetXml, mergedHeaderData, exportRow, positions) => {
  let sheetXml = baseSheetXml;
  (positions.employeeHeaderPositions || []).forEach((pos) => {
    const value = mergedHeaderData?.[pos.key];
    if (value != null && String(value).trim() !== '') {
      sheetXml = formXIXKAUpsertInlineStrCell(sheetXml, pos.cellRef, value);
    }
  });
  (positions.tablePositions || []).forEach((pos) => {
    const value = resolveKarnatakaExportCellValue(exportRow, pos.header);
    if (value !== '') {
      sheetXml = formXIXKAUpsertInlineStrCell(sheetXml, pos.cellRef, value);
    }
  });
  (positions.footerPositions || []).forEach((pos) => {
    const value = resolveKarnatakaExportCellValue(exportRow, pos.header);
    if (value !== '') {
      sheetXml = formXIXKAUpsertInlineStrCell(sheetXml, pos.cellRef, value);
    }
  });
  return sheetXml;
};

const buildFormXIXKarnatakaFastXlsxBytes = async (fastTemplate, mergedHeaderData, exportRow) => {
  const { sheetEntry, baseSheetXml, staticFiles, positions } = fastTemplate;
  const sheetXml = patchFormXIXKarnatakaFastSheetXml(baseSheetXml, mergedHeaderData, exportRow, positions);
  const entryZip = new JSZip();
  Object.entries(staticFiles).forEach(([path, data]) => {
    entryZip.file(path, data);
  });
  entryZip.file(sheetEntry, sheetXml);
  return entryZip.generateAsync({ type: 'uint8array', compression: 'STORE' });
};

export function canUseFormXIXKarnatakaFastExport(parsedFormHeader) {
  return !!parsedFormHeader?.formXIXKarnatakaTableLayout;
}

async function prepareFormXIXKarnatakaFastZipTemplate({
  templateArrayBuffer,
  parsedFormHeader,
  headerFormData,
}) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(templateArrayBuffer);
  const worksheet = workbook.worksheets[0];
  if (!worksheet) throw new Error('Template worksheet not found.');

  const staticHeaderData = { ...(headerFormData || {}) };
  FORM_XIX_KA_EMPLOYEE_HEADER_SPECS.forEach((spec) => {
    delete staticHeaderData[spec.key];
  });
  writeFormXIXKarnatakaHeaderFieldsToWorksheet(worksheet, staticHeaderData, parsedFormHeader);

  const positions = resolveFormXIXKarnatakaFastExportPositions(worksheet, parsedFormHeader);
  clearFormXIXKarnatakaPerEmployeeValueCells(worksheet, positions);

  const preparedBuffer = await workbook.xlsx.writeBuffer();
  const templateZip = await JSZip.loadAsync(preparedBuffer);
  const sheetEntry = resolveFormXIXKAWorksheetEntry(templateZip.files);
  if (!sheetEntry) throw new Error('Template worksheet XML not found.');
  const baseSheetXml = await templateZip.file(sheetEntry).async('string');
  const staticFiles = {};
  await Promise.all(
    Object.keys(templateZip.files).map(async (path) => {
      const file = templateZip.files[path];
      if (!file || file.dir || path === sheetEntry) return;
      staticFiles[path] = await file.async('uint8array');
    })
  );
  return { sheetEntry, baseSheetXml, staticFiles, positions };
}

const FORM_XIX_KA_FAST_ZIP_BATCH = 12;

async function buildFormXIXKarnatakaFastZipDownload({
  exportRows,
  fastTemplate,
  baseHeaderData,
  parsedFormHeader,
  formFileName,
  employees = [],
}) {
  const zip = new JSZip();
  const usedNames = new Map();
  for (let i = 0; i < exportRows.length; i += FORM_XIX_KA_FAST_ZIP_BATCH) {
    const batch = exportRows.slice(i, i + FORM_XIX_KA_FAST_ZIP_BATCH);
    const batchBytes = await Promise.all(
      batch.map(async (exportRow, batchIndex) => {
        const index = i + batchIndex;
        const emp = employees[index]?.Employee || employees[index]?.employee || employees[index] || null;
        const mergedHeaderData = buildEmployeeHeaderFormData(baseHeaderData, exportRow, emp);
        const xlsxBytes = await buildFormXIXKarnatakaFastXlsxBytes(
          fastTemplate,
          mergedHeaderData,
          exportRow
        );
        return { xlsxBytes, exportRow, index };
      })
    );
    batchBytes.forEach(({ xlsxBytes, exportRow, index }) => {
      const baseName = resolveEmployeeDownloadBaseName(exportRow, index);
      const count = usedNames.get(baseName) || 0;
      usedNames.set(baseName, count + 1);
      const suffix = count > 0 ? `_${count + 1}` : '';
      zip.file(`Form_XIX_Karnataka_${baseName}${suffix}.xlsx`, xlsxBytes);
    });
    if (i > 0) {
      await new Promise((resolve) => setTimeout(resolve, 0));
    }
  }
  const zipBase = String(formFileName || parsedFormHeader?.title || 'Form_XIX_Karnataka')
    .replace(/\.xlsx?$/i, '')
    .replace(/[^a-zA-Z0-9._-]+/g, '_');
  return {
    blob: await zip.generateAsync({ type: 'blob', compression: 'STORE' }),
    fileName: `${zipBase}_Employees.zip`,
  };
}

export function resolveFormXIXKarnatakaExportRows(mappedData, headers, employeesOverride = null, helpers = {}) {
  const hdrs = resolveFormXIXKarnatakaWageTableHeaders(headers);
  const tableRows = Array.isArray(mappedData) ? mappedData : [];
  const employees = Array.isArray(employeesOverride) ? employeesOverride : [];
  const {
    sanitizeValue = (v) => String(v ?? '').trim(),
    resolvePayrollRow = null,
    resolvePayrollFields = resolveFormXIXKarnatakaPayrollFields,
  } = helpers;
  if (employees.length > 0) {
    const mapHelpers = { sanitizeValue, resolvePayrollFields };
    if (typeof resolvePayrollRow === 'function') {
      mapHelpers.resolvePayrollRow = resolvePayrollRow;
    }
    const fromEmployees = mapFormXIXKarnatakaRowsFromEmployees(employees, hdrs, mapHelpers);
    enrichFormXIXKarnatakaStaticFieldRows(fromEmployees, hdrs);
    return overlayFormXIXKarnatakaUserEditsOntoRows(fromEmployees, tableRows, hdrs);
  }
  const rows = tableRows.filter((row) => rowHasMeaningfulFormXIXKarnatakaExportData(row, hdrs));
  enrichFormXIXKarnatakaStaticFieldRows(rows, hdrs);
  return rows;
}

export function triggerFormXIXKarnatakaZipDownload(blob, fileName) {
  if (!blob || !fileName) return;
  const downloadUrl = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = downloadUrl;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(downloadUrl);
}

export async function buildFormXIXKarnatakaPerEmployeeDownload({
  templateArrayBuffer,
  mappedData,
  headersToUse,
  parsedFormHeader,
  formFileName,
  headerFormData,
  employeesOverride,
  resolvePayrollRow,
}) {
  if (!templateArrayBuffer) {
    throw new Error('Original Form XIX Karnataka template could not be loaded. Open Autofill again, then Download.');
  }

  const hdrs = resolveFormXIXKarnatakaWageTableHeaders(headersToUse);
  const employees = Array.isArray(employeesOverride) ? employeesOverride : [];
  const tableRows = Array.isArray(mappedData) ? mappedData : [];
  const exportHelpers = {
    sanitizeValue: (v) => String(v ?? '').trim(),
    resolvePayrollRow: typeof resolvePayrollRow === 'function' ? resolvePayrollRow : null,
    resolvePayrollFields: resolveFormXIXKarnatakaPayrollFields,
  };
  const fromEmployees =
    employees.length > 0
      ? resolveFormXIXKarnatakaExportRows(tableRows, hdrs, employees, exportHelpers)
      : [];
  const fromTable = tableRows.filter((row) => rowHasMeaningfulFormXIXKarnatakaExportData(row, hdrs));
  let exportRows;
  if (fromTable.length > 0) {
    exportRows = fromTable.map((row) => ({ ...row }));
  } else if (employees.length > 0) {
    exportRows =
      fromEmployees.length > 0
        ? fromEmployees
        : mapFormXIXKarnatakaRowsFromEmployees(employees, hdrs, exportHelpers);
    if (exportRows.length < employees.length) {
      exportRows = mapFormXIXKarnatakaRowsFromEmployees(employees, hdrs, exportHelpers);
    }
    exportRows = overlayFormXIXKarnatakaUserEditsOntoRows(exportRows, tableRows, hdrs);
  } else {
    exportRows = [];
  }
  enrichFormXIXKarnatakaStaticFieldRows(exportRows, hdrs);

  const gridHasPayrollValues =
    fromTable.length > 0 &&
    fromTable.some((row) => {
      const daysHdr = hdrs.find(isFormXIXKADaysWorkedHeader);
      const grossHdr = hdrs.find(isFormXIXKAGrossHeader);
      return (
        (daysHdr && getKarnatakaRowValueForHeader(row, daysHdr) !== '') ||
        (grossHdr && getKarnatakaRowValueForHeader(row, grossHdr) !== '')
      );
    });

  if (
    exportRows.length > 0 &&
    typeof resolvePayrollRow === 'function' &&
    employees.length > 0 &&
    !gridHasPayrollValues
  ) {
    exportRows = exportRows.map((exportRow, index) => {
      const emp = employees[index]?.Employee || employees[index]?.employee || employees[index] || null;
      const payrollRow = resolvePayrollRow(emp);
      let merged = applyFormXIXKarnatakaEmployeeToRow(exportRow, emp, hdrs, {
        sanitizeValue: exportHelpers.sanitizeValue,
        payrollRow: payrollRow && !payrollRow.fetch_error ? payrollRow : null,
        resolvePayrollFields: resolveFormXIXKarnatakaPayrollFields,
      });
      if (payrollRow && !payrollRow.fetch_error) {
        merged = applyFormXIXKarnatakaFooterFieldsToRow(merged, payrollRow, hdrs);
      }
      return merged;
    });
    enrichFormXIXKarnatakaPayrollRows(exportRows, employees, hdrs, {
      sanitizeValue: exportHelpers.sanitizeValue,
      overwrite: false,
      resolvePayrollRow: (emp) => resolvePayrollRow(emp),
      resolvePayrollFields: resolveFormXIXKarnatakaPayrollFields,
    });
  } else if (
    exportRows.length > 0 &&
    typeof resolvePayrollRow === 'function' &&
    employees.length > 0 &&
    gridHasPayrollValues
  ) {
    const grossHdr = hdrs.find(isFormXIXKAGrossHeader);
    exportRows = exportRows.map((exportRow, index) => {
      if (grossHdr && getKarnatakaRowValueForHeader(exportRow, grossHdr) !== '') {
        return exportRow;
      }
      const emp = employees[index]?.Employee || employees[index]?.employee || employees[index] || null;
      const payrollRow = resolvePayrollRow(emp);
      if (!payrollRow || payrollRow.fetch_error) return exportRow;
      return applyFormXIXKarnatakaFooterFieldsToRow(exportRow, payrollRow, hdrs);
    });
  }

  const baseHeaderData = headerFormData && typeof headerFormData === 'object' ? { ...headerFormData } : {};
  const workbookArgs = {
    templateArrayBuffer,
    headersToUse: hdrs,
    parsedFormHeader,
    formFileName,
    headerFormData: baseHeaderData,
    resolvePayrollRow,
  };

  if (exportRows.length === 0) {
    return buildFormXIXKarnatakaWorkbookWithTemplateStyles({
      ...workbookArgs,
      mappedData: [],
      employeesOverride: employees,
    });
  }

  if (exportRows.length > 1 && canUseFormXIXKarnatakaFastExport(parsedFormHeader)) {
    try {
      const fastTemplate = await prepareFormXIXKarnatakaFastZipTemplate({
        templateArrayBuffer,
        parsedFormHeader,
        headerFormData: baseHeaderData,
      });
      const hasPatchTargets =
        fastTemplate.positions.employeeHeaderPositions.length > 0 ||
        fastTemplate.positions.tablePositions.length > 0;
      if (hasPatchTargets) {
        return buildFormXIXKarnatakaFastZipDownload({
          exportRows,
          fastTemplate,
          baseHeaderData,
          parsedFormHeader,
          formFileName,
          employees,
        });
      }
    } catch (fastZipErr) {
      console.warn('Form XIX Karnataka fast ZIP export failed, using standard export:', fastZipErr);
    }
  }

  const zip = new JSZip();
  const usedNames = new Map();
  for (let i = 0; i < exportRows.length; i += 1) {
    const emp = employees[i]?.Employee || employees[i]?.employee || employees[i] || null;
    const mergedHeaderData = buildEmployeeHeaderFormData(baseHeaderData, exportRows[i], emp);
    const { blob } = await buildFormXIXKarnatakaWorkbookWithTemplateStyles({
      ...workbookArgs,
      mappedData: [exportRows[i]],
      headerFormData: mergedHeaderData,
      employeesOverride: emp ? [emp] : [],
    });
    const bytes = new Uint8Array(await blob.arrayBuffer());
    const baseName = resolveEmployeeDownloadBaseName(exportRows[i], i);
    const count = usedNames.get(baseName) || 0;
    usedNames.set(baseName, count + 1);
    const suffix = count > 0 ? `_${count + 1}` : '';
    zip.file(`Form_XIX_Karnataka_${baseName}${suffix}.xlsx`, bytes);
    if (i > 0 && i % 15 === 0) {
      await new Promise((resolve) => setTimeout(resolve, 0));
    }
  }
  const zipBase = String(formFileName || parsedFormHeader?.title || 'Form_XIX_Karnataka')
    .replace(/\.xlsx?$/i, '')
    .replace(/[^a-zA-Z0-9._-]+/g, '_');
  return {
    blob: await zip.generateAsync({ type: 'blob', compression: 'STORE' }),
    fileName: `${zipBase}_Employees.zip`,
  };
}
