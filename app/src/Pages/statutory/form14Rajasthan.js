/** Rajasthan Form 14 — Record of the Hours of Work of Persons Employed [Rule 22(3)]. */

import ExcelJS from 'exceljs';
import { flattenPayrollEarningColumns, readPayrollScalar } from '../../utils/payrollEarnings';
import { ensureExcelJSDataRowsWithBorders } from '../../utils/excelTableBorders';

export const FORM_14_RJ_DISPLAY_TITLE = 'FORM 14';
export const FORM_14_RJ_DISPLAY_SUBTITLE = 'Record of the Hours of Work of Persons Employed';
export const FORM_14_RJ_DISPLAY_REFERENCE = 'Rule 22 - sub rule 3';
export const FORM_14_RJ_ACT_NAME = 'The Rajasthan Shops and Establishments Rules, 1959';
export const FORM_14_RJ_NOTICE_LINE =
  '(To be used only when Notice in Form 13 is exhibited)';
export const FORM_14_RJ_YOUNG_DEFAULT = 'Young';

/** Canonical Form 14 RJ table headers (template model — 2nd image). */
export const FORM_14_RJ_TABLE_HEADERS = [
  'Name of Persons Employed',
  'Whether young person or not',
  'Total hours worked during the month',
  '*Days on which overtime work is done and extent of such overtime on each day',
  'Extent of overtime worked during the month',
  'Extent of overtime worked previously during the month',
];

/** Official Form 14 RJ footnote under the overtime / days columns. */
export const FORM_14_RJ_FOOTER_NOTE =
  '*This column need not be filled by Commercial Establishments. In case of shops, Residential Hotels, Restaurants and Eating Houses and Theatres and other places of public amusement or entertainment, the extent of such overtime on each day shall be recorded in the days column against the employed person distinctively in red ink, indicating the time up to which such overtime work was taken from the employee';

/** Second Form 14 RJ note under the overtime footnote. */
export const FORM_14_RJ_DAY_ENTRIES_NOTE =
  'Note: Entries relating to any day must be made on that day';

const norm = (s) =>
  String(s || '')
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();

const stripDedupeSuffix = (h) =>
  String(h || '')
    .replace(/\s+\(\d+\)$/, '')
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

function buildContextBlob(formHeader, rowItem, fileName, sheetText, tableHeaders) {
  return [
    rowItem?.formName,
    rowItem?.FormName,
    rowItem?.description,
    rowItem?.Description,
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
}

/** Explicit Form 14 RJ filename / title markers (Arabic 14, not Roman XIV Employment Card). */
export function looksLikeForm14RajasthanFileName(text) {
  const s = String(text || '').toLowerCase();
  if (/form[\s._-]*xiv(?![a-z])/i.test(s)) return false;
  return /form[\s._-]*14[\s._-]*rj|\bform_14_rj\b|\b14_rj\b/.test(s);
}

/**
 * Rajasthan Shops Form 14 (Record of Hours of Work).
 * Must not match AP Form 14 child workers, CLRA Form XIV Employment Card,
 * or RJ Forms 11 / 12 / 15.
 */
export function isForm14RajasthanContext(
  formHeader,
  rowItem,
  fileName,
  sheetText = '',
  tableHeaders = []
) {
  const parts = buildContextBlob(formHeader, rowItem, fileName, sheetText, tableHeaders);

  if (/tamil\s*nadu|\b_tn\b/.test(parts)) return false;
  if (/child\s+worker|register\s+of\s+child|rule\s*86/.test(parts)) return false;
  if (/employment\s+card|form[\s._-]*xiv(?![a-z])/.test(parts) && !looksLikeForm14RajasthanFileName(parts)) {
    return false;
  }
  if (/form[\s._-]*11[\s._-]*rj|\bform_11_rj\b|\b11_rj\b/.test(parts)) return false;
  if (/form[\s._-]*12[\s._-]*rj|\bform_12_rj\b|\b12_rj\b/.test(parts)) return false;
  if (/form[\s._-]*15[\s._-]*rj|\bform_15_rj\b|\b15_rj\b/.test(parts)) return false;
  if (/andhra|\b_ap\b/.test(parts) && !/rajasthan|\b_rj\b/.test(parts)) return false;

  if (looksLikeForm14RajasthanFileName(parts)) return true;
  if (/record\s+of\s+(?:the\s+)?hours\s+of\s+work/.test(parts) && /rajasthan|\b_rj\b/.test(parts)) {
    return true;
  }
  if (
    /rajasthan/.test(parts) &&
    /form[\s._-]*14\b/.test(parts) &&
    (/rule\s*22\s*\(\s*3\s*\)|record\s+of\s+(?:the\s+)?hours|record\s+of\s+work/.test(parts) ||
      /shops\s+and\s+establishments/.test(parts))
  ) {
    return true;
  }
  return false;
}

export function enrichForm14RajasthanDisplayHeader(
  formHeader,
  fileName = '',
  item = null,
  tableHeaders = [],
  sheetText = ''
) {
  if (!isForm14RajasthanContext(formHeader, item, fileName, sheetText, tableHeaders)) {
    return formHeader;
  }
  const base = formHeader && typeof formHeader === 'object' ? { ...formHeader } : {};
  base.title = FORM_14_RJ_DISPLAY_TITLE;
  base.subtitle = FORM_14_RJ_DISPLAY_SUBTITLE;
  base.reference = FORM_14_RJ_DISPLAY_REFERENCE;
  base.actName = FORM_14_RJ_ACT_NAME;
  base.noticeLine = FORM_14_RJ_NOTICE_LINE;
  base.form14RJHoursOfWorkLayout = true;
  base.form11RJEmploymentLayout = false;
  base.form12RJEmploymentLayout = false;
  // Modal / preview: title stack centered (Excel model).
  base.headerTextAlign = 'center';
  base.form14RJCenterHeader = true;
  base.form14RJLeftAlignHeader = false;

  let fields = Array.isArray(base.fields) ? [...base.fields] : [];
  // Drop Form 11 RJ month/year keys if a prior misclassification added them.
  fields = fields.filter(
    (f) => f?.key !== 'form_11_rj_month' && f?.key !== 'form_11_rj_year'
  );
  const hasKey = (k) => fields.some((f) => f.key === k);
  if (!hasKey('form_14_rj_month')) {
    fields.push({ label: 'Month', value: '', key: 'form_14_rj_month' });
  }
  if (!hasKey('form_14_rj_year')) {
    fields.push({ label: 'Year', value: '', key: 'form_14_rj_year' });
  }
  base.fields = fields;
  return base;
}

export function isForm14RajasthanHoursOfWorkLayoutFormHeader(formHeader) {
  return !!formHeader?.form14RJHoursOfWorkLayout;
}

export function classifyForm14RJHeader(header) {
  const s = norm(stripDedupeSuffix(header));
  if (!s) return '';
  if (/name\s+of\s+persons?\s+employ/.test(s) || (s.includes('name') && s.includes('employ') && !/young/.test(s))) {
    return 'name';
  }
  if (/young\s+person/.test(s)) return 'young';
  if (/total\s+hours\s+worked/.test(s)) return 'totalHours';
  return '';
}

export function isForm14RajasthanNameHeader(header) {
  return classifyForm14RJHeader(header) === 'name';
}

export function isForm14RajasthanYoungPersonHeader(header) {
  return classifyForm14RJHeader(header) === 'young';
}

export function isForm14RajasthanTotalHoursHeader(header) {
  return classifyForm14RJHeader(header) === 'totalHours';
}

/** Whether young person or not — default "Young". */
export function resolveForm14RajasthanYoungPersonValue(_ageYears) {
  return FORM_14_RJ_YOUNG_DEFAULT;
}

export function isForm14RajasthanJunkTotalHoursValue(value) {
  const s = String(value ?? '').trim();
  if (!s) return false;
  if (/month|year|week|day\s*\(|experience|tenure/i.test(s) && !/^\d+(\.\d+)?$/.test(s)) {
    return true;
  }
  if (!/^\d+(\.\d+)?$/.test(s.replace(/,/g, ''))) return true;
  return false;
}

export function sanitizeForm14RajasthanTotalHoursValue(value) {
  const s = String(value ?? '').trim();
  if (!s || isForm14RajasthanJunkTotalHoursValue(s)) return '';
  const n = Number(s.replace(/,/g, ''));
  if (!Number.isFinite(n) || n < 0) return '';
  return String(Math.round(n * 100) / 100);
}

/** Total hours worked during the month = Paid_days × 8. */
export function computeForm14RajasthanTotalHours(paidDays) {
  if (paidDays === '' || paidDays == null) return '';
  const raw = String(paidDays).trim();
  if (!raw || /month|year|week|experience|tenure/i.test(raw)) return '';
  const n = Number(raw.replace(/,/g, ''));
  if (!Number.isFinite(n) || n < 0) return '';
  return String(Math.round(n * 8 * 100) / 100);
}

/** Read Paid_days (and aliases) from a Sample Payroll / Payroll table row. */
export function readForm14RajasthanPaidDays(payrollRow) {
  if (!payrollRow || typeof payrollRow !== 'object' || payrollRow.fetch_error) return '';
  const flat = flattenPayrollEarningColumns(payrollRow);
  const keys = [
    'Paid_days',
    'paid_days',
    'Paid Days',
    'paidDays',
    'PaidDays',
    'days_worked',
    'Days Worked',
    'daysWorked',
    'no_of_days_worked',
    'effective_paid_days',
  ];
  const patterns = [
    /^paid_days$/,
    /^paiddays$/,
    /^days_worked$/,
    /^daysworked$/,
    /^no_of_days_worked$/,
    /^effective_paid_days$/,
  ];
  const fromFlat = readPayrollScalar({ ...flat, ...payrollRow }, keys, patterns);
  if (fromFlat !== '' && fromFlat != null) {
    const raw = String(fromFlat).trim();
    if (/month|year|week|experience|tenure/i.test(raw)) return '';
    const num = Number(raw.replace(/,/g, ''));
    if (Number.isFinite(num) && num >= 0) return String(num);
  }
  return '';
}

function findForm14RJHeader(headers, role) {
  const hdrs = Array.isArray(headers) ? headers : [];
  return hdrs.find((h) => classifyForm14RJHeader(h) === role) || '';
}

function setForm14RajasthanRoleValue(row, headers, role, value) {
  if (!row || !role) return false;
  const hdrs = Array.isArray(headers) ? headers : [];
  let wrote = false;
  hdrs.forEach((h) => {
    if (classifyForm14RJHeader(h) === role) {
      row[h] = value;
      wrote = true;
    }
  });
  return wrote;
}

/**
 * Force Young default onto every employee row.
 * Total hours left for Paid_days pass (junk cleared here).
 * @returns {number} rows updated
 */
export function applyForm14RajasthanAutofillDefaultsToRows(mappedData, tableHeaders, options = {}) {
  if (!Array.isArray(mappedData) || mappedData.length === 0) return 0;
  const headers = Array.isArray(tableHeaders) ? tableHeaders : [];
  const sanitizeValue =
    typeof options.sanitizeValue === 'function'
      ? options.sanitizeValue
      : (v) => String(v ?? '').trim();

  let hits = 0;
  mappedData.forEach((row) => {
    if (!row || typeof row !== 'object') return;
    setForm14RajasthanRoleValue(row, headers, 'young', sanitizeValue(FORM_14_RJ_YOUNG_DEFAULT));

    const hoursHeader = findForm14RJHeader(headers, 'totalHours');
    if (hoursHeader) {
      const cur = String(row[hoursHeader] ?? '').trim();
      if (!cur || isForm14RajasthanJunkTotalHoursValue(cur)) {
        row[hoursHeader] = '';
      }
    }
    hits += 1;
  });
  return hits;
}

/**
 * Fill Total hours worked during the month ← Paid_days × 8.
 * Clears junk like "1 month(s)" when Paid_days is missing.
 * @returns {number} rows updated
 */
export function applyForm14RajasthanPaidDaysToRows(
  mappedData,
  employeesForMapping,
  tableHeaders,
  options = {}
) {
  if (!Array.isArray(mappedData) || mappedData.length === 0) return 0;
  const headers = Array.isArray(tableHeaders) ? tableHeaders : [];
  const hoursHeader = findForm14RJHeader(headers, 'totalHours');
  if (!hoursHeader) return 0;

  const resolvePayrollRow =
    typeof options.resolvePayrollRow === 'function' ? options.resolvePayrollRow : () => null;
  const unwrapEmp =
    typeof options.unwrapEmp === 'function'
      ? options.unwrapEmp
      : (empItem) => empItem && (empItem.Employee || empItem.employee || empItem);
  const sanitizeValue =
    typeof options.sanitizeValue === 'function'
      ? options.sanitizeValue
      : (v) => String(v ?? '').trim();
  const overwrite = options.overwrite !== false;

  let hits = 0;
  mappedData.forEach((row, idx) => {
    if (!row || typeof row !== 'object') return;
    const emp = unwrapEmp(employeesForMapping?.[idx]);
    const payrollRow = resolvePayrollRow(emp, row, idx);
    const paidDays = readForm14RajasthanPaidDays(payrollRow);
    const hours = computeForm14RajasthanTotalHours(paidDays);
    const cur = String(row[hoursHeader] ?? '').trim();
    if (hours === '') {
      if (cur && (overwrite || isForm14RajasthanJunkTotalHoursValue(cur))) {
        row[hoursHeader] = '';
        hits += 1;
      }
      return;
    }
    if (!overwrite && cur && !isForm14RajasthanJunkTotalHoursValue(cur)) return;
    row[hoursHeader] = sanitizeValue(hours);
    hits += 1;
  });
  return hits;
}

function excelCellValueToString(val) {
  if (val == null) return '';
  if (typeof val === 'string' || typeof val === 'number' || typeof val === 'boolean') return String(val);
  if (val instanceof Date) return val.toISOString();
  if (typeof val === 'object') {
    if (Array.isArray(val.richText)) return val.richText.map((rt) => rt?.text || '').join('');
    if (val.text != null) return String(val.text);
    if (val.result != null) return String(val.result);
  }
  return '';
}

function buildMergeTopLeftResolver(worksheet) {
  const merges = Array.isArray(worksheet?._merges)
    ? Object.values(worksheet._merges)
    : Array.isArray(worksheet?.model?.merges)
      ? worksheet.model.merges
      : [];
  const map = new Map();
  merges.forEach((m) => {
    const top = m?.top ?? m?.s?.r;
    const left = m?.left ?? m?.s?.c;
    const bottom = m?.bottom ?? m?.e?.r;
    const right = m?.right ?? m?.e?.c;
    if (top == null || left == null || bottom == null || right == null) return;
    for (let r = top; r <= bottom; r += 1) {
      for (let c = left; c <= right; c += 1) {
        map.set(`${r}:${c}`, { r: top, c: left });
      }
    }
  });
  return (r, c) => map.get(`${r}:${c}`) || { r, c };
}

function getForm14RajasthanRowValueForHeader(row, header) {
  if (!row || typeof row !== 'object') return '';
  if (Object.prototype.hasOwnProperty.call(row, header) && row[header] != null) {
    const direct = String(row[header]).trim();
    if (direct && direct !== 'false') return direct;
  }
  const role = classifyForm14RJHeader(header);
  if (role) {
    for (const [k, v] of Object.entries(row)) {
      if (String(k).startsWith('__')) continue;
      if (classifyForm14RJHeader(k) !== role) continue;
      if (v != null && String(v).trim() !== '' && String(v) !== 'false') return String(v).trim();
    }
  }
  const target = norm(stripDedupeSuffix(header));
  if (target) {
    for (const [k, v] of Object.entries(row)) {
      if (String(k).startsWith('__')) continue;
      if (norm(stripDedupeSuffix(k)) !== target) continue;
      if (v != null && String(v).trim() !== '' && String(v) !== 'false') return String(v).trim();
    }
  }
  return '';
}

export function applyForm14RajasthanMonthYearToHeaderData(headerData, fullMonth, year) {
  const out = headerData && typeof headerData === 'object' ? { ...headerData } : {};
  const monthVal = String(fullMonth || '').trim();
  const yearVal = String(year || '').trim();
  if (monthVal) {
    out.form_14_rj_month = monthVal;
    out.form_11_rj_month = out.form_11_rj_month || monthVal;
  }
  if (yearVal) {
    out.form_14_rj_year = yearVal;
    out.form_11_rj_year = out.form_11_rj_year || yearVal;
  }
  return out;
}

function writeForm14RajasthanMonthYearFields(worksheet, headerFormData, headerRowEnd = 20) {
  if (!worksheet || !headerFormData || typeof headerFormData !== 'object') return;
  const monthVal = String(
    headerFormData.form_14_rj_month ||
      headerFormData.form_11_rj_month ||
      headerFormData.form_x_month ||
      headerFormData.Month ||
      headerFormData.month ||
      ''
  ).trim();
  const yearVal = String(
    headerFormData.form_14_rj_year ||
      headerFormData.form_11_rj_year ||
      headerFormData.form_x_year ||
      headerFormData.Year ||
      headerFormData.year ||
      ''
  ).trim();
  if (!monthVal && !yearVal) return;

  const getMergeTopLeft = buildMergeTopLeftResolver(worksheet);
  const maxCols = Math.max(20, worksheet.columnCount || 20);
  for (let r = 1; r <= Math.max(1, headerRowEnd); r += 1) {
    for (let c = 1; c <= maxCols; c += 1) {
      const tl = getMergeTopLeft(r, c);
      if (tl.r !== r || tl.c !== c) continue;
      const cell = worksheet.getCell(tl.r, tl.c);
      const text = excelCellValueToString(cell?.value).trim();
      if (!text) continue;
      const lower = text.toLowerCase();

      // "Month---------" / "Month :" / "Month" → "Month : August"
      if (/^month\b/i.test(text) && !/\byear\b/i.test(text) && monthVal) {
        if (
          /^month\s*[-–—:_.\s]*$/i.test(text) ||
          /^month\s*[-–—:_]+/i.test(text) ||
          !/[a-z]{3,}/i.test(text.replace(/^month/i, '').replace(/[-–—:_.\s]/g, ''))
        ) {
          cell.value = `Month : ${monthVal}`;
          cell.alignment = {
            ...(cell.alignment || {}),
            wrapText: false,
            vertical: 'middle',
            horizontal: 'left',
          };
        }
        continue;
      }

      // "Year----------" / "Year :" / "Year" → "Year : 2026"
      if (/^year\b/i.test(text) && !/\bmonth\b/i.test(text) && yearVal) {
        if (
          /^year\s*[-–—:_.\s]*$/i.test(text) ||
          /^year\s*[-–—:_]+/i.test(text) ||
          !/\d{4}/.test(text)
        ) {
          cell.value = `Year : ${yearVal}`;
          cell.alignment = {
            ...(cell.alignment || {}),
            wrapText: false,
            vertical: 'middle',
            horizontal: 'left',
          };
        }
        continue;
      }

      // Combined "Month ———— Year ————" style line.
      if (/month/.test(lower) && /year/.test(lower)) {
        let next = text;
        if (monthVal) {
          next = next.replace(
            /(month\s*[-–—:_]*\s*)(?:[^y]*?)(?=\s*year|$)/i,
            `$1${monthVal} `
          );
        }
        if (yearVal) {
          next = next.replace(/(year\s*[-–—:_]*\s*).*$/i, `$1${yearVal}`);
        }
        if (next !== text) {
          cell.value = next.replace(/\s+/g, ' ').trim();
          cell.alignment = {
            ...(cell.alignment || {}),
            wrapText: false,
            vertical: 'middle',
            horizontal: 'left',
          };
        }
      }
    }
  }
}

/** Center the Form 14 title / act / notice block across the table width. */
export function applyForm14RajasthanTitleBlockCenter(
  worksheet,
  { headerRowEnd = 12, mergeFromCol = 2, mergeToCol = 7 } = {}
) {
  if (!worksheet) return;
  const getMergeTopLeft = buildMergeTopLeftResolver(worksheet);
  const maxCols = Math.max(20, worksheet.columnCount || 20);
  const end = Math.max(1, Number(headerRowEnd) || 12);
  const fromCol = Math.max(1, Number(mergeFromCol) || 2);
  const toCol = Math.max(fromCol, Number(mergeToCol) || fromCol + 5);

  const isTitleLine = (text) => {
    const lower = String(text || '').toLowerCase();
    return (
      /rajasthan\s+shops/.test(lower) ||
      /^form\s*14\b/.test(lower) ||
      /rule\s*22/.test(lower) ||
      /hours\s+of\s+work/.test(lower) ||
      /notice\s+in\s+form\s*13/.test(lower)
    );
  };

  const colToLetter = (n) => {
    let s = '';
    let num = n;
    while (num > 0) {
      const rem = (num - 1) % 26;
      s = String.fromCharCode(65 + rem) + s;
      num = Math.floor((num - 1) / 26);
    }
    return s;
  };

  for (let r = 1; r <= end; r += 1) {
    let titleText = '';
    let titleCol = 0;
    for (let c = 1; c <= maxCols; c += 1) {
      const tl = getMergeTopLeft(r, c);
      if (tl.r !== r || tl.c !== c) continue;
      const cell = worksheet.getCell(tl.r, tl.c);
      const text = excelCellValueToString(cell?.value).trim();
      if (!text || !isTitleLine(text)) continue;
      titleText = text;
      titleCol = tl.c;
      break;
    }
    if (!titleText || !titleCol) continue;

    // Clear / unmerge any existing merges covering this title row in the table band.
    {
      const merges = Array.isArray(worksheet?.model?.merges) ? [...worksheet.model.merges] : [];
      merges.forEach((range) => {
        const parts = String(range || '').split(':');
        if (parts.length !== 2) return;
        const start = parts[0].match(/^([A-Z]+)(\d+)$/i);
        const endM = parts[1].match(/^([A-Z]+)(\d+)$/i);
        if (!start || !endM) return;
        const r1 = parseInt(start[2], 10);
        const r2 = parseInt(endM[2], 10);
        if (r < r1 || r > r2) return;
        try {
          worksheet.unMergeCells(String(range));
        } catch (_) {
          /* ignore */
        }
      });
    }

    // Place title in the leftmost table column and merge across the table for true center.
    for (let c = fromCol; c <= toCol; c += 1) {
      const cell = worksheet.getCell(r, c);
      if (c === fromCol) cell.value = titleText;
      else if (excelCellValueToString(cell?.value).trim() === titleText) cell.value = '';
    }
    // Also clear the old title cell if it sat outside the merge band (e.g. column F).
    if (titleCol < fromCol || titleCol > toCol) {
      try {
        worksheet.getCell(r, titleCol).value = '';
      } catch (_) {
        /* ignore */
      }
    }

    const rangeAddr = `${colToLetter(fromCol)}${r}:${colToLetter(toCol)}${r}`;
    try {
      worksheet.mergeCells(rangeAddr);
    } catch (_) {
      /* already merged / overlap */
    }
    const anchor = worksheet.getCell(r, fromCol);
    anchor.value = titleText;
    anchor.alignment = {
      horizontal: 'center',
      vertical: 'middle',
      wrapText: false,
    };
    try {
      anchor.style = {
        ...(anchor.style && typeof anchor.style === 'object' ? anchor.style : {}),
        alignment: { horizontal: 'center', vertical: 'middle', wrapText: false },
      };
    } catch (_) {
      /* ignore */
    }
  }
}

/** @deprecated Use applyForm14RajasthanTitleBlockCenter — kept as alias. */
export function applyForm14RajasthanHeaderBlockLeftAlign(worksheet, headerRowEnd = 12) {
  applyForm14RajasthanTitleBlockCenter(worksheet, { headerRowEnd });
}

/** True when a cell looks like data (Young/Nil/name/hours) instead of a real column title. */
export function isForm14RajasthanCorruptHeaderLabel(text) {
  const s = String(text || '').trim();
  if (!s) return true;
  if (/^nil$/i.test(s)) return true;
  if (/^young$/i.test(s)) return true;
  if (/^\d+(\.\d+)?$/.test(s.replace(/,/g, ''))) return true;
  if (/^month\b|^year\b/i.test(s)) return true;
  if (/description\s+of\s+department/i.test(s)) return true;
  // Real Form 14 RJ headers always include one of these phrases.
  if (
    /name\s+of\s+persons?\s+employ/i.test(s) ||
    /young\s+person/i.test(s) ||
    /total\s+hours\s+worked/i.test(s) ||
    /days\s+on\s+which\s+overtime/i.test(s) ||
    /extent\s+of\s+overtime/i.test(s)
  ) {
    return false;
  }
  // Anything else (employee name, junk) is not a model header.
  return true;
}

/** Ensure the table header row matches the template model (never Young/Nil/data). */
export function ensureForm14RajasthanTableHeaderLabels(worksheet, layout) {
  if (!worksheet || !layout) return;
  const headerRow = Number(layout.headerRow) || 0;
  if (headerRow < 1) return;
  const row = worksheet.getRow(headerRow);
  if (row) row.height = Math.max(Number(row.height) || 0, 48);

  let cols = Array.isArray(layout.templateCols) ? [...layout.templateCols] : [];
  const startCol =
    Number(layout.startCol) ||
    (cols.length ? Math.min(...cols.map((c) => c.col)) : 2);

  // Always keep the 6 model columns from the official template.
  if (cols.length < FORM_14_RJ_TABLE_HEADERS.length) {
    const existing = new Set(cols.map((c) => c.col));
    for (let i = 0; i < FORM_14_RJ_TABLE_HEADERS.length; i += 1) {
      const col = startCol + i;
      if (existing.has(col)) continue;
      cols.push({
        col,
        label: FORM_14_RJ_TABLE_HEADERS[i],
        role: classifyForm14RJHeader(FORM_14_RJ_TABLE_HEADERS[i]),
      });
      existing.add(col);
    }
    cols = cols.sort((a, b) => a.col - b.col).slice(0, FORM_14_RJ_TABLE_HEADERS.length);
    layout.templateCols = cols;
  }

  cols.slice(0, FORM_14_RJ_TABLE_HEADERS.length).forEach(({ col }, idx) => {
    const modelLabel = FORM_14_RJ_TABLE_HEADERS[idx];
    const cell = worksheet.getCell(headerRow, col);
    const current = excelCellValueToString(cell?.value).trim();
    // Always restore when blank/corrupt (Young, Nil, employee name, hours).
    if (!current || isForm14RajasthanCorruptHeaderLabel(current) || norm(current) !== norm(modelLabel)) {
      // Prefer exact model wording so download matches the 2nd image.
      cell.value = modelLabel;
    }
    cell.alignment = {
      ...(cell.alignment || {}),
      horizontal: 'left',
      vertical: 'middle',
      wrapText: true,
    };
    if (cols[idx]) {
      cols[idx].label = modelLabel;
      cols[idx].role = classifyForm14RJHeader(modelLabel);
    }
  });
  layout.templateCols = cols.slice(0, FORM_14_RJ_TABLE_HEADERS.length);
}

/** Widen Form 14 RJ columns so names / Young / hours / Nil fit like a readable register. */
export function applyForm14RajasthanColumnLayout(worksheet, templateCols = [], dataStartRow = 1, dataRowCount = 0) {
  if (!worksheet || !Array.isArray(templateCols) || templateCols.length === 0) return;

  const widthForRole = (role, label) => {
    if (role === 'name') return 32;
    if (role === 'young') return 24;
    if (role === 'totalHours') return 28;
    const s = norm(label);
    if (/overtime/.test(s) && /each\s+day|days\s+on\s+which/.test(s)) return 30;
    if (/overtime/.test(s)) return 26;
    if (/^nil$/i.test(String(label || '').trim())) return 14;
    return 20;
  };

  const cols = templateCols.map((c) => c.col).filter((n) => Number.isFinite(n) && n > 0);
  const firstCol = Math.min(...cols);
  const lastCol = Math.max(...cols);

  templateCols.forEach(({ col, role, label }) => {
    const column = worksheet.getColumn(col);
    if (!column) return;
    const target = widthForRole(role, label);
    const current = Number(column.width) || 0;
    column.width = Math.max(current, target);
  });

  // Widen every column in the table band (covers blank/Nil leaf cols).
  for (let c = firstCol; c <= lastCol; c += 1) {
    const column = worksheet.getColumn(c);
    if (!column) continue;
    column.width = Math.max(Number(column.width) || 0, 14);
  }

  // Labels like Description of Department / Month / Year sit just left of the table.
  for (let c = Math.max(1, firstCol - 2); c < firstCol; c += 1) {
    const column = worksheet.getColumn(c);
    if (column) column.width = Math.max(Number(column.width) || 0, 26);
  }

  // Readable row height for employee rows (avoid squashed / overlapping names).
  const rowCount = Math.max(0, Number(dataRowCount) || 0);
  for (let i = 0; i < rowCount; i += 1) {
    const row = worksheet.getRow(dataStartRow + i);
    if (!row) continue;
    row.height = Math.max(Number(row.height) || 0, 24);
    for (let c = firstCol; c <= lastCol; c += 1) {
      const cell = worksheet.getCell(dataStartRow + i, c);
      cell.alignment = {
        ...(cell.alignment || {}),
        wrapText: false,
        vertical: 'middle',
        horizontal: 'left',
      };
    }
  }
}

/**
 * Locate Form 14 RJ table header band + leaf columns from the ExcelJS template.
 * Prefers "Name of Persons Employed" / "Whether young person" anchors.
 */
export function detectForm14RajasthanTableLayout(worksheet, hints = {}) {
  if (!worksheet) return null;
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

  const maxScanRows = Math.max(40, (worksheet.rowCount || 0) + 5);
  // Prefer a row that still has the official "Name of Persons Employed" label.
  if (headerRow < 1) {
    for (let r = 1; r <= maxScanRows; r += 1) {
      for (let c = 1; c <= 40; c += 1) {
        const t = norm(getMergedAwareCellText(r, c));
        if (/name\s+of\s+persons?\s+employ/.test(t)) {
          headerRow = r;
          startCol = c;
          break;
        }
      }
      if (headerRow > 0) break;
    }
  }
  if (headerRow < 1) {
    for (let r = 1; r <= maxScanRows; r += 1) {
      for (let c = 1; c <= 40; c += 1) {
        const t = norm(getMergedAwareCellText(r, c));
        if (/young\s+person/.test(t) || /total\s+hours\s+worked/.test(t)) {
          headerRow = r;
          startCol = Math.max(1, c - 1);
          break;
        }
      }
      if (headerRow > 0) break;
    }
  }
  // Corrupted draft: headers became Young/Nil — anchor under Month / Description row.
  if (headerRow < 1) {
    for (let r = 1; r <= maxScanRows; r += 1) {
      for (let c = 1; c <= 20; c += 1) {
        const t = norm(getMergedAwareCellText(r, c));
        if (/description\s+of\s+department/.test(t) || /^month\b/.test(t)) {
          headerRow = r + (/^month\b/.test(t) ? 2 : 3);
          startCol = c;
          break;
        }
      }
      if (headerRow > 0) break;
    }
  }
  if (headerRow < 1) return null;

  // Prefer the leaf header row when a stacked band sits above day numbers / rest interval.
  let leafRow = headerRow;
  for (let r = headerRow; r <= Math.min(headerRow + 4, maxScanRows); r += 1) {
    let youngHits = 0;
    let nameHits = 0;
    let hoursHits = 0;
    for (let c = startCol; c <= startCol + 40; c += 1) {
      const role = classifyForm14RJHeader(getMergedAwareCellText(r, c));
      if (role === 'young') youngHits += 1;
      if (role === 'name') nameHits += 1;
      if (role === 'totalHours') hoursHits += 1;
    }
    if (nameHits && (youngHits || hoursHits)) {
      leafRow = r;
      break;
    }
  }

  let templateCols = [];
  const seenCols = new Set();
  for (let c = startCol; c <= startCol + 50; c += 1) {
    const tl = getMergeTopLeft(leafRow, c);
    if (seenCols.has(tl.c)) continue;
    seenCols.add(tl.c);
    const label =
      getMergedAwareCellText(leafRow, tl.c) ||
      getMergedAwareCellText(headerRow, tl.c);
    if (!label) {
      if (templateCols.length >= 4) break;
      continue;
    }
    templateCols.push({
      col: tl.c,
      label: stripDedupeSuffix(label),
      role: classifyForm14RJHeader(label),
    });
    if (templateCols.length >= FORM_14_RJ_TABLE_HEADERS.length) break;
  }

  // If headers were wiped into Young/Nil/names, rebuild the 6 model columns.
  const corruptCount = templateCols.filter((c) =>
    isForm14RajasthanCorruptHeaderLabel(c.label)
  ).length;
  // Prefer the leftmost column that actually has table content (templates usually start at B).
  let contentStartCol = startCol;
  for (let c = 1; c <= 12; c += 1) {
    if (getMergedAwareCellText(leafRow, c) || getMergedAwareCellText(headerRow, c)) {
      contentStartCol = c;
      break;
    }
  }
  if (
    templateCols.length < 2 ||
    corruptCount >= Math.ceil(Math.max(templateCols.length, 1) / 2)
  ) {
    startCol = contentStartCol;
    templateCols = FORM_14_RJ_TABLE_HEADERS.map((label, idx) => ({
      col: startCol + idx,
      label,
      role: classifyForm14RJHeader(label),
    }));
  } else {
    // Pad to 6 columns using model labels.
    while (templateCols.length < FORM_14_RJ_TABLE_HEADERS.length) {
      const idx = templateCols.length;
      const lastCol = templateCols[templateCols.length - 1]?.col || startCol - 1;
      templateCols.push({
        col: lastCol + 1,
        label: FORM_14_RJ_TABLE_HEADERS[idx],
        role: classifyForm14RJHeader(FORM_14_RJ_TABLE_HEADERS[idx]),
      });
    }
    templateCols = templateCols.slice(0, FORM_14_RJ_TABLE_HEADERS.length).map((c, idx) => {
      if (isForm14RajasthanCorruptHeaderLabel(c.label)) {
        return {
          ...c,
          label: FORM_14_RJ_TABLE_HEADERS[idx],
          role: classifyForm14RJHeader(FORM_14_RJ_TABLE_HEADERS[idx]),
        };
      }
      return c;
    });
    startCol = templateCols[0]?.col || startCol;
  }

  // Bottom of any vertical merges covering the header cells (tall wrapped headers).
  let headerBandEnd = leafRow;
  {
    const merges = Array.isArray(worksheet?.model?.merges) ? worksheet.model.merges : [];
    templateCols.forEach(({ col }) => {
      merges.forEach((range) => {
        const parts = String(range || '').split(':');
        if (parts.length !== 2) return;
        const start = parts[0].match(/^([A-Z]+)(\d+)$/i);
        const end = parts[1].match(/^([A-Z]+)(\d+)$/i);
        if (!start || !end) return;
        const colToNum = (letters) =>
          String(letters)
            .toUpperCase()
            .split('')
            .reduce((n, ch) => n * 26 + (ch.charCodeAt(0) - 64), 0);
        const c1 = colToNum(start[1]);
        const c2 = colToNum(end[1]);
        const r1 = parseInt(start[2], 10);
        const r2 = parseInt(end[2], 10);
        if (col >= c1 && col <= c2 && leafRow >= r1 && leafRow <= r2) {
          headerBandEnd = Math.max(headerBandEnd, r2);
        }
      });
    });
  }

  let dataStartRow =
    hints.parsedDataStartIndex != null && hints.parsedDataStartIndex >= 0
      ? hints.parsedDataStartIndex + 1
      : headerBandEnd + 1;

  // Never treat the header band as data — bad modal hints wipe model headers.
  if (dataStartRow <= headerBandEnd) dataStartRow = headerBandEnd + 1;

  // Skip a 1…N index row under the header band when present.
  {
    let numHits = 0;
    for (let i = 0; i < Math.min(12, templateCols.length); i += 1) {
      const t = String(getMergedAwareCellText(dataStartRow, templateCols[i].col) || '')
        .replace(/\s+/g, '')
        .trim();
      if (t === String(i + 1)) numHits += 1;
    }
    if (numHits >= 6) dataStartRow += 1;
  }

  // Skip another header-looking row (wrapped duplicate band).
  {
    const probe = getMergedAwareCellText(dataStartRow, templateCols[0]?.col || startCol);
    if (
      /name\s+of\s+persons?\s+employ/i.test(probe) ||
      /young\s+person/i.test(probe) ||
      /total\s+hours\s+worked/i.test(probe)
    ) {
      dataStartRow += 1;
      headerBandEnd = Math.max(headerBandEnd, dataStartRow - 1);
    }
  }

  if (dataStartRow <= headerBandEnd) dataStartRow = headerBandEnd + 1;

  return {
    headerRow: leafRow,
    dataStartRow,
    templateCols,
    startCol,
    headerBandEnd,
  };
}

function looksLikeForm14RajasthanFooterNote(text) {
  const s = String(text || '').toLowerCase();
  if (!s) return false;
  return (
    /need\s+not\s+be\s+filled/.test(s) ||
    (/commercial\s+establishments/.test(s) && /red\s+ink/.test(s)) ||
    (/overtime/.test(s) && /red\s+ink/.test(s) && /days\s+column/.test(s)) ||
    (/^\*\s*this\s+column/.test(s) && /commercial|overtime|red\s+ink|shops/.test(s))
  );
}

function looksLikeForm14RajasthanFooterFragment(text) {
  const s = String(text || '').trim();
  if (!s) return false;
  if (looksLikeForm14RajasthanFooterNote(s)) return true;
  const lower = s.toLowerCase();
  return (
    /^\*\s*this\s+column/i.test(s) ||
    /need\s+not\s+be\s+filled/i.test(lower) ||
    /commercial\s+establishments/i.test(lower) ||
    /residential\s+hotels/i.test(lower) ||
    /places\s+of\s+public\s+amusement/i.test(lower) ||
    /distinctively\s+in\s+red\s+ink/i.test(lower) ||
    /time\s+up\s+to\s+which\s+such\s+overtime/i.test(lower)
  );
}

/** Collapse wrapped / multi-cell footnote pieces into one continuous sentence. */
export function normalizeForm14RajasthanFooterNoteText(text) {
  const raw = String(text || '')
    .replace(/\r\n/g, '\n')
    .replace(/\r/g, '\n')
    .replace(/\u00a0/g, ' ')
    .trim();
  if (!raw) return FORM_14_RJ_FOOTER_NOTE;
  const single = raw
    .split(/\n+/)
    .map((part) => String(part || '').trim())
    .filter(Boolean)
    .join(' ')
    .replace(/\s+/g, ' ')
    .trim();
  return single || FORM_14_RJ_FOOTER_NOTE;
}

export function normalizeForm14RajasthanDayEntriesNoteText(text) {
  const raw = String(text || '')
    .replace(/\r\n/g, '\n')
    .replace(/\r/g, '\n')
    .replace(/\u00a0/g, ' ')
    .trim();
  if (!raw) return FORM_14_RJ_DAY_ENTRIES_NOTE;
  const single = raw
    .split(/\n+/)
    .map((part) => String(part || '').trim())
    .filter(Boolean)
    .join(' ')
    .replace(/\s+/g, ' ')
    .trim();
  if (/note\s*:\s*entries\s+relating\s+to\s+any\s+day/i.test(single)) {
    return FORM_14_RJ_DAY_ENTRIES_NOTE;
  }
  return single || FORM_14_RJ_DAY_ENTRIES_NOTE;
}

function looksLikeForm14RajasthanDayEntriesNote(text) {
  const s = String(text || '').toLowerCase();
  if (!s) return false;
  return (
    /note\s*:\s*entries\s+relating\s+to\s+any\s+day/.test(s) ||
    (/entries\s+relating\s+to\s+any\s+day/.test(s) && /made\s+on\s+that\s+day/.test(s))
  );
}

function excelColLettersToNumber(letters) {
  const s = String(letters || '').toUpperCase();
  let n = 0;
  for (let i = 0; i < s.length; i += 1) {
    n = n * 26 + (s.charCodeAt(i) - 64);
  }
  return n;
}

/** Capture template footnote (text + style + merge) before the data band is cleared. */
export function findForm14RajasthanFooterNote(worksheet, dataStartRow = 1) {
  if (!worksheet) return null;
  const getMergeTopLeft = buildMergeTopLeftResolver(worksheet);
  const maxRows = Math.max(worksheet.rowCount || 0, dataStartRow + 40);
  const maxCols = Math.max(20, worksheet.columnCount || 20);
  for (let r = Math.max(1, dataStartRow); r <= maxRows; r += 1) {
    for (let c = 1; c <= maxCols; c += 1) {
      const tl = getMergeTopLeft(r, c);
      if (tl.r !== r || tl.c !== c) continue;
      const cell = worksheet.getCell(r, c);
      const text = excelCellValueToString(cell?.value).trim();
      if (!looksLikeForm14RajasthanFooterFragment(text)) continue;

      // Stitch adjacent rows if the template note was split across lines/cells.
      const parts = [text];
      for (let rr = r + 1; rr <= Math.min(r + 6, maxRows); rr += 1) {
        let rowPiece = '';
        for (let cc = 1; cc <= maxCols; cc += 1) {
          const piece = excelCellValueToString(worksheet.getCell(rr, cc)?.value).trim();
          if (!piece) continue;
          if (/^note\s*:/i.test(piece)) break;
          if (!looksLikeForm14RajasthanFooterFragment(piece) && parts.join(' ').length > 40) {
            // Stop once we leave the footnote band.
            if (!/red\s+ink|overtime|employee|entertainment|shops/i.test(piece)) break;
          }
          if (looksLikeForm14RajasthanFooterFragment(piece) || /red\s+ink|overtime|employee/i.test(piece)) {
            rowPiece = piece;
            break;
          }
        }
        if (!rowPiece) break;
        if (/^note\s*:/i.test(rowPiece)) break;
        parts.push(rowPiece);
      }

      let mergeEndCol = c;
      const merges = worksheet?.model?.merges;
      if (Array.isArray(merges)) {
        for (const range of merges) {
          const rangeParts = String(range || '').split(':');
          if (rangeParts.length !== 2) continue;
          const start = rangeParts[0].match(/^([A-Z]+)(\d+)$/i);
          const end = rangeParts[1].match(/^([A-Z]+)(\d+)$/i);
          if (!start || !end) continue;
          if (parseInt(start[2], 10) !== r) continue;
          const c1 = excelColLettersToNumber(start[1]);
          const c2 = excelColLettersToNumber(end[1]);
          if (c1 <= c && c2 >= c) mergeEndCol = Math.max(mergeEndCol, c2);
        }
      }

      return {
        row: r,
        col: Math.min(c, 1),
        text: normalizeForm14RajasthanFooterNoteText(parts.join(' ')),
        mergeEndCol: Math.max(mergeEndCol, c + 5),
        font: cell.font ? { ...cell.font } : { italic: true, size: 9, name: 'Times New Roman' },
        alignment: { wrapText: true, vertical: 'top', horizontal: 'left' },
      };
    }
  }
  return null;
}

export function writeForm14RajasthanFooterNote(worksheet, options = {}) {
  if (!worksheet) return null;
  const {
    afterRow,
    startCol = 1,
    endCol = Math.max(8, startCol + 7),
    footer = null,
  } = options;
  const row =
    Number.isFinite(Number(afterRow)) && Number(afterRow) > 0
      ? Number(afterRow) + 1
      : footer?.row || null;
  if (!row) return null;

  const col = Math.max(1, startCol);
  const mergeTo = Math.max(col, endCol, footer?.mergeEndCol || col);
  const text = normalizeForm14RajasthanFooterNoteText(
    footer?.text || FORM_14_RJ_FOOTER_NOTE
  );

  // Clear old footnote fragments across several rows so only one continuous line remains.
  const maxRows = Math.max(worksheet.rowCount || 0, row + 10);
  for (let r = row; r <= maxRows; r += 1) {
    for (let c = col; c <= mergeTo; c += 1) {
      const existing = excelCellValueToString(worksheet.getCell(r, c)?.value).trim();
      if (
        looksLikeForm14RajasthanFooterFragment(existing) ||
        /red\s+ink|commercial\s+establishments|need\s+not\s+be\s+filled/i.test(existing)
      ) {
        try {
          worksheet.unMergeCells(r, col, r, mergeTo);
        } catch (_) {
          /* ignore */
        }
        worksheet.getCell(r, c).value = '';
      }
    }
  }

  try {
    worksheet.unMergeCells(row, col, row, mergeTo);
  } catch (_) {
    /* ignore */
  }
  try {
    worksheet.mergeCells(row, col, row, mergeTo);
  } catch (_) {
    /* ignore if already merged / invalid */
  }

  const cell = worksheet.getCell(row, col);
  cell.value = text;
  cell.font = footer?.font || { italic: true, size: 9, name: 'Times New Roman' };
  // One continuous note across the full table width (template-style single paragraph).
  cell.alignment = {
    wrapText: true,
    vertical: 'top',
    horizontal: 'left',
    shrinkToFit: false,
  };
  const targetRow = worksheet.getRow(row);
  if (targetRow) {
    // Wide merged row — keep a modest height; Excel wraps within the merge like the model.
    targetRow.height = Math.max(Number(targetRow.height) || 0, 36);
  }
  return row;
}

/** Capture "Note: Entries relating to any day..." from the template. */
export function findForm14RajasthanDayEntriesNote(worksheet, dataStartRow = 1) {
  if (!worksheet) return null;
  const getMergeTopLeft = buildMergeTopLeftResolver(worksheet);
  const maxRows = Math.max(worksheet.rowCount || 0, dataStartRow + 50);
  const maxCols = Math.max(20, worksheet.columnCount || 20);
  for (let r = Math.max(1, dataStartRow); r <= maxRows; r += 1) {
    for (let c = 1; c <= maxCols; c += 1) {
      const tl = getMergeTopLeft(r, c);
      if (tl.r !== r || tl.c !== c) continue;
      const cell = worksheet.getCell(r, c);
      const text = excelCellValueToString(cell?.value).trim();
      if (!looksLikeForm14RajasthanDayEntriesNote(text)) continue;

      const parts = [text];
      for (let rr = r + 1; rr <= Math.min(r + 3, maxRows); rr += 1) {
        let rowPiece = '';
        for (let cc = 1; cc <= maxCols; cc += 1) {
          const piece = excelCellValueToString(worksheet.getCell(rr, cc)?.value).trim();
          if (!piece) continue;
          if (looksLikeForm14RajasthanFooterFragment(piece)) break;
          if (/made\s+on\s+that\s+day|entries\s+relating|any\s+day/i.test(piece)) {
            rowPiece = piece;
            break;
          }
        }
        if (!rowPiece) break;
        parts.push(rowPiece);
      }

      let mergeEndCol = c;
      const merges = worksheet?.model?.merges;
      if (Array.isArray(merges)) {
        for (const range of merges) {
          const rangeParts = String(range || '').split(':');
          if (rangeParts.length !== 2) continue;
          const start = rangeParts[0].match(/^([A-Z]+)(\d+)$/i);
          const end = rangeParts[1].match(/^([A-Z]+)(\d+)$/i);
          if (!start || !end) continue;
          if (parseInt(start[2], 10) !== r) continue;
          const c1 = excelColLettersToNumber(start[1]);
          const c2 = excelColLettersToNumber(end[1]);
          if (c1 <= c && c2 >= c) mergeEndCol = Math.max(mergeEndCol, c2);
        }
      }

      return {
        row: r,
        col: Math.min(c, 1),
        text: normalizeForm14RajasthanDayEntriesNoteText(parts.join(' ')),
        mergeEndCol: Math.max(mergeEndCol, c + 5),
        font: cell.font ? { ...cell.font } : { italic: true, size: 9, name: 'Times New Roman' },
      };
    }
  }
  return null;
}

export function writeForm14RajasthanDayEntriesNote(worksheet, options = {}) {
  if (!worksheet) return null;
  const {
    afterRow,
    startCol = 1,
    endCol = Math.max(8, startCol + 7),
    note = null,
  } = options;
  const row =
    Number.isFinite(Number(afterRow)) && Number(afterRow) > 0
      ? Number(afterRow) + 1
      : note?.row || null;
  if (!row) return null;

  const col = Math.max(1, startCol);
  const mergeTo = Math.max(col, endCol, note?.mergeEndCol || col);
  const text = normalizeForm14RajasthanDayEntriesNoteText(
    note?.text || FORM_14_RJ_DAY_ENTRIES_NOTE
  );

  const maxRows = Math.max(worksheet.rowCount || 0, row + 8);
  for (let r = row; r <= maxRows; r += 1) {
    for (let c = col; c <= mergeTo; c += 1) {
      const existing = excelCellValueToString(worksheet.getCell(r, c)?.value).trim();
      if (looksLikeForm14RajasthanDayEntriesNote(existing)) {
        try {
          worksheet.unMergeCells(r, col, r, mergeTo);
        } catch (_) {
          /* ignore */
        }
        worksheet.getCell(r, c).value = '';
      }
    }
  }

  try {
    worksheet.unMergeCells(row, col, row, mergeTo);
  } catch (_) {
    /* ignore */
  }
  try {
    worksheet.mergeCells(row, col, row, mergeTo);
  } catch (_) {
    /* ignore */
  }

  const cell = worksheet.getCell(row, col);
  cell.value = text;
  cell.font = note?.font || { italic: true, size: 9, name: 'Times New Roman' };
  cell.alignment = {
    wrapText: false,
    vertical: 'middle',
    horizontal: 'left',
    shrinkToFit: false,
  };
  const targetRow = worksheet.getRow(row);
  if (targetRow) {
    targetRow.height = Math.max(Number(targetRow.height) || 0, 18);
  }
  return row;
}

/**
 * ExcelJS export — load the original Form 14 RJ template (not a SheetJS round-trip)
 * so Excel does not repair/corrupt merges on open.
 */
export async function buildForm14RajasthanWorkbookWithTemplateStyles({
  templateArrayBuffer,
  mappedData,
  headersToUse,
  parsedFormHeader,
  formFileName,
  headerFormData,
  parsedHeaderRowIndex,
  parsedDataStartIndex,
  parsedTableStartCol,
}) {
  if (!templateArrayBuffer) {
    throw new Error('Form 14 RJ template buffer is required.');
  }
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(templateArrayBuffer);
  const worksheet = workbook.worksheets[0];
  if (!worksheet) throw new Error('Template worksheet not found.');

  const layout = detectForm14RajasthanTableLayout(worksheet, {
    parsedHeaderRowIndex,
    parsedDataStartIndex,
    parsedTableStartCol,
  });
  if (!layout) throw new Error('Could not locate Rajasthan Form 14 table header row.');

  const { dataStartRow, templateCols, headerRow, headerBandEnd } = layout;
  const protectedHeaderEnd = Math.max(headerRow || 1, headerBandEnd || headerRow || 1);
  const tableColMin = Math.min(...templateCols.map(({ col }) => col));
  const tableColMax = Math.max(...templateCols.map(({ col }) => col));
  const capturedFooter = findForm14RajasthanFooterNote(worksheet, dataStartRow);
  const capturedDayEntriesNote = findForm14RajasthanDayEntriesNote(worksheet, dataStartRow);
  writeForm14RajasthanMonthYearFields(
    worksheet,
    headerFormData,
    Math.max(1, protectedHeaderEnd - 1)
  );
  applyForm14RajasthanTitleBlockCenter(worksheet, {
    headerRowEnd: Math.max(1, protectedHeaderEnd - 1),
    mergeFromCol: tableColMin,
    mergeToCol: Math.max(tableColMax, tableColMin + 5),
  });
  ensureForm14RajasthanTableHeaderLabels(worksheet, layout);

  const headers = Array.isArray(headersToUse) && headersToUse.length > 0
    ? headersToUse
    : templateCols.map((c) => c.label);
  const rows = (Array.isArray(mappedData) ? mappedData : []).filter(
    (row) => row && typeof row === 'object'
  );

  applyForm14RajasthanAutofillDefaultsToRows(rows, headers);

  const safeDataStart = Math.max(dataStartRow, protectedHeaderEnd + 1);
  const lastDataRow = safeDataStart + Math.max(rows.length, 1) - 1;
  // Clear only through the data band (or just above a known footer), never wipe the footnote forever.
  const clearToRow = capturedFooter?.row
    ? Math.max(lastDataRow, Math.min(capturedFooter.row - 1, safeDataStart + rows.length + 10))
    : Math.max(safeDataStart + rows.length + 2, safeDataStart + 8);

  // Unmerge body cells in the data band so Excel does not repair overlapping merges.
  {
    const merges = worksheet?.model?.merges;
    if (Array.isArray(merges) && merges.length > 0) {
      const toRemove = [];
      merges.forEach((range) => {
        const parts = String(range || '').split(':');
        if (parts.length !== 2) return;
        const start = parts[0].match(/^([A-Z]+)(\d+)$/i);
        const end = parts[1].match(/^([A-Z]+)(\d+)$/i);
        if (!start || !end) return;
        const r1 = parseInt(start[2], 10);
        const r2 = parseInt(end[2], 10);
        // Keep title + table header merges intact.
        if (r1 <= protectedHeaderEnd) return;
        if (r2 >= safeDataStart && r1 <= clearToRow) toRemove.push(String(range));
      });
      toRemove.forEach((range) => {
        try {
          worksheet.unMergeCells(range);
        } catch (_) {
          /* ignore invalid merge ranges */
        }
      });
    }
  }

  for (let r = safeDataStart; r <= clearToRow; r += 1) {
    for (let c = tableColMin; c <= tableColMax; c += 1) {
      worksheet.getCell(r, c).value = '';
    }
  }

  rows.forEach((row, idx) => {
    const targetRowNum = safeDataStart + idx;
    const targetRow = worksheet.getRow(targetRowNum);
    if (targetRow) targetRow.height = Math.max(Number(targetRow.height) || 0, 24);
    templateCols.forEach(({ col, label, role }) => {
      let val = getForm14RajasthanRowValueForHeader(row, label);
      if ((!val || !String(val).trim()) && role) {
        for (const h of headers) {
          if (classifyForm14RJHeader(h) === role) {
            const viaHeader = getForm14RajasthanRowValueForHeader(row, h);
            if (viaHeader) {
              val = viaHeader;
              break;
            }
          }
        }
      }
      if (role === 'young' && (!val || !String(val).trim())) {
        val = FORM_14_RJ_YOUNG_DEFAULT;
      }
      // Blank overtime / extent columns → Nil (matches register practice).
      if (
        (!val || !String(val).trim()) &&
        role !== 'name' &&
        role !== 'young' &&
        role !== 'totalHours'
      ) {
        val = 'Nil';
      }
      const cell = worksheet.getCell(targetRowNum, col);
      // Always write as text so finalize won't right-align hours as numbers.
      const display =
        val == null || String(val).trim() === ''
          ? ''
          : role === 'totalHours'
            ? String(val).replace(/,/g, '').trim()
            : String(val);
      cell.value = display;
      const leftAlign = {
        horizontal: 'left',
        vertical: 'middle',
        wrapText: false,
      };
      cell.alignment = leftAlign;
      // ExcelJS sometimes keeps a shared style.alignment — set both.
      try {
        cell.style = {
          ...(cell.style && typeof cell.style === 'object' ? cell.style : {}),
          alignment: { ...leftAlign },
        };
      } catch (_) {
        /* ignore */
      }
    });
  });

  if (rows.length > 0) {
    ensureExcelJSDataRowsWithBorders(worksheet, {
      dataStartRow: safeDataStart,
      dataRowCount: rows.length,
      colFrom: tableColMin,
      colTo: tableColMax,
      templateRow: safeDataStart,
      templateBodyRows: 1,
    });
  }

  applyForm14RajasthanColumnLayout(worksheet, templateCols, safeDataStart, rows.length);
  // Re-center title block; keep table data left-aligned.
  applyForm14RajasthanTitleBlockCenter(worksheet, {
    headerRowEnd: Math.max(1, protectedHeaderEnd - 1),
    mergeFromCol: tableColMin,
    mergeToCol: Math.max(tableColMax, tableColMin + 5),
  });
  ensureForm14RajasthanTableHeaderLabels(worksheet, layout);
  for (let i = 0; i < rows.length; i += 1) {
    for (let c = tableColMin; c <= tableColMax; c += 1) {
      const cell = worksheet.getCell(safeDataStart + i, c);
      const leftAlign = {
        horizontal: 'left',
        vertical: 'middle',
        wrapText: false,
      };
      cell.alignment = leftAlign;
      try {
        cell.style = {
          ...(cell.style && typeof cell.style === 'object' ? cell.style : {}),
          alignment: { ...leftAlign },
        };
      } catch (_) {
        /* ignore */
      }
    }
  }

  // Always keep the Form 14 RJ overtime footnote as one continuous line under the rows.
  const footerRow = writeForm14RajasthanFooterNote(worksheet, {
    afterRow: Math.max(lastDataRow, safeDataStart),
    startCol: tableColMin,
    endCol: Math.max(tableColMax, tableColMin + 5),
    footer: {
      text: normalizeForm14RajasthanFooterNoteText(
        capturedFooter?.text || FORM_14_RJ_FOOTER_NOTE
      ),
      col: tableColMin,
      mergeEndCol: Math.max(tableColMax, tableColMin + 5),
      font: capturedFooter?.font,
    },
  });

  // Second note — also one continuous sentence (template model).
  writeForm14RajasthanDayEntriesNote(worksheet, {
    afterRow: footerRow || Math.max(lastDataRow, safeDataStart),
    startCol: tableColMin,
    endCol: Math.max(tableColMax, tableColMin + 5),
    note: {
      text: normalizeForm14RajasthanDayEntriesNoteText(
        capturedDayEntriesNote?.text || FORM_14_RJ_DAY_ENTRIES_NOTE
      ),
      col: tableColMin,
      mergeEndCol: Math.max(tableColMax, tableColMin + 5),
      font: capturedDayEntriesNote?.font,
    },
  });

  const out = await workbook.xlsx.writeBuffer();
  const fileName =
    formFileName ||
    parsedFormHeader?.title?.replace(/[^a-zA-Z0-9]/g, '_') ||
    'Form_14_RJ_Rajasthan.xlsx';
  return {
    blob: new Blob([out], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    }),
    fileName,
  };
}
