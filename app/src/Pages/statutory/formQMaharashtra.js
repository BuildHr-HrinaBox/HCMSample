/** Maharashtra Form Q — muster-roll cum wage register payroll mapping. */

import {
  flattenPayrollEarningColumns,
  readPayrollScalar,
} from '../../utils/payrollEarnings';
import { buildFormPGJEmployeePayrollCandidates } from './formPGJGujarat';

function normHeaderLabel(h) {
  return String(h || '')
    .replace(/^\(?\d+\)?\s*[\.\)]?\s*/i, '')
    .replace(/\r?\n/g, ' ')
    .replace(/[^a-z0-9 ]/gi, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

function parsePayrollNumber(value) {
  const n = Number(String(value ?? '').replace(/[,₹]/g, '').trim());
  return Number.isFinite(n) ? n : NaN;
}

function formatCellValue(value, { allowZero = false } = {}) {
  if (value == null || value === '') return '';
  if (typeof value === 'string' && /^nil$/i.test(value.trim())) return 'Nil';
  const num = parsePayrollNumber(value);
  if (Number.isFinite(num) && (allowZero || num !== 0)) return num;
  return String(value).trim();
}

export function matchFormQMaharashtraWageTailBucket(header) {
  const n = normHeaderLabel(header);
  if (!n) return '';
  if (/total\s+days?\s+worked/.test(n)) return 'paidDays';
  if (/minimum\s+rate\s+of\s+wages/.test(n)) return 'minimumRateWages';
  if (/total\s+production/.test(n) && /piece\s+rate/.test(n)) return 'skip';
  if (/actual\s+wages\s+paid/.test(n)) return 'basic';
  if (/house\s+rent\s+allowance/.test(n)) return 'hra';
  if (/dearness\s+allowance/.test(n)) return 'skip';
  if (/gross\s+amount\s+payable/.test(n)) return 'grossPay';
  if (/total\s+hours?\s+of\s+overtime/.test(n)) return 'overtimeHours';
  if (/overtime\s+earnings?/.test(n)) return 'overtimeEarnings';
  if (/provident\s+fund/.test(n) || /^pf$/.test(n)) return 'providentFund';
  if (/family\s+pension/.test(n)) return 'skip';
  if (/esi\s+contribution/.test(n)) return 'skip';
  if (/professional\s+tax/.test(n) || /^pt$/.test(n)) return 'professionalTax';
  if (/income\s+tax/.test(n) || /^tds$/.test(n) || /^incometax$/.test(n.replace(/\s+/g, ''))) {
    return 'incomeTax';
  }
  if (/loan.*interest/.test(n)) return 'skip';
  if (/^advances$/.test(n)) return 'skip';
  if (/other\s+deductions?/.test(n)) return 'skip';
  if (/total\s+deductions?/.test(n)) return 'totalDeduction';
  if (/net\s+payable/.test(n)) return 'netPay';
  if (/date\s+of\s+payment/.test(n)) return 'paymentDate';
  if (/signature|thumb\s+impression/.test(n)) return 'skip';
  return '';
}

export function resolveFormQMaharashtraPayrollFields(payrollRow, helpers = {}) {
  const monthEndDate = String(helpers.monthEndDate || '').trim();
  const empty = {
    paidDays: '',
    minimumRateWages: '',
    basic: '',
    hra: '',
    grossPay: '',
    overtimeHours: 'Nil',
    overtimeEarnings: 'Nil',
    providentFund: '',
    professionalTax: '',
    incomeTax: '',
    totalDeduction: '',
    netPay: '',
    paymentDate: monthEndDate,
  };
  if (!payrollRow || payrollRow.fetch_error) return empty;

  const flat = flattenPayrollEarningColumns(payrollRow);
  const merged = { ...flat, ...payrollRow };

  const pickAmount = (...values) => {
    for (let i = 0; i < values.length; i += 1) {
      const value = values[i];
      if (value === '' || value == null) continue;
      const n = parsePayrollNumber(value);
      if (Number.isFinite(n)) return n;
    }
    return '';
  };

  const paidDays = readPayrollScalar(
    merged,
    ['paid_days', 'Paid_days', 'Paid Days', 'paidDays', 'days_worked', 'Days Worked', 'no_of_days_worked'],
    [/^paid_days$/, /^paiddays$/, /daysworked/]
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

  const netPay = readPayrollScalar(
    merged,
    ['net_pay', 'Net Pay', 'netPay', 'monthly_salary', 'Net Payable'],
    [/^net_pay$/]
  );

  // Sample Payroll: PF ← epf_contribution / PF, PT ← professional_tax, IT ← income_tax / TDS.
  const providentFund = pickAmount(
    flat.epf_contribution,
    flat.pf,
    flat.PF,
    flat.provident_fund,
    readPayrollScalar(
      merged,
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

  const professionalTax = pickAmount(
    flat.professional_tax,
    flat.ProfessionalTax,
    flat.professionalTax,
    flat.pt,
    flat.PT,
    readPayrollScalar(
      merged,
      ['professional_tax', 'Professional Tax', 'ProfessionalTax', 'professionalTax', 'pt', 'PT'],
      [/^professional_tax$/, /^pt$/]
    )
  );

  const incomeTax = pickAmount(
    flat.income_tax,
    flat.IncomeTax,
    flat.incomeTax,
    flat.tds,
    readPayrollScalar(
      merged,
      [
        'income_tax',
        'Income Tax',
        'IncomeTax',
        'incomeTax',
        'tds',
        'TDS',
        'tax_deducted_at_source',
      ],
      [/income.*tax/i, /^tds$/, /^incometax$/]
    )
  );

  const grossNum = parsePayrollNumber(grossPay);
  const netNum = parsePayrollNumber(netPay);

  let minimumRateWages = '';
  if (Number.isFinite(grossNum) && grossNum > 0) {
    minimumRateWages = Math.round((grossNum / 26) * 100) / 100;
  }

  let totalDeduction = '';
  if (Number.isFinite(grossNum) && Number.isFinite(netNum)) {
    totalDeduction = Math.round((grossNum - netNum) * 100) / 100;
    if (totalDeduction < 0) totalDeduction = '';
  }

  return {
    paidDays,
    minimumRateWages,
    basic,
    hra,
    grossPay,
    overtimeHours: 'Nil',
    overtimeEarnings: 'Nil',
    providentFund,
    professionalTax,
    incomeTax,
    totalDeduction,
    netPay,
    paymentDate: monthEndDate,
  };
}

export function applyFormQMaharashtraPayrollToRow(row, payrollRow, headers, helpers = {}) {
  const hdrs = Array.isArray(headers) ? headers : [];
  const out = row && typeof row === 'object' ? row : {};
  const {
    sanitizeValue = (v) => String(v ?? '').trim(),
    monthEndDate = '',
    overwrite = true,
  } = helpers;
  const payroll = resolveFormQMaharashtraPayrollFields(
    payrollRow && !payrollRow.fetch_error ? payrollRow : null,
    { monthEndDate }
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
    const bucket = matchFormQMaharashtraWageTailBucket(header);
    if (!bucket || bucket === 'skip') return;
    if (bucket === 'overtimeHours' || bucket === 'overtimeEarnings') {
      if (setCell(header, payroll[bucket], { allowZero: true })) wrote = true;
      return;
    }
    if (bucket === 'paymentDate') {
      if (setCell(header, payroll.paymentDate, { allowZero: true })) wrote = true;
      return;
    }
    if (Object.prototype.hasOwnProperty.call(payroll, bucket)) {
      if (setCell(header, payroll[bucket], { allowZero: true })) wrote = true;
    }
  });
  return wrote;
}

export function enrichFormQMaharashtraPayrollRows(mappedData, employees, headers, helpers = {}) {
  const hdrs = Array.isArray(headers) && headers.length > 0 ? headers : [];
  const {
    resolvePayrollRow = null,
    sanitizeValue = (v) => String(v ?? '').trim(),
    monthEndDate = '',
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
      payrollRow = resolvePayrollRow(empItem, row, globalRowIndex);
      if (!payrollRow || payrollRow.fetch_error) {
        const candidates = buildFormPGJEmployeePayrollCandidates(empItem, row);
        for (let ci = 0; ci < candidates.length; ci += 1) {
          const hit = resolvePayrollRow(candidates[ci], row, globalRowIndex);
          if (hit && !hit.fetch_error) {
            payrollRow = hit;
            break;
          }
        }
      }
    }
    const wrote = applyFormQMaharashtraPayrollToRow(row, payrollRow, hdrs, {
      sanitizeValue,
      monthEndDate,
      overwrite,
    });
    if (wrote) hits += 1;
  });
  return hits;
}

function formQExcelJsCellText(val) {
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

const FORM_Q_MH_TITLE = 'FORM Q';
const FORM_Q_MH_REFERENCE = '(See Rule 26(1))';
const FORM_Q_MH_SUBTITLE = 'MUSTER-ROLL CUM WAGE REGISTER';

function looksLikeFormQMaharashtraMainTitle(text) {
  const t = String(text || '')
    .replace(/\s+/g, ' ')
    .trim();
  return /^form\s*q$/i.test(t);
}

function looksLikeFormQMaharashtraReference(text) {
  const t = String(text || '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
  return /see\s*rule\s*26/.test(t);
}

function looksLikeFormQMaharashtraSubtitle(text) {
  const t = String(text || '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
  return /muster[-\s]*roll/.test(t) && /wage\s*register/.test(t);
}

/**
 * Ensure Form Q MH title band:
 *   FORM Q
 *   (See Rule 26(1))
 *   MUSTER-ROLL CUM WAGE REGISTER
 * All three bold; muster-roll sits directly under the See Rule line.
 */
export function ensureFormQMaharashtraTitleLayout(worksheet, options = {}) {
  if (!worksheet) return false;

  const cellText = formQExcelJsCellText;
  const scanTo = Math.max(1, Number(options.scanTo) || 10);
  const maxC = Math.max(1, Number(options.maxCol) || 20);

  const findHit = (matcher) => {
    for (let r = 1; r <= scanTo; r += 1) {
      for (let c = 1; c <= maxC; c += 1) {
        const t = cellText(worksheet.getCell(r, c)?.value).trim();
        if (!t) continue;
        if (matcher(t)) return { row: r, col: c, text: t };
      }
    }
    return null;
  };

  let titleHit = findHit(looksLikeFormQMaharashtraMainTitle);
  let refHit = findHit(looksLikeFormQMaharashtraReference);
  let subHit = findHit(looksLikeFormQMaharashtraSubtitle);

  // Only touch Form Q muster sheets (avoid rewriting unrelated workbooks).
  const sheetName = String(worksheet.name || '');
  const looksFormQ =
    /form\s*q/i.test(sheetName) || titleHit || refHit || subHit;
  if (!looksFormQ) return false;
  if (!titleHit && !refHit && !subHit) return false;

  // Prefer writing into existing title columns; default to col A.
  const writeCol =
    Number(options.writeCol) ||
    titleHit?.col ||
    refHit?.col ||
    subHit?.col ||
    1;

  // Keep relative order: FORM Q → See Rule → MUSTER-ROLL on consecutive rows when possible.
  let titleRow = titleHit?.row || 3;
  let refRow = refHit?.row || titleRow + 1;
  let subRow = subHit?.row || refRow + 1;

  if (refRow <= titleRow) refRow = titleRow + 1;
  if (subRow <= refRow) subRow = refRow + 1;

  // If muster-roll was above See Rule, move it under See Rule.
  if (subHit && refHit && subHit.row < refHit.row) {
    subRow = refRow + 1;
  }

  const applyBoldTitle = (row, text, { size } = {}) => {
    if (row < 1 || !text) return;
    // Clear duplicate title fragments on this row (keep other content like blanks).
    for (let c = 1; c <= maxC; c += 1) {
      const cell = worksheet.getCell(row, c);
      const t = cellText(cell?.value).trim();
      if (
        looksLikeFormQMaharashtraMainTitle(t) ||
        looksLikeFormQMaharashtraReference(t) ||
        looksLikeFormQMaharashtraSubtitle(t)
      ) {
        if (c !== writeCol) cell.value = null;
      }
    }
    const cell = worksheet.getCell(row, writeCol);
    cell.value = text;
    cell.font = {
      ...(cell.font || {}),
      bold: true,
      ...(size ? { size } : cell.font?.size ? {} : { size: 12 }),
      name: cell.font?.name || 'Calibri',
    };
    cell.alignment = {
      ...(cell.alignment || {}),
      horizontal: cell.alignment?.horizontal || 'left',
      vertical: 'middle',
    };
  };

  applyBoldTitle(titleRow, FORM_Q_MH_TITLE, { size: 14 });
  applyBoldTitle(refRow, FORM_Q_MH_REFERENCE, { size: 11 });
  applyBoldTitle(subRow, FORM_Q_MH_SUBTITLE, { size: 12 });
  return true;
}

/** Table body + header grid font — matches PDF Form Q MH size 8. */
export const FORM_Q_MH_TABLE_FONT_SIZE = 8;

export function applyFormQMaharashtraTableFontSize(
  worksheet,
  { rowFrom, rowTo, colFrom, colTo, size = FORM_Q_MH_TABLE_FONT_SIZE } = {}
) {
  if (!worksheet || !(rowFrom > 0) || !(rowTo >= rowFrom) || !(colTo >= colFrom)) return false;
  for (let r = rowFrom; r <= rowTo; r += 1) {
    for (let c = colFrom; c <= colTo; c += 1) {
      const cell = worksheet.getCell(r, c);
      cell.font = {
        ...(cell.font || {}),
        size,
        name: cell.font?.name || 'Calibri',
      };
    }
  }
  return true;
}

function formQMhCellText(val) {
  return formQExcelJsCellText(val).replace(/\s+/g, ' ').trim();
}

function formQMhUnmergeCovering(worksheet, r1, c1, r2, c2) {
  const merges = Array.isArray(worksheet.model?.merges) ? [...worksheet.model.merges] : [];
  merges.forEach((label) => {
    const parts = String(label || '').split(':');
    if (parts.length !== 2) return;
    try {
      const tl = worksheet.getCell(parts[0]);
      const br = worksheet.getCell(parts[1]);
      if (!tl || !br || br.row < r1 || tl.row > r2 || br.col < c1 || tl.col > c2) return;
      worksheet.unMergeCells(label);
    } catch (_) {
      /* ignore */
    }
  });
}

function formQMhApplyParentMerge(worksheet, headerRow, startCol, endCol, label) {
  if (!(endCol > startCol) || headerRow < 1 || startCol < 1) return false;
  formQMhUnmergeCovering(worksheet, headerRow, startCol, headerRow, endCol);
  for (let c = startCol + 1; c <= endCol; c += 1) {
    try {
      worksheet.getCell(headerRow, c).value = null;
    } catch (_) {
      /* ignore */
    }
  }
  try {
    worksheet.mergeCells(headerRow, startCol, headerRow, endCol);
  } catch (_) {
    /* already merged */
  }
  const cell = worksheet.getCell(headerRow, startCol);
  cell.value = label;
  cell.alignment = {
    ...(cell.alignment || {}),
    horizontal: 'center',
    vertical: 'middle',
    wrapText: true,
    textRotation: 0,
  };
  cell.font = {
    ...(cell.font || {}),
    bold: true,
    size: cell.font?.size || FORM_Q_MH_TABLE_FONT_SIZE,
    name: cell.font?.name || 'Calibri',
  };
  return true;
}

/**
 * Re-apply Excel-model parent merges so download matches the official Form Q grid:
 * Working hours (From|To), Interval for Rest (From|To), Date of the Month (1..N),
 * Deductions (PF … Other Deductions).
 */
export function ensureFormQMaharashtraGroupHeaderMerges(worksheet, options = {}) {
  if (!worksheet) return false;
  const maxR = Math.min(Math.max(worksheet.rowCount || 40, 40), 80);
  const maxC = Math.max(worksheet.columnCount || 0, worksheet.actualColumnCount || 0, 120);
  const scanTo = Math.max(1, Number(options.scanTo) || 40);

  let headerRow = Number(options.headerRow) || 0;
  let startCol = Number(options.startCol) || 0;
  if (!(headerRow > 0)) {
    for (let r = 1; r <= Math.min(scanTo, maxR); r += 1) {
      for (let c = 1; c <= maxC; c += 1) {
        const t = formQMhCellText(worksheet.getCell(r, c)?.value).toLowerCase();
        if (/full\s+name\s+of\s+the\s+worker/.test(t) || /^sr\.?\s*no/.test(t)) {
          headerRow = r;
          startCol = startCol || Math.max(1, c - 1);
          break;
        }
      }
      if (headerRow > 0) break;
    }
  }
  if (!(headerRow > 0)) return false;
  if (!(startCol > 0)) startCol = 3;

  // Parent group labels may sit on headerRow or one row above (Sr. No. on leaf row).
  let parentRow = headerRow;
  for (let r = Math.max(1, headerRow - 1); r <= headerRow; r += 1) {
    let hits = 0;
    for (let c = startCol; c <= Math.min(startCol + 80, maxC); c += 1) {
      const t = formQMhCellText(worksheet.getCell(r, c)?.value).toLowerCase();
      if (
        /working\s+hours/.test(t) ||
        /interval\s+for\s+rest/.test(t) ||
        /date\s+of\s+(the\s+)?month/.test(t) ||
        /^deductions?$/.test(t)
      ) {
        hits += 1;
      }
    }
    if (hits >= 2) {
      parentRow = r;
      break;
    }
  }
  const leafRow = parentRow === headerRow ? headerRow + 1 : headerRow;
  const getLeaf = (c) => formQMhCellText(worksheet.getCell(leafRow, c)?.value);
  const getParent = (c) => formQMhCellText(worksheet.getCell(parentRow, c)?.value);

  let wrote = false;

  // Working hours → From|To
  for (let c = startCol; c <= Math.min(startCol + 40, maxC); c += 1) {
    const t = getParent(c).toLowerCase();
    if (!/^working\s+hours/.test(t) || /overtime/.test(t)) continue;
    let end = c;
    const leaf0 = getLeaf(c).toLowerCase();
    const leaf1 = getLeaf(c + 1).toLowerCase();
    if (/^from$/.test(leaf0) && /^to$/.test(leaf1)) end = c + 1;
    else if (/^to$/.test(leaf1) || !getParent(c + 1)) end = c + 1;
    if (formQMhApplyParentMerge(worksheet, parentRow, c, end, 'Working hours')) wrote = true;
    break;
  }

  // Interval for Rest → From|To (always span two columns like the Excel model)
  for (let c = startCol; c <= Math.min(startCol + 40, maxC); c += 1) {
    const t = getParent(c).toLowerCase();
    if (!/interval\s+for\s+rest/.test(t)) continue;
    let end = c + 1;
    const leaf0 = getLeaf(c).toLowerCase();
    const leaf1 = getLeaf(c + 1).toLowerCase();
    if (/^from$/.test(leaf0) && /^to$/.test(leaf1)) end = c + 1;
    else if (/^to$/.test(leaf1) || !getParent(c + 1) || /interval\s+for\s+rest/.test(getParent(c + 1).toLowerCase())) {
      end = c + 1;
    }
    // Clear a split second "Interval for Rest" label on the adjacent parent cell.
    if (/interval\s+for\s+rest/.test(getParent(c + 1).toLowerCase())) {
      try {
        worksheet.getCell(parentRow, c + 1).value = null;
      } catch (_) {
        /* ignore */
      }
    }
    if (formQMhApplyParentMerge(worksheet, parentRow, c, end, 'Interval for Rest')) wrote = true;
    break;
  }

  // Date of the Month → days 1..N
  for (let c = startCol; c <= Math.min(startCol + 50, maxC); c += 1) {
    const t = getParent(c).toLowerCase();
    if (!/date\s+of\s+(the\s+)?month/.test(t) || /payment|entry/.test(t)) continue;
    let end = c;
    for (let nc = c; nc <= Math.min(c + 31, maxC); nc += 1) {
      const leaf = getLeaf(nc).replace(/[()]/g, '');
      const dayNum = parseInt(leaf, 10);
      const isDay = /^\d{1,2}$/.test(leaf) && dayNum >= 1 && dayNum <= 31;
      const parentHere = getParent(nc).toLowerCase();
      if (isDay || (nc === c && /date\s+of\s+(the\s+)?month/.test(parentHere))) {
        end = nc;
        continue;
      }
      if (!leaf && !parentHere && end > c) {
        end = nc;
        continue;
      }
      break;
    }
    if (formQMhApplyParentMerge(worksheet, parentRow, c, end, 'Date of the Month')) wrote = true;
    break;
  }

  // Deductions → one merge covering PF … Other Deductions (coalesce split "Deductions" cells)
  let deductionsStart = -1;
  for (let c = startCol; c <= Math.min(startCol + 90, maxC); c += 1) {
    const t = getParent(c).toLowerCase();
    if (/^deductions?$/.test(t) || /^deductions?\s*\(/.test(t)) {
      deductionsStart = c;
      break;
    }
  }
  if (deductionsStart > 0) {
    let end = deductionsStart;
    for (let nc = deductionsStart; nc <= Math.min(deductionsStart + 12, maxC); nc += 1) {
      const leaf = getLeaf(nc).toLowerCase();
      const parentHere = getParent(nc).toLowerCase();
      if (nc === deductionsStart) {
        end = nc;
        continue;
      }
      if (/^deductions?$/.test(parentHere) || /^deductions?\s*\(/.test(parentHere)) {
        // Second split banner — absorb into the same merge.
        try {
          worksheet.getCell(parentRow, nc).value = null;
        } catch (_) {
          /* ignore */
        }
        end = nc;
        continue;
      }
      if (
        /provident|family\s+pension|esi|professional\s+tax|income\s+tax|loan|advances?|other\s+deductions?/.test(
          leaf
        ) ||
        (!leaf && !parentHere)
      ) {
        // Never absorb Total Deduction / Net even when a stray Advances/Other leaf is present.
        if (/total\s+deduction|net\s+payable|date\s+of\s+payment|signature|amount\s+deposited/.test(parentHere)) {
          break;
        }
        end = nc;
        if (/other\s+deductions?/.test(leaf)) break;
        continue;
      }
      if (
        /total\s+deduction|net\s+payable|date\s+of\s+payment|signature|amount\s+deposited/.test(
          parentHere
        ) ||
        /total\s+deduction|net\s+payable|date\s+of\s+payment|signature|amount\s+deposited/.test(leaf)
      ) {
        break;
      }
      break;
    }
    if (formQMhApplyParentMerge(worksheet, parentRow, deductionsStart, end, 'Deductions')) {
      wrote = true;
    }
  }

  // Remove duplicate Signature / Amount Deposited columns after the first Signature
  // (template sometimes leaves an empty col + repeated Signature at the far right).
  try {
    let firstSigCol = -1;
    for (let c = startCol; c <= maxC; c += 1) {
      const stack = `${getParent(c)} ${getLeaf(c)}`.toLowerCase();
      if (/signature|thumb\s+impression/.test(stack)) {
        firstSigCol = c;
        break;
      }
    }
    if (firstSigCol > 0) {
      for (let c = firstSigCol + 1; c <= maxC; c += 1) {
        const stack = `${getParent(c)} ${getLeaf(c)}`.toLowerCase();
        const empty =
          !String(getParent(c) || '').trim() &&
          !String(getLeaf(c) || '').trim() &&
          !String(formQMhCellText(worksheet.getCell(leafRow + 1, c)?.value) || '').trim();
        if (
          empty ||
          /signature|thumb\s+impression/.test(stack) ||
          /amount\s+deposited/.test(stack)
        ) {
          formQMhUnmergeCovering(worksheet, parentRow, c, leafRow + 2, c);
          for (let r = parentRow; r <= Math.min(leafRow + 2, parentRow + 4); r += 1) {
            try {
              worksheet.getCell(r, c).value = null;
            } catch (_) {
              /* ignore */
            }
          }
          wrote = true;
        } else if (stack.trim()) {
          // Non-duplicate content after signature — stop clearing.
          break;
        }
      }
    }
  } catch (_) {
    /* ignore */
  }

  return wrote;
}

/**
 * Heading boxes (Name of the worker, Designation, wage/deduction leaves, etc.):
 * Excel textRotation 90 = bottom→top, matching the muster-roll model.
 * Group banners (Working hours / Interval / Date of the Month / Deductions) stay horizontal.
 */
export function applyFormQMaharashtraVerticalHeaderText(worksheet, options = {}) {
  if (!worksheet) return false;
  const maxR = Math.min(Math.max(worksheet.rowCount || 40, 40), 80);
  const maxC = Math.max(worksheet.columnCount || 0, worksheet.actualColumnCount || 0, 120);
  const scanTo = Math.max(1, Number(options.scanTo) || 40);

  let headerRow = Number(options.headerRow) || 0;
  let startCol = Number(options.startCol) || 0;
  if (!(headerRow > 0)) {
    for (let r = 1; r <= Math.min(scanTo, maxR); r += 1) {
      for (let c = 1; c <= maxC; c += 1) {
        const t = formQMhCellText(worksheet.getCell(r, c)?.value).toLowerCase();
        if (/full\s+name\s+of\s+the\s+worker|name\s+of\s+the\s+worker/.test(t) || /^sr\.?\s*no/.test(t)) {
          headerRow = r;
          startCol = startCol || Math.max(1, c);
          break;
        }
      }
      if (headerRow > 0) break;
    }
  }
  if (!(headerRow > 0)) return false;
  if (!(startCol > 0)) startCol = 1;

  let parentRow = headerRow;
  for (let r = Math.max(1, headerRow - 1); r <= headerRow; r += 1) {
    let hits = 0;
    for (let c = startCol; c <= Math.min(startCol + 80, maxC); c += 1) {
      const t = formQMhCellText(worksheet.getCell(r, c)?.value).toLowerCase();
      if (
        /working\s+hours/.test(t) ||
        /interval\s+for\s+rest/.test(t) ||
        /date\s+of\s+(the\s+)?month/.test(t) ||
        /^deductions?$/.test(t)
      ) {
        hits += 1;
      }
    }
    if (hits >= 2) {
      parentRow = r;
      break;
    }
  }
  const leafRow = parentRow === headerRow ? headerRow + 1 : headerRow;
  const indexRow = leafRow + 1;

  const isVerticalLabel = (raw) => {
    const t = String(raw || '')
      .replace(/\s+/g, ' ')
      .trim()
      .toLowerCase();
    if (!t) return false;
    if (/^working\s+hours/.test(t) && !/overtime/.test(t)) return false;
    if (/interval\s+for\s+rest/.test(t)) return false;
    if (/^date\s+of\s+(the\s+)?month/.test(t) && !/payment|entry/.test(t)) return false;
    if (/^deductions?$/.test(t) || /^deductions?\s*\(/.test(t)) return false;
    if (/^(from|to)$/.test(t)) return false;
    if (/^\(?\d{1,2}\)?$/.test(t)) return false;
    if (/^\d{1,3}$/.test(t)) return false;
    return (
      /^sr\.?\s*no/.test(t) ||
      /full\s+name\s+of\s+the\s+worker|name\s+of\s+the\s+worker/.test(t) ||
      /^designation|nature\s+of\s+work/.test(t) ||
      /^age$/.test(t) ||
      /^sex$/.test(t) ||
      /entry\s+into\s+service/.test(t) ||
      /total\s+days?\s+worked|minimum\s+rate|total\s+production|actual\s+wages?|house\s+rent|dearness|gross\s+amount|overtime|provident|family\s+pension|\besi\b|professional\s+tax|income\s+tax|loan|advances?|other\s+deductions?|total\s+deductions?|net\s+payable|date\s+of\s+payment|bank\s+account|cheque|amount\s+deposited|signature|thumb/.test(
        t
      )
    );
  };

  let wrote = false;
  const applyRotation = (row, col) => {
    const cell = worksheet.getCell(row, col);
    const text = formQMhCellText(cell?.value);
    if (!isVerticalLabel(text)) return;
    cell.alignment = {
      ...(cell.alignment || {}),
      textRotation: 90,
      wrapText: true,
      horizontal: 'center',
      vertical: 'middle',
    };
    wrote = true;
  };

  for (let c = startCol; c <= maxC; c += 1) {
    applyRotation(parentRow, c);
    applyRotation(leafRow, c);
  }

  try {
    const parentExcelRow = worksheet.getRow(parentRow);
    parentExcelRow.height = Math.max(Number(parentExcelRow.height) || 0, 72);
    const leafExcelRow = worksheet.getRow(leafRow);
    leafExcelRow.height = Math.max(Number(leafExcelRow.height) || 0, 72);
    if (indexRow <= maxR) {
      const indexExcelRow = worksheet.getRow(indexRow);
      indexExcelRow.height = Math.max(Number(indexExcelRow.height) || 0, 18);
    }
  } catch (_) {
    /* ignore */
  }

  return wrote;
}

