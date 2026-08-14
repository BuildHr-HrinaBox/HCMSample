/** Form XVIII — Madhya Pradesh MP Combined Register (Muster Roll-cum Register of Wages). */

import ExcelJS from 'exceljs';
import * as XLSX from 'xlsx';
import {
  clearExcelJSTrailingTableCells,
  countExcelJSTemplateBodyRows,
  ensureExcelJSDataRowsWithBorders
} from '../../utils/excelTableBorders';
import {
  flattenPayrollEarningColumns,
  getEarningsArray,
  mergePayrollRunEmployeePayload,
  readPayrollForm15WageAmounts,
  readPayrollScalar,
} from '../../utils/payrollEarnings';
import {
  parseLeaveCellObject,
  sanitizeLeaveMetricDisplayValue,
} from '../../utils/leaveMetrics';
import {
  getCachedForm15PayrollTableRows,
  getLatestCachedPayrollTableRows,
  getPayrollBulkRowsForAutofill,
  yieldToMain,
} from '../../utils/statutoryAutofillCache';
import { writeStatutoryHeaderFieldsToExcelJsWorksheet } from '../../utils/statutorySiteCompanyHeaders';
import {
  buildFormXIXMPPayrollRowResolver,
  collectEmployeeNameCandidates,
  employeeGidCandidates,
  loadFormXIXMPPayrollRowsForAutofill,
  resolveFormXIXMPPayrollRowForEmployee,
  resolveFormXIXMPPayrollRowsForAutofill,
  resolvePayrollRowByNameAndGid,
} from './formXIXMPWageSlip';
import { leaveTypeLabelMatchesAliases } from './formXTamilNaduLeave';

export const loadFormXVIIIMPPayrollRowsForAutofill = loadFormXIXMPPayrollRowsForAutofill;

/** True when Sample Payroll / pay-run row carries PF or Professional Tax scalars. */
export function formXVIIIMPPayrollRowHasSampleDeductionFields(payrollRow) {
  if (!payrollRow || payrollRow.fetch_error) return false;
  const flat = flattenPayrollEarningColumns(payrollRow);
  const pf = resolveFormXVIIIMPSamplePayrollPf(flat, payrollRow, payrollRow);
  const pt = resolveFormXVIIIMPSamplePayrollPt(flat, payrollRow, payrollRow);
  return (pf !== '' && pf != null) || (pt !== '' && pt != null);
}

export function formXVIIIMPPayrollRowHasSamplePf(payrollRow) {
  if (!payrollRow || payrollRow.fetch_error) return false;
  const flat = flattenPayrollEarningColumns(payrollRow);
  const pf = resolveFormXVIIIMPSamplePayrollPf(flat, payrollRow, payrollRow);
  return pf !== '' && pf != null;
}

/**
 * Prefer Sample Payroll rows that carry PF / Professional Tax (Form XVIII MP col 22).
 * Zoho-only pay-run rows often have gross/net but empty PF / PT scalars —
 * never early-return those when a Sample Payroll snapshot with deductions exists.
 */
export function resolveFormXVIIIMPPayrollRowsForAutofill(
  statutoryPayrollRows,
  monthCandidates = [],
  helpers = {}
) {
  const flattenRows = (rows) =>
    (Array.isArray(rows) ? rows : [])
      .filter((row) => row && typeof row === 'object' && row.fetch_error !== true)
      .map((row) => flattenPayrollEarningColumns(row));

  const preferPfRows = (rows) => {
    const list = flattenRows(rows);
    if (list.length === 0) return [];
    const withPf = list.filter((row) => formXVIIIMPPayrollRowHasSamplePf(row));
    if (withPf.length > 0) return withPf;
    const withDed = list.filter((row) => formXVIIIMPPayrollRowHasSampleDeductionFields(row));
    return withDed.length > 0 ? withDed : list;
  };

  const takeIfHasPf = (rows) => {
    const preferred = preferPfRows(rows);
    if (preferred.length > 0 && preferred.some((row) => formXVIIIMPPayrollRowHasSamplePf(row))) {
      return preferred;
    }
    return null;
  };
  const takeIfHasDed = (rows) => {
    const preferred = preferPfRows(rows);
    if (
      preferred.length > 0 &&
      preferred.some((row) => formXVIIIMPPayrollRowHasSampleDeductionFields(row))
    ) {
      return preferred;
    }
    return null;
  };

  const cached =
    helpers.cachedSampleRows || getCachedForm15PayrollTableRows(monthCandidates)?.rows || null;
  const bulk = helpers.bulkSampleRows || getPayrollBulkRowsForAutofill() || null;
  const latest = getLatestCachedPayrollTableRows()?.rows || null;

  const pfHit =
    takeIfHasPf(cached) ||
    takeIfHasPf(bulk) ||
    takeIfHasPf(latest) ||
    takeIfHasPf(statutoryPayrollRows);
  if (pfHit) return pfHit;

  const dedHit =
    takeIfHasDed(cached) ||
    takeIfHasDed(bulk) ||
    takeIfHasDed(latest) ||
    takeIfHasDed(statutoryPayrollRows);
  if (dedHit) return dedHit;

  const base = resolveFormXIXMPPayrollRowsForAutofill(statutoryPayrollRows, monthCandidates);
  const fromBase = takeIfHasPf(base) || takeIfHasDed(base);
  if (fromBase) return fromBase;

  const fallbackCandidates = [cached, bulk, latest, statutoryPayrollRows, base];
  for (let i = 0; i < fallbackCandidates.length; i += 1) {
    const fallback = preferPfRows(fallbackCandidates[i]);
    if (fallback.length > 0) return fallback;
  }
  return [];
}

export function buildFormXVIIIMPPayrollRowResolver(payrollRows) {
  return buildFormXIXMPPayrollRowResolver(payrollRows);
}

const mpPayrollRowNameCandidates = (row) => {
  if (!row || typeof row !== 'object') return [];
  const full = String(
    row.employee_name || row.full_name || row.name || `${row.first_name || ''} ${row.last_name || ''}`
  )
    .trim()
    .toLowerCase();
  const out = full ? [full] : [];
  const paren = full.match(/^(.+?)\s*\(([^)]+)\)\s*$/);
  if (paren) {
    out.push(paren[1].trim());
    out.push(paren[2].trim());
  }
  return [...new Set(out.filter(Boolean))];
};

const mpFormRowNameCandidates = (row, headers) => {
  if (!row || typeof row !== 'object') return [];
  const list = Array.isArray(headers) ? headers : [];
  const out = [];
  list.forEach((header) => {
    const s = mpCombinedRegisterHeaderNorm(header);
    if (!s.includes('name')) return;
    if (
      !(
        s.includes('workman') ||
        s.includes('employee') ||
        s.includes('worker') ||
        s.includes('workmen')
      )
    ) {
      return;
    }
    const raw = String(row[header] ?? '').trim().toLowerCase();
    if (!raw) return;
    out.push(raw);
    const paren = raw.match(/^(.+?)\s*\(([^)]+)\)\s*$/);
    if (paren) {
      out.push(paren[1].trim());
      out.push(paren[2].trim());
    }
  });
  return [...new Set(out.filter(Boolean))];
};

/** Match pay-run payroll row — People employee, then grid name (same strategy as leave API). */
export function resolveFormXVIIIMPPayrollRowForAutofillRow(emp, row, payrollRows, headers, resolver) {
  const rows = Array.isArray(payrollRows) ? payrollRows : [];
  if (rows.length === 0) return null;
  const resolve =
    typeof resolver === 'function' ? resolver : buildFormXVIIIMPPayrollRowResolver(rows);
  const fromPeople = resolve(emp);
  if (fromPeople && !fromPeople.fetch_error) return fromPeople;

  const nameCandidates = [
    ...collectEmployeeNameCandidates(emp),
    ...mpFormRowNameCandidates(row, headers),
  ];
  const gidCandidates = [
    ...employeeGidCandidates(emp),
    ...employeeGidCandidates(null, row),
  ];
  const compositeHit = resolvePayrollRowByNameAndGid(
    [...new Set(nameCandidates)],
    [...new Set(gidCandidates)],
    rows
  );
  if (compositeHit) return compositeHit;
  if (gidCandidates.length > 0 && nameCandidates.length > 0) return null;

  const rowNames = mpFormRowNameCandidates(row, headers);
  if (rowNames.length > 0) {
    const hit = rows.find((pr) => {
      if (!pr || pr.fetch_error) return false;
      const payrollNames = mpPayrollRowNameCandidates(pr);
      return rowNames.some((rowName) =>
        payrollNames.some(
          (payrollName) =>
            payrollName === rowName ||
            payrollName.includes(rowName) ||
            rowName.includes(payrollName)
        )
      );
    });
    if (hit) return hit;
  }

  return resolveFormXIXMPPayrollRowForEmployee(emp, rows);
}

export function mpCombinedRegisterHeaderNorm(txt) {
  return String(txt || '')
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

/**
 * Karnataka Form T ("COMBINED MUSTER ROLL CUM REGISTER OF WAGES") shares wording with
 * MP Form XVIII — never treat Form T as the Madhya Pradesh combined register.
 */
export function looksLikeKarnatakaFormTSheet(blob) {
  const s = String(blob || '').toLowerCase();
  if (/\bform[\s._-]*xviii\b|\bform_xviii\b|\bform-xviii\b/.test(s)) return false;
  if (/\bmadhya\s+pradesh\b/.test(s)) return false;
  // Matches "Form T", "Form_T", "Form___T_-_Karnataka.xlsx", etc.
  const namedFormT =
    /\bform[\s._-]*t\b/.test(s) ||
    /form[_-]{1,}t(?:[_-]|$)/.test(s) ||
    /form_{2,}t/.test(s) ||
    /form[\s._-]*t[\s._-]*ka\b/.test(s);
  if (!namedFormT) return false;
  return (
    /karnataka/.test(s) ||
    /rule\s+24\s*\(\s*9[\s-]*b/.test(s) ||
    /shops?\s*.{0,24}commercial/.test(s) ||
    /combined\s+muster\s+roll/.test(s) ||
    /muster[\s._-]*roll[\s._-]*cum[\s._-]*register[\s._-]*of[\s._-]*wages/.test(s) ||
    (/s[\s&._-]*e\b/.test(s) && /form[\s._-]*t/.test(s))
  );
}

/**
 * Tamil Nadu CLRA Form XVIII (Register of Wages-cum-Muster Roll) shares "Form XVIII" /
 * "wages-cum-muster" wording with MP Form XVIII — never treat it as the MP combined register.
 */
export function looksLikeTamilNaduFormXVIIISheet(blob) {
  const s = String(blob || '').toLowerCase();
  if (!/tamil[\s._-]*nadu|tamilnadu/.test(s) && !/form_xviii[_\s-]*tamil|form[\s._-]*xviii[_\s-]*tamil/.test(s)) {
    return false;
  }
  // Explicit Madhya Pradesh identity wins over a stray "Tamil Nadu" mention elsewhere.
  if (/\bmadhya\s+pradesh\b/.test(s) && !/form_xviii[_\s-]*tamil|form[\s._-]*xviii[_\s-]*tamil/.test(s)) {
    return false;
  }
  return (
    /form[\s._-]*xviii(?![a-z])/i.test(s) ||
    /form[\s._-]*18(?!\d)/i.test(s) ||
    /register[\s._-]*of[\s._-]*wages[\s._-]*cum[\s._-]*muster/i.test(s) ||
    /register\s+of\s+wages[\s-]*cum[\s-]*muster\s+roll/i.test(s)
  );
}

export function looksLikeMPCombinedRegisterSheet(blob) {
  const s = String(blob || '').toLowerCase();
  if (looksLikeKarnatakaFormTSheet(s)) return false;
  if (looksLikeTamilNaduFormXVIIISheet(s)) return false;
  return (
    /muster[\s._-]*roll[\s._-]*cum[\s._-]*register[\s._-]*of[\s._-]*wages/i.test(s) ||
    /register[\s._-]*of[\s._-]*wages[\s._-]*cum[\s._-]*muster/i.test(s) ||
    /mp\s+combined\s+register/i.test(s) ||
    (/\bmadhya\s+pradesh\b/.test(s) &&
      (/\bform\s*xviii\b/.test(s) || /combined\s+register/.test(s) || /muster\s+roll/.test(s))) ||
    (s.includes('name of the establishment and address') && s.includes('location of work'))
  );
}

export function isFormXVIIIMPCombinedRegisterContext(formHeader, rowItem, fileName, sheetText = '') {
  const parts = [
    rowItem?.formName,
    rowItem?.FormName,
    rowItem?.state,
    rowItem?.State,
    rowItem?.description,
    rowItem?.Description,
    fileName,
    formHeader?.title,
    formHeader?.subtitle,
    formHeader?.reference,
    sheetText
  ]
    .join(' ')
    .toLowerCase();
  if (looksLikeKarnatakaFormTSheet(parts)) return false;
  // Form_XVIII_-_TamilNadu.xlsx must use Central/TN wages-cum-muster parse — not MP rebuild.
  if (looksLikeTamilNaduFormXVIIISheet(parts)) return false;
  return (
    /\bform\s*xviii\b|\bform_xviii\b|\bform-xviii\b/i.test(parts) ||
    looksLikeMPCombinedRegisterSheet(parts)
  );
}

function excelCellLooksLikeSerialHeader(rawText) {
  const t = mpCombinedRegisterHeaderNorm(rawText);
  return !!(
    t &&
    (/\bs\.?\s*no\.?\b/.test(t) ||
      /\bsr\.?\s*no\.?\b/.test(t) ||
      /\bsl\.?\s*no\.?\b/.test(t) ||
      t.includes('serial'))
  );
}

export function isWorksheetMergeAnchorCell(merges, rowIndex, colIndex) {
  if (!Array.isArray(merges) || merges.length === 0) return true;
  for (let i = 0; i < merges.length; i += 1) {
    const m = merges[i];
    if (!m?.s || !m?.e) continue;
    if (rowIndex >= m.s.r && rowIndex <= m.e.r && colIndex >= m.s.c && colIndex <= m.e.c) {
      return rowIndex === m.s.r && colIndex === m.s.c;
    }
  }
  return true;
}

function isEffectivelyBlankHeaderCell(text) {
  const t = String(text || '').replace(/\s+/g, ' ').trim();
  if (!t) return true;
  return /^[-–—_\s.]+$/i.test(t);
}

export function normalizeStatutoryColumnIndexCell(text) {
  return String(text || '')
    .replace(/[()]/g, '')
    .replace(/\s+/g, '')
    .trim();
}

export function rowLooksLikeStatutoryColumnIndexRow(
  rowIndex,
  startCol,
  getMergedAwareCellText,
  numCols = 32,
  maxScan = 32
) {
  let hits = 0;
  let total = 0;
  const lim = Math.min(numCols, maxScan);
  for (let i = 0; i < lim; i += 1) {
    const c = startCol + i;
    const t = normalizeStatutoryColumnIndexCell(getMergedAwareCellText(rowIndex, c));
    if (!t) continue;
    total += 1;
    if (/^\d+$/.test(t)) hits += 1;
  }
  return total >= 4 && hits >= Math.max(3, Math.floor(total * 0.7));
}

export function findSlNoColumnInSheetRow(rowIndex, getMergedAwareCellText, maxCols = 40) {
  for (let c = 0; c < maxCols; c += 1) {
    const raw = getMergedAwareCellText(rowIndex, c);
    if (!excelCellLooksLikeSerialHeader(raw)) continue;
    const s = mpCombinedRegisterHeaderNorm(raw).replace(/[^a-z0-9]/g, '');
    if (
      s === 'srno' ||
      s === 'sno' ||
      s === 'slno' ||
      (s.includes('sr') && s.includes('no')) ||
      (s.includes('sl') && s.includes('no'))
    ) {
      return c;
    }
  }
  return -1;
}

function inferMPTableEndCol({
  mainRow,
  startCol,
  getMergedAwareCellText,
  getRawCellText,
  maxCols,
  dataStartRow = -1,
  bandEndRow = -1
}) {
  const readRaw = typeof getRawCellText === 'function' ? getRawCellText : getMergedAwareCellText;
  const bandEnd =
    bandEndRow >= mainRow
      ? Math.min(bandEndRow, maxCols - 1)
      : dataStartRow > mainRow
        ? Math.min(dataStartRow + 1, mainRow + 10, maxCols - 1)
        : Math.min(mainRow + 8, maxCols - 1);
  let endCol = startCol;
  let maxIndex = 0;

  for (let r = Math.max(0, mainRow - 1); r <= bandEnd; r += 1) {
    for (let c = startCol; c < maxCols; c += 1) {
      const idx = normalizeStatutoryColumnIndexCell(getMergedAwareCellText(r, c));
      if (idx && /^\d+$/.test(idx)) {
        maxIndex = Math.max(maxIndex, Number(idx));
        endCol = Math.max(endCol, c);
      }
      const raw = String(readRaw(r, c) || '').trim();
      const merged = String(getMergedAwareCellText(r, c) || '').trim();
      const t = raw || merged;
      if (!t || isEffectivelyBlankHeaderCell(t) || /^\d+$/.test(normalizeStatutoryColumnIndexCell(t))) {
        continue;
      }
      const norm = mpCombinedRegisterHeaderNorm(t);
      if (
        norm.includes('remark') ||
        norm.includes('signature') ||
        norm.includes('thumb') ||
        norm.includes('bank') ||
        norm.includes('utr') ||
        norm.includes('net amount') ||
        norm.includes('payable') ||
        norm === 'pf' ||
        norm.startsWith('pf ') ||
        norm.includes('esic') ||
        norm === 'pt' ||
        norm.startsWith('pt ') ||
        norm.includes('lwf') ||
        norm.includes('welfare fund') ||
        norm.includes('deduction') ||
        norm.includes('allowance') ||
        norm.includes('wage') ||
        norm.includes('overtime')
      ) {
        endCol = Math.max(endCol, c);
      }
    }
  }

  if (maxIndex > 0) {
    for (let r = Math.max(0, mainRow - 1); r <= bandEnd; r += 1) {
      for (let c = startCol; c < maxCols; c += 1) {
        const idx = normalizeStatutoryColumnIndexCell(getMergedAwareCellText(r, c));
        if (idx === String(maxIndex)) {
          endCol = Math.max(endCol, c);
        }
      }
    }
  }

  // MP Combined Register template has 26 numbered columns (Sl. No. through Remarks).
  if (maxIndex >= 20 && endCol < startCol + 25) {
    endCol = Math.min(startCol + 25, maxCols - 1);
  }

  return endCol;
}

function stripTrailingSectionIndex(label) {
  return String(label || '')
    .replace(/\s*\(\s*\d+\s*\)\s*$/i, '')
    .trim()
    .toLowerCase();
}

/** Section-index → canonical header when merged cells inherit the previous column label. */
export const FORM_XVIII_MP_SECTION_INDEX_HEADERS = {
  13: 'Wage rate/ pay or (piece rate/ wages per unit)',
  14: 'Other allowances',
  15: 'Over time worked (Number of hours in the month)',
  16: 'Amount of overtime wages',
  17: 'Amount of Maternity benefit (If any)',
  18: 'Any other Amount (Please mention)',
  19: 'Total/ gross Wages/ Earnings',
  23: 'Net amount payable',
};

/** Merged parent "Other Deductions Like EPF/ESI/…" used as the first child (PF) column key. */
export function isFormXVIIIMPOtherDeductionsGroupParentHeader(header) {
  const s = mpCombinedRegisterHeaderNorm(header).replace(/[^a-z0-9 ]/g, ' ').replace(/\s+/g, ' ').trim();
  if (!s.includes('other') || !s.includes('deduction')) return false;
  // Exclude "Other allowances" / "Any other Amount" (not the EPF/ESI/Welfare group).
  if (s.includes('allowance')) return false;
  if (s.includes('amount') && (s.includes('mention') || /^any\s+other/.test(s))) return false;
  // Full group title, or shortened "Other Deductions (if any)" / "Other Deductions" from merge bleed.
  return (
    s.includes('epf') ||
    s.includes('esi') ||
    s.includes('welfare') ||
    s.includes('provident') ||
    /\bpf\b/.test(s) ||
    /^other\s+deductions?(?:\s+if\s+any)?$/.test(s) ||
    /^other\s+deductions?\s+like\b/.test(s)
  );
}

export function isFormXVIIIMPPfHeader(header) {
  const s = mpCombinedRegisterHeaderNorm(header).replace(/[^a-z0-9 ]/g, ' ').replace(/\s+/g, ' ').trim();
  if (isFormXVIIIMPOtherDeductionsGroupParentHeader(header)) return true;
  if (s.includes('registration') || s.includes('uan') || (s.includes('account') && s.includes('number'))) {
    return false;
  }
  // "Other Deductions…_PF (22)" composite keys from grouped Excel models.
  if (/_pf(?:\s*\(?\s*\d+\s*\)?)?\s*$/i.test(String(header || '').trim())) return true;
  if (/\bpf\b/.test(s) && (s.includes('other') || s.includes('deduction'))) return true;
  return (
    s === 'pf' ||
    /^pf\s*\d*$/.test(s) ||
    /^pf\s+\d+/.test(s) ||
    s.startsWith('pf ') ||
    (s.includes('provident') && !s.includes('voluntary'))
  );
}

export function isFormXVIIIMPEsicHeader(header) {
  const s = mpCombinedRegisterHeaderNorm(header).replace(/[^a-z0-9 ]/g, ' ').replace(/\s+/g, ' ').trim();
  if (s.includes('registration') || s.includes('number') || s.includes('no')) return false;
  return s === 'esic' || s.startsWith('esic ') || /^esic\s*\d*/.test(s) || (s.includes('employee') && s.includes('insurance'));
}

export function isFormXVIIIMPPtHeader(header) {
  const raw = String(header || '').trim();
  const s = mpCombinedRegisterHeaderNorm(header).replace(/[^a-z0-9 ]/g, ' ').replace(/\s+/g, ' ').trim();
  // "Other Deductions…_PT (22)" composite keys from grouped Excel models.
  if (/_pt(?:\s*\(?\s*\d+\s*\)?)?\s*$/i.test(raw)) return true;
  if (/\bpt\b/.test(s) && (s.includes('other') || s.includes('deduction'))) return true;
  return (
    s === 'pt' ||
    /^pt\s*\d*$/.test(s) ||
    /^pt\s+\d+/.test(s) ||
    s.startsWith('pt ') ||
    s.includes('professional tax')
  );
}

export function isFormXVIIIMPLwfHeader(header) {
  const s = mpCombinedRegisterHeaderNorm(header).replace(/[^a-z0-9 ]/g, ' ').replace(/\s+/g, ' ').trim();
  return s === 'lwf' || s.startsWith('lwf ') || /^lwf\s*\d*/.test(s) || s.includes('labour welfare') || s.includes('labor welfare');
}

/** Prefer canonical section label when sheet text is blank or wrongly copied from a merge. */
export function resolveFormXVIIIMPSectionHeaderLabel(label, sectionIdx, prevLabel = '') {
  const idx = String(sectionIdx || '').trim();
  const cur = String(label || '').replace(/\s+/g, ' ').trim();
  const prev = String(prevLabel || '').replace(/\s+/g, ' ').trim();
  const curBase = stripTrailingSectionIndex(cur);
  const prevBase = stripTrailingSectionIndex(prev);
  const alias = FORM_XVIII_MP_SECTION_INDEX_HEADERS[idx];
  const prevIsWage = isFormXVIIIMPWageRateHeader(prev);
  const curIsWage = isFormXVIIIMPWageRateHeader(cur);
  const curBlankOrColumn = !cur || /^column\s+\d+$/i.test(cur);

  // Merged-header bleed: column after Wage rate repeating the same label ⇒ Other allowances.
  // Also covers wrong section index "13" inherited from a merged index cell.
  if (prevIsWage && (curIsWage || (curBlankOrColumn && idx !== '13'))) {
    return FORM_XVIII_MP_SECTION_INDEX_HEADERS[14];
  }
  // Merged parent "Other Deductions Like EPF/ESI/…" — first column of the band is PF.
  // Leave later duplicate parent labels for repairFormXVIIIMPOtherDeductionsPfHeader.
  if (
    isFormXVIIIMPOtherDeductionsGroupParentHeader(cur) &&
    !isFormXVIIIMPOtherDeductionsGroupParentHeader(prev) &&
    !isFormXVIIIMPPfHeader(prev)
  ) {
    return idx && /^\d+$/.test(idx) ? `PF (${idx})` : 'PF (22)';
  }
  if (!alias) return cur;
  if (!cur || /^column\s+\d+$/i.test(cur)) return alias;
  if (prevBase && curBase && prevBase === curBase) return alias;
  // Col 14 must never keep a Wage-rate label inherited from a merged parent cell.
  if (idx === '14' && curIsWage && !isFormXVIIIMPOtherAllowanceHeader(cur)) {
    return alias;
  }
  return cur;
}

/** Turn consecutive duplicate Wage-rate headers into Other allowances (download safety net). */
export function repairFormXVIIIMPDuplicateWageHeaders(headers) {
  const list = Array.isArray(headers) ? [...headers] : [];
  for (let i = 1; i < list.length; i += 1) {
    if (isFormXVIIIMPWageRateHeader(list[i]) && isFormXVIIIMPWageRateHeader(list[i - 1])) {
      list[i] = FORM_XVIII_MP_SECTION_INDEX_HEADERS[14];
    }
  }
  return repairFormXVIIIMPOtherDeductionsPfHeader(list);
}

/**
 * Merged "Other Deductions Like EPF/ESI/…" parent text on child columns:
 * 1st → PF (22), 2nd → ESIC (22), 3rd → PT (22), 4th → LWF (22).
 * Also maps parent bleed after an already-renamed PF (22).
 */
export function repairFormXVIIIMPOtherDeductionsPfHeader(headers) {
  const list = Array.isArray(headers) ? [...headers] : [];
  const childDefaults = ['PF (22)', 'ESIC (22)', 'PT (22)', 'LWF (22)'];
  let offset = -1;
  for (let i = 0; i < list.length; i += 1) {
    const isParent = isFormXVIIIMPOtherDeductionsGroupParentHeader(list[i]);
    const isPfOnly =
      !isParent &&
      isFormXVIIIMPPfHeader(list[i]) &&
      (offset < 0 || offset === 0);
    if (isParent || (isPfOnly && offset < 0)) {
      offset = offset < 0 ? 0 : offset + 1;
      if (isParent && offset < childDefaults.length) {
        list[i] = childDefaults[offset];
      }
      continue;
    }
    if (offset >= 0) {
      // Still inside deduction band when leaf ESIC/PT/LWF follows PF.
      const s = mpCombinedRegisterHeaderNorm(list[i])
        .replace(/[^a-z0-9 ]/g, ' ')
        .replace(/\s+/g, ' ')
        .trim();
      if (
        s.startsWith('esic') ||
        s === 'pt' ||
        s.startsWith('pt ') ||
        s.includes('professional tax') ||
        s === 'lwf' ||
        s.startsWith('lwf ') ||
        s.includes('labour welfare') ||
        s.includes('labor welfare')
      ) {
        offset += 1;
      } else {
        offset = -1;
      }
    }
  }
  return list;
}

/** MP Combined Register: header row + optional (1)(2)… index row; map grid 1:1 to Excel columns. */
export function rebuildFormXVIIIMPCombinedRegisterTableHeadersFromSheet({
  headerRowIndex = -1,
  getMergedAwareCellText,
  getRawCellText = null,
  jsonData,
  effectiveSheetCols,
  sheetTextBlob = '',
  forceCombinedRegister = false,
  merges = null
}) {
  const readRaw = typeof getRawCellText === 'function' ? getRawCellText : getMergedAwareCellText;
  if (!Array.isArray(jsonData) || jsonData.length === 0) return null;
  const blob =
    sheetTextBlob ||
    jsonData
      .slice(0, Math.min(25, jsonData.length))
      .map((row) => (Array.isArray(row) ? row : []).map((c) => String(c || '')).join(' '))
      .join(' ');
  if (!forceCombinedRegister && !looksLikeMPCombinedRegisterSheet(blob)) return null;

  const scanMaxCol = Math.max(effectiveSheetCols || 0, 55);
  const rowLooksLikeMpTableHeader = (joined) =>
    (/(\bsr\.?\s*no\b|\bsl\.?\s*no\b|\bs\.?\s*no\b|\bserial\b)/.test(joined) ||
      joined.includes('sr no') ||
      joined.includes('sl no')) &&
    (joined.includes('emp') ||
      joined.includes('full name') ||
      joined.includes('employee') ||
      joined.includes('serial number in') ||
      joined.includes('workmen'));

  let mainRow = headerRowIndex >= 0 ? headerRowIndex : -1;
  if (mainRow < 0) {
    for (let r = 0; r < Math.min(35, jsonData.length); r += 1) {
      const parts = [];
      for (let c = 0; c < scanMaxCol; c += 1) {
        const t = String(getMergedAwareCellText(r, c) || '').trim().toLowerCase();
        if (t) parts.push(t);
      }
      if (rowLooksLikeMpTableHeader(parts.join(' '))) {
        mainRow = r;
        break;
      }
    }
  }
  if (mainRow < 0) return null;

  let startCol = findSlNoColumnInSheetRow(mainRow, getMergedAwareCellText, scanMaxCol);
  if (startCol < 0) startCol = 0;

  let sectionRow = -1;
  for (let r = mainRow + 1; r <= Math.min(mainRow + 5, jsonData.length - 1); r += 1) {
    if (rowLooksLikeStatutoryColumnIndexRow(r, startCol, getMergedAwareCellText, 50, 50)) {
      sectionRow = r;
      break;
    }
  }

  const dataStartRow = sectionRow >= 0 ? sectionRow + 1 : mainRow + 2;
  const bandEnd = Math.min(
    sectionRow >= 0 ? sectionRow + 3 : mainRow + 8,
    jsonData.length - 1
  );

  const pickColumnHeaderLabel = (c, sectionIdx) => {
    const rawMain = String(readRaw(mainRow, c) || '').trim();
    const findLeafLabel = () => {
      for (let r = Math.max(0, mainRow - 2); r <= bandEnd; r += 1) {
        if (r === mainRow) continue;
        const leaf = String(readRaw(r, c) || '').trim();
        if (leaf && !isEffectivelyBlankHeaderCell(leaf) && !/^\d+$/.test(leaf)) {
          return leaf.replace(/\s+/g, ' ').trim();
        }
      }
      return '';
    };
    // Merged parent "Other Deductions Like EPF/…" — prefer leaf PF/ESIC/PT/LWF when present.
    if (rawMain && isFormXVIIIMPOtherDeductionsGroupParentHeader(rawMain)) {
      const leaf = findLeafLabel();
      if (leaf && !isFormXVIIIMPOtherDeductionsGroupParentHeader(leaf)) {
        return leaf;
      }
      return rawMain.replace(/\s+/g, ' ').trim();
    }
    if (rawMain && !isEffectivelyBlankHeaderCell(rawMain) && !/^\d+$/.test(rawMain)) {
      return rawMain.replace(/\s+/g, ' ').trim();
    }
    const leafFallback = findLeafLabel();
    if (leafFallback) return leafFallback;
    const merged = String(getMergedAwareCellText(mainRow, c) || '').trim();
    if (merged && !isEffectivelyBlankHeaderCell(merged) && !/^\d+$/.test(merged)) {
      if (isFormXVIIIMPOtherDeductionsGroupParentHeader(merged)) {
        const leaf = findLeafLabel();
        if (leaf && !isFormXVIIIMPOtherDeductionsGroupParentHeader(leaf)) return leaf;
      }
      let label = merged.replace(/\s+/g, ' ').trim();
      if (sectionIdx && /^\d+$/.test(sectionIdx) && !/\(\s*\d+\s*\)/.test(label)) {
        label = `${label} (${sectionIdx})`;
      }
      return label;
    }
    if (sectionIdx && /^\d+$/.test(sectionIdx)) return `Column ${sectionIdx}`;
    return '';
  };

  const headers = [];
  const excelCols = [];
  const maxCols = Math.max(effectiveSheetCols || 0, startCol + 60);
  const endCol = inferMPTableEndCol({
    mainRow,
    startCol,
    getMergedAwareCellText,
    getRawCellText: readRaw,
    maxCols,
    dataStartRow,
    bandEndRow: bandEnd
  });

  for (let c = startCol; c <= endCol; c += 1) {
    // Prefer raw section index so merged cells do not copy "13" into column 14.
    const sectionIdx =
      sectionRow >= 0
        ? normalizeStatutoryColumnIndexCell(
            readRaw(sectionRow, c) || getMergedAwareCellText(sectionRow, c)
          )
        : '';
    let label = pickColumnHeaderLabel(c, sectionIdx);
    const prevLabel = headers.length > 0 ? headers[headers.length - 1] : '';
    if (!label) {
      if (headers.length === 0) continue;
      label =
        sectionIdx && /^\d+$/.test(sectionIdx)
          ? `Column ${sectionIdx}`
          : `Column ${c - startCol + 1}`;
    } else if (sectionIdx && /^\d+$/.test(sectionIdx) && !/\(\s*\d+\s*\)/.test(label)) {
      const prevBase = stripTrailingSectionIndex(prevLabel);
      const curBase = stripTrailingSectionIndex(label);
      if (prevBase && prevBase === curBase) {
        label = `${label} (${sectionIdx})`;
      }
    }
    // Fix merged-header bleed: col 14 "Other allowances" must not inherit Wage rate text.
    label = resolveFormXVIIIMPSectionHeaderLabel(label, sectionIdx, prevLabel) || label;
    headers.push(label);
    excelCols.push(c);
  }

  if (headers.length < 4) return null;
  const repairedHeaders = repairFormXVIIIMPDuplicateWageHeaders(headers);
  return {
    headers: repairedHeaders,
    excelCols,
    headerRowIndex: mainRow,
    dataStartIndex: dataStartRow,
    tableStartCol: startCol,
    sectionRowIndex: sectionRow
  };
}

/** Parse Zoho People `tabularSections` (string or object) from employee record. */
export function parseEmployeeTabularSections(emp) {
  const sources = [emp, emp?.Employee, emp?.employee].filter(Boolean);
  for (let i = 0; i < sources.length; i += 1) {
    const src = sources[i];
    const raw = src.tabularSections ?? src.TabularSections ?? src.tabularsections;
    if (raw == null || raw === '') continue;
    if (typeof raw === 'string') {
      try {
        const parsed = JSON.parse(raw);
        if (parsed && typeof parsed === 'object') return parsed;
      } catch (_) {
        /* ignore */
      }
      continue;
    }
    if (typeof raw === 'object') return raw;
  }
  return {};
}

/** Format Education Details (and optional Skill sections) for MP Education/ Skill column. */
export function formatEducationSkillFromEmployee(emp) {
  const tabs = parseEmployeeTabularSections(emp);
  if (!tabs || typeof tabs !== 'object') return '';

  const parts = [];
  const eduKey =
    Object.keys(tabs).find((k) => /education\s*details/i.test(String(k || ''))) || 'Education Details';
  const eduList = tabs[eduKey];
  if (Array.isArray(eduList)) {
    eduList.forEach((edu) => {
      if (!edu || typeof edu !== 'object') return;
      const degree = String(edu.Degree ?? edu.degree ?? edu.Qualification ?? edu.qualification ?? '').trim();
      const spec = String(
        edu.Specialization ?? edu.specialization ?? edu.Field ?? edu.field ?? edu['Field of Study'] ?? ''
      ).trim();
      const inst = String(
        edu.University ?? edu.university ?? edu.College ?? edu.college ?? edu.Institution ?? edu.institution ?? ''
      ).trim();
      const line = [degree, spec].filter(Boolean).join(' - ');
      const full = [line, inst].filter(Boolean).join(', ');
      if (full) parts.push(full);
    });
  }

  const skillKey = Object.keys(tabs).find((k) => /skill/i.test(String(k || '')) && !/education/i.test(String(k || '')));
  if (skillKey && Array.isArray(tabs[skillKey])) {
    tabs[skillKey].forEach((sk) => {
      if (!sk || typeof sk !== 'object') return;
      const name = String(
        sk.Skill ?? sk.skill ?? sk.Skill_Name ?? sk.skill_name ?? sk.name ?? sk.Name ?? ''
      ).trim();
      if (name) parts.push(name);
    });
  }

  return parts.join('; ');
}

/** MP Combined Register — default total days worked when payroll lacks paid_days. */
export const FORM_XVIII_MP_DEFAULT_DAYS_WORKED = '30';

/** Fixed Category of Leave text for Form XVIII MP. */
export const FORM_XVIII_MP_LEAVE_CATEGORY =
  'Earned Leave,Legacy Earned Leave,Paternity Leave,Contingency Leave';

/** NIL defaults for OT / maternity / advances / fines / other amount columns. */
export const FORM_XVIII_MP_NIL = 'NIL';

/** Leave types whose booked/balance feed Leaves availed + Total Balance Leaves. */
export const FORM_XVIII_MP_LEAVE_METRIC_ALIASES = [
  'earned leave',
  'earned leave (test)',
  'earned leave(test)',
  'legacy earned leave',
  'legacy earned',
  'contingency leave',
  'contingency',
];

const mpPeopleHeaderNorm = (header) =>
  mpCombinedRegisterHeaderNorm(header).replace(/[^a-z0-9 ]/g, ' ').replace(/\s+/g, ' ').trim();

function parseFormXVIIIMPMoney(value) {
  if (value === '' || value == null) return NaN;
  const n = Number(String(value).replace(/,/g, '').trim());
  return Number.isFinite(n) ? n : NaN;
}

/** First finite money amount — same pattern as Form P / Form B Sample Payroll deductions. */
function pickFormXVIIIMPAmount(...values) {
  for (let i = 0; i < values.length; i += 1) {
    const value = values[i];
    if (value === '' || value == null) continue;
    const n = parseFormXVIIIMPMoney(value);
    if (Number.isFinite(n)) return n;
  }
  return '';
}

/** Other allowances = gross_pay − basic − hra (requires positive basic/hra so it never dumps full gross). */
export function computeFormXVIIIMPOtherAllowances(grossPay, basic, hra) {
  const g = parseFormXVIIIMPMoney(grossPay);
  if (!Number.isFinite(g)) return '';
  const b = parseFormXVIIIMPMoney(basic);
  const h = parseFormXVIIIMPMoney(hra);
  const hasBasic = Number.isFinite(b) && b > 0;
  const hasHra = Number.isFinite(h) && h > 0;
  if (!hasBasic && !hasHra) return '';
  const known = (hasBasic ? b : 0) + (hasHra ? h : 0);
  if (!(known > 0)) return '';
  const other = Math.round((g - known) * 100) / 100;
  // Never mirror Wage rate / gross into Other allowances.
  if (!Number.isFinite(other) || other < 0 || other === g) return '';
  return other;
}

export function isFormXVIIIMPEmpIdHeader(header) {
  const s = mpPeopleHeaderNorm(header);
  return s === 'emp id' || s === 'empid' || (s.includes('emp') && s.includes('id'));
}

export function isFormXVIIIMPWorkerNameHeader(header) {
  const s = mpPeopleHeaderNorm(header);
  return (s.includes('full name') && s.includes('worker')) || (s.includes('name') && s.includes('workman'));
}

export function isFormXVIIIMPDobHeader(header) {
  const s = mpPeopleHeaderNorm(header);
  return (s.includes('age') && s.includes('birth')) || s.includes('date of birth') || s === 'age';
}

export function isFormXVIIIMPAddressHeader(header) {
  const s = mpPeopleHeaderNorm(header);
  if (s.includes('nominee')) return false;
  if (s.includes('contractor') || s.includes('establishment') || s.includes('principal employer')) return false;
  return s === 'address' || (s.includes('address') && !s.includes('name and'));
}

export function isFormXVIIIMPSexHeader(header) {
  const s = mpPeopleHeaderNorm(header);
  return s === 'sex' || s.includes('sex m f') || (s.startsWith('sex') && s.includes('m'));
}

export function isFormXVIIIMPFatherOrHusbandHeader(header) {
  const s = mpPeopleHeaderNorm(header);
  return (s.includes('father') || s.includes('husband')) && s.includes('name');
}

export function isFormXVIIIMPNomineeHeader(header) {
  const s = mpPeopleHeaderNorm(header);
  return s.includes('nominee');
}

export function isFormXVIIIMPDesignationHeader(header) {
  const s = mpPeopleHeaderNorm(header);
  return (
    s.includes('designation') ||
    (s.includes('nature') && s.includes('work')) ||
    (s.includes('category') && s.includes('work'))
  );
}

export function isFormXVIIIMPPeopleAutofillHeader(header) {
  return (
    isFormXVIIIMPEmpIdHeader(header) ||
    isFormXVIIIMPWorkerNameHeader(header) ||
    isFormXVIIIMPDobHeader(header) ||
    isFormXVIIIMPAddressHeader(header) ||
    isFormXVIIIMPEducationSkillHeader(header) ||
    isFormXVIIIMPSexHeader(header) ||
    isFormXVIIIMPFatherOrHusbandHeader(header) ||
    isFormXVIIIMPNomineeHeader(header) ||
    isFormXVIIIMPDesignationHeader(header) ||
    isFormXVIIIMPDaysWorkedHeader(header)
  );
}

function unwrapFormXVIIIMPEmployee(empItem) {
  return empItem?.Employee || empItem?.employee || empItem || {};
}

function pickFormXVIIIMPEmployeeValue(emp, keys) {
  const src = unwrapFormXVIIIMPEmployee(emp);
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

export function isFormXVIIIMPSerialNoHeader(header) {
  const s = mpPeopleHeaderNorm(header);
  return (
    s === 'sr no' ||
    s === 's no' ||
    s === 'sl no' ||
    s === 'serial no' ||
    (s.includes('serial') && !s.includes('register') && !s.includes('workmen') && !s.includes('workman'))
  );
}

export function readFormXVIIIMPEmployeeId(emp = {}) {
  return pickFormXVIIIMPEmployeeValue(emp, [
    'gidNumber',
    'GIDNumber',
    'gid_number',
    'GID Number',
    'EmployeeID',
    'Employee ID',
    'Employee_ID',
    'EmployeeId',
    'employeeId',
    'Employee.ID',
    'employee_number',
    'Employee Number',
    'Employee_Number',
    'EmployeeCode',
    'Employee Code',
    'erecno',
    'Erecno',
    'Zoho_ID',
    'Zoho_ID',
    'Role.ID',
  ]);
}

export function readFormXVIIIMPWorkerFullName(emp = {}) {
  const fn = pickFormXVIIIMPEmployeeValue(emp, [
    'FirstName',
    'First Name',
    'First_Name',
    'firstName',
    'Name',
    'Name1',
  ]);
  const ln = pickFormXVIIIMPEmployeeValue(emp, ['LastName', 'Last Name', 'Last_Name', 'lastName', 'Surname']);
  if (fn && ln) return `${fn} ${ln}`;
  return (
    fn ||
    ln ||
    pickFormXVIIIMPEmployeeValue(emp, ['DisplayName', 'EmployeeName', 'Employee Name', 'Employee_Name', 'Full Name'])
  );
}

export function readFormXVIIIMPDateOfBirth(emp = {}, formatDateFn = (v) => String(v ?? '').trim()) {
  const raw = pickFormXVIIIMPEmployeeValue(emp, [
    'Date_of_birth',
    'Date of Birth',
    'DateofBirth',
    'Dateofbirth',
    'DOB',
    'dob',
  ]);
  return formatDateFn(raw);
}

export function readFormXVIIIMPAddress(emp = {}) {
  const line = pickFormXVIIIMPEmployeeValue(emp, [
    'PresentAddress',
    'Present Address',
    'Present_Address',
    'Address_Line_1',
    'Address Line 1',
    'PermanentAddress',
    'Permanent Address',
    'Permanent_Address',
    'Address',
  ]);
  if (line) return line;
  const city = pickFormXVIIIMPEmployeeValue(emp, ['City', 'City1', 'Present City', 'Work_location', 'LocationName']);
  const state = pickFormXVIIIMPEmployeeValue(emp, ['State', 'State1']);
  return [city, state].filter(Boolean).join(', ');
}

export function readFormXVIIIMPGender(emp = {}) {
  const raw = pickFormXVIIIMPEmployeeValue(emp, ['Gender', 'Sex', 'gender', 'sex']);
  if (!raw) return '';
  const low = raw.toLowerCase();
  if (low === 'm' || low === 'male' || low.startsWith('male')) return 'M';
  if (low === 'f' || low === 'female' || low.startsWith('female')) return 'F';
  return raw;
}

export function readFormXVIIIMPFatherOrHusbandName(emp = {}) {
  return pickFormXVIIIMPEmployeeValue(emp, [
    'Father_s_Name',
    "Father's Name",
    'Father_s Name',
    'FatherName',
    'Father Name',
    'HusbandName',
    'Husband Name',
    'SpouseName',
    'Spouse Name',
  ]);
}

export function readFormXVIIIMPDesignation(emp = {}) {
  return pickFormXVIIIMPEmployeeValue(emp, [
    'Designation',
    'Designation.displayValue',
    'designation',
    'JobTitle',
    'Job Title',
    'Department',
    'department',
  ]);
}

export function isFormXVIIIMPEducationSkillHeader(header) {
  const s = mpCombinedRegisterHeaderNorm(header).replace(/[^a-z0-9 ]/g, ' ').replace(/\s+/g, ' ').trim();
  return s.includes('education') && (s.includes('skill') || s.includes('skil'));
}

export function isFormXVIIIMPGrossWagesHeader(header) {
  const s = mpCombinedRegisterHeaderNorm(header).replace(/[^a-z0-9 ]/g, ' ').replace(/\s+/g, ' ').trim();
  if (s.includes('net') || s.includes('overtime') || (s.includes('wage') && s.includes('rate'))) return false;
  if (s.includes('gross') && (s.includes('wage') || s.includes('earning'))) return true;
  if (s.includes('total') && s.includes('gross')) return true;
  return s.includes('total') && s.includes('wage') && s.includes('earning');
}

export function isFormXVIIIMPNetPayableHeader(header) {
  const s = mpCombinedRegisterHeaderNorm(header).replace(/[^a-z0-9 ]/g, ' ').replace(/\s+/g, ' ').trim();
  return s.includes('net') && (s.includes('payable') || s.includes('paid') || s.includes('amount'));
}

export function isFormXVIIIMPOtherAllowanceHeader(header) {
  const s = mpCombinedRegisterHeaderNorm(header).replace(/[^a-z0-9 ]/g, ' ').replace(/\s+/g, ' ').trim();
  return s.includes('other') && s.includes('allowance');
}

export function isFormXVIIIMPWageRateHeader(header) {
  const s = mpCombinedRegisterHeaderNorm(header).replace(/[^a-z0-9 ]/g, ' ').replace(/\s+/g, ' ').trim();
  if (s.includes('overtime') || s.includes('allowance') || s.includes('gross') || s.includes('earning')) {
    return false;
  }
  if (s.includes('net') || s.includes('deduction') || s.includes('total')) return false;
  return (
    (s.includes('wage') && (s.includes('rate') || s.includes('pay') || s.includes('piece'))) ||
    (s.includes('piece') && s.includes('rate'))
  );
}

export function isFormXVIIIMPDaysWorkedHeader(header) {
  const s = mpCombinedRegisterHeaderNorm(header).replace(/[^a-z0-9 ]/g, ' ').replace(/\s+/g, ' ').trim();
  return (
    (s.includes('total') && s.includes('day') && s.includes('work')) ||
    s.includes('no of days worked') ||
    (s.includes('days') && s.includes('worked'))
  );
}

export function isFormXVIIIMPLeaveCategoryHeader(header) {
  const s = mpCombinedRegisterHeaderNorm(header).replace(/[^a-z0-9 ]/g, ' ').replace(/\s+/g, ' ').trim();
  return s.includes('category') && s.includes('leave');
}

export function isFormXVIIIMPLeavesAvailedHeader(header) {
  const s = mpCombinedRegisterHeaderNorm(header).replace(/[^a-z0-9 ]/g, ' ').replace(/\s+/g, ' ').trim();
  return (
    (s.includes('leave') && (s.includes('avail') || s.includes('availed'))) ||
    (s.includes('availed') && s.includes('day'))
  );
}

export function isFormXVIIIMPLeaveBalanceHeader(header) {
  const s = mpCombinedRegisterHeaderNorm(header).replace(/[^a-z0-9 ]/g, ' ').replace(/\s+/g, ' ').trim();
  return s.includes('total') && s.includes('balance') && s.includes('leave');
}

export function isFormXVIIIMPOvertimeHoursHeader(header) {
  const s = mpCombinedRegisterHeaderNorm(header)
    .replace(/[^a-z0-9 ]/g, ' ')
    .replace(/\s+/g, ' ')
    .replace(/\bover\s+time\b/g, 'overtime')
    .trim();
  return s.includes('overtime') && (s.includes('hour') || s.includes('worked'));
}

export function isFormXVIIIMPOvertimeWagesHeader(header) {
  const s = mpCombinedRegisterHeaderNorm(header)
    .replace(/[^a-z0-9 ]/g, ' ')
    .replace(/\s+/g, ' ')
    .replace(/\bover\s+time\b/g, 'overtime')
    .trim();
  if (s.includes('hour') || s.includes('worked')) return false;
  return s.includes('overtime') && (s.includes('wage') || s.includes('amount'));
}

export function isFormXVIIIMPMaternityBenefitHeader(header) {
  const s = mpCombinedRegisterHeaderNorm(header).replace(/[^a-z0-9 ]/g, ' ').replace(/\s+/g, ' ').trim();
  return s.includes('maternity');
}

export function isFormXVIIIMPAnyOtherAmountHeader(header) {
  const s = mpCombinedRegisterHeaderNorm(header).replace(/[^a-z0-9 ]/g, ' ').replace(/\s+/g, ' ').trim();
  return (
    (s.includes('any') && s.includes('other') && s.includes('amount')) ||
    (s.includes('other') && s.includes('amount') && s.includes('mention'))
  );
}

export function isFormXVIIIMPAdvancesHeader(header) {
  const s = mpCombinedRegisterHeaderNorm(header).replace(/[^a-z0-9 ]/g, ' ').replace(/\s+/g, ' ').trim();
  return s.includes('advance') || s.includes('loan');
}

export function isFormXVIIIMPFinesHeader(header) {
  const s = mpCombinedRegisterHeaderNorm(header).replace(/[^a-z0-9 ]/g, ' ').replace(/\s+/g, ' ').trim();
  return s.includes('fine');
}

export function isFormXVIIIMPNilDefaultHeader(header) {
  return (
    isFormXVIIIMPOvertimeHoursHeader(header) ||
    isFormXVIIIMPOvertimeWagesHeader(header) ||
    isFormXVIIIMPMaternityBenefitHeader(header) ||
    isFormXVIIIMPAnyOtherAmountHeader(header) ||
    isFormXVIIIMPAdvancesHeader(header) ||
    isFormXVIIIMPFinesHeader(header)
  );
}

/** PF/ESIC/PT/LWF deduction amount columns — not EPF/UAN registration columns. */
export function isFormXVIIIMPPayrollDeductionHeader(header) {
  return (
    isFormXVIIIMPPfHeader(header) ||
    isFormXVIIIMPEsicHeader(header) ||
    isFormXVIIIMPPtHeader(header) ||
    isFormXVIIIMPLwfHeader(header)
  );
}

export function readFormXVIIIMPPayrollGrossNet(payrollPayload, helpers = {}) {
  const flatFn = helpers.flattenPayrollEarningColumns || flattenPayrollEarningColumns;
  const merged = mergePayrollRunEmployeePayload(payrollPayload);
  const flat =
    merged && typeof flatFn === 'function' ? flatFn(merged) : merged || payrollPayload || {};
  const p = helpers.getPayrollPayloadObject
    ? helpers.getPayrollPayloadObject(flat)
    : flat;
  let gross =
    flat.gross_pay ??
    flat['gross_pay'] ??
    flat.total_earnings ??
    flat['total_earnings'] ??
    flat.monthly_gross_amount ??
    p.gross_pay ??
    p['gross_pay'] ??
    p.total_earnings ??
    p['total_earnings'] ??
    '';
  let net =
    flat.net_pay ??
    flat['net_pay'] ??
    flat.net_amount ??
    p.net_pay ??
    p['net_pay'] ??
    p.net_amount ??
    '';
  if ((gross === '' || gross == null) && typeof helpers.sumPayrollLineItems === 'function') {
    const earnings = helpers.getEarningsArray ? helpers.getEarningsArray(p) : [];
    const sum = helpers.sumPayrollLineItems(earnings);
    if (sum !== '') gross = sum;
  }
  if (typeof helpers.resolvePayrollAmounts === 'function') {
    const amounts = helpers.resolvePayrollAmounts(payrollPayload) || {};
    if ((gross === '' || gross == null) && amounts.grossPay !== '' && amounts.grossPay != null) {
      gross = amounts.grossPay;
    }
    if ((net === '' || net == null) && amounts.netPay !== '' && amounts.netPay != null) {
      net = amounts.netPay;
    }
  }
  if (gross === '' || gross == null) {
    const componentKeys = [
      'earned_basic',
      'basic',
      'hra',
      'hra_fbp',
      'other_allowance',
      'dearness_allowance',
      'conveyance_allowance',
      'overtime'
    ];
    let sum = 0;
    let has = false;
    componentKeys.forEach((key) => {
      const n = Number(String(flat[key] ?? p[key] ?? '').replace(/,/g, '').trim());
      if (Number.isFinite(n)) {
        sum += n;
        has = true;
      }
    });
    if (has) gross = sum;
  }
  const other_allowance_raw =
    flat.other_allowance ??
    flat['other_allowance'] ??
    p.other_allowance ??
    p['other_allowance'] ??
    readPayrollScalar(flat, ['other_allowance', 'Other Allowance', 'otherAllowance'], [/other_allowance/]) ??
    readPayrollScalar(p, ['other_allowance', 'Other Allowance', 'otherAllowance'], [/other_allowance/]) ??
    '';
  const wageParts = readPayrollForm15WageAmounts(payrollPayload);
  let basic =
    wageParts.basic !== '' && wageParts.basic != null
      ? wageParts.basic
      : flat.earned_basic ??
        flat.basic ??
        flat['earned_basic'] ??
        flat['basic'] ??
        p.earned_basic ??
        p.basic ??
        readPayrollScalar(flat, ['earned_basic', 'basic', 'Basic', 'Basic Pay'], [/^earned_basic$/, /^basic$/]) ??
        readPayrollScalar(p, ['earned_basic', 'basic', 'Basic', 'Basic Pay'], [/^earned_basic$/, /^basic$/]) ??
        '';
  let hra =
    wageParts.hra !== '' && wageParts.hra != null
      ? wageParts.hra
      : flat.hra_fbp ??
        flat.hra ??
        flat['hra_fbp'] ??
        flat['hra'] ??
        p.hra_fbp ??
        p.hra ??
        readPayrollScalar(flat, ['hra_fbp', 'hra', 'HRA', 'House Rent Allowance'], [/^hra_fbp$/, /^hra$/]) ??
        readPayrollScalar(p, ['hra_fbp', 'hra', 'HRA', 'House Rent Allowance'], [/^hra_fbp$/, /^hra$/]) ??
        '';
  if (basic === '' || basic == null || hra === '' || hra == null) {
    const earnings =
      typeof helpers.getEarningsArray === 'function'
        ? helpers.getEarningsArray(p)
        : getEarningsArray(p);
    const findAmt =
      typeof helpers.findEarningAmount === 'function'
        ? helpers.findEarningAmount
        : (list, pred) => {
            for (const item of list || []) {
              const n = String(item?.name || item?.earning_name || '').toLowerCase();
              const t = String(item?.type || item?.earning_type || '').toLowerCase();
              const amt = item?.amount ?? item?.earning_amount;
              if (pred(t, n) && amt != null && amt !== '') return amt;
            }
            return '';
          };
    if (basic === '' || basic == null) {
      basic =
        findAmt(
          earnings,
          (t, n) =>
            t === 'basic' ||
            t === 'earned_basic' ||
            n === 'basic' ||
            n.includes('earned basic') ||
            n.includes('basic pay') ||
            n.includes('basic wage')
        ) || basic;
    }
    if (hra === '' || hra == null) {
      hra =
        findAmt(
          earnings,
          (t, n) =>
            t === 'hra' ||
            t === 'hra_fbp' ||
            n === 'hra' ||
            n.includes('house rent') ||
            (n.includes('hra') && !n.includes('other'))
        ) || hra;
    }
  }
  const paid_days =
    flat.paid_days ??
    flat['paid_days'] ??
    flat.paidDays ??
    p.paid_days ??
    p['paid_days'] ??
    readPayrollScalar(
      flat,
      ['paid_days', 'Paid Days', 'paidDays', 'days_worked', 'Days Worked', 'no_of_days_worked'],
      [/^paid_days$/, /paiddays/, /daysworked/, /days_present/, /noofdayspresent/]
    ) ??
    readPayrollScalar(
      p,
      ['paid_days', 'Paid Days', 'paidDays', 'days_worked', 'Days Worked', 'no_of_days_worked'],
      [/^paid_days$/, /paiddays/, /daysworked/, /days_present/, /noofdayspresent/]
    ) ??
    '';
  const grossResolved =
    gross !== '' && gross != null
      ? gross
      : readPayrollScalar(
          flat,
          ['gross_pay', 'Gross Pay', 'grossPay', 'total_earnings', 'monthly_gross_amount'],
          [/^gross_pay$/, /^total_earnings$/]
        ) ||
        readPayrollScalar(
          p,
          ['gross_pay', 'Gross Pay', 'grossPay', 'total_earnings', 'monthly_gross_amount'],
          [/^gross_pay$/, /^total_earnings$/]
        );
  const netResolved =
    net !== '' && net != null
      ? net
      : readPayrollScalar(flat, ['net_pay', 'Net Pay', 'netPay', 'monthly_salary'], [/^net_pay$/]) ||
        readPayrollScalar(p, ['net_pay', 'Net Pay', 'netPay', 'monthly_salary'], [/^net_pay$/]);
  const otherFromFormula = computeFormXVIIIMPOtherAllowances(grossResolved, basic, hra);
  const grossNum = parseFormXVIIIMPMoney(grossResolved);
  const rawNum = parseFormXVIIIMPMoney(other_allowance_raw);
  // Never fall back to a payroll "other_allowance" that is just gross / wage-rate.
  const rawSafe =
    other_allowance_raw !== '' &&
    other_allowance_raw != null &&
    !(Number.isFinite(rawNum) && Number.isFinite(grossNum) && rawNum === grossNum && grossNum !== 0)
      ? other_allowance_raw
      : '';
  const other_allowance =
    otherFromFormula !== '' && otherFromFormula != null ? otherFromFormula : rawSafe;
  return {
    gross: grossResolved,
    net: netResolved,
    basic,
    hra,
    other_allowance,
    paid_days,
    flat,
    p,
    pf: resolveFormXVIIIMPSamplePayrollPf(flat, p, payrollPayload),
    pt: resolveFormXVIIIMPSamplePayrollPt(flat, p, payrollPayload),
  };
}

/** Prefer Sample Payroll table PF column (then flattened epf_contribution / pf). */
export function resolveFormXVIIIMPSamplePayrollPf(flat = {}, p = {}, payrollPayload = null) {
  const sources = [flat, p, payrollPayload].filter((src) => src && typeof src === 'object');
  const keys = [
    'epf_contribution',
    'EPF Contribution',
    'PF',
    'pf',
    'epf',
    'EPF',
    'employee_pf',
    'Employee PF',
    'provident_fund',
    'Provident Fund',
    // Sample Payroll UI also maps employer PF into the PF column when employee EPF is blank.
    'employer_pf',
    'employer_epf',
    'employer_epf_contribution',
    'employer_amount',
  ];
  const patterns = [
    /^epf(_contribution)?$/,
    /^pf$/,
    /^provident_fund$/,
    /^employer_pf$/,
    /^employer_epf/,
  ];
  for (const src of sources) {
    const val = pickFormXVIIIMPAmount(
      src.epf_contribution,
      src.pf,
      src.PF,
      src.provident_fund,
      src.employee_pf,
      src.employer_pf,
      src.employer_epf,
      readPayrollScalar(src, keys, patterns)
    );
    if (val !== '' && val != null) return val;
  }
  return '';
}

/** Prefer Sample Payroll table Professional Tax column. */
export function resolveFormXVIIIMPSamplePayrollPt(flat = {}, p = {}, payrollPayload = null) {
  const sources = [flat, p, payrollPayload].filter((src) => src && typeof src === 'object');
  const keys = [
    'Professional Tax',
    'professional_tax',
    'ProfessionalTax',
    'professionalTax',
    'pt',
    'PT',
  ];
  const patterns = [/^professional_tax$/, /^professionaltax$/, /^pt$/];
  for (const src of sources) {
    const val = pickFormXVIIIMPAmount(
      src.professional_tax,
      src.professionalTax,
      src.ProfessionalTax,
      src['Professional Tax'],
      src.pt,
      src.PT,
      readPayrollScalar(src, keys, patterns)
    );
    if (val !== '' && val != null) return val;
  }
  return '';
}

/** Map MP Combined Register payroll columns (gross, PF/ESIC/PT/LWF, net, UTR). */
export function resolveFormXVIIIMPTableHeaders(headers) {
  const list = Array.isArray(headers) ? headers : [];
  const norm = (h) =>
    mpCombinedRegisterHeaderNorm(h)
      .replace(/[^a-z0-9 ]/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  const findHeader = (testFn) =>
    list.find((h) => {
      const s = norm(h);
      return s && testFn(s, h);
    }) || null;
  return {
    grossWages: findHeader((s) => {
      if (s.includes('net') || s.includes('overtime') || (s.includes('wage') && s.includes('rate'))) {
        return false;
      }
      if (s.includes('gross') && (s.includes('wage') || s.includes('earning'))) return true;
      if (s.includes('total') && s.includes('gross')) return true;
      return s.includes('total') && s.includes('wage') && s.includes('earning');
    }),
    overtimeWages: findHeader((s) => {
      const t = s.replace(/\bover\s+time\b/g, 'overtime');
      return t.includes('overtime') && (t.includes('wage') || t.includes('amount')) && !t.includes('hour') && !t.includes('worked');
    }),
    overtimeHours: findHeader((s) => {
      const t = s.replace(/\bover\s+time\b/g, 'overtime');
      return t.includes('overtime') && (t.includes('hour') || t.includes('worked'));
    }),
    otherAllowances: findHeader((s) => s.includes('other') && s.includes('allowance')),
    pf: findHeader((_s, h) => isFormXVIIIMPPfHeader(h)),
    esic: findHeader((_s, h) => isFormXVIIIMPEsicHeader(h)),
    pt: findHeader((_s, h) => isFormXVIIIMPPtHeader(h)),
    lwf: findHeader((_s, h) => isFormXVIIIMPLwfHeader(h)),
    fines: findHeader((s) => s.includes('fine')),
    advances: findHeader((s) => s.includes('advance') || s.includes('loan')),
    maternityBenefit: findHeader((s) => s.includes('maternity')),
    anyOtherAmount: findHeader(
      (s) =>
        (s.includes('any') && s.includes('other') && s.includes('amount')) ||
        (s.includes('other') && s.includes('amount') && s.includes('mention'))
    ),
    netPayable: findHeader((s) => s.includes('net') && (s.includes('payable') || s.includes('paid'))),
    bankUtr: findHeader((s) => s.includes('bank') && s.includes('utr')),
    educationSkill: findHeader((s) => s.includes('education') && (s.includes('skill') || s.includes('skil'))),
    empId: findHeader((s) => s === 'emp id' || s === 'empid' || (s.includes('emp') && s.includes('id'))),
    workerName: findHeader(
      (s) => (s.includes('full name') && s.includes('worker')) || (s.includes('name') && s.includes('workman'))
    ),
    dob: findHeader(
      (s) => (s.includes('age') && s.includes('birth')) || s.includes('date of birth') || s === 'age'
    ),
    address: findHeader((s) => {
      if (s.includes('nominee')) return false;
      if (s.includes('contractor') || s.includes('establishment') || s.includes('principal employer')) {
        return false;
      }
      return s === 'address' || (s.includes('address') && !s.includes('name and'));
    }),
    sex: findHeader((s) => s === 'sex' || s.includes('sex m f') || (s.startsWith('sex') && s.includes('m'))),
    fatherOrHusband: findHeader(
      (s) => (s.includes('father') || s.includes('husband')) && s.includes('name')
    ),
    nominee: findHeader((s) => s.includes('nominee')),
    designation: findHeader(
      (s) =>
        s.includes('designation') ||
        (s.includes('nature') && s.includes('work')) ||
        (s.includes('category') && s.includes('work'))
    ),
    wageRate: findHeader((s) => {
      if (s.includes('allowance') || s.includes('gross') || s.includes('earning') || s.includes('net')) {
        return false;
      }
      if (s.includes('deduction') || s.includes('total')) return false;
      return (
        (s.includes('wage') && (s.includes('rate') || s.includes('pay') || s.includes('piece'))) ||
        (s.includes('piece') && s.includes('rate'))
      );
    }),
    daysWorked: findHeader(
      (s) =>
        (s.includes('total') && s.includes('day') && s.includes('work')) ||
        s.includes('no of days worked') ||
        (s.includes('days') && s.includes('worked'))
    ),
    leaveCategory: findHeader((s) => s.includes('category') && s.includes('leave')),
    leavesAvailed: findHeader(
      (s) =>
        (s.includes('leave') && (s.includes('avail') || s.includes('availed'))) ||
        (s.includes('availed') && s.includes('day'))
    ),
    totalBalanceLeaves: findHeader((s) => s.includes('total') && s.includes('balance') && s.includes('leave'))
  };
}

export function applyFormXVIIIMPEmployeeToRow(row, emp, mpHeaders, helpers = {}) {
  if (!row || !emp) return false;
  const empRecord = unwrapFormXVIIIMPEmployee(emp);
  const sanitizeValue = helpers.sanitizeValue || ((v) => v);
  const formatStatutoryDateDisplay = helpers.formatStatutoryDateDisplay || ((v) => String(v ?? '').trim());
  const headers = Array.isArray(helpers.headers) ? helpers.headers : [];
  let hit = false;

  const write = (header, val) => {
    if (!header || val === '' || val == null) return;
    row[header] = sanitizeValue(val);
    hit = true;
  };

  const peopleValues = {
    empId: readFormXVIIIMPEmployeeId(empRecord),
    workerName: readFormXVIIIMPWorkerFullName(empRecord),
    dob: readFormXVIIIMPDateOfBirth(empRecord, formatStatutoryDateDisplay),
    address: readFormXVIIIMPAddress(empRecord),
    educationSkill: formatEducationSkillFromEmployee(empRecord),
    sex: readFormXVIIIMPGender(empRecord),
    fatherOrHusband: readFormXVIIIMPFatherOrHusbandName(empRecord),
    designation: readFormXVIIIMPDesignation(empRecord),
    daysWorked: FORM_XVIII_MP_DEFAULT_DAYS_WORKED,
  };

  if (mpHeaders) {
    write(mpHeaders.empId, peopleValues.empId);
    write(mpHeaders.workerName, peopleValues.workerName);
    write(mpHeaders.dob, peopleValues.dob);
    write(mpHeaders.address, peopleValues.address);
    write(mpHeaders.educationSkill, peopleValues.educationSkill);
    write(mpHeaders.sex, peopleValues.sex);
    write(mpHeaders.fatherOrHusband, peopleValues.fatherOrHusband);
    write(mpHeaders.designation, peopleValues.designation);
    if (!String(row[mpHeaders.daysWorked] ?? '').trim()) {
      write(mpHeaders.daysWorked, peopleValues.daysWorked);
    }
  }

  headers.forEach((header) => {
    if (isFormXVIIIMPEmpIdHeader(header)) write(header, peopleValues.empId);
    else if (isFormXVIIIMPWorkerNameHeader(header)) write(header, peopleValues.workerName);
    else if (isFormXVIIIMPDobHeader(header)) write(header, peopleValues.dob);
    else if (isFormXVIIIMPAddressHeader(header)) write(header, peopleValues.address);
    else if (isFormXVIIIMPEducationSkillHeader(header)) write(header, peopleValues.educationSkill);
    else if (isFormXVIIIMPSexHeader(header)) write(header, peopleValues.sex);
    else if (isFormXVIIIMPFatherOrHusbandHeader(header)) write(header, peopleValues.fatherOrHusband);
    else if (isFormXVIIIMPDesignationHeader(header)) write(header, peopleValues.designation);
    else if (isFormXVIIIMPDaysWorkedHeader(header) && !String(row[header] ?? '').trim()) {
      write(header, peopleValues.daysWorked);
    }
  });

  return hit;
}

const FORM_XVIII_MP_PEOPLE_HEADER_KEYS = new Set([
  'empId',
  'workerName',
  'dob',
  'address',
  'educationSkill',
  'sex',
  'fatherOrHusband',
  'nominee',
  'designation',
]);

/** Payroll/leave columns cleared before autofill — excludes people identity columns. */
export function getFormXVIIIMPPayrollSkipHeaders(mpHeaders) {
  if (!mpHeaders || typeof mpHeaders !== 'object') return [];
  return Object.entries(mpHeaders)
    .filter(([key]) => !FORM_XVIII_MP_PEOPLE_HEADER_KEYS.has(key))
    .map(([, header]) => header)
    .filter(Boolean);
}

export function applyFormXVIIIMPPeopleField(row, header, emp, helpers = {}) {
  if (!row || !header || !emp) return false;
  const empRecord = unwrapFormXVIIIMPEmployee(emp);
  const sanitizeValue = helpers.sanitizeValue || ((v) => v);
  const formatStatutoryDateDisplay = helpers.formatStatutoryDateDisplay || ((v) => String(v ?? '').trim());

  let val = '';
  if (isFormXVIIIMPEmpIdHeader(header)) val = readFormXVIIIMPEmployeeId(empRecord);
  else if (isFormXVIIIMPWorkerNameHeader(header)) val = readFormXVIIIMPWorkerFullName(empRecord);
  else if (isFormXVIIIMPDobHeader(header)) val = readFormXVIIIMPDateOfBirth(empRecord, formatStatutoryDateDisplay);
  else if (isFormXVIIIMPAddressHeader(header)) val = readFormXVIIIMPAddress(empRecord);
  else if (isFormXVIIIMPEducationSkillHeader(header)) val = formatEducationSkillFromEmployee(empRecord);
  else if (isFormXVIIIMPSexHeader(header)) val = readFormXVIIIMPGender(empRecord);
  else if (isFormXVIIIMPFatherOrHusbandHeader(header)) val = readFormXVIIIMPFatherOrHusbandName(empRecord);
  else if (isFormXVIIIMPDesignationHeader(header)) val = readFormXVIIIMPDesignation(empRecord);
  else if (isFormXVIIIMPDaysWorkedHeader(header)) {
    if (String(row[header] ?? '').trim()) return false;
    val = FORM_XVIII_MP_DEFAULT_DAYS_WORKED;
  } else return false;

  if (val === '' || val == null) return false;
  row[header] = sanitizeValue(val);
  return true;
}

export function enrichFormXVIIIMPEmployeeRows(mappedData, employees, headers, helpers = {}) {
  if (!Array.isArray(mappedData) || mappedData.length === 0) return 0;
  const mpHeaders = resolveFormXVIIIMPTableHeaders(headers);
  const headerList = Array.isArray(headers) ? headers : [];
  let hits = 0;
  mappedData.forEach((row, rowIndex) => {
    const empItem = employees?.[rowIndex];
    if (!empItem) return;
    if (
      applyFormXVIIIMPEmployeeToRow(row, empItem, mpHeaders, {
        ...helpers,
        headers: headerList,
      })
    ) {
      hits += 1;
    }
    applyFormXVIIIMPLeaveCategoryToRow(row, headerList, helpers);
    applyFormXVIIIMPNilDefaultsToRow(row, headerList, {
      nilText: FORM_XVIII_MP_NIL,
      overwrite: true,
    });
  });
  return hits;
}

export function applyFormXVIIIMPPayrollToRow(row, payrollPayload, mpHeaders, helpers = {}) {
  if (!row || !payrollPayload) return false;
  const sanitizeValue = helpers.sanitizeValue || ((v) => v);
  const { gross, net, other_allowance, paid_days, p, pf: samplePf, pt: samplePt } =
    readFormXVIIIMPPayrollGrossNet(payrollPayload, helpers);
  const headers = Array.isArray(helpers.headers) ? helpers.headers : Object.keys(row || {});
  let hit = false;

  const writeAmount = (header, val) => {
    if (!header || val === '' || val == null) return;
    row[header] = sanitizeValue(val);
    hit = true;
  };
  const writeOtherAllowance = (header, val) => {
    if (!header) return;
    const g = parseFormXVIIIMPMoney(gross);
    const v = parseFormXVIIIMPMoney(val);
    const existing = parseFormXVIIIMPMoney(row[header]);
    // Do not write Wage rate / gross into Other allowances.
    if (Number.isFinite(v) && Number.isFinite(g) && v === g && g !== 0) {
      if (Number.isFinite(existing) && existing === g) {
        row[header] = '';
        hit = true;
      }
      return;
    }
    // Clear a previously mirrored gross when formula could not be resolved.
    if (
      (val === '' || val == null) &&
      Number.isFinite(existing) &&
      Number.isFinite(g) &&
      existing === g &&
      g !== 0
    ) {
      row[header] = '';
      hit = true;
      return;
    }
    writeAmount(header, val);
  };
  const daysValue =
    paid_days !== '' && paid_days != null ? paid_days : FORM_XVIII_MP_DEFAULT_DAYS_WORKED;

  headers.forEach((h) => {
    if (isFormXVIIIMPGrossWagesHeader(h)) writeAmount(h, gross);
    if (isFormXVIIIMPNetPayableHeader(h)) writeAmount(h, net);
    if (isFormXVIIIMPOtherAllowanceHeader(h)) writeOtherAllowance(h, other_allowance);
    if (isFormXVIIIMPWageRateHeader(h)) writeAmount(h, gross);
    if (isFormXVIIIMPDaysWorkedHeader(h)) writeAmount(h, daysValue);
    if (isFormXVIIIMPLeaveCategoryHeader(h)) writeAmount(h, FORM_XVIII_MP_LEAVE_CATEGORY);
  });

  if (mpHeaders) {
    writeAmount(mpHeaders.grossWages, gross);
    writeAmount(mpHeaders.netPayable, net);
    writeOtherAllowance(mpHeaders.otherAllowances, other_allowance);
    writeAmount(mpHeaders.wageRate, gross);
    writeAmount(mpHeaders.daysWorked, daysValue);
    writeAmount(mpHeaders.leaveCategory, FORM_XVIII_MP_LEAVE_CATEGORY);
  }

  const findDeductionAmount = helpers.findDeductionAmount;
  const deductions =
    typeof findDeductionAmount === 'function' && helpers.getDeductionsArray
      ? helpers.getDeductionsArray(p)
      : typeof findDeductionAmount === 'function'
        ? []
        : null;
  const otherAllow = other_allowance;
  // PF (22) ← Sample Payroll PF; PT (22) ← Sample Payroll Professional Tax.
  // Fall back to Zoho deduction line items only when Sample Payroll columns are empty.
  const pfFromDeductions =
    typeof findDeductionAmount === 'function'
      ? findDeductionAmount(
          deductions || [],
          (t, n) => t === 'pf' || t === 'epf' || n.includes('provident') || n.includes('epf')
        ) || ''
      : '';
  const ptFromDeductions =
    typeof findDeductionAmount === 'function'
      ? findDeductionAmount(
          deductions || [],
          (t, n) => t === 'pt' || n.includes('professional tax')
        ) || ''
      : '';
  const pf =
    samplePf !== '' && samplePf != null
      ? samplePf
      : pfFromDeductions;
  const esic =
    typeof findDeductionAmount === 'function'
      ? findDeductionAmount(
          deductions || [],
          (t, n) => t === 'esi' || n.includes('esic') || n.includes('employee state insurance')
        ) || ''
      : '';
  const pt =
    samplePt !== '' && samplePt != null
      ? samplePt
      : ptFromDeductions;
  const lwf =
    typeof findDeductionAmount === 'function'
      ? findDeductionAmount(
          deductions || [],
          (t, n) => n.includes('labour welfare') || n.includes('labor welfare') || n.includes('lwf')
        ) || ''
      : '';

  headers.forEach((h) => {
    if (isFormXVIIIMPPfHeader(h)) writeAmount(h, pf);
    if (isFormXVIIIMPEsicHeader(h)) writeAmount(h, esic);
    if (isFormXVIIIMPPtHeader(h)) writeAmount(h, pt);
    if (isFormXVIIIMPLwfHeader(h)) writeAmount(h, lwf);
  });

  // Also write PF/PT onto leftover merged "Other Deductions…" / composite keys still present on the row.
  if (pf !== '' && pf != null) {
    Object.keys(row || {}).forEach((key) => {
      if (isFormXVIIIMPOtherDeductionsGroupParentHeader(key) || isFormXVIIIMPPfHeader(key)) {
        writeAmount(key, pf);
      }
    });
  }
  if (pt !== '' && pt != null) {
    Object.keys(row || {}).forEach((key) => {
      if (isFormXVIIIMPPtHeader(key)) {
        writeAmount(key, pt);
      }
    });
  }

  const set = (key, val) => {
    if (!mpHeaders?.[key] || val === '' || val == null) return;
    row[mpHeaders[key]] = sanitizeValue(val);
    hit = true;
  };

  set('grossWages', gross);
  writeOtherAllowance(mpHeaders?.otherAllowances, otherAllow);
  set('pf', pf);
  set('esic', esic);
  set('pt', pt);
  set('lwf', lwf);
  set('netPayable', net);
  set('wageRate', gross);
  set('daysWorked', paid_days);
  set('leaveCategory', FORM_XVIII_MP_LEAVE_CATEGORY);
  applyFormXVIIIMPNilDefaultsToRow(row, headers, { nilText: FORM_XVIII_MP_NIL, overwrite: true });
  return hit;
}

/** Build Form XVIII MP grid rows from People (+ optional cached payroll) — fast modal open path. */
export function mapFormXVIIIMPRowsFromEmployees(employees, headers, helpers = {}) {
  const mpHeaders = resolveFormXVIIIMPTableHeaders(headers);
  const headerList = Array.isArray(headers) ? headers : [];
  const list = Array.isArray(employees) ? employees : [];
  const {
    sanitizeValue = (v) => String(v ?? '').trim(),
    formatStatutoryDateDisplay = (v) => String(v ?? '').trim(),
    resolvePayrollRow = null,
    rowIndexOffset = 0,
    ...payrollHelpers
  } = helpers;

  return list.map((empItem, rowIndex) => {
    const row = {};
    headerList.forEach((header) => {
      row[header] = '';
    });
    const globalRowIndex = rowIndexOffset + rowIndex;
    headerList.forEach((header) => {
      if (isFormXVIIIMPSerialNoHeader(header)) {
        row[header] = String(globalRowIndex + 1);
      }
    });
    applyFormXVIIIMPEmployeeToRow(row, empItem, mpHeaders, {
      sanitizeValue,
      formatStatutoryDateDisplay,
      headers: headerList,
    });
    applyFormXVIIIMPLeaveCategoryToRow(row, headerList, { sanitizeValue });
    applyFormXVIIIMPNilDefaultsToRow(row, headerList, { nilText: FORM_XVIII_MP_NIL, overwrite: true });
    const emp = unwrapFormXVIIIMPEmployee(empItem);
    const payrollRow =
      typeof resolvePayrollRow === 'function' ? resolvePayrollRow(emp, rowIndex) : null;
    if (payrollRow && !payrollRow.fetch_error) {
      applyFormXVIIIMPPayrollToRow(row, payrollRow, mpHeaders, {
        sanitizeValue,
        headers: headerList,
        ...payrollHelpers,
      });
    }
    return row;
  });
}

/** Apply pay-run payroll fields onto Form XVIII MP grid rows. */
export function enrichFormXVIIIMPPayrollRows(mappedData, employees, headers, helpers = {}) {
  if (!Array.isArray(mappedData) || mappedData.length === 0) return 0;
  const mpHeaders = resolveFormXVIIIMPTableHeaders(headers);
  const {
    payrollRows = [],
    sanitizeValue = (v) => String(v ?? '').trim(),
    resolvePayrollRow = null,
    ...payrollHelpers
  } = helpers;
  const rows = Array.isArray(payrollRows) ? payrollRows : [];
  if (rows.length === 0) return 0;
  const resolver = buildFormXVIIIMPPayrollRowResolver(rows);
  let hits = 0;
  mappedData.forEach((row, rowIndex) => {
    const empItem = employees[rowIndex];
    const emp = empItem?.Employee || empItem?.employee || empItem;
    const payrollRow =
      typeof resolvePayrollRow === 'function'
        ? resolvePayrollRow(emp, row, rowIndex)
        : resolveFormXVIIIMPPayrollRowForAutofillRow(emp, row, rows, headers, resolver);
    if (!payrollRow || payrollRow.fetch_error) return;
    if (
      applyFormXVIIIMPPayrollToRow(row, payrollRow, mpHeaders, {
        sanitizeValue,
        headers,
        ...payrollHelpers,
      })
    ) {
      hits += 1;
    }
  });
  return hits;
}

/** True when saved/download rows have employee names and at least one data cell. */
export function formXVIIIMPDownloadHasSubstantiveRows(mappedData, headers) {
  if (!Array.isArray(mappedData) || mappedData.length === 0) return false;
  const mpHeaders = resolveFormXVIIIMPTableHeaders(headers);
  const payrollHeaders = [
    mpHeaders.grossWages,
    mpHeaders.netPayable,
    mpHeaders.wageRate,
    mpHeaders.otherAllowances,
    mpHeaders.daysWorked,
    mpHeaders.leavesAvailed,
    mpHeaders.totalBalanceLeaves,
  ].filter(Boolean);
  return mappedData.some((row) => {
    if (!row || typeof row !== 'object') return false;
    const hasName = Object.entries(row).some(([key, value]) => {
      const s = mpCombinedRegisterHeaderNorm(key);
      return (
        s.includes('name') &&
        (s.includes('workman') || s.includes('employee') || s.includes('worker')) &&
        String(value ?? '').trim()
      );
    });
    if (!hasName) return false;
    if (payrollHeaders.some((header) => String(row[header] ?? '').trim())) return true;
    return Object.values(row).some((value) => String(value ?? '').trim());
  });
}

/** True when any row is missing gross/net/wage-rate/PF/PT, or other-allowance equals gross (bad formula). */
export function formXVIIIMPRowsNeedPayrollEnrich(mappedData, headers) {
  if (!Array.isArray(mappedData) || mappedData.length === 0) return false;
  const mpHeaders = resolveFormXVIIIMPTableHeaders(headers);
  const deductionTargets = [mpHeaders.pf, mpHeaders.pt].filter(Boolean);
  if (
    deductionTargets.length > 0 &&
    mappedData.some((row) =>
      deductionTargets.some((header) => !String(row?.[header] ?? '').trim())
    )
  ) {
    return true;
  }
  const targets = [
    mpHeaders.grossWages,
    mpHeaders.netPayable,
    mpHeaders.wageRate,
    mpHeaders.otherAllowances,
  ].filter(Boolean);
  if (targets.length === 0) return false;
  return mappedData.some((row) => {
    if (!row || typeof row !== 'object') return false;
    if (targets.every((header) => !String(row?.[header] ?? '').trim())) return true;
    const otherHdr = mpHeaders.otherAllowances;
    const grossHdr = mpHeaders.grossWages || mpHeaders.wageRate;
    if (!otherHdr || !grossHdr) return false;
    const otherVal = String(row[otherHdr] ?? '').replace(/,/g, '').trim();
    const grossVal = String(row[grossHdr] ?? '').replace(/,/g, '').trim();
    if (!otherVal || !grossVal) return !otherVal && !!grossVal;
    const otherNum = Number(otherVal);
    const grossNum = Number(grossVal);
    return Number.isFinite(otherNum) && Number.isFinite(grossNum) && otherNum === grossNum && grossNum !== 0;
  });
}

/** Fast download enrich — Payroll table session cache (sync); use loadFormXVIIIMPPayrollRowsForAutofill for real-time. */
export function enrichFormXVIIIMPPayrollRowsFromCache(
  mappedData,
  employees,
  headers,
  monthCandidates,
  helpers = {}
) {
  const cached = getCachedForm15PayrollTableRows(monthCandidates);
  const latestCached = cached?.rows?.length > 0 ? cached : getLatestCachedPayrollTableRows();
  const payrollRows = resolveFormXVIIIMPPayrollRowsForAutofill(
    latestCached?.rows?.length > 0 ? latestCached.rows.map((row) => flattenPayrollEarningColumns(row)) : null,
    monthCandidates
  );
  if (payrollRows.length === 0) return 0;
  return enrichFormXVIIIMPPayrollRows(mappedData, employees, headers, {
    ...helpers,
    payrollRows,
  });
}

/**
 * Last-mile download fix: recompute Other allowances = gross − basic − hra and
 * never leave Wage rate / gross mirrored into that column.
 */
export function finalizeFormXVIIIMPOtherAllowancesForDownload(
  mappedData,
  headers,
  employees = [],
  helpers = {}
) {
  if (!Array.isArray(mappedData) || mappedData.length === 0) return 0;
  const headerList = Array.isArray(headers) ? headers : [];
  {
    const repaired = repairFormXVIIIMPDuplicateWageHeaders(headerList);
    for (let i = 0; i < repaired.length; i += 1) headerList[i] = repaired[i];
  }
  const mpHeaders = resolveFormXVIIIMPTableHeaders(headerList);
  let otherHeader = mpHeaders.otherAllowances;
  if (!otherHeader) {
    otherHeader = headerList.find((h) => isFormXVIIIMPOtherAllowanceHeader(h)) || null;
  }
  // Ensure a writable Other allowances key exists even if headers were mislabeled.
  if (!otherHeader) {
    otherHeader = FORM_XVIII_MP_SECTION_INDEX_HEADERS[14];
    if (!headerList.includes(otherHeader)) headerList.push(otherHeader);
  }
  const wageHeader = mpHeaders.wageRate;
  const grossHeader = mpHeaders.grossWages;
  const sanitizeValue = helpers.sanitizeValue || ((v) => String(v ?? '').trim());
  const rows = Array.isArray(helpers.payrollRows) ? helpers.payrollRows : [];
  const resolver =
    typeof helpers.resolvePayrollRow === 'function'
      ? helpers.resolvePayrollRow
      : rows.length > 0
        ? buildFormXVIIIMPPayrollRowResolver(rows)
        : null;
  let hits = 0;
  mappedData.forEach((row, rowIndex) => {
    if (!row || typeof row !== 'object') return;
    const empItem = employees[rowIndex];
    const emp = empItem?.Employee || empItem?.employee || empItem;
    let payrollRow =
      typeof resolver === 'function'
        ? resolveFormXVIIIMPPayrollRowForAutofillRow(emp, row, rows, headerList, resolver)
        : null;
    if ((!payrollRow || payrollRow.fetch_error) && rows[rowIndex] && !rows[rowIndex].fetch_error) {
      payrollRow = rows[rowIndex];
    }
    let otherVal = '';
    let grossVal = '';
    if (payrollRow && !payrollRow.fetch_error) {
      const parsed = readFormXVIIIMPPayrollGrossNet(payrollRow, helpers);
      otherVal = parsed.other_allowance;
      grossVal = parsed.gross;
      if (otherVal !== '' && otherVal != null) {
        row[otherHeader] = sanitizeValue(otherVal);
        hits += 1;
      }
    }
    const gCandidates = [
      parseFormXVIIIMPMoney(grossVal),
      parseFormXVIIIMPMoney(row[grossHeader]),
      parseFormXVIIIMPMoney(row[wageHeader]),
    ];
    const g = gCandidates.find((n) => Number.isFinite(n));
    const cur = parseFormXVIIIMPMoney(row[otherHeader]);
    // Strip mirrored Wage rate / gross when formula could not produce a distinct value.
    if (Number.isFinite(cur) && Number.isFinite(g) && cur === g && g !== 0) {
      if (otherVal === '' || otherVal == null || parseFormXVIIIMPMoney(otherVal) === g) {
        row[otherHeader] = '';
        hits += 1;
      }
    }
    // Copy correct value onto every Other-allowances header variant in the row.
    headerList.forEach((h) => {
      if (!isFormXVIIIMPOtherAllowanceHeader(h)) return;
      if (h === otherHeader) return;
      if (row[otherHeader] !== '' && row[otherHeader] != null) {
        row[h] = row[otherHeader];
      }
    });
  });
  return hits;
}

/** Apply NIL to Form XVIII MP OT / maternity / advances / fines / other-amount columns. */
export function applyFormXVIIIMPNilDefaultsToRow(row, headers, helpers = {}) {
  if (!row || typeof row !== 'object') return false;
  const hdrs = Array.isArray(headers) ? headers : Object.keys(row);
  const nilText = helpers.nilText != null ? String(helpers.nilText) : FORM_XVIII_MP_NIL;
  const { overwrite = true } = helpers;
  let hit = false;
  hdrs.forEach((header) => {
    if (!isFormXVIIIMPNilDefaultHeader(header)) return;
    const existing = String(row[header] ?? '').trim();
    if (
      !overwrite &&
      existing &&
      !/^enter\b/i.test(existing) &&
      !/^nil+$/i.test(existing) &&
      existing.toLowerCase() !== 'n/a' &&
      existing !== '-' &&
      existing !== '—'
    ) {
      return;
    }
    row[header] = nilText;
    hit = true;
  });
  return hit;
}

export function applyFormXVIIIMPNilDefaultsToMappedRows(
  mappedData,
  headers,
  nilText = FORM_XVIII_MP_NIL,
  helpers = {}
) {
  if (!Array.isArray(mappedData)) return [];
  const { overwrite = true } = helpers;
  mappedData.forEach((row) => {
    applyFormXVIIIMPNilDefaultsToRow(row, headers, { nilText, overwrite });
  });
  return mappedData;
}

/** Apply fixed Category of Leave text on Form XVIII MP rows. */
export function applyFormXVIIIMPLeaveCategoryToRow(row, headers, helpers = {}) {
  if (!row || typeof row !== 'object') return false;
  const hdrs = Array.isArray(headers) ? headers : Object.keys(row);
  const categoryText =
    helpers.categoryText != null ? String(helpers.categoryText) : FORM_XVIII_MP_LEAVE_CATEGORY;
  const sanitizeValue = helpers.sanitizeValue || ((v) => v);
  let hit = false;
  hdrs.forEach((header) => {
    if (!isFormXVIIIMPLeaveCategoryHeader(header)) return;
    row[header] = sanitizeValue(categoryText);
    hit = true;
  });
  return hit;
}

/**
 * Sum booked/balance for Earned Leave + Legacy Earned Leave + Contingency Leave only.
 * (Category of Leave also lists Paternity Leave, but availed/balance exclude it.)
 */
export function sumFormXVIIIMPLeaveBookedAndBalance(leaveRecord, leaveTypeLabels = {}) {
  if (!leaveRecord || typeof leaveRecord !== 'object') {
    return { booked: '', balance: '' };
  }
  let bookedTotal = 0;
  let balanceTotal = 0;
  let hasBooked = false;
  let hasBalance = false;
  const skipKeys = new Set([
    's.no',
    'sno',
    'employeeid',
    'employee_id',
    'employee',
    'employeename',
    'zohoid',
    'zoho_id',
    'zoho.id',
  ]);

  Object.entries(leaveRecord).forEach(([key, value]) => {
    const kl = String(key || '')
      .toLowerCase()
      .replace(/\s+/g, ' ')
      .trim();
    if (skipKeys.has(kl.replace(/\s+/g, '')) || skipKeys.has(kl)) return;
    const label = String(leaveTypeLabels?.[key] || key || '').trim();
    if (!leaveTypeLabelMatchesAliases(label, FORM_XVIII_MP_LEAVE_METRIC_ALIASES)) return;
    const obj = parseLeaveCellObject(value) ?? (value && typeof value === 'object' ? value : null);
    if (!obj || typeof obj !== 'object') {
      const plain = Number(String(value ?? '').trim());
      if (Number.isFinite(plain)) {
        // Plain numeric cells (e.g. Legacy Earned Leave = 0) count as balance when present.
        balanceTotal += plain;
        hasBalance = true;
      }
      return;
    }
    const bookRaw = obj.paidBooked ?? obj.booked ?? obj.Booked ?? obj.unpaidBooked;
    const balRaw = obj.paidBalance ?? obj.balance ?? obj.Balance ?? obj.unpaidBalance;
    const bookNum = Number(String(sanitizeLeaveMetricDisplayValue(bookRaw) || '').trim());
    const balNum = Number(String(sanitizeLeaveMetricDisplayValue(balRaw) || '').trim());
    if (Number.isFinite(bookNum)) {
      bookedTotal += bookNum;
      hasBooked = true;
    }
    if (Number.isFinite(balNum)) {
      balanceTotal += balNum;
      hasBalance = true;
    }
  });

  return {
    booked: hasBooked ? String(bookedTotal) : '',
    balance: hasBalance ? String(balanceTotal) : '',
  };
}

/** Apply leave API booked/balance totals to Form XVIII MP leave columns. */
export function applyFormXVIIIMPLeaveToRow(row, leaveRecord, mpHeaders, { leaveTypeLabels = {}, sanitizeValue = (v) => v, headers = [] } = {}) {
  if (!row || !mpHeaders) return false;
  let hit = false;
  const write = (header, val) => {
    if (!header || val === '' || val == null) return;
    row[header] = sanitizeValue(val);
    hit = true;
  };

  write(mpHeaders.leaveCategory, FORM_XVIII_MP_LEAVE_CATEGORY);
  const headerList = Array.isArray(headers) && headers.length ? headers : Object.keys(row);
  applyFormXVIIIMPLeaveCategoryToRow(row, headerList, { sanitizeValue });

  if (!leaveRecord) return hit;
  const { booked, balance } = sumFormXVIIIMPLeaveBookedAndBalance(leaveRecord, leaveTypeLabels);
  write(mpHeaders.leavesAvailed, booked);
  write(mpHeaders.totalBalanceLeaves, balance);
  return hit;
}

function readWorkbookSheetHelpers(workbook, sheetName) {
  const sheetNames = workbook?.SheetNames || [];
  const resolvedName =
    sheetName && workbook?.Sheets?.[sheetName] ? sheetName : sheetNames[0] || '';
  const worksheet = resolvedName ? workbook.Sheets[resolvedName] : null;
  const merges = worksheet?.['!merges'] || [];
  const rawCell = (r, c) => {
    if (r < 0 || c < 0 || !worksheet) return '';
    const ref = XLSX.utils.encode_cell({ r, c });
    const cell = worksheet[ref];
    return cell && cell.v != null ? String(cell.v).trim() : '';
  };
  const getMergedAwareCellText = (r, c) => {
    const direct = rawCell(r, c);
    if (direct) return direct;
    for (let i = 0; i < merges.length; i += 1) {
      const m = merges[i];
      if (!m?.s || !m?.e) continue;
      if (r >= m.s.r && r <= m.e.r && c >= m.s.c && c <= m.e.c) {
        const topLeft = rawCell(m.s.r, m.s.c);
        if (topLeft) return topLeft;
      }
    }
    return '';
  };
  const getRawCellText = rawCell;
  const jsonData = worksheet
    ? XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' })
    : [];
  const effectiveSheetCols = Math.max(
    ...(jsonData || []).map((row) => (Array.isArray(row) ? row.length : 0)),
    40
  );
  return { worksheet, jsonData, merges, getMergedAwareCellText, getRawCellText, effectiveSheetCols, resolvedName };
}

export function resolveMPCombinedRegisterVisibleTableColumns({
  workbook,
  sheetName,
  headerRowIndex = -1,
  fallbackHeaders = [],
  fallbackExcelCols = null,
  dataStartIndex = -1,
  tableStartCol = 0
}) {
  const { jsonData, merges, getMergedAwareCellText, getRawCellText, effectiveSheetCols } = readWorkbookSheetHelpers(
    workbook,
    sheetName
  );
  const rebuild = rebuildFormXVIIIMPCombinedRegisterTableHeadersFromSheet({
    headerRowIndex,
    getMergedAwareCellText,
    getRawCellText,
    jsonData,
    effectiveSheetCols,
    forceCombinedRegister: true,
    merges
  });
  if (rebuild?.headers?.length >= 4) {
    return {
      headers: repairFormXVIIIMPDuplicateWageHeaders(rebuild.headers),
      excelCols: rebuild.excelCols,
      headerRowIndex: rebuild.headerRowIndex,
      dataStartIndex: rebuild.dataStartIndex,
      tableStartCol: rebuild.tableStartCol
    };
  }
  const headers = repairFormXVIIIMPDuplicateWageHeaders(
    Array.isArray(fallbackHeaders) ? fallbackHeaders.filter((h) => String(h || '').trim()) : []
  );
  if (headers.length < 4) {
    return { headers: [], excelCols: [], headerRowIndex, dataStartIndex, tableStartCol };
  }
  const startCol = Math.max(0, Number(tableStartCol) || 0);
  const excelCols =
    Array.isArray(fallbackExcelCols) && fallbackExcelCols.length === headers.length
      ? fallbackExcelCols
      : headers.map((_, idx) => startCol + idx);
  return {
    headers,
    excelCols,
    headerRowIndex,
    dataStartIndex,
    tableStartCol: startCol
  };
}

export function repairFormXVIIIMPCombinedRegisterFromWorkbook(workbook, sheetName, parsed) {
  if (!workbook || !parsed) return parsed;
  const cols = resolveMPCombinedRegisterVisibleTableColumns({
    workbook,
    sheetName,
    headerRowIndex: parsed.headerRowIndex ?? -1,
    fallbackHeaders: parsed.headers,
    fallbackExcelCols: parsed.headerExcelCols,
    dataStartIndex: parsed.dataStartIndex ?? -1,
    tableStartCol: parsed.tableStartCol ?? 0
  });
  if (cols.headers.length < 4) return parsed;
  const priorHeaders = Array.isArray(parsed.headers) ? parsed.headers : [];
  parsed.headers = cols.headers;
  parsed.headerExcelCols = cols.excelCols;
  parsed.tableStartCol = cols.tableStartCol;
  parsed.headerRowIndex = cols.headerRowIndex;
  parsed.dataStartIndex = cols.dataStartIndex;
  parsed.originalHeaderRowIndex = cols.headerRowIndex;
  if (Array.isArray(parsed.tableData) && parsed.tableData.length > 0 && priorHeaders.length > 0) {
    parsed.tableData = remapMPCombinedRegisterRows(parsed.tableData, priorHeaders, cols.headers);
  }
  return parsed;
}

export function remapMPCombinedRegisterRows(rows, oldHeaders, newHeaders) {
  if (!Array.isArray(rows) || !Array.isArray(oldHeaders) || !Array.isArray(newHeaders)) return rows;
  const norm = (s) =>
    String(s || '')
      .replace(/\s+/g, ' ')
      .trim()
      .toLowerCase()
      .replace(/[^a-z0-9 ]/g, '');
  const samePayrollRole = (a, b) => {
    if (isFormXVIIIMPOtherAllowanceHeader(a) && isFormXVIIIMPOtherAllowanceHeader(b)) return true;
    if (isFormXVIIIMPWageRateHeader(a) && isFormXVIIIMPWageRateHeader(b)) return true;
    if (isFormXVIIIMPGrossWagesHeader(a) && isFormXVIIIMPGrossWagesHeader(b)) return true;
    if (isFormXVIIIMPNetPayableHeader(a) && isFormXVIIIMPNetPayableHeader(b)) return true;
    // PF (22) ↔ merged "Other Deductions Like EPF/…" parent key from the modal grid.
    if (isFormXVIIIMPPfHeader(a) && isFormXVIIIMPPfHeader(b)) return true;
    if (isFormXVIIIMPPtHeader(a) && isFormXVIIIMPPtHeader(b)) return true;
    if (isFormXVIIIMPEsicHeader(a) && isFormXVIIIMPEsicHeader(b)) return true;
    if (isFormXVIIIMPLwfHeader(a) && isFormXVIIIMPLwfHeader(b)) return true;
    return false;
  };
  const repairedNew = repairFormXVIIIMPDuplicateWageHeaders(newHeaders);
  // Keep caller header list aligned with remapped row keys.
  if (Array.isArray(newHeaders) && repairedNew.length === newHeaders.length) {
    for (let i = 0; i < repairedNew.length; i += 1) newHeaders[i] = repairedNew[i];
  }
  return rows.map((row) => {
    if (!row || typeof row !== 'object' || Array.isArray(row)) return row;
    const out = {};
    const otherSrcKey = Object.keys(row).find((k) => isFormXVIIIMPOtherAllowanceHeader(k));
    const pfSrcKey = Object.keys(row).find((k) => isFormXVIIIMPPfHeader(k));
    const ptSrcKey = Object.keys(row).find((k) => isFormXVIIIMPPtHeader(k));
    repairedNew.forEach((nh) => {
      const target = norm(nh);
      if (isFormXVIIIMPOtherAllowanceHeader(nh) && otherSrcKey) {
        out[nh] = row[otherSrcKey];
        return;
      }
      if (isFormXVIIIMPPfHeader(nh) && pfSrcKey) {
        out[nh] = row[pfSrcKey];
        return;
      }
      if (isFormXVIIIMPPtHeader(nh) && ptSrcKey) {
        out[nh] = row[ptSrcKey];
        return;
      }
      if (Object.prototype.hasOwnProperty.call(row, nh)) {
        out[nh] = row[nh];
        return;
      }
      const key =
        Object.keys(row).find((k) => norm(k) === target) ||
        Object.keys(row).find((k) => samePayrollRole(k, nh));
      if (key) out[nh] = row[key];
      else out[nh] = '';
    });
    return out;
  });
}

/** Read Other allowances from a grid row by role (not fuzzy wage-rate match). */
export function readFormXVIIIMPOtherAllowanceFromRow(row, headers = []) {
  if (!row || typeof row !== 'object' || Array.isArray(row)) return '';
  const list = Array.isArray(headers) ? headers : [];
  for (let i = 0; i < list.length; i += 1) {
    const h = list[i];
    if (!isFormXVIIIMPOtherAllowanceHeader(h)) continue;
    if (row[h] !== '' && row[h] != null) return row[h];
  }
  const keys = Object.keys(row).filter((k) => !String(k || '').startsWith('__'));
  for (let i = 0; i < keys.length; i += 1) {
    const k = keys[i];
    if (!isFormXVIIIMPOtherAllowanceHeader(k)) continue;
    if (row[k] !== '' && row[k] != null) return row[k];
  }
  return '';
}

/** Read wage-rate / gross amounts from a grid row by header role. */
export function readFormXVIIIMPWageAndGrossFromRow(row, headers = []) {
  if (!row || typeof row !== 'object' || Array.isArray(row)) {
    return { wage: '', gross: '', wageNum: NaN, grossNum: NaN };
  }
  const list = Array.isArray(headers) ? headers : [];
  const mp = resolveFormXVIIIMPTableHeaders(list);
  let wage = mp.wageRate && row[mp.wageRate] != null ? row[mp.wageRate] : '';
  let gross = mp.grossWages && row[mp.grossWages] != null ? row[mp.grossWages] : '';
  if (wage === '' || wage == null) {
    const k = Object.keys(row).find((key) => isFormXVIIIMPWageRateHeader(key));
    if (k) wage = row[k];
  }
  if (gross === '' || gross == null) {
    const k = Object.keys(row).find((key) => isFormXVIIIMPGrossWagesHeader(key));
    if (k) gross = row[k];
  }
  return {
    wage,
    gross,
    wageNum: parseFormXVIIIMPMoney(wage),
    grossNum: parseFormXVIIIMPMoney(gross),
  };
}

/**
 * Resolve Other allowances for Excel write: payroll formula first, then distinct grid value.
 * Never returns Wage rate / gross.
 */
export function resolveFormXVIIIMPOtherAllowanceForWrite(row, headers = [], payrollRow = null, helpers = {}) {
  const { wageNum, grossNum } = readFormXVIIIMPWageAndGrossFromRow(row, headers);
  const mirrorsWageOrGross = (val) => {
    const n = parseFormXVIIIMPMoney(val);
    if (!Number.isFinite(n) || n === 0) return !Number.isFinite(n);
    if (Number.isFinite(wageNum) && n === wageNum) return true;
    if (Number.isFinite(grossNum) && n === grossNum) return true;
    return false;
  };
  if (payrollRow && !payrollRow.fetch_error) {
    const parsed = readFormXVIIIMPPayrollGrossNet(payrollRow, helpers);
    if (parsed.other_allowance !== '' && parsed.other_allowance != null && !mirrorsWageOrGross(parsed.other_allowance)) {
      return parsed.other_allowance;
    }
  }
  const fromRow = readFormXVIIIMPOtherAllowanceFromRow(row, headers);
  if (fromRow !== '' && fromRow != null && !mirrorsWageOrGross(fromRow)) {
    return fromRow;
  }
  return '';
}

/** Snapshot Other allowances that are already distinct from Wage rate / gross (live UI values). */
export function snapshotFormXVIIIMPDistinctOtherAllowances(mappedData, headers) {
  if (!Array.isArray(mappedData)) return [];
  return mappedData.map((row) => {
    const fromRow = readFormXVIIIMPOtherAllowanceFromRow(row, headers);
    if (fromRow === '' || fromRow == null) return null;
    const { wageNum, grossNum } = readFormXVIIIMPWageAndGrossFromRow(row, headers);
    const n = parseFormXVIIIMPMoney(fromRow);
    if (!Number.isFinite(n) || n === 0) return null;
    if (Number.isFinite(wageNum) && n === wageNum) return null;
    if (Number.isFinite(grossNum) && n === grossNum) return null;
    return String(fromRow).trim();
  });
}

/** Restore snapshotted Other allowances when enrich/finalize left Wage rate mirrored or blank. */
export function restoreFormXVIIIMPDistinctOtherAllowances(mappedData, headers, snapshot) {
  if (!Array.isArray(mappedData) || !Array.isArray(snapshot)) return 0;
  const headerList = Array.isArray(headers) ? headers : [];
  {
    const repaired = repairFormXVIIIMPDuplicateWageHeaders(headerList);
    for (let i = 0; i < repaired.length; i += 1) headerList[i] = repaired[i];
  }
  const mp = resolveFormXVIIIMPTableHeaders(headerList);
  let otherHeader = mp.otherAllowances || headerList.find((h) => isFormXVIIIMPOtherAllowanceHeader(h));
  if (!otherHeader) {
    otherHeader = FORM_XVIII_MP_SECTION_INDEX_HEADERS[14];
    if (!headerList.includes(otherHeader)) headerList.push(otherHeader);
  }
  let hits = 0;
  mappedData.forEach((row, idx) => {
    if (!row || typeof row !== 'object') return;
    const preserved = snapshot[idx];
    if (preserved == null || preserved === '') return;
    const cur = parseFormXVIIIMPMoney(row[otherHeader]);
    const { wageNum, grossNum } = readFormXVIIIMPWageAndGrossFromRow(row, headerList);
    const bad =
      !Number.isFinite(cur) ||
      cur === 0 ||
      (Number.isFinite(wageNum) && cur === wageNum) ||
      (Number.isFinite(grossNum) && cur === grossNum);
    if (!bad) return;
    row[otherHeader] = preserved;
    headerList.forEach((h) => {
      if (isFormXVIIIMPOtherAllowanceHeader(h)) row[h] = preserved;
    });
    hits += 1;
  });
  return hits;
}

/** Header index for Other allowances (by role, or column immediately after Wage rate). */
export function findFormXVIIIMPOtherAllowanceHeaderIndex(headers) {
  const list = Array.isArray(headers) ? headers : [];
  const byRole = list.findIndex((h) => isFormXVIIIMPOtherAllowanceHeader(h));
  if (byRole >= 0) return byRole;
  const wageIdx = list.findIndex((h) => isFormXVIIIMPWageRateHeader(h));
  if (wageIdx >= 0 && wageIdx + 1 < list.length) return wageIdx + 1;
  return -1;
}

function unmergeFormXVIIIExcelJSRowsInRange(worksheet, startRow, endRow, colFrom, colTo) {
  const merges = worksheet?.model?.merges;
  if (!worksheet || !Array.isArray(merges) || merges.length === 0) return;
  const toRemove = [];
  for (let i = 0; i < merges.length; i += 1) {
    const range = merges[i];
    const parts = String(range || '').split(':');
    if (parts.length !== 2) continue;
    const start = parts[0].match(/^([A-Z]+)(\d+)$/i);
    const end = parts[1].match(/^([A-Z]+)(\d+)$/i);
    if (!start || !end) continue;
    const r1 = parseInt(start[2], 10);
    const r2 = parseInt(end[2], 10);
    const c1 = XLSX.utils.decode_col(start[1].toUpperCase()) + 1;
    const c2 = XLSX.utils.decode_col(end[1].toUpperCase()) + 1;
    if (r2 >= startRow && r1 <= endRow && c2 >= colFrom && c1 <= colTo) {
      toRemove.push(range);
    }
  }
  toRemove.forEach((range) => {
    try {
      worksheet.unMergeCells(range);
    } catch (_) {
      /* ignore */
    }
  });
}

export async function buildFormXVIIIWorkbookWithTemplateStyles({
  templateArrayBuffer,
  mappedData,
  mappedRowMatrix,
  headersToUse,
  headerExcelCols,
  parsedHeaderRowIndex,
  parsedDataStartIndex,
  parsedFormHeader,
  headerFormData = {},
  formFileName,
  currentItem = null,
  payrollRows = null,
  payrollHelpers = {},
  employees = [],
}) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(templateArrayBuffer);
  const worksheet = workbook.worksheets[0];
  if (!worksheet) throw new Error('Template worksheet not found.');

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

  const mergeTopLeftCache = new Map();
  const getMergeTopLeft = (r, c) => {
    const key = `${r}:${c}`;
    if (mergeTopLeftCache.has(key)) return mergeTopLeftCache.get(key);
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
    mergeTopLeftCache.set(key, topLeft);
    return topLeft;
  };
  const getMergedAwareCellText = (r, c) => {
    const tl = getMergeTopLeft(r, c);
    return excelCellValueToString(worksheet.getCell(tl.r, tl.c)?.value).trim();
  };

  const trimmedHeaders = Array.isArray(headersToUse)
    ? headersToUse.filter((h) => String(h || '').trim())
    : [];
  const layoutKnown =
    trimmedHeaders.length >= 4 &&
    Array.isArray(headerExcelCols) &&
    headerExcelCols.length === trimmedHeaders.length &&
    Number(parsedHeaderRowIndex) >= 0 &&
    Number(parsedDataStartIndex) >= 0;

  let mpLayout;
  if (layoutKnown) {
    mpLayout = {
      headers: trimmedHeaders,
      excelCols: headerExcelCols,
      headerRowIndex: parsedHeaderRowIndex,
      dataStartIndex: parsedDataStartIndex,
      tableStartCol: Math.max(0, Number(headerExcelCols[0]) || 0)
    };
  } else {
    const maxScanRows = Math.max(
      (Number(parsedDataStartIndex) >= 0 ? Number(parsedDataStartIndex) + 25 : 0),
      (Number(parsedHeaderRowIndex) >= 0 ? Number(parsedHeaderRowIndex) + 15 : 0),
      45,
      worksheet.rowCount + 5
    );
    const maxScanCols = Math.max(
      Array.isArray(headerExcelCols) && headerExcelCols.length > 0
        ? Math.max(...headerExcelCols.map((c) => Number(c) || 0)) + 3
        : 0,
      45,
      worksheet.columnCount + 3
    );
    const jsonData = [];
    for (let r = 1; r <= maxScanRows; r += 1) {
      const row = [];
      for (let c = 1; c <= maxScanCols; c += 1) {
        row.push(getMergedAwareCellText(r, c));
      }
      jsonData.push(row);
    }
    const getMergedAwareCellText0 = (r0, c0) => getMergedAwareCellText(r0 + 1, c0 + 1);

    mpLayout = rebuildFormXVIIIMPCombinedRegisterTableHeadersFromSheet({
      headerRowIndex: parsedHeaderRowIndex ?? -1,
      getMergedAwareCellText: getMergedAwareCellText0,
      jsonData,
      effectiveSheetCols: maxScanCols,
      forceCombinedRegister: true
    });
    if (!mpLayout?.headers?.length) {
      mpLayout = {
        headers: trimmedHeaders,
        excelCols: Array.isArray(headerExcelCols) ? headerExcelCols : null,
        headerRowIndex: parsedHeaderRowIndex ?? -1,
        dataStartIndex: parsedDataStartIndex ?? -1,
        tableStartCol: 0
      };
    }
  }

  const effectiveHeaders = Array.isArray(mpLayout.headers) ? [...mpLayout.headers] : [];
  if (effectiveHeaders.length < 4) {
    throw new Error('Could not locate Form XVIII MP Combined Register table columns.');
  }

  let fieldCols = [];
  if (Array.isArray(mpLayout.excelCols) && mpLayout.excelCols.length === effectiveHeaders.length) {
    fieldCols = mpLayout.excelCols.map((c) => Number(c) + 1);
  } else if (Array.isArray(headerExcelCols) && headerExcelCols.length === effectiveHeaders.length) {
    fieldCols = headerExcelCols.map((c) => Number(c) + 1);
  } else {
    const startCol = (Number(mpLayout.tableStartCol) || 0) + 1;
    fieldCols = effectiveHeaders.map((_, idx) => startCol + idx);
  }

  const startRow =
    mpLayout.dataStartIndex != null && mpLayout.dataStartIndex >= 0
      ? mpLayout.dataStartIndex + 1
      : parsedDataStartIndex != null && parsedDataStartIndex >= 0
        ? parsedDataStartIndex + 1
        : mpLayout.headerRowIndex >= 0
          ? mpLayout.headerRowIndex + 3
          : 12;

  // Repair headers from the section-index row (13=Wage rate, 14=Other allowances, …).
  // Prevents layoutKnown / merged-parent labels from writing Wage rate into col 14.
  const sectionRow1 =
    startRow > 1 ? startRow - 1 : mpLayout.headerRowIndex >= 0 ? mpLayout.headerRowIndex + 2 : -1;
  if (sectionRow1 > 0 && fieldCols.length === effectiveHeaders.length) {
    const repaired = effectiveHeaders.map((label, idx) => {
      const col = fieldCols[idx];
      // Prefer the cell's own value; only fall back to merge parent when blank.
      const rawSection = excelCellValueToString(worksheet.getCell(sectionRow1, col)?.value).trim();
      const mergedSection = getMergedAwareCellText(sectionRow1, col);
      const sectionRaw = rawSection || mergedSection;
      const sectionIdx = String(sectionRaw).replace(/[^\d]/g, '');
      const prev = idx > 0 ? effectiveHeaders[idx - 1] : '';
      return resolveFormXVIIIMPSectionHeaderLabel(label, sectionIdx, prev) || label;
    });
    for (let i = 0; i < repaired.length; i += 1) effectiveHeaders[i] = repaired[i];
  }
  {
    const deduped = repairFormXVIIIMPDuplicateWageHeaders(effectiveHeaders);
    for (let i = 0; i < deduped.length; i += 1) effectiveHeaders[i] = deduped[i];
  }

  // Align row keys with repaired headers (Other Deductions… → PF (22), etc.).
  const sourceMapped =
    Array.isArray(mappedData) && mappedData.length > 0
      ? remapMPCombinedRegisterRows(
          mappedData,
          trimmedHeaders.length > 0 ? trimmedHeaders : effectiveHeaders,
          effectiveHeaders
        )
      : mappedData;

  writeStatutoryHeaderFieldsToExcelJsWorksheet(worksheet, {
    headerFormData: headerFormData && typeof headerFormData === 'object' ? headerFormData : {},
    parsedFormHeader,
    headerRowEnd: Math.max(1, startRow - 1),
    maxScanCols: 80
  });

  const normalize = (txt) =>
    String(txt || '')
      .replace(/\r?\n/g, ' ')
      .replace(/\s+/g, ' ')
      .trim()
      .toLowerCase()
      .replace(/[^a-z0-9 ]/g, '');

  const getRowValueForHeader = (row, header, headerIndex, allHeaders) => {
    if (!row || typeof row !== 'object' || Array.isArray(row)) {
      return Array.isArray(row) ? row[headerIndex] : '';
    }
    if (Object.prototype.hasOwnProperty.call(row, header)) {
      const direct = row[header];
      // Empty Other allowances must not fuzzy-fall through to Wage rate.
      if (direct !== '' && direct != null) return direct;
      if (!isFormXVIIIMPOtherAllowanceHeader(header)) return direct;
    }
    const target = normalize(header);
    if (!target) return '';
    const rowKeys = Object.keys(row).filter((k) => !String(k || '').startsWith('__'));
    const exact = rowKeys.filter((k) => normalize(k) === target);
    if (exact.length === 1) return row[exact[0]];
    if (exact.length > 1) {
      const occur =
        allHeaders.slice(0, headerIndex + 1).filter((h) => normalize(h) === target).length - 1;
      return row[exact[Math.min(Math.max(occur, 0), exact.length - 1)]];
    }
    // Role-based match (Other allowances ↔ Other allowance) — never Wage rate ↔ Other allowances.
    if (isFormXVIIIMPOtherAllowanceHeader(header)) {
      const otherKey = rowKeys.find((k) => isFormXVIIIMPOtherAllowanceHeader(k));
      if (otherKey) return row[otherKey];
      return '';
    }
    if (isFormXVIIIMPWageRateHeader(header)) {
      const wageKey = rowKeys.find((k) => isFormXVIIIMPWageRateHeader(k));
      if (wageKey) return row[wageKey];
      return '';
    }
    // PF (22) ↔ merged "Other Deductions Like EPF/…" (modal key before download rebuild).
    if (isFormXVIIIMPPfHeader(header)) {
      const pfKey = rowKeys.find((k) => isFormXVIIIMPPfHeader(k));
      if (pfKey && row[pfKey] !== '' && row[pfKey] != null) return row[pfKey];
      return Object.prototype.hasOwnProperty.call(row, header) ? row[header] : '';
    }
    if (isFormXVIIIMPPtHeader(header)) {
      const ptKey = rowKeys.find((k) => isFormXVIIIMPPtHeader(k));
      if (ptKey && row[ptKey] !== '' && row[ptKey] != null) return row[ptKey];
      return Object.prototype.hasOwnProperty.call(row, header) ? row[header] : '';
    }
    if (isFormXVIIIMPEsicHeader(header)) {
      const esicKey = rowKeys.find((k) => isFormXVIIIMPEsicHeader(k));
      if (esicKey && row[esicKey] !== '' && row[esicKey] != null) return row[esicKey];
      return Object.prototype.hasOwnProperty.call(row, header) ? row[header] : '';
    }
    if (isFormXVIIIMPLwfHeader(header)) {
      const lwfKey = rowKeys.find((k) => isFormXVIIIMPLwfHeader(k));
      if (lwfKey && row[lwfKey] !== '' && row[lwfKey] != null) return row[lwfKey];
      return Object.prototype.hasOwnProperty.call(row, header) ? row[header] : '';
    }
    const fuzzy = rowKeys.find((k) => {
      const nk = normalize(k);
      if (!nk) return false;
      if (isFormXVIIIMPWageRateHeader(k) || isFormXVIIIMPOtherAllowanceHeader(k)) return false;
      if (isFormXVIIIMPPfHeader(k) || isFormXVIIIMPPtHeader(k)) return false;
      return nk.includes(target) || target.includes(nk);
    });
    return fuzzy ? row[fuzzy] : '';
  };

  const rowLooksMeaningful = (row) => {
    if (Array.isArray(row)) {
      const cells = row.map((v) => String(v ?? '').trim()).filter(Boolean);
      return cells.length > 0 && cells.some((v) => /[a-z]/i.test(v));
    }
    if (row && typeof row === 'object') {
      const vals = Object.values(row).map((v) => String(v ?? '').trim()).filter(Boolean);
      return vals.length > 0 && vals.some((v) => /[a-z]/i.test(v));
    }
    return false;
  };

  const sourcePrimary =
    Array.isArray(sourceMapped) && sourceMapped.length > 0
      ? sourceMapped
      : Array.isArray(mappedRowMatrix)
        ? mappedRowMatrix
        : [];

  const tableColMin = fieldCols.length > 0 ? Math.min(...fieldCols) : 1;
  const tableColMax = fieldCols.length > 0 ? Math.max(...fieldCols) : effectiveHeaders.length;
  const templateBodyRows = countExcelJSTemplateBodyRows(worksheet, startRow, tableColMin, tableColMax);
  const bodyRowsToPaint = Math.max(sourcePrimary.length, templateBodyRows);

  clearExcelJSTrailingTableCells(worksheet, {
    dataStartRow: startRow,
    dataRowCount: bodyRowsToPaint,
    afterCol: tableColMax,
    throughCol: tableColMax + 30
  });

  const mpWriteHeaders = resolveFormXVIIIMPTableHeaders(effectiveHeaders);
  let otherHeaderIdx = findFormXVIIIMPOtherAllowanceHeaderIndex(effectiveHeaders);
  if (otherHeaderIdx >= 0 && !isFormXVIIIMPOtherAllowanceHeader(effectiveHeaders[otherHeaderIdx])) {
    effectiveHeaders[otherHeaderIdx] = FORM_XVIII_MP_SECTION_INDEX_HEADERS[14];
  }
  const otherHdrForWrite =
    mpWriteHeaders.otherAllowances ||
    (otherHeaderIdx >= 0 ? effectiveHeaders[otherHeaderIdx] : null) ||
    FORM_XVIII_MP_SECTION_INDEX_HEADERS[14];
  const otherExcelCol = otherHeaderIdx >= 0 ? fieldCols[otherHeaderIdx] : null;

  const writeHelpers = {
    flattenPayrollEarningColumns,
    ...payrollHelpers,
  };
  const payrollList = Array.isArray(payrollRows) ? payrollRows : [];
  const payrollResolver =
    payrollList.length > 0 ? buildFormXVIIIMPPayrollRowResolver(payrollList) : null;

  // Unmerge Wage rate ↔ Other allowances spans so col 14 can hold its own value.
  if (otherExcelCol && sourcePrimary.length > 0) {
    const wageIdx = effectiveHeaders.findIndex((h) => isFormXVIIIMPWageRateHeader(h));
    const wageCol = wageIdx >= 0 ? fieldCols[wageIdx] : otherExcelCol - 1;
    const colFrom = Math.max(1, Math.min(wageCol || otherExcelCol, otherExcelCol) - 1);
    const colTo = Math.max(wageCol || otherExcelCol, otherExcelCol) + 1;
    unmergeFormXVIIIExcelJSRowsInRange(
      worksheet,
      startRow,
      startRow + Math.max(sourcePrimary.length, 1) - 1,
      colFrom,
      colTo
    );
  }

  // Unmerge Other Deductions band (PF/ESIC/PT/LWF) so each child column can hold its own value.
  {
    const dedCols = effectiveHeaders
      .map((h, idx) =>
        isFormXVIIIMPPfHeader(h) ||
        isFormXVIIIMPEsicHeader(h) ||
        isFormXVIIIMPPtHeader(h) ||
        isFormXVIIIMPLwfHeader(h)
          ? fieldCols[idx]
          : null
      )
      .filter((c) => Number.isFinite(c) && c > 0);
    if (dedCols.length > 0 && sourcePrimary.length > 0) {
      unmergeFormXVIIIExcelJSRowsInRange(
        worksheet,
        startRow,
        startRow + Math.max(sourcePrimary.length, 1) - 1,
        Math.min(...dedCols),
        Math.max(...dedCols)
      );
    }
  }

  const writeCellValue = (excelRow, targetCol, value, useExactCell) => {
    if (!targetCol || targetCol < 1) return;
    const cell = useExactCell
      ? worksheet.getCell(excelRow, targetCol)
      : (() => {
          const tl = getMergeTopLeft(excelRow, targetCol);
          return worksheet.getCell(tl.r, tl.c);
        })();
    if (value == null || value === '') {
      cell.value = '';
      return;
    }
    if (
      typeof value === 'number' ||
      (typeof value === 'string' && /^-?\d+(\.\d+)?$/.test(String(value).trim()))
    ) {
      cell.value = Number(value);
    } else {
      cell.value = String(value);
    }
    cell.font = { ...(cell.font || {}), bold: false };
  };

  for (let i = 0; i < sourcePrimary.length; i += 1) {
    if (i > 0 && i % 20 === 0) {
      await yieldToMain();
    }
    const row = sourcePrimary[i];
    if (!rowLooksMeaningful(row)) continue;
    const excelRow = startRow + i;
    const { wageNum, grossNum } = Array.isArray(row)
      ? { wageNum: NaN, grossNum: NaN }
      : readFormXVIIIMPWageAndGrossFromRow(row, effectiveHeaders);

    let payrollRow = null;
    if (!Array.isArray(row) && payrollList.length > 0) {
      const empItem = Array.isArray(employees) ? employees[i] : null;
      const emp = empItem?.Employee || empItem?.employee || empItem;
      payrollRow = resolveFormXVIIIMPPayrollRowForAutofillRow(
        emp,
        row,
        payrollList,
        effectiveHeaders,
        payrollResolver
      );
      if ((!payrollRow || payrollRow.fetch_error) && payrollList[i] && !payrollList[i].fetch_error) {
        payrollRow = payrollList[i];
      }
    }
    const otherWriteVal = Array.isArray(row)
      ? ''
      : resolveFormXVIIIMPOtherAllowanceForWrite(row, effectiveHeaders, payrollRow, writeHelpers);
    if (!Array.isArray(row) && otherWriteVal !== '' && otherHdrForWrite) {
      row[otherHdrForWrite] = otherWriteVal;
    }
    // Stamp Sample Payroll PF / PT onto the row before cell write (download safety net).
    if (!Array.isArray(row) && payrollRow && !payrollRow.fetch_error) {
      applyFormXVIIIMPPayrollToRow(row, payrollRow, mpWriteHeaders, {
        ...writeHelpers,
        headers: effectiveHeaders,
        sanitizeValue: writeHelpers.sanitizeValue || ((v) => v),
      });
    }

    for (let j = 0; j < effectiveHeaders.length; j += 1) {
      const header = effectiveHeaders[j];
      let value = Array.isArray(row) ? row[j] : getRowValueForHeader(row, header, j, effectiveHeaders);
      // Hard guard: never paint Wage rate / gross into Other allowances column.
      if (isFormXVIIIMPOtherAllowanceHeader(header) || j === otherHeaderIdx) {
        value = otherWriteVal;
        const otherNum = parseFormXVIIIMPMoney(value);
        if (
          Number.isFinite(otherNum) &&
          ((Number.isFinite(wageNum) && otherNum === wageNum && wageNum !== 0) ||
            (Number.isFinite(grossNum) && otherNum === grossNum && grossNum !== 0))
        ) {
          value = '';
        }
      }
      const targetCol = fieldCols[j];
      const useExactCell =
        isFormXVIIIMPOtherAllowanceHeader(header) ||
        j === otherHeaderIdx ||
        isFormXVIIIMPWageRateHeader(header) ||
        isFormXVIIIMPGrossWagesHeader(header) ||
        isFormXVIIIMPNetPayableHeader(header) ||
        isFormXVIIIMPPfHeader(header) ||
        isFormXVIIIMPPtHeader(header) ||
        isFormXVIIIMPEsicHeader(header) ||
        isFormXVIIIMPLwfHeader(header);
      writeCellValue(excelRow, targetCol, value, useExactCell);
    }

    // Final pass: always stamp Other allowances onto the exact sheet column (section 14).
    if (otherExcelCol) {
      writeCellValue(excelRow, otherExcelCol, otherWriteVal, true);
    }
  }

  if (sourcePrimary.length > 0 && fieldCols.length > 0) {
    ensureExcelJSDataRowsWithBorders(worksheet, {
      dataStartRow: startRow,
      dataRowCount: sourcePrimary.length,
      colFrom: tableColMin,
      colTo: tableColMax,
      templateRow: startRow,
      templateBodyRows: 1
    });
    clearExcelJSTrailingTableCells(worksheet, {
      dataStartRow: startRow,
      dataRowCount: bodyRowsToPaint,
      afterCol: tableColMax,
      throughCol: tableColMax + 30
    });
  }

  await yieldToMain();
  const out = await workbook.xlsx.writeBuffer();
  const outName =
    formFileName ||
    currentItem?.formName?.replace(/[^a-zA-Z0-9]/g, '_') ||
    parsedFormHeader?.title?.replace(/[^a-zA-Z0-9]/g, '_') ||
    `Form_XVIII_MP_${Date.now()}.xlsx`;
  const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  return { blob, fileName: outName };
}
