/** Form XVIII — Madhya Pradesh MP Combined Register (Muster Roll-cum Register of Wages). */

import ExcelJS from 'exceljs';
import * as XLSX from 'xlsx';
import {
  clearExcelJSTrailingTableCells,
  countExcelJSTemplateBodyRows,
  ensureExcelJSDataRowsWithBorders
} from '../../utils/excelTableBorders';
import { flattenPayrollEarningColumns } from '../../utils/payrollEarnings';
import { writeStatutoryHeaderFieldsToExcelJsWorksheet } from '../../utils/statutorySiteCompanyHeaders';

export function mpCombinedRegisterHeaderNorm(txt) {
  return String(txt || '')
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

export function looksLikeMPCombinedRegisterSheet(blob) {
  const s = String(blob || '').toLowerCase();
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
    if (rawMain && !isEffectivelyBlankHeaderCell(rawMain) && !/^\d+$/.test(rawMain)) {
      return rawMain.replace(/\s+/g, ' ').trim();
    }
    for (let r = Math.max(0, mainRow - 2); r <= bandEnd; r += 1) {
      const leaf = String(readRaw(r, c) || '').trim();
      if (leaf && !isEffectivelyBlankHeaderCell(leaf) && !/^\d+$/.test(leaf)) {
        return leaf.replace(/\s+/g, ' ').trim();
      }
    }
    const merged = String(getMergedAwareCellText(mainRow, c) || '').trim();
    if (merged && !isEffectivelyBlankHeaderCell(merged) && !/^\d+$/.test(merged)) {
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
    const sectionIdx =
      sectionRow >= 0 ? normalizeStatutoryColumnIndexCell(getMergedAwareCellText(sectionRow, c)) : '';
    let label = pickColumnHeaderLabel(c, sectionIdx);
    if (!label) {
      if (headers.length === 0) continue;
      label =
        sectionIdx && /^\d+$/.test(sectionIdx)
          ? `Column ${sectionIdx}`
          : `Column ${c - startCol + 1}`;
    } else if (sectionIdx && /^\d+$/.test(sectionIdx) && !/\(\s*\d+\s*\)/.test(label)) {
      const prevBase =
        headers.length > 0 ? stripTrailingSectionIndex(headers[headers.length - 1]) : '';
      const curBase = stripTrailingSectionIndex(label);
      if (prevBase && prevBase === curBase) {
        label = `${label} (${sectionIdx})`;
      }
    }
    headers.push(label);
    excelCols.push(c);
  }

  if (headers.length < 4) return null;
  return {
    headers,
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

export function isFormXVIIIMPEducationSkillHeader(header) {
  const s = mpCombinedRegisterHeaderNorm(header).replace(/[^a-z0-9 ]/g, ' ').replace(/\s+/g, ' ').trim();
  return s.includes('education') && (s.includes('skill') || s.includes('skil'));
}

export function isFormXVIIIMPGrossWagesHeader(header) {
  const s = mpCombinedRegisterHeaderNorm(header).replace(/[^a-z0-9 ]/g, ' ').replace(/\s+/g, ' ').trim();
  if (s.includes('net') || (s.includes('wage') && s.includes('rate'))) return false;
  if (s.includes('gross') && (s.includes('wage') || s.includes('earning'))) return true;
  return s.includes('total') && s.includes('wage') && s.includes('earning');
}

export function isFormXVIIIMPNetPayableHeader(header) {
  const s = mpCombinedRegisterHeaderNorm(header).replace(/[^a-z0-9 ]/g, ' ').replace(/\s+/g, ' ').trim();
  return s.includes('net') && (s.includes('payable') || s.includes('paid') || s.includes('amount'));
}

/** PF/ESIC/PT/LWF deduction amount columns — not EPF/UAN registration columns. */
export function isFormXVIIIMPPayrollDeductionHeader(header) {
  const s = mpCombinedRegisterHeaderNorm(header).replace(/[^a-z0-9 ]/g, ' ').replace(/\s+/g, ' ').trim();
  return (
    s === 'pf' ||
    /^pf\s*\(/.test(s) ||
    s.startsWith('pf ') ||
    s === 'esic' ||
    /^esic\s*\(/.test(s) ||
    s.startsWith('esic ') ||
    s === 'pt' ||
    /^pt\s*\(/.test(s) ||
    s.startsWith('pt ') ||
    s === 'lwf' ||
    /^lwf\s*\(/.test(s) ||
    s.startsWith('lwf ')
  );
}

export function readFormXVIIIMPPayrollGrossNet(payrollPayload, helpers = {}) {
  const flatFn = helpers.flattenPayrollEarningColumns || flattenPayrollEarningColumns;
  const flat =
    payrollPayload && typeof flatFn === 'function' ? flatFn(payrollPayload) : payrollPayload || {};
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
  return { gross, net, flat, p };
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
      return s && testFn(s);
    }) || null;
  return {
    grossWages: findHeader(
      (s) => (s.includes('gross') || s.includes('total')) && (s.includes('wage') || s.includes('earning'))
    ),
    overtimeWages: findHeader((s) => s.includes('overtime') && s.includes('wage')),
    overtimeHours: findHeader((s) => s.includes('overtime') && s.includes('hour')),
    otherAllowances: findHeader((s) => s.includes('other') && s.includes('allowance')),
    pf: findHeader((s) => s === 'pf' || s.startsWith('pf ') || s.includes('provident')),
    esic: findHeader((s) => s.includes('esic') || (s.includes('employee') && s.includes('insurance'))),
    pt: findHeader((s) => s === 'pt' || s.startsWith('pt ') || s.includes('professional tax')),
    lwf: findHeader((s) => s.includes('lwf') || s.includes('labour welfare') || s.includes('labor welfare')),
    fines: findHeader((s) => s.includes('fine') && s.includes('deduction')),
    advances: findHeader((s) => s.includes('advance') || s.includes('loan')),
    netPayable: findHeader((s) => s.includes('net') && (s.includes('payable') || s.includes('paid'))),
    bankUtr: findHeader((s) => s.includes('bank') && s.includes('utr')),
    educationSkill: findHeader((s) => s.includes('education') && (s.includes('skill') || s.includes('skil')))
  };
}

export function applyFormXVIIIMPEmployeeToRow(row, emp, mpHeaders, helpers = {}) {
  if (!row || !emp || !mpHeaders) return false;
  const sanitizeValue = helpers.sanitizeValue || ((v) => v);
  let hit = false;
  if (mpHeaders.educationSkill) {
    const edu = formatEducationSkillFromEmployee(emp);
    if (edu) {
      row[mpHeaders.educationSkill] = sanitizeValue(edu);
      hit = true;
    }
  }
  return hit;
}

export function applyFormXVIIIMPPayrollToRow(row, payrollPayload, mpHeaders, helpers = {}) {
  if (!row || !payrollPayload) return false;
  const sanitizeValue = helpers.sanitizeValue || ((v) => v);
  const { gross, net, flat, p } = readFormXVIIIMPPayrollGrossNet(payrollPayload, helpers);
  const headers = Array.isArray(helpers.headers) ? helpers.headers : Object.keys(row || {});
  let hit = false;

  const writeAmount = (header, val) => {
    if (!header || val === '' || val == null) return;
    row[header] = sanitizeValue(val);
    hit = true;
  };

  headers.forEach((h) => {
    if (isFormXVIIIMPGrossWagesHeader(h)) writeAmount(h, gross);
    if (isFormXVIIIMPNetPayableHeader(h)) writeAmount(h, net);
  });

  if (mpHeaders) {
    writeAmount(mpHeaders.grossWages, gross);
    writeAmount(mpHeaders.netPayable, net);
  }

  const findEarningAmount = helpers.findEarningAmount;
  const findDeductionAmount = helpers.findDeductionAmount;
  if (typeof findEarningAmount !== 'function' || typeof findDeductionAmount !== 'function') {
    return hit;
  }

  const earnings = helpers.getEarningsArray ? helpers.getEarningsArray(p) : [];
  const deductions = helpers.getDeductionsArray ? helpers.getDeductionsArray(p) : [];
  const ot =
    findEarningAmount(earnings, (t, n) => t === 'overtime' || t === 'ot' || n.includes('overtime')) || '';
  const otherAllow =
    findEarningAmount(
      earnings,
      (t, n) => n.includes('other') && (n.includes('allowance') || n.includes('allowances'))
    ) || '';
  const pf =
    findDeductionAmount(
      deductions,
      (t, n) => t === 'pf' || t === 'epf' || n.includes('provident') || n.includes('epf')
    ) || '';
  const esic =
    findDeductionAmount(
      deductions,
      (t, n) => t === 'esi' || n.includes('esic') || n.includes('employee state insurance')
    ) || '';
  const pt =
    findDeductionAmount(
      deductions,
      (t, n) => t === 'pt' || n.includes('professional tax')
    ) || '';
  const lwf =
    findDeductionAmount(
      deductions,
      (t, n) => n.includes('labour welfare') || n.includes('labor welfare') || n.includes('lwf')
    ) || '';

  const normHdr = (h) =>
    mpCombinedRegisterHeaderNorm(h).replace(/[^a-z0-9 ]/g, ' ').replace(/\s+/g, ' ').trim();
  headers.forEach((h) => {
    const s = normHdr(h);
    if (s === 'pf' || /^pf\s*\(/.test(s) || s.startsWith('pf ')) writeAmount(h, pf);
    if (s === 'esic' || /^esic\s*\(/.test(s) || s.startsWith('esic ')) writeAmount(h, esic);
    if (s === 'pt' || /^pt\s*\(/.test(s) || s.startsWith('pt ')) writeAmount(h, pt);
    if (s === 'lwf' || /^lwf\s*\(/.test(s) || s.startsWith('lwf ')) writeAmount(h, lwf);
  });

  const set = (key, val) => {
    if (!mpHeaders?.[key] || val === '' || val == null) return;
    row[mpHeaders[key]] = sanitizeValue(val);
    hit = true;
  };

  set('grossWages', gross);
  set('overtimeWages', ot);
  set('otherAllowances', otherAllow);
  set('pf', pf);
  set('esic', esic);
  set('pt', pt);
  set('lwf', lwf);
  set('netPayable', net);
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
      headers: rebuild.headers,
      excelCols: rebuild.excelCols,
      headerRowIndex: rebuild.headerRowIndex,
      dataStartIndex: rebuild.dataStartIndex,
      tableStartCol: rebuild.tableStartCol
    };
  }
  const headers = Array.isArray(fallbackHeaders)
    ? fallbackHeaders.filter((h) => String(h || '').trim())
    : [];
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
  return rows.map((row) => {
    if (!row || typeof row !== 'object' || Array.isArray(row)) return row;
    const out = {};
    newHeaders.forEach((nh) => {
      const target = norm(nh);
      if (Object.prototype.hasOwnProperty.call(row, nh)) {
        out[nh] = row[nh];
        return;
      }
      const key = Object.keys(row).find((k) => norm(k) === target);
      if (key) out[nh] = row[key];
      else out[nh] = '';
    });
    return out;
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
  currentItem = null
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

  const maxScanRows = Math.max(120, worksheet.rowCount + 10);
  const maxScanCols = Math.max(60, worksheet.columnCount + 5);
  const jsonData = [];
  for (let r = 1; r <= maxScanRows; r += 1) {
    const row = [];
    for (let c = 1; c <= maxScanCols; c += 1) {
      row.push(getMergedAwareCellText(r, c));
    }
    jsonData.push(row);
  }
  const getMergedAwareCellText0 = (r0, c0) => getMergedAwareCellText(r0 + 1, c0 + 1);

  let mpLayout = rebuildFormXVIIIMPCombinedRegisterTableHeadersFromSheet({
    headerRowIndex: parsedHeaderRowIndex ?? -1,
    getMergedAwareCellText: getMergedAwareCellText0,
    jsonData,
    effectiveSheetCols: maxScanCols,
    forceCombinedRegister: true
  });
  if (!mpLayout?.headers?.length) {
    mpLayout = {
      headers: Array.isArray(headersToUse) ? headersToUse.filter((h) => String(h || '').trim()) : [],
      excelCols: Array.isArray(headerExcelCols) ? headerExcelCols : null,
      headerRowIndex: parsedHeaderRowIndex ?? -1,
      dataStartIndex: parsedDataStartIndex ?? -1,
      tableStartCol: 0
    };
  }

  const effectiveHeaders = mpLayout.headers;
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
    if (Object.prototype.hasOwnProperty.call(row, header)) return row[header];
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
    const fuzzy = rowKeys.find((k) => {
      const nk = normalize(k);
      return nk && (nk.includes(target) || target.includes(nk));
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
    Array.isArray(mappedData) && mappedData.length > 0
      ? mappedData
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

  for (let i = 0; i < sourcePrimary.length; i += 1) {
    const row = sourcePrimary[i];
    if (!rowLooksMeaningful(row)) continue;
    const excelRow = startRow + i;
    for (let wc = 0; wc < fieldCols.length; wc += 1) {
      worksheet.getCell(excelRow, fieldCols[wc]).value = '';
    }
    for (let j = 0; j < effectiveHeaders.length; j += 1) {
      const header = effectiveHeaders[j];
      let value = Array.isArray(row) ? row[j] : getRowValueForHeader(row, header, j, effectiveHeaders);
      if (value == null || value === '') continue;
      const targetCol = fieldCols[j];
      if (!targetCol || targetCol < 1) continue;
      const tl = getMergeTopLeft(excelRow, targetCol);
      const cell = worksheet.getCell(tl.r, tl.c);
      if (
        typeof value === 'number' ||
        (typeof value === 'string' && /^-?\d+(\.\d+)?$/.test(String(value).trim()))
      ) {
        cell.value = Number(value);
      } else {
        cell.value = String(value);
      }
      cell.font = { ...(cell.font || {}), bold: false };
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

  const out = await workbook.xlsx.writeBuffer();
  const outName =
    formFileName ||
    currentItem?.formName?.replace(/[^a-zA-Z0-9]/g, '_') ||
    parsedFormHeader?.title?.replace(/[^a-zA-Z0-9]/g, '_') ||
    `Form_XVIII_MP_${Date.now()}.xlsx`;
  const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  return { blob, fileName: outName };
}
