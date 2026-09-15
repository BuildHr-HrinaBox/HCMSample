import {
  flattenPayrollEarningColumns,
  readPayrollScalar,
} from '../../utils/payrollEarnings';
import {
  excelCellValueToString,
  formatStatutoryHeaderLabelValueExport,
  isContractorHeaderLabel,
  isEstablishmentContractCarriedHeaderLabel,
  isNatureLocationHeaderLabel,
  isPrincipalEmployerHeaderLabel,
  normalizeStatutoryHeaderLabel,
  resolveHeaderFieldExportValue,
} from '../../utils/statutorySiteCompanyHeaders';

/**
 * Form XVII Andhra Pradesh — Register of Wages.
 * PF / PT deduction columns ← Sample Payroll PF and Professional Tax.
 */

export const FORM_XVII_AP_DEFAULT_CONTRACTOR_NAME = 'VAYONA ENERGY PRIVATE LIMITED';

function formXVIIAPHeaderNorm(header) {
  return String(header || '')
    .replace(/\r?\n/g, ' ')
    .replace(/['''`´]/g, '')
    .replace(/[^a-z0-9]+/gi, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

function parseFormXVIIAPMoney(value) {
  if (value === '' || value == null) return NaN;
  const n = Number(String(value).replace(/[,₹]/g, '').trim());
  return Number.isFinite(n) ? n : NaN;
}

function pickFormXVIIAPAmount(...values) {
  for (let i = 0; i < values.length; i += 1) {
    const value = values[i];
    if (value === '' || value == null) continue;
    const n = parseFormXVIIAPMoney(value);
    if (Number.isFinite(n)) return n;
  }
  return '';
}

function unwrapPayrollPayload(payrollRow) {
  if (!payrollRow || typeof payrollRow !== 'object') return payrollRow;
  if (typeof payrollRow.payroll_payload === 'string') {
    try {
      return { ...payrollRow, ...JSON.parse(payrollRow.payroll_payload) };
    } catch (_) {
      return payrollRow;
    }
  }
  if (payrollRow.payroll_payload && typeof payrollRow.payroll_payload === 'object') {
    return { ...payrollRow, ...payrollRow.payroll_payload };
  }
  return payrollRow;
}

/** PF amount column — not EPF/UAN registration number. */
export function isFormXVIIAPPfHeader(header) {
  const s = formXVIIAPHeaderNorm(header);
  if (!s) return false;
  // Deduped duplicate PF leaf (PF (2)) is the PT column on merged Form XVII templates.
  if (/^pf\s*\(\d+\)$/.test(s)) return false;
  if (s.includes('total') && s.includes('deduction')) return false;
  if (s.includes('uan') || s.includes('registration') || (s.includes('account') && s.includes('number'))) {
    return false;
  }
  if (/\bno\b/.test(s) && (/\bpf\b/.test(s) || s.includes('provident') || /\bepf\b/.test(s))) {
    return false;
  }
  if (s === 'pf' || /^pf\s*\d*$/.test(s) || s.startsWith('pf ')) return true;
  if (s === 'epf' || s.startsWith('epf ')) return true;
  if (s.includes('provident') && !s.includes('voluntary')) return true;
  return /\bpf\b/.test(s) && s.includes('deduction');
}

/** PT / Professional Tax amount column. */
export function isFormXVIIAPPtHeader(header) {
  const s = formXVIIAPHeaderNorm(header);
  if (!s) return false;
  if (s.includes('total') && s.includes('deduction')) return false;
  if (s === 'pt' || /^pt\s*\d*$/.test(s) || s.startsWith('pt ')) return true;
  if (s.includes('professional') && s.includes('tax')) return true;
  return /\bpt\b/.test(s) && s.includes('deduction');
}

/** Merged parent over ESI / PF / PT / SEVA / Total Deductions. */
export function isFormXVIIAPDeductionsParentBanner(header) {
  const s = formXVIIAPHeaderNorm(header);
  return s.includes('deduction') && s.includes('indicate') && s.includes('nature');
}

export function isFormXVIIAPTotalDeductionsHeader(header) {
  const s = formXVIIAPHeaderNorm(header);
  if (!s) return false;
  if (isFormXVIIAPDeductionsParentBanner(header)) return false;
  return (
    s === 'total deductions' ||
    s === 'total deduction' ||
    (s.includes('total') && s.includes('deduction') && !s.includes('other'))
  );
}

export function isFormXVIIAPNetAmountPaidHeader(header) {
  const s = formXVIIAPHeaderNorm(header);
  return s.includes('net') && s.includes('amount') && s.includes('paid');
}

function formXVIIAPHeaderRole(header) {
  if (isFormXVIIAPTotalDeductionsHeader(header)) return 'Total Deductions';
  if (isFormXVIIAPPfHeader(header)) return 'PF';
  if (isFormXVIIAPPtHeader(header)) return 'PT';
  const s = formXVIIAPHeaderNorm(header);
  if (s === 'esi' || s.startsWith('esi ')) return 'ESI';
  if (s === 'seva' || s.startsWith('seva ')) return 'SEVA';
  if (isFormXVIIAPNetAmountPaidHeader(header)) return 'Net Amount paid';
  return '';
}

function isFormXVIIAPPlaceholderExportHeader(header) {
  const s = String(header || '').trim();
  if (!s) return true;
  if (/^column\s+\d+$/i.test(s)) return true;
  return isFormXVIIAPDeductionsParentBanner(s);
}

/**
 * Excel leaf cells for PT / Total Deductions are often blank (merged parent).
 * Recover them from position: PT is immediately after PF; Total Deductions is
 * immediately before Net Amount paid.
 */
export function inferFormXVIIAPDeductionExportHeaders(headers) {
  const list = Array.isArray(headers) ? headers.map((h) => String(h ?? '')) : [];
  if (list.length === 0) return list;
  const roles = list.map((h) => formXVIIAPHeaderRole(h));
  const pfIdx = roles.indexOf('PF');
  const netIdx = roles.indexOf('Net Amount paid');
  if (pfIdx < 0 && netIdx < 0) return list;
  const out = [...list];
  const setRole = (idx, label) => {
    if (idx < 0 || idx >= out.length) return;
    out[idx] = label;
    roles[idx] = label;
  };
  const canRelabel = (idx, label) => {
    if (idx < 0 || idx >= out.length) return false;
    const cur = roles[idx];
    if (cur === label) return false;
    if (!cur || isFormXVIIAPPlaceholderExportHeader(out[idx])) return true;
    // Merged Form XVII headers can copy "PF" onto the PT leaf column.
    if (label === 'PT' && cur === 'PF') return true;
    return false;
  };
  if (pfIdx >= 0 && canRelabel(pfIdx + 1, 'PT')) {
    setRole(pfIdx + 1, 'PT');
  }
  let totalIdx = netIdx > 0 ? netIdx - 1 : -1;
  if (totalIdx >= 0 && roles[totalIdx] === 'SEVA' && pfIdx >= 0 && pfIdx + 3 < out.length) {
    totalIdx = pfIdx + 3;
  }
  if (totalIdx >= 0 && canRelabel(totalIdx, 'Total Deductions')) {
    setRole(totalIdx, 'Total Deductions');
  } else if (pfIdx >= 0 && pfIdx + 3 < out.length && canRelabel(pfIdx + 3, 'Total Deductions')) {
    setRole(pfIdx + 3, 'Total Deductions');
  }
  return out;
}

export function readFormXVIIAPExportAmount(row, predicate) {
  if (!row || typeof row !== 'object' || typeof predicate !== 'function') return '';
  for (const [key, raw] of Object.entries(row)) {
    if (raw == null || String(raw).trim() === '') continue;
    if (/^enter\b/i.test(String(raw).trim())) continue;
    if (!predicate(key)) continue;
    const n = parseFormXVIIAPMoney(raw);
    if (Number.isFinite(n)) return n;
  }
  return '';
}

/** Prefer Sample Payroll table PF column (then flattened epf_contribution / employer PF). */
export function resolveFormXVIIAPSamplePayrollPf(payrollRow) {
  if (!payrollRow || payrollRow.fetch_error) return '';
  const flat = flattenPayrollEarningColumns(payrollRow);
  const p = unwrapPayrollPayload(payrollRow);
  const sources = [flat, payrollRow, p].filter((src) => src && typeof src === 'object');
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
  for (let i = 0; i < sources.length; i += 1) {
    const src = sources[i];
    const val = pickFormXVIIAPAmount(
      src.epf_contribution,
      src.pf,
      src.PF,
      src.provident_fund,
      src.employee_pf,
      src['Provident Fund'],
      src.employer_pf,
      src.employer_epf,
      src.employer_amount,
      readPayrollScalar(src, keys, patterns)
    );
    if (val !== '' && val != null) return val;
  }
  return '';
}

/** Prefer Sample Payroll table Professional Tax column. */
export function resolveFormXVIIAPSamplePayrollPt(payrollRow) {
  if (!payrollRow || payrollRow.fetch_error) return '';
  const flat = flattenPayrollEarningColumns(payrollRow);
  const p = unwrapPayrollPayload(payrollRow);
  const sources = [flat, payrollRow, p].filter((src) => src && typeof src === 'object');
  const keys = [
    'Professional Tax',
    'professional_tax',
    'ProfessionalTax',
    'professionalTax',
    'pt',
    'PT',
  ];
  const patterns = [/^professional_tax$/, /^professionaltax$/, /^pt$/];
  for (let i = 0; i < sources.length; i += 1) {
    const src = sources[i];
    const val = pickFormXVIIAPAmount(
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

export function formXVIIAPPayrollRowHasPfOrPt(payrollRow) {
  if (!payrollRow || payrollRow.fetch_error) return false;
  const pf = resolveFormXVIIAPSamplePayrollPf(payrollRow);
  const pt = resolveFormXVIIAPSamplePayrollPt(payrollRow);
  return (pf !== '' && pf != null) || (pt !== '' && pt != null);
}

/** Prefer Sample Payroll snapshots that carry PF / Professional Tax over gross-only pay-run rows. */
export function preferFormXVIIAPPayrollRowsWithPfPt(...rowLists) {
  const flattenRows = (rows) =>
    (Array.isArray(rows) ? rows : [])
      .filter((row) => row && typeof row === 'object' && row.fetch_error !== true)
      .map((row) => flattenPayrollEarningColumns(row));

  let fallback = [];
  for (let i = 0; i < rowLists.length; i += 1) {
    const list = flattenRows(rowLists[i]);
    if (list.length === 0) continue;
    const withPfPt = list.filter((row) => formXVIIAPPayrollRowHasPfOrPt(row));
    if (withPfPt.length > 0) return withPfPt;
    if (fallback.length === 0) fallback = list;
  }
  return fallback;
}

export function applyFormXVIIAPPfPtToRow(row, payrollRow, headers, helpers = {}) {
  if (!row || !payrollRow || payrollRow.fetch_error || !Array.isArray(headers)) return false;
  const { sanitizeValue, overwrite = true } = helpers;
  const cellIsEmpty = (header) => {
    if (!header) return true;
    const v = String(row[header] || '').trim();
    if (!v) return true;
    const s = v.toLowerCase();
    return /^enter\b/.test(s) || s.includes('enter ') || s.includes('select ');
  };
  const setNum = (header, value) => {
    if (!header || value === '' || value == null) return false;
    const num = parseFormXVIIAPMoney(value);
    if (!Number.isFinite(num)) return false;
    if (!overwrite && !cellIsEmpty(header)) return false;
    row[header] = typeof sanitizeValue === 'function' ? sanitizeValue(num) : num;
    return true;
  };

  const pf = resolveFormXVIIAPSamplePayrollPf(payrollRow);
  const pt = resolveFormXVIIAPSamplePayrollPt(payrollRow);
  let hit = false;
  headers.forEach((header) => {
    if (isFormXVIIAPPfHeader(header) && setNum(header, pf)) hit = true;
    if (isFormXVIIAPPtHeader(header) && setNum(header, pt)) hit = true;
  });
  return hit;
}

function readFormXVIIAPExportGrossNet(row) {
  if (!row || typeof row !== 'object') return { gross: '', net: '' };
  const grossKeys = ['gross_pay', 'grossPay', 'Gross_Pay', 'gross_wages', 'grossWages'];
  const netKeys = ['net_pay', 'netPay', 'Net_Pay', 'net_wages', 'netWages'];
  let gross = '';
  let net = '';
  for (const [key, raw] of Object.entries(row)) {
    if (raw == null || String(raw).trim() === '') continue;
    if (/^enter\b/i.test(String(raw).trim())) continue;
    const nk = formXVIIAPHeaderNorm(key);
    const n = parseFormXVIIAPMoney(raw);
    if (!Number.isFinite(n)) continue;
    if (gross === '' && (grossKeys.includes(key) || (nk === 'total' && !nk.includes('deduction')))) {
      gross = n;
    }
    if (
      net === '' &&
      (netKeys.includes(key) ||
        (nk.includes('net') && nk.includes('amount') && nk.includes('paid')))
    ) {
      net = n;
    }
  }
  return { gross, net };
}

function resolveFormXVIIAPExportTotalDeductions(gross, net, pf, pt, pickedTotal) {
  if (pickedTotal !== '' && pickedTotal != null) return pickedTotal;
  const g = parseFormXVIIAPMoney(gross);
  const n = parseFormXVIIAPMoney(net);
  if (Number.isFinite(g) && Number.isFinite(n)) {
    const diff = Math.round((g - n) * 100) / 100;
    if (diff >= 0) return diff;
  }
  const pfN = parseFormXVIIAPMoney(pf);
  const ptN = parseFormXVIIAPMoney(pt);
  if (Number.isFinite(pfN) && Number.isFinite(ptN)) return Math.round((pfN + ptN) * 100) / 100;
  return '';
}

function readFormXVIIAPExportAmountAtHeader(row, header) {
  if (!row || typeof row !== 'object' || !header) return '';
  if (!Object.prototype.hasOwnProperty.call(row, header)) return '';
  const raw = row[header];
  if (raw == null || String(raw).trim() === '') return '';
  if (/^enter\b/i.test(String(raw).trim())) return '';
  const n = parseFormXVIIAPMoney(raw);
  return Number.isFinite(n) ? n : '';
}

/** PT / Total Deductions often sit under duplicate PF keys or blank merged leaves — read by column order. */
function readFormXVIIAPExportPtAndTotalFromHeaderOrder(row, fallbackRow, headers) {
  if (!Array.isArray(headers) || headers.length === 0) {
    return { pt: '', totalDeductions: '' };
  }
  const inferred = inferFormXVIIAPDeductionExportHeaders(headers);
  const pfIdx = inferred.findIndex((h) => isFormXVIIAPPfHeader(h));
  const netIdx = inferred.findIndex((h) => isFormXVIIAPNetAmountPaidHeader(h));
  let pt = '';
  let totalDeductions = '';
  const sources = [row, fallbackRow].filter((src) => src && typeof src === 'object');
  if (pfIdx >= 0 && pfIdx + 1 < inferred.length) {
    const ptHeader = inferred[pfIdx + 1];
    const origHeader = headers[pfIdx + 1];
    for (let i = 0; i < sources.length && pt === ''; i += 1) {
      pt =
        readFormXVIIAPExportAmountAtHeader(sources[i], ptHeader) ||
        (origHeader && origHeader !== ptHeader
          ? readFormXVIIAPExportAmountAtHeader(sources[i], origHeader)
          : '');
    }
  }
  const totalIdx =
    netIdx > 0
      ? netIdx - 1
      : pfIdx >= 0 && pfIdx + 3 < inferred.length
        ? pfIdx + 3
        : -1;
  if (totalIdx >= 0 && totalIdx < inferred.length) {
    const totalHeader = inferred[totalIdx];
    const origTotalHeader = headers[totalIdx];
    for (let i = 0; i < sources.length && totalDeductions === ''; i += 1) {
      totalDeductions =
        readFormXVIIAPExportAmountAtHeader(sources[i], totalHeader) ||
        (origTotalHeader && origTotalHeader !== totalHeader
          ? readFormXVIIAPExportAmountAtHeader(sources[i], origTotalHeader)
          : '');
    }
  }
  return { pt, totalDeductions };
}

export function pickFormXVIIAPExportPtAndTotalDeductions(row, fallbackRow = null, headers = null) {
  const pick = (predicate) => {
    const fromRow = readFormXVIIAPExportAmount(row, predicate);
    if (fromRow !== '' && fromRow != null) return fromRow;
    return readFormXVIIAPExportAmount(fallbackRow, predicate);
  };
  let pt = pick(isFormXVIIAPPtHeader);
  let totalDeductions = pick(isFormXVIIAPTotalDeductionsHeader);
  const pf = pick(isFormXVIIAPPfHeader);
  if (pt === '' || pt == null || totalDeductions === '' || totalDeductions == null) {
    const positional = readFormXVIIAPExportPtAndTotalFromHeaderOrder(row, fallbackRow, headers);
    if (pt === '' || pt == null) pt = positional.pt;
    if (totalDeductions === '' || totalDeductions == null) {
      totalDeductions = positional.totalDeductions;
    }
  }
  if (pt === '' || pt == null) {
    pt =
      resolveFormXVIIAPSamplePayrollPt(row) ||
      resolveFormXVIIAPSamplePayrollPt(fallbackRow) ||
      '';
  }
  const { gross, net } = readFormXVIIAPExportGrossNet(row);
  const fallbackGrossNet = readFormXVIIAPExportGrossNet(fallbackRow);
  totalDeductions = resolveFormXVIIAPExportTotalDeductions(
    gross !== '' ? gross : fallbackGrossNet.gross,
    net !== '' ? net : fallbackGrossNet.net,
    pf,
    pt,
    totalDeductions
  );
  return { pf, pt, totalDeductions };
}

/**
 * Find PF / PT / Total Deductions / Net columns from worksheet header cells.
 * Prefer the first PF so a merged child that inherited "PF" cannot shift PT right.
 */
export function locateFormXVIIAPDeductionColumnsFromHeaderCells(cells) {
  let pfCol = null;
  let ptCol = null;
  let totalDedCol = null;
  let netCol = null;
  (Array.isArray(cells) ? cells : []).forEach((cell) => {
    const col = Number(cell?.col);
    const text = cell?.text;
    if (!Number.isFinite(col) || col < 1) return;
    if (isFormXVIIAPDeductionsParentBanner(text)) return;
    if (pfCol == null && isFormXVIIAPPfHeader(text)) pfCol = col;
    if (isFormXVIIAPPtHeader(text)) ptCol = col;
    if (isFormXVIIAPTotalDeductionsHeader(text)) totalDedCol = col;
    if (isFormXVIIAPNetAmountPaidHeader(text)) netCol = col;
  });
  if (!Number.isFinite(ptCol) && Number.isFinite(pfCol)) ptCol = pfCol + 1;
  if (!Number.isFinite(totalDedCol) && Number.isFinite(netCol) && netCol > 1) {
    totalDedCol = netCol - 1;
  } else if (!Number.isFinite(totalDedCol) && Number.isFinite(pfCol)) {
    totalDedCol = pfCol + 3;
  }
  return { pfCol, ptCol, totalDedCol, netCol };
}

function formXVIIAPExcelCellText(val) {
  return excelCellValueToString(val).trim();
}

function formXVIIAPLabelOnly(raw) {
  return String(raw || '')
    .split(':')[0]
    .trim()
    .replace(/\.+$/, '');
}

function formXVIIAPValueAfterColon(raw) {
  const text = String(raw || '');
  const idx = text.indexOf(':');
  if (idx < 0) return '';
  return text.slice(idx + 1).trim();
}

/** Split AP template lines: "Establishemnt in/" + "under which contract is carried on:" */
function isFormXVIIAPEstablishmentLabelPart(text) {
  const s = normalizeStatutoryHeaderLabel(formXVIIAPLabelOnly(text));
  if (!s) return false;
  if (isEstablishmentContractCarriedHeaderLabel(text)) return true;
  return (
    /establ(?:ishment|ishemnt)/.test(s) ||
    /under\s+which\s+contract/.test(s) ||
    (/name\s+and\s+address/.test(s) && /establ/.test(s))
  );
}

function isFormXVIIAPNatureLabel(text) {
  return isNatureLocationHeaderLabel(formXVIIAPLabelOnly(text));
}

function isFormXVIIAPContractorLabel(text) {
  return isContractorHeaderLabel(formXVIIAPLabelOnly(text));
}

function isFormXVIIAPPrincipalEmployerLabel(text) {
  return isPrincipalEmployerHeaderLabel(formXVIIAPLabelOnly(text));
}

function isFormXVIIAPMetaHeaderLabelText(text) {
  const raw = String(text || '').trim();
  if (!raw) return false;
  return (
    isFormXVIIAPContractorLabel(raw) ||
    isFormXVIIAPNatureLabel(raw) ||
    isFormXVIIAPEstablishmentLabelPart(raw) ||
    isFormXVIIAPPrincipalEmployerLabel(raw) ||
    /wage\s+period/i.test(formXVIIAPLabelOnly(raw))
  );
}

function normalizeFormXVIIAPCompare(text) {
  return String(text || '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '')
    .trim();
}

function valuesLookLikeFormXVIIAPDuplicate(a, b) {
  const left = normalizeFormXVIIAPCompare(a);
  const right = normalizeFormXVIIAPCompare(b);
  if (!left || !right) return false;
  if (left === right) return true;
  if (left.length >= 4 && right.includes(left)) return true;
  if (right.length >= 4 && left.includes(right)) return true;
  return false;
}

/** "AP-NimbagalluAP-Nimbagallu" / "AP-Nimbagallu / AP-Nimbagallu" → one value. */
export function collapseFormXVIIAPRepeatedText(text) {
  let v = String(text || '').replace(/\s+/g, ' ').trim();
  if (!v) return '';
  const slash = v.split(/\s*[\/|]\s*/).map((p) => p.trim()).filter(Boolean);
  if (
    slash.length === 2 &&
    normalizeFormXVIIAPCompare(slash[0]) === normalizeFormXVIIAPCompare(slash[1])
  ) {
    return slash[0];
  }
  const tokens = v.split(/\s+/);
  if (tokens.length >= 2 && tokens.length % 2 === 0) {
    const mid = tokens.length / 2;
    const a = tokens.slice(0, mid).join(' ');
    const b = tokens.slice(mid).join(' ');
    if (normalizeFormXVIIAPCompare(a) === normalizeFormXVIIAPCompare(b)) return a;
  }
  const compact = v.replace(/\s+/g, '');
  if (compact.length >= 8) {
    const half = Math.floor(compact.length / 2);
    if (compact.slice(0, half).toLowerCase() === compact.slice(half, half * 2).toLowerCase()) {
      return v.slice(0, Math.ceil(v.length / 2)).replace(/[\/|,]+$/, '').trim();
    }
  }
  return v;
}

function decodeFormXVIIAPA1(a1) {
  const m = String(a1 || '').match(/^([A-Z]+)(\d+)$/i);
  if (!m) return null;
  let col = 0;
  const letters = m[1].toUpperCase();
  for (let i = 0; i < letters.length; i += 1) {
    col = col * 26 + (letters.charCodeAt(i) - 64);
  }
  return { row: Number(m[2]), col };
}

function listFormXVIIAPMergeBoxes(worksheet) {
  const fromModel = Array.isArray(worksheet?.model?.merges) ? worksheet.model.merges : [];
  const fromMap = worksheet?._merges && typeof worksheet._merges === 'object'
    ? Object.keys(worksheet._merges)
    : [];
  const raw = fromModel.length > 0 ? fromModel : fromMap;
  const out = [];
  for (let i = 0; i < raw.length; i += 1) {
    const parts = String(raw[i] || '').split(':');
    if (parts.length < 2) continue;
    const a = decodeFormXVIIAPA1(parts[0]);
    const b = decodeFormXVIIAPA1(parts[1]);
    if (!a || !b) continue;
    out.push({
      row: Math.min(a.row, b.row),
      col: Math.min(a.col, b.col),
      endRow: Math.max(a.row, b.row),
      endCol: Math.max(a.col, b.col),
    });
  }
  return out;
}

function findFormXVIIAPMergeCovering(worksheet, row, col) {
  const boxes = listFormXVIIAPMergeBoxes(worksheet);
  for (let i = 0; i < boxes.length; i += 1) {
    const box = boxes[i];
    if (row >= box.row && row <= box.endRow && col >= box.col && col <= box.endCol) {
      return box;
    }
  }
  return { row, col, endRow: row, endCol: col };
}

function findFormXVIIAPWagePeriodCell(worksheet, scanEnd, colLimit) {
  for (let r = 1; r <= scanEnd; r += 1) {
    for (let c = 1; c <= Math.min(8, colLimit); c += 1) {
      const raw = formXVIIAPExcelCellText(worksheet.getCell(r, c)?.value);
      if (/wage\s*period/i.test(formXVIIAPLabelOnly(raw))) {
        return { row: r, col: c };
      }
    }
  }
  return null;
}

function stripFormXVIIAPEstablishmentLabelPrefix(raw) {
  return String(raw || '')
    .replace(/^name\s+and\s+address[^\n:]*:\s*/i, '')
    .trim();
}

function unmergeFormXVIIAPRange(worksheet, rowFrom, rowTo, colFrom, colTo) {
  if (!worksheet) return;
  const boxes = listFormXVIIAPMergeBoxes(worksheet);
  const r0 = Math.max(1, Number(rowFrom) || 1);
  const r1 = Math.max(r0, Number(rowTo) || r0);
  const c0 = Math.max(1, Number(colFrom) || 1);
  const c1 = Math.max(c0, Number(colTo) || c0);
  for (let i = 0; i < boxes.length; i += 1) {
    const box = boxes[i];
    if (box.endRow < r0 || box.row > r1 || box.endCol < c0 || box.col > c1) continue;
    try {
      worksheet.unMergeCells(box.row, box.col, box.endRow, box.endCol);
    } catch (_) {
      try {
        if (typeof worksheet.unmergeCells === 'function') {
          worksheet.unmergeCells(box.row, box.col, box.endRow, box.endCol);
        }
      } catch (__) {
        /* already unmerged */
      }
    }
  }
}

/** Label + value on one horizontal line (no wrap / no tall stacked box). */
function writeFormXVIIAPSingleLine(worksheet, row, startCol, endCol, text) {
  if (!worksheet || !text || !row) return;
  const c0 = Math.max(1, Number(startCol) || 1);
  const c1 = Math.max(c0, Number(endCol) || c0);
  unmergeFormXVIIAPRange(worksheet, row, row, c0, c1);
  for (let c = c0; c <= c1; c += 1) {
    const cell = worksheet.getCell(row, c);
    if (c === c0) continue;
    if (cell?.isMerged && cell.master && (cell.master.row !== row || cell.master.col !== c)) continue;
    cell.value = null;
  }
  if (c1 > c0) {
    try {
      worksheet.mergeCells(row, c0, row, c1);
    } catch (_) {
      /* template may already merge */
    }
  }
  const cell = worksheet.getCell(row, c0);
  cell.value = text;
  cell.alignment = {
    ...(cell.alignment || {}),
    wrapText: false,
    shrinkToFit: false,
    vertical: 'middle',
    horizontal: 'left',
  };
  const wsRow = worksheet.getRow(row);
  if (wsRow) wsRow.height = 20;
}

function findFormXVIIAPEstablishmentValueBox(worksheet, establishmentLabels, scanEnd, colLimit) {
  const first = establishmentLabels[0];
  const last = establishmentLabels[establishmentLabels.length - 1];
  const labelCol = first?.col || 9;
  const wage = findFormXVIIAPWagePeriodCell(worksheet, scanEnd, colLimit);
  const boxes = listFormXVIIAPMergeBoxes(worksheet);

  if (wage) {
    let best = null;
    for (let i = 0; i < boxes.length; i += 1) {
      const box = boxes[i];
      if (wage.row >= box.row && wage.row <= box.endRow && box.col >= 8) {
        if (!best || box.col < best.col) best = box;
      }
    }
    if (best) return best;
    const startCol = Math.max(9, labelCol);
    let firstEmpty = 0;
    for (let c = startCol; c <= colLimit; c += 1) {
      const raw = formXVIIAPExcelCellText(worksheet.getCell(wage.row, c)?.value);
      if (!raw) {
        if (!firstEmpty) firstEmpty = c;
        continue;
      }
      if (isFormXVIIAPNatureLabel(raw) || isFormXVIIAPContractorLabel(raw)) continue;
      if (/wage\s*period/i.test(formXVIIAPLabelOnly(raw))) continue;
      return findFormXVIIAPMergeCovering(worksheet, wage.row, c);
    }
    return findFormXVIIAPMergeCovering(
      worksheet,
      wage.row,
      firstEmpty || Math.max(10, labelCol + 1)
    );
  }

  const valueRow = Math.max((first?.row || 5) + 1, last?.row || 6);
  const valueCol = Math.min(Math.max(labelCol + 1, 10), colLimit);
  return findFormXVIIAPMergeCovering(worksheet, valueRow, valueCol);
}

/**
 * Form XVII AP Excel meta band:
 * - Contractor defaults to VAYONA ENERGY PRIVATE LIMITED when site contractor is empty.
 * - Each heading + value sits on one horizontal line (no wrap / stacked columns).
 * - Nature / contractor / establishment values once (no adjacent duplicates).
 */
export function applyFormXVIIAPHeaderLayoutFixes(
  worksheet,
  {
    headerFormData,
    parsedFormHeader,
    headerRowEnd = 12,
    maxCol = 24,
  } = {}
) {
  if (!worksheet) return;
  const data = headerFormData && typeof headerFormData === 'object' ? headerFormData : {};

  const fields = Array.isArray(parsedFormHeader?.fields) ? parsedFormHeader.fields : [];
  const resolveByKeys = (keys, labelHint) => {
    for (let i = 0; i < keys.length; i += 1) {
      const v = resolveHeaderFieldExportValue(data, { key: keys[i], label: labelHint });
      if (v) return v;
    }
    for (let i = 0; i < fields.length; i += 1) {
      const field = fields[i];
      if (!field?.key) continue;
      if (labelHint === 'establishment' && !isEstablishmentContractCarriedHeaderLabel(field.label)) continue;
      if (labelHint === 'nature' && !isNatureLocationHeaderLabel(field.label)) continue;
      if (labelHint === 'contractor' && !isContractorHeaderLabel(field.label)) continue;
      const v = resolveHeaderFieldExportValue(data, field);
      if (v) return v;
    }
    return '';
  };

  const establishmentValue = resolveByKeys(
    [
      'form_xvii_establishment_contract_carried',
      'form_xiii_establishment_contract_carried',
      'form_xxiii_establishment_contract_carried',
      'form_xvi_establishment_contract_carried',
      'statutory_establishment_name_address',
      'statutory_employer_name_address',
    ],
    'establishment'
  );
  const natureValue = collapseFormXVIIAPRepeatedText(
    resolveByKeys(
      ['form_xvii_nature_location_work', 'form_xvi_nature_location_work', 'form_xxiii_nature_location_work'],
      'nature'
    )
  );
  const contractorValue =
    resolveByKeys(
      ['form_xvii_contractor', 'form_xxiii_contractor', 'form_xvi_contractor', 'form25_contractor'],
      'contractor'
    ) || FORM_XVII_AP_DEFAULT_CONTRACTOR_NAME;

  const scanEnd = Math.max(8, Number(headerRowEnd) || 12);
  const colLimit = Math.max(16, Number(maxCol) || 24);

  let contractorLabel = null;
  let natureLabel = null;
  let establishmentLabels = [];

  for (let r = 1; r <= scanEnd; r += 1) {
    for (let c = 1; c <= colLimit; c += 1) {
      const raw = formXVIIAPExcelCellText(worksheet.getCell(r, c)?.value);
      if (!raw) continue;
      if (/form[\s._-]*xvii|register\s+of\s+wages|vide\s+rule/i.test(raw) && r <= 3) continue;
      if (!contractorLabel && isFormXVIIAPContractorLabel(raw)) {
        contractorLabel = { row: r, col: c, raw };
      }
      if (!natureLabel && isFormXVIIAPNatureLabel(raw)) {
        natureLabel = { row: r, col: c, raw };
      }
      if (isFormXVIIAPEstablishmentLabelPart(raw)) {
        establishmentLabels.push({ row: r, col: c, raw });
      }
    }
  }

  if (!contractorLabel) {
    for (let r = 4; r <= scanEnd && !contractorLabel; r += 1) {
      for (let c = 1; c <= 8; c += 1) {
        const raw = formXVIIAPExcelCellText(worksheet.getCell(r, c)?.value);
        if (
          raw &&
          (valuesLookLikeFormXVIIAPDuplicate(contractorValue, raw) || /vayona\s+energy/i.test(raw)) &&
          !isFormXVIIAPNatureLabel(raw) &&
          !/wage\s*period/i.test(raw)
        ) {
          contractorLabel = { row: r, col: 1, raw: 'Name and address of Contractor:' };
          break;
        }
      }
    }
  }
  if (!contractorLabel) contractorLabel = { row: 5, col: 1, raw: 'Name and address of Contractor:' };
  if (!natureLabel) natureLabel = { row: 7, col: 1, raw: 'Nature and location of work:' };

  establishmentLabels.sort((a, b) => a.row - b.row || a.col - b.col);
  const rightStart =
    establishmentLabels[0]?.col > 4 ? establishmentLabels[0].col : 9;
  const leftEnd = Math.max(4, rightStart - 1);
  const wage = findFormXVIIAPWagePeriodCell(worksheet, scanEnd, colLimit);
  const oldBox = findFormXVIIAPEstablishmentValueBox(
    worksheet,
    establishmentLabels.length ? establishmentLabels : [{ row: 5, col: rightStart }],
    scanEnd,
    colLimit
  );

  const contractor = String(
    contractorValue ||
      formXVIIAPValueAfterColon(contractorLabel.raw) ||
      FORM_XVII_AP_DEFAULT_CONTRACTOR_NAME
  ).trim();
  const natureExisting = collapseFormXVIIAPRepeatedText(
    formXVIIAPValueAfterColon(natureLabel.raw) ||
      formXVIIAPExcelCellText(worksheet.getCell(natureLabel.row, natureLabel.col + 1)?.value) ||
      ''
  );
  const nature = collapseFormXVIIAPRepeatedText(natureValue || natureExisting || '');

  writeFormXVIIAPSingleLine(
    worksheet,
    contractorLabel.row,
    1,
    leftEnd,
    formatStatutoryHeaderLabelValueExport(
      'Name and address of Contractor',
      contractorLabel.raw,
      contractor
    )
  );

  if (nature) {
    writeFormXVIIAPSingleLine(
      worksheet,
      natureLabel.row,
      1,
      leftEnd,
      formatStatutoryHeaderLabelValueExport(
        'Nature and location of work',
        natureLabel.raw,
        nature
      )
    );
  }

  if (wage) {
    const wageVal = (
      formXVIIAPValueAfterColon(formXVIIAPExcelCellText(worksheet.getCell(wage.row, wage.col)?.value)) ||
      'monthly'
    ).replace(/:+\s*$/g, '').trim() || 'monthly';
    writeFormXVIIAPSingleLine(
      worksheet,
      wage.row,
      1,
      leftEnd,
      formatStatutoryHeaderLabelValueExport('Wage period', 'Wage period', wageVal)
    );
  }

  if (oldBox && (oldBox.endRow > oldBox.row || oldBox.endCol > oldBox.col)) {
    unmergeFormXVIIAPRange(worksheet, oldBox.row, oldBox.endRow, oldBox.col, oldBox.endCol);
    for (let r = oldBox.row; r <= (oldBox.endRow || oldBox.row); r += 1) {
      for (let c = oldBox.col; c <= (oldBox.endCol || oldBox.col); c += 1) {
        const cell = worksheet.getCell(r, c);
        const raw = formXVIIAPExcelCellText(cell?.value);
        if (/wage\s*period/i.test(formXVIIAPLabelOnly(raw))) continue;
        cell.value = null;
      }
      if (r >= 4) {
        const wsRow = worksheet.getRow(r);
        if (Number(wsRow.height) > 22) wsRow.height = 20;
      }
    }
  }

  if (establishmentValue) {
    const estRow = establishmentLabels[0]?.row || contractorLabel.row || 5;
    writeFormXVIIAPSingleLine(
      worksheet,
      estRow,
      rightStart,
      colLimit,
      formatStatutoryHeaderLabelValueExport(
        'Name and address of Establishment in/under which contract is carried on',
        establishmentLabels[0]?.raw || '',
        establishmentValue
      )
    );
    for (let i = 1; i < establishmentLabels.length; i += 1) {
      const extra = establishmentLabels[i];
      if (extra.row === estRow) continue;
      const cell = worksheet.getCell(extra.row, extra.col);
      if (isFormXVIIAPEstablishmentLabelPart(formXVIIAPExcelCellText(cell?.value))) {
        cell.value = null;
      }
    }
  }

  // Drop leftover duplicate values on the left band (same text in column F, etc.).
  [contractorLabel.row, natureLabel.row, wage?.row].filter(Boolean).forEach((row) => {
    for (let c = 2; c <= leftEnd; c += 1) {
      const cell = worksheet.getCell(row, c);
      if (cell?.isMerged && cell.master && (cell.master.row !== row || cell.master.col !== c)) continue;
      if (c === 1) continue;
      const raw = formXVIIAPExcelCellText(cell?.value);
      if (!raw) continue;
      if (isFormXVIIAPMetaHeaderLabelText(raw) && formXVIIAPValueAfterColon(raw)) continue;
      if (
        valuesLookLikeFormXVIIAPDuplicate(contractor, raw) ||
        valuesLookLikeFormXVIIAPDuplicate(nature, raw) ||
        valuesLookLikeFormXVIIAPDuplicate(establishmentValue, raw)
      ) {
        cell.value = null;
      }
    }
  });
}
