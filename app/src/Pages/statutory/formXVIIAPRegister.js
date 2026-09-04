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

function writeFormXVIIAPValueCell(cell, text) {
  if (!cell) return;
  const val = String(text || '').trim();
  cell.value = val || null;
  cell.alignment = {
    ...(cell.alignment || {}),
    wrapText: true,
    vertical: 'top',
    horizontal: 'left',
  };
}

/** Prefer an existing value cell to the right of the label (AP templates often skip a merge slave). */
function findFormXVIIAPValueCol(worksheet, row, labelCol, colLimit, stopAtCol = 0) {
  const start = Math.max(1, Number(labelCol) || 1) + 1;
  const end = Math.min(Number(colLimit) || start, start + 5, stopAtCol > 0 ? stopAtCol - 1 : Number(colLimit) || start);
  let firstEmpty = 0;
  for (let c = start; c <= end; c += 1) {
    const raw = formXVIIAPExcelCellText(worksheet.getCell(row, c)?.value);
    if (!raw) {
      if (!firstEmpty) firstEmpty = c;
      continue;
    }
    if (isFormXVIIAPMetaHeaderLabelText(raw)) return firstEmpty || Math.min(start, colLimit);
    return c;
  }
  return firstEmpty || Math.min(start, colLimit);
}

/**
 * Clear horizontal duplicate value cells on a meta row (e.g. AP-Tadipatri AP-Tadipatri).
 * Keeps the first value cell in [valueStartCol, valueEndCol]; clears other unlabeled values.
 */
function dedupeFormXVIIAPAdjacentValueCells(worksheet, row, valueStartCol, valueEndCol, keepValue) {
  if (!worksheet || !row) return;
  const start = Math.max(1, Number(valueStartCol) || 1);
  const end = Math.max(start, Number(valueEndCol) || start);
  const keep = String(keepValue || '').trim();
  let kept = false;
  for (let c = start; c <= end; c += 1) {
    const cell = worksheet.getCell(row, c);
    const raw = formXVIIAPExcelCellText(cell?.value);
    if (isFormXVIIAPMetaHeaderLabelText(raw)) continue;
    if (!kept && keep) {
      writeFormXVIIAPValueCell(cell, keep);
      kept = true;
      continue;
    }
    // Drop leftover template / adjacent spill cells in the value band.
    if (raw) cell.value = null;
  }
}

/**
 * Form XVII AP Excel meta band:
 * - Establishment (company name + address) once, on the row BELOW the first establishment label line
 *   (template splits "Establishemnt in/" / "under which contract…").
 * - Nature / contractor values once (no adjacent duplicates).
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
  if (!worksheet || !headerFormData || typeof headerFormData !== 'object') return;

  const fields = Array.isArray(parsedFormHeader?.fields) ? parsedFormHeader.fields : [];
  const resolveByKeys = (keys, labelHint) => {
    for (let i = 0; i < keys.length; i += 1) {
      const v = resolveHeaderFieldExportValue(headerFormData, { key: keys[i], label: labelHint });
      if (v) return v;
    }
    for (let i = 0; i < fields.length; i += 1) {
      const field = fields[i];
      if (!field?.key) continue;
      if (labelHint === 'establishment' && !isEstablishmentContractCarriedHeaderLabel(field.label)) continue;
      if (labelHint === 'nature' && !isNatureLocationHeaderLabel(field.label)) continue;
      if (labelHint === 'contractor' && !isContractorHeaderLabel(field.label)) continue;
      const v = resolveHeaderFieldExportValue(headerFormData, field);
      if (v) return v;
    }
    return '';
  };

  const establishmentValue = resolveByKeys(
    ['form_xvii_establishment_contract_carried', 'form_xiii_establishment_contract_carried'],
    'establishment'
  );
  const natureValue = resolveByKeys(['form_xvii_nature_location_work', 'form_xvi_nature_location_work'], 'nature');
  const contractorValue = resolveByKeys(['form_xvii_contractor', 'form_xxiii_contractor'], 'contractor');

  const scanEnd = Math.max(8, Number(headerRowEnd) || 12);
  const colLimit = Math.max(16, Number(maxCol) || 24);

  let contractorLabel = null;
  let natureLabel = null;
  let establishmentLabels = [];

  for (let r = 1; r <= scanEnd; r += 1) {
    for (let c = 1; c <= colLimit; c += 1) {
      const raw = formXVIIAPExcelCellText(worksheet.getCell(r, c)?.value);
      if (!raw) continue;
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

  // Nature: one value only, immediately beside the label.
  if (natureLabel) {
    const rightBound = establishmentLabels[0]?.col > natureLabel.col ? establishmentLabels[0].col : 0;
    const valueCol = findFormXVIIAPValueCol(worksheet, natureLabel.row, natureLabel.col, colLimit, rightBound);
    const existing =
      formXVIIAPValueAfterColon(natureLabel.raw) ||
      formXVIIAPExcelCellText(worksheet.getCell(natureLabel.row, valueCol)?.value) ||
      '';
    const nature = String(natureValue || existing || '').trim();
    // Keep label cell as label-only when value sits beside it.
    if (formXVIIAPValueAfterColon(natureLabel.raw)) {
      worksheet.getCell(natureLabel.row, natureLabel.col).value = `${formXVIIAPLabelOnly(natureLabel.raw)}:`;
    }
    if (nature) {
      writeFormXVIIAPValueCell(worksheet.getCell(natureLabel.row, valueCol), nature);
      dedupeFormXVIIAPAdjacentValueCells(
        worksheet,
        natureLabel.row,
        valueCol,
        Math.min(valueCol + 6, rightBound > 0 ? rightBound - 1 : Math.max(7, valueCol + 4)),
        nature
      );
    }
  }

  // Contractor: one value only beside the label (never duplicate horizontally).
  if (contractorLabel) {
    const rightBound = establishmentLabels[0]?.col > contractorLabel.col ? establishmentLabels[0].col : 0;
    const valueCol = findFormXVIIAPValueCol(
      worksheet,
      contractorLabel.row,
      contractorLabel.col,
      colLimit,
      rightBound
    );
    const existing =
      formXVIIAPValueAfterColon(contractorLabel.raw) ||
      formXVIIAPExcelCellText(worksheet.getCell(contractorLabel.row, valueCol)?.value) ||
      '';
    const contractor = String(contractorValue || existing || '').trim();
    if (formXVIIAPValueAfterColon(contractorLabel.raw)) {
      worksheet.getCell(contractorLabel.row, contractorLabel.col).value =
        `${formXVIIAPLabelOnly(contractorLabel.raw)}:`;
    }
    if (contractor) {
      writeFormXVIIAPValueCell(worksheet.getCell(contractorLabel.row, valueCol), contractor);
      dedupeFormXVIIAPAdjacentValueCells(
        worksheet,
        contractorLabel.row,
        valueCol,
        Math.min(valueCol + 6, rightBound > 0 ? rightBound - 1 : Math.max(7, valueCol + 4)),
        contractor
      );
    }
  }

  // Establishment: company name + address once, BELOW the first label line (not on the upper band).
  if (establishmentLabels.length > 0) {
    establishmentLabels.sort((a, b) => a.row - b.row || a.col - b.col);
    const first = establishmentLabels[0];
    const last = establishmentLabels[establishmentLabels.length - 1];
    const labelCol = first.col;
    // Prefer a value column to the right of the establishment label block.
    const valueCol = Math.min(Math.max(labelCol + 1, 10), colLimit);
    // Put company text on the continuation row (below "Establishemnt in/"), not above it.
    const valueRow = Math.max(first.row + 1, last.row);

    // Clear company / establishment spill from the upper establishment band and duplicate cells.
    for (let r = first.row; r <= Math.min(valueRow + 1, scanEnd); r += 1) {
      for (let c = labelCol; c <= colLimit; c += 1) {
        if (r === valueRow && c === valueCol) continue;
        const cell = worksheet.getCell(r, c);
        const raw = formXVIIAPExcelCellText(cell?.value);
        if (!raw) continue;
        if (isFormXVIIAPEstablishmentLabelPart(raw)) {
          // Keep label text only (strip any inline value on the upper line).
          if (formXVIIAPValueAfterColon(raw)) {
            cell.value = `${formXVIIAPLabelOnly(raw)}:`;
          }
          continue;
        }
        if (isFormXVIIAPPrincipalEmployerLabel(raw) || isFormXVIIAPNatureLabel(raw)) continue;
        if (
          establishmentValue &&
          (valuesLookLikeFormXVIIAPDuplicate(establishmentValue, raw) ||
            valuesLookLikeFormXVIIAPDuplicate(establishmentValue, formXVIIAPValueAfterColon(raw) || raw))
        ) {
          cell.value = null;
        }
      }
    }

    if (establishmentValue) {
      // Prefer plain value under the label; combine only if the target cell still holds a label fragment.
      const target = worksheet.getCell(valueRow, valueCol);
      const targetRaw = formXVIIAPExcelCellText(target?.value);
      if (isFormXVIIAPEstablishmentLabelPart(targetRaw)) {
        target.value = formatStatutoryHeaderLabelValueExport(
          formXVIIAPLabelOnly(targetRaw) || 'Name and address of Establishment in/under which contract is carried on',
          targetRaw,
          establishmentValue
        );
      } else {
        writeFormXVIIAPValueCell(target, establishmentValue);
      }
      dedupeFormXVIIAPAdjacentValueCells(
        worksheet,
        valueRow,
        valueCol,
        Math.min(valueCol + 8, colLimit),
        establishmentValue
      );
    }
  }
}
