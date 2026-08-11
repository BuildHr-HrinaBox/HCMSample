/** Rajasthan Form 11 — Register of Employment [Rule 22 sub-rule 1]. */

import * as XLSX from 'xlsx';

export const FORM_11_RJ_DISPLAY_TITLE = 'FORM 11';
export const FORM_11_RJ_DISPLAY_SUBTITLE = 'Register of Employment';
export const FORM_11_RJ_DISPLAY_REFERENCE = 'Rule 22 - sub rule 1';

/** Leaf columns (12) matching Form_11_RJ.xlsx after resolving the Days-of-Month band. */
export const FORM_11_RJ_TABLE_HEADERS = [
  'Name of Persons Employed',
  'Whether young person or not',
  'Time at which employment commences',
  'Time at which employment ceases',
  '1',
  '2 Rest interval',
  '3',
  'Total hours worked during the month',
  '*Days on which overtime work is done and extent of such overtime on each day',
  'Extent of overtime worked during the month',
  'Extent of overtime worked during the quarter',
  'Extent of overtime worked during the year',
];

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

/** Table looks like RJ Register of Employment (not ESIC Accident Book / AP period of work). */
export function headersIndicateForm11RajasthanEmployment(tableHeaders) {
  const list = (Array.isArray(tableHeaders) ? tableHeaders : [])
    .map((h) => stripDedupeSuffix(h))
    .filter(Boolean);
  if (list.length < 4) return false;
  const joined = list.map((h) => norm(h)).join(' | ');
  if (/accident\s+book|insurance\s+no|cause\s+of\s+injury|periods?\s+of\s+work|group\s+letter/.test(joined)) {
    return false;
  }
  const hasName = list.some((h) => /name\s+of\s+persons?\s+employ/.test(norm(h)));
  const hasYoung = list.some((h) => /young\s+person/.test(norm(h)));
  const hasCommences = list.some((h) => /employment\s+commences|time\s+at\s+which\s+employment\s+commence/.test(norm(h)));
  const hasRest = list.some((h) => /rest\s+interval/.test(norm(h)));
  const hasTotalHours = list.some((h) => /total\s+hours\s+worked/.test(norm(h)));
  const hasOtExtent = list.some((h) => /extent\s+of\s+overtime/.test(norm(h)));
  return hasName && (hasYoung || hasCommences) && (hasRest || hasTotalHours || hasOtExtent || hasCommences);
}

/**
 * Rajasthan Shops Form 11 (Register of Employment).
 * Must not match TN ESIC Form 11 Accident Book or AP Form 11 Period of Work.
 */
export function isForm11RajasthanContext(
  formHeader,
  rowItem,
  fileName,
  sheetText = '',
  tableHeaders = []
) {
  const parts = buildContextBlob(formHeader, rowItem, fileName, sheetText, tableHeaders);

  if (/tamil\s*nadu|\b_tn\b|accident\s+book|employees.? state insurance|regulation\s*66/.test(parts)) {
    return false;
  }
  if (/andhra|\b_ap\b|period\s+of\s+work|rule\s*79/.test(parts) && !/rajasthan|\b_rj\b/.test(parts)) {
    return false;
  }

  if (/form[\s._-]*11[\s._-]*rj|\bform_11_rj\b|\b11_rj\b/.test(parts)) return true;
  if (
    /rajasthan/.test(parts) &&
    /\bform[\s._-]*11\b/.test(parts) &&
    (/register\s+of\s+employment|rule\s*22/.test(parts) ||
      headersIndicateForm11RajasthanEmployment(tableHeaders))
  ) {
    return true;
  }
  if (
    headersIndicateForm11RajasthanEmployment(tableHeaders) &&
    /rajasthan|\b_rj\b|shops\s+and\s+establishments/.test(parts)
  ) {
    return true;
  }
  return false;
}

export function classifyForm11RJHeader(header) {
  const s = norm(stripDedupeSuffix(header));
  if (!s) return '';
  if (/name\s+of\s+persons?\s+employ/.test(s) || (s.includes('name') && s.includes('employ') && !/young/.test(s))) {
    return 'name';
  }
  if (/young\s+person/.test(s)) return 'young';
  if (/employment\s+commences|time\s+at\s+which\s+employment\s+commence/.test(s)) return 'commences';
  if (/employment\s+ceases|time\s+at\s+which\s+employment\s+cease/.test(s)) return 'ceases';
  if (/^1$/.test(s) || /^day\s*1\b/.test(s)) return 'day1';
  if (/rest\s+interval/.test(s) || /^2(\s|$)/.test(s)) return 'rest';
  if (/^3$/.test(s) || /^day\s*3\b/.test(s)) return 'day3';
  if (/total\s+hours\s+worked/.test(s)) return 'totalHours';
  if (/days\s+on\s+which\s+overtime|overtime\s+work\s+is\s+done/.test(s)) return 'otDays';
  if (/extent\s+of\s+overtime.*month|overtime\s+worked\s+during\s+the\s+month/.test(s)) return 'otMonth';
  if (/extent\s+of\s+overtime.*quarter|overtime\s+worked\s+during\s+the\s+quarter/.test(s)) {
    return 'otQuarter';
  }
  if (/extent\s+of\s+overtime.*year|overtime\s+worked\s+during\s+the\s+year/.test(s)) return 'otYear';
  return '';
}

const ROLE_TO_CANONICAL = {
  name: FORM_11_RJ_TABLE_HEADERS[0],
  young: FORM_11_RJ_TABLE_HEADERS[1],
  commences: FORM_11_RJ_TABLE_HEADERS[2],
  ceases: FORM_11_RJ_TABLE_HEADERS[3],
  day1: FORM_11_RJ_TABLE_HEADERS[4],
  rest: FORM_11_RJ_TABLE_HEADERS[5],
  day3: FORM_11_RJ_TABLE_HEADERS[6],
  totalHours: FORM_11_RJ_TABLE_HEADERS[7],
  otDays: FORM_11_RJ_TABLE_HEADERS[8],
  otMonth: FORM_11_RJ_TABLE_HEADERS[9],
  otQuarter: FORM_11_RJ_TABLE_HEADERS[10],
  otYear: FORM_11_RJ_TABLE_HEADERS[11],
};

/** Prefer template wording when all 12 roles are present; else canonical. */
export function resolveForm11RajasthanTableHeaders(tableHeaders) {
  const src = (Array.isArray(tableHeaders) ? tableHeaders : [])
    .map((h) => stripDedupeSuffix(h))
    .filter(Boolean);

  if (src.length >= 8 && headersIndicateForm11RajasthanEmployment(src)) {
    const byRole = {};
    src.forEach((h) => {
      const role = classifyForm11RJHeader(h);
      if (role && !byRole[role]) byRole[role] = h;
    });
    const roles = Object.keys(ROLE_TO_CANONICAL);
    const mapped = roles.map((role) => byRole[role] || ROLE_TO_CANONICAL[role]);
    // Day "1" / "3" often vanish when stacked-header combine skips pure digits.
    if (mapped.filter(Boolean).length >= 8) return mapped;
  }
  return [...FORM_11_RJ_TABLE_HEADERS];
}

function readFirstNonEmpty(row, headers, className) {
  if (!row || typeof row !== 'object') return '';
  for (let i = 0; i < headers.length; i += 1) {
    const h = headers[i];
    if (classifyForm11RJHeader(h) !== className) continue;
    const v = row[h];
    if (v != null && String(v).trim() !== '') return v;
  }
  for (const [k, v] of Object.entries(row)) {
    if (String(k).startsWith('__')) continue;
    if (classifyForm11RJHeader(k) !== className) continue;
    if (v != null && String(v).trim() !== '') return v;
  }
  return '';
}

/** Strip HTML / help text that Zoho sometimes dumps into cells. */
export function sanitizeForm11RajasthanCellValue(value) {
  let s = String(value ?? '').trim();
  if (!s) return '';
  if (/<[^>]+>/.test(s) || /please\s+follow\s+the\s+steps/i.test(s)) return '';
  s = s.replace(/<[^>]*>/g, '').trim();
  return s;
}

export function remapForm11RajasthanRowsToHeaders(rows, sourceHeaders, targetHeaders) {
  const src = Array.isArray(sourceHeaders) ? sourceHeaders : [];
  const tgt = resolveForm11RajasthanTableHeaders(targetHeaders?.length ? targetHeaders : src);
  if (!Array.isArray(rows)) return [];
  return rows.map((row) => {
    if (!row || typeof row !== 'object') {
      const blank = {};
      tgt.forEach((h) => {
        blank[h] = '';
      });
      return blank;
    }
    const out = {};
    tgt.forEach((targetHeader) => {
      const cls = classifyForm11RJHeader(targetHeader);
      let val = '';
      if (Object.prototype.hasOwnProperty.call(row, targetHeader)) {
        val = row[targetHeader];
      } else if (cls) {
        val = readFirstNonEmpty(row, src, cls);
      }
      out[targetHeader] = sanitizeForm11RajasthanCellValue(val);
    });
    Object.keys(row).forEach((k) => {
      if (String(k).startsWith('__')) out[k] = row[k];
    });
    return out;
  });
}

export function isForm11RajasthanNameHeader(header) {
  return classifyForm11RJHeader(header) === 'name';
}

export function isForm11RajasthanYoungPersonHeader(header) {
  return classifyForm11RJHeader(header) === 'young';
}

export function isForm11RajasthanSkipAutofillHeader(header) {
  const role = classifyForm11RJHeader(header);
  return role && role !== 'name' && role !== 'young';
}

/** Young person column: Yes when age &lt; 18, else No. */
export function resolveForm11RajasthanYoungPersonValue(ageYears) {
  const n = Number(ageYears);
  if (!Number.isFinite(n) || n <= 0) return 'No';
  return n < 18 ? 'Yes' : 'No';
}

export function enrichForm11RajasthanDisplayHeader(formHeader, fileName = '', item = null, tableHeaders = [], sheetText = '') {
  if (!isForm11RajasthanContext(formHeader, item, fileName, sheetText, tableHeaders)) {
    return formHeader;
  }
  const base = formHeader && typeof formHeader === 'object' ? { ...formHeader } : {};
  base.title = FORM_11_RJ_DISPLAY_TITLE;
  base.subtitle = FORM_11_RJ_DISPLAY_SUBTITLE;
  base.reference = FORM_11_RJ_DISPLAY_REFERENCE;
  base.form11RJEmploymentLayout = true;

  let fields = Array.isArray(base.fields) ? [...base.fields] : [];
  const hasKey = (k) => fields.some((f) => f.key === k);
  if (!hasKey('form_11_rj_month')) {
    fields.push({ label: 'Month', value: '', key: 'form_11_rj_month' });
  }
  if (!hasKey('form_11_rj_year')) {
    fields.push({ label: 'Year', value: '', key: 'form_11_rj_year' });
  }
  base.fields = fields;
  return base;
}

export function isForm11RajasthanEmploymentLayoutFormHeader(formHeader) {
  return !!formHeader?.form11RJEmploymentLayout;
}

/**
 * Locate header band + data start from workbook (skip column-number row 1…12).
 */
export function repairForm11RajasthanTableHeadersFromWorkbook(workbook, hints = {}) {
  if (!workbook || !workbook.SheetNames || !workbook.SheetNames.length) return null;
  const sheetName = hints.sheetName || workbook.SheetNames[0];
  const ws = workbook.Sheets[sheetName];
  if (!ws) return null;
  const merges = ws['!merges'] || [];
  const ref = ws['!ref'];
  if (!ref) return null;
  const range = XLSX.utils.decode_range(ref);

  const rawCell = (r, c) => {
    const cell = ws[XLSX.utils.encode_cell({ r, c })];
    return cell && cell.v != null ? String(cell.v).trim() : '';
  };
  const mergedAware = (r, c) => {
    const direct = rawCell(r, c);
    if (direct) return direct;
    for (let i = 0; i < merges.length; i += 1) {
      const m = merges[i];
      if (!m?.s || !m?.e) continue;
      if (r >= m.s.r && r <= m.e.r && c >= m.s.c && c <= m.e.c) {
        return rawCell(m.s.r, m.s.c);
      }
    }
    return '';
  };

  let headerRow = -1;
  let startCol = range.s.c;
  for (let r = range.s.r; r <= Math.min(range.e.r, range.s.r + 30); r += 1) {
    for (let c = range.s.c; c <= range.e.c; c += 1) {
      const t = norm(mergedAware(r, c));
      if (/name\s+of\s+persons?\s+employ/.test(t)) {
        headerRow = r;
        startCol = c;
        break;
      }
    }
    if (headerRow >= 0) break;
  }
  if (headerRow < 0) return null;

  const leafRow = headerRow + 1;
  const headers = [];
  for (let i = 0; i < 12; i += 1) {
    const c = startCol + i;
    const top = stripDedupeSuffix(mergedAware(headerRow, c));
    const leaf = stripDedupeSuffix(mergedAware(leafRow, c));
    let chosen = '';
    if (i === 0) chosen = top || FORM_11_RJ_TABLE_HEADERS[0];
    else if (i === 1) chosen = top || FORM_11_RJ_TABLE_HEADERS[1];
    else if (i === 2) chosen = leaf || FORM_11_RJ_TABLE_HEADERS[2];
    else if (i === 3) chosen = leaf || FORM_11_RJ_TABLE_HEADERS[3];
    else if (i === 4) chosen = /^\d+$/.test(top) ? top : FORM_11_RJ_TABLE_HEADERS[4];
    else if (i === 5) {
      const day = /^\d+$/.test(top) ? top : '2';
      const rest = /rest\s+interval/i.test(leaf) ? leaf : 'Rest interval';
      chosen = `${day} ${rest}`.replace(/\s+/g, ' ').trim();
    } else if (i === 6) chosen = /^\d+$/.test(top) ? top : FORM_11_RJ_TABLE_HEADERS[6];
    else if (i >= 7) chosen = top || FORM_11_RJ_TABLE_HEADERS[i];
    headers.push(chosen || FORM_11_RJ_TABLE_HEADERS[i]);
  }

  // Column-number row (1…12) sits under the leaf band.
  let dataStart = leafRow + 1;
  let numHits = 0;
  for (let i = 0; i < 12; i += 1) {
    const t = String(mergedAware(dataStart, startCol + i) || '').trim();
    if (t === String(i + 1)) numHits += 1;
  }
  if (numHits >= 8) dataStart += 1;

  return {
    headers: resolveForm11RajasthanTableHeaders(headers),
    headerRowIndex: headerRow,
    dataStartIndex: dataStart,
    tableStartCol: startCol,
    stackedHeaderBandEnd: leafRow,
  };
}
