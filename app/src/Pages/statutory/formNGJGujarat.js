/** Gujarat Form N — Leave Book (See rule 17). */

import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import * as XLSX from 'xlsx';
import { ensureExcelJSDataRowsWithBorders } from '../../utils/excelTableBorders';
import {
  buildLeaveRecordLookupMap,
  findLeaveRecordForFormRow,
} from '../../utils/leaveMetrics';
import {
  collectApprovedLeaveIdentityKeys,
  getApprovedLeaveRecordUniqueKey,
  normalizePersonNameKey,
  resolveApprovedLeavePeriodForFormNGJ,
} from './formFKarnataka';
import {
  filterApprovedLeaveRecordsForFormOGJMonth,
  readApprovedLeaveLeaveType,
} from './formOGJGujarat';
import {
  computeFormXLeaveBeginning,
  FORM_X_MEDICAL_LEAVE_TYPE_ALIASES,
  getFormXMedicalLeaveApiMetrics,
  leaveTypeLabelMatchesAliases,
} from './formXTamilNaduLeave';

export const FORM_NGJ_WORKER_HEADERS = [
  'Name of the worker',
  'Description of the Department (if applicable)',
  'Name of the employer',
  'Date of entry into service',
  'Receipt of level book',
];

export const FORM_NGJ_LEAVE_PARENT = {
  accumulation: 'Accumulation of leave',
  leaveAllowed: 'Leave allowed',
  paymentLeave: 'Payment for leave made on',
  refusal: 'Refusal of leave',
  paymentDischarge:
    'Payment for Leave on discharge of an worker quitting employment if admissible',
};

export const FORM_NGJ_FESTIVAL_PARENT = 'Details of Festival Leave';
export const FORM_NGJ_CASUAL_PARENT = 'Details of Casual Leave';

export const FORM_NGJ_LEAVE_TABLE_HEADERS = [
  `${FORM_NGJ_LEAVE_PARENT.accumulation}_Leave due on`,
  `${FORM_NGJ_LEAVE_PARENT.accumulation}_No. of days`,
  `${FORM_NGJ_LEAVE_PARENT.leaveAllowed}_From To`,
  `${FORM_NGJ_LEAVE_PARENT.paymentLeave}_1st Moiety`,
  `${FORM_NGJ_LEAVE_PARENT.paymentLeave}_2nd Moiety`,
  `${FORM_NGJ_LEAVE_PARENT.refusal}_Application Date`,
  `${FORM_NGJ_LEAVE_PARENT.refusal}_Date of Refusal`,
  `${FORM_NGJ_LEAVE_PARENT.paymentDischarge}_Date of discharge`,
  `${FORM_NGJ_LEAVE_PARENT.paymentDischarge}_Date and amount paid`,
  `${FORM_NGJ_LEAVE_PARENT.paymentDischarge}_Signature or thumb impression of worker`,
  'Remarks',
];

export const FORM_NGJ_FESTIVAL_HEADERS = [
  `${FORM_NGJ_FESTIVAL_PARENT}_Period_From`,
  `${FORM_NGJ_FESTIVAL_PARENT}_Period_To`,
  `${FORM_NGJ_FESTIVAL_PARENT}_Total Leave`,
  `${FORM_NGJ_FESTIVAL_PARENT}_Availed Leave`,
  `${FORM_NGJ_FESTIVAL_PARENT}_Balance Leave`,
  `${FORM_NGJ_FESTIVAL_PARENT}_Payment made in lieu of Festival Leave, when called`,
  `${FORM_NGJ_FESTIVAL_PARENT}_Remarks`,
];

export const FORM_NGJ_CASUAL_HEADERS = [
  `${FORM_NGJ_CASUAL_PARENT}_Period_From`,
  `${FORM_NGJ_CASUAL_PARENT}_Period_To`,
  `${FORM_NGJ_CASUAL_PARENT}_Total Leave`,
  `${FORM_NGJ_CASUAL_PARENT}_Availed Leave`,
  `${FORM_NGJ_CASUAL_PARENT}_Balance Leave`,
  `${FORM_NGJ_CASUAL_PARENT}_Remarks`,
];

export const FORM_NGJ_GJ_CANONICAL_TABLE_HEADERS = [
  ...FORM_NGJ_WORKER_HEADERS,
  ...FORM_NGJ_LEAVE_TABLE_HEADERS,
  ...FORM_NGJ_FESTIVAL_HEADERS,
  ...FORM_NGJ_CASUAL_HEADERS,
];

export function formNGJGujaratHeaderNorm(txt) {
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
  return formNGJGujaratHeaderNorm(stripLeadingNumber(h))
    .replace(/[^a-z0-9 ]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

export function isFormNGJGujaratContext(
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

  if (/\bform\s*m\b|\bform_m\b|form[\s._-]*m[\s._-]*gj/.test(parts) && /identity\s+card/.test(parts)) {
    return false;
  }

  const hasGujarat = /gujarat|\b_gj\b|form[\s._-]*n[\s._-]*gj|form_n_gj/.test(parts);
  const hasFormN = /\bform[\s._-]*n\b/.test(parts);
  if (hasGujarat && hasFormN) return true;
  if (/\bform_n_gj\b/.test(parts)) return true;
  if (hasFormN && /leave\s+book/.test(parts)) return true;
  if (hasFormN && /see\s+rule\s+17/.test(parts)) return true;

  const joined = parts;
  return (
    /leave\s+book/.test(joined) &&
    /accumulation\s+of\s+leave/.test(joined) &&
    (/leave\s+due\s+on/.test(joined) || /festival\s+leave/.test(joined))
  );
}

export function headersLookLikeBrokenFormNGJ(headers) {
  return (Array.isArray(headers) ? headers : []).some((h) => {
    const text = String(h || '');
    return (
      (/name\s+of\s+the\s+establishment/i.test(text) &&
        /name\s+of\s+the\s+worker/i.test(text)) ||
      (/accumulation\s+of\s+leave/i.test(text) &&
        text.split(/accumulation\s+of\s+leave/i).length > 2) ||
      (/description\s+of\s+the\s+department/i.test(text) &&
        /leave\s+due\s+on/i.test(text) &&
        text.length > 80)
    );
  });
}

export function isFormNGJGujaratHeaderFieldLayoutFormHeader(formHeader) {
  return !!formHeader?.formNGJGujaratHeaderFieldLayout;
}

export function isFormNGJGujaratTableLayoutFormHeader(formHeader) {
  return !!formHeader?.formNGJGujaratTableLayout;
}

function upgradeFormNGJPeriodHeaders(headers) {
  const out = [];
  const seen = new Set();
  (Array.isArray(headers) ? headers : []).forEach((header) => {
    const raw = String(header || '').trim();
    if (!raw) return;
    const periodOnly = raw.match(/^(Details of (?:Festival|Casual) Leave)_Period$/i);
    if (periodOnly) {
      const fromKey = `${periodOnly[1]}_Period_From`;
      const toKey = `${periodOnly[1]}_Period_To`;
      if (!seen.has(fromKey)) {
        out.push(fromKey);
        seen.add(fromKey);
      }
      if (!seen.has(toKey)) {
        out.push(toKey);
        seen.add(toKey);
      }
      return;
    }
    if (!seen.has(raw)) {
      out.push(raw);
      seen.add(raw);
    }
  });
  return out;
}

function ensureFormNGJCanonicalSectionHeaders(headers) {
  const list = upgradeFormNGJPeriodHeaders(Array.isArray(headers) ? headers : []);
  const hasPrefix = (prefix) => list.some((h) => String(h || '').startsWith(prefix));
  const out = [...list];
  if (!hasPrefix(`${FORM_NGJ_FESTIVAL_PARENT}_`)) {
    FORM_NGJ_FESTIVAL_HEADERS.forEach((h) => {
      if (!out.includes(h)) out.push(h);
    });
  }
  if (!hasPrefix(`${FORM_NGJ_CASUAL_PARENT}_`)) {
    FORM_NGJ_CASUAL_HEADERS.forEach((h) => {
      if (!out.includes(h)) out.push(h);
    });
  }
  return out;
}

export function resolveFormNGJGujaratTableHeaders(tableHeaders) {
  const normalizeFlat = (header) => {
    const raw = String(header || '').trim();
    if (!raw) return raw;
    // Keep Festival/Casual Period_From / Period_To (do not collapse to parent_From).
    const periodLeaf = raw.match(
      /^(Details of (?:Festival|Casual) Leave)_(Period)_(From|To)$/i
    );
    if (periodLeaf) {
      const parent = normalizeFormNGJParentLabel(periodLeaf[1]) || periodLeaf[1].trim();
      const fromTo = /^to$/i.test(periodLeaf[3]) ? 'To' : 'From';
      return `${parent}_Period_${fromTo}`;
    }
    const m = raw.match(/^(.+)_([\s\S]+)$/);
    if (!m) return raw;
    const parent = normalizeFormNGJParentLabel(m[1]) || m[1].trim();
    const sub = normalizeFormNGJSubLabel(m[2]) || m[2].trim();
    if (parent === 'Remarks' || (parent === 'Remarks' && !sub)) return 'Remarks';
    return `${parent}_${sub}`;
  };

  if (headersLookLikeBrokenFormNGJ(tableHeaders)) {
    return [...FORM_NGJ_GJ_CANONICAL_TABLE_HEADERS];
  }
  const parsed = Array.isArray(tableHeaders)
    ? ensureFormNGJCanonicalSectionHeaders(
        tableHeaders.map((h) => normalizeFlat(String(h || '').trim())).filter(Boolean)
      )
    : [];
  if (parsed.length >= 8) return parsed;
  return [...FORM_NGJ_GJ_CANONICAL_TABLE_HEADERS];
}

export function buildFormNGJSubColumnsFromHeaders(headers) {
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

function normalizeFormNGJParentLabel(raw) {
  const n = formNGJGujaratHeaderNorm(raw);
  if (!n) return '';
  if (/accumulation\s+of\s+leave/.test(n)) return FORM_NGJ_LEAVE_PARENT.accumulation;
  if (/leave\s+allowed/.test(n)) return FORM_NGJ_LEAVE_PARENT.leaveAllowed;
  if (/payment\s+for\s+leave\s+made\s+on/.test(n)) return FORM_NGJ_LEAVE_PARENT.paymentLeave;
  if (/refusal\s+of\s+leave/.test(n)) return FORM_NGJ_LEAVE_PARENT.refusal;
  if (/payment\s+for\s+leave\s+on\s+discharge/.test(n)) return FORM_NGJ_LEAVE_PARENT.paymentDischarge;
  if (/^remarks$/.test(n)) return 'Remarks';
  if (/details\s+of\s+festival/.test(n)) return FORM_NGJ_FESTIVAL_PARENT;
  if (/details\s+of\s+casual/.test(n)) return FORM_NGJ_CASUAL_PARENT;
  return String(raw || '')
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function normalizeFormNGJSubLabel(raw) {
  const n = formNGJGujaratHeaderNorm(raw);
  if (!n) return '';
  if (/leave\s+due\s+on/.test(n)) return 'Leave due on';
  if (/no\.?\s*of\s+days/.test(n)) return 'No. of days';
  if (/from\s+to/.test(n) || (/\bfrom\b/.test(n) && /\bto\b/.test(n))) return 'From To';
  if (/1st\s+moiety/.test(n)) return '1st Moiety';
  if (/2nd\s+moiety/.test(n)) return '2nd Moiety';
  if (/application\s+date/.test(n)) return 'Application Date';
  if (/date\s+of\s+refusal/.test(n)) return 'Date of Refusal';
  if (/date\s+of\s+discharge/.test(n)) return 'Date of discharge';
  if (/date\s+and\s+amount\s+paid/.test(n)) return 'Date and amount paid';
  if (/signature\s+or\s+thumb/.test(n)) return 'Signature or thumb impression of worker';
  if (/^remarks$/.test(n)) return 'Remarks';
  if (/^from$/.test(n)) return 'From';
  if (/^to$/.test(n)) return 'To';
  if (/^period$/.test(n)) return 'Period';
  if (/total\s+leave/.test(n)) return 'Total Leave';
  if (/availed\s+leave/.test(n)) return 'Availed Leave';
  if (/balance\s+leave/.test(n)) return 'Balance Leave';
  if (/payment\s+made\s+in\s+lieu/.test(n)) {
    return 'Payment made in lieu of Festival Leave, when called';
  }
  return String(raw || '')
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

const isFormNGJHeaderMetaLabel = (txt) => {
  const t = formNGJGujaratHeaderNorm(txt);
  return (
    /name\s+of\s+the\s+establishment/.test(t) ||
    (/name\s+of\s+the\s+worker/.test(t) && !/leave\s+due/.test(t)) ||
    /description\s+of\s+the\s+department/.test(t) ||
    (/name\s+of\s+the\s+employer/.test(t) && !/leave\s+allowed/.test(t)) ||
    /date\s+of\s+entry\s+into\s+service/.test(t) ||
    /receipt\s+of\s+level\s+book/.test(t) ||
    /signature\s+or\s+thumb\s+impression\s+of\s+worker/.test(t) ||
    /^form\s*[-–]?\s*n\b/.test(t) ||
    /^leave\s+book$/.test(t) ||
    /^see\s+rule\s+17/.test(t)
  );
};

const isNumericOnly = (txt) => /^\d+$/.test(String(txt || '').trim());

function appendFormNGJSectionFlatHeaders(flatHeaders, subColumnsData, sectionParent, mainLabel, subs) {
  if (!sectionParent || !mainLabel) return;
  const parentKey = `${sectionParent}_${mainLabel}`;
  if (!subColumnsData[sectionParent]) subColumnsData[sectionParent] = [];

  if (mainLabel === 'Period' && subs.length >= 2) {
    subs.forEach((sub) => {
      const key = `${parentKey}_${sub}`;
      if (!flatHeaders.includes(key)) {
        flatHeaders.push(key);
        subColumnsData[sectionParent].push(`${mainLabel}_${sub}`);
      }
    });
    return;
  }

  if (subs.length === 0 || (subs.length === 1 && subs[0] === mainLabel)) {
    const key = parentKey;
    if (!flatHeaders.includes(key)) {
      flatHeaders.push(key);
      subColumnsData[sectionParent].push(mainLabel);
    }
    return;
  }

  subs.forEach((sub) => {
    const key = `${parentKey}_${sub}`;
    if (!flatHeaders.includes(key)) {
      flatHeaders.push(key);
      subColumnsData[sectionParent].push(`${mainLabel}_${sub}`);
    }
  });
}

/**
 * Parse Festival / Casual leave band from Excel (Period From/To + leave summary columns).
 */
function parseFormNGJSectionLeaveHeadersFromSheet({
  jsonData,
  getMergedAwareCellText,
  merges,
  getMergeSpanEndCol,
  sectionLabelRe,
  sectionParent,
  searchStartRow,
  maxCols,
  norm,
  normLower,
}) {
  const flatHeaders = [];
  const sectionSubs = {};
  if (!sectionSubs[sectionParent]) sectionSubs[sectionParent] = [];

  let headerRowIndex = -1;
  for (let r = Math.max(0, searchStartRow); r < Math.min(searchStartRow + 15, jsonData.length); r += 1) {
    const parts = [];
    for (let col = 0; col < maxCols; col += 1) {
      const t = normLower(getMergedAwareCellText(r, col));
      if (t) parts.push(t);
    }
    const joined = parts.join(' ');
    if (sectionLabelRe.test(joined)) {
      headerRowIndex = r + 1;
      break;
    }
    if (/^period$/.test(parts[0] || '') && /total\s+leave/.test(joined)) {
      headerRowIndex = r;
      break;
    }
  }
  if (headerRowIndex < 0 || headerRowIndex >= jsonData.length) {
    return { flatHeaders, subColumnsData: sectionSubs };
  }

  let subRowIndex = headerRowIndex + 1;
  for (let sr = headerRowIndex + 1; sr <= headerRowIndex + 3; sr += 1) {
    const parts = [];
    for (let col = 0; col < maxCols; col += 1) {
      const t = normLower(getMergedAwareCellText(sr, col));
      if (t) parts.push(t);
    }
    const joined = parts.join(' ');
    if (/\bfrom\b/.test(joined) && /\bto\b/.test(joined)) {
      subRowIndex = sr;
      break;
    }
  }

  let c = 0;
  while (c < maxCols) {
    const mergeEnd = getMergeSpanEndCol(headerRowIndex, c);
    let endC = mergeEnd;
    let main = norm(getMergedAwareCellText(headerRowIndex, c));
    if (!main) {
      c = Math.max(c + 1, endC);
      continue;
    }
    while (endC < maxCols) {
      const nextMain = norm(getMergedAwareCellText(headerRowIndex, endC));
      if (!nextMain) break;
      if (normLower(nextMain) !== normLower(main)) break;
      endC = Math.max(endC, getMergeSpanEndCol(headerRowIndex, endC));
    }

    const mainLabel = normalizeFormNGJSubLabel(main) || main;
    if (!mainLabel || sectionLabelRe.test(normLower(mainLabel))) {
      c = endC;
      continue;
    }

    const subs = [];
    for (let sc = c; sc < endC; sc += 1) {
      const subRaw = norm(getMergedAwareCellText(subRowIndex, sc));
      if (!subRaw || isNumericOnly(subRaw)) continue;
      const subLower = normLower(subRaw);
      if (subLower === normLower(main) || subLower === normLower(mainLabel)) continue;
      if (subRaw.endsWith(':') && subRaw.length < 40) continue;
      const sub = normalizeFormNGJSubLabel(subRaw);
      if (sub && !subs.includes(sub)) subs.push(sub);
    }

    if (/^period$/i.test(mainLabel) && subs.length === 0) {
      subs.push('From', 'To');
    }

    appendFormNGJSectionFlatHeaders(flatHeaders, sectionSubs, sectionParent, mainLabel, subs);
    c = endC;
  }

  return { flatHeaders, subColumnsData: sectionSubs };
}

/**
 * Parse Form N GJ Excel — multi-row leave table + festival + casual leave sections.
 */
export function rebuildFormNGJGujaratTableHeadersFromSheet({
  jsonData,
  effectiveSheetCols,
  getMergedAwareCellText,
  merges,
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

  let parentRowIndex = -1;
  for (let r = 0; r < Math.min(40, jsonData.length); r += 1) {
    const parts = [];
    for (let c = 0; c < Math.max(effectiveSheetCols, 13); c += 1) {
      const t = normLower(getMergedAwareCellText(r, c));
      if (t) parts.push(t);
    }
    const joined = parts.join(' ');
    if (/accumulation\s+of\s+leave/.test(joined) && /leave\s+allowed/.test(joined)) {
      parentRowIndex = r;
      break;
    }
  }
  if (parentRowIndex < 0) return null;

  let subRowIndex = parentRowIndex + 1;
  for (let sr = parentRowIndex + 1; sr <= parentRowIndex + 4; sr += 1) {
    const parts = [];
    for (let c = 0; c < Math.max(effectiveSheetCols, 13); c += 1) {
      const t = normLower(getMergedAwareCellText(sr, c));
      if (t) parts.push(t);
    }
    const joined = parts.join(' ');
    if (/leave\s+due\s+on/.test(joined) || /no\.?\s*of\s+days/.test(joined)) {
      subRowIndex = sr;
      break;
    }
  }

  const flatLeaveHeaders = [];
  const subColumnsData = {};
  const mainHeaders = [];
  let startCol = 0;
  const maxCols = Math.max(effectiveSheetCols, (jsonData[parentRowIndex] || []).length, 13);

  let c = 0;
  while (c < maxCols) {
    const mergeEnd = getMergeSpanEndCol(parentRowIndex, c);
    let endC = mergeEnd;
    let main = norm(getMergedAwareCellText(parentRowIndex, c));
    if (!main || isFormNGJHeaderMetaLabel(main)) {
      c = Math.max(c + 1, endC);
      continue;
    }
    while (endC < maxCols) {
      const nextMain = norm(getMergedAwareCellText(parentRowIndex, endC));
      if (!nextMain) break;
      if (normLower(nextMain) !== normLower(main)) break;
      endC = Math.max(endC, getMergeSpanEndCol(parentRowIndex, endC));
    }

    const parent = normalizeFormNGJParentLabel(main);
    if (!parent) {
      c = endC;
      continue;
    }

    const subs = [];
    for (let sc = c; sc < endC; sc += 1) {
      const subRaw = norm(getMergedAwareCellText(subRowIndex, sc));
      if (!subRaw || isNumericOnly(subRaw)) continue;
      const subLower = normLower(subRaw);
      if (subLower === normLower(parent)) continue;
      if (subRaw.endsWith(':') && subRaw.length < 40) continue;
      const sub = normalizeFormNGJSubLabel(subRaw);
      if (sub && !subs.includes(sub)) subs.push(sub);
    }

    mainHeaders.push(parent);
    if (!subColumnsData[parent]) subColumnsData[parent] = [];

    if (parent === 'Remarks' && subs.length === 0) {
      flatLeaveHeaders.push('Remarks');
    } else if (subs.length === 0) {
      flatLeaveHeaders.push(parent);
    } else {
      subs.forEach((sub) => {
        flatLeaveHeaders.push(`${parent}_${sub}`);
        subColumnsData[parent].push(sub);
      });
    }
    c = endC;
  }

  const festivalParsed = parseFormNGJSectionLeaveHeadersFromSheet({
    jsonData,
    getMergedAwareCellText,
    merges,
    getMergeSpanEndCol,
    sectionLabelRe: /details\s+of\s+festival\s+leave/,
    sectionParent: FORM_NGJ_FESTIVAL_PARENT,
    searchStartRow: parentRowIndex + 1,
    maxCols,
    norm,
    normLower,
  });

  let casualSearchStart = parentRowIndex + 1;
  for (let r = parentRowIndex + 1; r < Math.min(parentRowIndex + 25, jsonData.length); r += 1) {
    const parts = [];
    for (let col = 0; col < maxCols; col += 1) {
      const t = normLower(getMergedAwareCellText(r, col));
      if (t) parts.push(t);
    }
    if (/details\s+of\s+casual\s+leave/.test(parts.join(' '))) {
      casualSearchStart = r;
      break;
    }
  }

  const casualParsed = parseFormNGJSectionLeaveHeadersFromSheet({
    jsonData,
    getMergedAwareCellText,
    merges,
    getMergeSpanEndCol,
    sectionLabelRe: /details\s+of\s+casual\s+leave/,
    sectionParent: FORM_NGJ_CASUAL_PARENT,
    searchStartRow: casualSearchStart,
    maxCols,
    norm,
    normLower,
  });

  const leaveHeaders =
    flatLeaveHeaders.length >= 5 ? flatLeaveHeaders : [...FORM_NGJ_LEAVE_TABLE_HEADERS];
  const festivalHeaders =
    festivalParsed.flatHeaders.length >= 3
      ? festivalParsed.flatHeaders
      : [...FORM_NGJ_FESTIVAL_HEADERS];
  const casualHeaders =
    casualParsed.flatHeaders.length >= 3 ? casualParsed.flatHeaders : [...FORM_NGJ_CASUAL_HEADERS];

  Object.assign(subColumnsData, festivalParsed.subColumnsData, casualParsed.subColumnsData);

  const expandedHeaders = [
    ...FORM_NGJ_WORKER_HEADERS,
    ...leaveHeaders,
    ...festivalHeaders,
    ...casualHeaders,
  ];

  return {
    headers: mainHeaders,
    expandedHeaders,
    subColumnsData,
    headerRowIndex: parentRowIndex,
    startCol,
    dataStartIndex: Math.max(subRowIndex + 1, parentRowIndex + 3),
  };
}

export function remapFormNGJGujaratRowsToHeaders(rows, sourceHeaders, targetHeaders) {
  const src = Array.isArray(sourceHeaders) ? sourceHeaders : [];
  const tgt = resolveFormNGJGujaratTableHeaders(targetHeaders);
  if (!Array.isArray(rows)) return [];

  // Use export buckets so Festival Total/Availed/Balance never collide with Casual.
  const bucketFor = (header) => formNGJExportHeaderBucket(header);

  return rows.map((row) => {
    if (!row || typeof row !== 'object') return {};
    const out = {};
    tgt.forEach((targetHeader, colIdx) => {
      let val = '';
      if (Object.prototype.hasOwnProperty.call(row, targetHeader)) {
        val = row[targetHeader];
      } else {
        const bucket = bucketFor(targetHeader);
        for (const [k, v] of Object.entries(row)) {
          if (String(k).startsWith('__')) continue;
          if (bucketFor(k) === bucket && v != null && String(v).trim() !== '') {
            val = v;
            break;
          }
        }
        // Index fallback only when source/target buckets match (never Festival→Casual).
        if (
          (val == null || String(val).trim() === '') &&
          src[colIdx] &&
          Object.prototype.hasOwnProperty.call(row, src[colIdx]) &&
          bucketFor(src[colIdx]) === bucket
        ) {
          val = row[src[colIdx]];
        }
      }
      out[targetHeader] = val == null ? '' : val;
    });
    if (row && typeof row === 'object') {
      Object.entries(row).forEach(([key, value]) => {
        if (String(key).startsWith('__') && out[key] == null) out[key] = value;
      });
    }
    return out;
  });
}

export function enrichFormNGJGujaratDisplayHeader(formHeader, fileName, rowItem, tableHeaders) {
  const ctx = isFormNGJGujaratContext(formHeader, rowItem, fileName, '', tableHeaders);
  if (!ctx) return formHeader;
  return {
    ...(formHeader || {}),
    title: formHeader?.title || 'FORM - N',
    subtitle: formHeader?.subtitle || 'LEAVE BOOK',
    reference: formHeader?.reference || '(See rule 17)',
    formNGJGujaratHeaderFieldLayout: true,
    formNGJGujaratTableLayout: true,
  };
}

export const FORM_NGJ_TEMPLATE_SPECS = [
  {
    key: 'form_n_gj_establishment',
    label: 'Name of the establishment',
    group: 'header',
    fieldType: 'textarea',
    match: /name\s+of\s+the\s+establishment/i,
  },
];

function buildTemplateField(spec, coords = {}) {
  return {
    key: spec.key,
    label: spec.label,
    group: spec.group || 'header',
    fieldType: spec.fieldType || 'text',
    labelRow: coords.labelRow,
    labelCol: coords.labelCol,
    valueCol: coords.valueCol,
    valueRow: coords.valueRow,
    value: coords.value || '',
  };
}

function buildWorkbookMergedCellAccessor(workbook, hints = {}) {
  if (!workbook?.SheetNames?.length) return null;
  const sheetName = hints.preferredSheetName || workbook.SheetNames[0];
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
        return getRawCellText(m.s.r, m.s.c);
      }
    }
    return '';
  };
  return { getMergedAwareCellText, effectiveSheetCols, sheetName };
}

function findHeaderLabelCell(getMergedAwareCellText, matchRe, effectiveSheetCols, maxRows = 80) {
  const maxC = Math.max(20, effectiveSheetCols || 0);
  for (let r = 0; r < maxRows; r += 1) {
    for (let c = 0; c < maxC; c += 1) {
      const raw = String(getMergedAwareCellText(r, c) || '').trim();
      if (!raw) continue;
      const norm = formNGJGujaratHeaderNorm(raw);
      if (!matchRe.test(norm) && !matchRe.test(raw)) continue;
      return { labelRow: r, labelCol: c, valueCol: c + 1, valueRow: r, value: '' };
    }
  }
  return null;
}

export function buildFormNGJTemplateFields(getMergedAwareCellText = null, effectiveSheetCols = 20) {
  return FORM_NGJ_TEMPLATE_SPECS.map((spec) => {
    let coords = {};
    if (typeof getMergedAwareCellText === 'function' && spec.match) {
      coords = findHeaderLabelCell(getMergedAwareCellText, spec.match, effectiveSheetCols) || {};
    }
    return buildTemplateField(spec, coords);
  });
}

export function resolveFormNGJGujaratHeaderFieldLayout(parsed, workbook, hints = {}) {
  const formHeader = hints.formHeader || parsed?.formHeader || null;
  const item = hints.item || null;
  const fileName = hints.fileName || hints.formFileName || '';
  const sheetText = hints.sheetText || '';
  const tableHeaders = hints.tableHeaders || parsed?.headers || [];
  if (
    !isFormNGJGujaratContext(formHeader, item, fileName, sheetText, tableHeaders) &&
    !headersLookLikeBrokenFormNGJ(tableHeaders)
  ) {
    return null;
  }

  const accessor = buildWorkbookMergedCellAccessor(workbook, hints);
  const getMergedAwareCellText = accessor?.getMergedAwareCellText ?? null;
  const effectiveSheetCols = accessor?.effectiveSheetCols ?? 20;
  const uiFields = buildFormNGJTemplateFields(getMergedAwareCellText, effectiveSheetCols);

  // Keep real sheet table anchors — draft export requires headerRowIndex >= 0.
  let headers = resolveFormNGJGujaratTableHeaders(parsed?.headers || hints.tableHeaders || []);
  let headerRowIndex = Number(parsed?.headerRowIndex);
  let dataStartIndex = Number(parsed?.dataStartIndex);
  let tableStartCol = Number(parsed?.tableStartCol);
  if (!Number.isFinite(headerRowIndex)) headerRowIndex = -1;
  if (!Number.isFinite(dataStartIndex)) dataStartIndex = -1;
  if (!Number.isFinite(tableStartCol) || tableStartCol < 0) tableStartCol = 0;

  if (workbook && accessor && (headerRowIndex < 0 || dataStartIndex < 0 || headers.length < 8)) {
    try {
      const sheetName = accessor.sheetName;
      const worksheet = sheetName ? workbook.Sheets?.[sheetName] : null;
      if (worksheet) {
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
        const rebuild = rebuildFormNGJGujaratTableHeadersFromSheet({
          jsonData,
          effectiveSheetCols,
          getMergedAwareCellText,
          merges: worksheet['!merges'] || [],
        });
        if (rebuild?.expandedHeaders?.length >= 8) {
          headers = resolveFormNGJGujaratTableHeaders(rebuild.expandedHeaders);
          if (headerRowIndex < 0 && Number.isFinite(Number(rebuild.headerRowIndex))) {
            headerRowIndex = Number(rebuild.headerRowIndex);
          }
          if (dataStartIndex < 0 && Number.isFinite(Number(rebuild.dataStartIndex))) {
            dataStartIndex = Number(rebuild.dataStartIndex);
          }
          if (Number.isFinite(Number(rebuild.startCol))) {
            tableStartCol = Number(rebuild.startCol);
          }
        }
      }
    } catch (_) {
      /* keep parsed anchors */
    }
  }

  if (dataStartIndex < 0 && headerRowIndex >= 0) {
    dataStartIndex = headerRowIndex + 3;
  }

  return {
    formHeader: {
      title: formHeader?.title || 'FORM - N',
      subtitle: formHeader?.subtitle || 'LEAVE BOOK',
      reference: formHeader?.reference || '(See rule 17)',
      formNGJGujaratHeaderFieldLayout: true,
      formNGJGujaratTableLayout: true,
      textRows: [],
      fields: uiFields,
      templateFields: uiFields,
    },
    headers,
    subColumns: buildFormNGJSubColumnsFromHeaders(headers),
    tableData: Array.isArray(parsed?.tableData) ? parsed.tableData : [],
    headerRowIndex,
    dataStartIndex: dataStartIndex >= 0 ? dataStartIndex : 0,
    tableStartCol,
    sheetName: accessor?.sheetName || hints.preferredSheetName || null,
  };
}

function unwrapEmployeeRecord(emp) {
  return emp?.Employee || emp?.employee || emp;
}

function readEmployeeFullName(emp) {
  const src = unwrapEmployeeRecord(emp);
  const fn = String(
    src?.FirstName || src?.['FirstName'] || src?.firstName || src?.['First Name'] || ''
  ).trim();
  const ln = String(
    src?.LastName || src?.['LastName'] || src?.lastName || src?.['Last Name'] || ''
  ).trim();
  if (fn && ln) return `${fn} ${ln}`;
  return (
    fn ||
    ln ||
    String(
      src?.DisplayName ||
        src?.['Display Name'] ||
        src?.Employee_Name ||
        src?.['Employee Name'] ||
        ''
    ).trim()
  );
}

function readEmployeeDepartment(emp) {
  const src = unwrapEmployeeRecord(emp);
  return String(
    src?.Department ||
      src?.['Department'] ||
      src?.department ||
      src?.Designation ||
      src?.['Designation'] ||
      src?.designation ||
      ''
  ).trim();
}

function readEmployeeDateOfJoining(emp) {
  const src = unwrapEmployeeRecord(emp);
  return String(
    src?.Dateofjoining ||
      src?.['Dateofjoining'] ||
      src?.DateOfJoining ||
      src?.['Date Of Joining'] ||
      src?.dateOfJoining ||
      src?.['Date of Joining'] ||
      ''
  ).trim();
}

/** Prefer EmployeeID (VE0471) so LeaveData "Name (VE0471)" matches; Zoho_ID as fallback. */
function readEmployeeLookupId(emp) {
  const src = unwrapEmployeeRecord(emp);
  if (!src || typeof src !== 'object') return '';
  const keys = [
    'EmployeeID',
    'Employee ID',
    'Employee_ID',
    'EmployeeId',
    'employeeId',
    'employee_id',
    'Employee.ID',
    'Employee_Code',
    'EmployeeCode',
    'Zoho_ID',
    'ZohoID',
    'zoho_id',
    'erecno',
    'Erecno',
  ];
  for (let i = 0; i < keys.length; i += 1) {
    const raw = src[keys[i]];
    if (raw == null || raw === '') continue;
    if (typeof raw === 'object') {
      const nested = String(raw.displayValue ?? raw.name ?? raw.Name ?? raw.ID ?? '').trim();
      if (nested) return nested;
      continue;
    }
    const text = String(raw).trim();
    if (text) return text;
  }
  return '';
}

function collectEmployeeLookupIdCandidates(emp, options = {}) {
  const ids = new Set();
  const add = (v) => {
    const s = String(v || '')
      .trim()
      .toLowerCase();
    if (s && s.length >= 2) ids.add(s);
  };
  add(readEmployeeLookupId(emp));
  const src = unwrapEmployeeRecord(emp);
  if (src && typeof src === 'object') {
    [
      'EmployeeID',
      'Employee ID',
      'Employee_ID',
      'EmployeeId',
      'employeeId',
      'Zoho_ID',
      'ZohoID',
      'zoho_id',
      'erecno',
    ].forEach((k) => add(src[k]));
  }
  if (typeof options.collectEmployeeIdCandidates === 'function' && emp) {
    options.collectEmployeeIdCandidates(emp, {}).forEach(add);
  }
  return [...ids];
}

export function isFormNGJWorkerNameHeader(header) {
  return /name\s+of\s+the\s+worker/i.test(normHeaderLabel(header));
}

export function isFormNGJDepartmentHeader(header) {
  return /description\s+of\s+the\s+department/i.test(normHeaderLabel(header));
}

export function isFormNGJEmployerHeader(header) {
  return /name\s+of\s+the\s+employer/i.test(normHeaderLabel(header));
}

export function isFormNGJDateOfEntryHeader(header) {
  return /date\s+of\s+entry\s+into\s+service/i.test(normHeaderLabel(header));
}

export function isFormNGJReceiptHeader(header) {
  return /receipt\s+of\s+level\s+book/i.test(normHeaderLabel(header));
}

/** Leave / festival columns — people overlay skips these (Casual Leave filled from Contingency API). */
export function isFormNGJSkipPeopleAutofillHeader(header) {
  const n = normHeaderLabel(header);
  if (isFormNGJWorkerNameHeader(header)) return false;
  if (isFormNGJDepartmentHeader(header)) return false;
  if (isFormNGJEmployerHeader(header)) return false;
  if (isFormNGJDateOfEntryHeader(header)) return false;
  if (isFormNGJReceiptHeader(header)) return false;
  return (
    /leave/.test(n) ||
    /moiety/.test(n) ||
    /refusal/.test(n) ||
    /discharge/.test(n) ||
    /festival/.test(n) ||
    /casual/.test(n) ||
    /^remarks$/.test(n) ||
    /period/.test(n) ||
    /availed/.test(n) ||
    /balance/.test(n)
  );
}

export function formatFormNGJGujaratEstablishmentFromSite(site) {
  if (!site || typeof site !== 'object') return '';
  const name = String(site.siteName ?? site.SiteName ?? '').trim();
  const addr = String(site.siteAddress ?? site.SiteAddress ?? '').trim();
  const city = String(site.siteCity ?? site.SiteCity ?? '').trim();
  const state = String(site.siteState ?? site.SiteState ?? '').trim();
  const addrLine = [addr, city, state].filter(Boolean).join(', ');
  return [name, addrLine].filter(Boolean).join('\n').trim();
}

export function formatFormNGJGujaratEmployerFromSite(site) {
  if (!site || typeof site !== 'object') return '';
  return String(
    site.employerName ??
      site.EmployerName ??
      site.companyName ??
      site.CompanyName ??
      site.siteName ??
      site.SiteName ??
      ''
  ).trim();
}

export function applyFormNGJGujaratAutofillFromSite(headerData, siteContext = {}, options = {}) {
  const { onlyEmpty = false } = options;
  const out = headerData && typeof headerData === 'object' ? { ...headerData } : {};
  const establishmentText = String(siteContext.establishmentText || '').trim();
  if (establishmentText) {
    const cur = String(out.form_n_gj_establishment ?? '').trim();
    if (!onlyEmpty || !cur || /^enter\b/i.test(cur)) {
      out.form_n_gj_establishment = establishmentText;
    }
  }
  return out;
}

export function applyFormNGJGujaratEmployeeToRow(row, emp, headers, helpers = {}) {
  const hdrs = resolveFormNGJGujaratTableHeaders(headers);
  const out = row && typeof row === 'object' ? { ...row } : {};
  const {
    sanitizeValue = (v) => String(v ?? '').trim(),
    overwrite = true,
    formatStatutoryDateDisplay = (v) => String(v || '').trim(),
    employerText = '',
  } = helpers;

  const shouldSet = (header) => overwrite || !String(out[header] ?? '').trim();
  const setCell = (header, value) => {
    if (!header || !shouldSet(header)) return;
    out[header] = sanitizeValue(value ?? '');
  };

  const fullName = readEmployeeFullName(emp);
  const employeeId = readEmployeeLookupId(emp);
  if (fullName) out.__employeeLookupName = fullName;
  if (employeeId) out.__employeeLookupId = employeeId;

  hdrs.forEach((header) => {
    if (isFormNGJSkipPeopleAutofillHeader(header)) {
      setCell(header, '');
      return;
    }
    if (isFormNGJWorkerNameHeader(header)) {
      setCell(header, fullName);
      return;
    }
    if (isFormNGJDepartmentHeader(header)) {
      setCell(header, readEmployeeDepartment(emp));
      return;
    }
    if (isFormNGJEmployerHeader(header)) {
      setCell(header, employerText);
      return;
    }
    if (isFormNGJDateOfEntryHeader(header)) {
      const doj = readEmployeeDateOfJoining(emp);
      setCell(header, doj ? formatStatutoryDateDisplay(doj) : '');
      return;
    }
    if (isFormNGJReceiptHeader(header)) {
      setCell(header, '');
    }
  });
  return out;
}

/** Casual Leave on Form N maps to Zoho Contingency Leave (same aliases as Form X Medical). */
export const FORM_NGJ_CASUAL_LEAVE_TYPE_ALIASES = [...FORM_X_MEDICAL_LEAVE_TYPE_ALIASES];

export function isFormNGJCasualSectionHeader(header) {
  return /details\s+of\s+casual\s+leave/i.test(String(header || ''));
}

export function isFormNGJCasualPeriodFromHeader(header) {
  if (!isFormNGJCasualSectionHeader(header)) return false;
  const sub = String(header || '')
    .split('_')
    .pop();
  if (/^from$/i.test(String(sub || '').trim())) return true;
  const n = normHeaderLabel(header);
  return /\bfrom\b/.test(n) && !/\bto\b/.test(n);
}

export function isFormNGJCasualPeriodToHeader(header) {
  if (!isFormNGJCasualSectionHeader(header)) return false;
  const sub = String(header || '')
    .split('_')
    .pop();
  if (/^to$/i.test(String(sub || '').trim())) return true;
  const n = normHeaderLabel(header);
  return /\bto\b/.test(n) && !/\bfrom\b/.test(n);
}

export function isFormNGJCasualTotalLeaveHeader(header) {
  return isFormNGJCasualSectionHeader(header) && /total\s+leave/.test(normHeaderLabel(header));
}

export function isFormNGJCasualAvailedLeaveHeader(header) {
  return isFormNGJCasualSectionHeader(header) && /availed\s+leave/.test(normHeaderLabel(header));
}

export function isFormNGJCasualBalanceLeaveHeader(header) {
  return isFormNGJCasualSectionHeader(header) && /balance\s+leave/.test(normHeaderLabel(header));
}

export function isFormNGJCasualLeaveAutofillHeader(header) {
  return (
    isFormNGJCasualPeriodFromHeader(header) ||
    isFormNGJCasualPeriodToHeader(header) ||
    isFormNGJCasualTotalLeaveHeader(header) ||
    isFormNGJCasualAvailedLeaveHeader(header) ||
    isFormNGJCasualBalanceLeaveHeader(header)
  );
}

export function isApprovedLeaveContingencyForFormNGJ(record) {
  const leaveType = readApprovedLeaveLeaveType(record);
  if (!leaveType) return false;
  return leaveTypeLabelMatchesAliases(leaveType, FORM_NGJ_CASUAL_LEAVE_TYPE_ALIASES);
}

export function filterApprovedLeaveRecordsForFormNGJCasual(records, monthFrom, monthTo) {
  const monthFiltered = filterApprovedLeaveRecordsForFormOGJMonth(records, monthFrom, monthTo);
  return monthFiltered.filter(isApprovedLeaveContingencyForFormNGJ);
}

function formatFormNGJLeaveNumber(value) {
  if (value == null || String(value).trim() === '') return '';
  const n = Number(String(value).replace(/,/g, '').trim());
  if (!Number.isFinite(n)) return '';
  return String(Math.round(n * 100) / 100);
}

function setFormNGJCasualCell(row, header, value, overwrite) {
  if (value == null || String(value).trim() === '') return false;
  if (!overwrite && String(row[header] ?? '').trim() !== '') return false;
  row[header] = String(value).trim();
  return true;
}

/**
 * Read FirstName / LastName for Form N matching.
 * Never match Contingency leave on first name or last name alone.
 */
export function readFormNGJEmployeeNameParts(emp) {
  const src = unwrapEmployeeRecord(emp);
  if (!src || typeof src !== 'object') return { first: '', last: '', full: '' };
  const first = String(
    src.FirstName || src['FirstName'] || src.firstName || src['First Name'] || ''
  ).trim();
  const last = String(
    src.LastName || src['LastName'] || src.lastName || src['Last Name'] || ''
  ).trim();
  const full = first && last ? `${first} ${last}` : first || last || '';
  return { first, last, full };
}

/**
 * Form N name match: require FirstName AND LastName (exact tokens).
 * Does NOT allow last-initial shortcuts ("Ajaykumar M" ↛ "Ajaykumar Mansingbhai"),
 * which previously assigned another worker's 5-day leave to Ajaykumar Mansingbhai.
 */
export function formNGJPersonNamesMatchExact(workerName, recordName, firstName = '', lastName = '') {
  const recordNorm = normalizePersonNameKey(recordName);
  if (!recordNorm) return false;

  const fn = normalizePersonNameKey(firstName);
  const ln = normalizePersonNameKey(lastName);
  if (fn && ln) {
    const recordParts = recordNorm.split(' ').filter(Boolean);
    if (recordParts.length < 2) return false;
    return recordParts[0] === fn && recordParts[recordParts.length - 1] === ln;
  }

  const workerNorm = normalizePersonNameKey(workerName);
  if (!workerNorm) return false;
  if (workerNorm === recordNorm) return true;

  const aParts = workerNorm.split(' ').filter(Boolean);
  const bParts = recordNorm.split(' ').filter(Boolean);
  if (aParts.length < 2 || bParts.length < 2) return false;
  return aParts[0] === bParts[0] && aParts[aParts.length - 1] === bParts[bParts.length - 1];
}

/** Only EmployeeID / Zoho People ids — never Role.ID or other shared codes. */
function collectFormNGJApprovedLeaveMatchIds(emp, options = {}) {
  const ids = new Set();
  const add = (value) => {
    const id = String(value || '')
      .trim()
      .toLowerCase();
    if (!id || id.length < 4 || /^\d{1,3}$/.test(id)) return;
    ids.add(id);
  };

  const src = unwrapEmployeeRecord(emp);
  if (src && typeof src === 'object') {
    add(readEmployeeLookupId(src));
    [
      'EmployeeID',
      'Employee ID',
      'Employee_ID',
      'EmployeeId',
      'employeeId',
      'Zoho_ID',
      'ZohoID',
      'zoho_id',
      'erecno',
      'Erecno',
    ].forEach((k) => add(src[k]));
  }

  // Narrow optional collector: keep VE/Zoho-like ids only (drop Role.ID etc.).
  if (typeof options.collectEmployeeIdCandidates === 'function' && emp) {
    options.collectEmployeeIdCandidates(emp, {}).forEach((raw) => {
      const id = String(raw || '')
        .trim()
        .toLowerCase();
      if (!id || id.length < 4) return;
      if (/^ve\d+/i.test(id) || /^[a-z]{1,3}\d{3,}$/i.test(id) || /^\d{6,}$/.test(id)) {
        ids.add(id);
      }
    });
  }
  return [...ids];
}

/**
 * Match approved Contingency leave to a Form N worker using EmployeeID/Zoho id
 * and FirstName+LastName — never first-name-only or last-initial-only.
 *
 * Zoho approved-leave rows can share the same erecno across different people
 * (seen in Approved Leaves). An ID hit alone is not enough when FirstName+LastName
 * are available and the leave row has a name.
 */
export function formNGJRecordMatchesWorker(record, workerName, emp, options = {}) {
  if (!record) return false;

  const { ids: recordIds, names: recordNames } = collectApprovedLeaveIdentityKeys(record);
  const workerIds = new Set(collectFormNGJApprovedLeaveMatchIds(emp, options));
  const { first, last, full } = readFormNGJEmployeeNameParts(emp);
  const nameForMatch = full || String(workerName || '').trim();

  const nameMatched = recordNames.some((rn) =>
    formNGJPersonNamesMatchExact(nameForMatch, rn, first, last)
  );
  const idMatched = workerIds.size > 0 && recordIds.some((id) => workerIds.has(id));

  if (idMatched) {
    // Duplicate Zoho IDs: reject when emp FirstName+LastName disagree with leave name.
    if (first && last && recordNames.length > 0 && !nameMatched) return false;
    return true;
  }

  if (!nameForMatch && !(first && last)) return false;
  return nameMatched;
}

/**
 * Among Contingency leaves for this worker, prefer LeaveCount that matches
 * LeaveData Contingency booked (avoids another person's leave via shared Zoho ID).
 */
export function pickBestFormNGJContingencyLeave(matches, bookedLimit = '') {
  if (!Array.isArray(matches) || matches.length === 0) return null;
  if (matches.length === 1) return matches[0];

  const bookedNum = Number(String(bookedLimit || '').replace(/,/g, '').trim());
  const scored = matches.map((rec) => {
    const period = resolveApprovedLeavePeriodForFormNGJ(rec, bookedLimit);
    const count = Number(String(period.daysCount || '').replace(/,/g, '').trim());
    return { rec, count: Number.isFinite(count) ? count : NaN, period };
  });

  if (Number.isFinite(bookedNum) && bookedNum > 0) {
    const exact = scored.find((s) => s.count === bookedNum);
    if (exact) return exact.rec;
    const under = scored
      .filter((s) => Number.isFinite(s.count) && s.count > 0 && s.count <= bookedNum)
      .sort((a, b) => b.count - a.count);
    if (under.length > 0) return under[0].rec;
  }
  return matches[0];
}

/**
 * Find Contingency Leave approved record for a Form N worker row.
 * Matches on EmployeeID / Zoho id and FirstName+LastName (not first/last alone).
 */
export function findApprovedContingencyLeaveForFormNRow(
  row,
  tableHeaders,
  approvedRecords,
  emp,
  options = {}
) {
  if (!Array.isArray(approvedRecords) || approvedRecords.length === 0) return null;

  const hdrs = resolveFormNGJGujaratTableHeaders(tableHeaders);
  const workerHeader = hdrs.find(isFormNGJWorkerNameHeader) || '';
  const { full: empFullName } = readFormNGJEmployeeNameParts(emp);
  const workerName = workerHeader
    ? String(row?.[workerHeader] ?? row?.__employeeLookupName ?? empFullName ?? '').trim()
    : String(row?.__employeeLookupName ?? empFullName ?? '').trim();
  if (!workerName && !empFullName) return null;

  const usedRecordKeys =
    options.usedRecordKeys instanceof Set ? options.usedRecordKeys : new Set();
  const monthFrom = options.monthFrom || '';
  const monthTo = options.monthTo || '';
  const contingencyRecords = filterApprovedLeaveRecordsForFormNGJCasual(
    approvedRecords,
    monthFrom,
    monthTo
  );

  const matches = [];
  for (let i = 0; i < contingencyRecords.length; i += 1) {
    const rec = contingencyRecords[i];
    if (!rec) continue;
    const key = getApprovedLeaveRecordUniqueKey(rec);
    if (key && usedRecordKeys.has(key)) continue;
    if (formNGJRecordMatchesWorker(rec, workerName || empFullName, emp, options)) {
      matches.push(rec);
    }
  }
  return pickBestFormNGJContingencyLeave(matches, options.bookedLimit || '');
}

/**
 * Build Casual Leave metrics:
 * From / To → Days with LeaveCount > 0 (drop weekly-off 0.0; drop weekends when
 *   LeaveCount exceeds Contingency booked).
 * Availed → that leave's LeaveCount, never above LeaveData Contingency booked.
 * Balance → Contingency LeaveData (end balance).
 * Total → beginning = balance + availed when both exist.
 */
export function buildFormNGJCasualLeaveValues(approvedRecord, leaveRecord, leaveTypeLabels = {}) {
  const apiMetrics = leaveRecord
    ? getFormXMedicalLeaveApiMetrics(leaveRecord, leaveTypeLabels)
    : { balance: '', booked: '', hasData: false };
  const booked =
    apiMetrics.hasData && apiMetrics.booked !== '' && apiMetrics.booked != null
      ? formatFormNGJLeaveNumber(apiMetrics.booked)
      : '';

  const metrics = resolveApprovedLeavePeriodForFormNGJ(approvedRecord || {}, booked);
  const hasApprovedPeriod = !!(metrics.from || metrics.to);
  if (!hasApprovedPeriod) {
    return {
      from: '',
      to: '',
      totalLeave: '',
      availedLeave: '',
      balanceLeave: '',
      hasApprovedPeriod: false,
      hasLeaveMetrics: false,
    };
  }

  let availed = formatFormNGJLeaveNumber(metrics.daysCount);
  // If LeaveCount still exceeds Contingency booked, trust LeaveData booked.
  if (availed !== '' && booked !== '') {
    const leaveCountNum = Number(availed);
    const bookedNum = Number(booked);
    if (Number.isFinite(leaveCountNum) && Number.isFinite(bookedNum) && leaveCountNum > bookedNum) {
      availed = booked;
    }
  }
  const balance =
    apiMetrics.hasData && apiMetrics.balance !== '' && apiMetrics.balance != null
      ? formatFormNGJLeaveNumber(apiMetrics.balance)
      : '';
  let total = '';
  if (availed !== '' && balance !== '') {
    total = computeFormXLeaveBeginning(balance, availed);
  }
  if (!total && availed !== '') total = availed;

  return {
    from: metrics.from || '',
    to: metrics.to || '',
    totalLeave: total || '',
    availedLeave: availed || '',
    balanceLeave: balance || '',
    hasApprovedPeriod: true,
    hasLeaveMetrics: !!(total || availed || balance),
  };
}

export function applyFormNGJGujaratCasualLeaveToRow(
  row,
  approvedRecord,
  leaveRecord,
  tableHeaders,
  { overwrite = true, leaveTypeLabels = {} } = {}
) {
  if (!row) return 0;
  const values = buildFormNGJCasualLeaveValues(approvedRecord, leaveRecord, leaveTypeLabels);
  if (!values.hasApprovedPeriod && !values.hasLeaveMetrics) return 0;

  const hdrs = resolveFormNGJGujaratTableHeaders(tableHeaders);
  let applied = 0;
  hdrs.forEach((header) => {
    if (isFormNGJCasualPeriodFromHeader(header) && values.from) {
      if (setFormNGJCasualCell(row, header, values.from, overwrite)) applied += 1;
    } else if (isFormNGJCasualPeriodToHeader(header) && values.to) {
      if (setFormNGJCasualCell(row, header, values.to, overwrite)) applied += 1;
    } else if (isFormNGJCasualTotalLeaveHeader(header) && values.totalLeave) {
      if (setFormNGJCasualCell(row, header, values.totalLeave, overwrite)) applied += 1;
    } else if (isFormNGJCasualAvailedLeaveHeader(header) && values.availedLeave) {
      if (setFormNGJCasualCell(row, header, values.availedLeave, overwrite)) applied += 1;
    } else if (isFormNGJCasualBalanceLeaveHeader(header) && values.balanceLeave) {
      if (setFormNGJCasualCell(row, header, values.balanceLeave, overwrite)) applied += 1;
    }
  });
  return applied;
}

/**
 * Autofill Form N "Details of Casual Leave" from Contingency approved leave
 * (From/To via Form O match model) + LeaveData Contingency metrics.
 */
export function applyFormNGJGujaratCasualLeaveAutofill(
  mappedData,
  employeesForMapping,
  tableHeaders,
  approvedLeaveRecords,
  leaveRecords = [],
  options = {}
) {
  if (!Array.isArray(mappedData) || mappedData.length === 0) return 0;

  const priorHeaders = Array.isArray(tableHeaders) ? tableHeaders : [];
  const canonicalHeaders = resolveFormNGJGujaratTableHeaders(tableHeaders);
  const monthFrom = options.monthFrom || '';
  const monthTo = options.monthTo || '';
  const leaveTypeLabels = options.leaveTypeLabels || {};
  const contingencyRecords = filterApprovedLeaveRecordsForFormNGJCasual(
    approvedLeaveRecords,
    monthFrom,
    monthTo
  );
  const leaveLookup = buildLeaveRecordLookupMap(leaveRecords);
  const workerHeader = canonicalHeaders.find(isFormNGJWorkerNameHeader) || '';
  const usedRecordKeys = new Set();
  let hits = 0;

  mappedData.forEach((row, index) => {
    const normalizedRow = remapFormNGJGujaratRowsToHeaders(
      [row],
      priorHeaders,
      canonicalHeaders
    )[0];
    Object.assign(row, normalizedRow);

    const empItem = employeesForMapping[index];
    const emp = empItem && (empItem.Employee || empItem.employee || empItem);
    const empName = readEmployeeFullName(emp);
    const empId = readEmployeeLookupId(emp);
    if (empName) row.__employeeLookupName = empName;
    if (empId) row.__employeeLookupId = empId;

    let leaveRecord = null;
    const idCandidates = collectEmployeeLookupIdCandidates(emp, options);
    for (let i = 0; i < idCandidates.length; i += 1) {
      const id = idCandidates[i];
      if (leaveLookup.byId.has(id)) {
        leaveRecord = leaveLookup.byId.get(id);
        break;
      }
    }
    if (!leaveRecord) {
      leaveRecord =
        findLeaveRecordForFormRow(leaveLookup, row, '', workerHeader) ||
        (typeof options.findLeaveRecord === 'function'
          ? options.findLeaveRecord(row, emp, index)
          : null);
    }

    const bookedLimit = leaveRecord
      ? String(
          getFormXMedicalLeaveApiMetrics(leaveRecord, leaveTypeLabels).booked ?? ''
        ).trim()
      : '';

    const approvedRecord = findApprovedContingencyLeaveForFormNRow(
      row,
      canonicalHeaders,
      contingencyRecords.length > 0 ? contingencyRecords : approvedLeaveRecords,
      emp,
      {
        usedRecordKeys,
        monthFrom,
        monthTo,
        bookedLimit,
        collectEmployeeIdCandidates: options.collectEmployeeIdCandidates,
      }
    );

    if (approvedRecord) {
      const recordKey = getApprovedLeaveRecordUniqueKey(approvedRecord);
      if (recordKey) usedRecordKeys.add(recordKey);
    } else {
      // No Contingency leave → clear all casual columns (no LeaveData defaults).
      canonicalHeaders.forEach((header) => {
        if (isFormNGJCasualLeaveAutofillHeader(header)) {
          row[header] = '';
        }
      });
      return;
    }

    const applied = applyFormNGJGujaratCasualLeaveToRow(
      row,
      approvedRecord,
      leaveRecord,
      canonicalHeaders,
      { overwrite: options.overwrite !== false, leaveTypeLabels }
    );
    if (applied > 0) hits += 1;
  });

  return hits;
}

/* -------------------------------------------------------------------------- */
/* ExcelJS download — preserves template merges (SheetJS buildDraft corrupts) */
/* -------------------------------------------------------------------------- */

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

function buildMergeTopLeftResolver(worksheet) {
  const cache = new Map();
  return (r, c) => {
    const key = `${r}:${c}`;
    if (cache.has(key)) return cache.get(key);
    let topLeft = { r, c };
    const merges = worksheet.model?.merges;
    if (Array.isArray(merges)) {
      for (let mi = 0; mi < merges.length; mi += 1) {
        const parts = String(merges[mi] || '').split(':');
        if (parts.length !== 2) continue;
        const tl = worksheet.getCell(parts[0]);
        const br = worksheet.getCell(parts[1]);
        if (!tl || !br) continue;
        const r1 = tl.fullAddress?.row ?? tl.row;
        const c1 = tl.fullAddress?.col ?? tl.col;
        const r2 = br.fullAddress?.row ?? br.row;
        const c2 = br.fullAddress?.col ?? br.col;
        if (r >= r1 && r <= r2 && c >= c1 && c <= c2) {
          topLeft = { r: r1, c: c1 };
          break;
        }
      }
    }
    cache.set(key, topLeft);
    return topLeft;
  };
}

function buildMergeSpanEndResolver(worksheet, getMergeTopLeft) {
  const cache = new Map();
  return (r, c) => {
    const tl = getMergeTopLeft(r, c);
    const key = `${tl.r}:${tl.c}`;
    if (cache.has(key)) return cache.get(key);
    let end = { r: tl.r, c: tl.c };
    const merges = worksheet.model?.merges;
    if (Array.isArray(merges)) {
      for (let mi = 0; mi < merges.length; mi += 1) {
        const parts = String(merges[mi] || '').split(':');
        if (parts.length !== 2) continue;
        const mtl = worksheet.getCell(parts[0]);
        const br = worksheet.getCell(parts[1]);
        if (!mtl || !br) continue;
        const r1 = mtl.fullAddress?.row ?? mtl.row;
        const c1 = mtl.fullAddress?.col ?? mtl.col;
        const r2 = br.fullAddress?.row ?? br.row;
        const c2 = br.fullAddress?.col ?? br.col;
        if (tl.r === r1 && tl.c === c1) {
          end = { r: r2, c: c2 };
          break;
        }
      }
    }
    cache.set(key, end);
    return end;
  };
}

function formNGJExportHeaderBucket(header) {
  const n = normHeaderLabel(header);
  if (/name\s+of\s+the\s+worker/.test(n)) return 'workerName';
  if (/description\s+of\s+the\s+department/.test(n)) return 'department';
  if (/name\s+of\s+the\s+employer/.test(n)) return 'employer';
  if (/date\s+of\s+entry\s+into\s+service/.test(n)) return 'dateOfEntry';
  if (/receipt\s+of\s+le[av]e?\s*book|receipt\s+of\s+level\s+book/.test(n)) return 'receipt';
  if (/leave\s+due\s+on/.test(n)) return 'leaveDueOn';
  if (/no\.?\s*of\s+days/.test(n)) return 'noOfDays';
  if (/from\s*-?\s*to|from\s+to/.test(n) && /leave\s+allowed|allowed/.test(String(header || '').toLowerCase())) {
    return 'leaveFromTo';
  }
  if (/from\s*-?\s*to|^from\s+to$/.test(n)) return 'leaveFromTo';
  if (/1st\s+moiety/.test(n)) return 'moiety1';
  if (/2nd\s+moiety/.test(n)) return 'moiety2';
  if (/application\s+date/.test(n)) return 'applicationDate';
  if (/date\s+of\s+refusal/.test(n)) return 'dateOfRefusal';
  if (/date\s+of\s+discharge/.test(n)) return 'dateOfDischarge';
  if (/date\s+and\s+amount\s+paid/.test(n)) return 'dateAndAmountPaid';
  if (/signature|thumb\s+impression/.test(n)) return 'signature';
  if (/^remarks$/.test(n) && !/festival|casual/.test(String(header || '').toLowerCase())) return 'leaveRemarks';
  if (/details\s+of\s+festival/.test(n) || String(header || '').startsWith(FORM_NGJ_FESTIVAL_PARENT)) {
    if (/period.*from|period_from|_period_from/.test(n) || /_period_from$/i.test(String(header || ''))) {
      return 'festivalPeriodFrom';
    }
    if (/period.*to|period_to|_period_to/.test(n) || /_period_to$/i.test(String(header || ''))) {
      return 'festivalPeriodTo';
    }
    if (/^period$|_period$/i.test(n) || /_period$/i.test(String(header || ''))) return 'festivalPeriod';
    if (/total\s+leave/.test(n)) return 'festivalTotal';
    if (/availed\s+leave/.test(n)) return 'festivalAvailed';
    if (/balance\s+leave/.test(n)) return 'festivalBalance';
    if (/payment\s+made\s+in\s+lieu/.test(n)) return 'festivalPayment';
    if (/remarks/.test(n)) return 'festivalRemarks';
  }
  if (/details\s+of\s+casual/.test(n) || String(header || '').startsWith(FORM_NGJ_CASUAL_PARENT)) {
    if (/period.*from|period_from|_period_from/.test(n) || /_period_from$/i.test(String(header || ''))) {
      return 'casualPeriodFrom';
    }
    if (/period.*to|period_to|_period_to/.test(n) || /_period_to$/i.test(String(header || ''))) {
      return 'casualPeriodTo';
    }
    if (/^period$|_period$/i.test(n) || /_period$/i.test(String(header || ''))) return 'casualPeriod';
    if (/total\s+leave/.test(n)) return 'casualTotal';
    if (/availed\s+leave/.test(n)) return 'casualAvailed';
    if (/balance\s+leave/.test(n)) return 'casualBalance';
    if (/remarks/.test(n)) return 'casualRemarks';
  }
  return n;
}

export function getFormNGJRowValueForHeader(row, header) {
  if (!row || !header) return '';
  const direct = row[header];
  if (direct != null && String(direct).trim() !== '') return String(direct).trim();
  const bucket = formNGJExportHeaderBucket(header);
  for (const [k, v] of Object.entries(row)) {
    if (String(k).startsWith('__')) continue;
    if (formNGJExportHeaderBucket(k) === bucket && v != null && String(v).trim() !== '') {
      return String(v).trim();
    }
  }
  return '';
}

export function rowHasMeaningfulFormNGJGujaratExportData(row, headers) {
  const hdrs = resolveFormNGJGujaratTableHeaders(headers);
  return hdrs.some((header) => getFormNGJRowValueForHeader(row, header) !== '');
}

export function filterFormNGJGujaratExportRows(rows, headers) {
  const hdrs = resolveFormNGJGujaratTableHeaders(headers);
  return (Array.isArray(rows) ? rows : []).filter((row) =>
    rowHasMeaningfulFormNGJGujaratExportData(row, hdrs)
  );
}

function detectFormNGJSectionColumns(worksheet, getMergedAwareCellText, getMergeSpanEnd, titleRe, maxCol) {
  let titleRow = -1;
  const maxScan = Math.max(60, worksheet.rowCount + 5);
  for (let r = 1; r <= maxScan; r += 1) {
    const parts = [];
    for (let c = 1; c <= maxCol; c += 1) {
      const t = getMergedAwareCellText(r, c);
      if (t) parts.push(t.toLowerCase());
    }
    if (titleRe.test(parts.join(' '))) {
      titleRow = r;
      break;
    }
  }
  if (titleRow < 1) return null;

  let headerRow = titleRow + 1;
  for (let r = titleRow + 1; r <= titleRow + 4; r += 1) {
    const parts = [];
    for (let c = 1; c <= maxCol; c += 1) {
      const t = getMergedAwareCellText(r, c);
      if (t) parts.push(t.toLowerCase());
    }
    const joined = parts.join(' ');
    if (/period/.test(joined) && (/total\s+leave/.test(joined) || /availed/.test(joined))) {
      headerRow = r;
      break;
    }
  }

  let subRow = -1;
  for (let r = headerRow + 1; r <= headerRow + 3; r += 1) {
    const parts = [];
    for (let c = 1; c <= maxCol; c += 1) {
      const t = getMergedAwareCellText(r, c);
      if (t) parts.push(t.toLowerCase());
    }
    const joined = parts.join(' ');
    if (/\bfrom\b/.test(joined) && /\bto\b/.test(joined)) {
      subRow = r;
      break;
    }
  }

  const cols = [];
  let c = 1;
  while (c <= maxCol) {
    const main = getMergedAwareCellText(headerRow, c);
    if (!main) {
      c += 1;
      continue;
    }
    const endC = getMergeSpanEnd(headerRow, c).c;
    const mainNorm = formNGJGujaratHeaderNorm(main);
    if (/^period$/i.test(mainNorm) || mainNorm === 'period') {
      if (subRow > 0) {
        let fromCol = -1;
        let toCol = -1;
        for (let pc = c; pc <= endC; pc += 1) {
          const sub = formNGJGujaratHeaderNorm(getMergedAwareCellText(subRow, pc));
          if (/^from$/.test(sub)) fromCol = pc;
          if (/^to$/.test(sub)) toCol = pc;
        }
        if (fromCol > 0) cols.push({ col: fromCol, bucket: 'periodFrom', label: 'Period_From' });
        if (toCol > 0) cols.push({ col: toCol, bucket: 'periodTo', label: 'Period_To' });
        if (fromCol < 0 && toCol < 0) {
          cols.push({ col: c, bucket: 'period', label: 'Period' });
        }
      } else {
        cols.push({ col: c, bucket: 'period', label: 'Period' });
      }
    } else if (/total\s+leave/.test(mainNorm)) {
      cols.push({ col: c, bucket: 'total', label: 'Total Leave' });
    } else if (/availed\s+leave/.test(mainNorm)) {
      cols.push({ col: c, bucket: 'availed', label: 'Availed Leave' });
    } else if (/balance\s+leave/.test(mainNorm)) {
      cols.push({ col: c, bucket: 'balance', label: 'Balance Leave' });
    } else if (/payment\s+made\s+in\s+lieu/.test(mainNorm)) {
      cols.push({ col: c, bucket: 'payment', label: 'Payment made in lieu of Festival Leave, when called' });
    } else if (/remarks/.test(mainNorm)) {
      cols.push({ col: c, bucket: 'remarks', label: 'Remarks' });
    }
    c = Math.max(c + 1, endC + 1);
  }

  const dataStartRow = (subRow > 0 ? subRow : headerRow) + 1;
  return { titleRow, headerRow, subRow, dataStartRow, cols };
}

function detectFormNGJGujaratSheetLayout(worksheet) {
  const getMergeTopLeft = buildMergeTopLeftResolver(worksheet);
  const getMergeSpanEnd = buildMergeSpanEndResolver(worksheet, getMergeTopLeft);
  const getMergedAwareCellText = (r, c) => {
    const tl = getMergeTopLeft(r, c);
    return excelCellValueToString(worksheet.getCell(tl.r, tl.c)?.value).trim();
  };

  const maxCol = 20;
  const maxScan = Math.max(45, worksheet.rowCount + 5);
  let leaveHeaderRow = -1;
  for (let r = 1; r <= maxScan; r += 1) {
    const parts = [];
    for (let c = 1; c <= maxCol; c += 1) {
      const t = getMergedAwareCellText(r, c);
      if (t) parts.push(t.toLowerCase());
    }
    const joined = parts.join(' ');
    if (/accumulation\s+of\s+leave/.test(joined) && /leave\s+allowed/.test(joined)) {
      leaveHeaderRow = r;
      break;
    }
  }
  if (leaveHeaderRow < 1) return null;

  let leaveSubRow = leaveHeaderRow + 1;
  for (let r = leaveHeaderRow + 1; r <= leaveHeaderRow + 4; r += 1) {
    const parts = [];
    for (let c = 1; c <= maxCol; c += 1) {
      const t = getMergedAwareCellText(r, c);
      if (t) parts.push(t.toLowerCase());
    }
    const joined = parts.join(' ');
    if (/leave\s+due\s+on/.test(joined) || /no\.?\s*of\s+days/.test(joined)) {
      leaveSubRow = r;
      break;
    }
  }

  const leaveCols = [];
  let c = 1;
  while (c <= maxCol) {
    const main = getMergedAwareCellText(leaveHeaderRow, c);
    if (!main || isFormNGJHeaderMetaLabel(main)) {
      c += 1;
      continue;
    }
    const endC = getMergeSpanEnd(leaveHeaderRow, c).c;
    const parent = normalizeFormNGJParentLabel(main);
    if (!parent) {
      c = Math.max(c + 1, endC + 1);
      continue;
    }

    if (parent === 'Remarks') {
      leaveCols.push({
        col: c,
        bucket: 'leaveRemarks',
        label: 'Remarks',
      });
    } else {
      let added = false;
      for (let sc = c; sc <= endC; sc += 1) {
        const subRaw = getMergedAwareCellText(leaveSubRow, sc);
        if (!subRaw || isNumericOnly(subRaw)) continue;
        const sub = normalizeFormNGJSubLabel(subRaw) || subRaw;
        const label = `${parent}_${sub}`;
        leaveCols.push({
          col: sc,
          bucket: formNGJExportHeaderBucket(label),
          label,
        });
        added = true;
      }
      if (!added) {
        leaveCols.push({
          col: c,
          bucket: formNGJExportHeaderBucket(parent),
          label: parent,
        });
      }
    }
    c = Math.max(c + 1, endC + 1);
  }

  const festival = detectFormNGJSectionColumns(
    worksheet,
    getMergedAwareCellText,
    getMergeSpanEnd,
    /details\s+of\s+festival\s+leave/,
    maxCol
  );
  const casual = detectFormNGJSectionColumns(
    worksheet,
    getMergedAwareCellText,
    getMergeSpanEnd,
    /details\s+of\s+casual\s+leave/,
    maxCol
  );

  return {
    leaveHeaderRow,
    leaveSubRow,
    leaveDataStartRow: leaveSubRow + 1,
    leaveCols,
    festival,
    casual,
    getMergedAwareCellText,
    getMergeTopLeft,
  };
}

function writeFormNGJValueBesideOrBelowLabel(worksheet, labelRow, labelCol, value, getMergedAwareCellText) {
  const val = String(value ?? '').trim();
  if (!val || labelRow < 1 || labelCol < 1) return;

  const rightText = getMergedAwareCellText(labelRow, labelCol + 1);
  const rightLooksLikeLabel =
    /name\s+of|description\s+of|date\s+of|receipt\s+of|department|employer|worker|establishment/i.test(
      rightText
    );
  if (!rightText || /^enter\b/i.test(rightText) || rightText === ':') {
    const cell = worksheet.getCell(labelRow, labelCol + 1);
    cell.value = val;
    cell.alignment = { ...(cell.alignment || {}), wrapText: true, vertical: 'middle' };
    return;
  }
  if (rightLooksLikeLabel) {
    const belowText = getMergedAwareCellText(labelRow + 1, labelCol);
    if (!belowText || /^enter\b/i.test(belowText)) {
      const cell = worksheet.getCell(labelRow + 1, labelCol);
      cell.value = val;
      cell.alignment = { ...(cell.alignment || {}), wrapText: true, vertical: 'middle' };
    }
    return;
  }
  const belowText = getMergedAwareCellText(labelRow + 1, labelCol);
  if (!belowText || /^enter\b/i.test(belowText)) {
    const cell = worksheet.getCell(labelRow + 1, labelCol);
    cell.value = val;
    cell.alignment = { ...(cell.alignment || {}), wrapText: true, vertical: 'middle' };
  }
}

function writeFormNGJWorkerHeaderFields(worksheet, layout, values) {
  const getText = layout.getMergedAwareCellText;
  const getMergeTopLeft = layout.getMergeTopLeft;
  const maxRow = Math.max(1, layout.leaveHeaderRow - 1);
  const specs = [
    { re: /name\s+of\s+the\s+establishment/i, value: values.establishment },
    { re: /name\s+of\s+the\s+worker/i, value: values.workerName },
    { re: /description\s+of\s+the\s+department/i, value: values.department },
    { re: /name\s+of\s+the\s+employer/i, value: values.employer },
    { re: /date\s+of\s+entry\s+into\s+service/i, value: values.dateOfEntry },
    { re: /receipt\s+of\s+le[av]e?\s*book|receipt\s+of\s+level\s+book/i, value: values.receipt },
  ];
  const written = new Set();

  for (let r = 1; r <= maxRow; r += 1) {
    for (let c = 1; c <= 20; c += 1) {
      const tl = getMergeTopLeft(r, c);
      if (tl.r !== r || tl.c !== c) continue;
      const raw = getText(r, c);
      if (!raw) continue;
      const labelOnly = String(raw).split(':')[0].trim();
      for (const spec of specs) {
        if (written.has(spec.re.source)) continue;
        if (!spec.re.test(labelOnly) && !spec.re.test(raw)) continue;
        if (!String(spec.value || '').trim()) {
          written.add(spec.re.source);
          continue;
        }
        writeFormNGJValueBesideOrBelowLabel(worksheet, r, c, spec.value, getText);
        written.add(spec.re.source);
      }
    }
  }
}

function writeFormNGJSectionRow(worksheet, section, rowValuesByBucket) {
  if (!section?.cols?.length || !(section.dataStartRow >= 1)) return;
  const targetRow = section.dataStartRow;
  section.cols.forEach(({ col, bucket }) => {
    let val = '';
    if (bucket === 'period') {
      const from = rowValuesByBucket.periodFrom || '';
      const to = rowValuesByBucket.periodTo || '';
      val = [from, to].filter(Boolean).join(' - ');
    } else if (bucket === 'periodFrom') {
      val = rowValuesByBucket.periodFrom || '';
    } else if (bucket === 'periodTo') {
      val = rowValuesByBucket.periodTo || '';
    } else {
      val = rowValuesByBucket[bucket] || '';
    }
    const cell = worksheet.getCell(targetRow, col);
    cell.value = val ? String(val) : '';
    if (val) {
      cell.alignment = { ...(cell.alignment || {}), wrapText: true, vertical: 'middle' };
    }
  });
}

function collectFormNGJHeaderValuesFromRow(row, headers, headerFormData, formatDate) {
  const hdrs = resolveFormNGJGujaratTableHeaders(headers);
  const workerHeader = hdrs.find(isFormNGJWorkerNameHeader);
  const deptHeader = hdrs.find(isFormNGJDepartmentHeader);
  const employerHeader = hdrs.find(isFormNGJEmployerHeader);
  const entryHeader = hdrs.find(isFormNGJDateOfEntryHeader);
  const receiptHeader = hdrs.find(isFormNGJReceiptHeader);
  const establishment =
    String(headerFormData?.form_n_gj_establishment ?? '').trim() ||
    String(headerFormData?.['Name of the establishment'] ?? '').trim();
  const dateOfEntryRaw = getFormNGJRowValueForHeader(
    row,
    entryHeader || 'Date of entry into service'
  );
  return {
    establishment,
    workerName: getFormNGJRowValueForHeader(row, workerHeader || 'Name of the worker'),
    department: getFormNGJRowValueForHeader(
      row,
      deptHeader || 'Description of the Department (if applicable)'
    ),
    employer: getFormNGJRowValueForHeader(row, employerHeader || 'Name of the employer'),
    dateOfEntry: dateOfEntryRaw ? formatDate(dateOfEntryRaw) : '',
    receipt: getFormNGJRowValueForHeader(row, receiptHeader || 'Receipt of level book'),
  };
}

function collectFormNGJLeaveValuesFromRow(row, headers, formatDate) {
  const hdrs = resolveFormNGJGujaratTableHeaders(headers);
  const pick = (pred, fallbackBucket) => {
    const header = hdrs.find(pred) || hdrs.find((h) => formNGJExportHeaderBucket(h) === fallbackBucket);
    const val = header ? getFormNGJRowValueForHeader(row, header) : '';
    return val;
  };
  const leaveDueOn = pick(
    (h) => formNGJExportHeaderBucket(h) === 'leaveDueOn',
    'leaveDueOn'
  );
  const applicationDate = pick(
    (h) => formNGJExportHeaderBucket(h) === 'applicationDate',
    'applicationDate'
  );
  const dateOfRefusal = pick(
    (h) => formNGJExportHeaderBucket(h) === 'dateOfRefusal',
    'dateOfRefusal'
  );
  const dateOfDischarge = pick(
    (h) => formNGJExportHeaderBucket(h) === 'dateOfDischarge',
    'dateOfDischarge'
  );
  return {
    leaveDueOn: leaveDueOn ? formatDate(leaveDueOn) : '',
    noOfDays: pick((h) => formNGJExportHeaderBucket(h) === 'noOfDays', 'noOfDays'),
    leaveFromTo: pick((h) => formNGJExportHeaderBucket(h) === 'leaveFromTo', 'leaveFromTo'),
    moiety1: pick((h) => formNGJExportHeaderBucket(h) === 'moiety1', 'moiety1'),
    moiety2: pick((h) => formNGJExportHeaderBucket(h) === 'moiety2', 'moiety2'),
    applicationDate: applicationDate ? formatDate(applicationDate) : '',
    dateOfRefusal: dateOfRefusal ? formatDate(dateOfRefusal) : '',
    dateOfDischarge: dateOfDischarge ? formatDate(dateOfDischarge) : '',
    dateAndAmountPaid: pick(
      (h) => formNGJExportHeaderBucket(h) === 'dateAndAmountPaid',
      'dateAndAmountPaid'
    ),
    signature: pick((h) => formNGJExportHeaderBucket(h) === 'signature', 'signature'),
    leaveRemarks: pick((h) => formNGJExportHeaderBucket(h) === 'leaveRemarks', 'leaveRemarks'),
  };
}

function collectFormNGJSectionValuesFromRow(row, headers, sectionParent, formatDate) {
  const hdrs = resolveFormNGJGujaratTableHeaders(headers);
  const findSection = (suffixRe, bucket) => {
    const header =
      hdrs.find(
        (h) =>
          String(h || '').startsWith(sectionParent) &&
          (suffixRe.test(String(h || '')) || formNGJExportHeaderBucket(h) === bucket)
      ) || '';
    return header ? getFormNGJRowValueForHeader(row, header) : '';
  };
  const from = findSection(/period_from|_from$/i, sectionParent.includes('Festival') ? 'festivalPeriodFrom' : 'casualPeriodFrom');
  const to = findSection(/period_to|_to$/i, sectionParent.includes('Festival') ? 'festivalPeriodTo' : 'casualPeriodTo');
  return {
    periodFrom: from ? formatDate(from) : '',
    periodTo: to ? formatDate(to) : '',
    total: findSection(/total\s+leave/i, sectionParent.includes('Festival') ? 'festivalTotal' : 'casualTotal'),
    availed: findSection(
      /availed\s+leave/i,
      sectionParent.includes('Festival') ? 'festivalAvailed' : 'casualAvailed'
    ),
    balance: findSection(
      /balance\s+leave/i,
      sectionParent.includes('Festival') ? 'festivalBalance' : 'casualBalance'
    ),
    payment: findSection(/payment\s+made\s+in\s+lieu/i, 'festivalPayment'),
    remarks: findSection(/remarks/i, sectionParent.includes('Festival') ? 'festivalRemarks' : 'casualRemarks'),
  };
}

/**
 * Write one Form N Leave Book into the original Excel template (ExcelJS keeps merges).
 */
export async function buildFormNGJGujaratWorkbookWithTemplateStyles({
  templateArrayBuffer,
  mappedData,
  headersToUse,
  parsedFormHeader,
  formFileName,
  headerFormData,
  sheetNameHint,
  formatStatutoryDateDisplay = null,
}) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(templateArrayBuffer);
  const preferred = String(sheetNameHint || '').trim();
  const worksheet =
    (preferred && workbook.worksheets.find((ws) => String(ws?.name || '').trim() === preferred)) ||
    workbook.worksheets[0] ||
    null;
  if (!worksheet) throw new Error('Template worksheet not found.');

  const layout = detectFormNGJGujaratSheetLayout(worksheet);
  if (!layout) throw new Error('Could not locate Gujarat Form N leave table header row.');

  const formatDate =
    typeof formatStatutoryDateDisplay === 'function'
      ? formatStatutoryDateDisplay
      : (v) => String(v || '').trim();

  const normalizedHeaders = resolveFormNGJGujaratTableHeaders(headersToUse);
  const rows = filterFormNGJGujaratExportRows(
    remapFormNGJGujaratRowsToHeaders(
      Array.isArray(mappedData) ? mappedData : [],
      normalizedHeaders,
      normalizedHeaders
    ),
    normalizedHeaders
  );
  const row = rows[0] || (Array.isArray(mappedData) ? mappedData[0] : {}) || {};

  const headerValues = collectFormNGJHeaderValuesFromRow(
    row,
    normalizedHeaders,
    headerFormData && typeof headerFormData === 'object' ? headerFormData : {},
    formatDate
  );
  writeFormNGJWorkerHeaderFields(worksheet, layout, headerValues);

  const leaveValues = collectFormNGJLeaveValuesFromRow(row, normalizedHeaders, formatDate);
  if (layout.leaveCols.length > 0) {
    const dataRow = layout.leaveDataStartRow;
    layout.leaveCols.forEach(({ col, bucket }) => {
      const val = leaveValues[bucket] || '';
      const cell = worksheet.getCell(dataRow, col);
      cell.value = val ? String(val) : '';
      if (val) {
        cell.alignment = { ...(cell.alignment || {}), wrapText: true, vertical: 'middle' };
      }
    });
    const colFrom = Math.min(...layout.leaveCols.map((x) => x.col));
    const colTo = Math.max(...layout.leaveCols.map((x) => x.col));
    ensureExcelJSDataRowsWithBorders(worksheet, {
      dataStartRow: dataRow,
      dataRowCount: 1,
      colFrom,
      colTo,
      templateRow: dataRow,
      templateBodyRows: 1,
    });
  }

  writeFormNGJSectionRow(
    worksheet,
    layout.festival,
    collectFormNGJSectionValuesFromRow(row, normalizedHeaders, FORM_NGJ_FESTIVAL_PARENT, formatDate)
  );
  writeFormNGJSectionRow(
    worksheet,
    layout.casual,
    collectFormNGJSectionValuesFromRow(row, normalizedHeaders, FORM_NGJ_CASUAL_PARENT, formatDate)
  );

  const out = await workbook.xlsx.writeBuffer();
  const fileName =
    formFileName ||
    parsedFormHeader?.title?.replace(/[^a-zA-Z0-9]/g, '_') ||
    'Form_N_GJ_Gujarat.xlsx';
  return {
    blob: new Blob([out], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    }),
    fileName,
  };
}

function sanitizeFormNGJGujaratDownloadFileName(name) {
  return (
    String(name || 'Employee')
      .replace(/[\\/:*?"<>|]+/g, '_')
      .replace(/\s+/g, '_')
      .replace(/_+/g, '_')
      .replace(/^_|_$/g, '')
      .slice(0, 80) || 'Employee'
  );
}

export function resolveFormNGJGujaratEmployeeDownloadBaseName(row, headers, rowIndex = 0) {
  const lookupName = String(row?.__employeeLookupName || '').trim();
  if (lookupName) return sanitizeFormNGJGujaratDownloadFileName(lookupName);
  const hdrs = resolveFormNGJGujaratTableHeaders(headers);
  const workerHeader = hdrs.find(isFormNGJWorkerNameHeader);
  const val = getFormNGJRowValueForHeader(row, workerHeader || 'Name of the worker');
  if (val) return sanitizeFormNGJGujaratDownloadFileName(val.split(/\r?\n/)[0].trim());
  return sanitizeFormNGJGujaratDownloadFileName(`Employee_${rowIndex + 1}`);
}

function allocateUniqueFormNGJGujaratDownloadFileName(baseName, usedNames) {
  const safeBase = sanitizeFormNGJGujaratDownloadFileName(baseName);
  const count = usedNames.get(safeBase) || 0;
  usedNames.set(safeBase, count + 1);
  if (count === 0) return `${safeBase}.xlsx`;
  return `${safeBase}_${count + 1}.xlsx`;
}

/** One Leave Book per worker — ZIP when multiple employees. */
export async function buildFormNGJGujaratPerEmployeeDownload({
  templateArrayBuffer,
  mappedData,
  headersToUse,
  parsedFormHeader,
  formFileName,
  headerFormData,
  sheetNameHint,
  formatStatutoryDateDisplay = null,
}) {
  const hdrs = resolveFormNGJGujaratTableHeaders(headersToUse);
  let exportRows = filterFormNGJGujaratExportRows(
    remapFormNGJGujaratRowsToHeaders(
      Array.isArray(mappedData) ? mappedData : [],
      hdrs,
      hdrs
    ),
    hdrs
  );
  if (exportRows.length === 0 && Array.isArray(mappedData) && mappedData.length > 0) {
    exportRows = [mappedData[0]];
  }

  const workbookArgs = {
    templateArrayBuffer,
    headersToUse: hdrs,
    parsedFormHeader,
    formFileName,
    headerFormData,
    sheetNameHint,
    formatStatutoryDateDisplay,
  };

  if (exportRows.length <= 1) {
    return buildFormNGJGujaratWorkbookWithTemplateStyles({
      ...workbookArgs,
      mappedData: exportRows,
    });
  }

  const zip = new JSZip();
  const usedNames = new Map();
  for (let i = 0; i < exportRows.length; i += 1) {
    const { blob } = await buildFormNGJGujaratWorkbookWithTemplateStyles({
      ...workbookArgs,
      mappedData: [exportRows[i]],
    });
    const xlsxBytes = await blob.arrayBuffer();
    const baseName = resolveFormNGJGujaratEmployeeDownloadBaseName(exportRows[i], hdrs, i);
    zip.file(allocateUniqueFormNGJGujaratDownloadFileName(baseName, usedNames), xlsxBytes);
    if (i > 0 && i % 25 === 0) {
      await new Promise((resolve) => setTimeout(resolve, 0));
    }
  }
  const zipBase = String(formFileName || parsedFormHeader?.title || 'Form_N_GJ')
    .replace(/\.xlsx?$/i, '')
    .replace(/[^a-zA-Z0-9._-]+/g, '_');
  return {
    blob: await zip.generateAsync({ type: 'blob', compression: 'STORE' }),
    fileName: `${zipBase}_Employees.zip`,
  };
}
