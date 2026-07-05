/** Gujarat Form N — Leave Book (See rule 17). */

import * as XLSX from 'xlsx';

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

  const bucketFor = (header) => {
    const n = normHeaderLabel(header);
    if (/name\s+of\s+the\s+worker/.test(n)) return 'workerName';
    if (/description\s+of\s+the\s+department/.test(n)) return 'department';
    if (/name\s+of\s+the\s+employer/.test(n)) return 'employer';
    if (/date\s+of\s+entry\s+into\s+service/.test(n)) return 'dateOfEntry';
    if (/receipt\s+of\s+level\s+book/.test(n)) return 'receipt';
    return n;
  };

  return rows.map((row) => {
    if (!row || typeof row !== 'object') return {};
    const out = {};
    tgt.forEach((targetHeader, colIdx) => {
      let val = '';
      if (Object.prototype.hasOwnProperty.call(row, targetHeader)) {
        val = row[targetHeader];
      } else if (src[colIdx] && Object.prototype.hasOwnProperty.call(row, src[colIdx])) {
        val = row[src[colIdx]];
      } else {
        const bucket = bucketFor(targetHeader);
        for (const [k, v] of Object.entries(row)) {
          if (bucketFor(k) === bucket) {
            val = v;
            break;
          }
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
    headers: resolveFormNGJGujaratTableHeaders(parsed?.headers || hints.tableHeaders || []),
    subColumns: buildFormNGJSubColumnsFromHeaders(
      resolveFormNGJGujaratTableHeaders(parsed?.headers || hints.tableHeaders || [])
    ),
    tableData: [],
    headerRowIndex: -1,
    dataStartIndex: 0,
    tableStartCol: 0,
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

/** Leave / festival columns — manual entry only. */
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
  if (fullName) out.__employeeLookupName = fullName;

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
