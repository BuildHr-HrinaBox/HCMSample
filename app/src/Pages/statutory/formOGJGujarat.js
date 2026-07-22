/** Gujarat Form O — notice of accumulated leave (See rule 18). */

import ExcelJS from 'exceljs';
import { ensureExcelJSDataRowsWithBorders } from '../../utils/excelTableBorders';
import { writeStatutoryHeaderFieldsToExcelJsWorksheet } from '../../utils/statutorySiteCompanyHeaders';
import {
  collectApprovedLeaveIdentityKeys,
  getApprovedLeaveRecordUniqueKey,
  normalizeApprovedLeaveRecord,
  normalizePersonNameKey,
} from './formFKarnataka';

export const FORM_OGJ_PERIOD_PARENT = 'Period for which leave is accumulated';

export const FORM_OGJ_GJ_CANONICAL_TABLE_HEADERS = [
  'Sr. No.',
  'Name of Workers',
  'Number of accumulated leave',
  `${FORM_OGJ_PERIOD_PARENT}_From`,
  `${FORM_OGJ_PERIOD_PARENT}_Till`,
];

export function formOGJGujaratHeaderNorm(txt) {
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
  return formOGJGujaratHeaderNorm(stripLeadingNumber(h))
    .replace(/[^a-z0-9 ]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

/** Fix common Excel typo and normalize period parent label. */
export function normalizeFormOGJPeriodParentLabel(raw) {
  const text = String(raw || '')
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  if (!text) return FORM_OGJ_PERIOD_PARENT;
  const norm = formOGJGujaratHeaderNorm(text);
  if (/perod|period/.test(norm) && /accumulated/.test(norm) && /leave/.test(norm)) {
    return FORM_OGJ_PERIOD_PARENT;
  }
  return text;
}

export function isFormOGJGujaratContext(
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

  if (/\bform\s*q\b|\bform_q\b|form[\s._-]*q[\s._-]*gj/.test(parts)) return false;
  if (/\bform\s*p\b|\bform_p\b|form[\s._-]*p[\s._-]*gj/.test(parts)) return false;

  const hasGujarat = /gujarat|\b_gj\b|form[\s._-]*o[\s._-]*gj|form_o_gj/.test(parts);
  const hasFormO = /\bform[\s._-]*o\b/.test(parts);
  if (hasGujarat && hasFormO) return true;
  if (/\bform_o_gj\b/.test(parts)) return true;
  if (hasFormO && /see\s+rule\s+18/.test(parts)) return true;

  const joined = parts;
  return (
    /sr\.?\s*no/.test(joined) &&
    /name\s+of\s+workers?/.test(joined) &&
    /accumulated\s+leave/.test(joined) &&
    (/period|perod/.test(joined) || /from/.test(joined) || /till/.test(joined))
  );
}

export function isFormOGJGujaratTableLayoutFormHeader(formHeader) {
  return !!formHeader?.formOGJGujaratTableLayout;
}

export function resolveFormOGJGujaratTableHeaders(_tableHeaders) {
  return [...FORM_OGJ_GJ_CANONICAL_TABLE_HEADERS];
}

export function buildFormOGJSubColumnsFromHeaders(headers) {
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

const isEmployerMetaLabel = (txt) => {
  const t = formOGJGujaratHeaderNorm(txt);
  return (
    /name\s+and\s+address\s+of\s+the\s+establishment/.test(t) ||
    /name\s+of\s+the\s+authorized/.test(t) ||
    /authorized\s+person/.test(t) ||
    /^notice$/.test(t) ||
    /^date\s*:?$/.test(t) ||
    /^place\s*:?$/.test(t)
  );
};

const isPeriodSubLabel = (sub) => /^(from|till|to)$/i.test(String(sub || '').trim());

/**
 * Parse Form O GJ Excel header rows (parent + From/Till under Period column).
 */
export function rebuildFormOGJGujaratTableHeadersFromSheet({
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

  let headerRowIndex = -1;
  for (let r = 0; r < Math.min(35, jsonData.length); r += 1) {
    const parts = [];
    for (let c = 0; c < Math.max(effectiveSheetCols, 8); c += 1) {
      const t = normLower(getMergedAwareCellText(r, c));
      if (t) parts.push(t);
    }
    const joined = parts.join(' ');
    if (
      /sr\.?\s*no/.test(joined) &&
      /name\s+of\s+workers?/.test(joined) &&
      /accumulated\s+leave/.test(joined)
    ) {
      headerRowIndex = r;
      break;
    }
  }
  if (headerRowIndex < 0) return null;

  let startCol = -1;
  for (let c = 0; c < Math.max(effectiveSheetCols, 8); c += 1) {
    const t = normLower(getMergedAwareCellText(headerRowIndex, c));
    if (/^sr\.?\s*no\.?$/.test(t) || t === 'sr no' || t.startsWith('sr. no')) {
      startCol = c;
      break;
    }
  }
  if (startCol < 0) startCol = 0;

  const subRowIndex = headerRowIndex + 1;
  const flatHeaders = [];
  const subColumnsData = {};
  const mainHeaders = [];

  let c = startCol;
  const maxCols = Math.max(effectiveSheetCols, (jsonData[headerRowIndex] || []).length, startCol + 6);
  while (c < maxCols) {
    const mergeEnd = getMergeSpanEndCol(headerRowIndex, c);
    let endC = mergeEnd;
    let main = norm(getMergedAwareCellText(headerRowIndex, c));
    if (!main || isEmployerMetaLabel(main)) {
      c = Math.max(c + 1, endC);
      continue;
    }
    while (endC < maxCols) {
      const nextMain = norm(getMergedAwareCellText(headerRowIndex, endC));
      if (!nextMain) break;
      if (normLower(nextMain) !== normLower(main)) break;
      endC = Math.max(endC, getMergeSpanEndCol(headerRowIndex, endC));
    }

    const mainLower = normLower(main);
    const isPeriodBand =
      (/perod|period/.test(mainLower) && /accumulated/.test(mainLower) && /leave/.test(mainLower)) ||
      (endC - c >= 2 && /perod|period/.test(mainLower));

    if (isPeriodBand) {
      main = normalizeFormOGJPeriodParentLabel(main);
    }

    const subs = [];
    for (let sc = c; sc < endC; sc += 1) {
      const sub = norm(getMergedAwareCellText(subRowIndex, sc));
      if (!sub || sub.toLowerCase() === main.toLowerCase()) continue;
      if (sub.endsWith(':') && sub.length < 40) continue;
      if (isPeriodSubLabel(sub) || (!/^\d+$/.test(sub) && sub.length > 0 && sub.length <= 20)) {
        subs.push(/^to$/i.test(sub) ? 'Till' : sub.charAt(0).toUpperCase() + sub.slice(1).toLowerCase());
      }
    }

    mainHeaders.push(main);
    if (!subColumnsData[main]) subColumnsData[main] = [];

    if (isPeriodBand && subs.length === 0) {
      subs.push('From', 'Till');
    }

    if (subs.length === 0) {
      flatHeaders.push(main);
    } else {
      subs.forEach((sub) => {
        flatHeaders.push(`${main}_${sub}`);
        subColumnsData[main].push(sub);
      });
    }
    c = endC;
  }

  if (flatHeaders.length < 4) return null;

  let subHeaderRowIndex = subRowIndex;
  outerSubScan: for (let sr = headerRowIndex + 1; sr <= headerRowIndex + 4; sr += 1) {
    for (let sc = startCol; sc < maxCols; sc += 1) {
      const sub = normLower(getMergedAwareCellText(sr, sc));
      if (sub === 'from' || sub === 'till' || sub === 'to') {
        subHeaderRowIndex = sr;
        break outerSubScan;
      }
    }
  }

  const canonical = resolveFormOGJGujaratTableHeaders();
  const expandedHeaders =
    flatHeaders.length >= 4
      ? flatHeaders.map((h) => {
          const m = String(h).match(/^(.+)_(From|Till|To)$/i);
          if (m) {
            return `${normalizeFormOGJPeriodParentLabel(m[1])}_${/^to$/i.test(m[2]) ? 'Till' : m[2].charAt(0).toUpperCase() + m[2].slice(1).toLowerCase()}`;
          }
          const n = normHeaderLabel(h);
          if (/^sr\.?\s*no/.test(n)) return 'Sr. No.';
          if (/name\s+of\s+workers?/.test(n)) return 'Name of Workers';
          if (/number\s+of\s+accumulated\s+leave/.test(n)) return 'Number of accumulated leave';
          // Template typo "Name of accumulated leave" ≡ official "Number of accumulated leave"
          if (/name\s+of\s+accumulated\s+leave/.test(n)) return 'Number of accumulated leave';
          if (/accumulated\s+leave/.test(n) && !/period|perod/.test(n)) return 'Number of accumulated leave';
          return h;
        })
      : canonical;

  const normalizedSubs = {};
  Object.keys(subColumnsData).forEach((key) => {
    const parent = normalizeFormOGJPeriodParentLabel(key);
    normalizedSubs[parent] = subColumnsData[key];
  });

  // Always start after From/Till sub-header (never write into the period labels row).
  const dataStartIndex = Math.max(headerRowIndex + 3, subHeaderRowIndex + 1);

  return {
    headers: mainHeaders,
    expandedHeaders,
    subColumnsData: normalizedSubs,
    headerRowIndex,
    startCol,
    dataStartIndex,
  };
}

export function remapFormOGJGujaratRowsToHeaders(rows, sourceHeaders, targetHeaders) {
  const src = Array.isArray(sourceHeaders) ? sourceHeaders : [];
  const tgt = resolveFormOGJGujaratTableHeaders(targetHeaders);
  if (!Array.isArray(rows)) return [];

  const bucketFor = (header) => {
    const n = normHeaderLabel(header);
    if (/^sr\.?\s*no/.test(n)) return 'sno';
    if (/name\s+of\s+workers?/.test(n)) return 'workerName';
    if (/number\s+of\s+accumulated\s+leave/.test(n)) return 'leaveCount';
    // Template typo "Name …" stores the same LeaveCount value as "Number …"
    if (/name\s+of\s+accumulated\s+leave/.test(n)) return 'leaveCount';
    if (/accumulated\s+leave/.test(n) && !/period|perod/.test(n)) return 'leaveCount';
    if (/period|perod/.test(n) && /from/.test(n)) return 'periodFrom';
    if (/period|perod/.test(n) && /(till|to)/.test(n)) return 'periodTill';
    if (/^from$/i.test(String(header).split('_').pop() || '')) return 'periodFrom';
    if (/^(till|to)$/i.test(String(header).split('_').pop() || '')) return 'periodTill';
    return n;
  };

  return rows.map((row, rowIndex) => {
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
      if ((val == null || val === '') && bucketFor(targetHeader) === 'sno') {
        val = rowIndex + 1;
      }
      out[targetHeader] = val == null ? '' : val;
    });
    return out;
  });
}

export function isFormOGJSerialHeader(header) {
  return /^sr\.?\s*no/i.test(normHeaderLabel(header));
}

export function isFormOGJWorkerNameHeader(header) {
  const n = normHeaderLabel(header);
  return /name\s+of\s+workers?/.test(n);
}

export function isFormOGJAccumulatedLeaveCountHeader(header) {
  const n = normHeaderLabel(header);
  return /number\s+of\s+accumulated\s+leave/.test(n);
}

export function isFormOGJAccumulatedLeaveNameHeader(header) {
  const n = normHeaderLabel(header);
  // Excel templates often label this "Name of accumulated leave" (typo for Number).
  return /name\s+of\s+accumulated\s+leave/.test(n);
}

export function isFormOGJAccumulatedLeaveHeader(header) {
  const n = normHeaderLabel(header);
  return /accumulated\s+leave/.test(n) && !/period|perod/.test(n);
}

export function isFormOGJPeriodFromHeader(header) {
  const sub = String(header || '').split('_').pop() || '';
  if (/^from$/i.test(sub.trim())) return true;
  const n = normHeaderLabel(header);
  return /period|perod/.test(n) && /\bfrom\b/.test(n);
}

export function isFormOGJPeriodTillHeader(header) {
  const sub = String(header || '').split('_').pop() || '';
  if (/^(till|to)$/i.test(sub.trim())) return true;
  const n = normHeaderLabel(header);
  return /period|perod/.test(n) && /\b(till|to)\b/.test(n);
}

/** Leave type + period dates come from approved leave API, not People overlay. */
export function isFormOGJSkipPeopleAutofillHeader(header) {
  return (
    isFormOGJAccumulatedLeaveHeader(header) ||
    isFormOGJPeriodFromHeader(header) ||
    isFormOGJPeriodTillHeader(header)
  );
}

export function readApprovedLeaveLeaveType(record) {
  if (!record || typeof record !== 'object') return '';
  const raw =
    record['Leave Type'] ??
    record['Leave type'] ??
    record.LeaveType ??
    record.leaveType ??
    record.Leavetype ??
    record['Leavetype'] ??
    record['Type of Leave'] ??
    record.TypeOfLeave ??
    record.leave_type ??
    record.LeaveName ??
    record['Leave Name'] ??
    '';
  if (raw != null && typeof raw === 'object') {
    return String(
      raw.name ?? raw.Name ?? raw.value ?? raw.label ?? raw.displayValue ?? raw.display_value ?? ''
    ).trim();
  }
  const text = String(raw || '').trim();
  if (text) return text;
  const reason = String(record.Reason ?? record.reason ?? '').trim();
  return reason;
}

function zohoLeaveDateToMs(dateStr) {
  const m = String(dateStr || '').match(/^(\d{1,2})-([A-Za-z]{3})-(\d{4})$/i);
  if (!m) return 0;
  const months = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];
  const mon = months.indexOf(m[2].toLowerCase());
  if (mon < 0) return 0;
  return new Date(parseInt(m[3], 10), mon, parseInt(m[1], 10)).getTime();
}

/** Form O: accept records with valid From/To overlapping the statutory month (no daysCount gate). */
export function approvedLeaveRecordOverlapsFormOGJMonth(record, monthFrom, monthTo) {
  if (!record || typeof record !== 'object') return false;
  const status = String(
    record.ApprovalStatus ?? record.approvalStatus ?? record.Status ?? record.status ?? 'APPROVED'
  )
    .trim()
    .toUpperCase();
  if (status && status !== 'APPROVED') return false;

  const metrics = normalizeApprovedLeaveRecord(record);
  if (!metrics.from) return false;

  const rangeStart = zohoLeaveDateToMs(monthFrom);
  const rangeEnd = zohoLeaveDateToMs(monthTo);
  const leaveStart = zohoLeaveDateToMs(metrics.from);
  const leaveEnd = zohoLeaveDateToMs(metrics.to || metrics.from);
  if (!rangeStart || !rangeEnd || !leaveStart) return true;
  const end = leaveEnd || leaveStart;
  return leaveStart <= rangeEnd && end >= rangeStart;
}

export function filterApprovedLeaveRecordsForFormOGJMonth(records, monthFrom, monthTo) {
  if (!monthFrom || !monthTo) return Array.isArray(records) ? records : [];
  return (Array.isArray(records) ? records : []).filter((record) =>
    approvedLeaveRecordOverlapsFormOGJMonth(record, monthFrom, monthTo)
  );
}

/**
 * Strict full-name match for Form O — never match on last name or first name alone.
 * "chetan kumar" must not match "kumar singh" or "ajikumar singh".
 */
export function formOGJPersonNamesMatchStrict(workerName, recordName) {
  const a = normalizePersonNameKey(workerName);
  const b = normalizePersonNameKey(recordName);
  if (!a || !b) return false;
  if (a === b) return true;

  const aParts = a.split(' ').filter(Boolean);
  const bParts = b.split(' ').filter(Boolean);
  if (aParts.length < 2 || bParts.length < 2) return false;

  if (aParts[0] !== bParts[0]) return false;

  const aLast = aParts[aParts.length - 1];
  const bLast = bParts[bParts.length - 1];
  if (aLast === bLast) return true;

  // Allow single-letter last initial on the form vs full surname in Zoho (e.g. "muthukrishnan p" ↔ "muthukrishnan paldurai").
  if (aLast.length === 1 && bLast.startsWith(aLast)) return true;
  if (bLast.length === 1 && aLast.startsWith(bLast)) return true;

  return false;
}

function readEmployeeZohoIds(emp) {
  if (!emp || typeof emp !== 'object') return [];
  const ids = new Set();
  const add = (value) => {
    const id = String(value || '').trim().toLowerCase();
    if (id && id.length >= 4 && !/^\d{1,3}$/.test(id)) ids.add(id);
  };
  add(emp.Zoho_ID);
  add(emp['Zoho_ID']);
  add(emp.ZohoID);
  add(emp['ZohoID']);
  add(emp.erecno);
  add(emp.Erecno);
  add(emp['Erecno']);
  add(emp.Employee_ID);
  add(emp['Employee ID']);
  add(emp['Employee.ID']);
  add(emp.employeeId);
  return [...ids];
}

function formOGJRecordMatchesWorker(record, workerName, emp, options = {}) {
  if (!record) return false;

  const { ids: recordIds, names: recordNames } = collectApprovedLeaveIdentityKeys(record);

  const workerIds = new Set(readEmployeeZohoIds(emp));
  if (typeof options.collectEmployeeIdCandidates === 'function' && emp) {
    options.collectEmployeeIdCandidates(emp, {}).forEach((id) => {
      const normalized = String(id || '').trim().toLowerCase();
      if (normalized && normalized.length >= 4) workerIds.add(normalized);
    });
  }
  if (workerIds.size > 0 && recordIds.some((id) => workerIds.has(id))) {
    return true;
  }

  const workerNorm = normalizePersonNameKey(workerName);
  if (!workerNorm) return false;

  for (const recordName of recordNames) {
    if (formOGJPersonNamesMatchStrict(workerName, recordName)) return true;
  }

  return false;
}

export function findApprovedLeaveForFormORow(
  row,
  tableHeaders,
  approvedRecords,
  emp,
  options = {}
) {
  if (!Array.isArray(approvedRecords) || approvedRecords.length === 0) return null;

  const canonicalHeaders = resolveFormOGJGujaratTableHeaders(tableHeaders);
  const workerHeader = canonicalHeaders.find(isFormOGJWorkerNameHeader) || '';
  const workerName = workerHeader ? String(row?.[workerHeader] ?? '').trim() : '';
  if (!workerName) return null;

  const usedRecordKeys =
    options.usedRecordKeys instanceof Set ? options.usedRecordKeys : new Set();
  const monthFrom = options.monthFrom || '';
  const monthTo = options.monthTo || '';

  for (let i = 0; i < approvedRecords.length; i += 1) {
    const rec = approvedRecords[i];
    if (!rec) continue;
    const key = getApprovedLeaveRecordUniqueKey(rec);
    if (key && usedRecordKeys.has(key)) continue;
    if (monthFrom && monthTo && !approvedLeaveRecordOverlapsFormOGJMonth(rec, monthFrom, monthTo)) {
      continue;
    }
    if (formOGJRecordMatchesWorker(rec, workerName, emp, options)) {
      return rec;
    }
  }
  return null;
}

export function getFormOGJRowValueForHeader(row, header) {
  if (!row || !header) return '';
  const direct = row[header];
  if (direct != null && String(direct).trim() !== '') return String(direct).trim();

  const canonicalHeaders = FORM_OGJ_GJ_CANONICAL_TABLE_HEADERS;
  const bucketFor = (h) => {
    if (isFormOGJSerialHeader(h)) return 'sno';
    if (isFormOGJWorkerNameHeader(h)) return 'workerName';
    if (isFormOGJAccumulatedLeaveCountHeader(h)) return 'leaveCount';
    // Template "Name of accumulated leave" ≡ LeaveCount / Number column
    if (isFormOGJAccumulatedLeaveNameHeader(h)) return 'leaveCount';
    if (isFormOGJAccumulatedLeaveHeader(h)) return 'leaveCount';
    if (isFormOGJPeriodFromHeader(h)) return 'periodFrom';
    if (isFormOGJPeriodTillHeader(h)) return 'periodTill';
    return normHeaderLabel(h);
  };
  const bucket = bucketFor(header);
  for (const [k, v] of Object.entries(row)) {
    if (String(k).startsWith('__')) continue;
    if (bucketFor(k) === bucket && v != null && String(v).trim() !== '') {
      return String(v).trim();
    }
  }
  const idx = canonicalHeaders.findIndex((h) => bucketFor(h) === bucket);
  if (idx >= 0) {
    const canonKey = canonicalHeaders[idx];
    const canonVal = row[canonKey];
    if (canonVal != null && String(canonVal).trim() !== '') return String(canonVal).trim();
  }
  return '';
}

function setFormOGJRowCell(row, header, value, overwrite) {
  if (value == null || String(value).trim() === '') return false;
  if (!overwrite && String(row[header] ?? '').trim() !== '') return false;
  row[header] = String(value).trim();
  return true;
}

export function applyFormOGJGujaratApprovedLeaveToRow(
  row,
  approvedRecord,
  tableHeaders,
  { overwrite = true, monthFrom = '', monthTo = '' } = {}
) {
  if (!row || !approvedRecord) return 0;
  if (
    monthFrom &&
    monthTo &&
    !approvedLeaveRecordOverlapsFormOGJMonth(approvedRecord, monthFrom, monthTo)
  ) {
    return 0;
  }
  const metrics = normalizeApprovedLeaveRecord(approvedRecord);
  const leaveType = readApprovedLeaveLeaveType(approvedRecord);
  const leaveCount = metrics.daysCount;
  const periodFrom = metrics.from;
  const periodTill = metrics.to;
  if (!leaveCount && !leaveType && !periodFrom && !periodTill) return 0;

  const hdrs = resolveFormOGJGujaratTableHeaders(tableHeaders);
  let applied = 0;
  hdrs.forEach((header) => {
    if (isFormOGJAccumulatedLeaveCountHeader(header) && leaveCount) {
      if (setFormOGJRowCell(row, header, leaveCount, overwrite)) applied += 1;
    } else if (isFormOGJAccumulatedLeaveNameHeader(header) && leaveCount) {
      // Template typo "Name …" still receives LeaveCount (same as Number column).
      if (setFormOGJRowCell(row, header, leaveCount, overwrite)) applied += 1;
    } else if (isFormOGJAccumulatedLeaveNameHeader(header) && leaveType) {
      if (setFormOGJRowCell(row, header, leaveType, overwrite)) applied += 1;
    } else if (
      !isFormOGJAccumulatedLeaveCountHeader(header) &&
      !isFormOGJAccumulatedLeaveNameHeader(header) &&
      isFormOGJAccumulatedLeaveHeader(header) &&
      leaveCount
    ) {
      if (setFormOGJRowCell(row, header, leaveCount, overwrite)) applied += 1;
    } else if (isFormOGJPeriodFromHeader(header) && periodFrom) {
      if (setFormOGJRowCell(row, header, periodFrom, overwrite)) applied += 1;
    } else if (isFormOGJPeriodTillHeader(header) && periodTill) {
      if (setFormOGJRowCell(row, header, periodTill, overwrite)) applied += 1;
    }
  });

  // Alias keys so UI reads values even if header labels differ slightly from canonical keys.
  if (applied > 0) {
    hdrs.forEach((header) => {
      const val = row[header];
      if (val == null || String(val).trim() === '') return;
      const bucket = normHeaderLabel(header);
      Object.keys(row).forEach((k) => {
        if (k === header) return;
        if (normHeaderLabel(k) === bucket) row[k] = val;
      });
    });
  }
  return applied;
}

export function applyFormOGJGujaratApprovedLeaveAutofill(
  mappedData,
  employeesForMapping,
  tableHeaders,
  approvedLeaveRecords,
  options = {}
) {
  if (!Array.isArray(mappedData) || mappedData.length === 0) return 0;
  if (!Array.isArray(approvedLeaveRecords) || approvedLeaveRecords.length === 0) return 0;

  const priorHeaders = Array.isArray(tableHeaders) ? tableHeaders : [];
  const canonicalHeaders = resolveFormOGJGujaratTableHeaders(tableHeaders);
  const monthFrom = options.monthFrom || '';
  const monthTo = options.monthTo || '';
  const monthFiltered = filterApprovedLeaveRecordsForFormOGJMonth(
    approvedLeaveRecords,
    monthFrom,
    monthTo
  );
  const recordsToUse =
    monthFrom && monthTo ? monthFiltered : approvedLeaveRecords;
  const usedRecordKeys = new Set();
  let hits = 0;

  mappedData.forEach((row, index) => {
    const normalizedRow = remapFormOGJGujaratRowsToHeaders([row], priorHeaders, canonicalHeaders)[0];
    Object.assign(row, normalizedRow);

    const empItem = employeesForMapping[index];
    const emp = empItem && (empItem.Employee || empItem.employee || empItem);
    const approvedRecord = findApprovedLeaveForFormORow(row, canonicalHeaders, recordsToUse, emp, {
      usedRecordKeys,
      monthFrom,
      monthTo,
      collectEmployeeIdCandidates: options.collectEmployeeIdCandidates,
    });
    if (!approvedRecord) {
      // Clear stale leave cells when no strict match (e.g. prior wrong "kumar" match).
      canonicalHeaders.forEach((header) => {
        if (isFormOGJSkipPeopleAutofillHeader(header)) row[header] = '';
      });
      return;
    }

    const recordKey = getApprovedLeaveRecordUniqueKey(approvedRecord);
    if (recordKey) usedRecordKeys.add(recordKey);

    const applied = applyFormOGJGujaratApprovedLeaveToRow(
      row,
      approvedRecord,
      canonicalHeaders,
      { overwrite: options.overwrite !== false, monthFrom, monthTo }
    );
    if (applied > 0) hits += 1;
  });
  return hits;
}

export function enrichFormOGJGujaratDisplayHeader(formHeader, fileName, rowItem, tableHeaders) {
  const ctx = isFormOGJGujaratContext(formHeader, rowItem, fileName, '', tableHeaders);
  if (!ctx) return formHeader;
  return {
    ...(formHeader || {}),
    title: formHeader?.title || 'FORM - O',
    subtitle: formHeader?.subtitle || '(See rule 18)',
    formOGJGujaratTableLayout: true,
  };
}

function formOGJHeaderAliasBucket(n) {
  const norm = normHeaderLabel(n);
  if (/^sr\.?\s*no/.test(norm)) return 'sno';
  if (/name\s+of\s+workers?/.test(norm)) return 'workerName';
  if (/number\s+of\s+accumulated\s+leave/.test(norm)) return 'leaveCount';
  // Template typo "Name of accumulated leave" ≡ Number / LeaveCount
  if (/name\s+of\s+accumulated\s+leave/.test(norm)) return 'leaveCount';
  if (/accumulated\s+leave/.test(norm) && !/period|perod/.test(norm)) return 'leaveCount';
  if (/period|perod/.test(norm) && /accumulated/.test(norm) && /leave/.test(norm)) return 'periodBand';
  if (norm === 'from') return 'periodFrom';
  if (norm === 'till' || norm === 'to') return 'periodTill';
  return norm;
}

export function rowHasMeaningfulFormOGJGujaratExportData(row, headers) {
  const hdrs = resolveFormOGJGujaratTableHeaders(headers);
  return hdrs.some((header) => getFormOGJRowValueForHeader(row, header) !== '');
}

export function filterFormOGJGujaratExportRows(rows, headers) {
  const hdrs = resolveFormOGJGujaratTableHeaders(headers);
  return (Array.isArray(rows) ? rows : []).filter((row) =>
    rowHasMeaningfulFormOGJGujaratExportData(row, hdrs)
  );
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

function excelCellLooksLikeSerialHeader(text) {
  const t = formOGJGujaratHeaderNorm(text).replace(/[^a-z0-9 ]/g, ' ').replace(/\s+/g, ' ').trim();
  return /^sr\.?\s*no|^s\.?\s*no|serial|sl\.?\s*no/.test(t);
}

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

function collectWorksheetRowText(worksheet, row, maxCol = 20) {
  const parts = [];
  for (let c = 1; c <= maxCol; c += 1) {
    const t = excelCellValueToString(worksheet.getCell(row, c)?.value)
      .replace(/\r?\n+/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
    if (t) parts.push(t);
  }
  return parts.join(' ').replace(/\s+/g, ' ').trim();
}

/** Unmerge any ranges that intersect the data body so each employee gets its own row. */
function unmergeFormOGJDataRegion(worksheet, rowFrom, rowTo, colFrom, colTo) {
  if (!worksheet || typeof worksheet.unMergeCells !== 'function') return;
  const merges = Array.isArray(worksheet.model?.merges) ? [...worksheet.model.merges] : [];
  merges.forEach((label) => {
    const parts = String(label || '').split(':');
    if (parts.length !== 2) return;
    try {
      const tl = worksheet.getCell(parts[0]);
      const br = worksheet.getCell(parts[1]);
      const r1 = tl.fullAddress?.row ?? tl.row;
      const c1 = tl.fullAddress?.col ?? tl.col;
      const r2 = br.fullAddress?.row ?? br.row;
      const c2 = br.fullAddress?.col ?? br.col;
      const rowOverlap = r1 <= rowTo && r2 >= rowFrom;
      const colOverlap = c1 <= colTo && c2 >= colFrom;
      // Vertical merges into the body collapse multiple employees into one cell.
      if (rowOverlap && colOverlap && r2 > r1) {
        worksheet.unMergeCells(label);
      }
    } catch (_) {
      /* ignore invalid merge refs */
    }
  });
}

function findFormOGJFooterRow(worksheet, fromRow) {
  const start = Math.max(1, fromRow);
  const end = Math.max(start + 40, (worksheet.rowCount || start) + 5);
  for (let r = start; r <= end; r += 1) {
    const text = collectWorksheetRowText(worksheet, r, 30).toLowerCase();
    if (!text) continue;
    if (
      /system\s+generated/.test(text) ||
      /authorised\s+signatory|authorized\s+signatory/.test(text) ||
      /signature\s+of\s+employer/.test(text) ||
      /^date\s*:/.test(text) ||
      /^place\s*:/.test(text) ||
      /copy\s+to\s+workers/.test(text)
    ) {
      return r;
    }
  }
  return null;
}

function ensureFormOGJDataRowCapacity(worksheet, dataStartRow, rowCount, colFrom, colTo) {
  if (!worksheet || rowCount < 1) return dataStartRow;

  // Unmerge first — vertical merges make every write land on the same top-left cell
  // (last employee overwrites the first; download shows only Sr. No. 2).
  unmergeFormOGJDataRegion(
    worksheet,
    dataStartRow,
    dataStartRow + Math.max(rowCount, 2) + 5,
    colFrom,
    colTo
  );

  const footerRow = findFormOGJFooterRow(worksheet, dataStartRow);
  if (footerRow != null && footerRow > dataStartRow) {
    const available = footerRow - dataStartRow;
    const need = rowCount - available;
    if (need > 0 && typeof worksheet.spliceRows === 'function') {
      const blanks = Array.from({ length: need }, () => []);
      worksheet.spliceRows(footerRow, 0, ...blanks);
    }
  } else if (typeof worksheet.spliceRows === 'function' && rowCount > 1) {
    const blanks = Array.from({ length: rowCount - 1 }, () => []);
    worksheet.spliceRows(dataStartRow + 1, 0, ...blanks);
  }
  return dataStartRow;
}

function detectFormOGJGujaratTableLayout(worksheet, hints = {}) {
  const getMergeTopLeft = buildMergeTopLeftResolver(worksheet);
  const getMergeSpanEnd = buildMergeSpanEndResolver(worksheet, getMergeTopLeft);
  const getMergedAwareCellText = (r, c) => {
    const tl = getMergeTopLeft(r, c);
    return excelCellValueToString(worksheet.getCell(tl.r, tl.c)?.value).trim();
  };

  // Prefer sheet scan over modal hints — stale dataStartIndex was writing emp #1 into From/Till.
  let headerRow = -1;
  let startCol =
    hints.parsedTableStartCol != null && hints.parsedTableStartCol >= 0
      ? hints.parsedTableStartCol + 1
      : 1;

  const maxScanRows = Math.max(45, worksheet.rowCount + 5);
  for (let r = 1; r <= maxScanRows; r += 1) {
    const parts = [];
    for (let c = 1; c <= 20; c += 1) {
      const t = getMergedAwareCellText(r, c);
      if (t) parts.push(t.toLowerCase());
    }
    const joined = parts.join(' ');
    if (
      /sr\.?\s*no/.test(joined) &&
      /name\s+of\s+workers?/.test(joined) &&
      /accumulated\s+leave/.test(joined)
    ) {
      headerRow = r;
      for (let c = 1; c <= 20; c += 1) {
        if (excelCellLooksLikeSerialHeader(getMergedAwareCellText(r, c))) {
          startCol = c;
          break;
        }
      }
      break;
    }
  }
  if (headerRow < 1) return null;

  const templateCols = [];
  let subHeaderRow = headerRow;
  let c = startCol;
  while (c <= startCol + 12) {
    const label = getMergedAwareCellText(headerRow, c);
    if (!label) {
      if (templateCols.length >= 4) break;
      c += 1;
      continue;
    }
    const spanEnd = getMergeSpanEnd(headerRow, c);
    const endC = spanEnd.c;
    const bucket = formOGJHeaderAliasBucket(label);

    if (bucket === 'periodBand') {
      let fromCol = -1;
      let tillCol = -1;
      for (let sr = headerRow + 1; sr <= headerRow + 4; sr += 1) {
        for (let pc = c; pc <= endC; pc += 1) {
          const sub = formOGJHeaderAliasBucket(getMergedAwareCellText(sr, pc));
          if (sub === 'periodFrom') fromCol = pc;
          if (sub === 'periodTill') tillCol = pc;
        }
        if (fromCol > 0 && tillCol > 0) {
          subHeaderRow = Math.max(subHeaderRow, sr);
          break;
        }
      }
      if (fromCol < 0) fromCol = c;
      if (tillCol < 0) tillCol = endC > c ? endC : c + 1;
      templateCols.push({
        col: fromCol,
        bucket: 'periodFrom',
        label: `${FORM_OGJ_PERIOD_PARENT}_From`,
      });
      templateCols.push({
        col: tillCol,
        bucket: 'periodTill',
        label: `${FORM_OGJ_PERIOD_PARENT}_Till`,
      });
    } else if (bucket && bucket !== 'periodFrom' && bucket !== 'periodTill') {
      const normalizedLabel =
        bucket === 'leaveCount' ? 'Number of accumulated leave' : label;
      templateCols.push({ col: c, bucket, label: normalizedLabel });
    }
    c = endC + 1;
    if (templateCols.length >= 5) break;
  }
  if (templateCols.length < 4) return null;

  // Independent From/Till scan (all columns) — do not trust modal dataStartIndex.
  let fromTillRow = -1;
  for (let r = headerRow; r <= headerRow + 6; r += 1) {
    let hasFrom = false;
    let hasTill = false;
    for (let col = 1; col <= 25; col += 1) {
      const t = formOGJGujaratHeaderNorm(getMergedAwareCellText(r, col));
      if (t === 'from') hasFrom = true;
      if (t === 'till' || t === 'to') hasTill = true;
    }
    if (hasFrom || hasTill) {
      fromTillRow = r;
      subHeaderRow = Math.max(subHeaderRow, r);
    }
  }

  const dataStartRow =
    fromTillRow > 0 ? fromTillRow + 1 : Math.max(headerRow + 3, subHeaderRow + 1);

  return { headerRow, subHeaderRow, dataStartRow, templateCols, startCol };
}

/** Write employee rows into the Gujarat Form O template (preserves merged headers & column widths). */
export async function buildFormOGJGujaratWorkbookWithTemplateStyles({
  templateArrayBuffer,
  mappedData,
  headersToUse,
  parsedFormHeader,
  formFileName,
  headerFormData,
  parsedHeaderRowIndex,
  parsedDataStartIndex,
  parsedTableStartCol,
  formatStatutoryDateDisplay = null,
}) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(templateArrayBuffer);
  const worksheet = workbook.worksheets[0];
  if (!worksheet) throw new Error('Template worksheet not found.');

  const layout = detectFormOGJGujaratTableLayout(worksheet, {
    parsedHeaderRowIndex,
    parsedTableStartCol,
    // Intentionally ignore parsedDataStartIndex — stale modal index wrote emp #1 into From/Till.
  });
  if (!layout) throw new Error('Could not locate Gujarat Form O table header row.');

  let { dataStartRow, templateCols } = layout;
  const formatDate =
    typeof formatStatutoryDateDisplay === 'function'
      ? formatStatutoryDateDisplay
      : (v) => String(v || '').trim();

  if (headerFormData && typeof headerFormData === 'object') {
    writeStatutoryHeaderFieldsToExcelJsWorksheet(worksheet, {
      headerFormData,
      parsedFormHeader,
      headerRowEnd: Math.max(1, layout.headerRow - 1),
      maxScanCols: 80,
    });
  }

  const normalizedHeaders = resolveFormOGJGujaratTableHeaders(headersToUse);
  const rows = filterFormOGJGujaratExportRows(
    remapFormOGJGujaratRowsToHeaders(
      Array.isArray(mappedData) ? mappedData : [],
      normalizedHeaders,
      normalizedHeaders
    ),
    normalizedHeaders
  );

  const tableColMin = Math.min(...templateCols.map(({ col }) => col));
  const tableColMax = Math.max(...templateCols.map(({ col }) => col));

  dataStartRow = ensureFormOGJDataRowCapacity(
    worksheet,
    dataStartRow,
    Math.max(rows.length, 1),
    tableColMin,
    tableColMax
  );

  const clearToRow = Math.max(dataStartRow + rows.length + 10, dataStartRow + 15);
  for (let r = dataStartRow; r <= clearToRow; r += 1) {
    for (let col = tableColMin; col <= tableColMax; col += 1) {
      worksheet.getCell(r, col).value = '';
    }
  }

  rows.forEach((row, idx) => {
    const targetRow = worksheet.getRow(dataStartRow + idx);
    if (targetRow) targetRow.height = 18;
    templateCols.forEach(({ col, bucket, label }) => {
      let val = getFormOGJRowValueForHeader(row, label);
      if ((val == null || val === '') && bucket) {
        for (const [k, v] of Object.entries(row)) {
          if (
            formOGJHeaderAliasBucket(normHeaderLabel(k)) === bucket &&
            v != null &&
            String(v).trim() !== ''
          ) {
            val = String(v).trim();
            break;
          }
        }
      }
      // Always renumber after filter/remap so Sr. No. starts at 1 in the download.
      if (bucket === 'sno') val = String(idx + 1);
      const cell = worksheet.getCell(dataStartRow + idx, col);
      if (val == null || val === '') {
        cell.value = '';
        return;
      }
      if (bucket === 'periodFrom' || bucket === 'periodTill') {
        cell.value = formatDate(val);
      } else if (bucket === 'sno') {
        const n = Number(String(val).replace(/[,]/g, '').trim());
        cell.value = Number.isFinite(n) ? n : String(val);
      } else {
        cell.value = String(val);
      }
      cell.alignment = { ...(cell.alignment || {}), vertical: 'middle', wrapText: true };
    });
  });

  if (rows.length > 0) {
    ensureExcelJSDataRowsWithBorders(worksheet, {
      dataStartRow,
      dataRowCount: rows.length,
      colFrom: tableColMin,
      colTo: tableColMax,
      templateRow: dataStartRow,
      templateBodyRows: 1,
    });
  }

  const out = await workbook.xlsx.writeBuffer();
  const fileName =
    formFileName ||
    parsedFormHeader?.title?.replace(/[^a-zA-Z0-9]/g, '_') ||
    'Form_O_GJ_Gujarat.xlsx';
  return {
    blob: new Blob([out], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    }),
    fileName,
  };
}
