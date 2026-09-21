/** Gujarat Form L — notice of change in shift (See rule 14). */

import ExcelJS from 'exceljs';
import { ensureExcelJSDataRowsWithBorders } from '../../utils/excelTableBorders';
import { writeStatutoryHeaderFieldsToExcelJsWorksheet } from '../../utils/statutorySiteCompanyHeaders';

export const FORM_LGJ_DATE_BAND = 'Date of the Month';

export const FORM_LGJ_WEEKLY_HOLIDAY_DEFAULT = 'Saturday & Sunday';
export const FORM_LGJ_FIRST_SHIFT_FROM_DEFAULT = '9 AM';
export const FORM_LGJ_FIRST_SHIFT_TO_DEFAULT = '6 PM';

export const FORM_LGJ_GJ_CANONICAL_TABLE_HEADERS = [
  'Sr. No.',
  'Name of the Worker',
  'Designation',
  'Weekly holiday day',
  `${FORM_LGJ_DATE_BAND}_1st Shift_From`,
  `${FORM_LGJ_DATE_BAND}_1st Shift_To`,
  `${FORM_LGJ_DATE_BAND}_2nd Shift_From`,
  `${FORM_LGJ_DATE_BAND}_2nd Shift_To`,
  `${FORM_LGJ_DATE_BAND}_3rd Shift_From`,
  `${FORM_LGJ_DATE_BAND}_3rd Shift_To`,
];

export function formLGJGujaratHeaderNorm(txt) {
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
  return formLGJGujaratHeaderNorm(stripLeadingNumber(h))
    .replace(/[^a-z0-9 ]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

export function isFormLGJGujaratContext(
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
  if (/\bform\s*o\b|\bform_o\b|form[\s._-]*o[\s._-]*gj/.test(parts)) return false;
  if (/working\s+hours/.test(parts) && /total\s+days\s+worked/.test(parts)) return false;

  const hasGujarat = /gujarat|\b_gj\b|form[\s._-]*l[\s._-]*gj|form_l_gj/.test(parts);
  const hasFormL = /\bform[\s._-]*l\b/.test(parts);
  if (hasGujarat && hasFormL) return true;
  if (/\bform_l_gj\b/.test(parts)) return true;
  if (hasFormL && /see\s+rule\s+14/.test(parts)) return true;

  const joined = parts;
  return (
    /sr\.?\s*no/.test(joined) &&
    /name\s+of\s+the\s+worker/.test(joined) &&
    /designation/.test(joined) &&
    (/date\s+of\s+the\s+month|date\s+of\s+month/.test(joined) || /\d+(?:st|nd|rd|th)\s+shift/.test(joined))
  );
}

export function isFormLGJGujaratTableLayoutFormHeader(formHeader) {
  return !!formHeader?.formLGJGujaratTableLayout;
}

export function headersLookLikeBrokenFormLGJ(headers) {
  return (Array.isArray(headers) ? headers : []).some((h) => {
    const text = String(h || '');
    return (
      /date\s+of\s+the\s+month.*shift.*from/i.test(text) &&
      !/^Date of the Month_\d/i.test(text) &&
      text.split('_').length < 3
    );
  });
}

export function resolveFormLGJGujaratTableHeaders(tableHeaders) {
  const src = Array.isArray(tableHeaders) ? tableHeaders : [];
  if (headersLookLikeBrokenFormLGJ(src)) {
    return [...FORM_LGJ_GJ_CANONICAL_TABLE_HEADERS];
  }
  if (src.length >= 5) {
    const normalized = src.map((h) => normalizeFormLGJHeaderKey(h));
    if (normalized.filter(Boolean).length >= 5) return normalized;
  }
  return [...FORM_LGJ_GJ_CANONICAL_TABLE_HEADERS];
}

function normalizeFormLGJHeaderKey(header) {
  const raw = String(header || '').trim();
  if (!raw) return '';
  const n = normHeaderLabel(raw);
  if (/^sr\.?\s*no/.test(n)) return 'Sr. No.';
  if (/name\s+of\s+the\s+worker/.test(n)) return 'Name of the Worker';
  if (/^designation$/.test(n)) return 'Designation';
  if (/weekly\s+holiday/.test(n)) return 'Weekly holiday day';

  const triple = raw.match(/^(.+?)_(.+?)_(From|To|Till)$/i);
  if (triple) {
    const band = /date\s+of\s+the\s+month|date\s+of\s+month/i.test(triple[1])
      ? FORM_LGJ_DATE_BAND
      : triple[1].trim();
    const shift = triple[2].trim();
    const field = /^to$/i.test(triple[3]) || /^till$/i.test(triple[3]) ? 'To' : 'From';
    return `${band}_${shift}_${field}`;
  }

  const broken = raw.match(/date\s+of\s+the\s+month\s+(\d+(?:st|nd|rd|th)\s+shift)\s+from\s*[-–—]?\s*to\s*[-–—]?\s*_?\1/i);
  if (broken) {
    return `${FORM_LGJ_DATE_BAND}_${broken[1]}_From`;
  }

  return raw;
}

export function parseFormLGJHeaderKey(header) {
  const normalized = normalizeFormLGJHeaderKey(header);
  const m = String(normalized || '').match(/^(.+?)_(.+?)_(From|To)$/i);
  if (!m) return null;
  return { band: m[1].trim(), shift: m[2].trim(), field: m[3].trim() };
}

export function buildFormLGJSubColumnsFromHeaders(headers) {
  const subs = {};
  (headers || []).forEach((h) => {
    const parsed = parseFormLGJHeaderKey(h);
    if (!parsed) return;
    if (!subs[parsed.band]) subs[parsed.band] = [];
    if (!subs[parsed.band].includes(parsed.shift)) subs[parsed.band].push(parsed.shift);
  });
  return subs;
}

const isEmployerMetaLabel = (txt) => {
  const t = formLGJGujaratHeaderNorm(txt);
  return (
    /name\s+and\s+address\s+of\s+the\s+establishment/.test(t) ||
    /name\s+of\s+the\s+authorized/.test(t) ||
    /^notice$/.test(t) ||
    /^date\s*:?$/.test(t) ||
    /^place\s*:?$/.test(t)
  );
};

export function parseFormLGJShiftCellValue(raw) {
  const text = String(raw || '').trim();
  if (!text) return { from: '', to: '' };
  const lines = text
    .split(/\r?\n/)
    .map((s) => s.trim())
    .filter(Boolean);
  if (lines.length >= 2) return { from: lines[0], to: lines[1] };
  const dash = text.match(/^(.+?)\s*[-–—]\s*(.+)$/);
  if (dash) return { from: dash[1].trim(), to: dash[2].trim() };
  return { from: text, to: '' };
}

export function formatFormLGJShiftCellForExport(from, to) {
  const f = String(from || '').trim();
  const t = String(to || '').trim();
  if (f && t) return `${f}\n${t}`;
  return f || t || '';
}

function shiftHeaderKeys(shift) {
  return {
    from: `${FORM_LGJ_DATE_BAND}_${shift}_From`,
    to: `${FORM_LGJ_DATE_BAND}_${shift}_To`,
  };
}

export function getFormLGJShiftValuesForExport(row, shift) {
  const keys = shiftHeaderKeys(shift);
  return {
    from: getFormLGJRowValueForHeader(row, keys.from),
    to: getFormLGJRowValueForHeader(row, keys.to),
  };
}

function normalizeShiftLabel(raw) {
  const text = String(raw || '')
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  const m = text.match(/^(\d+(?:st|nd|rd|th))\s+shift$/i);
  if (m) return `${m[1]} Shift`;
  return text;
}

/**
 * Parse Form L GJ Excel header rows (identity cols + Date of the Month / shift / From-To).
 */
export function rebuildFormLGJGujaratTableHeadersFromSheet({
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
  for (let r = 0; r < Math.min(40, jsonData.length); r += 1) {
    const parts = [];
    for (let c = 0; c < Math.max(effectiveSheetCols, 12); c += 1) {
      const t = normLower(getMergedAwareCellText(r, c));
      if (t) parts.push(t);
    }
    const joined = parts.join(' ');
    if (
      /sr\.?\s*no/.test(joined) &&
      /name\s+of\s+the\s+worker/.test(joined) &&
      /designation/.test(joined)
    ) {
      headerRowIndex = r;
      break;
    }
  }
  if (headerRowIndex < 0) return null;

  let startCol = -1;
  for (let c = 0; c < Math.max(effectiveSheetCols, 12); c += 1) {
    const t = normLower(getMergedAwareCellText(headerRowIndex, c));
    if (/^sr\.?\s*no\.?$/.test(t) || t === 'sr no' || t.startsWith('sr. no')) {
      startCol = c;
      break;
    }
  }
  if (startCol < 0) startCol = 0;

  const readIdentityLabel = (c) => {
    for (let r = headerRowIndex; r <= headerRowIndex + 3; r += 1) {
      const t = norm(getMergedAwareCellText(r, c));
      if (!t || isEmployerMetaLabel(t)) continue;
      const n = normHeaderLabel(t);
      if (/^sr\.?\s*no/.test(n)) return 'Sr. No.';
      if (/name\s+of\s+the\s+worker/.test(n)) return 'Name of the Worker';
      if (/^designation$/.test(n)) return 'Designation';
      if (/weekly\s+holiday/.test(n)) return 'Weekly holiday day';
    }
    return '';
  };

  const flatHeaders = [];
  const subColumnsData = {};
  const mainHeaders = [];
  let subHeaderRowIndex = headerRowIndex + 1;
  let dataStartIndex = headerRowIndex + 3;
  let weeklyHolidayCol = -1;

  let c = startCol;
  const maxCols = Math.max(effectiveSheetCols, (jsonData[headerRowIndex] || []).length, startCol + 10);
  while (c < startCol + 5 && c < maxCols) {
    const label = readIdentityLabel(c);
    if (label) {
      flatHeaders.push(label);
      mainHeaders.push(label);
      if (label === 'Weekly holiday day') weeklyHolidayCol = c;
    } else {
      // Stop identity scan once we hit Date-of-Month / shift band cells.
      const probe = normLower(getMergedAwareCellText(headerRowIndex, c));
      if (/date\s+of\s+the\s+month|date\s+of\s+month|\d+(?:st|nd|rd|th)\s+shift/.test(probe)) break;
    }
    c = Math.max(c + 1, getMergeSpanEndCol(headerRowIndex, c));
  }

  const shiftBands = [];
  const shiftBandKeys = new Set();
  for (let r = headerRowIndex; r <= headerRowIndex + 4; r += 1) {
    for (let col = startCol; col < maxCols; col += 1) {
      const raw = norm(getMergedAwareCellText(r, col));
      if (!/^\d+(?:st|nd|rd|th)\s+shift$/i.test(raw)) continue;
      const endC = getMergeSpanEndCol(r, col) - 1;
      const key = `${col}:${r}`;
      if (shiftBandKeys.has(key)) continue;
      shiftBandKeys.add(key);
      shiftBands.push({
        shift: normalizeShiftLabel(raw),
        row: r,
        startCol: col,
        endCol: endC,
      });
    }
  }

  shiftBands.sort((a, b) => a.startCol - b.startCol);

  if (shiftBands.length === 0) {
    for (let col = startCol + 3; col < maxCols; col += 1) {
      const parent = norm(getMergedAwareCellText(headerRowIndex, col));
      if (!/date\s+of\s+the\s+month|date\s+of\s+month/i.test(parent)) continue;
      const endC = getMergeSpanEndCol(headerRowIndex, col) - 1;
      const width = Math.max(2, endC - col + 1);
      const shiftCount = Math.floor(width / 2);
      for (let s = 1; s <= Math.max(shiftCount, 2); s += 1) {
        const ordinal =
          s === 1 ? '1st Shift' : s === 2 ? '2nd Shift' : s === 3 ? '3rd Shift' : `${s}th Shift`;
        shiftBands.push({
          shift: ordinal,
          row: headerRowIndex + 1,
          startCol: col + (s - 1) * 2,
          endCol: col + (s - 1) * 2 + 1,
        });
      }
      break;
    }
  }

  if (!subColumnsData[FORM_LGJ_DATE_BAND]) subColumnsData[FORM_LGJ_DATE_BAND] = [];
  const shiftColumnMap = [];

  shiftBands.forEach((band) => {
    const bandWidth = band.endCol - band.startCol + 1;
    let fromCol = -1;
    let toCol = -1;
    let singleColumnShift = bandWidth === 1;

    for (let r = band.row + 1; r <= band.row + 3; r += 1) {
      for (let col = band.startCol; col <= band.endCol; col += 1) {
        const subRaw = norm(getMergedAwareCellText(r, col));
        const sub = normLower(subRaw);
        if (!sub) continue;
        if (/from/.test(sub) && (/to|till/.test(sub) || sub.includes('-'))) {
          fromCol = col;
          toCol = col;
          singleColumnShift = true;
          subHeaderRowIndex = Math.max(subHeaderRowIndex, r);
        } else {
          if (sub === 'from' || /^from\b/.test(sub)) fromCol = col;
          if (sub === 'to' || sub === 'till' || /^to\b/.test(sub)) toCol = col;
          if (fromCol >= 0 || toCol >= 0) subHeaderRowIndex = Math.max(subHeaderRowIndex, r);
        }
      }
      if (fromCol >= 0 && toCol >= 0) break;
    }

    if (fromCol < 0) fromCol = band.startCol;
    if (toCol < 0) toCol = band.endCol;
    if (singleColumnShift || fromCol === toCol) {
      toCol = fromCol;
      singleColumnShift = true;
    } else if (toCol <= fromCol && bandWidth > 1) {
      toCol = Math.min(band.endCol, fromCol + 1);
    }

    flatHeaders.push(`${FORM_LGJ_DATE_BAND}_${band.shift}_From`);
    flatHeaders.push(`${FORM_LGJ_DATE_BAND}_${band.shift}_To`);
    if (!subColumnsData[FORM_LGJ_DATE_BAND].includes(band.shift)) {
      subColumnsData[FORM_LGJ_DATE_BAND].push(band.shift);
    }
    shiftColumnMap.push({
      shift: band.shift,
      physicalCol: band.startCol,
      singleColumn: singleColumnShift,
      fromCol,
      toCol,
    });
  });

  if (weeklyHolidayCol < 0) {
    for (let col = startCol; col < maxCols; col += 1) {
      for (let r = headerRowIndex; r <= headerRowIndex + 4; r += 1) {
        const t = normLower(getMergedAwareCellText(r, col));
        if (/weekly\s+holiday/.test(t)) {
          weeklyHolidayCol = col;
          break;
        }
      }
      if (weeklyHolidayCol >= 0) break;
    }
  }
  if (weeklyHolidayCol >= 0 && !flatHeaders.includes('Weekly holiday day')) {
    const desigIdx = flatHeaders.indexOf('Designation');
    const insertAt = desigIdx >= 0 ? desigIdx + 1 : Math.min(3, flatHeaders.length);
    flatHeaders.splice(insertAt, 0, 'Weekly holiday day');
    mainHeaders.splice(insertAt, 0, 'Weekly holiday day');
  }

  if (flatHeaders.length < 5) {
    return {
      headers: mainHeaders,
      expandedHeaders: resolveFormLGJGujaratTableHeaders(),
      subColumnsData: buildFormLGJSubColumnsFromHeaders(FORM_LGJ_GJ_CANONICAL_TABLE_HEADERS),
      headerRowIndex,
      startCol,
      dataStartIndex: headerRowIndex + 3,
      shiftColumnMap: [],
    };
  }

  dataStartIndex = Math.max(headerRowIndex + 3, subHeaderRowIndex + 1);

  return {
    headers: mainHeaders,
    expandedHeaders: flatHeaders.map((h) => normalizeFormLGJHeaderKey(h)),
    subColumnsData,
    headerRowIndex,
    startCol,
    dataStartIndex,
    shiftColumnMap,
    weeklyHolidayCol,
  };
}

export function remapFormLGJGujaratRowsToHeaders(rows, sourceHeaders, targetHeaders) {
  const src = Array.isArray(sourceHeaders) ? sourceHeaders : [];
  const tgt = resolveFormLGJGujaratTableHeaders(targetHeaders);
  if (!Array.isArray(rows)) return [];

  const bucketFor = (header) => {
    const n = normHeaderLabel(header);
    if (/^sr\.?\s*no/.test(n)) return 'sno';
    if (/name\s+of\s+the\s+worker/.test(n)) return 'workerName';
    if (/^designation$/.test(n)) return 'designation';
    if (/weekly\s+holiday/.test(n)) return 'weeklyHoliday';
    const parsed = parseFormLGJHeaderKey(header);
    if (parsed) return `shift:${parsed.shift.toLowerCase()}:${parsed.field.toLowerCase()}`;
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

export function isFormLGJSerialHeader(header) {
  return /^sr\.?\s*no/i.test(normHeaderLabel(header));
}

export function isFormLGJWorkerNameHeader(header) {
  return /name\s+of\s+the\s+worker/.test(normHeaderLabel(header));
}

/** Title-case worker names for Form L Excel (lookup helpers store lowercase). */
export function toFormLGJPersonNameDisplay(value) {
  const s = String(value || '')
    .trim()
    .replace(/\s+/g, ' ');
  if (!s) return '';
  return s
    .split(' ')
    .map((word) => {
      if (!word) return word;
      return word.charAt(0).toUpperCase() + word.slice(1).toLowerCase();
    })
    .join(' ');
}

export function isFormLGJDesignationHeader(header) {
  return /^designation$/.test(normHeaderLabel(header));
}

export function isFormLGJShiftDateHeader(header) {
  return !!parseFormLGJHeaderKey(header);
}

export function isFormLGJFirstShiftFromHeader(header) {
  const parsed = parseFormLGJHeaderKey(header);
  return !!parsed && /^1st shift$/i.test(parsed.shift) && /^from$/i.test(parsed.field);
}

export function isFormLGJFirstShiftToHeader(header) {
  const parsed = parseFormLGJHeaderKey(header);
  return !!parsed && /^1st shift$/i.test(parsed.shift) && /^to$/i.test(parsed.field);
}

/** 1st Shift From — default 9 AM (Form L rule 14 notice). */
export function resolveFormLGJFirstShiftFromValue(_emp, _readShiftStart, _formatShiftTime) {
  return FORM_LGJ_FIRST_SHIFT_FROM_DEFAULT;
}

/** 1st Shift To — default 6 PM (Form L rule 14 notice). */
export function resolveFormLGJFirstShiftToValue(_emp, _readShiftEnd, _formatShiftTime) {
  return FORM_LGJ_FIRST_SHIFT_TO_DEFAULT;
}

/** Weekly holiday day — default Saturday & Sunday (Form L rule 14 notice). */
export function resolveFormLGJWeeklyHolidayValue() {
  return FORM_LGJ_WEEKLY_HOLIDAY_DEFAULT;
}

export function isFormLGJWeeklyHolidayHeader(header) {
  return /weekly\s+holiday/.test(normHeaderLabel(header));
}

export function getFormLGJRowValueForHeader(row, header) {
  if (!row || !header) return '';
  const direct = row[header];
  if (direct != null && String(direct).trim() !== '') return String(direct).trim();

  const bucketFor = (h) => {
    if (isFormLGJSerialHeader(h)) return 'sno';
    if (isFormLGJWorkerNameHeader(h)) return 'workerName';
    if (isFormLGJDesignationHeader(h)) return 'designation';
    if (isFormLGJWeeklyHolidayHeader(h)) return 'weeklyHoliday';
    const parsed = parseFormLGJHeaderKey(h);
    if (parsed) return `shift:${parsed.shift.toLowerCase()}:${parsed.field.toLowerCase()}`;
    return normHeaderLabel(h);
  };
  const bucket = bucketFor(header);
  for (const [k, v] of Object.entries(row)) {
    if (String(k).startsWith('__')) continue;
    if (bucketFor(k) === bucket && v != null && String(v).trim() !== '') {
      return String(v).trim();
    }
  }
  if (bucket === 'weeklyHoliday') return FORM_LGJ_WEEKLY_HOLIDAY_DEFAULT;
  return '';
}

export function enrichFormLGJGujaratDisplayHeader(formHeader, fileName, rowItem, tableHeaders) {
  const ctx = isFormLGJGujaratContext(formHeader, rowItem, fileName, '', tableHeaders);
  if (!ctx) return formHeader;
  return {
    ...(formHeader || {}),
    title: formHeader?.title || 'FORM - L',
    subtitle: formHeader?.subtitle || '(See rule 14)',
    formLGJGujaratTableLayout: true,
    textRows: Array.isArray(formHeader?.textRows) ? formHeader.textRows : [],
  };
}

/** Collect notice paragraph lines above the table (for UI / export). */
export function extractFormLGJNoticeTextRows(jsonData, headerRowIndex, getMergedAwareCellText, maxCol = 10) {
  if (!Array.isArray(jsonData) || headerRowIndex <= 0) return [];
  const lines = [];
  for (let r = 0; r < headerRowIndex; r += 1) {
    const parts = [];
    for (let c = 0; c < maxCol; c += 1) {
      const t = String(getMergedAwareCellText(r, c) || '')
        .replace(/\r?\n+/g, ' ')
        .replace(/\s+/g, ' ')
        .trim();
      if (t) parts.push(t);
    }
    const rowText = parts.join(' ').replace(/\s+/g, ' ').trim();
    if (!rowText) continue;
    if (/^form\s*[-–]?\s*l$/i.test(rowText)) continue;
    if (/see rule\s*14/i.test(rowText)) continue;
    if (/list of workers engaged in shift/i.test(rowText)) continue;
    if (FORM_LGJ_NOTICE_LINE_PATTERNS.some((p) => p.test(rowText)) || rowText.length >= 35) {
      lines.push(rowText);
    }
  }
  return lines;
}

/** Chunk model for 3-row table header: identity cols + Date band + shift bands + From/To. */
export function buildFormLGJHeaderChunkModel(headers) {
  const hdrs = resolveFormLGJGujaratTableHeaders(headers);
  const identityBefore = [];
  const identityAfter = [];
  const dateCols = [];
  let seenDate = false;
  hdrs.forEach((header, index) => {
    const parsed = parseFormLGJHeaderKey(header);
    if (parsed) {
      seenDate = true;
      dateCols.push({ index, header, ...parsed });
      return;
    }
    if (!seenDate) identityBefore.push({ index, header });
    else identityAfter.push({ index, header });
  });
  if (dateCols.length === 0) return null;

  const shiftGroups = [];
  dateCols.forEach((col) => {
    let group = shiftGroups.find((g) => g.shift.toLowerCase() === col.shift.toLowerCase());
    if (!group) {
      group = { shift: col.shift, cols: [] };
      shiftGroups.push(group);
    }
    group.cols.push(col);
  });

  shiftGroups.forEach((group) => {
    group.cols.sort((a, b) => {
      const order = { from: 0, to: 1 };
      return (order[a.field.toLowerCase()] ?? 9) - (order[b.field.toLowerCase()] ?? 9);
    });
  });

  return {
    identity: [...identityBefore, ...identityAfter],
    identityBefore,
    identityAfter,
    shiftGroups,
    dateBand: FORM_LGJ_DATE_BAND,
  };
}

export function rowHasMeaningfulFormLGJGujaratExportData(row, headers) {
  const hdrs = resolveFormLGJGujaratTableHeaders(headers);
  return hdrs.some((header) => getFormLGJRowValueForHeader(row, header) !== '');
}

export function filterFormLGJGujaratExportRows(rows, headers) {
  const hdrs = resolveFormLGJGujaratTableHeaders(headers);
  return (Array.isArray(rows) ? rows : []).filter((row) =>
    rowHasMeaningfulFormLGJGujaratExportData(row, hdrs)
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
  const t = formLGJGujaratHeaderNorm(text).replace(/[^a-z0-9 ]/g, ' ').replace(/\s+/g, ' ').trim();
  return /^sr\.?\s*no|^s\.?\s*no|serial|sl\.?\s*no/.test(t);
}

function formLGJHeaderAliasBucket(n) {
  const norm = normHeaderLabel(n);
  if (/^sr\.?\s*no/.test(norm)) return 'sno';
  if (/name\s+of\s+the\s+worker/.test(norm)) return 'workerName';
  if (/^designation$/.test(norm)) return 'designation';
  if (/weekly\s+holiday/.test(norm)) return 'weeklyHoliday';
  return norm;
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

function detectFormLGJGujaratTableLayout(worksheet, hints = {}) {
  const getMergeTopLeft = buildMergeTopLeftResolver(worksheet);
  const getMergeSpanEnd = buildMergeSpanEndResolver(worksheet, getMergeTopLeft);
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

  if (headerRow < 1) {
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
        /name\s+of\s+the\s+worker/.test(joined) &&
        /designation/.test(joined)
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
  }
  if (headerRow < 1) return null;

  const templateCols = [];
  let subHeaderRow = headerRow;

  const readIdentityAt = (c) => {
    for (let r = headerRow; r <= headerRow + 4; r += 1) {
      const t = getMergedAwareCellText(r, c);
      const bucket = formLGJHeaderAliasBucket(t);
      if (bucket === 'sno') return { col: c, bucket, label: 'Sr. No.' };
      if (bucket === 'workerName') return { col: c, bucket, label: 'Name of the Worker' };
      if (bucket === 'designation') return { col: c, bucket, label: 'Designation' };
      if (bucket === 'weeklyHoliday') return { col: c, bucket, label: 'Weekly holiday day' };
    }
    return null;
  };

  const shiftBands = [];
  for (let r = headerRow; r <= headerRow + 4; r += 1) {
    for (let c = startCol; c <= startCol + 14; c += 1) {
      const raw = getMergedAwareCellText(r, c);
      if (!/^\d+(?:st|nd|rd|th)\s+shift$/i.test(raw)) continue;
      const end = getMergeSpanEnd(r, c);
      if (!shiftBands.some((b) => b.startCol === c && b.row === r)) {
        shiftBands.push({
          shift: normalizeShiftLabel(raw),
          row: r,
          startCol: c,
          endCol: end.c,
        });
      }
    }
  }
  shiftBands.sort((a, b) => a.startCol - b.startCol);

  const identityCols = new Set();
  for (let c = startCol; c < startCol + 4; c += 1) {
    const id = readIdentityAt(c);
    if (id && !identityCols.has(id.bucket)) {
      identityCols.add(id.bucket);
      templateCols.push(id);
    }
  }

  shiftBands.forEach((band) => {
    const bandWidth = band.endCol - band.startCol + 1;
    let fromCol = -1;
    let toCol = -1;
    let singleColumn = bandWidth === 1;

    for (let r = band.row + 1; r <= band.row + 3; r += 1) {
      for (let col = band.startCol; col <= band.endCol; col += 1) {
        const sub = formLGJGujaratHeaderNorm(getMergedAwareCellText(r, col));
        if (/from/.test(sub) && (/to|till/.test(sub) || sub.includes('-'))) {
          fromCol = col;
          toCol = col;
          singleColumn = true;
          subHeaderRow = Math.max(subHeaderRow, r);
        } else {
          if (sub === 'from' || /^from\b/.test(sub)) fromCol = col;
          if (sub === 'to' || sub === 'till' || /^to\b/.test(sub)) toCol = col;
          if (fromCol >= 0 || toCol >= 0) subHeaderRow = Math.max(subHeaderRow, r);
        }
      }
      if (fromCol >= 0 && toCol >= 0) break;
    }

    if (fromCol < 0) fromCol = band.startCol;
    if (toCol < 0) toCol = band.endCol;
    if (singleColumn || fromCol === toCol) {
      templateCols.push({
        col: band.startCol,
        bucket: 'shiftCombined',
        shift: band.shift,
        label: `${FORM_LGJ_DATE_BAND}_${band.shift}`,
      });
    } else {
      templateCols.push({
        col: fromCol,
        bucket: 'shiftFrom',
        shift: band.shift,
        label: `${FORM_LGJ_DATE_BAND}_${band.shift}_From`,
      });
      templateCols.push({
        col: toCol,
        bucket: 'shiftTo',
        shift: band.shift,
        label: `${FORM_LGJ_DATE_BAND}_${band.shift}_To`,
      });
    }
  });

  for (let c = startCol; c <= startCol + 14; c += 1) {
    const id = readIdentityAt(c);
    if (id?.bucket === 'weeklyHoliday' && !templateCols.some((t) => t.bucket === 'weeklyHoliday')) {
      templateCols.push(id);
    }
  }

  if (templateCols.length < 4) return null;

  const dataStartRow =
    hints.parsedDataStartIndex != null && hints.parsedDataStartIndex >= 0
      ? hints.parsedDataStartIndex + 1
      : Math.max(headerRow + 3, subHeaderRow + 1);

  templateCols.sort((a, b) => a.col - b.col);

  return { headerRow, subHeaderRow, dataStartRow, templateCols, startCol };
}

const FORM_LGJ_NOTICE_LINE_PATTERNS = [
  /workers in establishment/i,
  /shift schedule|shift scehdule/i,
  /schedule for the month/i,
];

function collectWorksheetRowText(worksheet, row, maxCol) {
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

/** Merge notice paragraphs across the table width — single line, no narrow-column wrap. */
function writeFormLGJGujaratNoticeAndColumnLayout(worksheet, layout, parsedFormHeader) {
  if (!worksheet || !layout) return;
  const { headerRow, startCol, templateCols } = layout;
  const tableColMax = Math.max(...templateCols.map(({ col }) => col), startCol + 6);
  const mergeFrom = Math.max(2, startCol);
  const mergeTo = Math.max(tableColMax, 8);

  const writeMergedNoticeLine = (r, text) => {
    const line = String(text || '').replace(/\r?\n+/g, ' ').replace(/\s+/g, ' ').trim();
    if (!line) return;
    for (let c = 1; c <= mergeTo + 2; c += 1) {
      worksheet.getCell(r, c).value = '';
    }
    try {
      worksheet.unMergeCells(r, mergeFrom, r, mergeTo);
    } catch (_) {
      /* not merged */
    }
    try {
      worksheet.mergeCells(r, mergeFrom, r, mergeTo);
    } catch (_) {
      /* template merge */
    }
    const cell = worksheet.getCell(r, mergeFrom);
    cell.value = line;
    cell.alignment = {
      horizontal: 'left',
      vertical: 'middle',
      wrapText: false,
      shrinkToFit: false,
    };
    const wsRow = worksheet.getRow(r);
    if (wsRow) wsRow.height = 18;
  };

  for (let r = 1; r < headerRow; r += 1) {
    const rowText = collectWorksheetRowText(worksheet, r, mergeTo + 2);
    if (!rowText) continue;
    if (/^form\s*[-–]?\s*l$/i.test(rowText)) continue;
    if (/see rule\s*14/i.test(rowText)) continue;
    if (/list of workers engaged in shift/i.test(rowText)) continue;

    const isNotice = FORM_LGJ_NOTICE_LINE_PATTERNS.some((p) => p.test(rowText));
    if (!isNotice && rowText.length < 35) continue;

    writeMergedNoticeLine(r, rowText);
  }

  const savedLines = Array.isArray(parsedFormHeader?.textRows)
    ? parsedFormHeader.textRows.filter(Boolean)
    : [];
  if (savedLines.length > 0) {
    let lineIdx = 0;
    for (let r = 1; r < headerRow && lineIdx < savedLines.length; r += 1) {
      const rowText = collectWorksheetRowText(worksheet, r, mergeTo + 2);
      if (!rowText) continue;
      if (/^form\s*[-–]?\s*l$/i.test(rowText)) continue;
      if (/see rule\s*14/i.test(rowText)) continue;
      if (/list of workers engaged in shift/i.test(rowText)) continue;
      if (FORM_LGJ_NOTICE_LINE_PATTERNS.some((p) => p.test(rowText)) || rowText.length >= 35) {
        writeMergedNoticeLine(r, savedLines[lineIdx] || rowText);
        lineIdx += 1;
      }
    }
  }

  templateCols.forEach(({ col, bucket }) => {
    const column = worksheet.getColumn(col);
    if (!column) return;
    if (bucket === 'sno') column.width = Math.max(column.width || 0, 8);
    else if (bucket === 'workerName') column.width = Math.max(column.width || 0, 26);
    else if (bucket === 'designation') column.width = Math.max(column.width || 0, 22);
    else if (bucket === 'shiftCombined' || bucket === 'shiftFrom' || bucket === 'shiftTo') {
      column.width = Math.max(column.width || 0, 16);
    } else if (bucket === 'weeklyHoliday') column.width = Math.max(column.width || 0, 18);
  });
}

function applyFormLGJDataCellStyle(cell, bucket, val) {
  const text = String(val ?? '');
  const isShift = bucket === 'shiftCombined' || bucket === 'shiftFrom' || bucket === 'shiftTo';
  cell.alignment = {
    ...(cell.alignment || {}),
    vertical: 'middle',
    horizontal: isShift ? 'center' : bucket === 'sno' ? 'center' : 'left',
    wrapText: isShift && text.includes('\n'),
    shrinkToFit: false,
  };
}

function estimateFormLGJDataRowHeight(row, templateCols) {
  let height = 24;
  templateCols.forEach(({ bucket, shift, label }) => {
    if (bucket === 'shiftCombined' && shift) {
      const { from, to } = getFormLGJShiftValuesForExport(row, shift);
      const combined = formatFormLGJShiftCellForExport(from, to);
      if (combined.includes('\n')) height = Math.max(height, 38);
    } else if (bucket === 'workerName' || bucket === 'designation') {
      const val = getFormLGJRowValueForHeader(row, label);
      if (String(val).length > 18) height = Math.max(height, 28);
    }
  });
  return height;
}

/** Write rows into the Gujarat Form L template (preserves merged headers & column layout). */
export async function buildFormLGJGujaratWorkbookWithTemplateStyles({
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
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(templateArrayBuffer);
  const worksheet = workbook.worksheets[0];
  if (!worksheet) throw new Error('Template worksheet not found.');

  const layout = detectFormLGJGujaratTableLayout(worksheet, {
    parsedHeaderRowIndex,
    parsedDataStartIndex,
    parsedTableStartCol,
  });
  if (!layout) throw new Error('Could not locate Gujarat Form L table header row.');

  const { dataStartRow, templateCols } = layout;

  writeFormLGJGujaratNoticeAndColumnLayout(worksheet, layout, parsedFormHeader);

  if (headerFormData && typeof headerFormData === 'object') {
    writeStatutoryHeaderFieldsToExcelJsWorksheet(worksheet, {
      headerFormData,
      parsedFormHeader,
      headerRowEnd: Math.max(1, layout.headerRow - 1),
      maxScanCols: 80,
    });
  }

  const normalizedHeaders = resolveFormLGJGujaratTableHeaders(headersToUse);
  const rows = filterFormLGJGujaratExportRows(
    remapFormLGJGujaratRowsToHeaders(
      Array.isArray(mappedData) ? mappedData : [],
      normalizedHeaders,
      normalizedHeaders
    ),
    normalizedHeaders
  );

  const tableColMin = Math.min(...templateCols.map(({ col }) => col));
  const tableColMax = Math.max(...templateCols.map(({ col }) => col));
  const clearToRow = Math.max(dataStartRow + rows.length + 10, dataStartRow + 15);
  for (let r = dataStartRow; r <= clearToRow; r += 1) {
    for (let col = tableColMin; col <= tableColMax; col += 1) {
      worksheet.getCell(r, col).value = '';
    }
  }

  rows.forEach((row, idx) => {
    const targetRow = worksheet.getRow(dataStartRow + idx);
    const rowHeight = estimateFormLGJDataRowHeight(row, templateCols);
    if (targetRow) targetRow.height = rowHeight;
    templateCols.forEach(({ col, bucket, shift, label }) => {
      let val = '';
      if (bucket === 'shiftCombined' && shift) {
        const { from, to } = getFormLGJShiftValuesForExport(row, shift);
        val = formatFormLGJShiftCellForExport(from, to);
      } else if (bucket === 'shiftFrom' && shift) {
        val = getFormLGJRowValueForHeader(row, `${FORM_LGJ_DATE_BAND}_${shift}_From`);
      } else if (bucket === 'shiftTo' && shift) {
        val = getFormLGJRowValueForHeader(row, `${FORM_LGJ_DATE_BAND}_${shift}_To`);
      } else {
        val = getFormLGJRowValueForHeader(row, label);
        if ((val == null || val === '') && bucket) {
          for (const [k, v] of Object.entries(row)) {
            if (
              formLGJHeaderAliasBucket(normHeaderLabel(k)) === bucket &&
              v != null &&
              String(v).trim() !== ''
            ) {
              val = String(v).trim();
              break;
            }
          }
        }
      }
      if ((val == null || val === '') && bucket === 'sno') val = String(idx + 1);
      const cell = worksheet.getCell(dataStartRow + idx, col);
      if (val == null || val === '') {
        cell.value = '';
        return;
      }
      if (bucket === 'sno') {
        const n = Number(String(val).replace(/[,]/g, '').trim());
        cell.value = Number.isFinite(n) ? n : String(val);
      } else if (bucket === 'workerName') {
        // Autofill lookup names are lowercased; write Title Case in Excel.
        cell.value = toFormLGJPersonNameDisplay(val);
      } else {
        cell.value = String(val);
      }
      applyFormLGJDataCellStyle(cell, bucket, val);
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
    'Form_L_GJ_Gujarat.xlsx';
  return {
    blob: new Blob([out], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    }),
    fileName,
  };
}
