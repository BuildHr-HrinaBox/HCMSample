/** Form No. 25 (AP) — Muster Roll [Rule 103]. */

export const FORM_25_AP_MUSTER_PREFIX_COUNT = 8;

export const FORM_25_AP_REMARKS_HEADER = 'Remarks';

/** Excel AP template splits days 1–15 under "For the period ending", 16–31 under month-year band. */
export const FORM_25_AP_FIRST_DAY_BAND_SIZE = 15;

export const FORM_25_AP_FACTORY_HEADER_SPECS = [
  {
    key: 'form25_ap_header_factory',
    label: 'Name of the factory:',
    test: (raw) => /^name\s+of\s+the\s+factory\s*:?\s*$/i.test(raw)
  },
  {
    key: 'form25_ap_header_place',
    label: 'Place:',
    test: (raw) => /^place\s*:?\s*$/i.test(raw)
  },
  {
    key: 'form25_ap_header_district',
    label: 'District:',
    test: (raw) => /^district\s*:?\s*$/i.test(raw)
  }
];

export function form25APMusterHeaderNorm(txt) {
  return String(txt || '')
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

/** Split merged Excel title block into centered lines like the template. */
export function splitForm25HeaderBlock(text) {
  const t = String(text || '')
    .replace(/\r\n/g, '\n')
    .replace(/\s+\n/g, '\n')
    .replace(/\n\s+/g, '\n')
    .trim();
  if (!t) return null;
  if (!/\bform[\s._-]*25\b/i.test(t) && !/\bfrom\s+no\.?\s*25\b/i.test(t)) return null;
  if (!/muster\s+roll/i.test(t)) return null;

  const titleMatch = t.match(/\b(FORM|FROM)\s+NO\.?\s*25\b/i);
  const title = titleMatch
    ? titleMatch[0].replace(/\s+/g, ' ').trim().toUpperCase().replace(/^FROM/i, 'FORM')
    : 'FORM NO. 25';

  const bracketMatch = t.match(/\[Prescribed under Rule\s+\d+\]/i);
  const referenceBracket = bracketMatch ? bracketMatch[0] : '';

  const subtitle = /\bMuster Roll\b/i.test(t) ? 'Muster Roll' : '';

  const actMatch = t.match(/The\s+Andhra\s+Pradesh\s+Factories\s+Rules?,?\s*\d{4}/i);
  const actLine = actMatch ? actMatch[0].replace(/\s+/g, ' ').trim() : '';

  const reference = [referenceBracket, actLine].filter(Boolean).join('\n').trim();

  return { title, subtitle, reference };
}

export function normalizeForm25HeaderForDisplay(formHeader) {
  if (!formHeader || typeof formHeader !== 'object') return formHeader;
  const joined = [formHeader.title, formHeader.subtitle, formHeader.reference]
    .filter(Boolean)
    .join('\n');
  const parsed = splitForm25HeaderBlock(joined);
  if (!parsed) return formHeader;
  return {
    ...formHeader,
    title: parsed.title || formHeader.title,
    subtitle: parsed.subtitle || formHeader.subtitle,
    reference: parsed.reference || formHeader.reference
  };
}

export function shouldPlaceForm25ReferenceBeforeSubtitle(formHeader) {
  const title = String(formHeader?.title || '').trim();
  const ref = String(formHeader?.reference || '').trim();
  if (!title || !ref) return false;
  if (!/\bform[\s._-]*25\b/i.test(title)) return false;
  return /\[Prescribed under Rule\s+\d+\]/i.test(ref.split('\n')[0] || ref);
}

export function isForm25APMusterDayHeaderKey(headerKey) {
  const h = String(headerKey || '').trim();
  if (/^for\s+the\s+period\s+ending_\d{1,2}$/i.test(h)) return true;
  if (/^\d{1,2}$/.test(h)) {
    const n = Number(h);
    return n >= 1 && n <= 31;
  }
  const m = h.match(/_(\d{1,2})$/);
  if (!m) return false;
  const n = Number(m[1]);
  return n >= 1 && n <= 31;
}

export function resolveForm25APMusterDayNumberFromHeader(header) {
  const h = String(header || '').trim();
  const m1 = h.match(/^For the period ending_(\d{1,2})$/i);
  if (m1) return Number(m1[1]);
  const m2 = h.match(/_(\d{1,2})$/);
  if (m2) return Number(m2[1]);
  if (/^\d{1,2}$/.test(h)) return Number(h);
  return 0;
}

/** Day-of-month columns for AP Muster (canonical or Excel-parsed). */
export function listForm25APMusterDayHeaders(headers) {
  const out = [];
  (Array.isArray(headers) ? headers : []).forEach((header) => {
    const day = resolveForm25APMusterDayNumberFromHeader(header);
    if (day >= 1 && day <= 31) out.push({ header, day });
  });
  return out;
}

/** Read cell across canonical / Excel day column key variants. */
export function readForm25APMusterCellValue(row, header) {
  if (!row || typeof row !== 'object') return '';
  const direct = row[header];
  if (direct != null && String(direct).trim() !== '') return direct;
  const day = resolveForm25APMusterDayNumberFromHeader(header);
  if (day < 1 || day > 31) return direct == null ? '' : direct;
  const aliases = [
    `For the period ending_${day}`,
    `For the period ending ${day}`,
    String(day),
    `Day ${day}`,
    `Dates_${day}`,
    `DATES_${day}`
  ];
  for (let i = 0; i < aliases.length; i += 1) {
    const val = row[aliases[i]];
    if (val != null && String(val).trim() !== '') return val;
  }
  const bucket = `day${day}`;
  for (const [key, val] of Object.entries(row)) {
    if (String(key).startsWith('__')) continue;
    if (form25APMusterHeaderAliasBucket(key) === bucket && val != null && String(val).trim() !== '') {
      return val;
    }
  }
  return '';
}

/**
 * Fill day columns with P / A / WO.
 * formatCode(attendanceRec, dayDate) → display code; getRecForDay optional per-row lookup.
 */
export function fillForm25APMusterDayAttendanceRows(
  rows,
  headers,
  {
    year,
    monthIndex,
    formatCode,
    getRecForDay,
    overwrite = false,
    skipFutureDaysInCurrentMonth = true,
    now = new Date()
  } = {}
) {
  const dayHeaders = listForm25APMusterDayHeaders(headers);
  if (!Array.isArray(rows) || dayHeaders.length === 0) return 0;
  const y = Number(year);
  const mi = Number(monthIndex);
  if (!Number.isFinite(y) || !Number.isFinite(mi) || mi < 0 || mi > 11) return 0;
  const format =
    typeof formatCode === 'function'
      ? formatCode
      : (_rec, dayDate) => {
          const dow = dayDate.getDay();
          return dow === 0 || dow === 6 ? 'WO' : '';
        };
  const isCurrentMonth = y === now.getFullYear() && mi === now.getMonth();
  let filled = 0;
  rows.forEach((row, rowIndex) => {
    if (!row || typeof row !== 'object') return;
    dayHeaders.forEach(({ header, day }) => {
      const cur = readForm25APMusterCellValue(row, header);
      if (!overwrite && String(cur ?? '').trim()) return;
      if (skipFutureDaysInCurrentMonth && isCurrentMonth && day > now.getDate()) return;
      const dayDate = new Date(y, mi, day);
      if (dayDate.getMonth() !== mi) return;
      const rec =
        typeof getRecForDay === 'function' ? getRecForDay(row, rowIndex, day, dayDate) : null;
      const code = format(rec, dayDate);
      if (code == null || String(code).trim() === '') return;
      row[header] = String(code).trim();
      filled += 1;
    });
  });
  return filled;
}

export function countForm25APMusterDayHeadersFromAny(tableHeaders) {
  const days = listForm25APMusterDayHeaders(tableHeaders).map((d) => d.day);
  return days.length > 0 ? Math.max(...days) : 0;
}

export function countForm25APMusterDayHeaders(tableHeaders) {
  return countForm25APMusterDayHeadersFromAny(tableHeaders);
}

export function looksLikeForm25APMusterPartialHeaders(tableHeaders) {
  const src = Array.isArray(tableHeaders) ? tableHeaders : [];
  if (src.length < FORM_25_AP_MUSTER_PREFIX_COUNT + 5) return false;
  const joined = src.map((h) => form25APMusterHeaderNorm(h)).join('\n');
  if (/relay/.test(joined) && /shift/.test(joined) && /period\s+of\s+work/.test(joined)) return true;
  const dayLike = src.filter((h) => isForm25APMusterDayHeaderKey(h)).length;
  return dayLike >= 5;
}

export function buildForm25APMusterSecondBandLabel(monthName, year) {
  const mon = String(monthName || '').trim();
  const y = Number(year);
  if (!mon) return '';
  if (Number.isFinite(y) && y > 0) return `${mon}-${y}`;
  return mon;
}

export function isForm25APMusterRemarksHeader(h) {
  return /^remarks?$/i.test(form25APMusterHeaderNorm(h));
}

export function buildForm25APMusterCanonicalHeaders(dayCount = 31) {
  const headers = [
    'Serial Number',
    'Name',
    "Father's name",
    'Designation',
    'Group',
    'Relay',
    'Shift number',
    'Period of work'
  ];
  const days = Math.min(Math.max(Number(dayCount) || 31, 1), 31);
  for (let d = 1; d <= days; d += 1) {
    headers.push(`For the period ending_${d}`);
  }
  headers.push(FORM_25_AP_REMARKS_HEADER);
  return headers;
}

export function detectForm25APMusterRemarksColumn(getCell, headerRow, startCol, maxCols) {
  for (let r = headerRow; r <= headerRow + 3; r += 1) {
    for (let c = startCol; c < maxCols; c += 1) {
      const raw = String(getCell(r, c) || '').trim();
      if (isForm25APMusterRemarksHeader(raw)) return c;
    }
  }
  return null;
}

export function detectForm25APMusterDayColumnMap(getCell, headerRow, startCol, maxCols) {
  const dayCols = new Map();
  const scanFrom = Math.max(startCol, 0);
  const scanTo = Math.max(maxCols, scanFrom + 40);
  for (let r = headerRow; r <= headerRow + 3; r += 1) {
    for (let c = scanFrom; c < scanTo; c += 1) {
      const raw = String(getCell(r, c) || '')
        .trim()
        .replace(/[()]/g, '');
      if (!/^\d{1,2}$/.test(raw)) continue;
      const n = Number(raw);
      if (n >= 1 && n <= 31 && !dayCols.has(n)) {
        dayCols.set(n, c);
      }
    }
  }
  return dayCols;
}

export function resolveForm25APMusterDayCount(parsedHeaders, dayColumnMap, monthDayCount) {
  const fromHeaders = countForm25APMusterDayHeadersFromAny(parsedHeaders);
  const fromSheet = dayColumnMap && dayColumnMap.size > 0 ? Math.max(...dayColumnMap.keys()) : 0;
  const fromMonth = Number(monthDayCount);
  if (Number.isFinite(fromMonth) && fromMonth > 0) {
    return Math.min(31, fromMonth);
  }
  if (!fromHeaders && !fromSheet) return 31;
  return Math.min(31, Math.max(fromHeaders, fromSheet, 31));
}

export function headersIndicateForm25APMusterTable(tableHeaders) {
  const joined = (Array.isArray(tableHeaders) ? tableHeaders : [])
    .map((h) => form25APMusterHeaderNorm(h))
    .join('\n');
  if (!joined) return false;
  if (/muster\s+roll/.test(joined) && /relay/.test(joined) && /shift/.test(joined)) return true;
  if (/serial/.test(joined) && /father/.test(joined) && /period\s+of\s+work/.test(joined)) return true;
  if (/for\s+the\s+period\s+ending/.test(joined) && /shift\s+number/.test(joined)) return true;
  const dayEndingCols = countForm25APMusterDayHeadersFromAny(tableHeaders);
  if (dayEndingCols >= 8 && /relay/.test(joined)) return true;
  if (looksLikeForm25APMusterPartialHeaders(tableHeaders)) return true;
  return false;
}

function buildForm25ContextBlob(formHeader, rowItem, fileName, tableHeaders, sheetText) {
  return [
    rowItem?.formName,
    rowItem?.FormName,
    rowItem?.formFileName,
    rowItem?.FormFileName,
    rowItem?.state,
    rowItem?.State,
    fileName,
    formHeader?.title,
    formHeader?.subtitle,
    formHeader?.reference,
    sheetText,
    Array.isArray(tableHeaders) ? tableHeaders.join(' ') : tableHeaders
  ]
    .join(' ')
    .toLowerCase();
}

/** Tamil Nadu Form 25 — festival holidays block + native Excel layout (not AP Muster). */
export function isForm25TamilNaduContext(partsOrBlob) {
  const p = String(partsOrBlob || '').toLowerCase();
  return /tamil[\s._-]*nadu|tamilnadu|form_25[_\s-]*tamil|form[\s._-]*25[_\s-]*tamil/.test(p);
}

export function form25HasFestivalHolidayHeaderBlock(partsOrBlob) {
  const p = String(partsOrBlob || '').toLowerCase();
  return (
    /festival\s+holidays?\s+approval\s+proceedings/.test(p) ||
    /approved\s+festival\s+holidays/.test(p)
  );
}

export function isForm25AndhraPradeshContext(partsOrBlob) {
  const p = String(partsOrBlob || '').toLowerCase();
  return (
    /andhra[\s._-]*pradesh/.test(p) ||
    /form_25[_\s-]*andhra/.test(p) ||
    /andhrapradesh/.test(p) ||
    (/muster\s+roll/.test(p) && /rule\s*103/.test(p))
  );
}

/** Form No. 25 AP Muster Roll [Rule 103] — Andhra Pradesh only. */
export function isForm25APMusterRollContext(
  formHeader,
  rowItem,
  fileName,
  tableHeaders = '',
  sheetText = ''
) {
  const parts = buildForm25ContextBlob(formHeader, rowItem, fileName, tableHeaders, sheetText);
  if (!/\bform[\s._-]*25\b/.test(parts)) return false;

  if (isForm25TamilNaduContext(parts)) return false;
  if (form25HasFestivalHolidayHeaderBlock(parts)) return false;

  if (/register\s+of\s+adult|worker\s+identity|scheme\s+of\s+shifts/.test(parts) && !/muster\s+roll/.test(parts)) {
    return false;
  }
  if (/contractor/.test(parts) && /principal\s+employer/.test(parts) && !/muster\s+roll/.test(parts)) {
    return false;
  }

  if (!isForm25AndhraPradeshContext(parts)) return false;

  if (/muster\s+roll/.test(parts)) return true;
  if (/rule\s*103/.test(parts)) return true;
  if (/form_25[_\s-]*andhra/i.test(parts)) return true;
  return headersIndicateForm25APMusterTable(tableHeaders);
}

export function normalizeForm25APMusterTableHeaders(parsedHeaders, dayCountHint) {
  const src = Array.isArray(parsedHeaders) ? parsedHeaders : [];
  const dayCount = resolveForm25APMusterDayCount(src, null, dayCountHint);
  const canonical = buildForm25APMusterCanonicalHeaders(dayCount);
  const looksLikeMuster =
    headersIndicateForm25APMusterTable(src) || looksLikeForm25APMusterPartialHeaders(src);
  if (!looksLikeMuster) {
    return canonical;
  }
  const out = [...canonical];
  src.forEach((h, i) => {
    if (i >= FORM_25_AP_MUSTER_PREFIX_COUNT) return;
    if (i < out.length && String(h || '').trim() && !isForm25APMusterDayHeaderKey(h)) {
      out[i] = String(h).trim();
    }
  });
  return out;
}

export function buildForm25APMusterSubColumnsFromHeaders(headers) {
  const subs = {};
  (headers || []).forEach((h) => {
    const m = String(h || '').match(/^For the period ending_(\d{1,2})$/i);
    if (!m) return;
    const parent = 'For the period ending';
    if (!subs[parent]) subs[parent] = [];
    if (!subs[parent].includes(m[1])) subs[parent].push(m[1]);
  });
  return subs;
}

export function form25APMusterHeaderAliasBucket(h) {
  const s = form25APMusterHeaderNorm(h);
  if (/serial|\(1\)/.test(s) && (/serial|number/.test(s) || s === '(1)')) return 'serial';
  if (s === 'name' || (/\bname\b/.test(s) && !/father/.test(s) && !/factory/.test(s))) return 'name';
  if (/father/.test(s)) return 'father';
  if (/designat/.test(s)) return 'designation';
  if (/^group$|\(5\)/.test(s)) return 'group';
  if (/relay/.test(s)) return 'relay';
  if (/shift/.test(s)) return 'shift';
  if (/period\s+of\s+work/.test(s)) return 'period';
  if (/^remarks?$/.test(s)) return 'remarks';
  const daySuffix = String(h || '').match(/_(\d{1,2})$/);
  if (daySuffix) {
    const n = Number(daySuffix[1]);
    if (n >= 1 && n <= 31) return `day${n}`;
  }
  if (/^\d{1,2}$/.test(s)) {
    const n = Number(s);
    if (n >= 1 && n <= 31) return `day${n}`;
  }
  return s;
}

export function remapForm25APMusterRowsToHeaders(rows, sourceHeaders, targetHeaders, dayCountHint) {
  const src = Array.isArray(sourceHeaders) ? sourceHeaders : [];
  const tgt = normalizeForm25APMusterTableHeaders(targetHeaders, dayCountHint);
  if (!Array.isArray(rows)) return [];
  const srcBuckets = src.map(form25APMusterHeaderAliasBucket);
  const tgtBuckets = tgt.map(form25APMusterHeaderAliasBucket);
  return rows.map((row) => {
    if (!row || typeof row !== 'object') return tgt.reduce((o, h) => ({ ...o, [h]: '' }), {});
    const out = {};
    tgt.forEach((th, ti) => {
      const tb = tgtBuckets[ti];
      let val = '';
      if (Object.prototype.hasOwnProperty.call(row, th)) {
        val = row[th];
      } else {
        const si = srcBuckets.indexOf(tb);
        if (si >= 0 && src[si]) {
          const sk = src[si];
          if (Object.prototype.hasOwnProperty.call(row, sk)) val = row[sk];
        }
      }
      out[th] = val == null ? '' : val;
    });
    return out;
  });
}

export function isForm25APPeriodOfWorkHeader(h) {
  return /^period\s+of\s+work$/i.test(form25APMusterHeaderNorm(h));
}

/** Shift name or "start - end" for Form 25 AP Muster "Period of work" column. */
export function formatForm25APPeriodOfWorkValue({ shiftName = '', shiftStart = '', shiftEnd = '' } = {}) {
  const name = String(shiftName || '').trim();
  if (name && name !== '-') return name;
  const start = String(shiftStart || '').trim();
  const end = String(shiftEnd || '').trim();
  if (start && end && start !== '-' && end !== '-') {
    return `${start} - ${end}`;
  }
  if (start && start !== '-') return start;
  if (end && end !== '-') return end;
  return '';
}

export function isForm25APMusterSkipAutofillHeader(h, normalizeLooseHeaderText) {
  const s = normalizeLooseHeaderText(h);
  if (!s) return false;
  if (isForm25APMusterDayHeaderKey(h)) return false;
  if (/^group$/.test(s.replace(/\s+/g, ' ').trim())) return true;
  if (/relay/.test(s)) return true;
  if (/shift/.test(s) && /number/.test(s)) return true;
  if (/period\s+of\s+work/.test(s)) return true;
  if (/^remarks?$/.test(s.replace(/\s+/g, ' ').trim())) return true;
  return false;
}

export function enrichForm25APFactoryHeadersFromSheet(
  headerRowIndex,
  effectiveSheetCols,
  getMergedAwareCellText,
  scanRowValueAfterLabel
) {
  const results = [];
  for (const spec of FORM_25_AP_FACTORY_HEADER_SPECS) {
    outer: for (let r = 0; r < headerRowIndex; r++) {
      for (let c = 0; c < Math.max(20, effectiveSheetCols); c++) {
        const raw = String(getMergedAwareCellText(r, c) || '').trim();
        if (!raw || !spec.test(raw)) continue;
        const inline = raw.match(/:\s*(.+)$/);
        let val = inline ? String(inline[1] || '').trim() : '';
        if (!val) val = scanRowValueAfterLabel(r, c, effectiveSheetCols, getMergedAwareCellText);
        results.push({ label: spec.label, value: val, key: spec.key });
        break outer;
      }
    }
  }
  return results;
}

export function rebuildForm25APMusterTableFromSheet({
  jsonData,
  effectiveSheetCols,
  getMergedAwareCellText,
  headerRowIndex = -1,
  tableStartCol = 0,
  startIndex = -1,
  headersToUse = [],
  tableData = [],
  monthDayCount,
  looksLikeSerialHeader
}) {
  let headerRow = headerRowIndex;
  let startCol = tableStartCol >= 0 ? tableStartCol : 0;
  const maxCols = Math.max(effectiveSheetCols, 100);
  if (headerRow < 0) {
    for (let r = 0; r < Math.min(15, jsonData.length); r += 1) {
      let rowProbe = '';
      for (let c = 0; c < maxCols; c += 1) {
        rowProbe += ` ${form25APMusterHeaderNorm(getMergedAwareCellText(r, c))}`;
      }
      if (!/serial/.test(rowProbe) || !/name/.test(rowProbe) || !/relay/.test(rowProbe)) continue;
      if (!/father|designat|period\s+of\s+work|shift/.test(rowProbe)) continue;
      headerRow = r;
      for (let c = 0; c < 5; c += 1) {
        const t = form25APMusterHeaderNorm(getMergedAwareCellText(r, c));
        const raw = getMergedAwareCellText(r, c);
        if (/serial/.test(t) || (typeof looksLikeSerialHeader === 'function' && looksLikeSerialHeader(raw))) {
          startCol = c;
          break;
        }
      }
      break;
    }
  }
  if (headerRow < 0) return null;

  let dataStart = startIndex >= 0 ? startIndex : headerRow + 1;
  for (let offset = 1; offset <= 4; offset += 1) {
    let markers = 0;
    for (let c = startCol; c < maxCols; c += 1) {
      const raw = String(getMergedAwareCellText(headerRow + offset, c) || '').trim();
      const v = raw.replace(/[()]/g, '').trim();
      if (/^\d{1,2}$/.test(v)) markers += 1;
    }
    if (markers >= 8) {
      dataStart = Math.max(dataStart, headerRow + offset + 1);
      break;
    }
  }

  const dayColumnMap = detectForm25APMusterDayColumnMap(
    (r, c) => getMergedAwareCellText(r, c),
    headerRow,
    startCol,
    maxCols
  );
  const dayCount = resolveForm25APMusterDayCount(headersToUse, dayColumnMap, monthDayCount);
  const canonical = buildForm25APMusterCanonicalHeaders(dayCount);

  const colByBucket = {
    serial: startCol,
    name: startCol + 1,
    father: startCol + 2,
    designation: startCol + 3,
    group: startCol + 4,
    relay: startCol + 5,
    shift: startCol + 6,
    period: startCol + 7
  };
  for (let d = 1; d <= dayCount; d += 1) {
    colByBucket[`day${d}`] = dayColumnMap.has(d) ? dayColumnMap.get(d) : startCol + 8 + (d - 1);
  }
  const remarksCol = detectForm25APMusterRemarksColumn(
    (r, c) => getMergedAwareCellText(r, c),
    headerRow,
    startCol,
    maxCols
  );
  colByBucket.remarks =
    remarksCol != null ? remarksCol : startCol + 8 + dayCount;

  const readCell = (rowIdx, bucket) => {
    const c = colByBucket[bucket];
    if (c == null) return '';
    return String(getMergedAwareCellText(rowIdx, c) || '').trim();
  };

  const rebuiltRows = [];
  for (let r = dataStart; r < jsonData.length; r += 1) {
    const rowObj = {};
    let hasData = false;
    canonical.forEach((h) => {
      const bucket = form25APMusterHeaderAliasBucket(h);
      const val = readCell(r, bucket);
      rowObj[h] = val;
      if (val) hasData = true;
    });
    if (hasData) rebuiltRows.push(rowObj);
  }

  const priorRows = Array.isArray(tableData) ? tableData.map((row) => ({ ...row })) : [];
  const mergedRows =
    priorRows.length > 0 && rebuiltRows.length === 0
      ? remapForm25APMusterRowsToHeaders(priorRows, headersToUse, canonical, dayCount)
      : rebuiltRows.length > 0
        ? rebuiltRows
        : remapForm25APMusterRowsToHeaders(priorRows, headersToUse, canonical, dayCount);

  return {
    headers: canonical,
    expandedHeaders: canonical,
    subColumnsData: buildForm25APMusterSubColumnsFromHeaders(canonical),
    headerRowIndex: headerRow,
    startIndex: dataStart,
    tableStartCol: startCol,
    tableData: mergedRows,
    dayColumnMap
  };
}

export function buildForm25APMusterTemplateColMap(startCol, dayCount, dayColumnMap) {
  const cols = [
    { col: startCol, bucket: 'serial' },
    { col: startCol + 1, bucket: 'name' },
    { col: startCol + 2, bucket: 'father' },
    { col: startCol + 3, bucket: 'designation' },
    { col: startCol + 4, bucket: 'group' },
    { col: startCol + 5, bucket: 'relay' },
    { col: startCol + 6, bucket: 'shift' },
    { col: startCol + 7, bucket: 'period' }
  ];
  for (let d = 1; d <= dayCount; d += 1) {
    cols.push({
      col: dayColumnMap && dayColumnMap.has(d) ? dayColumnMap.get(d) : startCol + 8 + (d - 1),
      bucket: `day${d}`
    });
  }
  cols.push({
    col: startCol + 8 + dayCount,
    bucket: 'remarks'
  });
  return cols;
}

/** Layout for two-row thead: prefix cols + day bands. */
export function getForm25APMusterTheadLayout(headers, secondBandLabel = '') {
  const hdrs = Array.isArray(headers) ? headers : [];
  if (hdrs.length < FORM_25_AP_MUSTER_PREFIX_COUNT + 5) return null;
  const dayCount = countForm25APMusterDayHeadersFromAny(hdrs);
  if (dayCount < 1) return null;
  const prefix = hdrs.slice(0, FORM_25_AP_MUSTER_PREFIX_COUNT);
  const firstBandSize = Math.min(FORM_25_AP_FIRST_DAY_BAND_SIZE, dayCount);
  const firstBandDays = Array.from({ length: firstBandSize }, (_, i) => i + 1);
  const secondBandDays =
    dayCount > firstBandSize
      ? Array.from({ length: dayCount - firstBandSize }, (_, i) => firstBandSize + i + 1)
      : [];
  const hasRemarks = hdrs.some((h) => isForm25APMusterRemarksHeader(h));
  return {
    prefix,
    firstBand: { label: 'For the period ending', days: firstBandDays },
    secondBand: {
      label: secondBandLabel || '',
      days: secondBandDays
    },
    hasRemarks,
    remarksLabel: FORM_25_AP_REMARKS_HEADER
  };
}
