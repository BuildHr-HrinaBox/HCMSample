/** Form No. 6 (AP) — Humidity register [Rule 22]. */

export function form6APHeaderNorm(txt) {
  return String(txt || '')
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

export function isForm6APHumidityRegisterContext(formHeader, rowItem, fileName, sheetText = '') {
  const parts = [
    rowItem?.formName,
    rowItem?.FormName,
    rowItem?.formFileName,
    rowItem?.FormFileName,
    fileName,
    formHeader?.title,
    formHeader?.subtitle,
    formHeader?.reference,
    sheetText
  ]
    .join(' ')
    .toLowerCase();
  if (/form[\s._-]*6[\s._-]*[-_.]?\s*ap|form_6[\s._-]*andhra|form6[\s._-]*andhra/i.test(parts)) {
    return true;
  }
  if (/\bform\s+no\.?\s*6\b|\bform\s+6\b/.test(parts) && /humidity|hygrometer|rule\s+22/.test(parts)) {
    return true;
  }
  if (/\bform\s+no\.?\s*6\b/.test(parts) && /andhra[\s._-]*pradesh|\bap\b/.test(parts)) {
    if (/humidity|hygrometer/.test(parts)) return true;
  }
  return false;
}

export function headersIndicateForm6APHumidityTable(tableHeaders) {
  if (!Array.isArray(tableHeaders) || tableHeaders.length < 5) return false;
  const joined = tableHeaders.map((h) => form6APHeaderNorm(h)).join(' ');
  return (
    joined.includes('dry bulb') &&
    joined.includes('wet bulb') &&
    (joined.includes('date') || joined.includes('month') || joined.includes('day'))
  );
}

export function isForm6APHumidityGroupedHeader(headerKey) {
  const m = String(headerKey || '').match(/^(.+)_([\s\S]+)$/);
  if (!m) return false;
  const sub = form6APHeaderNorm(m[2]);
  return sub === 'dry bulb' || sub === 'wet bulb';
}

export function formatForm6HumidityHeaderLabel(headerKey) {
  const h = String(headerKey || '');
  const m = h.match(/^(.+)_([\s\S]+)$/);
  if (!m) return h;
  const sub = String(m[2] || '').trim();
  if (/^dry\s+bulb$/i.test(sub) || /^wet\s+bulb$/i.test(sub)) return sub;
  return h;
}

/** Humidity table columns are manual entry — never Zoho People autofill. */
export function isForm6APHumiditySkipAutofillHeader(headerKey) {
  const t = form6APHeaderNorm(headerKey);
  if (!t) return false;
  if (t.includes('date') && t.includes('year') && t.includes('month')) return true;
  if (t.includes('dry bulb') || t.includes('wet bulb')) return true;
  if (t.includes('humidity') && t.includes('insert')) return true;
  if (t === 'remarks') return true;
  return isForm6APHumidityGroupedHeader(headerKey);
}

function form6APOrdinalDay(n) {
  const j = n % 10;
  const k = n % 100;
  if (k >= 11 && k <= 13) return `${n}th`;
  if (j === 1) return `${n}st`;
  if (j === 2) return `${n}nd`;
  if (j === 3) return `${n}rd`;
  return `${n}th`;
}

export function isForm6APDateOrdinalValue(value) {
  const m = String(value ?? '')
    .trim()
    .match(/^(\d{1,2})(st|nd|rd|th)$/i);
  if (!m) return false;
  const n = parseInt(m[1], 10);
  return n >= 1 && n <= 31;
}

/** Certification / signatory rows below the day grid must not appear as editable rows. */
export function isForm6APHumidityFooterRow(row, headers) {
  if (!row || typeof row !== 'object') return false;
  const vals = Object.values(row)
    .map((v) => String(v ?? '').trim())
    .filter(Boolean);
  if (vals.length === 0) return false;
  const joined = vals.join(' ').toLowerCase();
  if (/certified\s+that/.test(joined)) return true;
  if (/for\s*\(\s*company/.test(joined)) return true;
  if (/authorised\s+signatory|authorized\s+signatory/.test(joined)) return true;
  if (/signature\s+of\s+(the\s+)?(employer|manager|authorised|authorized)/.test(joined)) return true;
  const dateHeader = findForm6APDateHeader(headers);
  const dateVal = String(row[dateHeader] ?? '').trim();
  if (!isForm6APDateOrdinalValue(dateVal) && vals.some((v) => /certified\s+that/i.test(v))) return true;
  return false;
}

export function filterForm6APHumidityTableRows(rows, headers) {
  if (!Array.isArray(rows)) return [];
  return rows.filter((row) => !isForm6APHumidityFooterRow(row, headers));
}

export function getForm6APHumidityCellPlaceholder(headerKey) {
  const t = form6APHeaderNorm(headerKey);
  if (t.includes('date') && t.includes('year')) return 'Enter value';
  if (t.includes('dry bulb') || t.includes('wet bulb')) return 'Enter value';
  if (t === 'remarks') return 'Enter value';
  if (t.includes('humidity') && t.includes('insert')) return 'Enter value';
  return 'Enter value';
}

export function getForm6APHeaderFieldPlaceholder() {
  return 'Enter value';
}

export const FORM_6_AP_CERTIFICATION_LINE = 'Certified that the above entries are correct.';
export const FORM_6_AP_SIGNED_LABEL = 'Signed';

export function sanitizeForm6APHumidityCellValue(value) {
  const s = String(value ?? '').trim();
  if (!s) return '';
  if (/<\s*div[\s>]/i.test(s) || /please follow the steps/i.test(s)) return '';
  if (/certified\s+that/i.test(s) || /^signed$/i.test(s)) return '';
  return value;
}

/** Template day rows with manual humidity columns cleared — used when opening / autofill. */
export function applyForm6APHumidityTableRowsForAutofill(rows, headers) {
  return clearForm6APHumidityAutofillColumns(ensureForm6APTemplateRows(rows, headers), headers);
}

/** Keep saved humidity values; only normalize day rows and drop footer bleed (save / download). */
export function prepareForm6APHumidityRowsForExport(rows, headers) {
  return ensureForm6APTemplateRows(filterForm6APHumidityTableRows(rows, headers), headers);
}

/** True when any humidity / bulb / remarks cell has user-entered data (not day ordinals). */
export function form6APHumidityRowsHaveManualEntryData(rows, headers) {
  if (!Array.isArray(rows) || rows.length === 0) return false;
  const hdrs = Array.isArray(headers) ? headers : [];
  const dateHeader = findForm6APDateHeader(hdrs);
  return rows.some((row) =>
    hdrs.some((h) => {
      if (h === dateHeader) return false;
      if (!isForm6APHumiditySkipAutofillHeader(h)) return false;
      const v = String(row?.[h] ?? '').trim();
      if (!v) return false;
      if (/certified\s+that/i.test(v) || /^signed$/i.test(v)) return false;
      return true;
    })
  );
}

/** Trigger browser download of an Excel blob. */
export function triggerExcelBlobDownload(blob, fileName) {
  if (!blob) return;
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = fileName || 'form.xlsx';
  link.style.display = 'none';
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

/** Write certification + Signed labels on rows immediately below the 31 day grid (Excel model). */
export function writeForm6APHumidityCertificationFooter({
  dataStartIndex,
  tableStartCol = 0,
  maxCols = 9,
  getMergedAwareCellText,
  writeCell
}) {
  if (typeof writeCell !== 'function' || typeof getMergedAwareCellText !== 'function') return;
  const certRow = dataStartIndex + FORM_6_AP_TEMPLATE_DAY_COUNT;
  const signedRow = certRow + 1;
  const startCol = Math.max(0, tableStartCol);
  const colLimit = Math.max(startCol + 1, maxCols);
  let signedCol = colLimit - 1;

  for (let r = certRow; r <= signedRow + 1; r += 1) {
    for (let c = startCol; c < colLimit; c += 1) {
      const t = String(getMergedAwareCellText(r, c) || '').trim().toLowerCase();
      if (t === 'signed') {
        signedCol = c;
        break;
      }
    }
  }

  writeCell(certRow, startCol, FORM_6_AP_CERTIFICATION_LINE);
  writeCell(certRow, signedCol, FORM_6_AP_SIGNED_LABEL);
  writeCell(signedRow, signedCol, FORM_6_AP_SIGNED_LABEL);
}

/** ExcelJS variant — preserves template styles (SheetJS write corrupts Form 6 AP). */
export function writeForm6APHumidityCertificationFooterExcelJs(worksheet, {
  dataStartRow1Based,
  tableStartCol1Based = 1,
  maxCols = 9
}) {
  if (!worksheet) return;
  const certRow = dataStartRow1Based + FORM_6_AP_TEMPLATE_DAY_COUNT;
  const signedRow = certRow + 1;
  const startCol = Math.max(1, tableStartCol1Based);
  const colLimit = Math.max(startCol, maxCols);
  let signedCol = colLimit;

  const cellText = (r, c) => String(worksheet.getCell(r, c)?.value ?? '').trim().toLowerCase();
  for (let r = certRow; r <= signedRow + 1; r += 1) {
    for (let c = startCol; c <= colLimit; c += 1) {
      if (cellText(r, c) === 'signed') {
        signedCol = c;
        break;
      }
    }
  }

  worksheet.getCell(certRow, startCol).value = FORM_6_AP_CERTIFICATION_LINE;
  worksheet.getCell(certRow, signedCol).value = FORM_6_AP_SIGNED_LABEL;
  worksheet.getCell(signedRow, signedCol).value = FORM_6_AP_SIGNED_LABEL;
}

export function clearForm6APHumidityAutofillColumns(rows, headers) {
  if (!Array.isArray(rows) || !Array.isArray(headers)) return rows || [];
  const dateHeader = findForm6APDateHeader(headers);
  const skipHeaders = headers.filter((h) => isForm6APHumiditySkipAutofillHeader(h));
  return rows.map((row, rowIdx) => {
    if (!row || typeof row !== 'object') return row;
    const next = { ...row };
    skipHeaders.forEach((header) => {
      if (header === dateHeader) {
        next[header] =
          rowIdx < FORM_6_AP_TEMPLATE_ORDINALS.length ? FORM_6_AP_TEMPLATE_ORDINALS[rowIdx] : '';
        return;
      }
      next[header] = '';
    });
    return next;
  });
}

const FORM_6_AP_DATE_HEADER = 'Date, year, month, day';
export const FORM_6_AP_TEMPLATE_DAY_COUNT = 31;
const FORM_6_AP_TEMPLATE_ORDINALS = Array.from({ length: FORM_6_AP_TEMPLATE_DAY_COUNT }, (_, i) =>
  form6APOrdinalDay(i + 1)
);

const FORM_6_AP_HEADER_FIELD_SPECS = [
  {
    key: 'form6_ap_header_department',
    label: 'Department:',
    test: (raw) => /^department\s*:?\s*$/i.test(raw)
  },
  {
    key: 'form6_ap_header_hygrometer_1',
    label: 'Hygrometer:',
    test: (raw) => /^hygrometer\s*:?\s*$/i.test(raw),
    occurrence: 0
  },
  {
    key: 'form6_ap_header_hygrometer_2',
    label: 'Hygrometer:',
    test: (raw) => /^hygrometer\s*:?\s*$/i.test(raw),
    occurrence: 1
  },
  {
    key: 'form6_ap_header_distinctive_mark',
    label: 'Distinctive mark or No.:',
    test: (raw) => /^distinctive\s+mark\s+or\s+no\.?\s*:?\s*$/i.test(raw)
  },
  {
    key: 'form6_ap_header_position',
    label: 'Position in department:',
    test: (raw) => /^position\s+in\s+department\s*:?\s*$/i.test(raw)
  }
];

function scanRowValueAfterLabel(getMergedAwareCellText, rowIndex, labelCol, maxCols) {
  for (let nc = labelCol + 1; nc < Math.min(labelCol + 6, maxCols); nc += 1) {
    const v = String(getMergedAwareCellText(rowIndex, nc) || '').trim();
    if (!v || v.includes(':')) continue;
    return v;
  }
  return '';
}

export function enrichForm6APHeaderFieldsFromSheet(headerRowIndex, effectiveSheetCols, getMergedAwareCellText) {
  if (headerRowIndex <= 0) return [];
  const maxCols = Math.max(effectiveSheetCols, 9);
  const norm = (txt) =>
    String(txt || '')
      .replace(/\r?\n/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  const found = [];
  const hygrometerHits = { count: 0 };

  for (let r = 0; r < headerRowIndex; r += 1) {
    for (let c = 0; c < maxCols; c += 1) {
      const raw = norm(getMergedAwareCellText(r, c));
      if (!raw) continue;
      FORM_6_AP_HEADER_FIELD_SPECS.forEach((spec) => {
        if (!spec.test(raw)) return;
        if (spec.occurrence != null) {
          if (spec.occurrence !== hygrometerHits.count) return;
          hygrometerHits.count += 1;
        }
        if (found.some((f) => f.key === spec.key)) return;
        found.push({
          label: spec.label,
          value: scanRowValueAfterLabel(getMergedAwareCellText, r, c, maxCols),
          key: spec.key
        });
      });
    }
  }

  return found;
}

/** Always expose the five AP template header fields even when the sheet scan misses merged labels. */
export function ensureForm6APHeaderFields(fields) {
  const list = Array.isArray(fields) ? [...fields] : [];
  const hasKey = (key) => list.some((f) => f?.key === key);
  FORM_6_AP_HEADER_FIELD_SPECS.forEach((spec) => {
    if (hasKey(spec.key)) return;
    list.push({ label: spec.label, value: '', key: spec.key });
  });
  return list;
}

/**
 * Rebuild table headers from the Excel two-row band (time periods + Dry/Wet bulb).
 * Avoids merged-cell lookback pulling unrelated text into column keys.
 */
export function rebuildForm6APHumidityHeadersFromSheet({
  jsonData,
  effectiveSheetCols,
  getMergedAwareCellText,
  merges
}) {
  const norm = (txt) =>
    String(txt || '')
      .replace(/\r?\n/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  const normLower = (txt) => form6APHeaderNorm(txt);

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
  for (let r = 0; r < Math.min(30, jsonData.length); r += 1) {
    for (let c = 0; c < Math.max(effectiveSheetCols, 9); c += 1) {
      const t = normLower(getMergedAwareCellText(r, c));
      if (t.includes('date') && t.includes('year') && t.includes('month') && t.includes('day')) {
        headerRowIndex = r;
        break;
      }
    }
    if (headerRowIndex >= 0) break;
  }
  if (headerRowIndex < 0) return null;

  const subRowIndex = headerRowIndex + 1;
  const subRowHasBulbs = ['dry bulb', 'wet bulb'].some((needle) => {
    for (let c = 0; c < Math.max(effectiveSheetCols, 9); c += 1) {
      if (normLower(getMergedAwareCellText(subRowIndex, c)).includes(needle)) return true;
    }
    return false;
  });
  if (!subRowHasBulbs) return null;

  const flatHeaders = [];
  const subColumnsData = {};
  const mainHeaders = [];
  let c = 0;
  const maxCols = Math.max(effectiveSheetCols, 9);

  while (c < maxCols) {
    const mergeEnd = getMergeSpanEndCol(headerRowIndex, c);
    let endC = mergeEnd;
    const main = norm(getMergedAwareCellText(headerRowIndex, c));
    if (!main) {
      c += 1;
      continue;
    }
    while (endC < maxCols) {
      const nextMain = norm(getMergedAwareCellText(headerRowIndex, endC));
      if (!nextMain) break;
      if (normLower(nextMain) !== normLower(main)) break;
      endC = Math.max(endC, getMergeSpanEndCol(headerRowIndex, endC));
    }

    const subs = [];
    for (let sc = c; sc < endC; sc += 1) {
      const sub = norm(getMergedAwareCellText(subRowIndex, sc));
      if (!sub || normLower(sub) === normLower(main)) continue;
      if (/^dry\s+bulb$/i.test(sub) || /^wet\s+bulb$/i.test(sub)) subs.push(sub);
    }

    mainHeaders.push(main);
    if (!subColumnsData[main]) subColumnsData[main] = [];

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

  while (flatHeaders.length > 0 && !flatHeaders[flatHeaders.length - 1]) flatHeaders.pop();
  if (!headersIndicateForm6APHumidityTable(flatHeaders)) return null;

  const headerFields = enrichForm6APHeaderFieldsFromSheet(
    headerRowIndex,
    effectiveSheetCols,
    getMergedAwareCellText
  );

  return {
    headers: mainHeaders,
    expandedHeaders: flatHeaders,
    subColumnsData,
    headerRowIndex,
    startIndex: subRowIndex + 1,
    startCol: 0,
    headerFields
  };
}

export function findForm6APDateHeader(headers) {
  const list = Array.isArray(headers) ? headers : [];
  return (
    list.find((h) => {
      const t = form6APHeaderNorm(h);
      return t.includes('date') && t.includes('year') && t.includes('month');
    }) || FORM_6_AP_DATE_HEADER
  );
}

/** Ensure one row per calendar day (1st–31st) under the date column; drop certification footer rows. */
export function ensureForm6APTemplateRows(rows, headers) {
  const hdrs = Array.isArray(headers) ? headers : [];
  const dateHeader = findForm6APDateHeader(hdrs);
  const cleaned = filterForm6APHumidityTableRows(rows, hdrs);
  const byOrdinal = new Map();
  const unmatched = [];

  cleaned.forEach((row) => {
    const copy = { ...(row || {}) };
    const d = String(copy[dateHeader] ?? '').trim();
    if (isForm6APDateOrdinalValue(d)) {
      const key = d.toLowerCase();
      if (!byOrdinal.has(key)) byOrdinal.set(key, copy);
    } else {
      unmatched.push(copy);
    }
  });

  const missingOrdinals = FORM_6_AP_TEMPLATE_ORDINALS.filter((o) => !byOrdinal.has(o.toLowerCase()));
  unmatched.forEach((row, i) => {
    const ord = missingOrdinals[i];
    if (ord) byOrdinal.set(ord.toLowerCase(), row);
  });

  const out = [];

  FORM_6_AP_TEMPLATE_ORDINALS.forEach((ordinal) => {
    const key = ordinal.toLowerCase();
    const src = byOrdinal.get(key);
    const row = { ...(src || {}) };
    row[dateHeader] = ordinal;
    hdrs.forEach((h) => {
      if (h === dateHeader) return;
      if (row[h] == null) row[h] = '';
    });
    out.push(row);
  });

  return out;
}

export function remapForm6APRowsToHeaders(rows, oldHeaders, newHeaders) {
  if (!Array.isArray(rows) || !Array.isArray(newHeaders) || newHeaders.length === 0) return rows || [];
  const oldList = Array.isArray(oldHeaders) ? oldHeaders : [];
  const normKey = (k) => form6APHeaderNorm(k);

  return rows.map((row) => {
    if (!row || typeof row !== 'object') return row;
    const next = {};
    newHeaders.forEach((newH, i) => {
      if (row[newH] != null && String(row[newH]).trim() !== '') {
        next[newH] = row[newH];
        return;
      }
      const oldH = oldList[i];
      if (oldH != null && row[oldH] != null && String(row[oldH]).trim() !== '') {
        next[newH] = row[oldH];
        return;
      }
      const wanted = normKey(newH);
      const matchKey = Object.keys(row).find((k) => normKey(k) === wanted);
      next[newH] = matchKey != null ? row[matchKey] : '';
    });
    return next;
  });
}
