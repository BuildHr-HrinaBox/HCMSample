/** Karnataka Form F — Register of Leave with Wages (Rule 8): header fields + PART I earned leave table. */

export function formFKarnatakaHeaderNorm(txt) {
  return String(txt || '')
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

export function isFormFLeaveWithWagesContext(formHeader, rowItem, fileName, sheetText = '') {
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
  ]
    .filter((x) => x != null && String(x).trim() !== '')
    .join(' ')
    .toLowerCase();

  const hasFormF = /\bform[\s._-]*f\b/i.test(parts);
  const hasLeaveRegister =
    /leave\s+with\s+wages/i.test(parts) ||
    /register\s+of\s+leave/i.test(parts);
  if (hasFormF && hasLeaveRegister) return true;
  if (hasFormF && /karnataka/.test(parts)) return true;
  return false;
}

export function isFormFKarnatakaTableLayoutFormHeader(formHeader) {
  return !!formHeader?.formFKarnatakaTableLayout;
}

/** PART I — Earned Leave columns (Karnataka Form F template). */
export const FORM_F_KARNATAKA_PART_I_HEADERS = [
  'Days worked From',
  'Days worked To',
  'Total days worked',
  'Leave earned',
  'Leave at credit (incl balance if any)',
  'Leave availed From',
  'Leave availed To',
  'Leave availed No. of days',
  'Balance on return from leave',
  'Date on which wages for leave paid',
  'Remarks',
];

const PART_I_SUBHEADER_FRAGMENTS = [
  /^from$/i,
  /^to$/i,
  /total\s+days\s+worked/i,
  /leave\s+earned/i,
  /leave\s+at\s+credit/i,
  /no\.?\s*of\s+days/i,
  /balance\s+on\s+return\s+from\s+leave/i,
  /date\s+on\s+which\s+wages\s+for\s+leave\s+paid/i,
  /^remarks?$/i,
];

export function headersIndicateFormFKarnatakaPartITable(tableHeaders) {
  const joined = (Array.isArray(tableHeaders) ? tableHeaders : [])
    .map((h) => formFKarnatakaHeaderNorm(h))
    .join('\n');
  if (!joined) return false;
  if (/sl\s*no\s+in\s+the\s+register\s+of\s+adult/.test(joined) && !/leave\s+earned/.test(joined)) {
    return false;
  }
  return (
    /leave\s+earned/.test(joined) &&
    /balance\s+on\s+return\s+from\s+leave/.test(joined) &&
    (/total\s+days\s+worked/.test(joined) || /days\s+worked/.test(joined))
  );
}

export function resolveFormFKarnatakaTableHeaders(tableHeaders) {
  if (headersIndicateFormFKarnatakaPartITable(tableHeaders)) {
    return [...tableHeaders];
  }
  return [...FORM_F_KARNATAKA_PART_I_HEADERS];
}

const normalizeCell = (txt) =>
  String(txt || '')
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

const isPartIEarnedLeaveBanner = (text) => {
  const t = formFKarnatakaHeaderNorm(text);
  return /part\s*[-–]?\s*i/.test(t) && /earned\s+leave/.test(t);
};

const isEmployeeIdentityBannerRow = (rowCells) => {
  const joined = rowCells.map((c) => formFKarnatakaHeaderNorm(c)).join(' ');
  return (
    /sl\s*no\s+in\s+the\s+register\s+of\s+adult/.test(joined) &&
    (/name\s+of\s+the\s+person/.test(joined) || /date\s+of\s+entry\s+into\s+service/.test(joined))
  );
};

const rowLooksLikePartISubHeaderRow = (rowCells) => {
  const nonEmpty = rowCells.map((c) => normalizeCell(c)).filter(Boolean);
  if (nonEmpty.length < 4) return false;
  const joined = nonEmpty.join(' ').toLowerCase();
  if (isEmployeeIdentityBannerRow(nonEmpty)) return false;
  let hits = 0;
  nonEmpty.forEach((cell) => {
    if (PART_I_SUBHEADER_FRAGMENTS.some((re) => re.test(cell))) hits += 1;
  });
  return hits >= 3 && /leave\s+earned/.test(joined);
};

const rowLooksLikePartINumberedIndexRow = (rowCells) => {
  let nums = 0;
  let total = 0;
  rowCells.forEach((cell) => {
    const t = normalizeCell(cell);
    if (!t) return;
    total += 1;
    if (/^\d+$/.test(t)) nums += 1;
  });
  return total >= 6 && nums >= 6 && nums / total >= 0.55;
};

const combinePartIHeaderBand = (upperCells, lowerCells) => {
  const maxLen = Math.max(upperCells.length, lowerCells.length, 11);
  const headers = [];
  let currentGroup = '';
  for (let c = 0; c < maxLen; c += 1) {
    const upper = normalizeCell(upperCells[c]);
    const lower = normalizeCell(lowerCells[c]);
    if (upper && !/^\d+$/.test(upper)) {
      if (/no\s+of\s+days\s+worked/i.test(upper)) currentGroup = 'Days worked';
      else if (/leave\s+availed/i.test(upper)) currentGroup = 'Leave availed';
      else if (/leave\s+earned/i.test(upper)) currentGroup = 'Leave earned';
      else if (/leave\s+at\s+credit/i.test(upper)) currentGroup = 'Leave at credit';
      else if (/balance\s+on\s+return/i.test(upper)) currentGroup = 'Balance on return from leave';
      else if (/date\s+on\s+which\s+wages/i.test(upper)) currentGroup = 'Date on which wages for leave paid';
      else if (/^remarks?$/i.test(upper)) currentGroup = 'Remarks';
    }
    let label = lower || upper;
    if (/^\d+$/.test(label)) {
      const canonical = FORM_F_KARNATAKA_PART_I_HEADERS[parseInt(label, 10) - 1];
      if (canonical) label = canonical;
    }
    if (!label) continue;
    if (/^from$/i.test(label) && currentGroup === 'Days worked') label = 'Days worked From';
    else if (/^to$/i.test(label) && currentGroup === 'Days worked') label = 'Days worked To';
    else if (/^from$/i.test(label) && currentGroup === 'Leave availed') label = 'Leave availed From';
    else if (/^to$/i.test(label) && currentGroup === 'Leave availed') label = 'Leave availed To';
    else if (/no\.?\s*of\s+days/i.test(label) && currentGroup === 'Leave availed') {
      label = 'Leave availed No. of days';
    }
    headers.push(label);
  }
  return headers.filter(Boolean);
};

/**
 * Re-point parseExcelForm to PART I earned leave columns instead of the employee identity row.
 */
export function rebuildFormFKarnatakaPartIHeaders({
  jsonData,
  effectiveSheetCols,
  getMergedAwareCellText,
  headers,
  expandedHeaders,
  headerRowIndex,
  startIndex,
  tableData,
  firstSheetName,
  formHeaderInfo,
  parseHints = {},
}) {
  const fileName = parseHints?.fileName || parseHints?.formFileName || firstSheetName || '';
  const sheetTextBlob = (Array.isArray(jsonData) ? jsonData : [])
    .slice(0, 45)
    .map((row) => (Array.isArray(row) ? row : []).map((c) => String(c || '')).join(' '))
    .join(' ');
  if (!isFormFLeaveWithWagesContext(formHeaderInfo, parseHints?.item || null, fileName, sheetTextBlob)) {
    return null;
  }

  const maxCols = Math.max(effectiveSheetCols || 0, 20);
  const readRow = (r) => {
    const cells = [];
    for (let c = 0; c < maxCols; c += 1) {
      cells.push(normalizeCell(getMergedAwareCellText(r, c)));
    }
    return cells;
  };

  let partIRow = -1;
  for (let r = 0; r < Math.min(45, jsonData.length); r += 1) {
    const line = readRow(r).filter(Boolean).join(' ');
    if (isPartIEarnedLeaveBanner(line)) {
      partIRow = r;
      break;
    }
  }

  let subHeaderRow = -1;
  let groupHeaderRow = -1;
  const scanFrom = partIRow >= 0 ? partIRow + 1 : 0;
  const scanTo = Math.min(jsonData.length - 1, scanFrom + 12);
  for (let r = scanFrom; r <= scanTo; r += 1) {
    const cells = readRow(r);
    if (isEmployeeIdentityBannerRow(cells)) continue;
    if (rowLooksLikePartISubHeaderRow(cells)) {
      subHeaderRow = r;
      if (r > 0 && !rowLooksLikePartISubHeaderRow(readRow(r - 1))) {
        groupHeaderRow = r - 1;
      }
      break;
    }
  }

  if (subHeaderRow < 0) {
    for (let r = scanFrom; r <= scanTo; r += 1) {
      const cells = readRow(r);
      if (rowLooksLikePartINumberedIndexRow(cells)) {
        subHeaderRow = r;
        if (r > 0) groupHeaderRow = r - 1;
        break;
      }
    }
  }

  let rebuiltHeaders = [...FORM_F_KARNATAKA_PART_I_HEADERS];
  let newHeaderRowIndex = subHeaderRow;
  let newStartIndex = subHeaderRow >= 0 ? subHeaderRow + 1 : startIndex;

  if (subHeaderRow >= 0) {
    const lowerCells = readRow(subHeaderRow);
    const upperCells = groupHeaderRow >= 0 ? readRow(groupHeaderRow) : [];
    const combined = combinePartIHeaderBand(upperCells, lowerCells);
    if (combined.length >= 6 && headersIndicateFormFKarnatakaPartITable(combined)) {
      rebuiltHeaders = combined;
    }
    while (newStartIndex < jsonData.length) {
      const cells = readRow(newStartIndex).filter(Boolean);
      if (cells.length === 0) {
        newStartIndex += 1;
        continue;
      }
      if (isPartIEarnedLeaveBanner(cells.join(' ')) || /part\s*[-–]?\s*ii/i.test(cells.join(' ').toLowerCase())) {
        break;
      }
      if (rowLooksLikePartISubHeaderRow(cells) || rowLooksLikePartINumberedIndexRow(cells)) {
        newStartIndex += 1;
        continue;
      }
      break;
    }
  }

  const newTableData = [];
  for (let i = newStartIndex; i < jsonData.length; i += 1) {
    const cells = readRow(i);
    const line = cells.filter(Boolean).join(' ');
    if (!line) continue;
    if (/part\s*[-–]?\s*ii/i.test(formFKarnatakaHeaderNorm(line))) break;
    const rowData = {};
    rebuiltHeaders.forEach((header, idx) => {
      rowData[header] = cells[idx] || '';
    });
    const hasData = Object.values(rowData).some((v) => String(v ?? '').trim() !== '');
    if (hasData) newTableData.push(rowData);
  }

  return {
    headers: rebuiltHeaders,
    expandedHeaders: rebuiltHeaders,
    subColumnsData: {},
    headerRowIndex: newHeaderRowIndex >= 0 ? newHeaderRowIndex : headerRowIndex,
    startIndex: newStartIndex,
    tableData: newTableData,
    columnGroupLabels: null,
    originalHeaderRowIndex: newHeaderRowIndex >= 0 ? newHeaderRowIndex : headerRowIndex,
    formHeaderPatch: {
      title: formHeaderInfo?.title || 'FORM F',
      subtitle: formHeaderInfo?.subtitle || 'REGISTER OF LEAVE WITH WAGES',
      reference: formHeaderInfo?.reference || '(SEE RULE 8)',
      formFKarnatakaTableLayout: true,
    },
  };
}

export function isFormFKarnatakaLeaveEarnedHeader(header) {
  const s = formFKarnatakaHeaderNorm(header);
  return s.includes('leave') && s.includes('earned') && !s.includes('credit');
}

export function isFormFKarnatakaLeaveAvailedHeader(header) {
  const s = formFKarnatakaHeaderNorm(header);
  return s.includes('leave') && s.includes('avail');
}

export function isFormFKarnatakaLeaveBalanceHeader(header) {
  const s = formFKarnatakaHeaderNorm(header);
  return s.includes('balance') && s.includes('return') && s.includes('leave');
}
