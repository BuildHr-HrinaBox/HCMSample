import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import * as XLSX from 'xlsx';

/** Gujarat Form M — Identity Card (See Rule 16). Establishment header + worker table + per-employee ZIP. */

/** Modal table columns only — photograph stays on the Excel card, not in the grid. */
export const FORM_MGJ_GJ_CANONICAL_TABLE_HEADERS = [
  'S.No.',
  'The full name and address of the worker',
  'Date of birth of the worker',
  'Date of joining the service in the establishment',
  'Contact No.',
];

export function formMGJGujaratHeaderNorm(txt) {
  return String(txt || '')
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

const stripLeadingNumber = (s) =>
  String(s || '')
    .replace(/^\(?[a-e\d]+\)?\s*[\.\)]?\s*/i, '')
    .trim();

function normHeaderLabel(h) {
  return formMGJGujaratHeaderNorm(stripLeadingNumber(h))
    .replace(/[^a-z0-9 ]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

export function isFormMGJGujaratContext(
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

  const hasGujarat = /gujarat|\b_gj\b|form[\s._-]*m[\s._-]*gj|form_m_gj/.test(parts);
  const hasFormM = /\bform[\s._-]*m\b/.test(parts);
  const hasIdentityCard = /identity\s+card/.test(parts);

  if (hasGujarat && hasFormM) return true;
  if (hasFormM && hasIdentityCard && hasGujarat) return true;
  if (hasFormM && hasIdentityCard && /see\s+rule\s+16/.test(parts)) return true;
  if (
    hasIdentityCard &&
    /name\s+and\s+address\s+of\s+the\s+establishment/.test(parts) &&
    /full\s+name\s+and\s+address\s+of\s+the\s+worker/.test(parts) &&
    /passport\s+size\s+photograph/.test(parts)
  ) {
    return true;
  }

  return false;
}

export function isFormMGJGujaratHeaderFieldLayoutFormHeader(formHeader) {
  return !!formHeader?.formMGJGujaratHeaderFieldLayout;
}

export function isFormMGJGujaratTableLayoutFormHeader(formHeader) {
  return !!formHeader?.formMGJGujaratTableLayout;
}

export function resolveFormMGJGujaratTableHeaders(_tableHeaders) {
  return [...FORM_MGJ_GJ_CANONICAL_TABLE_HEADERS];
}

export const FORM_MGJ_TEMPLATE_SPECS = [
  {
    key: 'form_m_gj_establishment',
    label: 'Name and address of the establishment',
    group: 'header',
    fieldType: 'textarea',
    match: /name\s+and\s+address\s+of\s+the\s+establishment/i,
  },
  {
    key: 'form_m_gj_worker_name_address',
    label: 'The full name and address of the worker',
    group: 'worker',
    fieldType: 'textarea',
    tableHeader: 'The full name and address of the worker',
    match: /full\s+name\s+and\s+address\s+of\s+the\s+worker|name\s+and\s+address\s+of\s+the\s+worker/i,
  },
  {
    key: 'form_m_gj_date_of_birth',
    label: 'Date of birth of the worker',
    group: 'worker',
    fieldType: 'text',
    tableHeader: 'Date of birth of the worker',
    match: /date\s+of\s+birth\s+of\s+the\s+worker/i,
  },
  {
    key: 'form_m_gj_date_of_joining',
    label: 'Date of joining the service in the establishment',
    group: 'worker',
    fieldType: 'text',
    tableHeader: 'Date of joining the service in the establishment',
    match: /date\s+of\s+joining\s+the\s+service/i,
  },
  {
    key: 'form_m_gj_photograph',
    label: 'Recent passport size photograph of the worker',
    group: 'worker',
    fieldType: 'text',
    tableHeader: 'Recent passport size photograph of the worker',
    match: /passport\s+size\s+photograph|photograph\s+of\s+the\s+worker/i,
  },
  {
    key: 'form_m_gj_contact_no',
    label: 'Contact No.',
    group: 'worker',
    fieldType: 'text',
    tableHeader: 'Contact No.',
    match: /contact\s+no\.?/i,
  },
];

const WORKER_FIELD_KEY_BY_BUCKET = {
  workerNameAddress: 'form_m_gj_worker_name_address',
  dateOfBirth: 'form_m_gj_date_of_birth',
  dateOfJoining: 'form_m_gj_date_of_joining',
  photograph: 'form_m_gj_photograph',
  contactNo: 'form_m_gj_contact_no',
};

export function formMGJHeaderAliasBucket(norm) {
  const n = String(norm || '').trim();
  if (!n) return '';
  if (/contact\s+no|contact\s+number|^mobile$|^phone$|^contact$/.test(n)) return 'contactNo';
  if (/^sr\.?\s*no|^s\.?\s*no|^sl\.?\s*no|serial\s+number|^serial$/.test(n)) return 'sno';
  if (/full\s+name/.test(n) && /address/.test(n) && /worker/.test(n)) return 'workerNameAddress';
  if (/name/.test(n) && /address/.test(n) && /worker/.test(n)) return 'workerNameAddress';
  if (/date\s+of\s+birth/.test(n)) return 'dateOfBirth';
  if (/date\s+of\s+joining/.test(n) || (/joining/.test(n) && /service/.test(n))) return 'dateOfJoining';
  if (/photograph/.test(n) || /passport\s+size/.test(n)) return 'photograph';
  return n;
}

export function isFormMGJSerialHeader(header) {
  return formMGJHeaderAliasBucket(normHeaderLabel(header)) === 'sno';
}

export function isFormMGJWorkerNameAddressHeader(header) {
  return formMGJHeaderAliasBucket(normHeaderLabel(header)) === 'workerNameAddress';
}

export function isFormMGJDateOfBirthHeader(header) {
  return formMGJHeaderAliasBucket(normHeaderLabel(header)) === 'dateOfBirth';
}

export function isFormMGJDateOfJoiningHeader(header) {
  return formMGJHeaderAliasBucket(normHeaderLabel(header)) === 'dateOfJoining';
}

export function isFormMGJPhotographHeader(header) {
  return formMGJHeaderAliasBucket(normHeaderLabel(header)) === 'photograph';
}

export function isFormMGJContactHeader(header) {
  return formMGJHeaderAliasBucket(normHeaderLabel(header)) === 'contactNo';
}

export function isFormMGJGujaratHiddenTableHeader(header) {
  return isFormMGJPhotographHeader(header);
}

export function isFormMGJSkipPeopleAutofillHeader(header) {
  return isFormMGJPhotographHeader(header);
}

function looksLikePhoneValue(val) {
  const text = String(val ?? '').trim();
  if (!text || /^enter\b/i.test(text)) return false;
  const digits = text.replace(/\D/g, '');
  return digits.length >= 6 && digits.length <= 15;
}

export function sanitizeFormMGJGujaratMobileValue(val) {
  const text = String(val ?? '').trim();
  if (!text || /^enter\b/i.test(text)) return '';
  return looksLikePhoneValue(text) ? text : '';
}

export function remapFormMGJGujaratRowsToHeaders(rows, _sourceHeaders, _targetHeaders) {
  const hdrs = resolveFormMGJGujaratTableHeaders();
  return (Array.isArray(rows) ? rows : []).map((row, rowIndex) => {
    const out = {};
    hdrs.forEach((header) => {
      out[header] = getFormMGJGujaratRowValueForHeader(row, header, rowIndex);
    });
    if (row && typeof row === 'object') {
      Object.entries(row).forEach(([key, value]) => {
        if (String(key).startsWith('__') && out[key] == null) out[key] = value;
      });
    }
    return out;
  });
}

export function getFormMGJGujaratRowValueForHeader(row, header, rowIndex = 0) {
  if (!row || !header) return '';
  const bucket = formMGJHeaderAliasBucket(normHeaderLabel(header));
  const direct = row[header];
  if (direct != null && String(direct).trim() !== '') {
    const directText = String(direct).trim();
    if (bucket === 'contactNo') {
      return sanitizeFormMGJGujaratMobileValue(directText);
    }
    return directText;
  }
  if (bucket === 'contactNo') {
    for (const [k, v] of Object.entries(row)) {
      if (String(k).startsWith('__') || v == null || String(v).trim() === '') continue;
      if (formMGJHeaderAliasBucket(normHeaderLabel(k)) !== 'contactNo') continue;
      const text = sanitizeFormMGJGujaratMobileValue(v);
      if (text) return text;
    }
    return '';
  }
  for (const [k, v] of Object.entries(row)) {
    if (
      String(k).startsWith('__') ||
      v == null ||
      String(v).trim() === ''
    ) {
      continue;
    }
    if (formMGJHeaderAliasBucket(normHeaderLabel(k)) === bucket) {
      if (bucket === 'contactNo') {
        return sanitizeFormMGJGujaratMobileValue(v);
      }
      return String(v).trim();
    }
  }
  if (bucket === 'sno') return String(rowIndex + 1);
  return '';
}

export function rowHasMeaningfulFormMGJGujaratExportData(row, headers) {
  const hdrs = resolveFormMGJGujaratTableHeaders(headers);
  return hdrs.some((header) => {
    const bucket = formMGJHeaderAliasBucket(normHeaderLabel(header));
    if (bucket === 'sno' || bucket === 'photograph') return false;
    return getFormMGJGujaratRowValueForHeader(row, header) !== '';
  });
}

export function filterFormMGJGujaratExportRows(rows, headers) {
  const hdrs = resolveFormMGJGujaratTableHeaders(headers);
  return (Array.isArray(rows) ? rows : []).filter((row) =>
    rowHasMeaningfulFormMGJGujaratExportData(row, hdrs)
  );
}

const isNarrativeBlob = (raw) => {
  const n = formMGJGujaratHeaderNorm(raw);
  if (!n) return false;
  return (
    /^form\s*[-–]?\s*m\b/i.test(n) ||
    /identity\s+card/i.test(n) ||
    /see\s+rule\s+16/i.test(n) ||
    /^photograph$/i.test(n) ||
    n.length > 120
  );
};

const buildTemplateField = (spec, coords = {}) => ({
  label: spec.label,
  value: coords.value || '',
  key: spec.key,
  group: spec.group || 'header',
  fieldType: spec.fieldType || 'text',
  labelRow: coords.labelRow ?? null,
  labelCol: coords.labelCol ?? null,
  valueCol: coords.valueCol ?? null,
  valueRow: coords.valueRow ?? null,
});

const FORM_MGJ_DEFAULT_VALUE_COL = 3;
const FORM_MGJ_PHOTO_COL = 10;
const FORM_MGJ_VALUE_MAX_COL = 9;

const readValueBelowOrBesideLabel = (getMergedAwareCellText, labelRow, labelCol, effectiveSheetCols) => {
  const inlineValueCol = FORM_MGJ_DEFAULT_VALUE_COL;
  const v = String(getMergedAwareCellText(labelRow, inlineValueCol) || '').trim();
  if (v && !isNarrativeBlob(v) && !/^[\(\[]?[a-f][\)\]]\.?\s*$/i.test(v)) {
    return { value: v, valueCol: inlineValueCol, valueRow: labelRow };
  }
  return { value: '', valueCol: inlineValueCol, valueRow: labelRow };
};

const FORM_MGJ_EXPORT_ORDINALS = {
  form_m_gj_establishment: 'a',
  form_m_gj_worker_name_address: 'b',
  form_m_gj_date_of_birth: 'c',
  form_m_gj_date_of_joining: 'd',
  form_m_gj_photograph: 'e',
  form_m_gj_contact_no: 'f',
};

/** Form M GJ template: label col B (rows 7,9,11…), value col C on same row. */
const FORM_MGJ_CANONICAL_LABEL_POSITIONS = {
  form_m_gj_establishment: { row: 7, col: 2 },
  form_m_gj_worker_name_address: { row: 9, col: 2 },
  form_m_gj_date_of_birth: { row: 11, col: 2 },
  form_m_gj_date_of_joining: { row: 13, col: 2 },
  form_m_gj_photograph: { row: 15, col: 2 },
  form_m_gj_contact_no: { row: 17, col: 2 },
};

const FORM_MGJ_TEMPLATE_LABEL_TEXT = {
  form_m_gj_establishment: '(a) Name and address of the establishment;',
  form_m_gj_worker_name_address: '(b) The full name and address of the worker: -',
  form_m_gj_date_of_birth: '(c) Date of birth of the worker;',
  form_m_gj_date_of_joining: '(d) Date of joining the service in the establishment:',
  form_m_gj_photograph: '(e) Recent passport size photograph of the worker.',
  form_m_gj_contact_no: '(f) Contact No.',
};

const FORM_MGJ_LABEL_ROWS = [7, 9, 11, 13, 15, 17];
const FORM_MGJ_LEGACY_STACKED_ROWS = [8, 10, 12, 14, 16, 18];

const excelColLettersToNumber = (letters) => {
  let n = 0;
  for (const ch of String(letters || '').toUpperCase()) {
    n = n * 26 + (ch.charCodeAt(0) - 64);
  }
  return n;
};

const parseExcelJsMergeRanges = (worksheet) => {
  const raw = worksheet?.model?.merges;
  if (!Array.isArray(raw) || raw.length === 0) return [];
  return raw
    .map((range) => {
      const parts = String(range || '').split(':');
      if (parts.length !== 2) return null;
      const parseRef = (ref) => {
        const m = String(ref).match(/^([A-Z]+)(\d+)$/i);
        if (!m) return null;
        return { col: excelColLettersToNumber(m[1]), row: Number(m[2]) };
      };
      const start = parseRef(parts[0]);
      const end = parseRef(parts[1]);
      if (!start || !end) return null;
      return {
        top: Math.min(start.row, end.row),
        left: Math.min(start.col, end.col),
        bottom: Math.max(start.row, end.row),
        right: Math.max(start.col, end.col),
      };
    })
    .filter(Boolean);
};

const resolveMGJMergeTopLeft = (worksheet, row, col) => {
  if (!worksheet || row < 1 || col < 1) return { row, col };
  const ranges = parseExcelJsMergeRanges(worksheet);
  for (let i = 0; i < ranges.length; i += 1) {
    const m = ranges[i];
    if (row >= m.top && row <= m.bottom && col >= m.left && col <= m.right) {
      return { row: m.top, col: m.left };
    }
  }
  return { row, col };
};

const cellHasVisibleBorder = (border, side) => !!border?.[side]?.style;

const resolveMGJHeaderHalfBounds = (labelCol) => {
  const isRightHalf = labelCol >= 7;
  return {
    halfStart: isRightHalf ? 7 : 2,
    halfEnd: FORM_MGJ_VALUE_MAX_COL,
  };
};

const isMGJHeaderLabelBlob = (raw) => {
  const s = formMGJGujaratHeaderNorm(String(raw || '').replace(/:+$/, ''));
  if (!s) return false;
  return FORM_MGJ_TEMPLATE_SPECS.some(
    (spec) => spec.match.test(s) || spec.match.test(String(raw || ''))
  );
};

const isMGJHeaderLabelContinuation = (raw) => {
  const s = formMGJGujaratHeaderNorm(String(raw || '').replace(/:+$/, ''));
  if (!s) return false;
  if (/^form\s*[-–]?\s*m\b|identity\s+card|see\s+rule\s+16|^photograph$/i.test(s)) return false;
  if (/^\(?[a-f]\)?[\.\)]/i.test(s)) return false;
  return /establish|worker|joining|birth|contact|photograph|passport/i.test(s);
};

const isMGJLabelRow = (worksheet, row, labelCol, cellText) => {
  if (!worksheet || row < 1 || labelCol < 1) return false;
  const raw = String(cellText(worksheet.getCell(row, labelCol)?.value) || '').trim();
  if (!raw) return false;
  return isMGJHeaderLabelBlob(raw) || isMGJHeaderLabelContinuation(raw);
};

const ensureMGJCellRangeMerged = (worksheet, top, left, bottom, right) => {
  if (!worksheet || top < 1 || left < 1) return;
  if (top === bottom && left === right) return;
  const ranges = parseExcelJsMergeRanges(worksheet);
  const covered = ranges.some(
    (m) => top >= m.top && bottom <= m.bottom && left >= m.left && right <= m.right
  );
  if (covered) return;
  try {
    worksheet.mergeCells(top, left, bottom, right);
  } catch (_) {
    /* keep existing partial merges */
  }
};

const resolveMGJHeaderBoxArea = (worksheet, labelRow, labelCol) => {
  const { halfStart, halfEnd } = resolveMGJHeaderHalfBounds(labelCol);
  const minBoxRow = labelRow + 2;
  const ranges = parseExcelJsMergeRanges(worksheet);
  let bestMerge = null;
  ranges.forEach((m) => {
    if (m.top < minBoxRow) return;
    if (m.top > labelRow + 14) return;
    if (m.bottom - m.top < 1) return;
    const overlapLeft = Math.max(m.left, halfStart);
    const overlapRight = Math.min(m.right, halfEnd);
    if (overlapRight - overlapLeft < 1) return;
    if (!bestMerge || m.top < bestMerge.top || (m.top === bestMerge.top && m.left < bestMerge.left)) {
      bestMerge = m;
    }
  });
  if (bestMerge) {
    return {
      valueRow: bestMerge.top,
      valueCol: Math.max(bestMerge.left, halfStart),
      top: bestMerge.top,
      bottom: bestMerge.bottom,
      left: Math.max(bestMerge.left, halfStart),
      right: Math.min(bestMerge.right, halfEnd),
    };
  }
  for (let r = minBoxRow; r <= labelRow + 12; r += 1) {
    for (let c = halfStart; c <= halfEnd; c += 1) {
      const b = worksheet.getCell(r, c)?.border || {};
      if (!cellHasVisibleBorder(b, 'top') && !cellHasVisibleBorder(b, 'left')) continue;
      let bottom = r;
      let right = c;
      for (let br = r; br <= r + 8; br += 1) {
        const bb = worksheet.getCell(br, c)?.border || {};
        if (cellHasVisibleBorder(bb, 'left') || cellHasVisibleBorder(bb, 'right')) bottom = br;
        else if (br > r + 1) break;
      }
      for (let bc = c; bc <= halfEnd; bc += 1) {
        const bb = worksheet.getCell(r, bc)?.border || {};
        if (cellHasVisibleBorder(bb, 'top') || cellHasVisibleBorder(bb, 'bottom')) right = bc;
        else if (bc > c + 1) break;
      }
      return {
        valueRow: r,
        valueCol: Math.max(halfStart, labelCol),
        top: r,
        bottom,
        left: Math.max(halfStart, labelCol),
        right: Math.max(right, halfEnd),
      };
    }
  }
  return {
    valueRow: minBoxRow,
    valueCol: halfStart,
    top: minBoxRow,
    bottom: minBoxRow + 3,
    left: halfStart,
    right: halfEnd,
  };
};

const clearMGJHeaderStrayAboveBox = (worksheet, labelRow, labelCol, boxTop, cellText) => {
  if (!worksheet || labelRow < 1 || boxTop <= labelRow) return;
  const { halfStart, halfEnd } = resolveMGJHeaderHalfBounds(labelCol);
  for (let r = labelRow; r < boxTop; r += 1) {
    for (let c = halfStart; c <= halfEnd; c += 1) {
      if (r === labelRow && c === labelCol) continue;
      const raw = String(cellText(worksheet.getCell(r, c)?.value) || '').trim();
      if (!raw) continue;
      if (c === labelCol && (isMGJHeaderLabelBlob(raw) || isMGJHeaderLabelContinuation(raw))) continue;
      if (r === labelRow && isMGJHeaderLabelBlob(raw)) continue;
      worksheet.getCell(r, c).value = '';
    }
  }
};

const clearMGJHeaderRowSpill = (worksheet, labelRow, labelCol, cellText) => {
  if (!worksheet || labelRow < 1 || labelCol < 1) return;
  for (let c = labelCol + 1; c <= FORM_MGJ_VALUE_MAX_COL; c += 1) {
    const raw = String(cellText(worksheet.getCell(labelRow, c)?.value) || '').trim();
    if (!raw) continue;
    if (isMGJHeaderLabelBlob(raw)) break;
    worksheet.getCell(labelRow, c).value = '';
  }
};

const clearMGJHeaderBoxBand = (worksheet, labelRow, labelCol, area, cellText) => {
  if (!worksheet || !area || labelRow < 1) return;
  let startRow = area.top;
  const endRow = area.bottom;
  const startCol = area.left;
  const endCol = area.right;
  if (startRow <= labelRow) startRow = labelRow + 2;
  while (startRow <= endRow && isMGJLabelRow(worksheet, startRow, labelCol, cellText)) {
    startRow += 1;
  }
  if (startRow > endRow) return;
  for (let r = startRow; r <= endRow; r += 1) {
    if (isMGJLabelRow(worksheet, r, labelCol, cellText)) continue;
    for (let c = startCol; c <= endCol; c += 1) {
      worksheet.getCell(r, c).value = '';
    }
  }
};

const clearMGJHeaderSpillZones = (worksheet, cellText) => {
  if (!worksheet) return;
  for (let r = 1; r <= 28; r += 1) {
    for (let c = FORM_MGJ_PHOTO_COL; c <= 12; c += 1) {
      const raw = String(cellText(worksheet.getCell(r, c)?.value) || '').trim();
      if (!raw) continue;
      if (/^photograph$/i.test(raw)) continue;
      if (/^form\s*[-–]?\s*m\b|identity\s+card|see\s+rule\s+16/i.test(raw)) continue;
      if (isMGJHeaderLabelBlob(raw)) continue;
      worksheet.getCell(r, c).value = '';
    }
  }
};

const findMGJValueMergeOnRow = (worksheet, row, labelCol, cellText) => {
  const ranges = parseExcelJsMergeRanges(worksheet);
  let best = null;
  ranges.forEach((m) => {
    if (row < m.top || row > m.bottom) return;
    if (m.left <= labelCol) return;
    if (m.right <= m.left) return;
    if (!best || m.left < best.col) best = { row, col: m.left };
  });
  if (best) return best;
  for (let c = labelCol + 1; c <= FORM_MGJ_VALUE_MAX_COL; c += 1) {
    const t = String(cellText(worksheet.getCell(row, c)?.value) || '').trim();
    if (!t || /^[_\s.\-:…]+$/u.test(t)) return { row, col: c };
  }
  return { row, col: Math.max(labelCol + 1, FORM_MGJ_DEFAULT_VALUE_COL) };
};

const resolveMGJValueBounds = (worksheet, row, labelCol, cellText) => {
  const { col: valueCol } = findMGJValueMergeOnRow(worksheet, row, labelCol, cellText);
  const topLeft = resolveMGJMergeTopLeft(worksheet, row, valueCol);
  const ranges = parseExcelJsMergeRanges(worksheet);
  const merge = ranges.find(
    (m) =>
      topLeft.row >= m.top &&
      topLeft.row <= m.bottom &&
      topLeft.col >= m.left &&
      topLeft.col <= m.right
  );
  if (merge) {
    return {
      top: merge.top,
      bottom: merge.bottom,
      left: merge.left,
      right: Math.min(merge.right, FORM_MGJ_VALUE_MAX_COL),
    };
  }
  const left = Math.max(topLeft.col, labelCol + 1);
  let right = left;
  for (let c = left + 1; c <= FORM_MGJ_VALUE_MAX_COL; c += 1) {
    const t = String(cellText(worksheet.getCell(row, c)?.value) || '').trim();
    if (!t || /^[_\s.\-:…]+$/u.test(t)) right = c;
    else break;
  }
  if (right === left) right = Math.min(left + 5, FORM_MGJ_VALUE_MAX_COL);
  return { top: row, bottom: row, left, right };
};

const clearMGJValueBand = (worksheet, bounds) => {
  if (!worksheet || !bounds) return;
  for (let r = bounds.top; r <= bounds.bottom; r += 1) {
    for (let c = bounds.left; c <= bounds.right; c += 1) {
      worksheet.getCell(r, c).value = '';
    }
  }
};

const readMGJLabelWithContinuation = (worksheet, row, col, cellText) => {
  let text = formMGJGujaratHeaderNorm(
    String(cellText(worksheet.getCell(row, col)?.value) || '').replace(/:+$/, '')
  );
  if (!text) return '';
  const next = formMGJGujaratHeaderNorm(
    String(cellText(worksheet.getCell(row + 1, col)?.value) || '').replace(/:+$/, '')
  );
  if (
    next &&
    !/^\(?[a-f]\)?[\.\)]/i.test(next) &&
    !/^form\s*[-–]?\s*m\b|identity\s+card|see\s+rule\s+16/i.test(next) &&
    (text.includes('(') && !text.includes(')') || /establish|worker|joining|birth|contact|photograph/i.test(next))
  ) {
    text = `${text} ${next}`.trim();
  }
  return text;
};

const matchesFormMGJSpecLabel = (spec, cellStr) => {
  const s = formMGJGujaratHeaderNorm(String(cellStr || '').replace(/:+$/, ''));
  if (!s) return false;
  if (spec.match.test(s) || spec.match.test(cellStr)) return true;
  const ordinal = FORM_MGJ_EXPORT_ORDINALS[spec.key];
  if (ordinal && new RegExp(`^\\(?${ordinal}\\)?(?:[\\.)]|\\s|$)`, 'i').test(s)) {
    return spec.match.test(s.replace(/^\(?[a-f]\)?[\.\)]?\s*/i, '')) || spec.match.test(s);
  }
  return false;
};

const findFormMGJLabelPosition = (worksheet, spec, _parsedFields, cellText) => {
  const ordinal = FORM_MGJ_EXPORT_ORDINALS[spec.key];
  if (ordinal) {
    for (let r = 1; r <= 30; r += 1) {
      const raw = String(cellText(worksheet.getCell(r, 2)?.value) || '').trim();
      if (!raw) continue;
      if (new RegExp(`^\\(?${ordinal}\\)?(?:[\\.)]|\\s|$)`, 'i').test(raw)) {
        return { row: r, col: 2 };
      }
    }
  }
  for (let r = 1; r <= 30; r += 1) {
    for (let c = 1; c <= 4; c += 1) {
      const raw = String(cellText(worksheet.getCell(r, c)?.value) || '').replace(/:+$/, '').trim();
      if (!raw || !matchesFormMGJSpecLabel(spec, raw)) continue;
      return { row: r, col: c };
    }
  }
  return FORM_MGJ_CANONICAL_LABEL_POSITIONS[spec.key] || null;
};

const restoreMGJTemplateLabels = (worksheet) => {
  if (!worksheet) return;
  Object.entries(FORM_MGJ_TEMPLATE_LABEL_TEXT).forEach(([key, labelText]) => {
    const pos = FORM_MGJ_CANONICAL_LABEL_POSITIONS[key];
    if (!pos || !labelText) return;
    worksheet.getCell(pos.row, pos.col).value = labelText;
  });
};

const resolveMGJInlineValueBounds = (worksheet, labelRow) => {
  const ranges = parseExcelJsMergeRanges(worksheet);
  let merge = ranges.find(
    (m) =>
      labelRow >= m.top &&
      labelRow <= m.bottom &&
      m.left >= FORM_MGJ_DEFAULT_VALUE_COL &&
      m.right <= FORM_MGJ_VALUE_MAX_COL
  );
  if (merge) {
    return {
      top: merge.top,
      bottom: merge.bottom,
      left: Math.max(merge.left, FORM_MGJ_DEFAULT_VALUE_COL),
      right: Math.min(merge.right, FORM_MGJ_VALUE_MAX_COL),
    };
  }
  return {
    top: labelRow,
    bottom: labelRow,
    left: FORM_MGJ_DEFAULT_VALUE_COL,
    right: FORM_MGJ_VALUE_MAX_COL,
  };
};

const clearMGJExportValueSpill = (worksheet, cellText) => {
  if (!worksheet) return;
  FORM_MGJ_LEGACY_STACKED_ROWS.forEach((row) => {
    for (let c = 2; c <= FORM_MGJ_VALUE_MAX_COL; c += 1) {
      worksheet.getCell(row, c).value = '';
    }
  });
  FORM_MGJ_LABEL_ROWS.forEach((row) => {
    for (let c = FORM_MGJ_DEFAULT_VALUE_COL; c <= FORM_MGJ_VALUE_MAX_COL; c += 1) {
      const raw = String(cellText(worksheet.getCell(row, c)?.value) || '').trim();
      if (!raw) continue;
      if (isMGJHeaderLabelBlob(raw)) continue;
      worksheet.getCell(row, c).value = '';
    }
  });
};

const setMGJCellValue = (worksheet, row, col, value, options = {}) => {
  const { wrap = false, inBox = false } = options;
  const text = String(value ?? '').trim();
  const writePos = resolveMGJMergeTopLeft(worksheet, row, col);
  const cell = worksheet.getCell(writePos.row, writePos.col);
  if (!text) {
    cell.value = '';
    return;
  }
  cell.value = text;
  cell.alignment = {
    ...(cell.alignment || {}),
    horizontal: 'left',
    vertical: inBox || wrap || text.includes('\n') ? 'top' : 'middle',
    wrapText: inBox || wrap || text.includes('\n') || text.length > 36,
    shrinkToFit: false,
  };
};

const writeMGJInlineFieldValue = (worksheet, labelRow, value, options = {}) => {
  const { wrap = false } = options;
  const bounds = resolveMGJInlineValueBounds(worksheet, labelRow);
  clearMGJValueBand(worksheet, bounds);
  ensureMGJCellRangeMerged(worksheet, bounds.top, bounds.left, bounds.bottom, bounds.right);
  setMGJCellValue(worksheet, bounds.top, bounds.left, value, { wrap, inBox: wrap });
};

const clearMGJWorkmanValueBand = (worksheet, row, labelCol, cellText) => {
  if (!worksheet || row < 1) return;
  const bounds = resolveMGJValueBounds(worksheet, row, labelCol, cellText);
  clearMGJValueBand(worksheet, bounds);
};

const writeMGJBoxedHeaderValue = (worksheet, labelRow, labelCol, value, cellText) => {
  const area = resolveMGJHeaderBoxArea(worksheet, labelRow, labelCol);
  clearMGJHeaderStrayAboveBox(worksheet, labelRow, labelCol, area.top, cellText);
  clearMGJHeaderRowSpill(worksheet, labelRow, labelCol, cellText);
  clearMGJHeaderBoxBand(worksheet, labelRow, labelCol, area, cellText);
  ensureMGJCellRangeMerged(worksheet, area.top, area.left, area.bottom, area.right);
  setMGJCellValue(worksheet, area.top, area.left, value, { wrap: true, inBox: true });
};

const findHeaderLabelCell = (getMergedAwareCellText, matchRe, effectiveSheetCols, maxRows = 80) => {
  const maxC = Math.max(20, effectiveSheetCols || 0);
  for (let r = 0; r < maxRows; r += 1) {
    for (let c = 0; c < maxC; c += 1) {
      const raw = String(getMergedAwareCellText(r, c) || '').trim();
      if (!raw || isNarrativeBlob(raw)) continue;
      const norm = formMGJGujaratHeaderNorm(raw);
      if (!matchRe.test(norm) && !matchRe.test(raw)) continue;
      const read = readValueBelowOrBesideLabel(getMergedAwareCellText, r, c, effectiveSheetCols);
      return {
        labelRow: r,
        labelCol: c,
        valueCol: read.valueCol,
        valueRow: read.valueRow,
        value: read.value,
      };
    }
  }
  return null;
};

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

export function buildFormMGJTemplateFields(getMergedAwareCellText = null, effectiveSheetCols = 20) {
  return FORM_MGJ_TEMPLATE_SPECS.map((spec) => {
    let coords = {};
    if (typeof getMergedAwareCellText === 'function' && spec.match) {
      coords = findHeaderLabelCell(getMergedAwareCellText, spec.match, effectiveSheetCols) || {};
    }
    return buildTemplateField(spec, coords);
  });
}

export function resolveFormMGJGujaratHeaderFieldLayout(parsed, workbook, hints = {}) {
  const formHeader = hints.formHeader || parsed?.formHeader || null;
  const item = hints.item || null;
  const fileName = hints.fileName || hints.formFileName || '';
  const sheetText = hints.sheetText || '';
  if (!isFormMGJGujaratContext(formHeader, item, fileName, sheetText)) return null;

  const accessor = buildWorkbookMergedCellAccessor(workbook, hints);
  const getMergedAwareCellText = accessor?.getMergedAwareCellText ?? null;
  const effectiveSheetCols = accessor?.effectiveSheetCols ?? 20;
  const allTemplateFields = buildFormMGJTemplateFields(getMergedAwareCellText, effectiveSheetCols);
  let uiFields = allTemplateFields.filter((field) => field.group === 'header');
  if (uiFields.length === 0) {
    uiFields = [buildTemplateField(FORM_MGJ_TEMPLATE_SPECS[0])];
  }

  return {
    formHeader: {
      title: formHeader?.title || 'FORM - M',
      subtitle: formHeader?.subtitle || 'IDENTITY CARD',
      reference: formHeader?.reference || '(See Rule 16)',
      formMGJGujaratHeaderFieldLayout: true,
      formMGJGujaratTableLayout: true,
      textRows: [],
      fields: uiFields,
      templateFields: allTemplateFields,
    },
    headers: resolveFormMGJGujaratTableHeaders(parsed?.headers || hints.tableHeaders || []),
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
  const fn = String(src?.FirstName || src?.['First Name'] || src?.firstName || '').trim();
  const ln = String(src?.LastName || src?.['Last Name'] || src?.lastName || '').trim();
  if (fn && ln) return `${fn} ${ln}`;
  return fn || ln || String(src?.DisplayName || src?.['Display Name'] || src?.Employee_Name || '').trim();
}

function readEmployeeAddress(emp) {
  const src = unwrapEmployeeRecord(emp);
  return String(
    src?.PresentAddress ||
      src?.['Present Address'] ||
      src?.PermanentAddress ||
      src?.['Permanent Address'] ||
      src?.Address ||
      src?.['Address'] ||
      ''
  ).trim();
}

function cellIsEmpty(value) {
  const v = String(value ?? '').trim();
  if (!v) return true;
  const s = v.toLowerCase();
  return /^enter\b/.test(s) || s.includes('enter ');
}

function pickScalarEmployeeValue(raw) {
  if (raw == null || raw === '') return '';
  if (typeof raw === 'object') {
    return String(
      raw.display_value ??
        raw.displayValue ??
        raw.name ??
        raw.Name ??
        raw.value ??
        raw.text ??
        ''
    ).trim();
  }
  return String(raw).trim();
}

function collectEmployeeMobileSources(emp) {
  const sources = [];
  const seen = new Set();
  const push = (value) => {
    if (!value || typeof value !== 'object' || seen.has(value)) return;
    seen.add(value);
    sources.push(value);
  };
  push(emp);
  if (emp && typeof emp === 'object') {
    push(emp.Employee);
    push(emp.employee);
    push(unwrapEmployeeRecord(emp));
  }
  return sources;
}

function readMobileFromEmployeeSource(src) {
  if (!src || typeof src !== 'object') return '';
  // Zoho People Mobile only — do not fall back to Phone / emergency / nested contacts.
  const keys = ['Mobile', 'mobile', 'Mobile_Number', 'Mobile_Number1', 'Mobile Number', 'MobileNumber'];
  for (let i = 0; i < keys.length; i += 1) {
    if (!Object.prototype.hasOwnProperty.call(src, keys[i])) continue;
    const val = sanitizeFormMGJGujaratMobileValue(pickScalarEmployeeValue(src[keys[i]]));
    if (val) return val;
  }
  for (const [k, v] of Object.entries(src)) {
    const kn = String(k || '')
      .toLowerCase()
      .replace(/[^a-z0-9]/g, '');
    if (kn === 'mobile' || kn === 'mobilenumber' || kn === 'mobilenumber1') {
      const val = sanitizeFormMGJGujaratMobileValue(pickScalarEmployeeValue(v));
      if (val) return val;
    }
  }
  return '';
}

/** Contact No. uses Zoho People Mobile only; blank when missing or not a phone number. */
export function readFormMGJGujaratEmployeeMobile(emp) {
  const sources = collectEmployeeMobileSources(emp);
  for (let i = 0; i < sources.length; i += 1) {
    const val = readMobileFromEmployeeSource(sources[i]);
    if (val) return val;
  }
  return '';
}

function readScalarFromEmployeeSources(emp, keys, depthKeys = []) {
  const sources = collectEmployeeMobileSources(emp);
  for (let i = 0; i < sources.length; i += 1) {
    const src = sources[i];
    for (let k = 0; k < keys.length; k += 1) {
      const val = pickScalarEmployeeValue(src[keys[k]]);
      if (val) return val;
    }
    for (const [k, v] of Object.entries(src)) {
      const kn = String(k || '')
        .toLowerCase()
        .replace(/[^a-z0-9]/g, '');
      if (depthKeys.some((needle) => kn.includes(needle))) {
        const val = pickScalarEmployeeValue(v);
        if (val) return val;
      }
    }
  }
  return '';
}

export function readFormMGJGujaratEmployeeDateOfBirth(emp) {
  return readScalarFromEmployeeSources(
    emp,
    [
      'Date_of_birth',
      'Date_of_Birth',
      'DateofBirth',
      'Dateofbirth',
      'Date of Birth',
      'DateOfBirth',
      'DOB',
      'BirthDate',
      'Birth_Date',
    ],
    ['dateofbirth', 'dob', 'birthdate']
  );
}

export function readFormMGJGujaratEmployeeDateOfJoining(emp) {
  return readScalarFromEmployeeSources(
    emp,
    [
      'Dateofjoining',
      'Date_of_Joining',
      'Date of Joining',
      'DateofJoining',
      'DateOfJoining',
      'Date_of_joining',
      'DOJ',
      'JoiningDate',
    ],
    ['dateofjoining', 'joiningdate', 'doj']
  );
}

/** Force Contact No. from Mobile after overlay/remap — never keep names like "father". */
export function enrichFormMGJGujaratContactRows(mappedData, employees, headers, options = {}) {
  if (!Array.isArray(mappedData) || mappedData.length === 0) return mappedData;
  const hdrs = resolveFormMGJGujaratTableHeaders(headers);
  const contactHeader = hdrs.find((header) => isFormMGJContactHeader(header));
  if (!contactHeader) return mappedData;
  const { rowIndexOffset = 0 } = options;
  const employeeList = Array.isArray(employees) ? employees : [];
  mappedData.forEach((row, localIdx) => {
    if (!row || typeof row !== 'object') return;
    const absoluteIdx = rowIndexOffset + localIdx;
    const empItem = employeeList[absoluteIdx] || employeeList[localIdx];
    row.__employeeIndex = absoluteIdx;
    if (empItem && !row.__employeeLookupName) {
      const name = readEmployeeFullName(empItem);
      if (name) row.__employeeLookupName = name;
    }
    row[contactHeader] = empItem ? readFormMGJGujaratEmployeeMobile(empItem) : '';
  });
  return mappedData;
}

/** Remap grid rows for display/save and optionally fill Contact No. from People Mobile. */
export function resolveFormMGJGujaratDisplayRows(rows, sourceHeaders, targetHeaders, employees = null, options = {}) {
  const remapped = remapFormMGJGujaratRowsToHeaders(rows, sourceHeaders, targetHeaders);
  if (!Array.isArray(employees) || employees.length === 0) return remapped;
  const { rowIndexOffset = 0 } = options;
  enrichFormMGJGujaratContactRows(remapped, employees, targetHeaders, { rowIndexOffset });
  return remapFormMGJGujaratRowsToHeaders(remapped, targetHeaders, targetHeaders);
}

function readEmployeeContact(emp) {
  return readFormMGJGujaratEmployeeMobile(emp);
}

export function formatFormMGJGujaratEstablishmentFromSite(site) {
  if (!site || typeof site !== 'object') return '';
  const name = String(site.siteName ?? site.SiteName ?? '').trim();
  const addr = String(site.siteAddress ?? site.SiteAddress ?? '').trim();
  const city = String(site.siteCity ?? site.SiteCity ?? '').trim();
  const state = String(site.siteState ?? site.SiteState ?? '').trim();
  const addrLine = [addr, city, state].filter(Boolean).join(', ');
  return [name, addrLine].filter(Boolean).join('\n').trim();
}

export function applyFormMGJGujaratAutofillFromSite(headerData, siteContext = {}, options = {}) {
  const { onlyEmpty = false } = options;
  const out = headerData && typeof headerData === 'object' ? { ...headerData } : {};
  const establishmentText = String(siteContext.establishmentText || '').trim();
  if (establishmentText) {
    const cur = String(out.form_m_gj_establishment ?? '').trim();
    if (!onlyEmpty || !cur || /^enter\b/i.test(cur)) {
      out.form_m_gj_establishment = establishmentText;
    }
  }
  return out;
}

export function applyFormMGJGujaratAutofillFromEmployee(headerData, empItem, siteContext = {}, helpers = {}) {
  const { formatStatutoryDateDisplay = (v) => String(v || '').trim() } = helpers;
  const out = applyFormMGJGujaratAutofillFromSite(headerData, siteContext);
  const emp = unwrapEmployeeRecord(empItem);
  const fullName = readEmployeeFullName(emp);
  const address = readEmployeeAddress(emp);
  const workerLine = [fullName, address].filter(Boolean).join('\n');
  const dobRaw = readFormMGJGujaratEmployeeDateOfBirth(empItem);
  const dojRaw = readFormMGJGujaratEmployeeDateOfJoining(empItem);

  const fill = (key, value) => {
    const text = String(value ?? '').trim();
    if (!text || !key) return;
    const cur = String(out[key] ?? '').trim();
    if (!cur || /^enter\b/i.test(cur)) out[key] = text;
  };

  fill('form_m_gj_worker_name_address', workerLine);
  // Always sync dates from People — blank when missing (never keep shared header data).
  out.form_m_gj_date_of_birth = dobRaw ? formatStatutoryDateDisplay(dobRaw) : '';
  out.form_m_gj_date_of_joining = dojRaw ? formatStatutoryDateDisplay(dojRaw) : '';
  // Always sync Contact No. from People Mobile — blank when missing (never keep shared header data).
  out.form_m_gj_contact_no = readFormMGJGujaratEmployeeMobile(empItem);
  return out;
}

export function applyFormMGJGujaratEmployeeToRow(row, emp, headers, helpers = {}) {
  const hdrs = resolveFormMGJGujaratTableHeaders(headers);
  const out = row && typeof row === 'object' ? { ...row } : {};
  const {
    sanitizeValue = (v) => String(v ?? '').trim(),
    overwrite = true,
    rowIndex = 0,
    formatStatutoryDateDisplay = (v) => String(v || '').trim(),
  } = helpers;

  const shouldSet = (header) => overwrite || cellIsEmpty(out[header]);
  const setCell = (header, value) => {
    if (!header || !shouldSet(header)) return;
    out[header] = sanitizeValue(value ?? '');
  };

  const fullName = readEmployeeFullName(emp);
  const address = readEmployeeAddress(emp);
  const workerLine = [fullName, address].filter(Boolean).join('\n');
  const dobRaw = readFormMGJGujaratEmployeeDateOfBirth(emp);
  const dojRaw = readFormMGJGujaratEmployeeDateOfJoining(emp);
  const dob = dobRaw ? formatStatutoryDateDisplay(dobRaw) : '';
  const doj = dojRaw ? formatStatutoryDateDisplay(dojRaw) : '';
  if (fullName) out.__employeeLookupName = fullName;
  out.__employeeIndex = rowIndex;

  hdrs.forEach((header) => {
    if (isFormMGJSkipPeopleAutofillHeader(header)) {
      setCell(header, '');
      return;
    }
    if (isFormMGJSerialHeader(header)) {
      setCell(header, String(rowIndex + 1));
      return;
    }
    if (isFormMGJWorkerNameAddressHeader(header)) {
      setCell(header, workerLine);
      return;
    }
    if (isFormMGJDateOfBirthHeader(header)) {
      setCell(header, dob);
      return;
    }
    if (isFormMGJDateOfJoiningHeader(header)) {
      setCell(header, doj);
      return;
    }
    if (isFormMGJContactHeader(header)) {
      setCell(header, readFormMGJGujaratEmployeeMobile(emp));
    }
  });

  return out;
}

export function mapFormMGJGujaratRowsFromEmployees(employees, headers, helpers = {}) {
  const hdrs = resolveFormMGJGujaratTableHeaders(headers);
  const list = Array.isArray(employees) ? employees : [];
  return list.map((emp, rowIndex) =>
    applyFormMGJGujaratEmployeeToRow({}, emp, hdrs, { ...helpers, rowIndex })
  );
}

export function buildFormMGJGujaratHeaderFormDataForRow(baseHeaderData, row, headers, options = {}) {
  const hdrs = resolveFormMGJGujaratTableHeaders(headers);
  const {
    empItem = null,
    siteContext = {},
    formatStatutoryDateDisplay = (v) => String(v || '').trim(),
  } = options;
  let out = baseHeaderData && typeof baseHeaderData === 'object' ? { ...baseHeaderData } : {};
  // Drop shared worker fields before per-employee sync so download matches the autofill row.
  out.form_m_gj_date_of_birth = '';
  out.form_m_gj_date_of_joining = '';
  out.form_m_gj_contact_no = '';
  if (empItem) {
    out = applyFormMGJGujaratAutofillFromEmployee(out, empItem, siteContext, {
      formatStatutoryDateDisplay,
    });
  }
  hdrs.forEach((header) => {
    const bucket = formMGJHeaderAliasBucket(normHeaderLabel(header));
    const key = WORKER_FIELD_KEY_BY_BUCKET[bucket];
    if (!key) return;
    if (bucket === 'contactNo') {
      // Per-employee Contact No. must mirror People Mobile only; never keep shared headerFormData.
      out[key] = empItem
        ? readFormMGJGujaratEmployeeMobile(empItem)
        : sanitizeFormMGJGujaratMobileValue(getFormMGJGujaratRowValueForHeader(row, header));
      return;
    }
    if (bucket === 'dateOfBirth' || bucket === 'dateOfJoining') {
      // Prefer the autofill grid value (what the user sees), then People.
      const fromRow = String(getFormMGJGujaratRowValueForHeader(row, header) || '').trim();
      const fromEmp = empItem
        ? String(
            (bucket === 'dateOfBirth'
              ? readFormMGJGujaratEmployeeDateOfBirth(empItem)
              : readFormMGJGujaratEmployeeDateOfJoining(empItem)) || ''
          ).trim()
        : '';
      const raw = fromRow || fromEmp;
      out[key] = raw ? formatStatutoryDateDisplay(raw) : '';
      return;
    }
    let val = getFormMGJGujaratRowValueForHeader(row, header);
    if (val !== '') out[key] = val;
  });
  return out;
}

function resolveFormMGJGujaratEmployeeForExportRow(row, employees, index) {
  const list = Array.isArray(employees) ? employees : [];
  if (!list.length) return null;
  const storedIdx = Number(row?.__employeeIndex);
  if (Number.isFinite(storedIdx) && storedIdx >= 0 && list[storedIdx]) {
    return list[storedIdx];
  }
  const lookupName = String(row?.__employeeLookupName || '').trim().toLowerCase();
  if (lookupName) {
    const matched = list.find((emp) => readEmployeeFullName(emp).trim().toLowerCase() === lookupName);
    if (matched) return matched;
  }
  // __employeeIndex may be absolute while employeesOverride is a page slice — fall back to row order.
  if (Number.isFinite(storedIdx) && storedIdx >= list.length && list[index]) {
    return list[index];
  }
  return list[index] || null;
}

function resolveFormMGJGujaratExportFields(parsedFormHeader) {
  const templateFields = parsedFormHeader?.templateFields;
  if (Array.isArray(templateFields) && templateFields.length > 0) return templateFields;
  return Array.isArray(parsedFormHeader?.fields) ? parsedFormHeader.fields : [];
}

export function writeFormMGJGujaratFieldsToExcelJsWorksheet(worksheet, headerFormData, parsedFormHeader, helpers = {}) {
  if (!worksheet || !headerFormData || typeof headerFormData !== 'object') return;
  const excelCellValueToString =
    helpers.excelCellValueToString ||
    ((val) => {
      if (val == null) return '';
      if (typeof val === 'string' || typeof val === 'number' || typeof val === 'boolean') return String(val);
      if (val instanceof Date) return val.toISOString();
      if (typeof val === 'object') {
        if (Array.isArray(val.richText)) return val.richText.map((rt) => rt?.text || '').join('');
        if (val.text != null) return String(val.text);
        if (val.result != null) return String(val.result);
      }
      return '';
    });
  const parsedFields = resolveFormMGJGujaratExportFields(parsedFormHeader);

  restoreMGJTemplateLabels(worksheet);
  clearMGJExportValueSpill(worksheet, excelCellValueToString);
  clearMGJHeaderSpillZones(worksheet, excelCellValueToString);

  FORM_MGJ_TEMPLATE_SPECS.forEach((spec) => {
    if (spec.key === 'form_m_gj_photograph') return;
    const val = headerFormData[spec.key];
    if (val == null || String(val).trim() === '') return;

    const labelPos = findFormMGJLabelPosition(worksheet, spec, parsedFields, excelCellValueToString);
    if (!labelPos) return;

    writeMGJInlineFieldValue(worksheet, labelPos.row, val, {
      wrap:
        spec.group === 'header' ||
        String(val).includes('\n') ||
        spec.fieldType === 'textarea',
    });
  });
}

export async function buildFormMGJGujaratWorkbookWithTemplateStyles({
  templateArrayBuffer,
  parsedFormHeader,
  headerFormData,
  formFileName,
  sheetNameHint,
  excelCellValueToString = null,
}) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(templateArrayBuffer);
  const sheetCandidates = Array.isArray(workbook.worksheets) ? workbook.worksheets : [];
  const preferred = String(sheetNameHint || '').trim();
  const worksheet =
    (preferred && sheetCandidates.find((ws) => String(ws?.name || '').trim() === preferred)) ||
    sheetCandidates[0] ||
    null;
  if (!worksheet) throw new Error('Template worksheet not found.');

  writeFormMGJGujaratFieldsToExcelJsWorksheet(worksheet, headerFormData, parsedFormHeader, {
    excelCellValueToString: excelCellValueToString || undefined,
  });

  const out = await workbook.xlsx.writeBuffer();
  const fileName =
    formFileName ||
    parsedFormHeader?.title?.replace(/[^a-zA-Z0-9]/g, '_') ||
    'Form_M_GJ_Gujarat.xlsx';
  return {
    blob: new Blob([out], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    }),
    fileName,
  };
}

function sanitizeFormMGJGujaratDownloadFileName(name) {
  return String(name || 'Employee')
    .replace(/[\\/:*?"<>|]+/g, '_')
    .replace(/\s+/g, '_')
    .replace(/_+/g, '_')
    .replace(/^_|_$/g, '')
    .slice(0, 80) || 'Employee';
}

export function resolveFormMGJGujaratEmployeeDownloadBaseName(row, headers, rowIndex = 0) {
  const hdrs = resolveFormMGJGujaratTableHeaders(headers);
  const lookupName = String(row?.__employeeLookupName || '').trim();
  if (lookupName) return sanitizeFormMGJGujaratDownloadFileName(lookupName);
  for (const header of hdrs) {
    if (!isFormMGJWorkerNameAddressHeader(header)) continue;
    const val = getFormMGJGujaratRowValueForHeader(row, header, rowIndex);
    if (!val) continue;
    const firstLine = val.split(/\r?\n/)[0].trim();
    if (firstLine) return sanitizeFormMGJGujaratDownloadFileName(firstLine);
  }
  return sanitizeFormMGJGujaratDownloadFileName(`Employee_${rowIndex + 1}`);
}

function allocateUniqueFormMGJGujaratDownloadFileName(baseName, usedNames) {
  const safeBase = sanitizeFormMGJGujaratDownloadFileName(baseName);
  const count = usedNames.get(safeBase) || 0;
  usedNames.set(safeBase, count + 1);
  if (count === 0) return `${safeBase}.xlsx`;
  return `${safeBase}_${count + 1}.xlsx`;
}

export async function buildFormMGJGujaratPerEmployeeDownload({
  templateArrayBuffer,
  mappedData,
  headersToUse,
  parsedFormHeader,
  formFileName,
  headerFormData,
  sheetNameHint,
  excelCellValueToString = null,
  employeesOverride = null,
  formatStatutoryDateDisplay = null,
  siteContext = null,
}) {
  const hdrs = resolveFormMGJGujaratTableHeaders(headersToUse);
  const formatDate = formatStatutoryDateDisplay || ((v) => String(v || '').trim());
  const siteCtx = siteContext && typeof siteContext === 'object' ? siteContext : {};
  let exportRows = filterFormMGJGujaratExportRows(mappedData, hdrs);
  if (
    exportRows.length === 0 &&
    Array.isArray(employeesOverride) &&
    employeesOverride.length > 0
  ) {
    exportRows = mapFormMGJGujaratRowsFromEmployees(employeesOverride, hdrs, {
      formatStatutoryDateDisplay: formatDate,
    });
    exportRows = filterFormMGJGujaratExportRows(exportRows, hdrs);
  }

  const buildMergedHeaderData = (row, index) =>
    buildFormMGJGujaratHeaderFormDataForRow(headerFormData, row, hdrs, {
      empItem: resolveFormMGJGujaratEmployeeForExportRow(row, employeesOverride, index),
      siteContext: siteCtx,
      formatStatutoryDateDisplay: formatDate,
    });

  const exportFormHeader = {
    ...(parsedFormHeader || {}),
    fields: resolveFormMGJGujaratExportFields(parsedFormHeader),
  };
  const workbookArgs = {
    templateArrayBuffer,
    parsedFormHeader: exportFormHeader,
    formFileName,
    sheetNameHint,
    excelCellValueToString,
  };

  if (exportRows.length <= 1) {
    const mergedHeaderData = buildMergedHeaderData(exportRows[0] || {}, 0);
    return buildFormMGJGujaratWorkbookWithTemplateStyles({
      ...workbookArgs,
      headerFormData: mergedHeaderData,
    });
  }

  const zip = new JSZip();
  const usedNames = new Map();
  for (let i = 0; i < exportRows.length; i += 1) {
    const mergedHeaderData = buildMergedHeaderData(exportRows[i], i);
    const { blob } = await buildFormMGJGujaratWorkbookWithTemplateStyles({
      ...workbookArgs,
      headerFormData: mergedHeaderData,
    });
    const xlsxBytes = await blob.arrayBuffer();
    const baseName = resolveFormMGJGujaratEmployeeDownloadBaseName(exportRows[i], hdrs, i);
    zip.file(allocateUniqueFormMGJGujaratDownloadFileName(baseName, usedNames), xlsxBytes);
    if (i > 0 && i % 25 === 0) {
      await new Promise((resolve) => setTimeout(resolve, 0));
    }
  }
  const zipBase = String(formFileName || parsedFormHeader?.title || 'Form_M_GJ')
    .replace(/\.xlsx?$/i, '')
    .replace(/[^a-zA-Z0-9._-]+/g, '_');
  return {
    blob: await zip.generateAsync({ type: 'blob', compression: 'STORE' }),
    fileName: `${zipBase}_Employees.zip`,
  };
}
