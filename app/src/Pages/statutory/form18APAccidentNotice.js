import * as XLSX from 'xlsx';

/** Form No. 18 (AP) — Notice of Accident [Rule 96 / ESI Regulation 68]. */

export const FORM_18_AP_LABEL_COL = 2;
export const FORM_18_AP_VALUE_COL = 3;
export const FORM_18_AP_NUM_COL = 1;
export const FORM_18_AP_ESI_LABEL_COL = 4;

export function form18APHeaderNorm(txt) {
  return String(txt || '')
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

export function isForm18APAccidentNoticeContext(formHeader, rowItem, fileName, sheetText = '') {
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
  if (/form[\s._-]*18[\s._-]*[-_.]?\s*ap|form18[\s._-]*andhra|form_18[\s._-]*andhra/i.test(parts)) {
    return true;
  }
  if (/form18[\s._-]*andhra[\s._-]*pradesh/i.test(parts)) {
    return true;
  }
  if (/\bform[\s._-]*18\b/.test(parts) && /andhra[\s._-]*pradesh|\bap\b/.test(parts)) {
    if (/notice\s+of\s+accident|dangerous\s+occurrence|rule\s+96|regulation\s+68|employees?\s+state\s+insurance/i.test(parts)) {
      return true;
    }
  }
  if (/\bform\s+no\.?\s*18\b/.test(parts)) {
    if (/notice\s+of\s+accident|dangerous\s+occurrence|employees?\s+state\s+insurance|\besi\b/i.test(parts)) {
      return true;
    }
  }
  return false;
}

export function isForm18APHeaderFieldLayoutFormHeader(formHeader) {
  return !!formHeader?.form18APColumnBoxLayout;
}

/** Fallback labels only when the sheet cannot be scanned — order matches typical AP template. */
const FORM_18_AP_FALLBACK_LABELS = [
  'Name of Occupier (Factory / Employer)',
  "ESI Employer's Code No.",
  'Address of works/premises where accident or dangerous occurrence took place',
  'Nature of Industry',
  'Branch or department and exact place where the accident/dangerous occurrence took place',
  "Employee's State Insurance Number (if covered)",
  'Name and address of the injured person',
  '(a) Sex',
  '(b) Age (last birthday)',
  '(c) Occupation of the injured persons',
  '(d) Monthly wages of the person injured',
  'Local E.S.I. office to which the injured person attached',
  'Date, shift and hour of accident/dangerous occurrence',
  'If causes is by machinery',
  'State whether it was moved by mechanical power at that time',
  'its nature',
  'Nature of injury',
  'Location of injury (right leg, left hand or left eye etc.)',
  'Date and hour of return to work',
  '(a) Physician, dispensary or hospital from whom or in which the injured person received or is receiving treatment',
  'Has the injured person died',
  'If so, date of death',
  'Other particulars (e.g. fatal leg injury, arm injury, etc.)',
  'Result of investigation',
  'Number of accident/dangerous occurrence',
  'Date of investigation',
  'Date of receipt',
  'Signature Name & Designation of the occupier or Manager/Employer Employer\'s address and E.S.I Code No. Address ___________',
  'I certify that to the best of my knowledge and belief the above particulars are correct in every respect.'
];

/** Fields omitted from the autofill modal (not on this AP template). */
const FORM_18_AP_EXCLUDED_LABELS = [
  'Cause of accident',
  'If dangerous occurrence, nature thereof',
  'Names and addresses of witnesses',
  'Has the injured person been disabled? If so, give extent of disablement',
  'Has the injured person returned to work? If so, give date of return'
];

const isForm18APExcludedFieldLabel = (label) => {
  const norm = form18APHeaderNorm(label);
  if (!norm) return false;
  return FORM_18_AP_EXCLUDED_LABELS.some((excluded) => {
    const exNorm = form18APHeaderNorm(excluded);
    return exNorm === norm || norm.includes(exNorm) || exNorm.includes(norm);
  });
};

export const labelMatchScore = (canonicalNorm, candidateNorm) => {
  if (!canonicalNorm || !candidateNorm) return 0;
  if (canonicalNorm === candidateNorm) return 100;
  if (candidateNorm.includes(canonicalNorm) || canonicalNorm.includes(candidateNorm)) {
    return 80 + Math.min(canonicalNorm.length, candidateNorm.length) / 10;
  }
  const canonTokens = canonicalNorm.split(' ').filter((t) => t.length > 2);
  const candTokens = new Set(candidateNorm.split(' ').filter((t) => t.length > 2));
  if (canonTokens.length === 0) return 0;
  let hits = 0;
  canonTokens.forEach((t) => {
    if (candTokens.has(t)) hits += 1;
  });
  return (hits / canonTokens.length) * 70;
};

const filterForm18APFields = (fields) =>
  (Array.isArray(fields) ? fields : []).filter((f) => !isForm18APExcludedFieldLabel(f?.label));

const looksLikeForm18APFieldLabel = (raw) => {
  const t = String(raw || '').trim();
  if (!t || t.length < 4) return false;
  if (isForm18APExcludedFieldLabel(t)) return false;
  const n = form18APHeaderNorm(t);
  if (/^form\s+no\.?\s*18\b/.test(n)) return false;
  if (/prescribed under rule|regulation\s+68/i.test(n) && /notice of accident/i.test(n) && t.length > 72) return false;
  if (/^\([a-d]\)\s+/i.test(t)) return true;
  if (/^its nature$/i.test(n)) return true;
  if (/^(if |state whether|location of|result of|other particulars|i certify|has the|number of|date of|date and)/i.test(t)) {
    return true;
  }
  if (/^(name of|address|nature|branch|employee|local|date|sex|age|occupation|monthly|esi|cause|witness|extent|disabled|returned|death|signature)/i.test(t)) {
    return true;
  }
  return /accident|dangerous occurrence|injured|insurance|employer|occupier|premises|industry|department|shift|wages|witness|machinery|mechanical power|physician|dispensary|investigation|particulars|certify|return to work|receipt|designation|manager/i.test(n);
};

const looksLikeAnotherFieldLabel = (raw) => {
  const t = String(raw || '').trim();
  if (!t || t.length < 8) return false;
  return looksLikeForm18APFieldLabel(t) && !/^\d{1,3}$/.test(t);
};

const readValueBesideLabel = (getMergedAwareCellText, row, labelCol, effectiveSheetCols) => {
  const maxC = Math.max(12, effectiveSheetCols || 0);
  for (let c = labelCol + 1; c < Math.min(labelCol + 6, maxC); c += 1) {
    const v = String(getMergedAwareCellText(row, c) || '').trim();
    if (!v) continue;
    if (looksLikeAnotherFieldLabel(v)) break;
    return { value: v, valueCol: c };
  }
  return { value: '', valueCol: labelCol + 1 };
};

const makeFieldKey = (label) => {
  const slug = form18APHeaderNorm(label)
    .replace(/[^a-z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '')
    .slice(0, 56);
  return `form18_ap_${slug || 'field'}`;
};

const buildField = (label, coords = {}) => ({
  label: String(label || '').trim(),
  value: coords.value || '',
  key: makeFieldKey(label),
  labelRow: coords.labelRow ?? null,
  labelCol: coords.labelCol ?? FORM_18_AP_LABEL_COL,
  valueCol: coords.valueCol ?? FORM_18_AP_VALUE_COL
});

const fieldAlreadyPresent = (fields, label) => {
  const norm = form18APHeaderNorm(label);
  return fields.some((f) => labelMatchScore(norm, form18APHeaderNorm(f.label)) >= 45);
};

const sortFieldsBySheetPosition = (fields) =>
  [...fields].sort((a, b) => {
    const rowA = a.labelRow ?? 9999;
    const rowB = b.labelRow ?? 9999;
    if (rowA !== rowB) return rowA - rowB;
    return (a.labelCol ?? 0) - (b.labelCol ?? 0);
  });

/**
 * Keep Excel top-to-bottom order; insert missed labels at their sheet row when found.
 */
export function finalizeForm18APFieldOrder(sheetFields = [], getMergedAwareCellText = null) {
  const ordered = sortFieldsBySheetPosition(
    Array.isArray(sheetFields) ? sheetFields.filter((f) => String(f?.label || '').trim()) : []
  );

  const insertFieldInRowOrder = (field) => {
    const row = field.labelRow ?? 9999;
    const col = field.labelCol ?? 0;
    let insertAt = ordered.length;
    for (let i = 0; i < ordered.length; i += 1) {
      const existingRow = ordered[i].labelRow ?? 9999;
      const existingCol = ordered[i].labelCol ?? 0;
      if (existingRow > row || (existingRow === row && existingCol > col)) {
        insertAt = i;
        break;
      }
    }
    ordered.splice(insertAt, 0, field);
  };

  FORM_18_AP_FALLBACK_LABELS.forEach((label) => {
    if (fieldAlreadyPresent(ordered, label) || isForm18APExcludedFieldLabel(label)) return;

    let foundRow = null;
    let foundCol = FORM_18_AP_LABEL_COL;
    let foundValue = '';
    let foundValueCol = FORM_18_AP_VALUE_COL;

    if (typeof getMergedAwareCellText === 'function') {
      for (let r = 1; r < 220; r += 1) {
        for (const c of [FORM_18_AP_LABEL_COL, FORM_18_AP_ESI_LABEL_COL]) {
          const raw = String(getMergedAwareCellText(r, c) || '').trim();
          if (!raw) continue;
          if (labelMatchScore(form18APHeaderNorm(label), form18APHeaderNorm(raw)) >= 45) {
            foundRow = r;
            foundCol = c;
            const read = readValueBesideLabel(getMergedAwareCellText, r, c, 20);
            foundValue = read.value;
            foundValueCol = read.valueCol;
            break;
          }
        }
        if (foundRow != null) break;
      }
    }

    const field = buildField(label, {
      value: foundValue,
      labelRow: foundRow,
      labelCol: foundCol,
      valueCol: foundValueCol
    });

    if (foundRow != null) {
      insertFieldInRowOrder(field);
    } else {
      ordered.push(field);
    }
  });

  return filterForm18APFields(ordered);
};

/** Parse fields in Excel row order: column C labels, then column E (ESI) on the same row. */
export function enrichForm18APFieldsFromSheet(getMergedAwareCellText, effectiveSheetCols, maxRows = 200) {
  const fields = [];
  const seenLabels = new Set();
  const scanRows = Math.min(Math.max(maxRows, 40), 220);

  const pushField = (label, row, labelCol, value, valueCol) => {
    if (isForm18APExcludedFieldLabel(label)) return;
    const norm = form18APHeaderNorm(label);
    if (!norm || seenLabels.has(norm)) return;
    if (fields.some((f) => labelMatchScore(norm, form18APHeaderNorm(f.label)) >= 45)) return;
    seenLabels.add(norm);
    fields.push(
      buildField(label, {
        value: value || '',
        labelRow: row,
        labelCol,
        valueCol
      })
    );
  };

  for (let r = 1; r < scanRows; r += 1) {
    const primaryLabel = String(getMergedAwareCellText(r, FORM_18_AP_LABEL_COL) || '').trim();
    if (primaryLabel && looksLikeForm18APFieldLabel(primaryLabel)) {
      const { value, valueCol } = readValueBesideLabel(
        getMergedAwareCellText,
        r,
        FORM_18_AP_LABEL_COL,
        effectiveSheetCols
      );
      pushField(primaryLabel, r, FORM_18_AP_LABEL_COL, value, valueCol);
    }

    const esiLabel = String(getMergedAwareCellText(r, FORM_18_AP_ESI_LABEL_COL) || '').trim();
    if (esiLabel && /esi\s+employer'?s?\s+code/i.test(esiLabel)) {
      const esiRead = readValueBesideLabel(getMergedAwareCellText, r, FORM_18_AP_ESI_LABEL_COL, effectiveSheetCols);
      pushField(esiLabel, r, FORM_18_AP_ESI_LABEL_COL, esiRead.value, esiRead.valueCol);
    }
  }

  return filterForm18APFields(fields);
}

export function buildForm18APFallbackFields() {
  return filterForm18APFields(FORM_18_AP_FALLBACK_LABELS.map((label) => buildField(label)));
}

export function resolveForm18APHeaderFieldLayout(parsed, workbook, hints = {}) {
  const formHeader = hints.formHeader || parsed?.formHeader || null;
  const item = hints.item || null;
  const fileName = hints.fileName || hints.formFileName || '';
  const sheetText = hints.sheetText || '';
  if (!isForm18APAccidentNoticeContext(formHeader, item, fileName, sheetText)) return null;

  let sheetFields = [];

  if (workbook?.SheetNames?.[0]) {
    const sheetName =
      workbook.SheetNames.find((n) => /^18$/i.test(String(n || '').trim())) ||
      workbook.SheetNames.find((n) => /18/i.test(String(n || ''))) ||
      workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    if (worksheet) {
      const jsonData = worksheet['!ref']
        ? XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' })
        : [];
      const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1:A1');
      const maxSheetCols = range.e && typeof range.e.c === 'number' ? range.e.c + 1 : 0;
      const merges = worksheet['!merges'] || [];
      let mergeMaxCol = 0;
      merges.forEach((m) => {
        if (m?.e && typeof m.e.c === 'number') mergeMaxCol = Math.max(mergeMaxCol, m.e.c + 1);
      });
      const effectiveSheetCols = Math.max(maxSheetCols, mergeMaxCol);
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
            const topLeft = getRawCellText(m.s.r, m.s.c);
            if (topLeft) return topLeft;
          }
        }
        return '';
      };
      sheetFields = enrichForm18APFieldsFromSheet(
        getMergedAwareCellText,
        effectiveSheetCols,
        jsonData.length
      );
      const fields =
        sheetFields.length >= 3
          ? finalizeForm18APFieldOrder(sheetFields, getMergedAwareCellText)
          : buildForm18APFallbackFields();

      return {
        formHeader: {
          ...(formHeader || {}),
          form18APColumnBoxLayout: true,
          fields
        },
        headers: [],
        tableData: [],
        headerRowIndex: -1,
        dataStartIndex: 0,
        tableStartCol: 0
      };
    }
  }

  const fields = sheetFields.length >= 3 ? finalizeForm18APFieldOrder(sheetFields) : buildForm18APFallbackFields();

  return {
    formHeader: {
      ...(formHeader || {}),
      form18APColumnBoxLayout: true,
      fields
    },
    headers: [],
    tableData: [],
    headerRowIndex: -1,
    dataStartIndex: 0,
    tableStartCol: 0
  };
}

export function applyForm18APHeaderFieldLayoutToParsed(parsed, workbook, hints = {}) {
  if (!parsed) return parsed;
  const layout = resolveForm18APHeaderFieldLayout(parsed, workbook, {
    ...hints,
    formHeader: hints.formHeader || parsed.formHeader
  });
  if (!layout) return parsed;
  return {
    ...parsed,
    formHeader: layout.formHeader,
    headers: layout.headers,
    tableData: layout.tableData,
    headerRowIndex: layout.headerRowIndex,
    dataStartIndex: layout.dataStartIndex,
    tableStartCol: layout.tableStartCol
  };
}

export function writeForm18APFieldsToExcelJsWorksheet(worksheet, headerFormData, parsedFormHeader, helpers = {}) {
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
  const normalize = (txt) => form18APHeaderNorm(txt);
  const fields = Array.isArray(parsedFormHeader?.fields) ? parsedFormHeader.fields : [];
  const maxScanRows = Math.max(220, worksheet.rowCount + 10);
  const maxScanCols = 12;

  const writeAt = (row, col, value) => {
    const text = String(value ?? '').trim();
    if (!text || row == null || col == null) return;
    worksheet.getCell(row, col).value = text;
  };

  fields.forEach((field) => {
    const val = headerFormData[field.key];
    if (val == null || String(val).trim() === '') return;
    if (field.labelRow != null && field.valueCol != null) {
      writeAt(field.labelRow + 1, field.valueCol + 1, val);
      return;
    }
    const labelNorm = normalize(String(field.label || '').replace(/:+$/, ''));
    if (!labelNorm) return;
    let bestRow = -1;
    let bestCol = -1;
    let bestScore = 0;
    for (let r = 1; r <= maxScanRows; r += 1) {
      for (let c = 1; c <= maxScanCols; c += 1) {
        const raw = excelCellValueToString(worksheet.getCell(r, c)?.value);
        const score = labelMatchScore(labelNorm, normalize(raw.replace(/:+$/, '')));
        if (score > bestScore) {
          bestScore = score;
          bestRow = r;
          bestCol = c;
        }
      }
    }
    if (bestRow > 0 && bestCol > 0 && bestScore >= 45) {
      for (let nc = bestCol + 1; nc <= Math.min(bestCol + 6, maxScanCols + 4); nc += 1) {
        const nt = normalize(excelCellValueToString(worksheet.getCell(bestRow, nc)?.value));
        if (!nt || nt === ':') {
          writeAt(bestRow, nc, val);
          return;
        }
        if (looksLikeAnotherFieldLabel(excelCellValueToString(worksheet.getCell(bestRow, nc)?.value))) break;
      }
      writeAt(bestRow, Math.min(bestCol + 2, maxScanCols + 2), val);
    }
  });
}
