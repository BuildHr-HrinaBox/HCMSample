import * as XLSX from 'xlsx';

/** Form No. 2-A (AP) — Notice of Change of Manager / Occupier [Rule 12]. */

export const FORM_2AP_NUM_COL = 0;
export const FORM_2AP_LABEL_COL = 1;
export const FORM_2AP_VALUE_COL = 2;

export function form2APHeaderNorm(txt) {
  return String(txt || '')
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

export function isForm2APChangeNoticeContext(formHeader, rowItem, fileName, sheetText = '') {
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
  if (/form[\s._-]*2[\s._-]*[-_.]?\s*a|form_2[\s._-]*a|form2[\s._-]*a/i.test(parts)) {
    if (/andhra[\s._-]*pradesh|\bap\b/.test(parts)) return true;
    if (/change\s+of\s+manager|manager\s*\/\s*occupier|rule\s+12/.test(parts)) return true;
  }
  if (/\bform\s+no\.?\s*2[\s-]*a\b|\bform\s+2[\s-]*a\b/.test(parts)) {
    if (/notice\s+of\s+change|manager|occupier|rule\s+12|andhra|\bap\b/.test(parts)) return true;
  }
  if (/notice\s+of\s+change\s+of\s+manager/.test(parts) && /andhra[\s._-]*pradesh|\bap\b/.test(parts)) {
    return true;
  }
  return false;
}

export function isForm2APHeaderFieldLayoutFormHeader(formHeader) {
  return !!formHeader?.form2APColumnBoxLayout;
}

const FORM_2AP_FALLBACK_LABELS = [
  'Full name and Address of the Factory',
  'Licence Number',
  'Full postal Address for communications relating to the factory',
  'Whether change relates to Manager or Occupier',
  'Name of Outgoing Manager / Occupier',
  'Name and residential address of Incoming Manager / Occupier',
  'Date from which the change is to take effect',
  'Date of notice',
  'Signature of Occupier / Manager'
];

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

const looksLikeForm2APTitleOrInstruction = (raw) => {
  const n = form2APHeaderNorm(raw);
  if (!n) return true;
  if (/^form\s+no\.?\s*2[\s-]*a\b/.test(n)) return true;
  if (/prescribed under rule\s+12/.test(n) && /notice of change/.test(n)) return true;
  if (/^notice of change of manager/.test(n) && n.length > 48) return true;
  if (/this form is to be sent to the inspector/.test(n)) return true;
  if (/^page\s+\d/.test(n)) return true;
  return false;
};

const looksLikeForm2APFieldLabel = (raw) => {
  const t = String(raw || '').trim();
  if (!t || t.length < 4) return false;
  if (looksLikeForm2APTitleOrInstruction(t)) return false;
  if (/^\d{1,2}$/.test(t)) return false;
  const n = form2APHeaderNorm(t);
  return (
    /factory|licen[cs]e|address|manager|occupier|change|incoming|outgoing|postal|communication|signature|date|name|effect|notice/i.test(
      n
    ) || t.length >= 12
  );
};

const looksLikeAnotherFieldLabel = (raw) => {
  const t = String(raw || '').trim();
  if (!t || t.length < 8) return false;
  return looksLikeForm2APFieldLabel(t) && !/^\d{1,3}$/.test(t);
};

const readValueBesideLabel = (getMergedAwareCellText, row, labelCol, effectiveSheetCols) => {
  const maxC = Math.max(8, effectiveSheetCols || 0);
  for (let c = labelCol + 1; c < Math.min(labelCol + 4, maxC); c += 1) {
    const v = String(getMergedAwareCellText(row, c) || '').trim();
    if (!v) continue;
    if (looksLikeAnotherFieldLabel(v)) break;
    return { value: v, valueCol: c };
  }
  return { value: '', valueCol: labelCol + 1 };
};

const makeFieldKey = (label) => {
  const slug = form2APHeaderNorm(label)
    .replace(/[^a-z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '')
    .slice(0, 56);
  return `form2_ap_${slug || 'field'}`;
};

const buildField = (label, coords = {}) => ({
  label: String(label || '').trim(),
  value: coords.value || '',
  key: makeFieldKey(label),
  labelRow: coords.labelRow ?? null,
  labelCol: coords.labelCol ?? FORM_2AP_LABEL_COL,
  valueCol: coords.valueCol ?? FORM_2AP_VALUE_COL
});

const fieldAlreadyPresent = (fields, label) => {
  const norm = form2APHeaderNorm(label);
  return fields.some((f) => labelMatchScore(norm, form2APHeaderNorm(f.label)) >= 45);
};

const sortFieldsBySheetPosition = (fields) =>
  [...fields].sort((a, b) => {
    const rowA = a.labelRow ?? 9999;
    const rowB = b.labelRow ?? 9999;
    if (rowA !== rowB) return rowA - rowB;
    return (a.labelCol ?? 0) - (b.labelCol ?? 0);
  });

export function finalizeForm2APFieldOrder(sheetFields = [], getMergedAwareCellText = null) {
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

  FORM_2AP_FALLBACK_LABELS.forEach((label) => {
    if (fieldAlreadyPresent(ordered, label)) return;

    let foundRow = null;
    let foundValue = '';
    let foundValueCol = FORM_2AP_VALUE_COL;

    if (typeof getMergedAwareCellText === 'function') {
      for (let r = 0; r < 120; r += 1) {
        const raw = String(getMergedAwareCellText(r, FORM_2AP_LABEL_COL) || '').trim();
        if (!raw) continue;
        if (labelMatchScore(form2APHeaderNorm(label), form2APHeaderNorm(raw)) >= 45) {
          foundRow = r;
          const read = readValueBesideLabel(getMergedAwareCellText, r, FORM_2AP_LABEL_COL, 8);
          foundValue = read.value;
          foundValueCol = read.valueCol;
          break;
        }
      }
    }

    const field = buildField(label, {
      value: foundValue,
      labelRow: foundRow,
      valueCol: foundValueCol
    });

    if (foundRow != null) {
      insertFieldInRowOrder(field);
    } else {
      ordered.push(field);
    }
  });

  return ordered;
}

/** Parse fields in Excel row order: column B labels, column C values. */
export function enrichForm2APFieldsFromSheet(getMergedAwareCellText, effectiveSheetCols, maxRows = 120) {
  const fields = [];
  const seenLabels = new Set();
  const scanRows = Math.min(Math.max(maxRows, 30), 150);

  const pushField = (label, row, value, valueCol) => {
    const norm = form2APHeaderNorm(label);
    if (!norm || seenLabels.has(norm)) return;
    if (fields.some((f) => labelMatchScore(norm, form2APHeaderNorm(f.label)) >= 45)) return;
    seenLabels.add(norm);
    fields.push(
      buildField(label, {
        value: value || '',
        labelRow: row,
        valueCol
      })
    );
  };

  for (let r = 0; r < scanRows; r += 1) {
    const label = String(getMergedAwareCellText(r, FORM_2AP_LABEL_COL) || '').trim();
    if (!label || !looksLikeForm2APFieldLabel(label)) continue;
    const { value, valueCol } = readValueBesideLabel(
      getMergedAwareCellText,
      r,
      FORM_2AP_LABEL_COL,
      effectiveSheetCols
    );
    pushField(label, r, value, valueCol);
  }

  return fields;
}

export function buildForm2APFallbackFields() {
  return FORM_2AP_FALLBACK_LABELS.map((label) => buildField(label));
}

function resolveForm2APWorksheet(workbook) {
  if (!workbook?.SheetNames?.length) return null;
  const sheetName =
    workbook.SheetNames.find((n) => /^2[\s-]*a$/i.test(String(n || '').trim())) ||
    workbook.SheetNames.find((n) => /2[\s-]*a|change.*manager|manager.*occupier/i.test(String(n || ''))) ||
    workbook.SheetNames[0];
  return workbook.Sheets[sheetName] ? { sheetName, worksheet: workbook.Sheets[sheetName] } : null;
}

export function resolveForm2APWorkbookSheetName(workbook) {
  const resolved = resolveForm2APWorksheet(workbook);
  return resolved?.sheetName || workbook?.SheetNames?.[0] || '';
}

export function repickForm2APWorkbookSheetIfNeeded(workbook, hints, currentSheetName) {
  if (!isForm2APChangeNoticeContext(
    hints?.formHeader,
    hints?.item,
    hints?.fileName || hints?.formFileName || '',
    hints?.sheetText || ''
  )) {
    return null;
  }
  const target = resolveForm2APWorkbookSheetName(workbook);
  if (target && target !== currentSheetName) return target;
  return null;
}

function buildWorksheetHelpers(worksheet) {
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
  return { jsonData, effectiveSheetCols, getMergedAwareCellText };
}

export function resolveForm2APHeaderFieldLayout(parsed, workbook, hints = {}) {
  const formHeader = hints.formHeader || parsed?.formHeader || null;
  const item = hints.item || null;
  const fileName = hints.fileName || hints.formFileName || '';
  const sheetText = hints.sheetText || '';
  if (!isForm2APChangeNoticeContext(formHeader, item, fileName, sheetText)) return null;

  const resolved = resolveForm2APWorksheet(workbook);
  if (resolved?.worksheet) {
    const { jsonData, effectiveSheetCols, getMergedAwareCellText } = buildWorksheetHelpers(
      resolved.worksheet
    );
    const sheetFields = enrichForm2APFieldsFromSheet(
      getMergedAwareCellText,
      effectiveSheetCols,
      jsonData.length
    );
    const fields =
      sheetFields.length >= 3
        ? finalizeForm2APFieldOrder(sheetFields, getMergedAwareCellText)
        : buildForm2APFallbackFields();

    return {
      formHeader: {
        ...(formHeader || {}),
        title:
          formHeader?.title ||
          'Form No. 2-A – Notice of Change of Manager / Occupier',
        subtitle:
          formHeader?.subtitle ||
          'Prescribed under Rule 12 (Andhra Pradesh Factories Rules)',
        form2APColumnBoxLayout: true,
        fields
      },
      headers: [],
      tableData: [],
      headerRowIndex: -1,
      dataStartIndex: 0,
      tableStartCol: 0
    };
  }

  return {
    formHeader: {
      ...(formHeader || {}),
      title:
        formHeader?.title ||
        'Form No. 2-A – Notice of Change of Manager / Occupier',
      subtitle:
        formHeader?.subtitle ||
        'Prescribed under Rule 12 (Andhra Pradesh Factories Rules)',
      form2APColumnBoxLayout: true,
      fields: buildForm2APFallbackFields()
    },
    headers: [],
    tableData: [],
    headerRowIndex: -1,
    dataStartIndex: 0,
    tableStartCol: 0
  };
}

export function applyForm2APHeaderFieldLayoutToParsed(parsed, workbook, hints = {}) {
  if (!parsed) return parsed;
  const layout = resolveForm2APHeaderFieldLayout(parsed, workbook, {
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

export function writeForm2APFieldsToExcelJsWorksheet(worksheet, headerFormData, parsedFormHeader, helpers = {}) {
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
  const normalize = (txt) => form2APHeaderNorm(txt);
  const fields = Array.isArray(parsedFormHeader?.fields) ? parsedFormHeader.fields : [];
  const maxScanRows = Math.max(120, worksheet.rowCount + 10);
  const maxScanCols = 8;

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
      for (let nc = bestCol + 1; nc <= Math.min(bestCol + 4, maxScanCols + 2); nc += 1) {
        const nt = normalize(excelCellValueToString(worksheet.getCell(bestRow, nc)?.value));
        if (!nt || nt === ':') {
          writeAt(bestRow, nc, val);
          return;
        }
        if (looksLikeAnotherFieldLabel(excelCellValueToString(worksheet.getCell(bestRow, nc)?.value))) break;
      }
      writeAt(bestRow, Math.min(bestCol + 2, maxScanCols + 1), val);
    }
  });
}
