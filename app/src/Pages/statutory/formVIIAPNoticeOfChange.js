import * as XLSX from 'xlsx';
import { labelMatchScore } from './form18APAccidentNotice';
import { isFormXAPRegisterOfFinesContext } from './formXAPRegisterOfFines';

/** AP Shops Form VII — Notice of Change [Rule 5(4)]. */

export function formVIIAPHeaderNorm(txt) {
  return String(txt || '')
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

export function matchesFormVIIHint(blob) {
  const text = String(blob || '').toLowerCase();
  if (/\bform[\s._-]*xvii\b/.test(text)) return false;
  if (/\bform[\s._-]*xviii\b/.test(text)) return false;
  if (/\bform[\s._-]*vii\b/.test(text)) return true;
  if (/form___vii|form_-_vii/i.test(text)) return true;
  return false;
}

export function isFormVIIAPNoticeOfChangeContext(formHeader, rowItem, fileName, sheetText = '') {
  if (isFormXAPRegisterOfFinesContext(formHeader, rowItem, fileName, sheetText)) return false;
  const parts = [
    rowItem?.formName,
    rowItem?.FormName,
    rowItem?.description,
    rowItem?.Description,
    rowItem?.state,
    rowItem?.State,
    rowItem?.formFileName,
    rowItem?.FormFileName,
    fileName,
    formHeader?.title,
    formHeader?.subtitle,
    formHeader?.reference,
    sheetText
  ]
    .filter((x) => x != null && String(x).trim() !== '')
    .join(' ')
    .toLowerCase();
  const hasVII = matchesFormVIIHint(parts);
  const hasAP =
    /andhra[\s._-]*pradesh|a\.?\s*p\.?\s*shops/i.test(parts) ||
    String(rowItem?.state || rowItem?.State || '')
      .toLowerCase()
      .includes('andhra');
  const hasNotice = /notice\s+of\s+change/i.test(parts);
  const hasRule54 = /rule\s+5\s*\(\s*4\s*\)/i.test(parts);
  const hasEstablishmentRegistered = /establishment\s+already\s+registered/i.test(parts);
  if (hasVII && hasAP && hasNotice) return true;
  if (hasVII && hasNotice && (hasRule54 || hasEstablishmentRegistered)) return true;
  if (hasNotice && hasAP && hasVII) return true;
  if (hasVII && hasAP && hasEstablishmentRegistered) return true;
  return false;
}

export function isFormVIIAPHeaderFieldLayoutFormHeader(formHeader) {
  return !!formHeader?.formVIIAPHeaderFieldLayout;
}

export const FORM_VII_AP_TEMPLATE_SPECS = [
  {
    key: 'form_vii_ap_establishment_registered',
    label: 'Name of the Establishment already registered:',
    group: 'details',
    fieldType: 'textarea',
    match: /name\s+of\s+the\s+establishment\s+already\s+registered/i
  },
  {
    key: 'form_vii_ap_employer',
    label: 'Name of the Employer:',
    group: 'details',
    match: /name\s+of\s+the\s+employer/i
  },
  {
    key: 'form_vii_ap_registration_certificate',
    label: 'Registration Certificate Number:',
    group: 'details',
    match: /registration\s+certificate\s+number/i
  },
  {
    key: 'form_vii_ap_address',
    label: 'Address:',
    group: 'details',
    fieldType: 'textarea',
    match: /\baddress\b/i
  },
  {
    key: 'form_vii_ap_notice_day',
    label: 'Dated — Day:',
    group: 'date',
    sectionMatch: /dated\s+the/i
  },
  {
    key: 'form_vii_ap_notice_month',
    label: 'Dated — Month:',
    group: 'date',
    sectionMatch: /dated\s+the/i
  },
  {
    key: 'form_vii_ap_notice_year',
    label: 'Dated — Year (20__):',
    group: 'date',
    sectionMatch: /dated\s+the/i
  },
  {
    key: 'form_vii_ap_inspector_address',
    label: 'To — The Inspector (address):',
    group: 'addressee',
    fieldType: 'textarea',
    match: /^to\s*,?\s*$/i
  },
  {
    key: 'form_vii_ap_change_description',
    label: 'Describe the change:',
    group: 'change',
    fieldType: 'textarea',
    match: /describe\s+the\s+change/i
  },
  {
    key: 'form_vii_ap_challan_number',
    label: 'Challan No.:',
    group: 'enclosure',
    sectionMatch: /challan\s+no/i
  },
  {
    key: 'form_vii_ap_challan_date',
    label: 'Challan Dated:',
    group: 'enclosure',
    sectionMatch: /challan\s+no.*dated/i
  },
  {
    key: 'form_vii_ap_challan_amount',
    label: 'Challan Amount (Rs.):',
    group: 'enclosure',
    sectionMatch: /for\s+rs\.?/i
  }
];

export const FORM_VII_AP_FIELD_GROUPS = [
  { id: 'details', title: 'Establishment details' },
  { id: 'date', title: 'Notice date' },
  { id: 'addressee', title: 'To — The Inspector' },
  { id: 'change', title: 'Change particulars' },
  { id: 'enclosure', title: 'Registration Certificate & Challan' }
];

const isNarrativeBlob = (raw) => {
  const n = formVIIAPHeaderNorm(raw);
  if (!n) return false;
  return (
    /notice\s+is\s+hereby\s+given/.test(n) ||
    /following\s+change\s+has\s+taken\s+place/.test(n) ||
    /registration\s+certificate\s+and\s+challan/i.test(n) ||
    /signature\s+of\s+employer/i.test(n) ||
    /^note\s*:/i.test(n) ||
    /^\s*\d+[\.\)]\s+name\s+of\s+the\s+establishment/.test(n) ||
    n.length > 140
  );
};

const buildTemplateField = (spec, coords = {}) => ({
  label: spec.label,
  value: coords.value || '',
  key: spec.key,
  group: spec.group || 'details',
  fieldType: spec.fieldType || 'text',
  labelRow: coords.labelRow ?? null,
  labelCol: coords.labelCol ?? null,
  valueCol: coords.valueCol ?? null,
  valueRow: coords.valueRow ?? null,
  sectionRow: coords.sectionRow ?? null
});

const readValueBelowOrBesideLabel = (getMergedAwareCellText, labelRow, labelCol, effectiveSheetCols) => {
  const maxC = Math.max(20, effectiveSheetCols || 0);
  for (let c = labelCol + 1; c < Math.min(labelCol + 12, maxC); c += 1) {
    const v = String(getMergedAwareCellText(labelRow, c) || '').trim();
    if (v && !isNarrativeBlob(v)) {
      return { value: v, valueCol: c, valueRow: labelRow };
    }
  }
  for (let r = labelRow + 1; r <= Math.min(labelRow + 8, labelRow + 15); r += 1) {
    const v = String(getMergedAwareCellText(r, labelCol) || '').trim();
    if (v && !isNarrativeBlob(v)) {
      return { value: v, valueCol: labelCol, valueRow: r };
    }
  }
  return { value: '', valueCol: labelCol + 1, valueRow: labelRow };
};

const readMultilineValueBelowLabel = (getMergedAwareCellText, labelRow, labelCol, effectiveSheetCols, maxLines = 6) => {
  const maxC = Math.max(20, effectiveSheetCols || 0);
  const lines = [];
  for (let r = labelRow + 1; r <= labelRow + maxLines; r += 1) {
    let line = '';
    for (let c = labelCol; c < Math.min(labelCol + 5, maxC); c += 1) {
      const v = String(getMergedAwareCellText(r, c) || '').trim();
      if (v) line = line ? `${line} ${v}` : v;
    }
    if (!line) break;
    if (isNarrativeBlob(line)) break;
    const norm = formVIIAPHeaderNorm(line);
    if (/^name\s+of\s+the|^registration\s+certificate|^address\s*$|^dated\s+the|^to\s*,?$/i.test(norm)) break;
    lines.push(line);
  }
  return lines.join('\n').trim();
};

const findHeaderLabelCell = (getMergedAwareCellText, matchRe, effectiveSheetCols, maxRows = 120) => {
  const maxC = Math.max(20, effectiveSheetCols || 0);
  for (let r = 0; r < maxRows; r += 1) {
    for (let c = 0; c < maxC; c += 1) {
      const raw = String(getMergedAwareCellText(r, c) || '').trim();
      if (!raw || isNarrativeBlob(raw)) continue;
      const norm = formVIIAPHeaderNorm(raw);
      if (!matchRe.test(norm) && !matchRe.test(raw)) continue;
      if (norm.length > 100) continue;
      const read = readValueBelowOrBesideLabel(getMergedAwareCellText, r, c, effectiveSheetCols);
      const value = isNarrativeBlob(read.value) ? '' : read.value;
      return {
        labelRow: r,
        labelCol: c,
        valueCol: read.valueCol,
        valueRow: read.valueRow,
        value
      };
    }
  }
  return null;
};

const findNumberedLabelCell = (getMergedAwareCellText, matchRe, effectiveSheetCols, maxRows = 120) => {
  const maxC = Math.max(20, effectiveSheetCols || 0);
  for (let r = 0; r < maxRows; r += 1) {
    for (let c = 0; c < maxC; c += 1) {
      const raw = String(getMergedAwareCellText(r, c) || '').trim();
      if (!raw || isNarrativeBlob(raw)) continue;
      const norm = formVIIAPHeaderNorm(raw);
      if (!/^\s*\d+[\.\)]/.test(raw) && !matchRe.test(norm)) continue;
      if (!matchRe.test(norm) && !matchRe.test(raw)) continue;
      const read = readValueBelowOrBesideLabel(getMergedAwareCellText, r, c, effectiveSheetCols);
      return {
        labelRow: r,
        labelCol: c,
        valueCol: read.valueCol,
        valueRow: read.valueRow,
        value: isNarrativeBlob(read.value) ? '' : read.value
      };
    }
  }
  return findHeaderLabelCell(getMergedAwareCellText, matchRe, effectiveSheetCols, maxRows);
};

const findSectionRow = (getMergedAwareCellText, sectionMatch, effectiveSheetCols, maxRows = 120) => {
  const maxC = Math.max(20, effectiveSheetCols || 0);
  for (let r = 0; r < maxRows; r += 1) {
    let rowText = '';
    for (let c = 0; c < maxC; c += 1) {
      const t = String(getMergedAwareCellText(r, c) || '').trim();
      if (t) rowText += `${t} `;
    }
    const joined = rowText.replace(/\s+/g, ' ').trim();
    if (!joined) continue;
    if (joined && sectionMatch.test(joined)) {
      return { sectionRow: r, labelRow: r, labelCol: 0, value: '' };
    }
  }
  return null;
};

function pickWorkbookSheet(workbook, hints = {}) {
  const names = Array.isArray(workbook?.SheetNames) ? workbook.SheetNames : [];
  if (!names.length) return null;
  const preferred = String(hints.preferredSheetName || '').trim();
  if (preferred && workbook.Sheets?.[preferred]) return preferred;
  const blob = [
    hints.fileName,
    hints.formFileName,
    hints.item?.formName,
    hints.item?.FormName
  ]
    .filter(Boolean)
    .join(' ')
    .toLowerCase();
  let best = null;
  let bestScore = -Infinity;
  for (const name of names) {
    const sheetLower = String(name || '').toLowerCase();
    let score = 0;
    if (matchesFormVIIHint(sheetLower) || /notice\s+of\s+change/i.test(sheetLower)) score += 120;
    if (/appointment|register\s+of|employee\s+identification|name\s+of\s+the\s+employee/i.test(sheetLower)) {
      score -= 100;
    }
    if (score > bestScore) {
      bestScore = score;
      best = name;
    }
  }
  if (bestScore > 0 && best) return best;
  const noticeSheet = names.find((n) => /notice\s+of\s+change|form\s*vii/i.test(String(n || '')));
  if (noticeSheet) return noticeSheet;
  const viiSheet = names.find((n) => matchesFormVIIHint(String(n || '')));
  if (viiSheet) return viiSheet;
  if (matchesFormVIIHint(blob)) {
    return names.find((n) => matchesFormVIIHint(String(n || ''))) || names[0];
  }
  return names[0];
}

function buildWorkbookMergedCellAccessor(workbook, hints = {}) {
  if (!workbook?.SheetNames?.length) return null;
  const sheetName = pickWorkbookSheet(workbook, hints);
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
        const topLeft = getRawCellText(m.s.r, m.s.c);
        if (topLeft) return topLeft;
      }
    }
    return '';
  };

  return { getMergedAwareCellText, effectiveSheetCols, sheetName };
}

export function buildFormVIIAPTemplateFields(getMergedAwareCellText = null, effectiveSheetCols = 20) {
  return FORM_VII_AP_TEMPLATE_SPECS.map((spec) => {
    let coords = {};
    if (typeof getMergedAwareCellText === 'function') {
      if (spec.match) {
        coords =
          findNumberedLabelCell(getMergedAwareCellText, spec.match, effectiveSheetCols) ||
          findHeaderLabelCell(getMergedAwareCellText, spec.match, effectiveSheetCols) ||
          {};
      } else if (spec.sectionMatch) {
        coords = findSectionRow(getMergedAwareCellText, spec.sectionMatch, effectiveSheetCols) || {};
      }
    }
    return buildTemplateField(spec, coords);
  });
}

export function resolveFormVIIAPHeaderFieldLayout(parsed, workbook, hints = {}) {
  const formHeader = hints.formHeader || parsed?.formHeader || null;
  const item = hints.item || null;
  const fileName = hints.fileName || hints.formFileName || '';
  const sheetText = hints.sheetText || '';
  if (!isFormVIIAPNoticeOfChangeContext(formHeader, item, fileName, sheetText)) return null;

  const accessor = buildWorkbookMergedCellAccessor(workbook, hints);
  const getMergedAwareCellText = accessor?.getMergedAwareCellText ?? null;
  const effectiveSheetCols = accessor?.effectiveSheetCols ?? 20;
  const finalFields = buildFormVIIAPTemplateFields(getMergedAwareCellText, effectiveSheetCols);
  const title =
    formHeader?.title ||
    (String(sheetText).toLowerCase().includes('notice of change')
      ? 'Form VII – Notice of Change'
      : 'Form VII');

  return {
    formHeader: {
      title,
      subtitle: formHeader?.subtitle || '',
      reference: formHeader?.reference || '',
      formVIIAPHeaderFieldLayout: true,
      textRows: [],
      fields: finalFields
    },
    headers: [],
    tableData: [],
    headerRowIndex: -1,
    dataStartIndex: 0,
    tableStartCol: 0
  };
}

export function resolveFormVIIWorkbookSheetName(workbook, hints = {}) {
  const names = Array.isArray(workbook?.SheetNames) ? workbook.SheetNames : [];
  if (names.length <= 1) return names[0] || null;
  const blob = [
    hints.fileName,
    hints.formFileName,
    hints.formName,
    hints.item?.formName,
    hints.item?.FormName,
    hints.formHeaderTitle,
    hints.formHeader?.title
  ]
    .filter(Boolean)
    .join(' ')
    .toLowerCase();
  const buildSheetTextBlob = (sheetName) => {
    const ws = workbook.Sheets?.[sheetName];
    if (!ws) return '';
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    return rows
      .slice(0, 40)
      .map((row) => (Array.isArray(row) ? row : []).map((c) => String(c || '')).join(' '))
      .join(' ')
      .toLowerCase();
  };
  let best = null;
  let bestScore = -Infinity;
  for (const name of names) {
    const sheetLower = String(name || '').toLowerCase();
    const sheetText = buildSheetTextBlob(name);
    const sheetBlob = `${sheetLower} ${sheetText}`;
    let score = 0;
    if (matchesFormVIIHint(sheetLower) || /notice\s+of\s+change/i.test(sheetLower)) score += 120;
    if (/notice\s+of\s+change/i.test(sheetBlob)) score += 95;
    if (/rule\s+5\s*\(\s*4\s*\)/i.test(sheetText)) score += 45;
    if (/establishment\s+already\s+registered/i.test(sheetText)) score += 40;
    if (/registration\s+certificate\s+number/i.test(sheetText)) score += 30;
    if (/describe\s+the\s+change/i.test(sheetText)) score += 25;
    if (/appointment|register\s+of\s+fines|register\s+of\s+advances/i.test(sheetBlob)) score -= 120;
    if (/name\s+of\s+the\s+employee|employee\s+identification/i.test(sheetBlob)) score -= 150;
    if (score > bestScore) {
      bestScore = score;
      best = name;
    }
  }
  return bestScore > 0 ? best : pickWorkbookSheet(workbook, hints);
}

export function repickFormVIIWorkbookSheetIfNeeded(workbook, hints, currentSheetName, parsedFormHeader) {
  const names = Array.isArray(workbook?.SheetNames) ? workbook.SheetNames : [];
  if (names.length <= 1) return null;
  const wantsFormVII =
    matchesFormVIIHint(
      [hints.fileName, hints.formFileName, hints.item?.formName, hints.item?.FormName]
        .filter(Boolean)
        .join(' ')
    ) ||
    isFormVIIAPNoticeOfChangeContext(
      parsedFormHeader,
      hints.item,
      hints.fileName || hints.formFileName
    );
  if (!wantsFormVII) return null;
  const currentBlob = `${parsedFormHeader?.title || ''} ${currentSheetName || ''}`.toLowerCase();
  if (
    matchesFormVIIHint(currentBlob) &&
    /notice\s+of\s+change|establishment\s+already\s+registered/i.test(currentBlob)
  ) {
    return null;
  }
  return resolveFormVIIWorkbookSheetName(workbook, {
    ...hints,
    formHeader: parsedFormHeader
  });
}

export function applyFormVIIAPAutofillFromSite(headerData, siteContext = {}) {
  const out = headerData && typeof headerData === 'object' ? { ...headerData } : {};
  const fill = (key, value) => {
    const text = String(value ?? '').trim();
    if (!text || !key) return;
    const cur = String(out[key] ?? '').trim();
    if (!cur || /^enter\b/i.test(cur)) out[key] = text;
  };
  fill('form_vii_ap_establishment_registered', siteContext.establishmentText || '');
  fill('form_vii_ap_employer', siteContext.employerText || '');
  fill('form_vii_ap_registration_certificate', siteContext.registrationNumber || '');
  fill('form_vii_ap_address', siteContext.addressText || '');
  const today = new Date();
  fill('form_vii_ap_notice_day', String(today.getDate()));
  fill(
    'form_vii_ap_notice_month',
    today.toLocaleString('en-GB', { month: 'long' })
  );
  fill('form_vii_ap_notice_year', String(today.getFullYear()).slice(-2));
  return out;
}

export function importFormVIIAPHeaderFieldsFromWorkbook(workbook, hints = {}) {
  const accessor = buildWorkbookMergedCellAccessor(workbook, hints);
  if (!accessor) return null;
  const { getMergedAwareCellText, effectiveSheetCols } = accessor;
  const fields = buildFormVIIAPTemplateFields(getMergedAwareCellText, effectiveSheetCols);
  const out = {};
  fields.forEach((field) => {
    const spec = FORM_VII_AP_TEMPLATE_SPECS.find((s) => s.key === field.key);
    let val = '';
    if (field.labelRow != null && spec?.fieldType === 'textarea') {
      val = readMultilineValueBelowLabel(
        getMergedAwareCellText,
        field.labelRow,
        field.labelCol ?? 0,
        effectiveSheetCols
      );
    } else if (field.labelRow != null && field.valueCol != null) {
      const raw = String(
        getMergedAwareCellText(field.valueRow ?? field.labelRow, field.valueCol) || field.value || ''
      ).trim();
      val = isNarrativeBlob(raw) ? '' : raw;
    } else if (field.value) {
      val = String(field.value).trim();
    }
    if (val) out[field.key] = val;
  });
  return out;
}

export function writeFormVIIAPFieldsToExcelJsWorksheet(worksheet, headerFormData, parsedFormHeader, helpers = {}) {
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
  const normalize = (txt) => formVIIAPHeaderNorm(txt);
  const fields = Array.isArray(parsedFormHeader?.fields) ? parsedFormHeader.fields : [];
  const maxScanRows = Math.max(200, worksheet.rowCount + 10);
  const maxScanCols = 20;

  const writeAt = (row, col, value) => {
    const text = String(value ?? '').trim();
    if (!text || row == null || col == null) return;
    worksheet.getCell(row, col).value = text;
  };

  fields.forEach((field) => {
    const val = headerFormData[field.key];
    if (val == null || String(val).trim() === '') return;

    if (field.labelRow != null && field.valueCol != null) {
      const targetRow = (field.valueRow ?? field.labelRow) + 1;
      writeAt(targetRow, field.valueCol + 1, val);
      return;
    }

    const spec = FORM_VII_AP_TEMPLATE_SPECS.find((s) => s.key === field.key);
    const labelNorm = normalize(String(field.label || spec?.label || '').replace(/:+$/, ''));
    if (!labelNorm) return;
    for (let r = 1; r <= maxScanRows; r += 1) {
      for (let c = 1; c <= maxScanCols; c += 1) {
        const raw = excelCellValueToString(worksheet.getCell(r, c)?.value);
        const score = labelMatchScore(labelNorm, normalize(raw.replace(/:+$/, '')));
        if (score < 45) continue;
        for (let nc = c + 1; nc <= Math.min(c + 10, maxScanCols + 4); nc += 1) {
          const nt = normalize(excelCellValueToString(worksheet.getCell(r, nc)?.value));
          if (!nt || nt === ':') {
            writeAt(r, nc, val);
            return;
          }
        }
        writeAt(r + 1, c, val);
        return;
      }
    }
  });
}
