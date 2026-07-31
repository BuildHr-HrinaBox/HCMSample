import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import {
  flattenPayrollEarningColumns,
  readForm10GrossPayAmount,
  readPayrollForm15WageAmounts,
  readPayrollScalar,
} from '../../utils/payrollEarnings';
import { isFormXXVITamilNaduClraContext, matchesFormXXVIHint } from './formXXVITamilNaduMuster';
import { labelMatchScore } from './form18APAccidentNotice';
import { isForm2APChangeNoticeContext } from './form2APChangeNotice';

/** AP Shops Form XXVI — Letter of Appointment [Rule 30]. Header fields + employee table. */

export function formXXVIAPHeaderNorm(txt) {
  return String(txt || '')
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

export function isFormXXVIAPAppointmentLetterContext(formHeader, rowItem, fileName, sheetText = '') {
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
  if (isForm2APChangeNoticeContext(formHeader, rowItem, fileName, sheetText)) return false;
  if (isFormXXVITamilNaduClraContext(formHeader, rowItem, fileName, '', sheetText)) return false;
  if (/daily\s+hours|muster\s+roll/i.test(parts) && !/appointment/i.test(parts)) return false;
  const hasAppointment =
    /appointment(?:\s*(?:ltr|letter|order))?/i.test(parts) ||
    /letter\s+of\s+appointment/i.test(parts);
  const hasAP =
    /andhra[\s._-]*pradesh|a\.?\s*p\.?\s*shops/i.test(parts) ||
    String(rowItem?.state || rowItem?.State || '')
      .toLowerCase()
      .includes('andhra');
  const hasXXVI = matchesFormXXVIHint(parts);
  const hasRule30 = /rule\s+30/i.test(parts);
  if (hasXXVI && hasAP && (hasAppointment || hasRule30)) return true;
  if (hasAppointment && hasAP && !/register\s+of/i.test(parts)) return true;
  if (hasXXVI && hasAppointment) return true;
  if (/letter\s+of\s+appointment/i.test(parts) && /rule\s+30/i.test(parts)) return true;
  return false;
}

export function isFormXXVIAPHeaderFieldLayoutFormHeader(formHeader) {
  return !!formHeader?.formXXVIAPHeaderFieldLayout;
}

export function isFormXXVIAPTableLayoutFormHeader(formHeader) {
  return !!formHeader?.formXXVIAPTableLayout;
}

/**
 * Modal/export employee grid — appointment particulars (clauses 1–4 + To).
 * Establishment / employer / registration stay in header fields.
 */
export const FORM_XXVI_AP_TABLE_HEADERS = [
  'Name of Employee (Sri / Srimathi / Kumari)',
  'Son / Wife / Daughter of',
  'Aged',
  'Date of Birth',
  'Appointed as',
  'With effect from',
  'Category No.',
  'Scale of pay / Rate of increment in wages',
  'Total wages per day / month / week',
  'Basic Pay',
  'Dearness Allowance',
  'Other Allowance',
  'To (Name and Address of Employee)',
];

export function resolveFormXXVIAPTableHeaders(tableHeaders) {
  const list = Array.isArray(tableHeaders)
    ? tableHeaders.map((h) => String(h || '').trim()).filter(Boolean)
    : [];
  if (list.length >= FORM_XXVI_AP_TABLE_HEADERS.length - 2) return list;
  return [...FORM_XXVI_AP_TABLE_HEADERS];
}

/**
 * Fields that exist on the AP Form XXVI appointment letter Excel template
 * (header boxes + numbered clauses 1–4 + To address).
 */
export const FORM_XXVI_AP_TEMPLATE_SPECS = [
  {
    key: 'form_xxvi_ap_establishment',
    label: 'Name and Address of the Establishment:',
    group: 'header',
    fieldType: 'textarea',
    match: /name\s+and\s+address\s+of\s+the\s+establishment/i
  },
  {
    key: 'form_xxvi_ap_employer',
    label: 'Name and Address of the Employer:',
    group: 'header',
    fieldType: 'textarea',
    match: /name\s+and\s+address\s+of\s+the\s+employer/i
  },
  {
    key: 'form_xxvi_ap_registration_number',
    label: 'Registration Number:',
    group: 'header',
    match: /^registration\s+number\s*:?\s*$/i
  },
  {
    key: 'form_xxvi_ap_employee_name',
    label: 'Name of Employee (Sri / Srimathi / Kumari):',
    group: 'clause1',
    sectionMatch: /^\s*1[\.\)]\s+.*sri\s*\/\s*srimathi|kumari.*son\s*\/\s*wife\s*\/\s*daughter/i
  },
  {
    key: 'form_xxvi_ap_father_husband',
    label: 'Son / Wife / Daughter of:',
    group: 'clause1',
    sectionMatch: /son\s*\/\s*wife\s*\/\s*daughter\s+of/i
  },
  {
    key: 'form_xxvi_ap_age',
    label: 'Aged:',
    group: 'clause1',
    sectionMatch: /\baged\b/i
  },
  {
    key: 'form_xxvi_ap_date_of_birth',
    label: 'Date of Birth:',
    group: 'clause1',
    sectionMatch: /date\s+of\s+birth/i
  },
  {
    key: 'form_xxvi_ap_designation',
    label: 'Appointed as:',
    group: 'clause1',
    sectionMatch: /is\s+appointed\s+as/i
  },
  {
    key: 'form_xxvi_ap_appointment_date',
    label: 'With effect from:',
    group: 'clause1',
    sectionMatch: /with\s+effect\s+from/i
  },
  {
    key: 'form_xxvi_ap_category_no',
    label: 'Category No.:',
    group: 'clause2',
    sectionMatch: /^\s*2[\.\)]\s+.*category\s+no/i
  },
  {
    key: 'form_xxvi_ap_scale_of_pay',
    label: 'Scale of pay / Rate of increment in wages:',
    group: 'clause3',
    sectionMatch: /^\s*3[\.\)]\s+.*scale\s+of\s+pay|rate\s+of\s+increment/i
  },
  {
    key: 'form_xxvi_ap_total_wages',
    label: 'Total wages per day / month / week:',
    group: 'clause4',
    sectionMatch: /will\s+draw\s+a\s+total|per\s+day\s*\/\s*month\s*\/\s*week/i
  },
  {
    key: 'form_xxvi_ap_basic_pay',
    label: 'Basic Pay:',
    group: 'clause4',
    sectionMatch: /basic\s+pay/i
  },
  {
    key: 'form_xxvi_ap_dearness_allowance',
    label: 'Dearness Allowance:',
    group: 'clause4',
    sectionMatch: /dearness\s+allowance/i
  },
  {
    key: 'form_xxvi_ap_other_allowance',
    label: 'Other Allowance:',
    group: 'clause4',
    sectionMatch: /other\s+allowance/i
  },
  {
    key: 'form_xxvi_ap_to_address',
    label: 'To (Name and Address of Employee):',
    group: 'footer',
    fieldType: 'textarea',
    match: /^to\s*,?\s*$/i
  }
];

/** Header/footer only — appointment particulars render in the employee table. */
export const FORM_XXVI_AP_FIELD_GROUPS = [
  { id: 'header', title: 'Establishment & Employer' },
];

const isNarrativeBlob = (raw) => {
  const n = formXXVIAPHeaderNorm(raw);
  if (!n) return false;
  return (
    /passport\s+size\s+photo/.test(n) ||
    /^\s*1[\.\)]\s+/.test(n) ||
    /^\s*2[\.\)]\s+/.test(n) ||
    /^\s*3[\.\)]\s+/.test(n) ||
    /^\s*4[\.\)]\s+/.test(n) ||
    /sri\s*\/\s*srimathi|kumari.*son\s*\/\s*wife/.test(n) ||
    /is\s+appointed\s+as/.test(n) ||
    /will\s+draw\s+a\s+total/.test(n) ||
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
  sectionRow: coords.sectionRow ?? null
});

const readValueBelowOrBesideLabel = (getMergedAwareCellText, labelRow, labelCol, effectiveSheetCols) => {
  const maxC = Math.max(20, effectiveSheetCols || 0);
  for (let c = labelCol + 1; c < Math.min(labelCol + 12, maxC); c += 1) {
    const v = String(getMergedAwareCellText(labelRow, c) || '').trim();
    if (v && !/^name\s+and\s+address|^registration\s+number/i.test(formXXVIAPHeaderNorm(v))) {
      return { value: v, valueCol: c };
    }
  }
  for (let r = labelRow + 1; r <= Math.min(labelRow + 8, labelRow + 15); r += 1) {
    const v = String(getMergedAwareCellText(r, labelCol) || '').trim();
    if (v && !/^name\s+and\s+address|^registration\s+number/i.test(formXXVIAPHeaderNorm(v))) {
      return { value: v, valueCol: labelCol, valueRow: r };
    }
  }
  return { value: '', valueCol: labelCol + 1 };
};

const findHeaderLabelCell = (getMergedAwareCellText, matchRe, effectiveSheetCols, maxRows = 120) => {
  const maxC = Math.max(20, effectiveSheetCols || 0);
  for (let r = 0; r < maxRows; r += 1) {
    for (let c = 0; c < maxC; c += 1) {
      const raw = String(getMergedAwareCellText(r, c) || '').trim();
      if (!raw || isNarrativeBlob(raw)) continue;
      const norm = formXXVIAPHeaderNorm(raw);
      if (!matchRe.test(norm) && !matchRe.test(raw)) continue;
      if (norm.length > 90) continue;
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

const findToSection = (getMergedAwareCellText, effectiveSheetCols, maxRows = 120) => {
  const maxC = Math.max(20, effectiveSheetCols || 0);
  for (let r = 0; r < maxRows; r += 1) {
    for (let c = 0; c < maxC; c += 1) {
      const raw = String(getMergedAwareCellText(r, c) || '').trim();
      if (!/^to\s*,?\s*$/i.test(raw)) continue;
      const read = readValueBelowOrBesideLabel(getMergedAwareCellText, r, c, effectiveSheetCols);
      return {
        labelRow: r,
        labelCol: c,
        valueCol: read.valueCol,
        valueRow: read.valueRow ?? r + 1,
        value: read.value,
        sectionRow: r
      };
    }
  }
  return null;
};

/** Build modal fields from the real AP template — no generic HR register columns. */
export function buildFormXXVIAPTemplateFields(getMergedAwareCellText = null, effectiveSheetCols = 20) {
  return FORM_XXVI_AP_TEMPLATE_SPECS.map((spec) => {
    let coords = {};
    if (typeof getMergedAwareCellText === 'function') {
      if (spec.match) {
        coords = findHeaderLabelCell(getMergedAwareCellText, spec.match, effectiveSheetCols) || {};
      } else if (spec.key === 'form_xxvi_ap_to_address') {
        coords = findToSection(getMergedAwareCellText, effectiveSheetCols) || {};
      } else if (spec.sectionMatch) {
        coords = findSectionRow(getMergedAwareCellText, spec.sectionMatch, effectiveSheetCols) || {};
      }
    }
    return buildTemplateField(spec, coords);
  });
}

export function enrichFormXXVIAPFieldsFromSheet(getMergedAwareCellText, effectiveSheetCols, maxRows = 120) {
  return buildFormXXVIAPTemplateFields(getMergedAwareCellText, effectiveSheetCols);
}

export function buildFormXXVIAPFallbackFields() {
  return buildFormXXVIAPTemplateFields();
}

const getRowJoinedText = (getMergedAwareCellText, row, effectiveSheetCols) => {
  const maxC = Math.max(20, effectiveSheetCols || 0);
  let rowText = '';
  for (let c = 0; c < maxC; c += 1) {
    const t = String(getMergedAwareCellText(row, c) || '').trim();
    if (t) rowText += `${t} `;
  }
  return rowText.replace(/\s+/g, ' ').trim();
};

const readMultilineValueBelowLabel = (getMergedAwareCellText, labelRow, labelCol, effectiveSheetCols, maxLines = 6) => {
  const maxC = Math.max(20, effectiveSheetCols || 0);
  const lines = [];
  for (let r = labelRow + 1; r <= labelRow + maxLines; r += 1) {
    let line = '';
    for (let c = labelCol; c < Math.min(labelCol + 4, maxC); c += 1) {
      const v = String(getMergedAwareCellText(r, c) || '').trim();
      if (v) line = line ? `${line} ${v}` : v;
    }
    if (!line) break;
    if (isNarrativeBlob(line)) break;
    const norm = formXXVIAPHeaderNorm(line);
    if (/^name\s+and\s+address|^registration\s+number|^passport\s+size/i.test(norm)) break;
    lines.push(line);
  }
  return lines.join('\n').trim();
};

const cleanExtractedValue = (value) =>
  String(value ?? '')
    .replace(/_+/g, '')
    .replace(/\s+/g, ' ')
    .trim();

const extractNarrativeValueForKey = (key, rowText) => {
  const t = String(rowText || '').replace(/\s+/g, ' ').trim();
  if (!t) return '';
  const pick = (re) => {
    const m = t.match(re);
    return m && m[1] ? cleanExtractedValue(m[1]) : '';
  };

  switch (key) {
    case 'form_xxvi_ap_employee_name':
      return pick(
        /(?:Sri|Srimathi|Kumari)\s*[.,]?\s*(.+?)\s+son\s*\/?\s*wife\s*\/?\s*daughter\s+of/i
      );
    case 'form_xxvi_ap_father_husband':
      return pick(/son\s*\/?\s*wife\s*\/?\s*daughter\s+of\s+(.+?)\s+Aged/i);
    case 'form_xxvi_ap_age':
      return pick(/\bAged\s+(\d{1,3})\b/i);
    case 'form_xxvi_ap_date_of_birth':
      return pick(/Date\s+of\s+Birth\s+(.+?)\s+is\s+appointed\s+as/i);
    case 'form_xxvi_ap_designation':
      return pick(/is\s+appointed\s+as\s+(.+?)\s+with\s+effect\s+from/i);
    case 'form_xxvi_ap_appointment_date':
      return pick(/with\s+effect\s+from\s+(.+?)(?:\s*\.|\s+2[\.\)]|$)/i);
    case 'form_xxvi_ap_category_no':
      return pick(/category\s+No\.?\s*([A-Za-z0-9\-/]+)/i);
    case 'form_xxvi_ap_scale_of_pay':
      return pick(/scale\s+of\s+pay\s+(.+?)(?:\s+rate\s+of\s+increment|\s*\.|\s+4[\.\)]|$)/i) ||
        pick(/rate\s+of\s+increment\s+in\s+wages\s+(.+?)(?:\s*\.|\s+4[\.\)]|$)/i);
    case 'form_xxvi_ap_total_wages':
      return pick(/will\s+draw\s+a\s+total\s+(.+?)\s+per\s+(?:day|month|week)/i);
    case 'form_xxvi_ap_basic_pay':
      return pick(/\(i\)\s*Basic\s+Pay(?:\s+of)?\s+(.+?)(?:\(ii\)|$)/i) ||
        pick(/\bBasic\s+Pay(?:\s+of)?\s+(.+?)(?:\(ii\)|Dearness|$)/i);
    case 'form_xxvi_ap_dearness_allowance':
      return pick(/\(ii\)\s*Dearness\s+Allowance\s+(.+?)(?:\(iii\)|$)/i) ||
        pick(/Dearness\s+Allowance\s+(.+?)(?:\(iii\)|Other|$)/i);
    case 'form_xxvi_ap_other_allowance':
      return pick(/\(iii\)\s*Other\s+Allowance?s?\s+(.+?)(?:\.|$)/i) ||
        pick(/Other\s+Allowance?s?\s+(.+?)(?:\.|$)/i);
    default:
      return '';
  }
};

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

/** Read header-field values from a filled Form XXVI AP Excel workbook. */
export function importFormXXVIAPHeaderFieldsFromWorkbook(workbook, hints = {}) {
  const accessor = buildWorkbookMergedCellAccessor(workbook, hints);
  if (!accessor) return null;

  const { getMergedAwareCellText, effectiveSheetCols } = accessor;
  const fields = buildFormXXVIAPTemplateFields(getMergedAwareCellText, effectiveSheetCols);
  const out = {};
  const narrativeByRow = new Map();

  fields.forEach((field) => {
    const spec = FORM_XXVI_AP_TEMPLATE_SPECS.find((s) => s.key === field.key);
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

    if (field.sectionRow != null) {
      if (!narrativeByRow.has(field.sectionRow)) {
        narrativeByRow.set(
          field.sectionRow,
          getRowJoinedText(getMergedAwareCellText, field.sectionRow, effectiveSheetCols)
        );
      }
      const narrativeVal = extractNarrativeValueForKey(
        field.key,
        narrativeByRow.get(field.sectionRow)
      );
      if (narrativeVal) val = narrativeVal;
    }

    if (field.key === 'form_xxvi_ap_to_address' && field.labelRow != null && !val) {
      val = readMultilineValueBelowLabel(
        getMergedAwareCellText,
        field.labelRow,
        field.labelCol ?? 0,
        effectiveSheetCols
      );
    }

    if (val) out[field.key] = val;
  });

  return out;
}

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
  const appointmentSheet = names.find((n) => /appointment|letter\s+of\s+appointment/i.test(String(n || '')));
  if (appointmentSheet) return appointmentSheet;
  const xxviSheet = names.find((n) => matchesFormXXVIHint(String(n || '')));
  if (xxviSheet) return xxviSheet;
  if (matchesFormXXVIHint(blob)) {
    return names.find((n) => matchesFormXXVIHint(String(n || ''))) || names[0];
  }
  return names[0];
}

export function resolveFormXXVIAPHeaderFieldLayout(parsed, workbook, hints = {}) {
  const formHeader = hints.formHeader || parsed?.formHeader || null;
  const item = hints.item || null;
  const fileName = hints.fileName || hints.formFileName || '';
  const sheetText = hints.sheetText || '';
  if (!isFormXXVIAPAppointmentLetterContext(formHeader, item, fileName, sheetText)) return null;

  let getMergedAwareCellText = null;
  let effectiveSheetCols = 20;

  const accessor = buildWorkbookMergedCellAccessor(workbook, hints);
  if (accessor) {
    getMergedAwareCellText = accessor.getMergedAwareCellText;
    effectiveSheetCols = accessor.effectiveSheetCols;
  }

  const allFields = buildFormXXVIAPTemplateFields(getMergedAwareCellText, effectiveSheetCols);
  // Appointment particulars live in the table — keep establishment / employer / registration in header fields.
  const finalFields = allFields.filter((f) => f.group === 'header');
  const title =
    formHeader?.title ||
    (String(sheetText).toLowerCase().includes('letter of appointment')
      ? 'Form XXVI – Letter of Appointment'
      : 'Form XXVI');

  return {
    formHeader: {
      title,
      subtitle: formHeader?.subtitle || 'Letter of Appointment',
      reference: formHeader?.reference || '[Rule 30]',
      formXXVIAPHeaderFieldLayout: true,
      formXXVIAPTableLayout: true,
      textRows: [],
      fields: finalFields
    },
    headers: resolveFormXXVIAPTableHeaders(parsed?.headers || hints.tableHeaders || null),
    tableData: [],
    headerRowIndex: -1,
    dataStartIndex: 0,
    tableStartCol: 0,
    sheetName: accessor?.sheetName || hints.preferredSheetName || null
  };
}

/** Map site + employee data onto template field keys (not generic register columns). */
export function applyFormXXVIAPAutofillFromEmployee(headerData, empItem, siteContext = {}) {
  const out = headerData && typeof headerData === 'object' ? { ...headerData } : {};
  const emp = empItem || {};
  const establishmentText = String(siteContext.establishmentText || '').trim();
  const employerText = String(siteContext.employerText || '').trim();
  const employeeName = resolveFormXXVIAPEmployeeName(emp);
  const fatherName = resolveFormXXVIAPFatherHusband(emp);
  const dob = resolveFormXXVIAPDateOfBirth(emp);
  const doj =
    emp.Dateofjoining ||
    emp['Dateofjoining'] ||
    emp['Date of Joining'] ||
    emp.DateofJoining ||
    '';
  const address =
    emp.PresentAddress ||
    emp['Present Address'] ||
    emp.Address ||
    emp['Address'] ||
    emp.PermanentAddress ||
    emp['Permanent Address'] ||
    '';

  const fill = (key, value) => {
    const text = String(value ?? '').trim();
    if (!text || !key) return;
    const cur = String(out[key] ?? '').trim();
    if (!cur || /^enter\b/i.test(cur)) out[key] = text;
  };

  fill('form_xxvi_ap_establishment', establishmentText);
  fill('form_xxvi_ap_employer', employerText);
  fill('form_xxvi_ap_employee_name', employeeName);
  fill('form_xxvi_ap_father_husband', fatherName);
  fill('form_xxvi_ap_date_of_birth', dob);
  fill('form_xxvi_ap_designation', emp.Designation || emp['Designation'] || '');
  fill('form_xxvi_ap_appointment_date', doj);
  fill('form_xxvi_ap_to_address', address ? `${employeeName}\n${address}`.trim() : employeeName);

  if (dob) {
    const d = new Date(dob);
    if (!Number.isNaN(d.getTime())) {
      const today = new Date();
      let age = today.getFullYear() - d.getFullYear();
      const md = today.getMonth() - d.getMonth();
      if (md < 0 || (md === 0 && today.getDate() < d.getDate())) age -= 1;
      if (age >= 0 && age < 120) fill('form_xxvi_ap_age', String(age));
    }
  } else if (emp.Age || emp['Age']) {
    fill('form_xxvi_ap_age', emp.Age || emp['Age']);
  }

  return out;
}

export function writeFormXXVIAPFieldsToExcelJsWorksheet(worksheet, headerFormData, parsedFormHeader, helpers = {}) {
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
  const normalize = (txt) => formXXVIAPHeaderNorm(txt);
  const fields = Array.isArray(parsedFormHeader?.fields) ? parsedFormHeader.fields : [];
  const maxScanRows = Math.max(260, worksheet.rowCount + 10);
  const maxScanCols = 20;

  const writeAt = (row, col, value) => {
    const text = String(value ?? '').trim();
    if (!text || row == null || col == null) return;
    worksheet.getCell(row, col).value = text;
  };

  const fillNarrativeSection = (sectionRow, valuesByPhrase) => {
    if (sectionRow == null || sectionRow < 0) return;
    const excelRow = sectionRow + 1;
    for (let c = 1; c <= maxScanCols; c += 1) {
      const raw = excelCellValueToString(worksheet.getCell(excelRow, c)?.value);
      if (!raw) continue;
      let next = raw;
      let changed = false;
      Object.entries(valuesByPhrase).forEach(([phrase, val]) => {
        const v = String(val || '').trim();
        if (!v) return;
        const re = new RegExp(`(${phrase.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')})\\s*_{2,}`, 'i');
        if (re.test(next)) {
          next = next.replace(re, `$1 ${v}`);
          changed = true;
        }
      });
      if (changed) {
        worksheet.getCell(excelRow, c).value = next;
        return;
      }
    }
  };

  fields.forEach((field) => {
    const val = headerFormData[field.key];
    if (val == null || String(val).trim() === '') return;

    if (field.sectionRow != null && field.key !== 'form_xxvi_ap_to_address') {
      if (field.key === 'form_xxvi_ap_employee_name') {
        fillNarrativeSection(field.sectionRow, { 'Sri / Srimathi / Kumari': val });
      } else if (field.key === 'form_xxvi_ap_father_husband') {
        fillNarrativeSection(field.sectionRow, { 'son/wife/daughter of': val, 'son / wife / daughter of': val });
      } else if (field.key === 'form_xxvi_ap_age') {
        fillNarrativeSection(field.sectionRow, { Aged: val });
      } else if (field.key === 'form_xxvi_ap_date_of_birth') {
        fillNarrativeSection(field.sectionRow, { 'Date of Birth': val });
      } else if (field.key === 'form_xxvi_ap_designation') {
        fillNarrativeSection(field.sectionRow, { 'is appointed as': val });
      } else if (field.key === 'form_xxvi_ap_appointment_date') {
        fillNarrativeSection(field.sectionRow, { 'with effect from': val });
      } else if (field.key === 'form_xxvi_ap_category_no') {
        fillNarrativeSection(field.sectionRow, { 'category No.': val, 'category No': val });
      } else if (field.key === 'form_xxvi_ap_scale_of_pay') {
        fillNarrativeSection(field.sectionRow, { 'scale of pay': val, 'rate of increment': val });
      } else if (field.key === 'form_xxvi_ap_total_wages') {
        fillNarrativeSection(field.sectionRow, { 'will draw a total': val });
      } else if (field.key === 'form_xxvi_ap_basic_pay') {
        fillNarrativeSection(field.sectionRow, { 'Basic Pay': val, 'Basic Pay of': val });
      } else if (field.key === 'form_xxvi_ap_dearness_allowance') {
        fillNarrativeSection(field.sectionRow, { 'Dearness Allowance': val });
      } else if (field.key === 'form_xxvi_ap_other_allowance') {
        fillNarrativeSection(field.sectionRow, { 'Other Allowance': val });
      }
      return;
    }

    if (field.labelRow != null && field.valueCol != null) {
      const targetRow = (field.valueRow ?? field.labelRow) + 1;
      writeAt(targetRow, field.valueCol + 1, val);
      return;
    }

    const spec = FORM_XXVI_AP_TEMPLATE_SPECS.find((s) => s.key === field.key);
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

function formatPayrollAmount(value) {
  if (value == null || value === '') return '';
  const n = Number(String(value).replace(/,/g, '').trim());
  if (!Number.isFinite(n)) return String(value).trim();
  return String(Math.round(n * 100) / 100);
}

export function readFormXXVIAPPayrollBasicPay(payrollRow) {
  if (!payrollRow || payrollRow.fetch_error) return '';
  const flat = flattenPayrollEarningColumns(payrollRow);
  return formatPayrollAmount(
    readPayrollScalar(flat, [
      'earned_basic',
      'basic_pay',
      'basic',
      'Basic',
      'Basic Pay',
      'Earned Basic',
    ]) || flat.earned_basic || flat.basic_pay || flat.basic
  );
}

export function readFormXXVIAPPayrollDearnessAllowance(payrollRow) {
  if (!payrollRow || payrollRow.fetch_error) return '';
  const flat = flattenPayrollEarningColumns(payrollRow);
  return formatPayrollAmount(
    readPayrollScalar(flat, ['dearness_allowance', 'Dearness Allowance', 'da', 'DA']) ||
      flat.dearness_allowance
  );
}

function parseFormXXVIAPMoney(value) {
  if (value === '' || value == null) return NaN;
  const n = Number(String(value).replace(/,/g, '').trim());
  return Number.isFinite(n) ? n : NaN;
}

/** Other Allowance = gross_pay − basic − hra */
export function computeFormXXVIAPOtherAllowance(grossPay, basic, hra) {
  const g = parseFormXXVIAPMoney(grossPay);
  if (!Number.isFinite(g)) return '';
  const b = parseFormXXVIAPMoney(basic);
  const h = parseFormXXVIAPMoney(hra);
  const known = (Number.isFinite(b) ? b : 0) + (Number.isFinite(h) ? h : 0);
  const other = Math.round((g - known) * 100) / 100;
  return Number.isFinite(other) ? other : '';
}

export function readFormXXVIAPPayrollHra(payrollRow) {
  if (!payrollRow || payrollRow.fetch_error) return '';
  const wages = readPayrollForm15WageAmounts(payrollRow);
  const flat = flattenPayrollEarningColumns(payrollRow);
  return formatPayrollAmount(
    readPayrollScalar(flat, [
      'hra',
      'HRA',
      'hra_fbp',
      'house_rent_allowance',
      'House Rent Allowance',
    ]) ||
      wages.hra_fbp ||
      wages.hra ||
      flat.hra_fbp ||
      flat.hra
  );
}

/** Other Allowance ← gross_pay − basic − hra (Sample Payroll). */
export function readFormXXVIAPPayrollOtherAllowance(payrollRow) {
  if (!payrollRow || payrollRow.fetch_error) return '';
  const gross =
    readForm10GrossPayAmount(payrollRow) ||
    readFormXXVIAPPayrollTotalWages(payrollRow);
  const basic = readFormXXVIAPPayrollBasicPay(payrollRow);
  const hra = readFormXXVIAPPayrollHra(payrollRow);
  const computed = computeFormXXVIAPOtherAllowance(gross, basic, hra);
  return computed === '' ? '' : formatPayrollAmount(computed);
}

export function readFormXXVIAPPayrollTotalWages(payrollRow) {
  if (!payrollRow || payrollRow.fetch_error) return '';
  const flat = flattenPayrollEarningColumns(payrollRow);
  const grossFromForm10 = readForm10GrossPayAmount(payrollRow);
  if (grossFromForm10 !== '' && grossFromForm10 != null) {
    return formatPayrollAmount(grossFromForm10);
  }
  const gross = readPayrollScalar(flat, [
    'gross_pay',
    'Gross Pay',
    'grossPay',
    'total_earnings',
    'monthly_gross_amount',
  ]);
  if (gross != null && String(gross).trim() !== '') return formatPayrollAmount(gross);
  const basic = Number(readFormXXVIAPPayrollBasicPay(payrollRow) || 0);
  const da = Number(readFormXXVIAPPayrollDearnessAllowance(payrollRow) || 0);
  const hra = Number(readFormXXVIAPPayrollHra(payrollRow) || 0);
  const sum = basic + da + hra;
  return sum > 0 ? formatPayrollAmount(sum) : '';
}

function resolveFormXXVIAPEmployeeAge(emp = {}, dobRaw = '') {
  if (emp.Age || emp['Age']) return String(emp.Age || emp['Age']).trim();
  if (!dobRaw) return '';
  const d = new Date(dobRaw);
  if (Number.isNaN(d.getTime())) return '';
  const today = new Date();
  let age = today.getFullYear() - d.getFullYear();
  const md = today.getMonth() - d.getMonth();
  if (md < 0 || (md === 0 && today.getDate() < d.getDate())) age -= 1;
  if (age >= 0 && age < 120) return String(age);
  return '';
}

/** Prefer EmployeeName from People / Sample Payroll. */
export function resolveFormXXVIAPEmployeeName(emp = {}) {
  const fromEmployeeName = String(
    emp.EmployeeName ||
      emp['EmployeeName'] ||
      emp['Employee Name'] ||
      emp.Employee_Name ||
      emp.employee_name ||
      emp.DisplayName ||
      emp['Display Name'] ||
      emp.displayName ||
      ''
  ).trim();
  if (fromEmployeeName) return fromEmployeeName;
  const fn = String(
    emp.FirstName || emp['FirstName'] || emp.firstName || emp['First Name'] || ''
  ).trim();
  const ln = String(
    emp.LastName || emp['LastName'] || emp.lastName || emp['Last Name'] || ''
  ).trim();
  if (fn && ln) return `${fn} ${ln}`;
  if (fn || ln) return fn || ln;
  return String(emp.Name || emp['Name'] || emp.full_name || emp['Full Name'] || '').trim();
}

export function resolveFormXXVIAPDateOfBirth(emp = {}) {
  return String(
    emp.Date_of_birth ||
      emp['Date_of_birth'] ||
      emp.DateofBirth ||
      emp['Date of Birth'] ||
      emp.Dateofbirth ||
      emp.date_of_birth ||
      emp.DOB ||
      emp.dob ||
      ''
  ).trim();
}

function resolveFormXXVIAPFatherHusband(emp = {}) {
  return String(
    emp.FatherName ||
      emp['Father Name'] ||
      emp.Father_s_Name ||
      emp['Father_s_Name'] ||
      emp.SpouseName ||
      emp['Spouse Name'] ||
      emp.Father_SpouseName ||
      emp['Father/Husband Name'] ||
      ''
  ).trim();
}

function resolveFormXXVIAPToAddress(emp = {}) {
  const name = resolveFormXXVIAPEmployeeName(emp);
  const address = String(
    emp.PresentAddress ||
      emp['Present Address'] ||
      emp.Address ||
      emp['Address'] ||
      emp.PermanentAddress ||
      emp['Permanent Address'] ||
      ''
  ).trim();
  if (name && address) return `${name}\n${address}`;
  return name || address || '';
}

export function isFormXXVIAPEmployeeNameHeader(h) {
  const s = formXXVIAPHeaderNorm(h).replace(/^\d+[\.\)]\s*/, '');
  if (/name\s+and\s+address\s+of\s+employee|^to\b/.test(s)) return false;
  return /name\s+of\s+employee|sri\s*\/\s*srimathi|kumari/.test(s);
}

export function isFormXXVIAPFatherHusbandHeader(h) {
  const s = formXXVIAPHeaderNorm(h).replace(/^\d+[\.\)]\s*/, '');
  return /son\s*\/\s*wife\s*\/\s*daughter\s+of/.test(s);
}

export function isFormXXVIAPAgeHeader(h) {
  const s = formXXVIAPHeaderNorm(h).replace(/^\d+[\.\)]\s*/, '');
  return /^aged$/.test(s) || /^age$/.test(s);
}

export function isFormXXVIAPDateOfBirthHeader(h) {
  const s = formXXVIAPHeaderNorm(h).replace(/^\d+[\.\)]\s*/, '');
  return /date\s+of\s+birth/.test(s);
}

export function isFormXXVIAPDesignationHeader(h) {
  const s = formXXVIAPHeaderNorm(h).replace(/^\d+[\.\)]\s*/, '');
  return /appointed\s+as|designation/.test(s);
}

export function isFormXXVIAPAppointmentDateHeader(h) {
  const s = formXXVIAPHeaderNorm(h).replace(/^\d+[\.\)]\s*/, '');
  return /with\s+effect\s+from|date\s+of\s+joining|appointment\s+date/.test(s);
}

export function isFormXXVIAPCategoryHeader(h) {
  const s = formXXVIAPHeaderNorm(h).replace(/^\d+[\.\)]\s*/, '');
  return /category\s+no/.test(s);
}

export function isFormXXVIAPScaleOfPayHeader(h) {
  const s = formXXVIAPHeaderNorm(h).replace(/^\d+[\.\)]\s*/, '');
  return /scale\s+of\s+pay|rate\s+of\s+increment/.test(s);
}

export function isFormXXVIAPTotalWagesHeader(h) {
  const s = formXXVIAPHeaderNorm(h).replace(/^\d+[\.\)]\s*/, '');
  return /total\s+wages|per\s+day\s*\/\s*month\s*\/\s*week/.test(s);
}

export function isFormXXVIAPBasicPayHeader(h) {
  const s = formXXVIAPHeaderNorm(h).replace(/^\d+[\.\)]\s*/, '');
  return /^basic\s+pay$/.test(s) || /basic\s+pay/.test(s);
}

export function isFormXXVIAPDearnessAllowanceHeader(h) {
  const s = formXXVIAPHeaderNorm(h).replace(/^\d+[\.\)]\s*/, '');
  return /dearness\s+allowance/.test(s);
}

export function isFormXXVIAPOtherAllowanceHeader(h) {
  const s = formXXVIAPHeaderNorm(h).replace(/^\d+[\.\)]\s*/, '');
  return /other\s+allowance/.test(s);
}

export function isFormXXVIAPToAddressHeader(h) {
  const s = formXXVIAPHeaderNorm(h).replace(/^\d+[\.\)]\s*/, '');
  return /^to\b|name\s+and\s+address\s+of\s+employee/.test(s);
}

/** Map one employee (+ optional payroll) onto Form XXVI AP table columns. */
export function applyFormXXVIAPEmployeeToRow(row, emp, headers, helpers = {}) {
  if (!row || !emp || !Array.isArray(headers)) return row;
  const {
    sanitizeValue = (v) => String(v ?? '').trim(),
    formatStatutoryDateDisplay = (v) => String(v || '').trim(),
    payrollRow = null,
  } = helpers;
  const out = { ...row };
  const employeeSource =
    payrollRow && !payrollRow.fetch_error
      ? {
          ...emp,
          EmployeeName:
            emp.EmployeeName ||
            emp['Employee Name'] ||
            emp.Name ||
            payrollRow.EmployeeName ||
            payrollRow.employee_name ||
            payrollRow['Employee Name'] ||
            payrollRow.full_name ||
            '',
          Date_of_birth:
            emp.Date_of_birth ||
            emp.DateofBirth ||
            emp['Date of Birth'] ||
            payrollRow.Date_of_birth ||
            payrollRow.DateofBirth ||
            payrollRow.date_of_birth ||
            '',
        }
      : emp;
  const employeeName = resolveFormXXVIAPEmployeeName(employeeSource);
  const fatherName = resolveFormXXVIAPFatherHusband(emp);
  const dobRaw = resolveFormXXVIAPDateOfBirth(employeeSource);
  const dojRaw =
    emp.Dateofjoining ||
    emp['Dateofjoining'] ||
    emp['Date of Joining'] ||
    emp.DateofJoining ||
    '';
  const designation = emp.Designation || emp['Designation'] || '';
  const category =
    emp.Category ||
    emp['Category'] ||
    emp.CategoryNo ||
    emp['Category No'] ||
    emp.EmployeeCategory ||
    '';
  const scaleOfPay =
    emp.ScaleOfPay ||
    emp['Scale of Pay'] ||
    emp.PayScale ||
    emp['Pay Scale'] ||
    emp.RateOfIncrement ||
    '';

  headers.forEach((header) => {
    if (isFormXXVIAPEmployeeNameHeader(header)) {
      out[header] = sanitizeValue(employeeName);
      return;
    }
    if (isFormXXVIAPFatherHusbandHeader(header)) {
      out[header] = sanitizeValue(fatherName);
      return;
    }
    if (isFormXXVIAPAgeHeader(header)) {
      out[header] = sanitizeValue(resolveFormXXVIAPEmployeeAge(emp, dobRaw));
      return;
    }
    if (isFormXXVIAPDateOfBirthHeader(header)) {
      out[header] = sanitizeValue(formatStatutoryDateDisplay(dobRaw));
      return;
    }
    if (isFormXXVIAPDesignationHeader(header)) {
      out[header] = sanitizeValue(designation);
      return;
    }
    if (isFormXXVIAPAppointmentDateHeader(header)) {
      out[header] = sanitizeValue(formatStatutoryDateDisplay(dojRaw));
      return;
    }
    if (isFormXXVIAPCategoryHeader(header)) {
      out[header] = sanitizeValue(category);
      return;
    }
    if (isFormXXVIAPScaleOfPayHeader(header)) {
      out[header] = sanitizeValue(scaleOfPay);
      return;
    }
    if (isFormXXVIAPTotalWagesHeader(header)) {
      out[header] = sanitizeValue(readFormXXVIAPPayrollTotalWages(payrollRow));
      return;
    }
    if (isFormXXVIAPBasicPayHeader(header)) {
      out[header] = sanitizeValue(readFormXXVIAPPayrollBasicPay(payrollRow));
      return;
    }
    if (isFormXXVIAPDearnessAllowanceHeader(header)) {
      out[header] = sanitizeValue(readFormXXVIAPPayrollDearnessAllowance(payrollRow));
      return;
    }
    if (isFormXXVIAPOtherAllowanceHeader(header)) {
      out[header] = sanitizeValue(readFormXXVIAPPayrollOtherAllowance(payrollRow));
      return;
    }
    if (isFormXXVIAPToAddressHeader(header)) {
      out[header] = sanitizeValue(resolveFormXXVIAPToAddress(employeeSource));
    }
  });
  return out;
}

/** Map modal table row cells onto form_xxvi_ap_* header keys for Excel export. */
export function mergeFormXXVIAPTableRowIntoHeaderData(
  headerData,
  row,
  headers = FORM_XXVI_AP_TABLE_HEADERS
) {
  const out = headerData && typeof headerData === 'object' ? { ...headerData } : {};
  const hdrs = resolveFormXXVIAPTableHeaders(headers);
  const read = (pred) => {
    const header = hdrs.find(pred);
    if (!header || !row || typeof row !== 'object') return '';
    return String(row[header] ?? '').trim();
  };
  const setIf = (key, value) => {
    const text = String(value ?? '').trim();
    if (!text || !key) return;
    out[key] = text;
  };

  setIf('form_xxvi_ap_employee_name', read(isFormXXVIAPEmployeeNameHeader));
  setIf('form_xxvi_ap_father_husband', read(isFormXXVIAPFatherHusbandHeader));
  setIf('form_xxvi_ap_age', read(isFormXXVIAPAgeHeader));
  setIf('form_xxvi_ap_date_of_birth', read(isFormXXVIAPDateOfBirthHeader));
  setIf('form_xxvi_ap_designation', read(isFormXXVIAPDesignationHeader));
  setIf('form_xxvi_ap_appointment_date', read(isFormXXVIAPAppointmentDateHeader));
  setIf('form_xxvi_ap_category_no', read(isFormXXVIAPCategoryHeader));
  setIf('form_xxvi_ap_scale_of_pay', read(isFormXXVIAPScaleOfPayHeader));
  setIf('form_xxvi_ap_total_wages', read(isFormXXVIAPTotalWagesHeader));
  setIf('form_xxvi_ap_basic_pay', read(isFormXXVIAPBasicPayHeader));
  setIf('form_xxvi_ap_dearness_allowance', read(isFormXXVIAPDearnessAllowanceHeader));
  setIf('form_xxvi_ap_other_allowance', read(isFormXXVIAPOtherAllowanceHeader));
  setIf('form_xxvi_ap_to_address', read(isFormXXVIAPToAddressHeader));
  return out;
}

function formXXVIAPRowHasExportData(row, headers) {
  if (!row || typeof row !== 'object') return false;
  const hdrs = resolveFormXXVIAPTableHeaders(headers);
  return hdrs.some((h) => String(row[h] ?? '').trim() !== '');
}

function sanitizeFormXXVIAPHeaderFormData(headerFormData) {
  const out = headerFormData && typeof headerFormData === 'object' ? { ...headerFormData } : {};
  Object.keys(out).forEach((key) => {
    if (out[key] == null) return;
    out[key] = String(out[key]).trim();
  });
  return out;
}

function resolveFormXXVIAPEmployeeDownloadBaseName(row, headers, index) {
  const hdrs = resolveFormXXVIAPTableHeaders(headers);
  const nameHeader = hdrs.find(isFormXXVIAPEmployeeNameHeader);
  const raw = nameHeader ? String(row?.[nameHeader] ?? '').split(/\r?\n/)[0].trim() : '';
  const safe = String(raw || `Employee_${index + 1}`)
    .replace(/[^a-zA-Z0-9._-]+/g, '_')
    .replace(/^_+|_+$/g, '')
    .slice(0, 60);
  return safe || `Employee_${index + 1}`;
}

function allocateUniqueFormXXVIAPDownloadFileName(baseName, usedNames) {
  const root = String(baseName || 'Employee').replace(/\.xlsx?$/i, '');
  const count = usedNames.get(root) || 0;
  usedNames.set(root, count + 1);
  const suffix = count > 0 ? `_${count + 1}` : '';
  return `Form_XXVI_AP_${root}${suffix}.xlsx`;
}

/** One appointment letter XLSX per employee, packaged as ZIP. */
export async function buildFormXXVIAPPerEmployeeDownload({
  templateArrayBuffer,
  mappedData,
  headersToUse,
  parsedFormHeader,
  formFileName,
  headerFormData,
  sheetNameHint,
}) {
  if (!templateArrayBuffer) {
    throw new Error('Original form template buffer is required for Form XXVI AP export.');
  }

  const hdrs = resolveFormXXVIAPTableHeaders(headersToUse);
  const exportRows = (Array.isArray(mappedData) ? mappedData : []).filter((row) =>
    formXXVIAPRowHasExportData(row, hdrs)
  );
  const rowsForZip = exportRows.length > 0 ? exportRows : [null];
  const baseHeaderData = sanitizeFormXXVIAPHeaderFormData(headerFormData);

  // Export must include clause/footer particulars (modal UI keeps them only in the table).
  const formXXVIFields = FORM_XXVI_AP_TEMPLATE_SPECS.map((spec) => {
    const fromParsed = Array.isArray(parsedFormHeader?.fields)
      ? parsedFormHeader.fields.find((f) => f?.key === spec.key)
      : null;
    return {
      key: spec.key,
      label: fromParsed?.label || spec.label,
      group: spec.group,
      fieldType: spec.fieldType || 'text',
      ...(fromParsed || {}),
      key: spec.key,
      label: fromParsed?.label || spec.label,
      group: spec.group,
    };
  });
  const headerForWrite = {
    ...(parsedFormHeader || {}),
    fields: formXXVIFields,
    formXXVIAPHeaderFieldLayout: true,
    formXXVIAPTableLayout: true,
  };

  const zip = new JSZip();
  const usedNames = new Map();
  for (let i = 0; i < rowsForZip.length; i += 1) {
    const row = rowsForZip[i];
    const mergedHeaderData = mergeFormXXVIAPTableRowIntoHeaderData(baseHeaderData, row, hdrs);
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(templateArrayBuffer);
    const sheetCandidates = Array.isArray(workbook.worksheets) ? workbook.worksheets : [];
    const preferred = String(sheetNameHint || '').trim();
    const worksheet =
      (preferred && sheetCandidates.find((ws) => String(ws?.name || '').trim() === preferred)) ||
      sheetCandidates.find((ws) => /appointment|letter\s+of\s+appointment/i.test(String(ws?.name || ''))) ||
      sheetCandidates.find((ws) => /xxvi/i.test(String(ws?.name || ''))) ||
      sheetCandidates[0] ||
      null;
    if (!worksheet) throw new Error('Template worksheet not found.');
    writeFormXXVIAPFieldsToExcelJsWorksheet(worksheet, mergedHeaderData, headerForWrite);
    const out = await workbook.xlsx.writeBuffer();
    const xlsxBytes = out instanceof Uint8Array ? out : new Uint8Array(out);
    const baseName = resolveFormXXVIAPEmployeeDownloadBaseName(row, hdrs, i);
    zip.file(allocateUniqueFormXXVIAPDownloadFileName(baseName, usedNames), xlsxBytes);
    if (i > 0 && i % 15 === 0) {
      await new Promise((resolve) => setTimeout(resolve, 0));
    }
  }

  const zipBase = String(formFileName || parsedFormHeader?.title || 'Form_XXVI_AP')
    .replace(/\.xlsx?$/i, '')
    .replace(/[^a-zA-Z0-9._-]+/g, '_');
  return {
    blob: await zip.generateAsync({ type: 'blob', compression: 'STORE' }),
    fileName: `${zipBase}_Employees.zip`,
  };
}
