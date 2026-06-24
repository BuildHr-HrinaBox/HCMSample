import * as XLSX from 'xlsx';
import { isFormXXVITamilNaduClraContext, matchesFormXXVIHint } from './formXXVITamilNaduMuster';
import { labelMatchScore } from './form18APAccidentNotice';
import { isForm2APChangeNoticeContext } from './form2APChangeNotice';

/** AP Shops Form XXVI — Letter of Appointment [Rule 30]. Matches the official Excel narrative layout. */

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

export const FORM_XXVI_AP_FIELD_GROUPS = [
  { id: 'header', title: 'Establishment & Employer' },
  { id: 'clause1', title: '1. Appointment particulars' },
  { id: 'clause2', title: '2. Category' },
  { id: 'clause3', title: '3. Scale of pay' },
  { id: 'clause4', title: '4. Wages composition' },
  { id: 'footer', title: 'To' }
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

  const finalFields = buildFormXXVIAPTemplateFields(getMergedAwareCellText, effectiveSheetCols);
  const title =
    formHeader?.title ||
    (String(sheetText).toLowerCase().includes('letter of appointment')
      ? 'Form XXVI – Letter of Appointment'
      : 'Form XXVI');

  return {
    formHeader: {
      title,
      subtitle: formHeader?.subtitle || '',
      reference: formHeader?.reference || '',
      formXXVIAPHeaderFieldLayout: true,
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

/** Map site + employee data onto template field keys (not generic register columns). */
export function applyFormXXVIAPAutofillFromEmployee(headerData, empItem, siteContext = {}) {
  const out = headerData && typeof headerData === 'object' ? { ...headerData } : {};
  const emp = empItem || {};
  const establishmentText = String(siteContext.establishmentText || '').trim();
  const employerText = String(siteContext.employerText || '').trim();
  const employeeName =
    emp.Name || emp['Name'] || emp.EmployeeName || emp['Employee Name'] || '';
  const fatherName =
    emp.FatherName ||
    emp['Father Name'] ||
    emp.SpouseName ||
    emp['Spouse Name'] ||
    emp.Father_SpouseName ||
    '';
  const dob = emp.DateofBirth || emp['Date of Birth'] || emp.DOB || '';
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
