import * as XLSX from 'xlsx';
import { formatStatutoryHeaderLabelValueExport } from '../../utils/statutorySiteCompanyHeaders';

/** Tamil Nadu CLRA Form XVIII — Register of Wages-cum-Muster Roll [Rule 78(1)(a)(i)]. */

export const FORM_XVIII_TN_TITLE = 'Form XVIII – Register of Wages-cum-Muster Roll';
export const FORM_XVIII_TN_SUBTITLE = 'Form of Register of Wages-cum-Muster Roll';
export const FORM_XVIII_TN_REFERENCE = '[See rule 78(1)(a)(i)]';

export function formXVIIITamilNaduHeaderNorm(txt) {
  return String(txt || '')
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

export function matchesFormXVIIIHint(blob) {
  const parts = String(blob || '').toLowerCase();
  if (/form[\s._-]*xxviii(?![a-z])/i.test(parts)) return false;
  return (
    /form[\s._-]*xviii(?![a-z])/i.test(parts) ||
    /form[\s._-]*18(?!\d)/i.test(parts) ||
    /(?:^|[\s._-])xviii(?=[\s._\W-]|$)/i.test(parts)
  );
}

export function isFormXVIIITamilNaduContext(partsOrBlob) {
  const p = String(partsOrBlob || '').toLowerCase();
  return /tamil[\s._-]*nadu|tamilnadu|form_xviii[_\s-]*tamil|form[\s._-]*xviii[_\s-]*tamil/.test(p);
}

function buildFormXVIIIContextBlob(formHeader, rowItem, fileName, tableHeaders, sheetText) {
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

export function isRegisterOfWagesCumMusterRollBlob(blob) {
  const parts = String(blob || '').toLowerCase();
  return (
    matchesFormXVIIIHint(parts) ||
    /register[\s._-]*of[\s._-]*wages[\s._-]*cum[\s._-]*muster/i.test(parts) ||
    /register\s+of\s+wages[\s-]*cum[\s-]*muster\s+roll/i.test(parts)
  );
}

/** Filename / form-name identity for Form XXVI (26) — must not be treated as Form XVIII. */
function identityIndicatesFormXXVINotXVIII(rowItem, fileName) {
  const identity = [
    rowItem?.formName,
    rowItem?.FormName,
    rowItem?.formFileName,
    rowItem?.FormFileName,
    fileName
  ]
    .filter(Boolean)
    .join(' ')
    .toLowerCase();
  if (!identity) return false;
  // Explicit Form XVIII / 18 identity wins.
  if (/form[\s._-]*xviii(?![a-z])/i.test(identity) || /form[\s._-]*18(?!\d)/i.test(identity)) {
    return false;
  }
  // Form 26-A is a different register; still not Form XVIII.
  if (/form[\s._-]*xxvi(?![a-z])/i.test(identity) || /form[\s._-]*26(?!\d)/i.test(identity)) {
    return true;
  }
  return false;
}

/** Tamil Nadu CLRA Form XVIII — wages-cum-muster (distinct from generic/AP Form XVIII workbooks). */
export function isFormXVIIITamilNaduClraContext(
  formHeader,
  rowItem,
  fileName,
  tableHeaders = '',
  sheetText = ''
) {
  // Form_26 / Form_XXVI Tamil Nadu templates sometimes contain leftover Form XVIII title text.
  // Identity (filename / form name) must win over sheet wording.
  if (identityIndicatesFormXXVINotXVIII(rowItem, fileName)) return false;
  const parts = buildFormXVIIIContextBlob(formHeader, rowItem, fileName, tableHeaders, sheetText);
  if (!isRegisterOfWagesCumMusterRollBlob(parts)) return false;
  return isFormXVIIITamilNaduContext(parts);
}

export function isFormXVIIITamilNaduHeaderFieldLayoutFormHeader(formHeader) {
  return !!formHeader?.formXVIIITamilNaduHeaderFieldLayout;
}

/** Ordered header fields for TN Form XVIII — 2-column grid reads left-to-right, top-to-bottom. */
export const FORM_XVIII_TN_HEADER_SPECS = [
  {
    key: 'form_xviii_contractor',
    label: '1. Name and Address of Contractor.',
    fieldType: 'textarea',
    match:
      /^1\.?\s*name\s+and\s+address\s+of\s+(?:the\s+)?contractor|name\s+and\s+address\s+of\s+(?:the\s+)?contractor/i
  },
  {
    key: 'form_xviii_establishment_contract_carried',
    label: '2. Name and address of establishment in/under which contract is carried on.',
    fieldType: 'textarea',
    match: /^2\.?\s*name\s+and\s+address\s+of\s+establishment|establishment[\s\S]*contract\s+is\s+carried\s+on/i
  },
  {
    key: 'form_xviii_nature_location_work',
    label: '3. Nature and location of work.',
    fieldType: 'textarea',
    match: /^3\.?\s*nature\s+and\s+location\s+of\s+work|nature\s+and\s+location\s+of\s+work/i
  },
  {
    key: 'form_xviii_principal_employer',
    label: '4. Name and address of Principal Employer.',
    fieldType: 'textarea',
    match: /^4\.?\s*name\s+and\s+address\s+of\s+principal\s+employer|principal\s+employer/i
  },
  {
    key: 'form_xviii_wage_period',
    label: 'Wage period : Weekly/Fortnightly',
    match: /wage\s*period\s*:?\s*weekly\s*\/\s*fortnightly|wage\s*period/i
  },
  {
    key: 'form_xviii_month_year',
    label: 'Month/Year',
    match: /^month\s*\/\s*year$|^month\s+year$/i
  }
];

const FORM_XVIII_TN_HEADER_KEYS = new Set(FORM_XVIII_TN_HEADER_SPECS.map((s) => s.key));

const GENERIC_SITE_HEADER_KEY_RE =
  /^(statutory_establishment_name|statutory_establishment_address|statutory_establishment_name_shop|form25_establishment)$/;

const isExcludedTnHeaderField = (field) => {
  const key = String(field?.key || '');
  const label = formXVIIITamilNaduHeaderNorm(field?.label);
  if (GENERIC_SITE_HEADER_KEY_RE.test(key)) return true;
  if (/^name\s+of\s+the\s+establishment/.test(label)) return true;
  if (/^address\s+of\s+the\s+establishment/.test(label)) return true;
  if (/^name\s+of\s+establishment\s*\/\s*shop/.test(label)) return true;
  return false;
};

const isMeaningfulHeaderValue = (value, label) => {
  const raw = String(value || '').trim();
  if (!raw) return false;
  const normalizedRaw = formXVIIITamilNaduHeaderNorm(raw);
  const normalizedLabel = formXVIIITamilNaduHeaderNorm(label);
  if (!normalizedRaw || normalizedRaw === normalizedLabel) return false;
  if (normalizedRaw === formXVIIITamilNaduHeaderNorm(`Enter ${String(label || '').replace(/:+$/, '').trim()}`)) {
    return false;
  }
  return true;
};

const labelMatchesSpec = (fieldLabel, specLabel) => {
  const a = formXVIIITamilNaduHeaderNorm(fieldLabel).replace(/^\d+\.\s*/, '');
  const b = formXVIIITamilNaduHeaderNorm(specLabel).replace(/^\d+\.\s*/, '');
  if (!a || !b) return false;
  if (a === b) return true;
  if (a.includes(b) || b.includes(a)) return true;
  return false;
};

const pickValueForSpec = (spec, existingFields = []) => {
  const byKey = existingFields.find((f) => f?.key === spec.key);
  if (byKey && isMeaningfulHeaderValue(byKey.value, spec.label)) {
    return String(byKey.value).trim();
  }
  for (const field of existingFields) {
    if (!field) continue;
    if (field.key === spec.key) continue;
    if (labelMatchesSpec(field.label, spec.label) && isMeaningfulHeaderValue(field.value, spec.label)) {
      return String(field.value).trim();
    }
    if (spec.match.test(formXVIIITamilNaduHeaderNorm(field.label)) && isMeaningfulHeaderValue(field.value, spec.label)) {
      return String(field.value).trim();
    }
  }
  return '';
};

export function finalizeFormXVIIITamilNaduHeaderFields(existingFields = []) {
  const fields = Array.isArray(existingFields) ? existingFields.filter((f) => !isExcludedTnHeaderField(f)) : [];
  return FORM_XVIII_TN_HEADER_SPECS.map((spec) => ({
    label: spec.label,
    key: spec.key,
    fieldType: spec.fieldType || 'text',
    value: pickValueForSpec(spec, fields)
  }));
}

function buildWorkbookMergedCellAccessor(workbook, hints = {}) {
  if (!workbook?.SheetNames?.length) return null;
  const preferred = hints.preferredSheetName || hints.sheetName;
  const sheetName =
    (preferred && workbook.SheetNames.includes(preferred) && preferred) ||
    workbook.SheetNames.find((n) => /form[\s._-]*xviii|wages[\s-]*cum[\s-]*muster/i.test(String(n))) ||
    workbook.SheetNames[0];
  const ws = workbook.Sheets[sheetName];
  if (!ws) return null;
  const merges = ws['!merges'] || [];
  const ref = ws['!ref'];
  let effectiveSheetCols = 20;
  if (ref) {
    try {
      const range = typeof XLSX !== 'undefined' ? XLSX.utils.decode_range(ref) : null;
      if (range) effectiveSheetCols = Math.max(20, range.e.c + 1);
    } catch {
      effectiveSheetCols = 20;
    }
  }
  const rawCell = (r, c) => {
    if (r < 0 || c < 0) return '';
    const cellRef = typeof XLSX !== 'undefined' ? XLSX.utils.encode_cell({ r, c }) : '';
    const cell = cellRef ? ws[cellRef] : null;
    return cell && cell.v != null ? String(cell.v).trim() : '';
  };
  const getMergedAwareCellText = (r, c) => {
    const direct = rawCell(r, c);
    if (direct) return direct;
    for (let i = 0; i < merges.length; i += 1) {
      const m = merges[i];
      if (!m?.s || !m?.e) continue;
      if (r >= m.s.r && r <= m.e.r && c >= m.s.c && c <= m.e.c) {
        const topLeft = rawCell(m.s.r, m.s.c);
        if (topLeft) return topLeft;
      }
    }
    return '';
  };
  return { getMergedAwareCellText, effectiveSheetCols, sheetName };
}

const readValueBesideLabel = (getMergedAwareCellText, labelRow, labelCol, effectiveSheetCols) => {
  const maxC = Math.max(20, effectiveSheetCols || 0);
  for (let c = labelCol + 1; c < Math.min(labelCol + 16, maxC); c += 1) {
    const v = String(getMergedAwareCellText(labelRow, c) || '').trim();
    if (v && !/^enter\b/i.test(v)) return v;
  }
  for (let r = labelRow + 1; r <= labelRow + 4; r += 1) {
    const v = String(getMergedAwareCellText(r, labelCol) || '').trim();
    if (v && !/^enter\b/i.test(v)) return v;
  }
  return '';
};

const scanSheetValueForSpec = (getMergedAwareCellText, spec, effectiveSheetCols, maxRows = 80) => {
  if (!getMergedAwareCellText) return '';
  const maxC = Math.max(20, effectiveSheetCols || 0);
  for (let r = 0; r < maxRows; r += 1) {
    for (let c = 0; c < maxC; c += 1) {
      const raw = String(getMergedAwareCellText(r, c) || '').trim();
      if (!raw) continue;
      const norm = formXVIIITamilNaduHeaderNorm(raw);
      if (!spec.match.test(norm) && !spec.match.test(raw)) continue;
      const value = readValueBesideLabel(getMergedAwareCellText, r, c, effectiveSheetCols);
      if (value) return value;
    }
  }
  return '';
};

export function buildFormXVIIITamilNaduTemplateFields(getMergedAwareCellText, effectiveSheetCols, existingFields = []) {
  return FORM_XVIII_TN_HEADER_SPECS.map((spec) => {
    const fromExisting = pickValueForSpec(spec, existingFields);
    const fromSheet = fromExisting || scanSheetValueForSpec(getMergedAwareCellText, spec, effectiveSheetCols);
    return {
      label: spec.label,
      key: spec.key,
      fieldType: spec.fieldType || 'text',
      value: fromSheet
    };
  });
}

export function enrichFormXVIIITamilNaduDisplayHeader(
  formHeader,
  item = null,
  fileName = '',
  tableHeaders = [],
  sheetText = ''
) {
  if (!isFormXVIIITamilNaduClraContext(formHeader, item, fileName, tableHeaders, sheetText)) {
    return formHeader;
  }
  const base = formHeader && typeof formHeader === 'object' ? { ...formHeader } : {};
  const existing = Array.isArray(base.fields) ? base.fields : [];
  base.title = FORM_XVIII_TN_TITLE;
  if (!/wages[\s-]*cum[\s-]*muster/i.test(String(base.subtitle || ''))) {
    base.subtitle = FORM_XVIII_TN_SUBTITLE;
  }
  if (!/\brule\s*78/i.test(String(base.reference || ''))) {
    base.reference = base.reference || FORM_XVIII_TN_REFERENCE;
  }
  base.formXVIIITamilNaduHeaderFieldLayout = true;
  base.fields = finalizeFormXVIIITamilNaduHeaderFields(existing);
  return base;
}

export function resolveFormXVIIITamilNaduHeaderFieldLayout(parsed, workbook, hints = {}) {
  const formHeader = hints.formHeader || parsed?.formHeader || null;
  const item = hints.item || null;
  const fileName = hints.fileName || hints.formFileName || '';
  const sheetText = hints.sheetText || '';
  const tableHeaders = hints.tableHeaders || parsed?.headers || [];
  if (!isFormXVIIITamilNaduClraContext(formHeader, item, fileName, tableHeaders, sheetText)) {
    return null;
  }

  let getMergedAwareCellText = null;
  let effectiveSheetCols = 20;
  const accessor = buildWorkbookMergedCellAccessor(workbook, hints);
  if (accessor) {
    getMergedAwareCellText = accessor.getMergedAwareCellText;
    effectiveSheetCols = accessor.effectiveSheetCols;
  }

  const existingFields = Array.isArray(formHeader?.fields) ? formHeader.fields : [];
  const finalFields = buildFormXVIIITamilNaduTemplateFields(
    getMergedAwareCellText,
    effectiveSheetCols,
    existingFields
  );

  return {
    formHeader: {
      ...formHeader,
      title: FORM_XVIII_TN_TITLE,
      subtitle: /wages[\s-]*cum[\s-]*muster/i.test(String(formHeader?.subtitle || ''))
        ? formHeader.subtitle
        : FORM_XVIII_TN_SUBTITLE,
      reference: formHeader?.reference || FORM_XVIII_TN_REFERENCE,
      formXVIIITamilNaduHeaderFieldLayout: true,
      fields: finalFields
    }
  };
}

export function isFormXVIIITamilNaduPlaceholderHeaderValue(value) {
  const s = formXVIIITamilNaduHeaderNorm(value);
  if (!s) return true;
  if (/^enter\b/.test(s)) return true;
  if (/^name\s+and\s+address\s+of\s+contractor/.test(s)) return true;
  if (/^name\s+and\s+address\s+of\s+establishment/.test(s)) return true;
  if (/^nature\s+and\s+location\s+of\s+work$/.test(s)) return true;
  if (/^name\s+and\s+address\s+of\s+principal\s+employer/.test(s)) return true;
  return false;
}

const isFormXVIIITamilNaduContractorFieldLabel = (label) => {
  const s = formXVIIITamilNaduHeaderNorm(label).replace(/^\d+\.\s*/, '');
  if (!s || /principal/.test(s)) return false;
  return /name\s+and\s+address\s+of\s+(?:the\s+)?contractor/.test(s);
};

export function applyFormXVIIITamilNaduAutofillFromSite(headerData, context = {}, options = {}) {
  const overwrite = options.overwrite === true;
  const out = headerData && typeof headerData === 'object' ? { ...headerData } : {};
  const fill = (key, value) => {
    const text = String(value ?? '').trim();
    if (!text || !key) return;
    const cur = String(out[key] ?? '').trim();
    if (
      overwrite ||
      !cur ||
      /^enter\b/i.test(cur) ||
      isFormXVIIITamilNaduPlaceholderHeaderValue(cur)
    ) {
      out[key] = text;
    }
  };
  fill('form_xviii_contractor', context.contractorText || '');
  fill('form_xviii_establishment_contract_carried', context.establishmentText || '');
  fill('form_xviii_nature_location_work', context.natureLocationText || '');
  fill('form_xviii_principal_employer', context.principalEmployerText || '');
  fill('form_xviii_month_year', context.monthYearText || '');
  const contractorText = String(context.contractorText ?? '').trim();
  if (contractorText) {
    const headerFields = Array.isArray(options.formHeaderFields) ? options.formHeaderFields : [];
    headerFields.forEach((field) => {
      if (!field?.key || !isFormXVIIITamilNaduContractorFieldLabel(field.label)) return;
      fill(field.key, contractorText);
    });
  }
  return out;
}

/** Write TN Form XVIII header band as "Label : value" (SheetJS download path). */
export function writeFormXVIIITamilNaduHeaderFieldsToSheetJs(
  ws,
  headerFormData,
  parsedFormHeader,
  helpers = {}
) {
  const writeCell = helpers.writeCell;
  if (!ws || !headerFormData || typeof writeCell !== 'function') return;
  const getText =
    typeof helpers.getMergedAwareCellText === 'function'
      ? helpers.getMergedAwareCellText
      : (r, c) => {
          const cell = ws[XLSX.utils.encode_cell({ r, c })];
          return cell && cell.v != null ? String(cell.v).trim() : '';
        };
  const maxRows = Math.max(1, Number(helpers.maxRows) || 14);
  const maxCols = Math.max(20, Number(helpers.maxCols) || 80);
  const written = new Set();

  for (let r = 0; r < maxRows; r += 1) {
    for (let c = 0; c < maxCols; c += 1) {
      const raw = String(getText(r, c) || '').trim();
      if (!raw) continue;
      for (let i = 0; i < FORM_XVIII_TN_HEADER_SPECS.length; i += 1) {
        const spec = FORM_XVIII_TN_HEADER_SPECS[i];
        if (written.has(spec.key)) continue;
        const labelPart = raw.split(':')[0].trim();
        const norm = formXVIIITamilNaduHeaderNorm(labelPart);
        if (!spec.match.test(norm) && !spec.match.test(raw)) continue;
        let val = String(headerFormData[spec.key] ?? '').trim();
        if (!val || isFormXVIIITamilNaduPlaceholderHeaderValue(val)) {
          const fields = Array.isArray(parsedFormHeader?.fields) ? parsedFormHeader.fields : [];
          const hit = fields.find(
            (f) =>
              f?.key === spec.key &&
              String(f.value ?? '').trim() &&
              !isFormXVIIITamilNaduPlaceholderHeaderValue(f.value)
          );
          if (hit) val = String(hit.value).trim();
        }
        if (!val || isFormXVIIITamilNaduPlaceholderHeaderValue(val)) continue;
        const existingAfterColon = raw.includes(':')
          ? String(raw.split(':').slice(1).join(':') || '').trim()
          : '';
        const shouldWrite =
          !existingAfterColon ||
          isFormXVIIITamilNaduPlaceholderHeaderValue(existingAfterColon) ||
          existingAfterColon !== val;
        if (!shouldWrite) continue;
        const text = formatStatutoryHeaderLabelValueExport(spec.label, labelPart, val);
        writeCell(XLSX.utils.encode_cell({ r, c }), text);
        written.add(spec.key);
        break;
      }
      if (/^wage\s*period/i.test(formXVIIITamilNaduHeaderNorm(raw.split(':')[0]))) {
        const periodVal = String(headerFormData.form_xviii_wage_period ?? headerFormData.form_xviii_month_year ?? '').trim();
        if (periodVal && !isFormXVIIITamilNaduPlaceholderHeaderValue(periodVal)) {
          writeCell(
            XLSX.utils.encode_cell({ r, c }),
            formatStatutoryHeaderLabelValueExport('Wage period', raw.split(':')[0], periodVal)
          );
        }
      }
      if (/^month\s*\/\s*year$/i.test(formXVIIITamilNaduHeaderNorm(raw.split(':')[0]))) {
        const monthVal = String(headerFormData.form_xviii_month_year ?? '').trim();
        if (monthVal && !isFormXVIIITamilNaduPlaceholderHeaderValue(monthVal)) {
          writeCell(
            XLSX.utils.encode_cell({ r, c }),
            formatStatutoryHeaderLabelValueExport('Month/Year', raw.split(':')[0], monthVal)
          );
        }
      }
    }
  }
}

/** Clear orphan bordered columns S–V (0-based 18–21) after the 16-column register band. */
export function stripFormXVIIITamilNaduTrailingSheetColumns(ws, options = {}) {
  const writeCell = options.writeCell;
  if (!ws || typeof writeCell !== 'function') return;
  const clearFrom = Number.isFinite(Number(options.clearFromCol))
    ? Number(options.clearFromCol)
    : 18;
  const clearThrough = Number.isFinite(Number(options.clearThroughCol))
    ? Number(options.clearThroughCol)
    : 21;
  if (clearThrough < clearFrom) return;
  const rowEnd = Math.max(
    Number(options.maxRow) || 0,
    (Number(options.dataStartRow) || 0) + (Number(options.dataRowCount) || 0) + 8
  );
  for (let r = 0; r <= rowEnd; r += 1) {
    for (let c = clearFrom; c <= clearThrough; c += 1) {
      const ref = XLSX.utils.encode_cell({ r, c });
      writeCell(ref, '');
      if (ws[ref]) {
        delete ws[ref].s;
      }
    }
  }
  if (ws['!ref']) {
    try {
      const range = XLSX.utils.decode_range(ws['!ref']);
      if (range.e.c >= clearFrom) {
        range.e.c = Math.min(range.e.c, clearFrom - 1);
        ws['!ref'] = XLSX.utils.encode_range(range);
      }
    } catch {
      /* ignore */
    }
  }
}

function firstPayrollMoney(source, keys) {
  if (!source || typeof source !== 'object') return '';
  for (let i = 0; i < keys.length; i += 1) {
    const n = parseFormXVIIITamilNaduMoney(source[keys[i]]);
    if (Number.isFinite(n)) return n;
  }
  return '';
}

export function resolveFormXVIIITamilNaduWageMusterColumnHeaders(headers = []) {
  const list = Array.isArray(headers) ? headers : [];
  const findHeader = (testFn) => list.find((h) => testFn(h)) || null;
  return {
    ...resolveFormXVIIITamilNaduPayrollColumnHeaders(headers),
    serialRegister: findHeader(isFormXVIIITamilNaduSerialRegisterHeader),
    dailyAttendance: findHeader(isFormXVIIITamilNaduDailyAttendanceHeader),
    attendanceUnits: findHeader(isFormXVIIITamilNaduTotalAttendanceHeader),
    dailyRate: findHeader(isFormXVIIITamilNaduDailyRateHeader),
    basicWages: findHeader((h) => {
      const s = formXVIIITamilNaduHeaderNorm(h);
      return s.includes('basic') && s.includes('wage') && !s.includes('daily');
    }),
    dearnessAllowance: findHeader((h) => {
      const s = formXVIIITamilNaduHeaderNorm(h);
      return s.includes('dearness') && s.includes('allowance');
    }),
    overtime: findHeader(isFormXVIIITamilNaduOvertimeHeader),
    otherCash: findHeader(isFormXVIIITamilNaduOtherCashPaymentsHeader),
  };
}

/** Apply payroll gross_pay / net_pay (and wage band) onto one export row. */
export function applyFormXVIIITamilNaduPayrollToExportRow(row, payrollRow, headers, options = {}) {
  if (!row || typeof row !== 'object' || !payrollRow || payrollRow.fetch_error) return row;
  const overwrite = options.overwrite !== false;
  const cols = resolveFormXVIIITamilNaduWageMusterColumnHeaders(headers);
  const out = { ...row };
  const setNum = (header, value) => {
    if (!header || value === '' || value == null) return;
    const n = parseFormXVIIITamilNaduMoney(value);
    if (!Number.isFinite(n)) return;
    if (!overwrite && String(out[header] ?? '').trim() !== '') return;
    out[header] = n;
  };

  const gross = firstPayrollMoney(payrollRow, ['gross_pay', 'Gross_Pay', 'grossPay', 'total_earnings']);
  const net = firstPayrollMoney(payrollRow, ['net_pay', 'Net_Pay', 'netPay']);
  const basic = firstPayrollMoney(payrollRow, ['basic', 'basic_pay', 'Basic', 'basic_wage']);
  const hra = firstPayrollMoney(payrollRow, ['hra', 'hra_fbp', 'HRA']);
  const paidDays = firstPayrollMoney(payrollRow, ['paid_days', 'Paid_days', 'paidDays']);

  setNum(cols.dailyRate, gross);
  setNum(cols.totalAmount, gross);
  setNum(cols.netAmountPaid, net);
  setNum(cols.basicWages, basic);
  setNum(cols.dailyAttendance, paidDays);
  setNum(cols.attendanceUnits, paidDays);
  if (cols.dearnessAllowance) {
    const da = firstPayrollMoney(payrollRow, ['da', 'dearness_allowance', 'dearness']);
    setNum(cols.dearnessAllowance, da);
  }
  if (cols.otherCash) {
    const other = computeFormXVIIITamilNaduOtherCashPayments(gross, basic, hra);
    if (other !== '') setNum(cols.otherCash, other);
  }
  if (cols.deductions) {
    const ded = computeFormXVIIITamilNaduDeductions(gross, net);
    if (ded !== '') {
      if (overwrite || looksLikeFormXVIIITamilNaduStaleDeductionCell(out[cols.deductions])) {
        out[cols.deductions] = ded;
      }
    }
  }
  return out;
}

export function enrichFormXVIIITamilNaduPayrollExportRows(
  rows,
  headers,
  { resolvePayrollRow, employees = [], overwrite = true } = {}
) {
  if (!Array.isArray(rows) || rows.length === 0) return [];
  return rows.map((row, index) => {
    let next = row;
    const resolver = typeof resolvePayrollRow === 'function' ? resolvePayrollRow : null;
    const empItem = Array.isArray(employees) ? employees[index] : null;
    const payrollRow = resolver ? resolver(empItem, row, index) : null;
    if (payrollRow && !payrollRow.fetch_error) {
      next = applyFormXVIIITamilNaduPayrollToExportRow(row, payrollRow, headers, { overwrite });
    }
    return enrichFormXVIIITamilNaduExportRows([next], headers, { overwriteDeductions: true })[0];
  });
}

/** SL.No / Serial No. in register of workmen (not the leading Sl No. column). */
export function isFormXVIIITamilNaduSerialRegisterHeader(header) {
  const s = formXVIIITamilNaduHeaderNorm(header);
  if (!s) return false;
  if (!(s.includes('register') && (s.includes('workmen') || s.includes('workman')))) return false;
  return (
    s.includes('serial') ||
    /\bsl\.?\s*no\b/.test(s) ||
    /\bs\.?\s*no\b/.test(s) ||
    s.includes('sl.no') ||
    s.includes('sl no')
  );
}

/** Daily attendance /units worked — not the Total attendance column. */
export function isFormXVIIITamilNaduDailyAttendanceHeader(header) {
  const s = formXVIIITamilNaduHeaderNorm(header);
  if (!s) return false;
  if (s.includes('total') || s.includes('done')) return false;
  return (
    (s.includes('daily') && s.includes('attendance')) ||
    (s.includes('daily') && s.includes('units') && s.includes('work'))
  );
}

/** Total attendance/ units of work done. */
export function isFormXVIIITamilNaduTotalAttendanceHeader(header) {
  const s = formXVIIITamilNaduHeaderNorm(header);
  if (!s) return false;
  if (isFormXVIIITamilNaduDailyAttendanceHeader(header)) return false;
  return (
    (s.includes('total') && s.includes('attendance')) ||
    (s.includes('units') && s.includes('work') && s.includes('done'))
  );
}

/** Daily rate of wages/piece-rate ← gross_pay. */
export function isFormXVIIITamilNaduDailyRateHeader(header) {
  const s = formXVIIITamilNaduHeaderNorm(header);
  if (!s) return false;
  return (
    (s.includes('daily') && s.includes('rate') && s.includes('wage')) ||
    s.includes('piece-rate') ||
    s.includes('piece rate')
  );
}

/** Other cash payments (nature of payment to be indicated) ← gross_pay − basic − hra. */
export function isFormXVIIITamilNaduOtherCashPaymentsHeader(header) {
  const s = formXVIIITamilNaduHeaderNorm(header);
  if (!s) return false;
  return s.includes('other') && s.includes('cash') && (s.includes('payment') || s.includes('nature'));
}

/** Overtime under Amount of wages earned — default display text when no OT earning. */
export const FORM_XVIII_TN_OVERTIME_NIL = 'Nil';

export function isFormXVIIITamilNaduOvertimeHeader(header) {
  const s = formXVIIITamilNaduHeaderNorm(header);
  if (!s) return false;
  if (s.includes('other')) return false;
  return /\bover[\s-]*time\b/.test(s) || s === 'ot' || s.startsWith('ot ');
}

export function applyFormXVIIITamilNaduOvertimeNilToRow(row, headers, helpers = {}) {
  const hdrs = Array.isArray(headers) ? headers : [];
  const out = row && typeof row === 'object' ? { ...row } : {};
  const nilText = helpers.nilText != null ? String(helpers.nilText) : FORM_XVIII_TN_OVERTIME_NIL;
  const { overwrite = false } = helpers;

  hdrs.forEach((header) => {
    if (!isFormXVIIITamilNaduOvertimeHeader(header)) return;
    const existing = String(out[header] ?? '').trim();
    if (
      !overwrite &&
      existing &&
      !/^enter\b/i.test(existing) &&
      !/^nil+$/i.test(existing) &&
      existing.toLowerCase() !== 'n/a' &&
      existing !== '-' &&
      existing !== '—'
    ) {
      return;
    }
    // Keep a real overtime amount when present (unless overwrite).
    if (!overwrite && existing && /^-?\d+(\.\d+)?$/.test(existing.replace(/,/g, ''))) {
      return;
    }
    out[header] = nilText;
  });
  return out;
}

export function applyFormXVIIITamilNaduOvertimeNilToMappedRows(
  mappedData,
  headers,
  nilText = FORM_XVIII_TN_OVERTIME_NIL,
  helpers = {}
) {
  if (!Array.isArray(mappedData)) return [];
  const { overwrite = false } = helpers;
  return mappedData.map((row) =>
    applyFormXVIIITamilNaduOvertimeNilToRow(row, headers, { nilText, overwrite })
  );
}

function parseFormXVIIITamilNaduMoney(value) {
  if (value == null || value === '') return NaN;
  const n = Number(String(value).replace(/,/g, '').trim());
  return Number.isFinite(n) ? n : NaN;
}

/** Count filled wage-muster body columns (excludes serial-only stubs from sparse SampleData). */
export function countFormXVIIITamilNaduWageMusterBodyCells(row, headers = []) {
  if (!row || typeof row !== 'object') return 0;
  const hdrs =
    Array.isArray(headers) && headers.length > 0 ? headers : Object.keys(row);
  let n = 0;
  for (let i = 0; i < hdrs.length; i += 1) {
    const h = hdrs[i];
    const direct = row[h];
    let v = direct;
    if (v == null || String(v).trim() === '') {
      const target = formXVIIITamilNaduHeaderNorm(h);
      if (target) {
        const hit = Object.keys(row).find((k) => formXVIIITamilNaduHeaderNorm(k) === target);
        if (hit) v = row[hit];
      }
    }
    if (v == null || String(v).trim() === '') continue;
    const s = formXVIIITamilNaduHeaderNorm(h);
    if (/^(sl|s)\s*no|serial\s*number|^column\s+\d+$/i.test(s)) continue;
    if (
      /name of employee|name of workman|designation|nature of work|basic\s+wages|dearness|overtime|deduction|net\s+amount|total\b/i.test(
        s
      ) ||
      isFormXVIIITamilNaduDailyAttendanceHeader(h) ||
      isFormXVIIITamilNaduTotalAttendanceHeader(h) ||
      isFormXVIIITamilNaduDailyRateHeader(h) ||
      isFormXVIIITamilNaduOtherCashPaymentsHeader(h)
    ) {
      n += 1;
    }
  }
  return n;
}

export function formXVIIITamilNaduRowsRicherThan(candidateRows, candidateHeaders, baseRows, baseHeaders) {
  const cand = Array.isArray(candidateRows) ? candidateRows : [];
  const base = Array.isArray(baseRows) ? baseRows : [];
  if (cand.length === 0) return false;
  if (base.length === 0) return true;
  const candScore = cand.reduce(
    (sum, row) => sum + countFormXVIIITamilNaduWageMusterBodyCells(row, candidateHeaders),
    0
  );
  const baseScore = base.reduce(
    (sum, row) => sum + countFormXVIIITamilNaduWageMusterBodyCells(row, baseHeaders),
    0
  );
  return candScore > baseScore;
}

/** Deductions, if any (indicate nature) — not Net Amount Paid. */
export function isFormXVIIITamilNaduDeductionsHeader(header) {
  const s = formXVIIITamilNaduHeaderNorm(header);
  if (!s) return false;
  if (s.includes('net') && s.includes('amount') && s.includes('paid')) return false;
  if (s.includes('deduction') && s.includes('other') && !s.includes('if any')) return false;
  return (
    (s.includes('deduction') &&
      (s.includes('indicate') || s.includes('nature') || s.includes('if any'))) ||
    (s.includes('total') && s.includes('deduction'))
  );
}

/** Net Amount Paid column under Amount of wages earned. */
export function isFormXVIIITamilNaduNetAmountPaidHeader(header) {
  const s = formXVIIITamilNaduHeaderNorm(header);
  return s.includes('net') && s.includes('amount') && s.includes('paid');
}

/** Gross Total under Amount of wages earned (not attendance/units totals). */
export function isFormXVIIITamilNaduTotalAmountHeader(header) {
  const s = formXVIIITamilNaduHeaderNorm(header);
  if (!s) return false;
  if (s.includes('attendance') || s.includes('units') || s.includes('work done')) return false;
  if (s.includes('net') || s.includes('deduction')) return false;
  if (s === 'total') return true;
  return s.includes('total') && s.includes('amount');
}

export function isFormXVIIITamilNaduEmployeeNameHeader(header) {
  const s = formXVIIITamilNaduHeaderNorm(header);
  return (
    s.includes('name') &&
    (s.includes('employee') || s.includes('workman') || s.includes('worker')) &&
    !s.includes('establishment') &&
    !s.includes('employer')
  );
}

/** Title-case person names for Excel export (lookup helpers store lowercase). */
export function toFormXVIIITamilNaduPersonNameDisplay(value) {
  const text = String(value || '')
    .trim()
    .replace(/\s+/g, ' ');
  if (!text) return '';
  return text
    .split(' ')
    .map((word) => {
      if (!word) return word;
      return word.charAt(0).toUpperCase() + word.slice(1).toLowerCase();
    })
    .join(' ');
}

/** Deductions = gross_pay − net_pay (Total − Net Amount Paid on the export row). */
export function computeFormXVIIITamilNaduDeductions(grossPay, netPay) {
  const g = parseFormXVIIITamilNaduMoney(grossPay);
  const n = parseFormXVIIITamilNaduMoney(netPay);
  if (!Number.isFinite(g) || !Number.isFinite(n) || g < n) return '';
  return Math.round((g - n) * 100) / 100;
}

/** TN Form XVIII template column-number row: 12=Total, 13=Deductions, 14=Net Amount Paid. */
const FORM_XVIII_TN_WAGE_COLUMN_NUM = {
  total: 12,
  deductions: 13,
  net: 14,
};

function findFormXVIIITamilNaduHeaderByColumnNumber(headers, columnNumber) {
  const list = Array.isArray(headers) ? headers : [];
  const want = String(columnNumber);
  const direct = list.find((h) => String(h ?? '').trim() === want);
  if (direct) return direct;
  return (
    list.find((h) => {
      const m = String(h ?? '').match(/^(.+)_(\d{1,2})$/);
      return m && m[2] === want;
    }) || null
  );
}

/** Wage band uses Excel column numbers (13) when merged leaf text is missing. */
export function formXVIIITamilNaduWageRegisterHasNumericBand(headers = []) {
  const list = Array.isArray(headers) ? headers : [];
  const dedNum = findFormXVIIITamilNaduHeaderByColumnNumber(
    list,
    FORM_XVIII_TN_WAGE_COLUMN_NUM.deductions
  );
  if (!dedNum || isFormXVIIITamilNaduDeductionsHeader(dedNum)) return false;
  const hasTotal =
    list.some((h) => isFormXVIIITamilNaduTotalAmountHeader(h)) ||
    Boolean(
      findFormXVIIITamilNaduHeaderByColumnNumber(list, FORM_XVIII_TN_WAGE_COLUMN_NUM.total)
    );
  const hasNet =
    list.some((h) => isFormXVIIITamilNaduNetAmountPaidHeader(h)) ||
    Boolean(findFormXVIIITamilNaduHeaderByColumnNumber(list, FORM_XVIII_TN_WAGE_COLUMN_NUM.net));
  return hasTotal && hasNet;
}

export function resolveFormXVIIITamilNaduPayrollColumnHeaders(headers = []) {
  const list = Array.isArray(headers) ? headers : [];
  const findHeader = (testFn) => list.find((h) => testFn(h)) || null;
  const numericBand = formXVIIITamilNaduWageRegisterHasNumericBand(list);
  const byNum = (n) =>
    numericBand ? findFormXVIIITamilNaduHeaderByColumnNumber(list, n) : null;
  return {
    employeeName: findHeader(isFormXVIIITamilNaduEmployeeNameHeader),
    totalAmount:
      findHeader(isFormXVIIITamilNaduTotalAmountHeader) ||
      byNum(FORM_XVIII_TN_WAGE_COLUMN_NUM.total),
    netAmountPaid:
      findHeader(isFormXVIIITamilNaduNetAmountPaidHeader) ||
      byNum(FORM_XVIII_TN_WAGE_COLUMN_NUM.net),
    deductions:
      findHeader(isFormXVIIITamilNaduDeductionsHeader) ||
      byNum(FORM_XVIII_TN_WAGE_COLUMN_NUM.deductions),
  };
}

/** Match deductions column by full header text or numeric key "13" / "…_13". */
export function isFormXVIIITamilNaduResolvedDeductionsHeader(header, headers = []) {
  if (!header) return false;
  if (isFormXVIIITamilNaduDeductionsHeader(header)) return true;
  const cols = resolveFormXVIIITamilNaduPayrollColumnHeaders(headers);
  return cols.deductions != null && cols.deductions === header;
}

/** Template/sample cells often show Excel time text (e.g. 07:45) instead of rupee deductions. */
export function looksLikeFormXVIIITamilNaduStaleDeductionCell(value) {
  const s = String(value ?? '').trim();
  if (!s) return true;
  if (/^enter\b/i.test(s)) return true;
  if (/^\d{1,2}:\d{2}$/.test(s)) return true;
  const n = parseFormXVIIITamilNaduMoney(s);
  return !Number.isFinite(n);
}

/** Modal display: replace stale template time with Total − Net Amount Paid when available. */
export function resolveFormXVIIITamilNaduDeductionsCellValue(row, headers) {
  if (!row || typeof row !== 'object') return '';
  const cols = resolveFormXVIIITamilNaduPayrollColumnHeaders(headers);
  if (!cols.deductions) return '';
  const existing = row[cols.deductions];
  if (!looksLikeFormXVIIITamilNaduStaleDeductionCell(existing)) {
    return existing;
  }
  const gross = cols.totalAmount ? row[cols.totalAmount] : '';
  const net = cols.netAmountPaid ? row[cols.netAmountPaid] : '';
  const ded = computeFormXVIIITamilNaduDeductions(gross, net);
  return ded !== '' ? ded : '';
}

/** Download/save pass: Title Case names and Deductions ← Total − Net Amount Paid. */
export function enrichFormXVIIITamilNaduExportRows(rows, headers, options = {}) {
  if (!Array.isArray(rows) || rows.length === 0) return [];
  const cols = resolveFormXVIIITamilNaduPayrollColumnHeaders(headers);
  const { overwriteDeductions = true } = options;
  return rows.map((row) => {
    if (!row || typeof row !== 'object') return row;
    const out = { ...row };
    if (cols.employeeName) {
      const cur = String(out[cols.employeeName] ?? '').trim();
      if (cur && !/^enter\b/i.test(cur)) {
        out[cols.employeeName] = toFormXVIIITamilNaduPersonNameDisplay(cur);
      }
    }
    const gross = cols.totalAmount ? out[cols.totalAmount] : '';
    const net = cols.netAmountPaid ? out[cols.netAmountPaid] : '';
    const ded = computeFormXVIIITamilNaduDeductions(gross, net);
    if (cols.deductions && ded !== '') {
      const existing = out[cols.deductions];
      if (
        overwriteDeductions ||
        looksLikeFormXVIIITamilNaduStaleDeductionCell(existing)
      ) {
        out[cols.deductions] = ded;
      }
    }
    return out;
  });
}

/** Other cash = gross_pay − basic − hra (blank when basic/hra missing so we never dump full gross). */
export function computeFormXVIIITamilNaduOtherCashPayments(grossPay, basic, hra) {
  const g = parseFormXVIIITamilNaduMoney(grossPay);
  if (!Number.isFinite(g)) return '';
  const b = parseFormXVIIITamilNaduMoney(basic);
  const h = parseFormXVIIITamilNaduMoney(hra);
  const hasBasic = Number.isFinite(b) && b > 0;
  const hasHra = Number.isFinite(h) && h > 0;
  if (!hasBasic && !hasHra) return '';
  const known = (hasBasic ? b : 0) + (hasHra ? h : 0);
  if (!(known > 0)) return '';
  const other = Math.round((g - known) * 100) / 100;
  if (!Number.isFinite(other) || other < 0 || other === g) return '';
  return other;
}

export { FORM_XVIII_TN_HEADER_KEYS };
