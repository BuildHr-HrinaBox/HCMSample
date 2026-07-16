const FORM_I_TAMIL_NADU_NIL_DEFAULT = 'NIL';

const normalizeFormITamilNaduHeaderText = (text) =>
  String(text ?? '')
    .replace(/\r?\n/g, ' ')
    .replace(/[_-]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();

const isBlankLikeFormITamilNaduValue = (value) => {
  const text = String(value ?? '').trim();
  return !text || /^enter\b/i.test(text) || /^select\b/i.test(text);
};

/**
 * Defaults from app/src/Form 1 Tamilnadu.xlsx → sheet "FORM 1"
 * (Register of Workmen). Used for Form I TN autofill + download when
 * People/payroll rows are missing — including PW Form I Register of Fines.
 */
export const FORM_I_TN_DEFAULT_EMPLOYEES = [
  {
    sno: '1',
    empId: 'VE0147',
    name: 'Avudaiappan',
    names: ['Avudaiappan', 'Avudaiappa'],
    designation: 'Assistant Manager',
    department: 'Service',
    dateOfFirstEntry: '01-Dec-2025',
    dateCompleted480Days: '26 Mar 2027',
    dateMadePermanent: '30 May 2026'
  },
  {
    sno: '2',
    empId: 'VE0079',
    name: 'Vijayakumar',
    names: ['Vijayakumar', 'Vijayakumal'],
    designation: 'Senior Engineer',
    department: 'Service',
    dateOfFirstEntry: '01-Dec-2025',
    dateCompleted480Days: '26 Mar 2027',
    dateMadePermanent: '30 May 2026'
  },
  {
    sno: '3',
    empId: 'VE0054',
    name: 'Ajikumar',
    names: ['Ajikumar'],
    designation: 'Assistant Manager',
    department: 'Service',
    dateOfFirstEntry: '01-Dec-2025',
    dateCompleted480Days: '26 Mar 2027',
    dateMadePermanent: '30 May 2026'
  },
  {
    sno: '4',
    empId: 'VE0051',
    name: 'Senthilkannan',
    names: ['Senthilkannan', 'Senthilkann'],
    designation: 'Senior Engineer',
    department: 'Service',
    dateOfFirstEntry: '01-Dec-2025',
    dateCompleted480Days: '26 Mar 2027',
    dateMadePermanent: '30 May 2026'
  },
  {
    sno: '5',
    empId: 'VE0042',
    name: 'Vinu Monikandan',
    names: ['Vinu Monikandan', 'Vinu Monik'],
    designation: 'Senior Engineer',
    department: 'Service',
    dateOfFirstEntry: '01-Dec-2025',
    dateCompleted480Days: '26 Mar 2027',
    dateMadePermanent: '30 May 2026'
  }
];

const formITnHeaderBare = (header) =>
  normalizeFormITamilNaduHeaderText(header)
    .replace(/^\(?\d+\)?\s*/, '')
    .replace(/\s*[-–]\s*\d+\s*$/, '')
    .trim();

export function isFormITamilNaduFinesOrWorkmenDefaultContext({
  fileName = '',
  formFileName = '',
  formHeader = {},
  headers = [],
  sheetText = ''
} = {}) {
  const headerBlob = Array.isArray(headers)
    ? headers.map(normalizeFormITamilNaduHeaderText).join(' ')
    : '';
  const blob = [
    fileName,
    formFileName,
    formHeader?.title,
    formHeader?.subtitle,
    formHeader?.reference,
    sheetText,
    headerBlob
  ]
    .map((part) => normalizeFormITamilNaduHeaderText(part))
    .filter(Boolean)
    .join(' ');

  if (!blob) return false;
  const looksLikeFormI = /\bform\s*i\b/.test(blob) || /\bform\s*1\b/.test(blob);
  const looksLikeTamilNadu = /tamil\s*nadu|tamilnadu/.test(blob);
  if (!looksLikeFormI || !looksLikeTamilNadu) return false;

  // Suspension / subsistence Form_I_-_TamilNadu is a different register.
  if (
    /date\s+of\s+suspension|subsistence\s+allowance|kept\s+under\s+suspension|employees\s+placed\s+under\s+suspension/.test(
      blob
    )
  ) {
    return false;
  }

  return (
    /register\s+of\s+fines/.test(blob) ||
    /register\s+of\s+workmen/.test(blob) ||
    /pw\s*form\s*i/.test(blob) ||
    /act\s+or\s+omission/.test(blob) ||
    /name\s+and\s+address\s+of\s+the\s+workman/.test(blob) ||
    (Array.isArray(headers) &&
      headers.some((h) => {
        const bare = formITnHeaderBare(h);
        return bare === 'name' || /act\s+or\s+omission|fine\s+imposed/.test(bare);
      }))
  );
}

function classifyFormITnDefaultColumn(header) {
  const bare = formITnHeaderBare(header);
  if (!bare) return '';
  if (/^(s\.?\s*no\.?|sl\.?\s*no\.?|serial(\s+number)?)$/.test(bare)) return 'sno';
  if (/emp\s*id|employee\s*id|identification/.test(bare)) return 'empId';
  if (/father|husband|spouse/.test(bare)) return 'father';
  if (/department\s+of\s+gang|^department$/.test(bare) && !/workshop\s+departmental/.test(bare)) {
    return 'department';
  }
  if (/designation/.test(bare)) return 'designation';
  if (/first\s+entry|date\s+of\s+joining|dateofjoining/.test(bare)) return 'dateOfFirstEntry';
  if (/480\s+days/.test(bare)) return 'dateCompleted480Days';
  if (/made\s+permanent/.test(bare)) return 'dateMadePermanent';
  if (
    bare === 'name' ||
    /name\s+and\s+address\s+of\s+the\s+workman/.test(bare) ||
    (/name/.test(bare) && /workman|workmen|employee|worker/.test(bare))
  ) {
    return 'name';
  }
  if (
    /act\s+or\s+omission|fine\s+imposed|showed\s+cause|realis|whether\s+temporary|remarks|signature|total\s+wages|amount\s+of\s+and\s+date/.test(
      bare
    )
  ) {
    return 'leaveBlank';
  }
  return '';
}

/** Register of Fines fine/cause/wage columns — autofill as NIL when no fine was imposed. */
export function isFormITamilNaduFinesNilDefaultHeader(header) {
  const bare = formITnHeaderBare(header);
  if (!bare) return false;
  if (/^(s\.?\s*no\.?|sl\.?\s*no\.?|serial(\s+number)?)$/.test(bare)) return false;
  if (bare === 'name') return false;
  if (/father|husband|spouse/.test(bare)) return false;
  if (/department\s+of\s+gang|^department$/.test(bare) && !/workshop\s+departmental/.test(bare)) {
    return false;
  }
  if (
    /name\s+and\s+address\s+of\s+the\s+workman/.test(bare) ||
    (/name/.test(bare) && /workman|workmen|employee|worker/.test(bare))
  ) {
    return false;
  }
  // Workmen-register-only columns — never NIL from this helper.
  if (/whether\s+temporary|signature|480\s+days|made\s+permanent|first\s+entry|designation/.test(bare)) {
    return false;
  }
  return (
    /act\s+or\s+omission/.test(bare) ||
    /whether\s+workman/.test(bare) ||
    /showed\s+cause/.test(bare) ||
    /total\s+wages/.test(bare) ||
    /amount\s+of\s+and\s+date/.test(bare) ||
    /fine\s+imposed/.test(bare) ||
    /fine\s+realis/.test(bare) ||
    bare === 'remarks'
  );
}

/** Build one row object for Register of Fines or Register of Workmen headers. */
export function buildFormITamilNaduDefaultRow(emp, headers, index = 0) {
  const row = {};
  const list = Array.isArray(headers) ? headers : [];
  list.forEach((header) => {
    row[header] = '';
  });
  list.forEach((header) => {
    const kind = classifyFormITnDefaultColumn(header);
    if (kind === 'sno') row[header] = String(emp.sno || index + 1);
    else if (kind === 'empId') row[header] = emp.empId || '';
    else if (kind === 'name') row[header] = emp.name || '';
    else if (kind === 'designation') row[header] = emp.designation || '';
    else if (kind === 'department') row[header] = emp.department || emp.designation || '';
    else if (kind === 'dateOfFirstEntry') row[header] = emp.dateOfFirstEntry || '';
    else if (kind === 'dateCompleted480Days') row[header] = emp.dateCompleted480Days || '';
    else if (kind === 'dateMadePermanent') row[header] = emp.dateMadePermanent || '';
    else if (kind === 'father') row[header] = '';
    else if (isFormITamilNaduFinesNilDefaultHeader(header)) {
      row[header] = FORM_I_TAMIL_NADU_NIL_DEFAULT;
    } else if (kind === 'leaveBlank') {
      row[header] = '';
    }
  });
  row.__employeeLookupName = emp.name || '';
  row.__employeeLookupId = emp.empId || '';
  return row;
}

export function buildFormITamilNaduDefaultRows(headers) {
  const hdrs = Array.isArray(headers) ? headers : [];
  return FORM_I_TN_DEFAULT_EMPLOYEES.map((emp, index) =>
    buildFormITamilNaduDefaultRow(emp, hdrs, index)
  );
}

function formITnRowHasNamedEmployee(row, headers) {
  if (!row || typeof row !== 'object') return false;
  const list = Array.isArray(headers) && headers.length > 0 ? headers : Object.keys(row);
  for (const header of list) {
    if (String(header || '').startsWith('__')) continue;
    if (classifyFormITnDefaultColumn(header) !== 'name') continue;
    const v = String(row[header] ?? '').trim();
    if (v && !isBlankLikeFormITamilNaduValue(v) && !/^\d+$/.test(v)) return true;
  }
  // Fallback: any known default name value on the row.
  const values = Object.entries(row)
    .filter(([k]) => !String(k).startsWith('__'))
    .map(([, v]) => String(v ?? '').trim().toLowerCase())
    .filter(Boolean);
  return FORM_I_TN_DEFAULT_EMPLOYEES.some((emp) =>
    values.some((v) => emp.names.some((n) => v === String(n).toLowerCase() || v.includes(String(n).toLowerCase())))
  );
}

/**
 * Ensure Form I TN download/autofill has the Form 1 Excel default employees
 * when Name columns are empty (template Sl.No-only rows).
 */
export function ensureFormITamilNaduDefaultEmployeeRows(rows, headers, { force = false } = {}) {
  const hdrs = Array.isArray(headers) ? headers : [];
  if (hdrs.length === 0) return rows;
  const list = Array.isArray(rows) ? rows : [];
  if (!force && list.some((row) => formITnRowHasNamedEmployee(row, hdrs))) {
    return list;
  }
  return buildFormITamilNaduDefaultRows(hdrs);
}

/** Monthly emoluments (Wages) — never autofetch from People/payroll for Form I TN suspension register. */
export function isFormITamilNaduSkipAutofillHeader(header) {
  const text = normalizeFormITamilNaduHeaderText(header);
  if (!text) return false;
  if (text.includes('monthly') && text.includes('emolument')) return true;
  if (text.includes('emolument') && text.includes('wage') && text.includes('employee')) return true;
  return false;
}

export function isFormITamilNaduNilDefaultHeader(header) {
  const text = normalizeFormITamilNaduHeaderText(header);
  if (!text) return false;
  if (isFormITamilNaduSkipAutofillHeader(header)) return false;

  const hasSuspensionDate =
    text.includes('date of suspension') || (text.includes('date') && text.includes('suspension'));
  const hasRevocationDate =
    text.includes('date of revocation of suspension') ||
    (text.includes('revocation') && text.includes('suspension'));
  const hasPunishmentText =
    text.includes('whether the employee') &&
    text.includes('punishment') &&
    (text.includes('exon') || text.includes('award') || text.includes('awaerd'));
  return (
    (text.includes('nature of offence') && text.includes('date of offence')) ||
    hasSuspensionDate ||
    hasRevocationDate ||
    (text.includes('rate') &&
      text.includes('allowance') &&
      text.includes('calculated') &&
      text.includes('period for which') &&
      text.includes('calculation made')) ||
    (text.includes('amount') &&
      text.includes('allowance paid') &&
      text.includes('date of payment')) ||
    hasPunishmentText ||
    text === 'remarks' ||
    (text.includes('signature of employee') &&
      text.includes('receiving money') &&
      text.includes('postal acknowledgement of money order'))
  );
}

export function isFormITamilNaduSuspensionWorkbookContext({
  fileName = '',
  formFileName = '',
  formHeader = {},
  headers = [],
  sheetText = ''
} = {}) {
  const headerBlob = Array.isArray(headers)
    ? headers.map(normalizeFormITamilNaduHeaderText).join(' ')
    : '';
  const blob = [
    fileName,
    formFileName,
    formHeader?.title,
    formHeader?.subtitle,
    formHeader?.reference,
    sheetText,
    headerBlob
  ]
    .map((part) => normalizeFormITamilNaduHeaderText(part))
    .filter(Boolean)
    .join(' ');

  if (!blob) return false;
  const looksLikeFormI = /\bform\s*i\b/.test(blob) || /\bform\s*1\b/.test(blob);
  const looksLikeTamilNadu = /tamil\s*nadu|tamilnadu/.test(blob);
  const looksLikeSuspensionRegister =
    blob.includes('nature of offence') ||
    blob.includes('date of suspension') ||
    blob.includes('revocation of suspension') ||
    blob.includes('subsistence allowance') ||
    blob.includes('postal acknowledgement of money order') ||
    blob.includes('exoncrated') ||
    blob.includes('awaerded') ||
    blob.includes('kept under suspension') ||
    blob.includes('employees placed under suspension') ||
    // Sheet / workbook naming for Form_I_-_TamilNadu.xlsx (SA Form 1).
    /\bsa\s*form\s*1\b/.test(blob);

  return looksLikeFormI && looksLikeTamilNadu && looksLikeSuspensionRegister;
}

/** Clear Monthly emoluments; fill blank offence/suspension/remarks/signature cells with NIL. */
export function applyFormITamilNaduNilDefaultsToRows(rows, headers, { overwriteNil = false } = {}) {
  if (!Array.isArray(rows) || rows.length === 0 || !Array.isArray(headers) || headers.length === 0) {
    return rows;
  }

  return rows.map((row) => {
    const next = row && typeof row === 'object' && !Array.isArray(row) ? { ...row } : {};
    let changed = false;

    headers.forEach((header) => {
      if (isFormITamilNaduSkipAutofillHeader(header)) {
        if (String(next[header] ?? '').trim() !== '') {
          next[header] = '';
          changed = true;
        }
        return;
      }
      const isNilCol =
        isFormITamilNaduNilDefaultHeader(header) || isFormITamilNaduFinesNilDefaultHeader(header);
      if (!isNilCol) return;
      if (!overwriteNil && !isBlankLikeFormITamilNaduValue(next[header])) return;
      if (next[header] === FORM_I_TAMIL_NADU_NIL_DEFAULT) return;
      next[header] = FORM_I_TAMIL_NADU_NIL_DEFAULT;
      changed = true;
    });

    return changed ? next : row;
  });
}

export { FORM_I_TAMIL_NADU_NIL_DEFAULT, normalizeFormITamilNaduHeaderText };

/** Canonical Form I TN Register of Workmen layout from Form 1 Tamilnadu.xlsx → "FORM 1". */
export const FORM_I_TN_WORKMEN_TABLE_COLS = 10;
export const FORM_I_TN_WORKMEN_COL_WIDTHS = [
  5.82, 9, 30.73, 40.45, 45, 15.82, 23.73, 16.45, 9.82, 25.82
];
export const FORM_I_TN_WORKMEN_MERGES = [
  'A1:J1',
  'A2:D2',
  'E2:J2',
  'A3:A4',
  'B3:B4',
  'C3:C4',
  'D3:D4',
  'E3:E4',
  'F3:F4',
  'G3:G4',
  'H3:H4',
  'I3:I4',
  'J3:J4'
];
export const FORM_I_TN_WORKMEN_ROW_HEIGHTS = {
  1: 91.5,
  2: 73.5,
  3: 33.75,
  4: 26.25
};

function formIExcelJsCellText(val) {
  if (val == null) return '';
  if (typeof val === 'string' || typeof val === 'number' || typeof val === 'boolean') return String(val);
  if (val instanceof Date) return val.toISOString();
  if (typeof val === 'object') {
    if (Array.isArray(val.richText)) return val.richText.map((rt) => rt?.text || '').join('');
    if (val.text != null) return String(val.text);
    if (val.result != null) return String(val.result);
  }
  return '';
}

/**
 * ExcelJS Cell.text throws on merge slaves when the master value is null:
 * "Cannot read properties of null (reading 'toString')".
 * Always prefer value extraction — never call cell.text in Form I paths.
 */
export function safeFormIExcelJsCellText(cell) {
  if (!cell) return '';
  try {
    return formIExcelJsCellText(cell.value);
  } catch (_) {
    return '';
  }
}

function cloneFormICellStyle(style) {
  if (!style) return null;
  try {
    const cloned = JSON.parse(JSON.stringify(style));
    if (cloned.font?.color && cloned.font.color.indexed != null) {
      delete cloned.font.color.indexed;
      if (!cloned.font.color.argb && !cloned.font.color.theme) {
        cloned.font.color = { argb: 'FF000000' };
      }
    }
    return cloned;
  } catch (_) {
    return null;
  }
}

function normalizeFormITitleText(raw) {
  return String(raw || '')
    .replace(/_x000d_/gi, '\n')
    .replace(/\u000d/g, '\n')
    .replace(/\r\n/g, '\n')
    .replace(/\r/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

/**
 * True when the ExcelJS sheet is Tamil Nadu Form I Register of Workmen
 * (Conferment of Permanent Status) — sheet "FORM 1".
 */
export function sheetLooksLikeFormITamilNaduWorkmenRegister(worksheet) {
  if (!worksheet) return false;
  const name = String(worksheet.name || '').trim().toLowerCase();
  if (/^form\s*1$/.test(name) || /^form\s*i$/.test(name)) {
    // Prefer name match only when content also looks like workmen (avoid Form U collision).
  }
  let blob = ` ${name} `;
  for (let r = 1; r <= 8; r += 1) {
    for (let c = 1; c <= 12; c += 1) {
      const t = safeFormIExcelJsCellText(worksheet.getCell(r, c)).trim();
      if (t) blob += ` ${t}`;
    }
  }
  const text = blob.toLowerCase();
  if (/register\s+of\s+fines/.test(text)) return false;
  if (/act\s+or\s+omission/.test(text) && /fine\s+imposed/.test(text)) return false;
  if (/date\s+of\s+suspension|subsistence\s+allowance|kept\s+under\s+suspension/.test(text)) {
    return false;
  }
  if (/register\s+of\s+workmen/.test(text)) return true;
  if (/conferment\s+of\s+permanent\s+status/.test(text)) return true;
  if (/name\s+and\s+address\s+of\s+the\s+workman/.test(text) && /480\s+days/.test(text)) return true;
  return false;
}

/**
 * Clone Form I workmen register into a fresh workbook limited to columns A–J.
 * Phantom col groups (数百) on the Form 1 sheet make Excel repair downloads and
 * collapse the A1:J1 title merge into a narrow column A.
 */
export function cloneFormITamilNaduWorkmenWorksheetClean(sourceWs) {
  // eslint-disable-next-line global-require
  const ExcelJS = require('exceljs');
  const outWb = new ExcelJS.Workbook();
  const outWs = outWb.addWorksheet('FORM 1');
  const maxCol = FORM_I_TN_WORKMEN_TABLE_COLS;
  let maxRow = Math.max(17, sourceWs.actualRowCount || 0, sourceWs.rowCount || 0);
  maxRow = Math.min(Math.max(maxRow, 17), 200);

  for (let r = 1; r <= maxRow; r += 1) {
    const srcRow = sourceWs.getRow(r);
    const dstRow = outWs.getRow(r);
    if (srcRow.height != null) dstRow.height = srcRow.height;
    for (let c = 1; c <= maxCol; c += 1) {
      const src = sourceWs.getCell(r, c);
      const dst = outWs.getCell(r, c);
      const raw = src.value;
      if (raw != null && typeof raw === 'object' && Array.isArray(raw.richText)) {
        dst.value = {
          richText: raw.richText.map((rt) => ({
            text: normalizeFormITitleText(rt?.text || ''),
            font: rt?.font ? JSON.parse(JSON.stringify(rt.font)) : undefined
          }))
        };
      } else if (typeof raw === 'string') {
        dst.value = normalizeFormITitleText(raw);
      } else if (raw != null) {
        const asText = formIExcelJsCellText(raw);
        dst.value = asText ? normalizeFormITitleText(asText) : raw;
      }
      const cloned = cloneFormICellStyle(src.style);
      if (cloned) dst.style = cloned;
    }
  }

  FORM_I_TN_WORKMEN_COL_WIDTHS.forEach((w, i) => {
    outWs.getColumn(i + 1).width = w;
  });

  Object.entries(FORM_I_TN_WORKMEN_ROW_HEIGHTS).forEach(([row, height]) => {
    const r = Number(row);
    if (!outWs.getRow(r).height) outWs.getRow(r).height = height;
  });

  FORM_I_TN_WORKMEN_MERGES.forEach((label) => {
    try {
      outWs.mergeCells(label);
    } catch (_) {
      /* ignore */
    }
  });

  const titleCell = outWs.getCell(1, 1);
  const titleText = normalizeFormITitleText(safeFormIExcelJsCellText(titleCell));
  if (titleText) {
    // Prefer plain multiline string — richer and more stable across ExcelJS rewrite.
    titleCell.value = titleText;
  }
  titleCell.alignment = {
    horizontal: 'center',
    vertical: 'middle',
    wrapText: true
  };
  titleCell.font = {
    ...(titleCell.font || {}),
    bold: true,
    size: titleCell.font?.size || 12,
    name: titleCell.font?.name || 'Palatino Linotype'
  };

  // Re-assert header wrap/center (template uses tall merged headers, not always textRotation).
  for (let c = 1; c <= maxCol; c += 1) {
    const cell = outWs.getCell(3, c);
    cell.alignment = {
      ...(cell.alignment || {}),
      horizontal: 'center',
      vertical: 'middle',
      wrapText: true
    };
    if (!cell.font?.bold) {
      cell.font = { ...(cell.font || {}), bold: true };
    }
  }

  return { workbook: outWb, worksheet: outWs };
}

/**
 * Ensure Form I workmen title band stays merged A1:J1 (centered wrap).
 */
export function ensureFormITamilNaduWorkmenTitleLayout(worksheet) {
  if (!worksheet) return false;
  const titleText = normalizeFormITitleText(safeFormIExcelJsCellText(worksheet.getCell(1, 1)));
  if (!titleText || !/form\s*[-–]?\s*i/i.test(titleText)) return false;

  const merges = Array.isArray(worksheet.model?.merges) ? [...worksheet.model.merges] : [];
  const hasTitleMerge = merges.some((m) => /^A1:J1$/i.test(String(m)));
  if (!hasTitleMerge) {
    merges.forEach((label) => {
      if (/^A1:/i.test(String(label))) {
        try {
          worksheet.unMergeCells(label);
        } catch (_) {
          /* ignore */
        }
      }
    });
    try {
      worksheet.mergeCells('A1:J1');
    } catch (_) {
      /* ignore */
    }
  }

  const cell = worksheet.getCell(1, 1);
  cell.value = titleText;
  cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  cell.font = {
    ...(cell.font || {}),
    bold: true,
    size: cell.font?.size || 12,
    name: cell.font?.name || 'Palatino Linotype'
  };
  if (!worksheet.getRow(1).height || worksheet.getRow(1).height < 60) {
    worksheet.getRow(1).height = FORM_I_TN_WORKMEN_ROW_HEIGHTS[1];
  }
  FORM_I_TN_WORKMEN_COL_WIDTHS.forEach((w, i) => {
    worksheet.getColumn(i + 1).width = w;
  });
  return true;
}
