import * as XLSX from 'xlsx';

const FORM_VI_TN_HOLIDAY_KEYWORDS = [
  'national and festival holidays',
  'festival holidays',
  'approved festival holidays',
  'pongal',
  'republic day',
  'tamil new year',
  'good friday',
  'may day',
  'ramzan',
  'independence day',
  'krishna jayanthi',
  'vinayakar chathurthi',
  'ayudha pooja',
  'vijaya dashami',
  'diwali',
  'christmas'
];

const FORM_VI_HINT_REGEX = /\bform[\s._-]*vi(?:[\s._-]|$)|\bform[\s._-]*6(?:[\s._-]|$)/;
const FORM_VI_TN_HOLIDAY_TITLE_REGEX =
  /register\s+of\s+national\s+and\s+festival\s+holidays?|national\s+and\s+festival\s+holidays?|festival\s+holidays\s+approval\s+proceedings/i;
const FORM_VI_TN_EMPLOYEE_DETAIL_KEYWORDS = [
  'employee identification no',
  'employee identification number',
  'employee details',
  'personal details',
  'date of birth',
  'date of joining',
  'gender',
  'father / spouse',
  "father's/husband's name",
  'insurance corporation',
  'aadhaar'
];

function normalizeFormVITamilNaduText(txt) {
  return String(txt || '')
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

function buildFormVITamilNaduContextBlob(formHeader, rowItem, fileName, sheetText = '', tableHeaders = []) {
  return [
    rowItem?.formName,
    rowItem?.FormName,
    rowItem?.formFileName,
    rowItem?.FormFileName,
    fileName,
    formHeader?.title,
    formHeader?.subtitle,
    formHeader?.reference,
    sheetText,
    Array.isArray(tableHeaders) ? tableHeaders.join(' ') : ''
  ]
    .filter((x) => x != null && String(x).trim() !== '')
    .join(' ')
    .toLowerCase();
}

export function isFormVITamilNaduHolidayContext(formHeader, rowItem, fileName, sheetText = '', tableHeaders = []) {
  const parts = buildFormVITamilNaduContextBlob(formHeader, rowItem, fileName, sheetText, tableHeaders);
  const hasFormVI =
    FORM_VI_HINT_REGEX.test(parts) ||
    parts.includes('form_vi') ||
    parts.includes('form vi') ||
    parts.includes('form-vi');
  const hasTamilNadu =
    /tamil\s*nadu|tamilnadu/.test(parts) || parts.includes('tamil nadu') || parts.includes('tn');
  const hasHolidayHint =
    /festival\s+holidays?|national\s+and\s+festival\s+holidays?/.test(parts) ||
    FORM_VI_TN_HOLIDAY_KEYWORDS.some((keyword) => parts.includes(keyword));
  if (!hasFormVI) return false;
  if (!hasTamilNadu && !hasHolidayHint) {
    return false;
  }
  return true;
}

export function buildFormVITamilNaduSheetTextBlob(workbook, sheetName) {
  if (!workbook?.Sheets || !sheetName || !workbook.Sheets[sheetName]) return '';
  const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: '' });
  return rows
    .slice(0, 45)
    .map((row) => (Array.isArray(row) ? row : []).map((c) => String(c || '').trim()).join(' '))
    .join(' ');
}

function sheetLooksLikeFormVITamilNaduHolidayWorkbook(workbook, sheetName) {
  const sheetText = normalizeFormVITamilNaduText(buildFormVITamilNaduSheetTextBlob(workbook, sheetName));
  if (!sheetText) return false;
  if (FORM_VI_TN_HOLIDAY_TITLE_REGEX.test(sheetText)) return true;

  const holidayHits = FORM_VI_TN_HOLIDAY_KEYWORDS.reduce((acc, keyword) => {
    if (sheetText.includes(keyword)) return acc + 1;
    return acc;
  }, 0);
  const holidayDateMatches = sheetText.match(/\b\d{1,2}[/-]\d{1,2}[/-]\d{2,4}\b/g) || [];
  const hasEmployeeCode = /employee\s+code/.test(sheetText);
  const hasRemarks = /remarks?/.test(sheetText);
  if (hasEmployeeCode && hasRemarks && holidayHits >= 1 && holidayDateMatches.length >= 2) {
    return true;
  }
  if (holidayHits >= 2 && holidayDateMatches.length >= 3 && hasRemarks) {
    return true;
  }
  return false;
}

export function scoreFormVITamilNaduWorkbookSheet(workbook, sheetName, hints = {}) {
  const sheetText = normalizeFormVITamilNaduText(buildFormVITamilNaduSheetTextBlob(workbook, sheetName));
  const sheetNameText = normalizeFormVITamilNaduText(sheetName);
  const wantsFormVI = isFormVITamilNaduHolidayContext(
    hints?.formHeader || {
      title: hints?.formHeaderTitle,
      subtitle: hints?.formHeaderSubtitle,
      reference: hints?.formHeaderReference
    },
    hints?.item,
    hints?.fileName || hints?.formFileName || '',
    hints?.sheetText || '',
    hints?.tableHeaders || []
  );
  if (!wantsFormVI) return -Infinity;

  let score = 0;
  if (sheetLooksLikeFormVITamilNaduHolidayWorkbook(workbook, sheetName)) score += 420;
  if (FORM_VI_HINT_REGEX.test(sheetNameText)) score += 30;
  if (/form\s*vi/.test(sheetText) || /form\s*6/.test(sheetText)) score += 20;
  if (/national\s+and\s+festival\s+holidays?/.test(sheetText)) score += 220;
  if (/festival\s+holidays?/.test(sheetText)) score += 140;
  if (/approved\s+festival\s+holidays?/.test(sheetText)) score += 90;

  const holidayHits = FORM_VI_TN_HOLIDAY_KEYWORDS.reduce((acc, keyword) => {
    if (sheetText.includes(keyword)) return acc + 1;
    return acc;
  }, 0);
  score += Math.min(holidayHits * 18, 144);

  const holidayDateMatches = sheetText.match(/\b\d{1,2}[/-]\d{1,2}[/-]\d{2,4}\b/g) || [];
  score += Math.min(holidayDateMatches.length * 8, 96);

  const employeeDetailHits = FORM_VI_TN_EMPLOYEE_DETAIL_KEYWORDS.reduce((acc, keyword) => {
    if (sheetText.includes(keyword)) return acc + 1;
    return acc;
  }, 0);
  if (
    /insurance\s+corporation|aadhaar|480\s+days|made\s+permanent|subsistence\s+allowance/.test(sheetText) ||
    employeeDetailHits >= 2
  ) {
    score -= 240;
  }
  if (
    /employee\s+details|permanent\s+status|date\s+of\s+joining|date\s+of\s+suspension/.test(sheetText) ||
    /employee\s+identification\s+no|gender|date\s+of\s+birth/.test(sheetText)
  ) {
    score -= 80;
  }
  if (/s\.?\s*no|serial/.test(sheetText)) score += 8;
  if (/employee\s+code/.test(sheetText)) score += 25;
  if (/name\s+of\s+the\s+employee|father|husband|doj/.test(sheetText)) score += 16;
  if (/remarks?/.test(sheetText)) score += 12;

  return score;
}

export function resolveFormVITamilNaduWorkbookSheetName(workbook, hints = {}) {
  const names = Array.isArray(workbook?.SheetNames) ? workbook.SheetNames : [];
  if (names.length === 0) return '';
  const wantsFormVI = isFormVITamilNaduHolidayContext(
    hints?.formHeader || {
      title: hints?.formHeaderTitle,
      subtitle: hints?.formHeaderSubtitle,
      reference: hints?.formHeaderReference
    },
    hints?.item,
    hints?.fileName || hints?.formFileName || '',
    hints?.sheetText || '',
    hints?.tableHeaders || []
  );
  if (!wantsFormVI) {
    return hints?.preferredSheetName && names.includes(hints.preferredSheetName)
      ? hints.preferredSheetName
      : names[0] || '';
  }

  const explicitHolidaySheet = names.find((name) => sheetLooksLikeFormVITamilNaduHolidayWorkbook(workbook, name));
  if (explicitHolidaySheet) return explicitHolidaySheet;

  const preferred = hints?.preferredSheetName && names.includes(hints.preferredSheetName)
    ? hints.preferredSheetName
    : names[0];
  let best = preferred || names[0] || '';
  let bestScore = best ? scoreFormVITamilNaduWorkbookSheet(workbook, best, hints) : -Infinity;

  for (const name of names) {
    const score = scoreFormVITamilNaduWorkbookSheet(workbook, name, hints);
    if (score > bestScore) {
      bestScore = score;
      best = name;
    }
  }

  return bestScore > 0 ? best : (preferred || names[0] || '');
}

export function repickFormVITamilNaduWorkbookSheetIfNeeded(workbook, hints = {}, currentSheetName = '') {
  if (
    !isFormVITamilNaduHolidayContext(
      hints?.formHeader || {
        title: hints?.formHeaderTitle,
        subtitle: hints?.formHeaderSubtitle,
        reference: hints?.formHeaderReference
      },
      hints?.item,
      hints?.fileName || hints?.formFileName || '',
      hints?.sheetText || '',
      hints?.tableHeaders || []
    )
  ) {
    return null;
  }
  const target = resolveFormVITamilNaduWorkbookSheetName(workbook, {
    ...hints,
    preferredSheetName: currentSheetName || hints?.preferredSheetName || ''
  });
  if (target && target !== currentSheetName) return target;
  return null;
}

/** Form VI TN: merged heading often stores FORM / REGISTER / rule / Act lines in one cell. */
export function splitFormVIHeaderBlock(text) {
  const t = String(text || '')
    .replace(/\r\n/g, '\n')
    .replace(/\u000d/g, '\n')
    .replace(/_x000d_/gi, '\n')
    .replace(/\s+\n/g, '\n')
    .replace(/\n\s+/g, '\n')
    .trim();
  if (!t) return null;
  if (!/\bform\s*[-–]?\s*vi\b/i.test(t) && !/register\s+of\s+national\s+and\s+festival\s+holidays/i.test(t)) {
    return null;
  }

  const lines = t
    .split('\n')
    .map((l) => l.replace(/\s+/g, ' ').trim())
    .filter(Boolean);
  if (lines.length === 0) return null;

  const titleLine = lines.find((l) => /\bform\s*[-–]?\s*vi\b/i.test(l)) || 'FORM - VI';
  const subtitleLine =
    lines.find((l) => /register\s+of\s+national\s+and\s+festival\s+holidays/i.test(l)) ||
    'REGISTER OF NATIONAL AND FESTIVAL HOLIDAYS';
  const bracketRule =
    lines.find((l) => /^\[.*\]$/.test(l) || /see\s+sub-?rule|see\s+rule/i.test(l)) || '';
  const actLine =
    lines.find((l) => /tamil\s+nadu\s+industrial\s+establishments/i.test(l)) ||
    lines.find((l) => /national\s+and\s+festival\s+holidays?\s+rules?/i.test(l)) ||
    '';
  const reference = [bracketRule, actLine].filter(Boolean).join('\n').trim();

  return {
    title: titleLine,
    subtitle: subtitleLine,
    reference
  };
}

function formVIExcelJsCellText(val) {
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

/** Normalize SheetJS / Excel repair artifacts that trigger "We found a problem". */
export function normalizeFormVITitleCellText(raw) {
  return String(raw || '')
    .replace(/_x000d_/gi, '\n')
    .replace(/\u000d/g, '\n')
    .replace(/\r\n/g, '\n')
    .replace(/\r/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

function looksLikeFormVITitleBandText(raw) {
  const s = normalizeFormVITitleCellText(raw).replace(/\s+/g, ' ').trim().toLowerCase();
  if (!s) return false;
  // Festival date-band group header uses similar legal wording — never treat it as the form title.
  if (/days,?\s+dates\s+and\s+months/.test(s)) return false;
  if (/under\s+(the\s+)?section\s*3/.test(s)) return false;
  if (/enter\s+days\s+dates?/.test(s)) return false;
  if (/\bform\s*[-–]?\s*vi\b/.test(s)) return true;
  if (/register\s+of\s+national\s+and\s+festival\s+holidays/.test(s)) return true;
  if (/see\s+sub-?rule/.test(s) && /rule\s*\(?\s*7/.test(s)) return true;
  if (
    /tamil\s+nadu\s+industrial\s+establishments/.test(s) &&
    /festival\s+holidays?\s+rules?/.test(s)
  ) {
    return true;
  }
  return false;
}

function decodeExcelA1Col(letters) {
  let n = 0;
  const s = String(letters || '').toUpperCase();
  for (let i = 0; i < s.length; i += 1) n = n * 26 + (s.charCodeAt(i) - 64);
  return n;
}

function parseExcelMergeLabel(label) {
  const m = String(label || '').match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/i);
  if (!m) return null;
  return {
    c1: decodeExcelA1Col(m[1]),
    r1: Number(m[2]),
    c2: decodeExcelA1Col(m[3]),
    r2: Number(m[4]),
    label: String(label)
  };
}

/**
 * Pick the Form VI holiday register worksheet from an ExcelJS workbook.
 */
export function pickFormVITamilNaduExcelJsWorksheet(workbook, hints = {}) {
  const sheets = Array.isArray(workbook?.worksheets) ? workbook.worksheets : [];
  if (sheets.length === 0) return null;

  const preferred = String(hints?.sheetNameHint || hints?.preferredSheetName || '')
    .trim()
    .toLowerCase();
  if (preferred) {
    const exact = sheets.find((ws) => String(ws?.name || '').trim().toLowerCase() === preferred);
    if (exact) return exact;
  }

  const scoreSheet = (ws) => {
    const n = String(ws?.name || '').trim().toLowerCase();
    let score = 0;
    if (n === 'form vi' || n === 'form 6' || n === 'form-vi') score += 80;
    if (/\bform\s*vi\b/.test(n) || /\bform\s*6\b/.test(n)) score += 40;
    if (/festival|holiday|national/.test(n)) score += 30;
    if (/employee\s+detail/.test(n)) score -= 120;
    if (/^sheet\d+$/i.test(n)) score -= 20;

    const rowLimit = Math.min(20, ws.rowCount || 20);
    const colLimit = Math.min(30, ws.columnCount || 30);
    let blob = '';
    for (let r = 1; r <= rowLimit; r += 1) {
      for (let c = 1; c <= colLimit; c += 1) {
        const t = formVIExcelJsCellText(ws.getCell(r, c)?.value).trim();
        if (t) blob += ` ${t}`;
      }
    }
    const text = blob.toLowerCase();
    if (/register\s+of\s+national\s+and\s+festival\s+holidays/.test(text)) score += 220;
    if (/festival\s+holidays?/.test(text)) score += 80;
    if (FORM_VI_TN_HOLIDAY_KEYWORDS.some((k) => text.includes(k))) score += 40;
    if (/employee\s+identification|date\s+of\s+birth|aadhaar/.test(text)) score -= 80;
    return score;
  };

  let best = sheets[0];
  let bestScore = scoreSheet(best);
  sheets.forEach((ws) => {
    const score = scoreSheet(ws);
    if (score > bestScore) {
      bestScore = score;
      best = ws;
    }
  });
  return best;
}

/**
 * Restore Form VI title band as a centered multi-row merge across the table width
 * (template: A1:S4). When merges are lost, the heading collapses into narrow column A.
 */
export function ensureFormVITamilNaduTitleLayout(worksheet, options = {}) {
  if (!worksheet) return false;

  const colFrom = Math.max(1, Number(options.colFrom) || 1);
  const colTo = Math.max(colFrom, Number(options.colTo) || colFrom + 18);
  const headerRow = Math.max(2, Number(options.headerRow) || 6);
  const scanRows = Math.min(Math.max(1, headerRow - 1), 8);

  let titleRow = -1;
  let titleCol = -1;
  let titleText = '';
  for (let r = 1; r <= scanRows; r += 1) {
    for (let c = 1; c <= Math.min(colTo + 2, 40); c += 1) {
      const raw = formVIExcelJsCellText(worksheet.getCell(r, c)?.value);
      if (!looksLikeFormVITitleBandText(raw)) continue;
      titleRow = r;
      titleCol = c;
      titleText = normalizeFormVITitleCellText(raw);
      break;
    }
    if (titleRow > 0) break;
  }
  if (titleRow < 1 || !titleText) return false;

  const mergeLabels = Array.isArray(worksheet.model?.merges) ? [...worksheet.model.merges] : [];
  const titleMerges = mergeLabels
    .map(parseExcelMergeLabel)
    .filter((m) => m && titleRow >= m.r1 && titleRow <= m.r2 && titleCol >= m.c1 && titleCol <= m.c2);

  const wideEnough = titleMerges.some((m) => m.c2 - m.c1 >= Math.min(10, colTo - colFrom));
  const multiRow = titleMerges.some((m) => m.r2 > m.r1);
  const needsRebuild = !wideEnough || !multiRow || titleCol !== colFrom;

  // Always normalize CR artifacts that cause Excel repair dialogs.
  const titleCell = worksheet.getCell(titleRow, titleCol);
  if (String(titleCell.value || '') !== titleText) {
    titleCell.value = titleText;
  }

  if (!needsRebuild) {
    const master = worksheet.getCell(titleMerges[0].r1, titleMerges[0].c1);
    master.value = titleText;
    master.alignment = {
      ...(master.alignment || {}),
      horizontal: 'center',
      vertical: 'middle',
      wrapText: true
    };
    if (!master.font?.bold) {
      master.font = { ...(master.font || {}), bold: true };
    }
    return true;
  }

  // Prefer template span: title row through row above factory-address / header (usually rows 1–4).
  let titleEndRow = Math.max(titleRow, Math.min(scanRows, 4));
  if (titleMerges.length > 0) {
    titleEndRow = Math.max(titleEndRow, ...titleMerges.map((m) => m.r2));
  }
  titleEndRow = Math.min(titleEndRow, headerRow - 1);

  const unmergeCovering = (r1, c1, r2, c2) => {
    mergeLabels.forEach((label) => {
      const m = parseExcelMergeLabel(label);
      if (!m) return;
      const overlaps = !(m.r2 < r1 || m.r1 > r2 || m.c2 < c1 || m.c1 > c2);
      if (!overlaps) return;
      try {
        worksheet.unMergeCells(label);
      } catch (_) {
        /* ignore */
      }
    });
  };

  unmergeCovering(titleRow, colFrom, titleEndRow, colTo);
  try {
    worksheet.mergeCells(titleRow, colFrom, titleEndRow, colTo);
  } catch (_) {
    try {
      worksheet.mergeCells(titleRow, colFrom, titleRow, colTo);
      titleEndRow = titleRow;
    } catch (__) {
      return false;
    }
  }

  const cell = worksheet.getCell(titleRow, colFrom);
  cell.value = titleText;
  if (titleCol !== colFrom) {
    try {
      worksheet.getCell(titleRow, titleCol).value = null;
    } catch (_) {
      /* ignore */
    }
  }
  cell.alignment = {
    ...(cell.alignment || {}),
    horizontal: 'center',
    vertical: 'middle',
    wrapText: true
  };
  cell.font = {
    ...(cell.font || {}),
    bold: true,
    size: cell.font?.size || 12,
    name: cell.font?.name || 'Palatino Linotype'
  };

  const lineCount = Math.max(1, titleText.split(/\n/).filter((l) => String(l).trim()).length);
  const totalHeight = Math.min(120, Math.max(72, lineCount * 18));
  const perRow = Math.round(totalHeight / Math.max(1, titleEndRow - titleRow + 1));
  for (let r = titleRow; r <= titleEndRow; r += 1) {
    const row = worksheet.getRow(r);
    if (!row.height || row.height < perRow) row.height = perRow;
  }

  return true;
}

/** Canonical Form VI (Tamil Nadu) layout from Form_VI_-_TamilNadu.xlsx. */
export const FORM_VI_TN_TABLE_COLS = 19;
export const FORM_VI_TN_COL_WIDTHS = [
  9.18, 9.82, 31.54, 26.27, 14.27, 9.18, 9.18, 9.18, 9.18, 9.18, 9.18, 9.18, 9.18, 9.18, 9.18,
  9.18, 9.18, 9.18, 9.18
];
export const FORM_VI_TN_MERGES = [
  'A1:S4',
  'A5:D5',
  'E5:S5',
  'A6:A8',
  'B6:B8',
  'C6:C8',
  'D6:D8',
  'F6:S6',
  'S7:S8'
];
export const FORM_VI_TN_ROW_HEIGHTS = {
  1: 18,
  2: 18,
  5: 69,
  6: 35.15,
  7: 242
};

function cloneFormVICellStyle(style) {
  if (!style) return null;
  try {
    const cloned = JSON.parse(JSON.stringify(style));
    // Indexed theme colors often trip Excel repair after ExcelJS rewrite.
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

function formVILooksLikeSerialHeader(txt) {
  const s = String(txt || '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
  return /^(s\.?\s*no\.?|sl\.?\s*no\.?|serial(\s*no\.?)?)$/i.test(s) || s === 's.no' || s === 'sno';
}

function formVINormalizeHeader(txt) {
  return String(txt || '')
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

function formVIHolidayMatchToken(txt) {
  const s = formVINormalizeHeader(txt);
  const tokens = [
    'pongal',
    'republic day',
    'tamil new year',
    'good friday',
    'may day',
    'ramzan',
    'independence day',
    'krishna jayanthi',
    'vinayakar chathurthi',
    'ayudha pooja',
    'vijaya dashami',
    'diwali',
    'christmas'
  ];
  const hit = tokens.find((t) => s.includes(t));
  if (hit) return hit;
  const date = s.match(/\d{1,2}[-/.]\d{1,2}[-/.]\d{2,4}/);
  return date ? date[0].replace(/\//g, '-') : s;
}

function formVIHolidayLabelsMatch(appHeader, excelLabel) {
  const a = formVIHolidayMatchToken(appHeader);
  const e = formVIHolidayMatchToken(excelLabel);
  if (!a || !e) return false;
  if (a === e) return true;
  return a.includes(e) || e.includes(a);
}

function isFormVIHolidayColumnHeaderLocal(header) {
  const s = formVINormalizeHeader(header);
  if (!s || s === 'remark' || s === 'remarks') return false;
  if (/days,?\s+dates\s+and\s+months/.test(s)) return false;
  if (/national/.test(s) && /festival/.test(s) && /holiday/.test(s)) return false;
  if (/\d{1,2}[-/.]\d{1,2}[-/.]\d{2,4}/.test(s)) return true;
  return /pongal|republic\s+day|tamil\s+new\s+year|good\s+friday|may\s+day|ramzan|independence\s+day|krishna|vinayakar|ayudha|vijaya\s+dashami|diwali|christmas/.test(
    s
  );
}

function isFormVIDojHeaderLocal(header) {
  const s = formVINormalizeHeader(header);
  return s === 'doj' || s === 'd.o.j' || s === 'd.o.j.' || (s.includes('date') && s.includes('join'));
}

/**
 * Clone Form VI into a fresh workbook limited to columns A–S.
 * The stock template used-range (A1:MV19) plus phantom col groups make Excel
 * repair downloads and drop merges/widths — that is what collapses the title into col A.
 */
export function cloneFormVITamilNaduWorksheetClean(sourceWs) {
  // eslint-disable-next-line global-require
  const ExcelJS = require('exceljs');
  const outWb = new ExcelJS.Workbook();
  const outWs = outWb.addWorksheet('Form VI');
  const maxCol = FORM_VI_TN_TABLE_COLS;
  let maxRow = Math.max(19, sourceWs.actualRowCount || 0, sourceWs.rowCount || 0);
  maxRow = Math.min(Math.max(maxRow, 19), 200);

  for (let r = 1; r <= maxRow; r += 1) {
    const srcRow = sourceWs.getRow(r);
    const dstRow = outWs.getRow(r);
    if (srcRow.height != null) dstRow.height = srcRow.height;
    for (let c = 1; c <= maxCol; c += 1) {
      const src = sourceWs.getCell(r, c);
      const dst = outWs.getCell(r, c);
      const raw = src.value;
      if (typeof raw === 'string') {
        dst.value = normalizeFormVITitleCellText(raw);
      } else if (raw != null) {
        dst.value = raw;
      }
      const cloned = cloneFormVICellStyle(src.style);
      if (cloned) dst.style = cloned;
    }
  }

  FORM_VI_TN_COL_WIDTHS.forEach((w, i) => {
    outWs.getColumn(i + 1).width = w;
  });

  Object.entries(FORM_VI_TN_ROW_HEIGHTS).forEach(([row, height]) => {
    const r = Number(row);
    if (!outWs.getRow(r).height) outWs.getRow(r).height = height;
  });

  FORM_VI_TN_MERGES.forEach((label) => {
    try {
      outWs.mergeCells(label);
    } catch (_) {
      /* ignore duplicate */
    }
  });

  const titleCell = outWs.getCell(1, 1);
  if (titleCell.value != null) {
    titleCell.value = normalizeFormVITitleCellText(formVIExcelJsCellText(titleCell.value));
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

  return { workbook: outWb, worksheet: outWs };
}

/**
 * Build a Form VI Tamil Nadu download workbook with template alignment (no Excel repair).
 */
export async function buildFormVITamilNaduWorkbookWithTemplateStyles({
  templateArrayBuffer,
  mappedData,
  mappedRowMatrix,
  headersToUse,
  parsedHeaderRowIndex,
  parsedDataStartIndex,
  parsedFormHeader,
  formFileName,
  sheetNameHint,
  ExcelJS: ExcelJSLib
} = {}) {
  const ExcelJS = ExcelJSLib || require('exceljs');
  const templateWb = new ExcelJS.Workbook();
  await templateWb.xlsx.load(templateArrayBuffer);

  const sourceWs = pickFormVITamilNaduExcelJsWorksheet(templateWb, {
    sheetNameHint,
    preferredSheetName: sheetNameHint,
    formHeader: parsedFormHeader,
    fileName: formFileName
  });
  if (!sourceWs) throw new Error('Template worksheet not found.');

  const { workbook, worksheet } = cloneFormVITamilNaduWorksheetClean(sourceWs);

  let headerRow =
    parsedHeaderRowIndex != null && parsedHeaderRowIndex >= 0 ? parsedHeaderRowIndex + 1 : -1;
  let startCol = 1;
  for (let r = 1; r <= 40; r += 1) {
    for (let c = 1; c <= FORM_VI_TN_TABLE_COLS; c += 1) {
      const raw = formVIExcelJsCellText(worksheet.getCell(r, c)?.value);
      if (formVILooksLikeSerialHeader(raw)) {
        headerRow = r;
        startCol = c;
        break;
      }
    }
    if (headerRow > 0) break;
  }
  if (headerRow < 1) headerRow = 6;

  const getColLabel = (col) => {
    const parts = [];
    for (let r = Math.max(1, headerRow - 2); r <= headerRow + 2; r += 1) {
      const v = formVINormalizeHeader(formVIExcelJsCellText(worksheet.getCell(r, col)?.value));
      if (v) parts.push(v);
    }
    return parts.join(' ');
  };

  let dataStartRow =
    parsedDataStartIndex != null && parsedDataStartIndex >= 0 ? parsedDataStartIndex + 1 : headerRow + 1;
  let seqMarkers = 0;
  for (let c = startCol; c <= startCol + 18; c += 1) {
    const v = formVINormalizeHeader(formVIExcelJsCellText(worksheet.getCell(headerRow + 1, c)?.value));
    if (/^\d{1,2}$/.test(v)) seqMarkers += 1;
  }
  if (seqMarkers >= 4) dataStartRow = Math.max(dataStartRow, headerRow + 2);
  if (dataStartRow < 9) dataStartRow = 9;

  const hdrs = Array.isArray(headersToUse) ? headersToUse : [];
  const colByHeader = new Map();
  hdrs.forEach((header) => {
    const hNorm = formVINormalizeHeader(header);
    for (let c = startCol; c <= FORM_VI_TN_TABLE_COLS; c += 1) {
      const label = getColLabel(c);
      if (!label) continue;
      let matched = false;
      if (/\bs\.?\s*no\b|serial/.test(hNorm) && /\bs\.?\s*no\b|serial/.test(label)) matched = true;
      else if (/employee\s+code/.test(hNorm) && /employee\s+code/.test(label)) matched = true;
      else if (/name\s+of\s+the\s+employee/.test(hNorm) && /name\s+of\s+the\s+employee/.test(label)) matched = true;
      else if (/father|husband/.test(hNorm) && /father|husband/.test(label)) matched = true;
      else if (isFormVIDojHeaderLocal(header) && (/\bdoj\b/.test(label) || /date\s+of\s+joining/.test(label))) matched = true;
      else if (/^remarks?$/.test(hNorm) && /^remarks?$/.test(label)) matched = true;
      else if (isFormVIHolidayColumnHeaderLocal(header) && formVIHolidayLabelsMatch(header, label)) matched = true;
      if (matched && !colByHeader.has(header)) colByHeader.set(header, c);
    }
  });

  let dojCol = null;
  hdrs.forEach((h) => {
    if (isFormVIDojHeaderLocal(h) && colByHeader.has(h)) dojCol = colByHeader.get(h);
  });
  if (!dojCol) {
    for (let c = startCol; c <= FORM_VI_TN_TABLE_COLS; c += 1) {
      const label = getColLabel(c);
      if (/\bdoj\b/.test(label) || /date\s+of\s+joining/.test(label)) {
        dojCol = c;
        break;
      }
    }
  }

  const holidayHeaders = hdrs.filter((h) => isFormVIHolidayColumnHeaderLocal(h));
  if (dojCol && holidayHeaders.some((h) => !colByHeader.has(h))) {
    let col = dojCol + 1;
    holidayHeaders.forEach((h) => {
      if (colByHeader.has(h)) return;
      while (col <= FORM_VI_TN_TABLE_COLS && /^remarks?$/.test(getColLabel(col))) col += 1;
      if (col <= FORM_VI_TN_TABLE_COLS) {
        colByHeader.set(h, col);
        col += 1;
      }
    });
  }

  hdrs.forEach((h, idx) => {
    if (!colByHeader.has(h)) {
      const col = startCol + idx;
      if (col <= FORM_VI_TN_TABLE_COLS) colByHeader.set(h, col);
    }
  });

  const rowLooksMeaningful = (row) => {
    if (Array.isArray(row)) return row.some((v) => String(v ?? '').trim() !== '');
    if (row && typeof row === 'object') return Object.values(row).some((v) => String(v ?? '').trim() !== '');
    return false;
  };

  const sourcePrimary =
    Array.isArray(mappedData) && mappedData.length > 0
      ? mappedData
      : Array.isArray(mappedRowMatrix)
        ? mappedRowMatrix
        : [];

  const sourceRows = sourcePrimary
    .filter((row) => rowLooksMeaningful(row))
    .map((row, idx) => {
      if (Array.isArray(row)) return row;
      const obj = {};
      hdrs.forEach((h) => {
        obj[h] = row?.[h] ?? '';
      });
      const snoHeader = hdrs.find((h) => /\bs\.?\s*no\b|serial/i.test(String(h || '')));
      if (snoHeader && !String(obj[snoHeader] ?? '').trim()) obj[snoHeader] = idx + 1;
      holidayHeaders.forEach((h) => {
        if (!String(obj[h] ?? '').trim()) obj[h] = 'H';
      });
      return obj;
    });

  const footerScanStart = dataStartRow;
  let footerRow = -1;
  for (
    let r = Math.min(worksheet.rowCount || 40, dataStartRow + Math.max(sourceRows.length, 5) + 5);
    r >= footerScanStart;
    r -= 1
  ) {
    const blob = Array.from({ length: FORM_VI_TN_TABLE_COLS }, (_, i) =>
      formVIExcelJsCellText(worksheet.getCell(r, i + 1)?.value)
    )
      .join(' ')
      .toLowerCase();
    if (/authorised|authorized|signatory|signature\s+of\s+employer/.test(blob)) {
      footerRow = r;
    }
  }
  const clearTo =
    footerRow > dataStartRow
      ? footerRow - 1
      : Math.min(Math.max(dataStartRow + sourceRows.length + 5, dataStartRow + 10), 200);

  for (let r = dataStartRow; r <= clearTo; r += 1) {
    for (let c = 1; c <= FORM_VI_TN_TABLE_COLS; c += 1) {
      worksheet.getCell(r, c).value = null;
    }
  }

  const border = {
    top: { style: 'thin', color: { argb: 'FF000000' } },
    left: { style: 'thin', color: { argb: 'FF000000' } },
    bottom: { style: 'thin', color: { argb: 'FF000000' } },
    right: { style: 'thin', color: { argb: 'FF000000' } }
  };

  for (let i = 0; i < sourceRows.length; i += 1) {
    const row = sourceRows[i] || {};
    const excelRow = dataStartRow + i;
    hdrs.forEach((header) => {
      const col = colByHeader.get(header);
      if (!col) return;
      const value = Array.isArray(row) ? row[hdrs.indexOf(header)] : row[header];
      if (value == null || String(value).trim() === '') return;
      const cell = worksheet.getCell(excelRow, col);
      if (
        /\bs\.?\s*no\b|serial/i.test(String(header || '')) &&
        (typeof value === 'number' ||
          (typeof value === 'string' && /^-?\d+(\.\d+)?$/.test(String(value).trim())))
      ) {
        cell.value = Number(value);
      } else {
        cell.value = String(value);
      }
    });
    for (let c = 1; c <= FORM_VI_TN_TABLE_COLS; c += 1) {
      const cell = worksheet.getCell(excelRow, c);
      cell.border = border;
      cell.alignment = {
        ...(cell.alignment || {}),
        horizontal: 'center',
        vertical: 'middle',
        wrapText: true
      };
    }
  }

  FORM_VI_TN_COL_WIDTHS.forEach((w, i) => {
    worksheet.getColumn(i + 1).width = w;
  });
  if (!worksheet.getRow(7).height || worksheet.getRow(7).height < 200) {
    worksheet.getRow(7).height = FORM_VI_TN_ROW_HEIGHTS[7];
  }
  if (!worksheet.getRow(5).height) worksheet.getRow(5).height = FORM_VI_TN_ROW_HEIGHTS[5];
  if (!worksheet.getRow(6).height) worksheet.getRow(6).height = FORM_VI_TN_ROW_HEIGHTS[6];

  ensureFormVITamilNaduTitleLayout(worksheet, {
    colFrom: 1,
    colTo: FORM_VI_TN_TABLE_COLS,
    headerRow
  });

  const out = await workbook.xlsx.writeBuffer();
  const fileName =
    formFileName ||
    parsedFormHeader?.title?.replace(/[^a-zA-Z0-9]/g, '_') ||
    `Form_VI_${Date.now()}.xlsx`;
  const blob = new Blob([out], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  });
  return { blob, fileName, arrayBuffer: out };
}
