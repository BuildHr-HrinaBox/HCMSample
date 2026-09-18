import ExcelJS from 'exceljs';
import {
  applyExcelJSFullBoxBordersToRange,
  ensureExcelJSDataRowsWithBorders,
} from '../../utils/excelTableBorders';
import {
  applyFormCRJContractorFromSite,
  enrichEstablishmentPrincipalEmployerHeaderFields,
  writeStatutoryHeaderFieldsToExcelJsWorksheet,
} from '../../utils/statutorySiteCompanyHeaders';
import {
  formAGJGujaratHeaderNorm,
  formAGJHeaderAliasBucket,
  getFormAGJGujaratRowValueForHeader,
  mapFormAGJGujaratRowsFromEmployees,
} from './formAGJGujarat';

/** Rajasthan Form A — employee register (original Form A_RJ Excel template). */

/** Official Part-A columns on the blank Form A RJ template (image / Form Master file). */
export const FORM_A_RJ_CANONICAL_TABLE_HEADERS = [
  'Sl. No.',
  'Employee Code',
  'Name',
  'Surname',
  "Father's/Spouse Name",
  'Date of Birth',
  'Nationality',
  'Education Level',
  'Date of Joining',
  'Designation',
  'Category (HS/S/SS/US)*',
  'Type of Employment',
  'Mobile',
  'UAN',
  'PAN',
];

/** Bucket order → physical columns A–O on the official Form A RJ register. */
export const FORM_A_RJ_BUCKET_COL_ORDER = [
  'sno',
  'employeeId',
  'firstName',
  'surname',
  'father',
  'dob',
  'nationality',
  'education',
  'doj',
  'designation',
  'skillCategory',
  'employmentType',
  'mobile',
  'uan',
  'pan',
];

const FORM_A_RJ_TITLE_COL_FROM = 5; // E
const FORM_A_RJ_TITLE_COL_TO = 7; // G
const FORM_A_RJ_DEFAULT_MAIN_TITLE = 'FORM A';
const FORM_A_RJ_DEFAULT_SUBTITLE = 'FORMAT OF EMPLOYEE REGISTER';
const FORM_A_RJ_SEE_RULE_TEXT = '[See rule 2(1)]';
const FORM_A_RJ_PART_A_TEXT = '[Part-A: For all Establishments]';
/** Employee data rows need extra height so wrapped DOB / designation stay readable. */
const FORM_A_RJ_DATA_ROW_HEIGHT = 32;
const FORM_A_RJ_TABLE_START_COL = 1; // A

/** Narrow Employee Code; Name immediately after (column C). */
const FORM_A_RJ_COLUMN_WIDTHS = {
  1: 6, // Sl. No.
  2: 12, // Employee Code
  3: 16, // Name
  4: 14, // Surname
  5: 22,
  6: 12,
  7: 12,
  8: 16,
  9: 12,
  10: 16,
  11: 14,
  12: 16,
  13: 12,
  14: 14,
  15: 12,
};

function isPlaceholderBandText(text) {
  const raw = String(text || '').trim();
  if (!raw) return true;
  if (/^[-–—_.\s]+$/.test(raw)) return true;
  return false;
}

function sheetTextLooksLikeFormARJMainTitle(text) {
  const n = formAGJGujaratHeaderNorm(text).replace(/[^a-z0-9 ]/g, ' ').replace(/\s+/g, ' ').trim();
  return n === 'form a' || /^form\s*a$/.test(n);
}

function sheetTextLooksLikeFormARJSubtitle(text) {
  if (isPlaceholderBandText(text)) return false;
  const n = formAGJGujaratHeaderNorm(text);
  return /format\s+of\s+employee/.test(n);
}

function sheetTextLooksLikeFormARJSeeRule(text) {
  if (isPlaceholderBandText(text)) return false;
  const n = formAGJGujaratHeaderNorm(text);
  return /see\s*rule\s*2/.test(n);
}

function sheetTextLooksLikeFormARJPartA(text) {
  if (isPlaceholderBandText(text)) return false;
  const n = formAGJGujaratHeaderNorm(text);
  return /part\s*[-\s]?a\b/.test(n) && /establish/.test(n);
}

/** Move See-rule / FORM A / subtitle / Part-A into columns E–G (Part-A below subtitle). */
export function writeFormARajasthanTitleBandsInColumnsEG(worksheet, parsedFormHeader = {}) {
  if (!worksheet) return;
  const getMergeTopLeft = buildMergeTopLeftResolver(worksheet);
  const getMergedAwareCellText = (r, c) => {
    const tl = getMergeTopLeft(r, c);
    return excelCellValueToString(worksheet.getCell(tl.r, tl.c)?.value).trim();
  };

  const parsedTitle = String(parsedFormHeader?.title || '').trim();
  const parsedSubtitle = String(parsedFormHeader?.subtitle || '').trim();
  const mainText =
    parsedTitle && sheetTextLooksLikeFormARJMainTitle(parsedTitle)
      ? parsedTitle
      : FORM_A_RJ_DEFAULT_MAIN_TITLE;
  const subText =
    parsedSubtitle && sheetTextLooksLikeFormARJSubtitle(parsedSubtitle)
      ? parsedSubtitle
      : FORM_A_RJ_DEFAULT_SUBTITLE;

  let seeRuleRow = -1;
  let seeRuleText = FORM_A_RJ_SEE_RULE_TEXT;
  let mainRow = -1;
  let subRow = -1;
  let partARow = -1;
  let partAText = FORM_A_RJ_PART_A_TEXT;
  // Scan far enough for Form A MH extras (ESIC / LWF / Photo / Signature).
  const titleScanColTo = Math.max(FORM_A_RJ_TITLE_COL_TO + 30, 40);
  for (let r = 1; r <= 12; r += 1) {
    for (let c = 1; c <= titleScanColTo; c += 1) {
      const raw = getMergedAwareCellText(r, c);
      if (!raw || isPlaceholderBandText(raw)) continue;
      if (seeRuleRow < 0 && sheetTextLooksLikeFormARJSeeRule(raw)) {
        seeRuleRow = r;
        seeRuleText = raw;
      }
      if (mainRow < 0 && sheetTextLooksLikeFormARJMainTitle(raw)) mainRow = r;
      if (subRow < 0 && sheetTextLooksLikeFormARJSubtitle(raw)) subRow = r;
      if (partARow < 0 && sheetTextLooksLikeFormARJPartA(raw)) {
        partARow = r;
        partAText = raw;
      }
    }
  }
  if (mainRow < 0) mainRow = 2;
  if (subRow < 0) subRow = 3;
  // Always place See-rule citation on the row above FORM A (never leave it in column O).
  if (seeRuleRow < 0 || seeRuleRow >= mainRow) {
    seeRuleRow = Math.max(1, mainRow - 1);
  }
  if (seeRuleRow === mainRow) {
    if (mainRow > 1) {
      seeRuleRow = mainRow - 1;
    } else {
      mainRow = 2;
      if (subRow <= mainRow) subRow = 3;
      seeRuleRow = 1;
    }
  }

  const rowLooksLikeEstablishmentField = (row) => {
    for (let c = 1; c <= 8; c += 1) {
      const raw = getMergedAwareCellText(row, c);
      if (!raw) continue;
      if (sheetTextLooksLikeFormARJPartA(raw)) continue;
      if (sheetTextLooksLikeFormARJSeeRule(raw)) continue;
      if (sheetTextLooksLikeFormARJMainTitle(raw)) continue;
      if (sheetTextLooksLikeFormARJSubtitle(raw)) continue;
      if (/name\s+(and\s+address\s+)?of|establishment|contractor|principal\s+employer/i.test(raw)) {
        return true;
      }
    }
    return false;
  };

  // Part-A always sits directly under FORMAT OF EMPLOYEE REGISTER (own row).
  if (partARow < 0 || partARow <= subRow) {
    const candidate = subRow + 1;
    if (rowLooksLikeEstablishmentField(candidate)) {
      try {
        worksheet.spliceRows(candidate, 0, []);
      } catch (_) {
        /* ignore */
      }
      partARow = candidate;
    } else {
      partARow = candidate;
    }
  }

  const clearMatchingTextInRow = (row, matcher) => {
    for (let c = 1; c <= titleScanColTo; c += 1) {
      const raw = getMergedAwareCellText(row, c);
      if (!raw || !matcher(raw)) continue;
      const tl = getMergeTopLeft(row, c);
      worksheet.getCell(tl.r, tl.c).value = '';
    }
  };

  const clearTitleFromLeft = (row, kind) => {
    // Part-A writes into column A — do not clear the destination band.
    if (kind === 'partA') return;
    for (let c = 1; c < FORM_A_RJ_TITLE_COL_FROM; c += 1) {
      const raw = getMergedAwareCellText(row, c);
      if (!raw) continue;
      const match =
        kind === 'subtitle'
          ? sheetTextLooksLikeFormARJSubtitle(raw)
          : kind === 'seeRule'
            ? sheetTextLooksLikeFormARJSeeRule(raw)
            : sheetTextLooksLikeFormARJMainTitle(raw);
      if (!match) continue;
      const tl = getMergeTopLeft(row, c);
      worksheet.getCell(tl.r, tl.c).value = '';
    }
  };

  // Clear See-rule / Part-A text left in far-right columns across the title band.
  const clearScanTo = Math.max(12, mainRow, subRow, seeRuleRow, partARow);
  for (let r = 1; r <= clearScanTo; r += 1) {
    clearMatchingTextInRow(r, sheetTextLooksLikeFormARJSeeRule);
    clearMatchingTextInRow(r, sheetTextLooksLikeFormARJPartA);
  }

  const writeBand = (row, text, kind) => {
    clearTitleFromLeft(row, kind);
    // Part-A must sit on the left (col A); other title bands stay in E–G.
    const colFrom = kind === 'partA' ? 1 : FORM_A_RJ_TITLE_COL_FROM;
    const colTo = kind === 'partA' ? 4 : FORM_A_RJ_TITLE_COL_TO;
    try {
      worksheet.unMergeCells(row, colFrom, row, colTo);
    } catch (_) {
      /* not merged */
    }
    try {
      worksheet.mergeCells(row, colFrom, row, colTo);
    } catch (_) {
      /* ignore */
    }
    const cell = worksheet.getCell(row, colFrom);
    cell.value = text;
    cell.alignment = {
      horizontal: kind === 'partA' ? 'left' : 'center',
      vertical: 'middle',
      wrapText: true,
    };
  };

  writeBand(seeRuleRow, seeRuleText || FORM_A_RJ_SEE_RULE_TEXT, 'seeRule');
  writeBand(mainRow, mainText, 'main');
  writeBand(subRow, subText, 'subtitle');
  writeBand(partARow, partAText || FORM_A_RJ_PART_A_TEXT, 'partA');
}

/** Last non-empty header column (Form A MH keeps ESIC IP…Signature after PAN). */
export function findFormARegisterHeaderColTo(worksheet, headerRow, startCol = FORM_A_RJ_TABLE_START_COL) {
  if (!worksheet || headerRow < 1) return FORM_A_RJ_BUCKET_COL_ORDER.length;
  const getMergeTopLeft = buildMergeTopLeftResolver(worksheet);
  let last = startCol + FORM_A_RJ_BUCKET_COL_ORDER.length - 1;
  for (let c = startCol; c <= startCol + 40; c += 1) {
    const tl = getMergeTopLeft(headerRow, c);
    const raw = excelCellValueToString(worksheet.getCell(tl.r, tl.c)?.value).trim();
    if (raw) last = c;
  }
  return last;
}

/** Break vertical merges in the employee data band so each row gets its own box. */
export function unmergeFormARegisterDataBand(worksheet, rowFrom, rowTo, colFrom, colTo) {
  if (!worksheet || rowFrom < 1 || rowTo < rowFrom || colTo < colFrom) return;

  const colLettersToNum = (letters) =>
    String(letters || '')
      .toUpperCase()
      .split('')
      .reduce((n, ch) => n * 26 + (ch.charCodeAt(0) - 64), 0);

  const collectRanges = () => {
    const out = [];
    const raw = worksheet.model?.merges;
    if (Array.isArray(raw)) {
      raw.forEach((range) => {
        const m = String(range || '').match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/i);
        if (!m) return;
        out.push({
          ref: range,
          top: Math.min(Number(m[2]), Number(m[4])),
          bottom: Math.max(Number(m[2]), Number(m[4])),
          left: Math.min(colLettersToNum(m[1]), colLettersToNum(m[3])),
          right: Math.max(colLettersToNum(m[1]), colLettersToNum(m[3])),
        });
      });
    }
    const dict = worksheet._merges;
    if (dict && typeof dict === 'object') {
      Object.keys(dict).forEach((key) => {
        const entry = dict[key];
        const top = Number(entry?.top || entry?.model?.top);
        const left = Number(entry?.left || entry?.model?.left);
        const bottom = Number(entry?.bottom || entry?.model?.bottom);
        const right = Number(entry?.right || entry?.model?.right);
        if (![top, left, bottom, right].every((n) => Number.isFinite(n) && n >= 1)) return;
        out.push({ ref: key, top, left, bottom, right });
      });
    }
    return out;
  };

  collectRanges().forEach((m) => {
    const overlapsRow = m.bottom >= rowFrom && m.top <= rowTo;
    const overlapsCol = m.right >= colFrom && m.left <= colTo;
    if (!overlapsRow || !overlapsCol) return;
    if (m.bottom <= m.top) return; // single-row merge — keep
    try {
      if (m.ref && typeof worksheet.unMergeCells === 'function') {
        worksheet.unMergeCells(m.ref);
      }
    } catch (_) {
      /* try numeric */
    }
    try {
      if (typeof worksheet.unMergeCells === 'function') {
        worksheet.unMergeCells(m.top, m.left, m.bottom, m.right);
      }
    } catch (_) {
      /* ignore */
    }
  });
}

function cellTextLooksLikeEmploymentCardBodyLabel(text) {
  const t = formAGJGujaratHeaderNorm(text);
  return (
    /name of the workman/.test(t) ||
    /serial number in the register/.test(t) ||
    /nature of employment/.test(t) ||
    /wage rate with particulars/.test(t) ||
    /wage period/.test(t) ||
    /tenure of employment/.test(t) ||
    /^remarks$/.test(t)
  );
}

export function sheetBodyHasEmploymentCardLabels(worksheet, dataStartRow, footerStartRow = -1) {
  if (!worksheet || dataStartRow < 1) return false;
  const end = footerStartRow > dataStartRow ? footerStartRow - 1 : dataStartRow + 25;
  for (let r = dataStartRow; r <= end; r += 1) {
    for (let c = 1; c <= 4; c += 1) {
      const t = excelCellValueToString(worksheet.getCell(r, c)?.value).trim();
      if (cellTextLooksLikeEmploymentCardBodyLabel(t)) return true;
    }
  }
  return false;
}

/** Rewrite header row + column widths so Employee Code (B) is narrow and Name follows in C. */
export function applyFormARajasthanCanonicalHeadersAndWidths(worksheet, headerRow, startCol = FORM_A_RJ_TABLE_START_COL) {
  if (!worksheet || headerRow < 1) return;
  FORM_A_RJ_CANONICAL_TABLE_HEADERS.forEach((label, i) => {
    const col = startCol + i;
    const cell = worksheet.getCell(headerRow, col);
    cell.value = label;
    cell.alignment = { ...(cell.alignment || {}), vertical: 'middle', wrapText: true, horizontal: 'center' };
    const width = FORM_A_RJ_COLUMN_WIDTHS[col];
    if (width) worksheet.getColumn(col).width = width;
  });
}

export function buildFormARajasthanCanonicalTemplateCols(headerRow, startCol = FORM_A_RJ_TABLE_START_COL) {
  return FORM_A_RJ_BUCKET_COL_ORDER.map((bucket, i) => ({
    col: startCol + i,
    label: FORM_A_RJ_CANONICAL_TABLE_HEADERS[i],
    bucket,
    subRowOffset: 0,
  }));
}

/** Columns that appear on Gujarat Form A but not on the RJ Part-A grid. */
const FORM_A_RJ_HIDDEN_HEADER_BUCKETS = new Set([
  'gender',
  'categoryAddress',
  'bankAccount',
  'bankName',
  'bankBranchIfsc',
  'aadhaar',
  'esic',
  'presentAddress',
  'permanentAddress',
  'lwf',
]);

/** All export fields on one consecutive row per employee (no blank spacer rows). */
const FORM_A_RJ_MAIN_ROW_OFFSET = 0;
const FORM_A_RJ_EXPORT_ROWS_PER_EMPLOYEE = 1;

export function isFormARajasthanHiddenTableHeader(header) {
  const bucket = formAGJHeaderAliasBucket(formAGJGujaratHeaderNorm(header));
  return FORM_A_RJ_HIDDEN_HEADER_BUCKETS.has(bucket);
}

export function headersIndicateFormARajasthanTable(tableHeaders) {
  if (!Array.isArray(tableHeaders) || tableHeaders.length < 6) return false;
  const buckets = tableHeaders.map((h) => formAGJHeaderAliasBucket(formAGJGujaratHeaderNorm(h)));
  const joined = tableHeaders.map((h) => formAGJGujaratHeaderNorm(h)).join(' ');
  const hasSno = buckets.includes('sno');
  const hasName = buckets.includes('firstName') || /\bname\b/.test(joined);
  const hasSurname = buckets.includes('surname');
  const hasEmployeeCode = buckets.includes('employeeId') || (/employee/.test(joined) && /code/.test(joined));
  const hasSkillCategory =
    buckets.includes('skillCategory') || (/category/.test(joined) && /\(hs|hs\s*\/|\/ss\//.test(joined));
  const hasEducation = buckets.includes('education') || /education/.test(joined);
  return hasSno && hasName && hasSurname && (hasEmployeeCode || hasSkillCategory || hasEducation);
}

/** Hybrid Form A RJ templates: Employee Code + Surname but column C still says Nature of work. */
export function headersIndicateFormARajasthanHybridTemplate(tableHeaders) {
  if (!Array.isArray(tableHeaders) || tableHeaders.length < 5) return false;
  const joined = tableHeaders.map((h) => formAGJGujaratHeaderNorm(h)).join(' ');
  const buckets = tableHeaders.map((h) => formAGJHeaderAliasBucket(formAGJGujaratHeaderNorm(h)));
  const hasSno =
    buckets.includes('sno') || /^sl\.?\s*n|^sr\.?\s*no|^s\.?\s*no|serial\s*no/.test(joined);
  const hasEmployeeCode = buckets.includes('employeeId') || (/employee/.test(joined) && /code/.test(joined));
  const hasSurname = buckets.includes('surname');
  const hybridNatureCol = /nature of work/.test(joined);
  return hasSno && hasEmployeeCode && hasSurname && (hybridNatureCol || buckets.includes('firstName'));
}

export function isFormARajasthanLikeExport(
  formHeader,
  rowItem,
  fileName,
  sheetText = '',
  tableHeaders = []
) {
  if (isFormARajasthanContext(formHeader, rowItem, fileName, sheetText, tableHeaders)) return true;
  if (headersIndicateFormARajasthanHybridTemplate(tableHeaders)) return true;
  const blob = [
    rowItem?.formName,
    rowItem?.FormName,
    rowItem?.state,
    rowItem?.State,
    rowItem?.siteState,
    rowItem?.SiteState,
    fileName,
    formHeader?.title,
    formHeader?.subtitle,
    sheetText,
    Array.isArray(tableHeaders) ? tableHeaders.join(' ') : '',
  ]
    .filter((x) => x != null && String(x).trim() !== '')
    .join(' ')
    .toLowerCase();
  if (/format\s+of\s+employee/.test(blob) && /rajasthan|\brj\b/.test(blob)) return true;
  if (
    /format\s+of\s+employee/.test(blob) &&
    /employee\s+code/.test(blob) &&
    /surname/.test(blob)
  ) {
    return true;
  }
  return false;
}

export function isFormARajasthanContext(
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
    rowItem?.siteState,
    rowItem?.SiteState,
    rowItem?.state,
    rowItem?.State,
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

  if (/gujarat|\b_gj\b|form[\s._-]*a[\s._-]*gj|form_a_gj/.test(parts)) return false;
  if (/maternity\s+benefit|register\s+of\s+muster\s+roll/.test(parts)) return false;

  // Form_A_MH / Form_A_RJ — underscore after A breaks \bform…a\b.
  if (/form[\s._-]*a[\s._-]*(rj|mh)\b|form_a_(rj|mh)\b/.test(parts)) return true;

  const hasRajasthan =
    /rajasthan|\brj\b|form[\s._-]*a[\s._-]*rj|form_a_rj/.test(parts);
  const hasMaharashtra =
    /maharashtra|\b_mh\b|form[\s._-]*a[\s._-]*mh|form_a_mh/.test(parts);
  const hasFormA =
    /form[\s._-]*a(?=[\s._-]|$)/.test(parts) || /\bform[\s._-]*a\b/.test(parts);
  const hasEmployeeWorkmanFormat =
    /format\s+of\s+employee/.test(parts) ||
    (/employee\s*\/\s*workman\s*\/\s*worker/.test(parts) &&
      !/register\s+of\s+wages|rate\s+of\s+wage/.test(parts));

  // Rajasthan / Maharashtra Form A — even when the uploaded template still has employment-card body labels.
  if ((hasRajasthan || hasMaharashtra) && (hasFormA || hasEmployeeWorkmanFormat)) return true;
  if (hasRajasthan && /form_a_rj/.test(parts)) return true;
  if (hasMaharashtra && /form_a_mh/.test(parts)) return true;
  if ((hasRajasthan || hasMaharashtra) && headersIndicateFormARajasthanTable(tableHeaders)) return true;
  if (
    headersIndicateFormARajasthanHybridTemplate(tableHeaders) &&
    (hasRajasthan || hasMaharashtra || hasEmployeeWorkmanFormat || hasFormA)
  ) {
    return true;
  }

  if (/employment\s+card|form\s*xiv|form\s*14/.test(parts) && /name of the workman|tenure of employment/.test(parts)) {
    return false;
  }

  return false;
}

/** Same contractor header field Form C_RJ uses, so Site Management name+address can bind. */
export function enrichFormARajasthanDisplayHeader(formHeader) {
  const base = formHeader && typeof formHeader === 'object' ? { ...formHeader } : {};
  let fields = enrichEstablishmentPrincipalEmployerHeaderFields(base.fields || []);
  const hasContractor = fields.some((field) =>
    /name\s+and\s+address\s+of\s+(?:the\s+)?contractor/i.test(String(field?.label || ''))
  );
  if (!hasContractor) {
    fields = [
      ...fields,
      { label: 'Name and address of contractor', value: '', key: 'form_a_rj_contractor' },
    ];
  }
  return { ...base, fields };
}

export function applyFormARajasthanContractorFromSite(headerData, site, formHeaderFields = []) {
  return applyFormCRJContractorFromSite(headerData, site, formHeaderFields);
}

export function resolveFormARajasthanTableHeaders(tableHeaders) {
  const parsed = Array.isArray(tableHeaders)
    ? tableHeaders.map((h) => String(h || '').trim()).filter(Boolean)
    : [];
  const visible = parsed.filter((h) => !isFormARajasthanHiddenTableHeader(h));
  if (headersIndicateFormARajasthanTable(visible) || headersIndicateFormARajasthanTable(parsed)) {
    return visible.length >= 6 ? visible : [...FORM_A_RJ_CANONICAL_TABLE_HEADERS];
  }
  if (visible.length >= 8) return visible;
  return [...FORM_A_RJ_CANONICAL_TABLE_HEADERS];
}

/** Full template headers for Excel export — official RJ Part-A columns only. */
export function resolveFormARajasthanExportHeaders(tableHeaders) {
  const parsed = Array.isArray(tableHeaders)
    ? tableHeaders.map((h) => String(h || '').trim()).filter(Boolean)
    : [];
  const visible = parsed.filter((h) => !isFormARajasthanHiddenTableHeader(h));
  if (
    headersIndicateFormARajasthanTable(visible) ||
    headersIndicateFormARajasthanHybridTemplate(visible) ||
    headersIndicateFormARajasthanTable(parsed)
  ) {
    return [...FORM_A_RJ_CANONICAL_TABLE_HEADERS];
  }
  if (visible.length >= 8) return visible;
  return [...FORM_A_RJ_CANONICAL_TABLE_HEADERS];
}

export function remapFormARajasthanRowsToHeaders(rows, sourceHeaders, targetHeaders) {
  const src = Array.isArray(sourceHeaders) ? sourceHeaders : [];
  const tgt = resolveFormARajasthanExportHeaders(targetHeaders);
  if (!Array.isArray(rows)) return [];
  return rows.map((row, rowIndex) => {
    if (!row || typeof row !== 'object') return {};
    const out = {};
    tgt.forEach((targetHeader, colIdx) => {
      let val = '';
      if (Object.prototype.hasOwnProperty.call(row, targetHeader)) {
        val = row[targetHeader];
      } else if (src[colIdx] && Object.prototype.hasOwnProperty.call(row, src[colIdx])) {
        val = row[src[colIdx]];
      } else {
        const bucket = formAGJHeaderAliasBucket(formAGJGujaratHeaderNorm(targetHeader));
        for (const [k, v] of Object.entries(row)) {
          if (formAGJHeaderAliasBucket(formAGJGujaratHeaderNorm(k)) === bucket) {
            val = v;
            break;
          }
        }
      }
      const bucket = formAGJHeaderAliasBucket(formAGJGujaratHeaderNorm(targetHeader));
      if ((val == null || val === '') && bucket === 'firstName') {
        for (const [k, v] of Object.entries(row)) {
          const kn = formAGJGujaratHeaderNorm(k);
          if (/nature of work/.test(kn) && v != null && String(v).trim() !== '') {
            val = v;
            break;
          }
        }
      }
      if ((val == null || val === '') && bucket === 'sno') {
        val = rowIndex + 1;
      }
      out[targetHeader] = val == null ? '' : val;
    });
    return out;
  });
}

export function filterFormARajasthanExportRows(rows, headers) {
  const exportHdrs = resolveFormARajasthanExportHeaders(headers);
  return (Array.isArray(rows) ? rows : []).filter((row) =>
    rowHasMeaningfulFormARajasthanExportData(row, exportHdrs)
  );
}

export function rowHasMeaningfulFormARajasthanExportData(row, headers) {
  const hdrs = resolveFormARajasthanExportHeaders(headers);
  return hdrs.some((header) => {
    const bucket = formAGJHeaderAliasBucket(formAGJGujaratHeaderNorm(header));
    if (bucket === 'sno') return false;
    return getFormAGJGujaratRowValueForHeader(row, header) !== '';
  });
}

export function enrichFormARajasthanRowsForExport(
  mappedData,
  headers,
  employeesOverride = null,
  helpers = {}
) {
  const hdrs = resolveFormARajasthanExportHeaders(headers);
  const rows = filterFormARajasthanExportRows(
    remapFormARajasthanRowsToHeaders(Array.isArray(mappedData) ? mappedData : [], hdrs, hdrs),
    hdrs
  );
  const employees = Array.isArray(employeesOverride) ? employeesOverride : [];
  if (employees.length === 0 || rows.length === 0) return rows;

  return rows.map((row, index) => {
    const emp = employees[index];
    if (!emp) return row;
    // Pass RJ headers; mapFormAGJ may expand keys — copy back by alias bucket.
    const fromEmp = mapFormAGJGujaratRowsFromEmployees([emp], hdrs, {
      ...helpers,
      rowIndexOffset: index,
    })[0];
    const out = { ...row };
    hdrs.forEach((header) => {
      const existing = out[header];
      let empVal = fromEmp?.[header];
      if (empVal == null || String(empVal).trim() === '') {
        empVal = getFormAGJGujaratRowValueForHeader(fromEmp || {}, header, index);
      }
      if (
        (existing == null || String(existing).trim() === '') &&
        empVal != null &&
        String(empVal).trim() !== ''
      ) {
        out[header] = empVal;
      }
    });
    if (fromEmp?.__employeeLookupName) out.__employeeLookupName = fromEmp.__employeeLookupName;
    if (fromEmp?.__employeeLookupId) out.__employeeLookupId = fromEmp.__employeeLookupId;
    return out;
  });
}

const excelCellValueToString = (val) => {
  if (val == null) return '';
  if (typeof val === 'string' || typeof val === 'number' || typeof val === 'boolean') return String(val);
  if (val instanceof Date) return val.toISOString();
  if (typeof val === 'object') {
    if (Array.isArray(val.richText)) return val.richText.map((rt) => rt?.text || '').join('');
    if (val.text != null) return String(val.text);
    if (val.result != null) return String(val.result);
  }
  return '';
};

function excelCellLooksLikeSerialHeader(text) {
  const t = formAGJGujaratHeaderNorm(text).replace(/[^a-z0-9 ]/g, ' ').replace(/\s+/g, ' ').trim();
  if (t.length > 22) return false;
  return (
    /^sr\.?\s*no|^s\.?\s*no|^sl\.?\s*no|^sl\s*n$|^serial\s*no/.test(t) || t === 'sl n'
  );
}

function rowLooksLikeFormARajasthanTableHeaderRow(getMergedAwareCellText, r, startCol) {
  if (!excelCellLooksLikeSerialHeader(getMergedAwareCellText(r, startCol))) return false;
  for (let c = startCol + 1; c <= startCol + 4; c += 1) {
    const t = formAGJGujaratHeaderNorm(getMergedAwareCellText(r, c));
    if (cellTextLooksLikeEmploymentCardBodyLabel(getMergedAwareCellText(r, c))) return false;
    if (formAGJHeaderAliasBucket(t) === 'employeeId') return true;
    if (/employee/.test(t) && /code/.test(t)) return true;
  }
  return false;
}

function buildMergeTopLeftResolver(worksheet) {
  const cache = new Map();
  return (r, c) => {
    const key = `${r}:${c}`;
    if (cache.has(key)) return cache.get(key);
    let topLeft = { r, c };
    const merges = worksheet.model?.merges;
    if (Array.isArray(merges)) {
      for (let mi = 0; mi < merges.length; mi += 1) {
        const parts = String(merges[mi] || '').split(':');
        if (parts.length !== 2) continue;
        const tl = worksheet.getCell(parts[0]);
        const br = worksheet.getCell(parts[1]);
        if (!tl || !br) continue;
        if (r >= tl.row && r <= br.row && c >= tl.col && c <= br.col) {
          topLeft = { r: tl.row, c: tl.col };
          break;
        }
      }
    }
    cache.set(key, topLeft);
    return topLeft;
  };
}

/** Reject Form XIV Employment Card templates mistaken for Form A RJ. */
export function sheetLooksLikeFormXIVEmploymentCard(worksheet) {
  if (!worksheet) return false;
  const getMergeTopLeft = buildMergeTopLeftResolver(worksheet);
  let workmanHits = 0;
  let tenureHits = 0;
  let natureHits = 0;
  const maxR = Math.min(40, Math.max(worksheet.rowCount || 0, 25));
  for (let r = 1; r <= maxR; r += 1) {
    for (let c = 1; c <= 6; c += 1) {
      const tl = getMergeTopLeft(r, c);
      const t = formAGJGujaratHeaderNorm(
        excelCellValueToString(worksheet.getCell(tl.r, tl.c)?.value)
      );
      if (/name of the workman/.test(t)) workmanHits += 1;
      if (/tenure of employment/.test(t)) tenureHits += 1;
      if (/nature of work and location/.test(t) && !/category/.test(t)) natureHits += 1;
    }
  }
  return workmanHits >= 1 && (tenureHits >= 1 || natureHits >= 1);
}

function scoreFormARajasthanHeaderRow(getMergedAwareCellText, r, startCol) {
  let score = 0;
  let hasName = false;
  let hasEmployeeCode = false;
  for (let c = startCol; c < startCol + 20; c += 1) {
    const raw = getMergedAwareCellText(r, c);
    if (!raw) continue;
    const t = formAGJGujaratHeaderNorm(raw);
    const bucket = formAGJHeaderAliasBucket(t);
    if (excelCellLooksLikeSerialHeader(raw)) score += 3;
    if (bucket === 'employeeId') {
      score += 4;
      hasEmployeeCode = true;
    }
    if (bucket === 'firstName') {
      score += 4;
      hasName = true;
    }
    if (bucket === 'surname') score += 2;
    if (bucket === 'skillCategory') score += 3;
    if (bucket === 'education') score += 1;
    if (bucket === 'employmentType') score += 1;
    // Wrong / hybrid templates
    if (/nature of work and location/.test(t)) score -= 8;
    if (/name of the workman/.test(t)) score -= 8;
    if (/tenure of employment/.test(t)) score -= 8;
  }
  if (hasName && hasEmployeeCode) score += 5;
  return score;
}

function findFormARajasthanFooterStartRow(worksheet, dataStartRow, getMergedAwareCellText) {
  const maxR = Math.max(dataStartRow + 80, worksheet.rowCount || 0);
  for (let r = dataStartRow; r <= maxR; r += 1) {
    const a = getMergedAwareCellText(r, 1);
    const joined = [1, 2, 3]
      .map((c) => getMergedAwareCellText(r, c))
      .filter(Boolean)
      .join(' ');
    if (/^\*/.test(a) || /^#\s*note/i.test(a) || /highly\s*skilled\/skilled/i.test(joined)) {
      return r;
    }
  }
  return -1;
}

function detectFormARajasthanTableLayout(worksheet, hints = {}) {
  const getMergeTopLeft = buildMergeTopLeftResolver(worksheet);
  const getMergedAwareCellText = (r, c) => {
    const tl = getMergeTopLeft(r, c);
    return excelCellValueToString(worksheet.getCell(tl.r, tl.c)?.value).trim();
  };

  let headerRow =
    hints.parsedHeaderRowIndex != null && hints.parsedHeaderRowIndex >= 0
      ? hints.parsedHeaderRowIndex + 1
      : -1;
  // Data always maps to columns A–O on the official RJ register; ignore parsed start-col hints.
  let startCol = FORM_A_RJ_TABLE_START_COL;

  if (
    headerRow > 0 &&
    !rowLooksLikeFormARajasthanTableHeaderRow(getMergedAwareCellText, headerRow, FORM_A_RJ_TABLE_START_COL)
  ) {
    headerRow = -1;
  }

  if (headerRow < 1) {
    const maxScanRows = Math.max(35, worksheet.rowCount + 5);
    let best = { score: -1, row: -1, col: FORM_A_RJ_TABLE_START_COL };
    for (let r = 1; r <= maxScanRows; r += 1) {
      if (!rowLooksLikeFormARajasthanTableHeaderRow(getMergedAwareCellText, r, FORM_A_RJ_TABLE_START_COL)) {
        continue;
      }
      const score = scoreFormARajasthanHeaderRow(
        getMergedAwareCellText,
        r,
        FORM_A_RJ_TABLE_START_COL
      );
      if (score > best.score) best = { score, row: r, col: FORM_A_RJ_TABLE_START_COL };
    }
    if (best.row > 0 && best.score >= 1) {
      headerRow = best.row;
      startCol = FORM_A_RJ_TABLE_START_COL;
    }
  }
  if (headerRow < 1) return null;

  let dataStartRow =
    hints.parsedDataStartIndex != null && hints.parsedDataStartIndex >= 0
      ? hints.parsedDataStartIndex + 1
      : headerRow + 1;

  if (dataStartRow <= headerRow) dataStartRow = headerRow + 1;
  if (
    hints.parsedDataStartIndex != null &&
    hints.parsedDataStartIndex >= 0 &&
    cellTextLooksLikeEmploymentCardBodyLabel(getMergedAwareCellText(dataStartRow, 2))
  ) {
    dataStartRow = headerRow + 1;
  }

  let seqHits = 0;
  for (let i = 0; i < 8; i += 1) {
    const t = getMergedAwareCellText(dataStartRow, startCol + i).replace(/\s+/g, '').trim();
    if (t === String(i + 1)) seqHits += 1;
  }
  if (seqHits >= 3) {
    dataStartRow = headerRow + 2;
  }

  const footerStartRow = findFormARajasthanFooterStartRow(
    worksheet,
    dataStartRow,
    getMergedAwareCellText
  );

  const templateCols = buildFormARajasthanCanonicalTemplateCols(headerRow, startCol);

  if (templateCols.length < 4) return null;

  return {
    headerRow,
    dataStartRow,
    templateCols,
    footerStartRow,
    startCol,
    rowsPerEmployee: FORM_A_RJ_EXPORT_ROWS_PER_EMPLOYEE,
  };
}

function writeFormARajasthanCellValue(cell, bucket, val) {
  if (val == null || val === '') {
    cell.value = '';
    return;
  }
  if (bucket === 'sno') {
    const n = Number(String(val).replace(/[,]/g, '').trim());
    cell.value = Number.isFinite(n) ? n : String(val);
  } else if (
    bucket === 'mobile' ||
    bucket === 'uan' ||
    bucket === 'bankAccount' ||
    bucket === 'aadhaar' ||
    bucket === 'esic'
  ) {
    const digits = String(val).replace(/\D/g, '');
    cell.value = digits || String(val);
    if (digits) cell.numFmt = '0';
  } else if (bucket === 'pan') {
    cell.value = String(val).trim().toUpperCase();
  } else if (bucket === 'skillCategory') {
    cell.value = String(val).trim();
  } else {
    cell.value = String(val);
  }
  cell.alignment = { ...(cell.alignment || {}), vertical: 'middle', wrapText: true };
}

/**
 * Rajasthan Form A Excel export — fill the original Form Master template in place.
 * Never creates a new blank workbook; requires the official Form A_RJ .xlsx buffer.
 */
export async function buildFormARajasthanWorkbookWithTemplateStyles({
  templateArrayBuffer,
  mappedData,
  headersToUse,
  parsedFormHeader,
  formFileName,
  headerFormData,
  parsedHeaderRowIndex,
  parsedDataStartIndex,
  parsedTableStartCol,
  employeesOverride = null,
  formatStatutoryDateDisplay = null,
}) {
  if (!templateArrayBuffer) {
    throw new Error('Original form template buffer is required for Form A RJ export.');
  }

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(templateArrayBuffer);
  const worksheet = workbook.worksheets[0];
  if (!worksheet) throw new Error('Template worksheet not found.');

  if (sheetLooksLikeFormXIVEmploymentCard(worksheet)) {
    // Hybrid templates (table header + employment-card body labels) are repaired in place.
  }

  // Title bands may insert a Part-A row — detect table layout after that shift.
  writeFormARajasthanTitleBandsInColumnsEG(worksheet, parsedFormHeader);

  const layout = detectFormARajasthanTableLayout(worksheet, {
    parsedHeaderRowIndex,
    parsedDataStartIndex,
    parsedTableStartCol,
  });
  if (!layout) throw new Error('Could not locate Rajasthan Form A table header row.');

  const { dataStartRow, templateCols, footerStartRow, headerRow } = layout;

  applyFormARajasthanCanonicalHeadersAndWidths(worksheet, headerRow, FORM_A_RJ_TABLE_START_COL);

  // Keep original establishment fields above the table.
  if (headerFormData && typeof headerFormData === 'object') {
    writeStatutoryHeaderFieldsToExcelJsWorksheet(worksheet, {
      headerFormData,
      parsedFormHeader,
      headerRowEnd: Math.max(1, layout.headerRow - 1),
      maxScanCols: 80,
    });
  }

  const templateHeaderLabels = templateCols.map((entry) => entry.label);
  const normalizedHeaders = resolveFormARajasthanExportHeaders(
    Array.isArray(headersToUse) && headersToUse.length >= 6
      ? headersToUse
      : templateHeaderLabels.length > 0
        ? templateHeaderLabels
        : FORM_A_RJ_CANONICAL_TABLE_HEADERS
  );
  const exportHelpers = {
    formatStatutoryDateDisplay:
      typeof formatStatutoryDateDisplay === 'function'
        ? formatStatutoryDateDisplay
        : (v) => String(v || '').trim(),
  };
  const rows = enrichFormARajasthanRowsForExport(
    mappedData,
    normalizedHeaders,
    employeesOverride,
    exportHelpers
  ).map((row, index) => {
    const emp = Array.isArray(employeesOverride) ? employeesOverride[index] : null;
    if (!emp) return row;
    const fromEmp = mapFormAGJGujaratRowsFromEmployees([emp], normalizedHeaders, {
      ...exportHelpers,
      rowIndexOffset: index,
    })[0];
    const out = { ...row };
    normalizedHeaders.forEach((header) => {
      const empVal = getFormAGJGujaratRowValueForHeader(fromEmp || {}, header, index);
      if (
        (out[header] == null || String(out[header]).trim() === '') &&
        empVal != null &&
        String(empVal).trim() !== ''
      ) {
        out[header] = empVal;
      }
    });
    if (fromEmp?.__employeeLookupId && !String(out.__employeeLookupId || '').trim()) {
      out.__employeeLookupId = fromEmp.__employeeLookupId;
    }
    return out;
  });

  const tableColMin = Math.min(...templateCols.map(({ col }) => col));
  const tableColMax = Math.max(...templateCols.map(({ col }) => col));
  // Form A MH (and similar) keep ESIC IP / LWF / AADHAAR… after PAN — border the full header span.
  const fullTableColTo = Math.max(
    tableColMax,
    findFormARegisterHeaderColTo(worksheet, headerRow, FORM_A_RJ_TABLE_START_COL)
  );

  // Preserve footer notes (* HS/S/SS/US and #Note…) — clear only the data band.
  const clearEndExclusive =
    footerStartRow > dataStartRow
      ? footerStartRow
      : Math.max(dataStartRow + Math.max(rows.length, 1), dataStartRow + 1);

  // Tall merged blanks after PAN (Form A MH) hide per-row boxes — split before clear/write.
  unmergeFormARegisterDataBand(
    worksheet,
    dataStartRow,
    Math.max(clearEndExclusive - 1, dataStartRow + Math.max(rows.length, 1) - 1),
    tableColMin,
    fullTableColTo
  );

  for (let r = dataStartRow; r < clearEndExclusive; r += 1) {
    for (let c = tableColMin; c <= fullTableColTo; c += 1) {
      worksheet.getCell(r, c).value = '';
    }
  }

  // Restore canonical headers if the clear band overlapped the header row.
  applyFormARajasthanCanonicalHeadersAndWidths(worksheet, headerRow, FORM_A_RJ_TABLE_START_COL);

  // Insert rows before footer when employee count exceeds the blank template body.
  if (footerStartRow > dataStartRow && rows.length > footerStartRow - dataStartRow) {
    const need = rows.length - (footerStartRow - dataStartRow);
    worksheet.spliceRows(footerStartRow, 0, ...Array.from({ length: need }, () => []));
  }

  // Clear employment-card label cells in column B (hybrid templates).
  const clearEndForLabels =
    footerStartRow > dataStartRow ? footerStartRow : dataStartRow + Math.max(rows.length, 14);
  for (let r = dataStartRow; r < clearEndForLabels; r += 1) {
    const bText = excelCellValueToString(worksheet.getCell(r, 2)?.value).trim();
    if (cellTextLooksLikeEmploymentCardBodyLabel(bText)) {
      worksheet.getCell(r, 2).value = '';
    }
  }

  rows.forEach((row, idx) => {
    const targetRowNum = dataStartRow + idx;
    const targetRow = worksheet.getRow(targetRowNum);
    if (targetRow) targetRow.height = FORM_A_RJ_DATA_ROW_HEIGHT;
    templateCols.forEach(({ col, bucket, label }) => {
      let val = getFormAGJGujaratRowValueForHeader(row, label, idx);
      if ((val == null || val === '') && bucket) {
        for (const [k, v] of Object.entries(row)) {
          if (
            formAGJHeaderAliasBucket(formAGJGujaratHeaderNorm(k)) === bucket &&
            v != null &&
            String(v).trim() !== ''
          ) {
            val = String(v).trim();
            break;
          }
        }
      }
      // Employee Code from enriched employee lookup when grid cell is empty.
      if ((val == null || val === '') && bucket === 'employeeId' && row.__employeeLookupId) {
        val = String(row.__employeeLookupId).trim();
      }
      if ((val == null || val === '') && bucket === 'sno') val = String(idx + 1);
      const cell = worksheet.getCell(targetRowNum, col);
      writeFormARajasthanCellValue(cell, bucket, val);
    });
  });

  const dataRowCount = Math.max(rows.length, 1);
  const borderRowTo = dataStartRow + dataRowCount - 1;
  // Unmerge again in case spliceRows/write reintroduced tall merges after PAN.
  unmergeFormARegisterDataBand(
    worksheet,
    dataStartRow,
    borderRowTo,
    tableColMin,
    fullTableColTo
  );
  // Header + data: full box through last MH column (ESIC IP, LWF, AADHAAR, …).
  applyExcelJSFullBoxBordersToRange(worksheet, {
    rowFrom: headerRow,
    rowTo: borderRowTo,
    colFrom: tableColMin,
    colTo: fullTableColTo,
  });
  if (rows.length > 0) {
    ensureExcelJSDataRowsWithBorders(worksheet, {
      dataStartRow,
      dataRowCount: rows.length,
      colFrom: tableColMin,
      colTo: fullTableColTo,
      templateRow: dataStartRow,
      templateBodyRows: 1,
    });
  }

  const out = await workbook.xlsx.writeBuffer();
  const fileName =
    formFileName ||
    parsedFormHeader?.title?.replace(/[^a-zA-Z0-9]/g, '_') ||
    'Form_A_RJ_Rajasthan.xlsx';
  return {
    blob: new Blob([out], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    }),
    buffer: out,
    fileName,
  };
}
