import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import { headersIndicateFormAGJEmployeeRegisterTable } from './formCGJGujarat';
import {
  applyExcelJSFullBoxBordersToRange,
  countExcelJSTemplateBodyRows,
  ensureExcelJSDataRowsWithBorders,
  excelJSCellHasBorder,
} from '../../utils/excelTableBorders';
import { flattenPayrollEarningColumns, readPayrollScalar } from '../../utils/payrollEarnings';
import {
  collectForm10RowNameParts,
  findForm10PayrollRowByFirstAndLastName,
} from './form10TamilNadu';
import {
  enrichEstablishmentPrincipalEmployerHeaderFields,
  excelCellValueToString,
  writeStatutoryHeaderFieldsToExcelJsWorksheet,
} from '../../utils/statutorySiteCompanyHeaders';

/** Gujarat Form D — Muster Roll (Relay or Set Work, Summary days/hours). */

export const FORM_DGJ_GJ_CANONICAL_TABLE_HEADERS = [
  'Sr. No. in Employee / Workman / Worker Register',
  'Name',
  'Relay or Set Work',
  'Summary No. of Days',
  'Remarks No. of Hours',
  'Signature of Register Keeper*',
];

/** Legal notes painted outside the Form D bordered box (Excel + PDF). */
export const FORM_DGJ_GJ_ELECTRONIC_FORMAT_NOTE_1 =
  '*Not necessary in case of electronic format';
export const FORM_DGJ_GJ_ELECTRONIC_FORMAT_NOTE_2 =
  '**Not necessary in case of electronic format';
export const FORM_DGJ_GJ_GOVERNOR_ORDER_NOTE =
  'By order and in the name of the Governor of Gujarat';
export const FORM_DGJ_GJ_OUTER_FOOTNOTES = [
  FORM_DGJ_GJ_ELECTRONIC_FORMAT_NOTE_1,
  FORM_DGJ_GJ_ELECTRONIC_FORMAT_NOTE_2,
  FORM_DGJ_GJ_GOVERNOR_ORDER_NOTE,
];

export function isFormDGJGujaratOuterFootnoteText(text) {
  const norm = String(text || '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase()
    .replace(/elecrtonic/g, 'electronic');
  if (!norm) return false;
  if (/^\*{1,2}\s*not\s+necessary\s+in\s+case\s+of\s+electronic\s+format\.?$/.test(norm)) {
    return true;
  }
  if (/^by\s+order\s+and\s+in\s+the\s+name\s+of\s+the\s+governor\s+of\s+gujarat\.?$/.test(norm)) {
    return true;
  }
  return false;
}

function formDGJExcelCellText(worksheet, r, c) {
  return excelCellValueToString(worksheet?.getCell(r, c)?.value).trim();
}

function clearFormDGJCellBorder(cell) {
  if (!cell) return;
  try {
    cell.border = {};
    const prev = cell.style && typeof cell.style === 'object' ? { ...cell.style } : {};
    cell.style = { ...prev, border: {} };
  } catch (_) {
    try {
      cell.border = {};
    } catch (__) {
      /* ignore */
    }
  }
}

/**
 * First 5 Form D GJ header meta labels — export as plain text (no box borders).
 * Keeps "For the period From …" and the muster-roll table box untouched.
 */
export function isFormDGJGujaratPlainHeaderMetaLabel(text) {
  const raw = String(text || '')
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  if (!raw) return false;
  const labelOnly = raw.split(':')[0].trim().toLowerCase().replace(/^\*+\s*/, '');
  if (/for\s+the\s+period\s+from/.test(labelOnly)) return false;
  if (/^name\s+of\s+establishment\b/.test(labelOnly) && !/principal|employer|contractor/.test(labelOnly)) {
    return true;
  }
  if (/^name\s+of\s+owner\b/.test(labelOnly)) return true;
  if (
    /labour\s+identification\s+no/.test(labelOnly) &&
    /principal\s+employer/.test(labelOnly)
  ) {
    return true;
  }
  if (
    /labour\s+identification\s+no/.test(labelOnly) &&
    !/principal\s+employer/.test(labelOnly)
  ) {
    return true;
  }
  if (/name\s+and\s+address\s+of\s+principal\s+employer/.test(labelOnly)) return true;
  return false;
}

function formDGJFindMergeBounds(worksheet, row, col) {
  const merges = worksheet?.model?.merges;
  if (Array.isArray(merges)) {
    for (let i = 0; i < merges.length; i += 1) {
      const parts = String(merges[i] || '').split(':');
      if (parts.length !== 2) continue;
      try {
        const tl = worksheet.getCell(parts[0]);
        const br = worksheet.getCell(parts[1]);
        if (!tl || !br) continue;
        if (row >= tl.row && row <= br.row && col >= tl.col && col <= br.col) {
          return { r0: tl.row, r1: br.row, c0: tl.col, c1: br.col };
        }
      } catch (_) {
        /* ignore */
      }
    }
  }
  return { r0: row, r1: row, c0: col, c1: col };
}

/**
 * Strip template box borders from the first 5 header meta fields so Excel shows
 * plain text only (Name of Establishment / Owner / LIN / Principal Employer / PE LIN).
 */
export function clearFormDGJGujaratHeaderMetaBoxes(
  worksheet,
  { headerRowEnd = 25, colTo = 12 } = {}
) {
  if (!worksheet) return;
  const rowEnd = Math.max(1, Number(headerRowEnd) || 25);
  const colEnd = Math.max(4, Number(colTo) || 12);
  const cleared = new Set();

  for (let r = 1; r <= rowEnd; r += 1) {
    for (let c = 1; c <= colEnd; c += 1) {
      const text = formDGJExcelCellText(worksheet, r, c);
      if (!isFormDGJGujaratPlainHeaderMetaLabel(text)) continue;
      const bounds = formDGJFindMergeBounds(worksheet, r, c);
      // Template often boxes an empty row under a single-row label (e.g. A10:A11).
      let r1 = bounds.r1;
      if (bounds.r0 === bounds.r1 && bounds.r0 + 1 <= rowEnd) {
        const nextText = formDGJExcelCellText(worksheet, bounds.r0 + 1, bounds.c0);
        if (
          !nextText &&
          !/for\s+the\s+period\s+from/i.test(
            formDGJExcelCellText(worksheet, bounds.r0 + 1, 1)
          )
        ) {
          r1 = bounds.r0 + 1;
        }
      }
      const c1 = Math.max(bounds.c1, Math.min(colEnd, bounds.c0 + 3));
      for (let rr = bounds.r0; rr <= Math.min(r1, rowEnd); rr += 1) {
        for (let cc = bounds.c0; cc <= Math.min(c1, colEnd + 4); cc += 1) {
          const key = `${rr}:${cc}`;
          if (cleared.has(key)) continue;
          cleared.add(key);
          clearFormDGJCellBorder(worksheet.getCell(rr, cc));
        }
      }
    }
  }
}

export function isFormDGJGujaratSystemGeneratedText(text) {
  return /this\s+is\s+a\s+system\s+generated\s+document|system\s+generated\s+document/i.test(
    String(text || '').trim()
  );
}

/** Locate the template "System Generated" row below the muster table. */
export function findFormDGJGujaratSystemGeneratedRow(
  worksheet,
  dataStartRow = 1,
  colFrom = 1,
  colTo = 6
) {
  if (!worksheet) return null;
  const r0 = Math.max(1, Number(dataStartRow) || 1);
  const c0 = Math.max(1, Number(colFrom) || 1);
  const c1 = Math.max(c0, Number(colTo) || c0);
  const maxRows = Math.max(worksheet.rowCount || 0, r0 + 80);
  for (let r = r0; r <= maxRows; r += 1) {
    for (let c = 1; c <= Math.max(c1 + 8, 12); c += 1) {
      if (isFormDGJGujaratSystemGeneratedText(formDGJExcelCellText(worksheet, r, c))) {
        return r;
      }
    }
  }
  return null;
}

/**
 * Last bordered table-body row (before System Generated / outer footnotes).
 * Keeps empty template body rows inside the box — matches the Form D model.
 */
export function findFormDGJGujaratTableBoxEndRow(
  worksheet,
  dataStartRow = 1,
  colFrom = 1,
  colTo = 6,
  hints = {}
) {
  if (!worksheet) return Math.max(1, dataStartRow);
  const r0 = Math.max(1, Number(dataStartRow) || 1);
  const c0 = Math.max(1, Number(colFrom) || 1);
  const c1 = Math.max(c0, Number(colTo) || c0);
  const systemRow =
    Number(hints.systemGeneratedRow) > 0
      ? Number(hints.systemGeneratedRow)
      : findFormDGJGujaratSystemGeneratedRow(worksheet, r0, c0, c1);
  const scanEnd =
    systemRow && systemRow > r0
      ? systemRow - 1
      : Math.max(worksheet.rowCount || 0, r0 + 40);

  let lastBordered = r0;
  for (let r = r0; r <= scanEnd; r += 1) {
    let bordered = false;
    let footnoteOrSystem = false;
    for (let c = c0; c <= c1; c += 1) {
      const cell = worksheet.getCell(r, c);
      const t = formDGJExcelCellText(worksheet, r, c);
      if (isFormDGJGujaratOuterFootnoteText(t) || isFormDGJGujaratSystemGeneratedText(t)) {
        footnoteOrSystem = true;
      }
      if (excelJSCellHasBorder(cell)) bordered = true;
    }
    if (footnoteOrSystem) break;
    if (bordered) lastBordered = r;
    else if (r > r0 && lastBordered >= r0) break;
  }

  const minEmptyAfterData = Number(hints.minEmptyAfterData);
  const lastDataRow = Number(hints.lastDataRow);
  let end = lastBordered;
  if (Number.isFinite(lastDataRow) && lastDataRow >= r0) {
    const withPadding =
      lastDataRow + (Number.isFinite(minEmptyAfterData) ? Math.max(0, minEmptyAfterData) : 2);
    end = Math.max(end, withPadding);
  }
  if (systemRow && systemRow > r0) end = Math.min(end, systemRow - 1);
  return Math.max(r0, end);
}

function clearFormDGJRowOutsideBox(worksheet, row, colFrom, colTo) {
  if (!worksheet || !row) return;
  const c0 = Math.max(1, colFrom);
  const c1 = Math.max(c0, colTo);
  try {
    worksheet.unMergeCells(row, c0, row, Math.max(c1, c0 + 8));
  } catch (_) {
    try {
      worksheet.unMergeCells(row, c0, row, c1);
    } catch (__) {
      /* ignore */
    }
  }
  for (let c = c0; c <= Math.max(c1, c0 + 8); c += 1) {
    const cell = worksheet.getCell(row, c);
    cell.value = '';
    clearFormDGJCellBorder(cell);
  }
}

/**
 * Centered System Generated line directly under the table box (no borders).
 */
export function writeFormDGJGujaratSystemGeneratedNote(worksheet, options = {}) {
  if (!worksheet) return null;
  const {
    afterRow,
    startCol = 1,
    endCol = 6,
    text = 'This is a System Generated Document',
  } = options;
  const baseRow =
    Number.isFinite(Number(afterRow)) && Number(afterRow) > 0 ? Number(afterRow) : null;
  if (!baseRow) return null;
  const row = baseRow + 1;
  const colFrom = Math.max(1, Number(startCol) || 1);
  const colTo = Math.max(colFrom, Number(endCol) || colFrom);

  // Remove any other System Generated copies below the box.
  const scanEnd = Math.max(worksheet.rowCount || 0, row + 20);
  for (let r = row; r <= scanEnd; r += 1) {
    for (let c = 1; c <= Math.max(colTo + 8, 12); c += 1) {
      if (!isFormDGJGujaratSystemGeneratedText(formDGJExcelCellText(worksheet, r, c))) continue;
      clearFormDGJRowOutsideBox(worksheet, r, 1, Math.max(colTo + 8, 12));
      break;
    }
  }

  clearFormDGJRowOutsideBox(worksheet, row, colFrom, colTo);
  try {
    worksheet.mergeCells(row, colFrom, row, colTo);
  } catch (_) {
    /* ignore */
  }
  const cell = worksheet.getCell(row, colFrom);
  cell.value = text;
  cell.font = { name: 'Arial', size: 9, bold: false, italic: false };
  cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: false };
  clearFormDGJCellBorder(cell);
  for (let c = colFrom; c <= colTo; c += 1) clearFormDGJCellBorder(worksheet.getCell(row, c));
  const excelRow = worksheet.getRow(row);
  if (excelRow) excelRow.height = Math.max(Number(excelRow.height) || 0, 18);
  return row;
}

/**
 * Write Form D Gujarat legal footnotes below the register box / System Generated
 * line, with no borders — matches the PDF outer notes.
 */
export function writeFormDGJGujaratOuterFootnotes(worksheet, options = {}) {
  if (!worksheet) return null;
  const {
    afterRow,
    startCol = 1,
    endCol = 6,
    footnotes = FORM_DGJ_GJ_OUTER_FOOTNOTES,
  } = options;
  const baseRow =
    Number.isFinite(Number(afterRow)) && Number(afterRow) > 0 ? Number(afterRow) : null;
  if (!baseRow) return null;

  const colFrom = Math.max(1, Number(startCol) || 1);
  const colTo = Math.max(colFrom, Number(endCol) || colFrom);
  const lines = (Array.isArray(footnotes) ? footnotes : FORM_DGJ_GJ_OUTER_FOOTNOTES)
    .map((t) => String(t || '').trim())
    .filter(Boolean);
  if (!lines.length) return null;

  // Drop any prior copies of these notes so re-download stays clean.
  const scanEnd = Math.max(worksheet.rowCount || 0, baseRow + 16);
  for (let r = Math.max(1, baseRow - 2); r <= scanEnd; r += 1) {
    for (let c = 1; c <= Math.max(colTo + 4, 12); c += 1) {
      const existing = formDGJExcelCellText(worksheet, r, c);
      if (!isFormDGJGujaratOuterFootnoteText(existing)) continue;
      clearFormDGJRowOutsideBox(worksheet, r, colFrom, colTo);
      break;
    }
  }

  let firstWrittenRow = null;
  lines.forEach((text, idx) => {
    const row = baseRow + 1 + idx;
    if (firstWrittenRow == null) firstWrittenRow = row;
    clearFormDGJRowOutsideBox(worksheet, row, colFrom, colTo);
    try {
      worksheet.mergeCells(row, colFrom, row, colTo);
    } catch (_) {
      /* overlap — keep A-cell text */
    }
    const cell = worksheet.getCell(row, colFrom);
    cell.value = text;
    cell.font = { name: 'Arial', size: 9, italic: false, bold: false };
    cell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
    clearFormDGJCellBorder(cell);
    for (let c = colFrom; c <= colTo; c += 1) clearFormDGJCellBorder(worksheet.getCell(row, c));
    const excelRow = worksheet.getRow(row);
    if (excelRow) excelRow.height = Math.max(Number(excelRow.height) || 0, 16);
  });

  return firstWrittenRow;
}

/** @deprecated use findFormDGJGujaratTableBoxEndRow */
export function findFormDGJGujaratBoxEndRow(worksheet, dataStartRow = 1, colFrom = 1, colTo = 6) {
  return findFormDGJGujaratTableBoxEndRow(worksheet, dataStartRow, colFrom, colTo);
}

export function formDGJGujaratHeaderNorm(txt) {
  return String(txt || '')
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

export function blobIndicatesFormDGJMusterRoll(blob, tableHeaders = null) {
  const text = String(blob || '').toLowerCase();
  if (/\bform[\s._-]*d\b/.test(text) && /gujarat|\b_gj\b|form_d_gj/.test(text)) return true;
  if (/relay\s+or\s+set\s+work/.test(text)) return true;
  if (/summary\s+no\.?\s*of\s+days/.test(text) && /register\s+keeper/.test(text)) return true;
  if (Array.isArray(tableHeaders) && tableHeaders.length > 0) {
    const hdr = tableHeaders.map((h) => formDGJGujaratHeaderNorm(h)).join('\n');
    if (/relay\s+or\s+set/.test(hdr)) return true;
    if (/summary.*days/.test(hdr) && /register\s+keeper|signature/.test(hdr)) return true;
    if (/sr\.?\s*no.*register/.test(hdr) && /relay\s+or\s+set/.test(hdr)) return true;
  }
  return false;
}

export function headersIndicateFormDGJMusterTable(tableHeaders) {
  if (!Array.isArray(tableHeaders) || tableHeaders.length < 3) return false;
  const joined = tableHeaders.map((h) => formDGJGujaratHeaderNorm(h)).join('\n');
  const hasRelay = /relay\s+or\s+set/.test(joined);
  const hasSummaryDays = /summary.*days|summery.*days|no\.?\s*of\s+days/.test(joined);
  const hasName = /\bname\b/.test(joined);
  const hasRegisterKeeper = /register\s+keeper|signature/.test(joined);
  const hasSrRegister =
    /sr\.?\s*no/.test(joined) &&
    (/employee.*register|workman.*register|worker.*register/.test(joined) ||
      /register/.test(joined));
  if (hasRelay && hasSummaryDays) return true;
  if (hasRelay && hasName && hasRegisterKeeper) return true;
  if (hasSrRegister && hasRelay) return true;
  return false;
}

export function isFormDGJGujaratContext(
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

  if (/rajasthan|\b_rj\b|form[\s._-]*d[\s._-]*rj|form_d_rj/.test(parts)) return false;

  const hasGujarat = /gujarat|\b_gj\b|form[\s._-]*d[\s._-]*gj|form_d_gj/.test(parts);
  const hasFormD = /\bform[\s._-]*d\b/.test(parts);

  if (hasGujarat && hasFormD) return true;
  if (hasFormD && blobIndicatesFormDGJMusterRoll(parts, tableHeaders)) return true;
  if (hasGujarat && blobIndicatesFormDGJMusterRoll(parts, tableHeaders)) return true;
  if (headersIndicateFormDGJMusterTable(tableHeaders) && (hasGujarat || hasFormD)) return true;
  return false;
}

export function isFormDGJGujaratUsingWrongEmployeeRegisterTable(tableHeaders) {
  return (
    headersIndicateFormAGJEmployeeRegisterTable(tableHeaders) &&
    !headersIndicateFormDGJMusterTable(tableHeaders)
  );
}

function buildSheetTextBlob(workbook, sheetName) {
  const ws = workbook?.Sheets?.[sheetName];
  if (!ws) return '';
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  return rows
    .slice(0, 45)
    .map((row) => (Array.isArray(row) ? row : []).map((c) => String(c || '')).join(' '))
    .join(' ')
    .toLowerCase();
}

function scoreFormDGJGujaratSheet(sheetName, sheetText, hintsBlob) {
  const sheetLower = String(sheetName || '').toLowerCase().replace(/\s+/g, ' ').trim();
  const sheetBlob = `${sheetLower} ${sheetText}`;
  let score = 0;

  if (/\bform[\s._-]*d\b/.test(sheetLower) && !/\bform[\s._-]*d[\s._-]*(a|b|c|e|f)\b/.test(sheetLower)) {
    score += 240;
  }
  if (/\bform\s*d\b/.test(sheetText) && /relay\s+or\s+set/.test(sheetText)) score += 200;
  if (/relay\s+or\s+set\s+work/.test(sheetText)) score += 90;
  if (/summary\s+no\.?\s*of\s+days/.test(sheetText)) score += 70;
  if (/register\s+keeper/.test(sheetText)) score += 50;
  if (/sr\.?\s*no.*register/.test(sheetText)) score += 40;
  if (/remarks.*hours|no\.?\s*of\s+hours/.test(sheetText)) score += 30;

  if (/surname/.test(sheetText) && /\bgender\b/.test(sheetText) && /education/.test(sheetText)) {
    score -= 150;
  }
  if (/format\s+of\s+employee|type\s+of\s+employment/.test(sheetText) && /education/.test(sheetText)) {
    score -= 120;
  }
  if (/\bform[\s._-]*a\b/.test(sheetLower) && hintsBlob.includes('form') && hintsBlob.includes('d')) {
    score -= 180;
  }
  if (/\bform[\s._-]*b\b/.test(sheetLower) && /register\s+of\s+wages|rate\s+of\s+wage/.test(sheetText)) {
    score -= 100;
  }
  if (/\bform[\s._-]*c\b/.test(sheetLower) && /register\s+of\s+deductions|damage/.test(sheetText)) {
    score -= 100;
  }
  if (/gujarat|\b_gj\b|form_d_gj/.test(hintsBlob) && /gujarat|\b_gj\b/.test(sheetBlob)) score += 20;

  return score;
}

export function resolveFormDGJGujaratWorkbookSheetName(workbook, hints = {}) {
  const names = Array.isArray(workbook?.SheetNames) ? workbook.SheetNames : [];
  if (!names.length) return null;
  if (names.length === 1) return names[0];

  const blob = [
    hints.fileName,
    hints.formFileName,
    hints.formName,
    hints.item?.formName,
    hints.item?.FormName,
    hints.formHeaderTitle,
    hints.formHeader?.title,
  ]
    .filter(Boolean)
    .join(' ')
    .toLowerCase();

  const wantsFormD =
    /\bform[\s._-]*d\b/.test(blob) ||
    /form[\s._-]*d[\s._-]*gj|form_d_gj/.test(blob) ||
    blobIndicatesFormDGJMusterRoll(blob);

  if (!wantsFormD) return hints.preferredSheetName || names[0];

  let best = hints.preferredSheetName && names.includes(hints.preferredSheetName)
    ? hints.preferredSheetName
    : names[0];
  let bestScore = scoreFormDGJGujaratSheet(best, buildSheetTextBlob(workbook, best), blob);

  for (const name of names) {
    const score = scoreFormDGJGujaratSheet(name, buildSheetTextBlob(workbook, name), blob);
    if (score > bestScore) {
      bestScore = score;
      best = name;
    }
  }
  return bestScore > 0 ? best : hints.preferredSheetName || names[0];
}

function getSheetMergedCellText(rows, merges, r, c) {
  const direct = String((rows[r] || [])[c] ?? '').trim();
  if (direct) return direct;
  for (const m of merges || []) {
    if (!m?.s || !m?.e) continue;
    if (r >= m.s.r && r <= m.e.r && c >= m.s.c && c <= m.e.c) {
      return String((rows[m.s.r] || [])[m.s.c] ?? '').trim();
    }
  }
  return '';
}

export function resolveFormDGJGujaratTableLayout(workbook, hints = {}) {
  const names = Array.isArray(workbook?.SheetNames) ? workbook.SheetNames : [];
  if (!names.length) return null;

  const sheetName = resolveFormDGJGujaratWorkbookSheetName(workbook, hints);
  if (!sheetName) return null;

  const ws = workbook?.Sheets?.[sheetName];
  if (!ws) return null;

  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  const merges = ws['!merges'] || [];
  const maxScan = Math.min(rows.length, 90);
  let best = null;
  let bestScore = -1;

  for (let r = 0; r < maxScan; r += 1) {
    const cells = [];
    const maxCols = Math.max((rows[r] || []).length, (rows[r + 1] || []).length, 16);
    for (let c = 0; c < maxCols; c += 1) {
      const text = getSheetMergedCellText(rows, merges, r, c);
      if (text) cells.push({ c, text, lower: text.toLowerCase() });
    }
    if (cells.length < 2) continue;

    const nextRowCells = [];
    for (let c = 0; c < maxCols; c += 1) {
      const text = getSheetMergedCellText(rows, merges, r + 1, c);
      if (text) nextRowCells.push(text.toLowerCase());
    }
    const rowText = cells.map((x) => x.lower).join(' ');
    const combinedText = `${rowText} ${nextRowCells.join(' ')}`.trim();

    let score = 0;
    if (/relay\s+or\s+set/.test(combinedText)) score += 80;
    if (/summary\s+no\.?\s*of\s+days|summery\s+no\.?\s*of\s+days/.test(combinedText)) score += 60;
    if (/register\s+keeper|signature/.test(combinedText)) score += 40;
    if (/remarks.*hours|no\.?\s*of\s+hours/.test(combinedText)) score += 30;
    if (/sr\.?\s*no.*register/.test(combinedText)) score += 35;
    if (/\bname\b/.test(combinedText) && !/surname/.test(combinedText)) score += 20;
    if (/s\.?\s*no|serial|sl\.?\s*no/.test(combinedText)) score += 15;

    if (/surname/.test(combinedText) && /\bgender\b/.test(combinedText) && /education/.test(combinedText)) {
      score -= 150;
    }
    if (/employees?\s*\/\s*workme|worker\s+code/.test(combinedText) && /surname/.test(combinedText)) {
      score -= 120;
    }
    if (/register\s+of\s+wages|rate\s+of\s+wage/.test(combinedText)) score -= 80;
    if (/register\s+of\s+deductions|damage\s+or\s+loss/.test(combinedText)) score -= 80;
    if (/for\s+the\s+period\s+from/.test(combinedText) && !/relay\s+or\s+set/.test(combinedText)) {
      score -= 200;
    }

    score += cells.length * 2;
    if (score < 15) continue;

    let startCol = cells[0].c;
    for (const cell of cells) {
      if (/^s\.?\s*no\.?$|^sl\.?\s*no\.?$|serial/i.test(cell.lower)) {
        startCol = cell.c;
        break;
      }
    }

    const headers = [];
    for (let c = startCol; c < maxCols; c += 1) {
      let h = getSheetMergedCellText(rows, merges, r, c);
      if (!h) h = getSheetMergedCellText(rows, merges, r + 1, c);
      if (!h && r > 0) h = getSheetMergedCellText(rows, merges, r - 1, c);
      h = String(h || '').replace(/\s+/g, ' ').trim();
      if (!h) {
        if (headers.length > 0) break;
        continue;
      }
      if (/^\d+$/.test(h) && headers.length === 0) continue;
      headers.push(h);
    }
    if (headers.length < 3) continue;

    let headerRowIndex = r;
    let dataStartIndex = r + 1;
    const probeRow = rows[dataStartIndex] || [];
    if (
      headers.some((h) => /sr\.?\s*no|serial|sl\.?\s*no/i.test(String(h))) &&
      probeRow.length > 0 &&
      !/^\d+$/.test(String(probeRow[startCol] ?? '').trim())
    ) {
      dataStartIndex = r + 2;
    }

    const candidate = {
      sheetName,
      headerRowIndex,
      dataStartIndex,
      tableStartCol: startCol,
      headers,
    };
    if (score > bestScore) {
      bestScore = score;
      best = candidate;
    }
  }

  return best;
}

export function resolveFormDGJGujaratTableHeaders(tableHeaders) {
  const parsed = Array.isArray(tableHeaders)
    ? tableHeaders.map((h) => String(h || '').trim()).filter(Boolean)
    : [];
  if (parsed.length >= 3 && headersIndicateFormDGJMusterTable(parsed)) return parsed;
  return parsed.length > 0 ? parsed : [...FORM_DGJ_GJ_CANONICAL_TABLE_HEADERS];
}

export function repairFormDGJGujaratTableHeadersFromWorkbook(workbook, hints = {}) {
  if (!workbook) return null;
  const layout = resolveFormDGJGujaratTableLayout(workbook, {
    preferredSheetName: hints.preferredSheetName || hints.sheetName || '',
    fileName: hints.fileName || '',
    formFileName: hints.formFileName || '',
    formName: hints.formName || '',
    item: hints.item || null,
    formHeader: hints.formHeader || hints.parsedFormHeader || null,
    formHeaderTitle: hints.formHeaderTitle || hints.formHeader?.title || '',
  });
  if (!layout?.headers?.length) return null;
  return layout;
}

export function enrichFormDGJGujaratDisplayHeader(formHeader, fileName, item, tableHeaders, sheetText) {
  const base = formHeader && typeof formHeader === 'object' ? { ...formHeader } : {};
  const parts = [item?.formName, item?.FormName, fileName, base.title, base.subtitle, sheetText]
    .filter(Boolean)
    .join(' ')
    .toLowerCase();
  const isFormD =
    isFormDGJGujaratContext(base, item, fileName, sheetText, tableHeaders) ||
    (/\bform[\s._-]*d\b/.test(parts) && !/rajasthan|\b_rj\b|form_d_rj/.test(parts));

  if (!isFormD) return base;

  const fields = enrichEstablishmentPrincipalEmployerHeaderFields(base.fields || []);
  const hasPeriod = fields.some((f) => /for\s+the\s+period\s+from/i.test(String(f?.label || '')));
  if (!hasPeriod) {
    fields.push({ label: 'For the period From', value: '', key: 'form_d_gj_period' });
  }

  return {
    ...base,
    fields,
    formVIIAPHeaderFieldLayout: false,
    formXXVIAPHeaderFieldLayout: false,
    formXIXAPHeaderFieldLayout: false,
    title:
      /\bform[\s._-]*d\b/i.test(String(base.title || '')) ||
      /form[\s._-]*d[\s._-]*gj/i.test(parts)
        ? base.title || 'FORM D'
        : 'FORM D',
    subtitle:
      /muster|relay\s+or\s+set|register\s+keeper/i.test(String(base.subtitle || ''))
        ? base.subtitle
        : 'Muster Roll (Relay or Set Work)',
  };
}

const stripLeadingNumber = (s) =>
  String(s || '')
    .replace(/^\(?\d+\)?[\.\)]\s*/, '')
    .trim();

export function parseFormDGJHeaderLabel(h) {
  const raw = stripLeadingNumber(h);
  const m = String(raw || '').match(/^(.+)_([\s\S]+)$/);
  if (m) return String(m[2] || '').trim();
  return raw;
}

export function normDGJHeader(h) {
  return formDGJGujaratHeaderNorm(parseFormDGJHeaderLabel(h))
    .replace(/[^a-z0-9 ]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function unwrapEmployeeRecord(emp) {
  return emp?.Employee || emp?.employee || emp;
}

function pickEmployeeValue(emp, keys) {
  const src = unwrapEmployeeRecord(emp);
  if (!src || typeof src !== 'object') return '';
  for (let i = 0; i < keys.length; i += 1) {
    const raw = src[keys[i]];
    if (raw == null || raw === '') continue;
    if (typeof raw === 'object') {
      const nested = String(raw.displayValue ?? raw.name ?? raw.Name ?? '').trim();
      if (nested) return nested;
      continue;
    }
    const text = String(raw).trim();
    if (text) return text;
  }
  return '';
}

export function readFormDGJEmployeeFullName(emp) {
  const fn = pickEmployeeValue(emp, ['FirstName', 'First Name', 'firstName']);
  const ln = pickEmployeeValue(emp, ['LastName', 'Last Name', 'lastName']);
  if (fn && ln) return `${fn} ${ln}`;
  return fn || ln || '';
}

export function readFormDGJDesignation(emp) {
  return pickEmployeeValue(emp, [
    'Designation',
    'designation',
    'Designation.displayValue',
    'Department',
    'department',
  ]);
}

export function isFormDGJNameHeader(h) {
  const s = normDGJHeader(h);
  if (!/\bname\b/.test(s)) return false;
  if (/father|spouse|husband|employer|establishment|person|presence|explanation|bank|register\s+keeper|signature/.test(s)) {
    return false;
  }
  return s === 'name' || /^name\b/.test(s);
}

export function isFormDGJRelayOrSetWorkHeader(h) {
  const s = normDGJHeader(h);
  return /relay\s+or\s+set/.test(s) || (s.includes('relay') && s.includes('set') && s.includes('work'));
}

/** Summary / Summery No. of Days ← Sample Payroll Paid_days. */
export function isFormDGJSummaryNoOfDaysHeader(h) {
  const s = normDGJHeader(h);
  if (!s) return false;
  // Exclude neighbouring columns (use word boundaries — avoid matching "name" inside "workman").
  if (
    /\bremarks\b/.test(s) ||
    /\bhours?\b/.test(s) ||
    /\bsignature\b/.test(s) ||
    /register\s+keeper/.test(s) ||
    /relay\s+or\s+set/.test(s) ||
    /^(sr|sl|s)\s*no\b/.test(s) ||
    /\bserial\b/.test(s)
  ) {
    return false;
  }
  // Template / UI may spell "Summary" or "Summery".
  return (
    (/summary|summery/.test(s) && /\bdays?\b/.test(s)) ||
    /^no\.?\s*of\s+days$/.test(s)
  );
}

/** Remarks No. of Hours ← Summary No. of Days × 8. */
export function isFormDGJRemarksNoOfHoursHeader(h) {
  const s = normDGJHeader(h);
  if (!s) return false;
  if (/\bsignature\b/.test(s) || /register\s+keeper/.test(s) || /relay\s+or\s+set/.test(s)) {
    return false;
  }
  if (isFormDGJSummaryNoOfDaysHeader(h)) return false;
  return (
    (/\bremarks\b/.test(s) && /\bhours?\b/.test(s)) ||
    /^no\.?\s*of\s+hours$/.test(s) ||
    /remarks\s+no\.?\s*(of\s+)?hours?/.test(s)
  );
}

/** Remarks hours = Summary / Paid_days × 8. */
export function computeFormDGJGujaratRemarksHours(paidDays) {
  const n = Number(String(paidDays ?? '').replace(/,/g, '').trim());
  if (!Number.isFinite(n) || n < 0) return '';
  return String(Math.round(n * 8 * 100) / 100);
}

/** Read Paid_days (and aliases) from a Sample Payroll / Payroll table row. */
export function readFormDGJGujaratPaidDays(payrollRow) {
  if (!payrollRow || typeof payrollRow !== 'object' || payrollRow.fetch_error) return '';
  const flat = flattenPayrollEarningColumns(payrollRow);
  const keys = [
    'Paid_days',
    'paid_days',
    'Paid Days',
    'paidDays',
    'PaidDays',
    'days_worked',
    'Days Worked',
    'daysWorked',
    'no_of_days_worked',
    'effective_paid_days',
  ];
  const patterns = [
    /^paid_days$/,
    /^paiddays$/,
    /paid_days/,
    /daysworked/,
    /days_present/,
    /noofdayspresent/,
    /no_of_days_present/,
    /effective_paid_days/,
  ];
  const fromFlat = readPayrollScalar({ ...flat, ...payrollRow }, keys, patterns);
  if (fromFlat !== '' && fromFlat != null) {
    const num = Number(String(fromFlat).replace(/,/g, '').trim());
    if (Number.isFinite(num) && num >= 0) return String(num);
  }
  return '';
}

/**
 * Fill Summary No. of Days from Sample Payroll Paid_days,
 * and Remarks No. of Hours = Summary No. of Days × 8.
 * Match payroll by FirstName AND LastName only (never employee ID / first-name-only).
 * @returns {number} rows updated
 */
export function applyFormDGJGujaratPaidDaysToRows(
  mappedData,
  employeesForMapping,
  tableHeaders,
  options = {}
) {
  if (!Array.isArray(mappedData) || mappedData.length === 0) return 0;
  const headers = Array.isArray(tableHeaders) ? tableHeaders : [];
  const daysHeaders = headers.filter((h) => isFormDGJSummaryNoOfDaysHeader(h));
  const hoursHeaders = headers.filter((h) => isFormDGJRemarksNoOfHoursHeader(h));
  if (daysHeaders.length === 0 && hoursHeaders.length === 0) return 0;

  const resolvePayrollRow =
    typeof options.resolvePayrollRow === 'function' ? options.resolvePayrollRow : null;
  const unwrapEmp =
    typeof options.unwrapEmp === 'function'
      ? options.unwrapEmp
      : (item) => (item && (item.Employee || item.employee || item)) || null;
  const sanitize =
    typeof options.sanitizeValue === 'function' ? options.sanitizeValue : (v) => String(v ?? '').trim();
  const overwrite = options.overwrite !== false;
  const payrollList = Array.isArray(options.payrollRows) ? options.payrollRows : [];

  const resolvePayrollForRow = (row, index) => {
    const emp = unwrapEmp(Array.isArray(employeesForMapping) ? employeesForMapping[index] : null);
    const extraParts = collectForm10RowNameParts(row, headers);
    if (resolvePayrollRow) {
      const payrollRow = resolvePayrollRow(emp, row, index);
      if (payrollRow && !payrollRow.fetch_error && readFormDGJGujaratPaidDays(payrollRow) !== '') {
        return payrollRow;
      }
    }
    const byName = findForm10PayrollRowByFirstAndLastName(emp || row, payrollList, extraParts);
    if (byName && readFormDGJGujaratPaidDays(byName) !== '') return byName;
    // No ID / first-name-only / loose fallbacks — FirstName AND LastName only.
    return null;
  };

  const setIfAllowed = (row, key, value) => {
    if (!key || value == null || value === '') return false;
    const cur = String(row[key] ?? '').trim();
    const isPlaceholder = !cur || /^enter\b/i.test(cur);
    if (!overwrite && cur && !isPlaceholder) return false;
    row[key] = sanitize(value);
    return true;
  };

  let hits = 0;
  mappedData.forEach((row, index) => {
    if (!row || typeof row !== 'object') return;
    const payrollRow = resolvePayrollForRow(row, index);
    let paidDays = readFormDGJGujaratPaidDays(payrollRow);
    // If payroll miss but Summary days already on the row, still compute hours.
    if (paidDays === '') {
      for (let i = 0; i < daysHeaders.length; i += 1) {
        const existing = String(row[daysHeaders[i]] ?? '').trim();
        if (existing && !/^enter\b/i.test(existing)) {
          paidDays = existing;
          break;
        }
      }
    }
    if (paidDays === '') return;
    const hours = computeFormDGJGujaratRemarksHours(paidDays);
    let changed = false;
    daysHeaders.forEach((daysHeader) => {
      if (setIfAllowed(row, daysHeader, paidDays)) changed = true;
    });
    hoursHeaders.forEach((hoursHeader) => {
      if (setIfAllowed(row, hoursHeader, hours)) changed = true;
    });
    // Also write onto any existing row keys that look like Summary days / Remarks hours.
    Object.keys(row).forEach((key) => {
      if (isFormDGJSummaryNoOfDaysHeader(key) && !daysHeaders.includes(key)) {
        if (setIfAllowed(row, key, paidDays)) changed = true;
      }
      if (isFormDGJRemarksNoOfHoursHeader(key) && !hoursHeaders.includes(key)) {
        if (setIfAllowed(row, key, hours)) changed = true;
      }
    });
    if (changed) hits += 1;
  });
  return hits;
}

export function applyFormDGJGujaratEmployeeToRow(row, emp, headers, helpers = {}) {
  const hdrs = Array.isArray(headers) ? headers : [];
  const out = row && typeof row === 'object' ? { ...row } : {};
  const {
    sanitizeValue = (v) => String(v ?? '').trim(),
    overwrite = true,
    payrollRow = null,
  } = helpers;

  const cellIsEmpty = (header) => {
    const v = String(out[header] || '').trim();
    if (!v) return true;
    const s = v.toLowerCase();
    return /^enter\b/.test(s) || s.includes('enter ');
  };
  const setCell = (header, value) => {
    if (!header || value == null || value === '') return;
    if (!overwrite && !cellIsEmpty(header)) return;
    out[header] = sanitizeValue(value);
  };

  const fullName = readFormDGJEmployeeFullName(emp);
  const designation = readFormDGJDesignation(emp);
  if (fullName) out.__employeeLookupName = fullName;
  const paidDays = readFormDGJGujaratPaidDays(payrollRow);
  const remarksHours = computeFormDGJGujaratRemarksHours(paidDays);

  hdrs.forEach((header) => {
    if (isFormDGJNameHeader(header)) {
      setCell(header, fullName);
      return;
    }
    if (isFormDGJRelayOrSetWorkHeader(header)) {
      setCell(header, designation);
      return;
    }
    if (isFormDGJSummaryNoOfDaysHeader(header)) {
      setCell(header, paidDays);
      return;
    }
    if (isFormDGJRemarksNoOfHoursHeader(header)) {
      setCell(header, remarksHours);
    }
  });

  return out;
}

const MONTH_NAMES_DGJ = [
  'January',
  'February',
  'March',
  'April',
  'May',
  'June',
  'July',
  'August',
  'September',
  'October',
  'November',
  'December',
];

export function buildFormDGJGujaratPeriodLine(monthName, year) {
  let idx = MONTH_NAMES_DGJ.findIndex(
    (m) => m.toLowerCase() === String(monthName || '').toLowerCase().trim()
  );
  if (idx < 0) {
    const token = String(monthName || '').toLowerCase().trim().slice(0, 3);
    idx = MONTH_NAMES_DGJ.findIndex((m) => m.toLowerCase().startsWith(token));
  }
  if (idx < 0) idx = new Date().getMonth();
  const y = Number(year);
  if (!Number.isFinite(y) || y < 1900) return '';
  const m = idx + 1;
  const lastDay = new Date(y, idx + 1, 0).getDate();
  const pad = (n) => String(n).padStart(2, '0');
  return `For the period From ${pad(1)}-${pad(m)}-${y} to ${pad(lastDay)}-${pad(m)}-${y}`;
}

export function prepareFormDGJGujaratDownloadHeaderData(
  headerFormData,
  parsedFormHeader = null,
  siteContext = {}
) {
  const out = headerFormData && typeof headerFormData === 'object' ? { ...headerFormData } : {};
  const {
    establishmentText = '',
    principalEmployerText = '',
    periodText = '',
  } = siteContext;

  if (establishmentText) {
    out.statutory_establishment_name = establishmentText;
    out.form_d_gj_establishment = establishmentText;
  }
  if (principalEmployerText) {
    out.statutory_principal_employer = principalEmployerText;
    out.form_d_gj_principal_employer = principalEmployerText;
  }
  if (periodText) {
    out.form_d_gj_period = periodText;
    out.statutory_period_from = periodText;
  }

  const fields = Array.isArray(parsedFormHeader?.fields) ? parsedFormHeader.fields : [];
  fields.forEach((field) => {
    const key = field?.key;
    if (!key || String(out[key] ?? '').trim()) return;
    if (/establishment/i.test(String(field.label || '')) && !/principal|employer|contractor/.test(String(field.label || ''))) {
      if (establishmentText) out[key] = establishmentText;
    } else if (/principal\s+employer/i.test(String(field.label || ''))) {
      if (principalEmployerText) out[key] = principalEmployerText;
    } else if (/for\s+the\s+period\s+from/i.test(String(field.label || ''))) {
      if (periodText) out[key] = periodText;
    }
  });

  return out;
}

function writeFormDGJGujaratPeriodCell(worksheet, periodText, headerRowEnd = 25) {
  const line = String(periodText || '').trim();
  if (!line || !worksheet) return;
  for (let r = 1; r <= headerRowEnd; r += 1) {
    for (let c = 1; c <= 24; c += 1) {
      const raw = excelCellValueToString(worksheet.getCell(r, c)?.value).trim();
      if (/for\s+the\s+period\s+from/i.test(raw)) {
        worksheet.getCell(r, c).value = line;
        return;
      }
    }
  }
}

export function writeFormDGJGujaratHeaderFieldsToWorksheet(
  worksheet,
  headerFormData,
  parsedFormHeader,
  helpers = {}
) {
  if (!worksheet || !headerFormData || typeof headerFormData !== 'object') return;
  const headerRowEnd = Math.max(1, Number(helpers.headerRowEnd) || 25);
  writeStatutoryHeaderFieldsToExcelJsWorksheet(worksheet, {
    headerFormData,
    parsedFormHeader,
    headerRowEnd,
    maxScanCols: 80,
    writeMode: 'both',
  });
  writeFormDGJGujaratPeriodCell(
    worksheet,
    headerFormData.form_d_gj_period || headerFormData.statutory_period_from,
    headerRowEnd
  );
  // First 5 meta fields: plain text only (strip template box borders).
  clearFormDGJGujaratHeaderMetaBoxes(worksheet, {
    headerRowEnd,
    colTo: Math.max(6, Number(helpers.colTo) || 12),
  });
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

function excelCellLooksLikeFormDGJTableHeader(text) {
  const t = formDGJGujaratHeaderNorm(text).replace(/[^a-z0-9 ]/g, ' ').replace(/\s+/g, ' ').trim();
  return (
    /^sr\.?\s*no|^s\.?\s*no|serial|sl\.?\s*no/.test(t) ||
    /relay\s+or\s+set/.test(t) ||
    (t === 'name' || /^name\b/.test(t))
  );
}

function detectFormDGJGujaratTableLayout(worksheet, hints = {}) {
  const getMergeTopLeft = buildMergeTopLeftResolver(worksheet);
  const getMergedAwareCellText = (r, c) => {
    const tl = getMergeTopLeft(r, c);
    return excelCellValueToString(worksheet.getCell(tl.r, tl.c)?.value).trim();
  };

  let headerRow =
    hints.parsedHeaderRowIndex != null && hints.parsedHeaderRowIndex >= 0
      ? hints.parsedHeaderRowIndex + 1
      : -1;
  let startCol =
    hints.parsedTableStartCol != null && hints.parsedTableStartCol >= 0
      ? hints.parsedTableStartCol + 1
      : 1;

  if (headerRow < 1) {
    const maxScanRows = Math.max(40, worksheet.rowCount + 5);
    for (let r = 1; r <= maxScanRows; r += 1) {
      let foundRelay = false;
      for (let c = 1; c <= 20; c += 1) {
        const text = getMergedAwareCellText(r, c);
        if (/relay\s+or\s+set/i.test(text)) {
          headerRow = r;
          foundRelay = true;
          break;
        }
      }
      if (foundRelay) break;
      for (let c = 1; c <= 20; c += 1) {
        if (excelCellLooksLikeFormDGJTableHeader(getMergedAwareCellText(r, c))) {
          headerRow = r;
          startCol = c;
          break;
        }
      }
      if (headerRow > 0) break;
    }
  }

  if (headerRow < 1) return null;

  if (startCol === 1) {
    for (let c = 1; c <= 20; c += 1) {
      if (excelCellLooksLikeFormDGJTableHeader(getMergedAwareCellText(headerRow, c))) {
        startCol = c;
        break;
      }
    }
  }

  const templateCols = [];
  for (let c = startCol; c <= startCol + 20; c += 1) {
    const label = getMergedAwareCellText(headerRow, c);
    if (!label) {
      if (templateCols.length >= 4) break;
      continue;
    }
    templateCols.push({ col: c, label });
    if (templateCols.length >= 12) break;
  }
  if (templateCols.length < 3) return null;

  const dataStartRow =
    hints.parsedDataStartIndex != null && hints.parsedDataStartIndex >= 0
      ? hints.parsedDataStartIndex + 1
      : headerRow + 1;

  return { headerRow, dataStartRow, templateCols, startCol };
}

export function getFormDGJGujaratRowValueForHeader(row, header, rowIndex = 0) {
  if (!row || !header) return '';
  const direct = row[header];
  if (direct != null && String(direct).trim() !== '') return String(direct).trim();
  if (isFormDGJNameHeader(header)) {
    for (const [k, v] of Object.entries(row)) {
      if (isFormDGJNameHeader(k) && v != null && String(v).trim() !== '') return String(v).trim();
    }
  }
  if (isFormDGJRelayOrSetWorkHeader(header)) {
    for (const [k, v] of Object.entries(row)) {
      if (isFormDGJRelayOrSetWorkHeader(k) && v != null && String(v).trim() !== '') {
        return String(v).trim();
      }
    }
  }
  if (isFormDGJSummaryNoOfDaysHeader(header)) {
    for (const [k, v] of Object.entries(row)) {
      if (isFormDGJSummaryNoOfDaysHeader(k) && v != null && String(v).trim() !== '') {
        return String(v).trim();
      }
    }
  }
  if (isFormDGJRemarksNoOfHoursHeader(header)) {
    for (const [k, v] of Object.entries(row)) {
      if (isFormDGJRemarksNoOfHoursHeader(k) && v != null && String(v).trim() !== '') {
        return String(v).trim();
      }
    }
  }
  if (/sr\.?\s*no|serial|sl\.?\s*no/i.test(normDGJHeader(header))) {
    return String(rowIndex + 1);
  }
  return '';
}

export function rowHasMeaningfulFormDGJGujaratExportData(row, headers) {
  const hdrs = resolveFormDGJGujaratTableHeaders(headers);
  return hdrs.some((header) => {
    if (/sr\.?\s*no|serial|sl\.?\s*no/i.test(normDGJHeader(header))) return false;
    return getFormDGJGujaratRowValueForHeader(row, header) !== '';
  });
}

export function filterFormDGJGujaratExportRows(rows, headers) {
  const hdrs = resolveFormDGJGujaratTableHeaders(headers);
  return (Array.isArray(rows) ? rows : []).filter((row) =>
    rowHasMeaningfulFormDGJGujaratExportData(row, hdrs)
  );
}

export function remapFormDGJGujaratRowsToHeaders(rows, sourceHeaders, targetHeaders) {
  const src = Array.isArray(sourceHeaders) ? sourceHeaders : [];
  const tgt = resolveFormDGJGujaratTableHeaders(targetHeaders);
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
        for (const [k, v] of Object.entries(row)) {
          if (normDGJHeader(k) === normDGJHeader(targetHeader)) {
            val = v;
            break;
          }
        }
      }
      if ((val == null || val === '') && /sr\.?\s*no|serial|sl\.?\s*no/i.test(normDGJHeader(targetHeader))) {
        val = rowIndex + 1;
      }
      out[targetHeader] = val == null ? '' : val;
    });
    return out;
  });
}

function resolveFormDGJGujaratWorksheet(workbook, hints = {}) {
  if (!workbook) return null;
  const preferred = String(hints.preferredSheetName || '').trim();
  if (preferred) {
    const ws = workbook.getWorksheet(preferred);
    if (ws) return ws;
  }
  for (const ws of workbook.worksheets || []) {
    if (/\bform[\s._-]*d\b/i.test(String(ws.name || ''))) return ws;
  }
  return workbook.worksheets?.[0] || null;
}

/** Write employee rows into the Gujarat Form D template (FORM D sheet). */
export async function buildFormDGJGujaratWorkbookWithTemplateStyles({
  templateArrayBuffer,
  mappedData,
  headersToUse,
  parsedFormHeader,
  formFileName,
  headerFormData,
  parsedHeaderRowIndex,
  parsedDataStartIndex,
  parsedTableStartCol,
  preferredSheetName = '',
}) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(templateArrayBuffer);
  const worksheet = resolveFormDGJGujaratWorksheet(workbook, { preferredSheetName });
  if (!worksheet) throw new Error('Template worksheet not found.');

  const layout = detectFormDGJGujaratTableLayout(worksheet, {
    parsedHeaderRowIndex,
    parsedDataStartIndex,
    parsedTableStartCol,
  });
  if (!layout) throw new Error('Could not locate Gujarat Form D table header row.');

  const { dataStartRow, templateCols } = layout;

  writeFormDGJGujaratHeaderFieldsToWorksheet(worksheet, headerFormData, parsedFormHeader, {
    headerRowEnd: Math.max(1, layout.headerRow - 1),
  });

  const templateHeaderLabels = templateCols.map((entry) => entry.label);
  const normalizedHeaders = resolveFormDGJGujaratTableHeaders(
    Array.isArray(headersToUse) && headersToUse.length >= templateHeaderLabels.length
      ? headersToUse
      : templateHeaderLabels.length > 0
        ? templateHeaderLabels
        : headersToUse
  );
  const rows = filterFormDGJGujaratExportRows(
    remapFormDGJGujaratRowsToHeaders(
      Array.isArray(mappedData) ? mappedData : [],
      normalizedHeaders,
      normalizedHeaders
    ),
    normalizedHeaders
  );

  const tableColMin = Math.min(...templateCols.map(({ col }) => col));
  const tableColMax = Math.max(...templateCols.map(({ col }) => col));

  // Capture template box geometry before clearing values.
  const templateSystemRow = findFormDGJGujaratSystemGeneratedRow(
    worksheet,
    dataStartRow,
    tableColMin,
    tableColMax
  );
  const templateBodyRows = countExcelJSTemplateBodyRows(
    worksheet,
    dataStartRow,
    tableColMin,
    tableColMax,
    80
  );
  const templateBoxEndBeforeWrite = findFormDGJGujaratTableBoxEndRow(
    worksheet,
    dataStartRow,
    tableColMin,
    tableColMax,
    { systemGeneratedRow: templateSystemRow }
  );

  const lastDataRow =
    rows.length > 0 ? dataStartRow + rows.length - 1 : Math.max(1, dataStartRow - 1);
  // Keep empty bordered rows under the data so the outer box matches the model.
  let tableBoxEndRow = Math.max(
    templateBoxEndBeforeWrite,
    lastDataRow + 2,
    dataStartRow + Math.max(templateBodyRows, rows.length + 2) - 1
  );
  if (templateSystemRow && templateSystemRow > dataStartRow) {
    tableBoxEndRow = Math.min(tableBoxEndRow, templateSystemRow - 1);
  }
  tableBoxEndRow = Math.max(tableBoxEndRow, lastDataRow);

  // Clear employee values inside the box only (keep structure for borders).
  for (let r = dataStartRow; r <= tableBoxEndRow; r += 1) {
    for (let c = tableColMin; c <= tableColMax; c += 1) {
      const text = excelCellValueToString(worksheet.getCell(r, c)?.value).trim();
      if (isFormDGJGujaratSystemGeneratedText(text) || isFormDGJGujaratOuterFootnoteText(text)) {
        worksheet.getCell(r, c).value = '';
        continue;
      }
      worksheet.getCell(r, c).value = '';
    }
  }

  // Clear old System Generated / footnotes that sat below the template box.
  const cleanupEnd = Math.max(
    worksheet.rowCount || 0,
    (templateSystemRow || tableBoxEndRow) + 20,
    tableBoxEndRow + 20
  );
  for (let r = tableBoxEndRow + 1; r <= cleanupEnd; r += 1) {
    for (let c = 1; c <= Math.max(tableColMax + 8, 12); c += 1) {
      const text = formDGJExcelCellText(worksheet, r, c);
      if (
        isFormDGJGujaratSystemGeneratedText(text) ||
        isFormDGJGujaratOuterFootnoteText(text)
      ) {
        clearFormDGJRowOutsideBox(worksheet, r, 1, Math.max(tableColMax + 8, 12));
        break;
      }
    }
  }

  rows.forEach((row, idx) => {
    const targetRow = worksheet.getRow(dataStartRow + idx);
    if (targetRow) targetRow.height = 18;
    templateCols.forEach(({ col, label }) => {
      const val = getFormDGJGujaratRowValueForHeader(row, label, idx);
      const cell = worksheet.getCell(dataStartRow + idx, col);
      if (val == null || val === '') {
        cell.value = '';
        return;
      }
      if (/sr\.?\s*no|serial|sl\.?\s*no/i.test(normDGJHeader(label))) {
        const n = Number(String(val).replace(/[,]/g, '').trim());
        cell.value = Number.isFinite(n) ? n : String(val);
      } else {
        cell.value = String(val);
      }
      cell.alignment = { ...(cell.alignment || {}), vertical: 'middle', wrapText: true };
    });
  });

  // Continuous full box from table header through empty body rows (model layout).
  const borderFromRow = Math.max(1, layout.headerRow || dataStartRow);
  applyExcelJSFullBoxBordersToRange(worksheet, {
    rowFrom: borderFromRow,
    rowTo: tableBoxEndRow,
    colFrom: tableColMin,
    colTo: tableColMax,
    borderStyle: 'thin',
  });
  if (rows.length > 0) {
    ensureExcelJSDataRowsWithBorders(worksheet, {
      dataStartRow,
      dataRowCount: Math.max(rows.length, tableBoxEndRow - dataStartRow + 1),
      colFrom: tableColMin,
      colTo: tableColMax,
      templateRow: dataStartRow,
      templateBodyRows: Math.max(1, templateBodyRows),
    });
  }

  // Outside the box (same as PDF model): System Generated, then legal footnotes.
  const systemNoteRow = writeFormDGJGujaratSystemGeneratedNote(worksheet, {
    afterRow: tableBoxEndRow,
    startCol: tableColMin,
    endCol: tableColMax,
  });
  writeFormDGJGujaratOuterFootnotes(worksheet, {
    afterRow: systemNoteRow || tableBoxEndRow,
    startCol: tableColMin,
    endCol: tableColMax,
  });

  const out = await workbook.xlsx.writeBuffer();
  const fileName =
    formFileName ||
    parsedFormHeader?.title?.replace(/[^a-zA-Z0-9]/g, '_') ||
    'Form_D_GJ_Gujarat.xlsx';
  return {
    blob: new Blob([out], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    }),
    fileName,
  };
}
