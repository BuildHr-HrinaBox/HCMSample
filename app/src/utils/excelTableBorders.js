/**
 * ExcelJS helpers — copy "all borders" table styling from a template data row
 * onto every populated data row in statutory form exports.
 */

export function excelJSCellHasBorder(cell) {
  const b = cell?.border;
  if (!b) return false;
  return !!(b.top?.style || b.bottom?.style || b.left?.style || b.right?.style);
}

export function excelJSCellHasFullBoxBorder(cell) {
  const b = cell?.border;
  if (!b) return false;
  return !!(b.top?.style && b.bottom?.style && b.left?.style && b.right?.style);
}

function cloneBorderSide(side) {
  if (!side || !side.style) return undefined;
  const out = { style: side.style };
  if (side.color) out.color = { ...side.color };
  return out;
}

export function cloneExcelJSBorder(border) {
  if (!border) return null;
  const out = {};
  const top = cloneBorderSide(border.top);
  const left = cloneBorderSide(border.left);
  const bottom = cloneBorderSide(border.bottom);
  const right = cloneBorderSide(border.right);
  if (top) out.top = top;
  if (left) out.left = left;
  if (bottom) out.bottom = bottom;
  if (right) out.right = right;
  if (border.diagonal?.style) out.diagonal = cloneBorderSide(border.diagonal);
  return Object.keys(out).length > 0 ? out : null;
}

function cloneExcelJSStyle(style) {
  if (!style) return null;
  try {
    return JSON.parse(JSON.stringify(style));
  } catch (_) {
    return { ...style };
  }
}

export const EXCELJS_THIN_BLACK_BOX_BORDER = {
  top: { style: 'thin', color: { argb: 'FF000000' } },
  left: { style: 'thin', color: { argb: 'FF000000' } },
  bottom: { style: 'thin', color: { argb: 'FF000000' } },
  right: { style: 'thin', color: { argb: 'FF000000' } }
};

export const EXCELJS_MEDIUM_BLACK_BOX_BORDER = {
  top: { style: 'medium', color: { argb: 'FF000000' } },
  left: { style: 'medium', color: { argb: 'FF000000' } },
  bottom: { style: 'medium', color: { argb: 'FF000000' } },
  right: { style: 'medium', color: { argb: 'FF000000' } }
};

const DEFAULT_ALL_BORDERS = EXCELJS_THIN_BLACK_BOX_BORDER;

function buildExcelJSFullBoxBorder(borderStyle = 'thin') {
  const style = borderStyle === 'medium' ? 'medium' : 'thin';
  return {
    top: { style, color: { argb: 'FF000000' } },
    left: { style, color: { argb: 'FF000000' } },
    bottom: { style, color: { argb: 'FF000000' } },
    right: { style, color: { argb: 'FF000000' } }
  };
}

/** Apply a full box border to every cell in [rowFrom..rowTo] × [colFrom..colTo] (1-based). */
export function applyExcelJSFullBoxBordersToRange(
  worksheet,
  { rowFrom, rowTo, colFrom, colTo, borderStyle = 'thin' } = {}
) {
  if (!worksheet) return;
  const r0 = Math.max(1, Number(rowFrom) || 1);
  const r1 = Math.max(r0, Number(rowTo) || r0);
  const c0 = Math.max(1, Number(colFrom) || 1);
  const c1 = Math.max(c0, Number(colTo) || c0);
  for (let r = r0; r <= r1; r += 1) {
    for (let c = c0; c <= c1; c += 1) {
      const cell = worksheet.getCell(r, c);
      // Fresh border object every cell — ExcelJS shares style refs if reused.
      const box = buildExcelJSFullBoxBorder(borderStyle);
      // Set both ways so borders survive ExcelJS writeBuffer after SheetJS round-trips.
      cell.border = box;
      try {
        const prev = cell.style && typeof cell.style === 'object' ? { ...cell.style } : {};
        cell.style = { ...prev, border: box };
      } catch (_) {
        /* border property above is enough for most ExcelJS versions */
      }
    }
  }
}

/** First row in [rowFrom, rowTo] that has at least one bordered cell. */
export function findExcelJSTemplateBorderRow(worksheet, rowFrom, rowTo, colFrom, colTo) {
  const r0 = Math.max(1, rowFrom);
  const r1 = Math.max(r0, rowTo);
  const c0 = Math.max(1, colFrom);
  const c1 = Math.max(c0, colTo);
  for (let r = r0; r <= r1; r += 1) {
    for (let c = c0; c <= c1; c += 1) {
      if (excelJSCellHasBorder(worksheet.getCell(r, c))) return r;
    }
  }
  return r0;
}

/** Infer table column span from bordered cells or header row text. */
export function inferExcelJSTableColumnRange(worksheet, { headerRow, dataStartRow, colScanMax = 80 } = {}) {
  const cMax = Math.max(1, colScanMax);
  let colFrom = cMax;
  let colTo = 0;
  const r0 = Math.max(1, dataStartRow || headerRow || 1);
  const r1 = r0 + 8;
  for (let r = r0; r <= r1; r += 1) {
    for (let c = 1; c <= cMax; c += 1) {
      const cell = worksheet.getCell(r, c);
      if (excelJSCellHasBorder(cell)) {
        colFrom = Math.min(colFrom, c);
        colTo = Math.max(colTo, c);
      }
    }
  }
  if (colTo > 0) return { colFrom: colFrom || 1, colTo };
  const hr = Math.max(1, headerRow || r0);
  for (let c = 1; c <= cMax; c += 1) {
    const v = cellText(worksheet.getCell(hr, c)?.value);
    if (v) {
      colFrom = Math.min(colFrom, c);
      colTo = Math.max(colTo, c);
    }
  }
  return { colFrom: colFrom <= cMax ? colFrom : 1, colTo: colTo || cMax };
}

function cellText(val) {
  if (val == null) return '';
  if (typeof val === 'object' && val.richText) return val.richText.map((t) => t.text || '').join('');
  if (typeof val === 'object' && val.text != null) return String(val.text);
  return String(val).trim();
}

/**
 * Apply thin box borders to every cell in the data table range.
 * Always forces a full box on each cell so rows beyond the template body stay bordered.
 */
export function applyExcelJSDataRowBorders(
  worksheet,
  { dataStartRow, dataRowCount, colFrom, colTo, templateRow = null, forceFullBox = true } = {}
) {
  if (!worksheet || !dataRowCount || dataRowCount < 1) return;

  const r0 = Math.max(1, dataStartRow);
  const r1 = r0 + dataRowCount - 1;
  let c0 = Math.max(1, colFrom);
  let c1 = Math.max(c0, colTo);
  if (!colFrom || !colTo) {
    const inferred = inferExcelJSTableColumnRange(worksheet, {
      headerRow: templateRow != null ? templateRow : r0 - 1,
      dataStartRow: r0
    });
    c0 = inferred.colFrom;
    c1 = inferred.colTo;
  }

  const styleRow =
    templateRow != null && templateRow >= 1
      ? templateRow
      : findExcelJSTemplateBorderRow(worksheet, r0, r1, c0, c1);

  let templateHasStyle = false;
  for (let c = c0; c <= c1; c += 1) {
    const src = worksheet.getCell(styleRow, c);
    if (src?.style && (excelJSCellHasBorder(src) || src.style.font || src.style.alignment)) {
      templateHasStyle = true;
      break;
    }
  }

  for (let r = r0; r <= r1; r += 1) {
    for (let c = c0; c <= c1; c += 1) {
      const dst = worksheet.getCell(r, c);
      const src = worksheet.getCell(styleRow, c);

      if (templateHasStyle && src?.style) {
        try {
          const cloned = cloneExcelJSStyle(src.style);
          if (cloned) {
            dst.style = cloned;
          }
        } catch (_) {
          /* keep existing style */
        }
      }

      if (forceFullBox) {
        // Fresh border object per cell — ExcelJS shares style refs if reused.
        dst.border = {
          top: { style: 'thin', color: { argb: 'FF000000' } },
          left: { style: 'thin', color: { argb: 'FF000000' } },
          bottom: { style: 'thin', color: { argb: 'FF000000' } },
          right: { style: 'thin', color: { argb: 'FF000000' } }
        };
      } else if (!excelJSCellHasFullBoxBorder(dst)) {
        if (excelJSCellHasBorder(src)) {
          const border = cloneExcelJSBorder(src.border);
          if (border) {
            dst.border = {
              top: border.top ? { ...border.top } : undefined,
              left: border.left ? { ...border.left } : undefined,
              bottom: border.bottom ? { ...border.bottom } : undefined,
              right: border.right ? { ...border.right } : undefined
            };
          } else {
            dst.border = {
              top: { style: 'thin', color: { argb: 'FF000000' } },
              left: { style: 'thin', color: { argb: 'FF000000' } },
              bottom: { style: 'thin', color: { argb: 'FF000000' } },
              right: { style: 'thin', color: { argb: 'FF000000' } }
            };
          }
        } else {
          dst.border = {
            top: { style: 'thin', color: { argb: 'FF000000' } },
            left: { style: 'thin', color: { argb: 'FF000000' } },
            bottom: { style: 'thin', color: { argb: 'FF000000' } },
            right: { style: 'thin', color: { argb: 'FF000000' } }
          };
        }
      }
    }
  }
}

/**
 * Apply all-borders styling to every populated data row.
 * Does not insert/duplicate worksheet rows — callers write values directly to
 * dataStartRow … dataStartRow + dataRowCount - 1. (duplicateRow after write
 * would clone the first employee row across the sheet.)
 */
export function ensureExcelJSDataRowsWithBorders(
  worksheet,
  { dataStartRow, dataRowCount, colFrom, colTo, templateRow = null, templateBodyRows = 1 } = {}
) {
  if (!worksheet || !dataRowCount || dataRowCount < 1) return;

  const startRow = Math.max(1, dataStartRow);
  void templateBodyRows;

  applyExcelJSDataRowBorders(worksheet, {
    dataStartRow: startRow,
    dataRowCount,
    colFrom,
    colTo,
    templateRow: templateRow != null ? templateRow : startRow,
    forceFullBox: true
  });
}

/**
 * After ExcelJS removeWorksheet(), workbook.views often keep activeTab/firstSheet
 * pointing at a removed sheet index. Excel then opens the file as [Repaired]
 * (Form XXIII AP multi-tab templates are a common case).
 */
export function resetExcelJsWorkbookActiveSheet(workbook) {
  if (!workbook) return;
  const sheetCount = Array.isArray(workbook.worksheets) ? workbook.worksheets.length : 0;
  if (sheetCount < 1) return;
  const maxIdx = Math.max(0, sheetCount - 1);
  const base =
    (Array.isArray(workbook.views) && workbook.views[0] && typeof workbook.views[0] === 'object'
      ? workbook.views[0]
      : null) || {
      x: 0,
      y: 0,
      width: 12000,
      height: 16000,
      visibility: 'visible',
    };
  let activeTab = Number(base.activeTab);
  let firstSheet = Number(base.firstSheet);
  if (!Number.isFinite(activeTab) || activeTab < 0 || activeTab > maxIdx) activeTab = 0;
  if (!Number.isFinite(firstSheet) || firstSheet < 0 || firstSheet > maxIdx) firstSheet = 0;
  workbook.views = [{ ...base, activeTab, firstSheet }];
}

/**
 * ExcelJS writeBuffer throws when a shared-formula clone's master was cleared or
 * overwritten ("Shared Formula master must exist above and or left of clone").
 * Convert orphaned clones to their cached result (or empty) so download can finish.
 */
export function detachOrphanedExcelJSSharedFormulas(worksheet) {
  if (!worksheet || typeof worksheet.eachRow !== 'function') return 0;
  let fixed = 0;
  worksheet.eachRow({ includeEmpty: false }, (row) => {
    row.eachCell({ includeEmpty: false }, (cell) => {
      const v = cell.value;
      if (!v || typeof v !== 'object' || Array.isArray(v)) return;
      // Clone: points at master address, no own formula text.
      if (!v.sharedFormula || v.formula) return;
      let masterOk = false;
      try {
        const master = worksheet.getCell(String(v.sharedFormula));
        const mv = master?.value;
        masterOk = !!(
          mv &&
          typeof mv === 'object' &&
          !Array.isArray(mv) &&
          (typeof mv.formula === 'string' || mv.shareType === 'shared')
        );
      } catch (_) {
        masterOk = false;
      }
      if (masterOk) return;
      cell.value = v.result != null && v.result !== undefined ? v.result : null;
      fixed += 1;
    });
  });
  return fixed;
}

/** Remove values and borders from cells to the right of the table (template stray boxes). */
export function clearExcelJSTrailingTableCells(
  worksheet,
  { dataStartRow, dataRowCount, afterCol, throughCol = null } = {}
) {
  if (!worksheet || !dataRowCount || dataRowCount < 1 || !afterCol) return;
  const r0 = Math.max(1, dataStartRow);
  const r1 = r0 + dataRowCount - 1;
  const c0 = afterCol + 1;
  const c1 = Math.max(c0, throughCol != null ? throughCol : afterCol + 40);
  for (let r = r0; r <= r1; r += 1) {
    for (let c = c0; c <= c1; c += 1) {
      const cell = worksheet.getCell(r, c);
      cell.value = null;
      cell.border = {};
    }
  }
}

/**
 * Form XXI AP: last real table column is Remarks (usually L). Used to strip M–W boxes.
 */
export function findExcelJSRemarksOrLastHeaderCol(
  worksheet,
  { headerRow = 1, startCol = 1, scanCols = 40 } = {}
) {
  if (!worksheet) return 0;
  const hr = Math.max(1, Number(headerRow) || 1);
  const c0 = Math.max(1, Number(startCol) || 1);
  const cMax = Math.max(c0, Number(scanCols) || 40);
  let remarksCol = 0;
  let lastLabelCol = 0;
  for (let r = Math.max(1, hr - 1); r <= hr + 2; r += 1) {
    for (let c = c0; c <= cMax; c += 1) {
      const raw = worksheet.getCell(r, c)?.value;
      let text = '';
      if (raw == null) text = '';
      else if (typeof raw === 'string' || typeof raw === 'number' || typeof raw === 'boolean') {
        text = String(raw).trim();
      } else if (typeof raw === 'object') {
        if (Array.isArray(raw.richText)) text = raw.richText.map((p) => p?.text || '').join('').trim();
        else if (raw.text != null) text = String(raw.text).trim();
      }
      if (!text) continue;
      lastLabelCol = Math.max(lastLabelCol, c);
      if (/^remarks?$/i.test(text)) remarksCol = Math.max(remarksCol, c);
    }
  }
  return remarksCol || lastLabelCol || 0;
}

/** Count contiguous template body rows that have borders or values in [colFrom, colTo]. */
export function countExcelJSTemplateBodyRows(worksheet, dataStartRow, colFrom, colTo, maxScan = 120) {
  if (!worksheet) return 1;
  const r0 = Math.max(1, dataStartRow);
  const c0 = Math.max(1, colFrom);
  const c1 = Math.max(c0, colTo);
  let rows = 0;
  for (let r = r0; r < r0 + maxScan; r += 1) {
    let hit = false;
    for (let c = c0; c <= c1; c += 1) {
      const cell = worksheet.getCell(r, c);
      const val = cell?.value;
      if (val != null && String(val).trim() !== '') hit = true;
      if (excelJSCellHasBorder(cell)) hit = true;
    }
    if (hit) rows += 1;
    else if (rows > 0) break;
  }
  return Math.max(rows, 1);
}

/** Plain display text from an ExcelJS cell value (for alignment checks). */
export function statutoryCellValueToPlainText(value) {
  if (value == null || value === '') return '';
  if (typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') {
    return String(value).trim();
  }
  if (value instanceof Date) {
    try {
      return value.toLocaleDateString('en-IN');
    } catch (_) {
      return value.toISOString().slice(0, 10);
    }
  }
  if (typeof value === 'object') {
    if (Array.isArray(value.richText)) {
      return value.richText
        .map((p) => p?.text || '')
        .join('')
        .trim();
    }
    if (value.text != null) return String(value.text).trim();
    if (value.result != null) return statutoryCellValueToPlainText(value.result);
    if (value.w != null && String(value.w).trim() !== '') return String(value.w).trim();
    if (value.hyperlink != null && value.text != null) return String(value.text).trim();
    if (value.v != null) return statutoryCellValueToPlainText(value.v);
  }
  return String(value).trim();
}

/**
 * True for pure numeric display values (amounts, counts, hours).
 * Text / codes / dates with separators stay non-numeric.
 * e.g. 77765, 7,776.50, ₹1200, 12% → true; hhhbhj, VE123, P, 23-07-2026 → false
 */
export function isStatutoryNumericCellValue(value) {
  if (typeof value === 'number') return Number.isFinite(value);
  if (value instanceof Date) return false;
  const raw = statutoryCellValueToPlainText(value);
  if (!raw) return false;
  let s = raw.replace(/,/g, '').trim();
  s = s.replace(/^[₹$€£]\s?/, '').trim();
  if (/^\((.+)\)$/.test(s)) {
    s = `-${RegExp.$1}`.replace(/,/g, '').trim();
  }
  if (/%$/.test(s)) s = s.replace(/%$/, '').trim();
  if (!s) return false;
  return /^-?\d+(\.\d+)?$/.test(s);
}

/**
 * Form Number / See Rule / Form Name title band (Form XXIII model).
 * These stay center-aligned on draft download — never forced left.
 * Exception: Rajasthan Form 14 table data is left-aligned (see applyStatutoryDownloadContentAlignment).
 */
export function isStatutoryFormTitleBandText(text) {
  const t = String(text || '')
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  if (!t || t.length > 140) return false;
  const lower = t.toLowerCase();

  // Form number: FORM XXIII, FORM - W, Form XXVI, Form U, Form 25, FORM-W
  if (
    /^\*?form[\s._-]*([a-z]{1,5}|[ivxlcdm]+|\d+[a-z]?)\b/i.test(t) &&
    t.length <= 72 &&
    !/name and address|nature and location/i.test(lower)
  ) {
    return true;
  }

  // Rule / sub-rule citation line (See Rule… or A.P. Vide Rule…)
  if (
    (/see\s+(sub-)?rule/i.test(lower) ||
      /vide\s+rule/i.test(lower) ||
      /^\*?\[?\s*see\s+/i.test(lower) ||
      /\(see\s+/i.test(lower) ||
      /\(vide\s+/i.test(lower)) &&
    t.length <= 120
  ) {
    return true;
  }

  // Form name / register title — short, no "Label : value" field
  if (/name and address|nature and location|principal employer|wage period from/i.test(lower)) {
    return false;
  }
  // Allow "Label: value" only when it is a see-rule line; otherwise field lines stay as-is
  if (/:/.test(t) && !/see\s+(sub-)?rule/i.test(lower)) return false;

  if (/^register\s+of\b/i.test(t) && t.length <= 100) return true;
  if (/^muster\s+roll\b/i.test(t) && t.length <= 100) return true;
  if (/^list\s+of\b/i.test(t) && t.length <= 100) return true;
  if (/^notice\s+of\b/i.test(t) && t.length <= 100) return true;
  if (/^identity\s+card\b/i.test(t) && t.length <= 100) return true;
  if (/^wage\s+slip\b/i.test(t) && t.length <= 100) return true;
  if (/^combined\s+muster\b/i.test(t) && t.length <= 100) return true;
  if (/^register\s+of\s+overtime\b/i.test(t)) return true;
  // CLRA Form XV / XI service certificate title (Gujarat / AP / TN / MP).
  if (/^service\s+certificate\b/i.test(t) && t.length <= 80) return true;
  return false;
}

/**
 * Form XV Service Certificate (esp. Gujarat two-column) — employment table values stay centered.
 */
function worksheetLooksLikeFormXVServiceCertificate(worksheet) {
  if (!worksheet || typeof worksheet.getCell !== 'function') return false;
  let hasFormXv = false;
  let hasServiceCert = false;
  const maxR = Math.min(8, Number(worksheet.rowCount) || 8);
  for (let r = 1; r <= maxR; r += 1) {
    for (let c = 1; c <= 12; c += 1) {
      let value;
      try {
        value = worksheet.getCell(r, c)?.value;
      } catch (_) {
        continue;
      }
      const text = statutoryCellValueToPlainText(value);
      if (!text) continue;
      if (/^form\s*xv\b/i.test(text)) hasFormXv = true;
      if (/^service\s+certificate\b/i.test(text)) hasServiceCert = true;
      if (hasFormXv && hasServiceCert) return true;
    }
  }
  return hasFormXv && hasServiceCert;
}

/** AP CLRA Form XXI — Register of Fines (boxed table values stay left-aligned). */
export function worksheetLooksLikeFormXXIAPRegisterOfFines(worksheet) {
  if (!worksheet || typeof worksheet.getCell !== 'function') return false;
  const parts = [];
  const maxR = Math.min(Number(worksheet.rowCount) || 24, 24);
  for (let r = 1; r <= maxR; r += 1) {
    for (let c = 1; c <= 16; c += 1) {
      let value;
      try {
        value = worksheet.getCell(r, c)?.value;
      } catch (_) {
        continue;
      }
      const text = statutoryCellValueToPlainText(value);
      if (text) parts.push(text);
    }
  }
  const blob = `${String(worksheet.name || '')} ${parts.join(' ')}`.toLowerCase();
  if (/form\s*[-._ ]*xxi\b/.test(blob) && /register\s+of\s+fines/.test(blob)) return true;
  if (
    /act\s*\/\s*omission\s+for\s+which\s+fine\s+imposed/.test(blob) &&
    /date\s+on\s+which\s+fine\s+reali[sz]/.test(blob)
  ) {
    return true;
  }
  return false;
}

/** Tamil Nadu CLRA Form XVIII — Register of Wages-cum-Muster Roll. */
export function worksheetLooksLikeFormXVIIITamilNadu(worksheet) {
  if (!worksheet || typeof worksheet.getCell !== 'function') return false;
  const parts = [String(worksheet.name || '')];
  const maxR = Math.min(Number(worksheet.rowCount) || 20, 20);
  for (let r = 1; r <= maxR; r += 1) {
    for (let c = 1; c <= 18; c += 1) {
      try {
        const text = statutoryCellValueToPlainText(worksheet.getCell(r, c)?.value);
        if (text) parts.push(text);
      } catch {
        /* ignore */
      }
    }
  }
  const blob = parts.join(' ').toLowerCase();
  if (/form[\s._-]*xxviii(?![a-z])/.test(blob)) return false;
  const isXviii = /form[\s._-]*xviii(?![a-z])/.test(blob) || /form[\s._-]*18(?!\d)/.test(blob);
  const isWagesCumMuster =
    /wages[\s._-]*cum[\s._-]*muster/.test(blob) ||
    /register\s+of\s+wages[\s-]*cum[\s-]*muster/.test(blob);
  const hasAttendancePair = /daily\s+attendance/.test(blob) && /total\s+attendance/.test(blob);
  return (isXviii && (isWagesCumMuster || hasAttendancePair)) || (isWagesCumMuster && hasAttendancePair);
}

function findFormXVIIITamilNaduDataStartRow(worksheet) {
  const maxR = Math.min(Number(worksheet?.rowCount) || 22, 22);
  for (let r = 8; r <= maxR; r += 1) {
    let numbered = 0;
    for (let c = 1; c <= 18; c += 1) {
      const t = statutoryCellValueToPlainText(worksheet.getCell(r, c)?.value).trim();
      if (/^\d{1,2}$/.test(t)) {
        const n = Number(t);
        if (n >= 1 && n <= 16) numbered += 1;
      }
    }
    if (numbered >= 8) return r + 1;
  }
  return 13;
}

/** Rajasthan Form 14 — Record of Hours of Work (table values left-aligned like template model). */
export function worksheetLooksLikeForm14Rajasthan(worksheet) {
  if (!worksheet || typeof worksheet.getCell !== 'function') return false;
  let hasForm14 = false;
  let hasHoursOfWork = false;
  let hasNameOfPersons = false;
  let hasYoungPerson = false;
  const maxR = Math.min(40, Number(worksheet.rowCount) || 40);
  for (let r = 1; r <= maxR; r += 1) {
    for (let c = 1; c <= 12; c += 1) {
      let value;
      try {
        value = worksheet.getCell(r, c)?.value;
      } catch (_) {
        continue;
      }
      const text = statutoryCellValueToPlainText(value);
      if (!text) continue;
      const lower = text.toLowerCase();
      if (/^form\s*14\b/i.test(text) && !/xiv/i.test(text)) hasForm14 = true;
      if (/record\s+of\s+(?:the\s+)?hours\s+of\s+work/i.test(lower)) hasHoursOfWork = true;
      if (/name\s+of\s+persons?\s+employ/i.test(lower)) hasNameOfPersons = true;
      if (/whether\s+young\s+person/i.test(lower)) hasYoungPerson = true;
      if ((hasForm14 && hasHoursOfWork) || (hasNameOfPersons && hasYoungPerson && hasForm14)) {
        return true;
      }
    }
  }
  return (hasNameOfPersons && hasYoungPerson) || (hasForm14 && hasHoursOfWork);
}

/**
 * Download alignment:
 * - Form Number / Rule / Form Name → center
 * - Numeric values → right (except Form XV service-certificate table → center)
 * - Rajasthan Form 14 table / month-year → left (matches official template model)
 * - Form XXI AP Register of Fines boxed cells → left
 * - Form XVIII TN table body (below column-number row) → left
 * - No left-align override for other text (keeps template centering)
 */
export function applyStatutoryDownloadContentAlignment(worksheet) {
  if (!worksheet) return;
  const formXvServiceCertificate = worksheetLooksLikeFormXVServiceCertificate(worksheet);
  const form14Rajasthan = worksheetLooksLikeForm14Rajasthan(worksheet);
  const formXXIAPFines = worksheetLooksLikeFormXXIAPRegisterOfFines(worksheet);
  const formXVIIITamilNadu = worksheetLooksLikeFormXVIIITamilNadu(worksheet);
  const formXVIIIDataStart = formXVIIITamilNadu
    ? findFormXVIIITamilNaduDataStartRow(worksheet)
    : 0;
  const sheetNameLower = String(worksheet?.name || '').toLowerCase();
  let formIFinesProbe = ` ${sheetNameLower} `;
  for (let r = 1; r <= 4; r += 1) {
    for (let c = 1; c <= 10; c += 1) {
      const t = statutoryCellValueToPlainText(worksheet.getCell(r, c)?.value);
      if (t) formIFinesProbe += ` ${t}`;
    }
  }
  const formITamilNaduFines =
    /pw\s*form\s*i/.test(sheetNameLower) ||
    (/register\s+of\s+fines/i.test(formIFinesProbe) &&
      !/register\s+of\s+workmen|conferment/i.test(formIFinesProbe));
  worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    row.eachCell({ includeEmpty: false }, (cell) => {
      const text = statutoryCellValueToPlainText(cell?.value);
      if (!text) return;
      if (/^this\s+is\s+a\s+system\s+generated\s+document\.?$/i.test(text)) return;

      const next = {
        ...(cell.alignment || {}),
        vertical: cell.alignment?.vertical || 'middle'
      };
      // Preserve template wrapText — do not force wrap on every cell (shreds Form XXIII headers).

      // Always center Form Number / Rule / Form Name (do not leave/force left).
      // Keep wrapText false on title band so a narrow column cannot stack the title vertically.
      // Long Gujarat "Vide rule 77 … Gujarat Rules" line may need wrap inside its merge.
      if (rowNumber <= 20 && isStatutoryFormTitleBandText(text)) {
        // Form 14 RJ model: center the title band across the sheet.
        next.horizontal = 'center';
        next.wrapText = /vide\s+rule/i.test(text) && text.length > 60 ? true : false;
        next.textRotation = 0;
        cell.alignment = next;
        return;
      }

      // Form I TN Register of Fines — table body (names + NIL) stays left.
      if (formITamilNaduFines && rowNumber >= 5) {
        next.horizontal = 'left';
        next.wrapText = true;
        cell.alignment = next;
        return;
      }

      // Form XXI AP — every boxed cell (headers, S.No, NIL, names) stays left.
      if (formXXIAPFines) {
        next.horizontal = 'left';
        cell.alignment = next;
        return;
      }

      // Form XVIII TN — table body text/numbers sit left in each box.
      if (formXVIIITamilNadu && formXVIIIDataStart > 0 && rowNumber >= formXVIIIDataStart) {
        next.horizontal = 'left';
        next.vertical = cell.alignment?.vertical || 'middle';
        cell.alignment = next;
        return;
      }

      // Form 14 RJ — table values (names, Young, hours, Nil) stay left like the template model.
      // Do not left-align title lines (already handled above).
      if (form14Rajasthan) {
        const lower = text.toLowerCase();
        if (
          /rajasthan\s+shops/.test(lower) ||
          /^form\s*14\b/.test(lower) ||
          /rule\s*22/.test(lower) ||
          /hours\s+of\s+work/.test(lower) ||
          /notice\s+in\s+form\s*13/.test(lower)
        ) {
          next.horizontal = 'center';
          cell.alignment = next;
          return;
        }
        next.horizontal = 'left';
        // Long header labels / footnotes may wrap; short data cells stay single-line.
        if (text.length < 40 && !/name of persons|young person|total hours|extent of overtime|days on which overtime/i.test(text)) {
          next.wrapText = false;
        }
        cell.alignment = next;
        return;
      }

      // Numbers only — never force normal text to left.
      if (isStatutoryNumericCellValue(cell?.value)) {
        // Form XV employment table (Sl No / rate / period) matches the certificate template.
        next.horizontal = formXvServiceCertificate && rowNumber >= 30 ? 'center' : 'right';
        cell.alignment = next;
      }
    });
  });
}
