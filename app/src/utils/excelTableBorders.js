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

const DEFAULT_ALL_BORDERS = EXCELJS_THIN_BLACK_BOX_BORDER;

/** Apply a full thin box border to every cell in [rowFrom..rowTo] × [colFrom..colTo] (1-based). */
export function applyExcelJSFullBoxBordersToRange(
  worksheet,
  { rowFrom, rowTo, colFrom, colTo } = {}
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
      const box = {
        top: { style: 'thin', color: { argb: 'FF000000' } },
        left: { style: 'thin', color: { argb: 'FF000000' } },
        bottom: { style: 'thin', color: { argb: 'FF000000' } },
        right: { style: 'thin', color: { argb: 'FF000000' } }
      };
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
