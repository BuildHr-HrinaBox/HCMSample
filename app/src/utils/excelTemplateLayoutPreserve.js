/**
 * Original statutory Excel templates are the single source of truth.
 *
 * Load the original workbook, write only dynamic values, then restore
 * sheet structure / formatting that fill code must not change.
 *
 * Per-form layouts stay as they are in each template — this module does not
 * apply one global column count, alignment, or autofit rule.
 */

import ExcelJS from 'exceljs';
import * as XLSX from 'xlsx';

const WIDTH_TOLERANCE = 0.2;
const HEIGHT_TOLERANCE = 0.75;

function cloneJson(value) {
  if (value == null) return value;
  try {
    return JSON.parse(JSON.stringify(value));
  } catch {
    if (Array.isArray(value)) return value.map((v) => cloneJson(v));
    if (typeof value === 'object') return { ...value };
    return value;
  }
}

function cellHasDisplayValue(value) {
  if (value == null || value === '') return false;
  if (typeof value === 'object' && !Array.isArray(value)) {
    if (Array.isArray(value.richText)) {
      return value.richText.some((p) => String(p?.text || '').trim());
    }
    if (value.text != null) return String(value.text).trim() !== '';
    if (value.result != null) return cellHasDisplayValue(value.result);
    if (value.formula || value.sharedFormula) return false;
  }
  return String(value).trim() !== '';
}

function parseExcelRange(range) {
  const m = String(range || '').match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/i);
  if (!m) return null;
  const colNum = (letters) =>
    String(letters)
      .toUpperCase()
      .split('')
      .reduce((n, ch) => n * 26 + (ch.charCodeAt(0) - 64), 0);
  return {
    c1: colNum(m[1]),
    r1: parseInt(m[2], 10),
    c2: colNum(m[3]),
    r2: parseInt(m[4], 10),
    raw: String(range)
  };
}

function snapshotExcelJsCellStyle(cell) {
  if (!cell) return null;
  const style = {
    alignment: cell.alignment ? cloneJson(cell.alignment) : undefined,
    font: cell.font ? cloneJson(cell.font) : undefined,
    border: cell.border ? cloneJson(cell.border) : undefined,
    fill: cell.fill ? cloneJson(cell.fill) : undefined,
    numFmt: cell.numFmt || undefined,
    protection: cell.protection ? cloneJson(cell.protection) : undefined
  };
  if (!style.alignment && !style.font && !style.border && !style.fill && !style.numFmt && !style.protection) {
    return null;
  }
  return style;
}

function applyExcelJsCellStyle(cell, style) {
  if (!cell || !style) return;
  if (style.alignment) cell.alignment = cloneJson(style.alignment);
  if (style.font) cell.font = cloneJson(style.font);
  if (style.border) cell.border = cloneJson(style.border);
  if (style.fill) cell.fill = cloneJson(style.fill);
  if (style.numFmt) cell.numFmt = style.numFmt;
  if (style.protection) cell.protection = cloneJson(style.protection);
}

function worksheetScanBounds(worksheet) {
  const rowCount = Math.max(
    1,
    Number(worksheet.rowCount) || 0,
    Number(worksheet.actualRowCount) || 0,
    40
  );
  const colCount = Math.max(1, Number(worksheet.columnCount) || 0, 16);
  return { rowCount: Math.min(rowCount, 400), colCount: Math.min(colCount, 80) };
}

export function snapshotExcelJsWorksheetLayout(worksheet) {
  if (!worksheet) return null;
  const { rowCount, colCount } = worksheetScanBounds(worksheet);
  const columns = [];
  for (let c = 1; c <= colCount; c += 1) {
    const col = worksheet.getColumn(c);
    columns.push({
      width: col?.width,
      hidden: !!col?.hidden,
      outlineLevel: col?.outlineLevel
    });
  }
  const rows = [];
  for (let r = 1; r <= rowCount; r += 1) {
    const row = worksheet.getRow(r);
    rows.push({
      height: row?.height,
      hidden: !!row?.hidden
    });
  }
  const cellStyles = {};
  for (let r = 1; r <= rowCount; r += 1) {
    for (let c = 1; c <= colCount; c += 1) {
      const style = snapshotExcelJsCellStyle(worksheet.getCell(r, c));
      if (style) cellStyles[`${r}:${c}`] = style;
    }
  }
  return {
    name: worksheet.name,
    merges: Array.isArray(worksheet.model?.merges) ? [...worksheet.model.merges] : [],
    columns,
    rows,
    pageSetup: cloneJson(worksheet.pageSetup || {}),
    headerFooter: cloneJson(worksheet.headerFooter || {}),
    views: cloneJson(worksheet.views || []),
    properties: cloneJson(worksheet.properties || {}),
    cellStyles
  };
}

export function snapshotExcelJsWorkbookLayout(workbook) {
  if (!workbook) return null;
  return {
    sheetNames: (workbook.worksheets || []).map((ws) => ws.name),
    sheets: (workbook.worksheets || []).map((ws) => snapshotExcelJsWorksheetLayout(ws))
  };
}

function canRestoreExcelJsMerge(worksheet, rangeStr) {
  const parsed = parseExcelRange(rangeStr);
  if (!parsed) return false;
  if (parsed.r1 === parsed.r2) return true;
  for (let r = parsed.r1; r <= parsed.r2; r += 1) {
    for (let c = parsed.c1; c <= parsed.c2; c += 1) {
      if (r === parsed.r1 && c === parsed.c1) continue;
      try {
        if (cellHasDisplayValue(worksheet.getCell(r, c)?.value)) return false;
      } catch {
        return false;
      }
    }
  }
  return true;
}

export function restoreExcelJsWorksheetLayout(worksheet, snapshot) {
  if (!worksheet || !snapshot) return;
  const columns = Array.isArray(snapshot.columns) ? snapshot.columns : [];
  for (let i = 0; i < columns.length; i += 1) {
    const spec = columns[i] || {};
    const col = worksheet.getColumn(i + 1);
    if (spec.width != null && spec.width !== '') col.width = spec.width;
    col.hidden = !!spec.hidden;
    if (spec.outlineLevel != null) col.outlineLevel = spec.outlineLevel;
  }
  const rows = Array.isArray(snapshot.rows) ? snapshot.rows : [];
  for (let i = 0; i < rows.length; i += 1) {
    const spec = rows[i] || {};
    const row = worksheet.getRow(i + 1);
    if (spec.height != null && spec.height !== '') row.height = spec.height;
    row.hidden = !!spec.hidden;
  }
  if (snapshot.pageSetup && typeof snapshot.pageSetup === 'object') {
    worksheet.pageSetup = { ...(worksheet.pageSetup || {}), ...cloneJson(snapshot.pageSetup) };
  }
  if (snapshot.headerFooter && typeof snapshot.headerFooter === 'object') {
    worksheet.headerFooter = {
      ...(worksheet.headerFooter || {}),
      ...cloneJson(snapshot.headerFooter)
    };
  }
  if (Array.isArray(snapshot.views)) {
    worksheet.views = cloneJson(snapshot.views);
  }
  if (snapshot.properties && typeof snapshot.properties === 'object') {
    Object.assign(worksheet.properties || {}, cloneJson(snapshot.properties));
  }

  const wantedMerges = Array.isArray(snapshot.merges) ? snapshot.merges : [];
  wantedMerges.forEach((range) => {
    const current = Array.isArray(worksheet.model?.merges) ? worksheet.model.merges : [];
    if (current.includes(range)) return;
    if (!canRestoreExcelJsMerge(worksheet, range)) return;
    try {
      worksheet.mergeCells(range);
    } catch {
      /* overlapping merge after extra data rows */
    }
  });

  const styles = snapshot.cellStyles || {};
  Object.keys(styles).forEach((key) => {
    const parts = key.split(':');
    const r = Number(parts[0]);
    const c = Number(parts[1]);
    if (!Number.isFinite(r) || !Number.isFinite(c)) return;
    try {
      applyExcelJsCellStyle(worksheet.getCell(r, c), styles[key]);
    } catch {
      /* ignore */
    }
  });
}

export function restoreExcelJsWorkbookLayout(workbook, snapshot) {
  if (!workbook || !snapshot) return;
  const sheets = Array.isArray(snapshot.sheets) ? snapshot.sheets : [];
  sheets.forEach((sheetSnap) => {
    if (!sheetSnap) return;
    const ws =
      (sheetSnap.name && workbook.getWorksheet(sheetSnap.name)) ||
      workbook.worksheets.find((s) => s.name === sheetSnap.name) ||
      null;
    if (ws) restoreExcelJsWorksheetLayout(ws, sheetSnap);
  });
}

function nearEqual(a, b, tol) {
  const na = Number(a);
  const nb = Number(b);
  if (!Number.isFinite(na) && !Number.isFinite(nb)) return true;
  if (!Number.isFinite(na) || !Number.isFinite(nb)) return na === nb;
  return Math.abs(na - nb) <= tol;
}

export function validateExcelJsWorksheetLayout(worksheet, snapshot) {
  const mismatches = [];
  if (!worksheet || !snapshot) {
    return { ok: false, mismatches: ['missing worksheet or snapshot'] };
  }
  const currentMerges = [...(Array.isArray(worksheet.model?.merges) ? worksheet.model.merges : [])].sort();
  const originalMerges = [...(snapshot.merges || [])].sort();
  originalMerges.forEach((range) => {
    if (currentMerges.includes(range)) return;
    if (!canRestoreExcelJsMerge(worksheet, range)) return;
    mismatches.push(`merge missing: ${range}`);
  });
  const columns = snapshot.columns || [];
  for (let i = 0; i < columns.length; i += 1) {
    const want = columns[i]?.width;
    if (want == null || want === '') continue;
    const got = worksheet.getColumn(i + 1)?.width;
    if (!nearEqual(want, got, WIDTH_TOLERANCE)) {
      mismatches.push(`col ${i + 1} width ${got} != ${want}`);
    }
    if (!!worksheet.getColumn(i + 1)?.hidden !== !!columns[i]?.hidden) {
      mismatches.push(`col ${i + 1} hidden flag`);
    }
  }
  const rows = snapshot.rows || [];
  for (let i = 0; i < rows.length; i += 1) {
    const want = rows[i]?.height;
    if (want == null || want === '') continue;
    const got = worksheet.getRow(i + 1)?.height;
    if (want != null && want !== '' && !nearEqual(want, got, HEIGHT_TOLERANCE)) {
      mismatches.push(`row ${i + 1} height ${got} != ${want}`);
    }
  }
  const ps = worksheet.pageSetup || {};
  const wantPs = snapshot.pageSetup || {};
  ['orientation', 'fitToPage', 'fitToWidth', 'fitToHeight', 'scale', 'paperSize', 'printArea'].forEach(
    (key) => {
      if (wantPs[key] != null && wantPs[key] !== '' && ps[key] !== wantPs[key]) {
        mismatches.push(`pageSetup.${key}`);
      }
    }
  );
  const hf = worksheet.headerFooter || {};
  const wantHf = snapshot.headerFooter || {};
  ['oddHeader', 'oddFooter', 'evenHeader', 'evenFooter', 'firstHeader', 'firstFooter'].forEach((key) => {
    if (wantHf[key] != null && wantHf[key] !== '' && hf[key] !== wantHf[key]) {
      mismatches.push(`headerFooter.${key}`);
    }
  });
  return { ok: mismatches.length === 0, mismatches };
}

export function validateExcelJsWorkbookLayout(workbook, snapshot) {
  const mismatches = [];
  if (!workbook || !snapshot) return { ok: false, mismatches: ['missing workbook or snapshot'] };
  const sheets = Array.isArray(snapshot.sheets) ? snapshot.sheets : [];
  sheets.forEach((sheetSnap) => {
    const ws = sheetSnap?.name ? workbook.getWorksheet(sheetSnap.name) : null;
    if (!ws) {
      mismatches.push(`sheet missing: ${sheetSnap?.name}`);
      return;
    }
    const report = validateExcelJsWorksheetLayout(ws, sheetSnap);
    report.mismatches.forEach((m) => mismatches.push(`${sheetSnap.name}: ${m}`));
  });
  return { ok: mismatches.length === 0, mismatches };
}

export function bindOriginalTemplateLayoutPreserve(workbook) {
  if (!workbook || workbook.__statutoryTemplateLayoutBound) return workbook;
  const snapshot = snapshotExcelJsWorkbookLayout(workbook);
  workbook.__statutoryTemplateLayout = snapshot;
  workbook.__statutoryTemplateLayoutBound = true;
  const xlsx = workbook.xlsx;
  if (!xlsx || xlsx.__statutoryPreserveWriteBound) return workbook;
  xlsx.__statutoryPreserveWriteBound = true;
  const origWrite = xlsx.writeBuffer.bind(xlsx);
  xlsx.writeBuffer = async (...args) => {
    restoreExcelJsWorkbookLayout(workbook, snapshot);
    const report = validateExcelJsWorkbookLayout(workbook, snapshot);
    if (!report.ok) {
      restoreExcelJsWorkbookLayout(workbook, snapshot);
      if (typeof console !== 'undefined' && console.warn) {
        console.warn(
          'Statutory Excel: layout restored to original template',
          report.mismatches.slice(0, 12)
        );
      }
    }
    if (typeof workbook.__applyStatutoryFilledRowRanges === 'function') {
      workbook.__applyStatutoryFilledRowRanges(workbook);
    }
    return origWrite(...args);
  };
  return workbook;
}

export async function loadOriginalStatutoryExcelTemplate(arrayBuffer) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(arrayBuffer);
  return bindOriginalTemplateLayoutPreserve(workbook);
}

function sheetJsCellStyle(cell) {
  if (!cell || typeof cell !== 'object') return undefined;
  if (cell.s) return cloneJson(cell.s);
  const style = {};
  if (cell.z) style.z = cell.z;
  if (cell.l) style.l = cloneJson(cell.l);
  return Object.keys(style).length ? style : undefined;
}

export function snapshotSheetJsWorksheetLayout(ws) {
  if (!ws) return null;
  const cellStyles = {};
  Object.keys(ws).forEach((key) => {
    if (!key || key[0] === '!') return;
    const style = sheetJsCellStyle(ws[key]);
    if (style) cellStyles[key] = style;
  });
  return {
    cols: cloneJson(ws['!cols'] || []),
    rows: cloneJson(ws['!rows'] || []),
    merges: cloneJson(ws['!merges'] || []),
    margins: cloneJson(ws['!margins'] || null),
    pageSetup: cloneJson(ws['!pageSetup'] || null),
    printHeader: cloneJson(ws['!printHeader'] || null),
    header: cloneJson(ws['!header'] || null),
    footer: cloneJson(ws['!footer'] || null),
    freeze: cloneJson(ws['!freeze'] || null),
    views: cloneJson(ws['!views'] || null),
    ref: ws['!ref'] || null,
    cellStyles
  };
}

export function snapshotSheetJsWorkbookLayout(wb) {
  if (!wb || !Array.isArray(wb.SheetNames)) return null;
  const sheets = {};
  wb.SheetNames.forEach((name) => {
    sheets[name] = snapshotSheetJsWorksheetLayout(wb.Sheets?.[name]);
  });
  return { sheetNames: [...wb.SheetNames], sheets };
}

function sheetJsMergeKey(m) {
  if (!m?.s || !m?.e) return '';
  return `${m.s.r}:${m.s.c}:${m.e.r}:${m.e.c}`;
}

function canRestoreSheetJsMerge(ws, merge) {
  if (!merge?.s || !merge?.e) return false;
  if (merge.s.r === merge.e.r) return true;
  for (let r = merge.s.r; r <= merge.e.r; r += 1) {
    for (let c = merge.s.c; c <= merge.e.c; c += 1) {
      if (r === merge.s.r && c === merge.s.c) continue;
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = ws[addr];
      const v = cell && typeof cell === 'object' ? cell.v : cell;
      if (v != null && String(v).trim() !== '') return false;
    }
  }
  return true;
}

export function restoreSheetJsWorksheetLayout(ws, snapshot) {
  if (!ws || !snapshot) return;
  if (snapshot.cols) ws['!cols'] = cloneJson(snapshot.cols);
  if (snapshot.rows) ws['!rows'] = cloneJson(snapshot.rows);
  if (snapshot.margins) ws['!margins'] = cloneJson(snapshot.margins);
  if (snapshot.pageSetup) ws['!pageSetup'] = cloneJson(snapshot.pageSetup);
  if (snapshot.printHeader) ws['!printHeader'] = cloneJson(snapshot.printHeader);
  if (snapshot.header) ws['!header'] = cloneJson(snapshot.header);
  if (snapshot.footer) ws['!footer'] = cloneJson(snapshot.footer);
  if (snapshot.freeze) ws['!freeze'] = cloneJson(snapshot.freeze);
  if (snapshot.views) ws['!views'] = cloneJson(snapshot.views);

  const wanted = Array.isArray(snapshot.merges) ? snapshot.merges : [];
  const current = Array.isArray(ws['!merges']) ? ws['!merges'] : [];
  const have = new Set(current.map(sheetJsMergeKey));
  wanted.forEach((merge) => {
    const key = sheetJsMergeKey(merge);
    if (!key || have.has(key)) return;
    if (!canRestoreSheetJsMerge(ws, merge)) return;
    current.push(cloneJson(merge));
    have.add(key);
  });
  ws['!merges'] = current;

  const styles = snapshot.cellStyles || {};
  Object.keys(styles).forEach((addr) => {
    const cell = ws[addr];
    if (cell == null) return;
    if (typeof cell !== 'object') {
      ws[addr] = { v: cell, t: typeof cell === 'number' ? 'n' : 's', s: cloneJson(styles[addr]) };
      return;
    }
    cell.s = { ...(cell.s || {}), ...cloneJson(styles[addr]) };
  });
}

export function restoreSheetJsWorkbookLayout(wb, snapshot) {
  if (!wb || !snapshot?.sheets) return;
  const names = Array.isArray(wb.SheetNames) ? wb.SheetNames : [];
  names.forEach((name) => {
    const snap = snapshot.sheets[name];
    if (snap && wb.Sheets?.[name]) restoreSheetJsWorksheetLayout(wb.Sheets[name], snap);
  });
}

export function validateSheetJsWorkbookLayout(wb, snapshot) {
  const mismatches = [];
  if (!wb || !snapshot) return { ok: false, mismatches: ['missing workbook or snapshot'] };
  Object.keys(snapshot.sheets || {}).forEach((name) => {
    if (!wb.Sheets?.[name]) return;
    const snap = snapshot.sheets[name];
    const ws = wb.Sheets[name];
    const wantCols = snap.cols || [];
    const gotCols = ws['!cols'] || [];
    for (let i = 0; i < wantCols.length; i += 1) {
      const want = wantCols[i]?.wch ?? wantCols[i]?.wpx;
      const got = gotCols[i]?.wch ?? gotCols[i]?.wpx;
      if (want != null && got != null && !nearEqual(want, got, WIDTH_TOLERANCE)) {
        mismatches.push(`${name} col ${i} width`);
      }
    }
    const wantRows = snap.rows || [];
    const gotRows = ws['!rows'] || [];
    for (let i = 0; i < wantRows.length; i += 1) {
      const want = wantRows[i]?.hpt ?? wantRows[i]?.h;
      const got = gotRows[i]?.hpt ?? gotRows[i]?.h;
      if (want != null && got != null && !nearEqual(want, got, HEIGHT_TOLERANCE)) {
        mismatches.push(`${name} row ${i} height`);
      }
    }
  });
  return { ok: mismatches.length === 0, mismatches };
}

export function preserveSheetJsOriginalTemplateLayout(wb, snapshot) {
  if (!wb || !snapshot) return { ok: true, mismatches: [] };
  restoreSheetJsWorkbookLayout(wb, snapshot);
  const report = validateSheetJsWorkbookLayout(wb, snapshot);
  if (!report.ok) restoreSheetJsWorkbookLayout(wb, snapshot);
  const fit = wb.__statutorySheetJsFitRange;
  if (fit && wb.Sheets?.[fit.sheetName]) {
    trimSheetJsTableToRecordCount(wb.Sheets[fit.sheetName], fit);
  }
  return report;
}

export function trimSheetJsTableToRecordCount(ws, {
  dataStartRow0,
  dataRowCount,
  colFrom0,
  colTo0
} = {}) {
  if (!ws) return;
  const start = Math.max(0, Number(dataStartRow0) || 0);
  const n = Math.max(0, Number(dataRowCount) || 0);
  const c0 = Math.max(0, Number(colFrom0) || 0);
  const c1 = Math.max(c0, Number(colTo0) || c0);
  if (n < 1) return;
  const lastFilled = start + n - 1;
  let maxR = lastFilled + 40;
  try {
    if (ws['!ref']) maxR = Math.max(maxR, XLSX.utils.decode_range(ws['!ref']).e.r);
  } catch {
    /* ignore */
  }
  const rowsMeta = Array.isArray(ws['!rows']) ? ws['!rows'] : [];
  const boxHeight = rowsMeta[start]?.hpt ?? rowsMeta[start]?.h;
  for (let i = 0; i < n; i += 1) {
    if (boxHeight == null) continue;
    if (!ws['!rows']) ws['!rows'] = [];
    const prev = ws['!rows'][start + i] && typeof ws['!rows'][start + i] === 'object'
      ? ws['!rows'][start + i]
      : {};
    ws['!rows'][start + i] = { ...prev, hpt: boxHeight, h: boxHeight };
  }
  for (let r = lastFilled + 1; r <= maxR; r += 1) {
    let noteRow = false;
    for (let c = c0; c <= c1; c += 1) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = ws[addr];
      const v = cell && typeof cell === 'object' ? cell.v : cell;
      if (/^this\s+is\s+a\s+system\s+generated\s+document\.?$/i.test(String(v || '').trim())) {
        noteRow = true;
        break;
      }
    }
    if (noteRow) continue;
    for (let c = c0; c <= c1; c += 1) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = ws[addr];
      const v = cell && typeof cell === 'object' ? cell.v : cell;
      if (/^this\s+is\s+a\s+system\s+generated\s+document\.?$/i.test(String(v || '').trim())) continue;
      if (ws[addr]) delete ws[addr];
    }
  }
}

/** Every ExcelJS template load snapshots original layout; writeBuffer restores it. */
export function installExcelJsOriginalTemplatePreserve() {
  if (ExcelJS.Workbook?.__statutoryPreserveInstalled) return;
  const OrigWorkbook = ExcelJS.Workbook;
  class StatutoryPreserveWorkbook extends OrigWorkbook {
    constructor(...args) {
      super(...args);
      const xlsx = this.xlsx;
      if (!xlsx || xlsx.__statutoryLoadPreserveBound) return;
      xlsx.__statutoryLoadPreserveBound = true;
      const origLoad = xlsx.load.bind(xlsx);
      xlsx.load = async (data, options) => {
        const result = await origLoad(data, options);
        if (!(options && options.skipLayoutPreserve) && !this.__skipStatutoryLayoutPreserve) {
          bindOriginalTemplateLayoutPreserve(this);
        }
        return result;
      };
    }
  }
  StatutoryPreserveWorkbook.__statutoryPreserveInstalled = true;
  ExcelJS.Workbook = StatutoryPreserveWorkbook;
}

installExcelJsOriginalTemplatePreserve();

