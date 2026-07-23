import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { jsPDF } from 'jspdf';
import * as XLSX from 'xlsx';

const EXCEL_EXT_RE = /\.(xlsx|xls|xlsm|xlsb)$/i;
/** Form XXVI TN needs ~44 leaf cols (9 identity + 31 days + 4 trailing). */
const MAX_PDF_COLS = 64;
const MAX_PDF_DATA_ROWS = 2500;
const MAX_ZIP_EXCEL_FILES = 60;
const MAX_TRAILING_EMPTY_AFTER_CONTENT = 2;
const SYSTEM_GENERATED_DOCUMENT_NOTE = 'This is a System Generated Document';

const isSystemGeneratedDocumentNote = (text) =>
  /^this\s+is\s+a\s+system\s+generated\s+document\.?$/i.test(
    String(text || '')
      .replace(/\s+/g, ' ')
      .trim()
  );

const isSystemGeneratedDocumentNoteRow = (row) => {
  if (!Array.isArray(row)) return false;
  const filled = row.map((c) => String(c || '').trim()).filter(Boolean);
  return filled.length === 1 && isSystemGeneratedDocumentNote(filled[0]);
};

const isZipArrayBuffer = (buf) => {
  if (!buf || buf.byteLength < 4) return false;
  const u8 = buf instanceof Uint8Array ? buf : new Uint8Array(buf);
  return u8[0] === 0x50 && u8[1] === 0x4b;
};

const cellToText = (value) => {
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
    if (value.richText && Array.isArray(value.richText)) {
      return value.richText
        .map((p) => p?.text || '')
        .join('')
        .trim();
    }
    if (value.t === 'd' && value.v instanceof Date) return cellToText(value.v);
    if (value.w != null && String(value.w).trim() !== '') return String(value.w).trim();
    if (value.text != null) return String(value.text).trim();
    if (value.result != null) return cellToText(value.result);
    if (value.v != null) return cellToText(value.v);
    if (value.hyperlink != null && value.text != null) return String(value.text).trim();
  }
  return String(value).trim();
};

const collectExcelBuffersFromDraft = async (arrayBuffer, fileName = '') => {
  const nameHint = String(fileName || '');
  if (!arrayBuffer || arrayBuffer.byteLength < 32) return [];

  const asSingleWorkbook = () => [
    {
      arrayBuffer,
      label: nameHint.replace(EXCEL_EXT_RE, '').replace(/\.zip$/i, '') || 'Draft'
    }
  ];

  if (EXCEL_EXT_RE.test(nameHint) && !/\.zip$/i.test(nameHint)) {
    return asSingleWorkbook();
  }

  if (isZipArrayBuffer(arrayBuffer) || /\.zip$/i.test(nameHint)) {
    try {
      const zip = await JSZip.loadAsync(arrayBuffer);
      const paths = Object.keys(zip.files || {});
      const isOoxmlWorkbook = paths.some(
        (p) =>
          /(^|\/)\[Content_Types\]\.xml$/i.test(p) ||
          /(^|\/)xl\/workbook\.xml$/i.test(p) ||
          /(^|\/)xl\/worksheets\//i.test(p)
      );
      if (isOoxmlWorkbook) return asSingleWorkbook();

      const entries = paths
        .filter((path) => {
          const f = zip.files[path];
          return f && !f.dir && EXCEL_EXT_RE.test(path) && !path.startsWith('__MACOSX');
        })
        .sort((a, b) => a.localeCompare(b))
        .slice(0, MAX_ZIP_EXCEL_FILES);

      if (entries.length > 0) {
        const out = [];
        for (const path of entries) {
          const buf = await zip.files[path].async('arraybuffer');
          if (!buf || buf.byteLength < 32) continue;
          const base = path.split('/').pop() || path;
          out.push({ arrayBuffer: buf, label: base.replace(EXCEL_EXT_RE, '') });
        }
        if (out.length > 0) return out;
      }
    } catch (zipErr) {
      console.warn('Draft ZIP inspect failed; trying as workbook:', zipErr);
    }
  }

  return asSingleWorkbook();
};

/**
 * Keep merged-cell value only in the top-left cell; clear duplicates in the merge range.
 * Prevents FORM XVIII title/address repeating in every column.
 */
const collapseMergedCellDuplicates = (rows, merges) => {
  if (!Array.isArray(merges) || !merges.length || !rows.length) return rows;
  merges.forEach((m) => {
    if (!m?.s || !m?.e) return;
    const r0 = m.s.r;
    const c0 = m.s.c;
    const r1 = m.e.r;
    const c1 = m.e.c;
    let master = '';
    if (rows[r0] && rows[r0][c0] != null) master = cellToText(rows[r0][c0]);
    // If master empty, take first non-empty in range
    if (!master) {
      for (let r = r0; r <= r1 && r < rows.length; r += 1) {
        for (let c = c0; c <= c1; c += 1) {
          const t = cellToText(rows[r]?.[c]);
          if (t) {
            master = t;
            break;
          }
        }
        if (master) break;
      }
    }
    for (let r = r0; r <= r1 && r < rows.length; r += 1) {
      if (!rows[r]) continue;
      while (rows[r].length <= c1) rows[r].push('');
      for (let c = c0; c <= c1; c += 1) {
        rows[r][c] = r === r0 && c === c0 ? master : '';
      }
    }
  });
  return rows;
};

/**
 * If the same long text is copied across many columns (broken merge), keep it once.
 */
const dedupeRepeatedRowText = (row) => {
  if (!Array.isArray(row) || row.length < 3) return row;
  const nonEmpty = row
    .map((c, idx) => ({ text: cellToText(c), idx }))
    .filter((x) => x.text);
  if (nonEmpty.length < 3) return row;

  const counts = new Map();
  nonEmpty.forEach(({ text }) => {
    counts.set(text, (counts.get(text) || 0) + 1);
  });

  let dominant = '';
  let dominantCount = 0;
  counts.forEach((count, text) => {
    if (count > dominantCount) {
      dominant = text;
      dominantCount = count;
    }
  });

  // Same title/address pasted into most columns
  const ratio = dominantCount / nonEmpty.length;
  if (dominant && dominant.length >= 12 && dominantCount >= 3 && ratio >= 0.5) {
    const out = row.map(() => '');
    out[0] = dominant;
    // Preserve other unique short values (rare)
    nonEmpty.forEach(({ text, idx }) => {
      if (text !== dominant) out[idx] = text;
    });
    return out;
  }
  return row;
};

const isFormMetaText = (text) => {
  const t = String(text || '').toLowerCase();
  if (!t) return false;
  return (
    /^form\s*(xviii|xvii|xix|xv|xiv|[a-z0-9-]+)\b/.test(t) ||
    /see rule|register of wages|muster roll|name and address of contractor|nature and location of work|address of the establishment|principal employer|wage period|from\s*to|list of workers|shift schedule|shift scehdule|all the workers in establishment/.test(
      t
    )
  );
};

/** Match Sr.No / S.No / SI.No and common wage-register identity headers. */
const isWageRegisterColHeaderBlob = (blob) => {
  const t = String(blob || '').toLowerCase();
  if (!t) return false;
  return (
    /(?:^|[^a-z])s\.?\s*no(?:[^a-z]|$)|(?:^|[^a-z])sr\.?\s*no(?:[^a-z]|$)|(?:^|[^a-z])si\.?\s*no(?:[^a-z]|$)/.test(
      t
    ) ||
    /name of(?:\s+the)?\s+employee|employee identification|name of the work(?:er|man)|designation|daily (?:attendance|hours)|amount of wages|net amount|rate of wages|basic wage|number of days worked/.test(
      t
    )
  );
};

/** Form W admin band above the leaf table (employer / gender boxes / month-year). */
const isFormWAdminBandBlob = (blob) => {
  const t = String(blob || '').toLowerCase();
  if (!t || isWageRegisterColHeaderBlob(t)) return false;
  return (
    /name and address of the employer|name of the manager|registration certificate|wage period from/.test(
      t
    ) ||
    (/\bmen\b/.test(t) && /\bwomen\b/.test(t)) ||
    (/male young person|female young person/.test(t) && !/s\.?\s*no/.test(t)) ||
    (/^month\s*:/.test(t.trim()) || /\bmonth\s*:/.test(t)) ||
    (/^year\s*:/.test(t.trim()) || /\byear\s*:/.test(t))
  );
};

/** Gender count / Month-Year value row sitting between Form W admin labels and S.No. */
const isFormWAdminValueRow = (filled) => {
  if (!Array.isArray(filled) || filled.length === 0 || filled.length > 8) return false;
  const allShort = filled.every((t) => {
    const s = String(t || '').trim();
    if (!s) return true;
    if (/^\d{1,4}$/.test(s)) return true;
    if (/^(month|year)\s*:?\s*\S*$/i.test(s)) return true;
    // Month name alone (May / April 2024)
    if (/^(january|february|march|april|may|june|july|august|september|october|november|december)(\s+\d{4})?$/i.test(s)) {
      return true;
    }
    if (/^\d{4}$/.test(s)) return true;
    return false;
  });
  return allShort && filled.some((t) => String(t || '').trim());
};

const looksLikeFormWPdfContext = (metaLines, rows, tableStart = 0) => {
  const blob = [...(metaLines || []), ...(rows || []).slice(0, Math.min(rows.length, 14)).flat()]
    .map((t) => String(t || ''))
    .join(' ')
    .toLowerCase();
  const isFormW =
    /(?:^|[^a-z0-9])form[\s._-]*w(?:[^a-z0-9]|$)/.test(blob) ||
    (/register of wages/.test(blob) && /sub-rule\s*\(?\s*1\s*\)?\s*of\s*rule\s*\(?\s*16/.test(blob));
  if (!isFormW) return false;
  // Prefer Tamil Nadu Form W; also accept plain Form W templates without state text.
  return (
    /tamil[\s._-]*nadu|tamilnadu|sub-rule\s*\(?\s*1\s*\)?\s*of\s*rule\s*\(?\s*16/.test(blob) ||
    isWageRegisterColHeaderBlob(
      (rows || [])
        .slice(Math.max(0, tableStart), Math.min(rows.length, (tableStart || 0) + 4))
        .flat()
        .join(' ')
    )
  );
};

/**
 * Build a dense AOA matrix from a worksheet — all data cells, merges collapsed.
 */
const sheetToDenseMatrix = (worksheet, sheetName = 'Sheet') => {
  const aoa = XLSX.utils.sheet_to_json(worksheet, {
    header: 1,
    defval: '',
    raw: false,
    blankrows: true
  });
  if (!Array.isArray(aoa) || aoa.length === 0) {
    return { name: sheetName, rows: [], colCount: 0, metaLines: [], tableStartRow: 0 };
  }

  let rows = aoa.map((row) => {
    const src = Array.isArray(row) ? row : [];
    return src.map((c) => cellToText(c));
  });

  collapseMergedCellDuplicates(rows, worksheet['!merges'] || []);
  rows = rows.map((row) => dedupeRepeatedRowText(row));

  // Prefer Excel !ref width so empty trailing cols (signature etc.) stay in the PDF band.
  let refColCount = 0;
  try {
    const ref = worksheet['!ref'];
    if (ref) {
      const decoded = XLSX.utils.decode_range(ref);
      refColCount = Math.max(0, (decoded?.e?.c ?? -1) + 1);
    }
  } catch (_) {
    refColCount = 0;
  }

  let maxCol = 0;
  let lastContentRow = -1;
  for (let r = 0; r < rows.length && r < MAX_PDF_DATA_ROWS; r += 1) {
    // Footer note lives in col 0 after merge — do not shrink the table to that column.
    if (isSystemGeneratedDocumentNoteRow(rows[r])) {
      lastContentRow = Math.max(lastContentRow, r);
      continue;
    }
    let hasValue = false;
    for (let c = 0; c < rows[r].length; c += 1) {
      if (rows[r][c]) {
        hasValue = true;
        maxCol = Math.max(maxCol, c + 1);
      }
    }
    if (hasValue) lastContentRow = r;
  }

  if (lastContentRow < 0) {
    return { name: sheetName, rows: [], colCount: 0, metaLines: [], tableStartRow: 0 };
  }

  const endRow = Math.min(rows.length - 1, lastContentRow + MAX_TRAILING_EMPTY_AFTER_CONTENT);
  // Wide day-grid forms (e.g. Form XXVI): keep a few empty trailing cols from !ref
  // (signature / termination) without inflating narrow sheets to the template width.
  if (refColCount > maxCol && maxCol >= 20 && refColCount - maxCol <= 8) {
    maxCol = refColCount;
  }
  maxCol = Math.min(Math.max(maxCol, 1), MAX_PDF_COLS);

  const padded = [];
  for (let r = 0; r <= endRow; r += 1) {
    const line = [];
    for (let c = 0; c < maxCol; c += 1) {
      line.push(cellToText(rows[r]?.[c]));
    }
    padded.push(line);
  }

  // Split form title/meta (full-width once) from the data table
  let tableStartRow = 0;
  const metaLines = [];
  for (let r = 0; r < Math.min(padded.length, 20); r += 1) {
    const filled = padded[r].filter((c) => c);
    const blob = filled.join(' ').toLowerCase();
    const isColHeader = isWageRegisterColHeaderBlob(blob) && filled.length >= 3;
    const isNumberRow =
      filled.filter((c) => /^\d{1,2}$/.test(c)).length >= Math.max(6, filled.length * 0.6);

    if (isColHeader || isNumberRow) {
      tableStartRow = r;
      break;
    }

    if (filled.length === 0) continue;

    // Form W: employer / Men-Women / Month-Year sit above the leaf header — keep as meta.
    const alreadyFormWMeta = looksLikeFormWPdfContext(metaLines, [], 0);
    if (
      isFormWAdminBandBlob(blob) ||
      (alreadyFormWMeta && isFormWAdminValueRow(filled))
    ) {
      const unique = [];
      filled.forEach((t) => {
        if (!unique.includes(t)) unique.push(t);
      });
      unique.forEach((t) => metaLines.push(t));
      tableStartRow = r + 1;
      continue;
    }

    // One (or few) meta values — show once as a line
    if (filled.length <= 2 || filled.every((t) => isFormMetaText(t) || t.length > 20)) {
      const unique = [];
      filled.forEach((t) => {
        if (!unique.includes(t)) unique.push(t);
      });
      unique.forEach((t) => metaLines.push(t));
      tableStartRow = r + 1;
      continue;
    }

    // Many columns but duplicated meta text already collapsed to col0
    if (filled.length === 1 && isFormMetaText(filled[0])) {
      metaLines.push(filled[0]);
      tableStartRow = r + 1;
      continue;
    }

    tableStartRow = r;
    break;
  }

  return {
    name: sheetName,
    rows: padded,
    colCount: maxCol,
    metaLines,
    tableStartRow
  };
};

/** Fill any blank cells from ExcelJS (formulas / rich text) without re-duplicating merges. */
const enrichMatrixWithExcelJs = async (arrayBuffer, matrices) => {
  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(arrayBuffer);
    const byName = new Map(matrices.map((m) => [String(m.name || '').toLowerCase(), m]));

    workbook.eachSheet((ws) => {
      const key = String(ws.name || '').toLowerCase();
      const matrix = byName.get(key) || matrices[0];
      if (!matrix || !matrix.rows.length) return;

      // Build merge non-master set
      const nonMaster = new Set();
      const merges = ws._merges || ws.model?.merges || {};
      const list = Array.isArray(merges) ? merges : Object.values(merges || {});
      list.forEach((m) => {
        const top = m?.top ?? m?.model?.top;
        const left = m?.left ?? m?.model?.left;
        const bottom = m?.bottom ?? m?.model?.bottom;
        const right = m?.right ?? m?.model?.right;
        if (top == null || left == null) return;
        for (let r = top; r <= bottom; r += 1) {
          for (let c = left; c <= right; c += 1) {
            if (r === top && c === left) continue;
            nonMaster.add(`${r},${c}`);
          }
        }
      });

      ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
        const r = rowNumber - 1;
        if (r < 0 || r >= MAX_PDF_DATA_ROWS) return;
        while (matrix.rows.length <= r && matrix.rows.length < MAX_PDF_DATA_ROWS) {
          matrix.rows.push(Array.from({ length: matrix.colCount }, () => ''));
        }
        if (r >= matrix.rows.length) return;

        row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
          if (nonMaster.has(`${rowNumber},${colNumber}`)) return;
          const c = colNumber - 1;
          if (c < 0 || c >= MAX_PDF_COLS) return;
          while (matrix.colCount <= c) {
            matrix.colCount += 1;
            matrix.rows.forEach((line) => line.push(''));
          }
          const text = cellToText(cell.value);
          if (!text) return;
          if (!matrix.rows[r][c]) matrix.rows[r][c] = text;
        });
      });
    });
  } catch (err) {
    console.warn('ExcelJS enrich skipped:', err);
  }
  return matrices;
};

const collectWorkbookMatrices = async (arrayBuffer, label = 'Draft') => {
  const wb = XLSX.read(arrayBuffer, { type: 'array', cellDates: true, cellText: true });
  const matrices = [];
  (wb.SheetNames || []).forEach((name) => {
    const matrix = sheetToDenseMatrix(wb.Sheets[name], name || label);
    if (matrix.rows.length > 0 && matrix.colCount > 0) {
      matrices.push(matrix);
    }
  });
  if (!matrices.length) return [];
  return enrichMatrixWithExcelJs(arrayBuffer, matrices);
};

const rowHasContent = (row) =>
  Array.isArray(row) && row.some((c) => String(c || '').trim() !== '');

const isLikelyHeaderBandRow = (row, rowIndex) => {
  if (!rowHasContent(row)) return false;
  const filled = row.filter((c) => String(c || '').trim());
  const blob = filled.join(' ').toLowerCase();
  if (rowIndex < 20 && isWageRegisterColHeaderBlob(blob) && filled.length >= 3) {
    return true;
  }
  if (
    rowIndex < 20 &&
    /designation|wage|attendance|daily hours|amount of wages|net amount|deductions|gross wages|net wages/.test(
      blob
    ) &&
    filled.length >= 3
  ) {
    return true;
  }
  const nums = filled.filter((c) => /^\d{1,2}$/.test(String(c).trim()));
  return nums.length >= 8 && nums.length >= filled.length * 0.7;
};

const isPurePdfNumericText = (raw) => {
  const numericRaw = String(raw || '')
    .replace(/,/g, '')
    .replace(/^[₹$€£]\s?/, '')
    .replace(/^\((.+)\)$/, '-$1')
    .replace(/%$/, '')
    .trim();
  return /^-?\d+(\.\d+)?$/.test(numericRaw);
};

/** Leaf-header-aware weights so Form W amount columns stay wide enough for 6–7 digit figures. */
const formWTamilNaduColumnWeight = (headerText, maxDataLen) => {
  const h = String(headerText || '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
  if (!h) return Math.max(6, Math.min(maxDataLen || 4, 9));
  if (/^(?:s|sr|si)\.?\s*no\b/.test(h)) return 3.2;
  if (/name of(?:\s+the)?\s+employee/.test(h)) return 13;
  if (/employee identification|identification\s*no/.test(h)) return 7.5;
  if (/number of days worked|days worked/.test(h)) return 4.5;
  if (/date of payment/.test(h)) return 7.5;
  if (/remarks/.test(h)) return 7;
  if (/receipt by employee|bank transaction/.test(h)) return 8;
  if (/unpaid accumulation|subsistence allowance/.test(h)) return 7;
  if (
    /basic wage|dearness|house rent|other allowance|overtime|gross wages|net wages|total deductions|provident fund|state insurance|labour welfare|advance|damages|pending recovery|any other deduction/.test(
      h
    )
  ) {
    return Math.max(7.2, Math.min(Math.max(maxDataLen || 0, 6) + 1.5, 10));
  }
  return Math.max(6, Math.min(maxDataLen || 4, 9));
};

const resolveLeafHeaderTexts = (rows, tableStart, headerBandEnd, colCount) => {
  const headers = Array.from({ length: colCount }, () => '');
  for (let c = 0; c < colCount; c += 1) {
    let best = '';
    for (let r = headerBandEnd; r >= tableStart; r -= 1) {
      const t = String(rows[r]?.[c] || '')
        .replace(/\s+/g, ' ')
        .trim();
      if (!t) continue;
      // Prefer the deepest (leaf) non-group label; skip ultra-wide group banners alone.
      if (/^deductions$/i.test(t) || /^advances$/i.test(t) || /^damages\s*\/\s*fine$/i.test(t)) {
        if (!best) best = t;
        continue;
      }
      best = t;
      break;
    }
    headers[c] = best;
  }
  return headers;
};

const stripFieldLabelPrefix = (text, labelRe) =>
  String(text || '')
    .replace(labelRe, '')
    .replace(/^[\s.:\-–—]+/, '')
    .trim();

/** Detect Form XXVI CLRA muster from sheet meta / header band. */
const detectFormXXVIPdfLayout = (metaLines, rows, tableStart) => {
  const blob = [...(metaLines || []), ...(rows || []).slice(0, Math.min(rows.length, 12)).flat()]
    .map((t) => String(t || ''))
    .join(' ')
    .toLowerCase();
  const isFormXXVI =
    /form\s*xxvi\b/.test(blob) ||
    (/register of employment of contractual?\s*labour/.test(blob) &&
      /tamil\s+nadu\s+contract\s+labour|see\s+rule\s*75/.test(blob));
  if (!isFormXXVI) return null;

  let principal = '';
  let contractor = '';
  let worksite = '';
  let month = '';
  let date = '';

  (metaLines || []).forEach((raw) => {
    const line = String(raw || '').replace(/\s+/g, ' ').trim();
    if (!line || isSystemGeneratedDocumentNote(line)) return;
    const lower = line.toLowerCase();
    if (/^form\s*xxvi\b/.test(lower)) return;
    if (/see\s+rule\s*75/.test(lower)) return;
    if (/register of employment of contractual?\s*labour/.test(lower)) return;

    if (/principal\s+employer/.test(lower)) {
      principal = stripFieldLabelPrefix(line, /name\s+and\s+address\s+of\s+(?:the\s+)?principal\s+employer/i);
      return;
    }
    if (/name\s+and\s+address\s+of\s+(?:the\s+)?contractor/.test(lower)) {
      contractor = stripFieldLabelPrefix(line, /name\s+and\s+address\s+of\s+(?:the\s+)?contractor\.?/i);
      return;
    }
    if (/nature\s+and\s+location\s+of\s+work|name\s+and\s+location\s+of\s+(?:the\s+)?work/.test(lower)) {
      worksite = stripFieldLabelPrefix(
        line,
        /(?:nature\s+and\s+location\s+of\s+work|name\s+and\s+location\s+of\s+(?:the\s+)?work\s*site?)\.?/i
      );
      return;
    }
    if (/^month\s*:?\s*/i.test(line)) {
      month = stripFieldLabelPrefix(line, /^month/i);
      return;
    }
    if (/^date\s*:?\s*/i.test(line) || /^year\s*:?\s*/i.test(line)) {
      date = stripFieldLabelPrefix(line, /^(?:date|year)/i);
    }
  });

  // Also scan early sheet rows for label/value pairs that meta split across columns.
  for (let r = 0; r < Math.min(tableStart || 0, rows.length); r += 1) {
    const row = rows[r] || [];
    for (let c = 0; c < row.length; c += 1) {
      const cell = String(row[c] || '').replace(/\s+/g, ' ').trim();
      if (!cell) continue;
      const lower = cell.toLowerCase();
      const next = String(row[c + 1] || '').replace(/\s+/g, ' ').trim();
      if (!principal && /principal\s+employer/.test(lower)) {
        principal = stripFieldLabelPrefix(cell, /name\s+and\s+address\s+of\s+(?:the\s+)?principal\s+employer/i) || next;
      }
      if (!contractor && /name\s+and\s+address\s+of\s+(?:the\s+)?contractor/.test(lower)) {
        contractor = stripFieldLabelPrefix(cell, /name\s+and\s+address\s+of\s+(?:the\s+)?contractor\.?/i) || next;
      }
      if (
        !worksite &&
        /nature\s+and\s+location\s+of\s+work|name\s+and\s+location\s+of\s+(?:the\s+)?work/.test(lower)
      ) {
        worksite =
          stripFieldLabelPrefix(
            cell,
            /(?:nature\s+and\s+location\s+of\s+work|name\s+and\s+location\s+of\s+(?:the\s+)?work\s*site?)\.?/i
          ) || next;
      }
      if (!month && /^month\s*:?\s*$/i.test(cell) && next && !/date|year|name and/i.test(next)) month = next;
      if (!date && /^(?:date|year)\s*:?\s*$/i.test(cell) && next && !/name and|nature/i.test(next)) date = next;
    }
  }

  return {
    title: 'FORM XXVI',
    reference: 'See Rule 75 of the Tamil Nadu Contract Labour (Regulation and Abolition) Rules, 1975',
    subtitle: 'Register of Employment of Contractual Labour',
    principal,
    contractor,
    worksite,
    month,
    date
  };
};

/**
 * Find the "Daily hours of work" day band (typically cols after Rate of Wages through day 31).
 */
const detectDailyHoursBand = (rows, tableStart, headerBandEnd, colCount) => {
  let labelRow = -1;
  let labelCol = -1;
  let labelText = 'Daily hours of work';

  for (let r = tableStart; r <= headerBandEnd && r < rows.length; r += 1) {
    for (let c = 0; c < colCount; c += 1) {
      const t = String(rows[r]?.[c] || '').trim();
      if (/daily\s+hours\s+of\s+work/i.test(t)) {
        labelRow = r;
        labelCol = c;
        labelText = t.replace(/\s+/g, ' ').trim();
        break;
      }
    }
    if (labelCol >= 0) break;
  }

  if (labelCol < 0) {
    // Fallback: Rate of Wages followed by a long run of day numbers 1..31
    let rateCol = -1;
    for (let r = tableStart; r <= headerBandEnd && r < rows.length; r += 1) {
      for (let c = 0; c < colCount; c += 1) {
        if (/rate\s+of\s+wages?/i.test(String(rows[r]?.[c] || ''))) rateCol = c;
      }
    }
    if (rateCol < 0 || rateCol + 20 >= colCount) return null;
    labelCol = rateCol + 1;
  }

  let trailingStart = colCount;
  for (let r = tableStart; r <= headerBandEnd && r < rows.length; r += 1) {
    for (let c = labelCol + 1; c < colCount; c += 1) {
      const t = String(rows[r]?.[c] || '');
      if (/number\s+of\s+days|signature|thumb|termination|contractor/i.test(t)) {
        trailingStart = Math.min(trailingStart, c);
      }
    }
  }

  // Prefer an explicit 1..N day-number run in the header band.
  let dayEnd = -1;
  for (let r = tableStart; r <= Math.min(headerBandEnd + 2, rows.length - 1); r += 1) {
    let run = 0;
    let end = -1;
    for (let c = labelCol; c < Math.min(trailingStart, colCount); c += 1) {
      const n = Number(String(rows[r]?.[c] || '').trim());
      if (Number.isInteger(n) && n >= 1 && n <= 31) {
        run += 1;
        end = c;
      } else if (run > 0 && !String(rows[r]?.[c] || '').trim()) {
        // allow blanks inside a collapsed merge
        continue;
      } else if (run > 0) {
        break;
      }
    }
    if (run >= 20) dayEnd = Math.max(dayEnd, end);
  }

  if (dayEnd < labelCol) {
    // Default Form XXVI: 31 day leaf columns
    dayEnd = Math.min(labelCol + 30, trailingStart - 1, colCount - 1);
  }
  if (dayEnd < labelCol) return null;

  return {
    labelRow,
    start: labelCol,
    end: dayEnd,
    label: labelText || 'Daily hours of work'
  };
};

const extractFieldValue = (label, value, { requireColon = true } = {}) => {
  const v = String(value || '').trim();
  if (!v) return requireColon ? `${label} :` : label;
  return `${label} : ${v}`;
};

/** Parse Form W Men / Women / young-person count box from meta + pre-table rows. */
const extractFormWGenderBox = (metaLines, rows, tableStart) => {
  const out = {
    total: '',
    men: '',
    women: '',
    maleYoung: '',
    femaleYoung: ''
  };
  let foundLabels = false;

  const assignCounts = (vals) => {
    if (!Array.isArray(vals) || vals.length < 4) return false;
    const nums = vals.map((v) => String(v ?? '').trim());
    if (!nums.every((n) => /^\d{1,5}$/.test(n))) return false;
    out.men = nums[0];
    out.women = nums[1];
    out.maleYoung = nums[2];
    out.femaleYoung = nums[3];
    return true;
  };

  const meta = (metaLines || [])
    .map((l) => String(l || '').replace(/\s+/g, ' ').trim())
    .filter(Boolean);

  meta.forEach((line) => {
    const m = line.match(/total\s+number\s+of\s+persons?\s+employed\s*:?\s*(\d{1,5})?/i);
    if (m && m[1] != null && m[1] !== '') out.total = m[1];
  });

  // Labels then counts in flat meta (skip Month/Year tokens between labels and numbers)
  for (let i = 0; i < meta.length; i += 1) {
    const a = meta[i].toLowerCase();
    const b = String(meta[i + 1] || '').toLowerCase();
    const c = String(meta[i + 2] || '').toLowerCase();
    const d = String(meta[i + 3] || '').toLowerCase();
    if (
      a === 'men' &&
      b === 'women' &&
      /male young person/.test(c) &&
      /female young person/.test(d)
    ) {
      foundLabels = true;
      const after = [];
      for (let j = i + 4; j < meta.length && after.length < 4; j += 1) {
        const tok = meta[j];
        if (/^\d{1,5}$/.test(tok)) after.push(tok);
        else if (/^(month|year)\b/i.test(tok)) continue;
        else if (after.length > 0) break;
      }
      assignCounts(after);
      break;
    }
  }

  // Dense sheet rows above the wage table
  const scanEnd = Math.min(rows?.length || 0, Math.max(tableStart || 0, 0) + 2, 20);
  for (let r = 0; r < scanEnd; r += 1) {
    const row = rows[r] || [];
    for (let c = 0; c < row.length; c += 1) {
      const cell = String(row[c] || '')
        .replace(/\s+/g, ' ')
        .trim();
      if (!cell) continue;
      const totalMatch = cell.match(/total\s+number\s+of\s+persons?\s+employed\s*:?\s*(\d{1,5})?/i);
      if (totalMatch && totalMatch[1]) out.total = totalMatch[1];

      if (/^men$/i.test(cell)) {
        const labels = [0, 1, 2, 3].map((k) =>
          String(row[c + k] || '')
            .replace(/\s+/g, ' ')
            .trim()
            .toLowerCase()
        );
        if (
          labels[0] === 'men' &&
          labels[1] === 'women' &&
          /male young person/.test(labels[2]) &&
          /female young person/.test(labels[3])
        ) {
          foundLabels = true;
          const next = rows[r + 1] || [];
          assignCounts([next[c], next[c + 1], next[c + 2], next[c + 3]]);
        }
      }
    }
  }

  if (!foundLabels && out.men === '' && out.women === '' && out.maleYoung === '' && out.femaleYoung === '') {
    return null;
  }

  // Always render the four boxes once labels (or counts) are present.
  if (out.men === '') out.men = '0';
  if (out.women === '') out.women = '0';
  if (out.maleYoung === '') out.maleYoung = '0';
  if (out.femaleYoung === '') out.femaleYoung = '0';

  if (out.total === '') {
    const sum =
      (Number(out.men) || 0) +
      (Number(out.women) || 0) +
      (Number(out.maleYoung) || 0) +
      (Number(out.femaleYoung) || 0);
    out.total = String(sum);
  }
  return out;
};

const normalizeMetaKey = (text) =>
  String(text || '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();

const isStatutoryTitleMetaLine = (line) => {
  const t = String(line || '').replace(/\s+/g, ' ').trim();
  if (!t || isSystemGeneratedDocumentNote(t)) return false;
  const lower = t.toLowerCase();
  if (/^form\s*[a-z0-9xivlc.\-]+\b/i.test(t) && t.length < 40) return true;
  if (/see\s+rule|\[see\s+rule|\(see\s+rule/i.test(lower)) return true;
  if (/^register of\b/i.test(t)) return true;
  if (/^muster\s+roll\b/i.test(t)) return true;
  if (/^list of\b/i.test(t) && t.length < 60) return true;
  // License / return style short titles without a field colon
  if (!/:/.test(t) && t.length <= 72 && /form|register|return|notice|accident|wages|employment/i.test(lower)) {
    return true;
  }
  return false;
};

const isStatutoryFieldMetaLine = (line) => {
  const t = String(line || '').replace(/\s+/g, ' ').trim();
  if (!t || isSystemGeneratedDocumentNote(t)) return false;
  if (isStatutoryTitleMetaLine(t)) return false;
  const lower = t.toLowerCase();
  if (/^month\s*:?\s*/i.test(t) || /^date\s*:?\s*/i.test(t) || /^year\s*:?\s*/i.test(t)) return true;
  if (
    /name\s+and\s+address|nature\s+and\s+location|address\s+of|principal\s+employer|contractor|establishment|worksite|wage\s+period|from\s*to|location\s+of\s+work/i.test(
      lower
    )
  ) {
    return true;
  }
  // Generic "Label : value" administrative lines
  if (/^[^:]{3,80}:\s*\S/.test(t)) return true;
  return false;
};

/**
 * Build bordered header model for any statutory form (titles + field rows).
 * Form XXVI keeps Excel-style Month/Date right column when available.
 */
const buildStatutoryPdfHeaderModel = (metaLines, rows, tableStart, sheetName = '') => {
  const xxvi = detectFormXXVIPdfLayout(metaLines, rows, tableStart);
  if (xxvi) {
    return {
      titles: [xxvi.title, xxvi.reference, xxvi.subtitle].filter(Boolean),
      fields: [
        extractFieldValue('Name and address of Principal Employer', xxvi.principal),
        extractFieldValue('Name and Address of Contractor.', xxvi.contractor),
        extractFieldValue('Nature and location of work.', xxvi.worksite)
      ],
      rightFields: [
        extractFieldValue('Month', xxvi.month, { requireColon: false }),
        extractFieldValue('Date', xxvi.date, { requireColon: false }),
        ''
      ],
      hasSystemNote: (metaLines || []).some((l) => isSystemGeneratedDocumentNote(l)),
      isFormXXVI: true,
      isFormW: false
    };
  }

  const isFormW = looksLikeFormWPdfContext(metaLines, rows, tableStart);
  const genderBox = isFormW ? extractFormWGenderBox(metaLines, rows, tableStart) : null;

  const titles = [];
  const fields = [];
  const seen = new Set();
  let month = '';
  let date = '';
  let year = '';
  let hasSystemNote = false;

  const pushUnique = (list, text, { asTitle = false } = {}) => {
    const raw = String(text || '').replace(/\s+/g, ' ').trim();
    if (!raw) return;
    let key = normalizeMetaKey(raw);
    // Collapse "Form XXII" / "FORM XXII"
    if (asTitle && /^form\s+/i.test(raw)) {
      key = `form:${key.replace(/^form\s+/, '')}`;
    }
    if (seen.has(key)) return;
    seen.add(key);
    list.push(raw);
  };

  (metaLines || []).forEach((raw) => {
    const line = String(raw || '').replace(/\s+/g, ' ').trim();
    if (!line) return;
    if (isSystemGeneratedDocumentNote(line)) {
      hasSystemNote = true;
      return;
    }
    // Gender box labels/counts are painted as a dedicated bordered box — not as field lines
    if (/^(men|women|male young person|female young person)$/i.test(line)) return;
    if (/^\d{1,4}$/.test(line) && isFormW) return;
    // Total line is the top row of the gender box when counts exist
    if (genderBox && /total\s+number\s+of\s+persons?\s+employed/i.test(line)) {
      const m = line.match(/total\s+number\s+of\s+persons?\s+employed\s*:?\s*(\d{1,5})?/i);
      if (m?.[1] && !genderBox.total) genderBox.total = m[1];
      return;
    }

    if (/^month\s*:?\s*/i.test(line) && !/name and|address|nature/i.test(line)) {
      month = stripFieldLabelPrefix(line, /^month/i);
      return;
    }
    if (/^year\s*:?\s*/i.test(line) && !/name and|address|nature|entry|termination/i.test(line)) {
      year = stripFieldLabelPrefix(line, /^year/i);
      return;
    }
    if (/^(?:date)\s*:?\s*/i.test(line) && !/name and|address|nature|entry|termination|payment/i.test(line)) {
      date = stripFieldLabelPrefix(line, /^date/i);
      return;
    }
    if (isStatutoryTitleMetaLine(line)) {
      pushUnique(titles, line, { asTitle: true });
      return;
    }
    if (isStatutoryFieldMetaLine(line) || isFormMetaText(line)) {
      pushUnique(fields, line);
      return;
    }
    // Leftover meta — treat short lines as titles, longer as fields
    if (line.length <= 64 && !/:/.test(line)) pushUnique(titles, line, { asTitle: true });
    else pushUnique(fields, line);
  });

  if (isFormW) {
    const hasFormWTitle = titles.some((t) => /form[\s._-]*w/i.test(t));
    if (!hasFormWTitle) titles.unshift('FORM W');
  }

  const sheet = String(sheetName || '').trim();
  if (
    sheet &&
    sheet !== 'Sheet1' &&
    !/^form\s*w$/i.test(sheet) &&
    !titles.some((t) => normalizeMetaKey(t).includes(normalizeMetaKey(sheet).slice(0, 12)))
  ) {
    titles.unshift(sheet);
  }

  const rightFields = [];
  if (month !== '' || date !== '' || year !== '' || /month|date|year/i.test((metaLines || []).join(' '))) {
    const hasMonthDateHint = (metaLines || []).some((l) =>
      /^(month|date|year)\b/i.test(String(l || '').trim())
    );
    if (hasMonthDateHint || isFormW) {
      rightFields.push(extractFieldValue('Month', month, { requireColon: false }));
      if (isFormW || year) {
        rightFields.push(extractFieldValue('Year', year || date, { requireColon: false }));
      } else {
        rightFields.push(extractFieldValue('Date', date, { requireColon: false }));
      }
    }
  }

  return {
    titles,
    fields,
    rightFields,
    genderBox,
    hasSystemNote,
    isFormXXVI: false,
    isFormW
  };
};

/**
 * Excel-style bordered header for ALL statutory forms — vertical/horizontal line grid
 * for form name + administrative fields below.
 */
const paintBorderedStatutoryHeader = (doc, headerModel, layout, yStart) => {
  const { marginX, marginTop, marginBottom, usableWidth, pageHeight } = layout;
  let y = yStart;
  const x0 = marginX;
  const x1 = marginX + usableWidth;
  const padX = 6;
  const lineH = 10;
  const titles = Array.isArray(headerModel?.titles) ? headerModel.titles : [];
  const fields = Array.isArray(headerModel?.fields) ? headerModel.fields : [];
  const rightFields = Array.isArray(headerModel?.rightFields) ? headerModel.rightFields : [];
  const genderBox = headerModel?.genderBox || null;
  const useRightBand = rightFields.some((t) => String(t || '').trim());
  const rightBandW = useRightBand ? Math.min(150, usableWidth * 0.22) : 0;
  const splitX = useRightBand ? x1 - rightBandW : x1;
  let paintedGenderBox = false;

  const ensureSpace = (h) => {
    if (y + h > pageHeight - marginBottom) {
      doc.addPage();
      y = marginTop;
    }
  };

  const strokeRect = (x, top, w, h) => {
    doc.setDrawColor(0, 0, 0);
    doc.setLineWidth(0.7);
    doc.rect(x, top, w, h, 'S');
  };

  const paintFullBand = (text, { bold = false, size = 9, align = 'center', minH = 16 } = {}) => {
    const line = String(text || '').trim();
    if (!line) return;
    doc.setFont('helvetica', bold ? 'bold' : 'normal');
    doc.setFontSize(size);
    const wrapped = doc.splitTextToSize(line, usableWidth - padX * 2);
    const h = Math.max(minH, wrapped.length * (size + 2) + 8);
    ensureSpace(h);
    strokeRect(x0, y, usableWidth, h);
    doc.setTextColor(0, 0, 0);
    const textY = y + (h - wrapped.length * (size + 2)) / 2 + size;
    if (align === 'right') {
      doc.text(wrapped, x1 - padX, textY, { align: 'right' });
    } else if (align === 'left') {
      doc.text(wrapped, x0 + padX, textY);
    } else {
      doc.text(wrapped, (x0 + x1) / 2, textY, { align: 'center' });
    }
    y += h;
  };

  const paintSplitBand = (leftText, rightText, { minH = 16 } = {}) => {
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(8);
    const leftW = useRightBand ? splitX - x0 - padX * 2 : usableWidth - padX * 2;
    const leftWrapped = doc.splitTextToSize(String(leftText || ''), Math.max(leftW, 40));
    const rightWrapped =
      useRightBand && String(rightText || '').trim()
        ? doc.splitTextToSize(String(rightText || ''), rightBandW - padX * 2)
        : [];
    const h = Math.max(minH, Math.max(leftWrapped.length, rightWrapped.length || 1) * lineH + 6);
    ensureSpace(h);
    strokeRect(x0, y, usableWidth, h);
    if (useRightBand) {
      doc.line(splitX, y, splitX, y + h);
    }
    doc.setTextColor(0, 0, 0);
    doc.text(leftWrapped, x0 + padX, y + 10);
    if (rightWrapped.length) {
      doc.text(rightWrapped, splitX + padX, y + 10);
    }
    y += h;
  };

  /** Form W Excel-style Men / Women / young-person bordered count box. */
  const paintFormWGenderBox = () => {
    if (!genderBox || paintedGenderBox) return;
    paintedGenderBox = true;
    const labels = ['Men', 'Women', 'Male young person', 'Female young person'];
    const values = [
      genderBox.men !== '' ? genderBox.men : '0',
      genderBox.women !== '' ? genderBox.women : '0',
      genderBox.maleYoung !== '' ? genderBox.maleYoung : '0',
      genderBox.femaleYoung !== '' ? genderBox.femaleYoung : '0'
    ];
    const boxW = Math.min(460, Math.max(320, usableWidth * 0.42));
    // Excel places the box mid-sheet; keep it left of Month/Year when present.
    const boxX = useRightBand
      ? Math.max(x0, splitX - boxW - 8)
      : x0 + Math.max(0, (usableWidth - boxW) * 0.35);
    const totalH = 16;
    const labelH = 26;
    const valueH = 16;
    const h = totalH + labelH + valueH;
    ensureSpace(h + 4);

    const colW = boxW / 4;
    doc.setDrawColor(0, 0, 0);
    doc.setLineWidth(0.7);
    doc.setTextColor(0, 0, 0);

    // Total row (merged)
    doc.rect(boxX, y, boxW, totalH, 'S');
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(8);
    const totalText = `Total number of persons employed: ${genderBox.total !== '' ? genderBox.total : values.reduce((a, b) => a + (Number(b) || 0), 0)}`;
    doc.text(totalText, boxX + 4, y + 11);

    // Label row
    const labelY = y + totalH;
    labels.forEach((label, i) => {
      const cx = boxX + i * colW;
      doc.rect(cx, labelY, colW, labelH, 'S');
      doc.setFont('helvetica', 'bold');
      doc.setFontSize(7);
      const wrapped = doc.splitTextToSize(label, Math.max(colW - 4, 20)).slice(0, 3);
      const textH = wrapped.length * 8;
      doc.text(wrapped, cx + colW / 2, labelY + (labelH - textH) / 2 + 7, { align: 'center' });
    });

    // Value row
    const valueY = labelY + labelH;
    values.forEach((val, i) => {
      const cx = boxX + i * colW;
      doc.rect(cx, valueY, colW, valueH, 'S');
      doc.setFont('helvetica', 'normal');
      doc.setFontSize(9);
      doc.text(String(val), cx + colW / 2, valueY + 11, { align: 'center' });
    });

    y += h + 6;
  };

  titles.forEach((title, idx) => {
    const lower = String(title).toLowerCase();
    const isFormName = /^form\s+/i.test(title);
    const isRegister = /^register of\b|^muster\s+roll\b|^list of\b/i.test(title);
    const isRule = /see\s+rule/i.test(lower);
    paintFullBand(title, {
      bold: isFormName || isRegister || idx === 0,
      size: isFormName ? 11 : isRegister ? 10 : isRule ? 8 : 9,
      align: 'center',
      minH: isFormName ? 20 : 16
    });
  });

  fields.forEach((field, idx) => {
    const right = rightFields[idx] || '';
    if (useRightBand && idx < Math.max(rightFields.length, 2) && (right || idx < 2)) {
      paintSplitBand(field, right, { minH: 16 });
    } else {
      paintFullBand(field, { bold: false, size: 8, align: 'left', minH: 16 });
    }
    // Place gender box after establishment / total band (Excel layout)
    if (
      genderBox &&
      !paintedGenderBox &&
      (/address of the establishment|total number of persons/i.test(String(field || '')) ||
        idx === 0)
    ) {
      // Prefer after establishment line; if first field is employer, still paint once early
      if (/address of the establishment|total number of persons/i.test(String(field || ''))) {
        paintFormWGenderBox();
      }
    }
  });

  if (genderBox && !paintedGenderBox) {
    paintFormWGenderBox();
  }

  return y;
};

/**
 * Draw one sheet: meta headings once (full width), then full data table.
 */
const drawMatrixSheet = (doc, matrix, startY) => {
  const pageWidth = doc.internal.pageSize.getWidth();
  const pageHeight = doc.internal.pageSize.getHeight();
  const marginX = 24;
  const marginTop = 28;
  const marginBottom = 32;
  const usableWidth = pageWidth - marginX * 2;
  const colCount = Math.max(1, matrix.colCount);
  const rows = matrix.rows;
  const tableStart = Math.max(0, matrix.tableStartRow || 0);
  const metaLines = Array.isArray(matrix.metaLines) ? matrix.metaLines : [];
  const headerModel = buildStatutoryPdfHeaderModel(metaLines, rows, tableStart, matrix.name);

  let headerBandEnd = tableStart;
  for (let r = tableStart; r < Math.min(rows.length, tableStart + 8); r += 1) {
    if (isLikelyHeaderBandRow(rows[r], r)) headerBandEnd = r;
  }
  // Form W has a 3-row merged header band (group → mid → leaf).
  if (headerModel.isFormW) {
    for (let r = tableStart; r < Math.min(rows.length, tableStart + 5); r += 1) {
      const blob = (rows[r] || []).join(' ').toLowerCase();
      if (/basic wage|dearness|house rent|deductions|net wages|s\.?\s*no/.test(blob)) {
        headerBandEnd = Math.max(headerBandEnd, r);
      }
    }
  }
  const dayBand = detectDailyHoursBand(rows, tableStart, headerBandEnd, colCount);
  const leafHeaders = headerModel.isFormW
    ? resolveLeafHeaderTexts(rows, tableStart, headerBandEnd, colCount)
    : null;

  const weights = [];
  for (let c = 0; c < colCount; c += 1) {
    const inDayBand = dayBand && c >= dayBand.start && c <= dayBand.end;
    if (inDayBand) {
      weights.push(2.2);
      continue;
    }
    let maxLen = 4;
    let maxDataLen = 0;
    for (let r = tableStart; r < Math.min(rows.length, tableStart + 80); r += 1) {
      if (isSystemGeneratedDocumentNoteRow(rows[r])) continue;
      const len = String(rows[r][c] || '').length;
      maxLen = Math.max(maxLen, len);
      if (r > headerBandEnd) maxDataLen = Math.max(maxDataLen, len);
    }
    if (headerModel.isFormW && leafHeaders) {
      weights.push(formWTamilNaduColumnWeight(leafHeaders[c], maxDataLen || maxLen));
      continue;
    }
    const cap = colCount > 36 ? 22 : 36;
    weights.push(Math.min(Math.max(maxLen, 3), cap));
  }
  const weightSum = weights.reduce((a, b) => a + b, 0) || 1;
  const colWidths = weights.map((w) => (w / weightSum) * usableWidth);
  const colXs = [marginX];
  for (let i = 0; i < colWidths.length; i += 1) colXs.push(colXs[i] + colWidths[i]);

  const fontSize = headerModel.isFormW
    ? colCount > 28
      ? 5.2
      : 5.8
    : colCount > 40
      ? 4.5
      : colCount > 28
        ? 5
        : colCount > 20
          ? 5.5
          : colCount > 14
            ? 6.5
            : colCount > 10
              ? 7
              : 8;
  let y = startY;
  let pendingSystemNote = false;

  const paintMetaLine = (text, { bold = false, size = 9, align = 'left' } = {}) => {
    const line = String(text || '').replace(/\s+/g, ' ').trim();
    if (!line) return;
    doc.setFont('helvetica', bold ? 'bold' : 'normal');
    doc.setFontSize(size);
    doc.setTextColor(0, 0, 0);
    const wrapped = doc.splitTextToSize(line, usableWidth);
    const blockH = wrapped.length * (size + 2) + 3;
    if (y + blockH > pageHeight - marginBottom) {
      doc.addPage();
      y = marginTop;
    }
    if (align === 'center') {
      doc.text(wrapped, pageWidth / 2, y + size, { align: 'center' });
    } else if (align === 'right') {
      doc.text(wrapped, pageWidth - marginX, y + size, { align: 'right' });
    } else {
      doc.text(wrapped, marginX, y + size);
    }
    y += blockH;
  };

  const isDayBandLabelHeaderRow = (row, rowIndex) => {
    if (!dayBand || rowIndex > headerBandEnd) return false;
    if (dayBand.labelRow >= 0) return rowIndex === dayBand.labelRow;
    const blob = (row || []).map((c) => String(c || '')).join(' ').toLowerCase();
    // Fallback only when the leaf label row carries identity + wage headers together.
    return (
      /daily\s+hours\s+of\s+work/.test(blob) ||
      (/rate\s+of\s+wages?/.test(blob) && /name of the work(?:er|man)/.test(blob))
    );
  };

  const measureRowHeight = (row, rowIndex) => {
    const mergeDayLabel = isDayBandLabelHeaderRow(row, rowIndex);
    const isHeaderRow = rowIndex <= headerBandEnd;
    let maxLines = 1;
    for (let c = 0; c < colCount; c += 1) {
      if (mergeDayLabel && dayBand && c > dayBand.start && c <= dayBand.end) continue;
      const width =
        mergeDayLabel && dayBand && c === dayBand.start
          ? colXs[dayBand.end + 1] - colXs[dayBand.start] - 3
          : Math.max(colWidths[c] - 3, 6);
      const text =
        mergeDayLabel && dayBand && c === dayBand.start
          ? dayBand.label
          : String(row[c] ?? '') || ' ';
      // Data rows: never wrap pure numbers (they shrink-to-fit when painted).
      if (!isHeaderRow && isPurePdfNumericText(text)) {
        maxLines = Math.max(maxLines, 1);
        continue;
      }
      const wrapped = doc.splitTextToSize(text, Math.max(width, 6));
      const lineCap = isHeaderRow ? (headerModel.isFormW ? 6 : 8) : 8;
      maxLines = Math.max(maxLines, Math.min(wrapped.length, lineCap));
    }
    return Math.max(fontSize + 5, maxLines * (fontSize + 1.5) + 4);
  };

  const paintGridRow = (row, rowIndex, rowH) => {
    const bold = rowIndex <= headerBandEnd || isLikelyHeaderBandRow(row, rowIndex);
    const mergeDayLabel = isDayBandLabelHeaderRow(row, rowIndex);
    const isHeaderRow = rowIndex <= headerBandEnd;
    doc.setDrawColor(0, 0, 0);
    doc.setLineWidth(0.4);
    doc.setFont('helvetica', bold ? 'bold' : 'normal');
    doc.setFontSize(fontSize);
    doc.setTextColor(0, 0, 0);

    for (let c = 0; c < colCount; c += 1) {
      if (mergeDayLabel && dayBand && c > dayBand.start && c <= dayBand.end) continue;

      if (mergeDayLabel && dayBand && c === dayBand.start) {
        const bandW = colXs[dayBand.end + 1] - colXs[dayBand.start];
        doc.rect(colXs[c], y, bandW, rowH, 'S');
        const lines = doc
          .splitTextToSize(dayBand.label, Math.max(bandW - 4, 8))
          .slice(0, 3);
        const textH = lines.length * (fontSize + 1);
        doc.text(lines, colXs[c] + bandW / 2, y + (rowH - textH) / 2 + fontSize, {
          align: 'center'
        });
        continue;
      }

      doc.rect(colXs[c], y, colWidths[c], rowH, 'S');
      const raw = String(row[c] ?? '');
      if (!raw.trim()) continue;
      // Hide duplicate day-band label leftovers in non-label header rows
      if (
        dayBand &&
        rowIndex <= headerBandEnd &&
        c >= dayBand.start &&
        c <= dayBand.end &&
        /daily\s+hours/i.test(raw)
      ) {
        continue;
      }

      const cellW = Math.max(colWidths[c] - 3, 6);
      const alignRight = isPurePdfNumericText(raw);

      // Amounts / counts: single line, shrink font instead of mid-digit wrap.
      if (!isHeaderRow && alignRight) {
        let size = fontSize;
        doc.setFontSize(size);
        while (size > 3.2 && doc.getTextWidth(raw) > cellW) {
          size -= 0.4;
          doc.setFontSize(size);
        }
        doc.text(raw, colXs[c] + colWidths[c] - 1.5, y + (rowH + size) / 2 - 1, {
          align: 'right'
        });
        doc.setFontSize(fontSize);
        continue;
      }

      const lineCap = isHeaderRow && headerModel.isFormW ? 6 : 8;
      const lines = doc.splitTextToSize(raw, cellW).slice(0, lineCap);
      if (alignRight) {
        doc.text(lines, colXs[c] + colWidths[c] - 1.5, y + fontSize + 1, { align: 'right' });
      } else if (isHeaderRow) {
        const textH = lines.length * (fontSize + 1);
        doc.text(lines, colXs[c] + colWidths[c] / 2, y + (rowH - textH) / 2 + fontSize, {
          align: 'center'
        });
      } else {
        doc.text(lines, colXs[c] + 1.5, y + fontSize + 1);
      }
    }
    y += rowH;
  };

  const repeatColumnHeaders = () => {
    for (let hr = tableStart; hr <= headerBandEnd && hr < rows.length; hr += 1) {
      if (!rowHasContent(rows[hr]) && hr < headerBandEnd) continue;
      const rowH = measureRowHeight(rows[hr], hr);
      if (y + rowH > pageHeight - marginBottom) {
        doc.addPage();
        y = marginTop;
      }
      paintGridRow(rows[hr], hr, rowH);
    }
  };

  // All forms: bordered “vertical line” header (form name + fields), matching Excel model.
  if ((headerModel.titles || []).length || (headerModel.fields || []).length) {
    y = paintBorderedStatutoryHeader(
      doc,
      headerModel,
      { pageWidth, pageHeight, marginX, marginTop, marginBottom, usableWidth },
      y
    );
  } else if (matrix.name && matrix.name !== 'Sheet1') {
    y = paintBorderedStatutoryHeader(
      doc,
      { titles: [matrix.name], fields: [], rightFields: [] },
      { pageWidth, pageHeight, marginX, marginTop, marginBottom, usableWidth },
      y
    );
  }

  if (headerModel.hasSystemNote) pendingSystemNote = true;
  metaLines.forEach((line) => {
    if (isSystemGeneratedDocumentNote(line)) pendingSystemNote = true;
  });

  for (let r = tableStart; r < rows.length; r += 1) {
    const row = rows[r];
    if (!rowHasContent(row)) continue;

    // Never draw the footer inside Sr. No. — paint after the full column band.
    if (isSystemGeneratedDocumentNoteRow(row)) {
      pendingSystemNote = true;
      continue;
    }

    // Safety: duplicated long meta already shown in the bordered header — skip grid paint
    const filled = row.filter((c) => String(c || '').trim());
    const uniqueFilled = [...new Set(filled)];
    if (
      r <= headerBandEnd + 1 &&
      uniqueFilled.length === 1 &&
      uniqueFilled[0].length > 20 &&
      filled.length >= 3
    ) {
      continue;
    }

    const rowH = measureRowHeight(row, r);
    if (y + rowH > pageHeight - marginBottom) {
      doc.addPage();
      y = marginTop;
      if (r > headerBandEnd) repeatColumnHeaders();
    }
    paintGridRow(row, r, rowH);
  }

  // Also catch a note that landed in the meta band or above the table start.
  if (!pendingSystemNote) {
    for (let r = 0; r < rows.length; r += 1) {
      if (isSystemGeneratedDocumentNoteRow(rows[r])) {
        pendingSystemNote = true;
        break;
      }
    }
  }
  if (pendingSystemNote) {
    y += 8;
    paintMetaLine(SYSTEM_GENERATED_DOCUMENT_NOTE, { bold: false, size: 9, align: 'center' });
  }

  return y + 12;
};

/**
 * Convert statutory draft Excel/ZIP → PDF with all sheet data (attendance/payroll/leave
 * values included in the draft), headings shown once.
 */
export async function buildStatutoryDraftPdfBlob({
  arrayBuffer,
  fileName = 'draft.pdf',
  title = 'Statutory Draft',
  monthLabel = ''
} = {}) {
  void monthLabel;
  const excelFiles = await collectExcelBuffersFromDraft(arrayBuffer, fileName);
  if (!excelFiles.length) {
    throw new Error('No Excel data found in the draft file to convert to PDF.');
  }

  const allMatrices = [];
  for (const file of excelFiles) {
    try {
      const matrices = await collectWorkbookMatrices(file.arrayBuffer, file.label);
      matrices.forEach((m) => {
        if (excelFiles.length > 1 && (!m.name || m.name === 'Sheet1')) {
          m.name = file.label;
        }
        allMatrices.push(m);
      });
    } catch (err) {
      console.warn('Skipped draft Excel for PDF:', file.label, err);
    }
  }

  if (!allMatrices.length) {
    throw new Error('Draft Excel has no readable rows to put in the PDF.');
  }

  const maxCols = Math.max(...allMatrices.map((m) => m.colCount));
  const anyFormW = allMatrices.some((m) =>
    looksLikeFormWPdfContext(m.metaLines, m.rows, m.tableStartRow || 0)
  );
  const wide = maxCols > 8 || anyFormW;
  // Form W (~30 wage/deduction cols) needs A2 landscape so amounts stay on one line.
  // Form XXVI / day-grid musters need A3 landscape so cols 1–14 (incl. days 1–31) fit.
  const veryWide = anyFormW || maxCols > 28;
  const doc = new jsPDF({
    unit: 'pt',
    format: anyFormW ? 'a2' : veryWide ? 'a3' : 'a4',
    orientation: wide ? 'landscape' : 'portrait'
  });

  // Bordered header draws form name for every sheet — skip a floating duplicate title.
  const firstMatrix = allMatrices[0];
  const firstHeaderModel = buildStatutoryPdfHeaderModel(
    firstMatrix?.metaLines || [],
    firstMatrix?.rows || [],
    firstMatrix?.tableStartRow || 0,
    firstMatrix?.name || ''
  );
  const hasBorderedHeader =
    (firstHeaderModel.titles || []).length > 0 || (firstHeaderModel.fields || []).length > 0;
  const heading = String(title || fileName || 'Statutory Draft')
    .replace(/\.(xlsx|xls|xlsm|xlsb|zip)$/i, '')
    .slice(0, 120);
  if (!hasBorderedHeader) {
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(11);
    doc.setTextColor(0, 0, 0);
    doc.text(heading, doc.internal.pageSize.getWidth() / 2, 20, { align: 'center' });
  }

  let y = hasBorderedHeader ? 20 : 32;
  for (let i = 0; i < allMatrices.length; i += 1) {
    if (i > 0) {
      doc.addPage();
      y = 28;
    }
    y = drawMatrixSheet(doc, allMatrices[i], y);
  }

  return doc.output('blob');
}

export async function buildFormLGJDraftPdfBlob(opts = {}) {
  return buildStatutoryDraftPdfBlob(opts);
}

export function draftFileNameToPdfName(fileName) {
  const base = String(fileName || 'statutory-draft')
    .replace(/\.(xlsx|xls|xlsm|xlsb|zip)$/i, '')
    .replace(/[^\w.-]+/g, '_')
    .replace(/_+/g, '_')
    .replace(/^_|_$/g, '');
  return `${base || 'statutory-draft'}.pdf`;
}

/** Test-only helpers for Form W PDF layout detection. */
export const statutoryDraftPdfTestUtils = {
  isWageRegisterColHeaderBlob,
  isFormWAdminBandBlob,
  isFormWAdminValueRow,
  looksLikeFormWPdfContext,
  formWTamilNaduColumnWeight,
  isPurePdfNumericText,
  extractFormWGenderBox,
  sheetToDenseMatrix
};
