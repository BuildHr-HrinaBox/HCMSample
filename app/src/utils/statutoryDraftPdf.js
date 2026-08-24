import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { jsPDF } from 'jspdf';
import * as XLSX from 'xlsx';
import {
  FORM_14_RJ_DAY_ENTRIES_NOTE,
  FORM_14_RJ_FOOTER_NOTE
} from '../Pages/statutory/form14Rajasthan';

const EXCEL_EXT_RE = /\.(xlsx|xls|xlsm|xlsb)$/i;
/** Form XXVI TN needs ~44 leaf cols (9 identity + 31 days + 4 trailing). */
const MAX_PDF_COLS = 64;
const MAX_PDF_DATA_ROWS = 2500;
const MAX_ZIP_EXCEL_FILES = 60;
const MAX_TRAILING_EMPTY_AFTER_CONTENT = 2;
const SYSTEM_GENERATED_DOCUMENT_NOTE = 'This is a System Generated Document';

/**
 * True when text looks like a draft Excel/ZIP file base name
 * (e.g. "Form_W_-_TamilNadu"), not a real form heading.
 */
const looksLikeExcelDraftFileLabel = (text) => {
  const raw = String(text || '')
    .replace(EXCEL_EXT_RE, '')
    .replace(/\.zip$/i, '')
    .trim();
  if (!raw) return false;
  // Filename-style separators: Form_W_-_TamilNadu, Form-L-Gujarat.xlsx stem
  if (/_-_/.test(raw) || /__+/.test(raw)) return true;
  if (/^form[_\s.-]+[a-z0-9xivlc.]+[_\s.-]+-+[_\s.-]*[a-z]/i.test(raw)) return true;
  // Underscore-heavy catalog names that are not printed form titles
  const underscores = (raw.match(/_/g) || []).length;
  if (underscores >= 2 && /form/i.test(raw)) return true;
  return false;
};

/**
 * Excel sheet tab labels must not become PDF titles
 * (e.g. "LWF Act - Form C", bare "Form C").
 */
const looksLikeExcelSheetTabName = (text) => {
  const raw = String(text || '')
    .replace(EXCEL_EXT_RE, '')
    .replace(/\.zip$/i, '')
    .trim();
  if (!raw) return false;
  if (looksLikeExcelDraftFileLabel(raw)) return true;
  if (/\bact\s*[-–—]\s*form\b/i.test(raw)) return true;
  if (/^lwf\b/i.test(raw) && /\bform\b/i.test(raw)) return true;
  // Bare short tab names like "Form C" / "Form-C" when used as sheet.name
  if (/^form\s*[-–.]?\s*[a-z0-9xivlc.]+\s*$/i.test(raw) && raw.length <= 24) return true;
  return false;
};

/**
 * Normalize Excel soft-breaks / `_x000d_` and expand a title/meta cell into
 * separate lines (Excel wrapText headings must not become one PDF line).
 */
const normalizeStatutoryMultilineText = (raw) =>
  String(raw || '')
    .replace(/_x000d_/gi, '\n')
    .replace(/\u000d/g, '\n')
    .replace(/\r\n/g, '\n')
    .replace(/\r/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();

const TITLE_CONCAT_SPLIT_RE =
  /\s+(?=(?:REGISTER\s+OF|OVERTIME\s+MUSTER\s+ROLL|(?<!OVERTIME\s)MUSTER\s+ROLL|LIST\s+OF|WAGE\s+SLIP|LETTER\s+OF|NOTICE\s+OF|COMBINED\s+|\(Prescribed\s+under\b|\[Prescribed\s+under\b|\[?\(?See\b|THE\s+(?:TAMIL|ANDHRA|KARNATAKA|RAJASTHAN|MADHYA|GUJARAT|TELANGANA|KERALA|WEST\s+BENGAL|ODISHA|PUNJAB|HARYANA|CENTRAL)\b))/i;

const expandStatutoryMetaSegments = (raw) => {
  const normalized = normalizeStatutoryMultilineText(raw);
  if (!normalized) return [];
  let parts = normalized
    .split(/\n+/)
    .map((s) => s.replace(/[ \t\u00a0]+/g, ' ').trim())
    .filter(Boolean);

  // SheetJS / some exports flatten wrapText titles into one long line — re-split.
  if (parts.length === 1) {
    const blob = parts[0];
    const looksLikeStackedTitle =
      /^form\b/i.test(blob) &&
      blob.length > 36 &&
      (TITLE_CONCAT_SPLIT_RE.test(blob) ||
        /register of|see\s+(?:sub-)?rule|prescribed\s+under|overtime\s+muster/i.test(blob));
    if (looksLikeStackedTitle) {
      const split = blob
        .split(TITLE_CONCAT_SPLIT_RE)
        .map((s) => s.replace(/\s+/g, ' ').trim())
        .filter(Boolean);
      if (split.length > 1) parts = split;
    }
  }
  return parts;
};

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

/** Form C LWF legal footnote under the register grid. */
const FORM_C_UNPAID_ACCUMULATIONS_FOOTNOTE =
  '*See definition of "Unpaid Accumulations" under Section 2(I) of the Tamil Nadu Labour Welfare Fund Act, 1972';
const FORM_C_RJ_FOOTNOTE = '*Applicable only in case of damage/loss/fine';

const isUnpaidAccumulationsFootnoteText = (text) => {
  const norm = String(text || '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
  if (!norm) return false;
  if (/see\s+definition\s+of/.test(norm) && /unpaid\s+accumulations/.test(norm)) return true;
  if (
    /under\s+section\s*2\s*\(?\s*i\s*\)?/.test(norm) &&
    /labour\s+welfare\s+fund\s+act/.test(norm)
  ) {
    return true;
  }
  return false;
};

const isUnpaidAccumulationsFootnoteRow = (row) => {
  if (!Array.isArray(row)) return false;
  const filled = row.map((c) => String(c || '').trim()).filter(Boolean);
  if (!filled.length) return false;
  const unique = [...new Set(filled)];
  return unique.length === 1 && isUnpaidAccumulationsFootnoteText(unique[0]);
};

const looksLikeFormCRajasthanPdfContext = (metaLines, rows, sheetName = '') => {
  const blob = [...(metaLines || []), ...(rows || []).slice(0, 18).flat(), sheetName || '']
    .join(' ')
    .toLowerCase();
  return (
    /\bform\s*c\b|form[\s._-]*c[\s._-]/.test(blob) &&
    /register\s+of\s+(?:loan|recoveries|loan\s*\/\s*recoveries)/.test(blob) &&
    /rajasthan|damage\s*\/\s*loss\s*\/\s*fine/.test(blob)
  );
};

const extractUnpaidAccumulationsFootnoteText = (row) => {
  if (!Array.isArray(row)) return '';
  for (let i = 0; i < row.length; i += 1) {
    const t = String(row[i] || '')
      .replace(/\s+/g, ' ')
      .trim();
    if (isUnpaidAccumulationsFootnoteText(t)) return t;
  }
  return '';
};

/** Form 14 RJ overtime footnote — paint below the grid, above the system-generated note. */
const isForm14RajasthanOvertimeFootnoteText = (text) => {
  const norm = String(text || '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
  if (!norm) return false;
  if (/need\s+not\s+be\s+filled/.test(norm) && /commercial\s+establishments?/.test(norm)) {
    return true;
  }
  return (
    /red\s+ink/.test(norm) &&
    /overtime/.test(norm) &&
    (/days\s+column|commercial\s+establishments?|public\s+amusement/.test(norm) ||
      /^\*\s*this\s+column/.test(norm))
  );
};

const isForm14RajasthanDayEntriesNoteText = (text) => {
  const norm = String(text || '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
  if (!norm) return false;
  return (
    /note\s*:\s*entries\s+relating\s+to\s+any\s+day/.test(norm) ||
    (/entries\s+relating\s+to\s+any\s+day/.test(norm) && /made\s+on\s+that\s+day/.test(norm))
  );
};

const isForm14RajasthanFootnoteText = (text) =>
  isForm14RajasthanOvertimeFootnoteText(text) || isForm14RajasthanDayEntriesNoteText(text);

const isForm14RajasthanFootnoteRow = (row) => {
  if (!Array.isArray(row)) return false;
  const filled = row.map((c) => String(c || '').trim()).filter(Boolean);
  if (!filled.length) return false;
  const unique = [...new Set(filled)];
  // Merged footnote often collapses to one cell; fragments may repeat across the band.
  return unique.every((t) => isForm14RajasthanFootnoteText(t));
};

const extractForm14RajasthanFootnoteText = (row) => {
  if (!Array.isArray(row)) return '';
  const parts = [];
  for (let i = 0; i < row.length; i += 1) {
    const t = String(row[i] || '')
      .replace(/\s+/g, ' ')
      .trim();
    if (!t || !isForm14RajasthanFootnoteText(t)) continue;
    if (!parts.includes(t)) parts.push(t);
  }
  return parts.join(' ').replace(/\s+/g, ' ').trim();
};

const normalizeForm14RajasthanPdfFootnoteText = (text) => {
  const raw = String(text || '')
    .replace(/\s+/g, ' ')
    .trim();
  if (!raw) return '';
  if (isForm14RajasthanDayEntriesNoteText(raw)) return FORM_14_RJ_DAY_ENTRIES_NOTE;
  if (isForm14RajasthanOvertimeFootnoteText(raw)) {
    // Prefer the captured sheet wording when it is already a full sentence.
    if (raw.length > 80) return raw;
    return FORM_14_RJ_FOOTER_NOTE;
  }
  return raw;
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

/**
 * Tamil Nadu LWF Form C — "Details of Fines…" + "Quarter ending…" column headers.
 * Must stay in the grid (never promoted to full-width meta bands).
 */
const isFormCLwfColHeaderBlob = (blob) => {
  const t = String(blob || '').toLowerCase();
  if (!t) return false;
  const quarterHits = t.match(/quarter\s+ending/g);
  if (quarterHits && quarterHits.length >= 2) return true;
  return /details\s+of\s+fines/.test(t) && /quarter\s+ending/.test(t);
};

/**
 * Tamil Nadu Form 25 — S.No / Name / day-band column headers.
 * Must stay in the grid (not promoted to full-width meta).
 */
const isForm25TamilNaduColHeaderBlob = (blob) => {
  const t = String(blob || '').toLowerCase();
  if (!t) return false;
  if (
    /daily\s+hours\s+of\s+work/.test(t) &&
    (/s\.?\s*no|name of the worker|scheme of shifts/.test(t) || /\b1\b.*\b2\b.*\b3\b/.test(t))
  ) {
    return true;
  }
  if (
    /name of the worker/.test(t) &&
    /worker\s+identit|time at which work|scheme of shifts|rest\s+interval/.test(t)
  ) {
    return true;
  }
  return false;
};

/**
 * Rajasthan Form 14 — Record of Hours of Work column headers.
 * Long labels must stay in the grid (never full-width meta bands).
 */
const isForm14RajasthanColHeaderBlob = (blob) => {
  const t = String(blob || '').toLowerCase();
  if (!t) return false;
  if (/name\s+of\s+persons?\s+employ/.test(t) && /young\s+person|total\s+hours\s+worked|overtime/.test(t)) {
    return true;
  }
  if (
    /whether\s+young\s+person/.test(t) &&
    /total\s+hours\s+worked|days\s+on\s+which\s+overtime|extent\s+of\s+overtime/.test(t)
  ) {
    return true;
  }
  return false;
};

/** CLRA Form XIX — Wage Slip (Rule 78) — compact Excel layout must stay portrait in PDF. */
const looksLikeFormXIXWageSlipPdfContext = (metaLines, rows, sheetName = '') => {
  const blob = [...(metaLines || []), ...(rows || []).slice(0, 16).flat(), sheetName || '']
    .join(' ')
    .toLowerCase();
  if (!blob) return false;
  const isXix =
    /form\s*xix\b/.test(blob) ||
    /form[\s._-]*xix[\s._-]/.test(blob) ||
    (/wage\s*s?\s*slip/.test(blob) && /rule\s*78/.test(blob));
  if (!isXix) return false;
  return (
    /no\.?\s*of\s+days\s+worked|days\s+worked/.test(blob) &&
    (/overtime\s+hours|dates\s+on\s+which\s+overtime|gross\s+wages\s+payable|actual\s+wages\s+paid/.test(
      blob
    ) ||
      /rate\s+of\s+daily\s+wages/.test(blob))
  );
};

const isFormXIXWageSlipColHeaderBlob = (blob) => {
  const t = String(blob || '').toLowerCase();
  if (!t) return false;
  return (
    /no\.?\s*of\s+days\s+worked/.test(t) &&
    (/rate\s+of\s+daily\s+wages|units\s+worked|overtime/.test(t) ||
      /dates\s+on\s+which\s+overtime/.test(t))
  );
};

const FORM_XIX_PDF_WAGE_HEADERS = [
  'No. of days worked',
  'Rate of daily wages/piece - rate',
  'No. of units worked in case of piece rate',
  'Dates on which overtime worked',
  'Overtime hours and amount of overtime wages'
];

const FORM_XIX_PDF_FOOTER_LABELS = [
  'Gross wages payable',
  'Deductions, any',
  'If Actual wages paid'
];

const pickFormXIXPdfNonEmptyCells = (row) =>
  (Array.isArray(row) ? row : [])
    .map((c) => String(c || '').trim())
    .filter(Boolean);

/**
 * Collapse sparse Excel merges (A–M) into the official 5 wage columns + footer band
 * so Download PDF matches the Excel Form XIX wage-slip layout.
 */
const normalizeFormXIXWageSlipPdfMatrix = (rows, colCount, tableStartRow = 0) => {
  const src = Array.isArray(rows) ? rows : [];
  if (!src.length) return { rows: src, colCount, tableStartRow };

  let headerRow = -1;
  for (let r = Math.max(0, tableStartRow - 2); r < Math.min(src.length, tableStartRow + 8); r += 1) {
    const blob = pickFormXIXPdfNonEmptyCells(src[r]).join(' ').toLowerCase();
    if (isFormXIXWageSlipColHeaderBlob(blob) || /no\.?\s*of\s+days\s+worked/.test(blob)) {
      headerRow = r;
      break;
    }
  }
  if (headerRow < 0) {
    for (let r = 0; r < Math.min(src.length, 24); r += 1) {
      const blob = pickFormXIXPdfNonEmptyCells(src[r]).join(' ').toLowerCase();
      if (isFormXIXWageSlipColHeaderBlob(blob)) {
        headerRow = r;
        break;
      }
    }
  }
  if (headerRow < 0) return { rows: src, colCount, tableStartRow };

  // Keep original column indexes so blank "units" does not shift OT / NIL values left.
  const headerSrc = Array.isArray(src[headerRow]) ? src[headerRow] : [];
  const headerColIndexes = [];
  for (let c = 0; c < headerSrc.length; c += 1) {
    if (String(headerSrc[c] || '').trim()) headerColIndexes.push(c);
  }
  while (headerColIndexes.length < 5) {
    const last = headerColIndexes.length ? headerColIndexes[headerColIndexes.length - 1] + 1 : 0;
    headerColIndexes.push(last);
  }
  const wageCols = headerColIndexes.slice(0, 5);

  const cellAt = (rowIndex, colIndex) =>
    String(src[rowIndex]?.[colIndex] ?? '').trim();

  const headers = FORM_XIX_PDF_WAGE_HEADERS.map(
    (fallback, i) => cellAt(headerRow, wageCols[i]) || fallback
  );

  let numberRow = -1;
  for (let r = headerRow + 1; r <= Math.min(src.length - 1, headerRow + 3); r += 1) {
    const filled = wageCols.map((c) => cellAt(r, c)).filter(Boolean);
    if (filled.length >= 3 && filled.every((t) => /^\d{1,2}$/.test(t))) {
      numberRow = r;
      break;
    }
  }

  let dataRow = -1;
  for (
    let r = (numberRow >= 0 ? numberRow : headerRow) + 1;
    r <= Math.min(src.length - 1, headerRow + 6);
    r += 1
  ) {
    const filled = wageCols.map((c) => cellAt(r, c));
    if (!filled.some(Boolean)) continue;
    if (filled.filter(Boolean).every((t) => /^\d{1,2}$/.test(t))) continue;
    const blob = filled.join(' ').toLowerCase();
    if (/gross\s+wages|deductions|actual\s+wages|signature/.test(blob)) continue;
    dataRow = r;
    break;
  }

  let footerLabelRow = -1;
  let footerValueRow = -1;
  for (let r = Math.max(headerRow + 1, dataRow + 1); r < Math.min(src.length, headerRow + 14); r += 1) {
    const blob = pickFormXIXPdfNonEmptyCells(src[r]).join(' ').toLowerCase();
    if (/gross\s+wages\s+payable/.test(blob) && /deduction|actual\s+wages/.test(blob)) {
      footerLabelRow = r;
      break;
    }
    if (/gross\s+wages\s+payable/.test(blob)) {
      footerLabelRow = r;
      break;
    }
  }
  if (footerLabelRow >= 0) {
    for (let r = footerLabelRow + 1; r <= Math.min(src.length - 1, footerLabelRow + 2); r += 1) {
      const filled = pickFormXIXPdfNonEmptyCells(src[r]);
      if (!filled.length) continue;
      const blob = filled.join(' ').toLowerCase();
      if (/signature|contractor|representative/.test(blob)) continue;
      if (filled.some((t) => /^-?\d+(\.\d+)?$/.test(String(t).replace(/,/g, '')))) {
        footerValueRow = r;
        break;
      }
    }
  }

  const dataValues = wageCols.map((c) => (dataRow >= 0 ? cellAt(dataRow, c) : ''));

  // Footer labels/values sit under wage cols 1 / 3 / 5 in the Excel template.
  const footerAnchorCols = [wageCols[0], wageCols[2], wageCols[4]];
  const footerLabels = FORM_XIX_PDF_FOOTER_LABELS.map((fallback, i) => {
    if (footerLabelRow < 0) return fallback;
    const fromAnchor = cellAt(footerLabelRow, footerAnchorCols[i]);
    if (fromAnchor) return fromAnchor;
    const packed = pickFormXIXPdfNonEmptyCells(src[footerLabelRow]);
    return packed[i] || fallback;
  });

  const footerValues = ['', '', '', '', ''];
  if (footerValueRow >= 0) {
    const fromAnchors = footerAnchorCols.map((c) => cellAt(footerValueRow, c));
    if (fromAnchors.some(Boolean)) {
      footerValues[0] = fromAnchors[0];
      footerValues[2] = fromAnchors[1];
      footerValues[4] = fromAnchors[2];
    } else {
      const packed = pickFormXIXPdfNonEmptyCells(src[footerValueRow]);
      footerValues[0] = packed[0] || '';
      footerValues[2] = packed[1] || '';
      footerValues[4] = packed[2] || '';
    }
  }

  const compact = [];
  for (let r = 0; r < headerRow; r += 1) {
    compact.push(['', '', '', '', '']);
  }
  compact.push(headers);
  compact.push(['1', '2', '3', '4', '5']);
  compact.push(dataValues);
  compact.push(['', '', '', '', '']);
  compact.push([footerLabels[0], '', footerLabels[1], '', footerLabels[2]]);
  compact.push(footerValues);

  const afterFooter = Math.max(footerValueRow, footerLabelRow, dataRow, headerRow) + 1;
  for (let r = afterFooter; r < src.length; r += 1) {
    const filled = pickFormXIXPdfNonEmptyCells(src[r]);
    if (!filled.length) continue;
    compact.push([filled.join(' '), '', '', '', '']);
  }

  return {
    rows: compact,
    colCount: 5,
    tableStartRow: headerRow
  };
};

/** CLRA Form XIV / Employment Card (Rule 76) — stacked label|value card, not a multi-col register. */
const looksLikeFormXIVEmploymentCardPdfContext = (metaLines, rows, sheetName = '') => {
  const blob = [...(metaLines || []), ...(rows || []).slice(0, 24).flat(), sheetName || '']
    .join(' ')
    .toLowerCase();
  if (!blob) return false;
  // Rajasthan Form X Employment Card is a horizontal workman table — keep that layout.
  if (
    (/form[\s._-]*x[\s._-]*rj\b|\bx_rj\b/.test(blob) || /see\s+rule\s*75/.test(blob)) &&
    /employment\s+card/.test(blob) &&
    !/form\s*xiv\b/.test(blob)
  ) {
    return false;
  }
  const isXiv =
    /form\s*xiv\b/.test(blob) ||
    /form[\s._-]*xiv[\s._-]/.test(blob) ||
    (/employment\s+card/.test(blob) && /see\s+rule\s*76/.test(blob));
  if (!isXiv) return false;
  return (
    /employment\s+card/.test(blob) ||
    /name\s+of\s+the\s+workman/.test(blob) ||
    /serial\s+(?:no\.?|number)\s+in\s+the\s+register/.test(blob)
  );
};

const looksLikeFormXRajasthanEmploymentCardPdfContext = (metaLines, rows, sheetName = '') => {
  const blob = [...(metaLines || []), ...(rows || []).slice(0, 24).flat(), sheetName || '']
    .join(' ')
    .toLowerCase();
  return (
    /see\s+rule\s*75/.test(blob) &&
    /employment\s+card/.test(blob) &&
    /name\s+and\s+address\s+of\s+(?:the\s+)?contractor/.test(blob) &&
    !/form\s*xiv\b/.test(blob)
  );
};

const looksLikeFormXIRajasthanServiceCertificatePdfContext = (metaLines, rows, sheetName = '') => {
  const blob = [...(metaLines || []), ...(rows || []).slice(0, 28).flat(), sheetName || '']
    .join(' ')
    .toLowerCase();
  return (
    /see\s+rule\s*76/.test(blob) &&
    /service\s+certificate/.test(blob) &&
    /form\s*xi\b|form[\s._-]*xi[\s._-]/.test(blob)
  );
};

const FORM_XI_RJ_PDF_HEADER_LABELS = [
  'Name and address of contractor',
  'Nature and location of work',
  'Name and address of establishment under which contract is carried on',
  'Name and address of principal employer',
  'Name and address of the workman',
  'Age or date of birth',
  'Identification marks',
  "Father's / Husband's Name",
];

const matchFormXIRajasthanServiceHeaderIndex = (text) => {
  const normalized = String(text || '')
    .replace(/\.+$/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
  if (!normalized) return -1;
  if (/name\s+and\s+address\s+(?:of|if)\s+(?:the\s+)?contractor/.test(normalized)) return 0;
  if (/nature\s+and\s+location\s+of\s+work/.test(normalized)) return 1;
  if (/name\s+and\s+address\s+of\s+(?:the\s+)?establishment/.test(normalized)) return 2;
  if (/name\s+and\s+address\s+of\s+(?:the\s+)?principal\s+employer/.test(normalized)) return 3;
  if (/name\s+and\s+address\s+of\s+the\s+workm[ae]n/.test(normalized)) return 4;
  if (/age\s+or\s+date\s+of\s+birth/.test(normalized)) return 5;
  if (/identification\s+marks/.test(normalized)) return 6;
  if (/father.*husband.*name|husband.*father.*name/.test(normalized)) return 7;
  return -1;
};

const buildFormXIRajasthanServiceHeaderModel = (metaLines) => {
  const lines = [];
  (metaLines || []).forEach((raw) => {
    expandStatutoryMetaSegments(raw).forEach((line) => {
      const text = String(line || '').replace(/\s+/g, ' ').trim();
      if (text) lines.push(text);
    });
  });
  const values = FORM_XI_RJ_PDF_HEADER_LABELS.map(() => '');
  const titles = [];
  let activeIndex = -1;
  lines.forEach((line) => {
    const index = matchFormXIRajasthanServiceHeaderIndex(line);
    if (index >= 0) {
      activeIndex = index;
      const inline = line.replace(/^[^:]*:\s*/, '').trim();
      if (inline && inline !== line) values[index] = inline;
      return;
    }
    if (activeIndex >= 0 && !values[activeIndex]) {
      if (!/^form\s+xi\b|^service\s+certificate$|^see\s+rule/i.test(line)) {
        values[activeIndex] = line;
        activeIndex = -1;
        return;
      }
    }
    if (activeIndex < 0 && !/^sample$|^rj[-\s]/i.test(line)) titles.push(line);
  });
  return {
    titles,
    fields: FORM_XI_RJ_PDF_HEADER_LABELS.map((label, index) =>
      `${label}: ${values[index]}`.trimEnd()
    ),
  };
};

const looksLikeFormXVRajasthanWageSlipPdfContext = (metaLines, rows, sheetName = '') => {
  const blob = [...(metaLines || []), ...(rows || []).slice(0, 24).flat(), sheetName || '']
    .join(' ')
    .toLowerCase();
  return (
    /form\s*xv\b|form[\s._-]*xv[\s._-]/.test(blob) &&
    /wage\s*slip|wages\s+slip/.test(blob) &&
    (/see\s+rule\s*77|rajasthan|rule\s*77\s*\(\s*2\s*\)\s*\(\s*b\s*\)/.test(blob)) &&
    !/form\s*xi\b|service\s+certificate/.test(blob)
  );
};

const FORM_XV_RJ_PDF_HEADER_LABELS = [
  'Name and address of contractor',
  'Name and location of work',
  'Name and address of establishment in/under which contract is carried on',
  'Name and address of principal employer',
  "Name and Father's name of the workman",
  'Sex and identification token/ticket No.',
  'For the week/fortnight/month',
];

const matchFormXVRajasthanWageSlipHeaderIndex = (text) => {
  const normalized = String(text || '')
    .replace(/\.+$/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
  if (!normalized) return -1;
  if (/name\s+and\s+address\s+of\s+(?:the\s+)?contractor/.test(normalized)) return 0;
  if (/(?:name|nature)\s+and\s+location\s+of\s+work/.test(normalized)) return 1;
  if (/name\s+and\s+address\s+of\s+(?:the\s+)?establishment/.test(normalized)) return 2;
  if (/name\s+and\s+address\s+of\s+(?:the\s+)?principal\s+employer/.test(normalized)) return 3;
  if (/name\s+and\s+father.*workman|father.*name\s+of\s+the\s+workman/.test(normalized)) return 4;
  if (/sex\s+and\s+identification|identification\s+token|ticket\s+no/.test(normalized)) return 5;
  if (/for\s+the\s+(?:week|fortnight)|week\s*\/\s*fortnight\s*\/\s*month/.test(normalized)) return 6;
  return -1;
};

const buildFormXVRajasthanWageSlipHeaderModel = (metaLines) => {
  const lines = [];
  (metaLines || []).forEach((raw) => {
    expandStatutoryMetaSegments(raw).forEach((line) => {
      const text = String(line || '').replace(/\s+/g, ' ').trim();
      if (text) lines.push(text);
    });
  });
  const values = FORM_XV_RJ_PDF_HEADER_LABELS.map(() => '');
  const titles = [];
  let activeIndex = -1;
  lines.forEach((line) => {
    const index = matchFormXVRajasthanWageSlipHeaderIndex(line);
    if (index >= 0) {
      activeIndex = index;
      const inline = line.replace(/^[^:]*:\s*/, '').trim();
      if (inline && inline !== line) values[index] = inline;
      return;
    }
    if (activeIndex >= 0 && !values[activeIndex]) {
      if (!/^form\s+xv\b|^wages?\s+slip$|^see\s+rule/i.test(line)) {
        values[activeIndex] = line;
        activeIndex = -1;
        return;
      }
    }
    if (activeIndex < 0 && !/^sample$|^rj[-\s]/i.test(line)) titles.push(line);
  });
  return {
    titles,
    fields: FORM_XV_RJ_PDF_HEADER_LABELS.map((label, index) =>
      `${label}: ${values[index]}`.trimEnd()
    ),
  };
};

const FORM_X_RJ_EMPLOYMENT_HEADER_LABELS = [
  'Name and address of contractor',
  'Nature and location of work',
  'Name and address of establishment under which contract is carried on',
  'Name and address of principal employer',
];

const matchFormXRajasthanEmploymentHeaderIndex = (text) => {
  const normalized = String(text || '')
    .replace(/\.+$/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
  if (!normalized) return -1;
  if (/name\s+and\s+address\s+(?:of|if)\s+(?:the\s+)?contractor/.test(normalized)) return 0;
  if (/nature\s+and\s+location\s+of\s+work/.test(normalized)) return 1;
  if (/name\s+and\s+address\s+of\s+(?:the\s+)?establishment/.test(normalized)) return 2;
  if (/name\s+and\s+address\s+of\s+(?:the\s+)?principal\s+employer/.test(normalized)) return 3;
  return -1;
};

const buildFormXRajasthanEmploymentHeaderModel = (metaLines) => {
  const lines = [];
  (metaLines || []).forEach((raw) => {
    expandStatutoryMetaSegments(raw).forEach((line) => {
      const text = String(line || '').replace(/\s+/g, ' ').trim();
      if (text) lines.push(text);
    });
  });
  const titles = [];
  const fields = [];
  const values = FORM_X_RJ_EMPLOYMENT_HEADER_LABELS.map(() => '');
  let activeIndex = -1;
  lines.forEach((line) => {
    const index = matchFormXRajasthanEmploymentHeaderIndex(line);
    if (index >= 0) {
      activeIndex = index;
      const inline = line.replace(/^[^:]*:\s*/, '').trim();
      if (inline && inline !== line) values[index] = inline;
      return;
    }
    if (activeIndex >= 0 && !values[activeIndex] && !/^form\s+x\b|^employment\s+card$|^see\s+rule/i.test(line)) {
      values[activeIndex] = line;
      activeIndex = -1;
      return;
    }
    if (activeIndex < 0 && !/^sample$|^rj[-\s]/i.test(line)) titles.push(line);
  });
  FORM_X_RJ_EMPLOYMENT_HEADER_LABELS.forEach((label, index) => {
    fields.push(`${label}: ${values[index]}`.trimEnd());
  });
  return { titles, fields };
};

/** True when several workman headers sit on one row (Form X_RJ tabular card). */
const isFormXIVEmploymentCardTabularHeaderRow = (row) => {
  const filled = (Array.isArray(row) ? row : [])
    .map((c) => String(c || '').trim())
    .filter(Boolean);
  if (filled.length < 3) return false;
  const blob = filled.join(' ').toLowerCase();
  let hits = 0;
  if (/name\s+of\s+the\s+workman/.test(blob)) hits += 1;
  if (/serial|s\.?\s*no/.test(blob) && /register|workman/.test(blob)) hits += 1;
  if (/nature/.test(blob) && /employ|designat/.test(blob)) hits += 1;
  if (/wage\s+period/.test(blob)) hits += 1;
  if (/tenure|period\s+of\s+employ/.test(blob)) hits += 1;
  if (/wage/.test(blob) && /rate|piece/.test(blob)) hits += 1;
  return hits >= 3;
};

const isFormXIVEmploymentCardWorkmanLabelText = (text) => {
  const t = String(text || '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase()
    .replace(/^\d+[\.\)]\s*/, '');
  if (!t) return false;
  return (
    /name\s+of\s+the\s+workman/.test(t) ||
    (/serial|s\.?\s*no/.test(t) && /register/.test(t)) ||
    (/nature/.test(t) && /employ|designat/.test(t)) ||
    (/wage/.test(t) && /rate|piece|particular|unit/.test(t)) ||
    /wage\s+period/.test(t) ||
    /tenure\s+of\s+employ|period\s+of\s+employ/.test(t) ||
    /^remarks?$/.test(t) ||
    /date\s+of\s+entry\s+into\s+service/.test(t)
  );
};

const isFormXIVEmploymentCardHeaderLabelText = (text) => {
  const t = String(text || '')
    .replace(/\s+/g, ' ')
    .replace(/\.+$/g, '')
    .trim()
    .toLowerCase();
  if (!t) return false;
  return (
    /name\s+and\s+address\s+(?:of|if)\s+(?:the\s+)?contractor/.test(t) ||
    /name\s+and\s+address\s+of\s+(?:the\s+)?establishment|establishment\s+in\s*\/?\s*under\s+which/.test(
      t
    ) ||
    /nature\s+and\s+location\s+of\s+work|nature\s+of\s+work\s+and\s+location/.test(t) ||
    /name\s+and\s+address\s+of\s+(?:the\s+)?principal\s+employer|(?:^|:\s*)principal\s+employer/.test(
      t
    )
  );
};

/** Canonical Form XIV header labels — always shown on PDF even when values are blank. */
const FORM_XIV_PDF_HEADER_FIELD_LABELS = [
  'Name and address if contractor',
  'Name and address of Establishment in/under which contract is carried on',
  'Nature and location of work',
  'Name and address of Principal Employer',
];

const matchFormXIVPdfHeaderFieldIndex = (text) => {
  const t = String(text || '')
    .replace(/\s+/g, ' ')
    .replace(/\.+$/g, '')
    .replace(/\s*:?\s*$/, '')
    .trim()
    .toLowerCase();
  if (!t) return -1;
  if (/name\s+and\s+address\s+(?:of|if)\s+(?:the\s+)?contractor/.test(t)) return 0;
  if (
    /name\s+and\s+address\s+of\s+(?:the\s+)?establishment/.test(t) ||
    /establishment\s+in\s*\/?\s*under\s+which/.test(t)
  ) {
    return 1;
  }
  if (/nature\s+and\s+location\s+of\s+work|nature\s+of\s+work\s+and\s+location/.test(t)) return 2;
  if (/name\s+and\s+address\s+of\s+(?:the\s+)?principal\s+employer|principal\s+employer/.test(t)) {
    return 3;
  }
  return -1;
};

const FORM_XIV_PDF_STACKED_VALUE_COL_INDEX = 4; // Excel column E (1-based 5)

/**
 * Collapse sparse Employment Card Excel columns into a single label|value pair per row
 * so Download PDF does not paint empty table columns.
 * The four establishment header fields are always emitted as label|value rows (even if blank).
 */
const normalizeFormXIVEmploymentCardPdfMatrix = (rows, colCount, tableStartRow = 0) => {
  const src = Array.isArray(rows) ? rows : [];
  if (!src.length) {
    return { rows: src, colCount, tableStartRow, metaLines: null };
  }

  // Horizontal Form X_RJ workman table — leave as-is.
  for (let r = 0; r < Math.min(src.length, 30); r += 1) {
    if (isFormXIVEmploymentCardTabularHeaderRow(src[r])) {
      return { rows: src, colCount, tableStartRow, metaLines: null };
    }
  }

  // Already normalized (titles in meta + 2-col body) — do not rebuild/wipe header fields.
  if (
    Number(colCount) === 2 &&
    src.some((row) => isFormXIVEmploymentCardWorkmanLabelText(String(row?.[0] || '')))
  ) {
    return { rows: src, colCount: 2, tableStartRow, metaLines: null };
  }

  const cellAt = (row, c) => String(row?.[c] ?? '').trim();

  const splitStackedRow = (row) => {
    const cells = Array.isArray(row) ? row : [];
    const nonEmpty = [];
    for (let c = 0; c < cells.length; c += 1) {
      const t = cellAt(cells, c);
      if (t) nonEmpty.push({ c, t });
    }
    if (!nonEmpty.length) return { left: '', value: '', kind: 'empty', headerIndex: -1 };

    let ordinal = '';
    let label = '';
    let labelCol = -1;
    let value = '';
    let kind = 'other';
    let headerIndex = -1;

    for (const { c, t } of nonEmpty) {
      if (/^\d{1,2}[\.\)]?$/.test(t)) {
        ordinal = t.replace(/\)/, '.');
        if (!/\.$/.test(ordinal)) ordinal = `${ordinal}.`;
        continue;
      }
      const combined = t.match(/^(\d{1,2})[\.\)]\s+(.+)$/);
      if (combined && isFormXIVEmploymentCardWorkmanLabelText(combined[2])) {
        ordinal = `${combined[1]}.`;
        label = combined[2].trim();
        labelCol = c;
        kind = 'workman';
        continue;
      }
      if (!label && isFormXIVEmploymentCardWorkmanLabelText(t)) {
        label = t.replace(/\.+$/g, '').trim();
        labelCol = c;
        kind = 'workman';
        continue;
      }
      const hdrIdx = matchFormXIVPdfHeaderFieldIndex(t);
      if (!label && hdrIdx >= 0) {
        label = FORM_XIV_PDF_HEADER_FIELD_LABELS[hdrIdx];
        labelCol = c;
        kind = 'header';
        headerIndex = hdrIdx;
      }
    }

    // Prefer stacked value column E when present (MP/KA Employment Card template).
    const stacked = cellAt(cells, FORM_XIV_PDF_STACKED_VALUE_COL_INDEX);
    if (
      stacked &&
      stacked !== label &&
      !isFormXIVEmploymentCardWorkmanLabelText(stacked) &&
      !isFormXIVEmploymentCardHeaderLabelText(stacked) &&
      !/^\d{1,2}[\.\)]?$/.test(stacked)
    ) {
      value = stacked;
    }

    if (!value) {
      for (let i = nonEmpty.length - 1; i >= 0; i -= 1) {
        const { c, t } = nonEmpty[i];
        if (c === labelCol) continue;
        if (/^\d{1,2}[\.\)]?$/.test(t)) continue;
        if (t === label) continue;
        if (isFormXIVEmploymentCardWorkmanLabelText(t) || isFormXIVEmploymentCardHeaderLabelText(t)) {
          continue;
        }
        // Dotted underline placeholders are not values.
        if (/^\.+$/.test(t) || /^[_\s.-]{3,}$/.test(t)) continue;
        value = t;
        break;
      }
    }

    if (!label && nonEmpty.length >= 1) {
      const first = nonEmpty[0].t;
      const hdrIdx = matchFormXIVPdfHeaderFieldIndex(first);
      if (hdrIdx >= 0) {
        label = FORM_XIV_PDF_HEADER_FIELD_LABELS[hdrIdx];
        kind = 'header';
        headerIndex = hdrIdx;
        value =
          value ||
          nonEmpty
            .slice(1)
            .map((x) => x.t)
            .filter(
              (t) =>
                t !== first &&
                !isFormXIVEmploymentCardHeaderLabelText(t) &&
                !/^\.+$/.test(t)
            )
            .join(' ');
      }
    }

    if (!label && !value) {
      return {
        left: nonEmpty.map((x) => x.t).join(' '),
        value: '',
        kind: 'title',
        headerIndex: -1,
      };
    }

    let left = label || '';
    if (kind === 'workman' && ordinal && left && !/^\d{1,2}[\.\)]/.test(left)) {
      left = `${ordinal} ${left}`.replace(/\.\s+/, '. ');
    } else if (ordinal && !left) {
      left = ordinal;
    }

    return { left, value, kind: kind || 'other', headerIndex };
  };

  const metaLines = [];
  const headerValues = FORM_XIV_PDF_HEADER_FIELD_LABELS.map(() => '');
  const headerSeen = FORM_XIV_PDF_HEADER_FIELD_LABELS.map(() => false);
  const workmanRows = [];
  const trailingRows = [];

  for (let r = 0; r < src.length; r += 1) {
    const row = src[r] || [];
    const filled = row.map((c) => String(c || '').trim()).filter(Boolean);
    if (!filled.length) continue;

    const blob = filled.join(' ').toLowerCase();
    if (isSystemGeneratedDocumentNote(filled.join(' '))) {
      trailingRows.push([filled[0], '']);
      continue;
    }
    if (/signature\s+of\s+(?:the\s+)?contractor/.test(blob)) {
      trailingRows.push([filled.join(' '), '']);
      continue;
    }

    const pair = splitStackedRow(row);

    if (pair.kind === 'header' || pair.headerIndex >= 0) {
      const idx =
        pair.headerIndex >= 0 ? pair.headerIndex : matchFormXIVPdfHeaderFieldIndex(pair.left);
      if (idx >= 0) {
        headerSeen[idx] = true;
        if (pair.value && !headerValues[idx]) headerValues[idx] = pair.value;
      }
      continue;
    }

    if (pair.kind === 'workman') {
      workmanRows.push([pair.left, pair.value]);
      continue;
    }

    // Titles / company / site — meta only.
    if (workmanRows.length === 0) {
      const titleLine = pair.left || filled.join(' ');
      // Skip header-like leftovers that failed classification.
      if (matchFormXIVPdfHeaderFieldIndex(titleLine) >= 0) {
        const idx = matchFormXIVPdfHeaderFieldIndex(titleLine);
        headerSeen[idx] = true;
        if (pair.value && !headerValues[idx]) headerValues[idx] = pair.value;
        continue;
      }
      if (
        /^form\s*xiv\b/i.test(titleLine) ||
        /employment\s+card/i.test(titleLine) ||
        /see\s+rule\s*76/i.test(titleLine) ||
        (titleLine.length < 120 &&
          !isFormXIVEmploymentCardWorkmanLabelText(titleLine) &&
          !/^\d{1,2}[\.\)]/.test(titleLine))
      ) {
        metaLines.push(titleLine);
        continue;
      }
    }

    if (workmanRows.length === 0) {
      metaLines.push(pair.left || filled.join(' '));
      continue;
    }

    trailingRows.push([pair.left || filled.join(' '), pair.value || '']);
  }

  // Always show the four header fields as label|value rows (blank value if missing).
  const headerRows = FORM_XIV_PDF_HEADER_FIELD_LABELS.map((label, idx) => [
    `${label}:`,
    headerValues[idx] || '',
  ]);

  const compact = [...headerRows, ...workmanRows, ...trailingRows];

  let lastContent = -1;
  for (let r = 0; r < compact.length; r += 1) {
    if ((compact[r] || []).some((c) => String(c || '').trim())) lastContent = r;
  }
  const trimmed =
    lastContent < 0
      ? compact
      : compact.slice(0, Math.min(compact.length, lastContent + 1 + MAX_TRAILING_EMPTY_AFTER_CONTENT));

  return {
    rows: trimmed.length ? trimmed : headerRows,
    colCount: 2,
    tableStartRow: 0,
    metaLines: metaLines.length ? metaLines : null,
  };
};

const looksLikeForm14RajasthanPdfContext = (metaLines, rows, sheetName = '') => {
  const blob = [...(metaLines || []), ...(rows || []).slice(0, 12).flat(), sheetName || '']
    .join(' ')
    .toLowerCase();
  if (/form\s*14\b/.test(blob) && /rajasthan|hours\s+of\s+work|rule\s*22/.test(blob)) return true;
  if (/name\s+of\s+persons?\s+employ/.test(blob) && /whether\s+young\s+person/.test(blob)) return true;
  return /record\s+of\s+(?:the\s+)?hours\s+of\s+work/.test(blob) && /rajasthan|form\s*14\b/.test(blob);
};

/** Rajasthan Form 11 — Register of Employment. */
const looksLikeForm11RajasthanPdfContext = (metaLines, rows, sheetName = '') => {
  const blob = [...(metaLines || []), ...(rows || []).slice(0, 14).flat(), sheetName || '']
    .join(' ')
    .toLowerCase();
  if (/form\s*14\b/.test(blob) && /hours\s+of\s+work/.test(blob)) return false;
  if (/register\s+of\s+employment/.test(blob) && (/form\s*11\b/.test(blob) || /rajasthan|rule\s*22/.test(blob))) {
    return true;
  }
  if (
    /name\s+of\s+persons?\s+employ/.test(blob) &&
    /employment\s+commence/.test(blob) &&
    /total\s+hours\s+worked/.test(blob)
  ) {
    return true;
  }
  return false;
};

/** Rajasthan Form XIX — Register of Overtime templates often keep an empty column A. */
const looksLikeFormXIXRajasthanOvertimePdfContext = (metaLines, rows, sheetName = '') => {
  const blob = [...(metaLines || []), ...(rows || []).slice(0, 16).flat(), sheetName || '']
    .join(' ')
    .toLowerCase();
  return (
    /form\s*xix\b|form[\s._-]*xix[\s._-]/.test(blob) &&
    /register\s+of\s+overtime/.test(blob) &&
    (/rajasthan|\brj\b|rule\s*77/.test(blob) || /serial\s+no/.test(blob))
  );
};

const looksLikeFormAPdfContext = (metaLines, rows, sheetName = '') => {
  const blob = [...(metaLines || []), ...(rows || []).slice(0, 16).flat(), sheetName || '']
    .join(' ')
    .toLowerCase();
  return /\bform\s*a\b|form[\s._-]*a[\s._-]/.test(blob) && /employee\s+register/.test(blob);
};

/**
 * Form 11 RJ: Total hours worked during the month must be numeric or blank.
 * Clear leaked times (5PM / 9AM) and NIL from that column in the PDF matrix.
 */
const scrubForm11RajasthanTotalHoursPdfColumn = (rows, tableStartRow = 0, colCount = 0) => {
  if (!Array.isArray(rows) || rows.length === 0) return rows;
  const start = Math.max(0, Number(tableStartRow) || 0);
  const cols = Math.max(1, Number(colCount) || (rows[start] || []).length || 1);
  let hoursCol = -1;
  for (let r = start; r < Math.min(rows.length, start + 6); r += 1) {
    const row = rows[r] || [];
    for (let c = 0; c < cols; c += 1) {
      const t = String(row[c] || '')
        .replace(/\s+/g, ' ')
        .trim()
        .toLowerCase();
      if (/total\s+hours\s+worked/.test(t)) {
        hoursCol = c;
        break;
      }
    }
    if (hoursCol >= 0) break;
  }
  if (hoursCol < 0) return rows;

  let headerEnd = start;
  for (let r = start; r < Math.min(rows.length, start + 8); r += 1) {
    const blob = (rows[r] || []).join(' ').toLowerCase();
    if (/total\s+hours\s+worked|employment\s+commence|rest\s+interval|young\s+person/.test(blob)) {
      headerEnd = r;
    }
    const filled = (rows[r] || []).filter((c) => String(c || '').trim());
    if (filled.filter((c) => /^\d{1,2}$/.test(String(c).trim())).length >= 8) {
      headerEnd = Math.max(headerEnd, r);
    }
  }

  return rows.map((row, r) => {
    if (!Array.isArray(row) || r <= headerEnd) return row;
    if (isSystemGeneratedDocumentNoteRow(row)) return row;
    const next = row.slice();
    const raw = String(next[hoursCol] ?? '').trim();
    if (!raw) return next;
    // Keep pure numeric hours only.
    if (/^\d+(\.\d+)?$/.test(raw.replace(/,/g, ''))) return next;
    next[hoursCol] = '';
    return next;
  });
};

const isPureNilPdfText = (raw) => /^nill?$/i.test(String(raw || '').trim());

/** Form 25 TN muster / compensatory holidays register. */
const looksLikeForm25TamilNaduPdfContext = (metaLines, rows, sheetName = '') => {
  const blob = [...(metaLines || []), ...(rows || []).slice(0, 12).flat(), sheetName || '']
    .join(' ')
    .toLowerCase();
  if (/compensatory\s+holidays|muster\s+roll\s+and\s+register\s+of\s+compensatory/.test(blob)) {
    return true;
  }
  if (/form\s*(?:no\.?\s*)?[-–]?\s*25\b/.test(blob) && /tamil\s*nadu|daily\s+hours\s+of\s+work/.test(blob)) {
    return true;
  }
  return /form\s*(?:no\.?\s*)?[-–]?\s*25\b/.test(blob) && /prescribed\s+under\s+rules?\s*77/.test(blob);
};

/** Last 0-based column that has any non-empty cell in the matrix. */
const findLastContentColumnIndex = (rows) => {
  let last = -1;
  (rows || []).forEach((row) => {
    if (!Array.isArray(row)) return;
    for (let c = 0; c < row.length; c += 1) {
      if (String(row[c] || '').trim()) last = Math.max(last, c);
    }
  });
  return last;
};

/** 0-based Remarks column from the header band (Form 25 ends here). */
const findRemarksColumnIndex = (rows, scanRows = 20) => {
  let remarksIdx = -1;
  const limit = Math.min((rows || []).length, scanRows);
  for (let r = 0; r < limit; r += 1) {
    const row = rows[r] || [];
    for (let c = 0; c < row.length; c += 1) {
      if (/^remarks?$/i.test(String(row[c] || '').trim())) {
        remarksIdx = Math.max(remarksIdx, c);
      }
    }
  }
  return remarksIdx;
};

/**
 * Drop trailing blank columns (e.g. empty cells after Remarks on Form 25).
 * Optionally clamp to Remarks when that is the last meaningful header.
 */
const trimTrailingBlankPdfColumns = (rows, colCount, { clampToRemarks = false } = {}) => {
  let last = findLastContentColumnIndex(rows);
  if (clampToRemarks) {
    const remarksIdx = findRemarksColumnIndex(rows);
    if (remarksIdx >= 0) last = Math.min(last < 0 ? remarksIdx : last, remarksIdx);
  }
  const nextCount = Math.max(1, Math.min(colCount, last + 1));
  if (nextCount >= colCount) return { rows, colCount };
  const trimmed = (rows || []).map((row) => {
    const src = Array.isArray(row) ? row : [];
    const line = [];
    for (let c = 0; c < nextCount; c += 1) line.push(String(src[c] ?? ''));
    return line;
  });
  return { rows: trimmed, colCount: nextCount };
};

/**
 * Form 14 RJ templates often keep an empty column A left of "Name of Persons Employed".
 * Drop leading blank columns so the PDF table starts at the name column.
 */
const trimForm14RajasthanLeadingBlankPdfColumns = (rows, colCount, tableStartRow = 0) => {
  if (!Array.isArray(rows) || rows.length === 0 || colCount <= 1) {
    return { rows, colCount };
  }
  const startRow = Math.max(0, Number(tableStartRow) || 0);
  let lead = 0;

  // Prefer the official name-column header as the left edge of the table.
  outer: for (let r = startRow; r < Math.min(rows.length, startRow + 8); r += 1) {
    const row = rows[r] || [];
    for (let c = 0; c < colCount; c += 1) {
      const t = String(row[c] || '')
        .replace(/\s+/g, ' ')
        .trim()
        .toLowerCase();
      if (/name\s+of\s+persons?\s+employ/.test(t)) {
        lead = c;
        break outer;
      }
    }
  }

  // Fallback: first column with any non-footnote table-band content.
  if (lead <= 0) {
    for (let c = 0; c < colCount; c += 1) {
      let hasContent = false;
      for (let r = startRow; r < rows.length; r += 1) {
        const row = rows[r];
        if (!row) continue;
        if (isSystemGeneratedDocumentNoteRow(row) || isForm14RajasthanFootnoteRow(row)) continue;
        if (String(row[c] || '').trim()) {
          hasContent = true;
          break;
        }
      }
      if (hasContent) {
        lead = c;
        break;
      }
    }
  }

  if (lead <= 0) return { rows, colCount };

  const nextCount = Math.max(1, colCount - lead);
  const trimmed = rows.map((row) => {
    const src = Array.isArray(row) ? row : [];
    const line = [];
    for (let c = 0; c < nextCount; c += 1) line.push(String(src[lead + c] ?? ''));
    return line;
  });
  return { rows: trimmed, colCount: nextCount };
};

/** Drop leading spacer columns so the Rajasthan Form XIX table starts at Serial No. */
const trimFormXIXRajasthanLeadingBlankPdfColumns = (rows, colCount, tableStartRow = 0) => {
  if (!Array.isArray(rows) || rows.length === 0 || colCount <= 1) {
    return { rows, colCount };
  }
  const startRow = Math.max(0, Number(tableStartRow) || 0);
  let lead = 0;
  outer: for (let r = startRow; r < Math.min(rows.length, startRow + 8); r += 1) {
    const row = rows[r] || [];
    for (let c = 0; c < colCount; c += 1) {
      const text = String(row[c] || '').replace(/\s+/g, ' ').trim();
      if (/^(?:serial\s+no\.?|s\.?\s*no\.?)$/i.test(text)) {
        lead = c;
        break outer;
      }
    }
  }
  if (lead <= 0) return { rows, colCount };
  const nextCount = Math.max(1, colCount - lead);
  const trimmed = rows.map((row) => {
    const src = Array.isArray(row) ? row : [];
    return Array.from({ length: nextCount }, (_, index) => String(src[lead + index] ?? ''));
  });
  return { rows: trimmed, colCount: nextCount };
};

/** Drop the empty template column before the Rajasthan Form X workman table. */
const trimFormXRajasthanEmploymentLeadingBlankPdfColumns = (rows, colCount, tableStartRow = 0) => {
  if (!Array.isArray(rows) || rows.length === 0 || colCount <= 1) {
    return { rows, colCount };
  }
  const startRow = Math.max(0, Number(tableStartRow) || 0);
  let lead = 0;
  let foundWorkmanHeader = false;
  outer: for (let r = startRow; r < Math.min(rows.length, startRow + 8); r += 1) {
    const row = rows[r] || [];
    for (let c = 0; c < colCount; c += 1) {
      const text = String(row[c] || '').replace(/\s+/g, ' ').trim();
      if (/^name\s+of\s+the\s+workman$/i.test(text)) {
        lead = c;
        foundWorkmanHeader = true;
        break outer;
      }
    }
  }
  if (!foundWorkmanHeader) return { rows, colCount };
  const nextCount = Math.max(1, colCount - lead);
  const trimmed = rows.map((row) => {
    const src = Array.isArray(row) ? row : [];
    return Array.from({ length: nextCount }, (_, index) => String(src[lead + index] ?? ''));
  });
  return { rows: trimmed, colCount: nextCount };
};

/** True when a Form 25 body row looks like a real employee (not leftover template shell). */
const isForm25TamilNaduEmployeePdfRow = (row) => {
  if (!Array.isArray(row)) return false;
  const filled = row.map((c) => String(c || '').trim()).filter(Boolean);
  if (!filled.length) return false;
  if (isSystemGeneratedDocumentNote(filled.join(' '))) return false;
  if (filled.some((v) => /^(P|A|WO|H|L|WOP|OD|SL|CL|EL|NH|FH)$/i.test(v))) return true;
  // Worker name / ID style tokens (not times, bare numbers, or shift labels alone).
  return filled.some((v) => {
    if (/^\d{1,2}:\d{2}/.test(v)) return false;
    if (/^\d{1,4}$/.test(v)) return false;
    if (/^(general\s+shift|scheme of shifts|rest(\s+interval)?|nil|n\/?a)$/i.test(v)) return false;
    if (/daily\s+hours|time at which|worker\s+identit|serial\s+number/i.test(v)) return false;
    return /[a-z]{2,}/i.test(v);
  });
};

/**
 * Form 25 templates often leave dozens of sample body rows — trim PDF after the last real employee.
 * Keep header band + footnote / system-note rows.
 */
const trimForm25TamilNaduPdfTrailingEmployeeRows = (rows, tableStartRow = 0) => {
  if (!Array.isArray(rows) || rows.length === 0) return rows;
  let headerEnd = Math.max(0, tableStartRow);
  for (let r = tableStartRow; r < Math.min(rows.length, tableStartRow + 8); r += 1) {
    const blob = (rows[r] || []).join(' ').toLowerCase();
    if (
      isForm25TamilNaduColHeaderBlob(blob) ||
      (/^\d{1,2}$/.test(String(rows[r]?.[0] || '').trim()) &&
        (rows[r] || []).filter((c) => /^\d{1,2}$/.test(String(c || '').trim())).length >= 10)
    ) {
      headerEnd = r;
    }
  }

  let lastEmp = -1;
  for (let r = headerEnd + 1; r < rows.length; r += 1) {
    if (isForm25TamilNaduEmployeePdfRow(rows[r])) lastEmp = r;
  }
  if (lastEmp < 0) return rows;

  const kept = [];
  for (let r = 0; r < rows.length; r += 1) {
    if (r <= lastEmp) {
      kept.push(rows[r]);
      continue;
    }
    // Preserve footer notes that sit below the employee block.
    if (isSystemGeneratedDocumentNoteRow(rows[r]) || isUnpaidAccumulationsFootnoteRow(rows[r])) {
      kept.push(rows[r]);
    }
  }
  return kept;
};

/** Drop rows that belong to a leaked Excel pivot (Row Labels / Count of MALE…). */
const stripLeakedPivotRowsFromPdfMatrix = (rows, tableStartRow = 0) => {
  if (!Array.isArray(rows) || rows.length === 0) return rows;
  const out = [];
  let dropping = false;
  for (let r = 0; r < rows.length; r += 1) {
    const blob = (rows[r] || []).join(' ').toLowerCase();
    if (
      r > tableStartRow &&
      (/^row\s+labels\b/.test(blob.trim()) ||
        (/row\s+labels/.test(blob) && /count of\s+mal/.test(blob)))
    ) {
      dropping = true;
    }
    if (dropping) {
      // Keep system note if it somehow sits after the pivot block.
      if (isSystemGeneratedDocumentNoteRow(rows[r])) {
        out.push(rows[r]);
      }
      continue;
    }
    out.push(rows[r]);
  }
  return out;
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
  const earlyBlob = rows
    .slice(0, Math.min(rows.length, 14))
    .flat()
    .concat(sheetName || '')
    .join(' ')
    .toLowerCase();
  const isForm25Tn =
    /compensatory\s+holidays|form\s*(?:no\.?\s*)?[-–]?\s*25\b/.test(earlyBlob) &&
    (/prescribed\s+under\s+rules?\s*77|daily\s+hours\s+of\s+work|muster\s+roll/.test(earlyBlob) ||
      /form\s*(?:no\.?\s*)?[-–]?\s*25\b/.test(earlyBlob));

  // Wide day-grid forms (e.g. Form XXVI): keep a few empty trailing cols from !ref
  // (signature / termination). Form 25 ends at Remarks — never pad blank columns after it.
  if (
    !isForm25Tn &&
    refColCount > maxCol &&
    maxCol >= 20 &&
    refColCount - maxCol <= 8
  ) {
    maxCol = refColCount;
  }
  maxCol = Math.min(Math.max(maxCol, 1), MAX_PDF_COLS);

  // Clamp to Remarks when present so empty template columns after it are dropped.
  if (isForm25Tn) {
    const remarksIdx = findRemarksColumnIndex(rows);
    if (remarksIdx >= 0) maxCol = Math.min(maxCol, remarksIdx + 1);
  }

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
  let metaLines = [];
  for (let r = 0; r < Math.min(padded.length, 20); r += 1) {
    const filled = padded[r].filter((c) => c);
    const blob = filled.join(' ').toLowerCase();
    const isColHeader =
      (isWageRegisterColHeaderBlob(blob) && filled.length >= 3) ||
      (isFormCLwfColHeaderBlob(blob) && filled.length >= 2) ||
      (isForm25TamilNaduColHeaderBlob(blob) && filled.length >= 3) ||
      (isForm14RajasthanColHeaderBlob(blob) && filled.length >= 2) ||
      (isFormXIXWageSlipColHeaderBlob(blob) && filled.length >= 3);
    const isNumberRow =
      filled.filter((c) => /^\d{1,2}$/.test(c)).length >= Math.max(6, filled.length * 0.6);

    if (isColHeader || isNumberRow) {
      tableStartRow = r;
      break;
    }

    if (filled.length === 0) continue;

    // Form 25 festival-holiday boxes are just 1–5 — never print as underlined meta rows.
    if (
      filled.length <= 8 &&
      filled.every((t) => /^\d{1,2}$/.test(String(t || '').trim()))
    ) {
      tableStartRow = r + 1;
      continue;
    }

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
      unique.forEach((t) => {
        expandStatutoryMetaSegments(t).forEach((seg) => metaLines.push(seg));
      });
      tableStartRow = r + 1;
      continue;
    }

    // One (or few) meta values — show once as a line.
    // Do NOT promote multi-column rows of long labels (Form 14 RJ headers are all >20 chars).
    if (
      filled.length <= 2 ||
      (filled.length <= 3 && filled.every((t) => isFormMetaText(t) || t.length > 20))
    ) {
      const unique = [];
      filled.forEach((t) => {
        if (!unique.includes(t)) unique.push(t);
      });
      unique.forEach((t) => {
        expandStatutoryMetaSegments(t).forEach((seg) => metaLines.push(seg));
      });
      tableStartRow = r + 1;
      continue;
    }

    // Many columns but duplicated meta text already collapsed to col0
    if (filled.length === 1 && isFormMetaText(filled[0])) {
      expandStatutoryMetaSegments(filled[0]).forEach((seg) => metaLines.push(seg));
      tableStartRow = r + 1;
      continue;
    }

    tableStartRow = r;
    break;
  }

  let finalRows = padded;
  if (isForm25Tn) {
    finalRows = trimForm25TamilNaduPdfTrailingEmployeeRows(padded, tableStartRow);
  }

  // Strip any pivot "Row Labels" block that leaked into a statutory sheet matrix.
  finalRows = stripLeakedPivotRowsFromPdfMatrix(finalRows, tableStartRow);

  let finalColCount = maxCol;
  if (looksLikeForm14RajasthanPdfContext(metaLines, finalRows, sheetName)) {
    const trimmedLead = trimForm14RajasthanLeadingBlankPdfColumns(
      finalRows,
      finalColCount,
      tableStartRow
    );
    finalRows = trimmedLead.rows;
    finalColCount = trimmedLead.colCount;
  }
  if (looksLikeFormXIXRajasthanOvertimePdfContext(metaLines, finalRows, sheetName)) {
    const trimmedLead = trimFormXIXRajasthanLeadingBlankPdfColumns(
      finalRows,
      finalColCount,
      tableStartRow
    );
    finalRows = trimmedLead.rows;
    finalColCount = trimmedLead.colCount;
  }
  if (looksLikeFormXRajasthanEmploymentCardPdfContext(metaLines, finalRows, sheetName)) {
    const trimmedLead = trimFormXRajasthanEmploymentLeadingBlankPdfColumns(
      finalRows,
      finalColCount,
      tableStartRow
    );
    finalRows = trimmedLead.rows;
    finalColCount = trimmedLead.colCount;
  }
  if (looksLikeForm11RajasthanPdfContext(metaLines, finalRows, sheetName)) {
    finalRows = scrubForm11RajasthanTotalHoursPdfColumn(
      finalRows,
      tableStartRow,
      finalColCount
    );
  }
  if (looksLikeFormXIXWageSlipPdfContext(metaLines, finalRows, sheetName)) {
    const normalized = normalizeFormXIXWageSlipPdfMatrix(
      finalRows,
      finalColCount,
      tableStartRow
    );
    finalRows = normalized.rows;
    finalColCount = normalized.colCount;
    tableStartRow = normalized.tableStartRow;
  }
  if (looksLikeFormXIVEmploymentCardPdfContext(metaLines, finalRows, sheetName)) {
    const normalized = normalizeFormXIVEmploymentCardPdfMatrix(
      finalRows,
      finalColCount,
      tableStartRow
    );
    finalRows = normalized.rows;
    finalColCount = normalized.colCount;
    tableStartRow = normalized.tableStartRow;
    if (Array.isArray(normalized.metaLines) && normalized.metaLines.length) {
      metaLines = normalized.metaLines;
    }
  }

  return {
    name: sheetName,
    rows: finalRows,
    colCount: finalColCount,
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
      // Never fall back to matrices[0] — that merged Sheet3 pivot rows into Form 15 / Form 25.
      const matrix = byName.get(key);
      if (!matrix || !matrix.rows.length) return;

      // Form XIV Employment Card is already compacted to label|value — never enrich by
      // Excel row index (that maps title/header rows onto workman rows and drops fields).
      const formXIVAlreadyNormalized =
        looksLikeFormXIVEmploymentCardPdfContext(matrix.metaLines, matrix.rows, matrix.name) &&
        Number(matrix.colCount) === 2 &&
        matrix.rows.some(
          (row) =>
            matchFormXIVPdfHeaderFieldIndex(String(row?.[0] || '')) >= 0 ||
            isFormXIVEmploymentCardWorkmanLabelText(String(row?.[0] || ''))
        );
      if (formXIVAlreadyNormalized) {
        return;
      }

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
        // Do not extend the matrix with rows from a mismatched / larger sheet.
        if (r >= matrix.rows.length) return;

        row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
          if (nonMaster.has(`${rowNumber},${colNumber}`)) return;
          const c = colNumber - 1;
          if (c < 0 || c >= matrix.colCount || c >= MAX_PDF_COLS) return;
          const text = cellToText(cell.value);
          if (!text) return;
          if (!matrix.rows[r][c]) matrix.rows[r][c] = text;
        });
      });

      // Form 25: ExcelJS enrich can widen past Remarks — trim blank trailing cols again.
      if (
        looksLikeForm25TamilNaduPdfContext(matrix.metaLines, matrix.rows, matrix.name)
      ) {
        const trimmed = trimTrailingBlankPdfColumns(matrix.rows, matrix.colCount, {
          clampToRemarks: true
        });
        matrix.rows = trimForm25TamilNaduPdfTrailingEmployeeRows(
          trimmed.rows,
          matrix.tableStartRow || 0
        );
        matrix.colCount = trimmed.colCount;
      }

      // Form 14 RJ: drop empty template column A left of the name column.
      if (looksLikeForm14RajasthanPdfContext(matrix.metaLines, matrix.rows, matrix.name)) {
        const trimmedLead = trimForm14RajasthanLeadingBlankPdfColumns(
          matrix.rows,
          matrix.colCount,
          matrix.tableStartRow || 0
        );
        matrix.rows = trimmedLead.rows;
        matrix.colCount = trimmedLead.colCount;
      }

      // Form XIX RJ: drop the empty template column before Serial No.
      if (looksLikeFormXIXRajasthanOvertimePdfContext(matrix.metaLines, matrix.rows, matrix.name)) {
        const trimmedLead = trimFormXIXRajasthanLeadingBlankPdfColumns(
          matrix.rows,
          matrix.colCount,
          matrix.tableStartRow || 0
        );
        matrix.rows = trimmedLead.rows;
        matrix.colCount = trimmedLead.colCount;
      }

      // Form X RJ Employment Card: remove the empty first table column.
      if (looksLikeFormXRajasthanEmploymentCardPdfContext(matrix.metaLines, matrix.rows, matrix.name)) {
        const trimmedLead = trimFormXRajasthanEmploymentLeadingBlankPdfColumns(
          matrix.rows,
          matrix.colCount,
          matrix.tableStartRow || 0
        );
        matrix.rows = trimmedLead.rows;
        matrix.colCount = trimmedLead.colCount;
      }

      // Form 11 RJ: blank non-numeric Total-hours cells (no leaked 5PM / 9AM).
      if (looksLikeForm11RajasthanPdfContext(matrix.metaLines, matrix.rows, matrix.name)) {
        matrix.rows = scrubForm11RajasthanTotalHoursPdfColumn(
          matrix.rows,
          matrix.tableStartRow || 0,
          matrix.colCount
        );
      }

      // Form XIX Wage Slip: collapse sparse merges into the Excel 5-column band.
      if (looksLikeFormXIXWageSlipPdfContext(matrix.metaLines, matrix.rows, matrix.name)) {
        const normalized = normalizeFormXIXWageSlipPdfMatrix(
          matrix.rows,
          matrix.colCount,
          matrix.tableStartRow || 0
        );
        matrix.rows = normalized.rows;
        matrix.colCount = normalized.colCount;
        matrix.tableStartRow = normalized.tableStartRow;
      }

      // Form XIV Employment Card: single label|value columns (no empty table cells).
      if (looksLikeFormXIVEmploymentCardPdfContext(matrix.metaLines, matrix.rows, matrix.name)) {
        const normalized = normalizeFormXIVEmploymentCardPdfMatrix(
          matrix.rows,
          matrix.colCount,
          matrix.tableStartRow || 0
        );
        matrix.rows = normalized.rows;
        matrix.colCount = normalized.colCount;
        matrix.tableStartRow = normalized.tableStartRow;
        if (Array.isArray(normalized.metaLines) && normalized.metaLines.length) {
          matrix.metaLines = normalized.metaLines;
        }
      }

      matrix.rows = stripLeakedPivotRowsFromPdfMatrix(matrix.rows, matrix.tableStartRow || 0);
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

/**
 * Pivot / helper sheets (e.g. "Sheet3" with Row Labels + Count of MALE) must not become PDF pages.
 */
const looksLikeAuxiliaryOrPivotPdfSheet = (matrix) => {
  const name = String(matrix?.name || '').trim();
  const blob = [
    name,
    ...(matrix?.metaLines || []),
    ...(matrix?.rows || []).slice(0, 20).flat()
  ]
    .join(' ')
    .toLowerCase();
  if (/row\s+labels/.test(blob)) return true;
  if (/count of\s+mal/.test(blob) && /count of\s+fem/.test(blob)) return true;
  if (/sum of\s+gross\s+wages/.test(blob) && /sum of\s+basic/.test(blob)) return true;
  if (/pivot|count of male|count of female/.test(blob) && /sum of\s+(gross|basic|house\s+rent)/.test(blob)) {
    return true;
  }
  // Generic "Sheet3" style tab with no statutory form markers — only when clearly not a register.
  if (
    /^sheet\s*\d+$/i.test(name) &&
    !/form\s*(?:no\.?\s*)?[-–.]?\s*\d+|form\s*15|muster\s+roll|register of|prescribed under|compensatory|see\s+sub-?rule/i.test(
      blob
    )
  ) {
    if (/count of|sum of|row labels|gross wages|basic wage|admin assistant|designation/.test(blob)) {
      return true;
    }
  }
  return false;
};

/** Form 15 Part 1 TN — Register of Leave with Wages. */
const looksLikeForm15TamilNaduPdfContext = (metaLines, rows, sheetName = '') => {
  const blob = [...(metaLines || []), ...(rows || []).slice(0, 12).flat(), sheetName || '']
    .join(' ')
    .toLowerCase();
  if (/form\s*15|form-15|form_15/.test(blob) && (/leave|part\s*i\b|part\s*1\b/.test(blob) || /tamil/.test(blob))) {
    return true;
  }
  if (/register of leave/.test(blob) && (/earned leave|medical leave|see\s+sub-?rule/.test(blob))) {
    return true;
  }
  if (/earned leave/.test(blob) && /medical leave/.test(blob) && /name of the employee/.test(blob)) {
    return true;
  }
  return false;
};

/**
 * Catalog / download name says Form 15 Part 1 even when the Excel template still has FORM-X titles.
 * Used to rewrite PDF (and export) headings — never rewrite a true Form X download.
 */
const looksLikeForm15Part1PreferredContext = (preferredTitle = '', fileName = '', sheetName = '') => {
  const blob = [preferredTitle, fileName, sheetName]
    .filter(Boolean)
    .join(' ')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ');
  if (!blob) return false;
  if (/form\s*15/.test(blob) && /part\s*2|partii|part_2/.test(blob)) return false;
  if (/form\s*15/.test(blob) && /part\s*1|parti|part_1/.test(blob)) return true;
  if (
    /form\s*15/.test(blob) &&
    /register\s*of\s*leave\s*with\s*wages/.test(blob) &&
    !/part\s*2|partii|part_2/.test(blob)
  ) {
    return true;
  }
  return /form\s*15/.test(blob) && /tamil\s*nadu|tamilnadu/.test(blob) && /leave/.test(blob);
};

const FORM_15_PART1_PDF_TITLE = 'FORM-15';
const FORM_15_PART1_PDF_SUBTITLE = 'REGISTER OF LEAVE WITH WAGES';

/** Normalize dashes/spaces so "FORM–X" / "FORM X" still match. */
const normalizeFormTitleToken = (raw) =>
  String(raw || '')
    .replace(/[\u2010-\u2015\u2212\uFE58\uFE63\uFF0D]/g, '-')
    .replace(/\s+/g, ' ')
    .trim();

/** True for bare Form X only — never Form XV / XIX / XX / XXVI, etc. */
const isStandaloneFormXTitle = (raw) => {
  const t = normalizeFormTitleToken(raw);
  if (!t) return false;
  // FORM-X, FORM X, FORMX, Form - X
  if (/^form[\s._-]*x$/i.test(t)) return true;
  const compact = t.toLowerCase().replace(/[\s._-]+/g, '');
  return compact === 'formx';
};

/** Rewrite Form X template titles when the download is Form 15 Part 1. */
const rewriteForm15Part1PdfTitles = (titles) => {
  const list = Array.isArray(titles) ? [...titles] : [];
  const mapped = list
    .map((t) => {
      const raw = normalizeFormTitleToken(t);
      if (!raw) return '';
      // Drop Form X entirely — FORM-15 is ensured below (do not keep a mapped duplicate).
      if (isStandaloneFormXTitle(raw)) return '';
      if (/register\s+of\s+leave\s+and\s+social\s+security/i.test(raw)) {
        return FORM_15_PART1_PDF_SUBTITLE;
      }
      return raw;
    })
    .filter(Boolean)
    // Safety: strip any leftover Form X variant that survived mapping.
    .filter((t) => !isStandaloneFormXTitle(t));
  if (!mapped.some((t) => /form[\s._-]*15\b/i.test(t))) {
    mapped.unshift(FORM_15_PART1_PDF_TITLE);
  }
  if (!mapped.some((t) => /register\s+of\s+leave\s+with\s+wages/i.test(t))) {
    const formIdx = mapped.findIndex((t) => /form[\s._-]*15\b/i.test(t));
    mapped.splice(formIdx >= 0 ? formIdx + 1 : 0, 0, FORM_15_PART1_PDF_SUBTITLE);
  }
  return mapped.filter(
    (t, i, arr) => arr.findIndex((x) => normalizeMetaKey(x) === normalizeMetaKey(t)) === i
  );
};

/** Scrub Form X strings from meta lines before title/field extraction. */
const scrubForm15Part1MetaLines = (metaLines) =>
  (Array.isArray(metaLines) ? metaLines : [])
    .map((line) => {
      const raw = normalizeFormTitleToken(line);
      if (!raw) return '';
      if (isStandaloneFormXTitle(raw)) return '';
      if (/register\s+of\s+leave\s+and\s+social\s+security/i.test(raw)) {
        return FORM_15_PART1_PDF_SUBTITLE;
      }
      // Combined cell: "FORM-X REGISTER OF …" — drop the Form X token.
      return raw
        .replace(/\bform[\s._\-–—]*x\b(?!\s*[ivxlcdm0-9])/gi, '')
        .replace(/\s+/g, ' ')
        .trim();
    })
    .filter(Boolean);

/** Group banner labels: leave categories + Form 15 Part 2 wage/deduction bands. */
const isLeaveCategoryGroupLabel = (text) => {
  const t = String(text || '')
    .replace(/\s+/g, ' ')
    .trim();
  if (!t) return false;
  const n = t.toLowerCase();
  if (/^(earned|medical|other)\s+leave$/.test(n)) return true;
  if (/^maternity\s+benefits?$/.test(n)) return true;
  if (
    /\bleave\b/.test(n) &&
    /earned|medical|other/.test(n) &&
    !/beginning|availed|balance|period|during|earned\s+during|wages?/.test(n)
  ) {
    return true;
  }
  return false;
};

/** Form 15 Part 2 / Form W / Form XXVII style merged group banners (not leaf metric titles). */
const isWageDeductionGroupLabel = (text) => {
  const t = String(text || '')
    .replace(/\s+/g, ' ')
    .trim();
  if (!t) return false;
  const n = t.toLowerCase();
  if (/^deductions?$/.test(n)) return true;
  if (/^advances?$/.test(n)) return true;
  if (/^damages?\s*\/\s*fines?$/.test(n) || /^damages?\s+or\s+fines?$/.test(n)) return true;
  // Short "Leave Wages" group only — not the long leaf "Leave Wages (Earned…)".
  if (/^leave\s+wages?$/.test(n)) return true;
  if (/^other\s+allowances?$/.test(n)) return true;
  // Form XXVII TN Register of Wages group banners
  if (/^wages?\s+earned$/.test(n)) return true;
  if (/other\s+allowances?\s*\/\s*cash\s+payment|cash\s+payment\s+nature/.test(n)) return true;
  // Mid-tier "OTHER" over PT / Uniform Deposits (not "Other Deductions" leaf)
  if (/^other$/.test(n)) return true;
  return false;
};

const isFormBRajasthanWageRateGroupLabel = (text) => {
  const normalized = String(text || '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
  return /^rate\s+of\s+minimum\s+wages\s+and\s+since\s+the\s+date\b/.test(normalized);
};

const looksLikeFormBRajasthanPdfContext = (metaLines, rows, sheetName = '') => {
  const blob = [...(metaLines || []), ...(rows || []).slice(0, 16).flat(), sheetName || '']
    .join(' ')
    .toLowerCase();
  return (
    /\bform\s*b\b|form[\s._-]*b[\s._-]/.test(blob) &&
    /format\s+of\s+wage\s+register/.test(blob) &&
    (/rajasthan|minimum\s+wages/.test(blob) || /rate\s+of\s+minimum\s+wages/.test(blob))
  );
};

/**
 * Form VI TN — merged banner above holiday date columns:
 * "Days, dates and months of the year on which National and Festival Holidays…"
 */
const isFormVIFestivalGroupLabel = (text) => {
  const t = String(text || '')
    .replace(/\s+/g, ' ')
    .trim();
  if (!t || t.length < 20) return false;
  const n = t.toLowerCase();
  if (/days,?\s+dates\s+and\s+months/.test(n)) return true;
  if (/enter\s+days\s+dates?/.test(n) && /months/.test(n)) return true;
  if (
    /national/.test(n) &&
    /festival/.test(n) &&
    /holiday/.test(n) &&
    (/section\s*3/.test(n) || /tamil\s+nadu\s+industrial\s+establishments/.test(n))
  ) {
    return true;
  }
  return false;
};

const isStatutoryGroupHeaderLabel = (text) =>
  isLeaveCategoryGroupLabel(text) ||
  isWageDeductionGroupLabel(text) ||
  isFormVIFestivalGroupLabel(text) ||
  isFormBRajasthanWageRateGroupLabel(text);

/** True when a leaf header marks the end of a Deductions / Leave Wages group span. */
const isGroupBandStopLeaf = (text) => {
  const n = String(text || '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
  if (!n) return false;
  return (
    /^net\s+wages?$/.test(n) ||
    /^gross\s+wages?$/.test(n) ||
    /^date\s+of\s+payment$/.test(n) ||
    /^unpaid\s+accumulations?$/.test(n) ||
    /^remarks?$/.test(n) ||
    /^overtime\s+wages?$/.test(n) ||
    /^basic\s+wages?$/.test(n) ||
    /signature|thumb\s+impression|authori[sz]ed/i.test(n)
  );
};

/**
 * Leaf titles that belong outside a mid-level group (Advances / Damages),
 * so expansion does not swallow PF / Total Deductions / sibling bands.
 */
const isForeignLeafForGroup = (groupLabel, leafText) => {
  const g = String(groupLabel || '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
  const n = String(leafText || '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
  if (!n) return false;
  if (isGroupBandStopLeaf(n)) return true;
  if (/^advances?$/.test(g)) {
    return /damage|fine|any\s+other\s+deduction|total\s+deduction|provident|insurance|labour\s+welfare|leave\s+wage|gross\s+wage|overtime|basic\s+wage/.test(
      n
    );
  }
  if (/damage|fine/.test(g)) {
    return /advance|any\s+other\s+deduction|total\s+deduction|provident|insurance|labour\s+welfare|leave\s+wage|gross\s+wage|overtime|basic\s+wage/.test(
      n
    );
  }
  if (/^leave\s+wages?$/.test(g)) {
    return /gross\s+wage|overtime|basic\s+wage|deduction|provident|net\s+wage/.test(n);
  }
  if (/^wages?\s+earned$/.test(g)) {
    return (
      /gross\s+wage|deduction|provident|insurance|net\s+wage|signature|unpaid|fines?/.test(n) ||
      isGroupBandStopLeaf(n)
    );
  }
  if (/other\s+allowances?\s*\/\s*cash\s+payment|cash\s+payment\s+nature/.test(g)) {
    return (
      /gross\s+wage|basic\s+wage|dearness|deduction|provident|net\s+wage|overtime/.test(n) ||
      isGroupBandStopLeaf(n)
    );
  }
  if (/^other$/.test(g)) {
    return (
      /fines?|other\s+deduction|total\s+deduction|provident|insurance|gross\s+wage|net\s+wage|signature|unpaid/.test(
        n
      ) || isGroupBandStopLeaf(n)
    );
  }
  if (/^deductions?$/.test(g)) {
    return isGroupBandStopLeaf(n);
  }
  // Form VI festival-holiday banner — stop before Remarks / identity columns.
  if (
    /days,?\s+dates\s+and\s+months/.test(g) ||
    (/national/.test(g) && /festival/.test(g) && /holiday/.test(g) && g.length > 40)
  ) {
    return (
      /^remarks?$/.test(n) ||
      /^(?:s|sr|sl)\.?\s*no\.?$/.test(n) ||
      /employee\s+code|name\s+of\s+(?:the\s+)?(?:employee|worker)|d\.?o\.?j|date\s+of\s+join/.test(n)
    );
  }
  return false;
};

const headerBandHasLeafAt = (rows, fromRow, toRow, col) => {
  for (let r = fromRow; r <= toRow && r < rows.length; r += 1) {
    const t = String(rows[r]?.[col] || '').trim();
    if (!t) continue;
    if (isStatutoryGroupHeaderLabel(t)) continue;
    if (/^\(?\s*\d{1,2}\s*\)?$/.test(t)) continue; // column index row
    return true;
  }
  return false;
};

const readLeafTextAtCol = (rows, fromRow, toRow, col) => {
  for (let r = toRow; r >= fromRow && r < rows.length; r -= 1) {
    const t = String(rows[r]?.[col] || '')
      .replace(/\s+/g, ' ')
      .trim();
    if (!t) continue;
    if (isStatutoryGroupHeaderLabel(t)) continue;
    if (/^\(?\s*\d{1,2}\s*\)?$/.test(t)) continue;
    return t;
  }
  return '';
};

const expandStatutoryGroupBandEnd = (rows, groupRow, start, colCount, headerBandEnd, nextGroupStart, groupLabel) => {
  if (Number.isFinite(nextGroupStart) && nextGroupStart > start) {
    return nextGroupStart - 1;
  }
  let end = start;
  for (let c = start + 1; c < colCount; c += 1) {
    const group = String(rows[groupRow]?.[c] || '')
      .replace(/\s+/g, ' ')
      .trim();
    if (group) {
      if (isStatutoryGroupHeaderLabel(group)) break;
      // Non-group text on the same banner row ends the span.
      break;
    }
    const leaf = readLeafTextAtCol(rows, groupRow + 1, headerBandEnd, c);
    if (leaf && isForeignLeafForGroup(groupLabel, leaf)) break;
    if (leaf || headerBandHasLeafAt(rows, groupRow + 1, headerBandEnd, c) || c === start) {
      end = c;
      continue;
    }
    // Allow a single blank inside a merge; stop on a longer empty run.
    const nextLeaf = headerBandHasLeafAt(rows, groupRow + 1, headerBandEnd, c + 1);
    if (!nextLeaf) break;
    end = c;
  }
  return Math.max(start, end);
};

/**
 * Detect merged group header bands for leave registers and Form 15 Part 2
 * (Deductions / Advances / Damages / Leave Wages). Excel merges collapse to the
 * top-left cell — PDF must re-span those columns and center the label.
 */
const detectStatutoryGroupHeaderBands = (rows, tableStart, headerBandEnd, colCount) => {
  const bands = [];
  for (let r = tableStart; r <= headerBandEnd && r < rows.length; r += 1) {
    const found = [];
    for (let c = 0; c < colCount; c += 1) {
      const t = String(rows[r]?.[c] || '')
        .replace(/\s+/g, ' ')
        .trim();
      if (isStatutoryGroupHeaderLabel(t)) found.push({ start: c, label: t });
    }
    if (!found.length) continue;
    found.forEach((item, i) => {
      const nextStart = i + 1 < found.length ? found[i + 1].start : null;
      const end = expandStatutoryGroupBandEnd(
        rows,
        r,
        item.start,
        colCount,
        headerBandEnd,
        nextStart,
        item.label
      );
      if (end > item.start) {
        bands.push({
          labelRow: r,
          start: item.start,
          end,
          label: item.label
        });
      }
    });
  }
  return bands;
};

/** @deprecated alias — leave-only detection now covered by detectStatutoryGroupHeaderBands. */
const detectLeaveCategoryBands = (rows, tableStart, headerBandEnd, colCount) =>
  detectStatutoryGroupHeaderBands(rows, tableStart, headerBandEnd, colCount).filter((b) =>
    isLeaveCategoryGroupLabel(b.label)
  );

/**
 * When the workbook contains a primary statutory register, drop unrelated sheets
 * so PDF is only the form — not Sheet2/Sheet3 pivots that inflate page count.
 */
const filterStatutoryPdfMatrices = (matrices) => {
  const list = Array.isArray(matrices) ? matrices.filter(Boolean) : [];
  if (!list.length) return list;

  // Always drop pivot / helper sheets first.
  const withoutAux = list.filter((m) => !looksLikeAuxiliaryOrPivotPdfSheet(m));
  const pool = withoutAux.length > 0 ? withoutAux : list;

  const form25 = pool.filter((m) =>
    looksLikeForm25TamilNaduPdfContext(m.metaLines, m.rows, m.name)
  );
  if (form25.length > 0) return form25;

  const form15 = pool.filter((m) =>
    looksLikeForm15TamilNaduPdfContext(m.metaLines, m.rows, m.name)
  );
  if (form15.length > 0) return form15;

  // Prefer real form/register sheets over bare SheetN tabs.
  const formish = pool.filter((m) => {
    const blob = [...(m.metaLines || []), ...(m.rows || []).slice(0, 8).flat(), m.name || '']
      .join(' ')
      .toLowerCase();
    return (
      /\bform\b|register of|muster roll|prescribed under|see\s+sub-?rule|earned leave/i.test(blob) &&
      !looksLikeAuxiliaryOrPivotPdfSheet(m)
    );
  });
  if (formish.length > 0) return formish;

  return pool;
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
  if (rowIndex < 20 && isForm14RajasthanColHeaderBlob(blob) && filled.length >= 2) {
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

/** Form 26 / 26-A / Form 1 TN style "Nil of the month" / "Nil for the month of …" cell text. */
const isNilOfTheMonthPdfText = (text) => {
  const t = String(text || '').trim();
  return /^nill?\s+of\s+the\s+month\b/i.test(t) || /^nil\s+for\s+the\s+month\b/i.test(t);
};

/** Form 1 TN template spill: signature block / payment date in the amount column. */
const isNilOfTheMonthIgnorableSpillPdfText = (text) => {
  const raw = String(text || '')
    .replace(/_x000d_/gi, '\n')
    .replace(/\r\n/g, '\n')
    .replace(/\r/g, '\n')
    .trim();
  if (!raw) return true;
  const flat = raw.replace(/\s+/g, ' ').trim();
  if (/^for\s*\(/i.test(flat)) return true;
  if (/authorised\s+signatory|authorized\s+signatory/i.test(flat)) return true;
  if (/signature\s+of\s+employer/i.test(flat)) return true;
  if (/manager\s*\/\s*authori[sz]ed\s+person/i.test(flat)) return true;
  // Bare calendar date (31-08-2026) left from template / pay-run stamp
  if (/^\d{1,2}[-/.]\d{1,2}[-/.]\d{2,4}$/.test(flat)) return true;
  // Multiline spill: date + For (… + Authorised Signatory
  if (
    /\d{1,2}[-/.]\d{1,2}[-/.]\d{2,4}/.test(flat) &&
    (/for\s*\(/i.test(flat) || /authori[sz]ed\s+signatory/i.test(flat) || /signature\s+of\s+employer/i.test(flat))
  ) {
    return true;
  }
  return false;
};

/**
 * Extra values allowed on a nil-of-month data row (Form 1 TN puts serial + month
 * beside "Nill of the month" instead of a single merged cell).
 */
const isNilOfTheMonthCompanionPdfText = (text) => {
  const t = String(text || '').trim();
  if (!t) return true;
  if (isNilOfTheMonthIgnorableSpillPdfText(t)) return true;
  if (isNilOfTheMonthPdfText(t)) return true;
  if (/^nill?$/i.test(t)) return true;
  if (/^\d{1,4}$/.test(t)) return true; // Sl.No
  // Aug 2026 / August 2026
  if (/^(jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)[a-z]*\.?\s+\d{4}$/i.test(t)) {
    return true;
  }
  if (/^\d{1,2}[-/]\d{4}$/.test(t) || /^\d{4}[-/]\d{1,2}$/.test(t)) return true;
  return false;
};

const NIL_OF_THE_MONTH_PDF_DISPLAY = 'Nil of the Month';

const NIL_OF_THE_MONTH_YEAR_RE =
  /\b((?:jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)[a-z]*\.?)\s+(\d{4})\b/i;

/** Normalize "Aug 2026" / "August 2026" → "Aug 2026". */
const formatNilOfTheMonthYearLabel = (monthToken, yearToken) => {
  const abbr = String(monthToken || '')
    .replace(/\./g, '')
    .slice(0, 3);
  const y = String(yearToken || '').trim();
  if (!abbr || !/^\d{4}$/.test(y)) return '';
  return `${abbr.charAt(0).toUpperCase()}${abbr.slice(1).toLowerCase()} ${y}`;
};

/**
 * Prefer "Nil of the Month Aug 2026" when the period is on the nil line or a companion cell.
 */
const formatNilOfTheMonthPdfDisplay = (nilText, filledTexts = []) => {
  const fromNil = String(nilText || '').match(NIL_OF_THE_MONTH_YEAR_RE);
  if (fromNil) {
    const label = formatNilOfTheMonthYearLabel(fromNil[1], fromNil[2]);
    return label ? `${NIL_OF_THE_MONTH_PDF_DISPLAY} ${label}` : NIL_OF_THE_MONTH_PDF_DISPLAY;
  }
  for (const t of filledTexts) {
    if (isNilOfTheMonthPdfText(t)) continue;
    const m = String(t || '').trim().match(
      /^(jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)[a-z]*\.?\s+(\d{4})$/i
    );
    if (!m) continue;
    const label = formatNilOfTheMonthYearLabel(m[1], m[2]);
    if (label) return `${NIL_OF_THE_MONTH_PDF_DISPLAY} ${label}`;
  }
  return NIL_OF_THE_MONTH_PDF_DISPLAY;
};

/**
 * Detect a data row that is a nil-month notice (after Excel merges collapse).
 * Returns a full-width span so PDF paints one merged, center-aligned box.
 * Form 1 TN / Form XXIX: "1 | Nill of the month | Aug 2026 | …" → "Nil of the Month Aug 2026".
 */
const resolveNilOfTheMonthPdfSpan = (row, colCount, rowIndex, headerBandEnd) => {
  if (rowIndex <= headerBandEnd || colCount < 2) return null;
  const filled = [];
  let nilText = '';
  for (let c = 0; c < colCount; c += 1) {
    const t = String(row?.[c] ?? '').trim();
    if (!t) continue;
    if (isNilOfTheMonthIgnorableSpillPdfText(t)) continue;
    filled.push(t);
    if (isNilOfTheMonthPdfText(t)) nilText = t;
  }
  if (!filled.length || !nilText) return null;
  if (!filled.every((t) => isNilOfTheMonthCompanionPdfText(t))) return null;
  return {
    start: 0,
    end: colCount - 1,
    text: formatNilOfTheMonthPdfDisplay(nilText, filled)
  };
};

/** True when a column header is a serial / Sl. No. style label. */
const isPdfSerialNumberHeader = (headerText) => {
  const h = String(headerText || '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
  if (!h) return false;
  if (/^\(?\s*\d{1,2}\s*\)?$/.test(h)) return false;
  return /^(?:s|sr|si|sl)\.?\s*no\.?$|^serial\s*(?:no\.?|number)\b|^sl\.?\s*no\b/i.test(h);
};

/**
 * Weight a column from its header name (+ data length) so PDF widths
 * follow column titles automatically for every statutory form.
 */
const statutoryHeaderColumnWeight = (headerText, maxDataLen = 0, colCount = 12) => {
  const h = String(headerText || '')
    .replace(/\s+/g, ' ')
    .trim();
  const lower = h.toLowerCase();
  const headerLen = h.length;
  const dataLen = Math.max(0, Number(maxDataLen) || 0);
  const wideForm = colCount > 16;

  if (isPdfSerialNumberHeader(h)) return Math.max(10, Math.min(12, 5 + Math.min(dataLen, 3)));

  // Short categorical headers — keep a floor so Sex/Age/Photo do not collapse.
  if (/^(sex|gender|age|photo|male|female)$/i.test(lower)) return 5.5;
  if (/^(nil|remarks?)$/i.test(lower)) return Math.max(5.5, Math.min(8, 4 + dataLen * 0.25));

  // ESI / insurance / id
  if (/insurance\s*no|esi\s*no|identification/i.test(lower)) {
    return Math.max(6.5, Math.min(9, 5 + Math.min(dataLen, 8) * 0.35));
  }

  // Date / time short labels
  if (/^(date|time)\b/i.test(lower) && headerLen <= 28) {
    return Math.max(5.5, Math.min(8.5, 4.5 + Math.min(dataLen, 12) * 0.25));
  }

  // Name / address identity columns
  if (/name.*address|address.*name|name of|injured person|name\s*&\s*address|name of persons?\s+employ/i.test(lower)) {
    return Math.max(12, Math.min(wideForm ? 18 : 20, Math.max(headerLen * 0.16, dataLen * 0.45, 12)));
  }

  // Form 14 RJ young-person / hours / overtime columns — keep readable in landscape.
  if (/young\s+person|total\s+hours\s+worked|days\s+on\s+which\s+overtime|extent\s+of\s+overtime/i.test(lower)) {
    return Math.max(10, Math.min(16, headerLen * 0.18 + Math.min(dataLen, 12) * 0.3));
  }

  // Long Form 11-style narrative headers
  if (headerLen >= 55) {
    return Math.max(12, Math.min(wideForm ? 20 : 22, headerLen * 0.2 + Math.min(dataLen, 10) * 0.25));
  }
  if (headerLen >= 28) {
    return Math.max(8.5, Math.min(16, headerLen * 0.2 + Math.min(dataLen, 14) * 0.3));
  }

  const fromHeader = Math.min(Math.max(headerLen * 0.22, 4.5), wideForm ? 12 : 14);
  const fromData = Math.min(Math.max(dataLen * 0.45, 3.5), wideForm ? 11 : 14);
  return Math.max(fromHeader, fromData, 4.5);
};

/**
 * Raise columns that fell below a header-based minimum, then renormalize to usableWidth.
 */
const enforcePdfColumnMinWidths = (widths, headers, usableWidth, options = {}) => {
  const src = Array.isArray(widths) ? widths.map((w) => Math.max(0, Number(w) || 0)) : [];
  if (!src.length) return src;
  const total = Math.max(1, Number(usableWidth) || src.reduce((a, b) => a + b, 0));
  const compactSerial = options?.compactSerial === true;
  const mins = src.map((_, i) => {
    const h = String(headers?.[i] || '')
      .replace(/\s+/g, ' ')
      .trim();
    const lower = h.toLowerCase();
    if (isPdfSerialNumberHeader(h)) {
      return compactSerial ? Math.max(32, total * 0.045) : Math.max(52, total * 0.08);
    }
    if (/^(sex|gender|age|photo)$/i.test(h)) return Math.max(28, total * 0.045);
    if (h.length >= 55) return Math.max(56, total * 0.09);
    if (h.length >= 28 || /name|address|witness|signature|occupation|department/i.test(lower)) {
      return Math.max(36, total * 0.06);
    }
    return Math.max(18, total * 0.03);
  });

  let next = src.map((w, i) => Math.max(w, mins[i]));
  let sum = next.reduce((a, b) => a + b, 0) || 1;
  if (sum > total) {
    let overflow = sum - total;
    for (let pass = 0; pass < 8 && overflow > 0.01; pass += 1) {
      const shrinkable = next.map((w, i) => Math.max(0, w - mins[i]));
      const shrinkSum = shrinkable.reduce((a, b) => a + b, 0);
      if (shrinkSum < 0.01) {
        // All columns already at mins — scale as a last resort.
        next = next.map((w) => (w / (sum || 1)) * total);
        sum = total;
        overflow = 0;
        break;
      }
      const take = Math.min(overflow, shrinkSum);
      next = next.map((w, i) => w - (take * shrinkable[i]) / shrinkSum);
      sum = next.reduce((a, b) => a + b, 0) || 1;
      overflow = sum - total;
    }
  } else if (Math.abs(sum - total) > 0.01) {
    const extra = total - sum;
    const growable = next.map((w, i) => Math.max(0, w - mins[i]));
    const growSum = growable.reduce((a, b) => a + b, 0);
    if (growSum > 0.01) {
      next = next.map((w, i) => w + (extra * growable[i]) / growSum);
    } else {
      next = next.map((w) => (w / sum) * total);
    }
  }
  // Final normalize only when needed — avoid FP shrink below hard mins.
  const finalSum = next.reduce((a, b) => a + b, 0) || 1;
  if (Math.abs(finalSum - total) <= 0.05) {
    return next;
  }
  return next.map((w) => (w / finalSum) * total);
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

/** Form XXVII TN Register of Wages — keep identity cols readable; amount cols compact. */
const formXXVIITamilNaduColumnWeight = (headerText, maxDataLen) => {
  const h = String(headerText || '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
  if (!h) return Math.max(4.5, Math.min(maxDataLen || 3, 7));
  if (/^(?:s|sr|si|sl)\.?\s*no\b|^serial\s*(?:no|number)/.test(h)) return 3.2;
  if (/name of(?:\s+the)?\s+(?:work(?:man|er)|employee)/.test(h)) return 11;
  if (/father|husband/.test(h)) return 8;
  if (/designation|employee\s*(?:number|id|code)/.test(h)) return 6.5;
  if (/signature|thumb\s+impression|cheque/.test(h)) return 8;
  if (/unpaid|umpaid/.test(h)) return 6.5;
  if (/wage\s*period|number\s+worked|units\s+of\s+work|days\s+worked/.test(h)) return 5;
  if (/daily\s+rated|piece\s+rate|overtime\s+rate/.test(h)) return 5.5;
  if (
    /basic\s+wage|dearness|wash\s+allow|\bhra\b|\bstb\b|cash\s+in\s+lieu|ecca|gross\s+wages?|net\s+wages?|providen|esi|fines?|other\s+deduction|total\s+deduction|\bpt\b|uniform/.test(
      h
    )
  ) {
    return Math.max(5.5, Math.min(Math.max(maxDataLen || 0, 5) + 1.2, 8.5));
  }
  return Math.max(5, Math.min(maxDataLen || 4, 7.5));
};

const resolveLeafHeaderTexts = (rows, tableStart, headerBandEnd, colCount) => {
  const headers = Array.from({ length: colCount }, () => '');
  for (let c = 0; c < colCount; c += 1) {
    let best = '';
    let numberFallback = '';
    for (let r = headerBandEnd; r >= tableStart; r -= 1) {
      const t = String(rows[r]?.[c] || '')
        .replace(/\s+/g, ' ')
        .trim();
      if (!t) continue;
      // Column index row "(9)" / "9" — keep looking upward for the real label.
      if (/^\(?\s*\d{1,2}\s*\)?$/.test(t)) {
        if (!numberFallback) numberFallback = t;
        continue;
      }
      // Prefer the deepest (leaf) non-group label; skip ultra-wide group banners alone.
      if (
        /^deductions$/i.test(t) ||
        /^advances$/i.test(t) ||
        /^damages\s*\/\s*fine$/i.test(t) ||
        /^details of injury$/i.test(t)
      ) {
        if (!best) best = t;
        continue;
      }
      best = t;
      break;
    }
    headers[c] = best || numberFallback;
  }
  return headers;
};

const stripFieldLabelPrefix = (text, labelRe) =>
  String(text || '')
    .replace(labelRe, '')
    .replace(/^[\s.:\-–—]+/, '')
    .trim();

/**
 * Tamil Nadu CLRA Form XXVII — Register of Wages [Rule 78(1)(a)].
 * Not quarterly returns (Form XXVII AP / other states).
 */
const looksLikeFormXXVIITamilNaduRegisterPdfContext = (
  metaLines,
  rows,
  sheetName = '',
  fileName = ''
) => {
  const blob = [
    ...(metaLines || []),
    ...(rows || []).slice(0, 14).flat(),
    sheetName || '',
    fileName || ''
  ]
    .join(' ')
    .toLowerCase();
  if (!blob.trim()) return false;
  if (/quarterly\s+return|form[\s._-]*xxvii[\s._-]*quarter|form[\s._-]*27[\s._-]*quarter/.test(blob)) {
    return false;
  }
  const isXxvii =
    /form\s*xxvii\b/.test(blob) ||
    /form[\s._-]*xxvii/.test(blob) ||
    /form_xxvii/.test(blob) ||
    (/form[\s._-]*27\b/.test(blob) && /tamil/.test(blob));
  const isWagesRegister =
    /register\s+of\s+wages/.test(blob) ||
    /rule\s*78\s*\(\s*1\s*\)\s*\(\s*a\s*\)/.test(blob) ||
    (/basic\s+wage/.test(blob) &&
      /gross\s+wages?/.test(blob) &&
      /net\s+wages?/.test(blob) &&
      /deductions?/.test(blob));
  return isXxvii && isWagesRegister;
};

/** Pull month name from Form XXVII "Wage Period : May" banner (not body WAGE PERIOD column). */
const extractMonthFromWagePeriodLine = (line) => {
  const t = String(line || '').replace(/\s+/g, ' ').trim();
  if (!t) return '';
  // Skip Form W style "Wage Period from 1st May … to …"
  if (/wage\s*period\s+from\b/i.test(t)) return '';
  const m = t.match(
    /^wage\s*period\s*:?\s*(january|february|march|april|may|june|july|august|september|october|november|december)\b/i
  );
  if (m?.[1]) return m[1];
  return '';
};

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
  if (/^form\s*(?:no\.?\s*)?[-–.]?\s*[a-z0-9xivlc.]+\b/i.test(t) && t.length < 48) return true;
  if (/see\s+(?:sub-)?rule|\[see\s+|^\(see\s+/i.test(lower)) return true;
  if (/^\(?prescribed\s+under\b|^\[prescribed\s+under\b/i.test(t)) return true;
  if (/^register of\b/i.test(t)) return true;
  if (/^overtime\s+muster\s+roll\b|^muster\s+roll\b/i.test(t)) return true;
  if (/^list of\b/i.test(t) && t.length < 60) return true;
  if (/^wage\s+slip\b|^letter\s+of\b|^notice\s+of\b|^combined\b/i.test(t) && t.length < 80) {
    return true;
  }
  // Act / rules banner under the form name (Excel line 3–4)
  if (/^the\s+.+\b(act|rules)\b/i.test(t) && t.length <= 180 && !/:/.test(t)) return true;
  // License / return style short titles without a field colon
  if (
    !/:/.test(t) &&
    t.length <= 100 &&
    /form|register|return|notice|accident|wages|employment|muster|overtime/i.test(lower)
  ) {
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
 * @param {object} [opts]
 * @param {string} [opts.preferredTitle] catalog form name (e.g. "Form 15 Part 1")
 * @param {string} [opts.fileName] draft / template file name
 */
const buildStatutoryPdfHeaderModel = (metaLines, rows, tableStart, sheetName = '', opts = {}) => {
  const preferredTitle = opts?.preferredTitle || '';
  const fileName = opts?.fileName || '';
  const preferForm15Part1 = looksLikeForm15Part1PreferredContext(
    preferredTitle,
    fileName,
    sheetName
  );
  const effectiveMetaLines = preferForm15Part1
    ? scrubForm15Part1MetaLines(metaLines)
    : metaLines;

  const xxvi = detectFormXXVIPdfLayout(effectiveMetaLines, rows, tableStart);
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
      hasSystemNote: (effectiveMetaLines || []).some((l) => isSystemGeneratedDocumentNote(l)),
      isFormXXVI: true,
      isFormW: false
    };
  }

  if (looksLikeFormXRajasthanEmploymentCardPdfContext(effectiveMetaLines, rows, sheetName)) {
    const employmentHeader = buildFormXRajasthanEmploymentHeaderModel(effectiveMetaLines);
    return {
      titles: employmentHeader.titles,
      fields: employmentHeader.fields,
      rightFields: [],
      genderBox: null,
      hasSystemNote: (effectiveMetaLines || []).some((l) => isSystemGeneratedDocumentNote(l)),
      isFormXXVI: false,
      isFormW: false,
      isFormXXVIIRegister: false,
      titleBoxFullBorder: false,
      hideRightBandSplit: false,
    };
  }

  if (looksLikeFormXIRajasthanServiceCertificatePdfContext(effectiveMetaLines, rows, sheetName)) {
    const serviceHeader = buildFormXIRajasthanServiceHeaderModel(effectiveMetaLines);
    return {
      titles: serviceHeader.titles,
      fields: serviceHeader.fields,
      rightFields: [],
      genderBox: null,
      hasSystemNote: (effectiveMetaLines || []).some((l) => isSystemGeneratedDocumentNote(l)),
      isFormXXVI: false,
      isFormW: false,
      isFormXXVIIRegister: false,
      titleBoxFullBorder: false,
      hideRightBandSplit: false,
    };
  }

  if (looksLikeFormXVRajasthanWageSlipPdfContext(effectiveMetaLines, rows, sheetName)) {
    const wageSlipHeader = buildFormXVRajasthanWageSlipHeaderModel(effectiveMetaLines);
    return {
      titles: wageSlipHeader.titles,
      fields: wageSlipHeader.fields,
      rightFields: [],
      genderBox: null,
      hasSystemNote: (effectiveMetaLines || []).some((l) => isSystemGeneratedDocumentNote(l)),
      isFormXXVI: false,
      isFormW: false,
      isFormXXVIIRegister: false,
      titleBoxFullBorder: false,
      hideRightBandSplit: false,
    };
  }

  const isFormW = looksLikeFormWPdfContext(effectiveMetaLines, rows, tableStart);
  const isFormXXVIIRegister = looksLikeFormXXVIITamilNaduRegisterPdfContext(
    effectiveMetaLines,
    rows,
    sheetName,
    fileName
  );
  const genderBox = isFormW ? extractFormWGenderBox(effectiveMetaLines, rows, tableStart) : null;

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

  (effectiveMetaLines || []).forEach((raw) => {
    expandStatutoryMetaSegments(raw).forEach((line) => {
      if (!line) return;
      if (preferForm15Part1 && isStandaloneFormXTitle(line)) return;
      if (isSystemGeneratedDocumentNote(line)) {
        hasSystemNote = true;
        return;
      }
      // Festival holiday boxes on Form 25 are numbered 1–5 — never print as title rows.
      if (/^\d{1,2}$/.test(line)) return;
      // Form 14 RJ column titles must never become full-width meta bands.
      if (
        /name\s+of\s+persons?\s+employ/i.test(line) ||
        /whether\s+young\s+person/i.test(line) ||
        /total\s+hours\s+worked\s+during/i.test(line) ||
        /days\s+on\s+which\s+overtime/i.test(line) ||
        /extent\s+of\s+overtime\s+worked/i.test(line)
      ) {
        return;
      }
      // Form 14 RJ footnotes are painted under the grid (above system note).
      if (isForm14RajasthanFootnoteText(line)) return;
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
        month = stripFieldLabelPrefix(line, /^month/i) || month;
        return;
      }
      // Form XXVII banner: "Wage Period : May" → Month right field (not a left meta row)
      if (/^wage\s*period\s*:?\s*/i.test(line) && !/wage\s*period\s+from\b/i.test(line)) {
        const fromWp = extractMonthFromWagePeriodLine(line);
        if (fromWp) month = month || fromWp;
        else {
          const rest = stripFieldLabelPrefix(line, /^wage\s*period/i);
          if (rest && !/weekly|monthly|fn\b/i.test(rest)) month = month || rest;
        }
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
      if (line.length <= 90 && !/:/.test(line)) pushUnique(titles, line, { asTitle: true });
      else pushUnique(fields, line);
    });
  });

  if (isFormW) {
    const hasFormWTitle = titles.some((t) => /form[\s._-]*w/i.test(t));
    if (!hasFormWTitle) titles.unshift('FORM W');
  }

  const sheet = String(sheetName || '').trim();
  // Never promote Excel sheet tab names (e.g. "LWF Act - Form C") into the PDF heading.
  // Titles already extracted from the sheet body are the printed form name.
  const hasFormTitle = titles.some((t) => /^form\s+/i.test(String(t || '')));
  if (
    sheet &&
    sheet !== 'Sheet1' &&
    titles.length === 0 &&
    !looksLikeExcelSheetTabName(sheet) &&
    !looksLikeExcelDraftFileLabel(sheet) &&
    !/^form\s*w$/i.test(sheet) &&
    !(hasFormTitle && /^form\s+/i.test(sheet)) &&
    !titles.some((t) => normalizeMetaKey(t).includes(normalizeMetaKey(sheet).slice(0, 12)))
  ) {
    titles.unshift(sheet);
  }

  const rightFields = [];
  if (
    month !== '' ||
    date !== '' ||
    year !== '' ||
    /month|date|year|wage\s*period/i.test((effectiveMetaLines || []).join(' '))
  ) {
    const hasMonthDateHint = (effectiveMetaLines || []).some((l) =>
      /^(month|date|year|wage\s*period)\b/i.test(String(l || '').trim())
    );
    if (hasMonthDateHint || isFormW || isFormXXVIIRegister) {
      rightFields.push(extractFieldValue('Month', month, { requireColon: false }));
      if (isFormW || year || isFormXXVIIRegister) {
        rightFields.push(extractFieldValue('Year', year || date, { requireColon: false }));
      } else {
        rightFields.push(extractFieldValue('Date', date, { requireColon: false }));
      }
    }
  }

  let finalTitles = titles.filter((t) => !looksLikeExcelDraftFileLabel(t));
  const alreadyForm15LeaveWithWages =
    finalTitles.some((t) => /form[\s._-]*15\b/i.test(String(t || ''))) &&
    finalTitles.some((t) => /register\s+of\s+leave\s+with\s+wages/i.test(String(t || '')));
  // Always drop bare FORM-X when this is Form 15 Part 1 (catalog) or the sheet already
  // shows FORM-15 + Leave with Wages (partial Excel rewrite left FORM-X behind).
  if (
    preferForm15Part1 ||
    alreadyForm15LeaveWithWages ||
    finalTitles.some((t) => isStandaloneFormXTitle(t)) &&
      finalTitles.some((t) => /form[\s._-]*15\b/i.test(String(t || '')))
  ) {
    finalTitles = rewriteForm15Part1PdfTitles(finalTitles);
  }
  // Last pass: never print Form X next to Form 15.
  finalTitles = finalTitles.filter((t) => !isStandaloneFormXTitle(t));

  return {
    titles: finalTitles,
    fields:
      fields.length === 0 && rightFields.length > 0
        ? rightFields.map(() => '')
        : fields,
    rightFields,
    genderBox,
    hasSystemNote,
    isFormXXVI: false,
    isFormW,
    isFormXXVIIRegister,
    /** Form XXVII: full title box, no Month/Year vertical divider. */
    titleBoxFullBorder: isFormXXVIIRegister,
    hideRightBandSplit: isFormXXVIIRegister
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
  const isFormA = titles.some((title) => /^form\s*a\b/i.test(String(title || '').trim()));
  const useRightBand = rightFields.some((t) => String(t || '').trim());
  const hideRightBandSplit = headerModel?.hideRightBandSplit === true;
  const titleBoxFullBorder = headerModel?.titleBoxFullBorder === true;
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

  /**
   * Form XXVII: one outer box around stacked title lines (full L/R borders),
   * horizontal rules only between lines — no internal verticals.
   */
  const paintTitleBoxBlock = (titleSegs) => {
    const prepared = [];
    titleSegs.forEach((title) => {
      expandStatutoryMetaSegments(title).forEach((seg) => {
        const line = String(seg || '').trim();
        if (!line) return;
        const lower = line.toLowerCase();
        const isFormName = /^form\s+/i.test(line);
        const isRegister =
          /^register of\b|^overtime\s+muster\s+roll\b|^muster\s+roll\b|^list of\b|^wage\s+slip\b|^letter\s+of\b|^notice\s+of\b|^combined\b/i.test(
            line
          );
        const isRule = /see\s+(?:sub-)?rule|prescribed\s+under/i.test(lower);
        const isActBanner = /^the\s+.+\b(act|rules)\b/i.test(line);
        const bold = isFormName || isRegister || prepared.length === 0;
        const size = isFormName ? 11 : isRegister ? 10 : isRule || isActBanner ? 8 : 9;
        const minH = isFormName || isRegister ? 20 : 16;
        doc.setFont('helvetica', bold ? 'bold' : 'normal');
        doc.setFontSize(size);
        const wrapped = doc.splitTextToSize(line, usableWidth - padX * 2);
        const h = Math.max(minH, wrapped.length * (size + 2) + 8);
        prepared.push({ line, bold, size, wrapped, h });
      });
    });
    if (!prepared.length) return;

    const totalH = prepared.reduce((a, p) => a + p.h, 0);
    ensureSpace(totalH);
    const boxTop = y;
    strokeRect(x0, boxTop, usableWidth, totalH);
    doc.setDrawColor(0, 0, 0);
    doc.setLineWidth(0.7);
    // Reinforce outer verticals so they meet the table edge (finish at full width).
    doc.line(x0, boxTop, x0, boxTop + totalH);
    doc.line(x1, boxTop, x1, boxTop + totalH);

    prepared.forEach((p, idx) => {
      if (idx > 0) {
        doc.line(x0, y, x1, y);
      }
      doc.setFont('helvetica', p.bold ? 'bold' : 'normal');
      doc.setFontSize(p.size);
      doc.setTextColor(0, 0, 0);
      const textY = y + (p.h - p.wrapped.length * (p.size + 2)) / 2 + p.size;
      doc.text(p.wrapped, (x0 + x1) / 2, textY, { align: 'center' });
      y += p.h;
    });
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
    // Form XXVII: keep Month/Year on the right without an internal vertical divider.
    if (useRightBand && !hideRightBandSplit) {
      doc.line(splitX, y, splitX, y + h);
    }
    doc.setTextColor(0, 0, 0);
    doc.text(leftWrapped, x0 + padX, y + 10);
    if (rightWrapped.length) {
      if (hideRightBandSplit) {
        doc.text(rightWrapped, x1 - padX, y + 10, { align: 'right' });
      } else {
        doc.text(rightWrapped, splitX + padX, y + 10);
      }
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

  // One Excel wrapText title cell → one centered band per line (Form I / Form 10 style).
  if (titleBoxFullBorder && titles.length) {
    paintTitleBoxBlock(titles);
  } else {
    let titlePaintIdx = 0;
    titles.forEach((title) => {
      expandStatutoryMetaSegments(title).forEach((seg) => {
        const lower = String(seg).toLowerCase();
        const isFormName = /^form\s+/i.test(seg);
        const isRegister =
          /^register of\b|^overtime\s+muster\s+roll\b|^muster\s+roll\b|^list of\b|^wage\s+slip\b|^letter\s+of\b|^notice\s+of\b|^combined\b/i.test(
            seg
          );
        const isRule = /see\s+(?:sub-)?rule|prescribed\s+under/i.test(lower);
        const isActBanner = /^the\s+.+\b(act|rules)\b/i.test(seg);
        paintFullBand(seg, {
          bold: isFormName || isRegister || titlePaintIdx === 0,
          size: isFormName ? 11 : isRegister ? 10 : isRule || isActBanner ? 8 : 9,
          align: 'center',
          minH: isFormName || isRegister ? 20 : 16
        });
        titlePaintIdx += 1;
      });
    });
  }

  fields.forEach((field, idx) => {
    const right = rightFields[idx] || '';
    if (useRightBand && idx < Math.max(rightFields.length, 2) && (right || idx < 2)) {
      paintSplitBand(field, right, { minH: 16 });
    } else {
      const centerFormAPartA = isFormA && /\[?\s*part\s*[-\s]?a\s*:/i.test(String(field || ''));
      paintFullBand(field, {
        bold: false,
        size: 8,
        align: centerFormAPartA ? 'center' : 'left',
        minH: 16
      });
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
const drawMatrixSheet = (doc, matrix, startY, pdfOpts = {}) => {
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
  const headerModel = buildStatutoryPdfHeaderModel(metaLines, rows, tableStart, matrix.name, {
    preferredTitle: pdfOpts.preferredTitle || '',
    fileName: pdfOpts.fileName || ''
  });

  let headerBandEnd = tableStart;
  for (let r = tableStart; r < Math.min(rows.length, tableStart + 8); r += 1) {
    if (isLikelyHeaderBandRow(rows[r], r)) headerBandEnd = r;
  }
  // Form XIV Employment Card: every workman row is label|value data — no header band.
  const isFormXIVEmploymentCardSheet = looksLikeFormXIVEmploymentCardPdfContext(
    metaLines,
    rows,
    matrix.name || ''
  );
  if (isFormXIVEmploymentCardSheet) {
    headerBandEnd = tableStart - 1;
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
  // Form XXVII Register of Wages — group → mid → leaf → column-number rows.
  if (headerModel.isFormXXVIIRegister) {
    for (let r = tableStart; r < Math.min(rows.length, tableStart + 6); r += 1) {
      const blob = (rows[r] || []).join(' ').toLowerCase();
      if (
        /wages?\s+earned|deductions?|basic\s+wage|gross\s+wages?|net\s+wages?|other\s+allowances?|wash\s+allow/.test(
          blob
        )
      ) {
        headerBandEnd = Math.max(headerBandEnd, r);
      }
      const filled = (rows[r] || []).filter((c) => String(c || '').trim());
      if (filled.filter((c) => /^\d{1,2}$/.test(String(c).trim())).length >= 8) {
        headerBandEnd = Math.max(headerBandEnd, r);
      }
    }
  }
  // Leave registers + Form 15 Part 2 wage bands + Form VI festival band.
  for (let r = tableStart; r < Math.min(rows.length, tableStart + 8); r += 1) {
    const blob = (rows[r] || []).join(' ').toLowerCase();
    if (
      /earned\s+leave|medical\s+leave|other\s+leave|maternity\s+benefit|leave\s+at\s+the\s+beginning|leave\s+availed|name\s+of\s+the\s+employee/.test(
        blob
      ) ||
      /deductions|advances|damages\s*\/\s*fine|leave\s+wages|provident\s+fund|gross\s+wages|net\s+wages|advance\s+paid/.test(
        blob
      ) ||
      /days,?\s+dates\s+and\s+months|national\s+and\s+festival\s+holidays|pongal|republic\s+day|diwali|christmas/.test(
        blob
      )
    ) {
      headerBandEnd = Math.max(headerBandEnd, r);
    }
  }
  const dayBand = detectDailyHoursBand(rows, tableStart, headerBandEnd, colCount);
  const groupBands = detectStatutoryGroupHeaderBands(rows, tableStart, headerBandEnd, colCount);
  // Resolve leaf headers for ALL forms so widths follow column names (Form 11, Form W, …).
  const leafHeaders = resolveLeafHeaderTexts(rows, tableStart, headerBandEnd, colCount);
  const isFormXIXRajasthanOvertimeSheet = looksLikeFormXIXRajasthanOvertimePdfContext(
    metaLines,
    rows,
    matrix.name || ''
  );
  const isFormASheet = looksLikeFormAPdfContext(metaLines, rows, matrix.name || '');
  const isFormBRajasthanSheet = looksLikeFormBRajasthanPdfContext(metaLines, rows, matrix.name || '');

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
      if (isUnpaidAccumulationsFootnoteRow(rows[r])) continue;
      if (isForm14RajasthanFootnoteRow(rows[r])) continue;
      const len = String(rows[r][c] || '').length;
      maxLen = Math.max(maxLen, len);
      if (r > headerBandEnd) maxDataLen = Math.max(maxDataLen, len);
    }
    if (headerModel.isFormW) {
      weights.push(formWTamilNaduColumnWeight(leafHeaders[c], maxDataLen || maxLen));
      continue;
    }
    if (headerModel.isFormXXVIIRegister) {
      weights.push(formXXVIITamilNaduColumnWeight(leafHeaders[c], maxDataLen || maxLen));
      continue;
    }
    const weight = statutoryHeaderColumnWeight(leafHeaders[c], maxDataLen || maxLen, colCount);
    weights.push(
      (isFormXIXRajasthanOvertimeSheet || isFormASheet || isFormBRajasthanSheet) &&
        isPdfSerialNumberHeader(leafHeaders[c])
        ? 5.5
        : weight
    );
  }
  const weightSum = weights.reduce((a, b) => a + b, 0) || 1;
  const rawWidths = weights.map((w) => (w / weightSum) * usableWidth);
  const colWidths = enforcePdfColumnMinWidths(rawWidths, leafHeaders, usableWidth, {
    compactSerial: isFormXIXRajasthanOvertimeSheet || isFormASheet || isFormBRajasthanSheet
  });
  const colXs = [marginX];
  for (let i = 0; i < colWidths.length; i += 1) colXs.push(colXs[i] + colWidths[i]);

  const fontSize = headerModel.isFormW
    ? colCount > 28
      ? 5.2
      : 5.8
    : headerModel.isFormXXVIIRegister
      ? colCount > 24
        ? 5
        : 5.4
      : colCount > 40
        ? 4.5
        : colCount > 28
          ? 5
          : colCount > 18
            ? 5.5
            : colCount > 14
              ? 6
              : colCount > 10
                ? 7
                : 8;
  let y = startY;
  let pendingSystemNote = false;
  let pendingUnpaidFootnote = '';
  const pendingForm14RjFootnotes = [];

  const looksLikeFormCLwfSheet = (() => {
    const blob = [...metaLines, ...(rows || []).slice(0, Math.min(rows.length, tableStart + 4)).flat(), matrix.name || '']
      .join(' ')
      .toLowerCase();
    return (
      (/form\s*-?\s*c\b/.test(blob) || /lwf\b/.test(blob)) &&
      (/unpaid\s+accumulations|details\s+of\s+fines|quarter\s+ending/.test(blob) ||
        isFormCLwfColHeaderBlob(blob))
    );
  })();
  const looksLikeFormCRajasthanSheet = looksLikeFormCRajasthanPdfContext(
    metaLines,
    rows,
    matrix.name || ''
  );
  const looksLikeForm14RjSheet = looksLikeForm14RajasthanPdfContext(
    metaLines,
    rows,
    matrix.name || ''
  );

  const pushForm14RjFootnote = (raw) => {
    const text = normalizeForm14RajasthanPdfFootnoteText(raw);
    if (!text) return;
    if (pendingForm14RjFootnotes.some((t) => t.toLowerCase() === text.toLowerCase())) return;
    // Keep overtime footnote before the day-entries note.
    if (isForm14RajasthanDayEntriesNoteText(text)) {
      pendingForm14RjFootnotes.push(text);
      return;
    }
    const dayIdx = pendingForm14RjFootnotes.findIndex((t) =>
      isForm14RajasthanDayEntriesNoteText(t)
    );
    if (dayIdx >= 0) pendingForm14RjFootnotes.splice(dayIdx, 0, text);
    else pendingForm14RjFootnotes.push(text);
  };

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

  const groupBandAt = (rowIndex, col) => {
    if (!groupBands.length || rowIndex > headerBandEnd) return null;
    return (
      groupBands.find(
        (b) => b.labelRow === rowIndex && col >= b.start && col <= b.end
      ) || null
    );
  };

  const measureRowHeight = (row, rowIndex) => {
    const mergeDayLabel = isDayBandLabelHeaderRow(row, rowIndex);
    const nilSpan = resolveNilOfTheMonthPdfSpan(row, colCount, rowIndex, headerBandEnd);
    const isHeaderRow = rowIndex <= headerBandEnd;
    let maxLines = 1;

    if (nilSpan) {
      const bandW = colXs[nilSpan.end + 1] - colXs[nilSpan.start] - 3;
      const wrapped = doc.splitTextToSize(nilSpan.text, Math.max(bandW, 6));
      maxLines = Math.max(1, Math.min(wrapped.length, 3));
      return Math.max(fontSize + 5, maxLines * (fontSize + 1.5) + 4);
    }

    for (let c = 0; c < colCount; c += 1) {
      if (mergeDayLabel && dayBand && c > dayBand.start && c <= dayBand.end) continue;
      const groupBand = groupBandAt(rowIndex, c);
      if (groupBand && c > groupBand.start && c <= groupBand.end) continue;
      const width =
        mergeDayLabel && dayBand && c === dayBand.start
          ? colXs[dayBand.end + 1] - colXs[dayBand.start] - 3
          : groupBand && c === groupBand.start
            ? colXs[groupBand.end + 1] - colXs[groupBand.start] - 3
            : Math.max(colWidths[c] - 3, 6);
      const text =
        mergeDayLabel && dayBand && c === dayBand.start
          ? dayBand.label
          : groupBand && c === groupBand.start
            ? groupBand.label
            : String(row[c] ?? '') || ' ';
      // Data rows: never wrap pure numbers (they shrink-to-fit when painted).
      if (!isHeaderRow && isPurePdfNumericText(text)) {
        maxLines = Math.max(maxLines, 1);
        continue;
      }
      const wrapped = doc.splitTextToSize(text, Math.max(width, 6));
      const lineCap =
        groupBand && isFormVIFestivalGroupLabel(groupBand.label)
          ? 10
          : isHeaderRow
            ? headerModel.isFormW
              ? 6
              : 8
            : 8;
      maxLines = Math.max(maxLines, Math.min(wrapped.length, lineCap));
    }
    return Math.max(fontSize + 5, maxLines * (fontSize + 1.5) + 4);
  };

  const paintGridRow = (row, rowIndex, rowH) => {
    const bold = rowIndex <= headerBandEnd || isLikelyHeaderBandRow(row, rowIndex);
    const mergeDayLabel = isDayBandLabelHeaderRow(row, rowIndex);
    const nilSpan = resolveNilOfTheMonthPdfSpan(row, colCount, rowIndex, headerBandEnd);
    const isHeaderRow = rowIndex <= headerBandEnd;
    doc.setDrawColor(0, 0, 0);
    doc.setLineWidth(0.4);
    doc.setFont('helvetica', bold ? 'bold' : 'normal');
    doc.setFontSize(fontSize);
    doc.setTextColor(0, 0, 0);

    if (nilSpan) {
      const bandW = colXs[nilSpan.end + 1] - colXs[nilSpan.start];
      doc.rect(colXs[nilSpan.start], y, bandW, rowH, 'S');
      const lines = doc
        .splitTextToSize(nilSpan.text, Math.max(bandW - 4, 8))
        .slice(0, 3);
      const textH = lines.length * (fontSize + 1);
      doc.text(lines, colXs[nilSpan.start] + bandW / 2, y + (rowH - textH) / 2 + fontSize, {
        align: 'center'
      });
      y += rowH;
      return;
    }

    for (let c = 0; c < colCount; c += 1) {
      if (mergeDayLabel && dayBand && c > dayBand.start && c <= dayBand.end) continue;
      const groupBand = groupBandAt(rowIndex, c);
      if (groupBand && c > groupBand.start && c <= groupBand.end) continue;

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

      if (groupBand && c === groupBand.start) {
        const bandW = colXs[groupBand.end + 1] - colXs[groupBand.start];
        doc.rect(colXs[c], y, bandW, rowH, 'S');
        const groupLineCap = isFormVIFestivalGroupLabel(groupBand.label) ? 10 : 3;
        const lines = doc
          .splitTextToSize(groupBand.label, Math.max(bandW - 4, 8))
          .slice(0, groupLineCap);
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
      // Hide leftover group labels that leaked into leaf header cells
      if (
        groupBands.length &&
        rowIndex <= headerBandEnd &&
        isStatutoryGroupHeaderLabel(raw) &&
        !groupBandAt(rowIndex, c)
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
      } else if (isHeaderRow || isPureNilPdfText(raw)) {
        // Header labels and Nil/NIL values — center like the Excel register model.
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
  } else if (
    matrix.name &&
    matrix.name !== 'Sheet1' &&
    !looksLikeExcelDraftFileLabel(matrix.name) &&
    !looksLikeExcelSheetTabName(matrix.name)
  ) {
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
    if (isUnpaidAccumulationsFootnoteText(line) && !pendingUnpaidFootnote) {
      pendingUnpaidFootnote = String(line).replace(/\s+/g, ' ').trim();
    }
    if (isForm14RajasthanFootnoteText(line)) pushForm14RjFootnote(line);
  });

  for (let r = tableStart; r < rows.length; r += 1) {
    const row = rows[r];
    if (!rowHasContent(row)) continue;

    // Never draw the footer inside Sr. No. — paint after the full column band.
    if (isSystemGeneratedDocumentNoteRow(row)) {
      pendingSystemNote = true;
      continue;
    }

    // Form C legal footnote — paint below the grid, above the system-generated note.
    if (isUnpaidAccumulationsFootnoteRow(row) || isUnpaidAccumulationsFootnoteText(row?.[0])) {
      const fn = extractUnpaidAccumulationsFootnoteText(row) || String(row?.[0] || '').trim();
      if (fn && !pendingUnpaidFootnote) pendingUnpaidFootnote = fn;
      continue;
    }

    // Form 14 RJ footnotes — paint below the grid, above the system-generated note.
    if (isForm14RajasthanFootnoteRow(row) || isForm14RajasthanFootnoteText(row?.[0])) {
      const fn = extractForm14RajasthanFootnoteText(row) || String(row?.[0] || '').trim();
      if (fn) pushForm14RjFootnote(fn);
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
  if (!pendingSystemNote || !pendingUnpaidFootnote || pendingForm14RjFootnotes.length < 2) {
    for (let r = 0; r < rows.length; r += 1) {
      if (!pendingSystemNote && isSystemGeneratedDocumentNoteRow(rows[r])) {
        pendingSystemNote = true;
      }
      if (!pendingUnpaidFootnote) {
        const fn = extractUnpaidAccumulationsFootnoteText(rows[r]);
        if (fn) pendingUnpaidFootnote = fn;
      }
      if (isForm14RajasthanFootnoteRow(rows[r]) || isForm14RajasthanFootnoteText(rows[r]?.[0])) {
        const fn = extractForm14RajasthanFootnoteText(rows[r]) || String(rows[r]?.[0] || '').trim();
        if (fn) pushForm14RjFootnote(fn);
      }
    }
  }

  // Form C always shows the unpaid-accumulations definition before the system note.
  if (looksLikeFormCLwfSheet && !pendingUnpaidFootnote) {
    pendingUnpaidFootnote = FORM_C_UNPAID_ACCUMULATIONS_FOOTNOTE;
  }
  if (looksLikeFormCLwfSheet) {
    pendingSystemNote = true;
  }
  if (looksLikeFormCRajasthanSheet) {
    pendingUnpaidFootnote = FORM_C_RJ_FOOTNOTE;
    pendingSystemNote = true;
  }

  // Form 14 RJ always keeps both legal notes under the register, above the system note.
  if (looksLikeForm14RjSheet) {
    if (!pendingForm14RjFootnotes.some((t) => isForm14RajasthanOvertimeFootnoteText(t))) {
      pushForm14RjFootnote(FORM_14_RJ_FOOTER_NOTE);
    }
    if (!pendingForm14RjFootnotes.some((t) => isForm14RajasthanDayEntriesNoteText(t))) {
      pushForm14RjFootnote(FORM_14_RJ_DAY_ENTRIES_NOTE);
    }
    pendingSystemNote = true;
  }

  if (pendingUnpaidFootnote) {
    y += 8;
    paintMetaLine(pendingUnpaidFootnote, { bold: false, size: 8, align: 'left' });
  }
  if (pendingForm14RjFootnotes.length) {
    y += pendingUnpaidFootnote ? 4 : 8;
    pendingForm14RjFootnotes.forEach((fn, idx) => {
      if (idx > 0) y += 2;
      paintMetaLine(fn, { bold: false, size: 8, align: 'left' });
    });
  }
  if (pendingSystemNote) {
    y += pendingUnpaidFootnote || pendingForm14RjFootnotes.length ? 4 : 8;
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

  const matricesForPdf = filterStatutoryPdfMatrices(allMatrices);
  if (!matricesForPdf.length) {
    throw new Error('Draft Excel has no readable rows to put in the PDF.');
  }

  const maxCols = Math.max(...matricesForPdf.map((m) => m.colCount));
  const anyFormW = matricesForPdf.some((m) =>
    looksLikeFormWPdfContext(m.metaLines, m.rows, m.tableStartRow || 0)
  );
  const anyAccidentBook = matricesForPdf.some((m) => {
    const blob = [...(m.metaLines || []), ...(m.rows || []).slice(0, 8).flat(), m.name || '']
      .join(' ')
      .toLowerCase();
    return /accident\s+book|form\s*(?:no\.?\s*)?11\b/.test(blob);
  });
  const anyForm25 = matricesForPdf.some((m) =>
    looksLikeForm25TamilNaduPdfContext(m.metaLines, m.rows, m.name)
  );
  const anyForm14Rj = matricesForPdf.some((m) =>
    looksLikeForm14RajasthanPdfContext(m.metaLines, m.rows, m.name)
  );
  const anyFormXIXWageSlip = matricesForPdf.some((m) =>
    looksLikeFormXIXWageSlipPdfContext(m.metaLines, m.rows, m.name)
  );
  const anyFormXIVEmploymentCard = matricesForPdf.some((m) =>
    looksLikeFormXIVEmploymentCardPdfContext(m.metaLines, m.rows, m.name)
  );
  // Form 14 RJ has few columns but very long header labels — landscape keeps them aligned.
  // Form XIX Wage Slip / Form XIV Employment Card stay portrait A4 (compact card layout).
  const forcePortraitCard = anyFormXIXWageSlip || anyFormXIVEmploymentCard;
  const wide =
    (!forcePortraitCard && maxCols > 8) ||
    anyFormW ||
    anyAccidentBook ||
    anyForm25 ||
    anyForm14Rj;
  // Form W (~30 wage/deduction cols) needs A2 landscape so amounts stay on one line.
  // Form 11 Accident Book (~18 cols with long headers) and other wide registers need A3.
  const veryWide =
    anyFormW || anyAccidentBook || anyForm25 || (!forcePortraitCard && maxCols > 14);
  const doc = new jsPDF({
    unit: 'pt',
    format: forcePortraitCard ? 'a4' : anyFormW ? 'a2' : veryWide ? 'a3' : 'a4',
    orientation: forcePortraitCard ? 'portrait' : wide ? 'landscape' : 'portrait'
  });

  // Bordered header draws form name for every sheet — skip a floating duplicate title.
  const firstMatrix = matricesForPdf[0];
  const pdfHeaderOpts = { preferredTitle: title, fileName };
  const firstHeaderModel = buildStatutoryPdfHeaderModel(
    firstMatrix?.metaLines || [],
    firstMatrix?.rows || [],
    firstMatrix?.tableStartRow || 0,
    firstMatrix?.name || '',
    pdfHeaderOpts
  );
  const hasBorderedHeader =
    (firstHeaderModel.titles || []).length > 0 || (firstHeaderModel.fields || []).length > 0;
  // Use a real form title only — never the Excel/ZIP draft file name as a PDF heading.
  const headingCandidate = String(title || '')
    .replace(EXCEL_EXT_RE, '')
    .replace(/\.zip$/i, '')
    .trim()
    .slice(0, 120);
  const heading =
    headingCandidate && !looksLikeExcelDraftFileLabel(headingCandidate) ? headingCandidate : '';
  if (!hasBorderedHeader && heading) {
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(11);
    doc.setTextColor(0, 0, 0);
    doc.text(heading, doc.internal.pageSize.getWidth() / 2, 20, { align: 'center' });
  }

  let y = hasBorderedHeader || !heading ? 20 : 32;
  for (let i = 0; i < matricesForPdf.length; i += 1) {
    if (i > 0) {
      doc.addPage();
      y = 28;
    }
    y = drawMatrixSheet(doc, matricesForPdf[i], y, pdfHeaderOpts);
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
  isFormCLwfColHeaderBlob,
  isForm14RajasthanColHeaderBlob,
  isFormBRajasthanWageRateGroupLabel,
  looksLikeFormBRajasthanPdfContext,
  looksLikeForm14RajasthanPdfContext,
  looksLikeForm11RajasthanPdfContext,
  looksLikeFormXIXRajasthanOvertimePdfContext,
  looksLikeFormXIXWageSlipPdfContext,
  looksLikeFormXRajasthanEmploymentCardPdfContext,
  buildFormXRajasthanEmploymentHeaderModel,
  looksLikeFormXIRajasthanServiceCertificatePdfContext,
  buildFormXIRajasthanServiceHeaderModel,
  looksLikeFormXVRajasthanWageSlipPdfContext,
  buildFormXVRajasthanWageSlipHeaderModel,
  isFormXIXWageSlipColHeaderBlob,
  normalizeFormXIXWageSlipPdfMatrix,
  looksLikeFormXIVEmploymentCardPdfContext,
  normalizeFormXIVEmploymentCardPdfMatrix,
  isFormXIVEmploymentCardWorkmanLabelText,
  scrubForm11RajasthanTotalHoursPdfColumn,
  trimForm14RajasthanLeadingBlankPdfColumns,
  trimFormXIXRajasthanLeadingBlankPdfColumns,
  trimFormXRajasthanEmploymentLeadingBlankPdfColumns,
  isForm25TamilNaduColHeaderBlob,
  looksLikeForm25TamilNaduPdfContext,
  findRemarksColumnIndex,
  trimTrailingBlankPdfColumns,
  isForm25TamilNaduEmployeePdfRow,
  trimForm25TamilNaduPdfTrailingEmployeeRows,
  looksLikeAuxiliaryOrPivotPdfSheet,
  looksLikeForm15TamilNaduPdfContext,
  looksLikeForm15Part1PreferredContext,
  rewriteForm15Part1PdfTitles,
  isStandaloneFormXTitle,
  detectLeaveCategoryBands,
  detectStatutoryGroupHeaderBands,
  isLeaveCategoryGroupLabel,
  isWageDeductionGroupLabel,
  isFormVIFestivalGroupLabel,
  isStatutoryGroupHeaderLabel,
  filterStatutoryPdfMatrices,
  stripLeakedPivotRowsFromPdfMatrix,
  isFormWAdminBandBlob,
  isFormWAdminValueRow,
  looksLikeFormWPdfContext,
  formWTamilNaduColumnWeight,
  formXXVIITamilNaduColumnWeight,
  looksLikeFormXXVIITamilNaduRegisterPdfContext,
  extractMonthFromWagePeriodLine,
  isPurePdfNumericText,
  isPureNilPdfText,
  isNilOfTheMonthPdfText,
  resolveNilOfTheMonthPdfSpan,
  isUnpaidAccumulationsFootnoteText,
  isUnpaidAccumulationsFootnoteRow,
  FORM_C_UNPAID_ACCUMULATIONS_FOOTNOTE,
  FORM_C_RJ_FOOTNOTE,
  looksLikeFormCRajasthanPdfContext,
  isForm14RajasthanOvertimeFootnoteText,
  isForm14RajasthanDayEntriesNoteText,
  isForm14RajasthanFootnoteText,
  isForm14RajasthanFootnoteRow,
  extractForm14RajasthanFootnoteText,
  FORM_14_RJ_FOOTER_NOTE,
  FORM_14_RJ_DAY_ENTRIES_NOTE,
  extractFormWGenderBox,
  sheetToDenseMatrix,
  looksLikeExcelDraftFileLabel,
  looksLikeExcelSheetTabName,
  buildStatutoryPdfHeaderModel,
  isStatutoryTitleMetaLine,
  expandStatutoryMetaSegments,
  normalizeStatutoryMultilineText,
  isPdfSerialNumberHeader,
  statutoryHeaderColumnWeight,
  enforcePdfColumnMinWidths,
  resolveLeafHeaderTexts
};
