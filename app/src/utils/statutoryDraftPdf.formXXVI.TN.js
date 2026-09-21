import * as XLSX from 'xlsx';
import { inferSmartbrowzPdfOptions, requestSmartbrowzPdf } from './smartbrowzPdf';
import { resolveToFullMonthName } from './statutoryMonthResolve';

export const FORM_XXVI_TN_PDF_TITLE = 'FORM XXVI';
export const FORM_XXVI_TN_PDF_REFERENCE =
  'See Rule 75 of the Tamil Nadu Contract Labour (Regulation and Abolition) Rules, 1975';
export const FORM_XXVI_TN_PDF_SUBTITLE = 'Register of Employment of Contractual Labour';

const escapeHtml = (value) =>
  String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');

const cellText = (value) => String(value ?? '').replace(/\s+/g, ' ').trim();

const stripFieldLabelPrefix = (text, labelRe) =>
  String(text || '')
    .replace(labelRe, '')
    .replace(/^[\s.:\-–—]+/, '')
    .trim();

export function looksLikeFormXXVITamilNaduPdfContext({
  fileName = '',
  title = '',
  sheetName = '',
  rows = []
} = {}) {
  const blob = [fileName, title, sheetName, ...(Array.isArray(rows) ? rows.slice(0, 16).flat() : [])]
    .join(' ')
    .toLowerCase();
  if (!blob.trim()) return false;
  if (/form[\s._-]*xxvii(?![a-z])/.test(blob)) return false;
  if (/\bform\s*26[\s._-]*a\b|\bform_26[\s._-]*a\b/.test(blob)) return false;
  if (/letter\s+of\s+appointment/.test(blob)) return false;
  if (/register\s+of\s+accident/.test(blob) && !/register\s+of\s+employment/.test(blob)) {
    return false;
  }
  const xxviHint =
    /form[\s._-]*xxvi(?![a-z])/.test(blob) ||
    /form[\s._-]*26(?!\d)/.test(blob) ||
    /register of employment of contractual?\s*labour/.test(blob);
  if (!xxviHint) return false;
  return (
    /tamil/.test(blob) ||
    /see\s+rule\s*75/.test(blob) ||
    /register of employment of contractual?\s*labour/.test(blob) ||
    /daily\s+hours\s+of\s+work/.test(blob)
  );
}

const isFormXXVIPeriodLabel = (text) => /^(?:month|date|year)\s*:?\s*$/i.test(String(text || '').trim());

const isFormXXVITableHeaderLike = (text) =>
  /serial|s\.?\s*no|name of the workman|name of the worker|daily hours|rate of wages|number of days|signature|thumb|termination|designation|permanent|local address/i.test(
    String(text || '')
  );

function pickValueNearLabel(rows, r, c, skipRe, tableStart) {
  const belowOk = tableStart == null || r + 1 < tableStart;
  const candidates = [
    cellText(rows[r]?.[c + 1]),
    belowOk ? cellText(rows[r + 1]?.[c]) : '',
    belowOk ? cellText(rows[r + 1]?.[c + 1]) : '',
    cellText(rows[r]?.[c + 2])
  ];
  for (let i = 0; i < candidates.length; i += 1) {
    const cand = candidates[i];
    if (!cand || isFormXXVIPeriodLabel(cand) || isFormXXVITableHeaderLike(cand)) continue;
    if (skipRe && skipRe.test(cand)) continue;
    return cand;
  }
  return '';
}

export function extractFormXXVITamilNaduPdfHeader(rows = [], tableStart = 0, fallbacks = {}) {
  const principalRe = /name\s+and\s+address\s+of\s+(?:the\s+)?principal\s+employer/i;
  const contractorRe = /name\s+and\s+address\s+of\s+(?:the\s+)?contractor/i;
  const worksiteRe =
    /(?:nature\s+and\s+location\s+of\s+work|name\s+and\s+location\s+of\s+(?:the\s+)?work\s*site?)/i;

  let principal = '';
  let contractor = '';
  let worksite = '';
  let month = '';
  let date = '';

  const scanLimit = Math.max(tableStart || 0, Math.min(rows.length, 24));
  for (let r = 0; r < scanLimit; r += 1) {
    const row = rows[r] || [];
    const colLimit = Math.max(row.length, 80);
    for (let c = 0; c < colLimit; c += 1) {
      const cell = cellText(row[c]);
      if (!cell) continue;
      const lower = cell.toLowerCase();
      const next = cellText(row[c + 1]);
      if (!principal && principalRe.test(cell)) {
        principal = stripFieldLabelPrefix(cell, principalRe) || next;
      }
      if (!contractor && contractorRe.test(cell) && !/workman/.test(lower)) {
        contractor = stripFieldLabelPrefix(cell, contractorRe) || next;
      }
      if (!worksite && worksiteRe.test(cell)) {
        worksite = stripFieldLabelPrefix(cell, worksiteRe) || next;
      }
      if (!month && /^month\s*:?\s*$/i.test(cell)) {
        month = pickValueNearLabel(rows, r, c, /date|year|name and/i, tableStart);
      }
      if (!month) {
        const monthMatch = cell.match(/^month\s*:?\s*(.+)$/i);
        if (monthMatch?.[1] && !/name and|date/i.test(monthMatch[1])) month = monthMatch[1].trim();
      }
      if (!date && /^(?:date|year)\s*:?\s*$/i.test(cell)) {
        date = pickValueNearLabel(rows, r, c, /month|name and|nature/i, tableStart);
      }
      if (!date) {
        const dateMatch = cell.match(/^(?:date|year)\s*:?\s*(.+)$/i);
        if (dateMatch?.[1] && !/name and|nature/i.test(dateMatch[1])) date = dateMatch[1].trim();
      }
    }
  }

  if (!month) {
    month = resolveToFullMonthName(fallbacks.month) || cellText(fallbacks.month);
  }
  if (!date) {
    const fbDate = cellText(fallbacks.date);
    date = (fbDate.match(/(?:19|20)\d{2}/) || [])[0] || fbDate;
  }

  return {
    title: FORM_XXVI_TN_PDF_TITLE,
    reference: FORM_XXVI_TN_PDF_REFERENCE,
    subtitle: FORM_XXVI_TN_PDF_SUBTITLE,
    principal,
    contractor,
    worksite,
    month,
    date
  };
}

/** Column widths follow heading box labels — never body-cell length. */
export function formXXVITamilNaduPdfColumnWeight(headerText, colIndex = 0, dayBand = null) {
  if (dayBand && colIndex >= dayBand.start && colIndex <= dayBand.end) return 1.15;
  const h = String(headerText || '')
    .replace(/\s+/g, ' ')
    .trim();
  const lower = h.toLowerCase();
  if (/serial|s\.?\s*no/.test(lower)) return 2.2;
  if (/age\s*(and|&)?\s*sex|^sex$|^age$/.test(lower)) return 2.4;
  if (/rate\s+of\s+wages?/.test(lower)) return 3;
  if (/date of (entry|termination)|entry into service/.test(lower)) return 3.2;
  if (/permanent|home address/.test(lower)) return 5.2;
  if (/local address/.test(lower)) return 5;
  if (/name of the workm|name of the worker/.test(lower)) return 5;
  if (/father|husband/.test(lower)) return 4.4;
  if (/designation|nature of work/.test(lower)) return 4.2;
  if (/number of days/.test(lower)) return 2.8;
  if (/signature|thumb|contractor|representative/.test(lower)) return 3.4;
  const len = h.length;
  return Math.min(6, Math.max(2.4, len * 0.16 || 2.6));
}

export function findFormXXVITamilNaduPdfTableStart(rows = []) {
  const list = Array.isArray(rows) ? rows : [];
  for (let r = 0; r < Math.min(list.length, 40); r += 1) {
    const blob = (list[r] || []).map((c) => String(c || '').toLowerCase()).join(' ');
    if (
      /serial\s+number|s\.?\s*no/.test(blob) &&
      /name of the workman|name of the worker/.test(blob)
    ) {
      return r;
    }
    if (/daily\s+hours\s+of\s+work/.test(blob) && /rate\s+of\s+wages/.test(blob)) {
      return r;
    }
  }
  return 0;
}

function detectDayBand(rows, tableStart, colCount) {
  const headerEnd = Math.min(tableStart + 6, rows.length - 1);
  let labelCol = -1;
  for (let r = tableStart; r <= headerEnd; r += 1) {
    for (let c = 0; c < colCount; c += 1) {
      if (/daily\s+hours\s+of\s+work/i.test(String(rows[r]?.[c] || ''))) {
        labelCol = c;
        break;
      }
    }
    if (labelCol >= 0) break;
  }
  if (labelCol < 0) {
    let rateCol = -1;
    for (let r = tableStart; r <= headerEnd; r += 1) {
      for (let c = 0; c < colCount; c += 1) {
        if (/rate\s+of\s+wages?/i.test(String(rows[r]?.[c] || ''))) rateCol = c;
      }
    }
    if (rateCol >= 0 && rateCol + 20 < colCount) labelCol = rateCol + 1;
  }
  if (labelCol < 0) return null;

  let trailingStart = colCount;
  for (let r = tableStart; r <= headerEnd; r += 1) {
    for (let c = labelCol + 1; c < colCount; c += 1) {
      const t = String(rows[r]?.[c] || '');
      if (/number\s+of\s+days|signature|thumb|termination|contractor/i.test(t)) {
        trailingStart = Math.min(trailingStart, c);
      }
    }
  }

  let dayEnd = -1;
  let dayRow = -1;
  for (let r = tableStart; r <= Math.min(headerEnd + 2, rows.length - 1); r += 1) {
    let run = 0;
    let end = -1;
    for (let c = labelCol; c < Math.min(trailingStart, colCount); c += 1) {
      const n = Number(String(rows[r]?.[c] || '').trim());
      if (Number.isInteger(n) && n >= 1 && n <= 31) {
        run += 1;
        end = c;
      } else if (run > 0 && !String(rows[r]?.[c] || '').trim()) {
        continue;
      } else if (run > 0) {
        break;
      }
    }
    if (run >= 8) {
      dayEnd = end;
      dayRow = r;
    }
  }
  if (dayEnd < labelCol) {
    dayEnd = Math.min(labelCol + 30, trailingStart - 1, colCount - 1);
  }
  if (dayEnd < labelCol) return null;
  return { start: labelCol, end: dayEnd, dayRow, trailingStart };
}

function findHeaderBandEnd(rows, tableStart, dayBand) {
  let end = tableStart;
  const limit = Math.min(rows.length - 1, tableStart + 6);
  if (dayBand?.dayRow >= tableStart) end = Math.max(end, dayBand.dayRow);
  for (let r = tableStart; r <= limit; r += 1) {
    const filled = (rows[r] || []).filter((c) => String(c || '').trim());
    const dayHits = filled.filter((c) => /^\d{1,2}$/.test(String(c).trim()) && Number(c) <= 31).length;
    const blob = filled.join(' ').toLowerCase();
    if (
      dayHits >= 8 ||
      /serial\s+number|name of the workman|daily hours|rate of wages|number of days/.test(blob)
    ) {
      end = Math.max(end, r);
    }
  }
  return end;
}

function worksheetToRows(ws) {
  const aoa = XLSX.utils.sheet_to_json(ws, {
    header: 1,
    defval: '',
    raw: false,
    blankrows: true
  });
  return (aoa || []).map((row) => (Array.isArray(row) ? row.map((c) => cellText(c)) : []));
}

export function workbookArrayBufferToFormXXVITamilNaduRows(arrayBuffer) {
  const wb = XLSX.read(arrayBuffer, { type: 'array', cellDates: true, cellText: true });
  const names = wb.SheetNames || [];
  const preferred =
    names.find((n) => /form[\s._-]*xxvi|form[\s._-]*26(?!\d)/i.test(String(n))) || names[0];
  const ws = preferred ? wb.Sheets[preferred] : null;
  if (!ws) return { rows: [], sheetName: '', colCount: 0 };
  const rows = worksheetToRows(ws);
  let colCount = 0;
  rows.forEach((row) => {
    colCount = Math.max(colCount, row.length);
  });
  try {
    if (ws['!ref']) {
      const range = XLSX.utils.decode_range(ws['!ref']);
      colCount = Math.max(colCount, range.e.c - range.s.c + 1);
    }
  } catch (_) {
    /* keep computed width */
  }
  const padded = rows.map((row) => {
    const next = [...row];
    while (next.length < colCount) next.push('');
    return next;
  });
  return { rows: padded, sheetName: preferred || '', colCount };
}

function renderMetaValue(value) {
  return escapeHtml(value || '');
}

function buildOfficialHeaderHtml(header) {
  return `
  <header class="xxvi-banner">
    <div class="xxvi-title">${escapeHtml(header.title)}</div>
    <div class="xxvi-ref">${escapeHtml(header.reference)}</div>
    <div class="xxvi-subtitle">${escapeHtml(header.subtitle)}</div>
  </header>
  <table class="xxvi-meta">
    <tr>
      <td>
        <div class="meta-label">Name and address of Principal Employer</div>
        <div class="meta-value">${renderMetaValue(header.principal)}</div>
      </td>
      <td class="meta-right">
        <div><span class="meta-label">Month</span> : ${renderMetaValue(header.month)}</div>
      </td>
    </tr>
    <tr>
      <td>
        <div class="meta-label">Name and Address of Contractor.</div>
        <div class="meta-value">${renderMetaValue(header.contractor)}</div>
      </td>
      <td class="meta-right">
        <div><span class="meta-label">Date</span> : ${renderMetaValue(header.date)}</div>
      </td>
    </tr>
    <tr>
      <td>
        <div class="meta-label">Nature and location of work.</div>
        <div class="meta-value">${renderMetaValue(header.worksite)}</div>
      </td>
      <td class="meta-right"></td>
    </tr>
  </table>`;
}

function leafHeaderText(theadRows, col) {
  for (let i = theadRows.length - 1; i >= 0; i -= 1) {
    const t = cellText(theadRows[i]?.[col]);
    if (t && !/^\d{1,2}$/.test(t)) return t;
  }
  for (let i = 0; i < theadRows.length; i += 1) {
    const t = cellText(theadRows[i]?.[col]);
    if (t) return t;
  }
  return '';
}

function buildGridHtml(rows, tableStart, colCount) {
  const dayBand = detectDayBand(rows, tableStart, colCount);
  const headerEnd = findHeaderBandEnd(rows, tableStart, dayBand);
  const theadRows = [];
  for (let r = tableStart; r <= headerEnd && r < rows.length; r += 1) {
    theadRows.push(rows[r] || []);
  }
  const bodyRows = rows.slice(headerEnd + 1).filter((row) =>
    (row || []).some((c) => String(c || '').trim())
  );

  const weights = [];
  for (let c = 0; c < colCount; c += 1) {
    weights.push(formXXVITamilNaduPdfColumnWeight(leafHeaderText(theadRows, c), c, dayBand));
  }
  const weightSum = weights.reduce((a, b) => a + b, 0) || 1;
  const colgroup = `<colgroup>${weights
    .map((w) => `<col style="width:${((w / weightSum) * 100).toFixed(3)}%" />`)
    .join('')}</colgroup>`;

  const renderRow = (row, tag) => {
    const cells = [];
    for (let c = 0; c < colCount; c += 1) {
      const text = cellText(row?.[c]);
      const inDay = dayBand && c >= dayBand.start && c <= dayBand.end;
      cells.push(
        `<${tag} class="${inDay ? 'day' : ''}">${escapeHtml(text)}</${tag}>`
      );
    }
    return `<tr>${cells.join('')}</tr>`;
  };

  const thead = theadRows.map((row) => renderRow(row, 'th')).join('');
  const tbody = (bodyRows.length ? bodyRows : [[]]).map((row) => renderRow(row, 'td')).join('');
  return `<table class="xxvi-grid">${colgroup}<thead>${thead}</thead><tbody>${tbody}</tbody></table>`;
}

export function buildFormXXVITamilNaduPdfHtml({
  rows = [],
  fileName = '',
  title = '',
  sheetName = '',
  monthLabel = '',
  dateLabel = ''
} = {}) {
  const tableStart = findFormXXVITamilNaduPdfTableStart(rows);
  const header = extractFormXXVITamilNaduPdfHeader(rows, tableStart, {
    month: monthLabel,
    date: dateLabel
  });
  const colCount = rows.reduce((max, row) => Math.max(max, (row || []).length), 0);
  return `<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8" />
  <title>${escapeHtml(header.title)}</title>
  <style>
    * { box-sizing: border-box; }
    body {
      font-family: "Times New Roman", Times, serif;
      color: #000;
      margin: 0;
      padding: 8px 10px;
    }
    .xxvi-banner { text-align: center; margin: 0 0 8px; }
    .xxvi-title { font-size: 18px; font-weight: 700; letter-spacing: 0.04em; }
    .xxvi-ref { font-size: 11px; margin-top: 2px; }
    .xxvi-subtitle { font-size: 13px; font-weight: 700; margin-top: 4px; }
    table.xxvi-meta {
      width: 100%;
      border-collapse: collapse;
      margin-bottom: 8px;
      font-size: 10px;
      table-layout: fixed;
    }
    table.xxvi-meta td {
      border: 1px solid #000;
      padding: 6px 8px;
      vertical-align: top;
      width: 74%;
    }
    table.xxvi-meta td.meta-right { width: 26%; }
    .meta-label { font-weight: 700; }
    .meta-value { min-height: 14px; margin-top: 2px; }
    table.xxvi-grid {
      border-collapse: collapse;
      width: 100%;
      font-size: 7px;
      table-layout: fixed;
    }
    table.xxvi-grid th, table.xxvi-grid td {
      border: 1px solid #000;
      padding: 2px 3px;
      vertical-align: top;
      overflow-wrap: anywhere;
      word-break: break-word;
    }
    table.xxvi-grid th {
      font-weight: 700;
      text-align: center;
      background: #f7f7f7;
    }
    table.xxvi-grid .day {
      text-align: center;
      white-space: nowrap;
    }
  </style>
</head>
<body data-form="xxvi-tn" data-file="${escapeHtml(fileName || title || sheetName)}">
  ${buildOfficialHeaderHtml(header)}
  ${buildGridHtml(rows, tableStart, colCount)}
</body>
</html>`;
}

export async function convertFormXXVITamilNaduExcelToSmartbrowzPdf({
  excelFiles = [],
  title = '',
  fileName = 'FORM-XXVI.pdf',
  monthLabel = ''
} = {}) {
  if (typeof process !== 'undefined' && process.env && process.env.NODE_ENV === 'test') {
    return null;
  }
  const files = Array.isArray(excelFiles) ? excelFiles : [];
  for (const file of files) {
    const label = file.label || fileName || title || '';
    const { rows, sheetName, colCount } = workbookArrayBufferToFormXXVITamilNaduRows(file.arrayBuffer);
    if (
      !looksLikeFormXXVITamilNaduPdfContext({
        fileName: label,
        title,
        sheetName,
        rows
      })
    ) {
      continue;
    }
    const html = buildFormXXVITamilNaduPdfHtml({
      rows,
      fileName: label,
      title,
      sheetName,
      monthLabel
    });
    const inferred = inferSmartbrowzPdfOptions({
      fileName: label || fileName,
      title: FORM_XXVI_TN_PDF_TITLE,
      maxCols: colCount
    });
    return requestSmartbrowzPdf({
      html,
      fileName: /\.pdf$/i.test(fileName) ? fileName : 'FORM-XXVI.pdf',
      pdf_options: inferred.pdf_options,
      page_options: inferred.page_options
    });
  }
  return null;
}
