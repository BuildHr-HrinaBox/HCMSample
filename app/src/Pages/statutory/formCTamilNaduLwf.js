/** Tamil Nadu LWF Form C — Register of Fines and Unpaid Accumulations (rule 29). */

const FORM_C_SUBTITLE_BASE = 'Register of Fines and Unpaid Accumulations for the year';
export const FORM_C_TN_ESTABLISHMENT_LABEL = 'Name of the Establishment';

/** "Name of the Establishment : Theni Site, …" — keep the label for Excel + PDF. */
export function formatFormCTamilNaduLwfEstablishmentLine(value) {
  const val = String(value || '')
    .replace(/\s+/g, ' ')
    .trim();
  if (!val) return `${FORM_C_TN_ESTABLISHMENT_LABEL} :`;
  if (/^name\s+of\s+the\s+establishment\s*:/i.test(val)) {
    const rest = val.replace(/^name\s+of\s+the\s+establishment\s*:?\s*/i, '').trim();
    return rest
      ? `${FORM_C_TN_ESTABLISHMENT_LABEL} : ${rest}`
      : `${FORM_C_TN_ESTABLISHMENT_LABEL} :`;
  }
  return `${FORM_C_TN_ESTABLISHMENT_LABEL} : ${val}`;
}

/** Signature / employer footer lines — must not appear on Form C PDF. */
export function isFormCTamilNaduLwfPdfSignatoryText(text) {
  const flat = String(text || '')
    .replace(/_x000d_/gi, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  if (!flat) return false;
  if (/^for\s*\(/i.test(flat)) return true;
  if (/authori[sz]ed\s+signatory/i.test(flat)) return true;
  if (/signature\s+of\s+employer/i.test(flat)) return true;
  if (/manager\s*\/\s*authori[sz]ed\s+person/i.test(flat)) return true;
  return false;
}

export function isFormCTamilNaduLwfPdfSignatoryRow(row) {
  if (!Array.isArray(row)) return isFormCTamilNaduLwfPdfSignatoryText(row);
  const filled = row
    .map((c) => String(c || '').replace(/\s+/g, ' ').trim())
    .filter(Boolean);
  if (!filled.length) return false;
  return filled.every((t) => isFormCTamilNaduLwfPdfSignatoryText(t));
}

/**
 * Keep Form C titles + force "Name of the Establishment :" field.
 * Drop Authorised Signatory / Signature of Employer footer lines from PDF header.
 */
export function rewriteFormCTamilNaduLwfPdfHeader(titles, fields) {
  const titleList = Array.isArray(titles) ? titles : [];
  const fieldList = Array.isArray(fields) ? fields : [];
  const nextTitles = [];
  const leftover = [];
  let establishment = '';

  const consider = (raw, asTitle) => {
    const line = String(raw || '').replace(/\s+/g, ' ').trim();
    if (!line) return;
    if (isFormCTamilNaduLwfPdfSignatoryText(line)) return;
    if (/name\s+of\s+the\s+establishment/i.test(line)) {
      const val = line.replace(/^name\s+of\s+the\s+establishment\s*:?\s*/i, '').trim();
      if (val && !/^name\s+of\s+the\s+establishment$/i.test(val)) establishment = val;
      return;
    }
    if (asTitle) {
      if (/pvt\.?\s*ltd|private\s+limited|energy|road|district|taluk|theni\s+site/i.test(line)) {
        leftover.push(line);
        return;
      }
      nextTitles.push(line);
      return;
    }
    leftover.push(line);
  };

  titleList.forEach((t) => consider(t, true));
  fieldList.forEach((f) => consider(f, false));

  if (!establishment) {
    const companyLike = leftover.find((l) =>
      /pvt\.?\s*ltd|private\s+limited|energy|road|district|taluk|theni\s+site/i.test(l)
    );
    if (companyLike) establishment = companyLike;
  }

  const nextFields = [formatFormCTamilNaduLwfEstablishmentLine(establishment)];
  leftover.forEach((l) => {
    if (
      /pvt\.?\s*ltd|private\s+limited|energy|road|district|taluk|theni\s+site/i.test(l) &&
      establishment
    ) {
      return;
    }
    if (isFormCTamilNaduLwfPdfSignatoryText(l)) return;
    nextFields.push(l);
  });

  const seen = new Set();
  return {
    titles: nextTitles.filter((t) => {
      const k = String(t || '').toLowerCase();
      if (seen.has(k)) return false;
      seen.add(k);
      return true;
    }),
    fields: nextFields.filter((f) => {
      const k = String(f || '').toLowerCase();
      if (seen.has(k)) return false;
      seen.add(k);
      return true;
    }),
  };
}

export function resolveFormCLabourWelfareYear(selectedMonthStr, item, fallbackLine = '') {
  const fromDue = extractYearToken(item?.dueDate || item?.DueDate || '');
  if (fromDue) return fromDue;
  const fromMonth = extractYearToken(selectedMonthStr);
  if (fromMonth) return fromMonth;
  const fromFallback = extractYearToken(fallbackLine);
  if (fromFallback) return fromFallback;
  return new Date().getFullYear();
}

function extractYearToken(value) {
  if (value == null || value === '') return null;
  const s = String(value);
  const m4 = s.match(/\b(19|20)\d{2}\b/);
  if (m4) return parseInt(m4[0], 10);
  return null;
}

export function applyFormCLabourWelfareYearToSubtitle(subtitle, year) {
  const y = Number(year);
  const yUse = Number.isFinite(y) && y >= 1900 ? y : new Date().getFullYear();
  const raw = String(subtitle || '').replace(/\s+/g, ' ').trim();
  if (!raw || /register\s+of\s+fines\s+and\s+unpaid|unpaid\s+accumulations/i.test(raw)) {
    return `${FORM_C_SUBTITLE_BASE} - ${yUse}`;
  }
  return raw;
}

/**
 * Rewrite quarter column labels so truncated / sample years become the corresponding year.
 * e.g. "Quarter ending 31-March-14 (2)" → "Quarter ending 31-March-2026 (2)"
 */
export function applyFormCLabourWelfareYearToHeaderLabel(header, year) {
  const y = Number(year);
  const yUse = Number.isFinite(y) && y >= 1900 ? y : new Date().getFullYear();
  const h = String(header || '');
  if (!/quarter\s+ending/i.test(h)) return h;
  return h.replace(
    /(quarter\s+ending\s+\d{1,2}[-/\s]*[A-Za-z]+)[-/\s]*(\d{2,4})(\b)/i,
    (_, prefix, _oldYear, boundary) => `${prefix}-${yUse}${boundary}`
  );
}

export function applyFormCLabourWelfareYearToHeaders(headers, year) {
  if (!Array.isArray(headers)) return headers;
  return headers.map((h) => applyFormCLabourWelfareYearToHeaderLabel(h, year));
}

export function applyFormCLabourWelfareYearToFormHeader(formHeader, year) {
  if (!formHeader || typeof formHeader !== 'object') return formHeader;
  const subtitle = applyFormCLabourWelfareYearToSubtitle(formHeader.subtitle, year);
  if (subtitle === formHeader.subtitle) return formHeader;
  return { ...formHeader, subtitle };
}

/** Simple column-index remap when Form C quarter header keys change. */
export function remapFormCLabourWelfareRowsToHeaders(rows, oldHeaders, newHeaders) {
  if (!Array.isArray(rows) || !Array.isArray(oldHeaders) || !Array.isArray(newHeaders)) {
    return Array.isArray(rows) ? rows : [];
  }
  if (oldHeaders.length === 0 || newHeaders.length === 0) return rows;
  const changed = oldHeaders.some((h, i) => h !== newHeaders[i]);
  if (!changed) return rows;
  return rows.map((row) => {
    if (!row || typeof row !== 'object' || Array.isArray(row)) return row;
    const out = {};
    const n = Math.max(oldHeaders.length, newHeaders.length);
    for (let i = 0; i < n; i += 1) {
      const oldH = oldHeaders[i];
      const newH = newHeaders[i] || oldH;
      if (!newH) continue;
      if (oldH != null && Object.prototype.hasOwnProperty.call(row, oldH)) {
        out[newH] = row[oldH];
      } else if (Object.prototype.hasOwnProperty.call(row, newH)) {
        out[newH] = row[newH];
      } else {
        out[newH] = '';
      }
    }
    Object.keys(row).forEach((k) => {
      if (!(k in out) && !oldHeaders.includes(k)) out[k] = row[k];
    });
    return out;
  });
}

function formCExcelJsCellText(val) {
  if (val == null) return '';
  if (typeof val === 'string' || typeof val === 'number' || typeof val === 'boolean') {
    return String(val).replace(/\r?\n/g, ' ').replace(/\s+/g, ' ').trim();
  }
  if (val instanceof Date) return val.toISOString();
  if (typeof val === 'object') {
    if (Array.isArray(val.richText)) {
      return val.richText
        .map((rt) => rt?.text || '')
        .join('')
        .replace(/\r?\n/g, ' ')
        .replace(/\s+/g, ' ')
        .trim();
    }
    if (val.text != null) return String(val.text).replace(/\r?\n/g, ' ').replace(/\s+/g, ' ').trim();
    if (val.result != null) return formCExcelJsCellText(val.result);
  }
  return '';
}

function parseFormCMergeLabel(label) {
  const parts = String(label || '').split(':');
  if (parts.length !== 2) return null;
  const start = parts[0].match(/^([A-Z]+)(\d+)$/i);
  const end = parts[1].match(/^([A-Z]+)(\d+)$/i);
  if (!start || !end) return null;
  const colToNum = (letters) => {
    let n = 0;
    const s = String(letters || '').toUpperCase();
    for (let i = 0; i < s.length; i += 1) n = n * 26 + (s.charCodeAt(i) - 64);
    return n;
  };
  return {
    r1: parseInt(start[2], 10),
    r2: parseInt(end[2], 10),
    c1: colToNum(start[1]),
    c2: colToNum(end[1]),
  };
}

function looksLikeFormCTitle(text) {
  return /^\*?form\s*[-–]?\s*c\b/i.test(String(text || '').trim());
}

function looksLikeFormCReference(text) {
  const t = String(text || '').trim();
  return /see\s+rule/i.test(t) && /labour\s+welfare/i.test(t);
}

function looksLikeFormCSubtitle(text) {
  return /register\s+of\s+fines\s+and\s+unpaid|unpaid\s+accumulations\s+for\s+the\s+year/i.test(
    String(text || '')
  );
}

function worksheetLooksLikeFormCTamilNaduLwf(worksheet) {
  if (!worksheet || typeof worksheet.getCell !== 'function') return false;
  let hits = 0;
  for (let r = 1; r <= 20; r += 1) {
    for (let c = 1; c <= 8; c += 1) {
      const t = formCExcelJsCellText(worksheet.getCell(r, c)?.value);
      if (!t) continue;
      if (looksLikeFormCTitle(t) || looksLikeFormCReference(t) || looksLikeFormCSubtitle(t)) {
        hits += 1;
      }
      if (/details\s+of\s+fines/i.test(t) && /unpaid\s+accumulations/i.test(t)) hits += 1;
      if (/quarter\s+ending/i.test(t)) hits += 1;
      if (hits >= 2) return true;
    }
  }
  return false;
}

/**
 * Center Form-C / See rule… / Register of Fines… across the table width (A–E style).
 * Matches the official Tamil Nadu LWF Form C header band.
 */
export function ensureFormCTamilNaduLwfTitleLayout(worksheet, options = {}) {
  if (!worksheet) return false;
  if (!options.force && !worksheetLooksLikeFormCTamilNaduLwf(worksheet)) return false;

  const colFrom = Math.max(1, Number(options.colFrom) || 1);
  const colTo = Math.max(colFrom, Number(options.colTo) || colFrom + 4);
  const scanTo = Math.max(1, Number(options.scanTo) || 12);

  const findRow = (matcher) => {
    for (let r = 1; r <= scanTo; r += 1) {
      for (let c = 1; c <= Math.max(colTo + 2, 8); c += 1) {
        const t = formCExcelJsCellText(worksheet.getCell(r, c)?.value);
        if (matcher(t)) return { row: r, col: c, text: t };
      }
    }
    return null;
  };

  const titleHit = findRow(looksLikeFormCTitle);
  const refHit = findRow(looksLikeFormCReference);
  const subHit = findRow(looksLikeFormCSubtitle);
  if (!titleHit && !refHit && !subHit) return false;

  const mergeLabels = Array.isArray(worksheet.model?.merges) ? [...worksheet.model.merges] : [];
  const unmergeCovering = (r1, c1, r2, c2) => {
    mergeLabels.forEach((label) => {
      const m = parseFormCMergeLabel(label);
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

  const writeCenteredBand = (hit, { bold = true, size } = {}) => {
    if (!hit || hit.row < 1 || !hit.text) return;
    const row = hit.row;
    // Clear stray copies across the band before re-centering.
    for (let c = 1; c <= Math.max(colTo + 4, 10); c += 1) {
      const cell = worksheet.getCell(row, c);
      const t = formCExcelJsCellText(cell?.value);
      if (
        looksLikeFormCTitle(t) ||
        looksLikeFormCReference(t) ||
        looksLikeFormCSubtitle(t)
      ) {
        cell.value = null;
      }
    }
    unmergeCovering(row, colFrom, row, colTo);
    try {
      worksheet.mergeCells(row, colFrom, row, colTo);
    } catch (_) {
      /* overlap — still write left cell */
    }
    const cell = worksheet.getCell(row, colFrom);
    cell.value = hit.text;
    cell.alignment = {
      ...(cell.alignment || {}),
      horizontal: 'center',
      vertical: 'middle',
      wrapText: false,
      textRotation: 0,
    };
    cell.font = {
      ...(cell.font || {}),
      bold,
      ...(size ? { size } : {}),
    };
  };

  writeCenteredBand(titleHit, { bold: true, size: 12 });
  writeCenteredBand(refHit, { bold: false, size: 10 });
  writeCenteredBand(subHit, { bold: true, size: 11 });
  return true;
}

/**
 * Align Form C table: centered headers, left labels, centered quarter values,
 * and vertical merges so "Nil" sits mid-block like the official register.
 */
export function ensureFormCTamilNaduLwfTableAlignment(worksheet, options = {}) {
  if (!worksheet) return false;

  const headerRow = Math.max(1, Number(options.headerRow) || 0);
  const dataStartRow = Math.max(headerRow + 1, Number(options.dataStartRow) || headerRow + 1);
  const colFrom = Math.max(1, Number(options.colFrom) || 1);
  const colTo = Math.max(colFrom, Number(options.colTo) || colFrom + 4);
  const quarterCols = Array.isArray(options.quarterCols)
    ? options.quarterCols.filter((c) => c > colFrom)
    : [];
  const categorySlots = Array.isArray(options.categorySlots) ? options.categorySlots : [];
  if (headerRow < 1) return false;

  // Table header row — center + wrap (quarter ending labels are multi-line).
  for (let c = colFrom; c <= colTo; c += 1) {
    const cell = worksheet.getCell(headerRow, c);
    const t = formCExcelJsCellText(cell?.value);
    if (!t) continue;
    cell.alignment = {
      ...(cell.alignment || {}),
      horizontal: 'center',
      vertical: 'middle',
      wrapText: true,
      textRotation: 0,
    };
  }
  try {
    const hdr = worksheet.getRow(headerRow);
    if (hdr) hdr.height = Math.max(Number(hdr.height) || 0, 36);
  } catch (_) {
    /* ignore */
  }

  const mergeLabels = () =>
    Array.isArray(worksheet.model?.merges) ? [...worksheet.model.merges] : [];

  const findLabelMergeExtent = (row, col) => {
    for (const label of mergeLabels()) {
      const m = parseFormCMergeLabel(label);
      if (!m) continue;
      if (row >= m.r1 && row <= m.r2 && col >= m.c1 && col <= m.c2) {
        return { r1: m.r1, r2: m.r2 };
      }
    }
    // Infer next slot boundary when template has no label merge.
    const idx = categorySlots.findIndex((s) => s.row === row);
    if (idx >= 0 && categorySlots[idx + 1]?.row > row) {
      return { r1: row, r2: categorySlots[idx + 1].row - 1 };
    }
    return { r1: row, r2: row };
  };

  const unmergeCovering = (r1, c1, r2, c2) => {
    mergeLabels().forEach((label) => {
      const m = parseFormCMergeLabel(label);
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

  const alignBody = (excelRow) => {
    const labelCell = worksheet.getCell(excelRow, colFrom);
    if (formCExcelJsCellText(labelCell?.value)) {
      labelCell.alignment = {
        ...(labelCell.alignment || {}),
        horizontal: 'left',
        vertical: 'top',
        wrapText: true,
      };
    }
    const extent = findLabelMergeExtent(excelRow, colFrom);
    const cols = quarterCols.length > 0 ? quarterCols : [];
    for (let qi = 0; qi < cols.length; qi += 1) {
      const c = cols[qi];
      // Prefer value on the slot master row; fall back to any cell in the block.
      let text = formCExcelJsCellText(worksheet.getCell(excelRow, c)?.value);
      if (!text) {
        for (let r = extent.r1; r <= extent.r2; r += 1) {
          text = formCExcelJsCellText(worksheet.getCell(r, c)?.value);
          if (text) break;
        }
      }
      if (!text) continue;

      unmergeCovering(extent.r1, c, extent.r2, c);
      for (let r = extent.r1; r <= extent.r2; r += 1) {
        worksheet.getCell(r, c).value = null;
      }
      if (extent.r2 > extent.r1) {
        try {
          worksheet.mergeCells(extent.r1, c, extent.r2, c);
        } catch (_) {
          /* ignore */
        }
      }
      const cell = worksheet.getCell(extent.r1, c);
      cell.value = text;
      cell.alignment = {
        ...(cell.alignment || {}),
        horizontal: 'center',
        vertical: 'middle',
        wrapText: true,
      };
    }
  };

  if (categorySlots.length > 0) {
    categorySlots.forEach((slot) => {
      if (slot?.row >= dataStartRow) alignBody(slot.row);
    });
  } else {
    // Scan a short body band when slots were not provided.
    for (let r = dataStartRow; r <= dataStartRow + 20; r += 1) {
      const label = formCExcelJsCellText(worksheet.getCell(r, colFrom)?.value);
      if (!label) continue;
      if (/see\s+definition\s+of|authorised\s+signator/i.test(label)) break;
      alignBody(r);
    }
  }

  return true;
}

/**
 * Left-align establishment line above the table and merge across the form width.
 */
export function ensureFormCTamilNaduLwfEstablishmentLayout(worksheet, options = {}) {
  if (!worksheet) return false;
  const headerRow = Math.max(1, Number(options.headerRow) || 0);
  const colFrom = Math.max(1, Number(options.colFrom) || 1);
  const colTo = Math.max(colFrom, Number(options.colTo) || colFrom + 4);
  const establishmentText = String(options.establishmentText || '').trim();
  if (headerRow < 2) return false;

  let estRow = Math.max(1, headerRow - 1);
  let existing = '';
  for (let r = Math.max(1, headerRow - 6); r < headerRow; r += 1) {
    const t = formCExcelJsCellText(worksheet.getCell(r, colFrom)?.value);
    if (
      !t ||
      /^name\s+of\s+the\s+establishment/i.test(t) ||
      (establishmentText && t.includes(establishmentText.slice(0, 24)))
    ) {
      estRow = r;
      existing = t;
      break;
    }
  }

  const stripEstLabel = (txt) =>
    String(txt || '')
      .replace(/^name\s+of\s+the\s+establishment\s*:?\s*/i, '')
      .trim();

  // Keep "Name of the Establishment :" prefix (Excel + PDF), value left-aligned across form width.
  const valueOnly = stripEstLabel(establishmentText) || stripEstLabel(existing) || '';
  const display = valueOnly
    ? formatFormCTamilNaduLwfEstablishmentLine(valueOnly)
    : existing
      ? formatFormCTamilNaduLwfEstablishmentLine(existing)
      : '';
  if (!display || display === `${FORM_C_TN_ESTABLISHMENT_LABEL} :`) return false;

  const mergeLabels = Array.isArray(worksheet.model?.merges) ? [...worksheet.model.merges] : [];
  mergeLabels.forEach((label) => {
    const m = parseFormCMergeLabel(label);
    if (!m) return;
    if (!(m.r2 < estRow || m.r1 > estRow || m.c2 < colFrom || m.c1 > colTo)) {
      try {
        worksheet.unMergeCells(label);
      } catch (_) {
        /* ignore */
      }
    }
  });

  try {
    worksheet.mergeCells(estRow, colFrom, estRow, colTo);
  } catch (_) {
    /* ignore */
  }
  for (let c = colFrom + 1; c <= colTo; c += 1) {
    try {
      worksheet.getCell(estRow, c).value = null;
    } catch (_) {
      /* ignore */
    }
  }

  const cell = worksheet.getCell(estRow, colFrom);
  cell.value = display;
  cell.alignment = {
    ...(cell.alignment || {}),
    horizontal: 'left',
    vertical: 'middle',
    wrapText: true,
    textRotation: 0,
  };
  cell.font = {
    ...(cell.font || {}),
    bold: true,
  };
  try {
    const wsRow = worksheet.getRow(estRow);
    if (wsRow) wsRow.height = Math.max(Number(wsRow.height) || 0, 30);
  } catch (_) {
    /* ignore */
  }
  return true;
}
