/** Tamil Nadu LWF Form C — Register of Fines and Unpaid Accumulations (rule 29). */

const FORM_C_SUBTITLE_BASE = 'Register of Fines and Unpaid Accumulations for the year';

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
