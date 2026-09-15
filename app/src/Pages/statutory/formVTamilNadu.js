/**
 * Tamil Nadu Form V — Register of Employment (daily attendance columns).
 * Day columns must match the selected month (May → 31, April → 30), not a fixed 30.
 * Total Days Worked ← payroll Paid_days.
 */

import { flattenPayrollEarningColumns, readPayrollScalar } from '../../utils/payrollEarnings';

/**
 * Filename Form_V_-_TamilNadu.xlsx: word-boundary after "v" fails because "_" is a word char.
 * Must not match Form VI / VII / XIV / XV / XVI / XXII.
 */
export function looksLikeFormVTamilNaduFilename(text) {
  const s = String(text || '').toLowerCase();
  return /(?:^|[^a-z0-9])form[\s._-]*v(?:[^a-z0-9]|$)/i.test(s);
}

/** True for Form V / Form_V_-_TamilNadu / Tamil Nadu Register of Employment — never Form VI or AP Form XXII. */
export function isFormVStatutoryDownloadHint(...parts) {
  const blob = parts
    .flat()
    .filter((p) => p != null && String(p).trim() !== '')
    .join(' ')
    .toLowerCase();
  if (!blob) return false;
  if (/\bform[\s._-]*xxii\b|\bform[\s._-]*22\b/.test(blob) && !looksLikeFormVTamilNaduFilename(blob)) {
    return false;
  }
  if (/\bform[\s._-]*11\b/.test(blob) && /rajasthan|\brj\b/.test(blob)) return false;
  if (/(?:^|[^a-z0-9])form[\s._-]*vi(?:[^a-z0-9]|$)/.test(blob) && !looksLikeFormVTamilNaduFilename(blob)) {
    return false;
  }
  if (looksLikeFormVTamilNaduFilename(blob)) return true;
  const isTamil = /tamil[\s._-]*nadu|tamilnadu/.test(blob);
  const isRegister = /register\s+of\s+employment/.test(blob);
  if (isTamil && isRegister) return true;
  if (
    isTamil &&
    /(?:^|[^a-z0-9])form[\s._-]*5(?:[^a-z0-9]|$)/.test(blob) &&
    !/\bform[\s._-]*5[0-9]/.test(blob)
  ) {
    return true;
  }
  return false;
}

export function resolveFormVTamilNaduDayNumberFromHeader(header) {
  const h = String(header || '').trim();
  if (/^\d{1,2}$/.test(h)) {
    const n = Number(h);
    return n >= 1 && n <= 31 ? n : 0;
  }
  const m =
    h.match(/(?:dates?|attendance|daily[^0-9]*)[^0-9]*(\d{1,2})$/i) ||
    h.match(/_(\d{1,2})$/) ||
    h.match(/(?:^|[^0-9])(\d{1,2})$/);
  if (m) {
    const n = Number(m[1]);
    return n >= 1 && n <= 31 ? n : 0;
  }
  return 0;
}

export function isFormVTamilNaduDayHeader(header) {
  return resolveFormVTamilNaduDayNumberFromHeader(header) >= 1;
}

/** Calendar day count for a month name (uses year for Feb leap years). */
export function getFormVTamilNaduMonthDayCount(monthName, year = new Date().getFullYear()) {
  const MONTHS = [
    'january',
    'february',
    'march',
    'april',
    'may',
    'june',
    'july',
    'august',
    'september',
    'october',
    'november',
    'december',
  ];
  const raw = String(monthName || '').trim().toLowerCase();
  if (!raw) return 31;
  const idx = MONTHS.findIndex((m) => m.startsWith(raw) || raw.startsWith(m.slice(0, 3)));
  if (idx < 0) return 31;
  const y = Number(year);
  const yUse = Number.isFinite(y) && y >= 1900 ? y : new Date().getFullYear();
  return new Date(yUse, idx + 1, 0).getDate();
}

function makeFormVTamilNaduDayHeader(sampleHeader, day) {
  const h = String(sampleHeader || '').trim();
  if (!h || /^\d{1,2}$/.test(h)) return String(day);
  if (/_\d{1,2}$/.test(h)) return h.replace(/_\d{1,2}$/, `_${day}`);
  if (/\d{1,2}$/.test(h)) return h.replace(/\d{1,2}$/, String(day));
  return String(day);
}

/**
 * Expand / trim Form V attendance day columns to the selected month length.
 * Example: template has 1–30 → May becomes 1–31; April stays 1–30.
 */
export function resolveFormVTamilNaduTableHeaders(headers, daysInMonth = 31) {
  const days = Math.min(Math.max(Number(daysInMonth) || 31, 28), 31);
  const list = Array.isArray(headers) ? headers.filter((h) => h != null && String(h).trim() !== '') : [];
  if (list.length === 0) return list;

  const dayEntries = [];
  list.forEach((header, index) => {
    const day = resolveFormVTamilNaduDayNumberFromHeader(header);
    if (day >= 1 && day <= 31) dayEntries.push({ header, day, index });
  });
  // Need a contiguous day band; otherwise leave headers alone.
  if (dayEntries.length < 5) return list;

  const firstDayIdx = Math.min(...dayEntries.map((e) => e.index));
  const lastDayIdx = Math.max(...dayEntries.map((e) => e.index));
  const prefix = list.slice(0, firstDayIdx);
  const suffix = list.slice(lastDayIdx + 1);
  const byDay = new Map();
  dayEntries.forEach((e) => {
    if (!byDay.has(e.day)) byDay.set(e.day, e.header);
  });

  const sample =
    byDay.get(1) ||
    byDay.get(Math.min(...byDay.keys())) ||
    String(dayEntries[0].header);

  const dayHeaders = [];
  for (let d = 1; d <= days; d += 1) {
    dayHeaders.push(byDay.has(d) ? byDay.get(d) : makeFormVTamilNaduDayHeader(sample, d));
  }

  const next = [...prefix, ...dayHeaders, ...suffix];
  // No-op when already exact.
  if (
    next.length === list.length &&
    next.every((h, i) => String(h) === String(list[i]))
  ) {
    return list;
  }
  return next;
}

/** True when headers look like Form V TN employment day grid (identity + many day cols). */
export function headersIndicateFormVTamilNaduDayGrid(headers) {
  const list = Array.isArray(headers) ? headers : [];
  const dayCount = list.filter((h) => isFormVTamilNaduDayHeader(h)).length;
  if (dayCount < 5) return false;
  const joined = list.map((h) => String(h || '').toLowerCase()).join(' ');
  if (/register\s+of\s+employment/.test(joined)) return true;
  if (/national\s+holiday|festival\s+holiday|remarks?/.test(joined) && dayCount >= 20) return true;
  return dayCount >= 28;
}

export function formVTamilNaduHeaderNorm(txt) {
  return String(txt || '')
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

/** "Total Days Worked" summary column (not a calendar day header). */
export function isFormVTamilNaduTotalDaysWorkedHeader(header) {
  if (isFormVTamilNaduDayHeader(header)) return false;
  const s = formVTamilNaduHeaderNorm(header);
  if (!s) return false;
  if (s.includes('overtime') || s.includes('loss') || s.includes('lop') || s.includes('holiday')) {
    return false;
  }
  if (/total\s*days?\s*worked/.test(s)) return true;
  return s.includes('total') && (s.includes('day') || s.includes('days')) && s.includes('work');
}

/** Total Hours Worked — Form V TN uses Paid_days × 8. */
export function isFormVTamilNaduTotalHoursWorkedHeader(header) {
  if (isFormVTamilNaduDayHeader(header)) return false;
  const s = formVTamilNaduHeaderNorm(header);
  if (!s) return false;
  if (s.includes('overtime') || s.includes('normal') || s.includes('daily')) return false;
  return /total\s*hours?\s*worked/.test(s) || (s.includes('total') && s.includes('hour') && s.includes('work'));
}

/** Number of days on Loss of Pay — Form V TN uses monthDays − Paid_days. */
export function isFormVTamilNaduLossOfPayHeader(header) {
  if (isFormVTamilNaduDayHeader(header)) return false;
  const s = formVTamilNaduHeaderNorm(header);
  if (!s) return false;
  return (/loss/.test(s) && /pay/.test(s)) || /\blop\b/.test(s);
}

/** Payroll Paid_days (and common aliases) for Form V Total Days Worked. */
export function readFormVTamilNaduPaidDays(payrollRow) {
  if (!payrollRow || payrollRow.fetch_error) return '';
  const flat = flattenPayrollEarningColumns(payrollRow);
  const keys = [
    'paid_days',
    'Paid_days',
    'Paid Days',
    'paidDays',
    'days_worked',
    'Days Worked',
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
  const fromFlat = readPayrollScalar(flat, keys, patterns);
  if (fromFlat !== '') {
    const raw = String(fromFlat).trim();
    if (isFormVTamilNaduJunkSummaryValue(raw)) return '';
    const num = Number(raw.replace(/,/g, ''));
    if (Number.isFinite(num) && num >= 0) return String(num);
  }
  if (payrollRow !== flat) {
    const fromRow = readPayrollScalar(payrollRow, keys, patterns);
    if (fromRow !== '') {
      const raw = String(fromRow).trim();
      if (isFormVTamilNaduJunkSummaryValue(raw)) return '';
      const num = Number(raw.replace(/,/g, ''));
      if (Number.isFinite(num) && num >= 0) return String(num);
    }
  }
  return '';
}

/**
 * Reject tenure/experience text like "1 month(s)" from Total Days / Hours cells.
 * Those columns must stay blank when Paid_days is missing — never invent defaults.
 */
export function isFormVTamilNaduJunkSummaryValue(value) {
  const s = String(value ?? '').trim();
  if (!s) return false;
  if (/^enter\s+/i.test(s) || /^select\s+/i.test(s) || /^nil$/i.test(s)) return true;
  if (/[ap]\.?m\.?$/i.test(s) || /^\d{1,2}:\d{2}\b/i.test(s)) return true;
  if (/month|year|week|day\s*\(|experience|tenure/i.test(s) && !/^\d+(\.\d+)?$/.test(s)) {
    return true;
  }
  if (!/^\d+(\.\d+)?$/.test(s.replace(/,/g, ''))) return true;
  return false;
}

/** Keep only numeric summary values; blank when missing or junk like "1 month(s)". */
export function sanitizeFormVTamilNaduSummaryNumericValue(value) {
  const s = String(value ?? '').trim();
  if (!s || isFormVTamilNaduJunkSummaryValue(s)) return '';
  const n = Number(s.replace(/,/g, ''));
  if (!Number.isFinite(n) || n < 0) return '';
  return String(Math.round(n * 100) / 100);
}

/** Total Hours Worked = Paid_days × 8. */
export function computeFormVTamilNaduTotalHoursWorked(paidDays) {
  if (paidDays === '' || paidDays == null) return '';
  const raw = String(paidDays).trim();
  if (!raw || isFormVTamilNaduJunkSummaryValue(raw)) return '';
  const n = Number(raw.replace(/,/g, ''));
  if (!Number.isFinite(n) || n < 0) return '';
  return String(Math.round(n * 8 * 100) / 100);
}

/** Loss of Pay days = monthDays − Paid_days (never negative). */
export function computeFormVTamilNaduLossOfPayDays(paidDays, daysInMonth = 31) {
  if (paidDays === '' || paidDays == null) return '';
  const raw = String(paidDays).trim();
  if (!raw || isFormVTamilNaduJunkSummaryValue(raw)) return '';
  const paid = Number(raw.replace(/,/g, ''));
  const monthDays = Math.min(Math.max(Number(daysInMonth) || 31, 28), 31);
  if (!Number.isFinite(paid)) return '';
  return String(Math.max(0, monthDays - paid));
}

/**
 * Fill Form V summary columns from payroll Paid_days:
 * - Total Days Worked ← Paid_days
 * - Total Hours Worked ← Paid_days × 8
 * - Loss of Pay ← monthDays − Paid_days
 * Clears junk like "1 month(s)" when Paid_days is missing.
 * @returns {number} rows updated
 */
export function applyFormVTamilNaduPaidDaysToRows(
  mappedData,
  employeesForMapping,
  tableHeaders,
  options = {}
) {
  if (!Array.isArray(mappedData) || mappedData.length === 0) return 0;
  const headers = Array.isArray(tableHeaders) ? tableHeaders : [];
  const daysHeader = headers.find((h) => isFormVTamilNaduTotalDaysWorkedHeader(h));
  const hoursHeader = headers.find((h) => isFormVTamilNaduTotalHoursWorkedHeader(h));
  const lopHeader = headers.find((h) => isFormVTamilNaduLossOfPayHeader(h));
  if (!daysHeader && !hoursHeader && !lopHeader) return 0;

  const resolvePayrollRow =
    typeof options.resolvePayrollRow === 'function' ? options.resolvePayrollRow : () => null;
  const unwrapEmp =
    typeof options.unwrapEmp === 'function'
      ? options.unwrapEmp
      : (item) => (item && (item.Employee || item.employee || item)) || null;
  const sanitize =
    typeof options.sanitizeValue === 'function' ? options.sanitizeValue : (v) => String(v ?? '');
  const overwrite = options.overwrite !== false;
  const daysInMonth = Math.min(Math.max(Number(options.daysInMonth) || 31, 28), 31);

  let hits = 0;
  mappedData.forEach((row, index) => {
    if (!row || typeof row !== 'object') return;
    const emp = unwrapEmp(Array.isArray(employeesForMapping) ? employeesForMapping[index] : null);
    const payrollRow = resolvePayrollRow(emp, row, index);
    const paidDays = readFormVTamilNaduPaidDays(payrollRow);
    let changed = false;

    const clearIfJunkOrOverwrite = (header) => {
      if (!header) return;
      const cur = String(row[header] ?? '').trim();
      if (!cur) return;
      if (overwrite || isFormVTamilNaduJunkSummaryValue(cur)) {
        row[header] = '';
        changed = true;
      }
    };

    if (paidDays === '') {
      // No Paid_days — blank days/hours/LOP (never keep "1 month(s)" / non-numeric junk).
      clearIfJunkOrOverwrite(daysHeader);
      clearIfJunkOrOverwrite(hoursHeader);
      clearIfJunkOrOverwrite(lopHeader);
      if (changed) hits += 1;
      return;
    }

    if (daysHeader && (overwrite || String(row[daysHeader] ?? '').trim() === '' || isFormVTamilNaduJunkSummaryValue(row[daysHeader]))) {
      row[daysHeader] = sanitize(paidDays);
      changed = true;
    }
    if (hoursHeader && (overwrite || String(row[hoursHeader] ?? '').trim() === '' || isFormVTamilNaduJunkSummaryValue(row[hoursHeader]))) {
      const hours = computeFormVTamilNaduTotalHoursWorked(paidDays);
      if (hours !== '') {
        row[hoursHeader] = sanitize(hours);
        changed = true;
      } else {
        clearIfJunkOrOverwrite(hoursHeader);
      }
    }
    if (lopHeader && (overwrite || String(row[lopHeader] ?? '').trim() === '' || isFormVTamilNaduJunkSummaryValue(row[lopHeader]))) {
      const lop = computeFormVTamilNaduLossOfPayDays(paidDays, daysInMonth);
      if (lop !== '') {
        row[lopHeader] = sanitize(lop);
        changed = true;
      } else {
        clearIfJunkOrOverwrite(lopHeader);
      }
    }
    if (changed) hits += 1;
  });
  return hits;
}

/**
 * Ensure the Excel template physically has day columns 1..daysInMonth.
 * Form_V template often stops at day 30 — for May we must insert day 31 before writing
 * so Total Days Worked / Hours / LOP stay aligned.
 *
 * Always prefer a manual right-shift: ExcelJS spliceColumns is unreliable on this template,
 * and summary titles are often split vertically ("Total" / "Days" / "Worked") so a
 * single-cell label check misses them.
 *
 * @returns {{ dayColsByNumber: Map<number,number>, markerCols: number[], inserted: number }}
 */
export function ensureFormVTamilNaduExcelDayColumns(
  worksheet,
  {
    dayRow,
    dayColsByNumber,
    markerCols,
    daysInMonth = 31,
    copyStylesFromDay = null,
    maxScanRows = 80,
    maxScanCols = 80,
  } = {}
) {
  const days = Math.min(Math.max(Number(daysInMonth) || 31, 28), 31);
  const cols = dayColsByNumber instanceof Map ? new Map(dayColsByNumber) : new Map();
  const ordered = Array.isArray(markerCols) ? [...markerCols] : [];
  if (!worksheet || cols.size === 0) {
    return { dayColsByNumber: cols, markerCols: ordered, inserted: 0 };
  }

  const cellText = (r, c) => {
    const val = worksheet.getCell(r, c)?.value;
    if (val == null) return '';
    if (typeof val === 'string' || typeof val === 'number' || typeof val === 'boolean') return String(val);
    if (typeof val === 'object') {
      if (Array.isArray(val.richText)) return val.richText.map((rt) => rt?.text || '').join('');
      if (val.text != null) return String(val.text);
      if (val.result != null) return String(val.result);
    }
    return '';
  };
  const norm = (txt) =>
    String(txt || '')
      .replace(/\r?\n/g, ' ')
      .replace(/\s+/g, ' ')
      .trim()
      .toLowerCase();
  const looksLikeSummaryLabel = (txt) => {
    const t = norm(txt);
    if (!t) return false;
    if (/total/.test(t) && /days?/.test(t) && /work/.test(t)) return true;
    if (/total/.test(t) && /hours?/.test(t) && /work/.test(t)) return true;
    if (/loss/.test(t) && /pay/.test(t)) return true;
    if (/benefit/.test(t) && /holiday/.test(t)) return true;
    if (/^remarks?$/.test(t)) return true;
    // Vertically split header fragments in Form V template
    if (/^total$/.test(t) || /^days?$/.test(t) || /^worked$/.test(t)) return true;
    if (/^hours?$/.test(t) || /^number$/.test(t) || /^loss$/.test(t) || /^of$/.test(t) || /^pay$/.test(t)) {
      return true;
    }
    return false;
  };
  const columnJoinedHeaderText = (col) => {
    const from = Math.max(1, (dayRow > 0 ? dayRow : 10) - 5);
    const to = Math.max(dayRow > 0 ? dayRow : 10, 12) + 3;
    const parts = [];
    for (let r = from; r <= to; r += 1) {
      const t = String(cellText(r, col) || '').trim();
      if (t) parts.push(t);
    }
    return parts.join(' ');
  };
  const columnLooksLikeSummary = (col) => {
    if (looksLikeSummaryLabel(columnJoinedHeaderText(col))) return true;
    const from = Math.max(1, (dayRow > 0 ? dayRow : 10) - 5);
    const to = Math.max(dayRow > 0 ? dayRow : 10, 12) + 3;
    for (let r = from; r <= to; r += 1) {
      if (looksLikeSummaryLabel(cellText(r, col))) return true;
    }
    return false;
  };
  const findFirstSummaryColAfter = (afterCol) => {
    const start = Math.max(1, Number(afterCol) || 0) + 1;
    const limit = Math.max(Number(maxScanCols) || 80, start + 40, 80);
    for (let c = start; c <= limit; c += 1) {
      if (columnLooksLikeSummary(c)) return c;
    }
    return start;
  };
  const shiftColumnsRightManual = (fromCol) => {
    const rowLimit = Math.max(Number(maxScanRows) || 80, worksheet.rowCount || 0, 60);
    const colLimit = Math.max(
      Number(maxScanCols) || 80,
      worksheet.columnCount || 0,
      fromCol + 40,
      90
    );
    const mergeLabels = Array.isArray(worksheet.model?.merges) ? [...worksheet.model.merges] : [];
    const decode = (a1) => {
      let n = 0;
      const s = String(a1 || '').toUpperCase();
      for (let i = 0; i < s.length; i += 1) n = n * 26 + (s.charCodeAt(i) - 64);
      return n;
    };
    const pendingRemerge = [];
    mergeLabels.forEach((label) => {
      const m = String(label || '').match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/i);
      if (!m) return;
      const c1 = decode(m[1]);
      const r1 = Number(m[2]);
      const c2 = decode(m[3]);
      const r2 = Number(m[4]);
      if (c2 < fromCol) return;
      try {
        worksheet.unMergeCells(label);
      } catch (_) {
        /* ignore */
      }
      if (c1 >= fromCol) pendingRemerge.push({ r1, c1: c1 + 1, r2, c2: c2 + 1 });
      else pendingRemerge.push({ r1, c1, r2, c2: c2 + 1 });
    });

    for (let c = colLimit; c >= fromCol; c -= 1) {
      for (let r = 1; r <= rowLimit; r += 1) {
        const src = worksheet.getCell(r, c);
        const dest = worksheet.getCell(r, c + 1);
        dest.value = src.value == null ? null : src.value;
        try {
          if (src.style) dest.style = JSON.parse(JSON.stringify(src.style));
        } catch (_) {
          /* ignore */
        }
      }
      try {
        const w = worksheet.getColumn(c)?.width;
        if (w) worksheet.getColumn(c + 1).width = w;
      } catch (_) {
        /* ignore */
      }
    }
    for (let r = 1; r <= rowLimit; r += 1) {
      worksheet.getCell(r, fromCol).value = null;
    }
    pendingRemerge.forEach(({ r1, c1, r2, c2 }) => {
      try {
        worksheet.mergeCells(r1, c1, r2, c2);
      } catch (_) {
        /* ignore */
      }
    });
  };
  const expandDailyHoursMergeTo = (insertAt, prevLastDayCol) => {
    try {
      const mergeLabels = Array.isArray(worksheet.model?.merges) ? [...worksheet.model.merges] : [];
      const decode = (a1) => {
        let n = 0;
        const s = String(a1 || '').toUpperCase();
        for (let i = 0; i < s.length; i += 1) n = n * 26 + (s.charCodeAt(i) - 64);
        return n;
      };
      mergeLabels.forEach((label) => {
        const m = String(label || '').match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/i);
        if (!m) return;
        const c1 = decode(m[1]);
        const r1 = Number(m[2]);
        const c2 = decode(m[3]);
        const r2 = Number(m[4]);
        if (c2 !== prevLastDayCol || c1 >= insertAt) return;
        const joined = norm(
          Array.from({ length: Math.max(1, r2 - r1 + 1) }, (_, i) => cellText(r1 + i, c1)).join(' ')
        );
        if (!/daily|hours|overtime|work/.test(joined) && c2 - c1 < 5) return;
        try {
          worksheet.unMergeCells(label);
        } catch (_) {
          /* ignore */
        }
        try {
          worksheet.mergeCells(r1, c1, r2, insertAt);
        } catch (_) {
          /* ignore */
        }
      });
    } catch (_) {
      /* ignore */
    }
  };

  let inserted = 0;
  const maxDayPresent = cols.size > 0 ? Math.max(...cols.keys()) : 0;
  for (let day = maxDayPresent + 1; day <= days; day += 1) {
    if (cols.has(day)) continue;
    let prevCol = -1;
    for (let d = day - 1; d >= 1; d -= 1) {
      if (cols.has(d)) {
        prevCol = cols.get(d);
        break;
      }
    }
    const maxExisting =
      cols.size > 0 ? Math.max(...[...cols.values()].map((n) => Number(n) || 0)) : 0;
    // Form V TN layout: days G…AJ (1–30), then AK = Total Days Worked.
    // For a 31-day month insert immediately after day 30 → new AK = day 31, totals shift to AL+.
    let insertAt = prevCol > 0 ? prevCol + 1 : maxExisting > 0 ? maxExisting + 1 : -1;
    if (insertAt < 1) continue;
    // If an empty gap sits before the summary band, land on the first summary column.
    if (prevCol > 0) {
      const summaryAt = findFirstSummaryColAfter(prevCol);
      if (summaryAt > insertAt && summaryAt <= prevCol + 2) insertAt = summaryAt;
    }

    const prevLastDayCol = prevCol;
    // Always manual-shift — spliceColumns often no-ops on loaded Form V templates.
    shiftColumnsRightManual(insertAt);

    const shifted = new Map();
    cols.forEach((c, d) => {
      shifted.set(d, c >= insertAt ? c + 1 : c);
    });
    cols.clear();
    shifted.forEach((c, d) => cols.set(d, c));
    for (let i = 0; i < ordered.length; i += 1) {
      if (ordered[i] >= insertAt) ordered[i] += 1;
    }

    expandDailyHoursMergeTo(insertAt, prevLastDayCol);

    cols.set(day, insertAt);
    ordered.push(insertAt);
    ordered.sort((a, b) => a - b);

    if (dayRow > 0) {
      const dayCell = worksheet.getCell(dayRow, insertAt);
      dayCell.value = day;
      const styleFromDay = copyStylesFromDay || day - 1;
      const styleCol = cols.get(styleFromDay) || prevCol;
      if (styleCol > 0 && styleCol !== insertAt) {
        try {
          const src = worksheet.getCell(dayRow, styleCol);
          if (src.style) dayCell.style = JSON.parse(JSON.stringify(src.style));
        } catch (_) {
          /* ignore */
        }
        try {
          const srcWidth = worksheet.getColumn(styleCol)?.width;
          if (srcWidth) worksheet.getColumn(insertAt).width = srcWidth;
        } catch (_) {
          /* ignore */
        }
      }
    }
    inserted += 1;
  }

  return { dayColsByNumber: cols, markerCols: ordered, inserted };
}

/**
 * Locate Form V trailing summary columns by header text (after day columns).
 */
export function locateFormVTamilNaduSummaryExcelColumns(
  worksheet,
  {
    headerRow = 10,
    dayRow = -1,
    afterCol = 0,
    maxScanCols = 80,
    cellText = (cell) => String(cell?.value ?? ''),
  } = {}
) {
  const norm = (txt) =>
    String(txt || '')
      .replace(/\r?\n/g, ' ')
      .replace(/\s+/g, ' ')
      .trim()
      .toLowerCase();
  const out = {
    totalDaysWorked: null,
    totalHoursWorked: null,
    lossOfPay: null,
    nationalHolidayBenefit: null,
    festivalHolidayBenefit: null,
    remarks: null,
  };
  if (!worksheet) return out;
  const fromRow = Math.max(1, Math.min(headerRow, dayRow > 0 ? dayRow : headerRow) - 5);
  const toRow = Math.max(headerRow, dayRow > 0 ? dayRow : headerRow) + 3;
  const startCol = Math.max(1, Number(afterCol) || 1);

  const classifyJoined = (joined, c) => {
    const t = norm(joined);
    if (!t) return;
    if (!out.totalDaysWorked && /total/.test(t) && /days?/.test(t) && /work/.test(t)) {
      out.totalDaysWorked = c;
      return;
    }
    if (!out.totalHoursWorked && /total/.test(t) && /hours?/.test(t) && /work/.test(t)) {
      out.totalHoursWorked = c;
      return;
    }
    if (!out.lossOfPay && /loss/.test(t) && /pay/.test(t)) {
      out.lossOfPay = c;
      return;
    }
    // Skip "Approved Festival Holidays" / approval proceedings header boxes.
    if (/approved\s+festival|approval\s+proceedings/.test(t)) return;
    if (
      !out.nationalHolidayBenefit &&
      /national/.test(t) &&
      /holiday/.test(t) &&
      (/benefit/.test(t) || /availed/.test(t) || /working/.test(t))
    ) {
      out.nationalHolidayBenefit = c;
      return;
    }
    if (
      !out.festivalHolidayBenefit &&
      /festival/.test(t) &&
      /holiday/.test(t) &&
      (/benefit/.test(t) || /availed/.test(t) || /working/.test(t))
    ) {
      out.festivalHolidayBenefit = c;
      return;
    }
    if (!out.remarks && /^remarks?$/.test(t)) {
      out.remarks = c;
    }
  };

  // Prefer joined vertical header text — Form V splits "Total Days Worked" across rows.
  for (let c = startCol; c <= maxScanCols; c += 1) {
    const parts = [];
    for (let r = fromRow; r <= toRow; r += 1) {
      const piece = String(cellText(worksheet.getCell(r, c)) || '').trim();
      if (piece) parts.push(piece);
    }
    classifyJoined(parts.join(' '), c);
  }

  // Fallback: single-cell matches (if join missed)
  for (let r = fromRow; r <= toRow; r += 1) {
    for (let c = startCol; c <= maxScanCols; c += 1) {
      classifyJoined(cellText(worksheet.getCell(r, c)), c);
    }
  }

  // Template layout: after LOP come National / Festival / Remarks in order.
  if (out.lossOfPay > 0) {
    if (!out.nationalHolidayBenefit) out.nationalHolidayBenefit = out.lossOfPay + 1;
    if (!out.festivalHolidayBenefit) out.festivalHolidayBenefit = out.lossOfPay + 2;
    if (!out.remarks) out.remarks = out.lossOfPay + 3;
  }
  return out;
}

/**
 * Excel columns that must stay blank on Form V TN download
 * (National Holiday benefit, Festival Holiday benefit, Remarks).
 * @returns {number[]}
 */
export function locateFormVTamilNaduLeaveBlankExcelColumns(
  worksheet,
  {
    headerRow = 10,
    dayRow = -1,
    afterCol = 0,
    lastDayCol = 0,
    maxScanCols = 80,
    cellText = (cell) => String(cell?.value ?? ''),
  } = {}
) {
  const summary = locateFormVTamilNaduSummaryExcelColumns(worksheet, {
    headerRow,
    dayRow,
    afterCol,
    maxScanCols,
    cellText,
  });
  const cols = new Set();
  [summary.nationalHolidayBenefit, summary.festivalHolidayBenefit, summary.remarks].forEach((c) => {
    if (Number.isFinite(c) && c > 0) cols.add(c);
  });
  // Hard fallback when labels were not found (day-30 template: 40/41/42).
  if (cols.size === 0) {
    const anchor = lastDayCol > 0 ? lastDayCol : afterCol > 0 ? afterCol - 1 : 36;
    [anchor + 4, anchor + 5, anchor + 6].forEach((c) => cols.add(c));
  }
  return [...cols].sort((a, b) => a - b);
}

/** Excel hides values written onto merged non-master cells. */
export function writeFormVTamilNaduExcelVisibleCell(worksheet, row, col, value) {
  if (!worksheet || !Number.isFinite(row) || !Number.isFinite(col) || row < 1 || col < 1) return;
  let targetRow = row;
  let targetCol = col;
  try {
    const cell = worksheet.getCell(row, col);
    if (cell?.isMerged && cell.master) {
      targetRow = cell.master.row || targetRow;
      targetCol = cell.master.col || targetCol;
    }
  } catch (_) {
    /* write to the requested cell */
  }
  worksheet.getCell(targetRow, targetCol).value = value;
}

function formVIdentityHeaderNorm(txt) {
  return String(txt || '')
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

/** Classify Form V identity columns (never day / summary). */
export function formVTamilNaduIdentityColumnKind(header) {
  if (isFormVTamilNaduDayHeader(header)) return '';
  const s = formVIdentityHeaderNorm(header);
  if (!s) return '';
  if (/^(s\.?\s*no\.?|sl\.?\s*no\.?|sr\.?\s*no\.?|serial(\s+no\.?)?)$/.test(s) || /\bs\.?\s*no\.?\b/.test(s)) {
    return 'sno';
  }
  if (/identification|employee\s*id|emp\.?\s*id|token\s*no/.test(s)) return 'empid';
  if (/commence|begins|start/.test(s) && /time|work/.test(s)) return 'commence';
  if (/^rest(\s+interval)?$/.test(s) || /rest\s+interval/.test(s)) return 'rest';
  if ((/ends|cease/.test(s) && /time|work/.test(s)) || /time\s+at\s+which\s+work\s+ends/.test(s)) {
    return 'ends';
  }
  if (
    /name\s+of\s+the\s+(employee|person|workman|worker)/.test(s) ||
    /name\s+of\s+(employee|person|worker)/.test(s) ||
    /^name$/.test(s)
  ) {
    return 'name';
  }
  return '';
}

/**
 * Map Autofill / array rows onto header keys so Excel export can read names and times.
 */
export function coerceFormVTamilNaduRowToObject(row, headers = []) {
  if (!row) return null;
  const hdrs = Array.isArray(headers) ? headers : [];
  if (Array.isArray(row)) {
    const obj = {};
    hdrs.forEach((h, i) => {
      if (h == null || String(h).trim() === '') return;
      obj[h] = row[i];
    });
    row.forEach((v, i) => {
      if (v == null || String(v).trim() === '') return;
      if (obj[String(i)] == null) obj[String(i)] = v;
    });
    return obj;
  }
  if (typeof row !== 'object') return null;
  const keys = Object.keys(row).filter((k) => !String(k).startsWith('__'));
  const hasNamedHeader = hdrs.some((h) => h && Object.prototype.hasOwnProperty.call(row, h));
  if (hasNamedHeader) return row;
  const numericKeys = keys.filter((k) => /^\d+$/.test(k));
  if (numericKeys.length >= 3 && hdrs.length > 0) {
    const obj = { ...row };
    hdrs.forEach((h, i) => {
      if (!h) return;
      const fromIndex = row[i] != null ? row[i] : row[String(i)];
      if (fromIndex != null && String(fromIndex).trim() !== '' && (obj[h] == null || String(obj[h]).trim() === '')) {
        obj[h] = fromIndex;
      }
    });
    return obj;
  }
  return row;
}

function isMeaningfulFormVExportValue(value) {
  const text = String(value ?? '').trim();
  if (!text) return false;
  if (/^enter\s+/i.test(text)) return false;
  if (/^select\s+/i.test(text)) return false;
  return true;
}

/** Autofill stores P/A/WO/CL; older grids may still have Present/Absent. */
export function normalizeFormVTamilNaduAttendanceCode(value) {
  const raw = String(value ?? '').trim();
  if (!raw) return '';
  if (/^(P|A|WO|H|L|WOP|OD|SL|CL|EL)$/i.test(raw)) return raw.toUpperCase();
  const lower = raw.toLowerCase().replace(/\s+/g, ' ');
  if (/^present$|^on duty|^od$|^p$/.test(lower)) return 'P';
  if (/^absent$|^loss of pay$|^lop$|^a$/.test(lower)) return 'A';
  if (/weekend|week off|weekoff|^wo$|^holiday$/.test(lower)) return 'WO';
  if (/contingency|^c\/?l$|^cl$/.test(lower)) return 'CL';
  return '';
}

/**
 * Resolve a Form V identity/summary cell from an Autofill row (object or array).
 */
export function resolveFormVTamilNaduExportCellValue(row, header, headers = [], headerIndex = null) {
  if (isFormVTamilNaduDayHeader(header)) {
    const day = resolveFormVTamilNaduDayNumberFromHeader(header);
    return resolveFormVTamilNaduDayValue(row, day, headers);
  }
  const coerced = coerceFormVTamilNaduRowToObject(row, headers);
  if (!coerced || typeof coerced !== 'object') return '';
  const kind = formVTamilNaduIdentityColumnKind(header);
  const isSummaryNumeric =
    isFormVTamilNaduTotalDaysWorkedHeader(header) ||
    isFormVTamilNaduTotalHoursWorkedHeader(header) ||
    isFormVTamilNaduLossOfPayHeader(header);
  const pickResolved = (raw) => {
    if (!isMeaningfulFormVExportValue(raw)) return '';
    if (isSummaryNumeric) return sanitizeFormVTamilNaduSummaryNumericValue(raw);
    return raw;
  };
  if (header && Object.prototype.hasOwnProperty.call(coerced, header)) {
    const picked = pickResolved(coerced[header]);
    if (picked !== '') return picked;
  }
  if (
    headerIndex != null &&
    headerIndex >= 0 &&
    Array.isArray(headers) &&
    headers[headerIndex] &&
    Object.prototype.hasOwnProperty.call(coerced, headers[headerIndex])
  ) {
    const picked = pickResolved(coerced[headers[headerIndex]]);
    if (picked !== '') return picked;
  }
  const rowKeys = Object.keys(coerced).filter((k) => !String(k).startsWith('__'));
  const target = formVTamilNaduHeaderNorm(header).replace(/[^a-z0-9]+/g, ' ').trim();
  if (target) {
    const exact = rowKeys.find((k) => formVTamilNaduHeaderNorm(k).replace(/[^a-z0-9]+/g, ' ').trim() === target);
    if (exact) {
      const picked = pickResolved(coerced[exact]);
      if (picked !== '') return picked;
    }
  }
  if (kind) {
    for (const key of rowKeys) {
      if (formVTamilNaduIdentityColumnKind(key) !== kind) continue;
      const picked = pickResolved(coerced[key]);
      if (picked !== '') return picked;
    }
  }
  if (kind === 'name') {
    for (const key of rowKeys) {
      if (/employee\s*name|^name$|full\s*name/i.test(String(key || ''))) {
        const picked = pickResolved(coerced[key]);
        if (picked !== '') return picked;
      }
    }
  }
  if (kind === 'empid') {
    for (const key of rowKeys) {
      if (/employeeid|employee_id|empcode|emp_id/i.test(String(key || ''))) {
        const picked = pickResolved(coerced[key]);
        if (picked !== '') return picked;
      }
    }
  }
  return '';
}

export function resolveFormVTamilNaduDayValue(row, dayNum, headers = []) {
  if (!row || !dayNum) return '';
  const day = Number(dayNum);
  if (!Number.isFinite(day) || day < 1 || day > 31) return '';
  const pick = (raw) => normalizeFormVTamilNaduAttendanceCode(raw);
  if (Array.isArray(row)) {
    const hdrs = Array.isArray(headers) ? headers : [];
    const idx = hdrs.findIndex((h) => resolveFormVTamilNaduDayNumberFromHeader(h) === day);
    if (idx >= 0) return pick(row[idx]);
    return pick(row[day]) || pick(row[day - 1]) || '';
  }
  if (typeof row !== 'object') return '';
  if (Object.prototype.hasOwnProperty.call(row, String(day))) {
    const v = pick(row[String(day)]);
    if (v) return v;
  }
  if (Object.prototype.hasOwnProperty.call(row, day)) {
    const v = pick(row[day]);
    if (v) return v;
  }
  for (const key of Object.keys(row)) {
    if (String(key).startsWith('__')) continue;
    if (resolveFormVTamilNaduDayNumberFromHeader(key) !== day) continue;
    const v = pick(row[key]);
    if (v) return v;
  }
  return '';
}

/**
 * Locate S.No / Name / Emp ID / shift-time columns on the Form V template header band.
 */
export function locateFormVTamilNaduIdentityExcelColumns(
  worksheet,
  { headerRow = 10, dayRow = -1, firstDayCol = 0, startCol = 1, maxScanCols = 20, cellText = (cell) => String(cell?.value ?? '') } = {}
) {
  const out = { sno: null, name: null, empid: null, commence: null, rest: null, ends: null };
  if (!worksheet) return out;
  const fromRow = Math.max(1, Math.min(headerRow, dayRow > 0 ? dayRow : headerRow) - 1);
  const toRow = Math.max(headerRow, dayRow > 0 ? dayRow : headerRow);
  const lastIdentityCol = firstDayCol > startCol ? firstDayCol - 1 : Math.max(6, startCol + 5);
  const colLimit = Math.min(Number(maxScanCols) || 20, lastIdentityCol);
  for (let c = Math.max(1, startCol); c <= colLimit; c += 1) {
    const parts = [];
    for (let r = fromRow; r <= toRow; r += 1) {
      const piece = String(cellText(worksheet.getCell(r, c)) || '').trim();
      if (piece && !/^\d{1,2}$/.test(piece)) parts.push(piece);
    }
    const kind = formVTamilNaduIdentityColumnKind(parts.join(' '));
    if (kind && out[kind] == null) out[kind] = c;
  }
  if (out.sno == null && startCol > 0) out.sno = startCol;
  if (out.name == null && out.sno > 0) out.name = out.sno + 1;
  return out;
}

const FORM_V_TN_MONTH_NAMES = [
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

function formVTamilNaduOrdinalDay(n) {
  const j = n % 10;
  const k = n % 100;
  if (k >= 11 && k <= 13) return `${n}th`;
  if (j === 1) return `${n}st`;
  if (j === 2) return `${n}nd`;
  if (j === 3) return `${n}rd`;
  return `${n}th`;
}

/**
 * "For the period from 1st April 2026 to 30th April 2026" from selected month/year.
 * Replaces truncated template text like "…April 202…".
 */
export function buildFormVTamilNaduPeriodLine(monthName, year) {
  const idx = FORM_V_TN_MONTH_NAMES.findIndex(
    (m) => m.toLowerCase() === String(monthName || '').toLowerCase().trim()
  );
  if (idx < 0) return '';
  const y = Number(year);
  if (!Number.isFinite(y) || y < 1900) return '';
  const lastDay = new Date(y, idx + 1, 0).getDate();
  const mon = FORM_V_TN_MONTH_NAMES[idx];
  return `For the period from ${formVTamilNaduOrdinalDay(1)} ${mon} ${y} to ${formVTamilNaduOrdinalDay(lastDay)} ${mon} ${y}`;
}

function formVTamilNaduCellText(val) {
  if (val == null) return '';
  if (typeof val === 'string' || typeof val === 'number' || typeof val === 'boolean') return String(val);
  if (typeof val === 'object') {
    if (Array.isArray(val.richText)) return val.richText.map((rt) => rt?.text || '').join('');
    if (val.text != null) return String(val.text);
    if (val.result != null) return String(val.result);
  }
  return String(val ?? '');
}

function isFormVTamilNaduTitleBannerText(raw) {
  const t = String(raw || '').replace(/\s+/g, ' ').trim();
  if (!t) return false;
  if (/^form[\s–—-]*v\.?$/i.test(t)) return true;
  if (/register\s+of\s+employment/i.test(t)) return true;
  if (/see\s+sub-?rule|rule\s*\(?\s*16\s*\)?/i.test(t)) return true;
  if (/for\s+the\s+period\s+from/i.test(t)) return true;
  return false;
}

function centerFormVTamilNaduBannerCell(cell) {
  if (!cell) return;
  const align = {
    ...(cell.alignment || {}),
    horizontal: 'center',
    vertical: 'middle',
    wrapText: true,
  };
  cell.alignment = align;
  try {
    if (cell.style && typeof cell.style === 'object') {
      cell.style = { ...cell.style, alignment: { ...(cell.style.alignment || {}), ...align } };
    }
  } catch (_) {
    /* alignment on cell is enough */
  }
}

/**
 * After day-31 insert the table grows past AP, but title merges stay A1:AP1 — headings then
 * look right-shifted vs Remarks. Re-merge title rows across the full table and force center.
 */
export function centerFormVTamilNaduTitleBannerRows(
  worksheet,
  { headerRowEnd = 8, tableEndCol = 0, periodText = '' } = {}
) {
  if (!worksheet) return { centered: 0, rematched: 0 };
  const rowEnd = Math.max(1, Math.min(Number(headerRowEnd) || 8, 12));
  const periodLine = String(periodText || '').trim();
  let endCol = Number(tableEndCol) || 0;
  if (!(endCol >= 10)) {
    // Fall back to widest existing title merge / last non-empty header cell.
    endCol = 42;
    try {
      const merges = Array.isArray(worksheet.model?.merges) ? worksheet.model.merges : [];
      merges.forEach((m) => {
        const match = String(m || '').match(/^A([1-4]):([A-Z]+)([1-4])$/i);
        if (!match) return;
        const colLetters = match[2].toUpperCase();
        let col = 0;
        for (let i = 0; i < colLetters.length; i += 1) {
          col = col * 26 + (colLetters.charCodeAt(i) - 64);
        }
        if (col > endCol) endCol = col;
      });
    } catch (_) {
      /* keep default */
    }
  }
  endCol = Math.max(10, Math.min(Number(endCol) || 42, 80));

  let centered = 0;
  let rematched = 0;
  for (let r = 1; r <= rowEnd; r += 1) {
    let bannerText = '';
    let foundCol = 0;
    for (let c = 1; c <= Math.max(endCol, 45); c += 1) {
      const raw = formVTamilNaduCellText(worksheet.getCell(r, c)?.value).trim();
      if (!raw) continue;
      if (!isFormVTamilNaduTitleBannerText(raw)) continue;
      bannerText = raw;
      foundCol = c;
      break;
    }
    if (!bannerText) continue;
    if (periodLine && /for\s+the\s+period\s+from/i.test(bannerText)) {
      bannerText = periodLine;
    }

    // Unmerge any existing merges that touch this title row, then merge A..endCol.
    try {
      const merges = Array.isArray(worksheet.model?.merges) ? [...worksheet.model.merges] : [];
      merges.forEach((range) => {
        const m = String(range || '').match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/i);
        if (!m) return;
        const r1 = Number(m[2]);
        const r2 = Number(m[4]);
        if (r < r1 || r > r2) return;
        try {
          worksheet.unMergeCells(range);
        } catch (_) {
          /* ignore */
        }
      });
    } catch (_) {
      /* keep existing merges */
    }

    // Clear non-master leftovers across the band, then write + merge + center.
    for (let c = 1; c <= endCol; c += 1) {
      try {
        const cell = worksheet.getCell(r, c);
        if (c === 1) continue;
        if (formVTamilNaduCellText(cell?.value).trim()) cell.value = null;
      } catch (_) {
        /* skip */
      }
    }
    writeFormVTamilNaduExcelVisibleCell(worksheet, r, 1, bannerText);
    try {
      worksheet.mergeCells(r, 1, r, endCol);
      rematched += 1;
    } catch (_) {
      /* merge may already exist */
    }
    const master = worksheet.getCell(r, 1);
    centerFormVTamilNaduBannerCell(master);
    // Also center any visible merge children ExcelJS exposes.
    for (let c = 1; c <= Math.min(endCol, 5); c += 1) {
      try {
        centerFormVTamilNaduBannerCell(worksheet.getCell(r, c));
      } catch (_) {
        /* skip */
      }
    }
    centered += 1;
    if (foundCol > 1) {
      /* value already moved to A */
    }
  }
  return { centered, rematched, endCol };
}

/** Replace template "For the period from …" banner with the selected month/year sentence (centered). */
export function writeFormVTamilNaduPeriodToWorksheet(worksheet, periodText, headerRowEnd = 12, tableEndCol = 0) {
  if (!worksheet) return false;
  const line = String(periodText || '').trim();
  const { centered } = centerFormVTamilNaduTitleBannerRows(worksheet, {
    headerRowEnd: Math.max(1, Number(headerRowEnd) || 12),
    tableEndCol,
    periodText: line,
  });
  return centered > 0;
}

