/**
 * Tamil Nadu Form V — Register of Employment (daily attendance columns).
 * Day columns must match the selected month (May → 31, April → 30), not a fixed 30.
 * Total Days Worked ← payroll Paid_days.
 */

import { flattenPayrollEarningColumns, readPayrollScalar } from '../../utils/payrollEarnings';


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
  if (fromFlat !== '') return String(fromFlat);
  if (payrollRow !== flat) {
    const fromRow = readPayrollScalar(payrollRow, keys, patterns);
    if (fromRow !== '') return String(fromRow);
  }
  return '';
}

/** Total Hours Worked = Paid_days × 8. */
export function computeFormVTamilNaduTotalHoursWorked(paidDays) {
  const n = Number(String(paidDays ?? '').replace(/,/g, '').trim());
  if (!Number.isFinite(n) || n < 0) return '';
  return String(Math.round(n * 8 * 100) / 100);
}

/** Loss of Pay days = monthDays − Paid_days (never negative). */
export function computeFormVTamilNaduLossOfPayDays(paidDays, daysInMonth = 31) {
  const paid = Number(String(paidDays ?? '').replace(/,/g, '').trim());
  const monthDays = Math.min(Math.max(Number(daysInMonth) || 31, 28), 31);
  if (!Number.isFinite(paid)) return '';
  return String(Math.max(0, monthDays - paid));
}

/**
 * Fill Form V summary columns from payroll Paid_days:
 * - Total Days Worked ← Paid_days
 * - Total Hours Worked ← Paid_days × 8
 * - Loss of Pay ← monthDays − Paid_days
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
    if (paidDays === '') return;
    let changed = false;
    if (daysHeader && (overwrite || String(row[daysHeader] ?? '').trim() === '')) {
      row[daysHeader] = sanitize(paidDays);
      changed = true;
    }
    if (hoursHeader && (overwrite || String(row[hoursHeader] ?? '').trim() === '')) {
      const hours = computeFormVTamilNaduTotalHoursWorked(paidDays);
      if (hours !== '') {
        row[hoursHeader] = sanitize(hours);
        changed = true;
      }
    }
    if (lopHeader && (overwrite || String(row[lopHeader] ?? '').trim() === '')) {
      const lop = computeFormVTamilNaduLossOfPayDays(paidDays, daysInMonth);
      if (lop !== '') {
        row[lopHeader] = sanitize(lop);
        changed = true;
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
