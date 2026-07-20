/**
 * Form 25 Tamil Nadu — Excel export column alignment.
 * Prefix + day-number columns (1–31) + trailing summary columns (Total Days / Hours / LOP / …).
 * Same failure mode as Form V: incomplete day headers caused summary values to land under day columns.
 */

export function form25TamilNaduHeaderNorm(txt) {
  return String(txt || '')
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

export function resolveForm25TamilNaduDayNumberFromHeader(header) {
  const h = String(header || '').trim();
  if (/^\d{1,2}$/.test(h)) {
    const n = Number(h);
    return n >= 1 && n <= 31 ? n : 0;
  }
  const m =
    h.match(/(?:dates?|daily\s*hours?|hours?\s+of\s+work)[^0-9]*(\d{1,2})$/i) ||
    h.match(/_(\d{1,2})$/);
  if (m) {
    const n = Number(m[1]);
    return n >= 1 && n <= 31 ? n : 0;
  }
  return 0;
}

export function isForm25TamilNaduDayHeaderKey(header) {
  return resolveForm25TamilNaduDayNumberFromHeader(header) >= 1;
}

/** "Serial Number in Register of adult workers and young persons" (not S.No). */
export function isForm25TamilNaduSerialInRegisterHeader(header) {
  const s = form25TamilNaduHeaderNorm(header);
  if (!s) return false;
  if (/^s\.?\s*no\.?$|^sl\.?\s*no\.?$|^serial$/.test(s)) return false;
  if (/name\s+of\s+the\s+worker/.test(s)) return false;
  const hasSerial = /\bserial\b/.test(s) || (/\bnumber\b/.test(s) && /\bregister\b/.test(s));
  if (!hasSerial || !/\bregister\b/.test(s)) return false;
  return (
    /\badult\b/.test(s) ||
    /\byoung\b/.test(s) ||
    /\bworkers?\b/.test(s) ||
    (/\bserial\b/.test(s) && /\bregister\b/.test(s))
  );
}

/** "Name of the Worker" — EmployeeName (not Serial Number in Register…). */
export function isForm25TamilNaduWorkerNameHeader(header) {
  const s = form25TamilNaduHeaderNorm(header);
  if (!s) return false;
  if (isForm25TamilNaduSerialInRegisterHeader(header)) return false;
  if (/signature|father|husband|address|establishment/.test(s)) return false;
  return /name\s+of\s+the\s+worker/.test(s) || /^name\s+of\s+worker$/.test(s);
}

export function isForm25TamilNaduRemarksHeader(header) {
  return /^remarks?$/.test(form25TamilNaduHeaderNorm(header));
}

export function isForm25TamilNaduSNoHeader(header) {
  const s = form25TamilNaduHeaderNorm(header);
  if (!s) return false;
  if (isForm25TamilNaduSerialInRegisterHeader(header)) return false;
  return (
    /^s\.?\s*no\.?$/.test(s) ||
    /^sl\.?\s*no\.?$/.test(s) ||
    /^sr\.?\s*no\.?$/.test(s) ||
    s === 'sno' ||
    s === 'serial no'
  );
}

/**
 * Ensure prefix headers include Serial Number in Register before Name of the Worker.
 * Parsed grids often drop Serial, so Name is written under the Serial Excel column.
 */
export function ensureForm25TamilNaduPrefixHeaders(headers) {
  const list = Array.isArray(headers) ? [...headers] : [];
  const hasSerial = list.some((h) => isForm25TamilNaduSerialInRegisterHeader(h));
  const nameIdx = list.findIndex((h) => isForm25TamilNaduWorkerNameHeader(h));
  if (hasSerial) return list;
  const serialLabel = 'Serial Number in Register of adult workers and young persons';
  if (nameIdx >= 0) {
    list.splice(nameIdx, 0, serialLabel);
    return list;
  }
  const snoIdx = list.findIndex((h) => isForm25TamilNaduSNoHeader(h));
  if (snoIdx >= 0) {
    list.splice(snoIdx + 1, 0, serialLabel);
    return list;
  }
  list.unshift(serialLabel);
  return list;
}

/** Prefer Zoho EmployeeName, then common name aliases. */
export function getForm25TamilNaduEmployeeName(emp = {}) {
  const pick = (...vals) => {
    for (const v of vals) {
      const s = v != null ? String(v).trim() : '';
      if (s) return s;
    }
    return '';
  };
  return pick(
    emp.EmployeeName,
    emp['EmployeeName'],
    emp['Employee Name'],
    emp.employeeName,
    emp.Nameoftheemployee,
    emp['Name of the employee'],
    emp.Full_Name,
    emp['Full Name'],
    emp.Employee_Name,
    emp['Employee_Name'],
    emp.Name,
    emp.name,
    [emp.FirstName || emp.firstName || emp['First Name'], emp.LastName || emp.lastName || emp['Last Name']]
      .filter(Boolean)
      .join(' ')
      .trim()
  );
}

/**
 * Drop stray columns after Remarks (template sometimes leaves 2 orphan data cols to the right).
 * Keep empty Rest Interval slots — filtering them shifts Ends/Scheme/day-1 left by one.
 */
export function trimForm25TamilNaduHeadersAfterRemarks(headers) {
  const list = Array.isArray(headers) ? [...headers] : [];
  const remarksIdx = list.findIndex((h) => isForm25TamilNaduRemarksHeader(h));
  if (remarksIdx < 0) return list;
  return list.slice(0, remarksIdx + 1);
}

/** When parsed headers omit day columns, synthesize 1…31 before summary fields. */
export function ensureForm25TamilNaduDayColumnHeaders(headers) {
  // Keep blank prefix headers (e.g. Rest Interval) so column indexes stay aligned with Excel.
  const base = Array.isArray(headers) ? [...headers] : [];
  while (base.length > 0 && !String(base[base.length - 1] || '').trim()) base.pop();
  const dayByNum = new Map();
  const prefix = [];
  const suffix = [];
  let seenDay = false;
  base.forEach((header) => {
    const day = resolveForm25TamilNaduDayNumberFromHeader(header);
    if (day >= 1 && day <= 31) {
      seenDay = true;
      if (!dayByNum.has(day)) dayByNum.set(day, header);
      return;
    }
    if (!seenDay) prefix.push(header);
    else suffix.push(header);
  });
  const trimmedSuffix = trimForm25TamilNaduHeadersAfterRemarks(suffix);
  if (dayByNum.size >= 28) {
    const days = [...dayByNum.entries()]
      .sort((a, b) => a[0] - b[0])
      .map(([, header]) => header);
    return [...prefix, ...days, ...trimmedSuffix];
  }
  const joined = [...prefix, ...trimmedSuffix].join(' ').toLowerCase();
  const looksLikeForm25 =
    /scheme\s+of\s+shift/.test(joined) ||
    /register\s+of\s+adult/.test(joined) ||
    /worker\s+identif/.test(joined) ||
    /hours\s+of\s+work/.test(joined) ||
    /time\s+at\s+which\s+work/.test(joined) ||
    dayByNum.size >= 3;
  if (!looksLikeForm25 && dayByNum.size === 0) return trimForm25TamilNaduHeadersAfterRemarks(base);
  for (let day = 1; day <= 31; day += 1) {
    if (!dayByNum.has(day)) dayByNum.set(day, String(day));
  }
  const days = [...dayByNum.entries()]
    .sort((a, b) => a[0] - b[0])
    .map(([, header]) => header);
  return [...prefix, ...days, ...trimmedSuffix];
}

/** Merge template/modal headers with day keys present on export rows. */
export function resolveForm25TamilNaduExportHeaders(headersToUse, mappedData) {
  const base = Array.isArray(headersToUse) ? [...headersToUse] : [];
  while (base.length > 0 && !String(base[base.length - 1] || '').trim()) base.pop();
  return ensureForm25TamilNaduDayColumnHeaders(ensureForm25TamilNaduPrefixHeaders(base));
}

/**
 * Locate Form 25 TN S.No / Serial-in-Register / Name columns from the template header band.
 */
export function locateForm25TamilNaduPrefixExcelColumns(
  worksheet,
  {
    headerRow = 9,
    startCol = 1,
    firstDayCol = -1,
    maxScanCols = 40,
    cellText = (cell) => String(cell?.value ?? ''),
  } = {}
) {
  const out = { sno: null, serialInRegister: null, workerName: null };
  if (!worksheet) return out;
  const fromRow = Math.max(1, headerRow - 2);
  const toRow = headerRow + 3;
  const endCol =
    firstDayCol > startCol ? firstDayCol - 1 : Math.min(Number(maxScanCols) || 40, startCol + 12);

  for (let c = startCol; c <= endCol; c += 1) {
    const parts = [];
    for (let r = fromRow; r <= toRow; r += 1) {
      const piece = String(cellText(worksheet.getCell(r, c)) || '').trim();
      if (piece && !/^\d{1,2}$/.test(piece)) parts.push(piece);
    }
    const joined = parts.join(' ');
    if (!out.sno && isForm25TamilNaduSNoHeader(joined)) out.sno = c;
    if (!out.serialInRegister && isForm25TamilNaduSerialInRegisterHeader(joined)) {
      out.serialInRegister = c;
    }
    if (!out.workerName && isForm25TamilNaduWorkerNameHeader(joined)) out.workerName = c;
  }

  // Canonical Form 25 TN layout: A=S.No, B=Serial-in-Register, C=Name
  if (!out.sno) out.sno = startCol;
  if (!out.serialInRegister) out.serialInRegister = startCol + 1;
  if (!out.workerName) out.workerName = startCol + 2;
  return out;
}

/**
 * Pull a worker name from a row even when it was stored under Serial Number in Register.
 */
export function extractForm25TamilNaduWorkerNameFromRow(row, headers = []) {
  if (!row) return '';
  const hdrs = Array.isArray(headers) ? headers : [];
  const nameHdr = hdrs.find((h) => isForm25TamilNaduWorkerNameHeader(h));
  const serialHdr = hdrs.find((h) => isForm25TamilNaduSerialInRegisterHeader(h));

  if (Array.isArray(row)) {
    const nameIdx = nameHdr ? hdrs.indexOf(nameHdr) : 2;
    const serialIdx = serialHdr ? hdrs.indexOf(serialHdr) : 1;
    const atName = nameIdx >= 0 ? String(row[nameIdx] ?? '').trim() : '';
    if (atName && !/^\d+$/.test(atName) && /[a-zA-Z]{2,}/.test(atName)) return atName;
    const atSerial = serialIdx >= 0 ? String(row[serialIdx] ?? '').trim() : '';
    if (atSerial && !/^\d+$/.test(atSerial) && /[a-zA-Z]{2,}/.test(atSerial)) return atSerial;
    // Common mis-aligned shapes: [sno, name, ...] or [sno, name, id, ...]
    for (const idx of [2, 1, 3]) {
      const v = String(row[idx] ?? '').trim();
      if (v && !/^\d+$/.test(v) && /[a-zA-Z]{2,}/.test(v)) return v;
    }
    return '';
  }

  if (typeof row !== 'object') return '';
  if (nameHdr) {
    const v = String(row[nameHdr] ?? '').trim();
    if (v && !/^\d+$/.test(v)) return v;
  }
  for (const [key, val] of Object.entries(row)) {
    if (String(key).startsWith('__')) continue;
    if (!isForm25TamilNaduWorkerNameHeader(key)) continue;
    const v = String(val ?? '').trim();
    if (v && !/^\d+$/.test(v)) return v;
  }
  if (serialHdr) {
    const v = String(row[serialHdr] ?? '').trim();
    if (v && !/^\d+$/.test(v) && /[a-zA-Z]{2,}/.test(v)) return v;
  }
  for (const [key, val] of Object.entries(row)) {
    if (String(key).startsWith('__')) continue;
    if (!isForm25TamilNaduSerialInRegisterHeader(key)) continue;
    const v = String(val ?? '').trim();
    if (v && !/^\d+$/.test(v) && /[a-zA-Z]{2,}/.test(v)) return v;
  }
  // Last resort: any person-like value under a name-ish key
  for (const [key, val] of Object.entries(row)) {
    if (String(key).startsWith('__')) continue;
    const nk = form25TamilNaduHeaderNorm(key);
    if (!/\bname\b/.test(nk) || /signature|father|husband|scheme|factory/.test(nk)) continue;
    const v = String(val ?? '').trim();
    if (v && !/^\d+$/.test(v) && /[a-zA-Z]{2,}/.test(v)) return v;
  }
  return '';
}

/**
 * Force Form 25 TN prefix columns:
 * - Serial Number in Register → 1, 2, 3…
 * - Name of the Worker → EmployeeName
 * Clears any values that belonged to columns after Remarks.
 */
export function applyForm25TamilNaduPrefixColumnsToRows(rows, headers, employees = []) {
  if (!Array.isArray(rows) || !Array.isArray(headers) || headers.length === 0) return rows;
  const originalHeaders = [...headers];
  const ensuredHeaders = ensureForm25TamilNaduPrefixHeaders(headers);
  const serialHdr = ensuredHeaders.find((h) => isForm25TamilNaduSerialInRegisterHeader(h));
  const nameHdr = ensuredHeaders.find((h) => isForm25TamilNaduWorkerNameHeader(h));
  const remarksIdx = ensuredHeaders.findIndex((h) => isForm25TamilNaduRemarksHeader(h));
  const empList = Array.isArray(employees) ? employees : [];
  const serialInserted =
    !originalHeaders.some((h) => isForm25TamilNaduSerialInRegisterHeader(h)) && Boolean(serialHdr);
  const serialInsertAt = serialInserted && serialHdr ? ensuredHeaders.indexOf(serialHdr) : -1;

  return rows.map((row, idx) => {
    const n = String(idx + 1);
    const emp = empList[idx] || null;
    const fromEmp = emp ? getForm25TamilNaduEmployeeName(emp) : '';

    if (Array.isArray(row)) {
      const out = [...row];
      // Keep cell positions aligned when Serial header was injected into the header list.
      if (serialInsertAt >= 0) out.splice(serialInsertAt, 0, '');
      while (out.length < ensuredHeaders.length) out.push('');
      const nameIdx = nameHdr ? ensuredHeaders.indexOf(nameHdr) : -1;
      const serialIdx = serialHdr ? ensuredHeaders.indexOf(serialHdr) : -1;
      const recovered = extractForm25TamilNaduWorkerNameFromRow(row, originalHeaders);
      const nameVal = fromEmp || recovered;
      if (nameIdx >= 0) out[nameIdx] = nameVal;
      if (serialIdx >= 0) out[serialIdx] = n;
      if (remarksIdx >= 0) {
        for (let i = remarksIdx + 1; i < out.length; i += 1) out[i] = '';
      }
      return out;
    }

    if (!row || typeof row !== 'object') return row;
    const out = { ...row };
    const recovered = extractForm25TamilNaduWorkerNameFromRow(row, originalHeaders);
    const nameVal = fromEmp || recovered;

    if (nameHdr) out[nameHdr] = nameVal || String(out[nameHdr] ?? '').trim();
    if (serialHdr) out[serialHdr] = n;

    // Clear stale name left under any serial-like key variants.
    Object.keys(out).forEach((key) => {
      if (key === serialHdr) return;
      if (isForm25TamilNaduSerialInRegisterHeader(key)) out[key] = n;
    });

    if (remarksIdx >= 0) {
      ensuredHeaders.forEach((h, hi) => {
        if (hi > remarksIdx && h) out[h] = '';
      });
    }
    return out;
  });
}

/**
 * Locate Form 25 TN trailing summary / signature columns by header text (after day band).
 */
export function locateForm25TamilNaduSummaryExcelColumns(
  worksheet,
  {
    headerRow = 9,
    dayRow = -1,
    afterCol = 0,
    maxScanCols = 80,
    cellText = (cell) => String(cell?.value ?? ''),
  } = {}
) {
  const norm = form25TamilNaduHeaderNorm;
  const out = {
    totalDaysWorked: null,
    totalHoursWorked: null,
    lossOfPay: null,
    nationalHolidayBenefit: null,
    festivalHolidayBenefit: null,
    leaveWithWages: null,
    casualLeave: null,
    signatureWorker: null,
    signatureManager: null,
    remarks: null,
  };
  if (!worksheet) return out;
  const fromRow = Math.max(1, Math.min(headerRow, dayRow > 0 ? dayRow : headerRow) - 5);
  const toRow = Math.max(headerRow, dayRow > 0 ? dayRow : headerRow) + 3;
  const startCol = Math.max(1, Number(afterCol) || 1);

  const classifyJoined = (joined, c) => {
    const t = norm(joined);
    if (!t) return;
    if (/approved\s+festival|approval\s+proceedings/.test(t)) return;
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
    if (
      !out.nationalHolidayBenefit &&
      /national/.test(t) &&
      /holiday/.test(t) &&
      (/benefit/.test(t) || /availed/.test(t) || /working/.test(t) || /entitled/.test(t))
    ) {
      out.nationalHolidayBenefit = c;
      return;
    }
    if (
      !out.festivalHolidayBenefit &&
      /festival/.test(t) &&
      /holiday/.test(t) &&
      (/benefit/.test(t) || /availed/.test(t) || /working/.test(t) || /entitled/.test(t))
    ) {
      out.festivalHolidayBenefit = c;
      return;
    }
    if (!out.leaveWithWages && /leave/.test(t) && /wages?/.test(t)) {
      out.leaveWithWages = c;
      return;
    }
    if (!out.casualLeave && /casual\s+leave/.test(t)) {
      out.casualLeave = c;
      return;
    }
    if (!out.signatureWorker && /signature/.test(t) && /worker/.test(t)) {
      out.signatureWorker = c;
      return;
    }
    if (!out.signatureManager && /signature/.test(t) && /manager/.test(t)) {
      out.signatureManager = c;
      return;
    }
    if (!out.remarks && /^remarks?$/.test(t)) {
      out.remarks = c;
    }
  };

  for (let c = startCol; c <= maxScanCols; c += 1) {
    const parts = [];
    for (let r = fromRow; r <= toRow; r += 1) {
      const piece = String(cellText(worksheet.getCell(r, c)) || '').trim();
      if (piece) parts.push(piece);
    }
    classifyJoined(parts.join(' '), c);
  }

  for (let r = fromRow; r <= toRow; r += 1) {
    for (let c = startCol; c <= maxScanCols; c += 1) {
      classifyJoined(cellText(worksheet.getCell(r, c)), c);
    }
  }

  // Template order after LOP when labels are split / missing.
  if (out.lossOfPay > 0) {
    if (!out.nationalHolidayBenefit) out.nationalHolidayBenefit = out.lossOfPay + 1;
    if (!out.festivalHolidayBenefit) out.festivalHolidayBenefit = out.lossOfPay + 2;
    if (!out.leaveWithWages) out.leaveWithWages = out.lossOfPay + 3;
    if (!out.casualLeave) out.casualLeave = out.lossOfPay + 4;
    if (!out.signatureWorker) out.signatureWorker = out.lossOfPay + 5;
    if (!out.signatureManager) out.signatureManager = out.lossOfPay + 6;
    if (!out.remarks) out.remarks = out.lossOfPay + 7;
  } else if (out.totalDaysWorked > 0) {
    if (!out.totalHoursWorked) out.totalHoursWorked = out.totalDaysWorked + 1;
    if (!out.lossOfPay) out.lossOfPay = out.totalDaysWorked + 2;
  }

  return out;
}

export function form25TamilNaduSummaryColForHeader(header, summaryCols) {
  const s = form25TamilNaduHeaderNorm(header);
  if (!s || !summaryCols) return null;
  if (/total/.test(s) && /days?/.test(s) && /work/.test(s)) return summaryCols.totalDaysWorked;
  if (/total/.test(s) && /hours?/.test(s) && /work/.test(s)) return summaryCols.totalHoursWorked;
  if (/loss/.test(s) && /pay/.test(s)) return summaryCols.lossOfPay;
  if (
    /national/.test(s) &&
    /holiday/.test(s) &&
    (/benefit/.test(s) || /availed/.test(s) || /working/.test(s) || /entitled/.test(s))
  ) {
    return summaryCols.nationalHolidayBenefit;
  }
  if (
    /festival/.test(s) &&
    /holiday/.test(s) &&
    (/benefit/.test(s) || /availed/.test(s) || /working/.test(s) || /entitled/.test(s))
  ) {
    return summaryCols.festivalHolidayBenefit;
  }
  if (/leave/.test(s) && /wages?/.test(s)) return summaryCols.leaveWithWages;
  if (/casual\s+leave/.test(s)) return summaryCols.casualLeave;
  if (/signature/.test(s) && /worker/.test(s)) return summaryCols.signatureWorker;
  if (/signature/.test(s) && /manager/.test(s)) return summaryCols.signatureManager;
  if (/^remarks?$/.test(s)) return summaryCols.remarks;
  return null;
}

/**
 * Map each Form 25 header to the physical Excel column.
 * Days → day-number columns; trailing summary → label columns after last day.
 * Never invent a day column past the physical day band (that overwrote Total Days Worked).
 */
export function buildForm25TamilNaduOrderedExcelCols({
  sourceHeaders,
  startCol = 1,
  dayStartIdx = -1,
  lastDayHeaderIdx = -1,
  firstDayCol = -1,
  lastDayCol = -1,
  dayColsByNumber = new Map(),
  markerCols = [],
  summaryCols = {},
  daysInMonth = 31,
} = {}) {
  const headers = Array.isArray(sourceHeaders) ? sourceHeaders : [];
  const monthDays = Math.min(Math.max(Number(daysInMonth) || 31, 28), 31);
  return headers.map((h, idx) => {
    const day = resolveForm25TamilNaduDayNumberFromHeader(h);
    if (day >= 1 && day <= 31) {
      if (dayColsByNumber.has(day)) return dayColsByNumber.get(day);
      if (day <= monthDays && markerCols[day - 1]) return markerCols[day - 1];
      return null;
    }
    const summaryCol = form25TamilNaduSummaryColForHeader(h, summaryCols);
    if (summaryCol) return summaryCol;
    if (dayStartIdx >= 0 && idx > lastDayHeaderIdx && lastDayCol > 0) {
      return lastDayCol + (idx - lastDayHeaderIdx);
    }
    if (dayStartIdx >= 0 && idx < dayStartIdx && firstDayCol > startCol) {
      const slots = firstDayCol - startCol;
      return startCol + Math.min(idx, Math.max(slots, 1) - 1);
    }
    return startCol + idx;
  });
}

const FORM25_TN_ALIAS_BUCKETS = [
  {
    re: /^(s\.?\s*no\.?|sl\.?\s*no\.?|serial)$/i,
    keys: [/^(s\.?\s*no\.?|sl\.?\s*no\.?)$/i, /^serial$/i],
  },
  {
    re: /serial\s+number\s+in\s+register|register\s+of\s+adult/i,
    keys: [/serial\s+number\s+in\s+register/i, /register\s+of\s+adult/i],
  },
  {
    re: /name\s+of\s+the\s+worker|^name$/i,
    keys: [/name\s+of\s+the\s+worker/i, /^name$/i],
  },
  {
    re: /worker\s+identif|employee\s*id|emp\.?\s*id/i,
    keys: [/worker\s+identif/i, /employee\s*id/i, /emp\.?\s*id/i],
  },
  {
    re: /commence|begins|start/i,
    keys: [/commence/i, /begins/i, /start/i],
  },
  {
    re: /rest\s+interval|^rest$/i,
    keys: [/rest\s+interval/i, /^rest$/i],
  },
  {
    re: /time\s+at\s+which\s+work\s+end|work\s+ends/i,
    keys: [/time\s+at\s+which\s+work\s+end/i, /work\s+ends/i],
  },
  {
    re: /scheme\s+of\s+shift/i,
    keys: [/scheme\s+of\s+shift/i],
  },
  {
    re: /total\s*days?\s*worked/i,
    keys: [/total\s*days?\s*worked/i],
  },
  {
    re: /total\s*hours?\s*worked/i,
    keys: [/total\s*hours?\s*worked/i],
  },
  {
    re: /loss\s*of\s*pay/i,
    keys: [/loss\s*of\s*pay/i],
  },
];

/**
 * Resolve a Form 25 TN cell value by header — never use Object.entries index fallback
 * (that put Names under Serial Number and Scheme under Work Ends).
 */
export function resolveForm25TamilNaduExportCellValue(rowObj, header, { pickExportCell } = {}) {
  const pick =
    typeof pickExportCell === 'function'
      ? pickExportCell
      : (v) => (v == null || String(v).trim() === '' ? '' : v);
  if (!rowObj || typeof rowObj !== 'object' || Array.isArray(rowObj)) return '';
  if (header && Object.prototype.hasOwnProperty.call(rowObj, header)) {
    const direct = pick(rowObj[header]);
    if (direct !== '') return direct;
  }
  const target = form25TamilNaduHeaderNorm(String(header || '').replace(/[^a-z0-9]+/gi, ' '));
  if (!target) return '';

  // Manual / signature / benefit columns — do not fuzzy-fill from unrelated People fields.
  if (
    /signature/.test(target) ||
    /^remarks?$/.test(target) ||
    (/benefit/.test(target) && /holiday/.test(target)) ||
    (/leave/.test(target) && /wages?/.test(target)) ||
    /casual\s+leave/.test(target)
  ) {
    return '';
  }

  const rowKeys = Object.keys(rowObj).filter((k) => !String(k).startsWith('__'));
  const exact = rowKeys.find(
    (k) => form25TamilNaduHeaderNorm(String(k).replace(/[^a-z0-9]+/gi, ' ')) === target
  );
  if (exact) {
    const v = pick(rowObj[exact]);
    if (v !== '') return v;
  }

  for (const bucket of FORM25_TN_ALIAS_BUCKETS) {
    if (!bucket.re.test(String(header || ''))) continue;
    for (const key of rowKeys) {
      if (!bucket.keys.some((rk) => rk.test(String(key || '')))) continue;
      if (isForm25TamilNaduDayHeaderKey(key)) continue;
      const v = pick(rowObj[key]);
      if (v !== '') return v;
    }
  }
  return '';
}

/** Widen day / summary columns so WO / totals / long IDs do not visually collide. */
export function applyForm25TamilNaduExportColumnWidths(
  worksheet,
  { dayCols = [], summaryCols = {}, startCol = 1, prefixCount = 8 } = {}
) {
  if (!worksheet) return;
  const setWidth = (col, width) => {
    if (!Number.isFinite(col) || col < 1) return;
    try {
      const column = worksheet.getColumn(col);
      const cur = Number(column.width) || 0;
      if (cur < width) column.width = width;
    } catch (_) {
      /* ignore */
    }
  };
  for (let i = 0; i < prefixCount; i += 1) {
    const w = i === 0 ? 6 : i <= 2 ? 18 : i === 3 ? 14 : 12;
    setWidth(startCol + i, w);
  }
  (Array.isArray(dayCols) ? dayCols : []).forEach((col) => setWidth(col, 4.5));
  Object.values(summaryCols || {}).forEach((col) => {
    if (Number.isFinite(col)) setWidth(col, 12);
  });
  if (summaryCols.signatureWorker) setWidth(summaryCols.signatureWorker, 16);
  if (summaryCols.signatureManager) setWidth(summaryCols.signatureManager, 16);
  if (summaryCols.remarks) setWidth(summaryCols.remarks, 14);
}
