import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import { headersIndicateFormAGJEmployeeRegisterTable } from './formCGJGujarat';
import { ensureExcelJSDataRowsWithBorders } from '../../utils/excelTableBorders';
import {
  enrichEstablishmentPrincipalEmployerHeaderFields,
  excelCellValueToString,
  writeStatutoryHeaderFieldsToExcelJsWorksheet,
} from '../../utils/statutorySiteCompanyHeaders';

/** Gujarat Form D — Muster Roll (Relay or Set Work, Summary days/hours). */

export const FORM_DGJ_GJ_CANONICAL_TABLE_HEADERS = [
  'Sr. No. in Employee / Workman / Worker Register',
  'Name',
  'Relay or Set Work',
  'Summary No. of Days',
  'Remarks No. of Hours',
  'Signature of Register Keeper*',
];

export function formDGJGujaratHeaderNorm(txt) {
  return String(txt || '')
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

export function blobIndicatesFormDGJMusterRoll(blob, tableHeaders = null) {
  const text = String(blob || '').toLowerCase();
  if (/\bform[\s._-]*d\b/.test(text) && /gujarat|\b_gj\b|form_d_gj/.test(text)) return true;
  if (/relay\s+or\s+set\s+work/.test(text)) return true;
  if (/summary\s+no\.?\s*of\s+days/.test(text) && /register\s+keeper/.test(text)) return true;
  if (Array.isArray(tableHeaders) && tableHeaders.length > 0) {
    const hdr = tableHeaders.map((h) => formDGJGujaratHeaderNorm(h)).join('\n');
    if (/relay\s+or\s+set/.test(hdr)) return true;
    if (/summary.*days/.test(hdr) && /register\s+keeper|signature/.test(hdr)) return true;
    if (/sr\.?\s*no.*register/.test(hdr) && /relay\s+or\s+set/.test(hdr)) return true;
  }
  return false;
}

export function headersIndicateFormDGJMusterTable(tableHeaders) {
  if (!Array.isArray(tableHeaders) || tableHeaders.length < 3) return false;
  const joined = tableHeaders.map((h) => formDGJGujaratHeaderNorm(h)).join('\n');
  const hasRelay = /relay\s+or\s+set/.test(joined);
  const hasSummaryDays = /summary.*days|no\.?\s*of\s+days/.test(joined);
  const hasName = /\bname\b/.test(joined);
  const hasRegisterKeeper = /register\s+keeper|signature/.test(joined);
  const hasSrRegister =
    /sr\.?\s*no/.test(joined) &&
    (/employee.*register|workman.*register|worker.*register/.test(joined) ||
      /register/.test(joined));
  if (hasRelay && hasSummaryDays) return true;
  if (hasRelay && hasName && hasRegisterKeeper) return true;
  if (hasSrRegister && hasRelay) return true;
  return false;
}

export function isFormDGJGujaratContext(
  formHeader,
  rowItem,
  fileName,
  sheetText = '',
  tableHeaders = []
) {
  const parts = [
    rowItem?.formName,
    rowItem?.FormName,
    rowItem?.formFileName,
    rowItem?.FormFileName,
    fileName,
    formHeader?.title,
    formHeader?.subtitle,
    formHeader?.reference,
    sheetText,
    Array.isArray(tableHeaders) ? tableHeaders.join(' ') : '',
  ]
    .filter((x) => x != null && String(x).trim() !== '')
    .join(' ')
    .toLowerCase();

  const hasGujarat = /gujarat|\b_gj\b|form[\s._-]*d[\s._-]*gj|form_d_gj/.test(parts);
  const hasFormD = /\bform[\s._-]*d\b/.test(parts);

  if (hasGujarat && hasFormD) return true;
  if (hasFormD && blobIndicatesFormDGJMusterRoll(parts, tableHeaders)) return true;
  if (hasGujarat && blobIndicatesFormDGJMusterRoll(parts, tableHeaders)) return true;
  if (headersIndicateFormDGJMusterTable(tableHeaders) && (hasGujarat || hasFormD)) return true;
  return false;
}

export function isFormDGJGujaratUsingWrongEmployeeRegisterTable(tableHeaders) {
  return (
    headersIndicateFormAGJEmployeeRegisterTable(tableHeaders) &&
    !headersIndicateFormDGJMusterTable(tableHeaders)
  );
}

function buildSheetTextBlob(workbook, sheetName) {
  const ws = workbook?.Sheets?.[sheetName];
  if (!ws) return '';
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  return rows
    .slice(0, 45)
    .map((row) => (Array.isArray(row) ? row : []).map((c) => String(c || '')).join(' '))
    .join(' ')
    .toLowerCase();
}

function scoreFormDGJGujaratSheet(sheetName, sheetText, hintsBlob) {
  const sheetLower = String(sheetName || '').toLowerCase().replace(/\s+/g, ' ').trim();
  const sheetBlob = `${sheetLower} ${sheetText}`;
  let score = 0;

  if (/\bform[\s._-]*d\b/.test(sheetLower) && !/\bform[\s._-]*d[\s._-]*(a|b|c|e|f)\b/.test(sheetLower)) {
    score += 240;
  }
  if (/\bform\s*d\b/.test(sheetText) && /relay\s+or\s+set/.test(sheetText)) score += 200;
  if (/relay\s+or\s+set\s+work/.test(sheetText)) score += 90;
  if (/summary\s+no\.?\s*of\s+days/.test(sheetText)) score += 70;
  if (/register\s+keeper/.test(sheetText)) score += 50;
  if (/sr\.?\s*no.*register/.test(sheetText)) score += 40;
  if (/remarks.*hours|no\.?\s*of\s+hours/.test(sheetText)) score += 30;

  if (/surname/.test(sheetText) && /\bgender\b/.test(sheetText) && /education/.test(sheetText)) {
    score -= 150;
  }
  if (/format\s+of\s+employee|type\s+of\s+employment/.test(sheetText) && /education/.test(sheetText)) {
    score -= 120;
  }
  if (/\bform[\s._-]*a\b/.test(sheetLower) && hintsBlob.includes('form') && hintsBlob.includes('d')) {
    score -= 180;
  }
  if (/\bform[\s._-]*b\b/.test(sheetLower) && /register\s+of\s+wages|rate\s+of\s+wage/.test(sheetText)) {
    score -= 100;
  }
  if (/\bform[\s._-]*c\b/.test(sheetLower) && /register\s+of\s+deductions|damage/.test(sheetText)) {
    score -= 100;
  }
  if (/gujarat|\b_gj\b|form_d_gj/.test(hintsBlob) && /gujarat|\b_gj\b/.test(sheetBlob)) score += 20;

  return score;
}

export function resolveFormDGJGujaratWorkbookSheetName(workbook, hints = {}) {
  const names = Array.isArray(workbook?.SheetNames) ? workbook.SheetNames : [];
  if (!names.length) return null;
  if (names.length === 1) return names[0];

  const blob = [
    hints.fileName,
    hints.formFileName,
    hints.formName,
    hints.item?.formName,
    hints.item?.FormName,
    hints.formHeaderTitle,
    hints.formHeader?.title,
  ]
    .filter(Boolean)
    .join(' ')
    .toLowerCase();

  const wantsFormD =
    /\bform[\s._-]*d\b/.test(blob) ||
    /form[\s._-]*d[\s._-]*gj|form_d_gj/.test(blob) ||
    blobIndicatesFormDGJMusterRoll(blob);

  if (!wantsFormD) return hints.preferredSheetName || names[0];

  let best = hints.preferredSheetName && names.includes(hints.preferredSheetName)
    ? hints.preferredSheetName
    : names[0];
  let bestScore = scoreFormDGJGujaratSheet(best, buildSheetTextBlob(workbook, best), blob);

  for (const name of names) {
    const score = scoreFormDGJGujaratSheet(name, buildSheetTextBlob(workbook, name), blob);
    if (score > bestScore) {
      bestScore = score;
      best = name;
    }
  }
  return bestScore > 0 ? best : hints.preferredSheetName || names[0];
}

function getSheetMergedCellText(rows, merges, r, c) {
  const direct = String((rows[r] || [])[c] ?? '').trim();
  if (direct) return direct;
  for (const m of merges || []) {
    if (!m?.s || !m?.e) continue;
    if (r >= m.s.r && r <= m.e.r && c >= m.s.c && c <= m.e.c) {
      return String((rows[m.s.r] || [])[m.s.c] ?? '').trim();
    }
  }
  return '';
}

export function resolveFormDGJGujaratTableLayout(workbook, hints = {}) {
  const names = Array.isArray(workbook?.SheetNames) ? workbook.SheetNames : [];
  if (!names.length) return null;

  const sheetName = resolveFormDGJGujaratWorkbookSheetName(workbook, hints);
  if (!sheetName) return null;

  const ws = workbook?.Sheets?.[sheetName];
  if (!ws) return null;

  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  const merges = ws['!merges'] || [];
  const maxScan = Math.min(rows.length, 90);
  let best = null;
  let bestScore = -1;

  for (let r = 0; r < maxScan; r += 1) {
    const cells = [];
    const maxCols = Math.max((rows[r] || []).length, (rows[r + 1] || []).length, 16);
    for (let c = 0; c < maxCols; c += 1) {
      const text = getSheetMergedCellText(rows, merges, r, c);
      if (text) cells.push({ c, text, lower: text.toLowerCase() });
    }
    if (cells.length < 2) continue;

    const nextRowCells = [];
    for (let c = 0; c < maxCols; c += 1) {
      const text = getSheetMergedCellText(rows, merges, r + 1, c);
      if (text) nextRowCells.push(text.toLowerCase());
    }
    const rowText = cells.map((x) => x.lower).join(' ');
    const combinedText = `${rowText} ${nextRowCells.join(' ')}`.trim();

    let score = 0;
    if (/relay\s+or\s+set/.test(combinedText)) score += 80;
    if (/summary\s+no\.?\s*of\s+days/.test(combinedText)) score += 60;
    if (/register\s+keeper|signature/.test(combinedText)) score += 40;
    if (/remarks.*hours|no\.?\s*of\s+hours/.test(combinedText)) score += 30;
    if (/sr\.?\s*no.*register/.test(combinedText)) score += 35;
    if (/\bname\b/.test(combinedText) && !/surname/.test(combinedText)) score += 20;
    if (/s\.?\s*no|serial|sl\.?\s*no/.test(combinedText)) score += 15;

    if (/surname/.test(combinedText) && /\bgender\b/.test(combinedText) && /education/.test(combinedText)) {
      score -= 150;
    }
    if (/employees?\s*\/\s*workme|worker\s+code/.test(combinedText) && /surname/.test(combinedText)) {
      score -= 120;
    }
    if (/register\s+of\s+wages|rate\s+of\s+wage/.test(combinedText)) score -= 80;
    if (/register\s+of\s+deductions|damage\s+or\s+loss/.test(combinedText)) score -= 80;

    score += cells.length * 2;
    if (score < 15) continue;

    let startCol = cells[0].c;
    for (const cell of cells) {
      if (/^s\.?\s*no\.?$|^sl\.?\s*no\.?$|serial/i.test(cell.lower)) {
        startCol = cell.c;
        break;
      }
    }

    const headers = [];
    for (let c = startCol; c < maxCols; c += 1) {
      let h = getSheetMergedCellText(rows, merges, r, c);
      if (!h) h = getSheetMergedCellText(rows, merges, r + 1, c);
      if (!h && r > 0) h = getSheetMergedCellText(rows, merges, r - 1, c);
      h = String(h || '').replace(/\s+/g, ' ').trim();
      if (!h) {
        if (headers.length > 0) break;
        continue;
      }
      if (/^\d+$/.test(h) && headers.length === 0) continue;
      headers.push(h);
    }
    if (headers.length < 3) continue;

    let headerRowIndex = r;
    let dataStartIndex = r + 1;
    const probeRow = rows[dataStartIndex] || [];
    if (
      headers.some((h) => /sr\.?\s*no|serial|sl\.?\s*no/i.test(String(h))) &&
      probeRow.length > 0 &&
      !/^\d+$/.test(String(probeRow[startCol] ?? '').trim())
    ) {
      dataStartIndex = r + 2;
    }

    const candidate = {
      sheetName,
      headerRowIndex,
      dataStartIndex,
      tableStartCol: startCol,
      headers,
    };
    if (score > bestScore) {
      bestScore = score;
      best = candidate;
    }
  }

  return best;
}

export function resolveFormDGJGujaratTableHeaders(tableHeaders) {
  const parsed = Array.isArray(tableHeaders)
    ? tableHeaders.map((h) => String(h || '').trim()).filter(Boolean)
    : [];
  if (parsed.length >= 3 && headersIndicateFormDGJMusterTable(parsed)) return parsed;
  return parsed.length > 0 ? parsed : [...FORM_DGJ_GJ_CANONICAL_TABLE_HEADERS];
}

export function repairFormDGJGujaratTableHeadersFromWorkbook(workbook, hints = {}) {
  if (!workbook) return null;
  const layout = resolveFormDGJGujaratTableLayout(workbook, {
    preferredSheetName: hints.preferredSheetName || hints.sheetName || '',
    fileName: hints.fileName || '',
    formFileName: hints.formFileName || '',
    formName: hints.formName || '',
    item: hints.item || null,
    formHeader: hints.formHeader || hints.parsedFormHeader || null,
    formHeaderTitle: hints.formHeaderTitle || hints.formHeader?.title || '',
  });
  if (!layout?.headers?.length) return null;
  return layout;
}

export function enrichFormDGJGujaratDisplayHeader(formHeader, fileName, item, tableHeaders, sheetText) {
  const base = formHeader && typeof formHeader === 'object' ? { ...formHeader } : {};
  const parts = [item?.formName, item?.FormName, fileName, base.title, base.subtitle, sheetText]
    .filter(Boolean)
    .join(' ')
    .toLowerCase();
  const isFormD =
    isFormDGJGujaratContext(base, item, fileName, sheetText, tableHeaders) ||
    /\bform[\s._-]*d\b/.test(parts);

  if (!isFormD) return base;

  const fields = enrichEstablishmentPrincipalEmployerHeaderFields(base.fields || []);
  const hasPeriod = fields.some((f) => /for\s+the\s+period\s+from/i.test(String(f?.label || '')));
  if (!hasPeriod) {
    fields.push({ label: 'For the period From', value: '', key: 'form_d_gj_period' });
  }

  return {
    ...base,
    fields,
    formVIIAPHeaderFieldLayout: false,
    formXXVIAPHeaderFieldLayout: false,
    formXIXAPHeaderFieldLayout: false,
    title:
      /\bform[\s._-]*d\b/i.test(String(base.title || '')) ||
      /form[\s._-]*d[\s._-]*gj/i.test(parts)
        ? base.title || 'FORM D'
        : 'FORM D',
    subtitle:
      /muster|relay\s+or\s+set|register\s+keeper/i.test(String(base.subtitle || ''))
        ? base.subtitle
        : 'Muster Roll (Relay or Set Work)',
  };
}

const stripLeadingNumber = (s) =>
  String(s || '')
    .replace(/^\(?\d+\)?[\.\)]\s*/, '')
    .trim();

export function parseFormDGJHeaderLabel(h) {
  const raw = stripLeadingNumber(h);
  const m = String(raw || '').match(/^(.+)_([\s\S]+)$/);
  if (m) return String(m[2] || '').trim();
  return raw;
}

export function normDGJHeader(h) {
  return formDGJGujaratHeaderNorm(parseFormDGJHeaderLabel(h))
    .replace(/[^a-z0-9 ]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function unwrapEmployeeRecord(emp) {
  return emp?.Employee || emp?.employee || emp;
}

function pickEmployeeValue(emp, keys) {
  const src = unwrapEmployeeRecord(emp);
  if (!src || typeof src !== 'object') return '';
  for (let i = 0; i < keys.length; i += 1) {
    const raw = src[keys[i]];
    if (raw == null || raw === '') continue;
    if (typeof raw === 'object') {
      const nested = String(raw.displayValue ?? raw.name ?? raw.Name ?? '').trim();
      if (nested) return nested;
      continue;
    }
    const text = String(raw).trim();
    if (text) return text;
  }
  return '';
}

export function readFormDGJEmployeeFullName(emp) {
  const fn = pickEmployeeValue(emp, ['FirstName', 'First Name', 'firstName']);
  const ln = pickEmployeeValue(emp, ['LastName', 'Last Name', 'lastName']);
  if (fn && ln) return `${fn} ${ln}`;
  return fn || ln || '';
}

export function readFormDGJDesignation(emp) {
  return pickEmployeeValue(emp, [
    'Designation',
    'designation',
    'Designation.displayValue',
    'Department',
    'department',
  ]);
}

export function isFormDGJNameHeader(h) {
  const s = normDGJHeader(h);
  if (!/\bname\b/.test(s)) return false;
  if (/father|spouse|husband|employer|establishment|person|presence|explanation|bank|register\s+keeper|signature/.test(s)) {
    return false;
  }
  return s === 'name' || /^name\b/.test(s);
}

export function isFormDGJRelayOrSetWorkHeader(h) {
  const s = normDGJHeader(h);
  return /relay\s+or\s+set/.test(s) || (s.includes('relay') && s.includes('set') && s.includes('work'));
}

export function applyFormDGJGujaratEmployeeToRow(row, emp, headers, helpers = {}) {
  const hdrs = Array.isArray(headers) ? headers : [];
  const out = row && typeof row === 'object' ? { ...row } : {};
  const {
    sanitizeValue = (v) => String(v ?? '').trim(),
    overwrite = true,
  } = helpers;

  const cellIsEmpty = (header) => {
    const v = String(out[header] || '').trim();
    if (!v) return true;
    const s = v.toLowerCase();
    return /^enter\b/.test(s) || s.includes('enter ');
  };
  const setCell = (header, value) => {
    if (!header || value == null || value === '') return;
    if (!overwrite && !cellIsEmpty(header)) return;
    out[header] = sanitizeValue(value);
  };

  const fullName = readFormDGJEmployeeFullName(emp);
  const designation = readFormDGJDesignation(emp);
  if (fullName) out.__employeeLookupName = fullName;

  hdrs.forEach((header) => {
    if (isFormDGJNameHeader(header)) {
      setCell(header, fullName);
      return;
    }
    if (isFormDGJRelayOrSetWorkHeader(header)) {
      setCell(header, designation);
    }
  });

  return out;
}

const MONTH_NAMES_DGJ = [
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

export function buildFormDGJGujaratPeriodLine(monthName, year) {
  let idx = MONTH_NAMES_DGJ.findIndex(
    (m) => m.toLowerCase() === String(monthName || '').toLowerCase().trim()
  );
  if (idx < 0) {
    const token = String(monthName || '').toLowerCase().trim().slice(0, 3);
    idx = MONTH_NAMES_DGJ.findIndex((m) => m.toLowerCase().startsWith(token));
  }
  if (idx < 0) idx = new Date().getMonth();
  const y = Number(year);
  if (!Number.isFinite(y) || y < 1900) return '';
  const m = idx + 1;
  const lastDay = new Date(y, idx + 1, 0).getDate();
  const pad = (n) => String(n).padStart(2, '0');
  return `For the period From ${pad(1)}-${pad(m)}-${y} to ${pad(lastDay)}-${pad(m)}-${y}`;
}

export function prepareFormDGJGujaratDownloadHeaderData(
  headerFormData,
  parsedFormHeader = null,
  siteContext = {}
) {
  const out = headerFormData && typeof headerFormData === 'object' ? { ...headerFormData } : {};
  const {
    establishmentText = '',
    principalEmployerText = '',
    periodText = '',
  } = siteContext;

  if (establishmentText) {
    out.statutory_establishment_name = establishmentText;
    out.form_d_gj_establishment = establishmentText;
  }
  if (principalEmployerText) {
    out.statutory_principal_employer = principalEmployerText;
    out.form_d_gj_principal_employer = principalEmployerText;
  }
  if (periodText) {
    out.form_d_gj_period = periodText;
    out.statutory_period_from = periodText;
  }

  const fields = Array.isArray(parsedFormHeader?.fields) ? parsedFormHeader.fields : [];
  fields.forEach((field) => {
    const key = field?.key;
    if (!key || String(out[key] ?? '').trim()) return;
    if (/establishment/i.test(String(field.label || '')) && !/principal|employer|contractor/.test(String(field.label || ''))) {
      if (establishmentText) out[key] = establishmentText;
    } else if (/principal\s+employer/i.test(String(field.label || ''))) {
      if (principalEmployerText) out[key] = principalEmployerText;
    } else if (/for\s+the\s+period\s+from/i.test(String(field.label || ''))) {
      if (periodText) out[key] = periodText;
    }
  });

  return out;
}

function writeFormDGJGujaratPeriodCell(worksheet, periodText, headerRowEnd = 25) {
  const line = String(periodText || '').trim();
  if (!line || !worksheet) return;
  for (let r = 1; r <= headerRowEnd; r += 1) {
    for (let c = 1; c <= 24; c += 1) {
      const raw = excelCellValueToString(worksheet.getCell(r, c)?.value).trim();
      if (/for\s+the\s+period\s+from/i.test(raw)) {
        worksheet.getCell(r, c).value = line;
        return;
      }
    }
  }
}

export function writeFormDGJGujaratHeaderFieldsToWorksheet(
  worksheet,
  headerFormData,
  parsedFormHeader,
  helpers = {}
) {
  if (!worksheet || !headerFormData || typeof headerFormData !== 'object') return;
  const headerRowEnd = Math.max(1, Number(helpers.headerRowEnd) || 25);
  writeStatutoryHeaderFieldsToExcelJsWorksheet(worksheet, {
    headerFormData,
    parsedFormHeader,
    headerRowEnd,
    maxScanCols: 80,
    writeMode: 'both',
  });
  writeFormDGJGujaratPeriodCell(
    worksheet,
    headerFormData.form_d_gj_period || headerFormData.statutory_period_from,
    headerRowEnd
  );
}

function buildMergeTopLeftResolver(worksheet) {
  const cache = new Map();
  return (r, c) => {
    const key = `${r}:${c}`;
    if (cache.has(key)) return cache.get(key);
    let topLeft = { r, c };
    const merges = worksheet.model?.merges;
    if (Array.isArray(merges)) {
      for (let mi = 0; mi < merges.length; mi += 1) {
        const parts = String(merges[mi] || '').split(':');
        if (parts.length !== 2) continue;
        const tl = worksheet.getCell(parts[0]);
        const br = worksheet.getCell(parts[1]);
        if (!tl || !br) continue;
        if (r >= tl.row && r <= br.row && c >= tl.col && c <= br.col) {
          topLeft = { r: tl.row, c: tl.col };
          break;
        }
      }
    }
    cache.set(key, topLeft);
    return topLeft;
  };
}

function excelCellLooksLikeFormDGJTableHeader(text) {
  const t = formDGJGujaratHeaderNorm(text).replace(/[^a-z0-9 ]/g, ' ').replace(/\s+/g, ' ').trim();
  return (
    /^sr\.?\s*no|^s\.?\s*no|serial|sl\.?\s*no/.test(t) ||
    /relay\s+or\s+set/.test(t) ||
    (t === 'name' || /^name\b/.test(t))
  );
}

function detectFormDGJGujaratTableLayout(worksheet, hints = {}) {
  const getMergeTopLeft = buildMergeTopLeftResolver(worksheet);
  const getMergedAwareCellText = (r, c) => {
    const tl = getMergeTopLeft(r, c);
    return excelCellValueToString(worksheet.getCell(tl.r, tl.c)?.value).trim();
  };

  let headerRow =
    hints.parsedHeaderRowIndex != null && hints.parsedHeaderRowIndex >= 0
      ? hints.parsedHeaderRowIndex + 1
      : -1;
  let startCol =
    hints.parsedTableStartCol != null && hints.parsedTableStartCol >= 0
      ? hints.parsedTableStartCol + 1
      : 1;

  if (headerRow < 1) {
    const maxScanRows = Math.max(40, worksheet.rowCount + 5);
    for (let r = 1; r <= maxScanRows; r += 1) {
      let foundRelay = false;
      for (let c = 1; c <= 20; c += 1) {
        const text = getMergedAwareCellText(r, c);
        if (/relay\s+or\s+set/i.test(text)) {
          headerRow = r;
          foundRelay = true;
          break;
        }
      }
      if (foundRelay) break;
      for (let c = 1; c <= 20; c += 1) {
        if (excelCellLooksLikeFormDGJTableHeader(getMergedAwareCellText(r, c))) {
          headerRow = r;
          startCol = c;
          break;
        }
      }
      if (headerRow > 0) break;
    }
  }

  if (headerRow < 1) return null;

  if (startCol === 1) {
    for (let c = 1; c <= 20; c += 1) {
      if (excelCellLooksLikeFormDGJTableHeader(getMergedAwareCellText(headerRow, c))) {
        startCol = c;
        break;
      }
    }
  }

  const templateCols = [];
  for (let c = startCol; c <= startCol + 20; c += 1) {
    const label = getMergedAwareCellText(headerRow, c);
    if (!label) {
      if (templateCols.length >= 4) break;
      continue;
    }
    templateCols.push({ col: c, label });
    if (templateCols.length >= 12) break;
  }
  if (templateCols.length < 3) return null;

  const dataStartRow =
    hints.parsedDataStartIndex != null && hints.parsedDataStartIndex >= 0
      ? hints.parsedDataStartIndex + 1
      : headerRow + 1;

  return { headerRow, dataStartRow, templateCols, startCol };
}

export function getFormDGJGujaratRowValueForHeader(row, header, rowIndex = 0) {
  if (!row || !header) return '';
  const direct = row[header];
  if (direct != null && String(direct).trim() !== '') return String(direct).trim();
  if (isFormDGJNameHeader(header)) {
    for (const [k, v] of Object.entries(row)) {
      if (isFormDGJNameHeader(k) && v != null && String(v).trim() !== '') return String(v).trim();
    }
  }
  if (isFormDGJRelayOrSetWorkHeader(header)) {
    for (const [k, v] of Object.entries(row)) {
      if (isFormDGJRelayOrSetWorkHeader(k) && v != null && String(v).trim() !== '') {
        return String(v).trim();
      }
    }
  }
  if (/sr\.?\s*no|serial|sl\.?\s*no/i.test(normDGJHeader(header))) {
    return String(rowIndex + 1);
  }
  return '';
}

export function rowHasMeaningfulFormDGJGujaratExportData(row, headers) {
  const hdrs = resolveFormDGJGujaratTableHeaders(headers);
  return hdrs.some((header) => {
    if (/sr\.?\s*no|serial|sl\.?\s*no/i.test(normDGJHeader(header))) return false;
    return getFormDGJGujaratRowValueForHeader(row, header) !== '';
  });
}

export function filterFormDGJGujaratExportRows(rows, headers) {
  const hdrs = resolveFormDGJGujaratTableHeaders(headers);
  return (Array.isArray(rows) ? rows : []).filter((row) =>
    rowHasMeaningfulFormDGJGujaratExportData(row, hdrs)
  );
}

export function remapFormDGJGujaratRowsToHeaders(rows, sourceHeaders, targetHeaders) {
  const src = Array.isArray(sourceHeaders) ? sourceHeaders : [];
  const tgt = resolveFormDGJGujaratTableHeaders(targetHeaders);
  if (!Array.isArray(rows)) return [];
  return rows.map((row, rowIndex) => {
    if (!row || typeof row !== 'object') return {};
    const out = {};
    tgt.forEach((targetHeader, colIdx) => {
      let val = '';
      if (Object.prototype.hasOwnProperty.call(row, targetHeader)) {
        val = row[targetHeader];
      } else if (src[colIdx] && Object.prototype.hasOwnProperty.call(row, src[colIdx])) {
        val = row[src[colIdx]];
      } else {
        for (const [k, v] of Object.entries(row)) {
          if (normDGJHeader(k) === normDGJHeader(targetHeader)) {
            val = v;
            break;
          }
        }
      }
      if ((val == null || val === '') && /sr\.?\s*no|serial|sl\.?\s*no/i.test(normDGJHeader(targetHeader))) {
        val = rowIndex + 1;
      }
      out[targetHeader] = val == null ? '' : val;
    });
    return out;
  });
}

function resolveFormDGJGujaratWorksheet(workbook, hints = {}) {
  if (!workbook) return null;
  const preferred = String(hints.preferredSheetName || '').trim();
  if (preferred) {
    const ws = workbook.getWorksheet(preferred);
    if (ws) return ws;
  }
  for (const ws of workbook.worksheets || []) {
    if (/\bform[\s._-]*d\b/i.test(String(ws.name || ''))) return ws;
  }
  return workbook.worksheets?.[0] || null;
}

/** Write employee rows into the Gujarat Form D template (FORM D sheet). */
export async function buildFormDGJGujaratWorkbookWithTemplateStyles({
  templateArrayBuffer,
  mappedData,
  headersToUse,
  parsedFormHeader,
  formFileName,
  headerFormData,
  parsedHeaderRowIndex,
  parsedDataStartIndex,
  parsedTableStartCol,
  preferredSheetName = '',
}) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(templateArrayBuffer);
  const worksheet = resolveFormDGJGujaratWorksheet(workbook, { preferredSheetName });
  if (!worksheet) throw new Error('Template worksheet not found.');

  const layout = detectFormDGJGujaratTableLayout(worksheet, {
    parsedHeaderRowIndex,
    parsedDataStartIndex,
    parsedTableStartCol,
  });
  if (!layout) throw new Error('Could not locate Gujarat Form D table header row.');

  const { dataStartRow, templateCols } = layout;

  writeFormDGJGujaratHeaderFieldsToWorksheet(worksheet, headerFormData, parsedFormHeader, {
    headerRowEnd: Math.max(1, layout.headerRow - 1),
  });

  const templateHeaderLabels = templateCols.map((entry) => entry.label);
  const normalizedHeaders = resolveFormDGJGujaratTableHeaders(
    Array.isArray(headersToUse) && headersToUse.length >= templateHeaderLabels.length
      ? headersToUse
      : templateHeaderLabels.length > 0
        ? templateHeaderLabels
        : headersToUse
  );
  const rows = filterFormDGJGujaratExportRows(
    remapFormDGJGujaratRowsToHeaders(
      Array.isArray(mappedData) ? mappedData : [],
      normalizedHeaders,
      normalizedHeaders
    ),
    normalizedHeaders
  );

  const tableColMin = Math.min(...templateCols.map(({ col }) => col));
  const tableColMax = Math.max(...templateCols.map(({ col }) => col));
  const clearToRow = Math.max(dataStartRow + rows.length + 10, dataStartRow + 15);
  for (let r = dataStartRow; r <= clearToRow; r += 1) {
    for (let c = tableColMin; c <= tableColMax; c += 1) {
      worksheet.getCell(r, c).value = '';
    }
  }

  rows.forEach((row, idx) => {
    const targetRow = worksheet.getRow(dataStartRow + idx);
    if (targetRow) targetRow.height = 18;
    templateCols.forEach(({ col, label }) => {
      const val = getFormDGJGujaratRowValueForHeader(row, label, idx);
      const cell = worksheet.getCell(dataStartRow + idx, col);
      if (val == null || val === '') {
        cell.value = '';
        return;
      }
      if (/sr\.?\s*no|serial|sl\.?\s*no/i.test(normDGJHeader(label))) {
        const n = Number(String(val).replace(/[,]/g, '').trim());
        cell.value = Number.isFinite(n) ? n : String(val);
      } else {
        cell.value = String(val);
      }
      cell.alignment = { ...(cell.alignment || {}), vertical: 'middle', wrapText: true };
    });
  });

  if (rows.length > 0) {
    ensureExcelJSDataRowsWithBorders(worksheet, {
      dataStartRow,
      dataRowCount: rows.length,
      colFrom: tableColMin,
      colTo: tableColMax,
      templateRow: dataStartRow,
      templateBodyRows: 1,
    });
  }

  const out = await workbook.xlsx.writeBuffer();
  const fileName =
    formFileName ||
    parsedFormHeader?.title?.replace(/[^a-zA-Z0-9]/g, '_') ||
    'Form_D_GJ_Gujarat.xlsx';
  return {
    blob: new Blob([out], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    }),
    fileName,
  };
}
