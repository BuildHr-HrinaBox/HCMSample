/** Form XXIII GJ — Register of Overtime (Gujarat). */

import {
  FORM_XXIII_GJ_OT_PDF_HEADERS,
  isFormXXIIIGJOtSpuriousPdfDataRow,
  scrubFormXXIIIGJOtPdfDataLine,
} from '../../utils/statutoryDraftPdf.formXXIII.GJ';

/** Canonical August-style 12-column OT table (never wages-register headers). */
export const FORM_XXIII_GJ_OT_TABLE_HEADERS = [...FORM_XXIII_GJ_OT_PDF_HEADERS];

export function isFormXXIIIGJContext(formHeader, rowItem, fileName, sheetText = '') {
  const parts = [
    rowItem?.formName,
    rowItem?.FormName,
    rowItem?.state,
    rowItem?.State,
    rowItem?.siteState,
    rowItem?.SiteState,
    rowItem?.formFileName,
    rowItem?.FormFileName,
    fileName,
    formHeader?.title,
    formHeader?.subtitle,
    sheetText,
  ]
    .filter((x) => x != null && String(x).trim() !== '')
    .join(' ')
    .toLowerCase();

  const hasGujarat =
    /gujarat/.test(parts) ||
    /(?:^|[\s,_-])gj(?:$|[\s,_-])/.test(parts) ||
    /\bform[\s._-]*xxiii[\s._-]*gj\b/.test(parts) ||
    /\bxxiii[\s._-]*gj\b/.test(parts) ||
    /form_xxiii_gj/.test(parts);
  const hasXXIII =
    /\bform[\s._-]*xxiii/i.test(parts) || /register\s+of\s+overtime/.test(parts);
  return hasGujarat && hasXXIII;
}

/** Form XXIII GJ overtime columns that always default to NIL (no attendance/payroll OT). */
export const FORM_XXIII_GJ_OT_NIL = 'NIL';

const normFormXXIIIGJHeader = (h) =>
  String(h || '')
    .replace(/\r?\n/g, ' ')
    .replace(/[\u2018\u2019\u2032\u0060]/g, "'")
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();

/**
 * Form XIV-style ordinal preamble lines on XXIII GJ templates — not the OT table S.No.
 */
export function isFormXXIIIGJOtPreambleFieldLabel(h) {
  const raw = String(h || '');
  const s = normFormXXIIIGJHeader(raw);
  if (!s) return false;
  if (/^(?:serial|sl\.?\s*no|s\.?\s*no)\.?$/.test(s.replace(/[^a-z0-9 ]/g, ' ').replace(/\s+/g, ' ').trim())) {
    return false;
  }
  if (/^\d+[.)]\s+/.test(s)) return true;
  if (s.includes('serial') && s.includes('register') && (s.includes('workmen') || s.includes('workman'))) {
    return true;
  }
  if (/\.{5,}|_{5,}|…/.test(raw)) {
    return (
      /name of the workman/.test(s) ||
      (/nature of employment/.test(s) && !/designation/.test(s)) ||
      (/wage rate/.test(s) && /piece/.test(s)) ||
      /wage period/.test(s) ||
      /tenure of employment/.test(s)
    );
  }
  return (
    /name of the workman/.test(s) ||
    (/nature of employment/.test(s) && !/designation/.test(s)) ||
    (/wage rate/.test(s) && /piece/.test(s)) ||
    /wage period/.test(s) ||
    /tenure of employment/.test(s)
  );
}

export function isFormXXIIIGJSexHeader(h) {
  const s = normFormXXIIIGJHeader(h);
  return s === 'sex' || s === 'gender';
}

export function resolveFormXXIIIGJOtTableHeaders(headers) {
  return [...FORM_XXIII_GJ_OT_TABLE_HEADERS];
}

/** Score a joined header-band row — OT thead beats Form XIV preamble. */
export function scoreFormXXIIIGJOtTableHeaderRow(text) {
  const joined = normFormXXIIIGJHeader(text);
  if (!joined) return -999;
  let score = 0;
  if (/\bovertime\b/.test(joined)) score += 10;
  if (/normal\s+rate/.test(joined)) score += 6;
  if (/name.*surname|surname.*workm|name.*workm|name of workman/.test(joined)) score += 4;
  if (/\bsex\b|designation|father/.test(joined)) score += 2;
  if (/sl\.?\s*no|s\.?\s*no|serial\s*no/.test(joined)) score += 5;
  if (/name of the workman|wage period|tenure of employment/.test(joined) && !/\bovertime\b/.test(joined)) {
    score -= 40;
  }
  if (/nature of employment/.test(joined) && !/designation/.test(joined) && !/\bovertime\b/.test(joined)) {
    score -= 20;
  }
  if (
    /name and address|principal employer|\bcontractor\b|establishment in which|nature and location|see rule|form\s*xxiii|register of overtime/i.test(
      joined
    )
  ) {
    score -= 35;
  }
  if (isFormXXIIIGJOtPreambleFieldLabel(text)) score -= 50;
  return score;
}

function findFormXXIIIGJPreambleEndRow(worksheet) {
  let last = -1;
  for (let r = 1; r <= 40; r += 1) {
    for (let c = 1; c <= 3; c += 1) {
      if (isFormXXIIIGJOtPreambleFieldLabel(excelJsPlainText(worksheet.getCell(r, c)))) {
        last = r;
        break;
      }
    }
  }
  return last;
}

function joinedFormXXIIIGJRowText(worksheet, r) {
  let joined = '';
  for (let c = 1; c <= 12; c += 1) {
    joined += ` ${excelJsPlainText(worksheet.getCell(r, c))}`;
  }
  return joined;
}

function resolveFormXXIIIGJOtHeaderRow1(worksheet) {
  const preambleEnd = findFormXXIIIGJPreambleEndRow(worksheet);
  const scanFrom = preambleEnd > 0 ? preambleEnd : 12;
  const scanTo = preambleEnd > 0 ? preambleEnd + 4 : 45;
  let headerRow1 = preambleEnd > 0 ? preambleEnd : 18;
  let bestScore = -Infinity;
  for (let r = scanFrom; r <= scanTo; r += 1) {
    const score = scoreFormXXIIIGJOtTableHeaderRow(joinedFormXXIIIGJRowText(worksheet, r));
    if (score > bestScore) {
      bestScore = score;
      headerRow1 = r;
    }
  }
  if (bestScore < 8) {
    headerRow1 = preambleEnd > 0 ? preambleEnd : 18;
  }
  return headerRow1;
}

function classifyFormXXIIIGJHeader(header) {
  const s = normFormXXIIIGJHeader(header);
  if (!s) return -1;
  if (/(?:^|\s)(?:sr|s|serial)\.?\s*no/.test(s)) return 0;
  if (/name of the employee|name and surname|name of workman|employee name/.test(s)) return 1;
  if (/father|husband/.test(s)) return 2;
  if (isFormXXIIIGJSexHeader(header)) return 3;
  if (/designation|nature of employment/.test(s)) return 4;
  if (isFormXXIIIGJOtWorkedDatesHeader(header)) return 5;
  if (isFormXXIIIGJTotalOvertimeWorkedHeader(header)) return 6;
  if (/normal\s+rate/.test(s) || (/rate of wages/.test(s) && !/overtime/.test(s))) return 7;
  if (isFormXXIIIGJOvertimeRateHeader(header)) return 8;
  if (isFormXXIIIGJOvertimeEarningsHeader(header)) return 9;
  if (isFormXXIIIGJOtWagesPaidDateHeader(header)) return 10;
  if (/remarks?/.test(s)) return 11;
  return -1;
}

function pickFormXXIIIGJFuzzyRowValue(row, header) {
  const target = normFormXXIIIGJHeader(header);
  if (!target || !row || typeof row !== 'object') return '';
  const keys = Object.keys(row).filter((k) => !String(k).startsWith('__'));
  const exact = keys.find((k) => normFormXXIIIGJHeader(k) === target);
  if (exact != null && row[exact] != null && String(row[exact]).trim() !== '') return row[exact];
  const fuzzy = keys.find((k) => {
    const n = normFormXXIIIGJHeader(k);
    return n && (n.includes(target) || target.includes(n));
  });
  return fuzzy != null ? row[fuzzy] : '';
}

export function remapFormXXIIIGJRowsToHeaders(rows, sourceHeaders, targetHeaders) {
  const targets = resolveFormXXIIIGJOtTableHeaders(targetHeaders);
  const sources = Array.isArray(sourceHeaders) ? sourceHeaders : targets;
  return (Array.isArray(rows) ? rows : []).map((row) => {
    if (!row || typeof row !== 'object' || Array.isArray(row)) return row;
    const next = {};
    Object.keys(row).forEach((k) => {
      if (String(k).startsWith('__')) next[k] = row[k];
    });
    targets.forEach((targetH) => {
      let val = row[targetH];
      if (val != null && String(val).trim() !== '') {
        next[targetH] = val;
        return;
      }
      const tClass = classifyFormXXIIIGJHeader(targetH);
      if (tClass >= 0) {
        for (let i = 0; i < sources.length; i += 1) {
          const sh = sources[i];
          if (classifyFormXXIIIGJHeader(sh) !== tClass) continue;
          const candidate = row[sh];
          if (candidate != null && String(candidate).trim() !== '') {
            val = candidate;
            break;
          }
        }
      }
      if (val == null || String(val).trim() === '') {
        val = pickFormXXIIIGJFuzzyRowValue(row, targetH);
      }
      next[targetH] = val != null ? val : '';
    });
    return next;
  });
}

function formXXIIIGJRowHasEmployeeBinding(row) {
  return !!(row.__employeeLookupId || row.__employeeLookupName || row.__employeeId || row.__payrollId);
}

export function isFormXXIIIGJFatherOrHusbandNameHeader(h) {
  const s = normFormXXIIIGJHeader(h);
  if (!s || s.includes('proof')) return false;
  return (
    (s.includes('father') || s.includes('husband') || s.includes('spouse')) &&
    (s.includes('name') || s.includes('father') || s.includes('husband'))
  );
}

export function formatFormXXIIIGJPersonNameDisplay(value) {
  const s = String(value || '')
    .trim()
    .replace(/\s+/g, ' ');
  if (!s) return '';
  return s
    .split(' ')
    .map((word) => {
      if (!word) return word;
      return word.charAt(0).toUpperCase() + word.slice(1).toLowerCase();
    })
    .join(' ');
}

function unwrapFormXXIIIGJEmployeeRecord(empItem) {
  if (!empItem || typeof empItem !== 'object') return null;
  if (empItem.Employee && typeof empItem.Employee === 'object') return empItem.Employee;
  if (empItem.employee && typeof empItem.employee === 'object') return empItem.employee;
  return empItem;
}

function pickFormXXIIIGJFatherOrSpouseFromEmployeeFields(emp) {
  if (!emp || typeof emp !== 'object') return '';
  const fatherName =
    emp.Father_s_Name ||
    emp['Father_s_Name'] ||
    emp.Father_Name ||
    emp['Father_Name'] ||
    emp.FathersName ||
    emp['FathersName'] ||
    emp['Father/Spouse Name'] ||
    emp['Father Name'] ||
    emp.fathersName ||
    emp['Father_s name'] ||
    '';
  if (String(fatherName || '').trim()) return String(fatherName).trim();
  const spouseName =
    emp.Spouse_Name ||
    emp['Spouse_Name'] ||
    emp.spouseName ||
    emp['Spouse Name'] ||
    emp.Husband_Name ||
    emp['Husband_Name'] ||
    emp['Husband Name'] ||
    '';
  if (String(spouseName || '').trim()) return String(spouseName).trim();
  for (const [key, raw] of Object.entries(emp)) {
    if (raw == null || String(raw).trim() === '') continue;
    const nk = normFormXXIIIGJHeader(key);
    if (!nk || nk.includes('proof') || /aadhaar|aadhar|pan\b|uan\b|pf\b|esi\b/.test(nk)) continue;
    if (!/father|husband|spouse/.test(nk)) continue;
    if (/name of (the )?(employee|workman|worker)/.test(nk)) continue;
    return String(raw).trim();
  }
  return '';
}

export function resolveFormXXIIIGJFatherOrSpouseFromEmployee(empItem) {
  const emp = unwrapFormXXIIIGJEmployeeRecord(empItem);
  if (!emp) return '';
  return pickFormXXIIIGJFatherOrSpouseFromEmployeeFields(emp);
}

function normalizeFormXXIIIGJEmployeeMatchToken(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ');
}

function collectFormXXIIIGJEmployeeIdTokens(empItem) {
  const emp = unwrapFormXXIIIGJEmployeeRecord(empItem);
  if (!emp) return [];
  return [
    emp.EmployeeID,
    emp['EmployeeID'],
    emp['Employee ID'],
    emp.employeeId,
    emp.Zoho_ID,
    emp['Zoho_ID'],
    emp.zoho_id,
    emp['Role.ID'],
    emp?.Role && typeof emp.Role === 'object' ? emp.Role.ID : '',
    emp.employeeCode,
    emp.EmployeeCode,
    emp['Employee Code'],
  ]
    .map((v) => String(v || '').trim().toLowerCase())
    .filter(Boolean);
}

export function findFormXXIIIGJEmployeeForRow(row, employees) {
  if (!row || typeof row !== 'object' || !Array.isArray(employees) || employees.length === 0) {
    return null;
  }
  const rowIds = new Set(
    [row.__employeeLookupId, row.__employeeId, row.__payrollId]
      .map((v) => String(v || '').trim().toLowerCase())
      .filter(Boolean)
  );
  const rowName = normalizeFormXXIIIGJEmployeeMatchToken(
    row.__employeeLookupName ?? row.__employeelookupname
  );
  for (let i = 0; i < employees.length; i += 1) {
    const emp = employees[i];
    const ids = collectFormXXIIIGJEmployeeIdTokens(emp);
    if (rowIds.size > 0 && ids.some((id) => rowIds.has(id))) return emp;
    if (rowName) {
      const e = unwrapFormXXIIIGJEmployeeRecord(emp);
      if (!e) continue;
      const parts = [
        e.FirstName,
        e['FirstName'],
        e.LastName,
        e['LastName'],
        e.EmployeeName,
        e['Employee Name'],
        e.Name,
      ]
        .map((v) => String(v || '').trim())
        .filter(Boolean);
      const full = normalizeFormXXIIIGJEmployeeMatchToken(parts.join(' '));
      if (full && full === rowName) return emp;
    }
  }
  return null;
}

export function attachFormXXIIIGJEmployeeMetaToRows(rows, employees) {
  if (!Array.isArray(rows)) return [];
  if (!Array.isArray(employees) || employees.length === 0) return rows;
  return rows.map((row) => {
    if (!row || typeof row !== 'object') return row;
    const emp = findFormXXIIIGJEmployeeForRow(row, employees);
    if (!emp) return row;
    const father = resolveFormXXIIIGJFatherOrSpouseFromEmployee(emp);
    const next = { ...row };
    if (father && !String(next.__employeeLookupFatherOrSpouse ?? '').trim()) {
      next.__employeeLookupFatherOrSpouse = father;
    }
    return next;
  });
}

export function resolveFormXXIIIGJFatherOrSpouseFromRow(row) {
  if (!row || typeof row !== 'object') return '';
  const meta = String(row.__employeeLookupFatherOrSpouse ?? row.__employeeLookupFather ?? '').trim();
  if (meta) return meta;
  const keys = Object.keys(row).filter((k) => !String(k).startsWith('__'));
  for (let i = 0; i < keys.length; i += 1) {
    const k = keys[i];
    const s = normFormXXIIIGJHeader(k);
    if (!s || s.includes('proof')) continue;
    if (!/father|husband|spouse/.test(s)) continue;
    const v = String(row[k] ?? '').trim();
    if (v && !/^enter\b/i.test(v)) return row[k];
  }
  return '';
}

function resolveFormXXIIIGJIdentityHeaderSlots(headers) {
  const list = Array.isArray(headers) ? headers : resolveFormXXIIIGJOtTableHeaders(headers);
  const canon = resolveFormXXIIIGJOtTableHeaders(list);
  const nameHdr = list.find((h) => classifyFormXXIIIGJHeader(h) === 1) || canon[1];
  const fatherHdr = list.find((h) => classifyFormXXIIIGJHeader(h) === 2) || canon[2];
  return { nameHdr, fatherHdr };
}

function applyFormXXIIIGJIdentityDisplay(row, headers) {
  const { nameHdr, fatherHdr } = resolveFormXXIIIGJIdentityHeaderSlots(headers);
  const next = row && typeof row === 'object' ? { ...row } : {};
  if (String(next[nameHdr] ?? '').trim()) {
    next[nameHdr] = formatFormXXIIIGJPersonNameDisplay(next[nameHdr]);
  }
  if (String(next[fatherHdr] ?? '').trim()) {
    next[fatherHdr] = formatFormXXIIIGJPersonNameDisplay(next[fatherHdr]);
  }
  return next;
}

/** Autofill grid often keeps the person name on __employeeLookupName while OT columns are filled. */
export function enrichFormXXIIIGJRowIdentityFromMeta(row, headers = FORM_XXIII_GJ_OT_TABLE_HEADERS, opts = {}) {
  if (!row || typeof row !== 'object') return row;
  const { nameHdr, fatherHdr } = resolveFormXXIIIGJIdentityHeaderSlots(headers);
  const next = { ...row };
  if (!String(next[nameHdr] ?? '').trim()) {
    const lookup = String(
      next.__employeeLookupName ?? next.__employeelookupname ?? ''
    ).trim();
    if (lookup) next[nameHdr] = lookup;
  }
  if (!String(next[fatherHdr] ?? '').trim()) {
    let father = resolveFormXXIIIGJFatherOrSpouseFromRow(next);
    if (!father && typeof opts.findEmployeeForRow === 'function') {
      const emp = opts.findEmployeeForRow(next);
      father = resolveFormXXIIIGJFatherOrSpouseFromEmployee(emp);
    }
    if (father !== '') next[fatherHdr] = father;
  }
  return applyFormXXIIIGJIdentityDisplay(next, headers);
}

/** Align autofill grid rows to table headers and fill name/father from People meta. */
export function resolveFormXXIIIGJDisplayRows(rows, sourceHeaders, targetHeaders, employees = null) {
  const targets =
    Array.isArray(targetHeaders) && targetHeaders.length > 0
      ? targetHeaders
      : resolveFormXXIIIGJOtTableHeaders(targetHeaders);
  const list = Array.isArray(employees) && employees.length > 0 ? employees : null;
  const withMeta = attachFormXXIIIGJEmployeeMetaToRows(Array.isArray(rows) ? rows : [], list);
  const remapped = remapFormXXIIIGJRowsToHeaders(withMeta, sourceHeaders, targets);
  const findEmployeeForRow =
    list != null ? (row) => findFormXXIIIGJEmployeeForRow(row, list) : null;
  return remapped.map((row) =>
    scrubFormXXIIIGJOtRowHeaderBleed(
      enrichFormXXIIIGJRowIdentityFromMeta(row, targets, {
        findEmployeeForRow: findEmployeeForRow || undefined,
      }),
      targets
    )
  );
}

function compactNormFormXXIIIGJCompare(text) {
  return normFormXXIIIGJHeader(text).replace(/[^a-z0-9 ]/g, '').trim();
}

export function isFormXXIIIGJRemarksHeader(h) {
  const s = compactNormFormXXIIIGJCompare(h);
  return s === 'remarks' || /^remarks?$/.test(s);
}

/** Remove OT thead / 1–12 index-row text accidentally stored on employee rows. */
export function sanitizeFormXXIIIGJOtExportCellValue(value, header, columnIndex = -1) {
  const s = String(value ?? '').trim();
  if (!s) return '';
  const cellNorm = compactNormFormXXIIIGJCompare(s);
  const hdrNorm = compactNormFormXXIIIGJCompare(header);
  if (cellNorm && hdrNorm && cellNorm === hdrNorm) return '';
  if (isFormXXIIIGJRemarksHeader(header) && /^remarks?$/i.test(s)) return '';
  const colIdx = Number(columnIndex);
  if (colIdx >= 0 && classifyFormXXIIIGJHeader(header) !== 0 && /^\d{1,2}$/.test(s)) {
    if (Number(s) === colIdx + 1) return '';
  }
  return s;
}

export function scrubFormXXIIIGJOtRowHeaderBleed(row, headers) {
  const hdrs = Array.isArray(headers) ? headers : resolveFormXXIIIGJOtTableHeaders(headers);
  if (!row || typeof row !== 'object') return row;
  const next = { ...row };
  hdrs.forEach((header, j) => {
    if (!header || !Object.prototype.hasOwnProperty.call(next, header)) return;
    next[header] = sanitizeFormXXIIIGJOtExportCellValue(next[header], header, j);
  });
  return next;
}

function isFormXXIIIGJTemplateSampleRow(row, opts = {}) {
  if (formXXIIIGJRowHasEmployeeBinding(row)) return false;
  const name = String(
    row['Name of the Employee'] ??
      row['Name and surname of workmen'] ??
      row[FORM_XXIII_GJ_OT_TABLE_HEADERS[1]] ??
      ''
  ).trim();
  const rate = String(row['Rate of Wages'] ?? row['Normal rate of wages'] ?? '').trim();
  if (name && rate === '0') return true;
  if (opts.autofillOnly && name) return true;
  return false;
}

export function prepareFormXXIIIGJOvertimeExportRows(rows, headers, opts = {}) {
  const canon = resolveFormXXIIIGJOtTableHeaders(headers);
  const employees = Array.isArray(opts.employees) ? opts.employees : [];
  const enrichOpts = {
    ...opts,
    findEmployeeForRow:
      typeof opts.findEmployeeForRow === 'function'
        ? opts.findEmployeeForRow
        : employees.length > 0
          ? (row) => findFormXXIIIGJEmployeeForRow(row, employees)
          : undefined,
  };
  const sourceRows = attachFormXXIIIGJEmployeeMetaToRows(
    Array.isArray(rows) ? rows : [],
    employees
  );
  const remapped = remapFormXXIIIGJRowsToHeaders(sourceRows, headers, canon);
  const out = [];
  remapped.forEach((row) => {
    if (!row || typeof row !== 'object') return;
    if (isFormXXIIIGJTemplateSampleRow(row, opts)) return;
    const next = scrubFormXXIIIGJOtRowHeaderBleed(
      applyFormXXIIIGJOtNilToRow(
        enrichFormXXIIIGJRowIdentityFromMeta(row, canon, enrichOpts),
        canon,
        { overwrite: true }
      ),
      canon
    );
    next[canon[0]] = String(out.length + 1);
    const name = String(next[canon[1]] ?? '').trim();
    const meaningful = canon.some(
      (h, i) =>
        i > 0 &&
        String(next[h] ?? '').trim() !== '' &&
        next[h] !== FORM_XXIII_GJ_OT_NIL
    );
    const pdfLine = scrubFormXXIIIGJOtPdfDataLine(canon.map((h) => next[h]));
    if (isFormXXIIIGJOtSpuriousPdfDataRow(pdfLine)) return;
    canon.forEach((h, j) => {
      next[h] = pdfLine[j] != null ? pdfLine[j] : next[h];
    });
    if (name || meaningful) out.push(next);
  });
  return out;
}

function excelJsPlainText(cell) {
  const v = cell?.value;
  if (v == null) return '';
  if (typeof v === 'string' || typeof v === 'number' || typeof v === 'boolean') return String(v).trim();
  if (typeof v === 'object') {
    if (Array.isArray(v.richText)) return v.richText.map((rt) => rt?.text || '').join('').trim();
    if (v.text != null) return String(v.text).trim();
    if (v.result != null) return String(v.result).trim();
  }
  return String(v).trim();
}

function isFormXXIIIGJWagesRegisterSpillLabel(text) {
  const s = normFormXXIIIGJHeader(text);
  if (!s) return false;
  return (
    /rate of wages/.test(s) &&
    (/normal wages|gross wage|actual wage|earned/.test(s) || /_/.test(String(text || '')))
  );
}

const thinBorderGj = {
  top: { style: 'thin' },
  left: { style: 'thin' },
  bottom: { style: 'thin' },
  right: { style: 'thin' },
};

function clearFormXXIIIGJMetaCellPresentation(cell) {
  if (!cell) return;
  cell.border = {};
}

/** Fit Form XXIII GJ onto the August 12-col OT band on an ExcelJS worksheet. */
export function applyFormXXIIIGJOtTemplateLayoutToExcelJs(worksheet, headersInput) {
  if (!worksheet) return null;
  const headers = resolveFormXXIIIGJOtTableHeaders(headersInput);
  const tableCols = headers.length;

  const headerRow1 = resolveFormXXIIIGJOtHeaderRow1(worksheet);

  const indexRow1 = headerRow1 + 1;
  const dataStartRow1 = indexRow1 + 1;

  for (let r = 1; r <= headerRow1; r += 1) {
    for (let c = 1; c <= 12; c += 1) {
      const cell = worksheet.getCell(r, c);
      const t = excelJsPlainText(cell);
      if (isFormXXIIIGJOtPreambleFieldLabel(t) || isFormXXIIIGJWagesRegisterSpillLabel(t)) {
        cell.value = null;
      }
      if (r < headerRow1 || isFormXXIIIGJOtPreambleFieldLabel(t) || isFormXXIIIGJWagesRegisterSpillLabel(t)) {
        clearFormXXIIIGJMetaCellPresentation(cell);
      }
    }
    const spillFrom = r < headerRow1 ? 10 : 13;
    for (let c = spillFrom; c <= 20; c += 1) {
      worksheet.getCell(r, c).value = null;
    }
  }

  headers.forEach((header, j) => {
    const cell = worksheet.getCell(headerRow1, j + 1);
    cell.value = header;
    cell.border = { ...thinBorderGj };
  });
  for (let j = 0; j < tableCols; j += 1) {
    const cell = worksheet.getCell(indexRow1, j + 1);
    cell.value = j + 1;
    cell.border = { ...thinBorderGj };
  }

  const excelWidths = [7, 22, 20, 7, 20, 12, 12, 11, 10, 10, 12, 12];
  for (let c = 1; c <= 60; c += 1) {
    if (typeof worksheet.getColumn !== 'function') break;
    const col = worksheet.getColumn(c);
    if (!col) continue;
    if (c <= tableCols) {
      col.hidden = false;
      if (col.model && typeof col.model === 'object') col.model.hidden = false;
      col.width = excelWidths[c - 1] != null ? excelWidths[c - 1] : 12;
    } else if (c <= 15) {
      col.hidden = true;
      if (col.model && typeof col.model === 'object') col.model.hidden = true;
    }
  }

  return {
    headers,
    headerRow1,
    indexRow1,
    dataStartRow1,
    snoRow0: headerRow1 - 1,
  };
}

/**
 * Unhide A–L, hide columns past L, plain header meta (no boxes), OT thead keeps borders.
 */
export function finalizeFormXXIIIGJOtWorksheetPresentation(worksheet, opts = {}) {
  if (!worksheet) return;
  const headerRowEnd = Math.max(1, Number(opts.headerRowEnd) || 15);
  const tableCols = Math.max(12, Number(opts.tableColCount) || 12);
  const excelWidths = [7, 22, 20, 7, 20, 12, 12, 11, 10, 10, 12, 12];

  for (let c = 1; c <= 60; c += 1) {
    if (typeof worksheet.getColumn !== 'function') break;
    const col = worksheet.getColumn(c);
    if (!col) continue;
    if (c <= tableCols) {
      col.hidden = false;
      if (col.model && typeof col.model === 'object') col.model.hidden = false;
      col.width = excelWidths[c - 1] != null ? excelWidths[c - 1] : 12;
    } else {
      col.hidden = true;
      if (col.model && typeof col.model === 'object') col.model.hidden = true;
    }
  }

  for (let r = 1; r <= headerRowEnd; r += 1) {
    for (let c = 1; c <= tableCols; c += 1) {
      clearFormXXIIIGJMetaCellPresentation(worksheet.getCell(r, c));
    }
    for (let c = tableCols + 1; c <= 15; c += 1) {
      worksheet.getCell(r, c).value = null;
    }
  }
}

/** Dates on which overtime worked */
export function isFormXXIIIGJOtWorkedDatesHeader(h) {
  const s = normFormXXIIIGJHeader(h);
  if (!s) return false;
  if (/payment|paid|rate|wage|earning/.test(s) && !/worked/.test(s)) return false;
  return (
    /dates?\s+on\s+which\s+over[\s-]*time\s+worked/.test(s) ||
    (s.includes('date') && s.includes('overtime') && s.includes('worked') && !/paid|payment/.test(s))
  );
}

/** Total overtime worked or production in case of piece rate */
export function isFormXXIIIGJTotalOvertimeWorkedHeader(h) {
  const s = normFormXXIIIGJHeader(h);
  if (!s) return false;
  return (
    (s.includes('total') && s.includes('overtime') && (s.includes('worked') || s.includes('production'))) ||
    (s.includes('overtime') && s.includes('piece') && s.includes('rate'))
  );
}

/** Overtime rate of wages */
export function isFormXXIIIGJOvertimeRateHeader(h) {
  const s = normFormXXIIIGJHeader(h);
  if (!s || /normal/.test(s)) return false;
  return s.includes('overtime') && s.includes('rate') && (s.includes('wage') || s.includes('pay'));
}

/** Overtime earnings */
export function isFormXXIIIGJOvertimeEarningsHeader(h) {
  const s = normFormXXIIIGJHeader(h);
  return s.includes('overtime') && s.includes('earning');
}

/** Date on which overtime wages paid */
export function isFormXXIIIGJOtWagesPaidDateHeader(h) {
  const s = normFormXXIIIGJHeader(h);
  return (
    (s.includes('overtime') && s.includes('paid')) ||
    /dates?\s+on\s+which\s+overtime\s+wage/.test(s)
  );
}

export function isFormXXIIIGJOtNilHeader(header) {
  return (
    isFormXXIIIGJOtWorkedDatesHeader(header) ||
    isFormXXIIIGJTotalOvertimeWorkedHeader(header) ||
    isFormXXIIIGJOvertimeRateHeader(header) ||
    isFormXXIIIGJOvertimeEarningsHeader(header) ||
    isFormXXIIIGJOtWagesPaidDateHeader(header)
  );
}

export function applyFormXXIIIGJOtNilToRow(row, headers, helpers = {}) {
  const hdrs = Array.isArray(headers) ? headers : [];
  const out = row && typeof row === 'object' ? { ...row } : {};
  const nilText = helpers.nilText != null ? String(helpers.nilText) : FORM_XXIII_GJ_OT_NIL;
  const { overwrite = true } = helpers;

  hdrs.forEach((header) => {
    if (!isFormXXIIIGJOtNilHeader(header)) return;
    const existing = String(out[header] ?? '').trim();
    if (
      !overwrite &&
      existing &&
      !/^enter\b/i.test(existing) &&
      !/^nil+$/i.test(existing) &&
      existing.toLowerCase() !== 'n/a' &&
      existing !== '-' &&
      existing !== '—'
    ) {
      return;
    }
    out[header] = nilText;
  });
  return out;
}

export function applyFormXXIIIGJOtNilToMappedRows(
  mappedData,
  headers,
  nilText = FORM_XXIII_GJ_OT_NIL,
  helpers = {}
) {
  if (!Array.isArray(mappedData)) return [];
  const { overwrite = true } = helpers;
  return mappedData.map((row) =>
    applyFormXXIIIGJOtNilToRow(row, headers, { nilText, overwrite })
  );
}
