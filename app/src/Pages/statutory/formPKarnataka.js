/**
 * Karnataka Form P — Notice of Maximum Leave Accumulated (See rule 20).
 * One notice letter per worker (Shri/Smt + Address + leave table) → ZIP of employee files.
 */

import JSZip from 'jszip';
import {
  buildFormOGJGujaratWorkbookWithTemplateStyles,
  getFormOGJRowValueForHeader,
  isFormOGJWorkerNameHeader,
  remapFormOGJGujaratRowsToHeaders,
  resolveFormOGJGujaratTableHeaders,
  toFormOGJPersonNameDisplay,
} from './formOGJGujarat';
import { formatFormFKarnatakaEmployeeName } from './formFKarnataka';
import {
  formatStatutoryHeaderLabelValueExport,
  resolveHeaderFieldExportValue,
} from '../../utils/statutorySiteCompanyHeaders';

export { isFormPKarnatakaAccumulatedLeaveContext } from './formOGJGujarat';

function unwrapEmployee(empItem) {
  if (!empItem || typeof empItem !== 'object') return null;
  return empItem.Employee || empItem.employee || empItem;
}

function excelCellText(value) {
  if (value == null || value === '') return '';
  if (typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') {
    return String(value).trim();
  }
  if (typeof value === 'object') {
    if (Array.isArray(value.richText)) {
      return value.richText.map((p) => p?.text || '').join('').trim();
    }
    if (value.text != null) return String(value.text).trim();
    if (value.result != null) return excelCellText(value.result);
  }
  const text = String(value).trim();
  return text === '[object Object]' ? '' : text;
}

export function resolveFormPKarnatakaEmployeeName(empItem, row = null) {
  const emp = unwrapEmployee(empItem);
  const fromEmp = formatFormFKarnatakaEmployeeName(emp || {});
  if (fromEmp) return fromEmp;
  if (!row || typeof row !== 'object') return '';
  const headers = Object.keys(row);
  const nameHeader = headers.find((h) => isFormOGJWorkerNameHeader(h));
  const raw = String(
    row.__employeeLookupName ||
      (nameHeader ? row[nameHeader] : '') ||
      row['Name of Workers'] ||
      row['Employee Name'] ||
      row.Name ||
      ''
  ).trim();
  return toFormOGJPersonNameDisplay(raw);
}

export function resolveFormPKarnatakaEmployeeAddress(empItem, row = null) {
  const emp = unwrapEmployee(empItem) || {};
  const bags = emp !== empItem && empItem && typeof empItem === 'object' ? [emp, empItem] : [emp];
  const keys = [
    'Permanent Address',
    'PermanentAddress',
    'Present Address',
    'PresentAddress',
    'Postal Address',
    'PostalAddress',
    'Address',
    'address',
    'Current Address',
    'Residence Address',
  ];
  for (const bag of bags) {
    if (!bag || typeof bag !== 'object') continue;
    for (const key of keys) {
      const raw = bag[key];
      if (raw == null || raw === '') continue;
      if (typeof raw === 'object') {
        const nested = String(
          raw.display_value ?? raw.displayValue ?? raw.name ?? raw.value ?? ''
        ).trim();
        if (nested && nested !== '[object Object]') return nested;
        continue;
      }
      const text = String(raw).trim();
      if (text && text !== '[object Object]') return text;
    }
  }
  return String(row?.Address || row?.['Permanent Address'] || row?.['Postal Address'] || '').trim();
}

/** 2nd-model To line: "Shri/Smt. Sornaraja Gurusamy" (one line, no dots / Name of worker). */
export function fillFormPKarnatakaShriSmtToLine(raw, name) {
  const text = String(name || '').trim();
  if (!text) return raw;
  const source = String(raw || '');
  const prefixMatch = source.match(/shri\s*\/?\s*smt\.?/i);
  let prefix = prefixMatch ? prefixMatch[0].replace(/\s+/g, ' ').trim() : 'Shri/Smt.';
  // Normalize "Shri / Smt." → "Shri/Smt."
  prefix = prefix.replace(/\s*\/\s*/g, '/').replace(/\s+/g, '');
  if (!/\.$/.test(prefix)) prefix = `${prefix}.`;
  return `${prefix} ${text}`.replace(/[ \t]{2,}/g, ' ').trim();
}

/** True when the Shri/Smt cell is still a blank template (dots / Name of worker). */
export function isFormPKarnatakaShriSmtPlaceholder(text) {
  const t = String(text || '')
    .replace(/\u2026/g, '.')
    .replace(/[.\u00b7\u2022_\s]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
  if (!t) return true;
  if (!/shri/.test(t) || !/smt/.test(t)) return false;
  const withoutLabel = t
    .replace(/shri\s*\/?\s*smt\.?/g, '')
    .replace(/name\s+of\s+(the\s+)?worker/g, '')
    .replace(/[()]/g, '')
    .trim();
  return withoutLabel === '';
}

export function resolveFormPKarnatakaEstablishmentText(headerFormData = {}) {
  if (!headerFormData || typeof headerFormData !== 'object') return '';
  const nameAndAddress = resolveHeaderFieldExportValue(headerFormData, {
    key: 'statutory_establishment_name_address',
    label: 'Name and address of the establishment',
  });
  if (nameAndAddress) return nameAndAddress;
  return resolveHeaderFieldExportValue(headerFormData, {
    key: 'statutory_establishment_address',
    label: 'Address of the Establishment',
  });
}

export function fillFormPKarnatakaEstablishmentLine(raw, value) {
  const text = String(value || '').trim();
  if (!text) return raw;
  const labelOnly = String(raw || '').split(':')[0].trim();
  const labelMatch = labelOnly.match(
    /^(name\s+and\s+address\s+of\s+(?:the\s+)?establishment|address\s+of\s+(?:the\s+)?establishment)/i
  );
  const label = labelMatch ? labelMatch[1].replace(/\s+/g, ' ').trim() : 'Name and address of the establishment';
  return formatStatutoryHeaderLabelValueExport(label, label, text);
}

function isFormPKarnatakaEstablishmentLabel(text) {
  const n = String(text || '').toLowerCase();
  if (/authorised person|authorized person/i.test(n)) return false;
  return (
    /name\s+and\s+address\s+of\s+(?:the\s+)?establishment/i.test(n) ||
    /address\s+of\s+(?:the\s+)?establishment/i.test(n)
  );
}

export function fillFormPKarnatakaToAddressLine(raw, address) {
  const text = String(address || '').trim();
  if (!text) return raw;
  if (isFormPKarnatakaEstablishmentLabel(raw)) return raw;
  return `Address: ${text}`;
}

function isFormPKarnatakaLegalParagraph(text) {
  return /it is hereby informed|maximum leave that can be accumulated|details of the leave accumulated/i.test(
    String(text || '')
  );
}

function applyFormPKarnatakaToLineLayout(worksheet, cell, row, col) {
  cell.alignment = {
    ...(cell.alignment || {}),
    wrapText: false,
    vertical: 'middle',
    horizontal: 'left',
  };
  const wsRow = typeof worksheet.getRow === 'function' ? worksheet.getRow(row) : null;
  if (wsRow && Number(wsRow.height) > 22) wsRow.height = 18;
  try {
    if (typeof worksheet.mergeCells === 'function') {
      worksheet.mergeCells(row, col, row, Math.max(col + 3, 6));
    }
  } catch (_) {
    /* template already merged this band */
  }
}

/**
 * Keep letterhead + To block on one line each (2nd Excel model). Legal paragraph still wraps.
 */
export function freezeFormPKarnatakaNoticeLetterheadLayout(worksheet) {
  if (!worksheet) return;
  const maxRows = Math.min(20, Number(worksheet.rowCount) || 20);
  for (let r = 1; r <= maxRows; r += 1) {
    for (let c = 1; c <= 8; c += 1) {
      const cell = worksheet.getCell(r, c);
      const raw = excelCellText(cell.value);
      if (!raw || isFormPKarnatakaLegalParagraph(raw)) continue;
      if (
        isFormPKarnatakaEstablishmentLabel(raw) ||
        /authorised person|authorized person/i.test(raw) ||
        /^to,?\s*$/i.test(raw.trim()) ||
        /shri\s*\/?\s*smt/i.test(raw) ||
        (/^address\b/i.test(raw.trim()) && !/establishment/i.test(raw))
      ) {
        applyFormPKarnatakaToLineLayout(worksheet, cell, r, c);
      }
    }
  }
}

/**
 * Fill establishment (fetched) + To block. Row 10 Address uses the same establishment address.
 */
export function writeFormPKarnatakaNoticeIdentity(worksheet, empItem, row = null, options = {}) {
  if (!worksheet) return 0;
  const name = resolveFormPKarnatakaEmployeeName(empItem, row);
  const establishmentText = String(
    options.establishmentText || resolveFormPKarnatakaEstablishmentText(options.headerFormData) || ''
  ).trim();
  const address = establishmentText;
  const maxRows = Math.min(40, Number(worksheet.rowCount) || 40);
  let written = 0;
  let wroteEstablishment = false;
  let wroteShriSmt = false;
  let wroteAddress = false;
  for (let r = 1; r <= maxRows; r += 1) {
    for (let c = 1; c <= 12; c += 1) {
      const cell = worksheet.getCell(r, c);
      const raw = excelCellText(cell.value);
      if (!raw || isFormPKarnatakaLegalParagraph(raw)) continue;
      const n = raw.toLowerCase();
      if (/authorised person|authorized person/i.test(n)) {
        continue;
      }
      if (!wroteEstablishment && establishmentText && isFormPKarnatakaEstablishmentLabel(raw)) {
        const next = fillFormPKarnatakaEstablishmentLine(raw, establishmentText);
        if (next !== raw) {
          cell.value = next;
          written += 1;
        }
        applyFormPKarnatakaToLineLayout(worksheet, cell, r, c);
        wroteEstablishment = true;
        continue;
      }
      if (
        !wroteShriSmt &&
        name &&
        /shri\s*\/?\s*smt/i.test(n) &&
        (isFormPKarnatakaShriSmtPlaceholder(raw) || !new RegExp(name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'i').test(raw))
      ) {
        const next = fillFormPKarnatakaShriSmtToLine(raw, name);
        if (next !== raw) {
          cell.value = next;
          written += 1;
        }
        applyFormPKarnatakaToLineLayout(worksheet, cell, r, c);
        wroteShriSmt = true;
        continue;
      }
      if (
        !wroteAddress &&
        address &&
        /^address\b/i.test(n.trim()) &&
        !/establishment/i.test(n)
      ) {
        const next = fillFormPKarnatakaToAddressLine(raw, address);
        if (next !== raw) {
          cell.value = next;
          written += 1;
        }
        applyFormPKarnatakaToLineLayout(worksheet, cell, r, c);
        wroteAddress = true;
      }
    }
  }
  freezeFormPKarnatakaNoticeLetterheadLayout(worksheet);
  return written;
}

function formPKarnatakaRowIdentity(row) {
  if (!row || typeof row !== 'object') return { id: '', name: '' };
  return {
    id: String(row.__employeeLookupId || row['Employee ID'] || row.EmployeeID || '')
      .trim()
      .toLowerCase(),
    name: String(
      row.__employeeLookupName ||
        getFormOGJRowValueForHeader(row, 'Name of Workers') ||
        row['Name of Workers'] ||
        ''
    )
      .trim()
      .toLowerCase(),
  };
}

function formPKarnatakaEmployeeIdentity(empItem) {
  const emp = unwrapEmployee(empItem);
  if (!emp || typeof emp !== 'object') return { id: '', name: '' };
  const id = String(
    emp.Employee_ID || emp['Employee ID'] || emp.employeeId || emp.EmployeeID || emp.Zoho_ID || ''
  )
    .trim()
    .toLowerCase();
  const name = resolveFormPKarnatakaEmployeeName(empItem).toLowerCase();
  return { id, name };
}

function identitiesMatch(rowId, rowName, empId, empName) {
  if (rowId && empId && rowId === empId) return true;
  if (rowName && empName && rowName === empName) return true;
  return false;
}

/**
 * Stamp the leave table body from mapped autofill (days / From / Till).
 * Survives fragile column detection and clears template shift-time residue.
 */
export function writeFormPKarnatakaLeaveTableRow(worksheet, row = null, options = {}) {
  if (!worksheet || !row || typeof row !== 'object') return 0;
  const formatDate =
    typeof options.formatStatutoryDateDisplay === 'function'
      ? options.formatStatutoryDateDisplay
      : (v) => String(v || '').trim();

  const leaveCount = String(
    getFormOGJRowValueForHeader(row, 'Number of accumulated leave') ||
      row['Number of accumulated leave'] ||
      ''
  ).trim();
  const periodFrom = String(
    getFormOGJRowValueForHeader(row, 'Period for which leave is accumulated_From') ||
      row['Period for which leave is accumulated_From'] ||
      row['Period for which leave is accumulated From_From'] ||
      ''
  ).trim();
  const periodTill = String(
    getFormOGJRowValueForHeader(row, 'Period for which leave is accumulated_Till') ||
      row['Period for which leave is accumulated_Till'] ||
      row['Period for which leave is accumulated Till_Till'] ||
      ''
  ).trim();
  if (!leaveCount && !periodFrom && !periodTill) return 0;

  const maxRows = Math.min(60, Number(worksheet.rowCount) || 60);
  let headerRow = -1;
  let snoCol = -1;
  let leaveCol = -1;
  let fromCol = -1;
  let tillCol = -1;

  for (let r = 1; r <= maxRows; r += 1) {
    const parts = [];
    for (let c = 1; c <= 12; c += 1) {
      const t = excelCellText(worksheet.getCell(r, c).value);
      if (t) parts.push(t.toLowerCase());
    }
    const joined = parts.join(' ');
    if (!/sr\.?\s*no/.test(joined) || !/accumulated\s+leave/.test(joined)) continue;
    if (/details of the leave/.test(joined) && !/number\s+of\s+accumulated/.test(joined)) continue;
    headerRow = r;
    for (let c = 1; c <= 12; c += 1) {
      const t = excelCellText(worksheet.getCell(r, c).value).toLowerCase();
      if (/^sr\.?\s*no/.test(t)) snoCol = c;
      if (/number\s+of\s+accumulated\s+leave/.test(t) || (/accumulated\s+leave/.test(t) && !/period|perod/.test(t))) {
        leaveCol = c;
      }
      if (/period|perod/.test(t) && /accumulated/.test(t)) {
        if (fromCol < 0) fromCol = c;
        tillCol = Math.max(tillCol, c + 1);
      }
    }
    break;
  }
  if (headerRow < 0 || snoCol < 0) return 0;

  for (let r = headerRow; r <= Math.min(headerRow + 4, maxRows); r += 1) {
    for (let c = 1; c <= 12; c += 1) {
      const t = excelCellText(worksheet.getCell(r, c).value).toLowerCase().trim();
      if (t === 'from') fromCol = c;
      if (t === 'till' || t === 'to') tillCol = c;
    }
  }
  if (leaveCol < 0) leaveCol = snoCol + 1;
  if (fromCol < 0) fromCol = leaveCol + 1;
  if (tillCol < 0 || tillCol === fromCol) tillCol = fromCol + 1;

  let dataRow = headerRow + 2;
  for (let r = headerRow; r <= Math.min(headerRow + 5, maxRows); r += 1) {
    let hasFrom = false;
    let hasTill = false;
    for (let c = 1; c <= 12; c += 1) {
      const t = excelCellText(worksheet.getCell(r, c).value).toLowerCase().trim();
      if (t === 'from') hasFrom = true;
      if (t === 'till' || t === 'to') hasTill = true;
    }
    if (hasFrom || hasTill) dataRow = r + 1;
  }

  const stamp = (col, value) => {
    if (col < 1) return;
    const cell = worksheet.getCell(dataRow, col);
    cell.numFmt = '@';
    cell.value = value == null || value === '' ? '' : String(value);
    cell.alignment = {
      ...(cell.alignment || {}),
      vertical: 'middle',
      horizontal: 'center',
      wrapText: true,
    };
  };

  stamp(snoCol, '1');
  stamp(leaveCol, leaveCount);
  stamp(fromCol, formatDate(periodFrom));
  stamp(tillCol, formatDate(periodTill));

  // Clear trailing body rows so leftover template times (09:00 AM) cannot leak into PDF.
  for (let r = dataRow + 1; r <= dataRow + 4; r += 1) {
    [snoCol, leaveCol, fromCol, tillCol].forEach((col) => {
      if (col < 1) return;
      worksheet.getCell(r, col).value = '';
    });
  }
  return 1;
}

export async function buildFormPKarnatakaWorkbookWithTemplateStyles(opts = {}) {
  const {
    employee = null,
    mappedData = [],
    headerFormData = null,
    formatStatutoryDateDisplay = null,
    ...rest
  } = opts;
  const rows = Array.isArray(mappedData) ? mappedData : [];
  const row = rows[0] || {};
  const establishmentText = resolveFormPKarnatakaEstablishmentText(headerFormData);
  return buildFormOGJGujaratWorkbookWithTemplateStyles({
    ...rest,
    formatStatutoryDateDisplay,
    // Generic header writer merges/wraps Address / Name into the To block (1st screenshot).
    headerFormData: null,
    mappedData: rows,
    afterWorksheetReady: (worksheet) => {
      writeFormPKarnatakaNoticeIdentity(worksheet, employee, row, { establishmentText, headerFormData });
      writeFormPKarnatakaLeaveTableRow(worksheet, row, { formatStatutoryDateDisplay });
    },
  });
}

function resolveFormPKarnatakaEmployeeDownloadBaseName(row, empItem, fallbackIndex = 0) {
  const fromRow = String(
    row?.__employeeLookupName ||
      getFormOGJRowValueForHeader(row || {}, 'Name of Workers') ||
      row?.['Name of Workers'] ||
      ''
  ).trim();
  const name = fromRow || resolveFormPKarnatakaEmployeeName(empItem, row);
  const slug = String(name || '')
    .replace(/[^\w\s.-]/g, '')
    .replace(/\s+/g, '_')
    .replace(/_+/g, '_')
    .replace(/^_|_$/g, '')
    .slice(0, 48);
  return slug || `Employee_${fallbackIndex + 1}`;
}

function allocateUniqueFormPKarnatakaDownloadFileName(baseName, usedNames, fileNamePrefix = 'Form_P_Karnataka') {
  const count = usedNames.get(baseName) || 0;
  usedNames.set(baseName, count + 1);
  const suffix = count > 0 ? `_${count + 1}` : '';
  const prefix = String(fileNamePrefix || 'Form_P_Karnataka').replace(/\.xlsx?$/i, '');
  return `${prefix}_${baseName}${suffix}.xlsx`;
}

/**
 * One Form P notice per employee. Single employee → one .xlsx; multiple → ZIP.
 * @param {string} [fileNamePrefix='Form_P_Karnataka'] — also used for MH (`Form_P_Maharashtra`).
 */
export async function buildFormPKarnatakaPerEmployeeDownload({
  templateArrayBuffer,
  mappedData = [],
  headersToUse = [],
  parsedFormHeader,
  formFileName,
  headerFormData,
  parsedHeaderRowIndex,
  parsedTableStartCol,
  employeesOverride = [],
  formatStatutoryDateDisplay = null,
  fileNamePrefix = 'Form_P_Karnataka',
}) {
  if (!templateArrayBuffer) {
    throw new Error(
      'Original Form P template could not be loaded. Open Autofill again, then Download.'
    );
  }

  const hdrs = resolveFormOGJGujaratTableHeaders(headersToUse);
  const employees = Array.isArray(employeesOverride) ? employeesOverride : [];
  const tableRows = remapFormOGJGujaratRowsToHeaders(
    Array.isArray(mappedData) ? mappedData : [],
    hdrs,
    hdrs
  );
  const exportEntries = [];
  if (tableRows.length > 0) {
    tableRows.forEach((row, index) => {
      if (!row || typeof row !== 'object') return;
      const emp =
        employees.length === tableRows.length
          ? employees[index]
          : employees.find((empItem) => {
              const { id: rowId, name: rowName } = formPKarnatakaRowIdentity(row);
              const { id, name } = formPKarnatakaEmployeeIdentity(empItem);
              return identitiesMatch(rowId, rowName, id, name);
            }) ||
            employees[index] ||
            null;
      const name = resolveFormPKarnatakaEmployeeName(emp, row);
      if (name && !String(row.__employeeLookupName || '').trim()) {
        row.__employeeLookupName = name;
      }
      exportEntries.push({ row, emp, index });
    });
  } else if (employees.length > 0) {
    employees.forEach((emp, index) => {
      const name = resolveFormPKarnatakaEmployeeName(emp, null);
      exportEntries.push({
        row: name ? { __employeeLookupName: name } : {},
        emp,
        index,
      });
    });
  }

  const prefix = String(fileNamePrefix || 'Form_P_Karnataka').replace(/\.xlsx?$/i, '');
  const zipBase = String(formFileName || parsedFormHeader?.title || prefix)
    .replace(/\.xlsx?$/i, '')
    .replace(/[^a-zA-Z0-9._-]+/g, '_');
  const templateBytes =
    templateArrayBuffer instanceof ArrayBuffer
      ? templateArrayBuffer.slice(0)
      : templateArrayBuffer;
  const workbookArgs = {
    headersToUse: hdrs,
    parsedFormHeader,
    headerFormData,
    parsedHeaderRowIndex,
    parsedTableStartCol,
    formatStatutoryDateDisplay,
  };

  if (exportEntries.length <= 1) {
    const entry = exportEntries[0];
    const rows = entry?.row ? [entry.row] : [];
    const baseName = resolveFormPKarnatakaEmployeeDownloadBaseName(
      entry?.row,
      entry?.emp,
      entry?.index || 0
    );
    return buildFormPKarnatakaWorkbookWithTemplateStyles({
      ...workbookArgs,
      templateArrayBuffer:
        templateBytes instanceof ArrayBuffer ? templateBytes.slice(0) : templateBytes,
      mappedData: rows,
      employee: entry?.emp || null,
      formFileName: allocateUniqueFormPKarnatakaDownloadFileName(baseName, new Map(), prefix),
    });
  }

  const zip = new JSZip();
  const usedNames = new Map();
  for (let i = 0; i < exportEntries.length; i += 1) {
    const entry = exportEntries[i];
    const baseName = resolveFormPKarnatakaEmployeeDownloadBaseName(entry.row, entry.emp, entry.index);
    const entryName = allocateUniqueFormPKarnatakaDownloadFileName(baseName, usedNames, prefix);
    const { blob } = await buildFormPKarnatakaWorkbookWithTemplateStyles({
      ...workbookArgs,
      templateArrayBuffer:
        templateBytes instanceof ArrayBuffer ? templateBytes.slice(0) : templateBytes,
      mappedData: entry.row ? [entry.row] : [],
      employee: entry.emp || null,
      formFileName: entryName,
    });
    const xlsxBytes =
      typeof blob?.arrayBuffer === 'function'
        ? new Uint8Array(await blob.arrayBuffer())
        : new Uint8Array(await new Response(blob).arrayBuffer());
    zip.file(entryName, xlsxBytes);
    if (i > 0 && i % 5 === 0) {
      await new Promise((resolve) => setTimeout(resolve, 0));
    }
  }

  return {
    blob: await zip.generateAsync({
      type: 'blob',
      mimeType: 'application/zip',
      compression: 'STORE',
    }),
    fileName: `${zipBase}_Employees.zip`,
  };
}
