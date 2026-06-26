'use strict';

const express = require('express');
const catalystSDK = require('zcatalyst-sdk-node');
const fileUpload = require('express-fileupload');
const fs = require('fs');
const os = require('os');
const path = require('path');
const { buildProofDataFromZoho, locateProofTemplateTable } = require('./zohoProofMerge');
let XLSX;
try {
  // Prefer style-preserving library for template borders
  XLSX = require('xlsx-js-style');
} catch (err) {
  XLSX = require('xlsx');
  // Fallback works for read/write; optional xlsx-js-style improves template border styles
}

const app = express();

app.use(express.json());
app.use(fileUpload());

// Attach catalyst instance per request
app.use((req, res, next) => {
  try {
    const catalyst = catalystSDK.initialize(req);
    res.locals.catalyst = catalyst;
    next();
  } catch (err) {
    res.status(500).send({ status: 'failure', message: 'Catalyst init failed' });
  }
});

// Health check
app.get('/', (req, res) => {
  res.status(200).json({ status: 'success', message: 'statutoryreg_function ready' });
});

/** Normalize file store IDs for Data Store (string or null). */
function normalizeFileRef(v) {
  if (v == null || v === '') return null;
  const s = String(v).trim();
  if (s === '' || s === 'null' || s === 'undefined') return null;
  return s;
}

/**
 * Catalyst sometimes returns INVALID_INPUT / "Invalid input value for column name" when the
 * Statutory table is missing optional columns. Strip layers and retry.
 */
function isDatastoreColumnError(err) {
  const msg = String(err?.message || '').toLowerCase();
  return (
    err?.code === 'INVALID_INPUT' ||
    msg.includes('column name') ||
    msg.includes('invalid column') ||
    msg.includes('no such column')
  );
}

const OPTIONAL_STATUTORY_COLUMN_LAYERS = [
  ['SubmittedDate', 'ApprovedDate'],
  ['MonthFilter', 'Sector', 'State', 'Site'],
  ['SampleFile', 'SampleFileName'],
  ['Autofill', 'Draft'],
  ['Approval', 'Status'],
  ['SendForApproval', 'Remarks']
];

/** Normalize Catalyst date column values to YYYY-MM-DD. */
function normalizeStatutoryDateValue(v) {
  if (v == null || String(v).trim() === '') return null;
  const s = String(v).trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.substring(0, 10);
  const d = new Date(s);
  if (!Number.isNaN(d.getTime())) return d.toISOString().substring(0, 10);
  return s;
}

function todayIsoDate() {
  return new Date().toISOString().substring(0, 10);
}

function statutoryMetaFromBody(body, key) {
  if (!body || typeof body !== 'object') return undefined;
  const pascal = key.charAt(0).toUpperCase() + key.slice(1);
  if (body[key] !== undefined) return body[key];
  if (body[pascal] !== undefined) return body[pascal];
  return undefined;
}

function normalizeStatutoryMetaValue(v) {
  if (v == null) return null;
  const s = String(v).trim();
  if (!s || s === 'null' || s === 'undefined') return null;
  return s;
}

/** Accept submittedDate (preferred) or legacy draftDate from API clients. */
function submittedDateFromBody(body) {
  if (!body || typeof body !== 'object') return undefined;
  if (body.submittedDate !== undefined) return body.submittedDate;
  if (body.SubmittedDate !== undefined) return body.SubmittedDate;
  if (body.draftDate !== undefined) return body.draftDate;
  if (body.DraftDate !== undefined) return body.DraftDate;
  return undefined;
}

/** Accept approvedDate (preferred) or legacy approvalDate from API clients. */
function approvedDateFromBody(body) {
  if (!body || typeof body !== 'object') return undefined;
  if (body.approvedDate !== undefined) return body.approvedDate;
  if (body.ApprovedDate !== undefined) return body.ApprovedDate;
  if (body.approvalDate !== undefined) return body.approvalDate;
  if (body.ApprovalDate !== undefined) return body.ApprovalDate;
  return undefined;
}

function mapStatutoryRowToApi(row) {
  if (!row) return null;
  return {
    id: row.ROWID,
    formName: row.FormName || '',
    act: row.Act || '',
    description: row.Description || '',
    dueDate: row.DueDate || '',
    sector: row.Sector || '',
    state: row.State || '',
    site: row.Site || '',
    monthfilter: row.MonthFilter || '',
    autofill: row.Autofill || '',
    draft: row.Draft || '',
    formFile: row.FormFile || null,
    formFileName: row.FormFileName || null,
    sampleFile: row.SampleFile || row.FormFile || null,
    sampleFileName: row.SampleFileName || row.FormFileName || null,
    proofSubmissionFile: row.ProofSubmissionFile || null,
    proofSubmissionFileName: row.ProofSubmissionFileName || null,
    draftFile: row.DraftFile || null,
    draftFileName: row.DraftFileName || null,
    submittedDate: row.SubmittedDate || row.DraftDate || '',
    approvedDate: row.ApprovedDate || row.ApprovalDate || '',
    approval: row.Approval || '',
    status: row.Status || '',
    sendForApproval: row.SendForApproval || '',
    remarks: row.Remarks || '',
    createdTime: row.CREATEDTIME,
    modifiedTime: row.MODIFIEDTIME
  };
}

async function insertStatutoryRow(table, insertData) {
  const data = { ...insertData };
  let lastErr;
  for (let layer = 0; layer <= OPTIONAL_STATUTORY_COLUMN_LAYERS.length; layer++) {
    try {
      return await table.insertRow(data);
    } catch (err) {
      lastErr = err;
      if (!isDatastoreColumnError(err) || layer >= OPTIONAL_STATUTORY_COLUMN_LAYERS.length) {
        throw err;
      }
      const strip = OPTIONAL_STATUTORY_COLUMN_LAYERS[layer];
      for (const k of strip) {
        delete data[k];
      }
      console.warn(`Statutory insertRow: retrying without [${strip.join(', ')}]:`, err.message);
    }
  }
  throw lastErr;
}

async function updateStatutoryRow(table, updateData) {
  const data = { ...updateData };
  let lastErr;
  for (let layer = 0; layer <= OPTIONAL_STATUTORY_COLUMN_LAYERS.length; layer++) {
    try {
      await table.updateRow(data);
      return;
    } catch (err) {
      lastErr = err;
      if (!isDatastoreColumnError(err) || layer >= OPTIONAL_STATUTORY_COLUMN_LAYERS.length) {
        throw err;
      }
      const strip = OPTIONAL_STATUTORY_COLUMN_LAYERS[layer];
      for (const k of strip) {
        delete data[k];
      }
      console.warn(`Statutory updateRow: retrying without [${strip.join(', ')}]:`, err.message);
    }
  }
  throw lastErr;
}

/** Statutory UI "Returned" / approver reject — Status is often stored as Rejected. */
function isReturnedStatutoryDecision(status, approval) {
  const s = String(status || '').trim().toLowerCase();
  const a = String(approval || '').trim().toLowerCase();
  return (
    s === 'rejected' ||
    s === 'returned' ||
    a === 'rejected' ||
    a === 'reject'
  );
}

/**
 * When a statutory line is returned/rejected with remarks, copy snapshot into Data Store table "Returned".
 * Non-fatal: statutory update still succeeds if Returned insert fails.
 */
async function persistReturnedTableRow(catalyst, source) {
  const remarks = String(source?.remarks ?? source?.Remarks ?? '').trim();
  if (!remarks) return null;

  const siteName =
    source?.site != null ? String(source.site).trim() : String(source?.Site ?? '').trim();
  const monthFilter = resolveMonthFilterFromSource(source);

  try {
    const returnedTable = catalyst.datastore().table('Returned');
    const row = {
      Site: siteName || null,
      MonthFilter: monthFilter,
      Act: source?.act != null ? String(source.act).trim() : String(source?.Act ?? '').trim(),
      Description:
        source?.description != null
          ? String(source.description).trim()
          : String(source?.Description ?? '').trim(),
      FormName:
        source?.formName != null
          ? String(source.formName).trim()
          : String(source?.FormName ?? '').trim(),
      Status: 'Returned',
      DueDate:
        source?.dueDate != null ? String(source.dueDate).trim() : String(source?.DueDate ?? '').trim(),
      Remarks: remarks
    };
    let inserted;
    try {
      inserted = await returnedTable.insertRow(row);
    } catch (err) {
      if (isDatastoreColumnError(err) && row.MonthFilter != null) {
        const { MonthFilter: _mf, ...withoutMonth } = row;
        inserted = await returnedTable.insertRow(withoutMonth);
        console.warn('Returned insert: retried without MonthFilter:', err.message);
      } else {
        throw err;
      }
    }
    console.log('Returned table row inserted:', inserted?.ROWID || inserted, {
      site: siteName,
      monthFilter
    });
    return inserted;
  } catch (err) {
    console.error('persistReturnedTableRow failed:', err.message || err);
    return null;
  }
}

async function syncReturnedTableIfApplicable(catalyst, statutoryRow, context = {}) {
  if (!statutoryRow || !isReturnedStatutoryDecision(statutoryRow.Status, statutoryRow.Approval)) {
    return null;
  }
  const reqBody = context.requestBody && typeof context.requestBody === 'object' ? context.requestBody : {};
  const existing = context.existingRecord && typeof context.existingRecord === 'object' ? context.existingRecord : {};
  const siteName = resolveSiteNameForReturned(statutoryRow, existing, reqBody);
  const monthFilter =
    trimMonthFilterForDatastore(statutoryRow.MonthFilter) ||
    resolveMonthFilterFromSource(reqBody) ||
    trimMonthFilterForDatastore(existing.MonthFilter);

  return persistReturnedTableRow(catalyst, {
    site: siteName,
    MonthFilter: monthFilter,
    act: pickNonEmptyText(reqBody.act, reqBody.Act, statutoryRow.Act, existing.Act),
    description: pickNonEmptyText(
      reqBody.description,
      reqBody.Description,
      statutoryRow.Description,
      existing.Description
    ),
    formName: pickNonEmptyText(
      reqBody.formName,
      reqBody.FormName,
      statutoryRow.FormName,
      existing.FormName
    ),
    dueDate: pickNonEmptyText(reqBody.dueDate, reqBody.DueDate, statutoryRow.DueDate, existing.DueDate),
    remarks: pickNonEmptyText(statutoryRow.Remarks, reqBody.remarks, reqBody.Remarks)
  });
}

const MONTH_NAMES = [
  'January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December'
];

function deriveMonthFilterFromDueDate(dueDate) {
  if (dueDate == null || String(dueDate).trim() === '') return null;
  const s = String(dueDate).trim();
  const lower = s.toLowerCase();
  for (let i = 0; i < MONTH_NAMES.length; i++) {
    const full = MONTH_NAMES[i].toLowerCase();
    const ab = full.slice(0, 3);
    if (lower.includes(full) || new RegExp(`\\b${ab}[a-z]*`, 'i').test(s)) {
      return MONTH_NAMES[i];
    }
  }
  const m = s.match(
    /\b(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*(?:[-/ ]+\d{4}|\d{4})?\b/i
  );
  if (m) {
    const abbr = m[1].slice(0, 3).toLowerCase();
    const idx = MONTH_NAMES.findIndex((name) => name.slice(0, 3).toLowerCase() === abbr);
    if (idx >= 0) return MONTH_NAMES[idx];
  }
  return null;
}

function resolveMonthFilterForSampleDataSave(payload, statutoryRow) {
  const fromPayload = resolveMonthFilterFromSource(payload);
  if (fromPayload) return fromPayload;
  const fromStatutory = trimMonthFilterForDatastore(statutoryRow?.MonthFilter);
  if (fromStatutory) return fromStatutory;
  const fromDue = deriveMonthFilterFromDueDate(statutoryRow?.DueDate);
  if (fromDue) return fromDue;
  return MONTH_NAMES[new Date().getMonth()];
}

/** Save current statutory grid snapshot into SampleData table. */
async function persistSampleDataSnapshot(catalyst, payload, statutoryRow = null) {
  const headers = Array.isArray(payload?.headers) ? payload.headers : [];
  const rows = Array.isArray(payload?.rows) ? payload.rows : [];
  const headerFormData =
    payload?.headerFormData && typeof payload.headerFormData === 'object' && !Array.isArray(payload.headerFormData)
      ? payload.headerFormData
      : {};
  const statutoryId =
    payload?.statutoryId != null ? String(payload.statutoryId).trim() : '';
  const hasUsableSnapshot = headers.length > 0 || rows.length > 0 || Object.keys(headerFormData).length > 0;
  if (!statutoryId || !hasUsableSnapshot) {
    return { saved: false, reason: 'empty_payload' };
  }
  const monthFilterValue = resolveMonthFilterForSampleDataSave(payload, statutoryRow);
  const table = catalyst.datastore().table('SampleData');
  const headerPayload = {
    formName: payload?.formName || null,
    statutoryId,
    monthfilter: monthFilterValue,
    headers,
    headerFormData
  };
  const dataPayload = {
    rows
  };
  const insertRow = {
    Header: JSON.stringify(headerPayload),
    Data: JSON.stringify(dataPayload),
    MonthFilter: monthFilterValue
  };
  let insertResp;
  try {
    insertResp = await table.insertRow(insertRow);
  } catch (err) {
    if (isDatastoreColumnError(err) && insertRow.MonthFilter != null) {
      const { MonthFilter: _mf, ...withoutMonth } = insertRow;
      insertResp = await table.insertRow(withoutMonth);
      console.warn('SampleData insert: retried without MonthFilter column:', err.message);
    } else {
      console.error('SampleData insert failed:', err.message || err, {
        statutoryId,
        rowCount: rows.length,
        monthFilter: monthFilterValue
      });
      throw err;
    }
  }
  let statutoryDataResult = { saved: false, inserted: 0 };
  try {
    statutoryDataResult = await persistStatutoryDataFields(
      catalyst,
      {
        ...payload,
        headers,
        rows,
        headerFormData,
        monthfilter: monthFilterValue,
        monthFilter: monthFilterValue,
        MonthFilter: monthFilterValue
      },
      statutoryRow
    );
  } catch (statDataErr) {
    console.warn('StatutoryData field save failed:', statDataErr?.message || statDataErr);
  }

  return {
    saved: true,
    rowId: insertResp?.ROWID || null,
    monthFilter: monthFilterValue,
    statutoryDataSaved: !!statutoryDataResult.saved,
    statutoryDataInserted: statutoryDataResult.inserted || 0
  };
}

const STATUTORY_DATA_TABLE = 'StatutoryData';

function pickStatutoryDataRow(entry) {
  if (!entry || typeof entry !== 'object') return null;
  return entry.StatutoryData || entry.statutorydata || entry;
}

function normalizeStatutoryDataHeaderKey(key) {
  return String(key || '')
    .replace(/[\r\n]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function statutoryDataHeaderKeysMatch(a, b) {
  return (
    normalizeStatutoryDataHeaderKey(a).toLowerCase() ===
    normalizeStatutoryDataHeaderKey(b).toLowerCase()
  );
}

function findStatutoryDataHeaderIndex(headers, targetHeader) {
  if (!Array.isArray(headers)) return -1;
  const exact = headers.findIndex((header) => statutoryDataHeaderKeysMatch(header, targetHeader));
  if (exact >= 0) return exact;
  const targetNorm = normalizeStatutoryDataHeaderForMatch(targetHeader);
  return headers.findIndex((header) => {
    const raw = String(header || '');
    const underscoreIdx = raw.indexOf('_');
    if (underscoreIdx < 0) return false;
    const sub = raw.slice(underscoreIdx + 1).trim();
    return sub && normalizeStatutoryDataHeaderForMatch(sub) === targetNorm;
  });
}

function findStatutoryDataRowObjectKey(row, targetHeader) {
  if (!row || typeof row !== 'object' || Array.isArray(row)) return null;
  const keys = Object.keys(row);
  const match = keys.find((key) => statutoryDataHeaderKeysMatch(key, targetHeader));
  if (match) return match;
  const targetNorm = normalizeStatutoryDataHeaderForMatch(targetHeader);
  return (
    keys.find((key) => {
      const raw = String(key || '');
      const underscoreIdx = raw.indexOf('_');
      if (underscoreIdx < 0) return false;
      const sub = raw.slice(underscoreIdx + 1).trim();
      return sub && normalizeStatutoryDataHeaderForMatch(sub) === targetNorm;
    }) || null
  );
}

const STATUTORY_ROMAN_PART_TO_DIGIT = {
  i: '1',
  ii: '2',
  iii: '3',
  iv: '4',
  v: '5',
  vi: '6',
  vii: '7',
  viii: '8',
  ix: '9',
  x: '10'
};

function extractStatutoryFormPartToken(formName, extraText = '') {
  const blob = `${formName || ''} ${extraText || ''}`
    .toLowerCase()
    .replace(/[^a-z0-9 ]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  const m = blob.match(/\bpart\s*(?:no\.?|number)?\s*(\d+|[ivx]+)\b/);
  if (!m) return '';
  const token = String(m[1] || '').trim();
  if (/^\d+$/.test(token)) return String(parseInt(token, 10));
  return STATUTORY_ROMAN_PART_TO_DIGIT[token] || token;
}

function formatStatutoryFormPartSuffix(partToken) {
  if (!partToken) return '';
  if (/^\d+$/.test(String(partToken))) return `Part ${parseInt(String(partToken), 10)}`;
  return `Part ${String(partToken).toUpperCase()}`;
}

function canonicalStatutoryDataFormName(name) {
  const raw = String(name || '').trim();
  if (!raw) return '';
  const lower = raw.toLowerCase().replace(/\s+/g, ' ');
  const m = lower.match(/^form\s+([a-z0-9]+)\b/);
  if (!m) return raw;
  const suffix = m[1];
  const base = /^\d+$/.test(suffix) ? `Form ${suffix}` : `Form ${suffix.toUpperCase()}`;
  const partSuffix = formatStatutoryFormPartSuffix(extractStatutoryFormPartToken(raw));
  return partSuffix ? `${base} ${partSuffix}` : base;
}

function resolveStatutoryDataHeaderLabel(key, fieldDefinitions) {
  if (Array.isArray(fieldDefinitions)) {
    const match = fieldDefinitions.find((f) => String(f?.key || '') === String(key));
    if (match?.label) return normalizeStatutoryDataHeaderKey(match.label);
  }
  return normalizeStatutoryDataHeaderKey(key);
}

function normalizeStatutoryDataHeaderForMatch(header) {
  return String(header || '')
    .toLowerCase()
    .replace(/[^a-z0-9 ]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function isStatutoryDataEmployeeNameHeader(header) {
  const n = normalizeStatutoryDataHeaderForMatch(header);
  return (
    /name\s+of\s+the\s+employee/.test(n) ||
    /name\s+of\s+employee/.test(n) ||
    /name\s+of\s+the\s+worker/.test(n) ||
    /name\s+of\s+the\s+workman/.test(n) ||
    /name\s+of\s+the\s+woman/.test(n) ||
    /name\s+of\s+the\s+person/.test(n) ||
    /full\s+name/.test(n) ||
    /employee\s+name/.test(n) ||
    /worker\s+name/.test(n) ||
    /workman\s+name/.test(n) ||
    (n.includes('name') &&
      (n.includes('employee') ||
        n.includes('worker') ||
        n.includes('workman') ||
        n.includes('woman') ||
        n.includes('person') ||
        n.includes('staff') ||
        n.includes('member')) &&
      !(/\bage\b/.test(n) && n.includes('woman'))) ||
    n === 'name'
  );
}

function isStatutoryDataEmployeeIdHeader(header) {
  const n = normalizeStatutoryDataHeaderForMatch(header);
  return (
    /employee\s+id/.test(n) ||
    /number\s+in\s+register/.test(n) ||
    /register\s+number/.test(n) ||
    /emp\s*id/.test(n) ||
    /employee\s+identification/.test(n) ||
    /employee\s+number/.test(n) ||
    /employee\s+code/.test(n) ||
    /emp\s*code/.test(n) ||
    /staff\s+id/.test(n) ||
    /token\s+number/.test(n) ||
    /ticket\s+number/.test(n) ||
    (/worker/.test(n) && /identity/.test(n)) ||
    (/worker/.test(n) && /identify/.test(n)) ||
    /identification\s+no/.test(n) ||
    /identity\s+no/.test(n) ||
    ((/\bid\b/.test(n) || n.includes('code') || n.includes('number')) &&
      (n.includes('employee') ||
        n.includes('emp') ||
        n.includes('worker') ||
        n.includes('staff') ||
        n.includes('roll') ||
        n.includes('register')))
  );
}

function isStatutoryDataSerialOnlyHeader(header) {
  const n = normalizeStatutoryDataHeaderForMatch(header);
  return (
    /^s\.?\s*no/.test(n) ||
    /serial\s+number/.test(n) ||
    /sl\.?\s*no/.test(n) ||
    n === 'sno' ||
    n === 'serial no'
  );
}

function isStatutoryDataPrimaryValueHeader(header) {
  const n = normalizeStatutoryDataHeaderForMatch(header);
  return (
    (n.includes('total') && n.includes('earn')) ||
    (n.includes('normal') && n.includes('earn')) ||
    (n.includes('overtime') && n.includes('earn')) ||
    n.includes('total earnings') ||
    n.includes('overtime earnings') ||
    n.includes('gross wage') ||
    n.includes('net wage') ||
    n.includes('basic wage') ||
    n === 'remarks' ||
    n.includes('remark') ||
    n.includes('rate of remuneration') ||
    n.includes('emolument') ||
    n.includes('basic wage') ||
    n.includes('dearness') ||
    n.includes('house rent') ||
    n.includes('other allowance')
  );
}

function isStatutoryDataDesignationRowKeyHeader(header) {
  const n = normalizeStatutoryDataHeaderForMatch(header);
  return (
    /category\s+of\s+workers?/.test(n) ||
    /category\s+of\s+workmen?/.test(n) ||
    /^designation$/.test(n) ||
    (n.includes('designation') && !n.includes('nature')) ||
    /post\s+held/.test(n) ||
    /job\s+title/.test(n) ||
    /class\s+of\s+workers?/.test(n)
  );
}

function isStatutoryDataIdentityHeader(header) {
  const n = normalizeStatutoryDataHeaderForMatch(header);
  return (
    isStatutoryDataEmployeeNameHeader(header) ||
    isStatutoryDataEmployeeIdHeader(header) ||
    isStatutoryDataDesignationRowKeyHeader(header) ||
    isStatutoryDataSerialOnlyHeader(header)
  );
}

function resolveStatutoryDataRowCell(row, headers, colIdx, normalizedHeader) {
  const originalHeader = headers[colIdx] || normalizedHeader;
  let cellVal = row[originalHeader];
  if (cellVal == null) {
    const matchKey = findStatutoryDataRowObjectKey(row, normalizedHeader);
    if (matchKey) cellVal = row[matchKey];
  }
  if (cellVal == null || String(cellVal).trim() === '') return '';
  return String(cellVal).trim();
}

function resolveStatutoryDataEmployeeKeyFromRowValues(row) {
  if (!row || typeof row !== 'object' || Array.isArray(row)) return '';
  for (const [key, val] of Object.entries(row)) {
    if (String(key).startsWith('__')) continue;
    const text = String(val ?? '').trim();
    if (!text) continue;
    if (/^VE\d{3,}$/i.test(text)) return text;
  }
  return '';
}

function resolveStatutoryDataEmployeeKeyFromFirstTextColumn(row, headers) {
  for (let colIdx = 0; colIdx < headers.length; colIdx += 1) {
    const header = headers[colIdx];
    if (isStatutoryDataIdentityHeader(header)) continue;
    const val = resolveStatutoryDataRowCell(
      row,
      headers,
      colIdx,
      normalizeStatutoryDataHeaderKey(header)
    );
    if (!val || /^\d+(\.\d+)?$/.test(val)) continue;
    if (/[a-zA-Z]{2,}/.test(val) && val.length >= 3) return val;
  }
  return '';
}

function headersHaveStatutoryDataEmployeeIdentityColumn(headers) {
  return (Array.isArray(headers) ? headers : []).some(
    (h) => isStatutoryDataEmployeeIdHeader(h) || isStatutoryDataEmployeeNameHeader(h)
  );
}

function resolveStatutoryDataEmployeeKey(row, headers, rowIdx = -1, employeeOrder = null) {
  let employeeId = '';
  let employeeName = '';
  headers.forEach((header, colIdx) => {
    const normalizedHeader = normalizeStatutoryDataHeaderKey(header);
    const val = resolveStatutoryDataRowCell(row, headers, colIdx, normalizedHeader);
    if (!val) return;
    if (isStatutoryDataEmployeeIdHeader(header)) employeeId = employeeId || val;
    else if (isStatutoryDataEmployeeNameHeader(header)) employeeName = employeeName || val;
  });
  if (employeeId || employeeName) return employeeId || employeeName;

  if (
    !headersHaveStatutoryDataEmployeeIdentityColumn(headers) &&
    Number.isInteger(rowIdx) &&
    rowIdx >= 0 &&
    Array.isArray(employeeOrder) &&
    employeeOrder.length > 0
  ) {
    const ordered = String(employeeOrder[rowIdx] ?? '').trim();
    if (ordered) return ordered;
    return `ROW_${rowIdx + 1}`;
  }

  let designationKey = '';
  headers.forEach((header, colIdx) => {
    const normalizedHeader = normalizeStatutoryDataHeaderKey(header);
    const val = resolveStatutoryDataRowCell(row, headers, colIdx, normalizedHeader);
    if (!val) return;
    if (isStatutoryDataDesignationRowKeyHeader(header)) designationKey = designationKey || val;
  });
  if (designationKey && (!Array.isArray(employeeOrder) || employeeOrder.length === 0)) {
    return designationKey;
  }

  const fromRowValues = resolveStatutoryDataEmployeeKeyFromRowValues(row);
  if (fromRowValues) return fromRowValues;
  const fromTextColumn = resolveStatutoryDataEmployeeKeyFromFirstTextColumn(row, headers);
  if (fromTextColumn) return fromTextColumn;
  if (Number.isInteger(rowIdx) && rowIdx >= 0) return `ROW_${rowIdx + 1}`;
  return '';
}

function resolveStatutoryDataPrimaryValue(row, headers) {
  for (let colIdx = 0; colIdx < headers.length; colIdx += 1) {
    const header = headers[colIdx];
    if (!isStatutoryDataPrimaryValueHeader(header)) continue;
    const val = resolveStatutoryDataRowCell(
      row,
      headers,
      colIdx,
      normalizeStatutoryDataHeaderKey(header)
    );
    if (val) return val;
  }
  for (let colIdx = 0; colIdx < headers.length; colIdx += 1) {
    const header = headers[colIdx];
    if (isStatutoryDataIdentityHeader(header)) continue;
    const val = resolveStatutoryDataRowCell(
      row,
      headers,
      colIdx,
      normalizeStatutoryDataHeaderKey(header)
    );
    if (val) return val;
  }
  return '';
}

function resolveStatutoryDataCellValueByHeader(row, headers, targetHeader) {
  const normalizedTarget = normalizeStatutoryDataHeaderKey(targetHeader);
  if (!normalizedTarget) return '';
  const colIdx = findStatutoryDataHeaderIndex(headers, normalizedTarget);
  if (colIdx >= 0) {
    const val = resolveStatutoryDataRowCell(row, headers, colIdx, normalizedTarget);
    if (val) return val;
  }
  const matchKey = findStatutoryDataRowObjectKey(row, normalizedTarget);
  if (matchKey && row[matchKey] != null && String(row[matchKey]).trim() !== '') {
    return String(row[matchKey]).trim();
  }
  return '';
}

function resolveStatutoryDataHeaderLabelFromGrid(headers, targetHeader) {
  const colIdx = findStatutoryDataHeaderIndex(headers, targetHeader);
  if (colIdx >= 0) return normalizeStatutoryDataHeaderKey(headers[colIdx]);
  return normalizeStatutoryDataHeaderKey(targetHeader);
}

function resolveStatutoryDataFieldsForRow(row, headers, editedHeaders = null) {
  const fields = [];
  const seen = new Set();
  const pushField = (columnName, value) => {
    const col = normalizeStatutoryDataHeaderKey(columnName);
    const val = value == null ? '' : String(value).trim();
    if (!col || !val) return;
    const key = `${col}\0${val}`;
    if (seen.has(key)) return;
    seen.add(key);
    fields.push({ columnName: col, value: val });
  };

  if (Array.isArray(editedHeaders) && editedHeaders.length > 0) {
    editedHeaders.forEach((header) => {
      if (!header || isStatutoryDataSerialOnlyHeader(header)) return;
      const val = resolveStatutoryDataCellValueByHeader(row, headers, header);
      if (val) pushField(resolveStatutoryDataHeaderLabelFromGrid(headers, header), val);
    });
    if (fields.length > 0) return fields;
  }

  if (Array.isArray(editedHeaders)) {
    for (let colIdx = 0; colIdx < headers.length; colIdx += 1) {
      const header = headers[colIdx];
      if (isStatutoryDataIdentityHeader(header)) continue;
      const val = resolveStatutoryDataRowCell(
        row,
        headers,
        colIdx,
        normalizeStatutoryDataHeaderKey(header)
      );
      if (val) pushField(header, val);
    }
    return fields;
  }

  for (let colIdx = 0; colIdx < headers.length; colIdx += 1) {
    const header = headers[colIdx];
    if (!isStatutoryDataPrimaryValueHeader(header)) continue;
    const val = resolveStatutoryDataRowCell(
      row,
      headers,
      colIdx,
      normalizeStatutoryDataHeaderKey(header)
    );
    if (val) {
      pushField(header, val);
      return fields;
    }
  }
  for (let colIdx = 0; colIdx < headers.length; colIdx += 1) {
    const header = headers[colIdx];
    if (isStatutoryDataIdentityHeader(header)) continue;
    const val = resolveStatutoryDataRowCell(
      row,
      headers,
      colIdx,
      normalizeStatutoryDataHeaderKey(header)
    );
    if (val) {
      pushField(header, val);
      return fields;
    }
  }
  return fields;
}

function buildStatutoryDataEditedHeadersByRow(payload) {
  const byRow = new Map();
  const cells = Array.isArray(payload?.statutoryDataEditedCells) ? payload.statutoryDataEditedCells : [];
  cells.forEach((cell) => {
    if (!cell || typeof cell !== 'object') return;
    const rowIdx = Number(cell.rowIndex);
    if (!Number.isInteger(rowIdx) || rowIdx < 0) return;
    const header = normalizeStatutoryDataHeaderKey(cell.header);
    if (!byRow.has(rowIdx)) byRow.set(rowIdx, []);
    if (header) byRow.get(rowIdx).push(header);
  });
  return byRow;
}

function resolveStatutoryDataEligibleRowIndexes(payload, rows, headers) {
  const indexes = new Set();
  const touched = Array.isArray(payload?.statutoryDataTouchedRowIndexes)
    ? payload.statutoryDataTouchedRowIndexes
        .map((idx) => Number(idx))
        .filter((idx) => Number.isInteger(idx) && idx >= 0 && idx < rows.length)
    : [];
  touched.forEach((idx) => indexes.add(idx));

  const editedCells = Array.isArray(payload?.statutoryDataEditedCells)
    ? payload.statutoryDataEditedCells
    : [];
  editedCells.forEach((cell) => {
    const rowIdx = Number(cell?.rowIndex);
    if (Number.isInteger(rowIdx) && rowIdx >= 0 && rowIdx < rows.length) indexes.add(rowIdx);
  });

  if (indexes.size > 0) return Array.from(indexes);

  const eligible = [];
  rows.forEach((row, rowIdx) => {
    if (!row || typeof row !== 'object' || Array.isArray(row)) return;
    const employeeKey = resolveStatutoryDataEmployeeKey(row, headers, rowIdx);
    const value = resolveStatutoryDataPrimaryValue(row, headers);
    if (employeeKey && value) eligible.push(rowIdx);
  });
  return eligible.length === 1 ? eligible : [];
}

function statutoryDataPayloadFromBody(body) {
  const headers = Array.isArray(body?.sampleDataHeader)
    ? body.sampleDataHeader
    : Array.isArray(body?.headers)
      ? body.headers
      : [];
  const rows = Array.isArray(body?.sampleData)
    ? body.sampleData
    : Array.isArray(body?.rows)
      ? body.rows
      : [];
  const headerFormData =
    body?.sampleHeaderFormData &&
    typeof body.sampleHeaderFormData === 'object' &&
    !Array.isArray(body.sampleHeaderFormData)
      ? body.sampleHeaderFormData
      : body?.headerFormData &&
          typeof body.headerFormData === 'object' &&
          !Array.isArray(body.headerFormData)
        ? body.headerFormData
        : {};
  const headerFieldDefinitions = Array.isArray(body?.headerFieldDefinitions)
    ? body.headerFieldDefinitions
    : [];
  const monthfilterRaw = normalizeMonthFilterValue(body);
  const statutoryDataTouchedRowIndexes = Array.isArray(body?.statutoryDataTouchedRowIndexes)
    ? body.statutoryDataTouchedRowIndexes
        .map((idx) => Number(idx))
        .filter((idx) => Number.isInteger(idx) && idx >= 0)
    : [];
  const statutoryDataEditedCells = Array.isArray(body?.statutoryDataEditedCells)
    ? body.statutoryDataEditedCells.filter((cell) => cell && typeof cell === 'object')
    : [];
  const statutoryDataEmployeeOrder = Array.isArray(body?.statutoryDataEmployeeOrder)
    ? body.statutoryDataEmployeeOrder.map((id) => String(id ?? '').trim())
    : [];
  return {
    headers,
    rows,
    headerFormData,
    headerFieldDefinitions,
    formName: pickNonEmptyText(body?.formName, body?.FormName, body?.formname),
    FormName: pickNonEmptyText(body?.FormName, body?.formName, body?.formname),
    formHeaderTitle: pickNonEmptyText(body?.formHeaderTitle, body?.parsedFormHeaderTitle),
    parsedFormHeaderTitle: pickNonEmptyText(body?.parsedFormHeaderTitle, body?.formHeaderTitle),
    statutoryDataTouchedRowIndexes,
    statutoryDataEditedCells,
    statutoryDataEmployeeOrder,
    monthfilter: trimMonthFilterForDatastore(monthfilterRaw),
    monthFilter: trimMonthFilterForDatastore(monthfilterRaw),
    MonthFilter: trimMonthFilterForDatastore(monthfilterRaw)
  };
}

function buildStatutoryDataFieldEntries(payload) {
  const headers = Array.isArray(payload?.headers) ? payload.headers : [];
  const rows = Array.isArray(payload?.rows) ? payload.rows : [];
  const employeeOrder = Array.isArray(payload?.statutoryDataEmployeeOrder)
    ? payload.statutoryDataEmployeeOrder.map((id) => String(id ?? '').trim())
    : [];
  const entries = [];
  const seen = new Set();
  const editedHeadersByRow = buildStatutoryDataEditedHeadersByRow(payload);
  const resolveRowKey = (row, rowIdx) =>
    resolveStatutoryDataEmployeeKey(row, headers, rowIdx, employeeOrder);

  const pushEntry = (changeHeader, columnName, value) => {
    const employeeKey = normalizeStatutoryDataHeaderKey(changeHeader);
    const col = normalizeStatutoryDataHeaderKey(columnName);
    const val = value == null ? '' : String(value).trim();
    if (!employeeKey || !col || val === '') return;
    const dedupeKey = `${employeeKey}\0${col}\0${val}`;
    if (seen.has(dedupeKey)) return;
    seen.add(dedupeKey);
    entries.push({ changeHeader: employeeKey, columnName: col, value: val });
  };

  const editedCells = Array.isArray(payload?.statutoryDataEditedCells)
    ? payload.statutoryDataEditedCells
    : [];
  editedCells.forEach((cell) => {
    if (!cell || typeof cell !== 'object') return;
    const rowIdx = Number(cell.rowIndex);
    const header = normalizeStatutoryDataHeaderKey(cell.header);
    if (!Number.isInteger(rowIdx) || rowIdx < 0 || rowIdx >= rows.length || !header) return;
    if (isStatutoryDataSerialOnlyHeader(header)) return;
    const row = rows[rowIdx];
    const employeeKey = resolveRowKey(row, rowIdx);
    if (!employeeKey) return;
    const val = resolveStatutoryDataCellValueByHeader(row, headers, header);
    if (!val) return;
    pushEntry(employeeKey, resolveStatutoryDataHeaderLabelFromGrid(headers, header), val);
  });

  const eligibleRowIndexes = resolveStatutoryDataEligibleRowIndexes(payload, rows, headers);
  const touchedRowSet = new Set(
    (Array.isArray(payload?.statutoryDataTouchedRowIndexes)
      ? payload.statutoryDataTouchedRowIndexes
      : []
    )
      .map((idx) => Number(idx))
      .filter((idx) => Number.isInteger(idx) && idx >= 0)
  );
  eligibleRowIndexes.forEach((rowIdx) => {
    const row = rows[rowIdx];
    if (!row || typeof row !== 'object' || Array.isArray(row)) return;
    const employeeKey = resolveRowKey(row, rowIdx);
    if (!employeeKey) return;
    const editedHeaders = editedHeadersByRow.has(rowIdx)
      ? editedHeadersByRow.get(rowIdx)
      : touchedRowSet.has(rowIdx)
        ? []
        : null;
    const fieldEntries = resolveStatutoryDataFieldsForRow(row, headers, editedHeaders);
    fieldEntries.forEach(({ columnName, value }) => {
      pushEntry(employeeKey, columnName, value);
    });
  });

  return entries;
}

function buildStatutoryDataReplaceKey(entry) {
  const form = normalizeStatutoryDataHeaderKey(entry.formName || '');
  const hdr = normalizeStatutoryDataHeaderKey(entry.changeHeader);
  const col = normalizeStatutoryDataHeaderKey(entry.columnName);
  return `${form}\0${hdr}\0${col}`;
}

function buildStatutoryDataRowReplaceKey(row) {
  const form = normalizeStatutoryDataHeaderKey(row?.FormName || row?.formName || '');
  const hdr = normalizeStatutoryDataHeaderKey(row?.ChangeDataHeader);
  const col = normalizeStatutoryDataHeaderKey(row?.ColumnName || row?.columnName || '');
  return `${form}\0${hdr}\0${col}`;
}

function resolveStatutoryDataFormName(payload, statutoryRow = null) {
  const titleLine = (value) => {
    const t = String(value || '').trim();
    if (!t) return '';
    return t.split('\n')[0].trim();
  };
  const raw =
    pickNonEmptyText(
      payload?.formName,
      payload?.FormName,
      payload?.formname,
      titleLine(payload?.formHeaderTitle),
      titleLine(payload?.parsedFormHeaderTitle),
      payload?.description,
      payload?.Description,
      payload?.act,
      payload?.Act
    ) ||
    pickNonEmptyText(statutoryRow?.FormName, statutoryRow?.Act, statutoryRow?.Description) ||
    '';
  return canonicalStatutoryDataFormName(raw) || raw;
}

async function insertStatutoryDataRow(table, entry, monthStore) {
  const baseRow = {
    ChangeDataHeader: entry.changeHeader,
    Value: entry.value,
    MonthStore: monthStore
  };
  const layerCandidates = [
    entry.formName || entry.columnName
      ? {
          ...baseRow,
          ...(entry.columnName ? { ColumnName: entry.columnName } : {}),
          ...(entry.formName ? { FormName: entry.formName } : {})
        }
      : null,
    entry.columnName ? { ...baseRow, ColumnName: entry.columnName } : null,
    baseRow
  ].filter(Boolean);
  const layers = [];
  const seen = new Set();
  layerCandidates.forEach((row) => {
    const key = JSON.stringify(row);
    if (seen.has(key)) return;
    seen.add(key);
    layers.push(row);
  });

  let lastErr = null;
  for (const row of layers) {
    try {
      await table.insertRow(row);
      return;
    } catch (err) {
      lastErr = err;
      if (!isDatastoreColumnError(err)) throw err;
    }
  }
  if (lastErr) throw lastErr;
}

/** Save employee field values into StatutoryData (FormName, ChangeDataHeader, ColumnName, Value, MonthStore). */
async function persistStatutoryDataFields(catalyst, payload, statutoryRow = null) {
  const monthStore =
    resolveMonthFilterForSampleDataSave(payload, statutoryRow) || MONTH_NAMES[new Date().getMonth()];
  const formName = resolveStatutoryDataFormName(payload, statutoryRow);
  const entries = buildStatutoryDataFieldEntries(payload).map((entry) => ({
    ...entry,
    formName
  }));
  if (entries.length === 0) {
    return { saved: false, reason: 'empty_entries' };
  }

  const table = catalyst.datastore().table(STATUTORY_DATA_TABLE);
  const zcql = catalyst.zcql();
  const replaceKeys = new Set(entries.map((entry) => buildStatutoryDataReplaceKey(entry)));
  const employeesToReplace = new Set(
    entries.map(
      (entry) =>
        `${normalizeStatutoryDataHeaderKey(entry.formName)}\0${normalizeStatutoryDataHeaderKey(entry.changeHeader)}`
    )
  );
  const replacingEmployeeColumns = new Set(
    entries.map(
      (entry) =>
        `${normalizeStatutoryDataHeaderKey(entry.changeHeader)}\0${normalizeStatutoryDataHeaderKey(entry.columnName)}`
    )
  );

  try {
    const escapedMonth = String(monthStore).replace(/'/g, "''");
    let existing = [];
    const queryAttempts = [
      `SELECT ROWID, FormName, ChangeDataHeader, ColumnName, MonthStore FROM ${STATUTORY_DATA_TABLE} WHERE MonthStore = '${escapedMonth}'`,
      `SELECT ROWID, ChangeDataHeader, ColumnName, MonthStore FROM ${STATUTORY_DATA_TABLE} WHERE MonthStore = '${escapedMonth}'`,
      `SELECT ROWID, ChangeDataHeader, MonthStore FROM ${STATUTORY_DATA_TABLE} WHERE MonthStore = '${escapedMonth}'`
    ];
    for (const query of queryAttempts) {
      try {
        existing = await zcql.executeZCQLQuery(query);
        break;
      } catch (queryErr) {
        if (!isDatastoreColumnError(queryErr)) throw queryErr;
      }
    }
    const list = Array.isArray(existing) ? existing : [];
    const toDelete = [];
    for (const item of list) {
      const row = pickStatutoryDataRow(item);
      if (!row?.ROWID) continue;
      const rowKey = buildStatutoryDataRowReplaceKey(row);
      const form = normalizeStatutoryDataHeaderKey(row.FormName || row.formName || '');
      const hdr = normalizeStatutoryDataHeaderKey(row.ChangeDataHeader);
      const col = normalizeStatutoryDataHeaderKey(row.ColumnName || row.columnName || '');
      if (
        replaceKeys.has(rowKey) ||
        (employeesToReplace.has(`${form}\0${hdr}`) && !col) ||
        (!form && col && replacingEmployeeColumns.has(`${hdr}\0${col}`))
      ) {
        toDelete.push(row.ROWID);
      }
    }
    if (toDelete.length > 0) {
      await table.deleteRows(toDelete);
    }
  } catch (queryErr) {
    console.warn('StatutoryData upsert lookup failed, inserting anyway:', queryErr.message);
  }

  let inserted = 0;
  for (const entry of entries) {
    try {
      await insertStatutoryDataRow(table, entry, monthStore);
      inserted += 1;
    } catch (insertErr) {
      console.error(
        'StatutoryData insert failed:',
        insertErr.message,
        entry.formName,
        entry.changeHeader,
        entry.columnName
      );
    }
  }

  return { saved: inserted > 0, inserted, monthStore, formName: formName || null };
}

/** Catalyst/API payloads vary: monthfilter | monthFilter | MonthFilter */
function normalizeMonthFilterValue(body) {
  if (!body || typeof body !== 'object') return undefined;
  if ('monthfilter' in body) return body.monthfilter;
  if ('monthFilter' in body) return body.monthFilter;
  if ('MonthFilter' in body) return body.MonthFilter;
  return undefined;
}

function trimMonthFilterForDatastore(value) {
  if (value == null) return null;
  const s = String(value).trim();
  if (s === '' || s === 'null' || s === 'undefined') return null;
  return s;
}

function resolveMonthFilterFromSource(source) {
  if (!source || typeof source !== 'object') return null;
  const raw =
    source.monthfilter !== undefined
      ? source.monthfilter
      : source.monthFilter !== undefined
        ? source.monthFilter
        : source.MonthFilter;
  return trimMonthFilterForDatastore(raw);
}

function resolveSampleDataRowMonthFilter(row, headerObj) {
  return (
    resolveMonthFilterFromSource(row) ||
    trimMonthFilterForDatastore(headerObj?.monthfilter) ||
    null
  );
}

function pickNonEmptyText(...values) {
  for (const v of values) {
    const s = String(v ?? '').trim();
    if (s && s !== 'null' && s !== 'undefined') return s;
  }
  return '';
}

/** Reject accidental time/id strings in Site (e.g. Catalyst CREATEDTIME shown in wrong column). */
function looksLikeInvalidReturnedSiteValue(value) {
  const s = String(value || '').trim();
  if (!s) return true;
  if (/^\d{1,2}:\d{2}:\d{2,}$/.test(s)) return true;
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return true;
  if (/^\d{10,}$/.test(s)) return true;
  return false;
}

function resolveSiteNameForReturned(statutoryRow, existingRecord, reqBody) {
  const candidates = [
    reqBody?.site,
    reqBody?.Site,
    statutoryRow?.Site,
    existingRecord?.Site
  ];
  for (const c of candidates) {
    const s = String(c ?? '').trim();
    if (s && !looksLikeInvalidReturnedSiteValue(s)) return s;
  }
  return '';
}

function safeJsonParse(value, fallback = null) {
  try {
    if (value == null) return fallback;
    if (typeof value === 'object') return value;
    const s = String(value).trim();
    if (!s) return fallback;
    return JSON.parse(s);
  } catch (_) {
    return fallback;
  }
}

// Map document types to folder IDs (your Catalyst Filestore folders)
const DOC_TYPE_TO_FOLDER_ID = {
  Form: '26741000000061758', // StatutoryFormFile
  ProofSubmission: '26741000000061739', // StatutoryProofSubmissionFile
  Draft: '26741000000061720' // DraftFileName
};

// Folder names used when configured folder ID does not exist (404/INVALID_ID)
const DOC_TYPE_TO_FOLDER_NAME = {
  Form: 'StatutoryFormFile',
  ProofSubmission: 'StatutoryProofSubmissionFile',
  Draft: 'DraftFileName'
};

// Resolve folder ID: find by name or create. Use when upload fails with INVALID_ID/404.
async function getOrCreateFolderId(catalyst, docType) {
  const folderName = DOC_TYPE_TO_FOLDER_NAME[docType];
  if (!folderName) return null;
  try {
    const allFolders = await catalyst.filestore().getAllFolders();
    const folderDetailsArr = allFolders.map((f) => (f.toJSON && f.toJSON()) || f);
    const found = folderDetailsArr.find((d) => (d.folder_name || d.folderName || '').trim() === folderName);
    if (found && (found.id != null)) {
      console.log('Using existing Filestore folder by name:', folderName, 'id:', found.id);
      return String(found.id);
    }
  } catch (listErr) {
    console.warn('Could not list Filestore folders:', listErr.message);
  }
  try {
    const newFolder = await catalyst.filestore().createFolder(folderName);
    const details = (newFolder.toJSON && newFolder.toJSON()) || newFolder;
    const id = details.id != null ? String(details.id) : null;
    if (id) console.log('Created Filestore folder:', folderName, 'id:', id);
    return id;
  } catch (createErr) {
    console.error('Could not create Filestore folder:', createErr);
    return null;
  }
}

// Form template folder (FormFolder) and file IDs from filestore
const FORM_TEMPLATE_FOLDER_ID = '26741000000061701';
/** Same folder as formmaster_function File Store "Templates" — fallback when FormFolder IDs differ */
const FORMMASTER_TEMPLATES_FOLDER_ID = '26741000000173291';
const FORM_TEMPLATE_FILE_MAP = {
  formu: { id: '26741000000162768', name: 'Form U.xlsx' },
  formi: { id: '26741000000162763', name: 'Form I.xlsx' },
  formv: { id: '26741000000162758', name: 'Form V.xlsx' },
  formw: { id: '26741000000162748', name: 'Form W.xlsx' },
  formx: { id: '26741000000162753', name: 'Form X.xlsx' }
};
/** Alternate file IDs in Cloud Scale → File Store → Templates (see Form U.xlsx etc.) */
const FORMMASTER_TEMPLATES_FILE_FALLBACK = {
  formu: { id: '26741000000173342', name: 'Form U.xlsx' },
  formi: { id: '26741000000173332', name: 'Form I.xlsx' },
  formv: { id: '26741000000173337', name: 'Form V.xlsx' },
  formw: { id: '26741000000173322', name: 'Form W.xlsx' },
  formx: { id: '26741000000173327', name: 'Form X.xlsx' }
};

// Local template fallbacks for statutory forms (used when filestore entry is missing)
const FORM_TEMPLATE_MAP = {
  formu: path.join(__dirname, 'templates', 'Form U.xlsx'),
  formi: path.join(__dirname, 'templates', 'Form I.xlsx'),
  formv: path.join(__dirname, 'templates', 'Form V.xlsx'),
  formw: path.join(__dirname, 'templates', 'Form W.xlsx'),
  formx: path.join(__dirname, 'templates', 'Form X.xlsx')
};

const normalizeFormKey = (value = '') =>
  String(value || '')
    .toLowerCase()
    .replace(/[^a-z0-9]/g, '');

const isFileMissingError = (err) =>
  err?.statusCode === 404 ||
  err?.code === 'INVALID_ID' ||
  (err?.message && err.message.includes('No such file'));

const getContentType = (filename) => {
  const ext = String(filename || '').toLowerCase().split('.').pop();
  const contentTypes = {
    pdf: 'application/pdf',
    jpg: 'image/jpeg',
    jpeg: 'image/jpeg',
    png: 'image/png',
    gif: 'image/gif',
    webp: 'image/webp',
    doc: 'application/msword',
    docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    xls: 'application/vnd.ms-excel',
    xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    txt: 'text/plain'
  };
  return contentTypes[ext] || 'application/octet-stream';
};

/** OOXML .xlsx is a ZIP file starting with PK — use when proof filename omits extension */
const proofBufferLooksLikeExcel = (buf) =>
  buf && Buffer.isBuffer(buf) && buf.length >= 4 && buf[0] === 0x50 && buf[1] === 0x4b;

const normalizeCellValue = (value) =>
  String(value == null ? '' : value).trim();

const isEmptyRow = (row) =>
  !row || row.length === 0 || row.every((cell) => !normalizeCellValue(cell));

const extractProofSubmissionData = (proofBuffer) => {
  const workbook = XLSX.read(proofBuffer, { type: 'buffer' });
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

  const rowMeta = rows.map((row) => {
    const values = (row || []).map((cell) => normalizeCellValue(cell));
    const nonEmptyCount = values.filter(Boolean).length;
    return { values, nonEmptyCount };
  });

  let tableHeaderIndex = -1;
  let bestScore = -1;
  for (let i = 0; i < rowMeta.length; i += 1) {
    const current = rowMeta[i];
    if (!current || current.nonEmptyCount === 0) continue;
    const nextHasData = rowMeta[i + 1]?.nonEmptyCount > 0;
    if (!nextHasData) continue;
    const prevEmpty = i === 0 || rowMeta[i - 1]?.nonEmptyCount === 0;
    const score = current.nonEmptyCount + (prevEmpty ? 2 : 0);
    if (score > bestScore) {
      bestScore = score;
      tableHeaderIndex = i;
    }
  }

  const tableHeaders =
    tableHeaderIndex >= 0 ? rowMeta[tableHeaderIndex].values : [];

  const tableData = [];
  if (tableHeaderIndex >= 0) {
    for (let i = tableHeaderIndex + 1; i < rows.length; i += 1) {
      const rowValues = (rows[i] || []).map((cell) => normalizeCellValue(cell));
      if (isEmptyRow(rowValues)) break;
      const rowData = {};
      tableHeaders.forEach((header, colIndex) => {
        const key = normalizeCellValue(header);
        if (key) {
          rowData[key] = normalizeCellValue(rowValues[colIndex] || '');
        }
      });
      if (Object.keys(rowData).length > 0) {
        tableData.push(rowData);
      }
    }
  }

  let headerFields = {};
  if (tableHeaderIndex > 1) {
    let candidateIndex = -1;
    for (let i = 0; i < tableHeaderIndex - 1; i += 1) {
      const labels = rowMeta[i];
      const values = rowMeta[i + 1];
      if (!labels || !values) continue;
      if (labels.nonEmptyCount === 0 || values.nonEmptyCount === 0) continue;
      const prevEmpty = i === 0 || rowMeta[i - 1]?.nonEmptyCount === 0;
      if (!prevEmpty) continue;
      candidateIndex = i;
    }
    if (candidateIndex >= 0) {
      const labels = rowMeta[candidateIndex].values;
      const values = rowMeta[candidateIndex + 1].values;
      const fields = {};
      labels.forEach((label, idx) => {
        const key = normalizeCellValue(label);
        if (key) {
          fields[key] = normalizeCellValue(values[idx] || '');
        }
      });
      headerFields = fields;
    }
  }

  return {
    tableHeaders,
    tableData,
    headerFields
  };
};

const findTemplateHeaderRow = (rows, headers) => {
  if (!headers || headers.length === 0) return -1;
  const normalizedHeaders = headers.map((header) => normalizeCellValue(header).toLowerCase());
  let bestIndex = -1;
  let bestMatches = 0;
  rows.forEach((row, idx) => {
    const rowValues = (row || []).map((cell) => normalizeCellValue(cell).toLowerCase());
    const matches = normalizedHeaders.filter((header) => header && rowValues.includes(header)).length;
    if (matches > bestMatches) {
      bestMatches = matches;
      bestIndex = idx;
    }
  });
  return bestMatches > 0 ? bestIndex : -1;
};

const applyProofDataToTemplate = (templateBuffer, proofData) => {
  const workbook = XLSX.read(templateBuffer, { type: 'buffer', cellStyles: true });
  let sheet;
  let rows;
  if (proofData.sheetName && workbook.Sheets[proofData.sheetName]) {
    sheet = workbook.Sheets[proofData.sheetName];
    rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
  } else {
    const loc = locateProofTemplateTable(workbook, XLSX);
    sheet = workbook.Sheets[loc.sheetName];
    rows = loc.rows;
  }

  let tableHeaderIndex =
    proofData.tableHeaderRowIndex != null && proofData.tableHeaderRowIndex >= 0
      ? proofData.tableHeaderRowIndex
      : findTemplateHeaderRow(rows, proofData.tableHeaders || []);

  if (tableHeaderIndex >= 0) {
    const templateHeaderRow = (rows[tableHeaderIndex] || []).map((cell) =>
      normalizeCellValue(cell).toLowerCase()
    );
    const headerMap = {};
    templateHeaderRow.forEach((header, colIndex) => {
      if (header) {
        headerMap[header] = colIndex;
      }
    });

    const tableRows = proofData.tableData || [];
    const styleRowIndex = tableHeaderIndex + 1;
    const cloneStyle = (style) => (style ? JSON.parse(JSON.stringify(style)) : undefined);

    // When we extend the table with new rows, the template borders/lines will only show
    // if styles (especially borders) exist on every cell in that row range.
    // This helper copies styles from an existing "style source row" into the target row.
    const applyRowStylesFromTemplate = (targetRowIndex, maxColInclusive) => {
      for (let colIndex = 0; colIndex <= maxColInclusive; colIndex += 1) {
        const targetAddr = XLSX.utils.encode_cell({ r: targetRowIndex, c: colIndex });
        const existingCell = sheet[targetAddr];
        const styleCell =
          sheet[XLSX.utils.encode_cell({ r: styleRowIndex, c: colIndex })] ||
          sheet[XLSX.utils.encode_cell({ r: tableHeaderIndex, c: colIndex })];

        if (!styleCell?.s) continue;

        if (existingCell) {
          if (!existingCell.s) existingCell.s = cloneStyle(styleCell.s);
          continue;
        }

        // Create an "empty" cell with style so borders render correctly.
        sheet[targetAddr] = { t: 's', v: '', s: cloneStyle(styleCell.s) };
      }
    };

    const tableColEnd = (() => {
      const headerCols = Object.values(headerMap);
      const headerMax = headerCols.length ? Math.max(...headerCols) : 0;
      // Fallback to current sheet range if available
      if (sheet['!ref']) {
        try {
          const r = XLSX.utils.decode_range(sheet['!ref']);
          return Math.max(headerMax, r.e.c);
        } catch (e) {
          return headerMax;
        }
      }
      return headerMax;
    })();

    tableRows.forEach((rowData, rowIndex) => {
      const targetRowIndex = tableHeaderIndex + 1 + rowIndex;
      // Ensure borders/lines exist across the entire row (not just the filled cells).
      applyRowStylesFromTemplate(targetRowIndex, tableColEnd);

      Object.entries(rowData).forEach(([header, value]) => {
        const columnIndex = headerMap[normalizeCellValue(header).toLowerCase()];
        if (columnIndex == null) return;
        const cellAddress = XLSX.utils.encode_cell({
          r: targetRowIndex,
          c: columnIndex
        });
        const existingCell = sheet[cellAddress] || {};
        let styleCell =
          sheet[XLSX.utils.encode_cell({ r: styleRowIndex, c: columnIndex })] ||
          sheet[XLSX.utils.encode_cell({ r: tableHeaderIndex, c: columnIndex })];
        if (!styleCell?.s) {
          for (let probeRow = styleRowIndex + 1; probeRow <= styleRowIndex + 20; probeRow += 1) {
            const probeCell = sheet[XLSX.utils.encode_cell({ r: probeRow, c: columnIndex })];
            if (probeCell?.s) {
              styleCell = probeCell;
              break;
            }
          }
        }
        const cellValue = normalizeCellValue(value);
        const numericValue = Number(cellValue);
        const nextCell = {
          ...existingCell,
          t: cellValue !== '' && !Number.isNaN(numericValue) && String(numericValue) === cellValue ? 'n' : 's',
          v: cellValue !== '' && !Number.isNaN(numericValue) && String(numericValue) === cellValue ? numericValue : cellValue
        };
        if (!nextCell.s && styleCell?.s) {
          nextCell.s = cloneStyle(styleCell.s);
        }
        sheet[cellAddress] = nextCell;
      });
    });

    const newDataEndRow = tableHeaderIndex + tableRows.length;
    let lastExistingRow = tableHeaderIndex;
    if (sheet['!ref']) {
      try {
        lastExistingRow = XLSX.utils.decode_range(sheet['!ref']).e.r;
      } catch (e) {
        /* keep tableHeaderIndex */
      }
    }
    for (let r = newDataEndRow + 1; r <= lastExistingRow; r += 1) {
      for (let c = 0; c <= tableColEnd; c += 1) {
        const cellAddress = XLSX.utils.encode_cell({ r, c });
        const existing = sheet[cellAddress];
        if (!existing) continue;
        sheet[cellAddress] = { ...existing, t: 's', v: '' };
      }
    }
  }

  const headerFields = proofData.headerFields || {};
  if (Object.keys(headerFields).length > 0) {
    const normalizedFieldMap = {};
    Object.entries(headerFields).forEach(([label, value]) => {
      const key = normalizeCellValue(label).toLowerCase();
      if (key) {
        normalizedFieldMap[key] = value;
      }
    });

    rows.forEach((row, rowIndex) => {
      (row || []).forEach((cell, colIndex) => {
        const label = normalizeCellValue(cell).toLowerCase();
        if (!label || !(label in normalizedFieldMap)) return;
        const value = normalizeCellValue(normalizedFieldMap[label]);
        if (value === '') return;
        const targetCell = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex + 1 });
        const existingCell = sheet[targetCell] || {};
        const numericValue = Number(value);
        const nextCell = {
          ...existingCell,
          t: value !== '' && !Number.isNaN(numericValue) && String(numericValue) === value ? 'n' : 's',
          v: value !== '' && !Number.isNaN(numericValue) && String(numericValue) === value ? numericValue : value
        };
        sheet[targetCell] = nextCell;
      });
    });
  }

  if (sheet['!ref']) {
    const range = XLSX.utils.decode_range(sheet['!ref']);
    const maxDataRow = (proofData.tableData || []).length;
    if (tableHeaderIndex >= 0 && maxDataRow > 0) {
      range.e.r = Math.max(range.e.r, tableHeaderIndex + maxDataRow);
    }
    sheet['!ref'] = XLSX.utils.encode_range(range);
  }

  return XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx', cellStyles: true });
};

/**
 * Proof excel with ?format=template: load the master template from File Store first
 * (e.g. Form U `26741000000173342` — full row-8 headers), then SEMaster FormFile / legacy folders.
 * Optional: ?templateFileId=...&templateFileName=Form%20U.xlsx to pin a workbook from Templates folder.
 */
async function tryLoadProofFormTemplateBuffer(catalyst, statutory, normalizedFormName, templatePath, req) {
  const templateInfo = FORM_TEMPLATE_FILE_MAP[normalizedFormName];
  const formmasterCanonical = FORMMASTER_TEMPLATES_FILE_FALLBACK[normalizedFormName];
  let templateBuffer = null;
  let templateFileName =
    formmasterCanonical?.name ||
    templateInfo?.name ||
    (templatePath ? path.basename(templatePath) : 'form_template.xlsx');

  const qId = (req.query.templateFileId || '').trim();
  if (qId) {
    try {
      templateBuffer = await catalyst.filestore().folder(FORMMASTER_TEMPLATES_FOLDER_ID).downloadFile(qId);
      const qName = (req.query.templateFileName || '').trim();
      if (qName) templateFileName = qName.endsWith('.xlsx') ? qName : `${qName}.xlsx`;
      console.log('ProofSubmission template: templateFileId query', qId);
      return { templateBuffer, templateFileName };
    } catch (qErr) {
      console.warn('templateFileId download failed:', qErr.message);
    }
  }

  if (formmasterCanonical?.id) {
    try {
      templateBuffer = await catalyst
        .filestore()
        .folder(FORMMASTER_TEMPLATES_FOLDER_ID)
        .downloadFile(formmasterCanonical.id);
      templateFileName = formmasterCanonical.name || templateFileName;
      console.log('ProofSubmission template: File Store Templates folder', formmasterCanonical.id);
      return { templateBuffer, templateFileName };
    } catch (fbErr) {
      console.warn('File Store Templates (canonical) download failed:', fbErr.message);
    }
  }

  const statutoryFormFileIdRaw =
    statutory.FormFile !== undefined && statutory.FormFile !== null ? String(statutory.FormFile).trim() : '';
  const statutoryFormFileId =
    statutoryFormFileIdRaw && statutoryFormFileIdRaw !== 'null' && statutoryFormFileIdRaw !== ''
      ? statutoryFormFileIdRaw
      : '';
  if (statutoryFormFileId) {
    try {
      const formFolderId = DOC_TYPE_TO_FOLDER_ID['Form'];
      templateBuffer = await catalyst.filestore().folder(formFolderId).downloadFile(statutoryFormFileId);
      console.log('ProofSubmission template: statutory FormFile', statutoryFormFileId);
      const fn = statutory.FormFileName && String(statutory.FormFileName).trim();
      if (fn && !/^view file$/i.test(fn)) {
        templateFileName = /\.(xlsx|xls)$/i.test(fn) ? fn : `${fn.replace(/\.+$/, '')}.xlsx`;
      } else if (formmasterCanonical?.name) {
        templateFileName = formmasterCanonical.name;
      } else if (templateInfo?.name) {
        templateFileName = templateInfo.name;
      }
      return { templateBuffer, templateFileName };
    } catch (formTemplateErr) {
      console.warn('ProofSubmission template: statutory FormFile failed:', formTemplateErr.message);
    }
  }

  if (templateInfo?.id) {
    try {
      templateBuffer = await catalyst
        .filestore()
        .folder(FORM_TEMPLATE_FOLDER_ID)
        .downloadFile(templateInfo.id);
      templateFileName = templateInfo.name || templateFileName;
      console.log('ProofSubmission template: legacy FormFolder', templateInfo.id);
      return { templateBuffer, templateFileName };
    } catch (templateErr) {
      console.warn('Legacy FormFolder template failed:', templateErr.message);
    }
  }

  if (templatePath && fs.existsSync(templatePath)) {
    templateBuffer = fs.readFileSync(templatePath);
    templateFileName = formmasterCanonical?.name || templateInfo?.name || path.basename(templatePath);
    console.log('ProofSubmission template: local FORM_TEMPLATE_MAP');
    return { templateBuffer, templateFileName };
  }

  return { templateBuffer: null, templateFileName };
}

async function mergeProofTemplateToBuffer(templateBuffer, proofBuffer, storedProofFileName, req) {
  let outputBuffer = templateBuffer;
  const fileExt = String(storedProofFileName || '').toLowerCase().split('.').pop();
  const canMergeProof =
    proofBuffer &&
    (fileExt === 'xlsx' || fileExt === 'xls' || proofBufferLooksLikeExcel(proofBuffer));
  const sourceParam = String(req.query.source || 'zoho').toLowerCase();
  const preferZoho = sourceParam !== 'proof';
  let proofData = null;
  if (preferZoho) {
    try {
      proofData = await buildProofDataFromZoho(templateBuffer, XLSX, { limit: 200 });
      if (proofData?.tableData?.length) {
        console.log(
          'ProofSubmission ?format=template: using Zoho People data, rows:',
          proofData.tableData.length
        );
      }
    } catch (zohoErr) {
      console.warn('Zoho People merge for proof template failed:', zohoErr.message);
    }
  }
  if (!proofData && canMergeProof) {
    proofData = extractProofSubmissionData(proofBuffer);
    if (proofData?.tableData?.length) {
      console.log(
        'ProofSubmission ?format=template: using stored proof Excel rows:',
        proofData.tableData.length
      );
    }
  }
  if (proofData && (proofData.tableData || []).length > 0) {
    outputBuffer = applyProofDataToTemplate(templateBuffer, proofData);
  }
  return outputBuffer;
}

// Upload file endpoint for statutory (Form or ProofSubmission)
app.post('/statutory/upload/:docType', async (req, res) => {
  try {
    const { docType } = req.params;
    console.log('Statutory file upload request for docType:', docType);
    console.log('Request headers:', req.headers);
    console.log('Request files:', req.files);
    
    const { catalyst } = res.locals;
    const folderId = DOC_TYPE_TO_FOLDER_ID[docType];
    console.log('Folder ID for docType:', docType, 'is:', folderId);
    
    if (!folderId) {
      console.error('Invalid document type:', docType);
      return res.status(400).json({ status: 'failure', message: `Invalid document type: ${docType}. Supported types: ${Object.keys(DOC_TYPE_TO_FOLDER_ID).join(', ')}` });
    }
    
    if (!req.files || !req.files.file) {
      console.error('No file uploaded for docType:', docType);
      return res.status(400).json({ status: 'failure', message: 'No file uploaded.' });
    }
    
    const file = req.files.file;
    console.log('File details:', {
      name: file.name,
      size: file.size,
      mimetype: file.mimetype,
      tempFilePath: file.tempFilePath
    });

    if (docType === 'ProofSubmission') {
      const allowedProof = /\.(pdf|png|jpe?g|gif|webp|bmp|tiff?|heic)$/i;
      if (!allowedProof.test(file.name || '')) {
        return res.status(400).json({
          status: 'failure',
          message: 'Proof submission must be a PDF or image file (not Excel or Word).'
        });
      }
    }
    
    const tempDir = os.tmpdir();
    const tempPath = path.join(tempDir, file.name);
    
    try {
      await file.mv(tempPath);
      console.log('File moved to temp path:', tempPath);
    } catch (mvErr) {
      console.error('File mv error:', mvErr);
      return res.status(500).json({ status: 'failure', message: 'Failed to save uploaded file.' });
    }

    const isFolderNotFoundError = (err) =>
      err?.statusCode === 404 ||
      err?.code === 'INVALID_ID' ||
      (err?.message && String(err.message).includes('No such resource'));

    let uploadResp;
    let effectiveFolderId = folderId;
    try {
      uploadResp = await catalyst.filestore().folder(effectiveFolderId).uploadFile({
        code: fs.createReadStream(tempPath),
        name: file.name
      });
    } catch (uploadErr) {
      if (isFolderNotFoundError(uploadErr)) {
        console.log('Configured folder ID not found, resolving by name or creating folder for docType:', docType);
        effectiveFolderId = await getOrCreateFolderId(catalyst, docType);
        if (effectiveFolderId) {
          try {
            uploadResp = await catalyst.filestore().folder(effectiveFolderId).uploadFile({
              code: fs.createReadStream(tempPath),
              name: file.name
            });
          } catch (retryErr) {
            console.error('Upload error after folder fallback:', retryErr);
            if (fs.existsSync(tempPath)) fs.unlinkSync(tempPath);
            return res.status(500).json({ status: 'failure', message: retryErr.message || 'File upload failed.' });
          }
        } else {
          console.error('Could not resolve or create folder for docType:', docType);
          if (fs.existsSync(tempPath)) fs.unlinkSync(tempPath);
          return res.status(500).json({ status: 'failure', message: 'Filestore folder not found and could not be created. Create a folder in Catalyst Filestore (e.g. StatutoryFormFile) and set its ID in the function config.' });
        }
      } else {
        console.error('Upload error during file processing:', uploadErr);
        if (fs.existsSync(tempPath)) fs.unlinkSync(tempPath);
        return res.status(500).json({ status: 'failure', message: uploadErr.message || 'File upload failed during processing.' });
      }
    } finally {
      if (fs.existsSync(tempPath)) fs.unlinkSync(tempPath);
    }

    try {
      // Defensive: handle different possible response structures
      let fileId, fileName;
      if (Array.isArray(uploadResp) && uploadResp[0]?.id) {
        fileId = uploadResp[0].id;
        fileName = file.name;
      } else if (uploadResp && uploadResp.id) {
        fileId = uploadResp.id;
        fileName = file.name;
      } else if (uploadResp && uploadResp.file_details && uploadResp.file_details[0]?.id) {
        fileId = uploadResp.file_details[0].id;
        fileName = file.name;
      } else {
        console.error('Unexpected upload response:', uploadResp);
        return res.status(500).json({ status: 'failure', message: 'File upload failed or invalid response from Catalyst.' });
      }

      console.log('FileId to be saved:', fileId);
      res.status(200).json({ status: 'success', fileId, fileName });
    } catch (parseErr) {
      console.error('Error parsing upload response:', parseErr);
      res.status(500).json({ status: 'failure', message: 'File upload response invalid.' });
    }
  } catch (err) {
    console.error('Upload error:', err);
    res.status(500).json({ status: 'failure', message: err.message || 'File upload failed.' });
  }
});

// Get all statutory records
app.get('/statutory', async (req, res) => {
  try {
    res.set('Cache-Control', 'no-store');
    const { catalyst } = res.locals;
    const table = catalyst.datastore().table('Statutory');
    
    // Use getAllRows() method to get all rows
    const rows = await table.getAllRows();
    
    const statutoryData = rows.map((row) => mapStatutoryRowToApi(row));
    
    res.status(200).json({
      status: 'success',
      data: { statutoryData }
    });
  } catch (err) {
    console.error('Error fetching statutory data:', err);
    res.status(500).json({ status: 'failure', message: err.message || 'Failed to fetch statutory data.' });
  }
});

// Get single statutory record
app.get('/statutory/:id', async (req, res) => {
  try {
    res.set('Cache-Control', 'no-store');
    const { id } = req.params;
    const { catalyst } = res.locals;
    const table = catalyst.datastore().table('Statutory');
    
    const row = await table.getRow(id);
    
    if (!row) {
      return res.status(404).json({ status: 'failure', message: 'Statutory record not found.' });
    }
    
    const statutory = mapStatutoryRowToApi(row);
    
    res.status(200).json({
      status: 'success',
      data: { statutory }
    });
  } catch (err) {
    console.error('Error fetching statutory record:', err);
    res.status(500).json({ status: 'failure', message: err.message || 'Failed to fetch statutory record.' });
  }
});

// Save imported/autofill grid snapshot into SampleData without full statutory save
app.post('/statutory/:id/sampledata', async (req, res) => {
  try {
    const { id } = req.params;
    const idStr = String(id || '').trim();
    if (!/^\d+$/.test(idStr)) {
      return res.status(400).json({ status: 'failure', message: 'Invalid statutory id for SampleData save.' });
    }

    const { catalyst } = res.locals;
    const table = catalyst.datastore().table('Statutory');

    let existingRecord = null;
    try {
      existingRecord = await table.getRow(idStr);
    } catch (getRowErr) {
      console.warn('SampleData save: statutory row lookup failed, saving snapshot anyway:', getRowErr.message);
    }

    const {
      sampleDataHeader,
      sampleHeaderFormData,
      sampleData,
      formName
    } = req.body;
    const monthfilterRaw = normalizeMonthFilterValue(req.body);

    const headers = Array.isArray(sampleDataHeader) ? sampleDataHeader : [];
    const rows = Array.isArray(sampleData) ? sampleData : [];
    const headerFormData =
      sampleHeaderFormData && typeof sampleHeaderFormData === 'object' && !Array.isArray(sampleHeaderFormData)
        ? sampleHeaderFormData
        : {};

    if (headers.length === 0 && rows.length === 0 && Object.keys(headerFormData).length === 0) {
      return res.status(400).json({ status: 'failure', message: 'No sample data to save.' });
    }

    const persistResult = await persistSampleDataSnapshot(
      catalyst,
      {
        statutoryId: existingRecord?.ROWID || idStr,
        formName: formName || existingRecord?.FormName || null,
        monthfilter:
          trimMonthFilterForDatastore(monthfilterRaw) ||
          existingRecord?.MonthFilter ||
          null,
        monthFilter:
          trimMonthFilterForDatastore(monthfilterRaw) ||
          existingRecord?.MonthFilter ||
          null,
        MonthFilter:
          trimMonthFilterForDatastore(monthfilterRaw) ||
          existingRecord?.MonthFilter ||
          null,
        headers,
        headerFormData,
        headerFieldDefinitions: Array.isArray(req.body?.headerFieldDefinitions)
          ? req.body.headerFieldDefinitions
          : [],
        rows
      },
      existingRecord
    );

    if (!persistResult?.saved) {
      return res.status(400).json({
        status: 'failure',
        message: 'Could not save SampleData snapshot.',
        reason: persistResult?.reason || 'unknown'
      });
    }

    console.log(
      'SampleData snapshot saved for statutory',
      idStr,
      'rows:',
      rows.length,
      'sampleDataRowId:',
      persistResult.rowId
    );

    return res.status(200).json({
      status: 'success',
      message: 'SampleData snapshot saved.',
      data: {
        statutoryId: String(existingRecord?.ROWID || idStr),
        rowCount: rows.length,
        sampleDataRowId: persistResult.rowId,
        MonthFilter: persistResult.monthFilter || null,
        monthfilter: persistResult.monthFilter || null,
        statutoryDataSaved: !!persistResult.statutoryDataSaved,
        statutoryDataInserted: persistResult.statutoryDataInserted || 0
      }
    });
  } catch (err) {
    console.error('Error saving SampleData snapshot:', err);
    return res.status(500).json({
      status: 'failure',
      message: err.message || 'Failed to save SampleData snapshot.'
    });
  }
});

// Save autofill/import field values into StatutoryData (ChangeDataHeader, Value, MonthStore)
app.post('/statutory/:id/statutorydata', async (req, res) => {
  try {
    const { id } = req.params;
    const idStr = String(id || '').trim();
    if (!/^\d+$/.test(idStr)) {
      return res.status(400).json({ status: 'failure', message: 'Invalid statutory id for StatutoryData save.' });
    }

    const { catalyst } = res.locals;
    const table = catalyst.datastore().table('Statutory');
    let existingRecord = null;
    try {
      existingRecord = await table.getRow(idStr);
    } catch (getRowErr) {
      console.warn('StatutoryData save: statutory row lookup failed, saving fields anyway:', getRowErr.message);
    }

    const dataPayload = statutoryDataPayloadFromBody(req.body);
    const { headers, rows, headerFormData, headerFieldDefinitions } = dataPayload;
    if (
      headers.length === 0 &&
      rows.length === 0 &&
      Object.keys(headerFormData).length === 0
    ) {
      return res.status(400).json({ status: 'failure', message: 'No field data to save.' });
    }

    const persistResult = await persistStatutoryDataFields(
      catalyst,
      {
        ...dataPayload,
        statutoryId: existingRecord?.ROWID || idStr,
        formName: req.body?.formName || existingRecord?.FormName || null
      },
      existingRecord
    );

    if (!persistResult?.saved) {
      return res.status(400).json({
        status: 'failure',
        message: 'Could not save StatutoryData fields.',
        reason: persistResult?.reason || 'unknown'
      });
    }

    return res.status(200).json({
      status: 'success',
      message: 'StatutoryData fields saved.',
      data: {
        statutoryId: String(existingRecord?.ROWID || idStr),
        inserted: persistResult.inserted || 0,
        MonthStore: persistResult.monthStore || null
      }
    });
  } catch (err) {
    console.error('Error saving StatutoryData fields:', err);
    return res.status(500).json({
      status: 'failure',
      message: err.message || 'Failed to save StatutoryData fields.'
    });
  }
});

function statutoryDataMonthsMatch(stored, query) {
  if (!query) return true;
  if (!stored) return false;
  const a = String(stored).trim().toLowerCase();
  const b = String(query).trim().toLowerCase();
  if (a === b) return true;
  if (a.slice(0, 3) === b.slice(0, 3)) return true;
  return false;
}

function statutoryDataFormsMatch(storedForm, queryForm) {
  const formNorm = canonicalStatutoryDataFormName(storedForm).toLowerCase();
  const queryNorm = canonicalStatutoryDataFormName(queryForm).toLowerCase();
  if (!queryNorm) return true;
  if (!formNorm) return false;
  if (formNorm === queryNorm) return true;
  const baseStored = formNorm.match(/^form\s+\d+/)?.[0] || formNorm;
  const baseQuery = queryNorm.match(/^form\s+\d+/)?.[0] || queryNorm;
  if (baseStored === baseQuery) {
    const partStored = extractStatutoryFormPartToken(storedForm);
    const partQuery = extractStatutoryFormPartToken(queryForm);
    if (partStored && partQuery) return partStored === partQuery;
  }
  const formAlnum = formNorm.replace(/[^a-z0-9]/g, '');
  const queryAlnum = queryNorm.replace(/[^a-z0-9]/g, '');
  if (formAlnum && queryAlnum && formAlnum === queryAlnum) return true;
  if (formNorm.length >= 6 && queryNorm.length >= 6) {
    return formNorm.includes(queryNorm) || queryNorm.includes(formNorm);
  }
  return false;
}

// Fetch saved StatutoryData field values for autofill/download restore (FormName + MonthStore).
app.get('/statutory/:id/statutorydata', async (req, res) => {
  try {
    const { catalyst } = res.locals;
    const zcql = catalyst.zcql();
    const formNameQ = String(
      req.query?.formName || req.query?.FormName || req.query?.formname || ''
    ).trim();
    const monthQ = trimMonthFilterForDatastore(
      req.query?.monthFilter || req.query?.monthfilter || req.query?.month || ''
    );

    const queryAttempts = [
      `SELECT ROWID, FormName, ChangeDataHeader, ColumnName, Value, MonthStore FROM ${STATUTORY_DATA_TABLE}`,
      `SELECT ROWID, ChangeDataHeader, ColumnName, Value, MonthStore FROM ${STATUTORY_DATA_TABLE}`,
      `SELECT ROWID, ChangeDataHeader, Value, MonthStore FROM ${STATUTORY_DATA_TABLE}`
    ];

    let result = [];
    for (const query of queryAttempts) {
      try {
        result = await zcql.executeZCQLQuery(query);
        break;
      } catch (queryErr) {
        if (!isDatastoreColumnError(queryErr)) throw queryErr;
      }
    }

    const rows = Array.isArray(result)
      ? result.map((entry) => pickStatutoryDataRow(entry)).filter(Boolean)
      : [];

    const filtered = rows.filter((row) => {
      if (monthQ && !statutoryDataMonthsMatch(row.MonthStore || row.monthStore, monthQ)) return false;
      if (formNameQ && !statutoryDataFormsMatch(row.FormName || row.formName, formNameQ)) return false;
      const employeeKey = String(row.ChangeDataHeader || '').trim();
      const value = row.Value != null ? String(row.Value).trim() : '';
      return Boolean(employeeKey && value);
    });

    return res.status(200).json({
      status: 'success',
      data: {
        statutoryData: filtered.map((row) => ({
          ROWID: row.ROWID,
          FormName: row.FormName || row.formName || null,
          ChangeDataHeader: row.ChangeDataHeader || null,
          ColumnName: row.ColumnName || row.columnName || null,
          Value: row.Value != null ? String(row.Value) : '',
          MonthStore: row.MonthStore || row.monthStore || null
        }))
      }
    });
  } catch (err) {
    console.error('Error fetching StatutoryData fields:', err);
    return res.status(500).json({
      status: 'failure',
      message: err.message || 'Failed to fetch StatutoryData fields.'
    });
  }
});

function normalizeStatutoryIdForSampleDataMatch(value) {
  const s = String(value ?? '').trim();
  if (!/^\d+$/.test(s)) return s;
  return s.replace(/^0+/, '') || '0';
}

// Fetch latest SampleData snapshot for a statutory record (used by View Draft generation)
app.get('/statutory/:id/sampledata', async (req, res) => {
  try {
    const { id } = req.params;
    const { catalyst } = res.locals;
    const zcql = catalyst.zcql();
    const query = 'SELECT ROWID, Header, Data, MonthFilter, MODIFIEDTIME, CREATEDTIME FROM SampleData';
    const result = await zcql.executeZCQLQuery(query);

    const rows = Array.isArray(result)
      ? result.map((entry) => entry.SampleData || entry.sampledata || entry)
      : [];

    const idStr = String(id).trim();
    const idNorm = normalizeStatutoryIdForSampleDataMatch(idStr);
    const formNameQ = String(req.query?.formName || req.query?.formname || '').trim().toLowerCase();
    const monthQ = trimMonthFilterForDatastore(
      req.query?.monthFilter || req.query?.monthfilter || req.query?.month || ''
    );
    const matchByFormMonthOnly =
      req.query?.matchByFormMonth === '1' ||
      req.query?.matchByFormMonth === 'true' ||
      (formNameQ && (idStr === '0' || !/^\d+$/.test(idStr)));

    const sampleDataMonthsMatch = (stored, query) => {
      if (!query) return true;
      if (!stored) return false;
      const a = String(stored).trim().toLowerCase();
      const b = String(query).trim().toLowerCase();
      if (a === b) return true;
      if (a.slice(0, 3) === b.slice(0, 3)) return true;
      return false;
    };

    const sampleDataFormsMatch = (storedForm, queryForm) => {
      const formNorm = String(storedForm || '').trim().toLowerCase();
      if (!formNorm || !queryForm) return false;
      return (
        formNorm === queryForm ||
        formNorm.includes(queryForm) ||
        queryForm.includes(formNorm) ||
        (/\bform\s*[-"']?\s*(\d+[a-z]?)\b/i.test(queryForm) &&
          /\bform\s*[-"']?\s*(\d+[a-z]?)\b/i.test(formNorm) &&
          queryForm.match(/\bform\s*[-"']?\s*(\d+[a-z]?)\b/i)?.[1] ===
            formNorm.match(/\bform\s*[-"']?\s*(\d+[a-z]?)\b/i)?.[1])
      );
    };

    const filterByFormMonth = (mappedRows) =>
      mappedRows.filter(
        (r) =>
          sampleDataFormsMatch(r.formName, formNameQ) &&
          sampleDataMonthsMatch(r.monthfilter, monthQ)
      );

    const mapSampleRow = (row) => {
      const headerObj = safeJsonParse(row?.Header, {});
      const dataObj = safeJsonParse(row?.Data, {});
      const statutoryId = headerObj?.statutoryId != null ? String(headerObj.statutoryId).trim() : '';
      const headers = Array.isArray(headerObj?.headers) ? headerObj.headers : [];
      const headerFormData =
        headerObj?.headerFormData && typeof headerObj.headerFormData === 'object' && !Array.isArray(headerObj.headerFormData)
          ? headerObj.headerFormData
          : {};
      const tableRows = Array.isArray(dataObj?.rows) ? dataObj.rows : [];
      const monthfilter = resolveSampleDataRowMonthFilter(row, headerObj);
      return {
        rowid: row?.ROWID != null ? String(row.ROWID) : '',
        modifiedTime: row?.MODIFIEDTIME || row?.CREATEDTIME || '',
        statutoryId,
        formName: headerObj?.formName || null,
        monthfilter,
        MonthFilter: monthfilter,
        headers,
        headerFormData,
        rows: tableRows
      };
    };

    const allMapped = rows.map(mapSampleRow);
    let matching = [];

    if (formNameQ && (matchByFormMonthOnly || monthQ)) {
      matching = filterByFormMonth(allMapped);
    }

    if (matching.length === 0 && formNameQ && !matchByFormMonthOnly) {
      matching = allMapped.filter((r) => sampleDataFormsMatch(r.formName, formNameQ));
    }

    if (matching.length === 0 && !matchByFormMonthOnly && /^\d+$/.test(idStr) && idStr !== '0') {
      matching = allMapped.filter(
        (r) =>
          r.statutoryId === idStr ||
          normalizeStatutoryIdForSampleDataMatch(r.statutoryId) === idNorm
      );
    }

    if (matching.length === 0 && !matchByFormMonthOnly && /^\d+$/.test(idStr) && idStr !== '0') {
      try {
        const statTable = catalyst.datastore().table('Statutory');
        const statRow = await statTable.getRow(idStr);
        const statFormNorm = String(statRow?.FormName || '').trim().toLowerCase();
        const statMonth = trimMonthFilterForDatastore(
          statRow?.MonthFilter ?? statRow?.monthfilter ?? statRow?.MonthFilter
        );
        if (statFormNorm) {
          matching = allMapped.filter((r) => {
            const formNorm = String(r.formName || '').trim().toLowerCase();
            if (!formNorm || formNorm !== statFormNorm) return false;
            if (statMonth && r.monthfilter) {
              return sampleDataMonthsMatch(r.monthfilter, statMonth);
            }
            return true;
          });
        }
      } catch (statLookupErr) {
        console.warn('SampleData GET: statutory row lookup fallback skipped:', statLookupErr.message);
      }
    }

    if (matching.length === 0 && formNameQ) {
      matching = filterByFormMonth(allMapped);
      if (matching.length === 0) {
        matching = allMapped.filter((r) => sampleDataFormsMatch(r.formName, formNameQ));
      }
    }

    if (matching.length === 0) {
      return res.status(200).json({ status: 'success', data: { sampleData: null } });
    }

    matching.sort((a, b) => {
      const ta = Date.parse(a.modifiedTime || '') || 0;
      const tb = Date.parse(b.modifiedTime || '') || 0;
      if (tb !== ta) return tb - ta;
      const ra = Number(a.rowid) || 0;
      const rb = Number(b.rowid) || 0;
      return rb - ra;
    });

    const latest =
      matching.find((row) => Array.isArray(row.rows) && row.rows.length > 0) ||
      matching.find((row) => Array.isArray(row.headers) && row.headers.length > 0) ||
      matching[0];
    return res.status(200).json({
      status: 'success',
      data: {
        sampleData: {
          statutoryId: latest.statutoryId,
          formName: latest.formName,
          monthfilter: latest.monthfilter,
          MonthFilter: latest.monthfilter,
          headers: latest.headers,
          headerFormData: latest.headerFormData,
          rows: latest.rows
        }
      }
    });
  } catch (err) {
    console.error('Error fetching SampleData snapshot:', err);
    return res.status(500).json({ status: 'failure', message: err.message || 'Failed to fetch SampleData.' });
  }
});

// Create statutory record
app.post('/statutory', async (req, res) => {
  try {
    console.log('Received request body:', JSON.stringify(req.body, null, 2));
    
    const {
      formName,
      act,
      description,
      dueDate,
      sector,
      state,
      site,
      autofill,
      draft,
      formFile,
      formFileName,
      sampleFile,
      sampleFileName,
      proofSubmissionFile,
      proofSubmissionFileName,
      draftFile,
      draftFileName,
      approval,
      status,
      sendForApproval,
      remarks,
      submittedDate,
      approvedDate,
      sampleDataHeader,
      sampleHeaderFormData,
      sampleData,
      headerFieldDefinitions
    } = req.body;
    const submittedDateRaw = submittedDateFromBody(req.body);
    const approvedDateRaw = approvedDateFromBody(req.body);

    const monthfilterRaw = normalizeMonthFilterValue(req.body);

    const sectorValue = normalizeStatutoryMetaValue(
      statutoryMetaFromBody(req.body, 'sector') ?? sector
    );
    const stateValue = normalizeStatutoryMetaValue(
      statutoryMetaFromBody(req.body, 'state') ?? state
    );
    const siteValue = normalizeStatutoryMetaValue(
      statutoryMetaFromBody(req.body, 'site') ?? site
    );

    // Validate required fields
    if (!formName || !String(formName).trim()) {
      return res.status(400).json({ status: 'failure', message: 'Form Name is required.' });
    }

    const { catalyst } = res.locals;
    const table = catalyst.datastore().table('Statutory');
    
    let effectiveProofFile = (proofSubmissionFile != null && proofSubmissionFile !== '') ? String(proofSubmissionFile).trim() : null;
    let effectiveProofFileName = (proofSubmissionFileName != null && proofSubmissionFileName !== '') ? String(proofSubmissionFileName).trim() : null;
    // Proof Submission is only populated by explicit PDF/image upload — never copy from Draft automatically.

    const normalizedDraftFile = normalizeFileRef(draftFile);
    const approvalNorm =
      approval != null && String(approval).trim() !== '' ? String(approval).trim().toLowerCase() : '';

    const insertData = {
      FormName: formName.trim(),
      Act: act || null,
      Description: description || null,
      DueDate: dueDate || null,
      Sector: sectorValue,
      State: stateValue,
      Site: siteValue,
      MonthFilter: trimMonthFilterForDatastore(monthfilterRaw),
      Autofill: autofill || null,
      Draft: draft || null,
      FormFile: normalizeFileRef(formFile),
      FormFileName: formFileName || null,
      SampleFile: normalizeFileRef(sampleFile) || normalizeFileRef(formFile),
      SampleFileName: sampleFileName || formFileName || null,
      ProofSubmissionFile: normalizeFileRef(effectiveProofFile),
      ProofSubmissionFileName: effectiveProofFileName || null,
      DraftFile: normalizedDraftFile,
      DraftFileName: draftFileName || null,
      SubmittedDate:
        submittedDateRaw !== undefined
          ? normalizeStatutoryDateValue(submittedDateRaw)
          : String(sendForApproval || '').trim().toLowerCase() === 'sent'
            ? todayIsoDate()
            : null,
      ApprovedDate:
        approvedDateRaw !== undefined
          ? normalizeStatutoryDateValue(approvedDateRaw)
          : approvalNorm === 'approved' || approvalNorm === 'approve'
            ? todayIsoDate()
            : null,
      Approval: approval != null && String(approval).trim() !== '' ? String(approval).trim() : null,
      Status: status != null && String(status).trim() !== '' ? String(status).trim() : null,
      SendForApproval:
        sendForApproval != null && String(sendForApproval).trim() !== '' ? String(sendForApproval).trim() : null,
      Remarks: remarks != null && String(remarks).trim() !== '' ? String(remarks).trim() : null
    };
    
    console.log('Inserting data to Statutory table:', JSON.stringify(insertData, null, 2));
    
    const insertResp = await insertStatutoryRow(table, insertData);
    console.log('Insert response:', JSON.stringify(insertResp, null, 2));
    
    const created = await table.getRow(insertResp.ROWID);
    console.log('Created record:', JSON.stringify(created, null, 2));
    
    const mappedRecord = mapStatutoryRowToApi(created);

    try {
      await persistSampleDataSnapshot(
        catalyst,
        {
          statutoryId: created.ROWID,
          formName: created.FormName || formName || null,
          monthfilter: created.MonthFilter || trimMonthFilterForDatastore(monthfilterRaw) || null,
          monthFilter: created.MonthFilter || trimMonthFilterForDatastore(monthfilterRaw) || null,
          MonthFilter: created.MonthFilter || trimMonthFilterForDatastore(monthfilterRaw) || null,
          headers: sampleDataHeader,
          headerFormData: sampleHeaderFormData,
          headerFieldDefinitions,
          rows: sampleData,
          statutoryDataTouchedRowIndexes: statutoryDataPayloadFromBody(req.body).statutoryDataTouchedRowIndexes,
          statutoryDataEditedCells: statutoryDataPayloadFromBody(req.body).statutoryDataEditedCells,
          statutoryDataEmployeeOrder: statutoryDataPayloadFromBody(req.body).statutoryDataEmployeeOrder
        },
        created
      );
    } catch (sampleErr) {
      // Do not block statutory save if SampleData write fails.
      console.warn('SampleData snapshot save failed:', sampleErr?.message || sampleErr);
    }

    await syncReturnedTableIfApplicable(catalyst, created, { requestBody: req.body });
    
    res.status(201).json({
      status: 'success',
      data: { statutory: mappedRecord }
    });
  } catch (err) {
    console.error('Error creating statutory record:', err);
    res.status(500).json({ status: 'failure', message: err.message || 'Failed to create statutory record.' });
  }
});

// Copy draft file to proof submission for a statutory record (server-side; no frontend download/upload)
app.post('/statutory/:id/copy-draft-to-proof', async (req, res) => {
  try {
    const { id } = req.params;
    console.log('Copy draft to proof: starting for statutory id', id);
    const { catalyst } = res.locals;
    const table = catalyst.datastore().table('Statutory');
    let row;
    try {
      row = await table.getRow(id);
    } catch (e) {
      console.error('Copy draft to proof: getRow failed', e.message);
      return res.status(404).json({ status: 'failure', message: 'Statutory record not found.' });
    }
    if (!row) {
      return res.status(404).json({ status: 'failure', message: 'Statutory record not found.' });
    }
    // If proof submission already exists, return it so frontend can send it in the update payload
    if (row.ProofSubmissionFile && row.ProofSubmissionFileName) {
      console.log('Copy draft to proof: proof already exists, returning existing', row.ProofSubmissionFile);
      return res.status(200).json({
        status: 'success',
        fileId: String(row.ProofSubmissionFile),
        fileName: row.ProofSubmissionFileName || ''
      });
    }
    const draftFolderId = DOC_TYPE_TO_FOLDER_ID['Draft'];
    const proofFolderId = DOC_TYPE_TO_FOLDER_ID['ProofSubmission'];
    if (!draftFolderId || !proofFolderId) {
      console.error('Copy draft to proof: missing folder config');
      return res.status(500).json({ status: 'failure', message: 'Filestore folder config missing.' });
    }
    if (!row.DraftFile || !row.DraftFileName) {
      console.log('Copy draft to proof: record has no draft', { hasDraftFile: !!row.DraftFile, hasDraftFileName: !!row.DraftFileName });
      return res.status(400).json({ status: 'failure', message: 'Record has no draft file to copy.' });
    }
    const draftFileId = row.DraftFile;
    const draftFileName = row.DraftFileName || 'proof.xlsx';
    const fileBuffer = await catalyst.filestore().folder(draftFolderId).downloadFile(draftFileId);
    const tempPath = path.join(os.tmpdir(), `proof_${id}_${Date.now()}_${draftFileName.replace(/[^a-zA-Z0-9._-]/g, '_')}`);
    fs.writeFileSync(tempPath, Buffer.isBuffer(fileBuffer) ? fileBuffer : Buffer.from(fileBuffer || []));
    let uploadResp;
    try {
      uploadResp = await catalyst.filestore().folder(proofFolderId).uploadFile({
        code: fs.createReadStream(tempPath),
        name: draftFileName
      });
    } finally {
      if (fs.existsSync(tempPath)) fs.unlinkSync(tempPath);
    }
    console.log('Copy draft to proof: upload response type', typeof uploadResp, Array.isArray(uploadResp) ? 'array' : '', uploadResp ? Object.keys(uploadResp) : '');
    let proofFileId = null;
    const proofFileName = draftFileName;
    if (Array.isArray(uploadResp) && uploadResp[0] != null) {
      proofFileId = uploadResp[0].id ?? uploadResp[0].ID ?? uploadResp[0].file_id;
    }
    if (!proofFileId && uploadResp && typeof uploadResp === 'object') {
      proofFileId = uploadResp.id ?? uploadResp.ID ?? uploadResp.file_id ?? uploadResp.fileId;
    }
    if (!proofFileId && uploadResp?.file_details) {
      const details = Array.isArray(uploadResp.file_details) ? uploadResp.file_details[0] : uploadResp.file_details;
      proofFileId = details?.id ?? details?.ID ?? details?.file_id;
    }
    if (!proofFileId) {
      console.error('Copy draft to proof: could not get file id from response', JSON.stringify(uploadResp));
      return res.status(500).json({ status: 'failure', message: 'Proof upload failed or invalid response.' });
    }
    proofFileId = String(proofFileId);
    await table.updateRow({
      ROWID: id,
      ProofSubmissionFile: proofFileId,
      ProofSubmissionFileName: proofFileName
    });
    console.log('Copy draft to proof: updated statutory', id, 'ProofSubmissionFile:', proofFileId, 'ProofSubmissionFileName:', proofFileName);
    res.status(200).json({ status: 'success', fileId: proofFileId, fileName: proofFileName });
  } catch (err) {
    console.error('Copy draft to proof error:', err);
    res.status(500).json({ status: 'failure', message: err.message || 'Copy draft to proof failed.' });
  }
});

// Update statutory record
app.put('/statutory/:id', async (req, res) => {
  try {
    const { id } = req.params;
    console.log('Received update request for ID:', id);
    console.log('Request body:', JSON.stringify(req.body, null, 2));
    
    const {
      formName,
      act,
      description,
      dueDate,
      sector,
      state,
      site,
      autofill,
      draft,
      formFile,
      formFileName,
      sampleFile,
      sampleFileName,
      draftFile,
      draftFileName,
      approval,
      status,
      sendForApproval,
      remarks,
      submittedDate,
      approvedDate,
      sampleDataHeader,
      sampleHeaderFormData,
      sampleData,
      headerFieldDefinitions
    } = req.body;
    const submittedDateRaw = submittedDateFromBody(req.body);
    const approvedDateRaw = approvedDateFromBody(req.body);
    const monthfilterRaw = normalizeMonthFilterValue(req.body);
    const sectorRaw = statutoryMetaFromBody(req.body, 'sector');
    const stateRaw = statutoryMetaFromBody(req.body, 'state');
    const siteRaw = statutoryMetaFromBody(req.body, 'site');
    // Support camelCase, PascalCase, and lowercase for proof (in case client or proxy normalizes keys)
    const proofSubmissionFile = req.body.proofSubmissionFile ?? req.body.ProofSubmissionFile ?? req.body.proofsubmissionfile;
    const proofSubmissionFileName = req.body.proofSubmissionFileName ?? req.body.ProofSubmissionFileName ?? req.body.proofsubmissionfilename;
    console.log('PUT statutory: proofSubmissionFile=', proofSubmissionFile, 'proofSubmissionFileName=', proofSubmissionFileName);

    const { catalyst } = res.locals;
    const table = catalyst.datastore().table('Statutory');
    
    // Check if record exists
    let existingRecord;
    try {
      existingRecord = await table.getRow(id);
    } catch (getRowErr) {
      console.error('Error fetching statutory record:', getRowErr);
      return res.status(404).json({ status: 'failure', message: 'Statutory record not found.' });
    }
    
    if (!existingRecord) {
      return res.status(404).json({ status: 'failure', message: 'Statutory record not found.' });
    }

    let effectiveProofFile = (proofSubmissionFile !== undefined && proofSubmissionFile !== null && proofSubmissionFile !== '') ? String(proofSubmissionFile).trim() : null;
    let effectiveProofFileName = (proofSubmissionFileName !== undefined && proofSubmissionFileName !== null && proofSubmissionFileName !== '') ? String(proofSubmissionFileName).trim() : null;
    if (effectiveProofFile === '' || effectiveProofFile === 'null' || effectiveProofFile === 'undefined') effectiveProofFile = null;
    if (effectiveProofFileName === '' || effectiveProofFileName === 'null' || effectiveProofFileName === 'undefined') effectiveProofFileName = null;
    // Proof Submission is only set from explicit upload; do not copy Draft to Proof here.

    const updateData = {
      ROWID: id
    };
    
    if (formName !== undefined) {
      updateData.FormName = formName.trim();
    }
    if (act !== undefined) updateData.Act = act;
    if (description !== undefined) updateData.Description = description;
    if (dueDate !== undefined) updateData.DueDate = dueDate;
    if (sectorRaw !== undefined || sector !== undefined) {
      updateData.Sector = normalizeStatutoryMetaValue(sectorRaw ?? sector);
    }
    if (stateRaw !== undefined || state !== undefined) {
      updateData.State = normalizeStatutoryMetaValue(stateRaw ?? state);
    }
    if (siteRaw !== undefined || site !== undefined) {
      updateData.Site = normalizeStatutoryMetaValue(siteRaw ?? site);
    }
    if (monthfilterRaw !== undefined) {
      updateData.MonthFilter = trimMonthFilterForDatastore(monthfilterRaw);
    }
    if (autofill !== undefined) updateData.Autofill = autofill;
    if (draft !== undefined) updateData.Draft = draft;
    // Preserve existing form file when client sends null/empty (e.g. Autofill save must not clear SEMaster form file)
    if (formFile !== undefined && formFile != null && String(formFile).trim() !== '') updateData.FormFile = normalizeFileRef(formFile);
    if (formFileName !== undefined && formFileName != null && String(formFileName).trim() !== '') updateData.FormFileName = formFileName;
    if (sampleFile !== undefined) {
      updateData.SampleFile = normalizeFileRef(sampleFile);
    }
    if (sampleFileName !== undefined) {
      updateData.SampleFileName = (sampleFileName != null && String(sampleFileName).trim() !== '') ? sampleFileName : null;
    }
    // Keep Sample in sync with Form when Sample not provided explicitly.
    if (sampleFile === undefined && updateData.FormFile !== undefined) {
      updateData.SampleFile = updateData.FormFile;
    }
    if (sampleFileName === undefined && updateData.FormFileName !== undefined) {
      updateData.SampleFileName = updateData.FormFileName;
    }
    // Always persist proof submission when provided (string IDs so backend stores reliably)
    if (proofSubmissionFile !== undefined || effectiveProofFile != null) {
      if (effectiveProofFile != null && String(effectiveProofFile).trim() !== '') {
        updateData.ProofSubmissionFile = normalizeFileRef(effectiveProofFile);
        updateData.ProofSubmissionFileName = (effectiveProofFileName != null && String(effectiveProofFileName).trim() !== '') ? String(effectiveProofFileName).trim() : null;
        console.log('PUT: persisting proof submission fileId=', updateData.ProofSubmissionFile, 'fileName=', updateData.ProofSubmissionFileName);
      } else {
        updateData.ProofSubmissionFile = null;
        updateData.ProofSubmissionFileName = null;
      }
    }
    if (draftFile !== undefined) updateData.DraftFile = normalizeFileRef(draftFile);
    if (draftFileName !== undefined) updateData.DraftFileName = draftFileName;
    if (approval !== undefined) {
      updateData.Approval = approval != null && String(approval).trim() !== '' ? String(approval).trim() : null;
    }
    if (status !== undefined) {
      updateData.Status = status != null && String(status).trim() !== '' ? String(status).trim() : null;
    }
    if (sendForApproval !== undefined) {
      updateData.SendForApproval =
        sendForApproval != null && String(sendForApproval).trim() !== ''
          ? String(sendForApproval).trim()
          : null;
      const sfaNorm = String(sendForApproval || '').trim().toLowerCase();
      if (sfaNorm === 'sent' && submittedDateRaw === undefined) {
        updateData.SubmittedDate = todayIsoDate();
      }
      if (sfaNorm === 'sent' && status === undefined) {
        updateData.Status = 'Pending';
      }
    }
    if (remarks !== undefined) {
      updateData.Remarks =
        remarks != null && String(remarks).trim() !== ''
          ? String(remarks).trim()
          : null;
    }
    if (submittedDateRaw !== undefined) {
      updateData.SubmittedDate = normalizeStatutoryDateValue(submittedDateRaw);
    }
    if (approvedDateRaw !== undefined) {
      updateData.ApprovedDate = normalizeStatutoryDateValue(approvedDateRaw);
    }
    if (updateData.Approval !== undefined) {
      const aNorm = String(updateData.Approval || '').trim().toLowerCase();
      if ((aNorm === 'approved' || aNorm === 'approve') && approvedDateRaw === undefined) {
        updateData.ApprovedDate = todayIsoDate();
      }
    }
    
    console.log('Updating Statutory record with data:', JSON.stringify(updateData, null, 2));

    // When importing a new file, remove old file from filestore to avoid orphaned data
    const newFormFile = updateData.FormFile != null ? String(updateData.FormFile).trim() : null;
    const newProofFile = updateData.ProofSubmissionFile != null ? String(updateData.ProofSubmissionFile).trim() : null;
    const newDraftFile = updateData.DraftFile != null ? String(updateData.DraftFile).trim() : null;
    const oldFormFile = existingRecord.FormFile ? String(existingRecord.FormFile).trim() : null;
    const oldProofFile = existingRecord.ProofSubmissionFile ? String(existingRecord.ProofSubmissionFile).trim() : null;
    const oldDraftFile = existingRecord.DraftFile ? String(existingRecord.DraftFile).trim() : null;

    if (newFormFile && oldFormFile && newFormFile !== oldFormFile) {
      try {
        await catalyst.filestore().folder(DOC_TYPE_TO_FOLDER_ID['Form']).deleteFile(oldFormFile);
        console.log('Removed old form file after import:', oldFormFile);
      } catch (fileErr) {
        console.warn('Could not delete old form file:', fileErr.message);
      }
    }
    if (newProofFile && oldProofFile && newProofFile !== oldProofFile) {
      try {
        await catalyst.filestore().folder(DOC_TYPE_TO_FOLDER_ID['ProofSubmission']).deleteFile(oldProofFile);
        console.log('Removed old proof submission file after import:', oldProofFile);
      } catch (fileErr) {
        console.warn('Could not delete old proof submission file:', fileErr.message);
      }
    }
    if (newDraftFile && oldDraftFile && newDraftFile !== oldDraftFile) {
      try {
        await catalyst.filestore().folder(DOC_TYPE_TO_FOLDER_ID['Draft']).deleteFile(oldDraftFile);
        console.log('Removed old draft file after import:', oldDraftFile);
      } catch (fileErr) {
        console.warn('Could not delete old draft file:', fileErr.message);
      }
    }
    
    await updateStatutoryRow(table, updateData);
    
    const updated = await table.getRow(id);
    
    const mappedRecord = mapStatutoryRowToApi(updated);

    const hasSamplePayload =
      (Array.isArray(sampleDataHeader) && sampleDataHeader.length > 0) ||
      (Array.isArray(sampleData) && sampleData.length > 0) ||
      (sampleHeaderFormData &&
        typeof sampleHeaderFormData === 'object' &&
        !Array.isArray(sampleHeaderFormData) &&
        Object.keys(sampleHeaderFormData).length > 0);

    let sampleDataSaved = false;
    if (hasSamplePayload) {
      try {
        const samplePersistResult = await persistSampleDataSnapshot(
          catalyst,
          {
            statutoryId: updated.ROWID || id,
            formName: updated.FormName || formName || null,
            monthfilter:
              updated.MonthFilter ||
              trimMonthFilterForDatastore(monthfilterRaw) ||
              existingRecord?.MonthFilter ||
              null,
            monthFilter:
              updated.MonthFilter ||
              trimMonthFilterForDatastore(monthfilterRaw) ||
              existingRecord?.MonthFilter ||
              null,
            MonthFilter:
              updated.MonthFilter ||
              trimMonthFilterForDatastore(monthfilterRaw) ||
              existingRecord?.MonthFilter ||
              null,
            headers: sampleDataHeader,
            headerFormData: sampleHeaderFormData,
            headerFieldDefinitions,
            rows: sampleData,
            statutoryDataTouchedRowIndexes: statutoryDataPayloadFromBody(req.body).statutoryDataTouchedRowIndexes,
            statutoryDataEditedCells: statutoryDataPayloadFromBody(req.body).statutoryDataEditedCells,
            statutoryDataEmployeeOrder: statutoryDataPayloadFromBody(req.body).statutoryDataEmployeeOrder
          },
          updated
        );
        sampleDataSaved = !!samplePersistResult?.saved;
      } catch (sampleErr) {
        // Do not block statutory save if SampleData write fails.
        console.warn('SampleData snapshot save failed:', sampleErr?.message || sampleErr);
      }
    }

    await syncReturnedTableIfApplicable(catalyst, updated, {
      existingRecord,
      requestBody: req.body
    });
    
    res.status(200).json({
      status: 'success',
      data: { statutory: mappedRecord, sampleDataSaved }
    });
  } catch (err) {
    console.error('Error updating statutory record:', err);
    res.status(500).json({ status: 'failure', message: err.message || 'Failed to update statutory record.' });
  }
});

// Delete statutory record
app.delete('/statutory/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const { catalyst } = res.locals;
    const table = catalyst.datastore().table('Statutory');
    
    // Check if record exists before deleting
    let existingRecord;
    try {
      existingRecord = await table.getRow(id);
    } catch (getRowErr) {
      console.error('Error fetching statutory record:', getRowErr);
      return res.status(404).json({ status: 'failure', message: 'Statutory record not found.' });
    }
    
    if (!existingRecord) {
      return res.status(404).json({ status: 'failure', message: 'Statutory record not found.' });
    }
    
    // Optionally delete associated files from filestore
    // Note: This is optional - you may want to keep files even if record is deleted
    const { catalyst: cat } = res.locals;
    if (existingRecord.FormFile) {
      try {
        await cat.filestore().folder(DOC_TYPE_TO_FOLDER_ID['Form']).deleteFile(existingRecord.FormFile);
        console.log('Deleted form file:', existingRecord.FormFile);
      } catch (fileErr) {
        console.warn('Could not delete form file:', fileErr.message);
      }
    }
    if (existingRecord.ProofSubmissionFile) {
      try {
        await cat.filestore().folder(DOC_TYPE_TO_FOLDER_ID['ProofSubmission']).deleteFile(existingRecord.ProofSubmissionFile);
        console.log('Deleted proof submission file:', existingRecord.ProofSubmissionFile);
      } catch (fileErr) {
        console.warn('Could not delete proof submission file:', fileErr.message);
      }
    }
    if (existingRecord.DraftFile) {
      try {
        await cat.filestore().folder(DOC_TYPE_TO_FOLDER_ID['Draft']).deleteFile(existingRecord.DraftFile);
        console.log('Deleted draft file:', existingRecord.DraftFile);
      } catch (fileErr) {
        console.warn('Could not delete draft file:', fileErr.message);
      }
    }
    
    await table.deleteRow(id);
    
    res.status(200).json({ status: 'success', data: { id } });
  } catch (err) {
    console.error('Error deleting statutory record:', err);
    res.status(500).json({ status: 'failure', message: err.message || 'Failed to delete statutory record.' });
  }
});

// Download file endpoint
app.get('/statutory/:id/file/:docType', async (req, res) => {
  try {
    const executionStartTime = Date.now();
    console.log('Execution started at:', executionStartTime);
    
    const { id, docType } = req.params;
    console.log('File download request:', { id, docType });
    const { catalyst } = res.locals;
    const folderId = DOC_TYPE_TO_FOLDER_ID[docType];
    console.log('Folder ID for docType:', docType, 'is:', folderId);
    
    if (!folderId) {
      console.error('Invalid document type:', docType);
      return res.status(400).json({ status: 'failure', message: 'Invalid document type.' });
    }
    
    const table = catalyst.datastore().table('Statutory');
    let statutory;
    try {
      statutory = await table.getRow(id);
    } catch (getRowErr) {
      console.error('Error fetching statutory record:', getRowErr);
      return res.status(404).json({ status: 'failure', message: 'Statutory record not found.' });
    }
    
    if (!statutory) {
      console.error('Statutory record not found for ID:', id);
      return res.status(404).json({ status: 'failure', message: 'Statutory record not found.' });
    }
    
    let fileId, fileName;
    if (docType === 'Form') {
      fileId = statutory.FormFile;
      fileName = statutory.FormFileName || 'form_file';
    } else if (docType === 'ProofSubmission') {
      fileId = statutory.ProofSubmissionFile;
      fileName = statutory.ProofSubmissionFileName || 'proof_submission_file';
    } else if (docType === 'Draft') {
      fileId = statutory.DraftFile;
      fileName = statutory.DraftFileName || 'draft_file';
    } else {
      return res.status(400).json({ status: 'failure', message: 'Invalid document type.' });
    }
    
    console.log('File info:', { fileId, fileName });
    const disposition = req.query.disposition === 'attachment' ? 'attachment' : 'inline';
    const normalizedFormName = normalizeFormKey(statutory.FormName || statutory.formName);
    const templatePath = FORM_TEMPLATE_MAP[normalizedFormName];
    const useTemplateForProof = docType === 'ProofSubmission' && req.query.format === 'template';

    const serveTemplateIfAvailable = (reason) => {
      if (!templatePath || !fs.existsSync(templatePath)) {
        return false;
      }
      try {
        const templateBuffer = fs.readFileSync(templatePath);
        const templateFileName = fileName || path.basename(templatePath);
        const contentType = getContentType(templateFileName);
        console.warn(
          `⚠️ Serving fallback template "${templateFileName}" for form "${statutory.FormName}" because ${reason}`
        );
        res.writeHead(200, {
          'Content-Type': contentType,
          'Content-Disposition': `${disposition}; filename="${templateFileName}"`,
          'Content-Length': templateBuffer.length
        });
        res.end(templateBuffer);
        return true;
      } catch (templateErr) {
        console.error('Error serving fallback template:', templateErr);
        return false;
      }
    };
    
    if (!fileId) {
      console.error('No file ID found for statutory:', id, 'docType:', docType);
      if (useTemplateForProof) {
        try {
          const { templateBuffer, templateFileName: tfName } = await tryLoadProofFormTemplateBuffer(
            catalyst,
            statutory,
            normalizedFormName,
            templatePath,
            req
          );
          if (templateBuffer) {
            const outputBuffer = await mergeProofTemplateToBuffer(templateBuffer, null, null, req);
            res.writeHead(200, {
              'Content-Type': getContentType(tfName),
              'Content-Disposition': `${disposition}; filename="${tfName}"`,
              'Content-Length': outputBuffer.length
            });
            res.end(outputBuffer);
            return;
          }
        } catch (noFileProofErr) {
          console.error('Proof template without stored proof file failed:', noFileProofErr);
        }
      }
      if (useTemplateForProof && serveTemplateIfAvailable('no proof submission file found')) {
        return;
      }
      if (serveTemplateIfAvailable('no file ID stored in database')) {
        return;
      }
      return res.status(404).json({ status: 'failure', message: 'File not found for this statutory record.' });
    }

    if (useTemplateForProof) {
      const { templateBuffer, templateFileName: tfName } = await tryLoadProofFormTemplateBuffer(
        catalyst,
        statutory,
        normalizedFormName,
        templatePath,
        req
      );

      let proofBuffer = null;
      try {
        proofBuffer = await catalyst.filestore().folder(folderId).downloadFile(fileId);
      } catch (downloadErr) {
        console.error('Error downloading proof submission file:', downloadErr);
      }

      if (templateBuffer) {
        try {
          const outputBuffer = await mergeProofTemplateToBuffer(templateBuffer, proofBuffer, fileName, req);
          res.writeHead(200, {
            'Content-Type': getContentType(tfName),
            'Content-Disposition': `${disposition}; filename="${tfName}"`,
            'Content-Length': outputBuffer.length
          });
          res.end(outputBuffer);
          return;
        } catch (fillErr) {
          console.error('Error applying proof submission data to template:', fillErr);
        }
      }
    }
    
    let fileBuffer;
    try {
      fileBuffer = await catalyst.filestore().folder(folderId).downloadFile(fileId);
    } catch (downloadErr) {
      console.error('Error downloading file from filestore:', downloadErr);
      
      // Check if it's a file not found error
      if (isFileMissingError(downloadErr)) {
        console.error(`File with ID ${fileId} does not exist in folder ${folderId}`);
        if (useTemplateForProof && serveTemplateIfAvailable('proof submission file missing from filestore')) {
          return;
        }
        if (serveTemplateIfAvailable('file missing from filestore')) {
          return;
        }
        // For Draft: clear broken file reference in DB so edit form can open and user can re-upload
        if (docType === 'Draft') {
          try {
            const table = catalyst.datastore().table('Statutory');
            await table.updateRow({
              ROWID: id,
              DraftFile: null,
              DraftFileName: null
            });
            console.log(`Cleared missing Draft file reference for statutory record ${id}`);
          } catch (clearErr) {
            console.error('Error clearing draft file reference:', clearErr);
          }
        }
        return res.status(404).json({ 
          status: 'failure', 
          message: 'File no longer exists in storage. The record has been updated. Please re-upload the file from the Edit form.'
        });
      }
      
      // Re-throw other errors to be caught by outer catch
      throw downloadErr;
    }
    
    const contentType = getContentType(fileName);
    
    console.log('Sending file:', { fileName, contentType, fileSize: fileBuffer.length, disposition });
    res.writeHead(200, {
      'Content-Type': contentType,
      'Content-Disposition': `${disposition}; filename="${fileName}"`,
      'Content-Length': fileBuffer.length
    });
    res.end(fileBuffer);
  } catch (err) {
    console.error('Error downloading file:', err);
    
    // Check if it's a file not found error that wasn't caught earlier
    if (err.statusCode === 404 || err.code === 'INVALID_ID' || 
        (err.message && err.message.includes('No such file'))) {
      return res.status(404).json({ 
        status: 'failure', 
        message: `File no longer exists in storage. The file may have been deleted or moved. Please re-upload the file.` 
      });
    }
    
    res.status(500).json({ status: 'failure', message: err.message || 'File download failed.' });
  }
});

// Delete file endpoint for statutory and docType
app.delete('/statutory/:id/file/:docType', async (req, res) => {
  try {
    const { id, docType } = req.params;
    const { catalyst } = res.locals;
    const folderId = DOC_TYPE_TO_FOLDER_ID[docType];
    
    if (!folderId) {
      return res.status(400).json({ status: 'failure', message: 'Invalid document type.' });
    }
    
    const table = catalyst.datastore().table('Statutory');
    const statutory = await table.getRow(id);
    
    let fileId;
    if (docType === 'Form') {
      fileId = statutory.FormFile;
    } else if (docType === 'ProofSubmission') {
      fileId = statutory.ProofSubmissionFile;
    } else if (docType === 'Draft') {
      fileId = statutory.DraftFile;
    } else {
      return res.status(400).json({ status: 'failure', message: 'Invalid document type.' });
    }
    
    if (!fileId) {
      return res.status(404).json({ status: 'failure', message: 'File not found for this statutory record.' });
    }
    
    // Delete file from file store
    await catalyst.filestore().folder(folderId).deleteFile(fileId);
    
    // Clear the appropriate column in the database
    const updateData = { ROWID: id };
    if (docType === 'Form') {
      updateData.FormFile = null;
      updateData.FormFileName = null;
    } else if (docType === 'ProofSubmission') {
      updateData.ProofSubmissionFile = null;
      updateData.ProofSubmissionFileName = null;
    } else if (docType === 'Draft') {
      updateData.DraftFile = null;
      updateData.DraftFileName = null;
    }
    await table.updateRow(updateData);
    
    res.status(200).json({ status: 'success' });
  } catch (err) {
    console.log(err);
    res.status(500).json({ status: 'failure', message: err.message || 'File delete failed.' });
  }
});

module.exports = app;