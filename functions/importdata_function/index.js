'use strict';

const express = require('express');
const catalystSDK = require('zcatalyst-sdk-node');

const app = express();
const IMPORT_DATA_TABLE = 'ImportData';

const MONTH_NAMES = [
  'January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December'
];

app.use(express.json({ limit: '50mb' }));

app.use((req, res, next) => {
  try {
    res.locals.catalyst = catalystSDK.initialize(req);
    next();
  } catch (err) {
    res.status(500).json({ status: 'failure', message: 'Catalyst init failed' });
  }
});

app.get('/', (req, res) => {
  res.status(200).json({ status: 'success', message: 'importdata_function ready' });
});

function isDatastoreColumnError(err) {
  const msg = String(err?.message || '').toLowerCase();
  return (
    err?.code === 'INVALID_INPUT' ||
    msg.includes('column name') ||
    msg.includes('invalid column') ||
    msg.includes('no such column')
  );
}

function trimMonthFilterForDatastore(value) {
  if (value == null) return null;
  const s = String(value).trim();
  if (s === '' || s === 'null' || s === 'undefined') return null;
  return s;
}

function normalizeMonthFilterValue(body) {
  if (!body || typeof body !== 'object') return undefined;
  if ('monthfilter' in body) return body.monthfilter;
  if ('monthFilter' in body) return body.monthFilter;
  if ('MonthFilter' in body) return body.MonthFilter;
  return undefined;
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

function deriveMonthFilterFromDueDate(dueDate) {
  if (!dueDate) return null;
  const s = String(dueDate).trim();
  if (!s) return null;
  const iso = s.match(/^(\d{4})-(\d{2})-\d{2}/);
  if (iso) {
    const monthIdx = Number(iso[2]) - 1;
    if (monthIdx >= 0 && monthIdx < 12) return MONTH_NAMES[monthIdx];
  }
  const d = new Date(s);
  if (!Number.isNaN(d.getTime())) return MONTH_NAMES[d.getMonth()];
  return null;
}

function resolveMonthFilterForSave(payload, statutoryRow) {
  const fromPayload = resolveMonthFilterFromSource(payload);
  if (fromPayload) return fromPayload;
  const fromStatutory = trimMonthFilterForDatastore(statutoryRow?.MonthFilter);
  if (fromStatutory) return fromStatutory;
  const fromDue = deriveMonthFilterFromDueDate(statutoryRow?.DueDate);
  if (fromDue) return fromDue;
  return MONTH_NAMES[new Date().getMonth()];
}

function pickImportDataRow(entry) {
  if (!entry || typeof entry !== 'object') return null;
  return entry.ImportData || entry.importdata || entry;
}

/** Save statutory grid snapshot into ImportData (Header, Data, MonthFilter). */
async function persistImportDataSnapshot(catalyst, payload, statutoryRow = null) {
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

  const monthFilterValue = resolveMonthFilterForSave(payload, statutoryRow);
  const table = catalyst.datastore().table(IMPORT_DATA_TABLE);
  const headerPayload = {
    formName: payload?.formName || null,
    statutoryId,
    monthfilter: monthFilterValue,
    headers,
    headerFormData
  };
  const dataPayload = { rows };
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
      console.warn('ImportData insert: retried without MonthFilter column:', err.message);
    } else {
      console.error('ImportData insert failed:', err.message || err, {
        statutoryId,
        rowCount: rows.length,
        monthFilter: monthFilterValue
      });
      throw err;
    }
  }

  return { saved: true, rowId: insertResp?.ROWID || null, monthFilter: monthFilterValue };
}

function normalizeStatutoryIdForMatch(value) {
  const s = String(value || '').trim();
  if (!/^\d+$/.test(s)) return '0';
  return s.replace(/^0+/, '') || '0';
}

function safeJsonParse(value, fallback = null) {
  if (value == null || value === '') return fallback;
  if (typeof value === 'object') return value;
  try {
    return JSON.parse(String(value));
  } catch (_) {
    return fallback;
  }
}

function gridDataMonthsMatch(stored, query) {
  if (!query) return true;
  if (!stored) return false;
  const a = String(stored).trim().toLowerCase();
  const b = String(query).trim().toLowerCase();
  if (a === b) return true;
  if (a.slice(0, 3) === b.slice(0, 3)) return true;
  return false;
}

function gridDataFormsMatch(storedForm, queryForm) {
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
}

function resolveImportDataRowMonthFilter(row, headerObj) {
  return (
    resolveMonthFilterFromSource(row) ||
    trimMonthFilterForDatastore(headerObj?.monthfilter) ||
    null
  );
}

// Save imported/autofill grid snapshot into ImportData
app.post('/importdata/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const idStr = String(id || '').trim();
    if (!/^\d+$/.test(idStr)) {
      return res.status(400).json({ status: 'failure', message: 'Invalid statutory id for ImportData save.' });
    }

    const { catalyst } = res.locals;
    const statutoryTable = catalyst.datastore().table('Statutory');

    let existingRecord = null;
    try {
      existingRecord = await statutoryTable.getRow(idStr);
    } catch (getRowErr) {
      console.warn('ImportData save: statutory row lookup failed, saving snapshot anyway:', getRowErr.message);
    }

    const {
      sampleDataHeader,
      sampleHeaderFormData,
      sampleData,
      formName,
      headers: headersAlt,
      rows: rowsAlt,
      headerFormData: headerFormDataAlt
    } = req.body;
    const monthfilterRaw = normalizeMonthFilterValue(req.body);

    const headers = Array.isArray(sampleDataHeader)
      ? sampleDataHeader
      : Array.isArray(headersAlt)
        ? headersAlt
        : [];
    const rows = Array.isArray(sampleData)
      ? sampleData
      : Array.isArray(rowsAlt)
        ? rowsAlt
        : [];
    const headerFormData =
      sampleHeaderFormData && typeof sampleHeaderFormData === 'object' && !Array.isArray(sampleHeaderFormData)
        ? sampleHeaderFormData
        : headerFormDataAlt && typeof headerFormDataAlt === 'object' && !Array.isArray(headerFormDataAlt)
          ? headerFormDataAlt
          : {};

    if (headers.length === 0 && rows.length === 0 && Object.keys(headerFormData).length === 0) {
      return res.status(400).json({ status: 'failure', message: 'No import data to save.' });
    }

    const persistResult = await persistImportDataSnapshot(
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
        rows
      },
      existingRecord
    );

    if (!persistResult?.saved) {
      return res.status(400).json({
        status: 'failure',
        message: 'Could not save ImportData snapshot.',
        reason: persistResult?.reason || 'unknown'
      });
    }

    return res.status(200).json({
      status: 'success',
      message: 'ImportData snapshot saved.',
      data: {
        statutoryId: String(existingRecord?.ROWID || idStr),
        rowCount: rows.length,
        importDataRowId: persistResult.rowId,
        MonthFilter: persistResult.monthFilter || null,
        monthfilter: persistResult.monthFilter || null
      }
    });
  } catch (err) {
    console.error('Error saving ImportData snapshot:', err);
    return res.status(500).json({
      status: 'failure',
      message: err.message || 'Failed to save ImportData snapshot.'
    });
  }
});

// Fetch latest ImportData snapshot for a statutory record (same matching rules as SampleData GET)
app.get('/importdata/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const { catalyst } = res.locals;
    const zcql = catalyst.zcql();
    const query = `SELECT ROWID, Header, Data, MonthFilter, MODIFIEDTIME, CREATEDTIME FROM ${IMPORT_DATA_TABLE}`;
    const result = await zcql.executeZCQLQuery(query);

    const rows = Array.isArray(result)
      ? result.map((entry) => pickImportDataRow(entry)).filter(Boolean)
      : [];

    const idStr = String(id).trim();
    const idNorm = normalizeStatutoryIdForMatch(idStr);
    const formNameQ = String(req.query?.formName || req.query?.formname || '').trim().toLowerCase();
    const monthQ = trimMonthFilterForDatastore(
      req.query?.monthFilter || req.query?.monthfilter || req.query?.month || ''
    );
    const matchByFormMonthOnly =
      req.query?.matchByFormMonth === '1' ||
      req.query?.matchByFormMonth === 'true' ||
      (formNameQ && (idStr === '0' || !/^\d+$/.test(idStr)));

    const filterByFormMonth = (mappedRows) =>
      mappedRows.filter(
        (r) =>
          gridDataFormsMatch(r.formName, formNameQ) &&
          gridDataMonthsMatch(r.monthfilter, monthQ)
      );

    const mapImportRow = (row) => {
      const headerObj = safeJsonParse(row?.Header, {});
      const dataObj = safeJsonParse(row?.Data, {});
      const statutoryId = headerObj?.statutoryId != null ? String(headerObj.statutoryId).trim() : '';
      const headers = Array.isArray(headerObj?.headers) ? headerObj.headers : [];
      const headerFormData =
        headerObj?.headerFormData && typeof headerObj.headerFormData === 'object' && !Array.isArray(headerObj.headerFormData)
          ? headerObj.headerFormData
          : {};
      const tableRows = Array.isArray(dataObj?.rows) ? dataObj.rows : [];
      const monthfilter = resolveImportDataRowMonthFilter(row, headerObj);
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

    const allMapped = rows.map(mapImportRow);
    let matching = [];

    if (formNameQ && (matchByFormMonthOnly || monthQ)) {
      matching = filterByFormMonth(allMapped);
    }

    if (matching.length === 0 && formNameQ && !matchByFormMonthOnly) {
      matching = allMapped.filter((r) => gridDataFormsMatch(r.formName, formNameQ));
    }

    if (matching.length === 0 && !matchByFormMonthOnly && /^\d+$/.test(idStr) && idStr !== '0') {
      matching = allMapped.filter(
        (r) =>
          r.statutoryId === idStr ||
          normalizeStatutoryIdForMatch(r.statutoryId) === idNorm
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
              return gridDataMonthsMatch(r.monthfilter, statMonth);
            }
            return true;
          });
        }
      } catch (statLookupErr) {
        console.warn('ImportData GET: statutory row lookup fallback skipped:', statLookupErr.message);
      }
    }

    if (matching.length === 0 && formNameQ) {
      matching = filterByFormMonth(allMapped);
      if (matching.length === 0) {
        matching = allMapped.filter((r) => gridDataFormsMatch(r.formName, formNameQ));
      }
    }

    if (matching.length === 0) {
      return res.status(200).json({ status: 'success', data: { importData: null } });
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
        importData: {
          statutoryId: latest.statutoryId,
          formName: latest.formName,
          monthfilter: latest.monthfilter,
          MonthFilter: latest.monthfilter,
          headers: latest.headers,
          headerFormData: latest.headerFormData,
          rows: latest.rows
        },
        importDataRowId: latest.rowid || null
      }
    });
  } catch (err) {
    console.error('Error fetching ImportData snapshot:', err);
    return res.status(500).json({
      status: 'failure',
      message: err.message || 'Failed to fetch ImportData.'
    });
  }
});

module.exports = app;
