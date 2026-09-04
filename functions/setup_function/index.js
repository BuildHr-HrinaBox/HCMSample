'use strict';

const express = require('express');
const catalystSDK = require('zcatalyst-sdk-node');

const app = express();
const TABLE_NAME = 'SetUp';

app.use(express.json({ limit: '2mb' }));

app.use((req, res, next) => {
  try {
    res.locals.catalyst = catalystSDK.initialize(req);
    next();
  } catch (err) {
    res.status(500).json({ status: 'failure', message: 'Catalyst init failed' });
  }
});

function pickRow(rowObj) {
  if (!rowObj || typeof rowObj !== 'object' || Array.isArray(rowObj)) return null;
  const nested =
    rowObj.SetUp || rowObj.setup || rowObj.SETUP || rowObj.Setup || null;
  if (nested && typeof nested === 'object' && !Array.isArray(nested)) {
    return nested;
  }
  const keys = Object.keys(rowObj);
  if (keys.length === 1) {
    const only = rowObj[keys[0]];
    if (only && typeof only === 'object' && !Array.isArray(only)) return only;
  }
  return rowObj;
}

function trimText(value) {
  return String(value ?? '').trim();
}

function normalizeFieldName(name) {
  return String(name || '')
    .toLowerCase()
    .replace(/[_\s-]/g, '');
}

function fieldFrom(data, names) {
  if (!data || typeof data !== 'object') return '';
  for (const name of names) {
    if (data[name] != null && trimText(data[name]) !== '') return trimText(data[name]);
  }
  const lookup = {};
  Object.keys(data).forEach((key) => {
    lookup[normalizeFieldName(key)] = data[key];
  });
  for (const name of names) {
    const value = lookup[normalizeFieldName(name)];
    if (value != null && trimText(value) !== '') return trimText(value);
  }
  return '';
}

function omitKeys(obj, keys) {
  const next = { ...obj };
  keys.forEach((key) => {
    delete next[key];
  });
  return next;
}

function isDatastoreColumnError(err) {
  const msg = String(err?.message || err || '').toLowerCase();
  return (
    msg.includes('column') ||
    msg.includes('invalid') ||
    msg.includes('unknown') ||
    msg.includes('not exist')
  );
}

async function upsertSetupRow(table, payload, isUpdate) {
  const attempts = [
    payload,
    omitKeys(payload, ['IndustryType']),
    (() => {
      const next = omitKeys(payload, ['IndustryType']);
      if (next.Role != null && next.RoleName == null) next.RoleName = next.Role;
      return next;
    })(),
    omitKeys(payload, ['IndustryType', 'Role']),
  ];
  let lastErr;
  for (let i = 0; i < attempts.length; i += 1) {
    try {
      return isUpdate ? await table.updateRow(attempts[i]) : await table.insertRow(attempts[i]);
    } catch (err) {
      lastErr = err;
      if (!isDatastoreColumnError(err) || i >= attempts.length - 1) throw err;
    }
  }
  throw lastErr;
}

function matchKey(state, site, formName, industryType = '', act = '', description = '') {
  return [
    trimText(state),
    trimText(site),
    trimText(formName),
    trimText(industryType),
    trimText(act),
    trimText(description),
  ]
    .map((part) => part.toLowerCase())
    .join('|');
}

function legacyMatchKey(state, site, formName) {
  return `${trimText(state).toLowerCase()}|${trimText(site).toLowerCase()}|${trimText(formName).toLowerCase()}`;
}

function emailsToStore(value) {
  if (Array.isArray(value)) {
    return value.map(trimText).filter(Boolean).join(', ');
  }
  return trimText(value);
}

function toAppRow(row) {
  const data = pickRow(row) || {};
  const id =
    fieldFrom(data, ['ROWID', 'rowId', 'rowid']) ||
    fieldFrom(row || {}, ['ROWID', 'rowId', 'rowid']);
  return {
    id: id !== '' ? id : null,
    state: fieldFrom(data, ['State', 'state']),
    site: fieldFrom(data, ['Site', 'site', 'SiteName', 'siteName']),
    formName: fieldFrom(data, ['FormName', 'formName', 'Name', 'name']),
    email: fieldFrom(data, ['Email', 'email', 'EmailId', 'emailId', 'Emails', 'emails']),
    act: fieldFrom(data, ['Act', 'act']),
    description: fieldFrom(data, ['Description', 'description']),
    role: fieldFrom(data, ['Role', 'role', 'RoleName', 'roleName']),
    industryType: fieldFrom(data, ['IndustryType', 'industryType', 'Industry', 'industry']),
  };
}

async function fetchRowsViaZcql(catalyst) {
  const zcql = catalyst.zcql();
  const all = [];
  const pageSize = 100;
  let offset = 0;
  while (offset < 5000) {
    const chunk = await zcql.executeZCQLQuery(
      `SELECT * FROM ${TABLE_NAME} LIMIT ${offset},${pageSize}`
    );
    if (!Array.isArray(chunk) || chunk.length === 0) break;
    chunk.forEach((row) => all.push(row));
    if (chunk.length < pageSize) break;
    offset += pageSize;
  }
  return all;
}

async function fetchAllSetupRows(catalyst) {
  let raw = [];
  try {
    raw = await fetchRowsViaZcql(catalyst);
  } catch (_) {
    raw = [];
  }
  if (!raw.length) {
    const table = catalyst.datastore().table(TABLE_NAME);
    raw = (await table.getAllRows()) || [];
  }
  return raw
    .map(toAppRow)
    .filter((row) => row.formName && (row.id || row.role || row.email));
}

app.get('/', (req, res) => {
  res.status(200).json({ status: 'success', message: 'setup_function ready' });
});

app.get('/setup', async (req, res) => {
  try {
    const { catalyst } = res.locals;
    const stateFilter = trimText(req.query.state);
    const siteFilter = trimText(req.query.site);
    let data = await fetchAllSetupRows(catalyst);
    if (stateFilter) {
      const want = stateFilter.toLowerCase();
      data = data.filter((row) => {
        const state = String(row.state || '').trim().toLowerCase();
        return state === want || state.includes(want) || want.includes(state);
      });
    }
    if (siteFilter) {
      const want = siteFilter.toLowerCase();
      data = data.filter((row) => {
        const site = String(row.site || '').trim().toLowerCase();
        if (!site) return false;
        return site === want || site.includes(want) || want.includes(site);
      });
    }
    res.status(200).json({ status: 'success', data });
  } catch (err) {
    res.status(500).json({ status: 'failure', message: err.message || 'Failed to load Setup records.' });
  }
});

app.put('/setup', async (req, res) => {
  try {
    const { catalyst } = res.locals;
    const state = trimText(req.body?.state);
    const site = trimText(req.body?.site);
    const industryTypeFromBody = trimText(req.body?.industryType ?? req.body?.IndustryType);
    const rows = Array.isArray(req.body?.rows) ? req.body.rows : [];

    if (!state) {
      return res.status(400).json({ status: 'failure', message: 'State is required.' });
    }
    if (!rows.length) {
      return res.status(400).json({ status: 'failure', message: 'At least one form row is required.' });
    }

    const table = catalyst.datastore().table(TABLE_NAME);
    const existing = await fetchAllSetupRows(catalyst);
    const byKey = new Map();
    const byLegacyKey = new Map();
    existing.forEach((row) => {
      byKey.set(
        matchKey(row.state, row.site, row.formName, row.industryType, row.act, row.description),
        row
      );
      const legacy = legacyMatchKey(row.state, row.site, row.formName);
      if (!byLegacyKey.has(legacy)) byLegacyKey.set(legacy, row);
    });

    const saved = [];
    for (const item of rows) {
      const formName = trimText(item?.formName || item?.FormName);
      if (!formName) continue;
      const email = emailsToStore(item?.emails ?? item?.email ?? item?.Email);
      const act = trimText(item?.act ?? item?.Act);
      const description = trimText(item?.description ?? item?.Description);
      const role = trimText(item?.role ?? item?.Role);
      const industryType = trimText(
        item?.industryType ?? item?.IndustryType ?? industryTypeFromBody
      );
      const key = matchKey(state, site, formName, industryType, act, description);
      const current =
        byKey.get(key) ||
        (!industryType && !act && !description
          ? byLegacyKey.get(legacyMatchKey(state, site, formName))
          : null) ||
        (industryType || act || description
          ? (() => {
              const legacy = byLegacyKey.get(legacyMatchKey(state, site, formName));
              if (!legacy) return null;
              const sameIndustry =
                !industryType ||
                !legacy.industryType ||
                industryType.toLowerCase() === legacy.industryType.toLowerCase();
              const sameAct =
                !act || !legacy.act || act.toLowerCase() === legacy.act.toLowerCase();
              const sameDescription =
                !description ||
                !legacy.description ||
                description.toLowerCase() === legacy.description.toLowerCase();
              return sameIndustry && sameAct && sameDescription ? legacy : null;
            })()
          : null);
      const rowPayload = {
        State: state,
        Site: site,
        FormName: formName,
        Email: email,
        Act: act,
        Description: description,
        Role: role,
        IndustryType: industryType || current?.industryType || '',
      };
      let stored;
      if (current?.id) {
        stored = await upsertSetupRow(table, { ROWID: current.id, ...rowPayload }, true);
      } else {
        stored = await upsertSetupRow(table, rowPayload, false);
      }
      const appRow = toAppRow(stored);
      const resolvedIndustry =
        appRow.industryType || industryType || current?.industryType || '';
      const resolvedAct = appRow.act || act;
      const resolvedDescription = appRow.description || description;
      const savedRow = {
        id: appRow.id || current?.id || null,
        state,
        site,
        formName,
        email,
        act: resolvedAct,
        description: resolvedDescription,
        role: appRow.role || role,
        industryType: resolvedIndustry,
      };
      saved.push(savedRow);
      byKey.set(
        matchKey(state, site, formName, resolvedIndustry, resolvedAct, resolvedDescription),
        { ...savedRow, id: savedRow.id }
      );
    }

    res.status(200).json({
      status: 'success',
      message: `Saved ${saved.length} Setup record(s).`,
      data: saved,
    });
  } catch (err) {
    res.status(500).json({ status: 'failure', message: err.message || 'Failed to save Setup records.' });
  }
});

module.exports = app;
