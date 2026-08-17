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
  if (!rowObj || typeof rowObj !== 'object') return null;
  return rowObj.SetUp || rowObj.setup || rowObj.SETUP || rowObj;
}

function trimText(value) {
  return String(value ?? '').trim();
}

function matchKey(state, site, formName) {
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
  return {
    id: data.ROWID || data.rowId || null,
    state: trimText(data.State ?? data.state),
    site: trimText(data.Site ?? data.site),
    formName: trimText(data.FormName ?? data.formName),
    email: trimText(data.Email ?? data.email),
    act: trimText(data.Act ?? data.act),
    description: trimText(data.Description ?? data.description),
    role: trimText(data.Role ?? data.role),
  };
}

async function fetchAllSetupRows(catalyst) {
  const table = catalyst.datastore().table(TABLE_NAME);
  const rows = await table.getAllRows();
  return (rows || []).map(toAppRow).filter((row) => row.id);
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
      data = data.filter((row) => row.state.toLowerCase() === want);
    }
    if (siteFilter) {
      const want = siteFilter.toLowerCase();
      data = data.filter((row) => row.site.toLowerCase() === want);
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
    const rows = Array.isArray(req.body?.rows) ? req.body.rows : [];

    if (!state) {
      return res.status(400).json({ status: 'failure', message: 'State is required.' });
    }
    if (!rows.length) {
      return res.status(400).json({ status: 'failure', message: 'At least one form row is required.' });
    }

    const table = catalyst.datastore().table(TABLE_NAME);
    const existing = await fetchAllSetupRows(catalyst);
    const byKey = new Map(existing.map((row) => [matchKey(row.state, row.site, row.formName), row]));

    const saved = [];
    for (const item of rows) {
      const formName = trimText(item?.formName || item?.FormName);
      if (!formName) continue;
      const email = emailsToStore(item?.emails ?? item?.email ?? item?.Email);
      const act = trimText(item?.act ?? item?.Act);
      const description = trimText(item?.description ?? item?.Description);
      const role = trimText(item?.role ?? item?.Role);
      const key = matchKey(state, site, formName);
      const current = byKey.get(key);
      const rowPayload = {
        State: state,
        Site: site,
        FormName: formName,
        Email: email,
        Act: act,
        Description: description,
        Role: role,
      };
      let stored;
      if (current?.id) {
        stored = await table.updateRow({
          ROWID: current.id,
          ...rowPayload,
        });
      } else {
        stored = await table.insertRow(rowPayload);
      }
      const appRow = toAppRow(stored);
      saved.push({
        id: appRow.id || current?.id || null,
        state,
        site,
        formName,
        email,
        act: appRow.act || act,
        description: appRow.description || description,
        role: appRow.role || role,
      });
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
