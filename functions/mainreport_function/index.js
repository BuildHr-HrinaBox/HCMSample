'use strict';

const express = require('express');
const catalystSDK = require('zcatalyst-sdk-node');

const app = express();
app.use(express.json());

/** Same Draft folder as statutoryreg_function (DraftFileName) */
const DRAFT_FOLDER_ID = '26741000000061720';

const MONTH_ORDER = [
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
  'December'
];

const SHORT_MONTH = {
  jan: 0,
  feb: 1,
  mar: 2,
  apr: 3,
  may: 4,
  jun: 5,
  jul: 6,
  aug: 7,
  sep: 8,
  oct: 9,
  nov: 10,
  dec: 11
};

app.use((req, res, next) => {
  try {
    const catalyst = catalystSDK.initialize(req);
    res.locals.catalyst = catalyst;
    next();
  } catch (err) {
    res.status(500).json({ status: 'failure', message: 'Catalyst init failed' });
  }
});

app.get('/', (req, res) => {
  res.status(200).json({ status: 'success', message: 'mainreport_function ready' });
});

function normalizeMonth(m) {
  const s = String(m || '')
    .trim()
    .toLowerCase();
  if (!s) return '';
  const full = MONTH_ORDER.find((x) => x.toLowerCase() === s);
  if (full) return full;
  if (SHORT_MONTH[s] !== undefined) return MONTH_ORDER[SHORT_MONTH[s]];
  if (/^\d+$/.test(s)) {
    const n = parseInt(s, 10);
    if (n >= 1 && n <= 12) return MONTH_ORDER[n - 1];
  }
  return String(m).trim();
}

function monthsMatch(a, b) {
  return normalizeMonth(a) === normalizeMonth(b);
}

function yearMatch(yr, yParam) {
  const a = parseInt(String(yr || '').trim(), 10);
  const b = parseInt(String(yParam || '').trim(), 10);
  if (isNaN(a) || isNaN(b)) return String(yr || '').trim() === String(yParam || '').trim();
  return a === b;
}

function sortMonths(months) {
  const uniq = [...new Set(months.map((m) => normalizeMonth(m) || m))].filter(Boolean);
  return uniq.sort((a, b) => {
    const ia = MONTH_ORDER.indexOf(normalizeMonth(a) || a);
    const ib = MONTH_ORDER.indexOf(normalizeMonth(b) || b);
    const va = ia === -1 ? 100 : ia;
    const vb = ib === -1 ? 100 : ib;
    return va - vb;
  });
}

function normalizeEmail(value) {
  return String(value || '')
    .trim()
    .toLowerCase();
}

function getActCategoryFromStatutoryRow(row) {
  if (!row) return null;
  const normalizeString = (str) => {
    if (!str) return '';
    return String(str)
      .trim()
      .toLowerCase();
  };
  const act = normalizeString(row.Act || row.act || '');
  const sector = normalizeString(row.Sector || row.sector || '');
  const isFactoriesAct =
    act.includes('factories act') ||
    act.includes('factory act') ||
    act.includes('factories') ||
    act.includes('factory') ||
    sector === 'factories act' ||
    sector === 'factory act' ||
    sector.includes('factories act') ||
    sector.includes('factory act') ||
    sector.includes('factories') ||
    sector.includes('factory');
  if (isFactoriesAct) return 'factories';
  const isShopsAndEstablishment =
    act.includes('shops and establishments') ||
    act.includes('shops and establishment') ||
    act.includes('shop and establishment') ||
    act.includes('the shops and establishments act') ||
    act.includes('the shops and establishment act') ||
    sector === 'shops and establishment' ||
    sector === 'shops and establishments' ||
    sector === 'shops and establishment act' ||
    sector === 'shop and establishment' ||
    sector.includes('shops and establishment') ||
    sector.includes('shop and establishment');
  if (isShopsAndEstablishment) return 'shops_and_establishment';
  const isCLRA =
    act.includes('clra') ||
    act.includes('contract labour') ||
    act.includes('contract labor') ||
    sector === 'clra' ||
    sector.includes('clra') ||
    sector.includes('contract labour') ||
    sector.includes('contract labor');
  if (isCLRA) return 'clra';
  return 'other';
}

function getActCategoryFromIndustryLabel(industry) {
  return getActCategoryFromStatutoryRow({ Sector: industry || '' });
}

/** Same site-name → act bucket as Statutory.js for Delphi URLs when Industry does not map. */
function getActCategoryFromSiteNameFallback(siteName) {
  const s = String(siteName || '')
    .trim()
    .toLowerCase();
  if (s === 'delphi kakinada') return 'clra';
  if (s === 'delphi') return 'shops_and_establishment';
  if (s === 'delphi oragadam') return 'factories';
  return null;
}

/**
 * Act category for a Site row: Industry first, then known site-name aliases.
 */
function getActCategoryForSiteRow(siteRow, siteNameParam) {
  const industry = siteRow ? String(siteRow.Industry || siteRow.industry || '').trim() : '';
  let cat = getActCategoryFromIndustryLabel(industry);
  if (cat && cat !== 'other') return cat;
  const fallback = getActCategoryFromSiteNameFallback(siteNameParam);
  return fallback || cat || 'other';
}

function normStateToken(v) {
  return String(v || '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ');
}

/** Same rules as app `statesFieldMatchesInchargeSiteStates` (strict state when Incharge sites have SiteState). */
function rowStatesMatchInchargeScope(rowStateField, stateLabels) {
  if (!stateLabels || stateLabels.length === 0) return true;
  const raw = String(rowStateField ?? '').trim();
  if (!raw) return false;
  const blob = normStateToken(raw);
  if (/\b(all india|pan india|pan-india|national|central|all states|all state)\b/.test(blob)) {
    return false;
  }
  const tokens = String(raw)
    .split(/[,;/|]/)
    .map((t) => normStateToken(t))
    .filter((t) => t.length > 0);
  if (tokens.length === 0) return false;
  const allowed = stateLabels.map(normStateToken).filter(Boolean);
  return tokens.some((tok) => allowed.some((a) => a === tok || tok.includes(a) || a.includes(tok)));
}

function filterRowsForInchargeLocation(rows, stateLabels, siteNamesLower) {
  if ((!stateLabels || stateLabels.length === 0) && (!siteNamesLower || siteNamesLower.size === 0)) {
    return rows;
  }
  if (!Array.isArray(rows)) return rows;
  return rows.filter((r) => {
    const site = String(r.Site || r.site || '').trim().toLowerCase();
    if (siteNamesLower && siteNamesLower.size > 0 && site && !siteNamesLower.has(site)) return false;
    return rowStatesMatchInchargeScope(r.State || r.state || '', stateLabels);
  });
}

/**
 * Site rows where Incharge email matches user: act categories, state labels, site names (lowercase).
 */
async function getInchargeReportScope(catalyst, userEmail) {
  const email = normalizeEmail(userEmail);
  const empty = { actCategories: null, stateLabels: null, siteNamesLower: null };
  if (!email) return empty;
  try {
    const siteRows = await catalyst.datastore().table('Site').getAllRows();
    const mine = siteRows.filter(
      (row) => normalizeEmail(row.InchargeEmail || row.inchargeEmail || row.incharge_email) === email
    );
    if (mine.length === 0) return empty;
    const cats = new Set();
    const states = new Set();
    const sites = new Set();
    mine.forEach((row) => {
      const sn = String(row.SiteName || row.siteName || '').trim();
      if (sn) sites.add(sn.toLowerCase());
      const st = String(row.SiteState || row.siteState || row.State || row.state || '').trim();
      if (st) states.add(st);
      const cat = getActCategoryFromIndustryLabel(row.Industry || row.industry || '');
      if (cat && cat !== 'other') cats.add(cat);
    });
    return {
      actCategories: cats.size ? [...cats] : null,
      stateLabels: states.size ? [...states] : null,
      siteNamesLower: sites.size ? sites : null
    };
  } catch (err) {
    console.warn('getInchargeReportScope:', err?.message || err);
    return empty;
  }
}

async function findSiteRowByName(catalyst, siteName) {
  const want = String(siteName || '')
    .trim()
    .toLowerCase();
  if (!want) return null;
  try {
    const rows = await catalyst.datastore().table('Site').getAllRows();
    for (const r of rows) {
      const n = String(r.SiteName || r.siteName || '')
        .trim()
        .toLowerCase();
      if (n === want) return r;
    }
  } catch (err) {
    console.warn('findSiteRowByName:', err?.message || err);
  }
  return null;
}

function filterRowsByActCategoryScope(rows, allowedActCategories) {
  if (!Array.isArray(allowedActCategories) || allowedActCategories.length === 0) return rows;
  const allow = new Set(allowedActCategories);
  return rows.filter((r) => allow.has(getActCategoryFromStatutoryRow(r)));
}

/** Month filter as stored on Statutory / Catalyst rows (field names vary by table mapping). */
function getStatutoryStoredMonthRaw(stRow) {
  if (!stRow) return '';
  const v =
    stRow.MonthFilter ??
    stRow.monthfilter ??
    stRow.Monthfilter ??
    stRow.monthFilter ??
    stRow.MONTHFILTER ??
    stRow.OriginalMonthFilter ??
    stRow.originalMonthFilter ??
    '';
  return String(v ?? '').trim();
}

/** Same field as Statutory "Month Filter" column (`monthFilter` / `MonthFilter`). */
function getStatutoryMonthFilterDisplay(stRow, reportMonthFallback) {
  const raw = getStatutoryStoredMonthRaw(stRow);
  if (raw !== '') {
    const n = normalizeMonth(raw);
    return n || raw;
  }
  const fb = normalizeMonth(reportMonthFallback);
  return fb || String(reportMonthFallback || '').trim();
}

function statutoryRowMatchesMonth(stRow, selectedMonth) {
  const selectedNorm = normalizeMonth(selectedMonth);
  if (!selectedNorm) return false;
  const selectedPrefix = selectedNorm.toLowerCase().substring(0, 3);
  const storedNorm = getStatutoryStoredMonthRaw(stRow).toLowerCase().substring(0, 3);
  if (storedNorm) return storedNorm === selectedPrefix;
  const dueDate = String(stRow.DueDate || stRow.dueDate || '')
    .toLowerCase()
    .trim();
  if (!dueDate) return false;
  const monthNames = MONTH_ORDER.map((m) => m.toLowerCase());
  const monthAbbr = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];
  const idx = MONTH_ORDER.indexOf(selectedNorm);
  if (idx === -1) return false;
  const fullMonthName = monthNames[idx];
  const monthAbbrName = monthAbbr[idx];
  /** Never treat "Monthly Basis" alone as matching every month — that showed May-only drafts under January Reports. */
  return dueDate.includes(fullMonthName) || dueDate.includes(monthAbbrName);
}

function getStatutoryRowStatusRaw(stRow) {
  if (!stRow) return '';
  const direct =
    stRow.Status ??
    stRow.status ??
    stRow.STATUS ??
    stRow.TransactionStatus ??
    stRow.transactionStatus;
  if (direct != null && String(direct).trim() !== '') return String(direct).trim();
  for (const k of Object.keys(stRow)) {
    if (String(k).toLowerCase() === 'status') {
      const v = stRow[k];
      if (v != null && String(v).trim() !== '') return String(v).trim();
    }
  }
  return '';
}

/** Match Statutory.js STATUS column: only transaction `Status` is Approved (not Pending / Yet to Complete). */
function isStatutoryTransactionApproved(stRow) {
  const st = getStatutoryRowStatusRaw(stRow).toLowerCase();
  return st === 'approved' || st === 'approve';
}

function pickBestStatutoryRow(matches) {
  if (!matches || matches.length === 0) return null;
  if (matches.length === 1) return matches[0];
  const score = (r) => {
    let s = 0;
    const draft = r.DraftFile;
    const hasDraft =
      draft != null && String(draft).trim() !== '' && String(draft).toLowerCase() !== 'null';
    if (hasDraft) s += 100;
    if (r.ROWID != null) s += 50;
    return s;
  };
  return [...matches].sort((a, b) => score(b) - score(a))[0];
}

function findStatutoryCandidatesForMainRow(mainRow, filteredStatutoryRows) {
  const raw = String(mainRow.FormName != null ? mainRow.FormName : '').trim();
  if (!raw) return [];
  if (/^\d+$/.test(raw)) {
    const byId = filteredStatutoryRows.filter((x) => String(x.ROWID) === raw);
    if (byId.length) return byId;
  }
  const lower = raw.toLowerCase();
  return filteredStatutoryRows.filter((x) => String(x.FormName || '').trim().toLowerCase() === lower);
}

function buildStatutoryDraftRow(stRow) {
  const id = stRow.ROWID;
  const formName = String(stRow.FormName || '').trim();
  const act = getActFromAnyRow(stRow);
  const description = getDescriptionFromAnyRow(stRow);
  const approved = isStatutoryTransactionApproved(stRow);
  const hasDraftStored =
    stRow.DraftFile != null && String(stRow.DraftFile).trim() !== '' && String(stRow.DraftFile).toLowerCase() !== 'null';
  const draftFile = hasDraftStored ? String(stRow.DraftFile).trim() : null;
  const draftFileName =
    stRow.DraftFileName != null && String(stRow.DraftFileName).trim() !== ''
      ? String(stRow.DraftFileName).trim()
      : 'draft';
  /** Reports: draft download only after transaction approval (same gate as statutory workflow). */
  const draftFileUrl =
    draftFile && approved
      ? `/server/statutoryreg_function/statutory/${id}/file/Draft?name=${encodeURIComponent(draftFileName)}`
      : null;
  const approval = String(stRow.Approval || stRow.approval || '').trim();
  const statutoryStatus = getStatutoryRowStatusRaw(stRow);
  const monthFilter = getStatutoryMonthFilterDisplay(stRow, '');
  return {
    rowId: id,
    statutoryRowId: id,
    formName,
    act,
    description,
    monthFilter,
    draftFile: approved ? draftFile : null,
    draftFileName: approved ? draftFileName : '',
    draftFileUrl,
    approval,
    statutoryStatus,
    hasStatutoryDraftStored: hasDraftStored
  };
}

function normalizeFormKey(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]/g, '');
}

function getFormNameFromAnyRow(row) {
  return String(row?.FormName || row?.formName || '').trim();
}

function getActFromAnyRow(row) {
  return String(row?.Act || row?.act || '').trim();
}

function getDescriptionFromAnyRow(row) {
  return String(row?.Description || row?.description || '').trim();
}

function getMasterKeyFromAnyRow(row) {
  const formKey = normalizeFormKey(getFormNameFromAnyRow(row));
  const actKey = normalizeFormKey(getActFromAnyRow(row));
  const descKey = normalizeFormKey(getDescriptionFromAnyRow(row));
  return `${formKey}|${actKey}|${descKey}`;
}

function buildEmptyFormRow(formKey, formName, act, description, reportMonthNorm) {
  const mf = normalizeMonth(reportMonthNorm) || String(reportMonthNorm || '').trim();
  return {
    rowId: `master_${formKey}`,
    statutoryRowId: null,
    formName: formName || '',
    act: act || '',
    description: description || '',
    monthFilter: mf,
    draftFile: null,
    draftFileName: '',
    draftFileUrl: null,
    approval: '',
    statutoryStatus: '',
    hasStatutoryDraftStored: false
  };
}

async function getAllSiteNames(catalyst, allowedActCategories, siteParam, inchargeSiteNamesLower) {
  try {
    const rows = await catalyst.datastore().table('Site').getAllRows();
    let scopedRows = rows;
    if (Array.isArray(allowedActCategories) && allowedActCategories.length > 0) {
      const allow = new Set(allowedActCategories);
      scopedRows = rows.filter((r) => allow.has(getActCategoryFromIndustryLabel(r.Industry || r.industry || '')));
    }
    if (inchargeSiteNamesLower && inchargeSiteNamesLower.size > 0) {
      scopedRows = scopedRows.filter((r) =>
        inchargeSiteNamesLower.has(String(r.SiteName || r.siteName || '').trim().toLowerCase())
      );
    }
    let names = scopedRows
      .map((r) => String(r.SiteName || r.siteName || '').trim())
      .filter(Boolean);

    if (siteParam) {
      names = names.filter((n) => n.toLowerCase() === String(siteParam).trim().toLowerCase());
    }

    return [...new Set(names)].sort((a, b) => a.localeCompare(b, undefined, { sensitivity: 'base' }));
  } catch (err) {
    console.warn('Could not fetch Site names:', err?.message || err);
    return [];
  }
}

function getContentType(filename) {
  const ext = String(filename || '')
    .toLowerCase()
    .split('.')
    .pop();
  const contentTypes = {
    pdf: 'application/pdf',
    doc: 'application/msword',
    docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    xls: 'application/vnd.ms-excel',
    xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    png: 'image/png',
    jpg: 'image/jpeg',
    jpeg: 'image/jpeg',
    gif: 'image/gif',
    txt: 'text/plain'
  };
  return contentTypes[ext] || 'application/octet-stream';
}

/**
 * Years = distinct years from MainReport plus the current calendar year (server time).
 * When the date rolls to 1 Jan of a new year, that year is included automatically.
 */
app.get('/mainreport/summary', async (req, res) => {
  try {
    res.set('Cache-Control', 'no-store');
    const { catalyst } = res.locals;
    const currentYear = new Date().getFullYear();
    const table = catalyst.datastore().table('MainReport');
    const rows = await table.getAllRows();

    const yearNums = new Set();
    yearNums.add(currentYear);

    const monthsByYear = {};

    for (const r of rows) {
      const y = parseInt(String(r.Year || '').trim(), 10);
      if (!isNaN(y)) {
        yearNums.add(y);
        const yk = String(y);
        if (!monthsByYear[yk]) monthsByYear[yk] = [];
        const m = r.Months != null ? String(r.Months).trim() : '';
        if (m) monthsByYear[yk].push(m);
      }
    }

    for (const k of Object.keys(monthsByYear)) {
      monthsByYear[k] = sortMonths(monthsByYear[k]);
    }

    const years = [...yearNums].sort((a, b) => b - a).map(String);

    res.status(200).json({
      status: 'success',
      data: {
        years,
        monthsByYear,
        serverYear: currentYear
      }
    });
  } catch (err) {
    console.error('mainreport/summary:', err);
    res.status(500).json({ status: 'failure', message: err.message || 'Failed to load main report summary.' });
  }
});

app.get('/mainreport/entries', async (req, res) => {
  try {
    res.set('Cache-Control', 'no-store');
    const year = req.query.year;
    const month = req.query.month;
    const userEmail = req.query.userEmail != null ? String(req.query.userEmail) : '';
    if (year == null || year === '' || month == null || month === '') {
      return res.status(400).json({ status: 'failure', message: 'year and month query parameters are required.' });
    }

    const { catalyst } = res.locals;
    const inchargeScope = await getInchargeReportScope(catalyst, userEmail);
    const allowedActCategories = inchargeScope.actCategories;

    const siteParamRaw = req.query.site != null ? String(req.query.site) : '';
    const siteParam = siteParamRaw.trim();

    /** When a site is chosen, narrow to that site's Industry (act bucket). Incharge scope still applies. */
    let narrowActCategories = allowedActCategories;
    if (siteParam) {
      const siteRow = await findSiteRowByName(catalyst, siteParam);
      if (!siteRow) {
        const siteNames = await getAllSiteNames(catalyst, allowedActCategories, '', inchargeScope.siteNamesLower);
        return res.status(200).json({
          status: 'success',
          data: { rows: [], siteNames }
        });
      }
      const siteCat = getActCategoryForSiteRow(siteRow, siteParam);
      const siteCats = [siteCat];
      if (Array.isArray(allowedActCategories) && allowedActCategories.length > 0) {
        const intersection = siteCats.filter((c) => allowedActCategories.includes(c));
        if (intersection.length === 0) {
          const siteNames = await getAllSiteNames(catalyst, allowedActCategories, '', inchargeScope.siteNamesLower);
          return res.status(200).json({
            status: 'success',
            data: { rows: [], siteNames }
          });
        }
        narrowActCategories = intersection;
      } else {
        narrowActCategories = siteCats;
      }
    }

    // Submitted rows (with draft linkage) come from Statutory.
    const statutoryTable = catalyst.datastore().table('Statutory');
    const statutoryAll = await statutoryTable.getAllRows();
    let statutoryForUser = filterRowsByActCategoryScope(statutoryAll, narrowActCategories);
    statutoryForUser = filterRowsForInchargeLocation(
      statutoryForUser,
      inchargeScope.stateLabels,
      inchargeScope.siteNamesLower
    );

    // Master form list comes from checklistbulk (same source used by Statutory page merge).
    const bulkTable = catalyst.datastore().table('checklistbulk');
    const bulkAll = await bulkTable.getAllRows();
    let bulkForUser = filterRowsByActCategoryScope(bulkAll, narrowActCategories);
    bulkForUser = filterRowsForInchargeLocation(
      bulkForUser,
      inchargeScope.stateLabels,
      inchargeScope.siteNamesLower
    );

    // Keep duplicates by form when act/description differs (same behavior user expects from Statutory list).
    const formsMap = new Map();
    for (const row of bulkForUser) {
      const name = getFormNameFromAnyRow(row);
      const formOnlyKey = normalizeFormKey(name);
      if (!formOnlyKey) continue;
      const key = getMasterKeyFromAnyRow(row);
      if (!formsMap.has(key)) {
        formsMap.set(key, {
          formName: name,
          act: getActFromAnyRow(row),
          description: getDescriptionFromAnyRow(row)
        });
      }
    }
    for (const row of statutoryForUser) {
      const name = getFormNameFromAnyRow(row);
      const formOnlyKey = normalizeFormKey(name);
      if (!formOnlyKey) continue;
      const key = getMasterKeyFromAnyRow(row);
      if (!formsMap.has(key)) {
        formsMap.set(key, {
          formName: name,
          act: getActFromAnyRow(row),
          description: getDescriptionFromAnyRow(row)
        });
      }
    }

    const reportMonthNorm = normalizeMonth(month);
    const out = [];
    for (const [formKey, meta] of formsMap.entries()) {
      const candidates = statutoryForUser.filter((r) => getMasterKeyFromAnyRow(r) === formKey);
      const monthCandidates = candidates.filter((r) => statutoryRowMatchesMonth(r, month));
      /** Do not fall back to other months' statutory rows — e.g. May submission must not show under January Reports. */
      const picked = monthCandidates.length > 0 ? pickBestStatutoryRow(monthCandidates) : null;
      if (picked) {
        const built = buildStatutoryDraftRow(picked);
        if (!built.act && meta?.act) built.act = meta.act;
        if (!built.description && meta?.description) built.description = meta.description;
        built.monthFilter = getStatutoryMonthFilterDisplay(picked, reportMonthNorm) || reportMonthNorm;
        out.push(built);
      } else {
        out.push(buildEmptyFormRow(formKey, meta?.formName, meta?.act, meta?.description, reportMonthNorm));
      }
    }

    out.sort((a, b) => String(a.formName || '').localeCompare(String(b.formName || ''), undefined, { sensitivity: 'base' }));

    // Full site list for the dropdown; do not filter by `site` query (that hid all other sites).
    const siteNames = await getAllSiteNames(catalyst, allowedActCategories, '', inchargeScope.siteNamesLower);

    res.status(200).json({
      status: 'success',
      data: { rows: out, siteNames }
    });
  } catch (err) {
    console.error('mainreport/entries:', err);
    res.status(500).json({ status: 'failure', message: err.message || 'Failed to load entries.' });
  }
});

app.get('/mainreport/file/:fileId', async (req, res) => {
  try {
    const { fileId } = req.params;
    const name = (req.query.name && String(req.query.name)) || 'download';
    const { catalyst } = res.locals;
    const folder = catalyst.filestore().folder(DRAFT_FOLDER_ID);
    const fileBuffer = await folder.downloadFile(fileId);
    const contentType = getContentType(name);
    const disposition = req.query.disposition === 'attachment' ? 'attachment' : 'inline';
    res.writeHead(200, {
      'Content-Type': contentType,
      'Content-Disposition': `${disposition}; filename="${name.replace(/"/g, '')}"`,
      'Content-Length': fileBuffer.length
    });
    res.end(fileBuffer);
  } catch (err) {
    console.error('mainreport/file:', err);
    res.status(404).json({ status: 'failure', message: err.message || 'File not found.' });
  }
});

module.exports = app;
