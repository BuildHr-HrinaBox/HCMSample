import { checklistStateMatchesSiteState, normalizeEmail } from './siteInchargeScope';

const SETUP_API = '/server/setup_function/setup';

function trimText(value) {
  return String(value ?? '').trim();
}

function squashFormName(value) {
  return trimText(value)
    .toLowerCase()
    .replace(/[_-]+/g, ' ')
    .replace(/\s+/g, ' ');
}

export function parseSetupEmails(value) {
  if (Array.isArray(value)) {
    return value.map((item) => normalizeEmail(item)).filter(Boolean);
  }
  return String(value ?? '')
    .split(/[,;]+/)
    .map((part) => normalizeEmail(part))
    .filter(Boolean);
}

export async function fetchSetupFormAssignments() {
  try {
    const res = await fetch(SETUP_API, { credentials: 'include', cache: 'no-store' });
    const data = await res.json().catch(() => null);
    if (data?.status === 'success' && Array.isArray(data.data)) return data.data;
  } catch (_) {
    /* keep Statutory usable if Setup API is down */
  }
  return [];
}

export function getSetupAssignmentsForEmail(setupRows, loginEmail) {
  const login = normalizeEmail(loginEmail);
  if (!login || !Array.isArray(setupRows) || setupRows.length === 0) return [];
  return setupRows.filter((row) => parseSetupEmails(row.email ?? row.Email).includes(login));
}

function formNamesMatch(formName, setupFormName) {
  const a = squashFormName(formName);
  const b = squashFormName(setupFormName);
  return Boolean(a && b && a === b);
}

function statesMatch(rowState, setupState) {
  const setup = trimText(setupState);
  if (!setup) return true;
  const row = trimText(rowState);
  if (!row) return true;
  return (
    checklistStateMatchesSiteState(row, setup) ||
    row.toLowerCase() === setup.toLowerCase()
  );
}

function sitesMatch(rowSite, setupSite, siteHint) {
  const setup = trimText(setupSite);
  if (!setup) return true;
  const candidates = [rowSite, siteHint]
    .map((value) => trimText(value).toLowerCase())
    .filter(Boolean);
  if (candidates.length === 0) return true;
  const want = setup.toLowerCase();
  return candidates.some((site) => site === want || site.includes(want) || want.includes(site));
}

function setupControlledRows(setupFormRows) {
  return (Array.isArray(setupFormRows) ? setupFormRows : []).filter((row) =>
    Boolean(trimText(row.role ?? row.Role))
  );
}

function setupRowsInItemScope(item, rows, { siteHint = '' } = {}) {
  const rowState = item?.state || item?.State || '';
  const rowSite = item?.site || item?.Site || '';
  return rows.filter((row) => {
    if (!statesMatch(rowState, row.state || row.State)) return false;
    if (!sitesMatch(rowSite, row.site || row.Site, siteHint)) return false;
    return true;
  });
}

function findMatchingSetupRows(item, setupFormRows, { siteHint = '' } = {}) {
  const formName = item?.formName || item?.FormName || '';
  const rowState = item?.state || item?.State || '';
  const rowSite = item?.site || item?.Site || '';
  return setupControlledRows(setupFormRows).filter((row) => {
    const setupFormName = row.formName || row.FormName || '';
    if (!formNamesMatch(formName, setupFormName)) return false;
    if (!statesMatch(rowState, row.state || row.State)) return false;
    if (!sitesMatch(rowSite, row.site || row.Site, siteHint)) return false;
    return true;
  });
}

export function statutoryRowAllowedBySetup(item, assignments, { siteHint = '' } = {}) {
  if (!Array.isArray(assignments) || assignments.length === 0) return true;
  const formName = item?.formName || item?.FormName || '';
  const rowState = item?.state || item?.State || '';
  const rowSite = item?.site || item?.Site || '';
  return assignments.some((row) => {
    const setupFormName = row.formName || row.FormName || '';
    if (!formNamesMatch(formName, setupFormName)) return false;
    if (!statesMatch(rowState, row.state || row.State)) return false;
    if (!sitesMatch(rowSite, row.site || row.Site, siteHint)) return false;
    return true;
  });
}

/** True when a statutory row may be shown for this login under Setup role access. */
export function statutoryRowVisibleForLogin(item, setupFormRows, loginEmail, { siteHint = '' } = {}) {
  const controlled = setupControlledRows(setupFormRows);
  const scopeRows = setupRowsInItemScope(item, controlled, { siteHint });
  if (scopeRows.length === 0) return true;

  const login = normalizeEmail(loginEmail);
  if (!login) return false;

  const matching = findMatchingSetupRows(item, setupFormRows, { siteHint });
  if (matching.length === 0) return false;

  return matching.some((row) => parseSetupEmails(row.email ?? row.Email).includes(login));
}
