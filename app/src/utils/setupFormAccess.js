import {
  checklistStateMatchesSiteState,
  getActCategoryFromActSector,
  industryLabelToActCategory,
  normalizeEmail,
} from './siteInchargeScope';

const SETUP_API = '/server/setup_function/setup';

function trimText(value) {
  return String(value ?? '').trim();
}

function squashText(value) {
  return trimText(value)
    .toLowerCase()
    .replace(/[_-]+/g, ' ')
    .replace(/\s+/g, ' ');
}

function squashFormName(value) {
  return squashText(value);
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
  if (!row) return false;
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
  if (candidates.length === 0) return false;
  const want = setup.toLowerCase();
  return candidates.some((site) => site === want || site.includes(want) || want.includes(site));
}

/** Soft match: empty setup value means "any". Otherwise require overlap / equality. */
function textsMatch(itemValue, setupValue) {
  const setup = squashText(setupValue);
  if (!setup) return true;
  const row = squashText(itemValue);
  if (!row) return false;
  return row === setup || row.includes(setup) || setup.includes(row);
}

function actsMatch(item, setupRow) {
  return textsMatch(item?.act ?? item?.Act, setupRow?.act ?? setupRow?.Act);
}

function descriptionsMatch(item, setupRow) {
  return textsMatch(
    item?.description ?? item?.Description,
    setupRow?.description ?? setupRow?.Description
  );
}

function industryTypesMatch(item, setupRow) {
  const setupIndustry = trimText(setupRow?.industryType ?? setupRow?.IndustryType);
  if (!setupIndustry) return true;

  const wantCategory = industryLabelToActCategory(setupIndustry);
  const itemIndustry = trimText(
    item?.industryType ?? item?.IndustryType ?? item?.industry ?? item?.Industry
  );
  if (itemIndustry) {
    if (itemIndustry.toLowerCase() === setupIndustry.toLowerCase()) return true;
    const itemCategory = industryLabelToActCategory(itemIndustry);
    if (wantCategory && itemCategory && itemCategory === wantCategory) return true;
  }

  const act = item?.act ?? item?.Act ?? '';
  const sector = item?.sector ?? item?.Sector ?? '';
  const rowCategory = getActCategoryFromActSector(act, sector);
  if (wantCategory && rowCategory && rowCategory !== 'other' && rowCategory === wantCategory) {
    return true;
  }

  if (!wantCategory) {
    const blob = squashText(`${act} ${sector} ${itemIndustry}`);
    const setup = squashText(setupIndustry);
    return Boolean(blob && setup && (blob.includes(setup) || setup.includes(blob)));
  }

  return false;
}

function setupRowMatchesItem(item, setupRow, { siteHint = '' } = {}) {
  const formName = item?.formName || item?.FormName || '';
  const rowState = item?.state || item?.State || '';
  const rowSite = item?.site || item?.Site || '';
  const setupFormName = setupRow?.formName || setupRow?.FormName || '';

  if (!formNamesMatch(formName, setupFormName)) return false;
  if (!statesMatch(rowState, setupRow?.state || setupRow?.State)) return false;
  if (!sitesMatch(rowSite, setupRow?.site || setupRow?.Site, siteHint)) return false;
  if (!actsMatch(item, setupRow)) return false;
  if (!descriptionsMatch(item, setupRow)) return false;
  if (!industryTypesMatch(item, setupRow)) return false;
  return true;
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
    if (!industryTypesMatch(item, row)) return false;
    return true;
  });
}

function findMatchingSetupRows(item, setupFormRows, { siteHint = '' } = {}) {
  return setupControlledRows(setupFormRows).filter((row) =>
    setupRowMatchesItem(item, row, { siteHint })
  );
}

export function statutoryRowAllowedBySetup(item, assignments, { siteHint = '' } = {}) {
  if (!Array.isArray(assignments) || assignments.length === 0) return true;
  return assignments.some((row) => setupRowMatchesItem(item, row, { siteHint }));
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
