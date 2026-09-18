import { resolveLoginEmailString } from './resolveLoginEmail';

const SITE_API = '/server/sitemanagement_function/sitemanagement';

export function normalizeEmail(value) {
  return String(value || '')
    .trim()
    .toLowerCase();
}

export function siteInchargeEmail(s) {
  if (!s || typeof s !== 'object') return '';
  return String(s.inchargeEmail ?? s.InchargeEmail ?? s.incharge_email ?? '').trim();
}

/**
 * All Incharge Mail Ids on a site (supports comma / semicolon / newline / space-separated lists).
 * Example: "2217002@nec.edu.in, ashwath021104@gmail.com" → both addresses.
 */
export function siteInchargeEmails(s) {
  const raw = siteInchargeEmail(s);
  if (!raw) return [];
  const fromRegex = raw.match(/[^\s,;<>]+@[^\s,;<>]+/g);
  if (fromRegex && fromRegex.length > 0) {
    return [...new Set(fromRegex.map((email) => normalizeEmail(email)).filter(Boolean))];
  }
  return raw
    .split(/[,;\n\r|]+/)
    .map((email) => normalizeEmail(email))
    .filter(Boolean);
}

/**
 * Mail Ids shown in Site Management for the current login.
 * When the login email is on the site, show ONLY that address (hide co-incharge emails).
 * Org admins viewing sites where they are not listed still see the full Mail Id list.
 */
export function siteInchargeEmailsForLoginDisplay(site, loginEmail, userRole) {
  const all = siteInchargeEmails(site);
  if (!all.length) return [];
  const loginNorm = normalizeEmail(loginEmail);
  if (loginNorm && all.includes(loginNorm)) return [loginNorm];
  if (isOrgWideSiteViewer(userRole)) return all;
  return all;
}

export function siteInchargeEmailForLoginDisplay(site, loginEmail, userRole) {
  return siteInchargeEmailsForLoginDisplay(site, loginEmail, userRole).join(', ');
}

export function siteIndustry(s) {
  if (!s || typeof s !== 'object') return '';
  return String(s.industry ?? s.Industry ?? '').trim();
}

export function siteStateFromRecord(s) {
  if (!s || typeof s !== 'object') return '';
  return String(s.siteState ?? s.SiteState ?? s.state ?? s.State ?? '').trim();
}

export function siteNameFromSiteRecord(s) {
  if (!s || typeof s !== 'object') return '';
  return (
    String(s.siteName ?? s.SiteName ?? '').trim() ||
    String(s.location ?? s.Location ?? '').trim()
  );
}

/** Location field from Site Management (RJ-Fatehgarh-2, GJ-Maliya, …). */
export function siteLocationFromSiteRecord(s) {
  if (!s || typeof s !== 'object') return '';
  return String(s.location ?? s.Location ?? '').trim();
}

/**
 * Distinct Site Management identities for a login Mail Id.
 * Uses Site Name, falls back to Location, and disambiguates duplicate names
 * across states (two rows both named "Fatehgarh Site" → use Location).
 */
export function buildLoginInchargeSiteRecords(siteDetails, loginEmail) {
  const loginNorm = normalizeEmail(loginEmail);
  const list = Array.isArray(siteDetails) ? siteDetails : [];
  const mine = loginNorm
    ? list.filter((s) => siteInchargeEmails(s).includes(loginNorm))
    : [];
  const drafts = mine
    .map((s) => {
      const rawName = String(s.siteName ?? s.SiteName ?? '').trim();
      const location = siteLocationFromSiteRecord(s);
      const siteState = siteStateFromRecord(s);
      const industry = siteIndustry(s);
      const siteName = rawName || location;
      if (!siteName) return null;
      return {
        siteName,
        rawName,
        location,
        siteState,
        industry,
        actCategory: industryLabelToActCategory(industry)
      };
    })
    .filter(Boolean);

  const nameCounts = new Map();
  drafts.forEach((r) => {
    const k = String(r.siteName || '')
      .trim()
      .toLowerCase();
    if (!k) return;
    nameCounts.set(k, (nameCounts.get(k) || 0) + 1);
  });

  const byKey = new Map();
  drafts.forEach((r) => {
    let display = r.siteName;
    const nameKey = String(r.siteName || '')
      .trim()
      .toLowerCase();
    if ((nameCounts.get(nameKey) || 0) > 1 && r.location) {
      display = r.location;
    }
    const uniq = [
      String(r.siteState || '')
        .trim()
        .toLowerCase(),
      String(display || '')
        .trim()
        .toLowerCase(),
      String(r.location || '')
        .trim()
        .toLowerCase()
    ].join('|');
    if (!byKey.has(uniq)) {
      byKey.set(uniq, {
        siteName: display,
        siteState: r.siteState,
        industry: r.industry,
        actCategory: r.actCategory,
        location: r.location
      });
    }
  });
  return [...byKey.values()];
}

function normScopeToken(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ');
}

/** Split actsbulk / statutory `states` field into comparable tokens. */
export function splitStateFieldTokens(statesOrState) {
  return String(statesOrState || '')
    .split(/[,;/|]/)
    .map((t) => normScopeToken(t))
    .filter((t) => t.length > 0);
}

/** Collapse punctuation/spacing so "Tamil Nadu", "TamilNadu", "tamil-nadu" compare equal. */
export function normalizeStateCompareKey(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '');
}

/**
 * True when a Checklist / Statutory row `state` value refers to the same state as Site Management `SiteState`.
 */
export function checklistStateMatchesSiteState(rowStatesField, siteStateFromSiteMgmt) {
  const target = normalizeStateCompareKey(siteStateFromSiteMgmt);
  if (!target) return false;
  const raw = String(rowStatesField ?? '').trim();
  if (!raw) return false;
  const tokens = splitStateFieldTokens(raw).map(normalizeStateCompareKey).filter(Boolean);
  const keys = tokens.length > 0 ? tokens : [normalizeStateCompareKey(raw)].filter(Boolean);
  return keys.some((k) => k && (k === target || k.includes(target) || target.includes(k)));
}

/**
 * When Site Management assigns state(s) for this Incharge login, only rows whose `state`/`states`
 * explicitly include one of those states are kept. Blank state, "All India", "National", etc. are
 * excluded so generic Shops & Establishment rows do not appear for a Maharashtra-only site.
 */
export function statesFieldMatchesInchargeSiteStates(statesOrState, siteStateLabels) {
  if (!siteStateLabels || siteStateLabels.length === 0) return true;
  const raw = String(statesOrState ?? '').trim();
  if (!raw) return false;
  const blob = normScopeToken(raw);
  if (/\b(all india|pan-india|pan india|national|central|all states|all state)\b/.test(blob)) {
    return false;
  }
  const tokens = splitStateFieldTokens(raw);
  if (tokens.length === 0) return false;
  const allowed = siteStateLabels.map((x) => normScopeToken(x)).filter(Boolean);
  return tokens.some((tok) =>
    allowed.some((a) => a === tok || tok.includes(a) || a.includes(tok))
  );
}

/** Checklist / Actsbulk sector vs Site Management Industry labels for Incharge sites. */
export function sectorMatchesInchargeSiteIndustries(sector, industryLabels) {
  if (!industryLabels || industryLabels.length === 0) return true;
  const sec = normScopeToken(sector);
  if (!sec) return true;
  return industryLabels.some((ind) => {
    const i = normScopeToken(ind);
    if (!i) return false;
    return sec === i || sec.includes(i) || i.includes(sec);
  });
}

/**
 * From Site Management: rows where Incharge email matches resolved login.
 * @returns {{ actCategories: string[] | null, siteNames: string[] | null, industryLabels: string[] | null, stateLabels: string[] | null }}
 */
export async function fetchInchargeDisplayScopeFromSites(userEmailProp) {
  const login = await resolveLoginEmailString(userEmailProp);
  const loginNorm = normalizeEmail(login);
  const empty = { actCategories: null, siteNames: null, industryLabels: null, stateLabels: null };
  if (!loginNorm) return empty;
  try {
    const res = await fetch(SITE_API, { cache: 'no-store' });
    if (!res.ok) return empty;
    const json = await res.json();
    const details = json?.data?.siteDetails;
    if (!Array.isArray(details)) return empty;
    const mine = details.filter((s) => siteInchargeEmails(s).includes(loginNorm));
    if (mine.length === 0) return empty;
    const cats = new Set();
    const siteNames = new Set();
    const industryLabels = new Set();
    const stateLabels = new Set();
    mine.forEach((s) => {
      const sn = siteNameFromSiteRecord(s);
      if (sn) siteNames.add(sn);
      const ind = siteIndustry(s);
      if (ind) industryLabels.add(ind);
      const st = siteStateFromRecord(s);
      if (st) stateLabels.add(st);
      const c = industryLabelToActCategory(ind);
      if (c) cats.add(c);
    });
    return {
      actCategories: cats.size ? [...cats] : null,
      siteNames: siteNames.size ? [...siteNames] : null,
      industryLabels: industryLabels.size ? [...industryLabels] : null,
      stateLabels: stateLabels.size ? [...stateLabels] : null
    };
  } catch {
    return empty;
  }
}

/**
 * Map Site Management "Industry" label to Statutory `getActCategory` bucket.
 */
export function industryLabelToActCategory(industry) {
  const s = String(industry || '').trim().toLowerCase();
  if (!s) return null;
  if (s.includes('factories') || s.includes('factory act') || s === 'factory') return 'factories';
  if (s.includes('shops') && s.includes('establishment')) return 'shops_and_establishment';
  if (s.includes('clra') || s.includes('contract labour') || s.includes('contract labor')) return 'clra';
  return null;
}

/** Derive Statutory-style act category from act + sector text (act wins when it clearly implies a bucket). */
export function getActCategoryFromActSector(act, sector) {
  const a = String(act || '')
    .trim()
    .toLowerCase();
  const s = String(sector || '')
    .trim()
    .toLowerCase();

  const factoriesFrom = (blob) =>
    blob &&
    (blob.includes('factories act') ||
      blob.includes('factory act') ||
      (blob.includes('factories') && !blob.includes('contract')) ||
      blob.includes('factory'));
  const shopsFrom = (blob) =>
    blob &&
    (blob.includes('shops and establishments') ||
      blob.includes('shops and establishment') ||
      blob.includes('shop and establishment') ||
      blob.includes('the shops and establishments act') ||
      blob.includes('the shops and establishment act'));
  const clraFrom = (blob) =>
    blob &&
    (blob.includes('clra') ||
      blob.includes('contract labour') ||
      blob.includes('contract labor'));

  if (clraFrom(a)) return 'clra';
  if (factoriesFrom(a)) return 'factories';
  if (shopsFrom(a)) return 'shops_and_establishment';

  if (clraFrom(s)) return 'clra';
  if (factoriesFrom(s)) return 'factories';
  if (shopsFrom(s)) return 'shops_and_establishment';

  return 'other';
}

/**
 * Act categories (factories | shops_and_establishment | clra) from sites where Incharge email matches login.
 * @returns {string[] | null} null = no scope (show all statutory)
 */
export async function fetchAllowedActCategoriesFromSites(userEmailProp) {
  const { actCategories } = await fetchInchargeDisplayScopeFromSites(userEmailProp);
  return actCategories;
}

/** Org-wide viewers (not restricted to Incharge sites). */
export function isOrgWideSiteViewer(role) {
  const r = String(role || '').trim();
  return r === 'App Administrator' || r === 'HR Admin';
}

/**
 * Scope labels from sites where the login email is Incharge.
 * @returns {{ stateLabels: string[], industryLabels: string[], mine: object[] } | null}
 */
export function buildInchargeSiteScopeFromList(sites, loginEmail) {
  const loginNorm = normalizeEmail(loginEmail);
  if (!loginNorm) return null;
  const list = Array.isArray(sites) ? sites : [];
  const mine = list.filter((s) => siteInchargeEmails(s).includes(loginNorm));
  if (!mine.length) return null;
  const stateLabels = [...new Set(mine.map(siteStateFromRecord).filter(Boolean))];
  const industryLabels = [...new Set(mine.map(siteIndustry).filter(Boolean))];
  return { stateLabels, industryLabels, mine };
}

/**
 * Site Management table rows visible to the current login.
 * Any site whose Mail Id list includes the login email is shown (multi-email OK).
 * Example: Mail Id "a@x.com, b@y.com" → login b@y.com still sees that site.
 * App User with no matching Mail Id sees none; App Administrator / HR Admin with no
 * personal assignment still see all sites (org-wide management).
 */
export function filterSitesForLoginUser(sites, loginEmail, userRole) {
  const list = Array.isArray(sites) ? sites : [];
  const loginNorm = normalizeEmail(loginEmail);
  if (!loginNorm) {
    return isOrgWideSiteViewer(userRole) ? list : [];
  }

  const mine = list.filter((s) => siteInchargeEmails(s).includes(loginNorm));
  if (mine.length > 0) return mine;

  if (isOrgWideSiteViewer(userRole)) return list;
  return [];
}

export function hasInchargeSiteScope(sites, loginEmail) {
  return buildInchargeSiteScopeFromList(sites, loginEmail) != null;
}
