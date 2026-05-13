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
  return String(s.siteName ?? s.SiteName ?? '').trim();
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
    const mine = details.filter((s) => normalizeEmail(siteInchargeEmail(s)) === loginNorm);
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

/** Derive Statutory-style act category from act + sector text. */
export function getActCategoryFromActSector(act, sector) {
  const a = String(act || '')
    .trim()
    .toLowerCase();
  const s = String(sector || '')
    .trim()
    .toLowerCase();
  if (
    a.includes('factories act') ||
    a.includes('factory act') ||
    a.includes('factories') ||
    a.includes('factory') ||
    s === 'factories act' ||
    s === 'factory act' ||
    s.includes('factories act') ||
    s.includes('factory act') ||
    s.includes('factories') ||
    s.includes('factory')
  ) {
    return 'factories';
  }
  if (
    a.includes('shops and establishments') ||
    a.includes('shops and establishment') ||
    a.includes('shop and establishment') ||
    a.includes('the shops and establishments act') ||
    a.includes('the shops and establishment act') ||
    s === 'shops and establishment' ||
    s === 'shops and establishments' ||
    s === 'shops and establishment act' ||
    s === 'shop and establishment' ||
    s.includes('shops and establishment') ||
    s.includes('shop and establishment')
  ) {
    return 'shops_and_establishment';
  }
  if (
    a.includes('clra') ||
    a.includes('contract labour') ||
    a.includes('contract labor') ||
    s === 'clra' ||
    s.includes('clra') ||
    s.includes('contract labour') ||
    s.includes('contract labor')
  ) {
    return 'clra';
  }
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
