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
  const login = await resolveLoginEmailString(userEmailProp);
  const loginNorm = normalizeEmail(login);
  if (!loginNorm) return null;
  try {
    const res = await fetch(SITE_API, { cache: 'no-store' });
    if (!res.ok) return null;
    const json = await res.json();
    const details = json?.data?.siteDetails;
    if (!Array.isArray(details)) return null;
    const cats = new Set();
    details.forEach((s) => {
      if (normalizeEmail(siteInchargeEmail(s)) !== loginNorm) return;
      const c = industryLabelToActCategory(siteIndustry(s));
      if (c) cats.add(c);
    });
    return cats.size ? [...cats] : null;
  } catch {
    return null;
  }
}
