'use strict';

function normalizeIndustryLabel(value) {
  const raw = String(value || '').trim();
  if (!raw || raw === '—') return 'Others';
  const lower = raw.toLowerCase();
  if (lower.includes('clra')) return 'CLRA';
  if (lower.includes('shop') && lower.includes('establishment')) return 'Shops and Establishment';
  if (lower.includes('factor')) return 'Factories Act';
  return raw;
}

function sortSiteSections(siteSections) {
  return [...(siteSections || [])].sort((a, b) =>
    String(a.label || a.siteName || '').localeCompare(String(b.label || b.siteName || ''), undefined, {
      sensitivity: 'base'
    })
  );
}

function groupPendingSitesBySite(pendingSites) {
  const groups = new Map();
  for (const item of pendingSites || []) {
    const key = String(item.siteKey || item.siteName || item.label || '')
      .trim()
      .toLowerCase();
    if (!key) continue;
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push({
      ...item,
      industry: normalizeIndustryLabel(item.industry)
    });
  }
  return [...groups.values()].sort((a, b) =>
    String(a[0]?.label || a[0]?.siteName || '').localeCompare(
      String(b[0]?.label || b[0]?.siteName || ''),
      undefined,
      { sensitivity: 'base' }
    )
  );
}

function groupSiteFormsByIndustry(forms) {
  const map = new Map();
  for (const form of forms || []) {
    const industry = normalizeIndustryLabel(form.industry);
    if (!map.has(industry)) map.set(industry, []);
    map.get(industry).push({ ...form, industry });
  }
  for (const list of map.values()) {
    list.sort((a, b) =>
      String(a.formName || '').localeCompare(String(b.formName || ''), undefined, { sensitivity: 'base' })
    );
  }
  return [...map.entries()].sort((a, b) => a[0].localeCompare(b[0], undefined, { sensitivity: 'base' }));
}

function getSiteIndustryLabels(forms) {
  return groupSiteFormsByIndustry(forms)
    .map(([industry]) => industry)
    .filter((industry) => industry && industry !== 'Others');
}

function formRowDedupeKey(form) {
  return `${String(form.formNumber || '').trim().toLowerCase()}|${String(form.formName || '')
    .trim()
    .toLowerCase()}|${normalizeIndustryLabel(form.industry).toLowerCase()}|${String(form.status || '')
    .trim()
    .toLowerCase()}`;
}

/** One email block per site name + industry (Site table may have duplicate rows). */
function siteSectionDisplayKey(section) {
  const siteName = String(section.siteName || section.label || '')
    .trim()
    .toLowerCase();
  const fromForms = getSiteIndustryLabels(section.forms || []);
  const industry =
    fromForms.length === 1
      ? fromForms[0]
      : normalizeIndustryLabel(section.siteIndustry);
  return `${siteName}|${String(industry || 'others').toLowerCase()}`;
}

function mergeDuplicateSiteSections(siteSections) {
  const map = new Map();
  for (const section of siteSections || []) {
    const key = siteSectionDisplayKey(section);
    if (!map.has(key)) {
      map.set(key, {
        ...section,
        forms: [...(section.forms || [])]
      });
      continue;
    }
    const existing = map.get(key);
    const seen = new Set((existing.forms || []).map(formRowDedupeKey));
    for (const form of section.forms || []) {
      const fk = formRowDedupeKey(form);
      if (seen.has(fk)) continue;
      seen.add(fk);
      existing.forms.push(form);
    }
  }
  return sortSiteSections([...map.values()]);
}

module.exports = {
  normalizeIndustryLabel,
  groupPendingSitesBySite,
  groupSiteFormsByIndustry,
  getSiteIndustryLabels,
  sortSiteSections,
  mergeDuplicateSiteSections
};
