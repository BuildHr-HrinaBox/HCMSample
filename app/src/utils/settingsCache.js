const API_BASE = '/server/settings_function/settings';
export const SETTINGS_CACHE_KEY = 'settingsData_v1';

let memoryCache = null;
let inflightRequest = null;

function normalizeSettingsPayload(data) {
  return {
    companyName: String(data?.companyName || ''),
    logoName: String(data?.logoName || ''),
    dueDate: String(data?.dueDate || '')
  };
}

export function readSettingsCache() {
  if (memoryCache) return memoryCache;
  try {
    const cached = localStorage.getItem(SETTINGS_CACHE_KEY);
    if (!cached) return null;
    const parsed = JSON.parse(cached);
    if (parsed && typeof parsed === 'object') {
      memoryCache = normalizeSettingsPayload(parsed);
      return memoryCache;
    }
  } catch (_) {
    /* ignore invalid cache */
  }
  return null;
}

export function writeSettingsCache(data) {
  memoryCache = normalizeSettingsPayload(data);
  try {
    localStorage.setItem(SETTINGS_CACHE_KEY, JSON.stringify(memoryCache));
  } catch (_) {
    /* ignore quota errors */
  }
}

export function clearSettingsCache() {
  memoryCache = null;
  try {
    localStorage.removeItem(SETTINGS_CACHE_KEY);
  } catch (_) {
    /* ignore */
  }
}

/** Shared fetch — reuses an in-flight request started by prefetch. */
export function fetchSettings(options = {}) {
  const { force = false } = options;
  if (!force && inflightRequest) return inflightRequest;

  inflightRequest = fetch(API_BASE, { cache: 'no-store' })
    .then(async (res) => {
      const data = await res.json();
      if (!res.ok || data?.status !== 'success') {
        throw new Error(data?.message || 'Failed to load settings');
      }
      writeSettingsCache(data?.data);
      return readSettingsCache();
    })
    .finally(() => {
      inflightRequest = null;
    });

  return inflightRequest;
}

export function prefetchSettings() {
  if (readSettingsCache()) return Promise.resolve(readSettingsCache());
  return fetchSettings().catch(() => null);
}
