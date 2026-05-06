/**
 * Extract email from Catalyst auth response (handles various shapes).
 * Keep in sync with getEmailFromAuth in App.js.
 */
export function getEmailFromAuth(auth) {
  if (!auth || typeof auth === 'boolean') return null;
  const candidates = [
    auth.content?.email_id,
    auth.content?.email,
    auth.user?.email_id,
    auth.user?.email,
    auth.email_id,
    auth.email,
    auth.identity?.email,
    auth.accounts?.[0]?.email,
    auth.user_id,
  ];
  for (const v of candidates) {
    const s = (v && typeof v === 'string' && v.trim()) || null;
    if (s && s.includes('@')) return s;
  }
  return null;
}

export function stringifyUserEmail(userEmail) {
  if (userEmail == null) return '';
  if (typeof userEmail === 'string') return userEmail.trim();
  if (typeof userEmail === 'object') {
    const s = (userEmail.email_id || userEmail.email || '').trim();
    return s;
  }
  return '';
}

/**
 * Best-effort login email for matching Site Incharge email: prop → Catalyst → localStorage.
 */
export async function resolveLoginEmailString(userEmailProp) {
  const fromProp = stringifyUserEmail(userEmailProp);
  if (fromProp) return fromProp;
  try {
    if (typeof window !== 'undefined' && window.catalyst?.auth?.isUserAuthenticated) {
      const auth = await window.catalyst.auth.isUserAuthenticated();
      const fromCat = getEmailFromAuth(auth);
      if (fromCat) return fromCat;
    }
  } catch (_) {
    /* ignore */
  }
  try {
    return (localStorage.getItem('userEmail') || '').trim();
  } catch (_) {
    return '';
  }
}
