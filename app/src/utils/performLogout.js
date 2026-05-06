/** Zoho Catalyst hosted login path — same as ProtectedRoute in App.js */
const CATALYST_AUTH_LOGIN = '/__catalyst/auth/login';

const CLEAR_KEYS = ['userRole', 'userEmail', 'userName', 'contractorInfo'];

export function performLogout() {
  try {
    CLEAR_KEYS.forEach((k) => localStorage.removeItem(k));
  } catch (_) {
    /* ignore */
  }
  if (typeof window === 'undefined') return;
  if (window.catalyst?.auth?.signOut) {
    window.catalyst.auth.signOut(CATALYST_AUTH_LOGIN);
  } else {
    window.location.href = CATALYST_AUTH_LOGIN;
  }
}
