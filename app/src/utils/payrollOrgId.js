/** Correct Zoho Payroll organisation ID for Vayona Energy. */
export const PAYROLL_ORGANIZATION_ID = '60065031807';

/** Legacy/wrong ID sometimes set in Catalyst env — always remap to PAYROLL_ORGANIZATION_ID. */
const LEGACY_WRONG_ORG_IDS = new Set(['60006183023']);

export function getPayrollOrganizationId() {
  const env = String(process.env.REACT_APP_ZOHO_PAYROLL_ORGANIZATION_ID || '').trim();
  if (!env || LEGACY_WRONG_ORG_IDS.has(env)) return PAYROLL_ORGANIZATION_ID;
  return env;
}
