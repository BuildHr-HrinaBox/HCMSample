import React, { useEffect, useState } from 'react';
import {
  Building,
  ClipboardList,
  BookMarked,
  CheckCircle2,
  MapPin,
  ShieldCheck,
} from 'lucide-react';
import HcmDashboardSidebar from '../components/HcmDashboardSidebar';
import HcmHeaderProfileMenu from '../components/HcmHeaderProfileMenu';
import ComplianceHealthDashboard from '../components/ComplianceHealthDashboard';
import {
  buildInchargeSiteScopeFromList,
  fetchInchargeDisplayScopeFromSites,
  filterSitesForLoginUser,
  getActCategoryFromActSector,
  industryLabelToActCategory,
  isOrgWideSiteViewer,
  sectorMatchesInchargeSiteIndustries,
  statesFieldMatchesInchargeSiteStates,
} from '../utils/siteInchargeScope';
import { resolveLoginEmailString } from '../utils/resolveLoginEmail';
import './newdashboard.css';

const DEFAULT_METRICS = [
  {
    key: 'companies',
    label: 'Active Companies',
    value: 0,
    hint: 'Registered organizations',
    icon: Building,
    tone: 'green',
  },
  {
    key: 'sites',
    label: 'Active Site',
    value: 0,
    hint: 'Locations under management',
    icon: MapPin,
    tone: 'orange',
  },
  {
    key: 'due',
    label: 'Approved',
    value: 0,
    hint: '0% of total',
    icon: CheckCircle2,
    tone: 'teal',
  },
  {
    key: 'policies',
    label: 'Returned',
    value: 0,
    hint: '0% of total',
    icon: BookMarked,
    tone: 'purple',
  },
  {
    key: 'approvals',
    label: 'Pending for Approval',
    value: 0,
    hint: '0% of total',
    icon: ClipboardList,
    tone: 'blue',
  },
];

const COMPANY_API = '/server/company_function/company';
/** Site count from Catalyst sitemanagement_function (same as Site Management page). */
const SITE_MANAGEMENT_API = '/server/sitemanagement_function/sitemanagement';
const STATUTORY_API = '/server/statutoryreg_function/statutory';

const SITE_MANAGEMENT_CACHE_KEY = 'siteManagementData';

function readSiteDetailsFromCache() {
  try {
    const cached = localStorage.getItem(SITE_MANAGEMENT_CACHE_KEY);
    if (!cached) return [];
    const parsed = JSON.parse(cached);
    return Array.isArray(parsed) ? parsed : [];
  } catch (_) {
    return [];
  }
}

/** Fetch site rows from sitemanagement_function (same source as Site Management). */
async function fetchSiteDetailsFromSiteManagement() {
  try {
    const res = await fetch(SITE_MANAGEMENT_API, { cache: 'no-store' });
    const text = await res.text();
    const trimmed = text.trim();
    if (!trimmed || trimmed[0] === '<') {
      return readSiteDetailsFromCache();
    }
    const data = JSON.parse(trimmed);
    const siteDetails = data?.data?.siteDetails;
    if (Array.isArray(siteDetails)) {
      if (siteDetails.length > 0) {
        try {
          localStorage.setItem(SITE_MANAGEMENT_CACHE_KEY, JSON.stringify(siteDetails));
        } catch (_) {}
      }
      return siteDetails;
    }
    return readSiteDetailsFromCache();
  } catch (_) {
    return readSiteDetailsFromCache();
  }
}

/**
 * Keep statutory rows that belong to the logged-in site incharge scope
 * (Site Management industry → act category + state), same rules as Calendar / Act Description.
 */
function statutoryRowMatchesSiteScope(row, scope) {
  const actCategories = scope?.actCategories;
  if (!Array.isArray(actCategories) || actCategories.length === 0) return true;

  const act = row?.act ?? row?.Act ?? '';
  const sector = row?.sector ?? row?.Sector ?? '';
  const cat = getActCategoryFromActSector(act, sector);
  if (!actCategories.includes(cat)) return false;

  if (String(sector || '').trim()) {
    const sectorCat =
      industryLabelToActCategory(sector) || getActCategoryFromActSector('', sector);
    if (sectorCat && sectorCat !== 'other' && !actCategories.includes(sectorCat)) {
      return false;
    }
    const inds = scope?.industryLabels;
    if (
      Array.isArray(inds) &&
      inds.length > 0 &&
      sectorCat &&
      sectorCat !== 'other' &&
      !sectorMatchesInchargeSiteIndustries(sector, inds)
    ) {
      return false;
    }
  }

  const sts = scope?.stateLabels;
  if (Array.isArray(sts) && sts.length > 0) {
    const stateField = row?.states ?? row?.state ?? row?.State ?? '';
    // Same as Statutory: blank / other states excluded for site incharge.
    if (!statesFieldMatchesInchargeSiteStates(stateField, sts)) {
      return false;
    }
  }

  const siteNames = scope?.siteNames;
  if (Array.isArray(siteNames) && siteNames.length > 0) {
    const rowSite = String(row?.site ?? row?.Site ?? row?.siteName ?? row?.SiteName ?? '')
      .trim()
      .toLowerCase();
    if (rowSite) {
      const allowed = new Set(siteNames.map((s) => String(s).trim().toLowerCase()).filter(Boolean));
      if (!allowed.has(rowSite)) return false;
    }
  }

  return true;
}

/** Align dashboard KPI buckets with Statutory Status column (Pending / Approved / Returned). */
function getStatutoryScorecardStatus(item) {
  const raw = item?.status != null && item.status !== '' ? item.status : item?.Status;
  const norm = String(raw || '').trim();
  const normalized = norm.toLowerCase();
  const rawApproval =
    item?.approval != null && item.approval !== '' ? item.approval : item?.Approval;
  const approvalNorm = String(rawApproval || '').trim().toLowerCase();

  if (
    approvalNorm === 'approved' ||
    approvalNorm === 'approve' ||
    normalized === 'approved' ||
    normalized === 'approve'
  ) {
    return 'approved';
  }
  if (
    approvalNorm === 'rejected' ||
    approvalNorm === 'reject' ||
    normalized === 'rejected' ||
    normalized === 'reject'
  ) {
    return 'returned';
  }

  const sendForApproval = String(item?.sendForApproval ?? item?.SendForApproval ?? '')
    .trim()
    .toLowerCase();
  if (sendForApproval === 'sent' || normalized === 'pending') {
    return 'pending';
  }

  return 'other';
}

function NewDashboard({ userName = 'Ravi Kumar', userRole = 'HR Admin', userInitials = 'RA', userEmail }) {
  const [metrics, setMetrics] = useState(DEFAULT_METRICS);

  const now = new Date();
  const welcomeDate = now.toLocaleDateString('en-IN', {
    weekday: 'long',
    day: 'numeric',
    month: 'long',
    year: 'numeric',
  });
  const welcomeTime = now.toLocaleTimeString('en-IN', {
    hour: '2-digit',
    minute: '2-digit',
    hour12: true,
    timeZoneName: 'short',
  });

  useEffect(() => {
    let cancelled = false;

    const loadMetrics = async () => {
      try {
        const [companyRes, siteDetails, statutoryRes, siteScope, loginEmail] =
          await Promise.all([
            fetch(COMPANY_API, { cache: 'no-store' }),
            fetchSiteDetailsFromSiteManagement(),
            fetch(STATUTORY_API, { cache: 'no-store' }),
            fetchInchargeDisplayScopeFromSites(userEmail),
            resolveLoginEmailString(userEmail),
          ]);

        // Active Site = sites assigned to this login in Site Management (incharge email match).
        const inchargeScope = buildInchargeSiteScopeFromList(siteDetails, loginEmail);
        const visibleSites = filterSitesForLoginUser(siteDetails, loginEmail, userRole);
        const activeSites = inchargeScope?.mine?.length
          ? inchargeScope.mine.length
          : isOrgWideSiteViewer(userRole)
            ? (Array.isArray(siteDetails) ? siteDetails.length : 0)
            : visibleSites.length;

        let activeCompanies = 0;
        if (companyRes.ok) {
          const companyJson = await companyRes.json();
          const companyDetails = companyJson?.data?.companyDetails;
          activeCompanies = Array.isArray(companyDetails) ? companyDetails.length : 0;
        }

        let pendingApprovals = 0;
        let approvedForms = 0;
        let rejectedForms = 0;
        const hasSiteScope =
          Array.isArray(siteScope?.actCategories) && siteScope.actCategories.length > 0;

        if (statutoryRes.ok) {
          const statutoryJson = await statutoryRes.json();
          const allStatutory = Array.isArray(statutoryJson?.data?.statutoryData)
            ? statutoryJson.data.statutoryData
            : [];

          // Site login: count only statutory rows for that Site Management assignment.
          const statutoryData = hasSiteScope
            ? allStatutory.filter((item) => statutoryRowMatchesSiteScope(item, siteScope))
            : allStatutory;

          statutoryData.forEach((item) => {
            const bucket = getStatutoryScorecardStatus(item);
            if (bucket === 'pending') pendingApprovals += 1;
            else if (bucket === 'approved') approvedForms += 1;
            else if (bucket === 'returned') rejectedForms += 1;
          });
        }

        if (!cancelled) {
          const scorecardTotal = pendingApprovals + approvedForms + rejectedForms;
          const pctOfTotal = (n) =>
            scorecardTotal > 0 ? `${Math.round((n / scorecardTotal) * 100)}% of total` : '0% of total';

          setMetrics((prev) =>
            prev.map((metric) => {
              if (metric.key === 'companies') {
                return { ...metric, value: activeCompanies, hint: 'Registered organizations' };
              }
              if (metric.key === 'sites') {
                return { ...metric, value: activeSites, hint: 'Locations under management' };
              }
              if (metric.key === 'approvals') {
                return { ...metric, value: pendingApprovals, hint: pctOfTotal(pendingApprovals) };
              }
              if (metric.key === 'due') {
                return { ...metric, value: approvedForms, hint: pctOfTotal(approvedForms) };
              }
              if (metric.key === 'policies') {
                return { ...metric, value: rejectedForms, hint: pctOfTotal(rejectedForms) };
              }
              return metric;
            })
          );
        }
      } catch (_) {
        if (!cancelled) {
          setMetrics((prev) =>
            prev.map((metric) => {
              if (
                metric.key === 'companies' ||
                metric.key === 'sites' ||
                metric.key === 'approvals' ||
                metric.key === 'due' ||
                metric.key === 'policies'
              ) {
                return { ...metric, value: 0 };
              }
              return metric;
            })
          );
        }
      }
    };

    loadMetrics();

    return () => {
      cancelled = true;
    };
  }, [userEmail, userRole]);

  return (
    <div className="nd-root">
      <HcmDashboardSidebar userName={userName} userRole={userRole} userInitials={userInitials} userEmail={userEmail} />

      <div className="nd-main">
        <header className="nd-header">
          <h1 className="nd-header-title">HR Compliance Management</h1>
          <div className="nd-header-actions">
            <HcmHeaderProfileMenu userInitials={userInitials} userEmail={userEmail} />
          </div>
        </header>

        <main className="nd-content">
          <section className="nd-welcome">
            <div className="nd-welcome-bg" aria-hidden />
            <div className="nd-welcome-inner">
              <div className="nd-welcome-copy">
                <h2 className="nd-welcome-heading">Welcome to Vayona Energy HCM 👋</h2>
                <p className="nd-welcome-desc">
                  Track compliance, approvals, and statutory tasks across your organization in one
                  place.
                </p>
                <p className="nd-welcome-meta">
                  {welcomeDate} | {welcomeTime}
                </p>
              </div>
              <div className="nd-welcome-art" aria-hidden>
                <div className="nd-welcome-checklist">
                  <div className="nd-welcome-check-row">
                    <CheckCircle2 size={18} className="nd-wc-ok" />
                    <span />
                  </div>
                  <div className="nd-welcome-check-row">
                    <CheckCircle2 size={18} className="nd-wc-ok" />
                    <span />
                  </div>
                  <div className="nd-welcome-check-row">
                    <CheckCircle2 size={18} className="nd-wc-dim" />
                    <span />
                  </div>
                </div>
                <div className="nd-welcome-shield">
                  <ShieldCheck size={40} strokeWidth={1.5} />
                </div>
              </div>
            </div>
          </section>

          <section className="nd-metrics">
            {metrics.map((m) => {
              const Icon = m.icon;
              return (
                <article key={m.key} className={`nd-metric nd-metric--${m.tone}`}>
                  <div className="nd-metric-icon-wrap">
                    <Icon size={20} strokeWidth={2} />
                  </div>
                  <div className="nd-metric-body">
                    <span className="nd-metric-label">{m.label}</span>
                    <span className="nd-metric-value">{m.value}</span>
                    {m.hint ? <span className="nd-metric-hint">{m.hint}</span> : null}
                  </div>
                </article>
              );
            })}
          </section>

          <ComplianceHealthDashboard userRole={userRole} userEmail={userEmail} />
        </main>
      </div>
    </div>
  );
}

export default NewDashboard;
