import React, { useCallback, useEffect, useState } from 'react';
import {
  Building,
  BookMarked,
  ClipboardList,
  CheckCircle2,
  FileText,
  MapPin,
  ShieldCheck,
} from 'lucide-react';
import HcmDashboardSidebar from '../components/HcmDashboardSidebar';
import HcmHeaderProfileMenu from '../components/HcmHeaderProfileMenu';
import ComplianceHealthDashboard from '../components/ComplianceHealthDashboard';
import {
  buildInchargeSiteScopeFromList,
  filterSitesForLoginUser,
  isOrgWideSiteViewer,
} from '../utils/siteInchargeScope';
import { resolveLoginEmailString } from '../utils/resolveLoginEmail';
import './newdashboard.css';

const OVERVIEW_METRICS = [
  {
    key: 'companies',
    label: 'Active Companies',
    value: 0,
    icon: Building,
    tone: 'green',
  },
  {
    key: 'sites',
    label: 'Active Sites',
    value: 0,
    icon: MapPin,
    tone: 'orange',
  },
];

const STATUS_METRICS = [
  {
    key: 'due',
    label: 'Approved',
    value: 0,
    icon: CheckCircle2,
    tone: 'green',
  },
  {
    key: 'policies',
    label: 'Returned',
    value: 0,
    icon: BookMarked,
    tone: 'purple',
  },
  {
    key: 'approvals',
    label: 'Pending for Approval',
    value: 0,
    icon: ClipboardList,
    tone: 'blue',
  },
  {
    key: 'yetToSubmit',
    label: 'Yet to Submit',
    value: 0,
    icon: FileText,
    tone: 'orange',
  },
];

const COMPANY_API = '/server/company_function/company';
/** Site count from Catalyst sitemanagement_function (same as Site Management page). */
const SITE_MANAGEMENT_API = '/server/sitemanagement_function/sitemanagement';

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

function NewDashboard({ userName = 'Ravi Kumar', userRole = 'HR Admin', userInitials = 'RA', userEmail }) {
  const [overviewMetrics, setOverviewMetrics] = useState(OVERVIEW_METRICS);
  const [statusMetrics, setStatusMetrics] = useState(STATUS_METRICS);

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

  /** Mirror Compliance Health bucket counts on the top Compliance Status cards. */
  const handleHealthCounts = useCallback((counts) => {
    if (!counts || typeof counts !== 'object') return;
    const approved = Number(counts.approved) || 0;
    const pending = Number(counts.pending) || 0;
    const returned = Number(counts.returned) || 0;
    const yetToSubmit = Number(counts.yetToSubmit) || 0;

    setStatusMetrics((prev) =>
      prev.map((metric) => {
        const valueByKey = {
          due: approved,
          policies: returned,
          approvals: pending,
          yetToSubmit,
        };
        if (!(metric.key in valueByKey)) return metric;
        return { ...metric, value: valueByKey[metric.key] };
      })
    );
  }, []);

  useEffect(() => {
    let cancelled = false;

    const loadOverviewMetrics = async () => {
      try {
        const [companyRes, siteDetails, loginEmail] = await Promise.all([
          fetch(COMPANY_API, { cache: 'no-store' }),
          fetchSiteDetailsFromSiteManagement(),
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

        if (!cancelled) {
          setOverviewMetrics((prev) =>
            prev.map((metric) => {
              if (metric.key === 'companies') {
                return { ...metric, value: activeCompanies };
              }
              if (metric.key === 'sites') {
                return { ...metric, value: activeSites };
              }
              return metric;
            })
          );
        }
      } catch (_) {
        if (!cancelled) {
          setOverviewMetrics((prev) =>
            prev.map((metric) =>
              metric.key === 'companies' || metric.key === 'sites'
                ? { ...metric, value: 0 }
                : metric
            )
          );
        }
      }
    };

    loadOverviewMetrics();

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
                    <CheckCircle2 size={20} className="nd-wc-ok" />
                    <span />
                  </div>
                  <div className="nd-welcome-check-row">
                    <CheckCircle2 size={20} className="nd-wc-ok" />
                    <span />
                  </div>
                  <div className="nd-welcome-check-row">
                    <CheckCircle2 size={20} className="nd-wc-dim" />
                    <span />
                  </div>
                </div>
                <div className="nd-welcome-shield">
                  <ShieldCheck size={48} strokeWidth={1.5} />
                </div>
              </div>
            </div>
          </section>

          <section className="nd-metrics" aria-label="Dashboard overview and compliance status">
            <div className="nd-metrics-panel nd-metrics-panel--overview">
              <h3 className="nd-metrics-panel-title nd-metrics-panel-title--overview">
                Organization & Site Overview
              </h3>
              <div className="nd-metrics-overview-grid">
                {overviewMetrics.map((m) => {
                  const Icon = m.icon;
                  return (
                    <article key={m.key} className={`nd-metric nd-metric--${m.tone}`}>
                      <div className="nd-metric-icon-wrap">
                        <Icon size={24} strokeWidth={2} />
                      </div>
                      <div className="nd-metric-body">
                        <span className="nd-metric-label">{m.label}</span>
                        <span className="nd-metric-value">{m.value}</span>
                        {m.hint ? <span className="nd-metric-hint">{m.hint}</span> : null}
                      </div>
                    </article>
                  );
                })}
              </div>
            </div>

            <div className="nd-metrics-panel nd-metrics-panel--status">
              <h3 className="nd-metrics-panel-title nd-metrics-panel-title--status">
                Compliance Status
              </h3>
              <div className="nd-metrics-status-grid">
                {statusMetrics.map((m) => {
                  const Icon = m.icon;
                  return (
                    <article key={m.key} className={`nd-metric nd-metric--status nd-metric--${m.tone}`}>
                      <div className="nd-metric-icon-wrap">
                        <Icon size={24} strokeWidth={2} />
                      </div>
                      <div className="nd-metric-body">
                        <span className="nd-metric-label">{m.label}</span>
                        <span className="nd-metric-value">{m.value}</span>
                      </div>
                    </article>
                  );
                })}
              </div>
            </div>
          </section>

          <ComplianceHealthDashboard
            userRole={userRole}
            userEmail={userEmail}
            onCountsChange={handleHealthCounts}
          />
        </main>
      </div>
    </div>
  );
}

export default NewDashboard;
