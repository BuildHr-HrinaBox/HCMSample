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
import Dashboard from './Dashboard';
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
    key: 'approvals',
    label: 'Pending for Approval',
    value: 0,
    hint: '0% of total',
    icon: ClipboardList,
    tone: 'blue',
  },
  {
    key: 'due',
    label: 'Approved',
    value: 0,
    hint: '0% of total',
    hintTone: 'success',
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
];

function formatShareHint(count, total) {
  if (!total) return '0% of total';
  const pct = ((count / total) * 100).toFixed(2).replace(/\.00$/, '');
  return `${pct}% of total`;
}

const COMPANY_API = '/server/company_function/company';
/** Site count from Catalyst sitemanagement_function (same as Site Management page). */
const SITE_MANAGEMENT_API = '/server/sitemanagement_function/sitemanagement';
const STATUTORY_API = '/server/statutoryreg_function/statutory';

const SITE_MANAGEMENT_CACHE_KEY = 'siteManagementData';

function activeSiteCountFromCache() {
  try {
    const cached = localStorage.getItem(SITE_MANAGEMENT_CACHE_KEY);
    if (!cached) return 0;
    const parsed = JSON.parse(cached);
    return Array.isArray(parsed) ? parsed.length : 0;
  } catch (_) {
    return 0;
  }
}

/** Parse active site count from sitemanagement_function GET /sitemanagement (same as Site Management). */
async function fetchActiveSiteCountFromSiteManagement() {
  try {
    const res = await fetch(SITE_MANAGEMENT_API, { cache: 'no-store' });
    const text = await res.text();
    const trimmed = text.trim();
    if (!trimmed || trimmed[0] === '<') {
      return activeSiteCountFromCache();
    }
    const data = JSON.parse(trimmed);
    const siteDetails = data?.data?.siteDetails;
    const listCount = Array.isArray(siteDetails) ? siteDetails.length : 0;
    const total = parseInt(data?.data?.total, 10);
    const fromTotal = !Number.isNaN(total) && total >= 0 ? total : 0;
    const count = Math.max(fromTotal, listCount);
    if (count > 0) {
      if (listCount > 0) {
        try {
          localStorage.setItem(SITE_MANAGEMENT_CACHE_KEY, JSON.stringify(siteDetails));
        } catch (_) {}
      }
      return count;
    }
    if (data?.status === 'success' || listCount > 0) return count;
    return activeSiteCountFromCache();
  } catch (_) {
    return activeSiteCountFromCache();
  }
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
        const [companyRes, activeSites, statutoryRes] = await Promise.all([
          fetch(COMPANY_API, { cache: 'no-store' }),
          fetchActiveSiteCountFromSiteManagement(),
          fetch(STATUTORY_API, { cache: 'no-store' }),
        ]);

        let activeCompanies = 0;
        if (companyRes.ok) {
          const companyJson = await companyRes.json();
          const companyDetails = companyJson?.data?.companyDetails;
          activeCompanies = Array.isArray(companyDetails) ? companyDetails.length : 0;
        }

        let pendingApprovals = 0;
        let approvedForms = 0;
        let rejectedForms = 0;
        let statutoryTotal = 0;
        if (statutoryRes.ok) {
          const statutoryJson = await statutoryRes.json();
          const statutoryData = Array.isArray(statutoryJson?.data?.statutoryData)
            ? statutoryJson.data.statutoryData
            : [];

          statutoryTotal = statutoryData.length;

          pendingApprovals = statutoryData.filter((item) => {
            const status = String(item?.status ?? item?.Status ?? '').trim().toLowerCase();
            return status === 'pending';
          }).length;

          approvedForms = statutoryData.filter((item) => {
            const status = String(item?.status ?? item?.Status ?? '').trim().toLowerCase();
            return status === 'approved';
          }).length;

          rejectedForms = statutoryData.filter((item) => {
            const status = String(item?.status ?? item?.Status ?? '').trim().toLowerCase();
            return status === 'rejected' || status === 'reject' || status === 'returned';
          }).length;
        }

        if (!cancelled) {
          setMetrics((prev) =>
            prev.map((metric) => {
              if (metric.key === 'companies') {
                return { ...metric, value: activeCompanies, hint: 'Registered organizations' };
              }
              if (metric.key === 'sites') {
                return { ...metric, value: activeSites, hint: 'Locations under management' };
              }
              if (metric.key === 'approvals') {
                return {
                  ...metric,
                  value: pendingApprovals,
                  hint: formatShareHint(pendingApprovals, statutoryTotal),
                };
              }
              if (metric.key === 'due') {
                return {
                  ...metric,
                  value: approvedForms,
                  hint: formatShareHint(approvedForms, statutoryTotal),
                  hintTone: 'success',
                };
              }
              if (metric.key === 'policies') {
                return {
                  ...metric,
                  value: rejectedForms,
                  hint: formatShareHint(rejectedForms, statutoryTotal),
                };
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
                return {
                  ...metric,
                  value: 0,
                  hint:
                    metric.key === 'companies'
                      ? 'Registered organizations'
                      : metric.key === 'sites'
                        ? 'Locations under management'
                        : '0% of total',
                };
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
  }, []);

  return (
    <div className="nd-root" style={{ background: '#ffffff', backgroundColor: '#ffffff' }}>
      <HcmDashboardSidebar userName={userName} userRole={userRole} userInitials={userInitials} userEmail={userEmail} />

      <div className="nd-main" style={{ background: '#ffffff', backgroundColor: '#ffffff' }}>
        <header className="nd-header" style={{ background: '#ffffff', backgroundColor: '#ffffff' }}>
          <h1 className="nd-header-title">HR Compliance Management</h1>
          <div className="nd-header-actions">
            <HcmHeaderProfileMenu userInitials={userInitials} userEmail={userEmail} />
          </div>
        </header>

        <main className="nd-content" style={{ background: '#ffffff', backgroundColor: '#ffffff' }}>
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
                    <Icon size={22} strokeWidth={2} />
                  </div>
                  <div className="nd-metric-body">
                    <span className="nd-metric-value">{m.value}</span>
                    <span className="nd-metric-label">{m.label}</span>
                    {m.hint ? (
                      <span
                        className={`nd-metric-hint${m.hintTone === 'success' ? ' nd-metric-hint--success' : ''}`}
                      >
                        {m.hint}
                      </span>
                    ) : null}
                  </div>
                </article>
              );
            })}
          </section>

          <Dashboard userRole={userRole} userEmail={userEmail} />
        </main>
      </div>
    </div>
  );
}

export default NewDashboard;
