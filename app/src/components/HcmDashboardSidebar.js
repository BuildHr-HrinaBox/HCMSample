import React, { useEffect, useState } from 'react';
import { Link, useLocation } from 'react-router-dom';
import {
  LayoutDashboard,
  Building2,
  BookOpen,
  Calendar,
  ArrowLeftRight,
  ChevronDown,
  ChevronRight,
  FileText,
  Download,
} from 'lucide-react';
import '../Pages/newdashboard.css';
import vayonaBrandLogo from './Yanona Logo.png';
import { fetchAllowedActCategoriesFromSites } from '../utils/siteInchargeScope';

/** Must match CompanyDetails.js COMPANY_DETAILS_CLOSE_MODAL_EVENT */
const COMPANY_DETAILS_CLOSE_MODAL_EVENT = 'company-details-close-modal';

export const HCM_ORGANIZATION_CHILDREN = [
  { label: 'Company Details', to: '/rule-book/company-details' },
  { label: 'Site Management', to: '/rule-book/site-management' },
];

/** Statutory / checklist bulk / form master — nested under sidebar "Transaction" */
export const HCM_TRANSACTION_CHILDREN = [
  { label: 'Transaction', to: '/rule-book/statutory' },
  { label: 'Checklist Master', to: '/rule-book/checklistbulk' },
  { label: 'Form Master', to: '/rule-book/formmaster' },
];

/** Form Fetch — matches NewHomeSidebar formFetch group */
export const HCM_FORM_FETCH_CHILDREN = [
  { label: 'People', to: '/rule-book/people' },
  { label: 'Attendance', to: '/rule-book/attendance' },
  { label: 'Leave', to: '/rule-book/leave' },
  { label: 'Payroll', to: '/rule-book/payroll' },
];

const NAV = [
  { type: 'link', id: 'dashboard', label: 'Dashboard', icon: LayoutDashboard, to: '/hcm-dashboard' },
  {
    type: 'group',
    id: 'org',
    label: 'Organization',
    icon: Building2,
    children: HCM_ORGANIZATION_CHILDREN,
  },
  {
    type: 'group',
    id: 'library',
    label: 'Library',
    icon: BookOpen,
    children: [
      { label: 'Act Description', to: '/rule-book/act-descriptions' },
      { label: 'Acts Bulk', to: '/rule-book/actsbulk' },
    ],
  },
  {
    type: 'group',
    id: 'transaction',
    label: 'Transaction',
    icon: ArrowLeftRight,
    children: HCM_TRANSACTION_CHILDREN,
  },
  {
    type: 'group',
    id: 'formFetch',
    label: 'Form Fetch',
    icon: Download,
    children: HCM_FORM_FETCH_CHILDREN,
  },
  { type: 'link', id: 'calendar', label: 'Calendar', icon: Calendar, to: '/calendar-picker' },
  { type: 'link', id: 'main-report', label: 'Reports', icon: FileText, to: '/mainreport' },
];

/**
 * Same sidebar as /hcm-dashboard (newdashboard) — Vayona nav + org dropdown + user.
 */
export default function HcmDashboardSidebar({ userName = 'User', userRole = 'App User', userInitials = 'U', userEmail }) {
  const location = useLocation();
  // null = scope not loaded yet: hide site-restricted items so they never flash on screen for site users
  const [isSiteLoginUser, setIsSiteLoginUser] = useState(null);

  useEffect(() => {
    let cancelled = false;
    setIsSiteLoginUser(null);
    fetchAllowedActCategoriesFromSites(userEmail)
      .then((cats) => {
        if (cancelled) return;
        setIsSiteLoginUser(Array.isArray(cats) && cats.length > 0);
      })
      .catch(() => {
        if (!cancelled) setIsSiteLoginUser(false);
      });
    return () => {
      cancelled = true;
    };
  }, [userEmail]);

  const hiddenForSiteLogin = new Set(['Company Details', 'Acts Bulk', 'Checklist Master', 'Form Master']);
  const shouldApplySiteNavFilter = isSiteLoginUser !== false;
  const filteredNav = NAV.map((item) => {
    if (item.type !== 'group' || !Array.isArray(item.children) || !shouldApplySiteNavFilter) return item;
    return {
      ...item,
      children: item.children.filter((child) => !hiddenForSiteLogin.has(child.label)),
    };
  });

  const orgSectionActive = HCM_ORGANIZATION_CHILDREN.some((c) => location.pathname === c.to);
  const librarySectionActive = filteredNav.some(
    (i) => i.type === 'group' && i.id === 'library' && (i.children || []).some((c) => location.pathname === c.to)
  );
  const transactionSectionActive = filteredNav.some(
    (i) => i.type === 'group' && i.id === 'transaction' && (i.children || []).some((c) => location.pathname === c.to)
  );
  const formFetchSectionActive = filteredNav.some(
    (i) => i.type === 'group' && i.id === 'formFetch' && (i.children || []).some((c) => location.pathname === c.to)
  );
  // Accordion open group: only one group open at a time.
  const [openGroupId, setOpenGroupId] = useState(() => {
    if (orgSectionActive) return 'org';
    if (librarySectionActive) return 'library';
    if (transactionSectionActive) return 'transaction';
    if (formFetchSectionActive) return 'formFetch';
    return null;
  });

  useEffect(() => {
    if (orgSectionActive) setOpenGroupId('org');
    else if (librarySectionActive) setOpenGroupId('library');
    else if (transactionSectionActive) setOpenGroupId('transaction');
    else if (formFetchSectionActive) setOpenGroupId('formFetch');
  }, [orgSectionActive, librarySectionActive, transactionSectionActive, formFetchSectionActive]);

  const isLinkActive = (item) => {
    if (item.to === '/') return location.pathname === '/';
    if (item.to === '/hcm-dashboard') return location.pathname === '/hcm-dashboard';
    return location.pathname === item.to;
  };

  return (
    <aside className="nd-sidebar">
      <div className="nd-brand">
        <img
          src={vayonaBrandLogo}
          alt="Vayona Energy"
          className="nd-brand-logo"
          width={200}
          height={56}
          decoding="async"
        />
      </div>

      <div className="nd-sidebar-body">
        <nav className="nd-nav" aria-label="Main">
          {filteredNav.map((item) => {
            if (item.type === 'link') {
              const Icon = item.icon;
              const active = isLinkActive(item) && openGroupId === null;
              return (
                <Link
                  key={item.id}
                  to={item.to}
                  className={`nd-nav-item${item.id === 'calendar' ? ' nd-nav-item--calendar-centered' : ''}${active ? ' nd-nav-item--active' : ''}`}
                >
                  <Icon size={20} strokeWidth={2} />
                  <span>{item.label}</span>
                </Link>
              );
            }
            if (item.type === 'group') {
              const Icon = item.icon;
              const groupOpen = openGroupId === item.id;
              const headerActive =
                item.id === 'org'
                  ? groupOpen || (orgSectionActive && openGroupId === null)
                  : item.id === 'library'
                    ? groupOpen || (librarySectionActive && openGroupId === null)
                    : item.id === 'transaction'
                      ? groupOpen || (transactionSectionActive && openGroupId === null)
                      : item.id === 'formFetch'
                        ? groupOpen || (formFetchSectionActive && openGroupId === null)
                        : false;
              return (
                <div key={item.id} className="nd-nav-group">
                  <button
                    type="button"
                    className={`nd-nav-group-header${headerActive ? ' nd-nav-group-header--active' : ''}${groupOpen ? ' nd-nav-group-header--open' : ''}`}
                    onClick={() => setOpenGroupId((prev) => (prev === item.id ? null : item.id))}
                    aria-expanded={groupOpen}
                  >
                    <Icon size={20} strokeWidth={2} />
                    <span className="nd-nav-group-label">{item.label}</span>
                    {groupOpen ? (
                      <ChevronDown size={18} strokeWidth={2} className="nd-nav-group-chevron" />
                    ) : (
                      <ChevronRight size={18} strokeWidth={2} className="nd-nav-group-chevron" />
                    )}
                  </button>
                  <div className={`nd-nav-group-panel${groupOpen ? ' nd-nav-group-panel--open' : ''}`}>
                    <div className="nd-nav-group-panel-inner">
                      <div className="nd-nav-subitem-list">
                        {item.children.map((child) => {
                          const subActive = location.pathname === child.to;
                          return (
                            <Link
                              key={child.to}
                              to={child.to}
                              className={`nd-nav-subitem${subActive ? ' nd-nav-subitem--active' : ''}`}
                              onClick={() => {
                                if (child.to === '/rule-book/company-details' && location.pathname === child.to) {
                                  window.dispatchEvent(new CustomEvent(COMPANY_DETAILS_CLOSE_MODAL_EVENT));
                                }
                              }}
                            >
                              {child.label}
                            </Link>
                          );
                        })}
                      </div>
                    </div>
                  </div>
                </div>
              );
            }
            return null;
          })}
        </nav>

        <div className="nd-sidebar-spacer" />

      </div>
    </aside>
  );
}