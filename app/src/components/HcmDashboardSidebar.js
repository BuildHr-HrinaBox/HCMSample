import React, { useEffect, useRef, useState } from 'react';
import { Link, useLocation } from 'react-router-dom';
import {
  LayoutDashboard,
  Building2,
  BookOpen,
  Calendar,
  ArrowLeftRight,
  ChevronRight,
  FileText,
  Download,
  Settings as SettingsIcon,
} from 'lucide-react';
import '../Pages/newdashboard.css';
import vayonaBrandLogo from './Yanona Logo.png';
import { fetchAllowedActCategoriesFromSites } from '../utils/siteInchargeScope';
import { prefetchSettings } from '../utils/settingsCache';

/** Must match CompanyDetails.js COMPANY_DETAILS_CLOSE_MODAL_EVENT */
const COMPANY_DETAILS_CLOSE_MODAL_EVENT = 'company-details-close-modal';

export const HCM_ORGANIZATION_CHILDREN = [
  { label: 'Company Details', to: '/rule-book/company-details' },
  { label: 'Site Management', to: '/rule-book/site-management' },
];

/** Statutory / checklist bulk / form master — nested under sidebar "Transaction" */
export const HCM_TRANSACTION_CHILDREN = [
  { label: 'Statutory', to: '/rule-book/statutory' },
  { label: 'Checklist Master', to: '/rule-book/checklistbulk' },
  { label: 'Form Master', to: '/rule-book/formmaster' },
];

/** Form Fetch — matches NewHomeSidebar formFetch group */
export const HCM_FORM_FETCH_CHILDREN = [
  { label: 'People', to: '/rule-book/people' },
  { label: 'Attendance', to: '/rule-book/attendance' },
  { label: 'Leave', to: '/rule-book/leave' },
  { label: 'Payroll', to: '/rule-book/payroll' },
  { label: 'CLRA', to: '/rule-book/clra' },
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
  { type: 'link', id: 'settings', label: 'Settings', icon: SettingsIcon, to: '/settings' },
  { type: 'link', id: 'main-report', label: 'Reports', icon: FileText, to: '/mainreport' },
  { type: 'link', id: 'returned-report', label: 'Returned Report', icon: FileText, to: '/returned-report' },
];

/**
 * Same sidebar as /hcm-dashboard (newdashboard) — Vayona nav + org dropdown + user.
 */
export default function HcmDashboardSidebar({ userName = 'User', userRole = 'App User', userInitials = 'U', userEmail }) {
  const location = useLocation();
  // null = scope not loaded yet: hide site-restricted items so they never flash on screen for site users
  const [isSiteLoginUser, setIsSiteLoginUser] = useState(null);

  useEffect(() => {
    prefetchSettings();
  }, []);

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

  const hiddenForSiteLogin = new Set(['Company Details', 'Acts Bulk', 'Checklist Master']);
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
  const [openGroups, setOpenGroups] = useState({
    org: orgSectionActive,
    library: librarySectionActive,
    transaction: transactionSectionActive,
    formFetch: formFetchSectionActive,
  });
  const userCollapsedGroupsRef = useRef({});

  useEffect(() => {
    userCollapsedGroupsRef.current = {};
  }, [location.pathname]);

  useEffect(() => {
    setOpenGroups((prev) => ({
      ...prev,
      ...(orgSectionActive && !userCollapsedGroupsRef.current.org ? { org: true } : {}),
      ...(librarySectionActive && !userCollapsedGroupsRef.current.library ? { library: true } : {}),
      ...(transactionSectionActive && !userCollapsedGroupsRef.current.transaction
        ? { transaction: true }
        : {}),
      ...(formFetchSectionActive && !userCollapsedGroupsRef.current.formFetch ? { formFetch: true } : {}),
    }));
  }, [orgSectionActive, librarySectionActive, transactionSectionActive, formFetchSectionActive]);

  const toggleGroup = (groupId) => {
    setOpenGroups((prev) => {
      const nextOpen = !prev[groupId];
      if (!nextOpen) userCollapsedGroupsRef.current[groupId] = true;
      else delete userCollapsedGroupsRef.current[groupId];
      return { ...prev, [groupId]: nextOpen };
    });
  };

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
              const active = isLinkActive(item);
              return (
                <Link
                  key={item.id}
                  to={item.to}
                  title={item.label}
                  className={`nd-nav-item nd-nav-item--leaf${item.id === 'calendar' ? ' nd-nav-item--calendar-centered' : ''}${item.id === 'returned-report' ? ' nd-nav-item--returned-report' : ''}${active ? ' nd-nav-item--active' : ''}`}
                  onMouseEnter={item.id === 'settings' ? prefetchSettings : undefined}
                  onFocus={item.id === 'settings' ? prefetchSettings : undefined}
                  onClick={item.id === 'settings' ? prefetchSettings : undefined}
                >
                  <Icon size={20} strokeWidth={2} aria-hidden />
                  <span className="nd-nav-item-label">{item.label}</span>
                </Link>
              );
            }
            if (item.type === 'group') {
              const Icon = item.icon;
              const groupOpen = !!openGroups[item.id];
              const sectionActive =
                item.id === 'org'
                  ? orgSectionActive
                  : item.id === 'library'
                    ? librarySectionActive
                    : item.id === 'transaction'
                      ? transactionSectionActive
                      : item.id === 'formFetch'
                        ? formFetchSectionActive
                        : false;
              return (
                <div key={item.id} className="nd-nav-group">
                  <button
                    type="button"
                    className={`nd-nav-group-header${groupOpen ? ' nd-nav-group-header--active nd-nav-group-header--open' : ''}${!groupOpen && sectionActive ? ' nd-nav-group-header--route' : ''}`}
                    onClick={() => toggleGroup(item.id)}
                    aria-expanded={groupOpen}
                  >
                    <Icon size={20} strokeWidth={2} />
                    <span className="nd-nav-group-label">{item.label}</span>
                    <ChevronRight
                      size={18}
                      strokeWidth={2}
                      className={`nd-nav-group-chevron${groupOpen ? ' nd-nav-group-chevron--open' : ''}`}
                      aria-hidden
                    />
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