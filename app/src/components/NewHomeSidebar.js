import React, { useState, useEffect } from 'react';
import { Link, useLocation } from 'react-router-dom';
import {
  Home,
  LayoutDashboard,
  LayoutTemplate,
  Building2,
  BookOpen,
  Calendar,
  ArrowLeftRight,
  Leaf,
  PieChart,
  FileText,
  BarChart3,
  X,
  ChevronDown,
  ChevronRight,
  MapPin,
  FileCheck,
} from 'lucide-react';
import '../Pages/NewHome.css';
import { fetchAllowedActCategoriesFromSites } from '../utils/siteInchargeScope';

const AUDIT_ONLY_USER_EMAIL = 'afrindinu14@gmail.com';

const navigationItems = [
  { type: 'item', icon: <Home size={20} strokeWidth={2} />, label: 'Home', path: '/', color: 'purple' },
  { type: 'item', icon: <LayoutDashboard size={20} strokeWidth={2} />, label: 'Dashboard', path: '/new-dashboard', color: 'blue' },
  {
    type: 'item',
    icon: <LayoutTemplate size={20} strokeWidth={2} />,
    label: 'HCM Dashboard',
    path: '/hcm-dashboard',
    color: 'green',
  },
  {
    type: 'group',
    key: 'organizationMaster',
    icon: <Building2 size={20} strokeWidth={2} />,
    label: 'Organization Master',
    children: [
      { label: 'Company Details', path: '/rule-book/company-details' },
      { label: 'Site Management', path: '/rule-book/site-management' },
      { label: 'Certificate Management', path: '/certificate' },
      { label: 'Employee Data', path: '/rule-book/employee-data' },
    ],
  },
  {
    type: 'group',
    key: 'library',
    icon: <BookOpen size={20} strokeWidth={2} />,
    label: 'Library',
    children: [
      { label: 'Applicable Acts Library', path: '/rule-book/act-descriptions' },
      { label: 'Acts Bulk', path: '/rule-book/actsbulk' },
      { label: 'Quiz creation', path: '/quiz/create' },
    ],
  },
  {
    type: 'group',
    key: 'calendar',
    icon: <Calendar size={20} strokeWidth={2} />,
    label: 'Calendar',
    children: [
      { label: 'Calendar Picker', path: '/calendar-picker' },
      { label: 'Calendar Master', path: '/rule-book/calendarbulk' },
      { label: 'Calendar PDF Master', path: '/rule-book/calendar-pdf-master' },
    ],
  },
  {
    type: 'group',
    key: 'checklist',
    icon: <ArrowLeftRight size={20} strokeWidth={2} />,
    label: 'Transaction',
    children: [
      { label: 'Checklist', path: '/rule-book/checklist' },
      { label: 'Checklist Master', path: '/rule-book/checklistbulk' },
      { label: 'Form Master', path: '/rule-book/formmaster' },
      { label: 'To-Do List', path: '/rule-book/todo-list' },
      { label: 'Statutory', path: '/rule-book/statutory' },
      {
        type: 'group',
        key: 'transactionStatutoryForms',
        label: 'Statutory Forms',
        children: [
          { label: 'Factory', path: '/rule-book/factory' },
          { label: 'S&E Master', path: '/rule-book/se-master' },
        ],
      },
    ],
  },
  {
    type: 'group',
    key: 'assessments',
    icon: <PieChart size={20} strokeWidth={2} />,
    label: 'Audit',
    children: [
      { label: 'Assessment Report', path: '/assessments/assessment-report' },
      { label: 'Compliance', path: '/compliance' },
      { label: 'Audit Master', path: '/rule-book/audit-master' },
      { label: 'Audit Master Import', path: '/rule-book/audit-master-import' },
      { label: 'Main Audit', path: '/rule-book/main-audit' },
      { label: 'Audit Report', path: '/audit-report' },
      { label: 'Reports', path: '/mainreport' },
      {
        type: 'group',
        key: 'ruleBook',
        label: 'Rule Book',
        children: [
          { label: 'Client Details', path: '/rule-book/client-details' },
          { label: 'Scoring Criteria', path: '/rule-book/scoring-criteria' },
          { label: 'Record Category', path: '/rule-book/records' },
          { label: 'States', path: '/rule-book/states' },
          { label: 'Unique ID', path: '/rule-book/unique-ids' },
          { label: 'Unique ID List', path: '/rule-book/uniqueid-list' },
        ],
      },
    ],
  },
  {
    type: 'group',
    key: 'formFetch',
    icon: <FileText size={20} strokeWidth={2} />,
    label: 'Form Fetch',
    children: [
      { label: 'People', path: '/rule-book/people' },
      { label: 'Attendance', path: '/rule-book/attendance' },
      { label: 'Leave', path: '/rule-book/leave' },
      { label: 'Payroll', path: '/rule-book/payroll' },
      { label: 'CLRA', path: '/rule-book/clra' },
    ],
  },
  { type: 'item', icon: <BarChart3 size={20} strokeWidth={2} />, label: 'Report', path: '/reports', color: 'blue' },
  { type: 'item', icon: <FileText size={20} strokeWidth={2} />, label: 'Returned Report', path: '/returned-report', color: 'blue' },
];

const auditOnlyNavItems = [
  { type: 'item', icon: <Home size={20} strokeWidth={2} />, label: 'Home', path: '/', color: 'purple' },
  {
    type: 'group',
    key: 'audit',
    icon: <PieChart size={20} strokeWidth={2} />,
    label: 'Audit',
    children: [
      { label: 'Main Audit', path: '/rule-book/main-audit' },
      { label: 'Audit Report', path: '/audit-report' },
      { label: 'Reports', path: '/mainreport' },
      { label: 'Returned Report', path: '/returned-report' },
    ],
  },
  {
    type: 'group',
    key: 'statutory',
    icon: <FileCheck size={20} strokeWidth={2} />,
    label: 'Statutory',
    children: [{ label: 'Statutory', path: '/rule-book/statutory' }],
  },
];

function sidebarUserInitials(name, email) {
  if (name && String(name).trim()) {
    const parts = String(name).trim().split(/\s+/).filter(Boolean);
    if (parts.length >= 2) {
      return `${parts[0][0]}${parts[1][0]}`.toUpperCase();
    }
    return parts[0].slice(0, 2).toUpperCase();
  }
  if (email && String(email).includes('@')) {
    return String(email).slice(0, 2).toUpperCase();
  }
  return 'U';
}

export default function NewHomeSidebar({ isOpen, onClose, userEmail, userName, userRole }) {
  const location = useLocation();
  const [sites, setSites] = useState([]);
  // null = scope not loaded yet: hide site-restricted items so they never flash for site users
  const [isSiteLoginUser, setIsSiteLoginUser] = useState(null);
  const [openGroups, setOpenGroups] = useState({
    organizationMaster: false,
    library: false,
    checklist: false,
    calendar: false,
    assessments: true,
    audit: true,
    statutory: true,
    formFetch: false,
    transactionStatutoryForms: false,
    ruleBook: false,
  });

  const isAuditOnlyUser = (userEmail || '').trim().toLowerCase() === AUDIT_ONLY_USER_EMAIL;

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

  const displayName =
    (userName && String(userName).trim()) ||
    (userEmail && userEmail.includes('@') ? userEmail.split('@')[0] : '') ||
    'User';
  const displayRole = (userRole && String(userRole).trim()) || 'App User';
  const initials = sidebarUserInitials(userName, userEmail);

  useEffect(() => {
    if (!isAuditOnlyUser) return;
    let cancelled = false;
    fetch('/server/sitemanagement_function/sitemanagement')
      .then((res) => (res.ok ? res.json() : null))
      .then((data) => {
        if (cancelled) return;
        if (data?.status === 'success' && Array.isArray(data?.data?.siteDetails)) {
          setSites(data.data.siteDetails);
        }
      })
      .catch(() => {});
    return () => {
      cancelled = true;
    };
  }, [isAuditOnlyUser]);

  const toggleGroup = (key) => {
    setOpenGroups((prev) => ({ ...prev, [key]: !prev[key] }));
  };

  const itemsToRender = (() => {
    if (!isAuditOnlyUser) return navigationItems;
    const auditGroup = auditOnlyNavItems.find((i) => i.type === 'group' && i.key === 'audit');
    const statutoryGroup = auditOnlyNavItems.find((i) => i.type === 'group' && i.key === 'statutory');
    const auditSiteLinks = sites.map((site) => ({
      label: site.siteName || 'Unnamed Site',
      path: `/rule-book/main-audit?site=${encodeURIComponent(site.siteName || '')}`,
      isSiteLink: true,
    }));
    const statutorySiteLinks = sites.map((site) => ({
      label: site.siteName || 'Unnamed Site',
      path: `/rule-book/statutory?site=${encodeURIComponent(site.siteName || '')}`,
      isSiteLink: true,
    }));
    const auditGroupWithSites = auditGroup
      ? { ...auditGroup, children: [...(auditGroup.children || []), ...auditSiteLinks] }
      : auditGroup;
    const statutoryGroupWithSites = statutoryGroup
      ? { ...statutoryGroup, children: [...(statutoryGroup.children || []), ...statutorySiteLinks] }
      : statutoryGroup;
    return auditOnlyNavItems.map((item) => {
      if (item.type === 'group' && item.key === 'audit') return auditGroupWithSites;
      if (item.type === 'group' && item.key === 'statutory') return statutoryGroupWithSites;
      return item;
    });
  })();

  const hiddenLabelsForSiteLogin = new Set(['Company Details', 'Acts Bulk', 'Checklist Master']);
  const shouldApplySiteNavFilter = isSiteLoginUser !== false;
  const filteredItemsToRender = shouldApplySiteNavFilter
    ? itemsToRender
        .map((item) => {
          if (item.type === 'item') {
            return hiddenLabelsForSiteLogin.has(item.label) ? null : item;
          }
          if (item.type === 'group' && Array.isArray(item.children)) {
            return {
              ...item,
              children: item.children.filter((child) => !hiddenLabelsForSiteLogin.has(child.label)),
            };
          }
          return item;
        })
        .filter(Boolean)
    : itemsToRender;

  const isGroupActive = (group) => {
    if (!group.children) return false;
    return group.children.some((child) => {
      if (child.path) {
        const fullPath = location.pathname + location.search;
        return fullPath === child.path || location.pathname === child.path;
      }
      if (child.type === 'group' && child.children) {
        return child.children.some((grandchild) => location.pathname === grandchild.path);
      }
      return false;
    });
  };

  return (
    <>
      {isOpen && <div className="new-home-sidebar-overlay" onClick={onClose} aria-hidden="true" />}
      <div className={`new-home-sidebar ${isOpen ? 'open' : ''}`}>
        <div className="new-home-sidebar-header">
          <div className="new-home-sidebar-brand">
            <div className="new-home-sidebar-brand-icon" aria-hidden>
              <Leaf size={22} strokeWidth={2.2} />
            </div>
            <span className="new-home-sidebar-brand-text">VAYONA ENERGY</span>
          </div>
          <button className="new-home-sidebar-close" onClick={onClose} type="button" aria-label="Close menu">
            <X size={22} strokeWidth={2} />
          </button>
        </div>
        <div className="new-home-sidebar-scroll">
          <nav className="new-home-sidebar-content" aria-label="Main navigation">
            {filteredItemsToRender.map((item, index) => {
              if (item.type === 'item') {
                const isActive =
                  item.path === '/'
                    ? location.pathname === '/' || location.pathname === '/new-home'
                    : location.pathname === item.path;
                return (
                  <Link
                    key={index}
                    to={item.path}
                    className={`new-home-sidebar-item ${isActive ? 'active' : ''}`}
                    onClick={onClose}
                  >
                    <span className="new-home-sidebar-icon">{item.icon}</span>
                    <span>{item.label}</span>
                  </Link>
                );
              }
              if (item.type === 'group') {
                const expanded = openGroups[item.key];
                const isActive = isGroupActive(item);
                return (
                  <div key={item.key} className="new-home-sidebar-group">
                    <div
                      className={`new-home-sidebar-group-header ${expanded ? 'open' : ''} ${isActive ? 'active' : ''}`}
                      onClick={() => toggleGroup(item.key)}
                      onKeyDown={(e) => e.key === 'Enter' && toggleGroup(item.key)}
                      role="button"
                      tabIndex={0}
                      aria-expanded={expanded}
                    >
                      <span className="new-home-sidebar-icon">{item.icon}</span>
                      <span>{item.label}</span>
                      {expanded ? <ChevronDown size={18} strokeWidth={2} /> : <ChevronRight size={18} strokeWidth={2} />}
                    </div>
                    <div className={`new-home-sidebar-group-content ${expanded ? 'open' : ''}`}>
                      {item.children.map((child, childIndex) => {
                        if (child.type === 'group') {
                          const nestedExpanded = openGroups[child.key];
                          const isNestedActive =
                            child.children && child.children.some((grandchild) => location.pathname === grandchild.path);
                          return (
                            <div key={childIndex} className="new-home-sidebar-nested-group">
                              <div
                                className={`new-home-sidebar-nested-group-header ${nestedExpanded ? 'open' : ''} ${isNestedActive ? 'active' : ''}`}
                                onClick={(e) => {
                                  e.stopPropagation();
                                  toggleGroup(child.key);
                                }}
                                onKeyDown={(e) => e.key === 'Enter' && (e.stopPropagation(), toggleGroup(child.key))}
                                role="button"
                                tabIndex={0}
                                aria-expanded={nestedExpanded}
                              >
                                <span>{child.label}</span>
                                {nestedExpanded ? <ChevronDown size={14} strokeWidth={2} /> : <ChevronRight size={14} strokeWidth={2} />}
                              </div>
                              <div className={`new-home-sidebar-nested-group-content ${nestedExpanded ? 'open' : ''}`}>
                                {child.children.map((grandchild, grandchildIndex) => {
                                  const isGrandchildActive = location.pathname === grandchild.path;
                                  return (
                                    <Link
                                      key={grandchildIndex}
                                      to={grandchild.path}
                                      className={`new-home-sidebar-nested-child ${isGrandchildActive ? 'active' : ''}`}
                                      onClick={onClose}
                                    >
                                      <span>{grandchild.label}</span>
                                    </Link>
                                  );
                                })}
                              </div>
                            </div>
                          );
                        }
                        const isChildActive = child.path.includes('?')
                          ? location.pathname + location.search === child.path
                          : location.pathname === child.path;
                        return (
                          <Link
                            key={childIndex}
                            to={child.path}
                            className={`new-home-sidebar-child ${isChildActive ? 'active' : ''} ${child.isSiteLink ? 'new-home-sidebar-site-link' : ''}`}
                            onClick={onClose}
                          >
                            {child.isSiteLink && <MapPin size={14} style={{ flexShrink: 0, marginRight: '6px' }} />}
                            <span>{child.label}</span>
                          </Link>
                        );
                      })}
                    </div>
                  </div>
                );
              }
              return null;
            })}
          </nav>
        </div>
        <div className="new-home-sidebar-footer">
          <div className="new-home-sidebar-user">
            <div className="new-home-sidebar-user-avatar" aria-hidden>
              {initials}
            </div>
            <div className="new-home-sidebar-user-meta">
              <span className="new-home-sidebar-user-name">{displayName}</span>
              <span className="new-home-sidebar-user-role">{displayRole}</span>
            </div>
          </div>
        </div>
      </div>
    </>
  );
}
