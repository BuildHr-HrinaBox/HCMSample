import React, { useState, useEffect, useRef } from 'react';
import { Link, useLocation, useNavigate } from 'react-router-dom';
import './NewHome.css';
import {
  Bell, MoreVertical, X, LayoutDashboard, Scale, BookOpen,
  Target, FileCheck, Users, Map, Database, Calendar as CalendarIcon,
  FileText, Gavel, Building, Settings, Award, ClipboardList, FileSignature,
  PieChart, Calendar, FolderOpen, Upload, CalendarDays, CalendarCheck, ChevronDown, ChevronRight,
  CheckCircle, AlertCircle, Clock, Activity, Home, BarChart3
} from 'lucide-react';
import { fetchChecklistNotifications } from '../utils/checklistNotifications';
import HCMLogo from '../components/HCM Logo.png';

const NewHome = ({ userEmail, userName, userRole, sidebarOpen: sidebarOpenProp, onMenuToggle }) => {
  const [loginDisplay, setLoginDisplay] = useState(() => {
    if (typeof localStorage === 'undefined') return { name: '', email: '' };
    const storedEmail = localStorage.getItem('userEmail') || '';
    const storedName = localStorage.getItem('userName') || storedEmail;
    return { name: storedName, email: storedEmail };
  });
  const displayEmail = userEmail || loginDisplay.email || (typeof localStorage !== 'undefined' ? localStorage.getItem('userEmail') : null) || '';
  const displayName = userName || loginDisplay.name || (typeof localStorage !== 'undefined' ? localStorage.getItem('userName') : null) || displayEmail || 'User';
  const displayRole = userRole || (typeof localStorage !== 'undefined' ? localStorage.getItem('userRole') : null) || '';
  const [sidebarOpenLocal, setSidebarOpenLocal] = useState(false);
  const sidebarOpen = onMenuToggle ? sidebarOpenProp : sidebarOpenLocal;
  const setSidebarOpen = onMenuToggle ? (() => {}) : setSidebarOpenLocal;
  const toggleSidebar = onMenuToggle ? () => onMenuToggle() : () => setSidebarOpenLocal(prev => !prev);
  const [userProfile, setUserProfile] = useState(null);
  const [userDetails, setUserDetails] = useState({
    firstName: '',
    lastName: '',
    mailid: '',
    timeZone: '',
    createdTime: ''
  });
  const [activeActsCount, setActiveActsCount] = useState(0);
  const [regulationsCount, setRegulationsCount] = useState(0);
  const [isLoadingCounts, setIsLoadingCounts] = useState(true);
  const [notificationCount, setNotificationCount] = useState(0);
  const [notificationItems, setNotificationItems] = useState([]);
  const [showNotificationDropdown, setShowNotificationDropdown] = useState(false);
  const [showProfileDropdown, setShowProfileDropdown] = useState(false);
  const [windowWidth, setWindowWidth] = useState(window.innerWidth);
  const notificationDropdownRef = useRef(null);
  const profileDropdownRef = useRef(null);
  const wheelContainerRef = useRef(null);
  const [wheelSize, setWheelSize] = useState({ w: 855, h: 708 });
  const [openGroups, setOpenGroups] = useState({
    organizationMaster: false,
    library: false,
    checklist: false,
    calendar: false,
    statutoryForms: false,
    assessments: true,
    transactionStatutoryForms: false,
    formFetch: false,
    ruleBook: false
  });
  const location = useLocation();
  const navigate = useNavigate();

  // Recent activities data
  const [recentActivities] = useState([
    {
      icon: <Bell size={18} />,
      title: 'Reminder',
      description: 'Upcoming compliance deadline approaching',
      time: '1 hour ago',
      color: '#F59E0B'
    },
    {
      icon: <FileText size={18} />,
      title: 'Reports',
      description: '',
      time: '2 hours ago',
      color: '#3B82F6'
    },
    {
      icon: <CheckCircle size={18} />,
      title: 'Compliance Updated',
      description: 'Safety regulations compliance verified',
      time: '6 hours ago',
      color: '#10B981'
    },
    {
      icon: <Calendar size={18} />,
      title: 'Deadline Reminder',
      description: 'Form 12 submission due tomorrow',
      time: '1 day ago',
      color: '#EF4444'
    },
    {
      icon: <ClipboardList size={18} />,
      title: 'Checklist Updated',
      description: 'New checklist items added to compliance',
      time: '2 days ago',
      color: '#8B5CF6'
    }
  ]);

  // Navigation items with groups
  const navigationItems = [
    { type: 'item', icon: <Home size={20} />, label: 'Home', path: '/', color: 'purple' },
    { type: 'item', icon: <LayoutDashboard size={20} />, label: 'Dashboard', path: '/new-dashboard', color: 'blue' },
    {
      type: 'group',
      key: 'organizationMaster',
      icon: <Users size={20} />,
      label: 'Organization Master',
      children: [
        { label: 'Company Details', path: '/rule-book/company-details' },
        { label: 'Site Management', path: '/rule-book/site-management' },
        { label: 'Certificate Management', path: '/certificate' },
        { label: 'Employee Data', path: '/rule-book/employee-data' }
      ]
    },
    {
      type: 'group',
      key: 'library',
      icon: <FolderOpen size={20} />,
      label: 'Library',
      children: [
        { label: 'Applicable Acts Library', path: '/rule-book/act-descriptions' },
        { label: 'Acts Bulk', path: '/rule-book/acts-bulk' },
        { label: 'Quiz creation', path: '/quiz/create' }
      ]
    },
    {
      type: 'group',
      key: 'calendar',
      icon: <Calendar size={20} />,
      label: 'Calendar',
      children: [
        { label: 'Calendar Picker', path: '/calendar-picker' },
        { label: 'Calendar Master', path: '/rule-book/calendarbulk' },
        { label: 'Calendar PDF Master', path: '/rule-book/calendar-pdf-master' }
      ]
    },
    {
      type: 'group',
      key: 'checklist',
      icon: <ClipboardList size={20} />,
      label: 'Transaction',
      children: [
        { label: 'Checklist', path: '/rule-book/checklist' },
        { label: 'Checklist Master', path: '/rule-book/checklistbulk' },
        { label: 'Statutory', path: '/rule-book/statutory' },
        {
          type: 'group',
          key: 'transactionStatutoryForms',
          label: 'Statutory Forms',
          children: [
            { label: 'Factory', path: '/rule-book/factory' },
            { label: 'S&E Master', path: '/rule-book/se-master' }
          ]
        }
      ]
    },
    {
      type: 'group',
      key: 'assessments',
      icon: <PieChart size={20} />,
      label: 'Audit',
      children: [
        { label: 'Assessment Report', path: '/assessments/assessment-report' },
        { label: 'Compliance', path: '/compliance' },
        { label: 'Audit Master', path: '/rule-book/audit-master' },
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
            { label: 'Unique ID', path: '/rule-book/unique-ids' }
          ]
        }
      ]
    },
    {
      type: 'group',
      key: 'formFetch',
      icon: <FileText size={20} />,
      label: 'Form Fetch',
      children: [
        { label: 'People', path: '/people' },
        { label: 'Leave', path: '/leave' }
      ]
    },
    { type: 'item', icon: <BarChart3 size={20} />, label: 'Report', path: '/reports', color: 'blue' }
  ];

  // Module wheel cards - model UI layout: Library (top), Applicable Acts (top right), Transaction (bottom left), Audit (bottom right), Checklist/Calendar (left), Report/Dashboard (right)
  const wheelModuleCards = [
    { icon: <FolderOpen size={20} />, label: 'Library', path: '/rule-book/act-descriptions', color: 'purple', pos: { x: 0, y: -280 }, boxClass: 'information-box' },
    { icon: <FileText size={20} />, label: 'Applicable Acts', path: '/rule-book/act-descriptions', color: 'indigo', pos: { x: 386, y: -266 }, boxClass: 'applicable-acts-box' },
    { icon: <ClipboardList size={20} />, label: 'Transaction', path: '/rule-book/checklist', color: 'green', pos: { x: -242.487, y: 140 }, boxClass: 'transaction-box' },
    { icon: <PieChart size={20} />, label: 'Audit', path: '/assessments/assessment-report', color: 'green', pos: { x: 242.487, y: 140 }, boxClass: 'audit-box' },
    { icon: <ClipboardList size={20} />, label: 'Checklist', path: '/rule-book/checklist', color: 'green', pos: { x: -362, y: -28 }, boxClass: 'checklist-box' },
    { icon: <Calendar size={20} />, label: 'Calendar', path: '/calendar-picker', color: 'orange', pos: { x: -362, y: 282 }, boxClass: 'calendar-box' },
    { icon: <FileText size={20} />, label: 'Report', path: '/reports', color: 'teal', pos: { x: 410, y: -28 }, boxClass: 'report-box' },
    { icon: <LayoutDashboard size={20} />, label: 'Dashboard', path: '/new-dashboard', color: 'blue', pos: { x: 420, y: 220 }, boxClass: 'dashboard-box' }
  ];

  const toggleGroup = (key) => {
    setOpenGroups(prev => ({ ...prev, [key]: !prev[key] }));
  };

  const isGroupActive = (group) => {
    if (!group.children) return false;
    return group.children.some(child => {
      if (child.path) {
        return location.pathname === child.path;
      } else if (child.type === 'group' && child.children) {
        return child.children.some(grandchild => location.pathname === grandchild.path);
      }
      return false;
    });
  };

  // Keep loginDisplay in sync with props and profile (name, email)
  useEffect(() => {
    const storedEmail = typeof localStorage !== 'undefined' ? localStorage.getItem('userEmail') : null;
    const storedName = typeof localStorage !== 'undefined' ? localStorage.getItem('userName') : null;
    const email = (userEmail || userProfile?.email || storedEmail || '').trim();
    const name = (userName || userProfile?.name || userProfile?.email || storedName || userEmail || storedEmail || '').trim();
    if (email || name) {
      setLoginDisplay(prev => ({
        name: name || prev.name,
        email: email || prev.email
      }));
    }
  }, [userEmail, userName, userProfile?.email, userProfile?.name]);

  // Re-read localStorage on mount and after a short delay (so we show email/name as soon as set by App.js)
  useEffect(() => {
    const syncFromStorage = () => {
      if (typeof localStorage === 'undefined') return;
      const storedEmail = localStorage.getItem('userEmail') || '';
      const storedName = localStorage.getItem('userName') || storedEmail;
      if (storedEmail || storedName) {
        setLoginDisplay(prev => ({
          name: storedName || prev.name,
          email: storedEmail || prev.email
        }));
      }
    };
    syncFromStorage();
    const t = setTimeout(syncFromStorage, 400);
    const t2 = setTimeout(syncFromStorage, 1200);
    return () => { clearTimeout(t); clearTimeout(t2); };
  }, []);

  // When email is still missing, fetch from Catalyst auth / backend so profile shows logged-in mail id
  useEffect(() => {
    const currentEmail = (userEmail || loginDisplay.email || (typeof localStorage !== 'undefined' ? localStorage.getItem('userEmail') : null) || '').trim();
    if (currentEmail) return;

    let isCancelled = false;
    const extractEmailFromAuth = (auth) => {
      if (!auth || typeof auth === 'boolean') return null;
      const candidates = [
        auth.content?.email_id,
        auth.content?.email,
        auth.user?.email_id,
        auth.user?.email,
        auth.email_id,
        auth.email,
        auth.identity?.email,
        auth.accounts?.[0]?.email
      ];
      for (const v of candidates) {
        const s = (v && typeof v === 'string' && v.trim()) || null;
        if (s && s.includes('@')) return s;
      }
      return null;
    };

    const fetchEmail = async () => {
      try {
        if (typeof window !== 'undefined' && window.catalyst?.auth?.isUserAuthenticated) {
          const auth = await window.catalyst.auth.isUserAuthenticated();
          const email = extractEmailFromAuth(auth);
          if (!isCancelled && email) {
            setLoginDisplay(prev => ({ ...prev, email, name: prev.name || email }));
            try { localStorage.setItem('userEmail', email); } catch (_) {}
            return;
          }
        }
        const res = await fetch('/server/authorization_portal_function', { method: 'GET', credentials: 'include', headers: { 'Content-Type': 'application/json' } });
        if (isCancelled || !res.ok) return;
        const data = await res.json();
        const raw = data?.output ? (typeof data.output === 'string' ? JSON.parse(data.output) : data.output) : data;
        const details = raw?.user_details || raw;
        const email = (details?.email_id || details?.email || '').trim() || null;
        if (email) {
          setLoginDisplay(prev => ({ ...prev, email, name: prev.name || email }));
          try { localStorage.setItem('userEmail', email); } catch (_) {}
        }
      } catch (_) {}
    };

    const t = setTimeout(fetchEmail, 400);
    return () => { isCancelled = true; clearTimeout(t); };
  }, [userEmail, loginDisplay.email]);

  // Fetch full user profile (UserProfile model) from Catalyst auth, with fallbacks
  useEffect(() => {
    let isCancelled = false;
    const normalize = (c) => {
      if (!c || typeof c !== 'object') return null;
      const first = c.first_name ?? c.firstName ?? '';
      const last = c.last_name ?? c.lastName ?? '';
      const mail = (c.email_id ?? c.email ?? c.mailid ?? '').trim();
      const tz = c.time_zone ?? c.timeZone ?? '';
      const created = c.created_time ?? c.createdTime ?? '';
      const avatar = c.profile_picture ?? c.profilePicture ?? c.avatar ?? null;
      return { first, last, mail, tz, created, avatar };
    };

    const fetchUserProfile = async () => {
      try {
        let data = null;
        if (typeof window !== 'undefined' && window.catalyst?.auth?.isUserAuthenticated) {
          const result = await window.catalyst.auth.isUserAuthenticated();
          if (!isCancelled && result) {
            const c = result.content ?? result.user ?? result;
            data = normalize(c);
          }
        }
        if (!isCancelled && data && (data.mail || data.first || data.last)) {
          setUserDetails({
            firstName: data.first,
            lastName: data.last,
            mailid: data.mail,
            timeZone: data.tz,
            createdTime: data.created
          });
          setUserProfile({
            name: [data.first, data.last].filter(Boolean).join(' ') || data.mail || 'User',
            email: data.mail,
            avatar: data.avatar
          });
          return;
        }
        // Fallback: get email from localStorage or authorized_portal and set profile
        const stored = typeof localStorage !== 'undefined' ? localStorage.getItem('userEmail') : null;
        const email = (stored && stored.trim()) || '';
        if (isCancelled) return;
        if (email) {
          setUserDetails(prev => (prev.mailid ? prev : { ...prev, mailid: email }));
          setUserProfile(prev => (prev?.email ? prev : { name: email, email, avatar: null }));
          setLoginDisplay(prev => ({ name: prev.name || email, email }));
        } else {
          const res = await fetch('/server/authorization_portal_function', { method: 'GET', credentials: 'include', headers: { 'Content-Type': 'application/json' } });
          if (isCancelled || !res.ok) return;
          const json = await res.json();
          const raw = json?.output ? (typeof json.output === 'string' ? JSON.parse(json.output) : json.output) : json;
          const details = raw?.user_details ?? raw;
          const d = normalize(details || {});
          const fallbackEmail = (d.mail || details?.email_id || details?.email || '').trim();
          if (fallbackEmail) {
            setUserDetails(prev => (prev.mailid ? prev : {
              firstName: d.first,
              lastName: d.last,
              mailid: fallbackEmail,
              timeZone: d.tz,
              createdTime: d.created
            }));
            setUserProfile(prev => (prev?.email ? prev : { name: d.first && d.last ? [d.first, d.last].join(' ') : fallbackEmail, email: fallbackEmail, avatar: d.avatar }));
            setLoginDisplay(prev => ({ name: prev.name || fallbackEmail, email: fallbackEmail }));
          }
        }
      } catch (_) {}
    };
    fetchUserProfile();
    return () => { isCancelled = true; };
  }, []);

  // Sync userDetails and userProfile from props / localStorage so profile displays for both App User and App Administrator
  useEffect(() => {
    const email = (userEmail || loginDisplay.email || (typeof localStorage !== 'undefined' ? localStorage.getItem('userEmail') : null) || '').trim();
    const name = (userName || loginDisplay.name || (typeof localStorage !== 'undefined' ? localStorage.getItem('userName') : null) || '').trim();
    if (!email && !name) return;
    setUserDetails(prev => {
      if (prev.mailid && prev.firstName) return prev;
      const parts = name ? name.split(/\s+/) : [];
      return {
        ...prev,
        mailid: prev.mailid || email,
        firstName: prev.firstName || (parts[0] || ''),
        lastName: prev.lastName || (parts.slice(1).join(' ') || ''),
      };
    });
    setUserProfile(prev => {
      if (prev?.email && prev?.name) return prev;
      return { name: name || prev?.name || email, email: email || prev?.email, avatar: prev?.avatar ?? null };
    });
  }, [userEmail, userName, loginDisplay.email, loginDisplay.name]);

  // Close profile dropdown when clicking outside
  useEffect(() => {
    const handleClickOutside = (e) => {
      if (profileDropdownRef.current && !profileDropdownRef.current.contains(e.target)) {
        setShowProfileDropdown(false);
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  // Fetch counts (Active Acts from actsbulk)
  useEffect(() => {
    let isCancelled = false;
    const fetchCounts = async () => {
      setIsLoadingCounts(true);
      try {
        // Try a lightweight count endpoint first (if available)
        const countResp = await fetch('/server/actsbulk_function/acts/count');
        if (!isCancelled && countResp.ok) {
          const data = await countResp.json();
          const count = typeof data?.count === 'number' ? data.count : 0;
          setActiveActsCount(count);
          setIsLoadingCounts(false);
          return;
        }
      } catch (_) {
        // Ignore and fallback to getAll
      }
      try {
        const resp = await fetch('/server/actsbulk_function/actsbulk?action=getAll');
        if (!isCancelled && resp.ok) {
          const data = await resp.json();
          const list = Array.isArray(data?.data) ? data.data : Array.isArray(data) ? data : [];
          setActiveActsCount(list.length || (typeof data?.count === 'number' ? data.count : 0));
        }
      } catch (_) {
        // Keep default 0 on failure
      } finally {
        if (!isCancelled) setIsLoadingCounts(false);
      }
    };
    fetchCounts();
    return () => { isCancelled = true; };
  }, []);

  // Fetch regulations count (from regulationsbulk)
  useEffect(() => {
    let isCancelled = false;
    const fetchRegulationsCount = async () => {
      setIsLoadingCounts(true);
      try {
        const countResp = await fetch('/server/regulationsbulk_function/regulations/count');
        if (!isCancelled && countResp.ok) {
          const data = await countResp.json();
          const count = typeof data?.count === 'number' ? data.count : 0;
          setRegulationsCount(count);
          setIsLoadingCounts(false);
          return;
        }
      } catch (_) {
        // Ignore and fallback
      }
      try {
        const resp = await fetch('/server/regulationsbulk_function/regulationsbulk?action=getAll');
        if (!isCancelled && resp.ok) {
          const data = await resp.json();
          const list = Array.isArray(data?.data) ? data.data : Array.isArray(data) ? data : [];
          setRegulationsCount(list.length || (typeof data?.count === 'number' ? data.count : 0));
        }
      } catch (_) {
        // Keep default 0 on failure
      } finally {
        if (!isCancelled) setIsLoadingCounts(false);
      }
    };
    fetchRegulationsCount();
    return () => { isCancelled = true; };
  }, []);

  // Fetch checklist notifications (filtered by logged-in user so only corresponding statutory forms show)
  useEffect(() => {
    const fetchNotifications = async () => {
      try {
        const email = displayEmail || (typeof localStorage !== 'undefined' ? localStorage.getItem('userEmail') : null) || '';
        const result = await fetchChecklistNotifications(email);
        setNotificationCount(result.count);
        setNotificationItems(result.items);
      } catch (error) {
        console.error('Error fetching notifications:', error);
        setNotificationCount(0);
        setNotificationItems([]);
      }
    };

    fetchNotifications();
    const interval = setInterval(fetchNotifications, 60000);
    return () => clearInterval(interval);
  }, [displayEmail]);

  // Close notification dropdown when clicking outside
  useEffect(() => {
    const handleClickOutside = (event) => {
      if (notificationDropdownRef.current && !notificationDropdownRef.current.contains(event.target)) {
        setShowNotificationDropdown(false);
      }
    };

    document.addEventListener('mousedown', handleClickOutside);
    return () => {
      document.removeEventListener('mousedown', handleClickOutside);
    };
  }, []);

  // Handle window resize for responsive wheel layout
  useEffect(() => {
    const handleResize = () => {
      setWindowWidth(window.innerWidth);
    };

    window.addEventListener('resize', handleResize);
    return () => window.removeEventListener('resize', handleResize);
  }, []);

  // Measure wheel container so Audit→Report/Dashboard arrows align with card positions
  useEffect(() => {
    const el = wheelContainerRef.current;
    if (!el) return;
    const ro = new ResizeObserver(() => {
      if (el) setWheelSize({ w: el.offsetWidth, h: el.offsetHeight });
    });
    ro.observe(el);
    setWheelSize({ w: el.offsetWidth, h: el.offsetHeight });
    return () => ro.disconnect();
  }, []);

  // When on Home: show NewHome sidebar. When navigating away, close it (main app sidebar will show).
  useEffect(() => {
    if (location.pathname === '/' || location.pathname === '/new-home') {
      setSidebarOpen(true);
    } else {
      setSidebarOpen(false);
    }
  }, [location.pathname]);

  const handleLogout = () => {
    const redirectURL = '/__catalyst/auth/login';
    if (window.catalyst?.auth?.signOut) {
      window.catalyst.auth.signOut(redirectURL);
    } else {
      window.location.href = redirectURL;
    }
  };

  return (
    <div style={{ display: 'flex' }}>
      {/* Sidebar - only when not using app-level sidebar (e.g. standalone NewHome) */}
      {!onMenuToggle && (
        <>
          {sidebarOpen && (
            <div className="new-home-sidebar-overlay" onClick={toggleSidebar} aria-hidden="true" />
          )}
          <div className={`new-home-sidebar ${sidebarOpen ? 'open' : ''}`}>
            <div className="new-home-sidebar-header">
              <div className="new-home-sidebar-logo-container">
                <span className="new-home-sidebar-logo" style={{ fontSize: '1.25rem', fontWeight: 700, color: '#1f2937' }}>HCM</span>
              </div>
              <button className="new-home-sidebar-close" onClick={toggleSidebar} type="button" aria-label="Close menu">
                <X size={24} />
              </button>
            </div>
            <div className="new-home-sidebar-content">
              {navigationItems.map((item, index) => {
                if (item.type === 'item') {
                  const isActive = item.path === '/'
                    ? (location.pathname === '/' || location.pathname === '/new-home')
                    : location.pathname === item.path;
                  return (
                    <Link
                      key={index}
                      to={item.path}
                      className={`new-home-sidebar-item ${isActive ? 'active' : ''}`}
                      onClick={() => setSidebarOpen(false)}
                    >
                      <span className="new-home-sidebar-icon">{item.icon}</span>
                      <span>{item.label}</span>
                    </Link>
                  );
                } else if (item.type === 'group') {
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
                        {expanded ? <ChevronDown size={16} /> : <ChevronRight size={16} />}
                      </div>
                      <div className={`new-home-sidebar-group-content ${expanded ? 'open' : ''}`}>
                        {item.children.map((child, childIndex) => {
                          if (child.type === 'group') {
                            const nestedExpanded = openGroups[child.key];
                            const isNestedActive = child.children && child.children.some(grandchild => location.pathname === grandchild.path);
                            return (
                              <div key={childIndex} className="new-home-sidebar-nested-group">
                                <div
                                  className={`new-home-sidebar-nested-group-header ${nestedExpanded ? 'open' : ''} ${isNestedActive ? 'active' : ''}`}
                                  onClick={(e) => { e.stopPropagation(); toggleGroup(child.key); }}
                                  onKeyDown={(e) => e.key === 'Enter' && (e.stopPropagation(), toggleGroup(child.key))}
                                  role="button"
                                  tabIndex={0}
                                  aria-expanded={nestedExpanded}
                                >
                                  <span>{child.label}</span>
                                  {nestedExpanded ? <ChevronDown size={14} /> : <ChevronRight size={14} />}
                                </div>
                                <div className={`new-home-sidebar-nested-group-content ${nestedExpanded ? 'open' : ''}`}>
                                  {child.children.map((grandchild, grandchildIndex) => {
                                    const isGrandchildActive = location.pathname === grandchild.path;
                                    return (
                                      <Link
                                        key={grandchildIndex}
                                        to={grandchild.path}
                                        className={`new-home-sidebar-nested-child ${isGrandchildActive ? 'active' : ''}`}
                                        onClick={() => setSidebarOpen(false)}
                                      >
                                        <span>{grandchild.label}</span>
                                      </Link>
                                    );
                                  })}
                                </div>
                              </div>
                            );
                          }
                          const isChildActive = location.pathname === child.path;
                          return (
                            <Link
                              key={childIndex}
                              to={child.path}
                              className={`new-home-sidebar-child ${isChildActive ? 'active' : ''}`}
                              onClick={() => setSidebarOpen(false)}
                            >
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
            </div>
          </div>
        </>
      )}

      {/* Main content area - model UI background */}
      <div className="new-home-main-area" style={{ marginLeft: 0, flex: '1 1 0%', transition: 'margin-left 0.2s', minHeight: '100vh' }}>
        <div style={{ padding: 0 }}>
          <div className="new-home-container">
            {/* Header Section - model: left (menu+logo+subtitle), center (HCM), right (bell+profile) */}
            <div className="new-home-header">
              <div className="new-home-header-top">
                <div className="new-home-header-left">
                  <button className="new-home-menu-btn" onClick={toggleSidebar} type="button" aria-label="Menu">
                    <MoreVertical size={24} />
                  </button>
                  <div className="new-home-logo-container">
                    <img src={HCMLogo} alt="HCM" className="new-home-logo" />
                  </div>
                </div>
                <h1 className="new-home-welcome">HCM</h1>
                <div className="new-home-header-right">
                  <div ref={notificationDropdownRef} style={{ position: 'relative' }}>
            <button
              className="new-home-bell-btn"
              style={{ position: 'relative' }}
              onClick={() => setShowNotificationDropdown(!showNotificationDropdown)}
            >
              <Bell size={24} />
              {notificationCount > 0 && (
                <span
                  style={{
                    position: 'absolute',
                    top: '-4px',
                    right: '-4px',
                    backgroundColor: '#ef4444',
                    color: 'white',
                    borderRadius: '50%',
                    width: '20px',
                    height: '20px',
                    display: 'flex',
                    alignItems: 'center',
                    justifyContent: 'center',
                    fontSize: '11px',
                    fontWeight: 'bold',
                    border: '2px solid white'
                  }}
                >
                  {notificationCount > 99 ? '99+' : notificationCount}
                </span>
              )}
            </button>

            {/* Notification Dropdown - aligned right under bell, no horizontal scroll */}
            {showNotificationDropdown && (
              <div
                style={{
                  position: 'absolute',
                  top: 'calc(100% + 10px)',
                  right: 0,
                  width: '320px',
                  backgroundColor: 'white',
                  borderRadius: '12px',
                  boxShadow: '0 10px 30px rgba(0, 0, 0, 0.2)',
                  maxHeight: '500px',
                  overflowX: 'hidden',
                  overflowY: 'auto',
                  zIndex: 1000,
                  border: '1px solid #e5e7eb'
                }}
              >
                <div
                  style={{
                    padding: '12px 16px',
                    borderBottom: '1px solid #e5e7eb',
                    display: 'flex',
                    justifyContent: 'space-between',
                    alignItems: 'center',
                    backgroundColor: '#f9fafb',
                    borderTopLeftRadius: '12px',
                    borderTopRightRadius: '12px',
                    flexShrink: 0
                  }}
                >
                  <h3 style={{ margin: 0, fontSize: '16px', fontWeight: '600', color: '#1f2937' }}>
                    Notifications ({notificationCount})
                  </h3>
                  <button
                    onClick={() => setShowNotificationDropdown(false)}
                    style={{
                      background: 'none',
                      border: 'none',
                      cursor: 'pointer',
                      fontSize: '18px',
                      color: '#6b7280',
                      padding: '4px 8px',
                      flexShrink: 0
                    }}
                  >
                    ×
                  </button>
                </div>
                <div style={{ padding: '8px', overflowX: 'hidden' }}>
                  {notificationItems.length === 0 ? (
                    <div
                      style={{
                        padding: '24px',
                        textAlign: 'center',
                        color: '#6b7280',
                        fontSize: '14px'
                      }}
                    >
                      No notifications
                    </div>
                  ) : (
                    notificationItems.map((item, index) => (
                      <div
                        key={item.id || index}
                        style={{
                          padding: '12px',
                          borderBottom: index < notificationItems.length - 1 ? '1px solid #e5e7eb' : 'none',
                          cursor: 'pointer',
                          transition: 'background-color 0.2s',
                          borderRadius: '8px',
                          margin: '4px 0',
                          overflow: 'hidden'
                        }}
                        onMouseEnter={(e) => { e.currentTarget.style.backgroundColor = '#f3f4f6'; }}
                        onMouseLeave={(e) => { e.currentTarget.style.backgroundColor = 'transparent'; }}
                        onClick={() => {
                          setShowNotificationDropdown(false);
                          const path = item.source === 'statutory' ? '/rule-book/statutory' : '/rule-book/checklist';
                          navigate(path);
                        }}
                      >
                        <div style={{ display: 'flex', alignItems: 'flex-start', gap: '12px', minWidth: 0 }}>
                          <div style={{ flex: 1, minWidth: 0 }}>
                            <div style={{ display: 'flex', alignItems: 'center', gap: '8px', marginBottom: '4px', flexWrap: 'wrap' }}>
                              <span style={{ fontWeight: '600', color: '#1f2937' }}>
                                {item.formName || 'Unknown Form'}
                              </span>
                              {item.daysRemaining !== null && (
                                <span
                                  style={{
                                    backgroundColor: item.daysRemaining <= 1 ? '#fee2e2' : '#dbeafe',
                                    color: item.daysRemaining <= 1 ? '#dc2626' : '#2563eb',
                                    padding: '2px 8px',
                                    borderRadius: '6px',
                                    fontSize: '12px',
                                    fontWeight: '600',
                                    whiteSpace: 'nowrap',
                                    flexShrink: 0
                                  }}
                                >
                                  {item.daysRemaining === 0 ? 'Today' : item.daysRemaining === 1 ? 'Tomorrow' : `${item.daysRemaining} days`}
                                </span>
                              )}
                            </div>
                            <div style={{ fontSize: '13px', color: '#6b7280', marginBottom: '0' }}>
                              {item.message}
                            </div>
                            {item.description && (
                              <div style={{ fontSize: '12px', color: '#9ca3af', marginTop: '4px', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                                {item.description}
                              </div>
                            )}
                          </div>
                        </div>
                      </div>
                    ))
                  )}
                </div>
              </div>
            )}
                  </div>
                  {/* Profile: UserProfile model - clickable avatar + dropdown */}
                  <div ref={profileDropdownRef} style={{ position: 'relative' }}>
                    <div
                      role="button"
                      tabIndex={0}
                      onClick={() => setShowProfileDropdown(!showProfileDropdown)}
                      onKeyDown={(e) => e.key === 'Enter' && setShowProfileDropdown(prev => !prev)}
                      className="new-home-login-user-display new-home-profile-email"
                      style={{
                        display: 'flex',
                        alignItems: 'center',
                        gap: '10px',
                        color: 'rgba(255,255,255,0.95)',
                        fontSize: '0.875rem',
                        fontWeight: 500,
                        cursor: 'pointer',
                        minWidth: '120px',
                        maxWidth: '320px',
                        padding: '4px 8px',
                        borderRadius: '8px'
                      }}
                    >
                      {userProfile?.avatar ? (
                        <img
                          src={userProfile.avatar}
                          alt="Profile"
                          style={{ width: 36, height: 36, borderRadius: '50%', objectFit: 'cover', flexShrink: 0 }}
                        />
                      ) : (
                        <div
                          style={{
                            width: 36,
                            height: 36,
                            borderRadius: '50%',
                            background: 'rgba(255,255,255,0.25)',
                            display: 'flex',
                            alignItems: 'center',
                            justifyContent: 'center',
                            fontSize: '1rem',
                            fontWeight: 600,
                            flexShrink: 0
                          }}
                        >
                          {(userDetails.firstName?.charAt(0) || userProfile?.name?.charAt(0) || displayName?.charAt(0) || loginDisplay.name?.charAt(0) || displayEmail?.charAt(0) || 'U').toUpperCase()}
                        </div>
                      )}
                      <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'flex-end', minWidth: 0, flex: 1, overflow: 'hidden' }}>
                        <span style={{ fontSize: '0.7rem', opacity: 0.9, whiteSpace: 'nowrap' }}>Logged in as</span>
                        <span
                          title={userDetails.mailid || loginDisplay.email || displayEmail || ''}
                          style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', maxWidth: '100%' }}
                        >
                          {displayName || displayEmail || 'User'}
                        </span>
                      </div>
                    </div>
                    {showProfileDropdown && (
                      <div
                        className="new-home-profile-dropdown"
                        style={{
                          position: 'absolute',
                          top: 'calc(100% + 10px)',
                          right: 0,
                          backgroundColor: 'white',
                          borderRadius: '12px',
                          boxShadow: '0 10px 30px rgba(0, 0, 0, 0.2)',
                          minWidth: '320px',
                          maxWidth: '380px',
                          zIndex: 1000,
                          border: '1px solid #e5e7eb',
                          overflow: 'hidden'
                        }}
                      >
                        <div style={{ padding: '20px', textAlign: 'center', borderBottom: '1px solid #e5e7eb', backgroundColor: '#f9fafb' }}>
                          <h3 style={{ margin: '0 0 12px', fontSize: '16px', fontWeight: 600, color: '#1f2937' }}>User Profile Information</h3>
                          {userProfile?.avatar ? (
                            <img src={userProfile.avatar} alt="Profile" style={{ width: 80, height: 80, borderRadius: '50%', objectFit: 'cover' }} />
                          ) : (
                            <div style={{ width: 80, height: 80, borderRadius: '50%', background: '#e5e7eb', display: 'inline-flex', alignItems: 'center', justifyContent: 'center', fontSize: '2rem', fontWeight: 600, color: '#6b7280' }}>
                              {(userDetails.firstName?.charAt(0) || userProfile?.name?.charAt(0) || displayName?.charAt(0) || loginDisplay.email?.charAt(0) || displayEmail?.charAt(0) || 'U').toUpperCase()}
                            </div>
                          )}
                        </div>
                        <div style={{ padding: '16px' }}>
                          {(() => {
                            const profileEmail = userDetails.mailid || loginDisplay.email || displayEmail || (typeof localStorage !== 'undefined' ? localStorage.getItem('userEmail') : null) || '';
                            const profileName = userDetails.firstName || userDetails.lastName ? [userDetails.firstName, userDetails.lastName].filter(Boolean).join(' ') : (loginDisplay.name || profileEmail || '');
                            return (
                              <>
                                <p style={{ margin: '8px 0', fontSize: '14px', color: '#374151' }}><strong>First Name:</strong> {userDetails.firstName || (profileName && !profileName.includes('@') ? profileName.split(/\s+/)[0] : null) || '—'}</p>
                                <p style={{ margin: '8px 0', fontSize: '14px', color: '#374151' }}><strong>Last Name:</strong> {userDetails.lastName || (profileName && profileName.includes(' ') ? profileName.split(/\s+/).slice(1).join(' ') : null) || '—'}</p>
                                <p style={{ margin: '8px 0', fontSize: '14px', color: '#374151' }}><strong>Email Address:</strong> {profileEmail || '—'}</p>
                                <p style={{ margin: '8px 0', fontSize: '14px', color: '#374151' }}><strong>Role:</strong> {displayRole || userDetails.role_identifier || userDetails.role || '—'}</p>
                                <p style={{ margin: '8px 0', fontSize: '14px', color: '#374151' }}><strong>Time Zone:</strong> {userDetails.timeZone || '—'}</p>
                                <p style={{ margin: '8px 0', fontSize: '14px', color: '#374151' }}><strong>Joined On:</strong> {userDetails.createdTime || '—'}</p>
                              </>
                            );
                          })()}
                        </div>
                        <div style={{ padding: '12px 16px', borderTop: '1px solid #e5e7eb' }}>
                          <button
                            type="button"
                            onClick={handleLogout}
                            style={{
                              width: '100%',
                              padding: '10px 16px',
                              borderRadius: '8px',
                              border: 'none',
                              background: '#7c3aed',
                              color: 'white',
                              fontSize: '14px',
                              fontWeight: 600,
                              cursor: 'pointer',
                              display: 'flex',
                              alignItems: 'center',
                              justifyContent: 'center',
                              gap: '8px'
                            }}
                          >
                            <span>→</span>
                            <span>Logout</span>
                          </button>
                        </div>
                      </div>
                    )}
                  </div>
                </div>
              </div>

              {/* Stats Cards - single row below header */}
              <div className="new-home-stats">
          <div className="new-home-stat-card">
            <div className="new-home-stat-label">Active Acts</div>
            <div className="new-home-stat-value">{isLoadingCounts ? '...' : activeActsCount}</div>
          </div>
          <div className="new-home-stat-card">
            <div className="new-home-stat-label">Regulation</div>
            <div className="new-home-stat-value">{isLoadingCounts ? '...' : regulationsCount}</div>
          </div>
          <div className="new-home-stat-card">
            <div className="new-home-stat-label">Form Submitted</div>
            <div className="new-home-stat-value">42</div>
          </div>
          <div className="new-home-stat-card">
            <div className="new-home-stat-label">Compliance Rate</div>
            <div className="new-home-stat-value">92%</div>
          </div>
            </div>
          </div>

            {/* Main Content - About full width, then Diagram | Recent Activity */}
            <div className="new-home-content">
              <div className="new-home-content-wrapper">
                {/* About HR Compliance Management - full width row */}
                <div className="new-home-about-section new-home-about-fullwidth">
                  <div className="new-home-about-card">
                    <h3 className="new-home-about-title">About HR Compliance Management</h3>
                    <p className="new-home-about-subtitle">HR Compliance Management - Your Complete Compliance Solution</p>
                    <div className="new-home-about-description">
                      <p>
                        HR Compliance Management is a comprehensive platform designed to streamline
                        and manage all aspects of statutory compliance, regulatory requirements, and legal obligations.
                      </p>
                      <p>
                        Built with modern technology and user-friendly interfaces, HR Compliance Management provides powerful tools for
                        administrators and users to efficiently manage their compliance processes.
                      </p>
                    </div>
                  </div>
                </div>
                <div className="new-home-main-layout">
                  {/* Left column: Diagram (System Modules) */}
                  <div className="new-home-left-content">
                    <div className="new-home-modules-section">
                  <div className="new-home-modules-wheel-container">
                    {/* Triangle Cards */}
                <div className="new-home-modules-wheel" ref={wheelContainerRef}>
                  {/* Model UI: SVG with triangle, gradient and dashed arrow lines (viewBox from reference) */}
                  <svg
                    className="new-home-modules-connections"
                    viewBox="0 0 855 708"
                    preserveAspectRatio="none"
                    style={{
                      position: 'absolute',
                      top: 0,
                      left: 0,
                      width: '100%',
                      height: '100%',
                      pointerEvents: 'none',
                      zIndex: 0,
                      overflow: 'visible'
                    }}
                  >
                    <defs>
                      <marker id="arrowhead" markerWidth="10" markerHeight="10" refX="9" refY="3" orient="auto" markerUnits="userSpaceOnUse" overflow="visible" viewBox="0 0 10 6">
                        <polygon points="0 0, 10 3, 0 6" fill="#7C3AED" fillOpacity="0.9" />
                      </marker>
                      <linearGradient id="systemTriangleGradient" x1="0%" y1="0%" x2="0%" y2="100%">
                        <stop offset="0%" stopColor="#0000FF" stopOpacity="0.9" />
                        <stop offset="100%" stopColor="#00FFFF" stopOpacity="0.85" />
                      </linearGradient>
                    </defs>
                    <polygon
                      points="427.5,172 309.5,436 545.5,436"
                      fill="url(#systemTriangleGradient)"
                      className="new-home-triangle-fill"
                      fillOpacity="0.8"
                    />
                    <polygon
                      points="427.5,172 309.5,436 545.5,436"
                      fill="none"
                      stroke="#8B4513"
                      strokeWidth="3"
                      className="new-home-triangle-animated-border"
                    />
                    {/* Arrow from Library to Applicable Acts - dashed line + explicit arrowhead so it shows when sidebar is removed */}
                    <line x1="462" y1="129" x2="604" y2="129" stroke="#7C3AED" strokeWidth="1.5" strokeDasharray="5,5" opacity="0.6" className="new-home-connection-line" />
                    <polygon points="604 129, 594 132, 594 126" fill="#7C3AED" fillOpacity="0.9" className="new-home-connection-line" />
                  </svg>
                  {/* Audit → Report & Dashboard arrows in container pixel space so they always hit the cards */}
                  {wheelSize.w > 0 && wheelSize.h > 0 && (() => {
                    const cx = wheelSize.w / 2;
                    const cy = wheelSize.h / 2;
                    const audit = { x: cx + 242.487, y: cy + 140 };
                    const report = { x: cx + 410, y: cy - 28 };
                    const dashboard = { x: cx + 420, y: cy + 220 };
                    const transaction = { x: cx - 242.487, y: cy + 140 };
                    const checklist = { x: cx - 362, y: cy - 28 };
                    const calendar = { x: cx - 362, y: cy + 282 };
                    const arrow = (from, to) => {
                      const dx = to.x - from.x, dy = to.y - from.y;
                      const len = Math.sqrt(dx * dx + dy * dy) || 1;
                      const ux = dx / len, uy = dy / len;
                      const bx = to.x - 10 * ux, by = to.y - 10 * uy;
                      const px = bx + 3 * uy; const mx = bx - 3 * uy;
                      const py = by - 3 * ux; const my = by + 3 * ux;
                      return { line: { x1: from.x, y1: from.y, x2: to.x, y2: to.y }, poly: `${to.x} ${to.y}, ${px} ${py}, ${mx} ${my}` };
                    };
                    const toReport = arrow(audit, report);
                    const toDashboard = arrow(audit, dashboard);
                    const toChecklist = arrow(transaction, checklist);
                    const toCalendar = arrow(transaction, calendar);
                    return (
                      <svg
                        className="new-home-modules-connections"
                        viewBox={`0 0 ${wheelSize.w} ${wheelSize.h}`}
                        preserveAspectRatio="none"
                        style={{ position: 'absolute', top: 0, left: 0, width: '100%', height: '100%', pointerEvents: 'none', zIndex: 0, overflow: 'visible' }}
                      >
                        {/* Audit V-shape: to Report and Dashboard */}
                        <line {...toReport.line} stroke="#7C3AED" strokeWidth="1.5" strokeDasharray="5,5" opacity={0.6} className="new-home-connection-line" />
                        <line {...toDashboard.line} stroke="#7C3AED" strokeWidth="1.5" strokeDasharray="5,5" opacity={0.6} className="new-home-connection-line" />
                        <polygon points={toReport.poly} fill="#7C3AED" fillOpacity={0.9} className="new-home-connection-line" />
                        <polygon points={toDashboard.poly} fill="#7C3AED" fillOpacity={0.9} className="new-home-connection-line" />
                        {/* Transaction V-shape: to Checklist and Calendar */}
                        <line {...toChecklist.line} stroke="#7C3AED" strokeWidth="1.5" strokeDasharray="5,5" opacity={0.6} className="new-home-connection-line" />
                        <line {...toCalendar.line} stroke="#7C3AED" strokeWidth="1.5" strokeDasharray="5,5" opacity={0.6} className="new-home-connection-line" />
                        <polygon points={toChecklist.poly} fill="#7C3AED" fillOpacity={0.9} className="new-home-connection-line" />
                        <polygon points={toCalendar.poly} fill="#7C3AED" fillOpacity={0.9} className="new-home-connection-line" />
                      </svg>
                    );
                  })()}

                  {/* Center title - Model UI */}
                  <div className="new-home-modules-center-title">
                    <h2 className="new-home-section-title">System Modules & Features</h2>
                  </div>

                  {/* Module cards - model UI positions */}
                  {wheelModuleCards.map((item, index) => {
                    const isActive = location.pathname === item.path;
                    const { x, y } = item.pos;
                    return (
                      <Link
                        key={index}
                        to={item.path}
                        className={`new-home-module-card-wheel claymorphism-card ${item.boxClass || ''} ${isActive ? 'active' : ''}`}
                        style={{
                          position: 'absolute',
                          left: '50%',
                          top: '50%',
                          transform: `translate(calc(-50% + ${x}px), calc(-50% + ${y}px))`,
                          ['--hover-x']: `${x}px`,
                          ['--hover-y']: `${y}px`
                        }}
                      >
                        <div className={`new-home-module-icon-wheel ${item.color || 'purple'}`}>
                          {item.icon}
                        </div>
                        <div className="new-home-module-label-wheel">{item.label}</div>
                      </Link>
                    );
                  })}
                  </div>
                  </div>
                </div>
              </div>
                  {/* Right column: Recent Activity */}
                  <div className="new-home-right-content">
                    <div className="new-home-activity-section">
                      <div className="new-home-activity-card">
                        <div className="new-home-activity-header">
                          <Activity size={22} />
                          <h3 className="new-home-activity-title">Recent Activity</h3>
                        </div>
                        <div className="new-home-activity-list">
                          {recentActivities.map((activity, index) => (
                            <div
                              key={index}
                              className="new-home-activity-item"
                              onClick={() => {
                                if (activity.title === 'Reminder' || activity.title === 'Deadline Reminder') {
                                  setShowNotificationDropdown(true);
                                }
                              }}
                              style={{
                                cursor: (activity.title === 'Reminder' || activity.title === 'Deadline Reminder') ? 'pointer' : 'default'
                              }}
                            >
                              <div
                                className="new-home-activity-icon-wrapper"
                                style={{ backgroundColor: `${activity.color}20`, color: activity.color }}
                              >
                                {activity.icon}
                              </div>
                              <div className="new-home-activity-content">
                                <h4 className="new-home-activity-item-title">{activity.title}</h4>
                                <p className="new-home-activity-item-description">{activity.description}</p>
                                <span className="new-home-activity-time">
                                  <Clock size={12} />
                                  {activity.time}
                                </span>
                              </div>
                            </div>
                          ))}
                        </div>
                      </div>
                    </div>
                  </div>
                </div>
            </div>
          </div>
        </div>
      </div>
    </div>
    </div>
  );
};

export default NewHome;