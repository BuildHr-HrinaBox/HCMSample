import React, { useState, useEffect, useCallback } from 'react';
import { HashRouter as Router, Routes, Route, Navigate, useLocation } from 'react-router-dom';
import './App.css';
import NewHomeSidebar from './components/NewHomeSidebar';
import Header from './components/Header';
import LoginPage from './Pages/LoginPage';
import SimpleLogin from './Pages/SimpleLogin';
import Dashboard from './Pages/Dashboard';
import NewDashboard from './Pages/newdashboard';
import HcmDashboardPageShell from './components/HcmDashboardPageShell';
import Acts from './Pages/Acts';
import ActDescription from './Pages/ActDescription';
import ScoringCriteria from './Pages/ScoringCriteria';
import RecordCategory from './Pages/RecordCategory';
import ClientDetails from './Pages/ClientDetails';
import Compliance from './Pages/Compliance';
import States from './Pages/States';
import ActsBulk from './Pages/ActsBulk';
import CalendarBulk from './Pages/CalendarBulk';
import CalendarPicker from './Pages/CalendarPicker';
import CalendarPDFMaster from './Pages/CalendarPDFMaster';
import RegulationsBulk from './Pages/RegulationsBulk';
import Regulations from './Pages/Regulations';
import CompanyDetails from './Pages/CompanyDetails';
import SiteManagement from './Pages/SiteManagement';
import Unique from './Pages/Unique';
import UniqueId from './Pages/UniqueId';
import AssessmentReport from './Pages/AssessmentReport';
import Form12 from './Pages/Form12';
import Form15I from './Pages/Form15I';
import Form15II from './Pages/Form15II';
import Form25 from './Pages/Form25';
import Checklist from './Pages/Checklist';
import Checklistbulk from './Pages/Checklistbulk';
import ToDoList from './Pages/ToDoList';
import StatutoryMaster from './Pages/StatutoryMaster';
import Statutorymasterfactory from './Pages/Statutorymasterfactory';
import Factory from './Pages/Factory';
import FormBuilder from './Pages/FormBuilder';
import Statutory from './Pages/Statutory';
import NewHome from './Pages/NewHome';
import Certificate from './Pages/Certificate';
import SEMaster from './Pages/SEMaster';
import Formmaster from './Pages/Formmaster';
import People from './Pages/People';
import Attendance from './Pages/Attendance';
import Leave from './Pages/Leave';
import NewLeave from './Pages/NewLeave';
import ApprovedLeaves from './Pages/ApprovedLeaves';
import Payroll from './Pages/Payroll';
import SamplePayroll from './Pages/SamplePayroll';
import CLRA from './utils/CLRA';
import Reports from './Pages/Reports';
import Mainreport from './Pages/Mainreport';
import ReturnedReport from './Pages/ReturnedReport';
import Settings from './Pages/Settings';
import Setup from './Pages/Setup';
import AuditMasterImport from './Pages/AuditMasterImport';
import Audit from './Pages/Audit';
import AuditReport from './Pages/AuditReport';

// Protected Route component
function ProtectedRoute({ children }) {
  const [isAuthenticated, setIsAuthenticated] = useState(false);
  const [isLoading, setIsLoading] = useState(true);

  useEffect(() => {
    const checkAuth = async () => {
      try {
        // Check if Catalyst is available
        if (window.catalyst && window.catalyst.auth && window.catalyst.auth.isUserAuthenticated) {
          const authStatus = await window.catalyst.auth.isUserAuthenticated();
          setIsAuthenticated(authStatus);
        } else {
          // If Catalyst is not available, assume not authenticated
          setIsAuthenticated(false);
        }
      } catch (err) {
        console.log('Authentication check failed:', err);
        setIsAuthenticated(false);
      } finally {
        setIsLoading(false);
      }
    };
    checkAuth();
  }, []);

  const logout = useCallback(() => {
    const redirectURL = "/__catalyst/auth/login";
    if (window.catalyst && window.catalyst.auth && window.catalyst.auth.signOut) {
      window.catalyst.auth.signOut(redirectURL);
    } else {
      window.location.href = redirectURL;
    }
  }, []);

  if (isLoading) {
    return (
      <div style={{ 
        display: 'flex', 
        justifyContent: 'center', 
        alignItems: 'center', 
        height: '100vh',
        fontSize: '1.2rem',
        color: '#3f51b5'
      }}>
        Loading...
      </div>
    );
  }

  // Skip login form: always show main app (no login page)
  return children;
}

// Role-based Protected Route component
function RoleProtectedRoute({ children, allowedRoles, userRole }) {
  if (!userRole) {
    return (
      <div style={{ 
        display: 'flex', 
        justifyContent: 'center', 
        alignItems: 'center', 
        height: '100vh',
        fontSize: '1.2rem',
        color: '#f44336'
      }}>
        Loading user role...
      </div>
    );
  }

  if (!allowedRoles.includes(userRole)) {
    return (
      <div style={{ 
        display: 'flex', 
        flexDirection: 'column',
        justifyContent: 'center', 
        alignItems: 'center', 
        height: '100vh',
        fontSize: '1.2rem',
        color: '#f44336'
      }}>
        <h2>Access Denied</h2>
        <p>You don't have permission to access this page.</p>
        <p>Your role: {userRole}</p>
        <button 
          onClick={() => window.history.back()} 
          style={{
            padding: '10px 20px',
            backgroundColor: '#3f51b5',
            color: 'white',
            border: 'none',
            borderRadius: '4px',
            cursor: 'pointer',
            marginTop: '20px'
          }}
        >
          Go Back
        </button>
      </div>
    );
  }

  return children;
}

function parseBackendResponse(data) {
  let responseData = data;
  if (data && data.output) {
    try {
      responseData = typeof data.output === 'string' ? JSON.parse(data.output) : data.output;
    } catch (_) {
      return null;
    }
  }
  if (responseData && responseData.status === 'success' && responseData.user_details) {
    const ud = responseData.user_details;
    const email = (ud.email_id || ud.email || '').trim() || null;
    const role = ud.role_identifier || ud.role || 'App User';
    const firstName = ud.first_name || '';
    const lastName = ud.last_name || '';
    return {
      role,
      email,
      first_name: firstName,
      last_name: lastName,
      user_details: ud,
    };
  }
  return null;
}

async function fetchFromEndpoint(url) {
  const response = await fetch(url, { credentials: 'include', method: 'GET', headers: { 'Content-Type': 'application/json' } });
  if (!response.ok) return null;
  const data = await response.json();
  return parseBackendResponse(data);
}

// Fetch user profile from backend and/or Catalyst (works for App User and App Administrator)
async function fetchUserRoleFromBackend() {
  try {
    // 1) Try Catalyst auth first – use for profile when available (both App User and App Administrator)
    if (typeof window !== 'undefined' && window.catalyst?.auth?.isUserAuthenticated) {
      const authData = await window.catalyst.auth.isUserAuthenticated();
      if (authData && authData.content) {
        const content = authData.content;
        const email = (content.email_id || content.email || '').trim() || null;
        const roleName = content.role_details?.role_name || content.role_identifier || content.role || 'App User';
        const firstName = content.first_name || '';
        const lastName = content.last_name || '';
        const displayName = [firstName, lastName].filter(Boolean).join(' ') || email;
        console.log('Profile from Catalyst auth:', { role: roleName, email, name: displayName });
        if (email || roleName) {
          return {
            role: roleName,
            email: email || null,
            first_name: firstName,
            last_name: lastName,
            user_details: content,
          };
        }
      }
    }

    // 2) Try authorized_portal_function (basicio – HTTP GET with user context)
    console.log('Trying backend: authorized_portal_function...');
    let result = await fetchFromEndpoint('/server/authorized_portal_function');
    if (result) {
      console.log('User from authorized_portal_function:', result);
      return result;
    }

    // 3) Fallback: authorization_portal_function
    console.log('Trying backend: authorization_portal_function...');
    result = await fetchFromEndpoint('/server/authorization_portal_function');
    if (result) {
      console.log('User from authorization_portal_function:', result);
      return result;
    }

    return null;
  } catch (err) {
    console.error('Error fetching user role:', err);
    return null;
  }
}

// Extract email from Catalyst auth response (handles various response shapes)
function getEmailFromAuth(auth) {
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
    auth.user_id
  ];
  for (const v of candidates) {
    const s = (v && typeof v === 'string' && v.trim()) || null;
    if (s && s.includes('@')) return s;
  }
  return null;
}

function App() {
  const [userRole, setUserRole] = useState(localStorage.getItem('userRole') || null);
  const [userEmail, setUserEmail] = useState(localStorage.getItem('userEmail') || null);
  const [userName, setUserName] = useState(localStorage.getItem('userName') || null);

  // Fetch user profile (Catalyst first, then backend – works for App User and App Administrator)
  useEffect(() => {
    async function fetchUserRole() {
      try {
        const result = await fetchUserRoleFromBackend();
        if (result && typeof result === 'object') {
          const role = result.role || 'App User';
          const email = result.email || null;
          const firstName = result.first_name || '';
          const lastName = result.last_name || '';
          const displayName = [firstName, lastName].filter(Boolean).join(' ') || email || null;

          console.log('User profile set:', { role, email, name: displayName });
          setUserRole(role);
          localStorage.setItem('userRole', role);
          if (displayName) {
            setUserName(displayName);
            localStorage.setItem('userName', displayName);
          }
          if (email) {
            setUserEmail(email);
            localStorage.setItem('userEmail', email);
          }
          if (result.user_details?.contractor_info) {
            localStorage.setItem('contractorInfo', JSON.stringify(result.user_details.contractor_info));
          }
        } else {
          // No profile from Catalyst or backend – use localStorage
          const storedRole = localStorage.getItem('userRole');
          const storedEmail = localStorage.getItem('userEmail');
          const storedName = localStorage.getItem('userName');
          setUserRole(storedRole || 'App User');
          if (storedEmail) setUserEmail(storedEmail);
          if (storedName) setUserName(storedName);
        }
      } catch (error) {
        console.error('Error fetching user profile:', error);
        setUserRole(localStorage.getItem('userRole') || 'App User');
        const stored = localStorage.getItem('userEmail');
        if (stored) setUserEmail(stored);
      }
    }
    fetchUserRole();
  }, []);

  // Retry: Catalyst or backend may return profile later – re-fetch so profile displays for both roles
  useEffect(() => {
    const timeouts = [500, 1500, 3000].map((ms) =>
      setTimeout(async () => {
        try {
          const result = await fetchUserRoleFromBackend();
          if (result && typeof result === 'object') {
            const role = result.role || 'App User';
            const email = result.email || null;
            const displayName = [result.first_name, result.last_name].filter(Boolean).join(' ') || email;
            setUserRole((prev) => prev || role);
            if (email) {
              setUserEmail((prev) => prev || email);
              localStorage.setItem('userEmail', email);
            }
            if (displayName) {
              setUserName((prev) => prev || displayName);
              localStorage.setItem('userName', displayName);
            }
            localStorage.setItem('userRole', role);
          }
        } catch (_) {}
      }, ms)
    );
    return () => timeouts.forEach(clearTimeout);
  }, []);

  return (
    <Router>
      <Routes>
        {/* All routes go through MainApp which includes the sidebar */}
        <Route
          path="/*"
          element={
            <ProtectedRoute>
              <MainApp userRole={userRole} userEmail={userEmail} userName={userName} />
            </ProtectedRoute>
          }
        />
      </Routes>
    </Router>
  );
}

function MainApp({ userRole, userEmail, userName }) {
  const [sidebarOpen, setSidebarOpen] = useState(true);
  const location = useLocation();

  useEffect(() => {
    console.log('MainApp user:', { userRole, userEmail, userName });
  }, [userRole, userEmail, userName]);

  // Restricted users: limited to allowed paths list. afrindinusha@gmail.com is NOT restricted so they can navigate to all pages (Statutory CLRA, Checklist, Calendar, etc.).
  const isRestrictedUser = userEmail === 'afrindinu29@gmail.com';
  // Audit-only user: sidebar shows only Home, Audit, Audit Report, and site-wise audit links.
  const isAuditOnlyUser = userEmail === 'afrindinu14@gmail.com';
  const auditOnlyAllowedPaths = ['/', '/rule-book/main-audit', '/audit-report', '/mainreport', '/returned-report', '/rule-book/statutory', '/settings'];

  const shouldShowOnlySiteManagement = false;

  // Custom Dashboard component that redirects specific users
  const CustomDashboard = () => {
    return <Dashboard userRole={userRole} />;
  };

  // Component to redirect restricted users
  const RestrictedRoute = ({ children, allowedPaths }) => {
    if (isRestrictedUser) {
      // Use React Router's location instead of window.location
      const currentPath = location.pathname;
      console.log('RestrictedRoute check:', { currentPath, allowedPaths, isRestrictedUser });
      // Allow access to all paths for restricted users if they're in the allowed list
      if (allowedPaths.includes('*') || allowedPaths.includes(currentPath)) {
        console.log('Access allowed for path:', currentPath);
        return children;
      } else {
        console.log('Access denied for path:', currentPath, 'redirecting to /');
        return <Navigate to="/" replace />;
      }
    }
    return children;
  };

  const isHomePage = location.pathname === '/';
  const isHcmDashboardPage = location.pathname === '/hcm-dashboard';
  const isNewDashboardPage = location.pathname === '/new-dashboard';
  const isCompanyDetailsHcmShell = location.pathname === '/rule-book/company-details';
  const isSiteManagementHcmShell = location.pathname === '/rule-book/site-management';
  const isActsBulkHcmShell = location.pathname === '/rule-book/actsbulk';
  const isActDescriptionsHcmShell = location.pathname === '/rule-book/act-descriptions';
  const isCalendarPickerHcmShell = location.pathname === '/calendar-picker';
  const isComplianceHcmShell = location.pathname === '/compliance';
  const isStatutoryHcmShell = location.pathname === '/rule-book/statutory';
  const isChecklistBulkHcmShell = location.pathname === '/rule-book/checklistbulk';
  const isFormmasterHcmShell = location.pathname === '/rule-book/formmaster';
  const isMainReportHcmShell = location.pathname === '/mainreport';
  const isReturnedReportHcmShell = location.pathname === '/returned-report';
  const isSettingsHcmShell = location.pathname === '/settings';
  const isSetupHcmShell = location.pathname === '/setup';
  const isPeopleHcmShell = location.pathname === '/rule-book/people';
  const isAttendanceHcmShell = location.pathname === '/rule-book/attendance';
  const isLeaveHcmShell = location.pathname === '/rule-book/leave';
  const isNewLeaveHcmShell = location.pathname === '/rule-book/newleave';
  const isApprovedLeavesHcmShell = location.pathname === '/rule-book/approved-leaves';
  const isPayrollHcmShell = location.pathname === '/rule-book/payroll';
  const isSamplePayrollHcmShell = location.pathname === '/rule-book/sample-payroll';
  const isClraHcmShell = location.pathname === '/rule-book/clra';
  const isHcmFullShell =
    isHcmDashboardPage ||
    isNewDashboardPage ||
    isCompanyDetailsHcmShell ||
    isSiteManagementHcmShell ||
    isActsBulkHcmShell ||
    isActDescriptionsHcmShell ||
    isCalendarPickerHcmShell ||
    isComplianceHcmShell ||
    isStatutoryHcmShell ||
    isChecklistBulkHcmShell ||
    isFormmasterHcmShell ||
    isMainReportHcmShell ||
    isReturnedReportHcmShell ||
    isSettingsHcmShell ||
    isSetupHcmShell ||
    isPeopleHcmShell ||
    isAttendanceHcmShell ||
    isLeaveHcmShell ||
    isNewLeaveHcmShell ||
    isApprovedLeavesHcmShell ||
    isPayrollHcmShell ||
    isSamplePayrollHcmShell ||
    isClraHcmShell;

  // When user navigates, ensure NewHome sidebar is open (so it's visible on both home and other pages)
  useEffect(() => {
    setSidebarOpen(true);
  }, [location.pathname]);

  const dashboardInitials = (() => {
    if (userName && userName.trim()) {
      const parts = userName.trim().split(/\s+/).filter(Boolean);
      if (parts.length >= 2) {
        return `${parts[0][0]}${parts[1][0]}`.toUpperCase();
      }
      return parts[0].slice(0, 2).toUpperCase();
    }
    if (userEmail && userEmail.includes('@')) {
      return userEmail.slice(0, 2).toUpperCase();
    }
    return 'RA';
  })();

  return (
    <div style={{ display: 'flex', height: '100%', minHeight: '100%' }}>
      {/* NewHome-style sidebar for all routes (home and non-home) */}
      {!isHcmFullShell && (
        <NewHomeSidebar
          isOpen={sidebarOpen}
          onClose={() => setSidebarOpen(false)}
          userEmail={userEmail}
          userName={userName}
          userRole={userRole}
        />
      )}
      <div
        style={{
          marginLeft: isHcmFullShell ? 0 : sidebarOpen ? 260 : 0,
          flex: 1,
          minWidth: 0,
          minHeight: 0,
          height: '100vh',
          maxHeight: '100vh',
          overflow: 'hidden',
          transition: 'margin-left 0.2s ease',
          display: 'flex',
          flexDirection: 'column',
          background: isHcmFullShell ? '#ffffff' : isHomePage ? 'transparent' : '#f8f9fa',
        }}
      >
        {!isHomePage && !isHcmFullShell && (
          <Header onMenuToggle={() => setSidebarOpen(!sidebarOpen)} userEmail={userEmail} userName={userName} />
        )}
        <div className="main-app-routes-shell" style={{ padding: 0, display: 'flex', flexDirection: 'column' }}>
          <RestrictedRoute allowedPaths={
            shouldShowOnlySiteManagement 
              ? ['/', '/rule-book/site-management', '/assessments/assessment-report', '/assessments/vendor-view', '/assessments/face-sheets', '/assessments/category-score'] 
              : isAuditOnlyUser
                ? auditOnlyAllowedPaths
                : isRestrictedUser 
                  ? ['/', '/new-dashboard', '/hcm-dashboard', '/rule-book/company-details', '/rule-book/site-management', '/rule-book/se-master', '/rule-book/formmaster', '/rule-book/actsbulk', '/rule-book/act-descriptions', '/rule-book/regulationsbulk', '/rule-book/regulations', '/rule-book/statutory', '/rule-book/checklist', '/rule-book/checklistbulk', '/rule-book/people', '/rule-book/attendance', '/rule-book/leave', '/rule-book/newleave', '/rule-book/approved-leaves', '/rule-book/payroll', '/rule-book/sample-payroll', '/rule-book/clra', '/rule-book/todo-list', '/rule-book/calendarbulk', '/rule-book/statutory-master', '/rule-book/statutorymasterfactory', '/form-builder', '/calendar-picker', '/compliance', '/certificate', '/mainreport', '/returned-report', '/settings', '/setup', '/assessments/assessment-report', '/assessments/vendor-view', '/assessments/face-sheets', '/assessments/category-score'] 
                  : ['*']
          }>
            <Routes>
              {/* Home page: open Dashboard for all users */}
              <Route
                path="/"
                element={<Navigate to="/hcm-dashboard" replace />}
              />
              <Route
                path="/new-dashboard"
                element={
                  <HcmDashboardPageShell
                    userName={userName || 'User'}
                    userRole={userRole || 'App User'}
                    userInitials={dashboardInitials}
                    userEmail={userEmail}
                  >
                    <Dashboard userRole={userRole} userEmail={userEmail} standalone />
                  </HcmDashboardPageShell>
                }
              />
              <Route
                path="/hcm-dashboard"
                element={
                  <NewDashboard
                    userName={userName || 'Ravi Kumar'}
                    userRole={userRole || 'HR Admin'}
                    userInitials={dashboardInitials}
                    userEmail={userEmail}
                  />
                }
              />
              <Route
                path="/compliance"
                element={
                  <HcmDashboardPageShell
                    userName={userName || 'User'}
                    userRole={userRole || 'App User'}
                    userInitials={dashboardInitials}
                    userEmail={userEmail}
                  >
                    <Compliance userRole={userRole} userEmail={userEmail} />
                  </HcmDashboardPageShell>
                }
              />
              <Route path="/rule-book/acts" element={<Acts userRole={userRole} userEmail={userEmail} />} />
              <Route
                path="/rule-book/act-descriptions"
                element={
                  <HcmDashboardPageShell
                    userName={userName || 'User'}
                    userRole={userRole || 'App User'}
                    userInitials={dashboardInitials}
                    userEmail={userEmail}
                    mainExtraClassName="nd-content--viewport-scroll"
                  >
                    <ActDescription userRole={userRole} userEmail={userEmail} />
                  </HcmDashboardPageShell>
                }
              />
              <Route path="/rule-book/scoring-criteria" element={<ScoringCriteria userRole={userRole} userEmail={userEmail} />} />
              <Route path="/rule-book/records" element={<RecordCategory userRole={userRole} userEmail={userEmail} />} />
              <Route path="/rule-book/client-details" element={<ClientDetails userRole={userRole} userEmail={userEmail} />} />
              <Route path="/rule-book/states" element={<States userRole={userRole} userEmail={userEmail} />} />
              <Route
                path="/rule-book/actsbulk"
                element={
                  <HcmDashboardPageShell
                    userName={userName || 'User'}
                    userRole={userRole || 'App User'}
                    userInitials={dashboardInitials}
                    userEmail={userEmail}
                    mainExtraClassName="nd-content--viewport-scroll"
                  >
                    <ActsBulk userRole={userRole} userEmail={userEmail} />
                  </HcmDashboardPageShell>
                }
              />
              <Route path="/rule-book/audit-master-import" element={<AuditMasterImport userRole={userRole} userEmail={userEmail} />} />
              <Route path="/rule-book/main-audit" element={<Audit userRole={userRole} userEmail={userEmail} />} />
              <Route path="/rule-book/calendarbulk" element={<CalendarBulk userRole={userRole} userEmail={userEmail} />} />
              <Route path="/rule-book/regulationsbulk" element={<RegulationsBulk userRole={userRole} userEmail={userEmail} />} />
              <Route path="/rule-book/regulations" element={<Regulations userRole={userRole} userEmail={userEmail} />} />
              <Route
                path="/rule-book/statutory"
                element={
                  <HcmDashboardPageShell
                    userName={userName || 'User'}
                    userRole={userRole || 'App User'}
                    userInitials={dashboardInitials}
                    userEmail={userEmail}
                  >
                    <Statutory userRole={userRole} userEmail={userEmail} />
                  </HcmDashboardPageShell>
                }
              />
              <Route
                path="/rule-book/company-details"
                element={
                  <HcmDashboardPageShell
                    userName={userName || 'User'}
                    userRole={userRole || 'App User'}
                    userInitials={dashboardInitials}
                    userEmail={userEmail}
                  >
                    <CompanyDetails userRole={userRole} userEmail={userEmail} />
                  </HcmDashboardPageShell>
                }
              />
              <Route
                path="/rule-book/site-management"
                element={
                  <HcmDashboardPageShell
                    userName={userName || 'User'}
                    userRole={userRole || 'App User'}
                    userInitials={dashboardInitials}
                    userEmail={userEmail}
                  >
                    <SiteManagement userEmail={userEmail} userRole={userRole} />
                  </HcmDashboardPageShell>
                }
              />
              <Route path="/rule-book/se-master" element={<SEMaster />} />
              <Route
                path="/rule-book/formmaster"
                element={
                  <HcmDashboardPageShell
                    userName={userName || 'User'}
                    userRole={userRole || 'App User'}
                    userInitials={dashboardInitials}
                    userEmail={userEmail}
                    mainExtraClassName="nd-content--viewport-scroll"
                  >
                    <Formmaster />
                  </HcmDashboardPageShell>
                }
              />
              <Route path="/certificate" element={<Certificate userRole={userRole} userEmail={userEmail} />} />
              <Route path="/rule-book/unique-ids" element={<Unique userRole={userRole} userEmail={userEmail} />} />
              <Route path="/rule-book/uniqueid-list" element={<UniqueId userRole={userRole} userEmail={userEmail} />} />
              <Route path="/rule-book/form12" element={<Form12 userRole={userRole} userEmail={userEmail} />} />
              <Route path="/rule-book/form15I" element={<Form15I userRole={userRole} userEmail={userEmail} />} />
              <Route path="/rule-book/form15II" element={<Form15II userRole={userRole} userEmail={userEmail} />} />
              <Route path="/rule-book/form25" element={<Form25 userRole={userRole} userEmail={userEmail} />} />
              <Route path="/rule-book/checklist" element={<Checklist userRole={userRole} userEmail={userEmail} />} />
              <Route
                path="/rule-book/checklistbulk"
                element={
                  <HcmDashboardPageShell
                    userName={userName || 'User'}
                    userRole={userRole || 'App User'}
                    userInitials={dashboardInitials}
                    userEmail={userEmail}
                    mainExtraClassName="nd-content--viewport-scroll"
                  >
                    <Checklistbulk userRole={userRole} userEmail={userEmail} />
                  </HcmDashboardPageShell>
                }
              />
              <Route path="/rule-book/todo-list" element={<ToDoList userEmail={userEmail} />} />
              <Route
                path="/rule-book/people"
                element={
                  <HcmDashboardPageShell
                    userName={userName || 'User'}
                    userRole={userRole || 'App User'}
                    userInitials={dashboardInitials}
                    userEmail={userEmail}
                    mainExtraClassName="nd-content--viewport-scroll"
                  >
                    <People userRole={userRole} userEmail={userEmail} />
                  </HcmDashboardPageShell>
                }
              />
              <Route
                path="/rule-book/attendance"
                element={
                  <HcmDashboardPageShell
                    userName={userName || 'User'}
                    userRole={userRole || 'App User'}
                    userInitials={dashboardInitials}
                    userEmail={userEmail}
                    mainExtraClassName="nd-content--viewport-scroll"
                  >
                    <Attendance userRole={userRole} userEmail={userEmail} />
                  </HcmDashboardPageShell>
                }
              />
              <Route
                path="/rule-book/leave"
                element={
                  <HcmDashboardPageShell
                    userName={userName || 'User'}
                    userRole={userRole || 'App User'}
                    userInitials={dashboardInitials}
                    userEmail={userEmail}
                    mainExtraClassName="nd-content--viewport-scroll"
                  >
                    <Leave userRole={userRole} userEmail={userEmail} />
                  </HcmDashboardPageShell>
                }
              />
              <Route
                path="/rule-book/newleave"
                element={
                  <HcmDashboardPageShell
                    userName={userName || 'User'}
                    userRole={userRole || 'App User'}
                    userInitials={dashboardInitials}
                    userEmail={userEmail}
                    mainExtraClassName="nd-content--viewport-scroll"
                  >
                    <NewLeave userRole={userRole} userEmail={userEmail} />
                  </HcmDashboardPageShell>
                }
              />
              <Route
                path="/rule-book/approved-leaves"
                element={
                  <HcmDashboardPageShell
                    userName={userName || 'User'}
                    userRole={userRole || 'App User'}
                    userInitials={dashboardInitials}
                    userEmail={userEmail}
                    mainExtraClassName="nd-content--viewport-scroll"
                  >
                    <ApprovedLeaves userRole={userRole} userEmail={userEmail} />
                  </HcmDashboardPageShell>
                }
              />
              <Route
                path="/rule-book/payroll"
                element={
                  <HcmDashboardPageShell
                    userName={userName || 'User'}
                    userRole={userRole || 'App User'}
                    userInitials={dashboardInitials}
                    userEmail={userEmail}
                    mainExtraClassName="nd-content--viewport-scroll"
                  >
                    <Payroll />
                  </HcmDashboardPageShell>
                }
              />
              <Route
                path="/rule-book/sample-payroll"
                element={
                  <HcmDashboardPageShell
                    userName={userName || 'User'}
                    userRole={userRole || 'App User'}
                    userInitials={dashboardInitials}
                    userEmail={userEmail}
                    mainExtraClassName="nd-content--viewport-scroll"
                  >
                    <SamplePayroll userRole={userRole} userEmail={userEmail} />
                  </HcmDashboardPageShell>
                }
              />
              <Route
                path="/rule-book/clra"
                element={
                  <HcmDashboardPageShell
                    userName={userName || 'User'}
                    userRole={userRole || 'App User'}
                    userInitials={dashboardInitials}
                    userEmail={userEmail}
                    mainExtraClassName="nd-content--viewport-scroll"
                  >
                    <CLRA userRole={userRole} userEmail={userEmail} />
                  </HcmDashboardPageShell>
                }
              />
              <Route path="/rule-book/statutory-master" element={<StatutoryMaster userRole={userRole} userEmail={userEmail} />} />
              <Route path="/rule-book/statutorymasterfactory" element={<Statutorymasterfactory userRole={userRole} userEmail={userEmail} />} />
              <Route path="/rule-book/factory" element={<Factory userRole={userRole} userEmail={userEmail} />} />
              <Route path="/form-builder" element={<FormBuilder userRole={userRole} userEmail={userEmail} />} />
              <Route path="/assessments/assessment-report" element={<AssessmentReport userRole={userRole} userEmail={userEmail} />} />
              <Route
                path="/calendar-picker"
                element={
                  <HcmDashboardPageShell
                    userName={userName || 'User'}
                    userRole={userRole || 'App User'}
                    userInitials={dashboardInitials}
                    userEmail={userEmail}
                    mainExtraClassName="nd-content--viewport-scroll"
                  >
                    <CalendarPicker userEmail={userEmail} userRole={userRole} />
                  </HcmDashboardPageShell>
                }
              />
              <Route path="/rule-book/calendar-pdf-master" element={<CalendarPDFMaster userEmail={userEmail} userRole={userRole} />} />
              <Route path="/trainer-module" element={<div>Trainer Module</div>} />
              <Route path="/trainee-module" element={<div>Trainee Module</div>} />
              <Route path="/course-creation" element={<div>Course Creation</div>} />
              <Route path="/curriculum-creation" element={<div>Curriculum Creation</div>} />
              <Route path="/learner-portal" element={<div>Learner Portal</div>} />
              <Route path="/quiz/:id" element={<div>Quiz Module</div>} />
              <Route path="/certification" element={<div>Certification</div>} />
              <Route path="/reports" element={<Reports userRole={userRole} userEmail={userEmail} />} />
              <Route
                path="/mainreport"
                element={
                  <HcmDashboardPageShell
                    userName={userName || 'User'}
                    userRole={userRole || 'App User'}
                    userInitials={dashboardInitials}
                    userEmail={userEmail}
                    mainExtraClassName="nd-content--viewport-scroll nd-content--mainreport"
                  >
                    <Mainreport userEmail={userEmail} />
                  </HcmDashboardPageShell>
                }
              />
              <Route
                path="/settings"
                element={
                  <HcmDashboardPageShell
                    userName={userName || 'User'}
                    userRole={userRole || 'App User'}
                    userInitials={dashboardInitials}
                    userEmail={userEmail}
                  >
                    <Settings />
                  </HcmDashboardPageShell>
                }
              />
              <Route
                path="/setup"
                element={
                  <HcmDashboardPageShell
                    userName={userName || 'User'}
                    userRole={userRole || 'App User'}
                    userInitials={dashboardInitials}
                    userEmail={userEmail}
                    mainExtraClassName="nd-content--viewport-scroll"
                  >
                    <Setup userEmail={userEmail} />
                  </HcmDashboardPageShell>
                }
              />
              <Route
                path="/returned-report"
                element={
                  <HcmDashboardPageShell
                    userName={userName || 'User'}
                    userRole={userRole || 'App User'}
                    userInitials={dashboardInitials}
                    userEmail={userEmail}
                    mainExtraClassName="nd-content--viewport-scroll nd-content--mainreport"
                  >
                    <ReturnedReport userEmail={userEmail} />
                  </HcmDashboardPageShell>
                }
              />
              <Route path="/audit-report" element={<AuditReport userRole={userRole} userEmail={userEmail} />} />
              <Route path="/notifications" element={<div>Notifications</div>} />
              <Route path="/feedback" element={<div>Feedback</div>} />
              <Route path="/medical-fitness" element={<div>Medical Fitness</div>} />
              {/* Catch-all route */}
              <Route path="*" element={<Navigate to="/" replace />} />
            </Routes>
          </RestrictedRoute>
        </div>
      </div>
    </div>
  );
}

export default App;