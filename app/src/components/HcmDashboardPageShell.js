import React from 'react';
import { Bell, CircleHelp } from 'lucide-react';
import HcmDashboardSidebar from './HcmDashboardSidebar';
import HcmHeaderProfileMenu from './HcmHeaderProfileMenu';
import '../Pages/newdashboard.css';

/**
 * Full /hcm-dashboard chrome (sidebar + top bar) wrapping arbitrary page content.
 * @param {string} [mainExtraClassName] — e.g. `nd-content--viewport-scroll` so `<main>` scrolls (fixes mouse wheel vs nested overflow).
 */
export default function HcmDashboardPageShell({ children, userName, userRole, userInitials, userEmail, mainExtraClassName = '' }) {
  const mainClass = ['nd-content', 'nd-content--embedded-page', mainExtraClassName].filter(Boolean).join(' ');
  return (
    <div className="nd-root nd-root--embed-shell">
      <HcmDashboardSidebar
        userName={userName}
        userRole={userRole}
        userInitials={userInitials}
        userEmail={userEmail}
      />
      <div className="nd-main">
        <header className="nd-header">
          <h1 className="nd-header-title">HR Compliance Management</h1>
          <div className="nd-header-actions">
            <button type="button" className="nd-icon-btn nd-icon-btn--badge" aria-label="Notifications">
              <Bell size={20} strokeWidth={2} />
              <span className="nd-badge-count">3</span>
            </button>
            <button type="button" className="nd-icon-btn" aria-label="Help">
              <CircleHelp size={20} strokeWidth={2} />
            </button>
            <HcmHeaderProfileMenu userInitials={userInitials} userEmail={userEmail} />
          </div>
        </header>
        <main className={mainClass}>{children}</main>
      </div>
    </div>
  );
}
