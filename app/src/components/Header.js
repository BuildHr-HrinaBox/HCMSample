import React, { useState, useEffect, useRef } from 'react';
import { MoreVertical } from 'lucide-react';
import { performLogout } from '../utils/performLogout';
import './Header.css';

function headerInitials(userName, userEmail) {
  if (userName && String(userName).trim()) {
    const parts = String(userName).trim().split(/\s+/).filter(Boolean);
    if (parts.length >= 2) {
      return `${parts[0][0]}${parts[1][0]}`.toUpperCase();
    }
    return parts[0].slice(0, 2).toUpperCase();
  }
  const raw = userEmail ?? (typeof localStorage !== 'undefined' ? localStorage.getItem('userEmail') : null) ?? '';
  const displayEmail = typeof raw === 'string' ? raw : '';
  if (displayEmail && displayEmail.includes('@')) {
    return displayEmail.slice(0, 2).toUpperCase();
  }
  return 'U';
}

const Header = ({ onMenuToggle, userEmail, userName }) => {
  const [showProfileDropdown, setShowProfileDropdown] = useState(false);
  const profileRef = useRef(null);
  const raw = userEmail ?? (typeof localStorage !== 'undefined' ? localStorage.getItem('userEmail') : null) ?? '';
  const displayEmail =
    typeof raw === 'string' ? raw : raw && typeof raw === 'object' && (raw.email_id || raw.email) ? raw.email_id || raw.email : '';

  useEffect(() => {
    const handleClickOutside = (e) => {
      if (profileRef.current && !profileRef.current.contains(e.target)) {
        setShowProfileDropdown(false);
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  const initials = headerInitials(userName, displayEmail);

  return (
    <header className="hcm-app-header">
      <div className="hcm-app-header-left">
        <button type="button" className="hcm-app-header-menu" onClick={onMenuToggle} aria-label="Toggle menu">
          <MoreVertical size={22} strokeWidth={2} />
        </button>
        <h1 className="hcm-app-header-title">HR Compliance Management</h1>
      </div>
      <div className="hcm-app-header-right">
        <div ref={profileRef} style={{ position: 'relative' }}>
          <button
            type="button"
            className="hcm-app-header-avatar"
            onClick={() => setShowProfileDropdown((prev) => !prev)}
            aria-label="Profile"
            aria-expanded={showProfileDropdown}
            title={displayEmail ? `Logged in as ${displayEmail}` : 'User'}
          >
            {initials}
          </button>
          {showProfileDropdown && (
            <div className="hcm-app-header-dropdown" role="menu">
              <div className="hcm-app-header-dropdown-label">Logged in as</div>
              <div className="hcm-app-header-dropdown-email">{displayEmail || 'User'}</div>
              <button
                type="button"
                className="hcm-app-header-dropdown-logout"
                role="menuitem"
                onClick={() => {
                  setShowProfileDropdown(false);
                  performLogout();
                }}
              >
                Logout
              </button>
            </div>
          )}
        </div>
      </div>
    </header>
  );
};

export default Header;
