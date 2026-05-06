import React, { useState, useEffect, useRef } from 'react';
import { performLogout } from '../utils/performLogout';

function resolveDisplayEmail(userEmail) {
  if (userEmail && typeof userEmail === 'string') return userEmail;
  try {
    return localStorage.getItem('userEmail') || '';
  } catch (_) {
    return '';
  }
}

/**
 * Profile avatar in HCM dashboard header: opens menu with Logout → Catalyst auth page.
 */
export default function HcmHeaderProfileMenu({ userInitials, userEmail }) {
  const [open, setOpen] = useState(false);
  const wrapRef = useRef(null);
  const displayEmail = resolveDisplayEmail(userEmail);

  useEffect(() => {
    const onDown = (e) => {
      if (wrapRef.current && !wrapRef.current.contains(e.target)) {
        setOpen(false);
      }
    };
    document.addEventListener('mousedown', onDown);
    return () => document.removeEventListener('mousedown', onDown);
  }, []);

  return (
    <div className="nd-profile-wrap" ref={wrapRef}>
      <button
        type="button"
        className="nd-avatar nd-avatar--header nd-avatar--header-btn"
        onClick={() => setOpen((v) => !v)}
        aria-label="Profile menu"
        aria-expanded={open}
        title={displayEmail ? `Logged in as ${displayEmail}` : 'Profile'}
      >
        {userInitials}
      </button>
      {open && (
        <div className="nd-header-dropdown" role="menu">
          <div className="nd-header-dropdown-label">Logged in as</div>
          <div className="nd-header-dropdown-email">{displayEmail || 'User'}</div>
          <button
            type="button"
            className="nd-header-dropdown-logout"
            role="menuitem"
            onClick={() => {
              setOpen(false);
              performLogout();
            }}
          >
            Logout
          </button>
        </div>
      )}
    </div>
  );
}
