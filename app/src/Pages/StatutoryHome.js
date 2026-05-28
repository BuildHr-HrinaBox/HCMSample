import React, { useState, useEffect, useRef } from 'react';
import { Link } from 'react-router-dom';
import './StatutoryHome.css';

const StatutoryHome = ({ userRole, userEmail, onMenuToggle }) => {
  const displayEmail = userEmail || (typeof localStorage !== 'undefined' ? localStorage.getItem('userEmail') : null) || '';
  const userInitial = (displayEmail && displayEmail.charAt(0).toUpperCase()) || 'A';
  const [showProfileDropdown, setShowProfileDropdown] = useState(false);
  const profileRef = useRef(null);

  useEffect(() => {
    const handleClickOutside = (e) => {
      if (profileRef.current && !profileRef.current.contains(e.target)) {
        setShowProfileDropdown(false);
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  return (
    <div className="statutory-home">
      {/* Header - reddish-brown gradient */}
      <header className="statutory-home-header">
        <div className="statutory-home-header-top">
          <div className="statutory-home-header-left">
            <button type="button" className="statutory-home-menu-btn" aria-label="Menu" onClick={onMenuToggle}>
              <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <circle cx="12" cy="6" r="1.5" fill="currentColor" />
                <circle cx="12" cy="12" r="1.5" fill="currentColor" />
                <circle cx="12" cy="18" r="1.5" fill="currentColor" />
              </svg>
            </button>
            <span className="statutory-home-logo-title" style={{ color: '#fff' }}>HCM</span>
          </div>

          <div className="statutory-home-header-center">
            <h1 className="statutory-home-title">HCM</h1>
          </div>

          <div className="statutory-home-header-right">
            <div className="statutory-home-profile-wrap" ref={profileRef}>
              <button
                type="button"
                className="statutory-home-avatar statutory-home-avatar-btn"
                title={displayEmail ? `Logged in as ${displayEmail}` : 'User'}
                onClick={() => setShowProfileDropdown((prev) => !prev)}
                aria-label="Profile"
                aria-expanded={showProfileDropdown}
              >
                {userInitial}
              </button>
              {showProfileDropdown && (
                <div className="statutory-home-profile-dropdown">
                  <div className="statutory-home-profile-dropdown-label">Logged in as</div>
                  <div className="statutory-home-profile-dropdown-email">{displayEmail || 'User'}</div>
                </div>
              )}
            </div>
          </div>
        </div>

        <div className="statutory-home-kpis-row">
          <div className="statutory-home-kpi">
            <div className="statutory-home-kpi-label">Active Acts</div>
            <div className="statutory-home-kpi-value">12</div>
          </div>
          <div className="statutory-home-kpi">
            <div className="statutory-home-kpi-label">Regulation</div>
            <div className="statutory-home-kpi-value">71</div>
          </div>
          <div className="statutory-home-kpi">
            <div className="statutory-home-kpi-label">Form Submitted</div>
            <div className="statutory-home-kpi-value">42</div>
          </div>
          <div className="statutory-home-kpi">
            <div className="statutory-home-kpi-label">Compliance Rate</div>
            <div className="statutory-home-kpi-value">92%</div>
          </div>
        </div>
      </header>

      {/* Main content */}
      <main className="statutory-home-main">
        <div className="statutory-home-content-left">
          {/* About card */}
          <section className="statutory-home-about-card">
            <h2 className="statutory-home-about-title">About HR Compliance Management</h2>
            <p className="statutory-home-about-subtitle">
              HR Compliance Management - Your Complete Compliance Solution
            </p>
            <p className="statutory-home-about-desc">
              HR Compliance Management is a comprehensive platform designed to streamline and manage all aspects of
              statutory compliance, regulatory requirements, and legal obligations. Built with modern technology and
              user-friendly interfaces, HR Compliance Management provides powerful tools for administrators and users
              to efficiently manage their compliance processes.
            </p>
          </section>

          {/* System Modules & Features - triangle view (same arrow format as reference model) */}
          <section className="statutory-home-diagram-card">
            <div className="statutory-home-diagram-triangle-view">
              {/* Single SVG: triangle + dashed connection lines with arrowheads */}
              <svg
                className="statutory-home-diagram-svg"
                viewBox="0 0 500 400"
                preserveAspectRatio="xMidYMid meet"
              >
                <defs>
                  <linearGradient id="statutorySystemTriangleGradient" x1="0%" y1="0%" x2="0%" y2="100%">
                    <stop offset="0%" stopColor="#0000FF" stopOpacity="0.9" />
                    <stop offset="100%" stopColor="#00FFFF" stopOpacity="0.85" />
                  </linearGradient>
                  <marker id="statutoryArrowhead" markerWidth="10" markerHeight="10" refX="9" refY="3" orient="auto">
                    <polygon points="0 0, 10 3, 0 6" fill="#7C3AED" opacity="0.6" />
                  </marker>
                </defs>
                {/* Triangle: top, left-bottom, right-bottom (same shape as reference) */}
                <polygon
                  points="250,55 100,320 400,320"
                  fill="url(#statutorySystemTriangleGradient)"
                  className="statutory-home-triangle-fill"
                  fillOpacity="0.8"
                />
                <polygon
                  points="250,55 100,320 400,320"
                  fill="none"
                  stroke="#8B4513"
                  strokeWidth="3"
                  className="statutory-home-triangle-animated-border"
                />
                {/* Connection lines with arrows: triangle vertices -> module card positions */}
                <line x1="250" y1="55" x2="250" y2="12" stroke="#7C3AED" strokeWidth="2" strokeDasharray="5,5" opacity="0.5" markerEnd="url(#statutoryArrowhead)" className="statutory-home-connection-line" />
                <line x1="250" y1="55" x2="410" y2="18" stroke="#7C3AED" strokeWidth="2" strokeDasharray="5,5" opacity="0.5" markerEnd="url(#statutoryArrowhead)" className="statutory-home-connection-line" />
                <line x1="100" y1="320" x2="25" y2="155" stroke="#7C3AED" strokeWidth="2.5" strokeDasharray="5,5" opacity="0.6" markerEnd="url(#statutoryArrowhead)" className="statutory-home-connection-line" />
                <line x1="100" y1="320" x2="25" y2="285" stroke="#7C3AED" strokeWidth="2.5" strokeDasharray="5,5" opacity="0.6" markerEnd="url(#statutoryArrowhead)" className="statutory-home-connection-line" />
                <line x1="100" y1="320" x2="165" y2="368" stroke="#7C3AED" strokeWidth="2.5" strokeDasharray="5,5" opacity="0.6" markerEnd="url(#statutoryArrowhead)" className="statutory-home-connection-line" />
                <line x1="400" y1="320" x2="335" y2="368" stroke="#7C3AED" strokeWidth="2.5" strokeDasharray="5,5" opacity="0.6" markerEnd="url(#statutoryArrowhead)" className="statutory-home-connection-line" />
                <line x1="400" y1="320" x2="475" y2="155" stroke="#7C3AED" strokeWidth="2.5" strokeDasharray="5,5" opacity="0.6" markerEnd="url(#statutoryArrowhead)" className="statutory-home-connection-line" />
                <line x1="400" y1="320" x2="475" y2="285" stroke="#7C3AED" strokeWidth="2.5" strokeDasharray="5,5" opacity="0.6" markerEnd="url(#statutoryArrowhead)" className="statutory-home-connection-line" />
              </svg>
              {/* Center label overlay */}
              <div className="statutory-home-triangle-label">System Modules &amp; Features</div>

              {/* Module cards - positions in % to align with line endpoints (viewBox 0 0 500 400) */}
              <Link to="/rule-book/act-descriptions" className="statutory-home-module yellow lib">
                <svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="#7c3aed" strokeWidth="2">
                  <path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z" />
                  <line x1="12" y1="11" x2="12" y2="17" />
                  <line x1="9" y1="14" x2="15" y2="14" />
                </svg>
                <span className="mod-label">Library</span>
              </Link>
              <Link to="/rule-book/act-descriptions" className="statutory-home-module yellow acts">
                <svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="#7c3aed" strokeWidth="2">
                  <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" />
                  <polyline points="14 2 14 8 20 8" />
                  <line x1="16" y1="13" x2="8" y2="13" />
                  <line x1="16" y1="17" x2="8" y2="17" />
                </svg>
                <span className="mod-label">Applicable Acts</span>
              </Link>

              {/* Left - Pink: Checklist, Calendar, Transaction */}
              <Link to="/rule-book/checklist" className="statutory-home-module pink checklist">
                <svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="#059669" strokeWidth="2">
                  <path d="M9 11l3 3L22 4" />
                  <path d="M21 12v7a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11" />
                </svg>
                <span className="mod-label">Checklist</span>
              </Link>
              <Link to="/calendar-picker" className="statutory-home-module pink calendar">
                <svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="#ea580c" strokeWidth="2">
                  <rect x="3" y="4" width="18" height="18" rx="2" ry="2" />
                  <line x1="16" y1="2" x2="16" y2="6" />
                  <line x1="8" y1="2" x2="8" y2="6" />
                  <line x1="3" y1="10" x2="21" y2="10" />
                </svg>
                <span className="mod-label">Calendar</span>
              </Link>
              <Link to="/rule-book/actsbulk" className="statutory-home-module pink transaction">
                <svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="#059669" strokeWidth="2">
                  <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" />
                  <polyline points="14 2 14 8 20 8" />
                  <line x1="16" y1="13" x2="8" y2="13" />
                  <line x1="16" y1="17" x2="8" y2="17" />
                </svg>
                <span className="mod-label">Transaction</span>
              </Link>

              {/* Right - Green: Report, Dashboard, Audit */}
              <Link to="/assessments/assessment-report" className="statutory-home-module green report">
                <svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="#22c55e" strokeWidth="2">
                  <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" />
                  <polyline points="14 2 14 8 20 8" />
                  <line x1="16" y1="13" x2="8" y2="13" />
                  <line x1="16" y1="17" x2="8" y2="17" />
                </svg>
                <span className="mod-label">Report</span>
              </Link>
              <Link to="/" className="statutory-home-module green dashboard">
                <svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="#22c55e" strokeWidth="2">
                  <rect x="3" y="3" width="7" height="9" />
                  <rect x="14" y="3" width="7" height="5" />
                  <rect x="14" y="12" width="7" height="9" />
                  <rect x="3" y="16" width="7" height="5" />
                </svg>
                <span className="mod-label">Dashboard</span>
              </Link>
              <Link to="/assessments/assessment-report" className="statutory-home-module green audit">
                <svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="#22c55e" strokeWidth="2">
                  <circle cx="12" cy="12" r="10" />
                  <polyline points="12 6 12 12 16 14" />
                </svg>
                <span className="mod-label">Audit</span>
              </Link>
            </div>
          </section>
        </div>

        {/* Recent Activity sidebar */}
        <aside className="statutory-home-activity">
          <h3 className="statutory-home-activity-title">
            <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <polygon points="13 2 3 14 12 14 11 22 21 10 12 10 13 2" />
            </svg>
            Recent Activity
          </h3>
          <div className="statutory-home-activity-item">
            <div className="statutory-home-activity-icon bell">
              <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <path d="M18 8A6 6 0 0 0 6 8c0 7-3 9-3 9h18s-3-2-3-9" />
                <path d="M13.73 21a2 2 0 0 1-3.46 0" />
              </svg>
            </div>
            <div className="statutory-home-activity-body">
              <p className="statutory-home-activity-desc">Upcoming compliance deadline approaching</p>
              <p className="statutory-home-activity-time">1 hour ago</p>
            </div>
          </div>
          <div className="statutory-home-activity-item">
            <div className="statutory-home-activity-icon report">
              <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" />
                <polyline points="14 2 14 8 20 8" />
                <line x1="16" y1="13" x2="8" y2="13" />
                <line x1="16" y1="17" x2="8" y2="17" />
                <polyline points="10 9 9 9 8 9" />
              </svg>
            </div>
            <div className="statutory-home-activity-body">
              <p className="statutory-home-activity-desc">Reports</p>
              <p className="statutory-home-activity-time">2 hours ago</p>
            </div>
          </div>
          <div className="statutory-home-activity-item">
            <div className="statutory-home-activity-icon check">
              <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <path d="M22 11.08V12a10 10 0 1 1-5.93-9.14" />
                <polyline points="22 4 12 14.01 9 11.01" />
              </svg>
            </div>
            <div className="statutory-home-activity-body">
              <p className="statutory-home-activity-desc">Safety regulations compliance verified</p>
              <p className="statutory-home-activity-time">3 hours ago</p>
            </div>
          </div>
        </aside>
      </main>
    </div>
  );
};

export default StatutoryHome;
