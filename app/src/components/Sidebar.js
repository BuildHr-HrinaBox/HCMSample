import React from 'react';
import { Link, useLocation } from 'react-router-dom';

const Sidebar = ({ isOpen, userEmail, userRole }) => {
  const location = useLocation();

  const navItems = [
    { path: '/', label: 'Home', icon: '⌂' },
    { path: '/compliance', label: 'Dashboard', icon: '◷' },
    { path: '/rule-book/acts', label: 'Acts', icon: '📄' },
    { path: '/rule-book/act-descriptions', label: 'Applicable Acts Library', icon: '📚' },
    { path: '/rule-book/actsbulk', label: 'Acts Bulk', icon: '📋' },
    { path: '/rule-book/calendarbulk', label: 'Calendar Bulk', icon: '📅' },
    { path: '/rule-book/regulationsbulk', label: 'Regulations Bulk', icon: '📜' },
    { path: '/rule-book/regulations', label: 'Regulations', icon: '📜' },
    { path: '/rule-book/statutory', label: 'Statutory', icon: '📋' },
    { path: '/rule-book/site-management', label: 'Site Management', icon: '🏢' },
    { path: '/rule-book/se-master', label: 'S&E Master', icon: '📑' },
    { path: '/rule-book/company-details', label: 'Company Details', icon: '🏛' },
    { path: '/rule-book/checklist', label: 'Checklist', icon: '☑' },
    { path: '/rule-book/checklistbulk', label: 'Checklist Bulk', icon: '📋' },
    { path: '/rule-book/people', label: 'People', icon: '👥' },
    { path: '/rule-book/attendance', label: 'Attendance', icon: '📅' },
    { path: '/rule-book/leave', label: 'Leave', icon: '🏖' },
    { path: '/rule-book/newleave', label: 'New Leave', icon: '🏖' },
    { path: '/rule-book/approved-leaves', label: 'Approved Leaves', icon: '✓' },
    { path: '/rule-book/payroll', label: 'Payroll', icon: '💰' },
    { path: '/rule-book/clra', label: 'CLRA', icon: '📋' },
    { path: '/calendar-picker', label: 'Calendar', icon: '📆' },
    { path: '/certificate', label: 'Certificate', icon: '📜' },
  ];

  return (
    <div
      style={{
        position: 'fixed',
        left: 0,
        top: 0,
        bottom: 0,
        width: isOpen ? 220 : 0,
        background: 'linear-gradient(180deg, #8B4513 0%, #A0522D 100%)',
        color: '#fff',
        transition: 'width 0.2s ease',
        zIndex: 100,
        overflow: 'hidden',
        display: 'flex',
        flexDirection: 'column',
        direction: 'ltr',
        boxSizing: 'border-box',
      }}
    >
      <div style={{
        padding: isOpen ? '16px' : '12px 8px',
        display: 'flex',
        alignItems: 'center',
        justifyContent: isOpen ? 'flex-start' : 'center',
        gap: 10,
        minHeight: 64,
        borderBottom: '1px solid rgba(255,255,255,0.2)',
        boxSizing: 'border-box',
        flexShrink: 0,
      }}>
        <div style={{ width: 36, height: 36, background: 'rgba(255,255,255,0.2)', borderRadius: 8, flexShrink: 0 }} />
        {isOpen && (
          <div style={{ minWidth: 0, flex: 1, overflow: 'hidden' }}>
            <div style={{ fontWeight: 700, fontSize: '0.95rem' }}>HCM</div>
            <div style={{ fontSize: '0.65rem', opacity: 0.9 }}>HR Compliance Management</div>
          </div>
        )}
      </div>
      <nav style={{ flex: 1, overflowY: 'auto', overflowX: 'hidden', padding: '8px 0', minWidth: 0 }}>
        {navItems.map((item) => {
          const active = location.pathname === item.path;
          return (
            <Link
              key={item.path}
              to={item.path}
              style={{
                display: 'flex',
                alignItems: 'center',
                justifyContent: isOpen ? 'flex-start' : 'center',
                gap: isOpen ? 12 : 0,
                padding: isOpen ? '12px 16px' : '14px 0',
                color: '#fff',
                textDecoration: 'none',
                background: active ? 'rgba(255,255,255,0.15)' : 'transparent',
                borderLeft: active ? '4px solid #fff' : '4px solid transparent',
                minHeight: 44,
                boxSizing: 'border-box',
                minWidth: 0,
              }}
            >
              <span style={{ fontSize: '1rem', flexShrink: 0, width: 20, textAlign: 'center', display: 'inline-block' }}>{item.icon}</span>
              {isOpen && <span style={{ fontSize: '0.8rem', flex: 1, minWidth: 0, wordBreak: 'break-word', overflowWrap: 'break-word', lineHeight: 1.3 }}>{item.label}</span>}
            </Link>
          );
        })}
      </nav>
    </div>
  );
};

export default Sidebar;
