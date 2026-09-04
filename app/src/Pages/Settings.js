import React, { useCallback, useEffect, useMemo, useState } from 'react';
import './CompanyDetails.css';
import './Settings.css';
import {
  clearSettingsCache,
  fetchSettings,
  readSettingsCache,
  writeSettingsCache
} from '../utils/settingsCache';

const API_BASE = '/server/settings_function';

const formatSettingsDueDate = (value) => {
  const raw = String(value || '').trim();
  if (!raw) return '-';
  if (/^\d{1,2}$/.test(raw)) {
    const day = Number(raw);
    if (day >= 1 && day <= 31) {
      const date = new Date();
      date.setDate(day);
      return date.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });
    }
  }
  const date = new Date(raw);
  if (!Number.isNaN(date.getTime())) {
    return date.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });
  }
  return raw;
};

const formatDueDateForInput = (value) => {
  const raw = String(value || '').trim();
  if (!raw) return '';
  if (/^\d{2}-\d{2}-\d{4}$/.test(raw)) return raw;
  if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) {
    const [year, month, day] = raw.split('-');
    return `${day}-${month}-${year}`;
  }
  if (/^\d{1,2}$/.test(raw)) {
    const day = String(Number(raw)).padStart(2, '0');
    const now = new Date();
    const month = String(now.getMonth() + 1).padStart(2, '0');
    return `${day}-${month}-${now.getFullYear()}`;
  }
  const date = new Date(raw);
  if (!Number.isNaN(date.getTime())) {
    const day = String(date.getDate()).padStart(2, '0');
    const month = String(date.getMonth() + 1).padStart(2, '0');
    return `${day}-${month}-${date.getFullYear()}`;
  }
  return raw;
};

const toDateInputValue = (value) => {
  const raw = String(value || '').trim();
  if (!raw) return '';
  if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) return raw;
  if (/^\d{1,2}$/.test(raw)) {
    const day = Number(raw);
    if (day >= 1 && day <= 31) {
      const now = new Date();
      const month = String(now.getMonth() + 1).padStart(2, '0');
      const dayStr = String(day).padStart(2, '0');
      return `${now.getFullYear()}-${month}-${dayStr}`;
    }
  }
  const date = new Date(raw);
  if (!Number.isNaN(date.getTime())) {
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    return `${date.getFullYear()}-${month}-${day}`;
  }
  return raw;
};

export function DueDateNotification({ embedded = false, hideAddButton = false, refreshKey = 0 }) {
  const cachedOnMount = readSettingsCache();
  const [showForm, setShowForm] = useState(false);
  const [savedCompanyName, setSavedCompanyName] = useState(() =>
    cachedOnMount ? cachedOnMount.companyName : ''
  );
  const [savedLogoName, setSavedLogoName] = useState(() =>
    cachedOnMount ? cachedOnMount.logoName : ''
  );
  const [savedDueDate, setSavedDueDate] = useState(() =>
    cachedOnMount ? cachedOnMount.dueDate : ''
  );
  const [formDueDate, setFormDueDate] = useState('');
  const [saving, setSaving] = useState(false);
  const [loading, setLoading] = useState(() => !cachedOnMount);
  const [message, setMessage] = useState('');

  const applySettings = useCallback((data) => {
    if (!data) return;
    setSavedCompanyName(String(data.companyName || ''));
    setSavedLogoName(String(data.logoName || ''));
    setSavedDueDate(String(data.dueDate || ''));
  }, []);

  const loadSettings = useCallback(async () => {
    const cached = readSettingsCache();
    if (cached) {
      applySettings(cached);
      setLoading(false);
    }

    try {
      const fresh = await fetchSettings({ force: !cached });
      if (fresh) applySettings(fresh);
    } catch (err) {
      if (!cached) setMessage(err?.message || 'Unable to load settings.');
    } finally {
      setLoading(false);
    }
  }, [applySettings]);

  useEffect(() => {
    loadSettings();
  }, [loadSettings, refreshKey]);

  const openAddForm = () => {
    setShowForm(true);
    setFormDueDate('');
    setMessage('');
  };

  const openEditForm = () => {
    setShowForm(true);
    setFormDueDate(formatDueDateForInput(savedDueDate));
    setMessage('');
  };

  const closeForm = () => {
    setShowForm(false);
    setMessage('');
  };

  const submitForm = async (event) => {
    event.preventDefault();
    const existingCompanyName = savedCompanyName.trim();

    setSaving(true);
    setMessage('');
    try {
      const formData = new FormData();
      formData.append('companyName', existingCompanyName);
      formData.append('dueDate', formDueDate.trim());

      const res = await fetch(`${API_BASE}/settings`, {
        method: 'PUT',
        body: formData,
      });
      const data = await res.json();
      if (!res.ok || data?.status !== 'success') {
        throw new Error(data?.message || 'Failed to save settings');
      }

      const nextCompanyName = String(data?.data?.companyName || existingCompanyName);
      const nextLogoName = String(data?.data?.logoName || '');
      const nextDueDate = String(data?.data?.dueDate || formDueDate.trim());
      setSavedCompanyName(nextCompanyName);
      setSavedLogoName(nextLogoName);
      setSavedDueDate(nextDueDate);
      writeSettingsCache(data?.data);
      setFormDueDate(nextDueDate);
      setMessage('');
      setShowForm(false);
    } catch (err) {
      setMessage(err?.message || 'Unable to save settings.');
    } finally {
      setSaving(false);
    }
  };

  const hasSavedSettings = savedCompanyName.trim() || savedLogoName.trim() || savedDueDate.trim();
  const settingsRows = hasSavedSettings
    ? [
        {
          id: '1',
          companyName: savedCompanyName || '-',
          logo: savedLogoName || '-',
          dueDate: savedDueDate || '-',
        },
      ]
    : [];
  const showLoadingPlaceholder = loading && settingsRows.length === 0;

  const handleDeleteRow = async () => {
    try {
      const res = await fetch(`${API_BASE}/settings`, { method: 'DELETE' });
      const data = await res.json();
      if (!res.ok || data?.status !== 'success') {
        throw new Error(data?.message || 'Failed to delete settings');
      }
      setSavedCompanyName('');
      setSavedLogoName('');
      setSavedDueDate('');
      clearSettingsCache();
      setFormDueDate('');
      setMessage('Deleted successfully.');
    } catch (err) {
      setMessage(err?.message || 'Unable to delete settings.');
    }
  };

  const listAndForm = !showForm ? (
    <>
      <div className="settings-page-heading settings-page-heading--row">
        {embedded ? (
          <h2 className="settings-hub-section-title">Due Date Notification</h2>
        ) : (
          <h1 className="settings-page-title">Email Notification</h1>
        )}
        <div className="settings-page-actions">
          {!hideAddButton ? (
            <button type="button" className="company-details-btn company-details-btn-add" onClick={openAddForm}>
              Add
            </button>
          ) : null}
        </div>
      </div>
      <div className={embedded ? 'settings-hub-section-body' : 'settings-list-body'}>
        {showLoadingPlaceholder ? (
          <p className="settings-list-empty">Loading...</p>
        ) : (
          <div className="company-details-table-wrap">
            <div className="settings-due-date-list">
              {settingsRows.length === 0 ? (
                <p className="settings-list-empty">
                  {hideAddButton
                    ? <>No settings added yet. Click <strong>+ Add</strong> above to open form.</>
                    : 'No settings added yet. Click Add to open form.'}
                </p>
              ) : (
                settingsRows.map((row) => (
                  <div key={row.id} className="settings-due-date-card">
                    <div className="settings-due-date-icon" aria-hidden="true">
                      <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8">
                        <rect x="3" y="4.5" width="18" height="16" rx="2" />
                        <path d="M7 2.5v4M17 2.5v4M3 9h18" />
                      </svg>
                    </div>
                    <div className="settings-due-date-info">
                      <strong>Due Date</strong>
                      <span>
                        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                          <rect x="3" y="4.5" width="18" height="16" rx="2" />
                          <path d="M7 2.5v4M17 2.5v4M3 9h18" />
                        </svg>
                        {formatSettingsDueDate(row.dueDate)}
                      </span>
                    </div>
                    <div className="settings-due-date-actions">
                      <button
                        type="button"
                        className="company-details-table-action company-details-table-action--view"
                        title="View"
                        aria-label="View settings"
                        onClick={openEditForm}
                      >
                        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                          <path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z" />
                          <circle cx="12" cy="12" r="3" />
                        </svg>
                      </button>
                      <button
                        type="button"
                        className="company-details-table-action company-details-table-action--edit"
                        title="Edit"
                        aria-label="Edit settings"
                        onClick={openEditForm}
                      >
                        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                          <path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7" />
                          <path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z" />
                        </svg>
                      </button>
                      <button
                        type="button"
                        className="company-details-table-action company-details-table-action--delete"
                        title="Delete"
                        aria-label="Delete settings"
                        onClick={handleDeleteRow}
                      >
                        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                          <polyline points="3 6 5 6 21 6" />
                          <path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3-3V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2" />
                          <line x1="10" y1="11" x2="10" y2="17" />
                          <line x1="14" y1="11" x2="14" y2="17" />
                        </svg>
                      </button>
                    </div>
                  </div>
                ))
              )}
            </div>
          </div>
        )}
        {message ? <div className="settings-message">{message}</div> : null}
      </div>
    </>
  ) : (
    <div className="settings-inline-form-wrap">
      <div className="settings-modal company-details-modal" aria-labelledby="settings-form-title">
        <div className="settings-modal-header company-details-modal-header">
          <h2 id="settings-form-title" className="settings-modal-title company-details-modal-title">
            Company Settings
          </h2>
        </div>

        <form className="settings-modal-form company-details-modal-form" onSubmit={submitForm}>
          <div className="settings-modal-scroll company-details-modal-scroll">
            <div className="settings-modal-body company-details-modal-body">
              <div className="settings-section-card company-details-section-card">
                <div className="settings-fields-grid company-details-fields-grid">
                  <div className="settings-field company-details-field company-details-field--span2">
                    <label htmlFor="settings-due-date">DueDate</label>
                    <input
                      id="settings-due-date"
                      type="text"
                      value={formDueDate}
                      onChange={(e) => setFormDueDate(e.target.value)}
                      placeholder="Enter due date"
                      disabled={saving}
                    />
                  </div>
                </div>
              </div>

              {message ? <div className="settings-message">{message}</div> : null}
            </div>
          </div>

          <div className="settings-modal-footer company-details-modal-footer">
            <button
              type="button"
              className="settings-btn company-details-btn company-details-btn-secondary"
              onClick={closeForm}
              disabled={saving}
            >
              Cancel
            </button>
            <button
              type="submit"
              className="settings-btn company-details-btn company-details-btn-primary company-details-btn-submit-modal"
              disabled={saving}
            >
              {saving ? 'Submitting...' : 'Submit'}
            </button>
          </div>
        </form>
      </div>
    </div>
  );

  if (embedded) {
    return <section className="settings-hub-section">{listAndForm}</section>;
  }

  return (
    <div className={`settings-page company-details-page${showForm ? ' company-details-page-form-open' : ''}`}>
      <div className="settings-shell">
        {!showForm ? <div className="settings-list-card company-details-shell-box">{listAndForm}</div> : listAndForm}
      </div>
    </div>
  );
}

const formatNotificationDate = (value) => {
  const raw = String(value || '').trim();
  if (!raw) return '-';
  if (/^\d{1,2}$/.test(raw)) {
    const day = Number(raw);
    if (day >= 1 && day <= 31) {
      const date = new Date();
      date.setDate(day);
      return date.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });
    }
  }
  const date = new Date(raw);
  if (!Number.isNaN(date.getTime())) {
    return date.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });
  }
  return raw;
};

const REMINDER_LEVEL_LABELS = ['Level 1 Notification', 'Level 2 Notification'];

const NOTIFICATION_ICON_TONES = {
  dueDate: 'purple',
  pending: 'orange',
  reminder: 'blue',
  chro: 'green',
};

const NOTIFICATION_STATUS_OPTIONS = [
  { value: 'Active', label: 'Active' },
  { value: 'Paused', label: 'Paused' },
];

function normalizeRowStatus(value) {
  return String(value || '').trim().toLowerCase() === 'paused' ? 'Paused' : 'Active';
}

function buildNotificationRows({ dueDate, dueDateStatus, pendingRows, reminderRows, chroRows }) {
  const rows = [];

  if (String(dueDate || '').trim()) {
    rows.push({
      id: 'dueDate-settings',
      type: 'dueDate',
      iconTone: NOTIFICATION_ICON_TONES.dueDate,
      name: 'Due Date Notification',
      subtitle: 'Compliance Settings',
      triggerPrimary: formatSettingsDueDate(dueDate),
      triggerSecondary: 'Due date',
      emailPrimary: '-',
      emailSecondary: 'All sites notification',
      status: normalizeRowStatus(dueDateStatus),
      raw: { dueDate, status: normalizeRowStatus(dueDateStatus) },
    });
  }

  (pendingRows || []).forEach((row) => {
    rows.push({
      id: `pending-${row.id}`,
      type: 'pending',
      rowId: String(row.id || ''),
      iconTone: NOTIFICATION_ICON_TONES.pending,
      name: 'Pending Notification',
      subtitle: 'HR Compliance',
      triggerPrimary: formatNotificationDate(row.notificationDate),
      triggerSecondary: 'Notification date',
      emailPrimary: row.email || '-',
      emailSecondary: 'Notification email',
      status: normalizeRowStatus(row.status),
      raw: row,
    });
  });

  (reminderRows || [])
    .filter((row) => row?.id && (row.email || row.notificationDate))
    .forEach((row) => {
      const level = Number(row.level) || 1;
      rows.push({
        id: `reminder-${row.id}`,
        type: 'reminder',
        rowId: String(row.id || ''),
        iconTone: NOTIFICATION_ICON_TONES.reminder,
        name: REMINDER_LEVEL_LABELS[level - 1] || `Level ${level} Notification`,
        subtitle: 'HR Compliance',
        triggerPrimary: formatNotificationDate(row.notificationDate),
        triggerSecondary: `Level ${level} reminder`,
        emailPrimary: row.email || '-',
        emailSecondary: 'Reminder email',
        status: normalizeRowStatus(row.status),
        raw: row,
      });
    });

  (chroRows || []).forEach((row) => {
    rows.push({
      id: `chro-${row.id}`,
      type: 'chro',
      rowId: String(row.id || ''),
      iconTone: NOTIFICATION_ICON_TONES.chro,
      name: 'CHRO Notification',
      subtitle: row.name || 'Monthly Report',
      triggerPrimary: '28th of month',
      triggerSecondary: 'Scheduled',
      emailPrimary: row.email || '-',
      emailSecondary: row.name ? `${row.name}` : 'CHRO contact',
      status: normalizeRowStatus(row.status),
      raw: row,
    });
  });

  return rows;
}

function NotificationIcon({ tone = 'purple' }) {
  return (
    <span className={`notification-table-icon notification-table-icon--${tone}`} aria-hidden="true">
      <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8">
        <rect x="3" y="4.5" width="18" height="16" rx="2" />
        <path d="M7 2.5v4M17 2.5v4M3 9h18" />
      </svg>
    </span>
  );
}

function NotificationCell({ primary, secondary }) {
  return (
    <div className="notification-table-cell-stack">
      <span className="notification-table-cell-primary">{primary}</span>
      <span className="notification-table-cell-secondary">{secondary}</span>
    </div>
  );
}

const EMAIL_VALUE_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

const NOTIFICATION_TYPES = [
  { id: 'dueDate', label: 'Due Date Notification' },
  { id: 'pending', label: 'Pending Notification' },
  { id: 'reminder', label: 'Reminder Notification' },
  { id: 'chro', label: 'CHRO Notification' },
];

function emptyHubForm() {
  return {
    rowId: '',
    notificationType: 'dueDate',
    dueDate: '',
    email: '',
    notificationDate: '',
    level: 1,
    name: '',
    notificationStatus: 'Active',
  };
}

export default function Settings() {
  const [showForm, setShowForm] = useState(false);
  const [editing, setEditing] = useState(false);
  const [refreshKey, setRefreshKey] = useState(0);
  const [form, setForm] = useState(emptyHubForm);
  const [saving, setSaving] = useState(false);
  const [message, setMessage] = useState('');
  const [listMessage, setListMessage] = useState('');
  const [loading, setLoading] = useState(true);
  const [tableSearch, setTableSearch] = useState('');
  const [statusFilter, setStatusFilter] = useState('all');
  const [dueDate, setDueDate] = useState(() => String(readSettingsCache()?.dueDate || ''));
  const [dueDateStatus, setDueDateStatus] = useState(
    () => readSettingsCache()?.notificationStatus || 'Active'
  );
  const [pendingRows, setPendingRows] = useState([]);
  const [reminderRows, setReminderRows] = useState([]);
  const [chroRows, setChroRows] = useState([]);
  const [savedCompanyName, setSavedCompanyName] = useState(() => {
    const cached = readSettingsCache();
    return cached ? String(cached.companyName || '') : '';
  });

  const loadNotifications = useCallback(async () => {
    setLoading(true);
    setListMessage('');
    try {
      const [settingsData, pendingRes, reminderRes, chroRes] = await Promise.all([
        fetchSettings({ force: true }),
        fetch(`${API_BASE}/pending-notification`).then((res) => res.json()),
        fetch(`${API_BASE}/reminder-notification`).then((res) => res.json()),
        fetch(`${API_BASE}/chro`).then((res) => res.json()),
      ]);

      if (settingsData) {
        setDueDate(String(settingsData.dueDate || ''));
        setDueDateStatus(String(settingsData.notificationStatus || 'Active'));
        setSavedCompanyName(String(settingsData.companyName || ''));
        writeSettingsCache(settingsData);
      }

      if (pendingRes?.status === 'success') {
        setPendingRows(Array.isArray(pendingRes.data) ? pendingRes.data : []);
      } else {
        throw new Error(pendingRes?.message || 'Failed to load pending notifications.');
      }

      if (reminderRes?.status === 'success') {
        setReminderRows(Array.isArray(reminderRes.data) ? reminderRes.data : []);
      } else {
        throw new Error(reminderRes?.message || 'Failed to load reminder notifications.');
      }

      if (chroRes?.status === 'success') {
        setChroRows(Array.isArray(chroRes.data) ? chroRes.data : []);
      } else {
        throw new Error(chroRes?.message || 'Failed to load CHRO notifications.');
      }
    } catch (err) {
      setListMessage(err?.message || 'Unable to load notifications.');
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    loadNotifications();
  }, [loadNotifications, refreshKey]);

  const notificationRows = useMemo(
    () => buildNotificationRows({ dueDate, dueDateStatus, pendingRows, reminderRows, chroRows }),
    [dueDate, dueDateStatus, pendingRows, reminderRows, chroRows]
  );

  const filteredRows = useMemo(() => {
    const query = tableSearch.trim().toLowerCase();
    return notificationRows.filter((row) => {
      const matchesStatus =
        statusFilter === 'all' || row.status.toLowerCase() === statusFilter.toLowerCase();
      if (!matchesStatus) return false;
      if (!query) return true;
      const haystack = [
        row.name,
        row.subtitle,
        row.triggerPrimary,
        row.triggerSecondary,
        row.emailPrimary,
        row.emailSecondary,
        row.status,
      ]
        .join(' ')
        .toLowerCase();
      return haystack.includes(query);
    });
  }, [notificationRows, statusFilter, tableSearch]);

  const openAddForm = () => {
    setEditing(false);
    setShowForm(true);
    setForm(emptyHubForm());
    setMessage('');
  };

  const openEditForm = (row) => {
    setEditing(true);
    setShowForm(true);
    setMessage('');

    if (row.type === 'dueDate') {
      setForm({
        ...emptyHubForm(),
        notificationType: 'dueDate',
        dueDate: formatDueDateForInput(row.raw?.dueDate),
        notificationStatus: normalizeRowStatus(row.raw?.status || row.status),
      });
      return;
    }

    if (row.type === 'pending') {
      setForm({
        ...emptyHubForm(),
        rowId: row.rowId,
        notificationType: 'pending',
        email: String(row.raw?.email || ''),
        notificationDate: toDateInputValue(row.raw?.notificationDate),
        notificationStatus: normalizeRowStatus(row.raw?.status || row.status),
      });
      return;
    }

    if (row.type === 'reminder') {
      setForm({
        ...emptyHubForm(),
        rowId: row.rowId,
        notificationType: 'reminder',
        level: Number(row.raw?.level) || 1,
        email: String(row.raw?.email || ''),
        notificationDate: toDateInputValue(row.raw?.notificationDate),
        notificationStatus: normalizeRowStatus(row.raw?.status || row.status),
      });
      return;
    }

    setForm({
      ...emptyHubForm(),
      rowId: row.rowId,
      notificationType: 'chro',
      name: String(row.raw?.name || ''),
      email: String(row.raw?.email || ''),
      notificationStatus: normalizeRowStatus(row.raw?.status || row.status),
    });
  };

  const closeForm = () => {
    setShowForm(false);
    setEditing(false);
    setForm(emptyHubForm());
    setMessage('');
  };

  const bumpRefresh = () => setRefreshKey((key) => key + 1);

  const handleDeleteRow = async (row) => {
    try {
      if (row.type === 'dueDate') {
        const res = await fetch(`${API_BASE}/settings`, { method: 'DELETE' });
        const data = await res.json();
        if (!res.ok || data?.status !== 'success') {
          throw new Error(data?.message || 'Failed to delete due date notification');
        }
        clearSettingsCache();
      } else if (row.type === 'pending') {
        const res = await fetch(`${API_BASE}/pending-notification?rowId=${encodeURIComponent(row.rowId)}`, {
          method: 'DELETE',
        });
        const data = await res.json();
        if (!res.ok || data?.status !== 'success') {
          throw new Error(data?.message || 'Failed to delete pending notification');
        }
      } else if (row.type === 'reminder') {
        const res = await fetch(`${API_BASE}/reminder-notification?rowId=${encodeURIComponent(row.rowId)}`, {
          method: 'DELETE',
        });
        const data = await res.json();
        if (!res.ok || data?.status !== 'success') {
          throw new Error(data?.message || 'Failed to delete reminder notification');
        }
      } else if (row.type === 'chro') {
        const res = await fetch(`${API_BASE}/chro?rowId=${encodeURIComponent(row.rowId)}`, {
          method: 'DELETE',
        });
        const data = await res.json();
        if (!res.ok || data?.status !== 'success') {
          throw new Error(data?.message || 'Failed to delete CHRO notification');
        }
      }
      setListMessage('Deleted successfully.');
      bumpRefresh();
    } catch (err) {
      setListMessage(err?.message || 'Unable to delete notification.');
    }
  };

  const submitForm = async (event) => {
    event.preventDefault();
    setSaving(true);
    setMessage('');

    try {
      if (form.notificationType === 'dueDate') {
        const dueDateValue = form.dueDate.trim();
        if (!dueDateValue) throw new Error('Due Date is required.');
        const formData = new FormData();
        formData.append('companyName', savedCompanyName.trim());
        formData.append('dueDate', dueDateValue);
        formData.append('notificationStatus', form.notificationStatus);

        const res = await fetch(`${API_BASE}/settings`, { method: 'PUT', body: formData });
        const data = await res.json();
        if (!res.ok || data?.status !== 'success') {
          throw new Error(data?.message || 'Failed to save due date notification');
        }
        writeSettingsCache(data?.data);
        if (data?.data?.companyName) setSavedCompanyName(String(data.data.companyName));
      } else if (form.notificationType === 'pending') {
        const email = form.email.trim();
        const notificationDate = form.notificationDate.trim();
        if (!email) throw new Error('Email is required.');
        if (!EMAIL_VALUE_RE.test(email)) throw new Error('Enter a valid email address.');
        if (!notificationDate) throw new Error('Notification Date is required.');

        const formData = new FormData();
        if (form.rowId) formData.append('rowId', form.rowId);
        formData.append('email', email);
        formData.append('notificationDate', notificationDate);
        formData.append('notificationStatus', form.notificationStatus);

        const res = await fetch(`${API_BASE}/pending-notification`, { method: 'PUT', body: formData });
        const data = await res.json();
        if (!res.ok || data?.status !== 'success') {
          throw new Error(data?.message || 'Failed to save pending notification record');
        }
      } else if (form.notificationType === 'reminder') {
        const email = form.email.trim();
        const notificationDate = form.notificationDate.trim();
        const level = Number(form.level);
        if (!level || level < 1 || level > 2) throw new Error('Invalid notification level.');
        if (!email) throw new Error('Email is required.');
        if (!EMAIL_VALUE_RE.test(email)) throw new Error('Enter a valid email address.');
        if (!notificationDate) throw new Error('Notification Date is required.');

        const formData = new FormData();
        if (form.rowId) formData.append('rowId', form.rowId);
        formData.append('level', String(level));
        formData.append('email', email);
        formData.append('notificationDate', notificationDate);
        formData.append('notificationStatus', form.notificationStatus);

        const res = await fetch(`${API_BASE}/reminder-notification`, { method: 'PUT', body: formData });
        const data = await res.json();
        if (!res.ok || data?.status !== 'success') {
          throw new Error(data?.message || 'Failed to save reminder notification record');
        }
      } else if (form.notificationType === 'chro') {
        const name = form.name.trim();
        const email = form.email.trim();
        if (!name) throw new Error('CHRO Name is required.');
        if (!email) throw new Error('Email is required.');
        if (!EMAIL_VALUE_RE.test(email)) throw new Error('Enter a valid email address.');

        const formData = new FormData();
        if (form.rowId) formData.append('rowId', form.rowId);
        formData.append('name', name);
        formData.append('email', email);
        formData.append('notificationStatus', form.notificationStatus);

        const res = await fetch(`${API_BASE}/chro`, { method: 'PUT', body: formData });
        const data = await res.json();
        if (!res.ok || data?.status !== 'success') {
          throw new Error(data?.message || 'Failed to save CHRO record');
        }
      }

      closeForm();
      bumpRefresh();
    } catch (err) {
      setMessage(err?.message || 'Unable to save notification.');
    } finally {
      setSaving(false);
    }
  };

  const selectedType = NOTIFICATION_TYPES.find((item) => item.id === form.notificationType);
  const reminderLevelTitle =
    REMINDER_LEVEL_LABELS[Number(form.level) - 1] || `Level ${form.level} Notification`;
  const formTitle = editing
    ? `Edit ${selectedType?.label || 'Notification'}`
    : `Add ${selectedType?.label || 'Notification'}`;

  const addForm = (
    <div className="settings-inline-form-wrap">
      <div className="settings-modal company-details-modal" aria-labelledby="notification-hub-form-title">
        <div className="settings-modal-header company-details-modal-header">
          <h2 id="notification-hub-form-title" className="settings-modal-title company-details-modal-title">
            {formTitle}
          </h2>
        </div>

        <form className="settings-modal-form company-details-modal-form" onSubmit={submitForm}>
          <div className="settings-modal-scroll company-details-modal-scroll">
            <div className="settings-modal-body company-details-modal-body">
              <div className="settings-section-card company-details-section-card">
                <div className="settings-fields-grid company-details-fields-grid">
                  <div className="settings-field company-details-field company-details-field--span2">
                    <label htmlFor="notification-type">
                      Notification Type <span className="required">*</span>
                    </label>
                    <select
                      id="notification-type"
                      value={form.notificationType}
                      onChange={(e) => {
                        const notificationType = e.target.value;
                        setForm({
                          ...emptyHubForm(),
                          notificationType,
                        });
                        setMessage('');
                      }}
                      disabled={saving || editing}
                    >
                      {NOTIFICATION_TYPES.map((type) => (
                        <option key={type.id} value={type.id}>
                          {type.label}
                        </option>
                      ))}
                    </select>
                  </div>

                  <div className="settings-field company-details-field company-details-field--span2">
                    <label htmlFor="hub-notification-status">
                      Status <span className="required">*</span>
                    </label>
                    <select
                      id="hub-notification-status"
                      value={form.notificationStatus}
                      onChange={(e) =>
                        setForm((prev) => ({ ...prev, notificationStatus: e.target.value }))
                      }
                      disabled={saving}
                    >
                      {NOTIFICATION_STATUS_OPTIONS.map((option) => (
                        <option key={option.value} value={option.value}>
                          {option.label}
                        </option>
                      ))}
                    </select>
                  </div>

                  {form.notificationType === 'dueDate' ? (
                    <div className="settings-field company-details-field company-details-field--span2">
                      <label htmlFor="hub-due-date">
                        Due Date <span className="required">*</span>
                      </label>
                      <input
                        id="hub-due-date"
                        type="text"
                        value={form.dueDate}
                        onChange={(e) => setForm((prev) => ({ ...prev, dueDate: e.target.value }))}
                        placeholder="Enter due date"
                        disabled={saving}
                      />
                    </div>
                  ) : null}

                  {form.notificationType === 'pending' ? (
                    <>
                      <div className="settings-field company-details-field company-details-field--span2">
                        <label htmlFor="hub-pending-email">
                          Email <span className="required">*</span>
                        </label>
                        <input
                          id="hub-pending-email"
                          type="email"
                          value={form.email}
                          onChange={(e) => setForm((prev) => ({ ...prev, email: e.target.value }))}
                          placeholder="Enter email"
                          disabled={saving}
                        />
                      </div>
                      <div className="settings-field company-details-field company-details-field--span2">
                        <label htmlFor="hub-pending-date">
                          Notification Date <span className="required">*</span>
                        </label>
                        <input
                          id="hub-pending-date"
                          type="date"
                          value={form.notificationDate}
                          onChange={(e) => setForm((prev) => ({ ...prev, notificationDate: e.target.value }))}
                          disabled={saving}
                        />
                      </div>
                    </>
                  ) : null}

                  {form.notificationType === 'reminder' ? (
                    <>
                      <div className="settings-field company-details-field company-details-field--span2">
                        <label htmlFor="hub-reminder-level">
                          Notification Level <span className="required">*</span>
                        </label>
                        <select
                          id="hub-reminder-level"
                          value={form.level}
                          onChange={(e) =>
                            setForm((prev) => ({ ...prev, level: Number(e.target.value) || 1 }))
                          }
                          disabled={saving || editing}
                        >
                          {REMINDER_LEVEL_LABELS.map((label, index) => (
                            <option key={label} value={index + 1}>
                              {label}
                            </option>
                          ))}
                        </select>
                      </div>
                      <div className="settings-field company-details-field company-details-field--span2">
                        <label htmlFor="hub-reminder-email">
                          Email <span className="required">*</span>
                        </label>
                        <input
                          id="hub-reminder-email"
                          type="email"
                          value={form.email}
                          onChange={(e) => setForm((prev) => ({ ...prev, email: e.target.value }))}
                          placeholder="Enter email"
                          disabled={saving}
                        />
                      </div>
                      <div className="settings-field company-details-field company-details-field--span2">
                        <label htmlFor="hub-reminder-date">
                          Notification Date <span className="required">*</span>
                        </label>
                        <input
                          id="hub-reminder-date"
                          type="date"
                          value={form.notificationDate}
                          onChange={(e) =>
                            setForm((prev) => ({ ...prev, notificationDate: e.target.value }))
                          }
                          disabled={saving}
                        />
                      </div>
                      <p className="settings-list-empty" style={{ gridColumn: '1 / -1', margin: 0 }}>
                        {reminderLevelTitle}: HR Compliance email uses Reports data for the current month —
                        only Pending and Yet to Submit forms are included.
                      </p>
                    </>
                  ) : null}

                  {form.notificationType === 'chro' ? (
                    <>
                      <div className="settings-field company-details-field company-details-field--span2">
                        <label htmlFor="hub-chro-name">
                          CHRO Name <span className="required">*</span>
                        </label>
                        <input
                          id="hub-chro-name"
                          type="text"
                          value={form.name}
                          onChange={(e) => setForm((prev) => ({ ...prev, name: e.target.value }))}
                          placeholder="Enter CHRO name"
                          disabled={saving}
                        />
                      </div>
                      <div className="settings-field company-details-field company-details-field--span2">
                        <label htmlFor="hub-chro-email">
                          Email <span className="required">*</span>
                        </label>
                        <input
                          id="hub-chro-email"
                          type="email"
                          value={form.email}
                          onChange={(e) => setForm((prev) => ({ ...prev, email: e.target.value }))}
                          placeholder="Enter email"
                          disabled={saving}
                        />
                      </div>
                    </>
                  ) : null}
                </div>
              </div>

              {message ? <div className="settings-message">{message}</div> : null}
            </div>
          </div>

          <div className="settings-modal-footer company-details-modal-footer">
            <button
              type="button"
              className="settings-btn company-details-btn company-details-btn-secondary"
              onClick={closeForm}
              disabled={saving}
            >
              Cancel
            </button>
            <button
              type="submit"
              className="settings-btn company-details-btn company-details-btn-primary company-details-btn-submit-modal"
              disabled={saving}
            >
              {saving ? 'Submitting...' : 'Submit'}
            </button>
          </div>
        </form>
      </div>
    </div>
  );

  if (showForm) {
    return (
      <div className="settings-page company-details-page company-details-page-form-open">
        <div className="settings-shell">{addForm}</div>
      </div>
    );
  }

  return (
    <div className="settings-page company-details-page">
      <div className="settings-shell">
        <div className="settings-list-card company-details-shell-box">
          <div className="settings-page-heading settings-page-heading--row">
            <h1 className="settings-page-title">Email Notification</h1>
            <div className="settings-page-actions">
              <button
                type="button"
                className="company-details-btn company-details-btn-add"
                onClick={openAddForm}
              >
                <span className="company-details-btn-add-icon">+</span> Add
              </button>
            </div>
          </div>

          <div className="settings-list-body">
            <div className="company-details-inner-toolbar notification-table-toolbar">
              <div className="company-details-table-search company-details-table-search--flex">
                <svg
                  className="company-details-table-search-icon"
                  width="18"
                  height="18"
                  viewBox="0 0 24 24"
                  fill="none"
                  stroke="currentColor"
                  strokeWidth="2"
                  aria-hidden
                >
                  <circle cx="11" cy="11" r="8" />
                  <path d="m21 21-4.35-4.35" />
                </svg>
                <input
                  type="search"
                  className="company-details-table-search-input"
                  placeholder="Search notifications..."
                  value={tableSearch}
                  onChange={(e) => setTableSearch(e.target.value)}
                  aria-label="Search notifications"
                />
              </div>
              <div className="company-details-inner-toolbar-right">
                <div className="company-details-sort-wrap notification-table-status-filter">
                  <svg
                    className="company-details-sort-icon"
                    width="16"
                    height="16"
                    viewBox="0 0 24 24"
                    fill="none"
                    stroke="currentColor"
                    strokeWidth="2"
                    aria-hidden
                  >
                    <polygon points="22 3 2 3 10 12.46 10 19 14 21 14 12.46 22 3" />
                  </svg>
                  <select
                    className="company-details-sort-select"
                    value={statusFilter}
                    onChange={(e) => setStatusFilter(e.target.value)}
                    aria-label="Filter by status"
                  >
                    <option value="all">Status: All</option>
                    <option value="active">Status: Active</option>
                    <option value="paused">Status: Paused</option>
                  </select>
                </div>
                <button
                  type="button"
                  className="company-details-btn company-details-btn-export notification-table-refresh-btn"
                  onClick={bumpRefresh}
                  title="Refresh"
                  aria-label="Refresh notifications"
                >
                  <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" aria-hidden>
                    <polyline points="23 4 23 10 17 10" />
                    <polyline points="1 20 1 14 7 14" />
                    <path d="M3.51 9a9 9 0 0 1 14.13-3.36L23 10M1 14l5.36 5.36A9 9 0 0 0 20.49 15" />
                  </svg>
                  Refresh
                </button>
              </div>
            </div>

            {listMessage ? <div className="settings-message">{listMessage}</div> : null}

            <div className="company-details-table-wrap">
              <table className="company-details-data-table notification-data-table">
                <thead>
                  <tr>
                    <th scope="col">Notification Name</th>
                    <th scope="col">Trigger</th>
                    <th scope="col">Email</th>
                    <th scope="col">Status</th>
                    <th className="company-details-th-actions" scope="col">
                      Actions
                    </th>
                  </tr>
                </thead>
                <tbody>
                  {loading ? (
                    <tr>
                      <td colSpan={5} className="notification-table-empty-cell">
                        Loading...
                      </td>
                    </tr>
                  ) : filteredRows.length === 0 ? (
                    <tr>
                      <td colSpan={5} className="notification-table-empty-cell">
                        {notificationRows.length === 0
                          ? <>No notifications added yet. Click <strong>+ Add</strong> to open form.</>
                          : 'No notifications match your search or filter.'}
                      </td>
                    </tr>
                  ) : (
                    filteredRows.map((row) => (
                      <tr key={row.id}>
                        <td>
                          <div className="notification-table-name-cell">
                            <NotificationIcon tone={row.iconTone} />
                            <NotificationCell primary={row.name} secondary={row.subtitle} />
                          </div>
                        </td>
                        <td>
                          <NotificationCell primary={row.triggerPrimary} secondary={row.triggerSecondary} />
                        </td>
                        <td>
                          <NotificationCell primary={row.emailPrimary} secondary={row.emailSecondary} />
                        </td>
                        <td>
                          <span
                            className={`notification-status-badge notification-status-badge--${row.status.toLowerCase()}`}
                          >
                            {row.status}
                          </span>
                        </td>
                        <td className="company-details-td-actions">
                          <div className="notification-table-actions">
                            <button
                              type="button"
                              className="company-details-table-action company-details-table-action--edit notification-table-edit-btn"
                              title="Edit"
                              aria-label={`Edit ${row.name}`}
                              onClick={() => openEditForm(row)}
                            >
                              <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                                <path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7" />
                                <path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z" />
                              </svg>
                            </button>
                            <button
                              type="button"
                              className="company-details-table-action company-details-table-action--delete notification-table-edit-btn"
                              title="Delete"
                              aria-label={`Delete ${row.name}`}
                              onClick={() => handleDeleteRow(row)}
                            >
                              <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                                <polyline points="3 6 5 6 21 6" />
                                <path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3-3V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2" />
                                <line x1="10" y1="11" x2="10" y2="17" />
                                <line x1="14" y1="11" x2="14" y2="17" />
                              </svg>
                            </button>
                          </div>
                        </td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}
