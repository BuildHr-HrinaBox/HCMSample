import React, { useCallback, useEffect, useState } from 'react';
import './CompanyDetails.css';
import './Settings.css';

const API_BASE = '/server/settings_function';
const EMAIL_VALUE_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
const LEVEL_LABELS = ['Level 1 Notification', 'Level 2 Notification'];

function emptyForm() {
  return { rowId: '', level: 1, email: '', notificationDate: '' };
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

export default function ReminderNotification({ embedded = false, refreshKey = 0 }) {
  const [showForm, setShowForm] = useState(false);
  const [rows, setRows] = useState([]);
  const [form, setForm] = useState(emptyForm);
  const [saving, setSaving] = useState(false);
  const [loading, setLoading] = useState(true);
  const [message, setMessage] = useState('');

  const loadRows = useCallback(async () => {
    try {
      const res = await fetch(`${API_BASE}/reminder-notification`);
      const data = await res.json();
      if (!res.ok || data?.status !== 'success') {
        throw new Error(data?.message || 'Failed to load reminder notification records');
      }
      setRows(Array.isArray(data?.data) ? data.data : []);
    } catch (err) {
      setMessage(err?.message || 'Unable to load reminder notification records.');
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    loadRows();
  }, [loadRows, refreshKey]);

  const openAddForm = (level) => {
    const existing = rows.find((row) => Number(row.level) === Number(level));
    setShowForm(true);
    setForm({
      rowId: String(existing?.id || ''),
      level: Number(level) || 1,
      email: String(existing?.email || ''),
      notificationDate: toDateInputValue(existing?.notificationDate),
    });
    setMessage('');
  };

  const closeForm = () => {
    setShowForm(false);
    setForm(emptyForm());
    setMessage('');
  };

  const submitForm = async (event) => {
    event.preventDefault();
    const email = form.email.trim();
    const notificationDate = form.notificationDate.trim();
    const level = Number(form.level);
    if (!level || level < 1 || level > 2) {
      setMessage('Invalid notification level.');
      return;
    }
    if (!email) {
      setMessage('Email is required.');
      return;
    }
    if (!EMAIL_VALUE_RE.test(email)) {
      setMessage('Enter a valid email address.');
      return;
    }
    if (!notificationDate) {
      setMessage('Notification Date is required.');
      return;
    }

    setSaving(true);
    setMessage('');
    try {
      const formData = new FormData();
      if (form.rowId) formData.append('rowId', form.rowId);
      formData.append('level', String(level));
      formData.append('email', email);
      formData.append('notificationDate', notificationDate);

      const res = await fetch(`${API_BASE}/reminder-notification`, {
        method: 'PUT',
        body: formData,
      });
      const data = await res.json();
      if (!res.ok || data?.status !== 'success') {
        throw new Error(data?.message || 'Failed to save reminder notification record');
      }
      setShowForm(false);
      setForm(emptyForm());
      await loadRows();
    } catch (err) {
      setMessage(err?.message || 'Unable to save reminder notification record.');
    } finally {
      setSaving(false);
    }
  };

  const handleDeleteRow = async (row) => {
    const rowId = String(row?.id || '').trim();
    if (!rowId) return;
    try {
      const res = await fetch(`${API_BASE}/reminder-notification?rowId=${encodeURIComponent(rowId)}`, {
        method: 'DELETE',
      });
      const data = await res.json();
      if (!res.ok || data?.status !== 'success') {
        throw new Error(data?.message || 'Failed to delete reminder notification record');
      }
      setMessage('Deleted successfully.');
      await loadRows();
    } catch (err) {
      setMessage(err?.message || 'Unable to delete reminder notification record.');
    }
  };

  const showLoadingPlaceholder = loading && rows.length === 0;
  const levelTitle = LEVEL_LABELS[Number(form.level) - 1] || `Level ${form.level} Notification`;

  const listAndForm = !showForm ? (
          <>
            <div className="settings-page-heading settings-page-heading--row">
              {embedded ? (
                <h2 className="settings-hub-section-title">Reminder Notification</h2>
              ) : (
                <h1 className="settings-page-title">Reminder Notification</h1>
              )}
            </div>
            <div className={embedded ? 'settings-hub-section-body' : 'settings-list-body'}>
              <p className="settings-list-empty" style={{ marginBottom: 16 }}>
                Configure up to 2 reminder levels. On each level&apos;s notification date, an HR Compliance email
                is sent using <strong>Reports</strong> data for the current month — only <strong>Pending</strong> and{' '}
                <strong>Yet to Submit</strong> forms are included.
              </p>
              {showLoadingPlaceholder ? (
                <p className="settings-list-empty">Loading...</p>
              ) : (
                <div className="company-details-table-wrap">
                  <div className="settings-due-date-list">
                    {LEVEL_LABELS.map((label, index) => {
                      const level = index + 1;
                      const row = rows.find((item) => Number(item.level) === level) || { level };
                      const configured = Boolean(row?.id && (row.email || row.notificationDate));
                      return (
                        <div
                          key={level}
                          className="settings-due-date-card"
                        >
                          <div className="settings-due-date-icon" aria-hidden="true">
                            <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8">
                              <rect x="3" y="4.5" width="18" height="16" rx="2" />
                              <path d="M7 2.5v4M17 2.5v4M3 9h18" />
                            </svg>
                          </div>
                          <div className="settings-chro-card-fields">
                            <div className="settings-due-date-info">
                              <strong>{label}</strong>
                              <span>{configured ? 'Configured' : 'Not configured'}</span>
                            </div>
                            <div className="settings-due-date-info">
                              <strong>Email</strong>
                              <span>{row.email || '-'}</span>
                            </div>
                            <div className="settings-due-date-info">
                              <strong>Notification Date</strong>
                              <span>{formatNotificationDate(row.notificationDate)}</span>
                            </div>
                          </div>
                          <div className="settings-due-date-actions">
                            <button
                              type="button"
                              className="company-details-table-action company-details-table-action--edit"
                              title={configured ? 'Edit' : 'Add'}
                              aria-label={configured ? `Edit ${label}` : `Add ${label}`}
                              onClick={() => openAddForm(level)}
                            >
                              <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                                <path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7" />
                                <path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z" />
                              </svg>
                            </button>
                            {configured ? (
                              <button
                                type="button"
                                className="company-details-table-action company-details-table-action--delete"
                                title="Delete"
                                aria-label={`Delete ${label}`}
                                onClick={() => handleDeleteRow(row)}
                              >
                                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                                  <polyline points="3 6 5 6 21 6" />
                                  <path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3-3V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2" />
                                  <line x1="10" y1="11" x2="10" y2="17" />
                                  <line x1="14" y1="11" x2="14" y2="17" />
                                </svg>
                              </button>
                            ) : null}
                          </div>
                        </div>
                      );
                    })}
                  </div>
                </div>
              )}
              {message ? <div className="settings-message">{message}</div> : null}
            </div>
          </>
        ) : (
          <div className="settings-inline-form-wrap">
            <div className="settings-modal company-details-modal" aria-labelledby="reminder-notification-form-title">
              <div className="settings-modal-header company-details-modal-header">
                <h2 id="reminder-notification-form-title" className="settings-modal-title company-details-modal-title">
                  {form.rowId ? `Edit ${levelTitle}` : `Add ${levelTitle}`}
                </h2>
              </div>

              <form className="settings-modal-form company-details-modal-form" onSubmit={submitForm}>
                <div className="settings-modal-scroll company-details-modal-scroll">
                  <div className="settings-modal-body company-details-modal-body">
                    <div className="settings-section-card company-details-section-card">
                      <div className="settings-fields-grid company-details-fields-grid">
                        <div className="settings-field company-details-field company-details-field--span2">
                          <label htmlFor="reminder-notification-level">Notification Level</label>
                          <input
                            id="reminder-notification-level"
                            type="text"
                            value={levelTitle}
                            readOnly
                            disabled
                          />
                        </div>
                        <div className="settings-field company-details-field company-details-field--span2">
                          <label htmlFor="reminder-notification-email">
                            Email <span className="required">*</span>
                          </label>
                          <input
                            id="reminder-notification-email"
                            type="email"
                            value={form.email}
                            onChange={(e) => setForm((prev) => ({ ...prev, email: e.target.value }))}
                            placeholder="Enter email"
                            disabled={saving}
                          />
                        </div>
                        <div className="settings-field company-details-field company-details-field--span2">
                          <label htmlFor="reminder-notification-date">
                            Notification Date <span className="required">*</span>
                          </label>
                          <input
                            id="reminder-notification-date"
                            type="date"
                            value={form.notificationDate}
                            onChange={(e) => setForm((prev) => ({ ...prev, notificationDate: e.target.value }))}
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
