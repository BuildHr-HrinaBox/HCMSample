import React, { useCallback, useEffect, useState } from 'react';
import './CompanyDetails.css';
import './Settings.css';

const API_BASE = '/server/settings_function';
const EMAIL_VALUE_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

function emptyForm() {
  return { rowId: '', email: '', notificationDate: '' };
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

export default function PendingNotification({ embedded = false, hideAddButton = false, refreshKey = 0 }) {
  const [showForm, setShowForm] = useState(false);
  const [rows, setRows] = useState([]);
  const [form, setForm] = useState(emptyForm);
  const [saving, setSaving] = useState(false);
  const [loading, setLoading] = useState(true);
  const [message, setMessage] = useState('');

  const loadRows = useCallback(async () => {
    try {
      const res = await fetch(`${API_BASE}/pending-notification`);
      const data = await res.json();
      if (!res.ok || data?.status !== 'success') {
        throw new Error(data?.message || 'Failed to load pending notification records');
      }
      setRows(Array.isArray(data?.data) ? data.data : []);
    } catch (err) {
      setMessage(err?.message || 'Unable to load pending notification records.');
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    loadRows();
  }, [loadRows, refreshKey]);

  const openAddForm = () => {
    setShowForm(true);
    setForm(emptyForm());
    setMessage('');
  };

  const openEditForm = (row) => {
    setShowForm(true);
    setForm({
      rowId: String(row?.id || ''),
      email: String(row?.email || ''),
      notificationDate: toDateInputValue(row?.notificationDate),
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
      formData.append('email', email);
      formData.append('notificationDate', notificationDate);

      const res = await fetch(`${API_BASE}/pending-notification`, {
        method: 'PUT',
        body: formData,
      });
      const data = await res.json();
      if (!res.ok || data?.status !== 'success') {
        throw new Error(data?.message || 'Failed to save pending notification record');
      }
      setShowForm(false);
      setForm(emptyForm());
      await loadRows();
    } catch (err) {
      setMessage(err?.message || 'Unable to save pending notification record.');
    } finally {
      setSaving(false);
    }
  };

  const handleDeleteRow = async (row) => {
    const rowId = String(row?.id || '').trim();
    if (!rowId) return;
    try {
      const res = await fetch(`${API_BASE}/pending-notification?rowId=${encodeURIComponent(rowId)}`, {
        method: 'DELETE',
      });
      const data = await res.json();
      if (!res.ok || data?.status !== 'success') {
        throw new Error(data?.message || 'Failed to delete pending notification record');
      }
      setMessage('Deleted successfully.');
      await loadRows();
    } catch (err) {
      setMessage(err?.message || 'Unable to delete pending notification record.');
    }
  };

  const showLoadingPlaceholder = loading && rows.length === 0;

  const listAndForm = !showForm ? (
          <>
            <div className="settings-page-heading settings-page-heading--row">
              {embedded ? (
                <h2 className="settings-hub-section-title">Pending Notification</h2>
              ) : (
                <h1 className="settings-page-title">Pending Notification</h1>
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
                    {rows.length === 0 ? (
                      <p className="settings-list-empty">
                        {hideAddButton
                          ? <>No pending notification added yet. Click <strong>+ Add</strong> above to open form.</>
                          : 'No pending notification added yet. Click Add to open form.'}
                      </p>
                    ) : (
                      rows.map((row) => (
                        <div
                          key={row.id || `${row.email}-${row.notificationDate}`}
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
                              className="company-details-table-action company-details-table-action--view"
                              title="View"
                              aria-label="View pending notification"
                              onClick={() => openEditForm(row)}
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
                              aria-label="Edit pending notification"
                              onClick={() => openEditForm(row)}
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
                              aria-label="Delete pending notification"
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
            <div className="settings-modal company-details-modal" aria-labelledby="pending-notification-form-title">
              <div className="settings-modal-header company-details-modal-header">
                <h2 id="pending-notification-form-title" className="settings-modal-title company-details-modal-title">
                  {form.rowId ? 'Edit Pending Notification' : 'Add Pending Notification'}
                </h2>
              </div>

              <form className="settings-modal-form company-details-modal-form" onSubmit={submitForm}>
                <div className="settings-modal-scroll company-details-modal-scroll">
                  <div className="settings-modal-body company-details-modal-body">
                    <div className="settings-section-card company-details-section-card">
                      <div className="settings-fields-grid company-details-fields-grid">
                        <div className="settings-field company-details-field company-details-field--span2">
                          <label htmlFor="pending-notification-email">
                            Email <span className="required">*</span>
                          </label>
                          <input
                            id="pending-notification-email"
                            type="email"
                            value={form.email}
                            onChange={(e) => setForm((prev) => ({ ...prev, email: e.target.value }))}
                            placeholder="Enter email"
                            disabled={saving}
                          />
                        </div>
                        <div className="settings-field company-details-field company-details-field--span2">
                          <label htmlFor="pending-notification-date">
                            Notification Date <span className="required">*</span>
                          </label>
                          <input
                            id="pending-notification-date"
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
