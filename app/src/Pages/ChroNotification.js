import React, { useCallback, useEffect, useState } from 'react';
import './CompanyDetails.css';
import './Settings.css';

const API_BASE = '/server/settings_function';
const EMAIL_VALUE_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

function emptyForm() {
  return { rowId: '', name: '', email: '' };
}

export default function ChroNotification({ embedded = false, hideAddButton = false, refreshKey = 0 }) {
  const [showForm, setShowForm] = useState(false);
  const [rows, setRows] = useState([]);
  const [form, setForm] = useState(emptyForm);
  const [saving, setSaving] = useState(false);
  const [loading, setLoading] = useState(true);
  const [message, setMessage] = useState('');

  const loadRows = useCallback(async () => {
    try {
      const res = await fetch(`${API_BASE}/chro`);
      const data = await res.json();
      if (!res.ok || data?.status !== 'success') {
        throw new Error(data?.message || 'Failed to load CHRO records');
      }
      setRows(Array.isArray(data?.data) ? data.data : []);
    } catch (err) {
      setMessage(err?.message || 'Unable to load CHRO records.');
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
      name: String(row?.name || ''),
      email: String(row?.email || ''),
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
    const name = form.name.trim();
    const email = form.email.trim();
    if (!name) {
      setMessage('CHRO Name is required.');
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

    setSaving(true);
    setMessage('');
    try {
      const formData = new FormData();
      if (form.rowId) formData.append('rowId', form.rowId);
      formData.append('name', name);
      formData.append('email', email);

      const res = await fetch(`${API_BASE}/chro`, {
        method: 'PUT',
        body: formData,
      });
      const data = await res.json();
      if (!res.ok || data?.status !== 'success') {
        throw new Error(data?.message || 'Failed to save CHRO record');
      }
      setShowForm(false);
      setForm(emptyForm());
      await loadRows();
    } catch (err) {
      setMessage(err?.message || 'Unable to save CHRO record.');
    } finally {
      setSaving(false);
    }
  };

  const handleDeleteRow = async (row) => {
    const rowId = String(row?.id || '').trim();
    if (!rowId) return;
    try {
      const res = await fetch(`${API_BASE}/chro?rowId=${encodeURIComponent(rowId)}`, {
        method: 'DELETE',
      });
      const data = await res.json();
      if (!res.ok || data?.status !== 'success') {
        throw new Error(data?.message || 'Failed to delete CHRO record');
      }
      setMessage('Deleted successfully.');
      await loadRows();
    } catch (err) {
      setMessage(err?.message || 'Unable to delete CHRO record.');
    }
  };

  const showLoadingPlaceholder = loading && rows.length === 0;

  const listAndForm = !showForm ? (
          <>
            <div className="settings-page-heading settings-page-heading--row">
              {embedded ? (
                <h2 className="settings-hub-section-title">CHRO Notification</h2>
              ) : (
                <h1 className="settings-page-title">CHRO Notification</h1>
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
              <p className="settings-list-empty" style={{ marginBottom: 16 }}>
                Configure CHRO contact details. On the 28th of each month, an email is sent with the{' '}
                <strong>Approved Sites</strong> report and <strong>Pending Forms</strong> from Reports data for the
                current month.
              </p>
              {showLoadingPlaceholder ? (
                <p className="settings-list-empty">Loading...</p>
              ) : (
                <div className="company-details-table-wrap">
                  <div className="settings-due-date-list">
                    {rows.length === 0 ? (
                      <p className="settings-list-empty">
                        {hideAddButton
                          ? <>No CHRO added yet. Click <strong>+ Add</strong> above to open form.</>
                          : 'No CHRO added yet. Click Add to open form.'}
                      </p>
                    ) : (
                      rows.map((row) => (
                        <div key={row.id || `${row.name}-${row.email}`} className="settings-due-date-card">
                          <div className="settings-due-date-icon" aria-hidden="true">
                            <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8">
                              <path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2" />
                              <circle cx="12" cy="7" r="4" />
                            </svg>
                          </div>
                          <div className="settings-chro-card-fields">
                            <div className="settings-due-date-info">
                              <strong>CHRO Name</strong>
                              <span>{row.name || '-'}</span>
                            </div>
                            <div className="settings-due-date-info">
                              <strong>Email</strong>
                              <span>{row.email || '-'}</span>
                            </div>
                          </div>
                          <div className="settings-due-date-actions">
                            <button
                              type="button"
                              className="company-details-table-action company-details-table-action--view"
                              title="View"
                              aria-label="View CHRO"
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
                              aria-label="Edit CHRO"
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
                              aria-label="Delete CHRO"
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
            <div className="settings-modal company-details-modal" aria-labelledby="chro-form-title">
              <div className="settings-modal-header company-details-modal-header">
                <h2 id="chro-form-title" className="settings-modal-title company-details-modal-title">
                  {form.rowId ? 'Edit CHRO Notification' : 'Add CHRO Notification'}
                </h2>
              </div>

              <form className="settings-modal-form company-details-modal-form" onSubmit={submitForm}>
                <div className="settings-modal-scroll company-details-modal-scroll">
                  <div className="settings-modal-body company-details-modal-body">
                    <div className="settings-section-card company-details-section-card">
                      <div className="settings-fields-grid company-details-fields-grid">
                        <div className="settings-field company-details-field company-details-field--span2">
                          <label htmlFor="chro-name">
                            CHRO Name <span className="required">*</span>
                          </label>
                          <input
                            id="chro-name"
                            type="text"
                            value={form.name}
                            onChange={(e) => setForm((prev) => ({ ...prev, name: e.target.value }))}
                            placeholder="Enter CHRO name"
                            disabled={saving}
                          />
                        </div>
                        <div className="settings-field company-details-field company-details-field--span2">
                          <label htmlFor="chro-email">
                            Email <span className="required">*</span>
                          </label>
                          <input
                            id="chro-email"
                            type="email"
                            value={form.email}
                            onChange={(e) => setForm((prev) => ({ ...prev, email: e.target.value }))}
                            placeholder="Enter email"
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
