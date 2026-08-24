import React, { useCallback, useEffect, useState } from 'react';
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

export default function Settings() {
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
  }, [loadSettings]);

  const openAddForm = () => {
    setShowForm(true);
    setFormDueDate('');
    setMessage('');
  };

  const openEditForm = () => {
    setShowForm(true);
    setFormDueDate(savedDueDate || '');
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

  return (
    <div className={`settings-page company-details-page${showForm ? ' company-details-page-form-open' : ''}`}>
      <div className="settings-shell">
        {!showForm ? (
          <div className="settings-list-card company-details-shell-box">
            <div className="settings-page-heading settings-page-heading--row">
              <h1 className="settings-page-title">Email Notification</h1>
              <div className="settings-page-actions">
                <button type="button" className="company-details-btn company-details-btn-add" onClick={openAddForm}>
                  Add
                </button>
              </div>
            </div>
            <div className="settings-list-body">
              {showLoadingPlaceholder ? (
                <p className="settings-list-empty">Loading...</p>
              ) : (
                <div className="company-details-table-wrap">
                  <div className="settings-due-date-list">
                    {settingsRows.length === 0 ? (
                      <p className="settings-list-empty">No settings added yet. Click Add to open form.</p>
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
          </div>
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
        )}
      </div>
    </div>
  );
}
