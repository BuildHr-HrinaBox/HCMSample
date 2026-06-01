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

export default function Settings() {
  const cachedOnMount = readSettingsCache();
  const [showForm, setShowForm] = useState(false);
  const [savedCompanyName, setSavedCompanyName] = useState(() =>
    cachedOnMount ? cachedOnMount.companyName : ''
  );
  const [savedLogoName, setSavedLogoName] = useState(() =>
    cachedOnMount ? cachedOnMount.logoName : ''
  );
  const [formCompanyName, setFormCompanyName] = useState('');
  const [formCurrentLogoName, setFormCurrentLogoName] = useState('');
  const [selectedLogo, setSelectedLogo] = useState(null);
  const [saving, setSaving] = useState(false);
  const [loading, setLoading] = useState(() => !cachedOnMount);
  const [message, setMessage] = useState('');
  const [selectedRowIds, setSelectedRowIds] = useState([]);

  const selectedLogoName = useMemo(() => {
    if (!selectedLogo) return '';
    return selectedLogo.name || '';
  }, [selectedLogo]);

  const applySettings = useCallback((data) => {
    if (!data) return;
    setSavedCompanyName(String(data.companyName || ''));
    setSavedLogoName(String(data.logoName || ''));
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
    setFormCompanyName('');
    setFormCurrentLogoName('');
    setSelectedLogo(null);
    setMessage('');
  };

  const openEditForm = () => {
    setShowForm(true);
    setFormCompanyName(savedCompanyName || '');
    setFormCurrentLogoName(savedLogoName || '');
    setSelectedLogo(null);
    setMessage('');
  };

  const closeForm = () => {
    setShowForm(false);
    setMessage('');
  };

  const submitForm = async (event) => {
    event.preventDefault();
    const trimmedName = formCompanyName.trim();
    if (!trimmedName) {
      setMessage('Company name is required.');
      return;
    }

    setSaving(true);
    setMessage('');
    try {
      const formData = new FormData();
      formData.append('companyName', trimmedName);
      if (selectedLogo) formData.append('logo', selectedLogo);

      const res = await fetch(`${API_BASE}/settings`, {
        method: 'PUT',
        body: formData,
      });
      const data = await res.json();
      if (!res.ok || data?.status !== 'success') {
        throw new Error(data?.message || 'Failed to save settings');
      }

      const nextCompanyName = String(data?.data?.companyName || trimmedName);
      const nextLogoName = String(data?.data?.logoName || '');
      setSavedCompanyName(nextCompanyName);
      setSavedLogoName(nextLogoName);
      writeSettingsCache(data?.data);
      setFormCurrentLogoName(nextLogoName);
      setSelectedLogo(null);
      setMessage('');
      setShowForm(false);
    } catch (err) {
      setMessage(err?.message || 'Unable to save settings.');
    } finally {
      setSaving(false);
    }
  };

  const hasSavedSettings = savedCompanyName.trim() || savedLogoName.trim();
  const settingsRows = hasSavedSettings
    ? [
        {
          id: '1',
          companyName: savedCompanyName || '-',
          logo: savedLogoName || '-',
        },
      ]
    : [];
  const allSelected = settingsRows.length > 0 && selectedRowIds.length === settingsRows.length;
  const showLoadingPlaceholder = loading && settingsRows.length === 0;

  const toggleSelectAll = (checked) => {
    if (checked) {
      setSelectedRowIds(settingsRows.map((r) => r.id));
    } else {
      setSelectedRowIds([]);
    }
  };

  const toggleSelectRow = (rowId, checked) => {
    setSelectedRowIds((prev) => {
      if (checked) return [...new Set([...prev, rowId])];
      return prev.filter((id) => id !== rowId);
    });
  };

  const handleDeleteSelected = async () => {
    if (selectedRowIds.length === 0) {
      setMessage('Select at least one row to delete.');
      return;
    }
    try {
      const res = await fetch(`${API_BASE}/settings`, { method: 'DELETE' });
      const data = await res.json();
      if (!res.ok || data?.status !== 'success') {
        throw new Error(data?.message || 'Failed to delete settings');
      }
      setSavedCompanyName('');
      setSavedLogoName('');
      clearSettingsCache();
      setFormCompanyName('');
      setFormCurrentLogoName('');
      setSelectedLogo(null);
      setSelectedRowIds([]);
      setMessage('Deleted successfully.');
    } catch (err) {
      setMessage(err?.message || 'Unable to delete settings.');
    }
  };

  const handleDeleteRow = async () => {
    try {
      const res = await fetch(`${API_BASE}/settings`, { method: 'DELETE' });
      const data = await res.json();
      if (!res.ok || data?.status !== 'success') {
        throw new Error(data?.message || 'Failed to delete settings');
      }
      setSavedCompanyName('');
      setSavedLogoName('');
      clearSettingsCache();
      setFormCompanyName('');
      setFormCurrentLogoName('');
      setSelectedLogo(null);
      setSelectedRowIds([]);
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
              <h1 className="settings-page-title">Settings</h1>
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
                  <table className="company-details-data-table settings-data-table">
                    <thead>
                      <tr>
                        <th style={{ width: 40 }}>
                          <input
                            type="checkbox"
                            className="company-details-table-checkbox"
                            checked={allSelected}
                            onChange={(e) => toggleSelectAll(e.target.checked)}
                            aria-label="Select all settings rows"
                          />
                        </th>
                        <th>#</th>
                        <th>Company Name</th>
                        <th>Logo</th>
                        <th>Actions</th>
                      </tr>
                    </thead>
                    <tbody>
                      {settingsRows.length === 0 ? (
                        <tr>
                          <td colSpan={5} className="company-details-table-empty-cell">
                            No settings added yet. Click Add to open form.
                          </td>
                        </tr>
                      ) : (
                        settingsRows.map((row, index) => (
                          <tr key={row.id} className="company-details-data-row">
                            <td>
                              <input
                                type="checkbox"
                                className="company-details-table-checkbox"
                                checked={selectedRowIds.includes(row.id)}
                                onChange={(e) => toggleSelectRow(row.id, e.target.checked)}
                                aria-label={`Select settings row ${index + 1}`}
                              />
                            </td>
                            <td>{index + 1}</td>
                            <td>{row.companyName}</td>
                            <td>{row.logo}</td>
                            <td>
                              <div className="company-details-table-actions settings-table-actions">
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
                            </td>
                          </tr>
                        ))
                      )}
                    </tbody>
                  </table>
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
                          <label htmlFor="settings-company-name">Company Name</label>
                          <input
                            id="settings-company-name"
                            type="text"
                            value={formCompanyName}
                            onChange={(e) => setFormCompanyName(e.target.value)}
                            placeholder="Enter company name"
                            disabled={saving}
                            required
                          />
                        </div>

                        <div className="settings-field company-details-field company-details-field--span2">
                          <label htmlFor="settings-logo">Logo</label>
                          <div className="company-details-doc-upload-input-wrap">
                            <div className="company-details-doc-upload-icon-inside" aria-hidden>
                              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                                <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
                                <polyline points="7 10 12 5 17 10" />
                                <line x1="12" y1="5" x2="12" y2="17" />
                              </svg>
                              <span>Import</span>
                            </div>
                            <input
                              id="settings-logo"
                              type="file"
                              className="company-details-doc-file-input"
                              accept=".png,.jpg,.jpeg,.webp,.svg"
                              onChange={(e) => setSelectedLogo(e.target.files?.[0] || null)}
                              disabled={saving}
                            />
                          </div>
                          {selectedLogoName ? (
                            <p className="company-details-doc-pending">Selected: {selectedLogoName}</p>
                          ) : formCurrentLogoName ? (
                            <p className="company-details-doc-pending">Current: {formCurrentLogoName}</p>
                          ) : null}
                          <p className="company-details-doc-hint-line">Max 5MB · JPG, JPEG, PNG, WEBP, SVG</p>
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
