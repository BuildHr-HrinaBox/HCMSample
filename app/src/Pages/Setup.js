import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import './Setup.css';
import './CompanyDetails.css';
import {
  checklistStateMatchesSiteState,
  normalizeStateCompareKey,
  siteInchargeEmail,
  siteNameFromSiteRecord,
  siteStateFromRecord,
} from '../utils/siteInchargeScope';

const FORMMASTER_API = '/server/formmaster_function/records';
const SITE_API = '/server/sitemanagement_function/sitemanagement';
const SETUP_API = '/server/setup_function/setup';
const EMAILS_STORAGE_KEY = 'hcm_setup_form_emails';
const EMAIL_VALUE_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

function readStoredEmails() {
  try {
    const raw = localStorage.getItem(EMAILS_STORAGE_KEY);
    const parsed = raw ? JSON.parse(raw) : {};
    return parsed && typeof parsed === 'object' ? parsed : {};
  } catch (_) {
    return {};
  }
}

function writeStoredEmails(map) {
  try {
    localStorage.setItem(EMAILS_STORAGE_KEY, JSON.stringify(map));
  } catch (_) {
    /* ignore quota / private mode */
  }
}

function emailStorageKey(state, site, formName) {
  return `${String(state || '').trim().toLowerCase()}|${String(site || '').trim().toLowerCase()}|${String(formName || '').trim().toLowerCase()}`;
}

function uniqueSorted(values) {
  return [...new Set(values.map((v) => String(v || '').trim()).filter(Boolean))].sort((a, b) =>
    a.localeCompare(b)
  );
}

function isValidEmail(value) {
  return EMAIL_VALUE_RE.test(String(value || '').trim());
}

function parseStoredEmails(value) {
  if (Array.isArray(value)) {
    return uniqueSorted(value.filter(isValidEmail));
  }
  if (typeof value !== 'string') return [];
  return uniqueSorted(
    value
      .split(/[,;]+/)
      .map((part) => part.trim())
      .filter(isValidEmail)
  );
}

function mapSetupRowsToEmails(rows) {
  const map = {};
  (Array.isArray(rows) ? rows : []).forEach((row) => {
    const formName = row?.formName || row?.FormName;
    if (!formName) return;
    const key = emailStorageKey(row.state || row.State, row.site || row.Site, formName);
    map[key] = parseStoredEmails(row.email ?? row.Email);
  });
  return map;
}

function getSavedEmailsForForm(emailsMap, formName, { state = '', site = '' } = {}) {
  const formKey = String(formName || '').trim().toLowerCase();
  const stateFilter = normalizeStateCompareKey(state);
  const siteFilter = String(site || '').trim().toLowerCase();
  const bucket = [];
  Object.entries(emailsMap || {}).forEach(([key, value]) => {
    const parts = key.split('|');
    if (parts.length < 3) return;
    const [keyState, keySite, keyForm] = parts;
    if (keyForm !== formKey) return;
    if (stateFilter && normalizeStateCompareKey(keyState) !== stateFilter) return;
    if (siteFilter && keySite !== siteFilter) return;
    parseStoredEmails(value).forEach((email) => bucket.push(email));
  });
  return uniqueSorted(bucket);
}

function EmailMultiSelect({ options, value, onChange }) {
  const [open, setOpen] = useState(false);
  const rootRef = useRef(null);
  const selected = Array.isArray(value) ? value : [];
  const selectedSet = useMemo(
    () => new Set(selected.map((email) => email.toLowerCase())),
    [selected]
  );

  useEffect(() => {
    if (!open) return undefined;
    const onPointerDown = (event) => {
      if (rootRef.current && !rootRef.current.contains(event.target)) {
        setOpen(false);
      }
    };
    document.addEventListener('mousedown', onPointerDown);
    return () => document.removeEventListener('mousedown', onPointerDown);
  }, [open]);

  const toggleEmail = (email) => {
    const exists = selectedSet.has(email.toLowerCase());
    const next = exists
      ? selected.filter((item) => item.toLowerCase() !== email.toLowerCase())
      : [...selected, email];
    onChange(uniqueSorted(next));
  };

  const summary = selected.length === 0
    ? 'Select email'
    : selected.length === 1
      ? selected[0]
      : `${selected.length} emails selected`;

  return (
    <div className={`setup-multi-select${open ? ' setup-multi-select--open' : ''}`} ref={rootRef}>
      <button
        type="button"
        className="setup-multi-select-trigger"
        onClick={() => setOpen((prev) => !prev)}
        aria-haspopup="listbox"
        aria-expanded={open}
        title={selected.join(', ')}
      >
        <span className={selected.length ? '' : 'setup-multi-select-placeholder'}>{summary}</span>
        <span className="setup-multi-select-caret" aria-hidden>▾</span>
      </button>
      {open ? (
        <div className="setup-multi-select-menu" role="listbox" aria-multiselectable="true">
          {options.length === 0 ? (
            <div className="setup-multi-select-empty">No emails for this state/site</div>
          ) : (
            options.map((email) => {
              const checked = selectedSet.has(email.toLowerCase());
              return (
                <label key={email} className="setup-multi-select-option">
                  <input
                    type="checkbox"
                    checked={checked}
                    onChange={() => toggleEmail(email)}
                  />
                  <span>{email}</span>
                </label>
              );
            })
          )}
        </div>
      ) : null}
    </div>
  );
}

export default function Setup({ userEmail = '' }) {
  const [sites, setSites] = useState([]);
  const [formRecords, setFormRecords] = useState([]);
  const [loading, setLoading] = useState(true);
  const [message, setMessage] = useState('');
  const [selectedState, setSelectedState] = useState('');
  const [selectedSite, setSelectedSite] = useState('');
  const [emails, setEmails] = useState(() => readStoredEmails());
  const [saving, setSaving] = useState(false);

  const loadData = useCallback(async () => {
    setLoading(true);
    setMessage('');
    try {
      let cachedSites = [];
      try {
        const cached = localStorage.getItem('siteManagementData');
        const parsed = cached ? JSON.parse(cached) : [];
        if (Array.isArray(parsed) && parsed.length) cachedSites = parsed;
      } catch (_) {}
      if (cachedSites.length) setSites(cachedSites);

      const [siteRes, formRes, setupRes] = await Promise.all([
        fetch(SITE_API, { credentials: 'include', cache: 'no-store' }),
        fetch(`${FORMMASTER_API}?action=getAll&_ts=${Date.now()}`, {
          credentials: 'include',
          cache: 'no-store',
        }),
        fetch(SETUP_API, { credentials: 'include', cache: 'no-store' }),
      ]);
      const siteData = await siteRes.json().catch(() => null);
      const formData = await formRes.json().catch(() => null);
      const setupData = await setupRes.json().catch(() => null);

      if (siteData?.status === 'success' && Array.isArray(siteData?.data?.siteDetails)) {
        setSites(siteData.data.siteDetails);
      } else if (!cachedSites.length) {
        setSites([]);
      }

      if (formData?.status === 'success' && Array.isArray(formData.data)) {
        setFormRecords(formData.data);
      } else {
        setFormRecords([]);
      }

      if (setupData?.status === 'success' && Array.isArray(setupData.data)) {
        const fromBackend = mapSetupRowsToEmails(setupData.data);
        setEmails((prev) => {
          const next = { ...prev, ...fromBackend };
          writeStoredEmails(next);
          return next;
        });
      }
    } catch (err) {
      setMessage(err?.message || 'Failed to load Setup data.');
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    loadData();
  }, [loadData]);

  const stateOptions = useMemo(
    () => uniqueSorted(sites.map(siteStateFromRecord)),
    [sites]
  );

  const sitesForState = useMemo(() => {
    if (!selectedState) return sites;
    return sites.filter((site) => checklistStateMatchesSiteState(siteStateFromRecord(site), selectedState));
  }, [sites, selectedState]);

  const siteOptions = useMemo(
    () => uniqueSorted(sitesForState.map(siteNameFromSiteRecord)),
    [sitesForState]
  );

  useEffect(() => {
    if (selectedState && !stateOptions.includes(selectedState)) {
      setSelectedState('');
      setSelectedSite('');
    }
  }, [selectedState, stateOptions]);

  useEffect(() => {
    if (selectedSite && !siteOptions.includes(selectedSite)) {
      setSelectedSite('');
    }
  }, [selectedSite, siteOptions]);

  const adminEmail = String(userEmail || '').trim();

  const allEmailOptions = useMemo(() => {
    const bucket = [];
    const admin = isValidEmail(adminEmail) ? adminEmail : '';

    const sitesForEmails = selectedSite
      ? sitesForState.filter((site) => siteNameFromSiteRecord(site) === selectedSite)
      : selectedState
        ? sitesForState
        : sites;

    sitesForEmails.forEach((site) => {
      const siteEmail = siteInchargeEmail(site);
      if (isValidEmail(siteEmail)) bucket.push(siteEmail);
    });

    Object.values(emails).forEach((value) => {
      parseStoredEmails(value).forEach((email) => bucket.push(email));
    });

    const merged = uniqueSorted(bucket);
    const withoutAdmin = merged.filter((email) => email.toLowerCase() !== admin.toLowerCase());
    return admin ? [admin, ...withoutAdmin] : withoutAdmin;
  }, [adminEmail, selectedState, selectedSite, sitesForState, sites, emails]);

  const filteredForms = useMemo(() => {
    const rows = selectedState
      ? formRecords.filter((row) => checklistStateMatchesSiteState(row.state, selectedState))
      : formRecords;
    const seen = new Set();
    const unique = [];
    rows.forEach((row) => {
      const formName = String(row.formName || row.FormName || row.act || '').trim();
      if (!formName) return;
      const key = formName.toLowerCase();
      if (seen.has(key)) return;
      seen.add(key);
      unique.push({
        id: row.id || key,
        formName,
      });
    });
    unique.sort((a, b) => a.formName.localeCompare(b.formName));
    return unique;
  }, [formRecords, selectedState]);

  const getEmailValue = useCallback(
    (formName) => {
      const exactKey = emailStorageKey(selectedState, selectedSite, formName);
      if (Object.prototype.hasOwnProperty.call(emails, exactKey)) {
        const exactSaved = parseStoredEmails(emails[exactKey]);
        if (exactSaved.length) return exactSaved;
      }

      const scopedSaved = getSavedEmailsForForm(emails, formName, {
        state: selectedState,
        site: selectedSite,
      });
      if (scopedSaved.length) return scopedSaved;

      return allEmailOptions;
    },
    [emails, selectedState, selectedSite, allEmailOptions]
  );

  const handleEmailChange = (formName, nextEmails) => {
    const key = emailStorageKey(selectedState, selectedSite, formName);
    setEmails((prev) => {
      const next = { ...prev, [key]: nextEmails };
      writeStoredEmails(next);
      return next;
    });
  };

  const handleSave = async () => {
    if (!selectedState) {
      setMessage('Select a state before saving.');
      return;
    }
    if (!filteredForms.length) {
      setMessage('No forms to save for the selected state.');
      return;
    }

    const rows = filteredForms.map((row) => ({
      formName: row.formName,
      emails: getEmailValue(row.formName),
    }));
    const next = { ...emails };
    rows.forEach((row) => {
      next[emailStorageKey(selectedState, selectedSite, row.formName)] = row.emails;
    });
    setEmails(next);
    writeStoredEmails(next);

    setSaving(true);
    setMessage('');
    try {
      const res = await fetch(SETUP_API, {
        method: 'PUT',
        credentials: 'include',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          state: selectedState,
          site: selectedSite,
          rows,
        }),
      });
      const data = await res.json().catch(() => null);
      if (!res.ok || data?.status !== 'success') {
        throw new Error(data?.message || 'Failed to save Setup data.');
      }
      setMessage(data.message || 'Saved form emails to Setup.');
    } catch (err) {
      setMessage(err?.message || 'Failed to save Setup data.');
    } finally {
      setSaving(false);
    }
  };

  return (
    <div className="setup-page company-details-page">
      <div className="setup-shell">
        <div className="setup-card company-details-shell-box">
          <div className="setup-heading">
            <h1 className="setup-title">Setup</h1>
          </div>

          <div className="setup-filter-row">
            <label className="setup-filter">
              <span>State</span>
              <select
                value={selectedState}
                onChange={(e) => {
                  setSelectedState(e.target.value);
                  setSelectedSite('');
                  setMessage('');
                }}
              >
                <option value="">All states</option>
                {stateOptions.map((state) => (
                  <option key={state} value={state}>
                    {state}
                  </option>
                ))}
              </select>
            </label>
            <label className="setup-filter">
              <span>Site</span>
              <select
                value={selectedSite}
                onChange={(e) => {
                  setSelectedSite(e.target.value);
                  setMessage('');
                }}
              >
                <option value="">All sites</option>
                {siteOptions.map((site) => (
                  <option key={site} value={site}>
                    {site}
                  </option>
                ))}
              </select>
            </label>
            <button type="button" className="setup-save-btn" onClick={handleSave} disabled={saving}>
              {saving ? 'Saving...' : 'Save'}
            </button>
          </div>

          {message ? (
            <p
              className={`setup-message${/fail|required|select a state/i.test(message) ? ' setup-message-error' : ''}`}
              role="status"
            >
              {message}
            </p>
          ) : null}

          <div className="setup-table-wrap company-details-table-wrap">
            {loading && !formRecords.length ? (
              <p className="setup-empty">Loading…</p>
            ) : (
              <table className="company-details-data-table setup-data-table">
                <thead>
                  <tr>
                    <th>Forms</th>
                    <th>Email</th>
                  </tr>
                </thead>
                <tbody>
                  {filteredForms.length === 0 ? (
                    <tr>
                      <td colSpan={2} className="company-details-table-empty-cell">
                        {selectedState
                          ? 'No forms found for the selected state.'
                          : 'No forms found. Import Form Master records first.'}
                      </td>
                    </tr>
                  ) : (
                    filteredForms.map((row) => (
                      <tr key={row.id} className="company-details-data-row">
                        <td>{row.formName}</td>
                        <td>
                          <EmailMultiSelect
                            options={allEmailOptions}
                            value={getEmailValue(row.formName)}
                            onChange={(nextEmails) => handleEmailChange(row.formName, nextEmails)}
                          />
                        </td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            )}
          </div>
        </div>
      </div>
    </div>
  );
}
