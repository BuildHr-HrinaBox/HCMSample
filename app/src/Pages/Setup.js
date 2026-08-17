import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { Download, RotateCcw, Search } from 'lucide-react';
import './Setup.css';
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

function mapSetupRowsToMeta(rows) {
  const map = {};
  (Array.isArray(rows) ? rows : []).forEach((row) => {
    const formName = row?.formName || row?.FormName;
    if (!formName) return;
    const key = emailStorageKey(row.state || row.State, row.site || row.Site, formName);
    map[key] = {
      act: String(row.act ?? row.Act ?? '').trim(),
      description: String(row.description ?? row.Description ?? '').trim(),
      role: String(row.role ?? row.Role ?? '').trim(),
    };
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

function exportAccessRowsCsv(rows) {
  const headers = ['Role Name', 'Act', 'Description', 'Form Name', 'Email ID', 'Status'];
  const escape = (value) => `"${String(value ?? '').replace(/"/g, '""')}"`;
  const lines = [headers.map(escape).join(',')];
  rows.forEach((row) => {
    lines.push(
      [row.roleName, row.act, row.description, row.formName, row.email, row.status]
        .map(escape)
        .join(',')
    );
  });
  const blob = new Blob([lines.join('\n')], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = 'role-based-access.csv';
  link.click();
  URL.revokeObjectURL(url);
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

  const summary =
    selected.length === 0
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
        <span className="setup-multi-select-caret" aria-hidden>
          ▾
        </span>
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
                  <input type="checkbox" checked={checked} onChange={() => toggleEmail(email)} />
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
  const [search, setSearch] = useState('');
  const [statusFilter, setStatusFilter] = useState('All');
  const [emails, setEmails] = useState(() => readStoredEmails());
  const [setupMeta, setSetupMeta] = useState({});
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
        const metaFromBackend = mapSetupRowsToMeta(setupData.data);
        setEmails((prev) => {
          const next = { ...prev, ...fromBackend };
          writeStoredEmails(next);
          return next;
        });
        setSetupMeta((prev) => ({ ...prev, ...metaFromBackend }));
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
      const formName = String(row.formName || row.FormName || '').trim();
      if (!formName) return;
      const key = formName.toLowerCase();
      const act = String(row.act || row.Act || '').trim();
      const description = String(row.description || row.Description || '').trim();
      const metaKey = emailStorageKey(selectedState, selectedSite, formName);
      const savedMeta = setupMeta[metaKey] || {};

      if (seen.has(key)) {
        const existing = unique.find((item) => item.formName.toLowerCase() === key);
        if (existing) {
          if (!existing.act && act) existing.act = act;
          if (!existing.description && description) existing.description = description;
          if (!existing.act && savedMeta.act) existing.act = savedMeta.act;
          if (!existing.description && savedMeta.description) existing.description = savedMeta.description;
        }
        return;
      }
      seen.add(key);
      unique.push({
        id: row.id || key,
        formName,
        act: act || savedMeta.act || '',
        description: description || savedMeta.description || '',
        role: String(savedMeta.role || '').trim(),
      });
    });
    unique.sort((a, b) => a.formName.localeCompare(b.formName));
    return unique;
  }, [formRecords, selectedState, selectedSite, setupMeta]);

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

  const accessRows = useMemo(() => {
    const rows = [];
    filteredForms.forEach((form) => {
      const assignedEmails = getEmailValue(form.formName);
      const emailList = assignedEmails.length ? assignedEmails : [''];
      emailList.forEach((email, index) => {
        rows.push({
          id: `${form.id}-${index}-${email || 'empty'}`,
          formName: form.formName,
          roleName: form.role,
          act: form.act,
          description: form.description,
          email,
          status: email ? 'Active' : 'Inactive',
        });
      });
    });
    return rows;
  }, [filteredForms, getEmailValue]);

  const displayedRows = useMemo(() => {
    let rows = accessRows;
    if (statusFilter !== 'All') {
      rows = rows.filter((row) => row.status === statusFilter);
    }
    const query = search.trim().toLowerCase();
    if (query) {
      rows = rows.filter((row) =>
        [row.roleName, row.act, row.description, row.formName, row.email, row.status].some((value) =>
          String(value || '').toLowerCase().includes(query)
        )
      );
    }
    return rows;
  }, [accessRows, statusFilter, search]);

  const handleEmailChange = (formName, nextEmails) => {
    const key = emailStorageKey(selectedState, selectedSite, formName);
    setEmails((prev) => {
      const next = { ...prev, [key]: nextEmails };
      writeStoredEmails(next);
      return next;
    });
  };

  const handleRoleChange = (formName, role) => {
    const key = emailStorageKey(selectedState, selectedSite, formName);
    setSetupMeta((prev) => ({
      ...prev,
      [key]: {
        ...(prev[key] || {}),
        role: String(role || '').trim(),
      },
    }));
  };

  const handleReset = () => {
    setSearch('');
    setStatusFilter('All');
    setSelectedState('');
    setSelectedSite('');
    setMessage('');
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
      act: row.act,
      description: row.description,
      role: row.role,
      emails: getEmailValue(row.formName),
    }));
    const next = { ...emails };
    const nextMeta = { ...setupMeta };
    rows.forEach((row) => {
      const key = emailStorageKey(selectedState, selectedSite, row.formName);
      next[key] = row.emails;
      nextMeta[key] = {
        act: row.act,
        description: row.description,
        role: row.role,
      };
    });
    setEmails(next);
    setSetupMeta(nextMeta);
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
      setMessage(data.message || 'Saved role based access settings.');
    } catch (err) {
      setMessage(err?.message || 'Failed to save Setup data.');
    } finally {
      setSaving(false);
    }
  };

  const totalRows = displayedRows.length;
  const footerStart = totalRows === 0 ? 0 : 1;
  const footerEnd = totalRows;

  return (
    <div className="setup-rba-page">
      <div className="setup-rba-card">
        <div className="setup-rba-header">
          <div className="setup-rba-header-text">
            <h1 className="setup-rba-title">Role Based Access</h1>
            <p className="setup-rba-subtitle">Manage roles, forms and user access for HCM.</p>
          </div>
          <div className="setup-rba-header-actions">
            <button
              type="button"
              className="setup-rba-btn setup-rba-btn--save"
              onClick={handleSave}
              disabled={saving}
            >
              {saving ? 'Saving…' : 'Save'}
            </button>
            <button
              type="button"
              className="setup-rba-btn setup-rba-btn--export"
              onClick={() => exportAccessRowsCsv(displayedRows)}
              disabled={!displayedRows.length}
            >
              <Download size={16} strokeWidth={2} aria-hidden />
              Export
            </button>
          </div>
        </div>

        <div className="setup-rba-toolbar">
          <div className="setup-rba-field setup-rba-search-wrap">
            <span className="setup-rba-field-label">Search</span>
            <Search className="setup-rba-search-icon" size={16} strokeWidth={2} aria-hidden />
            <input
              type="search"
              className="setup-rba-search-input"
              value={search}
              onChange={(e) => setSearch(e.target.value)}
              placeholder="Search by role name, form name, act..."
              aria-label="Search role based access rows"
            />
          </div>

          <label className="setup-rba-field">
            <span className="setup-rba-field-label">Status</span>
            <select
              className="setup-rba-select"
              value={statusFilter}
              onChange={(e) => setStatusFilter(e.target.value)}
            >
              <option value="All">All</option>
              <option value="Active">Active</option>
              <option value="Inactive">Inactive</option>
            </select>
          </label>

          <label className="setup-rba-field">
            <span className="setup-rba-field-label">State</span>
            <select
              className="setup-rba-select"
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

          <label className="setup-rba-field">
            <span className="setup-rba-field-label">Site</span>
            <select
              className="setup-rba-select"
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

          <button type="button" className="setup-rba-btn setup-rba-btn--reset" onClick={handleReset}>
            <RotateCcw size={16} strokeWidth={2} aria-hidden />
            Reset
          </button>
        </div>

        {message ? (
          <p
            className={`setup-rba-message${/fail|required|select a state/i.test(message) ? ' setup-rba-message--error' : ''}`}
            role="status"
          >
            {message}
          </p>
        ) : null}

        <div className="setup-rba-table-wrap">
          {loading && !formRecords.length ? (
            <p className="setup-rba-empty">Loading…</p>
          ) : (
            <table className="setup-rba-table">
              <thead>
                <tr>
                  <th>Role Name</th>
                  <th>Act</th>
                  <th>Description</th>
                  <th>Form Name</th>
                  <th>Email ID</th>
                  <th>Status</th>
                </tr>
              </thead>
              <tbody>
                {displayedRows.length === 0 ? (
                  <tr>
                    <td colSpan={6} className="setup-rba-empty">
                      {selectedState
                        ? 'No access rows found for the selected filters.'
                        : 'No forms found. Import Form Master records first.'}
                    </td>
                  </tr>
                ) : (
                  displayedRows.map((row) => (
                    <tr key={row.id}>
                      <td>
                        <input
                          type="text"
                          className="setup-rba-cell-input"
                          value={row.roleName}
                          onChange={(e) => handleRoleChange(row.formName, e.target.value)}
                          placeholder="Enter role name"
                          aria-label={`Role name for ${row.formName}`}
                        />
                      </td>
                      <td>{row.act || '—'}</td>
                      <td>{row.description || '—'}</td>
                      <td>{row.formName}</td>
                      <td>
                        <EmailMultiSelect
                          options={allEmailOptions}
                          value={getEmailValue(row.formName)}
                          onChange={(nextEmails) => handleEmailChange(row.formName, nextEmails)}
                        />
                      </td>
                      <td>
                        <span
                          className={`setup-rba-status setup-rba-status--${row.status === 'Active' ? 'active' : 'inactive'}`}
                        >
                          {row.status}
                        </span>
                      </td>
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          )}
        </div>

        <div className="setup-rba-footer">
          Showing {footerStart} to {footerEnd} of {totalRows} entries
        </div>
      </div>
    </div>
  );
}
