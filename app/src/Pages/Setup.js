import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { ChevronRight, Download, Info, Plus, RotateCcw, Search, X } from 'lucide-react';
import './Setup.css';
import {
  checklistStateMatchesSiteState,
  getActCategoryFromActSector,
  industryLabelToActCategory,
  normalizeStateCompareKey,
  siteInchargeEmail,
  siteIndustry,
  siteNameFromSiteRecord,
  siteStateFromRecord,
} from '../utils/siteInchargeScope';

const INDUSTRY_OPTIONS = ['CLRA', 'Shops and Establishment', 'Factories Act'];

const FORMMASTER_API = '/server/formmaster_function/records';
const SITE_API = '/server/sitemanagement_function/sitemanagement';
const SETUP_API = '/server/setup_function/setup';
const EMAILS_STORAGE_KEY = 'hcm_setup_form_emails';
const META_STORAGE_KEY = 'hcm_setup_form_meta';
const FILTERS_STORAGE_KEY = 'hcm_setup_filters';
const EMAIL_VALUE_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
const ROWS_PER_PAGE = 10;

/** @returns {(number | 'ellipsis')[]} */
function buildPaginationItems(currentPage, totalPages) {
  if (totalPages <= 0) return [];
  if (totalPages <= 9) {
    return Array.from({ length: totalPages }, (_, i) => i + 1);
  }
  const set = new Set([1, totalPages, currentPage]);
  for (let i = currentPage - 2; i <= currentPage + 2; i += 1) {
    if (i >= 1 && i <= totalPages) set.add(i);
  }
  const sorted = [...set].sort((a, b) => a - b);
  const out = [];
  for (let i = 0; i < sorted.length; i += 1) {
    if (i > 0 && sorted[i] - sorted[i - 1] > 1) {
      out.push('ellipsis');
    }
    out.push(sorted[i]);
  }
  return out;
}

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

function readStoredMeta() {
  try {
    const raw = localStorage.getItem(META_STORAGE_KEY);
    const parsed = raw ? JSON.parse(raw) : {};
    return parsed && typeof parsed === 'object' && !Array.isArray(parsed) ? parsed : {};
  } catch (_) {
    return {};
  }
}

function writeStoredMeta(map) {
  try {
    localStorage.setItem(META_STORAGE_KEY, JSON.stringify(map));
  } catch (_) {
    /* ignore quota / private mode */
  }
}

function persistSetupMaps(emailsMap, metaMap) {
  writeStoredEmails(emailsMap);
  writeStoredMeta(metaMap);
}

function readStoredFilters() {
  try {
    const raw = sessionStorage.getItem(FILTERS_STORAGE_KEY);
    const parsed = raw ? JSON.parse(raw) : null;
    if (!parsed || typeof parsed !== 'object') return null;
    return parsed;
  } catch (_) {
    return null;
  }
}

function writeStoredFilters(filters) {
  try {
    sessionStorage.setItem(FILTERS_STORAGE_KEY, JSON.stringify(filters));
  } catch (_) {
    /* ignore quota / private mode */
  }
}

function emailStorageKey(state, site, formName, industryType = '', act = '', description = '') {
  return [
    String(state || '').trim().toLowerCase(),
    String(site || '').trim().toLowerCase(),
    String(formName || '').trim().toLowerCase(),
    String(industryType || '').trim().toLowerCase(),
    String(act || '').trim().toLowerCase(),
    String(description || '').trim().toLowerCase(),
  ].join('|');
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
    const formName = row?.formName || row?.FormName || row?.Name;
    if (!formName) return;
    const industryType = row.industryType ?? row.IndustryType ?? '';
    const act = row.act ?? row.Act ?? '';
    const description = row.description ?? row.Description ?? '';
    const key = emailStorageKey(
      row.state || row.State,
      row.site || row.Site,
      formName,
      industryType,
      act,
      description
    );
    map[key] = parseStoredEmails(row.email ?? row.Email ?? row.emails);
  });
  return map;
}

function mapSetupRowsToMeta(rows) {
  const map = {};
  (Array.isArray(rows) ? rows : []).forEach((row) => {
    const formName = String(row?.formName || row?.FormName || row?.Name || '').trim();
    if (!formName) return;
    const industryType = String(row.industryType ?? row.IndustryType ?? '').trim();
    const act = String(row.act ?? row.Act ?? '').trim();
    const description = String(row.description ?? row.Description ?? '').trim();
    const key = emailStorageKey(
      row.state || row.State,
      row.site || row.Site,
      formName,
      industryType,
      act,
      description
    );
    map[key] = {
      formName,
      act,
      description,
      role: String(row.role ?? row.Role ?? row.roleName ?? row.RoleName ?? '').trim(),
      state: String(row.state ?? row.State ?? '').trim(),
      site: String(row.site ?? row.Site ?? '').trim(),
      industryType,
      emails: parseStoredEmails(row.email ?? row.Email ?? row.emails),
    };
  });
  return map;
}

function emailsForMetaKey(emailsMap, key, meta) {
  if (Object.prototype.hasOwnProperty.call(emailsMap || {}, key)) {
    const exact = parseStoredEmails(emailsMap[key]);
    if (exact.length) return exact;
  }
  const fromMeta = parseStoredEmails(meta?.emails);
  if (fromMeta.length) return fromMeta;

  const formKey = String(meta?.formName || '').trim().toLowerCase();
  const stateKey = normalizeStateCompareKey(meta?.state);
  const bucket = [];
  Object.entries(emailsMap || {}).forEach(([storedKey, value]) => {
    const parts = String(storedKey).split('|');
    if (parts.length < 3) return;
    const [keyState, keySite, keyForm] = parts;
    if (formKey && keyForm !== formKey) return;
    if (stateKey && normalizeStateCompareKey(keyState) !== stateKey) return;
    if (siteKey && !sitesMatchFilter(keySite, meta?.site)) return;
    parseStoredEmails(value).forEach((email) => bucket.push(email));
  });
  return uniqueSorted(bucket);
}

function mergeEmailMaps(base, incoming) {
  const next = { ...(base || {}) };
  Object.entries(incoming || {}).forEach(([key, value]) => {
    const incomingEmails = parseStoredEmails(value);
    if (incomingEmails.length || !Object.prototype.hasOwnProperty.call(next, key)) {
      next[key] = incomingEmails;
    }
  });
  return next;
}

function mergeMetaMaps(base, incoming) {
  const next = { ...(base || {}) };
  Object.entries(incoming || {}).forEach(([key, value]) => {
    const prev = next[key] && typeof next[key] === 'object' ? next[key] : {};
    const incomingEmails = parseStoredEmails(value?.emails);
    const prevEmails = parseStoredEmails(prev.emails);
    next[key] = {
      ...prev,
      ...value,
      formName: String(value?.formName || prev.formName || '').trim(),
      role: String(value?.role || prev.role || '').trim(),
      industryType: String(value?.industryType || prev.industryType || '').trim(),
      act: String(value?.act || prev.act || '').trim(),
      description: String(value?.description || prev.description || '').trim(),
      state: String(value?.state || prev.state || '').trim(),
      site: String(value?.site || prev.site || '').trim(),
      emails: incomingEmails.length ? incomingEmails : prevEmails,
    };
  });
  return next;
}

function buildSavedAccessRows(setupMeta, emailsMap, formRecords, { state = '', site = '' } = {}) {
  const siteFilter = String(site || '').trim();
  const formNameLookup = new Map();
  (Array.isArray(formRecords) ? formRecords : []).forEach((row) => {
    const fn = String(row.formName || row.FormName || '').trim();
    if (fn) formNameLookup.set(fn.toLowerCase(), fn);
  });

  const rows = [];
  const seen = new Set();

  Object.entries(setupMeta || {}).forEach(([key, meta]) => {
    const parts = key.split('|');
    const rowState = String(meta?.state || parts[0] || '').trim();
    const rowSite = String(meta?.site || parts[1] || '').trim();
    const keyFormLower = String(meta?.formName || parts[2] || '')
      .trim()
      .toLowerCase();

    if (state) {
      const stateOk =
        checklistStateMatchesSiteState(rowState, state) ||
        normalizeStateCompareKey(rowState) === normalizeStateCompareKey(state);
      if (!stateOk) return;
    }
    if (siteFilter && !sitesMatchFilter(rowSite, siteFilter)) return;
    if (seen.has(key)) return;
    seen.add(key);

    const formName =
      String(meta?.formName || '').trim() ||
      formNameLookup.get(keyFormLower) ||
      keyFormLower;
    if (!formName) return;

    const assignedEmails = emailsForMetaKey(emailsMap, key, {
      ...meta,
      formName,
      state: rowState,
      site: rowSite,
    });
    const roleName = String(meta?.role || '').trim();
    if (!roleName && assignedEmails.length === 0) return;

    rows.push({
      id: key,
      formName,
      roleName: roleName || '—',
      act: String(meta?.act || '').trim(),
      description: String(meta?.description || '').trim(),
      industryType: String(meta?.industryType || '').trim(),
      rowState,
      rowSite,
      emails: assignedEmails,
      emailDisplay: assignedEmails.join(', '),
      status: assignedEmails.length > 0 ? 'Active' : 'Inactive',
    });
  });

  rows.sort((a, b) => {
    const score = (row) => (row.industryType ? 2 : 0) + (row.emails.length ? 1 : 0);
    return score(b) - score(a) || a.formName.localeCompare(b.formName);
  });
  const collapsed = [];
  const seenDisplay = new Set();
  rows.forEach((row) => {
    const displayKey = [
      row.rowState,
      row.rowSite,
      row.formName,
      row.act,
      row.description,
      row.roleName,
    ]
      .map((value) => String(value || '').trim().toLowerCase())
      .join('|');
    if (seenDisplay.has(displayKey)) return;
    seenDisplay.add(displayKey);
    collapsed.push(row);
  });
  collapsed.sort((a, b) => a.formName.localeCompare(b.formName));
  return collapsed;
}

function sitesMatchFilter(keySite, selectedSite) {
  const want = String(selectedSite || '').trim().toLowerCase();
  if (!want) return true;
  const have = String(keySite || '').trim().toLowerCase();
  if (!have) return false;
  return have === want || have.includes(want) || want.includes(have);
}

function hasStoredEntry(map, formName, { state = '', site = '' } = {}) {
  const formKey = String(formName || '').trim().toLowerCase();
  const stateFilter = normalizeStateCompareKey(state);
  const siteFilter = String(site || '').trim();
  const exactKey = emailStorageKey(state, site, formName);
  if (Object.prototype.hasOwnProperty.call(map || {}, exactKey)) return true;
  return Object.keys(map || {}).some((key) => {
    const parts = key.split('|');
    if (parts.length < 3) return false;
    const [keyState, keySite, keyForm] = parts;
    if (keyForm !== formKey) return false;
    if (stateFilter && normalizeStateCompareKey(keyState) !== stateFilter) return false;
    if (siteFilter && !sitesMatchFilter(keySite, siteFilter)) return false;
    return true;
  });
}

function getSavedEmailsForForm(emailsMap, formName, { state = '', site = '' } = {}) {
  const formKey = String(formName || '').trim().toLowerCase();
  const stateFilter = normalizeStateCompareKey(state);
  const siteFilter = String(site || '').trim();
  const exactKey = emailStorageKey(state, site, formName);
  if (Object.prototype.hasOwnProperty.call(emailsMap || {}, exactKey)) {
    return parseStoredEmails(emailsMap[exactKey]);
  }
  const bucket = [];
  Object.entries(emailsMap || {}).forEach(([key, value]) => {
    const parts = key.split('|');
    if (parts.length < 3) return;
    const [keyState, keySite, keyForm] = parts;
    if (keyForm !== formKey) return;
    if (stateFilter && normalizeStateCompareKey(keyState) !== stateFilter) return;
    if (siteFilter && !sitesMatchFilter(keySite, siteFilter)) return;
    parseStoredEmails(value).forEach((email) => bucket.push(email));
  });
  return uniqueSorted(bucket);
}

function findSetupMetaForForm(metaMap, formName, { state = '', site = '' } = {}) {
  const exactKey = emailStorageKey(state, site, formName);
  if (metaMap?.[exactKey]) return metaMap[exactKey];
  const formKey = String(formName || '').trim().toLowerCase();
  const stateFilter = normalizeStateCompareKey(state);
  const siteFilter = String(site || '').trim();
  let found = null;
  Object.entries(metaMap || {}).forEach(([key, value]) => {
    if (found) return;
    const parts = key.split('|');
    if (parts.length < 3) return;
    const [keyState, keySite, keyForm] = parts;
    if (keyForm !== formKey) return;
    if (stateFilter && normalizeStateCompareKey(keyState) !== stateFilter) return;
    if (siteFilter && !sitesMatchFilter(keySite, siteFilter)) return;
    found = value;
  });
  return found || {};
}

function exportAccessRowsCsv(rows) {
  const headers = ['Role Name', 'Act', 'Industry Type', 'Description', 'Form Name', 'Email ID', 'Status'];
  const escape = (value) => `"${String(value ?? '').replace(/"/g, '""')}"`;
  const lines = [headers.map(escape).join(',')];
  rows.forEach((row) => {
    lines.push(
      [
        row.roleName,
        row.act,
        row.industryType,
        row.description,
        row.formName,
        row.emailDisplay,
        row.status,
      ]
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

function UserChipMultiSelect({ options, value, onChange, placeholder = 'Select users' }) {
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

  const removeEmail = (email) => {
    onChange(selected.filter((item) => item.toLowerCase() !== email.toLowerCase()));
  };

  const hiddenCount = Math.max(0, selected.length - 2);

  return (
    <div
      className={`setup-chip-multi-select${open ? ' setup-chip-multi-select--open' : ''}`}
      ref={rootRef}
    >
      <button
        type="button"
        className="setup-chip-multi-select-trigger"
        onClick={() => setOpen((prev) => !prev)}
        aria-haspopup="listbox"
        aria-expanded={open}
        title={selected.length ? selected.join(', ') : undefined}
      >
        <span className="setup-chip-multi-select-chips">
          {selected.length === 0 ? (
            <span className="setup-chip-multi-select-placeholder">{placeholder}</span>
          ) : (
            <>
              {selected.slice(0, 2).map((email) => (
                <span key={email} className="setup-user-chip">
                  <span className="setup-user-chip-label">{email}</span>
                  <button
                    type="button"
                    className="setup-user-chip-remove"
                    onClick={(event) => {
                      event.stopPropagation();
                      removeEmail(email);
                    }}
                    aria-label={`Remove ${email}`}
                  >
                    <X size={12} strokeWidth={2.5} aria-hidden />
                  </button>
                </span>
              ))}
              {hiddenCount > 0 ? (
                <span className="setup-user-chip setup-user-chip--more">+{hiddenCount}</span>
              ) : null}
            </>
          )}
        </span>
        <span className="setup-multi-select-caret" aria-hidden>
          ▾
        </span>
      </button>
      {open ? (
        <div className="setup-multi-select-menu" role="listbox" aria-multiselectable="true">
          {options.length === 0 ? (
            <div className="setup-multi-select-empty">No users for this scope</div>
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
  const [selectedState, setSelectedState] = useState(() => readStoredFilters()?.state || '');
  const [selectedSite, setSelectedSite] = useState(() => readStoredFilters()?.site || '');
  const [selectedIndustryType, setSelectedIndustryType] = useState(
    () => readStoredFilters()?.industryType || ''
  );
  const [statusFilter, setStatusFilter] = useState(() => readStoredFilters()?.status || 'All');
  const [emails, setEmails] = useState(() => readStoredEmails());
  const [setupMeta, setSetupMeta] = useState(() => readStoredMeta());
  const [saving, setSaving] = useState(false);
  const [currentPage, setCurrentPage] = useState(1);
  const [showCreateRole, setShowCreateRole] = useState(false);
  const [createRoleForm, setCreateRoleForm] = useState({
    roleName: '',
    status: 'Active',
    industryType: '',
    state: '',
    site: '',
    formUsers: {},
  });
  const [createRoleMessage, setCreateRoleMessage] = useState('');
  const [createRoleSaving, setCreateRoleSaving] = useState(false);

  const loadData = useCallback(async ({ preserveMessage = false } = {}) => {
    setLoading(true);
    if (!preserveMessage) setMessage('');
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
        const hasMappedMeta = Object.keys(metaFromBackend).length > 0;
        const hasMappedEmails = Object.keys(fromBackend).length > 0;

        setEmails((prev) => {
          const merged =
            hasMappedEmails || hasMappedMeta
              ? mergeEmailMaps(prev, fromBackend)
              : Object.keys(prev).length
                ? prev
                : readStoredEmails();
          writeStoredEmails(merged);
          return merged;
        });
        setSetupMeta((prev) => {
          const merged =
            hasMappedMeta
              ? mergeMetaMaps(prev, metaFromBackend)
              : Object.keys(prev).length
                ? prev
                : readStoredMeta();
          writeStoredMeta(merged);
          return merged;
        });
        return setupData.data;
      }
      setEmails((prev) => {
        const cached = Object.keys(prev).length ? prev : readStoredEmails();
        writeStoredEmails(cached);
        return cached;
      });
      setSetupMeta((prev) => {
        const cached = Object.keys(prev).length ? prev : readStoredMeta();
        writeStoredMeta(cached);
        return cached;
      });
      return [];
    } catch (err) {
      if (!preserveMessage) setMessage(err?.message || 'Failed to load Setup data.');
      return [];
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    loadData();
  }, [loadData]);

  useEffect(() => {
    writeStoredFilters({
      status: statusFilter,
      industryType: selectedIndustryType,
      state: selectedState,
      site: selectedSite,
    });
  }, [statusFilter, selectedIndustryType, selectedState, selectedSite]);

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
    if (!selectedState || stateOptions.length === 0) return;
    const stillValid = stateOptions.some(
      (option) =>
        option === selectedState ||
        normalizeStateCompareKey(option) === normalizeStateCompareKey(selectedState) ||
        checklistStateMatchesSiteState(option, selectedState)
    );
    if (!stillValid) {
      setSelectedState('');
      setSelectedSite('');
    }
  }, [selectedState, stateOptions]);

  useEffect(() => {
    if (!selectedSite || siteOptions.length === 0) return;
    const stillValid = siteOptions.some(
      (option) =>
        option === selectedSite ||
        option.toLowerCase() === selectedSite.toLowerCase() ||
        sitesMatchFilter(option, selectedSite)
    );
    if (!stillValid) {
      setSelectedSite('');
    }
  }, [selectedSite, siteOptions]);

  const adminEmail = String(userEmail || '').trim();

  const sitesForScope = useMemo(() => {
    if (selectedSite) {
      return sitesForState.filter((site) => siteNameFromSiteRecord(site) === selectedSite);
    }
    if (selectedState) return sitesForState;
    return sites;
  }, [selectedSite, selectedState, sitesForState, sites]);

  const defaultAssignedEmails = useMemo(() => {
    const bucket = [];
    const admin = isValidEmail(adminEmail) ? adminEmail : '';
    sitesForScope.forEach((site) => {
      parseStoredEmails(siteInchargeEmail(site)).forEach((email) => bucket.push(email));
    });
    const merged = uniqueSorted(bucket);
    if (!admin) return merged;
    const withoutAdmin = merged.filter((email) => email.toLowerCase() !== admin.toLowerCase());
    return [admin, ...withoutAdmin];
  }, [adminEmail, sitesForScope]);

  const allEmailOptions = useMemo(() => {
    const bucket = [...defaultAssignedEmails];
    const scopeState = normalizeStateCompareKey(selectedState);
    Object.entries(emails).forEach(([key, value]) => {
      const parts = key.split('|');
      if (parts.length < 3) return;
      const [keyState, keySite] = parts;
      if (scopeState && normalizeStateCompareKey(keyState) !== scopeState) return;
      if (selectedSite && !sitesMatchFilter(keySite, selectedSite)) return;
      parseStoredEmails(value).forEach((email) => bucket.push(email));
    });
    sitesForScope.forEach((site) => {
      parseStoredEmails(siteInchargeEmail(site)).forEach((email) => bucket.push(email));
    });
    const merged = uniqueSorted(bucket);
    const admin = isValidEmail(adminEmail) ? adminEmail : '';
    if (!admin) return merged;
    const withoutAdmin = merged.filter((email) => email.toLowerCase() !== admin.toLowerCase());
    return [admin, ...withoutAdmin];
  }, [defaultAssignedEmails, emails, selectedState, selectedSite, sitesForScope, adminEmail]);

  const accessRows = useMemo(
    () =>
      buildSavedAccessRows(setupMeta, emails, formRecords, {
        state: selectedState,
        site: selectedSite,
      }),
    [setupMeta, emails, formRecords, selectedState, selectedSite]
  );

  const displayedRows = useMemo(() => {
    let rows = accessRows;
    if (statusFilter !== 'All') {
      rows = rows.filter((row) => row.status === statusFilter);
    }
    if (selectedIndustryType) {
      const wantCategory = industryLabelToActCategory(selectedIndustryType);
      const wantLabel = String(selectedIndustryType).trim().toLowerCase();
      rows = rows.filter((row) => {
        const haveLabel = String(row.industryType || '').trim().toLowerCase();
        if (haveLabel) {
          if (haveLabel === wantLabel) return true;
          const haveCategory = industryLabelToActCategory(row.industryType);
          if (wantCategory && haveCategory && haveCategory === wantCategory) return true;
        }
        if (wantCategory) {
          const actCategory = getActCategoryFromActSector(row.act, '');
          if (actCategory && actCategory !== 'other' && actCategory === wantCategory) return true;
          const siteRec = sites.find((site) =>
            sitesMatchFilter(siteNameFromSiteRecord(site), row.rowSite)
          );
          if (siteRec) {
            const siteCategory = industryLabelToActCategory(siteIndustry(siteRec));
            if (siteCategory && siteCategory === wantCategory) return true;
          }
          // IndustryType was not persisted on older rows — keep them visible
          // when state/site already match the selected scope.
          if (!haveLabel) return true;
        }
        return false;
      });
    }
    return rows;
  }, [accessRows, statusFilter, selectedIndustryType, sites]);

  const totalRows = displayedRows.length;
  const totalPages = Math.max(1, Math.ceil(totalRows / ROWS_PER_PAGE));
  const effectivePage = Math.min(currentPage, totalPages);
  const paginatedRows = useMemo(
    () => displayedRows.slice((effectivePage - 1) * ROWS_PER_PAGE, effectivePage * ROWS_PER_PAGE),
    [displayedRows, effectivePage]
  );
  const paginationItems = useMemo(
    () => buildPaginationItems(effectivePage, totalPages),
    [effectivePage, totalPages]
  );

  useEffect(() => {
    setCurrentPage((page) => Math.min(page, totalPages));
  }, [totalPages]);

  useEffect(() => {
    setCurrentPage(1);
  }, [statusFilter, selectedState, selectedSite, selectedIndustryType]);

  const handleEmailChange = (
    formName,
    nextEmails,
    {
      state = selectedState,
      site = selectedSite,
      industryType = '',
      act = '',
      description = '',
    } = {}
  ) => {
    const key = emailStorageKey(state, site, formName, industryType, act, description);
    setEmails((prev) => {
      const next = { ...prev, [key]: nextEmails };
      writeStoredEmails(next);
      return next;
    });
  };

  const handleRoleChange = (
    formName,
    role,
    {
      state = selectedState,
      site = selectedSite,
      act = '',
      description = '',
      industryType = '',
    } = {}
  ) => {
    const key = emailStorageKey(state, site, formName, industryType, act, description);
    setSetupMeta((prev) => ({
      ...prev,
      [key]: {
        ...(prev[key] || {}),
        formName: String(formName || prev[key]?.formName || '').trim(),
        role: String(role || '').trim(),
        state: String(state || prev[key]?.state || '').trim(),
        site: String(site || prev[key]?.site || '').trim(),
        act: String(act || prev[key]?.act || '').trim(),
        description: String(description || prev[key]?.description || '').trim(),
        industryType: String(industryType || prev[key]?.industryType || '').trim(),
      },
    }));
  };

  const handleReset = () => {
    setStatusFilter('All');
    setSelectedIndustryType('');
    setSelectedState('');
    setSelectedSite('');
    setCurrentPage(1);
    setMessage('');
    writeStoredFilters({ status: 'All', industryType: '', state: '', site: '' });
  };

  const openCreateRole = () => {
    setCreateRoleForm({
      roleName: '',
      status: 'Active',
      industryType: selectedIndustryType || '',
      state: selectedState || '',
      site: selectedSite || '',
      formUsers: {},
    });
    setCreateRoleMessage('');
    setShowCreateRole(true);
  };

  const closeCreateRole = () => {
    setShowCreateRole(false);
    setCreateRoleMessage('');
  };

  const createRoleSitesForScope = useMemo(() => {
    let scoped = sites;
    if (createRoleForm.industryType) {
      const wantCategory = industryLabelToActCategory(createRoleForm.industryType);
      scoped = scoped.filter((site) => {
        const siteCategory = industryLabelToActCategory(siteIndustry(site));
        return !wantCategory || siteCategory === wantCategory;
      });
    }
    if (createRoleForm.state) {
      scoped = scoped.filter((site) =>
        checklistStateMatchesSiteState(siteStateFromRecord(site), createRoleForm.state)
      );
    }
    if (createRoleForm.site) {
      scoped = scoped.filter((site) => siteNameFromSiteRecord(site) === createRoleForm.site);
    }
    return scoped;
  }, [sites, createRoleForm.industryType, createRoleForm.state, createRoleForm.site]);

  const createRoleStateOptions = useMemo(() => {
    let scoped = sites;
    if (createRoleForm.industryType) {
      const wantCategory = industryLabelToActCategory(createRoleForm.industryType);
      scoped = scoped.filter((site) => {
        const siteCategory = industryLabelToActCategory(siteIndustry(site));
        return !wantCategory || siteCategory === wantCategory;
      });
    }
    return uniqueSorted(scoped.map(siteStateFromRecord));
  }, [sites, createRoleForm.industryType]);

  const createRoleSiteOptions = useMemo(() => {
    let scoped = sites;
    if (createRoleForm.industryType) {
      const wantCategory = industryLabelToActCategory(createRoleForm.industryType);
      scoped = scoped.filter((site) => {
        const siteCategory = industryLabelToActCategory(siteIndustry(site));
        return !wantCategory || siteCategory === wantCategory;
      });
    }
    if (createRoleForm.state) {
      scoped = scoped.filter((site) =>
        checklistStateMatchesSiteState(siteStateFromRecord(site), createRoleForm.state)
      );
    }
    return uniqueSorted(scoped.map(siteNameFromSiteRecord));
  }, [sites, createRoleForm.industryType, createRoleForm.state]);

  const createRoleEmailOptions = useMemo(() => {
    const bucket = [];
    const admin = isValidEmail(adminEmail) ? adminEmail : '';
    createRoleSitesForScope.forEach((site) => {
      parseStoredEmails(siteInchargeEmail(site)).forEach((email) => bucket.push(email));
    });
    const merged = uniqueSorted(bucket);
    if (!admin) return merged;
    const withoutAdmin = merged.filter((email) => email.toLowerCase() !== admin.toLowerCase());
    return [admin, ...withoutAdmin];
  }, [createRoleSitesForScope, adminEmail]);

  const createRoleForms = useMemo(() => {
    const wantCategory = createRoleForm.industryType
      ? industryLabelToActCategory(createRoleForm.industryType)
      : null;
    const rows = createRoleForm.state
      ? formRecords.filter((row) => checklistStateMatchesSiteState(row.state, createRoleForm.state))
      : formRecords;
    const seen = new Set();
    const unique = [];
    rows.forEach((row) => {
      const formName = String(row.formName || row.FormName || '').trim();
      if (!formName) return;
      const act = String(row.act || row.Act || '').trim();
      const description = String(row.description || row.Description || '').trim();
      if (wantCategory) {
        const rowCategory = getActCategoryFromActSector(row.act || row.Act, row.sector || row.Sector);
        if (rowCategory && rowCategory !== wantCategory) return;
      }
      const key = [
        formName.toLowerCase(),
        act.toLowerCase(),
        description.toLowerCase(),
        String(createRoleForm.state || '').trim().toLowerCase(),
        String(createRoleForm.industryType || '').trim().toLowerCase(),
      ].join('|');
      if (seen.has(key)) return;
      seen.add(key);
      unique.push({
        id: row.id || key,
        formName,
        act,
        description,
      });
    });
    unique.sort((a, b) => a.formName.localeCompare(b.formName));
    return unique;
  }, [formRecords, createRoleForm.state, createRoleForm.industryType]);

  const updateCreateRoleField = (field, value) => {
    setCreateRoleForm((prev) => {
      const next = { ...prev, [field]: value };
      if (field === 'industryType') {
        next.state = '';
        next.site = '';
        next.formUsers = {};
      } else if (field === 'state') {
        next.site = '';
        next.formUsers = {};
      } else if (field === 'site') {
        next.formUsers = {};
      } else if (field === 'status') {
        next.formUsers = {};
      }
      return next;
    });
    setCreateRoleMessage('');
  };

  const updateCreateRoleFormUsers = (formName, nextEmails, formMeta = {}) => {
    const userKey = [
      String(formName || '').trim().toLowerCase(),
      String(formMeta.act || '').trim().toLowerCase(),
      String(formMeta.description || '').trim().toLowerCase(),
    ].join('|');
    setCreateRoleForm((prev) => ({
      ...prev,
      formUsers: {
        ...prev.formUsers,
        [userKey]: nextEmails,
      },
    }));
  };

  const getCreateRoleFormUserValue = (form) => {
    const formName = typeof form === 'string' ? form : form?.formName;
    const act = typeof form === 'string' ? '' : form?.act || '';
    const description = typeof form === 'string' ? '' : form?.description || '';
    if (createRoleForm.status !== 'Active') return [];
    const userKey = [
      String(formName || '').trim().toLowerCase(),
      String(act || '').trim().toLowerCase(),
      String(description || '').trim().toLowerCase(),
    ].join('|');
    if (Object.prototype.hasOwnProperty.call(createRoleForm.formUsers, userKey)) {
      return parseStoredEmails(createRoleForm.formUsers[userKey]);
    }
    if (Object.prototype.hasOwnProperty.call(createRoleForm.formUsers, formName)) {
      return parseStoredEmails(createRoleForm.formUsers[formName]);
    }
    return createRoleEmailOptions;
  };

  const handleCreateRoleSave = async () => {
    const roleName = String(createRoleForm.roleName || '').trim();
    const state = String(createRoleForm.state || '').trim();
    const site = String(createRoleForm.site || '').trim();

    if (!roleName) {
      setCreateRoleMessage('Role name is required.');
      return;
    }
    if (!createRoleForm.industryType) {
      setCreateRoleMessage('Select an industry type.');
      return;
    }
    if (!state) {
      setCreateRoleMessage('Select a state.');
      return;
    }
    if (!site) {
      setCreateRoleMessage('Select a site.');
      return;
    }
    if (!createRoleForms.length) {
      setCreateRoleMessage('No forms found for the selected scope.');
      return;
    }

    const industryType = String(createRoleForm.industryType || '').trim();
    const isActive = createRoleForm.status === 'Active';
    const rows = createRoleForms
      .map((form) => ({
        formName: form.formName,
        act: form.act,
        description: form.description,
        role: roleName,
        industryType,
        emails: isActive ? getCreateRoleFormUserValue(form) : [],
      }))
      .filter((row) => row.emails.length > 0);

    if (!rows.length) {
      setCreateRoleMessage('Assign at least one user to at least one form.');
      return;
    }

    const next = { ...emails };
    const nextMeta = { ...setupMeta };
    rows.forEach((row) => {
      const key = emailStorageKey(
        state,
        site,
        row.formName,
        row.industryType,
        row.act,
        row.description
      );
      next[key] = row.emails;
      nextMeta[key] = {
        formName: row.formName,
        act: row.act,
        description: row.description,
        role: row.role,
        industryType: row.industryType,
        state,
        site,
        emails: row.emails,
      };
    });
    setEmails(next);
    setSetupMeta(nextMeta);
    persistSetupMaps(next, nextMeta);

    setCreateRoleSaving(true);
    setCreateRoleMessage('');
    try {
      const res = await fetch(SETUP_API, {
        method: 'PUT',
        credentials: 'include',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          state,
          site,
          industryType,
          rows,
        }),
      });
      const data = await res.json().catch(() => null);
      if (!res.ok || data?.status !== 'success') {
        throw new Error(data?.message || 'Failed to create role.');
      }

      const savedRows = Array.isArray(data.data) ? data.data : [];
      if (savedRows.length) {
        const enriched = savedRows.map((row) => ({
          ...row,
          formName: row.formName || row.FormName,
          role: row.role || roleName,
          industryType: row.industryType || industryType,
          state: row.state || state,
          site: row.site || site,
          act: row.act || '',
          description: row.description || '',
          email: row.email || (Array.isArray(row.emails) ? row.emails.join(', ') : ''),
        }));
        const savedMeta = mapSetupRowsToMeta(enriched);
        const savedEmails = mapSetupRowsToEmails(enriched);
        setSetupMeta((prev) => {
          const merged = mergeMetaMaps({ ...prev, ...nextMeta }, savedMeta);
          writeStoredMeta(merged);
          return merged;
        });
        setEmails((prev) => {
          const merged = mergeEmailMaps({ ...prev, ...next }, savedEmails);
          writeStoredEmails(merged);
          return merged;
        });
      }

      const successMessage = data.message || `Role "${roleName}" created successfully.`;
      setStatusFilter('All');
      setSelectedIndustryType(industryType);
      setSelectedState(state);
      setSelectedSite(site);
      setMessage(successMessage);
      closeCreateRole();
      await loadData({ preserveMessage: true });
      // Keep optimistic create data if GET is empty / lagging; backend rows win on key conflict
      setSetupMeta((prev) => {
        const merged = mergeMetaMaps(nextMeta, prev);
        writeStoredMeta(merged);
        return merged;
      });
      setEmails((prev) => {
        const merged = mergeEmailMaps(next, prev);
        writeStoredEmails(merged);
        return merged;
      });
      setSelectedIndustryType(industryType);
      setSelectedState(state);
      setSelectedSite(site);
      setStatusFilter('All');
      setMessage(successMessage);
    } catch (err) {
      setCreateRoleMessage(err?.message || 'Failed to create role.');
    } finally {
      setCreateRoleSaving(false);
    }
  };

  const handleSave = async () => {
    if (!accessRows.length) {
      setMessage('No roles to save. Create a role first.');
      return;
    }

    const groups = new Map();
    accessRows.forEach((row) => {
      const state = row.rowState || selectedState;
      const site = row.rowSite || selectedSite;
      if (!state) return;
      const groupKey = `${state}|||${site}`;
      if (!groups.has(groupKey)) {
        groups.set(groupKey, { state, site, rows: [] });
      }
      groups.get(groupKey).rows.push({
        formName: row.formName,
        act: row.act,
        description: row.description,
        role: row.roleName,
        industryType: row.industryType || '',
        emails: row.emails,
      });
    });

    if (!groups.size) {
      setMessage('Select a state before saving.');
      return;
    }

    const next = { ...emails };
    const nextMeta = { ...setupMeta };
    accessRows.forEach((row) => {
      const state = row.rowState || selectedState;
      const site = row.rowSite || selectedSite;
      const key = emailStorageKey(
        state,
        site,
        row.formName,
        row.industryType,
        row.act,
        row.description
      );
      next[key] = row.emails;
      nextMeta[key] = {
        formName: row.formName,
        act: row.act,
        description: row.description,
        role: row.roleName,
        industryType: row.industryType || '',
        state,
        site,
        emails: row.emails,
      };
    });
    setEmails(next);
    setSetupMeta(nextMeta);
    persistSetupMaps(next, nextMeta);

    setSaving(true);
    setMessage('');
    try {
      let totalSaved = 0;
      for (const group of groups.values()) {
        const res = await fetch(SETUP_API, {
          method: 'PUT',
          credentials: 'include',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            state: group.state,
            site: group.site,
            rows: group.rows,
          }),
        });
        const data = await res.json().catch(() => null);
        if (!res.ok || data?.status !== 'success') {
          throw new Error(data?.message || 'Failed to save Setup data.');
        }
        totalSaved += Array.isArray(data.data) ? data.data.length : group.rows.length;
      }
      const successMessage = `Saved ${totalSaved} Setup record(s).`;
      setMessage(successMessage);
      await loadData({ preserveMessage: true });
      setMessage(successMessage);
    } catch (err) {
      setMessage(err?.message || 'Failed to save Setup data.');
    } finally {
      setSaving(false);
    }
  };

  const footerStart = totalRows === 0 ? 0 : (effectivePage - 1) * ROWS_PER_PAGE + 1;
  const footerEnd = totalRows === 0 ? 0 : Math.min(effectivePage * ROWS_PER_PAGE, totalRows);

  return (
    <div className="setup-rba-page">
      {showCreateRole ? (
        <div className="setup-rba-card setup-create-role-card">
          <nav className="setup-create-role-breadcrumb" aria-label="Breadcrumb">
            <button type="button" className="setup-create-role-breadcrumb-link" onClick={closeCreateRole}>
              Role Based Access
            </button>
            <ChevronRight size={14} strokeWidth={2} aria-hidden />
            <span>Create Role</span>
          </nav>

          <div className="setup-create-role-header">
            <div>
              <h1 className="setup-rba-title">Create Role</h1>
              <p className="setup-rba-subtitle">
                Add role details, select scope and assign form access with users.
              </p>
            </div>
          </div>

          <section className="setup-create-role-section">
            <h2 className="setup-create-role-section-title">
              <span className="setup-create-role-section-num">1</span>
              Role Information
            </h2>
            <div className="setup-create-role-grid setup-create-role-grid--2">
              <label className="setup-create-role-field">
                <span className="setup-create-role-label">
                  Role Name <span className="setup-create-role-required">*</span>
                </span>
                <input
                  type="text"
                  className="setup-rba-search-input setup-create-role-input"
                  value={createRoleForm.roleName}
                  onChange={(e) => updateCreateRoleField('roleName', e.target.value)}
                  placeholder="Enter role name"
                />
              </label>
              <label className="setup-create-role-field">
                <span className="setup-create-role-label">
                  Status <span className="setup-create-role-required">*</span>
                </span>
                <select
                  className="setup-rba-select"
                  value={createRoleForm.status}
                  onChange={(e) => updateCreateRoleField('status', e.target.value)}
                >
                  <option value="Active">Active</option>
                  <option value="Inactive">Inactive</option>
                </select>
              </label>
            </div>
          </section>

          <section className="setup-create-role-section">
            <h2 className="setup-create-role-section-title">
              <span className="setup-create-role-section-num">2</span>
              Scope / Access
            </h2>
            <div className="setup-create-role-grid setup-create-role-grid--3">
              <label className="setup-create-role-field">
                <span className="setup-create-role-label">
                  Industry Type <span className="setup-create-role-required">*</span>
                </span>
                <select
                  className="setup-rba-select"
                  value={createRoleForm.industryType}
                  onChange={(e) => updateCreateRoleField('industryType', e.target.value)}
                >
                  <option value="">Select industry type</option>
                  {INDUSTRY_OPTIONS.map((option) => (
                    <option key={option} value={option}>
                      {option}
                    </option>
                  ))}
                </select>
              </label>
              <label className="setup-create-role-field">
                <span className="setup-create-role-label">
                  State <span className="setup-create-role-required">*</span>
                </span>
                <select
                  className="setup-rba-select"
                  value={createRoleForm.state}
                  onChange={(e) => updateCreateRoleField('state', e.target.value)}
                  disabled={!createRoleForm.industryType}
                >
                  <option value="">Select state</option>
                  {createRoleStateOptions.map((state) => (
                    <option key={state} value={state}>
                      {state}
                    </option>
                  ))}
                </select>
              </label>
              <label className="setup-create-role-field">
                <span className="setup-create-role-label">
                  Site <span className="setup-create-role-required">*</span>
                </span>
                <select
                  className="setup-rba-select"
                  value={createRoleForm.site}
                  onChange={(e) => updateCreateRoleField('site', e.target.value)}
                  disabled={!createRoleForm.state}
                >
                  <option value="">Select site</option>
                  {createRoleSiteOptions.map((site) => (
                    <option key={site} value={site}>
                      {site}
                    </option>
                  ))}
                </select>
              </label>
            </div>
          </section>

          <section className="setup-create-role-section">
            <h2 className="setup-create-role-section-title">
              <span className="setup-create-role-section-num">3</span>
              Form Access &amp; Users
            </h2>
            <p className="setup-create-role-section-hint">Assign users who can access each form.</p>

            {createRoleMessage ? (
              <p
                className={`setup-rba-message${/fail|required|select/i.test(createRoleMessage) ? ' setup-rba-message--error' : ''}`}
                role="status"
              >
                {createRoleMessage}
              </p>
            ) : null}

            <div className="setup-rba-table-wrap setup-create-role-table-wrap">
              <table className="setup-rba-table setup-create-role-table">
                <thead>
                  <tr>
                    <th>Act</th>
                    <th>Description</th>
                    <th>Form Name</th>
                    <th>Users</th>
                  </tr>
                </thead>
                <tbody>
                  {!createRoleForm.state ? (
                    <tr>
                      <td colSpan={4} className="setup-rba-empty">
                        Select state and site to load forms.
                      </td>
                    </tr>
                  ) : createRoleForms.length === 0 ? (
                    <tr>
                      <td colSpan={4} className="setup-rba-empty">
                        No forms found for the selected scope.
                      </td>
                    </tr>
                  ) : (
                    createRoleForms.map((form) => (
                      <tr key={form.id}>
                        <td className="setup-create-role-act-cell">{form.act || '—'}</td>
                        <td className="setup-create-role-desc-cell">{form.description || '—'}</td>
                        <td className="setup-create-role-form-cell">{form.formName}</td>
                        <td>
                          <UserChipMultiSelect
                            options={createRoleEmailOptions}
                            value={getCreateRoleFormUserValue(form)}
                            onChange={(nextEmails) =>
                              updateCreateRoleFormUsers(form.formName, nextEmails, {
                                act: form.act,
                                description: form.description,
                              })
                            }
                            placeholder="Select users"
                          />
                        </td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>

            <p className="setup-create-role-info">
              <Info size={14} strokeWidth={2} aria-hidden />
              Users added here will have access to the respective forms.
            </p>
          </section>

          <div className="setup-create-role-actions">
            <button
              type="button"
              className="setup-rba-btn setup-rba-btn--export"
              onClick={closeCreateRole}
              disabled={createRoleSaving}
            >
              Cancel
            </button>
            <button
              type="button"
              className="setup-rba-btn setup-rba-btn--save"
              onClick={handleCreateRoleSave}
              disabled={createRoleSaving}
            >
              {createRoleSaving ? 'Saving…' : 'Create Role'}
            </button>
          </div>
        </div>
      ) : (
      <div className="setup-rba-card">
        <div className="setup-rba-header">
          <div className="setup-rba-header-text">
            <h1 className="setup-rba-title">Role Based Access</h1>
            <p className="setup-rba-subtitle">Manage roles, forms and user access for HCM.</p>
          </div>
          <div className="setup-rba-header-actions">
            <button
              type="button"
              className="setup-rba-btn setup-rba-btn--export"
              onClick={() => exportAccessRowsCsv(displayedRows)}
              disabled={!displayedRows.length}
            >
              <Download size={16} strokeWidth={2} aria-hidden />
              Export
            </button>
            <button
              type="button"
              className="setup-rba-btn setup-rba-btn--create"
              onClick={openCreateRole}
            >
              <Plus size={16} strokeWidth={2} aria-hidden />
              Create Role
            </button>
          </div>
        </div>

        <div className="setup-rba-toolbar">
          <label className="setup-rba-field setup-rba-field--status">
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
            <span className="setup-rba-field-label">Industry Type</span>
            <select
              className="setup-rba-select"
              value={selectedIndustryType}
              onChange={(e) => {
                setSelectedIndustryType(e.target.value);
                setMessage('');
              }}
            >
              <option value="">All industry types</option>
              {INDUSTRY_OPTIONS.map((option) => (
                <option key={option} value={option}>
                  {option}
                </option>
              ))}
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
          {loading ? (
            <p className="setup-rba-empty">Loading…</p>
          ) : (
            <table className="setup-rba-table">
              <thead>
                <tr>
                  <th>Role Name</th>
                  <th>Act</th>
                  <th>Industry Type</th>
                  <th>Description</th>
                  <th>Form Name</th>
                  <th>Email ID</th>
                  <th>Status</th>
                </tr>
              </thead>
              <tbody>
                {displayedRows.length === 0 ? (
                  <tr>
                    <td colSpan={7} className="setup-rba-empty">
                      {selectedState || selectedIndustryType || selectedSite || accessRows.length > 0
                        ? 'No roles found for the selected filters. Click Create Role to add one.'
                        : 'No roles yet. Click Create Role to add one, or adjust filters.'}
                    </td>
                  </tr>
                ) : (
                  paginatedRows.map((row) => (
                    <tr key={row.id}>
                      <td>
                        <input
                          type="text"
                          className="setup-rba-cell-input"
                          value={row.roleName}
                          onChange={(e) =>
                            handleRoleChange(row.formName, e.target.value, {
                              state: row.rowState,
                              site: row.rowSite,
                              act: row.act,
                              description: row.description,
                              industryType: row.industryType,
                            })
                          }
                          placeholder="Enter role name"
                          aria-label={`Role name for ${row.formName}`}
                        />
                      </td>
                      <td>{row.act || '—'}</td>
                      <td>{row.industryType || '—'}</td>
                      <td>{row.description || '—'}</td>
                      <td>{row.formName}</td>
                      <td>
                        <EmailMultiSelect
                          options={selectedState || selectedSite ? allEmailOptions : row.emails}
                          value={row.emails}
                          onChange={(nextEmails) =>
                            handleEmailChange(row.formName, nextEmails, {
                              state: row.rowState,
                              site: row.rowSite,
                              industryType: row.industryType,
                              act: row.act,
                              description: row.description,
                            })
                          }
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
          <span className="setup-rba-footer-info">
            Showing {footerStart} to {footerEnd} of {totalRows} entries
          </span>
          {totalRows > 0 ? (
            <nav className="setup-rba-pagination" aria-label="Table pagination">
              <button
                type="button"
                className="setup-rba-pagination-nav"
                disabled={effectivePage <= 1}
                onClick={() => setCurrentPage(1)}
                title="First page"
                aria-label="First page"
              >
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden>
                  <polyline points="11 17 6 12 11 7" />
                  <polyline points="18 17 13 12 18 7" />
                </svg>
              </button>
              <button
                type="button"
                className="setup-rba-pagination-nav"
                disabled={effectivePage <= 1}
                onClick={() => setCurrentPage((page) => Math.max(1, page - 1))}
                title="Previous page"
                aria-label="Previous page"
              >
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden>
                  <polyline points="15 18 9 12 15 6" />
                </svg>
              </button>
              <div className="setup-rba-pagination-pages">
                {paginationItems.map((item, index) =>
                  item === 'ellipsis' ? (
                    <span key={`ellipsis-${index}`} className="setup-rba-pagination-ellipsis" aria-hidden>
                      ...
                    </span>
                  ) : (
                    <button
                      key={item}
                      type="button"
                      className={`setup-rba-pagination-page${item === effectivePage ? ' setup-rba-pagination-page--active' : ''}`}
                      onClick={() => setCurrentPage(item)}
                      aria-label={`Page ${item}`}
                      aria-current={item === effectivePage ? 'page' : undefined}
                    >
                      {item}
                    </button>
                  )
                )}
              </div>
              <button
                type="button"
                className="setup-rba-pagination-nav"
                disabled={effectivePage >= totalPages}
                onClick={() => setCurrentPage((page) => Math.min(totalPages, page + 1))}
                title="Next page"
                aria-label="Next page"
              >
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden>
                  <polyline points="9 18 15 12 9 6" />
                </svg>
              </button>
              <button
                type="button"
                className="setup-rba-pagination-nav"
                disabled={effectivePage >= totalPages}
                onClick={() => setCurrentPage(totalPages)}
                title="Last page"
                aria-label="Last page"
              >
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden>
                  <polyline points="13 17 18 12 13 7" />
                  <polyline points="6 17 11 12 6 7" />
                </svg>
              </button>
            </nav>
          ) : null}
        </div>
      </div>
      )}
    </div>
  );
}
