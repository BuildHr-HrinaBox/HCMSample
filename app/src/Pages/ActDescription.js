import React, { useState, useEffect, useCallback, useMemo } from 'react';
import { Link } from 'react-router-dom';
import {
  Briefcase,
  Scale,
  MapPin,
  Calendar,
  Check,
  Clock,
  Play,
  FileText,
  Filter,
  RefreshCw,
  MoreVertical,
  ChevronRight,
  LayoutGrid,
  Table2,
  Search,
  Pencil,
  Trash2,
  ArrowUpDown,
  Upload,
  FileStack,
} from 'lucide-react';
import './CompanyDetails.css';
import './ActDescription.css';
import {
  fetchInchargeDisplayScopeFromSites,
  getActCategoryFromActSector,
  sectorMatchesInchargeSiteIndustries,
  statesFieldMatchesInchargeSiteStates,
} from '../utils/siteInchargeScope';

const API_BASE = '/server/actdescriptions_function';
const ACTSBULK_API_BASE = '/server/actsbulk_function';

const STATUS_OPTIONS = ['yet to start', 'in progress', 'completed'];

function formatDueDateAct(row) {
  const s = String(row?.dueDate ?? '').trim();
  if (!s) return '—';
  const d = new Date(s);
  if (!Number.isNaN(d.getTime())) {
    return d.toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' });
  }
  return s;
}

function normalizeActType(row) {
  const t = String(row?.type ?? '')
    .trim()
    .toLowerCase();
  if (!t) return '';
  if (t.startsWith('central') || t === 'national' || t.includes('all india')) return 'Central';
  if (t.startsWith('state')) return 'State';
  return '';
}

function rowMatchesTypeFilter(row, filterType) {
  if (filterType === 'all') return true;
  const n = normalizeActType(row);
  if (filterType === 'Central') return n === 'Central';
  if (filterType === 'State') return n === 'State';
  return true;
}

function registerCountCell(registers) {
  const s = String(registers ?? '').trim();
  if (!s) return null;
  const n = parseInt(s, 10);
  if (String(n) === s && Number.isFinite(n)) return n;
  const parts = s.split(/[,;]/).map((x) => x.trim()).filter(Boolean);
  if (parts.length > 1) return parts.length;
  const digits = s.match(/\d+/);
  if (digits) return parseInt(digits[0], 10);
  return 1;
}

function truncateTableText(str, max = 80) {
  const t = String(str ?? '').trim();
  if (t.length <= max) return t;
  return `${t.slice(0, max)}…`;
}

function abbreviateStateLabel(statesStr) {
  const raw = String(statesStr ?? '')
    .trim()
    .split(/[,;/]/)[0]
    ?.trim();
  if (!raw) return '—';
  if (raw.length <= 3) return raw.toUpperCase();
  const words = raw.split(/\s+/).filter(Boolean);
  if (words.length >= 2) {
    return words
      .map((w) => w[0])
      .join('')
      .slice(0, 3)
      .toUpperCase();
  }
  return raw.slice(0, 2).toUpperCase();
}

function dueDateUrgency(row) {
  const s = String(row?.dueDate ?? '').trim();
  if (!s) return 'none';
  const d = new Date(s);
  if (Number.isNaN(d.getTime())) return 'none';
  const now = new Date();
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const due = new Date(d.getFullYear(), d.getMonth(), d.getDate());
  const diffDays = (due.getTime() - today.getTime()) / 86400000;
  if (diffDays < 0) return 'overdue';
  if (diffDays <= 7) return 'soon';
  return 'ok';
}

function sectorPillTone(sector) {
  const s = String(sector || '').toLowerCase();
  if (s.includes('shop')) return 'blue';
  if (s.includes('labour') || s.includes('labor')) return 'purple';
  return 'slate';
}

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
    if (i > 0 && sorted[i] - sorted[i - 1] > 1) out.push('ellipsis');
    out.push(sorted[i]);
  }
  return out;
}

const initialForm = {
  actName: '',
  description: '',
  type: '',
  state: '',
  sector: '',
  applicability: '',
  keyComplianceRequirements: '',
  dueDate: '',
  penaltyforNonCompliance: '',
  registers: '',
  status: 'yet to start',
};

const ActDescription = ({ userRole, userEmail }) => {
  const [list, setList] = useState([]);
  const [total, setTotal] = useState(0);
  const [loading, setLoading] = useState(false);
  const [message, setMessage] = useState('');
  const [showForm, setShowForm] = useState(false);
  const [editingId, setEditingId] = useState(null);
  const [form, setForm] = useState(initialForm);
  const [searchActName, setSearchActName] = useState('');
  const [page, setPage] = useState(1);
  const perPage = 20;
  /* Applicable Acts Library - data from actsbulk_function */
  const [actsBulkList, setActsBulkList] = useState([]);
  const [actsBulkLoading, setActsBulkLoading] = useState(false);
  const [statusMap, setStatusMap] = useState({}); // id -> 'Completed' | 'Pending' | 'Yet to Start'
  const [libraryFilter, setLibraryFilter] = useState('all'); // 'all' | 'completed' | 'pending' | 'yet-to-start'
  const [actDetailsModalAct, setActDetailsModalAct] = useState(null); // row object when modal open
  const [selectedLibraryIds, setSelectedLibraryIds] = useState(new Set()); // act card checkboxes
  const [statusFromActDescription, setStatusFromActDescription] = useState({}); // actName -> 'Completed'|'Pending'|'Yet to Start' from backend
  const [viewMode, setViewMode] = useState('card');
  const [libraryPage, setLibraryPage] = useState(1);
  const [libraryPageSize, setLibraryPageSize] = useState(8);
  const [filterSector, setFilterSector] = useState('');
  const [filterState, setFilterState] = useState('');
  const [filterType, setFilterType] = useState('all');
  const [openCardMenuId, setOpenCardMenuId] = useState(null);
  const [filterPanelOpen, setFilterPanelOpen] = useState(true);
  const [tableSearch, setTableSearch] = useState('');
  const [tableSortKey, setTableSortKey] = useState(null);
  const [tableSortDir, setTableSortDir] = useState('asc');
  const [inchargeDisplayScope, setInchargeDisplayScope] = useState({
    ready: false,
    actCategories: null,
    industryLabels: null,
    stateLabels: null,
  });
  const allowedActCategoryList = inchargeDisplayScope.actCategories;
  const hasSiteBasedScope = Array.isArray(allowedActCategoryList) && allowedActCategoryList.length > 0;

  useEffect(() => {
    let cancelled = false;
    (async () => {
      setInchargeDisplayScope({ ready: false, actCategories: null, industryLabels: null, stateLabels: null });
      try {
        const scope = await fetchInchargeDisplayScopeFromSites(userEmail);
        if (!cancelled) {
          setInchargeDisplayScope({
            ready: true,
            actCategories: scope.actCategories,
            industryLabels: scope.industryLabels,
            stateLabels: scope.stateLabels,
          });
        }
      } catch {
        if (!cancelled) {
          setInchargeDisplayScope({ ready: true, actCategories: null, industryLabels: null, stateLabels: null });
        }
      }
    })();
    return () => {
      cancelled = true;
    };
  }, [userEmail]);

  const fetchList = useCallback(async () => {
    setLoading(true);
    setMessage('');
    try {
      const params = new URLSearchParams();
      params.set('page', String(page));
      params.set('perPage', String(perPage));
      if (searchActName && searchActName.trim()) params.set('actName', searchActName.trim());
      const res = await fetch(`${API_BASE}/act-descriptions?${params.toString()}`);
      const data = await res.json();
      if (data.status === 'success' && data.data) {
        setList(data.data.actDescriptions || []);
        setTotal(data.data.total ?? 0);
      } else {
        setList([]);
        setTotal(0);
      }
    } catch (err) {
      setMessage('Failed to load act descriptions.');
      setList([]);
      setTotal(0);
    } finally {
      setLoading(false);
    }
  }, [page, searchActName]);

  useEffect(() => {
    fetchList();
  }, [fetchList]);

  const fetchActsBulk = useCallback(async (options = {}) => {
    const { useCacheFirst = false, silentRefresh = false } = options;

    if (useCacheFirst) {
      try {
        const cached = localStorage.getItem('actsDescriptionBulkData');
        if (cached) {
          const parsed = JSON.parse(cached);
          if (Array.isArray(parsed) && parsed.length > 0) {
            setActsBulkList(parsed);
          }
        }
      } catch (_) {}
    }

    if (!silentRefresh) setActsBulkLoading(true);
    try {
      const res = await fetch(`${ACTSBULK_API_BASE}/actsbulk?action=getAll`);
      const data = await res.json();
      if (data.status === 'success' && Array.isArray(data.data)) {
        setActsBulkList(data.data);
        localStorage.setItem('actsDescriptionBulkData', JSON.stringify(data.data));
        try {
          await fetch(`${API_BASE}/act-descriptions/sync-from-bulk`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ acts: data.data }),
          });
        } catch (syncErr) {
          console.error('Sync Actsbulk to ActDescription:', syncErr);
        }
        try {
          const descRes = await fetch(`${API_BASE}/act-descriptions`);
          const descData = await descRes.json();
          if (descData.status === 'success' && Array.isArray(descData.data?.actDescriptions)) {
            const map = {};
            descData.data.actDescriptions.forEach((ad) => {
              const s = (ad.status || '').toLowerCase();
              const displayStatus = s === 'completed' ? 'Completed' : s === 'in progress' ? 'Pending' : 'Yet to Start';
              const key = (ad.actName || '').trim();
              const keyLower = key.toLowerCase();
              map[key] = displayStatus;
              map[keyLower] = displayStatus;
              if (key !== ad.actName) map[ad.actName] = displayStatus;
            });
            setStatusFromActDescription(map);
          } else {
            setStatusFromActDescription({});
          }
        } catch (descErr) {
          console.error('Fetch act descriptions for status:', descErr);
          setStatusFromActDescription({});
        }
      } else {
        setActsBulkList([]);
        setStatusFromActDescription({});
      }
    } catch (err) {
      setActsBulkList([]);
      setStatusFromActDescription({});
    } finally {
      if (!silentRefresh) setActsBulkLoading(false);
    }
  }, []);

  useEffect(() => {
    if (!showForm) fetchActsBulk({ useCacheFirst: true, silentRefresh: true });
  }, [showForm, fetchActsBulk]);

  useEffect(() => {
    setStatusMap((prev) => {
      const next = { ...prev };
      actsBulkList.forEach((row) => {
        const actName = (row.acts || row.actName || '').trim();
        const actNameLower = actName.toLowerCase();
        next[row.id] = statusFromActDescription[actName] || statusFromActDescription[actNameLower] || row.status || 'Yet to Start';
      });
      return next;
    });
  }, [actsBulkList, statusFromActDescription]);

  useEffect(() => {
    if (openCardMenuId == null) return;
    const close = () => setOpenCardMenuId(null);
    document.addEventListener('click', close);
    return () => document.removeEventListener('click', close);
  }, [openCardMenuId]);

  const getStatus = (id) => statusMap[id] || 'Yet to Start';
  const setStatus = (id, status) => setStatusMap((prev) => ({ ...prev, [id]: status }));

  const updateLibraryStatus = async (id, status, row) => {
    const actRow = row || actsBulkList.find((r) => r.id === id);
    setStatus(id, status);
    if (actRow) {
      try {
        await fetch(`${API_BASE}/act-descriptions/upsert`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            actName: actRow.acts || actRow.actName || '',
            description: actRow.description || null,
            type: actRow.type || null,
            state: actRow.states || actRow.state || null,
            sector: actRow.sector || null,
            applicability: actRow.applicability || null,
            keyComplianceRequirements: actRow.keyComplianceRequirements || null,
            dueDate: actRow.dueDate || null,
            penaltyforNonCompliance: actRow.penaltyforNonCompliance || null,
            registers: actRow.registers || null,
            status,
          }),
        });
      } catch (err) {
        console.error('Failed to save status to ActDescription:', err);
      }
    }
    try {
      const res = await fetch(`${ACTSBULK_API_BASE}/actsbulk?action=updateStatus&id=${encodeURIComponent(id)}`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ status }),
      });
      const data = await res.json();
      if (data.status === 'success') setStatus(id, status);
    } catch (err) {
      console.error('Failed to persist status to Actsbulk:', err);
    }
  };

  const toggleLibrarySelection = (id) => {
    const currentlySelected = selectedLibraryIds.has(id);
    const willBeChecked = !currentlySelected;
    if (willBeChecked && getStatus(id) === 'Pending') {
      const row = actsBulkList.find((r) => r.id === id);
      updateLibraryStatus(id, 'Completed', row);
    }
    setSelectedLibraryIds((prev) => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  };
  const isLibrarySelected = (id) => selectedLibraryIds.has(id);

  const displayActsBulkList = useMemo(() => {
    if (!inchargeDisplayScope.ready) return [];
    if (!hasSiteBasedScope) return actsBulkList;
    const allow = new Set(allowedActCategoryList);
    let rows = actsBulkList.filter((row) => {
      const act = row.acts || row.actName || row.Act || '';
      const sector = row.sector || row.Sector || '';
      return allow.has(getActCategoryFromActSector(act, sector));
    });
    const inds = inchargeDisplayScope.industryLabels;
    if (inds && inds.length > 0) {
      rows = rows.filter((row) => sectorMatchesInchargeSiteIndustries(row.sector || row.Sector || '', inds));
    }
    const sts = inchargeDisplayScope.stateLabels;
    if (sts && sts.length > 0) {
      rows = rows.filter((row) =>
        statesFieldMatchesInchargeSiteStates(row.states || row.state || row.State || '', sts)
      );
    }
    return rows;
  }, [actsBulkList, hasSiteBasedScope, allowedActCategoryList, inchargeDisplayScope]);

  const libraryListPending = actsBulkLoading || !inchargeDisplayScope.ready;

  const completedCount = displayActsBulkList.filter((row) => getStatus(row.id) === 'Completed').length;
  const pendingCount = displayActsBulkList.filter((row) => getStatus(row.id) === 'Pending').length;
  const yetToStartCount = displayActsBulkList.filter((row) => getStatus(row.id) === 'Yet to Start').length;
  const progressPct = displayActsBulkList.length ? Math.round((completedCount / displayActsBulkList.length) * 100) : 0;

  const filteredLibraryList = displayActsBulkList.filter((row) => {
    const s = getStatus(row.id);
    if (libraryFilter === 'all') return true;
    if (libraryFilter === 'completed') return s === 'Completed';
    if (libraryFilter === 'pending') return s === 'Pending';
    if (libraryFilter === 'yet-to-start') return s === 'Yet to Start';
    return true;
  });

  const uniqueSectors = useMemo(() => {
    const set = new Set();
    displayActsBulkList.forEach((r) => {
      const s = (r.sector || '').trim();
      if (s) set.add(s);
    });
    return [...set].sort();
  }, [displayActsBulkList]);

  const uniqueStates = useMemo(() => {
    const set = new Set();
    displayActsBulkList.forEach((r) => {
      String(r.states || r.state || '')
        .split(/[,;/]/)
        .forEach((p) => {
          const t = p.trim();
          if (t) set.add(t);
        });
    });
    return [...set].sort();
  }, [displayActsBulkList]);

  const filteredWithAdvanced = useMemo(() => {
    return filteredLibraryList.filter((row) => {
      const sector = (row.sector || '').trim();
      if (filterSector && sector !== filterSector) return false;
      if (filterState) {
        const blob = String(row.states || row.state || '').toLowerCase();
        const want = filterState.toLowerCase();
        if (!blob.includes(want) && !blob.split(/[,;/]/).some((p) => p.trim().toLowerCase() === want)) return false;
      }
      if (!rowMatchesTypeFilter(row, filterType)) return false;
      return true;
    });
  }, [filteredLibraryList, filterSector, filterState, filterType]);

  const searchFilteredList = useMemo(() => {
    const q = tableSearch.trim().toLowerCase();
    if (!q) return filteredWithAdvanced;
    return filteredWithAdvanced.filter((row) => {
      const act = String(row.acts || row.actName || '').toLowerCase();
      const sector = String(row.sector || '').toLowerCase();
      return act.includes(q) || sector.includes(q);
    });
  }, [filteredWithAdvanced, tableSearch]);

  const sortedLibraryList = useMemo(() => {
    const list = [...searchFilteredList];
    if (!tableSortKey) return list;
    const dir = tableSortDir === 'asc' ? 1 : -1;
    list.sort((a, b) => {
      switch (tableSortKey) {
        case 'acts':
          return dir * String(a.acts || '').localeCompare(String(b.acts || ''), undefined, { sensitivity: 'base' });
        case 'sector':
          return dir * String(a.sector || '').localeCompare(String(b.sector || ''), undefined, { sensitivity: 'base' });
        case 'dueDate': {
          const da = new Date(a.dueDate || 0).getTime();
          const db = new Date(b.dueDate || 0).getTime();
          const na = Number.isNaN(da) ? 0 : da;
          const nb = Number.isNaN(db) ? 0 : db;
          return dir * (na - nb);
        }
        case 'status':
          return dir * String(statusMap[a.id] || 'Yet to Start').localeCompare(String(statusMap[b.id] || 'Yet to Start'));
        default:
          return 0;
      }
    });
    return list;
  }, [searchFilteredList, tableSortKey, tableSortDir, statusMap]);

  const libraryTotalPages = Math.max(1, Math.ceil(sortedLibraryList.length / libraryPageSize));

  const paginatedLibraryList = useMemo(() => {
    const start = (libraryPage - 1) * libraryPageSize;
    return sortedLibraryList.slice(start, start + libraryPageSize);
  }, [sortedLibraryList, libraryPage, libraryPageSize]);

  useEffect(() => {
    setLibraryPage(1);
  }, [libraryFilter, filterSector, filterState, filterType, tableSearch]);

  useEffect(() => {
    if (libraryPage > libraryTotalPages) setLibraryPage(libraryTotalPages);
  }, [libraryPage, libraryTotalPages]);

  const allFilteredIds = useMemo(() => sortedLibraryList.map((r) => r.id), [sortedLibraryList]);
  const allFilteredSelected =
    allFilteredIds.length > 0 && allFilteredIds.every((id) => selectedLibraryIds.has(id));

  const toggleSelectAllFiltered = () => {
    if (allFilteredSelected) {
      setSelectedLibraryIds((prev) => {
        const next = new Set(prev);
        allFilteredIds.forEach((id) => next.delete(id));
        return next;
      });
    } else {
      setSelectedLibraryIds((prev) => {
        const next = new Set(prev);
        allFilteredIds.forEach((id) => next.add(id));
        return next;
      });
    }
  };

  const clearAdvancedFilters = () => {
    setFilterSector('');
    setFilterState('');
    setFilterType('all');
  };

  const hasActiveFilters = Boolean(filterSector || filterState || filterType !== 'all');

  const toggleTableSort = (key) => {
    setTableSortKey((prev) => {
      if (prev === key) {
        setTableSortDir((d) => (d === 'asc' ? 'desc' : 'asc'));
        return prev;
      }
      setTableSortDir('asc');
      return key;
    });
  };

  const openEditFromBulkRow = (row) => {
    const dueRaw = row.dueDate;
    let dueStr = '';
    if (dueRaw) {
      const d = new Date(dueRaw);
      if (!Number.isNaN(d.getTime())) dueStr = d.toISOString().slice(0, 10);
    }
    const st = getStatus(row.id);
    const statusForm = st === 'Completed' ? 'completed' : st === 'Pending' ? 'in progress' : 'yet to start';
    setEditingId(null);
    setForm({
      actName: row.acts || row.actName || '',
      description: row.description || '',
      type: row.type || '',
      state: String(row.states || row.state || '')
        .split(/[,;/]/)[0]
        ?.trim() || '',
      sector: row.sector || '',
      applicability: row.applicability || '',
      keyComplianceRequirements: row.keyComplianceRequirements || '',
      dueDate: dueStr,
      penaltyforNonCompliance: row.penaltyforNonCompliance || '',
      registers: row.registers != null ? String(row.registers) : '',
      status: statusForm,
    });
    setShowForm(true);
  };

  const handleDeleteActsBulkRow = async (id) => {
    if (!window.confirm('Delete this act from the library?')) return;
    setMessage('');
    try {
      const res = await fetch(`${ACTSBULK_API_BASE}/actsbulk?action=delete&id=${encodeURIComponent(id)}`, {
        method: 'DELETE',
        credentials: 'include',
      });
      const data = await res.json();
      if (data.status === 'success') {
        setMessage('Act removed from library.');
        fetchActsBulk();
        setSelectedLibraryIds((prev) => {
          const next = new Set(prev);
          next.delete(id);
          return next;
        });
      } else {
        setMessage(data.message || 'Delete failed.');
      }
    } catch (err) {
      setMessage('Delete failed.');
    }
  };

  const openEdit = (item) => {
    setEditingId(item.id);
    setForm({
      actName: item.actName || '',
      description: item.description || '',
      type: item.type || '',
      state: item.state || '',
      sector: item.sector || '',
      applicability: item.applicability || '',
      keyComplianceRequirements: item.keyComplianceRequirements || '',
      dueDate: item.dueDate || '',
      penaltyforNonCompliance: item.penaltyforNonCompliance || '',
      registers: item.registers || '',
      status: item.status || 'yet to start',
    });
    setShowForm(true);
  };

  const closeForm = () => {
    setShowForm(false);
    setEditingId(null);
    setForm(initialForm);
  };

  const handleChange = (e) => {
    const { name, value } = e.target;
    setForm((prev) => ({ ...prev, [name]: value }));
  };

  const handleSubmit = async (e) => {
    e.preventDefault();
    setMessage('');
    const payload = {
      actName: String(form.actName || '').trim(),
      description: form.description || null,
      type: form.type || null,
      state: form.state || null,
      sector: form.sector || null,
      applicability: form.applicability || null,
      keyComplianceRequirements: form.keyComplianceRequirements || null,
      dueDate: form.dueDate || null,
      penaltyforNonCompliance: form.penaltyforNonCompliance || null,
      registers: form.registers || null,
      status: form.status || 'yet to start',
    };
    if (!payload.actName) {
      setMessage('Act name is required.');
      return;
    }
    try {
      if (editingId) {
        const res = await fetch(`${API_BASE}/act-descriptions/${editingId}`, {
          method: 'PUT',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(payload),
        });
        const data = await res.json();
        if (data.status === 'success') {
          setMessage('Act description updated successfully.');
          closeForm();
          fetchList();
        } else {
          setMessage(data.message || 'Update failed.');
        }
      } else {
        const res = await fetch(`${API_BASE}/act-descriptions`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(payload),
        });
        const data = await res.json();
        if (data.status === 'success') {
          setMessage('Act description added successfully.');
          closeForm();
          fetchList();
        } else {
          setMessage(data.message || 'Add failed.');
        }
      }
    } catch (err) {
      setMessage(err.message || 'Request failed.');
    }
  };

  const handleDelete = async (id) => {
    if (!window.confirm('Delete this act description?')) return;
    setMessage('');
    try {
      const res = await fetch(`${API_BASE}/act-descriptions/${id}`, { method: 'DELETE' });
      const data = await res.json();
      if (data.status === 'success') {
        setMessage('Act description deleted.');
        fetchList();
      } else {
        setMessage(data.message || 'Delete failed.');
      }
    } catch (err) {
      setMessage(err.message || 'Delete failed.');
    }
  };

  const totalPages = Math.max(1, Math.ceil(total / perPage));

  const formModalTitle = editingId ? 'Edit Act Description' : 'Add Act Description';
  const breadcrumbFormCrumb = editingId ? 'Edit' : 'Add';

  return (
    <div
      className={`company-details-page act-description-page${showForm ? ' company-details-page-form-open act-description-page-form-open' : ''}`}
    >
      {!showForm && message ? (
        <div className="company-details-message company-details-message--flush">{message}</div>
      ) : null}
      <header className="company-details-page-heading">
        <h1 className="company-details-page-title">Applicable Acts Library</h1>
        <nav className="company-details-breadcrumb" aria-label="Breadcrumb">
          <Link to="/hcm-dashboard" className="company-details-bc-link">
            Dashboard
          </Link>
          <span className="company-details-bc-sep" aria-hidden>
            ›
          </span>
          <span className="company-details-bc-muted">Library</span>
          <span className="company-details-bc-sep" aria-hidden>
            ›
          </span>
          {showForm ? (
            <button type="button" className="company-details-bc-link company-details-bc-as-link" onClick={closeForm} aria-label="Return to library">
              Act Description
            </button>
          ) : (
            <span className="company-details-bc-current">Act Description</span>
          )}
          {showForm ? (
            <>
              <span className="company-details-bc-sep" aria-hidden>
                ›
              </span>
              <span className="company-details-bc-current">{breadcrumbFormCrumb}</span>
            </>
          ) : null}
        </nav>
      </header>
      {showForm ? (
        <div className="company-details-form-center">
          {message ? <div className="company-details-message company-details-message--modal">{message}</div> : null}
          <div className="company-details-modal" role="dialog" aria-labelledby="act-description-modal-title" aria-modal="true">
            <div className="company-details-modal-header">
              <h2 id="act-description-modal-title" className="company-details-modal-title">
                {formModalTitle}
              </h2>
              <button type="button" className="company-details-modal-close" onClick={closeForm} aria-label="Close">
                ×
              </button>
            </div>
            <form onSubmit={handleSubmit} className="company-details-modal-form" noValidate>
              <div className="company-details-modal-scroll">
                <div className="company-details-modal-body">
                  <section className="company-details-section-card">
                    <header className="company-details-section-head">
                      <h3 className="company-details-section-title">
                        <svg className="company-details-section-title-icon" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" aria-hidden>
                          <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" />
                          <polyline points="14 2 14 8 20 8" />
                          <line x1="16" y1="13" x2="8" y2="13" />
                          <line x1="16" y1="17" x2="8" y2="17" />
                        </svg>
                        Act information
                      </h3>
                    </header>
                    <div className="company-details-form-row">
                      <div className="company-details-form-col company-details-form-col--modal">
                        <div className="company-details-field">
                          <label htmlFor="ad-actName">
                            Act name <span className="required" aria-hidden="true">*</span>
                          </label>
                          <input id="ad-actName" name="actName" value={form.actName} onChange={handleChange} placeholder="Act name" required />
                        </div>
                        <div className="company-details-field">
                          <label htmlFor="ad-description">Description</label>
                          <textarea id="ad-description" name="description" value={form.description} onChange={handleChange} placeholder="Description" rows={3} />
                        </div>
                        <div className="company-details-field">
                          <label htmlFor="ad-type">Type</label>
                          <input id="ad-type" name="type" value={form.type} onChange={handleChange} placeholder="Type" />
                        </div>
                        <div className="company-details-field">
                          <label htmlFor="ad-state">State</label>
                          <input id="ad-state" name="state" value={form.state} onChange={handleChange} placeholder="State" />
                        </div>
                        <div className="company-details-field">
                          <label htmlFor="ad-sector">Sector</label>
                          <input id="ad-sector" name="sector" value={form.sector} onChange={handleChange} placeholder="Sector" />
                        </div>
                      </div>
                      <div className="company-details-form-col company-details-form-col--modal">
                        <div className="company-details-field">
                          <label htmlFor="ad-applicability">Applicability</label>
                          <input id="ad-applicability" name="applicability" value={form.applicability} onChange={handleChange} placeholder="Applicability" />
                        </div>
                        <div className="company-details-field">
                          <label htmlFor="ad-keyComplianceRequirements">Key compliance requirements</label>
                          <textarea id="ad-keyComplianceRequirements" name="keyComplianceRequirements" value={form.keyComplianceRequirements} onChange={handleChange} placeholder="Key compliance requirements" rows={2} />
                        </div>
                        <div className="company-details-field">
                          <label htmlFor="ad-dueDate">Due date</label>
                          <input id="ad-dueDate" name="dueDate" type="date" value={form.dueDate} onChange={handleChange} />
                        </div>
                        <div className="company-details-field">
                          <label htmlFor="ad-penaltyforNonCompliance">Penalty for non-compliance</label>
                          <input id="ad-penaltyforNonCompliance" name="penaltyforNonCompliance" value={form.penaltyforNonCompliance} onChange={handleChange} placeholder="Penalty" />
                        </div>
                        <div className="company-details-field">
                          <label htmlFor="ad-registers">Registers</label>
                          <input id="ad-registers" name="registers" value={form.registers} onChange={handleChange} placeholder="Registers" />
                        </div>
                        <div className="company-details-field">
                          <label htmlFor="ad-status">Status</label>
                          <select id="ad-status" name="status" value={form.status} onChange={handleChange}>
                            {STATUS_OPTIONS.map((s) => (
                              <option key={s} value={s}>{s}</option>
                            ))}
                          </select>
                        </div>
                      </div>
                    </div>
                  </section>
                </div>
              </div>
              <div className="company-details-modal-footer">
                <button type="button" className="company-details-btn company-details-btn-secondary" onClick={closeForm}>
                  Cancel
                </button>
                <button type="submit" className="company-details-btn company-details-btn-primary company-details-btn-submit-modal">
                  {editingId ? 'Update' : 'Submit'}
                </button>
              </div>
            </form>
          </div>
        </div>
      ) : (
        <>
          <div className="company-details-shell-box ad-lib-shell">
            <div className="ad-lib">
              <div className="ad-lib-hero">
                <div className="ad-lib-hero-main">
                  <div className="ad-lib-hero-title-block">
                    <div className="ad-lib-hero-icon" aria-hidden>
                      <FileText size={26} strokeWidth={1.75} />
                    </div>
                    <div>
                      <h2 className="ad-lib-hero-title">Applicable Acts Library</h2>
                      <p className="ad-lib-hero-sub">
                        {!inchargeDisplayScope.ready
                          ? 'Loading your site scope and applicable acts…'
                          : hasSiteBasedScope
                            ? (() => {
                                const parts = [];
                                if (inchargeDisplayScope.stateLabels?.length)
                                  parts.push(`State: ${inchargeDisplayScope.stateLabels.join(', ')}`);
                                if (inchargeDisplayScope.industryLabels?.length)
                                  parts.push(`Industry: ${inchargeDisplayScope.industryLabels.join(', ')}`);
                                parts.push(`Act categories: ${allowedActCategoryList.join(', ')}`);
                                return `Filtered by your Site Management assignment (${parts.join(' · ')}).`;
                              })()
                            : 'Manage and track all imported Acts from bulk import.'}
                      </p>
                    </div>
                  </div>
                  <div className="ad-lib-hero-progress">
                    <div className="ad-lib-progress-ring" style={{ '--progress': progressPct }}>
                      <span className="ad-lib-progress-ring__value">{progressPct}%</span>
                      <span className="ad-lib-progress-ring__label">PROGRESS</span>
                    </div>
                  </div>
                  <div className="ad-lib-tabs" role="tablist" aria-label="Filter by status">
                    <button
                      type="button"
                      role="tab"
                      aria-selected={libraryFilter === 'all'}
                      className={`ad-lib-tab ad-lib-tab--all${libraryFilter === 'all' ? ' ad-lib-tab--active' : ''}`}
                      onClick={() => setLibraryFilter('all')}
                    >
                      All ({displayActsBulkList.length})
                    </button>
                    <button
                      type="button"
                      role="tab"
                      aria-selected={libraryFilter === 'completed'}
                      className={`ad-lib-tab ad-lib-tab--completed${libraryFilter === 'completed' ? ' ad-lib-tab--active' : ''}`}
                      onClick={() => setLibraryFilter('completed')}
                    >
                      <Check size={14} strokeWidth={2.5} aria-hidden />
                      Completed ({completedCount})
                    </button>
                    <button
                      type="button"
                      role="tab"
                      aria-selected={libraryFilter === 'pending'}
                      className={`ad-lib-tab ad-lib-tab--pending${libraryFilter === 'pending' ? ' ad-lib-tab--active' : ''}`}
                      onClick={() => setLibraryFilter('pending')}
                    >
                      <Clock size={14} strokeWidth={2} aria-hidden />
                      Pending ({pendingCount})
                    </button>
                    <button
                      type="button"
                      role="tab"
                      aria-selected={libraryFilter === 'yet-to-start'}
                      className={`ad-lib-tab ad-lib-tab--yet${libraryFilter === 'yet-to-start' ? ' ad-lib-tab--active' : ''}`}
                      onClick={() => setLibraryFilter('yet-to-start')}
                    >
                      <Play size={14} strokeWidth={2} aria-hidden />
                      Yet to Start ({yetToStartCount})
                    </button>
                  </div>
                </div>
              </div>

              {filterPanelOpen ? (
                <div className="ad-lib-filter-panel">
                  <div className="ad-lib-filter-panel__row">
                    <label className="ad-lib-filter-field">
                      <span>Sector</span>
                      <select value={filterSector} onChange={(e) => setFilterSector(e.target.value)}>
                        <option value="">All sectors</option>
                        {uniqueSectors.map((s) => (
                          <option key={s} value={s}>{s}</option>
                        ))}
                      </select>
                    </label>
                    <label className="ad-lib-filter-field">
                      <span>State</span>
                      <select value={filterState} onChange={(e) => setFilterState(e.target.value)}>
                        <option value="">All states</option>
                        {uniqueStates.map((s) => (
                          <option key={s} value={s}>{s}</option>
                        ))}
                      </select>
                    </label>
                    <label className="ad-lib-filter-field">
                      <span>Type</span>
                      <select value={filterType} onChange={(e) => setFilterType(e.target.value)}>
                        <option value="all">All Types</option>
                        <option value="Central">Central</option>
                        <option value="State">State</option>
                      </select>
                    </label>
                  </div>
                </div>
              ) : null}

              <div className={`ad-lib-toolbar${viewMode === 'table' ? ' ad-lib-toolbar--table' : ''}`}>
                <div className="ad-lib-view-toggle">
                  <span className="ad-lib-view-toggle__label">View Mode</span>
                  <div className="ad-lib-view-toggle__btns">
                    <button
                      type="button"
                      className={`ad-lib-view-btn${viewMode === 'card' ? ' ad-lib-view-btn--active' : ''}`}
                      onClick={() => setViewMode('card')}
                    >
                      <LayoutGrid size={16} strokeWidth={2} />
                      Card View
                    </button>
                    <button
                      type="button"
                      className={`ad-lib-view-btn${viewMode === 'table' ? ' ad-lib-view-btn--active' : ''}`}
                      onClick={() => setViewMode('table')}
                    >
                      <Table2 size={16} strokeWidth={2} />
                      Table View
                    </button>
                  </div>
                </div>
                {viewMode === 'table' ? (
                  <div className="ad-lib-table-search-wrap">
                    <Search size={18} className="ad-lib-table-search-icon" strokeWidth={2} aria-hidden />
                    <input
                      type="search"
                      className="ad-lib-table-search-input"
                      placeholder="Search by Act name or sector..."
                      value={tableSearch}
                      onChange={(e) => setTableSearch(e.target.value)}
                      aria-label="Search by act name or sector"
                    />
                  </div>
                ) : (
                  <div className="ad-lib-toolbar-spacer" aria-hidden />
                )}
                <div className="ad-lib-bulk">
                  <button
                    type="button"
                    className="company-details-btn-export"
                    onClick={() => {
                      // Export CSV logic for acts bulk
                      const cols = [
                        { key: 'acts', label: 'Act Name' },
                        { key: 'sector', label: 'Sector' },
                        { key: 'type', label: 'Type' },
                        { key: 'states', label: 'States' },
                        { key: 'description', label: 'Description' },
                        { key: 'applicability', label: 'Applicability' },
                        { key: 'keyComplianceRequirements', label: 'Key Compliance' },
                        { key: 'dueDate', label: 'Due Date' },
                        { key: 'penaltyforNonCompliance', label: 'Penalty for Non-Compliance' },
                        { key: 'registers', label: 'Registers' },
                        { key: 'status', label: 'Status' },
                      ];
                      const esc = (v) => {
                        const s = String(v ?? '').replace(/"/g, '""');
                        return `"${s}"`;
                      };
                      const header = cols.map((c) => esc(c.label)).join(',');
                      const rows = displayActsBulkList.map((row) =>
                        cols.map((c) => esc(row[c.key])).join(',')
                      );
                      const csv = [header, ...rows].join('\r\n');
                      const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
                      const url = URL.createObjectURL(blob);
                      const a = document.createElement('a');
                      a.href = url;
                      a.download = `acts-library-${new Date().toISOString().slice(0, 10)}.csv`;
                      a.click();
                      URL.revokeObjectURL(url);
                    }}
                    style={{ display: 'inline-flex', alignItems: 'center', gap: 6, fontWeight: 600, fontSize: 15 }}
                  >
                    <Upload size={18} style={{ marginRight: 2 }} /> Export CSV
                  </button>
                </div>
              </div>

              <div className="ad-lib-main">
                {libraryListPending ? (
                  <p className="ad-lib-loading">Loading applicable acts...</p>
                ) : sortedLibraryList.length === 0 ? (
                  <p className="ad-lib-empty">No acts match the current filters.</p>
                ) : viewMode === 'card' ? (
                  <div className="ad-lib-grid">
                    {paginatedLibraryList.map((row) => {
                      const status = getStatus(row.id);
                      const sk = status.replace(/\s+/g, '-').toLowerCase();
                      return (
                        <div key={row.id} className={`ad-lib-card ad-lib-card--${sk}`}>
                          <div className="ad-lib-card__accent" aria-hidden />
                          <div className="ad-lib-card__inner">
                            <div className="ad-lib-card__top">
                              <div className="ad-lib-card__doc-icon">
                                <FileText size={22} strokeWidth={1.75} />
                              </div>
                              <div className="ad-lib-card__top-right">
                                <span className={`ad-lib-card__badge ad-lib-card__badge--${sk}`}>{status}</span>
                                <div className="ad-lib-card__menu-wrap">
                                  <button
                                    type="button"
                                    className="ad-lib-card__kebab"
                                    aria-label="More actions"
                                    onClick={(e) => {
                                      e.stopPropagation();
                                      setOpenCardMenuId(openCardMenuId === row.id ? null : row.id);
                                    }}
                                  >
                                    <MoreVertical size={18} strokeWidth={2} />
                                  </button>
                                  {openCardMenuId === row.id ? (
                                    <div className="ad-lib-card__dropdown" onClick={(e) => e.stopPropagation()}>
                                      <button
                                        type="button"
                                        onClick={() => {
                                          updateLibraryStatus(row.id, 'Completed', row);
                                          setOpenCardMenuId(null);
                                        }}
                                      >
                                        Mark Completed
                                      </button>
                                      <button
                                        type="button"
                                        onClick={() => {
                                          updateLibraryStatus(row.id, 'Pending', row);
                                          setOpenCardMenuId(null);
                                        }}
                                      >
                                        Mark Pending
                                      </button>
                                      <button
                                        type="button"
                                        onClick={() => {
                                          updateLibraryStatus(row.id, 'Yet to Start', row);
                                          setOpenCardMenuId(null);
                                        }}
                                      >
                                        Mark Yet to Start
                                      </button>
                                      <button
                                        type="button"
                                        className="ad-lib-card__dropdown-divider"
                                        onClick={() => {
                                          setActDetailsModalAct(row);
                                          setOpenCardMenuId(null);
                                        }}
                                      >
                                        View details
                                      </button>
                                    </div>
                                  ) : null}
                                </div>
                              </div>
                            </div>
                            <button
                              type="button"
                              className="ad-lib-card__title"
                              onClick={() => {
                                setActDetailsModalAct(row);
                                if (getStatus(row.id) === 'Yet to Start') updateLibraryStatus(row.id, 'Pending', row);
                              }}
                            >
                              {(row.acts || '—').toUpperCase()}
                            </button>
                            <ul className="ad-lib-card__meta">
                              <li>
                                <Briefcase size={14} strokeWidth={2} aria-hidden />
                                <span>{(row.sector || '—').trim() || '—'}</span>
                              </li>
                              <li>
                                <Scale size={14} strokeWidth={2} aria-hidden />
                                <span>{normalizeActType(row) || '—'}</span>
                              </li>
                              <li>
                                <MapPin size={14} strokeWidth={2} aria-hidden />
                                <span>{String(row.states || row.state || '—').trim() || '—'}</span>
                              </li>
                            </ul>
                            <div className="ad-lib-card__foot">
                              <button
                                type="button"
                                className="ad-lib-card__view"
                                onClick={() => setActDetailsModalAct(row)}
                              >
                                View Details
                                <ChevronRight size={14} strokeWidth={2.5} />
                              </button>
                            </div>
                          </div>
                        </div>
                      );
                    })}
                  </div>
                ) : (
                  <div className="ad-lib-table-outer">
                    <div className="ad-lib-table-scroll">
                      <table className="ad-lib-table ad-lib-table--wide">
                        <thead>
                          <tr>
                            <th className="ad-lib-th ad-lib-th--check">
                              <span className="ad-lib-sr-only">Select</span>
                            </th>
                            <th className="ad-lib-th ad-lib-th--narrow">S.No</th>
                            <th className="ad-lib-th">
                              <button type="button" className="ad-lib-th-sort" onClick={() => toggleTableSort('sector')}>
                                Sector
                                <ArrowUpDown size={12} strokeWidth={2} aria-hidden />
                              </button>
                            </th>
                            <th className="ad-lib-th ad-lib-th--acts">
                              <button type="button" className="ad-lib-th-sort" onClick={() => toggleTableSort('acts')}>
                                Acts
                                <ArrowUpDown size={12} strokeWidth={2} aria-hidden />
                              </button>
                            </th>
                            <th className="ad-lib-th">Type</th>
                            <th className="ad-lib-th">States</th>
                            <th className="ad-lib-th">Description</th>
                            <th className="ad-lib-th ad-lib-th--col-applicability">Applicability</th>
                            <th className="ad-lib-th ad-lib-th--col-key-compliance">Key Compliance</th>
                            <th className="ad-lib-th">
                              <button type="button" className="ad-lib-th-sort" onClick={() => toggleTableSort('dueDate')}>
                                Due Date
                                <ArrowUpDown size={12} strokeWidth={2} aria-hidden />
                              </button>
                            </th>
                            <th className="ad-lib-th">Penalty for Non-Compliance</th>
                            <th className="ad-lib-th ad-lib-th--center">Registers</th>
                            <th className="ad-lib-th">
                              <button type="button" className="ad-lib-th-sort" onClick={() => toggleTableSort('status')}>
                                Status
                                <ArrowUpDown size={12} strokeWidth={2} aria-hidden />
                              </button>
                            </th>
                            <th className="ad-lib-th ad-lib-th--actions">Actions</th>
                          </tr>
                        </thead>
                        <tbody>
                          {paginatedLibraryList.map((row, idx) => {
                            const status = getStatus(row.id);
                            const sk = status.replace(/\s+/g, '-').toLowerCase();
                            const serialNo = (libraryPage - 1) * libraryPageSize + idx + 1;
                            const typeLabel = normalizeActType(row) || '—';
                            const typeIsCentral = typeLabel === 'Central';
                            const urgency = dueDateUrgency(row);
                            const regN = registerCountCell(row.registers);
                            const sector = (row.sector || '—').trim() || '—';
                            const tone = sectorPillTone(row.sector);
                            return (
                              <tr key={row.id} className="ad-lib-tr">
                                <td className="ad-lib-td ad-lib-td--check">
                                  <input
                                    type="checkbox"
                                    checked={isLibrarySelected(row.id)}
                                    onChange={() => toggleLibrarySelection(row.id)}
                                    aria-label={`Select ${row.acts || 'act'}`}
                                  />
                                </td>
                                <td className="ad-lib-td ad-lib-td--mono">{serialNo}</td>
                                <td className="ad-lib-td">
                                  <span className={`ad-lib-pill-sector ad-lib-pill-sector--${tone}`}>
                                    <Briefcase size={12} strokeWidth={2} aria-hidden />
                                    {sector}
                                  </span>
                                </td>
                                <td className="ad-lib-td ad-lib-td--act" title={row.acts || ''}>
                                  {row.acts || '—'}
                                </td>
                                <td className="ad-lib-td">
                                  {typeLabel !== '—' ? (
                                    <span className={`ad-lib-pill-type ${typeIsCentral ? 'ad-lib-pill-type--central' : 'ad-lib-pill-type--state'}`}>
                                      {typeLabel}
                                    </span>
                                  ) : (
                                    '—'
                                  )}
                                </td>
                                <td className="ad-lib-td">
                                  <span className="ad-lib-pill-state">{abbreviateStateLabel(row.states || row.state)}</span>
                                </td>
                                <td className="ad-lib-td ad-lib-td--clip" title={row.description || ''}>
                                  {truncateTableText(row.description, 100)}
                                </td>
                                <td className="ad-lib-td ad-lib-td--applicability" title={row.applicability || ''}>
                                  {truncateTableText(row.applicability, 220)}
                                </td>
                                <td className="ad-lib-td ad-lib-td--key-compliance" title={row.keyComplianceRequirements || ''}>
                                  {truncateTableText(row.keyComplianceRequirements, 220)}
                                </td>
                                <td className="ad-lib-td">
                                  <span className={`ad-lib-due ad-lib-due--${urgency}`}>
                                    <Calendar size={13} strokeWidth={2} aria-hidden />
                                    {formatDueDateAct(row)}
                                  </span>
                                </td>
                                <td className="ad-lib-td ad-lib-td--clip" title={row.penaltyforNonCompliance || ''}>
                                  {truncateTableText(row.penaltyforNonCompliance, 90)}
                                </td>
                                <td className="ad-lib-td ad-lib-td--center">
                                  {regN != null ? (
                                    <span className="ad-lib-registers-badge" title={String(row.registers ?? '')}>
                                      <FileStack size={14} strokeWidth={2} aria-hidden />
                                      {regN}
                                    </span>
                                  ) : (
                                    '—'
                                  )}
                                </td>
                                <td className="ad-lib-td ad-lib-td--status">
                                  <span className={`ad-lib-status-pill ad-lib-status-pill--${sk}`}>
                                    {status === 'Completed' && <Check size={12} strokeWidth={2.5} aria-hidden />}
                                    {status === 'Pending' && <Clock size={12} strokeWidth={2} aria-hidden />}
                                    {status === 'Yet to Start' && <Play size={12} strokeWidth={2} aria-hidden />}
                                    {status}
                                  </span>
                                </td>
                                <td className="ad-lib-td ad-lib-td--actions">
                                  <button
                                    type="button"
                                    className="ad-lib-icon-btn ad-lib-icon-btn--edit"
                                    title="Edit"
                                    aria-label="Edit"
                                    onClick={() => openEditFromBulkRow(row)}
                                  >
                                    <Pencil size={15} strokeWidth={2} />
                                  </button>
                                  <button
                                    type="button"
                                    className="ad-lib-icon-btn ad-lib-icon-btn--delete"
                                    title="Delete"
                                    aria-label="Delete"
                                    onClick={() => handleDeleteActsBulkRow(row.id)}
                                  >
                                    <Trash2 size={15} strokeWidth={2} />
                                  </button>
                                </td>
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    </div>
                  </div>
                )}
              </div>

              {!libraryListPending && sortedLibraryList.length > 0 ? (
                <div className="ad-lib-pagination ad-lib-pagination--company-model">
                  <p className="ad-lib-pagination__info">
                    Showing{' '}
                    {sortedLibraryList.length === 0
                      ? 0
                      : (libraryPage - 1) * libraryPageSize + 1}{' '}
                    to {Math.min(libraryPage * libraryPageSize, sortedLibraryList.length)} of {sortedLibraryList.length} Acts.
                  </p>
                  <nav className="company-details-pagination" aria-label="Acts library pagination">
                    <button
                      type="button"
                      className="company-details-pagination-nav"
                      disabled={libraryPage <= 1}
                      onClick={() => setLibraryPage(1)}
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
                      className="company-details-pagination-nav"
                      disabled={libraryPage <= 1}
                      onClick={() => setLibraryPage((p) => Math.max(1, p - 1))}
                      title="Previous page"
                      aria-label="Previous page"
                    >
                      <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden>
                        <polyline points="15 18 9 12 15 6" />
                      </svg>
                    </button>
                    <div className="company-details-pagination-pages">
                      {buildPaginationItems(libraryPage, libraryTotalPages).map((item, idx) =>
                      item === 'ellipsis' ? (
                        <span key={`e-${idx}`} className="company-details-pagination-ellipsis" aria-hidden>
                          …
                        </span>
                      ) : (
                        <button
                          key={item}
                          type="button"
                          className={`company-details-pagination-page${libraryPage === item ? ' company-details-pagination-page--active' : ''}`}
                          onClick={() => setLibraryPage(item)}
                          aria-label={`Page ${item}`}
                          aria-current={libraryPage === item ? 'page' : undefined}
                        >
                          {item}
                        </button>
                      )
                      )}
                    </div>
                    <button
                      type="button"
                      className="company-details-pagination-nav"
                      disabled={libraryPage >= libraryTotalPages}
                      onClick={() => setLibraryPage((p) => Math.min(libraryTotalPages, p + 1))}
                      title="Next page"
                      aria-label="Next page"
                    >
                      <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden>
                        <polyline points="9 18 15 12 9 6" />
                      </svg>
                    </button>
                    <button
                      type="button"
                      className="company-details-pagination-nav"
                      disabled={libraryPage >= libraryTotalPages}
                      onClick={() => setLibraryPage(libraryTotalPages)}
                      title="Last page"
                      aria-label="Last page"
                    >
                      <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden>
                        <polyline points="13 17 18 12 13 7" />
                        <polyline points="6 17 11 12 6 7" />
                      </svg>
                    </button>
                  </nav>
                  <label className="ad-lib-rows">
                    Rows per page
                    <select
                      value={libraryPageSize}
                      onChange={(e) => {
                        setLibraryPageSize(Number(e.target.value));
                        setLibraryPage(1);
                      }}
                    >
                      <option value={8}>8</option>
                      <option value={12}>12</option>
                      <option value={24}>24</option>
                    </select>
                  </label>
                </div>
              ) : null}
            </div>
          </div>
        </>
      )}

      {/* Act Details modal - card list style */}
      {actDetailsModalAct && (
        <div className="act-description-modal-overlay" onClick={() => setActDetailsModalAct(null)}>
          <div className="act-description-modal" onClick={(e) => e.stopPropagation()}>
            <div className="act-description-modal-header">
              <h2 className="act-description-modal-title">Act Details</h2>
              <button type="button" className="act-description-modal-close" onClick={() => setActDetailsModalAct(null)} aria-label="Close">×</button>
            </div>
            <div className="act-description-modal-body">
              <div className="act-details-layout">
                <div className="act-details-cover">
                  {(() => {
                    const stateLabel = String(actDetailsModalAct.states || actDetailsModalAct.state || '').trim();
                    return stateLabel ? (
                      <p className="act-details-cover__state">{stateLabel}</p>
                    ) : null;
                  })()}
                  <h3 className="act-details-cover__title">{actDetailsModalAct.acts || '—'}</h3>
                  {String(actDetailsModalAct.description || '').trim() ? (
                    <p className="act-details-cover__desc">{actDetailsModalAct.description.trim()}</p>
                  ) : null}
                </div>
                <div className="act-details-list">
                  <div className="act-details-item act-details-item--violet">
                    <div className="act-details-item__icon" aria-hidden>
                      <FileText size={18} strokeWidth={2} />
                    </div>
                    <div className="act-details-item__content">
                      <h4>Registers</h4>
                      <p>{actDetailsModalAct.registers || '—'}</p>
                    </div>
                    <span className="act-details-item__no">01</span>
                  </div>
                  <div className="act-details-item act-details-item--red">
                    <div className="act-details-item__icon" aria-hidden>
                      <Scale size={18} strokeWidth={2} />
                    </div>
                    <div className="act-details-item__content">
                      <h4>Penalty For Non Compliance</h4>
                      <p>{actDetailsModalAct.penaltyforNonCompliance || '—'}</p>
                    </div>
                    <span className="act-details-item__no">02</span>
                  </div>
                  <div className="act-details-item act-details-item--blue">
                    <div className="act-details-item__icon" aria-hidden>
                      <Briefcase size={18} strokeWidth={2} />
                    </div>
                    <div className="act-details-item__content">
                      <h4>Applicability</h4>
                      <p>{actDetailsModalAct.applicability || '—'}</p>
                    </div>
                    <span className="act-details-item__no">03</span>
                  </div>
                  <div className="act-details-item act-details-item--green">
                    <div className="act-details-item__icon" aria-hidden>
                      <Check size={18} strokeWidth={2.4} />
                    </div>
                    <div className="act-details-item__content">
                      <h4>Key Compliance Requirements</h4>
                      <p>{actDetailsModalAct.keyComplianceRequirements || '—'}</p>
                    </div>
                    <span className="act-details-item__no">04</span>
                  </div>
                  <div className="act-details-item act-details-item--orange">
                    <div className="act-details-item__icon" aria-hidden>
                      <Calendar size={18} strokeWidth={2} />
                    </div>
                    <div className="act-details-item__content">
                      <h4>Due Date</h4>
                      <p>{actDetailsModalAct.dueDate || '—'}</p>
                    </div>
                    <span className="act-details-item__no">05</span>
                  </div>
                </div>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

export default ActDescription;
