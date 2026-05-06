import React, { useState, useEffect } from 'react';
import { Search, FileSpreadsheet, Download, Database } from 'lucide-react';
import './Formmaster.css';

const FORMMASTER_API = '/server/formmaster_function/templates';
const FORMMASTER_SYNC_API = '/server/formmaster_function/templates/sync';
const DOWNLOAD_BASE = '/server/formmaster_function/templates/download';

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

const Formmaster = () => {
  const [formName, setFormName] = useState('');
  const [templates, setTemplates] = useState([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const [searched, setSearched] = useState(false);
  const [syncing, setSyncing] = useState(false);
  const [syncMessage, setSyncMessage] = useState('');
  const [currentPage, setCurrentPage] = useState(1);
  const itemsPerPage = 10;

  const loadTemplates = async (filter) => {
    setError('');
    setSearched(true);
    setCurrentPage(1);
    setLoading(true);
    try {
      const url = (filter && String(filter).trim())
        ? `${FORMMASTER_API}?formName=${encodeURIComponent(String(filter).trim())}`
        : FORMMASTER_API;
      const res = await fetch(url, { credentials: 'include' });
      const text = await res.text();
      let data;
      try {
        data = text ? JSON.parse(text) : {};
      } catch (_) {
        setTemplates([]);
        setError(res.ok ? 'Invalid response from server.' : (text || `Request failed (${res.status}).`));
        return;
      }
      if (!res.ok) {
        setTemplates([]);
        setError(data.message || text || 'Failed to load templates.');
        return;
      }
      const list = (data.data && Array.isArray(data.data)) ? data.data : [];
      setTemplates(list);
      if (list.length === 0 && filter && String(filter).trim()) {
        setError('No Excel templates found matching your search.');
      }
    } catch (err) {
      console.error('Formmaster fetch error:', err);
      setTemplates([]);
      setError(err.message || 'Failed to load templates.');
    } finally {
      setLoading(false);
    }
  };

  // Load all forms on page open so the full list is displayed immediately
  useEffect(() => {
    loadTemplates('');
  }, []);

  const handleSearch = () => {
    loadTemplates(formName);
  };

  const handleKeyDown = (e) => {
    if (e.key === 'Enter') handleSearch();
  };

  const handleSyncToDataStore = async () => {
    setSyncMessage('');
    setError('');
    setSyncing(true);
    try {
      const res = await fetch(FORMMASTER_SYNC_API, { method: 'POST', credentials: 'include' });
      const text = await res.text();
      let data;
      try {
        data = text ? JSON.parse(text) : {};
      } catch (_) {
        setSyncMessage(res.ok ? 'Invalid response.' : (text || `Request failed (${res.status}).`));
        return;
      }
      if (res.ok && data.status === 'success') {
        setSyncMessage(data.message || 'FormMaster table updated.');
      } else {
        setSyncMessage(data.message || text || 'Sync failed.');
      }
    } catch (err) {
      console.error('Sync error:', err);
      setSyncMessage(err.message || 'Sync failed.');
    } finally {
      setSyncing(false);
    }
  };

  const formatSize = (val) => {
    if (val == null || val === '') return '—';
    if (typeof val === 'string') return val;
    const bytes = Number(val);
    if (Number.isNaN(bytes)) return '—';
    if (bytes < 1024) return `${bytes} B`;
    if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
    return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
  };

  const formatTime = (val) => {
    if (val == null || val === '') return '—';
    if (typeof val === 'object') return '—';
    try {
      const d = new Date(val);
      return isNaN(d.getTime()) ? String(val) : d.toLocaleString();
    } catch (_) {
      return '—';
    }
  };

  // Ensure we never render an object (React error #31); Catalyst may return nested objects
  const toText = (val) => {
    if (val == null || val === '') return '—';
    if (typeof val === 'string') return val;
    if (typeof val === 'number' || typeof val === 'boolean') return String(val);
    if (typeof val === 'object' && val !== null && ('email_id' in val || 'email' in val)) return val.email_id || val.email || '—';
    return '—';
  };

  const downloadUrl = (fileId, fileName) => {
    const name = (fileName || 'template.xlsx').replace(/[^a-zA-Z0-9._-]/g, '_');
    return `${DOWNLOAD_BASE}/${fileId}?fileName=${encodeURIComponent(name)}`;
  };

  const totalPages = Math.max(1, Math.ceil(templates.length / itemsPerPage));
  const paginatedTemplates = templates.slice(
    (currentPage - 1) * itemsPerPage,
    currentPage * itemsPerPage
  );
  const paginationItems = buildPaginationItems(currentPage, totalPages);

  return (
    <div className="formmaster-container">
      <div className="formmaster-header">
        <div className="formmaster-header-content">
          <div className="formmaster-header-icon">
            <FileSpreadsheet size={28} />
          </div>
          <div className="formmaster-header-text">
            <h1 className="formmaster-title">Form Master</h1>
            <p className="formmaster-subtitle">Search and view backend Excel templates</p>
          </div>
        </div>
      </div>

      <div className="formmaster-content">
        <div className="formmaster-search-card">
          <label htmlFor="formmaster-form-name" className="formmaster-label">Form Name</label>
          <div className="formmaster-search-row">
            <input
              id="formmaster-form-name"
              type="text"
              className="formmaster-input"
              placeholder="e.g. Form 10, Form U, or leave empty for all"
              value={formName}
              onChange={(e) => setFormName(e.target.value)}
              onKeyDown={handleKeyDown}
            />
            <button
              type="button"
              className="formmaster-btn-search"
              onClick={handleSearch}
              disabled={loading}
            >
              <Search size={18} />
              {loading ? ' Searching...' : ' Search'}
            </button>
            <button
              type="button"
              className="formmaster-btn-sync"
              onClick={handleSyncToDataStore}
              disabled={syncing}
              title="Copy templates from File Store into FormMaster Data Store table"
            >
              <Database size={18} />
              {syncing ? ' Syncing...' : ' Sync to Data Store'}
            </button>
          </div>
          {syncMessage && (
            <p className="formmaster-sync-message">{syncMessage}</p>
          )}
        </div>

        {error && (
          <div className="formmaster-error" role="alert">
            {error}
          </div>
        )}

        {searched && !loading && (
          <div className="formmaster-table-card">
            <h2 className="formmaster-table-title">
              Templates {formName.trim() ? `matching "${formName}"` : '(all)'}
            </h2>
            {templates.length === 0 ? (
              <p className="formmaster-empty">No templates to display.</p>
            ) : (
              <>
                <div className="formmaster-table-wrapper">
                  <table className="formmaster-data-table">
                    <thead>
                      <tr>
                        <th>File Name</th>
                        <th>Created By</th>
                        <th>Created Time</th>
                        <th className="formmaster-th-actions">Action</th>
                      </tr>
                    </thead>
                    <tbody>
                      {paginatedTemplates.map((row) => (
                          <tr key={row.id || row.name} className="formmaster-data-row">
                            <td>
                              <span className="formmaster-file-name">
                                <FileSpreadsheet size={16} />
                                {toText(row.name)}
                              </span>
                            </td>
                            <td>{toText(row.createdBy)}</td>
                            <td>{formatTime(row.createdTime)}</td>
                            <td>
                              <div className="formmaster-table-actions">
                                <a
                                  href={downloadUrl(row.id, row.name)}
                                  className="formmaster-download-link"
                                  download
                                >
                                  <Download size={16} />
                                  Download
                                </a>
                              </div>
                            </td>
                          </tr>
                        ))}
                    </tbody>
                  </table>
                </div>
                <div className="formmaster-pagination">
                  <nav className="formmaster-pagination-nav-wrap" aria-label="Table pagination">
                    <button
                      type="button"
                      className="formmaster-pagination-nav"
                      disabled={currentPage <= 1}
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
                      className="formmaster-pagination-nav"
                      onClick={() => setCurrentPage((prev) => Math.max(1, prev - 1))}
                      disabled={currentPage === 1}
                      title="Previous page"
                      aria-label="Previous page"
                    >
                      <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden>
                        <polyline points="15 18 9 12 15 6" />
                      </svg>
                    </button>
                    <div className="formmaster-pagination-pages">
                      {paginationItems.map((item, i) =>
                        item === 'ellipsis' ? (
                          <span key={`e-${i}`} className="formmaster-pagination-ellipsis" aria-hidden>
                            ...
                          </span>
                        ) : (
                          <button
                            key={item}
                            type="button"
                            className={`formmaster-pagination-page${item === currentPage ? ' formmaster-pagination-page--active' : ''}`}
                            onClick={() => setCurrentPage(item)}
                            aria-label={`Page ${item}`}
                            aria-current={item === currentPage ? 'page' : undefined}
                          >
                            {item}
                          </button>
                        )
                      )}
                    </div>
                    <button
                      type="button"
                      className="formmaster-pagination-nav"
                      onClick={() => setCurrentPage((prev) => Math.min(totalPages, prev + 1))}
                      disabled={currentPage >= totalPages}
                      title="Next page"
                      aria-label="Next page"
                    >
                      <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden>
                        <polyline points="9 18 15 12 9 6" />
                      </svg>
                    </button>
                    <button
                      type="button"
                      className="formmaster-pagination-nav"
                      onClick={() => setCurrentPage(totalPages)}
                      disabled={currentPage >= totalPages}
                      title="Last page"
                      aria-label="Last page"
                    >
                      <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden>
                        <polyline points="13 17 18 12 13 7" />
                        <polyline points="6 17 11 12 6 7" />
                      </svg>
                    </button>
                  </nav>
                </div>
              </>
            )}
          </div>
        )}
      </div>
    </div>
  );
};

export default Formmaster;
