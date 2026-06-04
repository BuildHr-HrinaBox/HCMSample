import React, { useState, useEffect, useCallback, useMemo, useRef } from 'react';
import { Search, FileSpreadsheet, Download, Upload, Database } from 'lucide-react';
import * as XLSX from 'xlsx';
import './Formmaster.css';

const RECORDS_API = '/server/formmaster_function/records';
const FORMMASTER_SYNC_API = '/server/formmaster_function/templates/sync';
const DOWNLOAD_BASE = '/server/formmaster_function/templates/download';
const PAGE_SIZE = 10;

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

const normalizeHeaderKey = (key) =>
  String(key ?? '')
    .replace(/\u00a0/g, ' ')
    .replace(/\s+/g, ' ')
    .replace(/\.+$/g, '')
    .trim()
    .toLowerCase();

const buildNormRow = (row) => {
  const norm = {};
  for (const [k, v] of Object.entries(row)) {
    if (k === '__rowNum__') continue;
    norm[normalizeHeaderKey(k)] = v;
  }
  return norm;
};

const cellByAliases = (normRow, aliases) => {
  for (const a of aliases) {
    const nk = normalizeHeaderKey(a);
    if (!Object.prototype.hasOwnProperty.call(normRow, nk)) continue;
    const v = normRow[nk];
    if (v === undefined || v === null) continue;
    if (String(v).trim() === '') continue;
    return v;
  }
  return undefined;
};

const mapRowToRecord = (row) => {
  const norm = buildNormRow(row);
  const coalesce = (...vals) => {
    for (const v of vals) {
      if (v === undefined || v === null) continue;
      if (String(v).trim() !== '') return String(v).trim();
    }
    return '';
  };
  return {
    act: coalesce(cellByAliases(norm, ['Act', 'Acts', 'act']), row.Act, row.Acts),
    description: coalesce(cellByAliases(norm, ['Description', 'description']), row.Description),
    sector: coalesce(cellByAliases(norm, ['Sector', 'sector']), row.Sector),
    state: coalesce(cellByAliases(norm, ['State', 'States', 'state', 'states']), row.State, row.States),
    formName: coalesce(
      cellByAliases(norm, ['Form Name', 'FormName', 'Form', 'form name', 'formname']),
      row['Form Name'],
      row.FormName
    )
  };
};

const Formmaster = () => {
  const [records, setRecords] = useState([]);
  const [loading, setLoading] = useState(false);
  const [message, setMessage] = useState('');
  const [searchTerm, setSearchTerm] = useState('');
  const [tablePage, setTablePage] = useState(1);
  const [importFile, setImportFile] = useState(null);
  const [dragOver, setDragOver] = useState(false);
  const [syncing, setSyncing] = useState(false);
  const [syncMessage, setSyncMessage] = useState('');
  const [uploadingRowId, setUploadingRowId] = useState(null);
  const rowFileInputRefs = useRef({});

  const fetchRecords = useCallback(async () => {
    setLoading(true);
    setMessage('');
    try {
      const res = await fetch(`${RECORDS_API}?action=getAll&_ts=${Date.now()}`, { credentials: 'include', cache: 'no-store' });
      const data = await res.json();
      if (data.status === 'success' && Array.isArray(data.data)) {
        setRecords(data.data);
      } else {
        setRecords([]);
        setMessage(data.message || 'Failed to load Form Master records.');
      }
    } catch (err) {
      setRecords([]);
      setMessage(err.message || 'Failed to load records.');
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    fetchRecords();
  }, [fetchRecords]);

  useEffect(() => {
    setTablePage(1);
  }, [searchTerm]);

  const handleImportFile = (selectedFile) => {
    if (!selectedFile) return;
    const name = (selectedFile.name || '').toLowerCase();
    if (!name.endsWith('.xlsx') && !name.endsWith('.xls')) {
      setMessage('Supports .xlsx, .xls files only');
      return;
    }
    setImportFile(selectedFile);
    setMessage('');
  };

  const handleDrop = (e) => {
    e.preventDefault();
    setDragOver(false);
    handleImportFile(e.dataTransfer?.files?.[0]);
  };

  const handleBulkImport = async () => {
    if (!importFile) {
      setMessage('Please select an Excel file first');
      return;
    }
    setLoading(true);
    setMessage('');
    try {
      const arrayBuffer = await importFile.arrayBuffer();
      const wb = XLSX.read(arrayBuffer, { type: 'array' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws);
      if (!rows.length) {
        setMessage('No rows found in the Excel file');
        return;
      }
      const recordsPayload = rows.map(mapRowToRecord).filter(
        (r) => r.act || r.description || r.sector || r.state || r.formName
      );
      if (!recordsPayload.length) {
        setMessage('No valid rows. Excel must include Act, Description, Sector, or State columns.');
        return;
      }
      const res = await fetch(`${RECORDS_API}?action=bulkImport`, {
        method: 'POST',
        credentials: 'include',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(recordsPayload)
      });
      const data = await res.json();
      if (data.status === 'success') {
        setMessage(`Imported ${recordsPayload.length} row(s). Upload a template file on each row as needed.`);
        setImportFile(null);
        await fetchRecords();
      } else {
        setMessage(data.message || data.error || 'Import failed');
      }
    } catch (err) {
      setMessage('Import failed: ' + (err.message || 'Unknown error'));
    } finally {
      setLoading(false);
    }
  };

  const handleClearAll = async () => {
    if (!window.confirm('Remove all Form Master records from the database?')) return;
    setLoading(true);
    setMessage('');
    try {
      const res = await fetch(`${RECORDS_API}?action=deleteAll`, {
        method: 'DELETE',
        credentials: 'include'
      });
      const data = await res.json();
      if (data.status === 'success') {
        setMessage('All records cleared');
        setRecords([]);
      } else {
        setMessage(data.message || 'Clear failed');
      }
    } catch (err) {
      setMessage('Clear failed: ' + (err.message || 'Unknown error'));
    } finally {
      setLoading(false);
    }
  };

  const handleSyncToDataStore = async () => {
    setSyncMessage('');
    setSyncing(true);
    try {
      const res = await fetch(FORMMASTER_SYNC_API, { method: 'POST', credentials: 'include' });
      const data = await res.json();
      if (res.ok && data.status === 'success') {
        setSyncMessage(data.message || 'Templates synced from File Store.');
        await fetchRecords();
      } else {
        setSyncMessage(data.message || 'Sync failed.');
      }
    } catch (err) {
      setSyncMessage(err.message || 'Sync failed.');
    } finally {
      setSyncing(false);
    }
  };

  const downloadUrl = (fileId, fileName) => {
    const name = (fileName || 'template.xlsx').replace(/[^a-zA-Z0-9._-]/g, '_');
    return `${DOWNLOAD_BASE}/${fileId}?fileName=${encodeURIComponent(name)}`;
  };

  const triggerRowUpload = (rowId) => {
    const input = rowFileInputRefs.current[rowId];
    if (input) input.click();
  };

  const handleRowFileChange = async (row, file) => {
    if (!file || !row?.id) return;
    const name = (file.name || '').toLowerCase();
    if (!name.endsWith('.xlsx') && !name.endsWith('.xls')) {
      setMessage('Template must be .xlsx or .xls');
      return;
    }
    setUploadingRowId(String(row.id));
    setMessage('');
    try {
      const fd = new FormData();
      fd.append('file', file);
      const res = await fetch(`${RECORDS_API}/${row.id}/upload`, {
        method: 'POST',
        credentials: 'include',
        body: fd
      });
      const data = await res.json();
      if (res.ok && data.status === 'success') {
        const uploadedId = data.fileId || data.record?.fileId || data.record?.action;
        if (uploadedId || data.record) {
          setRecords((prev) =>
            prev.map((r) => {
              if (String(r.id) !== String(row.id)) return r;
              if (data.record) return { ...r, ...data.record };
              return {
                ...r,
                action: uploadedId,
                fileId: uploadedId,
                templateFileName: data.fileName || r.templateFileName,
                fileName: data.fileName || r.fileName
              };
            })
          );
        }
        setMessage(`File uploaded for ${row.formName || row.act || row.id}`);
        await fetchRecords();
      } else {
        setMessage(data.message || data.error || 'Upload failed');
      }
    } catch (err) {
      setMessage('Upload failed: ' + (err.message || 'Unknown error'));
    } finally {
      setUploadingRowId(null);
    }
  };

  const filteredRecords = useMemo(() => {
    const term = searchTerm.trim().toLowerCase();
    if (!term) return records;
    return records.filter((row) =>
      [row.act, row.description, row.sector, row.state, row.formName]
        .some((v) => String(v || '').toLowerCase().includes(term))
    );
  }, [records, searchTerm]);

  const totalPages = Math.max(1, Math.ceil(filteredRecords.length / PAGE_SIZE));
  const effectivePage = Math.min(tablePage, totalPages);
  const pagedRecords = useMemo(() => {
    const start = (effectivePage - 1) * PAGE_SIZE;
    return filteredRecords.slice(start, start + PAGE_SIZE);
  }, [filteredRecords, effectivePage]);
  const paginationItems = useMemo(
    () => buildPaginationItems(effectivePage, totalPages),
    [effectivePage, totalPages]
  );

  return (
    <div className="formmaster-container">
      <div className="formmaster-header">
        <div className="formmaster-header-content">
          <div className="formmaster-header-icon">
            <FileSpreadsheet size={28} />
          </div>
          <div className="formmaster-header-text">
            <h1 className="formmaster-title">Form Master</h1>
            <p className="formmaster-subtitle">
              Each row is unique by Form Name + Act + Description + Sector + State — upload the matching template per row
            </p>
          </div>
        </div>
      </div>

      <div className="formmaster-content">
        <div className="formmaster-import-grid">
          <section className="formmaster-search-card formmaster-import-card">
            <h2 className="formmaster-import-heading">Import Excel (metadata)</h2>
            <p className="formmaster-import-caption">
              Columns: Act, Description, Sector, State (optional: Form Name). Files are uploaded per row in the table below.
            </p>
            <div
              className={`formmaster-dropzone${dragOver ? ' formmaster-dropzone-active' : ''}`}
              onDrop={handleDrop}
              onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
              onDragLeave={() => setDragOver(false)}
            >
              <Upload size={28} className="formmaster-dropzone-icon" />
              <p className="formmaster-dropzone-title">Drag &amp; drop Excel here</p>
              <span className="formmaster-btn-choose">Choose file</span>
              <input
                type="file"
                accept=".xlsx,.xls"
                className="formmaster-file-input"
                onChange={(e) => handleImportFile(e.target.files?.[0])}
              />
            </div>
            <div className="formmaster-import-footer">
              <span className="formmaster-import-filename">{importFile ? importFile.name : 'No file selected'}</span>
              <button
                type="button"
                className="formmaster-btn-import"
                onClick={handleBulkImport}
                disabled={loading || !importFile}
              >
                Import Data
              </button>
            </div>
          </section>

          <section className="formmaster-search-card formmaster-guidelines-card">
            <h2 className="formmaster-import-heading">Guidelines</h2>
            <ul className="formmaster-guidelines-list">
              <li>Import row metadata from Excel first.</li>
              <li>Same form name (e.g. two Form U) with different Act / State — upload a separate file on each row (files are stored with unique names).</li>
              <li>After deploying this fix, re-upload each row&apos;s template once so Maharashtra and other states do not share one file.</li>
              <li>Use <strong>Upload</strong> on that row only (.xlsx / .xls).</li>
              <li><strong>Sync from File Store</strong> pulls existing templates into new rows (file only).</li>
            </ul>
            <button
              type="button"
              className="formmaster-btn-sync"
              onClick={handleSyncToDataStore}
              disabled={syncing}
            >
              <Database size={18} />
              {syncing ? ' Syncing...' : ' Sync from File Store'}
            </button>
            {syncMessage && <p className="formmaster-sync-message">{syncMessage}</p>}
          </section>
        </div>

        {message && (
          <div className={`formmaster-banner${message.toLowerCase().includes('fail') ? ' formmaster-banner-error' : ''}`} role="status">
            {message}
          </div>
        )}

        <div className="formmaster-table-card">
          <div className="formmaster-table-toolbar">
            <h2 className="formmaster-table-title">Form Master records ({filteredRecords.length})</h2>
            <div className="formmaster-toolbar-right">
              <div className="formmaster-search-row">
                <input
                  type="text"
                  className="formmaster-input"
                  placeholder="Search Act, Sector, State..."
                  value={searchTerm}
                  onChange={(e) => setSearchTerm(e.target.value)}
                />
                <button type="button" className="formmaster-btn-search" onClick={() => setTablePage(1)} disabled={loading}>
                  <Search size={18} />
                  Search
                </button>
              </div>
              <button
                type="button"
                className="formmaster-btn-clear"
                onClick={handleClearAll}
                disabled={loading || !records.length}
              >
                Clear All
              </button>
            </div>
          </div>

          {loading && !records.length ? (
            <p className="formmaster-empty">Loading…</p>
          ) : filteredRecords.length === 0 ? (
            <p className="formmaster-empty">No records. Import Excel metadata or sync from File Store.</p>
          ) : (
            <>
              <div className="formmaster-table-wrapper">
                <table className="formmaster-data-table formmaster-data-table-wide">
                  <thead>
                    <tr>
                      <th>S.No</th>
                      <th>Act</th>
                      <th>Description</th>
                      <th>Sector</th>
                      <th>State</th>
                      <th>Form Name</th>
                      <th>Template file</th>
                      <th className="formmaster-th-actions">Action</th>
                    </tr>
                  </thead>
                  <tbody>
                    {pagedRecords.map((row, i) => {
                      const fileId =
                        row.action ||
                        row.fileId ||
                        row.Action ||
                        row.FileId ||
                        null;
                      const templateFileName = row.templateFileName || row.fileName || '';
                      const isUploading = uploadingRowId === String(row.id);
                      return (
                        <tr key={row.id || i} className="formmaster-data-row">
                          <td>{(effectivePage - 1) * PAGE_SIZE + i + 1}</td>
                          <td>{row.act || '—'}</td>
                          <td className="formmaster-cell-desc" title={row.description}>{row.description || '—'}</td>
                          <td>{row.sector || '—'}</td>
                          <td>{row.state || '—'}</td>
                          <td>{row.formName || '—'}</td>
                          <td>
                            {fileId ? (
                              <span className="formmaster-file-name">
                                <FileSpreadsheet size={14} />
                                {templateFileName || 'template.xlsx'}
                              </span>
                            ) : (
                              <span className="formmaster-no-file">No file</span>
                            )}
                          </td>
                          <td>
                            <div className="formmaster-table-actions">
                              <input
                                type="file"
                                accept=".xlsx,.xls"
                                className="formmaster-row-file-input"
                                ref={(el) => { rowFileInputRefs.current[row.id] = el; }}
                                onChange={(e) => {
                                  const f = e.target.files?.[0];
                                  if (f) handleRowFileChange(row, f);
                                  e.target.value = '';
                                }}
                              />
                              <button
                                type="button"
                                className="formmaster-btn-upload"
                                onClick={() => triggerRowUpload(row.id)}
                                disabled={isUploading}
                              >
                                <Upload size={14} />
                                {isUploading ? 'Uploading…' : fileId ? 'Replace' : 'Upload'}
                              </button>
                              {fileId && (
                                <a
                                  href={downloadUrl(fileId, templateFileName || row.formName || 'template.xlsx')}
                                  className="formmaster-download-link"
                                  download
                                >
                                  <Download size={14} />
                                  Download
                                </a>
                              )}
                            </div>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
              <div className="formmaster-pagination">
                <nav className="formmaster-pagination-nav-wrap" aria-label="Table pagination">
                  <button
                    type="button"
                    className="formmaster-pagination-nav"
                    disabled={effectivePage <= 1}
                    onClick={() => setTablePage((p) => Math.max(1, p - 1))}
                    aria-label="Previous page"
                  >
                    ‹
                  </button>
                  <div className="formmaster-pagination-pages">
                    {paginationItems.map((item, idx) =>
                      item === 'ellipsis' ? (
                        <span key={`e-${idx}`} className="formmaster-pagination-ellipsis">…</span>
                      ) : (
                        <button
                          key={item}
                          type="button"
                          className={`formmaster-pagination-page${item === effectivePage ? ' formmaster-pagination-page--active' : ''}`}
                          onClick={() => setTablePage(item)}
                          aria-current={item === effectivePage ? 'page' : undefined}
                        >
                          {item}
                        </button>
                      )
                    )}
                  </div>
                  <button
                    type="button"
                    className="formmaster-pagination-nav"
                    disabled={effectivePage >= totalPages}
                    onClick={() => setTablePage((p) => Math.min(totalPages, p + 1))}
                    aria-label="Next page"
                  >
                    ›
                  </button>
                </nav>
              </div>
            </>
          )}
        </div>
      </div>
    </div>
  );
};

export default Formmaster;
