
import React, { useState, useEffect, useCallback, useMemo } from 'react';
import * as XLSX from 'xlsx';
import './ActsBulk.css';

const API_BASE = '/server/actsbulk_function';
const PAGE_SIZE = 10;

/** Exact same buildPaginationItems as CompanyDetails */
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

/** Trim, fix NBSP, collapse spaces, strip trailing dots — Excel headers vary by template/state. */
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

const coalesceCell = (...vals) => {
  for (const v of vals) {
    if (v === undefined || v === null) continue;
    const s = typeof v === 'string' ? v : String(v);
    if (s.trim() !== '') return v;
  }
  return '';
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

/** When header text differs (state templates) but clearly means applicability. */
const fuzzyApplicability = (normRow) => {
  for (const [nk, v] of Object.entries(normRow)) {
    if (v === undefined || v === null || String(v).trim() === '') continue;
    const compact = nk.replace(/[^a-z0-9]/g, '');
    if (compact.includes('applicability')) return v;
    if (nk.includes('applicable to') || nk === 'applicable' || compact === 'applicable') return v;
  }
  return undefined;
};

/** Match key-compliance column but not penalty / non-compliance / registers. */
const fuzzyKeyCompliance = (normRow) => {
  for (const [nk, v] of Object.entries(normRow)) {
    if (v === undefined || v === null || String(v).trim() === '') continue;
    if (/penalty|non[-\s]?compliance|register/i.test(nk)) continue;
    const compact = nk.replace(/[^a-z0-9]/g, '');
    if (/^keycompliance/.test(compact) || (compact.includes('key') && compact.includes('compliance'))) return v;
  }
  return undefined;
};

// Map Excel row to datastore format (headers may differ across state Excel templates)
const mapRowToRecord = (row) => {
  const norm = buildNormRow(row);
  return {
    sector: coalesceCell(cellByAliases(norm, ['Sector', 'sector']), row.Sector, row.sector),
    acts: coalesceCell(cellByAliases(norm, ['Acts', 'Act', 'acts']), row.Acts, row.Act, row.acts),
    type: coalesceCell(cellByAliases(norm, ['Type', 'type']), row.Type, row.type),
    states: coalesceCell(cellByAliases(norm, ['States', 'State', 'states', 'state']), row.States, row.states),
    description: coalesceCell(cellByAliases(norm, ['Description', 'description']), row.Description, row.description),
    applicability: coalesceCell(
      cellByAliases(norm, [
        'Applicability',
        'Applicable',
        'Applicable To',
        'Scope of Applicability',
        'Application',
        'Act Applicability'
      ]),
      fuzzyApplicability(norm),
      row.Applicability,
      row.applicability
    ),
    keyComplianceRequirements: coalesceCell(
      cellByAliases(norm, [
        'Key Compliance Requirements',
        'Key Compliance Requirement',
        'Key Compliance',
        'Key Compliance Req',
        'Key compliances',
        'Key Statutory Compliances',
        'Compliance Requirements (Key)',
        'KeyComplianceRequirements',
        'Key compliance requirements'
      ]),
      fuzzyKeyCompliance(norm),
      row['Key Compliance Requirements'],
      row.KeyComplianceRequirements,
      row.keyComplianceRequirements
    ),
    dueDate: coalesceCell(
      cellByAliases(norm, ['Due Date', 'DueDate', 'due date', 'duedate']),
      row['Due Date'],
      row.DueDate,
      row.dueDate
    ),
    penaltyforNonCompliance: coalesceCell(
      cellByAliases(norm, [
        'Penalty for Non-Compliance',
        'Penalty for Non Compliance',
        'PenaltyforNonCompliance',
        'Penalty'
      ]),
      row['Penalty for Non-Compliance'],
      row.PenaltyforNonCompliance,
      row.penaltyforNonCompliance
    ),
    registers: coalesceCell(cellByAliases(norm, ['Registers', 'Register', 'registers']), row.Registers, row.registers)
  };
};

const ActsBulk = () => {
  const [file, setFile] = useState(null);
  const [importedData, setImportedData] = useState([]);
  const [loading, setLoading] = useState(false);
  const [message, setMessage] = useState('');
  const [dragOver, setDragOver] = useState(false);
  const [searchTerm, setSearchTerm] = useState('');
  const [tablePage, setTablePage] = useState(1);

  const fetchData = useCallback(async () => {
    setLoading(true);
    setMessage('');
    try {
      const res = await fetch(`${API_BASE}/actsbulk?action=getAll`, { credentials: 'include' });
      const data = await res.json();
      if (data.status === 'success' && Array.isArray(data.data)) {
        setImportedData(data.data);
      } else {
        setImportedData([]);
      }
    } catch (err) {
      setMessage('Failed to load data from datastore');
      setImportedData([]);
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => { fetchData(); }, [fetchData]);

  // Reset to page 1 when search changes
  useEffect(() => { setTablePage(1); }, [searchTerm]);

  const handleFile = (selectedFile) => {
    if (!selectedFile) return;
    const name = (selectedFile.name || '').toLowerCase();
    if (!name.endsWith('.xlsx') && !name.endsWith('.xls')) {
      setMessage('Supports .xlsx, .xls files only');
      return;
    }
    setFile(selectedFile);
    setMessage('');
  };

  const handleDrop = (e) => {
    e.preventDefault();
    setDragOver(false);
    handleFile(e.dataTransfer?.files?.[0]);
  };
  const handleDragOver = (e) => { e.preventDefault(); setDragOver(true); };
  const handleDragLeave = () => setDragOver(false);

  const handleImport = async () => {
    if (!file) { setMessage('Please select an Excel file first'); return; }
    setLoading(true);
    setMessage('');
    try {
      const arrayBuffer = await file.arrayBuffer();
      const wb = XLSX.read(arrayBuffer, { type: 'array' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws);
      if (!rows.length) { setMessage('No rows found in the Excel file'); setLoading(false); return; }
      const records = rows.map(mapRowToRecord);
      const res = await fetch(`${API_BASE}/actsbulk?action=bulkImport`, {
        method: 'POST', credentials: 'include',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(records)
      });
      const data = await res.json();
      if (data.status === 'success') {
        setMessage('Import successful');
        setFile(null);
        await fetchData();
      } else {
        setMessage(data.message || data.error || `Import failed (HTTP ${res.status})`);
      }
    } catch (err) {
      setMessage('Import failed: ' + (err.message || 'Unknown error'));
    } finally {
      setLoading(false);
    }
  };

  const handleClear = async () => {
    if (!window.confirm('Remove all Acts bulk records from the database?')) return;
    setLoading(true);
    setMessage('');
    try {
      const res = await fetch(`${API_BASE}/actsbulk?action=deleteAll`, { method: 'DELETE', credentials: 'include' });
      const data = await res.json();
      if (data.status === 'success') { setMessage('Data cleared'); setImportedData([]); }
      else setMessage(data.message || 'Clear failed');
    } catch (err) {
      setMessage('Clear failed: ' + (err.message || 'Unknown error'));
    } finally {
      setLoading(false);
    }
  };

  // Filter
  const filteredData = useMemo(() => {
    const term = searchTerm.trim().toLowerCase();
    if (!term) return importedData;
    return importedData.filter((row) =>
      [row.acts, row.sector, row.type, row.states, row.description]
        .some((v) => String(v || '').toLowerCase().includes(term))
    );
  }, [importedData, searchTerm]);

  // Pagination
  const totalPages = Math.max(1, Math.ceil(filteredData.length / PAGE_SIZE));
  const effectivePage = Math.min(tablePage, totalPages);

  useEffect(() => {
    setTablePage((p) => Math.min(p, totalPages));
  }, [totalPages]);

  const pagedData = useMemo(() => {
    const start = (effectivePage - 1) * PAGE_SIZE;
    return filteredData.slice(start, start + PAGE_SIZE);
  }, [filteredData, effectivePage]);

  const paginationItems = useMemo(
    () => buildPaginationItems(effectivePage, totalPages),
    [effectivePage, totalPages]
  );

  const startRecord = filteredData.length === 0 ? 0 : (effectivePage - 1) * PAGE_SIZE + 1;
  const endRecord = Math.min(effectivePage * PAGE_SIZE, filteredData.length);

  return (
    <div className="actsbulk-page">

      {/* ── Page heading ── */}
      <header className="actsbulk-header">
        <div className="actsbulk-header-icon">📄</div>
        <div>
          <h1 className="actsbulk-title">Acts Bulk Import</h1>
          <p className="actsbulk-subtitle">Import Excel data with all acts and descriptions</p>
        </div>
      </header>

      {/* ── Upload + Guidelines cards ── */}
      <div className="actsbulk-cards">
        <section className="actsbulk-card actsbulk-upload-card">
          <div className="actsbulk-card-head">
            <h2>Upload Excel File</h2>
          </div>
          <p className="actsbulk-card-caption">Drag and drop your file or browse to upload</p>
          <div
            className={`actsbulk-dropzone${dragOver ? ' actsbulk-dropzone-active' : ''}`}
            onDrop={handleDrop} onDragOver={handleDragOver} onDragLeave={handleDragLeave}
          >
            <span className="actsbulk-dropzone-arrow">☁</span>
            <p className="actsbulk-dropzone-title">Drag &amp; Drop Excel File Here</p>
            <p className="actsbulk-dropzone-or">or</p>
            <span className="actsbulk-btn-choose">Choose File</span>
            <p className="actsbulk-dropzone-hint">Supports .xlsx, .xls files • Maximum file size: 5 MB</p>
            <input type="file" accept=".xlsx,.xls" className="actsbulk-file-input"
              onChange={(e) => handleFile(e.target.files?.[0])} />
          </div>
          <div className="actsbulk-upload-footer">
            <span className="actsbulk-file-name">{file ? file.name : 'No file selected'}</span>
            <button type="button" className="actsbulk-btn actsbulk-btn-import"
              onClick={handleImport} disabled={loading || !file}>
              Import Data
            </button>
          </div>
        </section>

        <section className="actsbulk-card actsbulk-guidelines-card">
          <div className="actsbulk-card-head">
            <h2>Import Guidelines</h2>
          </div>
          <div className="actsbulk-guidelines">
            <div className="actsbulk-guideline-row">
              <span className="actsbulk-guideline-label">File Format</span>
              <span className="actsbulk-guideline-value">Excel (.xlsx, .xls)</span>
            </div>
            <div className="actsbulk-guideline-row">
              <span className="actsbulk-guideline-label">Required Columns</span>
              <span className="actsbulk-guideline-value">
                Sector, Acts, Type, States, Description, Applicability, Key Compliance Requirements, Due Date,
                Penalty for Non-Compliance, Registers
              </span>
            </div>
            <div className="actsbulk-guideline-row">
              <span className="actsbulk-guideline-label">Data Validation</span>
              <span className="actsbulk-guideline-value">Automatic validation on import</span>
            </div>
            <div className="actsbulk-guideline-row">
              <span className="actsbulk-guideline-label">Tips</span>
              <span className="actsbulk-guideline-value">Use the template to ensure correct column format</span>
            </div>
          </div>
          <div className="actsbulk-best-practice">
            Best Practice: Ensure date format is DD-MM-YYYY for accurate date processing.
          </div>
        </section>
      </div>

      {/* ── Imported data card ── */}
      <section className="actsbulk-card actsbulk-data-card">

        {/* Card head */}
        <div className="actsbulk-data-head">
          <div style={{ display: 'flex', alignItems: 'center' }}>
            <span className="actsbulk-data-icon">📑</span>
            <h2>Imported Acts Preview</h2>
          </div>
          <div className="actsbulk-data-stats">
            <div className="actsbulk-stat-card actsbulk-stat-success">
              <strong>{importedData.length}</strong>
              <span>Records Imported</span>
            </div>
            <div className="actsbulk-stat-card actsbulk-stat-warning">
              <strong>0</strong>
              <span>Validation Errors</span>
            </div>
          </div>
        </div>

        {/* Status banner */}
        {message ? <div className="actsbulk-status-banner">{message}</div> : null}

        {/* Toolbar */}
        <div className="actsbulk-data-toolbar">
          <div className="actsbulk-filter-group">
            <input type="text" className="actsbulk-search"
              placeholder="Search by Acts or Sector..."
              value={searchTerm} onChange={(e) => setSearchTerm(e.target.value)} />
          </div>
          <div className="actsbulk-data-actions">
            <button type="button" className="actsbulk-btn actsbulk-btn-clear"
              onClick={handleClear} disabled={loading || !importedData.length}>
              Clear All Data
            </button>
          </div>
        </div>

        {/* Loading */}
        {loading && !importedData.length && (
          <p className="actsbulk-loading">Loading…</p>
        )}

        {/* Table */}
        <div className="actsbulk-table-wrap">
          <table className="actsbulk-table">
            <thead>
              <tr>
                <th>S.No</th>
                <th>Sector</th>
                <th>Acts</th>
                <th>Type</th>
                <th>States</th>
                <th>Description</th>
                <th>Applicability</th>
                <th>Key Compliance</th>
                <th>Due Date</th>
                <th>Penalty for Non-Compliance</th>
                <th>Registers</th>
              </tr>
            </thead>
            <tbody>
              {pagedData.length === 0 ? (
                <tr>
                  <td colSpan={11} className="actsbulk-empty">
                    {loading ? 'Loading data…' : 'No imported data. Upload an Excel file and click Import Data.'}
                  </td>
                </tr>
              ) : (
                pagedData.map((row, i) => (
                  <tr key={row.id || i}>
                    <td>{(effectivePage - 1) * PAGE_SIZE + i + 1}</td>
                    <td>{row.sector}</td>
                    <td>{row.acts}</td>
                    <td>{row.type}</td>
                    <td>{row.states}</td>
                    <td>{row.description}</td>
                    <td>{row.applicability}</td>
                    <td>{row.keyComplianceRequirements}</td>
                    <td>{row.dueDate}</td>
                    <td>{row.penaltyforNonCompliance}</td>
                    <td>{row.registers}</td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>

        {/* ── Pagination (CompanyDetails model) ── */}
        {filteredData.length > 0 && (
          <div className="actsbulk-pagination">
            <div className="actsbulk-pagination-pages">
              {/* Prev */}
              <button
                className="actsbulk-pagination-nav"
                onClick={() => setTablePage((p) => Math.max(1, p - 1))}
                disabled={effectivePage <= 1}
                aria-label="Previous page"
              >
                ‹
              </button>

              {/* Page numbers */}
              {paginationItems.map((item, idx) =>
                item === 'ellipsis' ? (
                  <span key={`ellipsis-${idx}`} className="actsbulk-pagination-ellipsis">…</span>
                ) : (
                  <button
                    key={item}
                    className={`actsbulk-pagination-page${item === effectivePage ? ' actsbulk-pagination-page--active' : ''}`}
                    onClick={() => setTablePage(item)}
                    disabled={item === effectivePage}
                  >
                    {item}
                  </button>
                )
              )}

              {/* Next */}
              <button
                className="actsbulk-pagination-nav"
                onClick={() => setTablePage((p) => Math.min(totalPages, p + 1))}
                disabled={effectivePage >= totalPages}
                aria-label="Next page"
              >
                ›
              </button>
            </div>
          </div>
        )}
      </section>
    </div>
  );
};

export default ActsBulk;
