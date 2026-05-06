import React, { useState, useEffect, useCallback } from 'react';
import * as XLSX from 'xlsx';
import './AuditMasterImport.css';

const API_BASE = '/server/audit_function';

// Map Excel row to Audit format (Sector, Act, Description, RecordCategory, MaximumMarks, Applicability, Remarks)
// Supports various column name variations
const mapRowToRecord = (row) => ({
  sector: row.Sector ?? row.sector ?? '',
  act: String(row.Act ?? row.Acts ?? row.act ?? row.acts ?? row['Act Name'] ?? '').trim(),
  description: row.Description ?? row.description ?? '',
  recordCategory: row.RecordCategory ?? row['Record Category'] ?? row.recordCategory ?? '',
  maximumMarks: String(row.MaximumMarks ?? row['Maximum Marks'] ?? row.maximumMarks ?? '').trim() || '',
  applicability: row.Applicability ?? row.applicability ?? '',
  remarks: row.Remarks ?? row.remarks ?? ''
});

const AuditMasterImport = ({ userRole, userEmail }) => {
  const [file, setFile] = useState(null);
  const [importedData, setImportedData] = useState([]);
  const [loading, setLoading] = useState(false);
  const [message, setMessage] = useState('');
  const [dragOver, setDragOver] = useState(false);

  const fetchData = useCallback(async () => {
    try {
      const res = await fetch(`${API_BASE}/auditmaster?action=getAll`);
      const data = await res.json();
      if (data.status === 'success' && Array.isArray(data.data)) {
        setImportedData(data.data);
      } else {
        setImportedData([]);
      }
    } catch {
      setImportedData([]);
    }
  }, []);

  useEffect(() => {
    fetchData();
  }, [fetchData]);

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

  const handleDragOver = (e) => {
    e.preventDefault();
    setDragOver(true);
  };

  const handleDragLeave = () => setDragOver(false);

  const handleImport = async () => {
    if (!file) {
      setMessage('Please select an Excel file first');
      return;
    }
    setLoading(true);
    setMessage('');
    try {
      const arrayBuffer = await file.arrayBuffer();
      const wb = XLSX.read(arrayBuffer, { type: 'array' });
      const firstSheet = wb.SheetNames[0];
      const ws = wb.Sheets[firstSheet];
      const rows = XLSX.utils.sheet_to_json(ws);
      if (!rows.length) {
        setMessage('No rows found in the Excel file');
        setLoading(false);
        return;
      }
      const records = rows.map(mapRowToRecord).filter((r) => (r.act || '').trim());
      if (!records.length) {
        setMessage('No valid rows with Act column. Excel must have Act (or Acts) column.');
        setLoading(false);
        return;
      }
      const res = await fetch(`${API_BASE}/auditmaster?action=bulkImport`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(records)
      });
      const data = await res.json();
      if (data.status === 'success') {
        const count = (data.data && data.data.length) || records.length;
        setMessage(`${count} records imported successfully`);
        setFile(null);
        await fetchData();
      } else {
        setMessage(data.message || data.error || 'Import failed');
      }
    } catch (err) {
      setMessage('Import failed: ' + (err.message || 'Unknown error'));
    } finally {
      setLoading(false);
    }
  };

  const handleClear = async () => {
    if (!window.confirm('Clear all imported audit master data?')) return;
    setLoading(true);
    setMessage('');
    try {
      const res = await fetch(`${API_BASE}/auditmaster?action=deleteAll`, { method: 'DELETE' });
      const data = await res.json();
      if (data.status === 'success') {
        setMessage('Data cleared');
        setImportedData([]);
      } else {
        setMessage(data.message || 'Clear failed');
      }
    } catch (err) {
      setMessage('Clear failed: ' + (err.message || 'Unknown error'));
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="audit-master-import-page">
      <header className="audit-master-import-header">
        <div className="audit-master-import-header-icon">
          <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
            <path d="M22 11.08V12a10 10 0 1 1-5.93-9.14" />
            <polyline points="22 4 12 14.01 9 11.01" />
          </svg>
        </div>
        <div>
          <h1 className="audit-master-import-title">Audit Master Import</h1>
          <p className="audit-master-import-subtitle">Import Excel data with Sector, Act, Description, RecordCategory, and MaximumMarks</p>
        </div>
      </header>

      <div className="audit-master-import-cards">
        <section className="audit-master-import-card audit-master-import-upload-card">
          <div className="audit-master-import-card-head">
            <span className="audit-master-import-card-icon">
              <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
                <polyline points="17 8 12 3 7 8" />
                <line x1="12" y1="3" x2="12" y2="15" />
              </svg>
            </span>
            <h2>UPLOAD EXCEL FILE</h2>
          </div>
          <div
            className={`audit-master-import-dropzone ${dragOver ? 'audit-master-import-dropzone-active' : ''}`}
            onDrop={handleDrop}
            onDragOver={handleDragOver}
            onDragLeave={handleDragLeave}
          >
            <span className="audit-master-import-dropzone-arrow">
              <svg width="40" height="40" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
                <polyline points="17 8 12 3 7 8" />
                <line x1="12" y1="3" x2="12" y2="15" />
              </svg>
            </span>
            <p>Drop your Excel file here</p>
            <p className="audit-master-import-dropzone-or">or click to browse files</p>
            <span className="audit-master-import-dropzone-badge">Supports .xlsx, .xls files</span>
            <input
              type="file"
              accept=".xlsx,.xls"
              className="audit-master-import-file-input"
              onChange={(e) => handleFile(e.target.files?.[0])}
            />
          </div>
          <button type="button" className="audit-master-import-btn audit-master-import-btn-primary" onClick={handleImport} disabled={loading || !file}>
            <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
              <polyline points="17 8 12 3 7 8" />
              <line x1="12" y1="3" x2="12" y2="15" />
            </svg>
            Import Data
          </button>
        </section>

        <section className="audit-master-import-card audit-master-import-guidelines-card">
          <div className="audit-master-import-card-head">
            <span className="audit-master-import-card-icon audit-master-import-icon-info">
              <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <circle cx="12" cy="12" r="10" />
                <line x1="12" y1="16" x2="12" y2="12" />
                <line x1="12" y1="8" x2="12.01" y2="8" />
              </svg>
            </span>
            <h2>IMPORT GUIDELINES</h2>
          </div>
          <ul className="audit-master-import-guidelines">
            <li>
              <span className="audit-master-import-guideline-icon audit-master-import-icon-doc">D</span>
              Excel (.xlsx, .xls)
            </li>
            <li>
              <span className="audit-master-import-guideline-icon audit-master-import-icon-check">✓</span>
              Sector, Act, Description, RecordCategory, Maximum Marks
            </li>
            <li>
              <span className="audit-master-import-guideline-icon audit-master-import-icon-star">★</span>
              Automatic validation on import
            </li>
          </ul>
        </section>
      </div>

      <section className="audit-master-import-data-card">
        <div className="audit-master-import-data-head">
          <div>
            <span className="audit-master-import-data-icon">
              <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <path d="M6 2L3 6v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2V6l-3-4z" />
                <line x1="3" y1="6" x2="21" y2="6" />
                <path d="M16 10a4 4 0 0 1-8 0" />
              </svg>
            </span>
            <h2>Imported Data</h2>
          </div>
          <div className="audit-master-import-data-actions">
            {message && <span className="audit-master-import-status-msg">{message}</span>}
            <button type="button" className="audit-master-import-btn audit-master-import-btn-clear" onClick={handleClear} disabled={loading || !importedData.length}>
              <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <line x1="18" y1="6" x2="6" y2="18" />
                <line x1="6" y1="6" x2="18" y2="18" />
              </svg>
              Clear Data
            </button>
          </div>
        </div>
        <div className="audit-master-import-table-wrap">
          {loading && !importedData.length ? (
            <p className="audit-master-import-loading">Loading...</p>
          ) : (
            <table className="audit-master-import-table">
              <thead>
                <tr>
                  <th>S.No</th>
                  <th>Sector</th>
                  <th>Act</th>
                  <th>Description</th>
                  <th>Record Category</th>
                  <th>Maximum Marks</th>
                </tr>
              </thead>
              <tbody>
                {importedData.length === 0 ? (
                  <tr>
                    <td colSpan={6} className="audit-master-import-empty">
                      No imported data. Upload an Excel file and click Import Data.
                    </td>
                  </tr>
                ) : (
                  importedData.map((row, i) => (
                    <tr key={row.id || i}>
                      <td>{i + 1}</td>
                      <td>{row.sector || '—'}</td>
                      <td>{row.act || '—'}</td>
                      <td>{row.description || '—'}</td>
                      <td>{row.recordCategory || '—'}</td>
                      <td>{row.maximumMarks || '—'}</td>
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          )}
        </div>
      </section>
    </div>
  );
};

export default AuditMasterImport;
