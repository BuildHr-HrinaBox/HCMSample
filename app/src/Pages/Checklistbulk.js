import React, { useState, useEffect, useCallback, useMemo } from 'react';
import * as XLSX from 'xlsx';
import './Checklistbulk.css';

const API_BASE = '/server/checklistbulk_function';

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

const isMonthlyFrequency = (frequency, dueDate) => {
  const f = String(frequency || '').toLowerCase();
  const d = String(dueDate || '').toLowerCase();
  return f.includes('monthly') || d.includes('monthly basis');
};

const extractMonthlyDayOfMonth = (dueDate) => {
  if (dueDate == null || dueDate === '') return null;
  if (typeof dueDate === 'number' && Number.isInteger(dueDate) && dueDate >= 1 && dueDate <= 31) {
    return dueDate;
  }
  const s = String(dueDate).trim();
  if (!s) return null;
  if (/^\d{1,2}$/.test(s)) {
    const day = parseInt(s, 10);
    if (day >= 1 && day <= 31) return day;
  }
  const legacy = s.match(/^1900-01-(\d{1,2})$/);
  if (legacy) {
    const recovered = parseInt(legacy[1], 10) + 1;
    if (recovered >= 1 && recovered <= 31) return recovered;
  }
  return null;
};

const ordinalDay = (day) => {
  const n = Number(day);
  const mod100 = n % 100;
  if (mod100 >= 11 && mod100 <= 13) return `${n}th`;
  const suffix = ['th', 'st', 'nd', 'rd'];
  return `${n}${suffix[n % 10] || 'th'}`;
};

/** e.g. dueDate 7 + Monthly Basis → "7th Jun" */
const formatChecklistDueDate = (dueDate, frequency) => {
  const raw = String(dueDate ?? '').trim();
  if (!raw) return '-';
  if (!isMonthlyFrequency(frequency, dueDate)) return raw;
  const day = extractMonthlyDayOfMonth(dueDate);
  if (day == null) return raw;
  const month = new Date().toLocaleString('en-GB', { month: 'short' });
  return `${ordinalDay(day)} ${month}`;
};

const normalizeDueDateForImport = (dueDate, frequency) => {
  if (dueDate == null || dueDate === '') return '';
  if (typeof dueDate === 'number' && Number.isInteger(dueDate) && dueDate >= 1 && dueDate <= 31) {
    return String(dueDate);
  }
  const s = String(dueDate).trim();
  if (isMonthlyFrequency(frequency, dueDate)) {
    const day = extractMonthlyDayOfMonth(dueDate);
    if (day != null) return String(day);
  }
  if (/^\d{1,2}$/.test(s)) {
    const day = parseInt(s, 10);
    if (day >= 1 && day <= 31) return String(day);
  }
  return dueDate;
};

// Map Excel row to datastore format (checklistbulk table)
const mapRowToRecord = (row) => {
  const formNumber = row.FormNumber ?? row['Form Number'] ?? '';
  const formNameCol = row['Form Name'] ?? '';
  const canonicalFormName = row.FormName ?? row.formName ?? '';
  const canonicalDescription = row.Description ?? row.description ?? '';

  const hasFormNumber = String(formNumber).trim() !== '';
  let formName;
  let description;

  if (hasFormNumber) {
    formName = String(formNumber).trim();
    description = String(formNameCol || canonicalFormName || canonicalDescription).trim();
  } else {
    formName = String(canonicalFormName).trim();
    description = String(canonicalDescription || formNameCol).trim();
    if (formNameCol && !description && formName) {
      description = String(formNameCol).trim();
    } else if (formNameCol && !formName) {
      formName = String(canonicalFormName || formNameCol).trim();
    }
  }

  const frequency = row.Frequency ?? row.frequency ?? '';
  const rawDueDate = row.DueDate ?? row['Due Date'] ?? row.dueDate ?? '';

  return {
  sector: row.Sector ?? row.sector ?? '',
  state: row.State ?? row.state ?? '',
  act: row.Act ?? row.act ?? '',
  formName,
  description,
  concernedGovtDepartment: row.ConcernedGovtDepartment ?? row['Concerned Govt Department'] ?? row['Govt Department'] ?? row.concernedGovtDepartment ?? '',
  dueDate: normalizeDueDateForImport(rawDueDate, frequency),
  nameOfTheCode: row.Nameofthecode ?? row.NameOfTheCode ?? row['Name of the Code'] ?? row.nameOfTheCode ?? row.nameofthecode ?? '',
  frequency,
  nameOfTheRule: row.NameoftheRule ?? row.NameOfTheRule ?? row['Name of the Rule'] ?? row.nameOfTheRule ?? row.nameoftheRule ?? ''
  };
};

const Checklistbulk = ({ userRole, userEmail }) => {
  const [file, setFile] = useState(null);
  const [importedData, setImportedData] = useState([]);
  const [loading, setLoading] = useState(false);
  const [message, setMessage] = useState('');
  const [dragOver, setDragOver] = useState(false);
  const [searchTerm, setSearchTerm] = useState('');
  const [lastImportAt, setLastImportAt] = useState('');
  const [currentPage, setCurrentPage] = useState(1);
  const rowsPerPage = 10;

  const fetchData = useCallback(async () => {
    setLoading(true);
    setMessage('');
    try {
      const res = await fetch(`${API_BASE}/checklistbulk?action=getAll`);
      const data = await res.json();
      if (data.status === 'success' && Array.isArray(data.data)) {
        setImportedData(data.data);
        if (data.count) {
          setLastImportAt(new Date().toLocaleString());
        }
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
      const records = rows.map(mapRowToRecord);
      const res = await fetch(`${API_BASE}/checklistbulk?action=bulkImport`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(records)
      });
      const data = await res.json();
      if (data.status === 'success') {
        setMessage(data.message || 'Import completed.');
        setFile(null);
        setLastImportAt(new Date().toLocaleString());
        await fetchData();
      } else {
        setMessage(data.message || 'Import failed');
      }
    } catch (err) {
      setMessage('Import failed: ' + (err.message || 'Unknown error'));
    } finally {
      setLoading(false);
    }
  };

  const handleSyncFormFiles = async () => {
    setLoading(true);
    setMessage('');
    try {
      const res = await fetch(`${API_BASE}/checklistbulk?action=syncFormFiles`, { method: 'POST' });
      const data = await res.json();
      if (data.status === 'success' || data.status === 'partial') {
        setMessage(data.message || 'Form files linked from Form Master.');
        await fetchData();
      } else {
        setMessage(data.message || 'Could not link form files.');
      }
    } catch (err) {
      setMessage('Link form files failed: ' + (err.message || 'Unknown error'));
    } finally {
      setLoading(false);
    }
  };

  const handleClear = async () => {
    setLoading(true);
    setMessage('');
    try {
      const res = await fetch(`${API_BASE}/checklistbulk?action=bulkDelete`, { method: 'DELETE' });
      const data = await res.json();
      if (data.status === 'success') {
        setMessage('');
        setImportedData([]);
        setLastImportAt('');
      } else {
        setMessage(data.message || 'Clear failed');
      }
    } catch (err) {
      setMessage('Clear failed: ' + (err.message || 'Unknown error'));
    } finally {
      setLoading(false);
    }
  };

  const filteredData = importedData.filter((row) => {
    const term = searchTerm.trim().toLowerCase();
    return !term || [row.act, row.formName, row.concernedGovtDepartment, row.state, row.sector, row.description, row.nameOfTheCode, row.frequency, row.nameOfTheRule]
      .some((value) => String(value || '').toLowerCase().includes(term));
  });

  // Pagination logic
  const totalPages = Math.max(1, Math.ceil(filteredData.length / rowsPerPage));
  const effectivePage = Math.min(currentPage, totalPages);
  const paginatedData = filteredData.slice((effectivePage - 1) * rowsPerPage, effectivePage * rowsPerPage);

  useEffect(() => {
    setCurrentPage((p) => Math.min(p, totalPages));
  }, [totalPages]);

  const paginationItems = useMemo(
    () => buildPaginationItems(effectivePage, totalPages),
    [effectivePage, totalPages]
  );

  // Reset to first page if filter/search changes
  useEffect(() => {
    setCurrentPage(1);
  }, [searchTerm, importedData]);

  const handleExport = () => {
    if (!filteredData.length) return;
    const worksheet = XLSX.utils.json_to_sheet(filteredData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'ChecklistBulk');
    XLSX.writeFile(workbook, 'checklistbulk-export.xlsx');
  };

  const handleTemplateDownload = () => {
    const templateRows = [{
      Sector: 'Factories Act',
      State: 'Tamilnadu',
      Act: 'The Factories Act, 1948',
      FormNumber: 'Form 11',
      FormName: 'Accident Book',
      ConcernedGovtDepartment: 'Labour Department',
      DueDate: 'Monthly Basis',
      Nameofthecode: 'FAC-AB-011',
      Frequency: 'Monthly',
      NameoftheRule: 'Factories Rules'
    }];
    const worksheet = XLSX.utils.json_to_sheet(templateRows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Template');
    XLSX.writeFile(workbook, 'checklistbulk-template.xlsx');
  };

  return (
    <div className="checklistbulk-page">
      <header className="checklistbulk-header">
        <div>
          <h1 className="checklistbulk-title">Checklist Bulk Import</h1>
          <p className="checklistbulk-subtitle">Import Excel data with sector, state, act, form details, rule details and due dates</p>
        </div>
        <button type="button" className="checklistbulk-btn checklistbulk-btn-outline" onClick={handleTemplateDownload}>
          Download Template
        </button>
      </header>

      <div className="checklistbulk-cards">
        <section className="checklistbulk-card checklistbulk-upload-card">
          <div className="checklistbulk-card-head">
            <h2>Upload Excel File</h2>
          </div>
          <p className="checklistbulk-card-caption">Drag and drop your file or browse to upload</p>
          <div
            className={`checklistbulk-dropzone ${dragOver ? 'checklistbulk-dropzone-active' : ''}`}
            onDrop={handleDrop}
            onDragOver={handleDragOver}
            onDragLeave={handleDragLeave}
          >
            <span className="checklistbulk-dropzone-arrow">☁</span>
            <p className="checklistbulk-dropzone-title">Drag & Drop Excel File Here</p>
            <p className="checklistbulk-dropzone-or">or</p>
            <span className="checklistbulk-btn checklistbulk-btn-choose">Choose File</span>
            <p className="checklistbulk-dropzone-hint">Supports .xlsx, .xls files • Maximum file size: 5 MB</p>
            <input
              type="file"
              accept=".xlsx,.xls"
              className="checklistbulk-file-input"
              onChange={(e) => handleFile(e.target.files?.[0])}
            />
          </div>
          <div className="checklistbulk-upload-footer">
            <span className="checklistbulk-file-name">{file ? file.name : 'No file selected'}</span>
            <button type="button" className="checklistbulk-btn checklistbulk-btn-import" onClick={handleImport} disabled={loading || !file}>
              Import Data
            </button>
          </div>
        </section>

        <section className="checklistbulk-card checklistbulk-guidelines-card">
          <div className="checklistbulk-card-head">
            <h2>Import Guidelines</h2>
          </div>
          <div className="checklistbulk-guidelines">
            <div className="checklistbulk-guideline-row"><span>File Format</span><strong>Excel (.xlsx, .xls)</strong></div>
            <div className="checklistbulk-guideline-row"><span>Required Columns</span><strong>Sector, State, Act, FormNumber, FormName, ConcernedGovtDepartment, DueDate, Nameofthecode, Frequency, NameoftheRule</strong></div>
            <div className="checklistbulk-guideline-row"><span>Data Validation</span><strong>Automatic validation on import</strong></div>
            <div className="checklistbulk-guideline-row"><span>Tips</span><strong>Use the template to ensure correct column format</strong></div>
            <div className="checklistbulk-guideline-row"><span>Form File</span><strong>Each row is linked to Form Master by Form Name + Act + Description + Sector + State (upload templates there first)</strong></div>
          </div>
          <div className="checklistbulk-best-practice">Best Practice: Ensure date format is DD-MM-YYYY for accurate date processing.</div>
        </section>
      </div>

      <section className="checklistbulk-card checklistbulk-data-card">
        <div className="checklistbulk-data-head">
          <div>
            <span className="checklistbulk-data-icon">🛒</span>
            <h2>Imported Data</h2>
          </div>
          <div className="checklistbulk-data-stats">
            <div className="checklistbulk-stat-card checklistbulk-stat-success"><strong>{importedData.length}</strong><span>Records Imported</span></div>
            <div className="checklistbulk-stat-card checklistbulk-stat-warning"><strong>0</strong><span>Validation Errors</span></div>
            <div className="checklistbulk-stat-card checklistbulk-stat-neutral"><strong>{lastImportAt || '-'}</strong><span>Last Import</span></div>
          </div>
        </div>
        {message ? <div className="checklistbulk-status-banner">{message}</div> : null}
        <div className="checklistbulk-data-toolbar">
          <div className="checklistbulk-filter-group">
            <input
              type="text"
              className="checklistbulk-search"
              placeholder="Search by act, form, code, frequency, or rule..."
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
            />
          </div>
          <div className="checklistbulk-data-actions">
            <button
              type="button"
              className="checklistbulk-btn checklistbulk-btn-outline"
              onClick={handleSyncFormFiles}
              disabled={loading || !importedData.length}
              title="Match each row to Form Master template by act, description, sector, state, and form name"
            >
              Link Form Files
            </button>
            <button type="button" className="checklistbulk-btn checklistbulk-btn-outline" onClick={handleExport} disabled={!filteredData.length}>
              Export Data
            </button>
            <button type="button" className="checklistbulk-btn checklistbulk-btn-clear" onClick={handleClear} disabled={loading || !importedData.length}>
              Clear All Data
            </button>
          </div>
        </div>
        <div className="checklistbulk-table-wrap">
          {loading && !importedData.length ? (
            <p className="checklistbulk-loading">Loading...</p>
          ) : (
            <table className="checklistbulk-table">
              <thead>
                <tr>
                  <th>S.No</th>
                  <th>Sector</th>
                  <th>State</th>
                  <th>Name of the Code</th>
                  <th>Act</th>
                  <th>Name of the Rule</th>
                  <th>Form Number</th>
                  <th>Form Name</th>
                  <th>Govt Department</th>
                  <th>Frequency</th>
                  <th>Due Date</th>
                  <th>Form File</th>
                </tr>
              </thead>
              <tbody>
                {filteredData.length === 0 ? (
                  <tr><td colSpan={12} className="checklistbulk-empty">No imported data. Upload an Excel file and click Import Data.</td></tr>
                ) : (
                  paginatedData.map((row, i) => (
                    <tr key={row.id || i}>
                      <td>{(effectivePage - 1) * rowsPerPage + i + 1}</td>
                      <td>{row.sector}</td>
                      <td>{row.state}</td>
                      <td>{row.nameOfTheCode || '-'}</td>
                      <td>{row.act}</td>
                      <td>{row.nameOfTheRule || '-'}</td>
                      <td><span className="checklistbulk-pill checklistbulk-pill-form">{row.formName || '-'}</span></td>
                      <td>{row.description}</td>
                      <td>{row.concernedGovtDepartment}</td>
                      <td>{row.frequency || '-'}</td>
                      <td>{formatChecklistDueDate(row.dueDate, row.frequency)}</td>
                      <td title={row.formFile ? `File id: ${row.formFile}` : 'Upload matching row in Form Master, then Link Form Files'}>
                        {row.formFileName || (row.formFile ? 'Linked' : '—')}
                      </td>
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          )}
        </div>
        {filteredData.length > 0 ? (
          <nav className="checklistbulk-pagination" aria-label="Table pagination">
            <button
              type="button"
              className="checklistbulk-pagination-nav"
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
              className="checklistbulk-pagination-nav"
              disabled={effectivePage <= 1}
              onClick={() => setCurrentPage((p) => Math.max(1, p - 1))}
              title="Previous page"
              aria-label="Previous page"
            >
              <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden>
                <polyline points="15 18 9 12 15 6" />
              </svg>
            </button>
            <div className="checklistbulk-pagination-pages">
              {paginationItems.map((item, i) =>
                item === 'ellipsis' ? (
                  <span key={`e-${i}`} className="checklistbulk-pagination-ellipsis" aria-hidden>
                    ...
                  </span>
                ) : (
                  <button
                    key={item}
                    type="button"
                    className={`checklistbulk-pagination-page${item === effectivePage ? ' checklistbulk-pagination-page--active' : ''}`}
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
              className="checklistbulk-pagination-nav"
              disabled={effectivePage >= totalPages}
              onClick={() => setCurrentPage((p) => Math.min(totalPages, p + 1))}
              title="Next page"
              aria-label="Next page"
            >
              <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden>
                <polyline points="9 18 15 12 9 6" />
              </svg>
            </button>
            <button
              type="button"
              className="checklistbulk-pagination-nav"
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
      </section>
    </div>
  );
};

export default Checklistbulk;