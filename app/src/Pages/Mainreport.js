import React, { useMemo, useState, useEffect, useCallback } from 'react';
import { useNavigate, useSearchParams } from 'react-router-dom';
import { jsPDF } from 'jspdf';
import './Mainreport.css';

const API = '/server/mainreport_function/mainreport';

/** Full month names - must match backend `normalizeMonth` / MainReport `Months` values */
const ALL_MONTHS = [
  'January',
  'February',
  'March',
  'April',
  'May',
  'June',
  'July',
  'August',
  'September',
  'October',
  'November',
  'December'
];

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

const getStatusLabel = (row) => {
  const raw = String(row?.status || '').trim().toLowerCase();
  if (raw.includes('complete')) return 'Completed';
  if (raw.includes('pending')) return 'Pending';
  if (raw.includes('yet') || raw.includes('start')) return 'Yet to Complete';
  return row?.draftFileUrl ? 'Completed' : 'Yet to Complete';
};

const toIntOrNull = (v) => {
  const n = Number.parseInt(String(v || ''), 10);
  return Number.isFinite(n) ? n : null;
};

const monthIndexFromName = (name) => {
  const idx = ALL_MONTHS.findIndex((m) => m.toLowerCase() === String(name || '').toLowerCase());
  return idx >= 0 ? idx : null;
};

const clamp = (n, min, max) => Math.min(max, Math.max(min, n));

const formatMaybeDate = (value) => {
  if (!value) return '';
  const d = new Date(value);
  if (Number.isNaN(d.getTime())) return String(value);
  return d.toLocaleString(undefined, { year: 'numeric', month: 'short', day: '2-digit' });
};

const getRowUpdatedAt = (row) =>
  row?.lastUpdated || row?.updatedAt || row?.updated_at || row?.modifiedAt || row?.createdAt || null;

const Mainreport = ({ userEmail: userEmailProp }) => {
  const navigate = useNavigate();
  const [searchParams] = useSearchParams();
  const siteFromUrl = searchParams.get('site') || '';
  const yearFromUrl = searchParams.get('year');
  const monthFromUrl = searchParams.get('month');

  const effectiveUserEmail = useMemo(
    () =>
      String(
        userEmailProp ||
          (typeof localStorage !== 'undefined' ? localStorage.getItem('userEmail') : '') ||
          ''
      ).trim(),
    [userEmailProp]
  );

  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');
  const [years, setYears] = useState([]);
  const [tableRows, setTableRows] = useState([]);
  const [siteNames, setSiteNames] = useState([]);
  const [tableLoading, setTableLoading] = useState(false);
  const selectedYear = yearFromUrl ? String(yearFromUrl) : '';
  const selectedMonth = monthFromUrl ? String(monthFromUrl) : '';
  const selectedSite = siteFromUrl ? String(siteFromUrl) : '';

  // Filter UI state (Apply/Reset like screenshot)
  const [draftYear, setDraftYear] = useState(selectedYear);
  const [draftMonth, setDraftMonth] = useState(selectedMonth);
  const [draftSite, setDraftSite] = useState(selectedSite);
  const [showDraftOnly, setShowDraftOnly] = useState(false);
  // Pagination state
  const [currentPage, setCurrentPage] = useState(1);
  const PAGE_SIZE = 10;

  const loadSummary = useCallback(async () => {
    setError('');
    setLoading(true);
    try {
      const res = await fetch(`${API}/summary`, { cache: 'no-store' });
      const json = await res.json();
      if (json.status !== 'success' || !json.data) {
        throw new Error(json.message || 'Invalid response');
      }
      const { years: yList } = json.data;
      setYears(Array.isArray(yList) ? yList : []);
    } catch (e) {
      setError(e.message || 'Failed to load main report');
      setYears([]);
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    loadSummary();
  }, [loadSummary]);

  const loadEntries = useCallback(async (year, month, siteValue = '') => {
    setTableLoading(true);
    setError('');
    try {
      const q = new URLSearchParams({
        year: String(year),
        month: String(month),
        userEmail: effectiveUserEmail,
        site: String(siteValue || '')
      });
      const res = await fetch(`${API}/entries?${q}`, { cache: 'no-store' });
      const json = await res.json();
      if (json.status !== 'success' || !json.data) {
        throw new Error(json.message || 'Invalid response');
      }
      setTableRows(Array.isArray(json.data.rows) ? json.data.rows : []);
      setSiteNames(Array.isArray(json.data.siteNames) ? json.data.siteNames : []);
    } catch (e) {
      setError(e.message || 'Failed to load rows');
      setTableRows([]);
      setSiteNames([]);
    } finally {
      setTableLoading(false);
    }
  }, [effectiveUserEmail]);

  // Load rows: before Apply use draft year/month + all sites; after Apply use URL (year/month/site) so site filter refetches.
  useEffect(() => {
    const hasApplied = Boolean(selectedYear && selectedMonth);
    const y = hasApplied ? selectedYear : draftYear;
    const m = hasApplied ? selectedMonth : draftMonth;
    const site = hasApplied ? String(selectedSite || '').trim() : '';
    if (!y || !m) {
      setSiteNames([]);
      return;
    }
    loadEntries(y, m, site);
  }, [selectedYear, selectedMonth, selectedSite, draftYear, draftMonth, loadEntries]);

  // Keep filter controls synced when URL changes (sidebar nav, back/forward, month pills).
  useEffect(() => {
    setDraftYear(selectedYear);
    setDraftMonth(selectedMonth);
    setDraftSite(selectedSite);
  }, [selectedYear, selectedMonth, selectedSite]);

  const applyFilters = useCallback(() => {
    const y = String(draftYear || '').trim();
    const m = String(draftMonth || '').trim();
    const s = String(draftSite || '').trim();

    if (!y || !m) {
      setError('Year and month are required.');
      return;
    }
    setError('');

    const q = new URLSearchParams();
    q.set('year', y);
    q.set('month', m);
    if (s) q.set('site', s);
    navigate(`/mainreport?${q.toString()}`);
  }, [draftYear, draftMonth, draftSite, navigate]);

  const resetFilters = useCallback(() => {
    setDraftYear('');
    setDraftMonth('');
    setDraftSite('');
    setShowDraftOnly(false);
    navigate('/mainreport');
    setTableRows([]);
    setSiteNames([]);
  }, [navigate]);

  const onMonthPillClick = (monthName) => {
    if (!draftYear) return;
    setDraftMonth(monthName);
    setDraftSite('');
    const q = new URLSearchParams();
    q.set('year', String(draftYear));
    q.set('month', String(monthName));
    navigate(`/mainreport?${q.toString()}`);
    setTableRows([]);
    setSiteNames([]);
  };

  const downloadConsolidatedPdf = () => {
    const doc = new jsPDF({ unit: 'pt', format: 'a4' });
    const pageWidth = doc.internal.pageSize.getWidth();

    const title = `Consolidated Report - ${selectedMonth || ''} ${selectedYear || ''}${selectedSite ? ` (${selectedSite})` : ''}`;
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(16);
    doc.text(title, 32, 36);

    const classified = tableRows.reduce(
      (acc, row) => {
        const status = getStatusLabel(row);
        if (status === 'Completed') acc.completed.push(row);
        else if (status === 'Pending') acc.pending.push(row);
        else acc.yetToStart.push(row);
        return acc;
      },
      { completed: [], pending: [], yetToStart: [] }
    );

    const boxes = [
      { label: 'Completed', count: classified.completed.length, color: [22, 163, 74] },
      { label: 'Pending', count: classified.pending.length, color: [217, 119, 6] },
      { label: 'Yet to Complete', count: classified.yetToStart.length, color: [234, 179, 8] }
    ];

    const startX = 32;
    let y = 64;
    const gap = 12;
    const boxW = (pageWidth - startX * 2 - gap * 2) / 3;
    const boxH = 72;

    boxes.forEach((b, idx) => {
      const xPos = startX + idx * (boxW + gap);
      const yPos = y;
      doc.setDrawColor(...b.color);
      doc.setFillColor(255, 255, 255);
      doc.roundedRect(xPos, yPos, boxW, boxH, 8, 8, 'FD');
      doc.setFont('helvetica', 'bold');
      doc.setFontSize(12);
      doc.setTextColor(31, 41, 55);
      doc.text(b.label, xPos + 12, yPos + 22);
      doc.setFontSize(22);
      doc.setTextColor(...b.color);
      doc.text(String(b.count), xPos + 12, yPos + 54);
    });

    y = y + boxH + 18;
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(12);
    doc.setTextColor(31, 41, 55);
    doc.text(`Month Data: ${selectedMonth || ''} ${selectedYear || ''}`, 32, y);
    y += 16;

    const header = ['S.NO', 'Form Name', 'DraftFile', 'Status'];
    const colW = [55, 245, 130, 90];
    const rowH = 18;
    const tableX = 32;
    let tableY = y;

    const drawHeader = () => {
      let x = tableX;
      doc.setFillColor(249, 250, 251);
      doc.rect(tableX, tableY, colW.reduce((a, b) => a + b, 0), rowH, 'F');
      doc.setLineWidth(1.2);
      doc.setDrawColor(120, 120, 120);
      doc.rect(tableX, tableY, colW.reduce((a, b) => a + b, 0), rowH);
      // Draw vertical separators for header columns
      let headerDividerX = tableX;
      for (let i = 0; i < colW.length - 1; i += 1) {
        headerDividerX += colW[i];
        doc.line(headerDividerX, tableY, headerDividerX, tableY + rowH);
      }
      doc.setFont('helvetica', 'bold');
      doc.setFontSize(10);
      doc.setTextColor(55, 65, 81);
      header.forEach((h, i) => {
        doc.text(h, x + 6, tableY + 12);
        x += colW[i];
      });
      tableY += rowH;
    };

    drawHeader();

    const rows = tableRows.length
      ? tableRows.map((row, idx) => [
          String(idx + 1),
          String(row.formName || '-'),
          String(row.draftFileName || '-'),
          getStatusLabel(row)
        ])
      : [['', 'No rows for this month.', '', '']];

    rows.forEach((r) => {
      if (tableY > 770) {
        doc.addPage();
        tableY = 36;
        drawHeader();
      }
      let x = tableX;
      doc.setLineWidth(1.1);
      doc.setDrawColor(130, 130, 130);
      doc.rect(tableX, tableY, colW.reduce((a, b) => a + b, 0), rowH);
      // Draw vertical separators for body columns
      let rowDividerX = tableX;
      for (let i = 0; i < colW.length - 1; i += 1) {
        rowDividerX += colW[i];
        doc.line(rowDividerX, tableY, rowDividerX, tableY + rowH);
      }
      doc.setFont('helvetica', 'normal');
      doc.setFontSize(9);
      r.forEach((cell, i) => {
        if (i === 3) {
          const statusLower = String(cell || '').toLowerCase();
          if (statusLower.includes('complete') && !statusLower.includes('yet')) doc.setTextColor(22, 163, 74);
          else if (statusLower.includes('yet')) doc.setTextColor(234, 179, 8);
          else if (statusLower.includes('pending')) doc.setTextColor(217, 119, 6);
          else doc.setTextColor(75, 85, 99);
        } else {
          doc.setTextColor(75, 85, 99);
        }
        doc.text(String(cell), x + 6, tableY + 12, { maxWidth: colW[i] - 10 });
        x += colW[i];
      });
      tableY += rowH;
    });

    const fname = `Consolidated_Report_${selectedYear || ''}_${selectedMonth || ''}_${selectedSite || 'All'}.pdf`;
    doc.save(fname.replace(/\s+/g, '_'));
  };

  const statusCounts = useMemo(() => {
    const base = { total: 0, yet: 0, pending: 0, drafts: 0 };
    if (!Array.isArray(tableRows) || tableRows.length === 0) return base;
    const counts = tableRows.reduce((acc, row) => {
      acc.total += 1;
      const st = getStatusLabel(row);
      if (st === 'Pending') acc.pending += 1;
      else if (st !== 'Completed') acc.yet += 1;
      if (row?.draftFileUrl) acc.drafts += 1;
      return acc;
    }, base);
    return counts;
  }, [tableRows]);

  const filteredRows = useMemo(() => {
    if (!showDraftOnly) return tableRows;
    return (Array.isArray(tableRows) ? tableRows : []).filter((r) => !!r?.draftFileUrl);
  }, [tableRows, showDraftOnly]);

  // Pagination: slice filteredRows for current page
  const totalPages = Math.max(1, Math.ceil(filteredRows.length / PAGE_SIZE));
  const effectivePage = Math.min(currentPage, totalPages);
  const pagedRows = filteredRows.slice((effectivePage - 1) * PAGE_SIZE, effectivePage * PAGE_SIZE);
  const paginationItems = useMemo(
    () => buildPaginationItems(effectivePage, totalPages),
    [effectivePage, totalPages]
  );

  // Reset to page 1 when filters change
  useEffect(() => {
    setCurrentPage(1);
  }, [filteredRows]);

  useEffect(() => {
    setCurrentPage((p) => Math.min(p, totalPages));
  }, [totalPages]);

  return (
    <div className="mainreport-page">
      <div className="mainreport-wrap">
        <div className="mr-shell">
          <header className="mr-topbar">
            <div className="mr-titleblock">
              <div className="mr-title-row">
                <div className="mr-title-icon" aria-hidden>
                  <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                    <rect x="3" y="4" width="18" height="18" rx="2" ry="2" />
                    <line x1="16" y1="2" x2="16" y2="6" />
                    <line x1="8" y1="2" x2="8" y2="6" />
                    <line x1="3" y1="10" x2="21" y2="10" />
                  </svg>
                </div>
                <h1 className="mr-title">Reports</h1>
              </div>
            </div>
          </header>

          {error ? <div className="mr-alert">{error}</div> : null}

          <section className="mr-card mr-filters" aria-label="Filters & Selection">
          <div className="mr-card-head">
            <div className="mr-card-title">
              <span>Filters &amp; Selection</span>
            </div>
          </div>
          <form
            className="mr-filter-form"
            onSubmit={(e) => {
              e.preventDefault();
              applyFilters();
            }}
          >
          <div className="mr-filter-grid">
            <label className="mr-field">
              <span className="mr-label">
                Year <abbr className="mr-required" title="Required">*</abbr>
              </span>
              <select
                className="mr-select"
                name="mainreport-year"
                value={draftYear}
                required
                aria-required="true"
                onChange={(e) => {
                  const y = e.target.value;
                  setDraftYear(y);
                  setDraftMonth('');
                  setDraftSite('');
                }}
              >
                <option value="">Select year</option>
                {(Array.isArray(years) ? years : []).map((y) => (
                  <option key={y} value={String(y)}>{y}</option>
                ))}
              </select>
            </label>

            <label className="mr-field">
              <span className="mr-label">
                Month <abbr className="mr-required" title="Required">*</abbr>
              </span>
              <select
                className="mr-select"
                name="mainreport-month"
                value={draftMonth}
                required={Boolean(draftYear)}
                aria-required="true"
                onChange={(e) => {
                  setDraftMonth(e.target.value);
                  setDraftSite('');
                }}
                disabled={!draftYear}
              >
                <option value="">Select month</option>
                {ALL_MONTHS.map((m) => (
                  <option key={m} value={m}>
                    {m}
                  </option>
                ))}
              </select>
            </label>

            <label className="mr-field">
              <span className="mr-label">Site</span>
              <select
                className="mr-select"
                value={draftSite}
                onChange={(e) => setDraftSite(e.target.value)}
                disabled={!draftYear || !draftMonth}
              >
                <option value="">All</option>
                {(Array.isArray(siteNames) ? siteNames : []).map((s) => (
                  <option key={s} value={s}>{s}</option>
                ))}
              </select>
            </label>

            <div className="mr-filter-actions">
              <button type="button" className="mr-btn mr-btn-ghost" onClick={resetFilters}>
                Reset
              </button>
              <button
                type="submit"
                className="mr-btn mr-btn-primary"
                disabled={!draftYear || !draftMonth}
                title={!draftYear || !draftMonth ? 'Select year and month (required).' : 'Apply filters'}
              >
                Apply Filters
              </button>
            </div>
          </div>
          </form>
          <div className="mr-filter-hint">
            {loading ? <span className="mr-muted">Loading years…</span> : null}
          </div>
          </section>

          <section className="mr-kpis" aria-label="Summary">
          <div className="mr-kpi mr-kpi--forms">
            <div className="mr-kpi-label">Total Forms</div>
            <div className="mr-kpi-value">{statusCounts.total}</div>
          </div>
          <div className="mr-kpi mr-kpi--drafts">
            <div className="mr-kpi-label">Completed</div>
            <div className="mr-kpi-value">{statusCounts.drafts}</div>
          </div>
          <div className="mr-kpi mr-kpi--yet">
            <div className="mr-kpi-label">Pending</div>
            <div className="mr-kpi-value">{statusCounts.yet}</div>
          </div>
          </section>


          <section className="mr-grid">
          <div className="mr-left">
            <section className="mr-card mr-table-card" aria-label="Forms Summary">
              <div className="mr-table-head">
                <div>
                  <div className="mr-table-title">Forms Summary</div>
                  <div className="mr-muted">List of forms with status and draft file availability for the selected period</div>
                </div>
                <div className="mr-table-actions">
                  <button
                    type="button"
                    className={`mr-seg${!showDraftOnly ? ' mr-seg--active' : ''}`}
                    onClick={() => setShowDraftOnly(false)}
                    disabled={!selectedYear || !selectedMonth}
                  >
                    Consolidated Report
                  </button>
                  <button
                    type="button"
                    className={`mr-seg${showDraftOnly ? ' mr-seg--active' : ''}`}
                    onClick={() => setShowDraftOnly(true)}
                    disabled={!selectedYear || !selectedMonth}
                  >
                    Draft Files Only
                  </button>
                  <button
                    type="button"
                    className="mr-btn mr-btn-primary"
                    onClick={downloadConsolidatedPdf}
                    disabled={!selectedYear || !selectedMonth}
                    title={!selectedYear || !selectedMonth ? 'Select Year/Month first' : 'Download PDF'}
                  >
                    Download Report
                  </button>
                </div>
              </div>

              <div className="mr-table-wrap">
                <table className="mr-table">
                  <thead>
                    <tr>
                      <th style={{ width: 70 }}>S.No</th>
                      <th style={{ maxWidth: 190, width: 190 }}>Form Name</th>
                      <th style={{ maxWidth: 260, width: 260 }}>Act</th>
                      <th style={{ maxWidth: 360, width: 360 }}>Description</th>
                      <th style={{ width: 180 }}>Status</th>
                      <th>Draft File</th>
                    </tr>
                  </thead>
                  <tbody>
                    {!selectedYear || !selectedMonth ? (
                      <tr>
                        <td colSpan={6} className="mr-empty">
                          Select Year and Month, then click <strong>Apply Filters</strong>.
                        </td>
                      </tr>
                    ) : tableLoading ? (
                      <tr>
                        <td colSpan={6} className="mr-empty">Loading…</td>
                      </tr>
                    ) : filteredRows.length === 0 ? (
                      <tr>
                        <td colSpan={6} className="mr-empty">No rows for this selection.</td>
                      </tr>
                    ) : (
                      pagedRows.map((row, idx) => {
                        const status = getStatusLabel(row);
                        const statusClass =
                          status === 'Completed'
                            ? 'mr-status--ok'
                            : status === 'Yet to Complete'
                              ? 'mr-status--yet'
                              : 'mr-status--pending';
                        const draftName = row?.draftFileName || '';
                        return (
                          <tr key={row?.rowId != null ? String(row.rowId) : `${(currentPage - 1) * PAGE_SIZE + idx}-${draftName}`}>
                            <td>{(currentPage - 1) * PAGE_SIZE + idx + 1}</td>
                            <td
                              className="mr-strong"
                              style={{ maxWidth: 190, width: 190, whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}
                              title={row?.formName || ''}
                            >
                              {row?.formName || ''}
                            </td>
                            <td style={{ maxWidth: 260, width: 260, whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>
                              {row?.act || ''}
                            </td>
                            <td
                              style={{
                                maxWidth: 360,
                                width: 360,
                                overflow: 'hidden',
                                textOverflow: 'ellipsis',
                                whiteSpace: 'normal',
                                display: '-webkit-box',
                                WebkitLineClamp: 2,
                                WebkitBoxOrient: 'vertical'
                              }}
                              title={row?.description || ''}
                            >
                              {row?.description || ''}
                            </td>
                            <td>
                              <span className={`mr-status ${statusClass}`}>{status}</span>
                            </td>
                            <td>
                              {row?.draftFileUrl ? (
                                <a className="mr-link" href={row.draftFileUrl} target="_blank" rel="noopener noreferrer">
                                  {draftName || 'Open draft'}
                                </a>
                              ) : (
                                <span className="mr-muted">{draftName}</span>
                              )}
                            </td>
                          </tr>
                        );
                      })
                    )}
                  </tbody>
                </table>
              </div>
              {filteredRows.length > PAGE_SIZE ? (
                <nav className="mr-pagination" aria-label="Table pagination">
                  <button
                    type="button"
                    className="mr-pagination-nav"
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
                    className="mr-pagination-nav"
                    disabled={effectivePage <= 1}
                    onClick={() => setCurrentPage((p) => Math.max(1, p - 1))}
                    title="Previous page"
                    aria-label="Previous page"
                  >
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden>
                      <polyline points="15 18 9 12 15 6" />
                    </svg>
                  </button>
                  <div className="mr-pagination-pages">
                    {paginationItems.map((item, i) =>
                      item === 'ellipsis' ? (
                        <span key={`e-${i}`} className="mr-pagination-ellipsis" aria-hidden>
                          ...
                        </span>
                      ) : (
                        <button
                          key={item}
                          type="button"
                          className={`mr-pagination-page${item === effectivePage ? ' mr-pagination-page--active' : ''}`}
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
                    className="mr-pagination-nav"
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
                    className="mr-pagination-nav"
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
          </section>
        </div>
      </div>
    </div>
  );
};

export default Mainreport;
