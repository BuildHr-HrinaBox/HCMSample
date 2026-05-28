import React, { useMemo, useState, useEffect, useCallback, useRef } from 'react';
import { Link, useNavigate, useSearchParams } from 'react-router-dom';
import { jsPDF } from 'jspdf';
import './Mainreport.css';

const API = '/server/mainreport_function/mainreport';

/** Same key as Statutory.js — Reports "Statutory" link primes Transaction month filter. */
const STATUTORY_MONTH_FILTER_KEY = 'statutory_month_filter';

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

/** Same rule as Statutory STATUS + `mainreport_function`: only explicit transaction status Approved may expose drafts. */
const getStatutoryTransactionStatusNorm = (row) =>
  String(row?.statutoryStatus ?? row?.status ?? '')
    .trim()
    .toLowerCase();

const isStatutoryTransactionStatusApproved = (row) => {
  const s = getStatutoryTransactionStatusNorm(row);
  return s === 'approved' || s === 'approve';
};

/** Strip draft URL/name if API is stale or fields disagree (prevents Pending + Excel link). */
const sanitizeMainReportApiRow = (row) => {
  if (!row || typeof row !== 'object') return row;
  if (isStatutoryTransactionStatusApproved(row)) return row;
  return { ...row, draftFileUrl: null, draftFile: null, draftFileName: '' };
};

const canShowDraftFileLink = (row) => isStatutoryTransactionStatusApproved(row) && !!row?.draftFileUrl;

const isYetToCompleteStatusText = (stLower) =>
  stLower.includes('yet to complete') ||
  (stLower.includes('yet') && stLower.includes('complete')) ||
  stLower.includes('not started') ||
  stLower.includes('nostart') ||
  stLower.includes('incomplete') ||
  stLower.includes('not complete');

const getStatusLabel = (row) => {
  const st = String(row?.statutoryStatus || row?.status || '').trim();
  const stLower = st.toLowerCase();
  const appr = String(row?.approval || '').trim().toLowerCase();
  const hasDraftStored = row?.hasStatutoryDraftStored === true;

  if (appr === 'rejected' || appr === 'reject' || stLower.includes('reject')) return 'Rejected';

  if (stLower === 'approved' || stLower === 'approve') return 'Approved';

  if (hasDraftStored && (st === '' || st === '-' || st === '—')) return 'Pending';
  if (hasDraftStored && stLower === 'pending') return 'Pending';

  if (isYetToCompleteStatusText(stLower)) return 'Yet to Complete';

  if (!hasDraftStored) return 'Yet to Complete';

  /** Never surface raw DB typos / stray values (e.g. "change") in Reports — bucket as Yet to Complete. */
  return 'Yet to Complete';
};

const getSectorGroupLabel = (row) => {
  const s = String(row?.sector || '').trim();
  return s || 'Others';
};

const sortRowsSectorWise = (rows) =>
  [...(Array.isArray(rows) ? rows : [])].sort((a, b) => {
    const sa = getSectorGroupLabel(a).toLowerCase();
    const sb = getSectorGroupLabel(b).toLowerCase();
    if (sa !== sb) return sa.localeCompare(sb);
    return String(a?.formName || '').localeCompare(String(b?.formName || ''), undefined, { sensitivity: 'base' });
  });

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
  const siteFromUrl = (searchParams.get('site') || '').trim();
  const yearFromUrl = (searchParams.get('year') || '').trim();
  const monthFromUrl = (searchParams.get('month') || '').trim();

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
  const selectedYear = yearFromUrl ? String(yearFromUrl).trim() : '';
  const selectedMonth = monthFromUrl ? String(monthFromUrl).trim() : '';
  const selectedSite = siteFromUrl ? String(siteFromUrl).trim() : '';

  // Filter UI state (Apply/Reset like screenshot)
  const [draftYear, setDraftYear] = useState(selectedYear);
  const [draftMonth, setDraftMonth] = useState(selectedMonth);
  const [draftSite, setDraftSite] = useState(selectedSite);
  const [showDraftOnly, setShowDraftOnly] = useState(false);

  const reportSiteForStatutory = (selectedSite || draftSite || '').trim();
  const reportMonthForStatutory = (selectedMonth || draftMonth || '').trim();

  const statutoryTransactionHref = useMemo(() => {
    if (!reportSiteForStatutory) return '/rule-book/statutory';
    return `/rule-book/statutory?site=${encodeURIComponent(reportSiteForStatutory)}`;
  }, [reportSiteForStatutory]);

  const primeStatutoryMonthFilter = useCallback((fullMonthName) => {
    const m = String(fullMonthName || '').trim();
    if (m && typeof localStorage !== 'undefined') {
      localStorage.setItem(STATUTORY_MONTH_FILTER_KEY, m);
    }
  }, []);
  // Pagination state
  const [currentPage, setCurrentPage] = useState(1);
  const PAGE_SIZE = 10;

  const calendarYearStr = useMemo(() => String(new Date().getFullYear()), []);

  /** Always include URL/draft year + current calendar year so the <select> value matches an <option> before /summary returns (avoids year “jumping”). */
  const yearSelectOptions = useMemo(() => {
    const uniq = new Set();
    (Array.isArray(years) ? years : []).forEach((y) => {
      const s = String(y != null ? y : '').trim();
      if (s) uniq.add(s);
    });
    [calendarYearStr, selectedYear, draftYear].forEach((y) => {
      const s = String(y != null ? y : '').trim();
      if (s) uniq.add(s);
    });
    return [...uniq].sort((a, b) => {
      const na = Number.parseInt(a, 10);
      const nb = Number.parseInt(b, 10);
      if (Number.isFinite(na) && Number.isFinite(nb)) return nb - na;
      return String(b).localeCompare(String(a), undefined, { sensitivity: 'base' });
    });
  }, [years, calendarYearStr, selectedYear, draftYear]);

  const lastUrlFilterKeyRef = useRef('');
  const reportLogoDataUrlRef = useRef(null);

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
      const rawRows = Array.isArray(json.data.rows) ? json.data.rows : [];
      setTableRows(rawRows.map(sanitizeMainReportApiRow));
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
    const key = `${selectedYear}|${selectedMonth}|${selectedSite}`;
    if (lastUrlFilterKeyRef.current === key) return;
    lastUrlFilterKeyRef.current = key;
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
    setSiteNames([]);
  };

  const getReportLogoDataUrl = useCallback(async () => {
    if (reportLogoDataUrlRef.current) return reportLogoDataUrlRef.current;
    try {
      const res = await fetch('/server/settings_function/settings/logo');
      if (!res.ok) return null;
      const blob = await res.blob();
      if (!blob || !blob.type.startsWith('image/')) return null;
      const dataUrl = await new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onloadend = () => resolve(String(reader.result || ''));
        reader.onerror = reject;
        reader.readAsDataURL(blob);
      });
      reportLogoDataUrlRef.current = dataUrl || null;
      return reportLogoDataUrlRef.current;
    } catch (_) {
      return null;
    }
  }, []);

  const downloadConsolidatedPdf = async () => {
    const doc = new jsPDF({ unit: 'pt', format: 'a4', orientation: 'landscape' });
    const pageWidth = doc.internal.pageSize.getWidth();
    const pageHeight = doc.internal.pageSize.getHeight();
    const logoDataUrl = await getReportLogoDataUrl();

    const title = `Consolidated Report - ${selectedMonth || draftMonth || ''} ${selectedYear || draftYear || ''}${(selectedSite || draftSite) ? ` (${selectedSite || draftSite})` : ''}`;

    const totalForms = tableRows.length;
    let approvedPdf = 0;
    for (const row of tableRows) {
      if (isStatutoryTransactionStatusApproved(row) && getStatusLabel(row) === 'Approved') approvedPdf += 1;
    }
    let pendingPdf = 0;
    let yetToCompletePdf = 0;
    for (const row of tableRows) {
      const status = getStatusLabel(row);
      if (status === 'Pending') pendingPdf += 1;
      else if (status === 'Yet to Complete') yetToCompletePdf += 1;
    }

    const boxes = [
      { label: 'Total Forms', count: totalForms, color: [34, 197, 94] },
      { label: 'Approved', count: approvedPdf, color: [99, 102, 241] },
      { label: 'Pending', count: pendingPdf, color: [245, 158, 11] },
      { label: 'Yet to Complete', count: yetToCompletePdf, color: [234, 179, 8] }
    ];

    const startX = 32;
    const gap = 12;
    const boxW = (pageWidth - startX * 2 - gap * 3) / 4;
    const boxH = 72;

    const header = ['S.NO', 'Sector', 'Form Name', 'Act', 'Description', 'Month Filter', 'Status', 'Draft File'];
    const colW = [28, 72, 112, 136, 200, 66, 82, 74];
    const rowH = 18;
    const tableX = 32;
    let tableY = 0;

    const getCellLines = (value, colIndex) => {
      const text = String(value ?? '');
      // Wrap all columns except serial number so row height accounts for line breaks.
      if (colIndex === 0) return [text];
      const wrapped = doc.splitTextToSize(text || '-', colW[colIndex] - 10);
      return Array.isArray(wrapped) && wrapped.length > 0 ? wrapped : ['-'];
    };

    const drawPageHeader = () => {
      let titleX = 32;
      const titleY = 36;
      if (logoDataUrl) {
        try {
          const logoX = 32;
          const logoY = 16;
          const logoW = 74;
          const logoH = 26;
          const logoFormat = logoDataUrl.includes('image/png')
            ? 'PNG'
            : logoDataUrl.includes('image/webp')
              ? 'WEBP'
              : 'JPEG';
          doc.addImage(logoDataUrl, logoFormat, logoX, logoY, logoW, logoH);
          titleX = logoX + logoW + 56;
        } catch (_) {
          titleX = 32;
        }
      }
      doc.setFont('helvetica', 'bold');
      doc.setFontSize(16);
      doc.setTextColor(17, 24, 39);
      doc.text(title, titleX, titleY);

      let y = 64;
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
      doc.text(`Month Data: ${selectedMonth || draftMonth || ''} ${selectedYear || draftYear || ''}`, 32, y);
      return y + 16;
    };

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
      doc.setFontSize(9);
      doc.setTextColor(55, 65, 81);
      header.forEach((h, i) => {
        doc.text(h, x + 6, tableY + 12);
        x += colW[i];
      });
      tableY += rowH;
    };

    tableY = drawPageHeader();
    drawHeader();

    const sortedPdfRows = sortRowsSectorWise(tableRows);
    const rows = sortedPdfRows.length
      ? sortedPdfRows.map((row, idx) => [
          String(idx + 1),
          String(row.sector || '—'),
          String(row.formName || '-'),
          String(row.act || '-'),
          String(row.description || '-'),
          String(row.monthFilter || row.MonthFilter || selectedMonth || draftMonth || '—'),
          getStatusLabel(row),
          String(canShowDraftFileLink(row) && row.draftFileName ? row.draftFileName : '-')
        ])
      : [['', '', '', '', 'No rows for this month.', '', '', '']];

    rows.forEach((r) => {
      const lineHeight = 10;
      const rowPaddingTop = 12;
      const rowPaddingBottom = 4;
      const cellLinesByCol = r.map((cell, i) => getCellLines(cell, i));
      const maxLines = Math.max(...cellLinesByCol.map((lines) => lines.length));
      const dynamicRowH = Math.max(rowH, rowPaddingTop + (maxLines - 1) * lineHeight + rowPaddingBottom);

      if (tableY + dynamicRowH > pageHeight - 28) {
        doc.addPage();
        tableY = drawPageHeader();
        drawHeader();
      }
      let x = tableX;
      doc.setLineWidth(1.1);
      doc.setDrawColor(130, 130, 130);
      doc.rect(tableX, tableY, colW.reduce((a, b) => a + b, 0), dynamicRowH);
      // Draw vertical separators for body columns
      let rowDividerX = tableX;
      for (let i = 0; i < colW.length - 1; i += 1) {
        rowDividerX += colW[i];
        doc.line(rowDividerX, tableY, rowDividerX, tableY + dynamicRowH);
      }
      doc.setFont('helvetica', 'normal');
      doc.setFontSize(8);
      r.forEach((cell, i) => {
        if (i === 6) {
          const statusLower = String(cell || '').toLowerCase();
          if (statusLower.includes('approved')) {
            doc.setTextColor(22, 163, 74);
          } else if (statusLower.includes('yet')) doc.setTextColor(234, 179, 8);
          else if (statusLower.includes('pending')) doc.setTextColor(217, 119, 6);
          else doc.setTextColor(75, 85, 99);
        } else {
          doc.setTextColor(75, 85, 99);
        }
        const lines = cellLinesByCol[i];
        lines.forEach((line, lineIdx) => {
          doc.text(String(line), x + 6, tableY + rowPaddingTop + lineIdx * lineHeight, {
            maxWidth: colW[i] - 10,
          });
        });
        x += colW[i];
      });
      tableY += dynamicRowH;
    });

    const totalPdfPages = doc.getNumberOfPages();
    for (let p = 1; p <= totalPdfPages; p += 1) {
      doc.setPage(p);
      doc.setFont('helvetica', 'normal');
      doc.setFontSize(9);
      doc.setTextColor(100, 116, 139);
      doc.text(`Page ${p} of ${totalPdfPages}`, pageWidth - 32, pageHeight - 14, { align: 'right' });
    }

    const fname = `Consolidated_Report_${selectedYear || draftYear || ''}_${selectedMonth || draftMonth || ''}_${(selectedSite || draftSite) || 'All'}.pdf`;
    doc.save(fname.replace(/\s+/g, '_'));
  };

  /**
   * KPIs match the Forms Summary STATUS column (`getStatusLabel`).
   */
  const statusCounts = useMemo(() => {
    if (!Array.isArray(tableRows) || tableRows.length === 0) {
      return { total: 0, approved: 0, pending: 0, yetToComplete: 0 };
    }
    let approved = 0;
    let pending = 0;
    let yetToComplete = 0;
    for (const row of tableRows) {
      const status = getStatusLabel(row);
      if (isStatutoryTransactionStatusApproved(row) && status === 'Approved') approved += 1;
      else if (status === 'Pending') pending += 1;
      else if (status === 'Yet to Complete') yetToComplete += 1;
    }
    const total = tableRows.length;
    return { total, approved, pending, yetToComplete };
  }, [tableRows]);

  const filteredRows = useMemo(() => {
    const rows = !showDraftOnly
      ? tableRows
      : (Array.isArray(tableRows) ? tableRows : []).filter((r) => canShowDraftFileLink(r));
    return sortRowsSectorWise(rows);
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

  const draftHasYearMonth = Boolean(draftYear && draftMonth);

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
                {yearSelectOptions.map((y) => (
                  <option key={y} value={String(y)}>
                    {y}
                  </option>
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
          <div className="mr-filter-hint" />
          </section>

          <section className="mr-kpis" aria-label="Summary">
          <div className="mr-kpi mr-kpi--forms">
            <div className="mr-kpi-label">Total Forms</div>
            <div className="mr-kpi-value">{statusCounts.total}</div>
          </div>
          <div className="mr-kpi mr-kpi--drafts">
            <div className="mr-kpi-label">Approved</div>
            <div className="mr-kpi-value">{statusCounts.approved}</div>
          </div>
          <div
            className="mr-kpi mr-kpi--pending"
            title="Forms that are in Pending status."
          >
            <div className="mr-kpi-label">Pending</div>
            <div className="mr-kpi-value">{statusCounts.pending}</div>
          </div>
          <div className="mr-kpi mr-kpi--yet" title="Forms that are Yet to Complete.">
            <div className="mr-kpi-label">Yet to Complete</div>
            <div className="mr-kpi-value">{statusCounts.yetToComplete}</div>
          </div>
          </section>


          <section className="mr-grid">
          <div className="mr-left">
            <section className="mr-card mr-table-card" aria-label="Forms Summary">
              <div className="mr-table-head">
                <div>
                  <div className="mr-table-title">Forms Summary</div>
                </div>
                <div className="mr-table-actions">
                  <button
                    type="button"
                    className={`mr-seg${!showDraftOnly ? ' mr-seg--active' : ''}`}
                    onClick={() => setShowDraftOnly(false)}
                    disabled={!draftHasYearMonth}
                  >
                    Consolidated Report
                  </button>
                  <button
                    type="button"
                    className={`mr-seg${showDraftOnly ? ' mr-seg--active' : ''}`}
                    onClick={() => setShowDraftOnly(true)}
                    disabled={!draftHasYearMonth}
                  >
                    Draft Files Only
                  </button>
                  <button
                    type="button"
                    className="mr-btn mr-btn-primary"
                    onClick={downloadConsolidatedPdf}
                    disabled={!draftHasYearMonth}
                    title={!draftHasYearMonth ? 'Select Year/Month first' : 'Download PDF'}
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
                      <th style={{ width: 180, minWidth: 180 }}>Sector</th>
                      <th style={{ maxWidth: 190, width: 190 }}>Form Name</th>
                      <th style={{ maxWidth: 260, width: 260 }}>Act</th>
                      <th style={{ maxWidth: 360, width: 360 }}>Description</th>
                      <th style={{ width: 120, minWidth: 120 }}>Month Filter</th>
                      <th style={{ width: 180 }}>Status</th>
                      <th>Draft File</th>
                    </tr>
                  </thead>
                  <tbody>
                    {!draftHasYearMonth ? (
                      <tr>
                        <td colSpan={8} className="mr-empty">
                          Select year and month to load the table.
                        </td>
                      </tr>
                    ) : tableLoading && filteredRows.length === 0 ? (
                      <tr aria-hidden="true">
                        <td colSpan={8} className="mr-table-placeholder" />
                      </tr>
                    ) : filteredRows.length === 0 ? (
                      <tr>
                        <td colSpan={8} className="mr-empty">No rows for this selection.</td>
                      </tr>
                    ) : (
                      pagedRows.map((row, idx) => {
                        const seq = (effectivePage - 1) * PAGE_SIZE + idx + 1;
                        const key = row?.rowId != null ? String(row.rowId) : `${(effectivePage - 1) * PAGE_SIZE + idx}`;
                        const status = getStatusLabel(row);
                        const statusClass =
                          status === 'Approved'
                            ? 'mr-status--ok'
                            : status === 'Yet to Complete'
                              ? 'mr-status--yet'
                              : 'mr-status--pending';
                        const draftName = canShowDraftFileLink(row) ? row?.draftFileName || '' : '';
                        const monthCell =
                          String(row?.monthFilter || row?.MonthFilter || reportMonthForStatutory || '').trim() || '—';
                        const monthToPrime = monthCell !== '—' ? monthCell : reportMonthForStatutory;
                        return (
                          <tr key={key}>
                            <td>{seq}</td>
                            <td style={{ width: 180, minWidth: 180, whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>
                              {row?.sector || '—'}
                            </td>
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
                            <td className="mr-month-filter-cell">
                              <span className="mr-month-filter-value" title={monthCell}>
                                {monthCell}
                              </span>
                            </td>
                            <td>
                              <span className={`mr-status ${statusClass}`}>{status}</span>
                            </td>
                            <td>
                              {canShowDraftFileLink(row) ? (
                                <a className="mr-link" href={row.draftFileUrl} target="_blank" rel="noopener noreferrer">
                                  {draftName || 'Open draft'}
                                </a>
                              ) : (
                                <span className="mr-muted">—</span>
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
