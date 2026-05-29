import React, { useMemo, useState, useEffect, useCallback, useRef } from 'react';
import { useNavigate, useSearchParams } from 'react-router-dom';
import { jsPDF } from 'jspdf';
import JSZip from 'jszip';
import './Mainreport.css';

const API = '/server/mainreport_function/mainreport';

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

const CONSOLIDATED_SECTOR_FOLDERS = ['CLRA', 'Factories Act', 'Shops and Establishment', 'Others'];

const getReturnedSectorFolder = (row) => {
  const sector = String(row?.sector || '').trim();
  if (sector) return sector;
  const act = String(row?.act || '').toLowerCase();
  const isFactories = act.includes('factories') || act.includes('factory');
  if (isFactories) return 'Factories Act';
  const isShops =
    act.includes('shops and establishment') || act.includes('shop and establishment');
  if (isShops) return 'Shops and Establishment';
  const isClra =
    act.includes('clra') || act.includes('contract labour') || act.includes('contract labor');
  if (isClra) return 'CLRA';
  return 'Others';
};

const getReturnedStateFolder = (row) => {
  const s = String(row?.state || '')
    .trim()
    .replace(/[\\/:*?"<>|]/g, '_')
    .replace(/\s+/g, ' ')
    .trim();
  return s.slice(0, 80) || 'Unknown';
};

const groupReturnedRowsByStateAndSector = (rows) => {
  const byState = new Map();
  for (const row of rows) {
    const stateKey = getReturnedStateFolder(row);
    const sectorKey = getReturnedSectorFolder(row);
    if (!byState.has(stateKey)) byState.set(stateKey, new Map());
    const bySector = byState.get(stateKey);
    if (!bySector.has(sectorKey)) bySector.set(sectorKey, []);
    bySector.get(sectorKey).push(row);
  }
  return byState;
};

const formatMaybeDate = (value) => {
  if (!value) return '';
  const d = new Date(value);
  if (Number.isNaN(d.getTime())) return String(value);
  return d.toLocaleString(undefined, { year: 'numeric', month: 'short', day: '2-digit' });
};

const formatPdfExportDate = (date = new Date()) =>
  date.toLocaleDateString(undefined, { year: 'numeric', month: 'short', day: '2-digit' });

const buildReturnedReportPdfArrayBuffer = (rows, { title, logoDataUrl, totalReturned }) => {
  const exportDate = new Date();
  const list = Array.isArray(rows) ? rows : [];
  const doc = new jsPDF({ unit: 'pt', format: 'a4', orientation: 'landscape' });
  const pageWidth = doc.internal.pageSize.getWidth();
  const pageHeight = doc.internal.pageSize.getHeight();
  const startX = 32;
  const tableX = startX;
  const tableWidth = pageWidth - startX * 2;
  const rowH = 18;
  const headerDark = [55, 65, 81];
  let tableY = 0;

  const header = ['S.No', 'Site', 'State', 'Sector', 'Form Name', 'Act', 'Description', 'Due Date', 'Status', 'Remarks'];
  const colPct = [0.05, 0.09, 0.08, 0.1, 0.09, 0.12, 0.2, 0.08, 0.08, 0.11];
  const colW = colPct.map((p) => Math.floor(tableWidth * p));
  const colWidthTotal = colW.reduce((a, b) => a + b, 0);
  if (colWidthTotal < tableWidth) colW[6] += tableWidth - colWidthTotal;

  const drawPageHeader = () => {
    if (logoDataUrl) {
      try {
        const logoFormat = logoDataUrl.includes('image/png')
          ? 'PNG'
          : logoDataUrl.includes('image/webp')
            ? 'WEBP'
            : 'JPEG';
        doc.addImage(logoDataUrl, logoFormat, startX, 16, 74, 26);
      } catch (_) {
        /* optional logo */
      }
    }
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(16);
    doc.setTextColor(17, 24, 39);
    doc.text(title, pageWidth / 2, 36, { align: 'center' });

    const boxW = 180;
    const boxH = 56;
    const y = 58;
    doc.setDrawColor(34, 197, 94);
    doc.setFillColor(255, 255, 255);
    doc.roundedRect(startX, y, boxW, boxH, 8, 8, 'FD');
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(11);
    doc.setTextColor(31, 41, 55);
    doc.text('Total Returned', startX + 12, y + 20);
    doc.setFontSize(20);
    doc.setTextColor(34, 197, 94);
    doc.text(String(totalReturned), startX + 12, y + 44);
    return y + boxH + 14;
  };

  const getCellLines = (value, colIndex) => {
    const text = String(value ?? '');
    if (colIndex === 0) return [text];
    const wrapped = doc.splitTextToSize(text || '-', colW[colIndex] - 8);
    return Array.isArray(wrapped) && wrapped.length > 0 ? wrapped : ['-'];
  };

  const drawTableHeader = () => {
    doc.setFillColor(...headerDark);
    doc.rect(tableX, tableY, tableWidth, rowH, 'F');
    doc.setLineWidth(1.1);
    doc.setDrawColor(120, 120, 120);
    doc.rect(tableX, tableY, tableWidth, rowH);
    let x = tableX;
    let dividerX = tableX;
    for (let i = 0; i < colW.length - 1; i += 1) {
      dividerX += colW[i];
      doc.line(dividerX, tableY, dividerX, tableY + rowH);
    }
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(8);
    doc.setTextColor(255, 255, 255);
    header.forEach((h, i) => {
      doc.text(h, x + 4, tableY + 12);
      x += colW[i];
    });
    tableY += rowH;
  };

  tableY = drawPageHeader();
  drawTableHeader();

  const pdfRows = list.length
    ? list.map((row, idx) => [
        String(idx + 1),
        String(row?.site || '').trim() || '—',
        String(row?.state || '').trim() || '—',
        String(row?.sector || '').trim() || '—',
        String(row?.formName || '').trim() || '—',
        String(row?.act || '').trim() || '—',
        String(row?.description || '').trim() || '—',
        String(row?.dueDate || '').trim() || '—',
        String(row?.status || 'Returned').trim() || 'Returned',
        String(row?.remarks || '').trim() || '—'
      ])
    : [['', '', '', '', '', '', 'No returned forms for this selection.', '', '', '']];

  pdfRows.forEach((r) => {
    const lineHeight = 10;
    const rowPaddingTop = 11;
    const cellLinesByCol = r.map((cell, i) => getCellLines(cell, i));
    const maxLines = Math.max(...cellLinesByCol.map((lines) => lines.length));
    const dynamicRowH = Math.max(rowH, rowPaddingTop + (maxLines - 1) * lineHeight + 6);

    if (tableY + dynamicRowH > pageHeight - 28) {
      doc.addPage();
      tableY = drawPageHeader();
      drawTableHeader();
    }

    let x = tableX;
    doc.setLineWidth(1);
    doc.setDrawColor(130, 130, 130);
    doc.rect(tableX, tableY, tableWidth, dynamicRowH);
    let dividerX = tableX;
    for (let i = 0; i < colW.length - 1; i += 1) {
      dividerX += colW[i];
      doc.line(dividerX, tableY, dividerX, tableY + dynamicRowH);
    }
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(7);
    r.forEach((cell, i) => {
      if (i === 8) doc.setTextColor(185, 28, 28);
      else doc.setTextColor(55, 65, 81);
      cellLinesByCol[i].forEach((line, lineIdx) => {
        doc.text(String(line), x + 4, tableY + rowPaddingTop + lineIdx * lineHeight, {
          maxWidth: colW[i] - 8
        });
      });
      x += colW[i];
    });
    tableY += dynamicRowH;
  });

  const totalPdfPages = doc.getNumberOfPages();
  const exportDateText = `Export Date: ${formatPdfExportDate(exportDate)}`;
  for (let p = 1; p <= totalPdfPages; p += 1) {
    doc.setPage(p);
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(9);
    doc.setTextColor(100, 116, 139);
    doc.text(exportDateText, pageWidth - startX, 30, { align: 'right' });
    doc.text(`Page ${p} of ${totalPdfPages}`, pageWidth - 32, pageHeight - 14, { align: 'right' });
  }

  return doc.output('arraybuffer');
};

const ReturnedReport = ({ userEmail: userEmailProp }) => {
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
  const [consolidatedZipLoading, setConsolidatedZipLoading] = useState(false);
  const selectedYear = yearFromUrl ? String(yearFromUrl).trim() : '';
  const selectedMonth = monthFromUrl ? String(monthFromUrl).trim() : '';
  const selectedSite = siteFromUrl ? String(siteFromUrl).trim() : '';

  const [draftYear, setDraftYear] = useState(selectedYear);
  const [draftMonth, setDraftMonth] = useState(selectedMonth);
  const [draftSite, setDraftSite] = useState(selectedSite);

  const [currentPage, setCurrentPage] = useState(1);
  const PAGE_SIZE = 10;

  const calendarYearStr = useMemo(() => String(new Date().getFullYear()), []);

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

  const hasAppliedFilters = Boolean(selectedYear && selectedMonth);

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

  const downloadReturnedPdf = useCallback(async () => {
    if (!hasAppliedFilters) {
      setError('Apply filters before downloading the report.');
      return;
    }
    if (!tableRows.length) {
      setError('No returned forms to download for this selection.');
      return;
    }
    try {
      const logoDataUrl = await getReportLogoDataUrl();
      const siteSuffix = selectedSite ? ` (${selectedSite})` : ' (All sites)';
      const title = `Returned Report - ${selectedMonth} ${selectedYear}${siteSuffix}`;
      const buf = buildReturnedReportPdfArrayBuffer(tableRows, {
        title,
        logoDataUrl,
        totalReturned: tableRows.length
      });
      const fname = `Returned_Report_${selectedYear}_${selectedMonth}_${selectedSite || 'All'}.pdf`;
      const blob = new Blob([buf], { type: 'application/pdf' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = fname.replace(/\s+/g, '_');
      a.rel = 'noopener';
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
    } catch (e) {
      setError(e.message || 'Failed to download returned report PDF.');
    }
  }, [hasAppliedFilters, selectedYear, selectedMonth, selectedSite, tableRows, getReportLogoDataUrl]);

  const downloadConsolidatedZip = useCallback(async () => {
    if (!hasAppliedFilters || !tableRows.length) {
      setError('Apply filters and load returned forms before downloading.');
      return;
    }
    setConsolidatedZipLoading(true);
    setError('');
    try {
      const logoDataUrl = await getReportLogoDataUrl();
      const siteSuffix = selectedSite ? ` (${selectedSite})` : ' (All sites)';
      const periodLabel = `${selectedMonth} ${selectedYear}${siteSuffix}`;

      const byState = groupReturnedRowsByStateAndSector(tableRows);
      const zip = new JSZip();
      let foldersAdded = 0;

      const statesSorted = [...byState.keys()].sort((a, b) =>
        a.localeCompare(b, undefined, { sensitivity: 'base' })
      );

      for (const stateFolder of statesSorted) {
        const bySector = byState.get(stateFolder);
        if (!bySector) continue;

        const sectorOrder = [
          ...CONSOLIDATED_SECTOR_FOLDERS,
          ...[...bySector.keys()].filter((k) => !CONSOLIDATED_SECTOR_FOLDERS.includes(k))
        ];

        for (const sectorFolder of sectorOrder) {
          const sectorRows = bySector.get(sectorFolder);
          if (!sectorRows?.length) continue;

          const stateDir = zip.folder(stateFolder);
          const sectorDir = stateDir?.folder(sectorFolder);
          if (!sectorDir) continue;

          const pdfTitle = `Returned Report — ${stateFolder} — ${sectorFolder} - ${periodLabel}`;
          const pdfBuf = buildReturnedReportPdfArrayBuffer(sectorRows, {
            title: pdfTitle,
            logoDataUrl,
            totalReturned: sectorRows.length
          });
          sectorDir.file(`${sectorFolder.replace(/\s+/g, '_')}_Returned_Report.pdf`, pdfBuf, {
            binary: true
          });
          foldersAdded += 1;
        }
      }

      if (!foldersAdded) throw new Error('No state/sector data found for the current filters.');

      const monthLabel = String(selectedMonth || '').replace(/\s+/g, '_');
      const siteLabel = String(selectedSite || 'All').replace(/\s+/g, '_');
      const fname = `Returned_Consolidated_Report_${selectedYear}_${monthLabel}_${siteLabel}.zip`;
      const blob = await zip.generateAsync({ type: 'blob' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = fname;
      a.rel = 'noopener';
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
    } catch (e) {
      setError(e.message || 'Failed to download consolidated returned report ZIP.');
    } finally {
      setConsolidatedZipLoading(false);
    }
  }, [
    hasAppliedFilters,
    selectedYear,
    selectedMonth,
    selectedSite,
    tableRows,
    getReportLogoDataUrl
  ]);

  const loadSummary = useCallback(async () => {
    setError('');
    setLoading(true);
    try {
      const res = await fetch(`${API}/returned-summary`, { cache: 'no-store' });
      const json = await res.json();
      if (json.status !== 'success' || !json.data) {
        throw new Error(json.message || 'Invalid response');
      }
      setYears(Array.isArray(json.data.years) ? json.data.years : []);
    } catch (e) {
      setError(e.message || 'Failed to load returned report');
      setYears([]);
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    loadSummary();
  }, [loadSummary]);

  const loadEntries = useCallback(
    async (year, month, siteValue = '') => {
      setTableLoading(true);
      setError('');
      try {
        const q = new URLSearchParams({
          year: String(year),
          month: String(month),
          userEmail: effectiveUserEmail,
          site: String(siteValue || '')
        });
        const res = await fetch(`${API}/returned-entries?${q}`, { cache: 'no-store' });
        const json = await res.json();
        if (json.status !== 'success' || !json.data) {
          throw new Error(json.message || 'Invalid response');
        }
        setTableRows(Array.isArray(json.data.rows) ? json.data.rows : []);
        setSiteNames(Array.isArray(json.data.siteNames) ? json.data.siteNames : []);
      } catch (e) {
        setError(e.message || 'Failed to load returned rows');
        setTableRows([]);
        setSiteNames([]);
      } finally {
        setTableLoading(false);
      }
    },
    [effectiveUserEmail]
  );

  useEffect(() => {
    const hasApplied = Boolean(selectedYear && selectedMonth);
    const y = hasApplied ? selectedYear : draftYear;
    const m = hasApplied ? selectedMonth : draftMonth;
    const site = hasApplied ? String(selectedSite || '').trim() : '';
    if (!y || !m) {
      if (!hasApplied) setSiteNames([]);
      return;
    }
    loadEntries(y, m, site);
  }, [selectedYear, selectedMonth, selectedSite, draftYear, draftMonth, loadEntries]);

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
    navigate(`/returned-report?${q.toString()}`);
  }, [draftYear, draftMonth, draftSite, navigate]);

  const resetFilters = useCallback(() => {
    setDraftYear('');
    setDraftMonth('');
    setDraftSite('');
    navigate('/returned-report');
    setTableRows([]);
    setSiteNames([]);
  }, [navigate]);

  const totalReturned = useMemo(() => (Array.isArray(tableRows) ? tableRows.length : 0), [tableRows]);

  const totalPages = Math.max(1, Math.ceil(tableRows.length / PAGE_SIZE));
  const effectivePage = Math.min(currentPage, totalPages);
  const pagedRows = tableRows.slice((effectivePage - 1) * PAGE_SIZE, effectivePage * PAGE_SIZE);
  const paginationItems = useMemo(
    () => buildPaginationItems(effectivePage, totalPages),
    [effectivePage, totalPages]
  );

  useEffect(() => {
    setCurrentPage(1);
  }, [tableRows, selectedYear, selectedMonth, selectedSite]);

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
                <h1 className="mr-title">Returned Report</h1>
              </div>
            </div>
          </header>

          {error ? <div className="mr-alert">{error}</div> : null}
          {loading ? <div className="mr-muted" style={{ padding: '8px 0' }}>Loading filters…</div> : null}

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
                    name="returned-report-year"
                    value={draftYear}
                    required
                    aria-required="true"
                    onChange={(e) => {
                      setDraftYear(e.target.value);
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
                    name="returned-report-month"
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
                      <option key={s} value={s}>
                        {s}
                      </option>
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
                  >
                    Apply Filters
                  </button>
                </div>
              </div>
            </form>
          </section>

          <section className="mr-kpis" aria-label="Summary">
            <div className="mr-kpi mr-kpi--forms">
              <div className="mr-kpi-label">Total Returned</div>
              <div className="mr-kpi-value">{totalReturned}</div>
            </div>
          </section>

          <section className="mr-grid">
            <div className="mr-left">
              <section className="mr-card mr-table-card" aria-label="Returned forms">
                <div className="mr-table-head">
                  <div>
                    <div className="mr-table-title">Returned Forms</div>
                  </div>
                  <div className="mr-table-actions">
                    <button
                      type="button"
                      className="mr-seg mr-seg--active"
                      onClick={downloadConsolidatedZip}
                      disabled={!hasAppliedFilters || consolidatedZipLoading || !tableRows.length}
                      title={
                        !hasAppliedFilters
                          ? 'Apply filters first'
                          : consolidatedZipLoading
                            ? 'Preparing ZIP…'
                            : !tableRows.length
                              ? 'No data to download'
                              : 'ZIP by State → Sector with returned forms PDF in each sector folder'
                      }
                    >
                      {consolidatedZipLoading ? 'Preparing ZIP…' : 'Consolidated Report'}
                    </button>
                    <button
                      type="button"
                      className="mr-btn mr-btn-primary"
                      onClick={downloadReturnedPdf}
                      disabled={!hasAppliedFilters || !tableRows.length}
                      title={
                        !hasAppliedFilters
                          ? 'Apply filters first'
                          : !tableRows.length
                            ? 'No data to download'
                            : 'Download returned forms as PDF'
                      }
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
                        <th style={{ width: 140 }}>Site</th>
                        <th style={{ width: 100, minWidth: 100 }}>State</th>
                        <th style={{ width: 120, minWidth: 120 }}>Sector</th>
                        <th style={{ maxWidth: 160, width: 160 }}>Form Name</th>
                        <th style={{ maxWidth: 200, width: 200 }}>Act</th>
                        <th style={{ maxWidth: 280, width: 280 }}>Description</th>
                        <th style={{ width: 120 }}>Due Date</th>
                        <th style={{ width: 100 }}>Status</th>
                        <th>Remarks</th>
                      </tr>
                    </thead>
                    <tbody>
                      {!hasAppliedFilters ? (
                        <tr>
                          <td colSpan={10} className="mr-empty">
                            Select year, month, and site, then click Apply Filters to load returned forms.
                          </td>
                        </tr>
                      ) : tableLoading && tableRows.length === 0 ? (
                        <tr aria-hidden="true">
                          <td colSpan={10} className="mr-table-placeholder" />
                        </tr>
                      ) : tableRows.length === 0 ? (
                        <tr>
                          <td colSpan={10} className="mr-empty">
                            No returned forms for this selection.
                          </td>
                        </tr>
                      ) : (
                        pagedRows.map((row, idx) => (
                            <tr key={row?.rowId != null ? String(row.rowId) : `${effectivePage}-${idx}`}>
                              <td>{(effectivePage - 1) * PAGE_SIZE + idx + 1}</td>
                              <td title={row?.site || ''}>{row?.site || '—'}</td>
                              <td title={row?.state || ''}>{row?.state || '—'}</td>
                              <td title={row?.sector || ''}>{row?.sector || '—'}</td>
                              <td
                                className="mr-strong"
                                style={{ maxWidth: 160, whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}
                                title={row?.formName || ''}
                              >
                                {row?.formName || ''}
                              </td>
                              <td
                                className="mr-act-col"
                                title={row?.act || ''}
                              >
                                {row?.act || ''}
                              </td>
                              <td
                                style={{
                                  maxWidth: 280,
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
                              <td>{row?.dueDate || '—'}</td>
                              <td>
                                <span className="mr-status mr-status--returned">
                                  {row?.status || 'Returned'}
                                </span>
                              </td>
                              <td title={row?.remarks || ''}>{row?.remarks || '—'}</td>
                            </tr>
                          ))
                      )}
                    </tbody>
                  </table>
                </div>

                {tableRows.length > PAGE_SIZE ? (
                  <nav className="mr-pagination" aria-label="Table pagination">
                    <button
                      type="button"
                      className="mr-pagination-nav"
                      disabled={effectivePage <= 1}
                      onClick={() => setCurrentPage(1)}
                      aria-label="First page"
                    >
                      «
                    </button>
                    <button
                      type="button"
                      className="mr-pagination-nav"
                      disabled={effectivePage <= 1}
                      onClick={() => setCurrentPage((p) => Math.max(1, p - 1))}
                      aria-label="Previous page"
                    >
                      ‹
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
                      aria-label="Next page"
                    >
                      ›
                    </button>
                    <button
                      type="button"
                      className="mr-pagination-nav"
                      disabled={effectivePage >= totalPages}
                      onClick={() => setCurrentPage(totalPages)}
                      aria-label="Last page"
                    >
                      »
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

export default ReturnedReport;
