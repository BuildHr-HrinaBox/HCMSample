import React, { useMemo, useState, useEffect, useCallback, useRef } from 'react';
import { useNavigate, useSearchParams } from 'react-router-dom';
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

const formatMaybeDate = (value) => {
  if (!value) return '';
  const d = new Date(value);
  if (Number.isNaN(d.getTime())) return String(value);
  return d.toLocaleString(undefined, { year: 'numeric', month: 'short', day: '2-digit' });
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
                </div>

                <div className="mr-table-wrap">
                  <table className="mr-table">
                    <thead>
                      <tr>
                        <th style={{ width: 70 }}>S.No</th>
                        <th style={{ maxWidth: 160, width: 160 }}>Form Name</th>
                        <th style={{ maxWidth: 200, width: 200 }}>Act</th>
                        <th style={{ maxWidth: 280, width: 280 }}>Description</th>
                        <th style={{ width: 110, minWidth: 110 }}>Month</th>
                        <th style={{ width: 140 }}>Site</th>
                        <th style={{ width: 100 }}>Status</th>
                        <th>Remarks</th>
                        <th style={{ width: 120 }}>Due Date</th>
                      </tr>
                    </thead>
                    <tbody>
                      {!draftHasYearMonth ? (
                        <tr>
                          <td colSpan={9} className="mr-empty">
                            Select year and month, then click <strong>Apply Filters</strong>.
                          </td>
                        </tr>
                      ) : tableLoading && tableRows.length === 0 ? (
                        <tr aria-hidden="true">
                          <td colSpan={9} className="mr-table-placeholder" />
                        </tr>
                      ) : tableRows.length === 0 ? (
                        <tr>
                          <td colSpan={9} className="mr-empty">
                            No returned forms for this selection.
                          </td>
                        </tr>
                      ) : (
                        pagedRows.map((row, idx) => {
                          const monthCell =
                            String(
                              row?.monthFilter || selectedMonth || draftMonth || ''
                            ).trim() || '—';
                          return (
                            <tr key={row?.rowId != null ? String(row.rowId) : `${effectivePage}-${idx}`}>
                              <td>{(effectivePage - 1) * PAGE_SIZE + idx + 1}</td>
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
                              <td className="mr-month-filter-cell" title={monthCell}>
                                {monthCell}
                              </td>
                              <td title={row?.site || ''}>{row?.site || '—'}</td>
                              <td>
                                <span className="mr-status mr-status--returned">
                                  {row?.status || 'Returned'}
                                </span>
                              </td>
                              <td title={row?.remarks || ''}>{row?.remarks || '—'}</td>
                              <td>{row?.dueDate || '—'}</td>
                            </tr>
                          );
                        })
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
