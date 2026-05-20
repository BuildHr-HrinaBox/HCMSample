import React, { useEffect, useMemo, useState } from 'react';
import {
  Bell,
  CircleHelp,
  Building,
  ClipboardList,
  AlertTriangle,
  BookMarked,
  CheckCircle2,
  ShieldCheck,
} from 'lucide-react';
import HcmDashboardSidebar from '../components/HcmDashboardSidebar';
import HcmHeaderProfileMenu from '../components/HcmHeaderProfileMenu';
import { Link } from 'react-router-dom';
import './newdashboard.css';

const DEFAULT_METRICS = [
  {
    key: 'companies',
    label: 'Active Companies',
    value: 0,
    icon: Building,
    tone: 'green',
  },
  {
    key: 'due',
    label: 'Approved',
    value: 0,
    icon: CheckCircle2,
    tone: 'green',
  },
  {
    key: 'approvals',
    label: 'Pending Approvals',
    value: 0,
    icon: ClipboardList,
    tone: 'blue',
  },
  {
    key: 'policies',
    label: 'Rejected',
    value: 0,
    icon: BookMarked,
    tone: 'purple',
  },
];

const COMPANY_API = '/server/company_function/company';
const STATUTORY_API = '/server/statutoryreg_function/statutory';
const CHECKLIST_BULK_API = '/server/checklistbulk_function/checklistbulk?action=getAll';

const CAL_WEEKDAYS = ['Su', 'Mo', 'Tu', 'We', 'Th', 'Fr', 'Sa'];

function buildCalendarGrid(year, monthIndex) {
  const first = new Date(year, monthIndex, 1);
  const startPad = first.getDay();
  const daysInMonth = new Date(year, monthIndex + 1, 0).getDate();
  const cells = [];
  for (let i = 0; i < startPad; i += 1) cells.push(null);
  for (let d = 1; d <= daysInMonth; d += 1) cells.push(d);
  const rows = [];
  for (let i = 0; i < cells.length; i += 7) {
    rows.push(cells.slice(i, i + 7));
  }
  return rows;
}

const TRANSACTION_PAGE = '/rule-book/statutory';

const hasValue = (value) => {
  if (value == null) return false;
  const text = String(value).trim().toLowerCase();
  return text !== '' && text !== 'null' && text !== 'undefined';
};

const parseDateValue = (value) => {
  const text = String(value || '').trim();
  if (!text) return null;
  const parsed = Date.parse(text);
  return Number.isNaN(parsed) ? null : parsed;
};

const formatDueDate = (value) => {
  const timestamp = parseDateValue(value);
  if (timestamp == null) return String(value || '-').trim() || '-';
  return new Date(timestamp).toLocaleDateString('en-GB', {
    day: '2-digit',
    month: 'short',
    year: 'numeric',
  });
};

const isMonthlyBasis = (value) => String(value || '').trim().toLowerCase() === 'monthly basis';

const buildMonthlyBasisDate = (calendarMonth) =>
  new Date(calendarMonth.year, calendarMonth.month, 15);

const getDueDateInfo = (value, calendarMonth) => {
  const raw = String(value || '').trim();
  if (isMonthlyBasis(raw)) {
    const date = buildMonthlyBasisDate(calendarMonth);
    return {
      timestamp: date.getTime(),
      display: date.toLocaleDateString('en-GB', { day: '2-digit', month: 'short' }),
      day: 15,
      isDerived: true,
    };
  }

  const timestamp = parseDateValue(raw);
  if (timestamp == null) {
    return {
      timestamp: null,
      display: raw || '-',
      day: null,
      isDerived: false,
    };
  }

  const date = new Date(timestamp);
  return {
    timestamp,
    display: date.toLocaleDateString('en-GB', { day: '2-digit', month: 'short' }),
    day: date.getDate(),
    isDerived: false,
  };
};

const getStatusMeta = (value) => {
  const rawLabel = String(value || '').trim();
  const normalized = rawLabel.toLowerCase();
  if (normalized === 'approved') {
    return { label: 'Approved', statusClass: 'nd-badge-upcoming', iconTone: 'green' };
  }
  if (normalized === 'pending') {
    return { label: 'Pending', statusClass: 'nd-badge-pending', iconTone: 'grey' };
  }
  if (normalized === 'yet to complete') {
    return { label: 'Yet to Complete', statusClass: 'nd-badge-pending', iconTone: 'grey' };
  }
  if (normalized === 'rejected' || normalized === 'reject') {
    return { label: 'Rejected', statusClass: 'nd-badge-high', iconTone: 'orange' };
  }
  if (normalized === 'in progress') {
    return { label: 'In Progress', statusClass: 'nd-badge-progress', iconTone: 'blue' };
  }
  return {
    label: rawLabel || 'Pending',
    statusClass: rawLabel ? 'nd-badge-upcoming' : 'nd-badge-pending',
    iconTone: rawLabel ? 'green' : 'grey',
  };
};

const hasSubmittedFile = (item) =>
  hasValue(item?.formFile ?? item?.FormFile) ||
  hasValue(item?.proofSubmissionFile ?? item?.ProofSubmissionFile) ||
  hasValue(item?.draftFile ?? item?.DraftFile);

const normalizeLookupValue = (value) => String(value || '').trim().toLowerCase();

const buildActDescriptionKey = (act, description) =>
  `${normalizeLookupValue(act)}|${normalizeLookupValue(description)}`;

const buildTaskKey = (item) =>
  `${String(item?.name ?? '').toLowerCase()}|${String(item?.due ?? '').toLowerCase()}|${String(item?.company ?? '').toLowerCase()}`;

const limitUpcomingTasks = (items, limit = 4) => {
  const submitted = [];
  const remaining = [];
  const seen = new Set();

  items.forEach((item) => {
    const key = buildTaskKey(item);
    if (seen.has(key)) return;
    seen.add(key);
    if (item?.hasFile) submitted.push(item);
    else remaining.push(item);
  });

  return [...submitted, ...remaining].slice(0, limit);
};

function NewDashboard({ userName = 'Ravi Kumar', userRole = 'HR Admin', userInitials = 'RA', userEmail }) {
  const [metrics, setMetrics] = useState(DEFAULT_METRICS);
  const [tasks, setTasks] = useState([]);
  const [calendarDueDates, setCalendarDueDates] = useState(new Set());
  const [selectedCalendarDay, setSelectedCalendarDay] = useState(new Date().getDate());

  const now = new Date();
  const calendarMonth = useMemo(
    () => ({
      year: now.getFullYear(),
      month: now.getMonth(),
      label: now.toLocaleDateString('en-IN', { month: 'long', year: 'numeric' }),
    }),
    [now]
  );
  const calendarRows = useMemo(() => buildCalendarGrid(calendarMonth.year, calendarMonth.month), [calendarMonth]);
  const todayDateOnly = useMemo(
    () => new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime(),
    [now]
  );
  const welcomeDate = now.toLocaleDateString('en-IN', {
    weekday: 'long',
    day: 'numeric',
    month: 'long',
    year: 'numeric',
  });
  const welcomeTime = now.toLocaleTimeString('en-IN', {
    hour: '2-digit',
    minute: '2-digit',
    hour12: true,
    timeZoneName: 'short',
  });

  useEffect(() => {
    let cancelled = false;

    const loadMetrics = async () => {
      try {
        const [companyRes, statutoryRes, bulkRes] = await Promise.all([
          fetch(COMPANY_API, { cache: 'no-store' }),
          fetch(STATUTORY_API, { cache: 'no-store' }),
          fetch(CHECKLIST_BULK_API, { cache: 'no-store' }),
        ]);

        let activeCompanies = 0;
        let companyNames = [];
        if (companyRes.ok) {
          const companyJson = await companyRes.json();
          const companyDetails = companyJson?.data?.companyDetails;
          activeCompanies = Array.isArray(companyDetails) ? companyDetails.length : 0;
          companyNames = Array.isArray(companyDetails)
            ? companyDetails
                .map((item) => String(item?.companyName ?? item?.CompanyName ?? '').trim())
                .filter(Boolean)
            : [];
        }

        let pendingApprovals = 0;
        let approvedForms = 0;
        let rejectedForms = 0;
        let upcomingTasks = [];
        const dueDateCandidates = [];
        const statutoryStatusByAct = new Map();
        const statutoryStatusByActDescription = new Map();
        if (statutoryRes.ok) {
          const statutoryJson = await statutoryRes.json();
          const statutoryData = Array.isArray(statutoryJson?.data?.statutoryData)
            ? statutoryJson.data.statutoryData
            : [];

          pendingApprovals = statutoryData.filter((item) => {
            const status = String(item?.status ?? item?.Status ?? '').trim().toLowerCase();
            return status === 'pending';
          }).length;

          approvedForms = statutoryData.filter((item) => {
            const status = String(item?.status ?? item?.Status ?? '').trim().toLowerCase();
            return status === 'approved';
          }).length;

          rejectedForms = statutoryData.filter((item) => {
            const status = String(item?.status ?? item?.Status ?? '').trim().toLowerCase();
            return status === 'rejected' || status === 'reject';
          }).length;

          statutoryData.forEach((item) => {
            const actName = String(item?.act ?? item?.Act ?? item?.formName ?? item?.FormName ?? '').trim().toLowerCase();
            const description = String(
              item?.description ?? item?.Description ?? item?.desc ?? item?.Desc ?? ''
            ).trim().toLowerCase();
            const rawStatus = String(item?.status ?? item?.Status ?? '').trim();
            if (actName && rawStatus && !statutoryStatusByAct.has(actName)) {
              statutoryStatusByAct.set(actName, rawStatus);
            }
            const actDescriptionKey = buildActDescriptionKey(actName, description);
            if (actName && description && rawStatus && !statutoryStatusByActDescription.has(actDescriptionKey)) {
              statutoryStatusByActDescription.set(actDescriptionKey, rawStatus);
            }
          });

          upcomingTasks = statutoryData
            .filter((item) => hasValue(item?.act ?? item?.Act) || hasValue(item?.formName ?? item?.FormName))
            .slice(0, 4)
            .map((item, index) => {
              const dueDateInfo = getDueDateInfo(item?.dueDate ?? item?.DueDate, calendarMonth);
              if (dueDateInfo.day != null) dueDateCandidates.push(dueDateInfo.day);
              const statusMeta = getStatusMeta(item?.status ?? item?.Status);
              return {
                id: item?.id ?? item?.ROWID ?? `${item?.act ?? item?.Act}-${item?.dueDate ?? item?.DueDate}`,
                name: String(item?.act ?? item?.Act ?? item?.formName ?? item?.FormName ?? '-').trim() || '-',
                description: String(item?.description ?? item?.Description ?? item?.desc ?? item?.Desc ?? '-').trim() || '-',
                due: dueDateInfo.display,
                dueClass: 'nd-due-normal',
                company: String(
                  item?.companyName ??
                  item?.CompanyName ??
                  companyNames[index % Math.max(companyNames.length, 1)] ??
                  item?.site ??
                  item?.Site ??
                  item?.company ??
                  item?.Company ??
                  '-'
                ).trim() || '-',
                status: statusMeta.label,
                statusClass: statusMeta.statusClass,
                iconTone: statusMeta.iconTone,
                hasFile: hasSubmittedFile(item),
                sortDate: dueDateInfo.timestamp,
                calendarDay: dueDateInfo.day,
              };
            });
        }

        if (bulkRes.ok) {
          const bulkJson = await bulkRes.json();
          const bulkData = Array.isArray(bulkJson?.data) ? bulkJson.data : [];
          const existingKeys = new Set(
            upcomingTasks.map(
              (item) =>
                `${String(item.name).toLowerCase()}|${String(item.due).toLowerCase()}|${String(item.company).toLowerCase()}`
            )
          );

          bulkData.forEach((item, index) => {
            const dueDateInfo = getDueDateInfo(item?.dueDate ?? item?.DueDate, calendarMonth);
            if (dueDateInfo.day != null) dueDateCandidates.push(dueDateInfo.day);
            const actName = String(item?.act ?? item?.Act ?? item?.formName ?? item?.FormName ?? '').trim().toLowerCase();
            const description = String(
              item?.description ?? item?.Description ?? item?.desc ?? item?.Desc ?? ''
            ).trim().toLowerCase();
            const matchedStatus =
              statutoryStatusByActDescription.get(buildActDescriptionKey(actName, description)) ||
              'Yet to Complete';
            const statusMeta = getStatusMeta(matchedStatus);
            const task = {
              id: item?.id ?? item?.ROWID ?? `bulk-${index}`,
              name: String(item?.act ?? item?.Act ?? item?.formName ?? item?.FormName ?? '-').trim() || '-',
              description: String(item?.description ?? item?.Description ?? item?.desc ?? item?.Desc ?? '-').trim() || '-',
              due: dueDateInfo.timestamp != null ? dueDateInfo.display : '-',
              dueClass: 'nd-due-normal',
              company: String(
                companyNames[index % Math.max(companyNames.length, 1)] ??
                item?.companyName ??
                item?.CompanyName ??
                item?.site ??
                item?.Site ??
                item?.company ??
                item?.Company ??
                '-'
              ).trim() || '-',
              status: statusMeta.label,
              statusClass: statusMeta.statusClass,
              iconTone: statusMeta.iconTone,
              hasFile: hasSubmittedFile(item),
              sortDate: dueDateInfo.timestamp,
              calendarDay: dueDateInfo.day,
            };

            const key = buildTaskKey(task);
            if (!existingKeys.has(key)) {
              existingKeys.add(key);
              upcomingTasks.push(task);
            }
          });
        }

        const sortedUpcomingTasks = upcomingTasks
          .sort((a, b) => {
            const aTime = a?.sortDate ?? null;
            const bTime = b?.sortDate ?? null;
            if (aTime == null && bTime == null) return 0;
            if (aTime == null) return 1;
            if (bTime == null) return -1;
            const aDistance = Math.abs(aTime - todayDateOnly);
            const bDistance = Math.abs(bTime - todayDateOnly);
            if (aDistance !== bDistance) return aDistance - bDistance;
            return aTime - bTime;
          })
          .map(({ sortDate, ...item }) => item);

        const filteredUpcomingTasks =
          selectedCalendarDay == null
            ? sortedUpcomingTasks
            : sortedUpcomingTasks.filter((item) => item.calendarDay === selectedCalendarDay);

        upcomingTasks = limitUpcomingTasks(filteredUpcomingTasks, 4);

        if (!cancelled) {
          setCalendarDueDates(
            new Set(dueDateCandidates.filter((day) => day != null))
          );
          setTasks(upcomingTasks);
          setMetrics((prev) =>
            prev.map((metric) => {
              if (metric.key === 'companies') return { ...metric, value: activeCompanies };
              if (metric.key === 'approvals') return { ...metric, value: pendingApprovals };
              if (metric.key === 'due') return { ...metric, value: approvedForms };
              if (metric.key === 'policies') return { ...metric, value: rejectedForms };
              return metric;
            })
          );
        }
      } catch (_) {
        if (!cancelled) {
          setCalendarDueDates(new Set());
          setTasks([]);
          setMetrics((prev) =>
            prev.map((metric) => {
              if (
                metric.key === 'companies' ||
                metric.key === 'approvals' ||
                metric.key === 'due' ||
                metric.key === 'policies'
              ) {
                return { ...metric, value: 0 };
              }
              return metric;
            })
          );
        }
      }
    };

    loadMetrics();

    return () => {
      cancelled = true;
    };
  }, [calendarMonth.month, calendarMonth.year, selectedCalendarDay, todayDateOnly]);

  return (
    <div className="nd-root">
      <HcmDashboardSidebar userName={userName} userRole={userRole} userInitials={userInitials} userEmail={userEmail} />

      <div className="nd-main">
        <header className="nd-header">
          <h1 className="nd-header-title">HR Compliance Management</h1>
          <div className="nd-header-actions">
            <button type="button" className="nd-icon-btn nd-icon-btn--badge" aria-label="Notifications">
              <Bell size={20} />
              <span className="nd-badge-count">3</span>
            </button>
            <button type="button" className="nd-icon-btn" aria-label="Help">
              <CircleHelp size={20} />
            </button>
            <HcmHeaderProfileMenu userInitials={userInitials} userEmail={userEmail} />
          </div>
        </header>

        <main className="nd-content">
          <section className="nd-welcome">
            <div className="nd-welcome-bg" aria-hidden />
            <div className="nd-welcome-inner">
              <div className="nd-welcome-copy">
                <h2 className="nd-welcome-heading">Welcome to Vayona Energy HCM 👋</h2>
                <p className="nd-welcome-desc">
                  Track compliance, approvals, and statutory tasks across your organization in one
                  place.
                </p>
                <p className="nd-welcome-meta">
                  {welcomeDate} | {welcomeTime}
                </p>
              </div>
              <div className="nd-welcome-art" aria-hidden>
                <div className="nd-welcome-checklist">
                  <div className="nd-welcome-check-row">
                    <CheckCircle2 size={18} className="nd-wc-ok" />
                    <span />
                  </div>
                  <div className="nd-welcome-check-row">
                    <CheckCircle2 size={18} className="nd-wc-ok" />
                    <span />
                  </div>
                  <div className="nd-welcome-check-row">
                    <CheckCircle2 size={18} className="nd-wc-dim" />
                    <span />
                  </div>
                </div>
                <div className="nd-welcome-shield">
                  <ShieldCheck size={40} strokeWidth={1.5} />
                </div>
              </div>
            </div>
          </section>

          <section className="nd-metrics">
            {metrics.map((m) => {
              const Icon = m.icon;
              return (
                <article key={m.key} className={`nd-metric nd-metric--${m.tone}`}>
                  <div className="nd-metric-icon-wrap">
                    <Icon size={22} strokeWidth={2} />
                  </div>
                  <div className="nd-metric-body">
                    <span className="nd-metric-label">{m.label}</span>
                    <span className="nd-metric-value">{m.value}</span>
                  </div>
                </article>
              );
            })}
          </section>

          <div className="nd-grid-2">
            <section className="nd-card nd-card--table">
              <div className="nd-card-head">
                <h3 className="nd-card-title">Upcoming Compliance &amp; Tasks</h3>
                <Link to={TRANSACTION_PAGE} className="nd-link-btn">
                  View All
                </Link>
              </div>
              <div className="nd-table-wrap">
                <table className="nd-table">
                  <thead>
                    <tr>
                      <th>Act</th>
                      <th>Description</th>
                      <th>Due Date</th>
                      <th>Status</th>
                    </tr>
                  </thead>
                  <tbody>
                    {tasks.map((row) => (
                      <tr key={row.id}>
                        <td>
                          <div className="nd-task-cell">
                            <span className={`nd-task-dot nd-task-dot--${row.iconTone}`} />
                            {row.name}
                          </div>
                        </td>
                        <td>{row.description || '-'}</td>
                        <td>
                          <span className={row.dueClass}>{row.due}</span>
                        </td>
                        <td>
                          <span className={`nd-badge ${row.statusClass}`}>{row.status}</span>
                        </td>
                      </tr>
                    ))}
                    {tasks.length === 0 && (
                      <tr>
                        <td colSpan={4}>No transaction data available.</td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
            </section>

            <section className="nd-card nd-card--calendar">
              <div className="nd-card-head">
                <h3 className="nd-card-title">{calendarMonth.label}</h3>
              </div>
              <div className="nd-cal">
                <div className="nd-cal-weekdays">
                  {CAL_WEEKDAYS.map((d) => (
                    <span key={d} className="nd-cal-wd">
                      {d}
                    </span>
                  ))}
                </div>
                <div className="nd-cal-grid">
                  {calendarRows.map((week, wi) => (
                    <React.Fragment key={wi}>
                      {week.map((day, di) => {
                        if (day === null) {
                          return <span key={`e-${wi}-${di}`} className="nd-cal-day nd-cal-day--empty" />;
                        }
                        const isToday =
                          day === now.getDate() &&
                          calendarMonth.month === now.getMonth() &&
                          calendarMonth.year === now.getFullYear();
                        const marks = [];
                        if (calendarDueDates.has(day)) marks.push('due');
                        return (
                          <button
                            type="button"
                            key={day}
                            className={`nd-cal-day${isToday ? ' nd-cal-day--today' : ''}${selectedCalendarDay === day ? ' nd-cal-day--today' : ''}`}
                            onClick={() =>
                              setSelectedCalendarDay((current) => (current === day ? null : day))
                            }
                          >
                            <span className="nd-cal-num">{day}</span>
                            {marks.length > 0 && (
                              <span className="nd-cal-marks">
                                {marks.map((m) => (
                                  <span key={m} className={`nd-cal-mark nd-cal-mark--${m}`} />
                                ))}
                              </span>
                            )}
                          </button>
                        );
                      })}
                    </React.Fragment>
                  ))}
                </div>
              </div>
            </section>
          </div>
        </main>
      </div>
    </div>
  );
}

export default NewDashboard;
