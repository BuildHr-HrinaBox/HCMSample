import React, { useMemo } from 'react';
import {
  Bell,
  CircleHelp,
  Building,
  ClipboardList,
  AlertTriangle,
  BookMarked,
  Upload,
  CheckCircle2,
  ShieldCheck,
  Building2,
} from 'lucide-react';
import HcmDashboardSidebar from '../components/HcmDashboardSidebar';
import HcmHeaderProfileMenu from '../components/HcmHeaderProfileMenu';
import './newdashboard.css';

const METRICS = [
  {
    key: 'companies',
    label: 'Active Companies',
    value: '24',
    trend: '↑ 8% vs last month',
    trendUp: true,
    icon: Building,
    tone: 'green',
  },
  {
    key: 'approvals',
    label: 'Pending Approvals',
    value: '12',
    trend: '↓ 2% vs last week',
    trendUp: false,
    icon: ClipboardList,
    tone: 'blue',
  },
  {
    key: 'due',
    label: 'Compliance Due',
    value: '5',
    trend: '↑ 3 vs last week',
    trendUp: true,
    icon: AlertTriangle,
    tone: 'orange',
  },
  {
    key: 'policies',
    label: 'Policies & Documents',
    value: '156',
    trend: '↑ 12 added this month',
    trendUp: true,
    icon: BookMarked,
    tone: 'purple',
  },
];

const TASKS = [
  {
    name: 'PF Return Filing',
    due: '18 Jun 2025',
    dueClass: 'nd-due-urgent',
    company: 'Vayona Solar Pvt Ltd',
    status: 'High Priority',
    statusClass: 'nd-badge-high',
    iconTone: 'orange',
  },
  {
    name: 'ESI Contribution',
    due: '22 Jun 2025',
    dueClass: 'nd-due-soon',
    company: 'GreenWind Energy',
    status: 'In Progress',
    statusClass: 'nd-badge-progress',
    iconTone: 'blue',
  },
  {
    name: 'Factory License Renewal',
    due: '28 Jun 2025',
    dueClass: 'nd-due-normal',
    company: 'Vayona Manufacturing',
    status: 'Pending',
    statusClass: 'nd-badge-pending',
    iconTone: 'grey',
  },
  {
    name: 'Annual Safety Audit',
    due: '05 Jul 2025',
    dueClass: 'nd-due-normal',
    company: 'Vayona Logistics',
    status: 'Upcoming',
    statusClass: 'nd-badge-upcoming',
    iconTone: 'green',
  },
];

const ACTIVITY = [
  {
    title: 'Policy Document Uploaded',
    meta: 'By Priya Singh • 2 hours ago',
    icon: Upload,
    tone: 'blue',
  },
  {
    title: 'Compliance checklist completed',
    meta: 'By Amit Verma • Yesterday',
    icon: CheckCircle2,
    tone: 'green',
  },
  {
    title: 'New subsidiary onboarded',
    meta: 'By System • 3 days ago',
    icon: Building2,
    tone: 'purple',
  },
];

/** June 2025: 1 = Sunday; highlight 17; markers for legend */
const CAL_MONTH = { year: 2025, month: 5, label: 'June 2025' };
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

const dueDates = new Set([3, 10, 18]);
const eventDates = new Set([12, 22]);
const holidayDates = new Set([15]);

function NewDashboard({ userName = 'Ravi Kumar', userRole = 'HR Admin', userInitials = 'RA', userEmail }) {
  const calendarRows = useMemo(() => buildCalendarGrid(CAL_MONTH.year, CAL_MONTH.month), []);

  const now = new Date();
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
            {METRICS.map((m) => {
              const Icon = m.icon;
              return (
                <article key={m.key} className={`nd-metric nd-metric--${m.tone}`}>
                  <div className="nd-metric-icon-wrap">
                    <Icon size={22} strokeWidth={2} />
                  </div>
                  <div className="nd-metric-body">
                    <span className="nd-metric-label">{m.label}</span>
                    <span className="nd-metric-value">{m.value}</span>
                    <span className={`nd-metric-trend${m.trendUp ? ' nd-metric-trend--up' : ' nd-metric-trend--down'}`}>
                      {m.trend}
                    </span>
                  </div>
                </article>
              );
            })}
          </section>

          <div className="nd-grid-2">
            <section className="nd-card nd-card--table">
              <div className="nd-card-head">
                <h3 className="nd-card-title">Upcoming Compliance &amp; Tasks</h3>
                <button type="button" className="nd-link-btn">
                  View All
                </button>
              </div>
              <div className="nd-table-wrap">
                <table className="nd-table">
                  <thead>
                    <tr>
                      <th>Task / Compliance</th>
                      <th>Due Date</th>
                      <th>Company</th>
                      <th>Status</th>
                    </tr>
                  </thead>
                  <tbody>
                    {TASKS.map((row) => (
                      <tr key={row.name}>
                        <td>
                          <div className="nd-task-cell">
                            <span className={`nd-task-dot nd-task-dot--${row.iconTone}`} />
                            {row.name}
                          </div>
                        </td>
                        <td>
                          <span className={row.dueClass}>{row.due}</span>
                        </td>
                        <td>{row.company}</td>
                        <td>
                          <span className={`nd-badge ${row.statusClass}`}>{row.status}</span>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </section>

            <section className="nd-card nd-card--calendar">
              <div className="nd-card-head">
                <h3 className="nd-card-title">{CAL_MONTH.label}</h3>
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
                        const isToday = day === 17;
                        const marks = [];
                        if (dueDates.has(day)) marks.push('due');
                        if (eventDates.has(day)) marks.push('event');
                        if (holidayDates.has(day)) marks.push('holiday');
                        return (
                          <span
                            key={day}
                            className={`nd-cal-day${isToday ? ' nd-cal-day--today' : ''}`}
                          >
                            <span className="nd-cal-num">{day}</span>
                            {marks.length > 0 && (
                              <span className="nd-cal-marks">
                                {marks.map((m) => (
                                  <span key={m} className={`nd-cal-mark nd-cal-mark--${m}`} />
                                ))}
                              </span>
                            )}
                          </span>
                        );
                      })}
                    </React.Fragment>
                  ))}
                </div>
              </div>
              <div className="nd-cal-legend">
                <span className="nd-cal-legend-item">
                  <span className="nd-cal-mark nd-cal-mark--due" /> 3 Due Dates
                </span>
                <span className="nd-cal-legend-item">
                  <span className="nd-cal-mark nd-cal-mark--event" /> 2 Events
                </span>
                <span className="nd-cal-legend-item">
                  <span className="nd-cal-mark nd-cal-mark--holiday" /> 1 Holiday
                </span>
              </div>
            </section>
          </div>

          <section className="nd-card nd-card--activity">
            <div className="nd-card-head">
              <h3 className="nd-card-title">Recent Activity</h3>
              <button type="button" className="nd-link-btn">
                View All Activity
              </button>
            </div>
            <ul className="nd-activity-list">
              {ACTIVITY.map((a) => {
                const Icon = a.icon;
                return (
                  <li key={a.title} className="nd-activity-item">
                    <div className={`nd-activity-icon nd-activity-icon--${a.tone}`}>
                      <Icon size={20} strokeWidth={2} />
                    </div>
                    <div className="nd-activity-text">
                      <span className="nd-activity-title">{a.title}</span>
                      <span className="nd-activity-meta">{a.meta}</span>
                    </div>
                  </li>
                );
              })}
            </ul>
          </section>
        </main>
      </div>
    </div>
  );
}

export default NewDashboard;
