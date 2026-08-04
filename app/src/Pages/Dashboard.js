import React, { useState, useMemo, useEffect, useCallback } from 'react';
import { MapContainer, TileLayer, Marker, Popup } from 'react-leaflet';
import { Link } from 'react-router-dom';
import L from 'leaflet';
import 'leaflet/dist/leaflet.css';
import './Dashboard.css';

const STATUTORY_API = '/server/statutoryreg_function/statutory';
const CHECKLISTBULK_API = '/server/checklistbulk_function/checklistbulk?action=getAll';
const SHOPS_ESTABLISHMENT_ONLY_EMAIL = 'afrindinusha.j@buildhr.co.in';
const CLRA_ONLY_EMAIL = 'afrindinusha@gmail.com';
const FACTORIES_ACT_ONLY_EMAIL = 'afrinatlin@gmail.com';

const COLORS = {
  approved: '#28a745',
  pending: '#ffc107',
  yetToSubmit: '#007bff',
  returned: '#dc3545',
  // Keep legacy keys for trend/act charts
  compliant: '#28a745',
  overdue: '#dc3545',
  inProgress: '#007bff',
};

const HEALTH_STATUS_LEGEND = [
  { key: 'approved', label: 'Approved', color: COLORS.approved },
  { key: 'pending', label: 'Pending', color: COLORS.pending },
  { key: 'yetToSubmit', label: 'Yet to submit', color: COLORS.yetToSubmit },
  { key: 'returned', label: 'Returned', color: COLORS.returned },
];

const STATUS_LEGEND = [
  { key: 'compliant', label: 'Compliant', color: COLORS.compliant },
  { key: 'pending', label: 'Pending', color: COLORS.pending },
  { key: 'overdue', label: 'Overdue', color: COLORS.overdue },
  { key: 'inProgress', label: 'In Progress', color: COLORS.inProgress },
];

const LOCATION_SCORES = [
  { name: 'Maharashtra', percent: 85, color: '#28a745', lat: 19.7515, lng: 75.7139 },
  { name: 'Tamil Nadu', percent: 78, color: '#28a745', lat: 11.1271, lng: 78.6569 },
  { name: 'Karnataka', percent: 72, color: '#ffc107', lat: 15.3173, lng: 75.7139 },
  { name: 'Gujarat', percent: 68, color: '#7c3aed', lat: 22.2587, lng: 71.1924 },
  { name: 'Rajasthan', percent: 60, color: '#e11d48', lat: 27.0238, lng: 74.2179 },
];

const MAP_MARKERS = [
  { name: 'Delphi Oragadam', state: 'Tamil Nadu', percent: 78, color: '#007bff', lat: 12.9063, lng: 79.7769 },
  { name: 'Delphi Kakinada', state: 'Andhra Pradesh', percent: 68, color: '#f97316', lat: 16.9891, lng: 82.2475 },
  { name: 'Delphi HQ', state: 'Delhi', percent: 85, color: '#28a745', lat: 28.6139, lng: 77.209 },
];

const TREND_MONTHS = ['Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug'];
const TREND_SERIES = {
  compliant: [52, 58, 65, 70, 78, 85],
  pending: [22, 20, 18, 21, 19, 17],
  overdue: [14, 13, 12, 11, 12, 10],
  inProgress: [8, 7, 9, 8, 7, 6],
};

const FALLBACK_ACT_BARS = [
  { name: 'Factories Act 1948', compliant: 18, pending: 4, overdue: 2, inProgress: 2 },
  { name: 'ESI Act 1948', compliant: 16, pending: 3, overdue: 2, inProgress: 1 },
  { name: 'EPF Act 1952', compliant: 14, pending: 3, overdue: 1, inProgress: 2 },
  { name: 'Shops & Establishments Act', compliant: 12, pending: 4, overdue: 2, inProgress: 1 },
  { name: 'Professional Tax Act', compliant: 10, pending: 3, overdue: 1, inProgress: 1 },
  { name: 'Other Acts', compliant: 2, pending: 8, overdue: 1, inProgress: 1 },
];

const isShopsAndEstablishment = (item) => {
  const act = String(item.act || item.Act || '').toLowerCase().trim();
  const sector = String(item.sector || item.Sector || '').toLowerCase().trim();
  const isSE = (s) => s === 'shops and establishment' || s.includes('shops and establishment');
  return isSE(act) || isSE(sector);
};

const isCLRA = (item) => {
  const act = String(item.act || item.Act || '').toLowerCase().trim();
  const sector = String(item.sector || item.Sector || '').toLowerCase().trim();
  const isCLRAS = (s) =>
    s === 'clra' || s.includes('clra') || s.includes('contract labour') || s.includes('contract labor');
  return isCLRAS(act) || isCLRAS(sector);
};

const isFactoriesAct = (item) => {
  const act = String(item.act || item.Act || '').toLowerCase().trim();
  const sector = String(item.sector || item.Sector || '').toLowerCase().trim();
  const isFA = (s) =>
    s.includes('factories act') || s.includes('factory act') || s.includes('factories') || s.includes('factory');
  return isFA(act) || isFA(sector);
};

const hasDraftFile = (item) => {
  const draftVal = item?.draftFile ?? item?.DraftFile ?? null;
  return (
    draftVal != null &&
    String(draftVal).trim() !== '' &&
    String(draftVal).trim() !== 'null' &&
    String(draftVal).trim() !== 'undefined'
  );
};

const statutorySendForApprovalIsSent = (row) => {
  if (!row) return false;
  return /^sent$/i.test(String(row?.sendForApproval ?? row?.SendForApproval ?? '').trim());
};

/** Exact same display labels as Statutory.js getStatutoryRowDisplayStatus. */
const getStatutoryRowDisplayStatus = (item) => {
  const raw = item?.status != null && item.status !== '' ? item.status : item?.Status;
  const norm = String(raw || '').trim();
  const normalized = norm.toLowerCase();
  const rawApproval = item?.approval != null && item.approval !== '' ? item.approval : item?.Approval;
  const approvalNorm = String(rawApproval || '').trim().toLowerCase();

  if (
    approvalNorm === 'approved' ||
    approvalNorm === 'approve' ||
    normalized === 'approved' ||
    normalized === 'approve'
  ) {
    return 'Approved';
  }
  if (
    approvalNorm === 'rejected' ||
    approvalNorm === 'reject' ||
    normalized === 'rejected' ||
    normalized === 'reject'
  ) {
    return 'Returned';
  }
  if (statutorySendForApprovalIsSent(item)) return 'Pending';
  if (!hasDraftFile(item)) return 'Yet to Complete';
  if (norm === '' || norm === '-' || norm === '—') return 'Pending';
  if (normalized === 'pending') return 'Pending';
  if (normalized === 'yet to complete' || normalized === 'yet to comply' || normalized === 'yet to submit') {
    return statutorySendForApprovalIsSent(item) ? 'Pending' : 'Yet to Complete';
  }
  if (normalized === 'returned') return 'Returned';
  return norm || 'Yet to Complete';
};

/** Compliance Health keys — Yet to submit = Statutory "Yet to Complete". */
const getStatutoryHealthKey = (item) => {
  const display = getStatutoryRowDisplayStatus(item);
  if (display === 'Approved') return 'approved';
  if (display === 'Pending') return 'pending';
  if (display === 'Returned') return 'returned';
  // Statutory "Yet to Complete" (and any yet-to-* label) → Yet to submit
  if (display === 'Yet to Complete' || /yet to/i.test(display)) return 'yetToSubmit';
  return 'yetToSubmit';
};

const classifyItem = (item) => {
  const healthKey = getStatutoryHealthKey(item);
  if (healthKey === 'approved') return 'compliant';
  if (healthKey === 'returned') return 'overdue';
  if (healthKey === 'pending') return 'pending';
  return 'inProgress';
};

const normalizeRow = (item) => ({
  formFile: item.formFile ?? item.FormFile ?? null,
  proofSubmissionFile: item.proofSubmissionFile ?? item.ProofSubmissionFile ?? null,
  draftFile: item.draftFile ?? item.DraftFile ?? null,
  sendForApproval: item.sendForApproval ?? item.SendForApproval ?? '',
  approval: item.approval ?? item.Approval ?? '',
  act: item.act ?? item.Act ?? '',
  sector: item.sector ?? item.Sector ?? '',
  state: item.state ?? item.State ?? '',
  description: item.description ?? item.Description ?? '',
  status: item.status ?? item.Status ?? '',
  formName: String(item.formName ?? item.FormName ?? '').trim() || 'Form',
});

const rowMatchKey = (item) => {
  const form = String(item.formName || item.FormName || '')
    .toLowerCase()
    .trim();
  const act = String(item.act || item.Act || '')
    .toLowerCase()
    .trim();
  const desc = String(item.description || item.Description || '')
    .toLowerCase()
    .trim();
  return `${form}|${act}|${desc}`;
};

const formOnlyKey = (item) =>
  String(item.formName || item.FormName || '')
    .toLowerCase()
    .trim();

const donorScore = (row) => {
  let score = 0;
  if (hasDraftFile(row)) score += 1000;
  if (statutorySendForApprovalIsSent(row)) score += 500;
  const approval = String(row.approval || row.Approval || '')
    .trim()
    .toLowerCase();
  if (approval === 'approved' || approval === 'approve') score += 200;
  if (approval === 'rejected' || approval === 'reject') score += 150;
  const status = String(row.status || row.Status || '')
    .trim()
    .toLowerCase();
  if (status === 'approved' || status === 'pending' || status === 'rejected') score += 50;
  return score;
};

/**
 * Mirror Statutory page row set: ChecklistBulk 1:1 with statutory draft/status overlay.
 * "Yet to Complete" on Statutory = no draft after this merge.
 */
const buildStatutoryDisplayRows = (statutoryItems, bulkItems) => {
  if (!Array.isArray(bulkItems) || bulkItems.length === 0) {
    return statutoryItems;
  }

  const byFullKey = new Map();
  const byFormKey = new Map();
  statutoryItems.forEach((item) => {
    const full = rowMatchKey(item);
    const form = formOnlyKey(item);
    if (full) {
      const list = byFullKey.get(full) || [];
      list.push(item);
      byFullKey.set(full, list);
    }
    if (form) {
      const list = byFormKey.get(form) || [];
      list.push(item);
      byFormKey.set(form, list);
    }
  });

  return bulkItems.map((bulkItem) => {
    const bulk = normalizeRow(bulkItem);
    const full = rowMatchKey(bulk);
    const form = formOnlyKey(bulk);
    const candidates = [...(byFullKey.get(full) || []), ...(byFormKey.get(form) || [])];
    const unique = [];
    const seen = new Set();
    candidates.forEach((c) => {
      const id = `${rowMatchKey(c)}|${c.draftFile}|${c.status}|${c.sendForApproval}`;
      if (seen.has(id)) return;
      seen.add(id);
      unique.push(c);
    });
    const donor = [...unique].sort((a, b) => donorScore(b) - donorScore(a))[0] || null;
    if (!donor) return bulk;
    return {
      ...bulk,
      draftFile: donor.draftFile ?? donor.DraftFile ?? bulk.draftFile,
      proofSubmissionFile: donor.proofSubmissionFile ?? bulk.proofSubmissionFile,
      sendForApproval: donor.sendForApproval || bulk.sendForApproval,
      approval: donor.approval || bulk.approval,
      status: donor.status || bulk.status,
      act: bulk.act || donor.act,
      sector: bulk.sector || donor.sector,
    };
  });
};

function createMarkerIcon(color, borderColor) {
  return L.divIcon({
    className: 'custom-marker',
    html: `
      <div style="
        width: 28px; height: 28px;
        background-color: ${color};
        border: 3px solid ${borderColor};
        border-radius: 50% 50% 50% 0;
        transform: rotate(-45deg);
        box-shadow: 0 2px 8px rgba(0,0,0,0.28);
        position: relative;
      ">
        <div style="
          width: 8px; height: 8px;
          background-color: white;
          border-radius: 50%;
          position: absolute;
          top: 50%; left: 50%;
          transform: translate(-50%, -50%) rotate(45deg);
        "></div>
      </div>
    `,
    iconSize: [28, 28],
    iconAnchor: [14, 28],
  });
}

function IndiaMap({ locations }) {
  const markers = useMemo(
    () =>
      locations.map((loc) => ({
        ...loc,
        icon: createMarkerIcon(loc.color, loc.color),
      })),
    [locations]
  );

  return (
    <MapContainer
      center={[22.5, 78.9]}
      zoom={4}
      style={{ height: '100%', width: '100%', minHeight: 220 }}
      scrollWheelZoom={false}
      zoomControl={false}
      attributionControl={false}
      className="india-map-container"
    >
      <TileLayer url="https://{s}.basemaps.cartocdn.com/light_nolabels/{z}/{x}/{y}{r}.png" />
      {markers.map((loc) => (
        <Marker key={loc.name} position={[loc.lat, loc.lng]} icon={loc.icon}>
          <Popup>
            <strong>{loc.name}</strong>
            <br />
            {loc.state} · {loc.percent}% compliant
          </Popup>
        </Marker>
      ))}
    </MapContainer>
  );
}

function HealthGauge({ segments, centerPercent }) {
  const total = segments.reduce((sum, s) => sum + s.value, 0) || 1;
  let angle = -90;
  const radius = 70;
  const stroke = 18;
  const cx = 90;
  const cy = 90;
  const arcs = segments.map((seg) => {
    const sweep = (seg.value / total) * 270;
    const start = angle;
    angle += sweep;
    const startRad = (start * Math.PI) / 180;
    const endRad = ((start + sweep) * Math.PI) / 180;
    const x1 = cx + radius * Math.cos(startRad);
    const y1 = cy + radius * Math.sin(startRad);
    const x2 = cx + radius * Math.cos(endRad);
    const y2 = cy + radius * Math.sin(endRad);
    const large = sweep > 180 ? 1 : 0;
    return {
      ...seg,
      d: `M ${x1} ${y1} A ${radius} ${radius} 0 ${large} 1 ${x2} ${y2}`,
    };
  });

  return (
    <svg className="cd-gauge" viewBox="0 0 180 150" aria-hidden>
      <path
        d="M 20 90 A 70 70 0 1 1 160 90"
        fill="none"
        stroke="#eef2f7"
        strokeWidth={stroke}
        strokeLinecap="round"
      />
      {arcs.map((arc) =>
        arc.value > 0 ? (
          <path
            key={arc.key}
            d={arc.d}
            fill="none"
            stroke={arc.color}
            strokeWidth={stroke}
            strokeLinecap="butt"
          />
        ) : null
      )}
      <text x="90" y="82" textAnchor="middle" className="cd-gauge-value">
        {typeof centerPercent === 'number' ? `${centerPercent}%` : centerPercent}
      </text>
      <text x="90" y="102" textAnchor="middle" className="cd-gauge-label">
        Overall Compliance Score
      </text>
    </svg>
  );
}

function TrendChart({ series }) {
  const width = 520;
  const height = 240;
  const pad = { top: 16, right: 16, bottom: 36, left: 36 };
  const chartW = width - pad.left - pad.right;
  const chartH = height - pad.top - pad.bottom;
  const maxY = 100;
  const xAt = (i) => pad.left + (i / (TREND_MONTHS.length - 1)) * chartW;
  const yAt = (v) => pad.top + chartH - (v / maxY) * chartH;

  const paths = STATUS_LEGEND.map((status) => {
    const values = series[status.key] || [];
    const d = values
      .map((v, i) => `${i === 0 ? 'M' : 'L'} ${xAt(i)} ${yAt(v)}`)
      .join(' ');
    return { ...status, d, values };
  });

  return (
    <svg className="cd-trend-svg" viewBox={`0 0 ${width} ${height}`} preserveAspectRatio="xMidYMid meet">
      {[0, 25, 50, 75, 100].map((tick) => {
        const y = yAt(tick);
        return (
          <g key={tick}>
            <line x1={pad.left} y1={y} x2={pad.left + chartW} y2={y} stroke="#eef2f7" strokeWidth="1" />
            <text x={pad.left - 8} y={y + 4} textAnchor="end" className="cd-axis-text">
              {tick}
            </text>
          </g>
        );
      })}
      {TREND_MONTHS.map((m, i) => (
        <text key={m} x={xAt(i)} y={height - 10} textAnchor="middle" className="cd-axis-text">
          {m}
        </text>
      ))}
      {paths.map((p) => (
        <g key={p.key}>
          <path d={p.d} fill="none" stroke={p.color} strokeWidth="2.5" />
          {p.values.map((v, i) => (
            <circle key={`${p.key}-${i}`} cx={xAt(i)} cy={yAt(v)} r="4" fill="#fff" stroke={p.color} strokeWidth="2" />
          ))}
        </g>
      ))}
    </svg>
  );
}

function ActBarChart({ data }) {
  const width = 560;
  const height = 280;
  const pad = { top: 28, right: 12, bottom: 58, left: 42 };
  const chartW = width - pad.left - pad.right;
  const chartH = height - pad.top - pad.bottom;
  const maxY = 30;
  const groupW = chartW / Math.max(data.length, 1);
  const barW = Math.min(12, groupW / 5.5);
  const keys = ['compliant', 'pending', 'overdue', 'inProgress'];

  return (
    <svg className="cd-act-svg" viewBox={`0 0 ${width} ${height}`} preserveAspectRatio="xMidYMid meet">
      <text x="14" y="140" transform="rotate(-90 14 140)" className="cd-axis-text">
        No. of Compliances
      </text>
      {[0, 5, 10, 15, 20, 25, 30].map((tick) => {
        const y = pad.top + chartH - (tick / maxY) * chartH;
        return (
          <g key={tick}>
            <line x1={pad.left} y1={y} x2={pad.left + chartW} y2={y} stroke="#eef2f7" strokeWidth="1" />
            <text x={pad.left - 8} y={y + 4} textAnchor="end" className="cd-axis-text">
              {tick}
            </text>
          </g>
        );
      })}
      {data.map((act, i) => {
        const gx = pad.left + i * groupW + groupW / 2;
        const short =
          act.name.length > 16 ? `${act.name.slice(0, 14)}…` : act.name;
        return (
          <g key={act.name}>
            {keys.map((key, ki) => {
              const value = Number(act[key] || 0);
              const h = Math.max(value > 0 ? 4 : 0, (value / maxY) * chartH);
              const x = gx - (keys.length * barW + (keys.length - 1) * 3) / 2 + ki * (barW + 3);
              const y = pad.top + chartH - h;
              return (
                <g key={key}>
                  <rect x={x} y={y} width={barW} height={h} rx="2" fill={COLORS[key]} />
                  {value > 0 && (
                    <text x={x + barW / 2} y={y - 4} textAnchor="middle" className="cd-bar-label">
                      {value}
                    </text>
                  )}
                </g>
              );
            })}
            <text x={gx} y={height - 28} textAnchor="middle" className="cd-axis-text cd-act-label">
              {short}
            </text>
          </g>
        );
      })}
    </svg>
  );
}

const Dashboard = ({ userRole, userEmail, standalone = false }) => {
  const [period, setPeriod] = useState('This Month');
  const [health, setHealth] = useState({
    approved: 0,
    pending: 0,
    yetToSubmit: 0,
    returned: 0,
  });
  const [actBars, setActBars] = useState(FALLBACK_ACT_BARS);
  const [loading, setLoading] = useState(true);

  const isShopsAndEstablishmentOnlyUser =
    (userEmail || '').trim().toLowerCase() === SHOPS_ESTABLISHMENT_ONLY_EMAIL;
  const isCLRAOnlyUser = (userEmail || '').trim().toLowerCase() === CLRA_ONLY_EMAIL;
  const isFactoriesActOnlyUser = (userEmail || '').trim().toLowerCase() === FACTORIES_ACT_ONLY_EMAIL;

  const fetchDashboardData = useCallback(async () => {
    setLoading(true);
    try {
      const [statRes, bulkRes] = await Promise.all([
        fetch(STATUTORY_API, { cache: 'no-store' }),
        fetch(CHECKLISTBULK_API, { cache: 'no-store' }),
      ]);

      let statutoryItems = [];
      if (statRes.ok) {
        const statJson = await statRes.json();
        const statutoryData =
          statJson.data?.statutoryData || (Array.isArray(statJson.data) ? statJson.data : []);
        if (Array.isArray(statutoryData)) {
          statutoryItems = statutoryData.map(normalizeRow);
        }
      }

      let bulkItems = [];
      if (bulkRes.ok) {
        const bulkJson = await bulkRes.json();
        const bulkList = Array.isArray(bulkJson.data) ? bulkJson.data : [];
        bulkItems = bulkList.map(normalizeRow);
      }

      // Same row set as Statutory page: ChecklistBulk lines + statutory draft/status overlay.
      // Yet to submit = rows that show as "Yet to Complete" there.
      let itemsToCount = buildStatutoryDisplayRows(statutoryItems, bulkItems);
      if (isShopsAndEstablishmentOnlyUser) itemsToCount = itemsToCount.filter(isShopsAndEstablishment);
      else if (isCLRAOnlyUser) itemsToCount = itemsToCount.filter(isCLRA);
      else if (isFactoriesActOnlyUser) itemsToCount = itemsToCount.filter(isFactoriesAct);

      const counts = { approved: 0, pending: 0, yetToSubmit: 0, returned: 0 };
      itemsToCount.forEach((item) => {
        const key = getStatutoryHealthKey(item);
        if (counts[key] != null) counts[key] += 1;
        else counts.yetToSubmit += 1;
      });
      setHealth(counts);

      if (itemsToCount.length > 0) {
        const byAct = {};
        itemsToCount.forEach((item) => {
          const actName = String(item.act || '').trim() || 'Other Acts';
          if (!byAct[actName]) {
            byAct[actName] = { name: actName, compliant: 0, pending: 0, overdue: 0, inProgress: 0 };
          }
          byAct[actName][classifyItem(item)] += 1;
        });
        const ranked = Object.values(byAct)
          .map((a) => ({
            ...a,
            total: a.compliant + a.pending + a.overdue + a.inProgress,
          }))
          .sort((a, b) => b.total - a.total);

        const top = ranked.slice(0, 5);
        const rest = ranked.slice(5);
        if (rest.length) {
          top.push(
            rest.reduce(
              (acc, cur) => ({
                name: 'Other Acts',
                compliant: acc.compliant + cur.compliant,
                pending: acc.pending + cur.pending,
                overdue: acc.overdue + cur.overdue,
                inProgress: acc.inProgress + cur.inProgress,
              }),
              { name: 'Other Acts', compliant: 0, pending: 0, overdue: 0, inProgress: 0 }
            )
          );
        }
        setActBars(top.length ? top : FALLBACK_ACT_BARS);
      } else {
        setActBars(FALLBACK_ACT_BARS);
      }
    } catch {
      setHealth({ approved: 0, pending: 0, yetToSubmit: 0, returned: 0 });
      setActBars(FALLBACK_ACT_BARS);
    } finally {
      setLoading(false);
    }
  }, [isShopsAndEstablishmentOnlyUser, isCLRAOnlyUser, isFactoriesActOnlyUser]);

  useEffect(() => {
    fetchDashboardData();
  }, [fetchDashboardData]);

  const healthTotal =
    health.approved + health.pending + health.yetToSubmit + health.returned || 1;
  const overallScore = Math.round((health.approved / healthTotal) * 100);
  const footerScore = overallScore;
  const vsLastMonth = 8;

  const gaugeSegments = HEALTH_STATUS_LEGEND.map((s) => ({
    ...s,
    value: health[s.key] || 0,
  }));

  const healthRows = HEALTH_STATUS_LEGEND.map((s) => {
    const value = health[s.key] || 0;
    const pct = Math.round((value / healthTotal) * 100);
    return { ...s, value, pct };
  });

  return (
    <div
      className={`cd-page${standalone ? ' cd-page--standalone' : ''}`}
      style={{ background: '#ffffff', backgroundColor: '#ffffff' }}
    >
      <div className="cd-grid" style={{ background: '#ffffff', backgroundColor: '#ffffff' }}>
        {/* Compliance Health */}
        <section className="cd-card" style={{ background: 'transparent', backgroundColor: 'transparent', boxShadow: 'none', border: 'none' }}>
          <div className="cd-card-header">
            <h2 className="cd-card-title">Compliance Health</h2>
            <select
              className="cd-select"
              value={period}
              onChange={(e) => setPeriod(e.target.value)}
              aria-label="Period filter"
            >
              <option>This Month</option>
              <option>Last Month</option>
              <option>This Quarter</option>
            </select>
          </div>
          <div className="cd-health-body">
            <div className="cd-health-gauge-wrap">
              <HealthGauge segments={gaugeSegments} centerPercent={loading ? '—' : overallScore} />
            </div>
            <ul className="cd-health-legend">
              {healthRows.map((row) => (
                <li key={row.key}>
                  <span className="cd-dot" style={{ background: row.color }} />
                  <span className="cd-legend-label">{row.label}</span>
                  <span className="cd-legend-value">
                    {loading ? '—' : row.value}{' '}
                    <em>({loading ? '—' : `${row.pct}%`})</em>
                  </span>
                </li>
              ))}
            </ul>
          </div>
          <div className="cd-health-footer">
            <span>Overall Compliance Score</span>
            <div className="cd-health-footer-score">
              <strong>{loading ? '—' : `${footerScore}%`}</strong>
              <small>↑ {vsLastMonth}% vs last month</small>
            </div>
          </div>
        </section>

        {/* Compliance Trend */}
        <section className="cd-card" style={{ background: 'transparent', backgroundColor: 'transparent', boxShadow: 'none', border: 'none' }}>
          <div className="cd-card-header">
            <h2 className="cd-card-title">Compliance Trend (Last 6 Months)</h2>
            <Link to="/mainreport" className="cd-link">
              View Report
            </Link>
          </div>
          <div className="cd-inline-legend">
            {STATUS_LEGEND.map((s) => (
              <span key={s.key} className="cd-inline-legend-item">
                <span className="cd-dot" style={{ background: s.color }} />
                {s.label}
              </span>
            ))}
          </div>
          <div className="cd-chart-wrap">
            <TrendChart series={TREND_SERIES} />
          </div>
        </section>

        {/* Compliance by Location */}
        <section className="cd-card" style={{ background: 'transparent', backgroundColor: 'transparent', boxShadow: 'none', border: 'none' }}>
          <div className="cd-card-header">
            <h2 className="cd-card-title">Compliance by Location</h2>
          </div>
          <div className="cd-location-body">
            <div className="cd-location-map">
              <IndiaMap locations={MAP_MARKERS} />
            </div>
            <ul className="cd-location-list">
              {LOCATION_SCORES.map((loc) => (
                <li key={loc.name}>
                  <div className="cd-location-row">
                    <span className="cd-dot" style={{ background: loc.color }} />
                    <span className="cd-location-name">{loc.name}</span>
                    <span className="cd-location-pct">{loc.percent}%</span>
                  </div>
                  <div className="cd-progress">
                    <div
                      className="cd-progress-fill"
                      style={{ width: `${loc.percent}%`, background: loc.color }}
                    />
                  </div>
                </li>
              ))}
            </ul>
          </div>
          <div className="cd-card-footer-center">
            <Link to="/rule-book/site-management" className="cd-link">
              View All
            </Link>
          </div>
        </section>

        {/* Compliance by Act */}
        <section className="cd-card" style={{ background: 'transparent', backgroundColor: 'transparent', boxShadow: 'none', border: 'none' }}>
          <div className="cd-card-header">
            <h2 className="cd-card-title">Compliance by Act</h2>
            <Link to="/rule-book/actsbulk" className="cd-link">
              View All
            </Link>
          </div>
          <div className="cd-inline-legend">
            {STATUS_LEGEND.map((s) => (
              <span key={s.key} className="cd-inline-legend-item">
                <span className="cd-swatch" style={{ background: s.color }} />
                {s.label}
              </span>
            ))}
          </div>
          <div className="cd-chart-wrap">
            <ActBarChart data={actBars} />
          </div>
        </section>
      </div>
    </div>
  );
};

export default Dashboard;
