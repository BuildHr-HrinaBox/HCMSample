import React, { useState, useMemo, useEffect, useCallback } from 'react';
import {
  Eye,
  Calendar,
  Flame,
  AlertTriangle,
  Minimize2,
} from 'lucide-react';
import { MapContainer, TileLayer, Marker, Popup } from 'react-leaflet';
import L from 'leaflet';
import 'leaflet/dist/leaflet.css';
import './Dashboard.css';

const STATUTORY_API = '/server/statutoryreg_function/statutory';
const CHECKLIST_API = '/server/checklist_function/checklist';
const CHECKLISTBULK_API = '/server/checklistbulk_function/checklistbulk?action=getAll';
const SHOPS_ESTABLISHMENT_ONLY_EMAIL = 'afrindinusha.j@buildhr.co.in';
const CLRA_ONLY_EMAIL = 'afrindinusha@gmail.com';
const FACTORIES_ACT_ONLY_EMAIL = 'afrinatlin@gmail.com';

const hasProof = (item) =>
  item.proofSubmissionFile &&
  item.proofSubmissionFile !== 'null' &&
  item.proofSubmissionFile !== null &&
  String(item.proofSubmissionFile).trim() !== '';

const hasForm = (item) =>
  item.formFile &&
  item.formFile !== 'null' &&
  item.formFile !== null &&
  String(item.formFile).trim() !== '' &&
  !String(item.formFile).startsWith('STATUTORY_MASTER_FORM:');

const isShopsAndEstablishment = (item) => {
  const act = String(item.act || item.Act || '').toLowerCase().trim();
  const sector = String(item.sector || item.Sector || '').toLowerCase().trim();
  const isSE = (s) => s === 'shops and establishment' || s.includes('shops and establishment');
  return isSE(act) || isSE(sector);
};

const isCLRA = (item) => {
  const act = String(item.act || item.Act || '').toLowerCase().trim();
  const sector = String(item.sector || item.Sector || '').toLowerCase().trim();
  const isCLRAS = (s) => s === 'clra' || s.includes('clra') || s.includes('contract labour') || s.includes('contract labor');
  return isCLRAS(act) || isCLRAS(sector);
};

const isFactoriesAct = (item) => {
  const act = String(item.act || item.Act || '').toLowerCase().trim();
  const sector = String(item.sector || item.Sector || '').toLowerCase().trim();
  const isFA = (s) => s.includes('factories act') || s.includes('factory act') || s.includes('factories') || s.includes('factory');
  return isFA(act) || isFA(sector);
};

// Locations with coordinates in India (for map markers)
const LOCATIONS = [
  { name: 'Delphi Oragadam', state: 'Tamil Nadu', percent: 68, color: 'rgb(59, 130, 246)', lat: 12.9063, lng: 79.7769 },
  { name: 'Delphi Kakinada', state: 'Andhra Pradesh', percent: 54, color: 'rgb(249, 115, 22)', lat: 16.9891, lng: 82.2475 },
  { name: 'Delphi HQ', state: 'Delhi', percent: 83, color: 'rgb(16, 185, 129)', lat: 28.6139, lng: 77.2090 },
];

function createMarkerIcon(color, borderColor) {
  return L.divIcon({
    className: 'custom-marker',
    html: `
      <div style="
        width: 30px; height: 30px;
        background-color: ${color};
        border: 4px solid ${borderColor};
        border-radius: 50% 50% 50% 0;
        transform: rotate(-45deg);
        box-shadow: 0 2px 8px rgba(0,0,0,0.3);
        position: relative;
      ">
        <div style="
          width: 9px; height: 9px;
          background-color: white;
          border-radius: 50%;
          position: absolute;
          top: 50%; left: 50%;
          transform: translate(-50%, -50%) rotate(45deg);
        "></div>
      </div>
    `,
    iconSize: [30, 30],
    iconAnchor: [15, 30],
  });
}

function IndiaMap({ locations }) {
  const markers = useMemo(() => [
    { ...locations[0], icon: createMarkerIcon('#2563eb', '#1e40af') },
    { ...locations[1], icon: createMarkerIcon('#f97316', '#c2410c') },
    { ...locations[2], icon: createMarkerIcon('#10b981', '#047857') },
  ], [locations]);

  return (
    <MapContainer
      center={[20.5, 78.9]}
      zoom={5}
      style={{ height: '100%', width: '100%', minHeight: 280 }}
      scrollWheelZoom={true}
      className="india-map-container"
    >
      <TileLayer
        attribution='&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a>'
        url="https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png"
      />
      {markers.map((loc) => (
        <Marker key={loc.name} position={[loc.lat, loc.lng]} icon={loc.icon}>
          <Popup>
            <strong>{loc.name}</strong><br />
            {loc.state} · {loc.percent}% compliant
          </Popup>
        </Marker>
      ))}
    </MapContainer>
  );
}

const Dashboard = ({ userRole, userEmail }) => {
  const [activeLocationIndex, setActiveLocationIndex] = useState(0);
  const [expandedActIndex, setExpandedActIndex] = useState(null);
  const [proofStats, setProofStats] = useState({ completed: 0, pending: 0 });
  const [complianceActsData, setComplianceActsData] = useState([]);
  const [metricsLoading, setMetricsLoading] = useState(true);
  const locations = LOCATIONS;

  const isShopsAndEstablishmentOnlyUser = (userEmail || '').trim().toLowerCase() === SHOPS_ESTABLISHMENT_ONLY_EMAIL;
  const isCLRAOnlyUser = (userEmail || '').trim().toLowerCase() === CLRA_ONLY_EMAIL;
  const isFactoriesActOnlyUser = (userEmail || '').trim().toLowerCase() === FACTORIES_ACT_ONLY_EMAIL;

  // Fetch statutory + checklist + checklistbulk (same sources as Statutory page) so counts match
  const fetchProofCounts = useCallback(async () => {
    setMetricsLoading(true);
    try {
      const [statRes, checklistRes, bulkRes] = await Promise.all([
        fetch(STATUTORY_API, { cache: 'no-store' }),
        fetch(CHECKLIST_API, { cache: 'no-store' }),
        fetch(CHECKLISTBULK_API, { cache: 'no-store' })
      ]);
      let allItems = [];
      if (statRes.ok) {
        const statJson = await statRes.json();
        const statutoryData = statJson.data?.statutoryData || (Array.isArray(statJson.data) ? statJson.data : []);
        if (Array.isArray(statutoryData)) {
          statutoryData.forEach(item => {
            const formFile = item.formFile ?? item.FormFile ?? null;
            const proofSubmissionFile = item.proofSubmissionFile ?? item.ProofSubmissionFile ?? null;
            const act = item.act ?? item.Act ?? '';
            const sector = item.sector ?? item.Sector ?? '';
            const formName = (item.formName ?? item.FormName ?? '').trim() || 'Form';
            allItems.push({ formFile, proofSubmissionFile, act, sector, formName });
          });
        }
      }
      if (checklistRes.ok) {
        const clJson = await checklistRes.json();
        const checklistData = clJson.data?.checklistData || (Array.isArray(clJson.data) ? clJson.data : []);
        if (Array.isArray(checklistData)) {
          checklistData.forEach(item => {
            const formFile = item.formFile ?? item.FormFile ?? null;
            const proofSubmissionFile = item.proofSubmissionFile ?? item.ProofSubmissionFile ?? null;
            const act = item.act ?? item.Act ?? '';
            const sector = item.sector ?? item.Sector ?? '';
            const formName = (item.formName ?? item.FormName ?? '').trim() || 'Form';
            if (formFile && formFile !== 'null' && String(formFile).trim() !== '' && !String(formFile).startsWith('STATUTORY_MASTER_FORM:')) {
              allItems.push({ formFile, proofSubmissionFile, act, sector, formName });
            }
          });
        }
      }
      // ChecklistBulk: same rows as Statutory page (no proof/file stored in bulk – each row counts as pending until proof in statutory)
      if (bulkRes.ok) {
        const bulkJson = await bulkRes.json();
        const bulkList = Array.isArray(bulkJson.data) ? bulkJson.data : [];
        bulkList.forEach(item => {
          const act = item.act ?? item.Act ?? '';
          const sector = item.sector ?? item.Sector ?? '';
          const formName = (item.formName ?? item.FormName ?? '').trim() || 'Form';
          allItems.push({ formFile: null, proofSubmissionFile: null, act, sector, formName });
        });
      }
      // For afrindinusha.j@buildhr.co.in show only Shops and Establishment; afrindinusha@gmail.com only CLRA; afrinatlin@gmail.com only Factories Act
      let itemsToCount = allItems;
      if (isShopsAndEstablishmentOnlyUser) itemsToCount = allItems.filter(isShopsAndEstablishment);
      else if (isCLRAOnlyUser) itemsToCount = allItems.filter(isCLRA);
      else if (isFactoriesActOnlyUser) itemsToCount = allItems.filter(isFactoriesAct);
      const completed = itemsToCount.filter(item => hasProof(item)).length;
      const pending = itemsToCount.filter(item => !hasProof(item)).length;
      setProofStats({ completed, pending });
      // Build Compliance Score By Acts from statutory data: group by act, list forms per act
      const byAct = {};
      itemsToCount.forEach(item => {
        const actName = String(item.act || '').trim() || 'Other';
        if (!byAct[actName]) byAct[actName] = { name: actName, forms: [], completed: 0 };
        const formName = String(item.formName || '').trim() || 'Form';
        byAct[actName].forms.push({ formName, hasProof: hasProof(item) });
        if (hasProof(item)) byAct[actName].completed++;
      });
      const actsList = Object.values(byAct).map(a => ({
        name: a.name,
        count: a.forms.length,
        score: a.forms.length ? Math.round((a.completed / a.forms.length) * 100) : 0,
        forms: a.forms,
      }));
      setComplianceActsData(actsList);
    } catch {
      setProofStats({ completed: 0, pending: 0 });
      setComplianceActsData([]);
    } finally {
      setMetricsLoading(false);
    }
  }, [isShopsAndEstablishmentOnlyUser, isCLRAOnlyUser, isFactoriesActOnlyUser]);

  useEffect(() => {
    fetchProofCounts();
  }, [fetchProofCounts]);

  // Completed = proof submission uploaded count; Pending = proof not uploaded count (e.g. 13 statutory, 0 proof → Completed 0, Pending 13)
  const metrics = [
    { label: 'Completed', value: metricsLoading ? '—' : Number(proofStats.completed), icon: <Eye size={24} />, colorClass: 'metric-card-green' },
    { label: 'Pending', value: metricsLoading ? '—' : Number(proofStats.pending), icon: <Calendar size={24} />, colorClass: 'metric-card-blue' },
    { label: 'Reject', value: 0, icon: <Flame size={24} />, colorClass: 'metric-card-pink' },
    { label: 'Overdue', value: 0, icon: <AlertTriangle size={24} />, colorClass: 'metric-card-orange' },
  ];

  const totalProofItems = proofStats.completed + proofStats.pending;
  const compliantPercent = totalProofItems > 0 ? Math.round((proofStats.completed / totalProofItems) * 100) : 0;

  const complianceActs = metricsLoading ? [] : complianceActsData;
  const barChartData = complianceActs.map((a, i) => ({ label: a.name, score: a.score, index: i }));

  return (
    <div className="new-dashboard-page">
      <div className="metric-cards-container">
        {metrics.map((m) => (
          <div key={m.label} className={`metric-card ${m.colorClass}`}>
            <div className="metric-card-icon">{m.icon}</div>
            <div className="metric-card-value">{m.value}</div>
            <div className="metric-card-label">{m.label}</div>
          </div>
        ))}
      </div>

      <div className="new-dashboard-grid">
        <div className="dashboard-top-row">
          {/* Overall Compliance */}
          <div className="overall-compliance-card">
            <div className="card-header">
              <span className="card-dot accent" />
              <span>Overall Compliance</span>
            </div>
            <div className="overall-compliance-ring">
              <div
                className="overall-compliance-segment"
                style={{
                  background: `conic-gradient(#53c26b 0%, #53c26b ${compliantPercent}%, #f26a6a ${compliantPercent}%, #f26a6a 100%)`,
                }}
              >
                <div className="overall-compliance-inner">
                  <span>{compliantPercent}%</span>
                  <small>Compliant</small>
                </div>
              </div>
            </div>
          </div>

          {/* Location Compliance Scores + Map */}
          <div className="map-card">
            <div className="card-header">
              <span className="card-dot primary" />
              <span>Location Compliance Scores</span>
              <button type="button" className="map-toggle-btn" title="Minimize map" aria-label="Minimize map">
                <Minimize2 size={18} />
              </button>
            </div>
            <div className="map-frame">
              <IndiaMap locations={locations} />
            </div>
            <div className="map-locations-list">
              {locations.map((loc, idx) => (
                <button
                  key={loc.name}
                  type="button"
                  className={`map-location-chip ${activeLocationIndex === idx ? 'active' : ''}`}
                  onClick={() => setActiveLocationIndex(idx)}
                >
                  <div className="map-chip-header">
                    <span className="map-chip-dot" style={{ background: loc.color }} />
                    <span className="map-chip-name">{loc.name}</span>
                  </div>
                  <div className="map-chip-meta">{loc.state}</div>
                  <div className="map-chip-score">{loc.percent}% compliant</div>
                </button>
              ))}
            </div>
          </div>
        </div>

        <div className="dashboard-bottom-row">
          {/* Compliance Score By Acts */}
          <div className="compliance-acts-card">
            <div className="card-header">
              <span className="card-dot accent" />
              <span>Compliance Score By Acts</span>
            </div>
            <div className="compliance-acts-list">
              {metricsLoading ? (
                <div className="compliance-act-row" style={{ color: '#6b7280', padding: '12px 0' }}>
                  Loading…
                </div>
              ) : complianceActs.length === 0 ? (
                <div className="compliance-act-row" style={{ color: '#6b7280', padding: '12px 0' }}>
                  No statutory forms for this login yet.
                </div>
              ) : (
                complianceActs.map((act, idx) => (
                  <div key={act.name || idx} className="compliance-act-container">
                    <div
                      className="compliance-act-row compliance-act-row-clickable"
                      style={{ cursor: 'pointer' }}
                      onClick={() => setExpandedActIndex(expandedActIndex === idx ? null : idx)}
                      role="button"
                      tabIndex={0}
                      onKeyDown={(e) => e.key === 'Enter' && setExpandedActIndex(expandedActIndex === idx ? null : idx)}
                    >
                      <span className="compliance-act-name">
                        {act.name}
                        <span className="checklist-count-badge">({act.count})</span>
                      </span>
                      <span className={`compliance-act-score ${act.score >= 80 ? 'score-high' : act.score >= 50 ? 'score-mid' : ''}`}>
                        {metricsLoading ? '—' : `${act.score}%`}
                      </span>
                      <span className="expand-icon" style={{ marginLeft: '8px' }}>
                        {expandedActIndex === idx ? '▼' : '▶'}
                      </span>
                    </div>
                    {expandedActIndex === idx && act.forms && act.forms.length > 0 && (
                      <div className="compliance-act-forms" style={{ paddingLeft: '16px', paddingBottom: '12px', marginTop: '4px', borderLeft: '2px solid rgba(82, 194, 107, 0.4)', marginLeft: '8px' }}>
                        <div style={{ fontSize: '12px', fontWeight: 600, color: '#374151', marginBottom: '6px' }}>Forms in statutory</div>
                        {act.forms.map((f, fi) => (
                          <div key={fi} style={{ fontSize: '13px', color: '#4b5563', marginBottom: '4px', display: 'flex', alignItems: 'center', gap: '8px' }}>
                            <span>{f.formName}</span>
                            <span style={{ fontSize: '11px', color: f.hasProof ? '#059669' : '#d97706' }}>
                              {f.hasProof ? '✓ Proof submitted' : 'Pending'}
                            </span>
                          </div>
                        ))}
                      </div>
                    )}
                  </div>
                ))
              )}
            </div>
          </div>

          {/* Compliance Score Chart */}
          <div className="bar-chart-card">
            <div className="card-header">
              <span className="card-dot primary" />
              <span>Compliance Score Chart</span>
            </div>
            <div className="bar-chart-container">
              <div className="bar-chart-wrapper">
                <BarChartSVG data={barChartData} />
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};

function BarChartSVG({ data }) {
  const width = 400;
  const height = 300;
  const padding = { top: 20, right: 20, bottom: 60, left: 45 };
  const chartWidth = width - padding.left - padding.right;
  const chartHeight = height - padding.top - padding.bottom;
  const maxScore = 100;
  const barGap = 8;
  const barWidth = Math.max(28, (chartWidth - (data.length - 1) * barGap) / data.length - barGap);

  return (
    <svg
      className="bar-chart-svg"
      viewBox={`0 0 ${width} ${height}`}
      preserveAspectRatio="xMidYMid meet"
      width="100%"
      height="350"
    >
      {/* Horizontal grid lines */}
      {[0, 25, 50, 75, 100].map((pct, i) => {
        const y = padding.top + chartHeight - (pct / maxScore) * chartHeight;
        return (
          <g key={pct}>
            <line
              x1={padding.left}
              y1={y}
              x2={padding.left + chartWidth}
              y2={y}
              stroke="#e0e0e0"
              strokeWidth="1"
              strokeDasharray="2,2"
            />
            <text x={padding.left - 5} y={y + 4} textAnchor="end" fontSize="10" fill="#7a7a7a">
              {pct}%
            </text>
          </g>
        );
      })}
      <line
        x1={padding.left}
        y1={padding.top + chartHeight}
        x2={padding.left + chartWidth}
        y2={padding.top + chartHeight}
        stroke="#1f1f1f"
        strokeWidth="2"
      />
      {data.map((d, i) => {
        const x = padding.left + i * (barWidth + barGap) + barGap / 2;
        const barHeight = Math.max(2, (d.score / maxScore) * chartHeight);
        const y = padding.top + chartHeight - barHeight;
        const labelShort = d.label.length > 18 ? d.label.slice(0, 15) + '...' : d.label;
        return (
          <g key={d.label}>
            <rect
              x={x}
              y={padding.top + chartHeight - 2}
              width={barWidth + 4}
              height={4}
              fill="#f26a6a"
              rx="4"
              opacity="0.3"
              stroke="#e0e0e0"
              strokeWidth="1"
            />
            <rect
              x={x + 2}
              y={y + 8}
              width={barWidth}
              height={Math.max(24, barHeight - 8)}
              fill={d.score >= 80 ? '#e8f5e9' : d.score >= 50 ? '#fff3e0' : '#ffebee'}
              rx="12"
              opacity="0.95"
            />
            <text
              x={x + 2 + barWidth / 2}
              y={y + 24}
              textAnchor="middle"
              fontSize="14"
              fill={d.score >= 80 ? '#2e7d32' : d.score >= 50 ? '#e65100' : '#c62828'}
              fontWeight="600"
            >
              {d.score}%
            </text>
            <text
              x={x + 2 + barWidth / 2}
              y={padding.top + chartHeight + 28}
              textAnchor="middle"
              fontSize="9"
              fill="#7a7a7a"
              transform={`rotate(-45 ${x + 2 + barWidth / 2} ${padding.top + chartHeight + 28})`}
            >
              {labelShort}
            </text>
          </g>
        );
      })}
    </svg>
  );
}

export default Dashboard;
