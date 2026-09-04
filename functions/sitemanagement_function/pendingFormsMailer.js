'use strict';

const DEFAULT_APP_BASE = 'https://hcm-60066164950.development.catalystserverless.in';

const MONTH_NAMES = [
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

const MONTHLY_SEND_DAY = 20;

function nowInKolkata() {
  const parts = new Intl.DateTimeFormat('en-GB', {
    timeZone: 'Asia/Kolkata',
    year: 'numeric',
    month: 'numeric',
    day: 'numeric',
    hour: 'numeric',
    minute: 'numeric',
    hour12: false
  }).formatToParts(new Date());
  const get = (type) => Number(parts.find((p) => p.type === type)?.value || 0);
  return {
    year: get('year'),
    month: get('month'),
    day: get('day'),
    hour: get('hour'),
    minute: get('minute')
  };
}

function isMonthlySendDayInKolkata(now = nowInKolkata()) {
  return Number(now.day) === MONTHLY_SEND_DAY;
}

/** @deprecated use isMonthlySendDayInKolkata */
function isNineteenthInKolkata(now = nowInKolkata()) {
  return isMonthlySendDayInKolkata(now);
}

function currentPeriodLabel(now = nowInKolkata()) {
  const monthName = MONTH_NAMES[Math.max(0, Math.min(11, now.month - 1))] || '';
  return `${monthName} ${now.year}`.trim();
}

function trimText(value) {
  return String(value ?? '').trim();
}

function normalizeEmail(value) {
  return trimText(value).toLowerCase();
}

function splitEmails(value) {
  return String(value || '')
    .split(/[,;]/)
    .map((email) => normalizeEmail(email))
    .filter((email) => email && email.includes('@'));
}

function pickField(obj, ...keys) {
  if (!obj || typeof obj !== 'object') return '';
  for (const key of keys) {
    if (obj[key] != null && String(obj[key]).trim() !== '') return String(obj[key]).trim();
  }
  for (const key of Object.keys(obj)) {
    const lower = String(key).toLowerCase();
    if (keys.some((k) => String(k).toLowerCase() === lower)) {
      const v = obj[key];
      if (v != null && String(v).trim() !== '') return String(v).trim();
    }
  }
  return '';
}

function unwrapRow(row, tableName) {
  if (!row || typeof row !== 'object') return row;
  const nested = row[tableName];
  if (nested && typeof nested === 'object') {
    return { ...nested, ROWID: nested.ROWID ?? row.ROWID };
  }
  return row;
}

function escapeHtml(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function normToken(value) {
  return trimText(value)
    .toLowerCase()
    .replace(/\s+/g, ' ');
}

function normalizeStateCompareKey(value) {
  return trimText(value)
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '');
}

function splitStateTokens(value) {
  return String(value || '')
    .split(/[,;/|]/)
    .map((t) => normToken(t))
    .filter(Boolean);
}

function statesMatch(rowStateField, siteState) {
  const target = normalizeStateCompareKey(siteState);
  if (!target) return false;
  const raw = trimText(rowStateField);
  if (!raw) return false;
  const blob = normToken(raw);
  if (/\b(all india|pan-india|pan india|national|central|all states|all state)\b/.test(blob)) {
    return false;
  }
  const tokens = splitStateTokens(raw);
  const keys = (tokens.length ? tokens : [raw]).map(normalizeStateCompareKey).filter(Boolean);
  return keys.some((k) => k && (k === target || k.includes(target) || target.includes(k)));
}

function sectorMatchesIndustry(sector, industry) {
  const sec = normToken(sector);
  const ind = normToken(industry);
  if (!sec || !ind) return !sec || !ind;
  return sec === ind || sec.includes(ind) || ind.includes(sec);
}

function industryToActCategory(industry) {
  const s = normToken(industry);
  if (!s) return null;
  if (s.includes('factories') || s.includes('factory act') || s === 'factory') return 'factories';
  if (s.includes('shops') && s.includes('establishment')) return 'shops_and_establishment';
  if (s.includes('clra') || s.includes('contract labour') || s.includes('contract labor')) return 'clra';
  return null;
}

function formToActCategory(row) {
  const act = normToken(pickField(row, 'Act', 'act'));
  const sector = normToken(pickField(row, 'Sector', 'sector'));
  const blob = `${act} ${sector}`;
  if (blob.includes('clra') || blob.includes('contract labour') || blob.includes('contract labor')) {
    return 'clra';
  }
  if (
    blob.includes('factories act') ||
    blob.includes('factory act') ||
    (blob.includes('factories') && !blob.includes('contract')) ||
    blob.includes('factory')
  ) {
    return 'factories';
  }
  if (
    blob.includes('shops and establishments') ||
    blob.includes('shops and establishment') ||
    blob.includes('shop and establishment')
  ) {
    return 'shops_and_establishment';
  }
  return 'other';
}

function hasDraftFile(row) {
  const draftVal = pickField(row, 'DraftFile', 'draftFile');
  return Boolean(draftVal) && draftVal.toLowerCase() !== 'null' && draftVal.toLowerCase() !== 'undefined';
}

function isApprovedForm(row) {
  const status = pickField(row, 'Status', 'status').toLowerCase();
  const approval = pickField(row, 'Approval', 'approval').toLowerCase();
  return (
    status === 'approved' ||
    status === 'approve' ||
    approval === 'approved' ||
    approval === 'approve'
  );
}

function isReturnedForm(row) {
  const status = pickField(row, 'Status', 'status').toLowerCase();
  const approval = pickField(row, 'Approval', 'approval').toLowerCase();
  return (
    status === 'rejected' ||
    status === 'reject' ||
    status === 'returned' ||
    approval === 'rejected' ||
    approval === 'reject' ||
    approval === 'returned'
  );
}

function sendForApprovalIsSent(row) {
  return /^sent$/i.test(pickField(row, 'SendForApproval', 'sendForApproval'));
}

function isPanIndiaOrBlankState(stateRaw) {
  const raw = trimText(stateRaw);
  if (!raw) return true;
  return /\b(all india|pan india|pan-india|national|central|all states|all state)\b/.test(normToken(raw));
}

function normalizeFormKey(value) {
  return trimText(value)
    .toLowerCase()
    .replace(/[^a-z0-9]/g, '');
}

function baseFormNameKey(name) {
  const n = normToken(name);
  const m = n.match(/^form\s+[a-z0-9]+/);
  return m ? m[0] : n;
}

function formKeysMatch(a, b) {
  const na = pickField(a, 'FormName', 'formName');
  const nb = pickField(b, 'FormName', 'formName');
  if (normalizeFormKey(na) && normalizeFormKey(na) === normalizeFormKey(nb)) return true;
  const ba = baseFormNameKey(na);
  const bb = baseFormNameKey(nb);
  return Boolean(ba && bb && ba === bb);
}

/**
 * Reports status (Mainreport):
 * - No statutory row / no draft → Yet to Complete (from checklistbulk)
 * - Statutory overlay → Pending, Returned, or Approved only
 */
function getReportDisplayStatus(stRow) {
  if (!stRow) return 'Yet to Complete';
  if (isApprovedForm(stRow)) return 'Approved';
  if (isReturnedForm(stRow)) return 'Returned';
  if (sendForApprovalIsSent(stRow)) return 'Pending';
  const status = pickField(stRow, 'Status', 'status');
  const normalized = status.toLowerCase();
  if (normalized === 'pending') return 'Pending';
  if (!hasDraftFile(stRow)) return 'Yet to Complete';
  if (!status || status === '-' || status === '—') return 'Pending';
  if (
    normalized.includes('yet to complete') ||
    normalized.includes('yet to comply') ||
    normalized.includes('yet to submit')
  ) {
    return 'Yet to Complete';
  }
  return 'Pending';
}

function siteNamesMatch(a, b) {
  const x = normToken(a);
  const y = normToken(b);
  if (!x || !y) return false;
  return x === y || x.includes(y) || y.includes(x);
}

function bulkBelongsToSite(bulk, site) {
  const siteName = pickField(site, 'SiteName', 'siteName');
  const bulkSite = pickField(bulk, 'Site', 'site');
  if (bulkSite && !siteNamesMatch(bulkSite, siteName)) return false;

  const siteState = pickField(site, 'SiteState', 'siteState', 'State', 'state');
  const bulkState = pickField(bulk, 'State', 'state');
  if (siteState && bulkState && !isPanIndiaOrBlankState(bulkState) && !statesMatch(bulkState, siteState)) {
    return false;
  }

  const siteIndustry = pickField(site, 'Industry', 'industry');
  const siteCat = industryToActCategory(siteIndustry);
  const bulkCat = formToActCategory(bulk);
  if (siteCat && bulkCat !== 'other' && siteCat !== bulkCat) return false;
  const bulkSector = pickField(bulk, 'Sector', 'sector');
  if (bulkSector && siteIndustry && !sectorMatchesIndustry(bulkSector, siteIndustry) && siteCat !== bulkCat) {
    return false;
  }
  return true;
}

function statutoryBelongsToSite(row, site) {
  const siteName = pickField(site, 'SiteName', 'siteName');
  const rowSite = pickField(row, 'Site', 'site');
  if (rowSite) return siteNamesMatch(rowSite, siteName);
  const rowState = pickField(row, 'State', 'state');
  const siteState = pickField(site, 'SiteState', 'siteState', 'State', 'state');
  if (!rowState || !siteState || !statesMatch(rowState, siteState)) return false;
  const rowSector = pickField(row, 'Sector', 'sector');
  const siteIndustry = pickField(site, 'Industry', 'industry');
  const rowCat = formToActCategory(row);
  const siteCat = industryToActCategory(siteIndustry);
  if (rowCat && siteCat && rowCat !== 'other' && rowCat === siteCat) return true;
  if (rowSector && siteIndustry) return sectorMatchesIndustry(rowSector, siteIndustry);
  return !rowSector || !siteIndustry;
}

function pickBestStatutoryOverlay(candidates) {
  if (!candidates.length) return null;
  const score = (r) => {
    const s = getReportDisplayStatus(r);
    if (s === 'Approved') return 400;
    if (s === 'Pending') return 300;
    if (s === 'Returned') return 200;
    if (hasDraftFile(r)) return 80;
    return 10;
  };
  return [...candidates].sort((a, b) => score(b) - score(a))[0];
}

function normalizeMonthName(value) {
  const s = trimText(value).toLowerCase();
  if (!s) return '';
  const full = MONTH_NAMES.find((x) => x.toLowerCase() === s);
  if (full) return full;
  const idx = MONTH_NAMES.findIndex((x) => x.toLowerCase().slice(0, 3) === s.slice(0, 3));
  return idx >= 0 ? MONTH_NAMES[idx] : '';
}

function statutoryMatchesPeriod(row, periodLabel) {
  const [monthName, yearText] = String(periodLabel || '').split(' ');
  const selected = normalizeMonthName(monthName);
  const selectedPrefix = selected.toLowerCase().slice(0, 3);
  const year = parseInt(yearText, 10);
  const storedMonth = normalizeMonthName(pickField(row, 'MonthFilter', 'monthFilter', 'monthfilter'));
  const due = pickField(row, 'DueDate', 'dueDate').toLowerCase();
  const blob = `${storedMonth} ${due}`.toLowerCase();
  const monthOk = selectedPrefix
    ? blob.includes(selectedPrefix) || blob.includes(selected.toLowerCase())
    : true;
  if (!monthOk && storedMonth) return false;
  if (!monthOk && due && selected) return false;
  if (!Number.isNaN(year)) {
    const yearHit = String(due).match(/\b(20\d{2})\b/);
    if (yearHit && parseInt(yearHit[1], 10) !== year) return false;
  }
  return true;
}

function findStatutoryOverlay(bulk, statutoryRows, siteName, periodLabel) {
  let candidates = statutoryRows.filter((r) => formKeysMatch(bulk, r));
  const siteMatches = candidates.filter((r) => siteNamesMatch(pickField(r, 'Site', 'site'), siteName));
  const noSite = candidates.filter((r) => !pickField(r, 'Site', 'site'));
  candidates = siteMatches.length ? siteMatches : noSite;
  const periodMatches = candidates.filter((r) => statutoryMatchesPeriod(r, periodLabel));
  if (periodMatches.length) return pickBestStatutoryOverlay(periodMatches);
  // Reports: no statutory match for this month → Yet to Complete (do not reuse old Approved/Pending).
  return null;
}

function mapDigestForm(source, status, site, periodLabel, overlay) {
  const monthName = String(periodLabel || '').split(' ')[0] || '';
  return {
    id: source.ROWID || source.id || '',
    formName: pickField(source, 'FormName', 'formName') || 'Untitled form',
    act: pickField(source, 'Act', 'act') || pickField(overlay, 'Act', 'act'),
    description:
      pickField(source, 'Description', 'description') || pickField(overlay, 'Description', 'description'),
    state: pickField(source, 'State', 'state') || pickField(site, 'SiteState', 'siteState'),
    site: pickField(site, 'SiteName', 'siteName'),
    dueDate: pickField(source, 'DueDate', 'dueDate') || pickField(overlay, 'DueDate', 'dueDate'),
    month: pickField(overlay, 'MonthFilter', 'monthFilter', 'monthfilter') || monthName,
    status
  };
}

/**
 * Reports inventory (checklistbulk) is the source of Yet to Complete.
 * Statutory is used only to overlay Pending / Returned / Approved.
 */
function buildSiteDigest(site, bulkRows, statutoryRows, periodLabel) {
  const pending = [];
  const seenKeys = new Set();
  const siteName = pickField(site, 'SiteName', 'siteName');

  for (const bulk of bulkRows) {
    if (!pickField(bulk, 'FormName', 'formName')) continue;
    if (!bulkBelongsToSite(bulk, site)) continue;
    const overlay = findStatutoryOverlay(bulk, statutoryRows, siteName, periodLabel);
    const status = getReportDisplayStatus(overlay);
    if (status === 'Approved') continue;
    const key = `${normalizeFormKey(pickField(bulk, 'FormName', 'formName'))}|${baseFormNameKey(
      pickField(bulk, 'FormName', 'formName')
    )}`;
    seenKeys.add(key);
    pending.push(mapDigestForm(bulk, status, site, periodLabel, overlay));
  }

  for (const st of statutoryRows) {
    if (!pickField(st, 'FormName', 'formName')) continue;
    if (!statutoryBelongsToSite(st, site)) continue;
    if (!statutoryMatchesPeriod(st, periodLabel)) continue;
    const status = getReportDisplayStatus(st);
    if (status === 'Approved' || status === 'Yet to Complete') continue;
    const key = `${normalizeFormKey(pickField(st, 'FormName', 'formName'))}|${baseFormNameKey(
      pickField(st, 'FormName', 'formName')
    )}`;
    if (seenKeys.has(key)) continue;
    seenKeys.add(key);
    pending.push(mapDigestForm(st, status, site, periodLabel, st));
  }

  pending.sort((a, b) => {
    const byStatus = String(a.status).localeCompare(String(b.status), undefined, { sensitivity: 'base' });
    if (byStatus !== 0) return byStatus;
    return String(a.formName).localeCompare(String(b.formName), undefined, { sensitivity: 'base' });
  });
  return pending;
}

function statusColor(status) {
  const s = String(status || '').toLowerCase();
  if (s === 'returned' || s === 'rejected') return '#b91c1c';
  if (s === 'pending') return '#b45309';
  return '#2563eb';
}

function formatAsOfLabel(now = nowInKolkata()) {
  const monthName = MONTH_NAMES[Math.max(0, Math.min(11, now.month - 1))] || '';
  return `${now.day} ${monthName} ${now.year}`.trim();
}

function siteDisplayName(siteName) {
  const name = trimText(siteName);
  if (!name) return 'Site';
  return /\bsite\b/i.test(name) ? name : `${name} Site`;
}

function resolveEmailLogoUrl() {
  const fromEnv = trimText(process.env.LOGO_URL);
  if (fromEnv) return fromEnv;
  const base = trimText(process.env.APP_BASE_URL) || DEFAULT_APP_BASE;
  const root = base.replace(/\/+$/, '');
  return `${root}/server/sitemanagement_function/sitemanagement/email-logo`;
}

function vayonaLogoHtml() {
  const logoUrl = resolveEmailLogoUrl();
  return `<table role="presentation" cellpadding="0" cellspacing="0" border="0">
    <tr>
      <td bgcolor="#ffffff" style="background-color:#ffffff;padding:10px 14px;">
        <img src="${escapeHtml(logoUrl)}" alt="VAYONA ENERGY" width="160" style="display:block;border:0;outline:none;text-decoration:none;width:160px;height:auto;max-width:160px;-ms-interpolation-mode:bicubic;" />
      </td>
    </tr>
  </table>`;
}

function mintIconCell(symbol) {
  return `<table role="presentation" width="32" height="32" cellpadding="0" cellspacing="0" border="0">
    <tr>
      <td width="32" height="32" align="center" valign="middle" bgcolor="#d8f3ec" style="background-color:#d8f3ec;font-family:Arial,Helvetica,sans-serif;font-size:16px;color:#0f766e;line-height:32px;">${symbol}</td>
    </tr>
  </table>`;
}

function rowIndexBadge(n) {
  return `<table role="presentation" width="28" height="28" cellpadding="0" cellspacing="0" border="0" align="center">
    <tr>
      <td width="28" height="28" align="center" valign="middle" bgcolor="#d8f3ec" style="background-color:#d8f3ec;font-family:Arial,Helvetica,sans-serif;font-size:13px;font-weight:700;color:#0f766e;line-height:28px;">${n}</td>
    </tr>
  </table>`;
}

function buildEmailHtml({ siteName, inchargeName, periodLabel, asOfLabel, forms }) {
  const greetingName = inchargeName ? escapeHtml(inchargeName) : 'Site In-Charge';
  const siteLabel = siteDisplayName(siteName);
  const asOf = asOfLabel || `${MONTHLY_SEND_DAY} ${periodLabel}`;
  const logoBlock = vayonaLogoHtml();
  const rows = forms
    .map((form, idx) => {
      const border = idx === forms.length - 1 ? '0' : '1px solid #e4edef';
      return `<tr>
          <td align="center" valign="middle" style="padding:14px 8px;border-bottom:${border};">${rowIndexBadge(idx + 1)}</td>
          <td valign="middle" style="padding:14px 12px;border-bottom:${border};font-family:Arial,Helvetica,sans-serif;font-size:14px;font-weight:700;color:#111827;">${escapeHtml(form.formName)}</td>
          <td valign="middle" style="padding:14px 12px;border-bottom:${border};font-family:Arial,Helvetica,sans-serif;font-size:13px;line-height:1.45;color:#334155;">${escapeHtml(form.act || '—')}</td>
          <td valign="middle" style="padding:14px 12px;border-bottom:${border};font-family:Arial,Helvetica,sans-serif;font-size:13px;line-height:1.45;color:#334155;">${escapeHtml(form.description || '—')}</td>
          <td valign="middle" style="padding:14px 12px;border-bottom:${border};font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#334155;">${escapeHtml(form.month || '—')}</td>
          <td valign="middle" style="padding:14px 12px;border-bottom:${border};font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#334155;">${escapeHtml(form.dueDate || '—')}</td>
          <td valign="middle" style="padding:14px 12px;border-bottom:${border};font-family:Arial,Helvetica,sans-serif;font-size:13px;font-weight:600;color:${statusColor(form.status)};">&#8226;&nbsp;${escapeHtml(form.status)}</td>
        </tr>`;
    })
    .join('');

  return `<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Forms yet to complete</title>
</head>
<body style="margin:0;padding:0;background-color:#f4f7f8;">
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" bgcolor="#f4f7f8" style="background-color:#f4f7f8;">
    <tr>
      <td align="center" style="padding:0;">
        <table role="presentation" width="920" cellpadding="0" cellspacing="0" border="0" bgcolor="#ffffff" style="width:920px;max-width:920px;background-color:#ffffff;">
          <tr>
            <td bgcolor="#005c61" style="padding:28px 32px;background-color:#005c61;">
              <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0">
                <tr>
                  <td width="170" valign="middle" style="padding-right:18px;border-right:1px solid #7db0ae;">
                    ${logoBlock}
                  </td>
                  <td valign="middle" style="padding-left:22px;font-family:Arial,Helvetica,sans-serif;">
                    <span style="display:block;font-size:28px;line-height:34px;font-weight:700;color:#ffffff;">Forms yet to complete</span>
                    <span style="display:block;font-size:16px;line-height:24px;color:#d7eeee;padding-top:8px;">${escapeHtml(siteLabel)} &bull; ${escapeHtml(periodLabel)}</span>
                  </td>
                  <td width="44" align="right" valign="middle" style="font-family:Arial,Helvetica,sans-serif;font-size:28px;color:#8fd0ce;line-height:1;">&#128196;</td>
                </tr>
              </table>
            </td>
          </tr>
          <tr>
            <td style="padding:28px 32px 12px;font-family:Arial,Helvetica,sans-serif;color:#111827;">
              <p style="margin:0 0 14px;font-size:15px;line-height:1.55;">Dear ${greetingName},</p>
              <p style="margin:0;font-size:15px;line-height:1.65;color:#1f2937;">
                The following report forms for the <strong>${escapeHtml(siteLabel)}</strong> are yet to be completed as of <strong>${escapeHtml(asOf)}</strong>. Please review and complete them in the <strong>HR Compliance Management Tool</strong>.
              </p>
            </td>
          </tr>
          <tr>
            <td style="padding:8px 32px 0;">
              <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0"><tr><td style="border-top:1px solid #e4edef;font-size:0;line-height:0;">&nbsp;</td></tr></table>
            </td>
          </tr>
          <tr>
            <td style="padding:22px 32px 10px;">
              <table role="presentation" cellpadding="0" cellspacing="0" border="0">
                <tr>
                  <td valign="middle" style="padding-right:10px;">${mintIconCell('&#128196;')}</td>
                  <td valign="middle" style="font-family:Arial,Helvetica,sans-serif;font-size:20px;font-weight:700;color:#111827;">Forms Overview</td>
                </tr>
              </table>
            </td>
          </tr>
          <tr>
            <td style="padding:8px 32px 18px;">
              <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="border:1px solid #dbe5e8;">
                <tr bgcolor="#005c61">
                  <th align="center" width="48" bgcolor="#005c61" style="padding:12px 8px;background-color:#005c61;font-family:Arial,Helvetica,sans-serif;font-size:13px;font-weight:700;color:#ffffff;">#</th>
                  <th align="left" bgcolor="#005c61" style="padding:12px;background-color:#005c61;font-family:Arial,Helvetica,sans-serif;font-size:13px;font-weight:700;color:#ffffff;">Form</th>
                  <th align="left" bgcolor="#005c61" style="padding:12px;background-color:#005c61;font-family:Arial,Helvetica,sans-serif;font-size:13px;font-weight:700;color:#ffffff;">Act</th>
                  <th align="left" bgcolor="#005c61" style="padding:12px;background-color:#005c61;font-family:Arial,Helvetica,sans-serif;font-size:13px;font-weight:700;color:#ffffff;">Description</th>
                  <th align="left" bgcolor="#005c61" style="padding:12px;background-color:#005c61;font-family:Arial,Helvetica,sans-serif;font-size:13px;font-weight:700;color:#ffffff;">Month</th>
                  <th align="left" bgcolor="#005c61" style="padding:12px;background-color:#005c61;font-family:Arial,Helvetica,sans-serif;font-size:13px;font-weight:700;color:#ffffff;">Due date</th>
                  <th align="left" bgcolor="#005c61" style="padding:12px;background-color:#005c61;font-family:Arial,Helvetica,sans-serif;font-size:13px;font-weight:700;color:#ffffff;">Status</th>
                </tr>
                ${rows}
              </table>
            </td>
          </tr>
          <tr>
            <td style="padding:4px 32px 0;">
              <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0"><tr><td style="border-top:1px solid #e4edef;font-size:0;line-height:0;">&nbsp;</td></tr></table>
            </td>
          </tr>
          <tr>
            <td style="padding:16px 32px 28px;">
              <table role="presentation" cellpadding="0" cellspacing="0" border="0">
                <tr>
                  <td valign="middle" style="padding-right:10px;">${mintIconCell('&#9993;')}</td>
                  <td valign="middle" style="font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#6b7280;">This is an automated notification email.</td>
                </tr>
              </table>
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
</body>
</html>`;
}

const DEFAULT_MAIL_FROM = 'afrindinusha.j@buildhr.co.in';
const DEFAULT_MAIL_DISPLAY_NAME = 'Vayona Energy HCM';

function envFromEmail() {
  return trimText(process.env.MAIL_FROM) || trimText(process.env.FROM_EMAIL) || DEFAULT_MAIL_FROM;
}

function looksLikeEmail(value) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(trimText(value));
}

/**
 * Catalyst sendMail requires the verified mailbox string, not the Mail unique ID.
 * MAIL_FROM / default is used whenever it is a real email address.
 */
async function resolveFromEmail() {
  const preferred = envFromEmail();
  if (!looksLikeEmail(preferred)) {
    throw new Error(
      'MAIL_FROM must be the verified Email Configuration address (afrindinusha.j@buildhr.co.in), not a Mail unique ID.'
    );
  }
  console.log('pendingFormsMailer: using configured from_email', preferred);
  return preferred;
}

async function fetchAllRows(catalyst, tableName) {
  try {
    const rows = await catalyst.datastore().table(tableName).getAllRows();
    return Array.isArray(rows) ? rows : [];
  } catch (err) {
    console.warn(`pendingFormsMailer: failed to load ${tableName}:`, err?.message || err);
    return [];
  }
}

async function fetchSiteRows(catalyst) {
  const queries = [
    'SELECT ROWID, SiteName, SiteState, Industry, InchargeName, InchargeEmail FROM Site',
    'SELECT ROWID, SiteName, SiteState, Industry, InchargeName FROM Site'
  ];
  try {
    const zcql = catalyst.zcql();
    for (const sql of queries) {
      try {
        const rows = await zcql.executeZCQLQuery(sql);
        const mapped = (Array.isArray(rows) ? rows : []).map((row) => unwrapRow(row, 'Site')).filter(Boolean);
        if (mapped.length) {
          console.log(`pendingFormsMailer: loaded ${mapped.length} site(s) via ZCQL`);
          return mapped;
        }
      } catch (err) {
        console.warn(`pendingFormsMailer: site ZCQL failed (${sql}):`, err?.message || err);
      }
    }
  } catch (err) {
    console.warn('pendingFormsMailer: ZCQL unavailable:', err?.message || err);
  }
  const fallback = (await fetchAllRows(catalyst, 'Site')).map((row) => unwrapRow(row, 'Site')).filter(Boolean);
  console.log(`pendingFormsMailer: loaded ${fallback.length} site(s) via getAllRows`);
  return fallback;
}

function siteInchargeEmails(site) {
  return splitEmails(
    pickField(
      site,
      'InchargeEmail',
      'inchargeEmail',
      'incharge_email',
      'InchargeMail',
      'InchargeMailId',
      'MailId',
      'mail_id'
    )
  );
}

function isFromEmailConfigError(err) {
  const msg = String(err?.message || err || '').toLowerCase();
  return (
    msg.includes('from_email') ||
    msg.includes('from email') ||
    msg.includes('from_address') ||
    msg.includes('invalid mail') ||
    msg.includes('needs to be verified') ||
    msg.includes('domain')
  );
}

async function sendSiteEmail(catalyst, { fromEmail, toEmails, siteName, inchargeName, periodLabel, asOfLabel, forms }) {
  const subject = `Yet to complete / pending forms · ${siteName} · ${periodLabel}`;
  const html = buildEmailHtml({ siteName, inchargeName, periodLabel, asOfLabel, forms });
  const bcc = looksLikeEmail(fromEmail) && !toEmails.includes(normalizeEmail(fromEmail))
    ? [fromEmail]
    : undefined;
  console.log('pendingFormsMailer: sendMail', {
    from_email: fromEmail,
    to_email: toEmails,
    bcc: bcc || [],
    siteName,
    pendingCount: forms.length,
    subject
  });
  const mailObj = {
    from_email: fromEmail,
    to_email: toEmails,
    subject,
    content: html,
    html_mode: true,
    display_name: DEFAULT_MAIL_DISPLAY_NAME
  };
  if (bcc) mailObj.bcc = bcc;
  return catalyst.email().sendMail(mailObj);
}

async function sendOperatorSummaryEmail(catalyst, {
  fromEmail,
  periodLabel,
  emailsSent,
  emailsFailed,
  sitesSkippedNoEmail,
  sitesSkippedNoPending,
  details
}) {
  const sentRows = details.filter((d) => d.status === 'sent');
  const list = sentRows
    .map(
      (d) =>
        `<li><strong>${escapeHtml(d.siteName)}</strong> — ${d.pendingCount} form(s) (Yet to Complete / Pending / Returned) to ${escapeHtml(
          (d.toEmails || []).join(', ')
        )}</li>`
    )
    .join('');
  const html = `<!DOCTYPE html>
<html>
  <body style="font-family:Segoe UI,Arial,sans-serif;color:#111827;">
    <h2>Yet to complete / pending forms digest sent · ${escapeHtml(periodLabel)}</h2>
    <p>Site In-Charge emails sent: <strong>${emailsSent}</strong>. Failed: ${emailsFailed}. Skipped (no incharge email): ${sitesSkippedNoEmail}. Skipped (no open forms): ${sitesSkippedNoPending}.</p>
    <ul>${list || '<li>None</li>'}</ul>
    <p style="color:#6b7280;font-size:12px;">This copy is sent to the verified Mail sender so you can confirm delivery. Each Site In-Charge also received their own site-wise list.</p>
  </body>
</html>`;
  await catalyst.email().sendMail({
    from_email: fromEmail,
    to_email: [fromEmail],
    subject: `HCM yet to complete digest sent · ${periodLabel} · ${emailsSent} site(s)`,
    content: html,
    html_mode: true,
    display_name: DEFAULT_MAIL_DISPLAY_NAME
  });
}

/**
 * Compile report forms (Yet to Complete from checklistbulk + Pending/Returned from Statutory)
 * site-wise and email each Site In-Charge. Approved forms are excluded.
 */
async function sendPendingFormsDigest(catalyst, options = {}) {
  const now = nowInKolkata();
  const periodLabel = currentPeriodLabel(now);
  const asOfLabel = formatAsOfLabel(now);
  const force = Boolean(options.force);
  const dryRun = Boolean(options.dryRun);
  const requireSendDay =
    options.requireMonthlySendDay !== false && options.requireNineteenth !== false;

  if (requireSendDay && !force && !isMonthlySendDayInKolkata(now)) {
    return {
      skipped: true,
      reason: `Not the ${MONTHLY_SEND_DAY}th in Asia/Kolkata (today is ${now.day}). Pass force=true to send anyway.`,
      periodLabel,
      sitesConsidered: 0,
      emailsSent: 0,
      emailsFailed: 0,
      sitesSkippedNoEmail: 0,
      sitesSkippedNoPending: 0,
      details: []
    };
  }

  const [siteRows, statutoryRaw, bulkRaw] = await Promise.all([
    fetchSiteRows(catalyst),
    fetchAllRows(catalyst, 'Statutory'),
    fetchAllRows(catalyst, 'checklistbulk')
  ]);
  const statutoryRows = (statutoryRaw || []).map((row) => unwrapRow(row, 'Statutory'));
  const bulkRows = (bulkRaw || []).map((row) => unwrapRow(row, 'checklistbulk'));
  console.log(
    `pendingFormsMailer: loaded ${siteRows.length} site(s), ${bulkRows.length} report form(s), ${statutoryRows.length} statutory row(s)`
  );

  let fromEmail = '';
  if (!dryRun) {
    try {
      fromEmail = await resolveFromEmail(catalyst);
      console.log('pendingFormsMailer: sending with from_email', fromEmail);
    } catch (err) {
      return {
        skipped: true,
        reason: err?.message || String(err),
        periodLabel,
        sitesConsidered: siteRows.length,
        emailsSent: 0,
        emailsFailed: 0,
        sitesSkippedNoEmail: 0,
        sitesSkippedNoPending: 0,
        details: []
      };
    }
  }

  const details = [];
  let emailsSent = 0;
  let emailsFailed = 0;
  let sitesSkippedNoEmail = 0;
  let sitesSkippedNoPending = 0;

  for (const site of siteRows) {
    const siteName = pickField(site, 'SiteName', 'siteName');
    if (!siteName) continue;
    const inchargeName = pickField(site, 'InchargeName', 'inchargeName');
    const toEmails = siteInchargeEmails(site);
    const forms = buildSiteDigest(site, bulkRows, statutoryRows, periodLabel);

    if (!toEmails.length) {
      sitesSkippedNoEmail += 1;
      details.push({
        siteName,
        status: 'skipped_no_email',
        pendingCount: forms.length
      });
      continue;
    }
    if (!forms.length) {
      sitesSkippedNoPending += 1;
      details.push({
        siteName,
        toEmails,
        status: 'skipped_no_pending',
        pendingCount: 0
      });
      continue;
    }

    const summary = {
      siteName,
      inchargeName,
      toEmails,
      pendingCount: forms.length,
      forms: forms.map((f) => ({
        formName: f.formName,
        act: f.act,
        description: f.description,
        month: f.month,
        dueDate: f.dueDate,
        status: f.status
      }))
    };

    if (dryRun) {
      summary.status = 'dry_run';
      details.push(summary);
      emailsSent += 1;
      continue;
    }

    try {
      await sendSiteEmail(catalyst, {
        fromEmail,
        toEmails,
        siteName,
        inchargeName,
        periodLabel,
        asOfLabel,
        forms
      });
      summary.status = 'sent';
      emailsSent += 1;
      details.push(summary);
      console.log(
        `pendingFormsMailer: sent ${forms.length} yet-to-complete/pending/returned form(s) for ${siteName} to ${toEmails.join(', ')}`
      );
    } catch (err) {
      console.error(`pendingFormsMailer: failed for ${siteName}:`, err?.message || err);
      summary.status = 'failed';
      summary.error = err?.message || String(err);
      emailsFailed += 1;
      details.push(summary);
      if (isFromEmailConfigError(err)) {
        console.error(
          'pendingFormsMailer: stopping remaining sends. Add and verify this from_email in Cloud Scale → Mail:',
          fromEmail
        );
        break;
      }
    }

    await new Promise((resolve) => setTimeout(resolve, 150));
  }

  if (!dryRun && emailsSent > 0 && looksLikeEmail(fromEmail)) {
    try {
      await sendOperatorSummaryEmail(catalyst, {
        fromEmail,
        periodLabel,
        emailsSent,
        emailsFailed,
        sitesSkippedNoEmail,
        sitesSkippedNoPending,
        details
      });
    } catch (err) {
      console.warn('pendingFormsMailer: operator summary email failed:', err?.message || err);
    }
  }

  return {
    skipped: false,
    fromEmail: fromEmail || envFromEmail() || null,
    periodLabel,
    sitesConsidered: siteRows.length,
    emailsSent,
    emailsFailed,
    sitesSkippedNoEmail,
    sitesSkippedNoPending,
    dryRun,
    details
  };
}

module.exports = {
  sendPendingFormsDigest,
  buildEmailHtml,
  resolveEmailLogoUrl,
  vayonaLogoHtml,
  isMonthlySendDayInKolkata,
  isNineteenthInKolkata,
  MONTHLY_SEND_DAY,
  currentPeriodLabel,
  nowInKolkata,
  fetchSiteRows,
  fetchAllRows,
  unwrapRow,
  pickField,
  bulkBelongsToSite,
  findStatutoryOverlay,
  getReportDisplayStatus,
  isApprovedForm,
  isReturnedForm,
  resolveFromEmail,
  looksLikeEmail,
  normalizeEmail,
  splitEmails,
  escapeHtml,
  isFromEmailConfigError,
  envFromEmail,
  DEFAULT_MAIL_DISPLAY_NAME,
  normalizeFormKey,
  baseFormNameKey,
  statutoryBelongsToSite,
  statutoryMatchesPeriod
};
