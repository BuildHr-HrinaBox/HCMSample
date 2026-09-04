'use strict';

const fs = require('fs');
const os = require('os');
const path = require('path');
const {
  fetchAllRows,
  fetchSiteRows,
  unwrapRow,
  pickField,
  resolveFromEmail,
  looksLikeEmail,
  splitEmails,
  escapeHtml,
  isFromEmailConfigError,
  vayonaLogoHtml,
  currentPeriodLabel,
  nowInKolkata,
  DEFAULT_MAIL_DISPLAY_NAME,
  statutoryMatchesPeriod
} = require('./pendingFormsMailer');
const { buildChroRollupFromReports, buildPendingRollupFromReports } = require('./mainReportData');
const { buildApprovedSitesPdfBuffer, buildApprovedSitesPdfFileName } = require('./approvedSitesPdf');

const CHRO_TABLE = 'CHRO';
/** Main email card width (640px + ~2in at 96dpi). */
const CHRO_EMAIL_WIDTH_PX = 832;
/** Approved Sites Report is emailed to CHRO on the 28th of every month (Asia/Kolkata). */
const CHRO_MONTHLY_SEND_DAY = 28;

function isChroSendDayInKolkata(now = nowInKolkata()) {
  return Number(now.day) === CHRO_MONTHLY_SEND_DAY;
}

const STATE_ABBR = {
  andhrapradesh: 'AP',
  arunachalpradesh: 'AR',
  assam: 'AS',
  bihar: 'BR',
  chhattisgarh: 'CG',
  goa: 'GA',
  gujarat: 'GJ',
  haryana: 'HR',
  himachalpradesh: 'HP',
  jharkhand: 'JH',
  karnataka: 'KA',
  kerala: 'KL',
  madhyapradesh: 'MP',
  maharashtra: 'MH',
  manipur: 'MN',
  meghalaya: 'ML',
  mizoram: 'MZ',
  nagaland: 'NL',
  odisha: 'OD',
  orissa: 'OD',
  punjab: 'PB',
  rajasthan: 'RJ',
  sikkim: 'SK',
  tamilnadu: 'TN',
  telangana: 'TS',
  tripura: 'TR',
  uttarpradesh: 'UP',
  uttarakhand: 'UK',
  westbengal: 'WB',
  delhi: 'DL',
  nctofdelhi: 'DL',
  puducherry: 'PY',
  pondicherry: 'PY'
};

function stateKey(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '');
}

function stateAbbreviation(state) {
  const key = stateKey(state);
  if (!key) return '';
  if (STATE_ABBR[key]) return STATE_ABBR[key];
  if (key.length <= 3) return key.toUpperCase();
  return '';
}

function formatStatutoryIndustry(sector) {
  const raw = String(sector || '').trim();
  return raw || '—';
}

function formatSiteLabel(name, state) {
  const abbr = stateAbbreviation(state);
  const label = String(name || '').trim();
  if (abbr && label && !label.toUpperCase().startsWith(`${abbr} `) && !label.toUpperCase().startsWith(`${abbr}–`)) {
    return `${abbr} – ${label}`;
  }
  return label || '—';
}

const MONTH_INDEX = {
  january: 1,
  february: 2,
  march: 3,
  april: 4,
  may: 5,
  june: 6,
  july: 7,
  august: 8,
  september: 9,
  october: 10,
  november: 11,
  december: 12
};

function parsePeriodParts(periodLabel) {
  const [monthName, yearText] = String(periodLabel || '').split(/\s+/);
  const month = MONTH_INDEX[String(monthName || '').toLowerCase()] || 0;
  const year = parseInt(yearText, 10);
  return { month, year, monthName: monthName || '' };
}

function periodToYearMonth(periodLabel) {
  const { month, year, monthName } = parsePeriodParts(periodLabel);
  const names = [
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
  return {
    year: String(year || ''),
    month: names[month - 1] || monthName || String(periodLabel || '').split(/\s+/)[0] || ''
  };
}

function dateValueInPeriod(value, periodLabel) {
  const raw = String(value || '').trim();
  if (!raw || raw.toLowerCase() === 'null') return false;
  const { month, year } = parsePeriodParts(periodLabel);
  if (!month || !year) return false;
  const iso = raw.match(/^(\d{4})-(\d{2})(?:-(\d{2}))?/);
  if (iso) {
    return parseInt(iso[1], 10) === year && parseInt(iso[2], 10) === month;
  }
  const parsed = new Date(raw);
  if (Number.isNaN(parsed.getTime())) {
    const blob = raw.toLowerCase();
    const monthToken = String(parsePeriodParts(periodLabel).monthName || '').slice(0, 3).toLowerCase();
    return blob.includes(String(year)) && monthToken && blob.includes(monthToken);
  }
  return parsed.getFullYear() === year && parsed.getMonth() + 1 === month;
}

function statutoryRowInReportPeriod(row, periodLabel) {
  const storedMonth = pickField(row, 'MonthFilter', 'monthFilter', 'monthfilter');
  if (storedMonth) return statutoryMatchesPeriod(row, periodLabel);
  const dateFields = [
    pickField(row, 'SubmittedDate', 'submittedDate'),
    pickField(row, 'ApprovedDate', 'approvedDate'),
    pickField(row, 'ApprovalDate', 'approvalDate'),
    pickField(row, 'DraftDate', 'draftDate'),
    pickField(row, 'DueDate', 'dueDate'),
    pickField(row, 'CREATEDTIME', 'createdTime'),
    pickField(row, 'MODIFIEDTIME', 'modifiedTime')
  ];
  return dateFields.some((value) => dateValueInPeriod(value, periodLabel));
}

function statutorySendForApprovalIsSent(row) {
  return /^sent$/i.test(pickField(row, 'SendForApproval', 'sendForApproval'));
}

function statutoryHasDraftFile(row) {
  const draftVal = pickField(row, 'DraftFile', 'draftFile');
  return Boolean(draftVal) && draftVal.toLowerCase() !== 'null' && draftVal.toLowerCase() !== 'undefined';
}

/** Same Status column rules as Statutory.js / statutoryreg_function. */
function getStatutoryRegDisplayStatus(row) {
  const raw = pickField(row, 'Status', 'status');
  const normalized = raw.toLowerCase();
  const approvalNorm = pickField(row, 'Approval', 'approval').toLowerCase();
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
    normalized === 'reject' ||
    normalized === 'returned' ||
    approvalNorm === 'returned'
  ) {
    return 'Returned';
  }
  if (statutorySendForApprovalIsSent(row)) return 'Pending';
  if (!statutoryHasDraftFile(row)) return 'Yet to Complete';
  if (raw === '' || raw === '-' || raw === '—') {
    return statutorySendForApprovalIsSent(row) ? 'Pending' : 'Yet to Complete';
  }
  if (normalized === 'pending') {
    return statutorySendForApprovalIsSent(row) ? 'Pending' : 'Yet to Complete';
  }
  if (normalized === 'yet to complete' || normalized === 'yet to comply') {
    return statutorySendForApprovalIsSent(row) ? 'Pending' : 'Yet to Complete';
  }
  return raw || 'Yet to Complete';
}

async function fetchStatutoryRows(catalyst) {
  const raw = await fetchAllRows(catalyst, 'Statutory');
  return raw.map((row) => unwrapRow(row, 'Statutory')).filter(Boolean);
}

async function countStatutoryPendingAndReturned(catalyst, periodLabel) {
  const rows = await fetchStatutoryRows(catalyst);
  const inPeriod = rows.filter((row) => statutoryRowInReportPeriod(row, periodLabel));
  let pendingCount = 0;
  let returnedCount = 0;
  for (const row of inPeriod) {
    const status = getStatutoryRegDisplayStatus(row);
    if (status === 'Pending') pendingCount += 1;
    else if (status === 'Returned') returnedCount += 1;
  }
  console.log('approvedSitesMailer: statutoryreg counts', {
    periodLabel,
    statutoryRows: rows.length,
    inPeriod: inPeriod.length,
    pendingCount,
    returnedCount
  });
  return { pendingCount, returnedCount };
}

async function countActiveSitesFromSiteManagement(catalyst) {
  const siteRows = await fetchSiteRows(catalyst);
  const activeCount = siteRows.filter((site) => pickField(site, 'SiteName', 'siteName')).length;
  console.log('approvedSitesMailer: sitemanagement active sites', activeCount);
  return activeCount;
}

function approvedBadge() {
  return `<span style="display:inline-block;padding:4px 10px;border-radius:999px;background-color:#dcfce7;color:#166534;font-family:Arial,Helvetica,sans-serif;font-size:12px;font-weight:700;">Approved</span>`;
}

function pendingStatusBadge(status) {
  if (status === 'Pending') {
    return `<span style="display:inline-block;padding:4px 10px;border-radius:999px;background-color:#fef3c7;color:#92400e;font-family:Arial,Helvetica,sans-serif;font-size:12px;font-weight:700;">Pending</span>`;
  }
  return `<span style="display:inline-block;padding:4px 10px;border-radius:999px;background-color:#fee2e2;color:#991b1b;font-family:Arial,Helvetica,sans-serif;font-size:12px;font-weight:700;">Yet to Submit</span>`;
}

function formatPendingStatusLabel(status) {
  return status === 'Yet to Complete' ? 'Yet to Submit' : status || 'Yet to Submit';
}

function buildApprovedFormsTableHtml(approvedSites) {
  if (!Array.isArray(approvedSites) || approvedSites.length === 0) return '';

  const groups = new Map();
  for (const item of approvedSites) {
    const key = String(item.siteKey || item.siteName || item.label || '')
      .trim()
      .toLowerCase();
    if (!key) continue;
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(item);
  }

  const sortedGroups = [...groups.values()].sort((a, b) =>
    String(a[0]?.label || a[0]?.siteName || '').localeCompare(
      String(b[0]?.label || b[0]?.siteName || ''),
      undefined,
      { sensitivity: 'base' }
    )
  );

  const thStyle =
    'padding:10px 12px;font-family:Arial,Helvetica,sans-serif;font-size:12px;font-weight:700;color:#4b5563;';
  const siteHeaderStyle =
    'padding:14px 12px 10px;background-color:#f0fdf4;font-family:Arial,Helvetica,sans-serif;font-size:15px;font-weight:700;color:#166534;border-bottom:1px solid #dcfce7;';

  let html = '';
  sortedGroups.forEach((forms, groupIdx) => {
    forms.sort((a, b) =>
      String(a.formName || '').localeCompare(String(b.formName || ''), undefined, { sensitivity: 'base' })
    );
    const siteLabel = forms[0]?.label || forms[0]?.siteName || '—';
    const groupGap = groupIdx > 0 ? 'border-top:10px solid #ffffff;' : '';

    html += `<tr>
      <td colspan="5" style="${groupGap}${siteHeaderStyle}">${escapeHtml(siteLabel)}</td>
    </tr>
    <tr bgcolor="#f3f4f6">
      <th align="left" style="${thStyle}">State</th>
      <th align="left" style="${thStyle}">Industry</th>
      <th align="left" style="${thStyle}">Form Number</th>
      <th align="left" style="${thStyle}">Form Name</th>
      <th align="left" style="${thStyle}">Status</th>
    </tr>`;

    forms.forEach((form, formIdx) => {
      const isLastGroup = groupIdx === sortedGroups.length - 1;
      const isLastForm = formIdx === forms.length - 1;
      const border = isLastGroup && isLastForm ? '0' : '1px solid #edf2f0';
      const cell = (content) =>
        `<td valign="middle" style="padding:12px;border-bottom:${border};">${content}</td>`;
      html += `<tr>
        ${cell(`<span style="font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#334155;">${escapeHtml(form.state)}</span>`)}
        ${cell(`<span style="font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#334155;">${escapeHtml(form.industry)}</span>`)}
        ${cell(`<span style="font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#334155;">${escapeHtml(form.formNumber)}</span>`)}
        ${cell(`<span style="font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#111827;">${escapeHtml(form.formName)}</span>`)}
        ${cell(approvedBadge())}
      </tr>`;
    });
  });
  return html;
}

function buildPendingFormsTableHtml(pendingSites) {
  if (!Array.isArray(pendingSites) || pendingSites.length === 0) return '';

  const groups = new Map();
  for (const item of pendingSites) {
    const key = String(item.siteKey || item.siteName || item.label || '')
      .trim()
      .toLowerCase();
    if (!key) continue;
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(item);
  }

  const sortedGroups = [...groups.values()].sort((a, b) =>
    String(a[0]?.label || a[0]?.siteName || '').localeCompare(
      String(b[0]?.label || b[0]?.siteName || ''),
      undefined,
      { sensitivity: 'base' }
    )
  );

  const thStyle =
    'padding:10px 12px;font-family:Arial,Helvetica,sans-serif;font-size:12px;font-weight:700;color:#4b5563;';
  const siteHeaderStyle =
    'padding:14px 12px 10px;background-color:#fffbeb;font-family:Arial,Helvetica,sans-serif;font-size:15px;font-weight:700;color:#92400e;border-bottom:1px solid #fde68a;';

  let html = '';
  sortedGroups.forEach((forms, groupIdx) => {
    forms.sort((a, b) =>
      String(a.formName || '').localeCompare(String(b.formName || ''), undefined, { sensitivity: 'base' })
    );
    const siteLabel = forms[0]?.label || forms[0]?.siteName || '—';
    const groupGap = groupIdx > 0 ? 'border-top:10px solid #ffffff;' : '';

    html += `<tr>
      <td colspan="5" style="${groupGap}${siteHeaderStyle}">${escapeHtml(siteLabel)}</td>
    </tr>
    <tr bgcolor="#f3f4f6">
      <th align="left" style="${thStyle}">State</th>
      <th align="left" style="${thStyle}">Industry</th>
      <th align="left" style="${thStyle}">Form Number</th>
      <th align="left" style="${thStyle}">Form Name</th>
      <th align="left" style="${thStyle}">Status</th>
    </tr>`;

    forms.forEach((form, formIdx) => {
      const isLastGroup = groupIdx === sortedGroups.length - 1;
      const isLastForm = formIdx === forms.length - 1;
      const border = isLastGroup && isLastForm ? '0' : '1px solid #edf2f0';
      const cell = (content) =>
        `<td valign="middle" style="padding:12px;border-bottom:${border};">${content}</td>`;
      const statusLabel = formatPendingStatusLabel(form.status);
      html += `<tr>
        ${cell(`<span style="font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#334155;">${escapeHtml(form.state)}</span>`)}
        ${cell(`<span style="font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#334155;">${escapeHtml(form.industry)}</span>`)}
        ${cell(`<span style="font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#334155;">${escapeHtml(form.formNumber)}</span>`)}
        ${cell(`<span style="font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#111827;">${escapeHtml(form.formName)}</span>`)}
        ${cell(pendingStatusBadge(statusLabel))}
      </tr>`;
    });
  });
  return html;
}

function buildApprovedSitesEmailHtml({ chroName, periodLabel, complianceScore, approvedSites, pendingSites }) {
  const greeting = chroName ? escapeHtml(chroName) : 'CHRO';
  const logoBlock = vayonaLogoHtml();
  const score = `${complianceScore}%`;
  const rows = buildApprovedFormsTableHtml(approvedSites);
  const pendingRows = buildPendingFormsTableHtml(pendingSites);
  const emptyRow = `<tr><td colspan="5" style="padding:18px 12px;font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#6b7280;">No approved forms for ${escapeHtml(periodLabel)} yet.</td></tr>`;
  const emptyPendingRow = `<tr><td colspan="5" style="padding:18px 12px;font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#6b7280;">No pending or yet to submit forms for ${escapeHtml(periodLabel)}.</td></tr>`;

  return `<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>HR Compliance Approved Sites Report</title>
</head>
<body style="margin:0;padding:0;background-color:#f4f7f6;">
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" bgcolor="#f4f7f6" style="background-color:#f4f7f6;">
    <tr>
      <td align="center" style="padding:24px 12px;">
        <table role="presentation" width="${CHRO_EMAIL_WIDTH_PX}" cellpadding="0" cellspacing="0" border="0" bgcolor="#ffffff" style="width:${CHRO_EMAIL_WIDTH_PX}px;max-width:${CHRO_EMAIL_WIDTH_PX}px;background-color:#ffffff;border-radius:8px;">
          <tr>
            <td align="center" style="padding:22px 24px 8px;">${logoBlock}</td>
          </tr>
          <tr>
            <td style="padding:8px 24px 18px;">
              <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" bgcolor="#166534" style="background-color:#166534;border-radius:16px;">
                <tr>
                  <td style="padding:22px 24px;">
                    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0">
                      <tr>
                        <td valign="middle" width="44">
                          <table role="presentation" width="36" height="36" cellpadding="0" cellspacing="0" border="0">
                            <tr>
                              <td width="36" height="36" align="center" valign="middle" bgcolor="#ffffff" style="background-color:#ffffff;border-radius:18px;font-size:18px;color:#166534;line-height:36px;">&#10003;</td>
                            </tr>
                          </table>
                        </td>
                        <td valign="middle" style="padding-left:12px;font-family:Arial,Helvetica,sans-serif;">
                          <div style="font-size:26px;line-height:30px;font-weight:700;color:#ffffff;">HR Compliance</div>
                          <div style="font-size:14px;line-height:20px;color:#d1fae5;padding-top:4px;">Approved Sites &amp; Pending Forms Report &bull; ${escapeHtml(periodLabel)}</div>
                        </td>
                        <td width="118" align="right" valign="middle">
                          <table role="presentation" cellpadding="0" cellspacing="0" border="0" bgcolor="#ffffff" style="background-color:#ffffff;border-radius:10px;">
                            <tr>
                              <td align="center" style="padding:10px 12px;font-family:Arial,Helvetica,sans-serif;">
                                <div style="font-size:22px;line-height:24px;font-weight:700;color:#166534;">${escapeHtml(score)}</div>
                                <div style="font-size:10px;line-height:13px;color:#6b7280;padding-top:3px;">Compliance Score</div>
                              </td>
                            </tr>
                          </table>
                        </td>
                      </tr>
                    </table>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
          <tr>
            <td style="padding:8px 28px 6px;font-family:Arial,Helvetica,sans-serif;color:#111827;">
              <p style="margin:0 0 12px;font-size:15px;line-height:1.5;font-weight:700;">Dear ${greeting},</p>
              <p style="margin:0;font-size:14px;line-height:1.65;color:#374151;">
                Please find below the HR Compliance report for <strong>${escapeHtml(periodLabel)}</strong>, including approved sites and pending forms. A PDF copy is attached for your records. This report provides an overview of the compliance status of sites across the organization.
              </p>
            </td>
          </tr>
          <tr>
            <td style="padding:18px 28px 8px;font-family:Arial,Helvetica,sans-serif;font-size:18px;font-weight:700;color:#166534;">Approved Sites</td>
          </tr>
          <tr>
            <td style="padding:0 24px 24px;">
              <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="border:1px solid #e5e7eb;border-radius:10px;">
                ${rows || emptyRow}
              </table>
            </td>
          </tr>
          <tr>
            <td style="padding:18px 28px 8px;font-family:Arial,Helvetica,sans-serif;font-size:18px;font-weight:700;color:#92400e;">Pending Forms</td>
          </tr>
          <tr>
            <td style="padding:0 24px 24px;">
              <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="border:1px solid #e5e7eb;border-radius:10px;">
                ${pendingRows || emptyPendingRow}
              </table>
            </td>
          </tr>
          <tr>
            <td style="padding:0 28px 28px;font-family:Arial,Helvetica,sans-serif;font-size:12px;color:#6b7280;">This is an automated notification email.</td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
</body>
</html>`;
}

function pickChroRow(rowObj) {
  if (!rowObj || typeof rowObj !== 'object' || Array.isArray(rowObj)) return null;
  const nested = rowObj.CHRO || rowObj.Chro || rowObj.chro || null;
  if (nested && typeof nested === 'object' && !Array.isArray(nested)) {
    return { ...nested, ROWID: nested.ROWID ?? rowObj.ROWID };
  }
  const keys = Object.keys(rowObj);
  if (keys.length === 1) {
    const only = rowObj[keys[0]];
    if (only && typeof only === 'object' && !Array.isArray(only)) {
      return { ...only, ROWID: only.ROWID ?? rowObj.ROWID };
    }
  }
  return rowObj;
}

function emailsFromChroRow(row) {
  const named = splitEmails(
    pickField(
      row,
      'Email',
      'email',
      'EmailId',
      'emailId',
      'MailId',
      'MailID',
      'Mail_Id',
      'mail_id',
      'Mail',
      'mail',
      'Mailid',
      'Mail Id',
      'mail id'
    )
  );
  if (named.length) return named;
  const lookup = {};
  Object.keys(row || {}).forEach((key) => {
    lookup[String(key).toLowerCase().replace(/[_\s-]/g, '')] = row[key];
  });
  const fromNormalized = splitEmails(
    lookup.email || lookup.emailid || lookup.mailid || lookup.mail || ''
  );
  if (fromNormalized.length) return fromNormalized;
  const found = [];
  Object.values(row || {}).forEach((value) => {
    if (typeof value === 'string' && value.includes('@')) {
      splitEmails(value).forEach((email) => found.push(email));
    }
  });
  return found;
}

async function fetchChroTableRows(catalyst) {
  let raw = [];
  try {
    const zcql = catalyst.zcql();
    const pageSize = 100;
    let offset = 0;
    while (offset < 2000) {
      const chunk = await zcql.executeZCQLQuery(
        `SELECT * FROM ${CHRO_TABLE} LIMIT ${offset},${pageSize}`
      );
      if (!Array.isArray(chunk) || chunk.length === 0) break;
      chunk.forEach((row) => raw.push(row));
      if (chunk.length < pageSize) break;
      offset += pageSize;
    }
    console.log(`approvedSitesMailer: loaded ${raw.length} CHRO row(s) via ZCQL`);
  } catch (err) {
    console.warn('approvedSitesMailer: CHRO ZCQL failed:', err?.message || err);
    raw = [];
  }
  if (!raw.length) {
    raw = await fetchAllRows(catalyst, CHRO_TABLE);
    console.log(`approvedSitesMailer: loaded ${raw.length} CHRO row(s) via getAllRows`);
  }
  return raw.map((row) => pickChroRow(row)).filter(Boolean);
}

async function fetchChroRecipients(catalyst) {
  const rows = await fetchChroTableRows(catalyst);
  const seen = new Set();
  const recipients = [];
  for (const row of rows) {
    const status = pickField(row, 'Status', 'status', 'NotificationStatus', 'notificationStatus');
    if (String(status || '').trim().toLowerCase() === 'paused') continue;
    const name = pickField(row, 'Name', 'name', 'CHROName', 'chroName');
    const emails = emailsFromChroRow(row);
    if (!emails.length) {
      console.warn('approvedSitesMailer: CHRO row has no mail id', {
        rowId: row.ROWID || row.rowId || null,
        keys: Object.keys(row || {})
      });
    }
    emails.forEach((email) => {
      if (seen.has(email)) return;
      seen.add(email);
      recipients.push({ email, name });
    });
  }
  console.log(
    `approvedSitesMailer: CHRO recipients ${recipients.length} from ${rows.length} row(s)`,
    recipients.map((r) => r.email)
  );
  return recipients;
}

async function buildSiteRollup(catalyst, periodLabel) {
  const { year, month } = periodToYearMonth(periodLabel);
  if (!year || !month) {
    throw new Error(`Invalid report period: ${periodLabel}`);
  }
  const [rollup, pendingRollup, statutoryCounts, activeCount] = await Promise.all([
    buildChroRollupFromReports(catalyst, year, month),
    buildPendingRollupFromReports(catalyst, year, month),
    countStatutoryPendingAndReturned(catalyst, periodLabel),
    countActiveSitesFromSiteManagement(catalyst)
  ]);
  const approvedSites = (rollup.approvedSites || []).map((row) => ({
    ...row,
    siteKey: String(row.siteName || row.label || '').trim().toLowerCase(),
    siteName: row.siteName || row.label,
    label: formatSiteLabel(row.siteName || row.label, row.state),
    industry: formatStatutoryIndustry(row.industry),
    formNumber: String(row.formNumber || '').trim() || '—',
    formName: String(row.formName || '').trim() || '—'
  }));
  const pendingSites = (pendingRollup.pendingSites || []).map((row) => ({
    ...row,
    siteKey: String(row.siteName || row.label || '').trim().toLowerCase(),
    siteName: row.siteName || row.label,
    label: formatSiteLabel(row.siteName || row.label, row.state),
    industry: formatStatutoryIndustry(row.industry),
    formNumber: String(row.formNumber || '').trim() || '—',
    formName: String(row.formName || '').trim() || '—',
    status: formatPendingStatusLabel(row.status)
  }));
  const pendingCount = statutoryCounts.pendingCount;
  const returnedCount = statutoryCounts.returnedCount;
  const approvedCount = rollup.approvedCount;
  const totalForms = Number(rollup.reportRowCount) || 0;
  const complianceScore = totalForms ? Math.round((approvedCount / totalForms) * 100) : 0;
  console.log('approvedSitesMailer: reports rollup', {
    periodLabel,
    year,
    month,
    approvedCount,
    pendingCount,
    returnedCount,
    activeCount,
    totalForms,
    complianceScore,
    reportRowCount: rollup.reportRowCount,
    approvedSites: approvedSites.map((s) => `${s.label} / ${s.formName}`),
    pendingSites: pendingSites.map((s) => `${s.label} / ${s.formName} (${s.status})`)
  });
  return {
    ...rollup,
    approvedSites,
    pendingSites,
    pendingCount,
    returnedCount,
    activeCount,
    complianceScore
  };
}

async function sendApprovedSitesReportToChro(catalyst, options = {}) {
  const now = nowInKolkata();
  const periodLabel = options.periodLabel || currentPeriodLabel(now);
  const force = Boolean(options.force);
  const dryRun = Boolean(options.dryRun);
  const requireSendDay = options.requireMonthlySendDay === true;

  if (requireSendDay && !force && !isChroSendDayInKolkata(now)) {
    return {
      skipped: true,
      reason: `Not the ${CHRO_MONTHLY_SEND_DAY}th in Asia/Kolkata (today is ${now.day}). CHRO mail sends on the ${CHRO_MONTHLY_SEND_DAY}th.`,
      periodLabel,
      emailsSent: 0,
      emailsFailed: 0
    };
  }

  const recipients = await fetchChroRecipients(catalyst);
  const toEmails = recipients.map((r) => r.email).filter((email) => looksLikeEmail(email));
  if (!toEmails.length) {
    return {
      skipped: false,
      reason: 'No CHRO mail id found in the CHRO table. Save Email/Mail Id in Settings → CHRO Notification.',
      periodLabel,
      emailsSent: 0,
      emailsFailed: 1
    };
  }

  const rollup = await buildSiteRollup(catalyst, periodLabel);
  const chroName = recipients.find((r) => r.name)?.name || '';
  const html = buildApprovedSitesEmailHtml({
    chroName,
    periodLabel,
    ...rollup
  });
  const subject = `HR Compliance – Approved Sites & Pending Forms Report · ${periodLabel}`;

  const pdfPayload = {
    chroName,
    periodLabel,
    complianceScore: rollup.complianceScore,
    approvedSites: rollup.approvedSites,
    pendingSites: rollup.pendingSites
  };
  const pdfFileName = buildApprovedSitesPdfFileName(periodLabel);
  let pdfBuffer = null;
  try {
    pdfBuffer = await buildApprovedSitesPdfBuffer(pdfPayload);
  } catch (pdfErr) {
    console.error('approvedSitesMailer: PDF build failed:', pdfErr?.message || pdfErr);
    return {
      skipped: false,
      reason: pdfErr?.message || 'Failed to build approved sites PDF attachment.',
      periodLabel,
      emailsSent: 0,
      emailsFailed: 1,
      ...rollup
    };
  }

  if (dryRun) {
    return {
      skipped: false,
      dryRun: true,
      periodLabel,
      toEmails,
      emailsSent: toEmails.length,
      emailsFailed: 0,
      pdfFileName,
      pdfByteLength: pdfBuffer.length,
      ...rollup,
      html
    };
  }

  let fromEmail = '';
  try {
    fromEmail = await resolveFromEmail();
  } catch (err) {
    return {
      skipped: true,
      reason: err?.message || String(err),
      periodLabel,
      emailsSent: 0,
      emailsFailed: 0
    };
  }

  const tmpPath = path.join(os.tmpdir(), pdfFileName);
  fs.writeFileSync(tmpPath, pdfBuffer);

  const mailObj = {
    from_email: fromEmail,
    to_email: toEmails,
    subject,
    content: html,
    html_mode: true,
    display_name: DEFAULT_MAIL_DISPLAY_NAME,
    attachments: [fs.createReadStream(tmpPath)]
  };

  try {
    await catalyst.email().sendMail(mailObj);
    console.log('approvedSitesMailer: sent report to', toEmails.join(', '), {
      periodLabel,
      pdfFileName,
      pdfByteLength: pdfBuffer.length,
      approvedCount: rollup.approvedCount,
      pendingCount: rollup.pendingCount,
      returnedCount: rollup.returnedCount,
      activeCount: rollup.activeCount
    });
    return {
      skipped: false,
      fromEmail,
      periodLabel,
      toEmails,
      emailsSent: toEmails.length,
      emailsFailed: 0,
      pdfFileName,
      pdfByteLength: pdfBuffer.length,
      ...rollup
    };
  } catch (err) {
    console.error('approvedSitesMailer: send failed:', err?.message || err);
    return {
      skipped: false,
      fromEmail,
      periodLabel,
      toEmails,
      emailsSent: 0,
      emailsFailed: 1,
      error: err?.message || String(err),
      fromEmailConfigError: isFromEmailConfigError(err),
      ...rollup
    };
  } finally {
    try {
      fs.unlinkSync(tmpPath);
    } catch (_) {
      /* ignore temp cleanup */
    }
  }
}

function isNewlyApprovedRecord(previousRow, nextRow) {
  const isApproved = (row) => {
    const status = String(row?.Status || row?.status || '').trim().toLowerCase();
    const approval = String(row?.Approval || row?.approval || '').trim().toLowerCase();
    return status === 'approved' || status === 'approve' || approval === 'approved' || approval === 'approve';
  };
  return !isApproved(previousRow) && isApproved(nextRow);
}

module.exports = {
  sendApprovedSitesReportToChro,
  buildApprovedSitesEmailHtml,
  buildSiteRollup,
  fetchChroRecipients,
  isNewlyApprovedRecord,
  isChroSendDayInKolkata,
  CHRO_MONTHLY_SEND_DAY
};
