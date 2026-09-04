'use strict';

const fs = require('fs');
const os = require('os');
const path = require('path');
const {
  buildPendingRollupFromReports
} = require('./mainReportData');
const {
  resolveFromEmail,
  looksLikeEmail,
  escapeHtml,
  vayonaLogoHtml,
  currentPeriodLabel,
  DEFAULT_MAIL_DISPLAY_NAME
} = require('./pendingFormsMailer');
const { buildPendingSitesPdfBuffer, buildPendingSitesPdfFileName } = require('./pendingSitesPdf');
const {
  groupSiteFormsByIndustry,
  getSiteIndustryLabels,
  sortSiteSections,
  normalizeIndustryLabel,
  mergeDuplicateSiteSections
} = require('./pendingSitesTable');

const EMAIL_WIDTH_PX = 832;

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

function trimText(value) {
  return String(value ?? '').trim();
}

function normalizeEmail(value) {
  return trimText(value).toLowerCase();
}

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

function formatSiteLabel(name, state) {
  const abbr = stateAbbreviation(state);
  const label = String(name || '').trim();
  if (abbr && label && !label.toUpperCase().startsWith(`${abbr} `) && !label.toUpperCase().startsWith(`${abbr}–`)) {
    return `${abbr} – ${label}`;
  }
  return label || '—';
}

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

function statusBadge(status) {
  if (status === 'Pending') {
    return `<span style="display:inline-block;padding:4px 10px;border-radius:999px;background-color:#fef3c7;color:#92400e;font-family:Arial,Helvetica,sans-serif;font-size:12px;font-weight:700;">Pending</span>`;
  }
  return `<span style="display:inline-block;padding:4px 10px;border-radius:999px;background-color:#fee2e2;color:#991b1b;font-family:Arial,Helvetica,sans-serif;font-size:12px;font-weight:700;">Yet to Submit</span>`;
}

async function buildPendingSitesRollup(catalyst, periodLabel) {
  const { year, month } = periodToYearMonth(periodLabel);
  if (!year || !month) {
    throw new Error(`Invalid report period: ${periodLabel}`);
  }

  const rollup = await buildPendingRollupFromReports(catalyst, year, month);
  const mergedSections = mergeDuplicateSiteSections(rollup.siteSections || []);
  const siteSections = sortSiteSections(mergedSections).map((section) => ({
    ...section,
    label: formatSiteLabel(section.siteName, section.state),
    forms: (section.forms || []).map((row) => ({
      ...row,
      label: formatSiteLabel(row.siteName, row.state),
      industry: normalizeIndustryLabel(row.industry),
      status: row.status === 'Yet to Complete' ? 'Yet to Submit' : row.status
    }))
  }));
  const pendingSites = siteSections.flatMap((section) => section.forms);
  const sitesWithPending = siteSections.filter((section) => (section.forms || []).length > 0).length;
  const sitesAllComplete = siteSections.length - sitesWithPending;

  return {
    ...rollup,
    siteSections,
    pendingSites,
    sitesWithPending,
    sitesAllComplete
  };
}

function buildPendingFormsTableHtml(siteSections) {
  if (!Array.isArray(siteSections) || siteSections.length === 0) return '';

  const sortedSections = sortSiteSections(siteSections);

  const thStyle =
    'padding:10px 12px;font-family:Arial,Helvetica,sans-serif;font-size:12px;font-weight:700;color:#4b5563;';
  const siteHeaderStyle =
    'padding:14px 12px 10px;background-color:#f0fdf4;font-family:Arial,Helvetica,sans-serif;font-size:15px;font-weight:700;color:#166534;border-bottom:1px solid #dcfce7;';
  const industryHeaderStyle =
    'padding:10px 12px 8px;background-color:#ecfdf5;font-family:Arial,Helvetica,sans-serif;font-size:13px;font-weight:700;color:#047857;border-bottom:1px solid #d1fae5;';

  let html = '';
  sortedSections.forEach((section, groupIdx) => {
    const forms = section.forms || [];
    const siteLabel = section.label || section.siteName || '—';
    const industries = getSiteIndustryLabels(forms);
    const siteIndustryFallback = normalizeIndustryLabel(section.siteIndustry);
    const industryNote =
      industries.length > 1
        ? `<span style="display:block;font-size:12px;font-weight:600;color:#15803d;padding-top:4px;">Industries: ${industries.map((ind) => escapeHtml(ind)).join(' · ')}</span>`
        : industries.length === 1
          ? `<span style="display:block;font-size:12px;font-weight:600;color:#15803d;padding-top:4px;">Industry: ${escapeHtml(industries[0])}</span>`
          : siteIndustryFallback && siteIndustryFallback !== 'Others'
            ? `<span style="display:block;font-size:12px;font-weight:600;color:#15803d;padding-top:4px;">Industry: ${escapeHtml(siteIndustryFallback)}</span>`
            : '';
    const groupGap = groupIdx > 0 ? 'border-top:10px solid #ffffff;' : '';

    html += `<tr>
      <td colspan="5" style="${groupGap}${siteHeaderStyle}">${escapeHtml(siteLabel)}${industryNote}</td>
    </tr>`;

    if (!forms.length) {
      html += `<tr>
        <td colspan="5" style="padding:14px 12px;font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#6b7280;border-bottom:1px solid #edf2f0;">
          No pending or yet to submit forms for this site.
        </td>
      </tr>`;
      return;
    }

    const industryGroups = groupSiteFormsByIndustry(forms);
    const multiIndustry = industryGroups.length > 1;

    industryGroups.forEach(([industryName, industryForms], industryIdx) => {
      if (multiIndustry) {
        html += `<tr>
          <td colspan="5" style="${industryHeaderStyle}">${escapeHtml(industryName)}</td>
        </tr>`;
      }

      html += `<tr bgcolor="#f3f4f6">
        <th align="left" style="${thStyle}">State</th>
        <th align="left" style="${thStyle}">Industry</th>
        <th align="left" style="${thStyle}">Form Number</th>
        <th align="left" style="${thStyle}">Form Name</th>
        <th align="left" style="${thStyle}">Status</th>
      </tr>`;

      industryForms.forEach((form, formIdx) => {
        const isLastGroup = groupIdx === sortedSections.length - 1;
        const isLastIndustry = industryIdx === industryGroups.length - 1;
        const isLastForm = formIdx === industryForms.length - 1;
        const border = isLastGroup && isLastIndustry && isLastForm ? '0' : '1px solid #edf2f0';
        const cell = (content) =>
          `<td valign="middle" style="padding:12px;border-bottom:${border};">${content}</td>`;
        html += `<tr>
          ${cell(`<span style="font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#334155;">${escapeHtml(form.state)}</span>`)}
          ${cell(`<span style="font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#334155;">${escapeHtml(form.industry)}</span>`)}
          ${cell(`<span style="font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#334155;">${escapeHtml(form.formNumber)}</span>`)}
          ${cell(`<span style="font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#111827;">${escapeHtml(form.formName)}</span>`)}
          ${cell(statusBadge(form.status))}
        </tr>`;
      });
    });
  });
  return html;
}

function buildPendingSitesEmailHtml({ recipientName, periodLabel, level, complianceScore, siteSections }) {
  const greeting = recipientName ? escapeHtml(recipientName) : 'Team';
  const logoBlock = vayonaLogoHtml();
  const score = `${complianceScore}%`;
  const rows = buildPendingFormsTableHtml(siteSections);
  const emptyRow = `<tr><td colspan="5" style="padding:18px 12px;font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#6b7280;">No pending or yet to submit forms in Reports for ${escapeHtml(periodLabel)}.</td></tr>`;

  return `<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>HR Compliance Pending Sites Reminder</title>
</head>
<body style="margin:0;padding:0;background-color:#f4f7f6;">
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" bgcolor="#f4f7f6" style="background-color:#f4f7f6;">
    <tr>
      <td align="center" style="padding:24px 12px;">
        <table role="presentation" width="${EMAIL_WIDTH_PX}" cellpadding="0" cellspacing="0" border="0" bgcolor="#ffffff" style="width:${EMAIL_WIDTH_PX}px;max-width:${EMAIL_WIDTH_PX}px;background-color:#ffffff;border-radius:8px;">
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
                          <div style="font-size:14px;line-height:20px;color:#d1fae5;padding-top:4px;">Pending Sites Reminder &bull; Level ${escapeHtml(level)} &bull; ${escapeHtml(periodLabel)}</div>
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
                Please find below the HR Compliance &ndash; Pending Sites Reminder for <strong>${escapeHtml(periodLabel)}</strong> (Level ${escapeHtml(level)}).
              </p>
            </td>
          </tr>
          <tr>
            <td style="padding:18px 28px 8px;font-family:Arial,Helvetica,sans-serif;font-size:18px;font-weight:700;color:#166534;">Pending Sites</td>
          </tr>
          <tr>
            <td style="padding:0 24px 24px;">
              <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="border:1px solid #e5e7eb;border-radius:10px;">
                ${rows || emptyRow}
              </table>
            </td>
          </tr>
          <tr>
            <td style="padding:0 28px 28px;font-family:Arial,Helvetica,sans-serif;font-size:12px;color:#6b7280;">This is an automated reminder notification email.</td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
</body>
</html>`;
}

function recipientDisplayName(email) {
  const local = String(email || '').split('@')[0] || '';
  if (!local) return '';
  return local
    .split(/[._-]+/)
    .filter(Boolean)
    .map((part) => part.charAt(0).toUpperCase() + part.slice(1))
    .join(' ');
}

async function sendReminderEmails(catalyst, dueRows, now, options = {}) {
  const dryRun = Boolean(options.dryRun);
  const periodLabel = options.periodLabel || currentPeriodLabel(now);
  const validRows = (dueRows || []).filter((row) => looksLikeEmail(row.email));
  if (!validRows.length) {
    return { skipped: true, reason: 'No valid recipient email configured.', emailsSent: 0, emailsFailed: 0 };
  }

  const rollup = await buildPendingSitesRollup(catalyst, periodLabel);
  if (!rollup.siteSections?.length) {
    return {
      skipped: true,
      reason: 'No sites found in Site Management.',
      periodLabel,
      emailsSent: 0,
      emailsFailed: 0,
      ...rollup
    };
  }

  if (dryRun) {
    return {
      skipped: false,
      dryRun: true,
      periodLabel,
      emailsSent: validRows.length,
      emailsFailed: 0,
      recipients: validRows.map((row) => ({
        level: row.level,
        email: normalizeEmail(row.email),
        pendingForms: rollup.pendingSites.length
      })),
      ...rollup
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
      emailsFailed: validRows.length,
      ...rollup
    };
  }

  let emailsSent = 0;
  let emailsFailed = 0;
  const details = [];

  for (const row of validRows) {
    const toEmail = normalizeEmail(row.email);
    const level = Number(row.level) || 1;
    const recipientName = recipientDisplayName(toEmail);
    const html = buildPendingSitesEmailHtml({
      recipientName,
      periodLabel,
      level,
      complianceScore: rollup.complianceScore,
      siteSections: rollup.siteSections,
      sitesWithPending: rollup.sitesWithPending,
      activeCount: rollup.activeCount,
      sitesAllComplete: rollup.sitesAllComplete
    });
    const subject = `HR Compliance – Pending Sites Reminder · Level ${level} · ${periodLabel}`;
    const pdfFileName = buildPendingSitesPdfFileName(periodLabel, level);
    let tmpPath = '';

    try {
      const pdfBuffer = await buildPendingSitesPdfBuffer({
        recipientName,
        periodLabel,
        level,
        complianceScore: rollup.complianceScore,
        siteSections: rollup.siteSections
      });
      tmpPath = path.join(os.tmpdir(), pdfFileName);
      fs.writeFileSync(tmpPath, pdfBuffer);

      await catalyst.email().sendMail({
        from_email: fromEmail,
        to_email: [toEmail],
        subject,
        content: html,
        html_mode: true,
        display_name: DEFAULT_MAIL_DISPLAY_NAME,
        attachments: [fs.createReadStream(tmpPath)]
      });
      emailsSent += 1;
      details.push({
        level,
        email: toEmail,
        status: 'sent',
        pendingForms: rollup.pendingSites.length,
        pdfFileName
      });
    } catch (err) {
      emailsFailed += 1;
      details.push({
        level,
        email: toEmail,
        status: 'failed',
        error: err?.message || String(err)
      });
    } finally {
      if (tmpPath) {
        try {
          fs.unlinkSync(tmpPath);
        } catch (_) {
          /* ignore */
        }
      }
    }
  }

  return {
    skipped: false,
    fromEmail,
    periodLabel,
    emailsSent,
    emailsFailed,
    details,
    ...rollup
  };
}

module.exports = {
  sendReminderEmails,
  buildPendingSitesEmailHtml,
  buildPendingSitesRollup,
  looksLikeEmail,
  normalizeEmail,
  resolveFromEmail
};
