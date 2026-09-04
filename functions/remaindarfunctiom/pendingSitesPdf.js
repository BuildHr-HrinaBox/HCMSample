'use strict';

const http = require('http');
const https = require('https');
const PDFDocument = require('pdfkit');
const {
  groupSiteFormsByIndustry,
  getSiteIndustryLabels,
  sortSiteSections,
  normalizeIndustryLabel
} = require('./pendingSitesTable');

const DEFAULT_APP_BASE = 'https://hcm-60066164950.development.catalystserverless.in';

const PDF_FONT = {
  bannerTitle: 24,
  bannerSubtitle: 12,
  scoreValue: 20,
  scoreLabel: 10,
  greeting: 13,
  intro: 12,
  sectionTitle: 15,
  siteHeader: 12,
  tableHeader: 10,
  tableCell: 11,
  status: 11,
  empty: 12
};

function safeText(value) {
  return String(value || '').trim() || '—';
}

function resolveLogoUrl() {
  const fromEnv = String(process.env.LOGO_URL || '').trim();
  if (fromEnv) return fromEnv;
  const base = String(process.env.APP_BASE_URL || DEFAULT_APP_BASE).trim().replace(/\/+$/, '');
  return `${base}/server/sitemanagement_function/sitemanagement/email-logo`;
}

function fetchUrlBuffer(url, redirects = 0) {
  if (redirects > 5) return Promise.reject(new Error('Too many redirects fetching logo'));
  return new Promise((resolve, reject) => {
    const client = String(url).startsWith('https') ? https : http;
    client
      .get(url, (res) => {
        if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location) {
          res.resume();
          fetchUrlBuffer(res.headers.location, redirects + 1).then(resolve).catch(reject);
          return;
        }
        if (res.statusCode && res.statusCode >= 400) {
          res.resume();
          reject(new Error(`Logo fetch failed: HTTP ${res.statusCode}`));
          return;
        }
        const chunks = [];
        res.on('data', (chunk) => chunks.push(chunk));
        res.on('end', () => resolve(Buffer.concat(chunks)));
        res.on('error', reject);
      })
      .on('error', reject);
  });
}

function detectImageFormat(buffer) {
  if (!buffer || buffer.length < 4) return null;
  if (buffer[0] === 0x89 && buffer[1] === 0x50 && buffer[2] === 0x4e && buffer[3] === 0x47) return 'PNG';
  if (buffer[0] === 0xff && buffer[1] === 0xd8) return 'JPEG';
  return null;
}

function buildPendingSitesPdfFileName(periodLabel, level) {
  const safe = String(periodLabel || 'report')
    .replace(/[^\w\s-]/g, '')
    .trim()
    .replace(/\s+/g, '-');
  const levelPart = level ? `-Level-${level}` : '';
  return `HR-Compliance-Pending-Sites-Reminder${levelPart}-${safe || 'report'}.pdf`;
}

function buildColumnLayout(left, pageWidth) {
  const ratios = [0.12, 0.16, 0.11, 0.46, 0.15];
  const colWidths = ratios.map((ratio) => Math.floor(pageWidth * ratio));
  const used = colWidths.reduce((sum, width) => sum + width, 0);
  colWidths[colWidths.length - 1] += pageWidth - used;
  const colX = [left];
  for (let i = 0; i < colWidths.length - 1; i += 1) {
    colX.push(colX[i] + colWidths[i]);
  }
  return { colWidths, colX };
}

function measureRowHeight(doc, values, colWidths, fontName, fontSize) {
  doc.font(fontName).fontSize(fontSize);
  let maxHeight = 0;
  values.forEach((value, idx) => {
    const height = doc.heightOfString(String(value), {
      width: colWidths[idx] - 12,
      lineGap: 2
    });
    if (height > maxHeight) maxHeight = height;
  });
  return Math.max(24, Math.ceil(maxHeight) + 12);
}

function drawTableRow(doc, { rowY, rowH, values, colX, colWidths, left, pageWidth, isHeader, fillColor }) {
  if (fillColor) {
    doc.save();
    doc.rect(left, rowY, pageWidth, rowH).fill(fillColor);
    doc.restore();
  }

  doc.save();
  doc.lineWidth(0.5).strokeColor('#d1d5db');
  doc.rect(left, rowY, pageWidth, rowH).stroke();
  for (let i = 1; i < colX.length; i += 1) {
    doc.moveTo(colX[i], rowY).lineTo(colX[i], rowY + rowH).stroke();
  }
  doc.restore();

  values.forEach((value, idx) => {
    doc.save();
    if (isHeader) {
      doc.fillColor('#4b5563').font('Helvetica-Bold').fontSize(PDF_FONT.tableHeader);
    } else if (idx === 4) {
      doc.fillColor('#92400e').font('Helvetica-Bold').fontSize(PDF_FONT.status);
    } else if (idx === 3) {
      doc.fillColor('#111827').font('Helvetica').fontSize(PDF_FONT.tableCell);
    } else {
      doc.fillColor('#334155').font('Helvetica').fontSize(PDF_FONT.tableCell);
    }
    doc.text(String(value), colX[idx] + 6, rowY + 6, {
      width: colWidths[idx] - 12,
      height: rowH - 10,
      lineGap: 2,
      lineBreak: true,
      ellipsis: true
    });
    doc.restore();
  });
}

async function buildPendingSitesPdfBuffer({ recipientName, periodLabel, level, complianceScore, siteSections }) {
  let logoBuffer = null;
  let logoFormat = null;
  try {
    logoBuffer = await fetchUrlBuffer(resolveLogoUrl());
    logoFormat = detectImageFormat(logoBuffer);
  } catch (err) {
    console.warn('approvedSitesPdf: logo fetch failed:', err?.message || err);
  }

  return new Promise((resolve, reject) => {
    const doc = new PDFDocument({ size: 'A4', layout: 'landscape', margin: 36, autoFirstPage: true });
    const chunks = [];
    doc.on('data', (chunk) => chunks.push(chunk));
    doc.on('end', () => resolve(Buffer.concat(chunks)));
    doc.on('error', reject);

    const left = doc.page.margins.left;
    const pageWidth = doc.page.width - doc.page.margins.left - doc.page.margins.right;
    const pageBottom = doc.page.height - doc.page.margins.bottom;
    const { colWidths, colX } = buildColumnLayout(left, pageWidth);
    const headerLabels = ['State', 'Industry', 'Form Number', 'Form Name', 'Status'];
    const green = '#166534';
    const lightGreen = '#f0fdf4';

    let y = doc.page.margins.top;

    const startNewPage = () => {
      doc.addPage({ size: 'A4', layout: 'landscape', margin: 36 });
      y = doc.page.margins.top;
    };

    const ensureSpace = (needed) => {
      if (y + needed <= pageBottom) return;
      startNewPage();
    };

    if (logoBuffer && logoFormat) {
      try {
        const logoW = 150;
        const logoX = left + (pageWidth - logoW) / 2;
        doc.image(logoBuffer, logoX, y, { fit: [logoW, 42], align: 'center', valign: 'center' });
        y += 48;
      } catch (logoErr) {
        console.warn('approvedSitesPdf: logo embed failed:', logoErr?.message || logoErr);
      }
    }

    doc.save();
    doc.rect(left, y, pageWidth, 58).fill(green);
    doc.fillColor('#ffffff').font('Helvetica-Bold').fontSize(PDF_FONT.bannerTitle).text('HR Compliance', left + 12, y + 10);
    doc.font('Helvetica').fontSize(PDF_FONT.bannerSubtitle).text(
      `Pending Sites Reminder • Level ${safeText(level)} • ${safeText(periodLabel)}`,
      left + 12,
      y + 36
    );
    const scoreBoxW = 96;
    const scoreBoxX = left + pageWidth - scoreBoxW - 10;
    doc.roundedRect(scoreBoxX, y + 8, scoreBoxW, 42, 6).fill('#ffffff');
    doc.fillColor(green).font('Helvetica-Bold').fontSize(PDF_FONT.scoreValue).text(
      `${Number(complianceScore) || 0}%`,
      scoreBoxX,
      y + 14,
      { width: scoreBoxW, align: 'center' }
    );
    doc.fillColor('#6b7280').font('Helvetica').fontSize(PDF_FONT.scoreLabel).text('Compliance Score', scoreBoxX, y + 36, {
      width: scoreBoxW,
      align: 'center'
    });
    doc.restore();
    y += 68;

    doc.fillColor('#111827').font('Helvetica-Bold').fontSize(PDF_FONT.greeting).text(`Dear ${safeText(recipientName || 'Team')},`, left, y);
    y += 18;
    const intro = `Please find below the HR Compliance – Pending Sites Reminder for ${safeText(periodLabel)} (Level ${safeText(level)}).`;
    doc.fillColor('#374151').font('Helvetica').fontSize(PDF_FONT.intro).text(intro, left, y, { width: pageWidth, lineGap: 2 });
    y += doc.heightOfString(intro, { width: pageWidth, lineGap: 2 }) + 16;
    doc.fillColor(green).font('Helvetica-Bold').fontSize(PDF_FONT.sectionTitle).text('Pending Sites', left, y);
    y += 24;

    const drawColumnHeaderRow = () => {
      const headerH = 22;
      drawTableRow(doc, {
        rowY: y,
        rowH: headerH,
        values: headerLabels,
        colX,
        colWidths,
        left,
        pageWidth,
        isHeader: true,
        fillColor: '#f3f4f6'
      });
      y += headerH;
    };

    const sortedSections = sortSiteSections(siteSections);
    if (!sortedSections.length) {
      doc.fillColor('#6b7280').font('Helvetica').fontSize(PDF_FONT.empty).text(
        `No sites found in Site Management for ${safeText(periodLabel)}.`,
        left,
        y
      );
      doc.end();
      return;
    }

    for (const section of sortedSections) {
      const forms = section.forms || [];
      const siteLabel = section.label || section.siteName || '—';
      const industries = getSiteIndustryLabels(forms);
      const siteIndustryFallback = normalizeIndustryLabel(section.siteIndustry);
      const industryNote =
        industries.length > 1
          ? `Industries: ${industries.join(' · ')}`
          : industries.length === 1
            ? `Industry: ${industries[0]}`
            : siteIndustryFallback && siteIndustryFallback !== 'Others'
              ? `Industry: ${siteIndustryFallback}`
              : '';
      const industryGroups = groupSiteFormsByIndustry(forms);
      const multiIndustry = industryGroups.length > 1;

      ensureSpace(24 + (industryNote ? 14 : 0) + 22 + 24);
      doc.rect(left, y, pageWidth, 24).fill(lightGreen);
      doc.fillColor(green).font('Helvetica-Bold').fontSize(PDF_FONT.siteHeader).text(siteLabel, left + 8, y + 6, {
        width: pageWidth - 16
      });
      y += 24;
      if (industryNote) {
        doc.fillColor('#15803d').font('Helvetica-Bold').fontSize(10).text(industryNote, left + 8, y + 2, {
          width: pageWidth - 16
        });
        y += 14;
      }

      if (!forms.length) {
        ensureSpace(20);
        doc.fillColor('#6b7280').font('Helvetica').fontSize(PDF_FONT.empty).text(
          'No pending or yet to submit forms for this site.',
          left + 8,
          y + 4,
          { width: pageWidth - 16 }
        );
        y += 22;
        continue;
      }

      for (const [industryName, industryForms] of industryGroups) {
        if (multiIndustry) {
          ensureSpace(20);
          doc.rect(left, y, pageWidth, 18).fill('#ecfdf5');
          doc.fillColor('#047857').font('Helvetica-Bold').fontSize(11).text(industryName, left + 8, y + 4, {
            width: pageWidth - 16
          });
          y += 18;
        }

        drawColumnHeaderRow();

        for (const form of industryForms) {
          const values = [
            safeText(form.state),
            safeText(form.industry),
            safeText(form.formNumber),
            safeText(form.formName),
            safeText(form.status)
          ];
          const rowH = measureRowHeight(doc, values, colWidths, 'Helvetica', PDF_FONT.tableCell);
          ensureSpace(rowH);
          drawTableRow(doc, {
            rowY: y,
            rowH,
            values,
            colX,
            colWidths,
            left,
            pageWidth,
            isHeader: false,
            fillColor: null
          });
          y += rowH;
        }
      }
      y += 10;
    }

    doc.end();
  });
}

module.exports = {
  buildPendingSitesPdfBuffer,
  buildPendingSitesPdfFileName
};
