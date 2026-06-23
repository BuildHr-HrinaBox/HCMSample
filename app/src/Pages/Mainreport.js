import React, { useMemo, useState, useEffect, useCallback, useRef } from 'react';
import { Link, useNavigate, useSearchParams } from 'react-router-dom';
import { jsPDF } from 'jspdf';
import JSZip from 'jszip';
import './Mainreport.css';
import {
  checklistStateMatchesSiteState,
  fetchInchargeDisplayScopeFromSites,
} from '../utils/siteInchargeScope';

const API = '/server/mainreport_function/mainreport';

/** Same key as Statutory.js — Reports "Statutory" link primes Transaction month filter. */
const STATUTORY_MONTH_FILTER_KEY = 'statutory_month_filter';

/** Full month names - must match backend `normalizeMonth` / MainReport `Months` values */
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

/** @returns {(number | 'ellipsis')[]} */
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
    if (i > 0 && sorted[i] - sorted[i - 1] > 1) {
      out.push('ellipsis');
    }
    out.push(sorted[i]);
  }
  return out;
}

/** Same rule as Statutory STATUS + `mainreport_function`: only explicit transaction status Approved may expose drafts. */
const getStatutoryTransactionStatusNorm = (row) =>
  String(row?.statutoryStatus ?? row?.status ?? '')
    .trim()
    .toLowerCase();

const isStatutoryRowApproved = (row) => {
  const st = getStatutoryTransactionStatusNorm(row);
  if (st === 'approved' || st === 'approve') return true;
  const appr = String(row?.approval ?? row?.Approval ?? '')
    .trim()
    .toLowerCase();
  return appr === 'approved' || appr === 'approve';
};

const isStatutoryTransactionStatusApproved = (row) => isStatutoryRowApproved(row);

/** Strip draft URL if not approved (table UI); keep draftFileId for consolidated ZIP. */
const sanitizeMainReportApiRow = (row) => {
  if (!row || typeof row !== 'object') return row;
  if (isStatutoryTransactionStatusApproved(row)) return row;
  return { ...row, draftFileUrl: null, draftFile: null };
};

const isYetToCompleteStatusText = (stLower) =>
  stLower.includes('yet to complete') ||
  (stLower.includes('yet') && stLower.includes('complete')) ||
  stLower.includes('not started') ||
  stLower.includes('nostart') ||
  stLower.includes('incomplete') ||
  stLower.includes('not complete');

const rowHasStoredDraft = (row) => {
  if (row?.hasStatutoryDraftStored === true) return true;
  const id = row?.draftFileId ?? row?.draftFile;
  return id != null && String(id).trim() !== '' && String(id).toLowerCase() !== 'null';
};

/** Match Statutory.js `getStatutoryRowDisplayStatus` for KPI cards and table STATUS column. */
const getStatusLabel = (row) => {
  const hasDraftStored = rowHasStoredDraft(row);
  if (!hasDraftStored) return 'Yet to Complete';

  const st = String(row?.statutoryStatus || row?.status || '').trim();
  const stLower = st.toLowerCase();
  const appr = String(row?.approval || '').trim().toLowerCase();

  if (appr === 'rejected' || appr === 'reject' || stLower.includes('reject')) return 'Rejected';
  if (stLower === 'approved' || stLower === 'approve') return 'Approved';
  if (appr === 'approved' || appr === 'approve') return 'Approved';
  if (st === '' || st === '-' || st === '—') return 'Pending';
  if (stLower === 'pending') return 'Pending';
  if (isYetToCompleteStatusText(stLower)) return 'Yet to Complete';

  return st || 'Pending';
};

const getSectorGroupLabel = (row) => {
  const s = String(row?.sector || '').trim();
  return s || 'Others';
};

/** Checklist bulk: FormName = form number; Description = full form title. */
const getFormNumberLabel = (row) => String(row?.formName || row?.formNumber || '').trim() || '—';

const getFormNameLabel = (row) => {
  const desc = String(row?.description || '').trim();
  if (desc) return desc;
  return String(row?.formName || '').trim() || '—';
};

const getReportStateLabel = (row) => String(row?.state || row?.zState || '').trim() || '—';

const isPanIndiaStateLabel = (value) => {
  const blob = String(value || '')
    .trim()
    .toLowerCase();
  return /\b(all india|pan india|pan-india|national|central|all states|all state)\b/.test(blob);
};

/** Keep only rows belonging to the selected site (name + state). */
const rowMatchesSiteFilter = (row, siteName, siteState) => {
  if (!siteName) return true;
  const rowSite = String(row?.site || '')
    .trim()
    .toLowerCase();
  const wantSite = String(siteName || '')
    .trim()
    .toLowerCase();
  if (rowSite && rowSite !== wantSite) return false;
  if (!siteState) return true;
  const rowState = getReportStateLabel(row);
  if (!rowState || rowState === '—' || isPanIndiaStateLabel(rowState)) return true;
  return checklistStateMatchesSiteState(rowState, siteState);
};

const formatReportDateDisplay = (value) => {
  if (value == null || String(value).trim() === '') return '';
  const s = String(value).trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) {
    const d = new Date(`${s.substring(0, 10)}T00:00:00`);
    if (!Number.isNaN(d.getTime())) {
      return d.toLocaleDateString('en-IN', { day: '2-digit', month: 'short', year: 'numeric' });
    }
  }
  const d = new Date(s);
  if (!Number.isNaN(d.getTime())) {
    return d.toLocaleDateString('en-IN', { day: '2-digit', month: 'short', year: 'numeric' });
  }
  return s;
};

const getDraftDateLabel = (row) => {
  const raw =
    row?.submittedDate ??
    row?.SubmittedDate ??
    row?.draftDate ??
    row?.DraftDate ??
    '';
  const formatted = formatReportDateDisplay(raw);
  return formatted || '—';
};

const getApprovedDateLabel = (row) => {
  const raw =
    row?.approvedDate ??
    row?.ApprovedDate ??
    row?.approvalDate ??
    row?.ApprovalDate ??
    '';
  const formatted = formatReportDateDisplay(raw);
  return formatted || '—';
};

const sortRowsSectorWise = (rows) =>
  [...(Array.isArray(rows) ? rows : [])].sort((a, b) => {
    const sa = getSectorGroupLabel(a).toLowerCase();
    const sb = getSectorGroupLabel(b).toLowerCase();
    if (sa !== sb) return sa.localeCompare(sb);
    return String(a?.formName || '').localeCompare(String(b?.formName || ''), undefined, { sensitivity: 'base' });
  });

const CONSOLIDATED_SECTOR_FOLDERS = ['CLRA', 'Factories Act', 'Shops and Establishment', 'Others'];

/** ZIP folder name: CLRA | Factories Act | Shops and Establishment | Others */
const getConsolidatedSectorFolder = (row) => {
  const act = String(row?.act || '').toLowerCase();
  const sector = String(row?.sector || '').toLowerCase();
  const isFactories =
    act.includes('factories') ||
    act.includes('factory') ||
    sector.includes('factories') ||
    sector.includes('factory');
  if (isFactories) return 'Factories Act';
  const isShops =
    act.includes('shops and establishment') ||
    act.includes('shop and establishment') ||
    sector.includes('shops and establishment') ||
    sector.includes('shop and establishment');
  if (isShops) return 'Shops and Establishment';
  const isClra =
    act.includes('clra') ||
    act.includes('contract labour') ||
    act.includes('contract labor') ||
    sector === 'clra' ||
    sector.includes('clra') ||
    sector.includes('contract labour');
  if (isClra) return 'CLRA';
  const label = String(row?.sector || '').trim();
  return label || 'Others';
};

const getConsolidatedStateFolder = (row) => {
  const label = getReportStateLabel(row);
  const s = String(label === '—' ? 'Unknown' : label)
    .trim()
    .replace(/[\\/:*?"<>|]/g, '_')
    .replace(/\s+/g, ' ')
    .trim();
  return s.slice(0, 80) || 'Unknown';
};

/** Group report rows: State → Sector (CLRA / Factories Act / Shops and Establishment). */
const groupRowsByStateAndSector = (rows) => {
  const byState = new Map();
  for (const row of rows) {
    const stateKey = getConsolidatedStateFolder(row);
    const sectorKey = getConsolidatedSectorFolder(row);
    if (!byState.has(stateKey)) byState.set(stateKey, new Map());
    const bySector = byState.get(stateKey);
    if (!bySector.has(sectorKey)) bySector.set(sectorKey, []);
    bySector.get(sectorKey).push(row);
  }
  return byState;
};

const getDraftDownloadUrl = (row) => {
  if (row?.draftFileUrl) return row.draftFileUrl;
  const id = row?.draftFileId || row?.draftFile;
  if (!id) return null;
  const name = encodeURIComponent(String(row?.draftFileName || 'draft').trim() || 'draft');
  return `/server/mainreport_function/mainreport/file/${id}?name=${name}&disposition=attachment`;
};

const ensureUniqueFileName = (baseName, used) => {
  let name = String(baseName || 'file').trim() || 'file';
  if (!used.has(name)) {
    used.add(name);
    return name;
  }
  const dot = name.lastIndexOf('.');
  const stem = dot > 0 ? name.slice(0, dot) : name;
  const ext = dot > 0 ? name.slice(dot) : '';
  let n = 2;
  while (used.has(`${stem}_${n}${ext}`)) n += 1;
  name = `${stem}_${n}${ext}`;
  used.add(name);
  return name;
};

const countRowStatusBucket = (row) => {
  const status = getStatusLabel(row);
  if (status === 'Approved') return 'approved';
  if (status === 'Pending') return 'pending';
  return 'yet';
};

/** Rows for State × Sector summary table (Approved / Pending / Yet to Complete). */
const buildStateSectorSummaryData = (sourceRows) => {
  const list = Array.isArray(sourceRows) ? sourceRows : [];
  const keyMap = new Map();
  const entries = [];

  for (const row of list) {
    const state = getReportStateLabel(row);
    const sector = getConsolidatedSectorFolder(row);
    const key = `${state}\0${sector}`;
    if (!keyMap.has(key)) {
      const rec = { state, sector, approved: 0, pending: 0, yetToComplete: 0 };
      keyMap.set(key, rec);
      entries.push(rec);
    }
    const rec = keyMap.get(key);
    const bucket = countRowStatusBucket(row);
    if (bucket === 'approved') rec.approved += 1;
    else if (bucket === 'pending') rec.pending += 1;
    else rec.yetToComplete += 1;
  }

  const sectorRank = (s) => {
    const i = CONSOLIDATED_SECTOR_FOLDERS.indexOf(s);
    return i === -1 ? CONSOLIDATED_SECTOR_FOLDERS.length : i;
  };

  entries.sort((a, b) => {
    const st = a.state.localeCompare(b.state, undefined, { sensitivity: 'base' });
    if (st) return st;
    const sr = sectorRank(a.sector) - sectorRank(b.sector);
    if (sr) return sr;
    return a.sector.localeCompare(b.sector, undefined, { sensitivity: 'base' });
  });

  const totals = { approved: 0, pending: 0, yetToComplete: 0 };
  for (const e of entries) {
    totals.approved += e.approved;
    totals.pending += e.pending;
    totals.yetToComplete += e.yetToComplete;
  }

  return { entries, totals };
};

const formatPdfExportDate = (date = new Date()) =>
  date.toLocaleDateString(undefined, { year: 'numeric', month: 'short', day: '2-digit' });

/** Landscape PDF: KPI cards → State/Sector summary → detailed forms table. */
const buildReportPdfArrayBuffer = (sourceRows, { title, logoDataUrl }) => {
  const exportDate = new Date();
  const tableRowsForPdf = Array.isArray(sourceRows) ? sourceRows : [];
  const doc = new jsPDF({ unit: 'pt', format: 'a4', orientation: 'landscape' });
  const pageWidth = doc.internal.pageSize.getWidth();
  const pageHeight = doc.internal.pageSize.getHeight();

  let approvedPdf = 0;
  let pendingPdf = 0;
  let yetToCompletePdf = 0;
  for (const row of tableRowsForPdf) {
    const bucket = countRowStatusBucket(row);
    if (bucket === 'approved') approvedPdf += 1;
    else if (bucket === 'pending') pendingPdf += 1;
    else yetToCompletePdf += 1;
  }

  const boxes = [
    { label: 'Total Forms', count: tableRowsForPdf.length, color: [34, 197, 94] },
    { label: 'Approved', count: approvedPdf, color: [99, 102, 241] },
    { label: 'Pending', count: pendingPdf, color: [245, 158, 11] },
    { label: 'Yet to Complete', count: yetToCompletePdf, color: [234, 179, 8] }
  ];

  const startX = 32;
  const gap = 12;
  const boxW = (pageWidth - startX * 2 - gap * 3) / 4;
  const boxH = 72;
  const tableX = startX;
  const tableWidth = pageWidth - startX * 2;
  const rowH = 18;
  const headerDark = [55, 65, 81];
  const headerSub = [71, 85, 105];
  let tableY = 0;

  const drawTitleAndKpis = () => {
    if (logoDataUrl) {
      try {
        const logoFormat = logoDataUrl.includes('image/png')
          ? 'PNG'
          : logoDataUrl.includes('image/webp')
            ? 'WEBP'
            : 'JPEG';
        doc.addImage(logoDataUrl, logoFormat, startX, 16, 74, 26);
      } catch (_) {
        /* logo optional */
      }
    }
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(16);
    doc.setTextColor(17, 24, 39);
    doc.text(title, pageWidth / 2, 36, { align: 'center' });

    let y = 64;
    boxes.forEach((b, idx) => {
      const xPos = startX + idx * (boxW + gap);
      doc.setDrawColor(...b.color);
      doc.setFillColor(255, 255, 255);
      doc.roundedRect(xPos, y, boxW, boxH, 8, 8, 'FD');
      doc.setFont('helvetica', 'bold');
      doc.setFontSize(12);
      doc.setTextColor(31, 41, 55);
      doc.text(b.label, xPos + 12, y + 22);
      doc.setFontSize(22);
      doc.setTextColor(...b.color);
      doc.text(String(b.count), xPos + 12, y + 54);
    });
    return y + boxH + 14;
  };

  const drawCompactPageHeader = () => {
    if (logoDataUrl) {
      try {
        const logoFormat = logoDataUrl.includes('image/png')
          ? 'PNG'
          : logoDataUrl.includes('image/webp')
            ? 'WEBP'
            : 'JPEG';
        doc.addImage(logoDataUrl, logoFormat, startX, 16, 74, 26);
      } catch (_) {
        /* logo optional */
      }
    }
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(14);
    doc.setTextColor(17, 24, 39);
    doc.text(title, pageWidth / 2, 36, { align: 'center' });
    return 52;
  };

  const ensureSpace = (neededH, repeatDetailHeader) => {
    if (tableY + neededH <= pageHeight - 28) return;
    doc.addPage();
    tableY = drawCompactPageHeader();
    if (repeatDetailHeader) drawDetailTableHeader();
  };

  const summaryColW = [
    Math.floor(tableWidth * 0.06),
    Math.floor(tableWidth * 0.2),
    Math.floor(tableWidth * 0.26),
    Math.floor(tableWidth * 0.16),
    Math.floor(tableWidth * 0.16),
    Math.floor(tableWidth * 0.16)
  ];
  const summaryWTotal = summaryColW.reduce((a, b) => a + b, 0);
  if (summaryWTotal < tableWidth) summaryColW[2] += tableWidth - summaryWTotal;

  const drawStateSectorSummaryTable = (startY) => {
    const { entries, totals } = buildStateSectorSummaryData(tableRowsForPdf);
    if (!entries.length) return startY;

    let y = startY;
    const headH1 = 20;
    const headH2 = 18;
    const bodyH = 20;

    doc.setFillColor(...headerDark);
    doc.rect(tableX, y, tableWidth, headH1 + headH2, 'F');
    doc.setDrawColor(120, 120, 120);
    doc.setLineWidth(1);
    doc.rect(tableX, y, tableWidth, headH1 + headH2);

    let x = tableX;
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(9);
    doc.setTextColor(255, 255, 255);
    doc.text('S.No', x + 6, y + 13);
    x += summaryColW[0];
    doc.line(x, y, x, y + headH1 + headH2);
    doc.text('State', x + 6, y + 13);
    x += summaryColW[1];
    doc.line(x, y, x, y + headH1 + headH2);
    doc.text('Sector', x + 6, y + 13);
    x += summaryColW[2];
    doc.line(x, y, x, y + headH1 + headH2);
    const statusX = x;
    const statusW = summaryColW[3] + summaryColW[4] + summaryColW[5];
    doc.text('Status', statusX + statusW / 2, y + 13, { align: 'center' });
    doc.line(x, y + headH1, tableX + tableWidth, y + headH1);

    x = statusX;
    doc.setFillColor(...headerSub);
    doc.rect(statusX, y + headH1, statusW, headH2, 'F');
    ['Approved', 'Pending', 'Yet to Complete'].forEach((label, i) => {
      doc.text(label, x + summaryColW[3 + i] / 2, y + headH1 + 12, { align: 'center' });
      if (i < 2) {
        x += summaryColW[3 + i];
        doc.line(x, y + headH1, x, y + headH1 + headH2);
      }
    });

    y += headH1 + headH2;

    entries.forEach((entry, idx) => {
      doc.setDrawColor(180, 180, 180);
      doc.rect(tableX, y, tableWidth, bodyH);
      x = tableX;
      doc.setFont('helvetica', 'normal');
      doc.setFontSize(9);
      doc.setTextColor(55, 65, 81);
      doc.text(String(idx + 1), x + 6, y + 13);
      x += summaryColW[0];
      doc.line(x, y, x, y + bodyH);
      doc.text(entry.state, x + 6, y + 13);
      x += summaryColW[1];
      doc.line(x, y, x, y + bodyH);
      doc.text(entry.sector, x + 6, y + 13);
      x += summaryColW[2];
      doc.line(x, y, x, y + bodyH);
      const counts = [entry.approved, entry.pending, entry.yetToComplete];
      const countColors = [[22, 163, 74], [217, 119, 6], [234, 179, 8]];
      counts.forEach((c, i) => {
        doc.setTextColor(...countColors[i]);
        doc.text(String(c), x + summaryColW[3 + i] / 2, y + 13, { align: 'center' });
        if (i < 2) {
          x += summaryColW[3 + i];
          doc.line(x, y, x, y + bodyH);
        }
      });
      y += bodyH;
    });

    doc.setFillColor(243, 244, 246);
    doc.rect(tableX, y, tableWidth, bodyH, 'F');
    doc.setDrawColor(120, 120, 120);
    doc.rect(tableX, y, tableWidth, bodyH);
    x = tableX;
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(9);
    doc.setTextColor(31, 41, 55);
    doc.text('', x + 6, y + 13);
    x += summaryColW[0];
    doc.line(x, y, x, y + bodyH);
    doc.text('Grand Total', x + 6, y + 13);
    x += summaryColW[1] + summaryColW[2];
    doc.line(x, y, x, y + bodyH);
    const gt = [totals.approved, totals.pending, totals.yetToComplete];
    const gtColors = [[22, 163, 74], [217, 119, 6], [234, 179, 8]];
    gt.forEach((c, i) => {
      doc.setTextColor(...gtColors[i]);
      doc.text(String(c), x + summaryColW[3 + i] / 2, y + 13, { align: 'center' });
      if (i < 2) {
        x += summaryColW[3 + i];
        doc.line(x, y, x, y + bodyH);
      }
    });

    return y + bodyH + 16;
  };

  const detailHeader = ['S.NO', 'State', 'Sector', 'Act', 'Form Number', 'Form Name', 'Submitted Date', 'Approved Date', 'Status'];
  const colPct = [0.05, 0.1, 0.12, 0.2, 0.1, 0.18, 0.09, 0.09, 0.07];
  const colW = colPct.map((p) => Math.floor(tableWidth * p));
  const colWidthTotal = colW.reduce((a, b) => a + b, 0);
  if (colWidthTotal < tableWidth) colW[5] += tableWidth - colWidthTotal;

  const getCellLines = (value, colIndex) => {
    const text = String(value ?? '');
    if (colIndex === 0) return [text];
    const wrapped = doc.splitTextToSize(text || '-', colW[colIndex] - 10);
    return Array.isArray(wrapped) && wrapped.length > 0 ? wrapped : ['-'];
  };

  const drawDetailTableHeader = () => {
    doc.setFillColor(...headerDark);
    doc.rect(tableX, tableY, tableWidth, rowH, 'F');
    doc.setLineWidth(1.2);
    doc.setDrawColor(120, 120, 120);
    doc.rect(tableX, tableY, tableWidth, rowH);
    let x = tableX;
    let dividerX = tableX;
    for (let i = 0; i < colW.length - 1; i += 1) {
      dividerX += colW[i];
      doc.line(dividerX, tableY, dividerX, tableY + rowH);
    }
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(9);
    doc.setTextColor(255, 255, 255);
    detailHeader.forEach((h, i) => {
      doc.text(h, x + 6, tableY + 12);
      x += colW[i];
    });
    tableY += rowH;
  };

  tableY = drawTitleAndKpis();
  tableY = drawStateSectorSummaryTable(tableY);
  drawDetailTableHeader();

  const pdfTableRows = sortRowsSectorWise(tableRowsForPdf);
  const rows = pdfTableRows.length
    ? pdfTableRows.map((row, idx) => [
        String(idx + 1),
        getReportStateLabel(row),
        getConsolidatedSectorFolder(row),
        String(row.act || '-'),
        getFormNumberLabel(row),
        getFormNameLabel(row),
        getDraftDateLabel(row),
        getApprovedDateLabel(row),
        getStatusLabel(row)
      ])
    : [['', '', '', '', 'No rows for this selection.', '', '', '', '']];

  rows.forEach((r) => {
    const lineHeight = 10;
    const rowPaddingTop = 12;
    const rowPaddingBottom = 4;
    const cellLinesByCol = r.map((cell, i) => getCellLines(cell, i));
    const maxLines = Math.max(...cellLinesByCol.map((lines) => lines.length));
    const dynamicRowH = Math.max(rowH, rowPaddingTop + (maxLines - 1) * lineHeight + rowPaddingBottom);

    ensureSpace(dynamicRowH, true);

    let x = tableX;
    doc.setLineWidth(1.1);
    doc.setDrawColor(130, 130, 130);
    doc.rect(tableX, tableY, tableWidth, dynamicRowH);
    let dividerX = tableX;
    for (let i = 0; i < colW.length - 1; i += 1) {
      dividerX += colW[i];
      doc.line(dividerX, tableY, dividerX, tableY + dynamicRowH);
    }
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(8);
    r.forEach((cell, i) => {
      if (i === 8) {
        const sl = String(cell || '').toLowerCase();
        if (sl.includes('approved')) doc.setTextColor(22, 163, 74);
        else if (sl.includes('yet')) doc.setTextColor(234, 179, 8);
        else if (sl.includes('pending')) doc.setTextColor(217, 119, 6);
        else doc.setTextColor(75, 85, 99);
      } else doc.setTextColor(75, 85, 99);
      cellLinesByCol[i].forEach((line, lineIdx) => {
        doc.text(String(line), x + 6, tableY + rowPaddingTop + lineIdx * lineHeight, { maxWidth: colW[i] - 10 });
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

const toIntOrNull = (v) => {
  const n = Number.parseInt(String(v || ''), 10);
  return Number.isFinite(n) ? n : null;
};

const monthIndexFromName = (name) => {
  const idx = ALL_MONTHS.findIndex((m) => m.toLowerCase() === String(name || '').toLowerCase());
  return idx >= 0 ? idx : null;
};

const clamp = (n, min, max) => Math.min(max, Math.max(min, n));

const formatMaybeDate = (value) => {
  if (!value) return '';
  const d = new Date(value);
  if (Number.isNaN(d.getTime())) return String(value);
  return d.toLocaleString(undefined, { year: 'numeric', month: 'short', day: '2-digit' });
};

const getRowUpdatedAt = (row) =>
  row?.lastUpdated || row?.updatedAt || row?.updated_at || row?.modifiedAt || row?.createdAt || null;

const Mainreport = ({ userEmail: userEmailProp }) => {
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
  const [reportSiteState, setReportSiteState] = useState('');
  const [inchargeSiteNames, setInchargeSiteNames] = useState(null);
  const [tableLoading, setTableLoading] = useState(false);
  const [consolidatedZipLoading, setConsolidatedZipLoading] = useState(false);
  const selectedYear = yearFromUrl ? String(yearFromUrl).trim() : '';
  const selectedMonth = monthFromUrl ? String(monthFromUrl).trim() : '';
  const selectedSite = siteFromUrl ? String(siteFromUrl).trim() : '';

  // Filter UI state (Apply/Reset like screenshot)
  const [draftYear, setDraftYear] = useState(selectedYear);
  const [draftMonth, setDraftMonth] = useState(selectedMonth);
  const [draftSite, setDraftSite] = useState(selectedSite);
  const reportSiteForStatutory = (selectedSite || draftSite || '').trim();
  const reportMonthForStatutory = (selectedMonth || draftMonth || '').trim();

  const statutoryTransactionHref = useMemo(() => {
    if (!reportSiteForStatutory) return '/rule-book/statutory';
    return `/rule-book/statutory?site=${encodeURIComponent(reportSiteForStatutory)}`;
  }, [reportSiteForStatutory]);

  const primeStatutoryMonthFilter = useCallback((fullMonthName) => {
    const m = String(fullMonthName || '').trim();
    if (m && typeof localStorage !== 'undefined') {
      localStorage.setItem(STATUTORY_MONTH_FILTER_KEY, m);
    }
  }, []);
  // Pagination state
  const [currentPage, setCurrentPage] = useState(1);
  const PAGE_SIZE = 10;

  const calendarYearStr = useMemo(() => String(new Date().getFullYear()), []);

  /** Always include URL/draft year + current calendar year so the <select> value matches an <option> before /summary returns (avoids year “jumping”). */
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

  const loadSummary = useCallback(async () => {
    setError('');
    setLoading(true);
    try {
      const res = await fetch(`${API}/summary`, { cache: 'no-store' });
      const json = await res.json();
      if (json.status !== 'success' || !json.data) {
        throw new Error(json.message || 'Invalid response');
      }
      const { years: yList } = json.data;
      setYears(Array.isArray(yList) ? yList : []);
    } catch (e) {
      setError(e.message || 'Failed to load main report');
      setYears([]);
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    loadSummary();
  }, [loadSummary]);

  useEffect(() => {
    let cancelled = false;
    fetchInchargeDisplayScopeFromSites(effectiveUserEmail).then((scope) => {
      if (cancelled) return;
      setInchargeSiteNames(Array.isArray(scope?.siteNames) ? scope.siteNames : []);
    });
    return () => {
      cancelled = true;
    };
  }, [effectiveUserEmail]);

  const inchargeScopeReady = inchargeSiteNames !== null;
  const isSiteInchargeView = inchargeScopeReady && inchargeSiteNames.length > 0;
  const requireSiteSelection = isSiteInchargeView;

  const siteSelectOptions = useMemo(() => {
    const fromApi = Array.isArray(siteNames) ? siteNames : [];
    if (!isSiteInchargeView) return fromApi;
    const allowed = new Set(inchargeSiteNames.map((s) => String(s).trim().toLowerCase()));
    return fromApi.filter((s) => allowed.has(String(s).trim().toLowerCase()));
  }, [siteNames, inchargeSiteNames, isSiteInchargeView]);

  const loadEntries = useCallback(async (year, month, siteValue = '') => {
    setTableLoading(true);
    setError('');
    try {
      const q = new URLSearchParams({
        year: String(year),
        month: String(month),
        userEmail: effectiveUserEmail,
        site: String(siteValue || '')
      });
      const res = await fetch(`${API}/entries?${q}`, { cache: 'no-store' });
      const json = await res.json();
      if (json.status !== 'success' || !json.data) {
        throw new Error(json.message || 'Invalid response');
      }
      const rawRows = Array.isArray(json.data.rows) ? json.data.rows : [];
      setTableRows(rawRows.map(sanitizeMainReportApiRow));
      setSiteNames(Array.isArray(json.data.siteNames) ? json.data.siteNames : []);
      setReportSiteState(String(json.data.selectedSiteState || '').trim());
    } catch (e) {
      setError(e.message || 'Failed to load rows');
      setTableRows([]);
      setSiteNames([]);
      setReportSiteState('');
    } finally {
      setTableLoading(false);
    }
  }, [effectiveUserEmail]);

  /** Site dropdown options only — does not load table/KPI data until Apply Filters. */
  const loadSiteNamesForDraft = useCallback(
    async (year, month) => {
      try {
        const q = new URLSearchParams({
          year: String(year),
          month: String(month),
          userEmail: effectiveUserEmail,
          site: ''
        });
        const res = await fetch(`${API}/entries?${q}`, { cache: 'no-store' });
        const json = await res.json();
        if (json.status !== 'success' || !json.data) return;
        setSiteNames(Array.isArray(json.data.siteNames) ? json.data.siteNames : []);
      } catch (_) {
        setSiteNames([]);
      }
    },
    [effectiveUserEmail]
  );

  const hasAppliedFilters = Boolean(
    selectedYear && selectedMonth && inchargeScopeReady && (!requireSiteSelection || selectedSite)
  );

  // Site Incharge with one assigned site: auto-apply site filter when year/month are set.
  useEffect(() => {
    if (!inchargeScopeReady || !selectedYear || !selectedMonth || selectedSite) return;
    if (inchargeSiteNames.length !== 1) return;
    const q = new URLSearchParams();
    q.set('year', selectedYear);
    q.set('month', selectedMonth);
    q.set('site', inchargeSiteNames[0]);
    navigate(`/mainreport?${q.toString()}`, { replace: true });
  }, [inchargeScopeReady, selectedYear, selectedMonth, selectedSite, inchargeSiteNames, navigate]);

  // Load report rows only after Apply Filters (URL has year + month [+ site for incharge]).
  useEffect(() => {
    if (!hasAppliedFilters) {
      setTableRows([]);
      setReportSiteState('');
      return;
    }
    loadEntries(selectedYear, selectedMonth, String(selectedSite || '').trim());
  }, [hasAppliedFilters, selectedYear, selectedMonth, selectedSite, loadEntries]);

  // Populate site dropdown while filters are being chosen (no table/KPI data yet).
  useEffect(() => {
    if (hasAppliedFilters) return;
    if (!draftYear || !draftMonth) {
      setSiteNames([]);
      return;
    }
    loadSiteNamesForDraft(draftYear, draftMonth);
  }, [hasAppliedFilters, draftYear, draftMonth, loadSiteNamesForDraft]);

  // Keep filter controls synced when URL changes (sidebar nav, back/forward, month pills).
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
    if (requireSiteSelection && !s) {
      setError('Select a site to view the report for that site only.');
      return;
    }
    setError('');

    const q = new URLSearchParams();
    q.set('year', y);
    q.set('month', m);
    if (s) q.set('site', s);
    navigate(`/mainreport?${q.toString()}`);
  }, [draftYear, draftMonth, draftSite, navigate, requireSiteSelection]);

  const resetFilters = useCallback(() => {
    setDraftYear('');
    setDraftMonth('');
    setDraftSite('');
    navigate('/mainreport');
    setTableRows([]);
    setSiteNames([]);
  }, [navigate]);

  const onMonthPillClick = (monthName) => {
    if (!draftYear) return;
    setDraftMonth(monthName);
    setDraftSite('');
    const q = new URLSearchParams();
    q.set('year', String(draftYear));
    q.set('month', String(monthName));
    navigate(`/mainreport?${q.toString()}`);
    setSiteNames([]);
  };

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

  const downloadConsolidatedZip = useCallback(async () => {
    if (!hasAppliedFilters || !tableRows.length) {
      setError('Apply filters and load report data before downloading.');
      return;
    }
    setConsolidatedZipLoading(true);
    setError('');
    try {
      const logoDataUrl = await getReportLogoDataUrl();
      const siteSuffix = selectedSite ? ` (${selectedSite})` : ' (All sites)';
      const periodLabel = `${selectedMonth} ${selectedYear}${siteSuffix}`;

      const byState = groupRowsByStateAndSector(tableRows);
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

          const pdfTitle = `${stateFolder} — ${sectorFolder} - ${periodLabel}`;
          const pdfBuf = buildReportPdfArrayBuffer(sectorRows, { title: pdfTitle, logoDataUrl });
          sectorDir.file(`${sectorFolder.replace(/\s+/g, '_')}_Report.pdf`, pdfBuf, { binary: true });

          const usedNames = new Set();
          for (const row of sectorRows) {
            const draftUrl = getDraftDownloadUrl(row);
            if (!draftUrl) continue;
            try {
              const res = await fetch(draftUrl, { cache: 'no-store', credentials: 'include' });
              if (!res.ok) continue;
              const blob = await res.blob();
              const baseName = String(row.draftFileName || row.formName || 'draft').trim() || 'draft';
              sectorDir.file(ensureUniqueFileName(baseName, usedNames), blob);
            } catch (_) {
              /* skip missing draft */
            }
          }

          foldersAdded += 1;
        }
      }

      if (!foldersAdded) throw new Error('No state/sector data found for the current filters.');

      const monthLabel = String(selectedMonth || '').replace(/\s+/g, '_');
      const siteLabel = String(selectedSite || 'All').replace(/\s+/g, '_');
      const fname = `Consolidated_Report_${selectedYear}_${monthLabel}_${siteLabel}.zip`;
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
      setError(e.message || 'Failed to download consolidated ZIP.');
    } finally {
      setConsolidatedZipLoading(false);
    }
  }, [hasAppliedFilters, selectedYear, selectedMonth, selectedSite, tableRows, getReportLogoDataUrl]);

  const downloadConsolidatedPdf = async () => {
    if (!hasAppliedFilters || !tableRows.length) {
      setError('Apply filters and load report data before downloading.');
      return;
    }
    const logoDataUrl = await getReportLogoDataUrl();
    const title = `Consolidated Report - ${selectedMonth || draftMonth || ''} ${selectedYear || draftYear || ''}${selectedSite || draftSite ? ` (${selectedSite || draftSite})` : ' (All sites)'}`;
    const buf = buildReportPdfArrayBuffer(tableRows, { title, logoDataUrl });
    const fname = `Consolidated_Report_${selectedYear || draftYear || ''}_${selectedMonth || draftMonth || ''}_${selectedSite || draftSite || 'All'}.pdf`;
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
  };

  const filteredRows = useMemo(() => {
    let rows = sortRowsSectorWise(tableRows);
    if (selectedSite) {
      rows = rows.filter((row) => rowMatchesSiteFilter(row, selectedSite, reportSiteState));
    }
    return rows;
  }, [tableRows, selectedSite, reportSiteState]);

  /**
   * KPIs match the Forms Summary STATUS column (`getStatusLabel`) for visible rows.
   */
  const statusCounts = useMemo(() => {
    if (!Array.isArray(filteredRows) || filteredRows.length === 0) {
      return { total: 0, approved: 0, pending: 0, yetToComplete: 0 };
    }
    let approved = 0;
    let pending = 0;
    let yetToComplete = 0;
    for (const row of filteredRows) {
      const status = getStatusLabel(row);
      if (status === 'Approved') approved += 1;
      else if (status === 'Pending') pending += 1;
      else if (status === 'Yet to Complete') yetToComplete += 1;
    }
    const total = filteredRows.length;
    return { total, approved, pending, yetToComplete };
  }, [filteredRows]);

  // Pagination: slice filteredRows for current page
  const totalPages = Math.max(1, Math.ceil(filteredRows.length / PAGE_SIZE));
  const effectivePage = Math.min(currentPage, totalPages);
  const pagedRows = filteredRows.slice((effectivePage - 1) * PAGE_SIZE, effectivePage * PAGE_SIZE);
  const paginationItems = useMemo(
    () => buildPaginationItems(effectivePage, totalPages),
    [effectivePage, totalPages]
  );

  // Reset to page 1 when filters change
  useEffect(() => {
    setCurrentPage(1);
  }, [filteredRows]);

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
                <h1 className="mr-title">Reports</h1>
              </div>
            </div>
          </header>

          {error ? <div className="mr-alert">{error}</div> : null}

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
                name="mainreport-year"
                value={draftYear}
                required
                aria-required="true"
                onChange={(e) => {
                  const y = e.target.value;
                  setDraftYear(y);
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
                name="mainreport-month"
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
              <span className="mr-label">
                Site
                {requireSiteSelection ? (
                  <abbr className="mr-required" title="Required">*</abbr>
                ) : null}
              </span>
              <select
                className="mr-select"
                value={draftSite}
                onChange={(e) => setDraftSite(e.target.value)}
                disabled={!draftYear || !draftMonth}
                required={requireSiteSelection}
                aria-required={requireSiteSelection ? 'true' : undefined}
              >
                {!requireSiteSelection ? <option value="">All</option> : (
                  <option value="">Select site</option>
                )}
                {siteSelectOptions.map((s) => (
                  <option key={s} value={s}>{s}</option>
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
                disabled={!draftYear || !draftMonth || (requireSiteSelection && !draftSite)}
                title={
                  !draftYear || !draftMonth
                    ? 'Select year and month (required).'
                    : requireSiteSelection && !draftSite
                      ? 'Select a site to view site-specific report.'
                      : 'Apply filters'
                }
              >
                Apply Filters
              </button>
            </div>
          </div>
          </form>
          <div className="mr-filter-hint" />
          </section>

          <section className="mr-kpis" aria-label="Summary">
          <div className="mr-kpi mr-kpi--forms">
            <div className="mr-kpi-label">Total Forms</div>
            <div className="mr-kpi-value">{statusCounts.total}</div>
          </div>
          <div className="mr-kpi mr-kpi--drafts">
            <div className="mr-kpi-label">Approved</div>
            <div className="mr-kpi-value">{statusCounts.approved}</div>
          </div>
          <div
            className="mr-kpi mr-kpi--pending"
            title="Forms that are in Pending status."
          >
            <div className="mr-kpi-label">Pending</div>
            <div className="mr-kpi-value">{statusCounts.pending}</div>
          </div>
          <div className="mr-kpi mr-kpi--yet" title="Forms that are Yet to Complete.">
            <div className="mr-kpi-label">Yet to Complete</div>
            <div className="mr-kpi-value">{statusCounts.yetToComplete}</div>
          </div>
          </section>


          <section className="mr-grid">
          <div className="mr-left">
            <section className="mr-card mr-table-card" aria-label="Forms Summary">
              <div className="mr-table-head">
                <div>
                  <div className="mr-table-title">Forms Summary</div>
                  {selectedSite ? (
                    <div className="mr-table-subtitle">Showing forms for site: {selectedSite}</div>
                  ) : null}
                </div>
                <div className="mr-table-actions">
                  <button
                    type="button"
                    className="mr-seg mr-seg--active"
                    onClick={downloadConsolidatedZip}
                    disabled={!hasAppliedFilters || consolidatedZipLoading}
                    title={
                      !hasAppliedFilters
                        ? 'Apply filters first'
                        : consolidatedZipLoading
                          ? 'Preparing ZIP…'
                          : 'ZIP by State → Sector (CLRA, Factories Act, Shops and Establishment) with PDF in each sector folder'
                    }
                  >
                    {consolidatedZipLoading ? 'Preparing ZIP…' : 'Consolidated Report'}
                  </button>
                  <button
                    type="button"
                    className="mr-btn mr-btn-primary"
                    onClick={downloadConsolidatedPdf}
                    disabled={!hasAppliedFilters}
                    title={!hasAppliedFilters ? 'Apply filters to load the report first' : 'Download PDF'}
                  >
                    Download Report
                  </button>
                </div>
              </div>

              <div className="mr-table-wrap">
                <table className="mr-table">
                  <thead>
                    <tr>
                      <th style={{ width: 70 }}>S.NO</th>
                      <th style={{ width: 120, minWidth: 120 }}>State</th>
                      <th style={{ width: 180, minWidth: 180 }}>Sector</th>
                      <th style={{ maxWidth: 260, width: 260 }}>Act</th>
                      <th style={{ maxWidth: 120, width: 120 }}>Form Number</th>
                      <th style={{ maxWidth: 280, width: 280 }}>Form Name</th>
                      <th style={{ width: 110, minWidth: 110 }}>Submitted Date</th>
                      <th style={{ width: 110, minWidth: 110 }}>Approved Date</th>
                      <th style={{ width: 180 }}>Status</th>
                    </tr>
                  </thead>
                  <tbody>
                    {!hasAppliedFilters ? (
                      <tr>
                        <td colSpan={9} className="mr-empty">
                          {requireSiteSelection
                            ? 'Select year, month, and your site, then click Apply Filters to load the site report.'
                            : 'Select year, month, and optionally a site, then click Apply Filters to load the report.'}
                        </td>
                      </tr>
                    ) : tableLoading && filteredRows.length === 0 ? (
                      <tr aria-hidden="true">
                        <td colSpan={9} className="mr-table-placeholder" />
                      </tr>
                    ) : filteredRows.length === 0 ? (
                      <tr>
                        <td colSpan={9} className="mr-empty">No rows for this selection.</td>
                      </tr>
                    ) : (
                      pagedRows.map((row, idx) => {
                        const seq = (effectivePage - 1) * PAGE_SIZE + idx + 1;
                        const key = row?.rowId != null ? String(row.rowId) : `${(effectivePage - 1) * PAGE_SIZE + idx}`;
                        const status = getStatusLabel(row);
                        const statusClass =
                          status === 'Approved'
                            ? 'mr-status--ok'
                            : status === 'Yet to Complete'
                              ? 'mr-status--yet'
                              : 'mr-status--pending';
                        const formNumber = getFormNumberLabel(row);
                        const formName = getFormNameLabel(row);
                        const stateLabel = getReportStateLabel(row);
                        const draftDateLabel = getDraftDateLabel(row);
                        const approvedDateLabel = getApprovedDateLabel(row);
                        return (
                          <tr key={key}>
                            <td>{seq}</td>
                            <td
                              style={{ width: 120, minWidth: 120, whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}
                              title={stateLabel !== '—' ? stateLabel : ''}
                            >
                              {stateLabel}
                            </td>
                            <td style={{ width: 180, minWidth: 180, whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>
                              {row?.sector || '—'}
                            </td>
                            <td
                              className="mr-act-col"
                              style={{ maxWidth: 260, width: 260, whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}
                              title={row?.act || ''}
                            >
                              {row?.act || ''}
                            </td>
                            <td
                              style={{ maxWidth: 120, width: 120, whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}
                              title={formNumber !== '—' ? formNumber : ''}
                            >
                              {formNumber}
                            </td>
                            <td
                              className="mr-strong"
                              style={{
                                maxWidth: 280,
                                width: 280,
                                overflow: 'hidden',
                                textOverflow: 'ellipsis',
                                whiteSpace: 'normal',
                                display: '-webkit-box',
                                WebkitLineClamp: 2,
                                WebkitBoxOrient: 'vertical'
                              }}
                              title={formName !== '—' ? formName : ''}
                            >
                              {formName}
                            </td>
                            <td style={{ whiteSpace: 'nowrap' }} title={draftDateLabel !== '—' ? draftDateLabel : ''}>
                              {draftDateLabel}
                            </td>
                            <td style={{ whiteSpace: 'nowrap' }} title={approvedDateLabel !== '—' ? approvedDateLabel : ''}>
                              {approvedDateLabel}
                            </td>
                            <td>
                              <span className={`mr-status ${statusClass}`}>{status}</span>
                            </td>
                          </tr>
                        );
                      })
                    )}
                  </tbody>
                </table>
              </div>
              {filteredRows.length > PAGE_SIZE ? (
                <nav className="mr-pagination" aria-label="Table pagination">
                  <button
                    type="button"
                    className="mr-pagination-nav"
                    disabled={effectivePage <= 1}
                    onClick={() => setCurrentPage(1)}
                    title="First page"
                    aria-label="First page"
                  >
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden>
                      <polyline points="11 17 6 12 11 7" />
                      <polyline points="18 17 13 12 18 7" />
                    </svg>
                  </button>
                  <button
                    type="button"
                    className="mr-pagination-nav"
                    disabled={effectivePage <= 1}
                    onClick={() => setCurrentPage((p) => Math.max(1, p - 1))}
                    title="Previous page"
                    aria-label="Previous page"
                  >
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden>
                      <polyline points="15 18 9 12 15 6" />
                    </svg>
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
                          aria-label={`Page ${item}`}
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
                    title="Next page"
                    aria-label="Next page"
                  >
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden>
                      <polyline points="9 18 15 12 9 6" />
                    </svg>
                  </button>
                  <button
                    type="button"
                    className="mr-pagination-nav"
                    disabled={effectivePage >= totalPages}
                    onClick={() => setCurrentPage(totalPages)}
                    title="Last page"
                    aria-label="Last page"
                  >
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden>
                      <polyline points="13 17 18 12 13 7" />
                      <polyline points="6 17 11 12 6 7" />
                    </svg>
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

export default Mainreport;
