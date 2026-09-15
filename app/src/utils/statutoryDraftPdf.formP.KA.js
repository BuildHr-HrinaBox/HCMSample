/**
 * Karnataka Form P — Notice of Maximum Leave Accumulated (See rule 20).
 * Excel is a letter (titles + To/Shri/Smt + legal paragraph + small leave table).
 * The generic PDF grid stacked every line as a table row and moved establishment
 * under "Details of the leave accumulated".
 */

export const FORM_P_KA_PDF_TITLE = "Form – 'P'";
export const FORM_P_KA_PDF_RULE = '(See rule 20)';
export const FORM_P_KA_PDF_SUBTITLE = 'NOTICE OF MAXIMUM LEAVE ACCUMULATED';
export const FORM_P_KA_PDF_DETAILS = 'Details of the leave accumulated';
export const FORM_P_KA_PDF_LEGAL = [
  'It is hereby informed that as per section 18 (5) of the Maharashtra Shops and Establishments',
  '(Regulation of Employment and Conditions of Service) Act, 2017 (Mah. LXI of 2017) the',
  'maximum leave that can be accumulated is for 45 days. Maximum leave of 45 days has been',
  'accumulated at your credit. Hence, no further leave due to you, but not availed by you will not',
  'be accumulated and it shall lapse, if unavailed.',
].join(' ');

const MIN_LEAVE_BODY_ROWS = 3;

const norm = (value) =>
  String(value ?? '')
    .replace(/_x000d_/gi, ' ')
    .replace(/\u000d/g, ' ')
    .replace(/\r?\n/g, ' ')
    .replace(/\u00a0/g, ' ')
    .replace(/\[object Object\]/gi, '')
    .replace(/[ \t]+/g, ' ')
    .trim();

const lower = (value) => norm(value).toLowerCase();

const blobFrom = (metaLines, rows, sheetName = '', fileName = '') =>
  [...(metaLines || []), ...(rows || []).slice(0, 24).flat(), sheetName || '', fileName || '']
    .map((t) => lower(t))
    .join(' ');

export function looksLikeFormPKarnatakaPdfContext(
  metaLines = [],
  rows = [],
  sheetName = '',
  fileName = ''
) {
  const blob = blobFrom(metaLines, rows, sheetName, fileName);
  if (!blob && !sheetName && !fileName) return false;
  const nameBlob = `${sheetName || ''} ${fileName || ''}`.toLowerCase();
  if (/muster[\s-]*roll/.test(blob) || /date\s+of\s+(the\s+)?month/.test(blob)) return false;
  if (/gujarat/.test(blob) && !/karnataka/.test(blob) && !/karnataka/.test(nameBlob)) return false;
  if (/\bform[\s._-]*p[\s._-]*gj\b/.test(blob) || /form_p_gj/.test(blob) || /form_p_gj/.test(nameBlob)) {
    return false;
  }

  const nameIsFormPKA =
    /form[\s._-]*p/.test(nameBlob) &&
    (/karnataka/.test(nameBlob) || /form[\s._-]*p[\s._-]*ka\b/.test(nameBlob) || /form_p_-_karnataka/.test(nameBlob));
  if (nameIsFormPKA) return true;

  const isFormP =
    /form[\s._\-–—'"]*p\b/.test(blob) || /form_p_/.test(blob) || /\bform\s*p\b/.test(blob);
  const isKA =
    /karnataka/.test(blob) ||
    /form[\s._-]*p[\s._-]*ka\b/.test(blob) ||
    /form_p_-_karnataka/.test(blob);
  const hasNotice = /notice of maximum leave accumulated/.test(blob);
  const hasRule20 = /see\s+rule\s*20/.test(blob);
  const hasShriSmt = /shri\s*\/?\s*smt/.test(blob);
  if (isFormP && (isKA || hasRule20) && (hasNotice || hasShriSmt || hasRule20)) return true;
  return Boolean(hasNotice && hasRule20 && hasShriSmt);
}

const cellText = (rows, r, c) => norm(rows?.[r]?.[c]);

const rowBlob = (row) =>
  (Array.isArray(row) ? row : [])
    .map((c) => lower(c))
    .filter(Boolean)
    .join(' ');

const valueAfterColon = (text) => {
  const src = norm(text);
  const idx = src.indexOf(':');
  if (idx < 0) return '';
  return src.slice(idx + 1).replace(/^[:\s.]+/, '').trim();
};

const isDotsOnly = (text) => /^\.+$/.test(norm(text).replace(/\s/g, ''));

const isSystemGeneratedNote = (text) =>
  /system\s+generated\s+document|computer\s+generated\s+document/.test(lower(text));

function collectNonEmptyCells(row) {
  return (Array.isArray(row) ? row : []).map((c) => norm(c)).filter((t) => t && !isDotsOnly(t));
}

function firstMatchingLine(rows, tester) {
  for (let r = 0; r < (rows || []).length; r += 1) {
    const cells = collectNonEmptyCells(rows[r]);
    for (const cell of cells) {
      if (tester(lower(cell), cell)) return cell;
    }
  }
  return '';
}

function isTitleLine(text) {
  const t = lower(text);
  return /^form\s*[-–]?\s*['']?p['']?$/.test(t) || (/^form\s*[-–]?\s*['']?p['']?/.test(t) && t.length < 24);
}

function isEstablishmentLine(text) {
  const t = lower(text);
  if (/authorised person|authorized person/.test(t)) return false;
  return (
    /name\s+and\s+address\s+of\s+(?:the\s+)?establishment/.test(t) ||
    /address\s+of\s+(?:the\s+)?establishment/.test(t)
  );
}

function isWorkerAddressLine(text) {
  const t = lower(text);
  if (isEstablishmentLine(t)) return false;
  if (/authorised|authorized/.test(t)) return false;
  return /^address\b/.test(t.trim());
}

function isLegalLine(text) {
  const t = lower(text);
  return (
    /it is hereby informed/.test(t) ||
    /maximum leave that can be accumulated/.test(t) ||
    /maharashtra shops and establishments/.test(t)
  );
}

function isLeaveHeaderRow(row) {
  const blob = rowBlob(row);
  if (!blob) return false;
  if (isLegalLine(blob) || /details of the leave/.test(blob)) return false;
  return /sr\.?\s*no/.test(blob) && /accumulated\s+leave/.test(blob);
}

function isFromTillRow(row) {
  const blob = rowBlob(row);
  return /\bfrom\b/.test(blob) && /\b(till|to)\b/.test(blob) && !/shri/.test(blob);
}

function packLeaveDataRow(row, startCol) {
  const src = Array.isArray(row) ? row : [];
  const slice = src.slice(startCol, startCol + 4).map((c) => norm(c));
  if (slice.some(Boolean)) return slice;
  const packed = src.map((c) => norm(c)).filter(Boolean);
  if (!packed.length) return ['', '', '', ''];
  const out = ['', '', '', ''];
  packed.slice(0, 4).forEach((v, i) => {
    out[i] = v;
  });
  return out;
}

function isLeaveDataRow(row, startCol) {
  const blob = rowBlob(row);
  if (!blob) return false;
  if (isTitleLine(blob) || isLegalLine(blob) || isLeaveHeaderRow(row) || isFromTillRow(row)) return false;
  if (/details of the leave/.test(blob)) return false;
  if (isEstablishmentLine(blob) || /shri\s*\/?\s*smt/.test(blob) || /^to,?\s*$/.test(blob.trim())) {
    return false;
  }
  if (isSystemGeneratedNote(blob)) return false;
  const packed = packLeaveDataRow(row, startCol);
  return packed.some((t) => t && !isDotsOnly(t));
}

function extractLeaveRows(rows) {
  let headerRow = -1;
  let startCol = 0;
  for (let r = 0; r < (rows || []).length; r += 1) {
    if (!isLeaveHeaderRow(rows[r])) continue;
    headerRow = r;
    const row = rows[r] || [];
    const idx = row.findIndex((c) => /sr\.?\s*no/i.test(norm(c)));
    if (idx >= 0) startCol = idx;
    break;
  }
  if (headerRow < 0) return [];

  let dataStart = headerRow + 1;
  for (let r = headerRow; r <= Math.min(headerRow + 4, (rows || []).length - 1); r += 1) {
    if (isFromTillRow(rows[r])) {
      dataStart = r + 1;
      break;
    }
  }

  const out = [];
  for (let r = dataStart; r < (rows || []).length; r += 1) {
    if (!isLeaveDataRow(rows[r], startCol)) {
      if (out.length) break;
      continue;
    }
    const packed = packLeaveDataRow(rows[r], startCol);
    if (packed.some(Boolean)) out.push(packed);
  }
  return out;
}

function padLeaveRows(rows) {
  const out = [...(rows || [])];
  while (out.length < MIN_LEAVE_BODY_ROWS) out.push(['', '', '', '']);
  return out;
}

export function normalizeFormPKarnatakaPdfMatrix(rows = [], colCount = 4, _tableStartRow = 0, metaLines = []) {
  const src = [];
  (Array.isArray(metaLines) ? metaLines : []).forEach((line) => {
    const t = norm(line);
    if (t) src.push([t]);
  });
  (Array.isArray(rows) ? rows : []).forEach((row) => {
    src.push(Array.isArray(row) ? row.map((c) => norm(c)) : [norm(row)]);
  });

  const establishment =
    firstMatchingLine(src, (n) => isEstablishmentLine(n)) ||
    firstMatchingLine(src, (n) => /name\s+and\s+address/.test(n));
  const authorised = firstMatchingLine(
    src,
    (n) => /authorised person|authorized person/.test(n) && /manager|authorised|authorized/.test(n)
  );
  const shriSmt = firstMatchingLine(src, (n) => /shri\s*\/?\s*smt/.test(n));
  const address = firstMatchingLine(src, (n, raw) => isWorkerAddressLine(raw));
  const legal =
    src
      .map((row) => collectNonEmptyCells(row).join(' '))
      .find((line) => isLegalLine(line)) || FORM_P_KA_PDF_LEGAL;

  const establishmentText = establishment
    ? /:/.test(establishment)
      ? establishment
      : `Name and address of the establishment : ${valueAfterColon(establishment) || ''}`.replace(
          /:\s*$/,
          ':'
        )
    : 'Name and address of the establishment :';
  const establishmentValue = valueAfterColon(establishmentText);
  const addressText = address
    ? /:/.test(address) && valueAfterColon(address)
      ? address.replace(/^address\.*/i, 'Address')
      : `Address: ${establishmentValue || valueAfterColon(address)}`.replace(/:\s*$/, ':')
    : establishmentValue
      ? `Address: ${establishmentValue}`
      : 'Address:';

  const model = {
    titles: [FORM_P_KA_PDF_TITLE, FORM_P_KA_PDF_RULE, FORM_P_KA_PDF_SUBTITLE],
    establishment: /:\s*$/.test(establishmentText) && establishmentValue
      ? `${establishmentText.replace(/:\s*$/, '')} : ${establishmentValue}`
      : establishmentText,
    authorised: authorised || 'Name of the Authorised person / Manager',
    to: 'To,',
    shriSmt: shriSmt || 'Shri/Smt.',
    address: addressText,
    legal,
    details: FORM_P_KA_PDF_DETAILS,
    leaveRows: padLeaveRows(extractLeaveRows(src)),
  };

  void colCount;
  return {
    rows: [
      [model.titles[0]],
      [model.titles[1]],
      [model.titles[2]],
      [model.establishment],
      [model.authorised],
      [model.to],
      [model.shriSmt],
      [model.address],
      [model.legal],
      [model.details],
      ['Sr. No.', 'Number of accumulated leave', 'From', 'Till'],
      ...model.leaveRows,
    ],
    colCount: 4,
    tableStartRow: 10,
    metaLines: model.titles,
    formPKALayout: true,
    formPKAModel: model,
  };
}

export const applyFormPKarnatakaPdfNormalization = (target, normalized) => {
  if (!target || !normalized) return target;
  target.rows = normalized.rows;
  target.colCount = normalized.colCount;
  target.tableStartRow = normalized.tableStartRow;
  target.metaLines = Array.isArray(normalized.metaLines) ? normalized.metaLines : [];
  target.formPKALayout = true;
  target.formPKAModel = normalized.formPKAModel;
  return target;
};

function strokeRect(doc, x, y, w, h) {
  doc.setDrawColor(0, 0, 0);
  doc.setLineWidth(0.6);
  doc.rect(x, y, w, h, 'S');
}

function drawWrapped(doc, text, x, y, w, size, align, bold) {
  doc.setFont('helvetica', bold ? 'bold' : 'normal');
  doc.setFontSize(size);
  const lines = doc.splitTextToSize(String(text || ''), Math.max(10, w - 4));
  const lineH = size + 1.6;
  const top = y + size;
  if (align === 'center') {
    doc.text(lines, x + w / 2, top, { align: 'center' });
  } else {
    doc.text(lines, x + 2, top);
  }
  return Math.max(lineH, lines.length * lineH);
}

function colWidths(total, weights) {
  const sum = weights.reduce((a, b) => a + b, 0) || 1;
  return weights.map((w) => (total * w) / sum);
}

/**
 * Draw Form P KA as the Excel notice letter, not a generic full-sheet grid.
 */
export function drawFormPKarnatakaNotice(doc, matrix, startY = 28) {
  const pageWidth = doc.internal.pageSize.getWidth();
  const pageHeight = doc.internal.pageSize.getHeight();
  const marginX = 36;
  const usableWidth = pageWidth - marginX * 2;
  const model =
    matrix?.formPKAModel ||
    normalizeFormPKarnatakaPdfMatrix(
      matrix?.rows || [],
      matrix?.colCount || 4,
      0,
      matrix?.metaLines || []
    ).formPKAModel;

  const boxTop = startY;
  const boxX = marginX - 6;
  const boxW = usableWidth + 12;
  let y = startY + 8;

  const paintCenter = (text, size, bold, h) => {
    doc.setFont('helvetica', bold ? 'bold' : 'normal');
    doc.setFontSize(size);
    doc.text(String(text || ''), marginX + usableWidth / 2, y + size + 2, { align: 'center' });
    y += h;
  };

  paintCenter(model.titles?.[0] || FORM_P_KA_PDF_TITLE, 12, true, 16);
  paintCenter(model.titles?.[1] || FORM_P_KA_PDF_RULE, 9, false, 14);
  paintCenter(model.titles?.[2] || FORM_P_KA_PDF_SUBTITLE, 11, true, 18);

  const fieldLines = [
    { text: model.establishment, size: 9, bold: false },
    { text: model.authorised, size: 9, bold: false },
    { text: model.to, size: 9, bold: false },
    { text: model.shriSmt, size: 9, bold: false },
    { text: model.address, size: 9, bold: false },
  ];
  fieldLines.forEach((field) => {
    const h = drawWrapped(doc, field.text, marginX, y, usableWidth, field.size, 'left', field.bold);
    y += Math.max(14, h + 2);
  });

  y += 4;
  const legalH = drawWrapped(doc, model.legal || FORM_P_KA_PDF_LEGAL, marginX, y, usableWidth, 8.5, 'left', false);
  y += legalH + 10;

  paintCenter(model.details || FORM_P_KA_PDF_DETAILS, 10, true, 16);

  const groups = [
    { label: 'Sr. No.', leafs: [{ key: 'sno', label: '', weight: 12 }] },
    { label: 'Number of accumulated leave', leafs: [{ key: 'leave', label: '', weight: 22 }] },
    {
      label: 'Period for which leave is accumulated',
      leafs: [
        { key: 'from', label: 'From', weight: 22 },
        { key: 'till', label: 'Till', weight: 22 },
      ],
    },
  ];
  const leafs = [];
  groups.forEach((g) => g.leafs.forEach((leaf) => leafs.push({ ...leaf, parent: g.label })));
  const widths = colWidths(
    usableWidth,
    leafs.map((leaf) => leaf.weight || 10)
  );
  const groupH = 22;
  const leafH = 16;
  const dataH = 18;
  const tableX = marginX;
  let cursor = tableX;
  groups.forEach((group) => {
    const spanW = group.leafs.reduce((sum, leaf) => {
      const idx = leafs.findIndex((l) => l.key === leaf.key);
      return sum + (widths[idx] || 0);
    }, 0);
    strokeRect(doc, cursor, y, spanW, groupH);
    drawWrapped(doc, group.label, cursor, y + 3, spanW, 8, 'center', true);
    cursor += spanW;
  });
  cursor = tableX;
  leafs.forEach((leaf, i) => {
    strokeRect(doc, cursor, y + groupH, widths[i], leafH);
    if (leaf.label) drawWrapped(doc, leaf.label, cursor, y + groupH + 2, widths[i], 8, 'center', true);
    cursor += widths[i];
  });

  const bodyY = y + groupH + leafH;
  const leaveRows = Array.isArray(model.leaveRows) ? model.leaveRows : [];
  leaveRows.forEach((row, ri) => {
    cursor = tableX;
    leafs.forEach((leaf, i) => {
      strokeRect(doc, cursor, bodyY + ri * dataH, widths[i], dataH);
      const val = Array.isArray(row) ? row[i] : '';
      drawWrapped(doc, val, cursor, bodyY + ri * dataH + 3, widths[i], 8, 'center', false);
      cursor += widths[i];
    });
  });
  y = bodyY + leaveRows.length * dataH + 10;

  const boxH = Math.max(y - boxTop + 8, 120);
  const maxH = pageHeight - boxTop - 28;
  strokeRect(doc, boxX, boxTop, boxW, Math.min(boxH, maxH));
  return y + 8;
}
